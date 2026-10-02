"""
北海道 就労継続支援A型事業所 情報収集

データソース:
  Phase 1 - shogaisha-shuro.com から A型事業所リスト取得 (253件)
            (requests + BeautifulSoup)
  Phase 2 - knowbe.jp/rakita (Next.js RSCストリーミングに埋め込まれたJSON) から
            A型事業所一覧を取得し、Phase 1 に無い事業所を追加 (全28ページ)
            住所は postal_code を 日本郵便の郵便番号データ(ken_all, UTF-8版)で
            都道府県・市区町村まで解決し、address_detail(番地)と結合する。
            (requests + 正規表現でRSC埋め込みJSONを解析)
  Phase 3 - 各公式サイトから メールアドレス・インスタURL・問い合わせフォームURL を取得
            (Playwright + BeautifulSoup)

出力列: 名称, メールアドレス, 公式サイトURL, 所在地, 電話番号, インスタURL, 問い合わせフォームURL

出力ファイル: hokkaido_a_type_YYYYMMDD_HHMMSS.xlsx
"""

import asyncio
import csv
import io
import logging
import random
import re
import sys
import time
import zipfile
from datetime import datetime
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup
from playwright.async_api import async_playwright

# ── ロギング ──────────────────────────────────────────────
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

LOG_FILE = "scraper_hokkaido_a_type.log"
_sh = logging.StreamHandler(
    io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    if hasattr(sys.stdout, "buffer") else sys.stdout
)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[_sh, logging.FileHandler(LOG_FILE, encoding="utf-8")],
)
logger = logging.getLogger(__name__)

# ── 定数 ─────────────────────────────────────────────────
OUTPUT_FILE = f"hokkaido_a_type_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/124.0.0.0 Safari/537.36"
)
HEADERS = {"User-Agent": USER_AGENT}

SHURO_BASE = "https://shogaisha-shuro.com"
SHURO_LIST_URL = f"{SHURO_BASE}/category/shuro/hokkaido/?t1=1&k="

KNOWBE_BASE = "https://knowbe.jp"
KNOWBE_LIST_URL = f"{KNOWBE_BASE}/rakita/keizoku/type_a/1"  # 1 = 北海道
KNOWBE_MAX_PAGES = 30  # 実件数に応じて自動停止するので余裕を持った上限

ZIP_DATA_URL = "https://www.post.japanpost.jp/service/search/zipcode/download/utf/zip/utf_ken_all.zip"

SKIP_DOMAINS = [
    "instagram.com", "facebook.com", "twitter.com", "x.com", "youtube.com",
    "tiktok.com", "wikipedia.org", "google.com", "bing.com",
    "shogaisha-shuro.com", "knowbe.jp", "wam.go.jp",
    "tabelog.com", "hotpepper.jp",
]

OUTPUT_COLS = [
    "名称", "メールアドレス", "公式サイトURL",
    "所在地", "電話番号", "インスタURL", "問い合わせフォームURL",
]


def _normalize_name(s: str) -> str:
    s = re.sub(r"[\s　・「」『』()（）\-－―！!]+", "", str(s))
    s = re.sub(r"[Ａ-Ｚａ-ｚ０-９]", lambda m: chr(ord(m.group(0)) - 0xFEE0), s)
    for pref in ["特定非営利活動法人", "一般社団法人", "一般財団法人",
                 "合同会社", "株式会社", "有限会社", "NPO法人"]:
        s = s.replace(pref, "")
    return s.lower()


def _normalize_phone(s: str) -> str:
    return re.sub(r"[^0-9]", "", str(s))


def _decode_cfemail(encoded: str) -> str:
    try:
        key = int(encoded[:2], 16)
        return bytes(
            int(encoded[i:i + 2], 16) ^ key for i in range(2, len(encoded), 2)
        ).decode("utf-8")
    except Exception:
        return ""


def _fetch_text(url: str) -> str:
    r = requests.get(url, headers=HEADERS, timeout=25)
    r.encoding = "utf-8"
    return r.text


# ── Phase 1: shogaisha-shuro.com ─────────────────────────

def collect_shuro_list() -> list[tuple[str, str]]:
    try:
        html = _fetch_text(SHURO_LIST_URL)
    except Exception as e:
        logger.warning(f"shuro一覧取得失敗: {e}")
        return []
    soup = BeautifulSoup(html, "lxml")
    results: list[tuple[str, str]] = []
    for li in soup.select("ul.facility_list li.clearfix"):
        if not li.find("div", class_="type-a-icon"):
            continue
        name_div = li.find("div", class_="institution-name")
        if not name_div:
            continue
        a = name_div.find("a", href=True)
        if not a:
            continue
        name = name_div.get_text(strip=True)
        url = urljoin(SHURO_BASE, a["href"])
        results.append((name, url))
    logger.info(f"shogaisha-shuro.com A型リスト: {len(results)} 件")
    return results


def parse_shuro_detail(detail_url: str) -> dict:
    rec: dict = {k: "" for k in OUTPUT_COLS}
    try:
        html = _fetch_text(detail_url)
    except Exception as e:
        logger.warning(f"詳細取得失敗 {detail_url}: {e}")
        return rec
    soup = BeautifulSoup(html, "lxml")

    h = soup.find("h1")
    if h:
        rec["名称"] = h.get_text(strip=True)

    for tr in soup.select("table tr"):
        th = tr.find("th")
        td = tr.find("td")
        if not th or not td:
            continue
        key = th.get_text(strip=True)
        val = td.get_text(separator=" ", strip=True)

        if key == "所在地":
            rec["所在地"] = re.sub(r"〒\s*\d{3}[-－]\d{4}\s*", "", val).strip()
        elif key == "電話番号":
            phone = re.search(r"[\d\-]{6,}", val.replace("－", "-"))
            rec["電話番号"] = phone.group(0) if phone else ""
        elif key == "Eメール":
            cf_el = td.find(attrs={"data-cfemail": True})
            if cf_el:
                decoded = _decode_cfemail(cf_el["data-cfemail"])
                if decoded:
                    rec["メールアドレス"] = decoded
        elif key == "URL":
            link = td.find("a", href=True)
            if link:
                href = link["href"].strip()
                if href.startswith("http") and not any(d in href for d in SKIP_DOMAINS):
                    rec["公式サイトURL"] = href

    return rec


# ── 郵便番号 → 都道府県/市区町村 ──────────────────────────

def build_zip_lookup() -> dict[str, str]:
    """日本郵便の郵便番号データから 郵便番号(7桁) -> '都道府県市区町村' の辞書を作る。"""
    try:
        r = requests.get(ZIP_DATA_URL, headers=HEADERS, timeout=60)
        r.raise_for_status()
        with zipfile.ZipFile(io.BytesIO(r.content)) as zf:
            csv_name = next(n for n in zf.namelist() if n.lower().endswith(".csv"))
            with zf.open(csv_name) as f:
                text = io.TextIOWrapper(f, encoding="utf-8")
                lookup: dict[str, str] = {}
                for row in csv.reader(text):
                    zipcode = row[2]
                    pref = row[6]
                    city = row[7]
                    if zipcode not in lookup:
                        lookup[zipcode] = f"{pref}{city}"
        logger.info(f"郵便番号データ読み込み完了: {len(lookup)} 件")
        return lookup
    except Exception as e:
        logger.warning(f"郵便番号データ取得失敗: {e}")
        return {}


# ── Phase 2: knowbe.jp/rakita ─────────────────────────────

_FACILITY_START_RE = re.compile(
    r'\{\\"id\\":(\d+),\\"facility_number\\":\\"\d+\\",\\"display_status\\":\\"active\\"'
)


def _get_field(block: str, key: str) -> str:
    m = re.search(rf'\\"{key}\\":\\"([^\\]*)\\"', block)
    if m:
        return m.group(1)
    m = re.search(rf'\\"{key}\\":(\d+)', block)
    if m:
        return m.group(1)
    return ""


def parse_knowbe_page(html: str, zip_lookup: dict[str, str]) -> list[dict]:
    matches = list(_FACILITY_START_RE.finditer(html))
    records: list[dict] = []
    for i, m in enumerate(matches):
        start = m.start()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(html)
        block = html[start:end]

        fd_idx = block.find('\\"facility_detail\\":{')
        wam_idx = block.find('\\"wam_facility_detail\\":{')
        fd_block = block[fd_idx:wam_idx] if fd_idx >= 0 and wam_idx > fd_idx else (
            block[fd_idx:] if fd_idx >= 0 else "")
        wam_block = block[wam_idx:] if wam_idx >= 0 else ""

        name = _get_field(fd_block, "name") or _get_field(wam_block, "name")
        if not name:
            continue

        postal_code = _get_field(fd_block, "postal_code")
        address_detail = _get_field(fd_block, "address_detail") or _get_field(wam_block, "address_detail")
        building_name = _get_field(fd_block, "building_name")
        tel = _get_field(fd_block, "tel")
        website_url = _get_field(fd_block, "website_url")

        if address_detail.startswith("北海道"):
            address = address_detail
        elif postal_code and postal_code in zip_lookup:
            address = zip_lookup[postal_code] + address_detail
        else:
            address = address_detail
        if building_name:
            address = f"{address} {building_name}"

        rec = {k: "" for k in OUTPUT_COLS}
        rec["名称"] = name
        rec["所在地"] = address
        rec["電話番号"] = tel
        if website_url.startswith("http") and not any(d in website_url for d in SKIP_DOMAINS):
            rec["公式サイトURL"] = website_url

        records.append({"facility_id": m.group(1), **rec})
    return records


def collect_knowbe_facilities(zip_lookup: dict[str, str]) -> list[dict]:
    seen_ids: set[str] = set()
    all_records: list[dict] = []
    for page in range(1, KNOWBE_MAX_PAGES + 1):
        url = KNOWBE_LIST_URL if page == 1 else f"{KNOWBE_LIST_URL}?page={page}"
        try:
            html = _fetch_text(url)
        except Exception as e:
            logger.warning(f"knowbe一覧取得失敗 page={page}: {e}")
            continue
        page_records = parse_knowbe_page(html, zip_lookup)
        if not page_records:
            logger.info(f"knowbe page={page}: 0件のため終了")
            break
        new_count = 0
        for rec in page_records:
            fid = rec.pop("facility_id")
            if fid in seen_ids:
                continue
            seen_ids.add(fid)
            all_records.append(rec)
            new_count += 1
        logger.info(f"knowbe page={page}: {new_count} 件 (累計 {len(all_records)} 件)")
        time.sleep(random.uniform(0.3, 0.6))
    return all_records


# ── Phase 3: 公式サイト追加情報 ──────────────────────────

async def scrape_official_site(url: str, page) -> dict:
    result: dict = {}
    if any(d in url for d in SKIP_DOMAINS):
        return result
    try:
        resp = await page.goto(url, timeout=18000, wait_until="domcontentloaded")
        if resp and resp.status >= 400:
            return result
        await page.wait_for_timeout(1500)
        html = await page.content()
    except Exception as e:
        logger.debug(f"Playwright失敗 {url}: {e}")
        try:
            r = requests.get(url, headers=HEADERS, timeout=10)
            r.encoding = r.apparent_encoding
            html = r.text
        except Exception:
            return result

    soup = BeautifulSoup(html, "lxml")
    text = soup.get_text(separator="\n")

    cf_els = soup.find_all(attrs={"data-cfemail": True})
    if cf_els:
        decoded = _decode_cfemail(cf_els[0]["data-cfemail"])
        if decoded:
            result["メールアドレス"] = decoded
    if not result.get("メールアドレス"):
        emails = re.findall(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}", text)
        emails = [e for e in emails if not re.search(r"\.(png|jpg|gif|svg|webp)$", e, re.I)]
        if emails:
            result["メールアドレス"] = emails[0]

    for a in soup.find_all("a", href=True):
        href = a["href"]
        if "instagram.com" in href and "/p/" not in href and "/reel/" not in href:
            if href.startswith("//"):
                href = "https:" + href
            result["インスタURL"] = href.rstrip("/")
            break

    contact_kws = ["contact", "inquiry", "お問い合わせ", "問い合わせ", "ご相談", "メールフォーム"]
    for a in soup.find_all("a", href=True):
        href = a["href"]
        link_text = a.get_text(strip=True)
        if not any(kw in href.lower() or kw in link_text for kw in contact_kws):
            continue
        if href.startswith("http"):
            result["問い合わせフォームURL"] = href
        elif href.startswith("//"):
            result["問い合わせフォームURL"] = "https:" + href
        elif not href.startswith(("#", "mailto:", "tel:")):
            result["問い合わせフォームURL"] = urljoin(url, href)
        if "問い合わせフォームURL" in result:
            break

    return result


# ── メイン ────────────────────────────────────────────────

async def main() -> None:
    logger.info("=" * 65)
    logger.info("北海道 就労継続支援A型事業所 収集開始")
    logger.info("=" * 65)
    start_time = datetime.now()

    # Phase 1
    logger.info("Phase 1: shogaisha-shuro.com から A型事業所収集")
    shuro_list = collect_shuro_list()
    records: list[dict] = []
    seen_names: set[str] = set()
    seen_phones: set[str] = set()

    for i, (name_hint, detail_url) in enumerate(shuro_list, 1):
        rec = parse_shuro_detail(detail_url)
        if not rec["名称"]:
            rec["名称"] = name_hint
        nk = _normalize_name(rec["名称"])
        pk = _normalize_phone(rec.get("電話番号", ""))
        if nk and nk not in seen_names:
            seen_names.add(nk)
            if pk:
                seen_phones.add(pk)
            records.append(rec)
        if i % 30 == 0 or i == len(shuro_list):
            logger.info(f"  [{i}/{len(shuro_list)}]")
        time.sleep(random.uniform(0.25, 0.45))

    logger.info(f"Phase 1 完了: {len(records)} 件")

    # Phase 2
    logger.info("Phase 2: knowbe.jp/rakita から A型事業所収集 (差分追加)")
    zip_lookup = build_zip_lookup()
    knowbe_records = collect_knowbe_facilities(zip_lookup)
    knowbe_added = 0
    for rec in knowbe_records:
        nk = _normalize_name(rec["名称"])
        pk = _normalize_phone(rec.get("電話番号", ""))
        if nk in seen_names:
            continue
        if pk and pk in seen_phones:
            continue
        seen_names.add(nk)
        if pk:
            seen_phones.add(pk)
        records.append(rec)
        knowbe_added += 1

    logger.info(f"Phase 2 完了: 新規追加 {knowbe_added} 件 (合計 {len(records)} 件)")

    # Phase 3
    logger.info("Phase 3: 公式サイト収集開始")
    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=True)
        context = await browser.new_context(user_agent=USER_AGENT)
        official_page = await context.new_page()

        with_url = [r for r in records if r.get("公式サイトURL")]
        logger.info(f"  公式URLあり: {len(with_url)} / {len(records)} 件")

        for i, rec in enumerate(with_url, 1):
            url = rec["公式サイトURL"]
            extras = await scrape_official_site(url, official_page)
            if rec.get("メールアドレス") and "メールアドレス" in extras:
                del extras["メールアドレス"]
            rec.update(extras)
            if i % 20 == 0 or i == len(with_url):
                logger.info(f"  公式サイト取得: {i}/{len(with_url)}")
            await asyncio.sleep(random.uniform(0.5, 1.2))

        await browser.close()

    # Excel 出力
    df = pd.DataFrame(records, columns=OUTPUT_COLS)
    df.drop_duplicates(subset=["名称", "電話番号"], keep="first", inplace=True)
    df.sort_values(by="所在地", inplace=True, kind="stable")
    df.to_excel(OUTPUT_FILE, index=False)

    elapsed = int((datetime.now() - start_time).total_seconds())
    logger.info("=" * 65)
    logger.info(f"完了: {OUTPUT_FILE} に {len(df)} 件を出力")
    logger.info(f"所要時間: {elapsed // 60}分{elapsed % 60}秒")
    logger.info("=" * 65)


if __name__ == "__main__":
    asyncio.run(main())
