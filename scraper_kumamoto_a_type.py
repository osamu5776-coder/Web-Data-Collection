"""
熊本県 就労継続支援A型事業所 情報収集

データソース:
  Phase 1 - shogaisha-shuro.com から A型事業所リスト取得 (134件)
            (requests + BeautifulSoup)
  Phase 2 - 熊本県公式Excel (pref.kumamoto.jp) から追加情報取得
            (requests でダウンロード → pandas で A型シートを読み込み)
  Phase 3 - 各公式サイトから インスタURL・問い合わせフォームURL を取得
            (Playwright + BeautifulSoup)

出力列: 名称, メールアドレス, 公式サイトURL, 所在地, 電話番号, インスタURL, 問い合わせフォームURL

出力ファイル: kumamoto_a_type_YYYYMMDD_HHMMSS.xlsx
"""

import asyncio
import io
import logging
import random
import re
import sys
import time
from datetime import datetime
from urllib.parse import urljoin
from urllib.robotparser import RobotFileParser

import pandas as pd
import requests
from bs4 import BeautifulSoup
from playwright.async_api import async_playwright

# ── ロギング ──────────────────────────────────────────────
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

LOG_FILE = "scraper_kumamoto_a_type.log"
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
OUTPUT_FILE = f"kumamoto_a_type_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/124.0.0.0 Safari/537.36"
)
HEADERS = {"User-Agent": USER_AGENT}

SHURO_BASE = "https://shogaisha-shuro.com"
SHURO_LIST_URL = f"{SHURO_BASE}/category/shuro/kumamoto/?t1=1&k="

KUMAMOTO_EXCEL_URL = (
    "https://www.pref.kumamoto.jp/uploaded/attachment/313385.xlsx"
)
KUMAMOTO_SHEET = "就労継続支援A型"

SKIP_DOMAINS = [
    "instagram.com", "facebook.com", "twitter.com", "x.com", "youtube.com",
    "tiktok.com", "wikipedia.org", "google.com", "bing.com",
    "shogaisha-shuro.com", "wam.go.jp", "pref.kumamoto.jp",
    "tabelog.com", "hotpepper.jp",
]

OUTPUT_COLS = [
    "名称", "メールアドレス", "公式サイトURL",
    "所在地", "電話番号", "インスタURL", "問い合わせフォームURL",
]

# ── robots.txt キャッシュ ─────────────────────────────────
_robots_cache: dict[str, RobotFileParser] = {}

def _get_robots(url: str) -> RobotFileParser:
    from urllib.parse import urlparse
    parsed = urlparse(url)
    origin = f"{parsed.scheme}://{parsed.netloc}"
    if origin in _robots_cache:
        return _robots_cache[origin]
    rp = RobotFileParser()
    try:
        rp.set_url(f"{origin}/robots.txt")
        rp.read()
    except Exception:
        pass
    _robots_cache[origin] = rp
    return rp

def _is_allowed(url: str) -> bool:
    return _get_robots(url).can_fetch(USER_AGENT, url)

def _fetch(url: str) -> BeautifulSoup | None:
    try:
        r = requests.get(url, headers=HEADERS, timeout=15)
        r.encoding = "utf-8"
        return BeautifulSoup(r.text, "lxml")
    except Exception as e:
        logger.warning(f"取得失敗: {url} / {e}")
        return None

def _decode_cfemail(encoded: str) -> str:
    try:
        key = int(encoded[:2], 16)
        return bytes(
            int(encoded[i:i+2], 16) ^ key for i in range(2, len(encoded), 2)
        ).decode("utf-8")
    except Exception:
        return ""

def _normalize_name(s: str) -> str:
    s = re.sub(r"[\s　]+", "", str(s))
    s = re.sub(r"[Ａ-Ｚａ-ｚ０-９]", lambda m: chr(ord(m.group(0)) - 0xFEE0), s)
    return s.lower()

def _normalize_phone(s: str) -> str:
    return re.sub(r"[^0-9]", "", str(s))


# ── Phase 1: shogaisha-shuro.com ─────────────────────────

def collect_shuro_list() -> list[tuple[str, str]]:
    """A型事業所の (名称ヒント, 詳細URL) リストを返す。"""
    soup = _fetch(SHURO_LIST_URL)
    if soup is None:
        return []
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
        href = a["href"]
        url = href if href.startswith("http") else urljoin(SHURO_BASE, href)
        results.append((name, url))
    logger.info(f"shogaisha-shuro.com A型リスト: {len(results)} 件")
    return results


def parse_shuro_detail(detail_url: str) -> dict:
    """shogaisha-shuro.com 詳細ページから基本情報を取得する。"""
    rec: dict = {k: "" for k in OUTPUT_COLS}
    soup = _fetch(detail_url)
    if soup is None:
        return rec

    h = soup.find("h1") or soup.find("h2")
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
            else:
                m = re.search(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}", val)
                if m:
                    rec["メールアドレス"] = m.group(0)
        elif key == "URL":
            link = td.find("a", href=True)
            if link:
                href = link["href"].strip()
                if not any(d in href for d in SKIP_DOMAINS):
                    rec["公式サイトURL"] = href

    return rec


# ── Phase 2: 熊本県公式Excel ──────────────────────────────

def collect_kumamoto_excel() -> list[dict]:
    """熊本県公式ExcelからA型事業所を取得する。"""
    try:
        r = requests.get(KUMAMOTO_EXCEL_URL, headers=HEADERS, timeout=30)
        r.raise_for_status()
        df = pd.read_excel(io.BytesIO(r.content), sheet_name=KUMAMOTO_SHEET)
    except Exception as e:
        logger.warning(f"熊本Excel取得失敗: {e}")
        return []

    # 提供中のみ
    if "指定状態" in df.columns:
        df = df[df["指定状態"] == "提供中"]

    records: list[dict] = []
    for _, row in df.iterrows():
        name = str(row.get("サービス事業所名", "")).strip()
        if not name or name == "nan":
            continue

        addr1 = str(row.get("サービス事業所所在地1", "")).replace("nan", "").strip()
        addr2 = str(row.get("サービス事業所所在地2", "")).replace("nan", "").strip()
        address = (addr1 + addr2).strip()

        tel = str(row.get("サービス事業所電話", "")).replace("nan", "").strip()
        phone = re.sub(r"[^0-9\-]", "", tel)

        rec: dict = {k: "" for k in OUTPUT_COLS}
        rec["名称"] = name
        rec["所在地"] = address
        rec["電話番号"] = phone
        records.append(rec)

    logger.info(f"熊本Excel A型: {len(records)} 件 (提供中)")
    return records


# ── Phase 3: 公式サイト追加情報 ──────────────────────────

async def scrape_official_site(url: str, page) -> dict:
    """公式サイトからメール・インスタ・問い合わせURLを取得する。"""
    result: dict = {}
    if not _is_allowed(url):
        return result
    if any(d in url for d in SKIP_DOMAINS):
        return result

    try:
        resp = await page.goto(url, timeout=18000, wait_until="domcontentloaded")
        if resp and resp.status >= 400:
            return result
        await page.wait_for_timeout(1500)
        html = await page.content()
    except Exception as e:
        logger.debug(f"Playwright 失敗 {url}: {e}")
        try:
            r = requests.get(url, headers=HEADERS, timeout=10)
            r.encoding = r.apparent_encoding
            html = r.text
        except Exception:
            return result

    soup = BeautifulSoup(html, "lxml")
    text = soup.get_text(separator="\n")

    # メールアドレス (Cloudflare 難読化対応)
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

    # Instagram URL
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if "instagram.com" in href and "/p/" not in href and "/reel/" not in href:
            if href.startswith("//"):
                href = "https:" + href
            result["インスタURL"] = href.rstrip("/")
            break

    # 問い合わせフォームURL
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
    logger.info("熊本県 就労継続支援A型事業所 収集開始")
    logger.info("=" * 65)
    start_time = datetime.now()

    # ── Phase 1: shogaisha-shuro.com ──────────────────────
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
        if i % 20 == 0 or i == len(shuro_list):
            logger.info(f"  [{i}/{len(shuro_list)}] {rec['名称']}")
        time.sleep(random.uniform(0.4, 1.0))

    logger.info(f"Phase 1 完了: {len(records)} 件")

    # ── Phase 2: 熊本県公式Excel ───────────────────────────
    logger.info("Phase 2: 熊本県公式Excel から A型事業所収集")
    excel_records = collect_kumamoto_excel()
    excel_added = 0

    for rec in excel_records:
        nk = _normalize_name(rec["名称"])
        pk = _normalize_phone(rec.get("電話番号", ""))
        # 名称・電話番号どちらかで既存チェック
        if nk in seen_names:
            continue
        if pk and pk in seen_phones:
            continue
        seen_names.add(nk)
        if pk:
            seen_phones.add(pk)
        records.append(rec)
        excel_added += 1

    logger.info(f"Phase 2 完了: 新規追加 {excel_added} 件 (合計 {len(records)} 件)")

    # ── Phase 3: 公式サイト追加情報 ────────────────────────
    logger.info("Phase 3: 公式サイト収集開始")
    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=True)
        context = await browser.new_context(user_agent=USER_AGENT)
        official_page = await context.new_page()

        with_url = [r for r in records if r.get("公式サイトURL")]
        logger.info(f"  公式URL あり: {len(with_url)} / {len(records)} 件")

        for i, rec in enumerate(with_url, 1):
            url = rec["公式サイトURL"]
            extras = await scrape_official_site(url, official_page)
            # shogaisha-shuro で取得済みのメールは上書きしない
            if rec.get("メールアドレス") and "メールアドレス" in extras:
                del extras["メールアドレス"]
            rec.update(extras)

            if i % 20 == 0 or i == len(with_url):
                logger.info(f"  公式サイト取得: {i}/{len(with_url)}")
            await asyncio.sleep(random.uniform(0.8, 2.0))

        await browser.close()

    # ── Excel 出力 ────────────────────────────────────────
    df = pd.DataFrame(records, columns=OUTPUT_COLS)
    df.drop_duplicates(subset=["名称", "電話番号"], keep="first", inplace=True)
    df.to_excel(OUTPUT_FILE, index=False)

    elapsed = int((datetime.now() - start_time).total_seconds())
    logger.info("=" * 65)
    logger.info(f"完了: {OUTPUT_FILE} に {len(df)} 件を出力")
    logger.info(f"所要時間: {elapsed // 60}分{elapsed % 60}秒")
    logger.info("=" * 65)


if __name__ == "__main__":
    asyncio.run(main())
