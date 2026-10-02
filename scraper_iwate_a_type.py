"""
岩手県 就労継続支援A型事業所 情報収集

データソース:
  Phase 1 - shogaisha-shuro.com から A型事業所リスト取得 (54件)
            (requests + BeautifulSoup)
  Phase 2 - snabi.jp (LITALICO仕事ナビ) から A型事業所一覧を取得し、
            Phase 1 に無い事業所を追加 (全3ページ, 49件中の差分)
            (requests + 正規表現でSSR埋め込みJSONを解析)
  Phase 2b - Phase 2 で住所が非公開(掲載プランの都合で詳細非表示)だった
             事業所について、公式サイト等で裏取りした住所・電話番号を補完
             (2026-10-02 時点の手動確認済み情報)
  Phase 3 - 各公式サイトから メールアドレス・インスタURL・問い合わせフォームURL を取得
            (Playwright + BeautifulSoup)

出力列: 名称, メールアドレス, 公式サイトURL, 所在地, 電話番号, インスタURL, 問い合わせフォームURL

出力ファイル: iwate_a_type_YYYYMMDD_HHMMSS.xlsx
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

import pandas as pd
import requests
from bs4 import BeautifulSoup
from playwright.async_api import async_playwright

# ── ロギング ──────────────────────────────────────────────
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

LOG_FILE = "scraper_iwate_a_type.log"
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
OUTPUT_FILE = f"iwate_a_type_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/124.0.0.0 Safari/537.36"
)
HEADERS = {"User-Agent": USER_AGENT}

SHURO_BASE = "https://shogaisha-shuro.com"
SHURO_LIST_URL = f"{SHURO_BASE}/category/shuro/iwate/?t1=1&k="

SNABI_BASE = "https://snabi.jp"
SNABI_LIST_URL = f"{SNABI_BASE}/keizoku/type_a/3"  # 3 = 岩手県
SNABI_PAGES = 3

SKIP_DOMAINS = [
    "instagram.com", "facebook.com", "twitter.com", "x.com", "youtube.com",
    "tiktok.com", "wikipedia.org", "google.com", "bing.com",
    "shogaisha-shuro.com", "snabi.jp", "wam.go.jp",
    "tabelog.com", "hotpepper.jp",
]

OUTPUT_COLS = [
    "名称", "メールアドレス", "公式サイトURL",
    "所在地", "電話番号", "インスタURL", "問い合わせフォームURL",
]

# Phase 2b: snabi.jp 側で掲載プランの都合により住所が非公開だった事業所について、
# 公式サイト・求人媒体等で裏取りした情報 (2026-10-02 手動確認)
MANUAL_PATCH = {
    "フォレストファーム": {
        "所在地": "岩手県紫波郡矢巾町煙山第31地割60番地5",
        "電話番号": "019-611-1110",
    },
    "Barrier Free 盛岡": {
        "所在地": "岩手県盛岡市柳町1-6-9 第10マルビル2-C",
        "電話番号": "019-613-3881",
        "公式サイトURL": "https://barrierfree.co.jp/locations/morioka/",
    },
    "ワンステップ": {
        "所在地": "岩手県盛岡市内丸13番1号 岩手県民会館1階",
        "電話番号": "019-681-0588",
        "公式サイトURL": "http://i-ppo.jp/",
    },
    "ハーモニー八幡平": {
        "所在地": "岩手県八幡平市野駄第14地割97番1",
        "電話番号": "0195-78-8801",
        "公式サイトURL": "https://harmony-hachimantai.jp/",
    },
    "カノン": {
        "所在地": "岩手県盛岡市中央通3-3-2 菱和ビル4-A",
        "電話番号": "019-613-4150",
        "公式サイトURL": "https://senju.co/canon/",
    },
    # BAGLAB(奥州市)は公開情報が見つからず、市区町村のみ判明
    "BAGLAB": {
        "所在地": "岩手県奥州市(詳細住所非公開)",
        "電話番号": "",
    },
}


def _normalize_name(s: str) -> str:
    s = re.sub(r"[\s　・「」『』()（）\-－―]+", "", str(s))
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
    r = requests.get(url, headers=HEADERS, timeout=20)
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


# ── Phase 2: snabi.jp (LITALICO仕事ナビ) ──────────────────

def collect_snabi_facilities() -> list[tuple[str, str]]:
    """(名称, facility_id) のリストを返す。全ページ分の SSR 埋め込みJSONを解析。"""
    found: dict[str, str] = {}
    for page in range(1, SNABI_PAGES + 1):
        url = SNABI_LIST_URL if page == 1 else f"{SNABI_LIST_URL}?page={page}"
        try:
            html = _fetch_text(url)
        except Exception as e:
            logger.warning(f"snabi一覧取得失敗 page={page}: {e}")
            continue
        for m in re.finditer(
            r'&quot;id&quot;:(\d+),&quot;name&quot;:&quot;([^&]*)&quot;,&quot;hurigana&quot;:',
            html,
        ):
            fid, name = m.group(1), m.group(2)
            found[fid] = name
        time.sleep(0.3)
    logger.info(f"snabi.jp A型リスト: {len(found)} 件")
    return list(found.items())


def parse_snabi_detail(fid: str) -> dict:
    rec: dict = {k: "" for k in OUTPUT_COLS}
    try:
        html = _fetch_text(f"{SNABI_BASE}/facility/{fid}")
    except Exception as e:
        logger.warning(f"snabi詳細取得失敗 {fid}: {e}")
        return rec
    m = re.search(r'&quot;full_address&quot;:&quot;([^&]*)&quot;', html)
    if m:
        rec["所在地"] = m.group(1)
    return rec


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
    logger.info("岩手県 就労継続支援A型事業所 収集開始")
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
        if i % 20 == 0 or i == len(shuro_list):
            logger.info(f"  [{i}/{len(shuro_list)}] {rec['名称']}")
        time.sleep(random.uniform(0.3, 0.6))

    logger.info(f"Phase 1 完了: {len(records)} 件")

    # Phase 2
    logger.info("Phase 2: snabi.jp から A型事業所収集 (差分追加)")
    snabi_list = collect_snabi_facilities()
    snabi_added = 0
    for i, (fid, name) in enumerate(snabi_list, 1):
        nk = _normalize_name(name)
        if nk in seen_names:
            continue
        rec = parse_snabi_detail(fid)
        rec["名称"] = name
        seen_names.add(nk)
        records.append(rec)
        snabi_added += 1
        logger.info(f"  [新規 {snabi_added}] {name} -> {rec.get('所在地', '')}")
        time.sleep(random.uniform(0.3, 0.6))

    logger.info(f"Phase 2 完了: 新規追加 {snabi_added} 件 (合計 {len(records)} 件)")

    # Phase 2b: 手動確認済みパッチ
    logger.info("Phase 2b: 住所非公開事業所の手動確認済み情報を補完")
    for rec in records:
        patch = MANUAL_PATCH.get(rec["名称"])
        if not patch:
            continue
        for k, v in patch.items():
            if v and not rec.get(k):
                rec[k] = v

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
            if i % 10 == 0 or i == len(with_url):
                logger.info(f"  公式サイト取得: {i}/{len(with_url)}")
            await asyncio.sleep(random.uniform(0.6, 1.5))

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
