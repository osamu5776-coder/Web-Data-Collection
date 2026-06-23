"""
埼玉県皮膚科クリニック 網羅的情報収集

データソース:
  Phase 1 - 埼玉県皮膚科医会 (saitamahifuka.org) から全クリニックの基本情報を取得
            (requests + BeautifulSoup で静的ページを収集)
  Phase 2 - 各クリニックの公式サイトから追加情報を取得
            (Playwright + BeautifulSoup で JS ページも対応)

出力: saitama_all_YYYYMMDD_HHMMSS.xlsx
"""

import asyncio
import io
import logging
import random
import re
import sys
import time
from datetime import datetime
from urllib.parse import urljoin, quote
from urllib.robotparser import RobotFileParser

import pandas as pd
import requests
from bs4 import BeautifulSoup
from playwright.async_api import async_playwright

# ── ロギング ──────────────────────────────────────────────
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

LOG_FILE = "scraper_saitama_all.log"
_stream_handler = logging.StreamHandler(
    io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    if hasattr(sys.stdout, "buffer") else sys.stdout
)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[_stream_handler, logging.FileHandler(LOG_FILE, encoding="utf-8")],
)
logger = logging.getLogger(__name__)

# ── 定数 ─────────────────────────────────────────────────
OUTPUT_FILE = f"saitama_all_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/124.0.0.0 Safari/537.36"
)
HEADERS = {"User-Agent": USER_AGENT}

BASE_URL    = "http://saitamahifuka.org"
SERP_URL    = BASE_URL + "/search2/public/serp"
DETAIL_URL  = BASE_URL + "/search2/public/detail"

# 埼玉県内の全市区町村（saitamahifuka.org の select 値）
SAITAMA_CITIES = [
    "上里町", "本庄市", "美里町", "深谷市", "熊谷市", "行田市", "羽生市", "加須市",
    "神川町", "長瀞町", "寄居町", "皆野町", "小鹿野町", "秩父市", "横瀬町", "東秩父村",
    "小川町", "嵐山町", "滑川町", "ときがわ町", "越生町", "毛呂山町", "飯能市", "鳩山町",
    "東松山市", "吉見町", "鴻巣市", "坂戸市", "日高市", "鶴ヶ島市", "川島町", "北本市",
    "桶川市", "川越市", "狭山市", "入間市", "所沢市", "ふじみ野市", "三芳町", "富士見市",
    "志木市", "朝霞市", "新座市", "和光市", "上尾市", "伊奈町", "久喜市", "幸手市",
    "白岡町", "宮代町", "杉戸町", "蓮田市", "春日部市", "松伏町", "吉川市", "三郷市",
    "八潮市", "草加市", "戸田市", "蕨市", "川口市", "越谷市",
    "さいたま市 岩槻区", "さいたま市 緑区", "さいたま市 見沼区", "さいたま市 北区",
    "さいたま市 西区", "さいたま市 大宮区", "さいたま市 中央区", "さいたま市 桜区",
    "さいたま市 浦和区", "さいたま市 南区",
]

# 公式サイト収集時に除外するドメイン
SKIP_DOMAINS = [
    "instagram.com", "facebook.com", "twitter.com", "x.com", "youtube.com",
    "tiktok.com", "wikipedia.org", "google.com", "bing.com",
    "tabelog.com", "hotpepper.jp", "caloo.jp", "dr-map.jp", "qlife.jp",
    "byoinnavi.jp", "medley.life", "minnano-clinic.jp",
    "saitamahifuka.org",
]

# ── robots.txt キャッシュ ──────────────────────────────────
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
    rp = _get_robots(url)
    allowed = rp.can_fetch(USER_AGENT, url)
    if not allowed:
        logger.warning(f"robots.txt 禁止: {url}")
    return allowed

# ── Phase 1: saitamahifuka.org スクレイピング ─────────────

def _fetch(url: str) -> BeautifulSoup | None:
    try:
        resp = requests.get(url, headers=HEADERS, timeout=15)
        resp.encoding = "utf-8"
        return BeautifulSoup(resp.text, "lxml")
    except Exception as e:
        logger.warning(f"取得失敗: {url} / {e}")
        return None

def collect_detail_ids(city: str) -> list[str]:
    """市区町村の全ページを巡回して detail ID を収集する。"""
    ids: list[str] = []
    page_num = 1
    while True:
        if page_num == 1:
            url = f"{SERP_URL}?city={quote(city)}"
        else:
            url = f"{SERP_URL}/page:{page_num}?city={quote(city)}"

        soup = _fetch(url)
        if soup is None:
            break

        found = [
            a["href"].split("/detail/")[-1]
            for a in soup.find_all("a", href=True)
            if "/search2/public/detail/" in a["href"]
        ]
        if not found:
            break

        ids.extend(did for did in found if did not in ids)

        # 次ページリンク確認
        next_link = soup.find("a", href=lambda h: h and f"page:{page_num + 1}" in h)
        if not next_link:
            break
        page_num += 1
        time.sleep(random.uniform(0.5, 1.2))

    return ids


def fetch_detail(detail_id: str) -> dict:
    """detail ページから基本情報を取得する。"""
    url = f"{DETAIL_URL}/{detail_id}"
    soup = _fetch(url)
    rec = {
        "名称": "", "メールアドレス": "", "公式サイトURL": "",
        "所在地": "", "電話番号": "", "インスタURL": "", "問い合わせフォームURL": "",
    }
    if soup is None:
        return rec

    for row in soup.select("table tr, dl dt"):
        th = row.find("th") or row
        td = row.find("td") or row.find_next_sibling("dd")
        if th is None or td is None:
            continue
        key = th.get_text(strip=True)
        val = td.get_text(separator=" ", strip=True)

        if key in ("病院・医院名",):
            rec["名称"] = val
        elif key == "電話番号":
            rec["電話番号"] = val.split()[0] if val else ""
        elif key == "住所":
            # 「〒336-0926 埼玉県さいたま市…」→ 郵便番号を除去
            addr = re.sub(r"〒\s*\d{3}[-－]\d{4}\s*", "", val).strip()
            rec["所在地"] = addr
        elif key == "ホームページ":
            link = td.find("a", href=True)
            if link:
                rec["公式サイトURL"] = link["href"].strip()

    return rec


# ── Phase 2: 公式サイト スクレイピング ──────────────────────

async def scrape_official_site(url: str, page) -> dict:
    """Playwright で公式サイトを開き追加情報を抽出する。"""
    result: dict = {}
    if not _is_allowed(url):
        return result

    try:
        resp = await page.goto(url, timeout=15000, wait_until="domcontentloaded")
        if resp and resp.status >= 400:
            return result
        await page.wait_for_timeout(1200)
        html = await page.content()
    except Exception as e:
        logger.debug(f"Playwright 失敗 {url}: {e}")
        # フォールバック: requests
        try:
            r = requests.get(url, headers=HEADERS, timeout=10)
            r.encoding = r.apparent_encoding
            html = r.text
        except Exception:
            return result

    soup = BeautifulSoup(html, "lxml")
    text = soup.get_text(separator="\n")

    # メールアドレス
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
    logger.info("=" * 55)
    logger.info("埼玉県 皮膚科 全件収集 開始")
    logger.info("=" * 55)
    start_time = datetime.now()

    # ── Phase 1A: 全市区町村から detail ID を収集 ──────────
    logger.info(f"Phase 1A: {len(SAITAMA_CITIES)} 市区町村を検索")
    all_ids: list[str] = []
    seen: set[str] = set()

    for i, city in enumerate(SAITAMA_CITIES, 1):
        ids = collect_detail_ids(city)
        new_ids = [d for d in ids if d not in seen]
        seen.update(new_ids)
        all_ids.extend(new_ids)
        logger.info(f"  [{i}/{len(SAITAMA_CITIES)}] {city}: {len(ids)} 件 (累計 {len(all_ids)} 件)")
        time.sleep(random.uniform(0.8, 1.5))

    logger.info(f"Phase 1A 完了: ユニーク {len(all_ids)} クリニック")

    # ── Phase 1B: 各 detail ページから基本情報取得 ──────────
    logger.info("Phase 1B: 詳細ページ収集開始")
    records: list[dict] = []

    for i, did in enumerate(all_ids, 1):
        rec = fetch_detail(did)
        records.append(rec)
        if i % 20 == 0 or i == len(all_ids):
            logger.info(f"  詳細取得: {i}/{len(all_ids)}")
        time.sleep(random.uniform(0.5, 1.0))

    logger.info(f"Phase 1B 完了: {len(records)} 件")

    # ── Phase 2: 公式サイトから追加情報取得 ─────────────────
    logger.info("Phase 2: 公式サイト収集開始")

    with_url = [r for r in records if r.get("公式サイトURL")]
    logger.info(f"  公式URL あり: {len(with_url)} / {len(records)} 件")

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=True)
        context = await browser.new_context(user_agent=USER_AGENT)
        page = await context.new_page()

        for i, rec in enumerate(with_url, 1):
            url = rec["公式サイトURL"]
            if any(d in url for d in SKIP_DOMAINS):
                continue

            extras = await scrape_official_site(url, page)
            rec.update(extras)

            if i % 10 == 0 or i == len(with_url):
                logger.info(f"  公式サイト取得: {i}/{len(with_url)}")
            await asyncio.sleep(random.uniform(1.0, 2.5))

        await browser.close()

    # ── Excel 出力 ───────────────────────────────────────────
    columns = [
        "名称", "メールアドレス", "公式サイトURL",
        "所在地", "電話番号", "インスタURL", "問い合わせフォームURL",
    ]
    df = pd.DataFrame(records, columns=columns)
    df.drop_duplicates(subset=["名称", "電話番号"], keep="first", inplace=True)
    df.to_excel(OUTPUT_FILE, index=False)

    elapsed = int((datetime.now() - start_time).total_seconds())
    logger.info("=" * 55)
    logger.info(f"完了: {OUTPUT_FILE} に {len(df)} 件を出力 (所要時間: {elapsed // 60}分{elapsed % 60}秒)")
    logger.info("=" * 55)


if __name__ == "__main__":
    asyncio.run(main())
