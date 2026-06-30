"""
宮城県 整体・接骨院 情報収集

データソース:
  Phase 1 - judo-ch.jp から宮城県全市区町村の接骨院を収集
            (requests + BeautifulSoup, 全エリア・全ページ巡回)
  Phase 2 - hotpepper.jp から追加整骨院を収集
            (Playwright + BeautifulSoup)
  Phase 3 - DuckDuckGo で各院の公式サイトを検索
            (ddgs)
  Phase 4 - 公式サイトから メール・インスタ・問い合わせURL を取得
            (Playwright + BeautifulSoup)

出力: miyagi_seitai_YYYYMMDD_HHMMSS.xlsx
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
from ddgs import DDGS
from playwright.async_api import async_playwright

# ── ロギング ──────────────────────────────────────────────
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

LOG_FILE = "scraper_miyagi_seitai.log"
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
OUTPUT_FILE = f"miyagi_seitai_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/124.0.0.0 Safari/537.36"
)
HEADERS = {"User-Agent": USER_AGENT}

JUDO_BASE = "https://www.judo-ch.jp"
JUDO_PREF_URL = f"{JUDO_BASE}/sekkotsuinsrch/04/"

HOTPEPPER_LIST_URL = "https://beauty.hotpepper.jp/genre/kgkw504/pre04/"
HOTPEPPER_BASE = "https://beauty.hotpepper.jp"

# DuckDuckGo 検索時に除外するディレクトリドメイン
SKIP_DOMAINS = [
    "judo-ch.jp", "homemate-research.com", "hotpepper.jp",
    "instagram.com", "facebook.com", "twitter.com", "x.com", "youtube.com",
    "tiktok.com", "wikipedia.org", "google.com", "bing.com",
    "mapion.co.jp", "iタウンページ", "itp.ne.jp", "ekiten.jp",
    "salonboard.com", "beauty.rakuten.co.jp", "minkabu.jp",
    "enma.jp", "seikotsuin.info", "wellyou.net",
    "judi-channel.jp", "medley.life", "caloo.jp",
    "toresei.com", "seitai-rct.jp",
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
        r = requests.get(url, headers=HEADERS, timeout=12)
        r.encoding = r.apparent_encoding or "utf-8"
        return BeautifulSoup(r.text, "lxml")
    except Exception as e:
        logger.debug(f"取得失敗: {url} / {e}")
        return None

def _normalize_name(s: str) -> str:
    s = re.sub(r"[\s　]+", "", str(s))
    s = re.sub(r"[Ａ-Ｚａ-ｚ０-９]", lambda m: chr(ord(m.group(0)) - 0xFEE0), s)
    return s.lower()

def _normalize_phone(s: str) -> str:
    return re.sub(r"[^0-9]", "", str(s))


# ── Phase 1: judo-ch.jp ──────────────────────────────────

def get_area_urls() -> list[tuple[str, str]]:
    """宮城県の全エリアURL (エリア名, URL) を返す。"""
    soup = _fetch(JUDO_PREF_URL)
    if soup is None:
        return []
    seen: set[str] = set()
    results: list[tuple[str, str]] = []
    for a in soup.find_all("a", href=True):
        href = a["href"]
        # パターン: /04/04XXX/ (5桁の市区町村コード)
        if re.search(r"/sekkotsuinsrch/04/\d{5}/$", href):
            if href not in seen:
                seen.add(href)
                full = href if href.startswith("http") else urljoin(JUDO_BASE, href)
                results.append((a.get_text(strip=True), full))
    logger.info(f"judo-ch.jp: {len(results)} エリア")
    return results


def scrape_judo_area(area_url: str) -> list[dict]:
    """エリアページ（全ページ）からクリニック情報を収集する。"""
    records: list[dict] = []
    page_num = 1
    while True:
        url = area_url if page_num == 1 else f"{area_url}{page_num}/"
        soup = _fetch(url)
        if soup is None:
            break

        for h3 in soup.find_all("h3"):
            a = h3.find("a", href=True)
            if a and "/sekkotsuinsrch/" in a["href"]:
                name = h3.get_text(strip=True)
                detail_href = a["href"]
                detail_url = detail_href if detail_href.startswith("http") else urljoin(JUDO_BASE, detail_href)
                # syuhenshisetsulist ページは除外
                if "syuhenshisetsulist" in detail_url:
                    continue
                rec = {k: "" for k in OUTPUT_COLS}
                rec["名称"] = name
                rec["_detail_url"] = detail_url
                records.append(rec)

        # 次ページ確認
        next_link = soup.find("a", string="次のページ")
        if not next_link:
            break
        page_num += 1
        if page_num > 20:
            break
        time.sleep(random.uniform(0.5, 1.0))

    return records


def _fetch_detail_sync(rec: dict) -> None:
    """judo-ch.jp 詳細ページから住所・電話を取得して rec を更新する (同期版)。"""
    detail_url = rec.pop("_detail_url", "")
    if not detail_url:
        return

    try:
        r = requests.get(detail_url, headers=HEADERS, timeout=12)
        r.encoding = r.apparent_encoding or "utf-8"
        soup = BeautifulSoup(r.text, "lxml")
    except Exception as e:
        logger.debug(f"詳細取得失敗: {detail_url} / {e}")
        return

    for row in soup.select("table tr"):
        cells = row.find_all(["th", "td"])
        if len(cells) < 2:
            continue
        key = cells[0].get_text(strip=True)
        val = cells[1].get_text(strip=True)

        if key == "施設名称" and not rec.get("名称"):
            rec["名称"] = val
        elif key == "所在地":
            addr = re.sub(r"〒\s*\d{3}[-－]\d{4}", "", val).strip()
            rec["所在地"] = addr
        elif key == "TEL":
            phone = re.sub(r"[^0-9\-]", "", val)
            rec["電話番号"] = phone
    time.sleep(random.uniform(0.1, 0.3))


async def fetch_judo_details_concurrent(records: list[dict], sem_count: int = 10) -> None:
    """詳細ページを並列 sem_count 接続で取得する。"""
    semaphore = asyncio.Semaphore(sem_count)
    total = len(records)
    done = {"n": 0}

    async def _one(rec: dict) -> None:
        async with semaphore:
            await asyncio.to_thread(_fetch_detail_sync, rec)
            done["n"] += 1
            if done["n"] % 50 == 0 or done["n"] == total:
                logger.info(f"  詳細取得: {done['n']}/{total}")

    await asyncio.gather(*[_one(rec) for rec in records])


# ── Phase 2: hotpepper.jp ────────────────────────────────

async def scrape_hotpepper(page) -> list[dict]:
    """hotpepper から宮城県の接骨院・整骨院リストを収集する。"""
    records: list[dict] = []
    url = HOTPEPPER_LIST_URL
    page_num = 1

    while True:
        try:
            await page.goto(url, timeout=20000, wait_until="networkidle")
            await page.wait_for_timeout(2000)
        except Exception as e:
            logger.warning(f"hotpepper ページ取得失敗 {url}: {e}")
            break

        html = await page.content()
        soup = BeautifulSoup(html, "lxml")
        items = soup.select("li.searchListCassette")
        if not items:
            break

        for item in items:
            a = item.find("a", href=True)
            if not a:
                continue
            name_el = item.find(["h2", "h3", "strong", "p"])
            name = a.get_text(strip=True) if not name_el else name_el.get_text(strip=True)
            # リスト内の住所テキストを探す
            addr_el = item.find(class_=re.compile("address|addr"))
            addr = addr_el.get_text(strip=True) if addr_el else ""
            hp_url = a["href"]
            if not hp_url.startswith("http"):
                hp_url = urljoin(HOTPEPPER_BASE, hp_url)

            rec = {k: "" for k in OUTPUT_COLS}
            rec["名称"] = name
            rec["所在地"] = addr
            rec["_hp_url"] = hp_url
            records.append(rec)

        logger.info(f"  hotpepper p{page_num}: {len(items)} 件")

        # 次ページ
        next_a = soup.find("a", string=re.compile(r"次へ|>|Next", re.I))
        if not next_a:
            # ページネーション番号で次ページを探す
            pager = soup.select(".pagination a, .pagerArea a")
            next_href = None
            for pa in pager:
                t = pa.get_text(strip=True)
                if t == str(page_num + 1):
                    next_href = pa.get("href")
                    break
            if not next_href:
                break
            url = next_href if next_href.startswith("http") else urljoin(HOTPEPPER_BASE, next_href)
        else:
            href = next_a.get("href", "")
            url = href if href.startswith("http") else urljoin(HOTPEPPER_BASE, href)

        page_num += 1
        if page_num > 10:
            break
        await asyncio.sleep(random.uniform(1.5, 2.5))

    # hotpepper 詳細ページから住所を取得 (address が空のもの)
    for rec in records:
        hp_url = rec.pop("_hp_url", "")
        if rec.get("所在地") or not hp_url:
            continue
        try:
            await page.goto(hp_url, timeout=15000, wait_until="domcontentloaded")
            await page.wait_for_timeout(1000)
            html2 = await page.content()
            soup2 = BeautifulSoup(html2, "lxml")
            # 住所抽出
            for row in soup2.select("table tr"):
                th = row.find("th")
                td = row.find("td")
                if th and td and "住所" in th.get_text():
                    rec["所在地"] = td.get_text(strip=True)
                    break
        except Exception:
            pass
        await asyncio.sleep(random.uniform(0.8, 1.5))

    logger.info(f"hotpepper 合計: {len(records)} 件")
    return records


# ── Phase 3: DuckDuckGo 公式サイト検索 ──────────────────

def search_official_site(name: str, address: str) -> str:
    """DuckDuckGo でクリニックの公式サイトを検索して返す。"""
    # 所在地から市区町村を抽出
    city_match = re.search(r"宮城県\s*([^\s市区町村]{2,5}[市区町村])", address)
    city = city_match.group(1) if city_match else "宮城"
    query = f"{name} {city} 公式サイト"

    try:
        with DDGS() as ddgs:
            results = list(ddgs.text(query, max_results=5))
        for r in results:
            url = r.get("href", "")
            if not url:
                continue
            if any(d in url for d in SKIP_DOMAINS):
                continue
            if url.startswith("http"):
                return url
    except Exception as e:
        logger.debug(f"DDG 検索失敗 {name}: {e}")
    return ""


# ── Phase 4: 公式サイト追加情報 ──────────────────────────

def _decode_cfemail(encoded: str) -> str:
    try:
        key = int(encoded[:2], 16)
        return bytes(
            int(encoded[i:i+2], 16) ^ key for i in range(2, len(encoded), 2)
        ).decode("utf-8")
    except Exception:
        return ""


def _parse_official_html(url: str, html: str) -> dict:
    """HTML からメール・インスタ・問い合わせURLを抽出する（同期）。"""
    result: dict = {}
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


def _fetch_official_sync(url: str) -> dict:
    """requests で公式サイトを取得してパースする（同期・スレッド実行用）。"""
    if any(d in url for d in SKIP_DOMAINS):
        return {}
    if not _is_allowed(url):
        return {}
    try:
        r = requests.get(url, headers=HEADERS, timeout=10, allow_redirects=True)
        if r.status_code >= 400:
            return {}
        r.encoding = r.apparent_encoding or "utf-8"
        return _parse_official_html(url, r.text)
    except Exception:
        return {}


async def scrape_official_sites_concurrent(records: list[dict], sem_count: int = 10) -> None:
    """公式サイトを並列 sem_count で取得する（requests + asyncio.to_thread）。"""
    semaphore = asyncio.Semaphore(sem_count)
    total = len(records)
    done = {"n": 0}

    async def _one(rec: dict) -> None:
        async with semaphore:
            url = rec.get("公式サイトURL", "")
            if url:
                extras = await asyncio.to_thread(_fetch_official_sync, url)
                rec.update(extras)
            done["n"] += 1
            if done["n"] % 30 == 0 or done["n"] == total:
                logger.info(f"  公式サイト取得: {done['n']}/{total}")

    await asyncio.gather(*[_one(rec) for rec in records])


# ── メイン ────────────────────────────────────────────────

async def main() -> None:
    logger.info("=" * 65)
    logger.info("宮城県 整体・接骨院 収集開始")
    logger.info("=" * 65)
    start_time = datetime.now()

    # ── Phase 1: judo-ch.jp ──────────────────────────────
    logger.info("Phase 1: judo-ch.jp から宮城県接骨院収集")
    area_urls = get_area_urls()
    all_records: list[dict] = []
    seen_names: set[str] = set()
    seen_phones: set[str] = set()

    for i, (area_name, area_url) in enumerate(area_urls, 1):
        recs = scrape_judo_area(area_url)
        new_recs: list[dict] = []
        for rec in recs:
            nk = _normalize_name(rec["名称"])
            if nk and nk not in seen_names:
                seen_names.add(nk)
                new_recs.append(rec)
        all_records.extend(new_recs)
        logger.info(f"  [{i}/{len(area_urls)}] {area_name}: {len(recs)} 件 (新規 {len(new_recs)} 件, 累計 {len(all_records)} 件)")
        time.sleep(random.uniform(0.5, 1.0))

    # 詳細ページから住所・電話を並列取得
    logger.info(f"Phase 1B: 詳細ページ取得 ({len(all_records)} 件, 並列10)")
    await fetch_judo_details_concurrent(all_records, sem_count=10)
    # 電話番号を seen_phones に追加
    for rec in all_records:
        pk = _normalize_phone(rec.get("電話番号", ""))
        if pk:
            seen_phones.add(pk)

    logger.info(f"Phase 1 完了: {len(all_records)} 件")

    # ── Phase 2: hotpepper ───────────────────────────────
    logger.info("Phase 2: hotpepper から追加収集")
    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=True)
        context = await browser.new_context(user_agent=USER_AGENT)
        hp_page = await context.new_page()

        hp_records = await scrape_hotpepper(hp_page)
        hp_added = 0
        for rec in hp_records:
            nk = _normalize_name(rec.get("名称", ""))
            pk = _normalize_phone(rec.get("電話番号", ""))
            if nk and nk not in seen_names:
                if not pk or pk not in seen_phones:
                    seen_names.add(nk)
                    if pk:
                        seen_phones.add(pk)
                    all_records.append(rec)
                    hp_added += 1
        logger.info(f"Phase 2 完了: 新規追加 {hp_added} 件 (合計 {len(all_records)} 件)")

        # ── Phase 3: DuckDuckGo 公式サイト検索 ────────────
        logger.info("Phase 3: DuckDuckGo で公式サイト検索")
        no_url = [r for r in all_records if not r.get("公式サイトURL")]
        logger.info(f"  公式URL 未取得: {len(no_url)} 件")

        for i, rec in enumerate(no_url, 1):
            url = search_official_site(rec["名称"], rec.get("所在地", ""))
            if url:
                rec["公式サイトURL"] = url
            if i % 30 == 0 or i == len(no_url):
                found = sum(1 for r in no_url[:i] if r.get("公式サイトURL"))
                logger.info(f"  DDG検索: {i}/{len(no_url)} (発見 {found} 件)")
            time.sleep(random.uniform(1.5, 2.5))

        # ── Phase 4: 公式サイト追加情報取得 ────────────────
        logger.info("Phase 4: 公式サイト収集")
        with_url = [r for r in all_records if r.get("公式サイトURL")]
        logger.info(f"  公式URL あり: {len(with_url)} / {len(all_records)} 件")
        await scrape_official_sites_concurrent(with_url, sem_count=10)

        await browser.close()

    # ── Excel 出力 ────────────────────────────────────────
    df = pd.DataFrame(all_records, columns=OUTPUT_COLS)
    df.drop_duplicates(subset=["名称", "電話番号"], keep="first", inplace=True)
    df.drop_duplicates(subset=["名称"], keep="first", inplace=True)
    df.to_excel(OUTPUT_FILE, index=False)

    elapsed = int((datetime.now() - start_time).total_seconds())
    logger.info("=" * 65)
    logger.info(f"完了: {OUTPUT_FILE} に {len(df)} 件を出力")
    logger.info(f"所要時間: {elapsed // 60}分{elapsed % 60}秒")
    logger.info("=" * 65)


if __name__ == "__main__":
    asyncio.run(main())
