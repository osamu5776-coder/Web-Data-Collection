"""
宮城県 整体・接骨院 クイック出力（名称・所在地・電話のみ）

Phase 1: judo-ch.jp (478件, 並列10)
Phase 2: hotpepper (約52件追加)
→ 即座に Excel 出力（DuckDuckGo・公式サイト収集なし）
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

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

LOG_FILE = "scraper_miyagi_quick.log"
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

OUTPUT_COLS = [
    "名称", "メールアドレス", "公式サイトURL",
    "所在地", "電話番号", "インスタURL", "問い合わせフォームURL",
]


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


def get_area_urls() -> list[tuple[str, str]]:
    soup = _fetch(JUDO_PREF_URL)
    if soup is None:
        return []
    seen: set[str] = set()
    results: list[tuple[str, str]] = []
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if re.search(r"/sekkotsuinsrch/04/\d{5}/$", href):
            if href not in seen:
                seen.add(href)
                full = href if href.startswith("http") else urljoin(JUDO_BASE, href)
                results.append((a.get_text(strip=True), full))
    logger.info(f"judo-ch.jp: {len(results)} エリア")
    return results


def scrape_judo_area(area_url: str) -> list[dict]:
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
                if "syuhenshisetsulist" in detail_url:
                    continue
                rec = {k: "" for k in OUTPUT_COLS}
                rec["名称"] = name
                rec["_detail_url"] = detail_url
                records.append(rec)
        next_link = soup.find("a", string="次のページ")
        if not next_link:
            break
        page_num += 1
        if page_num > 20:
            break
        time.sleep(random.uniform(0.5, 1.0))
    return records


def _fetch_detail_sync(rec: dict) -> None:
    detail_url = rec.pop("_detail_url", "")
    if not detail_url:
        return
    try:
        r = requests.get(detail_url, headers=HEADERS, timeout=12)
        r.encoding = r.apparent_encoding or "utf-8"
        soup = BeautifulSoup(r.text, "lxml")
    except Exception:
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
            rec["所在地"] = re.sub(r"〒\s*\d{3}[-－]\d{4}", "", val).strip()
        elif key == "TEL":
            rec["電話番号"] = re.sub(r"[^0-9\-]", "", val)
    time.sleep(random.uniform(0.1, 0.3))


async def fetch_judo_details_concurrent(records: list[dict], sem_count: int = 10) -> None:
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


async def scrape_hotpepper(page) -> list[dict]:
    records: list[dict] = []
    url = HOTPEPPER_LIST_URL
    page_num = 1
    while True:
        try:
            await page.goto(url, timeout=20000, wait_until="networkidle")
            await page.wait_for_timeout(2000)
        except Exception as e:
            logger.warning(f"hotpepper 失敗 {url}: {e}")
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
            addr_el = item.find(class_=re.compile("address|addr"))
            addr = addr_el.get_text(strip=True) if addr_el else ""
            rec = {k: "" for k in OUTPUT_COLS}
            rec["名称"] = name
            rec["所在地"] = addr
            records.append(rec)
        logger.info(f"  hotpepper p{page_num}: {len(items)} 件")
        pager = soup.select(".pagination a, .pagerArea a")
        next_href = None
        for pa in pager:
            if pa.get_text(strip=True) == str(page_num + 1):
                next_href = pa.get("href")
                break
        if not next_href:
            break
        url = next_href if next_href.startswith("http") else urljoin(HOTPEPPER_BASE, next_href)
        page_num += 1
        if page_num > 10:
            break
        await asyncio.sleep(random.uniform(1.5, 2.5))
    logger.info(f"hotpepper 合計: {len(records)} 件")
    return records


async def main() -> None:
    logger.info("=" * 65)
    logger.info("宮城県 整体・接骨院 クイック収集開始")
    logger.info("=" * 65)
    start_time = datetime.now()

    # Phase 1: judo-ch.jp
    logger.info("Phase 1: judo-ch.jp から宮城県接骨院収集")
    area_urls = get_area_urls()
    all_records: list[dict] = []
    seen_names: set[str] = set()
    seen_phones: set[str] = set()

    for i, (area_name, area_url) in enumerate(area_urls, 1):
        recs = scrape_judo_area(area_url)
        new_recs = []
        for rec in recs:
            nk = _normalize_name(rec["名称"])
            if nk and nk not in seen_names:
                seen_names.add(nk)
                new_recs.append(rec)
        all_records.extend(new_recs)
        logger.info(f"  [{i}/{len(area_urls)}] {area_name}: {len(recs)} 件 (新規 {len(new_recs)} 件, 累計 {len(all_records)} 件)")
        time.sleep(random.uniform(0.5, 1.0))

    logger.info(f"Phase 1B: 詳細ページ取得 ({len(all_records)} 件, 並列10)")
    await fetch_judo_details_concurrent(all_records, sem_count=10)
    for rec in all_records:
        pk = _normalize_phone(rec.get("電話番号", ""))
        if pk:
            seen_phones.add(pk)
    logger.info(f"Phase 1 完了: {len(all_records)} 件")

    # Phase 2: hotpepper
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
        await browser.close()
    logger.info(f"Phase 2 完了: 新規追加 {hp_added} 件 (合計 {len(all_records)} 件)")

    # Excel 出力
    df = pd.DataFrame(all_records, columns=OUTPUT_COLS)
    df.drop_duplicates(subset=["名称"], keep="first", inplace=True)
    df.to_excel(OUTPUT_FILE, index=False)

    elapsed = int((datetime.now() - start_time).total_seconds())
    logger.info("=" * 65)
    logger.info(f"完了: {OUTPUT_FILE} に {len(df)} 件を出力")
    logger.info(f"所要時間: {elapsed // 60}分{elapsed % 60}秒")
    logger.info("=" * 65)


if __name__ == "__main__":
    asyncio.run(main())
