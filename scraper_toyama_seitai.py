"""
富山県 整体・接骨院 情報収集

データソース:
  Phase 1 - judo-ch.jp から富山県全市区町村の接骨院を収集
            一覧ページ: 名称・住所
            詳細ページ (10並列): 公式サイトURL
  Phase 2 - body-care.expert から追加情報収集
            名称・住所・電話・公式URL (5件)
  Phase 3 - chikuchikuryoho.com から富山県分を収集
            名称・住所・電話・メール/URL (5件)
  Phase 4 - DuckDuckGo で公式サイト検索 (3並列)
  Phase 5 - 公式サイトからメール・インスタ・問い合わせURL取得 (10並列 requests)

出力列: 名称, メールアドレス, 公式サイトURL, 所在地, 電話番号, インスタURL, 問い合わせフォームURL
出力ファイル: toyama_seitai_YYYYMMDD_HHMMSS.xlsx
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

# ── ロギング ──────────────────────────────────────────────
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

LOG_FILE = "scraper_toyama_seitai.log"
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
OUTPUT_FILE = f"toyama_seitai_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/124.0.0.0 Safari/537.36"
)
HEADERS = {"User-Agent": USER_AGENT}

JUDO_BASE = "https://www.judo-ch.jp"
JUDO_PREF_URL = f"{JUDO_BASE}/sekkotsuinsrch/16/"

BODYCARE_URL = "https://www.body-care.expert/toyama/"
CHIKU_URL = "https://chikuchikuryoho.com/chiryoinlist"

SKIP_DOMAINS = [
    "judo-ch.jp", "homemate-research.com", "hotpepper.jp",
    "body-care.expert", "chikuchikuryoho.com",
    "instagram.com", "facebook.com", "twitter.com", "x.com", "youtube.com",
    "tiktok.com", "wikipedia.org", "google.com", "bing.com",
    "mapion.co.jp", "itp.ne.jp", "ekiten.jp", "salonboard.com",
    "beauty.rakuten.co.jp", "enma.jp", "seikotsuin.info",
    "medley.life", "caloo.jp", "toresei.com",
]

OUTPUT_COLS = [
    "名称", "メールアドレス", "公式サイトURL",
    "所在地", "電話番号", "インスタURL", "問い合わせフォームURL",
]

# ── 共通ユーティリティ ────────────────────────────────────
_robots_cache: dict[str, RobotFileParser] = {}


def _is_allowed(url: str) -> bool:
    from urllib.parse import urlparse
    parsed = urlparse(url)
    origin = f"{parsed.scheme}://{parsed.netloc}"
    if origin not in _robots_cache:
        rp = RobotFileParser()
        try:
            rp.set_url(f"{origin}/robots.txt")
            rp.read()
        except Exception:
            pass
        _robots_cache[origin] = rp
    return _robots_cache[origin].can_fetch(USER_AGENT, url)


def _fetch(url: str, timeout: int = 12) -> BeautifulSoup | None:
    try:
        r = requests.get(url, headers=HEADERS, timeout=timeout)
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

def get_judo_area_urls() -> list[tuple[str, str]]:
    """富山県の全エリアURL (エリア名, URL) を返す。"""
    soup = _fetch(JUDO_PREF_URL)
    if soup is None:
        return []
    seen: set[str] = set()
    results: list[tuple[str, str]] = []
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if re.search(r"/sekkotsuinsrch/16/\d{5}/$", href):
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
            if not a or "/sekkotsuinsrch/" not in a["href"]:
                continue
            name = h3.get_text(strip=True)
            detail_href = a["href"]
            detail_url = detail_href if detail_href.startswith("http") else urljoin(JUDO_BASE, detail_href)
            if "syuhenshisetsulist" in detail_url:
                continue

            # 同じ li 内の住所テーブルを取得
            parent = h3.find_parent("li") or h3.find_parent("dd") or h3
            address = ""
            for tr in parent.find_all("tr"):
                th = tr.find("th")
                td = tr.find("td")
                if th and td and "所在地" in th.get_text():
                    address = re.sub(r"〒\s*\d{3}[-－]\d{4}\s*", "", td.get_text(strip=True)).strip()
                    break

            rec = {k: "" for k in OUTPUT_COLS}
            rec["名称"] = name
            rec["所在地"] = address
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


def _fetch_judo_detail_sync(rec: dict) -> None:
    """judo-ch.jp 詳細ページから公式URLを取得して rec を更新する (同期版)。"""
    detail_url = rec.pop("_detail_url", "")
    if not detail_url:
        return
    try:
        r = requests.get(detail_url, headers=HEADERS, timeout=12)
        r.encoding = r.apparent_encoding or "utf-8"
        soup = BeautifulSoup(r.text, "lxml")
    except Exception:
        return

    for tr in soup.select("table tr"):
        cells = tr.find_all(["th", "td"])
        if len(cells) < 2:
            continue
        key = cells[0].get_text(strip=True)
        if key == "所在地" and not rec.get("所在地"):
            rec["所在地"] = re.sub(r"〒\s*\d{3}[-－]\d{4}\s*", "", cells[1].get_text(strip=True)).strip()
        elif key in ("URL", "ホームページ", "公式サイト", "WEB"):
            link = cells[1].find("a", href=True)
            if link:
                href = link["href"].strip()
                if href.startswith("http") and not any(d in href for d in SKIP_DOMAINS):
                    rec["公式サイトURL"] = href
    time.sleep(random.uniform(0.1, 0.3))


async def fetch_judo_details_concurrent(records: list[dict], sem_count: int = 10) -> None:
    """詳細ページを並列 sem_count 接続で取得する。"""
    semaphore = asyncio.Semaphore(sem_count)
    total = len(records)
    done = {"n": 0}

    async def _one(rec: dict) -> None:
        async with semaphore:
            await asyncio.to_thread(_fetch_judo_detail_sync, rec)
            done["n"] += 1
            if done["n"] % 50 == 0 or done["n"] == total:
                logger.info(f"  詳細取得: {done['n']}/{total}")

    await asyncio.gather(*[_one(rec) for rec in records])


# ── Phase 2: body-care.expert ─────────────────────────────

def scrape_bodycare() -> list[dict]:
    """body-care.expert/toyama/ から接骨院情報を収集する。"""
    soup = _fetch(BODYCARE_URL)
    if soup is None:
        return []
    records: list[dict] = []
    for li in soup.select("li"):
        h2 = li.find("h2")
        if not h2:
            continue
        a_name = h2.find("a", href=True)
        if not a_name:
            continue
        name = a_name.get_text(strip=True)

        # 住所・電話 (p タグ)
        paras = li.find_all("p")
        address = paras[0].get_text(strip=True).strip() if len(paras) > 0 else ""
        phone = paras[1].get_text(strip=True).strip() if len(paras) > 1 else ""
        phone = re.sub(r"[^0-9\-]", "", phone)

        # 公式URL (body-care.expert 以外の2番目のリンク)
        official_url = ""
        for a in li.find_all("a", href=True):
            href = a["href"]
            if href.startswith("http") and "body-care.expert" not in href:
                official_url = href
                break

        rec = {k: "" for k in OUTPUT_COLS}
        rec["名称"] = name
        rec["所在地"] = address
        rec["電話番号"] = phone
        rec["公式サイトURL"] = official_url
        records.append(rec)

    logger.info(f"body-care.expert: {len(records)} 件")
    return records


# ── Phase 3: chikuchikuryoho.com ─────────────────────────

def scrape_chiku() -> list[dict]:
    """chikuchikuryoho.com から富山県の接骨院情報を収集する。"""
    soup = _fetch(CHIKU_URL)
    if soup is None:
        return []

    records: list[dict] = []
    in_toyama = False
    for tag in soup.find_all(["h3", "article"]):
        if tag.name == "h3":
            text = tag.get_text(strip=True)
            in_toyama = "富山" in text
            continue
        if not in_toyama or tag.name != "article":
            continue

        p = tag.find("p")
        if not p:
            continue
        full_text = p.get_text(separator="\n", strip=True)

        name_el = p.find("strong")
        name = name_el.get_text(strip=True) if name_el else ""
        if not name:
            continue

        # テキストから住所・電話を抽出
        lines = [ln.strip() for ln in full_text.split("\n") if ln.strip()]
        address, phone = "", ""
        for ln in lines:
            if re.search(r"富山県", ln) and not address:
                address = ln
            m = re.search(r"電話[：:]\s*([\d\-]+)", ln)
            if m:
                phone = m.group(1)

        # メール (mailto:) または公式URL
        email, official_url = "", ""
        for a in p.find_all("a", href=True):
            href = a["href"]
            if href.startswith("mailto:"):
                email = href[7:]
            elif href.startswith("http") and not any(d in href for d in SKIP_DOMAINS):
                official_url = href

        rec = {k: "" for k in OUTPUT_COLS}
        rec["名称"] = name
        rec["所在地"] = address
        rec["電話番号"] = phone
        rec["メールアドレス"] = email
        rec["公式サイトURL"] = official_url
        records.append(rec)

    logger.info(f"chikuchikuryoho.com 富山: {len(records)} 件")
    return records


# ── Phase 4: DuckDuckGo 公式サイト検索 ──────────────────

def search_official_site(name: str, address: str) -> str:
    """DuckDuckGo でクリニックの公式サイトを検索して返す。"""
    city_match = re.search(r"富山県\s*([^\s市区町村]{2,5}[市区町村])", address)
    city = city_match.group(1) if city_match else "富山"
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
        logger.debug(f"DDG失敗 {name}: {e}")
    return ""


# ── Phase 5: 公式サイト追加情報 ──────────────────────────

def _decode_cfemail(encoded: str) -> str:
    try:
        key = int(encoded[:2], 16)
        return bytes(
            int(encoded[i:i+2], 16) ^ key for i in range(2, len(encoded), 2)
        ).decode("utf-8")
    except Exception:
        return ""


def _parse_official_html(url: str, html: str) -> dict:
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

    # 電話番号 (まだ取得できていない場合)
    if not result.get("電話番号"):
        phones = re.findall(r"0\d{1,4}[-‐‑‒–—－]\d{1,4}[-‐‑‒–—－]\d{3,4}", text)
        if phones:
            result["電話番号"] = re.sub(r"[^\d\-]", "", phones[0])

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
    """公式サイトを並列 sem_count で取得する。"""
    semaphore = asyncio.Semaphore(sem_count)
    total = len(records)
    done = {"n": 0}

    async def _one(rec: dict) -> None:
        async with semaphore:
            url = str(rec.get("公式サイトURL", ""))
            if url.startswith("http"):
                extras = await asyncio.to_thread(_fetch_official_sync, url)
                # 既存のメールは上書きしない
                if rec.get("メールアドレス") and "メールアドレス" in extras:
                    del extras["メールアドレス"]
                # 既存の電話番号は上書きしない
                if rec.get("電話番号") and "電話番号" in extras:
                    del extras["電話番号"]
                rec.update(extras)
            done["n"] += 1
            if done["n"] % 50 == 0 or done["n"] == total:
                logger.info(f"  公式サイト取得: {done['n']}/{total}")

    await asyncio.gather(*[_one(rec) for rec in records])


# ── メイン ────────────────────────────────────────────────

async def main() -> None:
    logger.info("=" * 65)
    logger.info("富山県 整体・接骨院 収集開始")
    logger.info("=" * 65)
    start_time = datetime.now()

    all_records: list[dict] = []
    seen_names: set[str] = set()
    seen_phones: set[str] = set()

    # ── Phase 1: judo-ch.jp ──────────────────────────────
    logger.info("Phase 1: judo-ch.jp から富山県接骨院収集")
    area_urls = get_judo_area_urls()

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
    logger.info(f"Phase 1 完了: {len(all_records)} 件 (URL取得: {sum(1 for r in all_records if r.get('公式サイトURL'))} 件)")

    # ── Phase 2: body-care.expert ─────────────────────────
    logger.info("Phase 2: body-care.expert から追加収集")
    bc_records = scrape_bodycare()
    bc_added = 0
    for rec in bc_records:
        nk = _normalize_name(rec["名称"])
        pk = _normalize_phone(rec.get("電話番号", ""))
        if nk in seen_names:
            # 既存レコードに電話番号や URL を補完
            for existing in all_records:
                if _normalize_name(existing["名称"]) == nk:
                    if not existing.get("電話番号") and rec.get("電話番号"):
                        existing["電話番号"] = rec["電話番号"]
                    if not existing.get("公式サイトURL") and rec.get("公式サイトURL"):
                        existing["公式サイトURL"] = rec["公式サイトURL"]
                    break
            continue
        if pk and pk in seen_phones:
            continue
        seen_names.add(nk)
        if pk:
            seen_phones.add(pk)
        all_records.append(rec)
        bc_added += 1
    logger.info(f"Phase 2 完了: 新規追加 {bc_added} 件 (合計 {len(all_records)} 件)")

    # ── Phase 3: chikuchikuryoho.com ─────────────────────
    logger.info("Phase 3: chikuchikuryoho.com から富山県分収集")
    chiku_records = scrape_chiku()
    chiku_added = 0
    for rec in chiku_records:
        nk = _normalize_name(rec["名称"])
        pk = _normalize_phone(rec.get("電話番号", ""))
        if nk in seen_names:
            for existing in all_records:
                if _normalize_name(existing["名称"]) == nk:
                    if not existing.get("電話番号") and rec.get("電話番号"):
                        existing["電話番号"] = rec["電話番号"]
                    if not existing.get("メールアドレス") and rec.get("メールアドレス"):
                        existing["メールアドレス"] = rec["メールアドレス"]
                    if not existing.get("公式サイトURL") and rec.get("公式サイトURL"):
                        existing["公式サイトURL"] = rec["公式サイトURL"]
                    break
            continue
        if pk and pk in seen_phones:
            continue
        seen_names.add(nk)
        if pk:
            seen_phones.add(pk)
        all_records.append(rec)
        chiku_added += 1
    logger.info(f"Phase 3 完了: 新規追加 {chiku_added} 件 (合計 {len(all_records)} 件)")

    # ── Phase 4: DuckDuckGo 公式サイト検索 (3並列) ───────
    logger.info("Phase 4: DuckDuckGo で公式サイト検索")
    no_url = [r for r in all_records if not r.get("公式サイトURL")]
    logger.info(f"  公式URL 未取得: {len(no_url)} 件 (取得済み: {len(all_records) - len(no_url)} 件)")

    ddg_sem = asyncio.Semaphore(3)
    ddg_done = {"n": 0, "found": 0}

    async def _ddg_one(rec: dict) -> None:
        async with ddg_sem:
            url = await asyncio.to_thread(
                search_official_site, rec["名称"], rec.get("所在地", "")
            )
            if url:
                rec["公式サイトURL"] = url
                ddg_done["found"] += 1
            ddg_done["n"] += 1
            n = ddg_done["n"]
            if n % 30 == 0 or n == len(no_url):
                logger.info(f"  DDG検索: {n}/{len(no_url)} (発見 {ddg_done['found']} 件)")
            await asyncio.sleep(random.uniform(1.0, 2.0))

    await asyncio.gather(*[_ddg_one(rec) for rec in no_url])
    logger.info(f"Phase 4 完了: 公式URL 合計 {sum(1 for r in all_records if r.get('公式サイトURL'))} 件")

    # ── Phase 5: 公式サイト収集 (10並列) ──────────────────
    logger.info("Phase 5: 公式サイト収集")
    with_url = [r for r in all_records if r.get("公式サイトURL")]
    logger.info(f"  公式URL あり: {len(with_url)} / {len(all_records)} 件")
    await scrape_official_sites_concurrent(with_url, sem_count=10)
    logger.info("Phase 5 完了")

    # ── Excel 出力 ────────────────────────────────────────
    df = pd.DataFrame(all_records, columns=OUTPUT_COLS)
    df.drop_duplicates(subset=["名称"], keep="first", inplace=True)
    df = df.reset_index(drop=True)
    df.to_excel(OUTPUT_FILE, index=False)

    elapsed = int((datetime.now() - start_time).total_seconds())
    logger.info("=" * 65)
    logger.info(f"完了: {OUTPUT_FILE} に {len(df)} 件を出力")
    logger.info(f"所要時間: {elapsed // 60}分{elapsed % 60}秒")
    logger.info("=" * 65)


if __name__ == "__main__":
    asyncio.run(main())
