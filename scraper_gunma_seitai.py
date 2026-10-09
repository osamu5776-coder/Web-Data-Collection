"""
群馬県 整体・接骨院 情報収集

データソース:
  Phase 1 - judo-ch.jp/sekkotsuinsrch/10/list/ から群馬県の接骨院・整骨院を全ページ収集
            一覧ページ (ページネーション全件): 名称・住所
            詳細ページ (10並列): 公式サイトURL
  Phase 2 - body-care.expert/gunma/ から Pick Up 接骨院・整骨院を収集
            名称・住所・電話・公式URL
  Phase 3 - chikuchikuryoho.com/chiryoinlist#gunma から群馬県分を収集
            名称・住所・電話・メール/URL
  Phase 4 - DuckDuckGo で公式サイト未取得分を検索 (3並列)
  Phase 5 - 公式サイトからメール・インスタ・問い合わせURL取得 (10並列 requests)

出力列: 名称, メールアドレス, 公式サイトURL, 所在地, 電話番号, インスタURL, 問い合わせフォームURL
出力ファイル: gunma_seitai_YYYYMMDD_HHMMSS.xlsx
"""

import asyncio
import io
import logging
import random
import re
import sys
import time
from datetime import datetime
from urllib.parse import urljoin, urlparse
from urllib.robotparser import RobotFileParser

import pandas as pd
import requests
from bs4 import BeautifulSoup
from ddgs import DDGS

# ── ロギング ──────────────────────────────────────────────
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

LOG_FILE = "scraper_gunma_seitai.log"
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
OUTPUT_FILE = f"gunma_seitai_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/124.0.0.0 Safari/537.36"
)
HEADERS = {"User-Agent": USER_AGENT}

JUDO_BASE = "https://www.judo-ch.jp"
JUDO_LIST_URL = "https://www.judo-ch.jp/sekkotsuinsrch/10/list/"

BODYCARE_URL = "https://body-care.expert/gunma/"
CHIKU_URL = "https://chikuchikuryoho.com/chiryoinlist"

SKIP_DOMAINS = [
    # データソース自身
    "judo-ch.jp", "homemate-research.com", "homemate-research-station.com",
    "homemate-research-bus.com", "body-care.expert", "chikuchikuryoho.com",
    "maps.homemate-research.com",
    # LINE
    "line.me", "lin.ee",
    # 総合プラットフォーム・予約・クーポン
    "hotpepper.jp", "beauty.hotpepper.jp", "epark.jp", "curama.jp",
    "reserva.be", "coubic.com", "salonboard.com", "minimo.jp",
    "rakuten.co.jp", "beauty.rakuten.co.jp",
    # 口コミサイト
    "ekiten.jp", "tabelog.com", "caloo.jp", "minkabu.jp",
    "google.com", "google.co.jp", "maps.google.com",
    # 情報・ディレクトリサイト
    "itp.ne.jp", "mapion.co.jp", "navitime.co.jp", "mapfan.com",
    "goo.ne.jp", "wikipedia.org", "homemate-research.com",
    "seikotsuin.info", "wellyou.net", "judi-channel.jp",
    "medley.life", "toresei.com", "seitai-rct.jp", "enma.jp",
    "seikotsuguide.jp", "health-more.jp", "bonbonesquare.com",
    "minnanochiryoin.jp", "karadarefre.jp", "jiko24.jp",
    "seikotsuin-navi.com", "seitai-osusume-select.com",
    "treasuredtime.org", "findglocal.com", "mypl.net", "jusei.gr.jp",
    "ozmall.co.jp", "biglobe.ne.jp", "town.or.jp",
    "koutsujiko-chiryo-biz.jp", "chiryo-biz.jp", "self-whitening.jp",
    "tsunagu-good.com", "jiko-navi.jp", "passpon.jp", "everytown.info",
    "jpaweb.jp", "ameblo.jp", "learning-with.us",
    "seikotsu-navi.net", "reserven.jp", "salonnavigation.com",
    # SNS
    "instagram.com", "facebook.com", "twitter.com", "x.com",
    "youtube.com", "tiktok.com",
    # 検索エンジン
    "bing.com", "yahoo.co.jp",
]

# 官公庁ドメイン（誤マッチ除外）
GOV_DOMAIN_RE = re.compile(r"\.lg\.jp|^https?://(www\.)?city\.|//city\.", re.I)

OUTPUT_COLS = [
    "名称", "メールアドレス", "公式サイトURL",
    "所在地", "電話番号", "インスタURL", "問い合わせフォームURL",
]

# ── 共通ユーティリティ ────────────────────────────────────
_robots_cache: dict[str, RobotFileParser] = {}


def _is_allowed(url: str) -> bool:
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


def _is_skip_domain(url: str) -> bool:
    if GOV_DOMAIN_RE.search(url):
        return True
    return any(d in url for d in SKIP_DOMAINS)


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


def _verify_match(html_text: str, name: str, phone: str, address: str = "") -> bool:
    """取得した公式サイトが実際にその院のものかを名称・電話番号・所在地で検証する。

    「姓+整骨院/接骨院」のような名称は全国どこにでも存在しうるため、名称一致
    だけでは同名の無関係な別事業者を誤って採用してしまう。電話番号の完全一致
    がない限り、名称一致に加えて住所の市区町村名がページ内に含まれることを
    必須とする。
    """
    norm_page = _normalize_name(re.sub(r"[\s　]+", "", html_text))

    phone_digits = _normalize_phone(phone)
    if phone_digits and phone_digits in re.sub(r"[^0-9]", "", html_text):
        return True

    norm_name = _normalize_name(name)
    name_hit = bool(norm_name) and norm_name in norm_page

    core = re.sub(
        r"(整骨院|接骨院|整体院|治療院|鍼灸院|鍼灸整骨院|はりきゅう整骨院|"
        r"整体|鍼灸|マッサージ|分院)+$",
        "", name
    ).strip()
    core_norm = _normalize_name(core)
    core_hit = len(core_norm) >= 2 and core_norm in norm_page

    if not (name_hit or core_hit):
        return False

    city_match = re.search(r"([^\s　]{2,8}?[市区町村])", address)
    city = city_match.group(1) if city_match else ""
    return bool(city) and _normalize_name(city) in norm_page


# ── Phase 1: judo-ch.jp ──────────────────────────────────

def get_total_pages() -> int:
    """群馬県エリアの総件数から総ページ数を算出する。"""
    soup = _fetch(JUDO_LIST_URL)
    if soup is None:
        return 1
    count_el = soup.find("span", class_="count")
    if count_el:
        m = re.search(r"/(\d+)院", count_el.get_text())
        if m:
            total = int(m.group(1))
            pages = (total + 29) // 30
            logger.info(f"judo-ch.jp: 総件数 {total} 院 / {pages} ページ")
            return pages
    # フォールバック: ページネーションの最大番号
    max_page = 1
    for a in soup.find_all("a", href=True):
        m = re.search(r"/sekkotsuinsrch/10/list/(\d+)/$", a["href"])
        if m:
            max_page = max(max_page, int(m.group(1)))
    return max_page


def _parse_judo_list_soup(soup: BeautifulSoup) -> list[dict]:
    records: list[dict] = []
    for box in soup.select(".mod_area_info_box"):
        a = box.find("a", href=True)
        if not a or "/sekkotsuinsrch/10/" not in a["href"]:
            continue
        detail_url = a["href"]
        if "syuhenshisetsulist" in detail_url or detail_url.rstrip("/").endswith("/10"):
            continue
        name = a.get_text(strip=True)
        if not name:
            continue

        li = box.find_parent("li") or box.find_parent("dl") or box
        address = ""
        for tr in li.find_all("tr"):
            th = tr.find("th")
            td = tr.find("td")
            if th and td and "所在地" in th.get_text():
                address = td.get_text(strip=True)
                break

        rec = {k: "" for k in OUTPUT_COLS}
        rec["名称"] = name
        rec["所在地"] = address
        rec["_detail_url"] = detail_url
        records.append(rec)
    return records


def scrape_judo_list(total_pages: int) -> list[dict]:
    """judo-ch.jp 群馬県の全ページを収集する。"""
    all_records: list[dict] = []
    for page_num in range(1, total_pages + 1):
        url = JUDO_LIST_URL if page_num == 1 else f"{JUDO_LIST_URL}{page_num}/"
        soup = _fetch(url)
        if soup is None:
            continue
        recs = _parse_judo_list_soup(soup)
        all_records.extend(recs)
        logger.info(f"  judo-ch p{page_num}/{total_pages}: {len(recs)} 件 (累計 {len(all_records)} 件)")
        time.sleep(random.uniform(0.5, 1.0))
    return all_records


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
                if href.startswith("http") and not _is_skip_domain(href):
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
    """body-care.expert/gunma/ から Pick Up 接骨院・整骨院情報を収集する。"""
    soup = _fetch(BODYCARE_URL)
    if soup is None:
        return []
    records: list[dict] = []
    for li in soup.select("li.clearfix"):
        header = li.find("div", class_="archive_list_header")
        if not header:
            continue
        a_name = header.select_one("h2.archive_list_title a")
        if not a_name:
            continue
        name = a_name.get_text(strip=True)

        addr_el = header.select_one("p.archive_list_address")
        address = addr_el.get_text(strip=True) if addr_el else ""
        tel_el = header.select_one("p.archive_list_tel")
        phone = re.sub(r"[^0-9\-]", "", tel_el.get_text(strip=True)) if tel_el else ""

        official_url = ""
        btn = li.find("div", class_="archive_list_btn")
        if btn:
            for a in btn.find_all("a", href=True):
                href = a["href"]
                if href.startswith("http") and not _is_skip_domain(href):
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
    """chikuchikuryoho.com から群馬県分の治療院情報を収集する。"""
    soup = _fetch(CHIKU_URL)
    if soup is None:
        return []

    anchor = soup.find(id="gunma") or soup.find("a", attrs={"name": "gunma"})
    if anchor is None:
        logger.info("chikuchikuryoho.com: 群馬県セクションが見つかりません")
        return []

    post_module = anchor.find_next("div", class_=lambda c: bool(c) and "module-post" in c)
    if post_module is None:
        return []

    records: list[dict] = []
    for article in post_module.find_all("article", class_="post"):
        p = article.find("p")
        if not p:
            continue
        full_text = p.get_text(separator="\n", strip=True)

        name_el = p.find("strong")
        name = name_el.get_text(strip=True) if name_el else ""
        if not name:
            continue

        lines = [ln.strip() for ln in full_text.split("\n") if ln.strip()]
        address, phone = "", ""
        for ln in lines:
            if "群馬県" in ln and not address:
                address = ln
            m = re.search(r"電話\s*[：:]\s*([\d\-]+)", ln)
            if m:
                phone = m.group(1)

        email, official_url = "", ""
        for a in p.find_all("a", href=True):
            href = a["href"]
            if href.startswith("mailto:"):
                email = href[7:]
            elif href.startswith("http") and not _is_skip_domain(href):
                official_url = href

        rec = {k: "" for k in OUTPUT_COLS}
        rec["名称"] = name
        rec["所在地"] = address
        rec["電話番号"] = phone
        rec["メールアドレス"] = email
        rec["公式サイトURL"] = official_url
        records.append(rec)

    logger.info(f"chikuchikuryoho.com 群馬: {len(records)} 件")
    return records


# ── Phase 4: DuckDuckGo 公式サイト検索 ──────────────────

def search_official_site(name: str, address: str, phone: str) -> tuple[str, str]:
    """DuckDuckGo でクリニックの公式サイトを検索し、名称/電話番号/所在地で一致検証
    した URL と本文HTMLを返す（見つからない場合は ("", "")）。"""
    city_match = re.search(r"群馬県\s*([^\s市区町村]{2,5}[市区町村])", address)
    city = city_match.group(1) if city_match else "群馬"
    query = f"{name} {city} 公式サイト"
    try:
        with DDGS() as ddgs:
            results = list(ddgs.text(query, max_results=6))
    except Exception as e:
        logger.debug(f"DDG失敗 {name}: {e}")
        return "", ""

    for r in results:
        url = r.get("href", "")
        if not url or not url.startswith("http") or _is_skip_domain(url):
            continue
        if not _is_allowed(url):
            continue
        try:
            resp = requests.get(url, headers=HEADERS, timeout=10, allow_redirects=True)
            if resp.status_code >= 400 or _is_skip_domain(resp.url):
                continue
            resp.encoding = resp.apparent_encoding or "utf-8"
        except Exception:
            continue
        if _verify_match(resp.text, name, phone, address):
            return resp.url, resp.text
    return "", ""


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

    base_host = urlparse(url).netloc
    contact_kws = ["contact", "inquiry", "お問い合わせ", "問い合わせ", "ご相談", "メールフォーム"]
    for a in soup.find_all("a", href=True):
        href = a["href"]
        link_text = a.get_text(strip=True)
        if not any(kw in href.lower() or kw in link_text for kw in contact_kws):
            continue
        if _is_skip_domain(href):
            continue
        if href.startswith("//"):
            href = "https:" + href
        elif not href.startswith(("http", "#", "mailto:", "tel:")):
            href = urljoin(url, href)
        if href.startswith(("#", "mailto:", "tel:")):
            continue
        if urlparse(href).netloc != base_host:
            continue
        result["問い合わせフォームURL"] = href
        break

    return result


def _fetch_official_sync(url: str) -> dict:
    if _is_skip_domain(url):
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
            if url.startswith("http") and not rec.pop("_skip_refetch", False):
                extras = await asyncio.to_thread(_fetch_official_sync, url)
                if rec.get("メールアドレス") and "メールアドレス" in extras:
                    del extras["メールアドレス"]
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
    logger.info("群馬県 整体・接骨院 収集開始")
    logger.info("=" * 65)
    start_time = datetime.now()

    all_records: list[dict] = []
    seen_names: set[str] = set()
    seen_phones: set[str] = set()

    # ── Phase 1: judo-ch.jp ──────────────────────────────
    logger.info("Phase 1: judo-ch.jp から群馬県接骨院収集")
    total_pages = get_total_pages()
    judo_records = scrape_judo_list(total_pages)
    for rec in judo_records:
        nk = _normalize_name(rec["名称"])
        if nk and nk not in seen_names:
            seen_names.add(nk)
            all_records.append(rec)
    logger.info(f"Phase 1 一覧完了: {len(all_records)} 件")

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
    logger.info("Phase 3: chikuchikuryoho.com から群馬県分収集")
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
            url, html = await asyncio.to_thread(
                search_official_site, rec["名称"], rec.get("所在地", ""), rec.get("電話番号", "")
            )
            if url:
                rec["公式サイトURL"] = url
                extras = _parse_official_html(url, html)
                if rec.get("電話番号") and "電話番号" in extras:
                    del extras["電話番号"]
                rec.update(extras)
                rec["_skip_refetch"] = True
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
