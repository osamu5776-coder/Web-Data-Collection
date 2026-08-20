"""
北海道 弁護士事務所 情報収集

データソース:
  Phase 1 - 弁護士ドットコム（bengo4.com）北海道事務所一覧
            https://www.bengo4.com/office/hokkaido/ (ページネーション全件)
            から事務所名・電話番号（掲載があれば）・詳細ページURLを収集し、
            各詳細ページから所在地・事務所URL（掲載があれば）を取得する。
  Phase 2 - 札幌弁護士会（satsuben.or.jp）弁護士検索
            https://satsuben.or.jp/search/list.php?h_kana=<かな>
            （五十音各行、全44音）から所属弁護士全885名の氏名・事務所名・
            電話番号を収集し、事務所名で集約。Phase 1 に事務所名が一致しない
            （＝bengo4.com未掲載の）事務所についてのみ、代表者のプロフィール
            ページ（profile.php）から所在地・電話番号・ホームページ・
            メールアドレスを取得する。
  Phase 3 - 公式サイトURLが未判明の事務所について、DuckDuckGo で公式サイトを
            検索し、ページ本文の電話番号一致または事務所名一致で検証する。
  Phase 4 - 判明した公式サイトから メール・インスタ・問い合わせフォームURL
            （・電話番号が bengo4.com のIP電話番号のみの場合はページ内の
            市外局番付き電話番号で補完）を取得する。

出力列: 名称, メールアドレス, 公式サイトURL, 所在地, 電話番号, インスタURL, 問い合わせフォームURL
出力ファイル: 北海道_弁護士リスト.xlsx
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

LOG_FILE = "scraper_hokkaido_bengoshi.log"
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
OUTPUT_FILE = "北海道_弁護士リスト.xlsx"
UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)
HEADERS = {"User-Agent": UA}

BENGO4_BASE = "https://www.bengo4.com/office/hokkaido/"
SATSUBEN_LIST = "https://satsuben.or.jp/search/list.php"
SATSUBEN_PROFILE = "https://satsuben.or.jp/search/profile.php"

KANA_LIST = (
    list("あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめも")
    + ["や", "ゆ", "よ"]
    + list("らりるれろ")
    + ["わ"]
)

SKIP_DOMAINS = [
    # データソース自身
    "bengo4.com", "satsuben.or.jp",
    # 弁護士会・公的機関（個別事務所の公式サイトではない）
    "nichibenren.or.jp", "houterasu.or.jp", "hakoben.or.jp",
    "asahiben.or.jp", "kushiroben.or.jp",
    # ディレクトリ・検索・比較サイト
    "itp.ne.jp", "mapion.co.jp", "navitime.co.jp", "mapfan.com",
    "goo.ne.jp", "wikipedia.org", "google.com", "google.co.jp",
    "maps.google.com", "bing.com", "yahoo.co.jp", "mynavi.jp",
    "houmu-cafe.com", "asoview.com",
    # 求人・転職
    "jp.indeed.com", "indeed.com", "jp.stanby.com", "stanby.com",
    # ブログプラットフォーム
    "ameblo.jp", "seesaa.net", "ldblog.jp", "livedoor.jp", "hatenablog.com",
    "blogspot.com", "exblog.jp", "note.com",
    # SNS
    "instagram.com", "facebook.com", "twitter.com", "x.com",
    "youtube.com", "tiktok.com", "line.me", "lin.ee",
]

OUTPUT_COLS = [
    "名称", "メールアドレス", "公式サイトURL",
    "所在地", "電話番号", "インスタURL", "問い合わせフォームURL",
]

ZEN2HAN_DIGITS = str.maketrans(
    "０１２３４５６７８９ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ"
    "ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ",
    "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    "abcdefghijklmnopqrstuvwxyz",
)

# ── 共通ユーティリティ ────────────────────────────────────
_robots_cache: dict[str, RobotFileParser] = {}


def _is_allowed(url: str) -> bool:
    # RobotFileParser.read() は既定の User-Agent で robots.txt を取得するため、
    # bot対策で 403 を返すサイト（bengo4.com 等）では disallow_all 扱いになってしまう。
    # そのためブラウザUAで明示的に取得してから parse() に渡す。
    parsed = urlparse(url)
    origin = f"{parsed.scheme}://{parsed.netloc}"
    if origin not in _robots_cache:
        rp = RobotFileParser()
        rp.set_url(f"{origin}/robots.txt")
        try:
            resp = requests.get(f"{origin}/robots.txt", headers=HEADERS, timeout=10)
            if resp.status_code < 400:
                rp.parse(resp.text.splitlines())
            else:
                rp.allow_all = True
        except Exception:
            rp.allow_all = True
        _robots_cache[origin] = rp
    return _robots_cache[origin].can_fetch(UA, url)


def _fetch(url: str, timeout: int = 12, **kwargs) -> BeautifulSoup | None:
    try:
        r = requests.get(url, headers=HEADERS, timeout=timeout, **kwargs)
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


def _is_skip_domain(url: str) -> bool:
    return any(d in url for d in SKIP_DOMAINS)


def _verify_match(html_text: str, name: str, phone: str) -> bool:
    """公式サイト候補が実際にその事務所のものかを名称・電話番号で検証する。

    bengo4.com 由来の電話番号はIP電話の通話計測用番号であることが多く
    公式サイトには掲載されないため、電話番号一致だけを必須とはせず、
    事務所名（法人格・支店等を除いたコア名称）一致でも可とする。
    """
    norm_page = _normalize_name(re.sub(r"[\s　]+", "", html_text))

    phone_digits = _normalize_phone(phone)
    if phone_digits and phone_digits in re.sub(r"[^0-9]", "", html_text):
        return True

    norm_name = _normalize_name(name)
    if norm_name and norm_name in norm_page:
        return True

    core = re.sub(r"^弁護士法人", "", name)
    core = re.sub(r"(法律事務所|事務所|支店|支部|オフィス)+$", "", core).strip()
    core_norm = _normalize_name(core)
    return len(core_norm) >= 3 and core_norm in norm_page


PORTAL_TITLE_KWS = [
    "求人", "アルバイト", "転職", "口コミ", "クチコミ", "ランキング",
    "一覧｜", "検索｜", "を探す", "事務所検索", "ナビ｜", "まとめ",
]


def _is_portal_page(html: str) -> bool:
    title_m = re.search(r"<title[^>]*>(.*?)</title>", html, re.I | re.S)
    title = title_m.group(1) if title_m else ""
    return any(kw in title for kw in PORTAL_TITLE_KWS)


def _ensure_hokkaido_prefix(address: str) -> str:
    address = re.sub(r"[\s　]+", " ", address).strip()
    if not address or "北海道" in address:
        return address
    m = re.match(r"^(〒\s*\d{3}-?\d{4})\s*(.*)$", address)
    if m:
        return f"{m.group(1)} 北海道{m.group(2)}"
    return f"北海道{address}"


# ── Phase 1: 弁護士ドットコム（bengo4.com）北海道事務所一覧 ──

def scrape_bengo4_list_page(page_num: int) -> list[dict]:
    url = BENGO4_BASE if page_num == 1 else f"{BENGO4_BASE}?page={page_num}"
    soup = _fetch(url)
    if soup is None:
        return []

    records: list[dict] = []
    for h3 in soup.select("h3.list_office__name"):
        a = h3.find("a", href=True)
        if not a:
            continue
        name = a.get_text(strip=True)
        detail_url = urljoin(url, a["href"])
        if not name or not detail_url:
            continue

        tel = ""
        parent = h3.find_parent("div", class_="list_office__summary")
        if parent:
            tel_p = parent.select_one("p.list_office__tel")
            if tel_p:
                tel_text = tel_p.get_text(strip=True)
                m = re.search(r"[\d-]{9,}", tel_text)
                tel = m.group(0) if m else ""

        records.append({"名称": name, "_detail_url": detail_url, "電話番号": tel})

    return records


def scrape_bengo4_list() -> list[dict]:
    logger.info("Phase 1: 弁護士ドットコム 北海道事務所一覧を収集")
    first_soup = _fetch(BENGO4_BASE)
    if first_soup is None:
        logger.error("  一覧ページ取得失敗")
        return []

    page_nums = [int(m) for m in re.findall(r"page=(\d+)", str(first_soup))]
    total_pages = max(page_nums) if page_nums else 1
    logger.info(f"  総ページ数: {total_pages}")

    all_records: list[dict] = []
    for h3 in first_soup.select("h3.list_office__name"):
        a = h3.find("a", href=True)
        if not a:
            continue
        name = a.get_text(strip=True)
        detail_url = urljoin(BENGO4_BASE, a["href"])
        tel = ""
        parent = h3.find_parent("div", class_="list_office__summary")
        if parent:
            tel_p = parent.select_one("p.list_office__tel")
            if tel_p:
                m = re.search(r"[\d-]{9,}", tel_p.get_text(strip=True))
                tel = m.group(0) if m else ""
        if name and detail_url:
            all_records.append({"名称": name, "_detail_url": detail_url, "電話番号": tel})
    logger.info(f"  page 1/{total_pages}: {len(all_records)} 件")

    for page_num in range(2, total_pages + 1):
        recs = scrape_bengo4_list_page(page_num)
        all_records.extend(recs)
        logger.info(f"  page {page_num}/{total_pages}: {len(recs)} 件 (累計 {len(all_records)} 件)")
        time.sleep(random.uniform(1.0, 2.0))

    logger.info(f"Phase 1 一覧取得完了: {len(all_records)} 件")
    return all_records


def _parse_bengo4_detail(html: str) -> dict:
    soup = BeautifulSoup(html, "lxml")
    result = {"所在地": "", "公式サイトURL": ""}

    for dt in soup.find_all("dt"):
        label = dt.get_text(strip=True)
        if label == "所在地":
            dd = dt.find_next_sibling("dd")
            if dd:
                addr = dd.get_text(separator=" ", strip=True)
                addr = re.sub(r"^〒\s*", "〒", addr)
                result["所在地"] = _ensure_hokkaido_prefix(addr)
        elif "事務所URL" in label:
            dd = dt.find_next_sibling("dd")
            a = dd.find("a", href=True) if dd else None
            if a:
                href = a["href"].strip()
                if href.startswith("http") and not _is_skip_domain(href):
                    result["公式サイトURL"] = href

    return result


def _fetch_bengo4_detail_sync(url: str) -> dict:
    if not _is_allowed(url):
        return {}
    try:
        r = requests.get(url, headers=HEADERS, timeout=12)
        if r.status_code >= 400:
            return {}
        r.encoding = r.apparent_encoding or "utf-8"
        return _parse_bengo4_detail(r.text)
    except Exception:
        return {}


async def enrich_bengo4_details(records: list[dict], sem_count: int = 4) -> None:
    semaphore = asyncio.Semaphore(sem_count)
    total = len(records)
    done = {"n": 0}

    async def _one(rec: dict) -> None:
        async with semaphore:
            extras = await asyncio.to_thread(_fetch_bengo4_detail_sync, rec["_detail_url"])
            rec.update(extras)
            await asyncio.sleep(random.uniform(0.8, 1.6))
            done["n"] += 1
            if done["n"] % 20 == 0 or done["n"] == total:
                logger.info(f"  詳細ページ取得: {done['n']}/{total}")

    await asyncio.gather(*[_one(rec) for rec in records])


# ── Phase 2: 札幌弁護士会（satsuben.or.jp） ──────────────

def scrape_satsuben_kana(kana: str) -> list[dict]:
    soup = _fetch(SATSUBEN_LIST, params={"h_kana": kana})
    if soup is None:
        return []

    records: list[dict] = []
    for li in soup.select("li.clearfix"):
        a = li.find("a", href=True) if li.find("a") is None else li.find("a")
        # gotoProfile(ID) は href ではなく javascript: 疑似リンクのため member_id を正規表現で抽出
        raw_a = li.find("a")
        if raw_a is None:
            continue
        href_js = raw_a.get("href", "")
        m_id = re.search(r"gotoProfile\((\d+)\)", href_js)
        if not m_id:
            continue
        member_id = m_id.group(1)

        name_span = li.select_one("span.name")
        office_span = li.select_one("span.office")
        tel_span = li.select_one("span.tel")
        name = name_span.get_text(strip=True) if name_span else ""
        office = office_span.get_text(strip=True) if office_span else ""
        tel = tel_span.get_text(strip=True) if tel_span else ""
        if not office:
            continue

        records.append({
            "弁護士名": name, "事務所名": office, "電話番号": tel, "member_id": member_id,
        })

    return records


def scrape_satsuben_roster() -> list[dict]:
    logger.info("Phase 2: 札幌弁護士会 会員名簿を収集（五十音44音）")
    all_records: list[dict] = []
    for i, kana in enumerate(KANA_LIST, 1):
        recs = scrape_satsuben_kana(kana)
        all_records.extend(recs)
        if i % 10 == 0 or i == len(KANA_LIST):
            logger.info(f"  {i}/{len(KANA_LIST)} 音 完了 (累計 {len(all_records)} 名)")
        time.sleep(random.uniform(0.8, 1.5))
    logger.info(f"Phase 2 名簿取得完了: {len(all_records)} 名")
    return all_records


def _parse_satsuben_profile(html: str) -> dict:
    soup = BeautifulSoup(html, "lxml")
    result: dict = {}
    for tr in soup.select("div.dataTable table tr"):
        th = tr.find("th")
        td = tr.find("td")
        if th is None or td is None:
            continue
        label = th.get_text(strip=True).translate(ZEN2HAN_DIGITS)
        label = label.replace("Ｔ", "T").replace("Ｅ", "E").replace("Ｌ", "L")
        text = td.get_text(separator=" ", strip=True)
        if not text or text == "":
            continue
        if label == "住所":
            result["所在地"] = _ensure_hokkaido_prefix(text.translate(ZEN2HAN_DIGITS))
        elif label == "TEL":
            result["電話番号"] = text
        elif label == "ホームページ":
            m = re.search(r"https?://\S+", text)
            if m and not _is_skip_domain(m.group(0)):
                result["公式サイトURL"] = m.group(0)
        elif label == "メールアドレス":
            m = re.search(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}", text)
            if m:
                result["メールアドレス"] = m.group(0)
    return result


def _fetch_satsuben_profile_sync(member_id: str, kana: str) -> dict:
    url = SATSUBEN_PROFILE
    params = {"h_kana": kana, "h_back_url": "/search/list.php", "h_member_id": member_id}
    full_url = f"{url}?" + "&".join(f"{k}={v}" for k, v in params.items())
    if not _is_allowed(full_url):
        return {}
    try:
        r = requests.get(url, params=params, headers=HEADERS, timeout=12)
        if r.status_code >= 400:
            return {}
        r.encoding = r.apparent_encoding or "utf-8"
        return _parse_satsuben_profile(r.text)
    except Exception:
        return {}


async def enrich_satsuben_profiles(records: list[dict], sem_count: int = 4) -> None:
    semaphore = asyncio.Semaphore(sem_count)
    total = len(records)
    done = {"n": 0}

    async def _one(rec: dict) -> None:
        async with semaphore:
            kana = rec.get("_kana", "あ")
            extras = await asyncio.to_thread(
                _fetch_satsuben_profile_sync, rec["member_id"], kana
            )
            rec.update(extras)
            await asyncio.sleep(random.uniform(0.8, 1.6))
            done["n"] += 1
            if done["n"] % 20 == 0 or done["n"] == total:
                logger.info(f"  会員詳細（未掲載事務所分）取得: {done['n']}/{total}")

    await asyncio.gather(*[_one(rec) for rec in records])


# ── Phase 3: DuckDuckGo 公式サイト検索+検証 ──────────────

def search_official_site(name: str, address: str, phone: str) -> tuple[str, str]:
    city_match = re.search(r"北海道\s*([^\s]{2,8}?[市区町村])", address)
    city = city_match.group(1) if city_match else "北海道"
    query = f"{name} {city} 弁護士 公式サイト"
    results = []
    for backend in ("yandex", "brave", "google"):
        try:
            with DDGS(timeout=8) as ddgs:
                results = list(ddgs.text(query, max_results=6, backend=backend))
            if results:
                break
        except Exception as e:
            logger.debug(f"検索失敗 {name} ({backend}): {e}")
            continue

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
        if _is_portal_page(resp.text):
            continue
        if _verify_match(resp.text, name, phone):
            return resp.url, resp.text
    return "", ""


# ── Phase 4: 公式サイト追加情報 ──────────────────────────

def _decode_cfemail(encoded: str) -> str:
    try:
        key = int(encoded[:2], 16)
        return bytes(
            int(encoded[i:i + 2], 16) ^ key for i in range(2, len(encoded), 2)
        ).decode("utf-8")
    except Exception:
        return ""


def _parse_official_html(url: str, html: str) -> dict:
    result: dict = {}
    soup = BeautifulSoup(html, "lxml")
    text = soup.get_text(separator="\n")
    base_host = urlparse(url).netloc

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

    contact_kws = ["contact", "inquiry", "お問い合わせ", "問い合わせ", "ご相談", "メールフォーム", "予約"]
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

    # bengo4.com のIP電話番号（050始まり）しか電話番号を持たない場合、
    # 公式サイト本文に掲載されている市外局番付き番号で補完する
    m = re.search(r"0\d{1,4}-\d{1,4}-\d{4}", text)
    if m:
        result.setdefault("_site_phone", m.group(0))

    return result


def _fetch_official_sync(url: str) -> dict:
    if _is_skip_domain(url) or not _is_allowed(url):
        return {}
    try:
        r = requests.get(url, headers=HEADERS, timeout=10, allow_redirects=True)
        if r.status_code >= 400:
            return {}
        r.encoding = r.apparent_encoding or "utf-8"
        return _parse_official_html(r.url, r.text)
    except Exception:
        return {}


async def enrich_known_urls(records: list[dict], sem_count: int = 6) -> None:
    targets = [r for r in records if r.get("公式サイトURL")]
    if not targets:
        return
    semaphore = asyncio.Semaphore(sem_count)
    total = len(targets)
    done = {"n": 0}

    async def _one(rec: dict) -> None:
        async with semaphore:
            extras = await asyncio.to_thread(_fetch_official_sync, rec["公式サイトURL"])
            site_phone = extras.pop("_site_phone", "")
            if site_phone and rec.get("電話番号", "").startswith("050"):
                rec["電話番号"] = site_phone
            rec.update(extras)
            done["n"] += 1
            if done["n"] % 20 == 0 or done["n"] == total:
                logger.info(f"  公式サイト情報取得: {done['n']}/{total}")

    await asyncio.gather(*[_one(rec) for rec in targets])


async def enrich_search_records(records: list[dict], sem_count: int = 3) -> None:
    targets = [r for r in records if not r.get("公式サイトURL")]
    semaphore = asyncio.Semaphore(sem_count)
    total = len(targets)
    done = {"n": 0, "found": 0}

    async def _one(rec: dict) -> None:
        async with semaphore:
            try:
                url, html = await asyncio.wait_for(
                    asyncio.to_thread(
                        search_official_site, rec["名称"], rec.get("所在地", ""), rec.get("電話番号", "")
                    ),
                    timeout=45,
                )
            except asyncio.TimeoutError:
                logger.debug(f"検索タイムアウト: {rec['名称']}")
                url, html = "", ""
            if url:
                rec["公式サイトURL"] = url
                extras = _parse_official_html(url, html)
                site_phone = extras.pop("_site_phone", "")
                if site_phone and rec.get("電話番号", "").startswith("050"):
                    rec["電話番号"] = site_phone
                rec.update(extras)
                done["found"] += 1
            done["n"] += 1
            n = done["n"]
            if n % 20 == 0 or n == total:
                logger.info(f"  検索・検証: {n}/{total} (公式サイト発見 {done['found']} 件)")
            await asyncio.sleep(random.uniform(1.0, 2.0))

    await asyncio.gather(*[_one(rec) for rec in targets])


# ── メイン ────────────────────────────────────────────────

async def main() -> None:
    logger.info("=" * 65)
    logger.info("北海道 弁護士事務所 収集開始")
    logger.info("=" * 65)
    start_time = datetime.now()

    bengo4_raw = scrape_bengo4_list()
    logger.info("Phase 1: 事務所詳細ページ（所在地・事務所URL）を取得")
    await enrich_bengo4_details(bengo4_raw, sem_count=4)

    deduped: list[dict] = []
    seen_names: set[str] = set()
    seen_phones: set[str] = set()

    for raw in bengo4_raw:
        nk = _normalize_name(raw["名称"])
        pk = _normalize_phone(raw.get("電話番号", ""))
        if nk and nk in seen_names:
            continue
        seen_names.add(nk)
        if pk:
            seen_phones.add(pk)
        rec = {k: "" for k in OUTPUT_COLS}
        rec["名称"] = raw["名称"]
        rec["所在地"] = raw.get("所在地", "")
        rec["電話番号"] = raw.get("電話番号", "")
        rec["公式サイトURL"] = raw.get("公式サイトURL", "")
        deduped.append(rec)

    logger.info(f"Phase 1 完了: {len(deduped)} 事務所")

    satsuben_raw = scrape_satsuben_roster()

    office_groups: dict[str, dict] = {}
    for r in satsuben_raw:
        nk = _normalize_name(r["事務所名"])
        if not nk or nk in office_groups:
            continue
        r["_kana"] = r["事務所名"][0] if r["事務所名"] else "あ"
        office_groups[nk] = r

    novel = [r for nk, r in office_groups.items() if nk not in seen_names]
    logger.info(
        f"Phase 2: 会員名簿から一意事務所 {len(office_groups)} 件中、"
        f"Phase 1 未掲載 {len(novel)} 件のプロフィールを取得"
    )
    await enrich_satsuben_profiles(novel, sem_count=4)

    added = 0
    for r in novel:
        nk = _normalize_name(r["事務所名"])
        pk = _normalize_phone(r.get("電話番号", ""))
        if nk in seen_names or (pk and pk in seen_phones):
            continue
        seen_names.add(nk)
        if pk:
            seen_phones.add(pk)
        rec = {k: "" for k in OUTPUT_COLS}
        rec["名称"] = r["事務所名"]
        rec["所在地"] = r.get("所在地", "")
        rec["電話番号"] = r.get("電話番号", "")
        rec["公式サイトURL"] = r.get("公式サイトURL", "")
        rec["メールアドレス"] = r.get("メールアドレス", "")
        deduped.append(rec)
        added += 1

    logger.info(f"Phase 2 完了: {added} 事務所を追加（合計 {len(deduped)} 事務所）")
    logger.info(f"  うち公式サイトURL判明: {sum(1 for r in deduped if r.get('公式サイトURL'))} 件")

    logger.info("Phase 3+4: 既知の公式サイトから追加情報取得")
    await enrich_known_urls(deduped, sem_count=6)

    logger.info("Phase 3+4: 未判明の事務所を DuckDuckGo で検索・検証・情報取得")
    await enrich_search_records(deduped, sem_count=3)
    logger.info(f"公式サイト取得 合計: {sum(1 for r in deduped if r.get('公式サイトURL'))} 件")

    df = pd.DataFrame(deduped, columns=OUTPUT_COLS)
    df.drop_duplicates(subset=["名称"], keep="first", inplace=True)
    for col in ["問い合わせフォームURL", "メールアドレス"]:
        mask_empty = df[col] == ""
        df = pd.concat([
            df[mask_empty],
            df[~mask_empty].drop_duplicates(subset=[col], keep="first"),
        ]).sort_index()
    df = df.reset_index(drop=True)
    df.to_excel(OUTPUT_FILE, index=False)

    elapsed = int((datetime.now() - start_time).total_seconds())
    logger.info("=" * 65)
    logger.info(f"完了: {OUTPUT_FILE} に {len(df)} 件を出力")
    logger.info(f"所要時間: {elapsed // 60}分{elapsed % 60}秒")
    logger.info("=" * 65)


if __name__ == "__main__":
    asyncio.run(main())
