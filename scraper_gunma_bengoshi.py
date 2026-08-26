"""
群馬県 法律事務所 情報収集

データソース:
  Phase 1 - 群馬弁護士会 弁護士情報検索システム（raijin3, 群馬弁護士会公式）
            http://www12.wind.ne.jp/raijin3/cgi-bin/bengosi-db/database.cgi
            性別（男・女）で絞り込み条件を満たしつつ実質全件（328名）を取得。
            一覧に氏名・事務所名・住所・電話番号が掲載されているため、
            事務所名で集約したうえで代表者の詳細ページ（cmd=j）から
            ホームページ（掲載があれば）を取得する。
  Phase 2 - ひまわりサーチ（弁護士情報提供サービス、bengoshikai.jp、kai_code=9）
            https://www.bengoshikai.jp/search/list.php
            群馬弁護士会所属で本サービスに登録した弁護士38名（2026年時点）の
            氏名・事務所名・住所・電話番号を取得。Phase 1 に事務所名が一致
            しない（＝raijin3未掲載の）事務所のみ追加する。
  Phase 3 - 公式サイトURLが未判明の事務所について、DuckDuckGo で公式サイトを
            検索し、ページ本文の電話番号一致または事務所名一致で検証する。
  Phase 4 - 判明した公式サイトから メール・インスタ・問い合わせフォームURL
            を取得する。

出力列: 名称, メールアドレス, 公式サイトURL, 所在地, 電話番号, インスタURL, 問い合わせフォームURL
出力ファイル: 群馬県_法律事務所リスト.xlsx
"""

import asyncio
import io
import logging
import random
import re
import sys
import time
from datetime import datetime
from urllib.parse import quote, urljoin, urlparse
from urllib.robotparser import RobotFileParser

import pandas as pd
import requests
from bs4 import BeautifulSoup
from ddgs import DDGS

# ── ロギング ──────────────────────────────────────────────
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

LOG_FILE = "scraper_gunma_bengoshi.log"
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
OUTPUT_FILE = "群馬県_法律事務所リスト.xlsx"
UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)
HEADERS = {"User-Agent": UA}

RAIJIN_CGI = "http://www12.wind.ne.jp/raijin3/cgi-bin/bengosi-db/database.cgi"
RAIJIN_REFERER = "http://www12.wind.ne.jp/raijin3/cgi-bin/bengosi-db/search.html"
RAIJIN_HYOJISU = 20

HIMAWARI_BASE = "https://www.bengoshikai.jp/search/"
HIMAWARI_KAI_CODE = "9"

SKIP_DOMAINS = [
    # データソース自身・弁護士会公式サイト（個別事務所ではない）
    "bengoshikai.jp", "wind.ne.jp", "gunben.or.jp", "nichibenren.or.jp",
    "houterasu.or.jp", "bengo4.com", "bengoshi-search.site", "bennavi.jp",
    "chuokeizai.co.jp",
    # 離婚・相続・債務整理などの弁護士紹介ディレクトリ（個別事務所ではない）
    "ricon-pro.com", "souzoku-pro.info", "saimuseiri110.net", "bengoshi.quest",
    "bengoshikensaku.com", "bkangunma.net", "ai-chosa-maebashi.com",
    "sashiireya.com", "agoora.co.jp", "souzokubengo-line.com",
    "rikon.asahi.com", "souzoku.asahi.com", "saimuseiri.asahi.com",
    "xn--zqs94l3txt9rgzaw2z12g.jp", "xn--u9jy76gb0o9pg1qbg18h15c.jp",
    # ディレクトリ・検索・比較サイト
    "itp.ne.jp", "mapion.co.jp", "navitime.co.jp", "mapfan.com",
    "goo.ne.jp", "wikipedia.org", "google.com", "google.co.jp",
    "maps.google.com", "maps.google.co.jp", "bing.com", "yahoo.co.jp", "mynavi.jp",
    "houmu-cafe.com",
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

ZEN2HAN = str.maketrans(
    "０１２３４５６７８９ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ"
    "ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ",
    "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    "abcdefghijklmnopqrstuvwxyz",
)

# ── 共通ユーティリティ ────────────────────────────────────
_robots_cache: dict[str, RobotFileParser] = {}


def _is_allowed(url: str) -> bool:
    # RobotFileParser.read() は既定の User-Agent で robots.txt を取得するため、
    # bot対策で 403 を返すサイトでは disallow_all 扱いになってしまう。
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
    s = s.translate(ZEN2HAN)
    return s.lower()


def _normalize_phone(s: str) -> str:
    return re.sub(r"[^0-9]", "", str(s))


SKIP_URL_PATTERNS = [
    r"\.pdf(\?|$)",       # 個別事務所の公式サイトではなくPDF資料であることが多い
    r"bengoshi[-_]?search", # 弁護士検索ディレクトリサイト
    r"bengoshi[-_]?navi",   # 弁護士ナビ系ディレクトリサイト
]


def _is_skip_domain(url: str) -> bool:
    if any(d in url for d in SKIP_DOMAINS):
        return True
    return any(re.search(p, url, re.I) for p in SKIP_URL_PATTERNS)


def _verify_match(html_text: str, name: str, phone: str) -> bool:
    """公式サイト候補が実際にその事務所のものかを名称・電話番号で検証する。"""
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


def _ensure_gunma_prefix(address: str) -> str:
    address = re.sub(r"[\s　]+", " ", address).strip()
    address = re.sub(r"^〒?\s*\d{3}-?\d{4}\s*", "", address).strip()
    if not address or "群馬県" in address:
        return address
    return f"群馬県{address}"


# ── Phase 1: 群馬弁護士会 弁護士情報検索システム（raijin3） ──

def _raijin_query(page: int) -> str:
    pairs = [
        ("cmd", "s"), ("S_20_Key_Sex", "男,女"), ("HyojiSu", str(RAIJIN_HYOJISU)),
        ("Tfile", "Data"), ("TrColor", "#ffffff,#e8ffe8"),
        ("Type_20", "Equal-or"), ("Option_20", "checkbox"), ("page", str(page)),
    ]
    return "&".join(f"{quote(k)}={quote(v, encoding='cp932')}" for k, v in pairs)


def _raijin_post_first_page() -> tuple[str, int]:
    pairs = [
        ("cmd", "s"), ("Tfile", ""), ("HTML", ""), ("DataHtml", ""),
        ("TrColor", "#ffffff,#e8ffe8"),
        ("S_20_Key_Sex", "男"), ("S_20_Key_Sex", "女"),
        ("Option_21", "checkbox"), ("Type_21", "Normal-or"),
        ("Option_20", "checkbox"), ("Type_20", "Equal-or"),
        ("Option_22", "checkbox"), ("Type_22", "Normal-or"),
    ]
    body = "&".join(f"{quote(k)}={quote(v, encoding='cp932')}" for k, v in pairs)
    headers = {**HEADERS, "Referer": RAIJIN_REFERER, "Content-Type": "application/x-www-form-urlencoded"}
    r = requests.post(RAIJIN_CGI, data=body.encode("ascii"), headers=headers, timeout=20)
    r.encoding = "shift_jis"
    count_m = re.search(r'<B>(\d+)</B></FONT>\s*件が該当しました', r.text)
    total = int(count_m.group(1)) if count_m else 0
    return r.text, total


def _parse_raijin_list(html: str) -> list[dict]:
    soup = BeautifulSoup(html, "lxml")
    records: list[dict] = []
    for a in soup.select('a[href*="cmd=j"]'):
        m = re.search(r"DataNum=(\d+)", a["href"])
        if not m:
            continue
        data_num = m.group(1)
        row = a.find_parent("tr")
        if row is None:
            continue
        tds = row.find_all("td")
        if len(tds) < 4:
            continue
        name = a.get_text(strip=True)
        office = tds[1].get_text(strip=True)
        address = tds[2].get_text(separator=" ", strip=True)
        phone = tds[3].get_text(strip=True)
        if not office:
            continue
        records.append({
            "弁護士名": name, "事務所名": office,
            "所在地": _ensure_gunma_prefix(address), "電話番号": phone,
            "data_num": data_num,
        })
    return records


def scrape_raijin_roster() -> list[dict]:
    logger.info("Phase 1: 群馬弁護士会 弁護士情報検索システム（raijin3）を収集")
    first_html, total = _raijin_post_first_page()
    if total == 0:
        logger.error("  一覧取得失敗（0件）")
        return []
    total_pages = (total + RAIJIN_HYOJISU - 1) // RAIJIN_HYOJISU
    logger.info(f"  対象: {total} 名 ({total_pages} ページ)")

    all_records = _parse_raijin_list(first_html)
    logger.info(f"  page 1/{total_pages}: {len(all_records)} 名")

    headers = {**HEADERS, "Referer": RAIJIN_REFERER}
    for page in range(2, total_pages + 1):
        url = f"{RAIJIN_CGI}?{_raijin_query(page)}"
        try:
            r = requests.get(url, headers=headers, timeout=15)
            r.encoding = "shift_jis"
            recs = _parse_raijin_list(r.text)
        except Exception as e:
            logger.debug(f"page {page} 取得失敗: {e}")
            recs = []
        all_records.extend(recs)
        logger.info(f"  page {page}/{total_pages}: {len(recs)} 名 (累計 {len(all_records)} 名)")
        time.sleep(random.uniform(0.8, 1.5))

    logger.info(f"Phase 1 名簿取得完了: {len(all_records)} 名")
    return all_records


def _parse_raijin_detail(html: str) -> dict:
    result: dict = {}
    m = re.search(r"ホームページ</td>\s*<td[^>]*>(.*?)</td>", html, re.S)
    if m:
        a_m = re.search(r'href="(https?://[^"]+)"', m.group(1))
        if a_m and not _is_skip_domain(a_m.group(1)):
            result["公式サイトURL"] = a_m.group(1)
    return result


def _fetch_raijin_detail_sync(data_num: str) -> dict:
    url = f"{RAIJIN_CGI}?cmd=j&DataNum={data_num}"
    if not _is_allowed(url):
        return {}
    headers = {**HEADERS, "Referer": RAIJIN_REFERER}
    try:
        r = requests.get(RAIJIN_CGI, params={"cmd": "j", "DataNum": data_num}, headers=headers, timeout=12)
        if r.status_code >= 400:
            return {}
        r.encoding = "shift_jis"
        return _parse_raijin_detail(r.text)
    except Exception:
        return {}


async def enrich_raijin_details(records: list[dict], sem_count: int = 4) -> None:
    semaphore = asyncio.Semaphore(sem_count)
    total = len(records)
    done = {"n": 0}

    async def _one(rec: dict) -> None:
        async with semaphore:
            extras = await asyncio.to_thread(_fetch_raijin_detail_sync, rec["data_num"])
            rec.update(extras)
            await asyncio.sleep(random.uniform(0.6, 1.2))
            done["n"] += 1
            if done["n"] % 20 == 0 or done["n"] == total:
                logger.info(f"  詳細ページ（ホームページ）取得: {done['n']}/{total}")

    await asyncio.gather(*[_one(rec) for rec in records])


# ── Phase 2: ひまわりサーチ（bengoshikai.jp, kai_code=9） ──

def scrape_himawari_roster() -> list[dict]:
    logger.info("Phase 2: ひまわりサーチ（群馬弁護士会）を収集")
    s = requests.Session()
    try:
        s.post(
            f"{HIMAWARI_BASE}form.php",
            data={"action_search_form": "true", "kai_code": HIMAWARI_KAI_CODE},
            headers={**HEADERS, "Referer": f"{HIMAWARI_BASE}?kai_code={HIMAWARI_KAI_CODE}"},
            timeout=15,
        )
        result_data = {
            "name": "", "kana": "", "member_section": "", "office_name": "",
            "hometown_pref_id": "", "office_addr": "", "hometown": "",
            "foreign_language": "", "genshikakukoku": "", "cond_if": "AND",
            "free_area": "", "kai_code": HIMAWARI_KAI_CODE, "action_search_result": "true",
        }
        s.post(
            f"{HIMAWARI_BASE}result.php", data=result_data,
            headers={**HEADERS, "Referer": f"{HIMAWARI_BASE}form.php?kai_code={HIMAWARI_KAI_CODE}"},
            timeout=15,
        )
        list_data = {
            "kai_code": HIMAWARI_KAI_CODE, "mode": "", "action_search_list": "true",
            "submit[list]": "氏名の一覧を表示する",
        }
        r = s.post(
            f"{HIMAWARI_BASE}list.php", data=list_data,
            headers={**HEADERS, "Referer": f"{HIMAWARI_BASE}result.php"},
            timeout=15,
        )
    except Exception as e:
        logger.error(f"  ひまわりサーチ取得失敗: {e}")
        return []

    r.encoding = r.apparent_encoding or "utf-8"
    soup = BeautifulSoup(r.text, "lxml")
    records: list[dict] = []
    for li in soup.select("li"):
        a = li.find("a", href=True)
        if a is None or "detail.php" not in a.get("href", ""):
            continue
        full_text = li.get_text(separator="　", strip=True)
        parts = [p for p in full_text.split("　") if p]
        # 例: 「小林有斗（こばやしゆうと）　あがつま法律事務所　群馬県...　TEL:0279-26-2100」
        if len(parts) < 3:
            continue
        office = parts[1]
        phone_m = re.search(r"TEL[:：]\s*([\d-]+)", full_text)
        phone = phone_m.group(1) if phone_m else ""
        addr_text = full_text
        addr_m = re.search(r"群馬県.*?(?=(?:　*TEL[:：]))", full_text)
        address = addr_m.group(0).strip() if addr_m else ""
        if not office:
            continue
        records.append({
            "事務所名": office, "所在地": _ensure_gunma_prefix(address), "電話番号": phone,
        })

    logger.info(f"Phase 2 名簿取得完了: {len(records)} 名")
    return records


# ── Phase 3: DuckDuckGo 公式サイト検索+検証 ──────────────

def search_official_site(name: str, address: str, phone: str) -> tuple[str, str]:
    city_match = re.search(r"群馬県\s*([^\s]{2,8}?[市区町村郡])", address)
    city = city_match.group(1) if city_match else "群馬県"
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
            if site_phone and not rec.get("電話番号"):
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
                if site_phone and not rec.get("電話番号"):
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
    logger.info("群馬県 法律事務所 収集開始")
    logger.info("=" * 65)
    start_time = datetime.now()

    raijin_raw = scrape_raijin_roster()

    office_groups: dict[str, dict] = {}
    for r in raijin_raw:
        nk = _normalize_name(r["事務所名"])
        if not nk or nk in office_groups:
            continue
        office_groups[nk] = r
    logger.info(f"Phase 1: 一意事務所 {len(office_groups)} 件")

    reps = list(office_groups.values())
    logger.info("Phase 1: 事務所詳細ページ（ホームページ）を取得")
    await enrich_raijin_details(reps, sem_count=4)

    deduped: list[dict] = []
    seen_names: set[str] = set()
    seen_phones: set[str] = set()

    for r in reps:
        nk = _normalize_name(r["事務所名"])
        pk = _normalize_phone(r.get("電話番号", ""))
        if nk and nk in seen_names:
            continue
        seen_names.add(nk)
        if pk:
            seen_phones.add(pk)
        rec = {k: "" for k in OUTPUT_COLS}
        rec["名称"] = r["事務所名"]
        rec["所在地"] = r.get("所在地", "")
        rec["電話番号"] = r.get("電話番号", "")
        rec["公式サイトURL"] = r.get("公式サイトURL", "")
        deduped.append(rec)

    logger.info(f"Phase 1 完了: {len(deduped)} 事務所")

    himawari_raw = scrape_himawari_roster()
    himawari_groups: dict[str, dict] = {}
    for r in himawari_raw:
        nk = _normalize_name(r["事務所名"])
        if not nk or nk in himawari_groups:
            continue
        himawari_groups[nk] = r

    added = 0
    for nk, r in himawari_groups.items():
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
        deduped.append(rec)
        added += 1

    logger.info(f"Phase 2 完了: {added} 事務所を追加（合計 {len(deduped)} 事務所）")
    logger.info(f"  うち公式サイトURL判明: {sum(1 for r in deduped if r.get('公式サイトURL'))} 件")

    logger.info("Phase 3+4: 既知の公式サイトから追加情報取得")
    await enrich_known_urls(deduped, sem_count=6)

    logger.info("Phase 3+4: 未判明の事務所を DuckDuckGo で検索・検証・情報取得")
    await enrich_search_records(deduped, sem_count=3)
    logger.info(f"公式サイト取得 合計: {sum(1 for r in deduped if r.get('公式サイトURL'))} 件")

    # 同一ドメインが異なる複数事務所に一致した場合、個別事務所の公式サイトではなく
    # 弁護士紹介ディレクトリ等のポータルサイトを誤って採用した可能性が高いため、
    # 該当ドメインから取得した情報をまとめて除去する（未知のディレクトリサイトに
    # 対する事後的な安全網。SKIP_DOMAINS の個別列挙だけでは新規サイトを防げないため）。
    domain_counts: dict[str, int] = {}
    for r in deduped:
        if r.get("公式サイトURL"):
            domain_counts[urlparse(r["公式サイトURL"]).netloc] = (
                domain_counts.get(urlparse(r["公式サイトURL"]).netloc, 0) + 1
            )
    portal_domains = {d for d, c in domain_counts.items() if c > 1}
    if portal_domains:
        logger.info(f"  ポータルサイトと判定しURLを除去したドメイン: {sorted(portal_domains)}")
        for r in deduped:
            if r.get("公式サイトURL") and urlparse(r["公式サイトURL"]).netloc in portal_domains:
                for col in ["公式サイトURL", "メールアドレス", "インスタURL", "問い合わせフォームURL"]:
                    r[col] = ""

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
