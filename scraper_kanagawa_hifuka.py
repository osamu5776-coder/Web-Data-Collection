"""
神奈川県 皮膚科 情報収集

データソース:
  Phase 1 - 神奈川県皮膚科医会（kanahifu.org）病院・医院検索
            全域チェックボックス（kn=on）で1回のPOSTから会員医療機関
            （医療機関名・所在地・電話番号・公式サイトURL）を取得。
            医会が直接管理するデータのため、掲載されている公式サイトURLは
            そのまま信頼できる（検索での再検証は行わない）。
            同一医療機関が複数の会員医師で重複掲載されるため、
            医療機関名＋所在地で重複排除する。
  Phase 2 - 既知の公式サイトURLについて メール・インスタ・問い合わせフォームURL
            を取得する。
  Phase 3 - Phase 1 で公式サイトURLが得られなかった医療機関について、
            DuckDuckGo で公式サイトを検索し、ページ本文の電話番号一致を
            必須として名称/市区町村の裏付けも要求する検証を行う。
  Phase 4 - Phase 3 で発見できた公式サイトから メール・インスタ・
            問い合わせフォームURL を取得する。

出力列: 名称, メールアドレス, 公式サイトURL, 所在地, 電話番号, インスタURL, 問い合わせフォームURL
出力ファイル: 神奈川県_皮膚科リスト.xlsx
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

LOG_FILE = "scraper_kanagawa_hifuka.log"
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
OUTPUT_FILE = "神奈川県_皮膚科リスト.xlsx"
KANAHIFU_SEARCH_URL = "https://www.kanahifu.org/search/search.php"
PREF_NAME = "神奈川県"
UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)
HEADERS = {"User-Agent": UA}

OUTPUT_COLS = [
    "名称", "メールアドレス", "公式サイトURL",
    "所在地", "電話番号", "インスタURL", "問い合わせフォームURL",
]

SKIP_DOMAINS = [
    # データソース自身
    "kanahifu.org",
    # 口コミ・紹介・検索ディレクトリ
    "doctorsfile.jp", "hospitalsfile.doctorsfile.jp",
    "caloo.jp", "ekiten.jp", "qlife.jp", "medley.life", "byoinnavi.jp",
    "kanagawa-doctors.com", "medicaldoc.jp", "epark.jp",
    "homemate-research.com", "homemate-research-clinic.com",
    "clinicfor.life", "hospita.jp", "10man-doc.jp", "10man-doc.co.jp",
    "e-doctor.co.jp", "kakaru.mynavi.jp", "my-best.com", "medicalnote.jp",
    # 官公庁・公的データベース
    "iryou.teikyouseido.mhlw.go.jp", "mhlw.go.jp", "kaigokensaku.mhlw.go.jp",
    "wam.go.jp", "dermatol.or.jp",
    # 医師会
    ".med.or.jp",
    # 求人・転職・アルバイト情報サイト
    "jp.indeed.com", "indeed.com", "jp.stanby.com", "stanby.com",
    "job-medley.com", "mynavi.jp",
    # ブログプラットフォーム
    "ameblo.jp", "seesaa.net", "ldblog.jp", "livedoor.jp", "hatenablog.com",
    "blogspot.com", "exblog.jp", "note.com",
    # 情報・地図・検索サイト
    "itp.ne.jp", "mapion.co.jp", "navitime.co.jp", "mapfan.com",
    "goo.ne.jp", "wikipedia.org", "google.com", "google.co.jp",
    "maps.google.com", "bing.com", "yahoo.co.jp",
    # SNS
    "instagram.com", "facebook.com", "twitter.com", "x.com",
    "youtube.com", "tiktok.com", "line.me", "lin.ee",
]

SKIP_URL_PATTERNS = [
    r"\.med\.or\.jp",
]

PORTAL_TITLE_KWS = [
    "求人", "アルバイト", "転職", "口コミ", "クチコミ", "ランキング",
    "医療連携", "医師会", "地域医療連携", "一覧｜", "検索｜", "を探す",
    "施設検索", "病院検索", "クリニック検索", "ナビ｜", "まとめ",
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
    try:
        return _robots_cache[origin].can_fetch(UA, url)
    except Exception:
        return True


def _is_skip_domain(url: str) -> bool:
    if any(d in url for d in SKIP_DOMAINS):
        return True
    return any(re.search(p, url, re.I) for p in SKIP_URL_PATTERNS)


def _normalize_name(s: str) -> str:
    s = re.sub(r"[\s　]+", "", str(s))
    s = re.sub(r"[Ａ-Ｚａ-ｚ０-９]", lambda m: chr(ord(m.group(0)) - 0xFEE0), s)
    return s.lower()


def _normalize_phone(s: str) -> str:
    return re.sub(r"[^0-9]", "", str(s))


def clean_name(name: str) -> str:
    name = re.sub(r"[\s　]+", " ", str(name)).strip()
    # 病院の場合「病院名 皮膚科」のように診療科が付与される。
    # 医会の会員一覧は皮膚科のみのため、冗長な診療科表記は除去する。
    name = re.sub(r"\s+(皮膚科|皮フ科|皮ふ科)$", "", name).strip()
    return name


def clean_address(addr: str) -> str:
    addr = re.sub(r"[\s　]+", " ", str(addr)).strip()
    if addr and not addr.startswith(PREF_NAME):
        addr = PREF_NAME + addr
    return addr


def clean_phone(phone: str) -> str:
    zen2han = str.maketrans("０１２３４５６７８９－", "0123456789-")
    return str(phone).translate(zen2han).strip()


def _page_title(html: str) -> str:
    m = re.search(r"<title[^>]*>(.*?)</title>", html, re.I | re.S)
    return re.sub(r"\s+", " ", m.group(1)).strip() if m else ""


def _is_portal_page(html: str) -> bool:
    title = _page_title(html)
    return any(kw in title for kw in PORTAL_TITLE_KWS)


def _verify_match(html_text: str, name: str, phone: str, address: str = "") -> bool:
    """検索で得た候補が本当にその医療機関の公式サイトかを検証する。

    電話番号一致を必須とし、その上で名称（または市区町村）による
    裏付けも要求する（口コミ・求人サイト等の誤マッチ防止）。
    """
    norm_page = _normalize_name(html_text)

    phone_digits = _normalize_phone(phone)
    phone_hit = bool(phone_digits) and len(phone_digits) >= 9 and phone_digits in re.sub(r"[^0-9]", "", html_text)
    if not phone_hit:
        return False

    norm_name = _normalize_name(name)
    name_hit = bool(norm_name) and norm_name in norm_page

    core = re.sub(r"(病院|医院|クリニック|診療所|センター)+$", "", name).strip()
    core_norm = _normalize_name(core)
    core_hit = len(core_norm) >= 2 and core_norm in norm_page

    if name_hit or core_hit:
        return True

    city_match = re.search(rf"{PREF_NAME}\s*([^\s]{{2,8}}?[市区町村])", address)
    city = city_match.group(1) if city_match else ""
    return bool(city) and _normalize_name(city) in norm_page


# ── Phase 1: 神奈川県皮膚科医会（kanahifu.org） ──────────

def fetch_kanahifu() -> list[dict]:
    logger.info("Phase 1: 神奈川県皮膚科医会（kanahifu.org）から会員医療機関を取得")
    r = requests.post(KANAHIFU_SEARCH_URL, headers=HEADERS, data={"kn": "on"}, timeout=30)
    r.encoding = r.apparent_encoding or "utf-8"
    soup = BeautifulSoup(r.text, "lxml")

    rows = soup.select("table.table_responsive tr")[1:]
    logger.info(f"  取得行数（会員医師単位・重複含む）: {len(rows)} 件")

    records: list[dict] = []
    for tr in rows:
        tds = tr.find_all("td")
        if len(tds) < 5:
            continue
        name = clean_name(tds[1].get_text(" ", strip=True))
        if not name:
            continue
        address = clean_address(tds[3].get_text(" ", strip=True))
        phone = clean_phone(tds[4].get_text(" ", strip=True))
        a = tds[1].find("a", href=True)
        url = a["href"].strip() if a else ""
        if url and _is_skip_domain(url):
            url = ""

        rec = {k: "" for k in OUTPUT_COLS}
        rec["名称"] = name
        rec["所在地"] = address
        rec["電話番号"] = phone
        rec["公式サイトURL"] = url
        records.append(rec)

    # 同一医療機関が複数の会員医師で重複掲載されるため重複排除する。
    # 電話番号は表記ゆれがなく同一施設で完全一致するため、名称＋電話番号を
    # 主キーとする（所在地はビル名表記等の細かな揺れがあり主キーに不向き）。
    seen: dict[tuple[str, str], dict] = {}
    order: list[tuple[str, str]] = []
    for rec in records:
        phone_key = _normalize_phone(rec["電話番号"]) or _normalize_name(rec["所在地"])
        key = (_normalize_name(rec["名称"]), phone_key)
        if key not in seen:
            seen[key] = rec
            order.append(key)
        else:
            if rec["公式サイトURL"] and not seen[key]["公式サイトURL"]:
                seen[key]["公式サイトURL"] = rec["公式サイトURL"]
            if len(rec["所在地"]) > len(seen[key]["所在地"]):
                seen[key]["所在地"] = rec["所在地"]

    deduped = [seen[k] for k in order]
    logger.info(
        f"Phase 1 完了: 重複排除後 {len(deduped)} 件"
        f"（公式サイトURL判明 {sum(1 for r in deduped if r['公式サイトURL'])} 件）"
    )
    return deduped


# ── Phase 3: DuckDuckGo 公式サイト検索+検証 ──────────────

def search_official_site(name: str, address: str, phone: str) -> tuple[str, str]:
    city_match = re.search(rf"{PREF_NAME}\s*([^\s]{{2,8}}?[市区町村])", address)
    city = city_match.group(1) if city_match else "神奈川"
    query = f"{name} {city} 皮膚科 公式サイト"
    results = []
    for backend in ("google", "bing", "brave", "duckduckgo"):
        try:
            with DDGS(timeout=8) as ddgs:
                results = list(ddgs.text(query, region="jp-jp", max_results=6, backend=backend))
            if results:
                break
        except Exception as e:
            logger.debug(f"検索失敗 {name} ({backend}): {e}")
            continue

    for res in results:
        url = res.get("href") or res.get("url") or ""
        if not url.startswith("http") or _is_skip_domain(url) or not _is_allowed(url):
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
        if _verify_match(resp.text, name, phone, address):
            return resp.url, resp.text
    return "", ""


# ── Phase 2 / 4: 公式サイトから追加情報取得 ───────────────

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
            result["インスタURL"] = href.split("?")[0].rstrip("/")
            break

    contact_kws = ["contact", "inquiry", "toiawase", "お問い合わせ", "問い合わせ",
                   "ご相談", "メールフォーム", "お問合せ", "問合せ", "予約"]
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


async def enrich_known_urls(records: list[dict], sem_count: int = 8) -> None:
    targets = [r for r in records if r.get("公式サイトURL")]
    if not targets:
        return
    logger.info(f"Phase 2: 既知の公式サイトから追加情報取得 ({len(targets)} 件)")
    semaphore = asyncio.Semaphore(sem_count)
    total = len(targets)
    done = {"n": 0}

    async def _one(rec: dict) -> None:
        async with semaphore:
            extras = await asyncio.to_thread(_fetch_official_sync, rec["公式サイトURL"])
            rec.update(extras)
            done["n"] += 1
            if done["n"] % 30 == 0 or done["n"] == total:
                logger.info(f"  公式サイト情報取得: {done['n']}/{total}")

    await asyncio.gather(*[_one(rec) for rec in targets])
    logger.info("Phase 2 完了")


async def enrich_search_records(records: list[dict], sem_count: int = 3) -> None:
    targets = [r for r in records if not r.get("公式サイトURL")]
    logger.info(f"Phase 3+4: 公式サイトURL未判明の医療機関を検索・検証 ({len(targets)} 件)")
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
                url, html = "", ""
            if url:
                rec["公式サイトURL"] = url
                rec.update(_parse_official_html(url, html))
                done["found"] += 1
            done["n"] += 1
            n = done["n"]
            if n % 20 == 0 or n == total:
                logger.info(f"  検索・検証: {n}/{total} (公式サイト発見 {done['found']} 件)")
            await asyncio.sleep(random.uniform(1.0, 2.0))

    await asyncio.gather(*[_one(r) for r in targets])
    logger.info("Phase 3+4 完了")


# ── メイン ────────────────────────────────────────────────

async def main() -> None:
    logger.info("=" * 65)
    logger.info("神奈川県 皮膚科 収集開始")
    logger.info("=" * 65)
    start_time = datetime.now()

    records = fetch_kanahifu()

    await enrich_known_urls(records, sem_count=8)
    await enrich_search_records(records, sem_count=3)
    logger.info(f"公式サイト取得 合計: {sum(1 for r in records if r.get('公式サイトURL'))} 件")

    df = pd.DataFrame(records, columns=OUTPUT_COLS)
    df.drop_duplicates(subset=["名称", "電話番号"], keep="first", inplace=True)
    for col in ["問い合わせフォームURL", "メールアドレス"]:
        mask_empty = df[col] == ""
        df = pd.concat([
            df[mask_empty],
            df[~mask_empty].drop_duplicates(subset=[col], keep="first"),
        ]).sort_index()
    df = df.sort_values("名称").reset_index(drop=True)
    df.to_excel(OUTPUT_FILE, index=False)

    n_site = (df["公式サイトURL"] != "").sum()
    n_mail = (df["メールアドレス"] != "").sum()
    n_ig = (df["インスタURL"] != "").sum()
    n_contact = (df["問い合わせフォームURL"] != "").sum()
    elapsed = int((datetime.now() - start_time).total_seconds())
    logger.info("=" * 65)
    logger.info(f"完了: {OUTPUT_FILE} に {len(df)} 件を出力")
    logger.info(f"  公式サイト {n_site} / メール {n_mail} / インスタ {n_ig} / 問い合わせフォーム {n_contact}")
    logger.info(f"所要時間: {elapsed // 60}分{elapsed % 60}秒")
    logger.info("=" * 65)


if __name__ == "__main__":
    asyncio.run(main())
