"""
福島県 精神科・心療内科 情報収集

データソース:
  Phase 1 - doctorsfile.jp から福島県の精神科・心療内科・神経精神科等を収集
            (https://doctorsfile.jp/search/ms61_ms62_ms78_ms93_ms94_pf7/)
            名称・所在地・電話番号を取得（公式サイトURLは掲載されていない）
  Phase 2 - DuckDuckGo で各院の公式サイトを検索し、ページ本文に名称または
            電話番号が含まれるかで一致検証（住所の市区町村の裏付けも必要な
            場合あり）。ディレクトリ/紹介/口コミ/官公庁サイトは除外。
  Phase 3 - 公式サイトから メール・インスタ・問い合わせフォームURL を取得

※ 医療機能情報提供制度（iryou.teikyouseido.mhlw.go.jp）は緯度経度＋距離半径
  方式の検索のみで都道府県境界での絞り込みができない（無指定だと全国
  8,000件超がヒットする）ため、データソースとしては採用せず、doctorsfile.jp
  の一覧を正としている。

出力列: 名称, メールアドレス, 公式サイトURL, 所在地, 電話番号, インスタURL, 問い合わせフォームURL
出力ファイル: fukushima_seishinka_YYYYMMDD_HHMMSS.xlsx
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

LOG_FILE = "scraper_fukushima_seishinka.log"
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
OUTPUT_FILE = f"fukushima_seishinka_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)
HEADERS = {"User-Agent": UA}

DOCTORSFILE_BASE = "https://doctorsfile.jp/search/ms61_ms62_ms78_ms93_ms94_pf7/"

SKIP_DOMAINS = [
    # データソース自身・同系列サイト
    "doctorsfile.jp", "hospitalsfile.doctorsfile.jp",
    # 口コミ・紹介・検索ディレクトリ
    "caloo.jp", "ekiten.jp", "qlife.jp", "medley.life", "byoinnavi.jp",
    "minds.jcqhc.or.jp", "e-doctor.co.jp", "10man-doc.jp", "10man-doc.co.jp",
    "kakaru.mynavi.jp", "clinic-navi.net", "clinicfor.life", "hospita.jp",
    "epark.jp", "junsuke.net", "mfs.jp", "minnanokaigo.com", "jiko24.jp",
    "oishasan.jp", "gurutto-aizu.com", "postmap.org", "opendata-japan.com",
    "gastro-health-now.org", "webkaigo.net", "cotohana.net",
    "doctor-concierge.jp", "fastdoctor.jp", "aeta-baby.jp", "itoshiihitoe.com",
    "fukushima-doctors.jp", "sanyokai-clinic.com", "kengikai.exblog.jp",
    "exblog.jp", "hospital.or.jp", "ajhc.or.jp", "jnss.or.jp",
    "e-resident.jp", "my-best.com", "kamponavi.com",
    "xn--db-vv5cr11j4rdj25d.com", "tokyo.asdj.org", "fertility-japan.com",
    "bigclear.org", "egg-room.com", "housingbazar.jp", "miyanet.net",
    "fukushima-kenchikutanbou.jp", "sync-sd0003.sakura.ne.jp",
    "hospia.jp", "ashitano.clinic", "remote-connect.jp", "funin-info.net",
    "toshi-ch.com",
    # ブログプラットフォーム（当事者本人以外の投稿が誤マッチしやすい）
    "ameblo.jp", "seesaa.net", "ldblog.jp", "livedoor.jp", "hatenablog.com",
    "blogspot.com",
    # 求人・転職・アルバイト情報サイト
    "im-nurse.com", "jp.indeed.com", "indeed.com", "jp.stanby.com",
    "stanby.com", "arubaito-ex.jp", "arubaito.sakura.ne.jp",
    "job.friendtree.co.jp", "friendtree.co.jp", "works.medical.nikkeibp.co.jp",
    "nikkeibp.co.jp", "f-kango.net",
    # 医師会・医療連携ネットワーク（個別医療機関の公式サイトではない）
    "fukushima.med.or.jp", "adachi-med.or.jp", "somagun.org",
    "f-renkei.net", "d-renkei.jp",
    # 官公庁・公的データベース（施設自体が公的機関である場合を除く）
    "iryou.teikyouseido.mhlw.go.jp", "mhlw.go.jp", "kaigokensaku.mhlw.go.jp",
    "wam.go.jp", "city.aizuwakamatsu.fukushima.jp", "city.minamisoma.lg.jp",
    # 情報・地図・検索サイト
    "itp.ne.jp", "mapion.co.jp", "navitime.co.jp", "mapfan.com",
    "goo.ne.jp", "wikipedia.org", "google.com", "google.co.jp",
    "maps.google.com", "bing.com", "yahoo.co.jp",
    # SNS
    "instagram.com", "facebook.com", "twitter.com", "x.com",
    "youtube.com", "tiktok.com", "line.me", "lin.ee",
]

# 上記の個別ドメイン列挙に加え、URLに含まれると医療ディレクトリ／求人・転職／
# 医療連携ネットワークサイトである可能性が非常に高いパターン
SKIP_URL_PATTERNS = [
    r"\.med\.or\.jp",   # 都道府県・郡市医師会
    r"renkei",           # 医療連携ネットワークシステム（例: f-renkei.net, d-renkei.jp）
    r"kuchikomi",         # 口コミサイト
]

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
    return _robots_cache[origin].can_fetch(UA, url)


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


def _is_skip_domain(url: str) -> bool:
    if any(d in url for d in SKIP_DOMAINS):
        return True
    return any(re.search(p, url, re.I) for p in SKIP_URL_PATTERNS)


def _verify_match(html_text: str, name: str, phone: str, address: str = "") -> bool:
    """取得した公式サイトが実際にその院のものかを名称・電話番号・所在地で検証する。

    医療機関のディレクトリ・職能団体の会員一覧・求人サイト・地域ブログ等は
    正確な名称や電話番号をそのまま掲載していることが多く、名称一致または
    電話番号一致のどちらか単独では無関係サイトとの誤マッチを防げない。
    そのため電話番号一致を必須とし、その上で名称（または市区町村）による
    裏付けも要求する。
    """
    norm_page = _normalize_name(re.sub(r"[\s　]+", "", html_text))

    phone_digits = _normalize_phone(phone)
    phone_hit = bool(phone_digits) and phone_digits in re.sub(r"[^0-9]", "", html_text)
    if not phone_hit:
        return False

    norm_name = _normalize_name(name)
    name_hit = bool(norm_name) and norm_name in norm_page

    core = re.sub(
        r"(病院|医院|クリニック|診療所|センター|治療院)+$",
        "", name
    ).strip()
    core_norm = _normalize_name(core)
    core_hit = len(core_norm) >= 2 and core_norm in norm_page

    if name_hit or core_hit:
        return True

    city_match = re.search(r"([^\s　]{2,8}?[市区町村])", address)
    city = city_match.group(1) if city_match else ""
    return bool(city) and _normalize_name(city) in norm_page


# 求人・転職・口コミ・ランキング・医療連携等のポータル/ディレクトリサイトに
# ありがちなキーワード。個別医療機関の公式サイトとして未知のディレクトリを
# 誤採用しないための、ドメイン列挙とは別の防御層。
PORTAL_TITLE_KWS = [
    "求人", "アルバイト", "転職", "口コミ", "クチコミ", "ランキング",
    "医療連携", "医師会", "地域医療連携", "一覧｜", "検索｜", "を探す",
    "施設検索", "病院検索", "クリニック検索", "ナビ｜", "まとめ",
]


def _is_portal_page(html: str) -> bool:
    title_m = re.search(r"<title[^>]*>(.*?)</title>", html, re.I | re.S)
    title = title_m.group(1) if title_m else ""
    return any(kw in title for kw in PORTAL_TITLE_KWS)


# ── Phase 1: doctorsfile.jp ──────────────────────────────

def scrape_doctorsfile_page(page_num: int) -> list[dict]:
    url = DOCTORSFILE_BASE if page_num == 1 else f"{DOCTORSFILE_BASE}page/{page_num}/"
    soup = _fetch(url)
    if soup is None:
        return []

    records: list[dict] = []
    for item in soup.select("div.result"):
        name_a = item.select_one("a.result__name")
        if not name_a:
            continue
        name = name_a.get_text(strip=True)

        address = ""
        area_icon = item.select_one("li.result-data__list i.ico-area-gray")
        if area_icon:
            address = area_icon.parent.get_text(strip=True)

        phone = ""
        tel_icon = item.select_one("li.result-data__list i.ico-tel-gray")
        if tel_icon:
            phone = tel_icon.parent.get_text(strip=True)

        if not name:
            continue
        rec = {k: "" for k in OUTPUT_COLS}
        rec["名称"] = name
        rec["所在地"] = address
        rec["電話番号"] = phone
        records.append(rec)

    return records


def scrape_doctorsfile() -> list[dict]:
    logger.info("Phase 1: doctorsfile.jp から福島県 精神科・心療内科を収集")
    first_soup = _fetch(DOCTORSFILE_BASE)
    if first_soup is None:
        return []

    m = re.search(r"(\d+)件中", first_soup.get_text())
    total = int(m.group(1)) if m else None
    total_pages = (total + 19) // 20 if total else 1
    logger.info(f"  対象: {total} 件 ({total_pages} ページ)")

    all_records: list[dict] = []
    for page_num in range(1, total_pages + 1):
        recs = scrape_doctorsfile_page(page_num)
        all_records.extend(recs)
        logger.info(f"  page {page_num}/{total_pages}: {len(recs)} 件 (累計 {len(all_records)} 件)")
        time.sleep(random.uniform(0.8, 1.5))

    logger.info(f"Phase 1 完了: {len(all_records)} 件")
    return all_records


# ── Phase 2: DuckDuckGo 公式サイト検索+検証 ──────────────

def search_official_site(name: str, address: str, phone: str) -> tuple[str, str]:
    city_match = re.search(r"福島県\s*([^\s]{2,8}?[市区町村])", address)
    city = city_match.group(1) if city_match else "福島"
    query = f"{name} {city} 公式サイト"
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
        if _verify_match(resp.text, name, phone, address):
            return resp.url, resp.text
    return "", ""


# ── Phase 3: 公式サイト追加情報 ──────────────────────────

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

    return result


async def enrich_records(records: list[dict], sem_count: int = 3) -> None:
    semaphore = asyncio.Semaphore(sem_count)
    total = len(records)
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
                rec.update(_parse_official_html(url, html))
                done["found"] += 1
            done["n"] += 1
            n = done["n"]
            if n % 20 == 0 or n == total:
                logger.info(f"  検索・検証: {n}/{total} (公式サイト発見 {done['found']} 件)")
            await asyncio.sleep(random.uniform(1.0, 2.0))

    await asyncio.gather(*[_one(rec) for rec in records])


# ── メイン ────────────────────────────────────────────────

async def main() -> None:
    logger.info("=" * 65)
    logger.info("福島県 精神科・心療内科 収集開始")
    logger.info("=" * 65)
    start_time = datetime.now()

    all_records = scrape_doctorsfile()

    seen_names: set[str] = set()
    seen_phones: set[str] = set()
    deduped: list[dict] = []
    for rec in all_records:
        nk = _normalize_name(rec["名称"])
        pk = _normalize_phone(rec.get("電話番号", ""))
        if nk and nk in seen_names:
            continue
        if pk and pk in seen_phones:
            continue
        seen_names.add(nk)
        if pk:
            seen_phones.add(pk)
        deduped.append(rec)
    logger.info(f"重複除去後: {len(deduped)} 件")

    logger.info("Phase 2+3: DuckDuckGo で公式サイト検索・検証・情報取得")
    await enrich_records(deduped, sem_count=3)
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
