"""
岡山県 精神科・心療内科 情報収集

データソース:
  Phase 1 - 岡山県精神保健福祉センター（okayama-pmc.jp）公開の医療機関一覧PDF
            「精神病床を有する病院」「精神科又は心療内科を有する病院」
            「精神科又は心療内科を有する診療所」の3表から名称・所在地・電話番号を収集。
            PDFは表の行がテキスト順に並ばないため、文字の座標から行と列を復元する。
            （「精神科デイケア等施設」の表は上記3表と施設が重複するため対象外）
  Phase 2 - doctorsfile.jp から岡山県の精神科・心療内科・神経精神科等を収集
            (https://doctorsfile.jp/search/ms61_ms62_ms78_ms93_ms94_pf33/)
  Phase 3 - GMOクリニック・マップ から岡山県の心療内科・精神科・児童精神科を収集
            一覧ページの __NEXT_DATA__ に公式サイトURLが含まれるため、それを採用する。
            電話番号は一覧に無いため、Phase 1/2 と名寄せできなかった施設のみ
            詳細ページから取得する（robots.txt の Crawl-delay: 10 を遵守）。
  Phase 4 - 公式サイトURLが得られなかった施設について DuckDuckGo で検索し、
            電話番号一致を必須とする検証を行う（千葉県版の検証ロジックを流用）。
  Phase 5 - 公式サイトから メール・インスタ・問い合わせフォームURL を取得
  ※ Googleマップは利用規約上スクレイピングできないため、件数の目安としてのみ参照。

出力列: 名称, メールアドレス, 公式サイトURL, 所在地, 電話番号, インスタURL, 問い合わせフォームURL
出力ファイル: okayama_seishinka_YYYYMMDD_HHMMSS.xlsx
"""

import asyncio
import io
import json
import logging
import random
import re
import sys
import time
from datetime import datetime
from urllib.parse import urljoin, urlparse
from urllib.robotparser import RobotFileParser

import pandas as pd
import pymupdf
import requests
from bs4 import BeautifulSoup
from ddgs import DDGS

# ── ロギング ──────────────────────────────────────────────
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

LOG_FILE = "scraper_okayama_seishinka.log"
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
OUTPUT_FILE = f"okayama_seishinka_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)
HEADERS = {"User-Agent": UA}

PMC_PDF_URL = (
    "https://www.okayama-pmc.jp/wp-content/uploads/2026/06/"
    "a13011ee3ee0145416440282680d8c96.pdf"
)
DOCTORSFILE_BASE = "https://doctorsfile.jp/search/ms61_ms62_ms78_ms93_ms94_pf33/"
GMO_BASE = "https://gmo-clinic-map.com"
# 12: 心療内科 / 14: 精神科 / 79: 児童精神科（80: 老年精神科は岡山県0件）
GMO_TREATS = [12, 14, 79]
GMO_DELAY = 10  # robots.txt の Crawl-delay

SKIP_DOMAINS = [
    # データソース自身・同系列サイト
    "doctorsfile.jp", "hospitalsfile.doctorsfile.jp", "gmo-clinic-map.com",
    "okayama-pmc.jp",
    # 口コミ・紹介・検索ディレクトリ
    "caloo.jp", "ekiten.jp", "qlife.jp", "medley.life", "byoinnavi.jp",
    "minds.jcqhc.or.jp", "e-doctor.co.jp", "10man-doc.jp", "10man-doc.co.jp",
    "kakaru.mynavi.jp", "clinic-navi.net", "clinicfor.life", "hospita.jp",
    "epark.jp", "junsuke.net", "mfs.jp", "minnanokaigo.com", "jiko24.jp",
    "oishasan.jp", "postmap.org", "opendata-japan.com",
    "gastro-health-now.org", "webkaigo.net", "cotohana.net",
    "doctor-concierge.jp", "fastdoctor.jp", "aeta-baby.jp", "itoshiihitoe.com",
    "hospital.or.jp", "ajhc.or.jp", "jnss.or.jp", "e-resident.jp", "my-best.com",
    "kamponavi.com", "fertility-japan.com", "hospia.jp", "ashitano.clinic",
    "remote-connect.jp", "toshi-ch.com", "mame-clinic.net", "medicaldoc.jp",
    "mamamap.jp", "eisai.jp", "mynavi.jp", "okayama-doctors.com",
    "shinryounaika.com", "kokoro-clinic.jp", "utsu-net.jp", "mental-navi.net",
    "medicalnote.jp", "scuel.me", "byoin.ne.jp", "iryou-navi.com",
    "okayama-iryo.jp", "hareiro.jp", "jimoken.jp", "ritoful.com",
    "byoin-machi.net",
    # ブログプラットフォーム（当事者本人以外の投稿が誤マッチしやすい）
    "ameblo.jp", "seesaa.net", "ldblog.jp", "livedoor.jp", "hatenablog.com",
    "blogspot.com", "exblog.jp",
    # 求人・転職・アルバイト情報サイト
    "im-nurse.com", "jp.indeed.com", "indeed.com", "jp.stanby.com",
    "stanby.com", "arubaito-ex.jp", "job.friendtree.co.jp", "friendtree.co.jp",
    "works.medical.nikkeibp.co.jp", "nikkeibp.co.jp", "f-kango.net",
    "arubaito.sakura.ne.jp", "job-gear.net",
    "guppy.jp", "job-medley.com", "kango-roo.com", "nursejinzaibank.com",
    # 官公庁・公的データベース
    "iryou.teikyouseido.mhlw.go.jp", "mhlw.go.jp", "kaigokensaku.mhlw.go.jp",
    "wam.go.jp",
    # 情報・地図・検索サイト
    "itp.ne.jp", "mapion.co.jp", "navitime.co.jp", "mapfan.com",
    "goo.ne.jp", "wikipedia.org", "google.com", "google.co.jp",
    "maps.google.com", "bing.com", "yahoo.co.jp",
    # SNS
    "instagram.com", "facebook.com", "twitter.com", "x.com",
    "youtube.com", "tiktok.com", "line.me", "lin.ee",
]

SKIP_URL_PATTERNS = [
    r"\.med\.or\.jp",   # 都道府県・郡市医師会
    r"renkei",           # 医療連携ネットワークシステム
    r"kuchikomi",         # 口コミサイト
]

OUTPUT_COLS = [
    "名称", "メールアドレス", "公式サイトURL",
    "所在地", "電話番号", "インスタURL", "問い合わせフォームURL",
]

# 名寄せ用に名称先頭から取り除く法人格（「医療法人○○会　」等）
CORP_PREFIX_RE = re.compile(
    r"^(?:(?:地方)?独立行政法人|社会医療法人(?:財団|社団)?|医療法人(?:財団|社団)?|"
    r"特定医療法人|公益財団法人|一般財団法人|公益社団法人|一般社団法人|"
    r"社会福祉法人|医療生活協同組合)"
    r"[\s　]*(?:[^\s　]+?(?:会|団|機構)[\s　]*)?"
)

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


def _fetch_text(url: str, timeout: int = 15) -> str | None:
    try:
        r = requests.get(url, headers=HEADERS, timeout=timeout)
        if r.status_code >= 400:
            return None
        r.encoding = r.apparent_encoding or "utf-8"
        return r.text
    except Exception as e:
        logger.debug(f"取得失敗: {url} / {e}")
        return None


def _fetch(url: str, timeout: int = 15) -> BeautifulSoup | None:
    text = _fetch_text(url, timeout)
    return BeautifulSoup(text, "lxml") if text is not None else None


def _zen_to_han(s: str) -> str:
    s = re.sub(r"[Ａ-Ｚａ-ｚ０-９]", lambda m: chr(ord(m.group(0)) - 0xFEE0), str(s))
    # 長音「ー」は名称に使われるため変換しない
    return s.translate(str.maketrans("－―‐（）　", "---() "))


def _normalize_name(s: str) -> str:
    s = re.sub(r"[\s　]+", "", _zen_to_han(s))
    return s.lower()


def _core_name(s: str) -> str:
    """名寄せ用: 法人格を除いた施設名（「医療法人○○会 △△クリニック」→「△△クリニック」）"""
    s = _zen_to_han(str(s)).strip()
    # 「医療法人宏仁会 医療法人宏仁会 まつうらクリニック」のような重複表記にも対応
    prev = None
    while prev != s:
        prev, s = s, CORP_PREFIX_RE.sub("", s)
    return _normalize_name(s)


def _normalize_phone(s: str) -> str:
    return re.sub(r"[^0-9]", "", _zen_to_han(s))


def _format_phone(s: str) -> str:
    phone = re.sub(r"[^0-9\-]", "", _zen_to_han(s))
    return phone if re.search(r"\d", phone) else ""


def _address_key(s: str) -> str:
    """名寄せ用の住所キー（市区町村以降・番地まで、表記ゆれを吸収）"""
    s = _zen_to_han(s)
    s = re.sub(r"\s+", "", s)
    s = re.sub(r"^〒?\d{3}-?\d{4}", "", s)
    s = s.replace("岡山県", "")
    s = re.sub(r"(\d+)丁目", r"\1-", s)
    s = re.sub(r"(\d+)番地?の?", r"\1-", s)
    s = re.sub(r"(\d+)号", r"\1", s)
    s = re.sub(r"字", "", s)
    m = re.match(r"(.+?\d+(?:-\d+)*)", s)
    return (m.group(1) if m else s).rstrip("-")


def _is_skip_domain(url: str) -> bool:
    if any(d in url for d in SKIP_DOMAINS):
        return True
    return any(re.search(p, url, re.I) for p in SKIP_URL_PATTERNS)


def _verify_match(html_text: str, name: str, phone: str, address: str = "") -> bool:
    """取得した公式サイトが実際にその院のものかを名称・電話番号・所在地で検証する。

    電話番号一致を必須とし、その上で名称（または市区町村）による裏付けも要求する。
    """
    norm_page = _normalize_name(html_text)

    phone_digits = _normalize_phone(phone)
    phone_hit = bool(phone_digits) and phone_digits in _normalize_phone(html_text)
    if not phone_hit:
        return False

    norm_name = _core_name(name)
    name_hit = bool(norm_name) and norm_name in norm_page

    core = re.sub(r"(病院|医院|クリニック|診療所|センター|ホスピタル)+$", "", norm_name)
    core_hit = len(core) >= 2 and core in norm_page

    if name_hit or core_hit:
        return True

    city_match = re.search(r"([^\s　]{2,8}?[市区町村])", address)
    city = city_match.group(1) if city_match else ""
    return bool(city) and _normalize_name(city) in norm_page


PORTAL_TITLE_KWS = [
    "求人", "アルバイト", "転職", "口コミ", "クチコミ", "ランキング",
    "医療連携", "医師会", "地域医療連携", "一覧｜", "検索｜", "を探す",
    "施設検索", "病院検索", "クリニック検索", "ナビ｜", "まとめ",
]


def _is_portal_page(html: str) -> bool:
    title_m = re.search(r"<title[^>]*>(.*?)</title>", html, re.I | re.S)
    title = title_m.group(1) if title_m else ""
    return any(kw in title for kw in PORTAL_TITLE_KWS)


def _new_record(name: str, address: str, phone: str, url: str = "") -> dict:
    rec = {k: "" for k in OUTPUT_COLS}
    rec["名称"] = re.sub(r"[\s　]+", " ", name).strip()
    rec["所在地"] = address
    rec["電話番号"] = _format_phone(phone)
    rec["公式サイトURL"] = url
    return rec


# ── Phase 1: 岡山県精神保健福祉センター PDF ─────────────

ZIP_RE = re.compile(r"^\d{3}-\d{4}$")
PHONE_RE = re.compile(r"^0\d{1,4}-\d{1,4}-\d{3,4}$")


def _pdf_rows(page) -> list[list[tuple[float, str]]]:
    """PDFページの文字列を y 座標で行にまとめ、各行を x 座標順に並べる。"""
    spans = []
    for b in page.get_text("dict")["blocks"]:
        for line in b.get("lines", []):
            for sp in line["spans"]:
                t = sp["text"].strip()
                if t:
                    spans.append((sp["bbox"][1], sp["bbox"][0], t))
    spans.sort()
    rows: list[list] = []
    for y, x, t in spans:
        if rows and abs(rows[-1][0] - y) <= 3:
            rows[-1][1].append((x, t))
        else:
            rows.append([y, [(x, t)]])
    return [sorted(items) for _, items in rows]


def _parse_pdf_table_row(items: list[tuple[float, str]]) -> list[dict]:
    """1行の中から（名称, 郵便番号, 所在地, 電話番号）の組を抽出する。

    1行に左右2施設が並ぶ表があるため、郵便番号を起点に
    直前の要素を名称、直後を所在地、その後の電話番号形式の要素を電話番号とする。
    """
    texts = [t for _, t in items]
    results = []
    for i, t in enumerate(texts):
        if not ZIP_RE.match(_zen_to_han(t)):
            continue
        if i == 0:
            continue
        name = texts[i - 1]
        address, phone = "", ""
        for t2 in texts[i + 1:i + 3]:
            t2h = _zen_to_han(t2)
            if PHONE_RE.match(t2h):
                phone = t2h
                break
            if not address:
                address = t2
        results.append({"name": name, "zip": t, "address": address, "phone": phone})
    return results


def scrape_pmc_pdf() -> list[dict]:
    logger.info("Phase 1: 岡山県精神保健福祉センター 医療機関一覧PDFから収集")
    try:
        r = requests.get(PMC_PDF_URL, headers=HEADERS, timeout=30)
        r.raise_for_status()
    except Exception as e:
        logger.warning(f"PDF取得失敗: {e}")
        return []

    doc = pymupdf.open(stream=r.content, filetype="pdf")
    records: list[dict] = []
    for page in doc:
        rows = _pdf_rows(page)
        section = ""
        for items in rows:
            line = " ".join(t for _, t in items)
            if "精神病床を有する病院" in line or "精神科又は心療内科を有する" in line:
                section = line.split()[0]
                continue
            if "デイケア等施設" in line:
                section = ""  # 他表と重複するため対象外
                continue
            if not section:
                continue
            for row in _parse_pdf_table_row(items):
                name = row["name"]
                # 名称の直前に所在地域（市区町村名）しかない場合の除外
                if re.fullmatch(r"\d+", name):
                    continue
                address = _zen_to_han(row["address"]).strip()
                if address and not address.startswith("岡山県"):
                    address = "岡山県" + address
                # PDFでは長音が全角ハイフンで書かれている箇所がある（エスポア－ル）
                name = re.sub(r"(?<=[ァ-ヶ])－", "ー", name)
                records.append(_new_record(_zen_to_han(name), address, row["phone"]))
    logger.info(f"Phase 1 完了: {len(records)} 件")
    return records


# ── Phase 2: doctorsfile.jp ──────────────────────────────

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
        if name:
            records.append(_new_record(name, address, phone))
    return records


def scrape_doctorsfile() -> list[dict]:
    logger.info("Phase 2: doctorsfile.jp から岡山県 精神科・心療内科を収集")
    first_soup = _fetch(DOCTORSFILE_BASE)
    if first_soup is None:
        return []
    m = re.search(r"(\d+)\s*件中", first_soup.get_text())
    total = int(m.group(1)) if m else None
    total_pages = (total + 19) // 20 if total else 1
    logger.info(f"  対象: {total} 件 ({total_pages} ページ)")

    all_records: list[dict] = []
    for page_num in range(1, total_pages + 1):
        recs = scrape_doctorsfile_page(page_num)
        all_records.extend(recs)
        logger.info(f"  page {page_num}/{total_pages}: {len(recs)} 件 (累計 {len(all_records)} 件)")
        time.sleep(random.uniform(2.0, 3.0))  # robots.txt Crawl-delay: 2
    logger.info(f"Phase 2 完了: {len(all_records)} 件")
    return all_records


# ── Phase 3: GMOクリニック・マップ ───────────────────────

def _gmo_next_data(url: str) -> dict | None:
    text = _fetch_text(url)
    if text is None:
        return None
    m = re.search(r'<script id="__NEXT_DATA__"[^>]*>(.*?)</script>', text, re.S)
    if not m:
        return None
    try:
        return json.loads(m.group(1))["props"]["pageProps"]
    except Exception:
        return None


def scrape_gmo() -> list[dict]:
    logger.info("Phase 3: GMOクリニック・マップから岡山県 心療内科・精神科を収集")
    by_id: dict[str, dict] = {}
    for treat in GMO_TREATS:
        page = 1
        while True:
            url = f"{GMO_BASE}/clinics/pref33_treat{treat}" + (f"?page={page}" if page > 1 else "")
            pp = _gmo_next_data(url)
            time.sleep(GMO_DELAY)
            hospitals = (pp or {}).get("searchResult", {}).get("hospitals", [])
            if not hospitals:
                break
            for h in hospitals:
                sid = h.get("scuelId")
                if not sid or sid in by_id:
                    continue
                homepage = (h.get("homepage") or "").strip()
                if homepage and _is_skip_domain(homepage):
                    homepage = ""
                address = (h.get("address1") or "") + (h.get("address2") or "")
                rec = _new_record(h.get("name", ""), _zen_to_han(address).strip(), "", homepage)
                rec["_gmo_id"] = sid
                by_id[sid] = rec
            logger.info(f"  treat{treat} page {page}: {len(hospitals)} 件 (累計 {len(by_id)} 件)")
            if len(hospitals) < 20:
                break
            page += 1
    logger.info(f"Phase 3 完了: {len(by_id)} 件")
    return list(by_id.values())


def fetch_gmo_phone(sid: str) -> str:
    text = _fetch_text(f"{GMO_BASE}/hospital/ref/{sid}")
    if not text:
        return ""
    m = re.search(r'"telephoneNumber"\s*:\s*"([^"]+)"', text)
    return _format_phone(m.group(1)) if m else ""


# ── 名寄せ ────────────────────────────────────────────────

class Merger:
    def __init__(self) -> None:
        self.records: list[dict] = []
        self.by_phone: dict[str, dict] = {}
        self.by_name: dict[str, dict] = {}
        self.by_addr: dict[str, list[dict]] = {}

    def find(self, rec: dict) -> dict | None:
        pk = _normalize_phone(rec.get("電話番号", ""))
        if pk and pk in self.by_phone:
            return self.by_phone[pk]
        nk = _core_name(rec["名称"])
        if nk and nk in self.by_name:
            return self.by_name[nk]
        ak = _address_key(rec.get("所在地", ""))
        if ak:
            for other in self.by_addr.get(ak, []):
                # 同一住所の別施設（医療ビル等）を誤って統合しないよう名称の包含も要求
                on = _core_name(other["名称"])
                if nk and on and (nk in on or on in nk):
                    return other
        return None

    def _index(self, rec: dict) -> None:
        pk = _normalize_phone(rec.get("電話番号", ""))
        if pk:
            self.by_phone.setdefault(pk, rec)
        nk = _core_name(rec["名称"])
        if nk:
            self.by_name.setdefault(nk, rec)
        ak = _address_key(rec.get("所在地", ""))
        if ak:
            self.by_addr.setdefault(ak, []).append(rec)

    def add(self, rec: dict) -> bool:
        """新規なら追加して True、既存と同一なら不足項目のみ補完して False。"""
        existing = self.find(rec)
        if existing is None:
            self.records.append(rec)
            self._index(rec)
            return True
        for col in ("所在地", "電話番号", "公式サイトURL"):
            if not existing.get(col) and rec.get(col):
                existing[col] = rec[col]
        self._index(existing)
        return False


# ── Phase 4: DuckDuckGo 公式サイト検索+検証 ──────────────

def search_official_site(name: str, address: str, phone: str) -> tuple[str, str]:
    city_match = re.search(r"岡山県\s*([^\s]{2,8}?[市町村])", address)
    city = city_match.group(1) if city_match else "岡山"
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
            return _prefer_top_page(resp.url, resp.text, name, phone, address)
    return "", ""


def _prefer_top_page(url: str, html: str, name: str, phone: str, address: str) -> tuple[str, str]:
    """検索で添付ファイルページ等のサブページが当たった場合、
    同じサイトのトップページでも検証が通ればトップページを採用する。"""
    parsed = urlparse(url)
    top = f"{parsed.scheme}://{parsed.netloc}/"
    if parsed.path in ("", "/"):
        return url, html
    try:
        resp = requests.get(top, headers=HEADERS, timeout=10, allow_redirects=True)
        if resp.status_code < 400:
            resp.encoding = resp.apparent_encoding or "utf-8"
            if _verify_match(resp.text, name, phone, address):
                return resp.url, resp.text
    except Exception:
        pass
    return url, html


# ── Phase 5: 公式サイト追加情報 ──────────────────────────

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
        mailtos = [a["href"][7:].split("?")[0] for a in soup.find_all("a", href=True)
                   if a["href"].lower().startswith("mailto:")]
        emails = mailtos + re.findall(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}", text)
        emails = [e for e in emails
                  if "@" in e and not re.search(r"\.(png|jpg|gif|svg|webp)$", e, re.I)
                  and not re.search(r"example\.|sentry|wixpress", e, re.I)]
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
    """GMOで公式サイトURLが判明している施設について追加情報のみ取得する。"""
    targets = [r for r in records if r.get("公式サイトURL")]
    if not targets:
        return
    semaphore = asyncio.Semaphore(sem_count)
    total = len(targets)
    done = {"n": 0}

    async def _one(rec: dict) -> None:
        async with semaphore:
            extras = await asyncio.to_thread(_fetch_official_sync, rec["公式サイトURL"])
            rec.update(extras)
            done["n"] += 1
            if done["n"] % 20 == 0 or done["n"] == total:
                logger.info(f"  既知公式サイト情報取得: {done['n']}/{total}")

    await asyncio.gather(*[_one(rec) for rec in targets])


async def enrich_search_records(records: list[dict], sem_count: int = 3) -> None:
    """公式サイトURLが未判明の施設について DuckDuckGo で検索・検証する。"""
    targets = [r for r in records if not r.get("公式サイトURL") and r.get("電話番号")]
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
                rec.update(_parse_official_html(url, html))
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
    logger.info("岡山県 精神科・心療内科 収集開始")
    logger.info("=" * 65)
    start_time = datetime.now()

    pmc_raw = scrape_pmc_pdf()
    doctorsfile_raw = scrape_doctorsfile()
    gmo_raw = scrape_gmo()

    merger = Merger()
    counts = {}
    # 県のPDFを優先（公的機関の一覧で名称・電話番号が正確なため）
    for label, raw in (("PDF", pmc_raw), ("doctorsfile", doctorsfile_raw)):
        counts[label] = sum(merger.add(rec) for rec in raw)

    gmo_new: list[dict] = []
    for rec in gmo_raw:
        if merger.find(rec) is None:
            gmo_new.append(rec)
        else:
            merger.add(rec)

    # GMOのみに掲載の施設は電話番号を詳細ページから取得してから名寄せする
    # （電話番号で既存施設と同一と判明する場合があるため）
    logger.info(f"GMOのみ掲載 {len(gmo_new)} 件の電話番号を詳細ページから取得")
    for i, rec in enumerate(gmo_new, 1):
        if _is_allowed(f"{GMO_BASE}/hospital/ref/{rec['_gmo_id']}"):
            rec["電話番号"] = fetch_gmo_phone(rec["_gmo_id"])
        time.sleep(GMO_DELAY)
        if i % 10 == 0 or i == len(gmo_new):
            logger.info(f"  GMO詳細: {i}/{len(gmo_new)}")
    counts["GMO"] = sum(merger.add(rec) for rec in gmo_new)

    deduped = merger.records
    logger.info(
        f"名寄せ後: {len(deduped)} 件 (PDF {counts['PDF']} / doctorsfile {counts['doctorsfile']}"
        f" / GMO {counts['GMO']})"
    )
    logger.info(f"  うちGMOで公式サイトURL判明: {sum(1 for r in deduped if r.get('公式サイトURL'))} 件")

    logger.info("Phase 5: 既知の公式サイトから追加情報取得")
    await enrich_known_urls(deduped, sem_count=8)

    logger.info("Phase 4+5: 未判明の施設を DuckDuckGo で検索・検証・情報取得")
    await enrich_search_records(deduped, sem_count=3)
    logger.info(f"公式サイト取得 合計: {sum(1 for r in deduped if r.get('公式サイトURL'))} 件")

    df = pd.DataFrame(deduped, columns=OUTPUT_COLS)
    df.drop_duplicates(subset=["名称", "電話番号"], keep="first", inplace=True)
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
