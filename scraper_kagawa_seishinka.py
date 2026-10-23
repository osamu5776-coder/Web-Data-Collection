"""
香川県 精神科・心療内科 情報収集

データソース:
  Phase 1 - 医療情報ネット（ナビイ／厚生労働省）で香川県の17市町ごとに
            「精神科・神経科・心療内科・老年精神科・老年心療内科・小児精神科・
            その他（精神科系）のいずれかに該当する」病院・診療所を検索し、
            名称・所在地・電話番号・URL・診療科目を収集（県単位の検索結果と件数を照合）。
            都道府県へ届け出られた情報のため、これを掲載可否の正とする。
  Phase 1b - 香川県「発達障害医療機関リスト」PDF（県医師会協力のアンケート結果）。
            小児科・小児神経科等のみの施設も含むため、診療科名に精神・心療・メンタル・
            こころを含む施設（診療科名欄が空なら施設名で判定）のみ採用する。
            表は縦中央寄せのセルが多いため、番号列の位置を基準に行範囲を決めて復元する。
  Phase 2 - GMOクリニック・マップ から香川県の心療内科・精神科・児童精神科を収集し、
            Phase 1 の施設の公式サイトURLを補完する。GMOの「精神科・神経科」分類は
            神経内科や総合病院にも付与されており不正確なため（岡山県版で判明）、
            GMOのみに掲載の施設は、公式サイトに精神科・心療内科の記載がある、GMOの分類が
            児童精神科・心療内科等の具体的な精神系科目である、または施設名に「メンタル」
            「こころ」「心療」を含む場合のみ追加する。
  Phase 3 - 公式サイトURLが得られなかった施設について DuckDuckGo で検索し、
            電話番号一致を必須とする検証を行う。
  Phase 4 - 公式サイトから メール・インスタ・問い合わせフォームURL を取得

出力列: 名称, メールアドレス, 公式サイトURL, 所在地, 電話番号, インスタURL, 問い合わせフォームURL
（精神系診療科目は掲載可否の判定とログ出力にのみ使用し、Excelには出力しない）
出力ファイル: kagawa_seishinka_YYYYMMDD_HHMMSS.xlsx
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
from urllib.parse import urlencode, urljoin, urlparse
from urllib.robotparser import RobotFileParser

import pandas as pd
import pymupdf
import requests
from bs4 import BeautifulSoup
from ddgs import DDGS

# ── ロギング ──────────────────────────────────────────────
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

LOG_FILE = "scraper_kagawa_seishinka.log"
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
OUTPUT_FILE = f"kagawa_seishinka_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)
HEADERS = {"User-Agent": UA}

PREF_NAME = "香川県"
PREF_CD = "37"
MHLW_BASE = "https://www.iryou.teikyouseido.mhlw.go.jp"
MHLW_SEARCH = f"{MHLW_BASE}/znk-web/juminkanja/S2400/initialize"
# 診療科目コード: 精神科 / 神経科 / 心療内科 / 老年精神科 / 老年心療内科 / 小児精神科 / その他（精神科系）
MHLW_DEPT_CODES = [
    "01-03-7-07001", "01-03-7-07002", "01-03-7-07003", "01-03-7-07004",
    "01-03-7-07005", "01-03-3-03005", "01-03-7-07991",
]
# 香川県の市町（全国地方公共団体コード）
MUNICIPALITIES = {
    "37201": "高松市", "37202": "丸亀市", "37203": "坂出市", "37204": "善通寺市",
    "37205": "観音寺市", "37206": "さぬき市", "37207": "東かがわ市", "37208": "三豊市",
    "37322": "土庄町", "37324": "小豆島町", "37341": "三木町", "37364": "直島町",
    "37386": "宇多津町", "37387": "綾川町", "37403": "琴平町", "37404": "多度津町",
    "37406": "まんのう町",
}
MHLW_DELAY = 2.0

HATTATSU_PDF_URL = "https://www.pref.kagawa.lg.jp/documents/6509/hattatu.pdf"
HATTATSU_PSY_DEPT_RE = re.compile(r"精神|心療|メンタル|こころ")
HATTATSU_PSY_NAME_RE = re.compile(r"メンタル|こころ|心療")

GMO_BASE = "https://gmo-clinic-map.com"
# 12: 心療内科 / 14: 精神科 / 79: 児童精神科 / 80: 老年精神科
GMO_TREATS = [12, 14, 79, 80]
GMO_DELAY = 10  # robots.txt の Crawl-delay
GMO_SPECIFIC_PSY_CATS = {"心療内科", "児童精神科", "老年精神科"}

PSY_DEPT_RE = re.compile(r"精神|心療|神経科$")
PSY_SITE_RE = re.compile(r"精神科|心療内科|精神神経科")

SKIP_DOMAINS = [
    # データソース自身・同系列サイト
    "doctorsfile.jp", "gmo-clinic-map.com", "teikyouseido.mhlw.go.jp",
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
    "mamamap.jp", "eisai.jp", "mynavi.jp", "shinryounaika.com", "kokoro-clinic.jp",
    "utsu-net.jp", "mental-navi.net", "medicalnote.jp", "scuel.me", "byoin.ne.jp",
    "iryou-navi.com", "jimoken.jp", "ritoful.com", "byoin-machi.net",
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
    "mhlw.go.jp", "kaigokensaku.mhlw.go.jp", "wam.go.jp",
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
        # RobotFileParser.read() は Python 既定の User-Agent で取得するため、Cloudflare 等に
        # 403 で弾かれて「全ページ禁止」と誤判定される。通常の UA で取得し、RFC 9309 に従い
        # 4xx・取得失敗は制限なし、正常取得できた場合のみその内容に従う。
        rp = RobotFileParser()
        try:
            r = requests.get(f"{origin}/robots.txt", headers=HEADERS, timeout=10)
            body = r.text if r.status_code == 200 else ""
            if "<html" in body[:500].lower():
                body = ""  # ボット対策ページ等が返された場合
            rp.parse(body.splitlines())
        except Exception:
            rp.parse([])
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
    prev = None
    while prev != s:
        prev, s = s, CORP_PREFIX_RE.sub("", s)
    return _normalize_name(s)


def _normalize_phone(s: str) -> str:
    return re.sub(r"[^0-9]", "", _zen_to_han(s))


def _format_phone(s: str) -> str:
    phone = re.sub(r"[^0-9\-]", "", _zen_to_han(s))
    if not re.search(r"\d", phone):
        return ""
    # 医療情報ネットには「087874-0003」「0878653387」のような区切りの乱れがあるため、
    # 香川県の市外局番（0875/0877/0879 は4桁、087 は3桁）で区切り直す
    digits = re.sub(r"\D", "", phone)
    if len(digits) == 10 and not re.fullmatch(r"0\d{1,4}-\d{1,4}-\d{4}", phone):
        if digits[:4] in ("0875", "0877", "0879"):
            return f"{digits[:4]}-{digits[4:6]}-{digits[6:]}"
        if digits[:3] == "087":
            return f"{digits[:3]}-{digits[3:6]}-{digits[6:]}"
    return phone


def _address_key(s: str) -> str:
    """名寄せ用の住所キー（市区町村以降・番地まで、表記ゆれを吸収）"""
    s = _zen_to_han(s)
    s = re.sub(r"\s+", "", s)
    s = re.sub(r"^〒?\d{3}-?\d{4}", "", s)
    s = s.replace(PREF_NAME, "")
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
    rec["精神系診療科目"] = ""
    rec["名称"] = re.sub(r"[\s　]+", " ", _zen_to_han(name)).strip()
    rec["所在地"] = address
    rec["電話番号"] = _format_phone(phone)
    rec["公式サイトURL"] = url
    return rec


# ── Phase 1: 医療情報ネット（ナビイ） ────────────────────

def _mhlw_url(page: int, lo: str = "", pref: str = "") -> str:
    params = {
        "sc": ",".join(MHLW_DEPT_CODES),
        "st": "01-2",       # 診療科目「いずれかに該当する」（01-1 は「全てに該当する」）
        "sjk": "3",
        "jc": "MC-01",
        "kkn": "1,2",       # 病院・診療所
        "page": str(page),  # 0 始まり
        "size": "20",
        "sortNo": "2",
    }
    if lo:
        params["lo"] = "1" + lo  # 先頭1桁「選択地点の市区町村内で検索する」+ 市区町村コード
    if pref:
        params["pref"] = pref
    return f"{MHLW_SEARCH}?{urlencode(params, safe=',')}"


def _parse_mhlw_item(item) -> dict | None:
    name_a = item.select_one("h2.name a")
    if not name_a:
        return None
    fields: dict[str, object] = {}
    for row in item.select("div.dl-row"):
        img = row.select_one("dt img")
        dd = row.find("dd")
        if not img or dd is None:
            continue
        label = img.get("alt", "")
        if label == "住所":
            p = dd.find("p")
            text = p.get_text(" ", strip=True) if p else dd.get_text(" ", strip=True)
            text = text.replace("Googleマップで見る", "")
            text = re.sub(r"^〒\s*\d{3}-\d{4}\s*", "", text)
            fields["住所"] = re.sub(r"\s+", " ", _zen_to_han(text)).strip()
        elif label == "電話番号":
            tel_a = dd.find("a", href=re.compile(r"^tel:"))
            fields["電話番号"] = tel_a.get_text(strip=True) if tel_a else dd.get_text(" ", strip=True)
        elif label == "URL":
            a = dd.find("a", href=True)
            fields["URL"] = a["href"].strip() if a else ""
        elif label == "診療科目":
            fields["診療科目"] = [td.get_text(strip=True) for td in dd.find_all("td") if td.get_text(strip=True)]

    url = str(fields.get("URL", ""))
    url = re.sub(r"^(https?://)www/", r"\1www.", url)  # 「http://www/tn-clinic.jp」等の誤記
    if url and (not url.startswith("http") or _is_skip_domain(url)):
        url = ""
    rec = _new_record(name_a.get_text(" ", strip=True), str(fields.get("住所", "")),
                      str(fields.get("電話番号", "")), url)
    depts = fields.get("診療科目", []) or []
    rec["精神系診療科目"] = "、".join(d for d in depts if PSY_DEPT_RE.search(d))
    rec["_kikan"] = re.search(r"kikanCd=(\d+)", name_a.get("href", "")).group(1) \
        if "kikanCd=" in name_a.get("href", "") else ""
    return rec


def _scrape_mhlw_query(label: str, **kw) -> tuple[list[dict], int | None]:
    records: list[dict] = []
    total = None
    page = 0
    while True:
        text = _fetch_text(_mhlw_url(page, **kw), timeout=30)
        time.sleep(MHLW_DELAY)
        if text is None:
            logger.warning(f"  {label} page {page} 取得失敗")
            break
        soup = BeautifulSoup(text, "lxml")
        if total is None:
            m = re.search(r"([\d,]+)\s*件中", soup.get_text())
            total = int(m.group(1).replace(",", "")) if m else 0
        items = soup.select("div.resultItems div.item")
        for it in items:
            rec = _parse_mhlw_item(it)
            if rec:
                records.append(rec)
        if not items or len(records) >= (total or 0):
            break
        page += 1
    return records, total


def scrape_mhlw() -> list[dict]:
    logger.info("Phase 1: 医療情報ネット（ナビイ）から市町ごとに収集")
    by_kikan: dict[str, dict] = {}
    for code, city in MUNICIPALITIES.items():
        recs, total = _scrape_mhlw_query(city, lo=code)
        for rec in recs:
            by_kikan.setdefault(rec["_kikan"] or rec["名称"], rec)
        logger.info(f"  {city}: {len(recs)} 件 (検索件数 {total} / 累計 {len(by_kikan)} 件)")

    # 県単位の検索と照合し、市町検索で漏れた施設があれば追加する
    pref_recs, pref_total = _scrape_mhlw_query(PREF_NAME, pref=PREF_CD)
    missing = [r for r in pref_recs if (r["_kikan"] or r["名称"]) not in by_kikan]
    for rec in missing:
        by_kikan[rec["_kikan"] or rec["名称"]] = rec
    logger.info(f"  県単位検索: {pref_total} 件 / 市町検索で漏れていた施設 {len(missing)} 件"
                + (f" ({', '.join(r['名称'] for r in missing)})" if missing else ""))
    outside = [r for r in by_kikan.values() if not r["所在地"].startswith(PREF_NAME)]
    if outside:
        logger.warning(f"  県外の住所: {[r['名称'] for r in outside]}")
    logger.info(f"Phase 1 完了: {len(by_kikan)} 件")
    return list(by_kikan.values())


# ── Phase 1b: 香川県 発達障害医療機関リスト PDF ──────────

def _hattatsu_rows(page) -> list[dict]:
    """番号列（左端）の位置を基準に、前後の番号との中間までを1行の範囲として各列を復元する。"""
    spans = []
    for b in page.get_text("dict")["blocks"]:
        for line in b.get("lines", []):
            for sp in line["spans"]:
                t = sp["text"].strip()
                if t:
                    spans.append(((sp["bbox"][1] + sp["bbox"][3]) / 2, sp["bbox"][0], t))
    spans.sort()
    anchors = [(y, t) for y, x, t in spans if x < 56 and re.fullmatch(r"\d+", t)]
    rows = []
    for i, (y, num) in enumerate(anchors):
        lo = (anchors[i - 1][0] + y) / 2 if i else y - 40
        hi = (anchors[i + 1][0] + y) / 2 if i + 1 < len(anchors) else y + 60

        def cell(x0: float, x1: float) -> list[str]:
            return [t for yy, x, t in spans if lo <= yy < hi and x0 <= x < x1]

        contact = [t for t in cell(625, 800) if t != "ホームページ"]
        phones = [t for t in contact if re.fullmatch(r"[\d\-]{10,13}", t)]
        url = "".join(t for t in contact if t not in phones)  # 折り返されたURLを連結
        rows.append({
            "番号": num,
            "名称": "".join(cell(85, 175)),
            "診療科名": "".join(cell(250, 312)),  # 「児童・思春期精」「神科」等の折り返しを連結
            "所在地": "".join(cell(500, 625)),
            "電話番号": phones[0] if phones else "",
            "URL": url,
        })
    return rows


def scrape_hattatsu_pdf() -> list[dict]:
    logger.info("Phase 1b: 香川県 発達障害医療機関リストPDFから収集")
    try:
        r = requests.get(HATTATSU_PDF_URL, headers=HEADERS, timeout=30)
        r.raise_for_status()
    except Exception as e:
        logger.warning(f"PDF取得失敗: {e}")
        return []

    records, skipped = [], []
    for page in pymupdf.open(stream=r.content, filetype="pdf"):
        for row in _hattatsu_rows(page):
            dept = row["診療科名"]
            if not (HATTATSU_PSY_DEPT_RE.search(dept)
                    or (not dept and HATTATSU_PSY_NAME_RE.search(row["名称"]))):
                skipped.append(f"{row['名称']}({dept or '科名なし'})")
                continue
            url = row["URL"]
            # 「https://www.chp-kagawa-jp」のようなドメイン誤記や紹介サイトURLは採用しない
            if url and (not re.match(r"https?://[^/]+\.[a-z]{2,}(/|$)", url) or _is_skip_domain(url)):
                url = ""
            address = _zen_to_han(row["所在地"]).strip()
            if address and not address.startswith(PREF_NAME):
                address = PREF_NAME + address
            rec = _new_record(row["名称"], address, row["電話番号"], url)
            rec["精神系診療科目"] = dept or "（科名記載なし）"
            records.append(rec)
    logger.info(f"Phase 1b 完了: 精神科・心療内科系 {len(records)} 件を採用 / "
                f"小児科・神経内科等のみで除外 {len(skipped)} 件: {', '.join(skipped)}")
    return records


# ── Phase 2: GMOクリニック・マップ ───────────────────────

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
    logger.info(f"Phase 2: GMOクリニック・マップから{PREF_NAME} 心療内科・精神科を収集")
    by_id: dict[str, dict] = {}
    for treat in GMO_TREATS:
        page = 1
        while True:
            url = f"{GMO_BASE}/clinics/pref{PREF_CD}_treat{treat}" + (f"?page={page}" if page > 1 else "")
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
                rec["_gmo_cats"] = sorted({c["name"] for c in h.get("treatmentCategories", [])})
                by_id[sid] = rec
            logger.info(f"  treat{treat} page {page}: {len(hospitals)} 件 (累計 {len(by_id)} 件)")
            if len(hospitals) < 20:
                break
            page += 1
    logger.info(f"Phase 2 完了: {len(by_id)} 件")
    return list(by_id.values())


def fetch_gmo_phone(sid: str) -> str:
    text = _fetch_text(f"{GMO_BASE}/hospital/ref/{sid}")
    if not text:
        return ""
    m = re.search(r'"telephoneNumber"\s*:\s*"([^"]+)"', text)
    return _format_phone(m.group(1)) if m else ""


def gmo_site_mentions_psychiatry(url: str) -> str:
    """GMOのみ掲載の施設について、公式サイト（トップと診療案内系ページ）に
    精神科・心療内科の記載があるかを確認し、見つかった科目名を返す。"""
    if not url or _is_skip_domain(url) or not _is_allowed(url):
        return ""
    top = _fetch_text(url)
    if not top:
        return ""
    found = set(PSY_SITE_RE.findall(BeautifulSoup(top, "lxml").get_text()))
    if found:
        return "、".join(sorted(found))
    soup = BeautifulSoup(top, "lxml")
    host = urlparse(url).netloc
    kw = re.compile(r"診療科|診療案内|診療内容|外来|department|shinryo|medical|guide")
    links = []
    for a in soup.find_all("a", href=True):
        href = urljoin(url, a["href"])
        if urlparse(href).netloc == host and (kw.search(href) or kw.search(a.get_text())) \
                and href not in links:
            links.append(href)
    for href in links[:10]:
        text = _fetch_text(href) or ""
        found |= set(PSY_SITE_RE.findall(BeautifulSoup(text, "lxml").get_text()))
        time.sleep(0.5)
    return "、".join(sorted(found))


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


# ── Phase 3: DuckDuckGo 公式サイト検索+検証 ──────────────

def search_official_site(name: str, address: str, phone: str) -> tuple[str, str]:
    city_match = re.search(PREF_NAME + r"\s*(?:[^\s]{2,4}郡)?([^\s]{1,8}?[市町村])", address)
    city = city_match.group(1) if city_match else PREF_NAME
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
    """公式サイトURLが判明している施設について追加情報のみ取得する。"""
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
    logger.info(f"{PREF_NAME} 精神科・心療内科 収集開始")
    logger.info("=" * 65)
    start_time = datetime.now()

    mhlw_raw = scrape_mhlw()
    hattatsu_raw = scrape_hattatsu_pdf()
    gmo_raw = scrape_gmo()

    merger = Merger()
    for rec in mhlw_raw:
        merger.add(rec)
    mhlw_count = len(merger.records)
    hattatsu_new = [rec["名称"] for rec in hattatsu_raw if merger.add(rec)]
    logger.info(f"発達障害医療機関リストから新規追加: {len(hattatsu_new)} 件 {hattatsu_new}")

    gmo_new = [rec for rec in gmo_raw if merger.find(rec) is None]
    for rec in gmo_raw:
        if merger.find(rec) is not None:
            merger.add(rec)  # 公式サイトURLの補完のみ

    # GMOのみ掲載の施設: 電話番号で既存施設と同一と判明する場合があるため取得してから名寄せし、
    # 新規施設は公式サイトで精神科・心療内科の記載を確認できたものだけ採用する
    logger.info(f"GMOのみ掲載 {len(gmo_new)} 件を確認（電話番号取得・公式サイトで診療科目確認）")
    adopted, rejected = [], []
    for rec in gmo_new:
        if _is_allowed(f"{GMO_BASE}/hospital/ref/{rec['_gmo_id']}"):
            rec["電話番号"] = fetch_gmo_phone(rec["_gmo_id"])
        time.sleep(GMO_DELAY)
        if merger.find(rec) is not None:
            merger.add(rec)
            continue
        depts = gmo_site_mentions_psychiatry(rec["公式サイトURL"])
        # GMOの「精神科・神経科」は神経内科・総合病院にも付与されるため単独では根拠にしないが、
        # 児童精神科・心療内科等の具体的な分類や、施設名（「メンタル」等）は根拠として扱う
        specific = [c for c in rec["_gmo_cats"] if c in GMO_SPECIFIC_PSY_CATS]
        if depts:
            rec["精神系診療科目"] = f"{depts}（公式サイト記載）"
        elif specific:
            rec["精神系診療科目"] = f"{'、'.join(specific)}（GMOクリニック・マップ）"
        elif HATTATSU_PSY_NAME_RE.search(rec["名称"]):
            rec["精神系診療科目"] = "（施設名より精神科系と判断・科目未確認）"
        if rec["精神系診療科目"]:
            merger.add(rec)
            adopted.append(rec["名称"])
        else:
            rejected.append(rec["名称"])
    logger.info(f"  GMOのみ掲載で採用: {len(adopted)} 件 {adopted}")
    logger.info(f"  GMOのみ掲載で除外（精神科・心療内科を確認できず）: {len(rejected)} 件 {rejected}")
    for rec in gmo_new:
        logger.info(f"    {rec['名称']}: GMO分類 {rec['_gmo_cats']} → {rec['精神系診療科目'] or '除外'}")

    deduped = merger.records
    logger.info(f"名寄せ後: {len(deduped)} 件 (医療情報ネット {mhlw_count} / "
                f"発達障害医療機関リスト {len(hattatsu_new)} / GMO {len(adopted)})")
    logger.info(f"  公式サイトURL判明: {sum(1 for r in deduped if r.get('公式サイトURL'))} 件")

    logger.info("Phase 4: 既知の公式サイトから追加情報取得")
    await enrich_known_urls(deduped, sem_count=8)

    logger.info("Phase 3+4: 未判明の施設を DuckDuckGo で検索・検証・情報取得")
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
