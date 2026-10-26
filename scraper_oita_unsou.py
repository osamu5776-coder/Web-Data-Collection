"""
大分県 運送会社リスト作成

データソース:
  Phase 1 - 大分県トラック運送業界紹介サイト（truck-info.ota.or.jp）
            「所在地: 大分市」の会員事業者（WordPress REST API で全件取得。
            一覧ページのページ送りはフィルタ付きだと常に1ページ目を返すため使わない）
  Phase 2 - Google マップ「運送会社 大分県」の検索結果（Playwright で一覧を
            最後までスクロールし、各スポットの住所・電話・ウェブサイトを取得）

  Phase 3 - 名寄せ: 法人格・支店/営業所/センター等の表記を除いた社名（コア名）で統合
  Phase 4 - 本社判定: 国税庁法人番号公表サイトでコア名を検索し、
              大分県内に同名法人あり              → 本社 大分県（登記上の本店所在地を採用）
              県内になし・県外に同名法人あり      → 除外（他県本社の支店・営業所）
              法人登録なし（個人事業等）          → 支店表記がなく所在地が県内なら採用
            支店表記の事業所でも県内に同名法人があれば、県内本社の営業所として採用し
            所在地は登記上の本店所在地とする
  Phase 5 - 公式サイトURLがない会社は DuckDuckGo で検索し、社名と市町村がページ内に
            あることを確認して採用
  Phase 6 - 公式サイト（トップ＋会社概要・問い合わせページ）からメールアドレス・
            インスタURL・問い合わせフォームURL・電話番号（未取得時）を取得

出力列: 名称, メールアドレス, 公式サイトURL, 所在地, 電話番号, インスタURL, 問い合わせフォームURL
出力ファイル: 大分県_運送会社リスト.xlsx
"""

import asyncio
import io
import json
import logging
import os
import random
import re
import sys
import time
import unicodedata
import warnings
from datetime import datetime
from urllib.parse import urljoin, urlparse
from urllib.robotparser import RobotFileParser

import pandas as pd
import phonenumbers
import requests
from bs4 import BeautifulSoup
from bs4 import XMLParsedAsHTMLWarning
from ddgs import DDGS
from phonenumbers import PhoneNumberFormat

# ── ロギング ──────────────────────────────────────────────
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

LOG_FILE = "scraper_oita_unsou.log"
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
warnings.filterwarnings("ignore", category=XMLParsedAsHTMLWarning)
for _name in ("httpx", "primp", "ddgs", "urllib3"):
    logging.getLogger(_name).setLevel(logging.WARNING)

# ── 定数 ─────────────────────────────────────────────────
OUTPUT_FILE = "大分県_運送会社リスト.xlsx"
CACHE_DIR = os.path.join("data", "oita_unsou")
UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36"
)
OTA_API = "https://truck-info.ota.or.jp/wp-json/wp/v2/organization"
OTA_LOCATION_OITA_CITY = 2
GMAPS_URL = (
    "https://www.google.com/maps/search/%E9%81%8B%E9%80%81%E4%BC%9A%E7%A4%BE+%E5%A4%A7%E5%88%86%E7%9C%8C/"
    "@33.5264702,129.8203822,9z?hl=ja"
)
NTA_BASE = "https://www.houjin-bangou.nta.go.jp"
NTA_PREF_OITA = "44"

OUTPUT_COLS = [
    "名称", "メールアドレス", "公式サイトURL",
    "所在地", "電話番号", "インスタURL", "問い合わせフォームURL",
]

def _cache(name: str):
    path = os.path.join(CACHE_DIR, name)
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    return None


def _save_cache(name: str, data) -> None:
    os.makedirs(CACHE_DIR, exist_ok=True)
    with open(os.path.join(CACHE_DIR, name), "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)


def _nfkc(s: str) -> str:
    return unicodedata.normalize("NFKC", str(s or "")).strip()


def _format_phone(s: str) -> str:
    """日本の市外局番の境界に従ってハイフンを入れる。"""
    s = _nfkc(s)
    m = re.search(r"0\d[\d\-‐ー−–()（） ]{7,14}\d", s)
    if not m:
        return ""
    try:
        num = phonenumbers.parse(re.sub(r"\D", "", m.group(0)), "JP")
        return phonenumbers.format_number(num, PhoneNumberFormat.NATIONAL)
    except phonenumbers.NumberParseException:
        return m.group(0)


def _norm_addr(s: str) -> str:
    """所在地を「大分県…」形式に整える（郵便番号・「住所:」を除去）。"""
    a = _nfkc(s)
    a = re.sub(r"^住所:\s*", "", a)
    a = re.sub(r"〒?\s*\d{3}-?\d{4}\s*", "", a).strip()
    if not a.startswith("大分県") and re.match(r"^(大分市|別府市|中津市|日田市|佐伯市|臼杵市|津久見市|竹田市|豊後高田市|杵築市|宇佐市|豊後大野市|由布市|国東市|東国東郡|速見郡|玖珠郡)", a):
        a = "大分県" + a
    return re.sub(r"\s+", " ", a)


# ── 社名の正規化（コア名） ─────────────────────────────────

_LEGAL_RE = re.compile(
    r"株式会社|有限会社|合同会社|合資会社|合名会社|一般社団法人|"
    r"\((?:株|有|同|資|名)\)|㈱|㈲"
)
# 「引越センター」「○○倉庫」「○○商店」は社名の一部なので拠点表記として扱わない
_BRANCH_KW = (
    r"営業所|支店|主管支店|支社|本社|本店|(?<!引越)センター|"
    r"事業所|事務所|出張所|荷扱所|営業部|運営部|分室|(?<![\s商])店"
)
_BRANCH_RE = re.compile(rf"({_BRANCH_KW})")
# 拠点名に使われる地名・区分（「大分豊海」「大分西」「森町」等）
_PLACES = (
    "西大分|大分|別府|中津|日田|佐伯|臼杵|津久見|竹田|豊後高田|杵築|宇佐|豊後大野|"
    "由布|国東|日出|九重|玖珠|大在|鶴崎|三佐|豊海|森町|畑中|みなと|空港|港|生協|デポ|"
    "チルド|航空貨物|福岡|北九州|熊本|宮崎|鹿児島|長崎|佐賀|[東西南北]"
)
_SUFFIX_RE = re.compile(rf"^(.+?)(?:{_PLACES})*(?:本社)?(?:{_BRANCH_KW})$")


def core_name(name: str) -> str:
    """法人格・括弧書き・支店/営業所表記を除いた社名。"""
    s = _nfkc(name)
    s = re.sub(r"[【\[(（](?!(?:株|有|同|資|名)\))[^】\])）]*[】\])）]", " ", s)
    segs = [x.strip() for x in _LEGAL_RE.split(s)]
    segs = [x for x in segs if x]
    if not segs:
        return ""
    # 法人格の後ろに続く「大分支店」等の区分は捨てる
    keep = [x for x in segs if not (_BRANCH_RE.search(x) and len(segs) > 1 and x is not segs[0])] or segs[:1]
    core = keep[0]
    parts = core.split(" ")
    if len(parts) > 1 and _BRANCH_RE.search(parts[-1]):
        core = " ".join(parts[:-1])
    core = core.replace(" ", "")
    # 「九州支店大分営業所」のような多段の拠点表記は繰り返し除去する
    while (m := _SUFFIX_RE.match(core)) and len(m.group(1)) >= 2:
        core = m.group(1)
    return core


def is_branch(name: str) -> bool:
    """名称が支店・営業所等の拠点を表しているか（「本社」「本店」表記は除く）。"""
    s = _nfkc(name)
    s = re.sub(r"本社|本店", "", s)
    return bool(re.search(r"営業所|支店|支社|(?<!引越)センター|事業所|出張所|荷扱所|(?<![\s商])店(?=$|[\s】)])", s))


# ── Phase 1: 大分県トラック協会 会員（大分市） ─────────────────

def fetch_ota() -> list[dict]:
    cached = _cache("ota.json")
    if cached is not None:
        logger.info(f"Phase 1: キャッシュから {len(cached)} 件")
        return cached
    s = requests.Session()
    s.headers["User-Agent"] = UA
    out = []
    page = 1
    while True:
        r = s.get(OTA_API, params={"per_page": 100, "page": page, "location": OTA_LOCATION_OITA_CITY}, timeout=30)
        if r.status_code != 200 or not r.json():
            break
        for x in r.json():
            soup = BeautifulSoup(x["content"]["rendered"], "lxml")
            a = soup.select_one(".OrganizationCard__title a")
            sel = lambda c: soup.select_one(c).get_text(strip=True) if soup.select_one(c) else ""
            out.append({
                "src": "協会",
                "name": sel(".OrganizationCard__title"),
                "addr": _norm_addr(sel(".OrganizationCard__addressText")),
                "tel": "",
                "web": a["href"] if a else "",
            })
        if page >= int(r.headers.get("X-WP-TotalPages", 1)):
            break
        page += 1
        time.sleep(1)
    _save_cache("ota.json", out)
    logger.info(f"Phase 1: 協会 大分市 {len(out)} 件")
    return out


# ── Phase 2: Google マップ ───────────────────────────────

async def fetch_gmaps() -> list[dict]:
    cached = _cache("gmaps.json")
    if cached is not None:
        logger.info(f"Phase 2: キャッシュから {len(cached)} 件")
        return cached
    from playwright.async_api import async_playwright

    out = []
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        ctx = await browser.new_context(locale="ja-JP", user_agent=UA)
        pg = await ctx.new_page()
        await pg.goto(GMAPS_URL, wait_until="domcontentloaded")
        await pg.wait_for_timeout(5000)
        feed = pg.locator('div[role="feed"]')
        last, same = 0, 0
        for _ in range(80):
            await feed.evaluate("e => e.scrollBy(0, 5000)")
            await pg.wait_for_timeout(1500)
            if await pg.locator("text=リストの最後に到達しました").count():
                break
            n = await pg.locator("a.hfpxzc").count()
            same = same + 1 if n == last else 0
            last = n
            if same >= 6:
                break
        links = await pg.eval_on_selector_all(
            "a.hfpxzc", "as => as.map(a => [a.getAttribute('aria-label'), a.href])"
        )
        logger.info(f"Phase 2: Google マップ {len(links)} 件の詳細を取得")
        for name, href in links:
            await pg.goto(href + "&hl=ja", wait_until="domcontentloaded")
            try:
                await pg.wait_for_selector("h1", timeout=10000)
            except Exception:
                pass
            await pg.wait_for_timeout(1500)
            rec = {"src": "Googleマップ", "name": name, "addr": "", "tel": "", "web": ""}
            addr = pg.locator('button[data-item-id="address"]').first
            if await addr.count():
                rec["addr"] = _norm_addr(await addr.get_attribute("aria-label") or "")
            tel = pg.locator('button[data-item-id^="phone:tel:"]').first
            if await tel.count():
                rec["tel"] = _format_phone(await tel.get_attribute("aria-label") or "")
            web = pg.locator('a[data-item-id="authority"]').first
            if await web.count():
                rec["web"] = await web.get_attribute("href") or ""
            out.append(rec)
        await browser.close()
    _save_cache("gmaps.json", out)
    return out


# ── Phase 3: 名寄せ ──────────────────────────────────────

def group_companies(records: list[dict]) -> dict[str, list[dict]]:
    groups: dict[str, list[dict]] = {}
    for r in records:
        if not r["addr"].startswith("大分県"):
            continue  # 県外の拠点（Googleマップの検索範囲が福岡県を含むため）
        c = core_name(r["name"])
        if c:
            groups.setdefault(c, []).append(r)
    return groups


# ── Phase 4: 国税庁法人番号公表サイトで本店所在地を確認 ──────────

class NTA:
    def __init__(self) -> None:
        self.s = requests.Session()
        self.s.headers["User-Agent"] = UA

    def search(self, name: str, pref: str = "") -> list[dict]:
        top = BeautifulSoup(self.s.get(NTA_BASE + "/", timeout=30).text, "lxml")
        tok = top.find("input", attrs={"name": re.compile("token")})
        data = {
            tok["name"]: tok["value"], "search": "true", "japanese": "true",
            "houzinNmShTypeRbtn": "2", "houzinNmTxtf": name,
            "_kanaCkbx": "on", "_noconvCkbx": "on", "_enCkbx": "on",
            "houzinAddrShTypeRbtn": "1", "prefectureLst": pref, "cityLst": "",
            "houzinNoShTyoumeSts": "0", "tyoumeTxtf": "", "kokugaiTxtf": "", "zipCdTxtf": "",
            "houzinNoShSonotaZyoukenSts": "0", "_houzinKdCkbx": "on", "_historyCkbx": "on",
            "_hideCkbx": "on", "_closeCkbx": "on", "_chgYmdShTargetCkbx": "on",
            "orderRbtn": "1", "houzinKdRbtn": "0", "searchFlg": "1",
            "preSyousaiScreenId": "KJSCR0101010",
        }
        r = self.s.post(NTA_BASE + "/kensaku-kekka.html", data=data, timeout=30)
        soup = BeautifulSoup(r.text, "lxml")
        out = []
        for tr in soup.select("table tr"):
            th = tr.find("th")
            tds = tr.find_all("td")
            if not th or len(tds) < 2:
                continue
            nm = tds[0].get_text("\n", strip=True).split("\n")[-1]
            out.append({"no": th.get_text(strip=True), "name": nm, "addr": tds[1].get_text(strip=True)})
        return out


def _same_core(nta_name: str, core: str) -> bool:
    return _nfkc(_LEGAL_RE.sub("", _nfkc(nta_name))).replace(" ", "") == core


def check_hq(groups: dict[str, list[dict]]) -> dict[str, dict]:
    cache = _cache("nta.json") or {}
    nta = NTA()
    todo = [c for c in groups if c not in cache]
    logger.info(f"Phase 4: 法人番号公表サイトで本店所在地確認 {len(todo)} 件（キャッシュ {len(cache)} 件）")
    for i, core in enumerate(todo, 1):
        try:
            oita = [x for x in nta.search(core, NTA_PREF_OITA) if _same_core(x["name"], core)]
            time.sleep(1)
            other = []
            if not oita:
                other = [x for x in nta.search(core) if _same_core(x["name"], core)]
                time.sleep(1)
            cache[core] = {"oita": oita, "other": other}
        except Exception as e:
            logger.warning(f"  法人番号検索失敗 {core}: {e}")
            continue
        if i % 20 == 0:
            _save_cache("nta.json", cache)
            logger.info(f"  {i}/{len(todo)}")
    _save_cache("nta.json", cache)
    return cache


def decide(core: str, recs: list[dict], nta: dict) -> tuple[bool, str]:
    """(大分県本社か, 判定理由)"""
    res = nta.get(core, {"oita": [], "other": []})
    all_branch = all(is_branch(r["name"]) for r in recs)
    if res["oita"]:
        return True, "法人番号: 県内本店"
    if res["other"]:
        return False, "法人番号: 県外本店"
    if all_branch:
        return False, "支店表記・法人登録なし"
    return True, "法人登録なし（個人事業等）"


# ── Phase 5: 公式サイト検索 ────────────────────────────────

SKIP_DOMAINS = [
    "google.", "yahoo.", "bing.com", "facebook.com", "instagram.com", "twitter.com", "x.com",
    "youtube.com", "wikipedia", "jpnumber", "telnavi", "houjin", "baseconnect", "mapion",
    "navitime", "itp.ne.jp", "ekiten", "indeed", "townwork", "mynavi", "hellowork", "baitoru",
    "en-gage", "doraever", "driver-job", "dora-ever", "careerindex", "jobmedley", "job-medley",
    "kyujin", "recruit", "rikunabi", "doda", "stanby", "ota.or.jp", "jta.or.jp", "tdb.co.jp",
    "salesnow", "musubu", "alarmbox", "info-houjin", "nta.go.jp", "gbiz", "catr.jp", "graffer",
    "trucks-info", "unsou-navi", "hikkoshi", "suumo", "homes.co.jp", "biz-maps", "jbplt",
    "ameblo", "note.com", "linkedin", "zehitomo", "tsunagu", "cnavi",
]
_SKIP_HOST_RE = re.compile(r"job|kyujin|recruit|career|navi|search|portal|ranking|matome|blog|houjin|kigyou|company-?list")


def _skip(url: str) -> bool:
    u = url.lower()
    return any(d in u for d in SKIP_DOMAINS) or bool(_SKIP_HOST_RE.search(urlparse(u).netloc))


_robots_cache: dict[str, RobotFileParser] = {}


def _is_allowed(url: str) -> bool:
    parsed = urlparse(url)
    origin = f"{parsed.scheme}://{parsed.netloc}"
    if origin not in _robots_cache:
        rp = RobotFileParser()
        try:
            r = requests.get(f"{origin}/robots.txt", headers={"User-Agent": UA}, timeout=8)
            rp.parse(r.text.splitlines() if r.status_code == 200 else [])
        except Exception:
            rp.parse([])
        _robots_cache[origin] = rp
    return _robots_cache[origin].can_fetch(UA, url)


def _get(url: str) -> tuple[str, str]:
    try:
        r = requests.get(url, headers={"User-Agent": UA}, timeout=12, allow_redirects=True)
        if r.status_code >= 400:
            return "", ""
        if r.encoding is None or r.encoding.lower() in ("iso-8859-1", "ascii"):
            r.encoding = r.apparent_encoding or "utf-8"
        return r.url, r.text
    except Exception:
        return "", ""


def _city(addr: str) -> str:
    m = re.match(r"大分県(.+?[市郡])", _nfkc(addr))
    return m.group(1) if m else ""


def _town(addr: str) -> str:
    """所在地の市町村より後ろの町名・大字名（番地の手前まで）。"""
    a = _nfkc(addr)
    a = a[a.find(_city(addr)) + len(_city(addr)):] if _city(addr) else a
    a = re.sub(r"^(.+?郡)?.+?[町村](?=大字|字)", "", a) if "郡" in _city(addr) else a
    a = re.sub(r"^大字|^字", "", a)
    return re.split(r"[\d一二三四五六七八九十〇]|字", a)[0]


# 企業情報のまとめ・検索サイト（社名と住所は載るが公式サイトではない）
DIRECTORY_DOMAINS = [
    "hakopro.jp", "data-link-plus.com", "kensetumap.com", "data-max.co.jp", "suke-dachi.jp",
    "jfsc.jp", "ivry.jp", "everytown.info", "x-work.jp", "cnavi", "nikkei", "buffett-code",
    "corporate-information", "kaisharesearch", "kigyou", "jigyosyo", "townpage", "goo.ne.jp",
    "compalyze", "tsujigawa", "apese.net", "town-guide", "hojin", "corp-", "bizdb", "clipbiz",
]
_DIRECTORY_TITLE_RE = re.compile(r"会社情報|企業情報|法人情報|口コミ|評判|求人|一覧|検索|電話番号|地図|アクセス情報")


def _oita_phones(text: str) -> list[tuple[int, str]]:
    """ページ内の大分県（097x）・フリーダイヤルの電話番号を (位置, 整形済み番号) で返す。"""
    out = []
    for m in re.finditer(r"(?<!\d)(0\d{1,4})[-‐ー−()（）\s]{0,2}(\d{1,4})[-‐ー−()（）\s]{0,2}(\d{3,4})(?!\d)", text):
        digits = "".join(m.groups())
        if re.search(r"FAX|Fax|fax|ファクス|ファックス", text[max(0, m.start() - 10):m.start()]):
            continue
        if len(digits) == 10 and digits.startswith(("097", "0120", "0800")):
            out.append((m.start(), _format_phone(digits)))
    return out


def search_official(core: str, addr: str) -> tuple[str, str]:
    """(公式サイトURL, まとめサイト等から得た電話番号) を返す。"""
    city, town = _city(addr), _town(addr)
    query = f"{core} {city or '大分県'} {town}"
    try:
        with DDGS() as ddgs:
            results = list(ddgs.text(query, region="jp-jp", max_results=8))
    except Exception as e:
        logger.debug(f"検索失敗 {core}: {e}")
        return "", ""
    tel = ""
    for res in results:
        url = res.get("href", "")
        if not url.startswith("http") or _skip(url) or not _is_allowed(url):
            continue
        final, html = _get(url)
        if not html or _skip(final):
            continue
        soup = BeautifulSoup(html, "lxml")
        page = _nfkc(re.sub(r"\s", "", soup.get_text()))
        if core not in page or (city and city not in page) or (len(town) >= 2 and town not in page):
            continue
        title = _nfkc(soup.title.get_text() if soup.title else "").replace(" ", "")
        directory = (
            any(d in final.lower() for d in DIRECTORY_DOMAINS)
            or bool(_DIRECTORY_TITLE_RE.search(title))
            or bool(re.search(r"\d{13}|[?&]id=", final))  # 法人番号・掲載IDで個社ページを表すまとめサイト
        )
        if not directory and core in title:
            return final, tel
        if not tel and len(town) >= 2:
            # 町名の記載から 300 文字以内にある番号のうち最も近いもの
            anchors = [m.start() for m in re.finditer(re.escape(town), page)]
            best = min(((abs(p - a), t) for p, t in _oita_phones(page) for a in anchors), default=None)
            if best and best[0] <= 300:
                tel = best[1]
    return "", tel


# ── Phase 6: 公式サイトからメール・インスタ・問い合わせフォーム・電話番号 ──

def _decode_cfemail(encoded: str) -> str:
    try:
        key = int(encoded[:2], 16)
        return bytes(int(encoded[i:i + 2], 16) ^ key for i in range(2, len(encoded), 2)).decode("utf-8")
    except Exception:
        return ""


_EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")
_CONTACT_KWS = ["contact", "inquiry", "toiawase", "otoiawase", "お問い合わせ", "お問合せ", "問い合わせ", "問合せ", "メールフォーム"]
_PROFILE_KWS = ["company", "about", "profile", "outline", "gaiyou", "gaiyo", "corporate", "会社概要", "会社案内", "企業情報", "会社情報"]


def _parse_site(url: str, html: str, result: dict, core: str = "") -> list[str]:
    """ページから情報を取り出して result を埋め、辿るべき同一ドメインのリンクを返す。"""
    soup = BeautifulSoup(html, "lxml")
    host = urlparse(url).netloc
    text = soup.get_text("\n")

    if not result.get("メールアドレス"):
        for el in soup.find_all(attrs={"data-cfemail": True}):
            if (m := _decode_cfemail(el["data-cfemail"])):
                result["メールアドレス"] = m
                break
    if not result.get("メールアドレス"):
        mails = [a["href"][7:].split("?")[0] for a in soup.find_all("a", href=True) if a["href"].lower().startswith("mailto:")]
        mails += _EMAIL_RE.findall(text)
        mails = [m for m in mails if _EMAIL_RE.fullmatch(m) and not re.search(r"\.(png|jpe?g|gif|svg|webp)$|example|sentry|wixpress", m, re.I)]
        if mails:
            result["メールアドレス"] = mails[0]

    if not result.get("_tel"):
        # グループ会社のサイトでは複数社の番号が並ぶため、社名の記載に最も近い番号を優先する
        page = _nfkc(text)
        phones = _oita_phones(page)
        anchors = [m.start() for m in re.finditer(re.escape(core), page)] if core else []
        near = min(((abs(p - a), t) for p, t in phones for a in anchors), default=None)
        labeled = [t for p, t in phones if re.search(r"TEL|Tel|tel|電話|☎|℡", page[max(0, p - 8):p])]
        if near and near[0] <= 300:
            result["_tel"] = near[1]
        elif labeled:
            result["_tel"] = labeled[0]
        elif phones:
            result["_tel"] = phones[0][1]

    nxt = []
    for a in soup.find_all("a", href=True):
        href = a["href"].strip()
        label = a.get_text(strip=True)
        if href.startswith("//"):
            href = "https:" + href
        if "instagram.com" in href and not result.get("インスタURL") and not re.search(r"/(p|reel|explore)/", href):
            result["インスタURL"] = href.split("?")[0].rstrip("/")
            continue
        if href.startswith(("#", "mailto:", "tel:", "javascript:")):
            continue
        absu = urljoin(url, href)
        if urlparse(absu).netloc != host:
            continue
        low = absu.lower()
        if any(k in low or k in label for k in _CONTACT_KWS):
            if not result.get("問い合わせフォームURL") and not re.search(r"\.(pdf|jpe?g|png)$", low):
                result["問い合わせフォームURL"] = absu
            nxt.append(absu)
        elif any(k in low or k in label for k in _PROFILE_KWS):
            nxt.append(absu)
    return nxt


def enrich_site(url: str, core: str = "") -> dict:
    result: dict = {}
    if not url.startswith("http") or not _is_allowed(url):
        return result
    final, html = _get(url)
    if not html:
        return result
    queue = _parse_site(final, html, result, core)
    seen = {final}
    for link in queue[:4]:
        if link in seen or not _is_allowed(link):
            continue
        seen.add(link)
        f2, h2 = _get(link)
        if h2:
            _parse_site(f2, h2, result, core)
        if result.get("メールアドレス") and result.get("_tel") and result.get("問い合わせフォームURL"):
            break
    return result


async def _enrich_sites(targets: list[dict]) -> None:
    sem = asyncio.Semaphore(8)

    async def _one(c: dict) -> None:
        async with sem:
            extra = await asyncio.to_thread(enrich_site, c["公式サイトURL"], c["_core"])
            for k in ("メールアドレス", "インスタURL", "問い合わせフォームURL"):
                if extra.get(k) and not c.get(k):
                    c[k] = extra[k]
            if not c["電話番号"] and extra.get("_tel"):
                c["電話番号"] = extra["_tel"]

    await asyncio.gather(*[_one(c) for c in targets])


async def enrich_all(companies: list[dict]) -> None:
    with_url = [c for c in companies if c["公式サイトURL"]]
    logger.info(f"Phase 6: 公式サイトから連絡先取得 {len(with_url)} 件")
    await _enrich_sites(with_url)

    targets = [c for c in companies if not c["公式サイトURL"] or not c["電話番号"]]
    logger.info(f"Phase 5: 公式サイトまたは電話番号が未取得の {len(targets)} 件を検索")
    sem = asyncio.Semaphore(3)

    async def _search(c: dict) -> None:
        async with sem:
            url, tel = await asyncio.to_thread(search_official, c["_core"], c["所在地"])
            if url and not c["公式サイトURL"]:
                c["公式サイトURL"] = url
                c["_new_site"] = True
            if tel and not c["電話番号"]:
                c["電話番号"] = tel
            await asyncio.sleep(random.uniform(1.0, 2.0))

    await asyncio.gather(*[_search(c) for c in targets])
    found = [c for c in targets if c.get("_new_site")]
    logger.info(f"  公式サイト発見: {len(found)} 件 / 電話番号あり: {sum(1 for c in targets if c['電話番号'])} 件")
    await _enrich_sites(found)

    for c in companies:
        if not c["電話番号"]:
            c["電話番号"] = c["_branch_tel"]


# ── メイン ────────────────────────────────────────────────

def _pick(recs: list[dict], key: str) -> str:
    """本社・非支店のレコードを優先して値を選ぶ。"""
    ordered = sorted(recs, key=lambda r: (is_branch(r["name"]), "本社" not in r["name"], r["src"] != "Googleマップ"))
    for r in ordered:
        if r.get(key):
            return r[key]
    return ""


def _clean_url(url: str) -> str:
    """広告計測用のクエリ（utm_*, yclid 等）を除去する。"""
    if re.search(r"[?&](utm_|yclid|gclid|fbclid)", url):
        return url.split("?")[0]
    return url


def _display_name(core: str, recs: list[dict], nta_hit: list[dict]) -> str:
    if nta_hit:
        return _nfkc(nta_hit[0]["name"]).replace("(株)", "株式会社").replace("(有)", "有限会社")
    nm = _nfkc(_pick(recs, "name"))
    nm = nm.replace("(株)", "株式会社").replace("(有)", "有限会社").replace("(同)", "合同会社")
    if is_branch(nm) or "本社" in nm:
        return core
    return re.sub(r"[【(][^】)]*[】)]", "", nm).strip()


async def main() -> None:
    logger.info("=" * 65)
    logger.info("大分県 運送会社リスト作成開始")
    logger.info("=" * 65)
    start = datetime.now()

    records = fetch_ota() + await fetch_gmaps()
    groups = group_companies(records)
    logger.info(f"Phase 3: 名寄せ後 {len(groups)} 社")

    nta = check_hq(groups)
    companies = []
    excluded = []
    for core, recs in groups.items():
        ok, reason = decide(core, recs, nta)
        if not ok:
            excluded.append((core, reason, recs[0]["name"]))
            continue
        hit = nta.get(core, {}).get("oita", [])
        non_branch_addr = [r["addr"] for r in recs if not is_branch(r["name"])]
        addr = non_branch_addr[0] if non_branch_addr else (_norm_addr(hit[0]["addr"]) if hit else _pick(recs, "addr"))
        companies.append({
            "_core": core,
            "_reason": reason,
            "名称": _display_name(core, recs, hit),
            "メールアドレス": "",
            "公式サイトURL": _clean_url(_pick(recs, "web")),
            "所在地": addr,
            # 本社の記載がない（県内本社の営業所のみ掲載）場合、所在地は登記上の本店なので
            # 営業所の電話番号は公式サイト等で本社番号が見つからなかったときの代替とする
            "電話番号": _format_phone(_pick(recs, "tel")) if non_branch_addr else "",
            "_branch_tel": "" if non_branch_addr else _format_phone(_pick(recs, "tel")),
            "インスタURL": "",
            "問い合わせフォームURL": "",
        })
    for core, reason, name in sorted(excluded):
        logger.info(f"  除外: {name}（{reason}）")
    logger.info(f"Phase 4: 大分県本社 {len(companies)} 社 / 除外 {len(excluded)} 社")

    await enrich_all(companies)

    df = pd.DataFrame(companies)
    df = df.sort_values(["所在地", "名称"]).reset_index(drop=True)
    _save_cache("companies.json", df.to_dict("records"))
    df = df[OUTPUT_COLS]
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="運送会社一覧")

    elapsed = int((datetime.now() - start).total_seconds())
    logger.info("=" * 65)
    logger.info(f"完了: {OUTPUT_FILE} に {len(df)} 件を出力")
    for col in OUTPUT_COLS[1:]:
        logger.info(f"  {col}: {(df[col] != '').sum()} 件")
    logger.info(f"所要時間: {elapsed // 60}分{elapsed % 60}秒")
    logger.info("=" * 65)


if __name__ == "__main__":
    asyncio.run(main())
