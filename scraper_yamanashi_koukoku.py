"""
山梨県 広告代理店リスト作成（本社が山梨県の会社のみ）

データソース（4 ソースを並列に収集する）:
  Phase 1 - 広告代理店の一覧（名簿） https://広告代理店の一覧.com/list/search.php?id=19 の一覧と各詳細ページ
            （会社名・住所・電話番号・ホームページ）。robots.txt なし。
            AgencyHub https://agencyhub.jp/prefecture/yamanashi/ の一覧と各詳細ページ
            （所在地・電話番号・公式サイト・問い合わせページ）
            NAVITIME 山梨県「広告代理店」タグ一覧（全ページ）と各スポットの詳細ページ
  Phase 2 - Google マップ「広告代理店 山梨県」と主要市町村別の検索結果（Playwright で一覧を
            最後までスクロールし、各スポットの住所・電話・ウェブサイト・カテゴリを取得）

  Phase 3 - 名寄せ: 法人格・支店/支局/営業所等の表記を除いた社名（コア名）で統合
  Phase 4 - 本社判定: 国税庁法人番号公表サイトでコア名を検索し、
              山梨県内に同名法人あり →  掲載拠点と同じ市町村に本店がある（同名法人が1社だけなら
                                        市町村が違っても可）とき本社 山梨県として採用
              県内になし・県外に同名法人あり → 除外（他県本社の支店・支局・営業所）
              法人登録なし（個人事業等）    → 支店/営業所表記がなければ採用
  Phase 7 - Google マップにだけ載っている会社は、カテゴリが広告・販促・制作系のものに限る
  Phase 6 - 公式サイト（トップ＋会社概要・問い合わせページ）からメールアドレス・
            インスタURL・問い合わせフォームURL・電話番号（未取得時）を取得
  Phase 5 - 公式サイトまたは電話番号が未取得の会社は DuckDuckGo で検索し、社名と
            市町村・町名がページ内にあることを確認して採用

電話番号は phonenumbers で日本の市外局番の境界に従ってハイフンを入れる。

出力列: 名称, メールアドレス, 公式サイトURL, 所在地, 電話番号, インスタURL, 問い合わせフォームURL
出力ファイル: 山梨県_広告代理店リスト.xlsx
"""

import asyncio
import io
import json
import logging
import os
import random
import re
import sys
import threading
import time
import unicodedata
import warnings
from concurrent.futures import ThreadPoolExecutor
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

LOG_FILE = "scraper_yamanashi_koukoku.log"
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
OUTPUT_FILE = "山梨県_広告代理店リスト.xlsx"
CACHE_DIR = os.path.join("data", "yamanashi_koukoku")
UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36"
)
MEIBO_BASE = "https://xn--bdk1d3bw36u9uu.com/list/"  # 広告代理店の一覧.com
MEIBO_LIST = MEIBO_BASE + "search.php?id=19"
AGENCYHUB_LIST = "https://agencyhub.jp/prefecture/yamanashi/"
NAVITIME_BASE = "https://www.navitime.co.jp/category/0516/19/"
NAVITIME_TAG = "010429"  # 広告代理店
YAMANASHI_CITIES = (
    "甲府市", "富士吉田市", "都留市", "山梨市", "大月市", "韮崎市", "南アルプス市", "北杜市", "甲斐市",
    "笛吹市", "上野原市", "甲州市", "中央市", "市川三郷町", "富士川町", "身延町", "南部町", "昭和町",
    "富士河口湖町", "忍野村", "山中湖村", "鳴沢村", "西桂町", "道志村", "小菅村", "丹波山村", "早川町",
)
# 未ログインの Google マップは1検索 50〜120 件で打ち切られるため、市町村別の検索も加える
GMAPS_QUERIES = ["広告代理店 山梨県", "広告会社 山梨県", "広告代理店 甲府市", "広告会社 甲府市"] + [
    f"広告代理店 山梨県{city}" for city in YAMANASHI_CITIES if city != "甲府市"
]
NTA_BASE = "https://www.houjin-bangou.nta.go.jp"
NTA_PREF_YAMANASHI = "19"
# 並列数（相手サイトへの負荷を抑えるため 4 まで）
GMAPS_TABS = 4
NTA_WORKERS = 4

# Google マップにだけ載っている会社で、広告代理店とみなすカテゴリ
_AD_CATEGORY_RE = re.compile(r"広告|広報|マーケティング|販促|PR|プロモーション|デザイン|デザイナー|ウェブサイト|Web|ホームページ|"
                             r"看板|サイン|印刷|映像|ビデオ制作|イベント|メディア|出版|新聞")
# 名称から広告代理店でないとわかるもの（Google マップの検索結果に混ざる店舗・公共施設等）
_NON_AD_NAME_RE = re.compile(r"役所|役場|郵便局|銀行|信用金庫|農協|JA|コンビニ|ローソン|セブン|ファミリーマート|"
                             r"不動産|ホテル|病院|クリニック|学校|大学|保育|美容|整体|飲食|レストラン|カフェ")

# 検索で誤って別の会社・無関係のページが公式サイトURLとして入るもの
WRONG_URLS: tuple[str, ...] = (
    "matomame.jp", "kensetumap.com", "tsukulink.net", "tellogs.com", "postmap.org", "voix.jp",
    "glarity.app", "hotfrog.jp", "telinquire.com", "baseconnect.in", "ipros.jp", "metoree.com",
    "bing.com", "big-advance.site", "agencyhub.jp", "xn--bdk1d3bw36u9uu.com", "imitsu.jp",
    "web-kanji.com", "rekaizen.com", "biz.ne.jp",
)


def clean_site_fields(c: dict) -> None:
    """検索で拾ったまとめサイト等を公式サイト欄から外し、そこから取った連絡先も消す。"""
    url = c["公式サイトURL"]
    if "instagram.com" in url:
        c["インスタURL"] = c["インスタURL"] or url.split("?")[0].rstrip("/")
        c["公式サイトURL"] = ""
    elif any(w in url for w in WRONG_URLS):
        host = urlparse(url).netloc.removeprefix("www.")
        if c["メールアドレス"].endswith("@" + host):
            c["メールアドレス"] = ""
        c["公式サイトURL"] = c["問い合わせフォームURL"] = c["インスタURL"] = ""


def is_ad_agency(core: str, name: str, recs: list[dict]) -> bool:
    if _NON_AD_NAME_RE.search(_LEGAL_RE.sub("", _nfkc(name)).strip()):
        return False
    if any(r["src"] != "Googleマップ" for r in recs):
        return True  # 名簿・AgencyHub・NAVITIME の広告代理店カテゴリに掲載
    cats = [r.get("category", "") for r in recs]
    if any(_AD_CATEGORY_RE.search(c) for c in cats):
        return True
    # カテゴリが「企業のオフィス」等の汎用で、社名が広告会社を表すもの（アドコム 等）
    generic = all(c in ("", "企業のオフィス") for c in cats)
    return generic and bool(re.search(r"広告|宣伝|アド|プロモーション|エージェンシー", _nfkc(name)))


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


_YAMANASHI_MUNI = r"^(" + "|".join(YAMANASHI_CITIES) + r"|南都留郡|北都留郡|西八代郡|南巨摩郡|中巨摩郡)"

def _norm_addr(s: str) -> str:
    """所在地を「山梨県…」形式に整える（郵便番号・「住所:」を除去）。"""
    a = _nfkc(s)
    a = re.sub(r"^(?:住所|所在地)[:：]?\s*", "", a)
    a = re.sub(r"^[(（【]?本社[)）】]?[:：]?", "", a)
    a = re.sub(r"〒?\s*\d{3}-?\d{4}\s*", "", a).strip()
    a = re.sub(r"^日本、\s*", "", a)
    i = a.find("山梨県")
    if i > 0 and not re.search(r"[都道府]|.県", a[:i]):
        a = a[i:]
    if not a.startswith("山梨県") and re.match(_YAMANASHI_MUNI, a):
        a = "山梨県" + a
    return re.sub(r"\s+", " ", a)


# ── 社名の正規化（コア名） ─────────────────────────────────

_LEGAL_RE = re.compile(
    r"株式会社|有限会社|合同会社|合資会社|合名会社|一般社団法人|社会福祉法人|"
    r"\((?:株|有|同|資|名)\)|㈱|㈲"
)
# 「九州鉄鋼センター」「深江工作所」等は社名の一部なので、地名等を伴わない「センター」「製作所」は拠点表記としない
_BRANCH_KW = (
    r"営業所|支店|支社|支局|事業所|出張所|分室|オフィス|営業部|事業部"
)
_BRANCH_RE = re.compile(rf"({_BRANCH_KW})")
# 拠点名に使われる地名・区分
_PLACES = (
    "山梨|甲府|富士吉田|都留|大月|韮崎|南アルプス|北杜|甲斐|笛吹|上野原|甲州|中央|昭和|富士五湖|河口湖|峡東|峡南|"
    "第[一二三四五六七八九1-9]|[東西南北]|新"
)
_SUFFIX_RE = re.compile(rf"^(.+?)(?:{_PLACES})*(?:{_BRANCH_KW})(?:地区|.{{0,6}})?$")


def core_name(name: str) -> str:
    """法人格・括弧書き・工場/営業所表記を除いた社名。"""
    s = _nfkc(name)
    s = re.sub(r"[【\[(（](?!(?:株|有|同|資|名)\))[^】\])）]*[】\])）]", " ", s)
    # Google マップの「本社」「営業課」「技術部・工機課」、NAVITIME の「代表」等の部署単位の掲載
    s = re.sub(r"\s+(?:本社|代表|\S*[課係](?:・\S*[課係])?|\S+部(?:・\S*[課係])?)$", "", s).strip()
    m = _LEGAL_RE.search(s)
    if m and m.start() > 0:
        # 「X株式会社 九州工場」→ 法人格の前が社名
        core = s[:m.start()]
    else:
        body = _LEGAL_RE.sub(" ", s).strip()
        head, _, rest = body.partition(" ")
        # 「OSL LIGHTING」のような空白入りの社名は、後ろが拠点表記のときだけ分ける
        core = head if rest and _BRANCH_RE.search(rest) else body
        while (mm := _SUFFIX_RE.match(core.replace(" ", ""))) and len(mm.group(1)) >= 2 and mm.group(1) != core:
            core = mm.group(1)
    return core.replace(" ", "")


def is_branch(name: str) -> bool:
    """名称が支店・支局・営業所等の拠点を表しているか。"""
    s = _nfkc(name)
    s = re.sub(r"^.*?(?:株式会社|有限会社|\((?:株|有)\))", "", s) if _LEGAL_RE.match(s) else _LEGAL_RE.split(s)[-1] if _LEGAL_RE.search(s) else s
    return bool(re.search(r"営業所|支店|支社|支局|事業所|出張所|分室|オフィス|営業部", s))


# ── Phase 1: ネット検索 / NAVITIME ───────────────

def _rec(src: str, name: str, addr: str, tel: str = "", web: str = "", mail: str = "", category: str = "") -> dict:
    return {"src": src, "name": _nfkc(name), "addr": _norm_addr(addr), "tel": _format_phone(tel),
            "web": web, "mail": mail, "category": category}


def _session_get(s: requests.Session, url: str) -> str:
    for i in range(3):
        try:
            res = s.get(url, timeout=25)
            if res.status_code == 200:
                res.encoding = "utf-8"
                return res.text
        except requests.RequestException:
            pass
        time.sleep(3 * (i + 1))
    return ""


def _th_fields(soup: BeautifulSoup) -> dict:
    f = {}
    for th in soup.find_all("th"):
        td = th.find_next_sibling("td")
        if td is not None:
            f.setdefault(th.get_text(strip=True), td)
    return f


def fetch_meibo() -> list[dict]:
    """広告代理店の一覧（名簿）: 一覧（1ページ）→ 各詳細ページ。"""
    cached = _cache("meibo.json")
    if cached is not None:
        logger.info(f"Phase 1: 名簿 キャッシュから {len(cached)} 件")
        return cached
    s = requests.Session()
    s.headers["User-Agent"] = UA
    soup = BeautifulSoup(_session_get(s, MEIBO_LIST), "lxml")
    rows = []
    for tr in soup.select("table tr"):
        a = tr.find("a", href=re.compile(r"detail\.php"))
        tds = tr.find_all("td")
        if not a or len(tds) < 2:
            continue
        hp = tr.find("a", href=re.compile(r"^https?://"))
        rows.append({"name": a.get_text(strip=True), "href": urljoin(MEIBO_BASE, a["href"]),
                     "addr": tds[0].get_text(" ", strip=True), "tel": tds[1].get_text(strip=True),
                     "web": hp["href"] if hp else ""})
    logger.info(f"Phase 1: 名簿 一覧 {len(rows)} 件の詳細を取得")
    out = []
    for r in rows:
        time.sleep(random.uniform(0.5, 1.0))
        f = _th_fields(BeautifulSoup(_session_get(s, r["href"]), "lxml"))
        addr = f["住所"].get_text(" ", strip=True) if f.get("住所") else r["addr"]
        tel = f["電話番号"].get_text(strip=True) if f.get("電話番号") else r["tel"]
        hp = f["ホームページ"].find("a", href=True) if f.get("ホームページ") else None
        out.append(_rec("名簿", r["name"], addr, tel, hp["href"].strip() if hp else r["web"]))
    _save_cache("meibo.json", out)
    logger.info(f"Phase 1: 名簿 完了 {len(out)} 件")
    return out


def fetch_agencyhub() -> list[dict]:
    """AgencyHub: 山梨県の一覧 → 各会社の詳細ページ（所在地・電話番号・公式サイト・問い合わせページ）。"""
    cached = _cache("agencyhub.json")
    if cached is not None:
        logger.info(f"Phase 1: AgencyHub キャッシュから {len(cached)} 件")
        return cached
    s = requests.Session()
    s.headers["User-Agent"] = UA
    links, page, url = {}, 1, AGENCYHUB_LIST
    while url:
        soup = BeautifulSoup(_session_get(s, url), "lxml")
        for a in soup.select('a[href*="/agency/"]'):
            href = a["href"]
            if href.rstrip("/").endswith("/agency"):
                continue
            links.setdefault(urljoin(AGENCYHUB_LIST, href), "")
        page += 1
        nxt = soup.find("a", href=re.compile(rf"/page/{page}/?$"))
        url = urljoin(AGENCYHUB_LIST, nxt["href"]) if nxt else None
    logger.info(f"Phase 1: AgencyHub 一覧 {len(links)} 件の詳細を取得")
    out = []
    for href in links:
        time.sleep(random.uniform(0.5, 1.0))
        soup = BeautifulSoup(_session_get(s, href), "lxml")
        h1 = soup.find("h1")
        f = _th_fields(soup)
        web = f["公式サイト"].find("a", href=True) if f.get("公式サイト") else None
        form = f["問い合わせ"].find("a", href=True) if f.get("問い合わせ") else None
        rec = _rec("AgencyHub", h1.get_text(strip=True) if h1 else "",
                   f["所在地"].get_text(" ", strip=True) if f.get("所在地") else "",
                   f["電話番号"].get_text(strip=True) if f.get("電話番号") else "",
                   web["href"].strip() if web else "")
        # 問い合わせページがトップページと同じ（専用ページがない）ものは採用しない
        if form and form["href"].rstrip("/") != rec["web"].rstrip("/"):
            rec["form"] = form["href"].strip()
        out.append(rec)
    _save_cache("agencyhub.json", out)
    logger.info(f"Phase 1: AgencyHub 完了 {len(out)} 件")
    return out


def fetch_navitime() -> list[dict]:
    """一覧（1ページ15件、2ページ目以降は ?page=N）→ 各スポット詳細ページの住所・電話番号。"""
    cached = _cache("navitime.json")
    if cached is not None:
        logger.info(f"Phase 1: NAVITIME キャッシュから {len(cached)} 件")
        return cached
    s = requests.Session()
    s.headers["User-Agent"] = UA

    def _get_html(url: str) -> str:
        for i in range(3):
            try:
                res = s.get(url, timeout=25)
                if res.status_code == 200:
                    res.encoding = "utf-8"
                    return res.text
            except requests.RequestException:
                pass
            time.sleep(3 * (i + 1))
        return ""

    spots, seen, page = [], set(), 1
    while True:
        url = f"{NAVITIME_BASE}?tags={NAVITIME_TAG}&from=tag" if page == 1 else f"{NAVITIME_BASE}?page={page}&tags={NAVITIME_TAG}"
        soup = BeautifulSoup(_get_html(url), "lxml")
        new = 0
        for li in soup.select("li.spot-section"):
            a = li.select_one('a[href*="/poi?spot="]')
            if not a:
                continue
            href = urljoin("https://www.navitime.co.jp/", a["href"]).split("&")[0]
            if href in seen:
                continue
            seen.add(href)
            nm = li.select_one(".spot-name-text") or li.select_one(".spot-name") or a
            spots.append({"name": nm.get_text(strip=True), "href": href})
            new += 1
        if not new:
            break
        page += 1
        time.sleep(random.uniform(0.5, 1.0))
    logger.info(f"Phase 1: NAVITIME 一覧 {page - 1} ページ {len(spots)} 件の詳細を取得")
    out = []
    for sp in spots:
        time.sleep(random.uniform(0.5, 1.0))
        soup = BeautifulSoup(_get_html(sp["href"]), "lxml")
        fields = {}
        for dt in soup.select("dl.detail-text-frame dt"):
            dd = dt.find_next_sibling("dd")
            if dd is not None:
                fields[dt.get_text(strip=True)] = re.sub(r"\s+", " ", dd.get_text(" ", strip=True))
        addr = fields.get("住所", "").replace("地図で見る", "").strip()
        out.append(_rec("NAVITIME", sp["name"], addr, fields.get("電話番号", "")))
    _save_cache("navitime.json", out)
    logger.info(f"Phase 1: NAVITIME 完了 {len(out)} 件")
    return out


# ── Phase 2: Google マップ ───────────────────────────────

async def fetch_gmaps() -> list[dict]:
    cached = _cache("gmaps.json")
    if cached is not None:
        logger.info(f"Phase 2: キャッシュから {len(cached)} 件")
        return cached
    from playwright.async_api import async_playwright

    links: dict[str, str] = _cache("gmaps_links.json") or {}
    done_q: list[str] = _cache("gmaps_done_queries.json") or []
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        ctx = await browser.new_context(locale="ja-JP", user_agent=UA)
        ssem = asyncio.Semaphore(3)

        async def _search(q: str) -> None:
            if q in done_q:
                return
            async with ssem:
                pg = await ctx.new_page()
                try:
                    await pg.goto("https://www.google.com/maps/search/" + q.replace(" ", "+") + "?hl=ja",
                                  wait_until="domcontentloaded", timeout=60000)
                    await pg.wait_for_timeout(5000)
                    feed = pg.locator('div[role="feed"]')
                    if await feed.count():
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
                        items = await pg.eval_on_selector_all(
                            "a.hfpxzc", "as => as.map(a => [a.getAttribute('aria-label'), a.href])")
                    else:
                        # 1件だけヒットした場合は詳細ページが直接開く
                        items = [[await pg.locator("h1").first.inner_text(), pg.url]] if "/place/" in pg.url else []
                    new = 0
                    for label, href in items:
                        key = href.split("?")[0]
                        if key not in links:
                            links[key] = label
                            new += 1
                    done_q.append(q)
                    logger.info(f"  検索「{q}」{len(items)} 件（新規 {new} / 累計 {len(links)}）")
                except Exception as e:
                    logger.warning(f"  検索失敗「{q}」: {e}")
                finally:
                    await pg.close()
                    _save_cache("gmaps_links.json", links)
                    _save_cache("gmaps_done_queries.json", done_q)

        logger.info(f"Phase 2: Google マップ 検索 {len(GMAPS_QUERIES)} クエリ（済 {len(done_q)}）")
        await asyncio.gather(*[_search(q) for q in GMAPS_QUERIES])

        details: dict[str, dict] = _cache("gmaps_details.json") or {}
        todo = [h for h in links if h not in details]
        logger.info(f"Phase 2: Google マップ {len(todo)} 件の詳細を取得（取得済 {len(details)}）")
        sem = asyncio.Semaphore(GMAPS_TABS)
        cnt = {"n": 0}

        async def _detail(href: str, name: str) -> None:
            async with sem:
                tab = await ctx.new_page()
                rec = _rec("Googleマップ", name, "")
                try:
                    await tab.goto(href + "?hl=ja", wait_until="domcontentloaded", timeout=60000)
                    try:
                        await tab.wait_for_selector(
                            'button[data-item-id="address"], button[data-item-id^="phone:tel:"]', timeout=20000)
                    except Exception:
                        pass
                    await tab.wait_for_timeout(1200)
                    if await (h1 := tab.locator("h1").first).count():
                        rec["name"] = _nfkc(await h1.inner_text()) or rec["name"]
                    if await (a := tab.locator('button[data-item-id="address"]').first).count():
                        rec["addr"] = _norm_addr(await a.get_attribute("aria-label") or "")
                    if await (t := tab.locator('button[data-item-id^="phone:tel:"]').first).count():
                        rec["tel"] = _format_phone(await t.get_attribute("aria-label") or "")
                    if await (w := tab.locator('a[data-item-id="authority"]').first).count():
                        rec["web"] = await w.get_attribute("href") or ""
                    if await (c := tab.locator("button.DkEaL").first).count():
                        rec["category"] = (await c.inner_text()).strip()
                    if await tab.locator("text=閉業").count():
                        rec["closed"] = True
                    if rec["addr"] or rec["tel"] or rec.get("closed"):
                        details[href] = rec  # 読み込めなかったものは保存せず次回再取得
                except Exception as e:
                    logger.debug(f"詳細失敗 {name}: {e}")
                finally:
                    await tab.close()
                cnt["n"] += 1
                if cnt["n"] % 25 == 0:
                    logger.info(f"  詳細 {cnt['n']}/{len(todo)}")
                    _save_cache("gmaps_details.json", details)

        await asyncio.gather(*[_detail(h, links[h]) for h in todo])
        _save_cache("gmaps_details.json", details)
        await browser.close()
    # 閉業した事業所は除く
    out = [r for r in details.values() if not r.get("closed")]
    _save_cache("gmaps.json", out)
    logger.info(f"Phase 2: Google マップ 完了 {len(out)} 件")
    return out


# ── Phase 3: 名寄せ ──────────────────────────────────────

def group_companies(records: list[dict]) -> dict[str, list[dict]]:
    groups: dict[str, list[dict]] = {}
    for r in records:
        r["addr"] = re.sub(r"^住所[:：]\s*", "", r["addr"])
        r["addr"] = _norm_addr(r["addr"]) if not r["addr"].startswith("山梨県") else r["addr"]
        if not r["addr"].startswith("山梨県"):
            continue  # 県外の拠点・県外本社の掲載
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
    todo = [c for c in groups if c not in cache]
    logger.info(f"Phase 4: 法人番号公表サイトで本店所在地確認 {len(todo)} 件（キャッシュ {len(cache)} 件）")
    local = threading.local()  # 検索トークンはセッション単位なのでスレッドごとにセッションを持つ
    lock = threading.Lock()
    done = {"n": 0}

    def _one(core: str) -> None:
        if not hasattr(local, "nta"):
            local.nta = NTA()
        try:
            in_pref = [x for x in local.nta.search(core, NTA_PREF_YAMANASHI) if _same_core(x["name"], core)]
            time.sleep(1)
            other = []
            if not in_pref:
                other = [x for x in local.nta.search(core) if _same_core(x["name"], core)]
                time.sleep(1)
        except Exception as e:
            logger.warning(f"  法人番号検索失敗 {core}: {e}")
            return
        with lock:
            cache[core] = {"in_pref": in_pref, "other": other}
            done["n"] += 1
            if done["n"] % 50 == 0:
                _save_cache("nta.json", cache)
                logger.info(f"  {done['n']}/{len(todo)}")

    with ThreadPoolExecutor(NTA_WORKERS) as ex:
        list(ex.map(_one, todo))
    _save_cache("nta.json", cache)
    return cache


def _city(addr: str) -> str:
    """「山梨県甲府市…」→「甲府市」、「山梨県南都留郡富士河口湖町…」→「南都留郡富士河口湖町」。"""
    a = _nfkc(addr).removeprefix("山梨県")
    m = re.match(r"(.+?郡.+?[町村]|.+?市)", a)
    return m.group(1) if m else ""


def _hq_hits(core: str, recs: list[dict], nta: dict, town: bool = False) -> list[dict]:
    """県内の同名法人のうち、掲載拠点と同じ市区町村（town=True なら町名まで）に本店があるもの。"""
    def key(addr: str) -> str:
        return _city(addr) + (_town(addr)[:2] if town else "")
    keys = {key(r["addr"]) for r in recs if _city(r["addr"])}
    return [h for h in nta.get(core, {}).get("in_pref", []) if key(_nfkc(h["addr"])) in keys]


def _legal_form(name: str) -> str:
    s = _nfkc(name)
    for form, pat in (("株", r"株式会社|\(株\)|㈱"), ("有", r"有限会社|\(有\)|㈲"), ("同", r"合同会社|\(同\)")):
        if re.search(pat, s):
            return form
    return ""


def decide(core: str, recs: list[dict], nta: dict) -> tuple[bool, str]:
    """(山梨県本社か, 判定理由)"""
    res = nta.get(core, {"in_pref": [], "other": []})
    all_branch = all(is_branch(r["name"]) for r in recs)
    if res["in_pref"]:
        if all_branch:
            # 支店・営業所名義の掲載しかないときは、本店が同じ町にある場合だけ採用
            if _hq_hits(core, recs, nta, town=True):
                return True, "法人番号: 県内本店（掲載拠点と同じ町）"
            return False, "支店・営業所の掲載のみ（本店は別の場所）"
        if _hq_hits(core, recs, nta):
            return True, "法人番号: 県内本店（所在地一致）"
        # 同名法人がすべて県内にある（県外の会社の拠点ではない）ので、所在地が一致しなくても県内本社とみなす
        return True, "法人番号: 県内本店"
    if res["other"]:
        if all_branch:
            return False, "法人番号: 県外本店（支店・営業所の掲載）"
        # 「〇〇看板店」「アートクラフト」のような個人事業や、ありふれた社名で県外に同名の別会社が
        # あるだけのものは除かない。県外の同名法人が1社だけで法人格も一致するときだけ県外本社の拠点とみなす
        forms = {_legal_form(r["name"]) for r in recs} - {""}
        if len(res["other"]) == 1 and _legal_form(res["other"][0]["name"]) in forms:
            return False, "法人番号: 県外本店"
        return True, "法人番号: 県内登録なし（県外の同名法人は別会社）"
    if all_branch:
        return False, "支店・営業所表記・法人登録なし"
    return True, "法人登録なし（個人事業等）"


# ── Phase 5: 公式サイト検索 ────────────────────────────────

SKIP_DOMAINS = [
    "google.", "yahoo.", "bing.com", "facebook.com", "instagram.com", "twitter.com", "x.com",
    "youtube.com", "wikipedia", "jpnumber", "telnavi", "houjin", "baseconnect", "mapion",
    "navitime", "itp.ne.jp", "ekiten", "indeed", "townwork", "mynavi", "hellowork", "baitoru",
    "en-gage", "careerindex", "jobmedley", "job-medley",
    "kyujin", "recruit", "rikunabi", "doda", "stanby", "tdb.co.jp",
    "salesnow", "musubu", "alarmbox", "info-houjin", "nta.go.jp", "gbiz", "catr.jp", "graffer",
    "suumo", "homes.co.jp", "biz-maps", "jbplt", "baseconnect.in",
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


def _town(addr: str) -> str:
    """所在地の市区町村より後ろの町名・大字名（番地の手前まで）。"""
    a = _nfkc(addr).removeprefix("山梨県")
    city = _city(addr)
    a = a[len(city):] if city and a.startswith(city) else a
    a = re.sub(r"^大字|^字", "", a)
    return re.split(r"[\d一二三四五六七八九十〇]|字", a)[0]


# 企業情報のまとめ・検索サイト（社名と住所は載るが公式サイトではない）
DIRECTORY_DOMAINS = [
    "hakopro.jp", "data-link-plus.com", "kensetumap.com", "data-max.co.jp", "suke-dachi.jp",
    "jfsc.jp", "ivry.jp", "everytown.info", "x-work.jp", "cnavi", "nikkei", "buffett-code",
    "corporate-information", "kaisharesearch", "kigyou", "jigyosyo", "townpage", "goo.ne.jp",
    "compalyze", "tsujigawa", "apese.net", "town-guide", "hojin", "corp-", "bizdb", "clipbiz",
    *WRONG_URLS,
]
_DIRECTORY_TITLE_RE = re.compile(r"会社情報|企業情報|法人情報|口コミ|評判|求人|一覧|検索|電話番号|地図|アクセス情報")


def _local_phones(text: str) -> list[tuple[int, str]]:
    """ページ内の山梨県（055x）・フリーダイヤルの電話番号を (位置, 整形済み番号) で返す。"""
    out = []
    for m in re.finditer(r"(?<!\d)(0\d{1,4})[-‐ー−()（）\s]{0,2}(\d{1,4})[-‐ー−()（）\s]{0,2}(\d{3,4})(?!\d)", text):
        digits = "".join(m.groups())
        if re.search(r"FAX|Fax|fax|ファクス|ファックス", text[max(0, m.start() - 10):m.start()]):
            continue
        if len(digits) == 10 and digits.startswith(("055", "0120", "0800")):
            out.append((m.start(), _format_phone(digits)))
    return out


def search_official(core: str, addr: str) -> tuple[str, str]:
    """(公式サイトURL, まとめサイト等から得た電話番号) を返す。"""
    city, town = _city(addr), _town(addr)
    query = f"{core} {city or '山梨県'} {town}"
    try:
        with DDGS() as ddgs:
            results = list(ddgs.text(query, region="jp-jp", max_results=8))
    except Exception as e:
        logger.debug(f"検索失敗 {core}: {e}")
        return "", ""
    tel = ""
    city_key = re.sub(r"^.+?郡", "", city)
    for res in results:
        url = res.get("href", "")
        if not url.startswith("http") or _skip(url) or not _is_allowed(url):
            continue
        final, html = _get(url)
        if not html or _skip(final):
            continue
        soup = BeautifulSoup(html, "lxml")
        page = _nfkc(re.sub(r"\s", "", soup.get_text()))
        if core not in page or (city_key and city_key not in page) or (len(town) >= 2 and town not in page):
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
            best = min(((abs(p - a), t) for p, t in _local_phones(page) for a in anchors), default=None)
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
        phones = _local_phones(page)
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
    scache: dict = _cache("search.json") or {}

    async def _search(c: dict) -> None:
        key = f"{c['_core']}|{c['所在地']}"
        if key in scache:
            url, tel = scache[key]
            if url and not c["公式サイトURL"]:
                c["公式サイトURL"] = url
                c["_new_site"] = True
            if tel and not c["電話番号"]:
                c["電話番号"] = tel
            return
        async with sem:
            url, tel = await asyncio.to_thread(search_official, c["_core"], c["所在地"])
            scache[key] = [url, tel]
            _save_cache("search.json", scache)
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

_SRC_ORDER = {"名簿": 0, "AgencyHub": 1, "NAVITIME": 2, "Googleマップ": 3}


def _pick(recs: list[dict], key: str) -> str:
    """本社・非拠点のレコードを優先して値を選ぶ。"""
    ordered = sorted(recs, key=lambda r: (is_branch(r["name"]), _SRC_ORDER.get(r["src"], 9)))
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
    if is_branch(nm):
        return core
    return re.sub(r"[【(][^】)]*[】)]|\s*本社$", "", nm).strip()


async def main() -> None:
    logger.info("=" * 65)
    logger.info("山梨県 広告代理店リスト作成開始")
    logger.info("=" * 65)
    start = datetime.now()

    # 4 ソースを並列に収集する（requests 系はスレッド、Google マップは Playwright）
    meibo, hub, navi, gm = await asyncio.gather(
        asyncio.to_thread(fetch_meibo), asyncio.to_thread(fetch_agencyhub),
        asyncio.to_thread(fetch_navitime), fetch_gmaps(),
    )
    records = meibo + hub + navi + gm
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
        hit = _hq_hits(core, recs, nta) or nta.get(core, {}).get("in_pref", [])
        non_branch = [r for r in recs if not is_branch(r["name"])]
        non_branch.sort(key=lambda r: _SRC_ORDER.get(r["src"], 9))
        # 「山梨県」だけで市町村がない掲載は使わず、登記上の本店所在地にする
        with_city = [r for r in non_branch if _city(r["addr"])]
        addr = with_city[0]["addr"] if with_city else _norm_addr(hit[0]["addr"]) if hit else _pick(recs, "addr")
        addr = re.sub(r"[-−]番地", "番地", addr)  # Google マップの「大俣原8−番地」
        companies.append({
            "_core": core,
            "_reason": reason,
            "_recs": recs,
            "名称": _display_name(core, recs, hit),
            "メールアドレス": _pick(recs, "mail"),
            "公式サイトURL": _clean_url(_pick(recs, "web")),
            "所在地": addr,
            # 本社の掲載がない（県内本社の工場のみ掲載）場合、所在地は登記上の本店なので
            # 工場の電話番号は公式サイト等で本社番号が見つからなかったときの代替とする
            "電話番号": (_pick([r for r in non_branch if r["addr"] == addr], "tel") or _pick(non_branch, "tel")) if non_branch else "",
            "_branch_tel": "" if non_branch else _pick(recs, "tel"),
            "インスタURL": "",
            "問い合わせフォームURL": _pick(recs, "form"),
        })
    for core, reason, name in sorted(excluded):
        logger.info(f"  除外: {name}（{reason}）")
    logger.info(f"Phase 4: 山梨県本社 {len(companies)} 社 / 除外 {len(excluded)} 社")

    ads = [c for c in companies if is_ad_agency(c["_core"], c["名称"], c["_recs"])]
    for c in companies:
        if c not in ads:
            logger.info(f"  広告代理店以外のため除外: {c['名称']}")
    logger.info(f"Phase 7: 広告代理店 {len(ads)} 社 / それ以外 {len(companies) - len(ads)} 社を除外")
    companies = ads
    for c in companies:
        del c["_recs"]
        if any(w in c["公式サイトURL"] for w in WRONG_URLS):
            c["公式サイトURL"] = ""

    await enrich_all(companies)
    for c in companies:
        clean_site_fields(c)

    df = pd.DataFrame(companies)
    # 同じ会社が別名義で載ることがあるので、電話番号が同じものは1社にまとめる
    # （法人番号で確認できた名義・公式サイトありを残す）
    df["_rank"] = df["_reason"].str.startswith("法人番号").astype(int) * 2 + (df["公式サイトURL"] != "").astype(int)
    df = df.sort_values("_rank", ascending=False)
    dup = (df["電話番号"] != "") & df.duplicated("電話番号")
    for name in df.loc[dup, "名称"]:
        logger.info(f"  電話番号重複のため統合: {name}")
    df = df[~dup]
    # 名称の制御文字（Google マップの「￼」）・同じ語の重複（「ダイワプリントダイワプリント」）を直す
    df["名称"] = df["名称"].str.replace("￼", "", regex=False).str.strip()
    df["名称"] = df["名称"].str.replace(r"^(.{3,})\1$", r"\1", regex=True)
    # サイトの記入例（tarou@zmail.com 等）はメールアドレスとして使わない
    df.loc[df["メールアドレス"].str.contains(r"@zmail\.com$|^tarou@|^sample@|@example\.", regex=True), "メールアドレス"] = ""
    # 独自ドメインのメールが同じものは同じ会社の別名義（pccメディアマーケティング / p.c.c. 等）
    own = (df["メールアドレス"] != "") & ~df["メールアドレス"].str.contains(
        r"@(?:gmail\.com|yahoo\.co\.jp|.*\.ocn\.ne\.jp|.*\.plala\.or\.jp|.*point\.ne\.jp|nns\.ne\.jp)$", regex=True)
    dup = own & df.duplicated("メールアドレス")
    for name in df.loc[dup, "名称"]:
        logger.info(f"  メールアドレス重複のため統合: {name}")
    df = df[~dup]
    df = df.sort_values(["所在地", "名称"]).reset_index(drop=True)
    _save_cache("companies.json", df.to_dict("records"))
    df = df[OUTPUT_COLS]
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="広告代理店一覧")

    elapsed = int((datetime.now() - start).total_seconds())
    logger.info("=" * 65)
    logger.info(f"完了: {OUTPUT_FILE} に {len(df)} 件を出力")
    for col in OUTPUT_COLS[1:]:
        logger.info(f"  {col}: {(df[col] != '').sum()} 件")
    logger.info(f"所要時間: {elapsed // 60}分{elapsed % 60}秒")
    logger.info("=" * 65)


if __name__ == "__main__":
    asyncio.run(main())
