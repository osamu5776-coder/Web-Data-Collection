"""
北海道 WEB制作会社リスト作成

使い方（収集は別プロセスで並列実行できる）:
  python scraper_hokkaido_web.py wk      # Web制作会社まとめ図鑑（一覧全ページ＋各詳細）
  python scraper_hokkaido_web.py ks      # 全国Web制作会社ナビ（北海道＋各市ページ＋各詳細）
  python scraper_hokkaido_web.py gmaps   # Google マップ 179市町村＋札幌市10区（各詳細）
  python scraper_hokkaido_web.py merge   # 統合・本店判定・公式サイト補完・Excel出力

判定:
  - 本店: 全国Web制作会社ナビは法人番号で北海道の登記法人。ほかは国税庁法人番号公表サイトで確認し、
    道外本社の支店・営業所は除外する（本社のみ）。
  - WEB制作: まとめ図鑑掲載、ナビの詳細に WEB 制作の記載、Google マップのカテゴリがウェブ系、
    または社名・公式サイトに WEB 制作の記載があるもの。

出力列: 名称, メールアドレス, 公式サイトURL, 所在地, 電話番号, インスタURL, 問い合わせフォームURL
出力ファイル: 北海道_WEB制作会社リスト.xlsx
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

MODE = sys.argv[1] if len(sys.argv) > 1 else "merge"
LOG_FILE = f"scraper_hokkaido_web_{MODE}.log"
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
OUTPUT_FILE = "北海道_WEB制作会社リスト.xlsx"
CACHE_DIR = os.path.join("data", "hokkaido_web")
UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36"
)
WK_LIST = ("https://wk-partners.co.jp/homepage/dictionary/area/hokkaido/nisekocho,hakodateshi,kitahiroshimashi,"
           "kitamishi,chitoseshi,muroranshi,otarushi,iwanaicho,obihiroshi,asahikawashi,sapporoshi,ikeda,"
           "sunakawashi,mombetsushi,tomakomaishi,kushiroshi/")
KS_BASE = "https://web-design-company.ks-product.com/"
KS_LIST = KS_BASE + "area/hokkaido/"
CITIES = [c + "市" for c in (
    "札幌 函館 小樽 旭川 室蘭 釧路 帯広 北見 夕張 岩見沢 網走 留萌 苫小牧 稚内 美唄 芦別 江別 赤平 紋別 士別 名寄 "
    "三笠 根室 千歳 滝川 砂川 歌志内 深川 富良野 登別 恵庭 伊達 北広島 石狩 北斗").split()]
TOWNS = [t if t.endswith("村") else t + "町" for t in (
    "当別 新篠津村 松前 福島 知内 木古内 七飯 鹿部 森 八雲 長万部 江差 上ノ国 厚沢部 乙部 奥尻 今金 せたな 島牧村 寿都 "
    "黒松内 蘭越 ニセコ 真狩村 留寿都村 喜茂別 京極 倶知安 共和 岩内 泊村 神恵内村 積丹 古平 仁木 余市 赤井川村 南幌 "
    "奈井江 上砂川 由仁 長沼 栗山 月形 浦臼 新十津川 妹背牛 秩父別 雨竜 北竜 沼田 鷹栖 東神楽 当麻 比布 愛別 上川 東川 "
    "美瑛 上富良野 中富良野 南富良野 占冠村 和寒 剣淵 下川 美深 音威子府村 中川 幌加内 増毛 小平 苫前 羽幌 初山別村 遠別 "
    "天塩 猿払村 浜頓別 中頓別 枝幸 豊富 礼文 利尻 利尻富士 幌延 美幌 津別 斜里 清里 小清水 訓子府 置戸 佐呂間 遠軽 湧別 "
    "滝上 興部 西興部村 雄武 大空 豊浦 壮瞥 白老 厚真 洞爺湖 安平 むかわ 日高 平取 新冠 浦河 様似 えりも 新ひだか 音更 "
    "士幌 上士幌 鹿追 新得 清水 芽室 中札内村 更別村 大樹 広尾 幕別 池田 豊頃 本別 足寄 陸別 浦幌 釧路 厚岸 浜中 標茶 "
    "弟子屈 鶴居村 白糠 別海 中標津 標津 羅臼").split()]
WARDS = [f"札幌市{w}区" for w in "中央 北 東 白石 豊平 南 西 厚別 手稲 清田".split()]
assert len(CITIES) + len(TOWNS) == 179
GMAPS_QUERIES = ["WEB制作会社 北海道"] + [f"WEB制作会社 北海道{a}" for a in CITIES + WARDS + TOWNS] \
    + [f"ホームページ制作会社 北海道{a}" for a in WARDS + ["旭川市", "函館市", "帯広市", "釧路市", "北見市", "苫小牧市"]]
NTA_BASE = "https://www.houjin-bangou.nta.go.jp"
NTA_PREF_HOKKAIDO = "01"
GMAPS_TABS = 3
NTA_WORKERS = 4

# WEB 制作を表す語（社名・事業内容・公式サイト本文）
_WEB_NAME_RE = re.compile(r"ウェブ|web|ホームページ|インターネット|HP制作|サイト制作", re.I)
_WEB_WORK_RE = re.compile(
    r"(?:ホームページ|(?:web|ウェブ)\s*(?:サイト|ページ)?|HP|LP|ランディングページ|(?:EC|採用|企業|コーポレート)サイト)"
    r"[^。\n]{0,12}?(?:制作|製作|作成|構築|企画|デザイン|リニューアル|開発)|(?:web|ウェブ)\s*デザイン|らくらくホームページ", re.I)
# Google マップに混ざる教室・就労支援等
_NG_NAME_RE = re.compile(r"教室|スクール|カレッジ|WebStudy|就労|ジョブトレ|建築金物|パソコンの販売|学校|職業安定所|役場")
_WEB_CATEGORY_RE = re.compile(r"ウェブ|Web|ホームページ|インターネット マーケティング")
# 公式サイト本文の「サイトのデザイン」等では判定が緩すぎるため、業種が無関係な Google マップの掲載は対象にしない
_MAYBE_CATEGORY_RE = re.compile(r"広告|マーケティング|デザイン|デザイナー|ソフトウェア|コンピュータ|IT |情報|システム|印刷|"
                                r"映像|ビデオ|動画|写真|コンサルタント|企業のオフィス|メディア|プロモーション|ブランディング|電気通信")

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


def web_level(recs: list[dict]) -> str:
    """yes=WEB制作と判断 / maybe=公式サイトの本文で確認する / no"""
    if not any(r.get("no") or r["src"] == "まとめ図鑑" for r in recs) and any(_NG_NAME_RE.search(_nfkc(r["name"])) for r in recs):
        return "no"
    for r in recs:
        if r["src"] == "まとめ図鑑" or r.get("web_work"):
            return "yes"
        if _WEB_NAME_RE.search(_nfkc(r["name"])) or _WEB_CATEGORY_RE.search(r.get("category", "")):
            return "yes"
    if any(r["src"] == "ナビ" and r.get("web") or _MAYBE_CATEGORY_RE.search(r.get("category", "")) for r in recs):
        return "maybe"
    return "no"


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


_HOKKAIDO_MUNI = r"^(" + "|".join(CITIES + TOWNS) + r"|.{1,4}郡)"

def _norm_addr(s: str) -> str:
    """所在地を「北海道…」形式に整える（郵便番号・「住所:」を除去）。"""
    a = _nfkc(s)
    a = re.sub(r"^(?:住所|所在地)[:：]?\s*", "", a)
    a = re.sub(r"^[(（【]?本社[)）】]?[:：]?", "", a)
    a = re.sub(r"〒?\s*\d{3}-?\d{4}\s*", "", a).strip()
    a = re.sub(r"^日本、\s*", "", a)
    i = a.find("北海道")
    if i > 0 and not re.search(r"[都府]|.県", a[:i]):
        a = a[i:]
    if not a.startswith("北海道") and re.match(_HOKKAIDO_MUNI, a):
        a = "北海道" + a
    return re.sub(r"\s+", " ", a)


# ── 社名の正規化（コア名） ─────────────────────────────────

_LEGAL_RE = re.compile(
    r"株式会社|有限会社|合同会社|合資会社|合名会社|一般社団法人|社会福祉法人|"
    r"\((?:株|有|同|資|名)\)|㈱|㈲"
)
# 「九州鉄鋼センター」「深江工作所」等は社名の一部なので、地名等を伴わない「センター」「製作所」は拠点表記としない
_BRANCH_KW = (
    r"営業所|支店|支社|支局|支所|事業所|出張所|分室|営業部|事業部"
)
_BRANCH_RE = re.compile(rf"({_BRANCH_KW})")
# 拠点名に使われる地名・区分
_PLACES = (
    "北海道|札幌|函館|小樽|旭川|室蘭|釧路|帯広|北見|苫小牧|千歳|江別|十勝|道東|道北|道南|道央|北日本|"
    "第[一二三四五六七八九1-9]|[東西南北]|中央|新"
)
_SUFFIX_RE = re.compile(rf"^(.+?)(?:{_PLACES})*(?:{_BRANCH_KW})(?:地区|.{{0,6}})?$")


def core_name(name: str) -> str:
    """法人格・括弧書き・工場/営業所表記を除いた社名。"""
    s = _nfkc(name)
    s = re.sub(r"[【\[(（](?!(?:株|有|同|資|名)\))[^】\])）]*[】\])）]", " ", s)
    # Google マップの「本社」「営業課」「技術部・工機課」、NAVITIME の「代表」等の部署単位の掲載
    s = re.sub(r"\s+(?:本社|代表|\S*[課係](?:・\S*[課係])?|\S+部(?:・\S*[課係])?)$", "", s).strip()
    s = re.sub(r"\s\S+オフィス$", "", s)
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
    # 「オフィス○○」「○○オフィス」は社名のことが多いので、地名＋オフィス／空白＋…オフィスのときだけ拠点とする
    return bool(re.search(rf"営業所|支店|支社|支局|支所|事業所|出張所|分室|営業部|(?:{_PLACES})オフィス$|\s\S+オフィス$", s))


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


def _new_session() -> requests.Session:
    ses = requests.Session()
    ses.headers["User-Agent"] = UA
    return ses


def fetch_wk() -> list[dict]:
    """Web制作会社まとめ図鑑: 一覧全ページ → 各詳細ページ。"""
    cached = _cache("wk.json")
    if cached is not None:
        return cached
    ses, ids, page = _new_session(), [], 1
    while True:
        html = _session_get(ses, WK_LIST if page == 1 else f"{WK_LIST}page/{page}/")
        new = [i for i in dict.fromkeys(re.findall(r"homepage/dictionary/(\d+)/", html)) if i not in ids]
        if not new:
            break
        ids += new
        logger.info(f"図鑑 {page} ページ: {len(new)} 件（累計 {len(ids)}）")
        if f"/page/{page + 1}/" not in html:
            break
        page += 1
        time.sleep(1)
    out = []
    for i in ids:
        time.sleep(random.uniform(0.7, 1.2))
        soup = BeautifulSoup(_session_get(ses, f"https://wk-partners.co.jp/homepage/dictionary/{i}/"), "lxml")
        f = {k: v.get_text(" ", strip=True) for k, v in _th_fields(soup).items()}
        if not f.get("会社名"):
            logger.warning(f"図鑑 詳細を読めない: {i}")
            continue
        rec = _rec("まとめ図鑑", f["会社名"], f.get("所在地", ""), f.get("電話番号", ""), f.get("WebサイトURL", ""))
        rec["addr_raw"] = f.get("所在地", "")
        out.append(rec)
    _save_cache("wk.json", out)
    logger.info(f"図鑑 完了 {len(out)} 件")
    return out


def _ks_get(ses: requests.Session, url: str) -> str:
    """全国Web制作会社ナビ用。10分あたり約100件で 429 になるため間隔を空け、Retry-After に従って待つ。"""
    for _ in range(4):
        time.sleep(6.5)
        try:
            r = ses.get(url, timeout=25)
        except requests.RequestException:
            continue
        if r.status_code == 429:
            wait = int(r.headers.get("Retry-After") or 600) + 5
            logger.info(f"  429: {wait} 秒待機")
            time.sleep(wait)
            continue
        if r.status_code == 200:
            r.encoding = "utf-8"
            return r.text
        break
    return ""


def fetch_ks() -> list[dict]:
    """全国Web制作会社ナビ: 北海道ページ＋各市ページ（「さらに表示」の JSON を含む）→ 詳細ページ。"""
    cached = _cache("ks.json")
    if cached is not None:
        return cached
    ses = _new_session()
    top = _ks_get(ses, KS_LIST)
    pages = [KS_LIST] + sorted({urljoin(KS_LIST, h) for h in re.findall(r'href="([^"]*area/hokkaido/[a-z]+-shi/)"', top)})
    nums: dict[str, bool] = {}
    plain: dict[str, tuple] = {}
    for url in pages:
        html = _ks_get(ses, url)
        for n in re.findall(r"company/(\d{13})/", html):
            nums[n] = True
        soup = BeautifulSoup(html, "lxml")
        # 公式サイト未確認の会社（詳細ページなし）は、社名と所在地だけが並ぶ
        lines = (soup.find("main") or soup.body).get_text("\n", strip=True).split("\n")
        for a, b in zip(lines, lines[1:]):
            if b.startswith("北海道") and re.search(_LEGAL_RE, a) and len(a) < 60 and len(b) < 90 and "。" not in a + b:
                plain[_nfkc(a)] = (a, b)
        for j in sorted(set(re.findall(r"[\w./-]*data/[\w-]+\.json", html))):
            txt = _ks_get(ses, urljoin(url, j))
            time.sleep(0.5)
            try:
                rows = json.loads(txt)
            except ValueError:
                continue
            for row in rows:
                if re.fullmatch(r"\d{13}", str(row[0])):
                    nums[row[0]] = True
                else:
                    plain[_nfkc(row[0])] = (row[0], row[1])
        logger.info(f"ナビ {url}: 詳細あり累計 {len(nums)} / 名称のみ累計 {len(plain)}")
        time.sleep(0.7)
    det: dict[str, dict] = _cache("ks_details.json") or {}
    todo = [n for n in nums if n not in det]
    logger.info(f"ナビ 詳細取得 {len(todo)} 件")
    for k, n in enumerate(todo, 1):
        soup = BeautifulSoup(_ks_get(ses, f"{KS_BASE}company/{n}/"), "lxml")
        f = _th_fields(soup)
        if "所在地" not in f:
            logger.warning(f"ナビ 詳細を読めない: {n}")
            continue
        text = (soup.find("main") or soup.body).get_text("\n", strip=True)
        a = text.find("の特徴")
        b = text.find("\nPR\n", a)
        desc = text[a:b] if a >= 0 and b > a else ""
        hp = f["公式サイト"].find("a", href=True) if f.get("公式サイト") else None
        rec = _rec("ナビ", f["商号"].get_text(strip=True) if f.get("商号") else soup.h1.get_text(strip=True),
                   f["所在地"].get_text(" ", strip=True), "", hp["href"] if hp else "")
        rec.update(no=n, web_work=bool(_WEB_WORK_RE.search(_nfkc(desc))))
        det[n] = rec
        if k % 10 == 0:
            _save_cache("ks_details.json", det)
            logger.info(f"  詳細 {k}/{len(todo)}")
    _save_cache("ks_details.json", det)
    have = {r["name"].replace(" ", "") for r in det.values()}
    out = list(det.values()) + [_rec("ナビ", a, b) for k, (a, b) in plain.items() if k.replace(" ", "") not in have]
    if len(det) == len(nums):
        _save_cache("ks.json", out)
    logger.info(f"ナビ 完了 詳細あり {len(det)}/{len(nums)} / 名称のみ {len(out) - len(det)}")
    return out


def _place_id(href: str) -> str:
    m = re.search(r"!1s(0x[0-9a-f]+:0x[0-9a-f]+)", href)
    return m.group(1) if m else href.split("?")[0]


async def _block(route) -> None:
    if route.request.resource_type in ("image", "media", "font"):
        await route.abort()
    else:
        await route.continue_()


async def run_gmaps() -> None:
    from playwright.async_api import async_playwright

    links: dict[str, list] = _cache("gmaps_links.json") or {}
    done_q: list[str] = _cache("gmaps_done_queries.json") or []
    details: dict[str, dict] = _cache("gmaps_details.json") or {}

    async def _search(ctx, q: str) -> None:
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
                items = await pg.eval_on_selector_all("a.hfpxzc", "as => as.map(a => [a.getAttribute('aria-label'), a.href])")
            else:
                items = [[await pg.locator("h1").first.inner_text(), pg.url]] if "/place/" in pg.url else []
            before = len(links)
            for label, href in items:
                links.setdefault(_place_id(href), [label, href.split("?")[0]])
            done_q.append(q)
            logger.info(f"  検索「{q}」{len(items)} 件（新規 {len(links) - before} / 累計 {len(links)}）")
        except Exception as e:
            logger.warning(f"  検索失敗「{q}」: {str(e)[:120]}")
        finally:
            await pg.close()

    async def _detail(ctx, pid: str) -> None:
        name, href = links[pid]
        tab = await ctx.new_page()
        try:
            rec = _rec("Googleマップ", name, "")
            await tab.goto(href + "?hl=ja", wait_until="domcontentloaded", timeout=45000)
            try:
                await tab.wait_for_selector('button[data-item-id="address"], button[data-item-id^="phone:tel:"]', timeout=15000)
            except Exception:
                pass
            await tab.wait_for_timeout(800)
            if await (h := tab.locator("h1").first).count():
                rec["name"] = _nfkc(await h.inner_text()) or rec["name"]
            if await (a := tab.locator('button[data-item-id="address"]').first).count():
                rec["addr_raw"] = await a.get_attribute("aria-label") or ""
                rec["addr"] = _norm_addr(rec["addr_raw"])
            if await (t := tab.locator('button[data-item-id^="phone:tel:"]').first).count():
                rec["tel"] = _format_phone(await t.get_attribute("aria-label") or "")
            if await (w := tab.locator('a[data-item-id="authority"]').first).count():
                rec["web"] = await w.get_attribute("href") or ""
            if await (c := tab.locator("button.DkEaL").first).count():
                rec["category"] = (await c.inner_text()).strip()
            if await tab.locator("text=閉業").count():
                rec["closed"] = True
            # 住所非公開（出張対応）の事業者はカテゴリだけ表示される
            if rec["addr"] or rec["tel"] or rec["category"] or rec.get("closed"):
                details[pid] = rec
        except Exception as e:
            logger.debug(f"詳細失敗 {name}: {e}")
        finally:
            await tab.close()

    async with async_playwright() as p:
        async def _batches(todo: list, worker, size: int, label: str, save) -> None:
            """size 件ごとにブラウザを開き直して処理する（長時間の連続取得で遅くなるため）。"""
            for i in range(0, len(todo), size):
                browser = await p.chromium.launch(headless=True)
                ctx = await browser.new_context(locale="ja-JP", user_agent=UA)
                await ctx.route("**/*", _block)
                sem = asyncio.Semaphore(GMAPS_TABS)

                async def _one(x):
                    async with sem:
                        try:
                            await asyncio.wait_for(worker(ctx, x), timeout=240)
                        except Exception as e:
                            logger.warning(f"  {label}失敗 {x}: {type(e).__name__}")

                await asyncio.gather(*[_one(x) for x in todo[i:i + size]])
                save()
                logger.info(f"  {label} {min(i + size, len(todo))}/{len(todo)}")
                await browser.close()

        def _save_links():
            _save_cache("gmaps_links.json", links)
            _save_cache("gmaps_done_queries.json", done_q)

        for _ in range(2):
            todo = [q for q in GMAPS_QUERIES if q not in done_q]
            if todo:
                logger.info(f"Google マップ 検索 {len(todo)} クエリ")
                await _batches(todo, _search, 15, "検索", _save_links)
        for rnd in range(1, 4):
            todo = [x for x in links if x not in details]
            if not todo:
                break
            logger.info(f"Google マップ 詳細 {rnd} 回目 {len(todo)} 件（取得済 {len(details)}）")
            await _batches(todo, _detail, 60, "詳細", lambda: _save_cache("gmaps_details.json", details))
    logger.info(f"Google マップ 完了 検索 {len(done_q)}/{len(GMAPS_QUERIES)} 詳細 {len(details)}/{len(links)}")


# ── Phase 3: 名寄せ ──────────────────────────────────────

def group_companies(records: list[dict]) -> dict[str, list[dict]]:
    by_core: dict[str, list[dict]] = {}
    for r in records:
        if not r["addr"].startswith("北海道"):
            continue
        if (c := core_name(r["name"])):
            by_core.setdefault(c.lower(), []).append(r)
    groups: dict[str, list[dict]] = {}
    for recs in by_core.values():
        core = core_name(recs[0]["name"])
        nos = list(dict.fromkeys(r["no"] for r in recs if r.get("no")))
        if len(nos) <= 1:
            groups[core] = recs
            continue
        parts = {n: [r for r in recs if r.get("no") == n] for n in nos}
        for r in recs:
            if not r.get("no"):
                n = next((n for n in nos if _city(parts[n][0]["addr"]) == _city(r["addr"])), nos[0])
                parts[n].append(r)
        for i, n in enumerate(nos):
            groups[core if i == 0 else f"{core}#{n}"] = parts[n]
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
    # ナビ掲載（法人番号が北海道の登記）の会社は確認不要
    todo = [c for c, recs in groups.items() if c not in cache and not any(r.get("no") for r in recs)]
    logger.info(f"Phase 4: 法人番号公表サイトで本店所在地確認 {len(todo)} 件（キャッシュ {len(cache)} 件）")
    local = threading.local()  # 検索トークンはセッション単位なのでスレッドごとにセッションを持つ
    lock = threading.Lock()
    done = {"n": 0}

    def _one(core: str) -> None:
        if not hasattr(local, "nta"):
            local.nta = NTA()
        try:
            in_pref = [x for x in local.nta.search(core, NTA_PREF_HOKKAIDO) if _same_core(x["name"], core)]
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
    """「北海道札幌市…」→「札幌市」、「北海道虻田郡ニセコ町…」→「虻田郡ニセコ町」。"""
    a = _nfkc(addr).removeprefix("北海道")
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
    """(北海道本社か, 判定理由)"""
    if any(r.get("no") for r in recs):
        return True, "法人番号: 道内本店（ナビ）"
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
        return True, "道内登録なし（道外の同名法人は別会社）"
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
    a = _nfkc(addr).removeprefix("北海道")
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
    """ページ内の北海道（01x）・フリーダイヤルの電話番号を (位置, 整形済み番号) で返す。"""
    out = []
    for m in re.finditer(r"(?<!\d)(0\d{1,4})[-‐ー−()（）\s]{0,2}(\d{1,4})[-‐ー−()（）\s]{0,2}(\d{3,4})(?!\d)", text):
        digits = "".join(m.groups())
        if re.search(r"FAX|Fax|fax|ファクス|ファックス", text[max(0, m.start() - 10):m.start()]):
            continue
        if len(digits) == 10 and digits.startswith(("01", "0800")) and not digits.startswith("0180"):
            out.append((m.start(), _format_phone(digits)))
    return out


def search_official(core: str, addr: str) -> tuple[str, str]:
    """(公式サイトURL, まとめサイト等から得た電話番号) を返す。"""
    city, town = _city(addr), _town(addr)
    query = f"{core} {city or '北海道'} {town}"
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
    if _WEB_WORK_RE.search(_nfkc(text)):
        result["_web"] = True

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
    cache: dict = _cache("sites.json") or {}
    sem = asyncio.Semaphore(10)
    cnt = {"n": 0}

    async def _one(c: dict) -> None:
        url = c["公式サイトURL"]
        async with sem:
            if url not in cache:
                try:
                    cache[url] = await asyncio.wait_for(asyncio.to_thread(enrich_site, url, c["_core"].split("#")[0]), 120)
                except asyncio.TimeoutError:
                    cache[url] = {}
                cnt["n"] += 1
                if cnt["n"] % 50 == 0:
                    _save_cache("sites.json", cache)
                    logger.info(f"  公式サイト {cnt['n']}/{len(targets)}")
        extra = cache[url]
        for k in ("メールアドレス", "インスタURL", "問い合わせフォームURL"):
            if extra.get(k) and not c.get(k):
                c[k] = extra[k]
        if not c["電話番号"] and extra.get("_tel"):
            c["電話番号"] = extra["_tel"]
        c["_web"] = c.get("_web") or bool(extra.get("_web"))

    await asyncio.gather(*[_one(c) for c in targets])
    _save_cache("sites.json", cache)


async def enrich_all(companies: list[dict]) -> None:
    with_url = [c for c in companies if c["公式サイトURL"]]
    logger.info(f"Phase 6: 公式サイトから連絡先取得 {len(with_url)} 件")
    await _enrich_sites(with_url)

    targets = [c for c in companies if c["_level"] == "yes" and (not c["公式サイトURL"] or not c["電話番号"])]
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
            url, tel = await asyncio.to_thread(search_official, c["_core"].split("#")[0], c["所在地"])
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

_SRC_ORDER = {"ナビ": 0, "まとめ図鑑": 1, "Googleマップ": 2}


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


def _own_host(url: str) -> str:
    host = urlparse(url).netloc.lower().removeprefix("www.")
    shared = r"wix|jimdo|goope|wordpress|ameba|ameblo|fc2|peraichi|instagram|facebook|x\.com|twitter|google|yahoo|line\.me|" \
             r"studio\.site|strikingly|crayonsite|note\.|lit\.link|hp\.peraichi|base\.shop|github|linktr|notion|sakura\.ne|moo\.jp"
    return "" if not host or re.search(shared, host) else host


def _merge_same(companies: list[dict], key) -> list[dict]:
    """key が同じ行を1社にまとめる。法人番号がともにあって異なるものは別会社として残す。"""
    out: list[dict] = []
    idx: dict[str, list[dict]] = {}
    for c in companies:
        k = key(c)
        hit = next((o for o in idx.get(k, []) if not (o["_no"] and c["_no"] and o["_no"] != c["_no"])), None) if k else None
        if hit:
            logger.info(f"  重複のため統合: {hit['名称']} ← {c['名称']}")
            for col in OUTPUT_COLS:
                hit[col] = hit[col] or c[col]
            hit["_no"] = hit["_no"] or c["_no"]
            continue
        out.append(c)
        if k:
            idx.setdefault(k, []).append(c)
    return out


async def main() -> None:
    start = datetime.now()
    gm = [r for r in (_cache("gmaps_details.json") or {}).values() if not r.get("closed")]
    wk, ks = _cache("wk.json") or [], (_cache("ks.json") or list((_cache("ks_details.json") or {}).values()))
    logger.info(f"統合: 図鑑 {len(wk)} / ナビ {len(ks)} / Google マップ {len(gm)}")
    groups = group_companies(wk + ks + gm)
    logger.info(f"名寄せ後 {len(groups)} 社")
    # WEB 制作と無関係な掲載は本店確認の前に落とす
    level = {c: web_level(r) for c, r in groups.items()}
    groups = {c: r for c, r in groups.items() if level[c] != "no"}
    logger.info(f"WEB制作: 確定 {sum(v == 'yes' for v in level.values())} / 公式サイトで確認 {sum(v == 'maybe' for v in level.values())}"
                f" / 対象外 {sum(v == 'no' for v in level.values())}")
    nta = check_hq(groups)
    companies, excluded = [], []
    for core, recs in groups.items():
        ok, reason = decide(core, recs, nta)
        hit = [] if any(r.get("no") for r in recs) else (_hq_hits(core, recs, nta) or nta.get(core, {}).get("in_pref", [])[:1])
        recs_s = sorted(recs, key=lambda r: (is_branch(r["name"]), _SRC_ORDER.get(r["src"], 9)))
        c = {
            "_core": core, "_reason": reason, "_level": level[core], "_web": False,
            "_no": next((r["no"] for r in recs if r.get("no")), ""), "_branch_tel": "",
            "_src": "+".join(sorted({r["src"] for r in recs})),
            "名称": _display_name(core.split("#")[0], recs, hit if ok else []),
            "メールアドレス": "", "公式サイトURL": _clean_url(_pick(recs, "web")),
            "所在地": recs_s[0]["addr"], "電話番号": _pick(recs, "tel"), "インスタURL": "", "問い合わせフォームURL": "",
        }
        if any(w in c["公式サイトURL"] for w in WRONG_URLS):
            c["公式サイトURL"] = ""
        if not ok:
            c["名称"] = _nfkc(recs_s[0]["name"])
        (companies if ok else excluded).append(c)
    logger.info(f"本店判定: 道内 {len(companies)} / 除外 {len(excluded)}")

    for c in excluded:
        logger.info(f"  除外: {c['名称']}（{c['_reason'].replace('県', '道')}）")
    await enrich_all(companies)
    for c in companies:
        clean_site_fields(c)
    dropped = [c for c in companies if c["_level"] == "maybe" and not c["_web"]]
    companies = [c for c in companies if c["_level"] == "yes" or c["_web"]]
    logger.info(f"公式サイトで WEB 制作を確認できず除外 {len(dropped)} / 残り {len(companies)}")

    bad = re.compile(r"@(?:zmail|mail|example|sample|xxx+|domain)\.|form-mailer|themeansar|@city\.|^(?:email|tarou?|taro|sample|test|your|xxx+)@|\.(?:png|jpe?g|webp|svg)$", re.I)
    for c in companies + excluded:
        n = re.sub(r"[\ufffc★☆◆◇■□●○♪]+", "", c["名称"]).strip()
        c["名称"] = re.sub(r"^(.{3,})\1$", r"\1", n)
        if c["電話番号"].startswith("010"):  # 存在しない市外局番（サイト内の番号の誤検出）
            c["電話番号"] = ""
        if c["名称"] == "株式会社KO":  # ナビの誤記（公式サイトは KOO）
            c["名称"] = "株式会社KOO"
        if bad.search(c["メールアドレス"]):
            c["メールアドレス"] = ""
    companies.sort(key=lambda c: (not c["_no"], not c["_reason"].startswith("法人番号"), not c["公式サイトURL"], len(c["名称"])))
    for key in (lambda c: re.sub(r"\D", "", c["電話番号"]) if not c["電話番号"].startswith(("0120", "0800")) else "",
                lambda c: _own_host(c["公式サイトURL"])):
        companies = _merge_same(companies, key)
    companies.sort(key=lambda c: (c["所在地"], c["名称"]))
    _save_cache("companies.json", companies)
    _save_cache("excluded.json", excluded)
    _save_cache("dropped.json", dropped)
    df = pd.DataFrame(companies)[OUTPUT_COLS]
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="WEB制作会社一覧")
        for ws in w.sheets.values():
            ws.freeze_panes = "A2"
            for col, width in zip("ABCDEFG", [38, 32, 40, 50, 16, 36, 46]):
                ws.column_dimensions[col].width = width
    logger.info(f"完了: {OUTPUT_FILE} に {len(df)} 件を出力（{(datetime.now() - start).seconds // 60}分）")
    for col in OUTPUT_COLS:
        logger.info(f"  {col}: {(df[col] != '').sum()} 件")


if __name__ == "__main__":
    if MODE == "wk":
        fetch_wk()
    elif MODE == "ks":
        fetch_ks()
    elif MODE == "gmaps":
        asyncio.run(run_gmaps())
    else:
        asyncio.run(main())
