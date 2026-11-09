"""
兵庫県 法律事務所 情報収集（本店・主たる事務所が兵庫県のもののみ）

使い方（収集フェーズは別プロセスで並列実行できる）:
  python scraper_hyogo_bengoshi.py bengo4   # 弁護士ドットコム 兵庫 全ページ＋詳細
  python scraper_hyogo_bengoshi.py gmaps    # Google マップ 市町村（神戸市は区）別 全件＋詳細
  python scraper_hyogo_bengoshi.py merge    # 統合・公式サイト補完・本店判定・Excel出力

データソース:
  - 弁護士ドットコム https://www.bengo4.com/office/hyogo/ （一覧全ページ・各詳細ページ）
  - Google マップ「法律事務所 兵庫県<市町村>」「弁護士事務所 兵庫県<市町村>」（各詳細ページ）
  ※ 日弁連 弁護士情報検索（member.nichibenren.or.jp）は robots.txt で全ページの
    巡回が禁止（Disallow: /）されているため自動取得しない。
    兵庫県弁護士会サイトの会員ページは認証必須で公開名簿がない。

本店判定:
  弁護士法人の従たる事務所（県外本部の「神戸支店」「神戸オフィス」等）を除くため、
  弁護士法人名・支店表記のある事務所は国税庁法人番号公表サイトで主たる事務所の
  所在地を確認し、県外のみに登録がある場合は除外する。個人事務所（法人でない）は
  その所在地が本拠のため県内所在であれば採用する。

出力列: 名称, メールアドレス, 公式サイトURL, 所在地, 電話番号, インスタURL, 問い合わせフォームURL
出力ファイル: 兵庫県_法律事務所リスト.xlsx
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
from urllib.parse import urljoin, urlparse
from urllib.robotparser import RobotFileParser

import pandas as pd
import requests
from bs4 import BeautifulSoup, XMLParsedAsHTMLWarning

MODE = sys.argv[1] if len(sys.argv) > 1 else "merge"

# ── ロギング ──────────────────────────────────────────────
LOG_FILE = f"scraper_hyogo_bengoshi_{MODE}.log"
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
OUTPUT_FILE = "兵庫県_法律事務所リスト.xlsx"
CACHE_DIR = os.path.join("data", "hyogo_bengoshi")
UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36"
)
HEADERS = {"User-Agent": UA}
BENGO4_BASE = "https://www.bengo4.com/office/hyogo/"
NTA_BASE = "https://www.houjin-bangou.nta.go.jp"
NTA_PREF_HYOGO = "28"

KOBE_WARDS = ["東灘区", "灘区", "中央区", "兵庫区", "北区", "長田区", "須磨区", "垂水区", "西区"]
CITIES = [
    "姫路市", "尼崎市", "明石市", "西宮市", "洲本市", "芦屋市", "伊丹市", "相生市", "豊岡市",
    "加古川市", "赤穂市", "西脇市", "宝塚市", "三木市", "高砂市", "川西市", "小野市", "三田市",
    "加西市", "丹波篠山市", "養父市", "丹波市", "南あわじ市", "朝来市", "淡路市", "宍粟市",
    "加東市", "たつの市", "猪名川町", "多可町", "稲美町", "播磨町", "市川町", "福崎町",
    "神河町", "太子町", "上郡町", "佐用町", "香美町", "新温泉町",
]
AREAS = [f"神戸市{w}" for w in KOBE_WARDS] + CITIES
# 未ログインの Google マップは1検索 120 件前後で打ち切られるため、事務所の多い地域は語を変えて重ねる
BIG_AREAS = {"神戸市中央区", "姫路市", "尼崎市", "西宮市", "明石市", "加古川市", "伊丹市", "宝塚市", "芦屋市", "川西市"}
GMAPS_QUERIES = [f"法律事務所 兵庫県{a}" for a in AREAS] + [f"弁護士事務所 兵庫県{a}" for a in AREAS if a in BIG_AREAS] \
    + [f"弁護士 兵庫県{a}" for a in AREAS if a in BIG_AREAS]
GMAPS_TABS = 2
GMAPS_SEARCH_TABS = 3

SKIP_DOMAINS = [
    "bengo4.com", "nichibenren.or.jp", "houterasu.or.jp", "hyogoben.or.jp",
    "itp.ne.jp", "mapion.co.jp", "navitime.co.jp", "mapfan.com", "goo.ne.jp", "wikipedia.org",
    "google.com", "google.co.jp", "bing.com", "yahoo.co.jp", "mynavi.jp", "houmu-cafe.com",
    "bennavi.jp", "bengoshi-search", "ben54.jp", "kakekomu.com", "ekiten.jp", "jinjibu.jp",
    "jp.indeed.com", "indeed.com", "stanby.com", "hellowork", "doda.jp", "en-japan.com",
    "ameblo.jp", "seesaa.net", "livedoor.jp", "hatenablog.com", "blogspot.com", "note.com",
    "instagram.com", "facebook.com", "twitter.com", "x.com", "youtube.com", "tiktok.com",
    "line.me", "lin.ee", "baseconnect.in", "houjin.info", "houjin-bangou", "kaisharesearch",
    "best-legal.jp", "legalus.jp", "keiji-pro.com", "rikon-pro.com", "souzoku-pro",
    "shakkin-pro", "roudou-pro", "jiko-pro", "hakopro", "compalyze",
]
OUTPUT_COLS = ["名称", "メールアドレス", "公式サイトURL", "所在地", "電話番号", "インスタURL", "問い合わせフォームURL"]
# 法律事務所ではない士業・機関
NON_LAWYER_RE = re.compile(r"司法書士|行政書士|税理士|社会保険労務士|社労士|土地家屋調査士|弁理士|特許|会計|公証|"
                           r"法テラス|司法支援センター|弁護士会|裁判所|検察|法務局|市役所|町役場|相談センター|大学|ロースクール")
LAWYER_RE = re.compile(r"法律事務所|弁護士|法律|法務事務所|lawoffice|law office|law firm|lawfirm|legal", re.I)


def _nfkc(s: str) -> str:
    return unicodedata.normalize("NFKC", str(s or "")).strip()


def _cache(name: str):
    p = os.path.join(CACHE_DIR, name)
    if os.path.exists(p):
        with open(p, encoding="utf-8") as f:
            return json.load(f)
    return None


def _save_cache(name: str, obj) -> None:
    os.makedirs(CACHE_DIR, exist_ok=True)
    with open(os.path.join(CACHE_DIR, name), "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=1)


_robots_cache: dict[str, RobotFileParser] = {}


def _is_allowed(url: str) -> bool:
    parsed = urlparse(url)
    origin = f"{parsed.scheme}://{parsed.netloc}"
    if origin not in _robots_cache:
        rp = RobotFileParser()
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


def _get(url: str, timeout: int = 15) -> requests.Response | None:
    for attempt in range(3):
        try:
            r = requests.get(url, headers=HEADERS, timeout=timeout, allow_redirects=True)
            if r.status_code == 429 or r.status_code >= 500:
                time.sleep(5 * (attempt + 1))
                continue
            r.encoding = r.apparent_encoding or "utf-8"
            return r
        except Exception as e:
            logger.debug(f"取得失敗 {url}: {e}")
            time.sleep(2 * (attempt + 1))
    return None


def _digits(s: str) -> str:
    return re.sub(r"\D", "", _nfkc(s))


def _format_phone(s: str) -> str:
    s = _nfkc(s)
    m = re.search(r"0\d{1,4}-\d{1,4}-\d{3,4}", s)
    if m:
        return m.group(0)
    d = _digits(s)
    if len(d) == 10:
        if d.startswith("078") or d.startswith("079") or d.startswith("06"):
            return f"{d[:2]}-{d[2:6]}-{d[6:]}" if d.startswith("06") else f"{d[:3]}-{d[3:6]}-{d[6:]}"
        return f"{d[:4]}-{d[4:6]}-{d[6:]}"
    if len(d) == 11:
        return f"{d[:3]}-{d[3:7]}-{d[7:]}"
    return s


def _norm_addr(s: str) -> str:
    s = _nfkc(s)
    s = re.sub(r"^(住所|所在地)\s*[:：]\s*", "", s)
    s = re.sub(r"^日本、?\s*", "", s)
    s = re.sub(r"〒?\s*\d{3}-?\d{4}\s*", "", s).strip()
    s = re.sub(r"\s+", " ", s)
    if s and not s.startswith("兵庫県") and re.match(r"(神戸|姫路|尼崎|明石|西宮|洲本|芦屋|伊丹|相生|豊岡|加古川|赤穂|西脇|宝塚|三木|高砂|川西|小野|三田|加西|丹波|養父|南あわじ|朝来|淡路|宍粟|加東|たつの|川辺郡|多可郡|加古郡|神崎郡|揖保郡|赤穂郡|佐用郡|美方郡)", s):
        s = "兵庫県" + s
    return s


# ── 弁護士ドットコム ─────────────────────────────────────

def _parse_bengo4_list(html: str, base: str) -> list[dict]:
    soup = BeautifulSoup(html, "lxml")
    out = []
    for h3 in soup.select("h3.list_office__name"):
        a = h3.find("a", href=True)
        if not a:
            continue
        tel = ""
        parent = h3.find_parent("div", class_="list_office__summary")
        if parent and (tp := parent.select_one("p.list_office__tel")):
            m = re.search(r"[\d-]{9,}", tp.get_text(strip=True))
            tel = m.group(0) if m else ""
        out.append({"src": "bengo4", "name": a.get_text(strip=True), "detail": urljoin(base, a["href"]), "tel": tel})
    return out


def _parse_bengo4_detail(html: str) -> dict:
    soup = BeautifulSoup(html, "lxml")
    res = {"addr": "", "web": ""}
    for dt in soup.find_all("dt"):
        label = dt.get_text(strip=True)
        dd = dt.find_next_sibling("dd")
        if not dd:
            continue
        if label == "所在地":
            res["addr"] = _norm_addr(dd.get_text(" ", strip=True))
        elif "事務所URL" in label and (a := dd.find("a", href=True)):
            href = a["href"].strip()
            if href.startswith("http") and not any(d in href for d in SKIP_DOMAINS):
                res["web"] = href
    return res


def run_bengo4() -> None:
    cached = _cache("bengo4.json")
    if cached is not None:
        logger.info(f"bengo4 キャッシュあり {len(cached)} 件")
        return
    r = _get(BENGO4_BASE)
    if r is None:
        logger.error("bengo4 一覧取得失敗")
        return
    pages = max([int(x) for x in re.findall(r"page=(\d+)", r.text)] or [1])
    total_m = re.search(r"([\d,]+)件", r.text)
    logger.info(f"bengo4 兵庫: {total_m.group(1) if total_m else '?'}件 / {pages}ページ")
    recs = _parse_bengo4_list(r.text, BENGO4_BASE)
    for p in range(2, pages + 1):
        time.sleep(random.uniform(1.0, 2.0))
        url = f"{BENGO4_BASE}?page={p}"
        rr = _get(url)
        got = _parse_bengo4_list(rr.text, url) if rr is not None else []
        recs.extend(got)
        logger.info(f"  page {p}/{pages}: {len(got)} 件（累計 {len(recs)}）")
    seen, uniq = set(), []
    for x in recs:
        if x["detail"] not in seen:
            seen.add(x["detail"])
            uniq.append(x)
    logger.info(f"bengo4 詳細取得 {len(uniq)} 件")
    for i, x in enumerate(uniq, 1):
        if _is_allowed(x["detail"]):
            rr = _get(x["detail"])
            if rr is not None and rr.status_code < 400:
                x.update(_parse_bengo4_detail(rr.text))
        if i % 20 == 0:
            logger.info(f"  詳細 {i}/{len(uniq)}")
        time.sleep(random.uniform(0.8, 1.5))
    _save_cache("bengo4.json", uniq)
    logger.info(f"bengo4 完了 {len(uniq)} 件")


# ── Google マップ ──────────────────────────────────────────

async def run_gmaps() -> None:
    cached = _cache("gmaps.json")
    if cached is not None:
        logger.info(f"gmaps キャッシュあり {len(cached)} 件")
        return
    from playwright.async_api import async_playwright

    links: dict[str, str] = _cache("gmaps_links.json") or {}
    done_q: list[str] = _cache("gmaps_done_queries.json") or []
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        ctx = await browser.new_context(locale="ja-JP", user_agent=UA)
        sem = asyncio.Semaphore(GMAPS_SEARCH_TABS)

        async def _search(q: str) -> None:
            if q in done_q:
                return
            async with sem:
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

        logger.info(f"gmaps 検索 {len(GMAPS_QUERIES)} クエリ（済 {len(done_q)}）")
        await asyncio.gather(*[_search(q) for q in GMAPS_QUERIES])

        # 名称で明らかに法律事務所でないものは詳細を開かない
        targets = {h: n for h, n in links.items() if not (NON_LAWYER_RE.search(n or "") and "法律事務所" not in (n or ""))}
        details: dict[str, dict] = _cache("gmaps_details.json") or {}
        todo = [h for h in targets if h not in details]
        logger.info(f"gmaps 詳細取得 {len(todo)} 件（全 {len(targets)} / 取得済 {len(details)}）")
        dsem = asyncio.Semaphore(GMAPS_TABS)
        cnt = {"n": 0}

        async def _detail(href: str, name: str) -> None:
            async with dsem:
                tab = await ctx.new_page()
                rec = {"src": "gmaps", "name": name, "addr": "", "tel": "", "web": "", "category": "", "detail": href}
                try:
                    await tab.goto(href + "?hl=ja", wait_until="domcontentloaded", timeout=60000)
                    try:
                        # 並列で開くと描画が遅れるため、住所か電話のボタンが出るまで待つ
                        await tab.wait_for_selector(
                            'button[data-item-id="address"], button[data-item-id^="phone:tel:"]', timeout=20000)
                    except Exception:
                        pass
                    await tab.wait_for_timeout(1500)
                    h1 = tab.locator("h1").first
                    if await h1.count():
                        t = (await h1.inner_text()).strip()
                        if t:
                            rec["name"] = t
                    a = tab.locator('button[data-item-id="address"]').first
                    if await a.count():
                        rec["addr"] = _norm_addr(await a.get_attribute("aria-label") or "")
                    t = tab.locator('button[data-item-id^="phone:tel:"]').first
                    if await t.count():
                        rec["tel"] = _format_phone(await t.get_attribute("aria-label") or "")
                    w = tab.locator('a[data-item-id="authority"]').first
                    if await w.count():
                        rec["web"] = await w.get_attribute("href") or ""
                    c = tab.locator("button.DkEaL").first
                    if await c.count():
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

        await asyncio.gather(*[_detail(h, targets[h]) for h in todo])
        _save_cache("gmaps_details.json", details)
        await browser.close()
    _save_cache("gmaps.json", list(details.values()))
    logger.info(f"gmaps 完了 {len(details)} 件")


async def run_gmaps_shard(idx: int, total: int) -> None:
    """未取得の詳細ページを total 分割した idx 番目だけ、独立したブラウザで取得する。"""
    from playwright.async_api import async_playwright

    links: dict[str, str] = _cache("gmaps_links.json") or {}
    base: dict = _cache("gmaps_details.json") or {}
    targets = {h: n for h, n in links.items() if not (NON_LAWYER_RE.search(n or "") and "法律事務所" not in (n or ""))}
    todo_all = sorted(h for h in targets if h not in base)
    shard_file = f"gmaps_details_shard{idx}.json"
    mine: dict = _cache(shard_file) or {}
    todo = [h for h in todo_all[idx::total] if h not in mine]
    logger.info(f"shard {idx}/{total}: {len(todo)} 件")
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        ctx = await browser.new_context(locale="ja-JP", user_agent=UA)
        tab = await ctx.new_page()
        for i, href in enumerate(todo, 1):
            rec = {"src": "gmaps", "name": targets[href], "addr": "", "tel": "", "web": "", "category": "", "detail": href}
            try:
                await tab.goto(href + "?hl=ja", wait_until="domcontentloaded", timeout=60000)
                try:
                    await tab.wait_for_selector(
                        'button[data-item-id="address"], button[data-item-id^="phone:tel:"]', timeout=20000)
                except Exception:
                    pass
                await tab.wait_for_timeout(1000)
                if await (h1 := tab.locator("h1").first).count():
                    rec["name"] = (await h1.inner_text()).strip() or rec["name"]
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
                    mine[href] = rec
            except Exception as e:
                logger.debug(f"詳細失敗 {href}: {e}")
            if i % 10 == 0 or i == len(todo):
                _save_cache(shard_file, mine)
                logger.info(f"  shard {idx}: {i}/{len(todo)}（取得 {len(mine)}）")
        await browser.close()
    _save_cache(shard_file, mine)


def combine_gmaps_shards() -> None:
    base: dict = _cache("gmaps_details.json") or {}
    for f in sorted(os.listdir(CACHE_DIR)):
        if f.startswith("gmaps_details_shard"):
            base.update(_cache(f) or {})
    _save_cache("gmaps_details.json", base)
    _save_cache("gmaps.json", list(base.values()))
    logger.info(f"gmaps シャード統合 {len(base)} 件")


# ── 補助金ナビ（認定経営革新等支援機関のうち弁護士） ─────────────
# robots.txt に Crawl-delay: 10 の指定があるため、このサイトだけは並列にせず 10 秒間隔で取得する

HOJO_BASE = "https://hojo-navi.com/shop/prefecture/hyogo/"
HOJO_DELAY = 10.5
_hojo_last = {"t": 0.0}


def _hojo_get(url: str) -> BeautifulSoup | None:
    wait = HOJO_DELAY - (time.time() - _hojo_last["t"])
    if wait > 0:
        time.sleep(wait)
    r = _get(url, timeout=30)
    _hojo_last["t"] = time.time()
    if r is None or r.status_code >= 400:
        logger.warning(f"  取得失敗 {url}")
        return None
    return BeautifulSoup(r.text, "lxml")


def run_hojonavi() -> None:
    st = _cache("hojo_state.json") or {"lines": {}, "stations": {}, "lawyers": {}, "details": {}}

    def save():
        _save_cache("hojo_state.json", st)

    # 1) 兵庫県の路線一覧
    if not st["lines"]:
        soup = _hojo_get(HOJO_BASE)
        for a in soup.select('a[href*="/shop/train-line/"]'):
            href = a["href"].split("/page/")[0].rstrip("/") + "/"
            st["lines"].setdefault(href, {"name": a.get_text(strip=True), "done": False})
        save()
    logger.info(f"補助金ナビ 路線 {len(st['lines'])} 本")

    # 2) 各路線の全ページから兵庫県の駅（掲載件数つき）
    for href, ln in st["lines"].items():
        if ln["done"]:
            continue
        page, url = 1, href
        while url:
            soup = _hojo_get(url)
            if soup is None:
                break
            for a in soup.select('a[href*="/shop/prefecture/hyogo/station/"]'):
                m = re.search(r"合計[:：]\s*(\d+)", a.get_text(" ", strip=True))
                cnt = int(m.group(1)) if m else -1
                st["stations"].setdefault(a["href"], {"name": a.get_text(" ", strip=True).split(" ")[0], "count": cnt, "done": False})
            page += 1
            nxt = f"{href}page/{page}/"
            url = nxt if soup.select_one(f'a[href="{nxt}"]') else None
        ln["done"] = True
        save()
        logger.info(f"  路線 {ln['name']}: 駅累計 {len(st['stations'])}")

    # 3) 掲載のある駅ページから種別「弁護士」の行
    todo = [h for h, s in st["stations"].items() if not s["done"] and s["count"] != 0]
    logger.info(f"補助金ナビ 駅 {len(st['stations'])}（掲載あり未取得 {len(todo)}）")
    for i, href in enumerate(todo, 1):
        soup = _hojo_get(href)
        if soup is not None:
            for tr in soup.select("table.tbl_detail_info tbody tr"):
                tds = tr.find_all("td")
                a = tds[0].find("a", href=True) if tds else None
                if not a or len(tds) < 3 or "弁護士" not in tds[2].get_text():
                    continue
                st["lawyers"].setdefault(a["href"], {"label": a.get_text("\n", strip=True), "addr": tds[1].get_text(strip=True)})
            st["stations"][href]["done"] = True
        if i % 10 == 0 or i == len(todo):
            save()
            logger.info(f"  駅 {i}/{len(todo)}（弁護士 {len(st['lawyers'])} 件）")

    # 4) 弁護士の詳細ページ
    todo = [h for h in st["lawyers"] if h not in st["details"]]
    logger.info(f"補助金ナビ 弁護士詳細 {len(todo)} 件")
    for i, href in enumerate(todo, 1):
        soup = _hojo_get(href)
        if soup is None:
            continue
        f = {}
        for th in soup.find_all(["th", "dt"]):
            td = th.find_next_sibling(["td", "dd"])
            if td is not None:
                f.setdefault(th.get_text(strip=True), td.get_text(" ", strip=True))
        st["details"][href] = f
        if i % 10 == 0 or i == len(todo):
            save()
            logger.info(f"  詳細 {i}/{len(todo)}")
    save()

    out = []
    for href, lw in st["lawyers"].items():
        f = st["details"].get(href, {})
        parts = lw["label"].split("\n")
        person = re.sub(r"[\s　]+", "", parts[0])
        office = parts[1].strip("()（） ") if len(parts) > 1 else ""
        name = office if office and office != parts[0] else person
        if not re.search(r"法律|弁護士|事務所|法人", name):
            name = f"弁護士 {person}"   # 事務所名の記載がない個人登録
        web = f.get("URL", "").strip()
        email = f.get("メール", "").strip()
        out.append({
            "src": "hojonavi", "name": _nfkc(name), "detail": href,
            "addr": _norm_addr(f.get("住所") or lw["addr"]),
            "tel": _format_phone(f.get("電話番号", "")) if f.get("電話番号") else "",
            "web": web if web.startswith("http") else (f"http://{web}" if re.match(r"[\w.-]+\.\w{2,}", web) else ""),
            "email": email if "@" in email else "",
            "branch": f.get("本店・支店", ""), "kind": f.get("認定支援機関種別", ""),
        })
    _save_cache("hojonavi.json", out)
    logger.info(f"補助金ナビ 完了: 弁護士 {len(out)} 件（メール {sum(1 for r in out if r['email'])} / URL {sum(1 for r in out if r['web'])}）")


# ── 統合 ─────────────────────────────────────────────────

def _key_name(name: str) -> str:
    s = _nfkc(name).lower()
    s = re.sub(r"[\s・･\-‐ー－（）()「」]", "", s)
    s = s.replace("弁護士法人", "").replace("法律事務所", "").replace("弁護士事務所", "")
    return s


def _city(addr: str) -> str:
    m = re.search(r"兵庫県\s*((?:神戸市)?[^\s市]{1,6}?[市区]|[^\s]{1,4}郡[^\s]{1,4}?町)", addr)
    return m.group(1) if m else ""


def _is_lawyer(rec: dict) -> bool:
    name, cat = rec.get("name", ""), rec.get("category", "")
    if rec.get("src") == "bengo4":
        return True
    if rec.get("closed"):
        return False
    if NON_LAWYER_RE.search(name) and "法律事務所" not in name:
        return False
    if re.search(r"司法書士|行政書士|税理士|探偵|保険|デイサービス|介護", name):
        return False
    name_ok = bool(re.search(r"法律事務所|弁護士", name))
    if re.search(r"弁護士|法律事務所", cat):
        return True
    # 司法書士・行政書士等のカテゴリや未分類は、名称に「法律事務所」「弁護士」がある場合のみ
    return name_ok


def _key_addr(addr: str) -> str:
    s = _nfkc(addr).replace("丁目", "-").replace("番地", "-").replace("番", "-").replace("号", "")
    s = re.sub(r"[\s‐－−ー―]", "-", s)
    s = re.sub(r"-+", "-", s)
    return re.sub(r"[\s-]", "", s)


def merge_records(b4: list[dict], gm: list[dict], hj: list[dict] | None = None) -> list[dict]:
    recs = [r for r in b4] + [r for r in gm if _is_lawyer(r)] + list(hj or [])
    recs = [r for r in recs if "兵庫県" in r.get("addr", "") or (r["src"] == "bengo4" and not r.get("addr"))]
    parent = list(range(len(recs)))

    def find(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    def union(i, j):
        parent[find(i)] = find(j)

    by_phone, by_name = {}, {}
    for i, r in enumerate(recs):
        d = _digits(r.get("tel", ""))
        if len(d) >= 10 and not d.startswith("050"):
            if d in by_phone:
                union(i, by_phone[d])
            else:
                by_phone[d] = i
        k = (_key_name(r["name"]), _city(r.get("addr", "")))
        if k[0]:
            if k in by_name:
                union(i, by_name[k])
            else:
                by_name[k] = i
    # 補助金ナビの個人名登録（「弁護士 ○○」）は同じ住所の事務所にまとめる
    by_addr: dict[str, int] = {}
    for i, r in enumerate(recs):
        if r["src"] != "hojonavi" and r.get("addr"):
            by_addr.setdefault(_key_addr(r["addr"]), i)
    for i, r in enumerate(recs):
        if r["src"] == "hojonavi" and r["name"].startswith("弁護士 ") and r.get("addr"):
            j = by_addr.get(_key_addr(r["addr"]))
            if j is not None:
                union(i, j)
    groups: dict[int, list[dict]] = {}
    for i, r in enumerate(recs):
        groups.setdefault(find(i), []).append(r)

    out = []
    for g in groups.values():
        b = next((r for r in g if r["src"] == "bengo4"), None)
        m = next((r for r in g if r["src"] == "gmaps"), None)
        h = next((r for r in g if r["src"] == "hojonavi"), None)
        name = (b or m or h)["name"]
        tels = [r.get("tel", "") for r in g if r.get("tel")]
        tel = next((t for t in tels if not _digits(t).startswith("050")), tels[0] if tels else "")
        addr = (b or {}).get("addr") or (m or {}).get("addr") or (h or {}).get("addr", "")
        web = next((r.get("web") for r in g if r.get("web") and not any(d in r["web"] for d in SKIP_DOMAINS)), "")
        # 補助金ナビの登録は改称・移転前のままのことがある（例: はりま中央→レクシード）ため、
        # メールは事務所名が一致する（または補助金ナビにしか載っていない）場合だけ採用する
        email = next((r["email"] for r in g if r.get("email")
                      and (r is (b or m or h) or _key_name(r["name"]) == _key_name(name))), "")
        addr = re.sub(r"^兵庫県\s+", "兵庫県", addr)
        out.append({"名称": _nfkc(name), "所在地": addr, "電話番号": _format_phone(tel) if tel else "",
                    "公式サイトURL": web, "メールアドレス": email, "インスタURL": "", "問い合わせフォームURL": "",
                    "_src": "+".join(sorted({r["src"] for r in g}))})
    return out


# ── 公式サイト補完 ───────────────────────────────────────

def _decode_cfemail(encoded: str) -> str:
    try:
        key = int(encoded[:2], 16)
        return bytes(int(encoded[i:i + 2], 16) ^ key for i in range(2, len(encoded), 2)).decode("utf-8")
    except Exception:
        return ""


BAD_EMAIL_RE = re.compile(r"\.(png|jpg|jpeg|gif|svg|webp)$|example\.|sentry|wixpress|@domain|your-?mail|xxx", re.I)


def _parse_official_html(url: str, html: str) -> dict:
    res: dict = {}
    soup = BeautifulSoup(html, "lxml")
    text = soup.get_text("\n")
    host = urlparse(url).netloc
    emails = []
    for el in soup.find_all(attrs={"data-cfemail": True}):
        if (e := _decode_cfemail(el["data-cfemail"])):
            emails.append(e)
    for a in soup.find_all("a", href=True):
        if a["href"].lower().startswith("mailto:"):
            emails.append(a["href"][7:].split("?")[0].strip())
    emails += re.findall(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}", text)
    emails = [e for e in emails if "@" in e and not BAD_EMAIL_RE.search(e)]
    if emails:
        res["メールアドレス"] = emails[0]
    for a in soup.find_all("a", href=True):
        h = a["href"]
        if "instagram.com" in h and "/p/" not in h and "/reel/" not in h and "share" not in h:
            res["インスタURL"] = ("https:" + h if h.startswith("//") else h).rstrip("/")
            break
    kws = ["contact", "inquiry", "toiawase", "form", "お問い合わせ", "問い合わせ", "お問合せ", "問合せ", "ご相談予約", "相談予約", "メールフォーム"]
    for a in soup.find_all("a", href=True):
        h, t = a["href"], a.get_text(strip=True)
        if not any(k in h.lower() or k in t for k in kws):
            continue
        if h.startswith(("#", "mailto:", "tel:", "javascript")):
            continue
        h = "https:" + h if h.startswith("//") else urljoin(url, h)
        if urlparse(h).netloc != host:
            continue
        res["問い合わせフォームURL"] = h
        break
    m = re.search(r"0(?:78|79\d?|6|7\d{2})[-‐(（ ]\d{1,4}[-‐)） ]\d{4}", _nfkc(text))
    if m:
        res["_site_phone"] = _format_phone(m.group(0))
    return res


def _fetch_official(url: str) -> tuple[str, dict]:
    if not url.startswith("http") or not _is_allowed(url):
        return url, {}
    r = _get(url, timeout=12)
    if r is None or r.status_code >= 400:
        return url, {}
    info = _parse_official_html(r.url, r.text)
    # トップに連絡先がなければ問い合わせ／アクセスページも見る
    if not info.get("メールアドレス") and info.get("問い合わせフォームURL"):
        r2 = _get(info["問い合わせフォームURL"], timeout=12)
        if r2 is not None and r2.status_code < 400:
            sub = _parse_official_html(r2.url, r2.text)
            if sub.get("メールアドレス"):
                info["メールアドレス"] = sub["メールアドレス"]
    return r.url, info


def _verify(html: str, name: str, phone: str) -> bool:
    page = _nfkc(html)
    d = _digits(phone)
    if d and not d.startswith("050") and d in re.sub(r"\D", "", page):
        return True
    core = re.sub(r"^弁護士法人", "", _nfkc(name))
    core = re.sub(r"\s.*$", "", core)
    core = re.sub(r"(法律事務所|弁護士事務所|事務所)$", "", core).strip()
    return len(core) >= 2 and (core + "法律事務所") in page.replace(" ", "") or (len(core) >= 4 and core in page)


PORTAL_TITLE_KWS = ["求人", "転職", "口コミ", "ランキング", "一覧", "検索", "を探す", "おすすめ", "比較", "まとめ"]


def _search_official(rec: dict) -> tuple[str, str]:
    from ddgs import DDGS
    city = _city(rec["所在地"]) or "兵庫県"
    q = f"{rec['名称']} {city} 弁護士"
    results = []
    for backend in ("brave", "yandex", "google", "duckduckgo"):
        try:
            with DDGS(timeout=10) as dd:
                results = list(dd.text(q, max_results=6, backend=backend))
            if results:
                break
        except Exception:
            continue
    for r in results:
        url = r.get("href", "")
        if not url.startswith("http") or any(d in url for d in SKIP_DOMAINS) or not _is_allowed(url):
            continue
        resp = _get(url, timeout=10)
        if resp is None or resp.status_code >= 400 or any(d in resp.url for d in SKIP_DOMAINS):
            continue
        tm = re.search(r"<title[^>]*>(.*?)</title>", resp.text, re.I | re.S)
        if tm and any(k in tm.group(1) for k in PORTAL_TITLE_KWS):
            continue
        if _verify(resp.text, rec["名称"], rec["電話番号"]):
            root = f"{urlparse(resp.url).scheme}://{urlparse(resp.url).netloc}/"
            return root, resp.url
    return "", ""


async def enrich(out: list[dict]) -> None:
    cache = _cache("official.json") or {}
    sem = asyncio.Semaphore(10)

    # 公式サイト未判明のものは Web 検索で探す
    search_cache = _cache("search.json") or {}
    need = [r for r in out if not r["公式サイトURL"] and r["名称"] not in search_cache]
    logger.info(f"公式サイト検索 {len(need)} 件（キャッシュ {len(search_cache)}）")
    ssem = asyncio.Semaphore(6)
    cnt = {"n": 0}

    async def _s(r):
        async with ssem:
            try:
                root, _ = await asyncio.wait_for(asyncio.to_thread(_search_official, r), timeout=60)
            except asyncio.TimeoutError:
                root = ""
            search_cache[r["名称"]] = root
            cnt["n"] += 1
            if cnt["n"] % 20 == 0:
                logger.info(f"  検索 {cnt['n']}/{len(need)}")
                _save_cache("search.json", search_cache)
            await asyncio.sleep(random.uniform(1.0, 2.0))

    await asyncio.gather(*[_s(r) for r in need])
    _save_cache("search.json", search_cache)
    for r in out:
        if not r["公式サイトURL"] and search_cache.get(r["名称"]):
            r["公式サイトURL"] = search_cache[r["名称"]]

    todo = sorted({r["公式サイトURL"] for r in out if r["公式サイトURL"] and r["公式サイトURL"] not in cache})
    logger.info(f"公式サイト取得 {len(todo)} 件（キャッシュ {len(cache)}）")

    async def _f(u):
        async with sem:
            try:
                _, info = await asyncio.wait_for(asyncio.to_thread(_fetch_official, u), timeout=40)
            except asyncio.TimeoutError:
                logger.info(f"  公式サイト取得タイムアウト: {u}")
                info = {}
            cache[u] = info

    await asyncio.gather(*[_f(u) for u in todo])
    _save_cache("official.json", cache)
    for r in out:
        info = dict(cache.get(r["公式サイトURL"], {}))
        sp = info.pop("_site_phone", "")
        if sp and (not r["電話番号"] or _digits(r["電話番号"]).startswith("050")):
            r["電話番号"] = sp
        for k, v in info.items():
            if v and not r.get(k):
                r[k] = v


# ── 本店判定（国税庁 法人番号公表サイト） ────────────────

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
            th, tds = tr.find("th"), tr.find_all("td")
            if not th or len(tds) < 2:
                continue
            out.append({"name": tds[0].get_text("\n", strip=True).split("\n")[-1], "addr": tds[1].get_text(strip=True)})
        return out


BRANCH_RE = re.compile(r"支店|支所|支部|分室|オフィス|出張所|事務所$")


def hq_core(name: str) -> str:
    """弁護士法人名・支店表記を除いた事務所の固有名（法人番号検索用）。"""
    s = _nfkc(name)
    s = re.sub(r"[（(][^）)]*[）)]", " ", s)
    s = s.replace("弁護士法人", " ").replace("弁護士事務所", "法律事務所 ")
    s = re.sub(r"法律\S{0,4}?事務所.*$", "", s)  # 「〇〇法律(経済)事務所 神戸支店」→「〇〇」
    s = re.sub(r"\s+\S*(支店|支所|支部|オフィス|事務所)$", "", s.strip())
    return s.replace(" ", "").strip()


# 法人番号公表サイトでの登録名が事務所名の表記と異なるもの
# 虎ノ門法律経済事務所 = 弁護士法人ＴＬＥＯ虎ノ門法律経済事務所（東京都港区）
NTA_ALIASES = {"アディーレ": "ADIRE", "虎ノ門": "TLEO虎ノ門"}


def _nta_core(name: str) -> str:
    s = _nfkc(name).lower().replace(" ", "")
    s = re.sub(r"法律\S{0,4}?事務所$", "", s)  # 法律事務所・法律経済事務所・法律会計事務所
    return re.sub(r"弁護士法人|法律事務所|弁護士事務所", "", s)


def needs_hq_check(name: str) -> bool:
    s = _nfkc(name)
    if "弁護士法人" in s or re.search(r"支店|支所|支部|分室|オフィス", s):
        return True
    after = re.split(r"法律\S{0,4}?事務所|弁護士事務所", s, maxsplit=1)
    return len(after) == 2 and bool(re.search(r"支店|支所|支部|オフィス|事務所|分室", after[1]))


def hq_candidates(name: str) -> list[str]:
    """法人番号検索に使う法人名の候補。

    登録名は「弁護士法人みお」「弁護士法人ＡＬＧ＆Ａｓｓｏｃｉａｔｅｓ」のように「弁護士法人」直後の
    部分だけのことが多いので、まずそれを使い、次に支店表記等を除いた事務所名を使う。
    """
    s = _nfkc(name)
    cands: list[str] = []
    m = re.search(r"弁護士法人[\s・]*([^()（）]*)", s)
    if m:
        rest = re.sub(r"^法律事務所", "", m.group(1).strip())
        tok = re.split(r"\s|支所|支店|分室|オフィス|法律会計事務所|法律事務所|事務所", rest)[0]
        if tok:
            cands.append(tok)
            short = re.split(r"綜合|総合", tok)[0]
            if short and short != tok:
                cands.append(short)
            # 「ALG&Associates神戸」「バディ神戸三宮」のような末尾の拠点地名を除いた形
            noplace = re.sub(r"(神戸|三宮|元町|姫路|明石|西宮|尼崎|芦屋|伊丹|宝塚|川西|三田|加古川|豊岡|洲本|淡路|篠山|北|中央|駅前)+$", "", tok)
            if noplace and noplace not in cands:
                cands.append(noplace)
    core = hq_core(name)
    if core and core not in cands:
        cands.append(core)
    return [NTA_ALIASES.get(c, c) for c in cands]


def check_hq(out: list[dict]) -> None:
    cache = _cache("nta2.json") or {}
    nta = NTA()
    keys = sorted({"|".join(hq_candidates(r["名称"])) for r in out if needs_hq_check(r["名称"])} - {""})
    todo = [k for k in keys if k not in cache]
    logger.info(f"法人番号で本店確認 {len(todo)} 件（キャッシュ {len(cache)}）")
    for i, key in enumerate(todo, 1):
        try:
            cands = key.split("|")
            want = {_nta_core(c) for c in cands}
            hits, seen = [], set()
            for c in cands:
                for query in ("弁護士法人" + c, c):
                    for x in nta.search(query):
                        if "弁護士法人" in x["name"] and _nta_core(x["name"]) in want and x["name"] + x["addr"] not in seen:
                            seen.add(x["name"] + x["addr"])
                            hits.append(x)
                    time.sleep(1)
                    if hits:
                        break
                if hits:
                    break
            cache[key] = hits
        except Exception as e:
            logger.warning(f"  法人番号検索失敗 {core}: {e}")
        time.sleep(1)
        if i % 20 == 0:
            _save_cache("nta2.json", cache)
            logger.info(f"  {i}/{len(todo)}")
    _save_cache("nta2.json", cache)
    for r in out:
        r["_hq"] = "対象外"
        if not needs_hq_check(r["名称"]):
            continue
        hits = cache.get("|".join(hq_candidates(r["名称"])), [])
        if not hits:
            r["_hq"] = "登録なし"
        elif any("兵庫県" in h["addr"] for h in hits):
            r["_hq"] = "兵庫県"
        else:
            r["_hq"] = "県外:" + hits[0]["addr"][:12]


def dedupe_by_phone(recs: list[dict]) -> list[dict]:
    """公式サイトで電話番号を補完した結果、同じ番号になった行（同一事務所）を1行にまとめる。"""
    groups: dict[str, list[dict]] = {}
    out = []
    for r in recs:
        d = _digits(r["電話番号"])
        if len(d) >= 10 and not d.startswith("050"):
            groups.setdefault(d, []).append(r)
        else:
            out.append(r)
    for g in groups.values():
        # 個人名の掲載（「弁護士 ○○(△△法律事務所)」）より事務所名の掲載を、bengo4 掲載を優先
        g.sort(key=lambda r: (bool(re.match(r"弁護士\s", r["名称"])), "bengo4" not in r["_src"], len(r["名称"])))
        best = dict(g[0])
        for other in g[1:]:
            for k in OUTPUT_COLS:
                if not best.get(k) and other.get(k):
                    best[k] = other[k]
            logger.info(f"  電話番号重複を統合: {best['名称']} ← {other['名称']}")
        out.append(best)
    return out


def run_merge() -> None:
    b4 = _cache("bengo4.json")
    gm = _cache("gmaps.json")
    if b4 is None or gm is None:
        logger.error("bengo4.json / gmaps.json がまだありません。先に bengo4 / gmaps を実行してください。")
        return
    logger.info(f"統合: bengo4 {len(b4)} 件 + gmaps {len(gm)} 件")
    hj = _cache("hojonavi.json") or []
    logger.info(f"補助金ナビ 弁護士 {len(hj)} 件")
    out = merge_records(b4, gm, hj)
    logger.info(f"統合後 {len(out)} 件")
    async def _both():
        # 公式サイト検索・取得と法人番号による本店確認を同時に進める
        await asyncio.gather(enrich(out), asyncio.to_thread(check_hq, out))

    asyncio.run(_both())
    excluded = [r for r in out if r["_hq"].startswith("県外")]
    for r in excluded:
        logger.info(f"  除外（本店県外）: {r['名称']} / {r['_hq']}")
    keep = dedupe_by_phone([r for r in out if not r["_hq"].startswith("県外")])
    keep.sort(key=lambda r: (r["所在地"] or "ん", r["名称"]))
    df = pd.DataFrame(keep)[OUTPUT_COLS]
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Sheet1")
        ws = w.sheets["Sheet1"]
        for col, width in zip("ABCDEFG", [40, 30, 40, 50, 16, 35, 45]):
            ws.column_dimensions[col].width = width
    _save_cache("merged_debug.json", out)
    logger.info(f"出力 {OUTPUT_FILE}: {len(keep)} 件（本店県外で除外 {len(excluded)} 件）")
    for c in OUTPUT_COLS:
        logger.info(f"  {c}: {(df[c].astype(str).str.len() > 0).sum()} 件")


if __name__ == "__main__":
    if MODE == "bengo4":
        run_bengo4()
    elif MODE == "gmaps":
        asyncio.run(run_gmaps())
    elif MODE == "merge":
        run_merge()
    elif MODE == "hojonavi":
        run_hojonavi()
    elif MODE == "gmaps_shard":
        asyncio.run(run_gmaps_shard(int(sys.argv[2]), int(sys.argv[3])))
    elif MODE == "gmaps_combine":
        combine_gmaps_shards()
    elif MODE == "presearch":
        # gmaps 収集中に bengo4 分の公式サイト検索・取得を先に進めてキャッシュしておく
        asyncio.run(enrich(merge_records(_cache("bengo4.json") or [], [])))
    else:
        print(__doc__)
