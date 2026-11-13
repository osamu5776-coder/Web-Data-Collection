"""
滋賀県 法律事務所 情報収集

使い方（収集フェーズは別プロセスで並列実行できる）:
  python scraper_shiga_bengoshi.py shigaben   # 滋賀弁護士会 会員一覧（全地域）＋各詳細
  python scraper_shiga_bengoshi.py bengo4     # 弁護士ドットコム 滋賀（弁護士一覧・事務所一覧）＋各詳細
  python scraper_shiga_bengoshi.py gmaps      # Google マップ 市町別 全件＋詳細
  python scraper_shiga_bengoshi.py merge      # 統合・公式サイト補完・Excel出力

データソース:
  - 滋賀弁護士会 会員一覧 https://shigaben.or.jp/introduce/member_list/ （全20地域・各詳細ページ）
    詳細ページに事務所名・所在地・電話番号・メールアドレス・ホームページがある。robots.txt なし。
  - 弁護士ドットコム https://www.bengo4.com/shiga/ （弁護士一覧）と /office/shiga/ （事務所一覧）
    robots.txt で禁止されている /*/contact/ は取得しない。
  - Google マップ「法律事務所 滋賀県<市町>」（各詳細ページ）

本店判定（本社＝主たる事務所が滋賀県のもののみ出力）:
  弁護士法人名・支店表記のある事務所は国税庁法人番号公表サイトで主たる事務所の所在地を確認し、
  県外のみに登録がある（県外本部の支店である）場合は除外する。個人事務所は所在地が本拠のため採用する。

出力列: 名称, メールアドレス, 公式サイトURL, 所在地, 電話番号, インスタURL, 問い合わせフォームURL
出力ファイル: 滋賀県_法律事務所リスト.xlsx
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
LOG_FILE = f"scraper_shiga_bengoshi_{MODE}.log"
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
OUTPUT_FILE = "滋賀県_法律事務所リスト.xlsx"
CACHE_DIR = os.path.join("data", "shiga_bengoshi")
UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36"
)
HEADERS = {"User-Agent": UA}
SHIGABEN_BASE = "https://shigaben.or.jp/introduce/member_list/"
SHIGABEN_LIST = SHIGABEN_BASE + "index.php?" + "&".join(f"area%5B%5D={i}" for i in range(1, 21))
BENGO4_LAWYER_BASE = "https://www.bengo4.com/shiga/"
BENGO4_OFFICE_BASE = "https://www.bengo4.com/office/shiga/"
NTA_BASE = "https://www.houjin-bangou.nta.go.jp"

AREAS = [
    "大津市", "彦根市", "長浜市", "近江八幡市", "草津市", "守山市", "栗東市", "甲賀市", "野洲市",
    "湖南市", "高島市", "東近江市", "米原市", "日野町", "竜王町", "愛荘町", "豊郷町", "甲良町", "多賀町",
]
# 未ログインの Google マップは1検索 120 件前後で打ち切られるため、事務所の多い地域は語を変えて重ねる
BIG_AREAS = {"大津市", "草津市", "彦根市", "長浜市", "近江八幡市", "守山市", "東近江市"}
GMAPS_QUERIES = [f"法律事務所 滋賀県{a}" for a in AREAS] + [f"弁護士事務所 滋賀県{a}" for a in AREAS if a in BIG_AREAS] \
    + [f"弁護士 滋賀県{a}" for a in AREAS if a in BIG_AREAS] + ["法律事務所 滋賀県"]
GMAPS_TABS = 3
GMAPS_SEARCH_TABS = 3

SKIP_DOMAINS = [
    "bengo4.com", "nichibenren.or.jp", "houterasu.or.jp", "shigaben.or.jp",
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
NON_LAWYER_RE = re.compile(r"司法書士|行政書士|税理士|社会保険労務士|社労士|土地家屋調査士|弁理士|特許|会計|公証|"
                           r"法テラス|司法支援センター|弁護士会|裁判所|検察|法務局|市役所|町役場|相談センター|大学|ロースクール")
SHIGA_CITY_RE = re.compile(r"^(大津|彦根|長浜|近江八幡|草津|守山|栗東|甲賀|野洲|湖南|高島|東近江|米原|蒲生郡|愛知郡|犬上郡)")


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
            # robots.txt が無いサイトは 404 の HTML を返すことがあるので text/plain のときだけ解釈する
            if resp.status_code < 400 and "html" not in resp.headers.get("Content-Type", ""):
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
            if r.encoding is None or r.encoding.lower() == "iso-8859-1":
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
        if d.startswith("077"):
            return f"{d[:3]}-{d[3:6]}-{d[6:]}"
        return f"{d[:4]}-{d[4:6]}-{d[6:]}"
    if len(d) == 11:
        return f"{d[:3]}-{d[3:7]}-{d[7:]}"
    return s


def _norm_addr(s: str) -> str:
    s = _nfkc(s)
    s = re.sub(r"^(住所|所在地)\s*[:：]?\s*", "", s)
    s = re.sub(r"^郵便番号\s*", "", s)
    s = re.sub(r"^日本、?\s*", "", s)
    s = re.sub(r"〒?\s*\d{3}-?\d{4}\s*", "", s).strip()
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"^滋賀県\s+", "滋賀県", s)
    if s and not s.startswith("滋賀県") and SHIGA_CITY_RE.match(s):
        s = "滋賀県" + s
    return s


# ── 滋賀弁護士会 会員一覧 ─────────────────────────────────

def _th_map(soup: BeautifulSoup) -> dict:
    f = {}
    for th in soup.find_all("th"):
        td = th.find_next_sibling("td")
        if td is not None:
            f.setdefault(th.get_text(strip=True), td)
    return f


def run_shigaben() -> None:
    cached = _cache("shigaben.json")
    if cached is not None:
        logger.info(f"滋賀弁護士会 キャッシュあり {len(cached)} 件")
        return
    r = _get(SHIGABEN_LIST, timeout=30)
    if r is None or r.status_code >= 400:
        logger.error("滋賀弁護士会 一覧取得失敗")
        return
    r.encoding = "utf-8"
    soup = BeautifulSoup(r.text, "lxml")
    rows = []
    for tr in soup.select("table tr"):
        tds = tr.find_all("td")
        if len(tds) < 5 or not (a := tds[0].find("a", href=True)):
            continue
        hp = tds[5].find("a", href=True) if len(tds) > 5 else None
        rows.append({
            "src": "shigaben", "detail": urljoin(SHIGABEN_BASE, a["href"]),
            "person": a.get_text(" ", strip=True).split(" ", 2)[-1] if a else "",
            "name": _nfkc(tds[2].get_text(" ", strip=True)),
            "addr": _norm_addr(tds[3].get_text(" ", strip=True)),
            "tel": _format_phone(tds[4].get_text(strip=True)),
            "web": hp["href"].strip() if hp else "",
        })
    logger.info(f"滋賀弁護士会 一覧 {len(rows)} 名 → 詳細取得")
    for i, x in enumerate(rows, 1):
        rr = _get(x["detail"])
        if rr is not None and rr.status_code < 400:
            rr.encoding = "utf-8"
            f = _th_map(BeautifulSoup(rr.text, "lxml"))
            if (td := f.get("氏名")):
                x["person"] = td.get_text("\n", strip=True).split("\n")[-1]
            if (td := f.get("事務所名")) and td.get_text(strip=True):
                x["name"] = _nfkc(td.get_text(" ", strip=True))
            if (td := f.get("所在地")) and td.get_text(strip=True):
                x["addr"] = _norm_addr(td.get_text(" ", strip=True))
            if (td := f.get("電話番号")) and td.get_text(strip=True):
                x["tel"] = _format_phone(td.get_text(strip=True))
            if (td := f.get("メールアドレス")):
                m = re.search(r"[\w.%+\-]+@[\w.\-]+\.[a-zA-Z]{2,}", _nfkc(td.get_text(" ", strip=True)))
                x["email"] = m.group(0) if m else ""
            if (td := f.get("ホームページ")) and (a := td.find("a", href=True)):
                x["web"] = a["href"].strip()
            if (td := f.get("地域")):
                x["area"] = td.get_text(strip=True)
        if i % 20 == 0:
            logger.info(f"  詳細 {i}/{len(rows)}")
        time.sleep(random.uniform(0.8, 1.5))
    _save_cache("shigaben.json", rows)
    logger.info(f"滋賀弁護士会 完了 {len(rows)} 名（メール {sum(1 for x in rows if x.get('email'))}）")


# ── 弁護士ドットコム ─────────────────────────────────────

def _parse_bengo4_office_list(html: str, base: str) -> list[dict]:
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


def _parse_bengo4_lawyer_list(html: str, base: str) -> list[dict]:
    soup = BeautifulSoup(html, "lxml")
    out = []
    for c in soup.select(".lawyer-cassette__root"):
        a = c.select_one("a.lawyer-cassette-profile__nameLink")
        o = c.select_one(".lawyer-cassette-profile__officeName")
        if not a or not o:
            continue
        m = re.search(r"0\d{1,4}-\d{1,4}-\d{3,4}", c.get_text(" ", strip=True))
        out.append({"src": "bengo4", "name": o.get_text(strip=True), "person": a.get_text(strip=True),
                    "detail": urljoin(base, a["href"]), "tel": m.group(0) if m else ""})
    return out


def _parse_bengo4_detail(html: str) -> dict:
    soup = BeautifulSoup(html, "lxml")
    res = {"addr": "", "web": ""}
    for dt in soup.find_all("dt"):
        label = dt.get_text(strip=True)
        dd = dt.find_next_sibling("dd")
        if not dd:
            continue
        if label == "所在地" and not res["addr"]:
            res["addr"] = _norm_addr(dd.get_text(" ", strip=True))
        elif ("事務所URL" in label or label == "ホームページ") and (a := dd.find("a", href=True)):
            href = a["href"].strip()
            if href.startswith("http") and not any(d in href for d in SKIP_DOMAINS):
                res["web"] = href
    return res


def _bengo4_pages(base: str, parser) -> list[dict]:
    r = _get(base)
    if r is None:
        logger.error(f"bengo4 一覧取得失敗 {base}")
        return []
    pages = max([int(x) for x in re.findall(r"page=(\d+)", r.text)] or [1])
    recs = parser(r.text, base)
    logger.info(f"bengo4 {base}: {pages} ページ")
    for p in range(2, pages + 1):
        time.sleep(random.uniform(1.0, 2.0))
        url = f"{base}?page={p}"
        rr = _get(url)
        got = parser(rr.text, url) if rr is not None else []
        recs.extend(got)
        logger.info(f"  page {p}/{pages}: {len(got)} 件（累計 {len(recs)}）")
    return recs


def run_bengo4() -> None:
    cached = _cache("bengo4.json")
    if cached is not None:
        logger.info(f"bengo4 キャッシュあり {len(cached)} 件")
        return
    recs = _bengo4_pages(BENGO4_OFFICE_BASE, _parse_bengo4_office_list) \
        + _bengo4_pages(BENGO4_LAWYER_BASE, _parse_bengo4_lawyer_list)
    seen, uniq = set(), []
    for x in recs:
        if x["detail"] not in seen:
            seen.add(x["detail"])
            uniq.append(x)
    logger.info(f"bengo4 詳細取得 {len(uniq)} 件")
    for i, x in enumerate(uniq, 1):
        if "/contact/" not in x["detail"] and _is_allowed(x["detail"]):
            rr = _get(x["detail"])
            if rr is not None and rr.status_code < 400:
                x.update(_parse_bengo4_detail(rr.text))
        if i % 20 == 0:
            logger.info(f"  詳細 {i}/{len(uniq)}")
        time.sleep(random.uniform(0.8, 1.5))
    _save_cache("bengo4.json", uniq)
    logger.info(f"bengo4 完了 {len(uniq)} 件")


# ── Google マップ ──────────────────────────────────────────

async def _gmaps_detail(tab, href: str, name: str) -> dict | None:
    rec = {"src": "gmaps", "name": name, "addr": "", "tel": "", "web": "", "category": "", "detail": href}
    await tab.goto(href + "?hl=ja", wait_until="domcontentloaded", timeout=60000)
    try:
        # 描画が遅れることがあるため、住所か電話のボタンが出るまで待つ
        await tab.wait_for_selector('button[data-item-id="address"], button[data-item-id^="phone:tel:"]', timeout=20000)
    except Exception:
        pass
    await tab.wait_for_timeout(1200)
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
    return rec if (rec["addr"] or rec["tel"] or rec.get("closed")) else None


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

        targets = {h: n for h, n in links.items() if not (NON_LAWYER_RE.search(n or "") and "法律事務所" not in (n or ""))}
        details: dict[str, dict] = _cache("gmaps_details.json") or {}
        todo = [h for h in targets if h not in details]
        logger.info(f"gmaps 詳細取得 {len(todo)} 件（全 {len(targets)} / 取得済 {len(details)}）")
        dsem = asyncio.Semaphore(GMAPS_TABS)
        cnt = {"n": 0}

        async def _detail(href: str, name: str) -> None:
            async with dsem:
                tab = await ctx.new_page()
                try:
                    rec = await _gmaps_detail(tab, href, name)
                    if rec:
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


# ── 統合 ─────────────────────────────────────────────────

def _key_name(name: str) -> str:
    s = _nfkc(name).lower()
    s = re.sub(r"[\s・･\-‐ー－（）()「」]", "", s)
    s = s.replace("弁護士法人", "").replace("法律事務所", "").replace("弁護士事務所", "")
    return s


def _city(addr: str) -> str:
    m = re.search(r"滋賀県\s*([^\s市]{1,4}市|[^\s]{1,3}郡[^\s]{1,3}?町)", addr)
    return m.group(1) if m else ""


def _key_addr(addr: str) -> str:
    s = _nfkc(addr).replace("丁目", "-").replace("番地", "-").replace("番", "-").replace("号", "")
    s = re.sub(r"[\s‐－−ー―]", "-", s)
    s = re.sub(r"-+", "-", s)
    s = re.sub(r"[\s-]", "", s)
    # 建物名・階数より前（番地まで）で比較する
    m = re.match(r"(.*?\d+(?:-\d+)*)", s)
    return m.group(1) if m else s


def _is_lawyer(rec: dict) -> bool:
    name, cat = rec.get("name", ""), rec.get("category", "")
    # 企業の法務部門（インハウス）や法テラスは対象外
    if re.search(r"株式会社|法テラス|司法支援センター", name) and not name.startswith("弁護士法人"):
        return False
    if rec.get("src") in ("bengo4", "shigaben"):
        return True
    # 「〇〇法務事務所」「〇〇法務グループ」は司法書士・行政書士のことが多いので除く
    if "法務" in name and not re.search(r"法律事務所|弁護士", name):
        return False
    if rec.get("closed"):
        return False
    if NON_LAWYER_RE.search(name) and "法律事務所" not in name:
        return False
    if re.search(r"司法書士|行政書士|税理士|探偵|保険|デイサービス|介護", name):
        return False
    if re.search(r"弁護士|法律事務所", cat):
        return True
    return bool(re.search(r"法律事務所|弁護士", name))


def merge_records(recs_in: list[dict]) -> list[dict]:
    recs = [r for r in recs_in if _is_lawyer(r)]
    recs = [r for r in recs if "滋賀県" in r.get("addr", "") or (r["src"] == "bengo4" and not r.get("addr"))]
    parent = list(range(len(recs)))

    def find(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    def union(i, j):
        parent[find(i)] = find(j)

    by_phone, by_name, by_addr = {}, {}, {}
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
        # 同じ番地で名称の核が一致するもの（表記ゆれ）もまとめる
        if r.get("addr"):
            ka = (_key_addr(r["addr"]), _key_name(r["name"])[:3])
            if ka in by_addr:
                union(i, by_addr[ka])
            else:
                by_addr[ka] = i
    groups: dict[int, list[dict]] = {}
    for i, r in enumerate(recs):
        groups.setdefault(find(i), []).append(r)

    out = []
    for g in groups.values():
        s = next((r for r in g if r["src"] == "shigaben"), None)
        b = next((r for r in g if r["src"] == "bengo4"), None)
        m = next((r for r in g if r["src"] == "gmaps"), None)
        name = (s or b or m)["name"]
        tels = [r.get("tel", "") for r in (([s] if s else []) + g) if r.get("tel")]
        tel = next((t for t in tels if not _digits(t).startswith("050")), tels[0] if tels else "")
        addr = (s or {}).get("addr") or (b or {}).get("addr") or (m or {}).get("addr", "")
        web = next((r.get("web") for r in g if r.get("web") and not any(d in r["web"] for d in SKIP_DOMAINS)), "")
        email = next((r["email"] for r in g if r.get("email")), "")
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


BAD_EMAIL_RE = re.compile(r"\.(png|jpg|jpeg|gif|svg|webp)$|example\.|sentry|wixpress|@domain|your-?mail|xxx|^sample@|^test@|@mail\.com$|yamadahp", re.I)


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
    m = re.search(r"0(?:77\d?|74\d)[-‐(（ ]\d{1,4}[-‐)） ]\d{4}", _nfkc(text))
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
    city = _city(rec["所在地"]) or "滋賀県"
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
            if k == "メールアドレス" and BAD_EMAIL_RE.search(v or ""):
                continue
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
            noplace = re.sub(r"(滋賀|大津|膳所|石山|瀬田|草津|南草津|守山|栗東|野洲|彦根|長浜|近江八幡|東近江|八日市|甲賀|水口|高島|米原|湖南|びわ湖|琵琶湖|駅前|中央)+$", "", tok)
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
            logger.warning(f"  法人番号検索失敗 {key}: {e}")
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
        elif any("滋賀県" in h["addr"] for h in hits):
            r["_hq"] = "滋賀県"
        else:
            r["_hq"] = "県外:" + hits[0]["addr"][:12]


def _addr_sig(addr: str):
    """市町＋番地・部屋番号の数字列（建物名の表記ゆれを無視して同一所在地を判定する）。"""
    nums = re.findall(r"\d+", _nfkc(addr))
    return (_city(addr), tuple(nums[:4])) if len(nums) >= 4 else None


def drop_branches(keep: list[dict], excluded: list[dict], hq_cache: dict) -> list[dict]:
    """本店県外の支店と同じ所在地の別表記（Google マップ等）と、県内本店法人の従たる事務所を除く。"""
    ex_sigs = {_addr_sig(r["所在地"]) for r in excluded} - {None}
    out = []
    for r in keep:
        # 弁護士会の名簿に独立の事務所として載っているものは、同じ部屋番号でも別事務所として残す
        if "shigaben" not in r["_src"] and _addr_sig(r["所在地"]) in ex_sigs:
            logger.info(f"  除外（本店県外の支店と同所在地）: {r['名称']} / {r['所在地']}")
            continue
        if r["_hq"] == "滋賀県":
            hits = hq_cache.get("|".join(hq_candidates(r["名称"])), [])
            hq_city = next((_city(_nfkc(h["addr"])) for h in hits if "滋賀県" in h["addr"]), "")
            if hq_city and _city(r["所在地"]) and _city(r["所在地"]) != hq_city:
                logger.info(f"  除外（県内本店法人の従たる事務所）: {r['名称']} / {r['所在地']}（本店 {hq_city}）")
                continue
        out.append(r)
    return out


def dedupe_by_email(recs: list[dict]) -> list[dict]:
    """同じ事務所ドメインのメールを持つ行（移転前後の重複掲載）を1行にまとめる。共用ドメインは対象外。"""
    seen: dict[str, dict] = {}
    out = []
    for r in sorted(recs, key=lambda r: "shigaben" not in r["_src"]):
        e = r["メールアドレス"].lower()
        if e and not re.search(r"@(shigaben\.or\.jp|gmail\.com|yahoo\.co\.jp|.*nifty\.com|.*ocn\.ne\.jp)$", e):
            if e in seen:
                logger.info(f"  メール重複を統合: {seen[e]['名称']} ← {r['名称']}")
                for k in OUTPUT_COLS:
                    if not seen[e].get(k) and r.get(k):
                        seen[e][k] = r[k]
                continue
            seen[e] = r
        out.append(r)
    return out


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
        # 弁護士会掲載の事務所名を優先
        g.sort(key=lambda r: ("shigaben" not in r["_src"], "bengo4" not in r["_src"], len(r["名称"])))
        best = dict(g[0])
        for other in g[1:]:
            for k in OUTPUT_COLS:
                if not best.get(k) and other.get(k):
                    best[k] = other[k]
            logger.info(f"  電話番号重複を統合: {best['名称']} ← {other['名称']}")
        out.append(best)
    return out


def run_merge() -> None:
    sb, b4, gm = _cache("shigaben.json"), _cache("bengo4.json"), _cache("gmaps.json")
    if sb is None or b4 is None or gm is None:
        logger.error("shigaben.json / bengo4.json / gmaps.json がまだありません。先に各収集を実行してください。")
        return
    logger.info(f"統合: 滋賀弁護士会 {len(sb)} 名 + bengo4 {len(b4)} 件 + gmaps {len(gm)} 件")
    out = merge_records(sb + b4 + gm)
    logger.info(f"統合後 {len(out)} 件")
    async def _both():
        # 公式サイト検索・取得と法人番号による本店確認を同時に進める
        await asyncio.gather(enrich(out), asyncio.to_thread(check_hq, out))

    asyncio.run(_both())
    excluded = [r for r in out if r["_hq"].startswith("県外")]
    for r in excluded:
        logger.info(f"  除外（本店県外）: {r['名称']} / {r['_hq']}")
    keep = dedupe_by_phone([r for r in out if not r["_hq"].startswith("県外")])
    keep = dedupe_by_email(drop_branches(keep, excluded, _cache("nta2.json") or {}))
    keep.sort(key=lambda r: (r["所在地"] or "ん", r["名称"]))
    df = pd.DataFrame(keep)[OUTPUT_COLS]
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Sheet1")
        ws = w.sheets["Sheet1"]
        for col, width in zip("ABCDEFG", [40, 30, 40, 50, 16, 35, 45]):
            ws.column_dimensions[col].width = width
    _save_cache("merged_debug.json", keep)
    logger.info(f"出力 {OUTPUT_FILE}: {len(keep)} 件（本店県外で除外 {len(excluded)} 件）")
    for c in OUTPUT_COLS:
        logger.info(f"  {c}: {(df[c].astype(str).str.len() > 0).sum()} 件")


if __name__ == "__main__":
    if MODE == "shigaben":
        run_shigaben()
    elif MODE == "bengo4":
        run_bengo4()
    elif MODE == "gmaps":
        asyncio.run(run_gmaps())
    elif MODE == "merge":
        run_merge()
    elif MODE == "presearch":
        # gmaps 収集中に弁護士会・bengo4 分の公式サイト検索・取得を先に進めてキャッシュしておく
        asyncio.run(enrich(merge_records((_cache("shigaben.json") or []) + (_cache("bengo4.json") or []))))
    else:
        print(__doc__)
