"""
栃木県 WEB制作会社 情報収集

データソース（seed_tochigi_web.json に事前収集して同梱）:
  - n-works.link のブログ記事「栃木県のWeb制作会社」
    （名称・所在地・電話番号・公式サイトURL が掲載）
  - wk-partners.co.jp「ホームページ制作会社辞典」栃木県ページ（全3ページ）
    （名称・所在地のみ）
  - imitsu.jp（PRONIアイミツ）「栃木県のホームページ制作会社」一覧（全6ページ・110件）
    （名称・所在地のみ）

処理:
  Phase 1 - 3ソースを会社名（法人格・支店表記を除いた主要部）で統合・重複排除。
            電話番号・URLが判明している方（主に n-works）を優先して残す。
  Phase 2 - 既知の公式サイトURLから メール・インスタ・問い合わせフォームURL・
            （未判明なら）電話番号 を取得する。
  Phase 3 - 公式サイトURLが未判明の会社について DuckDuckGo で公式サイトを検索し、
            電話番号一致または会社名一致で検証する。
  Phase 4 - 電話番号は市外局番の桁数（028=3桁 / 028x・029=4桁）を、
            数字だけでは判別できないため所在地の市区町村から判定してハイフン整形する。

出力列: 名称, メールアドレス, 公式サイトURL, 所在地, 電話番号, インスタURL, 問い合わせフォームURL
出力ファイル: 栃木県_WEB制作会社リスト.xlsx
"""

import asyncio
import io
import json
import logging
import random
import re
import sys
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

LOG_FILE = "scraper_tochigi_web.log"
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
OUTPUT_FILE = "栃木県_WEB制作会社リスト.xlsx"
SEED_JSON = "seed_tochigi_web.json"
PREF_NAME = "栃木県"
UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)
HEADERS = {"User-Agent": UA}

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

SKIP_DOMAINS = [
    "navitime.co.jp", "baseconnect.in", "itp.ne.jp", "mapion.co.jp",
    "mapfan.com", "goo.ne.jp", "wikipedia.org", "google.com", "google.co.jp",
    "maps.google.com", "maps.google.co.jp", "bing.com", "yahoo.co.jp",
    "yahoo.com", "mynavi.jp", "townpage.goo.ne.jp", "e-shops.jp", "navi-cat.jp",
    "jp.indeed.com", "indeed.com", "jp.stanby.com", "stanby.com", "job-medley.com",
    "hellowork.mhlw.go.jp", "houjin.jp", "houjin-bangou.nta.go.jp", "alarmbox.jp",
    "ameblo.jp", "seesaa.net", "ldblog.jp", "livedoor.jp", "hatenablog.com",
    "blogspot.com", "exblog.jp", "note.com", "wantedly.com", "en-hyouban.com",
    "baitoru.com", "townwork.net", "hellowork.careers", "jobtag.mhlw.go.jp",
    "instagram.com", "facebook.com", "twitter.com", "x.com", "prtimes.jp",
    "youtube.com", "tiktok.com", "line.me", "lin.ee", "nikkei.com",
    "tabelog.com", "retty.me", "hotpepper.jp", "jpnumber.com", "telnavi.jp",
    "csi-japan.com", "salesnow.jp", "musubu.in", "ipros.jp", "ekiten.jp",
    # 検証で誤検出したディレクトリ・名鑑・企業DB・行政/新聞・電話帳等
    "yomiuri.co.jp", "isico.or.jp", "24u.jp", "everytown.info", "yapy.jp",
    "alp-grp.jp", "agencyhub.jp", "companydata.tsujigawa.com", "tsujigawa.com",
    "kurobe-unazuki.jp", "town.shimanto.lg.jp", "shimanto.lg.jp", "g-ara.jp",
    "seino.co.jp", "fitness-aim.com", "jobnavi-i.jp", "info.gbiz.go.jp",
    "gbiz.go.jp", "baseconnect", "mapoo.jp", "navi-city.com", "jタウン",
    "chunichi.co.jp", "job.mynavi.jp", "rikunabi.com", "doda.jp",
    "nikkei.co.jp", "atengineer.com", "kaisyaneta.com", "urx.jp",
    "cyzo.com", "wwwc.jp", "cb-sokuho.com", "goo.gl", "bit.ly",
    "j-net21.smrj.go.jp", "smrj.go.jp", "ma-bank.jp", "ullet.com",
    "nikkyo.or.jp", "recme.jp", "job-terminal.com", "hellowork-navi.com",
    "town-guide.net", "jobofferad-agency.net", "syn-ad.com", "n-works.link",
    "sankeiliving.co.jp", "hokutetsukoku.jp", "themedia.jp", "web.fc2.com",
    "navita.co.jp", "imitsu.jp", "xn--vcki1fxhx94nwsb.com",
    # Web制作会社の比較・発注・まとめ系ディレクトリ
    "web-kanji.com", "hnavi.co.jp", "pronavi", "hp-mikata.jp", "mikata-web",
    "creive.me", "web-planner.jp", "webseisaku-madoguchi.com", "wixsite.com",
    "webnavi.biz", "seizo-navi.com", "hp-seisaku.net", "crowdworks.jp",
    "lancers.jp", "coconala.com", "web-consultants.jp", "map.yahoo.co.jp",
    "value-domain.com", "onamae.com", "jimdo", "amebaownd.com", "peraichi.com",
    "goope.jp", "shopify.com", "hubspot", "studio.design", "studio.site",
    "naviibaraki.com", "tsukuba-cci.or.jp", "refowork.com", "presspage.biz",
    "job-sign.com", "rs-hokkaido.net", "houjin.goo.to", "goo.to",
    "homepage.work", "astlink.jp", "nissenmedix.co.jp", "pc-rs.net",
    "houjinbase.com", "houjin.jp", "baseconnect.in",
    # 今回の元データ自体（会社の公式サイトではない）
    "n-works.link", "wk-partners.co.jp", "imitsu.jp",
    # 検証したが誤検出と判明したドメイン（同名の栃木県外の別会社／観光ポータル等）
    "tochinavi.net", "weblinks.jp", "gnx.co.jp", "r-suntec.co.jp",
    # 制作会社比較・企業データ・地域ショップ紹介ポータル(公式サイトではない)
    "homepage-kimete.com", "metoree.com", "tochigisi.com",
]

PORTAL_TITLE_KWS_EXTRA = ["比較", "おすすめ", "選び方", "発注", "ランキング", "一覧【"]

SKIP_URL_PATTERNS = [
    r"\.pdf(\?|$)",
    r"/company/\d+",
]


# ── 共通ユーティリティ ────────────────────────────────────
_robots_cache: dict[str, RobotFileParser] = {}


def _is_allowed(url: str) -> bool:
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
    try:
        return _robots_cache[origin].can_fetch(UA, url)
    except Exception:
        return True


def _normalize_name(s: str) -> str:
    s = re.sub(r"[\s　]+", "", str(s))
    s = s.translate(ZEN2HAN)
    s = s.replace("(株)", "株式会社").replace("（株）", "株式会社")
    s = s.replace("(有)", "有限会社").replace("（有）", "有限会社")
    return s.lower()


def _name_core(name: str) -> str:
    core = re.sub(r"(株式会社|有限会社|合同会社|合資会社|特定非営利活動法人|NPO法人)", "", str(name))
    core = re.sub(r"(栃木|北関東|関東|宇都宮|小山|栃木市|佐野|足利)\s*(支社|支店|営業所|支局|営業部)$", "", core)
    core = re.sub(r"[（(].*?[）)]", "", core)  # 「（DEN）」「(ファーイーストデザイン)」等の読み仮名注記を除去
    return _normalize_name(core)


def _normalize_phone(s: str) -> str:
    return re.sub(r"[^0-9]", "", str(s))


# 栃木県で使用される市外局番。028(3桁) と 028x(4桁) は数字だけでは
# 判別できない（例: "0282XXXXXX" は 028-2XX-XXXX か 0282-XX-XXXX か）。
# そのため所在地の市区町村名から正しい市外局番を判定する。
MOBILE_PREFIX = ("070", "080", "090", "050")
FREE_PREFIX_4 = ("0120", "0800", "0570")
TOCHIGI_AREA_4 = ("0282", "0283", "0284", "0285", "0287", "0288", "0289")

# 市区町村 → 市外局番（3桁 "028" または 4桁）
CITY_AREA: dict[str, str] = {}
for _cities, _code in (
    (["宇都宮", "上三川", "高根沢", "芳賀", "益子", "茂木", "市貝", "塩谷"], "028"),
    (["栃木市", "壬生"], "0282"),
    (["佐野"], "0283"),
    (["足利"], "0284"),
    (["小山", "下野", "野木", "真岡"], "0285"),
    (["那須塩原", "大田原", "矢板", "那須郡那須町", "那須町", "那珂川", "那須烏山", "さくら"], "0287"),
    (["日光"], "0288"),
    (["鹿沼"], "0289"),
):
    for _c in _cities:
        CITY_AREA[_c] = _code


def _area_from_address(address: str) -> str | None:
    for city, code in CITY_AREA.items():
        if city in address:
            return code
    return None


def format_phone(raw: str, address: str = "") -> str:
    """電話番号を市外局番でハイフン整形する。10桁の固定電話は所在地の
    市区町村から市外局番の桁数（028=3桁 / 028x=4桁）を判定する。

    既に「市外局番-局番-加入者番号」の3ブロックにハイフン等で区切られて
    いる場合（人手で書かれた n-works の記事や、企業サイト本文中の表記）は、
    その区切りが正しいとみなしてそのまま採用する。数字だけの表記
    （NAVITIME 等）の場合のみ、市外局番の桁数を所在地から推定して分割する。
    """
    pre = str(raw).translate(ZEN2HAN)
    pre = re.sub(r"[()（）\s]", "-", pre)
    parts = [re.sub(r"[^0-9]", "", p) for p in re.split(r"-+", pre) if p]
    parts = [p for p in parts if p]
    if len(parts) == 3 and all(part.isdigit() and part for part in parts):
        joined = "".join(parts)
        if joined.startswith("000") or len(set(joined)) == 1:
            return ""  # "000-XXX-XXXX" 等の無効値・掲載用ダミー番号を除去
        return "-".join(parts)

    d = _normalize_phone(raw)
    if not d:
        return ""
    if d.startswith("81") and len(d) >= 11:
        d = "0" + d[2:]
    if len(d) == 10:
        if d.startswith(FREE_PREFIX_4):
            return f"{d[:4]}-{d[4:7]}-{d[7:]}"
        if d[:2] in ("03", "06"):
            return f"{d[:2]}-{d[2:6]}-{d[6:]}"
        area = _area_from_address(address) if d.startswith("028") else None
        if area == "028":
            return f"{d[:3]}-{d[3:6]}-{d[6:]}"
        if area and d.startswith(area):
            return f"{d[:4]}-{d[4:6]}-{d[6:]}"
        # 所在地から判定できない場合は数字パターンで推定
        if d.startswith(TOCHIGI_AREA_4):
            return f"{d[:4]}-{d[4:6]}-{d[6:]}"
        if d.startswith("028"):
            return f"{d[:3]}-{d[3:6]}-{d[6:]}"
        # 栃木県外の番号（HQ番号等）はそのまま4-2-4想定で返す
        return f"{d[:4]}-{d[4:6]}-{d[6:]}"
    if len(d) == 11:
        if d.startswith(MOBILE_PREFIX):
            return f"{d[:3]}-{d[3:7]}-{d[7:]}"
        if d.startswith(FREE_PREFIX_4):
            return f"{d[:4]}-{d[4:7]}-{d[7:]}"
        return f"{d[:3]}-{d[3:7]}-{d[7:]}"
    if len(d) == 9 and d.startswith("0"):
        # 稀に下1桁欠け等。そのまま2-3-4で割る
        return f"{d[:2]}-{d[2:5]}-{d[5:]}"
    return d


def clean_address(addr: str) -> str:
    addr = re.sub(r"[\s　]+", " ", str(addr)).strip()
    addr = re.sub(r"^〒?\s*\d{3}-?\d{4}\s*", "", addr).strip()
    addr = addr.translate(ZEN2HAN)
    # 文字化けした長音・ハイフン（例: "53?1", "24ー30"）を統一
    addr = addr.replace("�", "-").replace("?", "-")
    addr = re.sub(r"(?<=\d)[ー−―‐](?=\d)", "-", addr)
    if addr and PREF_NAME not in addr:
        addr = PREF_NAME + addr
    return addr


# ── Phase 1: 3ソース（n-works / wk-partners / imitsu）を統合 ──

def load_and_merge_seed() -> list[dict]:
    logger.info("Phase 1: seed_tochigi_web.json（n-works/wk-partners/imitsu）を統合")
    with open(SEED_JSON, encoding="utf-8") as f:
        rows = json.load(f)
    logger.info(f"  読み込み: {len(rows)} 行（3ソース合算・重複含む）")

    merged: dict[str, dict] = {}
    order: list[str] = []
    for row in rows:
        name = row["name"].strip()
        ck = _name_core(name)
        if not ck:
            continue
        addr = clean_address(row.get("addr", ""))
        phone = format_phone(row.get("phone", ""), addr) if row.get("phone") else ""
        url = row.get("url", "")
        if ck not in merged:
            merged[ck] = {k: "" for k in OUTPUT_COLS}
            merged[ck]["名称"] = name
            order.append(ck)
        rec = merged[ck]
        # より詳細な住所（番地まで含む・文字数が長い方）を採用
        if addr and (not rec["所在地"] or len(addr) > len(rec["所在地"])):
            rec["所在地"] = addr
        if phone and not rec["電話番号"]:
            rec["電話番号"] = phone
        if url and not rec["公式サイトURL"] and not _is_skip_domain(url):
            rec["公式サイトURL"] = url

    records = [merged[ck] for ck in order]
    logger.info(f"Phase 1 完了: 重複排除後 {len(records)} 社"
                f"（電話番号あり {sum(1 for r in records if r['電話番号'])} /"
                f" 公式サイトURLあり {sum(1 for r in records if r['公式サイトURL'])}）")
    return records


# ── Phase 3: DuckDuckGo 公式サイト検索・検証 ──────────────

def _is_skip_domain(url: str) -> bool:
    if any(d in url for d in SKIP_DOMAINS):
        return True
    return any(re.search(p, url, re.I) for p in SKIP_URL_PATTERNS)


PORTAL_TITLE_KWS = [
    "求人", "アルバイト", "転職", "口コミ", "クチコミ", "ランキング",
    "一覧｜", "一覧 |", "検索｜", "検索 |", "を探す", "まとめ", "電話番号検索",
    "企業情報", "会社概要・沿革", "の情報", "企業データ", "法人番号", "の詳細",
    "スポンサー", "協賛", "ふるさと納税", "市役所", "町役場", "議会",
] + PORTAL_TITLE_KWS_EXTRA


def _page_title(html: str) -> str:
    m = re.search(r"<title[^>]*>(.*?)</title>", html, re.I | re.S)
    return re.sub(r"\s+", " ", m.group(1)).strip() if m else ""


def _is_portal_page(html: str) -> bool:
    title = _page_title(html)
    return any(kw in title for kw in PORTAL_TITLE_KWS)


def _verify_match(url: str, html_text: str, name: str, phone: str) -> bool:
    """検索で得た候補URLが本当にその会社の公式サイトかを厳しめに判定する。

    採用条件（いずれか）:
      A. ページ内に当該電話番号（数字列）が含まれる
      B. <title> に会社名の主要部（法人格・支店表記を除いた語）が含まれ、
         かつページが会社サイトらしい（会社概要/company 等の語を含む）
      C. ドメイン名（登録部分）に会社名由来のローマ字が含まれる
    """
    core = _name_core(name)
    if len(core) < 2:
        return False

    digits_page = re.sub(r"[^0-9]", "", html_text)
    pd_ = _normalize_phone(phone)
    if pd_ and len(pd_) >= 9 and pd_ in digits_page:
        return True

    title_norm = _normalize_name(_page_title(html_text))
    looks_corp = any(k in html_text for k in (
        "会社概要", "会社案内", "企業情報", "事業内容", "COMPANY", "company",
        "About", "ABOUT", "アクセス", "お問い合わせ", "プライバシーポリシー",
    ))
    if core in title_norm and looks_corp:
        return True

    host = urlparse(url).netloc.lower()
    reg = host.split(":")[0].removeprefix("www.")
    reg_main = reg.split(".")[0]
    # 会社名のローマ字化は難しいので、英字社名の場合のみドメイン一致を見る
    ascii_core = re.sub(r"[^a-z0-9]", "", _normalize_name(name))
    if len(reg_main) >= 4 and len(ascii_core) >= 4 and (
        reg_main in ascii_core or ascii_core in reg_main
    ):
        return True
    return False


def search_official_site(name: str, address: str, phone: str) -> tuple[str, str]:
    city_m = re.search(rf"{PREF_NAME}\s*([^\s0-9]{{2,8}}?[市区町村郡])", address)
    city = city_m.group(1) if city_m else PREF_NAME
    query = f"{name} {city} ホームページ制作 公式サイト"
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
        if _verify_match(resp.url, resp.text, name, phone):
            return resp.url, resp.text
    return "", ""


# ── Phase 4: 公式サイトから追加情報 ──────────────────────

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
    text = soup.get_text("\n")
    base_host = urlparse(url).netloc

    cf = soup.find(attrs={"data-cfemail": True})
    if cf:
        dec = _decode_cfemail(cf["data-cfemail"])
        if dec:
            result["メールアドレス"] = dec
    if not result.get("メールアドレス"):
        emails = re.findall(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}", text)
        emails = [e for e in emails if not re.search(r"\.(png|jpg|jpeg|gif|svg|webp)$", e, re.I)]
        emails = [e for e in emails if not e.lower().startswith(("example@", "info@example"))]
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
                   "ご相談", "メールフォーム", "お問合せ", "問合せ"]
    for a in soup.find_all("a", href=True):
        href = a["href"]
        ltext = a.get_text(strip=True)
        if not any(k in href.lower() or k in ltext for k in contact_kws):
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

    m = re.search(r"0\d{1,3}[-(]\d{1,4}[-)]\d{3,4}", text)
    if m:
        result["_site_phone"] = m.group(0)
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
    sem = asyncio.Semaphore(sem_count)
    done = {"n": 0}

    async def _one(rec: dict) -> None:
        async with sem:
            extras = await asyncio.to_thread(_fetch_official_sync, rec["公式サイトURL"])
            sp = extras.pop("_site_phone", "")
            if sp and not rec.get("電話番号"):
                rec["電話番号"] = format_phone(sp, rec.get("所在地", ""))
            for k, v in extras.items():
                if v and not rec.get(k):
                    rec[k] = v
            done["n"] += 1
            if done["n"] % 15 == 0 or done["n"] == len(targets):
                logger.info(f"  既知URL情報取得: {done['n']}/{len(targets)}")

    await asyncio.gather(*[_one(r) for r in targets])


async def enrich_search_records(records: list[dict], sem_count: int = 3) -> None:
    targets = [r for r in records if not r.get("公式サイトURL")]
    sem = asyncio.Semaphore(sem_count)
    done = {"n": 0, "found": 0}

    async def _one(rec: dict) -> None:
        async with sem:
            try:
                url, html = await asyncio.wait_for(
                    asyncio.to_thread(
                        search_official_site, rec["名称"], rec.get("所在地", ""), rec.get("電話番号", "")
                    ),
                    timeout=50,
                )
            except asyncio.TimeoutError:
                url, html = "", ""
            if url:
                rec["公式サイトURL"] = url
                extras = _parse_official_html(url, html)
                sp = extras.pop("_site_phone", "")
                if sp and not rec.get("電話番号"):
                    rec["電話番号"] = format_phone(sp, rec.get("所在地", ""))
                for k, v in extras.items():
                    if v and not rec.get(k):
                        rec[k] = v
                done["found"] += 1
            done["n"] += 1
            if done["n"] % 10 == 0 or done["n"] == len(targets):
                logger.info(f"  検索・検証: {done['n']}/{len(targets)} (発見 {done['found']})")
            await asyncio.sleep(random.uniform(1.0, 2.0))

    await asyncio.gather(*[_one(r) for r in targets])


# ── メイン ────────────────────────────────────────────────

async def main() -> None:
    logger.info("=" * 65)
    logger.info("栃木県 WEB制作会社 収集開始")
    logger.info("=" * 65)
    start = datetime.now()

    deduped = load_and_merge_seed()

    logger.info("Phase 2: 既知の公式サイトから追加情報を取得")
    await enrich_known_urls(deduped, sem_count=6)

    logger.info("Phase 3: 未判明の会社を検索・検証・情報取得")
    await enrich_search_records(deduped, sem_count=3)

    # 同一ドメインが複数社に一致 → ポータル誤採用として除去
    dom_count: dict[str, int] = {}
    for r in deduped:
        if r.get("公式サイトURL"):
            d = urlparse(r["公式サイトURL"]).netloc
            dom_count[d] = dom_count.get(d, 0) + 1
    portal = {d for d, c in dom_count.items() if c > 1}
    if portal:
        logger.info(f"  ポータル判定で除去: {sorted(portal)}")
        for r in deduped:
            if r.get("公式サイトURL") and urlparse(r["公式サイトURL"]).netloc in portal:
                # 電話番号もそのポータルページから拾った可能性が高いため合わせて除去する
                for col in ("公式サイトURL", "メールアドレス", "電話番号",
                            "インスタURL", "問い合わせフォームURL"):
                    r[col] = ""

    df = pd.DataFrame(deduped, columns=OUTPUT_COLS)
    df.drop_duplicates(subset=["名称"], keep="first", inplace=True)
    for col in ("問い合わせフォームURL", "メールアドレス"):
        m = df[col] == ""
        df = pd.concat([df[m], df[~m].drop_duplicates(subset=[col], keep="first")]).sort_index()
    df = df.sort_values("名称").reset_index(drop=True)
    df.to_excel(OUTPUT_FILE, index=False)

    n_phone = (df["電話番号"] != "").sum()
    n_site = (df["公式サイトURL"] != "").sum()
    n_mail = (df["メールアドレス"] != "").sum()
    elapsed = int((datetime.now() - start).total_seconds())
    logger.info("=" * 65)
    logger.info(f"完了: {OUTPUT_FILE} に {len(df)} 件を出力")
    logger.info(f"  電話番号あり {n_phone} / 公式サイト {n_site} / メール {n_mail}")
    logger.info(f"所要時間: {elapsed // 60}分{elapsed % 60}秒")
    logger.info("=" * 65)


if __name__ == "__main__":
    asyncio.run(main())
