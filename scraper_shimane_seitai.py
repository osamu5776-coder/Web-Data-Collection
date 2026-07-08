"""
島根県 整体・接骨院 情報収集

データソース:
  Phase 1 - shimane-jusei.or.jp/guide.html (島根県柔道整復師会 接骨院案内)
            から名称・電話番号・住所を収集
  Phase 2 - rairai.net/salontop/search/all/32 (らいらいネット 島根県)
            から整体・接骨・整骨系サロンを全22ページ収集
  Phase 3 - DuckDuckGo で各院の公式サイトを検索し、ページ本文に名称または電話番号が
            含まれるかで一致検証（不一致・ディレクトリ/紹介/官公庁サイトは除外）した上で
            メール・インスタ・問い合わせフォームURL を取得
            (ddgs + requests + BeautifulSoup, 3並列)

出力列: 名称, メールアドレス, 公式サイトURL, 所在地, 電話番号, インスタURL, 問い合わせフォームURL
出力ファイル: shimane_seitai_YYYYMMDD_HHMMSS.xlsx

※ 最終出力は「問い合わせフォームURL」「メールアドレス」の重複を除去する。
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

LOG_FILE = "scraper_shimane_seitai.log"
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
OUTPUT_FILE = f"shimane_seitai_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/124.0.0.0 Safari/537.36"
)
HEADERS = {"User-Agent": USER_AGENT}

JUSEI_URL = "https://shimane-jusei.or.jp/guide.html"

RAIRAI_BASE = "https://rairai.net"
RAIRAI_LIST_URL = "https://rairai.net/salontop/search/all/32"
RAIRAI_PAGES = 22

# 整体・接骨院として採用するカテゴリキーワード
CATEGORY_KEEP_KWS = ["整体", "接骨", "整骨"]

# LINE・総合プラットフォーム・予約サイト・口コミサイト・クーポンサイト・情報サイト等
# （公式サイトとして採用しない外部サイト）
SKIP_DOMAINS = [
    # データソース自身
    "shimane-jusei.or.jp", "rairai.net",
    # LINE
    "line.me", "lin.ee",
    # 総合プラットフォーム・予約・クーポン
    "hotpepper.jp", "beauty.hotpepper.jp", "epark.jp", "curama.jp",
    "reserva.be", "coubic.com", "salonboard.com", "minimo.jp",
    "rakuten.co.jp", "beauty.rakuten.co.jp",
    # 口コミサイト
    "ekiten.jp", "tabelog.com", "caloo.jp", "minkabu.jp",
    "google.com", "google.co.jp", "maps.google.com",
    # 情報・ディレクトリサイト（整骨院/接骨院まとめ・比較・検索サイト）
    "itp.ne.jp", "mapion.co.jp", "navitime.co.jp", "mapfan.com",
    "goo.ne.jp", "wikipedia.org", "homemate-research.com",
    "judo-ch.jp", "seikotsuin.info", "wellyou.net", "judi-channel.jp",
    "medley.life", "toresei.com", "seitai-rct.jp", "enma.jp",
    "chikuchikuryoho.com", "body-care.expert",
    "seikotsuguide.jp", "health-more.jp", "bonbonesquare.com",
    "minnanochiryoin.jp", "karadarefre.jp", "jiko24.jp",
    "seikotsuin-navi.com", "seitai-osusume-select.com",
    "treasuredtime.org", "findglocal.com", "mypl.net", "jusei.gr.jp",
    "ozmall.co.jp", "biglobe.ne.jp", "town.or.jp",
    "oue-c-clinic.com", "xn--t8j4aa4nmx460lnvbku9gi03awfof5y.jp",
    "koutsujiko-chiryo-biz.jp", "chiryo-biz.jp", "self-whitening.jp",
    "tsunagu-good.com", "jiko-navi.jp", "passpon.jp", "everytown.info",
    "jpaweb.jp", "shimane-hoken-hari9.com", "ameblo.jp",
    "otokoro.com", "learning-with.us", "ashigaru.jp",
    "seikotsu-navi.net", "reserven.jp", "salonnavigation.com",
    # SNS
    "instagram.com", "facebook.com", "twitter.com", "x.com",
    "youtube.com", "tiktok.com",
    # 検索エンジン
    "bing.com", "yahoo.co.jp",
]

# 官公庁ドメイン（市区町村サイトなどの誤マッチ除外）
GOV_DOMAIN_RE = re.compile(r"\.lg\.jp|^https?://(www\.)?city\.|//city\.", re.I)

OUTPUT_COLS = [
    "名称", "メールアドレス", "公式サイトURL",
    "所在地", "電話番号", "インスタURL", "問い合わせフォームURL",
]

# ── 共通ユーティリティ ────────────────────────────────────
_robots_cache: dict[str, RobotFileParser] = {}


def _is_allowed(url: str) -> bool:
    from urllib.parse import urlparse
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
    return _robots_cache[origin].can_fetch(USER_AGENT, url)


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
    if GOV_DOMAIN_RE.search(url):
        return True
    return any(d in url for d in SKIP_DOMAINS)


def _verify_match(html_text: str, name: str, phone: str, address: str = "") -> bool:
    """取得した公式サイトが実際にその院のものかを名称・電話番号・所在地で検証する。

    「姓+整骨院/接骨院」のような名称は全国どこにでも存在しうるため、名称一致
    だけでは同名の無関係な別事業者（他都道府県の同名チェーン等）を誤って
    採用してしまう。電話番号の完全一致がない限り、名称一致に加えて住所の
    市区町村名がページ内に含まれることを必須とする。
    """
    norm_page = _normalize_name(re.sub(r"[\s　]+", "", html_text))

    phone_digits = _normalize_phone(phone)
    if phone_digits and phone_digits in re.sub(r"[^0-9]", "", html_text):
        return True

    norm_name = _normalize_name(name)
    name_hit = bool(norm_name) and norm_name in norm_page

    # 「整骨院」「接骨院」等の一般的な業種語を除いた固有部分で再チェック
    core = re.sub(
        r"(整骨院|接骨院|整体院|治療院|鍼灸院|鍼灸整骨院|はりきゅう整骨院|"
        r"整体|鍼灸|マッサージ|分院)+$",
        "", name
    ).strip()
    core_norm = _normalize_name(core)
    core_hit = len(core_norm) >= 2 and core_norm in norm_page

    if not (name_hit or core_hit):
        return False

    # 電話番号一致がない名称一致は、住所（市区町村）の裏付けを必須とする
    city_match = re.search(r"([^\s　]{2,8}?[市区町村])", address)
    city = city_match.group(1) if city_match else ""
    return bool(city) and _normalize_name(city) in norm_page


# ── Phase 1: shimane-jusei.or.jp ─────────────────────────

def scrape_jusei() -> list[dict]:
    """島根県柔道整復師会の接骨院案内から名称・電話番号・住所を収集する。"""
    soup = _fetch(JUSEI_URL)
    if soup is None:
        return []

    records: list[dict] = []
    for tr in soup.select("table.info tr"):
        tds = tr.find_all("td")
        if len(tds) < 5:
            continue
        name = tds[0].get_text(strip=True)
        phone = re.sub(r"[^0-9\-]", "", tds[2].get_text(strip=True))
        address = tds[4].get_text(strip=True)
        if not address.startswith("島根県"):
            address = "島根県" + address

        rec = {k: "" for k in OUTPUT_COLS}
        rec["名称"] = name
        rec["電話番号"] = phone
        rec["所在地"] = address
        records.append(rec)

    logger.info(f"shimane-jusei.or.jp: {len(records)} 件")
    return records


# ── Phase 2: rairai.net ──────────────────────────────────

def scrape_rairai_page(page_num: int) -> list[dict]:
    """らいらいネット島根県の1ページ分を収集する。"""
    url = RAIRAI_LIST_URL if page_num == 1 else f"{RAIRAI_LIST_URL}?page={page_num}"
    soup = _fetch(url)
    if soup is None:
        return []

    records: list[dict] = []
    for article in soup.select("article.search_list_wrap"):
        h2 = article.find("h2")
        if not h2:
            continue
        name = h2.get_text(strip=True)

        cate_el = article.find("div", class_="search_list_cate")
        category = cate_el.get_text(strip=True) if cate_el else ""
        if not any(kw in category for kw in CATEGORY_KEEP_KWS):
            continue

        address, phone = "", ""
        for line in article.find_all("div", class_="search_list_txt_line"):
            left = line.find("span", class_="txt_left")
            right = line.find("span", class_="txt_right")
            if not left or not right:
                continue
            label = left.get_text(strip=True)
            value = right.get_text(strip=True)
            if label == "住所":
                address = value
            elif label == "電話番号":
                phone = re.sub(r"[^0-9\-]", "", value)

        if address and not address.startswith("島根県"):
            address = "島根県" + address

        rec = {k: "" for k in OUTPUT_COLS}
        rec["名称"] = name
        rec["所在地"] = address
        rec["電話番号"] = phone
        records.append(rec)

    return records


def scrape_rairai() -> list[dict]:
    """らいらいネット島根県の全ページを収集する。"""
    all_records: list[dict] = []
    for page_num in range(1, RAIRAI_PAGES + 1):
        recs = scrape_rairai_page(page_num)
        all_records.extend(recs)
        logger.info(f"  rairai p{page_num}/{RAIRAI_PAGES}: {len(recs)} 件 (累計 {len(all_records)} 件)")
        time.sleep(random.uniform(0.8, 1.5))
    return all_records


# ── Phase 3: DuckDuckGo 公式サイト検索 ──────────────────

def search_official_site(name: str, address: str, phone: str) -> tuple[str, str]:
    """DuckDuckGo でクリニックの公式サイトを検索し、名称/電話番号で一致検証した
    URLと本文HTMLを返す（見つからない場合は ("", "")）。"""
    city_match = re.search(r"島根県\s*([^\s]{2,8}?[市区町村])", address)
    city = city_match.group(1) if city_match else "島根"
    query = f"{name} {city} 公式サイト"
    try:
        with DDGS() as ddgs:
            results = list(ddgs.text(query, max_results=6))
    except Exception as e:
        logger.debug(f"DDG失敗 {name}: {e}")
        return "", ""

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
        if _verify_match(resp.text, name, phone, address):
            return resp.url, resp.text
    return "", ""


# ── Phase 4: 公式サイト追加情報 ──────────────────────────

def _decode_cfemail(encoded: str) -> str:
    try:
        key = int(encoded[:2], 16)
        return bytes(
            int(encoded[i:i+2], 16) ^ key for i in range(2, len(encoded), 2)
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

    if not result.get("電話番号"):
        phones = re.findall(r"0\d{1,4}[-‐‑‒–—－]\d{1,4}[-‐‑‒–—－]\d{3,4}", text)
        if phones:
            result["電話番号"] = re.sub(r"[^\d\-]", "", phones[0])

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
        # 無料ホームページサービス等の運営元自身のサポート/削除依頼フォームを除外
        # （公式サイトと異なるホストへのリンクは採用しない）
        if urlparse(href).netloc != base_host:
            continue
        result["問い合わせフォームURL"] = href
        break

    return result


# ── メイン ────────────────────────────────────────────────

async def main() -> None:
    logger.info("=" * 65)
    logger.info("島根県 整体・接骨院 収集開始")
    logger.info("=" * 65)
    start_time = datetime.now()

    all_records: list[dict] = []
    seen_names: set[str] = set()
    seen_phones: set[str] = set()

    # ── Phase 1: shimane-jusei.or.jp ─────────────────────
    logger.info("Phase 1: shimane-jusei.or.jp から接骨院案内収集")
    jusei_records = scrape_jusei()
    for rec in jusei_records:
        nk = _normalize_name(rec["名称"])
        pk = _normalize_phone(rec.get("電話番号", ""))
        if nk and nk not in seen_names:
            seen_names.add(nk)
            if pk:
                seen_phones.add(pk)
            all_records.append(rec)
    logger.info(f"Phase 1 完了: {len(all_records)} 件")

    # ── Phase 2: rairai.net ──────────────────────────────
    logger.info("Phase 2: rairai.net から整体・接骨系サロン収集")
    rairai_records = scrape_rairai()
    rairai_added = 0
    for rec in rairai_records:
        nk = _normalize_name(rec["名称"])
        pk = _normalize_phone(rec.get("電話番号", ""))
        if nk in seen_names:
            continue
        if pk and pk in seen_phones:
            continue
        seen_names.add(nk)
        if pk:
            seen_phones.add(pk)
        all_records.append(rec)
        rairai_added += 1
    logger.info(f"Phase 2 完了: 新規追加 {rairai_added} 件 (合計 {len(all_records)} 件)")

    # ── Phase 3: DuckDuckGo 公式サイト検索+検証 (3並列) ───
    logger.info("Phase 3: DuckDuckGo で公式サイト検索・名称/電話番号で一致検証")
    no_url = [r for r in all_records if not r.get("公式サイトURL")]
    logger.info(f"  公式URL 未取得: {len(no_url)} 件")

    ddg_sem = asyncio.Semaphore(3)
    ddg_done = {"n": 0, "found": 0}

    async def _ddg_one(rec: dict) -> None:
        async with ddg_sem:
            url, html = await asyncio.to_thread(
                search_official_site, rec["名称"], rec.get("所在地", ""), rec.get("電話番号", "")
            )
            if url:
                rec["公式サイトURL"] = url
                extras = _parse_official_html(url, html)
                # 出所元データの電話番号・所在地は上書きしない
                extras.pop("電話番号", None)
                rec.update(extras)
                ddg_done["found"] += 1
            ddg_done["n"] += 1
            n = ddg_done["n"]
            if n % 30 == 0 or n == len(no_url):
                logger.info(f"  DDG検索: {n}/{len(no_url)} (発見 {ddg_done['found']} 件)")
            await asyncio.sleep(random.uniform(1.0, 2.0))

    await asyncio.gather(*[_ddg_one(rec) for rec in no_url])
    logger.info(f"Phase 3 完了: 検証済み公式URL 合計 {sum(1 for r in all_records if r.get('公式サイトURL'))} 件")

    # ── Excel 出力 ────────────────────────────────────────
    df = pd.DataFrame(all_records, columns=OUTPUT_COLS)
    df.drop_duplicates(subset=["名称"], keep="first", inplace=True)
    # 問い合わせフォームURL・メールアドレスの重複除去（空欄は対象外）
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
