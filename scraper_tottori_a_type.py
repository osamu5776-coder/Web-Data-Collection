"""
鳥取県 就労継続支援A型事業所 情報収集

データソース:
  Phase 1 - shogaisha-shuro.com から A型事業所リスト取得 (32件)
            (requests + BeautifulSoup)
  Phase 2 - 鳥取県公式DB (db.pref.tottori.jp) から追加情報取得 (21件)
            (requests POST + BeautifulSoup)
  Phase 3 - WAM (wam.go.jp) から Playwright でフォーム操作して追加収集
            (Playwright + BeautifulSoup, ベストエフォート)
  Phase 4 - 各公式サイトから インスタURL・問い合わせフォームURL を取得
            (Playwright + BeautifulSoup)

出力: tottori_a_type_YYYYMMDD_HHMMSS.xlsx
"""

import asyncio
import io
import logging
import random
import re
import sys
import time
from datetime import datetime
from urllib.parse import urljoin
from urllib.robotparser import RobotFileParser

import pandas as pd
import requests
from bs4 import BeautifulSoup
from playwright.async_api import async_playwright

# ── ロギング ──────────────────────────────────────────────
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

LOG_FILE = "scraper_tottori_a_type.log"
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
OUTPUT_FILE = f"tottori_a_type_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/124.0.0.0 Safari/537.36"
)
HEADERS = {"User-Agent": USER_AGENT}

SHURO_BASE = "https://shogaisha-shuro.com"
SHURO_LIST_URL = f"{SHURO_BASE}/category/shuro/tottori/"

TOTTORI_DB_BASE = "https://db.pref.tottori.jp"
TOTTORI_DB_SEARCH = f"{TOTTORI_DB_BASE}/heartful.nsf/navi4.htm?OpenForm&Seq=1"
# 就労継続支援Ａ型事業所のフォーム値
TOTTORI_A_TYPE_VALUE = "_a227k244o9888os0gi6oh13jo224oe442c088ongghqb112u6227ok_"

WAM_TOP_URL = "https://www.wam.go.jp/sfkohyoout/COP000100E0000.do"
WAM_BASE = "https://www.wam.go.jp"

SKIP_DOMAINS = [
    "instagram.com", "facebook.com", "twitter.com", "x.com", "youtube.com",
    "tiktok.com", "wikipedia.org", "google.com", "bing.com",
    "shogaisha-shuro.com", "wam.go.jp", "db.pref.tottori.jp",
    "tabelog.com", "hotpepper.jp",
]

# ── robots.txt キャッシュ ─────────────────────────────────
_robots_cache: dict[str, RobotFileParser] = {}

def _get_robots(url: str) -> RobotFileParser:
    from urllib.parse import urlparse
    parsed = urlparse(url)
    origin = f"{parsed.scheme}://{parsed.netloc}"
    if origin in _robots_cache:
        return _robots_cache[origin]
    rp = RobotFileParser()
    try:
        rp.set_url(f"{origin}/robots.txt")
        rp.read()
    except Exception:
        pass
    _robots_cache[origin] = rp
    return rp

def _is_allowed(url: str) -> bool:
    return _get_robots(url).can_fetch(USER_AGENT, url)


def _fetch(url: str, encoding: str = "utf-8") -> BeautifulSoup | None:
    try:
        r = requests.get(url, headers=HEADERS, timeout=15)
        r.encoding = encoding
        return BeautifulSoup(r.text, "lxml")
    except Exception as e:
        logger.warning(f"取得失敗: {url} / {e}")
        return None


def _decode_cfemail(encoded: str) -> str:
    """Cloudflare の data-cfemail 属性を復号してメールアドレスを返す。"""
    try:
        key = int(encoded[:2], 16)
        return bytes(
            int(encoded[i:i+2], 16) ^ key for i in range(2, len(encoded), 2)
        ).decode("utf-8")
    except Exception:
        return ""


# ── Phase 1: shogaisha-shuro.com ─────────────────────────

def collect_shuro_list() -> list[tuple[str, str]]:
    """A型事業所の (名称ヒント, 詳細URL) リストを返す。"""
    soup = _fetch(SHURO_LIST_URL)
    if soup is None:
        return []

    results: list[tuple[str, str]] = []
    for li in soup.select("ul.facility_list li.clearfix"):
        if not li.find("div", class_="type-a-icon"):
            continue
        name_div = li.find("div", class_="institution-name")
        if not name_div:
            continue
        a = name_div.find("a", href=True)
        if not a:
            continue
        name = name_div.get_text(strip=True)
        href = a["href"]
        url = href if href.startswith("http") else urljoin(SHURO_BASE, href)
        results.append((name, url))

    logger.info(f"shogaisha-shuro.com A型リスト: {len(results)} 件")
    return results


def parse_shuro_detail(detail_url: str) -> dict:
    """shogaisha-shuro.com の詳細ページから基本情報を取得する。"""
    rec: dict = {
        "名称": "", "メールアドレス": "", "公式サイトURL": "",
        "所在地": "", "電話番号": "", "インスタURL": "", "問い合わせフォームURL": "",
        "事業所種別": "", "作業内容": "", "運営法人": "",
        "_source": "shogaisha-shuro",
    }
    soup = _fetch(detail_url)
    if soup is None:
        return rec

    h = soup.find("h1") or soup.find("h2")
    if h:
        rec["名称"] = h.get_text(strip=True)

    for tr in soup.select("table tr"):
        th = tr.find("th")
        td = tr.find("td")
        if not th or not td:
            continue
        key = th.get_text(strip=True)
        val = td.get_text(separator=" ", strip=True)

        if key == "所在地":
            rec["所在地"] = re.sub(r"〒\s*\d{3}[-－]\d{4}\s*", "", val).strip()
        elif key == "電話番号":
            phone = re.search(r"[\d\-]{6,}", val.replace("－", "-"))
            rec["電話番号"] = phone.group(0) if phone else ""
        elif key == "Eメール":
            cf_el = td.find(attrs={"data-cfemail": True})
            if cf_el:
                decoded = _decode_cfemail(cf_el["data-cfemail"])
                if decoded:
                    rec["メールアドレス"] = decoded
            else:
                m = re.search(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}", val)
                if m:
                    rec["メールアドレス"] = m.group(0)
        elif key == "URL":
            link = td.find("a", href=True)
            if link:
                href = link["href"].strip()
                # 鳥取県DBのURLは公式サイトではないため除外
                if "db.pref.tottori.jp" not in href:
                    rec["公式サイトURL"] = href
        elif key == "事業所種別":
            rec["事業所種別"] = val
        elif key == "作業内容":
            rec["作業内容"] = val[:200]
        elif key in ("運営法人等", "運営法人"):
            rec["運営法人"] = val

    return rec


# ── Phase 2: 鳥取県公式DB ────────────────────────────────

def collect_tottori_db() -> list[dict]:
    """鳥取県公式DBから就労継続支援A型事業所の一覧を取得する。"""
    post_data = {
        "__Click": "0",
        "Query": "",
        "%%Surrogate_area": "1",
        "area": "",
        "%%Surrogate_institutionq": "1",
        "institutionq": TOTTORI_A_TYPE_VALUE,
    }
    try:
        r = requests.post(TOTTORI_DB_SEARCH, data=post_data, headers=HEADERS, timeout=15)
        r.encoding = "shift_jis"
        soup = BeautifulSoup(r.text, "lxml")
    except Exception as e:
        logger.warning(f"鳥取県DB 一覧取得失敗: {e}")
        return []

    table = soup.find("table")
    if not table:
        logger.warning("鳥取県DB: テーブルが見つかりません")
        return []

    rows = table.find_all("tr")
    records: list[dict] = []

    for row in rows[1:]:  # ヘッダ行をスキップ
        cells = row.find_all("td")
        if len(cells) < 2:
            continue
        name = cells[1].get_text(strip=True) if len(cells) > 1 else ""
        if not name:
            continue

        # 詳細ページリンク
        detail_link = row.find("a", href=True)
        detail_url = ""
        if detail_link:
            href = detail_link["href"].replace("?OpenD", "?OpenDocument")
            detail_url = urljoin(TOTTORI_DB_BASE + "/heartful.nsf/", href)

        rec: dict = {
            "名称": name,
            "メールアドレス": "",
            "公式サイトURL": "",
            "所在地": cells[2].get_text(strip=True) if len(cells) > 2 else "",
            "電話番号": "",
            "インスタURL": "",
            "問い合わせフォームURL": "",
            "事業所種別": "就労継続支援A型",
            "作業内容": cells[3].get_text(strip=True)[:200] if len(cells) > 3 else "",
            "運営法人": "",
            "_source": "tottori-db",
            "_detail_url": detail_url,
        }
        records.append(rec)
        time.sleep(random.uniform(0.3, 0.7))

    logger.info(f"鳥取県DB 一覧: {len(records)} 件")
    return records


def fetch_tottori_db_detail(rec: dict) -> None:
    """鳥取県DB 詳細ページから電話・サイトURLを取得して rec を更新する（in-place）。"""
    detail_url = rec.pop("_detail_url", "")
    if not detail_url:
        return

    try:
        r = requests.get(detail_url, headers=HEADERS, timeout=15)
        r.encoding = "shift_jis"
        soup = BeautifulSoup(r.text, "lxml")
    except Exception as e:
        logger.debug(f"鳥取DB詳細取得失敗: {detail_url} / {e}")
        return

    # dl > dt + dd 構造
    dl = soup.find("dl")
    if dl:
        dts = dl.find_all("dt")
        dds = dl.find_all("dd")
        for dt, dd in zip(dts, dds):
            key = dt.get_text(strip=True)
            val = dd.get_text(strip=True)
            link = dd.find("a", href=True)

            if key == "電話":
                phone = re.search(r"[\d\-]{6,}", val.replace("－", "-"))
                if phone:
                    rec["電話番号"] = phone.group(0)
            elif key == "住所":
                if not rec.get("所在地"):
                    rec["所在地"] = val
            elif key == "サイト" and link:
                href = link["href"]
                if href.startswith("http") and not any(d in href for d in SKIP_DOMAINS):
                    rec["公式サイトURL"] = href
            elif key in ("運営主体",):
                if not rec.get("運営法人"):
                    rec["運営法人"] = val


# ── Phase 3: WAM (wam.go.jp) ─────────────────────────────

async def collect_wam_tottori(page) -> list[dict]:
    """
    Playwright で WAM の鳥取県 就労継続支援A型 を検索して全件収集する。
    取得できた場合のみリストを返す（失敗時は空リスト）。
    """
    records: list[dict] = []
    try:
        logger.info("WAM: トップページにアクセス")
        await page.goto(WAM_TOP_URL, timeout=30000, wait_until="load")
        await page.wait_for_timeout(3000)

        # 鳥取県ボタン (id="pref31")
        pref_btn = await page.query_selector("#pref31")
        if not pref_btn:
            logger.warning("WAM: #pref31 ボタンが見つかりません")
            return records

        # ボタンクリックでフォーム送信（POSTナビゲーション）
        await page.evaluate("""() => {
            const btn = document.querySelector('#pref31');
            if (btn && btn.onclick) btn.onclick();
        }""")
        await page.wait_for_timeout(2000)

        # フォーム経由で強制実行
        try:
            await page.evaluate(
                "doTransaction('COP000101E00',null,false,null,document.querySelector('form'),null,null)"
            )
        except Exception:
            pass

        await page.wait_for_load_state("networkidle", timeout=15000)
        await page.wait_for_timeout(2000)
        logger.info(f"WAM: 遷移後 URL = {page.url}")

        if "COP000100E0000" in page.url:
            logger.warning("WAM: 鳥取ページへの遷移失敗")
            return records

        # サービス種別「就労継続支援A型」を選択
        sel_handles = await page.query_selector_all("select")
        for sel in sel_handles:
            opts = await sel.evaluate("el => Array.from(el.options).map(o => o.text)")
            if any("就労継続支援" in o for o in opts):
                await sel.evaluate("""el => {
                    for (let o of el.options) {
                        if (o.text.includes('就労継続支援Ａ型') || o.text.includes('就労継続支援A型')) {
                            el.value = o.value;
                            break;
                        }
                    }
                }""")
                logger.info("WAM: サービス種別 = 就労継続支援A型 を選択")
                break

        # 検索実行 (form がない場合は submit ボタン経由)
        try:
            form_exists = await page.evaluate("!!document.querySelector('form')")
            if form_exists:
                await page.evaluate(
                    "doTransaction('COP000102E00',null,false,null,document.querySelector('form'),null,null)"
                )
            else:
                search_btn = await page.query_selector("input[type='submit'], button[type='submit']")
                if search_btn:
                    await search_btn.click()
        except Exception as e_search:
            logger.warning(f"WAM: 検索実行失敗 - {e_search}")
            return records

        await page.wait_for_load_state("networkidle", timeout=15000)
        await page.wait_for_timeout(2000)
        logger.info(f"WAM: 検索結果 URL = {page.url}")

        # 全ページを収集
        page_num = 1
        while True:
            html = await page.content()
            soup = BeautifulSoup(html, "lxml")
            rows = _parse_wam_result(soup)
            records.extend(rows)
            logger.info(f"WAM: ページ {page_num} → {len(rows)} 件")

            next_a = soup.find("a", string=re.compile(r"次[のへ]|>|next", re.I))
            if not next_a:
                break
            href = next_a.get("href", "")
            if href.startswith("javascript"):
                await page.evaluate(
                    "doTransaction('COP000102E00',null,false,null,document.querySelector('form'),null,null)"
                )
            else:
                await page.goto(urljoin(WAM_BASE, href), timeout=15000)
            await page.wait_for_load_state("networkidle", timeout=10000)
            await page.wait_for_timeout(1500)
            page_num += 1
            if page_num > 20:
                break

    except Exception as e:
        logger.warning(f"WAM 収集エラー: {e}")

    logger.info(f"WAM 合計: {len(records)} 件")
    return records


def _parse_wam_result(soup: BeautifulSoup) -> list[dict]:
    records: list[dict] = []
    for tr in soup.select("table.result tr, table.list tr, tbody tr"):
        cells = tr.find_all("td")
        if len(cells) < 2:
            continue
        name = cells[0].get_text(strip=True)
        if not name or name in ("施設名", "事業所名"):
            continue
        rec: dict = {
            "名称": name, "メールアドレス": "", "公式サイトURL": "",
            "所在地": "", "電話番号": "", "インスタURL": "", "問い合わせフォームURL": "",
            "事業所種別": "就労継続支援A型", "作業内容": "", "運営法人": "",
            "_source": "wam",
        }
        for cell in cells[1:]:
            text = cell.get_text(strip=True)
            if re.match(r"[\d\-]{6,}", text):
                rec["電話番号"] = text
            elif re.search(r"[県市町村]", text) and not rec["所在地"]:
                rec["所在地"] = re.sub(r"〒\s*\d{3}[-－]\d{4}\s*", "", text).strip()
        records.append(rec)
    return records


# ── Phase 4: 公式サイト追加情報 ──────────────────────────

async def scrape_official_site(url: str, page) -> dict:
    """公式サイトからメール・インスタ・問い合わせURLを取得する。"""
    result: dict = {}
    if not _is_allowed(url):
        return result
    if any(d in url for d in SKIP_DOMAINS):
        return result

    try:
        resp = await page.goto(url, timeout=18000, wait_until="domcontentloaded")
        if resp and resp.status >= 400:
            return result
        await page.wait_for_timeout(1500)
        html = await page.content()
    except Exception as e:
        logger.debug(f"Playwright 失敗 {url}: {e}")
        try:
            r = requests.get(url, headers=HEADERS, timeout=10)
            r.encoding = r.apparent_encoding
            html = r.text
        except Exception:
            return result

    soup = BeautifulSoup(html, "lxml")
    text = soup.get_text(separator="\n")

    # メールアドレス
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

    # Instagram URL
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if "instagram.com" in href and "/p/" not in href and "/reel/" not in href:
            if href.startswith("//"):
                href = "https:" + href
            result["インスタURL"] = href.rstrip("/")
            break

    # 問い合わせフォームURL
    contact_kws = ["contact", "inquiry", "お問い合わせ", "問い合わせ", "ご相談", "メールフォーム"]
    for a in soup.find_all("a", href=True):
        href = a["href"]
        link_text = a.get_text(strip=True)
        if not any(kw in href.lower() or kw in link_text for kw in contact_kws):
            continue
        if href.startswith("http"):
            result["問い合わせフォームURL"] = href
        elif href.startswith("//"):
            result["問い合わせフォームURL"] = "https:" + href
        elif not href.startswith(("#", "mailto:", "tel:")):
            result["問い合わせフォームURL"] = urljoin(url, href)
        if "問い合わせフォームURL" in result:
            break

    return result


# ── マージ補助 ────────────────────────────────────────────

def _normalize_name(s: str) -> str:
    """名称を正規化して重複チェックに使う。"""
    s = re.sub(r"[\s　　]+", "", s)
    s = s.replace("（", "(").replace("）", ")")
    s = re.sub(r"[Ａ-Ｚａ-ｚ０-９]", lambda m: chr(ord(m.group(0)) - 0xFEE0), s)
    return s.lower()


def _normalize_phone(s: str) -> str:
    return re.sub(r"[^0-9]", "", s)


# ── メイン ────────────────────────────────────────────────

async def main() -> None:
    logger.info("=" * 65)
    logger.info("鳥取県 就労継続支援A型事業所 収集開始")
    logger.info("=" * 65)
    start_time = datetime.now()

    # ── Phase 1: shogaisha-shuro.com ──────────────────────
    logger.info("Phase 1: shogaisha-shuro.com から A型事業所収集")
    shuro_list = collect_shuro_list()
    records: list[dict] = []
    seen_names: set[str] = set()
    seen_phones: set[str] = set()

    for i, (name_hint, detail_url) in enumerate(shuro_list, 1):
        rec = parse_shuro_detail(detail_url)
        if not rec["名称"]:
            rec["名称"] = name_hint
        nk = _normalize_name(rec["名称"])
        pk = _normalize_phone(rec.get("電話番号", ""))
        if nk and nk not in seen_names:
            seen_names.add(nk)
            if pk:
                seen_phones.add(pk)
            records.append(rec)
        if i % 10 == 0 or i == len(shuro_list):
            logger.info(f"  [{i}/{len(shuro_list)}] {rec['名称']}")
        time.sleep(random.uniform(0.5, 1.2))

    logger.info(f"Phase 1 完了: {len(records)} 件")

    # ── Phase 2: 鳥取県公式DB ─────────────────────────────
    logger.info("Phase 2: 鳥取県公式DB から A型事業所収集")
    db_records = collect_tottori_db()
    db_added = 0

    for rec in db_records:
        nk = _normalize_name(rec["名称"])
        # 重複チェック
        if nk in seen_names:
            continue
        # 詳細ページで電話・URLを補完
        fetch_tottori_db_detail(rec)
        pk = _normalize_phone(rec.get("電話番号", ""))
        if pk and pk in seen_phones:
            continue  # 電話番号一致 → 重複
        seen_names.add(nk)
        if pk:
            seen_phones.add(pk)
        records.append(rec)
        db_added += 1
        time.sleep(random.uniform(0.4, 0.9))

    logger.info(f"Phase 2 完了: 新規追加 {db_added} 件 (合計 {len(records)} 件)")

    # ── Phase 3: WAM (Playwright, ベストエフォート) ────────
    logger.info("Phase 3: WAM から鳥取県A型事業所収集")
    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=True)
        context = await browser.new_context(user_agent=USER_AGENT)
        wam_page = await context.new_page()

        wam_records = await collect_wam_tottori(wam_page)
        wam_added = 0
        for wr in wam_records:
            nk = _normalize_name(wr.get("名称", ""))
            pk = _normalize_phone(wr.get("電話番号", ""))
            if nk and nk not in seen_names:
                if not pk or pk not in seen_phones:
                    seen_names.add(nk)
                    if pk:
                        seen_phones.add(pk)
                    records.append(wr)
                    wam_added += 1
        logger.info(f"Phase 3 完了: 新規追加 {wam_added} 件 (合計 {len(records)} 件)")

        # ── Phase 4: 公式サイト追加情報取得 ────────────────
        logger.info("Phase 4: 公式サイト収集開始")
        official_page = await context.new_page()
        with_url = [r for r in records if r.get("公式サイトURL")]
        logger.info(f"  公式URL あり: {len(with_url)} / {len(records)} 件")

        for i, rec in enumerate(with_url, 1):
            url = rec["公式サイトURL"]
            extras = await scrape_official_site(url, official_page)
            # 既存メールは上書きしない
            if rec.get("メールアドレス") and "メールアドレス" in extras:
                del extras["メールアドレス"]
            rec.update(extras)

            if i % 10 == 0 or i == len(with_url):
                logger.info(f"  公式サイト取得: {i}/{len(with_url)}")
            await asyncio.sleep(random.uniform(1.0, 2.5))

        await browser.close()

    # ── Excel 出力 ───────────────────────────────────────────
    output_cols = [
        "名称", "メールアドレス", "公式サイトURL",
        "所在地", "電話番号", "インスタURL", "問い合わせフォームURL",
        "事業所種別", "作業内容", "運営法人",
    ]
    # 内部フィールドを除去
    for rec in records:
        rec.pop("_source", None)

    df = pd.DataFrame(records, columns=output_cols)
    df.drop_duplicates(subset=["名称", "電話番号"], keep="first", inplace=True)
    df.to_excel(OUTPUT_FILE, index=False)

    elapsed = int((datetime.now() - start_time).total_seconds())
    logger.info("=" * 65)
    logger.info(f"完了: {OUTPUT_FILE} に {len(df)} 件を出力")
    logger.info(f"所要時間: {elapsed // 60}分{elapsed % 60}秒")
    logger.info("=" * 65)


if __name__ == "__main__":
    asyncio.run(main())
