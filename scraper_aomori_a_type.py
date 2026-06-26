"""
青森県 就労継続支援A型事業所 情報収集

データソース:
  Phase 1 - shogaisha-shuro.com から A型事業所リスト取得
            (requests + BeautifulSoup)
  Phase 2 - WAM (wam.go.jp) から追加情報を Playwright で取得
            (Playwright 経由でフォーム操作して青森県A型を検索)
  Phase 3 - 各公式サイトから インスタURL・問い合わせフォームURL を取得
            (Playwright + BeautifulSoup)

出力: aomori_a_type_YYYYMMDD_HHMMSS.xlsx
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

LOG_FILE = "scraper_aomori_a_type.log"
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
OUTPUT_FILE = f"aomori_a_type_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/124.0.0.0 Safari/537.36"
)
HEADERS = {"User-Agent": USER_AGENT}

SHURO_BASE = "https://shogaisha-shuro.com"
SHURO_LIST_URL = f"{SHURO_BASE}/category/shuro/aomori/"

WAM_TOP_URL = "https://www.wam.go.jp/sfkohyoout/COP000100E0000.do"
WAM_BASE = "https://www.wam.go.jp"

SKIP_DOMAINS = [
    "instagram.com", "facebook.com", "twitter.com", "x.com", "youtube.com",
    "tiktok.com", "wikipedia.org", "google.com", "bing.com",
    "shogaisha-shuro.com", "wam.go.jp",
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
    rp = _get_robots(url)
    return rp.can_fetch(USER_AGENT, url)


def _fetch(url: str) -> BeautifulSoup | None:
    try:
        r = requests.get(url, headers=HEADERS, timeout=15)
        r.encoding = "utf-8"
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
    """A型事業所の (名称, 詳細URL) リストを返す。"""
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
    """詳細ページから基本情報を取得する。"""
    rec: dict = {
        "名称": "", "メールアドレス": "", "公式サイトURL": "",
        "所在地": "", "電話番号": "", "インスタURL": "", "問い合わせフォームURL": "",
        "事業所種別": "", "作業内容": "", "運営法人": "",
    }
    soup = _fetch(detail_url)
    if soup is None:
        return rec

    # 名称は H1 or H2
    h = soup.find("h1") or soup.find("h2")
    if h:
        rec["名称"] = h.get_text(strip=True)

    # テーブルの各行を解析
    for tr in soup.select("table tr"):
        th = tr.find("th")
        td = tr.find("td")
        if not th or not td:
            continue
        key = th.get_text(strip=True)
        val = td.get_text(separator=" ", strip=True)

        if key == "所在地":
            # 「〒XXX-XXXX 青森県…」→ 郵便番号除去
            addr = re.sub(r"〒\s*\d{3}[-－]\d{4}\s*", "", val).strip()
            rec["所在地"] = addr
        elif key == "電話番号":
            phone = re.search(r"[\d－\-()（）]{6,}", val)
            rec["電話番号"] = phone.group(0).strip() if phone else val.split()[0] if val else ""
        elif key == "Eメール":
            # Cloudflare のメール難読化を先に試みる
            cf_el = td.find(attrs={"data-cfemail": True})
            if cf_el:
                decoded = _decode_cfemail(cf_el["data-cfemail"])
                if decoded:
                    rec["メールアドレス"] = decoded
            else:
                email = re.search(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}", val)
                if email:
                    rec["メールアドレス"] = email.group(0)
        elif key == "URL":
            link = td.find("a", href=True)
            if link:
                rec["公式サイトURL"] = link["href"].strip()
        elif key == "事業所種別":
            rec["事業所種別"] = val
        elif key == "作業内容":
            rec["作業内容"] = val[:200]
        elif key in ("運営法人等", "運営法人"):
            rec["運営法人"] = val

    return rec


# ── Phase 2: WAM (wam.go.jp) ─────────────────────────────

async def collect_wam_aomori(page) -> list[dict]:
    """
    Playwright で WAM の青森県 就労継続支援A型 を検索して全件収集する。
    取得できた場合のみリストを返す（失敗時は空リスト）。
    """
    records: list[dict] = []
    try:
        logger.info("WAM: トップページにアクセス")
        await page.goto(WAM_TOP_URL, timeout=30000, wait_until="load")
        await page.wait_for_timeout(3000)

        # #pref02 ボタンの存在確認
        pref_btn = await page.query_selector("#pref02")
        if pref_btn:
            # ナビゲーション完了を待ちながらクリック
            async with page.expect_navigation(timeout=15000):
                await pref_btn.click()
            await page.wait_for_timeout(2000)
        else:
            # pref02 がない場合は form の hidden field に青森コードをセットして JS 実行
            await page.evaluate("""() => {
                const form = document.querySelector('form');
                if (!form) return;
                const hidden = form.querySelector('[name="vo_headVO_prefCode"]') ||
                               form.querySelector('[name="prefCode"]');
                if (hidden) hidden.value = '02';
            }""")
            await page.evaluate(
                "doTransaction('COP000101E00',null,false,null,document.querySelector('form'),null,null)"
            )
            await page.wait_for_load_state("load", timeout=15000)
            await page.wait_for_timeout(2000)

        logger.info(f"WAM: 青森県ページ URL = {page.url}")

        # サービス種別「就労継続支援A型」を選択
        # セレクトボックスを探す
        sel_handles = await page.query_selector_all("select")
        service_sel = None
        for sel in sel_handles:
            opts = await sel.evaluate("el => Array.from(el.options).map(o => o.text)")
            if any("就労継続支援" in o for o in opts):
                service_sel = sel
                break

        if service_sel:
            await service_sel.evaluate("""el => {
                for (let o of el.options) {
                    if (o.text.includes('就労継続支援Ａ型') || o.text.includes('就労継続支援A型')) {
                        el.value = o.value;
                        break;
                    }
                }
            }""")
            logger.info("WAM: サービス種別 = 就労継続支援A型 を選択")
        else:
            logger.warning("WAM: サービス種別セレクトが見つからず")

        # 検索実行
        search_btn = await page.query_selector("input[type='submit'], button[type='submit']")
        if search_btn:
            await search_btn.click()
        else:
            await page.evaluate(
                "doTransaction('COP000102E00',null,false,null,document.querySelector('form'),null,null)"
            )
        await page.wait_for_load_state("networkidle", timeout=15000)
        await page.wait_for_timeout(2000)
        logger.info(f"WAM: 検索結果 URL = {page.url}")

        # ページネーションを含む全件収集
        page_num = 1
        while True:
            html = await page.content()
            soup = BeautifulSoup(html, "lxml")
            rows = _parse_wam_result_page(soup)
            records.extend(rows)
            logger.info(f"WAM: ページ {page_num} → {len(rows)} 件 (累計 {len(records)} 件)")

            # 次ページリンク
            next_link = soup.find("a", string=re.compile(r"次[のへ]|next", re.I))
            if not next_link:
                # 数字ページリンクで次ページを探す
                current_links = soup.find_all("a", href=True)
                next_href = None
                for a in current_links:
                    if re.search(rf"page[=:]?{page_num + 1}", a.get("href", ""), re.I):
                        next_href = a["href"]
                        break
                if not next_href:
                    break
                abs_next = urljoin(WAM_BASE, next_href)
                await page.goto(abs_next, timeout=15000)
            else:
                href = next_link.get("href", "")
                if href.startswith("javascript"):
                    await page.evaluate(f"doTransaction('COP000102E00',null,false,null,document.querySelector('form'),null,null)")
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


def _parse_wam_result_page(soup: BeautifulSoup) -> list[dict]:
    """WAM 検索結果ページから事業所情報を抽出する。"""
    records: list[dict] = []
    # WAM は table 形式または dl 形式で結果を表示
    for tr in soup.select("table.result tr, table.list tr, tbody tr"):
        cells = tr.find_all("td")
        if len(cells) < 2:
            continue
        name = cells[0].get_text(strip=True)
        if not name or name in ("施設名", "事業所名"):
            continue
        rec: dict = {
            "名称": name,
            "メールアドレス": "",
            "公式サイトURL": "",
            "所在地": "",
            "電話番号": "",
            "インスタURL": "",
            "問い合わせフォームURL": "",
            "事業所種別": "",
            "作業内容": "",
            "運営法人": "",
        }
        for cell in cells[1:]:
            text = cell.get_text(strip=True)
            if re.match(r"[\d０-９]{2,4}[-－ー][\d０-９]{3,4}[-－ー][\d０-９]{4}", text):
                rec["電話番号"] = text
            elif "県" in text or "市" in text or "町" in text or "村" in text:
                addr = re.sub(r"〒\s*\d{3}[-－]\d{4}\s*", "", text).strip()
                if addr:
                    rec["所在地"] = addr
        records.append(rec)
    return records


# ── Phase 3: 公式サイト追加情報 ──────────────────────────

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

    # メールアドレス（ページテキストから）
    emails = re.findall(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}", text)
    emails = [e for e in emails if not re.search(r"\.(png|jpg|gif|svg|webp)$", e, re.I)]
    if emails and not result.get("メールアドレス"):
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


# ── メイン ────────────────────────────────────────────────

async def main() -> None:
    logger.info("=" * 60)
    logger.info("青森県 就労継続支援A型事業所 収集開始")
    logger.info("=" * 60)
    start_time = datetime.now()

    # Phase 1: shogaisha-shuro.com から A型リストと詳細情報取得
    logger.info("Phase 1: shogaisha-shuro.com から A型事業所収集")
    shuro_list = collect_shuro_list()
    records: list[dict] = []
    seen_keys: set[str] = set()

    for i, (name_hint, detail_url) in enumerate(shuro_list, 1):
        rec = parse_shuro_detail(detail_url)
        if not rec["名称"]:
            rec["名称"] = name_hint
        key = (rec["名称"], rec["電話番号"])
        if key not in seen_keys:
            seen_keys.add(key)
            records.append(rec)
        if i % 10 == 0 or i == len(shuro_list):
            logger.info(f"  [{i}/{len(shuro_list)}] {rec['名称']}")
        time.sleep(random.uniform(0.5, 1.2))

    logger.info(f"Phase 1 完了: {len(records)} 件")

    # Phase 2: WAM から追加取得 (Playwright)
    logger.info("Phase 2: WAM から青森県A型事業所収集")
    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=True)
        context = await browser.new_context(user_agent=USER_AGENT)
        wam_page = await context.new_page()

        wam_records = await collect_wam_aomori(wam_page)

        # WAM 結果をマージ（名称が既存に存在しなければ追加）
        added = 0
        existing_names = {r["名称"] for r in records}
        for wr in wam_records:
            if wr["名称"] and wr["名称"] not in existing_names:
                existing_names.add(wr["名称"])
                records.append(wr)
                added += 1
        logger.info(f"WAM 新規追加: {added} 件 (重複除外後 合計 {len(records)} 件)")

        # Phase 3: 公式サイトから追加情報取得
        logger.info("Phase 3: 公式サイト収集開始")
        official_page = await context.new_page()
        with_url = [r for r in records if r.get("公式サイトURL")]
        logger.info(f"  公式URL あり: {len(with_url)} / {len(records)} 件")

        for i, rec in enumerate(with_url, 1):
            url = rec["公式サイトURL"]
            extras = await scrape_official_site(url, official_page)
            # shogaisha-shuro でメール取得済みの場合は上書きしない
            if rec.get("メールアドレス") and "メールアドレス" in extras:
                del extras["メールアドレス"]
            rec.update(extras)

            if i % 10 == 0 or i == len(with_url):
                logger.info(f"  公式サイト取得: {i}/{len(with_url)}")
            await asyncio.sleep(random.uniform(1.0, 2.5))

        await browser.close()

    # Excel 出力
    columns = [
        "名称", "メールアドレス", "公式サイトURL",
        "所在地", "電話番号", "インスタURL", "問い合わせフォームURL",
        "事業所種別", "作業内容", "運営法人",
    ]
    df = pd.DataFrame(records, columns=columns)
    df.drop_duplicates(subset=["名称", "電話番号"], keep="first", inplace=True)
    df.to_excel(OUTPUT_FILE, index=False)

    elapsed = int((datetime.now() - start_time).total_seconds())
    logger.info("=" * 60)
    logger.info(f"完了: {OUTPUT_FILE} に {len(df)} 件を出力")
    logger.info(f"所要時間: {elapsed // 60}分{elapsed % 60}秒")
    logger.info("=" * 60)


if __name__ == "__main__":
    asyncio.run(main())
