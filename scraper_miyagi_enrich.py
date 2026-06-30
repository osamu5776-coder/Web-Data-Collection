"""
宮城県 整体・接骨院 追記スクリプト

既存の miyagi_seitai_*.xlsx を読み込み、
DuckDuckGo + 公式サイト収集でメール・インスタ・問い合わせURLを追記して
新しい Excel として出力する。
"""

import asyncio
import glob
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
from ddgs import DDGS

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

LOG_FILE = "scraper_miyagi_enrich.log"
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

OUTPUT_FILE = f"miyagi_seitai_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/124.0.0.0 Safari/537.36"
)
HEADERS = {"User-Agent": USER_AGENT}

SKIP_DOMAINS = [
    "judo-ch.jp", "homemate-research.com", "hotpepper.jp",
    "instagram.com", "facebook.com", "twitter.com", "x.com", "youtube.com",
    "tiktok.com", "wikipedia.org", "google.com", "bing.com",
    "mapion.co.jp", "itp.ne.jp", "ekiten.jp", "salonboard.com",
    "beauty.rakuten.co.jp", "enma.jp", "seikotsuin.info", "wellyou.net",
    "medley.life", "caloo.jp", "toresei.com",
]

OUTPUT_COLS = [
    "名称", "メールアドレス", "公式サイトURL",
    "所在地", "電話番号", "インスタURL", "問い合わせフォームURL",
]

_robots_cache: dict[str, RobotFileParser] = {}


def _is_allowed(url: str) -> bool:
    from urllib.parse import urlparse
    parsed = urlparse(url)
    origin = f"{parsed.scheme}://{parsed.netloc}"
    if origin in _robots_cache:
        rp = _robots_cache[origin]
    else:
        rp = RobotFileParser()
        try:
            rp.set_url(f"{origin}/robots.txt")
            rp.read()
        except Exception:
            pass
        _robots_cache[origin] = rp
    return rp.can_fetch(USER_AGENT, url)


def _decode_cfemail(encoded: str) -> str:
    try:
        key = int(encoded[:2], 16)
        return bytes(
            int(encoded[i:i+2], 16) ^ key for i in range(2, len(encoded), 2)
        ).decode("utf-8")
    except Exception:
        return ""


def search_official_site(name: str, address: str) -> str:
    city_match = re.search(r"宮城県\s*([^\s市区町村]{2,5}[市区町村])", address)
    city = city_match.group(1) if city_match else "宮城"
    query = f"{name} {city} 公式サイト"
    try:
        with DDGS() as ddgs:
            results = list(ddgs.text(query, max_results=5))
        for r in results:
            url = r.get("href", "")
            if not url:
                continue
            if any(d in url for d in SKIP_DOMAINS):
                continue
            if url.startswith("http"):
                return url
    except Exception as e:
        logger.debug(f"DDG失敗 {name}: {e}")
    return ""


def _parse_official_html(url: str, html: str) -> dict:
    result: dict = {}
    soup = BeautifulSoup(html, "lxml")
    text = soup.get_text(separator="\n")

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
        if href.startswith("http"):
            result["問い合わせフォームURL"] = href
        elif href.startswith("//"):
            result["問い合わせフォームURL"] = "https:" + href
        elif not href.startswith(("#", "mailto:", "tel:")):
            result["問い合わせフォームURL"] = urljoin(url, href)
        if "問い合わせフォームURL" in result:
            break

    return result


def _fetch_official_sync(url: str) -> dict:
    if any(d in url for d in SKIP_DOMAINS):
        return {}
    if not _is_allowed(url):
        return {}
    try:
        r = requests.get(url, headers=HEADERS, timeout=10, allow_redirects=True)
        if r.status_code >= 400:
            return {}
        r.encoding = r.apparent_encoding or "utf-8"
        return _parse_official_html(url, r.text)
    except Exception:
        return {}


async def main() -> None:
    logger.info("=" * 65)
    logger.info("宮城県 整体・接骨院 追記処理開始")
    logger.info("=" * 65)
    start_time = datetime.now()

    # 既存 Excel を読み込む
    files = sorted(glob.glob(
        "C:/Users/osamu/OneDrive/ドキュメント/Web Data Collection/miyagi_seitai_*.xlsx"
    ))
    if not files:
        logger.error("miyagi_seitai_*.xlsx が見つかりません")
        return
    src = files[-1]  # 最新ファイル
    logger.info(f"入力ファイル: {src}")
    df_in = pd.read_excel(src)
    # OUTPUT_COLS に合わせて列を揃える
    for col in OUTPUT_COLS:
        if col not in df_in.columns:
            df_in[col] = ""
    records = df_in[OUTPUT_COLS].fillna("").to_dict("records")
    logger.info(f"読み込み: {len(records)} 件")

    # ── Phase A: DuckDuckGo で公式サイト検索（3並列） ──────
    no_url = [r for r in records if not str(r.get("公式サイトURL", "")).startswith("http")]
    logger.info(f"Phase A: DDG検索 ({len(no_url)} 件, 3並列)")

    ddg_sem = asyncio.Semaphore(3)
    ddg_done = {"n": 0, "found": 0}

    async def _ddg_one(rec: dict) -> None:
        async with ddg_sem:
            url = await asyncio.to_thread(
                search_official_site, str(rec["名称"]), str(rec.get("所在地", ""))
            )
            if url:
                rec["公式サイトURL"] = url
                ddg_done["found"] += 1
            ddg_done["n"] += 1
            n = ddg_done["n"]
            if n % 30 == 0 or n == len(no_url):
                logger.info(f"  DDG: {n}/{len(no_url)} (発見 {ddg_done['found']} 件)")
            await asyncio.sleep(random.uniform(1.0, 2.0))

    await asyncio.gather(*[_ddg_one(rec) for rec in no_url])
    logger.info(f"Phase A 完了: 公式URL 取得済み {sum(1 for r in records if r.get('公式サイトURL'))} 件")

    # ── Phase B: 公式サイト収集（10並列 requests） ─────────
    with_url = [r for r in records if str(r.get("公式サイトURL", "")).startswith("http")]
    logger.info(f"Phase B: 公式サイト収集 ({len(with_url)} 件, 10並列)")

    b_sem = asyncio.Semaphore(10)
    b_done = {"n": 0}

    async def _b_one(rec: dict) -> None:
        async with b_sem:
            url = str(rec["公式サイトURL"])
            extras = await asyncio.to_thread(_fetch_official_sync, url)
            # 既存のメールは上書きしない
            if rec.get("メールアドレス") and "メールアドレス" in extras:
                del extras["メールアドレス"]
            rec.update(extras)
            b_done["n"] += 1
            n = b_done["n"]
            if n % 50 == 0 or n == len(with_url):
                logger.info(f"  公式サイト: {n}/{len(with_url)}")

    await asyncio.gather(*[_b_one(rec) for rec in with_url])
    logger.info("Phase B 完了")

    # Excel 出力
    df_out = pd.DataFrame(records, columns=OUTPUT_COLS)
    df_out.to_excel(OUTPUT_FILE, index=False)

    elapsed = int((datetime.now() - start_time).total_seconds())
    logger.info("=" * 65)
    logger.info(f"完了: {OUTPUT_FILE} に {len(df_out)} 件を出力")
    logger.info(f"所要時間: {elapsed // 60}分{elapsed % 60}秒")
    logger.info("=" * 65)


if __name__ == "__main__":
    asyncio.run(main())
