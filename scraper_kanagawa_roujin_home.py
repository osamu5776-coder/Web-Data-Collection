"""
神奈川県 老人ホーム情報収集

データソース: 介護サービス情報公表システム（kaigokensaku.mhlw.go.jp）神奈川県版
  Phase 1 - 特別養護老人ホーム（介護老人福祉施設 ServiceCd=510）
            介護事業所検索の JSON API から取得
  Phase 2 - 有料老人ホーム
            生活関連情報検索（SearchType=nursinghome）の HTML から取得
  Phase 3 - サービス付き高齢者向け住宅
            住まい検索（sumai_search）の HTML から取得。電話番号は
            サービス付き高齢者向け住宅情報提供システム
            （satsuki-jutaku.mlit.go.jp）の詳細ページから取得。

対象外: ショートステイ専用・デイサービス専用・グループホーム・訪問介護事業所
       （いずれも上記3種別の検索では取得されないため自然に除外される）

  Phase 4 - 公式サイトURLがある施設について、サイト本文からインスタURL・
            問い合わせフォームURLを取得（requests + BeautifulSoup, 10並列）

出力列: 施設種別, 名称, 所在地, 電話番号, 公式サイトURL, インスタURL, 問い合わせフォームURL
出力ファイル: kanagawa_roujin_home_YYYYMMDD_HHMMSS.xlsx
"""

import asyncio
import io
import json
import logging
import random
import re
import sys
import time
from datetime import datetime

import pandas as pd
import requests
from bs4 import BeautifulSoup

# ── ロギング ──────────────────────────────────────────────
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

LOG_FILE = "scraper_kanagawa_roujin_home.log"
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
OUTPUT_FILE = f"kanagawa_roujin_home_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)
BASE = "https://www.kaigokensaku.mhlw.go.jp/14"
PAGE_SIZE = 50

OUTPUT_COLS = [
    "施設種別", "名称", "所在地", "電話番号", "公式サイトURL",
    "インスタURL", "問い合わせフォームURL",
]


def _new_session() -> requests.Session:
    s = requests.Session()
    s.headers.update({"User-Agent": UA})
    return s


def _clean(s: str) -> str:
    return re.sub(r"[\s　]+", " ", str(s or "")).strip()


# ── Phase 1: 特別養護老人ホーム（JSON API） ──────────────

def fetch_tokuyou(session: requests.Session) -> list[dict]:
    logger.info("Phase 1: 特別養護老人ホーム（介護老人福祉施設）を取得")

    session.get(
        f"{BASE}/index.php?action_kouhyou_pref_search_condition_index=true",
        headers={"Referer": f"{BASE}/index.php?action_kouhyou_pref_topjigyosyo_index=true"},
        timeout=20,
    )
    conditions = json.dumps({"ServiceCd": ["510"]})
    session.post(
        f"{BASE}/index.php?action_kouhyou_pref_search_list_list=true",
        data={
            "method": "search",
            "action_kouhyou_pref_search_condition_index": "true",
            "PrefCd": "14",
            "FromPage": "kaigoConditionSearchPage",
            "SearchConditions": conditions,
        },
        headers={"Referer": f"{BASE}/index.php?action_kouhyou_pref_search_condition_index=true"},
        timeout=20,
    )

    records: list[dict] = []
    offset = 0
    total = None
    while total is None or offset < total:
        r = session.get(
            f"{BASE}/index.php",
            params={
                "action_kouhyou_pref_search_list_list": "true",
                "action_kouhyou_pref_search_search": "true",
                "method": "search",
                "p_count": str(PAGE_SIZE),
                "p_offset": str(offset),
                "p_sort_name": "",
                "p_order": "0",
            },
            headers={"Referer": f"{BASE}/index.php?action_kouhyou_pref_search_list_list=true"},
            timeout=20,
        )
        data = r.json()
        if data.get("status") != "success":
            break
        if total is None:
            total = data["pager"]["total"]
            logger.info(f"  対象: {total} 件")

        for item in data["data"]:
            name = _clean(item.get("JigyosyoName", ""))
            addr = _clean(item.get("JigyosyoJyusho", ""))
            zipcode = _clean(item.get("JigyosyoYubinbangou", "")).replace(",", "-")
            address = f"〒{zipcode} {addr}" if zipcode else addr
            phone = _clean(item.get("JigyosyoTel", ""))
            url = _clean(item.get("JHPUrl", "")) if str(item.get("UrlLinkFlag")) == "1" else ""

            records.append({
                "施設種別": "特別養護老人ホーム",
                "名称": name,
                "所在地": address,
                "電話番号": phone,
                "公式サイトURL": url,
            })

        offset += PAGE_SIZE
        logger.info(f"  取得: {min(offset, total)}/{total}")
        time.sleep(random.uniform(0.5, 1.0))

    logger.info(f"Phase 1 完了: {len(records)} 件")
    return records


# ── Phase 2: 有料老人ホーム（HTML） ───────────────────────

def _parse_shisetsu_list(html: str) -> list[dict]:
    soup = BeautifulSoup(html, "lxml")
    records: list[dict] = []
    for li in soup.select("li.listLi_shisetsu"):
        name_div = li.select_one("div.sel_jigyosyoName")
        name = _clean(name_div.get_text()) if name_div else ""
        if not name:
            continue

        addr_td = li.select_one("td.listAddress_shisetsu")
        zipcode = ""
        address = ""
        if addr_td:
            postal_el = addr_td.select_one("span.postalCode")
            if postal_el:
                zipcode = _clean(postal_el.get_text()).lstrip("〒")
                postal_el.decompose()
            for a in addr_td.select("a"):
                a.decompose()
            address = _clean(addr_td.get_text())

        tel_td = li.select_one("td.listTel_shisetsu")
        phone = _clean(tel_td.get_text()) if tel_td else ""

        url = ""
        onclick_el = li.select_one("a[onclick*=\"window.open\"]")
        if onclick_el:
            m = re.search(r"window\.open\('([^']+)'", onclick_el.get("onclick", ""))
            if m:
                url = m.group(1)

        full_address = f"〒{zipcode} {address}" if zipcode else address
        records.append({
            "施設種別": "有料老人ホーム",
            "名称": name,
            "所在地": full_address,
            "電話番号": phone,
            "公式サイトURL": url,
        })
    return records


def fetch_yuuryou(session: requests.Session) -> list[dict]:
    logger.info("Phase 2: 有料老人ホームを取得")

    session.get(
        f"{BASE}/index.php?action_kouhyou_pref_seikatu_search_list_list=true&PrefCd=14&SearchType=nursinghome",
        timeout=20,
    )

    records: list[dict] = []
    offset = 0
    total = None
    while total is None or offset < total:
        r = session.post(
            f"{BASE}/index.php?iframe=1",
            data={
                "method": "result", "PrefCd": "14", "OriPrefCd": "14",
                "SearchType": "nursinghome",
                "p_offset": str(offset), "p_count": str(PAGE_SIZE),
                "p_sort": "0", "p_order": "0",
                "action_kouhyou_pref_seikatu_search_list_list": "true",
            },
            headers={"Referer": f"{BASE}/index.php?action_kouhyou_pref_seikatu_search_list_list=true&PrefCd=14&SearchType=nursinghome"},
            timeout=20,
        )
        html = r.text.replace("\x00", "")
        m = re.search(r"対象事業所数：<span>(\d+)件", html)
        if total is None:
            if not m:
                break
            total = int(m.group(1))
            logger.info(f"  対象: {total} 件")

        recs = _parse_shisetsu_list(html)
        records.extend(recs)
        offset += PAGE_SIZE
        logger.info(f"  取得: {min(offset, total)}/{total}")
        time.sleep(random.uniform(0.5, 1.0))

    logger.info(f"Phase 2 完了: {len(records)} 件")
    return records


# ── Phase 3: サービス付き高齢者向け住宅（HTML + 詳細ページ） ──

def _parse_sumai_list(html: str) -> list[dict]:
    soup = BeautifulSoup(html, "lxml")
    records: list[dict] = []
    for li in soup.select("li.listLi_sumai"):
        name_div = li.select_one("div.sel_jigyosyoName")
        name = _clean(name_div.get_text()) if name_div else ""
        if not name:
            continue

        address = ""
        for th in li.find_all("th"):
            if th.get_text(strip=True) == "所在地":
                td = th.find_next_sibling("td")
                if td:
                    a = td.select_one("a")
                    if a:
                        a.decompose()
                    address = _clean(td.get_text())
                break

        detail_url = ""
        detail_a = li.select_one("div.btn_sumai_detail a[href]")
        if detail_a:
            detail_url = detail_a["href"]

        records.append({
            "施設種別": "サービス付き高齢者向け住宅",
            "名称": name,
            "所在地": address,
            "電話番号": "",
            "公式サイトURL": "",
            "_detail_url": detail_url,
        })
    return records


def fetch_sakoujuu(session: requests.Session) -> list[dict]:
    logger.info("Phase 3: サービス付き高齢者向け住宅を取得")

    session.get(
        f"{BASE}/index.php?action_kouhyou_pref_sumai_search_list_list=true&PrefCd=14",
        timeout=20,
    )

    records: list[dict] = []
    offset = 0
    total = None
    while total is None or offset < total:
        r = session.post(
            f"{BASE}/index.php?iframe=1",
            data={
                "method": "result", "PrefCd": "14", "OriPrefCd": "14",
                "p_offset": str(offset), "p_count": str(PAGE_SIZE),
                "p_sort": "0", "p_order": "0",
                "action_kouhyou_pref_sumai_search_list_list": "true",
            },
            headers={"Referer": f"{BASE}/index.php?action_kouhyou_pref_sumai_search_list_list=true&PrefCd=14"},
            timeout=20,
        )
        html = r.text.replace("\x00", "")
        m = re.search(r"対象施設数：<span>(\d+)件", html)
        if total is None:
            if not m:
                break
            total = int(m.group(1))
            logger.info(f"  対象: {total} 件")

        recs = _parse_sumai_list(html)
        records.extend(recs)
        offset += PAGE_SIZE
        logger.info(f"  取得: {min(offset, total)}/{total}")
        time.sleep(random.uniform(0.5, 1.0))

    logger.info(f"Phase 3A 完了（一覧取得）: {len(records)} 件")
    return records


def _fetch_sumai_detail_sync(rec: dict) -> None:
    """satsuki-jutaku.mlit.go.jp の詳細ページから電話番号を取得する。"""
    detail_url = rec.pop("_detail_url", "")
    if not detail_url:
        return
    try:
        r = requests.get(detail_url, headers={"User-Agent": UA}, timeout=12)
        if r.status_code >= 400:
            return
        r.encoding = r.apparent_encoding or "utf-8"
        soup = BeautifulSoup(r.text, "lxml")

        # 「問合せ先１」セルから氏名/名称・電話番号を取得
        for th in soup.find_all("th"):
            if "問合せ先" in th.get_text(strip=True):
                td = th.find_next_sibling("td")
                if td:
                    text = td.get_text(separator="\n")
                    m = re.search(r"電話番号[：:]\s*([\d\-]+)", text)
                    if m and not rec.get("電話番号"):
                        rec["電話番号"] = m.group(1)
                break
    except Exception as e:
        logger.debug(f"サ高住詳細取得失敗: {detail_url} / {e}")


async def enrich_sakoujuu(records: list[dict], sem_count: int = 8) -> None:
    semaphore = asyncio.Semaphore(sem_count)
    total = len(records)
    done = {"n": 0}

    async def _one(rec: dict) -> None:
        async with semaphore:
            await asyncio.to_thread(_fetch_sumai_detail_sync, rec)
            await asyncio.sleep(random.uniform(0.2, 0.4))
            done["n"] += 1
            if done["n"] % 50 == 0 or done["n"] == total:
                logger.info(f"  詳細取得: {done['n']}/{total}")

    await asyncio.gather(*[_one(rec) for rec in records])
    logger.info("Phase 3B 完了（詳細ページから電話番号取得）")


# ── Phase 4: 公式サイトからインスタ・問い合わせフォームURL取得 ──

from urllib.parse import urljoin, urlparse
from urllib.robotparser import RobotFileParser

_robots_cache: dict[str, RobotFileParser] = {}


def _is_allowed(url: str) -> bool:
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
    return _robots_cache[origin].can_fetch(UA, url)


def _parse_official_html(url: str, html: str) -> dict:
    result: dict = {}
    soup = BeautifulSoup(html, "lxml")
    base_host = urlparse(url).netloc

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
        if href.startswith("//"):
            href = "https:" + href
        elif not href.startswith(("http", "#", "mailto:", "tel:")):
            href = urljoin(url, href)
        if href.startswith(("#", "mailto:", "tel:")):
            continue
        # 別ドメイン（ホスティングサービス自身の問い合わせ窓口等）は採用しない
        if urlparse(href).netloc != base_host:
            continue
        result["問い合わせフォームURL"] = href
        break

    return result


def _fetch_official_sync(url: str) -> dict:
    if not url.startswith("http"):
        return {}
    if not _is_allowed(url):
        return {}
    try:
        r = requests.get(url, headers={"User-Agent": UA}, timeout=10, allow_redirects=True)
        if r.status_code >= 400:
            return {}
        r.encoding = r.apparent_encoding or "utf-8"
        return _parse_official_html(r.url, r.text)
    except Exception:
        return {}


async def enrich_official_sites(records: list[dict], sem_count: int = 10) -> None:
    semaphore = asyncio.Semaphore(sem_count)
    with_url = [r for r in records if r.get("公式サイトURL")]
    total = len(with_url)
    done = {"n": 0}
    logger.info(f"Phase 4: 公式サイトからインスタ・問い合わせフォームURL取得 ({total} 件)")

    async def _one(rec: dict) -> None:
        async with semaphore:
            extras = await asyncio.to_thread(_fetch_official_sync, rec["公式サイトURL"])
            rec.update(extras)
            done["n"] += 1
            if done["n"] % 50 == 0 or done["n"] == total:
                logger.info(f"  公式サイト取得: {done['n']}/{total}")

    await asyncio.gather(*[_one(rec) for rec in with_url])
    logger.info("Phase 4 完了")


# ── メイン ────────────────────────────────────────────────

async def main() -> None:
    logger.info("=" * 65)
    logger.info("神奈川県 老人ホーム情報収集開始")
    logger.info("対象: 特別養護老人ホーム / 有料老人ホーム / サービス付き高齢者向け住宅")
    logger.info("=" * 65)
    start_time = datetime.now()

    session = _new_session()

    all_records: list[dict] = []
    all_records.extend(fetch_tokuyou(session))
    all_records.extend(fetch_yuuryou(session))

    sakoujuu_records = fetch_sakoujuu(session)
    await enrich_sakoujuu(sakoujuu_records, sem_count=8)
    for rec in sakoujuu_records:
        rec.pop("_detail_url", None)
    all_records.extend(sakoujuu_records)

    for rec in all_records:
        rec.setdefault("インスタURL", "")
        rec.setdefault("問い合わせフォームURL", "")

    await enrich_official_sites(all_records, sem_count=10)

    df = pd.DataFrame(all_records, columns=OUTPUT_COLS)
    df.drop_duplicates(subset=["施設種別", "名称", "所在地"], keep="first", inplace=True)
    df = df.reset_index(drop=True)
    df.to_excel(OUTPUT_FILE, index=False)

    elapsed = int((datetime.now() - start_time).total_seconds())
    logger.info("=" * 65)
    logger.info(f"完了: {OUTPUT_FILE} に {len(df)} 件を出力")
    for t, cnt in df["施設種別"].value_counts().items():
        logger.info(f"  {t}: {cnt} 件")
    logger.info(f"  インスタURL取得: {(df['インスタURL'] != '').sum()} 件")
    logger.info(f"  問い合わせフォームURL取得: {(df['問い合わせフォームURL'] != '').sum()} 件")
    logger.info(f"所要時間: {elapsed // 60}分{elapsed % 60}秒")
    logger.info("=" * 65)


if __name__ == "__main__":
    asyncio.run(main())
