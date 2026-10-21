"""
京都府 老人ホーム情報収集

データソース:
  介護サービス情報公表システム（kaigokensaku.mhlw.go.jp）京都府版
    Phase 1 - 入所・居住系の介護保険施設（介護事業所検索の JSON API）
              特養(510/540)・特定施設の介護付有料老人ホーム(331/335/361)・
              特定施設のサ高住(334/337/364)
    Phase 2 - 有料老人ホーム（生活関連情報検索 SearchType=nursinghome の HTML）
    Phase 3 - サービス付き高齢者向け住宅（住まい検索の HTML。電話番号は
              satsuki-jutaku.mlit.go.jp の詳細ページから取得）
  みんかい（minkai.jp/kyoto）
    Phase 5 - 一覧ページ（約47ページ）から名称・種別・所在地を取得し、
              上記と名寄せして未掲載の施設（住宅型有料老人ホーム等）を追加
              （特養・有料老人ホーム・サ高住のみ）

  Phase 6 - 電話番号が未取得の施設は DuckDuckGo で公式サイトを検索し、名称と
            市区町村・町名がページ内にあることを確認した上で URL・電話番号を採用
  Phase 4 - 公式サイトURLがある施設について、サイト本文からメールアドレス・
            インスタURL・問い合わせフォームURLを取得（10並列）

名寄せ: 同一市区町村内で、正規化名称の一致、または住所（番地まで）一致かつ
        名称の包含関係で判定。「第二」「Ⅱ」「○号館」や特養/グループホーム等の
        区分違いは別施設として扱う。

対象外: グループホーム・介護老人保健施設・介護医療院・軽費老人ホーム・
        ショートステイ専用・デイサービス・訪問介護事業所（神奈川版と同じ対象範囲）

出力列: 名称, メールアドレス, 公式サイトURL, 所在地, 電話番号, インスタURL, 問い合わせフォームURL
出力ファイル: kyoto_roujin_home_YYYYMMDD_HHMMSS.xlsx
"""

import asyncio
import io
import json
import logging
import random
import re
import sys
import time
import unicodedata
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

LOG_FILE = "scraper_kyoto_roujin_home.log"
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
# 検索ライブラリ（ddgs）の通信ログでログファイルが埋まらないようにする
for _name in ("httpx", "primp", "ddgs", "urllib3"):
    logging.getLogger(_name).setLevel(logging.WARNING)

# ── 定数 ─────────────────────────────────────────────────
OUTPUT_FILE = f"kyoto_roujin_home_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)
BASE = "https://www.kaigokensaku.mhlw.go.jp/26"
PAGE_SIZE = 50

OUTPUT_COLS = [
    "名称", "メールアドレス", "公式サイトURL",
    "所在地", "電話番号", "インスタURL", "問い合わせフォームURL",
]


def _new_session() -> requests.Session:
    s = requests.Session()
    s.headers.update({"User-Agent": UA})
    return s


_KYOTO_AREA4 = ("0771", "0772", "0773", "0774", "0743")


def _format_phone(s: str) -> str:
    """全角・各種ダッシュを正規化し、ハイフン区切りに統一する。"""
    s = unicodedata.normalize("NFKC", str(s or ""))
    m = re.search(r"0\d[\d\-‐ー−–()（） ]{7,14}\d", s)
    if not m:
        return s.strip()
    raw = m.group(0)
    digits = re.sub(r"\D", "", raw)
    parts = [p for p in re.split(r"[^\d]+", raw) if p]
    if len(parts) == 3 and len(digits) in (10, 11):
        return "-".join(parts)
    if len(digits) == 10:
        if digits.startswith(_KYOTO_AREA4) or digits.startswith("0120"):
            return f"{digits[:4]}-{digits[4:6]}-{digits[6:]}"
        if digits.startswith("075"):
            return f"{digits[:3]}-{digits[3:6]}-{digits[6:]}"
        if digits.startswith(("03", "06")):
            return f"{digits[:2]}-{digits[2:6]}-{digits[6:]}"
    if len(digits) == 11 and digits[:3] in ("070", "080", "090"):
        return f"{digits[:3]}-{digits[3:7]}-{digits[7:]}"
    return "-".join(parts) if parts else digits


def _clean(s: str) -> str:
    return re.sub(r"[\s　]+", " ", str(s or "")).strip()


# ── Phase 1: 入所・居住系介護施設（JSON API） ──────────────

# 介護サービスコード → 出力上の施設種別
KAIGO_SERVICES = [
    ("510", "特別養護老人ホーム"),
    ("540", "特別養護老人ホーム（地域密着型）"),
    ("331", "介護付有料老人ホーム"),
    ("335", "介護付有料老人ホーム（外部サービス利用型）"),
    ("361", "介護付有料老人ホーム（地域密着型）"),
    ("334", "サービス付き高齢者向け住宅（特定施設）"),
    ("337", "サービス付き高齢者向け住宅（特定施設・外部サービス利用型）"),
    ("364", "サービス付き高齢者向け住宅（特定施設・地域密着型）"),
]


def fetch_kaigo_all(session: requests.Session) -> list[dict]:
    records: list[dict] = []
    for code, label in KAIGO_SERVICES:
        records.extend(fetch_kaigo_service(session, code, label))
    return records


def fetch_kaigo_service(session: requests.Session, service_cd: str, label: str) -> list[dict]:
    logger.info(f"Phase 1: {label}（ServiceCd={service_cd}）を取得")

    session.get(
        f"{BASE}/index.php?action_kouhyou_pref_search_condition_index=true",
        headers={"Referer": f"{BASE}/index.php?action_kouhyou_pref_topjigyosyo_index=true"},
        timeout=20,
    )
    conditions = json.dumps({"ServiceCd": [service_cd]})
    session.post(
        f"{BASE}/index.php?action_kouhyou_pref_search_list_list=true",
        data={
            "method": "search",
            "action_kouhyou_pref_search_condition_index": "true",
            "PrefCd": "26",
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
            address = _clean(item.get("JigyosyoJyusho", ""))
            phone = _clean(item.get("JigyosyoTel", ""))
            url = _clean(item.get("JHPUrl", "")) if str(item.get("UrlLinkFlag")) == "1" else ""

            records.append({
                "施設種別": label,
                "名称": name,
                "所在地": address,
                "電話番号": phone,
                "公式サイトURL": url,
            })

        offset += PAGE_SIZE
        logger.info(f"  取得: {min(offset, total)}/{total}")
        time.sleep(random.uniform(0.5, 1.0))

    logger.info(f"  {label} 完了: {len(records)} 件")
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
        address = ""
        if addr_td:
            postal_el = addr_td.select_one("span.postalCode")
            if postal_el:
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

        records.append({
            "施設種別": "有料老人ホーム",
            "名称": name,
            "所在地": address,
            "電話番号": phone,
            "公式サイトURL": url,
        })
    return records


def fetch_yuuryou(session: requests.Session) -> list[dict]:
    logger.info("Phase 2: 有料老人ホームを取得")

    session.get(
        f"{BASE}/index.php?action_kouhyou_pref_seikatu_search_list_list=true&PrefCd=26&SearchType=nursinghome",
        timeout=20,
    )

    records: list[dict] = []
    offset = 0
    total = None
    while total is None or offset < total:
        r = session.post(
            f"{BASE}/index.php?iframe=1",
            data={
                "method": "result", "PrefCd": "26", "OriPrefCd": "26",
                "SearchType": "nursinghome",
                "p_offset": str(offset), "p_count": str(PAGE_SIZE),
                "p_sort": "0", "p_order": "0",
                "action_kouhyou_pref_seikatu_search_list_list": "true",
            },
            headers={"Referer": f"{BASE}/index.php?action_kouhyou_pref_seikatu_search_list_list=true&PrefCd=26&SearchType=nursinghome"},
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
        f"{BASE}/index.php?action_kouhyou_pref_sumai_search_list_list=true&PrefCd=26",
        timeout=20,
    )

    records: list[dict] = []
    offset = 0
    total = None
    while total is None or offset < total:
        r = session.post(
            f"{BASE}/index.php?iframe=1",
            data={
                "method": "result", "PrefCd": "26", "OriPrefCd": "26",
                "p_offset": str(offset), "p_count": str(PAGE_SIZE),
                "p_sort": "0", "p_order": "0",
                "action_kouhyou_pref_sumai_search_list_list": "true",
            },
            headers={"Referer": f"{BASE}/index.php?action_kouhyou_pref_sumai_search_list_list=true&PrefCd=26"},
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


# ── Phase 5: みんかい（minkai.jp/kyoto）で漏れを補完 ──────

MINKAI_BASE = "https://minkai.jp"
# フルのブラウザUAだとブラウザ検証ページ(405)が返るため、簡素なUAで取得する
MINKAI_UA = "Mozilla/5.0"
# 神奈川版と同じ対象（グループホーム・老健・介護医療院・軽費老人ホーム等は除外）
MINKAI_KIND_RE = re.compile(r"特別養護|有料老人ホーム|サービス付き")


def _parse_minkai_list(html: str) -> list[dict]:
    soup = BeautifulSoup(html, "lxml")
    records: list[dict] = []
    for card in soup.select("div.c-card.js-jigyosho-card"):
        title = card.select_one("h3.c-card__title")
        name = _clean(title.get_text()) if title else ""
        if not name:
            continue
        label = card.select_one("div.c-card__main__labels label")
        kind = _clean(label.get_text()) if label else ""
        address = ""
        for head in card.select("div.c-jigyosho-summary__head"):
            if _clean(head.get_text()) == "住所":
                main = head.find_next_sibling("div")
                if main:
                    lines = [l.strip() for l in main.get_text("\n").split("\n") if l.strip()]
                    address = lines[0] if lines else ""
                break
        records.append({
            "施設種別": kind,
            "名称": name,
            "所在地": address,
            "電話番号": "",
            "公式サイトURL": "",
            "_minkai_id": card.get("data-jigyosho-id", ""),
        })
    return records


def fetch_minkai(session: requests.Session) -> list[dict]:
    logger.info("Phase 5: みんかい（minkai.jp/kyoto）を取得")
    records: list[dict] = []
    page = 1
    while True:
        url = f"{MINKAI_BASE}/kyoto" + (f"?page={page}" if page > 1 else "")
        try:
            r = session.get(url, headers={"User-Agent": MINKAI_UA}, timeout=20)
        except Exception as e:
            logger.warning(f"  minkai 取得失敗: {url} / {e}")
            break
        if r.status_code >= 400:
            break
        recs = _parse_minkai_list(r.text)
        if not recs:
            break
        # 種別での絞り込みは終端判定の後に行う（対象外種別のみのページで打ち切らないため）
        records.extend(rec for rec in recs if MINKAI_KIND_RE.search(rec["施設種別"]))
        if page % 10 == 0:
            logger.info(f"  minkai {page} ページ目: 累計 {len(records)} 件")
        page += 1
        time.sleep(random.uniform(1.0, 1.5))
    logger.info(f"Phase 5 完了: {len(records)} 件")
    return records


_TYPE_PREFIX_RE = re.compile(
    r"^(介護付き?|住宅型|健康型)?有料老人ホーム|^サービス付き?高齢者(向け)?住宅|"
    r"^特別養護老人ホーム|^介護老人保健施設|^グループホーム|^ケアハウス"
)


def _norm_name(s: str) -> str:
    s = unicodedata.normalize("NFKC", s or "")
    s = re.sub(r"[(（]\s*(旧名称|仮|特定施設).*?[)）]", "", s)
    s = re.sub(r"【.*?】", "", s)
    s = re.sub(r"\s", "", s).replace("サ-ビス", "サービス")
    s = _TYPE_PREFIX_RE.sub("", s) or s
    s = re.sub(r"[・･\-－ー‐「」]", "", s)
    return s.lower()


_KANJI_NUM = {c: str(i) for i, c in enumerate("〇一二三四五六七八九")}


def _norm_addr(s: str) -> str:
    """住所を番地までに正規化する（漢数字の丁目・区切り記号・建物名の差を吸収）。"""
    s = unicodedata.normalize("NFKC", s or "")
    s = re.sub(r"\s", "", s).replace("京都府", "").replace("ケ", "ヶ")
    s = re.sub(r"([一二三四五六七八九])丁目", lambda m: _KANJI_NUM[m.group(1)] + "丁目", s)
    s = re.sub(r"(丁目|番地の?|番|号)", "-", s)
    s = re.sub(r"[－ー‐−ｰ]", "-", s)
    m = re.match(r"^(.*?\d+(?:-\d+)*)", s)
    if m:
        s = m.group(1)
    return re.sub(r"-+$", "", re.sub(r"-+", "-", s))


# 名称の差分にこれらが含まれる場合は、同一建物・同一法人でも別施設として扱う（_norm_name 後の表記で照合）
_DISTINCT_RE = re.compile(r"^[ivx]+$|[0-9第二弐参三]|号館|番館|館|ユニット|従来|サテライト|グルプホム|ケアハウス|特別養護|老人保健")


def _city(addr: str) -> str:
    a = unicodedata.normalize("NFKC", addr or "").replace("京都府", "").strip()
    m = re.match(r"(京都市..?区|.{1,4}?郡.{1,4}?[町村]|.{1,4}?[市町村])", a)
    return m.group(1) if m else ""


def _kind_group(kind: str) -> str:
    """介護保険上の区分が明確な施設種別をグループ化する（住宅系はソース間で表記が揺れるため1グループ）。"""
    for key, group in (("グループホーム", "GH"), ("特別養護", "特養"), ("老人保健", "老健"),
                       ("医療院", "医療院")):
        if key in kind:
            return group
    return "住宅系"


def _same_facility(a: dict, b: dict) -> bool:
    if _city(a["所在地"]) != _city(b["所在地"]):
        return False
    # 同一法人が同名で特養とグループホームを併設する例が多いため、区分が異なれば別施設
    if _kind_group(a["施設種別"]) != _kind_group(b["施設種別"]):
        return False
    na, nb = _norm_name(a["名称"]), _norm_name(b["名称"])
    if na and na == nb:
        return True
    # 名称が異なる場合は、住所一致（片方が番地なしなら前方一致）かつ名称の包含関係を要求する
    aa, ab = _norm_addr(a["所在地"]), _norm_addr(b["所在地"])
    if not (aa and ab):
        return False
    exact = aa == ab
    prefix = (aa.startswith(ab) or ab.startswith(aa)) and len(min(aa, ab, key=len)) > len(_city(aa)) + 1
    if not (exact or prefix):
        return False
    short, long_ = sorted([na, nb], key=len)
    if len(short) < (3 if exact else 5) or short not in long_:
        return False
    return not _DISTINCT_RE.search(long_.replace(short, "", 1))


def merge_records(base: list[dict], extra: list[dict]) -> tuple[list[dict], int]:
    """extra のうち base と同一施設でないものを追加する。同一施設は空欄項目のみ補完。"""
    merged = list(base)
    added = 0
    by_city: dict[str, list[dict]] = {}
    for r in merged:
        by_city.setdefault(_city(r["所在地"]), []).append(r)
    name_idx = {(_city(r["所在地"]), _norm_name(r["名称"])): r for r in merged}
    for r in extra:
        hit = name_idx.get((_city(r["所在地"]), _norm_name(r["名称"])))
        if hit is not None and not _same_facility(hit, r):
            hit = None
        if hit is None:
            for cand in by_city.get(_city(r["所在地"]), []):
                if _same_facility(cand, r):
                    hit = cand
                    break
        if hit is not None:
            for k, v in r.items():
                if v and not hit.get(k) and not k.startswith("_"):
                    hit[k] = v
            continue
        merged.append(r)
        name_idx[(_city(r["所在地"]), _norm_name(r["名称"]))] = r
        by_city.setdefault(_city(r["所在地"]), []).append(r)
        added += 1
    return merged, added


# ── Phase 6: 電話番号未取得の施設を DuckDuckGo で公式サイト検索し補完 ──

SKIP_DOMAINS = [
    "minkai.jp", "kaigokensaku.mhlw.go.jp", "satsuki-jutaku.mlit.go.jp",
    "homes.co.jp", "lifull", "e-ninchisho", "sagasix", "kaigo-postseven",
    "oasis-navi", "minnanokaigo", "kaigodb", "roujin-homes", "s-roujin",
    "job-medley", "kaigojob", "indeed", "townwork", "mynavi", "hellowork",
    "baitoru", "en-gage", "careerindex", "glassdoor", "wantedly",
    "mapion", "navitime", "itp.ne.jp", "ekiten", "google.", "yahoo.", "bing.com",
    "facebook.com", "instagram.com", "twitter.com", "x.com", "youtube.com",
    "wikipedia", "jpnumber", "telnavi", "houjin", "baseconnect",
    "pref.kyoto", ".lg.jp", "city.", "rakuten", "suumo", "kaigonohonne",
]


# 介護施設ポータル・求人・医療情報サイト等のドメインに頻出する語
_SKIP_HOST_RE = re.compile(
    r"job|medicaldoc|co-medical|kyujin|recruit|career|caree|nurse|navi|kaigo-?home|nursingcare|kaigo-?db|"
    r"shisetsu|search|guide|portal|review|ranking|hikaku|matome|blog|ameblo|note\.com"
)


def _skip(url: str) -> bool:
    u = url.lower()
    return any(d in u for d in SKIP_DOMAINS) or bool(_SKIP_HOST_RE.search(urlparse(u).netloc))


def _town(addr: str) -> str:
    a = unicodedata.normalize("NFKC", addr or "").replace("京都府", "").strip()
    rest = a[len(_city(a)):]
    return re.sub(r"[\d一二三四五六七八九十〇]+.*$|字.*$", "", rest)


def _page_matches(text: str, rec: dict) -> bool:
    """検索で見つけたページが当該施設のものかを、名称と市区町村・町名で検証する。"""
    page = unicodedata.normalize("NFKC", re.sub(r"\s", "", text))
    core = _norm_name(rec["名称"])
    if len(core) < 2 or core not in _norm_name(page):
        return False
    city, town = _city(rec["所在地"]), _town(rec["所在地"])
    return bool(city) and city in page and (len(town) < 2 or town in page)


def _find_phone(text: str, rec: dict) -> str:
    """ページ内の京都府（075 / 077x / 0743）の電話番号を返す。複数ある場合は町名の記載に最も近いもの。"""
    page = unicodedata.normalize("NFKC", text)
    phones: list[tuple[int, str]] = []
    for m in re.finditer(r"(0\d{1,4})[-‐ー−(\s]{0,2}(\d{1,4})[-‐ー−)\s]{0,2}(\d{4})(?!\d)", page):
        digits = "".join(m.groups())
        if len(digits) == 10 and digits.startswith(("075",) + _KYOTO_AREA4):
            phones.append((m.start(), f"{m.group(1)}-{m.group(2)}-{m.group(3)}"))
    if not phones:
        return ""
    if len({p for _, p in phones}) == 1:
        return phones[0][1]
    town = _town(rec["所在地"])
    key = town if len(town) >= 2 else _city(rec["所在地"])
    anchors = [m.start() for m in re.finditer(re.escape(key), page)] if key else []
    best, best_d = "", 10 ** 9
    for pos, tel in phones:
        for anc in anchors:
            if abs(pos - anc) < best_d:
                best, best_d = tel, abs(pos - anc)
    return best if best_d < 400 else ""


def search_official(rec: dict) -> tuple[str, str]:
    name = re.sub(r"[(（]\s*(旧名称|仮).*?[)）]|【.*?】", "", rec["名称"]).strip()
    query = f"{name} {_city(rec['所在地'])} 京都府"
    try:
        with DDGS() as ddgs:
            results = list(ddgs.text(query, region="jp-jp", max_results=8))
    except Exception as e:
        logger.debug(f"DDG失敗 {name}: {e}")
        return "", ""
    for res in results:
        url = res.get("href", "")
        if not url.startswith("http") or _skip(url) or not _is_allowed(url):
            continue
        try:
            r = requests.get(url, headers={"User-Agent": UA}, timeout=10, allow_redirects=True)
            if r.status_code >= 400 or _skip(r.url):
                continue
            r.encoding = r.apparent_encoding or "utf-8"
        except Exception:
            continue
        text = BeautifulSoup(r.text, "lxml").get_text("\n")
        if _page_matches(text, rec):
            return r.url, text
    return "", ""


async def enrich_by_search(records: list[dict], sem_count: int = 3) -> None:
    targets = [r for r in records if not r.get("電話番号")]
    logger.info(f"Phase 6: 電話番号未取得 {len(targets)} 件を DuckDuckGo で公式サイト検索")
    sem = asyncio.Semaphore(sem_count)
    done = {"n": 0, "url": 0, "tel": 0}

    async def _one(rec: dict) -> None:
        async with sem:
            url, text = await asyncio.to_thread(search_official, rec)
            if url:
                if not rec.get("公式サイトURL"):
                    rec["公式サイトURL"] = url
                done["url"] += 1
                tel = _find_phone(text, rec)
                if tel:
                    rec["電話番号"] = tel
                    done["tel"] += 1
            done["n"] += 1
            if done["n"] % 20 == 0 or done["n"] == len(targets):
                logger.info(f"  検索: {done['n']}/{len(targets)}（公式URL {done['url']} / 電話 {done['tel']}）")
            await asyncio.sleep(random.uniform(1.0, 2.0))

    await asyncio.gather(*[_one(r) for r in targets])
    logger.info("Phase 6 完了")


# ── Phase 4: 公式サイトからインスタ・問い合わせフォームURL取得 ──

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
    base_host = urlparse(url).netloc
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
    logger.info("京都府 老人ホーム情報収集開始")
    logger.info("=" * 65)
    start_time = datetime.now()

    session = _new_session()

    kaigo = fetch_kaigo_all(session)
    yuuryou = fetch_yuuryou(session)
    sakoujuu = fetch_sakoujuu(session)
    await enrich_sakoujuu(sakoujuu, sem_count=8)
    for rec in sakoujuu:
        rec.pop("_detail_url", None)

    all_records, _ = merge_records([], kaigo + yuuryou + sakoujuu)
    logger.info(f"kaigokensaku 名寄せ後: {len(all_records)} 件")

    minkai = fetch_minkai(_new_session())
    all_records, added = merge_records(all_records, minkai)
    logger.info(f"minkai から追加: {added} 件（合計 {len(all_records)} 件）")

    for rec in all_records:
        rec.pop("_minkai_id", None)
        rec.setdefault("メールアドレス", "")
        rec.setdefault("インスタURL", "")
        rec.setdefault("問い合わせフォームURL", "")

    await enrich_by_search(all_records, sem_count=3)
    await enrich_official_sites(all_records, sem_count=10)

    df_full = pd.DataFrame(all_records).fillna("")
    df_full["電話番号"] = df_full["電話番号"].map(_format_phone)
    # 市外局番が省略された京都市内の番号（例: 622-8855）は 075 を補う
    local = df_full["電話番号"].str.fullmatch(r"\d{3}-\d{4}") & df_full["所在地"].str.contains("京都市")
    df_full.loc[local, "電話番号"] = "075-" + df_full.loc[local, "電話番号"]
    df = df_full[OUTPUT_COLS]
    df.to_excel(OUTPUT_FILE, index=False)

    elapsed = int((datetime.now() - start_time).total_seconds())
    logger.info("=" * 65)
    logger.info(f"完了: {OUTPUT_FILE} に {len(df)} 件を出力")
    for t, cnt in df_full["施設種別"].value_counts().items():
        logger.info(f"  {t}: {cnt} 件")
    logger.info(f"  電話番号取得: {(df['電話番号'] != '').sum()} 件")
    logger.info(f"  公式サイトURL取得: {(df['公式サイトURL'] != '').sum()} 件")
    logger.info(f"  メールアドレス取得: {(df['メールアドレス'] != '').sum()} 件")
    logger.info(f"  インスタURL取得: {(df['インスタURL'] != '').sum()} 件")
    logger.info(f"  問い合わせフォームURL取得: {(df['問い合わせフォームURL'] != '').sum()} 件")
    logger.info(f"所要時間: {elapsed // 60}分{elapsed % 60}秒")
    logger.info("=" * 65)


if __name__ == "__main__":
    asyncio.run(main())
