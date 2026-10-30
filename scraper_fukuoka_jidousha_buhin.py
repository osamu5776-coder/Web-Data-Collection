"""
福岡県 自動車部品製造会社リスト作成

データソース:
  Phase 1 - ふくおか自動車部品サプライヤーパーク（joho-fukuoka.or.jp）の福岡県企業一覧から
              各企業の企業概要 PDF を取得し、企業名・所在地・TEL・URL を読む
              （画像だけの PDF は PDF_FIX に目視で読んだ値を持つ）
            NAVITIME 福岡県「自動車部品・カー用品製造」タグ一覧（全ページ）と各スポットの詳細ページ
  Phase 2 - Google マップ「自動車部品製造会社 福岡県」等と主要20市町村別の検索結果（Playwright で一覧を
            最後までスクロールし、各スポットの住所・電話・ウェブサイト・カテゴリを取得）

  Phase 3 - 名寄せ: 法人格・工場/営業所/支店等の表記を除いた社名（コア名）で統合
  Phase 4 - 本社判定: 国税庁法人番号公表サイトでコア名を検索し、
              福岡県内に同名法人あり →  掲載拠点と同じ市区町村に本店がある（同名法人が1社だけなら
                                        市区町村が違っても可）とき本社 福岡県として採用
              県内になし・県外に同名法人あり → 除外（他県本社の工場・営業所）
              法人登録なし（個人事業等）    → 工場/営業所表記がなければ採用
            工場・営業所名義の掲載しかない会社は、本店がその拠点と同じ市区町村にあるときだけ採用
  Phase 7 - 自動車部品の製造・加工以外（部品商・用品販売・整備・解体・物流・ソフト開発・
            設備メーカー・完成車メーカー等）を除外。Google マップにだけ載っている会社は
            カテゴリが製造系（メーカー・製造業者・加工等）のものに限る
  Phase 6 - 公式サイト（トップ＋会社概要・問い合わせページ）からメールアドレス・
            インスタURL・問い合わせフォームURL・電話番号（未取得時）を取得
  Phase 5 - 公式サイトまたは電話番号が未取得の会社は DuckDuckGo で検索し、社名と
            市区町村・町名がページ内にあることを確認して採用

電話番号は phonenumbers で日本の市外局番の境界に従ってハイフンを入れる。

出力列: 名称, メールアドレス, 公式サイトURL, 所在地, 電話番号, インスタURL, 問い合わせフォームURL
出力ファイル: 福岡県_自動車部品製造会社リスト.xlsx
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

LOG_FILE = "scraper_fukuoka_jidousha_buhin.log"
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
OUTPUT_FILE = "福岡県_自動車部品製造会社リスト.xlsx"
CACHE_DIR = os.path.join("data", "fukuoka_jidousha_buhin")
UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36"
)
SUPPLIER_PARK = "https://www.joho-fukuoka.or.jp/k-car-supplierpark/search-area/fukuoka.html"
SUPPLIER_PDF_BASE = "https://www.joho-fukuoka.or.jp/k-car-supplierpark/pdf/"
NAVITIME_BASE = "https://www.navitime.co.jp/category/0516/40/"
NAVITIME_TAG = "010168"  # 自動車部品・カー用品製造
# 未ログインの Google マップは1検索 50〜120 件で打ち切られるため、主要市町村別の検索も加える
GMAPS_QUERIES = ["自動車部品製造会社 福岡県", "自動車部品メーカー 福岡県", "自動車部品工場 福岡県"] + [
    f"自動車部品メーカー 福岡県{city}" for city in (
        "北九州市", "福岡市", "久留米市", "飯塚市", "苅田町", "大牟田市", "宮若市", "糸島市", "古賀市", "筑紫野市",
        "直方市", "行橋市", "朝倉市", "八女市", "柳川市", "田川市", "大野城市", "粕屋町", "小郡市", "中間市",
    )
]
NTA_BASE = "https://www.houjin-bangou.nta.go.jp"
NTA_PREF_FUKUOKA = "40"
# 並列数（相手サイトへの負荷を抑えるため 4 まで）
GMAPS_TABS = 4
NTA_WORKERS = 4

# ── サプライヤーパーク PDF の補正 ───────────────────────────
# 画像だけの PDF・レイアウトが崩れて読めない PDF は目視で読んだ値を使う
PDF_FIX: dict[str, dict] = {
    "kyushu-denka.pdf": {"name": "株式会社九州電化", "addr": "福岡県福岡市東区社領3-4-8", "tel": "092-611-3461", "web": "https://www.k-denka.co.jp/"},
    "tobata-tekko.pdf": {"name": "戸畑鉄工株式会社", "addr": "福岡県北九州市小倉南区新曽根9番20号", "tel": "093-471-7955", "web": "http://www.tobata-iron.co.jp/"},
    "ybm.pdf": {"name": "株式会社ワイ・ビー・エム", "addr": "福岡県飯塚市口原1000-1", "tel": "0948-92-8800", "web": "http://pipe-ybm.co.jp"},
    "nihon-fine-tech.pdf": {"name": "日本ファインテック株式会社", "addr": "福岡県宮若市上有木1438-4", "tel": "0949-34-7780"},
    "rix.pdf": {"name": "リックス株式会社", "addr": "福岡県福岡市博多区山王1-15-15", "tel": "092-472-7311"},
    "dex.pdf": {"name": "株式会社デックス", "addr": "福岡県福岡市博多区住吉2丁目3-7-1305", "tel": "092-600-1913", "web": "http://www.dex-e.com"},
    "mitoyo.pdf": {"name": "株式会社ミトヨ 九州事業所", "addr": "福岡県田川市伊加利1824-17", "tel": "0947-46-5003"},
    "suou-gakuen-seikei-kojyo.pdf": {"name": "社会福祉法人周防学園 成形工場", "addr": "福岡県豊前市大字大村452-1", "tel": "0979-83-1338"},
    "alpha-matal.pdf": {"name": "株式会社アルファメタル", "addr": "福岡県宮若市稲光850番地", "tel": "0949-52-3355"},
    "metalx-japan.pdf": {"name": "株式会社メタリックスジャパン", "addr": "福岡県福岡市博多区博多駅東2-2-2 博多東ハニービル5F", "tel": "092-415-3151"},
    "watanabe-tekkou.pdf": {"name": "渡辺鉄工株式会社", "addr": "福岡県福岡市博多区相生町1-2-1", "tel": "092-581-0331"},
    "exenishimura.pdf": {"name": "株式会社エクセニシムラ", "addr": "福岡県北九州市戸畑区千防3丁目3-25", "tel": "093-881-5255"},
    "fukunetsu.pdf": {"name": "株式会社フクネツ", "addr": "福岡県糟屋郡篠栗町大字和田1034-8", "tel": "092-947-5652"},
    "tox.pdf": {"name": "トックス プレソテクニック株式会社", "addr": "福岡県糟屋郡須恵町佐谷2047-177", "tel": "092-934-4888", "web": "https://tox.com/ja-jp"},
    "aporojapan.pdf": {"name": "株式会社アポロジャパン", "addr": "福岡県北九州市若松区ひびきの1-8", "web": "http://www.sdl-apollo.com"},
    "nihon-biniron.pdf": {"name": "日本ビニロン株式会社 行橋工場", "addr": "福岡県行橋市大字稲童切堤237", "tel": "0930-23-8060", "web": "http://www.nvk.co.jp/"},
    "mituimiikeseisakusyo.pdf": {"name": "株式会社三井三池製作所 精密機器事業本部", "addr": ""},
    "atinde.pdf": {"name": "株式会社ATINDE", "addr": "福岡県行橋市泉中央3-4-9", "tel": "050-1572-6304"},
    "tada-seiki.pdf": {"name": "株式会社福岡多田精機", "addr": "福岡県朝倉郡筑前町四三嶋573-1", "tel": "0946-42-8828"},
    "toyosutenresu.pdf": {"name": "東洋ステンレス研磨工業株式会社"},
    "gokuden.pdf": {"name": "株式会社ゴク・デン", "addr": "福岡県福岡市博多区東那珂2丁目19番39号", "tel": "092-431-3470"},
}
# 企業名は県内でも、掲載されているのが他社本社の工場・営業所であるもの（ファイル名 → 拠点名）
PDF_BRANCH = {
    "iwamoto.pdf": "九州工場", "sanko-mino-kojyo.pdf": "福岡耳納工場", "sakaetekuno.pdf": "九州工場",
    "taiseipurasu.pdf": "九州工場", "nippo.pdf": "九州工場", "yawata-metal.pdf": "北九州若松工場",
    "fujipuresu.pdf": "北九州製作所", "hirai-seimitsu.pdf": "九州営業所", "nichiryo-sangyo.pdf": "福岡営業所",
    "tsuchiya.pdf": "福岡事業所", "shouwa-kinzokukogyo.pdf": "福岡工場", "sakaguchi-dennetu.pdf": "九州営業所",
    "tokyo-electron-device.pdf": "福岡営業所", "dts-insight.pdf": "九州オフィス", "core.pdf": "九州カンパニー",
    "siima-densi.pdf": "九州営業所", "tekunoasosie.pdf": "福岡営業所", "i-pex.pdf": "大野城事業所",
}

# ── 自動車部品製造の判定 ─────────────────────────────────────
# 社名から製造・加工以外が本業とわかる語
_NON_MFG_NAME_RE = re.compile(
    r"商会|商事|商店|商産|販売|物流|ロジ|整備|リサイクル|ガレージ|モータース|ショップ|"
    r"ホールディングス|建設|資材|組$|サービス|トレーディング|鉄鋼センター|コイルセンター|"
    r"カービューティ|ソフトウェア|自動車九州|デザイン$|設計$|テント|石油|建材|タイヤ|(?<!特殊)自動車$|Automobiles|"
    r"硝子$|グラスター|バッテリー|ペイント|紙器|新幹線|ENERGY|エナジー|ゲストホール"
)
# 事業内容（企業概要 PDF・公式サイト）で自動車部品の製造・加工以外が本業と確認した会社（コア名）
NON_MFG = {
    # サプライヤーパーク掲載のうち、ソフト開発・IT・商社・生産設備メーカー等
    "エフェクト", "ファインディックス", "ネクストシステム", "ビゼスト", "モビテック", "デックス",
    "デンシン・ソフトウェア", "アポロジャパン", "ATINDE", "エイチ・アイ・デー", "九州テン", "サン電子",
    "九州リオン", "リックス", "日本ファインテック", "オートシステム", "カナエ", "ティ・アイ・エス",
    "八光オートメーション", "光製作所", "森山鉄工", "渡辺鉄工", "シノハラ工業", "協和商工",
    "トックスプレソテクニック", "熱産ヒート", "エクセニシムラ", "周防学園",
    # NAVITIME / Google マップ掲載のうち、完成車メーカー・部品商・用品販売・修理等
    "トヨタ自動車九州", "日産自動車九州", "国鉄資材", "総合電装", "柳生電装", "八幡ラヂエーター工業",
    "綜合部品", "日本電測機", "三菱電機", "三菱電機パワーデバイス製作所",
    "TCSジャパン福岡",  # レンタカー
    "マッハ",            # 洗車・コーティングのフランチャイズ店
    "ストーク",          # 中古車販売
    # 産業用ロボット・工作機械・農業機械等のメーカー（自動車部品ではない）
    "安川マニュファクチャリング", "西部電機", "筑水キャニコム",
}
# Google マップのカテゴリで製造・加工とみなすもの（Google マップにだけ載っている会社に適用）
_MFG_CATEGORY_RE = re.compile(r"メーカー|製造|加工|金型|工場$|製作所|鋳造|鍛造|めっき|成形|企業のオフィス")
_NON_MFG_CATEGORY_RE = re.compile(r"整備|修理|用品店|部品商|部品店|専門店|販売|解体|中古|ショップ|ディーラー|自動車工場")

# 検索で誤って別の会社・無関係のページが公式サイトURLとして入るもの
WRONG_URLS: tuple[str, ...] = (
    "matomame.jp", "livecamera.", "kensetumap.com", "tsukulink.net", "tellogs.com", "goo-net.com",
    "carnext.jp", "postmap.org", "voix.jp", "glarity.app", "hotfrog.jp", "telinquire.com",
    "joho-fukuoka.or.jp", "jetro.go.jp", "ipros.jp", "metoree.com", "aperza.com",
    "bing.com", "f-dennou.jp", "tanushimaru.or.jp", "kyoka-db.com", "big-advance.site", "culebee.com",
    "atengineer.com", "car-premium.net", "rental-car.jp", "mach5.jp",
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


def is_manufacturer(core: str, name: str, recs: list[dict]) -> bool:
    if core in NON_MFG or _NON_MFG_NAME_RE.search(_LEGAL_RE.sub("", _nfkc(name)).strip()):
        return False
    if any(r["src"] != "Googleマップ" for r in recs):
        return True  # サプライヤーパーク・NAVITIME の製造カテゴリに掲載
    cats = [r.get("category", "") for r in recs]
    return any(_MFG_CATEGORY_RE.search(c) and not _NON_MFG_CATEGORY_RE.search(c) for c in cats)


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


_FUKUOKA_MUNI = (
    r"^(北九州市|福岡市|大牟田市|久留米市|直方市|飯塚市|田川市|柳川市|八女市|筑後市|大川市|"
    r"行橋市|豊前市|中間市|小郡市|筑紫野市|春日市|大野城市|宗像市|太宰府市|古賀市|福津市|"
    r"うきは市|宮若市|嘉麻市|朝倉市|みやま市|糸島市|那珂川市|"
    r"糟屋郡|遠賀郡|鞍手郡|嘉穂郡|朝倉郡|三井郡|三潴郡|八女郡|田川郡|京都郡|築上郡|糸島郡)"
)


def _norm_addr(s: str) -> str:
    """所在地を「福岡県…」形式に整える（郵便番号・「住所:」・拠点の括弧書きを除去）。"""
    a = _nfkc(s)
    a = re.sub(r"^住所:\s*", "", a)
    a = re.sub(r"^[(（【]?(?:本社・?工場|本社)[)）】]?[:：]?", "", a)
    a = re.sub(r"〒?\s*\d{3}-?\d{4}\s*", "", a).strip()
    a = re.sub(r"[(（](?:九州工場|九州営業所|戸畑工場|テクノパーク内)[)）]$", "", a).strip()
    # 「本社:福岡県…」「【九州営業所】福岡県…」のような見出しは落とす（前に他県の住所があるものは残す）
    i = a.find("福岡県")
    if i > 0 and not re.search(r"[都道府]|.県", a[:i]):
        a = a[i:]
    if not a.startswith("福岡県") and re.match(_FUKUOKA_MUNI, a):
        a = "福岡県" + a
    return re.sub(r"\s+", " ", a)


# ── 社名の正規化（コア名） ─────────────────────────────────

_LEGAL_RE = re.compile(
    r"株式会社|有限会社|合同会社|合資会社|合名会社|一般社団法人|社会福祉法人|"
    r"\((?:株|有|同|資|名)\)|㈱|㈲"
)
# 「九州鉄鋼センター」「深江工作所」等は社名の一部なので、地名等を伴わない「センター」「製作所」は拠点表記としない
_BRANCH_KW = (
    r"本社工場|分工場|工場|営業所|支店|支社|事業所|事務所|出張所|整備場|整備工場|カンパニー|"
    r"物流センター|配送センター|開発センター|サービスセンター|オフィス|ショールーム|倉庫|事業部"
)
_BRANCH_RE = re.compile(rf"({_BRANCH_KW})")
# 拠点名に使われる地名・区分
_PLACES = (
    "九州|東九州|西日本|福岡|北九州|小倉|八幡|門司|戸畑|若松|苅田|宮若|宮田|直方|飯塚|田川|久留米|大牟田|"
    "行橋|豊前|中間|朝倉|甘木|筑後|八女|柳川|大川|糸島|古賀|宗像|鞍手|小竹|水巻|犀川|勝山|磯浜|殿川|"
    "博多|土井|桧原|田主丸|大野城|耳納|鷹取|第[一二三四五六七八九1-9]|[東西南北]|中央|新|製造|生産|技術|精密機器|マイスター"
)
_SUFFIX_RE = re.compile(rf"^(.+?)(?:{_PLACES})*(?:{_BRANCH_KW})(?:地区|.{{0,6}})?$")


def core_name(name: str) -> str:
    """法人格・括弧書き・工場/営業所表記を除いた社名。"""
    s = _nfkc(name)
    s = re.sub(r"[【\[(（](?!(?:株|有|同|資|名)\))[^】\])）]*[】\])）]", " ", s)
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
    """名称が工場・営業所等の拠点を表しているか（「本社工場」表記は除く）。"""
    s = _nfkc(name).replace("本社工場", "")
    s = re.sub(r"^.*?(?:株式会社|有限会社|\((?:株|有)\))", "", s) if _LEGAL_RE.match(s) else _LEGAL_RE.split(s)[-1] if _LEGAL_RE.search(s) else s
    return bool(re.search(r"工場|営業所|支店|支社|事業所|事務所|出張所|整備場|カンパニー|センター|オフィス|ショールーム|倉庫|事業部|"
                          r"(?:九州|福岡|北九州)製作所", s))


# ── Phase 1: サプライヤーパーク企業概要 PDF / NAVITIME ───────────────

def _rec(src: str, name: str, addr: str, tel: str = "", web: str = "", mail: str = "", category: str = "") -> dict:
    return {"src": src, "name": _nfkc(name), "addr": _norm_addr(addr), "tel": _format_phone(tel),
            "web": web, "mail": mail, "category": category}


def _parse_supplier_pdf(text: str) -> dict:
    """企業概要 PDF の「企業名：」「所在地：〒… TEL …」「URL」を読む（1文字ずつ改行されている所は詰める）。"""
    flat = re.sub(r"[ \t　\n]+", "", _nfkc(text))
    name = re.search(r"企業名[:]?(.+?)(?:\(|設立|所在地|URL|事業内容|沿革|創業)", flat)
    addr = re.search(r"所在地[:]?(?:\d{3}-?\d{4})?(.+?)(?:TEL|電話)", flat)
    tel = re.search(r"(?:TEL|電話)[:]?([\d\-()]{10,16})", flat)
    web = re.search(r"URL[:]?\s*(https?://[^\s]+)", _nfkc(text))
    return {
        "name": name.group(1) if name else "",
        "addr": addr.group(1) if addr else "",
        "tel": tel.group(1) if tel else "",
        "web": web.group(1) if web else "",
    }


def fetch_supplier_park() -> list[dict]:
    cached = _cache("supplier_park.json")
    if cached is not None:
        logger.info(f"Phase 1: サプライヤーパーク キャッシュから {len(cached)} 件")
        return cached
    import pymupdf

    r = requests.get(SUPPLIER_PARK, headers={"User-Agent": UA}, timeout=30)
    r.encoding = "utf-8"
    pdfs: dict[str, str] = {}
    for m in re.finditer(r'<a href="\.\./pdf/([^"]+\.pdf)"[^>]*>([^<]+)</a>', r.text):
        pdfs.setdefault(m.group(1), m.group(2).strip())
    out = []
    for f, label in pdfs.items():
        time.sleep(0.5)
        res = requests.get(SUPPLIER_PDF_BASE + f, headers={"User-Agent": UA}, timeout=60)
        if res.status_code != 200:
            logger.warning(f"  PDF 取得失敗 {f}: {res.status_code}")
            continue
        doc = pymupdf.open(stream=res.content, filetype="pdf")
        info = _parse_supplier_pdf("\n".join(p.get_text() for p in doc))
        info.update(PDF_FIX.get(f, {}))
        name = info["name"] or label
        if f in PDF_BRANCH:
            name = f"{name} {PDF_BRANCH[f]}"
        rec = _rec("サプライヤーパーク", name, info["addr"], info.get("tel", ""), info.get("web", ""))
        rec["pdf"] = f
        out.append(rec)
    _save_cache("supplier_park.json", out)
    logger.info(f"Phase 1: サプライヤーパーク PDF {len(out)} 件")
    return out


def fetch_navitime() -> list[dict]:
    """一覧（1ページ15件、2ページ目以降は ?page=N）→ 各スポット詳細ページの住所・電話番号。"""
    cached = _cache("navitime.json")
    if cached is not None:
        logger.info(f"Phase 1: NAVITIME キャッシュから {len(cached)} 件")
        return cached
    s = requests.Session()
    s.headers["User-Agent"] = UA

    def _get_html(url: str) -> str:
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

    spots, seen, page = [], set(), 1
    while True:
        url = f"{NAVITIME_BASE}?tags={NAVITIME_TAG}&from=tag" if page == 1 else f"{NAVITIME_BASE}?page={page}&tags={NAVITIME_TAG}"
        soup = BeautifulSoup(_get_html(url), "lxml")
        new = 0
        for li in soup.select("li.spot-section"):
            a = li.select_one('a[href*="/poi?spot="]')
            if not a:
                continue
            href = urljoin("https://www.navitime.co.jp/", a["href"]).split("&")[0]
            if href in seen:
                continue
            seen.add(href)
            nm = li.select_one(".spot-name-text") or li.select_one(".spot-name") or a
            spots.append({"name": nm.get_text(strip=True), "href": href})
            new += 1
        if not new:
            break
        page += 1
        time.sleep(random.uniform(0.5, 1.0))
    logger.info(f"Phase 1: NAVITIME 一覧 {page - 1} ページ {len(spots)} 件の詳細を取得")
    out = []
    for sp in spots:
        time.sleep(random.uniform(0.5, 1.0))
        soup = BeautifulSoup(_get_html(sp["href"]), "lxml")
        fields = {}
        for dt in soup.select("dl.detail-text-frame dt"):
            dd = dt.find_next_sibling("dd")
            if dd is not None:
                fields[dt.get_text(strip=True)] = re.sub(r"\s+", " ", dd.get_text(" ", strip=True))
        addr = fields.get("住所", "").replace("地図で見る", "").strip()
        out.append(_rec("NAVITIME", sp["name"], addr, fields.get("電話番号", "")))
    _save_cache("navitime.json", out)
    return out


# ── Phase 2: Google マップ ───────────────────────────────

async def fetch_gmaps() -> list[dict]:
    cached = _cache("gmaps.json")
    if cached is not None:
        logger.info(f"Phase 2: キャッシュから {len(cached)} 件")
        return cached
    from playwright.async_api import async_playwright

    links: dict[str, str] = {}
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        ctx = await browser.new_context(locale="ja-JP", user_agent=UA)
        pg = await ctx.new_page()
        for q in GMAPS_QUERIES:
            await pg.goto("https://www.google.com/maps/search/" + q.replace(" ", "+") + "?hl=ja", wait_until="domcontentloaded")
            await pg.wait_for_timeout(5000)
            feed = pg.locator('div[role="feed"]')
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
            for label, href in await pg.eval_on_selector_all(
                "a.hfpxzc", "as => as.map(a => [a.getAttribute('aria-label'), a.href])"
            ):
                links.setdefault(href.split("?")[0], label)
        logger.info(f"Phase 2: Google マップ {len(links)} 件の詳細を取得")
        # 詳細ページは 4 タブで並列に開く
        sem = asyncio.Semaphore(GMAPS_TABS)

        async def _detail(href: str, name: str) -> dict:
            async with sem:
                tab = await ctx.new_page()
                rec = _rec("Googleマップ", name, "")
                try:
                    await tab.goto(href + "?hl=ja", wait_until="domcontentloaded")
                    try:
                        await tab.wait_for_selector("h1", timeout=10000)
                    except Exception:
                        pass
                    await tab.wait_for_timeout(1500)
                    addr = tab.locator('button[data-item-id="address"]').first
                    if await addr.count():
                        rec["addr"] = _norm_addr(await addr.get_attribute("aria-label") or "")
                    tel = tab.locator('button[data-item-id^="phone:tel:"]').first
                    if await tel.count():
                        rec["tel"] = _format_phone(await tel.get_attribute("aria-label") or "")
                    web = tab.locator('a[data-item-id="authority"]').first
                    if await web.count():
                        rec["web"] = await web.get_attribute("href") or ""
                    cat = tab.locator("button.DkEaL").first
                    if await cat.count():
                        rec["category"] = (await cat.inner_text()).strip()
                finally:
                    await tab.close()
                return rec

        out = list(await asyncio.gather(*[_detail(h, n) for h, n in links.items()]))
        await browser.close()
    _save_cache("gmaps.json", out)
    return out


# ── Phase 3: 名寄せ ──────────────────────────────────────

def group_companies(records: list[dict]) -> dict[str, list[dict]]:
    groups: dict[str, list[dict]] = {}
    for r in records:
        r["addr"] = re.sub(r"^住所[:：]\s*", "", r["addr"])
        r["addr"] = _norm_addr(r["addr"]) if not r["addr"].startswith("福岡県") else r["addr"]
        if not r["addr"].startswith("福岡県"):
            continue  # 県外の拠点・県外本社の掲載
        c = core_name(r["name"])
        if c:
            groups.setdefault(c, []).append(r)
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
    todo = [c for c in groups if c not in cache]
    logger.info(f"Phase 4: 法人番号公表サイトで本店所在地確認 {len(todo)} 件（キャッシュ {len(cache)} 件）")
    local = threading.local()  # 検索トークンはセッション単位なのでスレッドごとにセッションを持つ
    lock = threading.Lock()
    done = {"n": 0}

    def _one(core: str) -> None:
        if not hasattr(local, "nta"):
            local.nta = NTA()
        try:
            in_pref = [x for x in local.nta.search(core, NTA_PREF_FUKUOKA) if _same_core(x["name"], core)]
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
    """「福岡県北九州市小倉南区…」→「北九州市小倉南区」、「福岡県京都郡苅田町…」→「京都郡苅田町」。"""
    a = _nfkc(addr).removeprefix("福岡県")
    m = re.match(r"(北九州市.+?区|福岡市.+?区|.+?郡.+?[町村]|.+?市)", a)
    return m.group(1) if m else ""


def _hq_hits(core: str, recs: list[dict], nta: dict, town: bool = False) -> list[dict]:
    """県内の同名法人のうち、掲載拠点と同じ市区町村（town=True なら町名まで）に本店があるもの。"""
    def key(addr: str) -> str:
        return _city(addr) + (_town(addr)[:2] if town else "")
    keys = {key(r["addr"]) for r in recs if _city(r["addr"])}
    return [h for h in nta.get(core, {}).get("in_pref", []) if key(_nfkc(h["addr"])) in keys]


def decide(core: str, recs: list[dict], nta: dict) -> tuple[bool, str]:
    """(福岡県本社か, 判定理由)"""
    res = nta.get(core, {"in_pref": [], "other": []})
    all_branch = all(is_branch(r["name"]) for r in recs)
    if res["in_pref"]:
        if all_branch:
            # 工場・営業所名義の掲載しかないときは、本店が同じ町にある（本社工場等）場合だけ採用
            if _hq_hits(core, recs, nta, town=True):
                return True, "法人番号: 県内本店（工場と同じ町）"
            return False, "工場・営業所の掲載のみ（本店は別の場所）"
        if _hq_hits(core, recs, nta):
            return True, "法人番号: 県内本店（所在地一致）"
        if len(res["in_pref"]) == 1:
            return True, "法人番号: 県内本店"
        return False, "県内に同名法人が複数あり所在地が一致しない"
    if res["other"]:
        return False, "法人番号: 県外本店"
    if all_branch:
        return False, "工場・営業所表記・法人登録なし"
    return True, "法人登録なし（個人事業等）"


# ── Phase 5: 公式サイト検索 ────────────────────────────────

SKIP_DOMAINS = [
    "google.", "yahoo.", "bing.com", "facebook.com", "instagram.com", "twitter.com", "x.com",
    "youtube.com", "wikipedia", "jpnumber", "telnavi", "houjin", "baseconnect", "mapion",
    "navitime", "itp.ne.jp", "ekiten", "indeed", "townwork", "mynavi", "hellowork", "baitoru",
    "en-gage", "careerindex", "jobmedley", "job-medley",
    "kyujin", "recruit", "rikunabi", "doda", "stanby", "tdb.co.jp",
    "salesnow", "musubu", "alarmbox", "info-houjin", "nta.go.jp", "gbiz", "catr.jp", "graffer",
    "suumo", "homes.co.jp", "biz-maps", "jbplt", "joho-fukuoka.or.jp",
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
    a = _nfkc(addr).removeprefix("福岡県")
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
    """ページ内の福岡県（092・093・094x・0930/0940/0942〜0949）・フリーダイヤルの電話番号を (位置, 整形済み番号) で返す。"""
    out = []
    for m in re.finditer(r"(?<!\d)(0\d{1,4})[-‐ー−()（）\s]{0,2}(\d{1,4})[-‐ー−()（）\s]{0,2}(\d{3,4})(?!\d)", text):
        digits = "".join(m.groups())
        if re.search(r"FAX|Fax|fax|ファクス|ファックス", text[max(0, m.start() - 10):m.start()]):
            continue
        if len(digits) == 10 and digits.startswith(("092", "093", "094", "0930", "0979", "0120", "0800")):
            out.append((m.start(), _format_phone(digits)))
    return out


def search_official(core: str, addr: str) -> tuple[str, str]:
    """(公式サイトURL, まとめサイト等から得た電話番号) を返す。"""
    city, town = _city(addr), _town(addr)
    query = f"{core} {city or '福岡県'} {town}"
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
    sem = asyncio.Semaphore(8)

    async def _one(c: dict) -> None:
        async with sem:
            extra = await asyncio.to_thread(enrich_site, c["公式サイトURL"], c["_core"])
            for k in ("メールアドレス", "インスタURL", "問い合わせフォームURL"):
                if extra.get(k) and not c.get(k):
                    c[k] = extra[k]
            if not c["電話番号"] and extra.get("_tel"):
                c["電話番号"] = extra["_tel"]

    await asyncio.gather(*[_one(c) for c in targets])


async def enrich_all(companies: list[dict]) -> None:
    with_url = [c for c in companies if c["公式サイトURL"]]
    logger.info(f"Phase 6: 公式サイトから連絡先取得 {len(with_url)} 件")
    await _enrich_sites(with_url)

    targets = [c for c in companies if not c["公式サイトURL"] or not c["電話番号"]]
    logger.info(f"Phase 5: 公式サイトまたは電話番号が未取得の {len(targets)} 件を検索")
    sem = asyncio.Semaphore(3)

    async def _search(c: dict) -> None:
        async with sem:
            url, tel = await asyncio.to_thread(search_official, c["_core"], c["所在地"])
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

_SRC_ORDER = {"サプライヤーパーク": 0, "NAVITIME": 1, "Googleマップ": 2}


def _pick(recs: list[dict], key: str) -> str:
    """本社・非拠点のレコード、企業概要 PDF を優先して値を選ぶ。"""
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


async def main() -> None:
    logger.info("=" * 65)
    logger.info("福岡県 自動車部品製造会社リスト作成開始")
    logger.info("=" * 65)
    start = datetime.now()

    records = fetch_supplier_park() + fetch_navitime() + await fetch_gmaps()
    groups = group_companies(records)
    logger.info(f"Phase 3: 名寄せ後 {len(groups)} 社")

    nta = check_hq(groups)
    companies = []
    excluded = []
    for core, recs in groups.items():
        ok, reason = decide(core, recs, nta)
        if not ok:
            excluded.append((core, reason, recs[0]["name"]))
            continue
        hit = _hq_hits(core, recs, nta) or nta.get(core, {}).get("in_pref", [])
        non_branch = [r for r in recs if not is_branch(r["name"])]
        non_branch.sort(key=lambda r: _SRC_ORDER.get(r["src"], 9))
        addr = non_branch[0]["addr"] if non_branch else _norm_addr(hit[0]["addr"])
        companies.append({
            "_core": core,
            "_reason": reason,
            "_recs": recs,
            "名称": _display_name(core, recs, hit),
            "メールアドレス": _pick(recs, "mail"),
            "公式サイトURL": _clean_url(_pick(recs, "web")),
            "所在地": addr,
            # 本社の掲載がない（県内本社の工場のみ掲載）場合、所在地は登記上の本店なので
            # 工場の電話番号は公式サイト等で本社番号が見つからなかったときの代替とする
            "電話番号": _pick(non_branch, "tel") if non_branch else "",
            "_branch_tel": "" if non_branch else _pick(recs, "tel"),
            "インスタURL": "",
            "問い合わせフォームURL": "",
        })
    for core, reason, name in sorted(excluded):
        logger.info(f"  除外: {name}（{reason}）")
    logger.info(f"Phase 4: 福岡県本社 {len(companies)} 社 / 除外 {len(excluded)} 社")

    mfg = [c for c in companies if is_manufacturer(c["_core"], c["名称"], c["_recs"])]
    for c in companies:
        if c not in mfg:
            logger.info(f"  自動車部品製造以外のため除外: {c['名称']}")
    logger.info(f"Phase 7: 自動車部品製造 {len(mfg)} 社 / それ以外 {len(companies) - len(mfg)} 社を除外")
    companies = mfg
    for c in companies:
        del c["_recs"]
        if any(w in c["公式サイトURL"] for w in WRONG_URLS):
            c["公式サイトURL"] = ""

    await enrich_all(companies)
    for c in companies:
        clean_site_fields(c)

    df = pd.DataFrame(companies)
    # Google マップでは同じ会社が別名義（「ものづくり 金属加工 北九州(株)川村製作所」等）で載ることがあるので
    # 電話番号が同じものは1社にまとめる（法人番号で確認できた名義・公式サイトありを残す）
    df["_rank"] = df["_reason"].str.startswith("法人番号").astype(int) * 2 + (df["公式サイトURL"] != "").astype(int)
    df = df.sort_values("_rank", ascending=False)
    dup = (df["電話番号"] != "") & df.duplicated("電話番号")
    for name in df.loc[dup, "名称"]:
        logger.info(f"  電話番号重複のため統合: {name}")
    df = df[~dup]
    df = df.sort_values(["所在地", "名称"]).reset_index(drop=True)
    _save_cache("companies.json", df.to_dict("records"))
    df = df[OUTPUT_COLS]
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="自動車部品製造会社一覧")

    elapsed = int((datetime.now() - start).total_seconds())
    logger.info("=" * 65)
    logger.info(f"完了: {OUTPUT_FILE} に {len(df)} 件を出力")
    for col in OUTPUT_COLS[1:]:
        logger.info(f"  {col}: {(df[col] != '').sum()} 件")
    logger.info(f"所要時間: {elapsed // 60}分{elapsed % 60}秒")
    logger.info("=" * 65)


if __name__ == "__main__":
    asyncio.run(main())
