"""
熊本県 自動車部品製造会社リスト作成

データソース:
  Phase 1 - ふくおか自動車部品サプライヤーパーク（joho-fukuoka.or.jp）の熊本県企業一覧から
              各企業の企業概要 PDF を取得し、企業名・所在地・TEL・URL を読む
              （県外本社の会社は PDF 内の熊本工場の欄を使う。画像だけ・フォント崩れの PDF は
              PDF_FIX に目視で読んだ値を持つ）
            熊本県「自動車関連企業リスト」PDF（pref.kumamoto.jp/uploaded/attachment/123079.pdf）
              1ページ22社の表。主要製造品目が自動車部品のものを PREF_LIST に持つ
            企業クリップ「熊本県の製造業代表企業」全60ページ（1ページ15社）
              事業内容に自動車・二輪・ワイヤーハーネス等の語があり、かつ製造・加工を
              行っている会社を採用（販売・整備・設備メーカー等は NON_MFG で除外）
  Phase 3 - 名寄せ: 法人格・工場/事業所表記を除いた社名（コア名）で統合
  Phase 6 - 公式サイト（トップ＋会社概要・問い合わせページ）からメールアドレス・
            インスタURL・問い合わせフォームURL・電話番号（未取得時）を取得
  Phase 5 - 公式サイトまたは電話番号が未取得の会社は DuckDuckGo で検索し、社名と
            市町村・町名がページ内にあることを確認して採用
  Phase 7 - それでも電話番号がない会社は Google マップを「社名 市町村」で検索し、
            スポット名と市町村が一致するときだけ電話番号・ウェブサイトを使う

熊本県内の製造拠点をすべて対象とし、県外本社の会社は「熊本工場」等の拠点名を付けて載せる。
電話番号は phonenumbers で日本の市外局番の境界に従ってハイフンを入れる。

出力列: 名称, メールアドレス, 公式サイトURL, 所在地, 電話番号, インスタURL, 問い合わせフォームURL
出力ファイル: 熊本県_自動車部品製造会社リスト.xlsx
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

LOG_FILE = "scraper_kumamoto_jidousha_buhin.log"
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
OUTPUT_FILE = "熊本県_自動車部品製造会社リスト.xlsx"
CACHE_DIR = os.path.join("data", "kumamoto_jidousha_buhin")
UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36"
)
SUPPLIER_PARK = "https://www.joho-fukuoka.or.jp/k-car-supplierpark/search-area/kumamoto.html"
SUPPLIER_PDF_BASE = "https://www.joho-fukuoka.or.jp/k-car-supplierpark/pdf/"
CLIP_BASE = "https://clip-corporation.com/column/manufacturing-kumamoto/"
CLIP_PAGES = 60

# ── サプライヤーパーク PDF の補正 ───────────────────────────
PDF_FIX: dict[str, dict] = {
    # 画像だけの PDF（九州拠点の欄）
    "aoyamaseisakusyo.pdf": {"name": "株式会社青山製作所 熊本工場", "addr": "熊本県阿蘇郡高森町大字高森字豆塚2141-5", "tel": "0967-62-2121"},
    # フォントが崩れて文字化けする PDF
    "kfk-ogawa.pdf": {"name": "株式会社ケイエフケイ小川", "addr": "熊本県宇城市小川町南海東1961", "tel": "0964-43-4848"},
    # 県外本社の会社は熊本の工場の欄を使う
    "kobayashiseisakusyo.pdf": {"name": "株式会社小林製作所 熊本工場", "addr": "熊本県玉名郡南関町大字豊永5000番地", "tel": "0968-53-2817"},
    "mogiseisakusyo.pdf": {"name": "株式会社茂木製作所 九州工場", "addr": "熊本県玉名郡南関町肥猪4002-1", "tel": "0968-57-8485"},
    "tfo.pdf": {"name": "ティエフオー株式会社 熊本工場"},
    "kyushu-trix.pdf": {"addr": "熊本県荒尾市水野1092-9"},
    "kumabometaru.pdf": {"addr": "熊本県熊本市東区長嶺西1丁目4番15号"},
}

# 熊本県「自動車関連企業リスト」のうち主要製造品目が自動車部品（車載品）のもの
# （半導体装置・生産設備・ソフトウェア・高圧ガス等の会社は除いた）
# (企業名（工場名）, 所在地, 電話番号, URL)
PREF_LIST = [
    ("株式会社オジックテクノロジーズ", "熊本県熊本市西区上熊本2-9-9", "096-352-4450", "http://www.ogic.ne.jp/"),
    ("神田工業株式会社 熊本事業所", "熊本県八代市鏡町有佐1302", "0965-52-1071", "http://www.kanda-kogyo.co.jp/"),
    ("平井精密工業株式会社 熊本事業所", "熊本県荒尾市宮内字下山下902", "0968-62-0311", "https://www.hirai.co.jp/"),
    ("安高金属工業株式会社", "熊本県荒尾市高浜310-8", "0968-68-6111", "http://www.yasutaka.co.jp/"),
    ("九州三永金属工業株式会社", "熊本県玉名市大倉478", "0968-73-2251", "http://kyusyusanei.co.jp/public_html/index.html"),
    ("タツミ精工株式会社 熊本工場", "熊本県玉名市三ツ川1528", "0968-74-9126", "http://www.tatsumi-s.co.jp"),
    ("熊本部品株式会社", "熊本県天草市楠浦町1046-18", "0969-24-0800", "http://www.kumamotoparts.com/"),
    ("株式会社美加川製作所", "熊本県山鹿市久原4230-16", "0968-44-7444", "http://www.mikawa-ht.jp/"),
    ("ジヤスペック株式会社", "熊本県菊池市森北1812-25", "0968-24-5511", "http://jaspeq.co.jp/"),
    ("株式会社山清工業九州", "熊本県菊池市七城町蘇崎字植木山1241-7", "0968-26-4171", "http://www.yamaseikogyo.co.jp/index.html"),
    ("熊本工機株式会社", "熊本県玉名郡南関町大字四ツ原1280", "0968-53-9101", ""),
    ("株式会社白水ＤＨＣ", "熊本県阿蘇郡南阿蘇村吉田2111", "0967-62-8931", ""),
]

# ── 自動車部品製造の判定（企業クリップ） ───────────────────────
_AUTO_RE = re.compile(
    r"自動車|車載|カーエアコン|ハーネス|組電線|二輪|四輪|オートバイ|バイク|クラッチ|ショックアブソーバ|"
    r"マフラー|ミッション|フェンダー"
)
_MFG_RE = re.compile(r"製造|加工|成形|成型|生産|製作|組立|組付|メーカー|鋳|鍛|プレス|熱処理|表面処理|部品")
# 事業内容に自動車関連の語はあるが、自動車部品の製造・加工が本業でない会社（コア名）
NON_MFG = {
    "テラシステム",          # 生産設備・検査装置
    "トヨタエルアンドエフ熊本",  # 産業車両の販売・整備
    "福田屋", "夢・さぽーと", "木之内農園", "トミー食品", "中九州観光", "熊本コアミックス",
    "球磨郡公立多良木病院企業団", "熊本小野田レミコン", "九州産交バス",
    "三和工業",              # 建材販売・運送
    "Ｈ・Ｉシステック",      # 生産設備用の基板・制御盤
    "イズミ車体製作所",      # 福祉車両・特種車両（完成車）の製造
    "内村酸素", "熊本酸素",  # 高圧ガス・機械器具販売
    "ベストロジ熊本",        # 梱包
    "ＴＯＰ",                # 半導体検査装置
    "九州コーテッドサンド",  # 鋳物砂
    "イーザイ化成",          # 樹脂原料の加工販売
    "ティービーエム",        # 架装・鈑金・車検
    "ホンダパーツ九州", "フタバ",  # 部品卸
    "サントライ",            # 電子部品の検査
    "パリッシュ",            # 半導体装置部品加工・工場内メンテナンス
    "星山商店", "カネムラエコワークス",  # 金属リサイクル
    "吉本自動車", "山田電機", "本田商事",  # 修理・電装品販売
    "熊本オートバックス", "熊本三菱自動車販売", "スズキ自販熊本",  # 販売
    "ＪＮＣ開発",            # 樹脂資材（建材・農業資材等）
}

# 検索で誤って別の会社・無関係のページが公式サイトURLとして入るもの
WRONG_URLS: tuple[str, ...] = (
    "matomame.jp", "livecamera.", "kensetumap.com", "tsukulink.net", "tellogs.com", "goo-net.com",
    "carnext.jp", "postmap.org", "voix.jp", "glarity.app", "hotfrog.jp", "telinquire.com",
    "joho-fukuoka.or.jp", "jetro.go.jp", "ipros.jp", "metoree.com", "aperza.com",
    "bing.com", "kyoka-db.com", "big-advance.site", "culebee.com", "clip-corporation.com", "pref.kumamoto.jp",
    "irbank.net", "cat-studio.jp", "agritrio.co.jp", "kumamotobank.co.jp", "monomiza.jp",
    "eigyo-mfg.com", "nagomishoko.jp",
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


def _norm_addr(s: str) -> str:
    """所在地を「熊本県…」形式に整える（郵便番号・拠点の見出しを除去）。"""
    a = _nfkc(s)
    a = re.sub(r"〒?\s*\d{3}-?\d{4}\s*", "", a).strip()
    i = a.find("熊本県")
    if i > 0 and not re.search(r"[都道府]|.県", a[:i]):
        a = a[i:]  # 「本社、荒尾工場熊本県…」のような見出しを落とす
    a = a.replace("下益城郡城南町", "熊本市南区城南町")  # 2010年に熊本市へ編入
    return re.sub(r"\s+", " ", a)


_LEGAL_RE = re.compile(r"株式会社|有限会社|合同会社|合資会社|合名会社|\((?:株|有|同|資|名)\)|㈱|㈲")
_BRANCH_RE = re.compile(r"\s*(?:熊本|九州|八代|荒尾)?(?:工場|事業所|営業所)$|\s*[(（][^)）]*事業[)）]$")
# 資料によって表記が違う同じ会社（コア名 → 統一名）
_ALIASES = {"白水ディエイチシー": "白水DHC"}


def core_name(name: str) -> str:
    """法人格・工場/事業所表記・中黒を除いた社名（全角英数も半角に揃える）。"""
    s = re.sub(r"\s*\((?![株有同資名]\))[^)]*\)\s*$", "", _nfkc(name))  # 「( Kyushu F.tech inc.)」等の英語表記
    s = _BRANCH_RE.sub("", _LEGAL_RE.sub(" ", s).strip())
    s = s.replace(" ", "").replace("・", "")
    return _ALIASES.get(s, s)


def _rec(src: str, name: str, addr: str, tel: str = "", web: str = "", biz: str = "") -> dict:
    return {"src": src, "name": _nfkc(name), "addr": _norm_addr(addr), "tel": _format_phone(tel), "web": web, "biz": biz}


# ── Phase 1: データソース ─────────────────────────────────

def _parse_supplier_pdf(text: str) -> dict:
    """企業概要 PDF の「企業名：」「所在地：〒… TEL …」「URL」を読む（1文字ずつ改行されている所は詰める）。"""
    flat = re.sub(r"[ \t　\n]+", "", _nfkc(text))
    name = re.search(r"企業名[:]?(.+?)(?:\(|設立|所在地|URL|事業内容|沿革|創業)", flat)
    addr = re.search(r"所在地[:]?(.+?)(?:TEL|電話)", flat)
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
        out.append(_rec("サプライヤーパーク", info["name"] or label, info["addr"], info["tel"], info["web"]))
    _save_cache("supplier_park.json", out)
    logger.info(f"Phase 1: サプライヤーパーク PDF {len(out)} 件")
    return out


def fetch_pref_list() -> list[dict]:
    out = [_rec("熊本県リスト", n, a, t, w) for n, a, t, w in PREF_LIST]
    logger.info(f"Phase 1: 熊本県 自動車関連企業リスト {len(out)} 件")
    return out


def fetch_clip() -> list[dict]:
    """各社「社名 / 所在地 / … / 事業内容 / 本文 / 代表者 / 氏名 / 社名 / 詳しい事業内容」の行の並びを読む。"""
    cached = _cache("clip.json")
    if cached is None:
        cached = []
        for n in range(1, CLIP_PAGES + 1):
            url = CLIP_BASE + (f"{n}/" if n > 1 else "")
            soup = BeautifulSoup(requests.get(url, headers={"User-Agent": UA}, timeout=30).content, "lxml")
            lines = [x.strip() for x in (soup.find("main") or soup.body).get_text("\n").split("\n") if x.strip()]
            idx = [i for i, x in enumerate(lines) if x == "所在地" and i + 1 < len(lines) and lines[i + 1].startswith("熊本県")]
            for k, i in enumerate(idx):
                end = idx[k + 1] - 1 if k + 1 < len(idx) else len(lines)
                blk = lines[i - 1:end]
                biz = blk[blk.index("事業内容") + 1] if "事業内容" in blk else ""
                rep = blk.index("代表者") if "代表者" in blk else len(blk)
                # 最後の会社の後ろはページ送り・関連記事なので「前へ」で切る
                tail = blk[rep + 3:]
                tail = tail[:tail.index("前へ")] if "前へ" in tail else tail
                cached.append({"name": blk[0], "addr": blk[blk.index("所在地") + 1], "biz": biz + " " + " ".join(tail)})
            time.sleep(1)
        _save_cache("clip.json", cached)
    logger.info(f"Phase 1: 企業クリップ 熊本県の製造業 {len(cached)} 社")
    out = []
    for c in cached:
        if core_name(c["name"]) in {core_name(x) for x in NON_MFG}:
            continue
        if _AUTO_RE.search(c["biz"]) and _MFG_RE.search(c["biz"]):
            out.append(_rec("企業クリップ", c["name"], c["addr"], biz=c["biz"]))
    logger.info(f"  うち自動車部品の製造・加工 {len(out)} 社")
    return out


# ── Phase 5: 公式サイト検索 ────────────────────────────────

SKIP_DOMAINS = [
    "google.", "yahoo.", "bing.com", "facebook.com", "instagram.com", "twitter.com", "x.com",
    "youtube.com", "wikipedia", "jpnumber", "telnavi", "houjin", "baseconnect", "mapion",
    "navitime", "itp.ne.jp", "ekiten", "indeed", "townwork", "mynavi", "hellowork", "baitoru",
    "en-gage", "careerindex", "jobmedley", "job-medley",
    "kyujin", "recruit", "rikunabi", "doda", "stanby", "tdb.co.jp",
    "salesnow", "musubu", "alarmbox", "info-houjin", "nta.go.jp", "gbiz", "catr.jp", "graffer",
    "suumo", "homes.co.jp", "biz-maps", "jbplt", "joho-fukuoka.or.jp", "clip-corporation.com",
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


def _city(addr: str) -> str:
    """「熊本県熊本市東区…」→「熊本市東区」、「熊本県玉名郡南関町…」→「玉名郡南関町」。"""
    a = _nfkc(addr).removeprefix("熊本県")
    m = re.match(r"(熊本市.+?区|.+?郡.+?[町村]|.+?市)", a)
    return m.group(1) if m else ""


def _town(addr: str) -> str:
    """所在地の市町村より後ろの町名・大字名（番地の手前まで）。"""
    a = _nfkc(addr).removeprefix("熊本県")
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
    """ページ内の熊本県（096・0964〜0969）・フリーダイヤルの電話番号を (位置, 整形済み番号) で返す。"""
    out = []
    for m in re.finditer(r"(?<!\d)(0\d{1,4})[-‐ー−()（）\s]{0,2}(\d{1,4})[-‐ー−()（）\s]{0,2}(\d{3,4})(?!\d)", text):
        digits = "".join(m.groups())
        if re.search(r"FAX|Fax|fax|ファクス|ファックス", text[max(0, m.start() - 10):m.start()]):
            continue
        if len(digits) == 10 and digits.startswith(("096", "0120", "0800")):
            out.append((m.start(), _format_phone(digits)))
    return out


def search_official(core: str, addr: str) -> tuple[str, str]:
    """(公式サイトURL, まとめサイト等から得た電話番号) を返す。"""
    city, town = _city(addr), _town(addr)
    query = f"{core} {city or '熊本県'} {town}"
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


# ── Phase 7: Google マップで電話番号・公式サイトを補完 ─────────────

async def gmaps_lookup(companies: list[dict]) -> None:
    """電話番号が未取得の会社を「社名 市町村」で検索し、名称と市町村が一致するスポットの値を使う。"""
    cache = _cache("gmaps_lookup.json") or {}
    targets = [c for c in companies if not c["電話番号"]]
    todo = [c for c in targets if c["_core"] not in cache]
    logger.info(f"Phase 7: Google マップで補完 {len(targets)} 件（キャッシュ {len(targets) - len(todo)} 件）")
    if todo:
        from playwright.async_api import async_playwright

        async with async_playwright() as p:
            browser = await p.chromium.launch(headless=True)
            ctx = await browser.new_context(locale="ja-JP", user_agent=UA)
            pg = await ctx.new_page()
            for c in todo:
                rec = {"name": "", "addr": "", "tel": "", "web": ""}
                q = f"{c['名称']} {_city(c['所在地'])}"
                await pg.goto("https://www.google.com/maps/search/" + q.replace(" ", "+") + "?hl=ja", wait_until="domcontentloaded")
                await pg.wait_for_timeout(4000)
                # 検索結果が一覧になったときは先頭のスポットを開く
                first = pg.locator("a.hfpxzc").first
                if await first.count():
                    await first.click()
                    await pg.wait_for_timeout(3000)
                h1 = pg.locator("h1").last
                if await h1.count():
                    rec["name"] = (await h1.inner_text()).strip()
                for key, sel, attr in (("addr", 'button[data-item-id="address"]', "aria-label"),
                                       ("tel", 'button[data-item-id^="phone:tel:"]', "aria-label"),
                                       ("web", 'a[data-item-id="authority"]', "href")):
                    el = pg.locator(sel).first
                    if await el.count():
                        rec[key] = await el.get_attribute(attr) or ""
                cache[c["_core"]] = rec
                await pg.wait_for_timeout(random.randint(1000, 2000))
            await browser.close()
        _save_cache("gmaps_lookup.json", cache)
    n = 0
    for c in targets:
        rec = cache.get(c["_core"], {})
        spot = core_name(rec.get("name", ""))
        # 社名変更で「日立」が外れた「Astemo宇城」のように、どちらかがもう一方を含めば同じ会社とみなす
        if not spot or (spot not in c["_core"] and c["_core"] not in spot) or _city(c["所在地"]) not in _nfkc(rec.get("addr", "")):
            continue  # 別のスポット
        if rec.get("tel"):
            c["電話番号"] = _format_phone(rec["tel"])
            n += 1
        if rec.get("web") and not c["公式サイトURL"]:
            c["公式サイトURL"] = rec["web"]
    logger.info(f"  電話番号を補完: {n} 件")


# ── メイン ────────────────────────────────────────────────

_SRC_ORDER = {"サプライヤーパーク": 0, "熊本県リスト": 1, "企業クリップ": 2}


def _pick(recs: list[dict], key: str) -> str:
    for r in sorted(recs, key=lambda r: _SRC_ORDER.get(r["src"], 9)):
        if r.get(key):
            return r[key]
    return ""


def _display_name(recs: list[dict]) -> str:
    """工場名つきの名称（県外本社の拠点）があればそれを、なければ企業概要 PDF の名称を使う。"""
    branch = [r["name"] for r in recs if _BRANCH_RE.search(r["name"])]
    nm = branch[0] if branch else _pick(recs, "name")
    nm = _nfkc(nm).replace("(株)", "株式会社").replace("(有)", "有限会社")
    return re.sub(r"\s*[(（][^)）]*事業[)）]$", "", nm)


async def main() -> None:
    logger.info("=" * 65)
    logger.info("熊本県 自動車部品製造会社リスト作成開始")
    logger.info("=" * 65)
    start = datetime.now()

    records = fetch_supplier_park() + fetch_pref_list() + fetch_clip()
    groups: dict[str, list[dict]] = {}
    for r in records:
        if not r["addr"].startswith("熊本県"):
            logger.info(f"  県外の所在地のため除外: {r['name']}（{r['addr']}）")
            continue
        groups.setdefault(core_name(r["name"]), []).append(r)
    logger.info(f"Phase 3: 名寄せ後 {len(groups)} 社")

    companies = []
    for core, recs in groups.items():
        companies.append({
            "_core": core,
            "名称": _display_name(recs),
            "メールアドレス": "",
            "公式サイトURL": _pick(recs, "web"),
            "所在地": _pick(recs, "addr"),
            "電話番号": _pick(recs, "tel"),
            "インスタURL": "",
            "問い合わせフォームURL": "",
        })

    await enrich_all(companies)
    for c in companies:
        clean_site_fields(c)
    await gmaps_lookup(companies)
    for c in companies:
        clean_site_fields(c)

    df = pd.DataFrame(companies)
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
