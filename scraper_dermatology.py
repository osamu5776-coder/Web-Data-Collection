"""
埼玉県の皮膚科クリニック情報収集スクレイパー

- 検索: DuckDuckGo
- スクレイピング: Playwright（JS対応）+ BeautifulSoup（パース）
- robots.txt: ドメインごとにチェック・キャッシュ
- リクエスト間隔: 1〜3秒のランダム待機
- 出力: Excelファイル
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
from playwright.async_api import async_playwright

# ── ロギング設定 ──────────────────────────────────────────
# Windows の cp932 コンソールで絵文字などが原因のエンコードエラーを防ぐ
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

LOG_FILE = "scraper.log"
_stream_handler = logging.StreamHandler(
    io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    if hasattr(sys.stdout, "buffer") else sys.stdout
)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        _stream_handler,
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
    ],
)
logger = logging.getLogger(__name__)

# ── 定数 ─────────────────────────────────────────────────
SEARCH_KEYWORD = "埼玉県 皮膚科 クリニック 公式サイト"
MAX_SEARCH_RESULTS = 30
DELAY_MIN = 1.0
DELAY_MAX = 3.0
OUTPUT_FILE = f"saitama_dermatology_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/124.0.0.0 Safari/537.36"
)

# まとめサイト・SNS・無関係サイト等 除外ドメイン
SKIP_DOMAINS = [
    "instagram.com", "facebook.com", "twitter.com", "x.com",
    "youtube.com", "tiktok.com", "yelp.com",
    "tabelog.com", "hotpepper.jp", "gurunavi.com",
    "jalan.net", "wikipedia.org", "google.com", "bing.com",
    "caloo.jp", "dr-map.jp", "qlife.jp", "iryou.info",
    "byoinnavi.jp", "medley.life", "minnano-clinic.jp",
    "sitescorechecker.com", "fliphtml5.com", "jimcontent.com",
]

# 除外するURLパターン（拡張子など）
SKIP_URL_PATTERNS = [r"\.pdf$", r"\.docx?$"]

# ── robots.txt キャッシュ ──────────────────────────────────
_robots_cache: dict[str, RobotFileParser] = {}


def get_robots(base_url: str) -> RobotFileParser:
    """ドメインの robots.txt を取得・キャッシュして返す。"""
    parsed = urlparse(base_url)
    origin = f"{parsed.scheme}://{parsed.netloc}"
    if origin in _robots_cache:
        return _robots_cache[origin]

    rp = RobotFileParser()
    robots_url = f"{origin}/robots.txt"
    try:
        rp.set_url(robots_url)
        rp.read()
        logger.info(f"robots.txt 読み込み完了: {robots_url}")
    except Exception as e:
        logger.warning(f"robots.txt 取得失敗（全許可として扱う）: {robots_url} / {e}")

    _robots_cache[origin] = rp
    return rp


def is_allowed(url: str) -> bool:
    """robots.txt がアクセスを許可しているか確認する。"""
    rp = get_robots(url)
    allowed = rp.can_fetch(USER_AGENT, url)
    if not allowed:
        logger.warning(f"robots.txt により禁止: {url}")
    return allowed


def random_wait() -> None:
    """1〜3秒のランダム待機でサーバー負荷を軽減する。"""
    delay = random.uniform(DELAY_MIN, DELAY_MAX)
    logger.debug(f"待機: {delay:.1f}秒")
    time.sleep(delay)


# ── クリーニング ──────────────────────────────────────────

CLINIC_KEYWORDS = ["皮膚科", "皮フ科", "クリニック", "医院", "病院", "診療所"]
_STRONG_KW = ["クリニック", "医院", "病院", "診療所"]
_WEAK_KW   = ["皮膚科", "皮フ科"]

def _has_clinic_kw(s: str) -> bool:
    return any(kw in s for kw in CLINIC_KEYWORDS)

def _clinic_score(s: str) -> int:
    """セグメントがクリニック名らしいほど高スコアを返す。"""
    if len(s) <= 3:          # 「皮膚科」単独などは除外
        return 0
    score = sum(2 for kw in _STRONG_KW if kw in s)
    score += sum(1 for kw in _WEAK_KW  if kw in s)
    # 地名の羅列っぽい（スペース区切りの複数診療科リスト）は減点
    if s.count('・') >= 2 or s.count(' ') >= 2:
        score -= 1
    return score


def clean_name(raw: str) -> str:
    """ページタイトルからクリニック名だけを抽出する。"""
    if not raw or not raw.strip():
        return raw
    name = raw.strip()

    # 1. 【クリニック名】: 括弧内に強いキーワードがあり地名説明でなければ最優先
    for m in re.finditer(r'【([^】]{2,20})】', name):
        content = m.group(1)
        if (any(kw in content for kw in _STRONG_KW)
                and not re.search(r'[都道府県市区町村]', content)):
            return content.strip()

    # 2. 定型プレフィックス除去
    name = re.sub(r'^(?:トップページ|HOME|TOP)\s*[-－|｜]\s*', '', name).strip()

    # 3. 区切り文字で分割 → スコア最高のセグメントを選択
    segments = [name]
    for sep in ['｜', ' | ', '|', ' - ']:
        if sep in name:
            segments = [s.strip() for s in name.split(sep) if s.strip()]
            break
    if len(segments) > 1:
        scored = [(s, _clinic_score(s)) for s in segments]
        best = max(sc for _, sc in scored)
        if best > 0:
            name = min((s for s, sc in scored if sc == best), key=len)
        else:
            name = segments[0]

    # 4. 装飾ブラケットをスペースで置換（除去すると前後が結合するため）
    name = re.sub(r'【[^】]{0,15}】', ' ', name)
    name = re.sub(r'[ \t]+', ' ', name).strip()  # 半角スペースのみ正規化（全角スペースは保持）

    # 5. 地名プレフィックス除去 (複数パターンを順に適用)
    # a. "XX市 XX皮膚科" → 空白区切りの地名プレフィックス
    m = re.match(r'^.{1,12}(?:市|区|町|村|駅)[^\s　]*[ \t]+(.+)', name)
    if m and _has_clinic_kw(m.group(1)):
        name = m.group(1)
    # b. "埼玉県XX市のXXクリニック" → 「市の」＋強いキーワード
    m = re.match(r'^.{1,15}(?:市|区|町|村)の(.+?(?:クリニック|医院|病院|診療所))', name)
    if m:
        name = m.group(1)
    # c. "浦和の皮膚科いとう医院" → の＋弱キーワード＋強キーワード
    m = re.match(r'^.{1,8}の(?:皮膚科|皮フ科)(.{2,15}(?:医院|クリニック|病院|センター))', name)
    if m:
        name = m.group(1)

    # 6. 括弧内の補足説明除去: （毎週水曜日・祝日 予約制）等
    name = re.sub(r'[（(][^）)]{3,30}[）)]', '', name).strip()

    # 7. 空白区切りのトークンを順に評価して余分な説明を除去
    #    - クリニックキーワードを持つトークンの後に、キーワードも院名もない → 切り捨て
    #    - 「浦和院」など院名サフィックスは保持
    #    - 「国立病院機構 埼玉病院」のように後続トークンもキーワード持ちなら保持
    tokens = re.split(r'([ 　]+)', name)          # セパレータを保持して分割
    words   = tokens[0::2]                         # 奇数: テキスト
    seps    = tokens[1::2] + ['']                  # 偶数: セパレータ
    result_words, result_seps = [words[0]], []
    for word, sep in zip(words[1:], seps):
        cumulative = ''.join(result_words)
        is_branch  = bool(re.search(r'^\S{1,4}院$', word))
        has_kw     = _has_clinic_kw(word)
        if _has_clinic_kw(cumulative) and not has_kw and not is_branch:
            break                                  # 説明トークン → 切り捨て
        result_words.append(word)
        result_seps.append(sep)
    # 結合（セパレータを元通りに復元）
    out = result_words[0]
    for w, s in zip(result_words[1:], result_seps):
        out += s + w
    name = out.strip()

    # 7.5. 「病院名 皮膚科・美容皮膚科・...」→ 最初の診療科だけ残す
    m = re.match(r'^(.+?(?:病院|クリニック|医院|診療所))([ 　]+)((?:皮膚科|皮フ科))(・.+)$', name)
    if m:
        name = (m.group(1) + m.group(2) + m.group(3)).strip()

    # 8. 法人名プレフィックス除去
    name = re.sub(r'^独立行政法人\s+', '', name).strip()
    name = re.sub(r'^(?:医療法人(?:社団)?|社会福祉法人|学校法人)\s*\S+\s+', '', name).strip()

    return name.strip()


def clean_address(raw: str) -> str:
    """住所として有効な文字列のみ返す。無効なら空文字を返す。"""
    if not raw or str(raw).strip() in ("", "nan"):
        return ""
    raw = str(raw).strip()

    # 1. 〒XXX-XXXX パターンを含む → 有効
    postal = re.search(r'〒\s*\d{3}[-－ー]\d{4}(?:[^\n\r]{0,50})?', raw)
    if postal:
        addr = postal.group(0).strip()
        # パイプや括弧など不審な文字で切り詰め
        for bad in ['|', '｜', '【', '】', '\n']:
            if bad in addr:
                addr = addr[:addr.index(bad)].strip()
        return addr[:60]

    # 2. 都道府県＋市区町村＋番地のパターン（郵便番号なし）
    m = re.search(
        r'(?:埼玉県|東京都|神奈川県|\S{2,3}[都道府県])\S{1,10}[市区町村]\S{2,30}',
        raw,
    )
    if m:
        addr = m.group(0).strip()
        # 不審な文字・長すぎる・リスト的内容は無効
        if any(c in addr for c in ['|', '｜', '【', '】', '口コミ', 'TOP', 'おすすめ']):
            return ""
        if len(addr) > 60:
            return ""
        return addr

    return ""


# ── 情報抽出ユーティリティ ─────────────────────────────────

def extract_email(text: str) -> str:
    found = re.findall(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}", text)
    found = [e for e in found if not re.search(r"\.(png|jpg|gif|svg|webp)$", e, re.I)]
    return found[0] if found else ""


def extract_phone(text: str) -> str:
    found = re.findall(r"(?:0\d{1,4}[-－ー ]?\d{1,4}[-－ー ]?\d{3,4})", text)
    return found[0] if found else ""


def extract_address(text: str) -> str:
    found = re.findall(r"〒\s*\d{3}[-－]\d{4}[^\n\r]{0,60}", text)
    if found:
        return found[0].strip()
    found = re.findall(r"埼玉県[^\n\r]{5,60}", text)
    return found[0].strip() if found else ""


def extract_instagram(soup: BeautifulSoup) -> str:
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if "instagram.com" in href and "/p/" not in href and "/reel/" not in href:
            if href.startswith("//"):
                href = "https:" + href
            return href.rstrip("/")
    return ""


def find_contact_url(soup: BeautifulSoup, base_url: str) -> str:
    contact_keywords = [
        "contact", "inquiry", "お問い合わせ", "問い合わせ",
        "ご相談", "お申し込み", "メールフォーム",
    ]
    for a in soup.find_all("a", href=True):
        href = a["href"]
        link_text = a.get_text(strip=True)
        hit = any(kw in href.lower() or kw in link_text for kw in contact_keywords)
        if not hit:
            continue
        if href.startswith("http"):
            return href
        if href.startswith("//"):
            return "https:" + href
        if not href.startswith(("#", "mailto:", "tel:")):
            return urljoin(base_url, href)
    return ""


def get_page_title(soup: BeautifulSoup) -> str:
    tag = soup.find("title")
    if not tag:
        return ""
    raw = tag.get_text(strip=True)
    # 「クリニック名 | 説明」形式の場合、前半を名称とする
    for sep in ["|", "｜", "–", "—", " - "]:
        if sep in raw:
            return raw.split(sep)[0].strip()
    return raw


# ── 検索 ──────────────────────────────────────────────────

def search_clinics(keyword: str, max_results: int) -> list[dict]:
    """DuckDuckGo で皮膚科を検索してURLリストを返す。"""
    logger.info(f"検索開始: 「{keyword}」（最大 {max_results} 件）")
    try:
        with DDGS() as ddgs:
            results = list(ddgs.text(keyword, max_results=max_results))
        logger.info(f"検索結果: {len(results)} 件取得")
        return results
    except Exception as e:
        logger.error(f"検索エラー: {e}")
        raise


def filter_results(results: list[dict]) -> list[dict]:
    """まとめサイト・SNS・PDFなどを除外して公式サイトを優先する。"""
    def is_valid(r: dict) -> bool:
        url = r.get("href", "")
        if any(d in url for d in SKIP_DOMAINS):
            return False
        if any(re.search(pat, url, re.I) for pat in SKIP_URL_PATTERNS):
            return False
        return True

    filtered = [r for r in results if is_valid(r)]
    logger.info(f"フィルタリング後: {len(filtered)} 件（除外: {len(results) - len(filtered)} 件）")
    return filtered


# ── スクレイピング ─────────────────────────────────────────

async def fetch_html_playwright(url: str, page) -> tuple[str, str]:
    """
    Playwright でページを読み込み (JS 実行後の HTML, title) を返す。
    失敗時は ("", "") を返す。
    """
    try:
        response = await page.goto(url, timeout=15000, wait_until="domcontentloaded")
        if response and response.status >= 400:
            logger.warning(f"HTTP {response.status}: {url}")
            return "", ""
        await page.wait_for_timeout(1500)
        html = await page.content()
        title = await page.title()
        return html, title
    except Exception as e:
        logger.warning(f"Playwright 取得失敗 {url}: {e}")
        return "", ""


def fetch_html_requests(url: str) -> str:
    """
    requests でHTMLを取得するフォールバック。
    接続エラー時は ConnectionError を再送出する。
    """
    try:
        resp = requests.get(
            url,
            headers={"User-Agent": USER_AGENT},
            timeout=10,
        )
        resp.encoding = resp.apparent_encoding
        return resp.text
    except requests.ConnectionError as e:
        logger.error(f"接続エラー: {url} / {e}")
        raise
    except Exception as e:
        logger.warning(f"requests 取得失敗 {url}: {e}")
        return ""


async def collect_info(url: str, search_title: str, page) -> dict:
    """1件のクリニックURL から7項目の情報を収集して返す。"""
    record = {
        "名称": search_title,
        "メールアドレス": "",
        "公式サイトURL": url,
        "所在地": "",
        "電話番号": "",
        "インスタURL": "",
        "問い合わせフォームURL": "",
    }

    # robots.txt チェック
    if not is_allowed(url):
        return record

    # Playwright でページ取得（JS対応）
    html, pw_title = await fetch_html_playwright(url, page)

    if not html:
        # Playwright 失敗時は requests でフォールバック
        # ここで ConnectionError が発生した場合は呼び出し元に伝播させて停止
        html = fetch_html_requests(url)

    if not html:
        logger.warning(f"HTML 取得できず: {url}")
        return record

    soup = BeautifulSoup(html, "lxml")
    text = soup.get_text(separator="\n")

    # 名称: Playwright title → ページ <title> → 検索タイトル の優先順
    page_title = pw_title or get_page_title(soup)
    if page_title:
        record["名称"] = page_title

    record["メールアドレス"] = extract_email(text)
    record["電話番号"] = extract_phone(text)
    record["所在地"] = extract_address(text)
    record["インスタURL"] = extract_instagram(soup)
    record["問い合わせフォームURL"] = find_contact_url(soup, url)

    return record


# ── メイン ────────────────────────────────────────────────

async def main() -> None:
    logger.info("=" * 50)
    logger.info("埼玉県 皮膚科 情報収集 開始")
    logger.info("=" * 50)
    start_time = datetime.now()

    # ① 検索
    try:
        raw_results = search_clinics(SEARCH_KEYWORD, MAX_SEARCH_RESULTS)
    except Exception:
        logger.error("検索に失敗しました。終了します。")
        sys.exit(1)

    clinics = filter_results(raw_results)
    if not clinics:
        logger.error("収集対象が 0 件です。終了します。")
        sys.exit(1)

    records: list[dict] = []

    # ② Playwright ブラウザ起動
    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=True)
        context = await browser.new_context(user_agent=USER_AGENT)
        page = await context.new_page()

        for i, clinic in enumerate(clinics, 1):
            url = clinic.get("href", "").strip()
            title = clinic.get("title", "")
            if not url:
                continue

            logger.info(f"[{i}/{len(clinics)}] {title}")
            logger.info(f"  URL: {url}")

            try:
                info = await collect_info(url, title, page)
                records.append(info)
                logger.info(
                    f"  結果: 名称={info['名称'][:20]} / "
                    f"TEL={info['電話番号'] or 'なし'} / "
                    f"住所={info['所在地'][:15] + '...' if len(info['所在地']) > 15 else info['所在地'] or 'なし'}"
                )
            except requests.ConnectionError:
                # 接続エラーは収集済みデータを保存して終了
                logger.error("接続エラーが発生しました。収集を中断して結果を保存します。")
                break
            except Exception as e:
                logger.error(f"  予期しないエラー: {e}")
                records.append({
                    "名称": title, "メールアドレス": "", "公式サイトURL": url,
                    "所在地": "", "電話番号": "", "インスタURL": "", "問い合わせフォームURL": "",
                })

            random_wait()

        await browser.close()

    # ③ クリーニング（名称・住所の整形）
    if not records:
        logger.error("収集データが 0 件のため Excel 出力をスキップします。")
        sys.exit(1)

    for rec in records:
        rec["名称"] = clean_name(rec["名称"])
        rec["所在地"] = clean_address(rec["所在地"])

    # ④ Excel 出力
    columns = ["名称", "メールアドレス", "公式サイトURL", "所在地", "電話番号", "インスタURL", "問い合わせフォームURL"]
    df = pd.DataFrame(records, columns=columns)
    df.to_excel(OUTPUT_FILE, index=False)

    elapsed = int((datetime.now() - start_time).total_seconds())
    logger.info("=" * 50)
    logger.info(f"完了: {OUTPUT_FILE} に {len(records)} 件を出力 （所要時間: {elapsed}秒）")
    logger.info("=" * 50)


if __name__ == "__main__":
    asyncio.run(main())
