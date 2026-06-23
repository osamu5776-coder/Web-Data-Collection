"""
Web Data Collector
入力: input.xlsx（A列に検索キーワード）
出力: output.xlsx（7項目を収集）
"""
import re
import time
import sys
import pandas as pd
import requests
from bs4 import BeautifulSoup
from ddgs import DDGS

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    )
}

# 情報収集するサイト内のページ数上限（問い合わせページ探索用）
MAX_SUBPAGES = 3


def search_top_url(keyword: str) -> tuple[str, str]:
    """DuckDuckGoで検索し、上位の(タイトル, URL)を返す。"""
    try:
        with DDGS() as ddgs:
            results = list(ddgs.text(keyword, max_results=5))
        for r in results:
            url = r.get("href", "")
            # SNS・まとめサイト・地図サービスを除外して公式サイトを優先
            skip = ["instagram.com", "facebook.com", "twitter.com", "x.com",
                    "youtube.com", "tabelog.com", "hotpepper.jp", "gurunavi.com",
                    "ekiten.jp", "jalan.net", "rakuten.co.jp", "amazon.co.jp",
                    "wikipedia.org", "google.com", "maps.google"]
            if any(s in url for s in skip):
                continue
            return r.get("title", ""), url
        # フォールバック: 除外なしで最初の結果
        if results:
            return results[0].get("title", ""), results[0].get("href", "")
    except Exception as e:
        print(f"  [検索エラー] {e}")
    return "", ""


def fetch_soup(url: str, timeout: int = 10) -> BeautifulSoup | None:
    """URLのHTMLをBeautifulSoupで返す。失敗時はNone。"""
    try:
        resp = requests.get(url, headers=HEADERS, timeout=timeout)
        resp.encoding = resp.apparent_encoding
        return BeautifulSoup(resp.text, "lxml")
    except Exception:
        return None


def extract_emails(text: str) -> str:
    """テキストからメールアドレスを抽出。"""
    found = re.findall(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}", text)
    # ノイズ除去（画像ファイル名など）
    found = [e for e in found if not e.endswith((".png", ".jpg", ".gif", ".svg"))]
    return found[0] if found else ""


def extract_phone(text: str) -> str:
    """テキストから日本の電話番号を抽出。"""
    # 固定電話・携帯・フリーダイヤル対応
    pattern = r"(?:0\d{1,4}[-－ー]?\d{1,4}[-－ー]?\d{3,4})"
    found = re.findall(pattern, text)
    return found[0] if found else ""


def extract_address(text: str) -> str:
    """テキストから郵便番号付き住所を抽出。"""
    # 〒XXX-XXXX に続く住所
    pattern = r"〒\s*\d{3}[-－]\d{4}[^\n\r]{0,50}"
    found = re.findall(pattern, text)
    if found:
        return found[0].strip()
    # 都道府県から始まる住所
    pref_pattern = (
        r"(?:北海道|青森|岩手|宮城|秋田|山形|福島|茨城|栃木|群馬|埼玉|千葉|東京|神奈川"
        r"|新潟|富山|石川|福井|山梨|長野|岐阜|静岡|愛知|三重|滋賀|京都|大阪|兵庫|奈良"
        r"|和歌山|鳥取|島根|岡山|広島|山口|徳島|香川|愛媛|高知|福岡|佐賀|長崎|熊本"
        r"|大分|宮崎|鹿児島|沖縄)[^\n\r]{5,50}"
    )
    found = re.findall(pref_pattern, text)
    return found[0].strip() if found else ""


def extract_instagram(soup: BeautifulSoup) -> str:
    """ページ内のInstagramリンクを抽出。"""
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if "instagram.com" in href and "/p/" not in href:
            if href.startswith("//"):
                href = "https:" + href
            return href.rstrip("/")
    return ""


def find_contact_url(soup: BeautifulSoup, base_url: str) -> str:
    """問い合わせフォームのURLを探す。"""
    keywords = ["contact", "inquiry", "お問い合わせ", "問い合わせ", "ご相談", "お申し込み"]
    for a in soup.find_all("a", href=True):
        href = a["href"]
        text = a.get_text(strip=True)
        if any(kw in href.lower() or kw in text for kw in keywords):
            if href.startswith("http"):
                return href
            # 相対パスを絶対パスに変換
            from urllib.parse import urljoin
            return urljoin(base_url, href)
    return ""


def collect_info(keyword: str) -> dict:
    """1キーワードの情報をすべて収集して辞書で返す。"""
    result = {
        "名称": keyword,
        "メールアドレス": "",
        "公式サイトURL": "",
        "所在地": "",
        "電話番号": "",
        "インスタURL": "",
        "問い合わせフォームURL": "",
    }

    title, url = search_top_url(keyword)
    if not url:
        print(f"  [スキップ] 検索結果なし")
        return result

    result["公式サイトURL"] = url
    if title and not result["名称"]:
        result["名称"] = title

    print(f"  URL: {url}")

    soup = fetch_soup(url)
    if soup is None:
        print(f"  [スキップ] ページ取得失敗")
        return result

    text = soup.get_text(separator="\n")

    result["メールアドレス"] = extract_emails(text)
    result["電話番号"] = extract_phone(text)
    result["所在地"] = extract_address(text)
    result["インスタURL"] = extract_instagram(soup)
    result["問い合わせフォームURL"] = find_contact_url(soup, url)

    # 問い合わせURLが見つからない場合、サブページを追加探索
    if not result["問い合わせフォームURL"]:
        contact_url = find_contact_url(soup, url)
        result["問い合わせフォームURL"] = contact_url

    return result


def main():
    input_file = "input.xlsx"
    output_file = "output.xlsx"

    try:
        df_input = pd.read_excel(input_file)
    except FileNotFoundError:
        print(f"エラー: {input_file} が見つかりません。")
        print("先に create_sample_input.py を実行してサンプルを作成してください。")
        sys.exit(1)

    keywords = df_input.iloc[:, 0].dropna().astype(str).tolist()
    print(f"{len(keywords)} 件を処理します。\n")

    results = []
    for i, keyword in enumerate(keywords, 1):
        print(f"[{i}/{len(keywords)}] {keyword}")
        info = collect_info(keyword)
        results.append(info)
        print(f"  完了: メール={info['メールアドレス'] or 'なし'} / "
              f"TEL={info['電話番号'] or 'なし'} / "
              f"住所={info['所在地'][:20] + '...' if len(info['所在地']) > 20 else info['所在地'] or 'なし'}")
        print()
        time.sleep(2)  # サーバー負荷軽減

    df_output = pd.DataFrame(results, columns=[
        "名称", "メールアドレス", "公式サイトURL", "所在地",
        "電話番号", "インスタURL", "問い合わせフォームURL"
    ])

    df_output.to_excel(output_file, index=False)
    print(f"完了: {output_file} に出力しました（{len(results)} 件）")


if __name__ == "__main__":
    main()
