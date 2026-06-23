# Web Data Collection

Web 検索でビジネス情報を自動収集し、Excel に出力する Python ツール集です。

## 収集できる項目

| 列 | 内容 |
|----|------|
| ① 名称 | クリニック・店舗名 |
| ② メールアドレス | 公開メールアドレス |
| ③ 公式サイト URL | トップページ URL |
| ④ 所在地 | 住所（郵便番号なし） |
| ⑤ 電話番号 | 代表電話番号 |
| ⑥ インスタ URL | Instagram アカウント URL |
| ⑦ 問い合わせフォーム URL | お問い合わせページ URL |

---

## スクリプト一覧

### 1. `scraper_dermatology.py` — キーワード入力 → 汎用収集

`input.xlsx` の A 列に検索キーワードを入力すると、DuckDuckGo で検索して各キーワードのトップサイトから 7 項目を収集します。

**用途:** 業種・地域を問わず、任意のキーワードで情報収集したいとき

**入力:** `input.xlsx`（A 列にキーワード、1行1件）

```
A列
川越市 皮膚科
越谷市 美容院
大宮 税理士事務所
```

**出力:** `output_YYYYMMDD_HHMMSS.xlsx`

**実行:**
```bash
python scraper_dermatology.py
```

**主な設定（スクリプト冒頭の定数）:**

| 定数 | デフォルト | 説明 |
|------|-----------|------|
| `INPUT_FILE` | `input.xlsx` | 入力ファイル名 |
| `MAX_RESULTS_PER_KEYWORD` | `1` | キーワード1件あたりの収集URL数 |
| `DELAY_MIN` / `DELAY_MAX` | `1.0` / `3.0` | リクエスト間の待機秒数（ランダム） |

---

### 2. `scraper_saitama_all.py` — 埼玉県 皮膚科 全件収集

**埼玉県皮膚科医会** (`saitamahifuka.org`) を一次データソースとして、埼玉県全 72 市区町村の登録クリニックを網羅的に収集します。  
その後、各クリニックの公式サイトを Playwright で訪問して追加情報を取得します。

**用途:** 特定地域・特定業種を「できる限り全件」収集したいとき

**入力:** なし（データソースは自動取得）

**出力:** `saitama_all_YYYYMMDD_HHMMSS.xlsx`

**実行:**
```bash
python scraper_saitama_all.py
```

**収集フロー:**

```
Phase 1A  全72市区町村のリストページを巡回 → detail ID を収集
              (requests + BeautifulSoup)
    ↓
Phase 1B  各 detail ページから 名称・電話・住所・公式URL を取得
              (requests + BeautifulSoup)
    ↓
Phase 2   各公式サイトを Playwright で開いて メール・インスタ・問合せURL を抽出
              (async Playwright + BeautifulSoup)
    ↓
Excel 出力 (重複排除済み)
```

**実績値（2026年6月時点）:**

| 項目 | 件数 | 取得率 |
|------|------|--------|
| 名称・電話・住所 | 85 件 | 100% |
| 公式サイト URL | 63 件 | 74% |
| インスタ URL | 11 件 | 13% |
| 問い合わせ URL | 15 件 | 18% |

---

### 3. `collector.py` — シンプル版（Playwright なし）

requests + BeautifulSoup のみで動作する軽量版。JS を使わないシンプルなサイト向け。

**実行:**
```bash
python collector.py
```

---

## セットアップ

### 前提条件

- Python 3.11 以上
- Windows 10/11（他 OS でも動作するが未検証）

### インストール手順

```bash
# 1. 仮想環境を作成・有効化
python -m venv .venv
.venv\Scripts\activate        # Windows
# source .venv/bin/activate   # Mac/Linux

# 2. 依存ライブラリをインストール
pip install -r requirements.txt

# 3. Playwright のブラウザをインストール（初回のみ）
playwright install chromium
```

### 依存ライブラリ

| ライブラリ | 用途 |
|-----------|------|
| `ddgs` | DuckDuckGo 検索（API キー不要） |
| `playwright` | JavaScript 対応ブラウザ自動化 |
| `beautifulsoup4` + `lxml` | HTML パース |
| `requests` | HTTP 通信（フォールバック） |
| `pandas` + `openpyxl` | Excel 入出力 |

---

## ディレクトリ構成

```
Web Data Collection/
├── README.md
├── CLAUDE.md                    # AI エージェント向け設定
├── .gitignore
├── requirements.txt
│
├── scraper_dermatology.py       # キーワード入力→汎用収集（メイン）
├── scraper_saitama_all.py       # 埼玉県皮膚科 全件収集
├── collector.py                 # シンプル版（Playwright なし）
├── create_sample_input.py       # input.xlsx サンプル生成
│
├── input.xlsx                   # 検索キーワード入力ファイル
│
├── output_*.xlsx                # scraper_dermatology.py の出力（除外）
├── saitama_all_*.xlsx           # scraper_saitama_all.py の出力（除外）
├── scraper.log                  # 実行ログ（除外）
│
└── .venv/                       # 仮想環境（除外）
```

---

## 共通の設計方針

- **robots.txt 遵守:** アクセス前にドメインごとの robots.txt を取得・キャッシュし、禁止パスにはアクセスしない
- **ランダム待機:** リクエスト間に 1〜3 秒のランダム遅延を挿入してサーバー負荷を軽減
- **接続エラー対応:** 接続エラー発生時はそれまでの収集データを保存して終了
- **JS 対応:** Playwright（ヘッドレス Chromium）で動的ページを読み込み、失敗時は requests にフォールバック
- **名称クリーニング:** ページタイトルから「○○市 △△クリニック｜診療内容」などのサブタイトル・地名プレフィックスを除去してクリニック名のみを抽出
- **住所クリーニング:** 郵便番号（〒）を除去し、住所部分のみを出力

---

## 注意事項

- 収集対象サイトの**利用規約を確認**の上で使用してください
- 大量のリクエストを短時間に送ることは避けてください
- `.env` や認証情報を含むファイルは絶対にコミットしないでください
- 収集したデータは `data/` フォルダや `output_*.xlsx` として保存され、Git 管理外となっています
