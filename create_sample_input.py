"""
サンプルの input.xlsx を作成するスクリプト。
初回動作確認用。
"""
import pandas as pd

sample_data = {
    "検索キーワード": [
        "スターバックス 新宿",
        "ユニクロ 渋谷",
        "マクドナルド 銀座",
    ]
}

df = pd.DataFrame(sample_data)
df.to_excel("input.xlsx", index=False)
print("input.xlsx を作成しました。A列に検索キーワードを入力してください。")
