"""
山梨県 自動車部品製造会社 リスト作成

データソース:
  - https://www.japia.or.jp/japia/member/
    （日本自動車部品工業会 正会員リストから山梨県本社の企業を抽出）
  - https://www.navitime.co.jp/category/0516001/19/?tags=010168
    （山梨県の「自動車部品・カー用品製造」カテゴリ一覧、全7ページ）
  - 上記で抽出した候補企業ごとに公式サイト・企業情報サイトをWeb検索し、
    (1) 本社が山梨県内にあること（他県本社の支店・工場・営業所を除外）
    (2) 事業内容が自動車部品の製造であること（整備・販売・卸売・リサイクル業を除外）
    の2条件を個別に確認した上で採用。

電話番号は google/libphonenumber (phonenumbers) で日本の市外局番の境界に
従って正しい位置にハイフンを入れて整形する。

出力: 山梨県_自動車部品製造会社リスト.xlsx
"""

import phonenumbers
from phonenumbers import PhoneNumberFormat
import pandas as pd

# (会社名, 所在地, 電話番号（生数字）, 公式サイトURL)
COMPANIES = [
    ("株式会社甲府明電舎", "山梨県中央市中楯825", "0552747910", "https://www.meidensha.co.jp/kof/"),
    ("武甲産業株式会社", "山梨県甲府市和戸町492", "0552332228", "https://www.bukou.com/"),
    ("株式会社アスクテクニカ", "山梨県西八代郡市川三郷町高田610", "0552721151", "http://www.asktechnica.co.jp"),
    ("クリエイティブダイカスト株式会社", "山梨県都留市大幡3677-4", "0554431433", "https://creativediecast.co.jp/"),
    ("株式会社佐藤鋳造", "山梨県都留市玉川654", "0554435501", "https://www.satochuzou.co.jp/"),
    ("芦安精機株式会社", "山梨県南アルプス市有野3582", "0552854402", "http://asiyasuseiki.co.jp/"),
    ("道志ダンパー工業株式会社", "山梨県南都留郡道志村12260番地", "0554451398", "https://doshi-chemical.com/info_damper.html"),
    ("山梨宝栄工業株式会社", "山梨県韮崎市龍岡町下條南割480", "0551229900", "http://hoeikogyo.com/publics/index/24/"),
    ("三栄工業株式会社", "山梨県大月市富浜町鳥沢1845", "0554265321", "https://sanei-industries.com/"),
    ("株式会社富士製作所", "山梨県甲府市落合町817", "0552416001", "https://www.fujiss.co/"),
    ("サンペアー株式会社", "山梨県山梨市市川1292番地", "0553223383", "https://fruits.jp/~sunpear/"),
    ("信濃蚕業韮崎精密株式会社", "山梨県韮崎市富士見2丁目4-13", "0551220162", ""),
    ("サンコールエンジニアリング株式会社", "山梨県南アルプス市戸田970番地", "0552842981", "https://suncall-eng.co.jp/"),
    ("クラウンファスナー株式会社", "山梨県南アルプス市田島818", "0552843140", "https://www.crown-f.co.jp/"),
    ("株式会社吉沢鉄工所", "山梨県甲州市塩山小屋敷2010番地", "0553336010", ""),
    ("有限会社望月製作所", "山梨県南巨摩郡南部町万沢5200", "0556673126", "https://mochizuki-ss.co.jp/"),
    ("株式会社フジミ", "山梨県南都留郡富士河口湖町船津6663-2", "0555238411", ""),
    ("株式会社コーシン", "山梨県北杜市小淵沢町8187", "0551362600", "https://www.koshin-k.co.jp/"),
    ("エモスト株式会社", "山梨県大月市七保町下和田870番地", "0554222243", ""),
    ("有限会社サンテック", "山梨県南都留郡道志村12065番地", "0554522313", "https://www.suntec-pl.co.jp/"),
    ("株式会社信和製作所", "山梨県都留市田野倉290-1", "0554203917", "https://shinwaseisakusyo.com/"),
    ("明友機工株式会社", "山梨県韮崎市龍岡町下條南割字西原466番地", "0551457506", "https://www.e-meiyu.com/"),
    ("三井金属ダイカスト株式会社", "山梨県韮崎市大草町下條西割1200番地", "0551233121", "https://www.mitsui-kinzoku.co.jp/project/diecast/"),
    ("株式会社甲徳マシン", "山梨県甲斐市西八幡4422番地の11", "0552763541", "https://koutokumachin.com/"),
    ("株式会社ユキプラ", "山梨県南都留郡忍野村忍草1139番地", "0555842485", ""),
    ("株式会社フューチャーズクラフト", "山梨県北杜市高根町東井出1333番地1", "0551462842", "https://www.fc-carbon.com/"),
    ("株式会社ミクスター", "山梨県南アルプス市桃園293番地", "0552442039", ""),
    ("山梨大瀬工業株式会社", "山梨県上野原市秋山5840", "0554562121", ""),
    ("株式会社三吉", "山梨県大月市笹子町吉久保756番1", "0554569688", "http://www.miyoshi-nw.com/"),
]

OUTPUT_FILE = "山梨県_自動車部品製造会社リスト.xlsx"


def format_phone(raw: str) -> str:
    """生数字の電話番号を日本の市外局番境界に従って正しい位置にハイフン挿入する。"""
    try:
        num = phonenumbers.parse(raw, "JP")
        return phonenumbers.format_number(num, PhoneNumberFormat.NATIONAL)
    except phonenumbers.NumberParseException:
        return raw


def main() -> None:
    rows = []
    for name, addr, tel, url in COMPANIES:
        rows.append({
            "名称": name,
            "所在地": addr,
            "電話番号": format_phone(tel),
            "公式サイトURL": url,
        })

    df = pd.DataFrame(rows, columns=["名称", "所在地", "電話番号", "公式サイトURL"])
    df.to_excel(OUTPUT_FILE, index=False)
    print(f"完了: {OUTPUT_FILE} に {len(df)} 件を出力")


if __name__ == "__main__":
    main()
