## ライブラリ関連
# request... HTTPライブラリ(https://requests-docs-ja.readthedocs.io/en/latest/)
# openpyxl... Excel読み書きライブラリ(https://openpyxl.readthedocs.io/en/stable/index.html)
# BeautifulSoup... HTML,XML解読ライブラリ(https://www.crummy.com/software/BeautifulSoup/bs4/doc/)
# datetime... 日付モジュール(https://docs.python.org/ja/3/library/datetime.html)
import requests, openpyxl
from bs4 import BeautifulSoup
from datetime import datetime

# Excelファイルを新規作成し、シートタイトル「最新天気」
book  = openpyxl.Workbook()
sheet = book.active
sheet.title = '最新天気'

# スクレイピング対象のURL、エクセルファイルに使用する日付のフォーマット設定
TARGET_URL = 'https://weather.yahoo.co.jp/weather/jp/14/4610/14135.html'
DATE       = datetime.today().strftime('%Y%m%d%H%M%S')

# URL先のHTML情報を取得し、スクレイピング用に変換
r    = requests.get(TARGET_URL)
soup = BeautifulSoup(r.text, "html.parser")

# 特定のTABLEタグをクラス属性から。TABLEタグ範囲内のTRタグのvalign属性からmiddleのものを検索
TABLE_TAG = soup.find("table", class_="yjw_table")
TR_TAG = TABLE_TAG.find_all("tr", valign="middle")

# 「週間天気」テーブル内の情報を取得していく
i = 1
for tag in TR_TAG:
    value = None

    # TRタグ内のcolor属性に該当しないエリアであれば、検索をかける
    # #ff3300... 週間天気テーブル内の最高気温のフォント色
    if tag.find(color="#ff3300")==None:
        # 改行箇所があればそれを削除し、TDタグ内のタグ情報をリスト形式で取得
        value = [r.text.replace('\n', '') for r in tag.find_all("td")]
        # 取得したリスト内の情報をエクセルシートに書き込み
        for j, v in enumerate(value, 1):
            sheet.cell(i, j, v)
        print(i,value)
        i+=1

    # 気温の行の場合、検索をかける
    else:
        for rr in range(2):
            double = None
            # 最高気温の場合
            if rr == 0: 
                double = ['最高気温']
                double[len(double):len(double)] = [r.text.replace('\n', '') for r in tag.find_all("font",color="#ff3300")]
            # 最低気温の場合
            else:
                double = ['最低気温']
                double[len(double):len(double)] = [r.text.replace('\n', '') for r in tag.find_all("font",color="#0066ff")]
            # エクセルシートに書き込み
            for j, v in enumerate(double, 1):
                sheet.cell(i, j, v)
            print(i,double)
            i+=1

# 書き込んだ結果をセーブして、エクセルファイルとして出力する
book.save('sample_'+DATE+'.xlsx')