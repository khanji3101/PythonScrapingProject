# ライブラリ関連
import requests, openpyxl
from bs4 import BeautifulSoup
from datetime import datetime

book  = openpyxl.Workbook()
sheet = book.active
sheet.title = '最新天気'

TARGET_URL = 'https://weather.yahoo.co.jp/weather/jp/14/4610/14135.html'
DATE       = datetime.today().strftime('%Y%m%d%H%M%S')

r    = requests.get(TARGET_URL)
soup = BeautifulSoup(r.text, "html.parser")

counter = 0
TABLE_TAG = soup.find("table", class_="yjw_table")
TR_TAG = TABLE_TAG.find_all("tr", valign="middle")

i = 1
for tag in TR_TAG:
    value = None
    if tag.find(color="#ff3300")==None:
        value = [r.text.replace('\n', '') for r in tag.find_all("td")]
        for j, v in enumerate(value, 1):
            sheet.cell(i, j, v)
        print(i,value)
        i+=1
    else:
        for rr in range(2):
            double = None
            if rr == 0: 
                double = ['最高気温']
                double[len(double):len(double)] = [r.text.replace('\n', '') for r in tag.find_all("font",color="#ff3300")]
            else:
                double = ['最低気温']
                double[len(double):len(double)] = [r.text.replace('\n', '') for r in tag.find_all("font",color="#0066ff")]
            for j, v in enumerate(double, 1):
                sheet.cell(i, j, v)
            print(i,double)
            i+=1

book.save('sample_'+DATE+'.xlsx')