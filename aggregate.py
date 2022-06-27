import os, sys, glob
sys.path.append(os.path.join(os.path.dirname(__file__), 'site-packages'))
import openpyxl as px

NEW_FILE = "report.xlsx"

# エクセルのファイルを全て取得する
files = glob.glob("./*.xls*")

all_data = []
for f in files:
    # 開いているexcelがあったら無視。all.xlsxがあっても無視。
    if f.startswith('./~')  or f == NEW_FILE:
        continue
    # excelを開いて頂く
    wb=px.load_workbook(f, data_only=True)
    # シートを開いて頂く
    ws = wb.worksheets[0]
    # シートを読み込んで全行取得して頂く
    for row in ws.iter_rows(min_row=2):
        # 不要な行があったら飛ばす。
        if row[0].value is None or \
            not str(row[0].value).strip() or \
            row[1].value is None or \
            row[0].value == 'nanika zyogai sitai mozi':
            continue
        values = []
        # 全セルを舐め回してデータを取得する
        for col in row:
            values.append(col.value)
        # 全セルデータを一つの配列に保存する
        all_data.append(values)

# ここからall.xlsxを作る作業
# print(all_data)
wb = px.Workbook()
ws = wb.worksheets[0]
start_row = 2
start_col = 3
# 全セルデータを順番に書き込み
for y, row in enumerate(all_data):
    for x, cell in enumerate(row):
        ws.cell(row=start_row + y,
                    column=start_col + x,
                    value=all_data[y][x])

#名前を付けて保存
wb.save(NEW_FILE)