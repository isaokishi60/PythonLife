# データ表1.xlsx の Sheet4 から、指定項目の健康グラフを1つ作成する
# Power Automate対応版
#
# 実行例:
# python "Graph_from_Excel(Complete_2025_11_03).py" --start-date 2025-10-01 --end-date 2026-06-17 --item 1
import os
import argparse
import datetime

from openpyxl import load_workbook
from openpyxl.chart import Reference, LineChart
from openpyxl.chart.axis import DateAxis
from openpyxl.utils import get_column_letter


# =========================
# 引数
# =========================

parser = argparse.ArgumentParser()
parser.add_argument("--start-date", required=True, help="開始日 YYYY-MM-DD")
parser.add_argument("--end-date", required=True, help="終了日 YYYY-MM-DD")
parser.add_argument("--item", required=True, type=int, help="1:体重 2:血糖値 3:血圧 4:中程度運動 5:運動消費 6:歩数")

args = parser.parse_args()

real_start = datetime.datetime.strptime(args.start_date, "%Y-%m-%d").date()
real_end = datetime.datetime.strptime(args.end_date, "%Y-%m-%d").date()
Item_input = args.item

DayOne = datetime.timedelta(days=1)

# 元コードと同じく、不等号判定用に前後1日ずらす
dt_start_input = real_start - DayOne
dt_end_input = real_end + DayOne

print("開始日", real_start)
print("終了日", real_end)
print("項目", Item_input)

if real_end < real_start:
    raise ValueError("終了日が開始日より前です")

if Item_input not in [1, 2, 3, 4, 5, 6]:
    raise ValueError("item は 1～6 を指定してください")


# =========================
# パス設定
# =========================

def get_excel_path(filename, folder="ExcelDATA"):
    base = os.path.join(os.environ["OneDrive"], "ドキュメント", "PythonWork")
    return os.path.join(base, folder, filename)


filepath = get_excel_path("データ表1.xlsx", folder="ExcelDATA")


# =========================
# 項目定義
# =========================

ITEMS = {
    1: {
        "name": "体重",
        "headers": ["体重"],
        "sheet4_cols": [2],
        "ylabel": "体重",
    },
    2: {
        "name": "血糖値",
        "headers": ["血糖値"],
        "sheet4_cols": [3],
        "ylabel": "血糖値",
    },
    3: {
        "name": "血圧",
        "headers": ["血圧収縮期", "血圧拡張期", "心拍数"],
        "sheet4_cols": [4, 5, 6],
        "ylabel": "血圧",
    },
    4: {
        "name": "中程度運動",
        "headers": ["中程度運動"],
        "sheet4_cols": [7],
        "ylabel": "中程度運動(分)",
    },
    5: {
        "name": "運動消費エネルギー",
        "headers": ["運動消費エネルギー"],
        "sheet4_cols": [8],
        "ylabel": "運動消費カロリー",
    },
    6: {
        "name": "歩数",
        "headers": ["歩数"],
        "sheet4_cols": [9],
        "ylabel": "歩数",
    },
}

ITEM = ITEMS[Item_input]


# =========================
# Excel読み込み
# =========================

wb = load_workbook(filename=filepath)

if "Sheet4" not in wb.sheetnames:
    raise ValueError("データ表1.xlsx に Sheet4 がありません")

ws4 = wb["Sheet4"]

# Sheet2 / Sheet3 を毎回作り直す
for sheet_name in ["Sheet2", "Sheet3"]:
    if sheet_name in wb.sheetnames:
        wb.remove(wb[sheet_name])

ws2 = wb.create_sheet("Sheet2")
ws3 = wb.create_sheet("Sheet3")

ws2.column_dimensions["A"].width = 12
ws3.column_dimensions["A"].width = 12


# =========================
# Sheet2 見出し
# =========================

ws2.cell(row=1, column=1).value = "日付"

for i, header in enumerate(ITEM["headers"], start=2):
    ws2.cell(row=1, column=i).value = header
    ws2.column_dimensions[get_column_letter(i)].width = 15


# =========================
# Sheet4 → Sheet2 転記
# =========================

read_row = 2
write_row = 2
k = 0

while True:
    k += 1
    if k > 5000:
        print("5000行を超えたため停止")
        break

    date_value = ws4.cell(row=read_row, column=1).value

    if date_value is None:
        break

    if isinstance(date_value, datetime.datetime):
        date3 = date_value.date()
    elif isinstance(date_value, datetime.date):
        date3 = date_value
    else:
        read_row += 1
        continue

    if dt_start_input < date3 < dt_end_input:
        ws2.cell(row=write_row, column=1).value = date3

        for i, src_col in enumerate(ITEM["sheet4_cols"], start=2):
            ws2.cell(row=write_row, column=i).value = ws4.cell(row=read_row, column=src_col).value

        write_row += 1

    read_row += 1

print("転記行数:", write_row - 2)

if write_row <= 2:
    raise ValueError("指定期間のデータがありません")


# =========================
# Sheet2 → Sheet3 転記
# =========================

for r in range(1, ws2.max_row + 1):
    for c in range(1, len(ITEM["headers"]) + 2):
        ws3.cell(row=r, column=c).value = ws2.cell(row=r, column=c).value

for c in range(1, len(ITEM["headers"]) + 2):
    ws3.column_dimensions[get_column_letter(c)].width = 15


# =========================
# 折れ線グラフ作成
# =========================

graph_obj = LineChart()

values = Reference(
    ws3,
    min_row=1,
    min_col=2,
    max_row=ws3.max_row,
    max_col=len(ITEM["headers"]) + 1,
)

graph_obj.add_data(values, titles_from_data=True)

x_axis = Reference(
    ws3,
    min_col=1,
    min_row=2,
    max_row=ws3.max_row,
)

graph_obj.set_categories(x_axis)

graph_obj.y_axis.title = ITEM["ylabel"]
graph_obj.y_axis.crossAx = 500

graph_obj.x_axis = DateAxis(crossAx=500)
graph_obj.x_axis.number_format = "m/d"
graph_obj.x_axis.title = "年月日"

graph_obj.title = f"{ITEM['name']} ({real_start} ～ {real_end})"

graph_obj.anchor = "F8"
graph_obj.width = 30
graph_obj.height = 16

# 線とマーカー
colors = ["FF0000", "00AA00", "0000FF"]

for i, ser in enumerate(graph_obj.ser):
    color = colors[i % len(colors)]
    ser.graphicalProperties.line.solidFill = color
    ser.graphicalProperties.line.width = 15000
    ser.marker.symbol = "dot"
    ser.marker.graphicalProperties.line.solidFill = color

ws3.add_chart(graph_obj)


# =========================
# 保存
# =========================

wb.save(filepath)
wb.close()

print("グラフ作成完了:", ITEM["name"])
print(filepath)



