# %%
# HA1cと血糖値の前60日平均を求めてグラフ化する　20251113確認

import pandas as pd
import openpyxl
from openpyxl.chart import LineChart, Reference
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import timedelta

# ファイルパス
import os

def get_excel_path(filename, folder="ExcelDATA"):
    base = os.path.join(os.environ["OneDrive"], "ドキュメント", "PythonWork")
    return os.path.join(base, folder, filename)

hba1c_path = get_excel_path("データ表4.xlsx", folder="ExcelDATA")
glucose_path = get_excel_path("データ表1.xlsx", folder="ExcelDATA")

# データ読み込み
df_hba1c = pd.read_excel(hba1c_path, sheet_name="Sheet2", usecols=[1, 2]) #1と2の列を読み込む"日付”と"HA1C(%)"
df_glucose = pd.read_excel(glucose_path, sheet_name="Sheet4", usecols=[0, 2]) #"日付”と"血糖値"

# 列名を統一
df_hba1c.columns = ["日付", "HA1c"]
df_glucose.columns = ["日付", "血糖値"]

# 日付をdatetimeに変換
df_hba1c["日付"] = pd.to_datetime(df_hba1c["日付"]).dt.date  # 日付だけ
df_glucose["日付"] = pd.to_datetime(df_glucose["日付"]).dt.date

# 平均血糖値を追加
avg_glucose_list = []
for hba1c_date in df_hba1c["日付"]:
    start_date = hba1c_date - timedelta(days=60)
    mask = (df_glucose["日付"] >= start_date) & (df_glucose["日付"] < hba1c_date)
    avg_glucose = df_glucose.loc[mask, "血糖値"].mean()
    if pd.notna(avg_glucose):
        avg_glucose = round(avg_glucose, 1)
    avg_glucose_list.append(avg_glucose)

df_hba1c["前60日平均血糖値"] = avg_glucose_list

# 日付順に並び替え
df_hba1c = df_hba1c.sort_values("日付")

# Excelファイルに書き出し（openpyxl使用）
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Sheet2"

# DataFrameの内容をワークシートに書き込み
for r in dataframe_to_rows(df_hba1c, index=False, header=True):
    ws.append(r)

# 保存
wb.save(hba1c_path)
# %%
# 2軸のgraphを作成する（通常の折れ線グラフ）

from openpyxl.chart import (
    LineChart,
    Reference,
    Series,
)
from openpyxl.chart.axis import DateAxis, GraphicalProperties

# 最初のグラフオブジェクト（左側Y軸：HA1c）

graph_obj1 = LineChart()
graph_obj1.title = "HA1cと平均血糖値（60日前まで）"
graph_obj1.x_axis.title = "日付"
graph_obj1.y_axis.title = "HA1c(%)"

# 第2軸（右側Y軸：平均血糖値）
graph_obj2 = LineChart()
graph_obj2.y_axis.title = "平均血糖値"

# データ範囲の指定
v1 = Reference(
    ws,
    min_col=2,
    min_row=1,
    max_col=2,
    max_row=ws.max_row
)  # HA1c

h1 = Reference(
    ws,
    min_col=1,
    min_row=2,
    max_col=1,
    max_row=ws.max_row
)  # 日付

v2 = Reference(
    ws,
    min_col=3,
    min_row=1,
    max_col=3,
    max_row=ws.max_row
)  # 前60日平均血糖値

h2 = Reference(
    ws,
    min_col=1,
    min_row=2,
    max_col=1,
    max_row=ws.max_row
)  # 日付

# HA1c（左側Y軸）
graph_obj1.add_data(v1, titles_from_data=True)
graph_obj1.set_categories(h1)

# スムージングを解除（通常の折れ線）
for series in graph_obj1.series:
    series.smooth = False

# 平均血糖値（右側Y軸）
graph_obj2.add_data(v2, titles_from_data=True)
graph_obj2.set_categories(h2)

# スムージングを解除（通常の折れ線）
for series in graph_obj2.series:
    series.smooth = False

# 第2軸の設定
graph_obj2.y_axis.axId = 200
graph_obj2.y_axis.crosses = 'max'
graph_obj2.y_axis.majorGridlines = None

# 平均血糖値の表示範囲
graph_obj2.y_axis.scaling.min = 80
graph_obj2.y_axis.scaling.max = 140

# グラフサイズ
graph_obj1.width = 25
graph_obj1.height = 15

# 第1軸と第2軸を合成
graph_obj1 += graph_obj2

# グラフを貼り付け
ws.add_chart(graph_obj1, "F2")

# Sheet2をアクティブにする
wb.active = wb["Sheet2"]

# 保存
wb.save(hba1c_path)

# 完成したExcelファイルを開く
os.startfile(hba1c_path)

import sys
sys.exit()