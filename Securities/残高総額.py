# %%
#取得日基準保有総額　（川口PC、高萩PC兼用）

import os
import re
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
import requests
import yfinance as yf

# =========================
# パス設定（PC差を吸収）
# =========================

ONEDRIVE = os.environ.get("OneDrive")
if not ONEDRIVE:
    raise EnvironmentError("環境変数 'OneDrive' が見つかりません。OneDrive同期が有効か確認してください。")

BASE = Path(ONEDRIVE) / "有価証券"

HOLDINGS_PATH = str(BASE / "01_Excel入力" / "残高" / "保有総額.xlsx")


import argparse

def ask_target_date() -> date:
    parser = argparse.ArgumentParser()
    parser.add_argument("--date", help="対象日 (YYYY-MM-DD)")
    args = parser.parse_args()

    if args.date:
        td = datetime.strptime(args.date, "%Y-%m-%d").date()
        print("指定日:", td)
        return td

    print("日付を入力してください（この日の前営業日終値を使います）")

    y = int(input("開始年4桁: ").strip())
    m = int(input("開始月2桁: ").strip())
    d = int(input("開始日2桁: ").strip())

    td = date(y, m, d)
    print("指定日:", td)
    return td


# =========================
# 設定：対象日
# =========================
TARGET_DATE = ask_target_date()

# =========================
# 株価ファイル検索
# =========================

PRICE_HISTORY_DIR = BASE / "02_価格データ" / "株価" / "株価実績"

print(PRICE_HISTORY_DIR )

def find_price_file(target_date): 
    ymd = target_date.strftime("%Y%m%d") 
    exact = PRICE_HISTORY_DIR / f"株価_{ymd}_前営業日終値.xlsx" 
    if exact.exists(): 
        return str(exact) 
    # 保険検索（部分一致） 
    candidates = sorted(PRICE_HISTORY_DIR.glob(f"*{ymd}*終値*.xlsx")) 
    if candidates: 
        return str(candidates[-1]) 
    raise FileNotFoundError(
    f"{ymd} の株価ファイルが見つかりません。フォルダ={PRICE_HISTORY_DIR} / exact={exact}"
    )


import openpyxl
import pandas as pd
import numpy as np

# =========================
# 1. 株価ファイル読み込み
# =========================
price_path = find_price_file(TARGET_DATE)
df_price = pd.read_excel(price_path, sheet_name="Prices")

# SecurityID 型統一（重要）
df_price["SecurityID"] = df_price["SecurityID"].astype(str).str.strip()

# 前営業日取得（安全版）
s = pd.to_datetime(df_price["前営業日"], errors="coerce").dropna()
if s.empty:
    raise ValueError(f"前営業日が株価ファイルにありません: {price_path}")
prev_business_day = s.iloc[0].date()

# =========================
# 2. 保有総額ファイル読み込み
# =========================
balance_path = HOLDINGS_PATH
df_balance = pd.read_excel(balance_path)

# SecurityID 型統一（重要）
df_balance["SecurityID"] = df_balance["SecurityID"].astype(str).str.strip()

# 日付列の時刻削除
if "日付" in df_balance.columns:
    df_balance["日付"] = pd.to_datetime(df_balance["日付"]).dt.date

# =========================
# 3. 株価更新（merge）
# =========================
df_merged = df_balance.merge(
    df_price[["SecurityID", "終値"]],
    on="SecurityID",
    how="left",
    suffixes=("", "_new")
)

# 終値更新（新値優先）
df_merged["終値"] = df_merged["終値_new"].combine_first(df_merged["終値"])
df_merged.drop(columns=["終値_new"], inplace=True)

# =========================
# 4. 基準日セット
# =========================
df_merged["基準日"] = prev_business_day

# =========================
# 5. 数値型強制変換（最重要）
# =========================
for col in ["数量", "終値", "取得額"]:
    if col in df_merged.columns:
        df_merged[col] = (
            df_merged[col]
            .astype(str)
            .str.replace(",", "", regex=False)  # カンマ除去
        )
        df_merged[col] = pd.to_numeric(df_merged[col], errors="coerce")

# =========================
# 6. 終値欠損チェック
# =========================
missing_price = df_merged["終値"].isna()
if missing_price.any():
    print("⚠ 終値が取得できていない銘柄があります:")
    print(df_merged.loc[missing_price, ["SecurityID"]])

# =========================
# 7. 評価額計算
# =========================
df_merged["評価額"] = (
    df_merged["数量"] * df_merged["終値"]
).round(0)

# =========================
# 8. 利益計算
# =========================
df_merged["利益"] = (
    df_merged["評価額"] - df_merged["取得額"]
).round(0)

# NA許可整数型へ
df_merged["評価額"] = df_merged["評価額"].astype("Int64")
df_merged["利益"] = df_merged["利益"].astype("Int64")

# =========================
# 9. Excel 書き戻し
# =========================
with pd.ExcelWriter(balance_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
    df_merged.to_excel(w, sheet_name="保有総額", index=False)

print("✅ 保有総額.xlsx を更新しました。")

# =========================
#履歴追記コード　2026/02/19
# =========================

from pathlib import Path
from datetime import date
import pandas as pd

# 履歴ファイルのパス 
HISTORY_PATH = BASE / "01_Excel入力" / "残高" / "資産推移_History.xlsx" 
HISTORY_SHEET = "History"

# === 合計計算 ===
if "基準日" in df_merged.columns and df_merged["基準日"].notna().any():
    base_dt = pd.to_datetime(df_merged["基準日"].dropna().iloc[0]).date()
else:
    base_dt = date.today()

total_value = pd.to_numeric(df_merged["評価額"], errors="coerce").sum()
total_profit = pd.to_numeric(df_merged["利益"], errors="coerce").sum()
num_items = df_merged["SecurityID"].nunique()

new_row = pd.DataFrame([{
    "Date": base_dt,
    "TotalValue": round(float(total_value), 0),
    "TotalProfit": round(float(total_profit), 0),
    "NumItems": int(num_items)
}])

# === 履歴読み込み ===
if HISTORY_PATH.exists():
    hist = pd.read_excel(HISTORY_PATH, sheet_name=HISTORY_SHEET)
    hist.columns = hist.columns.astype(str).str.strip()
    if "Date" in hist.columns:
        hist["Date"] = pd.to_datetime(hist["Date"]).dt.date
    else:
        hist = pd.DataFrame(columns=["Date","TotalValue","TotalProfit","NumItems"])
else:
    hist = pd.DataFrame(columns=["Date","TotalValue","TotalProfit","NumItems"])


# === 同日上書き or 追記 ===
if (hist["Date"] == base_dt).any():
    hist.loc[hist["Date"] == base_dt, ["TotalValue","TotalProfit","NumItems"]] = \
        new_row[["TotalValue","TotalProfit","NumItems"]].iloc[0].values
else:
    hist = pd.concat([hist, new_row], ignore_index=True)

hist = hist.sort_values("Date").reset_index(drop=True)

with pd.ExcelWriter(HISTORY_PATH, engine="openpyxl", mode="w") as w:
    hist.to_excel(w, sheet_name=HISTORY_SHEET, index=False)

print("✅ 資産推移Historyを更新しました")

# =========================
# 証券会社別履歴
# =========================

BROKER_HISTORY_PATH = (
    BASE / "01_Excel入力" / "残高" / "資産推移_BrokerHistory.xlsx"
)

BROKER_HISTORY_SHEET = "BrokerHistory"

ACCOUNT_MASTER_PATH = (
    BASE / "01_Excel入力" / "マスター" / "口座マスター.xlsx"
)

BROKER_MASTER_PATH = (
    BASE / "01_Excel入力" / "マスター" / "証券会社マスター.xlsx"
)

# =========================
# マスター読み込み
# =========================

df_account = pd.read_excel(ACCOUNT_MASTER_PATH)
df_broker = pd.read_excel(BROKER_MASTER_PATH)

df_account.columns = df_account.columns.astype(str).str.strip()
df_broker.columns = df_broker.columns.astype(str).str.strip()

# 型統一
df_account["AccountID"] = df_account["AccountID"].astype(str).str.strip()
df_account["BrokerID"] = df_account["BrokerID"].astype(str).str.strip()

df_broker["BrokerID"] = df_broker["BrokerID"].astype(str).str.strip()

# =========================
# 保有データへ BrokerID 付与
# =========================

tmp = df_merged.merge(
    df_account[["AccountID", "BrokerID"]],
    on="AccountID",
    how="left"
)

tmp = tmp.merge(
    df_broker[["BrokerID", "証券会社名"]],
    on="BrokerID",
    how="left"
)

# =========================
# 証券会社別集計
# =========================

broker_hist = (
    tmp.groupby(
        ["BrokerID", "証券会社名"],
        as_index=False
    )[["評価額", "利益"]]
    .sum()
)

broker_hist["Date"] = base_dt

broker_hist = broker_hist.rename(columns={
    "証券会社名": "BrokerName",
    "評価額": "TotalValue",
    "利益": "TotalProfit"
})

broker_hist = broker_hist[
    ["Date", "BrokerID", "BrokerName", "TotalValue", "TotalProfit"]
]

# =========================
# 既存履歴読み込み
# =========================

if BROKER_HISTORY_PATH.exists():
    old = pd.read_excel(
        BROKER_HISTORY_PATH,
        sheet_name=BROKER_HISTORY_SHEET
    )

    old.columns = old.columns.astype(str).str.strip()

    if "Date" in old.columns:
        old["Date"] = pd.to_datetime(old["Date"]).dt.date

else:
    old = pd.DataFrame(columns=[
        "Date",
        "BrokerID",
        "BrokerName",
        "TotalValue",
        "TotalProfit"
    ])

# =========================
# 同日削除
# =========================

if len(old) > 0 and "Date" in old.columns:
    old = old[old["Date"] != base_dt]

# =========================
# 追加
# =========================

new_hist = pd.concat(
    [old, broker_hist],
    ignore_index=True
)

new_hist = new_hist.sort_values(
    ["Date", "BrokerID"]
).reset_index(drop=True)

# =========================
# 保存
# =========================

with pd.ExcelWriter(
    BROKER_HISTORY_PATH,
    engine="openpyxl",
    mode="w"
) as w:
    new_hist.to_excel(
        w,
        sheet_name=BROKER_HISTORY_SHEET,
        index=False
    )

print("✅ BrokerHistory 更新完了")

# %%
print(df_merged.columns.tolist())


