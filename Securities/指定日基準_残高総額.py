# %%
#指定日基準保有総額　（川口PC、高萩PC兼用
import os
import re
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
import requests
import yfinance as yf

import argparse

# =========================
# パス設定（PC差を吸収）
# =========================
ONEDRIVE = os.environ.get("OneDrive")
if not ONEDRIVE:
    raise EnvironmentError("環境変数 'OneDrive' が見つかりません。OneDrive同期が有効か確認してください。")

BASE = Path(ONEDRIVE) / "有価証券"

PRICE_DIR = str(BASE / r"01_Excel入力")
HOLDINGS_PATH = str(BASE / r"01_Excel入力\残高\保有総額_指定日基準.xlsx")

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
    print("終値が取得できていない銘柄があります:")
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
    df_merged.to_excel(w, sheet_name="指定日基準_保有総額", index=False)

print("指定日基準_保有総額.xlsx を更新しました。")


#df.to_excel


