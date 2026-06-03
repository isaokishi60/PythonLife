# %%
#日にちを指定して為替レートを取得（高萩PC、川口PC兼用）

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
FX_PATH = str(BASE / r"02_価格データ\為替\DimFx.xlsx")
FX_SHEET = "DimFx"
OUT_DIR = str(BASE / r"02_価格データ")

# =========================
# 指定日の USDJPY を取得
# =========================
def get_usdjpy_on_or_before(target_date):

    start = target_date - timedelta(days=10)
    end = target_date + timedelta(days=1)

    df = yf.download(
        "JPY=X",
        start=start.strftime("%Y-%m-%d"),
        end=end.strftime("%Y-%m-%d"),
        progress=False,
    )

    if df.empty:
        raise ValueError(f"USDJPY データ取得失敗: {target_date}")

    # ★ MultiIndex を解除（これが重要）
    df.columns = [col[0] if isinstance(col, tuple) else col for col in df.columns]

    df = df.reset_index()

    # Date を date 型に
    df["Date"] = pd.to_datetime(df["Date"]).dt.date

    # 指定日以前に絞る
    df = df[df["Date"] <= target_date].sort_values("Date")

    if df.empty:
        raise ValueError(f"{target_date} 以前の為替がありません")

    row = df.iloc[-1]

    fx = float(row["Close"])

    # ★ ここはもう安全に動く
    used_date = row["Date"]

    return fx, used_date




# =========================
# メイン処理
# =========================
def main():

    print("=== 指定日の USDJPY を DimFx に追記します ===")

    parser = argparse.ArgumentParser()
    parser.add_argument("--date", help="対象日 (YYYY-MM-DD)")
    args = parser.parse_args()

    if args.date:
        target_date = datetime.strptime(args.date, "%Y-%m-%d").date()
    else:
        y = int(input("年(YYYY): "))
        m = int(input("月(MM): "))
        d = int(input("日(DD): "))
        target_date = datetime(y, m, d).date()

    print("指定日:", target_date)

    # 為替取得
    fx, used_date = get_usdjpy_on_or_before(target_date)
    print(f"取得した為替: {used_date} USDJPY={fx:.4f}")

    # 既存 DimFx 読み込み（無ければ新規作成）
    try:
        df_fx = pd.read_excel(FX_PATH, sheet_name=FX_SHEET)
    except FileNotFoundError:
        df_fx = pd.DataFrame(columns=["Date", "USDJPY"])

    # ★ 既存の Date 列を必ず date 型に統一（時刻を消す）
    if "Date" in df_fx.columns:
        df_fx["Date"] = pd.to_datetime(df_fx["Date"]).dt.date

    # ★ レートを小数点以下1桁に丸める
    fx = round(fx, 1)

    # 追記データ
    new_row = pd.DataFrame({
        "Date": [used_date],
        "USDJPY": [fx]
    })

    # 結合（追記）
    df_fx = pd.concat([df_fx, new_row], ignore_index=True)

    # 保存
    with pd.ExcelWriter(FX_PATH, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
        df_fx.to_excel(w, sheet_name=FX_SHEET, index=False)

    # ★ 列幅を調整する（Date 列を広く） 
    from openpyxl import load_workbook

    wb = load_workbook(FX_PATH) 
    ws = wb[FX_SHEET] 
    
    # A列（Date）の幅を広げる 
    ws.column_dimensions["A"].width = 15 # お好みで 12〜20 くらい

    wb.save(FX_PATH) 

    print("DimFx.xlsx に追記しました。")
    print(FX_PATH)


if __name__ == "__main__":
    main()


