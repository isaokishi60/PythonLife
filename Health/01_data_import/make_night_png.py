import os
from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt

from datetime import datetime
from datetime import timedelta

# 共通関数
from hr_utils import night_window_from_two_days

# ============================
# 1) ベースフォルダを環境変数から取得
# ============================
BASE = os.environ.get("HEALTH_BASE")
if not BASE:
    raise RuntimeError(
        "環境変数 HEALTH_BASE が設定されていません。\n"
        "例: C:\\Users\\spax2\\OneDrive\\ドキュメント\\PythonWork\\Health"
    )

BASE = Path(BASE)

# ============================
# 2) Excel ファイルの場所を自動生成
# ============================
def get_week_xlsx(target_date):
    fname = f"RestHR_{target_date}_week.xlsx"
    return BASE / "01_Garmin_Import" / "outputs" / fname

# ============================
# 3) グラフ作成
# ============================
def make_night_png(target_date):
    # ★ ここで date 型に変換する
    if isinstance(target_date, str):
        target_date = datetime.strptime(target_date, "%Y-%m-%d").date()
    xlsx = get_week_xlsx(target_date)
    if not xlsx.exists():
        raise FileNotFoundError(f"Excel が見つかりません: {xlsx}")

    # Excel のシート名は文字列
    df_today = pd.read_excel(xlsx, sheet_name=str(target_date))

    # datetime を作る
    df_today["datetime"] = pd.to_datetime(df_today["date"].astype(str) + " " + df_today["time"])

    # 前日分のシートも読む
    prev_date = target_date - timedelta(days=1)
    df_prev = pd.read_excel(xlsx, sheet_name=str(prev_date))
    df_prev["datetime"] = pd.to_datetime(df_prev["date"].astype(str) + " " + df_prev["time"])

    # 2日分を結合
    df_two = pd.concat([df_prev, df_today], ignore_index=True)

    # 共通関数で夜間抽出
    df_n, start_dt, end_dt = night_window_from_two_days(df_two, target_date)

    # プロット
    plt.figure()
    plt.plot(df_n["night_hour"], df_n["heart_rate"])
    plt.title(f"Night Heart Rate {target_date} (JST) [21:00–06:00]")
    plt.ylabel("bpm")
    plt.ylim(0, 120)
    plt.tight_layout()

    out_png = BASE / "01_Garmin_Import" / "outputs" / "png" / f"NightHR_{target_date}.png"
    out_png.parent.mkdir(parents=True, exist_ok=True)
    plt.savefig(out_png, dpi=150)
    plt.close()

    print(f"Saved: {out_png}")

# ============================
# 4) 実行
# ============================
if __name__ == "__main__":
    make_night_png(datetime.now().strftime("%Y-%m-%d"))
