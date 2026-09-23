# %%
# make_bisoprolol_report.py
#
# ビソプロロール中止前後の心拍状態を比較する
#
# 基準
#   2026-08-27 : ビソプロロール服薬最終日
#   2026-08-28 : 中止後1日目
#
# 入力
#   1) HeartPeriod_YYYY-MM-DD_YYYY-MM-DD.xlsx / 日別集計
#   2) データ表1.xlsx / Sheet4
#
# 出力
#   Excel : ビソプロロール中止前後比較.xlsx
#   PNG   : Health/bisoprolol_report_png/
#
# 既存の make_ablation_report.py 等は変更しない

from __future__ import annotations

import os
import argparse
from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

# 日本語フォント
plt.rcParams["font.family"] = "Yu Gothic"
plt.rcParams["axes.unicode_minus"] = False


# ============================================================
# 1. 設定
# ============================================================

ONEDRIVE = Path(os.environ["OneDrive"])

BASE_HEALTH = (
    ONEDRIVE
    / "ドキュメント"
    / "PythonWork"
    / "Health"
)

EXCEL_DATA_DIR = (
    ONEDRIVE
    / "ドキュメント"
    / "PythonWork"
    / "ExcelDATA"
)

HEALTH_DATA_PATH = EXCEL_DATA_DIR / "データ表1.xlsx"

HEART_PERIOD_DIR = (
    BASE_HEALTH
    / "01_Garmin_Import"
    / "outputs_period"
)

OUTPUT_DIR = BASE_HEALTH / "bisoprolol_report"
PNG_DIR = OUTPUT_DIR / "png"

OUTPUT_EXCEL = OUTPUT_DIR / "ビソプロロール中止前後比較.xlsx"


# ============================================================
# 2. 引数
# ============================================================

def parse_args():
    parser = argparse.ArgumentParser(
        description="ビソプロロール中止前後の心拍状態を比較"
    )

    parser.add_argument(
        "--start-date",
        required=True,
        help="開始日 YYYY-MM-DD",
    )

    parser.add_argument(
        "--end-date",
        required=True,
        help="終了日 YYYY-MM-DD",
    )

    parser.add_argument(
        "--stop-date",
        default="2026-08-28",
        help="ビソプロロール中止後1日目 YYYY-MM-DD",
    )

    return parser.parse_args()


# ============================================================
# 3. HeartPeriod ファイル
# ============================================================

def get_heart_period_path(start_date, end_date):

    filename = (
        f"HeartPeriod_{start_date}_{end_date}.xlsx"
    )

    path = HEART_PERIOD_DIR / filename

    if not path.exists():
        raise FileNotFoundError(
            f"HeartPeriodファイルが見つかりません。\n{path}"
        )

    return path


# ============================================================
# 4. 心拍データ読み込み
# ============================================================

def load_heart_data(path):

    df = pd.read_excel(
        path,
        sheet_name="日別集計",
    )

    required = [
        "日付",
        "安静時心拍数",
        "最大心拍数",
        "最小心拍数",
        "RHR７日移動平均",
        "1日総拍動数",
        "頻脈時間(100bpm以上_分)",
        "測定不足",
    ]

    missing = [
        col for col in required
        if col not in df.columns
    ]

    if missing:
        raise RuntimeError(
            f"HeartPeriodに必要な列がありません: {missing}"
        )

    df["日付"] = pd.to_datetime(
        df["日付"],
        errors="coerce",
    )

    return df[required].copy()


# ============================================================
# 5. 運動データ読み込み
# ============================================================

def load_activity_data():

    if not HEALTH_DATA_PATH.exists():
        raise FileNotFoundError(
            f"データ表1.xlsxが見つかりません。\n"
            f"{HEALTH_DATA_PATH}"
        )

    df = pd.read_excel(
        HEALTH_DATA_PATH,
        sheet_name="Sheet4",
    )

    required = [
        "日付",
        "中程度運動量（分）",
        "運動消費カロリー",
        "歩数",
    ]

    missing = [
        col for col in required
        if col not in df.columns
    ]

    if missing:
        raise RuntimeError(
            f"Sheet4に必要な列がありません: {missing}"
        )

    df["日付"] = pd.to_datetime(
        df["日付"],
        errors="coerce",
    )

    return df[required].copy()


# ============================================================
# 6. データ結合
# ============================================================

def merge_data(
    heart_df,
    activity_df,
    start_date,
    end_date,
    stop_date,
):

    df = pd.merge(
        heart_df,
        activity_df,
        on="日付",
        how="left",
    )

    start_ts = pd.Timestamp(start_date)
    end_ts = pd.Timestamp(end_date)
    stop_ts = pd.Timestamp(stop_date)

    df = df[
        (df["日付"] >= start_ts)
        & (df["日付"] <= end_ts)
    ].copy()

    df = df.sort_values("日付")

    df["服薬状態"] = np.where(
        df["日付"] < stop_ts,
        "服薬中",
        "中止後",
    )

    # 活動時の心拍上昇幅
    df["MAX-RHR"] = (
        df["最大心拍数"]
        - df["安静時心拍数"]
    )

    return df


# ============================================================
# 7. 前後比較
# ============================================================

def create_summary(df):

    metrics = [
        "安静時心拍数",
        "最大心拍数",
        "MAX-RHR",
        "1日総拍動数",
        "頻脈時間(100bpm以上_分)",
        "中程度運動量（分）",
        "運動消費カロリー",
        "歩数",
    ]

    rows = []

    for metric in metrics:

        before = df.loc[
            df["服薬状態"] == "服薬中",
            metric,
        ].dropna()

        after = df.loc[
            df["服薬状態"] == "中止後",
            metric,
        ].dropna()

        before_mean = before.mean()
        after_mean = after.mean()

        rows.append(
            {
                "指標": metric,
                "服薬中平均": before_mean,
                "中止後平均": after_mean,
                "差": after_mean - before_mean,
                "変化率（%）":
                    (
                        (after_mean / before_mean - 1) * 100
                        if before_mean != 0
                        else np.nan
                    ),
                "服薬中中央値": before.median(),
                "中止後中央値": after.median(),
                "服薬中日数": len(before),
                "中止後日数": len(after),
            }
        )

    return pd.DataFrame(rows)


# ============================================================
# 8. 回帰用データ
# ============================================================

def regression_stats(df, x_col, y_col):

    # 測定不足日は除外
    d = df[
        df["測定不足"].fillna(0) == 0
    ][
        ["服薬状態", x_col, y_col]
    ].dropna()

    rows = []

    for state in ["服薬中", "中止後"]:

        part = d[
            d["服薬状態"] == state
        ]

        if len(part) < 3:
            continue

        x = part[x_col].astype(float)
        y = part[y_col].astype(float)

        slope, intercept = np.polyfit(
            x,
            y,
            1,
        )

        r = x.corr(y)

        rows.append(
            {
                "比較": f"{x_col} → {y_col}",
                "服薬状態": state,
                "データ数": len(part),
                "傾き": slope,
                "切片": intercept,
                "相関係数r": r,
                "R2": r ** 2,
            }
        )

    return pd.DataFrame(rows)


# ============================================================
# 9. 時系列グラフ
# ============================================================

def plot_timeseries(
    df,
    stop_date,
    column,
    ylabel,
    filename,
):

    fig, ax = plt.subplots(
        figsize=(12, 5)
    )

    ax.plot(
        df["日付"],
        df[column],
        marker="o",
        linewidth=1.5,
    )

    ax.axvline(
        pd.Timestamp(stop_date),
        linestyle="--",
        linewidth=2,
        label="ビソプロロール中止後1日目",
    )

    ax.set_title(column)
    ax.set_ylabel(ylabel)
    ax.set_xlabel("日付")

    ax.xaxis.set_major_formatter(
        mdates.DateFormatter("%m/%d")
    )

    ax.grid(
        True,
        alpha=0.3,
    )

    ax.legend()

    fig.autofmt_xdate()

    fig.tight_layout()

    path = PNG_DIR / filename

    fig.savefig(
        path,
        dpi=180,
        bbox_inches="tight",
    )

    plt.close(fig)

    print(f"PNG保存: {path}")


# ============================================================
# 10. 散布図
# ============================================================

def plot_scatter(
    df,
    x_col,
    y_col,
    filename,
):

    fig, ax = plt.subplots(
        figsize=(7, 6)
    )

    # 測定不足日は除外
    work = df[
        df["測定不足"].fillna(0) == 0
    ].copy()

    markers = {
        "服薬中": "o",
        "中止後": "^",
    }

    for state in ["服薬中", "中止後"]:

        part = work[
            work["服薬状態"] == state
        ][
            [x_col, y_col]
        ].dropna()

        if part.empty:
            continue

        ax.scatter(
            part[x_col],
            part[y_col],
            marker=markers[state],
            label=state,
            alpha=0.75,
        )

        if len(part) >= 3:

            x = part[x_col].astype(float)
            y = part[y_col].astype(float)

            slope, intercept = np.polyfit(
                x,
                y,
                1,
            )

            xx = np.linspace(
                x.min(),
                x.max(),
                100,
            )

            yy = slope * xx + intercept

            ax.plot(
                xx,
                yy,
                linewidth=1.5,
            )

    ax.set_xlabel(x_col)
    ax.set_ylabel(y_col)

    ax.set_title(
        f"{x_col} と {y_col}"
    )

    ax.grid(
        True,
        alpha=0.3,
    )

    ax.legend()

    fig.tight_layout()

    path = PNG_DIR / filename

    fig.savefig(
        path,
        dpi=180,
        bbox_inches="tight",
    )

    plt.close(fig)

    print(f"PNG保存: {path}")


# ============================================================
# 11. Excel保存
# ============================================================

def save_excel(
    df,
    summary_df,
    regression_df,
):

    with pd.ExcelWriter(
        OUTPUT_EXCEL,
        engine="openpyxl",
    ) as writer:

        summary_df.to_excel(
            writer,
            sheet_name="前後比較",
            index=False,
        )

        regression_df.to_excel(
            writer,
            sheet_name="運動量との関係",
            index=False,
        )

        df.to_excel(
            writer,
            sheet_name="日別結合データ",
            index=False,
        )

    print(
        f"Excel保存: {OUTPUT_EXCEL}"
    )


# ============================================================
# 12. main
# ============================================================

def main():

    args = parse_args()

    print("====================================")
    print("ビソプロロール中止前後比較")
    print("====================================")
    print("開始日:", args.start_date)
    print("終了日:", args.end_date)
    print("中止後1日目:", args.stop_date)

    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    PNG_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    heart_path = get_heart_period_path(
        args.start_date,
        args.end_date,
    )

    print(
        "HeartPeriod:",
        heart_path,
    )

    heart_df = load_heart_data(
        heart_path
    )

    activity_df = load_activity_data()

    df = merge_data(
        heart_df,
        activity_df,
        args.start_date,
        args.end_date,
        args.stop_date,
    )

    print(
        "結合データ件数:",
        len(df),
    )

    print(
        "服薬中:",
        (df["服薬状態"] == "服薬中").sum(),
        "日",
    )

    print(
        "中止後:",
        (df["服薬状態"] == "中止後").sum(),
        "日",
    )

    summary_df = create_summary(df)

    # ----------------------------
    # 運動量と心拍の関係
    # ----------------------------

    reg1 = regression_stats(
        df,
        "運動消費カロリー",
        "頻脈時間(100bpm以上_分)",
    )

    reg2 = regression_stats(
        df,
        "運動消費カロリー",
        "最大心拍数",
    )

    regression_df = pd.concat(
        [reg1, reg2],
        ignore_index=True,
    )

    # ----------------------------
    # Excel
    # ----------------------------

    save_excel(
        df,
        summary_df,
        regression_df,
    )

    # ----------------------------
    # 時系列
    # ----------------------------

    plot_timeseries(
        df,
        args.stop_date,
        "安静時心拍数",
        "bpm",
        "01_RHR.png",
    )

    plot_timeseries(
        df,
        args.stop_date,
        "最大心拍数",
        "bpm",
        "02_MAX_HR.png",
    )

    plot_timeseries(
        df,
        args.stop_date,
        "1日総拍動数",
        "拍/日",
        "03_Daily_Beats.png",
    )

    plot_timeseries(
        df,
        args.stop_date,
        "頻脈時間(100bpm以上_分)",
        "分/日",
        "04_Tachy_100.png",
    )

    # ----------------------------
    # 散布図
    # ----------------------------

    plot_scatter(
        df,
        "運動消費カロリー",
        "頻脈時間(100bpm以上_分)",
        "05_Calorie_vs_Tachy.png",
    )

    plot_scatter(
        df,
        "運動消費カロリー",
        "最大心拍数",
        "06_Calorie_vs_MAX.png",
    )

    print()
    print("=== 完了しました ===")
    print(
        "Excel:",
        OUTPUT_EXCEL,
    )
    print(
        "PNG:",
        PNG_DIR,
    )


if __name__ == "__main__":
    main()