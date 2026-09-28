# Garmin 1日心拍数 0:00～24:00 グラフ
# 例:
# python garmin_hr_24h.py --date 2026-09-23

from __future__ import annotations

import os
import sys
import argparse
import time
from pathlib import Path
from datetime import datetime, date

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

# rotation=90,
# ============================================================
# 1. Matplotlib
# ============================================================

plt.rcParams["font.family"] = "Meiryo"
plt.rcParams["axes.unicode_minus"] = False


# ============================================================
# 2. Garmin
# ============================================================

def get_garmin_client(
    email: str | None,
    password: str | None,
    max_retries: int = 5,
):

    try:
        from garminconnect import Garmin

    except Exception as e:
        raise RuntimeError(
            "garminconnect が import できません。"
            "venv311 でインストール状況を確認してください。"
        ) from e

    if not email:
        email = os.environ.get("GARMIN_EMAIL")

    if not password:
        password = os.environ.get("GARMIN_PASSWORD")

    if not email or not password:
        raise RuntimeError(
            "Garminログイン情報がありません。\n"
            "環境変数 GARMIN_EMAIL / GARMIN_PASSWORD "
            "を設定してください。"
        )

    g = Garmin(
        email,
        password,
    )

    for i in range(max_retries):

        try:
            g.login()
            return g

        except Exception as e:

            msg = str(e)

            if (
                "429" in msg
                or "Too Many Requests" in msg
            ):
                wait = (2 ** i) + 1

                print(
                    f"Garmin 429: {wait} 秒待機して再試行します "
                    f"({i + 1}/{max_retries})"
                )

                time.sleep(wait)
                continue

            raise

    raise RuntimeError(
        "Garminログインがレート制限で失敗しました。"
        "時間を空けて再実行してください。"
    )


def fetch_heart_rates(
    g,
    target_date: date,
) -> dict:

    d_str = target_date.isoformat()

    if hasattr(g, "get_heart_rates"):
        return g.get_heart_rates(d_str)

    candidates = [
        "get_heart_rate",
        "get_daily_heart_rate",
        "get_day_heart_rate",
    ]

    for fn in candidates:

        if hasattr(g, fn):
            return getattr(g, fn)(d_str)

    raise RuntimeError(
        "Garminクライアントに"
        "心拍取得メソッドが見つかりません。"
    )


# ============================================================
# 3. heartRateValues → DataFrame
# ============================================================

def make_hr_dataframe(
    target_date: date,
    hr_dict: dict,
) -> pd.DataFrame:

    hr_values = (
        hr_dict.get("heartRateValues")
        or []
    )

    rows = []

    for item in hr_values:

        if (
            not isinstance(item, (list, tuple))
            or len(item) < 2
        ):
            continue

        timestamp_ms = item[0]
        heart_rate = item[1]

        if heart_rate is None:
            continue

        try:
            dt = pd.to_datetime(
                timestamp_ms,
                unit="ms",
                utc=True,
            ).tz_convert(
                "Asia/Tokyo"
            )

        except Exception:
            continue

        # 念のため指定日のデータだけ残す
        if dt.date() != target_date:
            continue

        rows.append(
            {
                "日時": dt,
                "心拍数": heart_rate,
            }
        )

    df = pd.DataFrame(rows)

    if df.empty:
        raise RuntimeError(
            f"{target_date.isoformat()} の"
            "心拍時系列データがありません。"
        )

    df = (
        df.sort_values("日時")
        .reset_index(drop=True)
    )

    return df


# ============================================================
# 4. 0～24時グラフ
# ============================================================

def make_24h_chart(
    df: pd.DataFrame,
    target_date: date,
    png_path: Path,
    activities: list[str] | None = None,
) -> None:

    fig, ax = plt.subplots(
        figsize=(14, 8)
    )

    ax.plot(
        df["日時"],
        df["心拍数"],
        linewidth=1.8,
    )

    # ----------------------------
    # 行動時間帯を表示
    # ----------------------------

    if activities:

        for activity in activities:

            try:
                start_text, end_text, label = activity.split(",", 2)

                start_time = pd.Timestamp(
                    f"{target_date.isoformat()} {start_text}",
                    tz="Asia/Tokyo",
                )

                end_time = pd.Timestamp(
                    f"{target_date.isoformat()} {end_text}",
                    tz="Asia/Tokyo",
                )

                # 行動時間帯を薄い帯で表示
                ax.axvspan(
                    start_time,
                    end_time,
                    alpha=0.15,
                )

                # 行動名をグラフ上部に表示
                middle_time = start_time + (end_time - start_time) / 2

                ax.text(
                    middle_time,
                    0.97,
                    label,
                    transform=ax.get_xaxis_transform(),
                    ha="center",
                    va="top",
                    fontsize=10,
                    rotation=0,
                )

            except Exception as e:

                print(
                    f"[WARN] activity を解釈できません: "
                    f"{activity} ({e})"
                )

    # ----------------------------
    # 横軸を当日 0:00～翌日 0:00 に固定
    # ----------------------------

    start = pd.Timestamp(
        target_date,
        tz="Asia/Tokyo",
    )

    end = start + pd.Timedelta(days=1)

    ax.set_xlim(
        start,
        end,
    )

    # ----------------------------
    # X軸
    # ----------------------------

    ax.xaxis.set_major_locator(
        mdates.HourLocator(
            interval=1,
            tz=start.tz,
        )
    )

    ax.xaxis.set_major_formatter(
        mdates.DateFormatter(
            "%H",
            tz=start.tz,
        )
    )

    # ----------------------------
    # Y軸
    # ----------------------------

    hr_max = pd.to_numeric(
        df["心拍数"],
        errors="coerce",
    ).max()

    if pd.isna(hr_max):
        y_max = 160
    else:
        y_max = max(
            120,
            int(
                ((hr_max + 19) // 20)
                * 20
            ),
        )

    ax.set_ylim(
        0,
        y_max,
    )

    # ----------------------------
    # タイトル・ラベル
    # ----------------------------

    ax.set_title(
        f"Heart Rate {target_date.isoformat()} "
        f"(JST) [00:00-24:00]",
        fontsize=18,
    )

    ax.set_xlabel(
        "Time",
        fontsize=14,
    )

    ax.set_ylabel(
        "bpm",
        fontsize=14,
    )

    ax.grid(
        True,
        alpha=0.25,
    )

    # ----------------------------
    # 基本情報
    # ----------------------------

    valid_hr = pd.to_numeric(
        df["心拍数"],
        errors="coerce",
    ).dropna()

    if not valid_hr.empty:

        info = (
            f"測定点数: {len(valid_hr)}\n"
            f"最小: {valid_hr.min():.0f} bpm\n"
            f"最大: {valid_hr.max():.0f} bpm\n"
            f"平均: {valid_hr.mean():.1f} bpm"
        )

        ax.text(
            0.985,
            0.97,
            info,
            transform=ax.transAxes,
            ha="right",
            va="top",
            fontsize=11,
            bbox=dict(
                boxstyle="round",
                facecolor="white",
                alpha=0.8,
            ),
        )

    fig.tight_layout()

    fig.savefig(
        png_path,
        dpi=160,
    )

    plt.close(fig)


# ============================================================
# 5. Excel保存
# ============================================================

def save_excel(
    df: pd.DataFrame,
    excel_path: Path,
) -> None:

    # Excelはtimezone付きdatetimeを保存できないため、
    # JSTの時刻情報を保ったままtimezoneだけ外す
    out_df = df.copy()

    out_df["日時"] = (
        out_df["日時"]
        .dt.tz_localize(None)
    )

    out_df["時刻"] = (
        out_df["日時"]
        .dt.strftime("%H:%M:%S")
    )

    out_df.to_excel(
        excel_path,
        sheet_name="心拍時系列",
        index=False,
        engine="openpyxl",
    )


# ============================================================
# 6. 引数
# ============================================================

def parse_args() -> argparse.Namespace:

    p = argparse.ArgumentParser(
        description=(
            "Garmin 心拍数を"
            "指定日の0:00～24:00で表示"
        )
    )

    p.add_argument(
        "--date",
        required=True,
        help="対象日 YYYY-MM-DD",
    )

    p.add_argument(
        "--out-dir",
        default="",
        help="出力先フォルダ",
    )

    p.add_argument(
        "--email",
        default="",
        help=(
            "Garmin email "
            "（省略時=環境変数 GARMIN_EMAIL）"
        ),
    )

    p.add_argument(
        "--password",
        default="",
        help=(
            "Garmin password "
            "（省略時=環境変数 GARMIN_PASSWORD）"
        ),
    )

    p.add_argument(
        "--activity",
        action="append",
        default=[],
        help=(
            "行動時間帯を 開始時刻,終了時刻,内容 で指定。"
            "複数回指定可能。"
            '例: --activity "10:00,12:30,畑仕事"'
        ),
    )

    return p.parse_args()


# ============================================================
# 7. main
# ============================================================

def main() -> int:

    args = parse_args()

    target_date = datetime.strptime(
        args.date,
        "%Y-%m-%d",
    ).date()

    base = (
        Path(os.environ["OneDrive"])
        / "ドキュメント"
        / "PythonWork"
        / "Health"
    )

    if args.out_dir:

        out_dir = Path(
            args.out_dir
        )

    else:

        out_dir = (
            base
            / "01_Garmin_Import"
            / "outputs_24h"
        )

    out_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    print(
        "======================================="
    )
    print(
        "[START] Garmin HR 24h"
    )
    print(
        f"[INFO] date = {target_date.isoformat()}"
    )
    print(
        f"[INFO] out  = {out_dir}"
    )
    print(
        "======================================="
    )

    g = get_garmin_client(
        args.email or None,
        args.password or None,
    )

    print(
        "[INFO] Garmin心拍データ取得中..."
    )

    hr_dict = fetch_heart_rates(
        g,
        target_date,
    )

    df = make_hr_dataframe(
        target_date,
        hr_dict,
    )

    print(
        f"[INFO] 心拍測定点数: {len(df)}"
    )

    png_path = (
        out_dir
        / f"HeartRate_24h_{target_date.isoformat()}.png"
    )

    excel_path = (
        out_dir
        / f"HeartRate_24h_{target_date.isoformat()}.xlsx"
    )

    make_24h_chart(
        df,
        target_date,
        png_path,
        activities=args.activity,
    )

    save_excel(
        df,
        excel_path,
    )

    print(
        f"[DONE] PNG  : {png_path}"
    )

    print(
        f"[DONE] Excel: {excel_path}"
    )

    return 0


if __name__ == "__main__":

    try:
        raise SystemExit(
            main()
        )

    except Exception as e:

        print(
            f"[ERROR] {e}",
            file=sys.stderr,
        )

        raise