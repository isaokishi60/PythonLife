# 2026/01/01 から今日までのように、開始日〜終了日をまとめて取れる 期間版



from __future__ import annotations

import os
import sys
import json
import argparse
from pathlib import Path
from datetime import datetime, date, timedelta

import pandas as pd

try:
    import matplotlib.pyplot as plt
    import matplotlib
    matplotlib.rcParams["font.family"] = "Meiryo"
except Exception:
    plt = None


# =========================
# 1) Utils
# =========================
def jst_today() -> date:
    return datetime.now().date()


def ensure_dir(p: Path) -> None:
    p.mkdir(parents=True, exist_ok=True)


def log_print(msg: str) -> None:
    print(msg, flush=True)

#with pd.ExcelWriter(
# =========================
# 2) Garmin
# =========================
import time

def get_garmin_client(email: str | None, password: str | None, max_retries=5):
    try:
        from garminconnect import Garmin
    except Exception as e:
        raise RuntimeError(
            "garminconnect が import できません。venv311で `pip install garminconnect` を確認してください。"
        ) from e

    if not email:
        email = os.environ.get("GARMIN_EMAIL")
    if not password:
        password = os.environ.get("GARMIN_PASSWORD")

    if not email or not password:
        raise RuntimeError(
            "Garminログイン情報がありません。\n"
            "環境変数 GARMIN_EMAIL / GARMIN_PASSWORD を設定してください。"
        )

    g = Garmin(email, password)

    # -------------------------
    # 429 対策：ログインをリトライ
    # -------------------------
    for i in range(max_retries):
        try:
            g.login()
            return g  # 成功したら返す

        except Exception as e:
            msg = str(e)

            # 429 の場合
            if "429" in msg or "Too Many Requests" in msg:
                # Exponential Backoff（Garmin は Retry-After を返さないことが多い）
                wait = (2 ** i) + 1
                print(f"Garmin 429: {wait} 秒待機して再試行します ({i+1}/{max_retries})")
                time.sleep(wait)
                continue

            # 429 以外のエラーは即終了
            raise

    raise RuntimeError("Garminログインがレート制限で失敗しました。時間を空けて再実行してください。")




def fetch_heart_rates(g, d: date) -> dict:
    d_str = d.isoformat()
    if hasattr(g, "get_heart_rates"):
        return g.get_heart_rates(d_str)

    cand = ["get_heart_rate", "get_daily_heart_rate", "get_day_heart_rate"]
    for fn in cand:
        if hasattr(g, fn):
            return getattr(g, fn)(d_str)

    raise RuntimeError("Garminクライアントに心拍取得メソッドが見つかりません。")


# =========================
# 3) 心拍解析
# =========================
def calc_daily_total_beats(hr_values: list) -> int | None:
    if not hr_values:
        return None

    total_beats = 0.0
    prev_ts = None
    prev_hr = None

    for item in hr_values:
        if not isinstance(item, (list, tuple)) or len(item) < 2:
            continue

        ts = item[0]
        hr = item[1]

        if hr is None:
            continue

        if prev_ts is not None and prev_hr is not None:
            dt_sec = (ts - prev_ts) / 1000.0
            if dt_sec > 0:
                total_beats += prev_hr * dt_sec / 60.0

        prev_ts = ts
        prev_hr = hr

    return int(round(total_beats)) if total_beats > 0 else None


def calc_tachy_minutes(hr_values: list, threshold: int = 100) -> float | None:
    if not hr_values:
        return None

    total_sec = 0.0
    prev_ts = None
    prev_hr = None

    for item in hr_values:
        if not isinstance(item, (list, tuple)) or len(item) < 2:
            continue

        ts = item[0]
        hr = item[1]

        if hr is None:
            continue

        if prev_ts is not None and prev_hr is not None:
            dt_sec = (ts - prev_ts) / 1000.0
            if dt_sec > 0 and prev_hr >= threshold:
                total_sec += dt_sec

        prev_ts = ts
        prev_hr = hr

    return round(total_sec / 60.0, 1) if total_sec > 0 else 0.0


def detect_insufficient_measurement(hr_values: list, total_beats: int | None) -> tuple[int, int]:
    valid_points = 0

    for item in hr_values or []:
        if not isinstance(item, (list, tuple)) or len(item) < 2:
            continue
        if item[1] is not None:
            valid_points += 1

    insufficient = 0

    # Garminの2分刻みならフル日で約720点
    if valid_points < 550:
        insufficient = 1

    if total_beats is None or total_beats < 75000:
        insufficient = 1

    return valid_points, insufficient


def build_daily_summary_row(d: date, hr_dict: dict, prev_rhr: int | None) -> tuple[dict, int | None]:
    rhr = hr_dict.get("restingHeartRate")
    hr_max = hr_dict.get("maxHeartRate")
    hr_min = hr_dict.get("minHeartRate")
    rhr7 = hr_dict.get("lastSevenDaysAvgRestingHeartRate")
    hr_values = hr_dict.get("heartRateValues") or []

    total_beats = calc_daily_total_beats(hr_values)
    tachy_minutes = calc_tachy_minutes(hr_values, 100)
    valid_points, insufficient = detect_insufficient_measurement(hr_values, total_beats)

    diff = None
    if prev_rhr is not None and rhr is not None:
        diff = rhr - prev_rhr

    level = 0
    reason = ""

    if rhr is not None and rhr7 is not None and rhr >= rhr7 + 10:
        level = 2
        reason = "RHR↑"

    row = {
        "日付": d,
        "安静時心拍数": rhr,
        "前日差": diff,
        "最大心拍数": hr_max,
        "最小心拍数": hr_min,
        "RHR７日移動平均": rhr7,
        "1日総拍動数": total_beats,
        "頻脈時間(100bpm以上_分)": tachy_minutes,
        "測定点数": valid_points,
        "測定不足": insufficient,
        "警報レベル": level,
        "理由": reason,
    }
    return row, rhr


# =========================
# 4) 月別集計
# =========================
def make_monthly_summary(df: pd.DataFrame) -> pd.DataFrame:
    work = df.copy()
    work["月"] = pd.to_datetime(work["日付"], errors="coerce").dt.to_period("M").astype(str)

    agg = work.groupby("月").agg(
        {
            "安静時心拍数": "mean",
            "RHR７日移動平均": "mean",
            "最大心拍数": "max",
            "最小心拍数": "min",
            "1日総拍動数": ["sum", "mean"],
            "頻脈時間(100bpm以上_分)": ["sum", "mean"],
            "測定不足": "sum",
            "警報レベル": [
                lambda s: (s == 2).sum(),
                lambda s: (s == 1).sum(),
                lambda s: (s == 0).sum(),
            ],
            "日付": "count",
        }
    ).reset_index()

    agg.columns = [
        "月",
        "平均RHR",
        "平均7d",
        "月最大MAX",
        "月最小MIN",
        "総拍動数",
        "1日平均拍動数",
        "頻脈時間合計(分)",
        "1日平均頻脈時間(分)",
        "測定不足日数",
        "赤日数",
        "黄日数",
        "通常日数",
        "日数",
    ]

    for col in ["平均RHR", "平均7d", "1日平均拍動数", "1日平均頻脈時間(分)"]:
        agg[col] = pd.to_numeric(agg[col], errors="coerce").round(1)

    return agg


# =========================
# 5) グラフ
# =========================

def make_beats_chart(df: pd.DataFrame, fig_path: Path) -> None:
    if plt is None:
        return

    import matplotlib.dates as mdates
    import matplotlib.ticker as mticker

    d = df.copy()
    d["日付"] = pd.to_datetime(d["日付"], errors="coerce")
    d["1日総拍動数"] = pd.to_numeric(d["1日総拍動数"], errors="coerce")
    d["測定不足"] = pd.to_numeric(d["測定不足"], errors="coerce")

    # 確定データだけ使う
    d = d.dropna(subset=["日付", "1日総拍動数"]).copy()
    d = d[d["測定不足"] == 0].copy()
    d = d[d["1日総拍動数"] >= 75000].copy()
    d = d.sort_values("日付").copy()

    if d.empty:
        log_print("[WARN] 総拍動数グラフ対象データがありません。")
        return

    # 7日平均
    d["総拍動数_7日平均"] = d["1日総拍動数"].rolling(window=7, min_periods=1).mean()

    # 全期間平均
    mean_val = d["1日総拍動数"].mean()

    plt.figure(figsize=(11, 6))

    # 日次値
    plt.plot(
        d["日付"],
        d["1日総拍動数"],
        label="1日総拍動数"
    )

    # 7日平均
    plt.plot(
        d["日付"],
        d["総拍動数_7日平均"],
        linewidth=2.5,
        label="7日平均"
    )

    # 処置日
    plt.axvline(
        pd.to_datetime("2026-03-19"),
        color="red",
        linestyle="--",
        linewidth=2,
        label="処置日"
    )

    # 全期間平均
    plt.axhline(
        mean_val,
        linestyle="--",
        label=f"全期間平均 {mean_val:,.0f}"
    )

    plt.title("1日総拍動数の推移")
    plt.xlabel("日付")
    plt.ylabel("拍動数 / 日")

    ax = plt.gca()
    ax.xaxis.set_major_locator(mdates.DayLocator(interval=4))
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%m-%d"))
    ax.yaxis.set_major_formatter(mticker.StrMethodFormatter("{x:,.0f}"))

    plt.xticks(rotation=90, fontsize=8)
    plt.legend(loc="upper left")
    plt.tight_layout()
    plt.savefig(fig_path, dpi=160)
    plt.close()


def make_rhr_chart(df: pd.DataFrame, fig_path: Path) -> None:
    if plt is None:
        return

    import matplotlib.dates as mdates

    d = df.copy()
    d["日付"] = pd.to_datetime(d["日付"], errors="coerce")
    d["安静時心拍数"] = pd.to_numeric(d["安静時心拍数"], errors="coerce")
    d["RHR７日移動平均"] = pd.to_numeric(d["RHR７日移動平均"], errors="coerce")
    d = d.dropna(subset=["日付", "安静時心拍数"]).copy()
    d = d.sort_values("日付")

    if d.empty:
        return

    plt.figure(figsize=(11, 6))
    plt.plot(d["日付"], d["安静時心拍数"], label="安静時心拍数")
    plt.plot(d["日付"], d["RHR７日移動平均"], linewidth=2.5, label="7日平均")

    plt.axvline(pd.to_datetime("2026-03-19"), color="red", linestyle="--", linewidth=2, label="処置日")

    plt.title("安静時心拍数の推移")
    plt.xlabel("日付")
    plt.ylabel("心拍数 / 分")

    ax = plt.gca()
    ax.xaxis.set_major_locator(mdates.DayLocator(interval=4))  # ★追加
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%m-%d"))

    plt.xticks(rotation=90, fontsize=8)
    plt.legend(loc="upper left")
    plt.tight_layout()
    plt.savefig(fig_path, dpi=160)
    plt.close()


def make_tachy_chart(df: pd.DataFrame, fig_path: Path) -> None:
    if plt is None:
        return

    import matplotlib.dates as mdates

    d = df.copy()
    d["日付"] = pd.to_datetime(d["日付"], errors="coerce")
    d["頻脈時間(100bpm以上_分)"] = pd.to_numeric(d["頻脈時間(100bpm以上_分)"], errors="coerce")
    d["測定不足"] = pd.to_numeric(d["測定不足"], errors="coerce")
    d = d.dropna(subset=["日付", "頻脈時間(100bpm以上_分)"]).copy()
    d = d[d["測定不足"] == 0].copy()
    d = d.sort_values("日付")

    if d.empty:
        return

    plt.figure(figsize=(11, 6))
    plt.bar(d["日付"], d["頻脈時間(100bpm以上_分)"], label="100bpm以上")

    plt.axvline(pd.to_datetime("2026-03-19"), color="red", linestyle="--", linewidth=2, label="処置日")

    plt.title("頻脈時間（100bpm以上）")
    plt.xlabel("日付")
    plt.ylabel("分 / 日")

    ax = plt.gca()
    ax.xaxis.set_major_locator(mdates.DayLocator(interval=4))  # ★追加
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%m-%d"))

    plt.xticks(rotation=90, fontsize=8)
    plt.legend(loc="upper left")
    plt.tight_layout()
    plt.savefig(fig_path, dpi=160)
    plt.close()


# =========================
# 6) 期間取得
# =========================
def export_period(
    g,
    start_date: date,
    end_date: date,
    out_dir: Path,
    prefix: str = "HeartPeriod",
    save_png: bool = True,
) -> Path:
    ensure_dir(out_dir)
    png_dir = out_dir / "png"
    if save_png:
        ensure_dir(png_dir)

    xlsx_path = out_dir / f"{prefix}_{start_date.isoformat()}_{end_date.isoformat()}.xlsx"

    rows: list[dict] = []
    prev_rhr = None
    cur = start_date

    while cur <= end_date:
        log_print(f"[INFO] Fetch HR: {cur.isoformat()}")
        try:
            hr_dict = fetch_heart_rates(g, cur)
            row, prev_rhr = build_daily_summary_row(cur, hr_dict, prev_rhr)
            rows.append(row)
        except Exception as e:
            log_print(f"[WARN] {cur.isoformat()} 取得失敗: {e}")
        cur += timedelta(days=1)

    if not rows:
        raise RuntimeError("取得できた日次データがありません。")

    df_daily = pd.DataFrame(rows).sort_values("日付").reset_index(drop=True)

    df_daily["1日総拍動数"] = pd.to_numeric(df_daily["1日総拍動数"], errors="coerce")

    df_daily["総拍動数_7日平均"] = (
        df_daily["1日総拍動数"]
        .rolling(window=7, min_periods=1)
        .mean()
        .round(1)
    )

    df_month = make_monthly_summary(df_daily)

    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        df_daily.to_excel(writer, sheet_name="日別集計", index=False)
        df_month.to_excel(writer, sheet_name="月別集計", index=False)

    if save_png and plt is not None:
        rhr_png = png_dir / f"{prefix}_RHR_{start_date.isoformat()}_{end_date.isoformat()}.png"
        beats_png = png_dir / f"{prefix}_DailyBeats_{start_date.isoformat()}_{end_date.isoformat()}.png"
        tachy_png = png_dir / f"{prefix}_Tachy_{start_date.isoformat()}_{end_date.isoformat()}.png"

        make_rhr_chart(df_daily, rhr_png)
        make_beats_chart(df_daily, beats_png)
        make_tachy_chart(df_daily, tachy_png)

        log_print(f"[INFO] RHR chart saved:   {rhr_png}")
        log_print(f"[INFO] Beats chart saved: {beats_png}")
        log_print(f"[INFO] Tachy chart saved: {tachy_png}")

    return xlsx_path


# =========================
# 7) Main
# =========================
def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Garmin HR -> Excel (period)")
    p.add_argument("--start-date", type=str, default="2026-01-01", help="開始日 YYYY-MM-DD")
    p.add_argument("--end-date", type=str, default="", help="終了日 YYYY-MM-DD（省略時=今日）")
    p.add_argument("--out-dir", type=str, default="", help="出力先フォルダ（省略時=outputs_period）")
    p.add_argument("--prefix", type=str, default="HeartPeriod", help="出力ファイル名のprefix")
    p.add_argument("--no-png", action="store_true", help="PNG出力をしない")
    p.add_argument("--email", type=str, default="", help="Garmin email（省略時=環境変数GARMIN_EMAIL）")
    p.add_argument("--password", type=str, default="", help="Garmin password（省略時=環境変数GARMIN_PASSWORD）")
    p.add_argument("--dump-raw", action="store_true", help="終了日の生データjson保存（デバッグ用）")
    return p.parse_args()


def main() -> int:
    args = parse_args()

    start_date = datetime.strptime(args.start_date, "%Y-%m-%d").date()
    end_date = jst_today() if not args.end_date else datetime.strptime(args.end_date, "%Y-%m-%d").date()

    if start_date > end_date:
        raise RuntimeError("開始日が終了日より後です。")

    script_dir = Path(__file__).resolve().parent
    out_dir = Path(args.out_dir) if args.out_dir else (script_dir / "outputs_period")
    ensure_dir(out_dir)

    log_print("=======================================")
    log_print("[START] Garmin HR export (period)")
    log_print(f"[INFO] start_date  = {start_date.isoformat()}")
    log_print(f"[INFO] end_date    = {end_date.isoformat()}")
    log_print(f"[INFO] out_dir     = {out_dir}")
    log_print(f"[INFO] prefix      = {args.prefix}")
    log_print(f"[INFO] png         = {'OFF' if args.no_png else 'ON'}")
    log_print("=======================================")

    g = get_garmin_client(args.email or None, args.password or None)

    if args.dump_raw:
        raw = fetch_heart_rates(g, end_date)
        raw_path = out_dir / f"raw_hr_{end_date.isoformat()}.json"
        raw_path.write_text(json.dumps(raw, ensure_ascii=False, indent=2), encoding="utf-8")
        log_print(f"[INFO] raw saved: {raw_path}")

    xlsx = export_period(
        g=g,
        start_date=start_date,
        end_date=end_date,
        out_dir=out_dir,
        prefix=args.prefix,
        save_png=(not args.no_png),
    )

    log_print(f"[DONE] Excel saved: {xlsx}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as e:
        print(f"[ERROR] {e}", file=sys.stderr)
        raise