from __future__ import annotations

import os
import sys
import json
import argparse
from pathlib import Path
from datetime import datetime, date, timedelta, time

import pandas as pd

try:
    import matplotlib.pyplot as plt
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


def safe_sheet_name(name: str) -> str:
    bad = ["\\", "/", "*", "?", ":", "[", "]"]
    for ch in bad:
        name = name.replace(ch, "_")
    return name[:31]


# =========================
# 2) Garmin
# =========================
def get_garmin_client(email: str | None, password: str | None):
    try:
        from garminconnect import Garmin
    except Exception as e:
        raise RuntimeError("garminconnect が import できません。venv311で `pip install garminconnect` を確認してください。") from e

    if not email:
        email = os.environ.get("GARMIN_EMAIL")
    if not password:
        password = os.environ.get("GARMIN_PASSWORD")

    if not email or not password:
        raise RuntimeError(
            "Garminログイン情報がありません。\n"
            "環境変数 GARMIN_EMAIL / GARMIN_PASSWORD を設定してから実行してください。"
        )

    g = Garmin(email, password)
    g.login()
    return g


def fetch_heart_rates(g, d: date) -> dict:
    d_str = d.isoformat()
    if hasattr(g, "get_heart_rates"):
        return g.get_heart_rates(d_str)

    cand = ["get_heart_rate", "get_daily_heart_rate", "get_day_heart_rate"]
    for fn in cand:
        if hasattr(g, fn):
            return getattr(g, fn)(d_str)

    raise RuntimeError("Garminクライアントに心拍取得メソッドが見つかりません（garminconnect差異の可能性）。")


# =========================
# 3) dict -> DataFrame
# =========================
def hr_day_to_df(hr_dict: dict) -> pd.DataFrame:
    # ケースA: {"heartRateValues":[[timestamp_ms, hr], ...]}
    if isinstance(hr_dict, dict):
        if "heartRateValues" in hr_dict and isinstance(hr_dict["heartRateValues"], list):
            out = []
            for r in hr_dict["heartRateValues"]:
                if not r or len(r) < 2:
                    continue
                ts_ms, hr = r[0], r[1]
                try:
                    t = datetime.fromtimestamp(ts_ms / 1000.0)  # ローカル時刻（JST環境ならJST）
                except Exception:
                    continue
                out.append((t, hr))
            return pd.DataFrame(out, columns=["datetime", "heart_rate"])

        # ケースB: {"values":[{"startTimeInSeconds":..., "value":...}, ...]}
        if "values" in hr_dict and isinstance(hr_dict["values"], list):
            out = []
            for r in hr_dict["values"]:
                if not isinstance(r, dict):
                    continue
                hr = r.get("value")
                sec = r.get("startTimeInSeconds")
                ms = r.get("startTimeInMillis")
                t = None
                if sec is not None:
                    try:
                        t = datetime.fromtimestamp(int(sec))
                    except Exception:
                        t = None
                if t is None and ms is not None:
                    try:
                        t = datetime.fromtimestamp(int(ms) / 1000.0)
                    except Exception:
                        t = None
                if t is None:
                    continue
                out.append((t, hr))
            return pd.DataFrame(out, columns=["datetime", "heart_rate"])

    raise RuntimeError("心拍データの形式を解析できません（hr_dictの構造が想定外）。")


def df_add_time_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["date"] = df["datetime"].dt.date
    df["time"] = df["datetime"].dt.strftime("%H:%M:%S")
    return df[["date", "time", "heart_rate", "datetime"]]


# =========================
# 4) Night window (prev 21:00 -> today 06:00) using TWO days data
# =========================
def night_window_from_two_days(df_two: pd.DataFrame, target_date: date, start_h: int = 21, end_h: int = 6) -> tuple[pd.DataFrame, datetime, datetime]:
    start_dt = datetime.combine(target_date - timedelta(days=1), time(start_h, 0, 0))
    end_dt = datetime.combine(target_date, time(end_h, 0, 0))

    df = df_two.copy()
    df["datetime"] = pd.to_datetime(df["datetime"])
    df = df.sort_values("datetime")

    mask = (df["datetime"] >= start_dt) & (df["datetime"] <= end_dt)
    out = df.loc[mask].copy()

    # 21:00 を 0 とする連続軸（0..9）
    out["night_hour"] = (out["datetime"] - start_dt).dt.total_seconds() / 3600.0
    return out, start_dt, end_dt


def save_night_png_from_two_days(df_two: pd.DataFrame, target_date: date, out_png: Path, start_h: int = 21, end_h: int = 6):
    if plt is None:
        return

    df_n, start_dt, end_dt = night_window_from_two_days(df_two, target_date, start_h=start_h, end_h=end_h)
    if df_n.empty:
        log_print(f"[WARN] Night window empty: {target_date.isoformat()} (prev {start_h}:00 -> {end_h}:00)")
        return

    fig = plt.figure()
    plt.plot(df_n["night_hour"], df_n["heart_rate"])

    plt.title(f"Night Heart Rate {target_date.isoformat()} (JST) [{start_h:02d}:00-{end_h:02d}:00]")
    plt.xlabel("Time")
    plt.ylabel("bpm")

    # ★ ここを追加（y軸固定）
    plt.ylim(0, 120)

    # tick: 21,22,23,00,01,...,06
    hours = [(start_h + i) % 24 for i in range(0, (24 - start_h) + end_h + 1)]
    ticks = list(range(0, len(hours)))
    labels = [f"{h:02d}" for h in hours]
    plt.xticks(ticks=ticks, labels=labels)

    plt.tight_layout()
    fig.savefig(out_png, dpi=150)
    plt.close(fig)


# =========================
# 5) Export
# =========================
def export_week(
    g,
    base_date: date,
    out_dir: Path,
    prefix: str = "RestHR",
    save_png: bool = True,
    night_start_h: int = 21,
    night_end_h: int = 6,
) -> Path:
    ensure_dir(out_dir)
    png_dir = out_dir / "png"
    if save_png:
        ensure_dir(png_dir)

    # 7日分（古い→新しい）
    dates = [base_date - timedelta(days=i) for i in range(6, -1, -1)]
    xlsx_path = out_dir / f"{prefix}_{base_date.isoformat()}_week.xlsx"

    # 2日分結合用にキャッシュ（同じ日を何度も取りに行かない）
    df_cache: dict[date, pd.DataFrame] = {}

    def get_df_for_day(d: date) -> pd.DataFrame:
        if d in df_cache:
            return df_cache[d]
        hr_dict = fetch_heart_rates(g, d)
        df = hr_day_to_df(hr_dict)
        if not df.empty:
            df["datetime"] = pd.to_datetime(df["datetime"])
            df = df.sort_values("datetime")
        df_cache[d] = df
        return df

    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        for d in dates:
            log_print(f"[INFO] Fetch HR: {d.isoformat()}")

            # 当日分（Excel用）
            df_today = get_df_for_day(d)
            if df_today.empty:
                log_print(f"[WARN] No HR data: {d.isoformat()}")
                continue

            df_excel = df_add_time_cols(df_today)
            df_excel.to_excel(writer, sheet_name=safe_sheet_name(d.isoformat()), index=False)

            # 夜間PNG用（前日+当日を結合）
            if save_png and plt is not None:
                df_prev = get_df_for_day(d - timedelta(days=1))
                df_two = pd.concat([df_prev, df_today], ignore_index=True)
                out_png = png_dir / f"{prefix}_Night_{d.isoformat()}_{night_start_h:02d}-{night_end_h:02d}.png"
                try:
                    save_night_png_from_two_days(df_two, d, out_png, start_h=night_start_h, end_h=night_end_h)
                except Exception as e:
                    log_print(f"[WARN] Night PNG failed ({d.isoformat()}): {e}")

    return xlsx_path


# =========================
# 6) Main
# =========================
def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Garmin HR -> Excel (week) + Night PNG (prev21->today06)")
    p.add_argument("--base-date", type=str, default="", help="YYYY-MM-DD（省略時=今日）")
    p.add_argument("--out-dir", type=str, default="", help="出力先フォルダ（省略時=outputs）")
    p.add_argument("--prefix", type=str, default="RestHR", help="出力ファイル名のprefix")
    p.add_argument("--no-png", action="store_true", help="PNG出力をしない")
    p.add_argument("--night-start", type=int, default=21, help="夜間開始（前日側）時刻（時）")
    p.add_argument("--night-end", type=int, default=6, help="夜間終了（当日側）時刻（時）")
    p.add_argument("--email", type=str, default="", help="Garmin email（省略時=環境変数GARMIN_EMAIL）")
    p.add_argument("--password", type=str, default="", help="Garmin password（省略時=環境変数GARMIN_PASSWORD）")
    p.add_argument("--dump-raw", action="store_true", help="当日の生データjson保存（デバッグ用）")
    return p.parse_args()


def main() -> int:
    args = parse_args()
    base = jst_today() if not args.base_date else datetime.strptime(args.base_date, "%Y-%m-%d").date()

    script_dir = Path(__file__).resolve().parent
    out_dir = Path(args.out_dir) if args.out_dir else (script_dir / "outputs")
    ensure_dir(out_dir)

    log_print("=======================================")
    log_print("[START] Garmin HR export")
    log_print(f"[INFO] base_date   = {base.isoformat()}")
    log_print(f"[INFO] out_dir     = {out_dir}")
    log_print(f"[INFO] prefix      = {args.prefix}")
    log_print(f"[INFO] png         = {'OFF' if args.no_png else 'ON'}")
    log_print(f"[INFO] night_win   = prev {args.night_start:02d}:00 -> {args.night_end:02d}:00")
    log_print("=======================================")

    g = get_garmin_client(args.email or None, args.password or None)

    if args.dump_raw:
        raw = fetch_heart_rates(g, base)
        raw_path = out_dir / f"raw_hr_{base.isoformat()}.json"
        raw_path.write_text(json.dumps(raw, ensure_ascii=False, indent=2), encoding="utf-8")
        log_print(f"[INFO] raw saved: {raw_path}")

    xlsx = export_week(
        g=g,
        base_date=base,
        out_dir=out_dir,
        prefix=args.prefix,
        save_png=(not args.no_png),
        night_start_h=args.night_start,
        night_end_h=args.night_end,
    )

    log_print(f"[DONE] Excel saved: {xlsx}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as e:
        print(f"[ERROR] {e}", file=sys.stderr)
        raise