# -*- coding: utf-8 -*-
import os
import re
from datetime import datetime, timedelta, date
from pathlib import Path

import pandas as pd
from garminconnect import Garmin

import matplotlib.pyplot as plt


# =========================
# 出力先（両PC対応：spax2 / kawgu）
# =========================
OUT_XLSX = (
    Path.home()
    / "OneDrive" / "ドキュメント" / "PythonWork" / "ExcelDATA"
    / "安静時心拍_時系列_7日.xlsx"
)

DEFAULT_TARGET = date.today()


def input_date(prompt: str, default: date) -> date:
    s = input(f"{prompt} [Enterで {default.isoformat()}]: ").strip()
    if not s:
        return default
    return datetime.strptime(s, "%Y-%m-%d").date()


def _to_dt(x):
    """
    Garminから返る timestamp が
    - epoch(ms)
    - '2026-02-23T12:34:56.000'
    - '2026-02-23 12:34:56'
    など混在し得るので吸収
    """
    if x is None:
        return None

    # epoch (ms / sec)
    if isinstance(x, (int, float)):
        # msっぽい
        if x > 10_000_000_000:
            return datetime.fromtimestamp(x / 1000.0)
        return datetime.fromtimestamp(x)

    s = str(x).strip()
    if not s:
        return None
    # ISOっぽいのを適当に
    s = s.replace("Z", "")
    try:
        return datetime.fromisoformat(s)
    except Exception:
        pass

    # 最後の手段：数字を拾う
    m = re.match(r"(\d{4}-\d{2}-\d{2})[ T](\d{2}:\d{2}:\d{2})", s)
    if m:
        return datetime.fromisoformat(f"{m.group(1)}T{m.group(2)}")

    return None


def fetch_hr_timeseries_for_day(g: Garmin, day: date) -> pd.DataFrame:
    """
    1日分の心拍“時系列”を取る。
    garminconnect のバージョン差を吸収するため、複数候補を順に試す。
    """
    ds = day.isoformat()

    # 候補1: get_heart_rates(ds) の中にタイムラインが入っている場合
    # 候補2: get_heart_rate_data(ds) のような関数がある場合
    # 候補3: get_heart_rate( / get_heart_rate_summary 等) がある場合
    candidates = []

    # 1) get_heart_rate_data があれば最優先で試す
    if hasattr(g, "get_heart_rate_data"):
        candidates.append(("get_heart_rate_data", lambda: g.get_heart_rate_data(ds)))

    # 2) get_heart_rate があれば試す
    if hasattr(g, "get_heart_rate"):
        candidates.append(("get_heart_rate", lambda: g.get_heart_rate(ds)))

    # 3) 既に使っている get_heart_rates を試す（summary＋timelineが入る版もある）
    candidates.append(("get_heart_rates", lambda: g.get_heart_rates(ds)))

    last_err = None
    payload = None
    used = None
    for name, fn in candidates:
        try:
            payload = fn() or {}
            used = name
            break
        except Exception as e:
            last_err = e

    if payload is None:
        raise RuntimeError(f"{ds}: 心拍時系列取得に失敗: {last_err}")

    # --- payload からタイムライン配列を探す ---
    # よくあるキー候補（版によって違う）
    timeline_keys = [
        "heartRateValues", "heartRateValueDescriptors",
        "heartRateSamples", "samples",
        "timeOffsetHeartRateSamples",
        "heartRate"
    ]

    timeline = None

    # まずは直接キーで探す
    for k in timeline_keys:
        if k in payload and isinstance(payload[k], list) and payload[k]:
            timeline = payload[k]
            break

    # それでも無い場合、入れ子をざっくり探す
    if timeline is None:
        for k, v in payload.items():
            if isinstance(v, dict):
                for kk in timeline_keys:
                    if kk in v and isinstance(v[kk], list) and v[kk]:
                        timeline = v[kk]
                        break
            if timeline is not None:
                break

    if timeline is None:
        # 取得自体はできているが、時系列が無いケース
        # ここで payload のキーだけ出して原因調査できるようにする
        keys = list(payload.keys())
        print(f"⚠️ {ds}: 時系列が見つかりません（使用={used} / keys={keys}）")
        return pd.DataFrame(columns=["timestamp", "bpm", "date"])

    # timeline の中身が
    # - [ [timestamp, bpm], ... ]
    # - [{"timestamp":..., "value":...}, ...]
    # - [{"startGMT":..., "heartRate":...}, ...]
    # など混在し得るので吸収
    rows = []
    for item in timeline:
        ts = None
        bpm = None

        if isinstance(item, (list, tuple)) and len(item) >= 2:
            ts, bpm = item[0], item[1]
        elif isinstance(item, dict):
            # timestamp候補
            for tk in ["timestamp", "time", "startTime", "startGMT", "start", "dateTime"]:
                if tk in item:
                    ts = item[tk]
                    break
            # bpm候補
            for bk in ["bpm", "value", "heartRate", "hr"]:
                if bk in item:
                    bpm = item[bk]
                    break

        dt = _to_dt(ts)
        try:
            bpm_i = int(bpm) if bpm is not None else None
        except Exception:
            bpm_i = None

        if dt and bpm_i:
            rows.append((dt, bpm_i))

    df = pd.DataFrame(rows, columns=["timestamp", "bpm"])
    if df.empty:
        print(f"⚠️ {ds}: timelineはあるがデータ化できませんでした（使用={used}）")
        return pd.DataFrame(columns=["timestamp", "bpm", "date"])

    # その日の範囲に絞る（念のため）
    day_start = datetime.combine(day, datetime.min.time())
    day_end = day_start + timedelta(days=1)
    df = df[(df["timestamp"] >= day_start) & (df["timestamp"] < day_end)].copy()

    df["date"] = day
    df = df.sort_values("timestamp").reset_index(drop=True)

    print(f"✅ {ds}: 時系列 {len(df)}点 取得（使用={used}）")
    return df


def main():
    target = input_date("指定日 (YYYY-MM-DD)", DEFAULT_TARGET)
    start = target - timedelta(days=6)
    end = target

    email = os.environ["GARMIN_EMAIL"]
    password = os.environ["GARMIN_PASSWORD"]

    g = Garmin(email, password)
    g.login()

    all_days = []
    cur = start
    while cur <= end:
        df_day = fetch_hr_timeseries_for_day(g, cur)
        all_days.append(df_day)
        cur += timedelta(days=1)

    df_all = pd.concat(all_days, ignore_index=True)
    if df_all.empty:
        print("⚠️ 7日分の時系列データが空でした（Garmin側に時系列が無い/取得API差異の可能性）")
        return

    # 解析しやすいように 1分平均に落とす（重い場合の対策）
    # 不要ならコメントアウトOK
    df_all = df_all.set_index("timestamp").sort_index()
    df_1min = df_all["bpm"].resample("1min").mean().dropna().reset_index()
    df_1min["date"] = df_1min["timestamp"].dt.date

    OUT_XLSX.parent.mkdir(parents=True, exist_ok=True)

    # Excel出力（まとめ）
    with pd.ExcelWriter(OUT_XLSX, engine="openpyxl", mode="w") as w:
        df_1min.to_excel(w, sheet_name="raw_timeseries_1min", index=False)

        # 日別シート（必要なら）
        for d, gdf in df_1min.groupby("date"):
            name = f"daily_{d}"
            gdf.to_excel(w, sheet_name=name[:31], index=False)

    # グラフ（7日分まとめ）
    fig_path = OUT_XLSX.with_suffix(".png")
    plt.figure()
    plt.plot(df_1min["timestamp"], df_1min["bpm"])
    plt.title(f"Heart Rate Timeseries (7 days) ending {target.isoformat()}")
    plt.xlabel("Time")
    plt.ylabel("bpm")
    plt.tight_layout()
    plt.savefig(fig_path, dpi=160)
    plt.close()

    print("✅ 完了")
    print("Excel:", OUT_XLSX)
    print("Graph:", fig_path)


if __name__ == "__main__":
    main()