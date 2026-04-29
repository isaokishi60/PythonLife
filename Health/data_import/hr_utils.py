# hr_utils.py
# 共通ユーティリティ：夜間心拍データ抽出など

from datetime import datetime, date, timedelta, time
import pandas as pd


def night_window_from_two_days(
    df_two: pd.DataFrame,
    target_date: date,
    start_h: int = 21,
    end_h: int = 6
):
    """
    前日 start_h:00 〜 当日 end_h:00 の心拍データを抽出する。
    df_two は「前日＋当日」の2日分を結合した DataFrame。

    Parameters
    ----------
    df_two : pd.DataFrame
        datetime列とheart_rate列を含む2日分のデータ
    target_date : date
        対象日（当日側）
    start_h : int
        夜間開始（前日側）時刻
    end_h : int
        夜間終了（当日側）時刻

    Returns
    -------
    out : pd.DataFrame
        夜間データ（night_hour列を含む）
    start_dt : datetime
        夜間開始の datetime
    end_dt : datetime
        夜間終了の datetime
    """

    # 夜間の開始・終了時刻
    start_dt = datetime.combine(target_date - timedelta(days=1), time(start_h, 0, 0))
    end_dt = datetime.combine(target_date, time(end_h, 0, 0))

    # DataFrame の整形
    df = df_two.copy()
    df["datetime"] = pd.to_datetime(df["datetime"])
    df = df.sort_values("datetime")

    # 夜間の範囲で抽出
    mask = (df["datetime"] >= start_dt) & (df["datetime"] <= end_dt)
    out = df.loc[mask].copy()

    # 21:00 を 0 とする連続軸（0〜9時間）
    out["night_hour"] = (out["datetime"] - start_dt).dt.total_seconds() / 3600.0

    return out, start_dt, end_dt
