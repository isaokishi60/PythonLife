# %%
"""
夜間心拍数から、トイレに起きた可能性のある心拍上昇イベントを検出する。

入力:
    OneDrive\ドキュメント\PythonWork\Health
    \01_Garmin_Import\outputs
    \RestHR_YYYY-MM-DD_week.xlsx

各Excelファイル:
    7日分のシートを持つ。
    ファイル名の日付と同じシートを、その日の代表データとして読む。

対象時間:
    00:00:00以上、06:00:00未満

出力:
    OneDrive\ドキュメント\PythonWork\Health
    \07_BPH\toilet_count_data.xlsx
"""

import os
import re
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

# ============================================================
# パス設定
# ============================================================

ONEDRIVE = Path(os.environ["OneDrive"])

INPUT_DIR = (
    ONEDRIVE
    / "ドキュメント"
    / "PythonWork"
    / "Health"
    / "01_Garmin_Import"
    / "outputs"
)

OUTPUT_DIR = (
    ONEDRIVE
    / "ドキュメント"
    / "PythonWork"
    / "Health"
    / "07_BPH"
)

OUTPUT_PATH = OUTPUT_DIR / "toilet_count_data.xlsx"

FILE_PATTERN = "RestHR_*_week.xlsx"


# ============================================================
# 元データの列名
# ============================================================

DATETIME_COLUMN = "datetime"
HEART_RATE_COLUMN = "heart_rate"


# ============================================================
# 対象時間
# ============================================================

START_HOUR = 0
END_HOUR = 6


# ============================================================
# 心拍上昇イベントの判定条件
# ============================================================

# 周辺の基準心拍数より何拍以上高ければ、上昇候補とするか
RISE_THRESHOLD = 8

# 心拍上昇イベントとして認める最短時間
MIN_EVENT_DURATION_MIN = 1

# 長時間の心拍上昇は、トイレ以外の活動の可能性があるため除外
MAX_EVENT_DURATION_MIN = 30

# 離れた上昇イベントを、同じ1回としてまとめる時間
MERGE_GAP_MIN = 12

# 基準心拍数を求める移動中央値の範囲
BASELINE_WINDOW_MIN = 60

# 基準値計算に最低限必要な有効データ数
MIN_BASELINE_POINTS = 10

# 連続イベント判定時に許容するデータ間隔
# Garminデータが2分間隔なので、4分までを連続として扱う
MAX_CONTINUOUS_GAP_MIN = 5

# 短い欠測を補間する最大時間
INTERPOLATE_LIMIT_MIN = 4


# ============================================================
# ファイル名から対象日を取得
# ============================================================

def extract_target_date_from_filename(file_path: Path) -> str | None:
    """
    RestHR_YYYY-MM-DD_week.xlsxからYYYY-MM-DDを取り出す。
    """

    pattern = re.compile(
        r"^RestHR_(\d{4}-\d{2}-\d{2})_week\.xlsx$",
        re.IGNORECASE,
    )

    match = pattern.match(file_path.name)

    if match is None:
        return None

    return match.group(1)


# ============================================================
# 1ファイルから最終日シートを読み込む
# ============================================================

def read_target_sheet(
    file_path: Path,
    target_date_text: str,
) -> pd.DataFrame:
    """
    ファイル名の日付と同じシートを読み込む。
    """

    excel_file = pd.ExcelFile(
        file_path,
        engine="openpyxl",
    )

    if target_date_text not in excel_file.sheet_names:
        raise ValueError(
            f"対象シートがありません: {target_date_text}"
        )

    df = pd.read_excel(
        file_path,
        sheet_name=target_date_text,
        engine="openpyxl",
    )

    required_columns = {
        DATETIME_COLUMN,
        HEART_RATE_COLUMN,
    }

    missing_columns = required_columns - set(df.columns)

    if missing_columns:
        raise ValueError(
            f"必要な列がありません: {sorted(missing_columns)}"
        )

    df = df[
        [
            DATETIME_COLUMN,
            HEART_RATE_COLUMN,
        ]
    ].copy()

    df[DATETIME_COLUMN] = pd.to_datetime(
        df[DATETIME_COLUMN],
        errors="coerce",
    )

    df[HEART_RATE_COLUMN] = pd.to_numeric(
        df[HEART_RATE_COLUMN],
        errors="coerce",
    )

    df = df.dropna(
        subset=[
            DATETIME_COLUMN,
            HEART_RATE_COLUMN,
        ]
    )

    target_date = pd.Timestamp(target_date_text).normalize()

    # 念のため、ファイル名の日付と一致する行だけを残す
    df = df[
        df[DATETIME_COLUMN].dt.normalize() == target_date
    ].copy()

    df = df.sort_values(DATETIME_COLUMN)

    df = df.drop_duplicates(
        subset=[DATETIME_COLUMN],
        keep="last",
    )

    return df


# ============================================================
# 全ファイルの読み込み
# ============================================================

def load_all_daily_files(
    input_dir: Path,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    RestHR_*_week.xlsxを検索し、
    各ファイルの日付と同名のシートだけを読み込む。
    """

    files = sorted(input_dir.glob(FILE_PATTERN))

    if not files:
        raise FileNotFoundError(
            "対象ファイルがありません。\n"
            f"検索フォルダー: {input_dir}\n"
            f"検索条件: {FILE_PATTERN}"
        )

    all_data: list[pd.DataFrame] = []
    log_rows: list[dict] = []

    print(f"対象ファイル数: {len(files)}")

    for file_path in files:
        target_date_text = extract_target_date_from_filename(
            file_path
        )

        if target_date_text is None:
            log_rows.append(
                {
                    "ファイル名": file_path.name,
                    "対象日": pd.NaT,
                    "シート名": "",
                    "読込結果": "ファイル名形式不一致",
                    "読込件数": 0,
                }
            )
            continue

        try:
            df = read_target_sheet(
                file_path,
                target_date_text,
            )

            all_data.append(df)

            log_rows.append(
                {
                    "ファイル名": file_path.name,
                    "対象日": pd.Timestamp(target_date_text),
                    "シート名": target_date_text,
                    "読込結果": "正常",
                    "読込件数": len(df),
                }
            )

            print(
                f"読込: {file_path.name} "
                f"/ {target_date_text} "
                f"/ {len(df)}件"
            )

        except Exception as exc:
            log_rows.append(
                {
                    "ファイル名": file_path.name,
                    "対象日": pd.Timestamp(target_date_text),
                    "シート名": target_date_text,
                    "読込結果": f"エラー: {exc}",
                    "読込件数": 0,
                }
            )

            print(
                f"読込エラー: {file_path.name} "
                f"/ {exc}"
            )

    if not all_data:
        raise ValueError(
            "正常に読み込めた心拍数データがありません。"
        )

    heart_rate_df = pd.concat(
        all_data,
        ignore_index=True,
    )

    heart_rate_df = heart_rate_df.sort_values(
        DATETIME_COLUMN
    )

    heart_rate_df = heart_rate_df.drop_duplicates(
        subset=[DATETIME_COLUMN],
        keep="last",
    )

    file_log_df = pd.DataFrame(log_rows)

    return heart_rate_df, file_log_df


# ============================================================
# 夜間データ抽出
# ============================================================

def extract_night_data(
    df: pd.DataFrame,
) -> pd.DataFrame:
    """
    00:00以上06:00未満のデータだけを抽出する。
    """

    night_df = df[
        (df[DATETIME_COLUMN].dt.hour >= START_HOUR)
        & (df[DATETIME_COLUMN].dt.hour < END_HOUR)
    ].copy()

    if night_df.empty:
        raise ValueError(
            "00:00～06:00の心拍数データがありません。"
        )

    night_df["日付"] = (
        night_df[DATETIME_COLUMN]
        .dt.normalize()
    )

    return night_df


# ============================================================
# 移動中央値から基準心拍数を作成
# ============================================================

def calculate_baseline(
    work: pd.DataFrame,
) -> pd.DataFrame:
    """
    夜間心拍数を1分間隔に整え、
    移動中央値を基準心拍数として計算する。
    """

    work = work.copy()

    work = work.set_index(DATETIME_COLUMN)

    work = work[
        [HEART_RATE_COLUMN]
    ].resample("1min").mean()

    # Garminは2分間隔のため、短い間だけ補間する
    work[HEART_RATE_COLUMN] = (
        work[HEART_RATE_COLUMN]
        .interpolate(
            method="time",
            limit=INTERPOLATE_LIMIT_MIN,
            limit_direction="both",
        )
    )

    window_points = BASELINE_WINDOW_MIN + 1

    work["基準心拍数"] = (
        work[HEART_RATE_COLUMN]
        .rolling(
            window=window_points,
            center=True,
            min_periods=MIN_BASELINE_POINTS,
        )
        .median()
    )

    work["上昇幅"] = (
        work[HEART_RATE_COLUMN]
        - work["基準心拍数"]
    )

    work["上昇候補"] = (
        work["上昇幅"] >= RISE_THRESHOLD
    )

    return work


# ============================================================
# 連続候補をイベントにまとめる
# ============================================================

def build_raw_events(
    work: pd.DataFrame,
    target_date: pd.Timestamp,
) -> list[dict]:
    """
    上昇候補となった時間を連続イベントにまとめる。
    """

    candidate_times = work.index[
        work["上昇候補"].fillna(False)
    ]

    if len(candidate_times) == 0:
        return []

    groups: list[list[pd.Timestamp]] = []
    current_group = [candidate_times[0]]

    for current_time in candidate_times[1:]:
        previous_time = current_group[-1]

        gap_min = (
            current_time - previous_time
        ).total_seconds() / 60

        if gap_min <= MAX_CONTINUOUS_GAP_MIN:
            current_group.append(current_time)
        else:
            groups.append(current_group)
            current_group = [current_time]

    groups.append(current_group)

    raw_events: list[dict] = []

    for group in groups:
        start_time = group[0]
        end_time = group[-1]

        duration_min = (
            (end_time - start_time).total_seconds() / 60
        ) + 1

        if duration_min < MIN_EVENT_DURATION_MIN:
            continue

        if duration_min > MAX_EVENT_DURATION_MIN:
            continue

        event_data = work.loc[
            start_time:end_time
        ].copy()

        valid_hr = event_data[
            HEART_RATE_COLUMN
        ].dropna()

        valid_baseline = event_data[
            "基準心拍数"
        ].dropna()

        valid_rise = event_data[
            "上昇幅"
        ].dropna()

        if valid_hr.empty:
            continue

        if valid_baseline.empty:
            continue

        if valid_rise.empty:
            continue

        peak_hr = float(valid_hr.max())
        baseline_hr = float(valid_baseline.median())
        maximum_rise = float(valid_rise.max())

        peak_time = valid_hr.idxmax()

        raw_events.append(
            {
                "日付": target_date,
                "開始時刻": start_time,
                "最高心拍時刻": peak_time,
                "終了時刻": end_time,
                "継続時間_分": round(duration_min, 1),
                "最高心拍数": round(peak_hr, 1),
                "基準心拍数": round(baseline_hr, 1),
                "最大上昇幅": round(maximum_rise, 1),
            }
        )

    return raw_events


# ============================================================
# 近接イベントを1回にまとめる
# ============================================================

def merge_nearby_events(
    events: list[dict],
) -> list[dict]:
    """
    MERGE_GAP_MIN以内に続くイベントを同じ1回としてまとめる。
    """

    if not events:
        return []

    events = sorted(
        events,
        key=lambda x: x["開始時刻"],
    )

    merged_events = [events[0].copy()]

    for event in events[1:]:
        previous = merged_events[-1]

        gap_min = (
            event["開始時刻"]
            - previous["終了時刻"]
        ).total_seconds() / 60

        if gap_min <= MERGE_GAP_MIN:
            previous["終了時刻"] = max(
                previous["終了時刻"],
                event["終了時刻"],
            )

            previous["継続時間_分"] = round(
                (
                    previous["終了時刻"]
                    - previous["開始時刻"]
                ).total_seconds() / 60 + 1,
                1,
            )

            if (
                event["最高心拍数"]
                > previous["最高心拍数"]
            ):
                previous["最高心拍数"] = (
                    event["最高心拍数"]
                )

                previous["最高心拍時刻"] = (
                    event["最高心拍時刻"]
                )

            previous["最大上昇幅"] = max(
                previous["最大上昇幅"],
                event["最大上昇幅"],
            )

            previous["基準心拍数"] = round(
                np.nanmedian(
                    [
                        previous["基準心拍数"],
                        event["基準心拍数"],
                    ]
                ),
                1,
            )

        else:
            merged_events.append(event.copy())

    return merged_events


# ============================================================
# 1日分のイベント検出
# ============================================================

def detect_events_for_one_night(
    one_night_df: pd.DataFrame,
    target_date: pd.Timestamp,
) -> tuple[list[dict], pd.DataFrame]:
    """
    1日分の00:00～06:00からイベントを検出する。
    """

    if one_night_df.empty:
        return [], pd.DataFrame()

    work = one_night_df[
        [
            DATETIME_COLUMN,
            HEART_RATE_COLUMN,
        ]
    ].copy()

    work = calculate_baseline(work)

    raw_events = build_raw_events(
        work,
        target_date,
    )

    merged_events = merge_nearby_events(
        raw_events
    )

    return merged_events, work


# ============================================================
# 全日付を解析
# ============================================================

ANALYSIS_START_DATE = pd.Timestamp("2026-05-23")

def analyze_all_nights(
    heart_rate_df: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    日付ごとの夜間心拍イベント回数を計算する。
    """

    night_df = extract_night_data(
        heart_rate_df
    )

    # アブレーション翌日以降だけを解析する
    night_df = night_df[
        night_df["日付"] >= ANALYSIS_START_DATE
    ].copy()

    if night_df.empty:
        raise ValueError(
            f"{ANALYSIS_START_DATE:%Y-%m-%d}以降の"
            "夜間心拍数データがありません。"
        )

    summary_rows: list[dict] = []
    detail_rows: list[dict] = []
    analysis_rows: list[pd.DataFrame] = []

    target_dates = sorted(
        night_df["日付"].dropna().unique()
    )

    for target_date_value in target_dates:
        target_date = pd.Timestamp(
            target_date_value
        ).normalize()

        one_night_df = night_df[
            night_df["日付"] == target_date
        ].copy()

        events, work = detect_events_for_one_night(
            one_night_df,
            target_date,
        )

        valid_hr = one_night_df[
            HEART_RATE_COLUMN
        ].dropna()

        if valid_hr.empty:
            minimum_hr = np.nan
            median_hr = np.nan
            maximum_hr = np.nan
        else:
            minimum_hr = round(
                float(valid_hr.min()),
                1,
            )
            median_hr = round(
                float(valid_hr.median()),
                1,
            )
            maximum_hr = round(
                float(valid_hr.max()),
                1,
            )

        summary_rows.append(
            {
                "日付": target_date,
                "夜間心拍上昇回数": len(events),
                "元データ件数": len(one_night_df),
                "夜間最低心拍数": minimum_hr,
                "夜間中央値心拍数": median_hr,
                "夜間最高心拍数": maximum_hr,
            }
        )

        for event_number, event in enumerate(
            events,
            start=1,
        ):
            event_row = event.copy()
            event_row["イベント番号"] = event_number
            detail_rows.append(event_row)

        if not work.empty:
            work_export = work.reset_index()

            work_export["日付"] = target_date

            work_export["イベント候補"] = (
                work_export["上昇候補"]
                .fillna(False)
                .astype(int)
            )

            analysis_rows.append(
                work_export[
                    [
                        "日付",
                        DATETIME_COLUMN,
                        HEART_RATE_COLUMN,
                        "基準心拍数",
                        "上昇幅",
                        "イベント候補",
                    ]
                ]
            )

        print(
            f"{target_date:%Y-%m-%d}: "
            f"{len(events)}回"
        )

    summary_df = pd.DataFrame(summary_rows)

    if detail_rows:
        detail_df = pd.DataFrame(detail_rows)

        detail_df = detail_df[
            [
                "日付",
                "イベント番号",
                "開始時刻",
                "最高心拍時刻",
                "終了時刻",
                "継続時間_分",
                "最高心拍数",
                "基準心拍数",
                "最大上昇幅",
            ]
        ]
    else:
        detail_df = pd.DataFrame(
            columns=[
                "日付",
                "イベント番号",
                "開始時刻",
                "最高心拍時刻",
                "終了時刻",
                "継続時間_分",
                "最高心拍数",
                "基準心拍数",
                "最大上昇幅",
            ]
        )

    if analysis_rows:
        analysis_df = pd.concat(
            analysis_rows,
            ignore_index=True,
        )
    else:
        analysis_df = pd.DataFrame(
            columns=[
                "日付",
                DATETIME_COLUMN,
                HEART_RATE_COLUMN,
                "基準心拍数",
                "上昇幅",
                "イベント候補",
            ]
        )

    return summary_df, detail_df, analysis_df


# ============================================================
# 判定条件シート
# ============================================================

def create_conditions_df() -> pd.DataFrame:
    """
    使用した判定条件をExcel保存用の表にする。
    """

    return pd.DataFrame(
        {
            "項目": [
                "対象時間",
                "入力ファイル",
                "読込対象シート",
                "基準心拍数",
                "上昇閾値",
                "最短継続時間",
                "最長継続時間",
                "連続とみなす最大間隔",
                "同一イベントにまとめる間隔",
                "短時間欠測の補間",
                "注意",
            ],
            "設定値": [
                "00:00以上、06:00未満",
                FILE_PATTERN,
                "ファイル名の日付と同名のシート",
                f"前後約{BASELINE_WINDOW_MIN}分の移動中央値",
                f"基準心拍数より{RISE_THRESHOLD}拍/分以上",
                f"{MIN_EVENT_DURATION_MIN}分",
                f"{MAX_EVENT_DURATION_MIN}分",
                f"{MAX_CONTINUOUS_GAP_MIN}分以内",
                f"{MERGE_GAP_MIN}分以内",
                f"最大{INTERPOLATE_LIMIT_MIN}分",
                (
                    "結果はトイレ回数そのものではなく、"
                    "トイレに起きた可能性のある"
                    "夜間心拍上昇イベント回数"
                ),
            ],
        }
    )


# ============================================================
# 既存の目視入力シートを保存
# ============================================================

def read_existing_manual_sheet(
    output_path: Path,
) -> pd.DataFrame:
    """
    既存のtoilet_count_data.xlsxに
    「トイレ回数目視」シートがあれば読み込む。
    """

    default_df = pd.DataFrame(
        columns=[
            "日付",
            "目視トイレ回数",
            "メモ",
        ]
    )

    if not output_path.exists():
        return default_df

    try:
        excel_file = pd.ExcelFile(
            output_path,
            engine="openpyxl",
        )

        if "トイレ回数目視" not in excel_file.sheet_names:
            return default_df

        manual_df = pd.read_excel(
            output_path,
            sheet_name="トイレ回数目視",
            engine="openpyxl",
        )

        return manual_df

    except Exception as exc:
        print(
            "既存の「トイレ回数目視」シートを"
            f"読み込めませんでした: {exc}"
        )

        return default_df


# ============================================================
# Excelの書式調整
# ============================================================

def format_excel_file(
    output_path: Path,
) -> None:
    """
    出力Excelの列幅や表示形式を調整する。
    """

    workbook = load_workbook(output_path)

    for worksheet in workbook.worksheets:
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions

        for cell in worksheet[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
            )

        for column_cells in worksheet.columns:
            max_length = 0
            column_index = column_cells[0].column

            for cell in column_cells:
                value = cell.value

                if value is None:
                    continue

                text = str(value)
                max_length = max(
                    max_length,
                    len(text),
                )

            adjusted_width = min(
                max(max_length + 2, 10),
                35,
            )

            worksheet.column_dimensions[
                get_column_letter(column_index)
            ].width = adjusted_width

    # 日別回数
    if "日別回数" in workbook.sheetnames:
        worksheet = workbook["日別回数"]

        for cell in worksheet["A"][1:]:
            cell.number_format = "yyyy-mm-dd"

    # イベント詳細
    if "イベント詳細" in workbook.sheetnames:
        worksheet = workbook["イベント詳細"]

        for cell in worksheet["A"][1:]:
            cell.number_format = "yyyy-mm-dd"

        for column_letter in ["C", "D", "E"]:
            for cell in worksheet[column_letter][1:]:
                cell.number_format = (
                    "yyyy-mm-dd hh:mm"
                )

    # トイレ回数目視
    if "トイレ回数目視" in workbook.sheetnames:
        worksheet = workbook["トイレ回数目視"]

        for cell in worksheet["A"][1:]:
            cell.number_format = "yyyy-mm-dd"

    # 読込ファイル一覧
    if "読込ファイル一覧" in workbook.sheetnames:
        worksheet = workbook["読込ファイル一覧"]

        for cell in worksheet["B"][1:]:
            cell.number_format = "yyyy-mm-dd"

    # 心拍判定データ
    if "心拍判定データ" in workbook.sheetnames:
        worksheet = workbook["心拍判定データ"]

        for cell in worksheet["A"][1:]:
            cell.number_format = "yyyy-mm-dd"

        for cell in worksheet["B"][1:]:
            cell.number_format = (
                "yyyy-mm-dd hh:mm"
            )

    workbook.save(output_path)


# ============================================================
# Excel保存
# ============================================================

def save_results(
    summary_df: pd.DataFrame,
    detail_df: pd.DataFrame,
    analysis_df: pd.DataFrame,
    file_log_df: pd.DataFrame,
    output_path: Path,
) -> None:
    """
    結果をtoilet_count_data.xlsxへ保存する。
    """

    output_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    manual_df = read_existing_manual_sheet(
        output_path
    )

    conditions_df = create_conditions_df()

    with pd.ExcelWriter(
        output_path,
        engine="openpyxl",
        mode="w",
        datetime_format="yyyy-mm-dd hh:mm",
        date_format="yyyy-mm-dd",
    ) as writer:

        summary_df.to_excel(
            writer,
            sheet_name="日別回数",
            index=False,
        )

        detail_df.to_excel(
            writer,
            sheet_name="イベント詳細",
            index=False,
        )

        manual_df.to_excel(
            writer,
            sheet_name="トイレ回数目視",
            index=False,
        )

        conditions_df.to_excel(
            writer,
            sheet_name="判定条件",
            index=False,
        )

        file_log_df.to_excel(
            writer,
            sheet_name="読込ファイル一覧",
            index=False,
        )

        analysis_df.to_excel(
            writer,
            sheet_name="心拍判定データ",
            index=False,
        )

    format_excel_file(output_path)

    print("")
    print("Excelを保存しました。")
    print(output_path)


# ============================================================
# メイン処理
# ============================================================

def main() -> None:
    print("=" * 60)
    print("夜間心拍上昇イベントを集計します。")
    print("=" * 60)

    print(f"入力フォルダー: {INPUT_DIR}")
    print(f"出力ファイル:   {OUTPUT_PATH}")
    print("")

    heart_rate_df, file_log_df = (
        load_all_daily_files(INPUT_DIR)
    )

    print("")
    print(
        f"読み込み後の総データ件数: "
        f"{len(heart_rate_df):,}"
    )

    summary_df, detail_df, analysis_df = (
        analyze_all_nights(heart_rate_df)
    )

    total_events = int(
        summary_df["夜間心拍上昇回数"].sum()
    )

    print("")
    print(f"解析対象日数: {len(summary_df):,}")
    print(f"検出イベント合計: {total_events:,}")

    save_results(
        summary_df=summary_df,
        detail_df=detail_df,
        analysis_df=analysis_df,
        file_log_df=file_log_df,
        output_path=OUTPUT_PATH,
    )

    try:
        os.startfile(OUTPUT_PATH)
    except Exception as exc:
        print(
            "Excelファイルを自動で開けませんでした。"
        )
        print(exc)


if __name__ == "__main__":
    main()