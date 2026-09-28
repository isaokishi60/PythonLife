# -*- coding: utf-8 -*-
"""
next_crop_candidates.py

目的
----
「作物マスター.xlsx」「連作障害.xlsx」「作付け履歴」を使って、
指定日時点で
  1) 今の時期に開始できる作物
  2) 連作障害の回避年数を満たす畝
の組合せを候補として抽出する。

重要ルール
----------
・候補作物は「作物マスター.xlsx」の作物名から選ぶ。
・連作障害の年数は、各作物の「撤収日」から起算する。
・連作判定は「同じ作物」ではなく「同じ科」の履歴で判定する。
・年数=0 は連作制限なし。
・将来のジャガイモ、スイカ、タマネギ等の「畝予約との衝突判定」は
  次段階で追加する。本版では未実装。

実行例
------
C:\\Users\\spax2\\venv311\\Scripts\\python.exe ^
"C:\\Users\\spax2\\GitHub\\PythonLife\\Garden\\app\\next_crop_candidates.py"

日付指定:
C:\\Users\\spax2\\venv311\\Scripts\\python.exe ^
"C:\\Users\\spax2\\GitHub\\PythonLife\\Garden\\app\\next_crop_candidates.py" ^
--as-of-date 2026-09-11

履歴ファイルを明示:
... --history-path "C:\\...\\vegetable_garden_location.xlsx"

出力:
同じ「農作業関係Excel」フォルダーに
「次作候補.xlsx」を作成する。
"""

from __future__ import annotations

import argparse
import calendar
import os
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Iterable, Optional

import pandas as pd

# HISTORY_CROP_COLUMNS
# ============================================================
# 基本設定
# ============================================================

DEFAULT_EXCEL_DIR = Path(
    os.environ.get("OneDrive", r"C:\Users\spax2\OneDrive")
) / "ドキュメント" / "PythonWork" / "農作業" / "農作業関係Excel"

DEFAULT_CROP_MASTER = DEFAULT_EXCEL_DIR / "作物マスター.xlsx"
DEFAULT_ROTATION_MASTER = DEFAULT_EXCEL_DIR / "連作障害.xlsx"

# 現在の作付け履歴として使えそうなファイル名。
# 実ファイルが異なる場合は --history-path で指定する。
DEFAULT_HISTORY_CANDIDATES = [
    DEFAULT_EXCEL_DIR / "vegetable_garden_location.xlsx",
    DEFAULT_EXCEL_DIR / "vegetable_garden_photo_ex2_bed.xlsx",
]

DEFAULT_OUTPUT = DEFAULT_EXCEL_DIR / "次作候補.xlsx"


# ============================================================
# 列名候補
# ============================================================

CROP_NAME_COLUMNS = ["作物名", "Name", "name", "作物", "野菜名", "Item", "item"]
FAMILY_COLUMNS = ["科（連作用）", "科", "family", "Family"]
ROTATION_YEARS_COLUMNS = ["年数", "連作回避年数", "回避年数"]

START_WINDOW_COLUMNS = ["播種開始月", "開始月"]
END_WINDOW_COLUMNS = ["播種終了月", "終了月"]
CROP_TYPE_COLUMNS = ["作型", "種別"]
MIN_MONTHS_COLUMNS = ["最低栽培月数", "栽培月数"]

# 履歴
BED_COLUMNS = ["畝番号", "畝", "Bed", "bed", "畝No", "畝NO", "畝No."]
BED_AB_COLUMNS = ["畝AB", "畝A/B", "AB"]
SECTION_COLUMNS = ["区画", "区画番号", "Section", "section"]
HISTORY_CROP_COLUMNS = [
    "Name",
    "Name or Item",
    "作物名",
    "Item",
    "作物",
    "野菜名",
]
REMOVAL_DATE_COLUMNS = [
    "撤収日", "撤去日", "有効終了日", "終了日", "収穫終了日",
    "EndDate", "end_date", "Date_end"
]
START_DATE_COLUMNS = [
    "開始日", "有効開始日", "播種日", "定植日",
    "StartDate", "start_date", "Date_start"
]


# ============================================================
# ユーティリティ
# ============================================================

def normalize_text(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip().replace("　", " ")


def find_column(df: pd.DataFrame, candidates: Iterable[str], required: bool = False) -> Optional[str]:
    """候補名から実在列を探す。前後空白も吸収する。"""
    normalized = {str(c).strip(): c for c in df.columns}
    for name in candidates:
        if name in normalized:
            return normalized[name]

    # 英字だけ大文字小文字を無視
    lower_map = {str(c).strip().lower(): c for c in df.columns}
    for name in candidates:
        if name.lower() in lower_map:
            return lower_map[name.lower()]

    if required:
        raise KeyError(
            f"必要な列が見つかりません。候補={list(candidates)}\n"
            f"実際の列={list(df.columns)}"
        )
    return None


def read_first_nonempty_sheet(path: Path) -> tuple[pd.DataFrame, str]:
    """
    Excel内の各シートを見て、データのある最初のシートを返す。
    """
    xls = pd.ExcelFile(path)
    for sheet in xls.sheet_names:
        df = pd.read_excel(path, sheet_name=sheet)
        if not df.empty:
            return df, sheet
    raise ValueError(f"データのあるシートがありません: {path}")


def read_sheet_matching_columns(
    path: Path,
    required_groups: list[list[str]],
) -> tuple[pd.DataFrame, str]:
    """
    必要な列群を最も多く含むシートを自動選択する。
    required_groups:
      例 [[作物名候補], [撤収日候補]]
    """
    xls = pd.ExcelFile(path)
    best = None
    best_score = -1

    for sheet in xls.sheet_names:
        df = pd.read_excel(path, sheet_name=sheet)
        if df.empty:
            continue
        score = 0
        for group in required_groups:
            if find_column(df, group) is not None:
                score += 1
        if score > best_score:
            best = (df, sheet)
            best_score = score

    if best is None:
        raise ValueError(f"データのあるシートがありません: {path}")
    return best


def parse_as_of_date(text: Optional[str]) -> date:
    if not text:
        return date.today()
    return datetime.strptime(text, "%Y-%m-%d").date()


def month_decimal(d: date) -> float:
    """
    Excelマスターの 9.0, 9.5, 9.9 のような値と比較するため、
    日付を「月 + 月内進捗」に変換する。

    例:
      9月1日  ≒ 9.00
      9月中旬 ≒ 9.47
      9月末   ≒ 9.97
    """
    days = calendar.monthrange(d.year, d.month)[1]
    return d.month + (d.day - 1) / days


def is_in_month_window(value: float, start, end) -> bool:
    """現在値が開始～終了に入っているか。年跨ぎにも対応。"""
    if pd.isna(start) or pd.isna(end):
        return False
    start = float(start)
    end = float(end)
    if start <= end:
        return start <= value <= end
    # 例: 11.0 ～ 2.9
    return value >= start or value <= end


def add_years_safe(d: date, years: int) -> date:
    """2/29を考慮して年を加算。"""
    try:
        return d.replace(year=d.year + years)
    except ValueError:
        # 2/29 -> 2/28
        return d.replace(month=2, day=28, year=d.year + years)


def add_months_approx(d: date, months: float) -> date:
    """
    最低栽培月数から想定終了日を概算。
    1か月=30.4375日として扱う。
    """
    days = round(float(months) * 30.4375)
    return d + pd.Timedelta(days=days)


def choose_history_path(explicit: Optional[str]) -> Path:
    if explicit:
        p = Path(explicit)
        if not p.exists():
            raise FileNotFoundError(f"履歴ファイルがありません: {p}")
        return p

    for p in DEFAULT_HISTORY_CANDIDATES:
        if p.exists():
            return p

    raise FileNotFoundError(
        "作付け履歴ファイルを自動検出できませんでした。\n"
        "--history-path で指定してください。\n候補:\n"
        + "\n".join(str(p) for p in DEFAULT_HISTORY_CANDIDATES)
    )


def make_bed_key(row: pd.Series, bed_col, bed_ab_col, section_col) -> str:
    parts = []
    if bed_col:
        v = normalize_text(row.get(bed_col))
        if v:
            parts.append(v)
    if bed_ab_col:
        v = normalize_text(row.get(bed_ab_col))
        if v:
            parts.append(v)
    if section_col:
        v = normalize_text(row.get(section_col))
        if v:
            parts.append(v)
    return "-".join(parts)


# ============================================================
# データ読み込み
# ============================================================

@dataclass
class CropMasterColumns:
    crop: str
    family: str
    start_window: str
    end_window: str
    crop_type: Optional[str]
    min_months: Optional[str]


def load_crop_master(path: Path) -> tuple[pd.DataFrame, CropMasterColumns, str]:
    df, sheet = read_sheet_matching_columns(
        path,
        [
            CROP_NAME_COLUMNS,
            FAMILY_COLUMNS,
            START_WINDOW_COLUMNS,
            END_WINDOW_COLUMNS,
        ],
    )

    cols = CropMasterColumns(
        crop=find_column(df, CROP_NAME_COLUMNS, required=True),
        family=find_column(df, FAMILY_COLUMNS, required=True),
        start_window=find_column(df, START_WINDOW_COLUMNS, required=True),
        end_window=find_column(df, END_WINDOW_COLUMNS, required=True),
        crop_type=find_column(df, CROP_TYPE_COLUMNS),
        min_months=find_column(df, MIN_MONTHS_COLUMNS),
    )
    return df, cols, sheet


def load_rotation_master(path: Path) -> tuple[pd.DataFrame, str, str, str, str]:
    df, sheet = read_sheet_matching_columns(
        path,
        [CROP_NAME_COLUMNS, FAMILY_COLUMNS, ROTATION_YEARS_COLUMNS],
    )
    crop_col = find_column(df, CROP_NAME_COLUMNS, required=True)
    family_col = find_column(df, FAMILY_COLUMNS, required=True)
    years_col = find_column(df, ROTATION_YEARS_COLUMNS, required=True)
    risk_col = find_column(df, ["リスク"])
    return df, crop_col, family_col, years_col, risk_col


def load_history(path: Path):
    df, sheet = read_sheet_matching_columns(
        path,
        [HISTORY_CROP_COLUMNS, REMOVAL_DATE_COLUMNS],
    )
    crop_col = find_column(df, HISTORY_CROP_COLUMNS, required=True)
    removal_col = find_column(df, REMOVAL_DATE_COLUMNS, required=True)
    start_col = find_column(df, START_DATE_COLUMNS)

    bed_col = find_column(df, BED_COLUMNS)
    bed_ab_col = find_column(df, BED_AB_COLUMNS)
    section_col = find_column(df, SECTION_COLUMNS)

    if not any([bed_col, bed_ab_col, section_col]):
        raise KeyError(
            "畝を識別できる列がありません。\n"
            f"候補: {BED_COLUMNS + BED_AB_COLUMNS + SECTION_COLUMNS}\n"
            f"実際の列: {list(df.columns)}"
        )

    df = df.copy()
    df["_作物"] = df[crop_col].map(normalize_text)
    df["_撤収日"] = pd.to_datetime(df[removal_col], errors="coerce")
    if start_col:
        df["_開始日"] = pd.to_datetime(df[start_col], errors="coerce")
    else:
        df["_開始日"] = pd.NaT

    df["_畝キー"] = df.apply(
        lambda r: make_bed_key(r, bed_col, bed_ab_col, section_col), axis=1
    )
    df = df[df["_畝キー"] != ""].copy()

    return df, sheet


# ============================================================
# 判定ロジック
# ============================================================

def build_crop_family_map(
    crop_master: pd.DataFrame,
    crop_cols: CropMasterColumns,
    rotation_master: pd.DataFrame,
    rot_crop_col: str,
    rot_family_col: str,
) -> dict[str, str]:
    mapping: dict[str, str] = {}

    for _, r in crop_master.iterrows():
        crop = normalize_text(r[crop_cols.crop])
        family = normalize_text(r[crop_cols.family])
        if crop and family:
            mapping[crop] = family

    for _, r in rotation_master.iterrows():
        crop = normalize_text(r[rot_crop_col])
        family = normalize_text(r[rot_family_col])
        if crop and family:
            mapping[crop] = family

    return mapping


def build_rotation_year_map(
    rotation_master: pd.DataFrame,
    crop_col: str,
    years_col: str,
) -> dict[str, int]:
    result = {}
    for _, r in rotation_master.iterrows():
        crop = normalize_text(r[crop_col])
        if not crop:
            continue
        try:
            years = int(float(r[years_col])) if not pd.isna(r[years_col]) else 0
        except (TypeError, ValueError):
            years = 0
        result[crop] = years
    return result


def current_season_candidates(
    crop_master: pd.DataFrame,
    cols: CropMasterColumns,
    as_of: date,
) -> pd.DataFrame:
    now_value = month_decimal(as_of)
    rows = []

    for _, r in crop_master.iterrows():
        crop = normalize_text(r[cols.crop])
        family = normalize_text(r[cols.family])
        if not crop:
            continue

        start = r[cols.start_window]
        end = r[cols.end_window]

        if not is_in_month_window(now_value, start, end):
            continue

        crop_type = normalize_text(r[cols.crop_type]) if cols.crop_type else ""
        min_months = r[cols.min_months] if cols.min_months else None

        if crop_type.startswith("苗"):
            work_type = "苗・定植"
        else:
            work_type = "播種"

        expected_end = pd.NaT
        if min_months is not None and not pd.isna(min_months):
            expected_end = add_months_approx(as_of, float(min_months))

        rows.append(
            {
                "作物名": crop,
                "作型": crop_type,
                "科": family,
                "作業": work_type,
                "適期開始": start,
                "適期終了": end,
                "最低栽培月数": min_months,
                "今回開始日": pd.Timestamp(as_of),
                "想定最低撤収日": expected_end,
            }
        )

    return pd.DataFrame(rows)


def latest_family_removal_by_bed(
    history: pd.DataFrame,
    crop_family_map: dict[str, str],
) -> pd.DataFrame:
    df = history.copy()
    df["_科"] = df["_作物"].map(crop_family_map).fillna("")

    # 撤収日があるものだけが「過去の完了作」として連作起算に使える
    done = df[df["_撤収日"].notna() & (df["_科"] != "")].copy()

    if done.empty:
        return pd.DataFrame(columns=["_畝キー", "_科", "_最終撤収日", "_前作"])

    idx = done.groupby(["_畝キー", "_科"])["_撤収日"].idxmax()
    latest = done.loc[idx, ["_畝キー", "_科", "_撤収日", "_作物"]].copy()
    latest = latest.rename(
        columns={"_撤収日": "_最終撤収日", "_作物": "_前作"}
    )
    return latest


def determine_beds(history: pd.DataFrame, as_of: date) -> pd.DataFrame:
    """
    履歴に登場する畝を列挙。
    現在作付中かどうかも判定する。

    現在作付中:
      開始日 <= 基準日 かつ
      撤収日が空欄、または 撤収日 >= 基準日
    """
    base_date = pd.Timestamp(as_of)
    rows = []

    for bed, g in history.groupby("_畝キー"):
        active = g[
            (
                g["_開始日"].isna()
                | (g["_開始日"] <= base_date)
            )
            & (
                g["_撤収日"].isna()
                | (g["_撤収日"] >= base_date)
            )
        ].copy()

        # 開始日が全くない古い形式では、撤収日空欄を現在作付中とみなす
        if g["_開始日"].isna().all():
            active = g[g["_撤収日"].isna()].copy()

        active_crops = "・".join(
            sorted({x for x in active["_作物"].tolist() if x})
        )

        rows.append(
            {
                "畝": bed,
                "現在作付中": bool(active_crops),
                "現在作物": active_crops,
            }
        )

    return pd.DataFrame(rows)


def evaluate_candidates(
    season_crops: pd.DataFrame,
    beds: pd.DataFrame,
    family_history: pd.DataFrame,
    rotation_years: dict[str, int],
    as_of: date,
) -> pd.DataFrame:
    rows = []
    base_date = pd.Timestamp(as_of)

    # 検索しやすい辞書
    hist_lookup = {}
    for _, r in family_history.iterrows():
        hist_lookup[(r["_畝キー"], r["_科"])] = (
            r["_最終撤収日"],
            r["_前作"],
        )

    for _, crop in season_crops.iterrows():
        crop_name = crop["作物名"]
        family = crop["科"]
        years = int(rotation_years.get(crop_name, 0))

        for _, bed in beds.iterrows():
            bed_name = bed["畝"]

            # 現在使用中の畝は候補から除外するが、理由を残す
            if bed["現在作付中"]:
                rows.append(
                    {
                        **crop.to_dict(),
                        "畝": bed_name,
                        "連作回避年数": years,
                        "同科の直近作物": "",
                        "同科の最終撤収日": pd.NaT,
                        "連作解禁日": pd.NaT,
                        "判定": "不可",
                        "理由": f"現在作付中: {bed['現在作物']}",
                    }
                )
                continue

            latest = hist_lookup.get((bed_name, family))

            if years <= 0 or latest is None:
                rows.append(
                    {
                        **crop.to_dict(),
                        "畝": bed_name,
                        "連作回避年数": years,
                        "同科の直近作物": latest[1] if latest else "",
                        "同科の最終撤収日": latest[0] if latest else pd.NaT,
                        "連作解禁日": latest[0] if latest and years > 0 else pd.NaT,
                        "判定": "候補",
                        "理由": "連作制限なし" if years <= 0 else "同じ科の過去履歴なし",
                    }
                )
                continue

            last_removal_ts, last_crop = latest
            last_removal = pd.Timestamp(last_removal_ts).date()
            available_from = add_years_safe(last_removal, years)

            if as_of >= available_from:
                judgment = "候補"
                reason = "連作回避期間を経過"
            else:
                judgment = "不可"
                reason = f"{family}の連作回避期間中"

            rows.append(
                {
                    **crop.to_dict(),
                    "畝": bed_name,
                    "連作回避年数": years,
                    "同科の直近作物": last_crop,
                    "同科の最終撤収日": pd.Timestamp(last_removal),
                    "連作解禁日": pd.Timestamp(available_from),
                    "判定": judgment,
                    "理由": reason,
                }
            )

    result = pd.DataFrame(rows)
    if not result.empty:
        result = result.sort_values(
            ["判定", "作物名", "畝"],
            key=lambda s: s.map({"候補": 0, "不可": 1}).fillna(s)
            if s.name == "判定"
            else s,
        ).reset_index(drop=True)
    return result


# ============================================================
# 出力
# ============================================================

def write_output(
    output_path: Path,
    season_crops: pd.DataFrame,
    beds: pd.DataFrame,
    result: pd.DataFrame,
):
    output_path.parent.mkdir(parents=True, exist_ok=True)

    candidates = result[result["判定"] == "候補"].copy() if not result.empty else result
    rejected = result[result["判定"] == "不可"].copy() if not result.empty else result

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        candidates.to_excel(writer, sheet_name="候補", index=False)
        rejected.to_excel(writer, sheet_name="除外理由", index=False)
        season_crops.to_excel(writer, sheet_name="今の時期の作物", index=False)
        beds.to_excel(writer, sheet_name="畝状態", index=False)

        # 見やすさを最低限整える
        wb = writer.book
        for ws in wb.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for col_cells in ws.columns:
                max_len = 0
                col_letter = col_cells[0].column_letter
                for cell in col_cells[:200]:
                    value = "" if cell.value is None else str(cell.value)
                    max_len = max(max_len, len(value))
                ws.column_dimensions[col_letter].width = min(max(max_len + 2, 10), 35)


# ============================================================
# main
# ============================================================

def main():
    parser = argparse.ArgumentParser(description="次に育てる作物と畝の候補を抽出")
    parser.add_argument("--as-of-date", help="基準日 YYYY-MM-DD。省略時は今日")
    parser.add_argument("--crop-master", default=str(DEFAULT_CROP_MASTER))
    parser.add_argument("--rotation-master", default=str(DEFAULT_ROTATION_MASTER))
    parser.add_argument("--history-path", help="作付け履歴Excel")
    parser.add_argument("--output", default=str(DEFAULT_OUTPUT))
    args = parser.parse_args()

    as_of = parse_as_of_date(args.as_of_date)
    crop_master_path = Path(args.crop_master)
    rotation_master_path = Path(args.rotation_master)
    history_path = choose_history_path(args.history_path)
    output_path = Path(args.output)

    print("=== 次作候補判定 ===")
    print(f"基準日          : {as_of}")
    print(f"作物マスター    : {crop_master_path}")
    print(f"連作障害マスター: {rotation_master_path}")
    print(f"作付け履歴      : {history_path}")

    for p in [crop_master_path, rotation_master_path, history_path]:
        if not p.exists():
            raise FileNotFoundError(p)

    crop_master, crop_cols, crop_sheet = load_crop_master(crop_master_path)
    rot_master, rot_crop_col, rot_family_col, rot_years_col, _ = load_rotation_master(
        rotation_master_path
    )
    history, history_sheet = load_history(history_path)

    print(f"作物マスター sheet: {crop_sheet}")
    print(f"連作障害 sheet      : {rot_master.attrs.get('sheet', '自動選択')}")
    print(f"履歴 sheet          : {history_sheet}")

    crop_family_map = build_crop_family_map(
        crop_master,
        crop_cols,
        rot_master,
        rot_crop_col,
        rot_family_col,
    )
    rotation_years = build_rotation_year_map(
        rot_master,
        rot_crop_col,
        rot_years_col,
    )

    # 履歴にある作物で「科」が引けないものを通知
    unknown_crops = sorted(
        {
            c
            for c in history["_作物"].dropna().astype(str)
            if normalize_text(c) and normalize_text(c) not in crop_family_map
        }
    )
    if unknown_crops:
        print("\n[注意] 科を特定できない履歴作物:")
        for c in unknown_crops:
            print(f"  - {c}")

    season_crops = current_season_candidates(crop_master, crop_cols, as_of)
    family_history = latest_family_removal_by_bed(history, crop_family_map)
    beds = determine_beds(history, as_of)

    result = evaluate_candidates(
        season_crops,
        beds,
        family_history,
        rotation_years,
        as_of,
    )

    write_output(output_path, season_crops, beds, result)

    candidates = result[result["判定"] == "候補"].copy() if not result.empty else result

    print()
    print(f"今の時期に開始可能な作型数: {len(season_crops)}")
    print(f"履歴から認識した畝数      : {len(beds)}")
    print(f"作物×畝の候補数           : {len(candidates)}")
    print(f"出力                       : {output_path}")

    if not candidates.empty:
        print("\n=== 候補（先頭30件） ===")
        show_cols = [
            "作物名", "作型", "作業", "科", "畝",
            "連作回避年数", "同科の直近作物",
            "同科の最終撤収日", "連作解禁日",
            "想定最低撤収日", "理由"
        ]
        show_cols = [c for c in show_cols if c in candidates.columns]
        print(candidates[show_cols].head(30).to_string(index=False))
    else:
        print("\n候補が0件です。")
        print("「除外理由」シートを確認してください。")


if __name__ == "__main__":
    main()
