# -*- coding: utf-8 -*-

"""
garden_current_status.py

vegetable_garden_location.xlsx から、
指定日時点の各畝の最新 Name or Item を取得し、
「現在の作付け.xlsx」に畑の配置形式で出力する。

判定ルール：
    ・「畝」をキーにする
    ・Date <= 指定日 のデータだけを対象にする
    ・各畝について指定日に最も近い最新日付の行を取得する
    ・その行の「Name or Item」を表示する

実行例：
    python garden_current_status.py --date 2026-09-03

--date を省略した場合は、画面から日付を入力する。
"""

import os
import argparse
from pathlib import Path
from datetime import datetime

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import (
    Font,
    PatternFill,
    Alignment,
    Border,
    Side,
)
from openpyxl.utils import get_column_letter


# ============================================================
# パス設定
# ============================================================

ONEDRIVE = os.environ.get("OneDrive")

if not ONEDRIVE:
    raise RuntimeError(
        "環境変数 OneDrive が見つかりません。"
    )

EXCEL_DIR = (
    Path(ONEDRIVE)
    / "ドキュメント"
    / "PythonWork"
    / "農作業"
    / "農作業関係Excel"
)

SOURCE_FILE = (
    EXCEL_DIR
    / "vegetable_garden_location.xlsx"
)

CROP_MASTER_FILE = (
    EXCEL_DIR
    / "作物マスター.xlsx"
)

OUTPUT_FILE = (
    EXCEL_DIR
    / "現在の作付け.xlsx"
)


# ============================================================
# 畑のレイアウト
# ============================================================

# 左側A畝と右側B畝の物理配置
#
# A08 の次に通路があり、
# その右側には B09 が存在する。
#
# A側はA16まで。A側には通路があるため、B側より1畝少ない。

GARDEN_LAYOUT = [
    ("A01", "B01"),
    ("A02", "B02"),
    ("A03", "B03"),
    ("A04", "B04"),
    ("A05", "B05"),
    ("A06", "B06"),
    ("A07", "B07"),
    ("A08", "B08"),
    ("通路", "B09"),
    ("A09", "B10"),
    ("A10", "B11"),
    ("A11", "B12"),
    ("A12", "B13"),
    ("A13", "B14"),
    ("A14", "B15"),
    ("A15", "B16"),
    ("A16", "B17"),
]


# ============================================================
# Excelデータ読み込み
# ============================================================

def load_garden_data():
    """
    vegetable_garden_location.xlsx を読み込む。
    """

    if not SOURCE_FILE.exists():
        raise FileNotFoundError(
            f"入力ファイルが見つかりません。\n"
            f"{SOURCE_FILE}"
        )

    df = pd.read_excel(SOURCE_FILE)

    # 必須列
    required_columns = [
        "Date",
        "Name or Item",
        "畝",
    ]

    missing_columns = [
        col
        for col in required_columns
        if col not in df.columns
    ]

    if missing_columns:
        raise ValueError(
            "必要な列がありません。\n"
            f"不足列: {missing_columns}\n"
            f"現在の列: {list(df.columns)}"
        )

    # Dateを日付型へ変換
    df["Date"] = pd.to_datetime(
        df["Date"],
        errors="coerce"
    )

    # 日付が読み取れない行は除外
    df = df.dropna(subset=["Date"]).copy()

    # 畝番号を文字列へ統一
    df["畝"] = (
        df["畝"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )

    # Name or Itemも文字列へ
    df["Name or Item"] = (
        df["Name or Item"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    return df

# ============================================================
# 作物名 → 科名 辞書
# ============================================================

def load_crop_family_dict():

    df_master = pd.read_excel(
        CROP_MASTER_FILE
    )

    df_master["作物名"] = (
        df_master["作物名"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    df_master["科（連作用）"] = (
        df_master["科（連作用）"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    return (
        df_master
        .drop_duplicates(subset=["作物名"])
        .set_index("作物名")["科（連作用）"]
        .to_dict()
    )

# ============================================================
# 指定日時点の各畝の最新データ取得
# ============================================================

def get_latest_item_by_bed(
    df,
    target_date,
    crop_family_dict
):
    """
    指定日以前について、各畝の最新日付を求め、
    その最新日付に複数レコードがあればまとめて返す。

    例:
        A14:
            ラッカセイ
            ツルムラサキ

        A03:
            スイカ　撤収

        B14:
            ジャガイモ　撤収完了
    """

    target_date = pd.to_datetime(target_date)

    # 指定日以前、かつ畝番号あり
    data = df[
        (df["Date"] <= target_date)
        & (df["畝"] != "")
    ].copy()

    if data.empty:
        return {}

    # Tag1 がなければ空欄を作る
    if "Tag1" not in data.columns:
        data["Tag1"] = ""

    data["Tag1"] = (
        data["Tag1"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    # 同じ日付内での元データ順を保持
    data["_original_order"] = range(len(data))

    result = {}

    # 各畝ごとに処理
    for bed, group in data.groupby("畝", sort=False):

        # その畝の最新日付
        latest_date = group["Date"].max()

        # 最新日付の全レコード
        latest_rows = group[
            group["Date"] == latest_date
        ].copy()

        latest_rows = latest_rows.sort_values(
            "_original_order"
        )

        display_items = []

        for _, row in latest_rows.iterrows():

            item = str(
                row["Name or Item"]
            ).strip()

            tag1 = str(
                row.get("Tag1", "")
            ).strip()

            if not item:
                continue

            # -----------------------------------------
            # 現況として重要な状態だけ付加
            # -----------------------------------------

            important_status = [
                "撤収",
                "撤収中",
                "撤収開始",
                "撤収完了",
                "収穫中",
                "収穫完了",
                "耕運",
                "耕運済",
                "耕運後",
            ]

            status = ""

            for keyword in important_status:
                if keyword in tag1:
                    status = tag1
                    break

            # 科名
            family = crop_family_dict.get(
                item,
                ""
            )

            if status:
                text = f"{item}　{family}　{status}"
            else:
                text = f"{item}　{family}"

            text = text.strip()

            # 重複除外
            if text not in display_items:
                display_items.append(text)

        result[bed] = {
            "item": " / ".join(display_items),
            "date": latest_date,
        }

    return result

    # --------------------------------------------------------
    # 元データの行順を保持
    #
    # 同じ畝・同じ日付に複数レコードがある場合、
    # Excel上で後に登録されているものを最新とする。
    # --------------------------------------------------------

    data["_original_order"] = range(len(data))

    data = data.sort_values(
        ["Date", "_original_order"]
    )

    # --------------------------------------------------------
    # 各畝について最後の1件
    # --------------------------------------------------------

    latest = (
        data
        .groupby(
            "畝",
            sort=False,
            as_index=False
        )
        .tail(1)
    )

    result = {}

    for _, row in latest.iterrows():

        bed = row["畝"]

        result[bed] = {
            "item": row["Name or Item"],
            "date": row["Date"],
        }

    return result


# ============================================================
# Excel出力
# ============================================================

def save_current_layout(
    latest_items,
    target_date,
):
    """
    指定日時点の作付け状態を
    「現在の作付け.xlsx」に出力する。
    """

    target_date = pd.to_datetime(target_date)

    wb = Workbook()

    ws = wb.active
    ws.title = "現在の作付け"

    # ========================================================
    # 色・罫線
    # ========================================================

    thin = Side(
        style="thin",
        color="000000"
    )

    border = Border(
        left=thin,
        right=thin,
        top=thin,
        bottom=thin,
    )

    header_fill = PatternFill(
        fill_type="solid",
        fgColor="D9EAD3"
    )

    bed_fill = PatternFill(
        fill_type="solid",
        fgColor="FFF2CC"
    )

    passage_fill = PatternFill(
        fill_type="solid",
        fgColor="D9D9D9"
    )

    # ========================================================
    # タイトル
    # ========================================================

    ws["A1"] = "日付"
    ws["B1"] = target_date.to_pydatetime()

    ws["A1"].font = Font(
        bold=True,
        size=12
    )

    ws["B1"].font = Font(
        bold=True,
        size=12
    )

    ws["B1"].number_format = "yyyy/mm/dd"

    # 1行空ける
    # 2行目は空白

    # ========================================================
    # 見出し
    # ========================================================

    header_row = 3

    headers = [
        "畝番号A",
        "作物　作業",
        "通路",
        "畝番号B",
        "作物　作業",
    ]

    for col, value in enumerate(
        headers,
        start=1
    ):

        cell = ws.cell(
            row=header_row,
            column=col,
            value=value
        )

        cell.font = Font(
            bold=True
        )

        cell.fill = header_fill

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

        cell.border = border

    # ========================================================
    # 畑レイアウト書き込み
    # ========================================================

    start_row = 4

    for index, (
        bed_a,
        bed_b
    ) in enumerate(GARDEN_LAYOUT):

        row = start_row + index

        # ----------------------------------------------------
        # A側 畝番号
        # ----------------------------------------------------

        ws.cell(
            row=row,
            column=1,
            value=bed_a
        )

        ws.cell(
            row=row,
            column=1
        ).border = border

        ws.cell(
            row=row,
            column=1
        ).alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

        # A側が通路の場合
        if bed_a == "通路":

            ws.cell(
                row=row,
                column=1
            ).fill = passage_fill

            ws.cell(
                row=row,
                column=2,
                value=""
            )

        else:

            ws.cell(
                row=row,
                column=1
            ).fill = bed_fill

            item_a = ""

            if bed_a in latest_items:
                item_a = latest_items[
                    bed_a
                ]["item"]

            ws.cell(
                row=row,
                column=2,
                value=item_a
            )

        ws.cell(
            row=row,
            column=2
        ).border = border

        ws.cell(
            row=row,
            column=2
        ).alignment = Alignment(
            horizontal="left",
            vertical="center"
        )

        # ----------------------------------------------------
        # 中央通路
        # ----------------------------------------------------

        ws.cell(
            row=row,
            column=3,
            value=""
        )

        ws.cell(
            row=row,
            column=3
        ).fill = passage_fill

        ws.cell(
            row=row,
            column=3
        ).border = border

        # ----------------------------------------------------
        # B側 畝番号
        # ----------------------------------------------------

        ws.cell(
            row=row,
            column=4,
            value=bed_b
        )

        ws.cell(
            row=row,
            column=4
        ).border = border

        ws.cell(
            row=row,
            column=4
        ).alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

        if bed_b:

            ws.cell(
                row=row,
                column=4
            ).fill = bed_fill

            item_b = ""

            if bed_b in latest_items:
                item_b = latest_items[
                    bed_b
                ]["item"]

            ws.cell(
                row=row,
                column=5,
                value=item_b
            )

        else:

            ws.cell(
                row=row,
                column=5,
                value=""
            )

        ws.cell(
            row=row,
            column=5
        ).border = border

        ws.cell(
            row=row,
            column=5
        ).alignment = Alignment(
            horizontal="left",
            vertical="center"
        )

    # ========================================================
    # 列幅
    # ========================================================

    column_widths = {
        "A": 12,
        "B": 180,
        "C": 8,
        "D": 12,
        "E": 180,
    }

    for col, width in column_widths.items():
        ws.column_dimensions[col].width = width

    # ========================================================
    # 行高さ
    # ========================================================

    ws.row_dimensions[1].height = 24
    ws.row_dimensions[3].height = 26

    for row in range(
        4,
        4 + len(GARDEN_LAYOUT)
    ):
        ws.row_dimensions[row].height = 24

    # ========================================================
    # ウィンドウ固定
    # ========================================================

    ws.freeze_panes = "A4"

    # ========================================================
    # 印刷設定
    # ========================================================

    ws.page_setup.orientation = "portrait"

    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1

    ws.sheet_properties.pageSetUpPr.fitToPage = True

    ws.print_area = (
        f"A1:E{3 + len(GARDEN_LAYOUT)}"
    )

    # ========================================================
    # 保存
    # ========================================================

    OUTPUT_FILE.parent.mkdir(
        parents=True,
        exist_ok=True
    )

    wb.save(OUTPUT_FILE)

    print()
    print("=" * 60)
    print("現在の作付けを作成しました")
    print("=" * 60)

    print(
        "対象日:",
        target_date.strftime("%Y-%m-%d")
    )

    print(
        "入力:",
        SOURCE_FILE
    )

    print(
        "出力:",
        OUTPUT_FILE
    )


# ============================================================
# 確認用表示
# ============================================================

def print_latest_items(
    latest_items,
    target_date
):
    """
    PowerShell上にも結果を表示する。
    """

    print()
    print(
        f"指定日時点の状態: "
        f"{pd.to_datetime(target_date):%Y-%m-%d}"
    )

    print("-" * 45)

    all_beds = (
        [f"A{i:02d}" for i in range(1, 18)]
        +
        [f"B{i:02d}" for i in range(1, 18)]
    )

    for bed in all_beds:

        info = latest_items.get(bed)

        if info:

            print(
                f"{bed}  "
                f"{info['item']:<15} "
                f"({info['date']:%Y-%m-%d})"
            )

        else:

            print(
                f"{bed}  "
                "データなし"
            )


# ============================================================
# 日付取得
# ============================================================

def get_target_date(
    command_line_date=None
):
    """
    --date が指定されていればそれを使用。
    なければ画面入力。
    """

    if command_line_date:

        try:
            return pd.to_datetime(
                command_line_date,
                format="%Y-%m-%d"
            )

        except ValueError:

            raise ValueError(
                "--date は YYYY-MM-DD "
                "形式で指定してください。"
            )

    while True:

        value = input(
            "日付を入力してください "
            "(YYYY-MM-DD): "
        ).strip()

        try:

            return pd.to_datetime(
                value,
                format="%Y-%m-%d"
            )

        except ValueError:

            print(
                "日付の形式が正しくありません。"
            )

            print(
                "例: 2026-09-03"
            )


# ============================================================
# メイン
# ============================================================

def main():

    parser = argparse.ArgumentParser(
        description=(
            "指定日時点の畑の状態を"
            "現在の作付け.xlsxに出力します。"
        )
    )

    parser.add_argument(
        "--date",
        help=(
            "基準日 YYYY-MM-DD "
            "例: --date 2026-09-03"
        )
    )

    args = parser.parse_args()

    # 日付
    target_date = get_target_date(
        args.date
    )

    # 元データ読み込み
    df = load_garden_data()

    # 作物名 → 科名
    crop_family_dict = (
        load_crop_family_dict()
    )

    # 各畝の最新状態
    latest_items = (
        get_latest_item_by_bed(
            df,
            target_date,
            crop_family_dict
        )
    )

    # PowerShell表示
    print_latest_items(
        latest_items,
        target_date
    )

    # Excel作成
    save_current_layout(
        latest_items,
        target_date
    )


# ============================================================
# 実行
# ============================================================

if __name__ == "__main__":
    main()