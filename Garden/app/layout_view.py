# -*- coding: utf-8 -*-

"""
layout_view.py

garden_current_status.py の判定結果を使って、
指定日時点の畑のレイアウトを Streamlit に表示する。
"""

from datetime import date

import streamlit as st
import pandas as pd

from garden_current_status import (
    load_garden_data,
    load_crop_family_dict,
    get_latest_item_by_bed,
)


# ============================================================
# 畑の物理レイアウト
# garden_current_status.py と同じ配置
# ============================================================

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
# 表示用データ作成
# ============================================================

def make_layout_dataframe(latest_items):
    """
    garden_current_status.py の結果から
    Streamlit表示用DataFrameを作る。
    """

    rows = []

    for bed_a, bed_b in GARDEN_LAYOUT:

        # A側
        if bed_a == "通路":
            item_a = ""
        else:
            info_a = latest_items.get(bed_a)
            item_a = (
                info_a["item"]
                if info_a
                else ""
            )

        # B側
        if bed_b:
            info_b = latest_items.get(bed_b)
            item_b = (
                info_b["item"]
                if info_b
                else ""
            )
        else:
            item_b = ""

        rows.append(
            {
                "畝番号A": bed_a,
                "作物　作業A": item_a,
                "通路": "",
                "畝番号B": bed_b,
                "作物　作業B": item_b,
            }
        )

    return pd.DataFrame(rows)


# ============================================================
# Streamlit画面
# ============================================================

def show_layout():
    """
    現在の畑レイアウト画面
    """

    st.title("畑の作付け状況")

    # --------------------------------------------------------
    # 日付指定
    # --------------------------------------------------------

    target_date = st.date_input(
        "表示する日付",
        value=date.today(),
        format="YYYY/MM/DD",
    )

    st.caption(
        "指定日以前で各畝の最も新しい記録を表示します。"
    )

    # --------------------------------------------------------
    # データ読み込み
    # --------------------------------------------------------

    try:
        df = load_garden_data()

    except Exception as e:
        st.error(
            f"vegetable_garden_location.xlsx "
            f"の読み込みに失敗しました。\n\n{e}"
        )
        return

    # --------------------------------------------------------
    # 指定日時点の状態取得
    # --------------------------------------------------------

    try:
        # 作物名 → 科名
        crop_family_dict = load_crop_family_dict()

        latest_items = get_latest_item_by_bed(
            df,
            target_date,
            crop_family_dict,
        )

    except Exception as e:
        st.error(
            f"作付け状態の取得に失敗しました。\n\n{e}"
        )
        return

    # --------------------------------------------------------
    # レイアウト作成
    # --------------------------------------------------------

    layout_df = make_layout_dataframe(
        latest_items
    )

    st.subheader(
        f"{target_date:%Y年%m月%d日} の作付け"
    )

    # --------------------------------------------------------
    # CSS
    # --------------------------------------------------------

    st.markdown(
        """
        <style>

        .garden-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 18px;
        }

        .garden-table th {
            background-color: #d9ead3;
            border: 1px solid #888;
            padding: 8px;
            text-align: center;
        }

        .garden-table td {
            border: 1px solid #aaa;
            padding: 8px;
            vertical-align: middle;
        }

        .garden-bed {
            background-color: #fff2cc;
            text-align: center;
            font-weight: bold;
            width: 10%;
        }

        .garden-item {
            width: 32%;
            font-size: 18px;
        }

        .garden-path {
            background-color: #d9d9d9;
            width: 7%;
            text-align: center;
        }

        .garden-passage-row {
            background-color: #d9d9d9;
            font-weight: bold;
            text-align: center;
        }

        </style>
        """,
        unsafe_allow_html=True,
    )

    # --------------------------------------------------------
    # HTMLテーブル
    # --------------------------------------------------------

    html = (
        '<table class="garden-table">'
        '<tr>'
        '<th>畝番号A</th>'
        '<th>作物　作業</th>'
        '<th>通路</th>'
        '<th>畝番号B</th>'
        '<th>作物　作業</th>'
        '</tr>'
    )

    for _, row in layout_df.iterrows():

        bed_a = row["畝番号A"]
        item_a = row["作物　作業A"]

        bed_b = row["畝番号B"]
        item_b = row["作物　作業B"]

        # A側に横通路がある場所
        if bed_a == "通路":

            html += (
                '<tr>'
                '<td class="garden-passage-row">通路</td>'
                '<td class="garden-passage-row"></td>'
                '<td class="garden-path"></td>'
                f'<td class="garden-bed">{bed_b}</td>'
                f'<td class="garden-item">{item_b}</td>'
                '</tr>'
            )

        else:

            html += (
                '<tr>'
                f'<td class="garden-bed">{bed_a}</td>'
                f'<td class="garden-item">{item_a}</td>'
                '<td class="garden-path"></td>'
                f'<td class="garden-bed">{bed_b}</td>'
                f'<td class="garden-item">{item_b}</td>'
                '</tr>'
            )

    html += '</table>'

    st.markdown(
        html,
        unsafe_allow_html=True,
    )


# ============================================================
# 単独実行時
# ============================================================

if __name__ == "__main__":
    show_layout()


