
import streamlit as st
import pandas as pd
import os
from PIL import Image
import datetime
import subprocess
import sys
from pathlib import Path

# 🍄 CSSでボタンと品名の見た目をカスタマイズ
st.markdown("""
    <style>
    /* Streamlit標準ボタン（記録を見る など）用：薄い青 */
    div.stButton > button {
        background-color: #d0f0ff;  /* 薄い青 */
        border: 2px solid #88c;
        border-radius: 8px;
        box-shadow: 2px 2px 5px rgba(0,0,0,0.2);
        padding: 10px 20px;
        font-weight: bold;
        color: #003366;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown("""
    <style>
    /* ビュー切り替え用リンクボタン：薄い緑 */
    .view-link-button {
        background-color: #d8f6d3;  /* 薄い緑 */
        border: 2px solid #6ca86c;
        border-radius: 8px;
        box-shadow: 2px 2px 5px rgba(0,0,0,0.2);
        padding: 10px 20px;
        font-weight: bold;
        color: #003300;
        text-decoration: none;
        display: inline-block;
        text-align: center;
        min-width: 150px;
    }
    </style>
""", unsafe_allow_html=True)

# ===== パス設定（Excel/写真）ここから =====
# このファイル(start_page.py)がある場所：...\農作業\Streamlit画面
BASE_DIR = Path(__file__).resolve().parent

# 農作業のルート：...\農作業
ROOT_DIR = BASE_DIR.parent

# Excelフォルダ：...\農作業\農作業関係Excel
EXCEL_DIR = ROOT_DIR / "農作業関係Excel"

# 以前 DATA_DIR を使っている箇所があるので、互換のために揃えておく
DATA_DIR = EXCEL_DIR

# データ読み込み（Excelの場所をEXCEL_DIRに変更）
df_photo = pd.read_excel(EXCEL_DIR / "vegetable_garden_photo_ex2_with_bed.xlsx")

# 写真フォルダ候補（上から順に探す）
PHOTO_DIR_CANDIDATES = [
    Path(r"G:\その他のパソコン\マイ ノートパソコン\Pictures\Vegetables"),  # 旧（存在すれば使う）
    ROOT_DIR / "Pictures" / "Vegetables",                                  # 推奨：農作業配下
    Path(os.environ.get("OneDrive", "")) / "Pictures" / "Vegetables",      # OneDrive配下
]

photo_dir_path = None
for p in PHOTO_DIR_CANDIDATES:
    try:
        if str(p) and p.exists():
            photo_dir_path = p
            break
    except Exception:
        pass

# 文字列として使いたい箇所のために str も用意
photo_dir = str(photo_dir_path) if photo_dir_path is not None else ""

if photo_dir_path is None:
    st.warning(
        "写真フォルダが見つかりません。画像は表示できません。\n候補:\n- "
        + "\n- ".join(map(str, PHOTO_DIR_CANDIDATES))
    )
# ===== パス設定（Excel/写真）ここまで =====


st.title("家庭菜園スタート画面")

# ==== ビュー切り替え（薄い緑ボタン） ====
st.subheader("ビュー切り替え")

col_nav1, col_nav2 = st.columns(2)

with col_nav1:
    st.markdown(
        '<a class="view-link-button" href="http://localhost:8502" target="_blank">ガントチャート</a>',
        unsafe_allow_html=True
    )

with col_nav2:
    st.markdown(
        '<a class="view-link-button" href="http://localhost:8503" target="_blank">レイアウトビュー</a>',
        unsafe_allow_html=True
    )


st.write("---")

# このファイル(start_page.py)があるフォルダ
BASE_DIR = Path(__file__).resolve().parent

# 空き畝検索スクリプトのパス
AKI_SCRIPT = BASE_DIR / "空き畝検索.py"


# ==== 空き畝検索ツール ====
st.subheader("空き畝ツール")

if st.button("空き畝ガントチャートを表示", key="btn_free_bed_gantt"):
    if not AKI_SCRIPT.exists():
        st.error(f"空き畝検索スクリプトが見つかりません: {AKI_SCRIPT}")
    else:
        python_exe = sys.executable  # 今動いている Python (venv311)
        try:
            # 空き畝検索.py を別プロセスで起動（非同期で立ち上げる）
            subprocess.Popen([python_exe, str(AKI_SCRIPT)])
            st.info("空き畝ガントチャートを別ウィンドウ／タブで開きます。少し待ってブラウザを確認してください。")
        except Exception as e:
            st.error(f"空き畝検索スクリプトの起動に失敗しました: {e}")


# ==== 機能ボタン（薄い青・横一列） ====
st.subheader("家庭菜園スタートメニュー")

btn_cols = st.columns(5)

with btn_cols[0]:
    show_record = st.button("記録を見る", key="btn_record")

with btn_cols[1]:
    show_schedule = st.button("スケジュールを見る", key="btn_schedule")

with btn_cols[2]:
    show_crop = st.button("連作障害を見る", key="btn_crop")

with btn_cols[3]:
    show_comment = st.button("コメントを見る", key="btn_comment")

with btn_cols[4]:
    show_month_tasks = st.button("各月の作業を見る", key="btn_month_tasks")


st.write("---")

# 品名と期間を選択
# --- Name or Item を整形して一覧を作る ---
names = (
    df_photo["Name or Item"]
    .dropna()
    .astype(str)
    .str.replace("　", "", regex=False)  # 全角スペース削除（お好みで）
    .str.strip()                         # 前後の空白削除
)


options = sorted(names.unique().tolist())

# --- 品名を選択（selectbox） ---
selected_item = st.selectbox(
    "品名を選んでください（カタカナ）",
    options
)


# CSSで文字を大きく太く
st.markdown("""
    <style>
    .selected-item {
        font-size: 28px;
        font-weight: 900;
        font-family: 'Yu Gothic', 'Meiryo', sans-serif;
        color: #333;
        margin-top: 20px;
    }
    </style>
""", unsafe_allow_html=True)

# 選ばれた品名を表示
st.markdown(f'<p class="selected-item">選択中の品名：{selected_item}</p>', unsafe_allow_html=True)

default_start = datetime.date(2025, 1, 1)
selected_start_date = st.date_input("開始日を選んでください", value=default_start)

selected_end_date = st.date_input("終了日を入力してください")


# 記録ボタン
if show_record:
    df_filtered = df_photo[
        (df_photo["Name or Item"] == selected_item) &
        (pd.to_datetime(df_photo["Date"]).dt.date >= selected_start_date) &
        (pd.to_datetime(df_photo["Date"]).dt.date <= selected_end_date)
    ]

    if df_filtered.empty:
        st.warning("データが見つかりませんでした。")
    else:
        for _, row in df_filtered.iterrows():
            st.write(f"📅 日付: {row['Date']}")
            st.write(f"📌 品名: {row['Name or Item']}")

            # 🌱 畝番号の表示（区画がある場合は併記）
            bed = row.get("畝", None)
            block = row.get("区画", None)

            if pd.notna(bed):
                if pd.notna(block) and str(block).strip() != "":
                    st.write(f"🌱 畝: {bed}（{block}）")   # 例）A02（南）
                else:
                    st.write(f"🌱 畝: {bed}")             # 例）A05

            st.write(f"🏷️ タグ: {', '.join(str(row[tag]) for tag in ['Tag1','Tag2','Tag3','Tag4','Tag5'] if pd.notna(row[tag]))}")

            # JPG_Photo から絶対パスを生成
            if pd.notna(row["JPG_Photo"]):
                photo_path = os.path.join(photo_dir, str(row["JPG_Photo"]))
                if os.path.exists(photo_path):
                    img = Image.open(photo_path)
                    st.image(img, caption=row["Name or Item"], width=400)

                else:
                    st.error(f"写真が見つかりません: {photo_path}")
            else:
                st.write(f"Photo_id: {row['Photo_id']}")


# 絶対パスを作成
#photo_path = os.path.join(photo_dir, str(row["JPG_Photo"]))

# コメント表示ボタン
if show_comment:
    comment_path = DATA_DIR / "野菜育成コメント.xlsx"

    try:
        df_comment = pd.read_excel(comment_path, sheet_name="Sheet1", header=None)
        df_comment.columns = ["野菜名", "コメント1", "コメント2", "コメント3", "コメント4", "コメント5"]

        # 選択された品名に一致する行を取得
        comment_row = df_comment[df_comment["野菜名"] == selected_item]

        if comment_row.empty:
            st.info("この品目の育成コメントは登録されていません。")
        else:
            st.subheader(f"{selected_item} の育成コメント")
            for i in range(1, 6):
                comment = comment_row.iloc[0][f"コメント{i}"]
                if pd.notna(comment):
                    st.write(f"💬 コメント{i}: {comment}")
                else:
                    st.write(f"💬 コメント{i}: （未登録）")

    except Exception as e:
        st.error(f"コメントファイルの読み込み中にエラーが発生しました: {e}")


# 作業カレンダーの読み込み（ヘッダーなしで読み込む）
calendar_path = DATA_DIR / "作業カレンダー.xlsx"
df_calendar = pd.read_excel(calendar_path, sheet_name="Sheet1", header=None)

# 品目一覧（1行目のカタカナ名）
item_names = df_calendar.iloc[0, 2:].tolist()

# 選択された品目の列番号を取得
try:
    item_index = item_names.index(selected_item) + 2  # 0-based → Excel列は2列目から
except ValueError:
    st.warning("選択された品目のスケジュールが見つかりませんでした。")
    item_index = None

# スケジュール表示
# 作業スケジュールボタン
if show_schedule:
    calendar_path = DATA_DIR / "作業カレンダー.xlsx"
    df_calendar = pd.read_excel(calendar_path, sheet_name="Sheet1", header=None)



    item_names = df_calendar.iloc[0, 2:].tolist()
    try:
        item_index = item_names.index(selected_item) + 2
    except ValueError:
        st.warning("選択された品目のスケジュールが見つかりませんでした。")
        item_index = None

    months = ["1月", "2月", "3月", "4月", "5月", "6月",
              "7月", "8月", "9月", "10月", "11月", "12月"]
    periods = ["上旬", "中旬", "下旬"]
    month_start_rows = [3 + i * 3 for i in range(12)]

    if item_index is not None:
        st.subheader(f"{selected_item} の年間作業スケジュール")
        for month, start_row in zip(months, month_start_rows):
            st.markdown(f"### {month}")
            for i, period in enumerate(periods):
                row_index = start_row + i
                try:
                    task = df_calendar.iloc[row_index, item_index]
                    task_display = task if pd.notna(task) else "-"
                except IndexError:
                    task_display = "-"
                st.write(f"🕒 {period}: {task_display}")

# 連作障害ボタン
if show_crop:
    crop_path = DATA_DIR / "連作障害.xlsx"
    df_crop = pd.read_excel(crop_path, sheet_name="Crop_Performance")

    df_crop_filtered = df_crop[df_crop["Name"] == selected_item]

    if df_crop_filtered.empty:
        st.info("この品目の連作障害情報は登録されていません。")
    else:
        row = df_crop_filtered.iloc[0]
        st.subheader(f"{selected_item} の連作障害情報")
        st.write(f"🧬 野菜名: {row['野菜名']}")
        st.write(f"🌿 科: {row['科']}")
        st.write(f"⚠️ リスク: {row['リスク']}")
        st.write(f"🩺 主な障害: {row['主な障害']}")
        st.write(f"📆 推奨年数: {row['年数']}")
        st.write(f"📝 備考: {row['備考']}")


# ===== 各月の作業（作業カレンダー.xlsx から作る）ここから =====
# 作業カレンダー（ヘッダーなし）
calendar_path = DATA_DIR / "作業カレンダー.xlsx"
df_calendar = pd.read_excel(calendar_path, sheet_name="Sheet1", header=None)

# 品目一覧（1行目のカタカナ名）: C列(=index2)以降
item_names = df_calendar.iloc[0, 2:].tolist()

months = ["1月", "2月", "3月", "4月", "5月", "6月",
          "7月", "8月", "9月", "10月", "11月", "12月"]
periods = ["上旬", "中旬", "下旬"]

# 月選択（ボタンの外に置く）
selected_month_for_tasks = st.selectbox(
    "【月別作業】表示する月を選んでください",
    months,
    key="month_tasks_month"
)

# ボタンが押されたときにだけ表示する
if show_month_tasks:
    month_index = months.index(selected_month_for_tasks)
    start_row = 3 + month_index * 3  # 1月上旬が row=3 の前提（あなたの既存ロジックと同じ）

    rows_out = []
    for j, item in enumerate(item_names):
        col = 2 + j  # 品目列の実データ列

        # 上旬/中旬/下旬の3マスを取る
        vals = []
        for i, period in enumerate(periods):
            r = start_row + i
            v = df_calendar.iloc[r, col] if r < len(df_calendar) else None
            if pd.notna(v) and str(v).strip() != "":
                vals.append(f"{period}:{v}")

        if vals:
            rows_out.append({"品名": item, "作業": " / ".join(vals)})

    if not rows_out:
        st.info(f"{selected_month_for_tasks} の作業は登録されていません。")
    else:
        st.subheader(f"{selected_month_for_tasks} の作業リスト（作業カレンダー.xlsx）")

        # 既存のCSS（横並び表示）を流用
        st.markdown("""
            <style>
            .task-line { font-size: 18px; font-weight: 500; font-family: 'Yu Gothic','Meiryo',sans-serif; margin: 4px 0; }
            .task-name { display: inline-block; width: 150px; font-weight: 600; color: #333; }
            .task-content { display: inline-block; color: #444; }
            </style>
        """, unsafe_allow_html=True)

        for r in rows_out:
            st.markdown(
                f'<div class="task-line">'
                f'<span class="task-name">{r["品名"]}</span>'
                f'<span class="task-content">{r["作業"]}</span>'
                f'</div>',
                unsafe_allow_html=True
            )

        st.caption(f"元データ: {calendar_path}")



