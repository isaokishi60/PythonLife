
import streamlit as st
import pandas as pd
import os
from PIL import Image
import datetime
import subprocess
import sys
from pathlib import Path
from layout_view import show_layout

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

# このファイル(start_page.py)がある場所
BASE_DIR = Path(__file__).resolve().parent

# OneDrive
ONEDRIVE = os.environ.get("OneDrive")

if not ONEDRIVE:
    raise RuntimeError("環境変数 OneDrive が見つかりません。")

# 農作業のルート
ROOT_DIR = (
    Path(ONEDRIVE)
    / "ドキュメント"
    / "PythonWork"
    / "農作業"
)

# Excelフォルダ
EXCEL_DIR = ROOT_DIR / "農作業関係Excel"

# 既存コードとの互換用
DATA_DIR = EXCEL_DIR

# データ読み込み
df_photo = pd.read_excel(
    EXCEL_DIR / "vegetable_garden_location.xlsx"
)

# ============================================================
# 作物マスター読み込み（科名取得用）
# ============================================================

df_crop_master = pd.read_excel(
    EXCEL_DIR / "作物マスター.xlsx"
)

# 作物名・科名の前後空白を除去
df_crop_master["作物名"] = (
    df_crop_master["作物名"]
    .astype(str)
    .str.strip()
)

df_crop_master["科（連作用）"] = (
    df_crop_master["科（連作用）"]
    .astype(str)
    .str.strip()
)

# 作型違いで同じ作物が複数行あるため、
# 作物名ごとに1件にして辞書化
crop_family_dict = (
    df_crop_master
    .drop_duplicates(
        subset=["作物名"]
    )
    .set_index("作物名")["科（連作用）"]
    .to_dict()
)

# 写真フォルダ候補
PHOTO_DIR_CANDIDATES = [
    Path(r"G:\その他のパソコン\マイ ノートパソコン\Pictures\Vegetables"),
    ROOT_DIR / "Pictures" / "Vegetables",
    Path(ONEDRIVE) / "Pictures" / "Vegetables",
]

photo_dir_path = None

for p in PHOTO_DIR_CANDIDATES:
    try:
        if p.exists():
            photo_dir_path = p
            break
    except Exception:
        pass

photo_dir = (
    str(photo_dir_path)
    if photo_dir_path is not None
    else ""
)

if photo_dir_path is None:
    st.warning(
        "写真フォルダが見つかりません。画像は表示できません。\n候補:\n- "
        + "\n- ".join(map(str, PHOTO_DIR_CANDIDATES))
    )

# ===== パス設定（Excel/写真）ここまで =====

# ============================================================
# 画面切り替え
# ============================================================

if "garden_view" not in st.session_state:
    st.session_state.garden_view = "start"


# レイアウト画面
if st.session_state.garden_view == "layout":

    if st.button("← スタート画面に戻る"):
        st.session_state.garden_view = "start"
        st.rerun()

    show_layout()

    # 下にスタート画面の内容を表示しない
    st.stop()

st.title("家庭菜園スタート画面")

# ==== ビュー切り替え ====
st.subheader("ビュー切り替え")

col_nav1, col_nav2 = st.columns(2)

# ガントチャートは従来どおり別画面
with col_nav1:
    st.markdown(
        '<a class="view-link-button" '
        'href="http://localhost:8502" '
        'target="_blank">'
        'ガントチャート'
        '</a>',
        unsafe_allow_html=True
    )

# レイアウトビューは同じStreamlit内で表示
with col_nav2:

    if st.button(
        "レイアウトビュー",
        key="btn_layout_view"
    ):
        st.session_state.garden_view = "layout"
        st.rerun()


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

btn_cols = st.columns(6)

with btn_cols[0]:
    show_record = st.button("品目別記録を見る", key="btn_record")

with btn_cols[1]:
    show_schedule = st.button("スケジュールを見る", key="btn_schedule")

with btn_cols[2]:
    show_crop = st.button("連作障害を見る", key="btn_crop")

with btn_cols[3]:
    show_comment = st.button("コメントを見る", key="btn_comment")

with btn_cols[4]:
    show_month_tasks = st.button("各月の作業を見る", key="btn_month_tasks")

with btn_cols[5]:
    show_bed_photos = st.button(
        "畝の写真を見る",
        key="btn_bed_photos"
    )


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

# ============================================================
# 検索方法を選択
# ============================================================

search_mode = st.radio(
    "検索方法を選んでください",
    [
        "品名から見る",
        "畝番号から見る",
    ],
    horizontal=True,
)

selected_item = None
selected_bed = None


# ============================================================
# 品名から見る
# ============================================================

if search_mode == "品名から見る":

    selected_item = st.selectbox(
        "品名を選んでください（カタカナ）",
        options,
        key="selected_item"
    )


# ============================================================
# 畝番号から見る
# ============================================================

else:

    # 畝番号一覧
    bed_options = sorted(
        df_photo["畝"]
        .dropna()
        .astype(str)
        .str.strip()
        .unique()
        .tolist()
    )

    selected_bed = st.selectbox(
        "畝番号を選んでください",
        bed_options,
        key="selected_bed"
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

# 選ばれた品名または畝番号を表示
if selected_item is not None:
    st.markdown(
        f'<p class="selected-item">選択中の品名：{selected_item}</p>',
        unsafe_allow_html=True
    )

if selected_bed is not None:
    st.markdown(
        f'<p class="selected-item">選択中の畝番号：{selected_bed}</p>',
        unsafe_allow_html=True
    )

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
                st.write("写真なし")


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


# ============================================================
# 畝ごとの過去写真一覧
# ============================================================

if show_bed_photos:

    st.subheader(
        f"{selected_bed} の過去写真"
    )

    # ============================================================
    # 指定畝 ＋ 指定期間の履歴
    # ============================================================

    df_bed = df_photo.copy()

    # 日付をdatetimeに変換
    df_bed["Date"] = pd.to_datetime(
        df_bed["Date"],
        errors="coerce"
    )

    # 畝番号 ＋ 開始日～終了日で絞り込み
    df_bed = df_bed[
        (
            df_bed["畝"]
            .astype(str)
            .str.strip()
            == selected_bed
        )
        &
        (
            df_bed["Date"].dt.date
            >= selected_start_date
        )
        &
        (
            df_bed["Date"].dt.date
            <= selected_end_date
        )
    ].copy()


    if df_bed.empty:

        st.info(
            f"{selected_bed} の "
            f"{selected_start_date} ～ {selected_end_date} "
            "の記録はありません。"
        )

    else:

        # 新しい順
        df_bed = df_bed.sort_values(
            "Date",
            ascending=False
        )

        # 同じ写真が複数の作物・記録で使われている場合、
        # JPG_Photo単位で重複表示しない
        df_bed_photo = (
            df_bed[
                df_bed["JPG_Photo"].notna()
            ]
            .drop_duplicates(
                subset=["JPG_Photo"],
                keep="first"
            )
        )

        if df_bed_photo.empty:

            st.info(
                f"{selected_bed} の写真は登録されていません。"
            )

        else:

            st.write(
                f"写真件数: {len(df_bed_photo)}"
            )

            for _, row in df_bed_photo.iterrows():

                photo_name = str(
                    row["JPG_Photo"]
                ).strip()

                photo_path = os.path.join(
                    photo_dir,
                    photo_name
                )

                # 日付
                if pd.notna(row["Date"]):
                    date_text = row["Date"].strftime(
                        "%Y-%m-%d"
                    )
                else:
                    date_text = "日付不明"

                # 作物・作業
                item = str(
                    row.get(
                        "Name or Item",
                        ""
                    )
                ).strip()

                # 科名
                family = crop_family_dict.get(
                    item,
                    ""
                )

                # 区画
                block = row.get(
                    "区画",
                    ""
                )

                if pd.isna(block):
                    block = ""
                else:
                    block = str(block).strip()

                # Tag
                tags = []

                for tag_col in [
                    "Tag1",
                    "Tag2",
                    "Tag3",
                    "Tag4",
                    "Tag5",
                ]:

                    if (
                        tag_col in row.index
                        and pd.notna(
                            row[tag_col]
                        )
                    ):

                        tag_text = str(
                            row[tag_col]
                        ).strip()

                        if tag_text:
                            tags.append(
                                tag_text
                            )
                # 畝状態
                bed_status = ""

                if (
                    "畝状態" in row.index
                    and pd.notna(row["畝状態"])
                ):
                    bed_status = str(
                        row["畝状態"]
                    ).strip()


                # 見出し
                title = (
                    f"{date_text}　"
                    f"{item}"
                )

                # 科名
                if family:
                    title += (
                        f"　{family}"
                    )

                # 区画
                if block:
                    title += (
                        f"　区画:{block}"
                    )

                st.markdown(
                    f"### {title}"
                )

                # Tag ＋ 畝状態を表示
                status_parts = []

                if tags:
                    status_parts.append(
                        " / ".join(tags)
                    )

                if bed_status:
                    status_parts.append(
                        f"畝状態：{bed_status}"
                    )

                if status_parts:
                    st.write(
                        "　｜　".join(status_parts)
                    )

                # 写真
                if os.path.exists(photo_path):

                    img = Image.open(
                        photo_path
                    )

                    st.image(
                        img,
                        caption=photo_name,
                        width=500
                    )

                else:

                    st.warning(
                        f"写真が見つかりません: "
                        f"{photo_path}"
                    )

                st.write("---")

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



