# 連続して空いている畝を探す20251202

import plotly.express as px
import pandas as pd
from datetime import timedelta, date
from pathlib import Path


# ==================================
# 設定
# ==================================
from pathlib import Path

# このファイル(空き畝検索.py)がある場所：...\農作業\Streamlit画面
BASE_DIR = Path(__file__).resolve().parent

# 農作業のルート：...\農作業
ROOT_DIR = BASE_DIR.parent

# Excelフォルダ：...\農作業\農作業関係Excel
EXCEL_DIR = ROOT_DIR / "農作業関係Excel"

# 出力フォルダ（無ければ画面フォルダに作るより、農作業直下にまとめるのがおすすめ）
OUT_DIR = ROOT_DIR / "outputs"
OUT_DIR.mkdir(exist_ok=True)

# 入力
FILE_PATH  = EXCEL_DIR / "作付計画.xlsx"      # ★最重要
RISK_FILE  = EXCEL_DIR / "連作障害.xlsx"
SHEET_NAME = "Sheet1"

# 出力
OUT_PATH = OUT_DIR / "畝候補一覧_品目単位.xlsx"


# 何か月以上空いていたら対象にするか
MIN_MONTHS_LIST = [2, 3]   # 2か月以上と3か月以上


# ==================================
# 1. 作付計画の読み込み
# ==================================
df_all = pd.read_excel(FILE_PATH, sheet_name=SHEET_NAME)

# 日付型にそろえる
df_all["開始日"] = pd.to_datetime(df_all["開始日"])
df_all["終了日"] = pd.to_datetime(df_all["終了日"])

# ★ 空き畝計算用（畝が入っている行だけ）
df = df_all.dropna(subset=["畝", "品目", "開始日", "終了日"]).copy()


# ==================================
# 2. 品目×畝ごとの「占有期間」を求める
#    畝つくり～撤収 を優先し、なければ最初～最後
# ==================================
occ_rows = []

for (bed, item), g in df.groupby(["畝", "品目"]):
    g = g.sort_values("開始日")
    work = g["作業"].astype(str)

    # 畝つくり
    mask_start = work.str.contains("畝つくり", na=False)
    if mask_start.any():
        occ_start = g.loc[mask_start, "開始日"].min()
    else:
        occ_start = g["開始日"].min()

    # 撤収
    mask_end = work.str.contains("撤収", na=False)
    if mask_end.any():
        occ_end = g.loc[mask_end, "終了日"].max()
    else:
        occ_end = g["終了日"].max()

    occ_rows.append({
        "畝": bed,
        "品目": item,
        "占有開始": occ_start,
        "占有終了": occ_end,
    })

df_occ = pd.DataFrame(occ_rows)

if df_occ.empty:
    print("占有データがありません。")
    raise SystemExit

# 全体期間（とりあえずデータ範囲）
PLANNING_START = df_occ["占有開始"].min().normalize()
PLANNING_END   = df_occ["占有終了"].max().normalize()
print(f"全体期間: {PLANNING_START.date()} ～ {PLANNING_END.date()}")


# ==================================
# 3. 畝ごとに占有期間をマージ → 空き期間を求める
# ==================================
def merge_intervals(g: pd.DataFrame):
    g = g.sort_values("占有開始").reset_index(drop=True)
    merged = []
    current_start = g.loc[0, "占有開始"]
    current_end = g.loc[0, "占有終了"]

    for i in range(1, len(g)):
        s = g.loc[i, "占有開始"]
        e = g.loc[i, "占有終了"]

        if s <= current_end + timedelta(days=1):
            if e > current_end:
                current_end = e
        else:
            merged.append((current_start, current_end))
            current_start, current_end = s, e

    merged.append((current_start, current_end))
    return merged

def bed_order(bed):
    """A01, A02, ..., A16, B01, ... B17 の順になるよう数値化する"""
    if pd.isna(bed):
        return 9999
    s = str(bed)
    block = s[0]           # A or B
    num = s[1:] if len(s) > 1 else "0"
    try:
        num = int(num)
    except ValueError:
        num = 0
    block_base = 0 if block == "A" else 100
    return block_base + num

free_rows = []

for bed, g in df_occ.groupby("畝"):
    merged = merge_intervals(g)

    # 1) 全体開始～最初の占有開始まで
    first_start, first_end = merged[0]
    if first_start > PLANNING_START:
        free_start = PLANNING_START
        free_end = first_start - timedelta(days=1)
        if free_end >= free_start:
            free_rows.append({"畝": bed, "空き開始": free_start, "空き終了": free_end})

    # 2) 占有と占有の間
    for (prev_s, prev_e), (next_s, next_e) in zip(merged, merged[1:]):
        free_start = prev_e + timedelta(days=1)
        free_end = next_s - timedelta(days=1)
        if free_end >= free_start:
            free_rows.append({"畝": bed, "空き開始": free_start, "空き終了": free_end})

    # 3) 最後の占有終了～全体終了まで
    last_s, last_e = merged[-1]
    if last_e < PLANNING_END:
        free_start = last_e + timedelta(days=1)
        free_end = PLANNING_END
        if free_end >= free_start:
            free_rows.append({"畝": bed, "空き開始": free_start, "空き終了": free_end})

df_free = pd.DataFrame(free_rows)

if df_free.empty:
    print("空き期間はありませんでした。")
    raise SystemExit

# ==================================
# 4. 冬(1,2,12月)を除いた「実際に使える空き期間」に切り分け
#    → 各空き期間を年ごとに 3/1〜11/30 にクリップ
# ==================================
usable_rows = []

for _, row in df_free.iterrows():
    bed = row["畝"]
    start = row["空き開始"].date()
    end = row["空き終了"].date()

    year = start.year
    while year <= end.year:
        # その年の栽培可能シーズン（3/1〜11/30）
        season_start = date(year, 3, 1)
        season_end   = date(year, 11, 30)

        # 元の空き期間との共通部分
        seg_start = max(start, season_start)
        seg_end   = min(end, season_end)

        if seg_start <= seg_end:
            usable_rows.append({
                "畝": bed,
                "空き開始": pd.Timestamp(seg_start),
                "空き終了": pd.Timestamp(seg_end),
            })

        year += 1

df_free = pd.DataFrame(usable_rows)

if df_free.empty:
    print("（3〜11月の間に使える空き期間はありませんでした）")
    raise SystemExit

# 日数・月数を計算
df_free["空き日数"] = (df_free["空き終了"] - df_free["空き開始"]).dt.days + 1
df_free["空き月数(概算)"] = (df_free["空き日数"] / 30.0).round(1)

# ==================================
# 連作障害リスクが「高」の畝リスト（品目別）を作成
# ==================================
RISK_FILE = BASE_DIR / "連作障害.xlsx"

RISK_SHEET = "Crop_Performance"

try:
    df_risk = pd.read_excel(RISK_FILE, sheet_name=RISK_SHEET)

    # 列名は実際のファイルに合わせてここを調整してください
    # ここでは「品目」「畝」「リスク」という列名を想定
    df_risk["品目"] = df_risk["品目"].astype(str).str.strip()
    df_risk["畝"] = df_risk["畝"].astype(str).str.strip()

    # 品目ごとに「リスク=高」の畝セットを作る
    risk_high_map = (
        df_risk[df_risk["リスク"] == "高"]
        .groupby("品目")["畝"]
        .apply(lambda s: set(s.astype(str)))
        .to_dict()
    )

    print("★ 連作障害リスク【高】の畝（品目別）:", risk_high_map)

except Exception as e:
    print("※ 連作障害ファイルの読み込みでエラーが出ました。リスク判定はスキップします:", e)
    risk_high_map = {}


# ==================================
# 6. 畝が決まっていない「品目」ごとに、全作業期間をカバーできる畝候補を列挙
# ==================================

# 畝が空欄 or NaN の行だけ抽出（作付計画そのものから）
df_no_bed = df_all[
    (df_all["畝"].isna() | (df_all["畝"].astype(str).str.strip() == ""))
    & df_all["品目"].notna()
    & df_all["開始日"].notna()
    & df_all["終了日"].notna()
].copy()

if df_no_bed.empty:
    print("\n※ 畝が未入力の作付行はありませんでした。")
else:
    print("\n★ 畝未決の品目ごとに、全作業期間をカバーできる畝候補一覧を作成中…")

    # まず「品目ごと」に占有開始〜終了を決める
    occ_rows = []
    for item, g in df_no_bed.groupby("品目"):
        g = g.sort_values("開始日")
        work = g["作業"].astype(str)

        # 畝つくり～撤収 のルールは、既存の占有計算と同じ考え方
        mask_start = work.str.contains("畝つくり", na=False)
        if mask_start.any():
            occ_start = g.loc[mask_start, "開始日"].min()
        else:
            occ_start = g["開始日"].min()

        mask_end = work.str.contains("撤収", na=False)
        if mask_end.any():
            occ_end = g.loc[mask_end, "終了日"].max()
        else:
            occ_end = g["終了日"].max()

        occ_rows.append({
            "品目": item,
            "占有開始": occ_start,
            "占有終了": occ_end,
        })

    df_no_bed_occ = pd.DataFrame(occ_rows)

    rows = []
    for _, r in df_no_bed_occ.iterrows():
        item  = r["品目"]
        start = pd.to_datetime(r["占有開始"])
        end   = pd.to_datetime(r["占有終了"])

        # ★ この品目の「全作業期間（占有開始〜占有終了）」を丸ごと含む空き畝を検索
        cand = df_free[
            (df_free["空き開始"] <= start) &
            (df_free["空き終了"] >= end)
        ]

        # ★ ここで「リスク=高」の畝を除外（品目別）
        bad_beds = risk_high_map.get(str(item).strip(), set())
        if bad_beds:
            cand = cand[~cand["畝"].astype(str).str.strip().isin(bad_beds)]


        # 畝を並び順どおりにソート
        beds = sorted(cand["畝"].dropna().unique(), key=bed_order)

        if beds:
            cand_str = ", ".join(map(str, beds))
        else:
            cand_str = "候補なし"

        rows.append({
            "品目": item,
            "占有開始": start.date(),
            "占有終了": end.date(),
            "候補畝一覧": cand_str,
        })

    df_candidates = pd.DataFrame(rows)

    # Excel に出力（パスはお好みで）
    OUT_PATH = BASE_DIR / "畝候補一覧_品目単位.xlsx"

    df_candidates.to_excel(OUT_PATH , index=False)
    print("畝候補一覧（品目単位）を出力しました:", OUT_PATH )




# ==================================
# 5. 2か月以上 / 3か月以上の空き畝リストを出力
# ==================================
for min_months in MIN_MONTHS_LIST:
    min_days = int(min_months * 30)

    df_long = df_free[df_free["空き日数"] >= min_days].copy()
    if df_long.empty:
        #print(f"\n★ {min_months}か月以上連続して空いている畝はありませんでした。")
        continue

    #print(f"\n★ {min_months}か月以上連続して空いている畝のリスト")
    df_show = df_long.sort_values(["畝", "空き開始"]).copy()
    df_show["空き開始"] = df_show["空き開始"].dt.date
    df_show["空き終了"] = df_show["空き終了"].dt.date

    #print(df_show[["畝", "空き開始", "空き終了", "空き日数", "空き月数(概算)"]].to_string(index=False))

# ==================================
# ★ 空き畝ガントチャートの作成
# ==================================

# 2か月以上 or 3か月以上の空きを選ぶ
MIN_MONTHS_FOR_GANTT = 3.0   # ← 3か月以上にしたいときは 3.0 に変更

df_gantt = df_free[df_free["空き月数(概算)"] >= MIN_MONTHS_FOR_GANTT].copy()
if df_gantt.empty:
    print("\n※ ガントチャートに表示できる空き畝がありません。")
else:
    # 畝の並び順用
    df_gantt["畝順"] = df_gantt["畝"].apply(bed_order)

    # 年度（x軸の表示範囲用）を自動取得（空き開始の年を使う）
    year_min = int(df_gantt["空き開始"].dt.year.min())
    year_max = int(df_gantt["空き終了"].dt.year.max())
    # 普通は同じ年だと思うが、念のため
    year = year_min

    # 並び替え
    df_gantt = df_gantt.sort_values(["畝順", "空き開始"])

    # Plotly のタイムライン（ガントチャート風）
    fig = px.timeline(
        df_gantt,
        x_start="空き開始",
        x_end="空き終了",
        y="畝",
        hover_data=["空き開始", "空き終了", "空き日数", "空き月数(概算)"],
    )

    # 畝を上から A01, A02, ... Bxx の順に
    category_order = (
        df_gantt[["畝", "畝順"]]
        .drop_duplicates()
        .sort_values("畝順")["畝"]
        .tolist()
    )
    fig.update_yaxes(
        categoryorder="array",
        categoryarray=category_order,
        autorange="reversed"   # ★ これで期待どおり上から昇順になる
    )

    # x軸を 1〜12月のイメージに（dtick=月ごと、フォーマットを「m月」に）
    fig.update_xaxes(
        dtick="M1",
        tickformat="%m月",
        # もし 1〜12月全体を出したいなら range を固定
        range=[
            pd.Timestamp(f"{year}-01-01"),
            pd.Timestamp(f"{year}-12-31"),
        ],
    )

    fig.update_layout(
        title=f"空き畝ガントチャート（空き {MIN_MONTHS_FOR_GANTT:.1f} ヶ月以上）",
        xaxis_title="月",
        yaxis_title="畝",
        height=800,
        margin=dict(l=80, r=20, t=60, b=40),
    )

    # ブラウザで表示
    fig.show()
