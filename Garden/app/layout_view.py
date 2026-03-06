import streamlit as st
import pandas as pd
import calendar
import plotly.graph_objects as go
from pathlib import Path

# =========================
# 1. データ読み込み
# =========================
BASE_DIR = Path(__file__).resolve().parent
ROOT_DIR = BASE_DIR.parent
EXCEL_DIR = ROOT_DIR / "農作業関係Excel"

PLAN_PATH = EXCEL_DIR / "作付計画.xlsx"
LAYOUT_PATH = EXCEL_DIR / "Layout.xlsx"

df_plan = pd.read_excel(PLAN_PATH)
df_plan.columns = df_plan.columns.map(lambda x: str(x).strip())
df_plan["開始日"] = pd.to_datetime(df_plan["開始日"], errors="coerce")
df_plan["終了日"] = pd.to_datetime(df_plan["終了日"], errors="coerce")
df_plan["畝"] = df_plan["畝"].astype(str)

df_layout = pd.read_excel(LAYOUT_PATH, sheet_name="レイアウト")
df_layout.columns = df_layout.columns.map(lambda x: str(x).strip())
df_layout["畝"] = df_layout["畝"].astype(str)
df_layout["行"] = pd.to_numeric(df_layout["行"], errors="coerce")
df_layout["列"] = pd.to_numeric(df_layout["列"], errors="coerce")

# ★ 通路フラグ（無ければ False）
if "通路" in df_layout.columns:
    df_layout["通路"] = df_layout["通路"].fillna(0).astype(bool)
else:
    df_layout["通路"] = False

# =========================
# 2. 補助関数
# =========================
def get_period_range(year: int, month: int, part: str):
    _, last_day = calendar.monthrange(year, month)
    if part == "上旬":
        start_day, end_day = 1, 10
    elif part == "中旬":
        start_day, end_day = 11, 20
    else:
        start_day, end_day = 21, last_day
    return (
        pd.Timestamp(year=year, month=month, day=start_day),
        pd.Timestamp(year=year, month=month, day=end_day),
    )

def compute_bed_occupancy(df_plan: pd.DataFrame, df_layout: pd.DataFrame, year: int, month: int, part: str):
    period_start, period_end = get_period_range(year, month, part)
    dfp = df_plan.copy()
    dfp = dfp.dropna(subset=["畝", "品目", "開始日", "終了日"])

    occ_rows = []
    for (bed, item), g in dfp.groupby(["畝", "品目"]):
        g = g.sort_values("開始日")
        work = g["作業"].astype(str)

        mask_start = work.str.contains("畝つくり", na=False)
        occ_start = g.loc[mask_start, "開始日"].min() if mask_start.any() else g["開始日"].min()

        mask_end = work.str.contains("撤収", na=False)
        occ_end = g.loc[mask_end, "終了日"].max() if mask_end.any() else g["終了日"].max()

        occ_rows.append({"畝": str(bed), "品目": item, "占有開始": occ_start, "占有終了": occ_end})

    df_occ_range = pd.DataFrame(occ_rows)

    df_out = df_layout.copy()
    if df_occ_range.empty:
        df_out["占有中"] = False
        df_out["品目一覧"] = ""
        return df_out

    df_occ_range["占有中"] = ~(
        (df_occ_range["占有終了"] < period_start) |
        (df_occ_range["占有開始"] > period_end)
    )

    df_items = (
        df_occ_range[df_occ_range["占有中"]]
        .groupby("畝")["品目"]
        .apply(lambda x: "、".join(x))
    )
    occ_by_bed = df_occ_range.groupby("畝")["占有中"].any()

    df_out["占有中"] = df_out["畝"].map(occ_by_bed).fillna(False)
    df_out["品目一覧"] = df_out["畝"].map(df_items).fillna("")
    return df_out

# =========================
# 3. Streamlit UI
# =========================
st.title("畝の占有レイアウト（10日ごと）")

years = sorted(int(y) for y in df_plan["開始日"].dt.year.dropna().unique())
if not years:
    st.error("作付計画.xlsx の「開始日」から年が取得できません。列名や日付が正しいか確認してください。")
    st.write("df_plan columns:", df_plan.columns.tolist())
    st.stop()

default_year = years[-1]
year = st.selectbox("年を選択", years, index=years.index(default_year))

month = st.selectbox("月を選択", list(range(1, 13)), index=0)
part = st.selectbox("期間を選択", ["上旬", "中旬", "下旬"])

df_out = compute_bed_occupancy(df_plan, df_layout, year, month, part)

st.write(f"表示対象: {year}年{month}月{part}")

# =========================
# 4. プロット作成
# =========================
max_row = int(df_layout["行"].max())
max_col = int(df_layout["列"].max())

fig = go.Figure()

for _, row in df_out.iterrows():
    r = int(row["行"])
    c = int(row["列"])
    bed = row["畝"]
    items = row["品目一覧"]
    is_aisle = bool(row.get("通路", False))
    occupied = bool(row["占有中"])

    if is_aisle:
        fill_color = "#666666"   # 濃い灰色（通路）
    elif occupied:
        fill_color = "#ffcccc"   # 占有中
    else:
        fill_color = "#ccffcc"   # 空き


    fig.add_shape(
        type="rect",
        x0=c - 0.5, y0=r - 0.5,
        x1=c + 0.5, y1=r + 0.5,
        line=dict(color="black", width=1),
        fillcolor=fill_color,
    )

    text = ""
    if not is_aisle:
        text = f"{bed}"
        if items:
            text += f"\n{items}"

    fig.add_trace(go.Scatter(
        x=[c], y=[r],
        text=[text],
        mode="text",
        textposition="middle center",
        showlegend=False,
    ))

fig.update_xaxes(range=[0.5, max_col + 0.5], dtick=1, title="列")
fig.update_yaxes(range=[max_row + 0.5, 0.5], dtick=1, title="行")
fig.update_layout(height=40 * max_row, margin=dict(l=40, r=40, t=40, b=40))

st.plotly_chart(fig, use_container_width=True)
st.dataframe(df_out, use_container_width=True)


