import streamlit as st
import pandas as pd
import plotly.express as px
from pathlib import Path

# =========================
# 基本設定
# =========================
st.set_page_config(page_title="作付ガント", layout="wide")

BASE_DIR = Path(__file__).resolve().parent          # ...\農作業\Streamlit画面
ROOT_DIR = BASE_DIR.parent                           # ...\農作業
EXCEL_DIR = ROOT_DIR / "農作業関係Excel"             # ...\農作業\農作業関係Excel

PLAN_PATH = EXCEL_DIR / "作付計画.xlsx"
RULE_PATH = EXCEL_DIR / "畝配分.xlsx"               # 品目マスター / 作物グループ が入っている想定

RISK_PATH = EXCEL_DIR / "連作障害.xlsx"
RISK_SHEET = "Crop_Performance"

CROP_RISK_PATH = EXCEL_DIR / "連作障害.xlsx"

# =========================
# 便利関数
# =========================
def _strip_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = df.columns.map(lambda x: str(x).strip())
    return df

def load_item_master(rule_path: Path) -> pd.DataFrame:
    """
    畝配分.xlsx の「品目マスター」
      品目（作付計画）, 品目_base, 品目_var
    """
    try:
        m = pd.read_excel(rule_path, sheet_name="品目マスター")
        m = _strip_columns(m)
        # 想定列名に寄せる（念のため）
        col_map = {}
        for c in m.columns:
            if c.replace(" ", "") in ["品目(作付計画)", "品目（作付計画）"]:
                col_map[c] = "品目（作付計画）"
        if col_map:
            m = m.rename(columns=col_map)

        need = ["品目（作付計画）", "品目_base", "品目_var"]
        for c in need:
            if c not in m.columns:
                raise KeyError(f"品目マスターに列 {c} がありません。現在列={m.columns.tolist()}")

        m["品目（作付計画）"] = m["品目（作付計画）"].astype(str).str.strip()
        m["品目_base"] = m["品目_base"].astype(str).str.strip()
        m["品目_var"] = m["品目_var"].astype(str).fillna("").str.strip()
        return m[need].drop_duplicates()
    except Exception:
        # 無くても動く（品目_base=品目）にフォールバック
        return pd.DataFrame(columns=["品目（作付計画）", "品目_base", "品目_var"])

def apply_item_master(df: pd.DataFrame, master: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["品目"] = df["品目"].astype(str).str.strip()

    if master.empty:
        df["品目_base"] = df["品目"]
        df["品目_var"] = ""
        return df

    mp_base = dict(zip(master["品目（作付計画）"], master["品目_base"]))
    mp_var  = dict(zip(master["品目（作付計画）"], master["品目_var"]))

    df["品目_base"] = df["品目"].map(mp_base).fillna(df["品目"])
    df["品目_var"]  = df["品目"].map(mp_var).fillna("")
    return df

def load_crop_groups(rule_path: Path) -> dict:
    """
    畝配分.xlsx の「作物グループ」
    （あなたのシートは1列目=品目_base、2列目=グループ の形式でOK）
    """
    try:
        g = pd.read_excel(rule_path, sheet_name="作物グループ", header=None)
        g = g.dropna(how="all")
        if g.shape[1] < 2:
            return {}
        g = g.iloc[:, :2]
        g.columns = ["品目_base", "グループ"]
        g["品目_base"] = g["品目_base"].astype(str).str.strip()
        g["グループ"] = g["グループ"].astype(str).str.strip()
        return dict(zip(g["品目_base"], g["グループ"]))
    except Exception:
        return {}
    
def load_crop_family_map(path: Path) -> dict:
    """
    連作障害.xlsx / Crop_Performance
    野菜名 → 科 の辞書を作る
    """
    try:
        df = pd.read_excel(path, sheet_name="Crop_Performance")
        df.columns = df.columns.map(lambda x: str(x).strip())

        required = ["野菜名", "科"]
        for c in required:
            if c not in df.columns:
                raise KeyError(f"Crop_Performance に列 {c} がありません")

        df["野菜名"] = df["野菜名"].astype(str).str.strip()
        df["科"] = df["科"].astype(str).str.strip()

        return dict(zip(df["野菜名"], df["科"]))

    except Exception as e:
        st.warning(f"連作障害ファイルの読込に失敗: {e}")
        return {}


def has_overlap(a_start, a_end, b_start, b_end, min_overlap_days=1) -> bool:
    """2区間が min_overlap_days 日以上重なっているか"""
    latest_start = max(a_start, b_start)
    earliest_end = min(a_end, b_end)
    overlap_days = (earliest_end - latest_start).days + 1
    return overlap_days >= min_overlap_days


# =========================
# 1) データ読み込み
# =========================
if not PLAN_PATH.exists():
    st.error(f"作付計画.xlsx が見つかりません: {PLAN_PATH}")
    st.stop()

df_plan = pd.read_excel(PLAN_PATH)
df_plan = _strip_columns(df_plan)

# 必須列チェック
required = ["畝", "品目", "開始日", "終了日"]
missing = [c for c in required if c not in df_plan.columns]
if missing:
    st.error(f"作付計画.xlsx に必須列がありません: {missing}\n現在列={df_plan.columns.tolist()}")
    st.stop()

# 日付型
df_plan["開始日"] = pd.to_datetime(df_plan["開始日"], errors="coerce")
df_plan["終了日"] = pd.to_datetime(df_plan["終了日"], errors="coerce")

# 欠損除外
df_plan = df_plan.dropna(subset=["畝", "品目", "開始日", "終了日"]).copy()
df_plan["畝"] = df_plan["畝"].astype(str).str.strip()

# 品目マスター適用（品目_base/varを追加）
item_master = load_item_master(RULE_PATH) if RULE_PATH.exists() else pd.DataFrame()
df_plan = apply_item_master(df_plan, item_master)

# ガント表示用のラベル（任意）
df_plan["品目表示"] = df_plan["品目_base"]
df_plan.loc[df_plan["品目_var"].astype(str).str.strip() != "", "品目表示"] = (
    df_plan["品目_base"] + "@" + df_plan["品目_var"]
)

# =========================
# 2) UI（絞り込み）
# =========================
st.title("作付計画 ガントチャート")

with st.sidebar:
    st.header("絞り込み")

    beds_all = sorted(df_plan["畝"].unique().tolist())
    items_all = sorted(df_plan["品目表示"].unique().tolist())

    sel_beds = st.multiselect("畝（複数選択可）", beds_all, default=[])
    sel_items = st.multiselect("品目（複数選択可）", items_all, default=[])

    min_date = df_plan["開始日"].min()
    max_date = df_plan["終了日"].max()

    date_range = st.date_input(
        "表示期間（開始〜終了）",
        value=(min_date.date(), max_date.date()),
        min_value=min_date.date(),
        max_value=max_date.date()
    )

    st.caption("※ 絞り込み後のデータでガントとグループ要約を作ります")

# 絞り込み適用
df_plot = df_plan.copy()

if sel_beds:
    df_plot = df_plot[df_plot["畝"].isin(sel_beds)]

if sel_items:
    df_plot = df_plot[df_plot["品目表示"].isin(sel_items)]

# 日付レンジ（オーバーラップ表示）
start_d = pd.Timestamp(date_range[0])
end_d   = pd.Timestamp(date_range[1])

df_plot = df_plot[~((df_plot["終了日"] < start_d) | (df_plot["開始日"] > end_d))].copy()

if df_plot.empty:
    st.warning("条件に合うデータがありません。絞り込み条件を緩めてください。")
    st.stop()



# =========================
# 3) ガント描画
# =========================

def build_occupancy(df: pd.DataFrame) -> pd.DataFrame:
    """
    df（作付計画の行レベル）から、畝×品目_base（＋品目_var）単位の占有期間に集約した表を作る
    返り値: 畝, 品目表示, 占有開始, 占有終了, 品目_base, 品目_var
    """
    df = df.copy()

    # 念のため
    df["作業"] = df.get("作業", "").astype(str)

    rows = []
    key_cols = ["畝", "品目_base", "品目_var", "品目表示"]

    for (bed, base, var, disp), g in df.groupby(key_cols, dropna=False):
        g = g.sort_values("開始日")

        work = g["作業"].astype(str)

        # 開始（石灰肥料散布 → 畝つくり → 最初）
        m_lime = work.str.contains("石灰肥料散布", na=False)
        m_bed  = work.str.contains("畝つくり", na=False)

        if m_lime.any():
            occ_start = g.loc[m_lime, "開始日"].min()
        elif m_bed.any():
            occ_start = g.loc[m_bed, "開始日"].min()
        else:
            occ_start = g["開始日"].min()

        # 終了（撤収優先）
        m_end = work.str.contains("撤収", na=False)
        if m_end.any():
            occ_end = g.loc[m_end, "終了日"].max()
        else:
            occ_end = g["終了日"].max()

        rows.append({
            "畝": str(bed),
            "品目表示": disp,
            "品目_base": base,
            "品目_var": var,
            "占有開始": occ_start,
            "占有終了": occ_end,
        })

    occ = pd.DataFrame(rows)
    return occ


def assign_sublanes(occ: pd.DataFrame, min_overlap_days: int = 1) -> pd.DataFrame:
    """
    同じ畝内で占有期間が重なる場合、サブレーン番号（1,2,3...）を付与
    """
    occ = occ.copy()
    occ["サブ"] = 1

    def overlaps(a_s, a_e, b_s, b_e) -> bool:
        latest = max(a_s, b_s)
        earliest = min(a_e, b_e)
        # 1日重なりも重複扱い
        return (earliest - latest).days + 1 >= min_overlap_days

    out = []
    for bed, g in occ.groupby("畝"):
        g = g.sort_values("占有開始").reset_index(drop=True)

        # レーンごとに「最後の終了日」を管理
        lane_ends = []  # index=lane-1

        subs = []
        for _, r in g.iterrows():
            placed = False
            for lane_idx, lane_end in enumerate(lane_ends):
                # そのレーンの最後の区間と重ならないなら置ける
                if not overlaps(r["占有開始"], r["占有終了"], r["占有開始"], lane_end):
                    subs.append(lane_idx + 1)
                    lane_ends[lane_idx] = max(lane_end, r["占有終了"])
                    placed = True
                    break

            if not placed:
                lane_ends.append(r["占有終了"])
                subs.append(len(lane_ends))

        g["サブ"] = subs
        out.append(g)

    return pd.concat(out, ignore_index=True)


# ===== ここで占有ガント用データを作る =====
occ = build_occupancy(df_plot)
occ = assign_sublanes(occ, min_overlap_days=1)

# 表示用のY軸（畝-サブ）
occ["畝表示"] = occ["畝"] + "-" + occ["サブ"].astype(int).astype(str)

# 畝の並び（畝順→サブ順）
bed_order = sorted(occ["畝"].unique().tolist())
y_order = []
for b in bed_order:
    subs = sorted(occ.loc[occ["畝"] == b, "サブ"].unique().tolist())
    y_order += [f"{b}-{int(s)}" for s in subs]


def has_overlap(a_start, a_end, b_start, b_end, min_overlap_days=1) -> bool:
    """2区間が min_overlap_days 日以上重なっているか"""
    latest_start = max(a_start, b_start)
    earliest_end = min(a_end, b_end)
    overlap_days = (earliest_end - latest_start).days + 1
    return overlap_days >= min_overlap_days

# --- ガント用：同時混在（期間が重なる）だけ判定 ---
min_overlap_days = 1  # 1日でも重なれば同時混在扱い（必要なら 3 などに）

mixed_beds = set()

for bed, g in df_plot.groupby("畝"):
    gg = g.sort_values("開始日")

    # 品目_baseごとに、その畝での占有期間（最小開始〜最大終了）
    intervals = (
        gg.groupby("品目_base")[["開始日", "終了日"]]
          .agg(開始日=("開始日", "min"), 終了日=("終了日", "max"))
          .reset_index()
    )

    if len(intervals) < 2:
        continue

    iv = intervals.to_dict("records")
    for i in range(len(iv)):
        for j in range(i + 1, len(iv)):
            if has_overlap(
                iv[i]["開始日"], iv[i]["終了日"],
                iv[j]["開始日"], iv[j]["終了日"],
                min_overlap_days=min_overlap_days
            ):
                mixed_beds.add(bed)
                break
        if bed in mixed_beds:
            break

# 畝ラベルをハイライト（同時混在だけ 🟧）
df_plot["畝表示"] = df_plot["畝"].apply(lambda b: f"🟧 {b}" if b in mixed_beds else b)

# 畝の並びは昇順（表示名も合わせる）
bed_order = sorted(df_plot["畝"].unique().tolist())
bed_order_disp = [f"🟧 {b}" if b in mixed_beds else b for b in bed_order]

fig = px.timeline(
    occ,
    x_start="占有開始",
    x_end="占有終了",
    y="畝表示",
    color="品目表示",
    hover_data=["畝", "品目_base", "品目_var", "占有開始", "占有終了", "サブ"],
    category_orders={"畝表示": y_order},
)

fig.update_yaxes(autorange="reversed")

# ===== 区切り線（例：7/1 と 12/1）=====
years = range(
    occ["占有開始"].dt.year.min(),
    occ["占有終了"].dt.year.max() + 1
)

vlines = (
    [pd.Timestamp(y, 7, 1) for y in years] +
    [pd.Timestamp(y, 12, 1) for y in years]
)


# 表示中の範囲に入っている線だけ描く（任意）
xmin = occ["占有開始"].min()
xmax = occ["占有終了"].max()

for d in vlines:
    if xmin <= d <= xmax:
        fig.add_vline(
            x=d,
            line_width=1,
            line_dash="dash",
            opacity=0.6,
        )
        fig.add_annotation(
            x=d, y=1.02, xref="x", yref="paper",
            text=d.strftime("%m/%d"),
            showarrow=False
        )


st.plotly_chart(fig, use_container_width=True)

# =========================
# 高リスク科（アブラナ科／ナス科／マメ科）の占有ガント
# =========================

st.subheader("高リスク科（アブラナ科／ナス科／マメ科）の占有ガント")

target_families = ["アブラナ科", "ナス科", "マメ科"]

# 連作障害から 科マップを読む
try:
    family_map = load_crop_family_map(CROP_RISK_PATH)
except Exception as e:
    st.error(f"連作障害ファイルの読込に失敗: {e}")
    family_map = {}

# ※「全畝」を縦軸に残したいので、畝一覧は layout から取れるならそれが最良
# ここでは簡易に occ から作る（Layout.xlsx を使うならそっちでもOK）
beds_all = sorted(occ["畝"].unique().tolist())

# occ に「科」を付与（品目_baseで引く）
occ_fam = occ.copy()
occ_fam["科"] = occ_fam["品目_base"].map(family_map)

# 対象科だけ残す（NaN は自動的に落ちる）
target_families = ["アブラナ科", "ナス科", "マメ科"]
occ_risk = occ_fam[occ_fam["科"].isin(target_families)].copy()

if occ_risk.empty:
    st.info("対象の科（アブラナ科／ナス科／マメ科）が見つかりませんでした。Crop_Performance の品目_base と作付計画の品目_base を確認してください。")
else:
    # 全畝を軸に残すため、category_orders に全畝を入れる
    fig2 = px.timeline(
        occ_risk,
        x_start="占有開始",
        x_end="占有終了",
        y="畝",
        color="科",
        hover_data=["品目表示", "品目_base", "品目_var", "占有開始", "占有終了"],
        category_orders={"畝": beds_all, "科": target_families},
    )
    fig2.update_yaxes(autorange="reversed")
    fig2.update_layout(
        height=max(500, 22 * len(beds_all) + 200),
        margin=dict(l=20, r=20, t=30, b=20),
        legend_title_text="科",
    )
    st.plotly_chart(fig2, use_container_width=True)

    with st.expander("対象データ（確認用）"):
        st.dataframe(
            occ_risk.sort_values(["畝", "占有開始"]),
            use_container_width=True
        )


# =========================
# 4) ③ 作物グループ要約（ガント補助）★ここが追加ブロック
# =========================
st.subheader("畝の作物グループ要約（ガント補助）")

group_map = load_crop_groups(RULE_PATH) if RULE_PATH.exists() else {}

# ここでは「表示中 df_plot」から畝ごとの品目_baseを集約（= occ2相当）
bed_items = (
    df_plot.groupby("畝")["品目_base"]
          .apply(lambda s: sorted(set([str(x).strip() for x in s.dropna().tolist() if str(x).strip() != ""])))
)

rows = []
for bed, bases in bed_items.items():
    groups = [group_map.get(b, "未分類") for b in bases]
    vc = pd.Series(groups).value_counts()

    cohesion = float(vc.iloc[0] / vc.sum()) if len(vc) else 0.0
    rows.append({
        "畝": bed,
        "品目_base一覧": "、".join(bases),
        "グループ内訳": " / ".join([f"{g}:{n}" for g, n in vc.items()]),
        "中心グループ": vc.index[0] if len(vc) else "",
        "まとまり度": round(cohesion, 2),
        "グループ数": int(vc.size),
        "未分類数": int((pd.Series(groups) == "未分類").sum()) if len(groups) else 0,
    })

df_groupview = pd.DataFrame(rows)

# 表示オプション
col1, col2, col3, col4 = st.columns([1, 1, 1, 2])
with col1:
    only_mixed = st.checkbox("混在（グループ数>=2）の畝だけ", value=True)
with col2:
    only_unclassified = st.checkbox("未分類を含む畝だけ", value=False)
with col3:
    sort_key = st.selectbox("並び順", ["まとまり度（低い順）", "未分類数（多い順）", "畝（昇順）"], index=0)
with col4:
    st.caption("※ 未分類が多いのは「作物グループ未登録の品目」が含まれるためです（必ずしも問題ではありません）")


view = df_groupview.copy()
if only_mixed:
    # 同時混在（期間が重なる畝）だけを表示
    view = view[view["畝"].isin(mixed_beds)].copy()

if only_unclassified:
    view = view[view["未分類数"] >= 1].copy()

if sort_key == "まとまり度（低い順）":
    view = view.sort_values(["まとまり度", "グループ数", "未分類数"], ascending=[True, False, False])
elif sort_key == "未分類数（多い順）":
    view = view.sort_values(["未分類数", "まとまり度"], ascending=[False, True])
else:
    view = view.sort_values(["畝"], ascending=True)

st.dataframe(view, use_container_width=True)

# =========================
# 5)（任意）データ確認用
# =========================
with st.expander("表示中データ（df_plot）を確認"):
    st.dataframe(df_plot.sort_values(["畝", "開始日"]), use_container_width=True)




