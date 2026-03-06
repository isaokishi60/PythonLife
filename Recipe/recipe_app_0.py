import re
import os
from pathlib import Path
from urllib.parse import urlsplit, urlunsplit

import pandas as pd
import streamlit as st

# =========================
# パス（OneDrive配下）
# =========================
BASE = Path(os.environ.get("OneDrive", "")) / "タブレット用"
DB_PATH = BASE / "recipe_db.xlsx"
SHOTS_ROOT = BASE / "RecipeShots"
SHEET_NAME = "Recipes"

# =========================
# ユーティリティ
# =========================
def norm(x) -> str:
    if x is None:
        return ""
    s = str(x).strip().lower()
    s = re.sub(r"\s+", " ", s)
    return s

def norm_rating(x) -> str:
    if x is None:
        return ""
    try:
        if pd.isna(x):
            return ""
    except Exception:
        pass
    return str(x).strip().upper()

def is_blank(x) -> bool:
    if x is None:
        return True
    try:
        if pd.isna(x):
            return True
    except Exception:
        pass
    return str(x).strip() == ""

def safe_str(x) -> str:
    if is_blank(x):
        return ""
    return str(x).strip()

def list_images(folder: Path):
    exts = ["*.png", "*.jpg", "*.jpeg", "*.webp"]
    files = []
    for pat in exts:
        files.extend(folder.glob(pat))
    return sorted(files)

def clean_url_for_mobile(u: str) -> str:
    u = (u or "").strip()
    if not u:
        return ""
    sp = urlsplit(u)
    # query（?以降）と fragment（#以降）を両方落とす
    return urlunsplit((sp.scheme, sp.netloc, sp.path, "", ""))

def make_label(row: pd.Series) -> str:
    rid = safe_str(row.get("RecipeID", ""))
    rate = norm_rating(row.get("評価", ""))
    name = safe_str(row.get("レシピ名", ""))
    rate_disp = rate if rate else "-"
    return f"{rid} | {rate_disp} | {name}"

# =========================
# データ読み込み
# =========================
@st.cache_data
def load_recipes():
    if not DB_PATH.exists():
        raise FileNotFoundError(f"DBが見つかりません: {DB_PATH}")

    df = pd.read_excel(DB_PATH, sheet_name=SHEET_NAME)

    need = ["RecipeID", "レシピ名", "主材料1", "主材料2", "主材料3"]
    miss = [c for c in need if c not in df.columns]
    if miss:
        raise ValueError(f"{SHEET_NAME}に必要列がありません: {miss}")

    # 任意列
    for c in ["評価", "コメント", "URL"]:
        if c not in df.columns:
            df[c] = ""

    df["RecipeID"] = df["RecipeID"].astype(str).str.strip()
    return df

def filter_recipes(df: pd.DataFrame, kw: str, m1: str, rating_opt: str) -> pd.DataFrame:
    f = pd.Series(True, index=df.index)

    # 評価フィルター（案2）
    rr = df["評価"].map(norm_rating)
    if rating_opt == "選択しない":
        pass
    elif rating_opt == "全て":
        f &= rr.isin(["A", "B", "C"])
    else:
        f &= rr.eq(rating_opt)

    # レシピ名（部分一致）
    if kw:
        f &= df["レシピ名"].astype(str).map(norm).str.contains(re.escape(norm(kw)), na=False)

    # 主材料（部分一致）
    if m1:
        v = re.escape(norm(m1))
        f &= (
            df["主材料1"].astype(str).map(norm).str.contains(v, na=False)
            | df["主材料2"].astype(str).map(norm).str.contains(v, na=False)
            | df["主材料3"].astype(str).map(norm).str.contains(v, na=False)
        )

    return df[f].copy()

# =========================
# UI（ページ管理）
# =========================
st.set_page_config(page_title="レシピ検索", layout="wide")
st.title("レシピ検索（材料表＋調理法）")

with st.expander("現在のDB/画像フォルダ（確認）", expanded=False):
    st.write(f"DB: {DB_PATH}")
    st.write(f"Shots root: {SHOTS_ROOT}")

df = load_recipes().copy()

# --- session_state 初期化 ---
if "page" not in st.session_state:
    st.session_state.page = "search"   # search / select / view
if "kw" not in st.session_state:
    st.session_state.kw = ""
if "m1" not in st.session_state:
    st.session_state.m1 = ""
if "rating_opt" not in st.session_state:
    st.session_state.rating_opt = "選択しない"
if "selected" not in st.session_state:
    st.session_state.selected = []

show_cols = ["RecipeID", "レシピ名", "主材料1", "主材料2", "主材料3", "評価"]

# =========================
# ① 検索ページ
# =========================
if st.session_state.page == "search":
    st.header("① 検索条件")

    kw = st.text_input("レシピ名（部分一致）", value=st.session_state.kw)
    m1 = st.text_input("主材料（任意）", value=st.session_state.m1)
    rating_opt = st.selectbox(
        "評価（作ったレシピ）",
        ["選択しない", "全て", "A", "B", "C"],
        index=["選択しない", "全て", "A", "B", "C"].index(st.session_state.rating_opt)
        if st.session_state.rating_opt in ["選択しない", "全て", "A", "B", "C"] else 0
    )

    col1, col2 = st.columns(2)
    with col1:
        if st.button("検索する", use_container_width=True):
            st.session_state.kw = kw
            st.session_state.m1 = m1
            st.session_state.rating_opt = rating_opt
            st.session_state.selected = []
            st.session_state.page = "select"
            st.rerun()

    with col2:
        if st.button("条件クリア", use_container_width=True):
            st.session_state.kw = ""
            st.session_state.m1 = ""
            st.session_state.rating_opt = "選択しない"
            st.session_state.selected = []
            st.rerun()

# =========================
# ② 候補選択ページ（ここではキーボード不要）
# =========================
elif st.session_state.page == "select":
    st.header("② 候補レシピ（タップで選択）")

    res = filter_recipes(df, st.session_state.kw, st.session_state.m1, st.session_state.rating_opt)
    st.caption(f"該当: {len(res)}件")

    if len(res) == 0:
        st.info("条件に一致するレシピがありません。")
        if st.button("← 検索条件に戻る"):
            st.session_state.page = "search"
            st.rerun()
        st.stop()

    st.dataframe(res[show_cols], hide_index=True, use_container_width=True)

    options = res["RecipeID"].astype(str).tolist()
    label_map = {rid: make_label(res[res["RecipeID"] == rid].iloc[0]) for rid in options}

    selected = st.multiselect(
        "表示するレシピ（複数選択可）",
        options=options,
        default=[x for x in st.session_state.selected if x in options],
        format_func=lambda x: label_map.get(x, x),
    )

    c1, c2 = st.columns(2)
    with c1:
        if st.button("← 検索条件に戻る", use_container_width=True):
            st.session_state.page = "search"
            st.rerun()

    with c2:
        if st.button("レシピを見る", use_container_width=True):
            st.session_state.selected = selected
            st.session_state.page = "view"
            st.rerun()

# =========================
# ③ 表示ページ
# =========================
elif st.session_state.page == "view":
    st.header("③ レシピ表示")

    # いまの条件で再計算（DB更新があっても整合する）
    res = filter_recipes(df, st.session_state.kw, st.session_state.m1, st.session_state.rating_opt)

    selected = [x for x in st.session_state.selected if x in res["RecipeID"].astype(str).tolist()]

    if len(selected) == 0:
        st.info("表示するレシピが選択されていません。")
        if st.button("← 候補一覧に戻る"):
            st.session_state.page = "select"
            st.rerun()
        st.stop()

    # 戻るボタン（上）
    if st.button("← 候補一覧に戻る"):
        st.session_state.page = "select"
        st.rerun()

    # 詳細表示（複数）
    for rid in selected:
        row = res[res["RecipeID"].astype(str) == str(rid)].iloc[0]

        st.divider()
        st.subheader(f"{safe_str(row['レシピ名'])}（{rid}）")
        st.write(
            f"主材料: {safe_str(row.get('主材料1',''))} / "
            f"{safe_str(row.get('主材料2',''))} / "
            f"{safe_str(row.get('主材料3',''))}"
        )
        st.write(f"評価: {norm_rating(row.get('評価','')) or '（未入力）'}")

        # URL
        url_raw = safe_str(row.get("URL", ""))
        url = clean_url_for_mobile(url_raw)
        if url:
            if "kurashiru.com" in url:
                st.caption("※ クラシルはタブレットでは会員/誘導表示になる場合があります（画像が正本です）。")
            st.link_button("参考URLを開く", url)
        else:
            st.write("URL: URLはない")

        # コメント
        comment = safe_str(row.get("コメント", ""))
        if comment:
            st.caption(f"コメント: {comment}")

        st.divider()

        # 画像
        rdir = SHOTS_ROOT / str(rid)
        if not rdir.exists():
            st.warning(f"画像フォルダが見つかりません: {rdir}")
            continue

        imgs = list_images(rdir)
        if not imgs:
            st.warning(f"{rdir} に画像がありません。")
            continue

        st.info("上から順に表示します。最初を材料、続きが調理法になるように並び（ファイル名）を揃えると便利です。")
        for p in imgs:
            st.image(str(p), caption=p.name, use_container_width=True)

    # 戻るボタン（下）
    st.divider()
    if st.button("← 候補一覧に戻る（上に戻らず）", use_container_width=True):
        st.session_state.page = "select"
        st.rerun()








