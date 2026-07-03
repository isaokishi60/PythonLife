# %%
# ============================================================
# Top10履歴を「株価実績ファイル」から一括生成（Top10専用・全面差し替え版）
# 出力: 資産推移_Top10History.xlsx / sheet: Top10History
# ============================================================

import os
from pathlib import Path
import re
import pandas as pd

# =========================
# 設定（OneDriveベース）
# =========================
ONEDRIVE = os.environ.get("OneDrive")
if not ONEDRIVE:
    raise EnvironmentError("OneDriveの環境変数が見つかりません。OneDriveが有効か確認してください。")

BASE = Path(ONEDRIVE) / "有価証券"

PRICE_HISTORY_DIR = BASE / "02_価格データ" / "株価" / "株価実績"
HOLDINGS_PATH     = BASE / "01_Excel入力" / "残高" / "保有総額_指定日基準.xlsx"  # 数量/取得額/銘柄/コード
OUT_PATH          = BASE / "01_Excel入力" / "残高" / "資産推移_Top10History.xlsx"
TOP10_SHEET       = "Top10History"

# =========================
# ユーティリティ
# =========================
def _to_num(x: pd.Series) -> pd.Series:
    """カンマ除去して数値化（空/文字も許容）"""
    return pd.to_numeric(
        x.astype(str).str.replace(",", "", regex=False),
        errors="coerce"
    )

def _make_security_code_from_code(code_series: pd.Series) -> pd.Series:
    """
    SecurityCodeを作る：
    - 株式: 4桁など -> そのまま
    - 投信: 5〜8桁の数字 -> 8桁ゼロ埋め（03315177 など）
    """
    s = code_series.astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    s = s.where(~s.isin(["", "nan", "None"]), pd.NA)

    is_digits_1to8 = s.fillna("").str.fullmatch(r"\d{1,8}", na=False)
    is_fund = is_digits_1to8 & (s.fillna("").str.len() > 4)  # 5〜8桁は投信扱い
    s = s.where(~is_fund, s.fillna("").str.zfill(8))
    return s

def _list_price_files() -> list[Path]:
    files = sorted(PRICE_HISTORY_DIR.glob("株価_*_前営業日終値.xlsx"))
    if not files:
        raise FileNotFoundError(f"株価実績ファイルが見つかりません: {PRICE_HISTORY_DIR}")
    return files

def _read_prices(fp: Path) -> pd.DataFrame:
    """Pricesシートを読み、必要列だけ返す（列名ゆらぎに強く）"""
    dfp = pd.read_excel(fp, sheet_name="Prices")
    dfp.columns = dfp.columns.astype(str).str.strip()

    need = ["SecurityID", "終値", "前営業日"]
    miss = [c for c in need if c not in dfp.columns]
    if miss:
        raise KeyError(f"Prices列不足: {miss}  file={fp.name}")

    dfp["SecurityID"] = dfp["SecurityID"].astype(str).str.strip()
    dfp["終値"] = _to_num(dfp["終値"])

    # 前営業日（基準日）はファイル内で同一の想定
    base_dt = pd.to_datetime(dfp["前営業日"].dropna().iloc[0]).date()

    out = dfp[["SecurityID", "終値"]].copy()
    out.rename(columns={"終値": "終値_px"}, inplace=True)
    out["base_dt"] = base_dt
    return out

# =========================
# 0) 保有（いまの数量/取得額）を読み込む
# =========================
hold = pd.read_excel(HOLDINGS_PATH)
hold.columns = hold.columns.astype(str).str.strip()

need_hold = ["SecurityID", "銘柄", "コード", "数量", "取得額"]
missing_hold = [c for c in need_hold if c not in hold.columns]
if missing_hold:
    raise KeyError(f"保有ファイルに必要列がありません: {missing_hold}\ncolumns={hold.columns.tolist()}")

hold["SecurityID"] = hold["SecurityID"].astype(str).str.strip()
hold["Name"] = hold["銘柄"].astype(str).str.strip()
hold["SecurityCode"] = _make_security_code_from_code(hold["コード"])

hold["数量"] = _to_num(hold["数量"])
hold["取得額"] = _to_num(hold["取得額"])

# （保険）SecurityCodeが欠けるとTop10表が壊れるので警告
if hold["SecurityCode"].isna().any():
    bad = hold.loc[hold["SecurityCode"].isna(), ["SecurityID", "Name", "コード"]].head(10)
    print("WARNING: SecurityCode を作れない行があります（先頭10件）")
    print(bad.to_string(index=False))

# =========================
# 1) 株価実績ファイルを回して Top10 を作る（完全作り直し）
# =========================
price_files = _list_price_files()
print(f"価格ファイル数: {len(price_files)}")

rows_all = []  # 最終Top10Historyを積む（完全生成）

for fp in price_files:
    try:
        px = _read_prices(fp)
    except Exception as e:
        print(f"SKIP: {fp.name} ({e})")
        continue

    base_dt = px["base_dt"].iloc[0]

    # 保有×価格
    m = hold.merge(px[["SecurityID", "終値_px"]], on="SecurityID", how="left")

    m["評価額"] = (m["数量"] * m["終値_px"]).round(0)
    m["利益"]   = (m["評価額"] - m["取得額"]).round(0)

    # 価格が無い銘柄は除外（Top10計算不能）
    m = m.dropna(subset=["評価額"])

    # 銘柄単位に集約（同一銘柄を複数口座で持つケースに対応）
    agg = (
        m.groupby(["SecurityCode", "Name"], as_index=False)[["評価額", "利益"]]
         .sum()
    )

    top10 = (
        agg.sort_values("評価額", ascending=False)
           .head(10)
           .copy()
    )

    if len(top10) == 0:
        print(f"WARNING: Top10が作れません: {fp.name} (価格欠損が多い可能性)")
        continue

    # 行追加（この日付ぶん）
    tmp = pd.DataFrame({
        "Date":         [base_dt] * len(top10),
        "SecurityCode": top10["SecurityCode"].astype(str).str.strip().tolist(),
        "Name":         top10["Name"].astype(str).str.strip().tolist(),
        "Value":        pd.to_numeric(top10["評価額"], errors="coerce").round(0),
        "Profit":       pd.to_numeric(top10["利益"], errors="coerce").round(0),
        "Reference":    [fp.name] * len(top10),
    })
    rows_all.append(tmp)

# =========================
# 2) まとめて整形
# =========================
if not rows_all:
    raise RuntimeError("Top10History を生成できませんでした（価格ファイルが全てSKIPされた可能性）")

hist10 = pd.concat(rows_all, ignore_index=True)

# 型・並び
hist10["Date"] = pd.to_datetime(hist10["Date"]).dt.date
hist10["Value"]  = pd.to_numeric(hist10["Value"],  errors="coerce").astype("Int64")
hist10["Profit"] = pd.to_numeric(hist10["Profit"], errors="coerce").astype("Int64")

# DateごとにValue降順（見やすく）
hist10 = hist10.sort_values(["Date", "Value"], ascending=[True, False]).reset_index(drop=True)

# 重複防止（同じDate×SecurityCodeが重複したら最後を採用）
hist10 = hist10.drop_duplicates(subset=["Date", "SecurityCode"], keep="last")

# 列順固定
cols = ["Date", "SecurityCode", "Name", "Value", "Profit", "Reference"]
hist10 = hist10.reindex(columns=cols)

# =========================
# 3) 保存（Top10History だけを書き込む）
# =========================
OUT_PATH.parent.mkdir(parents=True, exist_ok=True)

with pd.ExcelWriter(OUT_PATH, engine="openpyxl", mode="w") as w:
    hist10.to_excel(w, sheet_name=TOP10_SHEET, index=False)

print("Top10History を一括生成しました:", str(OUT_PATH))
print("Sheet:", TOP10_SHEET)
print("last 10 rows:")
print(hist10.tail(10).to_string(index=False))



