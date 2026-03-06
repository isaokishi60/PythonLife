# make_price_file.py
# 株価（株式= yfinance / 投信= Yahoo Japan 履歴HTML）を取得して
# 株価_YYYYMMDD_前営業日終値.xlsx を出力する（Pricesシート）

import os
import re
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
import requests
import yfinance as yf


# =========================
# パス設定（OneDrive優先）
# =========================
def resolve_base() -> Path:
    od = os.environ.get("OneDrive")
    if od:
        p = Path(od) / "有価証券"
        if p.exists():
            return p

    # フォールバック（OneDrive環境変数が無い等）
    p2 = Path(r"O:\有価証券")
    if p2.exists():
        print("WARNING: OneDrive が見つからないため O:\\有価証券 を使用します")
        return p2

    raise EnvironmentError("有価証券フォルダが見つかりません。OneDrive または O: を確認してください。")


BASE = resolve_base()

MASTER_PATH = BASE / r"01_Excel入力\マスター\有価証券マスター.xlsx"
MASTER_SHEET = "Securities"

OUT_DIR = BASE / r"02_価格データ\株価\株価実績"  # 出力先


# =========================
# マスター読込
# =========================
def load_securities_master() -> pd.DataFrame:
    df = pd.read_excel(MASTER_PATH, sheet_name=MASTER_SHEET)
    df.columns = df.columns.astype(str).str.strip()

    if "SecurityID" not in df.columns:
        raise ValueError("有価証券マスターに 'SecurityID' 列がありません")

    out = pd.DataFrame()
    out["SecurityID"] = df["SecurityID"].astype(str).str.strip()

    # 種類 / 名称
    out["種類"] = df["種類"].astype(str).str.strip() if "種類" in df.columns else ""
    out["名称"] = df["名称"].astype(str).str.strip() if "名称" in df.columns else ""

    # 価格取得対象（無ければ True 扱い）
    if "価格取得対象" in df.columns:
        out["価格取得対象"] = df["価格取得対象"].fillna(False).astype(bool)
    else:
        out["価格取得対象"] = True

    # 株式コード（例 7203）
    if "コード" in df.columns:
        out["コード"] = (
            df["コード"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
        )
    else:
        out["コード"] = pd.NA
    out.loc[out["コード"].isin(["", "nan", "None"]), "コード"] = pd.NA

    # 投信コード（別名対応）
    fund_col = None
    for c in ["投資信託コード", "投信協会コード"]:
        if c in df.columns:
            fund_col = c
            break

    if fund_col:
        out["投資信託コード"] = (
            df[fund_col].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
        )
    else:
        out["投資信託コード"] = pd.NA

    out.loc[out["投資信託コード"].isin(["", "nan", "None"]), "投資信託コード"] = pd.NA

    # ★投資信託コードは8桁ゼロ埋め
    out["投資信託コード"] = out["投資信託コード"].astype("string")
    m = out["投資信託コード"].notna()
    out.loc[m, "投資信託コード"] = out.loc[m, "投資信託コード"].str.zfill(8)

    return out


# =========================
# ユーティリティ
# =========================
def normalize_code(x) -> str:
    """Excel由来の 1234.0 対策 + 前後空白除去"""
    if pd.isna(x):
        return ""
    s = str(x).strip()
    s = re.sub(r"\.0$", "", s)
    return s


def ask_target_date() -> date:
    print("日付を入力してください（この日の前営業日終値を取得します）")
    y = int(input("開始年4桁: ").strip())
    m = int(input("開始月2桁: ").strip())
    d = int(input("開始日2桁: ").strip())
    td = date(y, m, d)
    print("指定日:", td)
    return td


# =========================
# 株式：yfinance（頑丈版）
# =========================
def last_close_before_yfinance(symbol: str, target: date, days_back: int = 20):
    """
    target日の「前営業日終値」を返す（株式用：yfinance）
    """
    start = pd.Timestamp(target - timedelta(days=days_back))
    end = pd.Timestamp(target + timedelta(days=1))

    try:
        hist = yf.download(
            symbol,
            start=start,
            end=end,
            progress=False,
            auto_adjust=False,
            group_by="column",
        )
    except Exception as e:
        return None, None, f"download error: {e}"

    if hist is None or len(hist) == 0:
        return None, None, "no history"

    # yfinanceがMultiIndex列を返す場合あり
    if isinstance(hist.columns, pd.MultiIndex):
        hist.columns = hist.columns.get_level_values(0)

    # 日付は index に入っているので reset_index
    hist = hist.reset_index()

    # dtcol の同定（Dateが無いケース対策）
    dtcol = None
    for c in ["Date", "Datetime", "index"]:
        if c in hist.columns:
            dtcol = c
            break
    if dtcol is None:
        # 最初の列を日付扱い（最後の保険）
        dtcol = hist.columns[0]

    if "Close" not in hist.columns:
        return None, None, f"no Close column. columns={hist.columns.tolist()}"

    hist[dtcol] = pd.to_datetime(hist[dtcol], errors="coerce")
    hist = hist.dropna(subset=[dtcol]).copy()
    hist["__dt"] = hist[dtcol].dt.date

    hist = hist[hist["__dt"] < target].copy()
    if len(hist) == 0:
        return None, None, "no rows before target"

    row = hist.iloc[-1]
    px = row["Close"]
    px_date = row["__dt"]
    return float(px), px_date, None


# =========================
# 投資信託：Yahoo Japan（履歴HTML）
# =========================
def last_nav_before_yahoo_fund(fund_code_8: str, target: date, timeout: int = 20):
    """
    Yahooファイナンス（日本）の投信履歴ページから target日の直近（<=target）基準価額を取得。
    URL例: https://finance.yahoo.co.jp/quote/03315177/history
    """
    fund_code_8 = str(fund_code_8).strip().zfill(8)

    url = f"https://finance.yahoo.co.jp/quote/{fund_code_8}/history"
    headers = {"User-Agent": "Mozilla/5.0"}

    try:
        r = requests.get(url, headers=headers, timeout=timeout)
    except Exception as e:
        return None, None, f"request error: {e}"

    if r.status_code != 200:
        return None, None, f"HTTP {r.status_code}"

    try:
        tables = pd.read_html(r.text)
    except Exception as e:
        return None, None, f"read_html failed: {e}"

    if not tables:
        return None, None, "no tables"

    # それっぽい表を探す
    cand = None
    for t in tables:
        cols = [str(c).strip() for c in t.columns]
        if any(("日付" in c) or ("基準日" in c) for c in cols) and any(
            (k in c) for k in ["基準価額", "終値", "基準値", "値"] for c in cols
        ):
            cand = t.copy()
            break
    if cand is None:
        cand = tables[0].copy()

    cand.columns = [str(c).strip() for c in cand.columns]

    # 日付列
    date_col = None
    for c in cand.columns:
        if ("日付" in c) or ("基準日" in c):
            date_col = c
            break
    if date_col is None:
        return None, None, f"date col not found. columns={cand.columns.tolist()}"

    # 価格列
    price_col = None
    for key in ["基準価額", "終値", "基準値", "値"]:
        for c in cand.columns:
            if key in c:
                price_col = c
                break
        if price_col is not None:
            break
    if price_col is None:
        return None, None, f"price col not found. columns={cand.columns.tolist()}"

    # 日付・価格の型変換（強化版）
    s = cand[date_col].astype(str).str.strip()
    s = (
        s.str.replace("年", "-", regex=False)
        .str.replace("月", "-", regex=False)
        .str.replace("日", "", regex=False)
        .str.replace("/", "-", regex=False)
    )
    cand[date_col] = pd.to_datetime(s, errors="coerce").dt.normalize()

    cand[price_col] = cand[price_col].astype(str).str.replace(",", "", regex=False)
    cand[price_col] = pd.to_numeric(cand[price_col], errors="coerce")

    cand = cand.dropna(subset=[date_col, price_col]).copy()

    target_ts = pd.Timestamp(target).normalize()
    cand = cand[cand[date_col] <= target_ts].sort_values(date_col)

    if len(cand) == 0:
        return None, None, f"no rows <= target ({target_ts.date()})"

    row = cand.iloc[-1]
    px_date = row[date_col].date()
    px = float(row[price_col])

    return px, px_date, None


# =========================
# メイン
# =========================
def main():
    target = ask_target_date()

    OUT_DIR.mkdir(parents=True, exist_ok=True)

    try:
        sec = load_securities_master()
    except PermissionError:
        raise PermissionError(
            f"PermissionError: {MASTER_PATH}\n"
            "有価証券マスター.xlsx をExcelで開いていませんか？閉じてから再実行してください。"
        )

    # 取得対象だけ
    sec = sec[sec["価格取得対象"] == True].copy()

    print("SEC MASTER COLUMNS:", sec.columns.tolist())
    print(sec.head(10))

    # 投信コード確認
    print("USING SEC MASTER (fund codes should be 8 digits):")
    if "投資信託コード" in sec.columns:
        print(sec.loc[sec["種類"] != "株式", ["SecurityID", "種類", "投資信託コード"]].head(30).to_string(index=False))

    rows = []

    for _, r in sec.iterrows():
        secid = str(r["SecurityID"]).strip()
        kind = str(r["種類"]).strip()
        name = str(r["名称"]).strip()

        # --- 株式 ---
        if kind == "株式":
            code = normalize_code(r.get("コード", ""))
            if code == "":
                rows.append({
                    "SecurityID": secid, "種類": kind, "名称": name,
                    "コード": pd.NA, "前営業日": pd.NA, "終値": pd.NA,
                    "取得元": "", "エラー": "stock code empty"
                })
                continue

            symbol = f"{code}.T"
            px, px_date, err = last_close_before_yfinance(symbol, target)
            used = symbol

        # --- 投資信託（その他） ---
        else:
            fund = normalize_code(r.get("投資信託コード", ""))
            fund = fund.zfill(8) if fund else ""
            if fund == "":
                rows.append({
                    "SecurityID": secid, "種類": kind, "名称": name,
                    "コード": pd.NA, "前営業日": pd.NA, "終値": pd.NA,
                    "取得元": "", "エラー": "fund code empty"
                })
                continue

            px, px_date, err = last_nav_before_yahoo_fund(fund, target)
            used = f"YahooJP:{fund}"
            code = fund  # コード列は統一でfundコードを入れる

        if err is None and px is not None and px_date is not None:
            rows.append({
                "SecurityID": secid,
                "種類": kind,
                "名称": name,
                "コード": code,
                "前営業日": px_date,
                "終値": px,
                "取得元": used,
                "エラー": ""
            })
        else:
            rows.append({
                "SecurityID": secid,
                "種類": kind,
                "名称": name,
                "コード": code if kind != "株式" else code,
                "前営業日": pd.NA,
                "終値": pd.NA,
                "取得元": used,
                "エラー": f"{used}: {err}"
            })

    out = pd.DataFrame(rows)

    ymd = target.strftime("%Y%m%d")
    out_path = OUT_DIR / f"株価_{ymd}_前営業日終値.xlsx"

    with pd.ExcelWriter(out_path, engine="openpyxl") as w:
        out.to_excel(w, sheet_name="Prices", index=False)

    print("出力完了:", str(out_path))

    ng = out[out["エラー"].astype(str).str.len() > 0]
    if len(ng) > 0:
        print("取得できなかった行があります（上位10件）:")
        print(ng[["SecurityID", "種類", "名称", "コード", "エラー"]].head(10).to_string(index=False))


if __name__ == "__main__":
    main()