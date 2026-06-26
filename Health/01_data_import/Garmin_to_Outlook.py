# %%
#　日にちを指定して、ガーミンアプリからデータ（中程度運動　分、運動消費エネルギーkcal、歩数）を取得し、Outlook 予定表に貼り付ける
# 20251125確認済

# %%
# 期間を指定してデータを取得　2025/05/31確認 20250120以後のデータのみ
import os

import argparse
import datetime

parser = argparse.ArgumentParser()

parser.add_argument("--start-date", required=True)
parser.add_argument("--end-date", required=True)

args = parser.parse_args()

dt_start_input = datetime.datetime.strptime(
    args.start_date, "%Y-%m-%d"
).date()

dt_end_input = datetime.datetime.strptime(
    args.end_date, "%Y-%m-%d"
).date()

print("開始日", dt_start_input)
print("終了日", dt_end_input)


# %%


# %%
# 20251124 12:57  Outlook連携（運動/消費/歩数/心拍 + 前日差つき）
from datetime import datetime, timedelta, time
import os
import re
import win32com.client
import time as pytime

# =========================
# 設定
# =========================
SUBJECT_PREFIX = "[Garmin自動]"   # 今回はCategories判定なので必須ではないが残してOK
DISPLAY_FIX_OFFSET_HOURS = 9     # Outlook表示が-9hズラす前提の補正

email = os.environ.get("GARMIN_EMAIL")
password = os.environ.get("GARMIN_PASSWORD")

# =========================
# Garmin クライアント取得（429 対応版）
# =========================
def get_garmin_client(email: str | None, password: str | None, max_retries=5):
    try:
        from garminconnect import Garmin
    except Exception as e:
        raise RuntimeError(
            "garminconnect が import できません。venv311で `pip install garminconnect` を確認してください。"
        ) from e

    if not email:
        email = os.environ.get("GARMIN_EMAIL")
    if not password:
        password = os.environ.get("GARMIN_PASSWORD")

    if not email or not password:
        raise RuntimeError(
            "Garminログイン情報がありません。\n"
            "環境変数 GARMIN_EMAIL / GARMIN_PASSWORD を設定してください。"
        )

    g = Garmin(email, password)

    # -------------------------
    # 429 対策：ログインをリトライ
    # -------------------------
    for i in range(max_retries):
        try:
            g.login()
            return g  # 成功したら返す

        except Exception as e:
            msg = str(e)

            # 429 の場合
            if "429" in msg or "Too Many Requests" in msg:
                wait = (2 ** i) + 1  # Exponential Backoff
                print(f"Garmin 429: {wait} 秒待機して再試行します ({i+1}/{max_retries})")
                pytime.sleep(wait)
                continue

            # 429 以外のエラーは即終了
            raise

    raise RuntimeError("Garminログインがレート制限で失敗しました。時間を空けて再実行してください。")


# =========================
# 1) Garminログイン（安全版）
# =========================
garmin = get_garmin_client(email, password)


# =========================
# 2) Outlookカレンダー取得（アカウント指定）
# =========================
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
calendar = namespace.Folders["kishi_isao@outlook.jp"].Folders["予定表"]

items = calendar.Items
items.IncludeRecurrences = True
items.Sort("[Start]")

# =========================
# Outlook終日イベント upsert
# =========================
def upsert_all_day_event(day_date, tag, subject_text):
    """
    day_date: datetime.date（その日）
    tag: "運動" / "消費" / "歩数" / "心拍" / "心拍_黄" / "心拍_赤"
    """

    # 0:00ではなく 9:00 をStartにする（表示補正用）
    start_local = datetime.combine(day_date, time(0, 0, 0)) + timedelta(hours=DISPLAY_FIX_OFFSET_HOURS)
    end_local   = start_local + timedelta(days=1)

    # 同日のアイテムを拾う（Restrictは“実日付”で探す）
    day_start_for_search = datetime.combine(day_date, time(0, 0, 0))
    day_end_for_search   = day_start_for_search + timedelta(days=1)

    restriction = (
        "[Start] >= '" + day_start_for_search.strftime("%m/%d/%Y 00:00 AM") + "' AND "
        "[Start] < '"  + day_end_for_search.strftime("%m/%d/%Y 00:00 AM") + "'"
    )
    day_items = items.Restrict(restriction)

    # 既存イベントを探す（Categoriesで判定）
    target = None
    for it in day_items:
        cats = str(it.Categories or "")

        # 心拍は "心拍" 系（心拍/心拍_黄/心拍_赤）を探す
        if tag.startswith("心拍"):
            if cats.startswith("心拍"):
                target = it
                break
        else:
            # 運動/消費/歩数 は Garmin自動;タグ のもの
            if ("Garmin自動" in cats) and (tag in cats):
                target = it
                break

    if target is None:
        target = calendar.Items.Add()

    # 絶対値で1日終日に固定
    target.AllDayEvent = True
    target.Start = start_local
    target.End   = end_local
    target.Duration = 1440
    target.BusyStatus = 0
    target.ReminderSet = False

    # Categories 設定
    if tag.startswith("心拍"):
        # 例: "心拍" / "心拍_黄" / "心拍_赤"
        target.Categories = tag
    else:
        target.Categories = f"Garmin自動;{tag}"

    target.Subject = subject_text
    target.Body = ""
    target.Save()


# =========================
# 前日のRHRをOutlookから取得
# =========================
def get_prev_rhr_from_outlook(prev_date):
    """
    前日prev_dateの心拍イベント（Categoriesが心拍で始まる）から Subject の
    'RHR 46...' の 46 を抜き出す。無ければNone。
    """
    day_start = datetime.combine(prev_date, time(0, 0, 0))
    day_end   = day_start + timedelta(days=1)

    restriction = (
        "[Start] >= '" + day_start.strftime("%m/%d/%Y 00:00 AM") + "' AND "
        "[Start] < '"  + day_end.strftime("%m/%d/%Y 00:00 AM") + "'"
    )
    day_items = items.Restrict(restriction)

    found = None
    for it in day_items:
        cats = str(it.Categories or "")
        if cats.startswith("心拍"):
            subj = str(it.Subject or "")
            m = re.search(r"RHR\s*(\d+)", subj)
            if m:
                found = int(m.group(1))
                # 複数あっても最後に見つかったものを採用（重複耐性）
    return found


# =========================
# 3) 指定期間のGarminデータをOutlook終日イベントに反映
# =========================
# START_DATE, END_DATE は前のセルで date 型として入力している前提
start_date = datetime.combine(dt_start_input, datetime.min.time())
end_date   = datetime.combine(dt_end_input,   datetime.min.time())

current = start_date
while current <= end_date:
    date_str = current.strftime("%Y-%m-%d")
    day_date = current.date()

    # =========================================================
    # 1) Garmin 日次データ（歩数/消費/運動分）
    # =========================================================
    stats = garmin.get_stats(date_str) or {}

    steps = stats.get("totalSteps", 0) or 0
    calories = stats.get("activeKilocalories", 0) or 0
    moderate_minutes = stats.get("moderateIntensityMinutes") or 0
    vigorous_minutes = stats.get("vigorousIntensityMinutes") or 0
    exercise_minutes = moderate_minutes + vigorous_minutes

    exercise_minutes_i = int(exercise_minutes or 0)
    calories_i = int(calories or 0)
    steps_i = int(steps or 0)

    ex_str = f"{exercise_minutes_i:03d}"
    cal_str = f"{calories_i:04d}"
    steps_str = f"{steps_i:05d}"

    # =========================================================
    # 2) Garmin 心拍データ（RHR/MAX/MIN/RHR7）
    # =========================================================
    try:
        hr = garmin.get_heart_rates(date_str) or {}
    except Exception as e:
        print(f"⚠️ {date_str} 心拍取得に失敗: {e}")
        hr = {}

    rhr = int(hr.get("restingHeartRate") or 0)
    hr_max = int(hr.get("maxHeartRate") or 0)
    hr_min = int(hr.get("minHeartRate") or 0)
    rhr7 = int(hr.get("lastSevenDaysAvgRestingHeartRate") or 0)

    # =========================================================
    # 3) アラート判定（黄/赤/通常）
    # =========================================================
    delta = rhr - rhr7  # 今日RHRが7日平均よりどれだけ高いか

    Y_RHR_DELTA = 8
    R_RHR_DELTA = 12
    Y_MAX = 120
    R_MAX = 135

    level = 0
    if rhr7 > 0:
        if delta >= R_RHR_DELTA:
            level = 2
        elif delta >= Y_RHR_DELTA:
            level = 1

    if hr_max >= R_MAX:
        level = 2
    elif hr_max >= Y_MAX and level < 2:
        level = 1

    if level == 2:
        category_tag = "心拍_赤"
    elif level == 1:
        category_tag = "心拍_黄"
    else:
        category_tag = "心拍"

    # =========================================================
    # 4) 前日差（Outlookの前日イベントから）
    # =========================================================
    prev_rhr = get_prev_rhr_from_outlook(day_date - timedelta(days=1))

    if (prev_rhr is None) or (rhr == 0):
        diff_txt = ""
    else:
        diff_prev = rhr - prev_rhr
        if diff_prev > 0:
            diff_txt = f"(+{diff_prev})"
        elif diff_prev < 0:
            diff_txt = f"({diff_prev})"
        else:
            diff_txt = "(0)"

    hr_all = f"RHR {rhr}{diff_txt} | MAX {hr_max} | MIN {hr_min} | 7d {rhr7}"

    # =========================================================
    # 5) Outlook 終日イベント反映（upsert）
    # =========================================================
    upsert_all_day_event(day_date, "運動", f"{ex_str}分")
    upsert_all_day_event(day_date, "消費", f"運動消費{cal_str}kcal")
    upsert_all_day_event(day_date, "歩数", f"歩数{steps_str}")
    upsert_all_day_event(day_date, category_tag, hr_all)

    print(
        f"✅ {date_str} 反映："
        f"運動{exercise_minutes_i}分 / {calories_i}kcal / {steps_i}歩 / "
        f"{hr_all}"
    )

    current += timedelta(days=1)

print("=== 完了しました ===")



