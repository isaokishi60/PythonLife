import pandas as pd
from datetime import date

EXCEL_PATH = r"C:\Users\kawgu\OneDrive\ドキュメント\USB Backup\PythonWork\農作業\vegetable_garden_photo_ex2_with_bed.xlsx"


# ===== 2) 既存ファイルの読み込み or 新規作成 =====
base_columns = [
    "Date", "Name or Item", "畝", "区画",
    "Location", "Photo_id", "JPG_Photo",
    "Tag1", "Tag2", "Tag3", "Tag4", "Tag5"
]

try:
    df = pd.read_excel(EXCEL_PATH)
    print("既存ファイルを読み込みました。行数:", len(df))

    # 必要な列がなければ追加しておく（空列）
    for col in base_columns:
        if col not in df.columns:
            df[col] = None
            print(f"列 '{col}' がなかったので追加しました。")

except FileNotFoundError:
    print("ファイルが見つからなかったので、新規作成します。")
    df = pd.DataFrame(columns=base_columns)


def input_one_session():
    """1回分の入力を受け取り、追加する行のリストを返す"""
    print("\n=== 新しい作業記録を追加します ===")

    # 日付（Enterだけなら今日）
    today_str = date.today().isoformat()
    d_str = input(f"日付 (YYYY-MM-DD) [{today_str}]: ").strip()
    if not d_str:
        d_str = today_str

    item = input("品目 (Name or Item): ").strip()
    if not item:
        print("品目は必須です。キャンセルします。")
        return []

    # 畝入力：例) A02-南, A05-中  のようにカンマ区切り
    beds_str = input("畝（区画込み、例: A02-南, A05-中。畝1つなら A02-南）: ").strip()
    if not beds_str:
        print("畝が未入力です。キャンセルします。")
        return []

    # ★ 全角の読点・全角カンマにも対応
    beds_str = beds_str.replace("、", ",").replace("，", ",")


    # Photo_id / JPG_Photo は任意
    photo_id = input("Photo_id（任意）: ").strip()
    jpg_photo = input("JPG_Photo（写真ファイル名やパス、任意）: ").strip()

    # タグ（最大5個）
    tags_input = input("タグをカンマ区切りで入力（最大5個、例: 播種,寒冷紗）: ").strip()

    # ★ 全角の「、」「，」を半角カンマに揃える（あれば）
    tags_input = tags_input.replace("、", ",").replace("，", ",")

    tags = [t.strip() for t in tags_input.split(",") if t.strip()] if tags_input else []
    tags = (tags + [""] * 5)[:5]


    new_rows = []

    # ===== 畝を分解して複数行にする =====
    for bed_part in beds_str.split(","):
        bed_part = bed_part.strip()
        if not bed_part:
            continue

        # 例: "A02-南" → bed="A02", block="南"
        if "-" in bed_part:
            bed, block = bed_part.split("-", 1)
            bed = bed.strip()
            block = block.strip()
        else:
            bed = bed_part
            block = ""

        location = f"{bed}-{block}" if block else bed

        new_rows.append({
            "Date": d_str,
            "Name or Item": item,
            "畝": bed,
            "区画": block if block else None,
            "Location": location,
            "Photo_id": photo_id if photo_id else None,
            "JPG_Photo": jpg_photo if jpg_photo else None,
            "Tag1": tags[0] if len(tags) > 0 else None,
            "Tag2": tags[1] if len(tags) > 1 else None,
            "Tag3": tags[2] if len(tags) > 2 else None,
            "Tag4": tags[3] if len(tags) > 3 else None,
            "Tag5": tags[4] if len(tags) > 4 else None,
        })

    return new_rows


# ===== 3) メインループ =====
all_new_rows = []

while True:
    rows = input_one_session()
    if rows:
        all_new_rows.extend(rows)
        print(f"{len(rows)} 行のデータを追加予定です。")
    cont = input("続けて入力しますか？ (y=続ける / その他=終了): ").strip().lower()
    if cont != "y":
        break

# ===== 4) 保存 =====
if all_new_rows:
    df_new = pd.DataFrame(all_new_rows)
    df = pd.concat([df, df_new], ignore_index=True)
    df.to_excel(EXCEL_PATH, index=False)
    print("Excel に保存しました。全体の行数:", len(df))
else:
    print("新しいデータは追加されませんでした。")
