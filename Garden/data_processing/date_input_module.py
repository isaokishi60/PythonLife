from datetime import datetime

def input_date():  # 日付入力の関数
    while True:  # ユーザーが正しい入力をするまでループ
        try:
            year = int(input("年を入力してください（例: 2025）："))
            month = int(input("月を入力してください（例: 5）："))
            day = int(input("日を入力してください（例: 31）："))
            dt = datetime(year, month, day)
            return dt.date()
        except ValueError as e:
            print("無効な日付です。もう一度入力してください：", e)

def validate_date(input_date):
    if input_date < datetime(2025, 1, 20).date():
        print("2025年1月20日より前の日付は無効です。")
        return None
    return input_date

