#健康関連のデータをPower Automateのフローで取得するために、最初にデータ取得期間を設定する。

import datetime

# データ取得の開始日付入力
print("開始日は2021/12/29以降とする")
Nen_start=input("開始年4桁")
Tuki_start=input("開始月2桁")
Niti_start=input("開始日2桁")
           
dt_start_input=datetime.date(int(Nen_start), int(Tuki_start), int(Niti_start))
print("開始日", dt_start_input)

#データ取得の開始日付入力
Nen_end=input("終了年4桁")
Tuki_end=input("終了月2桁")
Niti_end=input("終了日2桁")
           
dt_end_input=datetime.date(int(Nen_end), int(Tuki_end), int(Niti_end))
print("終了日", dt_end_input)

if dt_end_input<dt_start_input:
    print("Eroor: 開始日が終了日より後")
else:
   print(f"START_DATE={dt_start_input}")
   print(f"END_DATE={dt_end_input}") 