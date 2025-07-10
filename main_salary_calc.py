import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from cover_utils import load_cover_mapping, compute_cover_summaries

# 固定路徑讀檔（請自行修改）
filename = r"照顧_20250706.xls"
cover_filename = r"cover.xlsx"

# 要視為假日加班的日期（MM/DD 格式，含前導 0）
holidays = [""]

# 載入 Excel
xls = pd.ExcelFile(filename)
sheet_names = xls.sheet_names
df = pd.read_excel(xls, sheet_name=sheet_names[1])
print("已成功讀取第二工作表")

# ROC 7碼字串轉 datetime
def roc7_to_datetime(roc_str):
    roc_str = str(int(roc_str)).zfill(7)
    year = 1911 + int(roc_str[:3])
    month = int(roc_str[3:5])
    day = int(roc_str[5:7])
    return datetime(year, month, day)

# 載入代班設定
cover_mapping = load_cover_mapping(cover_filename)
print(f"[DEBUG] 代班設定共載入：{len(cover_mapping)} 筆")

# 補齊完整起訖時間欄位
df['服務日期'] = df['服務日期(請輸入7碼)'].apply(roc7_to_datetime)
df['完整開始時間'] = df.apply(
    lambda row: row['服務日期'] + timedelta(hours=int(row['起始時段-小時(24小時制)']),
                                        minutes=int(row['起始時段-分鐘'])),
    axis=1
)
df['完整結束時間'] = df.apply(
    lambda row: row['服務日期'] + timedelta(hours=int(row['結束時段-小時(24小時制)']),
                                        minutes=int(row['結束時段-分鐘'])),
    axis=1
)

# === 主要欄位統計 ===
ba_summary = {}              # 員工 -> 平日/加班/假日/轉場統計
ba_daily_records = []       # 工時日表記錄
ga_sc_summary = {}          # 員工 -> GA/SC 次數與薪資

for emp, emp_group in df.groupby('服務人員姓名'):
    emp_ba_p1 = emp_ba_p1_34 = emp_ba_p1_67 = emp_ba_p2 = trans_count = 0
    emp_ba_days = {}
    emp_total_minutes = {}
    ga_sc_count = ga_sc_salary = 0

    emp_group_sorted = emp_group.sort_values(['服務日期', '完整開始時間']).reset_index(drop=True)
    
    # 計算轉場
    ba_rows = emp_group_sorted[emp_group_sorted['服務項目代碼'].astype(str).str.contains('BA')]
    for date, day_group in ba_rows.groupby('服務日期'):
        day_group = day_group.sort_values('完整開始時間').reset_index(drop=True)
        if len(day_group) > 0:
            trans_count += 1
        for idx in range(1, len(day_group)):
            prev = day_group.loc[idx - 1]
            curr = day_group.loc[idx]
            gap = (curr['完整開始時間'] - prev['完整結束時間']).total_seconds() / 60
            if gap > 1:
                trans_count += 1

    for idx, row in emp_group_sorted.iterrows():
        minutes = (row['完整結束時間'] - row['完整開始時間']).total_seconds() / 60
        date_str = f"{row['服務日期'].month}/{row['服務日期'].day}"
        day_key = row['服務日期'].strftime('%Y-%m-%d')
        case = row['個案姓名']

        # 分類統計
        if "BA" in str(row['服務項目代碼']):
            if (emp, case) not in emp_total_minutes:
                emp_total_minutes[(emp, case)] = {}
            emp_total_minutes[(emp, case)][date_str] = emp_total_minutes[(emp, case)].get(date_str, 0) + minutes
            emp_ba_days[day_key] = emp_ba_days.get(day_key, 0) + minutes
        elif any(x in str(row['服務項目代碼']) for x in ["GA", "SC"]):
            qty = int(row['數量\n(僅整數)']) if not pd.isna(row['數量\n(僅整數)']) else 1
            ga_sc_count += qty
            ga_sc_salary += qty * 650

    for day, total_mins in emp_ba_days.items():
        date_obj = datetime.strptime(day, '%Y-%m-%d')
        date_str = date_obj.strftime('%m/%d')
        weekday = date_obj.weekday()
        is_holiday = date_str in holidays or weekday >= 5
        hours = total_mins / 60
        if is_holiday:
            emp_ba_p2 += hours
        else:
            if total_mins > 600:
                emp_ba_p1_67 += (total_mins - 600) / 60
                emp_ba_p1_34 += (600 - 480) / 60
                emp_ba_p1 += 8
            elif total_mins > 480:
                emp_ba_p1_34 += (total_mins - 480) / 60
                emp_ba_p1 += 8
            else:
                emp_ba_p1 += hours

    salary = emp_ba_p1*220 + emp_ba_p1_34*220*1.34 + emp_ba_p1_67*220*1.67 + emp_ba_p2*220*2 + trans_count*35
    ba_summary[emp] = [round(emp_ba_p1,1), round(emp_ba_p1_34,1), round(emp_ba_p1_67,1),
                       round(emp_ba_p2,1), trans_count, int(salary)]

    for (emp_name, case), minutes_dict in emp_total_minutes.items():
        row_vals = [emp_name, case]
        total = 0
        for d in range(1, 31):
            day_label = f"6/{d}"
            mins = minutes_dict.get(day_label, 0)
            row_vals.append(int(mins) if mins else "")
            total += mins
        row_vals.append(int(total))
        ba_daily_records.append(row_vals)

    if ga_sc_count > 0:
        ga_sc_summary[emp] = {"員工": emp, "次數": ga_sc_count, "喘息薪資": ga_sc_salary}

# == 主表產出 ==
ba_df = pd.DataFrame([[emp] + vals for emp, vals in ba_summary.items()],
                     columns=['員工','平日*1','平日*1.34','平日*1.67','假日*2','轉場','BA+轉場薪資'])

ba_columns = ['服務人員姓名','個案姓名'] + [f"6/{i}" for i in range(1,31)] + ['總計']
ba_hours_df = pd.DataFrame(ba_daily_records, columns=ba_columns)

summary_rows = []
for emp in ba_hours_df['服務人員姓名'].unique():
    emp_df = ba_hours_df[ba_hours_df['服務人員姓名'] == emp]
    sums = emp_df.iloc[:, 2:-1].apply(pd.to_numeric, errors='coerce').sum()
    total = sums.sum()
    row = [emp, '合計'] + list(sums) + [total]
    summary_rows.append(row)
ba_hours_df = pd.concat([ba_hours_df, pd.DataFrame(summary_rows, columns=ba_columns)], ignore_index=True)

ga_sc_df = pd.DataFrame(list(ga_sc_summary.values()), columns=['員工','次數','喘息薪資'])

# == 代班表（來自模組）==
ba_cover_df, ga_cover_df = compute_cover_summaries(df, cover_mapping, holidays)

# 匯出
now = datetime.now().strftime("%Y%m%d_%H%M")
with pd.ExcelWriter(f'時數統計_{now}.xlsx') as writer:
    ba_df.to_excel(writer, sheet_name='BA時數表', index=False)
    ba_hours_df.to_excel(writer, sheet_name='BA工時表', index=False)
    ga_sc_df.to_excel(writer, sheet_name='GA_SC表', index=False)
    ba_cover_df.to_excel(writer, sheet_name='BA代班薪資轉交', index=False)
    ga_cover_df.to_excel(writer, sheet_name='GA_SC代班薪資轉交', index=False)

print("已匯出：時數統計.xlsx")
