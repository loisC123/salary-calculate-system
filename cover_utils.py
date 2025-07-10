import pandas as pd
from datetime import datetime, timedelta

def load_cover_mapping(cover_filename):
    """
    從 cover.xlsx 載入代班設定：
    回傳一個 dict： key = (原員工, 個案姓名, 服務日期) -> 代班員工
    """
    cover_df = pd.read_excel(cover_filename, sheet_name=0)
    mapping = {}
    for idx, row in cover_df.iterrows():
        original_emp = str(row['原員工']).strip()
        cover_emp = str(row['代班員工']).strip()
        case_name = str(row['個案姓名']).strip()
        roc_date_str = str(int(row['服務日期(民國7碼)'])).zfill(7)
        year = 1911 + int(roc_date_str[:3])
        month = int(roc_date_str[3:5])
        day = int(roc_date_str[5:7])
        service_date = pd.Timestamp(year, month, day)
        mapping[(original_emp, case_name, service_date)] = cover_emp
    return mapping

def compute_cover_summaries(df, cover_mapping, holidays):
    """
    根據主表與代班設定，回傳兩個 DataFrame：
    - BA代班薪資轉交
    - GA/SC代班薪資轉交
    """
    ba_cover_summary = {}
    ga_sc_cover_summary = {}

    for emp, emp_group in df.groupby('服務人員姓名'):
        emp_group_sorted = emp_group.sort_values(['服務日期', '完整開始時間']).reset_index(drop=True)

        # 專門針對 BA 轉場判斷代班
        ba_rows = emp_group_sorted[emp_group_sorted['服務項目代碼'].astype(str).str.contains('BA')]
        for date, day_group in ba_rows.groupby('服務日期'):
            day_group = day_group.sort_values('完整開始時間').reset_index(drop=True)
            if len(day_group) > 0:
                first_row = day_group.loc[0]
                cover_key = (emp, first_row['個案姓名'], first_row['服務日期'])
                if cover_key in cover_mapping:
                    cover_emp = cover_mapping[cover_key]
                    print(f"[DEBUG] 找到代班對應：{cover_key} -> {cover_mapping[cover_key]}")
                    k = (emp, cover_emp)
                    if k not in ba_cover_summary:
                        ba_cover_summary[k] = {'平日': 0, '假日': 0, '轉場': 0, '薪資': 0}
                    ba_cover_summary[k]['轉場'] += 1
                    ba_cover_summary[k]['薪資'] += 35
            for idx in range(1, len(day_group)):
                prev = day_group.loc[idx - 1]
                curr = day_group.loc[idx]
                gap = (curr['完整開始時間'] - prev['完整結束時間']).total_seconds() / 60
                if gap > 1:
                    cover_key = (emp, curr['個案姓名'], curr['服務日期'])
                    if cover_key in cover_mapping:
                        cover_emp = cover_mapping[cover_key]
                        k = (emp, cover_emp)
                        if k not in ba_cover_summary:
                            ba_cover_summary[k] = {'平日': 0, '假日': 0, '轉場': 0, '薪資': 0}
                        ba_cover_summary[k]['轉場'] += 1
                        ba_cover_summary[k]['薪資'] += 35

        # 每筆資料判斷代班
        for idx, row in emp_group_sorted.iterrows():
            row_service_date = row['服務日期']
            case = row['個案姓名']
            cover_key = (emp, case, row_service_date)
            if cover_key not in cover_mapping:
                continue
            cover_emp = cover_mapping[cover_key]

            weekday = row_service_date.weekday()
            date_str = row_service_date.strftime('%m/%d')
            is_holiday = date_str in holidays or weekday >= 5
            minutes = (row['完整結束時間'] - row['完整開始時間']).total_seconds() / 60

            # BA 處理時數與薪資
            if "BA" in str(row['服務項目代碼']):
                hours = minutes / 60
                k = (emp, cover_emp)
                if k not in ba_cover_summary:
                    ba_cover_summary[k] = {'平日': 0, '假日': 0, '轉場': 0, '薪資': 0}
                if is_holiday:
                    ba_cover_summary[k]['假日'] += hours
                    ba_cover_summary[k]['薪資'] += hours * 220 * 2
                else:
                    ba_cover_summary[k]['平日'] += hours
                    ba_cover_summary[k]['薪資'] += hours * 220
            # GA/SC 處理次數與薪資
            elif any(x in str(row['服務項目代碼']) for x in ["GA", "SC"]):
                qty = int(row['數量\n(僅整數)']) if not pd.isna(row['數量\n(僅整數)']) else 1
                k = (emp, cover_emp)
                if k not in ga_sc_cover_summary:
                    ga_sc_cover_summary[k] = {'次數': 0, '薪資': 0}
                ga_sc_cover_summary[k]['次數'] += qty
                ga_sc_cover_summary[k]['薪資'] += qty * 650

    # 轉為 DataFrame
    ba_cover_rows = []
    for (origin, cover), vals in ba_cover_summary.items():
        row = [origin, cover, round(vals['平日'], 2), round(vals['假日'], 2), int(vals['轉場']), int(vals['薪資'])]
        ba_cover_rows.append(row)
    ba_cover_df = pd.DataFrame(ba_cover_rows, columns=['員工', '代班', '平日代班', '假日代班', '轉場', '代班薪資'])

    ga_cover_rows = []
    for (origin, cover), vals in ga_sc_cover_summary.items():
        row = [origin, cover, vals['次數'], vals['薪資']]
        ga_cover_rows.append(row)
    ga_cover_df = pd.DataFrame(ga_cover_rows, columns=['員工', '代班', '次數', '薪資'])

    return ba_cover_df, ga_cover_df
