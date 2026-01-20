import pandas as pd
from typing import Tuple

# 範例資料
data = [
    {"預計開工日":"2025/09/10", "工單編號": "81A03708", "品號": "TTDW11030N.", "良品數": 232, "公分":11, "車數":8, "刀次":29},
    {"預計開工日":"2025/09/10", "工單編號": "81A03708", "品號": "TTDW11030N.", "良品數": 238, "公分":11, "車數":8, "刀次":26},
    {"預計開工日":"2025/09/10", "工單編號": "81A03722", "品號": "UTM065070.", "良品數": 420, "公分":6.5, "車數":7, "刀次":46},
    {"預計開工日":"2025/09/10", "工單編號": "81A03723", "品號": "UTM083070.", "良品數": 230, "公分":8.3, "車數":5, "刀次":0},
    {"預計開工日":"2025/09/11", "工單編號": "81A03722", "品號": "UTM065070.", "良品數": 420, "公分":6.5, "車數":7, "刀次":46},
    {"預計開工日":"2025/09/11", "工單編號": "81A03723", "品號": "UTM083070.", "良品數": 230, "公分":8.3, "車數":5, "刀次":0},
    {"預計開工日":"2025/09/11", "工單編號": "81A03722", "品號": "UTM065070.", "良品數": 14,  "公分":6.5, "車數":7, "刀次":2},
    {"預計開工日":"2025/09/11", "工單編號": "81A03723", "品號": "UTM083070.", "良品數": 10,  "公分":8.3, "車數":5, "刀次":0},
    {"預計開工日":"2025/09/11", "工單編號": "81A03722", "品號": "UTM065070.", "良品數": 68,  "公分":6.5, "車數":7, "刀次":7},
    {"預計開工日":"2025/09/11", "工單編號": "",          "品號": "UTM083070.", "良品數": 35,  "公分":8.3, "車數":5, "刀次":0},
]

data2 = []

df = pd.DataFrame(data)

df2 = pd.DataFrame(data2)

def merge_same_day_orders_multi(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df

    df = df.copy()
    df["預計開工日"] = pd.to_datetime(df["預計開工日"])
    merged_rows = []

    i = 0
    while i < len(df):
        row = df.iloc[i]
        row_date = row["預計開工日"]

        if row["刀次"] > 0:
            main_cut = row["刀次"]
            main_car = row["車數"]
            main_id = row["工單編號"]

            # 找對應子工單
            j = i + 1
            subs = []
            while j < len(df) and df.iloc[j]["刀次"] == 0 and df.iloc[j]["預計開工日"] == row_date:
                sub_id = str(df.iloc[j]["工單編號"]).strip()
                if sub_id == "":
                    sub_id = "EMPTY_SUB"   # 空白視為一種工單
                subs.append(sub_id)
                j += 1

            merged = False
            if merged_rows:
                # 找前一筆主工單
                prev_idx = -1
                for k in range(len(merged_rows)-1, -1, -1):
                    if merged_rows[k]["刀次"] > 0:
                        prev_idx = k
                        break
                if prev_idx >= 0:
                    prev_main = merged_rows[prev_idx]
                    prev_main_id = prev_main["工單編號"]
                    prev_date = prev_main["預計開工日"]

                    # 找前一組子工單
                    prev_subs = []
                    m = prev_idx + 1
                    while m < len(merged_rows) and merged_rows[m]["刀次"] == 0 and merged_rows[m]["預計開工日"] == prev_date:
                        prev_sub_id = str(merged_rows[m]["工單編號"]).strip()
                        if prev_sub_id == "":
                            prev_sub_id = "EMPTY_SUB"
                        prev_subs.append(prev_sub_id)
                        m += 1

                    # 判斷是否可合併
                    if prev_main_id == main_id and prev_date == row_date:
                        # 沒子工單 → 可合併
                        if len(subs) == 0 and len(prev_subs) == 0:
                            merged_rows[prev_idx]["刀次"] += main_cut
                            merged_rows[prev_idx]["良品數"] += main_cut * main_car
                            merged = True
                        # 子工單完全相同 → 可合併
                        elif subs == prev_subs:
                            merged_rows[prev_idx]["刀次"] += main_cut
                            merged_rows[prev_idx]["良品數"] += main_cut * main_car
                            for k, sub_id in enumerate(subs):
                                merged_rows[prev_idx + 1 + k]["良品數"] += main_cut * merged_rows[prev_idx + 1 + k]["車數"]
                            merged = True

            if not merged:
                # 新增主工單
                new_main = row.copy()
                new_main["良品數"] = main_cut * main_car
                merged_rows.append(new_main)
                # 新增子工單
                j2 = i + 1
                while j2 < len(df) and df.iloc[j2]["刀次"] == 0 and df.iloc[j2]["預計開工日"] == row_date:
                    new_sub = df.iloc[j2].copy()
                    new_sub["良品數"] = main_cut * new_sub["車數"]
                    merged_rows.append(new_sub)
                    j2 += 1

            i = j
        else:
            merged_rows.append(row.copy())
            i += 1

    merged_df = pd.DataFrame(merged_rows)
    merged_df["預計開工日"] = merged_df["預計開工日"].dt.strftime("%Y-%m-%d")
    
    return merged_df


'''
def merge_same_day_orders_multi(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df

    df = df.copy()
    df["預計開工日"] = pd.to_datetime(df["預計開工日"])
    merged_rows = []

    i = 0
    while i < len(df):
        row = df.iloc[i]
        row_date = row["預計開工日"]
        if row["刀次"] > 0:
            main_cut = row["刀次"]
            main_car = row["車數"]
            main_id = row["工單編號"]

            # 找對應子工單
            j = i + 1
            subs = []
            while j < len(df) and df.iloc[j]["刀次"] == 0 and df.iloc[j]["預計開工日"] == row_date:
                subs.append(df.iloc[j])
                j += 1

            merged = False
            if merged_rows:
                prev_idx = -(len(subs)+1)
                if abs(prev_idx) <= len(merged_rows):
                    prev_main = merged_rows[prev_idx]
                    prev_main_id = prev_main["工單編號"]
                    prev_date = prev_main["預計開工日"]

                    if prev_main_id == main_id and prev_date == row_date:
                        prev_subs = merged_rows[prev_idx+1 : prev_idx+1+len(subs)]
                        if len(prev_subs) == len(subs) and len(subs) > 0:
                            can_merge = True
                            for k in range(len(subs)):
                                if str(subs[k]["工單編號"]).strip() == "" or subs[k]["工單編號"] != prev_subs[k]["工單編號"]:
                                    can_merge = False
                                    break
                            if can_merge:
                                merged_rows[prev_idx]["刀次"] += main_cut
                                merged_rows[prev_idx]["良品數"] += main_cut * main_car
                                for k, sub in enumerate(subs):
                                    merged_rows[prev_idx+1+k]["良品數"] += main_cut * sub["車數"]
                                merged = True
                        elif len(subs) == 0:
                            merged_rows[prev_idx]["刀次"] += main_cut
                            merged_rows[prev_idx]["良品數"] += main_cut * main_car
                            merged = True

            if not merged:
                new_main = row.copy()
                new_main["良品數"] = main_cut * main_car
                merged_rows.append(new_main)
                for sub in subs:
                    new_sub = sub.copy()
                    new_sub["良品數"] = main_cut * sub["車數"]
                    merged_rows.append(new_sub)

            i = j
        else:
            merged_rows.append(row.copy())
            i += 1

    merged_df = pd.DataFrame(merged_rows)
    merged_df["預計開工日"] = merged_df["預計開工日"].dt.strftime("%Y-%m-%d")
    return merged_df
'''

def reorder_and_merge(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["預計開工日"] = pd.to_datetime(df["預計開工日"])
    merged_rows = []

    i = 0
    while i < len(df):
        row = df.iloc[i]
        row_date = row["預計開工日"]
        main_cut = row["刀次"]
        main_car = row["車數"]
        main_id = row["工單編號"]

        if main_cut > 0:
            # 找對應子工單（直到下一個主工單或空白子工單）
            j = i + 1
            subs = []
            while j < len(df) and df.iloc[j]["刀次"] == 0 and df.iloc[j]["預計開工日"] == row_date:
                subs.append(df.iloc[j])
                j += 1

            merged = False
            if merged_rows:
                # 找前一筆主工單
                prev_idx = -1
                for k in range(len(merged_rows)-1, -1, -1):
                    if merged_rows[k]["刀次"] > 0:
                        prev_idx = k
                        break
                prev = merged_rows[prev_idx]
                prev_main_id = prev["工單編號"]
                prev_date = prev["預計開工日"]

                # 檢查前後子工單是否存在
                prev_subs_empty = True
                next_idx = prev_idx + 1
                if next_idx < len(merged_rows) and merged_rows[next_idx]["刀次"] == 0:
                    prev_subs_empty = False

                # 合併條件：
                # 1. 有子工單集合完全相同
                # 2. 或前後兩筆主工單都沒有子工單
                if prev_main_id == main_id and prev_date == row_date and (
                    (len(subs) > 0 and all(str(sub["工單編號"]).strip() != "" for sub in subs) and prev_subs_empty) 
                    or (len(subs) == 0 and prev_subs_empty)
                ):
                    merged_rows[prev_idx]["刀次"] += main_cut
                    merged_rows[prev_idx]["良品數"] += main_cut * main_car
                    for k, sub in enumerate(subs):
                        merged_rows[prev_idx + 1 + k]["良品數"] += main_cut * sub["車數"]
                    merged = True

            if not merged:
                # 新增主工單
                new_main = row.copy()
                new_main["良品數"] = main_cut * main_car
                merged_rows.append(new_main)
                # 新增子工單
                for sub in subs:
                    new_sub = sub.copy()
                    new_sub["良品數"] = main_cut * sub["車數"]
                    merged_rows.append(new_sub)

            i = j
        else:
            # 子工單但前面沒有主工單，不合併
            merged_rows.append(row.copy())
            i += 1

    return pd.DataFrame(merged_rows)

#result = reorder_and_merge(df)
result = merge_same_day_orders_multi(df)
print(result[["預計開工日", "工單編號","品號","車數","刀次","良品數"]])
