import pandas as pd
from typing import Any, Tuple

# -----------------------------
# 先定義你的函數
# -----------------------------
def merge_order_cutNum(A_df: pd.DataFrame, B_df: pd.DataFrame, C_df: pd.DataFrame,
                       daily_standard: int = 55) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    MAX_KNIFE_MAP = {1:65, 2:55, 3:51, 4:47, 5:43, 6:39}
    prefix_rules = {"B110":5, "TTR":4, "UTMX":4, "UTM":3}
    default_prefix_len = 3

    def get_item_group_key(item_code: str) -> str:
        if not isinstance(item_code, str):
            return "UNKNOWN"
        item_code = item_code.strip().replace('.', '')
        for prefix in sorted(prefix_rules.keys(), key=len, reverse=True):
            length = prefix_rules[prefix]
            if item_code.startswith(prefix):
                return item_code[:length] if len(item_code)>=length else item_code
        return item_code[:default_prefix_len] if len(item_code)>=default_prefix_len else item_code

    def safe_int_conversion(value: Any, default: int = 0) -> int:
        if pd.isna(value) or value is None or str(value).strip() == "":
            return default
        try:
            return int(float(value))
        except:
            return default

    def process_one(df: pd.DataFrame, person_name: str) -> pd.DataFrame:
        if df is None or df.empty:
            return df.copy()
        df = df.copy().reset_index(drop=True)
        df["預計開工日"] = pd.to_datetime(df["預計開工日"], errors="coerce")
        df["預計完工日"] = pd.to_datetime(df["預計完工日"], errors="coerce")

        orders = []
        last_main = None
        for i, row in df.iterrows():
            knives = safe_int_conversion(row.get("刀次",0))
            cars = safe_int_conversion(row.get("車數",0))
            od = {"orig_idx": i, "orig_row": row.copy(), "remain": knives, "car": cars, "children": [], "slices": {}}
            if knives > 0:
                od["slices"][row["預計開工日"].normalize()] = knives
                last_main = od
                orders.append(od)
            else:
                if last_main is not None:
                    last_main["children"].append(row.copy())
                else:
                    orders.append(od)
                    last_main = od

        if not orders:
            return df

        all_dates = sorted({d.normalize() for od in orders for d in od["slices"].keys() if pd.notna(d)})
        current_date = all_dates[0]

        while True:
            # 1. 計算當日狀態
            daily_ods = [od for od in orders if current_date in od["slices"] and od["slices"][current_date]>0]
            total_knives = sum(od["slices"][current_date] for od in daily_ods)
            existing_keys = {get_item_group_key(od["orig_row"].get("品號")) for od in daily_ods}

            # 每次都即時計算上限
            num_items = len(existing_keys)
            day_limit = MAX_KNIFE_MAP.get(num_items, 65 if num_items==0 else 39)
            available = max(day_limit - total_knives, 0)

            # 2. 找出未分配的未來工單
            future_cand = []
            for od in orders:
                for sdate, amt in od["slices"].items():
                    if pd.notna(sdate) and sdate.normalize() > current_date and amt>0:
                        future_cand.append((sdate, od, amt))
            future_cand.sort(key=lambda x:(x[0], x[1]["orig_idx"]))

            while available>0 and future_cand:
                pick_date, pick_od, pick_amt = future_cand[0]
                pick_key = get_item_group_key(pick_od["orig_row"].get("品號"))
                # 計算即時新上限
                new_existing_keys = existing_keys | {pick_key}
                new_num_items = len(new_existing_keys)
                new_day_limit = MAX_KNIFE_MAP.get(new_num_items, 39)
                fill_target = min(available, pick_amt, new_day_limit - total_knives)
                if fill_target <=0:
                    break

                # 執行遞補
                pick_od["slices"][pick_date] -= fill_target
                if pick_od["slices"][pick_date]<=0:
                    del pick_od["slices"][pick_date]
                pick_od["slices"][current_date] = pick_od["slices"].get(current_date,0)+fill_target

                # 更新狀態
                total_knives += fill_target
                existing_keys.add(pick_key)
                available = new_day_limit - total_knives

                # 更新候選
                future_cand = [(d,o,a) for d,o,a in future_cand if o["slices"].get(d,0)>0]

            # 下一天
            future_remaining = [(d,od) for od in orders for d,amt in od["slices"].items() if d>current_date and amt>0]
            if not future_remaining:
                break
            current_date += pd.Timedelta(days=1)

        # 展開成 DataFrame
        out_rows = []
        for od in orders:
            for sdate, amt in sorted(od["slices"].items(), key=lambda x:x[0]):
                if amt>0:
                    main_row = od["orig_row"].copy()
                    main_row["預計開工日"] = sdate
                    main_row["預計完工日"] = sdate
                    main_row["刀次"] = int(amt)
                    main_car = safe_int_conversion(main_row.get("車數",0))
                    main_row["預估良品數"] = main_car*int(amt)
                    out_rows.append(main_row)
                    for ch in od["children"]:
                        ch_row = ch.copy()
                        ch_row["預計開工日"] = sdate
                        ch_row["預計完工日"] = sdate
                        ch_row["刀次"] = 0
                        ch_car = safe_int_conversion(ch_row.get("車數",0))
                        ch_row["預估良品數"] = ch_car*int(amt)
                        if "餘量" in ch_row.index:
                            ch_row["餘量"] = ch_row["預估良品數"]
                        out_rows.append(ch_row)

        out_df = pd.DataFrame(out_rows)
        out_df["預計開工日"] = out_df["預計開工日"].dt.strftime("%Y-%m-%d")
        out_df["預計完工日"] = out_df["預計完工日"].dt.strftime("%Y-%m-%d")
        return out_df.reset_index(drop=True)

    A_out = process_one(A_df, "容合 (A830)")
    B_out = process_one(B_df, "旺斌 (B201)")
    C_out = process_one(C_df, "家偉 (A159)")
    return A_out, B_out, C_out
# -----------------------------
# 建立小測試資料
# -----------------------------
data = [
    {"預計開工日":"2025-10-20","人員":"A","工單編號":"O001","品號":"X1","刀次":30,"車數":1,"預計完工日":"2025-10-20"},
    {"預計開工日":"2025-10-21","人員":"A","工單編號":"O002","品號":"X2","刀次":40,"車數":1,"預計完工日":"2025-10-20"},
    {"預計開工日":"2025-10-22","人員":"A","工單編號":"O003","品號":"X3","刀次":20,"車數":1,"預計完工日":"2025-10-20"},
    {"預計開工日":"2025-10-23","人員":"A","工單編號":"O004","品號":"X4","刀次":50,"車數":1,"預計完工日":"2025-10-20"},
]

df_A = pd.DataFrame([row for row in data if row["人員"]=="A"])
df_B = pd.DataFrame([])  # 空資料
df_C = pd.DataFrame([])  # 空資料

# -----------------------------
# 執行函數
# -----------------------------
A_out, B_out, C_out = merge_order_cutNum(df_A, df_B, df_C)

# -----------------------------
# 印出結果
# -----------------------------
print("======= A_out =======")
print(A_out[["預計開工日","工單編號","刀次"]])
