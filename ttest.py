
import pandas as pd
from datetime import timedelta

def generate_schedule_whole_day(df: pd.DataFrame, break_dates: set, max_lookback_days: int = 30) -> pd.DataFrame:
    df = df.copy()
    df["實際排程日期"] = None

    # 每日產能限制（週一到週五）
    daily_limits = {0: 65, 1: 65, 2: 55, 3: 65, 4: 55}

    # 建立主工單索引（刀次 > 0 為主工單）
    last_main_idx = None
    df["主工單索引"] = None
    for idx, row in df.iterrows():
        if row["刀次"] > 0:
            last_main_idx = idx
        df.at[idx, "主工單索引"] = last_main_idx

    # 全域產能初始化
    overall_end = df["預計完工日"].max()
    overall_start = overall_end - timedelta(days=max_lookback_days)
    all_dates = pd.date_range(overall_start, overall_end)

    schedule_capacity = {}
    for d in all_dates:
        if d.weekday() < 5 and d not in break_dates:
            schedule_capacity[d] = daily_limits.get(d.weekday(), 55)

    # 分組處理（依主工單索引），排序邏輯：組內最早完工日，整體由晚到早
    df["組內最早完工日"] = df.groupby("主工單索引")["預計完工日"].transform("min")
    groups = df.groupby("主工單索引")
    sorted_groups = sorted(groups, key=lambda g: df.loc[g[0], "組內最早完工日"], reverse=True)

    for main_idx, group in sorted_groups:
        group_rows = df.loc[group.index]
        earliest_end = group_rows["預計完工日"].min()
        start_date = earliest_end - timedelta(days=max_lookback_days)

        total_doz = group_rows["刀次"].sum()

        # 從 earliest_end 往前找可排日期
        possible_dates = [d for d in pd.date_range(start_date, earliest_end)
                          if d.weekday() < 5 and d not in break_dates]
        possible_dates = sorted(possible_dates, reverse=True)

        assigned_date = None
        for d in possible_dates:
            if schedule_capacity.get(d, 0) >= total_doz:
                schedule_capacity[d] -= total_doz
                assigned_date = d
                break

        if assigned_date is None:
            raise Exception(f"排程無法完成，工單 {df.loc[main_idx, '工單編號']} 無法安排")

        # 整組工單一起設置排程日期
        for idx in group.index:
            df.at[idx, "實際排程日期"] = assigned_date.strftime("%Y-%m-%d")

    df.drop(columns=["組內最早完工日"], inplace=True)
    return df


if __name__ == "__main__":
    '''
    data = {
        "預計開工日": ["2025-08-11", "2025-08-04", "2025-08-28", "2025-08-28", "2025-08-28"],
        "預計完工日": ["2025-08-11", "2025-08-04", "2025-08-28", "2025-08-28", "2025-08-28"],
        "刀次": [60, 0, 55, 55, 40],
        "工單編號": ["81A03587", "81A03569", "81A03609", "81A03609", "81A03609"],
    }
    '''
    data = {
        "預計開工日": [
            "2025-08-01", "2025-08-01", "2025-08-04", "2025-08-05", "2025-08-07", "2025-08-07", "2025-08-07",
            "2025-08-08", "2025-08-08", "2025-08-08", "2025-08-01", "2025-08-08", "2025-08-11", "2025-08-11",
            "2025-08-11", "2025-08-11", "2025-08-04", "2025-08-12", "2025-08-06", "2025-08-13", "2025-08-13",
            "2025-08-13", "2025-08-15", "2025-08-19", "2025-08-18", "2025-08-18", "2025-08-18", "2025-08-18",
            "2025-08-19", "2025-08-19", "2025-08-19", "2025-08-21", "2025-08-21"
        ],
        "預計完工日": [
            "2025-08-01", "2025-08-01", "2025-08-04", "2025-08-05", "2025-08-07", "2025-08-07", "2025-08-07",
            "2025-08-08", "2025-08-08", "2025-08-08", "2025-08-01", "2025-08-08", "2025-08-11", "2025-08-11",
            "2025-08-11", "2025-08-11", "2025-08-04", "2025-08-12", "2025-08-06", "2025-08-13", "2025-08-13",
            "2025-08-13", "2025-08-15", "2025-08-19", "2025-08-18", "2025-08-18", "2025-08-18", "2025-08-18",
            "2025-08-19", "2025-08-19", "2025-08-19", "2025-08-21", "2025-08-21"
        ],
        "刀次": [
            55, 6, 30, 4, 55, 55, 40, 18, 0, 10, 0, 25, 9, 31, 24, 60, 0, 22, 0, 55, 55, 15, 30, 0, 55, 55, 55, 1,
            15, 55, 7, 55, 19
        ],
        "工單編號": [
            "81A03565", "81A03565", "81A03568", "81A03572", "81A03551", "81A03551", "81A03551",
            "81A03578", "81A03581", "81A03580", "81A03565", "81A03582", "81A03584", "81A03585",
            "81A03586", "81A03587", "81A03569", "81A03591", "81A03576", "81A03553", "81A03553",
            "81A03553", "81A03600", "81A03602", "81A03599", "81A03599", "81A03599", "81A03599",
            "81A03604", "81A03605", "81A03605", "81A03606", "81A03606"
        ]
    }


    df_test = pd.DataFrame(data)
    df_test["預計開工日"] = pd.to_datetime(df_test["預計開工日"])
    df_test["預計完工日"] = pd.to_datetime(df_test["預計完工日"])

    break_dates = {
        pd.Timestamp("2025-08-08"),  # 休假日
    }

    result_df = generate_schedule_whole_day(df_test, break_dates)
    print(result_df[["工單編號", "刀次", "實際排程日期"]])

    test_data = {
        "預計開工日": ["2025-07-01", "2025-07-05"],
        "預計完工日": ["2025-07-10", "2025-07-15"],
        "刀次": [10, 20],
        "需求數量": [100, 200],
        "工單編號": ["W001", "W002"],
        "人員": ["A830", "A830"],
    }

    df_test = pd.DataFrame(test_data)

    # 確保欄位轉 datetime
    df_test["預計開工日"] = pd.to_datetime(df_test["預計開工日"])
    df_test["預計完工日"] = pd.to_datetime(df_test["預計完工日"])

    print(df_test)
