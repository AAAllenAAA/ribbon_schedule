# predict_AI_result.py (最終優化版 - 新增前後比較報告)
'''
主控台:
這是實際運行排程優化流程的主程式。它負責整個流程的控制和邏輯判斷。
核心工作： 讀取排程 $\to$ 呼叫 model_1.py 獲取分數 $\to$ 執行 optimize_schedule_order 優化演算法 $\to$ 生成最終的 Excel 報告。
'''
import pandas as pd
import numpy as np
import model_1 # 引入核心預測模組
import sys
import os

# -----------------------------------------------------
# 輔助函數：根據刀次定義主子工單群組
# -----------------------------------------------------
def create_group_id(df_schedule):
    """
    根據 '刀次' 規則，為工單創建 '群組ID'。
    """
    df = df_schedule.copy()
    
    # 確保刀次是數值，且缺失值為 0
    # 由於主程式中已做 to_numeric，這裡只需處理 NaN
    df['刀次_numeric'] = df['刀次'].fillna(0)
    df['群組ID'] = None # 初始化群組ID
    
    def assign_group_ids(group):
        group_ids = []
        current_group_id_suffix = 0
        current_group_id = None
        staff_name = group.name[1] 
        
        for _, row in group.iterrows():
            if row['is_valid_job']:
                if row['刀次_numeric'] > 0:
                    current_group_id_suffix += 1
                    current_group_id = f"{staff_name}_GROUP_{current_group_id_suffix}"
                
                group_ids.append(current_group_id)
            else:
                group_ids.append(current_group_id)
                
        return pd.Series(group_ids, index=group.index)

    df['群組ID'] = df.groupby(['預計開工日', '人員'], sort=False).apply(assign_group_ids).reset_index(level=[0, 1], drop=True)
    
    df.drop(columns=['刀次_numeric'], inplace=True)
    return df


# -----------------------------------------------------
# 核心優化演算法：基於 AI 機率調整排程順序 (以群組為單元)
# -----------------------------------------------------
def optimize_schedule_order(df_schedule, df_scores, threshold=0.7):
    """
    根據 AI 機率調整排程順序，以「群組ID」為單元進行移動。
    """
    
    optimized_order_indices = [] 
    
    # AI 高機率配對列表，按機率降序排列
    high_prob_pairs = df_scores[df_scores['AI_Prob'] >= threshold].sort_values(
        by='AI_Prob', ascending=False
    )
    
    grouped_schedules = df_schedule.groupby(['預計開工日', '人員'], sort=False) 
    
    for (date, staff), group_df in grouped_schedules:
        
        all_group_ids = group_df['群組ID'].unique().tolist()
        valid_group_sequence = [g_id for g_id in all_group_ids if g_id is not None]
        original_valid_groups = valid_group_sequence
        
        if not original_valid_groups:
            # 如果沒有有效工單群組，則保留原始順序
            optimized_order_indices.extend(group_df.index.tolist())
            continue
            
        first_group_id = original_valid_groups[0]
        
        current_optimized_group_sequence = [first_group_id]
        current_group_pool = set(original_valid_groups[1:]) 
        
        while current_group_pool:
            
            last_group_id = current_optimized_group_sequence[-1]
            
            # 找到 last_group_id 群組中的最後一張**有效工單** (Group Tail)
            last_group_tail_row = group_df[
                (group_df['群組ID'] == last_group_id) & (group_df['is_valid_job'])
            ].iloc[-1]
            tail_job_id = last_group_tail_row['工單編號']
            
            # 建立 Pool 的 Group Head 列表 (潛在的下一個主工單)
            # 為了效率，先將 pool_heads 轉為字典 {工單編號: 群組ID}
            pool_heads_df = group_df[
                (group_df['群組ID'].isin(current_group_pool)) & (group_df['is_valid_job'])
            ].groupby('群組ID')['工單編號'].first().reset_index()
            pool_head_to_group = dict(zip(pool_heads_df['工單編號'], pool_heads_df['群組ID']))
            
            pool_head_job_ids = list(pool_head_to_group.keys())
            
            # 查找 (tail_job_id) 與 (pool_head_job_ids) 的最佳高機率連接
            best_pair_row = high_prob_pairs[
                ((high_prob_pairs['工單編號_A'] == tail_job_id) & (high_prob_pairs['工單編號_B'].isin(pool_head_job_ids))) |
                ((high_prob_pairs['工單編號_B'] == tail_job_id) & (high_prob_pairs['工單編號_A'].isin(pool_head_job_ids)))
            ].head(1)
            
            next_group_id = None
            if not best_pair_row.empty:
                
                paired_job_A = best_pair_row['工單編號_A'].iloc[0]
                paired_job_B = best_pair_row['工單編號_B'].iloc[0]
                
                # 確定哪個工單是 Pool Head (即下一組的頭)
                next_head_job_id = paired_job_B if paired_job_A == tail_job_id else paired_job_A
                
                # *** 修正點 1: 確保 next_head_job_id 確實是 Pool 中的一個工單 ***
                if next_head_job_id in pool_head_to_group:
                    next_group_id = pool_head_to_group[next_head_job_id]
                
            # AI 優先選擇 (閾值判斷)
            if next_group_id is not None and next_group_id in current_group_pool:
                   current_optimized_group_sequence.append(next_group_id)
                   current_group_pool.remove(next_group_id)
                   continue 

            # *** 修正點 2: 回退邏輯優化 - 優先選擇公分數接近的群組 ***
            
            # 找到當前群組尾部的公分數
            last_group_cm = last_group_tail_row.get('公分')
            
            remaining_groups_df = group_df[group_df['群組ID'].isin(current_group_pool)].copy()
            
            if not remaining_groups_df.empty and last_group_cm is not None:
                
                # 輔助函數：計算與當前公分數的差異
                def get_cm_diff(g_id):
                    head_row = group_df[(group_df['群組ID'] == g_id) & (group_df['is_valid_job'])].iloc[0]
                    # 使用絕對差值，公分數缺失則視為差異極大 (無限大)
                    diff = abs(head_row.get('公分', np.inf) - last_group_cm) if head_row.get('公分') is not None else np.inf
                    return diff

                remaining_groups_df['cm_diff'] = remaining_groups_df['群組ID'].apply(get_cm_diff)

                # 優先選擇公分數差異最小的群組，如果差異相同，則按原始順序
                # 這裡使用 groupby().first() 來確保每個群組只計算一次
                cm_ranked_groups = remaining_groups_df.groupby('群組ID').first().sort_values(
                    by=['cm_diff', '原始順序'], ascending=[True, True]
                ).index.tolist()
                
                if cm_ranked_groups:
                    next_group_id = cm_ranked_groups[0]
                    current_optimized_group_sequence.append(next_group_id)
                    current_group_pool.remove(next_group_id)
                    continue

            # 最終回退：如果公分數差異也無法判斷（例如公分數缺失），則按原始順序取
            remaining_original_groups = [
                g_id for g_id in original_valid_groups if g_id in current_group_pool
            ]
            
            if remaining_original_groups:
                next_group_id = remaining_original_groups[0]
                current_optimized_group_sequence.append(next_group_id)
                current_group_pool.remove(next_group_id)
            else:
                 break
            
        # 展開群組ID，形成最終索引序列
        reordered_indices = []
        
        for group_id in current_optimized_group_sequence:
            group_indices = group_df[group_df['群組ID'] == group_id].index.tolist()
            reordered_indices.extend(group_indices)

        # 將沒有群組ID的工單/空行（NaN/None）放到最後
        none_group_indices = group_df[group_df['群組ID'].isna()].index.tolist()
        reordered_indices.extend(none_group_indices)
        
        optimized_order_indices.extend(reordered_indices)
    
    return df_schedule.loc[optimized_order_indices].reset_index(drop=True)

# -----------------------------------------------------
# 主程式
# -----------------------------------------------------
if __name__ == "__main__":

    # === 修正點 1: 只接收一個命令行參數 ===
    if len(sys.argv) < 2:
        print("❌ 錯誤: 請提供 Touring 產出的 Excel 檔案路徑作為參數。", file=sys.stderr)
        sys.exit(1)
        
    SCHEDULE_INPUT_PATH = sys.argv[1] # 接收第一個參數 (Touring 產出的檔案)
    
    # === 修正點 2: 動態生成輸出路徑 (在原檔名後加上 _ML_OPTIMIZED) ===
    base_name = os.path.splitext(SCHEDULE_INPUT_PATH)
    # e.g., /path/to/ribbon_schedule_20251121.xlsx -> /path/to/ribbon_schedule_20251121_ML_OPTIMIZED.xlsx
    OUTPUT_RESULT_PATH = base_name[0] + "_ML_RESULT" + base_name[1] 
    
    print(f"模型將讀取: {SCHEDULE_INPUT_PATH}")
    print(f"模型將輸出至: {OUTPUT_RESULT_PATH}")

    
    # ====================================
    # 配置區：請根據您的環境調整
    # ====================================

    # ❗ 必須檢查：所有應該是數字的欄位名稱 ❗
    NUMERIC_COLUMNS = ['刀次', '米平方', '餘量', '公分', '車數', '預估良品數'] 

    INPUT_SHEET_NAME = 0 # 0 表示讀取第一個工作表，最安全
    # ====================================
    
    all_sheets_data = []
    
    print("--- 1. 載入程式初版排程 (合併多 Sheet) ---")
    
    try:
        # 步驟 1.1: 直接讀取整個檔案的第一個 Sheet
        df_schedule_original = pd.read_excel(
            SCHEDULE_INPUT_PATH, 
            sheet_name=INPUT_SHEET_NAME, 
            dtype=str, 
            keep_default_na=False
        )
        df_schedule_original.columns = df_schedule_original.columns.str.strip()
        
        # 由於主程式已經在 assign_personnel_by_similarity 步驟後給了「人員」欄位，
        # 這裡不需要「原始分組」，但需要記錄原始順序。
        
        # 設置原始排程順序
        df_schedule_original['原始順序'] = range(1, len(df_schedule_original) + 1)
        df_schedule = df_schedule_original.copy()
        
        # 這裡假設 sorted_df_end 已經包含了「人員」欄位，這對 create_group_id 很重要。
        # 由於 sorted_df_end 是單一 DataFrame，這裡不再需要 '原始分組' 欄位。
        
        print(f"✅ 成功載入排程，總計 {len(df_schedule)} 筆記錄。")
        
    except Exception as e:
        print(f"❌ 致命錯誤：無法讀取 AI 輸入檔的 Sheet。錯誤: {e}")
        sys.exit(1) # 讀取失敗則直接退出

    # 修正點 A: 標記空工單，不移除
    df_schedule['工單編號'] = df_schedule['工單編號'].replace({None: ''}).astype(str).str.strip()
    df_schedule['is_valid_job'] = df_schedule['工單編號'].apply(lambda x: len(x) > 0)
    
    valid_count = df_schedule['is_valid_job'].sum()
    empty_count = len(df_schedule) - valid_count
    print(f"💡 數據分析：有效工單 {valid_count} 筆，待配對空缺 {empty_count} 筆。")
    
    # 修正點 B: 數值欄位類型轉換
    print("--- 1.2. 轉換數值欄位類型 ---")
    for col in NUMERIC_COLUMNS:
        if col in df_schedule.columns:
            # 嘗試轉換成數字。errors='coerce' 會將無法轉換的值 (如空字串) 設為 NaN
            df_schedule[col] = pd.to_numeric(df_schedule[col], errors='coerce')
    print("✅ 數值欄位已轉換回數字類型。")
        
    # 步驟 1.3: 創建群組ID
    print("--- 1.3. 根據刀次規則創建主子工單群組ID ---")
    df_schedule = create_group_id(df_schedule)
    print(f"✅ 已識別 {df_schedule['群組ID'].nunique()} 個群組。")
    
    # 步驟 2: 呼叫 ML Scorer 獲取所有潛在配對的 AI 機率
    print("--- 2. 計算 AI 配對機率 (模型預測) ---")
    df_ai_scores = model_1.get_ai_pairing_scores(df_schedule[df_schedule['is_valid_job']])
    
    df_optimized = pd.DataFrame() # 預設空DataFrame

    if df_ai_scores.empty:
        print("💡 提示：沒有找到任何潛在配對，不執行優化。")
        df_optimized = df_schedule
        df_optimized['優化後順序'] = df_optimized['原始順序']
    else:
        print(f"✅ 成功獲得 {len(df_ai_scores)} 筆潛在配對分數。最高機率: {df_ai_scores['AI_Prob'].max():.4f}")
        
        # 步驟 3: 執行優化排序演算法
        print("--- 3. 執行基於 AI 機率的排程排序優化 (群組連動) ---")
        df_optimized = optimize_schedule_order(df_schedule, df_ai_scores, threshold=0.7)
        
        # 輸出最終順序
        df_optimized['優化後順序'] = range(1, len(df_optimized) + 1)

    # 移除最終輸出中不需要的輔助欄位
    df_optimized_final = df_optimized.drop(columns=['is_valid_job', '群組ID'], errors='ignore')
    
    # 步驟 4: 生成優化對比報告
    print("--- 4. 生成優化對比報告 ---")
    
    # 提取對比報告需要的欄位
    df_report = df_optimized_final[['工單編號', '預計開工日', '人員', '原始順序', '優化後順序']].copy()
    
    # 計算順序變動量，並僅在有工單編號的行上計算 (空行變動沒有意義)
    df_report['順序變動量'] = np.where(
        df_report['工單編號'].str.len() > 0, 
        df_report['優化後順序'] - df_report['原始順序'], 
        0
    )
    
    # 計算關鍵績效指標 (KPI)
    valid_jobs_report = df_report[df_report['工單編號'].str.len() > 0]
    total_jobs_moved = (valid_jobs_report['順序變動量'] != 0).sum()
    avg_shift = valid_jobs_report['順序變動量'].abs().mean()
    
    print(f"📊 總計移動的有效工單/群組數量: {total_jobs_moved}")
    print(f"📊 平均順序變動幅度 (絕對值): {avg_shift:.2f}")


    # 步驟 5: 輸出最終結果 (寫入多個 Sheet)
    print(f"--- 5. 輸出最終結果至 {OUTPUT_RESULT_PATH} (Excel) ---")
    
    try:
        with pd.ExcelWriter(OUTPUT_RESULT_PATH, engine='xlsxwriter') as writer:
            # Sheet 1: 最終優化排程 (這是主程式讀取並使用的結果)
            df_optimized_final.to_excel(writer, index=False, sheet_name='AI優化排程結果')
            
            # Sheet 2: 原始 vs 優化對比報告 (保留供分析)
            #df_report.to_excel(writer, index=False, sheet_name='原始_VS_優化對比')
            
            # 可選：Sheet 3: AI配對分數 (保留供調參)
            #if 'df_ai_scores' in locals() and not df_ai_scores.empty:
                #df_ai_scores.to_excel(writer, index=False, sheet_name='AI配對分數(for_Debug)')

        print("\n🎉 AI 優化排程已完成！")
        
        # *** 關鍵輸出：告訴主程式最終路徑 ***
        print(f"FINAL_PATH:{OUTPUT_RESULT_PATH}") 
        
    except Exception as e:
        print(f"❌ 輸出錯誤：請確認輸出檔案是否已被開啟。錯誤: {e}", file=sys.stderr)
        sys.exit(1) # 輸出失敗則退出

    print("下一步：請檢查輸出的 Excel 檔案，特別是 '原始_VS_優化對比' Sheet，以評估 AI 效果。")

    # 正常退出
    sys.exit(0)