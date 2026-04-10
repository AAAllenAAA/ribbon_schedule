# create_ml_data.py (最終特徵版：符合品號優先級)
'''
核心工作： 載入歷史排程 -> 轉換為「工單配對」-> 計算「優化特徵」-> 生成訓練數據集。
'''
import pandas as pd
import os
import sys
import io
from datetime import datetime
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
# ====================================
# 配置區：請根據您的環境調整
# ====================================
#HISTORY_PATH = r'E:\ribbon_schedule\history\history_10_actual.csv' # 歷史/實際排程資料
HISTORY_PATH = r'E:\ribbon_schedule\history\history.csv' 
OUTPUT_ML_DATA_PATH = r"E:\ribbon_schedule\training_AI\ML\ML_Training.csv"
# ❗ 關鍵修正：CSV 檔案的編碼 ❗
CSV_ENCODING = 'big5'  # 嘗試最適合繁體中文的編碼
ENCODINGS_TO_TRY = ['big5', 'utf-8-sig', 'cp950']
# ====================================

# ------------------------------
# 輔助函數：相似度/數值/日期轉換
# ------------------------------
def product_similarity(p1, p2):
    """計算品號字串相似度"""
    p1, p2 = str(p1), str(p2)
    max_len = max(len(p1), len(p2))
    if max_len == 0:
        return 0.0
    matches = sum(a == b for a, b in zip(p1, p2))
    return matches / max_len

def safe_to_numeric(value):
    """將單一值轉換為數值，如果失敗/缺失，則設為 0.0"""
    try:
        num_value = pd.to_numeric(value, errors='coerce')
        return num_value if pd.notna(num_value) else 0.0
    except:
        return 0.0

# ------------------------------
# 核心函數：計算優化特徵 (最終版)
# ------------------------------
def create_pair_features(row_A, row_B):
    """
    計算兩張工單之間的 ML 特徵 (符合您的層級優化需求)
    特徵列表：
    1. is_item_identical (品號完全相同 - 最高優先級)
    2. is_cm_identical (公分數完全相同 - 第二優先級)
    3. item_similarity_score (品號相似度 - 第三優先級)
    4. operator_is_same (人員是否相同)
    """
    
    # 數值處理
    公分_A = safe_to_numeric(row_A['公分'])
    公分_B = safe_to_numeric(row_B['公分'])
    
    # 字串處理
    品號_A = str(row_A.get('品號', '')).strip()
    品號_B = str(row_B.get('品號', '')).strip()
    
    features = {
        # 原始識別碼 (不作為特徵，用於驗證，最後會移除)
        '工單編號_A': row_A['工單編號'],
        '工單編號_B': row_B['工單編號'],
        
        # ❗ 最終定案特徵 (英文 Key for ML) ❗
        
        # 優先級 1：品號完全相同 (最高優化)
        'is_item_identical': 1 if 品號_A == 品號_B else 0,
        
        # 優先級 2：公分數完全相同 (次要優化)
        'is_cm_identical': 1 if 公分_A == 公分_B else 0,
        
        # 優先級 3：品號相似度
        'item_similarity_score': product_similarity(品號_A, 品號_B),
        
        # 流程約束：人員是否相同
        'operator_is_same': 1 if row_A['人員'] == row_B['人員'] else 0,
        
        # ❗ 原有的 '公分數_差', '公分數是否相同' 已被移除/取代 ❗
        # ❗ is_parent_child_pair 已被移除 ❗
    }
    
    return features

def create_training_data(df_history):
    """生成訓練數據集，將單行工單轉換為兩兩配對行"""
    
    df_history = df_history.copy()
    
    # 統一欄位名稱並處理類型
    for col in ['公分']:
        df_history[col] = pd.to_numeric(df_history[col], errors='coerce').fillna(0)

    df_history['排程日'] = pd.to_datetime(df_history['預計開工日'], errors='coerce')
    
    
    # 移除關鍵欄位缺失的資料
    df_history.dropna(subset=['工單編號', '排程日', '公分', '人員'], inplace=True)
    
    # 排序：這是定義「相鄰」工單的基礎 (人工排程的順序)
    df_history.sort_values(by=['排程日', '人員', '工單編號'], inplace=True)
    df_history.reset_index(drop=True, inplace=True)
    
    all_pairs_data = []
    
    # 迭代生成配對：考慮 i 及其後面的幾張工單作為潛在配對
    for i in range(len(df_history)):
        row_A = df_history.iloc[i]
        
        # 只看接下來 10 筆工單的潛在配對 (這是模型學習的範圍)
        for j in range(i + 1, min(i + 11, len(df_history))): 
            row_B = df_history.iloc[j]
            
            # --- 1. 建立特徵 (X) ---
            features = create_pair_features(row_A, row_B)
            
            # --- 2. 建立標籤 (Y)：判斷是否為人工優化配對 ---
            # 判斷標準：如果 A 和 B 被人工排在「同一天、同一人員」下，則視為人工優化配對 (Label=1)。
            is_optimized_paired = (row_A['排程日'] == row_B['排程日']) and \
                                  (row_A['人員'] == row_B['人員']) 
            
            # 標籤：1 表示這對工單被人工優化配對了，0 表示沒有
            features['是否人工配對到同一天且同人員'] = 1 if is_optimized_paired else 0

            # 標籤：是否屬於同一唯一主工單ID
            #features['是否人工配對'] = 1 if row_A['唯一主工單ID'] == row_B['唯一主工單ID'] else 0
            
            all_pairs_data.append(features)

    return pd.DataFrame(all_pairs_data)


if __name__ == "__main__":
    print(f"--- [%s] 1. 載入歷史資料 ---" % datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    
    df_history = None
    for enc in ENCODINGS_TO_TRY:
        try:
            df_history = pd.read_csv(HISTORY_PATH, encoding=enc)
            df_history.columns = df_history.columns.str.strip()
            print(f"✅ 成功使用 {enc} 編碼載入資料")
            break
        except Exception:
            continue

    if df_history is None:
        print(f"❌ 錯誤：無法載入檔案 {HISTORY_PATH}，請檢查檔案是否存在或編碼是否特殊。")
        exit(1)

    print("--- 2. 轉換為工單配對數據集 ---")
    try:
        df_ml_data = create_training_data(df_history)
        
        print("\n標籤分佈檢查:")
        print(df_ml_data['是否人工配對到同一天且同人員'].value_counts())
        
        # 確保輸出目錄存在
        os.makedirs(os.path.dirname(OUTPUT_ML_DATA_PATH), exist_ok=True)
        
        # 5. 儲存 ML 訓練數據
        df_ml_data.drop(columns=['工單編號_A', '工單編號_B'], inplace=True)
        # 💡 使用 utf-8-sig 確保 Excel 打開不亂碼，且後續 train_model 讀取穩定
        df_ml_data.to_csv(OUTPUT_ML_DATA_PATH, index=False, encoding='utf-8-sig')
        
        print(f"\n✅ 成功生成訓練數據: {OUTPUT_ML_DATA_PATH}")
        
    except Exception as e:
        print(f"❌ 轉換過程出錯: {e}")
        exit(1)