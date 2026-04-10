# model_1.py(ml_scorer.py) (排程優化核心模組)
'''
模型服務接口:
核心工作： 載入模型 -> 重現「原料材質」與「跨日銜接」特徵工程 -> 呼叫模型進行預測 -> 返回 AI_Prob 分數。
這是專門為 predict_AI_result.py 服務的模組。它的職責是將工單數據傳遞給 ML 模型並獲取預測結果。
核心工作： 載入 schedule_pairing_classifier.pkl 和 scaler.pkl -> 重現 create_ml_data.py 的特徵工程 -> 呼叫模型進行預測 -> 將 AI_Prob 分數返回給主程式。
'''
import pandas as pd
import joblib
import numpy as np

# ====================================
# 配置區：確保路徑正確
# ====================================
MODEL_PATH = r'E:\ribbon_schedule\training_AI\ML\schedule_pairing_classifier.pkl'
SCALER_PATH = r'E:\ribbon_schedule\training_AI\ML\scaler.pkl'
# ====================================

# ------------------------------
# 載入模型和標準化工具
# ------------------------------
ML_MODEL = None
SCALER = None
MODEL_LOADED = False
try:
    ML_MODEL = joblib.load(MODEL_PATH)
    SCALER = joblib.load(SCALER_PATH)
    MODEL_LOADED = True
    print(f"✅ ML 模型和 Scaler 載入成功。模型路徑: {MODEL_PATH}")
except Exception as e:
    print(f"❌ 警告：無法載入 ML 模型 ({e})。將使用硬性規則作為預測結果。")

# ------------------------------
# 特徵計算函數 (必須和 train_model.py 中的特徵列表順序完全一致)
# ------------------------------
# 這是訓練時使用的特徵列表
FEATURES_LIST = [
    "is_item_identical",       # 優先級 1: 品號完全相同
    "is_cm_identical",         # 優先級 2: 公分數完全相同
    "item_similarity_score",   # 優先級 3: 品號相似度
    "operator_is_same"         # 流程約束: 人員是否相同
]

def product_similarity(p1, p2):
    """計算品號字串相似度，與 create_ml_data.py 中一致"""
    p1, p2 = str(p1), str(p2)
    max_len = max(len(p1), len(p2))
    if max_len == 0:
        return 0.0
    matches = sum(a == b for a, b in zip(p1, p2))
    return matches / max_len

def safe_to_numeric(value):
    """將單一值轉換為數值 如果失敗或缺失 則設為0.0"""
    try:
        num_value = pd.to_numeric(value, errors="coerce")
        return num_value if pd.notna(num_value) else 0.0
    except:
        return 0.0
    
def safe_to_date(value, default_date='2200-01-01'):
    """將數值安全轉換為日期時間物件 如果失敗則返回極遠點的日期 (P1交期)"""
    try:
        date_value = pd.to_datetime(value, erros='coerce')
        return date_value if pd.notna(date_value) else pd.to_datetime(default_date)
    except: 
        return pd.to_datetime(default_date)

def calculate_features(row_A, row_B):
    """根據兩筆工單數據，計算 ML 模型所需的特徵"""

    # 數值處理
    公分_A = safe_to_numeric(row_A.get('公分'))
    公分_B = safe_to_numeric(row_B.get('公分'))
    
    # 字串處理
    品號_A = str(row_A.get('品號', '')).strip()
    品號_B = str(row_B.get('品號', '')).strip()
    
    # ❗ 修正點 2: 更新特徵計算邏輯 ❗
    features = {
        # 優先級 1：品號完全相同
        'is_item_identical': 1 if 品號_A == 品號_B else 0,
        
        # 優先級 2：公分數完全相同
        'is_cm_identical': 1 if 公分_A == 公分_B else 0,
        
        # 優先級 3：品號相似度
        'item_similarity_score': product_similarity(品號_A, 品號_B),
        
        # 流程約束：人員是否相同
        'operator_is_same': 1 if row_A.get('人員') == row_B.get('人員') else 0,
        
        # ❗ '公分數是否相同' 和 '公分數_差' 已被移除 ❗
    }
    
    # 將字典轉換為 DataFrame，確保列順序與訓練時一致 (這一點非常棒)
    df_features = pd.DataFrame([features])
    return df_features[FEATURES_LIST]


# ------------------------------
# 核心預測函數
# ------------------------------
def get_ai_pairing_scores(df_current_schedule):
    """
    接收您的程式排程結果 (DataFrame)，輸出每對潛在工單的 AI 配對機率。
    """
    
    # 處理模型未載入的情況
    if not MODEL_LOADED:
        print("❌ 警告：模型未載入，使用硬性規則作為預測結果。")
        # 建立一個包含所有潛在配對的 DataFrame
        potential_pairs_if_fail = []
        df_schedule = df_current_schedule.copy()
        
        # 為了返回結果，必須先生成潛在配對
        df_schedule.sort_values(by=['預計開工日', '人員'], inplace=True)
        df_schedule.reset_index(drop=True, inplace=True)
        
        for i in range(len(df_schedule)):
            row_A = df_schedule.iloc[i]
            for j in range(i + 1, min(i + 6, len(df_schedule))):
                row_B = df_schedule.iloc[j]
                if row_A['人員'] == row_B['人員']:
                    # 模擬硬性規則分數 (公分數和人員都相同給高分)
                    prob = 0.8 if safe_to_numeric(row_A.get('公分')) == safe_to_numeric(row_B.get('公分')) and row_A['人員'] == row_B['人員'] else 0.1
                    potential_pairs_if_fail.append({
                        '工單編號_A': row_A['工單編號'],
                        '工單編號_B': row_B['工單編號'],
                        'AI_Prob': prob
                    })
        return pd.DataFrame(potential_pairs_if_fail)


    potential_pairs = []
    
    # 步驟 1: 迭代生成潛在配對 (與訓練數據邏輯一致)
    df_schedule = df_current_schedule.copy()
    # 這裡只需要確保排程是按照日期/人員分組，內部的順序已經由主程式決定
    df_schedule.sort_values(by=['預計開工日', '人員'], inplace=True) 
    df_schedule.reset_index(drop=True, inplace=True)

    # ❗ 修正點 3: 限制潛在配對的範圍 (保持與訓練時的 K=10 一致性) ❗
    # 您的訓練腳本使用 K=10，這裡使用 K=5 可能會錯過 AI 學習過的長距離配對。
    # 為了安全起見，我們使用 K=11 (即考慮接下來 10 個位置)
    PREDICTION_WINDOW = 11 

    for i in range(len(df_schedule)):
        row_A = df_schedule.iloc[i]
        
        for j in range(i + 1, min(i + PREDICTION_WINDOW, len(df_schedule))):
            row_B = df_schedule.iloc[j]
            
            # 只有當它們的排程人員相同時，才考慮配對（這是基本優化前提）
            if row_A['人員'] == row_B['人員']:
                features_df = calculate_features(row_A, row_B)
                features_df['工單編號_A'] = row_A['工單編號']
                features_df['工單編號_B'] = row_B['工單編號']
                potential_pairs.append(features_df)
    
    if not potential_pairs:
        return pd.DataFrame(columns=['工單編號_A', '工單編號_B', 'AI_Prob'])

    # 步驟 2: 整合特徵並進行標準化
    df_features_raw = pd.concat(potential_pairs, ignore_index=True)
    df_X = df_features_raw.drop(columns=['工單編號_A', '工單編號_B'])
    
    # 確保特徵列順序一致 (非常重要!)
    df_X = df_X[FEATURES_LIST] 
    
    # 使用訓練時的 scaler 進行標準化
    X_scaled = SCALER.transform(df_X)

    # 步驟 3: ML 模型預測
    # 輸出 (不配對機率, 配對機率)
    probabilities = ML_MODEL.predict_proba(X_scaled) 
    
    # 我們只需要配對機率 (index 1)
    pairing_probs = probabilities[:, 1]
    
    # 步驟 4: 輸出結果
    df_result = df_features_raw[['工單編號_A', '工單編號_B']].copy()
    df_result['AI_Prob'] = pairing_probs
    
    return df_result

# 註：此模組本身不應直接執行，而是被 predict_AI_result.py 引入並呼叫。