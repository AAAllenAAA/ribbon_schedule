# train_model.py
'''
模型訓練:
負責載入 create_ml_data.py 產生的訓練數據集，
並使用機器學習演算法（如您的分類器）進行學習。 
核心工作： 執行訓練、選擇最佳參數、並將訓練好的模型和數據預處理工具保存下來。 
輸出檔案： schedule_pairing_classifier.pkl (模型主體) 和 scaler.pkl (標準化工具)。
核心工作： 載入新版 ML_Training.csv -> 學習「原料材質」與「跨日銜接」權重 -> 儲存模型與 Scaler。
'''
import pandas as pd
import joblib
import os
from datetime import datetime
from sklearn.model_selection import train_test_split
from xgboost import XGBClassifier
from sklearn.metrics import f1_score, classification_report, confusion_matrix
from sklearn.preprocessing import StandardScaler
from imblearn.over_sampling import SMOTE # 處理數據不平衡問題

# ====================================
# 配置區：請根據您的環境調整
# ====================================
DATA_PATH = r"E:\ribbon_schedule\training_AI\ML\ML_Training.csv"
MODEL_OUTPUT_PATH = r"E:\ribbon_schedule\training_AI\ML\schedule_pairing_classifier.pkl"
SCALER_OUTPUT_PATH = r'E:\ribbon_schedule\training_AI\ML\scaler.pkl' # 標準化工具儲存路徑
BACKUP_DIR = r"E:\ribbon_schedule\training_AI\ML\backup" #歷史模型存放處
TARGET_COLUMN = '是否人工配對到同一天且同人員'
# ====================================

def train_and_save_model():
    print(f"--- [%s] 1. 載入數據 ---" % datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    try:
        data = pd.read_csv(DATA_PATH)
    except FileNotFoundError:
        print(f"❌ 錯誤：找不到數據檔案 {DATA_PATH}。請先執行 create_ml_data.py！")
        return

    # 1. 定義特徵 (X) 和標籤 (Y)
    # 這裡的特徵必須和 create_ml_data.py 中生成的欄位一致！
    FEATURES = [
        "is_item_identical",       # 優先級 1: 品號完全相同
        "is_cm_identical",         # 優先級 2: 公分數完全相同
        "item_similarity_score",   # 優先級 3: 品號相似度
        "operator_is_same"         # 流程約束: 人員是否相同
    ]
    
    X = data[FEATURES]
    Y = data[TARGET_COLUMN]

    print("--- 2. 數據標準化 (StandardScaler) ---")

    # 2. 數據標準化 (Standardization)
    # 對於樹模型（如 XGBoost），標準化並非必須，但對許多數值特徵（如差異、分數）可能有幫助。
    scaler = StandardScaler()
    X_scaled = scaler.fit_transform(X)
    X_scaled = pd.DataFrame(X_scaled, columns=X.columns)
    
    # 3. 劃分訓練集和測試集
    X_train, X_test, Y_train, Y_test = train_test_split(X_scaled, Y, test_size=0.2, random_state=42, stratify=Y)
    
    # 4. 處理數據不平衡 (如果 '是否人工配對' 為 0 和 1 的數量差距很大)
    # SMOTE 只在訓練集上執行，避免數據洩露 (Data Leakage)
    if Y.value_counts()[1] / len(Y) < 0.2: # 如果正樣本(1)少於20%，則考慮SMOTE
        print("--- 3. 偵測到數據不平衡，開始 SMOTE 過採樣 ---")
        try:
            smote = SMOTE(random_state=42)
            X_train_res, Y_train_res = smote.fit_resample(X_train, Y_train)
            print(f"原始訓練集大小: {len(X_train)} -> 處理後大小: {len(X_train_res)}")
        except ImportError:
            print("💡 提示：SMOTE 需要安裝 imblearn 庫: pip install imbalanced-learn")
            X_train_res, Y_train_res = X_train, Y_train
    else:
        X_train_res, Y_train_res = X_train, Y_train
        

    print("--- 4. 開始 XGBoost 分類模型訓練 ---")
    
    # 5. 初始化和訓練模型
    
    # 假設正樣本佔總樣本的比例為 p_ratio
    # 權重計算: (負樣本數 / 正樣本數)
    neg_count = Y_train_res.value_counts().get(0, 1)
    pos_count = Y_train_res.value_counts().get(1, 1)
    
    # 避免除以零
    scale_weight = neg_count / pos_count
    
    # 5. 初始化和訓練模型
    model = XGBClassifier(
        objective='binary:logistic',  # 二元分類目標
        n_estimators=300,             # 增加迭代次數: 弱學習器的數量。數值越大，模型越強大，但訓練時間越久，過擬合風險越高
        learning_rate=0.05,           # 較低的學習率，更穩定: 學習步長。控制模型對錯誤的修正力度。較小的學習率（如 0.05）通常需要更多的 n_estimators，但結果更穩定、泛化能力可能更好。
        max_depth=5,                  # 樹的深度: 每棵樹的最大深度。控制模型的複雜度。深度越大，模型越容易記住訓練數據（過擬合）。5 是一個平衡複雜度和泛化能力的保守值。
        random_state=42,              # 確保訓練過程的隨機性可重現。
        eval_metric='logloss',        # 評估指標，用於訓練過程中衡量模型的性能，logloss 適用於機率輸出的二元分類
        #use_label_encoder=False,      # 這是 XGBoost 為了未來版本兼容性建議關閉的參數。
        scale_pos_weight=scale_weight # 設定正樣本權重 (提高模型對 '配對' 的敏感度)
    )
    
    model.fit(X_train_res, Y_train_res)
    print("✅ 模型訓練完成！")

    # 6. 模型評估
    print("--- 5. 模型評估 ---")
    Y_pred = model.predict(X_test)
    
    # 關鍵指標：F1-Score
    f1 = f1_score(Y_test, Y_pred)
    print(f"F1-Score (平衡配對準確度): {f1:.4f}")
    
    # ... (後續評估報告保持不變) ...
    print("\n混淆矩陣 (Confusion Matrix):")
    print(confusion_matrix(Y_test, Y_pred))
    
    print("\n分類報告:")
    print(classification_report(Y_test, Y_pred))

    # --- 7. 儲存與備份邏輯 ---
    # 確保資料夾存在
    os.makedirs(BACKUP_DIR, exist_ok=True)
    
    # 產生帶有日期的備份檔名
    timestamp = datetime.now().strftime("%Y%m%d")
    backup_model_path = os.path.join(BACKUP_DIR, f"classifier_{timestamp}.pkl")
    backup_scaler_path = os.path.join(BACKUP_DIR, f"scaler_{timestamp}.pkl")

    # 1. 儲存供主程式使用的正式版 (直接覆蓋)
    joblib.dump(model, MODEL_OUTPUT_PATH)
    joblib.dump(scaler, SCALER_OUTPUT_PATH)
    
    # 2. 儲存一份帶時間戳記的備份版
    joblib.dump(model, backup_model_path)
    joblib.dump(scaler, backup_scaler_path)
    
    print(f"✅ 模型已存: {MODEL_OUTPUT_PATH}")
    print(f"✅ 標準化工具 (scaler) 已存: {SCALER_OUTPUT_PATH}")
    print(f"📦 歷史備份已存: {backup_model_path}")
    print("--- 接下來可以進行模型整合 (model_1.py) ---")

if __name__ == "__main__":
    train_and_save_model()