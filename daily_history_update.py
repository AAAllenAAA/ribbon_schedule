import os
import json
import pandas as pd
import mysql.connector
import datetime
from mysql.connector import Error

def load_config(file_path):
    """
    載入指定的 JSON 設定檔。
    """
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        print(f"❌ 找不到設定檔: {file_path}")
        return None
    except json.JSONDecodeError:
        print(f"❌ 設定檔格式錯誤 (JSON Decode Error): {file_path}")
        return None
    except Exception as e:
        print(f"❌ 載入設定檔時發生未知錯誤 {file_path}: {e}")
        return None

def fetch_data_from_db(db_config):
    """
    使用給定的資料庫設定連線並讀取 pairingrules 資料表。
    """
    conn = None
    try:
        today = datetime.date.today()
        today_date = today.strftime('%Y-%m-%d')
        #today_date = '2026-0-03'
        print("🟢 嘗試連線資料庫...")
        conn = mysql.connector.connect(**db_config)
        
        # SQL 查詢語句
        sql_base_df = f"""
            SELECT 
                wo_ForceSDate, woToUserUID, wo_SN, wo_ProductInfo, wo_TotalNum, 
                wo_Width, wo_CarNum, wo_CutNum, wo_ForceProNum, wo_ForceEDate 
            FROM 
                `workorder` 
            WHERE 
                wo_ForceSDate = '{today_date}'
        """
        # 使用 Pandas 讀取資料
        print("🟢 讀取 pairingrules 資料表中...")
        base_df = pd.read_sql(sql_base_df, conn)
        
        # 1. 確保 wo_ForceSDate 是日期格式 (避免有些是字串有些是日期)
        base_df['wo_ForceSDate'] = pd.to_datetime(base_df['wo_ForceSDate'])
        base_df['wo_ForceEDate'] = pd.to_datetime(base_df['wo_ForceEDate'])
        
        # 2. 強制轉換成你想要的格式：年/月/日 (不補零則用 %-m, %-d，Windows 環境有時需用 %#m, %#d)
        #base_df['wo_ForceSDate'] = base_df['wo_ForceSDate'].dt.strftime('%Y/%m/%d')
        #base_df['wo_ForceEDate'] = base_df['wo_ForceEDate'].dt.strftime('%Y/%m/%d')
        base_df['wo_ForceSDate'] = base_df['wo_ForceSDate'].dt.strftime('%Y/%#m/%#d')
        base_df['wo_ForceEDate'] = base_df['wo_ForceEDate'].dt.strftime('%Y/%#m/%#d')
        
        # 3. 接著再把所有欄位轉字串
        base_df = base_df.astype(str)
        
        print(f"✅ 資料成功讀取。總共 {len(base_df)} 筆資料。")
        return base_df
    
    except Error as e:
        print(f"❌ 資料庫操作失敗: MySQL Error: {e}")
        return None
    except Exception as e:
        print(f"❌ 從資料庫讀取基本資料失敗: {e}")
        return None
    finally:
        # 確保連線被關閉
        if conn and conn.is_connected():
            conn.close()
            print("🟢 資料庫連線已關閉。")


def append_date_to_csv(df, history_path):
    try:
        # 檢查檔案是否存在，以便決定是否要寫入 header
        file_exists = os.path.exists(history_path)
        
        print(f"🟢 嘗試將 {len(df)} 筆資料追加到 CSV 檔案: {history_path}...")
        
        df.to_csv(
            history_path,
            mode='a',              # 關鍵設定：使用 'a' (append) 模式進行追加
            header=not file_exists,  # 如果檔案不存在 (第一次寫入)，則寫入 header；否則不寫入
            index=False,           # 不將 DataFrame 的索引寫入 CSV 檔案
            encoding='utf-8'       # 確保編碼正確，以支援中文
        )
        
        print(f"✅ 成功將資料追加到歷史 CSV 檔案: {history_path}。")
        return True
    except Exception as e:
        print(f"X 將資料寫入檔案失敗: {e}")
        return False


def main():
    # 1. 處理路徑和載入 JSON 設定
    
    # 取得腳本所在目錄
    script_dir = os.path.dirname(os.path.abspath(__file__)) 
    
    # config_ribbon.json 放在腳本同目錄下
    config_ribbon_path = os.path.join(script_dir, "config_ribbon.json")
    config = load_config(config_ribbon_path)
    if not config:
        return

    print("✅ JSON 設定檔載入成功。")
    print("-" * 30)

    # 2. 檢查資料庫設定
    db_config = config.get("db_config")
    if not db_config:
        print("❌ config_ribbon.json 缺少 db_config 設定")
        return
    
    # 3. 連線資料庫並讀取資料
    base_df = fetch_data_from_db(db_config)
    
    if base_df is None:
        print("❌ 程式終止，無法取得所需的資料。")
        return
    
    # 4. (可在此處繼續使用 base_df 進行後續處理)
    #print(base_df)
    history_csv_path = config.get("schedule_history_path")
    append_date_to_csv(base_df, history_csv_path)

if __name__ == "__main__":
    main()