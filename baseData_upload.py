import mysql.connector
import pandas as pd
import math

db_config = {
    "host": "192.168.117.55",
    "user": "allen",
    "password": "unitech",
    "database": "ribbon_test"
}

# path about base_df
base_excel_path = r"E:\ribbon_schedule\data\基本資料-20250617-All.xlsx"
df = pd.read_excel(base_excel_path)

print("讀取筆數: ", len(df))


def normalize(value):
    """將空值或 nan 轉為空字串，保留其他字串或數值"""
    if value is None or pd.isna(value):
        return ''
    if isinstance(value, str):
        cleaned = value.strip()
        if cleaned.lower() in ["", "nan", "none", "null"]:
            return ''
        return cleaned
    return value  # 保留數值原樣


def normalize_numeric(value):
    """數值欄位：空值轉空字串，其他轉 float 保留小數"""
    val = normalize(value)
    if val == '':
        return ''
    try:
        return float(val)
    except:
        return val  # 若非數值就直接返回原值


def normalize_area(value):
    """面積欄位四捨五入到整數，空值轉空字串"""
    val = normalize_numeric(value)
    if val == '':
        return ''
    return int(math.floor(float(val) + 0.5))  # 傳統四捨五入

# connect MySQL
conn = mysql.connector.connect(**db_config)
cursor = conn.cursor()

insert_SQL = """
INSERT IGNORE INTO pairingrules
(main_ProductInfo, mi_SN, mi_Width, main_CarNum, 
 1st_ProductInfo, 1st_CarNum, 
 2nd_ProductInfo, 2nd_CarNum, 
 3th_ProductInfo, 3th_CarNum, 
 4th_ProductInfo, 4th_CarNum, 
 mi_Area)
VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
"""

# insert into database
for _, row in df.iterrows():

    data = (
        normalize(row["料號"]),
        normalize(row["原料材質"]),
        normalize_numeric(row["寬度Cm"]),
        normalize(row["車數"]),
        normalize(row["搭1料號"]),
        normalize(row["搭1產出車數"]),
        normalize(row["搭2料號"]),
        normalize(row["搭2產出車數"]),
        normalize(row["搭3料號"]),
        normalize(row["搭3產出車數"]),
        normalize(row["搭4料號"]),
        normalize(row["搭4產出車數"]),
        normalize_area(row["面積  M2"])
    )

    cursor.execute(insert_SQL, data)

conn.commit()
print("Insert success")

cursor.close()
conn.close()
