@echo off
E:
cd E:\ribbon_schedule\training_AI\ML

:: 設定環境變數，強迫 Python 使用 UTF-8 輸出
set PYTHONIOENCODING=utf-8

echo [%date% %time%] === Start Monthly Retraining === >> result_log.txt

:: 執行程式
python create_ml_data.py >> result_log.txt 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] create_ml_data failed >> pipeline_log.txt
    exit /b %errorlevel%
)

python train_model.py >> result_log.txt 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] train_model failed >> result_log.txt
    exit /b %errorlevel%
)

echo [%date% %time%] === Retraining Completed === >> result_log.txt
echo =============================================================================== >> result_log.txt