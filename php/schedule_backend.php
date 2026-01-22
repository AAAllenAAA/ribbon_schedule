<!DOCTYPE html>
<html lang="zh-Hant">

<head>

    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>排程處理結果</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            margin: 0;
            padding: 0;
            background-color: #f4f7f6;
        }

        .container {
            max-width: 800px;
            margin: 50px auto;
            padding: 30px;
            background-color: #ffffff;
            border-radius: 10px;
            box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
            text-align: center;
        }

        h1 {
            color: #2c3e50;
            border-bottom: 3px solid #3498db;
            padding-bottom: 10px;
            margin-bottom: 30px;
        }

        h3 {
            color: #27ae60;
            margin-top: 20px;
        }

        p {
            line-height: 1.6;
            color: #34495e;
        }

        /* 按鈕容器樣式 */
        .btn-group {
            margin-top: 30px;
            display: flex;
            justify-content: center;
            gap: 20px;
            /* 兩個按鈕的間距 */
            flex-wrap: wrap;
        }

        .btn-download {
            padding: 12px 25px;
            font-size: 17px;
            border: none;
            border-radius: 5px;
            color: white;
            cursor: pointer;
            font-weight: bold;
            transition: all 0.3s;
            text-decoration: none;
            display: inline-block;
        }

        /* Excel 按鈕顏色 (綠色) */
        .btn-excel {
            background-color: #2ecc71;
            box-shadow: 0 4px 6px rgba(46, 204, 113, 0.4);
        }

        .btn-excel:hover {
            background-color: #27ae60;
            box-shadow: 0 6px 8px rgba(46, 204, 113, 0.6);
        }

        /* CSV 按鈕顏色 (藍色) */
        .btn-csv {
            background-color: #3498db;
            box-shadow: 0 4px 6px rgba(52, 152, 219, 0.4);
        }

        .btn-csv:hover {
            background-color: #2980b9;
            box-shadow: 0 6px 8px rgba(52, 152, 219, 0.6);
        }

        .error {
            color: #e74c3c;
            font-weight: bold;
        }

        .btn-upload-data {
            background-color: #f39c12;
            box-shadow: 0 4px 6px rgba(243, 156, 18, 0.4);
        }

        /* 鎖定狀態的樣式 */
        .btn-disabled {
            background-color: #bdc3c7 !important;
            /* 灰色 */
            color: #ffffff !important;
            cursor: not-allowed !important;
            /* 禁止符號游標 */
            box-shadow: none !important;
            opacity: 0.7;
        }
    </style>
</head>

<body>
    <div class="container">
        <h1>排程任務結果</h1>
        <?php

        if (isset($_POST['submit'])) {
            // 1. 設定存檔與讀取路徑
            $uploadDir = "E:/ribbon_schedule/test_report_upload/";
            $uploadDir_json = "E:/ribbon_schedule/test_report_upload/json/";

            if (!is_dir($uploadDir))
                mkdir($uploadDir, 0777, true);
            if (!is_dir($uploadDir_json))
                mkdir($uploadDir_json, 0777, true);

            $filePath = '';

            // 2. 檔案上傳處理 (電子報)
            if (isset($_FILES['file']) && $_FILES['file']['error'] === 0) {
                $fileName = basename($_FILES['file']['name']);
                $fileTmp = $_FILES['file']['tmp_name'];
                $filePath = $uploadDir . $fileName;

                if (move_uploaded_file($fileTmp, $filePath)) {
                    //echo "<p>電子報檔案儲存成功</p>";
                } else {
                    die("<p class='error'>檔案儲存失敗！</p>");
                }
            } else {
                die("<p class='error'>檔案上傳錯誤或未上傳！</p>");
            }

            // 3. 準備 JSON 設定檔傳給 Python
            function formatHolidayData($prefix)
            {
                $starts = $_POST[$prefix . '_start'] ?? [];
                $ends = $_POST[$prefix . '_end'] ?? [];
                $combined = [];

                foreach ($starts as $i => $startTime) {
                    // 只有在開始跟結束都有填值的時候才加入
                    if (!empty($startTime) && !empty($ends[$i])) {

                        $cleanStart = date('Y-m-d H:i', strtotime($startTime));
                        $cleanEnd = date('Y-m-d H:i', strtotime($ends[$i]));

                        $combined[] = [
                            "start" => $cleanStart,
                            "end" => $cleanEnd
                        ];
                    }
                }
                return $combined;
            }

            function cleanDate($dateStr)
            {
                if (empty($dateStr))
                    return '';
                return date('Y-m-d', strtotime($dateStr));
            }

            function addBusinessDays($dateStr, $days)
            {
                $date = new DateTime($dateStr);
                $count = abs($days); // abs = 取絕對值
                $step = $days > 0 ? 1 : -1;

                while ($count > 0) {
                    // modify: PHP DataTime內建方法 專門用於改變日期物件的值
                    // 所以 以下這段程式可以解釋為 加/減 一天
                    $date->modify("$step day");
                    // 檢查是否為週一至週五 (1-5)
                    if ($date->format('N') <= 5) {
                        $count--;
                    }
                }
                return $date->format('Y-m-d');
            }

            $data = [
                //"start_range" => cleanDate($_POST['start_range']) ?? '',
                //"end_range" => cleanDate($_POST['end_range']) ?? '',
                "start_date" => cleanDate($_POST['start_date']) ?? '',
                "week_81B_user" => $_POST['81B_user'] ?? '',

                // 改用處理過的函式抓取資料
                "A159" => formatHolidayData('personA'),
                "A830" => formatHolidayData('personB'),
                "B201" => formatHolidayData('personC'),

                "uploaded_file" => $filePath,
                "run_mode" => $_POST['run_mode'] ?? 'daily'
            ];

            // --- 新增：啟動排隊鎖定機制 ---
            // --- 啟動排隊鎖定機制 ---
            $lockFile = $uploadDir_json . "system_running.lock";
            $fp = fopen($lockFile, "w+");

            if (flock($fp, LOCK_EX)) {
                // 【保護區開始】這裡面的動作一次只能有一個人做
        
                $startDateStr = $data['start_date'];

                if ($data['run_mode'] === 'schedule') {
                    // --- 預排模式：使用 User 從前端輸入的日期 ---
                    // 建議還是用 cleanDate 處理一下確保安全
                    $data['start_range'] = cleanDate($_POST['start_range']) ?? $startDateStr;
                    $data['end_range'] = cleanDate($_POST['end_range']) ?? $startDateStr;
                } else {
                    // --- 當日模式 (daily)：自動計算前後 20 個工作天 ---
                    $data['start_range'] = addBusinessDays($startDateStr, -30);
                    $data['end_range'] = addBusinessDays($startDateStr, 30);
                }

                // A. 寫入 JSON 設定檔
                $jsonConfigFile = $uploadDir_json . "config_data.json";
                file_put_contents($jsonConfigFile, json_encode($data, JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES));

                // B. 執行 Python 程式
                $pythonScript = ($data['run_mode'] == 'daily')
                    ? "E:/ribbon_schedule/schedule_web_v4.py"
                    : "E:/ribbon_schedule/prospective_month.py";

                //$resultJsonFile = $uploadDir_json . "result_info.json";
                $resultJsonFile = ($data['run_mode'] == 'daily')
                    ? $uploadDir_json . "result_info.json"
                    : $uploadDir_json . "result_info_month.json";

                if (file_exists($resultJsonFile)) {
                    unlink($resultJsonFile);
                }

                /*
                $cmd = "python \"$pythonScript\"";
                shell_exec($cmd);
                */
                // 💡 關鍵 1：加上 2>&1，這會把 Python 的報錯訊息一起抓回來
                $cmd = "python \"$pythonScript\" 2>&1";

                // 💡 關鍵 2：改用 exec 抓取詳細輸出的數組
                $output = [];
                $return_var = 0;
                exec($cmd, $output, $return_var);

                // 💡 關鍵 3：如果執行失敗 (status 不等於 0)，顯示錯誤內容
                if ($return_var !== 0) {
                    echo "<h2>❌ 排程執行失敗！</h2>";
                    echo "<p>錯誤代碼 (Status): $return_var</p>";
                    echo "<pre style='background: #000; color: #adff2f; padding: 20px; border-radius: 5px;'>";
                    echo "執行指令: $cmd\n";
                    echo "----------------------------------------\n";
                    if (empty($output)) {
                        echo "完全沒有輸出訊息。請檢查：\n";
                        echo "1. 是否有安裝 Python？\n";
                        echo "2. 指令 python 是否能在終端機執行？\n";
                        echo "3. 檔案路徑 $pythonScript 是否正確？";
                    } else {
                        echo implode("\n", $output);
                    }
                    echo "</pre>";
                    exit; // 停止執行，不要跳轉，讓你看清楚報錯訊息
                }

                // C. 【關鍵】立刻讀取 Python 產出的結果到 PHP 變數
        
                $myResult = null;
                if (file_exists($resultJsonFile)) {
                    $myResult = json_decode(file_get_contents($resultJsonFile), true);
                }

                // D. 執行完畢，釋放鎖定 (讓下一位可以進來寫 config_data.json 了)
                flock($fp, LOCK_UN);
                // 【保護區結束】
            }
            fclose($fp);

            // --- 5. 根據剛才存在 $myResult 變數裡的資料來顯示結果 ---
            if ($myResult && $myResult['status'] === 'success') {
                $excelFilename = basename($myResult['excel_path']);
                $csvFilename = basename($myResult['csv_path']);

                echo "<h3>✔ 排程已完成</h3>";
                echo "<p>完成時間：{$myResult['finish_time']}</p>";
                echo "<div class='btn-group'>";
                echo "<a href='download.php?file=" . urlencode($excelFilename) . "' class='btn-download btn-excel'>📊 下載 Excel 排程</a>";

                if ($data['run_mode'] === 'daily') {
                    echo "<a href='download.php?file=" . urlencode($csvFilename) . "' class='btn-download btn-csv'>📄 下載 CSV 資料</a>";
                    echo "<a href='javascript:void(0)' onclick=\"uploadToSystem('$csvFilename')\" class='btn-download btn-upload-data' id='uploadBtn'>📤 Upload 排程資料</a><br>";

                    echo "<a href='../csv-to-mysql/index.php' class=''>#若有更新csv檔案要upload請點選我#</a>";
                }
                echo "</div>";
            } else {
                echo "<p class='error'>處理失敗或找不到結果檔案，請檢查原始資料或 Python 執行狀況。</p>";
            }
        }
        ?>
        <br>
        <button onclick="location.href='schedule_index.php'" style="
            margin-top:40px;
            padding:10px 25px;
            font-size:16px;
            border: 2px solid #95a5a6; 
            background-color: transparent; 
            color:#34495e;
            border-radius:5px;
            cursor:pointer;
            transition: all 0.3s;
        " onmouseover="this.style.backgroundColor='#ecf0f1'; this.style.color='#2c3e50'"
            onmouseout="this.style.backgroundColor='transparent'; this.style.color='#34495e'">
            返回上一頁
        </button>
    </div>
    <script>
        function uploadToSystem(filename) {
            if (!confirm('確定要將檔案上傳至系統嗎？')) return;

            const btn = document.getElementById('uploadBtn');
            const originalText = btn.innerHTML;

            // 進入處理狀態：變色、改字、禁用
            btn.classList.add('btn-disabled');
            btn.innerHTML = '⏳ 處理中...';
            btn.style.pointerEvents = 'none';

            // 建立要傳送的資料
            const formData = new URLSearchParams();
            formData.append('action', 'process_csv');
            formData.append('filename', filename);

            // 呼叫 RESTful API
            fetch('api_handler.php', {
                method: 'POST',
                headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                body: formData
            })
                .then(response => response.json())
                .then(result => {
                    if (result.status === 'success') {
                        alert('✅ 成功：' + result.message);
                    } else {
                        alert('❌ 失敗：' + result.message);
                    }
                })
                .catch(error => {
                    console.error('Error:', error);
                    alert('連線發生錯誤，請檢查網路或執行路徑是否存在。');
                })
                .finally(() => {
                    // 恢復按鈕狀態
                    btn.classList.remove('btn-disabled');
                    btn.innerHTML = originalText;
                    btn.style.pointerEvents = 'auto';
                });
        }
    </script>
</body>

</html>