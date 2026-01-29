<?php
// api_handler.php
header("Content-Type: application/json; charset=UTF-8");

// 1. 引入類別檔案
require_once '..\csv-to-mysql\lib\WorkOrder_Allen.php'; 

// 2. 因為有 Namespace，必須使用完整類別名稱或宣告 use
// 假設您的 WorkOrder.php 第一行是 namespace Phppot;
$db = new \Phppot\WorkOrder(); 

$action = $_POST['action'] ?? '';
$filename = $_POST['filename'] ?? '';

if ($action === 'process_csv' && !empty($filename)) {
    // 定義 Python 產出 CSV 的物理路徑
    $uploadDir = "D:/ribbon_schedule/data/";
    $fullPath = $uploadDir . $filename;

    // 檢查實體檔案是否存在
    if (!file_exists($fullPath)) {
        echo json_encode(["status" => "error", "message" => "伺服器找不到檔案：$filename"]);
        exit;
    }

    // 3. ⭐ 重要：偽造 $_FILES 全域變數 ⭐
    // 這是為了讓不改動的 WorkOrder.php 能抓到檔案路徑與大小
    $_FILES["file"] = [
        "tmp_name" => $fullPath,
        "size" => filesize($fullPath),
        "error" => 0,
        "name" => $filename
    ];

    // 4. 執行原本的方法
    // 注意：您原本的方法不接收參數，它直接讀取 $_FILES
    $result = $db->readCSVRecords(); 

    // 5. 解析 WorkOrder 回傳的結果陣列 ($output)
    if (isset($result['type']) && $result['type'] === 'success') {
        echo json_encode([
            "status" => "success", 
            "message" => "匯入成功！" . $result['message']
        ]);
    } else {
        echo json_encode([
            "status" => "error", 
            "message" => "匯入失敗：" . ($result['message'] ?? '未知錯誤')
        ]);
    }
    exit;
}

// 若非法存取
echo json_encode(["status" => "error", "message" => "無效的請求或參數缺失"]);