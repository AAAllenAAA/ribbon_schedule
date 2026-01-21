<?php
// save_mapping.php
include 'db_config.php';
session_start();
$user_ID = $_SESSION['username'] ?? 'unknown';

$action = $_POST['action'] ?? '';

function writeTextLog($action, $main_p, $details, $user_ID)
{
    $logPath = 'E:\ribbon_schedule\data\data_adjust';
    $logFile = $logPath . '\mapping_change_log.txt';

    if (!is_dir($logPath)) {
        mkdir($logPath, 0777, true);
    }
    date_default_timezone_set('Asia/Taipei');
    $time = date('Y-m-d H:i:s');
    $logContent = "[$time] [$user_ID] [$action] 料號: $main_p | 詳情: $details" . PHP_EOL;

    file_put_contents($logFile, $logContent, FILE_APPEND | LOCK_EX);
}

if ($action == 'move_top') {
    $main_p = $_POST['main_p'];
    $sub_p = $_POST['sub_p'];

    // 1. 先將目標設為 0
    $sql_top = "UPDATE pairingrules SET sort_num = 0 WHERE main_ProductInfo = ? AND 1st_ProductInfo = ?";
    $stmt = $conn->prepare($sql_top);
    $stmt->bind_param("ss", $main_p, $sub_p);
    $stmt->execute();

    // 2. 【核心修正】將初始化變數與更新合併為一條指令
    // 利用 (SELECT @i:=0) 作為臨時表來初始化變數，確保在同一個 query 內完成
    $shuffle_sql = "
        UPDATE pairingrules, (SELECT @i:=0) AS temp 
        SET sort_num = (@i:=@i+1) 
        ORDER BY main_ProductInfo ASC, sort_num ASC 
    ";
    $conn->query($shuffle_sql);

    // 3. 強迫物理重排
    $conn->query("ALTER TABLE pairingrules ORDER BY sort_num ASC");

    $log_details = "執行置頂操作：將子料號 [$sub_p] 移至主料號 [$main_p] 的第一順位";

    writeTextLog("MOVE_TOP", $main_p, $log_details, $user_ID);

    //header("Location: adjust_mapping.php?status=success&msg=已將 $main_p 置頂");
    echo "success";
    exit();
}

if ($action == 'add') {
    $main_p = trim($_POST['main_ProductInfo']);
    $sub1 = trim($_POST['sub_1'] ?? '');
    $sub2 = trim($_POST['sub_2'] ?? '');
    $sub3 = trim($_POST['sub_3'] ?? '');
    $sub4 = trim($_POST['sub_4'] ?? '');
    $main_CarNum = $_POST['main_CarNum'] ?? 0;
    $sub1_Car = (int) ($_POST['sub1_CarNum'] ?? 0);
    $sub2_Car = (int) ($_POST['sub2_CarNum'] ?? 0);
    $sub3_Car = (int) ($_POST['sub3_CarNum'] ?? 0);
    $sub4_Car = (int) ($_POST['sub4_CarNum'] ?? 0);

    // --- 步驟 1: 檢查主料號是否存在於資料庫 ---
    $check_sql = "SELECT * FROM pairingrules WHERE main_ProductInfo = ? LIMIT 1";
    $stmt_check = $conn->prepare($check_sql);
    $stmt_check->bind_param("s", $main_p);
    $stmt_check->execute();
    $result = $stmt_check->get_result();

    if ($result->num_rows == 0) {
        // 如果找不到料號，跳回並彈窗
        echo "error_not_found";
        exit();
    }

    // 撈出原始資料
    $orig_data = $result->fetch_assoc();

    // --- 步驟 2: 執行新增 ---
    // 將新規則的 sort_num 設為 0
    $mi_SN = $orig_data['mi_SN'];
    $mi_Width = $orig_data['mi_Width'];
    $mi_Area = $orig_data['mi_Area'];

    /*
    // 定義一個內部函式來查子料號寬度
    function fetchWidth($conn, $part_no)
    {
        if (empty($part_no))
            return 0;
        $q = "SELECT mi_Width FROM main_part_table WHERE main_ProductInfo = ? LIMIT 1";
        $st = $conn->prepare($q);
        $st->bind_param("s", $part_no);
        $st->execute();
        $r = $st->get_result()->fetch_assoc();
        return $r ? (float) $r['mi_Width'] : 0;
    }

    // 查出所有子料號寬度
    $w1 = fetchWidth($conn, $sub1);
    $w2 = fetchWidth($conn, $sub2);
    $w3 = fetchWidth($conn, $sub3);
    $w4 = fetchWidth($conn, $sub4);
    // 若輸入的料號是空 則 車數就是 0
    if ($w1 == 0) {
        $sub1_car = 0;
    }
    if ($w2 == 0) {
        $sub2_car = 0;
    }
    if ($w3 == 0) {
        $sub3_car = 0;
    }
    if ($w4 == 0) {
        $sub4_car = 0;
    }

    // --- 步驟 3: 計算主車數 (main_CarNum) 把非0的寬度帶進來計算 ---
    // 尚未完成
    $sub1_car = 0;
    $sub2_car = 0;
    $sub3_car = 0;
    $sub4_car = 0;
    $main_CarNum = 0;
    */

    // insert
    $sql_ins = "INSERT INTO pairingrules (
    `main_ProductInfo`, `mi_SN`, `mi_Width`, `main_CarNum`, 
    `1st_ProductInfo`, `1st_CarNum`, 
    `2nd_ProductInfo`, `2nd_CarNum`, 
    `3th_ProductInfo`, `3th_CarNum`, 
    `4th_ProductInfo`, `4th_CarNum`, 
    `mi_Area`, `sort_num`
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0)";

    $stmt_ins = $conn->prepare($sql_ins);
    if (!$stmt_ins) {
        die("SQL 準備失敗: " . $conn->error); // 萬一欄位名還是錯，這行會告訴你
    }


    $stmt_ins->bind_param(
        "ssdisisisisis",
        $main_p,
        $mi_SN,
        $mi_Width,
        $main_CarNum,
        $sub1,
        $sub1_Car,
        $sub2,
        $sub2_Car,
        $sub3,
        $sub3_Car,
        $sub4,
        $sub4_Car,
        $mi_Area
    );

    $stmt_ins->execute();

    // --- 步驟 3: 重新洗牌 (確保新加入的 0 變為 1) ---
    $shuffle_sql = "
        UPDATE pairingrules, (SELECT @i:=0) AS temp 
        SET sort_num = (@i:=@i+1) 
        ORDER BY main_ProductInfo ASC, sort_num ASC
    ";

    if ($conn->query($shuffle_sql)) {
        // --- 步驟 4: 物理排序 (將資料表結構依據 sort_num 重排) ---
        $conn->query("ALTER TABLE pairingrules ORDER BY sort_num ASC");

        $log_details = "主車數: $main_CarNum, 子1: $sub1($sub1_Car), 子2: $sub2($sub2_Car)";

        // 2. 呼叫函式寫入 .txt
        writeTextLog("ADD", $main_p, $log_details, $user_ID);

        echo "success";
    } else {
        echo "Error during shuffle: " . $conn->error;
    }

    exit();
}

if ($action == 'delete') {
    $main_p = $_POST['main_p'];
    $sub_p = $_POST['sub_p'];

    $sql_del = "DELETE FROM pairingrules WHERE main_ProductInfo = ? AND 1st_ProductInfo = ?";
    $stmt = $conn->prepare($sql_del);
    $stmt->bind_param("ss", $main_p, $sub_p);
    $stmt->execute();

    header("Location: adjust_mapping.php?status=deleted");
    exit();
}
?>