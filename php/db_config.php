<?php
// db_config.php

$json_path = "D:/ribbon_schedule/config_ribbon.json";

if (!file_exists($json_path)) {
    die("❌ 找不到資料庫設定檔 (JSON)");
}

$json_data = json_decode(file_get_contents($json_path), true);
$db_config = $json_data['db_config']; // 取得 JSON 內的 db_config 區塊

/*
"db_config": {
        "host": "192.168.117.55",
        "user": "allen",
        "password": "unitech",
        "database": "ribbon_test"
    },
*/

$conn = new mysqli(
    $db_config['host'], 
    $db_config['user'], 
    $db_config['password'], 
    $db_config['database']
);

if ($conn->connect_error) {
    die("❌ 資料庫連線失敗: " . $conn->connect_error);
}

$conn->set_charset("utf8mb4");
?>