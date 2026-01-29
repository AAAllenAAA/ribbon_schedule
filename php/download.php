<?php
$baseDir = "D:/ribbon_schedule/data/";  // 你的 Excel 檔案放置位置

if (!isset($_GET['file'])) {
    die("沒有指定下載檔案");
}

$file = basename($_GET['file']);    // 防止跳脫路徑
$filePath = $baseDir . $file;

if (!file_exists($filePath)) {
    die("找不到檔案：$filePath");
}

// --- 強制停用快取 ---
header("Cache-Control: no-cache, must-revalidate"); 
header("Pragma: no-cache"); 
header("Expires: 0");

// 設定下載 header
header("Content-Type: application/octet-stream");
header("Content-Disposition: attachment; filename=\"$file\"");
header("Content-Length: " . filesize($filePath));

readfile($filePath);
exit;
?>
