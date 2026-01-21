<?php
// get_mapping.php
include 'db_config.php';

$keyword = $_GET['keyword'] ?? '';

// 依照 sort_num 排序，確保顯示順序
$sql = "SELECT * FROM pairingrules WHERE main_ProductInfo LIKE ? ORDER BY sort_num ASC";
$stmt = $conn->prepare($sql);
$search = "%$keyword%";
$stmt->bind_param("s", $search);
$stmt->execute();
$result = $stmt->get_result();

if ($result->num_rows > 0) {
    echo '<table class="table table-striped table-hover mt-3">';
    echo '<thead class="table-dark">
            <tr>
                <th width="7%">優先序</th>
                <th width="15%">主工單料號</th>
                <th width="8%">主車數</th>
                <th width="30%">搭配子料號組合 (1~4)</th>
                <th width="25%">子料號車數 (1~4)</th>
                <th width="15%">功能</th>
            </tr>
          </thead><tbody>';

    while ($row = $result->fetch_assoc()) {
        // --- 處理子料號顯示邏輯 ---
        $subs = [];
        if (!empty($row['1st_ProductInfo'])) $subs[] = $row['1st_ProductInfo'];
        if (!empty($row['2nd_ProductInfo'])) $subs[] = $row['2nd_ProductInfo'];
        if (!empty($row['3th_ProductInfo'])) $subs[] = $row['3th_ProductInfo'];
        if (!empty($row['4th_ProductInfo'])) $subs[] = $row['4th_ProductInfo'];
        $carnum = [];
        if ($row['1st_CarNum'] != 0) $carnum[] = $row['1st_CarNum'];
        if ($row['2nd_CarNum'] != 0) $carnum[] = $row['2nd_CarNum'];
        if ($row['3th_CarNum'] != 0) $carnum[] = $row['3th_CarNum'];
        if ($row['4th_CarNum'] != 0) $carnum[] = $row['4th_CarNum'];
        
        // 用 " + " 串接多個子料號
        $sub_display = implode(" <b class='text-primary'>+</b> ", array_map('htmlspecialchars', $subs));
        $sub_car_display = implode(" <b class='text-primary'>, </b> ", array_map('htmlspecialchars', $carnum));

        echo "<tr>";
        echo "<td><span class='badge bg-secondary'>{$row['sort_num']}</span></td>";
        echo "<td><b>" . htmlspecialchars($row['main_ProductInfo']) . "</b></td>";
        echo "<td><span class='badge bg-info text-dark'>{$row['main_CarNum']}</span></td>";
        echo "<td>" . $sub_display . "</td>";
        echo "<td>" . $sub_car_display . "</td>";
        echo "<td>
                <div class='btn-group' role='group'>
                    <button type='button' class='btn btn-warning btn-sm' 
                        onclick='moveTop(\"" . htmlspecialchars($row['main_ProductInfo']) . "\", \"" . htmlspecialchars($row['1st_ProductInfo']) . "\")'>
                        ⭐ 置頂
                    </button>
                </div>
              </td>";
        echo "</tr>";
    }
    echo '</tbody></table>';
} else {
    echo '<div class="alert alert-warning">找不到符合 "' . htmlspecialchars($keyword) . '" 的主品號規則。</div>';
}
?>