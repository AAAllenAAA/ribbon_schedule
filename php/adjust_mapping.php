<!DOCTYPE html>
<html lang="zh-TW">

<head>
    <meta charset="UTF-8">
    <title>異動搭配調整</title>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css">
    <style>
        .form-section {
            background: #CAFFFF;
            padding: 20px;
            border-radius: 8px;
            margin-bottom: 30px;
            border: 1px solid #dee2e6;
        }
    </style>
</head>

<body class="container mt-5">

    <div class="d-flex justify-content-between align-items-center mb-4">
        <h2>🛠 異動搭配調整</h2>
        <a href="schedule_index.php" class="btn btn-outline-secondary">返回</a>
    </div>

    <div class="form-section">
        <h4>🔍 搜尋主料號</h4>
        <div class="row g-3">
            <div class="col-md-9">
                <input type="text" id="search_main_part" class="form-control" placeholder="輸入主料號關鍵字（注意連同料號最後的'.'也需輸入）">
            </div>
            <div class="col-md-3">
                <button type="button" onclick="searchMapping()" class="btn btn-info w-100">搜尋</button>
            </div>
        </div>
    </div>



    <h4>📋 目前已設定之搭配 (由上而下代表優先順序)</h4>
    <div id="mapping_result">
        <div class="alert alert-secondary">請輸入主料號進行搜尋...</div>
    </div>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script>
        function searchMapping() {
            var keyword = $('#search_main_part').val().trim();
            if (keyword === "") {
                alert("請輸入主料號進行搜尋");
                return;
            }

            $.ajax({
                url: 'get_mapping.php',
                type: 'GET',
                data: { keyword: keyword },
                success: function (response) {
                    // 1. 顯示搜尋結果清單
                    $('#mapping_result').html(response);

                    // 2. 顯示新增區塊，並自動填入主料號
                    $('#add_section').show();
                    $('input[name="main_ProductInfo"]').val(keyword);
                },
                error: function () {
                    alert('搜尋失敗，請檢查資料庫連線');
                }
            });
        }

        $(document).ready(function () {
            // 當按下「確認新增」按鈕
            $('#btn_save').on('click', function () {
                // 基本防呆：檢查主車數
                var carNum = $('input[name="main_CarNum"]').val();
                if (!carNum || carNum <= 0) {
                    alert("請輸入正確的主工單車數！");
                    return;
                }

                // 確認視窗
                if (!confirm('確定要新增此筆搭配嗎？')) return;

                var btn = $(this);
                btn.prop('disabled', true).text('存檔並排序中...');

                // 收集表單資料
                var formData = $('#add_form').serialize();
                formData += "&action=add"; // 告知後端執行新增動作

                $.ajax({
                    url: 'save_mapping.php',
                    type: 'POST',
                    data: formData,
                    success: function (response) {
                        // response 應該是後端 echo 出來的 "success"
                        if (response.trim() === "success") {
                            alert('新增成功！資料表已完成重新排序。');

                            // 刷新上方清單
                            searchMapping();

                            // 清空表單內容 (除了唯讀的主料號)
                            $('#add_form')[0].reset();
                            var keyword = $('#search_main_part').val();
                            $('input[name="main_ProductInfo"]').val(keyword);
                        } else {
                            // 如果後端報錯，顯示錯誤訊息
                            alert('存檔失敗：' + response);
                        }
                    },
                    error: function () {
                        alert('系統錯誤，請洽管理員');
                    },
                    complete: function () {
                        btn.prop('disabled', false).text('確認新增');
                    }
                });
            });
        });

        // 置頂功能 (維持你原本的 AJAX，成功後呼叫 searchMapping 即可)
        function moveTop(mainP, subP) {
            if (!confirm('確定要將此配對置頂嗎？')) return;
            $.ajax({
                url: 'save_mapping.php',
                type: 'POST',
                data: { action: 'move_top', main_p: mainP, sub_p: subP },
                success: function () {
                    searchMapping(); // 成功後自動刷新
                }
            });
        }
    </script>

    <div class="form-section" id="add_section" , style="display: none;">
        <h4>➕ 新增搭配關係</h4>
        <form id="add_form" class="row g-3" autocomplete="off">
            <div class="col-md-12">
                <label class="form-label fw-bold text-danger">主工單料號</label>
                <input type="text" name="main_ProductInfo" class="form-control" placeholder="請輸入料號 (注意連同料號最後的'.'也需輸入)"
                    readonly required>
                <label class="form-label fw-bold text-danger">主工單料號車數 (必填)</label>
                <input type="text" name="main_CarNum" class="form-control" placeholder="請輸入車數" required>
            </div>
            <div class="col-md-3">
                <label class="form-label">子料號 1 </label>
                <input type="text" name="sub_1" class="form-control">
            </div>
            <div class="col-md-3">
                <label class="form-label">子料號 2 </label>
                <input type="text" name="sub_2" class="form-control">
            </div>
            <div class="col-md-3">
                <label class="form-label">子料號 3 </label>
                <input type="text" name="sub_3" class="form-control">
            </div>
            <div class="col-md-3">
                <label class="form-label">子料號 4 </label>
                <input type="text" name="sub_4" class="form-control">
            </div>
            <div class="col-md-3">
                <label class="form-label">子料號車數 1 </label>
                <input type="text" name="sub1_CarNum" class="form-control">
            </div>
            <div class="col-md-3">
                <label class="form-label">子料號車數 2 </label>
                <input type="text" name="sub2_CarNum" class="form-control">
            </div>
            <div class="col-md-3">
                <label class="form-label">子料號車數 3 </label>
                <input type="text" name="sub3_CarNum" class="form-control">
            </div>
            <div class="col-md-3">
                <label class="form-label">子料號車數 4 </label>
                <input type="text" name="sub4_CarNum" class="form-control">
            </div>
            <div class="col-md-12 text-end mt-3">
                <button type="button" id="btn_save" class="btn btn-primary px-5">確認新增</button>
            </div>
        </form>
    </div>

</body>

</html>