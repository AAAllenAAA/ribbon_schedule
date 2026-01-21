<?php
// 強制瀏覽器不要快取此頁面
header("Cache-Control: no-store, no-cache, must-revalidate, max-age=0");
header("Cache-Control: post-check=0, pre-check=0", false);
header("Pragma: no-cache");
header("Expires: Wed, 11 Jan 1984 05:00:00 GMT");
?>
<html>

<head>
    <meta charset="utf-8">
    <title>Ribbon Shedule Artificial Intelligence</title>
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css"
        integrity="sha384-ggOyR0iXCbMQv3Xipma34MD+dH/1fQ784/j6cY/iJTQUOhcWr7x9JvoRxT2MZw1T" crossorigin="anonymous">
    <script src="https://d3js.org/d3.v4.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/billboard.js/dist/billboard.min.js"></script>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/billboard.js/dist/billboard.min.css" />
    <meta http-equiv="content-type" content="application/vnd.ms-excel; charset=UTF-8" />
    <meta http-equiv="refresh" content="1800;url=../login.php">
    <link rel="preconnect" href="https://fonts.gstatic.com">
    <link href="https://fonts.googleapis.com/css2?family=Nunito:wght@300;400;600;700;800&display=swap" rel="stylesheet">
    <link rel="stylesheet" href="../assets/css/bootstrap.css">

    <link rel="stylesheet" href="../assets/vendors/iconly/bold.css">

    <link rel="stylesheet" href="../assets/vendors/perfect-scrollbar/perfect-scrollbar.css">
    <link rel="stylesheet" href="../assets/vendors/bootstrap-icons/bootstrap-icons.css">
    <link rel="stylesheet" href="../assets/css/app.css">

    <script src="https://ajax.googleapis.com/ajax/libs/jquery/2.1.3/jquery.min.js" type="text/javascript"></script>


    <script type="text/javascript">
        function validateFile() {
            var csvInputFile = document.forms["frmCSVImport"]["file"].value;
            if (csvInputFile == "") {
                error = "No source found to import";
                $("#response").html(error).addClass("error");;
                return false;
            }
            return true;
        }

    </script>

    <style>
        /* Loading 遮罩樣式 */
        #loadingOverlay {
            display: none;
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0, 0, 0, 0.7);
            z-index: 9999;
            text-align: center;
            padding-top: 20%;
        }

        .spinner {
            border: 16px solid #f3f3f3;
            border-top: 16px solid #3498db;
            border-radius: 50%;
            width: 120px;
            height: 120px;
            animation: spin 2s linear infinite;
            display: inline-block;
        }

        @keyframes spin {
            0% {
                transform: rotate(0deg);
            }

            100% {
                transform: rotate(360deg);
            }
        }

        .loading-text {
            color: white;
            font-size: 24px;
            margin-top: 20px;
            font-family: "Nunito", sans-serif;
        }
    </style>
</head>

<body style='background-color:#ECFFFF;'>

    <div id="loadingOverlay">
        <div class="spinner"></div>
        <div class="loading-text">排程計算中，請勿關閉視窗...</div>
    </div>

    <header class="mb-3">
        <a href="#" class="burger-btn d-block d-xl-none">
            <i class="bi bi-justify fs-3"></i>
        </a>
    </header>
    <div id="app">


        <div id="sidebar" class="active">
            <div class="sidebar-wrapper active">
                <br>
                <img src="../unitechlogo.png" width="250" height="50" />
                <div class="sidebar-toggler x">
                    <a href="#" class="sidebar-hide d-xl-none d-block"><i class="bi bi-x bi-middle"></i></a>
                </div>
                <div class="sidebar-menu">
                    <ul class="menu">
                        <li class="sidebar-item  has-sub">
                            <a class='sidebar-link'>
                                <i class="bi bi-stack"></i>
                                <span>工單管理</span>
                            </a>
                            <ul class="submenu ">
                                <li class="submenu-item">
                                    <a href='../datatableto.php'>當日工單</a>
                                </li>
                                <li class="submenu-item">
                                    <a href='../datatablewk.php'>當週工單</a>
                                </li>
                                <li class="submenu-item">
                                    <a href='../datatablemo.php'>當月工單</a>
                                </li>
                                <li class="submenu-item">
                                    <a href='../selectwork.php'>區間搜尋</a>
                                </li>
                                <li class="submenu-item">
                                    <a href='../datamanage.php'>管理工單</a>
                                </li>
                            </ul>
                        </li>
                        <li class="sidebar-item  has-sub">
                            <a class='sidebar-link'>
                                <i class="bi bi-collection-fill"></i>
                                <span>匯入</span>
                            </a>
                            <ul class="submenu">
                                <li class="submenu-item">
                                    <a href='../csv-to-mysql/index.php'>匯入工單</a>
                                </li>
                                <li class="submenu-item">
                                    <a href="../csv-to-mysql/import_Desc.php">匯入品名描述</a>
                                </li>
                                <li class="submenu-item">
                                    <a href="../csv-to-mysql/baseinfo.php">匯入基本資料</a>
                                </li>
                                <li class="submenu-item">
                                    <a href="../csv-to-mysql/materimport.php">匯入原料</a>
                                </li>
                                <li class="submenu-item">
                                    <a href="./schedule_index.php">排程</a>
                                </li>
                                <li class='submenu-item'>
                                    <a>匯入其他</a>
                                </li>
                            </ul>
                        </li>
                        <li class="sidebar-item  has-sub">
                            <a class='sidebar-link'>
                                <i class="bi bi-file-earmark-medical-fill"></i>
                                <span>報表</span>
                            </a>
                            <ul class="submenu">
                                <li class="submenu-item">
                                    <a href='../selectreport.php'>匯出製造報表</a>
                                </li>
                                <li class="submenu-item">
                                    <a href='../select_total.php'>匯出總工單報表</a>
                                </li>
                        </li>
                        <li class="submenu-item">
                            <a href='../rawview.php'>原料報表</a>
                        </li>
                    </ul>
                    </li>
                    <li class="sidebar-item">
                        <a onclick="change_password()" class="sidebar-link">
                            <i class="bi bi-file-lock2-fill"></i>
                            <span>修改密碼</span>
                        </a>
                    </li>
                    <li class="sidebar-item">
                        <a onclick="logout()" class="sidebar-link">
                            <i class="bi bi-x-octagon-fill"></i>
                            <span>登出</span>
                        </a>
                    </li>
                    </ul>
                </div>
            </div>
        </div>


        <div id='main'>

            <div class="d-flex justify-content-between align-items-center mb-3">
                <h2>排程</h2>
                <a href="adjust_mapping.php" class="btn btn-outline-dark shadow-sm">
                    <i class="bi bi-gear-fill"></i> 異動搭配調整
                </a>
            </div>
            <form class="form-horizontal" name="frmSchedule" action="schedule_backend.php" method="post"
                enctype="multipart/form-data" onsubmit="return validateFile()">

                <div class="form-group">
                    <label><strong>執行模式：</strong></label>
                    <div class="mt-2">
                        <div class="form-check form-check-inline">
                            <input class="form-check-input" type="radio" name="run_mode" id="mode_daily" value="daily"
                                checked onclick="toggleDateConstraint()"> <label class="form-check-label"
                                for="mode_daily">當日模式</label>
                        </div>
                        <div class="form-check form-check-inline">
                            <input class="form-check-input" type="radio" name="run_mode" id="mode_schedule"
                                value="schedule" onclick="toggleDateConstraint()"> <label class="form-check-label"
                                for="mode_schedule">預排模式</label>
                        </div>
                    </div>
                </div>



                <div class="form-group">
                    <label>匯入電子報檔案：</label>
                    <input type="file" name="file" id="file" class="form-control" accept=".csv,.xls,.xlsx">
                </div>

                <div class="form-group">
                    <label>電子報日期區間：</label>
                    <input type="date" id="start_range" name="start_range" class="form-control" required>
                    <input type="date" id="end_range" name="end_range" class="form-control" required>
                </div>

                <div class="form-group">
                    <label>排程開始日期：</label>
                    <input type="date" id="start_date" name="start_date" class="form-control" required>
                </div>

                <!--
                <div class="form-group">
                    <label><strong>81B 工單負責人 (週)</strong></label>
                    <div class="mt-2">
                        <div class="form-check form-check-inline">
                            <input class="form-check-input" type="radio" name="81B_user" id="userA159" value="A159"
                                <?php //echo isChecked('A159', $saved_user); ?>>
                            <label class="form-check-label" for="userA159">A159 (家偉)</label>
                        </div>
                        <div class="form-check form-check-inline">
                            <input class="form-check-input" type="radio" name="81B_user" id="userA830" value="A830"
                                <?php //echo isChecked('A830', $saved_user); ?>>
                            <label class="form-check-label" for="userA830">A830 (容合)</label>
                        </div>
                        <div class="form-check form-check-inline">
                            <input class="form-check-input" type="radio" name="81B_user" id="userB201" value="B201"
                                <?php //echo isChecked('B201', $saved_user); ?>>
                            <label class="form-check-label" for="userB201">B201 (旺斌)</label>
                        </div>
                    </div>
                </div>
                -->

                <div class="form-group">
                    <label>家偉(A159) 休假區間：</label>
                    <small class="text-muted">(中午 12:00 為"下午12:00")</small>
                    <div id="personA_holidays">
                        <div class="input-group mb-2 holiday-entry shadow-sm p-2 bg-white rounded">
                            <input type="datetime-local" name="personA_start[]" class="form-control mr-1"
                                onfocus="setPresetTime(this, '08:00')">
                            <span class="p-2">至</span>
                            <input type="datetime-local" name="personA_end[]" class="form-control mr-1"
                                onfocus="setPresetTime(this, '17:00')">
                            <div class="input-group-append">
                                <button type="button" class="btn btn-outline-danger btn-sm"
                                    onclick="removeHoliday(this.closest('.holiday-entry'), 'personA_holidays')">
                                    <i class="bi bi-x"></i>
                                </button>
                            </div>
                        </div>
                    </div>
                    <button type="button" class="btn btn-sm btn-secondary" id="addHolidayA"
                        onclick="addHoliday('personA_holidays')" disabled>新增日期區間</button>
                </div>

                <div class="form-group">
                    <label>容合(A830) 休假區間：</label>
                    <small class="text-muted">(中午 12:00 為"下午12:00")</small>
                    <div id="personB_holidays">
                        <div class="input-group mb-2 holiday-entry shadow-sm p-2 bg-white rounded">
                            <input type="datetime-local" name="personB_start[]" class="form-control mr-1"
                                onfocus="setPresetTime(this, '08:00')">
                            <span class="p-2">至</span>
                            <input type="datetime-local" name="personB_end[]" class="form-control mr-1"
                                onfocus="setPresetTime(this, '17:00')">
                            <div class="input-group-append">
                                <button type="button" class="btn btn-outline-danger btn-sm"
                                    onclick="removeHoliday(this.closest('.holiday-entry'), 'personB_holidays')">
                                    <i class="bi bi-x"></i>
                                </button>
                            </div>
                        </div>
                    </div>
                    <button type="button" class="btn btn-sm btn-secondary" id="addHolidayB"
                        onclick="addHoliday('personB_holidays')" disabled>新增日期區間</button>
                </div>

                <div class="form-group">
                    <label>旺斌(B201) 休假區間：</label>
                    <small class="text-muted">(中午 12:00 為"下午12:00")</small>
                    <div id="personC_holidays">
                        <div class="input-group mb-2 holiday-entry shadow-sm p-2 bg-white rounded">
                            <input type="datetime-local" name="personC_start[]" class="form-control mr-1"
                                onfocus="setPresetTime(this, '08:00')">
                            <span class="p-2">至</span>
                            <input type="datetime-local" name="personC_end[]" class="form-control mr-1"
                                onfocus="setPresetTime(this, '17:00')">
                            <div class="input-group-append">
                                <button type="button" class="btn btn-outline-danger btn-sm"
                                    onclick="removeHoliday(this.closest('.holiday-entry'), 'personC_holidays')">
                                    <i class="bi bi-x"></i>
                                </button>
                            </div>
                        </div>
                    </div>
                    <button type="button" class="btn btn-sm btn-secondary" id="addHolidayC"
                        onclick="addHoliday('personC_holidays')" disabled>新增日期區間</button>
                </div>

                <div class="form-group">
                    <button type="submit" id="submit" name="submit" class="btn btn-primary">開始排程</button>
                </div>

            </form>

            <script>
                // 核心邏輯：點擊時如果沒值，自動填入預設時段
                function setPresetTime(input, timeStr) {
                    if (!input.value) {
                        const today = new Date().toISOString().split('T')[0];
                        input.value = today + "T" + timeStr;
                    }
                }

                // 輔助函式：檢查最後一個日期輸入框並控制新增按鈕狀態
                function checkLastInput(containerId) {
                    const container = document.getElementById(containerId);
                    // ⭐ 重要：這裡要改成搜尋 datetime-local
                    const dateInputs = container.querySelectorAll('input[type="datetime-local"]');

                    const buttonId =
                        containerId === 'personA_holidays' ? 'addHolidayA' :
                            containerId === 'personB_holidays' ? 'addHolidayB' :
                                'addHolidayC';
                    const addButton = document.getElementById(buttonId);

                    if (dateInputs.length > 0) {
                        // ⭐ 邏輯修正：我們要檢查每一組的「最後一個框」（即結束時間）是否有值
                        const lastInput = dateInputs[dateInputs.length - 1];

                        // 設定按鈕狀態：最後一格有值才亮起 (disabled = false)
                        addButton.disabled = !lastInput.value;

                        // 確保為新的輸入框綁定監聽事件
                        if (!lastInput.classList.contains('input-listener-added')) {
                            lastInput.addEventListener('change', function () {
                                checkLastInput(containerId);
                            });
                            lastInput.classList.add('input-listener-added');
                        }
                    } else {
                        addButton.disabled = true;
                    }
                }

                // 動態新增休假日輸入框
                function addHoliday(containerId) {
                    const container = document.getElementById(containerId);

                    const buttonId =
                        containerId === 'personA_holidays' ? 'addHolidayA' :
                            containerId === 'personB_holidays' ? 'addHolidayB' :
                                'addHolidayC';
                    const addButton = document.getElementById(buttonId);

                    // 檢查是否有未填寫的欄位 (現在要檢查一對中的結束時間是否填寫)
                    const allInputs = container.querySelectorAll('input[type="datetime-local"]');
                    if (allInputs.length > 0 && !allInputs[allInputs.length - 1].value) {
                        return; // 最後一格沒填，不准新增
                    }

                    // 1. 創建包裝容器 (一組兩格)
                    const wrapperDiv = document.createElement("div");
                    wrapperDiv.className = "input-group mb-1 holiday-entry";

                    // 決定 Prefix (A, B, 或 C)
                    const personPrefix = containerId === 'personA_holidays' ? 'personA' :
                        containerId === 'personB_holidays' ? 'personB' : 'personC';

                    // 2. 建立開始與結束輸入框
                    // 開始時間
                    const inputStart = document.createElement("input");
                    inputStart.type = "datetime-local";
                    inputStart.name = personPrefix + "_start[]";
                    inputStart.className = "form-control";
                    inputStart.placeholder = "開始時間";
                    // 綁定 focus 事件
                    inputStart.onfocus = function () { setPresetTime(this, '08:00'); };

                    // 中間的文字
                    const span = document.createElement("span");
                    span.className = "input-group-text";
                    span.innerText = "至";

                    // 結束時間
                    const inputEnd = document.createElement("input");
                    inputEnd.type = "datetime-local";
                    inputEnd.name = personPrefix + "_end[]";
                    inputEnd.className = "form-control";
                    inputEnd.placeholder = "結束時間";
                    // 綁定 focus 事件
                    inputEnd.onfocus = function () { setPresetTime(this, '17:00'); };

                    // 3. 創建刪除按鈕
                    const deleteButtonWrapper = document.createElement("div");
                    deleteButtonWrapper.className = "input-group-append";
                    const deleteButton = document.createElement("button");
                    deleteButton.type = "button";
                    deleteButton.className = "btn btn-outline-danger btn-sm";
                    deleteButton.innerHTML = `<i class="bi bi-x"></i>`;
                    deleteButton.onclick = function () {
                        removeHoliday(wrapperDiv, containerId);
                    };
                    deleteButtonWrapper.appendChild(deleteButton);

                    // 4. 組裝並加入容器
                    wrapperDiv.appendChild(inputStart);
                    wrapperDiv.appendChild(span);
                    wrapperDiv.appendChild(inputEnd);
                    wrapperDiv.appendChild(deleteButtonWrapper);
                    container.appendChild(wrapperDiv);

                    // 5. 重新禁用按鈕，並為「結束時間」添加監聽器
                    addButton.disabled = true;
                    inputEnd.addEventListener('change', function () {
                        checkLastInput(containerId);
                    });
                }

                function removeHoliday(elementToRemove, containerId) {
                    const container = document.getElementById(containerId);

                    // 移除指定的一組
                    container.removeChild(elementToRemove);

                    // 檢查剩餘數量
                    const entries = container.querySelectorAll('.holiday-entry');

                    // 如果全被刪光了，自動補回一組空的（確保 UI 不會空掉）
                    if (entries.length === 0) {
                        addHoliday(containerId);
                    }

                    // 重新更新新增按鈕狀態
                    checkLastInput(containerId);
                }

                // 表單驗證（驗證匯入檔案）
                function validateFile() {
                    // 原有的檔案檢查
                    var csvInputFile = document.getElementById("file").value;
                    if (csvInputFile == "") {
                        alert("請選擇要匯入的檔案");
                        return false;
                    }

                    // --- 【新增：休假時間邏輯檢查】 ---
                    const containers = ['personA_holidays', 'personB_holidays', 'personC_holidays'];
                    for (let cid of containers) {
                        const container = document.getElementById(cid);
                        const starts = container.querySelectorAll('input[name*="_start"]');
                        const ends = container.querySelectorAll('input[name*="_end"]');

                        for (let i = 0; i < starts.length; i++) {
                            if (starts[i].value && ends[i].value) {
                                const startTime = new Date(starts[i].value);
                                const endTime = new Date(ends[i].value);

                                if (endTime <= startTime) {
                                    const personName = cid.includes('A') ? '家偉' : cid.includes('B') ? '容合' : '旺斌';
                                    alert(`錯誤：[${personName}] 的休假結束時間必須晚於開始時間！\n\n小提醒：\n中午 12 點請選"下午" 12:00\n凌晨 00 點是"上午" 12:00`);
                                    ends[i].focus();
                                    return false; // 擋住不准送出
                                }
                            }
                        }
                    }
                    // 3. 【核心修改】所有檢查通過後，彈出最後確認
                    if (confirm("確認參數正確並開始排程計算？\n(計算過程請勿關閉網頁)")) {
                        // 顯示遮罩
                        document.getElementById("loadingOverlay").style.display = "block";

                        // 鎖定送出按鈕
                        const btn = document.getElementById("submit");
                        if (btn) {
                            btn.innerText = "處理中...";
                        }

                        return true; // 回傳 true，瀏覽器才會真正把資料送到 schedule_backend.php
                    }

                    return false; // 如果按「取消」，就不送出
                }

                // 限制排程開始日期在日期區間內
                const startRange = document.getElementById("start_range");
                const endRange = document.getElementById("end_range");
                const startDate = document.getElementById("start_date");

                // 2. 定義切換限制的函式 (讓 HTML 的 onclick 呼叫)
                function toggleDateConstraint() {
                    // 檢查目前選中的模式
                    const isDaily = document.getElementById("mode_daily").checked;

                    if (isDaily) {
                        // 當日模式：立刻套用目前的區間限制
                        startDate.min = startRange.value;
                        startDate.max = endRange.value;
                    } else {
                        // 預排模式：移除所有日期限制
                        startDate.removeAttribute("min");
                        startDate.removeAttribute("max");
                    }
                }

                // 3. 修改你原本的監聽器
                startRange.addEventListener("change", function () {
                    // 只有在「當日模式」被選中時，才去連動限制
                    if (document.getElementById("mode_daily").checked) {
                        startDate.min = startRange.value;
                    }
                });

                endRange.addEventListener("change", function () {
                    if (document.getElementById("mode_daily").checked) {
                        startDate.max = endRange.value;
                    }
                });

                // 頁面載入/回退時的監聽器 (解決返回時資料殘留問題)
                window.addEventListener('pageshow', function (event) {
                    // 如果頁面是從快取（按返回鍵）載入的
                    if (event.persisted) {
                        // 強制頁面重新整理，這會發送新的 Request 到伺服器，Log 就會出現了
                        window.location.reload();
                    }
                });
            </script>


        </div>
    </div>


</body>
<script>
    function logout() {
        if (confirm("確定要登出嗎?") == true) {
            location.href = "../login.html";
        }
        else { }
    }
    function change_password() {
        if (confirm("確定要更改密碼嗎?") == true) {
            location.href = "../change_password.php";
        }
        else { }
    }
</script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.14.7/umd/popper.min.js"
    integrity="sha384-UO2eT0CpHqdSJQ6hJty5KVphtPhzWj9WO1clHTMGa3JDZwrnQq4sF86dIHNDz0W1"
    crossorigin="anonymous"></script>
<script src="https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/js/bootstrap.min.js"
    integrity="sha384-JjSmVgyd0p3pXB1rRibZUAYoIIy6OrQ6VrjIEaFf/nJGzIxFDsf4x0xIM+B07jRM"
    crossorigin="anonymous"></script>
<script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/js/bootstrap.min.js"></script>
<script src="../assets/vendors/perfect-scrollbar/perfect-scrollbar.min.js"></script>
<script src="../assets/js/bootstrap.bundle.min.js"></script>

<script src="https://cdn.jsdelivr.net/npm/echarts@5.2.2/dist/echarts.min.js"></script>

<script src="../assets/js/main.js"></script>

</html>