// ==UserScript==
// @name         OA 实验检测委托单自动填写 (终极融合加强版)
// @namespace    http://tampermonkey.net/
// @version      7.2
// @description  合并抗异步联动错峰填写、本地 Excel 数据导入及明细表多行联动功能
// @author       YourName
// @match        *://oa.ween-semi.com/spa/workflow/static4form/*
// @updateURL    https://cdn.jsdelivr.net/gh/YuchenJi/my-tampermonkey-scripts@main/autofill_oa_system.user.js
// @downloadURL  https://cdn.jsdelivr.net/gh/YuchenJi/my-tampermonkey-scripts@main/autofill_oa_system.user.js
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// @grant        none
// ==/UserScript==

(function() {
    'use strict';

    // =========================================
    // 工具函数：获取 1 个月后的日期
    // =========================================
    function getOneMonthLater() {
        let date = new Date();
        date.setMonth(date.getMonth() + 1);
        let year = date.getFullYear();
        let month = ("0" + (date.getMonth() + 1)).slice(-2);
        let day = ("0" + date.getDate()).slice(-2);
        return `${year}-${month}-${day}`;
    }

    // =========================================
    // 🌟 自定义翻译字典 (人话 -> 系统底层黑话)
    // 你可以在 Excel 的表头直接写左边的人话，脚本会自动翻译成右边的底层字段
    // =========================================
    const fieldDict = {
        // --- ATR 表单 ---
        "Recipient": "Contact",                     // 左边是你想用的名字，右边是真实的 data-fieldname
        "Telephone": "TEL",
        "Sample deliver address": "Address",
        "Date When Needed": "DateNeeded",
        "Charge Department": "ChargeDept",

        // --- FA 表单 ---
        "车规器件": "cgqj",
        "客户验证失效": "sfwkhyzjdsx",
        "Client Name": "ClientName"
    };

    // =========================================
    // 核心黑科技 1：主表智能寻址填充函数 (带翻译功能)
    // =========================================
    // =========================================
    // 核心黑科技 1：主表智能寻址填充函数 (增加选人框支持)
    // =========================================
    function smartFill(key, value, showName = "") {
        if (!window.WfForm) return;
        let valStr = String(value);

        // 1. 如果直接传入 fieldID
        if (key.toLowerCase().startsWith('field')) {
            window.WfForm.changeFieldValue(key, { value: valStr });
            return;
        }

        // 2. 查字典
        let actualFieldName = fieldDict[key] || key;

        // 3. 寻找格子
        let element = document.querySelector(`.excelMainTable [data-fieldname="${actualFieldName}"]`);
        if (element) {
            let realFieldId = element.getAttribute('data-fieldmark');
            if (realFieldId) {
                // 【核心修复】：如果传入了第三个参数 showName，就按照泛微选人框的特殊格式赋值
                if (showName !== "") {
                    window.WfForm.changeFieldValue(realFieldId, {
                        value: valStr,
                        specialobj: [{ id: valStr, name: showName }] // 告诉系统 UI 该显示什么名字
                    });
                    console.log(`[选人框写入] ${key}(${actualFieldName}) -> ID:${valStr}, Name:${showName}`);
                } else {
                    // 普通文本、下拉框走这里的常规赋值
                    window.WfForm.changeFieldValue(realFieldId, { value: valStr });
                    console.log(`[主表写入] ${key}(${actualFieldName}) -> ${valStr}`);
                }
            }
        } else {
            console.warn(`[跳过] 主表未找到字段: ${key}`);
        }
    }

    // =========================================
    // 核心黑科技 2：明细表智能寻址填充函数 (带翻译与行号解析)
    // =========================================
    function smartFillDetail(key, value, rowIndex = 0) {
        if (!window.WfForm) return;
        let valStr = String(value);

        let actualFieldName = fieldDict[key] || key;

        // 寻找满足 fieldname 且 fieldmark 以后缀 _rowIndex 结尾的明细表元素
        let element = document.querySelector(`[data-fieldname="${actualFieldName}"][data-fieldmark$="_${rowIndex}"]`);
        if (element) {
            let realFieldId = element.getAttribute('data-fieldmark');
            if (realFieldId) {
                window.WfForm.changeFieldValue(realFieldId, { value: valStr });
                console.log(`[明细表第${rowIndex+1}行写入] ${key}(${realFieldId}) -> ${valStr}`);
            }
        }
    }

    // =========================================
    // 功能 1：日常默认自动填表逻辑 (保留了抗异步和联动错峰)
    // =========================================
    function doAutoFill() {
        if (!window.WfForm) {
            alert("错误：未找到 WfForm 接口！");
            return false;
        }

        console.log("====== 开始执行默认自动填表 ======");
        try {
            let currentUrl = window.location.href;

            // ==========================================
            // 表单 A：FA Test Request form
            // ==========================================
            if (currentUrl.includes("workflowid=622")) {
                console.log("识别为: FA Test Request form，执行方案 A");

                // 【第一批次】：填写主字段、文本和复选框
                window.WfForm.changeFieldValue("field8133", { value: "bob.ji@ween-semi.com" });
                window.WfForm.changeFieldValue("field8102", { value: "15901754757" });
                window.WfForm.changeFieldValue("field8123", { value: "1" });
                window.WfForm.changeFieldValue("field8126", { value: "1" });
                window.WfForm.changeFieldValue("field11149", { value: "1" });
                window.WfForm.changeFieldValue("field11959", { value: "1" });
                window.WfForm.changeFieldValue("field8598", { value: "0" });
                window.WfForm.changeFieldValue("field8118", { value: "1" });

                // 关键点：触发联动的“父级”下拉框
                window.WfForm.changeFieldValue("field8100", { value: "1" }); // Client Name
                window.WfForm.changeFieldValue("field8101", { value: "7" }); // Department
                window.WfForm.changeFieldValue("field8145", { value: "0" }); // Address
            }

            // ==========================================
            // 表单 B : ATR
            // ==========================================
            else if (currentUrl.includes("workflowid=629")) {
                console.log("识别为: 第二个表单 ATR，执行方案 B");

                window.WfForm.changeFieldValue("field8316", { value: "1" });
                window.WfForm.changeFieldValue("field8321", { value: "5" });
                window.WfForm.changeFieldValue("field8305", { value: "3" });
                window.WfForm.changeFieldValue("field8339", { value: "1" });
                window.WfForm.changeFieldValue("field8320", { value: "1" });
                window.WfForm.changeFieldValue("field8319", { value: "0" });
                window.WfForm.changeFieldValue("field8338", { value: "1" });
                window.WfForm.changeFieldValue("field11149", { value: "1" });

                window.WfForm.changeFieldValue("field8307", { value: "2" }); // Charge Department

                smartFill("ChipsFrom", "1");
                smartFill("Outline0", "39");
                smartFill("Address", "0");
                smartFill("Contact", "Bob Ji");
                smartFill("TEL", "15901754757");
                smartFill("DateNeeded", getOneMonthLater());

                // 【主表联动到明细表演示】：将主表填好的 DeviceType 抄写到明细表第1行
                let mainDeviceType = window.WfForm.getFieldValue("field8308");
                if (mainDeviceType) {
                    smartFillDetail("DeviceType", mainDeviceType, 0);
                }
            }

            // ==========================================
            // 表单 C : 量产转工程
            // ==========================================
            else if (currentUrl.includes("workflowid=594")) {
                console.log("识别为: 第三个表单 量产转工程ATR，执行方案 C");

                smartFill("companycode", "3");
                smartFill("approver", "283", "James Zhang");
                smartFill("productline", "3");
                smartFill("costcenter", "3");
                smartFill("description", " request wafer_name \n 12NC \n wafer_lot_ID, good die = \n for xxx production in xxx ");

            }

            // 表单 D : 可靠性

            // ==========================================

            else if (currentUrl.includes("workflowid=625")) {
                console.log("识别为: 可靠性 Test Request form，执行方案 D");
                // 【第一批次】：填写主字段、文本和复选框
                smartFill("ClientName", "1");
                smartFill("ClientAdd", "0");
                smartFill("Department", "8");
                smartFill("Phone", "15901754757");
                smartFill("ProductCustomer", "1");
                smartFill("SampleReturn", "1");
                smartFill("Email", "bob.ji@ween-semi.com");
                smartFill("testway", "1");
                smartFill("DataOnly", "1");
                smartFill("Onlydata", "1");
                // 【第二批次】：自动填写明细表 (根据你 HTML 里的结构预设数据)
                // TestItem 对应下拉框的值(42=HTS, 43=HTGB正压, 44=HTGB负压, 110=HTRB)
                // Testpoints 对应下拉框的值(11="168,500,1000hrs")
                let detailData = [
                    //HTSL
                    { TestItem: "42", Qty: "80", Testpoints: "11", Remark: "Tj = 175°C" },
                    //HTGBP
                    { TestItem: "43", Qty: "80", Testpoints: "11", Remark: "Tj = 175°C, Vgs= 22V" },
                    //HTGBN
                    { TestItem: "44", Qty: "80", Testpoints: "11", Remark: "Tj = 175°C, Vgs= -12V" },
                    //HTRB
                    { TestItem: "110", Qty: "80", Testpoints: "11", Remark: "Tj = 175C, VDS=1200V" }
                ];

                const DETAIL_TABLE_ID = "detail_1"; // 泛微明细表1 (对应 groupid=0)

                // 获取当前已有的明细表行数
                let existingRowStr = window.WfForm.getDetailAllRowIndexStr(DETAIL_TABLE_ID);
                let existingRowCount = existingRowStr === "" ? 0 : existingRowStr.split(",").length;

                // 自动点击 "+" 号新增行，补齐到需要的行数

                for (let i = existingRowCount; i < detailData.length; i++) {
                    window.WfForm.addDetailRow(DETAIL_TABLE_ID);
                }

                // 延时 300ms 等待新增的 DOM 行渲染完毕，再向里面填字
                setTimeout(() => {
                    detailData.forEach((rowData, index) => {
                        // 利用已有的 smartFillDetail 函数进行智能填表
                        smartFillDetail("TestItem", rowData.TestItem, index);
                        smartFillDetail("Qty", rowData.Qty, index);
                        smartFillDetail("Testpoints", rowData.Testpoints, index);
                        smartFillDetail("Remark", rowData.Remark, index);
                    });
                    console.log("====== 方案 D 明细表填充完毕 ======");
                }, 300);
            }
            
            console.log("====== 自动填表指令派发完毕 ======");
            return true;
        } catch (error) {
            console.error("填表时发生错误:", error);
        }
    }

    // =========================================
    // 功能 2：读取并解析 Excel 数据 (支持多行自动新增)
    // =========================================
    function handleExcelFile(event, btnElement) {
        let file = event.target.files[0];
        if (!file) return;

        let reader = new FileReader();
        reader.onload = function(e) {
            try {
                let data = new Uint8Array(e.target.result);
                // 使用 SheetJS 解析 Excel
                let workbook = XLSX.read(data, { type: 'array' });
                let firstSheetName = workbook.SheetNames[0];
                let worksheet = workbook.Sheets[firstSheetName];
                let json = XLSX.utils.sheet_to_json(worksheet);

                if (json.length > 0) {
                    console.log("====== 开始解析 Excel 数据 ======", json);

                    const DETAIL_TABLE_ID = "detail_1"; // 泛微系统明细表1 的代号

                    // 遍历每一行 Excel 数据
                    json.forEach((rowData, index) => {
                        // 第一行 (index:0) 既填主表，又填明细表第1行
                        if (index === 0) {
                            for (let key in rowData) {
                                smartFill(key, rowData[key]);         // 试填主表
                                smartFillDetail(key, rowData[key], 0); // 试填明细表
                            }
                        }
                        // 第二行及以上：检查明细表行数是否足够，不够则自动点击“+”号新增
                        else {
                            if (window.WfForm) {
                                let existingRowStr = window.WfForm.getDetailAllRowIndexStr(DETAIL_TABLE_ID);
                                let existingRowCount = existingRowStr === "" ? 0 : existingRowStr.split(",").length;

                                if (existingRowCount <= index) {
                                    window.WfForm.addDetailRow(DETAIL_TABLE_ID);
                                }
                            }

                            // 延时 200ms 等待新增的行渲染完毕，再向里面填字
                            setTimeout(() => {
                                for (let key in rowData) {
                                    smartFillDetail(key, rowData[key], index);
                                }
                            }, 200);
                        }
                    });

                    // 导入成功的按钮动画反馈
                    let originalBg = "#20b082";
                    btnElement.innerHTML = "导入<br>成功";
                    btnElement.style.background = "#1890ff";
                    btnElement.style.transform = "scale(0.95)";
                    setTimeout(() => {
                        btnElement.innerHTML = "导入<br>Excel";
                        btnElement.style.background = originalBg;
                        btnElement.style.transform = "scale(1)";
                    }, 1500);

                } else {
                    alert("Excel 文件中没有读取到有效数据！");
                }
            } catch (err) {
                console.error("Excel 解析失败:", err);
                alert("Excel 解析失败，请确保文件格式正确 (.xlsx)。");
            }
        };
        reader.readAsArrayBuffer(file);
        event.target.value = ''; // 允许重复导入同一个文件
    }

    // =========================================
    // 动作 1：智能轮询 (等待页面加载与网络拉取)
    // =========================================
    let currentUrl = window.location.href;
    let targetFieldId = "";

    if (currentUrl.includes("workflowid=622")) {
        targetFieldId = "field8106";
    } else if (currentUrl.includes("workflowid=562")) {
        targetFieldId = "field8316";
    } else if (currentUrl.includes("workflowid=594")) {
        targetFieldId = "field9372";
    }

    let checkCount = 0;
    let timer = setInterval(function() {
        checkCount++;
        if (window.WfForm && document.getElementById(targetFieldId)) {
            clearInterval(timer);
            console.log("探测到表单实体，耐心等待下拉菜单数据从服务器拉取完毕...");
            // 保持 2秒 延迟，防止被下拉框的异步刷新覆盖
            setTimeout(doAutoFill, 2000);
        }
        if (checkCount > 30) {
            clearInterval(timer);
        }
    }, 500);

    // =========================================
    // 动作 2：创建操作面板 (双按钮 UI)
    // =========================================
    function createUI() {
        if(document.getElementById("my-autofill-panel")) return;

        // 面板容器
        let panel = document.createElement("div");
        panel.id = "my-autofill-panel";
        panel.style.cssText = "position:fixed; top:40%; right:20px; z-index:999999; display:flex; flex-direction:column; gap:12px;";

        // 按钮通用样式
        let btnStyle = "width:54px; height:54px; border:none; border-radius:50%; cursor:pointer; box-shadow:0 4px 8px rgba(0,0,0,0.2); font-weight:bold; font-size:13px; line-height:1.2; color:#fff; transition:all 0.3s;";

        // --- 按钮 1：一键填表 (蓝色) ---
        let btnDefault = document.createElement("button");
        btnDefault.innerHTML = "一键<br>填表";
        btnDefault.style.cssText = btnStyle + "background:#0079DE;";
        btnDefault.onmouseover = function(){ this.style.transform = "scale(1.1)"; };
        btnDefault.onmouseout = function(){ this.style.transform = "scale(1)"; };

        btnDefault.onclick = function(e) {
            e.preventDefault();
            doAutoFill();
            let originalBg = "#0079DE";
            btnDefault.innerHTML = "填表<br>成功";
            btnDefault.style.background = "#20b082";
            btnDefault.style.transform = "scale(0.95)";
            setTimeout(() => {
                btnDefault.innerHTML = "重新<br>填表";
                btnDefault.style.background = originalBg;
                btnDefault.style.transform = "scale(1)";
            }, 1500);
        };

        // --- 按钮 2：导入 Excel (绿色) ---
        let btnExcel = document.createElement("button");
        btnExcel.innerHTML = "导入<br>Excel";
        btnExcel.style.cssText = btnStyle + "background:#20b082;";
        btnExcel.onmouseover = function(){ this.style.transform = "scale(1.1)"; };
        btnExcel.onmouseout = function(){ this.style.transform = "scale(1)"; };

        let fileInput = document.createElement("input");
        fileInput.type = "file";
        fileInput.accept = ".xlsx, .xls, .csv";
        fileInput.style.display = "none";
        fileInput.onchange = (e) => handleExcelFile(e, btnExcel);

        btnExcel.onclick = (e) => {
            e.preventDefault();
            fileInput.click();
        };

        // 组装并推送到页面
        panel.appendChild(btnDefault);
        panel.appendChild(btnExcel);
        panel.appendChild(fileInput);
        document.body.appendChild(panel);
    }

    setTimeout(createUI, 3000);

})();
