// ========================================
// 弁当サイズの表示順（共通で使う）
// ========================================
const SIZE_ORDER = ["特大", "大", "中", "小"];

// ========================================
// 「〇日目の〇〇」→ Excel列番号 の対応表
// ========================================
const DAY_COLUMNS = {
    1: [
        { label: "出欠", col: 3 },
        { label: "昼食", col: 4 },
        { label: "おやつ", col: 5 },
        { label: "夕食", col: 6 },
        { label: "宿泊", col: 7 }
    ],
    2: [
        { label: "出欠", col: 8 },
        { label: "朝食", col: 9 },
        { label: "昼食", col: 10 },
        { label: "おやつ", col: 11 },
        { label: "夕食", col: 12 },
        { label: "宿泊", col: 13 }
    ],
    3: [
        { label: "出欠", col: 14 },
        { label: "朝食", col: 15 },
        { label: "昼食", col: 16 },
        { label: "おやつ", col: 17 }
    ]
};

// ========================================
// Excel列の設定
// ========================================
const COL = {
    NAME: 2,
    START: 3,
    END: 17,
    REASON: 19,
    LUNCH_SIZE: 20
};

// ==============================
// 編集モード
// ==============================
let editMode = false;

let changeHistory = [];

// ========================================
// 読み込んだExcelデータを一時保存
// ========================================
let selectedRows = null;
let selectedFileName = "";
let selectedDays = 3;
let dayLabels = {1: "1日目", 2: "2日目", 3: "3日目"};
// ★追加：直近の検索結果に対する弁当サイズ内訳（無ければnull）
let lastLunchSizeSummary = null;

// ★追加：出欠状態のリアルタイム監視を後で止められるようにしておく変数
let unsubscribeAttendanceListener = null;


// ========================================
// Excelファイル選択
// ========================================
document
    .getElementById("excel-file")
    .addEventListener("change", function (event) {

        const file = event.target.files[0];
        if (!file) return;

        selectedFileName = file.name;

        const reader = new FileReader();

        reader.onload = function (e) {

            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: "array" });
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];

            const rows = XLSX.utils.sheet_to_json(
                worksheet,
                { header: 1, defval: "", raw: false }
            );

            selectedRows = rows;

            document.getElementById("day-select-modal").style.display = "flex";
        };

        reader.readAsArrayBuffer(file);
    });


// ========================================
// 合宿日数選択
// ========================================
document
    .querySelectorAll(".day-select-buttons button")
    .forEach(function (button) {

        button.addEventListener("click", async function () {   // ★async追加

            const days = Number(button.dataset.days);
            selectedDays = days;

            document.getElementById("day-select-modal").style.display = "none";

            await saveSheetToFirestore(selectedRows, selectedFileName, days);

        });
    });


// ========================================
// Excelの内容をWeb表に表示
// ========================================
async function displayAttendance(rows, days) {   // ★async化

    createTableHeader(days);
    displayDateFromFileName(selectedFileName);

    const tbody = document.getElementById("attendance-body");
    tbody.innerHTML = "";

    changeHistory = [];
    updateUndoButton();

    let startCol;

    if (days === 1) {
        startCol = 14;
    } else if (days === 2) {
        startCol = 8;
    } else {
        startCol = 3;
    }

    const endCol = 17;

    for (let i = 2; i < rows.length; i++) {

        const row = rows[i];
        const name = row[COL.NAME];

        if (name === "計") break;
        if (!name) continue;

        const tr = document.createElement("tr");

        const nameCell = document.createElement("td");
        nameCell.className = "sticky-name";
        nameCell.textContent = name;
        tr.appendChild(nameCell);

        const lunchCell = document.createElement("td");
        lunchCell.className = "sticky-lunch";
        lunchCell.textContent = row[COL.LUNCH_SIZE] || "-";
        tr.appendChild(lunchCell);

        for (let col = startCol; col <= endCol; col++) {

            const cell = document.createElement("td");
            const value = row[col];

            if (value == 1) {
                cell.textContent = "〇";
                cell.className = "present";
            } else {
                cell.textContent = "-";
            }

            cell.addEventListener("click", function () {

                if (!editMode) return;

                const previousText = cell.textContent;

                if (cell.textContent === "〇") {
                    cell.textContent = "-";
                    cell.classList.remove("present");
                } else {
                    cell.textContent = "〇";
                    cell.classList.add("present");
                }

                recordChange(cell, previousText, cell.textContent);

                calculateTotals();
                saveAttendanceState();
            });

            tr.appendChild(cell);
        }

        const reasonCell = document.createElement("td");
        const reason = row[COL.REASON];
        reasonCell.textContent = reason || "-";
        tr.appendChild(reasonCell);

        tbody.appendChild(tr);
    }

    calculateTotals();

    populateRosterNamesDatalist();

    // ★変更：無条件保存ではなく、Firestoreとの同期を開始する
    await initAttendanceSync();
}


// ========================================
// 表のヘッダーを作成
// ========================================
function createTableHeader(days) {

    const thead = document.querySelector(".attendance-table thead");
    thead.innerHTML = "";

    const firstRow = document.createElement("tr");

    const nameHeader = document.createElement("th");
    nameHeader.rowSpan = 2;
    nameHeader.className = "sticky-name";
    nameHeader.textContent = "名前";
    firstRow.appendChild(nameHeader);

    const lunchHeader = document.createElement("th");
    lunchHeader.rowSpan = 2;
    lunchHeader.className = "sticky-lunch";
    lunchHeader.textContent = "弁当サイズ";
    firstRow.appendChild(lunchHeader);

    if (days === 1) {
        createDayHeader(firstRow, "day3-title", "3日目", 4);
    } else if (days === 2) {
        createDayHeader(firstRow, "day2-title", "2日目", 6);
        createDayHeader(firstRow, "day3-title", "3日目", 4);
    } else {
        createDayHeader(firstRow, "day1-title", "1日目", 5);
        createDayHeader(firstRow, "day2-title", "2日目", 6);
        createDayHeader(firstRow, "day3-title", "3日目", 4);
    }

    const reasonHeader = document.createElement("th");
    reasonHeader.rowSpan = 2;
    reasonHeader.textContent = "理由";
    firstRow.appendChild(reasonHeader);

    thead.appendChild(firstRow);

    const secondRow = document.createElement("tr");

    if (days === 1) {
        addSubHeaders(secondRow, ["出欠", "朝食", "昼食", "おやつ"]);
    } else if (days === 2) {
        addSubHeaders(secondRow, ["出欠", "朝食", "昼食", "おやつ", "夕食", "宿泊"]);
        addSubHeaders(secondRow, ["出欠", "朝食", "昼食", "おやつ"]);
    } else {
        addSubHeaders(secondRow, ["出欠", "昼食", "おやつ", "夕食", "宿泊"]);
        addSubHeaders(secondRow, ["出欠", "朝食", "昼食", "おやつ", "夕食", "宿泊"]);
        addSubHeaders(secondRow, ["出欠", "朝食", "昼食", "おやつ"]);
    }

    thead.appendChild(secondRow);
}

function createDayHeader(row, id, text, colspan) {
    const th = document.createElement("th");
    th.id = id;
    th.colSpan = colspan;
    th.textContent = text;
    row.appendChild(th);
}

function addSubHeaders(row, headers) {
    headers.forEach(function (text) {
        const th = document.createElement("th");
        th.textContent = text;
        row.appendChild(th);
    });
}


// ========================================
// 「計」を計算
// ========================================
function calculateTotals() {

    const tbody = document.getElementById("attendance-body");

    const oldTotal = document.getElementById("total-row");
    if (oldTotal) oldTotal.remove();

    // ★弁当サイズ内訳の古い行も、いったんまとめて削除
    document.querySelectorAll(".lunch-size-row").forEach(function (row) {
        row.remove();
    });

    const rows = Array.from(tbody.querySelectorAll("tr:not(#total-row):not(.lunch-size-row)"))
        .filter(function (row) { return row.style.display !== "none"; });

    if (rows.length === 0) return;

    const firstRow = rows[0];
    const cells = firstRow.querySelectorAll("td");
    const attendanceColumnCount = cells.length - 3;

    const totals = Array(attendanceColumnCount).fill(0);

    rows.forEach(function (row) {
        const cells = row.querySelectorAll("td");
        for (let i = 0; i < attendanceColumnCount; i++) {
            const cell = cells[i + 2];
            if (cell && cell.textContent === "〇") {
                totals[i]++;
            }
        }
    });

    const totalRow = document.createElement("tr");
    totalRow.id = "total-row";

    const totalName = document.createElement("td");
    totalName.className = "sticky-name";
    totalName.textContent = "計";
    totalRow.appendChild(totalName);

    const totalLunch = document.createElement("td");
    totalLunch.className = "sticky-lunch";
    totalLunch.textContent = "-";
    totalRow.appendChild(totalLunch);

    totals.forEach(function (total) {
        const cell = document.createElement("td");
        cell.textContent = total;
        totalRow.appendChild(cell);
    });

    const totalReason = document.createElement("td");
    totalReason.textContent = "-";
    totalRow.appendChild(totalReason);

    tbody.appendChild(totalRow);

    calculateLunchSizeTotals();
}


// ========================================
// 「計」行の下に、弁当サイズ別の人数を1行でまとめて表示
// ========================================
function calculateLunchSizeTotals() {

    const tbody = document.getElementById("attendance-body");

    // 今表示されている（フィルター等で隠れていない）人の行だけ対象にする
    const rows = Array.from(tbody.querySelectorAll("tr:not(#total-row):not(.lunch-size-row)"))
        .filter(function (row) { return row.style.display !== "none"; });

    if (rows.length === 0) return;

    // サイズごとに人数を数える
    const sizeCounts = {};

    rows.forEach(function (row) {

        const lunchCell = row.querySelector(".sticky-lunch");
        if (!lunchCell) return;

        const size = lunchCell.textContent.trim();

        // 空欄や「-」は集計に含めない
        if (!size || size === "-") return;

        sizeCounts[size] = (sizeCounts[size] || 0) + 1;
    });

    if (Object.keys(sizeCounts).length === 0) return;

    // ========================================
    // 弁当サイズの表示順（共通で使う）
    // ========================================
    const SIZE_ORDER = ["特大", "大", "中", "小"];

    // 指定した順番でサイズを並び替える（Excel側の表記ゆれで
    // sizeOrderに無いサイズがあれば、最後にまとめて追加）
    const sortedSizes = Object.keys(sizeCounts).sort(function (a, b) {

        let indexA = SIZE_ORDER.indexOf(a);
        let indexB = SIZE_ORDER.indexOf(b);

        // sizeOrderに無いサイズは一番後ろに回す
        if (indexA === -1) indexA = SIZE_ORDER.length;
        if (indexB === -1) indexB = SIZE_ORDER.length;

        return indexA - indexB;
    });

    // 「大：10人、中：8人、小：3人」のような文字列を作る
    const summaryText = sortedSizes
        .map(function (size) {
            return `${size}：${sizeCounts[size]}人`;
        })
        .join("　　");

    // 列数を「計」行に合わせる
    const totalRow = document.getElementById("total-row");
    if (!totalRow) return;

    const columnCount = totalRow.querySelectorAll("td").length;

    // ====================================
    // 1行だけ作る
    // ====================================
    const tr = document.createElement("tr");
    tr.className = "lunch-size-row";

    const labelCell = document.createElement("td");
    labelCell.className = "sticky-name";
    labelCell.textContent = "弁当内訳";
    tr.appendChild(labelCell);

    const summaryCell = document.createElement("td");
    summaryCell.colSpan = columnCount - 1;   // 名前列を除いた残り全部を1セルにまとめる
    summaryCell.textContent = summaryText;
    summaryCell.style.textAlign = "left";
    tr.appendChild(summaryCell);

    tbody.appendChild(tr);
}

// ========================================
// 列番号から「出欠」「昼食」などの項目名を逆引き
// ========================================
function getItemLabelForCol(col) {

    for (const day in DAY_COLUMNS) {
        const found = DAY_COLUMNS[day].find(function (item) {
            return item.col === col;
        });
        if (found) return found.label;
    }

    return null;
}


// ========================================
// 指定した名前リストの、弁当サイズ内訳を計算
// ========================================
function calculateLunchSizeBreakdownForNames(names) {

    const sizeCounts = {};

    names.forEach(function (name) {

        for (let i = 2; i < selectedRows.length; i++) {

            const row = selectedRows[i];

            if (row[COL.NAME] === name) {

                const size = (row[COL.LUNCH_SIZE] || "").toString().trim();

                if (size && size !== "-") {
                    sizeCounts[size] = (sizeCounts[size] || 0) + 1;
                }

                break;
            }
        }
    });

    if (Object.keys(sizeCounts).length === 0) return null;

    const sortedSizes = Object.keys(sizeCounts).sort(function (a, b) {

        let indexA = SIZE_ORDER.indexOf(a);
        let indexB = SIZE_ORDER.indexOf(b);

        if (indexA === -1) indexA = SIZE_ORDER.length;
        if (indexB === -1) indexB = SIZE_ORDER.length;

        return indexA - indexB;
    });

    return sortedSizes.map(function (size) {
        return `${size}：${sizeCounts[size]}人`;
    });
}


// ========================================
// ファイル名から日付を表示
// ========================================
function displayDateFromFileName(fileName) {

    const name = fileName.replace(/\.[^/.]+$/, "");

    const match = name.match(/^\d{2}-(\d{2})(\d{2}(?:\.\d{2})*)/);

    if (!match) {
        document.getElementById("file-name").textContent = "〇月〇日";
        return;
    }

    const month = Number(match[1]);
    const days = match[2].split(".").map(Number);

    const dateText = `${month}月` + days.map(day => `${day}日`).join("、");
    document.getElementById("file-name").textContent = dateText;

    const day1 = document.getElementById("day1-title");
    const day2 = document.getElementById("day2-title");
    const day3 = document.getElementById("day3-title");

    if (day1) day1.textContent = "1日目";
    if (day2) day2.textContent = "2日目";
    if (day3) day3.textContent = "3日目";

    if (days.length === 1) {

        if (day3) {
            day3.textContent = `${month}月${days[0]}日`;
            dayLabels[3] = `${month}月${days[0]}日`;
        }

    } else if (days.length === 2) {

        if (day2) day2.textContent = `${month}月${days[0]}日`;
        if (day3) day3.textContent = `${month}月${days[1]}日`;

        dayLabels[2] = `${month}月${days[0]}日`;
        dayLabels[3] = `${month}月${days[1]}日`;

    } else if (days.length === 3) {

        if (day1) day1.textContent = `${month}月${days[0]}日`;
        if (day2) day2.textContent = `${month}月${days[1]}日`;
        if (day3) day3.textContent = `${month}月${days[2]}日`;

        dayLabels[1] = `${month}月${days[0]}日`;
        dayLabels[2] = `${month}月${days[1]}日`;
        dayLabels[3] = `${month}月${days[2]}日`;
    }
}


// ==============================
// ①名前検索 ＋ フィルター選択（統合）
// ==============================
function updateAttendanceRowVisibility() {

    const keyword = document.getElementById("search-input").value.trim().toLowerCase();
    const filterIndex = document.getElementById("attendance-filter-select").value;
    const filterMembers = filterIndex !== "" ? filters[Number(filterIndex)].members : null;

    const rows = document.querySelectorAll("#attendance-body tr:not(#total-row)");

    rows.forEach(function (row) {

        const nameCell = row.querySelector(".sticky-name");
        if (!nameCell) return;

        const name = nameCell.textContent;
        const nameLower = name.toLowerCase();

        const matchesSearch = keyword === "" || nameLower.includes(keyword);
        const matchesFilter = !filterMembers || filterMembers.includes(name);

        row.style.display = (matchesSearch && matchesFilter) ? "" : "none";
    });

    calculateTotals();
}

document.getElementById("search-input").addEventListener("input", updateAttendanceRowVisibility);
document.getElementById("attendance-filter-select").addEventListener("change", updateAttendanceRowVisibility);


// ==============================
// フィルター選択肢を更新（①・③共通）
// ==============================
function populateFilterSelects() {

    const selects = [
        document.getElementById("attendance-filter-select"),
        document.getElementById("member-filter-select")
    ];

    selects.forEach(function (select) {

        if (!select) return;

        const currentValue = select.value;

        select.innerHTML =
            '<option value="">' +
            (select.id === "attendance-filter-select" ? "フィルターなし" : "フィルターを使わない") +
            "</option>";

        filters.forEach(function (filter, index) {
            const option = document.createElement("option");
            option.value = index;
            option.textContent = filter.name;
            select.appendChild(option);
        });

        if (Number(currentValue) < filters.length) {
            select.value = currentValue;
        }
    });
}


// ==============================
// 編集モード切り替え
// ==============================
document
    .getElementById("edit-mode-button")
    .addEventListener("click", function () {

        if(!editMode){
            const password = prompt("パスワードを入力してください");
            if(password !== "2525"){
                alert("パスワードが違います");
                return;
            };
        }

        editMode = !editMode;

        if (editMode) {
            this.textContent = "🔓 編集モード：ON";
            this.classList.add("editing");

            changeHistory = [];
            document.querySelectorAll(".cell-changed").forEach(function (cell) {
                cell.classList.remove("cell-changed");
            });
            updateUndoButton();
        } else {
            this.textContent = "🔒 編集モード：OFF";
            this.classList.remove("editing");
        }
    });


// ==============================
// タブ切り替え
// ==============================
document.querySelectorAll(".tab-button").forEach(function (button) {

    button.addEventListener("click", function () {

        document.querySelectorAll(".tab-button").forEach(function (btn) {
            btn.classList.remove("active");
        });

        this.classList.add("active");

        document.querySelectorAll(".tab-content").forEach(function (content) {
            content.classList.remove("active");
        });

        const tabId = this.dataset.tab;
        document.getElementById(tabId).classList.add("active");
    });
});


// ==============================
// フィルター管理
// ==============================
let filters = []; // 中身はFirestoreから受け取って入れる

const { collection, doc, setDoc, deleteDoc, onSnapshot, getDoc } = window.firestoreFns; // ★getDocを追加
const db = window.db;

const filtersRef = collection(db, "filters");

// Firestoreの変更をリアルタイムで受け取る
onSnapshot(filtersRef, function (snapshot) {

    filters = [];

    snapshot.forEach(function (docSnap) {
        filters.push({
            id: docSnap.id,
            name: docSnap.data().name,
            members: docSnap.data().members
        });
    });

    displayFilters();
    populateFilterSelects();
});


// ==============================
// フィルター一覧を表示
// ==============================
function displayFilters() {

    const filterList = document.getElementById("filter-list");
    if (!filterList) return;

    filterList.innerHTML = "";

    if (filters.length === 0) {
        filterList.innerHTML = "<p>まだフィルターがありません。</p>";
        return;
    }

    filters.forEach(function (filter) {

        const item = document.createElement("div");
        item.className = "filter-item";

        item.innerHTML = `
            <div class="filter-item-header">
                <div>
                    <span class="filter-item-name"></span>
                    <span class="filter-item-count"></span>
                </div>
                <span class="filter-arrow">▼</span>
            </div>
            <div class="filter-members-list"></div>
            <button class="delete-filter-button" data-id="${filter.id}">削除</button>
        `;

        item.querySelector(".filter-item-name").textContent = filter.name;
        item.querySelector(".filter-item-count").textContent = filter.members.length + "人";

        const membersList = item.querySelector(".filter-members-list");

        filter.members.forEach(function (member) {
            const div = document.createElement("div");
            div.className = "filter-member";
            div.textContent = member;
            membersList.appendChild(div);
        });

        filterList.appendChild(item);
    });

    document.querySelectorAll(".filter-item-header").forEach(function (header) {
        header.addEventListener("click", function () {
            const item = this.closest(".filter-item");
            item.classList.toggle("open");
        });
    });

    document.querySelectorAll(".delete-filter-button").forEach(function (button) {
        button.addEventListener("click", async function () {

            const id = this.dataset.id;
            const target = filters.find(function (f) { return f.id === id; });
            const filterName = target ? target.name : "";

            const result = confirm(`「${filterName}」を削除しますか？`);
            if (!result) return;

            await deleteDoc(doc(db, "filters", id));
        });
    });

    populateFilterSelects();
}
// ★以前ここにあった重複コード（未使用のfilterList参照ブロック、
//   重複した削除ボタン・開閉クリックの処理）はすべて削除しました。


// ==============================
// フィルター保存
// ==============================
document
    .getElementById("save-filter-button")
    .addEventListener("click", async function () {

        const nameInput = document.getElementById("filter-name");
        const membersInput = document.getElementById("filter-members");

        const filterName = nameInput.value.trim();

        const members = membersInput.value
            .split(/\r?\n/)
            .map(function (name) { return name.trim(); })
            .filter(function (name) { return name !== ""; });

        if (filterName === "") {
            alert("フィルター名を入力してください。");
            return;
        }

        if (members.length === 0) {
            alert("名簿を入力してください。");
            return;
        }

        const newDocRef = doc(filtersRef);

        await setDoc(newDocRef, {
            name: filterName,
            members: members
        });

        nameInput.value = "";
        membersInput.value = "";

        alert("フィルターを保存しました。");
    });

displayFilters();


// ==============================
// ③メンバー一覧：条件行
// ==============================
function addConditionRow() {

    const row = document.createElement("div");
    row.className = "condition-row";

    row.innerHTML = `
        <select class="condition-day"></select>
        <select class="condition-item"></select>
        <select class="condition-state">
            <option value="1">〇（参加）</option>
            <option value="0">-（不参加）</option>
        </select>
        <button type="button" class="remove-condition-button">✕</button>
    `;

    document.getElementById("condition-list").appendChild(row);

    const daySelect = row.querySelector(".condition-day");
    const itemSelect = row.querySelector(".condition-item");

    function updateItemOptions() {

        const day = Number(daySelect.value);
        itemSelect.innerHTML = "";

        DAY_COLUMNS[day].forEach(function (item) {
            const option = document.createElement("option");
            option.value = item.col;
            option.textContent = item.label;
            itemSelect.appendChild(option);
        });
    }

    daySelect.addEventListener("change", updateItemOptions);

    populateDaySelect(daySelect);
    updateItemOptions();

    row.querySelector(".remove-condition-button").addEventListener("click", function () {
        row.remove();
    });
}

document.getElementById("add-condition-button").addEventListener("click", addConditionRow);

addConditionRow();


// ==============================
// ③メンバー一覧：絞り込み実行
// ==============================
document
    .getElementById("run-member-search-button")
    .addEventListener("click", function () {

        if (!selectedRows) {
            alert("先に①でExcelを取り込んでください。");
            return;
        }

        const conditions = [];

        document.querySelectorAll(".condition-row").forEach(function (row) {
            const col = Number(row.querySelector(".condition-item").value);
            const state = row.querySelector(".condition-state").value;
            conditions.push({ col: col, state: state });
        });

        let matchedNames = [];

        for (let i = 2; i < selectedRows.length; i++) {

            const row = selectedRows[i];
            const name = row[COL.NAME];

            if (name === "計") break;
            if (!name) continue;

            let ok = true;

            conditions.forEach(function (cond) {
                const isPresent = row[cond.col] == 1;
                if (cond.state === "1" && !isPresent) ok = false;
                if (cond.state === "0" && isPresent) ok = false;
            });

            if (ok) matchedNames.push(name);
        }

        const filterIndex = document.getElementById("member-filter-select").value;

        if (filterIndex !== "") {

            const filterMembers =
                filters[Number(filterIndex)].members;

            matchedNames = matchedNames.filter(function (name) {
                return filterMembers.includes(name);
            });
        }

        // ★条件の中に「昼食」が含まれているか判定
        const hasLunchCondition = conditions.some(function (cond) {
            return getItemLabelForCol(cond.col) === "昼食";
        });

        // ★含まれていれば弁当サイズを集計、無ければnull
        lastLunchSizeSummary = hasLunchCondition
            ? calculateLunchSizeBreakdownForNames(matchedNames)
            : null;

        displayMemberResult(matchedNames);
    });


// ==============================
// ③メンバー一覧：結果表示
// ==============================
function displayMemberResult(names) {

    const list = document.getElementById("member-result-list");
    const count = document.getElementById("member-result-count");

    list.innerHTML = "";
    count.textContent = names.length;

    if (names.length === 0) {
        list.innerHTML = "<p>該当するメンバーがいません。</p>";
        return;
    }

    names.forEach(function (name) {

        const div = document.createElement("div");

        div.className = "member-result-item";
        div.textContent = name;

        list.appendChild(div);
    });

    /* ★弁当サイズ内訳があれば、同じリストの一番下に続けて表示*/
    if (lastLunchSizeSummary) {

        const summaryDiv = document.createElement("div");
        summaryDiv.className = "member-result-item member-lunch-summary-item";
        summaryDiv.textContent = "【弁当サイズ内訳】" + lastLunchSizeSummary.join("、");

        list.appendChild(summaryDiv);
    }
}


// ==============================
// ③メンバー一覧：コピー
// ==============================
document
    .getElementById("copy-members-button")
    .addEventListener("click", function () {

        const items =
            document.querySelectorAll(
                "#member-result-list .member-result-item:not(.member-lunch-summary-item)"
            );

        if (items.length === 0) {
            alert("コピーするメンバーがいません。");
            return;
        }

        const names =
            Array.from(items).map(function (item) {
                return item.textContent;
            });

        let text = names.join("\n");

        // ★弁当サイズ内訳があれば、末尾に追記
        if (lastLunchSizeSummary) {
            text += "\n\n【弁当サイズ内訳】\n" + lastLunchSizeSummary.join("\n");
        }

        navigator.clipboard.writeText(text)
            .then(function () {
                alert("メンバーをコピーしました。");
            })
            .catch(function () {
                alert("コピーに失敗しました。お使いのブラウザがコピー機能に対応していない可能性があります。");
            });
    });


// ========================================
// 今の合宿日数で、実際にデータがある「日」の番号一覧
// ========================================
function getActiveDayNumbers() {
    if (selectedDays === 1) return [3];
    if (selectedDays === 2) return [2, 3];
    return [1, 2, 3];
}


// ========================================
// 日プルダウンに、現在の日付ラベルをセットする
// ========================================
function populateDaySelect(selectEl) {

    selectEl.innerHTML = "";

    getActiveDayNumbers().forEach(function (dayNum) {
        const option = document.createElement("option");
        option.value = dayNum;
        option.textContent = dayLabels[dayNum];
        selectEl.appendChild(option);
    });
}


// ========================================
// すでにある条件行の「日」プルダウンを最新の日付に更新
// ========================================
function refreshConditionDayOptions() {

    document.querySelectorAll(".condition-row").forEach(function (row) {

        const daySelect = row.querySelector(".condition-day");
        const previousValue = daySelect.value;

        populateDaySelect(daySelect);

        if (getActiveDayNumbers().includes(Number(previousValue))) {
            daySelect.value = previousValue;
        }

        daySelect.dispatchEvent(new Event("change"));
    });
}


// ====================================
// 出欠データの保存先ID（ファイル名から拡張子を除いたもの）
// ====================================
function getAttendanceDocId() {
    return selectedFileName.replace(/\.[^/.]+$/, "");
}


// ====================================
// Firestoreと出欠データを同期開始する
// ====================================
async function initAttendanceSync() {

    // 前のファイルの監視を止める（重複監視の防止）
    if (unsubscribeAttendanceListener) {
        unsubscribeAttendanceListener();
        unsubscribeAttendanceListener = null;
    }

    const docId = getAttendanceDocId();
    const docRef = doc(db, "attendanceStates", docId);

    // すでに保存されているデータがあるか確認
    const snap = await getDoc(docRef);

    if (snap.exists()) {

        // 保存済みのデータがあれば、それを表に反映する
        applyAttendanceMarks(snap.data().marks);

    } else {

        // まだ無ければ、今のExcelの内容を初期値として保存する
        await saveAttendanceState();
    }

    // これ以降の変更をリアルタイムで受け取る
    unsubscribeAttendanceListener = onSnapshot(docRef, function (docSnap) {

        if (!docSnap.exists()) return;

        applyAttendanceMarks(docSnap.data().marks);
    });
}


// ====================================
// 保存済みの〇/-のデータを、今表示している表に反映する
// ====================================
function applyAttendanceMarks(marks) {

    const tbody = document.getElementById("attendance-body");
    const rows = tbody.querySelectorAll("tr:not(#total-row)");

    rows.forEach(function (row, rowIndex) {

        const rowMarks = marks[rowIndex];

        // 行数が合わない（Excelが違う等）場合はスキップ
        if (!rowMarks) return;

        const cells = row.querySelectorAll("td");

        for (let i = 0; i < rowMarks.length; i++) {

            const cell = cells[i + 2];
            if (!cell) continue;

            const shouldBePresent = rowMarks[i];
            const isPresent = cell.textContent === "〇";

            if (shouldBePresent && !isPresent) {
                cell.textContent = "〇";
                cell.classList.add("present");
            } else if (!shouldBePresent && isPresent) {
                cell.textContent = "-";
                cell.classList.remove("present");
            }
        }
    });

    calculateTotals();
}


// ====================================
// 現在の出席表データをFirestoreに保存
// ====================================
async function saveAttendanceState() {

    if (!selectedFileName) return;

    const tbody = document.getElementById("attendance-body");
    const rows = tbody.querySelectorAll("tr:not(#total-row)");

    const state = {};

    rows.forEach(function (row, rowIndex) {

        const cells = row.querySelectorAll("td");
        const marks = [];

        for (let i = 2; i < cells.length - 1; i++) {
            marks.push(cells[i].textContent === "〇");
        }

        state[rowIndex] = marks;
    });

    const docId = getAttendanceDocId();

    await setDoc(doc(db, "attendanceStates", docId), {
        marks: state
    });
}

// ==============================
// 共有シート（読み込んだExcel全体）の管理
// ==============================
const currentSheetRef = doc(db, "appState", "current");

async function saveSheetToFirestore(rows, fileName, days) {

    const docId = fileName.replace(/\.[^/.]+$/, "");

    // Firestoreは配列を直接ネストできないため、JSON文字列にして保存
    await setDoc(doc(db, "sheets", docId), {
        fileName: fileName,
        rowsJson: JSON.stringify(rows),
        days: days
    });

    // 「今アクティブなファイル」のポインターを更新 → 全員に配信される
    await setDoc(currentSheetRef, {
        docId: docId
    });
}

// 共有シートの変更をリアルタイムで受け取る
onSnapshot(currentSheetRef, async function (snap) {

    if (!snap.exists()) return;

    const docId = snap.data().docId;

    const sheetSnap = await getDoc(doc(db, "sheets", docId));
    if (!sheetSnap.exists()) return;

    const data = sheetSnap.data();

    selectedRows = JSON.parse(data.rowsJson);
    selectedFileName = data.fileName;
    selectedDays = data.days;

    await displayAttendance(selectedRows, selectedDays);
    refreshConditionDayOptions();
});

// ==============================
// 変更履歴の記録・ハイライト
// ==============================
function recordChange(cell, previousText, newText) {

    changeHistory.push({
        cell: cell,
        previousText: previousText,
        newText: newText
    });

    cell.classList.add("cell-changed");

    updateUndoButton();
}


// ==============================
// 「元に戻す」ボタンの見た目を更新
// ==============================
function updateUndoButton() {

    const button = document.getElementById("undo-button");
    if (!button) return;

    button.textContent = `↩ 元に戻す（${changeHistory.length}件）`;
    button.disabled = changeHistory.length === 0;
}


// ==============================
// 「元に戻す」ボタンのクリック処理
// ==============================
document
    .getElementById("undo-button")
    .addEventListener("click", function () {

        if (changeHistory.length === 0) return;

        // 履歴の一番最後（直近の変更）を取り出す
        const lastChange = changeHistory.pop();

        // セルの内容を変更前に戻す
        lastChange.cell.textContent = lastChange.previousText;

        if (lastChange.previousText === "〇") {
            lastChange.cell.classList.add("present");
        } else {
            lastChange.cell.classList.remove("present");
        }

        // ★同じセルが履歴に他にまだ残っていなければ、ハイライトを消す
        const stillInHistory = changeHistory.some(function (change) {
            return change.cell === lastChange.cell;
        });

        if (!stillInHistory) {
            lastChange.cell.classList.remove("cell-changed");
        }

        calculateTotals();
        saveAttendanceState();
        updateUndoButton();
    });

    // ==============================
// キャンセル対応：文字の正規化
// ==============================
function normalizeLine(line) {
    // コピペ時に紛れ込む見えない文字を除去
    return line.replace(/[\u200B-\u200F\u202A-\u202E\u2060\uFEFF]/g, "").trim();
}

const POSITIVE_WORDS = ["参加", "出席", "追加", "〇", "○"];
const NEGATIVE_WORDS = ["欠席", "キャンセル", "ｷｬﾝｾﾙ", "×"];

const ITEM_ALIASES = {
    "出欠": "出欠",
    "朝食": "朝食",
    "昼食": "昼食",
    "おやつ": "おやつ",
    "夕食": "夕食",
    "宿泊": "宿泊",
    "泊": "宿泊"
};


// ==============================
// 「9月12日」のような文字列から、その日が
// 何日目（1/2/3）に当たるかを調べる
// ==============================
function dayNumberToSlot(dayNum) {
    for (let slot = 1; slot <= 3; slot++) {
        if (dayLabels[slot] && dayLabels[slot].includes(dayNum + "日")) {
            return slot;
        }
    }
    return null;
}


// ==============================
// 「夕食、宿泊追加」「おやつ〇、夕食〇、泊〇」のような
// 項目リスト行を解析する
// ==============================
function splitLineIntoItemSegments(line) {

    const rawSegments = line.split(/[、,]/).map(function (s) { return s.trim(); }).filter(Boolean);
    if (rawSegments.length === 0) return null;

    function extractStatus(segment) {
        const allWords = POSITIVE_WORDS.concat(NEGATIVE_WORDS);

        for (let i = 0; i < allWords.length; i++) {
            const w = allWords[i];
            if (segment.endsWith(w)) {
                return {
                    rest: segment.slice(0, segment.length - w.length).replace(/→$/, "").trim(),
                    present: POSITIVE_WORDS.indexOf(w) !== -1
                };
            }
        }

        return null;
    }

    const parsedEach = rawSegments.map(extractStatus);

    // 全セグメントに個別の状態がある場合（おやつ〇、夕食〇、泊〇）
    if (parsedEach.every(function (p) { return p !== null; })) {
        return parsedEach.map(function (p) {
            return { itemLabel: p.rest, present: p.present };
        });
    }

    // 最後のセグメントだけに状態がある場合（夕食、宿泊追加）
    const lastParsed = parsedEach[parsedEach.length - 1];

    if (lastParsed !== null && parsedEach.slice(0, -1).every(function (p) { return p === null; })) {
        return rawSegments.map(function (seg, i) {
            const itemLabel = (i === rawSegments.length - 1) ? lastParsed.rest : seg;
            return { itemLabel: itemLabel, present: lastParsed.present };
        });
    }

    return null;
}


// ==============================
// 1人分のメモ（複数行）を解析する
// ==============================
function parsePersonBlock(lines, rosterNames) {

    let name = null;
    let currentDaySlot = null;
    const changes = [];
    const reasonLines = [];

    lines.forEach(function (rawLine) {

        const line = normalizeLine(rawLine);
        if (!line) return;

        // 名前を検出（見つかるまでの行は無視）
        if (!name && rosterNames.indexOf(line) !== -1) {
            name = line;
            return;
        }

        if (!name) return;

        // 「12日13日欠席」のような複数日＋状態
        let m = line.match(/^((?:\d{1,2}日)+)(参加|出席|欠席|キャンセル|追加)$/);
        if (m) {
            const dayNumbers = m[1].match(/\d{1,2}/g).map(Number);
            const present = POSITIVE_WORDS.indexOf(m[2]) !== -1;

            dayNumbers.forEach(function (dNum) {
                const slot = dayNumberToSlot(dNum);
                if (slot) {
                    changes.push({ daySlot: slot, item: "出欠", present: present });
                    currentDaySlot = slot;
                }
            });
            return;
        }

        // 「13日朝食追加」のような日付＋項目＋状態
        m = line.match(/^(\d{1,2})日(出欠|朝食|昼食|おやつ|夕食|宿泊|泊)(参加|出席|欠席|キャンセル|追加)$/);
        if (m) {
            const slot = dayNumberToSlot(Number(m[1]));
            const item = ITEM_ALIASES[m[2]];
            const present = POSITIVE_WORDS.indexOf(m[3]) !== -1;

            if (slot) {
                changes.push({ daySlot: slot, item: item, present: present });
                currentDaySlot = slot;
            }
            return;
        }

        // 項目リスト行（日付指定なし＝直前の日付を引き継ぐ）
        const segments = splitLineIntoItemSegments(line);

        if (segments && currentDaySlot) {

            segments.forEach(function (seg) {

                const item = ITEM_ALIASES[seg.itemLabel];

                // 表に無い項目（例：勉強会）は黙って無視する
                if (!item) return;

                changes.push({ daySlot: currentDaySlot, item: item, present: seg.present });
            });

            return;
        }

        // どれにも当てはまらなければ「理由」として扱う
        reasonLines.push(line);
    });

    return { name: name, changes: changes, reason: reasonLines.join(" ") };
}


// ==============================
// メモ全体（複数人分）を解析する
// ==============================
function parseMemoText(text) {

    const rosterNames = getAllRosterNames();

    const blocks = text.split(/\n\s*\n/);

    return blocks
        .map(function (block) {
            return parsePersonBlock(block.split(/\r?\n/), rosterNames);
        })
        .filter(function (result) { return result.name; });
}


// ==============================
// 今の出席表に載っている全員の名前を取得
// ==============================
function getAllRosterNames() {

    const names = [];

    document.querySelectorAll("#attendance-body .sticky-name").forEach(function (cell) {
        const text = cell.textContent.trim();
        if (text && text !== "計" && text !== "弁当内訳") {
            names.push(text);
        }
    });

    return names;
}


// ==============================
// 名前から、実際の出席表の行（tr）を探す
// ==============================
function findRowByName(name) {

    let found = null;

    document.querySelectorAll("#attendance-body tr:not(#total-row):not(.lunch-size-row)")
        .forEach(function (tr) {
            const cell = tr.querySelector(".sticky-name");
            if (cell && cell.textContent.trim() === name) {
                found = tr;
            }
        });

    return found;
}


// ==============================
// 「〇日目の〇〇」から、実際の<td>セルを探す
// ==============================
function getStartColForDays(days) {
    if (days === 1) return 14;
    if (days === 2) return 8;
    return 3;
}

function getCellForDayItem(row, daySlot, itemLabel) {

    const dayCols = DAY_COLUMNS[daySlot];
    if (!dayCols) return null;

    const itemDef = dayCols.find(function (i) { return i.label === itemLabel; });
    if (!itemDef) return null;

    const startCol = getStartColForDays(selectedDays);
    const cells = row.querySelectorAll("td");
    const index = 2 + (itemDef.col - startCol);

    return cells[index] || null;
}


// ==============================
// 名前の候補（datalist）を更新
// ==============================
function populateRosterNamesDatalist() {

    const datalist = document.getElementById("roster-names");
    if (!datalist) return;

    datalist.innerHTML = "";

    getAllRosterNames().forEach(function (name) {
        const option = document.createElement("option");
        option.value = name;
        datalist.appendChild(option);
    });
}


// ==============================
// パネルの行データ管理
// ==============================
let bulkRows = [];
let bulkRowIdCounter = 0;

function initRowToggles(rowData) {

    rowData.toggles = {};

    if (!rowData.name) {
        rowData.originalToggles = {};
        return;
    }

    const tr = findRowByName(rowData.name);
    if (!tr) {
        rowData.originalToggles = {};
        return;
    }

    getActiveDayNumbers().forEach(function (daySlot) {
        DAY_COLUMNS[daySlot].forEach(function (itemDef) {

            const cell = getCellForDayItem(tr, daySlot, itemDef.label);

            if (cell) {
                rowData.toggles[daySlot + "-" + itemDef.label] = (cell.textContent === "〇");
            }
        });
    });

    // ★今の実際の表の値を「元の値」として別に保存しておく
    rowData.originalToggles = Object.assign({}, rowData.toggles);
}

function renderRowToggles(rowEl, rowData) {

    const toggleArea = rowEl.querySelector(".bulk-toggle-area");
    toggleArea.innerHTML = "";

    getActiveDayNumbers().forEach(function (daySlot) {
        DAY_COLUMNS[daySlot].forEach(function (itemDef) {

            const key = daySlot + "-" + itemDef.label;
            const isPresent = rowData.toggles[key] === true;

            // ★元の値と比べて、変更されようとしているかを判定
            const originalValue = rowData.originalToggles ? rowData.originalToggles[key] : undefined;
            const isChanged = originalValue !== undefined && originalValue !== isPresent;

            const btn = document.createElement("button");
            btn.type = "button";
            btn.className = "bulk-toggle-button";
            btn.classList.toggle("on", isPresent);
            btn.classList.toggle("changed", isChanged);   // ★追加

            const dayLabel = dayLabels[daySlot] || (daySlot + "日目");

            // ★変更ありの場合は目印を付ける
            const mark = isChanged ? " ⚠変更" : "";
            btn.textContent = `${dayLabel} ${itemDef.label}：${isPresent ? "〇" : "-"}${mark}`;

            btn.addEventListener("click", function () {
                rowData.toggles[key] = !isPresent;
                renderRowToggles(rowEl, rowData);
            });

            toggleArea.appendChild(btn);
        });
    });
}

function renderBulkRow(rowData) {

    const container = document.getElementById("bulk-rows-container");

    const div = document.createElement("div");
    div.className = "bulk-row";
    div.dataset.rowId = rowData.id;

    const nameInput = document.createElement("input");
    nameInput.type = "text";
    nameInput.className = "bulk-name-input";
    nameInput.setAttribute("list", "roster-names");
    nameInput.placeholder = "名前を検索";
    nameInput.value = rowData.name || "";

    nameInput.addEventListener("change", function () {
        rowData.name = nameInput.value.trim();
        initRowToggles(rowData);
        renderRowToggles(div, rowData);
    });

    div.appendChild(nameInput);

    const toggleArea = document.createElement("div");
    toggleArea.className = "bulk-toggle-area";
    div.appendChild(toggleArea);

    const deleteButton = document.createElement("button");
    deleteButton.type = "button";
    deleteButton.className = "bulk-delete-row-button";
    deleteButton.textContent = "🗑";

    deleteButton.addEventListener("click", function () {
        bulkRows = bulkRows.filter(function (r) { return r.id !== rowData.id; });
        div.remove();
    });

    div.appendChild(deleteButton);

    container.appendChild(div);

    renderRowToggles(div, rowData);
}

function addBulkRow() {

    const rowData = { id: "row" + (bulkRowIdCounter++), name: "", toggles: {} };

    bulkRows.push(rowData);
    renderBulkRow(rowData);
}


// ==============================
// 「＋ 行を追加」ボタン
// ==============================
document
    .getElementById("bulk-add-row-button")
    .addEventListener("click", addBulkRow);

    addBulkRow;


// ==============================
// 「メモから自動入力」ボタン
// ==============================
document
    .getElementById("bulk-parse-button")
    .addEventListener("click", function () {

        const text = document.getElementById("bulk-memo").value;
        const parsed = parseMemoText(text);

        if (parsed.length === 0) {
            alert("メモから対象者を検出できませんでした。名簿に登録されている名前と完全に一致しているか確認してください。");
            return;
        }

        parsed.forEach(function (person) {

            const rowData = { id: "row" + (bulkRowIdCounter++), name: person.name, toggles: {} };

            // まず現在の実際の状態を初期値にする
            initRowToggles(rowData);

            // そこにメモから読み取った変更を上書きする
            person.changes.forEach(function (change) {
                rowData.toggles[change.daySlot + "-" + change.item] = change.present;
            });

            bulkRows.push(rowData);
            renderBulkRow(rowData);
        });

        alert(`${parsed.length}人分を自動入力しました。内容を確認してから「決定」を押してください。`);
    });


// ==============================
// 「決定」ボタン：まとめて実際の表に反映
// ==============================
document
    .getElementById("bulk-apply-button")
    .addEventListener("click", async function () {

        if (bulkRows.length === 0) {
            alert("反映する行がありません。");
            return;
        }

        let appliedCount = 0;

        bulkRows.forEach(function (rowData) {

            if (!rowData.name) return;

            const tr = findRowByName(rowData.name);
            if (!tr) return;

            Object.keys(rowData.toggles).forEach(function (key) {

                const parts = key.split("-");
                const daySlot = Number(parts[0]);
                const item = parts[1];

                const cell = getCellForDayItem(tr, daySlot, item);
                if (!cell) return;

                const newText = rowData.toggles[key] ? "〇" : "-";
                const previousText = cell.textContent;

                if (previousText === newText) return;

                cell.textContent = newText;
                cell.classList.toggle("present", newText === "〇");

                recordChange(cell, previousText, newText);
                appliedCount++;
            });
        });

        calculateTotals();
        await saveAttendanceState();

        alert(`${appliedCount}件のセルを変更しました。`);

        bulkRows = [];
        document.getElementById("bulk-rows-container").innerHTML = "";
    });


// ==============================
// メモをFirestoreに保存・共有する
// ==============================
function debounce(fn, wait) {
    let timer = null;
    return function () {
        const args = arguments;
        const self = this;
        clearTimeout(timer);
        timer = setTimeout(function () { fn.apply(self, args); }, wait);
    };
}

const bulkMemoRef = doc(db, "appState", "bulkMemo");

document
    .getElementById("bulk-memo")
    .addEventListener("input", debounce(async function () {
        await setDoc(bulkMemoRef, { text: document.getElementById("bulk-memo").value });
    }, 800));

onSnapshot(bulkMemoRef, function (snap) {

    if (!snap.exists()) return;

    const textarea = document.getElementById("bulk-memo");

    // 自分が今まさに入力中でなければ、他の人の更新を反映する
    if (document.activeElement !== textarea) {
        textarea.value = snap.data().text || "";
    }
});
