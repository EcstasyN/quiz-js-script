let allRawData = [], quizData = [], currentIdx = 0, score = 0, errorStats = {}, currentSessionWrongs = [], currentFileName = "";
const getEl = id => document.getElementById(id);
const toggle = (id, isHidden) => getEl(id).classList.toggle('hidden', isHidden);
const saveStats = () => localStorage.setItem("quizErrorStats_" + currentFileName, JSON.stringify(errorStats));

function toggleThreshold() { toggle('threshold-div', getEl('quiz-mode').value !== 'threshold'); }
function updateWrongInfo() { getEl('wrong-total').innerText = Object.keys(errorStats).length; }

document.getElementById('fileInput').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;
    currentFileName = file.name;
    const reader = new FileReader();
    reader.onload = (event) => {
        const workbook = XLSX.read(new Uint8Array(event.target.result), { type: 'array' });
        let raw = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]).filter(item => item["題目"]);
        // --- 新增：自動補編號邏輯 ---
        allRawData = raw.map((item, index) => {
            // 如果 Excel 沒給題號，就用目前的「列數索引+1」當作編號
            if (!item["題號"]) {
                item["題號"] = (index + 1).toString();
            }
            return item;
        });

        errorStats = JSON.parse(localStorage.getItem("quizErrorStats_" + currentFileName)) || {};
        if (allRawData.length > 0) {
            resetRange();
            toggle('setting-section', false);
            updateWrongInfo();
        } else alert("無題目資料！");
    };
    reader.readAsArrayBuffer(file);
});

function validateRangeLimit() {
    const total = allRawData.length;
    if (total === 0) return;
    const startInput = getEl('range-start');
    const endInput = getEl('range-end');
    const limitInput = getEl('quiz-limit');
    const mode = getEl('quiz-mode').value;
    const threshold = parseInt(getEl('error-threshold').value) || 1;
    let s = parseInt(startInput.value) || 1;
    let e = parseInt(endInput.value) || total;

    // 1. 修正索引邊界
    if (s < 1) s = 1;
    if (s > total) s = total;
    if (e > total) e = total;
    if (e < s) e = s;
    startInput.value = s;
    endInput.value = e;

    // 2. 根據目前模式與範圍，計算「真正可用」的題目清單
    const sourceRange = allRawData.slice(s - 1, e);
    const filtered = sourceRange.filter(item => {
        const qID = item["題號"]?.toString().trim() || "no-id";
        const count = errorStats[qID] || 0;
        
        if (mode === "wrong-only") return count > 0;
        if (mode === "threshold") return count >= threshold;
        return true; // 隨機模式
    });

    const available = filtered.length;
    // 3. 更新 UI 上的總量提示（可選：讓使用者知道過濾後剩幾題）
    getEl('total-count').innerText = `${available}`;
    // 4. 修正出題數量
    let currentLimit = parseInt(limitInput.value) || 1;
    if (currentLimit > available) currentLimit = available;
    if (available === 0) currentLimit = 0; // 若無錯題，設為 0
    if (currentLimit < 1 && available > 0) currentLimit = 1;

    limitInput.max = available; 
    limitInput.value = currentLimit;
}

function resetRange() {
    getEl('range-start').value = 1;
    getEl('range-end').value = allRawData.length;
    getEl('total-count').innerText = allRawData.length;
    getEl('quiz-limit').value = Math.min(10, allRawData.length);
    validateRangeLimit();
}

function prepareQuiz() {
    validateRangeLimit();
    const mode = getEl('quiz-mode').value;
    const limit = parseInt(getEl('quiz-limit').value);
    const startIndex = parseInt(getEl('range-start').value) - 1; // 轉為 0-based index
    const endIndex = parseInt(getEl('range-end').value) - 1;

    // 從總資料中切出使用者選擇的列數範圍
    let sourceRange = allRawData.slice(startIndex, endIndex + 1);

    // 根據模式過濾 (錯題模式仍需參考題號作為唯一識別碼)
    let filtered = sourceRange.filter(item => {
        const qID = item["題號"]?.toString().trim() || "no-id";
        const count = errorStats[qID] || 0;
        
        if (mode === "wrong-only") return count > 0;
        if (mode === "threshold") return count >= parseInt(getEl('error-threshold').value);
        return true;
    });

    if (filtered.length === 0) {
        alert("在此列數範圍內無符合條件的題目！");
        return;
    }

    // 隨機洗牌並取樣
    quizData = filtered.sort(() => Math.random() - 0.5).slice(0, limit);
    
    currentIdx = 0;
    score = 0;
    currentSessionWrongs = [];
    toggle('upload-section', true);
    toggle('result-section', true);
    toggle('quiz-section', false);
    showQuestion();
    document.getElementById('current-file-name').innerText = currentFileName;
    document.getElementById('footer-section').classList.remove('hidden'); 
}

function showQuestion() {
    const item = quizData[currentIdx];
    getEl('quiz-progress').innerText = `進度：(${currentIdx + 1}/${quizData.length})`;
    getEl('question-text').innerText = `Q${currentIdx + 1}: ${item["題目"]}`;
    getEl('msg').innerHTML = "";
    toggle('next-btn', true);
    
    const container = getEl('options-container');
    container.innerHTML = '';
    ["選項1", "選項2", "選項3", "選項4"].forEach((key, i) => {
        if (item[key] !== undefined) {
            const btn = document.createElement('button');
            btn.className = 'option';
            btn.innerText = item[key];
            btn.onclick = () => checkAnswer(i, item["答案"]);
            container.appendChild(btn);
        }
    });
}

function checkAnswer(selected, correct) {
    const item = quizData[currentIdx], qID = item["題號"].toString(), correctIdx = parseInt(correct) - 1;
    const btns = document.querySelectorAll('.option');
    btns.forEach(b => b.disabled = true);
    toggle('next-btn', false);
    
    if (btns[correctIdx]) btns[correctIdx].classList.add('correct-option');

    if (selected === correctIdx) {
        getEl('msg').innerHTML = `<div style="color: green; display: flex; justify-content: center; font-size: 1.2rem; font-weight:bold;">✅ 正確！</div>`;
        score++;
    } else {
        getEl('msg').innerHTML = `<div style="color: #f44336; display: flex; justify-content: center; font-size: 1.2rem; font-weight:bold;">❌ 答錯了！</div>`;
        if (btns[selected]) btns[selected].classList.add('wrong-option');
        errorStats[qID] = (errorStats[qID] || 0) + 1;
        saveStats();
        currentSessionWrongs.push({ id: qID, question: item["題目"] });
        updateWrongInfo();
    }
}

function nextQuestion() {
    if (++currentIdx < quizData.length) showQuestion();
    else showFinalResult();
}

function showFinalResult() {
    toggle('quiz-section', true);
    toggle('result-section', false);
    getEl('res-total').innerText = quizData.length;
    getEl('res-correct').innerText = score;
    getEl('res-wrong').innerText = quizData.length - score;
    getEl('res-rate').innerText = ((score / quizData.length) * 100).toFixed(1);
}

// function exportWrongQuestions() {
//     if (currentSessionWrongs.length === 0) return alert("無錯題紀錄！");
//     const data = allRawData.filter(item => currentSessionWrongs.some(w => w.id === item["題號"].toString().trim()));
//     downloadExcel(data, "單次測驗錯題");
// }

function exportAllHistoryWrongs() {
    const data = allRawData.filter(item => (errorStats[item["題號"]?.toString().trim()] || 0) > 0)
                          .map(item => ({...item, "錯誤次數": errorStats[item["題號"].toString().trim()]}));
    if (data.length === 0) return alert("無歷史紀錄！");
    downloadExcel(data, "錯題統計");
}

function downloadExcel(data, sheetName) {
    const ws = excelAlter(data), wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    XLSX.writeFile(wb, `${sheetName}-${currentFileName.split('.')[0]}.xlsx`);
}

function excelAlter(data) {
    const ws = XLSX.utils.json_to_sheet(data), range = XLSX.utils.decode_range(ws['!ref']);
    const colW = [{wch:8}, {wch:8}, {wch:40}, {wch:20}, {wch:20}, {wch:20}, {wch:20}, {wch:8}];
    ws['!cols'] = colW;
    const rows = [];
    for (let R = range.s.r; R <= range.e.r; ++R) {
        let ml = 1;
        for (let C = range.s.c; C <= range.e.c; ++C) {
            const cell = ws[XLSX.utils.encode_cell({c:C, r:R})];
            if (!cell) continue;
            cell.s = { alignment: { wrapText: true, vertical: "center", horizontal: [0,1,7].includes(C) ? "center" : "left" }, font: { name: "微軟正黑體" } };
            if (cell.v && R > 0) ml = Math.max(ml, Math.ceil((cell.v.toString().length * 2) / (colW[C]?.wch || 10)));
        }
        rows.push({ hpt: R === 0 ? 25 : ml * 17 + 5 });
    }
    ws['!rows'] = rows;
    return ws;
}

function clearErrorStats() {
    if (currentFileName && confirm(`確定重置 [${currentFileName}] 的紀錄？`)) {
        localStorage.removeItem("quizErrorStats_" + currentFileName);
        errorStats = {}; updateWrongInfo();
    }
}

function restartSettings() {
    const isQuizOngoing = !getEl('quiz-section').classList.contains('hidden');
    
    // 如果正在測驗中，則彈出確認視窗
    if (isQuizOngoing) {
        if (!confirm("確定要中斷測驗並返回嗎？")) {
            return; // 使用者按了取消，保留在原地
        }
    }

    // 隱藏測驗與結果區塊，顯示上傳/設定區塊
    getEl('quiz-section').classList.add('hidden');
    getEl('result-section').classList.add('hidden');
    getEl('upload-section').classList.remove('hidden');
    document.getElementById('footer-section').classList.add('hidden');
}

document.addEventListener('keydown', (e) => {
    const isAnswered = !getEl('next-btn').classList.contains('hidden');
    if (e.key === "Enter" && isAnswered) nextQuestion();
    else if (!isAnswered && ["1","2","3","4"].includes(e.key)) {
        const btns = document.querySelectorAll('.option');
        if (btns[e.key-1]) btns[e.key-1].click();
    }
});
