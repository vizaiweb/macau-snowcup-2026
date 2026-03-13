// 全局變量
let excelData = {};
const DATA_FILE_PATH = 'data/matches_data.xlsx';

// 頁面加載完成後執行
document.addEventListener('DOMContentLoaded', function() {
    // 加載Excel數據
    loadExcelData();
    
    // 綁定選項卡點擊事件
    bindTabEvents();

    // 綁定頂部導航欄點擊事件
    bindNavEvents();
});

// ========== 工具函數 ==========
// 將 Excel 時間小數轉為 HH:MM 格式
function excelTimeToHHMM(excelTime) {
    // 非數字直接返回（如已為文本格式的時間）
    if (typeof excelTime !== 'number') return excelTime || '-';
    // Excel時間是當天的小數比例，轉換為分鐘
    const totalMinutes = Math.round(excelTime * 24 * 60);
    const hours = Math.floor(totalMinutes / 60);
    const minutes = totalMinutes % 60;
    // 補零格式化為 HH:MM
    return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}`;
}

// 安全獲取數據欄位（兼容不同表頭命名）
function getSafeValue(item, keys) {
    for (let key of keys) {
        if (item.hasOwnProperty(key) && item[key] !== undefined && item[key] !== null) {
            return item[key];
        }
    }
    return '-';
}

// ========== 數據加載 ==========
// 加載Excel數據
function loadExcelData() {
    fetch(DATA_FILE_PATH)
        .then(response => {
            if (!response.ok) throw new Error('加載數據文件失敗');
            return response.arrayBuffer();
        })
        .then(data => {
            // 解析Excel文件
            const workbook = XLSX.read(data, { type: 'array' });
            
            // 存儲各工作表數據
            workbook.SheetNames.forEach(sheetName => {
                const worksheet = workbook.Sheets[sheetName];
                excelData[sheetName] = XLSX.utils.sheet_to_json(worksheet);
            });
            
            // 更新最後更新時間
            document.getElementById('update-time').textContent = new Date().toLocaleString('zh-Hant-MO');
            
            // 初始化顯示數據
            renderMatches('初級組');
            renderResults('初級組');
            renderRankings('初級組');
        })
        .catch(error => {
            console.error('加載數據錯誤:', error);
            // 顯示錯誤信息
            document.querySelectorAll('.tab-content').forEach(content => {
                content.innerHTML = `<div class="loading">加載數據失敗，請檢查文件路徑或格式</div>`;
            });
        });
}

// ========== 事件綁定 ==========
// 綁定選項卡點擊事件
function bindTabEvents() {
    // 對賽安排選項卡
    document.querySelectorAll('#matches .tab-btn').forEach(btn => {
        btn.addEventListener('click', function() {
            switchTab(this, '#matches');
            renderMatches(this.dataset.group);
        });
    });
    
    // 對賽成績選項卡
    document.querySelectorAll('#results .tab-btn').forEach(btn => {
        btn.addEventListener('click', function() {
            switchTab(this, '#results');
            renderResults(this.dataset.group);
        });
    });
    
    // 積分榜選項卡
    document.querySelectorAll('#rankings .tab-btn').forEach(btn => {
        btn.addEventListener('click', function() {
            switchTab(this, '#rankings');
            renderRankings(this.dataset.group);
        });
    });
}

// 綁定頂部導航欄事件
function bindNavEvents() {
    document.querySelectorAll('nav a').forEach(link => {
        link.addEventListener('click', function(e) {
            e.preventDefault();
            // 移除所有導航激活狀態
            document.querySelectorAll('nav a').forEach(item => {
                item.classList.remove('nav-active');
            });
            // 當前導航激活
            this.classList.add('nav-active');
            // 滾動到對應區域
            const targetId = this.getAttribute('href');
            document.querySelector(targetId).scrollIntoView({
                behavior: 'smooth'
            });
        });
    });
}

// 切換選項卡
function switchTab(btn, sectionId) {
    // 移除所有同區域選項卡的active類
    document.querySelectorAll(`${sectionId} .tab-btn`).forEach(tab => {
        tab.classList.remove('active');
    });
    // 當前選項卡添加active類
    btn.classList.add('active');
}

// ========== 渲染函數 ==========
// 渲染對賽安排（修復時間/隊伍/組別顯示）
function renderMatches(group) {
    const contentEl = document.getElementById('matches-content');
    const matchesData = excelData[`${group}_對賽安排`] || [];
    
    if (matchesData.length === 0) {
        contentEl.innerHTML = `<div class="loading">${group}暫無對賽安排數據</div>`;
        return;
    }
    
    // 生成表格（新增「組別」欄位，修復時間格式）
    let tableHtml = `
        <table>
            <thead>
                <tr>
                    <th>日期</th>
                    <th>時間</th>
                    <th>對賽隊伍A</th>
                    <th>對賽隊伍B</th>
                    <th>場地</th>
                    <th>組別</th>
                </tr>
            </thead>
            <tbody>
    `;
    
    matchesData.forEach(item => {
        // 兼容多種表頭命名方式
        const date = getSafeValue(item, ['日期', '比賽日期']);
        const time = excelTimeToHHMM(getSafeValue(item, ['時間', '比賽時間']));
        const teamA = getSafeValue(item, ['隊伍A', '對賽隊伍A', '參賽隊伍A']);
        const teamB = getSafeValue(item, ['隊伍B', '對賽隊伍B', '參賽隊伍B']);
        const venue = getSafeValue(item, ['場地', '比賽場地']);
        const subgroup = getSafeValue(item, ['組別', '小組']);
        
        tableHtml += `
            <tr>
                <td>${date}</td>
                <td>${time}</td>
                <td>${teamA}</td>
                <td>${teamB}</td>
                <td>${venue}</td>
                <td>${subgroup}</td>
            </tr>
        `;
    });
    
    tableHtml += `
            </tbody>
        </table>
    `;
    
    contentEl.innerHTML = tableHtml;
}

// 渲染對賽成績（可選添加組別欄位）
function renderResults(group) {
    const contentEl = document.getElementById('results-content');
    const resultsData = excelData[`${group}_對賽成績`] || [];
    
    if (resultsData.length === 0) {
        contentEl.innerHTML = `<div class="loading">${group}暫無對賽成績數據</div>`;
        return;
    }
    
    // 生成表格（可選添加組別欄位）
    let tableHtml = `
        <table>
            <thead>
                <tr>
                    <th>日期</th>
                    <th>隊伍A</th>
                    <th>比分</th>
                    <th>隊伍B</th>
                    <th>備註</th>
                    <th>組別</th>
                </tr>
            </thead>
            <tbody>
    `;
    
    resultsData.forEach(item => {
        const date = getSafeValue(item, ['日期', '比賽日期']);
        const teamA = getSafeValue(item, ['隊伍A', '對賽隊伍A']);
        const score = getSafeValue(item, ['比分', '成績']);
        const teamB = getSafeValue(item, ['隊伍B', '對賽隊伍B']);
        const remark = getSafeValue(item, ['備註', '附註']);
        const subgroup = getSafeValue(item, ['組別', '小組']);
        
        tableHtml += `
            <tr>
                <td>${date}</td>
                <td>${teamA}</td>
                <td>${score}</td>
                <td>${teamB}</td>
                <td>${remark}</td>
                <td>${subgroup}</td>
            </tr>
        `;
    });
    
    tableHtml += `
            </tbody>
        </table>
    `;
    
    contentEl.innerHTML = tableHtml;
}

// 渲染積分榜（可選添加組別欄位）
function renderRankings(group) {
    const contentEl = document.getElementById('rankings-content');
    const rankingsData = excelData[`${group}_積分榜`] || [];
    
    if (rankingsData.length === 0) {
        contentEl.innerHTML = `<div class="loading">${group}暫無積分榜數據</div>`;
        return;
    }
    
    // 按積分降序排序
    rankingsData.sort((a, b) => (Number(getSafeValue(b, ['積分'])) || 0) - (Number(getSafeValue(a, ['積分'])) || 0));
    
    // 生成表格（可選添加組別欄位）
    let tableHtml = `
        <table>
            <thead>
                <tr>
                    <th>排名</th>
                    <th>隊伍名稱</th>
                    <th>勝場</th>
                    <th>負場</th>
                    <th>積分</th>
                    <th>組別</th>
                </tr>
            </thead>
            <tbody>
    `;
    
    rankingsData.forEach((item, index) => {
        const teamName = getSafeValue(item, ['隊伍名稱', '隊伍']);
        const win = getSafeValue(item, ['勝場', '勝']);
        const lose = getSafeValue(item, ['負場', '負']);
        const score = getSafeValue(item, ['積分']);
        const subgroup = getSafeValue(item, ['組別', '小組']);
        
        tableHtml += `
            <tr>
                <td>${index + 1}</td>
                <td>${teamName}</td>
                <td>${win}</td>
                <td>${lose}</td>
                <td>${score}</td>
                <td>${subgroup}</td>
            </tr>
        `;
    });
    
    tableHtml += `
            </tbody>
        </table>
    `;
    
    contentEl.innerHTML = tableHtml;
}
