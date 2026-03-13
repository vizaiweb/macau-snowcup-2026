// 全局變量
let excelData = {};
const DATA_FILE_PATH = 'data/matches_data.xlsx';

// 頁面加載完成後執行
document.addEventListener('DOMContentLoaded', function() {
    // 加載Excel數據
    loadExcelData();
    
    // 綁定選項卡點擊事件
    bindTabEvents();
});

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

// 切換選項卡
function switchTab(btn, sectionId) {
    // 移除所有同區域選項卡的active類
    document.querySelectorAll(`${sectionId} .tab-btn`).forEach(tab => {
        tab.classList.remove('active');
    });
    // 當前選項卡添加active類
    btn.classList.add('active');
}

// 渲染對賽安排
function renderMatches(group) {
    const contentEl = document.getElementById('matches-content');
    const matchesData = excelData[`${group}_對賽安排`] || [];
    
    if (matchesData.length === 0) {
        contentEl.innerHTML = `<div class="loading">${group}暫無對賽安排數據</div>`;
        return;
    }
    
    // 生成表格
    let tableHtml = `
        <table>
            <thead>
                <tr>
                    <th>日期</th>
                    <th>時間</th>
                    <th>對賽隊伍A</th>
                    <th>對賽隊伍B</th>
                    <th>場地</th>
                </tr>
            </thead>
            <tbody>
    `;
    
    matchesData.forEach(item => {
        tableHtml += `
            <tr>
                <td>${item.日期 || '-'}</td>
                <td>${item.時間 || '-'}</td>
                <td>${item.隊伍A || '-'}</td>
                <td>${item.隊伍B || '-'}</td>
                <td>${item.場地 || '-'}</td>
            </tr>
        `;
    });
    
    tableHtml += `
            </tbody>
        </table>
    `;
    
    contentEl.innerHTML = tableHtml;
}

// 渲染對賽成績
function renderResults(group) {
    const contentEl = document.getElementById('results-content');
    const resultsData = excelData[`${group}_對賽成績`] || [];
    
    if (resultsData.length === 0) {
        contentEl.innerHTML = `<div class="loading">${group}暫無對賽成績數據</div>`;
        return;
    }
    
    // 生成表格
    let tableHtml = `
        <table>
            <thead>
                <tr>
                    <th>日期</th>
                    <th>隊伍A</th>
                    <th>比分</th>
                    <th>隊伍B</th>
                    <th>備註</th>
                </tr>
            </thead>
            <tbody>
    `;
    
    resultsData.forEach(item => {
        tableHtml += `
            <tr>
                <td>${item.日期 || '-'}</td>
                <td>${item.隊伍A || '-'}</td>
                <td>${item.比分 || '-'}</td>
                <td>${item.隊伍B || '-'}</td>
                <td>${item.備註 || '-'}</td>
            </tr>
        `;
    });
    
    tableHtml += `
            </tbody>
        </table>
    `;
    
    contentEl.innerHTML = tableHtml;
}

// 渲染積分榜
function renderRankings(group) {
    const contentEl = document.getElementById('rankings-content');
    const rankingsData = excelData[`${group}_積分榜`] || [];
    
    if (rankingsData.length === 0) {
        contentEl.innerHTML = `<div class="loading">${group}暫無積分榜數據</div>`;
        return;
    }
    
    // 按積分降序排序
    rankingsData.sort((a, b) => (b.積分 || 0) - (a.積分 || 0));
    
    // 生成表格
    let tableHtml = `
        <table>
            <thead>
                <tr>
                    <th>排名</th>
                    <th>隊伍名稱</th>
                    <th>勝場</th>
                    <th>負場</th>
                    <th>積分</th>
                </tr>
            </thead>
            <tbody>
    `;
    
    rankingsData.forEach((item, index) => {
        tableHtml += `
            <tr>
                <td>${index + 1}</td>
                <td>${item.隊伍名稱 || '-'}</td>
                <td>${item.勝場 || 0}</td>
                <td>${item.負場 || 0}</td>
                <td>${item.積分 || 0}</td>
            </tr>
        `;
    });
    
    tableHtml += `
            </tbody>
        </table>
    `;
    
    contentEl.innerHTML = tableHtml;
}
