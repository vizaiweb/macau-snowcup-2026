// 全局變量
let excelData = {};
const DATA_FILE_PATH = 'data/matches_data.xlsx';

// 頁面加載完成後執行
document.addEventListener('DOMContentLoaded', function() {
    loadExcelData();
    bindTabEvents();
    bindNavEvents();
});

// ========== 工具函數 ==========
function excelTimeToHHMM(excelTime) {
    if (typeof excelTime !== 'number') return excelTime || '-';
    const totalMinutes = Math.round(excelTime * 24 * 60);
    const hours = Math.floor(totalMinutes / 60);
    const minutes = totalMinutes % 60;
    return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}`;
}

// ========== 數據加載 ==========
function loadExcelData() {
    // 禁用緩存，強制讀取最新Excel
    fetch(DATA_FILE_PATH + '?t=' + new Date().getTime())
        .then(response => {
            if (!response.ok) throw new Error('加載數據失敗：' + response.status);
            return response.arrayBuffer();
        })
        .then(data => {
            const workbook = XLSX.read(data, { type: 'array' });
            // 只讀取需要的工作表（簡化邏輯）
            const targetSheet = '初級組_對賽安排';
            if (workbook.SheetNames.includes(targetSheet)) {
                excelData[targetSheet] = XLSX.utils.sheet_to_json(workbook.Sheets[targetSheet]);
                console.log('✅ 數據讀取成功：', excelData[targetSheet]); // 控制台查看數據
            }
            // 強制渲染
            renderMatches('初級組');
            document.getElementById('update-time').textContent = new Date().toLocaleString('zh-Hant-MO');
        })
        .catch(error => {
            console.error('❌ 加載錯誤：', error);
            document.getElementById('matches-content').innerHTML = `<div>加載失敗：${error.message}</div>`;
        });
}

// ========== 事件綁定 ==========
function bindTabEvents() {
    document.querySelectorAll('#matches .tab-btn').forEach(btn => {
        btn.addEventListener('click', function() {
            switchTab(this, '#matches');
            renderMatches(this.dataset.group);
        });
    });
    document.querySelectorAll('#results .tab-btn').forEach(btn => {
        btn.addEventListener('click', function() {
            switchTab(this, '#results');
            renderResults(this.dataset.group);
        });
    });
    document.querySelectorAll('#rankings .tab-btn').forEach(btn => {
        btn.addEventListener('click', function() {
            switchTab(this, '#rankings');
            renderRankings(this.dataset.group);
        });
    });
}

function bindNavEvents() {
    document.querySelectorAll('nav a').forEach(link => {
        link.addEventListener('click', function(e) {
            e.preventDefault();
            document.querySelectorAll('nav a').forEach(item => item.classList.remove('nav-active'));
            this.classList.add('nav-active');
            document.querySelector(this.getAttribute('href')).scrollIntoView({ behavior: 'smooth' });
        });
    });
}

function switchTab(btn, sectionId) {
    document.querySelectorAll(`${sectionId} .tab-btn`).forEach(tab => tab.classList.remove('active'));
    btn.classList.add('active');
}

// ========== 渲染函數 ==========
// 核心：匹配Excel表頭「隊伍A」「隊伍B」（無空格）
function renderMatches(group) {
    const contentEl = document.getElementById('matches-content');
    const matchesData = excelData[`${group}_對賽安排`] || [];

    if (matchesData.length === 0) {
        contentEl.innerHTML = `<div>${group}暫無對賽安排數據</div>`;
        return;
    }

    let tableHtml = `
        <table border="1" style="width:100%;border-collapse:collapse;">
            <thead>
                <tr style="background:#f0f0f0;">
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
        // 100%匹配Excel表頭（無空格）
        const date = item['日期'] || '-';
        const time = excelTimeToHHMM(item['時間']);
        const teamA = item['隊伍A'] || '-'; // 無空格
        const teamB = item['隊伍B'] || '-'; // 無空格
        const venue = item['場地'] || '-';
        const subgroup = item['組別'] || '-';

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

    tableHtml += `</tbody></table>`;
    contentEl.innerHTML = tableHtml;
}

function renderResults(group) {
    const contentEl = document.getElementById('results-content');
    const resultsData = excelData[`${group}_對賽成績`] || [];

    if (resultsData.length === 0) {
        contentEl.innerHTML = `<div>${group}暫無對賽成績數據</div>`;
        return;
    }

    let tableHtml = `
        <table border="1" style="width:100%;border-collapse:collapse;">
            <thead>
                <tr style="background:#f0f0f0;">
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
        const date = item['日期'] || '-';
        const teamA = item['隊伍A'] || '-';
        const score = item['比分'] || '-';
        const teamB = item['隊伍B'] || '-';
        const remark = item['備註'] || '-';
        const subgroup = item['組別'] || '-';

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

    tableHtml += `</tbody></table>`;
    contentEl.innerHTML = tableHtml;
}

function renderRankings(group) {
    const contentEl = document.getElementById('rankings-content');
    const rankingsData = excelData[`${group}_積分榜`] || [];

    if (rankingsData.length === 0) {
        contentEl.innerHTML = `<div>${group}暫無積分榜數據</div>`;
        return;
    }

    rankingsData.sort((a, b) => (Number(b['積分']) || 0) - (Number(a['積分']) || 0));

    let tableHtml = `
        <table border="1" style="width:100%;border-collapse:collapse;">
            <thead>
                <tr style="background:#f0f0f0;">
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
        const teamName = item['隊伍名稱'] || '-';
        const win = item['勝場'] || 0;
        const lose = item['負場'] || 0;
        const score = item['積分'] || 0;
        const subgroup = item['組別'] || '-';

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

    tableHtml += `</tbody></table>`;
    contentEl.innerHTML = tableHtml;
}
