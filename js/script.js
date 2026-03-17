// 全局變量
let excelData = {};
const DATA_FILE_PATH = 'matches_data.xlsx';

// 頁面加載完成後執行
document.addEventListener('DOMContentLoaded', function() {
    loadExcelData();
    bindAllEvents();
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
    // 1. 先請求獲取Excel文件的最後修改時間（需後端接口返回）
    fetch('/api/get-excel-modify-time') // 後端接口示例
        .then(timeRes => timeRes.json())
        .then(timeData => {
            const excelModifyTime = new Date(timeData.modifyTime);
            
            // 2. 再加載Excel數據
            return fetch(DATA_FILE_PATH + '?t=' + new Date().getTime())
                .then(response => {
                    if (!response.ok) throw new Error(`文件加載失敗（狀態碼：${response.status}）`);
                    return response.arrayBuffer();
                })
                .then(data => {
                    const workbook = XLSX.read(data, { type: 'array' });
                    workbook.SheetNames.forEach(sheetName => {
                        excelData[sheetName] = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
                    });
                    renderMatches('初級組');
                    renderResults('初級組');
                    renderRankings('初級組');
                    
                    // 顯示Excel文件的最後修改時間
                    const formattedTime = excelModifyTime.toLocaleString('zh-Hant-MO', {
                        day: 'numeric',
                        month: 'numeric',
                        year: 'numeric',
                        hour: '2-digit',
                        minute: '2-digit',
                        second: '2-digit',
                        hour12: true
                    });
                    document.getElementById('update-time').textContent = formattedTime;
                });
        })
        .catch(error => {
            // 錯誤處理（不變）
            const errorHtml = `
            <div style="padding:20px; color:#e74c3c; text-align:center;">
                <h4>數據加載失敗！</h4>
                <p>錯誤詳情：${error.message}</p>
                <p>請確認 Excel 路徑：${DATA_FILE_PATH}</p>
            </div>`;
            document.querySelectorAll('#matches-content, #results-content, #rankings-content').forEach(el => {
                el.innerHTML = errorHtml;
            });
        });
}

// ========== 事件綁定 ==========
function bindAllEvents() {
    document.querySelectorAll('#matches .tab-btn').forEach(btn => {
        btn.addEventListener('click', function() {
            const group = this.dataset.group;
            switchTab(this, '#matches');
            renderMatches(group);
        });
    });

    document.querySelectorAll('#results .tab-btn').forEach(btn => {
        btn.addEventListener('click', function() {
            const group = this.dataset.group;
            switchTab(this, '#results');
            renderResults(group);
        });
    });

    document.querySelectorAll('#rankings .tab-btn').forEach(btn => {
        btn.addEventListener('click', function() {
            const group = this.dataset.group;
            switchTab(this, '#rankings');
            renderRankings(group);
        });
    });

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

// ========== 渲染對賽安排 ==========
function renderMatches(group) {
    const contentEl = document.getElementById('matches-content');
    const sheetName = `${group}_對賽安排`;
    const matchesData = excelData[sheetName] || [];

    if (matchesData.length === 0) {
        contentEl.innerHTML = `<div class="loading">${group}暫無對賽安排數據</div>`;
        return;
    }

    let tableHtml = `
    <table style="width:100%; border-collapse:collapse;">
        <thead>
            <tr style="background:#3498db; color:white;">
                <th style="border:1px solid #ccc; padding:8px;">日期</th>
                <th style="border:1px solid #ccc; padding:8px;">時間</th>
                <th style="border:1px solid #ccc; padding:8px;">隊伍A</th>
                <th style="border:1px solid #ccc; padding:8px;">隊伍B</th>
                <th style="border:1px solid #ccc; padding:8px;">場地</th>
                <th style="border:1px solid #ccc; padding:8px;">組別</th>
                <th style="border:1px solid #ccc; padding:8px;">比分</th>
                <th style="border:1px solid #ccc; padding:8px;">備註</th>
            </tr>
        </thead>
        <tbody>
    `;

    matchesData.forEach(item => {
        const date = item['日期'] || '-';
        const time = excelTimeToHHMM(item['時間']);
        const teamA = item['隊伍A'] || '-';
        const teamB = item['隊伍B'] || '-';
        const venue = item['場地'] || '-';
        const subgroup = item['組別'] || '-';
        const score = item['比分'] || '未開賽';
        const remark = item['備註'] || '-';

        tableHtml += `
        <tr style="${item['比分'] ? 'background:#f0f9ff;' : ''}">
            <td style="border:1px solid #ccc; padding:8px;">${date}</td>
            <td style="border:1px solid #ccc; padding:8px;">${time}</td>
            <td style="border:1px solid #ccc; padding:8px;">${teamA}</td>
            <td style="border:1px solid #ccc; padding:8px;">${teamB}</td>
            <td style="border:1px solid #ccc; padding:8px;">${venue}</td>
            <td style="border:1px solid #ccc; padding:8px;">${subgroup}</td>
            <td style="border:1px solid #ccc; padding:8px; font-weight:bold;">${score}</td>
            <td style="border:1px solid #ccc; padding:8px;">${remark}</td>
        </tr>
        `;
    });

    tableHtml += `</tbody></table>`;
    contentEl.innerHTML = tableHtml;
}

// ========== 渲染對賽成績（顯示所有場次！） ==========
function renderResults(group) {
    const contentEl = document.getElementById('results-content');
    const sheetName = `${group}_對賽安排`;
    const matchesData = excelData[sheetName] || [];

    if (matchesData.length === 0) {
        contentEl.innerHTML = `<div class="loading">${group}暫無對賽數據</div>`;
        return;
    }

    // 按組別分類全部賽事（不論有沒有比數）
    const grouped = {};
    matchesData.forEach(item => {
        const subgroup = item['組別'] || '未分組';
        if (!grouped[subgroup]) grouped[subgroup] = [];
        grouped[subgroup].push(item);
    });

    let html = '';
    Object.keys(grouped).forEach(subgroup => {
        html += `
        <div style="margin-bottom:24px;">
            <h3 style="margin:0 0 8px; font-size:16px; font-weight:bold;">${subgroup} 組 - 對賽成績</h3>
            <table style="width:100%; border-collapse:collapse; table-layout:fixed;">
                <thead>
                    <tr style="background:#3498db; color:white;">
                        <th style="border:1px solid #ccc; padding:8px; width:12%; white-space:nowrap;">日期</th>
                        <th style="border:1px solid #ccc; padding:8px; width:12%; white-space:nowrap;">時間</th>
                        <th style="border:1px solid #ccc; padding:8px; width:25%; white-space:nowrap;">隊伍A</th>
                        <th style="border:1px solid #ccc; padding:8px; width:14%; white-space:nowrap;">比分</th>
                        <th style="border:1px solid #ccc; padding:8px; width:25%; white-space:nowrap;">隊伍B</th>
                        <th style="border:1px solid #ccc; padding:8px; width:12%; white-space:nowrap;">場地</th>
                    </tr>
                </thead>
                <tbody>
        `;

        grouped[subgroup].forEach(item => {
            html += `
            <tr>
                <td style="border:1px solid #ccc; padding:8px; white-space:nowrap;">${item['日期'] || '-'}</td>
                <td style="border:1px solid #ccc; padding:8px; white-space:nowrap;">${excelTimeToHHMM(item['時間'])}</td>
                <td style="border:1px solid #ccc; padding:8px; white-space:nowrap;">${item['隊伍A'] || '-'}</td>
                <td style="border:1px solid #ccc; padding:8px; font-weight:bold; white-space:nowrap;">${item['比分'] || '-'}</td>
                <td style="border:1px solid #ccc; padding:8px; white-space:nowrap;">${item['隊伍B'] || '-'}</td>
                <td style="border:1px solid #ccc; padding:8px; white-space:nowrap;">${item['場地'] || '-'}</td>
            </tr>
            `;
        });

        html += `</tbody></table></div>`;
    });

    contentEl.innerHTML = html;
}

// ========== 渲染積分榜（不顯示"補賽 組"） ==========
function renderRankings(group) {
    const contentEl = document.getElementById('rankings-content');
    const sheetName = `${group}_對賽安排`;
    const matchesData = excelData[sheetName] || [];

    // 收集所有隊伍
    const groupStats = {};
    matchesData.forEach(item => {
        const teamA = item['隊伍A'] || '-';
        const teamB = item['隊伍B'] || '-';
        const subgroup = item['組別'] || '未分組';
        if (!groupStats[subgroup]) groupStats[subgroup] = {};
        const initStats = () => ({ win:0, draw:0, lose:0, goal:0, concede:0, score:0 });
        if (!groupStats[subgroup][teamA]) groupStats[subgroup][teamA] = initStats();
        if (!groupStats[subgroup][teamB]) groupStats[subgroup][teamB] = initStats();
    });

    // 計算已完賽積分
    matchesData.forEach(item => {
        const teamA = item['隊伍A'] || '-';
        const teamB = item['隊伍B'] || '-';
        const subgroup = item['組別'] || '未分組';
        const score = item['比分'];
        if (typeof score === 'string' && /^\d+-\d+$/.test(score.trim())) {
            const [scoreA, scoreB] = score.trim().split('-').map(Number);
            groupStats[subgroup][teamA].goal += scoreA;
            groupStats[subgroup][teamA].concede += scoreB;
            groupStats[subgroup][teamB].goal += scoreB;
            groupStats[subgroup][teamB].concede += scoreA;

            if (scoreA > scoreB) {
                groupStats[subgroup][teamA].win += 1;
                groupStats[subgroup][teamA].score += 3;
                groupStats[subgroup][teamB].lose += 1;
            } else if (scoreA < scoreB) {
                groupStats[subgroup][teamB].win += 1;
                groupStats[subgroup][teamB].score += 3;
                groupStats[subgroup][teamA].lose += 1;
            } else {
                groupStats[subgroup][teamA].draw += 1;
                groupStats[subgroup][teamA].score += 1;
                groupStats[subgroup][teamB].draw += 1;
                groupStats[subgroup][teamB].score += 1;
            }
        }
    });

    let html = '';
    Object.keys(groupStats).forEach(subgroup => {
        if (subgroup === '補賽 組') return; // 跳過補賽組

        const sortedTeams = Object.keys(groupStats[subgroup]).sort((a, b) => {
            const ta = groupStats[subgroup][a], tb = groupStats[subgroup][b];
            if (tb.score !== ta.score) return tb.score - ta.score;
            return (tb.goal - tb.concede) - (ta.goal - ta.concede);
        });
        html += `
        <div style="margin-bottom:24px;">
            <h3 style="margin:0 0 8px; font-size:16px; font-weight:bold;">${subgroup} 組 - 積分榜</h3>
            <table style="width:100%; border-collapse:collapse; table-layout:auto;">
                <thead>
                    <tr style="background:#3498db; color:white;">
                        <th style="border:1px solid #ccc; padding:8px; min-width:60px; width:8%; white-space:nowrap;">排名</th>
                        <th style="border:1px solid #ccc; padding:8px; min-width:160px; width:30%; white-space:nowrap;">隊伍名稱</th>
                        <th style="border:1px solid #ccc; padding:8px; min-width:40px; width:8%; white-space:nowrap;">勝</th>
                        <th style="border:1px solid #ccc; padding:8px; min-width:40px; width:8%; white-space:nowrap;">平</th>
                        <th style="border:1px solid #ccc; padding:8px; min-width:40px; width:8%; white-space:nowrap;">負</th>
                        <th style="border:1px solid #ccc; padding:8px; min-width:40px; width:8%; white-space:nowrap;">進球</th>
                        <th style="border:1px solid #ccc; padding:8px; min-width:40px; width:8%; white-space:nowrap;">失球</th>
                        <th style="border:1px solid #ccc; padding:8px; min-width:50px; width:10%; white-space:nowrap;">積分</th>
                    </tr>
                </thead>
                <tbody>
        `;
        sortedTeams.forEach((team, index) => {
            const t = groupStats[subgroup][team];
            html += `
            <tr style="${index % 2 === 0 ? 'background:#f9f9f9;' : ''}">
                <td style="border:1px solid #ccc; padding:8px; text-align:center; white-space:nowrap;">${index + 1}</td>
                <td style="border:1px solid #ccc; padding:8px; white-space:nowrap;">${team}</td>
                <td style="border:1px solid #ccc; padding:8px; text-align:center; white-space:nowrap;">${t.win}</td>
                <td style="border:1px solid #ccc; padding:8px; text-align:center; white-space:nowrap;">${t.draw}</td>
                <td style="border:1px solid #ccc; padding:8px; text-align:center; white-space:nowrap;">${t.lose}</td>
                <td style="border:1px solid #ccc; padding:8px; text-align:center; white-space:nowrap;">${t.goal}</td>
                <td style="border:1px solid #ccc; padding:8px; text-align:center; white-space:nowrap;">${t.concede}</td>
                <td style="border:1px solid #ccc; padding:8px; text-align:center; font-weight:bold; color:#e74c3c; white-space:nowrap;">${t.score}</td>
            </tr>
            `;
        });
        html += `</tbody></table></div>`;
    });
    contentEl.innerHTML = html;
}
