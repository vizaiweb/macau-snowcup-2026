// 全局變量（完全匹配你的文件結構）
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

// ========== 數據加載核心函數（不變，只是路徑已修正） ==========
function loadExcelData() {
    console.log('🔍 嘗試加載 Excel：', DATA_FILE_PATH);
    
    fetch(DATA_FILE_PATH + '?t=' + new Date().getTime())
        .then(response => {
            console.log('🔍 服務器響應狀態：', response.status);
            if (!response.ok) throw new Error(`文件加載失敗（狀態碼：${response.status}）`);
            return response.arrayBuffer();
        })
        .then(data => {
            console.log('✅ Excel 加載成功，大小：', data.byteLength, '字節');
            const workbook = XLSX.read(data, { type: 'array' });
            console.log('📋 Excel 所有分頁：', workbook.SheetNames);

            workbook.SheetNames.forEach(sheetName => {
                excelData[sheetName] = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
                console.log(`✅ 加載分頁「${sheetName}」，數據行數：`, excelData[sheetName].length);
            });

            renderMatches('初級組');
            renderResults('初級組');
            renderRankings('初級組');
            document.getElementById('update-time').textContent = new Date().toLocaleString('zh-Hant-MO');
        })
        .catch(error => {
            console.error('❌ 加載錯誤：', error.message);
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
    // 對賽安排選項卡
    document.querySelectorAll('#matches .tab-btn').forEach(btn => {
        btn.addEventListener('click', function() {
            const group = this.dataset.group;
            switchTab(this, '#matches');
            renderMatches(group);
        });
    });

    // 對賽成績選項卡
    document.querySelectorAll('#results .tab-btn').forEach(btn => {
        btn.addEventListener('click', function() {
            const group = this.dataset.group;
            switchTab(this, '#results');
            renderResults(group);
        });
    });

    // 積分榜選項卡
    document.querySelectorAll('#rankings .tab-btn').forEach(btn => {
        btn.addEventListener('click', function() {
            const group = this.dataset.group;
            switchTab(this, '#rankings');
            renderRankings(group);
        });
    });

    // 頂部導航
    document.querySelectorAll('nav a').forEach(link => {
        link.addEventListener('click', function(e) {
            e.preventDefault();
            document.querySelectorAll('nav a').forEach(item => item.classList.remove('nav-active'));
            this.classList.add('nav-active');
            document.querySelector(this.getAttribute('href')).scrollIntoView({ behavior: 'smooth' });
        });
    });
}

// 選項卡切換
function switchTab(btn, sectionId) {
    document.querySelectorAll(`${sectionId} .tab-btn`).forEach(tab => tab.classList.remove('active'));
    btn.classList.add('active');
}

// ========== 渲染對賽安排（匹配Excel 8行數據） ==========
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

// ========== 渲染對賽成績（只顯示已完賽的4場） ==========
function renderResults(group) {
    const contentEl = document.getElementById('results-content');
    const sheetName = `${group}_對賽安排`;
    const matchesData = excelData[sheetName] || [];

    // 只篩選已完賽的賽事（有比分且不是空）
    const finishedMatches = matchesData.filter(item => {
        const score = item['比分'] || '';
        return score.trim() !== '' && score.includes('-');
    });

    if (finishedMatches.length === 0) {
        contentEl.innerHTML = `<div class="loading">${group}暫無已完賽賽事</div>`;
        return;
    }

    // 按組別分表（初A、初B、初C、初D）
    const grouped = {};
    finishedMatches.forEach(item => {
        const subgroup = item['組別'] || '未分組';
        if (!grouped[subgroup]) grouped[subgroup] = [];
        grouped[subgroup].push(item);
    });

    let html = '';
    Object.keys(grouped).forEach(subgroup => {
        html += `
        <div style="margin-bottom:24px;">
            <h3 style="margin:0 0 8px; font-size:16px; font-weight:bold;">${subgroup} 組 - 對賽成績</h3>
            <table style="width:100%; border-collapse:collapse;">
                <thead>
                    <tr style="background:#3498db; color:white;">
                        <th style="border:1px solid #ccc; padding:8px;">日期</th>
                        <th style="border:1px solid #ccc; padding:8px;">時間</th>
                        <th style="border:1px solid #ccc; padding:8px;">隊伍A</th>
                        <th style="border:1px solid #ccc; padding:8px;">比分</th>
                        <th style="border:1px solid #ccc; padding:8px;">隊伍B</th>
                        <th style="border:1px solid #ccc; padding:8px;">場地</th>
                    </tr>
                </thead>
                <tbody>
        `;

        grouped[subgroup].forEach(item => {
            html += `
            <tr>
                <td style="border:1px solid #ccc; padding:8px;">${item['日期'] || '-'}</td>
                <td style="border:1px solid #ccc; padding:8px;">${excelTimeToHHMM(item['時間'])}</td>
                <td style="border:1px solid #ccc; padding:8px;">${item['隊伍A'] || '-'}</td>
                <td style="border:1px solid #ccc; padding:8px; font-weight:bold;">${item['比分'] || '-'}</td>
                <td style="border:1px solid #ccc; padding:8px;">${item['隊伍B'] || '-'}</td>
                <td style="border:1px solid #ccc; padding:8px;">${item['場地'] || '-'}</td>
            </tr>
            `;
        });

        html += `</tbody></table></div>`;
    });

    contentEl.innerHTML = html;
}

// ========== 渲染積分榜（基於已完賽的4場計算） ==========
function renderRankings(group) {
    const contentEl = document.getElementById('rankings-content');
    const sheetName = `${group}_對賽安排`;
    const matchesData = excelData[sheetName] || [];

    // 篩選已完賽賽事
    const finishedMatches = matchesData.filter(item => {
        const score = item['比分'] || '';
        return score.trim() !== '' && score.includes('-');
    });

    if (finishedMatches.length === 0) {
        contentEl.innerHTML = `<div class="loading">${group}暫無已完賽賽事，無法計算積分</div>`;
        return;
    }

    // 初始化積分統計
    const groupStats = {};
    finishedMatches.forEach(item => {
        const teamA = item['隊伍A'] || '-';
        const teamB = item['隊伍B'] || '-';
        const [scoreA, scoreB] = (item['比分'] || '0-0').split('-').map(Number);
        const subgroup = item['組別'] || '未分組';

        // 初始化組別和隊伍
        if (!groupStats[subgroup]) groupStats[subgroup] = {};
        const initStats = () => ({ win:0, draw:0, lose:0, goal:0, concede:0, score:0 });
        if (!groupStats[subgroup][teamA]) groupStats[subgroup][teamA] = initStats();
        if (!groupStats[subgroup][teamB]) groupStats[subgroup][teamB] = initStats();

        // 更新進失球
        groupStats[subgroup][teamA].goal += scoreA;
        groupStats[subgroup][teamA].concede += scoreB;
        groupStats[subgroup][teamB].goal += scoreB;
        groupStats[subgroup][teamB].concede += scoreA;

        // 計算積分
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
    });

    // 生成積分表
    let html = '';
    Object.keys(groupStats).forEach(subgroup => {
        // 按積分排序（同分按淨勝球）
        const sortedTeams = Object.keys(groupStats[subgroup]).sort((a, b) => {
            const ta = groupStats[subgroup][a];
            const tb = groupStats[subgroup][b];
            if (tb.score !== ta.score) return tb.score - ta.score;
            return (tb.goal - tb.concede) - (ta.goal - ta.concede);
        });

        html += `
        <div style="margin-bottom:24px;">
            <h3 style="margin:0 0 8px; font-size:16px; font-weight:bold;">${subgroup} 組 - 積分榜</h3>
            <table style="width:100%; border-collapse:collapse;">
                <thead>
                    <tr style="background:#3498db; color:white;">
                        <th style="border:1px solid #ccc; padding:8px;">排名</th>
                        <th style="border:1px solid #ccc; padding:8px;">隊伍名稱</th>
                        <th style="border:1px solid #ccc; padding:8px;">勝</th>
                        <th style="border:1px solid #ccc; padding:8px;">平</th>
                        <th style="border:1px solid #ccc; padding:8px;">負</th>
                        <th style="border:1px solid #ccc; padding:8px;">進球</th>
                        <th style="border:1px solid #ccc; padding:8px;">失球</th>
                        <th style="border:1px solid #ccc; padding:8px;">積分</th>
                    </tr>
                </thead>
                <tbody>
        `;

        sortedTeams.forEach((team, index) => {
            const t = groupStats[subgroup][team];
            html += `
            <tr style="${index % 2 === 0 ? 'background:#f9f9f9;' : ''}">
                <td style="border:1px solid #ccc; padding:8px; text-align:center;">${index + 1}</td>
                <td style="border:1px solid #ccc; padding:8px;">${team}</td>
                <td style="border:1px solid #ccc; padding:8px; text-align:center;">${t.win}</td>
                <td style="border:1px solid #ccc; padding:8px; text-align:center;">${t.draw}</td>
                <td style="border:1px solid #ccc; padding:8px; text-align:center;">${t.lose}</td>
                <td style="border:1px solid #ccc; padding:8px; text-align:center;">${t.goal}</td>
                <td style="border:1px solid #ccc; padding:8px; text-align:center;">${t.concede}</td>
                <td style="border:1px solid #ccc; padding:8px; text-align:center; font-weight:bold; color:#e74c3c;">${t.score}</td>
            </tr>
            `;
        });

        html += `</tbody></table></div>`;
    });

    contentEl.innerHTML = html;
}
