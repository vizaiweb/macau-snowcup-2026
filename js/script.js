// 全局變量
let excelData = {};
const DATA_FILE_PATH = 'matches_data.xlsx';

// 頁面載入
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

// ========== 數據加載（動態取得最後修改時間） ==========
function loadExcelData() {
    fetch(DATA_FILE_PATH + '?t=' + Date.now())
        .then(response => {
            if (!response.ok) {
                return response.text().then(text => {
                    throw new Error(`HTTP ${response.status} ${response.statusText} - 回應內容開頭: ${text.substring(0, 100)}`);
                });
            }

            // 取得檔案的 Last-Modified 時間
            const lastModified = response.headers.get('last-modified');
            if (lastModified) {
                const lastModifiedDate = new Date(lastModified);
                // 格式化為澳門/香港慣用格式（例如 17/3/2026 上午11:27:40）
                document.getElementById('update-time').textContent = lastModifiedDate.toLocaleString('zh-Hant-MO', {
                    year: 'numeric',
                    month: 'numeric',
                    day: 'numeric',
                    hour: '2-digit',
                    minute: '2-digit',
                    second: '2-digit',
                    hour12: true
                });
            } else {
                // 若無 last-modified，則使用當前時間（或保留原訊息）
                document.getElementById('update-time').textContent = new Date().toLocaleString('zh-Hant-MO');
            }

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
        })
        .catch(error => {
            const errorHtml = `
                <div style="padding:20px; color:#e74c3c; text-align:center;">
                    <h4>數據加載失敗！</h4>
                    <p>錯誤詳情：${error.message}</p>
                    <p>請確認：</p>
                    <ul style="list-style:none; padding:0;">
                        <li>1. Excel 檔案 <strong>${DATA_FILE_PATH}</strong> 是否存在於網站根目錄？</li>
                        <li>2. 若在本機直接開啟 HTML，請改用本地伺服器（如 Live Server）。</li>
                        <li>3. 工作表名稱是否符合預期（例如「初級組_對賽安排」）？</li>
                    </ul>
                </div>`;
            document.querySelectorAll('#matches-content, #results-content, #rankings-content').forEach(el => {
                el.innerHTML = errorHtml;
            });
        });
}

// ========== 事件綁定（完全與先前相同） ==========
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

// ========== 渲染對賽安排（與先前相同） ==========
function renderMatches(group) {
    const contentEl = document.getElementById('matches-content');
    const sheetName = `${group}_對賽安排`;
    const matchesData = excelData[sheetName] || [];

    if (matchesData.length === 0) {
        contentEl.innerHTML = `<div class="loading">${group} 暫無對賽安排數據</div>`;
        return;
    }

    let html = `
    <table>
        <thead>
            <tr>
                <th>日期</th><th>時間</th><th>隊伍A</th><th>隊伍B</th><th>場地</th><th>組別</th><th>比分</th><th>備註</th>
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

        html += `
        <tr${score !== '未開賽' ? ' style="background:#f0f9ff;"' : ''}>
            <td>${date}</td><td>${time}</td><td>${teamA}</td><td>${teamB}</td>
            <td>${venue}</td><td>${subgroup}</td><td><strong>${score}</strong></td><td>${remark}</td>
        </tr>
        `;
    });

    html += `</tbody></table>`;
    contentEl.innerHTML = html;
}

// ========== 渲染對賽成績（與先前相同） ==========
function renderResults(group) {
    const contentEl = document.getElementById('results-content');
    const sheetName = `${group}_對賽安排`;
    const matchesData = excelData[sheetName] || [];

    if (matchesData.length === 0) {
        contentEl.innerHTML = `<div class="loading">${group} 暫無對賽數據</div>`;
        return;
    }

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
            <h3 style="margin:0 0 8px; font-size:16px;">${subgroup} 組 - 對賽成績</h3>
            <table>
                <thead>
                    <tr>
                        <th>日期</th><th>時間</th><th>隊伍A</th><th>比分</th><th>隊伍B</th><th>場地</th>
                    </tr>
                </thead>
                <tbody>
        `;

        grouped[subgroup].forEach(item => {
            html += `
            <tr>
                <td>${item['日期'] || '-'}</td>
                <td>${excelTimeToHHMM(item['時間'])}</td>
                <td>${item['隊伍A'] || '-'}</td>
                <td><strong>${item['比分'] || '-'}</strong></td>
                <td>${item['隊伍B'] || '-'}</td>
                <td>${item['場地'] || '-'}</td>
            </tr>
            `;
        });

        html += `</tbody></table></div>`;
    });

    contentEl.innerHTML = html;
}

// ========== 渲染積分榜（與先前相同） ==========
function renderRankings(group) {
    const contentEl = document.getElementById('rankings-content');
    const sheetName = `${group}_對賽安排`;
    const matchesData = excelData[sheetName] || [];

    if (matchesData.length === 0) {
        contentEl.innerHTML = `<div class="loading">${group} 暫無比賽數據，無法產生積分榜</div>`;
        return;
    }

    const groupStats = {};
    matchesData.forEach(item => {
        const teamA = item['隊伍A'] || '-';
        const teamB = item['隊伍B'] || '-';
        const subgroup = item['組別'] || '未分組';
        if (!groupStats[subgroup]) groupStats[subgroup] = {};
        if (!groupStats[subgroup][teamA]) groupStats[subgroup][teamA] = { win:0, draw:0, lose:0, goal:0, concede:0, score:0 };
        if (!groupStats[subgroup][teamB]) groupStats[subgroup][teamB] = { win:0, draw:0, lose:0, goal:0, concede:0, score:0 };
    });

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
        if (subgroup === '補賽') return;

        const sortedTeams = Object.keys(groupStats[subgroup]).sort((a, b) => {
            const ta = groupStats[subgroup][a], tb = groupStats[subgroup][b];
            if (tb.score !== ta.score) return tb.score - ta.score;
            return (tb.goal - tb.concede) - (ta.goal - ta.concede);
        });

        html += `
        <div style="margin-bottom:24px;">
            <h3 style="margin:0 0 8px; font-size:16px;">${subgroup} 組 - 積分榜</h3>
            <table class="leaderboard">
                <thead>
                    <tr>
                        <th>排名</th><th>隊伍名稱</th><th>勝</th><th>平</th><th>負</th><th>進球</th><th>失球</th><th>積分</th>
                    </tr>
                </thead>
                <tbody>
        `;

        sortedTeams.forEach((team, index) => {
            const t = groupStats[subgroup][team];
            html += `
            <tr ${index % 2 === 0 ? 'style="background:#f9f9f9;"' : ''}>
                <td>${index + 1}</td>
                <td>${team}</td>
                <td>${t.win}</td>
                <td>${t.draw}</td>
                <td>${t.lose}</td>
                <td>${t.goal}</td>
                <td>${t.concede}</td>
                <td style="font-weight:bold; color:#e74c3c;">${t.score}</td>
            </tr>
            `;
        });

        html += `</tbody></table></div>`;
    });

    contentEl.innerHTML = html || '<div class="loading">無有效組別資料</div>';
}
