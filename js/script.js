// 渲染積分榜 —— 包含棄權3:0規則 + 完整賽制
function renderRankings(group) {
    const contentEl = document.getElementById('rankings-content');
    const matchesData = excelData[`${group}_對賽安排`] || [];

    if (matchesData.length === 0) {
        contentEl.innerHTML = `<div class="loading">${group}暫無積分數據</div>`;
        return;
    }

    // 1. 處理所有賽事（包含正常賽事 + 棄權賽事）
    const processedMatches = [];
    matchesData.forEach(item => {
        const teamA = item['隊伍A'] || '-';
        const teamB = item['隊伍B'] || '-';
        const score = item['比分'] || '';
        const remark = item['備註'] || ''; // 用備註標註棄權：如「A隊棄權」「B隊棄權」
        const subgroup = item['組別'] || '未分組';

        // 情況1：正常賽事（有比分且不是--）
        if (score.trim() !== '' && score.trim() !== '--' && score.includes('-')) {
            processedMatches.push({
                teamA, teamB,
                scoreA: parseInt(score.split('-')[0]) || 0,
                scoreB: parseInt(score.split('-')[1]) || 0,
                subgroup,
                isForfeit: false,
                forfeitTeam: '' // 記錄棄權隊伍
            });
        }
        // 情況2：棄權賽事（備註標註棄權 或 比分為--）
        else if (remark.includes('棄權') || score.trim() === '--') {
            let scoreA = 0, scoreB = 0, forfeitTeam = '';
            // A隊棄權 → B隊3:0勝
            if (remark.includes('A隊棄權') || remark.includes('隊伍A棄權')) {
                scoreA = 0;
                scoreB = 3;
                forfeitTeam = teamA;
            }
            // B隊棄權 → A隊3:0勝
            else if (remark.includes('B隊棄權') || remark.includes('隊伍B棄權')) {
                scoreA = 3;
                scoreB = 0;
                forfeitTeam = teamB;
            }
            // 未標註具體隊伍 → 默認按章程補充規則（需手動確認）
            else {
                scoreA = 3;
                scoreB = 0;
                forfeitTeam = teamB;
            }
            processedMatches.push({
                teamA, teamB, scoreA, scoreB,
                subgroup,
                isForfeit: true,
                forfeitTeam
            });
        }
    });

    if (processedMatches.length === 0) {
        contentEl.innerHTML = `<div class="loading">${group}暫無已完賽/棄權賽事，積分榜待更新</div>`;
        return;
    }

    // 2. 初始化數據結構
    const groupStats = {}; // { 組別: { 隊伍名: { ...stats } } }
    const forfeitTeams = new Set(); // 記錄有棄權記錄的隊伍

    // 3. 計算積分/勝負/進失球
    processedMatches.forEach(match => {
        const { teamA, teamB, scoreA, scoreB, subgroup, isForfeit, forfeitTeam } = match;

        if (!teamA || !teamB || teamA === '-' || teamB === '-') return;

        // 初始化組別和隊伍
        if (!groupStats[subgroup]) groupStats[subgroup] = {};
        ['win', 'draw', 'lose', 'goal', 'concede', 'score', 'headToHead', 'forfeitCount'].forEach(key => {
            if (!groupStats[subgroup][teamA]) groupStats[subgroup][teamA] = {
                win: 0, draw: 0, lose: 0,
                goal: 0, concede: 0, score: 0,
                headToHead: {},
                forfeitCount: 0 // 棄權次數
            };
            if (!groupStats[subgroup][teamB]) groupStats[subgroup][teamB] = {
                win: 0, draw: 0, lose: 0,
                goal: 0, concede: 0, score: 0,
                headToHead: {},
                forfeitCount: 0
            };
        });

        // 標記棄權隊伍
        if (isForfeit && forfeitTeam) {
            groupStats[subgroup][forfeitTeam].forfeitCount += 1;
            forfeitTeams.add(forfeitTeam);
        }

        // 更新進失球
        groupStats[subgroup][teamA].goal += scoreA;
        groupStats[subgroup][teamA].concede += scoreB;
        groupStats[subgroup][teamB].goal += scoreB;
        groupStats[subgroup][teamB].concede += scoreA;

        // 記錄對賽數據
        groupStats[subgroup][teamA].headToHead[teamB] = { a: scoreA, b: scoreB };
        groupStats[subgroup][teamB].headToHead[teamA] = { a: scoreB, b: scoreA };

        // 計算積分（勝3/平1/負0）
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

    // 4. 按賽制排序（優先：棄權次數→積分→對賽成績→淨勝球→進球數）
    let html = '';
    Object.keys(groupStats).forEach(subgroup => {
        const teams = groupStats[subgroup];
        const sortedTeams = Object.keys(teams).sort((a, b) => {
            const ta = teams[a], tb = teams[b];

            // 規則1：棄權次數少者在前
            if (ta.forfeitCount !== tb.forfeitCount) return ta.forfeitCount - tb.forfeitCount;

            // 規則2：積分高者在前
            if (tb.score !== ta.score) return tb.score - ta.score;

            // 規則3：對賽成績（勝場→淨勝球→進球數）
            const h2h = ta.headToHead[b];
            if (h2h) {
                if (h2h.a > h2h.b) return -1;
                if (h2h.a < h2h.b) return 1;
                const netA = h2h.a - h2h.b;
                const netB = h2h.b - h2h.a;
                if (netA !== netB) return netB - netA;
                if (h2h.a !== h2h.b) return h2h.b - h2h.a;
            }

            // 規則4：總淨勝球
            const netA = ta.goal - ta.concede;
            const netB = tb.goal - tb.concede;
            if (netA !== netB) return netB - netA;

            // 規則5：總進球數
            if (ta.goal !== tb.goal) return tb.goal - ta.goal;

            // 規則6：抽籤
            return 0;
        });

        // 5. 生成積分表
        html += `
        <div style="margin-bottom:24px;">
            <h3 style="margin:0 0 8px; font-size:16px; font-weight:bold;">${subgroup} 組積分榜</h3>
            <table style="width:100%; border-collapse: collapse;">
                <thead>
                    <tr style="background:#3498db; color:white;">
                        <th style="border:1px solid #ccc; padding:8px;">排名</th>
                        <th style="border:1px solid #ccc; padding:8px;">隊伍名稱</th>
                        <th style="border:1px solid #ccc; padding:8px;">勝</th>
                        <th style="border:1px solid #ccc; padding:8px;">平</th>
                        <th style="border:1px solid #ccc; padding:8px;">負</th>
                        <th style="border:1px solid #ccc; padding:8px;">進球</th>
                        <th style="border:1px solid #ccc; padding:8px;">失球</th>
                        <th style="border:1px solid #ccc; padding:8px;">淨勝球</th>
                        <th style="border:1px solid #ccc; padding:8px;">積分</th>
                        <th style="border:1px solid #ccc; padding:8px;">棄權次數</th>
                    </tr>
                </thead>
                <tbody>
        `;

        sortedTeams.forEach((team, index) => {
            const t = teams[team];
            html += `
            <tr style="${index % 2 === 0 ? 'background:#f9f9f9;' : ''}">
                <td style="border:1px solid #ccc; padding:8px; text-align:center;">${index + 1}</td>
                <td style="border:1px solid #ccc; padding:8px;">${team}</td>
                <td style="border:1px solid #ccc; padding:8px; text-align:center;">${t.win}</td>
                <td style="border:1px solid #ccc; padding:8px; text-align:center;">${t.draw}</td>
                <td style="border:1px solid #ccc; padding:8px; text-align:center;">${t.lose}</td>
                <td style="border:1px solid #ccc; padding:8px; text-align:center;">${t.goal}</td>
                <td style="border:1px solid #ccc; padding:8px; text-align:center;">${t.concede}</td>
                <td style="border:1px solid #ccc; padding:8px; text-align:center;">${t.goal - t.concede}</td>
                <td style="border:1px solid #ccc; padding:8px; text-align:center; font-weight:bold; color:#e74c3c;">${t.score}</td>
                <td style="border:1px solid #ccc; padding:8px; text-align:center; color:${t.forfeitCount > 0 ? '#e74c3c' : '#2ecc71'};">${t.forfeitCount}</td>
            </tr>
            `;
        });

        html += `
                </tbody>
            </table>
            <p style="font-size:12px; color:#666; margin-top:4px;">* 棄權按3:0計算；同分時棄權次數少者在前，仍相同則抽籤</p>
        </div>
        `;
    });

    contentEl.innerHTML = html;
}
