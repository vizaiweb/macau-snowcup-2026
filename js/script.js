function renderRankings(group) {
    const contentEl = document.getElementById('rankings-content');
    const sheetName = `${group}_對賽安排`;
    const matchesData = excelData[sheetName] || [];
  
    // 收集所有隊伍（即使沒比數）
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
  
    // 對有比數的場次進行積分統計
    matchesData.forEach(item => {
        const teamA = item['隊伍A'] || '-';
        const teamB = item['隊伍B'] || '-';
        const subgroup = item['組別'] || '未分組';
        const score = item['比分'];
        // 僅計算標準比數（數字-數字）
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
        const sortedTeams = Object.keys(groupStats[subgroup]).sort((a, b) => {
            const ta = groupStats[subgroup][a], tb = groupStats[subgroup][b];
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
