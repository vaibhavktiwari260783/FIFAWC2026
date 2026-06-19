const EXCEL_FILE = 'FIFA World Cup Leaderboard - New.xlsx';

let participants = [];
let standings = [];
let activeGroup = '';

const el = (id) => document.getElementById(id);

function value(row, names) {
  for (const name of names) {
    if (row[name] !== undefined && row[name] !== null && row[name] !== '') return row[name];
  }
  return '';
}

function normaliseParticipant(row) {
  const team1 = value(row, ['Team 1', 'Team1']);
  const team2 = value(row, ['Team 2', 'Team2']);
  const team3 = value(row, ['Team 3', 'Team3']);
  const totalFromExcel = Number(value(row, ['Total Points', 'Points', 'Total'])) || 0;
  const numericCells = Object.values(row).filter(v => typeof v === 'number');
  const calculated = numericCells.length >= 4 ? Number(numericCells[numericCells.length - 1]) : totalFromExcel;
  return {
    participant: value(row, ['Participant', 'Name']),
    teams: [team1, team2, team3].filter(Boolean),
    totalPoints: totalFromExcel || calculated
  };
}

function normaliseStanding(row) {
  return {
    group: value(row, ['Group']),
    position: Number(value(row, ['Position', 'Pos'])) || 0,
    team: value(row, ['Team']),
    code: value(row, ['Code']),
    played: Number(value(row, ['Played', 'Pld'])) || 0,
    won: Number(value(row, ['Won', 'W'])) || 0,
    draw: Number(value(row, ['Draw', 'D'])) || 0,
    lost: Number(value(row, ['Lost', 'L'])) || 0,
    points: Number(value(row, ['Points', 'Pts'])) || 0,
    goalsFor: Number(value(row, ['GoalsFor', 'GF'])) || 0,
    goalsAgainst: Number(value(row, ['GoalsAgainst', 'GA'])) || 0,
    goalDifference: Number(value(row, ['GoalDifference', 'GD'])) || 0
  };
}

async function loadDefaultExcel() {
  const response = await fetch(`${EXCEL_FILE}?v=${Date.now()}`);
  if (!response.ok) throw new Error(`Could not load ${EXCEL_FILE}`);
  const arrayBuffer = await response.arrayBuffer();
  readWorkbook(arrayBuffer);
}

function readWorkbook(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: 'array' });
  const participantsSheet = workbook.Sheets['Participants'] || workbook.Sheets[workbook.SheetNames[0]];
  const standingsSheet = workbook.Sheets['standings'] || workbook.Sheets['Standings'] || workbook.Sheets[workbook.SheetNames[1]];

  participants = XLSX.utils.sheet_to_json(participantsSheet, { defval: '' })
    .map(normaliseParticipant)
    .filter(p => p.participant)
    .sort((a, b) => b.totalPoints - a.totalPoints || a.participant.localeCompare(b.participant));

  standings = XLSX.utils.sheet_to_json(standingsSheet, { defval: '' })
    .map(normaliseStanding)
    .filter(s => s.group && s.team)
    .sort((a, b) => a.group.localeCompare(b.group) || a.position - b.position);

  activeGroup = activeGroup || [...new Set(standings.map(s => s.group))][0] || '';
  renderAll();
  el('loadStatus').textContent = 'Excel loaded';
  el('lastUpdated').textContent = `Browser refresh: ${new Date().toLocaleString()}`;
}

function renderSummary() {
  const top = participants[0];
  const groups = new Set(standings.map(s => s.group)).size;
  const teams = new Set(standings.map(s => s.team)).size;
  el('summaryGrid').innerHTML = `
    <div class="kpi"><strong>${participants.length}</strong><span>Participants</span></div>
    <div class="kpi"><strong>${teams}</strong><span>Teams in standings</span></div>
    <div class="kpi"><strong>${groups}</strong><span>Groups</span></div>
    <div class="kpi"><strong>${top ? `${top.participant} (${top.totalPoints})` : '-'}</strong><span>Current leader</span></div>
  `;
}

function renderLeaderboard() {
  const search = el('searchInput').value.trim().toLowerCase();
  const rows = participants.filter(p => !search || p.participant.toLowerCase().includes(search) || p.teams.join(' ').toLowerCase().includes(search));
  el('leaderboardTable').innerHTML = `
    <thead><tr><th>Rank</th><th>Participant</th><th>Assigned Teams</th><th>Total Points</th></tr></thead>
    <tbody>${rows.map((p, index) => `
      <tr>
        <td class="rank">${index + 1}</td>
        <td><strong>${p.participant}</strong></td>
        <td>${p.teams.map(t => `<span class="team-pill">${t}</span>`).join('')}</td>
        <td class="points">${p.totalPoints}</td>
      </tr>`).join('')}</tbody>
  `;
}

function renderTabs() {
  const groups = [...new Set(standings.map(s => s.group))];
  el('groupTabs').innerHTML = groups.map(g => `<button class="tab ${g === activeGroup ? 'active' : ''}" data-group="${g}">${g}</button>`).join('');
  document.querySelectorAll('.tab').forEach(btn => btn.addEventListener('click', () => {
    activeGroup = btn.dataset.group;
    renderTabs();
    renderStandings();
  }));
}

function renderStandings() {
  const rows = standings.filter(s => s.group === activeGroup);
  el('standingsTable').innerHTML = `
    <thead><tr><th>Pos</th><th>Team</th><th>Code</th><th>Played</th><th>Won</th><th>Draw</th><th>Lost</th><th>Pts</th><th>GF</th><th>GA</th><th>GD</th></tr></thead>
    <tbody>${rows.map(s => `
      <tr>
        <td class="rank">${s.position}</td><td><strong>${s.team}</strong></td><td>${s.code}</td>
        <td>${s.played}</td><td>${s.won}</td><td>${s.draw}</td><td>${s.lost}</td>
        <td class="points">${s.points}</td><td>${s.goalsFor}</td><td>${s.goalsAgainst}</td><td>${s.goalDifference}</td>
      </tr>`).join('')}</tbody>
  `;
}

function renderAll() {
  renderSummary();
  renderLeaderboard();
  renderTabs();
  renderStandings();
}

el('refreshBtn').addEventListener('click', () => {
  el('loadStatus').textContent = 'Refreshing Excel…';
  loadDefaultExcel().catch(showError);
});
el('searchInput').addEventListener('input', renderLeaderboard);
el('fileInput').addEventListener('change', async (event) => {
  const file = event.target.files[0];
  if (!file) return;
  readWorkbook(await file.arrayBuffer());
  el('loadStatus').textContent = `Loaded: ${file.name}`;
});

function showError(error) {
  el('loadStatus').innerHTML = `<span class="error">${error.message}</span>`;
}

loadDefaultExcel().catch(showError);
