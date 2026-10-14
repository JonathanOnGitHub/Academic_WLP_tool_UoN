// ═══════════════════════════════════════════════════════
// SIMULATIONS TAB
// ═══════════════════════════════════════════════════════
let simAllData = [];

window.getSimHoursTotals = function() {
  const totals = {};
  simAllData.forEach(r => {
    if (r.name && r.hours > 0) totals[r.name] = (totals[r.name] || 0) + r.hours;
  });
  return totals;
};

// Default hours for the Y3 / Y4 sim half-day blocks. AM and PM are each 3h
// because the timetabled session spans 9-12 (or 9.30-12) and 1-4 (or 1.30-4).
const SIM_AM_HOURS = 3;
const SIM_PM_HOURS = 3;

// Parse the "Time" column of a Fishbowl row into hours. Inputs observed in
// 2026-27: "1-3pm", "2-4pm", "9-11am", "11-1pm" — all 2h sessions. Anything
// we can't parse defaults to 2h (the documented Fishbowl length).
function fishbowlHours(timeStr) {
  if (!timeStr) return 2;
  const m = String(timeStr).match(/(\d{1,2})(?::(\d{2}))?\s*-\s*(\d{1,2})(?::(\d{2}))?/);
  if (!m) return 2;
  const start = parseInt(m[1], 10) + (m[2] ? parseInt(m[2], 10) / 60 : 0);
  const end   = parseInt(m[3], 10) + (m[4] ? parseInt(m[4], 10) / 60 : 0);
  if (!isFinite(start) || !isFinite(end) || end <= start) return 2;
  return end - start;
}

function parseSimXlsx(file) {
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
      const data = [];

      for (const sheetName of wb.SheetNames) {
        const ws = wb.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });
        if (rows.length < 2) continue;

        const hdrs = rows[0].map(c => String(c).toLowerCase().trim());

        // 2026-27 layout differs between the two sheets. The AM/PM header
        // labels in row 1 are unreliable (the "PM" cell is also the first
        // PM staff slot in data rows), so detect the sheet layout by
        // looking for known landmark headers instead:
        //   * Sem 1 has a "Time to arrive at A25" column (idx 4)
        //   * Sem 2 has a standalone "Location" column (idx 5)
        const cDate   = hdrs.findIndex(h => h.includes('date'));
        const cDay    = hdrs.findIndex(h => h === 'day');
        const cYrGrp  = hdrs.findIndex(h => h.includes('year group') || h.includes('year'));
        const cType   = hdrs.findIndex(h => h === 'type');
        const hasLocation = cType > 0 && hdrs[cType - 1] === 'location';

        // Sem 1: AM at idx 7,8,9 and PM at idx 11,12,13 (idx 10 is blank).
        // Sem 2: AM at idx 8,9,10 and PM at idx 12,13,14 (idx 11 is blank).
        const layout = hasLocation
          ? { amStart: 8, pmStart: 12, slots: 3 }
          : { amStart: 7, pmStart: 11, slots: 3 };

        // Find the Time column for Fishbowl hour calculation (if any).
        const cTime = hdrs.findIndex(h => h === 'time' || h.startsWith('time '));

        if (cDate < 0) continue;

        for (let i = 1; i < rows.length; i++) {
          const r = rows[i];
          const rawDate = String(r[cDate] || '').trim();
          if (!rawDate) continue;
          // Skip duplicate header rows and footer notes (no numeric Week).
          const weekCell = String(r[0] || '').trim();
          if (!weekCell || !/^\d+$/.test(weekCell)) continue;

          const day     = cDay   >= 0 ? String(r[cDay]   || '').trim() : '';
          const yrGrp   = cYrGrp >= 0 ? String(r[cYrGrp] || '').trim() : '';
          const timeStr = cTime  >= 0 ? String(r[cTime]  || '').trim() : '';
          const typeStr = cType  >= 0 ? String(r[cType]  || '').trim() : '';
          const semester = sheetName;
          // No dedicated "Group" column in 2026-27; embed the session type
          // (e.g. "Y4 sim", "Fishbowl 1") so the activity string still
          // distinguishes them.
          const grp = typeStr;

          const isFishbowl = /^fishbowl/i.test(typeStr);
          const amHours = isFishbowl ? fishbowlHours(timeStr) : SIM_AM_HOURS;
          const pmHours = isFishbowl ? fishbowlHours(timeStr) : SIM_PM_HOURS;
          const typeLabel = isFishbowl
            ? `Fishbowl ${typeStr.replace(/^fishbowl\s*/i, '').trim() || ''}`.trim()
            : (typeStr || 'Simulation');

          function clean(n) {
            let s = String(n).trim();
            s = s.replace(/\s*\([^)]*\)\s*$/, '').trim();
            s = s.replace(/\s+DNA\s*$/i, '').trim();
            s = s.replace(/[,\s]+$/, '');
            return s;
          }
          function pushSlot(slot, hours) {
            const start = slot === 'AM' ? layout.amStart : layout.pmStart;
            for (let ci = start; ci < start + layout.slots && ci < r.length; ci++) {
              const raw = r[ci];
              if (raw === null || raw === undefined || String(raw).trim() === '') continue;
              const name = clean(raw);
              if (!name) continue;
              if (/^(9-12|1-4|noon|pm|am|fishbowl)/i.test(name)) continue;
              if (/^(no\s|none|tbc|dna|n\/a)$/i.test(name)) continue;
              if (!/[A-Za-z]/.test(name)) continue;
              data.push({
                name,
                date: rawDate,
                day,
                yearGroup: yrGrp,
                group: grp,
                timeSlot: slot,
                hours,
                semester,
                activity: `${typeLabel} (${semester}) — ${rawDate} ${yrGrp} ${slot}`.replace(/\s+/g, ' ').trim()
              });
            }
          }
          pushSlot('AM', amHours);
          pushSlot('PM', pmHours);
        }
      }

      if (data.length === 0) throw new Error('No valid data rows found. Ensure each sheet has columns: Week, Date, Day, Year Group and staff columns for AM and PM slots.');
      renderSim(data);
      document.getElementById('simError').classList.remove('show');
      WLP_SESSION.saveFile('simulation',file.name,e.target.result,file.type);
    } catch (err) {
      const el = document.getElementById('simError');
      el.textContent = 'Error: ' + err.message;
      el.classList.add('show');
    }
  };
  reader.readAsArrayBuffer(file);
}

function simGetSorted() {
  const q = document.getElementById('simSearch')?.value.toLowerCase() || '';
  let data = simAllData.filter(r =>
    r.name.toLowerCase().includes(q) || r.activity.toLowerCase().includes(q)
  );
  const sel = document.getElementById('simSortSel')?.value || 'hours-desc';
  const [sc, sd] = sel.split('-');
  // Aggregate by staff member
  const aggregated = [];
  const seen = new Set();
  data.forEach(r => {
    if (!seen.has(r.name)) {
      seen.add(r.name);
      const personRows = simAllData.filter(d => d.name === r.name);
      const totalHours = personRows.reduce((sum, d) => sum + d.hours, 0);
      aggregated.push({ name: r.name, hours: totalHours, sessions: personRows.length });
    }
  });
  aggregated.sort((a, b) => {
    if (sc === 'name') return sd === 'asc' ? a.name.localeCompare(b.name) : b.name.localeCompare(a.name);
    if (sc === 'sessions') return sd === 'asc' ? a.sessions - b.sessions : b.sessions - a.sessions;
    return sd === 'asc' ? a.hours - b.hours : b.hours - a.hours;
  });
  return aggregated;
}

function simOpenDetail(name) {
  const rows = simAllData.filter(r => r.name === name);
  if (rows.length === 0) return;
  const totalHours = rows.reduce((s, r) => s + r.hours, 0);
  let html = `<div class="panel-section">
    <h4>Summary</h4>
    <div class="panel-row"><span class="k">Sessions</span><span class="v">${rows.length}</span></div>
    <div class="panel-row"><span class="k">Total hours</span><span class="v big">${totalHours.toFixed(1)}h</span></div>
  </div>
  <div class="panel-section"><h4>Session detail</h4>`;
  // Sort by semester, then date
  const sorted = rows.slice().sort((a, b) => {
    if (a.semester !== b.semester) return a.semester.localeCompare(b.semester);
    return a.date.localeCompare(b.date);
  });
  sorted.forEach(r => {
    html += `<div style="padding:8px 0;border-bottom:1px solid var(--border)">
      <div style="font-weight:600;font-size:0.85rem;margin-bottom:4px">${r.semester} — ${r.date} ${r.day} (${r.yearGroup} ${r.group})</div>
      <div class="panel-row"><span class="k">Time</span><span class="v">${r.timeSlot} (${r.hours.toFixed(1)}h)</span></div>
    </div>`;
  });
  html += '</div>';
  openPanel(name, `${totalHours.toFixed(1)}h total · ${rows.length} session(s)`, html);
}

function renderSimTable() {
  const data = simGetSorted();
  const maxHours = Math.max(...data.map(r => r.hours), 1);
  document.getElementById('simTbody').innerHTML = data.map(r => {
    const barPct = Math.min(r.hours / maxHours * 100, 100).toFixed(1);
    return `<tr>
      <td class="name-f" data-name="${encodeURIComponent(r.name)}" style="cursor:pointer">${r.name}</td>
      <td class="num">${r.sessions}</td>
      <td class="num">${r.hours.toFixed(1)}</td>
      <td><div class="pgr-hrs-bar-wrap"><div class="pgr-hrs-bar"><div class="pgr-hrs-bar-fill" style="width:${barPct}%" title="${r.hours.toFixed(1)} hours"></div></div></td>
    </tr>`;
  }).join('');
  document.querySelectorAll('#simTbody .name-f').forEach(cell => {
    cell.addEventListener('click', () => simOpenDetail(decodeURIComponent(cell.dataset.name)));
  });
  const totalHours = data.reduce((s, r) => s + r.hours, 0);
  const totalSessions = data.reduce((s, r) => s + r.sessions, 0);
  document.getElementById('simFoot').innerHTML = `<tr>
    <td><strong>Total (${data.length} staff)</strong></td>
    <td class="num"><strong>${totalSessions}</strong></td>
    <td class="num"><strong>${totalHours.toFixed(1)}</strong></td>
    <td></td>
  </tr>`;
}

function renderSim(data) {
  simAllData = data;
  const totalHours = data.reduce((s, r) => s + r.hours, 0);
  const totalStaff = new Set(data.map(r => r.name)).size;
  document.getElementById('simStatsBar').innerHTML = `
    <div class="stat-card" style="border-left:3px solid var(--mid-blue)"><div class="sc-v">${totalStaff}</div><div class="sc-l">Staff with simulation sessions</div></div>
    <div class="stat-card" style="border-left:3px solid var(--gold)"><div class="sc-v">${totalHours.toFixed(0)}</div><div class="sc-l">Total Simulation Hours</div></div>
    <div class="stat-card" style="border-left:3px solid var(--teal)"><div class="sc-v">${data.length}</div><div class="sc-l">Session entries (AM/PM)</div></div>
  `;
  renderSimTable();
  document.getElementById('sim-landing').style.display = 'none';
  document.getElementById('sim-content').style.display = 'block';
  document.getElementById('badge-simulation').textContent = totalStaff;
  document.getElementById('simMeta').textContent = `${totalStaff} staff · ${totalHours.toFixed(1)} hours · ${data.length} entries`;
  updateCombStatus();
}

document.addEventListener('DOMContentLoaded', () => {
  const inp = document.getElementById('simFileInput');
  if (inp) inp.addEventListener('change', e => { if (e.target.files[0]) parseSimXlsx(e.target.files[0]); });

  const dz = document.getElementById('simDropZone');
  if (dz) {
    dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('drag-over'); });
    dz.addEventListener('dragleave', () => dz.classList.remove('drag-over'));
    dz.addEventListener('drop', e => { e.preventDefault(); dz.classList.remove('drag-over'); const f = e.dataTransfer.files[0]; if (f) parseSimXlsx(f); });
  }

  // Wire the analyse button (loads if file already picked, otherwise prompts)
  const analyseBtn = document.getElementById('simAnalyseBtn');
  if (analyseBtn) {
    analyseBtn.addEventListener('click', () => {
      // If we already have data from a file drop, simulate a click
      const fi = document.getElementById('simFileInput');
      if (fi && fi.files[0]) parseSimXlsx(fi.files[0]);
    });
  }

  const back = document.getElementById('simBtnBack');
  if (back) back.addEventListener('click', () => {
    document.getElementById('sim-content').style.display = 'none';
    document.getElementById('sim-landing').style.display = 'block';
    simAllData = [];
    document.getElementById('badge-simulation').textContent = '—';
    updateCombStatus();
  });

  const sortSel = document.getElementById('simSortSel');
  if (sortSel) sortSel.addEventListener('change', renderSimTable);
  const search = document.getElementById('simSearch');
  if (search) search.addEventListener('input', renderSimTable);

  document.addEventListener('click', e => {
    const th = e.target.closest('th[data-simsort]');
    if (!th) return;
    const col = th.dataset.simsort;
    const sel = document.getElementById('simSortSel');
    if (!sel) return;
    const isNum = ['hours', 'sessions'].includes(col);
    const dir = (sel.value.startsWith(col) && sel.value.endsWith('desc')) ? 'asc' : (isNum ? 'desc' : 'asc');
    const opt = [...sel.options].find(o => o.value === `${col}-${dir}`);
    if (opt) { sel.value = opt.value; renderSimTable(); }
  });

  const expBtn = document.getElementById('simBtnExport');
  if (expBtn) expBtn.addEventListener('click', () => {
    if (simAllData.length === 0) return;
    const hdr = ['Staff name', 'Semester', 'Date', 'Day', 'Year Group', 'Group', 'Time slot', 'Hours', 'Activity'];
    const rows = [hdr];
    const sorted = simAllData.slice().sort((a, b) => a.name.localeCompare(b.name) || a.date.localeCompare(b.date));
    sorted.forEach(r => rows.push([r.name, r.semester, r.date, r.day, r.yearGroup, r.group, r.timeSlot, r.hours, r.activity]));
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Simulations');
    XLSX.writeFile(wb, 'simulation_hours.xlsx');
  });
});
