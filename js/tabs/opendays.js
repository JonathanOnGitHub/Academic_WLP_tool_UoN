// ═══════════════════════════════════════════════════════
// OPEN DAYS TAB
// ═══════════════════════════════════════════════════════
let odAllData = [];

window.getOpenDaysHoursTotals = function() {
  const totals = {};
  odAllData.forEach(r => {
    if (r.name && r.hours > 0) totals[r.name] = (totals[r.name] || 0) + r.hours;
  });
  return totals;
};

/** Strip bracket content and anything after from a name */
function odCleanName(raw) {
  const s = String(raw || '').trim();
  // Remove everything from the first '(' onwards (bracket content + trailing text)
  const idx = s.indexOf('(');
  return idx >= 0 ? s.slice(0, idx).trim() : s;
}

/** Parse a time string like "12-2pm", "13.00-14.00", "9-5", or extract from "Staff volunteers (12-2pm)" */
function odParseTime(cellText) {
  const s = String(cellText || '').trim();
  if (!s) return 0;

  // Try to extract time pattern from inside brackets first
  let timeStr = s;
  const bracketMatch = s.match(/\(([^)]+)\)/);
  if (bracketMatch) timeStr = bracketMatch[1].trim();

  // Normalise dots to colons but only in time-like patterns
  timeStr = timeStr.replace(/\.(\d)/g, ':$1');

  // Strip "pm" / "am" markers (store for later adjustment)
  let pm = false;
  if (/pm/i.test(timeStr)) { pm = true; timeStr = timeStr.replace(/pm/i, '').trim(); }
  timeStr = timeStr.replace(/am/i, '').trim();

  // Try range format: "12-2", "11-12:30", "9-5", "13.00-14.00"
  const rangeMatch = timeStr.match(/^(\d{1,2})(?::(\d{2}))?\s*[-–]\s*(\d{1,2})(?::(\d{2}))?$/);
  if (rangeMatch) {
    let startH = parseInt(rangeMatch[1], 10);
    const startM = parseInt(rangeMatch[2] || '0', 10);
    let endH = parseInt(rangeMatch[3], 10);
    const endM = parseInt(rangeMatch[4] || '0', 10);

    // If pm suffix or end <= start and start >= 7, assume PM
    if (pm || (endH <= startH && startH >= 7)) endH += 12;
    // Also handle 12-2 (noon-2pm)
    if (startH === 12 && endH <= 7) endH += 12;

    const startDec = startH + startM / 60;
    const endDec = endH + endM / 60;
    return Math.max(0, endDec - startDec);
  }

  // Try plain decimal
  const decMatch = timeStr.match(/^(\d+(?:\.\d+)?)\s*(?:h|hr|hrs)?$/);
  if (decMatch) return parseFloat(decMatch[1]);

  return 0;
}

function parseOpenDaysXlsx(file) {
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

      // Need at least 3 rows (header + date + at least one name)
      if (rows.length < 3) throw new Error('File has too few rows. Need times (row 1), dates (row 2), and staff names.');

      // Columns D:G = indices 3,4,5,6
      const cols = [3, 4, 5, 6];
      const data = [];

      for (const ci of cols) {
        const timeCell = rows[0]?.[ci] || '';
        const dateCell = rows[1]?.[ci] || '';
        const hours = odParseTime(timeCell);
        if (hours <= 0) continue;

        for (let ri = 2; ri < rows.length; ri++) {
          const nameCell = String(rows[ri]?.[ci] || '').trim();
          if (!nameCell) continue;
          const name = odCleanName(nameCell);
          if (!name) continue;
          data.push({
            name,
            date: String(dateCell).trim(),
            time: String(timeCell).trim(),
            hours,
            sessionLabel: `${dateCell} (${hours}h)`
          });
        }
      }

      if (data.length === 0) throw new Error('No valid data found. Check that columns D–G contain times (row 1), dates (row 2), and staff names.');

      renderOpenDays(data);
      document.getElementById('odError').classList.remove('show');
      WLP_SESSION.saveFile('opendays', file.name, e.target.result, file.type);
    } catch (err) {
      const el = document.getElementById('odError');
      el.textContent = 'Error: ' + err.message;
      el.classList.add('show');
    }
  };
  reader.readAsArrayBuffer(file);
}

// ── Aggregation, sorting, search ─────────────────────

function odGetSorted() {
  const q = document.getElementById('odSearch')?.value.toLowerCase() || '';
  let data = odAllData.filter(r =>
    r.name.toLowerCase().includes(q)
  );
  const sel = document.getElementById('odSortSel')?.value || 'hours-desc';
  const [sc, sd] = sel.split('-');

  const aggregated = [];
  const seen = new Set();
  data.forEach(r => {
    if (!seen.has(r.name)) {
      seen.add(r.name);
      const personRows = odAllData.filter(d => d.name === r.name);
      const totalHours = personRows.reduce((sum, d) => sum + d.hours, 0);
      aggregated.push({ name: r.name, hours: totalHours, sessions: personRows.length });
    }
  });
  aggregated.sort((a, b) => {
    if (sc === 'name') return sd === 'asc' ? a.name.localeCompare(b.name) : b.name.localeCompare(a.name);
    return sd === 'asc' ? a.hours - b.hours : b.hours - a.hours;
  });
  return aggregated;
}

function odOpenDetail(name) {
  const rows = odAllData.filter(r => r.name === name);
  if (rows.length === 0) return;
  const totalHours = rows.reduce((s, r) => s + r.hours, 0);
  let html = `<div class="panel-section">
    <h4>Summary</h4>
    <div class="panel-row"><span class="k">Sessions</span><span class="v">${rows.length}</span></div>
    <div class="panel-row"><span class="k">Total hours</span><span class="v big">${totalHours.toFixed(1)}h</span></div>
  </div>
  <div class="panel-section"><h4>Session detail</h4>`;
  rows.forEach(r => {
    html += `<div style="padding:8px 0;border-bottom:1px solid var(--border)">
      <div style="font-weight:600;font-size:0.85rem;margin-bottom:4px">${r.date || '—'}</div>
      <div class="panel-row"><span class="k">Time</span><span class="v">${r.time || '—'}</span></div>
      <div class="panel-row"><span class="k">Hours</span><span class="v big" style="color:var(--teal)">${r.hours.toFixed(1)}h</span></div>
    </div>`;
  });
  html += '</div>';
  openPanel(name, `${totalHours.toFixed(1)}h total · ${rows.length} session(s)`, html);
}

function renderOpenDaysTable() {
  const data = odGetSorted();
  const maxHours = Math.max(...data.map(r => r.hours), 1);
  document.getElementById('odTbody').innerHTML = data.map(r => {
    const barPct = Math.min(r.hours / maxHours * 100, 100).toFixed(1);
    return `<tr>
      <td class="name-f" data-name="${encodeURIComponent(r.name)}" style="cursor:pointer">${r.name}</td>
      <td class="num">${r.sessions}</td>
      <td class="num">${r.hours.toFixed(1)}</td>
      <td><div class="pgr-hrs-bar-wrap"><div class="pgr-hrs-bar"><div class="pgr-hrs-bar-fill" style="width:${barPct}%" title="${r.hours.toFixed(1)} hours"></div></div></td>
    </tr>`;
  }).join('');
  document.querySelectorAll('#odTbody .name-f').forEach(cell => {
    cell.addEventListener('click', () => odOpenDetail(decodeURIComponent(cell.dataset.name)));
  });
  const totalHours = data.reduce((s, r) => s + r.hours, 0);
  const totalSessions = data.reduce((s, r) => s + r.sessions, 0);
  document.getElementById('odFoot').innerHTML = `<tr>
    <td><strong>Total (${data.length} staff)</strong></td>
    <td class="num"><strong>${totalSessions}</strong></td>
    <td class="num"><strong>${totalHours.toFixed(1)}</strong></td>
    <td></td>
  </tr>`;
}

function renderOpenDays(data) {
  odAllData = data;
  const totalHours = data.reduce((s, r) => s + r.hours, 0);
  const totalStaff = new Set(data.map(r => r.name)).size;
  document.getElementById('odStatsBar').innerHTML = `
    <div class="stat-card" style="border-left:3px solid var(--teal)"><div class="sc-v">${totalStaff}</div><div class="sc-l">Staff with Open Day duties</div></div>
    <div class="stat-card" style="border-left:3px solid var(--gold)"><div class="sc-v">${totalHours.toFixed(0)}</div><div class="sc-l">Total Open Day Hours</div></div>
    <div class="stat-card" style="border-left:3px solid var(--mid-blue)"><div class="sc-v">${data.length}</div><div class="sc-l">Session entries</div></div>
  `;
  renderOpenDaysTable();
  document.getElementById('opendays-landing').style.display = 'none';
  document.getElementById('opendays-content').style.display = 'block';
  document.getElementById('badge-opendays').textContent = totalStaff;
  document.getElementById('odMeta').textContent = `${totalStaff} staff · ${totalHours.toFixed(1)} hours · ${data.length} entries`;
  updateCombStatus();
}

document.addEventListener('DOMContentLoaded', () => {
  const inp = document.getElementById('odFileInput');
  if (inp) inp.addEventListener('change', e => { if (e.target.files[0]) parseOpenDaysXlsx(e.target.files[0]); });

  const dz = document.getElementById('odDropZone');
  if (dz) {
    dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('drag-over'); });
    dz.addEventListener('dragleave', () => dz.classList.remove('drag-over'));
    dz.addEventListener('drop', e => { e.preventDefault(); dz.classList.remove('drag-over'); const f = e.dataTransfer.files[0]; if (f) parseOpenDaysXlsx(f); });
  }

  const back = document.getElementById('odBtnBack');
  if (back) back.addEventListener('click', () => {
    document.getElementById('opendays-content').style.display = 'none';
    document.getElementById('opendays-landing').style.display = 'block';
    odAllData = [];
    document.getElementById('badge-opendays').textContent = '—';
    updateCombStatus();
  });

  const sortSel = document.getElementById('odSortSel');
  if (sortSel) sortSel.addEventListener('change', renderOpenDaysTable);
  const search = document.getElementById('odSearch');
  if (search) search.addEventListener('input', renderOpenDaysTable);

  // Header click sorting
  document.addEventListener('click', e => {
    const th = e.target.closest('th[data-odsort]');
    if (!th) return;
    const col = th.dataset.odsort;
    const sel = document.getElementById('odSortSel');
    if (!sel) return;
    const isNum = ['hours', 'sessions'].includes(col);
    const dir = (sel.value.startsWith(col) && sel.value.endsWith('desc')) ? 'asc' : (isNum ? 'desc' : 'asc');
    const opt = [...sel.options].find(o => o.value === `${col}-${dir}`);
    if (opt) { sel.value = opt.value; renderOpenDaysTable(); }
  });

  const expBtn = document.getElementById('odBtnExport');
  if (expBtn) expBtn.addEventListener('click', () => {
    if (odAllData.length === 0) return;
    const hdr = ['Staff member', 'Date', 'Time', 'Hours'];
    const rows = [hdr];
    odAllData.forEach(r => rows.push([r.name, r.date, r.time, r.hours]));
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Open Days');
    XLSX.writeFile(wb, 'open_days.xlsx');
  });
});
