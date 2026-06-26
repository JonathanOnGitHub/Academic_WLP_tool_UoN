// ═══════════════════════════════════════════════════════
// WELCOME WEEK TAB
// ═══════════════════════════════════════════════════════
let wwAllData = [];

window.getWelcomeWeekHoursTotals = function() {
  const totals = {};
  wwAllData.forEach(r => {
    if (r.name && r.hours > 0) totals[r.name] = (totals[r.name] || 0) + r.hours;
  });
  return totals;
};

function parseWwXlsx(file) {
  const reader = new FileReader();
  reader.onload = e => {
    try {
      let raw;
      if (file.name.match(/\.csv$/i)) {
        const text = new TextDecoder().decode(e.target.result);
        const lines = text.split(/\r?\n/).filter(l => l.trim());
        raw = lines.map(l => l.split(',').map(c => c.replace(/^"|"$/g, '').trim()));
      } else {
        const wb = XLSX.read(e.target.result, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      }

      if (raw.length < 3) throw new Error('File has too few rows. Expected headers on row 3.');

      // Header row is row 3 (0-indexed: raw[2])
      const hdrIdx = 2;
      const hdrs = raw[hdrIdx].map(c => String(c).toLowerCase().trim());

      function col(keywords) {
        const idx = hdrs.findIndex(h => keywords.some(k => h.includes(k)));
        return idx >= 0 ? idx : -1;
      }
      const cStaff   = col(['staff member', 'staff']);
      const cActivity = col(['activity']);
      const cHours   = col(['hours', 'hrs']);
      const cNotes   = col(['notes']);

      if (cStaff < 0) throw new Error('Could not find "Staff member" column in row 3.');
      if (cActivity < 0) throw new Error('Could not find "Activity" column in row 3.');
      if (cHours < 0) throw new Error('Could not find "Hours" column in row 3.');

      const data = [];
      for (let i = hdrIdx + 1; i < raw.length; i++) {
        const r = raw[i];
        const name = String(r[cStaff] || '').trim();
        const activity = String(r[cActivity] || '').trim();
        const rawHours = parseFloat(r[cHours]);
        if (!name || !activity || isNaN(rawHours) || rawHours <= 0) continue;
        const notes = cNotes >= 0 ? String(r[cNotes] || '').trim() : '';
        data.push({ name, activity, hours: rawHours, notes });
      }

      if (data.length === 0) throw new Error('No valid data rows found.');
      renderWw(data);
      document.getElementById('wwError').classList.remove('show');
    } catch (err) {
      const el = document.getElementById('wwError');
      el.textContent = 'Error: ' + err.message;
      el.classList.add('show');
    }
  };
  reader.readAsArrayBuffer(file);
}

function wwGetSorted() {
  const q = document.getElementById('wwSearch')?.value.toLowerCase() || '';
  let data = wwAllData.filter(r =>
    r.name.toLowerCase().includes(q) || r.activity.toLowerCase().includes(q)
  );
  const sel = document.getElementById('wwSortSel')?.value || 'hours-desc';
  const [sc, sd] = sel.split('-');
  // Aggregate by staff member
  const aggregated = [];
  const seen = new Set();
  data.forEach(r => {
    if (!seen.has(r.name)) {
      seen.add(r.name);
      const personRows = wwAllData.filter(d => d.name === r.name);
      const totalHours = personRows.reduce((sum, d) => sum + d.hours, 0);
      aggregated.push({ name: r.name, hours: totalHours, activities: personRows.length });
    }
  });
  aggregated.sort((a, b) => {
    if (sc === 'name') return sd === 'asc' ? a.name.localeCompare(b.name) : b.name.localeCompare(a.name);
    return sd === 'asc' ? a.hours - b.hours : b.hours - a.hours;
  });
  return aggregated;
}

function wwOpenDetail(name) {
  const rows = wwAllData.filter(r => r.name === name);
  if (rows.length === 0) return;
  const totalHours = rows.reduce((s, r) => s + r.hours, 0);
  let html = `<div class="panel-section">
    <h4>Summary</h4>
    <div class="panel-row"><span class="k">Activities</span><span class="v">${rows.length}</span></div>
    <div class="panel-row"><span class="k">Total hours</span><span class="v big">${totalHours.toFixed(1)}h</span></div>
  </div>
  <div class="panel-section"><h4>Activity detail</h4>`;
  rows.forEach(r => {
    html += `<div style="padding:8px 0;border-bottom:1px solid var(--border)">
      <div style="font-weight:600;font-size:0.85rem;margin-bottom:4px">${r.activity || '—'}</div>
      <div class="panel-row"><span class="k">Hours</span><span class="v big" style="color:var(--teal)">${r.hours.toFixed(1)}h</span></div>
      ${r.notes ? `<div class="panel-row"><span class="k">Notes</span><span class="v" style="font-size:0.78rem;color:var(--muted)">${r.notes}</span></div>` : ''}
    </div>`;
  });
  html += '</div>';
  openPanel(name, `${totalHours.toFixed(1)}h total · ${rows.length} activity(s)`, html);
}

function renderWwTable() {
  const data = wwGetSorted();
  const maxHours = Math.max(...data.map(r => r.hours), 1);
  document.getElementById('wwTbody').innerHTML = data.map(r => {
    const barPct = Math.min(r.hours / maxHours * 100, 100).toFixed(1);
    return `<tr>
      <td class="name-f" data-name="${encodeURIComponent(r.name)}" style="cursor:pointer">${r.name}</td>
      <td class="num">${r.activities}</td>
      <td class="num">${r.hours.toFixed(1)}</td>
      <td><div class="pgr-hrs-bar-wrap"><div class="pgr-hrs-bar"><div class="pgr-hrs-bar-fill" style="width:${barPct}%" title="${r.hours.toFixed(1)} hours"></div></div></td>
    </tr>`;
  }).join('');
  document.querySelectorAll('#wwTbody .name-f').forEach(cell => {
    cell.addEventListener('click', () => wwOpenDetail(decodeURIComponent(cell.dataset.name)));
  });
  const totalHours = data.reduce((s, r) => s + r.hours, 0);
  const totalActivities = data.reduce((s, r) => s + r.activities, 0);
  document.getElementById('wwFoot').innerHTML = `<tr>
    <td><strong>Total (${data.length} staff)</strong></td>
    <td class="num"><strong>${totalActivities}</strong></td>
    <td class="num"><strong>${totalHours.toFixed(1)}</strong></td>
    <td></td>
  </tr>`;
}

function renderWw(data) {
  wwAllData = data;
  const totalHours = data.reduce((s, r) => s + r.hours, 0);
  const totalStaff = new Set(data.map(r => r.name)).size;
  document.getElementById('wwStatsBar').innerHTML = `
    <div class="stat-card" style="border-left:3px solid var(--teal)"><div class="sc-v">${totalStaff}</div><div class="sc-l">Staff with Welcome Week activities</div></div>
    <div class="stat-card" style="border-left:3px solid var(--gold)"><div class="sc-v">${totalHours.toFixed(0)}</div><div class="sc-l">Total Welcome Week Hours</div></div>
    <div class="stat-card" style="border-left:3px solid var(--mid-blue)"><div class="sc-v">${data.length}</div><div class="sc-l">Activity entries</div></div>
  `;
  renderWwTable();
  document.getElementById('ww-landing').style.display = 'none';
  document.getElementById('ww-content').style.display = 'block';
  document.getElementById('badge-welcomeweek').textContent = totalStaff;
  document.getElementById('wwMeta').textContent = `${totalStaff} staff · ${totalHours.toFixed(1)} hours · ${data.length} entries`;
  updateCombStatus();
}

document.addEventListener('DOMContentLoaded', () => {
  const inp = document.getElementById('wwFileInput');
  if (inp) inp.addEventListener('change', e => { if (e.target.files[0]) parseWwXlsx(e.target.files[0]); });

  const dz = document.getElementById('wwDropZone');
  if (dz) {
    dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('drag-over'); });
    dz.addEventListener('dragleave', () => dz.classList.remove('drag-over'));
    dz.addEventListener('drop', e => { e.preventDefault(); dz.classList.remove('drag-over'); const f = e.dataTransfer.files[0]; if (f) parseWwXlsx(f); });
  }

  const back = document.getElementById('wwBtnBack');
  if (back) back.addEventListener('click', () => {
    document.getElementById('ww-content').style.display = 'none';
    document.getElementById('ww-landing').style.display = 'block';
    wwAllData = [];
    document.getElementById('badge-welcomeweek').textContent = '—';
    updateCombStatus();
  });

  const sortSel = document.getElementById('wwSortSel');
  if (sortSel) sortSel.addEventListener('change', renderWwTable);
  const search = document.getElementById('wwSearch');
  if (search) search.addEventListener('input', renderWwTable);

  document.addEventListener('click', e => {
    const th = e.target.closest('th[data-wwsort]');
    if (!th) return;
    const col = th.dataset.wwsort;
    const sel = document.getElementById('wwSortSel');
    if (!sel) return;
    const isNum = ['hours', 'activities'].includes(col);
    const dir = (sel.value.startsWith(col) && sel.value.endsWith('desc')) ? 'asc' : (isNum ? 'desc' : 'asc');
    const opt = [...sel.options].find(o => o.value === `${col}-${dir}`);
    if (opt) { sel.value = opt.value; renderWwTable(); }
  });

  const expBtn = document.getElementById('wwBtnExport');
  if (expBtn) expBtn.addEventListener('click', () => {
    if (wwAllData.length === 0) return;
    const hdr = ['Staff member', 'Activity', 'Hours', 'Notes'];
    const rows = [hdr];
    wwAllData.forEach(r => rows.push([r.name, r.activity, r.hours, r.notes]));
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Welcome Week');
    XLSX.writeFile(wb, 'welcome_week.xlsx');
  });
});
