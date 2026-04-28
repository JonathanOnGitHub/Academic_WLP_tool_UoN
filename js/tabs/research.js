// ═══════════════════════════════════════════════════════
// RESEARCH TAB
// ═══════════════════════════════════════════════════════
let resAllData = [];
let resSortCol = 'currhours', resSortDir = 'desc';
let resYearPref = 'current';

window.getResHoursTotals = function() {
  const totals = {};
  resAllData.forEach(r => {
    const hrs = resYearPref === 'next' ? r.nextHours : r.currHours;
    if (r.name && hrs > 0) totals[r.name] = (totals[r.name] || 0) + hrs;
  });
  return totals;
};

function parseResearchXlsx(file) {
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      if (rows.length < 2) throw new Error('No data rows found.');

      let hdrIdx = 0;
      for (let i = 0; i < Math.min(5, rows.length); i++) {
        const lower = rows[i].map(c => String(c).toLowerCase());
        if (lower.some(c => c.includes('name'))) { hdrIdx = i; break; }
      }
      const hdrs = rows[hdrIdx].map(c => String(c).toLowerCase().trim());

      function col(keywords) {
        const idx = hdrs.findIndex(h => keywords.some(k => h.includes(k)));
        return idx >= 0 ? idx : -1;
      }
      const cId   = col(['staff identifier','identifier','staff id','id']);
      const cName = col(['name']);
      const cDept = col(['department','dept']);
      const cFte  = col(['max fte','fte']);
      const cProj = col(['project count','projects']);
      const cHpw  = col(['hrs per week','hrs/week','hours per week']);
      const cCurr = col(['current year','curr']);
      const cNext = col(['next year','next']);

      if (cName < 0) throw new Error('Could not find a "Name" column.');

      const data = [];
      for (let i = hdrIdx + 1; i < rows.length; i++) {
        const r = rows[i];
        const name = String(r[cName] || '').trim();
        if (!name) continue;
        data.push({
          identifier: cId >= 0  ? String(r[cId]  || '').trim() : '',
          name,
          dept:      cDept >= 0 ? String(r[cDept] || '').trim() : '',
          fte:       cFte  >= 0 ? (parseFloat(r[cFte])  || 0) : 0,
          projects:  cProj >= 0 ? (parseInt(r[cProj])   || 0) : 0,
          hrsWeek:   cHpw  >= 0 ? (parseFloat(r[cHpw])  || 0) : 0,
          currHours: cCurr >= 0 ? (parseFloat(r[cCurr]) || 0) : 0,
          nextHours: cNext >= 0 ? (parseFloat(r[cNext]) || 0) : 0,
        });
      }
      if (data.length === 0) throw new Error('No data rows could be parsed.');
      renderResearch(data);
      document.getElementById('resError').style.display = 'none';
    } catch(err) {
      const el = document.getElementById('resError');
      el.textContent = 'Error: ' + err.message;
      el.style.display = 'block';
    }
  };
  reader.readAsArrayBuffer(file);
}

function resGetSorted() {
  const q = document.getElementById('resSearch')?.value.toLowerCase() || '';
  let data = resAllData.filter(r =>
    r.name.toLowerCase().includes(q) || r.dept.toLowerCase().includes(q) || r.identifier.toLowerCase().includes(q)
  );
  const sel = document.getElementById('resSortSel')?.value || 'currhours-desc';
  const [sc, sd] = sel.split('-');
  data.sort((a, b) => {
    let av, bv;
    if (sc === 'name') { av = a.name; bv = b.name; return sd === 'asc' ? av.localeCompare(bv) : bv.localeCompare(av); }
    if (sc === 'dept') { av = a.dept; bv = b.dept; return sd === 'asc' ? av.localeCompare(bv) : bv.localeCompare(av); }
    if (sc === 'identifier') { av = a.identifier; bv = b.identifier; return sd === 'asc' ? av.localeCompare(bv) : bv.localeCompare(av); }
    av = sc === 'fte' ? a.fte : sc === 'projects' ? a.projects : sc === 'hrsweek' ? a.hrsWeek : sc === 'nexthours' ? a.nextHours : a.currHours;
    bv = sc === 'fte' ? b.fte : sc === 'projects' ? b.projects : sc === 'hrsweek' ? b.hrsWeek : sc === 'nexthours' ? b.nextHours : b.currHours;
    return sd === 'asc' ? av - bv : bv - av;
  });
  return data;
}

function renderResTable() {
  const data = resGetSorted();
  const maxCurr = Math.max(...resAllData.map(r => r.currHours), 1);
  document.getElementById('resTbody').innerHTML = data.map(r => {
    const barPct = Math.min(r.currHours / maxCurr * 100, 100).toFixed(1);
    return `<tr>
      <td style="font-family:'IBM Plex Mono',monospace;font-size:0.78rem;color:var(--muted)">${r.identifier || '—'}</td>
      <td class="name-f">${r.name}</td>
      <td>${r.dept || '—'}</td>
      <td class="num">${r.fte > 0 ? r.fte.toFixed(2) : '—'}</td>
      <td class="num">${r.projects > 0 ? r.projects : '—'}</td>
      <td class="num">${r.hrsWeek > 0 ? r.hrsWeek.toFixed(1) : '—'}</td>
      <td class="num">${r.currHours > 0 ? r.currHours.toFixed(1) : '—'}</td>
      <td class="num">${r.nextHours > 0 ? r.nextHours.toFixed(1) : '—'}</td>
      <td><div class="res-hrs-bar-wrap"><div class="res-hrs-bar"><div class="res-hrs-bar-fill" style="width:${barPct}%"></div></div><span class="hours-val">${r.currHours > 0 ? r.currHours.toFixed(0) : '—'}</span></div></td>
    </tr>`;
  }).join('');
  const totCurr = data.reduce((s, r) => s + r.currHours, 0);
  const totNext = data.reduce((s, r) => s + r.nextHours, 0);
  const totProj = data.reduce((s, r) => s + r.projects, 0);
  document.getElementById('resFoot').innerHTML = `<tr>
    <td colspan="3"><strong>Total (${data.length} staff)</strong></td>
    <td class="num">—</td>
    <td class="num"><strong>${totProj}</strong></td>
    <td class="num">—</td>
    <td class="num"><strong>${totCurr.toFixed(1)}</strong></td>
    <td class="num"><strong>${totNext.toFixed(1)}</strong></td>
    <td></td>
  </tr>`;
}

function renderResearch(data) {
  resAllData = data;
  const totCurr  = data.reduce((s, r) => s + r.currHours, 0);
  const totNext  = data.reduce((s, r) => s + r.nextHours, 0);
  const avgFte   = data.filter(r => r.fte > 0).reduce((s, r) => s + r.fte, 0) / (data.filter(r => r.fte > 0).length || 1);
  const totProj  = data.reduce((s, r) => s + r.projects, 0);
  document.getElementById('resStatsBar').innerHTML = `
    <div class="stat-card teal"><div class="sc-v">${data.length}</div><div class="sc-l">Staff members</div></div>
    <div class="stat-card gold"><div class="sc-v">${totCurr.toFixed(0)}</div><div class="sc-l">Current Year Total Hrs</div></div>
    <div class="stat-card rust"><div class="sc-v">${totNext.toFixed(0)}</div><div class="sc-l">Next Year Total Hrs</div></div>
    <div class="stat-card purple"><div class="sc-v">${totProj}</div><div class="sc-l">Total Projects</div></div>
    <div class="stat-card"><div class="sc-v">${avgFte.toFixed(2)}</div><div class="sc-l">Avg Max FTE</div></div>
  `;
  renderResTable();
  document.getElementById('res-landing').style.display = 'none';
  document.getElementById('res-content').style.display = 'block';
  document.getElementById('badge-research').textContent = data.length;
  document.getElementById('resMeta').textContent = `${data.length} staff · uploaded`;
  updateCombStatus();
}

document.addEventListener('DOMContentLoaded', () => {
  const inp = document.getElementById('resFileInput');
  if (inp) inp.addEventListener('change', e => { if (e.target.files[0]) parseResearchXlsx(e.target.files[0]); });

  const dz = document.getElementById('resDropZone');
  if (dz) {
    dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('drag-over'); });
    dz.addEventListener('dragleave', () => dz.classList.remove('drag-over'));
    dz.addEventListener('drop', e => { e.preventDefault(); dz.classList.remove('drag-over'); const f = e.dataTransfer.files[0]; if (f) parseResearchXlsx(f); });
  }

  const back = document.getElementById('resBtnBack');
  if (back) back.addEventListener('click', () => {
    document.getElementById('res-content').style.display = 'none';
    document.getElementById('res-landing').style.display = 'block';
    resAllData = [];
    document.getElementById('badge-research').textContent = '—';
    updateCombStatus();
  });

  const sortSel = document.getElementById('resSortSel');
  if (sortSel) sortSel.addEventListener('change', renderResTable);
  const search = document.getElementById('resSearch');
  if (search) search.addEventListener('input', renderResTable);

  document.querySelectorAll('input[name="resYearPref"]').forEach(radio => {
    radio.addEventListener('change', e => {
      resYearPref = e.target.value;
      updateCombStatus();
    });
  });

  document.addEventListener('click', e => {
    const th = e.target.closest('th[data-ressort]');
    if (!th) return;
    const col = th.dataset.ressort;
    const sel = document.getElementById('resSortSel');
    if (!sel) return;
    const isNum = ['fte','projects','hrsweek','currhours','nexthours'].includes(col);
    const dir = (sel.value.startsWith(col) && sel.value.endsWith('desc')) ? 'asc' : (isNum ? 'desc' : 'asc');
    const opt = [...sel.options].find(o => o.value === `${col}-${dir}`);
    if (opt) { sel.value = opt.value; renderResTable(); }
  });

  const expBtn = document.getElementById('resBtnExport');
  if (expBtn) expBtn.addEventListener('click', () => {
    if (resAllData.length === 0) return;
    const rows = [['Staff Identifier','Name','Department','Max FTE','Project Count','Hrs Per Week','Current Year Hours','Next Year Hours']];
    resGetSorted().forEach(r => rows.push([r.identifier, r.name, r.dept, r.fte, r.projects, r.hrsWeek, r.currHours, r.nextHours]));
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Research Hours');
    XLSX.writeFile(wb, 'staff_research_hours.xlsx');
  });
});
