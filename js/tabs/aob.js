// ═══════════════════════════════════════════════════════
// AOB (Any other Business) TAB
// ═══════════════════════════════════════════════════════
let aobAllData = [];

window.getAobHoursTotals = function() {
  const totals = {};
  aobAllData.forEach(r => {
    if (r.name && r.hours > 0) totals[r.name] = (totals[r.name] || 0) + r.hours;
  });
  return totals;
};

function parseAobXlsx(file) {
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
      let parsedData = null;

      // Try each sheet in order; use the first one with matching headers and data
      for (const sheetName of wb.SheetNames) {
        if (parsedData) break;
        const ws = wb.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });
        if (rows.length < 2) continue;

        let hdrIdx = -1;
        for (let i = 0; i < Math.min(10, rows.length); i++) {
          const lower = rows[i].map(c => String(c).toLowerCase().trim());
          if (lower.some(c => c.includes('activity description'))) { hdrIdx = i; break; }
        }
        if (hdrIdx < 0) continue;

        const hdrs = rows[hdrIdx].map(c => String(c).toLowerCase().trim());

        function col(keywords) {
          const idx = hdrs.findIndex(h => keywords.some(k => h.includes(k)));
          return idx >= 0 ? idx : -1;
        }
        const cActDesc = col(['activity description', 'activity', 'description']);
        const cStaff   = col(['staff member(s)', 'staff members', 'staff', 'staff member']);
        const cHours   = col(['hours per activity', 'hours', 'hrs', 'total time per staff member']);
        const cStart   = col(['start date']);
        const cEnd     = col(['end date']);
        const cNotes   = col(['notes']);

        if (cActDesc < 0 || cStaff < 0 || cHours < 0) continue;

        const data = [];
        for (let i = hdrIdx + 1; i < rows.length; i++) {
          const r = rows[i];
          const activityDesc = String(r[cActDesc] || '').trim();
          const rawStaff = String(r[cStaff] || '').trim();
          const rawHours = parseFloat(r[cHours]);
          if (!activityDesc || !rawStaff || isNaN(rawHours) || rawHours <= 0) continue;

          const staffNames = rawStaff.split(/[,;]/).map(s => s.trim()).filter(s => s.length > 0);
          if (staffNames.length === 0) continue;

          const hrsPerStaff = rawHours / staffNames.length;
          const startDate = cStart >= 0 ? String(r[cStart] || '').trim() : '';
          const endDate = cEnd >= 0 ? String(r[cEnd] || '').trim() : '';
          const notes = cNotes >= 0 ? String(r[cNotes] || '').trim() : '';

          staffNames.forEach(name => {
            data.push({
              activityDesc,
              staffMembers: rawStaff,
              hours: hrsPerStaff,
              startDate,
              endDate,
              notes,
              name
            });
          });
        }
        if (data.length > 0) parsedData = data;
      }

      if (!parsedData) throw new Error('No valid data rows found. Ensure the file has a sheet with columns: Activity description, Staff member(s), Hours per activity.');
      renderAob(parsedData);
      document.getElementById('aobError').style.display = 'none';
    } catch(err) {
      const el = document.getElementById('aobError');
      el.textContent = 'Error: ' + err.message;
      el.style.display = 'block';
    }
  };
  reader.readAsArrayBuffer(file);
}

function aobGetSorted() {
  const q = document.getElementById('aobSearch')?.value.toLowerCase() || '';
  let data = aobAllData.filter(r =>
    r.name.toLowerCase().includes(q) || r.activityDesc.toLowerCase().includes(q)
  );
  const sel = document.getElementById('aobSortSel')?.value || 'hours-desc';
  const [sc, sd] = sel.split('-');
  data.sort((a, b) => {
    if (sc === 'name') { return sd === 'asc' ? a.name.localeCompare(b.name) : b.name.localeCompare(a.name); }
    if (sc === 'activity') { return sd === 'asc' ? a.activityDesc.localeCompare(b.activityDesc) : b.activityDesc.localeCompare(a.activityDesc); }
    const av = a.hours, bv = b.hours;
    return sd === 'asc' ? av - bv : bv - av;
  });
  // Aggregate by staff member for the table view
  const aggregated = [];
  const seen = new Set();
  data.forEach(r => {
    if (!seen.has(r.name)) {
      seen.add(r.name);
      const personRows = aobAllData.filter(d => d.name === r.name);
      const totalHours = personRows.reduce((sum, d) => sum + d.hours, 0);
      aggregated.push({ name: r.name, hours: totalHours, activities: personRows.length });
    }
  });
  aggregated.sort((a, b) => {
    if (sc === 'name') { return sd === 'asc' ? a.name.localeCompare(b.name) : b.name.localeCompare(a.name); }
    return sd === 'asc' ? a.hours - b.hours : b.hours - a.hours;
  });
  return aggregated;
}

function aobOpenDetail(name) {
  const rows = aobAllData.filter(r => r.name === name);
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
      <div style="font-weight:600;font-size:0.85rem;margin-bottom:4px">${r.activityDesc || '—'}</div>
      <div class="panel-row"><span class="k">Hours</span><span class="v big" style="color:var(--mid-blue)">${r.hours.toFixed(1)}h</span></div>
      <div class="panel-row"><span class="k">Staff</span><span class="v">${r.staffMembers}</span></div>
      ${r.startDate ? `<div class="panel-row"><span class="k">Start date</span><span class="v">${r.startDate}</span></div>` : ''}
      ${r.endDate ? `<div class="panel-row"><span class="k">End date</span><span class="v">${r.endDate}</span></div>` : ''}
      ${r.notes ? `<div class="panel-row"><span class="k">Notes</span><span class="v" style="font-size:0.78rem;color:var(--muted)">${r.notes}</span></div>` : ''}
    </div>`;
  });
  html += '</div>';
  openPanel(name, `${totalHours.toFixed(1)}h total · ${rows.length} activity(s)`, html);
}

function renderAobTable() {
  const data = aobGetSorted();
  const maxHours = Math.max(...data.map(r => r.hours), 1);
  document.getElementById('aobTbody').innerHTML = data.map(r => {
    const barPct = Math.min(r.hours / maxHours * 100, 100).toFixed(1);
    return `<tr>
      <td class="name-f" data-name="${encodeURIComponent(r.name)}" style="cursor:pointer">${r.name}</td>
      <td class="num">${r.activities}</td>
      <td class="num">${r.hours.toFixed(1)}</td>
      <td><div class="pgr-hrs-bar-wrap"><div class="pgr-hrs-bar"><div class="pgr-hrs-bar-fill" style="width:${barPct}%" title="${r.hours.toFixed(1)} hours"></div></div></td>
    </tr>`;
  }).join('');
  document.querySelectorAll('#aobTbody .name-f').forEach(cell => {
    cell.addEventListener('click', () => aobOpenDetail(decodeURIComponent(cell.dataset.name)));
  });
  const totalHours = data.reduce((s, r) => s + r.hours, 0);
  const totalActivities = data.reduce((s, r) => s + r.activities, 0);
  document.getElementById('aobFoot').innerHTML = `<tr>
    <td><strong>Total (${data.length} staff)</strong></td>
    <td class="num"><strong>${totalActivities}</strong></td>
    <td class="num"><strong>${totalHours.toFixed(1)}</strong></td>
    <td></td>
  </tr>`;
}

function renderAob(data) {
  aobAllData = data;
  const totalHours = data.reduce((s, r) => s + r.hours, 0);
  const totalStaff = new Set(data.map(r => r.name)).size;
  document.getElementById('aobStatsBar').innerHTML = `
    <div class="stat-card teal"><div class="sc-v">${totalStaff}</div><div class="sc-l">Staff with AoB activities</div></div>
    <div class="stat-card gold"><div class="sc-v">${totalHours.toFixed(0)}</div><div class="sc-l">Total AoB Hours</div></div>
    <div class="stat-card"><div class="sc-v">${data.length}</div><div class="sc-l">Activity entries</div></div>
  `;
  renderAobTable();
  document.getElementById('aob-landing').style.display = 'none';
  document.getElementById('aob-content').style.display = 'block';
  document.getElementById('badge-aob').textContent = totalStaff;
  document.getElementById('aobMeta').textContent = `${totalStaff} staff · ${totalHours.toFixed(1)} hours · ${data.length} entries`;
  updateCombStatus();
}

document.addEventListener('DOMContentLoaded', () => {
  const inp = document.getElementById('aobFileInput');
  if (inp) inp.addEventListener('change', e => { if (e.target.files[0]) parseAobXlsx(e.target.files[0]); });

  const dz = document.getElementById('aobDropZone');
  if (dz) {
    dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('drag-over'); });
    dz.addEventListener('dragleave', () => dz.classList.remove('drag-over'));
    dz.addEventListener('drop', e => { e.preventDefault(); dz.classList.remove('drag-over'); const f = e.dataTransfer.files[0]; if (f) parseAobXlsx(f); });
  }

  const back = document.getElementById('aobBtnBack');
  if (back) back.addEventListener('click', () => {
    document.getElementById('aob-content').style.display = 'none';
    document.getElementById('aob-landing').style.display = 'block';
    aobAllData = [];
    document.getElementById('badge-aob').textContent = '—';
    updateCombStatus();
  });

  const sortSel = document.getElementById('aobSortSel');
  if (sortSel) sortSel.addEventListener('change', renderAobTable);
  const search = document.getElementById('aobSearch');
  if (search) search.addEventListener('input', renderAobTable);

  document.addEventListener('click', e => {
    const th = e.target.closest('th[data-aobsort]');
    if (!th) return;
    const col = th.dataset.aobsort;
    const sel = document.getElementById('aobSortSel');
    if (!sel) return;
    const isNum = ['hours', 'activities'].includes(col);
    const dir = (sel.value.startsWith(col) && sel.value.endsWith('desc')) ? 'asc' : (isNum ? 'desc' : 'asc');
    const opt = [...sel.options].find(o => o.value === `${col}-${dir}`);
    if (opt) { sel.value = opt.value; renderAobTable(); }
  });

  const expBtn = document.getElementById('aobBtnExport');
  if (expBtn) expBtn.addEventListener('click', () => {
    if (aobAllData.length === 0) return;
    const hdr = ['Activity description', 'Staff member(s)', 'Hours per activity', 'Start date', 'End date', 'Notes'];
    const rows = [hdr];
    aobAllData.forEach(r => rows.push([r.activityDesc, r.staffMembers, r.hours, r.startDate, r.endDate, r.notes]));
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'AoB Activities');
    XLSX.writeFile(wb, 'aob_activities.xlsx');
  });
});
