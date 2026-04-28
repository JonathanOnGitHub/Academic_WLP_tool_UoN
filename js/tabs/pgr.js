// ═══════════════════════════════════════════════════════
// PGR SUPERVISION TAB
// ═══════════════════════════════════════════════════════
let pgrAllData = [];
let pgrSortCol = 'hours', pgrSortDir = 'desc';
let pgrYearPref = 'current';

window.getPgrHoursTotals = function() {
  const totals = {};
  pgrAllData.forEach(r => {
    const hrs = r.hours;
    if (r.supervisor && hrs > 0) totals[r.supervisor] = (totals[r.supervisor] || 0) + hrs;
  });
  return totals;
};

function parsePgrXlsx(file) {
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
        if (lower.some(c => c.includes('first') || c.includes('surname'))) { hdrIdx = i; break; }
      }
      const hdrs = rows[hdrIdx].map(c => String(c).toLowerCase().trim());

      function col(keywords) {
        const idx = hdrs.findIndex(h => keywords.some(k => h.includes(k)));
        return idx >= 0 ? idx : -1;
      }
      const cFirstName = col(['first name', 'first']);
      const cSurname = col(['surname', 'last name', 'last']);
      const cStartDate = col(['start date', 'start']);
      const cEsd = col(['esd']);
      const cEndDate = col(['end date', 'end']);
      const cThesisSubmitted = col(['thesis submitted date', 'thesis submitted', 'submitted']);
      const cPlan = col(['plan/programme', 'plan', 'programme']);
      const cPi = col(['pi', 'principal investigator']);
      const cPercentPi = col(['% pi', 'percent pi', '%pi']);
      const cSupervisor2 = col(['2nd supervisor', 'second supervisor']);
      const cPercent2 = col(['% supervisor 2', '% supervisor2', '% 2nd supervisor']);
      const cSupervisor3 = col(['3rd supervisor', 'third supervisor']);
      const cPercent3 = col(['% supervisor 3', '% supervisor3', '% 3rd supervisor']);
      const cSupervisor4 = col(['4th supervisor', 'fourth supervisor']);
      const cPercent4 = col(['% supervisor 4', '% supervisor4', '% 4th supervisor']);
      const cSupervisor5 = col(['5th supervisor', 'fifth supervisor']);
      const cPercent5 = col(['% supervisor 5', '% supervisor5', '% 5th supervisor']);
      const cAssistant = col(['assistant supervisor', 'assistant']);
      const cMode = col(['mode']);

      if (cFirstName < 0 && cSurname < 0) throw new Error('Could not find a "First Name" or "Surname" column.');

      const data = [];
      const today = new Date();
      for (let i = hdrIdx + 1; i < rows.length; i++) {
        const r = rows[i];
        const firstName = cFirstName >= 0 ? String(r[cFirstName] || '').trim() : '';
        const surname = cSurname >= 0 ? String(r[cSurname] || '').trim() : '';
        const studentName = (firstName + ' ' + surname).trim();
        if (!studentName) continue;
        const startDate = parseDate(r[cStartDate]);
        const endDate = parseDate(r[cEndDate]);
        if (startDate && endDate) {
          if (today < startDate || today > endDate) continue;
        }
        const supervisors = [];
        const addSupervisor = (nameCol, percentCol, defaultPercent) => {
          const name = nameCol >= 0 ? String(r[nameCol] || '').trim() : '';
          if (!name) return;
          let percent = percentCol >= 0 ? parseFloat(r[percentCol]) : defaultPercent;
          if (isNaN(percent) || percent <= 0) percent = defaultPercent;
          supervisors.push({ name, percent });
        };
        addSupervisor(cPi, cPercentPi, 100);
        addSupervisor(cSupervisor2, cPercent2, 0);
        addSupervisor(cSupervisor3, cPercent3, 0);
        addSupervisor(cSupervisor4, cPercent4, 0);
        addSupervisor(cSupervisor5, cPercent5, 0);
        if (cAssistant >= 0) {
          const assistantName = String(r[cAssistant] || '').trim();
          if (assistantName) supervisors.push({ name: assistantName, percent: 0 });
        }

        const hoursPerStudent = 100;
        supervisors.forEach(s => {
          if (s.percent > 0) {
            const hours = hoursPerStudent * (s.percent / 100);
            data.push({
              studentName,
              supervisor: s.name,
              percent: s.percent,
              hours,
              startDate: startDate ? startDate.toISOString().split('T')[0] : '',
              endDate: endDate ? endDate.toISOString().split('T')[0] : '',
              plan: cPlan >= 0 ? String(r[cPlan] || '').trim() : '',
              mode: cMode >= 0 ? String(r[cMode] || '').trim() : '',
            });
          }
        });
      }
      if (data.length === 0) throw new Error('No active PGR students found (current date within start/end dates).');
      renderPgr(data);
      document.getElementById('pgrError').style.display = 'none';
    } catch(err) {
      const el = document.getElementById('pgrError');
      el.textContent = 'Error: ' + err.message;
      el.style.display = 'block';
    }
  };
  reader.readAsArrayBuffer(file);
}

function pgrGetSorted() {
  const q = document.getElementById('pgrSearch')?.value.toLowerCase() || '';
  let data = pgrAllData.filter(r =>
    r.supervisor.toLowerCase().includes(q) || r.studentName.toLowerCase().includes(q)
  );
  const sel = document.getElementById('pgrSortSel')?.value || 'hours-desc';
  const [sc, sd] = sel.split('-');
  data.sort((a, b) => {
    let av, bv;
    if (sc === 'supervisor') { av = a.supervisor; bv = b.supervisor; return sd === 'asc' ? av.localeCompare(bv) : bv.localeCompare(av); }
    if (sc === 'students') {
      const aStudents = new Set(pgrAllData.filter(d => d.supervisor === a.supervisor).map(d => d.studentName));
      const bStudents = new Set(pgrAllData.filter(d => d.supervisor === b.supervisor).map(d => d.studentName));
      av = aStudents.size; bv = bStudents.size;
      return sd === 'asc' ? av - bv : bv - av;
    }
    av = sc === 'hours' ? a.hours : 0;
    bv = sc === 'hours' ? b.hours : 0;
    return sd === 'asc' ? av - bv : bv - av;
  });
  const aggregated = [];
  const seen = new Set();
  data.forEach(r => {
    if (!seen.has(r.supervisor)) {
      seen.add(r.supervisor);
      const supervisorRows = pgrAllData.filter(d => d.supervisor === r.supervisor);
      const totalHours = supervisorRows.reduce((sum, d) => sum + d.hours, 0);
      const studentSet = new Set(supervisorRows.map(d => d.studentName));
      aggregated.push({ supervisor: r.supervisor, hours: totalHours, studentCount: studentSet.size });
    }
  });
  aggregated.sort((a, b) => {
    let av, bv;
    if (sc === 'supervisor') { av = a.supervisor; bv = b.supervisor; return sd === 'asc' ? av.localeCompare(bv) : bv.localeCompare(av); }
    if (sc === 'students') { av = a.studentCount; bv = b.studentCount; return sd === 'asc' ? av - bv : bv - av; }
    av = a.hours; bv = b.hours;
    return sd === 'asc' ? av - bv : bv - av;
  });
  return aggregated;
}

function renderPgrTable() {
  const data = pgrGetSorted();
  const maxHours = Math.max(...data.map(r => r.hours), 1);
  document.getElementById('pgrTbody').innerHTML = data.map(r => {
    const enc = encodeURIComponent(r.supervisor);
    const barPct = Math.min(r.hours / maxHours * 100, 100).toFixed(1);
    return `<tr>
      <td class="name-f" data-supervisor="${enc}" style="cursor:pointer">${r.supervisor}</td>
      <td class="num">${r.studentCount}</td>
      <td class="num">${r.hours.toFixed(1)}</td>
      <td><div class="pgr-hrs-bar-wrap"><div class="pgr-hrs-bar"><div class="pgr-hrs-bar-fill" style="width:${barPct}%" title="${r.hours.toFixed(1)} hours"></div></div></td>
    </tr>`;
  }).join('');
  document.querySelectorAll('#pgrTbody .name-f').forEach(cell => {
    cell.addEventListener('click', () => pgrOpenDetail(decodeURIComponent(cell.dataset.supervisor)));
  });
  const totalHours = data.reduce((s, r) => s + r.hours, 0);
  const totalStudents = new Set(pgrAllData.map(d => d.studentName)).size;
  document.getElementById('pgrFoot').innerHTML = `<tr>
    <td><strong>Total (${data.length} supervisors)</strong></td>
    <td class="num"><strong>${totalStudents}</strong></td>
    <td class="num"><strong>${totalHours.toFixed(1)}</strong></td>
    <td></td>
  </tr>`;
}

function pgrOpenDetail(supervisorName) {
  const rows = pgrAllData.filter(r => r.supervisor === supervisorName);
  if (rows.length === 0) return;
  const totalHours = rows.reduce((s, r) => r.hours + s, 0);
  const studentCount = new Set(rows.map(r => r.studentName)).size;

  let html = `<div class="panel-section"><h4>Hours Breakdown</h4>
    <div class="panel-row"><span class="k">Total hours</span><span class="v big">${totalHours.toFixed(1)}h</span></div>
    <div class="panel-row"><span class="k">Students</span><span class="v">${studentCount}</span></div>
  </div>`;
  html += `<div class="panel-section"><h4>Students (${rows.length} supervisor rows)</h4>`;
  for (const r of rows) {
    const startStr = r.startDate ? r.startDate : '—';
    const endStr = r.endDate ? r.endDate : '—';
    html += `<div class="proj-student">
      <div class="sn">${r.studentName}</div>
      <div class="sr">
        ${r.plan ? '<span style="margin-right:8px">📋 ' + r.plan + '</span>' : ''}
        ${r.mode ? '<span style="margin-right:8px">📐 ' + r.mode + '</span>' : ''}
        <span style="margin-right:8px">📅 ${startStr} → ${endStr}</span>
        <span style="font-weight:600;color:var(--teal)">${r.hours.toFixed(1)}h @ ${r.percent}%</span>
      </div>
    </div>`;
  }
  html += '</div>';
  openPanel(supervisorName, `${totalHours.toFixed(1)}h total · ${studentCount} student${studentCount !== 1 ? 's' : ''}`, html);
}

function renderPgr(data) {
  pgrAllData = data;
  const totalHours = data.reduce((s, r) => s + r.hours, 0);
  const totalStudents = new Set(data.map(r => r.studentName)).size;
  const totalSupervisors = new Set(data.map(r => r.supervisor)).size;
  document.getElementById('pgrStatsBar').innerHTML = `
    <div class="stat-card teal"><div class="sc-v">${totalStudents}</div><div class="sc-l">Active PGR Students</div></div>
    <div class="stat-card gold"><div class="sc-v">${totalHours.toFixed(0)}</div><div class="sc-l">Total Supervision Hours</div></div>
    <div class="stat-card rust"><div class="sc-v">${totalSupervisors}</div><div class="sc-l">Supervisors</div></div>
  `;
  renderPgrTable();
  document.getElementById('pgr-landing').style.display = 'none';
  document.getElementById('pgr-content').style.display = 'block';
  document.getElementById('badge-pgr').textContent = totalSupervisors;
  document.getElementById('pgrMeta').textContent = `${totalStudents} students · ${totalSupervisors} supervisors`;
  updateCombStatus();
}

document.addEventListener('DOMContentLoaded', () => {
  const inp = document.getElementById('pgrFileInput');
  if (inp) inp.addEventListener('change', e => { if (e.target.files[0]) parsePgrXlsx(e.target.files[0]); });

  const dz = document.getElementById('pgrDropZone');
  if (dz) {
    dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('drag-over'); });
    dz.addEventListener('dragleave', () => dz.classList.remove('drag-over'));
    dz.addEventListener('drop', e => { e.preventDefault(); dz.classList.remove('drag-over'); const f = e.dataTransfer.files[0]; if (f) parsePgrXlsx(f); });
  }

  const back = document.getElementById('pgrBtnBack');
  if (back) back.addEventListener('click', () => {
    document.getElementById('pgr-content').style.display = 'none';
    document.getElementById('pgr-landing').style.display = 'block';
    pgrAllData = [];
    document.getElementById('badge-pgr').textContent = '—';
    updateCombStatus();
  });

  const sortSel = document.getElementById('pgrSortSel');
  if (sortSel) sortSel.addEventListener('change', renderPgrTable);
  const search = document.getElementById('pgrSearch');
  if (search) search.addEventListener('input', renderPgrTable);

  document.addEventListener('click', e => {
    const th = e.target.closest('th[data-pgrsort]');
    if (!th) return;
    const col = th.dataset.pgrsort;
    const sel = document.getElementById('pgrSortSel');
    if (!sel) return;
    const isNum = ['students','hours'].includes(col);
    const dir = (sel.value.startsWith(col) && sel.value.endsWith('desc')) ? 'asc' : (isNum ? 'desc' : 'asc');
    const opt = [...sel.options].find(o => o.value === `${col}-${dir}`);
    if (opt) { sel.value = opt.value; renderPgrTable(); }
  });

  const expBtn = document.getElementById('pgrBtnExport');
  if (expBtn) expBtn.addEventListener('click', () => {
    if (pgrAllData.length === 0) return;
    const rows = [['Student Name', 'Supervisor', 'Percent', 'Hours', 'Start Date', 'End Date', 'Plan/Programme', 'Mode']];
    pgrAllData.forEach(r => rows.push([r.studentName, r.supervisor, r.percent, r.hours, r.startDate, r.endDate, r.plan, r.mode]));
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'PGR Supervision');
    XLSX.writeFile(wb, 'pgr_supervision.xlsx');
  });
});
