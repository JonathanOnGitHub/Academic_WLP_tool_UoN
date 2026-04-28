// ═══════════════════════════════════════════════════════
// ASSESSMENT TAB
// ═══════════════════════════════════════════════════════
let assessmentAllData = [];

window.getAssessmentHoursTotals = function() {
  const totals = {};
  assessmentAllData.forEach(r => {
    if (r.supervisor && r.hours > 0) totals[r.supervisor] = (totals[r.supervisor] || 0) + r.hours;
  });
  return totals;
};

function parseAssessmentXlsx(file) {
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      if (rows.length < 3) throw new Error('Need at least header rows and one data row.');

      const hdrs = rows[1].map(c => String(c).toLowerCase().trim());
      const cDesc = 0;
      const cYear = 1;
      const cCourse = 2;
      const cLoadPerStudent = 3;
      const cTotalStudents = 4;
      const cTotalLoad = 5;
      const cAllStaff = 6;
      const cStaff1 = 7;
      const cExpiry1 = 8;
      const cStaff2 = 9;
      const cExpiry2 = 10;

      const data = [];
      const today = new Date();
      for (let i = 2; i < rows.length; i++) {
        const r = rows[i];
        const totalLoad = parseFloat(r[cTotalLoad]);
        if (isNaN(totalLoad) || totalLoad <= 0) continue;

        const allStaffRaw = String(r[cAllStaff] || '').trim();
        if (!allStaffRaw) continue;
        const allStaff = allStaffRaw.split(',').map(s => s.trim()).filter(s => s);
        if (allStaff.length === 0) continue;

        const baseHoursPerStaff = totalLoad / allStaff.length;

        const multipliers = new Map();
        allStaff.forEach(name => multipliers.set(name, 1.0));

        const staff1 = String(r[cStaff1] || '').trim();
        const expiry1 = parseDate(r[cExpiry1]);
        if (staff1 && expiry1 && today < expiry1) {
          if (multipliers.has(staff1)) {
            multipliers.set(staff1, multipliers.get(staff1) + 1.0);
          }
        }

        const staff2 = String(r[cStaff2] || '').trim();
        const expiry2 = parseDate(r[cExpiry2]);
        if (staff2 && expiry2 && today < expiry2) {
          if (multipliers.has(staff2)) {
            multipliers.set(staff2, multipliers.get(staff2) + 0.5);
          }
        }

        multipliers.forEach((multiplier, name) => {
          const hours = baseHoursPerStaff * multiplier;
          data.push({
            assessmentDesc: r[cDesc],
            year: r[cYear],
            course: r[cCourse],
            loadPerStudent: r[cLoadPerStudent],
            totalStudents: r[cTotalStudents],
            totalLoad,
            supervisor: name,
            hours,
            multiplier,
            staff1,
            expiry1: expiry1 ? expiry1.toISOString().split('T')[0] : '',
            staff2,
            expiry2: expiry2 ? expiry2.toISOString().split('T')[0] : ''
          });
        });
      }
      if (data.length === 0) throw new Error('No valid assessment rows found.');
      renderAssessment(data);
      document.getElementById('assessmentError').style.display = 'none';
    } catch(err) {
      const el = document.getElementById('assessmentError');
      el.textContent = 'Error: ' + err.message;
      el.style.display = 'block';
    }
  };
  reader.readAsArrayBuffer(file);
}

function assessmentGetSorted() {
  const q = document.getElementById('assessmentSearch')?.value.toLowerCase() || '';
  let data = assessmentAllData.filter(r =>
    r.supervisor.toLowerCase().includes(q)
  );
  const sel = document.getElementById('assessmentSortSel')?.value || 'hours-desc';
  const [sc, sd] = sel.split('-');
  data.sort((a, b) => {
    let av, bv;
    if (sc === 'name') { av = a.supervisor; bv = b.supervisor; return sd === 'asc' ? av.localeCompare(bv) : bv.localeCompare(av); }
    av = a.hours; bv = b.hours;
    return sd === 'asc' ? av - bv : bv - av;
  });
  const aggregated = [];
  const seen = new Set();
  data.forEach(r => {
    if (!seen.has(r.supervisor)) {
      seen.add(r.supervisor);
      const supervisorRows = assessmentAllData.filter(d => d.supervisor === r.supervisor);
      const totalHours = supervisorRows.reduce((sum, d) => sum + d.hours, 0);
      aggregated.push({ supervisor: r.supervisor, hours: totalHours });
    }
  });
  aggregated.sort((a, b) => {
    let av, bv;
    if (sc === 'name') { av = a.supervisor; bv = b.supervisor; return sd === 'asc' ? av.localeCompare(bv) : bv.localeCompare(av); }
    av = a.hours; bv = b.hours;
    return sd === 'asc' ? av - bv : bv - av;
  });
  return aggregated;
}

function renderAssessmentTable() {
  const data = assessmentGetSorted();
  const maxHours = Math.max(...data.map(r => r.hours), 1);
  document.getElementById('assessmentTbody').innerHTML = data.map(r => {
    const barPct = Math.min(r.hours / maxHours * 100, 100).toFixed(1);
    return `<tr>
      <td class="name-f">${r.supervisor}</td>
      <td class="num">${r.hours.toFixed(1)}</td>
      <td><div class="pgr-hrs-bar-wrap"><div class="pgr-hrs-bar"><div class="pgr-hrs-bar-fill" style="width:${barPct}%" title="${r.hours.toFixed(1)} hours"></div></div></td>
    </tr>`;
  }).join('');
  const totalHours = data.reduce((s, r) => s + r.hours, 0);
  document.getElementById('assessmentFoot').innerHTML = `<tr>
    <td><strong>Total (${data.length} staff)</strong></td>
    <td class="num"><strong>${totalHours.toFixed(1)}</strong></td>
    <td></td>
  </tr>`;
}

function renderAssessment(data) {
  assessmentAllData = data;
  const totalHours = data.reduce((s, r) => s + r.hours, 0);
  const totalStaff = new Set(data.map(r => r.supervisor)).size;
  document.getElementById('assessmentStatsBar').innerHTML = `
    <div class="stat-card teal"><div class="sc-v">${totalStaff}</div><div class="sc-l">Staff with Non-timetabled Assessment Load</div></div>
    <div class="stat-card gold"><div class="sc-v">${totalHours.toFixed(0)}</div><div class="sc-l">Total Non-timetabled Assessment Hours</div></div>
  `;
  renderAssessmentTable();
  document.getElementById('assessment-landing').style.display = 'none';
  document.getElementById('assessment-content').style.display = 'block';
  document.getElementById('badge-assessment').textContent = totalStaff;
  document.getElementById('assessmentMeta').textContent = `${totalStaff} staff · ${totalHours.toFixed(1)} hours`;
  updateCombStatus();
}

document.addEventListener('DOMContentLoaded', () => {
  const inp = document.getElementById('assessmentFileInput');
  if (inp) inp.addEventListener('change', e => { if (e.target.files[0]) parseAssessmentXlsx(e.target.files[0]); });

  const dz = document.getElementById('assessmentDropZone');
  if (dz) {
    dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('drag-over'); });
    dz.addEventListener('dragleave', () => dz.classList.remove('drag-over'));
    dz.addEventListener('drop', e => { e.preventDefault(); dz.classList.remove('drag-over'); const f = e.dataTransfer.files[0]; if (f) parseAssessmentXlsx(f); });
  }

  const back = document.getElementById('assessmentBtnBack');
  if (back) back.addEventListener('click', () => {
    document.getElementById('assessment-content').style.display = 'none';
    document.getElementById('assessment-landing').style.display = 'block';
    assessmentAllData = [];
    document.getElementById('badge-assessment').textContent = '—';
    updateCombStatus();
  });

  const sortSel = document.getElementById('assessmentSortSel');
  if (sortSel) sortSel.addEventListener('change', renderAssessmentTable);
  const search = document.getElementById('assessmentSearch');
  if (search) search.addEventListener('input', renderAssessmentTable);

  document.addEventListener('click', e => {
    const th = e.target.closest('th[data-assessmentsort]');
    if (!th) return;
    const col = th.dataset.assessmentsort;
    const sel = document.getElementById('assessmentSortSel');
    if (!sel) return;
    const isNum = ['hours'].includes(col);
    const dir = (sel.value.startsWith(col) && sel.value.endsWith('desc')) ? 'asc' : (isNum ? 'desc' : 'asc');
    const opt = [...sel.options].find(o => o.value === `${col}-${dir}`);
    if (opt) { sel.value = opt.value; renderAssessmentTable(); }
  });

  const expBtn = document.getElementById('assessmentBtnExport');
  if (expBtn) expBtn.addEventListener('click', () => {
    if (assessmentAllData.length === 0) return;
    const rows = [['Non-timetabled Assessment Description', 'Year', 'Course', 'Load per student (min)', 'Total Students', 'Total Load (hours)', 'Staff', 'Hours', 'Multiplier', 'Staff 1.0x', 'Expiry 1', 'Staff 0.5x', 'Expiry 2']];
    assessmentAllData.forEach(r => rows.push([r.assessmentDesc, r.year, r.course, r.loadPerStudent, r.totalStudents, r.totalLoad, r.supervisor, r.hours, r.multiplier, r.staff1, r.expiry1, r.staff2, r.expiry2]));
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Non-timetabled Assessment Load');
    XLSX.writeFile(wb, 'non_timetabled_assessment_load.xlsx');
  });
});
