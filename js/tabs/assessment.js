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
      if (rows.length < 2) throw new Error('Need at least header row and one data row.');

      // Single header row at index 0. Expected columns:
      //   0  Assessment description
      //   1  Year of study
      //   2  Course code(s)
      //   3  Load per student (minutes)
      //   4  Total students
      //   5  Total assessment load (hours)
      //   6  All staff (comma separated)
      //   7  New to assessment (+ 1.0x)
      //   8  Handing over assessment (+ 0.5x)
      //   9  Moderators (0.1x)
      //  10  Notes

      const data = [];
      for (let i = 1; i < rows.length; i++) {
        const r = rows[i];
        const totalLoad = parseFloat(r[5]);
        if (isNaN(totalLoad) || totalLoad <= 0) continue;

        const splitNames = idx => {
          const raw = String(r[idx] || '').trim();
          return raw ? raw.split(',').map(s => s.trim()).filter(s => s) : [];
        };

        const allStaff = splitNames(6);
        const newToAss = splitNames(7);
        const handingOver = splitNames(8);
        const moderators = splitNames(9);

        // Collect all distinct staff across all four name columns
        const allNames = new Set([...allStaff, ...newToAss, ...handingOver, ...moderators]);
        if (allNames.size === 0) continue;

        const baseHoursPerStaff = totalLoad / allNames.size;

        allNames.forEach(name => {
          let multiplier = 1.0;
          if (newToAss.includes(name)) multiplier += 1.0;
          if (handingOver.includes(name)) multiplier += 0.5;
          if (moderators.includes(name)) multiplier += 0.1;

          const hours = baseHoursPerStaff * multiplier;
          data.push({
            assessmentDesc: r[0],
            year: r[1],
            course: r[2],
            loadPerStudent: r[3],
            totalStudents: r[4],
            totalLoad,
            allStaff,
            newToAssessment: newToAss,
            handingOver,
            moderators,
            notes: String(r[10] || '').trim(),
            supervisor: name,
            hours,
            multiplier
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

function assessmentOpenDetail(name) {
  const rows = assessmentAllData.filter(r => r.supervisor === name);
  if (rows.length === 0) return;
  const totalHours = rows.reduce((s, r) => s + r.hours, 0);
  let html = `<div class="panel-section">
    <h4>Summary</h4>
    <div class="panel-row"><span class="k">Assessments</span><span class="v">${rows.length}</span></div>
    <div class="panel-row"><span class="k">Total hours</span><span class="v big">${totalHours.toFixed(1)}h</span></div>
  </div>
  <div class="panel-section"><h4>Assessment detail</h4>`;
  rows.forEach(r => {
    html += `<div style="padding:8px 0;border-bottom:1px solid var(--border)">
      <div style="font-weight:600;font-size:0.85rem;margin-bottom:4px">${r.assessmentDesc || '—'}</div>
      <div class="panel-row"><span class="k">Year / Course</span><span class="v">${r.year || '—'} · ${r.course || '—'}</span></div>
      <div class="panel-row"><span class="k">Students</span><span class="v">${r.totalStudents || '—'}</span></div>
      <div class="panel-row"><span class="k">Total load</span><span class="v">${r.totalLoad.toFixed(1)}h</span></div>
      <div class="panel-row"><span class="k">Multiplier</span><span class="v">${r.multiplier.toFixed(1)}x</span></div>
      <div class="panel-row"><span class="k">Hours</span><span class="v big" style="color:var(--mid-blue)">${r.hours.toFixed(1)}h</span></div>
      ${r.notes ? `<div class="panel-row"><span class="k">Notes</span><span class="v" style="font-size:0.78rem;color:var(--muted)">${r.notes}</span></div>` : ''}
    </div>`;
  });
  html += '</div>';
  openPanel(name, `${totalHours.toFixed(1)}h total · ${rows.length} assessment(s)`, html);
}

function renderAssessmentTable() {
  const data = assessmentGetSorted();
  const maxHours = Math.max(...data.map(r => r.hours), 1);
  document.getElementById('assessmentTbody').innerHTML = data.map(r => {
    const barPct = Math.min(r.hours / maxHours * 100, 100).toFixed(1);
    return `<tr>
      <td class="name-f" data-name="${encodeURIComponent(r.supervisor)}" style="cursor:pointer">${r.supervisor}</td>
      <td class="num">${r.hours.toFixed(1)}</td>
      <td><div class="pgr-hrs-bar-wrap"><div class="pgr-hrs-bar"><div class="pgr-hrs-bar-fill" style="width:${barPct}%" title="${r.hours.toFixed(1)} hours"></div></div></td>
    </tr>`;
  }).join('');
  document.querySelectorAll('#assessmentTbody .name-f').forEach(cell => {
    cell.addEventListener('click', () => assessmentOpenDetail(decodeURIComponent(cell.dataset.name)));
  });
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
    const hdr = ['Assessment description', 'Year of study', 'Course code(s)', 'Load per student (minutes)', 'Total students', 'Total assessment load (hours)', 'All staff', 'New to assessment (+ 1.0x)', 'Handing over assessment (+ 0.5x)', 'Moderators (0.1x)', 'Notes', 'Staff', 'Hours', 'Multiplier'];
    const rows = [hdr];
    assessmentAllData.forEach(r => rows.push([
      r.assessmentDesc, r.year, r.course, r.loadPerStudent, r.totalStudents, r.totalLoad,
      (r.allStaff||[]).join(', '), (r.newToAssessment||[]).join(', '), (r.handingOver||[]).join(', '), (r.moderators||[]).join(', '),
      r.notes, r.supervisor, r.hours, r.multiplier
    ]));
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Non-timetabled Assessment Load');
    XLSX.writeFile(wb, 'non_timetabled_assessment_load.xlsx');
  });
});
