// ═══════════════════════════════════════════════════════
// PGR SUPERVISION TAB — CSV-based import
// ═══════════════════════════════════════════════════════

let pgrAllData = [];

window.getPgrHoursTotals = function() {
  const totals = {};
  pgrAllData.forEach(r => {
    if (r.supervisor && r.hours > 0) totals[r.supervisor] = (totals[r.supervisor] || 0) + r.hours;
  });
  return totals;
};

function parsePgrCsv(file) {
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(e.target.result, { type: 'string' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      if (rows.length < 2) throw new Error('No data rows found.');

      // Find header row
      let hdrIdx = 0;
      for (let i = 0; i < Math.min(5, rows.length); i++) {
        const lower = rows[i].map(c => String(c).toLowerCase().trim());
        if (lower.some(c => c.includes('supervisor name') || c.includes('supervisor id'))) { hdrIdx = i; break; }
      }
      const hdrs = rows[hdrIdx].map(c => String(c).toLowerCase().trim());

      function col(keywords) {
        const idx = hdrs.findIndex(h => keywords.some(k => h.includes(k)));
        return idx >= 0 ? idx : -1;
      }

      const cSupervisorName = col(['supervisor name']);
      const cSupervisorEmail = col(['supervisor email']);
      const cSupervisorRole = col(['supervisor role']);
      const cStudentId = col(['student id']);
      const cFirstName = col(['firstname', 'first name', 'first']);
      const cLastName = col(['lastname', 'last name', 'surname', 'last']);
      const cSchool = col(['school/department', 'school', 'department']);
      const cPlanTitle = col(['plan title']);
      const cPlanCode = col(['plan code']);
      const cProgramTitle = col(['program title', 'programme title']);
      const cProgramCode = col(['program code', 'programme code']);
      const cStudentStatus = col(['student status']);
      const cResearchGroup = col(['research group']);
      const cPercent = col(['supervision percentage', 'percentage', '%']);

      if (cSupervisorName < 0) throw new Error('Could not find a "Supervisor Name" column.');
      if (cPercent < 0) throw new Error('Could not find a "Supervision Percentage" column.');

      const data = [];
      const schools = new Set();

      for (let i = hdrIdx + 1; i < rows.length; i++) {
        const r = rows[i];
        const supervisor = String(r[cSupervisorName] || '').trim();
        const studentId = cStudentId >= 0 ? String(r[cStudentId] || '').trim() : '';
        const firstName = cFirstName >= 0 ? String(r[cFirstName] || '').trim() : '';
        const lastName = cLastName >= 0 ? String(r[cLastName] || '').trim() : '';
        const studentName = (firstName + ' ' + lastName).trim() || studentId;
        if (!supervisor || (!studentName && !studentId)) continue;

        const school = cSchool >= 0 ? String(r[cSchool] || '').trim() : '';
        if (school) schools.add(school);

        const percentStr = String(r[cPercent] || '').trim();
        const percent = parseFloat(percentStr);
        if (isNaN(percent) || percent <= 0) continue;

        // Only include active students
        const status = cStudentStatus >= 0 ? String(r[cStudentStatus] || '').trim() : '';
        if (status && status !== 'Active in Programme') continue;

        const hours = 100 * (percent / 100);

        data.push({
          studentName,
          studentId,
          supervisor,
          supervisorEmail: cSupervisorEmail >= 0 ? String(r[cSupervisorEmail] || '').trim() : '',
          supervisorRole: cSupervisorRole >= 0 ? String(r[cSupervisorRole] || '').trim() : '',
          school,
          planTitle: cPlanTitle >= 0 ? String(r[cPlanTitle] || '').trim() : '',
          planCode: cPlanCode >= 0 ? String(r[cPlanCode] || '').trim() : '',
          programTitle: cProgramTitle >= 0 ? String(r[cProgramTitle] || '').trim() : '',
          programCode: cProgramCode >= 0 ? String(r[cProgramCode] || '').trim() : '',
          studentStatus: status,
          researchGroup: cResearchGroup >= 0 ? String(r[cResearchGroup] || '').trim() : '',
          percent,
          hours,
        });
      }

      if (data.length === 0) throw new Error('No active PGR students found (all had 0% supervision or non-active status).');

      pgrAllData = data;
      populateSchoolFilter(schools);
      applySchoolFilterAndRender();
      document.getElementById('pgrError').style.display = 'none';
    } catch(err) {
      const el = document.getElementById('pgrError');
      el.textContent = 'Error: ' + err.message;
      el.style.display = 'block';
    }
  };
  reader.readAsText(file);
}

function populateSchoolFilter(schools) {
  const sel = document.getElementById('pgrSchoolFilter');
  if (!sel) return;
  sel.innerHTML = '<option value="__all__">All Schools</option>';
  const sorted = [...schools].sort();
  sorted.forEach(s => {
    const opt = document.createElement('option');
    opt.value = s;
    opt.textContent = s;
    sel.appendChild(opt);
  });
  if (sorted.includes('Pharmacy')) sel.value = 'Pharmacy';
}

function getPgrSelectedSchool() {
  const sel = document.getElementById('pgrSchoolFilter');
  return sel ? sel.value : '__all__';
}

function getPgrFilteredData() {
  const selected = getPgrSelectedSchool();
  if (selected === '__all__') return pgrAllData;
  // Find all supervisors who have at least one student in the selected school,
  // then return ALL of their supervision rows (including students from other schools).
  const schoolSupervisors = new Set(
    pgrAllData.filter(r => r.school === selected).map(r => r.supervisor)
  );
  return pgrAllData.filter(r => schoolSupervisors.has(r.supervisor));
}

function applySchoolFilterAndRender() {
  renderPgr(getPgrFilteredData());
}

function pgrGetSorted() {
  const data = getPgrFilteredData();
  const selected = getPgrSelectedSchool();
  const q = document.getElementById('pgrSearch')?.value.toLowerCase() || '';
  let filtered = data.filter(r =>
    r.supervisor.toLowerCase().includes(q) || r.studentName.toLowerCase().includes(q)
  );
  const sel = document.getElementById('pgrSortSel')?.value || 'hours-desc';
  const lastDash = sel.lastIndexOf('-');
  const sc = sel.substring(0, lastDash);
  const sd = sel.substring(lastDash + 1);

  const map = new Map();
  filtered.forEach(r => {
    if (!map.has(r.supervisor)) {
      map.set(r.supervisor, { supervisor: r.supervisor, hours: 0, schoolStudentNames: new Set(), totalStudentNames: new Set() });
    }
    const entry = map.get(r.supervisor);
    entry.hours += r.hours;
    entry.totalStudentNames.add(r.studentName);
    if (selected === '__all__' || r.school === selected) {
      entry.schoolStudentNames.add(r.studentName);
    }
  });

  let aggregated = [...map.values()].map(e => ({
    supervisor: e.supervisor,
    hours: e.hours,
    schoolStudentCount: e.schoolStudentNames.size,
    totalStudentCount: e.totalStudentNames.size,
  }));

  aggregated.sort((a, b) => {
    if (sc === 'supervisor' || sc === 'name') {
      return sd === 'asc' ? a.supervisor.localeCompare(b.supervisor) : b.supervisor.localeCompare(a.supervisor);
    }
    if (sc === 'school-students' || sc === 'students') return sd === 'asc' ? a.schoolStudentCount - b.schoolStudentCount : b.schoolStudentCount - a.schoolStudentCount;
    if (sc === 'total-students') return sd === 'asc' ? a.totalStudentCount - b.totalStudentCount : b.totalStudentCount - a.totalStudentCount;
    return sd === 'asc' ? a.hours - b.hours : b.hours - a.hours;
  });
  return aggregated;
}

function renderPgrTable() {
  const data = pgrGetSorted();
  const selected = getPgrSelectedSchool();
  const isAll = selected === '__all__';
  const maxHours = Math.max(...data.map(r => r.hours), 1);
  document.getElementById('pgrTbody').innerHTML = data.map(r => {
    const enc = encodeURIComponent(r.supervisor);
    const barPct = Math.min(r.hours / maxHours * 100, 100).toFixed(1);
    return `<tr>
      <td class="name-f" data-supervisor="${enc}" style="cursor:pointer">${r.supervisor}</td>
      <td class="num">${r.schoolStudentCount}</td>
      <td class="num">${r.totalStudentCount}</td>
      <td class="num">${r.hours.toFixed(1)}</td>
      <td><div class="pgr-hrs-bar-wrap"><div class="pgr-hrs-bar"><div class="pgr-hrs-bar-fill" style="width:${barPct}%" title="${r.hours.toFixed(1)} hours"></div></div></td>
    </tr>`;
  }).join('');
  document.querySelectorAll('#pgrTbody .name-f').forEach(cell => {
    cell.addEventListener('click', () => pgrOpenDetail(decodeURIComponent(cell.dataset.supervisor)));
  });

  const filteredData = getPgrFilteredData();
  const totalHours = data.reduce((s, r) => s + r.hours, 0);
  const schoolStudents = isAll ? new Set(filteredData.map(d => d.studentName)).size : new Set(filteredData.filter(d => d.school === selected).map(d => d.studentName)).size;
  const totalStudents = new Set(filteredData.map(d => d.studentName)).size;
  document.getElementById('pgrFoot').innerHTML = `<tr>
    <td><strong>Total (${data.length} supervisors)</strong></td>
    <td class="num"><strong>${schoolStudents}</strong></td>
    <td class="num"><strong>${totalStudents}</strong></td>
    <td class="num"><strong>${totalHours.toFixed(1)}</strong></td>
    <td></td>
  </tr>`;
}

function pgrOpenDetail(supervisorName) {
  const filteredData = getPgrFilteredData();
  const rows = filteredData.filter(r => r.supervisor === supervisorName);
  if (rows.length === 0) return;
  const totalHours = rows.reduce((s, r) => r.hours + s, 0);
  const studentCount = new Set(rows.map(r => r.studentName)).size;

  let html = `<div class="panel-section"><h4>Hours Breakdown</h4>
    <div class="panel-row"><span class="k">Total hours</span><span class="v big">${totalHours.toFixed(1)}h</span></div>
    <div class="panel-row"><span class="k">Students</span><span class="v">${studentCount}</span></div>
  </div>`;
  html += `<div class="panel-section"><h4>Students (${rows.length} supervisor rows)</h4>`;
  const sorted = [...rows].sort((a, b) => b.percent - a.percent);
  for (const r of sorted) {
    const encStudent = encodeURIComponent(r.studentName);
    html += `<div class="proj-student">
      <div class="sn pgr-student-link" data-student="${encStudent}" style="cursor:pointer">${r.studentName}${r.studentId ? ' <span class="text-muted">#' + r.studentId + '</span>' : ''}</div>
      <div class="sr">
        ${r.school ? '<span style="margin-right:8px">🏫 ' + r.school + '</span>' : ''}
        ${r.planTitle ? '<span style="margin-right:8px">📋 ' + r.planTitle + '</span>' : ''}
        ${r.supervisorRole ? '<span style="margin-right:8px">👤 ' + r.supervisorRole + '</span>' : ''}
        ${r.researchGroup && r.researchGroup !== 'None' ? '<span style="margin-right:8px">🔬 ' + r.researchGroup + '</span>' : ''}
        <span style="font-weight:600;color:var(--teal)">${r.hours.toFixed(1)}h @ ${r.percent}%</span>
      </div>
    </div>`;
  }
  html += '</div>';
  openPanel(supervisorName, `${totalHours.toFixed(1)}h total · ${studentCount} student${studentCount !== 1 ? 's' : ''}`, html);
  // Attach click handlers for student deep-dive
  document.querySelectorAll('#panelBody .pgr-student-link').forEach(el => {
    el.addEventListener('click', e => {
      e.stopPropagation();
      pgrOpenStudentDetail(decodeURIComponent(el.dataset.student));
    });
  });
}

function pgrOpenStudentDetail(studentName) {
  // Look across ALL data (not just filtered) to find all supervisors for this student
  const rows = pgrAllData.filter(r => r.studentName === studentName);
  if (rows.length === 0) return;

  const totalPct = rows.reduce((s, r) => s + r.percent, 0);
  const totalHours = rows.reduce((s, r) => s + r.hours, 0);
  const studentId = rows.find(r => r.studentId)?.studentId || '—';
  const school = rows.find(r => r.school)?.school || '—';
  const planTitle = rows.find(r => r.planTitle)?.planTitle || '—';
  const researchGroup = rows.find(r => r.researchGroup && r.researchGroup !== 'None')?.researchGroup;
  const status = rows.find(r => r.studentStatus)?.studentStatus || '—';

  let html = `<div class="panel-section"><h4>Student Info</h4>
    <div class="panel-row"><span class="k">Student</span><span class="v">${studentName}</span></div>
    ${studentId !== '—' ? '<div class="panel-row"><span class="k">ID</span><span class="v">' + studentId + '</span></div>' : ''}
    <div class="panel-row"><span class="k">School</span><span class="v">${school}</span></div>
    <div class="panel-row"><span class="k">Plan</span><span class="v">${planTitle}</span></div>
    ${researchGroup ? '<div class="panel-row"><span class="k">Research Group</span><span class="v">' + researchGroup + '</span></div>' : ''}
    <div class="panel-row"><span class="k">Status</span><span class="v">${status}</span></div>
    <div class="panel-row"><span class="k">Total supervision</span><span class="v">${totalPct.toFixed(0)}% · ${totalHours.toFixed(1)}h</span></div>
  </div>`;

  html += `<div class="panel-section"><h4>Supervision Team (${rows.length} supervisor${rows.length !== 1 ? 's' : ''})</h4>`;
  html += `<div style="display:grid;gap:0.5rem;">`;
  for (const r of rows) {
    // Build a visual bar to show this supervisor's share
    const barPct = Math.min(r.percent / (totalPct || 100) * 100, 100).toFixed(0);
    html += `<div style="background:var(--surface);border:1px solid var(--border);border-radius:8px;padding:0.6rem 0.8rem;display:flex;align-items:center;gap:0.8rem;">
      <div style="flex:1">
        <div style="font-weight:600;font-size:0.9rem">${r.supervisor}</div>
        <div style="font-size:0.78rem;color:var(--muted)">
          ${r.supervisorRole || 'Supervisor'}
          ${r.supervisorEmail ? ' · ' + r.supervisorEmail : ''}
        </div>
      </div>
      <div style="text-align:right;min-width:100px">
        <div style="font-weight:700;font-size:1rem;color:var(--teal)">${r.hours.toFixed(1)}h</div>
        <div style="font-size:0.75rem;color:var(--muted)">${r.percent}% share</div>
      </div>
      <div style="width:80px;height:6px;background:var(--border);border-radius:3px;overflow:hidden">
        <div style="height:100%;width:${barPct}%;background:var(--teal);border-radius:3px;transition:width 0.3s" title="${barPct}% of total"></div>
      </div>
    </div>`;
  }
  html += `</div></div>`;

  openPanel(studentName, `🏫 ${school} · ${totalHours.toFixed(1)}h total`, html);
}

function renderPgr(data) {
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

  const sel = document.getElementById('pgrSchoolFilter');
  const schoolLabel = sel && sel.value !== '__all__' ? ` · ${sel.value}` : '';
  document.getElementById('pgrMeta').textContent = `${totalStudents} students · ${totalSupervisors} supervisors${schoolLabel}`;
  updateCombStatus();
}

document.addEventListener('DOMContentLoaded', () => {
  const inp = document.getElementById('pgrFileInput');
  if (inp) inp.addEventListener('change', e => { if (e.target.files[0]) parsePgrCsv(e.target.files[0]); });

  const dz = document.getElementById('pgrDropZone');
  if (dz) {
    dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('drag-over'); });
    dz.addEventListener('dragleave', () => dz.classList.remove('drag-over'));
    dz.addEventListener('drop', e => { e.preventDefault(); dz.classList.remove('drag-over'); const f = e.dataTransfer.files[0]; if (f) parsePgrCsv(f); });
  }

  const schoolFilter = document.getElementById('pgrSchoolFilter');
  if (schoolFilter) {
    schoolFilter.addEventListener('change', () => {
      if (pgrAllData.length > 0) applySchoolFilterAndRender();
    });
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
    const s = document.getElementById('pgrSortSel');
    if (!s) return;
    const isNum = ['school-students', 'total-students', 'hours'].includes(col);
    let dir;
    if (s.value.startsWith(col)) {
      dir = s.value.endsWith('desc') ? 'asc' : 'desc';
    } else {
      dir = isNum ? 'desc' : 'asc';
    }
    const opt = [...s.options].find(o => o.value === `${col}-${dir}`);
    if (opt) { s.value = opt.value; renderPgrTable(); }
  });

  const expBtn = document.getElementById('pgrBtnExport');
  if (expBtn) expBtn.addEventListener('click', () => {
    const filteredData = getPgrFilteredData();
    if (filteredData.length === 0) return;
    const rows = [['Student Name', 'Student ID', 'Supervisor', 'Supervisor Role', 'School/Department', 'Plan Title', 'Plan Code', 'Student Status', 'Research Group', 'Supervision %', 'Hours']];
    filteredData.forEach(r => rows.push([r.studentName, r.studentId, r.supervisor, r.supervisorRole, r.school, r.planTitle, r.planCode, r.studentStatus, r.researchGroup, r.percent, r.hours.toFixed(1)]));
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'PGR Supervision');
    XLSX.writeFile(wb, 'pgr_supervision.xlsx');
  });
});
