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

// Parse a date-ish cell from the source XLSX into a Date. Handles:
//   * Date objects (when sheet_to_json is called with cellDates: true)
//   * Excel serial date numbers (days since 1900-01-01, including the
//     famous 1900 leap-year bug — SheetJS treats 1 as 1900-01-01, so
//     serial 1 = 1900-01-00, but the canonical mapping for dates after
//     1900-02-28 is serial - 2 → day count from 1900-01-01)
//   * DD/MM/YYYY UK-style strings (used in the ESD column of the 2026-27 file)
//   * ISO YYYY-MM-DD strings (used in the Start date column)
function pgrParseDate(v) {
  if (v == null || v === '') return null;
  if (v instanceof Date && !isNaN(v)) return v;
  if (typeof v === 'number' && isFinite(v) && v > 0) {
    // Excel serial → JS Date. Account for the 1900 leap-year bug: the
    // serial 60 corresponds to 1900-02-29 (which never existed), so
    // anything after serial 60 is effectively serial - 1 day.
    const epoch = new Date(Date.UTC(1899, 11, 30));
    const ms = v * 86400 * 1000;
    return new Date(epoch.getTime() + ms);
  }
  if (typeof v === 'string') {
    const s = v.trim();
    const uk = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (uk) {
      const d = parseInt(uk[1], 10), mo = parseInt(uk[2], 10) - 1, yr = parseInt(uk[3], 10);
      const dt = new Date(yr, mo, d);
      return isNaN(dt) ? null : dt;
    }
    const iso = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (iso) {
      const yr = parseInt(iso[1], 10), mo = parseInt(iso[2], 10) - 1, d = parseInt(iso[3], 10);
      const dt = new Date(yr, mo, d);
      return isNaN(dt) ? null : dt;
    }
  }
  return null;
}

// Derive a school/department label from the plan/programme string. The
// 2026-27 file has no School/Department column, so we infer it for the
// school filter dropdown.
function pgrInferSchool(plan) {
  if (!plan) return 'Other';
  const p = String(plan).toLowerCase();
  if (p.includes('pharmacy')) return 'Pharmacy';
  if (p.includes('chemistry')) return 'Chemistry';
  if (p.includes('physics')) return 'Physics';
  if (p.includes('engineering')) return 'Engineering';
  if (p.includes('medicine') || p.includes('medical')) return 'Medicine';
  if (p.includes('bioscience') || p.includes('biology')) return 'Biosciences';
  if (p.includes('math')) return 'Mathematical Sciences';
  if (p.includes('computer')) return 'Computer Science';
  return 'Other';
}

function parsePgrCsv(file) {
  const reader = new FileReader();
  reader.onload = e => {
    try {
      // Read as ArrayBuffer so XLSX can sniff XLSX (binary) vs CSV (text).
      const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', cellDates: true });
      if (rows.length < 2) throw new Error('No data rows found.');

      // Find header row (scan first 5 rows for a known marker).
      let hdrIdx = 0;
      for (let i = 0; i < Math.min(5, rows.length); i++) {
        const lower = rows[i].map(c => String(c).toLowerCase().trim());
        if (lower.some(c =>
              c.includes('supervisor name') ||
              c.includes('supervisor id') ||
              c === 'pi' ||
              c === 'first name'
            )) { hdrIdx = i; break; }
      }
      const hdrs = rows[hdrIdx].map(c => String(c).toLowerCase().trim());

      function col(keywords) {
        const idx = hdrs.findIndex(h => keywords.some(k => h.includes(k)));
        return idx >= 0 ? idx : -1;
      }
      function colExact(name) {
        const idx = hdrs.findIndex(h => h === name);
        return idx >= 0 ? idx : -1;
      }

      // Detect which layout this file uses.
      //   * Legacy: one row per supervisor–student pair, with a single
      //     "Supervisor Name" + "Supervision Percentage" column.
      //   * 2026-27: one row per student, with up to 5 supervisor pairs
      //     (PI, 2nd, 3rd, 4th, 5th) and matching % columns.
      const cSupervisorName = col(['supervisor name']);
      const cPercentLegacy = col(['supervision percentage', 'percentage', '%']);
      const isLegacyLayout = cSupervisorName >= 0 && cPercentLegacy >= 0;

      const cStudentId = col(['student id']);
      const cFirstName = col(['firstname', 'first name', 'first']);
      const cLastName  = col(['lastname', 'last name', 'surname', 'last']);
      const cSchool    = col(['school/department', 'school', 'department']);
      const cPlanTitle = col(['plan title', 'plan/programme', 'plan / programme']);
      const cPlanCode  = col(['plan code']);
      const cProgramTitle = col(['program title', 'programme title']);
      const cProgramCode  = col(['program code', 'programme code']);
      const cStudentStatus = col(['student status']);
      const cResearchGroup = col(['research group']);
      const cStartDate = col(['start date']);
      const cEndDate   = col(['end date']);

      // 2026-27: index the supervisor / % pairs by their exact "PI",
      // "2nd Supervisor", "% Supervisor 2" header text.
      const supervisorPairs = isLegacyLayout ? [] : [
        { role: 'PI',           name: colExact('pi'),                pct: colExact('% pi') },
        { role: '2nd Supervisor', name: col(['2nd supervisor']),       pct: col(['% supervisor 2']) },
        { role: '3rd Supervisor', name: col(['3rd supervisor']),       pct: col(['% supervisor 3']) },
        { role: '4th Supervisor', name: col(['4th supervisor']),       pct: col(['% supervisor 4']) },
        { role: '5th Supervisor', name: col(['5th supervisor']),       pct: col(['% supervisor 5']) },
      ].filter(p => p.name >= 0 && p.pct >= 0);

      if (isLegacyLayout) {
        if (cSupervisorName < 0) throw new Error('Could not find a "Supervisor Name" column.');
        if (cPercentLegacy < 0)  throw new Error('Could not find a "Supervision Percentage" column.');
      } else if (supervisorPairs.length === 0) {
        throw new Error('Could not find supervisor/percentage columns. Expected either a "Supervisor Name" + "Supervision Percentage" pair (legacy) or PI / 2nd Supervisor / 3rd Supervisor / 4th Supervisor / 5th Supervisor pairs (2026-27).');
      }

      const data = [];
      const schools = new Set();
      // Default filter: today (end date >= today). 2026-27 files have no
      // status column, so we use the end date as the activity signal.
      const today = new Date();
      today.setHours(0, 0, 0, 0);

      for (let i = hdrIdx + 1; i < rows.length; i++) {
        const r = rows[i];
        if (!r) continue;
        const firstName = cFirstName >= 0 ? String(r[cFirstName] || '').trim() : '';
        const lastName  = cLastName  >= 0 ? String(r[cLastName]  || '').trim() : '';
        const studentName = (firstName + ' ' + lastName).trim();
        const studentId = cStudentId >= 0 ? String(r[cStudentId] || '').trim() : '';
        if (!studentName && !studentId) continue;

        // Active filter.
        if (cStudentStatus >= 0) {
          const status = String(r[cStudentStatus] || '').trim();
          if (status && status !== 'Active in Programme') continue;
        } else if (cEndDate >= 0) {
          const endDate = pgrParseDate(r[cEndDate]);
          if (endDate && endDate < today) continue; // past end date
        }

        const school = cSchool >= 0
          ? String(r[cSchool] || '').trim()
          : pgrInferSchool(cPlanTitle >= 0 ? r[cPlanTitle] : '');
        if (school) schools.add(school);

        const planTitle = cPlanTitle >= 0 ? String(r[cPlanTitle] || '').trim() : '';
        const programTitle = cProgramTitle >= 0 ? String(r[cProgramTitle] || '').trim() : '';
        const researchGroup = cResearchGroup >= 0 ? String(r[cResearchGroup] || '').trim() : '';
        const studentStatus = cStudentStatus >= 0 ? String(r[cStudentStatus] || '').trim() : '';
        const startDate = cStartDate >= 0 ? String(r[cStartDate] || '').trim() : '';
        const endDateStr = cEndDate >= 0 ? String(r[cEndDate] || '').trim() : '';

        if (isLegacyLayout) {
          // One supervisor per row.
          const supervisor = String(r[cSupervisorName] || '').trim();
          if (!supervisor) continue;
          const percentStr = String(r[cPercentLegacy] || '').trim();
          const percent = parseFloat(percentStr);
          if (isNaN(percent) || percent <= 0) continue;
          data.push({
            studentName, studentId, supervisor,
            supervisorRole: '',
            school, planTitle, planCode: '',
            programTitle, programCode: '',
            studentStatus, researchGroup, percent,
            hours: percent, // legacy: hours = percentage
            startDate, endDate: endDateStr,
          });
        } else {
          // 2026-27: expand into one record per non-empty supervisor.
          let addedForStudent = false;
          for (const p of supervisorPairs) {
            const name = String(r[p.name] || '').trim();
            if (!name) continue;
            const pct = parseFloat(String(r[p.pct] || '').trim());
            if (isNaN(pct) || pct <= 0) continue;
            data.push({
              studentName, studentId,
              supervisor: name,
              supervisorRole: p.role,
              school, planTitle,
              planCode: '', programTitle, programCode: '',
              studentStatus, researchGroup,
              percent: pct,
              hours: pct, // hours = percentage (same convention as legacy)
              startDate, endDate: endDateStr,
            });
            addedForStudent = true;
          }
          if (!addedForStudent) continue; // skip rows with no supervisors
        }
      }

      if (data.length === 0) throw new Error('No active PGR students found (all had 0% supervision or were past their end date).');

      pgrAllData = data;
      populateSchoolFilter(schools);
      applySchoolFilterAndRender();
      document.getElementById('pgrError').style.display = 'none';
      WLP_SESSION.saveFile('pgr', file.name, e.target.result,
        file.name.toLowerCase().endsWith('.csv') ? 'text/csv' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    } catch(err) {
      const el = document.getElementById('pgrError');
      el.textContent = 'Error: ' + err.message;
      el.style.display = 'block';
    }
  };
  reader.readAsArrayBuffer(file);
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
