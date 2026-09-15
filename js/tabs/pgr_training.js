// ═══════════════════════════════════════════════════════
// PGT TRAINING TAB
// ═══════════════════════════════════════════════════════
let pgrTrainingAllData = [];

window.getPgtTrainingHoursTotals = function() {
  const totals = {};
  pgrTrainingAllData.forEach(r => {
    if (r.name && r.hours > 0) totals[r.name] = (totals[r.name] || 0) + r.hours;
  });
  return totals;
};

/** Parse a time string like "12-2", "11-12:30", "09:00-16:00", "2-4 pm", "11-12.30", "1.5", "1.5h" into decimal hours */
function pgrParseTime(timeStr) {
  if (!timeStr) return 0;
  let s = String(timeStr).trim();
  if (!s) return 0;

  // Strip am/pm markers (with optional leading space)
  let pm = false;
  if (/pm/i.test(s)) { pm = true; s = s.replace(/\s*pm\s*/i, '').trim(); }
  s = s.replace(/\s*am\s*/i, '').trim();

  // Try explicit decimal format FIRST: "1.5", "1.5h", "2hrs", "2 hours"
  // (Must be before dot normalisation which would break "1.5" -> "1:5")
  const decimalMatch = s.match(/^(\d+(?:\.\d+)?)\s*(?:h|hr|hrs|hours)?$/i);
  if (decimalMatch) return parseFloat(decimalMatch[1]);

  // Normalise dots to colons in time patterns (e.g. "12.30" -> "12:30")
  s = s.replace(/\.(\d{2})\b/g, ':$1');

  // Try range format: "12-2", "11-12:30", "09:00-16:00", "9-5", "2-4 pm" (pm already stripped), "11-12.30"
  const rangeMatch = s.match(/^(\d{1,2})(?::(\d{2}))?\s*[-–]\s*(\d{1,2})(?::(\d{2}))?$/);
  if (rangeMatch) {
    let startH = parseInt(rangeMatch[1], 10);
    const startM = parseInt(rangeMatch[2] || '0', 10);
    let endH = parseInt(rangeMatch[3], 10);
    const endM = parseInt(rangeMatch[4] || '0', 10);

    // If pm was specified, convert both times
    if (pm) {
      if (startH < 12) startH += 12;
      if (endH < 12) endH += 12;
    } else {
      // Smart PM detection: end < start or end <= 7 && start >= 8
      if (endH < startH && startH <= 12) endH += 12;
      else if (endH <= 7 && startH >= 8) endH += 12;
    }

    const startDec = startH + startM / 60;
    const endDec = endH + endM / 60;
    return Math.max(0, endDec - startDec);
  }

  return 0;
}

/** Parse a single table's text content into rows */
function pgrParseTable(text) {
  const rows = [];
  if (!text || !text.trim()) return rows;

  const rawLines = text.split('\n');
  // Remove trailing blank lines but keep internal structure
  while (rawLines.length && !rawLines[rawLines.length - 1].trim()) rawLines.pop();

  // ── Detect headers ──────────────────────────────────────
  // Try vertical headers first: 5 consecutive single-field lines
  function normW(s) { return String(s).toLowerCase().replace(/[^a-z0-9 /]/g, '').trim(); }
  const VERT_HDRS = ['course name', 'course title', 'date', 'time', 'course provider', 'provider', 'further information', 'further info'];

  function isVertHeader(lines, start) {
    if (start + 4 >= lines.length) return false;
    const norms = [];
    for (let i = 0; i < 5; i++) {
      const trimmed = lines[start + i].trim();
      if (!trimmed) return false;
      norms.push(normW(trimmed));
    }
    // Check we have at least: a course-type, date, time, and provider-type
    const hasCourse = norms.slice(0, 2).some(n => /course/.test(n) || /course.?name/.test(n) || /course.?title/.test(n));
    const hasDate = norms.slice(1, 3).some(n => n === 'date' || n === 'date ');
    const hasTime = norms.slice(2, 4).some(n => n === 'time' || n === 'times');
    const hasProvider = norms.slice(2, 5).some(n => /provider/.test(n) || /course.?provider/.test(n) || n === 'staff');
    const hasInfo = norms.some(n => /further/.test(n) || n === 'notes');
    // Must have course, date/time-like, and provider
    const timeOrDate = hasDate || hasTime;
    if (hasCourse && hasProvider && timeOrDate) return { course: start, date: start + 1, time: start + 2, provider: start + 3, info: start + 4 };
    return false;
  }

  // Try horizontal headers: one line with all 5 fields
  function isHdrLine(line) {
    const trimmed = line.trim().toLowerCase();
    return trimmed.includes('course name') || trimmed.includes('course title') || trimmed.includes('course provider');
  }

  // Find best header approach
  let hdrType = null; // 'vertical' or 'horizontal'
  let hdrInfo = null;

  // Check for vertical headers
  for (let i = 0; i < Math.min(8, rawLines.length); i++) {
    const v = isVertHeader(rawLines, i);
    if (v) { hdrType = 'vertical'; hdrInfo = v; break; }
  }

  // Check for horizontal headers
  let hdrLineIdx = -1;
  if (!hdrType) {
    for (let i = 0; i < Math.min(10, rawLines.length); i++) {
      if (isHdrLine(rawLines[i])) { hdrLineIdx = i; break; }
    }
    if (hdrLineIdx >= 0) hdrType = 'horizontal';
  }

  if (!hdrType) return rows;  // Can't find headers

  // ── Parse data ──────────────────────────────────────────
  // Determine delimiter from a line that has multiple fields
  let delim = '\t';
  function sniffDelimiter(line) {
    if (line.includes('\t') && line.split('\t').length >= 4) return '\t';
    const cCount = (line.match(/,/g) || []).length;
    if (cCount >= 4) return ',';
    const sp = line.split(/\s{2,}/);
    if (sp.length >= 4) return /\s{2,}/;
    return '\t';
  }

  // Get data starting line
  let dataStart;
  if (hdrType === 'vertical') {
    dataStart = hdrInfo.info + 1;
    // Sniff delimiter from the first data row
    for (let i = dataStart; i < Math.min(dataStart + 5, rawLines.length); i++) {
      const l = rawLines[i].trim();
      if (l) { delim = sniffDelimiter(l); break; }
    }
  } else {
    dataStart = hdrLineIdx + 1;
    delim = sniffDelimiter(rawLines[hdrLineIdx]);
  }

  // ── Reassemble multi-line rows ────────────────────────────
  // A row should have course name + date/time + provider.
  // Lines that don't form a complete row are appended to the previous incomplete row.
  const dataLines = [];
  for (let i = dataStart; i < rawLines.length; i++) {
    const line = rawLines[i].trim();
    if (!line) continue;  // skip blank lines

    const cells = delim === '\t' ? line.split('\t')
      : delim === ',' ? line.split(',')
      : line.split(delim);
    const hasCourse = cells[0] && cells[0].trim().length > 2;  // course names are more than 2 chars
    const hasProvider = delim === '\t' || delim === ',' ? cells.length >= 4
      : cells.length >= 4;
    // Check if this line starts a new record (has a course name-ish first cell)
    const looksLikeNewRow = hasCourse && cells.length >= 3;

    if (looksLikeNewRow) {
      dataLines.push(line);
    } else if (dataLines.length > 0) {
      // Append to previous line with a space
      dataLines[dataLines.length - 1] += ' ' + line;
    }
  }

  // ── Parse reassembled data lines ─────────────────────────
  // Build column index mapping based on header type
  let ciCourse, ciDate, ciTime, ciProvider, ciInfo;

  if (hdrType === 'vertical') {
    ciCourse = 0; ciDate = 1; ciTime = 2; ciProvider = 3; ciInfo = 4;
  } else {
    const hdrLine = rawLines[hdrLineIdx];
    const hdrs = delim === '\t' ? hdrLine.split('\t')
      : delim === ',' ? hdrLine.split(',')
      : hdrLine.split(delim);
    const hdrLower = hdrs.map(h => h.trim().toLowerCase());
    function fIdx(kw) {
      const idx = hdrLower.findIndex(h => {
        for (const k of kw) {
          if (h === k) return true;
          const re = new RegExp('\\b' + k.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\b', 'i');
          if (re.test(h)) {
            // For 'name' / 'staff' keywords, exclude if the column also contains 'course' (to avoid matching 'course name')
            if ((k === 'name' || k === 'staff') && /\bcourse\b/i.test(h) && !/\bprovider\b/i.test(h)) continue;
            return true;
          }
        }
        return false;
      });
      return idx >= 0 ? idx : -1;
    }
    ciCourse   = fIdx(['course name', 'course title', 'course']);
    ciDate     = fIdx(['date']);
    ciTime     = fIdx(['time', 'times', 'duration']);
    ciProvider = fIdx(['course provider', 'provider', 'course provider/contact', 'staff', 'name']);
    ciInfo     = fIdx(['further information', 'further info', 'notes', 'information', 'info']);
    if (ciCourse < 0 || ciProvider < 0) return rows;
  }

  for (const line of dataLines) {
    const cells = delim === '\t' ? line.split('\t')
      : delim === ',' ? line.split(',')
      : line.split(delim);

    let courseName, date, timeRaw, providerRaw, info;

    if (hdrType === 'vertical') {
      // Columns are positional in the split: treat first cell as course name, etc.
      // But the entire line is one row after reassembly, split by delimiter
      courseName = cells[0] ? cells[0].trim() : '';
      date = cells[1] ? cells[1].trim() : '';
      timeRaw = cells[2] ? cells[2].trim() : '';
      providerRaw = cells[3] ? cells[3].trim() : '';
      info = cells[4] ? cells.slice(5).join(delim === '\t' ? ' ' : ',').trim() : '';
    } else {
      courseName = ciCourse >= 0 && cells[ciCourse] ? cells[ciCourse].trim() : '';
      date = ciDate >= 0 && cells[ciDate] ? cells[ciDate].trim() : '';
      timeRaw = ciTime >= 0 && cells[ciTime] ? cells[ciTime].trim() : '';
      providerRaw = ciProvider >= 0 && cells[ciProvider] ? cells[ciProvider].trim() : '';
      info = ciInfo >= 0 && cells[ciInfo] ? cells[ciInfo].trim() : '';
    }

    if (!courseName || !providerRaw) continue;

    const hours = pgrParseTime(timeRaw);
    if (hours <= 0) continue;

    // Split comma-delimited providers — each gets full hours
    const providers = providerRaw.split(',').map(p => p.trim()).filter(p => p);
    providers.forEach(name => {
      if (name) {
        rows.push({ courseName, date, time: timeRaw, provider: name, furtherInfo: info, name, hours });
      }
    });
  }

  return rows;
}

/** Parse all three textareas and render */
function parsePgtTrainingTables() {
  const t1 = document.getElementById('pgrtr-t1')?.value || '';
  const t2 = document.getElementById('pgrtr-t2')?.value || '';
  const t3 = document.getElementById('pgrtr-t3')?.value || '';
  const errorEl = document.getElementById('pgrtrError');

  const data = [
    ...pgrParseTable(t1),
    ...pgrParseTable(t2),
    ...pgrParseTable(t3)
  ];

  if (data.length === 0) {
    errorEl.textContent = 'No valid data found. Check that each table has headers: Course name, Date, Time, Course provider, Further Information';
    errorEl.classList.add('show');
    return;
  }

  errorEl.classList.remove('show');
  renderPgtTraining(data);
  WLP_SESSION.saveFile('pgr_training', 'pgr_training.txt', new TextEncoder().encode(JSON.stringify({
    t1, t2, t3
    })).buffer, 'application/json');
}

// ── Aggregation, sorting, search ─────────────────────

function pgrGetSorted() {
  const q = document.getElementById('pgrtrSearch')?.value.toLowerCase() || '';
  let data = pgrTrainingAllData.filter(r =>
    r.name.toLowerCase().includes(q) || r.courseName.toLowerCase().includes(q)
  );
  const sel = document.getElementById('pgrtrSortSel')?.value || 'hours-desc';
  const [sc, sd] = sel.split('-');

  // Aggregate by staff member
  const aggregated = [];
  const seen = new Set();
  data.forEach(r => {
    if (!seen.has(r.name)) {
      seen.add(r.name);
      const personRows = pgrTrainingAllData.filter(d => d.name === r.name);
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

function pgrOpenDetail(name) {
  const rows = pgrTrainingAllData.filter(r => r.name === name);
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
      <div style="font-weight:600;font-size:0.85rem;margin-bottom:4px">${r.courseName || '—'}</div>
      <div class="panel-row"><span class="k">Date</span><span class="v">${r.date || '—'}</span></div>
      <div class="panel-row"><span class="k">Time</span><span class="v">${r.time || '—'}</span></div>
      <div class="panel-row"><span class="k">Hours</span><span class="v big" style="color:var(--teal)">${r.hours.toFixed(1)}h</span></div>
      ${r.furtherInfo ? `<div class="panel-row"><span class="k">Notes</span><span class="v" style="font-size:0.78rem;color:var(--muted)">${r.furtherInfo}</span></div>` : ''}
    </div>`;
  });
  html += '</div>';
  openPanel(name, `${totalHours.toFixed(1)}h total · ${rows.length} session(s)`, html);
}

function renderPgtTrainingTable() {
  const data = pgrGetSorted();
  const maxHours = Math.max(...data.map(r => r.hours), 1);
  document.getElementById('pgrtrTbody').innerHTML = data.map(r => {
    const barPct = Math.min(r.hours / maxHours * 100, 100).toFixed(1);
    return `<tr>
      <td class="name-f" data-name="${encodeURIComponent(r.name)}" style="cursor:pointer">${r.name}</td>
      <td class="num">${r.sessions}</td>
      <td class="num">${r.hours.toFixed(1)}</td>
      <td><div class="pgr-hrs-bar-wrap"><div class="pgr-hrs-bar"><div class="pgr-hrs-bar-fill" style="width:${barPct}%" title="${r.hours.toFixed(1)} hours"></div></div></td>
    </tr>`;
  }).join('');
  document.querySelectorAll('#pgrtrTbody .name-f').forEach(cell => {
    cell.addEventListener('click', () => pgrOpenDetail(decodeURIComponent(cell.dataset.name)));
  });
  const totalHours = data.reduce((s, r) => s + r.hours, 0);
  const totalSessions = data.reduce((s, r) => s + r.sessions, 0);
  document.getElementById('pgrtrFoot').innerHTML = `<tr>
    <td><strong>Total (${data.length} staff)</strong></td>
    <td class="num"><strong>${totalSessions}</strong></td>
    <td class="num"><strong>${totalHours.toFixed(1)}</strong></td>
    <td></td>
  </tr>`;
}

function renderPgtTraining(data) {
  pgrTrainingAllData = data;
  const totalHours = data.reduce((s, r) => s + r.hours, 0);
  const totalStaff = new Set(data.map(r => r.name)).size;
  document.getElementById('pgrtrStatsBar').innerHTML = `
    <div class="stat-card" style="border-left:3px solid var(--teal)"><div class="sc-v">${totalStaff}</div><div class="sc-l">Staff with PGR Training</div></div>
    <div class="stat-card" style="border-left:3px solid var(--gold)"><div class="sc-v">${totalHours.toFixed(0)}</div><div class="sc-l">Total PGR Training Hours</div></div>
    <div class="stat-card" style="border-left:3px solid var(--mid-blue)"><div class="sc-v">${data.length}</div><div class="sc-l">Session entries</div></div>
  `;
  renderPgtTrainingTable();
  document.getElementById('pgr-training-landing').style.display = 'none';
  document.getElementById('pgr-training-content').style.display = 'block';
  document.getElementById('badge-pgr_training').textContent = totalStaff;
  document.getElementById('pgrtrMeta').textContent = `${totalStaff} staff · ${totalHours.toFixed(1)} hours · ${data.length} entries`;
  updateCombStatus();
}

document.addEventListener('DOMContentLoaded', () => {
  const btn = document.getElementById('pgrtrAnalyseBtn');
  if (btn) btn.addEventListener('click', parsePgtTrainingTables);

  const back = document.getElementById('pgrtrBtnBack');
  if (back) back.addEventListener('click', () => {
    document.getElementById('pgr-training-content').style.display = 'none';
    document.getElementById('pgr-training-landing').style.display = 'block';
    pgrTrainingAllData = [];
    document.getElementById('badge-pgr_training').textContent = '—';
    updateCombStatus();
  });

  const sortSel = document.getElementById('pgrtrSortSel');
  if (sortSel) sortSel.addEventListener('change', renderPgtTrainingTable);
  const search = document.getElementById('pgrtrSearch');
  if (search) search.addEventListener('input', renderPgtTrainingTable);

  // Header click sorting
  document.addEventListener('click', e => {
    const th = e.target.closest('th[data-pgrsort]');
    if (!th) return;
    const col = th.dataset.pgrsort;
    const sel = document.getElementById('pgrtrSortSel');
    if (!sel) return;
    const isNum = ['hours', 'sessions'].includes(col);
    const dir = (sel.value.startsWith(col) && sel.value.endsWith('desc')) ? 'asc' : (isNum ? 'desc' : 'asc');
    const opt = [...sel.options].find(o => o.value === `${col}-${dir}`);
    if (opt) { sel.value = opt.value; renderPgtTrainingTable(); }
  });

  const expBtn = document.getElementById('pgrtrBtnExport');
  if (expBtn) expBtn.addEventListener('click', () => {
    if (pgrTrainingAllData.length === 0) return;
    const hdr = ['Staff member', 'Course name', 'Date', 'Time', 'Hours', 'Further Information'];
    const rows = [hdr];
    pgrTrainingAllData.forEach(r => rows.push([r.name, r.courseName, r.date, r.time, r.hours, r.furtherInfo]));
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'PGR Training');
    XLSX.writeFile(wb, 'pgr_training.xlsx');
  });
});
