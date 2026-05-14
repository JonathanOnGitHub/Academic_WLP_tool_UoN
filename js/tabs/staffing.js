// ═══════════════════════════════════════════════════════
// STAFFING & CONCURRENCY ANALYSIS
// ═══════════════════════════════════════════════════════
// Analyses timetable sessions to find peak concurrency:
// how many distinct staff are needed simultaneously,
// when, and for which activity types.

let staffingResult = null;

// ── Utility ──────────────────────────────────────────

function timeToMinutes(t) {
  if (!t) return 0;
  const m = String(t).match(/(\d+):(\d+)/);
  if (m) return +m[1] * 60 + +m[2];
  return Math.round(timeToHours(t) * 60);
}

// ── Core analysis ────────────────────────────────────

function computeStaffingAnalysis() {
  staffingResult = null;
  if (!tlParsedSessions || tlParsedSessions.length === 0) return null;

  const [wFrom, wTo] = tlWeekRange;
  const SLOT = 30; // minutes

  // Phase 1: occupancy grid key → [staff names]
  const gridRaw = new Map();

  for (const sess of tlParsedSessions) {
    const fw = sess.weeks.filter(w => w >= wFrom && w <= wTo);
    if (fw.length === 0 || !sess.staff || sess.staff.length === 0) continue;

    const startMin = timeToMinutes(sess.start);
    const endMin = timeToMinutes(sess.end);
    if (startMin >= endMin) continue;

    for (const w of fw) {
      for (let m = startMin; m < endMin; m += SLOT) {
        const key = `${w}|${sess.day}|${m}`;
        if (!gridRaw.has(key)) gridRaw.set(key, []);
        const arr = gridRaw.get(key);
        for (const name of sess.staff) arr.push(name);
      }
    }
  }

  if (gridRaw.size === 0) return null;

  // Phase 2: dedup to distinct staff per slot, find peak
  const grid = new Map();
  let peak = 0, peakKey = null;
  for (const [key, staffArr] of gridRaw) {
    const set = new Set(staffArr);
    grid.set(key, set);
    if (set.size > peak) { peak = set.size; peakKey = key; }
  }

  // Phase 3: parse peak
  const [peakWeekStr, peakDay, peakMinStr] = (peakKey || '||').split('|');
  const peakWeek = +peakWeekStr;
  const peakMin = +peakMinStr;
  const peakTimeStart = formatHour(peakMin / 60);
  const peakTimeEnd = formatHour((peakMin + SLOT) / 60);

  // Sessions overlapping with peak slot
  const peakSessions = tlParsedSessions.filter(s =>
    s.weeks.includes(peakWeek) &&
    s.day === peakDay &&
    timeToMinutes(s.start) < peakMin + SLOT &&
    timeToMinutes(s.end) > peakMin
  );

  // Type breakdown at peak
  const typeAtPeak = {};
  for (const s of peakSessions) {
    const t = (s.type || '').trim() || 'Undefined';
    typeAtPeak[t] = (typeAtPeak[t] || 0) + s.staff.length;
  }

  // Phase 4: histogram bins
  const slotCounts = [...grid.values()].map(s => s.size);
  const binDefs = [
    { label: '0',       min: 0,  max: 0 },
    { label: '1–3', min: 1,  max: 3 },
    { label: '4–6', min: 4,  max: 6 },
    { label: '7–10', min: 7,  max: 10 },
    { label: '11–15',min: 11, max: 15 },
    { label: '16–20',min: 16, max: 20 },
    { label: '21–30',min: 21, max: 30 },
    { label: '31–50',min: 31, max: 50 },
    { label: '50+',     min: 50, max: Infinity },
  ];
  const bins = binDefs.map(b => ({
    ...b, count: slotCounts.filter(c => c >= b.min && c <= b.max).length
  }));
  const totalSlots = slotCounts.length;

  // Phase 5: week peaks
  const allWeeks = [...new Set(
    tlParsedSessions.flatMap(s => s.weeks.filter(w => w >= wFrom && w <= wTo))
  )].sort((a, b) => a - b);

  const weekPeaks = allWeeks.map(w => {
    let maxForWeek = 0;
    for (const [key, set] of grid) {
      if (key.startsWith(w + '|') && set.size > maxForWeek) maxForWeek = set.size;
    }
    return { week: w, peak: maxForWeek };
  });

  // Phase 6: totals
  const allStaff = new Set();
  let totalStaffHours = 0;
  for (const sess of tlParsedSessions) {
    for (const name of sess.staff) allStaff.add(name);
    const dur = sessionDuration(sess);
    const wCount = sess.weeks.filter(w => w >= wFrom && w <= wTo).length;
    totalStaffHours += dur * sess.staff.length * wCount;
  }

  // Phase 7: high-demand sessions (staff count >= 5, top 15)
  const highDemand = tlParsedSessions
    .filter(s => s.staff && s.staff.length >= 5)
    .sort((a, b) => b.staff.length - a.staff.length)
    .slice(0, 15);

  staffingResult = {
    grid, peak, peakKey, peakWeek, peakDay, peakTimeStart, peakTimeEnd,
    peakSessions, typeAtPeak,
    totalSlots, bins, slotCounts, weekPeaks,
    totalStaff: allStaff.size,
    totalStaffHours,
    totalSessionCount: tlParsedSessions.length,
    highDemand,
  };

  return staffingResult;
}

// ── Rendering ────────────────────────────────────────

function renderStaffingView() {
  const landing = document.getElementById('st-landing');
  const results = document.getElementById('st-results');

  const result = computeStaffingAnalysis();

  if (!result) {
    landing.style.display = '';
    results.style.display = 'none';
    const meta = document.getElementById('stMeta');
    if (meta) meta.textContent = '';
    document.getElementById('badge-staffing').textContent = '—';
    return;
  }

  landing.style.display = 'none';
  results.style.display = 'block';

  document.getElementById('badge-staffing').textContent =
    result.peak + ' peak · ' + result.totalStaff + ' staff';

  // Meta line
  document.getElementById('stMeta').textContent =
    `${result.totalSessionCount} sessions · ${result.totalSlots.toLocaleString()} time slots analysed · ${result.totalStaff} distinct staff · ${result.totalStaffHours.toFixed(0)} total staff-hours`;

  // ── Stats bar ──
  const avgStaffPerSlot = (result.slotCounts.reduce((s, c) => s + c, 0) / result.totalSlots).toFixed(1);
  document.getElementById('stStatsBar').innerHTML =
    `<div class="stat-card" style="border-left:3px solid var(--purple)">
      <div class="sc-v">${result.peak}</div>
      <div class="sc-l">Peak concurrency (staff)</div>
    </div>
    <div class="stat-card" style="border-left:3px solid var(--gold)">
      <div class="sc-v">${result.peakWeek}<span style="font-size:0.7rem;color:var(--muted)"> · ${result.peakDay}</span></div>
      <div class="sc-l">Peak time — ${result.peakTimeStart}–${result.peakTimeEnd}</div>
    </div>
    <div class="stat-card" style="border-left:3px solid var(--navy)">
      <div class="sc-v">${result.totalStaff}</div>
      <div class="sc-l">Distinct staff in timetable</div>
    </div>
    <div class="stat-card" style="border-left:3px solid var(--teal)">
      <div class="sc-v">${avgStaffPerSlot}</div>
      <div class="sc-l">Avg staff per time slot</div>
    </div>
    <div class="stat-card" style="border-left:3px solid var(--mid-blue)">
      <div class="sc-v">${result.totalStaffHours.toFixed(0)}</div>
      <div class="sc-l">Total staff-hours (all weeks)</div>
    </div>`;

  // ── Peak detail card ──
  const peakSessHtml = result.peakSessions.map(s =>
    `<div class="st-peak-session${((s.type||'').toLowerCase().includes('osc')?' st-osc':'')}">
      <span class="sps-type">${s.type || 'Session'}</span>
      <span style="color:var(--text)">${s.moduleCode || s.moduleTitle || s.sessionTitle || ''}</span>
      <span class="sps-count">${s.staff.length} staff</span>
      ${s.location ? `<span style="color:var(--muted);font-size:0.75rem">${s.location}</span>` : ''}
    </div>`
  ).join('');

  const typeBreakdown = Object.entries(result.typeAtPeak)
    .sort((a, b) => b[1] - a[1])
    .map(([t, c]) => `<span class="st-type-badge">${t}: ${c}</span>`)
    .join('');

  document.getElementById('stPeakCard').innerHTML =
    `<div class="st-peak-card">
      <div class="st-pc-title">Peak Slot: Week ${result.peakWeek} · ${result.peakDay} · ${result.peakTimeStart}–${result.peakTimeEnd}</div>
      <div style="font-size:0.82rem;color:var(--muted);margin-bottom:0.4rem">
        <strong style="color:var(--navy)">${result.peak}</strong> distinct staff needed simultaneously across
        <strong>${result.peakSessions.length}</strong> session${result.peakSessions.length > 1 ? 's' : ''}
        &nbsp;&middot;&nbsp; Types: ${typeBreakdown}
      </div>
      <div class="st-peak-sessions">${peakSessHtml}</div>
    </div>`;

  // ── Histogram ──
  // Show 0-bin as a note, scale bars from 1+ bins for readability
  const zeroBin = result.bins.find(b => b.label === '0');
  const nonZeroBins = result.bins.filter(b => b.label !== '0');
  const maxBinCount = Math.max(...nonZeroBins.map(b => b.count), 1);
  const histHtml = nonZeroBins.map(b => {
    const pct = result.totalSlots > 0 ? (b.count / result.totalSlots * 100) : 0;
    const barW = (b.count / maxBinCount) * 100;
    return `<div class="st-hist-row" onclick="showHistogramDetail('${b.label}')" style="cursor:pointer" title="Click to see example time slots">
      <span class="st-hist-label">${b.label}</span>
      <div class="st-hist-bar-wrap"><div class="st-hist-bar" style="width:${Math.max(barW, 1)}%"></div></div>
      <span class="st-hist-pct">${pct < 0.1 ? '<1' : pct.toFixed(0)}%</span>
      <span style="font-size:0.72rem;color:var(--muted);width:36px;flex-shrink:0">(${b.count})</span>
    </div>`;
  }).join('');
  const zeroPct = zeroBin ? (zeroBin.count / result.totalSlots * 100).toFixed(1) : '0';
  document.getElementById('stHistogram').innerHTML = histHtml +
    `<div class="st-hist-note">
      ${zeroPct}% of time slots have <strong>0</strong> staff (${zeroBin ? zeroBin.count.toLocaleString() : '0'} of ${result.totalSlots.toLocaleString()} empty)
      &nbsp;&middot;&nbsp; Chart bars scaled relative to the busiest non-empty bin
    </div>`;

  // ── Week peaks ──
  const maxWeekPeak = Math.max(...result.weekPeaks.map(w => w.peak), 1);
  const weekHtml = result.weekPeaks.map(w => {
    const barW = (w.peak / maxWeekPeak) * 100;
    const barColor = w.peak === result.peak
      ? 'linear-gradient(90deg,var(--rust),var(--gold))'
      : w.peak >= maxWeekPeak * 0.5
        ? 'linear-gradient(90deg,var(--mid-blue),var(--accent))'
        : 'var(--mid-blue)';
    return `<div class="st-week-row" onclick="showWeekPeakDetail(${w.week})" style="cursor:pointer" title="Click to see peak detail for this week">
      <span class="st-week-label">W${w.week}</span>
      <div class="st-week-bar-wrap"><div class="st-week-bar" style="width:${barW}%;background:${barColor}"></div></div>
      <span class="st-week-val">${w.peak}</span>
    </div>`;
  }).join('');
  document.getElementById('stWeekPeaks').innerHTML = weekHtml;

  // ── High-demand sessions table ──
  if (result.highDemand.length > 0) {
    const weekStr = s => s.weeksRaw || (s.weeks ? 'W' + s.weeks.join(',') : '');
    const hdHtml = `<table class="st-hd-table">
      <thead><tr>
        <th>#</th>
        <th>Module</th>
        <th>Session</th>
        <th>Type</th>
        <th>Day</th>
        <th>Time</th>
        <th>Weeks</th>
        <th class="num">Staff</th>
        <th>Location</th>
      </tr></thead>
      <tbody>
      ${result.highDemand.map((s, i) => `<tr>
        <td>${i + 1}</td>
        <td>${s.moduleCode || s.moduleTitle || '—'}</td>
        <td>${s.sessionTitle || s.activity || '—'}</td>
        <td>${s.type || '—'}</td>
        <td>${s.day || '—'}</td>
        <td>${s.start && s.end ? s.start + '–' + s.end : '—'}</td>
        <td>${weekStr(s)}</td>
        <td class="num"><strong>${s.staff.length}</strong></td>
        <td>${s.location || '—'}</td>
      </tr>`).join('')}
      </tbody>
    </table>`;
    document.getElementById('stHighDemand').innerHTML = hdHtml;
  } else {
    document.getElementById('stHighDemand').innerHTML =
      '<div style="color:var(--muted);font-size:0.85rem;padding:0.5rem 0">No sessions with 5+ staff.</div>';
  }

  // ── Term-time section ──
  renderTermTimeSection();
}

// ── Term-time staffing analysis ─────────────────────

let stTermResult = null;

function computeTermTimeStaffing(maxIntensity) {
  stTermResult = null;
  if (!tlParsedSessions || tlParsedSessions.length === 0) return null;

  const teachingWeeks = tlAllWeeks; // already filtered by week range
  if (teachingWeeks.length === 0) return null;

  const firstWeek = teachingWeeks[0];
  const lastWeek = teachingWeeks[teachingWeeks.length - 1];
  const numWeeks = teachingWeeks.length;

  // Per-week data
  let totalTeachingHours = 0;
  const weekly = {};

  for (const w of teachingWeeks) {
    let weekTotal = 0;
    const staffHours = {};

    for (const name of tlAllStaff) {
      const sessions = tlStaffData[name]?.[w];
      if (!sessions || sessions.length === 0) continue;
      const h = calcHours(sessions, tlRealisticMode);
      if (h > 0) {
        staffHours[name] = h;
        weekTotal += h;
      }
    }

    weekly[w] = {
      totalHours: weekTotal,
      staffHours,
      staffCount: Object.keys(staffHours).length,
      requiredStaff: Math.ceil(weekTotal / maxIntensity),
    };
    totalTeachingHours += weekTotal;
  }

  // Find peak week (by required staff)
  const peakWeek = [...teachingWeeks].sort(
    (a, b) => weekly[b].requiredStaff - weekly[a].requiredStaff
  )[0];

  // Average weekly hours
  const avgWeeklyHours = totalTeachingHours / numWeeks;

  stTermResult = {
    teachingWeeks, firstWeek, lastWeek, numWeeks,
    totalTeachingHours, avgWeeklyHours,
    weekly, peakWeek, maxIntensity,
  };

  return stTermResult;
}

function renderTermTimeSection() {
  const intensityEl = document.getElementById('stMaxIntensity');
  const maxIntensity = +(intensityEl?.value || 14);

  const result = computeTermTimeStaffing(maxIntensity);

  // Stats row
  const statsEl = document.getElementById('stTermStats');
  if (!result) {
    statsEl.innerHTML = '<div style="color:var(--muted);font-size:0.85rem;padding:0.5rem 0">No timetable data loaded.</div>';
    document.getElementById('stTermWeekly').innerHTML = '';
    document.getElementById('stTermIntensity').innerHTML = '';
    return;
  }

  // ── Stats cards ──
  const peakReq = result.weekly[result.peakWeek].requiredStaff;

  statsEl.innerHTML =
    `<div class="stat-card" style="border-left:3px solid var(--purple)">
      <div class="sc-v">${result.firstWeek}–${result.lastWeek}</div>
      <div class="sc-l">Teaching window (${result.numWeeks} weeks)</div>
    </div>
    <div class="stat-card" style="border-left:3px solid var(--mid-blue)">
      <div class="sc-v">${result.totalTeachingHours.toFixed(0)}</div>
      <div class="sc-l">Total teaching hours (contact)</div>
    </div>
    <div class="stat-card" style="border-left:3px solid var(--gold)">
      <div class="sc-v">${result.avgWeeklyHours.toFixed(0)}</div>
      <div class="sc-l">Avg teaching hours/week</div>
    </div>
    <div class="stat-card" style="border-left:3px solid var(--rust)">
      <div class="sc-v">${peakReq}</div>
      <div class="sc-l">Min staff needed (peak W${result.peakWeek}, ${maxIntensity}h/wk cap)</div>
    </div>
    <div class="stat-card" style="border-left:3px solid var(--teal)">
      <div class="sc-v">${tlAllStaff.length}</div>
      <div class="sc-l">Actual staff in timetable</div>
    </div>`;

  // ── Weekly table ──
  const maxWeekHours = Math.max(...result.teachingWeeks.map(w => result.weekly[w].totalHours), 1);
  const maxStaffNeeded = Math.max(...result.teachingWeeks.map(w => result.weekly[w].requiredStaff), 1);

  const weekRows = result.teachingWeeks.map(w => {
    const d = result.weekly[w];
    const gap = d.staffCount - d.requiredStaff;
    const gapCls = gap < 0 ? 'deficit' : gap > 0 ? 'surplus' : '';
    const gapLabel = gap < 0 ? `${gap}` : `+${gap}`;
    const isPeak = w === result.peakWeek;
    const barW = (d.totalHours / maxWeekHours) * 100;
    const staffBarW = (d.staffCount / (maxStaffNeeded * 1.3)) * 100;

    return `<tr class="${isPeak ? 'peak-row' : ''}" onclick="showWeekStaffDetail(${w})" style="cursor:pointer" title="Click to see staff breakdown for this week">
      <td>W${w}</td>
      <td class="num">${d.totalHours.toFixed(0)}</td>
      <td>
        <div class="bar-wrap">
          <div class="bar" style="width:${Math.max(barW, 1)}%"></div>
          <span>${d.totalHours.toFixed(0)}h</span>
        </div>
      </td>
      <td class="num">${d.requiredStaff}</td>
      <td class="num">${d.staffCount}</td>
      <td class="num">
        <div class="bar-wrap">
          <div class="bar bar-staff" style="width:${Math.max(staffBarW, 1)}%"></div>
          <span>${d.staffCount}</span>
        </div>
      </td>
      <td class="num ${gapCls}">${gapLabel}</td>
    </tr>`;
  }).join('');

  document.getElementById('stTermWeekly').innerHTML =
    `<div class="st-term-weekly">
      <table class="st-tw-table">
        <thead><tr>
          <th>Week</th>
          <th class="num">Teaching (h)</th>
          <th style="min-width:120px">Total hours</th>
          <th class="num">Min staff</th>
          <th class="num">Actual staff</th>
          <th style="min-width:100px">Actual staff</th>
          <th class="num">Gap</th>
        </tr></thead>
        <tbody>${weekRows}</tbody>
      </table>
      <div class="st-term-note">
        Min staff = weekly teaching hours ÷ ${maxIntensity}h/person, rounded up.
        Gap = actual staff − min staff (negative = potential shortfall).
        Highlighted row = peak week (highest min staff requirement).
      </div>
    </div>`;

  // ── Intensity distribution for peak week ──
  const peakData = result.weekly[result.peakWeek];
  const hours = Object.values(peakData.staffHours);
  const buckets = [
    { label: '0h',   min: 0,    max: 0,    count: 0 },
    { label: '1–5h', min: 0.5,  max: 5,    count: 0 },
    { label: '6–10h',min: 5.1,  max: 10,   count: 0 },
    { label: '10h+', min: 10.1, max: Infinity, count: 0 },
  ];

  // Count staff NOT teaching this week
  let staffNotTeaching = tlAllStaff.length - Object.keys(peakData.staffHours).length;
  buckets[0].count = staffNotTeaching;
  for (const h of hours) {
    for (const b of buckets) {
      if (h >= b.min && h <= b.max) { b.count++; break; }
    }
  }

  const maxBucket = Math.max(...buckets.map(b => b.count), 1);
  const tiHtml = buckets.map(b => {
    const barW = (b.count / maxBucket) * 100;
    const pctLabel = b.label === '0h' ? 'not teaching' : '';
    return `<div class="st-ti-row" onclick="showIntensityDetail(${result.peakWeek},'${b.label}',${b.min},${b.max})" style="cursor:pointer" title="Click to see staff list">
      <span class="st-ti-label">${b.label}</span>
      <div class="st-ti-bar-wrap">
        <div class="st-ti-bar" style="width:${Math.max(barW, 2)}%;background:${b.label === '10h+' ? 'var(--rust)' : 'var(--mid-blue)'}"></div>
      </div>
      <span class="st-ti-count">${b.count}</span>
      <span style="font-size:0.7rem;color:var(--muted)">${pctLabel}</span>
    </div>`;
  }).join('');

  document.getElementById('stTermIntensity').innerHTML =
    `<div class="st-term-intensity">
      <h4>Staff Intensity Distribution — Peak Week (W${result.peakWeek})</h4>
      <div style="font-size:0.78rem;color:var(--muted);margin-bottom:0.6rem">
        ${peakData.totalHours.toFixed(0)}h of teaching across ${peakData.staffCount} staff
        &nbsp;·&nbsp; Avg ${peakData.staffCount > 0 ? (peakData.totalHours / peakData.staffCount).toFixed(1) : '—'}h/person
        &nbsp;·&nbsp; ${staffNotTeaching} staff not teaching this week
      </div>
      ${tiHtml}
    </div>`;
}

// ── Deep-dive handlers ──────────────────────────────

function showWeekPeakDetail(week) {
  if (!staffingResult) return;
  const grid = staffingResult.grid;
  let peak = 0, peakKey = null;
  for (const [key, set] of grid) {
    if (key.startsWith(week + '|') && set.size > peak) { peak = set.size; peakKey = key; }
  }
  if (!peakKey) {
    openPanel('Week ' + week, 'No sessions',
      '<p style="color:var(--muted);padding:0.5rem 0">No teaching found for this week.</p>');
    return;
  }
  const [, day, minStr] = peakKey.split('|');
  const min = +minStr;
  const startH = formatHour(min / 60);
  const endH = formatHour((min + 30) / 60);
  const peakSessions = tlParsedSessions.filter(s =>
    s.weeks.includes(+week) && s.day === day &&
    timeToMinutes(s.start) < min + 30 && timeToMinutes(s.end) > min
  );
  const allStaff = [...new Set(peakSessions.flatMap(s => s.staff))].sort();
  const sessList = peakSessions.map(s =>
    `<div class="session-card">
      <div class="sc-title">${s.type || 'Session'} <span class="sc-hours">${s.staff.length} staff</span></div>
      <div class="sc-meta">
        ${s.moduleCode ? '<span class="sc-tag">' + s.moduleCode + '</span>' : ''}
        ${s.sessionTitle || s.activity || ''}
        ${s.location ? ' <span>📍 ' + s.location + '</span>' : ''}
      </div>
    </div>`
  ).join('');
  openPanel('Week ' + week + ' Peak',
    peak + ' staff · ' + day + ' ' + startH + '–' + endH,
    '<div class="panel-section"><h4>Sessions (' + peakSessions.length + ')</h4>' + sessList + '</div>' +
    '<div class="panel-section" style="margin-top:0.8rem"><h4>Staff (' + allStaff.length + ')</h4>' +
    '<p style="font-size:0.82rem;line-height:1.7">' + allStaff.join(', ') + '</p></div>');
}

function showHistogramDetail(label) {
  if (!staffingResult) return;
  const bin = staffingResult.bins.find(b => b.label === label);
  if (!bin) return;
  const matching = [...staffingResult.grid.entries()]
    .filter(([, set]) => set.size >= bin.min && set.size <= bin.max)
    .sort((a, b) => b[1].size - a[1].size)
    .slice(0, 30);
  if (matching.length === 0) {
    openPanel('Concurrency: ' + label + ' staff', 'No matching slots',
      '<p style="color:var(--muted);padding:0.5rem 0">No time slots in this range.</p>');
    return;
  }
  const rows = matching.map(([key, set]) => {
    const [w, d, m] = key.split('|');
    const startH = formatHour(+m / 60);
    const endH = formatHour((+m + 30) / 60);
    const sessions = tlParsedSessions.filter(s =>
      s.weeks.includes(+w) && s.day === d &&
      timeToMinutes(s.start) < +m + 30 && timeToMinutes(s.end) > +m
    );
    const types = [...new Set(sessions.map(s => s.type || '?'))].filter(Boolean).join(', ');
    return '<div class="session-card" style="margin-bottom:4px">' +
      '<div class="sc-title" style="font-size:0.8rem">W' + w + ' ' + d + ' ' + startH + '–' + endH +
        ' <span class="sc-hours">' + set.size + ' staff</span></div>' +
      '<div class="sc-meta">' + types + '</div></div>';
  }).join('');
  const pct = (matching.length / staffingResult.totalSlots * 100).toFixed(1);
  openPanel('Concurrency: ' + label + ' staff',
    matching.length + ' of ' + staffingResult.totalSlots.toLocaleString() + ' slots (' + pct + '%)',
    '<div class="panel-section"><h4>Example Slots (up to 30)</h4>' + rows + '</div>');
}

function showIntensityDetail(week, label, min, max) {
  const termResult = stTermResult;
  if (!termResult) return;
  const weekData = termResult.weekly[week];
  if (!weekData) return;
  const isZero = label === '0h';
  let items = [];
  if (isZero) {
    items = tlAllStaff.filter(n => !weekData.staffHours[n]);
  } else {
    items = Object.entries(weekData.staffHours)
      .filter(([, h]) => h >= min && h <= max)
      .sort((a, b) => b[1] - a[1]);
  }
  if (items.length === 0) {
    openPanel('Week ' + week + ': ' + label, '0 staff',
      '<p style="color:var(--muted);padding:0.5rem 0">No staff in this category.</p>');
    return;
  }
  const listHtml = items.map(n => {
    const name = Array.isArray(n) ? n[0] : n;
    const hrs = Array.isArray(n) ? ' (' + n[1].toFixed(1) + 'h)' : '';
    return '<div style="padding:4px 0;border-bottom:1px solid var(--border);font-size:0.82rem">' + name + hrs + '</div>';
  }).join('');
  openPanel('Week ' + week + ': ' + label,
    items.length + ' staff member' + (items.length > 1 ? 's' : ''),
    '<div class="panel-section">' + listHtml + '</div>');
}

function showWeekStaffDetail(week) {
  const termResult = stTermResult;
  if (!termResult) return;
  const weekData = termResult.weekly[week];
  if (!weekData) return;

  const sorted = Object.entries(weekData.staffHours)
    .sort((a, b) => b[1] - a[1]);

  const maxH = sorted.length > 0 ? sorted[0][1] : 1;
  const listHtml = sorted.map(([name, h]) =>
    '<div style="display:flex;align-items:center;gap:10px;padding:5px 0;border-bottom:1px solid var(--border)">' +
      '<span style="flex:1;font-size:0.85rem">' + name + '</span>' +
      '<div style="width:100px;height:8px;background:var(--light-blue);border-radius:4px;overflow:hidden">' +
        '<div style="height:100%;width:' + (h / maxH * 100) + '%;background:var(--mid-blue);border-radius:4px"></div>' +
      '</div>' +
      '<span style="font-family:monospace;font-size:0.85rem;width:50px;text-align:right;font-weight:600">' + h.toFixed(1) + 'h</span>' +
    '</div>'
  ).join('');

  openPanel(
    'Week ' + week + ' — Staff Teaching Hours',
    sorted.length + ' staff teaching · ' + weekData.totalHours.toFixed(0) + 'h total',
    '<div class="panel-section">' + listHtml + '</div>' +
    '<div class="panel-section" style="margin-top:0.8rem">' +
    '<h4>Staff Not Teaching This Week (' + (tlAllStaff.length - sorted.length) + ')</h4>' +
    '<p style="font-size:0.82rem;color:var(--muted)">' + tlAllStaff.filter(n => !weekData.staffHours[n]).join(', ') + '</p></div>'
  );
}

// ── Wire up ──────────────────────────────────────────

// Run when the staffing subtab is clicked
document.addEventListener('DOMContentLoaded', () => {
  const stBtn = document.querySelector('.sub-tab-btn[data-panel="panel-staffing"]');
  if (stBtn) {
    stBtn.addEventListener('click', () => {
      // render on next tick so the panel is visible first
      setTimeout(renderStaffingView, 50);
    });
  }

  // Re-render term-time section when max intensity changes
  const intensityEl = document.getElementById('stMaxIntensity');
  if (intensityEl) {
    intensityEl.addEventListener('input', () => {
      // Only re-render the term-time portion if the results panel is visible
      const results = document.getElementById('st-results');
      if (results && results.style.display !== 'none') {
        renderTermTimeSection();
      }
    });
  }
});
