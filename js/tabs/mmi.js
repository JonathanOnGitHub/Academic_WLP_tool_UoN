// ═══════════════════════════════════════════════════════
// TAB 4 — MMI WORKLOAD
// ═══════════════════════════════════════════════════════
let mmiSessions=[]; // [{sheetName, date, startH, endH, durationH, label, staff:[{name,isReserve}]}]
let mmiResults=[];  // [{name, sessions:[{...session, isReserve}], totalHours, isReserve(ever active)}]
let mmiRawWb=null;
let mmiSortKey='total-desc';

const mmiDropZone=document.getElementById('mmiDropZone'),mmiFileInput=document.getElementById('mmiFileInput'),mmiAnalyseBtn=document.getElementById('mmiAnalyseBtn');

mmiDropZone.addEventListener('dragover',e=>{e.preventDefault();mmiDropZone.classList.add('drag-over');});
mmiDropZone.addEventListener('dragleave',()=>mmiDropZone.classList.remove('drag-over'));
mmiDropZone.addEventListener('drop',e=>{e.preventDefault();mmiDropZone.classList.remove('drag-over');if(e.dataTransfer.files[0])mmiLoadFile(e.dataTransfer.files[0]);});
mmiFileInput.addEventListener('change',e=>{if(e.target.files[0])mmiLoadFile(e.target.files[0]);});

function mmiShowError(msg){const el=document.getElementById('mmiError');el.textContent=msg;el.classList.add('show');}
function mmiClearError(){document.getElementById('mmiError').classList.remove('show');}
function mmiShowWarn(msg){const el=document.getElementById('mmiWarn');el.textContent=msg;el.classList.add('show');}
function mmiClearWarn(){document.getElementById('mmiWarn').classList.remove('show');}

// Parse a single time string ("9.30am", "1.45", "12.30pm", "09.30") into
// fractional hours. The 2026-27 MMI source uses two unsuffixed formats:
//   * 24-hour style with a leading zero, e.g. "09.30" — always morning
//     when in the 0-9 range, always interpreted literally
//   * 12-hour style without a leading zero, e.g. "1.45" — disambiguated
//     by the previous time (legacy sheet-name parser) or by the end
//     time's am/pm suffix (title-based parser passes that as prevHour).
function mmiParseTime(raw, prevHour){
  raw=String(raw).trim().toLowerCase();
  const isPm=raw.includes('pm');const isAm=raw.includes('am');
  raw=raw.replace(/[apm]/g,'').trim();
  if(!raw)return null;
  const hStr=raw.split(/[.:]/)[0];
  let h=parseInt(hStr,10);
  let m=0;
  const parts=raw.split(/[.:]/);
  if(parts.length>1)m=parseInt(parts[1],10)||0;
  if(isNaN(h))return null;
  // Two-digit hours with a leading zero (e.g. "09", "12") are 24-hour
  // clock and never get the +12 rollover. This handles the
  // "09.30 - 12.30pm" case where the start must be 9:30am.
  const is24hStyle=hStr.length===2&&hStr.startsWith('0');
  if(isPm&&h<12)h+=12;
  if(isAm&&h===12)h=0;
  if(!isPm&&!isAm&&!is24hStyle&&prevHour&&h>0&&h<prevHour)h+=12;
  return h+m/60;
}

// Detect a leading-zero two-digit hour (e.g. "09", "12") — these are
// 24-hour-clock values and should not get the 12→24h rollover.
function is24hHourStr(raw){
  const m=String(raw).trim().match(/^(\d{1,2})/);
  if(!m)return false;
  const s=m[1];
  return s.length===2&&s.startsWith('0');
}

// Pull a date+time range from a sheet-title cell, the 2026-27 format. The
// title looks like one of:
//   "MMI session ONLINE Nov 25,  1.45 - 4.45pm"
//   "MMI session IN PERSON Dec 09,  1.45 - 4.45pm"
//   "MMI session ONLINE Dec 15,  09.30 - 12.30pm"
//   "MMI session INPERSON (UNIC/SOLS FY) March 24,  1.45 - 4.45pm"
//   "MMI session INPERSON (UNIC/SOLS FY) Jun 02 or 09,  1.45 - 4.45pm"
//   "MMI session ONLINE April 13,  9.30am - 12.30pm"        (am suffix on start)
//   "MMI session ONLINE April 14,  1.45- 4.45pm"            (missing space after -)
//
// Returns {dateStr, label, startH, endH, durationH} or null if the title
// doesn't look like an MMI session header.
function mmiParseTitle(title){
  if(!title)return null;
  const t=String(title).trim();
  if(!/mmi\s*session/i.test(t))return null;
  // Date: "Month DD[, or DD], " — capture month + first day. If there's
  // an "or DD" we keep the first date only (workload counted once).
  const MONTHS={jan:1,january:1,feb:2,february:2,mar:3,march:3,apr:4,april:4,
                may:5,jun:6,june:6,jul:7,july:7,aug:8,august:8,
                sep:9,sept:9,september:9,oct:10,october:10,nov:11,november:11,
                dec:12,december:12};
  const dateM=t.match(/\b(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|june?|july?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)\s+(\d{1,2})(?:\s+or\s+\d{1,2})?/i);
  if(!dateM)return null;
  const month=MONTHS[dateM[1].toLowerCase().slice(0,3)];
  const day=parseInt(dateM[2],10);
  if(!month||!day)return null;
  // Sheet names look like 25.11.26, 9.12.26, 6.01.27 — i.e. the year is the
  // last 2 digits of the academic year. We don't have a year in the title
  // itself, so derive it from the sheet name (caller passes it). For now
  // return null year and let the caller supply it.
  return { month, day, monthName:dateM[1] };
}

// Read date (year) from a sheet name like "25.11.26", "9.12.26", " 13.04.27 (am)",
// "02.06 OR 09.0.27" — return "DD/MM/YYYY" for the first date.
function mmiParseSheetDate(name){
  // Strip leading whitespace and an optional opening parenthesis
  const s=name.trim();
  // Find dd.mm(.yy)yy anywhere in the sheet name, take the first match.
  const m=s.match(/(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{2,4})/);
  if(!m)return null;
  const d=m[1].padStart(2,'0');
  const mo=m[2].padStart(2,'0');
  let yr=m[3];
  if(yr.length===2)yr='20'+yr;
  return `${d}/${mo}/${yr}`;
}

// Find a date+time range either from the sheet's row-2 title (2026-27
// format) or fall back to the sheet name (old format). Returns
// {dateStr,label,startH,endH,durationH} or null if neither yields a usable
// result.
function parseMmiTabName(name, ws){
  const s=name.trim();
  let dateStr='',label='',startH=null,endH=null,durationH=null;

  // 1. Try the row-2 title first (2026-27 format).
  if(ws){
    const rawRows=XLSX.utils.sheet_to_json(ws,{header:1,defval:'',raw:false,range:0});
    // Look in the first 4 rows for an "MMI session ..." line.
    let title=null;
    for(let i=0;i<Math.min(rawRows.length,4);i++){
      const cell=String(rawRows[i]?.[0]||'').trim();
      if(/mmi\s*session/i.test(cell)){title=cell;break;}
    }
    if(title){
      const parsed=mmiParseTitle(title);
      if(parsed){
        const sheetDate=mmiParseSheetDate(s);
        if(sheetDate){
          // Replace year portion in dd/mm/yyyy with the sheet-derived year.
          const yr=sheetDate.slice(6);
          const mm=String(parsed.month).padStart(2,'0');
          const dd=String(parsed.day).padStart(2,'0');
          dateStr=`${dd}/${mm}/${yr}`;
        }
        // Strip the leading "MMI session" + mode prefix to get a short label.
        label=title.replace(/^mmi\s*session\s+[^,]+,/i,'').trim();
        // Now grab the time range. Tolerate "9.30am - 12.30pm",
        // "1.45- 4.45pm", "09.30 - 12.30pm", etc. For unsuffixed starts
        // like "1.45" we use the end's hour as a 24h-style reference
        // point so the 12h→24h rollover kicks in (1.45 + 12 = 13.45);
        // for leading-zero 24h-style starts ("09.30") we pass null so
        // no rollover is applied.
        const timeRangeM=label.match(/([\d.:]+\s*(?:am|pm)?)\s*[-–—]\s*([\d.:]+\s*(?:am|pm)?)/i);
        if(timeRangeM){
          const rawStart=timeRangeM[1].trim(),rawEnd=timeRangeM[2].trim();
          const endHasPm=/pm/i.test(rawEnd);
          // End is parsed first without a hint; if it carries a pm suffix
          // and the unsuffixed start value is < 12, the start is PM too.
          // Pass end's raw hour (e.g. 12 for "12.30pm", 4 for "4.45pm") as
          // prevHour for the start. The 24h-style guard inside
          // mmiParseTime skips the rollover for "09.30"-shaped starts.
          const endHourRaw=parseInt(rawEnd.replace(/[^\d]/g,'').slice(0,2),10);
          const startHint=(endHasPm&&!is24hHourStr(rawStart))?endHourRaw:null;
          startH=mmiParseTime(rawStart,startHint);
          endH=mmiParseTime(rawEnd,startH);
          if(startH!==null&&endH!==null&&endH>startH){
            durationH=Math.round((endH-startH)*100)/100;
          }
          // Strip the time range from the label so what's left is just
          // the descriptive suffix (e.g. "INPERSON (UNIC/SOLS FY)").
          label=label.slice(0,timeRangeM.index).replace(/[,–—\-:\s]+$/,'').trim();
        }
      }
    }
  }

  // 2. Fall back to parsing the sheet name (legacy "DD.MM.YY -Online 1.45-4.45pm").
  if(durationH===null){
    const dateM=s.match(/^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{2,4})/);
    if(dateM){
      const d=dateM[1].padStart(2,'0'),mo=dateM[2].padStart(2,'0');
      let yr=dateM[3];if(yr.length===2)yr='20'+yr;
      dateStr=`${d}/${mo}/${yr}`;
    }
    let remainder=s.replace(/^\d{1,2}[.\/-]\d{1,2}[.\/-]\d{2,4}\s*/,'').replace(/^[-–—\s]+/,'').trim();
    const timeRangeM=remainder.match(/([\d.:]+\s*(?:am|pm)?)\s*[-–—]\s*([\d.:]+\s*(?:am|pm)?)\s*$/i);
    if(timeRangeM){
      const rawStart=timeRangeM[1].trim(),rawEnd=timeRangeM[2].trim();
      label=remainder.slice(0,remainder.length-timeRangeM[0].length).replace(/[-–—\s]+$/,'').trim();
      const endHasPm=/pm/i.test(rawEnd);
      startH=mmiParseTime(rawStart,endHasPm?12:null);
      endH=mmiParseTime(rawEnd,startH);
      if(startH!==null&&endH!==null&&endH>startH){
        durationH=Math.round((endH-startH)*100)/100;
      }
    }
  }

  if(durationH===null)return{dateStr,label,startH:null,endH:null,durationH:null};
  return{dateStr,label,startH,endH,durationH};
}

function mmiLoadFile(file){
  mmiClearError();mmiClearWarn();mmiAnalyseBtn.disabled=true;
  const reader=new FileReader();
  reader.onload=e=>{
    try{
      const wb=XLSX.read(e.target.result,{type:'array'});
      mmiRawWb=wb;
      const parsed=[];const unparsed=[];
      const visibleSheets=wb.SheetNames.filter((_,i)=>{const s=wb.Workbook?.Sheets?.[i];return !s||s.Hidden===0||s.Hidden===undefined;});
      for(const sheetName of visibleSheets){
        const ws=wb.Sheets[sheetName];
        const info=parseMmiTabName(sheetName,ws);
        if(info.durationH===null){unparsed.push(sheetName);continue;}
        parsed.push({sheetName,...info});
      }
      if(parsed.length===0){mmiShowError('No sheets with parseable time ranges found. Expected a row-2 title like "MMI session ONLINE Nov 25, 1.45 - 4.45pm" or tab names like "18.03.26 -Online 1.45-4.45pm".');return;}
      if(unparsed.length>0)mmiShowWarn(`${unparsed.length} sheet${unparsed.length>1?'s':''} skipped (time not recognised): ${unparsed.slice(0,5).join(', ')}${unparsed.length>5?'…':''}`);
      document.getElementById('mmiPreview').style.display='block';
      document.getElementById('mmiPreviewList').innerHTML=parsed.map(p=>`<div style="padding:4px 6px;border-bottom:1px solid var(--border);display:flex;justify-content:space-between;align-items:center">
        <span style="font-weight:500">${p.sheetName}</span>
        <span style="color:var(--purple);font-family:'IBM Plex Mono',monospace;font-size:0.75rem">${p.dateStr} · ${formatHour(p.startH)}–${formatHour(p.endH)} · ${p.durationH.toFixed(2)}h</span>
      </div>`).join('');
      mmiAnalyseBtn.disabled=false;
      mmiAnalyseBtn.textContent=`🩺 Calculate MMI Workload (${parsed.length} session${parsed.length>1?'s':''}) →`;
      WLP_SESSION.saveFile('mmi',file.name,e.target.result,file.type);
    }catch(err){mmiShowError('Error reading file: '+err.message);}
  };
  reader.readAsArrayBuffer(file);
}

function mmiAnalyse(){
  mmiClearError();
  const wb=mmiRawWb;if(!wb)return;
  const sessions=[];
  const visibleSheets=wb.SheetNames.filter((_,i)=>{const s=wb.Workbook?.Sheets?.[i];return!s||s.Hidden===0||s.Hidden===undefined;});
  for(const sheetName of visibleSheets){
    const ws=wb.Sheets[sheetName];
    const info=parseMmiTabName(sheetName,ws);if(info.durationH===null)continue;
    const raw=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});
    let reserveRowIdx=Infinity;
    for(let i=4;i<raw.length;i++){const cellA=String(raw[i][0]||'').trim().toLowerCase();if(cellA.includes('reserve')){reserveRowIdx=i;break;}}
    const staff=[];
    for(let i=4;i<raw.length;i++){
      const name=String(raw[i][1]||'').trim();if(!name)continue;
      if(/^(name|staff|reserve|assessor)/i.test(name))continue;
      staff.push({name,isReserve:i>=reserveRowIdx});
    }
    if(staff.length>0)sessions.push({sheetName,...info,staff});
  }
  if(sessions.length===0){mmiShowError('No staff names found in any session sheets. Ensure staff names are in column B from row 5.');return;}
  mmiSessions=sessions;
  const map={};
  for(const sess of sessions){for(const{name,isReserve}of sess.staff){if(!map[name])map[name]={name,sessions:[],totalHours:0,isActiveStaff:false};map[name].sessions.push({...sess,isReserve});if(!isReserve){map[name].totalHours+=sess.durationH;map[name].isActiveStaff=true;}}}
  mmiResults=Object.values(map).sort((a,b)=>b.totalHours-a.totalHours);
  const activeResults=mmiResults.filter(r=>r.isActiveStaff);
  const totalH=activeResults.reduce((s,r)=>s+r.totalHours,0);
  const totalSessions=sessions.length;
  document.getElementById('mmiMeta').textContent=`${totalSessions} sessions · ${activeResults.length} active staff · ${mmiResults.filter(r=>!r.isActiveStaff).length} reserve-only · ${totalH.toFixed(1)}h total`;
  document.getElementById('mmiStatsBar').innerHTML=[['Sessions',totalSessions],['Active Staff',activeResults.length],['Reserve Only',mmiResults.filter(r=>!r.isActiveStaff).length],['Total hrs',totalH.toFixed(1)],['Avg hrs',activeResults.length?fmt(totalH/activeResults.length):'—']].map(([l,v])=>`<div class="stat-card purple"><div class="sc-v">${v}</div><div class="sc-l">${l}</div></div>`).join('');
  document.getElementById('mmi-landing').style.display='none';document.getElementById('mmi-content').style.display='block';
  document.getElementById('badge-mmi').textContent=activeResults.length+' staff';
  mmiRenderTable();updateCombStatus();
}

mmiAnalyseBtn.addEventListener('click',mmiAnalyse);
document.getElementById('mmiBtnBack').addEventListener('click',()=>{document.getElementById('mmi-landing').style.display='';document.getElementById('mmi-content').style.display='none';});

function mmiGetSorted(){
  const q=document.getElementById('mmiSearch').value.toLowerCase();
  const showReserve=document.getElementById('mmiShowReserve').checked;
  let data=mmiResults.filter(r=>r.name.toLowerCase().includes(q));
  if(!showReserve)data=data.filter(r=>r.isActiveStaff);
  const[col,dir]=mmiSortKey.split('-');
  data.sort((a,b)=>{
    if(col==='name')return dir==='asc'?a.name.localeCompare(b.name):b.name.localeCompare(a.name);
    if(col==='sessions'){const an=a.sessions.filter(s=>!s.isReserve).length,bn=b.sessions.filter(s=>!s.isReserve).length;return dir==='desc'?bn-an:an-bn;}
    return dir==='desc'?b.totalHours-a.totalHours:a.totalHours-b.totalHours;
  });
  return data;
}

function mmiRenderTable(){
  const data=mmiGetSorted();
  const maxH=Math.max(...mmiResults.filter(r=>r.isActiveStaff).map(r=>r.totalHours),1);
  document.getElementById('mmiTbody').innerHTML=data.map(r=>{
    const activeSessions=r.sessions.filter(s=>!s.isReserve);
    const reserveSessions=r.sessions.filter(s=>s.isReserve);
    const isReserveOnly=!r.isActiveStaff;
    return`<tr style="cursor:pointer${isReserveOnly?';opacity:0.65':''}" data-name="${encodeURIComponent(r.name)}">
      <td class="name-f">${r.name}${isReserveOnly?'<span class="reserve-badge">reserve only</span>':''}</td>
      <td class="num">${activeSessions.length}${reserveSessions.length>0?`<span style="color:var(--muted);font-size:0.72rem"> (+${reserveSessions.length}R)</span>`:''}</td>
      <td class="tot">${isReserveOnly?'<span style="color:var(--muted)">—</span>':fmt(r.totalHours)+'h'}</td>
      <td>${isReserveOnly?'<span style="color:var(--muted);font-size:0.75rem">Reserve — hours not counted</span>':`<div class="hours-bar-wrap"><div class="hours-bar"><div class="hours-bar-fill" style="width:${r.totalHours/maxH*100}%;background:linear-gradient(90deg,var(--purple),#9333ea)"></div></div><span class="hours-val">${fmt(r.totalHours)}h</span></div>`}</td>
    </tr>`;
  }).join('');
  const activeData=mmiResults.filter(r=>r.isActiveStaff);
  const totH=activeData.reduce((s,r)=>s+r.totalHours,0);
  document.getElementById('mmiFoot').innerHTML=`<tr><td class="name-f">Grand Total (active staff)</td><td class="num">${mmiSessions.length} sessions</td><td class="tot" style="color:var(--purple)">${fmt(totH)}h</td><td></td></tr>`;
  document.querySelectorAll('#mmiTbody tr').forEach(row=>{row.addEventListener('click',()=>mmiOpenDetail(decodeURIComponent(row.dataset.name)));});
}

function mmiOpenDetail(name){
  const r=mmiResults.find(x=>x.name===name);if(!r)return;
  const activeSess=r.sessions.filter(s=>!s.isReserve);
  const reserveSess=r.sessions.filter(s=>s.isReserve);
  let html=`<div class="panel-section"><h4>Summary</h4>
    <div class="panel-row"><span class="k">Active sessions</span><span class="v">${activeSess.length}</span></div>
    <div class="panel-row"><span class="k">Reserve sessions</span><span class="v">${reserveSess.length}</span></div>
    <div class="panel-row"><span class="k">Total hours (active)</span><span class="v big">${fmt(r.totalHours)}h</span></div>
  </div>`;
  if(activeSess.length>0){html+=`<div class="panel-section"><h4>Active Sessions</h4>`;for(const s of activeSess)html+=`<div class="mmi-session-card"><div class="ms-title">${s.dateStr}${s.label?' · '+s.label:''}</div><div class="ms-meta">🕐 ${formatHour(s.startH)}–${formatHour(s.endH)} · ${s.durationH.toFixed(2)}h</div></div>`;html+='</div>';}
  if(reserveSess.length>0){html+=`<div class="panel-section"><h4>Reserve Sessions (hours not counted)</h4>`;for(const s of reserveSess)html+=`<div class="mmi-session-card" style="border-color:var(--muted);opacity:0.7"><div class="ms-title">${s.dateStr}${s.label?' · '+s.label:''}</div><div class="ms-meta">🕐 ${formatHour(s.startH)}–${formatHour(s.endH)} · ${s.durationH.toFixed(2)}h (reserve)</div></div>`;html+='</div>';}
  openPanel(name,`${fmt(r.totalHours)}h active · ${r.sessions.length} session${r.sessions.length!==1?'s':''}`,html);
}

document.getElementById('mmiSearch').addEventListener('input',mmiRenderTable);
document.getElementById('mmiSortSel').addEventListener('change',e=>{mmiSortKey=e.target.value;mmiRenderTable();});
document.getElementById('mmiShowReserve').addEventListener('change',mmiRenderTable);
document.querySelector('#mmi-content table.mmi-table thead').addEventListener('click',e=>{const th=e.target.closest('th[data-mmisort]');if(!th)return;const col=th.dataset.mmisort;const[curCol,curDir]=mmiSortKey.split('-');if(curCol===col)mmiSortKey=col+'-'+(curDir==='desc'?'asc':'desc');else mmiSortKey=col+'-'+(col==='name'?'asc':'desc');mmiRenderTable();});

document.getElementById('mmiBtnExport').addEventListener('click',()=>{
  const wb2=XLSX.utils.book_new();
  const rows=[['Academic','Active Sessions','Total Hours (active)','Is Reserve Only']];
  for(const r of mmiResults){const activeSess=r.sessions.filter(s=>!s.isReserve);rows.push([r.name,activeSess.length,+r.totalHours.toFixed(2),r.isActiveStaff?'No':'Yes']);}
  XLSX.utils.book_append_sheet(wb2,XLSX.utils.aoa_to_sheet(rows),'MMI Workload');
  const sRows=[['Session (Tab)','Date','Start','End','Duration (h)','Staff Name','Is Reserve']];
  for(const sess of mmiSessions)for(const st of sess.staff)sRows.push([sess.sheetName,sess.dateStr,formatHour(sess.startH),formatHour(sess.endH),+sess.durationH.toFixed(2),st.name,st.isReserve?'Yes':'No']);
  XLSX.utils.book_append_sheet(wb2,XLSX.utils.aoa_to_sheet(sRows),'Session Detail');
  XLSX.writeFile(wb2,'mmi_workload.xlsx');
});
