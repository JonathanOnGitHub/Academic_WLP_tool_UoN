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

function parseMmiTabName(name){
  const s=name.trim();
  const dateM=s.match(/^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{2,4})/);
  let dateStr='';
  if(dateM){const d=dateM[1].padStart(2,'0'),mo=dateM[2].padStart(2,'0');let yr=dateM[3];if(yr.length===2)yr='20'+yr;dateStr=`${d}/${mo}/${yr}`;}

  let remainder=s.replace(/^\d{1,2}[.\/-]\d{1,2}[.\/-]\d{2,4}\s*/,'').replace(/^[-–—\s]+/,'').trim();

  function parseTime(raw, prevHour, pmHint){
    raw=String(raw).trim().toLowerCase();
    const isPm=raw.includes('pm');const isAm=raw.includes('am');
    raw=raw.replace(/[apm]/g,'').trim();
    let h,m=0;
    const parts=raw.split(/[.:]/);
    h=parseInt(parts[0],10);
    if(parts.length>1)m=parseInt(parts[1],10)||0;
    if(isNaN(h))return null;
    if(isPm&&h<12)h+=12;
    if(isAm&&h===12)h=0;
    if(!isPm&&!isAm&&pmHint&&h<12)h+=12;
    if(!isPm&&!isAm&&h<(prevHour||0)&&h>0)h+=12;
    return h+m/60;
  }

  const timeRangeM=remainder.match(/([\d.:]+\s*(?:am|pm)?)\s*[-–—]\s*([\d.:]+\s*(?:am|pm)?)\s*$/i);
  if(!timeRangeM) return{dateStr,label:remainder,startH:null,endH:null,durationH:null};

  const rawStart=timeRangeM[1].trim(),rawEnd=timeRangeM[2].trim();
  const label=remainder.slice(0,remainder.length-timeRangeM[0].length).replace(/[-–—\s]+$/,'').trim();
  const endHasPm=/pm/i.test(rawEnd);
  const startH=parseTime(rawStart,null,endHasPm&&!/am/i.test(rawStart));
  const endH=parseTime(rawEnd,startH,false);

  if(startH===null||endH===null||endH<=startH) return{dateStr,label,startH:null,endH:null,durationH:null};
  const durationH=Math.round((endH-startH)*100)/100;
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
        const info=parseMmiTabName(sheetName);
        if(info.durationH===null){unparsed.push(sheetName);continue;}
        parsed.push({sheetName,...info});
      }
      if(parsed.length===0){mmiShowError('No sheets with parseable time ranges found. Expected tab names like "18.03.26 -Online 1.45-4.45pm".');return;}
      if(unparsed.length>0)mmiShowWarn(`${unparsed.length} sheet${unparsed.length>1?'s':''} skipped (time not recognised): ${unparsed.slice(0,5).join(', ')}${unparsed.length>5?'…':''}`);
      document.getElementById('mmiPreview').style.display='block';
      document.getElementById('mmiPreviewList').innerHTML=parsed.map(p=>`<div style="padding:4px 6px;border-bottom:1px solid var(--border);display:flex;justify-content:space-between;align-items:center">
        <span style="font-weight:500">${p.sheetName}</span>
        <span style="color:var(--purple);font-family:'IBM Plex Mono',monospace;font-size:0.75rem">${p.dateStr} · ${formatHour(p.startH)}–${formatHour(p.endH)} · ${p.durationH.toFixed(2)}h</span>
      </div>`).join('');
      mmiAnalyseBtn.disabled=false;
      mmiAnalyseBtn.textContent=`🩺 Calculate MMI Workload (${parsed.length} session${parsed.length>1?'s':''}) →`;
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
    const info=parseMmiTabName(sheetName);if(info.durationH===null)continue;
    const ws=wb.Sheets[sheetName];
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
