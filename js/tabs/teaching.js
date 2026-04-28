// ═══════════════════════════════════════════════════════
// TAB 1 — TEACHING LOAD
// ═══════════════════════════════════════════════════════
let tlUploadedFiles=[],tlParsedSessions=[],tlWeekRange=[1,52],tlRealisticMode=false;
let tlStaffData={},tlModuleData={},tlTypeData={};
let tlAllWeeks=[],tlAllStaff=[],tlAllModules=[],tlAllTypes=[];
let tlStaffSort={col:'total',dir:-1},tlModSort={col:'total',dir:-1},tlTypeSort={col:'total',dir:-1};
let tlPrepRatio=0;

const tlFileInput=document.getElementById('tlFileInput'),tlDropZone=document.getElementById('tlDropZone'),tlFileList=document.getElementById('tlFileList'),tlAnalyseBtn=document.getElementById('tlAnalyseBtn');

function tlUpdateFileList(){
  tlFileList.innerHTML='';
  tlUploadedFiles.forEach((f,i)=>{
    const d=document.createElement('div');d.className='file-item';
    d.innerHTML=`<span>📄</span><span class="fi-name">${f.name}</span><button class="fi-remove" data-i="${i}">✕</button>`;
    tlFileList.appendChild(d);
  });
  tlAnalyseBtn.disabled=tlUploadedFiles.length===0;
}

tlFileInput.addEventListener('change',e=>{tlUploadedFiles.push(...Array.from(e.target.files));tlUpdateFileList();tlFileInput.value='';});
tlDropZone.addEventListener('dragover',e=>{e.preventDefault();tlDropZone.classList.add('drag-over');});
tlDropZone.addEventListener('dragleave',()=>tlDropZone.classList.remove('drag-over'));
tlDropZone.addEventListener('drop',e=>{e.preventDefault();tlDropZone.classList.remove('drag-over');tlUploadedFiles.push(...Array.from(e.dataTransfer.files).filter(f=>f.name.match(/\.html?$/i)));tlUpdateFileList();});
tlFileList.addEventListener('click',e=>{if(e.target.classList.contains('fi-remove')){tlUploadedFiles.splice(+e.target.dataset.i,1);tlUpdateFileList();}});

function parseHTMLFile(html){
  const doc=new DOMParser().parseFromString(html,'text/html'),sessions=[];
  for(const table of doc.querySelectorAll('table')){
    const rows=table.querySelectorAll('tr');if(rows.length<2)continue;
    let headerRow=null,headerIdx=0;
    for(let i=0;i<Math.min(rows.length,5);i++){
      const cells=rows[i].querySelectorAll('th,td'),texts=[...cells].map(c=>c.textContent.trim().toLowerCase());
      if(texts.some(t=>t.includes('module')||t.includes('activity')||t.includes('staff'))){headerRow=cells;headerIdx=i;break;}
    }
    if(!headerRow)continue;
    const cols={},headerMap={
      activity:['activity'],moduleCode:['module code','module_code','modulecode'],moduleTitle:['module title','module_title','moduletitle'],
      sessionTitle:['session title','session_title','sessiontitle'],type:['type'],weeks:['weeks','week'],day:['day'],
      start:['start'],end:['end'],staff:['staff'],location:['location','room'],groupInfo:['group information','group info','group'],notes:['notes']
    };
    [...headerRow].forEach((cell,i)=>{
      const t=cell.textContent.trim().toLowerCase();
      for(const[key,variants]of Object.entries(headerMap)){
        if(variants.some(v=>t.includes(v))&&cols[key]===undefined)cols[key]=i;
      }
    });
    if(cols.weeks===undefined&&cols.staff===undefined)continue;
    for(let i=headerIdx+1;i<rows.length;i++){
      const cells=rows[i].querySelectorAll('td,th');if(cells.length<2)continue;
      const get=k=>(cols[k]!==undefined&&cells[cols[k]])?cells[cols[k]].innerHTML:'';
      const getText=k=>(cols[k]!==undefined&&cells[cols[k]])?cells[cols[k]].textContent.trim():'';
      const weeks=parseWeeks(getText('weeks'));if(weeks.length===0)continue;
      const staffNames=get('staff').split(/<br\s*\/?>/gi).map(s=>s.replace(/<[^>]+>/g,'').trim()).filter(Boolean);
      if(staffNames.length===0)continue;
      sessions.push({
        activity:getText('activity'),moduleCode:getText('moduleCode'),moduleTitle:getText('moduleTitle'),
        sessionTitle:getText('sessionTitle'),type:getText('type'),weeks,weeksRaw:getText('weeks'),
        day:getText('day'),start:getText('start'),end:getText('end'),staff:staffNames,
        location:getText('location'),groupInfo:getText('groupInfo'),notes:getText('notes')
      });
    }
  }
  return sessions;
}

function tlAggregate(sessions,wFrom,wTo){
  tlStaffData={};tlModuleData={};tlTypeData={};
  for(const sess of sessions){
    const fw=sess.weeks.filter(w=>w>=wFrom&&w<=wTo);if(fw.length===0)continue;
    for(const name of sess.staff){
      if(!tlStaffData[name])tlStaffData[name]={};
      for(const w of fw){if(!tlStaffData[name][w])tlStaffData[name][w]=[];tlStaffData[name][w].push(sess);}
      const mk=sess.moduleCode||sess.moduleTitle||'Unknown';
      if(!tlModuleData[mk])tlModuleData[mk]={};
      for(const w of fw){
        if(!tlModuleData[mk][w]){const _arr=[];Object.defineProperty(_arr,'staffMap',{value:new Map(),enumerable:false});tlModuleData[mk][w]=_arr;}
        tlModuleData[mk][w].push(sess);
        if(!tlModuleData[mk][w].staffMap.has(name))tlModuleData[mk][w].staffMap.set(name,[]);
        tlModuleData[mk][w].staffMap.get(name).push(sess);
      }
      const tk=(sess.type||'').trim()||'Undefined';
      if(!tlTypeData[tk])tlTypeData[tk]={};
      for(const w of fw){
        if(!tlTypeData[tk][w]){const _arr=[];Object.defineProperty(_arr,'staffMap',{value:new Map(),enumerable:false});tlTypeData[tk][w]=_arr;}
        tlTypeData[tk][w].push(sess);
        if(!tlTypeData[tk][w].staffMap.has(name))tlTypeData[tk][w].staffMap.set(name,[]);
        tlTypeData[tk][w].staffMap.get(name).push(sess);
      }
    }
  }
  tlAllStaff=Object.keys(tlStaffData).sort();
  tlAllModules=Object.keys(tlModuleData).sort();
  tlAllTypes=Object.keys(tlTypeData).sort((a,b)=>a==='Undefined'?1:b==='Undefined'?-1:a.localeCompare(b));
  const ws=new Set();
  for(const s of sessions)for(const w of s.weeks)if(w>=wFrom&&w<=wTo)ws.add(w);
  tlAllWeeks=[...ws].sort((a,b)=>a-b);
}

function buildGrid(dataMap,entities,weeks,sortConfig,realistic,prepRatio){
  const pr=prepRatio||0;
  if(entities.length===0)return{html:'<div style="padding:2rem;color:var(--muted)">No data found.</div>',legendHtml:''};
  const sorted=[...entities];
  if(sortConfig.col==='name')sorted.sort((a,b)=>sortConfig.dir*a.localeCompare(b));
  else if(sortConfig.col==='total')sorted.sort((a,b)=>sortConfig.dir*(weeks.reduce((s,w)=>s+calcHours(dataMap[a]?.[w],realistic)*(1+pr),0)-weeks.reduce((s,w)=>s+calcHours(dataMap[b]?.[w],realistic)*(1+pr),0)));
  else sorted.sort((a,b)=>sortConfig.dir*(calcHours(dataMap[a]?.[sortConfig.col],realistic)*(1+pr)-calcHours(dataMap[b]?.[sortConfig.col],realistic)*(1+pr)));
  let maxH=0;
  for(const e of entities)for(const w of weeks){const h=calcHours(dataMap[e]?.[w],realistic)*(1+pr);if(h>maxH)maxH=h;}
  const breaks=[0,maxH*0.1,maxH*0.25,maxH*0.45,maxH*0.65,maxH*0.85];
  const heatClass=h=>{if(h<=0)return'';for(let i=breaks.length-1;i>=0;i--)if(h>=breaks[i])return`heat-${i}`;return'heat-0';};
  const totArr=sortConfig.col==='total'?(sortConfig.dir>0?' ↑':' ↓'):'';
  const nameArr=sortConfig.col==='name'?(sortConfig.dir>0?' ↑':' ↓'):'';
  let html=`<table class="grid-table"><thead><tr><th style="text-align:left" data-sort="name">Name${nameArr}</th>`;
  for(const w of weeks){const arr=sortConfig.col===w?(sortConfig.dir>0?' ↑':' ↓'):'';html+=`<th class="week-header" data-sort-week="${w}">W${w}${arr}</th>`;}
  html+=`<th class="week-header" data-sort="total" style="background:#0a3060">Total${totArr}</th></tr></thead><tbody>`;
  for(const entity of sorted){
    html+=`<tr><td class="name-cell" data-entity="${encodeURIComponent(entity)}" title="${entity}">${entity}</td>`;
    let rowTotal=0;
    for(const w of weeks){
      const contact=calcHours(dataMap[entity]?.[w],realistic);const h=contact*(1+pr);rowTotal+=h;
      if(h>0){const tip=pr>0?`title="${contact.toFixed(1)}h contact + ${(contact*pr).toFixed(1)}h prep"`:'' ;html+=`<td class="data-cell ${heatClass(h)}" data-entity="${encodeURIComponent(entity)}" data-week="${w}" ${tip}><span>${h.toFixed(1)}</span><small>${pr>0?'incl. prep':'hrs'}</small></td>`;}
      else html+=`<td class="data-cell empty" data-entity="${encodeURIComponent(entity)}" data-week="${w}">–</td>`;
    }
    html+=`<td class="data-cell heat-3" data-entity="${encodeURIComponent(entity)}" data-week="total"><span>${rowTotal.toFixed(1)}</span><small>${pr>0?'incl. prep':'hrs'}</small></td></tr>`;
  }
  html+='<tr class="totals-row"><td class="name-cell" data-week="total" data-entity="all">Grand Total</td>';
  let gt=0;
  for(const w of weeks){const wt=entities.reduce((s,e)=>s+calcHours(dataMap[e]?.[w],realistic)*(1+pr),0);gt+=wt;html+=`<td class="data-cell" data-week="${w}" data-entity="all"><span>${wt.toFixed(1)}</span><small>hrs</small></td>`;}
  html+=`<td class="data-cell" data-week="total" data-entity="all"><span>${gt.toFixed(1)}</span><small>hrs</small></td></tr></tbody></table>`;
  const legendHtml=`<span>Colour scale:</span>${breaks.map((b,i)=>`<span class="legend-item"><span class="legend-swatch heat-${i}"></span>${b.toFixed(0)}${i<breaks.length-1?'–'+breaks[i+1].toFixed(0):'+'}</span>`).join('')}`;
  return{html,legendHtml};
}

function tlRenderGrid(id,dataMap,allEntities,weeks,sortCfg,realistic,legendId,prepRatio){
  const res=buildGrid(dataMap,allEntities,weeks,sortCfg,realistic,prepRatio||0);
  const wrap=document.getElementById(id);wrap.innerHTML=res.html;
  if(legendId)document.getElementById(legendId).innerHTML=res.legendHtml;
  tlAttachGridEvents(wrap,dataMap,id.includes('Staff')?'staff':id.includes('Mod')?'module':'type');
}

function tlRenderStaffGrid(){tlRenderGrid('tlStaffGridWrap',tlStaffData,tlAllStaff,tlAllWeeks,tlStaffSort,tlRealisticMode,'tlStaffLegend',tlPrepRatio);document.getElementById('tlStaffSortInfo').textContent=`Sorted by: ${tlStaffSort.col==='name'?'Name':tlStaffSort.col==='total'?'Total':'Week '+tlStaffSort.col} (${tlStaffSort.dir>0?'asc':'desc'})${tlPrepRatio>0?' · incl. '+tlPrepRatio+'× prep':''}`;}
function tlRenderModGrid(){tlRenderGrid('tlModGridWrap',tlModuleData,modTagFilteredModules(),tlAllWeeks,tlModSort,tlRealisticMode,'tlModLegend',tlPrepRatio);renderModuleTagChips();}
function tlRenderTypeGrid(){tlRenderGrid('tlTypeGridWrap',tlTypeData,tlAllTypes,tlAllWeeks,tlTypeSort,tlRealisticMode,'tlTypeLegend',tlPrepRatio);}

function tlAttachGridEvents(wrap,dataMap,type){
  wrap.querySelectorAll('.data-cell,.name-cell').forEach(cell=>{
    cell.addEventListener('click',()=>{
      const entity=cell.dataset.entity?decodeURIComponent(cell.dataset.entity):null,week=cell.dataset.week;
      if(!entity)return;
      const entities=entity==='all'?Object.keys(dataMap):[entity],weeks=(week==='total'||!week)?tlAllWeeks:[+week];
      const applyR=tlRealisticMode;
      let totalH=0,rawS=[];
      for(const e of entities)for(const w of weeks){totalH+=calcHours(dataMap[e]?.[w],applyR);rawS.push(...(dataMap[e]?.[w]||[]));}
      const unique=deduplicateSessions(rawS),realH=calcHours(unique,true);
      let html=`<div class="modal-stat-row"><div class="modal-stat"><div class="ms-value">${totalH.toFixed(1)}</div><div class="ms-label">Staff-hours</div></div><div class="modal-stat"><div class="ms-value">${realH.toFixed(1)}</div><div class="ms-label">Deduped hrs</div></div><div class="modal-stat"><div class="ms-value">${unique.length}</div><div class="ms-label">Sessions</div></div></div><div>`;
      const dayOrder={monday:1,tuesday:2,wednesday:3,thursday:4,friday:5,saturday:6,sunday:7};
      const sorted=unique.slice().sort((a,b)=>{
        const aWeek=Math.min(...a.weeks),bWeek=Math.min(...b.weeks);
        if(aWeek!==bWeek)return aWeek-bWeek;
        const aDay=dayOrder[(a.day||'').toLowerCase()]||99,bDay=dayOrder[(b.day||'').toLowerCase()]||99;
        if(aDay!==bDay)return aDay-bDay;
        const aStart=timeToHours(a.start),bStart=timeToHours(b.start);
        return aStart-bStart;
      });
      for(const s of sorted){
        const dur=sessionDuration(s);
        const weekStr=s.weeksRaw||(s.weeks.length?'W'+s.weeks.join(','):'');
        html+=`<div class="session-card"><div class="sc-title">${s.moduleCode?`<span class="sc-tag">${s.moduleCode}</span>`:''} ${s.moduleTitle||s.sessionTitle||s.activity||'Session'}<span class="sc-hours">${dur.toFixed(1)}h</span></div><div class="sc-meta">${s.day?`<span>📅 ${s.day}</span>`:''} ${weekStr?`<span>📆 ${weekStr}</span>`:''} ${s.start&&s.end?`<span>🕐 ${s.start}–${s.end}</span>`:''} ${s.location?`<span>📍 ${s.location}</span>`:''}</div></div>`;
      }
      html+='</div>';
      openPanel(entity==='all'?'Grand Total':entity,week==='total'?'All weeks':`Week ${week}`,html);
    });
  });
  const thead=wrap.querySelector('thead');
  if(thead)thead.addEventListener('click',e=>{
    const th=e.target.closest('th');if(!th)return;
    const col=th.dataset.sort,w=th.dataset.sortWeek?+th.dataset.sortWeek:null;
    const s=type==='staff'?tlStaffSort:type==='module'?tlModSort:tlTypeSort;
    const render=type==='staff'?tlRenderStaffGrid:type==='module'?tlRenderModGrid:tlRenderTypeGrid;
    if(col){if(s.col===col)s.dir*=-1;else{s.col=col;s.dir=-1;}render();}
    else if(w!==null){if(s.col===w)s.dir*=-1;else{s.col=w;s.dir=-1;}render();}
  });
}

function tlUpdateStatsBar(){
  const staffH=tlAllStaff.reduce((sum,staff)=>sum+tlAllWeeks.reduce((wSum,week)=>wSum+calcHours(tlStaffData[staff]?.[week],tlRealisticMode),0),0);
  const prepH=staffH*tlPrepRatio;
  const totalH=staffH+prepH;
  const cards=[['Sessions',tlParsedSessions.length],['Staff',tlAllStaff.length],['Modules',tlAllModules.length],['Weeks',tlAllWeeks.length],['Staff hrs',staffH.toFixed(0)]];
  if(tlPrepRatio>0){cards.push(['Prep hrs ('+tlPrepRatio+'×)',prepH.toFixed(0)]);cards.push(['Total hrs',totalH.toFixed(0)]);}
  document.getElementById('tlStatsBar').innerHTML=cards.map(([l,v])=>`<div class="stat-card"><div class="sc-v">${v}</div><div class="sc-l">${l}</div></div>`).join('');
}

tlAnalyseBtn.addEventListener('click',async()=>{
  tlAnalyseBtn.disabled=true;tlAnalyseBtn.textContent='⏳ Processing…';
  tlParsedSessions=[];
  for(const file of tlUploadedFiles){const html=await file.text();tlParsedSessions.push(...parseHTMLFile(html));}
  const wFrom=+document.getElementById('tlWeekFrom').value||1,wTo=+document.getElementById('tlWeekTo').value||52;
  tlWeekRange=[wFrom,wTo];tlRealisticMode=document.getElementById('tlRealistic').checked;
  document.getElementById('tlRealistic2').checked=tlRealisticMode;
  tlAggregate(tlParsedSessions,wFrom,wTo);
  document.getElementById('tlAnalyserTitle').textContent=document.getElementById('tlTitle').value||'Teaching Load';
  document.getElementById('tlAnalyserMeta').textContent=`${tlParsedSessions.length} sessions · ${tlAllStaff.length} staff · ${tlAllModules.length} modules · ${tlAllTypes.length} types`;
  tlUpdateStatsBar();
  document.getElementById('tl-landing').style.display='none';document.getElementById('tl-analyser').style.display='block';
  document.getElementById('badge-teaching').textContent=tlAllStaff.length+' staff';
  tlRenderStaffGrid();tlRenderModGrid();tlRenderTypeGrid();renderModuleTagFilterBar();updateCombStatus();
  tlAnalyseBtn.disabled=false;tlAnalyseBtn.textContent='📊 Analyse Timetable';
});

document.getElementById('tlRealistic2').addEventListener('change',e=>{tlRealisticMode=e.target.checked;tlUpdateStatsBar();tlRenderStaffGrid();tlRenderModGrid();tlRenderTypeGrid();});
document.getElementById('tlBtnSettings').addEventListener('click',()=>{document.getElementById('tlInlineSettings').classList.toggle('open');});
document.getElementById('tlRecalcBtn').addEventListener('click',()=>{const enabled=document.getElementById('tl_prep_enabled').checked;tlPrepRatio=enabled?(+document.getElementById('tl_prep_ratio').value||0):0;tlUpdateStatsBar();tlRenderStaffGrid();tlRenderModGrid();tlRenderTypeGrid();updateCombStatus();});
document.getElementById('tl_prep_enabled').addEventListener('change',()=>{document.getElementById('tl_prep_ratio').disabled=!document.getElementById('tl_prep_enabled').checked;});
document.querySelectorAll('[data-tltab]').forEach(btn=>{btn.addEventListener('click',()=>{document.querySelectorAll('[data-tltab]').forEach(b=>b.classList.remove('active'));btn.classList.add('active');const t=btn.dataset.tltab;['staff','modules','types'].forEach(x=>document.getElementById(`tl-tab-${x}`).style.display=x===t?'':'none');});});
document.getElementById('tlBtnBack').addEventListener('click',()=>{document.getElementById('tl-landing').style.display='';document.getElementById('tl-analyser').style.display='none';});
['tlSortStaffName','tlSortStaffTotal','tlSortModName','tlSortModTotal','tlSortTypeName','tlSortTypeTotal'].forEach(id=>{document.getElementById(id).addEventListener('click',()=>{const[,,type,col]=id.match(/tlSort(Staff|Mod|Type)(Name|Total)/),s=type==='Staff'?tlStaffSort:type==='Mod'?tlModSort:tlTypeSort,c=col==='Name'?'name':'total';if(s.col===c)s.dir*=-1;else{s.col=c;s.dir=col==='Name'?1:-1;}if(type==='Staff')tlRenderStaffGrid();else if(type==='Mod')tlRenderModGrid();else tlRenderTypeGrid();});});
document.getElementById('modTagFilterClear').addEventListener('click',()=>{moduleTagFilter=null;renderModuleTagFilterBar();tlRenderModGrid();});

document.getElementById('tlBtnExport').addEventListener('click',()=>{
  const wb=XLSX.utils.book_new();const pr=tlPrepRatio;
  const headers=pr>0?['Staff',...tlAllWeeks.map(w=>`Week ${w} Contact`),...tlAllWeeks.map(w=>`Week ${w} Prep`),'Total Contact','Total Prep','Total (incl. prep)']:['Staff',...tlAllWeeks.map(w=>`Week ${w}`),'Total'];
  const rows=[headers];
  for(const name of tlAllStaff){
    if(pr>0){const contacts=tlAllWeeks.map(w=>calcHours(tlStaffData[name]?.[w],tlRealisticMode));const totC=contacts.reduce((a,b)=>a+b,0);const totP=totC*pr;rows.push([name,...contacts,...contacts.map(c=>+(c*pr).toFixed(2)),+totC.toFixed(2),+totP.toFixed(2),+(totC+totP).toFixed(2)]);}
    else{const row=[name];let tot=0;for(const w of tlAllWeeks){const h=calcHours(tlStaffData[name]?.[w],tlRealisticMode);row.push(h||'');tot+=h;}row.push(+tot.toFixed(2));rows.push(row);}
  }
  if(pr>0){const contacts=tlAllWeeks.map(w=>tlAllStaff.reduce((s,e)=>s+calcHours(tlStaffData[e]?.[w],tlRealisticMode),0));const gtC=contacts.reduce((a,b)=>a+b,0);const gtP=gtC*pr;rows.push(['Grand Total',...contacts,...contacts.map(c=>+(c*pr).toFixed(2)),+gtC.toFixed(2),+gtP.toFixed(2),+(gtC+gtP).toFixed(2)]);}
  else{const gr=['Grand Total'];let gt=0;for(const w of tlAllWeeks){const wt=tlAllStaff.reduce((s,e)=>s+calcHours(tlStaffData[e]?.[w],tlRealisticMode),0);gr.push(+wt.toFixed(2));gt+=wt;}gr.push(+gt.toFixed(2));rows.push(gr);}
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(rows),'Staff Load');
  XLSX.writeFile(wb,'teaching_load.xlsx');
});
