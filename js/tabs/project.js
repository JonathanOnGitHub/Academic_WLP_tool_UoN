// ═══════════════════════════════════════════════════════
// TAB 3 — PROJECT SUPERVISION
// ═══════════════════════════════════════════════════════
let projRawProjects=[],projAllResults=[],projSettings={supervision:12,cosupervision:6,diss_feedback:3,diss_marking:2,poster_feedback:0.5,poster_marking:1/3},projSortKey='total-desc';
const projDropZone=document.getElementById('projDropZone'),projFileInput=document.getElementById('projFileInput'),projAnalyseBtn=document.getElementById('projAnalyseBtn');

projDropZone.addEventListener('dragover',e=>{e.preventDefault();projDropZone.classList.add('drag-over');});
projDropZone.addEventListener('dragleave',()=>projDropZone.classList.remove('drag-over'));
projDropZone.addEventListener('drop',e=>{e.preventDefault();projDropZone.classList.remove('drag-over');if(e.dataTransfer.files[0])projLoadFile(e.dataTransfer.files[0]);});
projFileInput.addEventListener('change',e=>{if(e.target.files[0])projLoadFile(e.target.files[0]);});
document.getElementById('projSettingsHdr').addEventListener('click',()=>{document.getElementById('projSettingsBody').classList.toggle('open');document.getElementById('projSettingsHdr').classList.toggle('open');});

function projShowError(msg){const el=document.getElementById('projError');el.textContent=msg;el.classList.add('show');}
function projClearError(){document.getElementById('projError').classList.remove('show');}
function projNormH(s){return String(s).toLowerCase().replace(/[\s_\-]/g,'').replace(/[^a-z0-9]/g,'');}

const PROJ_COL_MAP={
  theme:['theme','projecttitle','title','project'],supervisor:['supervisor'],
  cosupervisors:['co_supervisors','cosupervisors','cosupervisor','cosup'],
  poster1:['poster_assessor_1','poster1assessor','poster_1_assessor','poster1','posterassessor1'],
  poster2:['poster_assessor_2','poster2assessor','poster_2_assessor','poster2','posterassessor2'],
  dissertation1:['dissertation_assessor_1','dissertation1assessor','dissertation_1_assessor','diss1','dissertationassessor1'],
  dissertation2:['dissertation_assessor_2','dissertation2assessor','dissertation_2_assessor','diss2','dissertationassessor2']
};

function projFindCol(headers,key){const variants=PROJ_COL_MAP[key];for(let i=0;i<headers.length;i++){const h=projNormH(headers[i]);if(variants.some(v=>h===v))return i;}return -1;}

function projLoadFile(file){
  projClearError();
  const reader=new FileReader();
  reader.onload=e=>{
    try{
      let raw;
      if(file.name.match(/\.csv$/i)){const text=new TextDecoder().decode(e.target.result),lines=text.split(/\r?\n/).filter(l=>l.trim());raw=lines.map(l=>l.split(',').map(c=>c.replace(/^"|"$/g,'').trim()));}
      else{const wb=XLSX.read(e.target.result,{type:'array'}),ws=wb.Sheets[wb.SheetNames[0]];raw=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});}
      let headerIdx=0,headerRow=null;
      for(let i=0;i<Math.min(5,raw.length);i++){const h=raw[i].map(projNormH);if(h.some(c=>c.includes('supervisor')||c.includes('theme')||c.includes('assessor'))){headerIdx=i;headerRow=raw[i];break;}}
      if(!headerRow){projShowError('Could not find header row.');return;}
      const iSup=projFindCol(headerRow,'supervisor');if(iSup===-1){projShowError('Could not find Supervisor column.');return;}
      const iTheme=projFindCol(headerRow,'theme'),iCoSup=projFindCol(headerRow,'cosupervisors'),iP1=projFindCol(headerRow,'poster1'),iP2=projFindCol(headerRow,'poster2'),iD1=projFindCol(headerRow,'dissertation1'),iD2=projFindCol(headerRow,'dissertation2');
      projRawProjects=[];
      for(let i=headerIdx+1;i<raw.length;i++){const row=raw[i];if(row.every(c=>!String(c).trim()))continue;const get=idx=>idx!==-1?String(row[idx]||'').trim():'';const splitNames=s=>s.split(/[;,\/|&]+/).map(n=>n.trim()).filter(Boolean);projRawProjects.push({theme:get(iTheme),supervisors:splitNames(get(iSup)),cosupervisors:splitNames(get(iCoSup)),poster1:get(iP1),poster2:get(iP2),diss1:get(iD1),diss2:get(iD2)});}
      if(projRawProjects.length===0){projShowError('No project rows found.');return;}
      projAnalyseBtn.disabled=false;projAnalyseBtn.textContent=`🎓 Calculate Project Workload (${projRawProjects.length} projects found) →`;
    }catch(err){projShowError('Error reading file: '+err.message);}
  };
  reader.readAsArrayBuffer(file);
}

function projGetSettings(){return{supervision:+document.getElementById('ps_sup').value||0,cosupervision:+document.getElementById('ps_cosup').value||0,diss_feedback:+document.getElementById('ps_diss_fb').value||0,diss_marking:+document.getElementById('ps_diss_mk').value||0,poster_feedback:+document.getElementById('ps_post_fb').value||0,poster_marking:+document.getElementById('ps_post_mk').value||0};}
function projSyncInlineSettings(s){document.getElementById('as_sup').value=s.supervision;document.getElementById('as_cosup').value=s.cosupervision;document.getElementById('as_diss_fb').value=s.diss_feedback;document.getElementById('as_diss_mk').value=s.diss_marking;document.getElementById('as_post_fb').value=s.poster_feedback;document.getElementById('as_post_mk').value=s.poster_marking;}

function projCalculate(projects,s){
  const map={};
  const ensure=name=>{if(!name)return null;if(!map[name])map[name]={name,supervised:[],supShare:[],cosupervised:[],coSupShare:[],diss_assessed:[],poster_assessed:[]};return map[name];};
  for(const p of projects){
    const nSups=p.supervisors.length||1;
    for(const sup of p.supervisors){if(sup){ensure(sup).supervised.push(p);ensure(sup).supShare.push(1/nSups);}}
    const nCoSups=p.cosupervisors.length||1;
    for(const co of p.cosupervisors){if(co){ensure(co).cosupervised.push(p);ensure(co).coSupShare.push(1/nCoSups);}}
    if(p.diss1)ensure(p.diss1).diss_assessed.push(p);
    if(p.diss2)ensure(p.diss2).diss_assessed.push(p);
    if(p.poster1)ensure(p.poster1).poster_assessed.push(p);
    if(p.poster2)ensure(p.poster2).poster_assessed.push(p);
  }
  return Object.values(map).map(a=>{
    const h_sup=a.supShare.reduce((t,sh)=>t+sh*s.supervision,0);
    const h_df=a.supShare.reduce((t,sh)=>t+sh*s.diss_feedback,0);
    const h_pf=a.supShare.reduce((t,sh)=>t+sh*s.poster_feedback,0);
    const h_cosup=a.coSupShare.reduce((t,sh)=>t+sh*s.cosupervision,0);
    const nDissAss=a.diss_assessed.length,nPostAss=a.poster_assessed.length;
    const h_dm=nDissAss*s.diss_marking,h_pm=nPostAss*s.poster_marking;
    const nSup=+a.supShare.reduce((t,sh)=>t+sh,0).toFixed(4);
    const nCoSup=+a.coSupShare.reduce((t,sh)=>t+sh,0).toFixed(4);
    const total=h_sup+h_cosup+h_df+h_dm+h_pf+h_pm;
    return{...a,name:a.name,nSup,nCoSup,nDissAss,nPostAss,h_sup,h_cosup,h_df,h_dm,h_pf,h_pm,total};
  });
}

function projGetSorted(){
  const q=document.getElementById('projSearch').value.toLowerCase();
  let data=projAllResults.filter(r=>r.name.toLowerCase().includes(q));
  const[col,dir]=projSortKey.split('-');
  data.sort((a,b)=>{if(col==='name')return dir==='asc'?a.name.localeCompare(b.name):b.name.localeCompare(a.name);if(col==='students')return dir==='desc'?b.nSup-a.nSup:a.nSup-b.nSup;return dir==='desc'?b.total-a.total:a.total-b.total;});
  return data;
}

function projRenderTable(){
  const data=projGetSorted(),maxTotal=Math.max(...projAllResults.map(r=>r.total),1);
  document.getElementById('projTbody').innerHTML=data.map(r=>`<tr style="cursor:pointer" data-name="${encodeURIComponent(r.name)}"><td class="name-f">${r.name}</td><td class="num" title="${r.supervised.length} project${r.supervised.length!==1?'s':''}, ${r.nSup.toFixed?r.nSup.toFixed(2):r.nSup} share">${r.supervised.length||'—'}</td><td class="num">${fh(r.h_sup)}</td><td class="num">${fh(r.h_cosup)}</td><td class="num">${fh(r.h_df)}</td><td class="num">${fh(r.h_dm)}</td><td class="num">${fh(r.h_pf)}</td><td class="num">${fh(r.h_pm)}</td><td class="tot">${fmt(r.total)}</td><td><div class="hours-bar-wrap"><div class="hours-bar"><div class="hours-bar-fill" style="width:${r.total/maxTotal*100}%;background:linear-gradient(90deg,var(--rust),var(--gold))"></div></div><span class="hours-val">${fmt(r.total)}h</span></div></td></tr>`).join('');
  const sumFn=key=>projAllResults.reduce((s,r)=>s+r[key],0);
  document.getElementById('projFoot').innerHTML=`<tr style="font-weight:600;background:var(--light-blue)"><td>Grand Total</td><td class="num">${projAllResults.reduce((s,r)=>s+r.supervised.length,0)}</td><td class="num">${fmt(sumFn('h_sup'))}</td><td class="num">${fmt(sumFn('h_cosup'))}</td><td class="num">${fmt(sumFn('h_df'))}</td><td class="num">${fmt(sumFn('h_dm'))}</td><td class="num">${fmt(sumFn('h_pf'))}</td><td class="num">${fmt(sumFn('h_pm'))}</td><td class="tot" style="color:var(--mid-blue)">${fmt(sumFn('total'))}</td><td></td></tr>`;
  document.querySelectorAll('#projTbody tr').forEach(row=>{row.addEventListener('click',()=>projOpenDetail(decodeURIComponent(row.dataset.name)));});
}

function projOpenDetail(name){
  const r=projAllResults.find(x=>x.name===name);if(!r)return;
  const pills=p=>`${p.supervisors.includes(name)?'<span class="role-pill pill-sup">Supervisor</span>':''}${p.cosupervisors.includes(name)?'<span class="role-pill pill-cosup">Co-supervisor</span>':''}${p.diss1===name||p.diss2===name?'<span class="role-pill pill-diss">Diss. Assessor</span>':''}${p.poster1===name||p.poster2===name?'<span class="role-pill pill-post">Poster Assessor</span>':''}`;
  const allProjects=[...new Map([...r.supervised,...r.cosupervised,...r.diss_assessed,...r.poster_assessed].map(p=>[p.theme+p.supervisors.join(','),p])).values()];
  let html=`<div class="panel-section"><h4>Hours Breakdown</h4>${r.h_sup>0?`<div class="panel-row"><span class="k">Supervision (${Number.isInteger(r.nSup)?r.nSup:r.nSup.toFixed(2)} student share${r.nSup!==1?'s':''})</span><span class="v">${fmt(r.h_sup)}h</span></div>`:''}${r.h_cosup>0?`<div class="panel-row"><span class="k">Co-supervision (${r.nCoSup})</span><span class="v">${fmt(r.h_cosup)}h</span></div>`:''}${r.h_df>0?`<div class="panel-row"><span class="k">Dissertation feedback</span><span class="v">${fmt(r.h_df)}h</span></div>`:''}${r.h_dm>0?`<div class="panel-row"><span class="k">Dissertation marking (${r.nDissAss})</span><span class="v">${fmt(r.h_dm)}h</span></div>`:''}${r.h_pf>0?`<div class="panel-row"><span class="k">Poster feedback</span><span class="v">${fmt(r.h_pf)}h</span></div>`:''}${r.h_pm>0?`<div class="panel-row"><span class="k">Poster marking (${r.nPostAss})</span><span class="v">${fmt(r.h_pm)}h</span></div>`:''}<div class="panel-row" style="margin-top:4px"><span class="k"><strong>Total</strong></span><span class="v big">${fmt(r.total)}h</span></div></div>`;
  if(allProjects.length>0){html+=`<div class="panel-section"><h4>Projects (${allProjects.length})</h4>`;for(const p of allProjects)html+=`<div class="proj-student"><div class="sn">${p.theme||'(No title)'}</div><div class="sr">${pills(p)}</div></div>`;html+='</div>';}
  openPanel(name,`${fmt(r.total)}h total · ${allProjects.length} project${allProjects.length!==1?'s':''}`,html);
}

projAnalyseBtn.addEventListener('click',()=>{
  projSettings=projGetSettings();projSyncInlineSettings(projSettings);projAllResults=projCalculate(projRawProjects,projSettings);
  const totalH=projAllResults.reduce((s,r)=>s+r.total,0),nAc=projAllResults.length;
  document.getElementById('projMeta').textContent=`${projRawProjects.length} projects · ${nAc} academics · ${projAllResults.reduce((s,r)=>s+r.supervised.length,0)} supervision roles`;
  document.getElementById('projStatsBar').innerHTML=[['Projects',projRawProjects.length],['Academics',nAc],['Total hrs',fmt(totalH)],['Avg hrs',fmt(totalH/nAc)]].map(([l,v])=>`<div class="stat-card rust"><div class="sc-v">${v}</div><div class="sc-l">${l}</div></div>`).join('');
  document.getElementById('proj-landing').style.display='none';document.getElementById('proj-content').style.display='block';
  document.getElementById('badge-project').textContent=nAc+' academics';
  projRenderTable();updateCombStatus();
});

document.getElementById('projBtnBack').addEventListener('click',()=>{document.getElementById('proj-landing').style.display='';document.getElementById('proj-content').style.display='none';});
document.getElementById('projBtnSettings').addEventListener('click',()=>document.getElementById('projInlineSettings').classList.toggle('open'));
document.getElementById('projRecalcBtn').addEventListener('click',()=>{projSettings={supervision:+document.getElementById('as_sup').value||0,cosupervision:+document.getElementById('as_cosup').value||0,diss_feedback:+document.getElementById('as_diss_fb').value||0,diss_marking:+document.getElementById('as_diss_mk').value||0,poster_feedback:+document.getElementById('as_post_fb').value||0,poster_marking:+document.getElementById('as_post_mk').value||0};projAllResults=projCalculate(projRawProjects,projSettings);projRenderTable();updateCombStatus();});
document.getElementById('projSearch').addEventListener('input',projRenderTable);
document.getElementById('projSortSel').addEventListener('change',e=>{projSortKey=e.target.value;projRenderTable();});
document.querySelector('#proj-content table.proj-table thead').addEventListener('click',e=>{const th=e.target.closest('th[data-projsort]');if(!th)return;const col=th.dataset.projsort;const[curCol,curDir]=projSortKey.split('-');if(curCol===col)projSortKey=col+'-'+(curDir==='desc'?'asc':'desc');else projSortKey=col+'-'+(col==='name'?'asc':'desc');projRenderTable();});

document.getElementById('projBtnExport').addEventListener('click',()=>{
  const wb=XLSX.utils.book_new();
  const rows=[['Academic','Supervised (projects)','Sup. share','Co-supervised (projects)','Co-sup. share','Diss.Assessed','Poster Assessed','Sup.hrs','Co-sup.hrs','Diss.Feedback','Diss.Marking','Poster Feedback','Poster Marking','Total']];
  for(const r of projAllResults)rows.push([r.name,r.supervised.length,+r.nSup.toFixed(2),r.cosupervised.length,+r.nCoSup.toFixed(2),r.nDissAss,r.nPostAss,+r.h_sup.toFixed(2),+r.h_cosup.toFixed(2),+r.h_df.toFixed(2),+r.h_dm.toFixed(2),+r.h_pf.toFixed(2),+r.h_pm.toFixed(2),+r.total.toFixed(2)]);
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(rows),'Project Workload');
  const sRows=[['Setting','Value'],['Supervision',projSettings.supervision],['Co-supervision',projSettings.cosupervision],['Diss. feedback',projSettings.diss_feedback],['Diss. marking',projSettings.diss_marking],['Poster feedback',projSettings.poster_feedback],['Poster marking',+projSettings.poster_marking.toFixed(4)]];
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(sRows),'Settings');
  XLSX.writeFile(wb,'project_workload.xlsx');
});
