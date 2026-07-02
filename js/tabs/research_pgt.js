// ═══════════════════════════════════════════════════════
// TAB — RESEARCH PROJECTS PGT
// ═══════════════════════════════════════════════════════
let pgtRawData=[],pgtAllResults=[];
const PGTSETTINGS_DEFAULTS={supervision:12,diss_feedback:3,poster_feedback:0.5,diss_marking:2,poster_marking:0.5,marking_students:2};
let pgtSettings={...PGTSETTINGS_DEFAULTS},pgtSortKey='total-desc';

const pgtDropZone=document.getElementById('pgtDropZone');
const pgtFileInput=document.getElementById('pgtFileInput');
const pgtAnalyseBtn=document.getElementById('pgtAnalyseBtn');

pgtDropZone.addEventListener('dragover',e=>{e.preventDefault();pgtDropZone.classList.add('drag-over');});
pgtDropZone.addEventListener('dragleave',()=>pgtDropZone.classList.remove('drag-over'));
pgtDropZone.addEventListener('drop',e=>{e.preventDefault();pgtDropZone.classList.remove('drag-over');if(e.dataTransfer.files[0])pgtLoadFile(e.dataTransfer.files[0]);});
pgtFileInput.addEventListener('change',e=>{if(e.target.files[0])pgtLoadFile(e.target.files[0]);});
document.getElementById('pgtSettingsHdr').addEventListener('click',()=>{document.getElementById('pgtSettingsBody').classList.toggle('open');document.getElementById('pgtSettingsHdr').classList.toggle('open');});

function pgtShowError(msg){const el=document.getElementById('pgtError');el.textContent=msg;el.classList.add('show');}
function pgtClearError(){document.getElementById('pgtError').classList.remove('show');}
function pgtNormH(s){return String(s).toLowerCase().replace(/[\s_\-]/g,'').replace(/[^a-z0-9]/g,'');}

const PGT_COL_MAP={
  long_name:['long_name','longname','studentname','student','name'],
  location:['location','loc'],
  allocation:['allocation','alloc','staff','staffmember','staffmember(s)','staffmember(s)','supervisor','supervisors']
};

function pgtFindCol(headers,key){
  const variants=PGT_COL_MAP[key];
  for(let i=0;i<headers.length;i++){
    const h=pgtNormH(headers[i]);
    if(variants.some(v=>h===v))return i;
  }
  return -1;
}

function pgtLoadFile(file){
  pgtClearError();
  const reader=new FileReader();
  reader.onload=e=>{
    try{
      let raw;
      const wb=XLSX.read(e.target.result,{type:'array'});
      const ws=wb.Sheets[wb.SheetNames[0]];
      raw=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});
      // Find header row — search first 10 rows
      let headerIdx=-1,headerRow=null;
      for(let i=0;i<Math.min(10,raw.length);i++){
        const h=raw[i].map(pgtNormH);
        if(h.some(c=>c==='long_name'||c==='longname'||c==='studentname'||c==='student') &&
           h.some(c=>c==='allocation'||c==='alloc'||c==='staff')){
          headerIdx=i;
          headerRow=raw[i];
          break;
        }
      }
      if(!headerRow){
        // Fallback: try simpler detection
        for(let i=0;i<Math.min(10,raw.length);i++){
          const h=raw[i].map(String);
          if(h.some(c=>/long/i.test(c)) && h.some(c=>/allocation/i.test(c))){
            headerIdx=i;
            headerRow=raw[i];
            break;
          }
        }
      }
      if(!headerRow){pgtShowError('Could not find header row. Expected columns: Long_name, Location, Allocation');return;}
      const iLong=pgtFindCol(headerRow,'long_name');
      const iLoc=pgtFindCol(headerRow,'location');
      const iAlloc=pgtFindCol(headerRow,'allocation');
      if(iLong===-1){pgtShowError('Could not find "Long_name" column.');return;}
      if(iAlloc===-1){pgtShowError('Could not find "Allocation" column.');return;}
      pgtRawData=[];
      for(let i=headerIdx+1;i<raw.length;i++){
        const row=raw[i];
        if(row.every(c=>!String(c).trim()))continue;
        const get=idx=>idx!==-1?String(row[idx]||'').trim():'';
        const studentName=get(iLong);
        if(!studentName)continue;
        const location=get(iLoc);
        const allocRaw=get(iAlloc);
        const staffMembers=allocRaw.split('/').map(n=>n.trim()).filter(Boolean);
        if(staffMembers.length===0)continue;
        pgtRawData.push({studentName,location,staffMembers});
      }
      if(pgtRawData.length===0){pgtShowError('No student rows found.');return;}
      pgtAnalyseBtn.disabled=false;
      pgtAnalyseBtn.textContent=`🔬 Calculate PGT Research Hours (${pgtRawData.length} students found) →`;
    }catch(err){pgtShowError('Error reading file: '+err.message);}
  };
  reader.readAsArrayBuffer(file);
}

function pgtGetSettings(){
  return{
    supervision:+document.getElementById('pgt_sup').value||0,
    diss_feedback:+document.getElementById('pgt_diss_fb').value||0,
    poster_feedback:+document.getElementById('pgt_post_fb').value||0,
    diss_marking:+document.getElementById('pgt_diss_mk').value||0,
    poster_marking:+document.getElementById('pgt_post_mk').value||0,
    marking_students:+document.getElementById('pgt_mark_students').value||2
  };
}
function pgtSyncInlineSettings(s){
  document.getElementById('as_pgt_sup').value=s.supervision;
  document.getElementById('as_pgt_diss_fb').value=s.diss_feedback;
  document.getElementById('as_pgt_post_fb').value=s.poster_feedback;
  document.getElementById('as_pgt_diss_mk').value=s.diss_marking;
  document.getElementById('as_pgt_post_mk').value=s.poster_marking;
  document.getElementById('as_pgt_mark_students').value=s.marking_students;
}

function pgtCalculate(students,s){
  const map={};
  const ensure=name=>{if(!name)return null;if(!map[name])map[name]={name,students:[],nShares:0};return map[name];};
  for(const st of students){
    const nStaff=st.staffMembers.length;
    const share=1/nStaff;
    for(const staff of st.staffMembers){
      if(staff){
        const entry=ensure(staff);
        entry.students.push(st);
        entry.nShares+=share;
      }
    }
  }
  return Object.values(map).map(a=>{
    const nSup=a.nShares;
    const h_sup=nSup*s.supervision;
    const h_df=nSup*s.diss_feedback;
    const h_pf=nSup*s.poster_feedback;
    const h_dm=s.diss_marking*s.marking_students*nSup;
    const h_pm=s.poster_marking*s.marking_students*nSup;
    const total=h_sup+h_df+h_pf+h_dm+h_pm;
    const studentList=a.students.map(st=>({name:st.studentName,location:st.location,staffMembers:st.staffMembers}));
    return{...a,name:a.name,nSup,studentList,h_sup,h_df,h_pf,h_dm,h_pm,total};
  });
}

function pgtGetSorted(){
  const q=document.getElementById('pgtSearch').value.toLowerCase();
  let data=pgtAllResults.filter(r=>r.name.toLowerCase().includes(q));
  const[col,dir]=pgtSortKey.split('-');
  data.sort((a,b)=>{
    if(col==='name')return dir==='asc'?a.name.localeCompare(b.name):b.name.localeCompare(a.name);
    if(col==='students')return dir==='desc'?b.nSup-a.nSup:a.nSup-b.nSup;
    return dir==='desc'?b.total-a.total:a.total-b.total;
  });
  return data;
}

function pgtRenderTable(){
  const data=pgtGetSorted(),maxTotal=Math.max(...pgtAllResults.map(r=>r.total),1);
  document.getElementById('pgtTbody').innerHTML=data.map(r=>{
    const studCount=r.studentList.length;
    return `<tr style="cursor:pointer" data-name="${encodeURIComponent(r.name)}">
      <td class="name-f">${r.name}</td>
      <td class="num">${studCount||'—'}</td>
      <td class="num">${fh(r.h_sup)}</td>
      <td class="num">${fh(r.h_df)}</td>
      <td class="num">${fh(r.h_pf)}</td>
      <td class="num">${fh(r.h_dm)}</td>
      <td class="num">${fh(r.h_pm)}</td>
      <td class="tot">${fmt(r.total)}</td>
      <td><div class="hours-bar-wrap"><div class="hours-bar"><div class="hours-bar-fill" style="width:${r.total/maxTotal*100}%;background:linear-gradient(90deg,#6a1b9a,#e040fb)"></div></div><span class="hours-val">${fmt(r.total)}h</span></div></td>
    </tr>`;
  }).join('');
  const sumFn=key=>pgtAllResults.reduce((s,r)=>s+r[key],0);
  const totalStudents=pgtAllResults.reduce((s,r)=>s+r.studentList.length,0);
  document.getElementById('pgtFoot').innerHTML=`<tr style="font-weight:600;background:var(--light-blue)"><td>Grand Total</td><td class="num">${totalStudents}</td><td class="num">${fmt(sumFn('h_sup'))}</td><td class="num">${fmt(sumFn('h_df'))}</td><td class="num">${fmt(sumFn('h_pf'))}</td><td class="num">${fmt(sumFn('h_dm'))}</td><td class="num">${fmt(sumFn('h_pm'))}</td><td class="tot" style="color:var(--mid-blue)">${fmt(sumFn('total'))}</td><td></td></tr>`;
  document.querySelectorAll('#pgtTbody tr').forEach(row=>{row.addEventListener('click',()=>pgtOpenDetail(decodeURIComponent(row.dataset.name)));});
}

function pgtOpenDetail(name){
  const r=pgtAllResults.find(x=>x.name===name);if(!r)return;
  let html=`<div class="panel-section"><h4>Hours Breakdown</h4>
    ${r.h_sup>0?`<div class="panel-row"><span class="k">Supervision (${r.studentList.length} student${r.studentList.length!==1?'s':''})</span><span class="v">${fmt(r.h_sup)}h</span></div>`:''}
    ${r.h_df>0?`<div class="panel-row"><span class="k">Dissertation feedback</span><span class="v">${fmt(r.h_df)}h</span></div>`:''}
    ${r.h_pf>0?`<div class="panel-row"><span class="k">Poster feedback</span><span class="v">${fmt(r.h_pf)}h</span></div>`:''}
    ${r.h_dm>0?`<div class="panel-row"><span class="k">Dissertation marking (${pgtSettings.marking_students} students)</span><span class="v">${fmt(r.h_dm)}h</span></div>`:''}
    ${r.h_pm>0?`<div class="panel-row"><span class="k">Poster marking (${pgtSettings.marking_students} students)</span><span class="v">${fmt(r.h_pm)}h</span></div>`:''}
    <div class="panel-row" style="margin-top:4px"><span class="k"><strong>Total</strong></span><span class="v big">${fmt(r.total)}h</span></div></div>`;
  if(r.studentList.length>0){
    html+=`<div class="panel-section"><h4>Students (${r.studentList.length})</h4>`;
    for(const st of r.studentList){
      html+=`<div class="proj-student"><div class="sn">${st.name}</div><div class="sr">${st.location?`<span class="role-pill pill-cosup" style="background:#aaa">${st.location}</span>`:''}${st.staffMembers.length>1?`<span class="role-pill pill-sup">${st.staffMembers.length} staff</span>`:''}</div></div>`;
    }
    html+='</div>';
  }
  openPanel(name,`${fmt(r.total)}h total · ${r.studentList.length} student${r.studentList.length!==1?'s':''}`,html);
}

pgtAnalyseBtn.addEventListener('click',()=>{
  pgtSettings=pgtGetSettings();pgtSyncInlineSettings(pgtSettings);
  pgtAllResults=pgtCalculate(pgtRawData,pgtSettings);
  const totalH=pgtAllResults.reduce((s,r)=>s+r.total,0),nAc=pgtAllResults.length;
  const totalStudents=pgtAllResults.reduce((s,r)=>s+r.studentList.length,0);
  document.getElementById('pgtMeta').textContent=`${pgtRawData.length} students · ${nAc} academics · ${fmt(totalH)} total hours`;
  document.getElementById('pgtStatsBar').innerHTML=[['Students',pgtRawData.length],['Academics',nAc],['Total hrs',fmt(totalH)],['Avg hrs',fmt(totalH/nAc)]].map(([l,v])=>`<div class="stat-card" style="background:#f3e5f5;color:#6a1b9a"><div class="sc-v">${v}</div><div class="sc-l">${l}</div></div>`).join('');
  document.getElementById('pgt-landing').style.display='none';document.getElementById('pgt-content').style.display='block';
  document.getElementById('badge-researchpgt').textContent=nAc+' academics';
  pgtRenderTable();updateCombStatus();
});

document.getElementById('pgtBtnBack').addEventListener('click',()=>{document.getElementById('pgt-landing').style.display='';document.getElementById('pgt-content').style.display='none';});
document.getElementById('pgtBtnSettings').addEventListener('click',()=>document.getElementById('pgtInlineSettings').classList.toggle('open'));
document.getElementById('pgtRecalcBtn').addEventListener('click',()=>{
  pgtSettings={
    supervision:+document.getElementById('as_pgt_sup').value||0,
    diss_feedback:+document.getElementById('as_pgt_diss_fb').value||0,
    poster_feedback:+document.getElementById('as_pgt_post_fb').value||0,
    diss_marking:+document.getElementById('as_pgt_diss_mk').value||0,
    poster_marking:+document.getElementById('as_pgt_post_mk').value||0,
    marking_students:+document.getElementById('as_pgt_mark_students').value||2
  };
  pgtAllResults=pgtCalculate(pgtRawData,pgtSettings);
  pgtRenderTable();updateCombStatus();
});
document.getElementById('pgtSearch').addEventListener('input',pgtRenderTable);
document.getElementById('pgtSortSel').addEventListener('change',e=>{pgtSortKey=e.target.value;pgtRenderTable();});
document.querySelector('#pgt-content table.pgt-table thead').addEventListener('click',e=>{
  const th=e.target.closest('th[data-pgtsort]');if(!th)return;
  const col=th.dataset.pgtsort;
  const[curCol,curDir]=pgtSortKey.split('-');
  if(curCol===col)pgtSortKey=col+'-'+(curDir==='desc'?'asc':'desc');
  else pgtSortKey=col+'-'+(col==='name'?'asc':'desc');
  pgtRenderTable();
});

document.getElementById('pgtBtnExport').addEventListener('click',()=>{
  const wb=XLSX.utils.book_new();
  const rows=[['Academic','Students','Sup.hrs','Diss.Feedback','Poster Feedback','Diss.Marking','Poster Marking','Total']];
  for(const r of pgtAllResults)rows.push([r.name,r.studentList.length,+r.h_sup.toFixed(2),+r.h_df.toFixed(2),+r.h_pf.toFixed(2),+r.h_dm.toFixed(2),+r.h_pm.toFixed(2),+r.total.toFixed(2)]);
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(rows),'PGT Research');
  const sRows=[['Setting','Value'],['Supervision',pgtSettings.supervision],['Diss. feedback',pgtSettings.diss_feedback],['Poster feedback',pgtSettings.poster_feedback],['Diss. marking',pgtSettings.diss_marking],['Poster marking',pgtSettings.poster_marking],['Marking students',pgtSettings.marking_students]];
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(sRows),'Settings');
  XLSX.writeFile(wb,'research_projects_pgt.xlsx');
});

// Provide totals to Combined view
window.getResearchPgtHoursTotals=function(){
  const totals={};
  for(const r of pgtAllResults)totals[r.name]=r.total;
  return totals;
};
