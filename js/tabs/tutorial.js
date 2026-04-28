// ═══════════════════════════════════════════════════════
// TAB 2 — TUTORIAL WORKLOAD
// ═══════════════════════════════════════════════════════
let tutAllTutors=[],tutMaxHours=0,tutSortCol='hours',tutSortDir='desc';
const tutDropZone=document.getElementById('tutDropZone'),tutFileInput=document.getElementById('tutFileInput');

tutDropZone.addEventListener('dragover',e=>{e.preventDefault();tutDropZone.classList.add('drag-over');});
tutDropZone.addEventListener('dragleave',()=>tutDropZone.classList.remove('drag-over'));
tutDropZone.addEventListener('drop',e=>{e.preventDefault();tutDropZone.classList.remove('drag-over');if(e.dataTransfer.files[0])tutProcessFile(e.dataTransfer.files[0]);});
tutFileInput.addEventListener('change',e=>{if(e.target.files[0])tutProcessFile(e.target.files[0]);});

function tutShowError(msg){const el=document.getElementById('tutError');el.textContent=msg;el.classList.add('show');}
function tutClearError(){document.getElementById('tutError').classList.remove('show');}

function tutProcessFile(file){
  tutClearError();
  const y1h=+document.getElementById('tutY1Hours').value||16,oh=+document.getElementById('tutOtherHours').value||8,extra=+document.getElementById('tutExtraAllowance').value||0;
  const courseCodes=document.getElementById('tutSelectedCourses').value.split(',').map(s=>s.trim().toUpperCase()).filter(Boolean);
  const reader=new FileReader();
  reader.onload=e=>{
    try{
      const wb=XLSX.read(e.target.result,{type:'array'}),ws=wb.Sheets[wb.SheetNames[0]],raw=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});
      const headerRow=raw[5]||[];const norm=s=>String(s).toLowerCase().replace(/\s+/g,'').replace(/[^a-z0-9]/g,''),headers=headerRow.map(norm),col=key=>headers.indexOf(norm(key));
      const iYear=col('Year of Study'),iSurname=col('Surname'),iFirst=col('First Name'),iCourse=col('Course'),iEmail=col('UoN Email'),iTutor=col('Tutor'),iTutorEmail=col('Tutor email'),iStaff=col('Staff Indicator');
      if(iTutor===-1){tutShowError('Could not find a "Tutor" column. Check headers are on row 6.');return;}
      const tutorMap={};
      raw.slice(6).forEach(row=>{
        if(iStaff!==-1&&String(row[iStaff]).trim().toLowerCase()==='yes')return;
        const tutorName=String(row[iTutor]||'').trim();if(!tutorName)return;
        const yearNum=parseInt(String(row[iYear]||'').trim(),10),isY1=yearNum===1;
        const tutorEmail=iTutorEmail!==-1?String(row[iTutorEmail]||'').trim():'';
        const studentName=[String(row[iFirst]||'').trim(),String(row[iSurname]||'').trim()].filter(Boolean).join(' ');
        const course=iCourse!==-1?String(row[iCourse]||'').trim():'',studentEmail=iEmail!==-1?String(row[iEmail]||'').trim():'';
        if(!tutorMap[tutorName])tutorMap[tutorName]={name:tutorName,email:tutorEmail,year1:[],other:[]};
        const entry={name:studentName,year:String(row[iYear]||'').trim(),course,email:studentEmail};
        if(isY1)tutorMap[tutorName].year1.push(entry);else tutorMap[tutorName].other.push(entry);
      });
      tutAllTutors=Object.values(tutorMap).map(t=>{
        const extraHours=t.year1.filter(s=>courseCodes.includes(s.course.toUpperCase())).length*extra + t.other.filter(s=>courseCodes.includes(s.course.toUpperCase())).length*extra;
        return {...t,totalTutees:t.year1.length+t.other.length,hours:t.year1.length*y1h+t.other.length*oh+extraHours,extraHours};
      });
      if(tutAllTutors.length===0){tutShowError('No tutee records found.');return;}
      tutMaxHours=Math.max(...tutAllTutors.map(t=>t.hours));
      document.getElementById('tut-landing').style.display='none';document.getElementById('tut-content').style.display='block';
      document.getElementById('badge-tutorial').textContent=tutAllTutors.length+' tutors';
      document.getElementById('tutMeta').textContent=`${tutAllTutors.length} tutors · ${tutAllTutors.reduce((s,t)=>s+t.totalTutees,0)} tutees`;
      tutRenderSummary();tutRenderTable();updateCombStatus();
    }catch(err){tutShowError('Error reading file: '+err.message);}
  };
  reader.readAsArrayBuffer(file);
}

function tutRenderSummary(){
  const totalH=tutAllTutors.reduce((s,t)=>s+t.hours,0),totalT=tutAllTutors.reduce((s,t)=>s+t.totalTutees,0),avg=totalH/tutAllTutors.length,maxT=tutAllTutors.reduce((a,b)=>a.hours>b.hours?a:b);
  document.getElementById('tutSummary').innerHTML=[['Tutors',tutAllTutors.length],['Total Tutees',totalT],['Total Hours',totalH],['Avg Hours',avg.toFixed(1)],['Peak Load',`${maxT.hours}h`]].map(([l,v])=>`<div class="tut-stat"><div class="val">${v}</div><div class="lbl">${l}</div></div>`).join('');
}

function tutRenderTable(){
  const q=document.getElementById('tutSearch').value.toLowerCase();
  let data=tutAllTutors.filter(t=>t.name.toLowerCase().includes(q)||t.email.toLowerCase().includes(q));
  const colMap={name:'name',email:'email',total:'totalTutees',hours:'hours'},key=colMap[tutSortCol]||tutSortCol;
  data.sort((a,b)=>{const av=a[key]??0,bv=b[key]??0;return typeof av==='string'?(tutSortDir==='asc'?av.localeCompare(bv):bv.localeCompare(av)):(tutSortDir==='asc'?av-bv:bv-av);});
  document.getElementById('tutTbody').innerHTML=data.map(t=>`<tr data-name="${encodeURIComponent(t.name)}" style="cursor:pointer"><td class="tutor-name">${t.name}</td><td>${t.email||'—'}</td><td style="text-align:right">${t.year1.length}</td><td style="text-align:right">${t.other.length}</td><td style="text-align:right">${t.totalTutees}</td><td><div class="hours-bar-wrap"><div class="hours-bar"><div class="hours-bar-fill" style="width:${tutMaxHours>0?t.hours/tutMaxHours*100:0}%"></div></div><span class="hours-val">${t.hours}h</span></div></td></tr>`).join('');
  document.querySelectorAll('#tutTbody tr').forEach(row=>{
    row.addEventListener('click',()=>{
      const name=decodeURIComponent(row.dataset.name),tutor=tutAllTutors.find(t=>t.name===name);
      if(!tutor)return;
      const y1=tutor.year1.map(s=>`<div class="panel-row"><span class="k">${s.name}</span><span class="v">${s.course||'—'}</span></div>`).join(''),ot=tutor.other.map(s=>`<div class="panel-row"><span class="k">${s.name}</span><span class="v">${s.course||'—'}</span></div>`).join('');
      openPanel(tutor.name,tutor.email||'',`<div class="panel-section"><div class="panel-row"><span class="k">Year 1 tutees</span><span class="v">${tutor.year1.length}</span></div><div class="panel-row"><span class="k">Other year tutees</span><span class="v">${tutor.other.length}</span></div>${tutor.extraHours>0?`<div class="panel-row"><span class="k">Extra allowance (selected courses)</span><span class="v">${tutor.extraHours}h</span></div>`:''}<div class="panel-row"><span class="k">Total hours</span><span class="v big">${tutor.hours}h</span></div></div>${tutor.year1.length?`<div class="panel-section"><h4>Year 1 Tutees</h4>${y1}</div>`:''}${tutor.other.length?`<div class="panel-section"><h4>Other Tutees</h4>${ot}</div>`:''}`);
    });
  });
}

document.getElementById('tutSearch').addEventListener('input',tutRenderTable);
document.getElementById('tutSort').addEventListener('change',e=>{const[c,d]=e.target.value.split('-');tutSortCol=c;tutSortDir=d;tutRenderTable();});
document.getElementById('tutBtnBack').addEventListener('click',()=>{document.getElementById('tut-landing').style.display='';document.getElementById('tut-content').style.display='none';});
document.getElementById('tutTable').querySelector('thead').addEventListener('click',e=>{
  const th=e.target.closest('th[data-tutsort]');if(!th)return;const col=th.dataset.tutsort;
  if(tutSortCol===col)tutSortDir=tutSortDir==='asc'?'desc':'asc';else{tutSortCol=col;tutSortDir=(col==='hours'||col==='total'||col==='year1'||col==='other')?'desc':'asc';}tutRenderTable();
});
