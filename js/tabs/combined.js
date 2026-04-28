// ═══════════════════════════════════════════════════════
// COMBINED TOTALS TAB
// ═══════════════════════════════════════════════════════
let combData=[],combSortKey='total-desc';
const SRC_LABELS={tl:'📅 Teaching',assessment:'📝 Non-timetabled assess.',proj:'🎓 Project',tut:'👥 Tutorial',mmi:'🩺 MMI',cit:'🏛 Citizenship',res:'🔬 Research',pgr:'👨‍🎓 PGR'};
const staffTags=new Map();
const tagRules=new Map();
let activeTagFilter=null;
let moduleTagFilter=null;
const moduleTags=new Map();
const manualMappings = new Map(); // normKey(sourceName) → targetCanonical

let fteTarget=1600;
const staffFte=new Map();

function getEffectiveFte(canonical){
  const nk=normKey(canonical);
  const manual=staffFte.get(nk);
  if(manual!=null)return manual;
  const today=todayDate();
  let combined=1.0;
  let anyTagFraction=false;
  activeTagsForPerson(canonical).forEach(tag=>{
    const rule=tagRules.get(tag);
    if(!rule)return;
    if(rule.expiry&&rule.expiry<today)return;
    if(rule.fte!=null&&rule.fte!==1){
      combined*=rule.fte;
      anyTagFraction=true;
    }
  });
  return anyTagFraction?combined:1.0;
}
function getFte(canonical){return getEffectiveFte(canonical);}
function setFte(canonical,fraction){
  const nk=normKey(canonical);
  if(fraction===1.0||fraction===null||isNaN(fraction))staffFte.delete(nk);
  else staffFte.set(nk,Math.min(1.0,Math.max(0.05,+fraction)));
}
function personalTarget(canonical){return fteTarget*getEffectiveFte(canonical);}
function fteClass(pct){
  if(pct<80)return'fte-under';
  if(pct<=100)return'fte-ok';
  if(pct<=115)return'fte-warn';
  return'fte-over';
}
function ftePct(canonical,totalHours){
  const target=personalTarget(canonical);
  return target>0?Math.round(totalHours/target*100):0;
}
function fteBarHtml(canonical,totalHours){
  const pct=ftePct(canonical,totalHours);
  const cls=fteClass(pct);
  const fillPct=Math.min(pct,130)/130*100;
  const markerLeft=(100/130*100).toFixed(2);
  const target=personalTarget(canonical);
  const fte=getEffectiveFte(canonical);
  const isManual=staffFte.get(normKey(canonical))!=null;
  const fteSource=isManual?'manual override':'tag rules';
  const tip=`${totalHours.toFixed(1)}h of ${target.toFixed(0)}h target (${fteTarget}h × ${fte.toFixed(2)} FTE via ${fteSource})`;
  return`<div class="fte-wrap" title="${tip}">
    <span class="fte-pct ${cls}">${pct}%</span>
    <div class="fte-bar-outer">
      <div class="fte-bar-fill ${cls}" style="width:${fillPct.toFixed(1)}%"></div>
      <div class="fte-target-mark" style="left:${markerLeft}%"></div>
    </div>
  </div>`;
}

function updateCombStatus(){
  const hasTL=tlAllStaff.length>0,hasTUT=tutAllTutors.length>0,hasProj=projAllResults.length>0;
  const hasMmi=mmiResults.filter(r=>r.isActiveStaff).length>0;
  const hasCit=Object.keys(citizenshipTotals).length>0;
  const resHours=typeof window.getResHoursTotals==='function'?window.getResHoursTotals():{};
  const hasRes=Object.keys(resHours).length>0;
  const pgrHours=typeof window.getPgrHoursTotals==='function'?window.getPgrHoursTotals():{};
  const hasPgr=Object.keys(pgrHours).length>0;
  const assessmentHours=typeof window.getAssessmentHoursTotals==='function'?window.getAssessmentHoursTotals():{};
  const hasAssessment=Object.keys(assessmentHours).length>0;
  const pill=(id,loaded,loadedText,defaultText)=>{const el=document.getElementById(id);if(!el)return;el.className='status-pill'+(loaded?' loaded':'');el.textContent=loaded?loadedText:defaultText;};
  pill('comb-status-tl',hasTL,`Teaching: ${tlAllStaff.length} staff`,'Teaching Load');
  pill('comb-status-assessment',hasAssessment,`Non-timetabled assess.: ${Object.keys(assessmentHours).length} staff`,'Non-timetabled assess.');
  pill('comb-status-proj',hasProj,`Project: ${projAllResults.length} academics`,'Project Supervision');
  pill('comb-status-tut',hasTUT,`Tutorial: ${tutAllTutors.length} tutors`,'Tutorial Workload');
  pill('comb-status-mmi',hasMmi,`MMI: ${mmiResults.filter(r=>r.isActiveStaff).length} staff`,'MMIs');
  pill('comb-status-cit',hasCit,`Citizenship: ${Object.keys(citizenshipTotals).length} staff`,'Citizenship');
  pill('comb-status-res',hasRes,`Research: ${Object.keys(resHours).length} staff`,'Research Hours');
  pill('comb-status-pgr',hasPgr,`PGR: ${Object.keys(pgrHours).length} staff`,'PGR Supervision');
  document.getElementById('combMergeBtn').disabled=!(hasTL||hasTUT||hasProj||hasMmi||hasCit||hasRes||hasPgr||hasAssessment);
}

function recomputeCombData(){
  purgeExpiredAssignments();
  const resHoursTotals=typeof window.getResHoursTotals==='function'?window.getResHoursTotals():{};
  const pgrHoursTotals=typeof window.getPgrHoursTotals==='function'?window.getPgrHoursTotals():{};
  const assessmentHoursTotals=typeof window.getAssessmentHoursTotals==='function'?window.getAssessmentHoursTotals():{};
  combData.forEach(d=>{
    const contactH=d.tlName?tlAllWeeks.reduce((s,w)=>s+calcHours(tlStaffData[d.tlName]?.[w],tlRealisticMode),0):0;
    const projBase=d.projName?(projAllResults.find(r=>r.name===d.projName)?.total||0):0;
    const{tlTotal,projTotal}=applyBonuses(d.canonical,contactH,projBase);
    d.tlHours=d.tlName?tlTotal:0;
    d.projHours=d.projName?projTotal:0;
    d.resHours=d.resName?(resHoursTotals[d.resName]||0):0;
    d.pgrHours=d.pgrName?(pgrHoursTotals[d.pgrName]||0):0;
    d.assessmentHours=d.assessmentName?(assessmentHoursTotals[d.assessmentName]||0):0;
    d.total=d.tlHours+d.assessmentHours+d.projHours+d.tutHours+d.mmiHours+d.citHours+d.resHours+d.pgrHours;
    d._bonuses=computeBonuses(d.canonical);
  });
}

function doMerge(){
  purgeExpiredAssignments();
  const rawLists=[];
  if(tlAllStaff.length>0)rawLists.push({source:'tl',names:tlAllStaff});
  const assessmentHoursTotals=typeof window.getAssessmentHoursTotals==='function'?window.getAssessmentHoursTotals():{};
  const assessmentNames=Object.keys(assessmentHoursTotals);
  if(assessmentNames.length>0)rawLists.push({source:'assessment',names:assessmentNames});
  if(projAllResults.length>0)rawLists.push({source:'proj',names:projAllResults.map(r=>r.name)});
  if(tutAllTutors.length>0)rawLists.push({source:'tut',names:tutAllTutors.map(t=>t.name)});
  const activeMmi=mmiResults.filter(r=>r.isActiveStaff);
  if(activeMmi.length>0)rawLists.push({source:'mmi',names:activeMmi.map(r=>r.name)});
  const citNames=Object.keys(citizenshipTotals);
  if(citNames.length>0)rawLists.push({source:'cit',names:citNames});
  const resHoursTotals=typeof window.getResHoursTotals==='function'?window.getResHoursTotals():{};
  const resNames=Object.keys(resHoursTotals);
  if(resNames.length>0)rawLists.push({source:'res',names:resNames});
  const pgrHoursTotals=typeof window.getPgrHoursTotals==='function'?window.getPgrHoursTotals():{};
  const pgrNames=Object.keys(pgrHoursTotals);
  if(pgrNames.length>0)rawLists.push({source:'pgr',names:pgrNames});

  // Merge names, then post-process manual mappings (post-merge avoids data loss
  // from source-specific lookups using a rewritten name)
  const groups=mergeNameLists(rawLists);

  for(const[fromNk,target]of manualMappings){
    const fromGroup=groups.find(g=>!g._merged&&normKey(g.canonical)===fromNk);
    const toGroup=groups.find(g=>!g._merged&&g.canonical===target);
    if(fromGroup&&toGroup){
      Object.keys(fromGroup.sources).forEach(src=>{
        if(!toGroup.sources[src]){
          toGroup.sources[src]=fromGroup.sources[src];
        }else{
          if(!toGroup._extraSources)toGroup._extraSources={};
          if(!toGroup._extraSources[src])toGroup._extraSources[src]=[];
          toGroup._extraSources[src].push(fromGroup.sources[src]);
        }
      });
      toGroup.matchType='exact';
      fromGroup._merged=true;
    }
  }

  combData=groups.filter(g=>!g._merged).map(g=>{
    const tlName=g.sources['tl']||null,assessmentName=g.sources['assessment']||null,projName=g.sources['proj']||null,tutName=g.sources['tut']||null,mmiName=g.sources['mmi']||null,citName=g.sources['cit']||null,resName=g.sources['res']||null,pgrName=g.sources['pgr']||null;
    const tlExtra=g._extraSources?.tl||[],assessmentExtra=g._extraSources?.assessment||[],projExtra=g._extraSources?.proj||[],tutExtra=g._extraSources?.tut||[],mmiExtra=g._extraSources?.mmi||[],citExtra=g._extraSources?.cit||[],resExtra=g._extraSources?.res||[],pgrExtra=g._extraSources?.pgr||[];
    const contactH=tlName||tlExtra.length?tlAllWeeks.reduce((s,w)=>{
      let h=0;
      if(tlName)h+=calcHours(tlStaffData[tlName]?.[w],tlRealisticMode);
      tlExtra.forEach(en=>h+=calcHours(tlStaffData[en]?.[w],tlRealisticMode));
      return s+h;
    },0):0;
    const projBase=(projName?projAllResults.find(r=>r.name===projName)?.total||0:0)+projExtra.reduce((s,en)=>s+(projAllResults.find(r=>r.name===en)?.total||0),0);
    const{tlTotal,projTotal}=applyBonuses(g.canonical,contactH,projBase);
    const tlHours=tlName||tlExtra.length?tlTotal:0;
    const tutHours=(tutName?tutAllTutors.find(t=>t.name===tutName)?.hours||0:0)+tutExtra.reduce((s,en)=>s+(tutAllTutors.find(t=>t.name===en)?.hours||0),0);
    const projHours=projName||projExtra.length?projTotal:0;
    const mmiHours=(mmiName?mmiResults.find(r=>r.name===mmiName)?.totalHours||0:0)+mmiExtra.reduce((s,en)=>s+(mmiResults.find(r=>r.name===en)?.totalHours||0),0);
    const citHours=(citName?citizenshipTotals[citName]||0:0)+citExtra.reduce((s,en)=>s+(citizenshipTotals[en]||0),0);
    const resHours=(resName?resHoursTotals[resName]||0:0)+resExtra.reduce((s,en)=>s+(resHoursTotals[en]||0),0);
    const pgrHours=(pgrName?pgrHoursTotals[pgrName]||0:0)+pgrExtra.reduce((s,en)=>s+(pgrHoursTotals[en]||0),0);
    const assessmentHours=(assessmentName?assessmentHoursTotals[assessmentName]||0:0)+assessmentExtra.reduce((s,en)=>s+(assessmentHoursTotals[en]||0),0);
    const total=tlHours+assessmentHours+projHours+tutHours+mmiHours+citHours+resHours+pgrHours;
    const matchType=Object.keys(g.sources).length>1||g._extraSources?(g.matchType||'exact'):'only';
    const _bonuses=computeBonuses(g.canonical);
    return{canonical:g.canonical,tlName,assessmentName,projName,tutName,mmiName,citName,resName,pgrName,tlHours,assessmentHours,projHours,tutHours,mmiHours,citHours,resHours,pgrHours,total,matchType,score:g.score,sources:g.sources,_bonuses};
  });
  const maxTotal=Math.max(...combData.map(d=>d.total),1);
  const fuzzy=combData.filter(d=>d.matchType==='fuzzy').length;
  const firstname=combData.filter(d=>d.matchType==='firstname').length;
  const nMappings=manualMappings.size;
  document.getElementById('combMeta').textContent=`${combData.length} academics · ${combData.filter(d=>Object.keys(d.sources).length>1).length} matched across sources · ${fuzzy} fuzzy · ${firstname} first-name matches${nMappings?` · ${nMappings} manual mapping${nMappings>1?'s':''}`:''}`;
  // Show active mappings inline
  let mapInfo=document.getElementById('combMapInfo');
  if(!mapInfo){
    mapInfo=document.createElement('div');
    mapInfo.id='combMapInfo';
    mapInfo.style.cssText='margin-top:4px';
    document.getElementById('combMeta').after(mapInfo);
  }
  if(manualMappings.size>0){
      mapInfo.innerHTML='<span style="font-size:0.78rem;color:var(--teal);cursor:pointer" id="mapToggle">📌 <strong>'+manualMappings.size+' manual mapping'+(manualMappings.size>1?'s':'')+'</strong> <span style="font-size:0.68rem">(show)</span></span><div id="mapList" style="display:none;margin-top:6px;padding:8px;background:var(--light-blue);border-radius:6px;font-size:0.78rem"></div>';
      document.getElementById('mapToggle').addEventListener('click',()=>{
        const ml=document.getElementById('mapList');
        const show=ml.style.display==='none';
        ml.style.display=show?'block':'none';
        document.getElementById('mapToggle').querySelector('span').textContent=show?'(hide)':'(show)';
        if(show){
          ml.innerHTML=[...manualMappings.entries()].map(([nk,target])=>{
            // Find the canonical name from combData that has this normKey (best effort)
            const src=combData.find(d=>normKey(d.canonical)===nk)?.canonical||nk;
            const enc=encodeURIComponent(nk);
            return`<div style="display:flex;justify-content:space-between;align-items:center;padding:3px 0;border-bottom:1px solid var(--border)">
              <span><span style="color:var(--muted)">${src}</span> → <strong>${target}</strong></span>
              <button class="map-del-btn" data-nk="${enc}" style="background:none;border:1px solid var(--rust);color:var(--rust);border-radius:4px;padding:1px 7px;cursor:pointer;font-size:0.7rem">Delete</button>
            </div>`;
          }).join('');
          ml.querySelectorAll('.map-del-btn').forEach(b=>{
            b.addEventListener('click',()=>{
              manualMappings.delete(decodeURIComponent(b.dataset.nk));
              saveManualMappings();
              doMerge();
            });
          });
        }
      });
    }else{
      mapInfo.innerHTML='';
    }
  const mi=document.getElementById('combMergeInfo');
  const warnings=[];
  if(fuzzy>0)warnings.push(`<strong>⚠ ${fuzzy} fuzzy match${fuzzy>1?'es':''}</strong> — matched on name similarity`);
  if(firstname>0)warnings.push(`<strong>⚠ ${firstname} first-name match${firstname>1?'es':''}</strong> — matched on unique first name from MMI data (verify these)`);
  if(warnings.length>0){mi.innerHTML=warnings.join(' &nbsp;·&nbsp; ')+' — click flagged rows to inspect source names.';mi.style.display='';}else mi.style.display='none';
  document.getElementById('badge-combined').textContent=combData.length+' academics';
  document.getElementById('comb-result').style.display='block';
  reattachStoredTags();
  recomputeCombData();
  renderRulesEditor();
  combRender(maxTotal);
}

document.getElementById('combMergeBtn').addEventListener('click',doMerge);

function combGetSorted(){
  const q=document.getElementById('combSearch').value.toLowerCase();
  let data=combData.filter(d=>d.canonical.toLowerCase().includes(q));
  if(activeTagFilter!==null) data=data.filter(d=>activeTagsForPerson(d.canonical).includes(activeTagFilter));
  const[col,dir]=combSortKey.split('-');
  data.sort((a,b)=>{
    if(col==='name')return dir==='asc'?a.canonical.localeCompare(b.canonical):b.canonical.localeCompare(a.canonical);
    if(col==='fte'){const ap=ftePct(a.canonical,a.total),bp=ftePct(b.canonical,b.total);return dir==='asc'?ap-bp:bp-ap;}
    const av=col==='teaching'?a.tlHours:col==='assessment'?a.assessmentHours:col==='project'?a.projHours:col==='tutorial'?a.tutHours:col==='mmi'?a.mmiHours:col==='citizenship'?a.citHours:col==='research'?(a.resHours||0):col==='pgr'?a.pgrHours:a.total;
    const bv=col==='teaching'?b.tlHours:col==='assessment'?b.assessmentHours:col==='project'?b.projHours:col==='tutorial'?b.tutHours:col==='mmi'?b.mmiHours:col==='citizenship'?b.citHours:col==='research'?(b.resHours||0):col==='pgr'?b.pgrHours:b.total;
    return dir==='asc'?av-bv:bv-av;
  });
  return data;
}

let combSelected=new Set();
function combUpdateDetailBtn(){
  document.getElementById('combDetailBtn').disabled=combSelected.size===0;
  document.getElementById('combDetailBtn').textContent=combSelected.size>0?`📄 Detailed Report (${combSelected.size})`:'📄 Detailed Report';
  const cBtn=document.getElementById('combCombinedBtn');
  if(cBtn){cBtn.disabled=combSelected.size===0;cBtn.textContent=combSelected.size>0?`📄 Combined PDF (${combSelected.size})`:'📄 Combined PDF';}
}

// ── Tag & Rule helpers ────────────────────────────────────────────────────────

function todayDate(){const d=new Date();d.setHours(0,0,0,0);return d;}

function purgeExpiredAssignments(){
  const today=todayDate();
  staffTags.forEach((tagMap,canonical)=>{
    tagMap.forEach((info,tag)=>{
      if(info.expiry&&info.expiry<today){
        const rule=tagRules.get(tag);
        if(rule&&rule.fte!=null&&rule.fte!==1){
          const nk=normKey(canonical);
          const manual=staffFte.get(nk);
          if(manual!=null&&Math.abs(manual-rule.fte)<0.005)staffFte.delete(nk);
        }
        tagMap.delete(tag);
      }
    });
    if(tagMap.size===0)staffTags.delete(canonical);
  });
  if(activeTagFilter!==null&&!allTagsSorted().includes(activeTagFilter))activeTagFilter=null;
}

function allTagsSorted(){
  const s=new Set();
  staffTags.forEach(tagMap=>tagMap.forEach((_,t)=>s.add(t)));
  tagRules.forEach((_,t)=>s.add(t));
  return[...s].sort();
}

function tagsForPerson(canonical){
  return staffTags.get(canonical)||new Map();
}

function activeTagsForPerson(canonical){
  const today=todayDate();
  const out=[];
  (staffTags.get(canonical)||new Map()).forEach((info,tag)=>{
    if(!info.expiry||info.expiry>=today)out.push(tag);
  });
  return out;
}

function addTag(canonical,tag,expiry=null){
  tag=tag.trim();if(!tag)return;
  if(!staffTags.has(canonical))staffTags.set(canonical,new Map());
  staffTags.get(canonical).set(tag,{expiry});
}

function removeTag(canonical,tag){
  staffTags.get(canonical)?.delete(tag);
  if(staffTags.get(canonical)?.size===0)staffTags.delete(canonical);
  const rule=tagRules.get(tag);
  if(rule&&rule.fte!=null&&rule.fte!==1){
    const nk=normKey(canonical);
    const manual=staffFte.get(nk);
    if(manual!=null&&Math.abs(manual-rule.fte)<0.005)staffFte.delete(nk);
  }
  if(activeTagFilter===tag&&!allTagsSorted().includes(tag))activeTagFilter=null;
}

function setTagExpiry(canonical,tag,expiry){
  if(staffTags.get(canonical)?.has(tag)){
    staffTags.get(canonical).get(tag).expiry=expiry;
  }
}

function computeBonuses(canonical){
  const today=todayDate();
  let tlLoad=0,tlPrep=0,proj=0;
  activeTagsForPerson(canonical).forEach(tag=>{
    const rule=tagRules.get(tag);
    if(!rule)return;
    if(rule.expiry&&rule.expiry<today)return;
    tlLoad+=rule.tlLoad||0;
    tlPrep+=rule.tlPrep||0;
    proj+=rule.proj||0;
  });
  return{tlLoad,tlPrep,proj};
}

function applyBonuses(canonical,contactHours,projHoursBase){
  const b=computeBonuses(canonical);
  const loadMult=1+b.tlLoad;
  const prepRatio=tlPrepRatio+b.tlPrep;
  const tlTotal=contactHours*loadMult*(1+prepRatio);
  const projTotal=projHoursBase*(1+b.proj);
  return{tlTotal,projTotal,bonuses:b};
}

// ── Module tag helpers ──────────────────────────────────────────────────
function modTagsFor(moduleName){return moduleTags.get(normKey(moduleName))||new Set();}
function addModuleTag(moduleName,tag){
  tag=tag.trim();if(!tag)return;
  const nk=normKey(moduleName);
  if(!moduleTags.has(nk))moduleTags.set(nk,new Set());
  moduleTags.get(nk).add(tag);
  saveModuleTags();renderModuleTagFilterBar();tlRenderModGrid();
}
function removeModuleTag(moduleName,tag){
  const nk=normKey(moduleName);
  moduleTags.get(nk)?.delete(tag);
  if(moduleTags.get(nk)?.size===0)moduleTags.delete(nk);
  if(moduleTagFilter&&!allModuleTags().includes(moduleTagFilter))moduleTagFilter=null;
  saveModuleTags();renderModuleTagFilterBar();tlRenderModGrid();
}
function allModuleTags(){
  const s=new Set();
  moduleTags.forEach(tags=>tags.forEach(t=>s.add(t)));
  return[...s].sort();
}
function modTagFilteredModules(){
  if(!moduleTagFilter)return tlAllModules;
  return tlAllModules.filter(m=>modTagsFor(m).has(moduleTagFilter));
}

function renderModuleTagChips(){
  const wrap=document.getElementById('tlModGridWrap');
  if(!wrap)return;
  wrap.querySelectorAll('.name-cell[data-entity]').forEach(cell=>{
    const mn=decodeURIComponent(cell.dataset.entity);
    if(mn==='all')return;
    cell.style.overflow='visible';cell.style.maxWidth='none';cell.style.textOverflow='clip';
    const tags=modTagsFor(mn);
    const row=document.createElement('span');
    row.style.cssText='display:inline-flex;flex-wrap:wrap;gap:3px;align-items:center;margin-left:6px';
    row.className='mod-tag-row';
    for(const t of tags){
      const chip=document.createElement('span');chip.className='tag-chip';
      chip.textContent=t;
      const x=document.createElement('span');x.className='tag-x';x.textContent='×';
      x.dataset.mod=encodeURIComponent(mn);x.dataset.tag=encodeURIComponent(t);
      chip.appendChild(x);row.appendChild(chip);
    }
    const addBtn=document.createElement('button');addBtn.className='tag-add-btn mod-tag-add';
    addBtn.textContent='+ tag';addBtn.dataset.mod=encodeURIComponent(mn);
    row.appendChild(addBtn);
    cell.appendChild(row);
  });
  wrap.querySelectorAll('.mod-tag-row .tag-x').forEach(x=>{
    x.addEventListener('click',e=>{e.stopPropagation();removeModuleTag(decodeURIComponent(x.dataset.mod),decodeURIComponent(x.dataset.tag));});
  });
  wrap.querySelectorAll('.mod-tag-add').forEach(btn=>{
    btn.addEventListener('click',e=>{e.stopPropagation();openModuleTagPopover(decodeURIComponent(btn.dataset.mod),btn);});
  });
}
function renderModuleTagFilterBar(){
  const all=allModuleTags();
  const pills=document.getElementById('modTagFilterPills');
  const clearBtn=document.getElementById('modTagFilterClear');
  if(!pills)return;
  pills.innerHTML=all.map(t=>`<button class="tag-filter-pill${moduleTagFilter===t?' active':''}" data-modtag="${encodeURIComponent(t)}">${t}<span class="tfc">${tlAllModules.filter(m=>modTagsFor(m).has(t)).length}</span></button>`).join('');
  pills.querySelectorAll('.tag-filter-pill').forEach(btn=>{btn.addEventListener('click',()=>{const t=decodeURIComponent(btn.dataset.modtag);moduleTagFilter=moduleTagFilter===t?null:t;renderModuleTagFilterBar();tlRenderModGrid();});});
  if(clearBtn)clearBtn.style.display=moduleTagFilter?'':'none';
}

// ── Module tag popover ────────────────────────────────────────────────────────
let modTagPopoverModule=null;
function openModuleTagPopover(moduleName,anchorEl){
  modTagPopoverModule=moduleName;
  const pop=document.getElementById('modTagPopover');
  document.getElementById('modTagPopoverTitle').textContent=moduleName;
  document.getElementById('modTagPopoverInput').value='';
  renderModuleTagPopoverContent();
  const rect=anchorEl.getBoundingClientRect();
  pop.style.display='block';
  const pw=240,ph=200;
  let left=rect.left,top=rect.bottom+6;
  if(left+pw>window.innerWidth-8)left=window.innerWidth-pw-8;
  if(top+ph>window.innerHeight-8)top=rect.top-ph-6;
  pop.style.left=Math.max(8,left)+'px';
  pop.style.top=Math.max(8,top)+'px';
  setTimeout(()=>document.getElementById('modTagPopoverInput').focus(),50);
}
function renderModuleTagPopoverContent(){
  const mn=modTagPopoverModule;if(!mn)return;
  const tags=modTagsFor(mn);
  const existingEl=document.getElementById('modTagPopoverExisting');
  if(tags.size===0){
    existingEl.innerHTML='<span style="font-size:0.75rem;color:var(--muted)">No tags yet</span>';
  }else{
    existingEl.innerHTML=[...tags].map(t=>{
      const enc=encodeURIComponent(t);
      return`<div class="tag-assign-row">
        <span class="tag-assign-name">${t}</span>
        <button class="tag-assign-x" data-tag="${enc}" title="Remove tag">×</button>
      </div>`;
    }).join('');
    existingEl.querySelectorAll('.tag-assign-x').forEach(x=>{
      x.addEventListener('click',()=>{removeModuleTag(mn,decodeURIComponent(x.dataset.tag));renderModuleTagPopoverContent();renderModuleTagFilterBar();});
    });
  }
  const suggestions=allModuleTags().filter(t=>!tags.has(t));
  document.getElementById('modTagPopoverSuggestions').innerHTML=suggestions.map(t=>`<button class="tag-add-btn" data-tag="${encodeURIComponent(t)}" style="font-size:0.72rem">${t}</button>`).join('');
  document.querySelectorAll('#modTagPopoverSuggestions .tag-add-btn').forEach(btn=>{
    btn.addEventListener('click',()=>{addModuleTag(mn,decodeURIComponent(btn.dataset.tag));renderModuleTagPopoverContent();renderModuleTagFilterBar();});
  });
}
function closeModuleTagPopover(){document.getElementById('modTagPopover').style.display='none';modTagPopoverModule=null;}
document.getElementById('modTagPopoverClose').addEventListener('click',closeModuleTagPopover);
document.getElementById('modTagPopoverAdd').addEventListener('click',()=>{
  const val=document.getElementById('modTagPopoverInput').value.trim();
  if(val&&modTagPopoverModule){addModuleTag(modTagPopoverModule,val);document.getElementById('modTagPopoverInput').value='';renderModuleTagPopoverContent();renderModuleTagFilterBar();}
});
document.getElementById('modTagPopoverInput').addEventListener('keydown',e=>{
  if(e.key==='Enter'){const val=e.target.value.trim();if(val&&modTagPopoverModule){addModuleTag(modTagPopoverModule,val);e.target.value='';renderModuleTagPopoverContent();renderModuleTagFilterBar();}}
  if(e.key==='Escape')closeModuleTagPopover();
});
document.addEventListener('click',e=>{
  const pop=document.getElementById('modTagPopover');
  if(pop.style.display!=='none'&&!pop.contains(e.target)&&!e.target.closest('.mod-tag-add'))closeModuleTagPopover();
});

// ── Rules editor ──────────────────────────────────────────────────────────────
function renderRulesEditor(){
  const today=todayDate();
  const tbody=document.getElementById('rulesTbody');
  if(tagRules.size===0){
    tbody.innerHTML=`<tr><td colspan="7" style="color:var(--muted);font-style:italic;padding:0.8rem">No rules defined yet — add one below.</td></tr>`;
  }else{
    tbody.innerHTML=[...tagRules.entries()].map(([tag,rule])=>{
      const expired=rule.expiry&&rule.expiry<today;
      const expiryStr=rule.expiry?rule.expiry.toISOString().slice(0,10):'';
      const badge=expired?`<span class="rule-expired-badge">expired</span>`:`<span class="rule-active-badge">active</span>`;
      const enc=encodeURIComponent(tag);
      return`<tr data-rule-tag="${enc}"${expired?' style="opacity:0.55"':''}>
        <td><span class="rule-tag-name">${tag}</span> ${badge}</td>
        <td><input type="number" class="rule-tl" data-tag="${enc}" value="${rule.tlLoad||0}" min="0" step="0.1" style="width:62px"></td>
        <td><input type="number" class="rule-tp" data-tag="${enc}" value="${rule.tlPrep||0}" min="0" step="0.1" style="width:62px"></td>
        <td><input type="number" class="rule-proj" data-tag="${enc}" value="${rule.proj||0}" min="0" step="0.1" style="width:62px"></td>
        <td><input type="number" class="rule-fte" data-tag="${enc}" value="${rule.fte??1}" min="0.1" max="1" step="0.05" style="width:62px" title="FTE fraction: multiplies into personal target. 0.75 = 75% of base target."></td>
        <td><input type="date" class="rule-expiry" data-tag="${enc}" value="${expiryStr}"></td>
        <td><button class="rule-del" data-tag="${enc}" title="Delete rule">🗑</button></td>
      </tr>`;
    }).join('');
    tbody.querySelectorAll('.rule-tl,.rule-tp,.rule-proj,.rule-fte').forEach(inp=>{
      inp.addEventListener('change',()=>{
        const tag=decodeURIComponent(inp.dataset.tag);
        const row=tbody.querySelector(`tr[data-rule-tag="${inp.dataset.tag}"]`);
        const r=tagRules.get(tag);if(!r)return;
        r.tlLoad=+row.querySelector('.rule-tl').value||0;
        r.tlPrep=+row.querySelector('.rule-tp').value||0;
        r.proj=+row.querySelector('.rule-proj').value||0;
        r.fte=+row.querySelector('.rule-fte').value||1;
        recomputeCombData();saveTagState();combRender();
      });
    });
    tbody.querySelectorAll('.rule-expiry').forEach(inp=>{
      inp.addEventListener('change',()=>{
        const tag=decodeURIComponent(inp.dataset.tag);
        const r=tagRules.get(tag);if(!r)return;
        r.expiry=inp.value?new Date(inp.value):null;
        renderRulesEditor();recomputeCombData();saveTagState();combRender();
      });
    });
    tbody.querySelectorAll('.rule-del').forEach(btn=>{
      btn.addEventListener('click',()=>{
        tagRules.delete(decodeURIComponent(btn.dataset.tag));
        renderRulesEditor();renderTagFilterBar();recomputeCombData();saveTagState();combRender();
      });
    });
  }
  const activeRules=[...tagRules.values()].filter(r=>!r.expiry||r.expiry>=today).length;
  document.getElementById('rulesActiveCount').textContent=tagRules.size>0?`(${activeRules} active, ${tagRules.size} total)`:'';
  document.getElementById('ruleTagSuggestions').innerHTML=allTagsSorted().map(t=>`<option value="${t}">`).join('');
}

document.getElementById('rulesPanelHdr').addEventListener('click',()=>{
  document.getElementById('rulesPanelBody').classList.toggle('open');
  document.getElementById('rulesPanelHdr').classList.toggle('open');
});

document.getElementById('ruleAddBtn').addEventListener('click',()=>{
  const tag=document.getElementById('ruleNewTag').value.trim();
  if(!tag){document.getElementById('ruleNewTag').focus();return;}
  const tlLoad=+document.getElementById('ruleNewTL').value||0;
  const tlPrep=+document.getElementById('ruleNewTP').value||0;
  const proj=+document.getElementById('ruleNewProj').value||0;
  const fte=+document.getElementById('ruleNewFte').value||1;
  const expiryVal=document.getElementById('ruleNewExpiry').value;
  const expiry=expiryVal?new Date(expiryVal):null;
  tagRules.set(tag,{tlLoad,tlPrep,proj,fte,expiry});
  document.getElementById('ruleNewTag').value='';
  document.getElementById('ruleNewTL').value='0';
  document.getElementById('ruleNewTP').value='0';
  document.getElementById('ruleNewProj').value='0';
  document.getElementById('ruleNewFte').value='1';
  document.getElementById('ruleNewExpiry').value='';
  renderRulesEditor();renderTagFilterBar();recomputeCombData();saveTagState();combRender();
});

function renderTagFilterBar(){
  const today=todayDate();
  const all=allTagsSorted();
  const pills=document.getElementById('tagFilterPills');
  pills.innerHTML=all.map(t=>{
    const count=[...combData].filter(d=>activeTagsForPerson(d.canonical).includes(t)).length;
    const rule=tagRules.get(t);
    const hasActiveRule=rule&&(!rule.expiry||rule.expiry>=today);
    const ruleIcon=hasActiveRule?` ⚖️`:'';
    return`<button class="tag-filter-pill${activeTagFilter===t?' active':''}" data-tag="${encodeURIComponent(t)}">${t}${ruleIcon}<span class="tfc">${count}</span></button>`;
  }).join('');
  pills.querySelectorAll('.tag-filter-pill').forEach(btn=>{
    btn.addEventListener('click',()=>{
      const t=decodeURIComponent(btn.dataset.tag);
      activeTagFilter=activeTagFilter===t?null:t;
      renderTagFilterBar();combRender();
    });
  });
  const clearBtn=document.getElementById('tagFilterClear');
  clearBtn.style.display=activeTagFilter?'':'none';
  const grpBtn=document.getElementById('tagGroupReportBtn');
  grpBtn.classList.toggle('visible',activeTagFilter!==null);
  if(activeTagFilter){grpBtn.textContent=`📄 Report for "${activeTagFilter}"`;}
}

// ── Tag popover ───────────────────────────────────────────────────────────────
let tagPopoverCanonical=null;
function openTagPopover(canonical,anchorEl){
  tagPopoverCanonical=canonical;
  const pop=document.getElementById('tagPopover');
  document.getElementById('tagPopoverTitle').textContent=canonical;
  document.getElementById('tagPopoverInput').value='';
  const manualFte=staffFte.get(normKey(canonical));
  document.getElementById('tagPopoverFte').value=manualFte!=null?manualFte.toFixed(2):'';
  renderTagPopoverContent();
  const rect=anchorEl.getBoundingClientRect();
  pop.style.display='block';
  const pw=260,ph=220;
  let left=rect.left,top=rect.bottom+6;
  if(left+pw>window.innerWidth-8)left=window.innerWidth-pw-8;
  if(top+ph>window.innerHeight-8)top=rect.top-ph-6;
  pop.style.left=Math.max(8,left)+'px';
  pop.style.top=Math.max(8,top)+'px';
  setTimeout(()=>document.getElementById('tagPopoverInput').focus(),50);
}
function renderTagPopoverContent(){
  const canonical=tagPopoverCanonical;if(!canonical)return;
  const today=todayDate();
  const myTagMap=tagsForPerson(canonical);
  const existingEl=document.getElementById('tagPopoverExisting');
  if(myTagMap.size===0){
    existingEl.innerHTML='<span style="font-size:0.75rem;color:var(--muted)">No tags yet</span>';
  }else{
    existingEl.innerHTML=[...myTagMap.entries()].map(([t,info])=>{
      const enc=encodeURIComponent(t);
      const expired=info.expiry&&info.expiry<today;
      const expiryStr=info.expiry?info.expiry.toISOString().slice(0,10):'';
      const rule=tagRules.get(t);
      const hasRule=rule&&(!rule.expiry||rule.expiry>=today);
      return`<div class="tag-assign-row${expired?' rule-expired-badge':''}" style="${expired?'opacity:0.5':''}">
        <span class="tag-assign-name">${t}${hasRule?' ⚖️':''}</span>
        <span class="tag-assign-expiry">
          <label style="font-size:0.68rem;color:var(--muted)">expires:</label>
          <input type="date" class="tag-assign-expiry-input" data-tag="${enc}" value="${expiryStr}" title="Assignment expiry — leaves tag definition intact">
        </span>
        <button class="tag-assign-x" data-tag="${enc}" title="Remove tag">×</button>
      </div>`;
    }).join('');
    existingEl.querySelectorAll('.tag-assign-x').forEach(x=>{
      x.addEventListener('click',()=>{removeTag(canonical,decodeURIComponent(x.dataset.tag));recomputeCombData();renderTagPopoverContent();renderTagFilterBar();renderRulesEditor();saveTagState();combRender();});
    });
    existingEl.querySelectorAll('.tag-assign-expiry-input').forEach(inp=>{
      inp.addEventListener('change',()=>{
        const expiry=inp.value?new Date(inp.value):null;
        setTagExpiry(canonical,decodeURIComponent(inp.dataset.tag),expiry);
        recomputeCombData();renderTagFilterBar();saveTagState();combRender();
      });
    });
  }
  const suggestions=allTagsSorted().filter(t=>!myTagMap.has(t));
  document.getElementById('tagPopoverSuggestions').innerHTML=suggestions.map(t=>`<button class="tag-add-btn" data-tag="${encodeURIComponent(t)}" style="font-size:0.72rem">${t}</button>`).join('');
  document.querySelectorAll('#tagPopoverSuggestions .tag-add-btn').forEach(btn=>{
    btn.addEventListener('click',()=>{addTag(canonical,decodeURIComponent(btn.dataset.tag));recomputeCombData();renderTagPopoverContent();renderTagFilterBar();renderRulesEditor();saveTagState();combRender();});
  });
}
function closeTagPopover(){document.getElementById('tagPopover').style.display='none';tagPopoverCanonical=null;}
document.getElementById('tagPopoverClose').addEventListener('click',closeTagPopover);
document.getElementById('tagPopoverFte').addEventListener('change',e=>{
  if(!tagPopoverCanonical)return;
  const val=e.target.value.trim();
  setFte(tagPopoverCanonical,val===''?null:+val);
  saveTagState();combRender();
});
document.getElementById('tagPopoverAdd').addEventListener('click',()=>{
  const val=document.getElementById('tagPopoverInput').value.trim();
  if(val&&tagPopoverCanonical){addTag(tagPopoverCanonical,val);document.getElementById('tagPopoverInput').value='';recomputeCombData();renderTagPopoverContent();renderTagFilterBar();renderRulesEditor();saveTagState();combRender();}
});
document.getElementById('tagPopoverInput').addEventListener('keydown',e=>{
  if(e.key==='Enter'){const val=e.target.value.trim();if(val&&tagPopoverCanonical){addTag(tagPopoverCanonical,val);e.target.value='';recomputeCombData();renderTagPopoverContent();renderTagFilterBar();renderRulesEditor();saveTagState();combRender();}}
  if(e.key==='Escape')closeTagPopover();
});
document.addEventListener('click',e=>{
  const pop=document.getElementById('tagPopover');
  if(pop.style.display!=='none'&&!pop.contains(e.target)&&!e.target.closest('.tag-add-btn'))closeTagPopover();
});
document.getElementById('tagFilterClear').addEventListener('click',()=>{activeTagFilter=null;renderTagFilterBar();combRender();});
document.getElementById('tagGroupReportBtn').addEventListener('click',()=>{
  if(!activeTagFilter)return;
  const group=combData.filter(d=>activeTagsForPerson(d.canonical).includes(activeTagFilter)).map(d=>d.canonical);
  if(group.length>0)generateDetailedReport(group);
});

// ── Manual name mapping popover ──────────────────────────────────────────────
let mapPopoverCanonical=null;

function ensureMapPopover(){
  if(document.getElementById('mapPopover'))return;
  const div=document.createElement('div');
  div.id='mapPopover';
  div.style.cssText='display:none;position:fixed;z-index:10000;background:white;border:1px solid var(--border);border-radius:10px;box-shadow:0 8px 30px rgba(0,0,0,0.15);padding:14px;width:280px;font-size:0.82rem';
  div.innerHTML=`<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px">
    <strong id="mapPopoverTitle" style="font-size:0.9rem"></strong>
    <button id="mapPopoverClose" style="background:none;border:none;font-size:1.2rem;cursor:pointer;color:var(--muted);line-height:1">×</button>
  </div>
  <div style="margin-bottom:6px;color:var(--muted);font-size:0.75rem">Map this name to:</div>
  <input id="mapPopoverSearch" type="text" placeholder="Search staff..." style="width:100%;padding:6px 8px;border:1px solid var(--border);border-radius:6px;font-size:0.82rem;margin-bottom:6px;box-sizing:border-box">
  <div id="mapPopoverList" style="max-height:200px;overflow-y:auto"></div>
  <div id="mapPopoverExisting" style="margin-top:6px;padding-top:6px;border-top:1px solid var(--border);font-size:0.75rem;color:var(--muted)"></div>`;
  document.body.appendChild(div);
  document.getElementById('mapPopoverClose').addEventListener('click',closeMapPopover);
  document.getElementById('mapPopoverSearch').addEventListener('input',renderMapPopoverList);
  document.addEventListener('click',e=>{
    const pop=document.getElementById('mapPopover');
    if(pop&&pop.style.display!=='none'&&!pop.contains(e.target)&&!e.target.closest('.map-btn'))closeMapPopover();
  });
}

function openMapPopover(canonical,anchorEl){
  ensureMapPopover();
  mapPopoverCanonical=canonical;
  document.getElementById('mapPopoverTitle').textContent=canonical;
  document.getElementById('mapPopoverSearch').value='';

  const fromNk=normKey(canonical);
  const existingTarget=manualMappings.get(fromNk);
  const existingEl=document.getElementById('mapPopoverExisting');
  if(existingTarget){
    existingEl.innerHTML=`Mapped to: <strong>${existingTarget}</strong> · <button class="map-unmap-btn" style="background:none;border:1px solid var(--rust);color:var(--rust);border-radius:4px;padding:2px 8px;cursor:pointer;font-size:0.72rem">Remove mapping</button>`;
    existingEl.querySelector('.map-unmap-btn').addEventListener('click',()=>{
      manualMappings.delete(fromNk);
      saveManualMappings();
      closeMapPopover();
      doMerge();
    });
  }else{
    existingEl.innerHTML='';
  }

  renderMapPopoverList();

  const pop=document.getElementById('mapPopover');
  const rect=anchorEl.getBoundingClientRect();
  pop.style.display='block';
  let left=rect.left,top=rect.bottom+6;
  if(left+280>window.innerWidth-8)left=window.innerWidth-288;
  if(top+260>window.innerHeight-8)top=rect.top-260;
  pop.style.left=Math.max(8,left)+'px';
  pop.style.top=Math.max(8,top)+'px';
  setTimeout(()=>document.getElementById('mapPopoverSearch').focus(),50);
}

function renderMapPopoverList(){
  const q=document.getElementById('mapPopoverSearch').value.toLowerCase();
  const fromNk=normKey(mapPopoverCanonical);
  const targets=combData.filter(d=>normKey(d.canonical)!==fromNk).filter(d=>d.canonical.toLowerCase().includes(q)).map(d=>d.canonical);

  const list=document.getElementById('mapPopoverList');
  if(targets.length===0){
    list.innerHTML='<div style="color:var(--muted);padding:8px 0;text-align:center">No matching staff</div>';
    return;
  }
  list.innerHTML=targets.map(t=>`<button class="map-target-btn" data-target="${encodeURIComponent(t)}" style="display:block;width:100%;text-align:left;padding:6px 8px;border:none;background:transparent;border-radius:6px;cursor:pointer;font-size:0.82rem">${t}</button>`).join('');
  list.querySelectorAll('.map-target-btn').forEach(btn=>{
    btn.addEventListener('click',()=>{
      manualMappings.set(fromNk,decodeURIComponent(btn.dataset.target));
      saveManualMappings();
      closeMapPopover();
      doMerge();
    });
    btn.addEventListener('mouseenter',()=>btn.style.background='var(--light-blue)');
    btn.addEventListener('mouseleave',()=>btn.style.background='transparent');
  });
}

function closeMapPopover(){
  const pop=document.getElementById('mapPopover');
  if(pop)pop.style.display='none';
  mapPopoverCanonical=null;
}

// Wire up map buttons on each render
function wireMapButtons(){
  document.querySelectorAll('.map-btn').forEach(btn=>{
    btn.addEventListener('click',e=>{e.stopPropagation();openMapPopover(decodeURIComponent(btn.dataset.canonical),btn);});
  });
}

function combRender(maxTotal){
  if(!maxTotal)maxTotal=Math.max(...combData.map(d=>d.total),1);
  const data=combGetSorted();
  const visMax=Math.max(...data.map(d=>d.total),1);
  const matchBadge=(d,enc)=>{
    const mapBtn=` <button class="map-btn" data-canonical="${enc}" title="Manually map this name to another staff member">map</button>`;
    if(d.matchType==='fuzzy')return`<span class="match-badge match-fuzzy">fuzzy${d.score?` ${Math.round(d.score*100)}%`:''}</span>${mapBtn}`;
    if(d.matchType==='firstname')return`<span class="match-badge match-fuzzy">first name</span>${mapBtn}`;
    if(d.matchType==='only'){const src=Object.keys(d.sources)[0];return`<span class="match-badge match-only">${SRC_LABELS[src]||src} only</span>${mapBtn}`;}
    return`<span class="match-badge match-exact">✓ matched</span>${mapBtn}`;
  };
  document.getElementById('combTbody').innerHTML=data.map(d=>{
    const enc=encodeURIComponent(d.canonical);
    const chk=combSelected.has(d.canonical)?'checked':'';
    const myTagMap=tagsForPerson(d.canonical);
    const tagHtml=[...myTagMap.entries()].map(([t,info])=>{
      const today=todayDate();
      const expired=info.expiry&&info.expiry<today;
      if(expired)return'';
      return`<span class="tag-chip">${t}<span class="tag-x" data-canonical="${enc}" data-tag="${encodeURIComponent(t)}" onclick="event.stopPropagation()">×</span></span>`;
    }).join('');
    const b=d._bonuses||{tlLoad:0,tlPrep:0,proj:0};
    const hasTLAdj=b.tlLoad!==0||b.tlPrep!==0;
    const hasProjAdj=b.proj!==0;
    const adjTip=(hasTLAdj||hasProjAdj)?`title="Adjustments active: ${hasTLAdj?`teaching load +${b.tlLoad}, prep +${b.tlPrep}`:''}${hasTLAdj&&hasProjAdj?'; ':''}${hasProjAdj?`project +${b.proj}`:''}"`:'';
    return`<tr data-canonical="${enc}">
    <td style="text-align:center;padding:8px 6px"><input type="checkbox" class="comb-chk" data-canonical="${enc}" ${chk} onclick="event.stopPropagation()"></td>
    <td class="cn" style="cursor:pointer">${d.canonical}${hasTLAdj||hasProjAdj?`<span class="adj-badge" ${adjTip}>⚖️ adj</span>`:''}</td>
    <td class="tag-cell">${tagHtml}<button class="tag-add-btn comb-tag-add" data-canonical="${enc}" onclick="event.stopPropagation()">+ tag</button></td>
    <td class="num">${d.tlHours>0?d.tlHours.toFixed(1):'—'}</td>
    <td class="num">${d.assessmentHours>0?d.assessmentHours.toFixed(1):'—'}</td>
    <td class="num">${d.projHours>0?d.projHours.toFixed(1):'—'}</td>
    <td class="num">${d.tutHours>0?d.tutHours.toFixed(1):'—'}</td>
    <td class="num">${d.mmiHours>0?d.mmiHours.toFixed(1):'—'}</td>
    <td class="num">${d.citHours>0?d.citHours.toFixed(1):'—'}</td>
    <td class="num">${(d.resHours||0)>0?(d.resHours).toFixed(1):'—'}</td>
    <td class="num">${d.pgrHours>0?d.pgrHours.toFixed(1):'—'}</td>
    <td class="tot">${d.total.toFixed(1)}</td>
    <td>${fteBarHtml(d.canonical,d.total)}</td>
    <td>${matchBadge(d,enc)}</td>
  </tr>`;}).join('');
  const totTL=data.reduce((s,d)=>s+d.tlHours,0),totAssessment=data.reduce((s,d)=>s+d.assessmentHours,0),totProj=data.reduce((s,d)=>s+d.projHours,0),totTut=data.reduce((s,d)=>s+d.tutHours,0),totMmi=data.reduce((s,d)=>s+d.mmiHours,0),totCit=data.reduce((s,d)=>s+d.citHours,0),totRes=data.reduce((s,d)=>s+(d.resHours||0),0),totPgr=data.reduce((s,d)=>s+d.pgrHours,0),totAll=data.reduce((s,d)=>s+d.total,0);
  const avgFte=data.length>0?Math.round(data.reduce((s,d)=>s+ftePct(d.canonical,d.total),0)/data.length):0;
  const avgCls=fteClass(avgFte);
  const filterNote=activeTagFilter?` <span style="font-size:0.72rem;font-weight:400;color:var(--gold);margin-left:6px">tag: ${activeTagFilter} (${data.length})</span>`:'';
  document.getElementById('combFoot').innerHTML=`<tr><td></td><td class="cn">Total${filterNote}</td><td></td><td class="num">${totTL.toFixed(1)}</td><td class="num">${totAssessment.toFixed(1)}</td><td class="num">${totProj.toFixed(1)}</td><td class="num">${totTut.toFixed(1)}</td><td class="num">${totMmi.toFixed(1)}</td><td class="num">${totCit.toFixed(1)}</td><td class="num">${totRes.toFixed(1)}</td><td class="num">${totPgr.toFixed(1)}</td><td class="tot">${totAll.toFixed(1)}</td><td><span style="font-size:0.78rem;font-weight:600" class="fte-pct ${avgCls}">avg ${avgFte}%</span></td><td></td></tr>`;
  document.querySelectorAll('#combTbody .tag-x').forEach(x=>{
    x.addEventListener('click',e=>{e.stopPropagation();const c=decodeURIComponent(x.dataset.canonical),t=decodeURIComponent(x.dataset.tag);removeTag(c,t);recomputeCombData();renderTagFilterBar();renderRulesEditor();saveTagState();combRender();});
  });
  document.querySelectorAll('.comb-tag-add').forEach(btn=>{
    btn.addEventListener('click',e=>{e.stopPropagation();openTagPopover(decodeURIComponent(btn.dataset.canonical),btn);});
  });
  document.querySelectorAll('.comb-chk').forEach(chk=>{chk.addEventListener('change',()=>{const c=decodeURIComponent(chk.dataset.canonical);if(chk.checked)combSelected.add(c);else combSelected.delete(c);combUpdateDetailBtn();const all=document.querySelectorAll('.comb-chk');const allChecked=[...all].every(c=>c.checked);document.getElementById('combSelectAll').checked=allChecked;document.getElementById('combSelectAll').indeterminate=!allChecked&&[...all].some(c=>c.checked);});});
  document.querySelectorAll('#combTbody tr').forEach(row=>{row.querySelector('.cn').addEventListener('click',()=>{
    const canonical=decodeURIComponent(row.dataset.canonical),d=combData.find(x=>x.canonical===canonical);if(!d)return;
    const tutor=d.tutName?tutAllTutors.find(t=>t.name===d.tutName):null;
    const proj=d.projName?projAllResults.find(r=>r.name===d.projName):null;
    const mmiR=d.mmiName?mmiResults.find(r=>r.name===d.mmiName):null;
    let html=`<div class="panel-section"><h4>Load Summary</h4>
      ${d.tlHours>0?`<div class="panel-row"><span class="k">📅 Teaching</span><span class="v">${d.tlHours.toFixed(1)}h</span></div>`:''}
      ${d.assessmentHours>0?`<div class="panel-row"><span class="k">📝 Non-timetabled assess.</span><span class="v">${d.assessmentHours.toFixed(1)}h</span></div>`:''}
      ${d.projHours>0?`<div class="panel-row"><span class="k">🎓 Project supervision</span><span class="v">${d.projHours.toFixed(1)}h</span></div>`:''}
      ${d.tutHours>0?`<div class="panel-row"><span class="k">👥 Tutorial</span><span class="v">${d.tutHours.toFixed(1)}h</span></div>`:''}
      ${d.mmiHours>0?`<div class="panel-row"><span class="k">🩺 MMIs</span><span class="v">${d.mmiHours.toFixed(1)}h</span></div>`:''}
      ${d.citHours>0?`<div class="panel-row"><span class="k">🏛 Citizenship</span><span class="v">${d.citHours.toFixed(1)}h</span></div>`:''}
      ${(d.resHours||0)>0?`<div class="panel-row"><span class="k">🔬 Research</span><span class="v">${d.resHours.toFixed(1)}h</span></div>`:''}
      ${d.pgrHours>0?`<div class="panel-row"><span class="k">👨‍🎓 PGR Supervision</span><span class="v">${d.pgrHours.toFixed(1)}h</span></div>`:''}
      <div class="panel-row"><span class="k"><strong>Total</strong></span><span class="v big">${d.total.toFixed(1)}h</span></div>
      <div class="panel-row"><span class="k">FTE target</span><span class="v">${personalTarget(d.canonical).toFixed(0)}h (${(getFte(d.canonical)*100).toFixed(0)}% FTE)</span></div>
      <div class="panel-row"><span class="k">FTE %</span><span class="v ${fteClass(ftePct(d.canonical,d.total))}">${ftePct(d.canonical,d.total)}%</span></div>
    </div>`;
    if(Object.keys(d.sources).length>1){html+=`<div class="panel-section"><h4>Source Names</h4>${Object.entries(d.sources).map(([src,nm])=>`<div class="panel-row"><span class="k">${SRC_LABELS[src]||src}</span><span class="v" style="font-family:inherit;font-size:0.82rem">${nm}</span></div>`).join('')}</div>`;}
    if(tutor)html+=`<div class="panel-section"><h4>Tutorial Detail</h4><div class="panel-row"><span class="k">Y1 tutees</span><span class="v">${tutor.year1.length}</span></div><div class="panel-row"><span class="k">Other tutees</span><span class="v">${tutor.other.length}</span></div></div>`;
    if(proj)html+=`<div class="panel-section"><h4>Project Breakdown</h4><div class="panel-row"><span class="k">Projects supervised</span><span class="v">${proj.supervised.length} (${proj.nSup.toFixed?proj.nSup.toFixed(2):proj.nSup} share)</span></div><div class="panel-row"><span class="k">Co-supervised</span><span class="v">${proj.cosupervised.length} (${proj.nCoSup.toFixed?proj.nCoSup.toFixed(2):proj.nCoSup} share)</span></div><div class="panel-row"><span class="k">Diss. assessed</span><span class="v">${proj.nDissAss}</span></div><div class="panel-row"><span class="k">Poster assessed</span><span class="v">${proj.nPostAss}</span></div></div>`;
    if(mmiR){const activeSess=mmiR.sessions.filter(s=>!s.isReserve);html+=`<div class="panel-section"><h4>MMI Detail</h4><div class="panel-row"><span class="k">Active sessions</span><span class="v">${activeSess.length}</span></div><div class="panel-row"><span class="k">Total hours</span><span class="v">${fmt(mmiR.totalHours)}h</span></div></div>`;}
    openPanel(d.canonical,`${d.total.toFixed(1)}h total · ${Object.keys(d.sources).length} source${Object.keys(d.sources).length!==1?'s':''}`,html);
  });});
  combUpdateDetailBtn();
  wireMapButtons();
  renderTagFilterBar();
}

document.getElementById('combSearch').addEventListener('input',()=>combRender());
document.getElementById('combSort').addEventListener('change',e=>{combSortKey=e.target.value;combRender();});
document.querySelector('#comb-result table.comb-table thead').addEventListener('click',e=>{const th=e.target.closest('th[data-combsort]');if(!th)return;const col=th.dataset.combsort;const[curCol,curDir]=combSortKey.split('-');if(curCol===col)combSortKey=col+'-'+(curDir==='desc'?'asc':'desc');else combSortKey=col+'-'+(col==='name'?'asc':'desc');combRender();});
document.getElementById('combSelectAll').addEventListener('change',e=>{const checked=e.target.checked;document.querySelectorAll('.comb-chk').forEach(chk=>{chk.checked=checked;const c=decodeURIComponent(chk.dataset.canonical);if(checked)combSelected.add(c);else combSelected.delete(c);});combUpdateDetailBtn();});

// FTE settings
document.getElementById('combFteBtn').addEventListener('click',()=>document.getElementById('fteSettings').classList.toggle('open'));
document.getElementById('fteTarget').addEventListener('change',e=>{fteTarget=+e.target.value||1600;saveTagState();combRender();});

// ── Detailed Report generation ────────────────────────────────────────────────
function generateDetailedReport(canonicals){
  const dateStr=new Date().toLocaleDateString('en-GB',{day:'numeric',month:'long',year:'numeric'});
  const safeWindowName=(name)=>name.replace(/[^a-zA-Z0-9_-]/g,'_').replace(/_+/g,'_').replace(/^_|_$/g,'').slice(0,50)||'report';

  for(const canonical of canonicals){
    const d=combData.find(x=>x.canonical===canonical);
    if(!d)continue;
    const tutor=d.tutName?tutAllTutors.find(t=>t.name===d.tutName):null;
    const proj=d.projName?projAllResults.find(r=>r.name===d.projName):null;
    const mmiR=d.mmiName?mmiResults.find(r=>r.name===d.mmiName):null;
    const citRows=d.citName?citAllData.filter(r=>r.holder===d.citName):[];

    // Teaching detail
    let teachingHtml='';
    if(d.tlName&&tlStaffData[d.tlName]){
      const staffWeekMap=tlStaffData[d.tlName];
      const allSess=[];
      for(const w of tlAllWeeks)(staffWeekMap[w]||[]).forEach(s=>allSess.push({...s,week:w}));
      const dayOrder={monday:1,tuesday:2,wednesday:3,thursday:4,friday:5,saturday:6,sunday:7};
      const sortedSess = allSess.slice().sort((a,b)=>{
        if(a.week!==b.week)return a.week-b.week;
        const aD=dayOrder[(a.day||'').toLowerCase()]||99,bD=dayOrder[(b.day||'').toLowerCase()]||99;
        if(aD!==bD)return aD-bD;
        return(timeToHours(a.start)||0)-(timeToHours(b.start)||0);
      });
      const contactTotal=tlAllWeeks.reduce((sum,w)=>sum+calcHours(staffWeekMap[w],tlRealisticMode),0);
      const prepTotal=contactTotal*tlPrepRatio;
      teachingHtml=`<div class="rpt-section">
        <h3>Teaching Load</h3>
        <div class="rpt-summary-row">
          <span>Contact hours: <strong>${contactTotal.toFixed(1)}h</strong></span>
          ${tlPrepRatio>0?`<span>Preparation (${tlPrepRatio}× ratio): <strong>${prepTotal.toFixed(1)}h</strong></span><span>Total incl. prep: <strong>${d.tlHours.toFixed(1)}h</strong></span>`:''}
        </div>
        <table class="rpt-table">
          <thead><tr><th>Module</th><th>Session / Activity</th><th>Type</th><th>Day</th><th>Time</th><th style="text-align:center">Week</th><th style="text-align:right">hrs/session</th><th style="text-align:right">Total hrs</th></tr></thead>
          <tbody>
          ${sortedSess.map(s=>{
            const dur=sessionDuration(s);
            return `<tr>
              <td>${s.moduleCode?`${s.moduleCode}${s.moduleTitle?' – '+s.moduleTitle:''}`:s.moduleTitle||'Unknown Module'}</td>
              <td>${s.sessionTitle||s.activity||'—'}</td>
              <td>${s.type||'—'}</td>
              <td>${s.day||'—'}</td>
              <td>${s.start&&s.end?s.start+'–'+s.end:'—'}</td>
              <td style="text-align:center">Wk ${s.week}</td>
              <td style="text-align:right">${dur.toFixed(1)}</td>
              <td style="text-align:right">${dur.toFixed(1)}</td>
            </tr>`;
          }).join('')}
          </tbody>
          <tfoot><tr><td colspan="7"><strong>Total contact hours</strong></td><td style="text-align:right"><strong>${contactTotal.toFixed(1)}h</strong></td></tr>
          ${tlPrepRatio>0?`<tr><td colspan="7">Preparation hours (${tlPrepRatio}× contact)</td><td style="text-align:right">${prepTotal.toFixed(1)}h</td></tr>
          <tr class="rpt-total-row"><td colspan="7"><strong>Total teaching load (incl. prep)</strong></td><td style="text-align:right"><strong>${d.tlHours.toFixed(1)}h</strong></td></tr>`:''}
          </tfoot>
        </table>
      </div>`;
    }

    // Tutorial detail
    let tutorialHtml='';
    if(tutor){
      tutorialHtml=`<div class="rpt-section">
        <h3>Personal Tutoring</h3>
        <div class="rpt-summary-row">
          <span>Year 1 tutees: <strong>${tutor.year1.length}</strong></span>
          <span>Other year tutees: <strong>${tutor.other.length}</strong></span>
          <span>Total: <strong>${tutor.totalTutees}</strong></span>
          <span>Hours: <strong>${d.tutHours.toFixed(1)}h</strong></span>
        </div>
        ${tutor.year1.length>0?`<h4 style="margin:0.8rem 0 0.4rem;font-size:0.82rem;color:#666;font-weight:600;text-transform:uppercase;letter-spacing:0.04em">Year 1 Tutees</h4>
        <table class="rpt-table"><thead><tr><th>Name</th><th>Course</th><th>Email</th></tr></thead><tbody>
        ${tutor.year1.map(s=>`<tr><td>${s.name||'—'}</td><td>${s.course||'—'}</td><td>${s.email||'—'}</td></tr>`).join('')}
        </tbody></table>`:''}
        ${tutor.other.length>0?`<h4 style="margin:0.8rem 0 0.4rem;font-size:0.82rem;color:#666;font-weight:600;text-transform:uppercase;letter-spacing:0.04em">Other Year Tutees</h4>
        <table class="rpt-table"><thead><tr><th>Name</th><th>Year</th><th>Course</th></tr></thead><tbody>
        ${tutor.other.map(s=>`<tr><td>${s.name||'—'}</td><td>${s.year||'—'}</td><td>${s.course||'—'}</td></tr>`).join('')}
        </tbody></table>`:''}
      </div>`;
    }

    // Project detail
    let projectHtml='';
    if(proj){
      const allProj=[...new Set([...proj.supervised,...proj.cosupervised,...proj.diss_assessed,...proj.poster_assessed])];
      const rolePills=p=>[
        p.supervisors.includes(proj.name)?'Supervisor':'',
        p.cosupervisors.includes(proj.name)?'Co-supervisor':'',
        p.diss1===proj.name||p.diss2===proj.name?'Diss. Assessor':'',
        p.poster1===proj.name||p.poster2===proj.name?'Poster Assessor':'',
      ].filter(Boolean).join(', ');
      projectHtml=`<div class="rpt-section">
        <h3>Project Supervision &amp; Assessment</h3>
        <div class="rpt-summary-row">
          <span>Supervised: <strong>${proj.supervised.length}</strong> (${proj.nSup.toFixed(2)} share)</span>
          <span>Co-supervised: <strong>${proj.cosupervised.length}</strong> (${proj.nCoSup.toFixed(2)} share)</span>
          <span>Diss. assessed: <strong>${proj.nDissAss}</strong></span>
          <span>Poster assessed: <strong>${proj.nPostAss}</strong></span>
        </div>
        <table class="rpt-table">
          <thead><tr><th>Project Title</th><th>Role(s)</th><th style="text-align:right">Hrs</th></tr></thead>
          <tbody>
          ${allProj.map(p=>{
            const supShare=p.supervisors.includes(proj.name)?1/(p.supervisors.length||1):0;
            const coSupShare=p.cosupervisors.includes(proj.name)?1/(p.cosupervisors.length||1):0;
            const hrs=(supShare*(projSettings.supervision+projSettings.diss_feedback+projSettings.poster_feedback))
                     +(coSupShare*projSettings.cosupervision)
                     +((p.diss1===proj.name||p.diss2===proj.name)?projSettings.diss_marking:0)
                     +((p.poster1===proj.name||p.poster2===proj.name)?projSettings.poster_marking:0);
            return`<tr><td>${p.theme||'(No title)'}</td><td>${rolePills(p)}</td><td style="text-align:right">${hrs.toFixed(1)}</td></tr>`;
          }).join('')}
          </tbody>
          <tfoot><tr><td colspan="2"><strong>Total</strong></td><td style="text-align:right"><strong>${d.projHours.toFixed(1)}h</strong></td></tr></tfoot>
        </table>
      </div>`;
    }

    // MMI detail
    let mmiHtml='';
    if(mmiR){
      const activeSess=mmiR.sessions.filter(s=>!s.isReserve);
      mmiHtml=`<div class="rpt-section">
        <h3>MMI Interviewing</h3>
        <div class="rpt-summary-row">
          <span>Sessions: <strong>${activeSess.length}</strong></span>
          <span>Total hours: <strong>${d.mmiHours.toFixed(1)}h</strong></span>
        </div>
        <table class="rpt-table">
          <thead><tr><th>Date</th><th>Label</th><th>Time</th><th style="text-align:right">Duration (h)</th></tr></thead>
          <tbody>
          ${activeSess.map(s=>`<tr><td>${s.dateStr||'—'}</td><td>${s.label||'—'}</td><td>${formatHour(s.startH)}–${formatHour(s.endH)}</td><td style="text-align:right">${s.durationH.toFixed(2)}</td></tr>`).join('')}
          </tbody>
          <tfoot><tr><td colspan="3"><strong>Total</strong></td><td style="text-align:right"><strong>${d.mmiHours.toFixed(1)}h</strong></td></tr></tfoot>
        </table>
      </div>`;
    }

    // Citizenship detail
    let citHtml='';
    if(citRows.length>0){
      citHtml=`<div class="rpt-section">
        <h3>Citizenship &amp; Service Roles</h3>
        <table class="rpt-table">
          <thead><tr><th>Role</th><th>Category</th><th style="text-align:right">hrs/yr</th><th>Term</th><th>End Date</th></tr></thead>
          <tbody>
          ${citRows.map(r=>`<tr><td>${r.role}</td><td>${r.category}</td><td style="text-align:right">${r.hours%1===0?r.hours.toFixed(0):r.hours.toFixed(2)}</td><td>${r.term||'—'}</td><td>${r.end||'—'}</td></tr>`).join('')}
          </tbody>
          <tfoot><tr><td colspan="2"><strong>Total</strong></td><td style="text-align:right"><strong>${d.citHours.toFixed(1)}h</strong></td><td colspan="2"></td></tr></tfoot>
        </table>
      </div>`;
    }

    // Assessment detail
    let assessmentHtml='';
    if(d.assessmentName && assessmentAllData.length>0){
      const rows=assessmentAllData.filter(r=>r.supervisor===d.assessmentName);
      if(rows.length>0){
        assessmentHtml=`<div class="rpt-section">
        <h3>Non-timetabled Assessment Workload</h3>
        <div class="rpt-summary-row">
          <span>Non-timetabled assessments: <strong>${rows.length}</strong></span>
          <span>Total hours: <strong>${d.assessmentHours.toFixed(1)}h</strong></span>
        </div>
        <table class="rpt-table">
          <thead><tr><th>Non-timetabled assess.</th><th>Year</th><th>Course</th><th style="text-align:right">Students</th><th style="text-align:right">Total Load (h)</th><th style="text-align:right">Hours</th></tr></thead>
          <tbody>
          ${rows.map(r=>`<tr><td>${r.assessmentDesc||'—'}</td><td>${r.year||'—'}</td><td>${r.course||'—'}</td><td style="text-align:right">${r.totalStudents||'—'}</td><td style="text-align:right">${r.totalLoad.toFixed(1)}</td><td style="text-align:right">${r.hours.toFixed(1)}</td></tr>`).join('')}
          </tbody>
          <tfoot><tr><td colspan="5"><strong>Total non-timetabled assessment hours</strong></td><td style="text-align:right"><strong>${d.assessmentHours.toFixed(1)}h</strong></td></tr></tfoot>
        </table>
      </div>`;
      }
    }

    // Research detail
    let researchHtml='';
    if(d.resName && resAllData.length>0){
      const row=resAllData.find(r=>r.name===d.resName);
      if(row){
        const resHrs=d.resHours||0;
        researchHtml=`<div class="rpt-section">
        <h3>Staff Research Hours</h3>
        <div class="rpt-summary-row">
          <span>Department: <strong>${row.dept||'—'}</strong></span>
          <span>FTE: <strong>${row.fte>0?row.fte.toFixed(2):'—'}</strong></span>
          <span>Projects: <strong>${row.projects||0}</strong></span>
          <span>Hours/week: <strong>${row.hrsWeek>0?row.hrsWeek.toFixed(1):'—'}</strong></span>
          <span>Total hours: <strong>${resHrs.toFixed(1)}h</strong></span>
        </div>
        <table class="rpt-table">
          <thead><tr><th>Staff ID</th><th>Name</th><th>Department</th><th style="text-align:right">FTE</th><th style="text-align:right">Projects</th><th style="text-align:right">Hrs/Week</th><th style="text-align:right">Curr. Year (h)</th>${resYearPref==='next'?'<th style="text-align:right">Next Year (h)</th>':''}</tr></thead>
          <tbody>
            <tr><td>${row.identifier||'—'}</td><td>${row.name}</td><td>${row.dept||'—'}</td><td style="text-align:right">${row.fte>0?row.fte.toFixed(2):'—'}</td><td style="text-align:right">${row.projects||0}</td><td style="text-align:right">${row.hrsWeek>0?row.hrsWeek.toFixed(1):'—'}</td><td style="text-align:right">${row.currHours.toFixed(1)}</td>${resYearPref==='next'?`<td style="text-align:right">${row.nextHours.toFixed(1)}</td>`:''}</tr>
          </tbody>
          <tfoot><tr><td colspan="${resYearPref==='next'?'7':'6'}"><strong>Total research hours</strong></td><td style="text-align:right"><strong>${resHrs.toFixed(1)}h</strong></td></tr></tfoot>
        </table>
      </div>`;
      }
    }

    // PGR Supervision detail
    let pgrHtml='';
    if(d.pgrName && pgrAllData.length>0){
      const rows=pgrAllData.filter(r=>r.supervisor===d.pgrName);
      if(rows.length>0){
        pgrHtml=`<div class="rpt-section">
        <h3>PGR Supervision</h3>
        <div class="rpt-summary-row">
          <span>Students: <strong>${rows.length}</strong></span>
          <span>Total hours: <strong>${d.pgrHours.toFixed(1)}h</strong></span>
        </div>
        <table class="rpt-table">
          <thead><tr><th>Student</th><th>Plan</th><th>Mode</th><th style="text-align:right">%</th><th>Start</th><th>End</th><th style="text-align:right">Hours</th></tr></thead>
          <tbody>
          ${rows.map(r=>`<tr><td>${r.studentName||'—'}</td><td>${r.plan||'—'}</td><td>${r.mode||'—'}</td><td style="text-align:right">${r.percent.toFixed(0)}%</td><td>${r.startDate||'—'}</td><td>${r.endDate||'—'}</td><td style="text-align:right">${r.hours.toFixed(1)}</td></tr>`).join('')}
          </tbody>
          <tfoot><tr><td colspan="6"><strong>Total PGR supervision hours</strong></td><td style="text-align:right"><strong>${d.pgrHours.toFixed(1)}h</strong></td></tr></tfoot>
        </table>
      </div>`;
      }
    }

    // Summary donut-style bar
    const cats=[['Teaching',d.tlHours,'#0066cc'],['Non-timetabled assess.',d.assessmentHours,'#8a2be2'],['Projects',d.projHours,'#b84c2a'],['Tutorial',d.tutHours,'#1a7a4a'],['MMI',d.mmiHours,'#6b21a8'],['Citizenship',d.citHours,'#c89b2a'],['Research',(d.resHours||0),'#0a7a9a'],['PGR',d.pgrHours,'#d2691e']].filter(([,h])=>h>0);
    const summaryBars=cats.map(([label,h,col])=>`
      <div style="display:flex;align-items:center;gap:10px;margin-bottom:5px">
        <div style="width:110px;font-size:0.82rem;color:#444">${label}</div>
        <div style="flex:1;height:12px;background:#eee;border-radius:6px;overflow:hidden"><div style="height:100%;width:${(h/d.total*100).toFixed(1)}%;background:${col};border-radius:6px"></div></div>
        <div style="width:52px;text-align:right;font-family:monospace;font-size:0.82rem;font-weight:600">${h.toFixed(1)}h</div>
        <div style="width:36px;text-align:right;font-size:0.75rem;color:#888">${(h/d.total*100).toFixed(0)}%</div>
      </div>`).join('');

    const html=`<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">
<title>Academic Workload Report – ${canonical} – ${dateStr}</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Source+Serif+4:ital,wght@0,300;0,400;0,600;0,700;1,400&family=IBM+Plex+Sans:wght@300;400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap');
  *,*::before,*::after{box-sizing:border-box;margin:0;padding:0;}
  body{font-family:'IBM Plex Sans',sans-serif;font-size:13px;color:#1a1f2e;background:#f5f6f8;line-height:1.5;}
  .rpt-toolbar{background:#041e42;color:white;padding:12px 2rem;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:100;gap:1rem;flex-wrap:wrap;}
  .rpt-toolbar h2{font-family:'Source Serif 4',serif;font-size:1rem;font-weight:600;opacity:0.9;}
  .rpt-toolbar-btns{display:flex;gap:8px;}
  .rpt-btn{background:rgba(255,255,255,0.15);border:1px solid rgba(255,255,255,0.3);color:white;padding:6px 14px;border-radius:6px;cursor:pointer;font-size:0.82rem;font-family:'IBM Plex Sans',sans-serif;transition:background 0.2s;}
  .rpt-btn:hover{background:rgba(255,255,255,0.25);}
  .rpt-btn.primary{background:#0066cc;border-color:#004fa3;}
  .rpt-wrap{max-width:900px;margin:2rem auto;padding:0 1.5rem 4rem;}
  .rpt-person{background:white;border-radius:12px;border:1px solid #d0d7e3;overflow:hidden;margin-bottom:2.5rem;}
  .rpt-person-header{background:#041e42;color:white;padding:2rem 2rem 1.5rem;display:flex;align-items:flex-start;justify-content:space-between;gap:1rem;}
  .rpt-person-name{font-family:'Source Serif 4',serif;font-size:1.8rem;font-weight:600;margin-bottom:4px;}
  .rpt-person-sub{font-size:0.78rem;opacity:0.6;}
  .rpt-total-badge{text-align:center;background:rgba(255,255,255,0.1);border:1px solid rgba(255,255,255,0.2);border-radius:12px;padding:1rem 1.5rem;font-family:'IBM Plex Mono',monospace;font-size:2rem;font-weight:500;line-height:1.1;flex-shrink:0;}
  .rpt-total-badge span{display:block;font-family:'IBM Plex Sans',sans-serif;font-size:0.68rem;font-weight:400;opacity:0.7;margin-top:2px;}
  .rpt-overview{display:block;padding:1.5rem 2rem;background:#f8f9fb;border-bottom:1px solid #d0d7e3;}
  .rpt-overview-left{background:white;border-radius:8px;border:1px solid #d0d7e3;padding:1rem 1.2rem;}
  .rpt-section{padding:1.5rem 2rem;border-bottom:1px solid #eef0f5;}
  .rpt-section:last-child{border-bottom:none;}
  .rpt-section h3{font-family:'Source Serif 4',serif;font-size:1.05rem;color:#041e42;margin-bottom:0.8rem;padding-left:12px;border-left:3px solid #0066cc;line-height:1.3;}
  .rpt-summary-row{display:flex;gap:1.5rem;flex-wrap:wrap;margin-bottom:1rem;background:#f0f4ff;border-radius:6px;padding:0.6rem 0.8rem;font-size:0.82rem;}
  .rpt-table{width:100%;border-collapse:collapse;font-size:0.8rem;margin-bottom:1rem;}
  .rpt-table th{background:#041e42;color:white;padding:6px 10px;font-weight:500;text-align:left;font-size:0.75rem;}
  .rpt-table td{padding:6px 10px;border-bottom:1px solid #eef0f5;}
  .rpt-table tfoot td{font-weight:600;background:#f0f4ff;border-top:2px solid #d0d7e3;}
  .rpt-total-row td{background:#e8f0fb;}
  @media print{
    body{background:white;font-size:11px;color:#000;-webkit-print-color-adjust:exact;print-color-adjust:exact;}
    .rpt-toolbar{display:none!important;}
    .rpt-wrap{max-width:100%;margin:0;padding:0 1.5cm;}
    .rpt-person{border:1px solid #ccc;border-radius:4px;box-shadow:none;margin-bottom:0;page-break-inside:avoid;}
    .rpt-person-header{background:#041e42!important;-webkit-print-color-adjust:exact;print-color-adjust:exact;padding:1.5rem 2rem 1rem;}
    .rpt-table th{background:#041e42!important;-webkit-print-color-adjust:exact;print-color-adjust:exact;}
    .rpt-summary-row{background:#f0f4ff!important;-webkit-print-color-adjust:exact;print-color-adjust:exact;}
    .rpt-table tfoot td{background:#f0f4ff!important;-webkit-print-color-adjust:exact;print-color-adjust:exact;}
    .rpt-overview{background:#f8f9fb!important;-webkit-print-color-adjust:exact;print-color-adjust:exact;}
    .rpt-total-badge{border:1px solid rgba(255,255,255,0.3);}
    .rpt-total-badge span{opacity:0.7;}
    .rpt-section h3{color:#041e42;}
    @page{margin:2cm 0;}
  }
</style>
</head><body>
<div class="rpt-toolbar">
  <div>
    <h2>Academic Workload Report — ${canonical}</h2>
  </div>
  <div class="rpt-toolbar-btns">
    <button class="rpt-btn" onclick="window.print()">Print / Save PDF</button>
    <button class="rpt-btn" onclick="window.close()">Close</button>
  </div>
</div>
<div class="rpt-wrap"><div class="rpt-person">
  <div class="rpt-person-header">
    <div>
      <div class="rpt-person-name">${canonical}</div>
      <div class="rpt-person-sub">Academic Workload Report · ${dateStr}</div>
    </div>
    <div style="display:flex;gap:1rem;flex-shrink:0">
      <div class="rpt-total-badge">${d.total.toFixed(1)}<span>hrs total</span></div>
      <div class="rpt-total-badge" style="font-size:1.4rem;background:rgba(255,255,255,0.08)">${ftePct(d.canonical,d.total)}%<span>of ${personalTarget(d.canonical).toFixed(0)}h target</span></div>
    </div>
  </div>
  <div class="rpt-overview">
    <div class="rpt-overview-left">
      <h3 style="margin:0 0 0.8rem;font-size:0.8rem;text-transform:uppercase;letter-spacing:0.06em;color:#666;padding-left:10px;border-left:2px solid #0066cc;line-height:1.2;">Load Summary</h3>
      ${summaryBars}
      <div style="border-top:2px solid #041e42;margin-top:8px;padding:10px 0 0;display:flex;justify-content:space-between;font-size:0.9rem;">
        <span style="font-weight:700;color:#041e42;">Total</span><span style="font-family:'IBM Plex Mono',monospace;font-weight:700;color:#0066cc;font-size:1rem;">${d.total.toFixed(1)}h</span>
      </div>
    </div>
  </div>
  ${teachingHtml}${assessmentHtml}${projectHtml}${tutorialHtml}${mmiHtml}${citHtml}${researchHtml}${pgrHtml}
</div></div>
</body></html>`;

    const w=window.open('', safeWindowName(canonical));
    w.document.write(html);
    w.document.close();
  }
}

function generateCombinedReport(canonicals){
  const dateStr=new Date().toLocaleDateString('en-GB',{day:'numeric',month:'long',year:'numeric'});
  let sections='';

  for(let i=0;i<canonicals.length;i++){
    const canonical=canonicals[i];
    const d=combData.find(x=>x.canonical===canonical);
    if(!d)continue;
    const tutor=d.tutName?tutAllTutors.find(t=>t.name===d.tutName):null;
    const proj=d.projName?projAllResults.find(r=>r.name===d.projName):null;
    const mmiR=d.mmiName?mmiResults.find(r=>r.name===d.mmiName):null;
    const citRows=d.citName?citAllData.filter(r=>r.holder===d.citName):[];

    // Teaching detail
    let teachingHtml='';
    if(d.tlName&&tlStaffData[d.tlName]){
      const staffWeekMap=tlStaffData[d.tlName];
      const allSess=[];
      for(const w of tlAllWeeks)(staffWeekMap[w]||[]).forEach(s=>allSess.push({...s,week:w}));
      const dayOrder={monday:1,tuesday:2,wednesday:3,thursday:4,friday:5,saturday:6,sunday:7};
      const sortedSess = allSess.slice().sort((a,b)=>{
        if(a.week!==b.week)return a.week-b.week;
        const aD=dayOrder[(a.day||'').toLowerCase()]||99,bD=dayOrder[(b.day||'').toLowerCase()]||99;
        if(aD!==bD)return aD-bD;
        return(timeToHours(a.start)||0)-(timeToHours(b.start)||0);
      });
      const contactTotal=tlAllWeeks.reduce((sum,w)=>sum+calcHours(staffWeekMap[w],tlRealisticMode),0);
      const prepTotal=contactTotal*tlPrepRatio;
      teachingHtml=`<div class="rpt-section">
        <h3>Teaching Load</h3>
        <div class="rpt-summary-row">
          <span>Contact hours: <strong>${contactTotal.toFixed(1)}h</strong></span>
          ${tlPrepRatio>0?`<span>Preparation (${tlPrepRatio}× ratio): <strong>${prepTotal.toFixed(1)}h</strong></span><span>Total incl. prep: <strong>${d.tlHours.toFixed(1)}h</strong></span>`:''}
        </div>
        <table class="rpt-table">
          <thead><tr><th>Module</th><th>Session / Activity</th><th>Type</th><th>Day</th><th>Time</th><th style="text-align:center">Week</th><th style="text-align:right">hrs/session</th><th style="text-align:right">Total hrs</th></tr></thead>
          <tbody>
          ${sortedSess.map(s=>{
            const dur=sessionDuration(s);
            return `<tr>
              <td>${s.moduleCode?`${s.moduleCode}${s.moduleTitle?' – '+s.moduleTitle:''}`:s.moduleTitle||'Unknown Module'}</td>
              <td>${s.sessionTitle||s.activity||'—'}</td>
              <td>${s.type||'—'}</td>
              <td>${s.day||'—'}</td>
              <td>${s.start&&s.end?s.start+'–'+s.end:'—'}</td>
              <td style="text-align:center">Wk ${s.week}</td>
              <td style="text-align:right">${dur.toFixed(1)}</td>
              <td style="text-align:right">${dur.toFixed(1)}</td>
            </tr>`;
          }).join('')}
          </tbody>
          <tfoot><tr><td colspan="7"><strong>Total contact hours</strong></td><td style="text-align:right"><strong>${contactTotal.toFixed(1)}h</strong></td></tr>
          ${tlPrepRatio>0?`<tr><td colspan="7">Preparation hours (${tlPrepRatio}× contact)</td><td style="text-align:right">${prepTotal.toFixed(1)}h</td></tr>
          <tr class="rpt-total-row"><td colspan="7"><strong>Total teaching load (incl. prep)</strong></td><td style="text-align:right"><strong>${d.tlHours.toFixed(1)}h</strong></td></tr>`:''}
          </tfoot>
        </table>
      </div>`;
    }

    // Tutorial detail
    let tutorialHtml='';
    if(tutor){
      tutorialHtml=`<div class="rpt-section">
        <h3>Personal Tutoring</h3>
        <div class="rpt-summary-row">
          <span>Year 1 tutees: <strong>${tutor.year1.length}</strong></span>
          <span>Other year tutees: <strong>${tutor.other.length}</strong></span>
          <span>Total: <strong>${tutor.totalTutees}</strong></span>
          <span>Hours: <strong>${d.tutHours.toFixed(1)}h</strong></span>
        </div>
        ${tutor.year1.length>0?`<h4 style="margin:0.8rem 0 0.4rem;font-size:0.82rem;color:#666;font-weight:600;text-transform:uppercase;letter-spacing:0.04em">Year 1 Tutees</h4>
        <table class="rpt-table"><thead><tr><th>Name</th><th>Course</th><th>Email</th></tr></thead><tbody>
        ${tutor.year1.map(s=>`<tr><td>${s.name||'—'}</td><td>${s.course||'—'}</td><td>${s.email||'—'}</td></tr>`).join('')}
        </tbody></table>`:''}
        ${tutor.other.length>0?`<h4 style="margin:0.8rem 0 0.4rem;font-size:0.82rem;color:#666;font-weight:600;text-transform:uppercase;letter-spacing:0.04em">Other Year Tutees</h4>
        <table class="rpt-table"><thead><tr><th>Name</th><th>Year</th><th>Course</th></tr></thead><tbody>
        ${tutor.other.map(s=>`<tr><td>${s.name||'—'}</td><td>${s.year||'—'}</td><td>${s.course||'—'}</td></tr>`).join('')}
        </tbody></table>`:''}
      </div>`;
    }

    // Project detail
    let projectHtml='';
    if(proj){
      const allProj=[...new Set([...proj.supervised,...proj.cosupervised,...proj.diss_assessed,...proj.poster_assessed])];
      const rolePills=p=>[
        p.supervisors.includes(proj.name)?'Supervisor':'',
        p.cosupervisors.includes(proj.name)?'Co-supervisor':'',
        p.diss1===proj.name||p.diss2===proj.name?'Diss. Assessor':'',
        p.poster1===proj.name||p.poster2===proj.name?'Poster Assessor':'',
      ].filter(Boolean).join(', ');
      projectHtml=`<div class="rpt-section">
        <h3>Project Supervision &amp; Assessment</h3>
        <div class="rpt-summary-row">
          <span>Supervised: <strong>${proj.supervised.length}</strong> (${proj.nSup.toFixed(2)} share)</span>
          <span>Co-supervised: <strong>${proj.cosupervised.length}</strong> (${proj.nCoSup.toFixed(2)} share)</span>
          <span>Diss. assessed: <strong>${proj.nDissAss}</strong></span>
          <span>Poster assessed: <strong>${proj.nPostAss}</strong></span>
        </div>
        <table class="rpt-table">
          <thead><tr><th>Project Title</th><th>Role(s)</th><th style="text-align:right">Hrs</th></tr></thead>
          <tbody>
          ${allProj.map(p=>{
            const supShare=p.supervisors.includes(proj.name)?1/(p.supervisors.length||1):0;
            const coSupShare=p.cosupervisors.includes(proj.name)?1/(p.cosupervisors.length||1):0;
            const hrs=(supShare*(projSettings.supervision+projSettings.diss_feedback+projSettings.poster_feedback))
                     +(coSupShare*projSettings.cosupervision)
                     +((p.diss1===proj.name||p.diss2===proj.name)?projSettings.diss_marking:0)
                     +((p.poster1===proj.name||p.poster2===proj.name)?projSettings.poster_marking:0);
            return`<tr><td>${p.theme||'(No title)'}</td><td>${rolePills(p)}</td><td style="text-align:right">${hrs.toFixed(1)}</td></tr>`;
          }).join('')}
          </tbody>
          <tfoot><tr><td colspan="2"><strong>Total</strong></td><td style="text-align:right"><strong>${d.projHours.toFixed(1)}h</strong></td></tr></tfoot>
        </table>
      </div>`;
    }

    // MMI detail
    let mmiHtml='';
    if(mmiR){
      const activeSess=mmiR.sessions.filter(s=>!s.isReserve);
      mmiHtml=`<div class="rpt-section">
        <h3>MMI Interviewing</h3>
        <div class="rpt-summary-row">
          <span>Sessions: <strong>${activeSess.length}</strong></span>
          <span>Total hours: <strong>${d.mmiHours.toFixed(1)}h</strong></span>
        </div>
        <table class="rpt-table">
          <thead><tr><th>Date</th><th>Label</th><th>Time</th><th style="text-align:right">Duration (h)</th></tr></thead>
          <tbody>
          ${activeSess.map(s=>`<tr><td>${s.dateStr||'—'}</td><td>${s.label||'—'}</td><td>${formatHour(s.startH)}–${formatHour(s.endH)}</td><td style="text-align:right">${s.durationH.toFixed(2)}</td></tr>`).join('')}
          </tbody>
          <tfoot><tr><td colspan="3"><strong>Total</strong></td><td style="text-align:right"><strong>${d.mmiHours.toFixed(1)}h</strong></td></tr></tfoot>
        </table>
      </div>`;
    }

    // Citizenship detail
    let citHtml='';
    if(citRows.length>0){
      citHtml=`<div class="rpt-section">
        <h3>Citizenship &amp; Service Roles</h3>
        <table class="rpt-table">
          <thead><tr><th>Role</th><th>Category</th><th style="text-align:right">hrs/yr</th><th>Term</th><th>End Date</th></tr></thead>
          <tbody>
          ${citRows.map(r=>`<tr><td>${r.role}</td><td>${r.category}</td><td style="text-align:right">${r.hours%1===0?r.hours.toFixed(0):r.hours.toFixed(2)}</td><td>${r.term||'—'}</td><td>${r.end||'—'}</td></tr>`).join('')}
          </tbody>
          <tfoot><tr><td colspan="2"><strong>Total</strong></td><td style="text-align:right"><strong>${d.citHours.toFixed(1)}h</strong></td><td colspan="2"></td></tr></tfoot>
        </table>
      </div>`;
    }

    // Assessment detail
    let assessmentHtml='';
    if(d.assessmentName && assessmentAllData.length>0){
      const rows=assessmentAllData.filter(r=>r.supervisor===d.assessmentName);
      if(rows.length>0){
        assessmentHtml=`<div class="rpt-section">
        <h3>Non-timetabled Assessment Workload</h3>
        <div class="rpt-summary-row">
          <span>Non-timetabled assessments: <strong>${rows.length}</strong></span>
          <span>Total hours: <strong>${d.assessmentHours.toFixed(1)}h</strong></span>
        </div>
        <table class="rpt-table">
          <thead><tr><th>Non-timetabled assess.</th><th>Year</th><th>Course</th><th style="text-align:right">Students</th><th style="text-align:right">Total Load (h)</th><th style="text-align:right">Hours</th></tr></thead>
          <tbody>
          ${rows.map(r=>`<tr><td>${r.assessmentDesc||'—'}</td><td>${r.year||'—'}</td><td>${r.course||'—'}</td><td style="text-align:right">${r.totalStudents||'—'}</td><td style="text-align:right">${r.totalLoad.toFixed(1)}</td><td style="text-align:right">${r.hours.toFixed(1)}</td></tr>`).join('')}
          </tbody>
          <tfoot><tr><td colspan="5"><strong>Total non-timetabled assessment hours</strong></td><td style="text-align:right"><strong>${d.assessmentHours.toFixed(1)}h</strong></td></tr></tfoot>
        </table>
      </div>`;
      }
    }

    // Research detail
    let researchHtml='';
    if(d.resName && resAllData.length>0){
      const row=resAllData.find(r=>r.name===d.resName);
      if(row){
        const resHrs=d.resHours||0;
        researchHtml=`<div class="rpt-section">
        <h3>Staff Research Hours</h3>
        <div class="rpt-summary-row">
          <span>Department: <strong>${row.dept||'—'}</strong></span>
          <span>FTE: <strong>${row.fte>0?row.fte.toFixed(2):'—'}</strong></span>
          <span>Projects: <strong>${row.projects||0}</strong></span>
          <span>Hours/week: <strong>${row.hrsWeek>0?row.hrsWeek.toFixed(1):'—'}</strong></span>
          <span>Total hours: <strong>${resHrs.toFixed(1)}h</strong></span>
        </div>
        <table class="rpt-table">
          <thead><tr><th>Staff ID</th><th>Name</th><th>Department</th><th style="text-align:right">FTE</th><th style="text-align:right">Projects</th><th style="text-align:right">Hrs/Week</th><th style="text-align:right">Curr. Year (h)</th>${resYearPref==='next'?'<th style="text-align:right">Next Year (h)</th>':''}</tr></thead>
          <tbody>
            <tr><td>${row.identifier||'—'}</td><td>${row.name}</td><td>${row.dept||'—'}</td><td style="text-align:right">${row.fte>0?row.fte.toFixed(2):'—'}</td><td style="text-align:right">${row.projects||0}</td><td style="text-align:right">${row.hrsWeek>0?row.hrsWeek.toFixed(1):'—'}</td><td style="text-align:right">${row.currHours.toFixed(1)}</td>${resYearPref==='next'?`<td style="text-align:right">${row.nextHours.toFixed(1)}</td>`:''}</tr>
          </tbody>
          <tfoot><tr><td colspan="${resYearPref==='next'?'7':'6'}"><strong>Total research hours</strong></td><td style="text-align:right"><strong>${resHrs.toFixed(1)}h</strong></td></tr></tfoot>
        </table>
      </div>`;
      }
    }

    // PGR Supervision detail
    let pgrHtml='';
    if(d.pgrName && pgrAllData.length>0){
      const rows=pgrAllData.filter(r=>r.supervisor===d.pgrName);
      if(rows.length>0){
        pgrHtml=`<div class="rpt-section">
        <h3>PGR Supervision</h3>
        <div class="rpt-summary-row">
          <span>Students: <strong>${rows.length}</strong></span>
          <span>Total hours: <strong>${d.pgrHours.toFixed(1)}h</strong></span>
        </div>
        <table class="rpt-table">
          <thead><tr><th>Student</th><th>Plan</th><th>Mode</th><th style="text-align:right">%</th><th>Start</th><th>End</th><th style="text-align:right">Hours</th></tr></thead>
          <tbody>
          ${rows.map(r=>`<tr><td>${r.studentName||'—'}</td><td>${r.plan||'—'}</td><td>${r.mode||'—'}</td><td style="text-align:right">${r.percent.toFixed(0)}%</td><td>${r.startDate||'—'}</td><td>${r.endDate||'—'}</td><td style="text-align:right">${r.hours.toFixed(1)}</td></tr>`).join('')}
          </tbody>
          <tfoot><tr><td colspan="6"><strong>Total PGR supervision hours</strong></td><td style="text-align:right"><strong>${d.pgrHours.toFixed(1)}h</strong></td></tr></tfoot>
        </table>
      </div>`;
      }
    }

    // Summary donut-style bar
    const cats=[['Teaching',d.tlHours,'#0066cc'],['Non-timetabled assess.',d.assessmentHours,'#8a2be2'],['Projects',d.projHours,'#b84c2a'],['Tutorial',d.tutHours,'#1a7a4a'],['MMI',d.mmiHours,'#6b21a8'],['Citizenship',d.citHours,'#c89b2a'],['Research',(d.resHours||0),'#0a7a9a'],['PGR',d.pgrHours,'#d2691e']].filter(([,h])=>h>0);
    const summaryBars=cats.map(([label,h,col])=>`
      <div style="display:flex;align-items:center;gap:10px;margin-bottom:5px">
        <div style="width:110px;font-size:0.82rem;color:#444">${label}</div>
        <div style="flex:1;height:12px;background:#eee;border-radius:6px;overflow:hidden"><div style="height:100%;width:${(h/d.total*100).toFixed(1)}%;background:${col};border-radius:6px"></div></div>
        <div style="width:52px;text-align:right;font-family:monospace;font-size:0.82rem;font-weight:600">${h.toFixed(1)}h</div>
        <div style="width:36px;text-align:right;font-size:0.75rem;color:#888">${(h/d.total*100).toFixed(0)}%</div>
      </div>`).join('');

    sections+=`
    <div class="rpt-person${i>0?' rpt-person-break':''}">
      <div class="rpt-person-header">
        <div>
          <div class="rpt-person-name">${canonical}</div>
          <div class="rpt-person-sub">Academic Workload Report · ${dateStr}</div>
        </div>
        <div style="display:flex;gap:1rem;flex-shrink:0">
          <div class="rpt-total-badge">${d.total.toFixed(1)}<span>hrs total</span></div>
          <div class="rpt-total-badge" style="font-size:1.4rem;background:rgba(255,255,255,0.08)">${ftePct(d.canonical,d.total)}%<span>of ${personalTarget(d.canonical).toFixed(0)}h target</span></div>
        </div>
      </div>
      <div class="rpt-overview">
        <div class="rpt-overview-left">
          <h3 style="margin:0 0 0.8rem;font-size:0.8rem;text-transform:uppercase;letter-spacing:0.06em;color:#666;padding-left:10px;border-left:2px solid #0066cc;line-height:1.2;">Load Summary</h3>
          ${summaryBars}
          <div style="border-top:2px solid #041e42;margin-top:8px;padding:10px 0 0;display:flex;justify-content:space-between;font-size:0.9rem;">
            <span style="font-weight:700;color:#041e42;">Total</span><span style="font-family:'IBM Plex Mono',monospace;font-weight:700;color:#0066cc;font-size:1rem;">${d.total.toFixed(1)}h</span>
          </div>
        </div>
      </div>
      ${teachingHtml}${assessmentHtml}${projectHtml}${tutorialHtml}${mmiHtml}${citHtml}${researchHtml}${pgrHtml}
    </div>`;
  }

  if(!sections)return;

  const html=`<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">
<title>Academic Workload Report – Combined – ${dateStr}</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Source+Serif+4:ital,wght@0,300;0,400;0,600;0,700;1,400&family=IBM+Plex+Sans:wght@300;400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap');
  *,*::before,*::after{box-sizing:border-box;margin:0;padding:0;}
  body{font-family:'IBM Plex Sans',sans-serif;font-size:13px;color:#1a1f2e;background:#f5f6f8;line-height:1.5;}
  .rpt-toolbar{background:#041e42;color:white;padding:12px 2rem;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:100;gap:1rem;flex-wrap:wrap;}
  .rpt-toolbar h2{font-family:'Source Serif 4',serif;font-size:1rem;font-weight:600;opacity:0.9;}
  .rpt-toolbar-btns{display:flex;gap:8px;}
  .rpt-btn{background:rgba(255,255,255,0.15);border:1px solid rgba(255,255,255,0.3);color:white;padding:6px 14px;border-radius:6px;cursor:pointer;font-size:0.82rem;font-family:'IBM Plex Sans',sans-serif;transition:background 0.2s;}
  .rpt-btn:hover{background:rgba(255,255,255,0.25);}
  .rpt-btn.primary{background:#0066cc;border-color:#004fa3;}
  .rpt-wrap{max-width:900px;margin:2rem auto;padding:0 1.5rem 4rem;}
  .rpt-person{background:white;border-radius:12px;border:1px solid #d0d7e3;overflow:hidden;margin-bottom:2.5rem;}
  .rpt-person-break{page-break-before:always;}
  .rpt-person-header{background:#041e42;color:white;padding:2rem 2rem 1.5rem;display:flex;align-items:flex-start;justify-content:space-between;gap:1rem;}
  .rpt-person-name{font-family:'Source Serif 4',serif;font-size:1.8rem;font-weight:600;margin-bottom:4px;}
  .rpt-person-sub{font-size:0.78rem;opacity:0.6;}
  .rpt-total-badge{text-align:center;background:rgba(255,255,255,0.1);border:1px solid rgba(255,255,255,0.2);border-radius:12px;padding:1rem 1.5rem;font-family:'IBM Plex Mono',monospace;font-size:2rem;font-weight:500;line-height:1.1;flex-shrink:0;}
  .rpt-total-badge span{display:block;font-family:'IBM Plex Sans',sans-serif;font-size:0.68rem;font-weight:400;opacity:0.7;margin-top:2px;}
  .rpt-overview{display:block;padding:1.5rem 2rem;background:#f8f9fb;border-bottom:1px solid #d0d7e3;}
  .rpt-overview-left{background:white;border-radius:8px;border:1px solid #d0d7e3;padding:1rem 1.2rem;}
  .rpt-section{padding:1.5rem 2rem;border-bottom:1px solid #eef0f5;}
  .rpt-section:last-child{border-bottom:none;}
  .rpt-section h3{font-family:'Source Serif 4',serif;font-size:1.05rem;color:#041e42;margin-bottom:0.8rem;padding-left:12px;border-left:3px solid #0066cc;line-height:1.3;}
  .rpt-summary-row{display:flex;gap:1.5rem;flex-wrap:wrap;margin-bottom:1rem;background:#f0f4ff;border-radius:6px;padding:0.6rem 0.8rem;font-size:0.82rem;}
  .rpt-table{width:100%;border-collapse:collapse;font-size:0.8rem;margin-bottom:1rem;}
  .rpt-table th{background:#041e42;color:white;padding:6px 10px;font-weight:500;text-align:left;font-size:0.75rem;}
  .rpt-table td{padding:6px 10px;border-bottom:1px solid #eef0f5;}
  .rpt-table tfoot td{font-weight:600;background:#f0f4ff;border-top:2px solid #d0d7e3;}
  .rpt-total-row td{background:#e8f0fb;}
  @media print{
    body{background:white;font-size:11px;color:#000;-webkit-print-color-adjust:exact;print-color-adjust:exact;}
    .rpt-toolbar{display:none!important;}
    .rpt-wrap{max-width:100%;margin:0;padding:0 1.5cm;}
    .rpt-person{border:1px solid #ccc;border-radius:4px;box-shadow:none;margin-bottom:0;page-break-inside:avoid;}
    .rpt-person-break{page-break-before:always!important;}
    .rpt-person-header{background:#041e42!important;-webkit-print-color-adjust:exact;print-color-adjust:exact;padding:1.5rem 2rem 1rem;}
    .rpt-table th{background:#041e42!important;-webkit-print-color-adjust:exact;print-color-adjust:exact;}
    .rpt-summary-row{background:#f0f4ff!important;-webkit-print-color-adjust:exact;print-color-adjust:exact;}
    .rpt-table tfoot td{background:#f0f4ff!important;-webkit-print-color-adjust:exact;print-color-adjust:exact;}
    .rpt-overview{background:#f8f9fb!important;-webkit-print-color-adjust:exact;print-color-adjust:exact;}
    .rpt-total-badge{border:1px solid rgba(255,255,255,0.3);}
    .rpt-total-badge span{opacity:0.7;}
    .rpt-section h3{color:#041e42;}
    @page{margin:2cm 0;}
  }
</style>
</head><body>
<div class="rpt-toolbar">
  <div>
    <h2>Academic Workload Report — ${canonicals.length} staff members</h2>
  </div>
  <div class="rpt-toolbar-btns">
    <button class="rpt-btn" onclick="window.print()">Print / Save PDF</button>
    <button class="rpt-btn" onclick="window.close()">Close</button>
  </div>
</div>
<div class="rpt-wrap">${sections}</div>
</body></html>`;

  const w=window.open('','_blank');
  w.document.write(html);
  w.document.close();
}

document.getElementById('combDetailBtn').addEventListener('click',()=>{
  if(combSelected.size===0)return;
  generateDetailedReport([...combSelected]);
});

document.getElementById('combCombinedBtn').addEventListener('click',()=>{
  if(combSelected.size===0)return;
  generateCombinedReport([...combSelected]);
});

document.getElementById('combExportBtn').addEventListener('click',()=>{
  const wb2=XLSX.utils.book_new();
  const rows=[['Academic','Teaching Name','Non-timetabled assess. Name','Project Name','Tutorial Name','MMI Name','Citizenship Name','Research Name','PGR Name','Teaching Hrs','Non-timetabled assess. Hrs','Project Hrs','Tutorial Hrs','MMI Hrs','Citizenship Hrs','Research Hrs','PGR Hrs','Total Hrs','Match Type']];
  for(const d of combData)rows.push([d.canonical,d.tlName||'',d.assessmentName||'',d.projName||'',d.tutName||'',d.mmiName||'',d.citName||'',d.resName||'',d.pgrName||'',+d.tlHours.toFixed(2),+d.assessmentHours.toFixed(2),+d.projHours.toFixed(2),+d.tutHours.toFixed(2),+d.mmiHours.toFixed(2),+d.citHours.toFixed(2),+(d.resHours||0).toFixed(2),+d.pgrHours.toFixed(2),+d.total.toFixed(2),d.matchType]);
  rows.push(['Grand Total','','','','','','','','',+combData.reduce((s,d)=>s+d.tlHours,0).toFixed(2),+combData.reduce((s,d)=>s+d.assessmentHours,0).toFixed(2),+combData.reduce((s,d)=>s+d.projHours,0).toFixed(2),+combData.reduce((s,d)=>s+d.tutHours,0).toFixed(2),+combData.reduce((s,d)=>s+d.mmiHours,0).toFixed(2),+combData.reduce((s,d)=>s+d.citHours,0).toFixed(2),+combData.reduce((s,d)=>s+(d.resHours||0),0).toFixed(2),+combData.reduce((s,d)=>s+d.pgrHours,0).toFixed(2),+combData.reduce((s,d)=>s+d.total,0).toFixed(2),'']);
  XLSX.utils.book_append_sheet(wb2,XLSX.utils.aoa_to_sheet(rows),'Combined Load');
  XLSX.writeFile(wb2,'academic_load_combined.xlsx');
});

updateCombStatus();

// ── Model export / import button handlers ────────────────────────────────────
document.getElementById('combModelExportBtn').addEventListener('click',exportModel);
document.getElementById('combModelImportInput').addEventListener('change',function(){
  if(this.files[0])importModel(this.files[0]);
  this.value='';
});

loadTagState();
