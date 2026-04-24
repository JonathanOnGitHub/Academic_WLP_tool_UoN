// ═══════════════════════════════════════════════════════
// SHARED UTILITIES
// ═══════════════════════════════════════════════════════
let citizenshipTotals = {};
// ── Name normalisation & matching utilities ──────────────────────────────────

// Common nickname → canonical first-name expansions (bidirectional lookup built below)
const NICKNAMES = {
  charlie:['charles'],charles:['charlie','chuck','chas'],
  bill:['william'],will:['william'],willie:['william'],william:['bill','will','willie','wills'],
  bob:['robert'],rob:['robert'],bobby:['robert'],robert:['bob','rob','robbie','bobby'],
  jim:['james'],jimmy:['james'],jamie:['james'],james:['jim','jimmy','jamie'],
  dave:['david'],davy:['david'],david:['dave','davy'],
  mike:['michael'],mick:['michael'],mickey:['michael'],michael:['mike','mick','mickey'],
  nick:['nicholas'],nicky:['nicholas'],nicholas:['nick','nicky'],
  tony:['anthony'],ant:['anthony'],anthony:['tony','ant'],
  sue:['susan'],susie:['susan'],susan:['sue','susie'],
  liz:['elizabeth'],beth:['elizabeth'],betty:['elizabeth'],elizabeth:['liz','beth','betty','lisa'],
  kate:['katherine','kathryn','catherine'],kath:['katherine','kathryn','catherine'],
  katherine:['kate','kath','kathy'],kathryn:['kate','kath','kathy'],catherine:['kate','kath','cathy'],
  alex:['alexander','alexandra'],alexander:['alex','al'],alexandra:['alex'],
  andy:['andrew'],drew:['andrew'],andrew:['andy','drew'],
  chris:['christopher','christian'],christopher:['chris'],
  dan:['daniel'],danny:['daniel'],daniel:['dan','danny'],
  jon:['jonathan'],jonathan:['jon'],
  matt:['matthew'],matthew:['matt'],
  sam:['samuel','samantha'],samuel:['sam'],samantha:['sam'],
  steve:['stephen','steven'],stephen:['steve'],steven:['steve'],
  tom:['thomas'],tommy:['thomas'],thomas:['tom','tommy'],
  tim:['timothy'],timothy:['tim'],
  ben:['benjamin'],benjamin:['ben'],
  joe:['joseph'],joseph:['joe'],
  phil:['philip','phillip'],philip:['phil'],phillip:['phil'],
  pete:['peter'],peter:['pete'],
  pat:['patrick','patricia'],patrick:['pat'],patricia:['pat'],
  rick:['richard'],rich:['richard'],dick:['richard'],richard:['rick','rich','dick'],
  fred:['frederick'],frederick:['fred'],
  ed:['edward'],edward:['ed','eddie','ted'],
  jenny:['jennifer'],jen:['jennifer'],jennifer:['jenny','jen'],
  maggie:['margaret'],meg:['margaret'],margaret:['maggie','meg','peggy'],
  wendy:['gwendolyn'],gwendolyn:['wendy'],
  vicky:['victoria'],victoria:['vicky','vic'],
  nicky:['nicola'],nicola:['nicky'],
  rosie:['rosalind','rose','rosa'],rosalind:['rosie'],rose:['rosie'],rosa:['rosie'],
  kati:['katalin'],katalin:['kati'],
  beppe:['giuseppe'],giuseppe:['beppe'],
  sj:['steve','stephen'],steve:['sj'],stephen:['sj'],
};

// Build reverse: for any name, return all equivalents (including itself)
function nicknameVariants(name){
  const n=name.toLowerCase();
  const vars=new Set([n]);
  if(NICKNAMES[n])NICKNAMES[n].forEach(v=>vars.add(v));
  // also check if n appears as a value in any entry
  for(const[k,vs]of Object.entries(NICKNAMES)){if(vs.includes(n)){vars.add(k);vs.forEach(v=>vars.add(v));}}
  return vars;
}

// Surname prefixes that should be kept with the surname
const SURNAME_PREFIXES = ['de','van','von','der','den','di','da','del','dos','du','la','le','mac','mc','o','st','saint','ter','ten','van de','van den','van der','von dem','von der'];

function normaliseName(raw){
  // Strip titles wherever they appear: leading, trailing, or after comma-flip
  // CRITICAL FIX: Removed 'g' flag to prevent lastIndex state issues!
  const TITLE_RE=/\b(prof\.?|professor|dr\.?|mr\.?|mrs\.?|ms\.?|mx\.?|rev\.?|sir)\b\s*/i;
  let n=String(raw||'').trim();

  // Handle "Surname, First" format by flipping
  const cm=n.match(/^([^,]+),\s*(.+)$/);if(cm)n=cm[2]+' '+cm[1];

  // Strip titles (using replace without 'g' flag - we handle all occurrences by re-evaluating)
  // We need to loop because after removing one title, another might be revealed
  let prev;
  do{
    prev=n;
    n=n.replace(TITLE_RE,'');
  }while(n!==prev);

  // Normalize whitespace, hyphens, apostrophes to spaces
  n=n.toLowerCase().replace(/[-.']/g,' ').replace(/\s+/g,' ').trim();

  return n;
}

// Extract tokens, handling surname prefixes specially
// CRITICAL FIX: Now correctly combines prefixes with what FOLLOWS them, not what precedes
function nameTokens(raw){
  const norm=normaliseName(raw);
  const tokens=norm.split(' ').filter(Boolean);

  // Handle surname prefixes: combine "de" + "moor" into "de moor" as single token
  // But do NOT combine "cornelia" + "de" - prefixes only combine with what follows
  const result=[];
  let i=0;
  while(i<tokens.length){
    // Check if CURRENT token is a prefix AND there's a next token to combine with
    if(i<tokens.length-1 && SURNAME_PREFIXES.includes(tokens[i])){
      // Combine prefix with the next token (the actual surname part)
      result.push(tokens[i]+' '+tokens[i+1]);
      i+=2; // skip both the prefix and the next token
    }else{
      result.push(tokens[i]);
      i++;
    }
  }
  return result;
}

// Surname = last token after normalisation (handles prefixes like "de moor")
function extractSurname(raw){const t=nameTokens(raw);return t.length>0?t[t.length-1]:'';}
// First token (given name or initial)
function extractFirst(raw){const t=nameTokens(raw);return t.length>0?t[0]:'';}
// All given names (everything except last token)
function extractGivenNames(raw){const t=nameTokens(raw);return t.slice(0,-1);}

// Does 'initial' (single char) match the start of 'fullToken'?
function initialMatches(initial,fullToken){return initial.length===1&&fullToken.startsWith(initial);}

// Does a run of initials like "gc" match a sequence of given-name tokens like ["gautam","c..."]?
// Each character in the initials string must match the start of a corresponding token (in order).
function multiInitialMatches(initials, tokens){
  if(!/^[a-z]+$/.test(initials))return false;
  if(initials.length<2)return false;
  if(initials.length>tokens.length)return false;
  for(let i=0;i<initials.length;i++){
    if(!tokens[i].startsWith(initials[i]))return false;
  }
  return true;
}

// Check if single token appears anywhere in the other name (including as part of multi-word token)
function tokenAppearsIn(singleToken, allTokens){
  for(const t of allTokens){
    if(t===singleToken)return true;
    if(t.includes(' ') && t.split(' ').includes(singleToken))return true;
  }
  return false;
}

// Core similarity: returns score 0–1 between two raw name strings
// Handles: token overlap, nickname expansion, initial matching, surname-only, single-token matching
function nameSimilarity(a,b){
  if(!a||!b)return 0;
  const na=normaliseName(a),nb=normaliseName(b);
  if(na===nb)return 1.0;

  const ta=nameTokens(a);
  const tb=nameTokens(b);

  // ── 0. Handle single-token vs multi-token matching ────────────────────────
  // Case: "Cornelia" vs "de Moor, Cornelia Dr" or "Cornelia de Moor"
  if(ta.length===1 || tb.length===1){
    const singleTokens=ta.length===1?ta:tb;
    const multiTokens=ta.length===1?tb:ta;
    const single=singleTokens[0];

    // Direct token match
    if(tokenAppearsIn(single, multiTokens))return 0.95;

    // Initial match: "C" matches "cornelia"
    if(multiTokens.some(t=>initialMatches(single,t)))return 0.90;

    // Nickname variant match
    const vars=nicknameVariants(single);
    for(const v of vars){
      if(tokenAppearsIn(v, multiTokens))return 0.88;
    }

    // Partial substring match (for compound names) — require both sides ≥3 chars to avoid
    // single-char tokens like "a" matching inside any first name (e.g. "a" ⊂ "frankie")
    if(single.length>=3){
      for(const mt of multiTokens){
        if(mt.length>=3 && (mt.includes(single) || single.includes(mt)))return 0.75;
      }
    }
  }

  // ── 1. Standard token sort ratio ─────────────────────────────────────────
  const sA=new Set(ta),sB=new Set(tb);
  let inter=0;for(const w of sA)if(sB.has(w))inter++;
  const union=sA.size+sB.size-inter;
  const tsr=union===0?0:inter/union;
  if(tsr>=1.0)return 1.0;

  // ── 2. Nickname expansion: replace first token with all variants, recheck ─
  function expandFirst(tokens){
    if(tokens.length===0)return[tokens];
    const variants=nicknameVariants(tokens[0]);
    return[...variants].map(v=>[v,...tokens.slice(1)]);
  }
  const aExpanded=expandFirst(ta);
  const bExpanded=expandFirst(tb);
  let bestNick=tsr;
  for(const ae of aExpanded)for(const be of bExpanded){
    const sAe=new Set(ae),sBe=new Set(be);
    let i2=0;for(const w of sAe)if(sBe.has(w))i2++;
    const u2=sAe.size+sBe.size-i2;
    const sc=u2===0?0:i2/u2;
    if(sc>bestNick)bestNick=sc;
  }
  if(bestNick>=0.75)return bestNick;

  // ── 3. Initial matching ──────────────────────────────────────────────────
  const surA=ta[ta.length-1],surB=tb[tb.length-1];
  const firstA=ta[0],firstB=tb[0];
  if(surA&&surB&&surA===surB){
    // surnames match — now check first names / initials
    if(firstA===firstB)return 0.9;
    if(initialMatches(firstA,firstB)||initialMatches(firstB,firstA))return 0.82;
    // Multi-char initials: "gc" vs ["gautam","c..."] — check each char against given-name token sequence
    const givenA=ta.slice(0,-1), givenB=tb.slice(0,-1);
    if(multiInitialMatches(firstA,givenB)||multiInitialMatches(firstB,givenA))return 0.82;
    // Also: first char of a multi-initial token matches the full first name ("gc" → "g" matches "gautam")
    if(firstA.length>1&&firstA[0]===firstB[0])return 0.78;
    if(firstB.length>1&&firstB[0]===firstA[0])return 0.78;
    const varA=nicknameVariants(firstA),varB=nicknameVariants(firstB);
    if([...varA].some(v=>v===firstB||v.startsWith(firstB))||[...varB].some(v=>v===firstA||v.startsWith(firstA)))return 0.80;
    return 0.55;
  }

  // ── 4. Partial / substring matching for surname prefixes ─────────────────
  // Handle "de Moor" vs "Moor" - check if one surname contains the other
  if(surA && surB){
    const surALast=surA.split(' ').pop(); // Get last part after prefixes
    const surBLast=surB.split(' ').pop();
    if(surALast===surBLast){
      // Core surnames match, check given names
      const givenA=ta.slice(0,-1),givenB=tb.slice(0,-1);
      if(givenA.length===0 || givenB.length===0){
        // One is surname-only
        return 0.85;
      }
      // Both have given names - check overlap
      const givenSetA=new Set(givenA),givenSetB=new Set(givenB);
      let gInter=0;
      for(const g of givenSetA)if(givenSetB.has(g))gInter++;
      if(gInter>0)return 0.80+gInter*0.05;
    }
  }

  // ── 5. Check if one name is a substring of the other ────────────────────
  if(na.length>=3 && nb.length>=3 && (na.includes(nb)||nb.includes(na)))return 0.70;

  return bestNick;
}

function tokenSortRatio(a,b){return nameSimilarity(a,b);}

// NEW: Pre-merge names within each source before cross-source merge
// This handles cases like MMI having both "Snow" and "Snow Stolnik"
function intraSourceMerge(names, source){
  if(names.length<=1)return names.map(n=>({raw:n,canonical:n}));

  const groups=[];
  const used=new Set();

  for(let i=0;i<names.length;i++){
    if(used.has(i))continue;
    const group={raws:[names[i]],canonical:names[i]};
    used.add(i);

    for(let j=i+1;j<names.length;j++){
      if(used.has(j))continue;
      const sc=nameSimilarity(names[i],names[j]);
      if(sc>=0.85){ // High threshold for intra-source matching
        group.raws.push(names[j]);
        used.add(j);
        // Update canonical to the most complete name
        if(nameTokens(names[j]).length>nameTokens(group.canonical).length){
          group.canonical=names[j];
        }
      }
    }
    groups.push(group);
  }

  return groups.map(g=>({raw:g.canonical,canonical:g.canonical,raws:g.raws}));
}

// THE DEFINITIVE mergeNameLists - only one definition, comprehensive logic
function mergeNameLists(lists){
  // NEW: First, pre-merge within each source
  const preMergedLists=lists.map(({source,names})=>{
    const merged=intraSourceMerge(names,source);
    return{source,names:merged.map(m=>m.raw),canonicalMap:merged.reduce((acc,m)=>{acc[m.raw]=m.canonical;return acc;},{})};
  });

  const THRESH=0.65,all=[];
  for(const{source,names}of preMergedLists)for(const n of names)all.push({norm:normaliseName(n),raw:n,source});
  const groups=[],used=new Set(),key=(r,s)=>`${s}::${r}`;

  // Pass 1: rich similarity merge with all matching logic
  for(const e of all){
    if(used.has(key(e.raw,e.source)))continue;
    const g={canonical:e.raw,sources:{[e.source]:e.raw}};
    used.add(key(e.raw,e.source));

    for(const o of all){
      if(used.has(key(o.raw,o.source)))continue;
      if(o.source===e.source)continue;
      if(g.sources[o.source])continue;
      const sc=nameSimilarity(e.raw,o.raw);
      if(sc>=THRESH){
        g.sources[o.source]=o.raw;
        used.add(key(o.raw,o.source));
        const mtype=sc>=0.99?'exact':sc>=0.80?'fuzzy':'fuzzy';
        g.matchType=g.matchType?g.matchType:mtype;
        if(!g.score||sc<g.score)g.score=sc;
      }
    }
    if(!g.matchType)g.matchType=Object.keys(g.sources).length>1?'exact':'only';
    groups.push(g);
  }

  // Pass 2: remaining single-token entries — match to any compatible group
  // Use nameSimilarity for proper matching (handles "Snow" matching "Snow Stolnik")
  for(const g of groups){
    if(Object.keys(g.sources).length>1||g._merged)continue;
    const src=Object.keys(g.sources)[0];
    const tokens=normaliseName(g.canonical).split(' ').filter(Boolean);
    if(tokens.length!==1)continue;

    const candidates=groups.filter(og=>{
      if(og===g||og._merged)return false;
      if(og.sources[src])return false;
      // Use full nameSimilarity check instead of simple token matching
      // This allows "Snow" to match "Snow Stolnik" via single-token logic
      const sc=nameSimilarity(g.canonical,og.canonical);
      return sc>=0.70; // Lower threshold for Pass 2 to catch more matches
    });

    if(candidates.length===1){
      candidates[0].sources[src]=g.canonical;
      candidates[0].matchType='firstname';
      candidates[0].firstnameMatch=true;
      g._merged=true;
    }
  }

  // Optimize canonical names: choose the name with most tokens (most complete)
  for(const g of groups){
    if(g._merged)continue;
    const allNames=Object.values(g.sources);
    if(allNames.length>1){
      // Find the name with the most tokens (prefer full names over single names)
      const bestName=allNames.reduce((best,name)=>{
        const score=nameTokens(name).length;
        return score>best.score?{name,score}:best;
      },{name:allNames[0],score:0});
      g.canonical=bestName.name;
    }
  }

  return groups.filter(g=>!g._merged);
}
function closePanel(){document.getElementById('detailPanel').classList.remove('open');document.getElementById('overlay').classList.remove('open');}
function openPanel(name,sub,bodyHtml){document.getElementById('panelName').textContent=name;document.getElementById('panelSub').textContent=sub;document.getElementById('panelBody').innerHTML=bodyHtml;document.getElementById('detailPanel').classList.add('open');document.getElementById('overlay').classList.add('open');}
function fh(h){return h>0?h.toFixed(1):'—';}
function fmt(h){return h.toFixed(1);}

document.querySelectorAll('.main-tab-btn').forEach(btn=>{btn.addEventListener('click',()=>{document.querySelectorAll('.main-tab-btn').forEach(b=>b.classList.remove('active'));document.querySelectorAll('.main-tab-panel').forEach(p=>p.classList.remove('active'));btn.classList.add('active');document.getElementById(btn.dataset.panel).classList.add('active');});});
// ═══════════════════════════════════════════════════════
// TAB 1 — TEACHING LOAD (unchanged from v2)
// ═══════════════════════════════════════════════════════
let tlUploadedFiles=[],tlParsedSessions=[],tlWeekRange=[1,52],tlRealisticMode=false;
let tlStaffData={},tlModuleData={},tlTypeData={};
let tlAllWeeks=[],tlAllStaff=[],tlAllModules=[],tlAllTypes=[];
let tlStaffSort={col:'total',dir:-1},tlModSort={col:'total',dir:-1},tlTypeSort={col:'total',dir:-1};
let tlPrepRatio=0; // 0 = no prep hours; >0 = prep hrs added per contact hr
const tlFileInput=document.getElementById('tlFileInput'),tlDropZone=document.getElementById('tlDropZone'),tlFileList=document.getElementById('tlFileList'),tlAnalyseBtn=document.getElementById('tlAnalyseBtn');
function tlUpdateFileList(){tlFileList.innerHTML='';tlUploadedFiles.forEach((f,i)=>{const d=document.createElement('div');d.className='file-item';d.innerHTML=`<span>📄</span><span class="fi-name">${f.name}</span><button class="fi-remove" data-i="${i}">✕</button>`;tlFileList.appendChild(d);});tlAnalyseBtn.disabled=tlUploadedFiles.length===0;}
tlFileInput.addEventListener('change',e=>{tlUploadedFiles.push(...Array.from(e.target.files));tlUpdateFileList();tlFileInput.value='';});
tlDropZone.addEventListener('dragover',e=>{e.preventDefault();tlDropZone.classList.add('drag-over');});tlDropZone.addEventListener('dragleave',()=>tlDropZone.classList.remove('drag-over'));
tlDropZone.addEventListener('drop',e=>{e.preventDefault();tlDropZone.classList.remove('drag-over');tlUploadedFiles.push(...Array.from(e.dataTransfer.files).filter(f=>f.name.match(/\.html?$/i)));tlUpdateFileList();});
tlFileList.addEventListener('click',e=>{if(e.target.classList.contains('fi-remove')){tlUploadedFiles.splice(+e.target.dataset.i,1);tlUpdateFileList();}});
function parseWeeks(str){if(!str)return[];const ws=new Set();for(const p of String(str).trim().split(',')){const t=p.trim(),r=t.match(/^(\d+)\s*[-–—]\s*(\d+)$/);if(r){for(let w=+r[1];w<=+r[2];w++)ws.add(w);}else{const n=+t;if(!isNaN(n)&&n>0)ws.add(n);}}return[...ws].sort((a,b)=>a-b);}
function timeToHours(t){if(!t)return 0;const m=String(t).match(/(\d+):(\d+)/);if(m)return+m[1]+ +m[2]/60;const h=parseFloat(t);return isNaN(h)?0:h;}
function sessionDuration(s){const d=timeToHours(s.end)-timeToHours(s.start);return d>0?Math.round(d*100)/100:0;}
function parseHTMLFile(html){const doc=new DOMParser().parseFromString(html,'text/html'),sessions=[];for(const table of doc.querySelectorAll('table')){const rows=table.querySelectorAll('tr');if(rows.length<2)continue;let headerRow=null,headerIdx=0;for(let i=0;i<Math.min(rows.length,5);i++){const cells=rows[i].querySelectorAll('th,td'),texts=[...cells].map(c=>c.textContent.trim().toLowerCase());if(texts.some(t=>t.includes('module')||t.includes('activity')||t.includes('staff'))){headerRow=cells;headerIdx=i;break;}}if(!headerRow)continue;const cols={},headerMap={activity:['activity'],moduleCode:['module code','module_code','modulecode'],moduleTitle:['module title','module_title','moduletitle'],sessionTitle:['session title','session_title','sessiontitle'],type:['type'],weeks:['weeks','week'],day:['day'],start:['start'],end:['end'],staff:['staff'],location:['location','room'],groupInfo:['group information','group info','group'],notes:['notes']};[...headerRow].forEach((cell,i)=>{const t=cell.textContent.trim().toLowerCase();for(const[key,variants]of Object.entries(headerMap)){if(variants.some(v=>t.includes(v))&&cols[key]===undefined)cols[key]=i;}});if(cols.weeks===undefined&&cols.staff===undefined)continue;for(let i=headerIdx+1;i<rows.length;i++){const cells=rows[i].querySelectorAll('td,th');if(cells.length<2)continue;const get=k=>(cols[k]!==undefined&&cells[cols[k]])?cells[cols[k]].innerHTML:'';const getText=k=>(cols[k]!==undefined&&cells[cols[k]])?cells[cols[k]].textContent.trim():'';const weeks=parseWeeks(getText('weeks'));if(weeks.length===0)continue;const staffNames=get('staff').split(/<br\s*\/?>/gi).map(s=>s.replace(/<[^>]+>/g,'').trim()).filter(Boolean);if(staffNames.length===0)continue;sessions.push({activity:getText('activity'),moduleCode:getText('moduleCode'),moduleTitle:getText('moduleTitle'),sessionTitle:getText('sessionTitle'),type:getText('type'),weeks,weeksRaw:getText('weeks'),day:getText('day'),start:getText('start'),end:getText('end'),staff:staffNames,location:getText('location'),groupInfo:getText('groupInfo'),notes:getText('notes')});}}return sessions;}
function calcHours(sessions,realistic){if(!sessions||sessions.length===0)return 0;if(realistic&&sessions.staffMap){let total=0;for(const ss of sessions.staffMap.values())total+=calcHours(ss,true);return Math.round(total*100)/100;}if(!realistic)return sessions.reduce((s,x)=>s+sessionDuration(x),0);const byDay={};for(const s of sessions){const d=s.day||'unknown';if(!byDay[d])byDay[d]=[];byDay[d].push({start:timeToHours(s.start),end:timeToHours(s.end)});}let total=0;for(const slots of Object.values(byDay)){const sorted=slots.filter(s=>s.end>s.start).sort((a,b)=>a.start-b.start);let merged=[];for(const s of sorted){if(merged.length&&s.start<merged[merged.length-1].end)merged[merged.length-1].end=Math.max(merged[merged.length-1].end,s.end);else merged.push({...s});}total+=merged.reduce((s,x)=>s+(x.end-x.start),0);}return Math.round(total*100)/100;}
function sessionKey(s){return[s.moduleCode,s.moduleTitle,s.sessionTitle,s.day,s.start,s.end,s.type,s.location].join('|');}
function deduplicateSessions(sessions){const seen=new Map();for(const s of sessions){const k=sessionKey(s);if(!seen.has(k))seen.set(k,s);}return[...seen.values()];}
function formatWeekRange(weeks){if(!weeks||weeks.length===0)return'—';const sorted=[...weeks].sort((a,b)=>a-b);if(sorted.length===1)return'Wk '+sorted[0];let contiguous=true;for(let i=1;i<sorted.length;i++){if(sorted[i]!==sorted[i-1]+1){contiguous=false;break;}}if(contiguous)return'Wk '+sorted[0]+'–'+sorted[sorted.length-1];return'Wk '+sorted.join(', ');}
function tlAggregate(sessions,wFrom,wTo){tlStaffData={};tlModuleData={};tlTypeData={};for(const sess of sessions){const fw=sess.weeks.filter(w=>w>=wFrom&&w<=wTo);if(fw.length===0)continue;for(const name of sess.staff){if(!tlStaffData[name])tlStaffData[name]={};for(const w of fw){if(!tlStaffData[name][w])tlStaffData[name][w]=[];tlStaffData[name][w].push(sess);}const mk=sess.moduleCode||sess.moduleTitle||'Unknown';if(!tlModuleData[mk])tlModuleData[mk]={};for(const w of fw){if(!tlModuleData[mk][w]){const _arr=[];Object.defineProperty(_arr,'staffMap',{value:new Map(),enumerable:false});tlModuleData[mk][w]=_arr;}tlModuleData[mk][w].push(sess);if(!tlModuleData[mk][w].staffMap.has(name))tlModuleData[mk][w].staffMap.set(name,[]);tlModuleData[mk][w].staffMap.get(name).push(sess);}const tk=(sess.type||'').trim()||'Undefined';if(!tlTypeData[tk])tlTypeData[tk]={};for(const w of fw){if(!tlTypeData[tk][w]){const _arr=[];Object.defineProperty(_arr,'staffMap',{value:new Map(),enumerable:false});tlTypeData[tk][w]=_arr;}tlTypeData[tk][w].push(sess);if(!tlTypeData[tk][w].staffMap.has(name))tlTypeData[tk][w].staffMap.set(name,[]);tlTypeData[tk][w].staffMap.get(name).push(sess);}}}tlAllStaff=Object.keys(tlStaffData).sort();tlAllModules=Object.keys(tlModuleData).sort();tlAllTypes=Object.keys(tlTypeData).sort((a,b)=>a==='Undefined'?1:b==='Undefined'?-1:a.localeCompare(b));const ws=new Set();for(const s of sessions)for(const w of s.weeks)if(w>=wFrom&&w<=wTo)ws.add(w);tlAllWeeks=[...ws].sort((a,b)=>a-b);}
function buildGrid(dataMap,entities,weeks,sortConfig,realistic,prepRatio){const pr=prepRatio||0;if(entities.length===0)return{html:'<div style="padding:2rem;color:var(--muted)">No data found.</div>',legendHtml:''};const sorted=[...entities];if(sortConfig.col==='name')sorted.sort((a,b)=>sortConfig.dir*a.localeCompare(b));else if(sortConfig.col==='total')sorted.sort((a,b)=>sortConfig.dir*(weeks.reduce((s,w)=>s+calcHours(dataMap[a]?.[w],realistic)*(1+pr),0)-weeks.reduce((s,w)=>s+calcHours(dataMap[b]?.[w],realistic)*(1+pr),0)));else sorted.sort((a,b)=>sortConfig.dir*(calcHours(dataMap[a]?.[sortConfig.col],realistic)*(1+pr)-calcHours(dataMap[b]?.[sortConfig.col],realistic)*(1+pr)));let maxH=0;for(const e of entities)for(const w of weeks){const h=calcHours(dataMap[e]?.[w],realistic)*(1+pr);if(h>maxH)maxH=h;}const breaks=[0,maxH*0.1,maxH*0.25,maxH*0.45,maxH*0.65,maxH*0.85];const heatClass=h=>{if(h<=0)return'';for(let i=breaks.length-1;i>=0;i--)if(h>=breaks[i])return`heat-${i}`;return'heat-0';};const totArr=sortConfig.col==='total'?(sortConfig.dir>0?' ↑':' ↓'):'';const nameArr=sortConfig.col==='name'?(sortConfig.dir>0?' ↑':' ↓'):'';let html=`<table class="grid-table"><thead><tr><th style="text-align:left" data-sort="name">Name${nameArr}</th>`;for(const w of weeks){const arr=sortConfig.col===w?(sortConfig.dir>0?' ↑':' ↓'):'';html+=`<th class="week-header" data-sort-week="${w}">W${w}${arr}</th>`;}html+=`<th class="week-header" data-sort="total" style="background:#0a3060">Total${totArr}</th></tr></thead><tbody>`;for(const entity of sorted){html+=`<tr><td class="name-cell" data-entity="${encodeURIComponent(entity)}" title="${entity}">${entity}</td>`;let rowTotal=0;for(const w of weeks){const contact=calcHours(dataMap[entity]?.[w],realistic);const h=contact*(1+pr);rowTotal+=h;if(h>0){const tip=pr>0?`title="${contact.toFixed(1)}h contact + ${(contact*pr).toFixed(1)}h prep"`:'' ;html+=`<td class="data-cell ${heatClass(h)}" data-entity="${encodeURIComponent(entity)}" data-week="${w}" ${tip}><span>${h.toFixed(1)}</span><small>${pr>0?'incl. prep':'hrs'}</small></td>`;}else html+=`<td class="data-cell empty" data-entity="${encodeURIComponent(entity)}" data-week="${w}">–</td>`;}html+=`<td class="data-cell heat-3" data-entity="${encodeURIComponent(entity)}" data-week="total"><span>${rowTotal.toFixed(1)}</span><small>${pr>0?'incl. prep':'hrs'}</small></td></tr>`;}html+='<tr class="totals-row"><td class="name-cell" data-week="total" data-entity="all">Grand Total</td>';let gt=0;for(const w of weeks){const wt=entities.reduce((s,e)=>s+calcHours(dataMap[e]?.[w],realistic)*(1+pr),0);gt+=wt;html+=`<td class="data-cell" data-week="${w}" data-entity="all"><span>${wt.toFixed(1)}</span><small>hrs</small></td>`;}html+=`<td class="data-cell" data-week="total" data-entity="all"><span>${gt.toFixed(1)}</span><small>hrs</small></td></tr></tbody></table>`;const legendHtml=`<span>Colour scale:</span>${breaks.map((b,i)=>`<span class="legend-item"><span class="legend-swatch heat-${i}"></span>${b.toFixed(0)}${i<breaks.length-1?'–'+breaks[i+1].toFixed(0):'+'}</span>`).join('')}`;return{html,legendHtml};}
function tlRenderGrid(id,dataMap,allEntities,weeks,sortCfg,realistic,legendId,prepRatio){const res=buildGrid(dataMap,allEntities,weeks,sortCfg,realistic,prepRatio||0);const wrap=document.getElementById(id);wrap.innerHTML=res.html;if(legendId)document.getElementById(legendId).innerHTML=res.legendHtml;tlAttachGridEvents(wrap,dataMap,id.includes('Staff')?'staff':id.includes('Mod')?'module':'type');}
function tlRenderStaffGrid(){tlRenderGrid('tlStaffGridWrap',tlStaffData,tlAllStaff,tlAllWeeks,tlStaffSort,tlRealisticMode,'tlStaffLegend',tlPrepRatio);document.getElementById('tlStaffSortInfo').textContent=`Sorted by: ${tlStaffSort.col==='name'?'Name':tlStaffSort.col==='total'?'Total':'Week '+tlStaffSort.col} (${tlStaffSort.dir>0?'asc':'desc'})${tlPrepRatio>0?' · incl. '+tlPrepRatio+'× prep':''}`;}
function tlRenderModGrid(){tlRenderGrid('tlModGridWrap',tlModuleData,modTagFilteredModules(),tlAllWeeks,tlModSort,tlRealisticMode,'tlModLegend',tlPrepRatio);renderModuleTagChips();}
function tlRenderTypeGrid(){tlRenderGrid('tlTypeGridWrap',tlTypeData,tlAllTypes,tlAllWeeks,tlTypeSort,tlRealisticMode,'tlTypeLegend',tlPrepRatio);}
function tlAttachGridEvents(wrap,dataMap,type){wrap.querySelectorAll('.data-cell,.name-cell').forEach(cell=>{cell.addEventListener('click',()=>{const entity=cell.dataset.entity?decodeURIComponent(cell.dataset.entity):null,week=cell.dataset.week;if(!entity)return;const entities=entity==='all'?Object.keys(dataMap):[entity],weeks=(week==='total'||!week)?tlAllWeeks:[+week];const applyR=tlRealisticMode;let totalH=0,rawS=[];for(const e of entities)for(const w of weeks){totalH+=calcHours(dataMap[e]?.[w],applyR);rawS.push(...(dataMap[e]?.[w]||[]));}const unique=deduplicateSessions(rawS),realH=calcHours(unique,true);let html=`<div class="modal-stat-row"><div class="modal-stat"><div class="ms-value">${totalH.toFixed(1)}</div><div class="ms-label">Staff-hours</div></div><div class="modal-stat"><div class="ms-value">${realH.toFixed(1)}</div><div class="ms-label">Deduped hrs</div></div><div class="modal-stat"><div class="ms-value">${unique.length}</div><div class="ms-label">Sessions</div></div></div><div>`;const dayOrder={monday:1,tuesday:2,wednesday:3,thursday:4,friday:5,saturday:6,sunday:7};const sorted=unique.slice().sort((a,b)=>{const aWeek=Math.min(...a.weeks),bWeek=Math.min(...b.weeks);if(aWeek!==bWeek)return aWeek-bWeek;const aDay=dayOrder[(a.day||'').toLowerCase()]||99,bDay=dayOrder[(b.day||'').toLowerCase()]||99;if(aDay!==bDay)return aDay-bDay;const aStart=timeToHours(a.start),bStart=timeToHours(b.start);return aStart-bStart;});for(const s of sorted){const dur=sessionDuration(s);const weekStr=s.weeksRaw||(s.weeks.length?'W'+s.weeks.join(','):'');html+=`<div class="session-card"><div class="sc-title">${s.moduleCode?`<span class="sc-tag">${s.moduleCode}</span>`:''} ${s.moduleTitle||s.sessionTitle||s.activity||'Session'}<span class="sc-hours">${dur.toFixed(1)}h</span></div><div class="sc-meta">${s.day?`<span>📅 ${s.day}</span>`:''} ${weekStr?`<span>📆 ${weekStr}</span>`:''} ${s.start&&s.end?`<span>🕐 ${s.start}–${s.end}</span>`:''} ${s.location?`<span>📍 ${s.location}</span>`:''}</div></div>`;}html+='</div>';openPanel(entity==='all'?'Grand Total':entity,week==='total'?'All weeks':`Week ${week}`,html);});});const thead=wrap.querySelector('thead');if(thead)thead.addEventListener('click',e=>{const th=e.target.closest('th');if(!th)return;const col=th.dataset.sort,w=th.dataset.sortWeek?+th.dataset.sortWeek:null;const s=type==='staff'?tlStaffSort:type==='module'?tlModSort:tlTypeSort;const render=type==='staff'?tlRenderStaffGrid:type==='module'?tlRenderModGrid:tlRenderTypeGrid;if(col){if(s.col===col)s.dir*=-1;else{s.col=col;s.dir=-1;}render();}else if(w!==null){if(s.col===w)s.dir*=-1;else{s.col=w;s.dir=-1;}render();}});}
function tlUpdateStatsBar(){const staffH=tlAllStaff.reduce((sum,staff)=>sum+tlAllWeeks.reduce((wSum,week)=>wSum+calcHours(tlStaffData[staff]?.[week],tlRealisticMode),0),0);const prepH=staffH*tlPrepRatio;const totalH=staffH+prepH;const cards=[['Sessions',tlParsedSessions.length],['Staff',tlAllStaff.length],['Modules',tlAllModules.length],['Weeks',tlAllWeeks.length],['Staff hrs',staffH.toFixed(0)]];if(tlPrepRatio>0){cards.push(['Prep hrs ('+tlPrepRatio+'×)',prepH.toFixed(0)]);cards.push(['Total hrs',totalH.toFixed(0)]);}document.getElementById('tlStatsBar').innerHTML=cards.map(([l,v])=>`<div class="stat-card"><div class="sc-v">${v}</div><div class="sc-l">${l}</div></div>`).join('');}
tlAnalyseBtn.addEventListener('click',async()=>{tlAnalyseBtn.disabled=true;tlAnalyseBtn.textContent='⏳ Processing…';tlParsedSessions=[];for(const file of tlUploadedFiles){const html=await file.text();tlParsedSessions.push(...parseHTMLFile(html));}const wFrom=+document.getElementById('tlWeekFrom').value||1,wTo=+document.getElementById('tlWeekTo').value||52;tlWeekRange=[wFrom,wTo];tlRealisticMode=document.getElementById('tlRealistic').checked;document.getElementById('tlRealistic2').checked=tlRealisticMode;tlAggregate(tlParsedSessions,wFrom,wTo);document.getElementById('tlAnalyserTitle').textContent=document.getElementById('tlTitle').value||'Teaching Load';document.getElementById('tlAnalyserMeta').textContent=`${tlParsedSessions.length} sessions · ${tlAllStaff.length} staff · ${tlAllModules.length} modules · ${tlAllTypes.length} types`;tlUpdateStatsBar();document.getElementById('tl-landing').style.display='none';document.getElementById('tl-analyser').style.display='block';document.getElementById('badge-teaching').textContent=tlAllStaff.length+' staff';tlRenderStaffGrid();tlRenderModGrid();tlRenderTypeGrid();renderModuleTagFilterBar();tlAnalyseBtn.disabled=false;tlAnalyseBtn.textContent='📊 Analyse Timetable';updateCombStatus();});
document.getElementById('tlRealistic2').addEventListener('change',e=>{tlRealisticMode=e.target.checked;tlUpdateStatsBar();tlRenderStaffGrid();tlRenderModGrid();tlRenderTypeGrid();});
document.getElementById('tlBtnSettings').addEventListener('click',()=>{document.getElementById('tlInlineSettings').classList.toggle('open');});
document.getElementById('tlRecalcBtn').addEventListener('click',()=>{const enabled=document.getElementById('tl_prep_enabled').checked;tlPrepRatio=enabled?(+document.getElementById('tl_prep_ratio').value||0):0;tlUpdateStatsBar();tlRenderStaffGrid();tlRenderModGrid();tlRenderTypeGrid();updateCombStatus();});
document.getElementById('tl_prep_enabled').addEventListener('change',()=>{document.getElementById('tl_prep_ratio').disabled=!document.getElementById('tl_prep_enabled').checked;});
document.querySelectorAll('[data-tltab]').forEach(btn=>{btn.addEventListener('click',()=>{document.querySelectorAll('[data-tltab]').forEach(b=>b.classList.remove('active'));btn.classList.add('active');const t=btn.dataset.tltab;['staff','modules','types'].forEach(x=>document.getElementById(`tl-tab-${x}`).style.display=x===t?'':'none');});});
document.getElementById('tlBtnBack').addEventListener('click',()=>{document.getElementById('tl-landing').style.display='';document.getElementById('tl-analyser').style.display='none';});
['tlSortStaffName','tlSortStaffTotal','tlSortModName','tlSortModTotal','tlSortTypeName','tlSortTypeTotal'].forEach(id=>{document.getElementById(id).addEventListener('click',()=>{const[,,type,col]=id.match(/tlSort(Staff|Mod|Type)(Name|Total)/),s=type==='Staff'?tlStaffSort:type==='Mod'?tlModSort:tlTypeSort,c=col==='Name'?'name':'total';if(s.col===c)s.dir*=-1;else{s.col=c;s.dir=col==='Name'?1:-1;}if(type==='Staff')tlRenderStaffGrid();else if(type==='Mod')tlRenderModGrid();else tlRenderTypeGrid();});});
document.getElementById('modTagFilterClear').addEventListener('click',()=>{moduleTagFilter=null;renderModuleTagFilterBar();tlRenderModGrid();});
document.getElementById('tlBtnExport').addEventListener('click',()=>{const wb=XLSX.utils.book_new();const pr=tlPrepRatio;const headers=pr>0?['Staff',...tlAllWeeks.map(w=>`Week ${w} Contact`),...tlAllWeeks.map(w=>`Week ${w} Prep`),'Total Contact','Total Prep','Total (incl. prep)']:['Staff',...tlAllWeeks.map(w=>`Week ${w}`),'Total'];const rows=[headers];for(const name of tlAllStaff){if(pr>0){const contacts=tlAllWeeks.map(w=>calcHours(tlStaffData[name]?.[w],tlRealisticMode));const totC=contacts.reduce((a,b)=>a+b,0);const totP=totC*pr;rows.push([name,...contacts,...contacts.map(c=>+(c*pr).toFixed(2)),+totC.toFixed(2),+totP.toFixed(2),+(totC+totP).toFixed(2)]);}else{const row=[name];let tot=0;for(const w of tlAllWeeks){const h=calcHours(tlStaffData[name]?.[w],tlRealisticMode);row.push(h||'');tot+=h;}row.push(+tot.toFixed(2));rows.push(row);}}if(pr>0){const contacts=tlAllWeeks.map(w=>tlAllStaff.reduce((s,e)=>s+calcHours(tlStaffData[e]?.[w],tlRealisticMode),0));const gtC=contacts.reduce((a,b)=>a+b,0);const gtP=gtC*pr;rows.push(['Grand Total',...contacts,...contacts.map(c=>+(c*pr).toFixed(2)),+gtC.toFixed(2),+gtP.toFixed(2),+(gtC+gtP).toFixed(2)]);}else{const gr=['Grand Total'];let gt=0;for(const w of tlAllWeeks){const wt=tlAllStaff.reduce((s,e)=>s+calcHours(tlStaffData[e]?.[w],tlRealisticMode),0);gr.push(+wt.toFixed(2));gt+=wt;}gr.push(+gt.toFixed(2));rows.push(gr);}XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(rows),'Staff Load');XLSX.writeFile(wb,'teaching_load.xlsx');});

// ═══════════════════════════════════════════════════════
// TAB 2 — TUTORIAL WORKLOAD (unchanged)
// ═══════════════════════════════════════════════════════
let tutAllTutors=[],tutMaxHours=0,tutSortCol='hours',tutSortDir='desc';
const tutDropZone=document.getElementById('tutDropZone'),tutFileInput=document.getElementById('tutFileInput');
tutDropZone.addEventListener('dragover',e=>{e.preventDefault();tutDropZone.classList.add('drag-over');});tutDropZone.addEventListener('dragleave',()=>tutDropZone.classList.remove('drag-over'));tutDropZone.addEventListener('drop',e=>{e.preventDefault();tutDropZone.classList.remove('drag-over');if(e.dataTransfer.files[0])tutProcessFile(e.dataTransfer.files[0]);});tutFileInput.addEventListener('change',e=>{if(e.target.files[0])tutProcessFile(e.target.files[0]);});
function tutShowError(msg){const el=document.getElementById('tutError');el.textContent=msg;el.classList.add('show');}function tutClearError(){document.getElementById('tutError').classList.remove('show');}
function tutProcessFile(file){tutClearError();const y1h=+document.getElementById('tutY1Hours').value||16,oh=+document.getElementById('tutOtherHours').value||8,extra=+document.getElementById('tutExtraAllowance').value||0;const courseCodes=document.getElementById('tutSelectedCourses').value.split(',').map(s=>s.trim().toUpperCase()).filter(Boolean);const reader=new FileReader();reader.onload=e=>{try{const wb=XLSX.read(e.target.result,{type:'array'}),ws=wb.Sheets[wb.SheetNames[0]],raw=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});const headerRow=raw[5]||[];const norm=s=>String(s).toLowerCase().replace(/\s+/g,'').replace(/[^a-z0-9]/g,''),headers=headerRow.map(norm),col=key=>headers.indexOf(norm(key));const iYear=col('Year of Study'),iSurname=col('Surname'),iFirst=col('First Name'),iCourse=col('Course'),iEmail=col('UoN Email'),iTutor=col('Tutor'),iTutorEmail=col('Tutor email'),iStaff=col('Staff Indicator');if(iTutor===-1){tutShowError('Could not find a "Tutor" column. Check headers are on row 6.');return;}const tutorMap={};raw.slice(6).forEach(row=>{if(iStaff!==-1&&String(row[iStaff]).trim().toLowerCase()==='yes')return;const tutorName=String(row[iTutor]||'').trim();if(!tutorName)return;const yearNum=parseInt(String(row[iYear]||'').trim(),10),isY1=yearNum===1;const tutorEmail=iTutorEmail!==-1?String(row[iTutorEmail]||'').trim():'';const studentName=[String(row[iFirst]||'').trim(),String(row[iSurname]||'').trim()].filter(Boolean).join(' ');const course=iCourse!==-1?String(row[iCourse]||'').trim():'',studentEmail=iEmail!==-1?String(row[iEmail]||'').trim():'';if(!tutorMap[tutorName])tutorMap[tutorName]={name:tutorName,email:tutorEmail,year1:[],other:[]};const entry={name:studentName,year:String(row[iYear]||'').trim(),course,email:studentEmail};if(isY1)tutorMap[tutorName].year1.push(entry);else tutorMap[tutorName].other.push(entry);});tutAllTutors=Object.values(tutorMap).map(t=>{const extraHours=t.year1.filter(s=>courseCodes.includes(s.course.toUpperCase())).length*extra + t.other.filter(s=>courseCodes.includes(s.course.toUpperCase())).length*extra; return {...t,totalTutees:t.year1.length+t.other.length,hours:t.year1.length*y1h+t.other.length*oh+extraHours,extraHours};});if(tutAllTutors.length===0){tutShowError('No tutee records found.');return;}tutMaxHours=Math.max(...tutAllTutors.map(t=>t.hours));document.getElementById('tut-landing').style.display='none';document.getElementById('tut-content').style.display='block';document.getElementById('badge-tutorial').textContent=tutAllTutors.length+' tutors';document.getElementById('tutMeta').textContent=`${tutAllTutors.length} tutors · ${tutAllTutors.reduce((s,t)=>s+t.totalTutees,0)} tutees`;tutRenderSummary();tutRenderTable();updateCombStatus();}catch(err){tutShowError('Error reading file: '+err.message);}};reader.readAsArrayBuffer(file);}
function tutRenderSummary(){const totalH=tutAllTutors.reduce((s,t)=>s+t.hours,0),totalT=tutAllTutors.reduce((s,t)=>s+t.totalTutees,0),avg=totalH/tutAllTutors.length,maxT=tutAllTutors.reduce((a,b)=>a.hours>b.hours?a:b);document.getElementById('tutSummary').innerHTML=[['Tutors',tutAllTutors.length],['Total Tutees',totalT],['Total Hours',totalH],['Avg Hours',avg.toFixed(1)],['Peak Load',`${maxT.hours}h`]].map(([l,v])=>`<div class="tut-stat"><div class="val">${v}</div><div class="lbl">${l}</div></div>`).join('');}
function tutRenderTable(){const q=document.getElementById('tutSearch').value.toLowerCase();let data=tutAllTutors.filter(t=>t.name.toLowerCase().includes(q)||t.email.toLowerCase().includes(q));const colMap={name:'name',email:'email',total:'totalTutees',hours:'hours'},key=colMap[tutSortCol]||tutSortCol;data.sort((a,b)=>{const av=a[key]??0,bv=b[key]??0;return typeof av==='string'?(tutSortDir==='asc'?av.localeCompare(bv):bv.localeCompare(av)):(tutSortDir==='asc'?av-bv:bv-av);});document.getElementById('tutTbody').innerHTML=data.map(t=>`<tr data-name="${encodeURIComponent(t.name)}" style="cursor:pointer"><td class="tutor-name">${t.name}</td><td>${t.email||'—'}</td><td style="text-align:right">${t.year1.length}</td><td style="text-align:right">${t.other.length}</td><td style="text-align:right">${t.totalTutees}</td><td><div class="hours-bar-wrap"><div class="hours-bar"><div class="hours-bar-fill" style="width:${tutMaxHours>0?t.hours/tutMaxHours*100:0}%"></div></div><span class="hours-val">${t.hours}h</span></div></td></tr>`).join('');document.querySelectorAll('#tutTbody tr').forEach(row=>{row.addEventListener('click',()=>{const name=decodeURIComponent(row.dataset.name),tutor=tutAllTutors.find(t=>t.name===name);if(!tutor)return;const y1=tutor.year1.map(s=>`<div class="panel-row"><span class="k">${s.name}</span><span class="v">${s.course||'—'}</span></div>`).join(''),ot=tutor.other.map(s=>`<div class="panel-row"><span class="k">${s.name}</span><span class="v">${s.course||'—'}</span></div>`).join('');openPanel(tutor.name,tutor.email||'',`<div class="panel-section"><div class="panel-row"><span class="k">Year 1 tutees</span><span class="v">${tutor.year1.length}</span></div><div class="panel-row"><span class="k">Other year tutees</span><span class="v">${tutor.other.length}</span></div>${tutor.extraHours>0?`<div class="panel-row"><span class="k">Extra allowance (selected courses)</span><span class="v">${tutor.extraHours}h</span></div>`:''}<div class="panel-row"><span class="k">Total hours</span><span class="v big">${tutor.hours}h</span></div></div>${tutor.year1.length?`<div class="panel-section"><h4>Year 1 Tutees</h4>${y1}</div>`:''}${tutor.other.length?`<div class="panel-section"><h4>Other Tutees</h4>${ot}</div>`:''}`);});});}
document.getElementById('tutSearch').addEventListener('input',tutRenderTable);document.getElementById('tutSort').addEventListener('change',e=>{const[c,d]=e.target.value.split('-');tutSortCol=c;tutSortDir=d;tutRenderTable();});document.getElementById('tutBtnBack').addEventListener('click',()=>{document.getElementById('tut-landing').style.display='';document.getElementById('tut-content').style.display='none';});document.getElementById('tutTable').querySelector('thead').addEventListener('click',e=>{const th=e.target.closest('th[data-tutsort]');if(!th)return;const col=th.dataset.tutsort;if(tutSortCol===col)tutSortDir=tutSortDir==='asc'?'desc':'asc';else{tutSortCol=col;tutSortDir=(col==='hours'||col==='total'||col==='year1'||col==='other')?'desc':'asc';}tutRenderTable();});

// ═══════════════════════════════════════════════════════
// TAB 3 — PROJECT SUPERVISION (unchanged)
// ═══════════════════════════════════════════════════════
let projRawProjects=[],projAllResults=[],projSettings={supervision:12,cosupervision:6,diss_feedback:3,diss_marking:2,poster_feedback:0.5,poster_marking:1/3},projSortKey='total-desc';
const projDropZone=document.getElementById('projDropZone'),projFileInput=document.getElementById('projFileInput'),projAnalyseBtn=document.getElementById('projAnalyseBtn');
projDropZone.addEventListener('dragover',e=>{e.preventDefault();projDropZone.classList.add('drag-over');});projDropZone.addEventListener('dragleave',()=>projDropZone.classList.remove('drag-over'));projDropZone.addEventListener('drop',e=>{e.preventDefault();projDropZone.classList.remove('drag-over');if(e.dataTransfer.files[0])projLoadFile(e.dataTransfer.files[0]);});projFileInput.addEventListener('change',e=>{if(e.target.files[0])projLoadFile(e.target.files[0]);});
document.getElementById('projSettingsHdr').addEventListener('click',()=>{document.getElementById('projSettingsBody').classList.toggle('open');document.getElementById('projSettingsHdr').classList.toggle('open');});
function projShowError(msg){const el=document.getElementById('projError');el.textContent=msg;el.classList.add('show');}function projClearError(){document.getElementById('projError').classList.remove('show');}
function projNormH(s){return String(s).toLowerCase().replace(/[\s_\-]/g,'').replace(/[^a-z0-9]/g,'');}
const PROJ_COL_MAP={theme:['theme','projecttitle','title','project'],supervisor:['supervisor'],cosupervisors:['co_supervisors','cosupervisors','cosupervisor','cosup'],poster1:['poster_assessor_1','poster1assessor','poster_1_assessor','poster1','posterassessor1'],poster2:['poster_assessor_2','poster2assessor','poster_2_assessor','poster2','posterassessor2'],dissertation1:['dissertation_assessor_1','dissertation1assessor','dissertation_1_assessor','diss1','dissertationassessor1'],dissertation2:['dissertation_assessor_2','dissertation2assessor','dissertation_2_assessor','diss2','dissertationassessor2']};
function projFindCol(headers,key){const variants=PROJ_COL_MAP[key];for(let i=0;i<headers.length;i++){const h=projNormH(headers[i]);if(variants.some(v=>h===v))return i;}return -1;}
function projLoadFile(file){projClearError();const reader=new FileReader();reader.onload=e=>{try{let raw;if(file.name.match(/\.csv$/i)){const text=new TextDecoder().decode(e.target.result),lines=text.split(/\r?\n/).filter(l=>l.trim());raw=lines.map(l=>l.split(',').map(c=>c.replace(/^"|"$/g,'').trim()));}else{const wb=XLSX.read(e.target.result,{type:'array'}),ws=wb.Sheets[wb.SheetNames[0]];raw=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});}let headerIdx=0,headerRow=null;for(let i=0;i<Math.min(5,raw.length);i++){const h=raw[i].map(projNormH);if(h.some(c=>c.includes('supervisor')||c.includes('theme')||c.includes('assessor'))){headerIdx=i;headerRow=raw[i];break;}}if(!headerRow){projShowError('Could not find header row.');return;}const iSup=projFindCol(headerRow,'supervisor');if(iSup===-1){projShowError('Could not find Supervisor column.');return;}const iTheme=projFindCol(headerRow,'theme'),iCoSup=projFindCol(headerRow,'cosupervisors'),iP1=projFindCol(headerRow,'poster1'),iP2=projFindCol(headerRow,'poster2'),iD1=projFindCol(headerRow,'dissertation1'),iD2=projFindCol(headerRow,'dissertation2');projRawProjects=[];for(let i=headerIdx+1;i<raw.length;i++){const row=raw[i];if(row.every(c=>!String(c).trim()))continue;const get=idx=>idx!==-1?String(row[idx]||'').trim():'';const splitNames=s=>s.split(/[;,\/|&]+/).map(n=>n.trim()).filter(Boolean);projRawProjects.push({theme:get(iTheme),supervisors:splitNames(get(iSup)),cosupervisors:splitNames(get(iCoSup)),poster1:get(iP1),poster2:get(iP2),diss1:get(iD1),diss2:get(iD2)});}if(projRawProjects.length===0){projShowError('No project rows found.');return;}projAnalyseBtn.disabled=false;projAnalyseBtn.textContent=`🎓 Calculate Project Workload (${projRawProjects.length} projects found) →`;}catch(err){projShowError('Error reading file: '+err.message);}};reader.readAsArrayBuffer(file);}
function projGetSettings(){return{supervision:+document.getElementById('ps_sup').value||0,cosupervision:+document.getElementById('ps_cosup').value||0,diss_feedback:+document.getElementById('ps_diss_fb').value||0,diss_marking:+document.getElementById('ps_diss_mk').value||0,poster_feedback:+document.getElementById('ps_post_fb').value||0,poster_marking:+document.getElementById('ps_post_mk').value||0};}
function projSyncInlineSettings(s){document.getElementById('as_sup').value=s.supervision;document.getElementById('as_cosup').value=s.cosupervision;document.getElementById('as_diss_fb').value=s.diss_feedback;document.getElementById('as_diss_mk').value=s.diss_marking;document.getElementById('as_post_fb').value=s.poster_feedback;document.getElementById('as_post_mk').value=s.poster_marking;}
function projCalculate(projects,s){
  const map={};
  const ensure=name=>{if(!name)return null;if(!map[name])map[name]={name,supervised:[],supShare:[],cosupervised:[],coSupShare:[],diss_assessed:[],poster_assessed:[]};return map[name];};
  for(const p of projects){
    // Supervisors — share load equally among all listed
    const nSups=p.supervisors.length||1;
    for(const sup of p.supervisors){
      if(sup){ensure(sup).supervised.push(p);ensure(sup).supShare.push(1/nSups);}
    }
    // Co-supervisors — share load equally among all listed
    const nCoSups=p.cosupervisors.length||1;
    for(const co of p.cosupervisors){
      if(co){ensure(co).cosupervised.push(p);ensure(co).coSupShare.push(1/nCoSups);}
    }
    if(p.diss1)ensure(p.diss1).diss_assessed.push(p);
    if(p.diss2)ensure(p.diss2).diss_assessed.push(p);
    if(p.poster1)ensure(p.poster1).poster_assessed.push(p);
    if(p.poster2)ensure(p.poster2).poster_assessed.push(p);
  }
  return Object.values(map).map(a=>{
    // Supervision hours: sum of (share × rate) per project
    const h_sup=a.supShare.reduce((t,sh)=>t+sh*s.supervision,0);
    const h_df=a.supShare.reduce((t,sh)=>t+sh*s.diss_feedback,0);
    const h_pf=a.supShare.reduce((t,sh)=>t+sh*s.poster_feedback,0);
    const h_cosup=a.coSupShare.reduce((t,sh)=>t+sh*s.cosupervision,0);
    const nDissAss=a.diss_assessed.length,nPostAss=a.poster_assessed.length;
    const h_dm=nDissAss*s.diss_marking,h_pm=nPostAss*s.poster_marking;
    // Store effective counts as sums of shares for display
    const nSup=+a.supShare.reduce((t,sh)=>t+sh,0).toFixed(4);
    const nCoSup=+a.coSupShare.reduce((t,sh)=>t+sh,0).toFixed(4);
    const total=h_sup+h_cosup+h_df+h_dm+h_pf+h_pm;
    return{...a,nSup,nCoSup,nDissAss,nPostAss,h_sup,h_cosup,h_df,h_dm,h_pf,h_pm,total};
  });
}
function projGetSorted(){const q=document.getElementById('projSearch').value.toLowerCase();let data=projAllResults.filter(r=>r.name.toLowerCase().includes(q));const[col,dir]=projSortKey.split('-');data.sort((a,b)=>{if(col==='name')return dir==='asc'?a.name.localeCompare(b.name):b.name.localeCompare(a.name);if(col==='students')return dir==='desc'?b.nSup-a.nSup:a.nSup-b.nSup;return dir==='desc'?b.total-a.total:a.total-b.total;});return data;}
function projRenderTable(){const data=projGetSorted(),maxTotal=Math.max(...projAllResults.map(r=>r.total),1);document.getElementById('projTbody').innerHTML=data.map(r=>`<tr style="cursor:pointer" data-name="${encodeURIComponent(r.name)}"><td class="name-f">${r.name}</td><td class="num" title="${r.supervised.length} project${r.supervised.length!==1?'s':''}, ${r.nSup.toFixed?r.nSup.toFixed(2):r.nSup} share">${r.supervised.length||'—'}</td><td class="num">${fh(r.h_sup)}</td><td class="num">${fh(r.h_cosup)}</td><td class="num">${fh(r.h_df)}</td><td class="num">${fh(r.h_dm)}</td><td class="num">${fh(r.h_pf)}</td><td class="num">${fh(r.h_pm)}</td><td class="tot">${fmt(r.total)}</td><td><div class="hours-bar-wrap"><div class="hours-bar"><div class="hours-bar-fill" style="width:${r.total/maxTotal*100}%;background:linear-gradient(90deg,var(--rust),var(--gold))"></div></div><span class="hours-val">${fmt(r.total)}h</span></div></td></tr>`).join('');const sumFn=key=>projAllResults.reduce((s,r)=>s+r[key],0);document.getElementById('projFoot').innerHTML=`<tr style="font-weight:600;background:var(--light-blue)"><td>Grand Total</td><td class="num">${projAllResults.reduce((s,r)=>s+r.supervised.length,0)}</td><td class="num">${fmt(sumFn('h_sup'))}</td><td class="num">${fmt(sumFn('h_cosup'))}</td><td class="num">${fmt(sumFn('h_df'))}</td><td class="num">${fmt(sumFn('h_dm'))}</td><td class="num">${fmt(sumFn('h_pf'))}</td><td class="num">${fmt(sumFn('h_pm'))}</td><td class="tot" style="color:var(--mid-blue)">${fmt(sumFn('total'))}</td><td></td></tr>`;document.querySelectorAll('#projTbody tr').forEach(row=>{row.addEventListener('click',()=>projOpenDetail(decodeURIComponent(row.dataset.name)));});}
function projOpenDetail(name){const r=projAllResults.find(x=>x.name===name);if(!r)return;const pills=p=>`${p.supervisors.includes(name)?'<span class="role-pill pill-sup">Supervisor</span>':''}${p.cosupervisors.includes(name)?'<span class="role-pill pill-cosup">Co-supervisor</span>':''}${p.diss1===name||p.diss2===name?'<span class="role-pill pill-diss">Diss. Assessor</span>':''}${p.poster1===name||p.poster2===name?'<span class="role-pill pill-post">Poster Assessor</span>':''}`;const allProjects=[...new Map([...r.supervised,...r.cosupervised,...r.diss_assessed,...r.poster_assessed].map(p=>[p.theme+p.supervisors.join(','),p])).values()];let html=`<div class="panel-section"><h4>Hours Breakdown</h4>${r.h_sup>0?`<div class="panel-row"><span class="k">Supervision (${Number.isInteger(r.nSup)?r.nSup:r.nSup.toFixed(2)} student share${r.nSup!==1?'s':''})</span><span class="v">${fmt(r.h_sup)}h</span></div>`:''}${r.h_cosup>0?`<div class="panel-row"><span class="k">Co-supervision (${r.nCoSup})</span><span class="v">${fmt(r.h_cosup)}h</span></div>`:''}${r.h_df>0?`<div class="panel-row"><span class="k">Dissertation feedback</span><span class="v">${fmt(r.h_df)}h</span></div>`:''}${r.h_dm>0?`<div class="panel-row"><span class="k">Dissertation marking (${r.nDissAss})</span><span class="v">${fmt(r.h_dm)}h</span></div>`:''}${r.h_pf>0?`<div class="panel-row"><span class="k">Poster feedback</span><span class="v">${fmt(r.h_pf)}h</span></div>`:''}${r.h_pm>0?`<div class="panel-row"><span class="k">Poster marking (${r.nPostAss})</span><span class="v">${fmt(r.h_pm)}h</span></div>`:''}<div class="panel-row" style="margin-top:4px"><span class="k"><strong>Total</strong></span><span class="v big">${fmt(r.total)}h</span></div></div>`;if(allProjects.length>0){html+=`<div class="panel-section"><h4>Projects (${allProjects.length})</h4>`;for(const p of allProjects)html+=`<div class="proj-student"><div class="sn">${p.theme||'(No title)'}</div><div class="sr">${pills(p)}</div></div>`;html+='</div>';}openPanel(name,`${fmt(r.total)}h total · ${allProjects.length} project${allProjects.length!==1?'s':''}`,html);}
projAnalyseBtn.addEventListener('click',()=>{projSettings=projGetSettings();projSyncInlineSettings(projSettings);projAllResults=projCalculate(projRawProjects,projSettings);const totalH=projAllResults.reduce((s,r)=>s+r.total,0),nAc=projAllResults.length;document.getElementById('projMeta').textContent=`${projRawProjects.length} projects · ${nAc} academics · ${projAllResults.reduce((s,r)=>s+r.supervised.length,0)} supervision roles`;document.getElementById('projStatsBar').innerHTML=[['Projects',projRawProjects.length],['Academics',nAc],['Total hrs',fmt(totalH)],['Avg hrs',fmt(totalH/nAc)]].map(([l,v])=>`<div class="stat-card rust"><div class="sc-v">${v}</div><div class="sc-l">${l}</div></div>`).join('');document.getElementById('proj-landing').style.display='none';document.getElementById('proj-content').style.display='block';document.getElementById('badge-project').textContent=nAc+' academics';projRenderTable();updateCombStatus();});
document.getElementById('projBtnBack').addEventListener('click',()=>{document.getElementById('proj-landing').style.display='';document.getElementById('proj-content').style.display='none';});document.getElementById('projBtnSettings').addEventListener('click',()=>document.getElementById('projInlineSettings').classList.toggle('open'));document.getElementById('projRecalcBtn').addEventListener('click',()=>{projSettings={supervision:+document.getElementById('as_sup').value||0,cosupervision:+document.getElementById('as_cosup').value||0,diss_feedback:+document.getElementById('as_diss_fb').value||0,diss_marking:+document.getElementById('as_diss_mk').value||0,poster_feedback:+document.getElementById('as_post_fb').value||0,poster_marking:+document.getElementById('as_post_mk').value||0};projAllResults=projCalculate(projRawProjects,projSettings);projRenderTable();updateCombStatus();});
document.getElementById('projSearch').addEventListener('input',projRenderTable);document.getElementById('projSortSel').addEventListener('change',e=>{projSortKey=e.target.value;projRenderTable();});document.querySelector('#proj-content table.proj-table thead').addEventListener('click',e=>{const th=e.target.closest('th[data-projsort]');if(!th)return;const col=th.dataset.projsort;const[curCol,curDir]=projSortKey.split('-');if(curCol===col)projSortKey=col+'-'+(curDir==='desc'?'asc':'desc');else projSortKey=col+'-'+(col==='name'?'asc':'desc');projRenderTable();});
document.getElementById('projBtnExport').addEventListener('click',()=>{const wb=XLSX.utils.book_new();const rows=[['Academic','Supervised (projects)','Sup. share','Co-supervised (projects)','Co-sup. share','Diss.Assessed','Poster Assessed','Sup.hrs','Co-sup.hrs','Diss.Feedback','Diss.Marking','Poster Feedback','Poster Marking','Total']];for(const r of projAllResults)rows.push([r.name,r.supervised.length,+r.nSup.toFixed(2),r.cosupervised.length,+r.nCoSup.toFixed(2),r.nDissAss,r.nPostAss,+r.h_sup.toFixed(2),+r.h_cosup.toFixed(2),+r.h_df.toFixed(2),+r.h_dm.toFixed(2),+r.h_pf.toFixed(2),+r.h_pm.toFixed(2),+r.total.toFixed(2)]);XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(rows),'Project Workload');const sRows=[['Setting','Value'],['Supervision',projSettings.supervision],['Co-supervision',projSettings.cosupervision],['Diss. feedback',projSettings.diss_feedback],['Diss. marking',projSettings.diss_marking],['Poster feedback',projSettings.poster_feedback],['Poster marking',+projSettings.poster_marking.toFixed(4)]];XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(sRows),'Settings');XLSX.writeFile(wb,'project_workload.xlsx');});

// ═══════════════════════════════════════════════════════
// TAB 4 — MMI WORKLOAD
// ═══════════════════════════════════════════════════════
let mmiSessions=[]; // [{sheetName, date, startH, endH, durationH, label, staff:[{name,isReserve}]}]
let mmiResults=[];  // [{name, sessions:[{...session, isReserve}], totalHours, isReserve(ever active)}]
let mmiRawWb=null;
let mmiSortKey='total-desc';

const mmiDropZone=document.getElementById('mmiDropZone'),mmiFileInput=document.getElementById('mmiFileInput'),mmiAnalyseBtn=document.getElementById('mmiAnalyseBtn');
mmiDropZone.addEventListener('dragover',e=>{e.preventDefault();mmiDropZone.classList.add('drag-over');});mmiDropZone.addEventListener('dragleave',()=>mmiDropZone.classList.remove('drag-over'));mmiDropZone.addEventListener('drop',e=>{e.preventDefault();mmiDropZone.classList.remove('drag-over');if(e.dataTransfer.files[0])mmiLoadFile(e.dataTransfer.files[0]);});mmiFileInput.addEventListener('change',e=>{if(e.target.files[0])mmiLoadFile(e.target.files[0]);});

function mmiShowError(msg){const el=document.getElementById('mmiError');el.textContent=msg;el.classList.add('show');}
function mmiClearError(){document.getElementById('mmiError').classList.remove('show');}
function mmiShowWarn(msg){const el=document.getElementById('mmiWarn');el.textContent=msg;el.classList.add('show');}
function mmiClearWarn(){document.getElementById('mmiWarn').classList.remove('show');}

/**
 * Parse tab name → {date, startH, endH, durationH, label}
 * Handles formats like:
 *   "18.03.26 -Online 1.45-4.45pm"
 *   "18.03.26 9am-12pm"
 *   "18.03.26 Morning 9.00-12.00"
 *   "20.03.26 -In Person 2.00-5.00pm"
 */
function parseMmiTabName(name){
  const s=name.trim();

  // Extract date DD.MM.YY or DD/MM/YY at start
  const dateM=s.match(/^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{2,4})/);
  let dateStr='';
  if(dateM){
    const d=dateM[1].padStart(2,'0'),mo=dateM[2].padStart(2,'0');
    let yr=dateM[3];if(yr.length===2)yr='20'+yr;
    dateStr=`${d}/${mo}/${yr}`;
  }

  // Extract label (text between date and time range), stripping leading dashes/spaces
  let remainder=s.replace(/^\d{1,2}[.\/-]\d{1,2}[.\/-]\d{2,4}\s*/,'').replace(/^[-–—\s]+/,'').trim();

  // Time parsing helper: "1.45", "1:45", "13.45", "1" → decimal hours
  // Handles am/pm suffix on either time
  function parseTime(raw, prevHour, pmHint){
    raw=String(raw).trim().toLowerCase();
    const isPm=raw.includes('pm');const isAm=raw.includes('am');
    raw=raw.replace(/[apm]/g,'').trim();
    let h,m=0;
    const parts=raw.split(/[.:]/);
    h=parseInt(parts[0],10);
    if(parts.length>1)m=parseInt(parts[1],10)||0;
    if(isNaN(h))return null;
    // apply am/pm
    if(isPm&&h<12)h+=12;
    if(isAm&&h===12)h=0;
    // if no explicit am/pm, use context
    if(!isPm&&!isAm&&pmHint&&h<12)h+=12;
    if(!isPm&&!isAm&&h<(prevHour||0)&&h>0)h+=12; // times that "go over noon"
    return h+m/60;
  }

  // Match time range: e.g. "1.45-4.45pm" or "9am-12pm" or "13.00-16.00" or "1.45pm-4.45pm"
  const timeRangeM=remainder.match(/([\d.:]+\s*(?:am|pm)?)\s*[-–—]\s*([\d.:]+\s*(?:am|pm)?)\s*$/i);
  if(!timeRangeM) return{dateStr,label:remainder,startH:null,endH:null,durationH:null};

  const rawStart=timeRangeM[1].trim(),rawEnd=timeRangeM[2].trim();
  const label=remainder.slice(0,remainder.length-timeRangeM[0].length).replace(/[-–—\s]+$/,'').trim();

  // Determine pm hint from end time
  const endHasPm=/pm/i.test(rawEnd);
  const startH=parseTime(rawStart,null,endHasPm&&!/am/i.test(rawStart));
  const endH=parseTime(rawEnd,startH,false);

  if(startH===null||endH===null||endH<=startH) return{dateStr,label,startH:null,endH:null,durationH:null};
  const durationH=Math.round((endH-startH)*100)/100;
  return{dateStr,label,startH,endH,durationH};
}

function formatHour(h){
  if(h===null)return'?';
  const hh=Math.floor(h),mm=Math.round((h-hh)*60);
  const period=hh>=12?'pm':'am';const hh12=hh>12?hh-12:hh===0?12:hh;
  return`${hh12}${mm>0?':'+String(mm).padStart(2,'0'):''}${period}`;
}

function mmiLoadFile(file){
  mmiClearError();mmiClearWarn();mmiAnalyseBtn.disabled=true;
  const reader=new FileReader();
  reader.onload=e=>{
    try{
      const wb=XLSX.read(e.target.result,{type:'array'});
      mmiRawWb=wb;
      const parsed=[];const unparsed=[];
      // Only process visible sheets (Hidden===0 or undefined); skip hidden/very-hidden
      const visibleSheets=wb.SheetNames.filter((_,i)=>{
        const s=wb.Workbook?.Sheets?.[i];
        return !s||s.Hidden===0||s.Hidden===undefined;
      });
      for(const sheetName of visibleSheets){
        const info=parseMmiTabName(sheetName);
        if(info.durationH===null){unparsed.push(sheetName);continue;}
        parsed.push({sheetName,...info});
      }
      if(parsed.length===0){mmiShowError('No sheets with parseable time ranges found. Expected tab names like "18.03.26 -Online 1.45-4.45pm".');return;}
      if(unparsed.length>0)mmiShowWarn(`${unparsed.length} sheet${unparsed.length>1?'s':''} skipped (time not recognised): ${unparsed.slice(0,5).join(', ')}${unparsed.length>5?'…':''}`);
      // Show preview
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
    // Find "reserve" row cutoff: scan col A (idx 0) from row 5 (idx 4) onwards for cell containing "reserve"
    let reserveRowIdx=Infinity;
    for(let i=4;i<raw.length;i++){
      const cellA=String(raw[i][0]||'').trim().toLowerCase();
      if(cellA.includes('reserve')){reserveRowIdx=i;break;}
    }
    // Read names from col B (idx 1), rows 5+ (idx 4+), skip blanks
    const staff=[];
    for(let i=4;i<raw.length;i++){
      const name=String(raw[i][1]||'').trim();
      if(!name)continue;
      // Skip if the name itself looks like a header/label (e.g. "Reserve Staff", "Name")
      if(/^(name|staff|reserve|assessor)/i.test(name))continue;
      staff.push({name,isReserve:i>=reserveRowIdx});
    }
    if(staff.length>0)sessions.push({sheetName,...info,staff});
  }
  if(sessions.length===0){mmiShowError('No staff names found in any session sheets. Ensure staff names are in column B from row 5.');return;}
  mmiSessions=sessions;
  // Build per-academic results
  const map={};
  for(const sess of sessions){
    for(const{name,isReserve}of sess.staff){
      if(!map[name])map[name]={name,sessions:[],totalHours:0,isActiveStaff:false};
      map[name].sessions.push({...sess,isReserve});
      if(!isReserve){map[name].totalHours+=sess.durationH;map[name].isActiveStaff=true;}
    }
  }
  mmiResults=Object.values(map).sort((a,b)=>b.totalHours-a.totalHours);
  // Stats
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
  // Session detail sheet
  const sRows=[['Session (Tab)','Date','Start','End','Duration (h)','Staff Name','Is Reserve']];
  for(const sess of mmiSessions)for(const st of sess.staff)sRows.push([sess.sheetName,sess.dateStr,formatHour(sess.startH),formatHour(sess.endH),+sess.durationH.toFixed(2),st.name,st.isReserve?'Yes':'No']);
  XLSX.utils.book_append_sheet(wb2,XLSX.utils.aoa_to_sheet(sRows),'Session Detail');
  XLSX.writeFile(wb2,'mmi_workload.xlsx');
});

// ═══════════════════════════════════════════════════════
// TAB 5 — COMBINED TOTALS
// ═══════════════════════════════════════════════════════
let combData=[],combSortKey='total-desc';
const SRC_LABELS={tl:'📅 Teaching',assessment:'📝 Assessment',proj:'🎓 Project',tut:'👥 Tutorial',mmi:'🩺 MMI',cit:'🏛 Citizenship',res:'🔬 Research',pgr:'👨‍🎓 PGR'};
// staffTags: canonical → Map<tagName, {expiry: Date|null}>
const staffTags=new Map();
// tagRules: tagName → {tlLoad, tlPrep, proj, expiry: Date|null}
const tagRules=new Map();
let activeTagFilter=null;
// Module tags (for filtering only, no rules)
let moduleTagFilter=null;
const moduleTags=new Map(); // normKey(moduleName) → Set<tagName>

// FTE state
let fteTarget=1600; // global target hours (full year; research-active staff use FTE fraction)
const staffFte=new Map(); // normKey → manual override fraction (takes priority over tag fractions)

// Compute effective FTE fraction for a person:
// 1. If manual override set, use that.
// 2. Otherwise, multiply all active tag rule FTE fractions together.
// 3. Default 1.0.
function getEffectiveFte(canonical){
  const nk=normKey(canonical);
  // Manual override takes priority
  const manual=staffFte.get(nk);
  if(manual!=null)return manual;
  // Multiply tag fractions
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
// Build the FTE bar+percentage HTML for a row
function fteBarHtml(canonical,totalHours){
  const pct=ftePct(canonical,totalHours);
  const cls=fteClass(pct);
  // Bar fills proportionally; cap visual at 130% so it doesn't overflow
  const fillPct=Math.min(pct,130)/130*100;
  // Target marker sits at 100/130 = 76.9% of the bar width
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
  pill('comb-status-assessment',hasAssessment,`Assessment: ${Object.keys(assessmentHours).length} staff`,'Assessment');
  pill('comb-status-proj',hasProj,`Project: ${projAllResults.length} academics`,'Project Supervision');
  pill('comb-status-tut',hasTUT,`Tutorial: ${tutAllTutors.length} tutors`,'Tutorial Workload');
  pill('comb-status-mmi',hasMmi,`MMI: ${mmiResults.filter(r=>r.isActiveStaff).length} staff`,'MMIs');
  pill('comb-status-cit',hasCit,`Citizenship: ${Object.keys(citizenshipTotals).length} staff`,'Citizenship');
  pill('comb-status-res',hasRes,`Research: ${Object.keys(resHours).length} staff`,'Research Hours');
  pill('comb-status-pgr',hasPgr,`PGR: ${Object.keys(pgrHours).length} staff`,'PGR Supervision');
  document.getElementById('combMergeBtn').disabled=!(hasTL||hasTUT||hasProj||hasMmi||hasCit||hasRes||hasPgr||hasAssessment);
}

// Recompute hours for all combData entries using current bonuses (no re-merge)
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

document.getElementById('combMergeBtn').addEventListener('click',()=>{
  purgeExpiredAssignments();
  const lists=[];
  if(tlAllStaff.length>0)lists.push({source:'tl',names:tlAllStaff});
  const assessmentHoursTotals=typeof window.getAssessmentHoursTotals==='function'?window.getAssessmentHoursTotals():{};
  const assessmentNames=Object.keys(assessmentHoursTotals);
  if(assessmentNames.length>0)lists.push({source:'assessment',names:assessmentNames});
  if(projAllResults.length>0)lists.push({source:'proj',names:projAllResults.map(r=>r.name)});
  if(tutAllTutors.length>0)lists.push({source:'tut',names:tutAllTutors.map(t=>t.name)});
  const activeMmi=mmiResults.filter(r=>r.isActiveStaff);
  if(activeMmi.length>0)lists.push({source:'mmi',names:activeMmi.map(r=>r.name)});
  const citNames=Object.keys(citizenshipTotals);
  if(citNames.length>0)lists.push({source:'cit',names:citNames});
  const resHoursTotals=typeof window.getResHoursTotals==='function'?window.getResHoursTotals():{};
  const resNames=Object.keys(resHoursTotals);
  if(resNames.length>0)lists.push({source:'res',names:resNames});
  const pgrHoursTotals=typeof window.getPgrHoursTotals==='function'?window.getPgrHoursTotals():{};
  const pgrNames=Object.keys(pgrHoursTotals);
  if(pgrNames.length>0)lists.push({source:'pgr',names:pgrNames});
  const groups=mergeNameLists(lists);
  combData=groups.map(g=>{
    const tlName=g.sources['tl']||null,assessmentName=g.sources['assessment']||null,projName=g.sources['proj']||null,tutName=g.sources['tut']||null,mmiName=g.sources['mmi']||null,citName=g.sources['cit']||null,resName=g.sources['res']||null,pgrName=g.sources['pgr']||null;
    const contactH=tlName?tlAllWeeks.reduce((s,w)=>s+calcHours(tlStaffData[tlName]?.[w],tlRealisticMode),0):0;
    const projBase=projName?(projAllResults.find(r=>r.name===projName)?.total||0):0;
    const{tlTotal,projTotal}=applyBonuses(g.canonical,contactH,projBase);
    const tlHours=tlName?tlTotal:0;
    const tutHours=tutName?(tutAllTutors.find(t=>t.name===tutName)?.hours||0):0;
    const projHours=projName?projTotal:0;
    const mmiHours=mmiName?(mmiResults.find(r=>r.name===mmiName)?.totalHours||0):0;
    const citHours=citName?(citizenshipTotals[citName]||0):0;
    const resHours=resName?(resHoursTotals[resName]||0):0;
    const pgrHours=pgrName?(pgrHoursTotals[pgrName]||0):0;
    const assessmentHours=assessmentName?(assessmentHoursTotals[assessmentName]||0):0;
    const total=tlHours+assessmentHours+projHours+tutHours+mmiHours+citHours+resHours+pgrHours;
    const matchType=Object.keys(g.sources).length>1?(g.matchType||'exact'):'only';
    const _bonuses=computeBonuses(g.canonical);
    return{canonical:g.canonical,tlName,assessmentName,projName,tutName,mmiName,citName,resName,pgrName,tlHours,assessmentHours,projHours,tutHours,mmiHours,citHours,resHours,pgrHours,total,matchType,score:g.score,sources:g.sources,_bonuses};
  });
  const maxTotal=Math.max(...combData.map(d=>d.total),1);
  const fuzzy=combData.filter(d=>d.matchType==='fuzzy').length;
  const firstname=combData.filter(d=>d.matchType==='firstname').length;
  document.getElementById('combMeta').textContent=`${combData.length} academics · ${combData.filter(d=>Object.keys(d.sources).length>1).length} matched across sources · ${fuzzy} fuzzy · ${firstname} first-name matches`;
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
});

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

// Purge expired tag assignments (assignment expiry auto-deletes the assignment)
function purgeExpiredAssignments(){
  const today=todayDate();
  staffTags.forEach((tagMap,canonical)=>{
    tagMap.forEach((info,tag)=>{
      if(info.expiry&&info.expiry<today){
        // If the expired tag's rule had an FTE fraction matching a manual override,
        // clear the override so FTE recalculates from remaining tags.
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
  // Clear active filter if its tag no longer exists
  if(activeTagFilter!==null&&!allTagsSorted().includes(activeTagFilter))activeTagFilter=null;
}

function allTagsSorted(){
  const s=new Set();
  // Tags from assignments
  staffTags.forEach(tagMap=>tagMap.forEach((_,t)=>s.add(t)));
  // Tags from rules (rules can exist without current assignments)
  tagRules.forEach((_,t)=>s.add(t));
  return[...s].sort();
}

function tagsForPerson(canonical){
  // Returns Map<tagName, {expiry}> for this person (only non-expired)
  return staffTags.get(canonical)||new Map();
}

function activeTagsForPerson(canonical){
  // Returns array of tag names that are currently active (not expired)
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
  // If the deleted tag's rule had an FTE fraction, and the person has a
  // manual override matching that fraction, clear the override so FTE
  // recalculates from remaining tags (or defaults to 1.0).
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

// Compute additive bonuses for a person from all their active tag rules
// Returns {tlLoad, tlPrep, proj} — all additive on top of global settings
function computeBonuses(canonical){
  const today=todayDate();
  let tlLoad=0,tlPrep=0,proj=0;
  activeTagsForPerson(canonical).forEach(tag=>{
    const rule=tagRules.get(tag);
    if(!rule)return;
    // Rule expiry: if rule has expired, skip (but rule stays in tagRules for audit)
    if(rule.expiry&&rule.expiry<today)return;
    tlLoad+=rule.tlLoad||0;
    tlPrep+=rule.tlPrep||0;
    proj+=rule.proj||0;
  });
  return{tlLoad,tlPrep,proj};
}

// Apply bonuses to raw hours
// contactHours = pure timetabled contact hours (before any prep)
// projHoursBase = project hours at 1× multiplier
function applyBonuses(canonical,contactHours,projHoursBase){
  const b=computeBonuses(canonical);
  const loadMult=1+b.tlLoad;           // e.g. +1.0 → 2× contact hours
  const prepRatio=tlPrepRatio+b.tlPrep; // global prep ratio + bonus
  const tlTotal=contactHours*loadMult*(1+prepRatio);
  const projTotal=projHoursBase*(1+b.proj);
  return{tlTotal,projTotal,bonuses:b};
}

// ── Module tag helpers (filtering only) ──────────────────────────────────
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
    // Remove overflow restriction so tag chips are visible
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
  // Suggestions: all tags not already on this module
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
    // Wire up live edits
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
  // Update active count badge
  const activeRules=[...tagRules.values()].filter(r=>!r.expiry||r.expiry>=today).length;
  document.getElementById('rulesActiveCount').textContent=tagRules.size>0?`(${activeRules} active, ${tagRules.size} total)`:'';
  // Update datalist suggestions
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
  // Clear inputs
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
  // Populate FTE fraction
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
  // Suggestions: all known tags not already on this person
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

function combRender(maxTotal){
  if(!maxTotal)maxTotal=Math.max(...combData.map(d=>d.total),1);
  const data=combGetSorted();
  // Recalc maxTotal from visible data so bars scale to the filtered subset
  const visMax=Math.max(...data.map(d=>d.total),1);
  const matchBadge=d=>{
    if(d.matchType==='fuzzy')return`<span class="match-badge match-fuzzy">fuzzy${d.score?` ${Math.round(d.score*100)}%`:''}</span>`;
    if(d.matchType==='firstname')return`<span class="match-badge match-fuzzy">first name</span>`;
    if(d.matchType==='only'){const src=Object.keys(d.sources)[0];return`<span class="match-badge match-only">${SRC_LABELS[src]||src} only</span>`;}
    return`<span class="match-badge match-exact">✓ matched</span>`;
  };
  document.getElementById('combTbody').innerHTML=data.map(d=>{
    const enc=encodeURIComponent(d.canonical);
    const chk=combSelected.has(d.canonical)?'checked':'';
    const myTagMap=tagsForPerson(d.canonical);
    const tagHtml=[...myTagMap.entries()].map(([t,info])=>{
      const today=todayDate();
      const expired=info.expiry&&info.expiry<today;
      if(expired)return''; // Don't show expired assignments in table
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
    <td>${matchBadge(d)}</td>
  </tr>`;}).join('');
  // Footer: totals + average FTE%
  const totTL=data.reduce((s,d)=>s+d.tlHours,0),totAssessment=data.reduce((s,d)=>s+d.assessmentHours,0),totProj=data.reduce((s,d)=>s+d.projHours,0),totTut=data.reduce((s,d)=>s+d.tutHours,0),totMmi=data.reduce((s,d)=>s+d.mmiHours,0),totCit=data.reduce((s,d)=>s+d.citHours,0),totRes=data.reduce((s,d)=>s+(d.resHours||0),0),totPgr=data.reduce((s,d)=>s+d.pgrHours,0),totAll=data.reduce((s,d)=>s+d.total,0);
  const avgFte=data.length>0?Math.round(data.reduce((s,d)=>s+ftePct(d.canonical,d.total),0)/data.length):0;
  const avgCls=fteClass(avgFte);
  const filterNote=activeTagFilter?` <span style="font-size:0.72rem;font-weight:400;color:var(--gold);margin-left:6px">tag: ${activeTagFilter} (${data.length})</span>`:'';
  document.getElementById('combFoot').innerHTML=`<tr><td></td><td class="cn">Total${filterNote}</td><td></td><td class="num">${totTL.toFixed(1)}</td><td class="num">${totAssessment.toFixed(1)}</td><td class="num">${totProj.toFixed(1)}</td><td class="num">${totTut.toFixed(1)}</td><td class="num">${totMmi.toFixed(1)}</td><td class="num">${totCit.toFixed(1)}</td><td class="num">${totRes.toFixed(1)}</td><td class="num">${totPgr.toFixed(1)}</td><td class="tot">${totAll.toFixed(1)}</td><td><span style="font-size:0.78rem;font-weight:600" class="fte-pct ${avgCls}">avg ${avgFte}%</span></td><td></td></tr>`;
  // Inline tag-x remove buttons
  document.querySelectorAll('#combTbody .tag-x').forEach(x=>{
    x.addEventListener('click',e=>{e.stopPropagation();const c=decodeURIComponent(x.dataset.canonical),t=decodeURIComponent(x.dataset.tag);removeTag(c,t);recomputeCombData();renderTagFilterBar();renderRulesEditor();saveTagState();combRender();});
  });
  // "+ tag" buttons
  document.querySelectorAll('.comb-tag-add').forEach(btn=>{
    btn.addEventListener('click',e=>{e.stopPropagation();openTagPopover(decodeURIComponent(btn.dataset.canonical),btn);});
  });
  // Checkboxes
  document.querySelectorAll('.comb-chk').forEach(chk=>{chk.addEventListener('change',()=>{const c=decodeURIComponent(chk.dataset.canonical);if(chk.checked)combSelected.add(c);else combSelected.delete(c);combUpdateDetailBtn();const all=document.querySelectorAll('.comb-chk');const allChecked=[...all].every(c=>c.checked);document.getElementById('combSelectAll').checked=allChecked;document.getElementById('combSelectAll').indeterminate=!allChecked&&[...all].some(c=>c.checked);});});
  // Row click → open side panel (click on name cell only)
  document.querySelectorAll('#combTbody tr').forEach(row=>{row.querySelector('.cn').addEventListener('click',()=>{
    const canonical=decodeURIComponent(row.dataset.canonical),d=combData.find(x=>x.canonical===canonical);if(!d)return;
    const tutor=d.tutName?tutAllTutors.find(t=>t.name===d.tutName):null;
    const proj=d.projName?projAllResults.find(r=>r.name===d.projName):null;
    const mmiR=d.mmiName?mmiResults.find(r=>r.name===d.mmiName):null;
    let html=`<div class="panel-section"><h4>Load Summary</h4>
      ${d.tlHours>0?`<div class="panel-row"><span class="k">📅 Teaching</span><span class="v">${d.tlHours.toFixed(1)}h</span></div>`:''}
      ${d.assessmentHours>0?`<div class="panel-row"><span class="k">📝 Assessment</span><span class="v">${d.assessmentHours.toFixed(1)}h</span></div>`:''}
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
        <h3>Assessment Workload</h3>
        <div class="rpt-summary-row">
          <span>Assessments: <strong>${rows.length}</strong></span>
          <span>Total hours: <strong>${d.assessmentHours.toFixed(1)}h</strong></span>
        </div>
        <table class="rpt-table">
          <thead><tr><th>Assessment</th><th>Year</th><th>Course</th><th style="text-align:right">Students</th><th style="text-align:right">Total Load (h)</th><th style="text-align:right">Hours</th></tr></thead>
          <tbody>
          ${rows.map(r=>`<tr><td>${r.assessmentDesc||'—'}</td><td>${r.year||'—'}</td><td>${r.course||'—'}</td><td style="text-align:right">${r.totalStudents||'—'}</td><td style="text-align:right">${r.totalLoad.toFixed(1)}</td><td style="text-align:right">${r.hours.toFixed(1)}</td></tr>`).join('')}
          </tbody>
          <tfoot><tr><td colspan="5"><strong>Total assessment hours</strong></td><td style="text-align:right"><strong>${d.assessmentHours.toFixed(1)}h</strong></td></tr></tfoot>
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
    const cats=[['Teaching',d.tlHours,'#0066cc'],['Assessment',d.assessmentHours,'#8a2be2'],['Projects',d.projHours,'#b84c2a'],['Tutorial',d.tutHours,'#1a7a4a'],['MMI',d.mmiHours,'#6b21a8'],['Citizenship',d.citHours,'#c89b2a'],['Research',(d.resHours||0),'#0a7a9a'],['PGR',d.pgrHours,'#d2691e']].filter(([,h])=>h>0);
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
        <h3>Assessment Workload</h3>
        <div class="rpt-summary-row">
          <span>Assessments: <strong>${rows.length}</strong></span>
          <span>Total hours: <strong>${d.assessmentHours.toFixed(1)}h</strong></span>
        </div>
        <table class="rpt-table">
          <thead><tr><th>Assessment</th><th>Year</th><th>Course</th><th style="text-align:right">Students</th><th style="text-align:right">Total Load (h)</th><th style="text-align:right">Hours</th></tr></thead>
          <tbody>
          ${rows.map(r=>`<tr><td>${r.assessmentDesc||'—'}</td><td>${r.year||'—'}</td><td>${r.course||'—'}</td><td style="text-align:right">${r.totalStudents||'—'}</td><td style="text-align:right">${r.totalLoad.toFixed(1)}</td><td style="text-align:right">${r.hours.toFixed(1)}</td></tr>`).join('')}
          </tbody>
          <tfoot><tr><td colspan="5"><strong>Total assessment hours</strong></td><td style="text-align:right"><strong>${d.assessmentHours.toFixed(1)}h</strong></td></tr></tfoot>
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
    const cats=[['Teaching',d.tlHours,'#0066cc'],['Assessment',d.assessmentHours,'#8a2be2'],['Projects',d.projHours,'#b84c2a'],['Tutorial',d.tutHours,'#1a7a4a'],['MMI',d.mmiHours,'#6b21a8'],['Citizenship',d.citHours,'#c89b2a'],['Research',(d.resHours||0),'#0a7a9a'],['PGR',d.pgrHours,'#d2691e']].filter(([,h])=>h>0);
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
  const rows=[['Academic','Teaching Name','Assessment Name','Project Name','Tutorial Name','MMI Name','Citizenship Name','Research Name','PGR Name','Teaching Hrs','Assessment Hrs','Project Hrs','Tutorial Hrs','MMI Hrs','Citizenship Hrs','Research Hrs','PGR Hrs','Total Hrs','Match Type']];
  for(const d of combData)rows.push([d.canonical,d.tlName||'',d.assessmentName||'',d.projName||'',d.tutName||'',d.mmiName||'',d.citName||'',d.resName||'',d.pgrName||'',+d.tlHours.toFixed(2),+d.assessmentHours.toFixed(2),+d.projHours.toFixed(2),+d.tutHours.toFixed(2),+d.mmiHours.toFixed(2),+d.citHours.toFixed(2),+(d.resHours||0).toFixed(2),+d.pgrHours.toFixed(2),+d.total.toFixed(2),d.matchType]);
  rows.push(['Grand Total','','','','','','','','',+combData.reduce((s,d)=>s+d.tlHours,0).toFixed(2),+combData.reduce((s,d)=>s+d.assessmentHours,0).toFixed(2),+combData.reduce((s,d)=>s+d.projHours,0).toFixed(2),+combData.reduce((s,d)=>s+d.tutHours,0).toFixed(2),+combData.reduce((s,d)=>s+d.mmiHours,0).toFixed(2),+combData.reduce((s,d)=>s+d.citHours,0).toFixed(2),+combData.reduce((s,d)=>s+(d.resHours||0),0).toFixed(2),+combData.reduce((s,d)=>s+d.pgrHours,0).toFixed(2),+combData.reduce((s,d)=>s+d.total,0).toFixed(2),'']);
  XLSX.utils.book_append_sheet(wb2,XLSX.utils.aoa_to_sheet(rows),'Combined Load');
  XLSX.writeFile(wb2,'academic_load_combined.xlsx');
});

updateCombStatus();

// ═══════════════════════════════════════════════════════
// PERSISTENCE — tags, rules & settings (no workload data)
// ═══════════════════════════════════════════════════════
const STORAGE_KEY_TAGS='al_staff_tags_v2';
const STORAGE_KEY_RULES='al_tag_rules_v2';
const STORAGE_KEY_SETTINGS='al_settings_v1';
const STORAGE_KEY_MODTAGS='al_module_tags_v1';

function normKey(canonical){return normaliseName(canonical);}

function saveTagState(){
  // Save tags (keyed by normKey), manual FTE overrides, rules, and global settings
  const tagsArr=[...staffTags.entries()].map(([canonical,tagMap])=>{
    const nk=normKey(canonical);
    return[
      nk,
      [...tagMap.entries()].map(([tag,info])=>[tag,info.expiry?info.expiry.toISOString():null]),
      staffFte.get(nk)??null
    ];
  });
  // Include manual FTE overrides for people with no tags
  staffFte.forEach((frac,nk)=>{
    if(!tagsArr.find(e=>e[0]===nk))tagsArr.push([nk,[],frac]);
  });
  const rulesArr=[...tagRules.entries()].map(([tag,rule])=>[
    tag,
    {tlLoad:rule.tlLoad||0,tlPrep:rule.tlPrep||0,proj:rule.proj||0,fte:rule.fte??1,expiry:rule.expiry?rule.expiry.toISOString():null}
  ]);
  const settings={fteTarget};
  // Fire all three saves; log any failures
  window.storage.set(STORAGE_KEY_TAGS,JSON.stringify(tagsArr))
    .catch(e=>console.warn('AL: tags save failed',e));
  window.storage.set(STORAGE_KEY_RULES,JSON.stringify(rulesArr))
    .catch(e=>console.warn('AL: rules save failed',e));
  window.storage.set(STORAGE_KEY_SETTINGS,JSON.stringify(settings))
    .catch(e=>console.warn('AL: settings save failed',e));
  saveModuleTags();
}
function saveModuleTags(){
  const arr=[...moduleTags.entries()].map(([nk,tags])=>[nk,[...tags]]);
  window.storage.set(STORAGE_KEY_MODTAGS,JSON.stringify(arr))
    .catch(e=>console.warn('AL: module tags save failed',e));
}

let _pendingTagsByNormKey=new Map(); // normKey → {tagMap, manualFte}

function reattachStoredTags(){
  if(_pendingTagsByNormKey.size===0)return;
  combData.forEach(d=>{
    const nk=normKey(d.canonical);
    if(_pendingTagsByNormKey.has(nk)){
      const{tagMap,manualFte}=_pendingTagsByNormKey.get(nk);
      if(!staffTags.has(d.canonical))staffTags.set(d.canonical,new Map());
      tagMap.forEach((info,tag)=>{
        if(!staffTags.get(d.canonical).has(tag))
          staffTags.get(d.canonical).set(tag,info);
      });
      if(manualFte!=null)staffFte.set(nk,manualFte);
      _pendingTagsByNormKey.delete(nk);
    }
  });
  purgeExpiredAssignments();
}

async function loadTagState(){
  let anyLoaded=false;
  // Load settings first (fteTarget needed before render)
  try{
    const settingsRes=await window.storage.get(STORAGE_KEY_SETTINGS);
    if(settingsRes){
      const s=JSON.parse(settingsRes.value);
      if(s.fteTarget){
        fteTarget=s.fteTarget;
        document.getElementById('fteTarget').value=fteTarget;
      }
      anyLoaded=true;
    }
  }catch(e){console.warn('AL: settings load failed',e);}

  // Load rules
  try{
    const rulesRes=await window.storage.get(STORAGE_KEY_RULES);
    if(rulesRes){
      JSON.parse(rulesRes.value).forEach(([tag,rule])=>{
        tagRules.set(tag,{
          tlLoad:rule.tlLoad||0,
          tlPrep:rule.tlPrep||0,
          proj:rule.proj||0,
          fte:rule.fte??1,
          expiry:rule.expiry?new Date(rule.expiry):null
        });
      });
      anyLoaded=true;
    }
  }catch(e){console.warn('AL: rules load failed',e);}

  // Load tags (parked until merge provides canonical names)
  try{
    const tagsRes=await window.storage.get(STORAGE_KEY_TAGS);
    if(tagsRes){
      JSON.parse(tagsRes.value).forEach(entry=>{
        const[nk,tagEntries,manualFte]=entry;
        const tagMap=new Map();
        tagEntries.forEach(([tag,expiryISO])=>
          tagMap.set(tag,{expiry:expiryISO?new Date(expiryISO):null})
        );
        _pendingTagsByNormKey.set(nk,{tagMap,manualFte:manualFte??null});
      });
      anyLoaded=true;
    }
  }catch(e){console.warn('AL: tags load failed',e);}

  // Load module tags
  try{
    const modRes=await window.storage.get(STORAGE_KEY_MODTAGS);
    if(modRes){
      JSON.parse(modRes.value).forEach(([nk,tags])=>{
        moduleTags.set(nk,new Set(tags));
      });
      anyLoaded=true;
    }
  }catch(e){console.warn('AL: module tags load failed',e);}

  if(anyLoaded){
    renderRulesEditor();
    renderTagFilterBar();
    renderModuleTagFilterBar();
    const el=document.createElement('div');
    el.style.cssText='position:fixed;bottom:1rem;right:1rem;background:#041e42;color:white;padding:8px 14px;border-radius:8px;font-size:0.78rem;z-index:9999;opacity:0;transition:opacity 0.3s';
    const pending=_pendingTagsByNormKey.size;
    el.textContent=pending>0
      ?`✓ Settings restored · ${pending} staff group${pending!==1?'s':''} ready to reattach on merge`
      :'✓ Settings restored';
    document.body.appendChild(el);
    requestAnimationFrame(()=>{
      el.style.opacity='1';
      setTimeout(()=>{el.style.opacity='0';setTimeout(()=>el.remove(),400);},2800);
    });
  }
}

// ── localStorage shim for window.storage ─────────────────────────────────────
window.storage = {
  set(key, value) {
    try { localStorage.setItem(key, value); return Promise.resolve(); }
    catch(e) { return Promise.reject(e); }
  },
  get(key) {
    try {
      const v = localStorage.getItem(key);
      return Promise.resolve(v != null ? {value: v} : null);
    } catch(e) { return Promise.reject(e); }
  }
};

// ── Model export / import ─────────────────────────────────────────────────────
function exportModel(){
  const tagsArr=[...staffTags.entries()].map(([canonical,tagMap])=>{
    const nk=normKey(canonical);
    return[nk,[...tagMap.entries()].map(([tag,info])=>[tag,info.expiry?info.expiry.toISOString():null]),staffFte.get(nk)??null];
  });
  staffFte.forEach((frac,nk)=>{
    if(!tagsArr.find(e=>e[0]===nk))tagsArr.push([nk,[],frac]);
  });
  const rulesArr=[...tagRules.entries()].map(([tag,rule])=>[
    tag,{tlLoad:rule.tlLoad||0,tlPrep:rule.tlPrep||0,proj:rule.proj||0,fte:rule.fte??1,expiry:rule.expiry?rule.expiry.toISOString():null}
  ]);
  const payload={version:1,exportedAt:new Date().toISOString(),tags:tagsArr,rules:rulesArr,settings:{fteTarget}};
  const blob=new Blob([JSON.stringify(payload,null,2)],{type:'application/json'});
  const a=document.createElement('a');
  a.href=URL.createObjectURL(blob);
  a.download='academic_load_model.json';
  a.click();
  URL.revokeObjectURL(a.href);
}

function importModel(file){
  const reader=new FileReader();
  reader.onload=e=>{
    try{
      const payload=JSON.parse(e.target.result);
      if(!payload.version||!payload.tags||!payload.rules)throw new Error('Unrecognised format');
      // Restore rules
      tagRules.clear();
      payload.rules.forEach(([tag,rule])=>{
        tagRules.set(tag,{tlLoad:rule.tlLoad||0,tlPrep:rule.tlPrep||0,proj:rule.proj||0,fte:rule.fte??1,expiry:rule.expiry?new Date(rule.expiry):null});
      });
      // Restore FTE target
      if(payload.settings?.fteTarget){
        fteTarget=payload.settings.fteTarget;
        document.getElementById('fteTarget').value=fteTarget;
      }
      // Restore tags (may need reattachment if data not yet loaded)
      _pendingTagsByNormKey.clear();
      staffTags.clear();
      staffFte.clear();
      payload.tags.forEach(([nk,tagEntries,manualFte])=>{
        const tagMap=new Map();
        tagEntries.forEach(([tag,expiryISO])=>tagMap.set(tag,{expiry:expiryISO?new Date(expiryISO):null}));
        _pendingTagsByNormKey.set(nk,{tagMap,manualFte:manualFte??null});
      });
      reattachStoredTags();
      saveTagState();
      recomputeCombData();
      renderRulesEditor();
      renderTagFilterBar();
      const msg=document.createElement('div');
      msg.style.cssText='position:fixed;bottom:1rem;right:1rem;background:#1a7a4a;color:white;padding:8px 14px;border-radius:8px;font-size:0.78rem;z-index:9999;opacity:0;transition:opacity 0.3s';
      msg.textContent='✓ Model imported successfully';
      document.body.appendChild(msg);
      requestAnimationFrame(()=>{msg.style.opacity='1';setTimeout(()=>{msg.style.opacity='0';setTimeout(()=>msg.remove(),400);},2800);});
    }catch(err){alert('Import failed: '+err.message);}
  };
  reader.readAsText(file);
}

document.getElementById('combModelExportBtn').addEventListener('click',exportModel);
document.getElementById('combModelImportInput').addEventListener('change',function(){
  if(this.files[0])importModel(this.files[0]);
  this.value='';
});

loadTagState();


// ── Research Hours data store ──────────────────────────────────────────────────
let resAllData = [];       // array of row objects
let resSortCol = 'currhours', resSortDir = 'desc';
let resYearPref = 'current'; // 'current' | 'next'

// Track which year pref to expose for Combined merge
window.getResHoursTotals = function() {
  // Returns {name: hours} using the selected year preference
  const totals = {};
  resAllData.forEach(r => {
    const hrs = resYearPref === 'next' ? r.nextHours : r.currHours;
    if (r.name && hrs > 0) totals[r.name] = (totals[r.name] || 0) + hrs;
  });
  return totals;
};

// ── PGR Supervision data store ──────────────────────────────────────────────────
let pgrAllData = [];       // array of row objects
let pgrSortCol = 'hours', pgrSortDir = 'desc';
let pgrYearPref = 'current'; // 'current' | 'next' (maybe not needed)

// Track which year pref to expose for Combined merge
window.getPgrHoursTotals = function() {
  // Returns {name: hours} using the selected year preference
  const totals = {};
  pgrAllData.forEach(r => {
    const hrs = r.hours; // each row already contains hours per supervisor
    if (r.supervisor && hrs > 0) totals[r.supervisor] = (totals[r.supervisor] || 0) + hrs;
  });
  return totals;
};

// ── Assessment data store ──────────────────────────────────────────────────
let assessmentAllData = [];       // array of row objects

// Track which year pref to expose for Combined merge
window.getAssessmentHoursTotals = function() {
  // Returns {name: hours}
  const totals = {};
  assessmentAllData.forEach(r => {
    if (r.supervisor && r.hours > 0) totals[r.supervisor] = (totals[r.supervisor] || 0) + r.hours;
  });
  return totals;
};

function parsePgrXlsx(file) {
  console.log('PGR file upload:', file.name);
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      if (rows.length < 2) throw new Error('No data rows found.');

      // Find header row — look for row containing "First Name" or "Surname" (case-insensitive)
      let hdrIdx = 0;
      for (let i = 0; i < Math.min(5, rows.length); i++) {
        const lower = rows[i].map(c => String(c).toLowerCase());
        if (lower.some(c => c.includes('first') || c.includes('surname'))) { hdrIdx = i; break; }
      }
      const hdrs = rows[hdrIdx].map(c => String(c).toLowerCase().trim());

      // Map columns by matching keywords
      function col(keywords) {
        const idx = hdrs.findIndex(h => keywords.some(k => h.includes(k)));
        return idx >= 0 ? idx : -1;
      }
      const cFirstName = col(['first name', 'first']);
      const cSurname = col(['surname', 'last name', 'last']);
      const cStartDate = col(['start date', 'start']);
      const cEsd = col(['esd']);
      const cEndDate = col(['end date', 'end']);
      const cThesisSubmitted = col(['thesis submitted date', 'thesis submitted', 'submitted']);
      const cPlan = col(['plan/programme', 'plan', 'programme']);
      const cPi = col(['pi', 'principal investigator']);
      const cPercentPi = col(['% pi', 'percent pi', '%pi']);
      const cSupervisor2 = col(['2nd supervisor', 'second supervisor']);
      const cPercent2 = col(['% supervisor 2', '% supervisor2', '% 2nd supervisor']);
      const cSupervisor3 = col(['3rd supervisor', 'third supervisor']);
      const cPercent3 = col(['% supervisor 3', '% supervisor3', '% 3rd supervisor']);
      const cSupervisor4 = col(['4th supervisor', 'fourth supervisor']);
      const cPercent4 = col(['% supervisor 4', '% supervisor4', '% 4th supervisor']);
      const cSupervisor5 = col(['5th supervisor', 'fifth supervisor']);
      const cPercent5 = col(['% supervisor 5', '% supervisor5', '% 5th supervisor']);
      const cAssistant = col(['assistant supervisor', 'assistant']);
      const cMode = col(['mode']);

      if (cFirstName < 0 && cSurname < 0) throw new Error('Could not find a "First Name" or "Surname" column.');

      const data = [];
      const today = new Date();
      for (let i = hdrIdx + 1; i < rows.length; i++) {
        const r = rows[i];
        const firstName = cFirstName >= 0 ? String(r[cFirstName] || '').trim() : '';
        const surname = cSurname >= 0 ? String(r[cSurname] || '').trim() : '';
        const studentName = (firstName + ' ' + surname).trim();
        if (!studentName) continue;

        // Parse dates
        const startDate = parseDate(r[cStartDate]);
        const endDate = parseDate(r[cEndDate]);
        // Check if current date is between start and end (inclusive)
        if (startDate && endDate) {
          if (today < startDate || today > endDate) continue;
        } else {
          // If missing dates, assume active
          // continue;
        }

        // Supervisor percentages
        const supervisors = [];
        const addSupervisor = (nameCol, percentCol, defaultPercent = 100) => {
          const name = nameCol >= 0 ? String(r[nameCol] || '').trim() : '';
          if (!name) return;
          let percent = percentCol >= 0 ? parseFloat(r[percentCol]) : defaultPercent;
          if (isNaN(percent) || percent <= 0) percent = defaultPercent;
          supervisors.push({ name, percent });
        };
        addSupervisor(cPi, cPercentPi, 100); // PI defaults to 100%
        addSupervisor(cSupervisor2, cPercent2, 0);
        addSupervisor(cSupervisor3, cPercent3, 0);
        addSupervisor(cSupervisor4, cPercent4, 0);
        addSupervisor(cSupervisor5, cPercent5, 0);
        // Assistant supervisor (no percentage? assume 0% contribution)
        if (cAssistant >= 0) {
          const assistantName = String(r[cAssistant] || '').trim();
          if (assistantName) supervisors.push({ name: assistantName, percent: 0 });
        }

        // Total hours per year per student = 100
        const hoursPerStudent = 100;
        // Distribute among supervisors based on percentages
        supervisors.forEach(s => {
          if (s.percent > 0) {
            const hours = hoursPerStudent * (s.percent / 100);
            data.push({
              studentName,
              supervisor: s.name,
              percent: s.percent,
              hours,
              startDate: startDate ? startDate.toISOString().split('T')[0] : '',
              endDate: endDate ? endDate.toISOString().split('T')[0] : '',
              plan: cPlan >= 0 ? String(r[cPlan] || '').trim() : '',
              mode: cMode >= 0 ? String(r[cMode] || '').trim() : '',
            });
          }
        });
      }
      console.log('PGR parsed rows:', data.length);
      if (data.length === 0) throw new Error('No active PGR students found (current date within start/end dates).');
      renderPgr(data);
      document.getElementById('pgrError').style.display = 'none';
    } catch(err) {
      const el = document.getElementById('pgrError');
      el.textContent = 'Error: ' + err.message;
      el.style.display = 'block';
    }
  };
  reader.readAsArrayBuffer(file);
}

function parseAssessmentXlsx(file) {
  console.log('Assessment file upload:', file.name);
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      if (rows.length < 3) throw new Error('Need at least header rows and one data row.');

      // Double-row header: rows[0] is row 1 (merged headers), rows[1] is row 2 (column headers)
      // Use rows[1] as header row
      const hdrs = rows[1].map(c => String(c).toLowerCase().trim());
      // Column indices
      const cDesc = 0; // Assessment description
      const cYear = 1; // Year of study
      const cCourse = 2; // Course code(s)
      const cLoadPerStudent = 3; // Load per student (minutes)
      const cTotalStudents = 4; // Total students
      const cTotalLoad = 5; // Total assessment load (hours)
      const cAllStaff = 6; // All staff (comma separated)
      const cStaff1 = 7; // Staff (1.0x extra)
      const cExpiry1 = 8; // Expiry (XX/YY/ZZ) for staff1
      const cStaff2 = 9; // Staff (0.5x extra)
      const cExpiry2 = 10; // Expiry (XX/YY/ZZ) for staff2

      const data = [];
      const today = new Date();
      for (let i = 2; i < rows.length; i++) {
        const r = rows[i];
        const totalLoad = parseFloat(r[cTotalLoad]);
        if (isNaN(totalLoad) || totalLoad <= 0) continue;

        // Parse all staff comma-separated list
        const allStaffRaw = String(r[cAllStaff] || '').trim();
        if (!allStaffRaw) continue;
        const allStaff = allStaffRaw.split(',').map(s => s.trim()).filter(s => s);
        if (allStaff.length === 0) continue;

        // Base hours per staff (equal distribution)
        const baseHoursPerStaff = totalLoad / allStaff.length;

        // Map staff -> multiplier (start with 1.0 base)
        const multipliers = new Map();
        allStaff.forEach(name => multipliers.set(name, 1.0));

        // Check 1.0x extra allowance
        const staff1 = String(r[cStaff1] || '').trim();
        const expiry1 = parseDate(r[cExpiry1]);
        if (staff1 && expiry1 && today < expiry1) {
          if (multipliers.has(staff1)) {
            multipliers.set(staff1, multipliers.get(staff1) + 1.0);
          }
        }

        // Check 0.5x extra allowance
        const staff2 = String(r[cStaff2] || '').trim();
        const expiry2 = parseDate(r[cExpiry2]);
        if (staff2 && expiry2 && today < expiry2) {
          if (multipliers.has(staff2)) {
            multipliers.set(staff2, multipliers.get(staff2) + 0.5);
          }
        }

        // Create entries for each staff
        multipliers.forEach((multiplier, name) => {
          const hours = baseHoursPerStaff * multiplier;
          data.push({
            assessmentDesc: r[cDesc],
            year: r[cYear],
            course: r[cCourse],
            loadPerStudent: r[cLoadPerStudent],
            totalStudents: r[cTotalStudents],
            totalLoad,
            supervisor: name,
            hours,
            multiplier,
            staff1,
            expiry1: expiry1 ? expiry1.toISOString().split('T')[0] : '',
            staff2,
            expiry2: expiry2 ? expiry2.toISOString().split('T')[0] : ''
          });
        });
      }
      console.log('Assessment parsed rows:', data.length);
      if (data.length === 0) throw new Error('No valid assessment rows found.');
      renderAssessment(data);
      document.getElementById('assessmentError').style.display = 'none';
    } catch(err) {
      const el = document.getElementById('assessmentError');
      el.textContent = 'Error: ' + err.message;
      el.style.display = 'block';
    }
  };
  reader.readAsArrayBuffer(file);
}

function pgrGetSorted() {
  const q = document.getElementById('pgrSearch')?.value.toLowerCase() || '';
  let data = pgrAllData.filter(r =>
    r.supervisor.toLowerCase().includes(q) || r.studentName.toLowerCase().includes(q)
  );
  const sel = document.getElementById('pgrSortSel')?.value || 'hours-desc';
  const [sc, sd] = sel.split('-');
  data.sort((a, b) => {
    let av, bv;
    if (sc === 'supervisor') { av = a.supervisor; bv = b.supervisor; return sd === 'asc' ? av.localeCompare(bv) : bv.localeCompare(av); }
    if (sc === 'students') {
      // Group by supervisor to count students
      const aStudents = new Set(pgrAllData.filter(d => d.supervisor === a.supervisor).map(d => d.studentName));
      const bStudents = new Set(pgrAllData.filter(d => d.supervisor === b.supervisor).map(d => d.studentName));
      av = aStudents.size;
      bv = bStudents.size;
      return sd === 'asc' ? av - bv : bv - av;
    }
    av = sc === 'hours' ? a.hours : 0;
    bv = sc === 'hours' ? b.hours : 0;
    return sd === 'asc' ? av - bv : bv - av;
  });
  // Deduplicate by supervisor for table display (aggregate hours)
  const aggregated = [];
  const seen = new Set();
  data.forEach(r => {
    if (!seen.has(r.supervisor)) {
      seen.add(r.supervisor);
      const supervisorRows = pgrAllData.filter(d => d.supervisor === r.supervisor);
      const totalHours = supervisorRows.reduce((sum, d) => sum + d.hours, 0);
      const studentSet = new Set(supervisorRows.map(d => d.studentName));
      aggregated.push({
        supervisor: r.supervisor,
        hours: totalHours,
        studentCount: studentSet.size,
      });
    }
  });
  // Re-sort aggregated by original sort criteria
  aggregated.sort((a, b) => {
    let av, bv;
    if (sc === 'supervisor') { av = a.supervisor; bv = b.supervisor; return sd === 'asc' ? av.localeCompare(bv) : bv.localeCompare(av); }
    if (sc === 'students') { av = a.studentCount; bv = b.studentCount; return sd === 'asc' ? av - bv : bv - av; }
    av = a.hours; bv = b.hours;
    return sd === 'asc' ? av - bv : bv - av;
  });
  return aggregated;
}

function renderPgrTable() {
  const data = pgrGetSorted();
  const maxHours = Math.max(...data.map(r => r.hours), 1);
  document.getElementById('pgrTbody').innerHTML = data.map(r => {
    const enc = encodeURIComponent(r.supervisor);
    const barPct = Math.min(r.hours / maxHours * 100, 100).toFixed(1);
    return `<tr>
      <td class="name-f" data-supervisor="${enc}" style="cursor:pointer">${r.supervisor}</td>
      <td class="num">${r.studentCount}</td>
      <td class="num">${r.hours.toFixed(1)}</td>
      <td><div class="pgr-hrs-bar-wrap"><div class="pgr-hrs-bar"><div class="pgr-hrs-bar-fill" style="width:${barPct}%" title="${r.hours.toFixed(1)} hours"></div></div></td>
    </tr>`;
  }).join('');
  // Wire up row clicks for deep-dive
  document.querySelectorAll('#pgrTbody .name-f').forEach(cell => {
    cell.addEventListener('click', () => pgrOpenDetail(decodeURIComponent(cell.dataset.supervisor)));
  });

  // Footer
  const totalHours = data.reduce((s, r) => s + r.hours, 0);
  const totalStudents = new Set(pgrAllData.map(d => d.studentName)).size;
  document.getElementById('pgrFoot').innerHTML = `<tr>
    <td><strong>Total (${data.length} supervisors)</strong></td>
    <td class="num"><strong>${totalStudents}</strong></td>
    <td class="num"><strong>${totalHours.toFixed(1)}</strong></td>
    <td></td>
  </tr>`;
}

function pgrOpenDetail(supervisorName) {
  const rows = pgrAllData.filter(r => r.supervisor === supervisorName);
  if (rows.length === 0) return;
  const totalHours = rows.reduce((s, r) => r.hours + s, 0);
  const studentCount = new Set(rows.map(r => r.studentName)).size;

  let html = `<div class="panel-section"><h4>Hours Breakdown</h4>
    <div class="panel-row"><span class="k">Total hours</span><span class="v big">${totalHours.toFixed(1)}h</span></div>
    <div class="panel-row"><span class="k">Students</span><span class="v">${studentCount}</span></div>
  </div>`;

  html += `<div class="panel-section"><h4>Students (${rows.length} supervisor rows)</h4>`;
  for (const r of rows) {
    const startStr = r.startDate ? r.startDate.toISOString().slice(0, 10) : '—';
    const endStr = r.endDate ? r.endDate.toISOString().slice(0, 10) : '—';
    html += `<div class="proj-student">
      <div class="sn">${r.studentName}</div>
      <div class="sr">
        ${r.plan ? '<span style="margin-right:8px">📋 ' + r.plan + '</span>' : ''}
        ${r.mode ? '<span style="margin-right:8px">📐 ' + r.mode + '</span>' : ''}
        <span style="margin-right:8px">📅 ${startStr} → ${endStr}</span>
        <span style="font-weight:600;color:var(--teal)">${r.hours.toFixed(1)}h @ ${r.percent}%</span>
      </div>
    </div>`;
  }
  html += '</div>';

  openPanel(supervisorName, `${totalHours.toFixed(1)}h total · ${studentCount} student${studentCount !== 1 ? 's' : ''}`, html);
}

function renderPgr(data) {
  pgrAllData = data;

  // Stats bar
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
  document.getElementById('pgrMeta').textContent = `${totalStudents} students · ${totalSupervisors} supervisors`;

  // Update combined status
  updateCombStatus();
}

function assessmentGetSorted() {
  const q = document.getElementById('assessmentSearch')?.value.toLowerCase() || '';
  let data = assessmentAllData.filter(r =>
    r.supervisor.toLowerCase().includes(q)
  );
  const sel = document.getElementById('assessmentSortSel')?.value || 'hours-desc';
  const [sc, sd] = sel.split('-');
  data.sort((a, b) => {
    let av, bv;
    if (sc === 'name') { av = a.supervisor; bv = b.supervisor; return sd === 'asc' ? av.localeCompare(bv) : bv.localeCompare(av); }
    av = a.hours; bv = b.hours;
    return sd === 'asc' ? av - bv : bv - av;
  });
  // Deduplicate by supervisor for table display (aggregate hours)
  const aggregated = [];
  const seen = new Set();
  data.forEach(r => {
    if (!seen.has(r.supervisor)) {
      seen.add(r.supervisor);
      const supervisorRows = assessmentAllData.filter(d => d.supervisor === r.supervisor);
      const totalHours = supervisorRows.reduce((sum, d) => sum + d.hours, 0);
      aggregated.push({
        supervisor: r.supervisor,
        hours: totalHours,
      });
    }
  });
  // Re-sort aggregated by original sort criteria
  aggregated.sort((a, b) => {
    let av, bv;
    if (sc === 'name') { av = a.supervisor; bv = b.supervisor; return sd === 'asc' ? av.localeCompare(bv) : bv.localeCompare(av); }
    av = a.hours; bv = b.hours;
    return sd === 'asc' ? av - bv : bv - av;
  });
  return aggregated;
}

function renderAssessmentTable() {
  const data = assessmentGetSorted();
  const maxHours = Math.max(...data.map(r => r.hours), 1);
  document.getElementById('assessmentTbody').innerHTML = data.map(r => {
    const barPct = Math.min(r.hours / maxHours * 100, 100).toFixed(1);
    return `<tr>
      <td class="name-f">${r.supervisor}</td>
      <td class="num">${r.hours.toFixed(1)}</td>
      <td><div class="pgr-hrs-bar-wrap"><div class="pgr-hrs-bar"><div class="pgr-hrs-bar-fill" style="width:${barPct}%" title="${r.hours.toFixed(1)} hours"></div></div></td>
    </tr>`;
  }).join('');

  // Footer
  const totalHours = data.reduce((s, r) => s + r.hours, 0);
  document.getElementById('assessmentFoot').innerHTML = `<tr>
    <td><strong>Total (${data.length} staff)</strong></td>
    <td class="num"><strong>${totalHours.toFixed(1)}</strong></td>
    <td></td>
  </tr>`;
}

function renderAssessment(data) {
  assessmentAllData = data;

  // Stats bar
  const totalHours = data.reduce((s, r) => s + r.hours, 0);
  const totalStaff = new Set(data.map(r => r.supervisor)).size;
  document.getElementById('assessmentStatsBar').innerHTML = `
    <div class="stat-card teal"><div class="sc-v">${totalStaff}</div><div class="sc-l">Staff with Assessment Load</div></div>
    <div class="stat-card gold"><div class="sc-v">${totalHours.toFixed(0)}</div><div class="sc-l">Total Assessment Hours</div></div>
  `;

  renderAssessmentTable();

  document.getElementById('assessment-landing').style.display = 'none';
  document.getElementById('assessment-content').style.display = 'block';
  document.getElementById('badge-assessment').textContent = totalStaff;
  document.getElementById('assessmentMeta').textContent = `${totalStaff} staff · ${totalHours.toFixed(1)} hours`;

  // Update combined status
  updateCombStatus();
}

function parseDate(val) {
  if (!val) return null;
  // Excel serial number
  if (typeof val === 'number') {
    const excelEpoch = new Date(1899, 11, 30);
    const date = new Date(excelEpoch.getTime() + val * 24 * 60 * 60 * 1000);
    return date;
  }
  // Try parsing as string
  const parsed = new Date(val);
  return isNaN(parsed.getTime()) ? null : parsed;
}



function parseResearchXlsx(file) {
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      if (rows.length < 2) throw new Error('No data rows found.');

      // Find header row — look for row containing "Name" (case-insensitive)
      let hdrIdx = 0;
      for (let i = 0; i < Math.min(5, rows.length); i++) {
        const lower = rows[i].map(c => String(c).toLowerCase());
        if (lower.some(c => c.includes('name'))) { hdrIdx = i; break; }
      }
      const hdrs = rows[hdrIdx].map(c => String(c).toLowerCase().trim());

      // Map columns by matching keywords
      function col(keywords) {
        const idx = hdrs.findIndex(h => keywords.some(k => h.includes(k)));
        return idx >= 0 ? idx : -1;
      }
      const cId   = col(['staff identifier','identifier','staff id','id']);
      const cName = col(['name']);
      const cDept = col(['department','dept']);
      const cFte  = col(['max fte','fte']);
      const cProj = col(['project count','projects']);
      const cHpw  = col(['hrs per week','hrs/week','hours per week']);
      const cCurr = col(['current year','curr']);
      const cNext = col(['next year','next']);

      if (cName < 0) throw new Error('Could not find a "Name" column.');

      const data = [];
      for (let i = hdrIdx + 1; i < rows.length; i++) {
        const r = rows[i];
        const name = String(r[cName] || '').trim();
        if (!name) continue;
        data.push({
          identifier: cId >= 0  ? String(r[cId]  || '').trim() : '',
          name,
          dept:      cDept >= 0 ? String(r[cDept] || '').trim() : '',
          fte:       cFte  >= 0 ? (parseFloat(r[cFte])  || 0) : 0,
          projects:  cProj >= 0 ? (parseInt(r[cProj])   || 0) : 0,
          hrsWeek:   cHpw  >= 0 ? (parseFloat(r[cHpw])  || 0) : 0,
          currHours: cCurr >= 0 ? (parseFloat(r[cCurr]) || 0) : 0,
          nextHours: cNext >= 0 ? (parseFloat(r[cNext]) || 0) : 0,
        });
      }
      if (data.length === 0) throw new Error('No data rows could be parsed.');
      renderResearch(data);
      document.getElementById('resError').style.display = 'none';
    } catch(err) {
      const el = document.getElementById('resError');
      el.textContent = 'Error: ' + err.message;
      el.style.display = 'block';
    }
  };
  reader.readAsArrayBuffer(file);
}

function resGetSorted() {
  const q = document.getElementById('resSearch')?.value.toLowerCase() || '';
  let data = resAllData.filter(r =>
    r.name.toLowerCase().includes(q) || r.dept.toLowerCase().includes(q) || r.identifier.toLowerCase().includes(q)
  );
  const sel = document.getElementById('resSortSel')?.value || 'currhours-desc';
  const [sc, sd] = sel.split('-');
  data.sort((a, b) => {
    let av, bv;
    if (sc === 'name') { av = a.name; bv = b.name; return sd === 'asc' ? av.localeCompare(bv) : bv.localeCompare(av); }
    if (sc === 'dept') { av = a.dept; bv = b.dept; return sd === 'asc' ? av.localeCompare(bv) : bv.localeCompare(av); }
    if (sc === 'identifier') { av = a.identifier; bv = b.identifier; return sd === 'asc' ? av.localeCompare(bv) : bv.localeCompare(av); }
    av = sc === 'fte' ? a.fte : sc === 'projects' ? a.projects : sc === 'hrsweek' ? a.hrsWeek : sc === 'nexthours' ? a.nextHours : a.currHours;
    bv = sc === 'fte' ? b.fte : sc === 'projects' ? b.projects : sc === 'hrsweek' ? b.hrsWeek : sc === 'nexthours' ? b.nextHours : b.currHours;
    return sd === 'asc' ? av - bv : bv - av;
  });
  return data;
}

function renderResTable() {
  const data = resGetSorted();
  const maxCurr = Math.max(...resAllData.map(r => r.currHours), 1);
  document.getElementById('resTbody').innerHTML = data.map(r => {
    const barPct = Math.min(r.currHours / maxCurr * 100, 100).toFixed(1);
    return `<tr>
      <td style="font-family:'IBM Plex Mono',monospace;font-size:0.78rem;color:var(--muted)">${r.identifier || '—'}</td>
      <td class="name-f">${r.name}</td>
      <td>${r.dept || '—'}</td>
      <td class="num">${r.fte > 0 ? r.fte.toFixed(2) : '—'}</td>
      <td class="num">${r.projects > 0 ? r.projects : '—'}</td>
      <td class="num">${r.hrsWeek > 0 ? r.hrsWeek.toFixed(1) : '—'}</td>
      <td class="num">${r.currHours > 0 ? r.currHours.toFixed(1) : '—'}</td>
      <td class="num">${r.nextHours > 0 ? r.nextHours.toFixed(1) : '—'}</td>
      <td><div class="res-hrs-bar-wrap"><div class="res-hrs-bar"><div class="res-hrs-bar-fill" style="width:${barPct}%"></div></div><span class="hours-val">${r.currHours > 0 ? r.currHours.toFixed(0) : '—'}</span></div></td>
    </tr>`;
  }).join('');

  // Footer
  const totCurr = data.reduce((s, r) => s + r.currHours, 0);
  const totNext = data.reduce((s, r) => s + r.nextHours, 0);
  const totProj = data.reduce((s, r) => s + r.projects, 0);
  document.getElementById('resFoot').innerHTML = `<tr>
    <td colspan="3"><strong>Total (${data.length} staff)</strong></td>
    <td class="num">—</td>
    <td class="num"><strong>${totProj}</strong></td>
    <td class="num">—</td>
    <td class="num"><strong>${totCurr.toFixed(1)}</strong></td>
    <td class="num"><strong>${totNext.toFixed(1)}</strong></td>
    <td></td>
  </tr>`;
}

function renderResearch(data) {
  resAllData = data;

  // Stats bar
  const totCurr  = data.reduce((s, r) => s + r.currHours, 0);
  const totNext  = data.reduce((s, r) => s + r.nextHours, 0);
  const avgFte   = data.filter(r => r.fte > 0).reduce((s, r) => s + r.fte, 0) / (data.filter(r => r.fte > 0).length || 1);
  const totProj  = data.reduce((s, r) => s + r.projects, 0);
  document.getElementById('resStatsBar').innerHTML = `
    <div class="stat-card teal"><div class="sc-v">${data.length}</div><div class="sc-l">Staff members</div></div>
    <div class="stat-card gold"><div class="sc-v">${totCurr.toFixed(0)}</div><div class="sc-l">Current Year Total Hrs</div></div>
    <div class="stat-card rust"><div class="sc-v">${totNext.toFixed(0)}</div><div class="sc-l">Next Year Total Hrs</div></div>
    <div class="stat-card purple"><div class="sc-v">${totProj}</div><div class="sc-l">Total Projects</div></div>
    <div class="stat-card"><div class="sc-v">${avgFte.toFixed(2)}</div><div class="sc-l">Avg Max FTE</div></div>
  `;

  renderResTable();

  document.getElementById('res-landing').style.display = 'none';
  document.getElementById('res-content').style.display = 'block';
  document.getElementById('badge-research').textContent = data.length;
  document.getElementById('resMeta').textContent = `${data.length} staff · uploaded`;

  // Update combined status
  updateCombStatus();
}

document.addEventListener('DOMContentLoaded', () => {
  // File input
  const inp = document.getElementById('resFileInput');
  if (inp) inp.addEventListener('change', e => { if (e.target.files[0]) parseResearchXlsx(e.target.files[0]); });

  // Drag and drop
  const dz = document.getElementById('resDropZone');
  if (dz) {
    dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('drag-over'); });
    dz.addEventListener('dragleave', () => dz.classList.remove('drag-over'));
    dz.addEventListener('drop', e => { e.preventDefault(); dz.classList.remove('drag-over'); const f = e.dataTransfer.files[0]; if (f) parseResearchXlsx(f); });
  }

  // Back button
  const back = document.getElementById('resBtnBack');
  if (back) back.addEventListener('click', () => {
    document.getElementById('res-content').style.display = 'none';
    document.getElementById('res-landing').style.display = 'block';
    resAllData = [];
    document.getElementById('badge-research').textContent = '—';
    updateCombStatus();
  });

  // Sort and search
  const sortSel = document.getElementById('resSortSel');
  if (sortSel) sortSel.addEventListener('change', renderResTable);
  const search = document.getElementById('resSearch');
  if (search) search.addEventListener('input', renderResTable);

  // Year preference radio
  document.querySelectorAll('input[name="resYearPref"]').forEach(radio => {
    radio.addEventListener('change', e => {
      resYearPref = e.target.value;
      updateCombStatus();
    });
  });

  // Column header sorting
  document.addEventListener('click', e => {
    const th = e.target.closest('th[data-ressort]');
    if (!th) return;
    const col = th.dataset.ressort;
    const sel = document.getElementById('resSortSel');
    if (!sel) return;
    const isNum = ['fte','projects','hrsweek','currhours','nexthours'].includes(col);
    const dir = (sel.value.startsWith(col) && sel.value.endsWith('desc')) ? 'asc' : (isNum ? 'desc' : 'asc');
    // Try to find matching option
    const opt = [...sel.options].find(o => o.value === `${col}-${dir}`);
    if (opt) { sel.value = opt.value; renderResTable(); }
  });

  // PGR column header sorting
  document.addEventListener('click', e => {
    const th = e.target.closest('th[data-pgrsort]');
    if (!th) return;
    const col = th.dataset.pgrsort;
    const sel = document.getElementById('pgrSortSel');
    if (!sel) return;
    const isNum = ['students','hours'].includes(col);
    const dir = (sel.value.startsWith(col) && sel.value.endsWith('desc')) ? 'asc' : (isNum ? 'desc' : 'asc');
    // Try to find matching option
    const opt = [...sel.options].find(o => o.value === `${col}-${dir}`);
    if (opt) { sel.value = opt.value; renderPgrTable(); }
  });

  // Export button
  const expBtn = document.getElementById('resBtnExport');
  if (expBtn) expBtn.addEventListener('click', () => {
    if (resAllData.length === 0) return;
    const rows = [['Staff Identifier','Name','Department','Max FTE','Project Count','Hrs Per Week','Current Year Hours','Next Year Hours']];
    resGetSorted().forEach(r => rows.push([r.identifier, r.name, r.dept, r.fte, r.projects, r.hrsWeek, r.currHours, r.nextHours]));
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Research Hours');
    XLSX.writeFile(wb, 'staff_research_hours.xlsx');
  });

  // PGR event listeners
  const pgrInp = document.getElementById('pgrFileInput');
  if (pgrInp) pgrInp.addEventListener('change', e => { if (e.target.files[0]) parsePgrXlsx(e.target.files[0]); });

  // Drag and drop
  const pgrDz = document.getElementById('pgrDropZone');
  if (pgrDz) {
    pgrDz.addEventListener('dragover', e => { e.preventDefault(); pgrDz.classList.add('drag-over'); });
    pgrDz.addEventListener('dragleave', () => pgrDz.classList.remove('drag-over'));
    pgrDz.addEventListener('drop', e => { e.preventDefault(); pgrDz.classList.remove('drag-over'); const f = e.dataTransfer.files[0]; if (f) parsePgrXlsx(f); });
  }

  // Back button
  const pgrBack = document.getElementById('pgrBtnBack');
  if (pgrBack) pgrBack.addEventListener('click', () => {
    document.getElementById('pgr-content').style.display = 'none';
    document.getElementById('pgr-landing').style.display = 'block';
    pgrAllData = [];
    document.getElementById('badge-pgr').textContent = '—';
    updateCombStatus();
  });

  // Sort and search
  const pgrSortSel = document.getElementById('pgrSortSel');
  if (pgrSortSel) pgrSortSel.addEventListener('change', renderPgrTable);
  const pgrSearch = document.getElementById('pgrSearch');
  if (pgrSearch) pgrSearch.addEventListener('input', renderPgrTable);

  // Export button
  const pgrExpBtn = document.getElementById('pgrBtnExport');
  if (pgrExpBtn) pgrExpBtn.addEventListener('click', () => {
    if (pgrAllData.length === 0) return;
    const rows = [['Student Name', 'Supervisor', 'Percent', 'Hours', 'Start Date', 'End Date', 'Plan/Programme', 'Mode']];
    // Flatten data
    pgrAllData.forEach(r => rows.push([r.studentName, r.supervisor, r.percent, r.hours, r.startDate, r.endDate, r.plan, r.mode]));
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'PGR Supervision');
    XLSX.writeFile(wb, 'pgr_supervision.xlsx');
  });

  // Assessment event listeners
  const assessmentInp = document.getElementById('assessmentFileInput');
  if (assessmentInp) assessmentInp.addEventListener('change', e => { if (e.target.files[0]) parseAssessmentXlsx(e.target.files[0]); });

  // Drag and drop
  const assessmentDz = document.getElementById('assessmentDropZone');
  if (assessmentDz) {
    assessmentDz.addEventListener('dragover', e => { e.preventDefault(); assessmentDz.classList.add('drag-over'); });
    assessmentDz.addEventListener('dragleave', () => assessmentDz.classList.remove('drag-over'));
    assessmentDz.addEventListener('drop', e => { e.preventDefault(); assessmentDz.classList.remove('drag-over'); const f = e.dataTransfer.files[0]; if (f) parseAssessmentXlsx(f); });
  }

  // Back button
  const assessmentBack = document.getElementById('assessmentBtnBack');
  if (assessmentBack) assessmentBack.addEventListener('click', () => {
    document.getElementById('assessment-content').style.display = 'none';
    document.getElementById('assessment-landing').style.display = 'block';
    assessmentAllData = [];
    document.getElementById('badge-assessment').textContent = '—';
    updateCombStatus();
  });

  // Sort and search
  const assessmentSortSel = document.getElementById('assessmentSortSel');
  if (assessmentSortSel) assessmentSortSel.addEventListener('change', renderAssessmentTable);
  const assessmentSearch = document.getElementById('assessmentSearch');
  if (assessmentSearch) assessmentSearch.addEventListener('input', renderAssessmentTable);

  // Assessment column header sorting
  document.addEventListener('click', e => {
    const th = e.target.closest('th[data-assessmentsort]');
    if (!th) return;
    const col = th.dataset.assessmentsort;
    const sel = document.getElementById('assessmentSortSel');
    if (!sel) return;
    const isNum = ['hours'].includes(col);
    const dir = (sel.value.startsWith(col) && sel.value.endsWith('desc')) ? 'asc' : (isNum ? 'desc' : 'asc');
    // Try to find matching option
    const opt = [...sel.options].find(o => o.value === `${col}-${dir}`);
    if (opt) { sel.value = opt.value; renderAssessmentTable(); }
  });

  // Export button
  const assessmentExpBtn = document.getElementById('assessmentBtnExport');
  if (assessmentExpBtn) assessmentExpBtn.addEventListener('click', () => {
    if (assessmentAllData.length === 0) return;
    const rows = [['Assessment Description', 'Year', 'Course', 'Load per student (min)', 'Total Students', 'Total Load (hours)', 'Staff', 'Hours', 'Multiplier', 'Staff 1.0x', 'Expiry 1', 'Staff 0.5x', 'Expiry 2']];
    assessmentAllData.forEach(r => rows.push([r.assessmentDesc, r.year, r.course, r.loadPerStudent, r.totalStudents, r.totalLoad, r.supervisor, r.hours, r.multiplier, r.staff1, r.expiry1, r.staff2, r.expiry2]));
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Assessment Load');
    XLSX.writeFile(wb, 'assessment_load.xlsx');
  });

});


/* Explicit full-name aliases for citizenship data: maps abbreviated/informal
   names (as they appear in SharePoint) to canonical full names.
   Keys are lower-cased for matching. */
const CIT_NAME_ALIASES = {
  'jing y':     'Dr Jing Yang',
  'jing yang':  'Dr Jing Yang',
  'laura k':    'Dr Laura Kilpatrick',
  'laura kilpatrick': 'Dr Laura Kilpatrick',
  'kati k':     'Dr Katalin Kovacs',
  'katalin kovacs': 'Dr Katalin Kovacs',
};

/* Strip "(Life Sciences)"-style parenthetical suffixes from a name token,
   then apply any explicit alias expansion. */
function citNormaliseName(raw){
  let s = raw.replace(/\s*\(.*?\)\s*/g,'').trim();
  const alias = CIT_NAME_ALIASES[s.toLowerCase()];
  return alias ? alias : s;
}

/* Split "&"-delimited shared roles into individual {name, hours} entries */
function citExpandHolders(rawHolder, totalHours){
  const tokens=rawHolder.split('&').map(t=>citNormaliseName(t)).filter(t=>t.length>0);
  if(tokens.length===0) return [];
  const share=totalHours/tokens.length;
  return tokens.map(name=>({name, hours:share}));
}

function parseSharePointTable(input, category){

  let data=[];

  function pushRow(rawRole, rawHolder, rawHours, rawTerm, rawEnd){
    const totalHours=parseFloat(rawHours)||0;
    const holders=citExpandHolders(rawHolder, totalHours);
    holders.forEach(({name, hours})=>{
      data.push({
        role:rawRole,
        holder:name,
        holderOriginal:rawHolder.trim(),
        hours:hours,
        shared:holders.length>1,
        term:rawTerm,
        end:rawEnd,
        category:category
      });
    });
  }

  /* ---------- CASE 1: HTML TABLE ---------- */

  if(input.includes("<table") || input.includes("<tr")){
    const parser=new DOMParser();
    const doc=parser.parseFromString(input,"text/html");
    const rows=[...doc.querySelectorAll("tr")];
    rows.slice(1).forEach(r=>{
      const cells=[...r.querySelectorAll("td")].map(c=>c.textContent.trim());
      if(cells.length>=5) pushRow(cells[0],cells[1],cells[2],cells[3],cells[4]);
    });
  }

  /* ---------- CASE 2: TAB DELIMITED TEXT ---------- */

  else{
    const rows=input.trim().split("\n");
    rows.slice(1).forEach(line=>{
      const cells=line.split("\t");
      if(cells.length>=5) pushRow(cells[0].trim(),cells[1].trim(),cells[2].trim(),cells[3].trim(),cells[4].trim());
    });
  }

  return data;
}

let citSortCol='hours', citSortDir='desc';
let citAllData=[];

function citGetSorted(){
  const data=[...citAllData];
  data.sort((a,b)=>{
    if(citSortCol==='hours') return citSortDir==='asc'?a.hours-b.hours:b.hours-a.hours;
    const av=a[citSortCol]||'', bv=b[citSortCol]||'';
    return citSortDir==='asc'?av.localeCompare(bv):bv.localeCompare(av);
  });
  return data;
}

function renderCitizenshipTable(){
  document.getElementById("citTbody").innerHTML=citGetSorted().map(r=>`<tr>
    <td>${r.role}</td>
    <td>${r.holder}${r.shared?` <span style="font-size:0.68rem;background:var(--gold-light);color:var(--gold);border-radius:4px;padding:1px 5px;font-weight:600">shared</span>`:''}</td>
    <td style="font-family:'IBM Plex Mono',monospace;text-align:right">${r.hours%1===0?r.hours.toFixed(0):r.hours.toFixed(2)}</td>
    <td>${r.term}</td>
    <td>${r.end}</td>
    <td>${r.category}</td>
  </tr>`).join('');
}

function renderCitizenship(data){
  citizenshipTotals={};
  citAllData=data;
  data.forEach(r=>{
    if(!citizenshipTotals[r.holder]) citizenshipTotals[r.holder]=0;
    citizenshipTotals[r.holder]+=r.hours;
  });
  renderCitizenshipTable();
  document.getElementById("cit-results").style.display="block";
  document.getElementById("badge-citizenship").textContent=data.length;
  updateCombStatus();
}

document.addEventListener("DOMContentLoaded",()=>{

 const btn=document.getElementById("citAnalyseBtn");
 if(!btn) return;

 btn.onclick=function(){
  const teaching=document.getElementById("cit-teaching").value;
  const research=document.getElementById("cit-research").value;
  const school=document.getElementById("cit-school").value;
  let results=[];
  results=results.concat(parseSharePointTable(teaching,"Teaching roles"));
  results=results.concat(parseSharePointTable(research,"Research roles"));
  results=results.concat(parseSharePointTable(school,"School & Citizenship"));

  /* Deduplicate: if the same Role + Role Holder pair appears in more than one
     data source, keep only the first occurrence (preserving its category). */
  const seen=new Set();
  results=results.filter(r=>{
    const key=(r.role+'|||'+r.holder).toLowerCase().trim();
    if(seen.has(key))return false;
    seen.add(key);
    return true;
  });

  renderCitizenship(results);
 };

 document.getElementById("citTable").querySelector("thead").addEventListener("click",e=>{
  const th=e.target.closest("th[data-citsort]");
  if(!th) return;
  const col=th.dataset.citsort;
  if(citSortCol===col) citSortDir=citSortDir==='asc'?'desc':'asc';
  else{citSortCol=col; citSortDir=col==='hours'?'desc':'asc';}
  renderCitizenshipTable();
 });

});

