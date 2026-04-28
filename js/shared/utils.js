// ═══════════════════════════════════════════════════════
// Name normalisation & matching utilities
// ═══════════════════════════════════════════════════════

let citizenshipTotals = {};

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

function nicknameVariants(name){
  const n=name.toLowerCase();
  const vars=new Set([n]);
  if(NICKNAMES[n])NICKNAMES[n].forEach(v=>vars.add(v));
  for(const[k,vs]of Object.entries(NICKNAMES)){if(vs.includes(n)){vars.add(k);vs.forEach(v=>vars.add(v));}}
  return vars;
}

const SURNAME_PREFIXES = ['de','van','von','der','den','di','da','del','dos','du','la','le','mac','mc','o','st','saint','ter','ten','van de','van den','van der','von dem','von der'];

function normaliseName(raw){
  const TITLE_RE=/\b(prof\.?|professor|dr\.?|mr\.?|mrs\.?|ms\.?|mx\.?|rev\.?|sir)\b\s*/i;
  let n=String(raw||'').trim();
  const cm=n.match(/^([^,]+),\s*(.+)$/);if(cm)n=cm[2]+' '+cm[1];
  let prev;
  do{
    prev=n;
    n=n.replace(TITLE_RE,'');
  }while(n!==prev);
  n=n.toLowerCase().replace(/[-.']/g,' ').replace(/\s+/g,' ').trim();
  return n;
}

function nameTokens(raw){
  const norm=normaliseName(raw);
  const tokens=norm.split(' ').filter(Boolean);
  const result=[];
  let i=0;
  while(i<tokens.length){
    if(i<tokens.length-1 && SURNAME_PREFIXES.includes(tokens[i])){
      result.push(tokens[i]+' '+tokens[i+1]);
      i+=2;
    }else{
      result.push(tokens[i]);
      i++;
    }
  }
  return result;
}

function extractSurname(raw){const t=nameTokens(raw);return t.length>0?t[t.length-1]:'';}
function extractFirst(raw){const t=nameTokens(raw);return t.length>0?t[0]:'';}
function extractGivenNames(raw){const t=nameTokens(raw);return t.slice(0,-1);}

function initialMatches(initial,fullToken){return initial.length===1&&fullToken.startsWith(initial);}

function multiInitialMatches(initials, tokens){
  if(!/^[a-z]+$/.test(initials))return false;
  if(initials.length<2)return false;
  if(initials.length>tokens.length)return false;
  for(let i=0;i<initials.length;i++){
    if(!tokens[i].startsWith(initials[i]))return false;
  }
  return true;
}

function tokenAppearsIn(singleToken, allTokens){
  for(const t of allTokens){
    if(t===singleToken)return true;
    if(t.includes(' ') && t.split(' ').includes(singleToken))return true;
  }
  return false;
}

function nameSimilarity(a,b){
  if(!a||!b)return 0;
  const na=normaliseName(a),nb=normaliseName(b);
  if(na===nb)return 1.0;
  const ta=nameTokens(a);
  const tb=nameTokens(b);

  if(ta.length===1 || tb.length===1){
    const singleTokens=ta.length===1?ta:tb;
    const multiTokens=ta.length===1?tb:ta;
    const single=singleTokens[0];
    if(tokenAppearsIn(single, multiTokens))return 0.95;
    if(multiTokens.some(t=>initialMatches(single,t)))return 0.90;
    const vars=nicknameVariants(single);
    for(const v of vars){
      if(tokenAppearsIn(v, multiTokens))return 0.88;
    }
    if(single.length>=3){
      for(const mt of multiTokens){
        if(mt.length>=3 && (mt.includes(single) || single.includes(mt)))return 0.75;
      }
    }
  }

  const sA=new Set(ta),sB=new Set(tb);
  let inter=0;for(const w of sA)if(sB.has(w))inter++;
  const union=sA.size+sB.size-inter;
  const tsr=union===0?0:inter/union;
  if(tsr>=1.0)return 1.0;

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

  const surA=ta[ta.length-1],surB=tb[tb.length-1];
  const firstA=ta[0],firstB=tb[0];
  if(surA&&surB&&surA===surB){
    if(firstA===firstB)return 0.9;
    if(initialMatches(firstA,firstB)||initialMatches(firstB,firstA))return 0.82;
    const givenA=ta.slice(0,-1), givenB=tb.slice(0,-1);
    if(multiInitialMatches(firstA,givenB)||multiInitialMatches(firstB,givenA))return 0.82;
    if(firstA.length>1&&firstA[0]===firstB[0])return 0.78;
    if(firstB.length>1&&firstB[0]===firstA[0])return 0.78;
    const varA=nicknameVariants(firstA),varB=nicknameVariants(firstB);
    if([...varA].some(v=>v===firstB||v.startsWith(firstB))||[...varB].some(v=>v===firstA||v.startsWith(firstA)))return 0.80;
    return 0.55;
  }

  if(surA && surB){
    const surALast=surA.split(' ').pop();
    const surBLast=surB.split(' ').pop();
    if(surALast===surBLast){
      const givenA=ta.slice(0,-1),givenB=tb.slice(0,-1);
      if(givenA.length===0 || givenB.length===0)return 0.85;
      const givenSetA=new Set(givenA),givenSetB=new Set(givenB);
      let gInter=0;
      for(const g of givenSetA)if(givenSetB.has(g))gInter++;
      if(gInter>0)return 0.80+gInter*0.05;
    }
  }

  if(na.length>=3 && nb.length>=3 && (na.includes(nb)||nb.includes(na)))return 0.70;
  return bestNick;
}

function tokenSortRatio(a,b){return nameSimilarity(a,b);}

// ═══════════════════════════════════════════════════════
// Name merge utilities
// ═══════════════════════════════════════════════════════

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
      if(sc>=0.85){
        group.raws.push(names[j]);
        used.add(j);
        if(nameTokens(names[j]).length>nameTokens(group.canonical).length){
          group.canonical=names[j];
        }
      }
    }
    groups.push(group);
  }
  return groups.map(g=>({raw:g.canonical,canonical:g.canonical,raws:g.raws}));
}

function mergeNameLists(lists){
  const preMergedLists=lists.map(({source,names})=>{
    const merged=intraSourceMerge(names,source);
    return{source,names:merged.map(m=>m.raw),canonicalMap:merged.reduce((acc,m)=>{acc[m.raw]=m.canonical;return acc;},{})};
  });

  const THRESH=0.65,all=[];
  for(const{source,names}of preMergedLists)for(const n of names)all.push({norm:normaliseName(n),raw:n,source});
  const groups=[],used=new Set(),key=(r,s)=>`${s}::${r}`;

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

  for(const g of groups){
    if(Object.keys(g.sources).length>1||g._merged)continue;
    const src=Object.keys(g.sources)[0];
    const tokens=normaliseName(g.canonical).split(' ').filter(Boolean);
    if(tokens.length!==1)continue;
    const candidates=groups.filter(og=>{
      if(og===g||og._merged)return false;
      if(og.sources[src])return false;
      const sc=nameSimilarity(g.canonical,og.canonical);
      return sc>=0.70;
    });
    if(candidates.length===1){
      candidates[0].sources[src]=g.canonical;
      candidates[0].matchType='firstname';
      candidates[0].firstnameMatch=true;
      g._merged=true;
    }
  }

  for(const g of groups){
    if(g._merged)continue;
    const allNames=Object.values(g.sources);
    if(allNames.length>1){
      const bestName=allNames.reduce((best,name)=>{
        const score=nameTokens(name).length;
        return score>best.score?{name,score}:best;
      },{name:allNames[0],score:0});
      g.canonical=bestName.name;
    }
  }

  return groups.filter(g=>!g._merged);
}

function normKey(canonical){return normaliseName(canonical);}

// ═══════════════════════════════════════════════════════
// Date / time utilities
// ═══════════════════════════════════════════════════════

function parseDate(val) {
  if (!val) return null;
  if (typeof val === 'number') {
    const excelEpoch = new Date(1899, 11, 30);
    return new Date(excelEpoch.getTime() + val * 24 * 60 * 60 * 1000);
  }
  const parsed = new Date(val);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function parseWeeks(str){
  if(!str)return[];
  const ws=new Set();
  for(const p of String(str).trim().split(',')){
    const t=p.trim(),r=t.match(/^(\d+)\s*[-–—]\s*(\d+)$/);
    if(r){for(let w=+r[1];w<=+r[2];w++)ws.add(w);}
    else{const n=+t;if(!isNaN(n)&&n>0)ws.add(n);}
  }
  return[...ws].sort((a,b)=>a-b);
}

function timeToHours(t){
  if(!t)return 0;
  const m=String(t).match(/(\d+):(\d+)/);
  if(m)return+m[1]+ +m[2]/60;
  const h=parseFloat(t);
  return isNaN(h)?0:h;
}

function sessionDuration(s){
  const d=timeToHours(s.end)-timeToHours(s.start);
  return d>0?Math.round(d*100)/100:0;
}

function calcHours(sessions,realistic){
  if(!sessions||sessions.length===0)return 0;
  if(realistic&&sessions.staffMap){
    let total=0;
    for(const ss of sessions.staffMap.values())total+=calcHours(ss,true);
    return Math.round(total*100)/100;
  }
  if(!realistic)return sessions.reduce((s,x)=>s+sessionDuration(x),0);
  const byDay={};
  for(const s of sessions){
    const d=s.day||'unknown';
    if(!byDay[d])byDay[d]=[];
    byDay[d].push({start:timeToHours(s.start),end:timeToHours(s.end)});
  }
  let total=0;
  for(const slots of Object.values(byDay)){
    const sorted=slots.filter(s=>s.end>s.start).sort((a,b)=>a.start-b.start);
    let merged=[];
    for(const s of sorted){
      if(merged.length&&s.start<merged[merged.length-1].end)merged[merged.length-1].end=Math.max(merged[merged.length-1].end,s.end);
      else merged.push({...s});
    }
    total+=merged.reduce((s,x)=>s+(x.end-x.start),0);
  }
  return Math.round(total*100)/100;
}

function sessionKey(s){return[s.moduleCode,s.moduleTitle,s.sessionTitle,s.day,s.start,s.end,s.type,s.location].join('|');}

function deduplicateSessions(sessions){
  const seen=new Map();
  for(const s of sessions){
    const k=sessionKey(s);
    if(!seen.has(k))seen.set(k,s);
  }
  return[...seen.values()];
}

function formatWeekRange(weeks){
  if(!weeks||weeks.length===0)return'—';
  const sorted=[...weeks].sort((a,b)=>a-b);
  if(sorted.length===1)return'Wk '+sorted[0];
  let contiguous=true;
  for(let i=1;i<sorted.length;i++){if(sorted[i]!==sorted[i-1]+1){contiguous=false;break;}}
  if(contiguous)return'Wk '+sorted[0]+'–'+sorted[sorted.length-1];
  return'Wk '+sorted.join(', ');
}

function formatHour(h){
  if(h===null)return'?';
  const hh=Math.floor(h),mm=Math.round((h-hh)*60);
  const period=hh>=12?'pm':'am';const hh12=hh>12?hh-12:hh===0?12:hh;
  return`${hh12}${mm>0?':'+String(mm).padStart(2,'0'):''}${period}`;
}

// ═══════════════════════════════════════════════════════
// Shared UI helpers
// ═══════════════════════════════════════════════════════

function closePanel(){document.getElementById('detailPanel').classList.remove('open');document.getElementById('overlay').classList.remove('open');}
function openPanel(name,sub,bodyHtml){document.getElementById('panelName').textContent=name;document.getElementById('panelSub').textContent=sub;document.getElementById('panelBody').innerHTML=bodyHtml;document.getElementById('detailPanel').classList.add('open');document.getElementById('overlay').classList.add('open');}

function fh(h){return h>0?h.toFixed(1):'—';}
function fmt(h){return h.toFixed(1);}

// Main tab switching
document.querySelectorAll('.main-tab-btn').forEach(btn=>{btn.addEventListener('click',()=>{
  document.querySelectorAll('.main-tab-btn').forEach(b=>b.classList.remove('active'));
  document.querySelectorAll('.main-tab-panel').forEach(p=>p.classList.remove('active'));
  btn.classList.add('active');
  document.getElementById(btn.dataset.panel).classList.add('active');
});});
