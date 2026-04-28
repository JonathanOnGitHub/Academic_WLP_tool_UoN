// ═══════════════════════════════════════════════════════
// CITIZENSHIP TAB
// ═══════════════════════════════════════════════════════
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
