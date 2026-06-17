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

/* Parse an XLSX file dropped onto the citizenship tab.
   Looks for columns: Role, Role Holder, Hours, Term, End Date. */
function parseCitizenshipXlsx(file){
  const reader=new FileReader();
  reader.onload=e=>{
    try{
      const wb=XLSX.read(new Uint8Array(e.target.result),{type:'array'});
      const ws=wb.Sheets[wb.SheetNames[0]];
      const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});
      if(rows.length<2) throw new Error('No data rows found.');

      /* Find header row — look for a "Role" or "Teaching" keyword */
      let hdrIdx=0;
      for(let i=0;i<Math.min(5,rows.length);i++){
        const lower=rows[i].map(c=>String(c).toLowerCase());
        if(lower.some(c=>c.includes('role')||c.includes('teaching')||c.includes('citizenship'))){hdrIdx=i;break;}
      }
      const hdrs=rows[hdrIdx].map(c=>String(c).toLowerCase().trim());

      function col(keywords){
        const idx=hdrs.findIndex(h=>keywords.some(k=>h.includes(k)));
        return idx>=0?idx:-1;
      }
      const cRole =col(['role','teaching role','citizenship role','leadership role','activity']);
      const cHolder=col(['role holder','holder','name','staff','person']);
      const cHours =col(['hours','hrs','hours per week','hpw','workload','wlc']);
      const cTerm  =col(['term','duration','length','period']);
      const cEnd   =col(['end date','end','expiry','expires','date']);

      if(cRole<0) throw new Error('Could not find a "Role" column.');
      if(cHolder<0) throw new Error('Could not find a "Role Holder" column.');

      let cat=file.name.toLowerCase().includes('teach')
        ?'Teaching roles':file.name.toLowerCase().includes('research')
        ?'Research roles':'School & Citizenship';

      const all=[];
      for(let i=hdrIdx+1;i<rows.length;i++){
        const r=rows[i];
        const role=String(r[cRole]||'').trim();
        const holder=String(r[cHolder]||'').trim();
        if(!role||!holder) continue;                 /* skip blank rows */
        const hoursStr=String(cHours>=0?r[cHours]:'0').trim();
        const term=String(cTerm>=0?r[cTerm]:'').trim();
        const end=String(cEnd>=0?r[cEnd]:'').trim();
        const totalHours=parseFloat(hoursStr)||0;
        const tokens=citExpandHolders(holder,totalHours);
        tokens.forEach(({name,hours:hrs})=>{
          all.push({role,holder:name,holderOriginal:holder,hours:hrs,shared:tokens.length>1,term,end,category:cat});
        });
      }

      if(!all.length) throw new Error('No valid data rows found after parsing.');

      /* Deduplicate the same way the Analyse button does */
      const seen=new Set();
      const deduped=all.filter(r=>{
        const key=(r.role+'|||'+r.holder).toLowerCase().trim();
        if(seen.has(key))return false;
        seen.add(key);
        return true;
      });

      renderCitizenship(deduped);
    } catch(err){
      alert('Error reading spreadsheet: '+err.message);
    }
  };
  reader.readAsArrayBuffer(file);
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

  /* ---------- CASE 2: PLAIN TEXT (tab / comma / multi-space) ---------- */
  else{
    const trimmed=input.trim();
    if(!trimmed) return data;

    /* Try delimiters in order – return first that gives 5+ columns */
    function splitRow(s){
      let c=s.split("\t");
      if(c.length>=5) return c;
      if((s.match(/,/g)||[]).length>=4){
        c=s.split(",");
        if(c.length>=5) return c;
      }
      c=s.split(/\s{2,}/);
      if(c.length>=5) return c;
      return null;
    }

    const KNOWN_HDR=new Set(["role","role holder","hours for wlp","hours","term","end date","end"]);
    function isHdr(cells){return cells.length>0 && KNOWN_HDR.has(cells[0].toLowerCase().trim());}

    const lines=trimmed.split("\n");
    const rows=[];
    let acc=[];

    /* Try to reconstruct a 5-column row from the accumulator.
       Only clears acc on success. */
    function tryFlush(){
      if(!acc.length) return false;

      /* A – join with tab, try all delimiters on the result */
      let cells=splitRow(acc.join("\t"));

      /* B – 4+ accumulated lines, each line ≈ one field.
         The last line often holds 2 sub-fields (term & end) separated
         by 2+ spaces, giving the 5th column. */
      if(!cells && acc.length>=4){
        const lastCells=acc[acc.length-1].split(/\s{2,}/);
        const rebuilt=[...acc.slice(0,-1),...lastCells];
        if(rebuilt.length>=5) cells=rebuilt;
      }

      if(cells){
        rows.push(cells);
        acc=[];
        return true;
      }
      return false;
    }

    for(const line of lines){
      const l=line.trim();
      if(!l) continue;           /* blank lines are NOT row boundaries here;
                                    they only separate fields in multi-line rows */
      const cells=splitRow(l);
      if(cells){
        tryFlush();              /* flush any prior fragment */
        rows.push(cells);
      } else {
        acc.push(l);
        if(acc.length>=4) tryFlush();  /* try to assemble a multi-line row */
      }
    }
    tryFlush();                   /* trailing fragment */

    /* Skip rows whose first cell is a known header label */
    rows.forEach(cells=>{
      if(isHdr(cells)) return;
      pushRow(cells[0].trim(),cells[1].trim(),cells[2].trim(),cells[3].trim(),cells[4].trim());
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

 /* ── XLSX upload ── */
 const citInp=document.getElementById('citFileInput');
 if(citInp) citInp.addEventListener('change',e=>{if(e.target.files[0]) parseCitizenshipXlsx(e.target.files[0]);});
 const citDz=document.getElementById('citDropZone');
 if(citDz){
   citDz.addEventListener('dragover',e=>{e.preventDefault();citDz.classList.add('drag-over');});
   citDz.addEventListener('dragleave',()=>citDz.classList.remove('drag-over'));
   citDz.addEventListener('drop',e=>{e.preventDefault();citDz.classList.remove('drag-over');const f=e.dataTransfer.files[0];if(f) parseCitizenshipXlsx(f);});
 }

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
