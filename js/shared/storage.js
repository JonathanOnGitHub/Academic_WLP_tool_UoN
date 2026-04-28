// ═══════════════════════════════════════════════════════
// Persistent storage — tags, rules, settings, model export/import
// ═══════════════════════════════════════════════════════

const STORAGE_KEY_TAGS='al_tags_v2';
const STORAGE_KEY_RULES='al_tag_rules_v2';
const STORAGE_KEY_SETTINGS='al_settings_v1';
const STORAGE_KEY_MODTAGS='al_module_tags_v1';
const STORAGE_KEY_MAPPINGS='al_manual_mappings_v1';

let _pendingTagsByNormKey=new Map(); // normKey → {tagMap, manualFte}

// localStorage shim
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

function saveTagState(){
  const tagsArr=[...staffTags.entries()].map(([canonical,tagMap])=>{
    const nk=normKey(canonical);
    return[
      nk,
      [...tagMap.entries()].map(([tag,info])=>[tag,info.expiry?info.expiry.toISOString():null]),
      staffFte.get(nk)??null
    ];
  });
  staffFte.forEach((frac,nk)=>{
    if(!tagsArr.find(e=>e[0]===nk))tagsArr.push([nk,[],frac]);
  });
  const rulesArr=[...tagRules.entries()].map(([tag,rule])=>[
    tag,
    {tlLoad:rule.tlLoad||0,tlPrep:rule.tlPrep||0,proj:rule.proj||0,fte:rule.fte??1,expiry:rule.expiry?rule.expiry.toISOString():null}
  ]);
  const settings={fteTarget};
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

function saveManualMappings(){
  const arr=[...manualMappings.entries()].map(([fromNk,toName])=>[fromNk,toName]);
  window.storage.set(STORAGE_KEY_MAPPINGS,JSON.stringify(arr))
    .catch(e=>console.warn('AL: mappings save failed',e));
}

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

  try{
    const modRes=await window.storage.get(STORAGE_KEY_MODTAGS);
    if(modRes){
      JSON.parse(modRes.value).forEach(([nk,tags])=>{
        moduleTags.set(nk,new Set(tags));
      });
      anyLoaded=true;
    }
  }catch(e){console.warn('AL: module tags load failed',e);}

  // Load manual name mappings
  try{
    const mapRes=await window.storage.get(STORAGE_KEY_MAPPINGS);
    if(mapRes){
      JSON.parse(mapRes.value).forEach(([fromNk,toName])=>{
        manualMappings.set(fromNk,toName);
      });
      anyLoaded=true;
    }
  }catch(e){console.warn('AL: mappings load failed',e);}

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
  const payload={version:1,exportedAt:new Date().toISOString(),tags:tagsArr,rules:rulesArr,settings:{fteTarget},mappings:[...manualMappings.entries()].map(([nk,t])=>[nk,t])};
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
      tagRules.clear();
      payload.rules.forEach(([tag,rule])=>{
        tagRules.set(tag,{tlLoad:rule.tlLoad||0,tlPrep:rule.tlPrep||0,proj:rule.proj||0,fte:rule.fte??1,expiry:rule.expiry?new Date(rule.expiry):null});
      });
      if(payload.settings?.fteTarget){
        fteTarget=payload.settings.fteTarget;
        document.getElementById('fteTarget').value=fteTarget;
      }
      _pendingTagsByNormKey.clear();
      staffTags.clear();
      staffFte.clear();
      payload.tags.forEach(([nk,tagEntries,manualFte])=>{
        const tagMap=new Map();
        tagEntries.forEach(([tag,expiryISO])=>tagMap.set(tag,{expiry:expiryISO?new Date(expiryISO):null}));
        _pendingTagsByNormKey.set(nk,{tagMap,manualFte:manualFte??null});
      });
      manualMappings.clear();
      if(payload.mappings)payload.mappings.forEach(([nk,t])=>manualMappings.set(nk,t));
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
