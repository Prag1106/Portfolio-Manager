// ── REFRESH / SAVE / EXPORT ───────────────────────────────────────────────────
async function refreshPrices(){
  // With manual entry, refresh just recalculates all values
  if(!AID){toast('Nessun cliente','err');return;}
  for(const ptfl of Object.keys(CD.portfolios))recalc(ptfl,getTotal(ptfl),getW(ptfl),getWOrder(ptfl));
  toast('Valori ricalcolati','ok');
}
async function savePtfls(silent=false){
  if(!AID||_saving)return;
  _saving=true;
  const payload=JSON.stringify(CD.portfolios);
  try{
    const r=await fetch(`/api/clients/${AID}/portfolios`,{method:'POST',headers:{'Content-Type':'application/json'},body:payload});
    if(!r.ok)throw new Error('HTTP '+r.status);
    _lastSavedJSON=payload;_lastSavedAt=new Date();
    // keep the instant-open cache aligned with what we just saved
    const now=new Date().toISOString();
    _ptfCache[AID]=JSON.stringify(Object.fromEntries(Object.entries(CD.portfolios).map(([t,d])=>[t,{data:d,updated_at:now}])));
    await loadClients();
    updateSaveInfo();
    if(!silent)toast('Salvato','ok');
  }catch(e){
    if(!silent)toast('Errore di salvataggio — riprova','err');
  }finally{_saving=false;}
}
const EXPORT_KMAP={isin:'isin',ctrv:'ctrv',cg:'cg',sc:'sc',cop:'cop',commTot:'comm_tot',costo:'costo',costoTot:'costo_tot',pctPort:'pct_port',qty:'qty',pctSettore:'pct_settore',pctAc:'pct_ac',ter:'ter',commBar:'comm_bar'};
async function doExport(opts={}){if(!AID){toast('Nessun cliente','err');return;}await savePtfls(true);const p=new URLSearchParams();
  if(opts.doc&&'overviewSheet' in opts.doc)p.set('overview_sheet',opts.doc.overviewSheet?'1':'0');
  Object.entries(opts.types||{}).forEach(([pt,t])=>{
    p.set(pt+'_include',t.include===false?'0':'1');
    Object.entries(t.cols||{}).forEach(([ck,v])=>{const bk=EXPORT_KMAP[ck];if(bk)p.set(pt+'_'+bk,v?'1':'0');});
  });
  p.set('_',Date.now());  // cache-buster: dynamic export must never be served stale from browser cache
  const qs=p.toString();window.location.href=`/api/clients/${AID}/export`+(qs?'?'+qs:'');toast('Download Excel…','info');}
async function renameClient(name){if(!AID||!name.trim())return;await fetch(`/api/clients/${AID}`,{method:'PUT',headers:{'Content-Type':'application/json'},body:JSON.stringify({name:name.trim()})});CD.name=name.trim();document.getElementById('nav-crumb').innerHTML=consCrumb()+`<span>${_esc(name.trim())}</span> › Modifica`;await loadClients();}
function startAutoSave(){clearInterval(autoTimer);_lastSavedJSON=JSON.stringify(CD?.portfolios||{});_lastSavedAt=null;autoTimer=setInterval(()=>{if(AID&&!_saving&&isDirty())savePtfls(true);},30000);}

