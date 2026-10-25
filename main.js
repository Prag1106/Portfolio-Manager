async function loadCgData(){try{const r=await fetch('/api/cg-lookup');const data=await r.json();if(data.pro&&data.chpv)cgData=data;}catch(e){}}
async function loadDartaCats(){try{const r=await fetch('/api/darta-categories');const data=await r.json();if(data.equity)dartaCats=data;}catch(e){}}
async function loadCategoryTree(){try{const r=await fetch('/api/category-tree');const data=await r.json();if(data.equity)catTree=data;}catch(e){}}
async function loadDartaFundsByCat(){try{const r=await fetch('/api/darta-funds-by-category');const data=await r.json();if(data.pro)dartaFundsByCat=data;}catch(e){}}
async function loadSicavFundsByCat(){try{const r=await fetch('/api/sicav-funds-by-category');const data=await r.json();sicavFundsByCat=data;}catch(e){}}

// ── Semaforo sincronizzazione prezzi (pallino sul bottone Titoli DB) ──────────
// verde = sincronizzato oggi · arancio = più di un giorno · rosso = più di 7 giorni
function updateSyncDot(){
  const dot=document.getElementById('sync-dot');if(!dot)return;
  let latest=null;
  [...(_stockUniverse||[]),...(_etfUniverse||[])].forEach(r=>{const u=r&&r.updated_at;if(u&&(!latest||u>latest))latest=u;});
  dot.style.display='inline-block';
  if(!latest){dot.className='sync-dot red';dot.title='Prezzi mai sincronizzati';return;}
  const d=new Date(latest),now=new Date();
  const cls=d.toDateString()===now.toDateString()?'green':(now-d>7*864e5?'red':'orange');
  dot.className='sync-dot '+cls;
  dot.title='Ultima sincronizzazione prezzi: '+d.toLocaleString('it-IT',{day:'2-digit',month:'2-digit',year:'numeric',hour:'2-digit',minute:'2-digit'});
}

// ── Rinnovo PythonAnywhere (bottone in nav) ───────────────────────────────────
async function paStatusRefresh(fresh=false){
  try{
    const r=await fetch('/api/pa/status'+(fresh?'?fresh=1':''));const d=await r.json();
    const btn=document.getElementById('pa-renew-btn');if(!btn)return;
    if(!r.ok||d.error||d.days_left==null)return; // token assente → bottone resta nascosto
    btn.style.display='';
    btn.className='nav-renew'+(d.days_left<=5?' pa-warn':'');
    // freccia circolare di refresh con i giorni rimanenti al centro
    btn.innerHTML=`<svg viewBox="0 0 32 32" width="26" height="26" aria-hidden="true">
      <path d="M16 3 A13 13 0 1 1 4.74 9.5" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round"/>
      <polygon points="6.99,5.6 7.34,11.0 2.14,8.0" fill="currentColor"/>
      <text x="16" y="16.5" text-anchor="middle" dominant-baseline="central" font-size="12" font-weight="800" fill="currentColor">${d.days_left}</text>
    </svg>`;
    btn.title=`Il sito online scade il ${d.expiry} (${d.days_left} giorni) — clicca per aprire PythonAnywhere e premere "Run until 1 month from today"`;
    btn.dataset.expiry=d.expiry||'';
  }catch(e){}
}
// Rinnovo del sito gratuito: da ottobre 2026 PythonAnywhere non lo rinnova più via API
// (disable→enable non sposta la scadenza). Si apre la pagina Web dell'account, dove l'utente
// (già loggato) preme "Run until 1 month from today"; al ritorno il contatore si aggiorna.
const PA_WEB_TAB='https://www.pythonanywhere.com/user/Pragyesh/webapps/#tab_id_pragyesh_pythonanywhere_com';
function paRenew(){
  const btn=document.getElementById('pa-renew-btn');const before=btn?.dataset.expiry||'';
  window.open(PA_WEB_TAB,'_blank','noopener');
  toast('Su PythonAnywhere premi "Run until 1 month from today", poi torna qui','info');
  const onBack=async()=>{
    if(document.visibilityState!=='visible')return;
    document.removeEventListener('visibilitychange',onBack);window.removeEventListener('focus',onBack);
    await paStatusRefresh(true);
    const now=document.getElementById('pa-renew-btn')?.dataset.expiry||'';
    if(now&&now!==before)toast(`✓ Sito rinnovato — scade il ${now}`,'ok');
  };
  setTimeout(()=>{document.addEventListener('visibilitychange',onBack);window.addEventListener('focus',onBack);},500);
}

// ── Aggiornamento prezzi all'apertura della pagina ───────────────────────────
// Chiede al server un refresh (saltato se i prezzi hanno meno di 15 min). In locale
// il bot parte subito; online lo esegue il Mac appena è acceso (controlla ogni 5 min).
// Il pallino di Titoli DB pulsa finché i nuovi prezzi non arrivano.
async function pricesOnOpen(){
  let d;try{const r=await fetch('/api/prices/request',{method:'POST'});if(!r.ok)return;d=await r.json();}catch(e){return;}
  if(d.fresh)return;
  const dot=document.getElementById('sync-dot');const since=d.prices_updated_at||0;
  const busy=on=>{if(!dot)return;dot.classList.toggle('busy',on);if(on)dot.title='Aggiornamento prezzi in corso…'+(d.local?'':' (lo esegue il Mac appena è acceso)');};
  busy(true);
  const t0=Date.now(),limit=(d.local?8:10)*60e3,every=d.local?8000:30000;
  while(Date.now()-t0<limit){
    await new Promise(r=>setTimeout(r,every));
    let s;try{s=await (await fetch('/api/prices/status')).json();}catch(e){continue;}
    if(s.running)continue;
    if(s.prices_updated_at&&s.prices_updated_at>since){
      if(!d.local)await new Promise(r=>setTimeout(r,5000)); // il Mac sincronizza il DB online subito dopo l'upload
      await Promise.all([loadStockUniverse(),loadEtfUniverse()]);
      busy(false);updateSyncDot();
      if(document.getElementById('view-summary')?.classList.contains('active')&&AID){const c=clients.find(x=>x.id===AID);if(c)renderSummary(c);}
      toast('Prezzi aggiornati','ok');return;
    }
    if(d.local&&Date.now()-t0>30000)break; // bot terminato senza aggiornare (Mac chiuso o errore: riprova da solo)
  }
  busy(false);updateSyncDot();
}

// ── INIT ──────────────────────────────────────────────────────────────────────
document.getElementById('nav-date').textContent=new Date().toLocaleDateString('it-IT',{day:'2-digit',month:'long',year:'numeric'});
document.addEventListener('DOMContentLoaded',()=>hydrateIcons());
// Critical path first: dashboard needs clients + universes. Lookup caches and the
// PA status check follow, so they don't compete for the server's single worker.
Promise.all([loadStockUniverse(),loadEtfUniverse(),loadFx(),loadClients(),loadConsultants()]).then(()=>{
  goHome();updateSyncDot();
  loadCgData();
  loadDartaCats();
  loadCategoryTree();
  loadDartaFundsByCat();
  loadSicavFundsByCat();
  paStatusRefresh();
  pricesOnOpen();
});
function goHome(){AID=null;CD=null;CONS=null;document.body.classList.add('at-cons');showView('cons');renderConsultants();document.getElementById('nav-crumb').textContent='Consulenti';renderCliList(clients);}

// ── EXPORT MODAL ──────────────────────────────────────────────────────────────
let _exportFmt='excel';

// Toggleable column labels (keys are the canonical JS option keys)
const XCOL_LABEL={
  isin:'ISIN', ctrv:'CTRV', cg:'CG %', ter:'TER %', sc:'SC %', cop:'Cop. %',
  commTot:'Comm.Tot %', costo:'Costo', costoTot:'Costo Tot.',
  pctPort:'% Portafoglio', qty:'Quantità', pctSettore:'% Settore',
  pctAc:'% Classe (quota nell’asset class)',
  commBar:'Barra commissioni totali',
  costoPost:'Costo totale portafoglio (fine switch)'
};
// Columns that default to OFF (opt-in) in the export modal.
const XCOL_DEFAULT_OFF=new Set(['pctAc']);
// Columns available per portfolio type, per format. Selections are stored
// separately per format, and the available list differs because PDF and Excel
// render different column sets for the same portfolio type.
const XCOLS_PDF={
  amm:['isin','ctrv','pctAc','ter','costo','pctPort','qty'],
  ammAz:['isin','ctrv','pctAc','ter','costo','pctPort','qty'],
  ammEtf:['ctrv','ter','costo','pctPort','qty'],
  ammStock:['ctrv','pctPort','qty'],
  darta:['isin','ctrv','pctAc','cg','sc','cop','commTot','costo','commBar'],
  pac:['isin','ctrv','pctAc','cg','sc','cop','commTot','costo','costoTot','commBar','costoPost'],
  sicav:['isin','ctrv','pctAc','cg','costo','pctPort','qty','commBar'],
  sicavPac:['isin','ctrv','pctAc','cg','costo','costoTot','pctPort','qty','commBar'],
};
const XCOLS_XLS={
  amm:['isin','ctrv','pctAc','ter','costo','pctPort','qty','pctSettore'],
  ammAz:['isin','ctrv','pctAc','ter','costo','pctPort','qty','pctSettore'],
  ammEtf:['ctrv','ter','costo','pctPort','qty'],
  ammStock:['ctrv','pctPort','qty'],
  darta:['isin','ctrv','pctAc','cg','sc','cop','commTot','costo','pctPort','qty','pctSettore','commBar'],
  pac:['isin','ctrv','pctAc','cg','sc','cop','commTot','costo','costoTot','pctPort','qty','pctSettore','commBar'],
  sicav:['isin','ctrv','pctAc','cg','costo','pctPort','qty','pctSettore','commBar'],
  sicavPac:['isin','ctrv','pctAc','cg','costo','costoTot','pctPort','qty','pctSettore','commBar'],
};
function xcolsFor(fmt,pt){return((fmt==='pdf'?XCOLS_PDF:XCOLS_XLS)[pt])||[];}

// Portfolio types (with data) available to export for the current client
function activeExportTypes(){
  const out=[];if(!CD)return out;
  for(const pt of ALL_PTFL){
    const pd=CD.portfolios[pt];if(!pd)continue;
    let total;
    if(pt==='ammEtf'||pt==='ammStock')total=(pd.rows||[]).reduce((s,r)=>s+(r.ctrv||0),0);
    else total=getTotal(pt);
    if(total>0)out.push(pt);
  }
  return out;
}

function _exportStoreKey(fmt){return 'pm_export_opts_'+fmt;}
function _loadExportOpts(fmt){try{return JSON.parse(localStorage.getItem(_exportStoreKey(fmt))||'{}');}catch(e){return{};}}

function openExportModal(fmt){
  _exportFmt=fmt;
  document.getElementById('xm-fmt-badge').textContent=fmt==='pdf'?'PDF':fmt==='pptx'?'PowerPoint':'Excel';
  document.getElementById('xm-fmt-badge').style.background=fmt==='pdf'?'#C03060':fmt==='pptx'?'#C43E1C':'#1A6E3C';
  document.getElementById('xm-types-wrap').style.display=fmt==='pptx'?'none':'';
  if(fmt==='pptx'){openPptxOptions();return;}
  const saved=_loadExportOpts(fmt);
  const savedDoc=saved.doc||{};const savedTypes=saved.types||{};

  // ── Document-level options (format specific) ──
  const docBox=document.getElementById('xm-doc-group');
  if(fmt==='pdf'){
    docBox.innerHTML=`
      <div class="xm-sec-hd"><span>Layout documento</span><span class="xm-tag" style="background:#7A1A3C">PDF</span></div>
      <div class="xm-doc-list">
        <label class="xm-lbl"><input type="checkbox" data-doc="overview" ${savedDoc.overview!==false?'checked':''}> Riepilogo iniziale (donut portafogli)</label>
        <label class="xm-lbl"><input type="checkbox" data-doc="sectionDonut" ${savedDoc.sectionDonut!==false?'checked':''}> Donut per sezione</label>
      </div>`;
  }else{
    docBox.innerHTML=`
      <div class="xm-sec-hd"><span>Fogli documento</span><span class="xm-tag" style="background:#1A6E3C">Excel</span></div>
      <div class="xm-doc-list">
        <label class="xm-lbl"><input type="checkbox" data-doc="overviewSheet" ${savedDoc.overviewSheet!==false?'checked':''}> Foglio "Panoramica generale"</label>
      </div>`;
  }

  // ── Per-type accordions ──
  const types=activeExportTypes();
  const acc=document.getElementById('xm-types');
  if(!types.length){
    acc.innerHTML='<div style="font-size:12px;color:#8892a8;padding:6px 2px">Nessun portafoglio con dati da esportare.</div>';
  }else{
    acc.innerHTML=types.map((pt,idx)=>{
      const st=savedTypes[pt]||{};
      const included=st.include!==false;
      const cols=xcolsFor(fmt,pt);
      const colsHtml=cols.map(ck=>{
        const saved=st.cols&&(ck in st.cols)?st.cols[ck]:undefined;
        const checked=saved!==undefined?saved!==false:!XCOL_DEFAULT_OFF.has(ck);
        return `<label class="xm-lbl"><input type="checkbox" class="xm-col" data-pt="${pt}" data-col="${ck}" ${checked?'checked':''}> ${XCOL_LABEL[ck]||ck}</label>`;
      }).join('');
      const body=cols.length?`<div class="xm-col-grid">${colsHtml}</div>`:'<div style="font-size:11px;color:#5a6a82;padding:2px 0">Nessuna colonna opzionale.</div>';
      return `
      <div class="xm-acc${included?'':' excluded'}" data-pt="${pt}">
        <div class="xm-acc-hd">
          <label class="xm-acc-inc"><input type="checkbox" class="xm-inc" data-pt="${pt}" ${included?'checked':''} onchange="xmToggleInc('${pt}')"> <span class="xm-acc-name">${TL[pt]||pt}</span></label>
          <button type="button" class="xm-acc-tg" onclick="xmAccToggle(this)">${idx===0?'▾':'▸'}</button>
        </div>
        <div class="xm-acc-body" style="display:${idx===0?'block':'none'}">${body}</div>
      </div>`;
    }).join('');
    types.forEach(pt=>applyIncState(pt,(savedTypes[pt]||{}).include!==false));
  }

  renderExportWarnings();
  document.getElementById('export-modal').style.display='flex';
}

function xmAccToggle(btn){
  const body=btn.closest('.xm-acc').querySelector('.xm-acc-body');
  const open=body.style.display!=='none';
  body.style.display=open?'none':'block';
  btn.textContent=open?'▸':'▾';
}
function applyIncState(pt,inc){
  const acc=document.querySelector(`.xm-acc[data-pt="${pt}"]`);if(!acc)return;
  acc.classList.toggle('excluded',!inc);
  acc.querySelectorAll('.xm-col').forEach(i=>i.disabled=!inc);
}
function xmToggleInc(pt){applyIncState(pt,document.querySelector(`.xm-inc[data-pt="${pt}"]`).checked);}

// Banner di avviso allocazioni nel modale di export
function renderExportWarnings(){
  const box=document.getElementById('xm-alloc-warn');if(!box)return;
  const warns=getAllocWarnings();
  if(!warns.length){box.style.display='none';box.innerHTML='';return;}
  const rows=warns.map(w=>{
    const delta=Math.round((w.sum-100)*100)/100;
    const over=delta>0;
    const sign=over?'+':'';
    const col=over?'#ff8a80':'#ffcf6b';
    const verdict=over?'sopra il 100%':'sotto il 100%';
    return `<div style="display:flex;justify-content:space-between;gap:10px;padding:3px 0;font-size:12px"><span style="color:#cfd6e4"><strong>${w.label}</strong> · ${w.scope}</span><span style="color:${col};font-weight:700;font-family:'JetBrains Mono',monospace;white-space:nowrap">${w.sum.toFixed(2)}% &nbsp;(${sign}${delta.toFixed(2)}% — ${verdict})</span></div>`;
  }).join('');
  box.style.display='block';
  box.innerHTML=`<div style="display:flex;align-items:center;gap:8px;margin-bottom:6px"><span style="font-size:16px">⚠️</span><span style="font-size:12px;font-weight:800;letter-spacing:.4px;color:#ffd27a;text-transform:uppercase">Allocazioni non bilanciate</span></div><div>${rows}</div><div style="font-size:11px;color:#9aa6bc;margin-top:7px;line-height:1.5">Le voci sopra non sommano al 100%. Puoi scaricare comunque, ma controlla i pesi prima di consegnare il documento al cliente.</div>`;
}

function closeExportModal(){
  document.getElementById('export-modal').style.display='none';
}

function getExportOpts(){
  const opts={doc:{},types:{}};
  document.querySelectorAll('#export-modal input[data-doc]').forEach(cb=>{opts.doc[cb.dataset.doc]=cb.checked;});
  document.querySelectorAll('#export-modal .xm-acc').forEach(acc=>{
    const pt=acc.dataset.pt;
    const inc=acc.querySelector('.xm-inc')?.checked!==false;
    const cols={};
    acc.querySelectorAll('.xm-col').forEach(cb=>{cols[cb.dataset.col]=cb.checked;});
    opts.types[pt]={include:inc,cols};
  });
  try{localStorage.setItem(_exportStoreKey(_exportFmt),JSON.stringify(opts));}catch(e){}
  return opts;
}

// PowerPoint: the deck is DARTA-only; the modal just asks for the Start & Go
// terms that the app doesn't otherwise store (saved on darta.pac per client).
function openPptxOptions(){
  const docBox=document.getElementById('xm-doc-group');
  const pd=CD?.portfolios?.darta;
  if(!pd||!(pd.total>0)){
    docBox.innerHTML='<div style="font-size:12px;color:#ffcf6b;padding:4px 2px">La presentazione PowerPoint è disponibile solo per i clienti con una polizza DARTA.</div>';
  }else{
    const cfg=pd.pac||{};const sg=!!(pd.weights||{})['START & GO'];
    const sw=!!(cfg.enabled&&cfg.mesi>0);
    const rate=cfg.sgRate!=null?cfg.sgRate:(sg?'4,50':'');
    const delay=cfg.sgDelay!=null?cfg.sgDelay:(sg&&sw?3:0);
    const inp='width:90px;padding:5px 8px;border-radius:7px;border:1px solid #25395a;background:#0d1a30;color:#e8eaf0;font-size:13px;text-align:right';
    docBox.innerHTML=`
      <div class="xm-sec-hd"><span>Presentazione cliente</span><span class="xm-tag" style="background:#C43E1C">PPTX</span></div>
      <div style="font-size:12px;color:#aab4c8;line-height:1.5;margin-bottom:10px">Proposta di portafoglio DARTA in stile Allianz: asset allocation, strumenti, costi${sw?', allocazione a fine switch':''} e dettagli.</div>
      ${sg||sw?`<div class="xm-doc-list">
        ${sg?`<label class="xm-lbl" style="justify-content:space-between">Tasso Start &amp; Go (% annuo) <input id="xm-sg-rate" type="text" inputmode="decimal" value="${_esc(String(rate))}" placeholder="—" style="${inp}"></label>`:''}
        ${sw?`<label class="xm-lbl" style="justify-content:space-between">Mesi prima del primo switch <input id="xm-sg-delay" type="number" min="0" max="24" step="1" value="${_esc(String(delay))}" style="${inp}"></label>`:''}
      </div>`:''}`;
  }
  document.getElementById('xm-types').innerHTML='';
  document.getElementById('xm-alloc-warn').style.display='none';
  document.getElementById('export-modal').style.display='flex';
}
function confirmExport(){
  if(_exportFmt==='pptx'){
    const pd=CD?.portfolios?.darta;
    const rEl=document.getElementById('xm-sg-rate'),dEl=document.getElementById('xm-sg-delay');
    const o={};
    if(pd&&(rEl||dEl)){if(!pd.pac)pd.pac={};
      if(rEl){pd.pac.sgRate=rEl.value.trim();o.sgRate=pd.pac.sgRate;}
      if(dEl){pd.pac.sgDelay=Math.max(0,parseInt(dEl.value,10)||0);o.sgDelay=pd.pac.sgDelay;}}
    closeExportModal();downloadPptx(o);return;
  }
  const opts=getExportOpts();
  closeExportModal();
  if(_exportFmt==='excel')doExport(opts);
  else downloadPdf(opts);
}
