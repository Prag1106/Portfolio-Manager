// ── EDIT VIEW ─────────────────────────────────────────────────────────────────
function renderEdit(){
  const el=document.getElementById('view-edit');el.innerHTML='';if(!CD)return;
  const h=document.createElement('div');h.className='edit-header';
  h.innerHTML=`<div style="display:flex;align-items:center;gap:10px"><button class="btn btn-ghost" onclick="goSummary()" style="font-size:11px;padding:5px 10px">← Riepilogo</button><span class="edit-title">Modifica — ${_esc(CD.name)}</span></div><div style="display:flex;align-items:center;gap:6px"><button id="undo-btn" class="btn btn-ghost" onclick="doUndo()" disabled title="Annulla l'ultima modifica (Ctrl+Z)" style="font-size:11px;padding:5px 12px">↶ Annulla</button><button id="redo-btn" class="btn btn-ghost" onclick="doRedo()" disabled title="Ripristina (Ctrl+Shift+Z)" style="font-size:11px;padding:5px 12px">↷ Ripristina</button></div>`;
  el.appendChild(h);
  // baseline per la cronologia annulla: cattura lo stato prima di ogni modifica di un campo
  el.addEventListener('focusin',e=>{if(e.target.matches('input,select')){_flushUndo();_undoBaseline=JSON.stringify(CD.portfolios);}});
  el.addEventListener('keydown',fundRowKeydown);
  _undo=[];_redo=[];_undoBaseline=null;updateUndoBtns();
  const setup=document.createElement('div');setup.className='card';
  setup.innerHTML=`<div class="card-head"><span class="card-head-title">Configurazione Cliente</span></div><div class="card-body"><input class="name-inp" type="text" value="${_esc(CD.name)}" placeholder="Nome cliente" onchange="renameClient(this.value)"/><div style="margin-top:.875rem"><div class="sec-title">Portafogli attivi</div><div class="ptfl-grid" id="ptfl-grid"></div></div></div>`;
  el.appendChild(setup);
  const totEl=document.createElement('div');totEl.className='totals';totEl.id='totals-strip';el.appendChild(totEl);
  const tabBar=document.createElement('div');tabBar.className='tab-bar';tabBar.id='tab-bar';el.appendChild(tabBar);
  ALL_PTFL.forEach(t=>{const p=document.createElement('div');p.className='tab-pane';p.id='tp-'+t;el.appendChild(p);});
  const sb=document.createElement('div');sb.className='save-bar';
  sb.innerHTML=`<span class="save-info">Cliente: <strong>${_esc(CD.name)}</strong></span><div class="save-actions"><button class="btn-out" title="Excel" onclick="openExportModal('excel')">${ic('sheet')} Excel</button><button class="btn-out" title="PDF" onclick="openExportModal('pdf')">${ic('file')} PDF</button><button class="btn-out" title="PowerPoint" onclick="openExportModal('pptx')">${ic('slides')} PowerPoint</button><button class="btn-out" title="Duplica" onclick="openDup(${AID},CD.name)">${ic('copy')} Duplica</button><button class="btn-save" onclick="savePtfls()">${ic('save')} Salva</button><button class="btn-danger" onclick="openDel(${AID},CD.name)" title="Elimina">${ic('trash')} Elimina</button></div>`;
  el.appendChild(sb);
  const enabled=ALL_PTFL.filter(t=>!isPacType(t)&&!!CD.portfolios[t]);
  renderPtflCards(enabled);renderTotals(enabled);renderTabs(enabled);
}

// ── PORTFOLIO TYPE CARDS ──────────────────────────────────────────────────────
function renderPtflCards(enabled){
  const g=document.getElementById('ptfl-grid');if(!g)return;
  const types=[{id:'ammStock',icon:'trend',name:'Amministrato Azioni',desc:'Titoli azionari da Excel',cls:''},{id:'ammEtf',icon:'bars',name:'Amministrato ETF',desc:'ETF da Excel',cls:''},{id:'darta',icon:'clipboard',name:'DARTA + Easy Switch',desc:'Incl. START & GO → Easy Switch',cls:'sg-sel'},{id:'sicav',icon:'bank',name:'SICAV & Fondi',desc:'Fondi e SICAV di terzi',cls:''},{id:'sicavPac',icon:'coins',name:'PAC',desc:'Piano di accumulo · totale ÷ mesi',cls:''}];
  g.innerHTML=types.map(t=>{const on=enabled.includes(t.id);return`<div class="ptfl-card${on?' on':''}${on&&t.cls?' '+t.cls:''}" onclick="togglePtfl('${t.id}')"><div class="ptfl-chk">${on?'✓':''}</div><div class="ptfl-icon">${ic(t.icon)}</div><div class="ptfl-name">${t.name}</div><div class="ptfl-desc">${t.desc}</div></div>`;}).join('');
}
function togglePtfl(type){recordUndo();
  if(!CD._ptflBackup)CD._ptflBackup={};
  const on=!!CD.portfolios[type];
  const child=PAC_CHILD[type];
  if(on){
    CD._ptflBackup[type]=JSON.parse(JSON.stringify(CD.portfolios[type]));
    delete CD.portfolios[type];
    if(child){CD._ptflBackup[child]=JSON.parse(JSON.stringify(CD.portfolios[child]||{}));delete CD.portfolios[child];}
  }else{
    if(CD._ptflBackup[type]){CD.portfolios[type]=CD._ptflBackup[type];delete CD._ptflBackup[type];}
    else{CD.portfolios[type]=makeDefault(type);}
    if(child){
      if(CD._ptflBackup[child]){CD.portfolios[child]=CD._ptflBackup[child];delete CD._ptflBackup[child];}
      else{CD.portfolios[child]=makeDefault(child);}
    }
  }
  const enabled=ALL_PTFL.filter(t=>!isPacType(t)&&!!CD.portfolios[t]);
  renderPtflCards(enabled);renderTotals(enabled);renderTabs(enabled);
}
const DEFAULT_EQ={'Azioni Globale':40,'Azioni USA':20,'Azioni Europa':30,'Azioni Giappone':5,'Azioni Paesi Emergenti':5};
const DEFAULT_EQ_ORD=['Azioni Globale','Azioni USA','Azioni Europa','Azioni Giappone','Azioni Paesi Emergenti'];
const DEFAULT_BD={'Obbligazioni Euro':35,'Obbligazioni Globale':25,'Obbligazioni Paesi Emergenti':20,'Monetari Europa':20};
const DEFAULT_BD_ORD=['Obbligazioni Euro','Obbligazioni Globale','Obbligazioni Paesi Emergenti','Monetari Europa'];
function makeDefault(type){
  const EQ=DEFAULT_EQ, EQ_ORD=DEFAULT_EQ_ORD, BD=DEFAULT_BD, BD_ORD=DEFAULT_BD_ORD;
  const D={
    amm:{total:0,_mode:'weights',weights:{'ETF Azionari':{pct:60,subs:{},_type:'equity'},'ETF Obbligazionari':{pct:30,subs:{},_type:'bond'},'ETC':{pct:10,subs:{}}},rows:{},_order:['ETF Azionari','ETF Obbligazionari','ETC']},
    ammAz:{total:0,_mode:'weights',weights:{'Azioni':{pct:100,subs:{},_type:'equity'}},rows:{},_order:['Azioni']},
    ammStock:{rows:[]},
    ammEtf:{rows:[]},
    darta:{total:0,_mode:'weights',plan:'pro',weights:{},rows:{},pac:{enabled:false,sourceKey:null,totalSwitch:0,mesi:null,useRounded:false,roundedTotal:0},_order:[]},
    // Easy Switch parte vuoto come DARTA: asset class e sotto-settori si aggiungono a mano
    pac:{total:0,_mode:'weights',weights:{},rows:{},_order:[]},
    sicav:{total:0,_mode:'weights',weights:{},rows:{},_order:[]},
    // PAC parte vuoto come DARTA ed Easy Switch
    sicavPac:{total:0,mesi:null,_mode:'weights',weights:{},rows:{},_order:[]}
  };
  return JSON.parse(JSON.stringify(D[type]||D.amm));
}

// ── TOTALS ────────────────────────────────────────────────────────────────────
function renderTotals(enabled){
  const el=document.getElementById('totals-strip');if(!el)return;
  // Only show amm/ammAz in totals strip if client actually has that legacy data
  const hasLegacyAmm=!!(CD?.portfolios?.amm);const hasLegacyAmmAz=!!(CD?.portfolios?.ammAz);
  const totalsItems=[
    ...(hasLegacyAmm?[{id:'amm',lbl:'Amm. ETF',sub:'ETF e ETC'}]:[]),
    ...(hasLegacyAmmAz?[{id:'ammAz',lbl:'Amm. Azioni',sub:'Titoli azionari'}]:[]),
    {id:'ammStock',lbl:'Amm. Azioni',sub:'Da Excel'},
    {id:'ammEtf',lbl:'Amm. ETF',sub:'Da Excel'},
    {id:'darta',lbl:'DARTA',sub:'Include START & GO'},
    {id:'pac',lbl:'Easy Switch — Durata',sub:'Rata = START & GO ÷ mesi',mesi:true,parent:'darta'},
    {id:'sicav',lbl:'SICAV & Fondi',sub:'Fondi e SICAV'},
    {id:'sicavPac',lbl:'PAC',sub:'Totale ÷ mesi = rata',pacStd:true}
  ];
  el.innerHTML=totalsItems.map(c=>{
    let on=enabled.includes(c.id)||(c.mesi&&enabled.includes(c.parent));
    if(c.id==='pac')on=on&&!!CD.portfolios?.darta?.pac?.enabled;
    if(c.pacStd){const pd=CD.portfolios?.[c.id]||{};const tot=pd.total||0;const ms=pd.mesi||null;const rata=sicavPacRata();return`<div class="tc pac-c${!on?' off':''}"><div class="tc-lbl g">${c.lbl}</div><div class="tc-row"><span class="tc-pre">€</span><input class="tc-inp w130" type="text" inputmode="numeric" value="${fmtEuroInp(tot)}" ${!on?'disabled':''} onfocus="onEuroFocus(this)" onchange="setTotal('${c.id}',this.value);onEuroBlur(this)"/></div><div class="tc-row" style="margin-top:5px"><input class="tc-inp w55" type="number" min="1" step="1" value="${ms||''}" ${!on?'disabled':''} onchange="setSicavPacMesi(this.value)"/><span style="font-size:12px;color:var(--gold-l);font-weight:500;margin-left:5px">mesi</span></div><div class="tc-sub">Rata: ${fmt(rata)}/mese</div><div id="tcbar-${c.id}">${ctrvBar(c.id,rata)}</div></div>`;}
    if(c.mesi){const ms=CD.portfolios?.[c.parent]?.pac?.mesi||null;const rata=pacTotal(c.id);return`<div class="tc pac-c${!on?' off':''}"><div class="tc-lbl g">${c.lbl}</div><div class="tc-row"><input class="tc-inp w55" type="number" min="1" step="1" value="${ms||''}" ${!on?'disabled':''} onchange="setPacMesi('${c.parent}',this.value)"/><span style="font-size:12px;color:var(--gold-l);font-weight:500;margin-left:5px">mesi</span></div><div class="tc-sub">Rata: ${fmt(rata)}/mese</div></div>`;}
    const total=CD.portfolios?.[c.id]?.total||0;
    return`<div class="tc${!on?' off':''}"><div class="tc-lbl">${c.lbl}</div><div class="tc-row"><span class="tc-pre">€</span><input class="tc-inp w130" type="text" inputmode="numeric" value="${fmtEuroInp(total)}" ${!on?'disabled':''} onfocus="onEuroFocus(this)" onchange="setTotal('${c.id}',this.value);onEuroBlur(this)"/></div><div class="tc-sub">${c.sub}</div><div id="tcbar-${c.id}">${ctrvBar(c.id,total)}</div></div>`;
  }).join('');
}
function setTotal(ptfl,v){if(!CD.portfolios[ptfl])return;CD.portfolios[ptfl].total=parseEuro(v);const mode=getPD(ptfl)?._mode||'weights';if(mode==='ctrv')deriveWeightsFromRows(ptfl);else buildTab(ptfl);if(PAC_CHILD[ptfl])updatePac(ptfl);if(ptfl==='sicavPac'){const e=ALL_PTFL.filter(t=>!isPacType(t)&&!!CD.portfolios[t]);renderTotals(e);}}
function setPacMesi(parent,v){if(!CD.portfolios[parent])return;if(!CD.portfolios[parent].pac)CD.portfolios[parent].pac={};const n=parseInt(v);CD.portfolios[parent].pac.mesi=isNaN(n)||v===''?null:n;const e=ALL_PTFL.filter(t=>!isPacType(t)&&!!CD.portfolios[t]);renderTotals(e);updatePac(parent);}
// PAC standalone: months entered directly on the sicavPac portfolio.
function setSicavPacMesi(v){const pd=CD.portfolios?.sicavPac;if(!pd)return;const n=parseInt(v);pd.mesi=isNaN(n)||v===''?null:n;const e=ALL_PTFL.filter(t=>!isPacType(t)&&!!CD.portfolios[t]);renderTotals(e);buildTab('sicavPac');}
function setPacUseRounded(child,checked){const par=CD.portfolios[PAC_PARENT[child]];if(!par)return;if(!par.pac)par.pac={};const cfg=par.pac;cfg.useRounded=checked;if(checked&&!(cfg.roundedTotal>0)){const a=pacActual(child);cfg.roundedTotal=Math.floor(a/100)*100||Math.round(a);}updatePac(PAC_PARENT[child]);}
function setPacRoundedTotal(child,v){const par=CD.portfolios[PAC_PARENT[child]];if(!par)return;if(!par.pac)par.pac={};par.pac.roundedTotal=parseEuro(v);updatePac(PAC_PARENT[child]);}
function setDartaPlan(plan){if(!CD.portfolios.darta)return;recordUndo();CD.portfolios.darta.plan=plan;buildTab('darta');}
function getDartaPrivateCG(total){if(total>=20000000)return 0.80;if(total>=10000000)return 0.90;if(total>=5000000)return 1.10;if(total>=3500000)return 1.20;if(total>=1500000)return 1.40;return 1.60;}
function getDartaPrivateClass(total){if(total>=20000000)return 'E';if(total>=10000000)return 'D';if(total>=5000000)return 'C';if(total>=3500000)return 'B';if(total>=1500000)return 'A';return 'AA';}
function getDartaCopertura(total){if(total<=500000)return 0.0025;return 0.0025*500000/total;}
function updatePac(parent='darta'){const ppd=CD?.portfolios?.[parent];if(ppd?.pac?.sourceKey==='__SG__')ppd.pac.totalSwitch=easySwitchAmount();buildTab(PAC_CHILD[parent]);const e=ALL_PTFL.filter(t=>!isPacType(t)&&!!CD.portfolios[t]);renderTotals(e);}

// ── EASY SWITCH (DARTA) ───────────────────────────────────────────────────────
// New model: a real DARTA fund is chosen as the source; a total amount is
// converted into the Easy Switch plan over N months (rata = total ÷ mesi).
// The source fund keeps its full value (the money sits there at day one).
function ensurePac(){if(CD.portfolios.darta&&!CD.portfolios.darta.pac)CD.portfolios.darta.pac={};}
// Flat list of named DARTA fund rows, each with its current controvalore.
function dartaFundRows(){
  const pd=CD?.portfolios?.darta;if(!pd)return[];
  const w=pd.weights||{};const rows=pd.rows||{};const total=pd.total||0;const out=[];
  (pd._order||Object.keys(w)).forEach(ac=>{
    const acd=w[ac];if(!acd||ac==='START & GO')return;
    getSubOrder('darta',ac).forEach(sub=>{
      const sp100=acd.subs[sub];if(sp100===undefined)return;
      Object.keys(rows).filter(k=>k.startsWith(`darta|${ac}|${sub}|`)).forEach(k=>{
        const rd=rows[k];const nm=(rd&&rd.name)?rd.name:((rd&&rd.isin)?rd.isin:null);
        if(!nm)return;
        const sv=total*(acd.pct/100)*(sp100/100);const {ctrv}=derive(rd,sv);
        out.push({key:k,label:nm,ac,sub,ctrv:ctrv>0?Math.round(ctrv):0});
      });
    });
  });
  return out;
}
function easySwitchSourceLabel(){const cfg=CD?.portfolios?.darta?.pac;if(!cfg?.sourceKey)return null;if(cfg.sourceKey==='__SG__')return 'START & GO';const r=dartaFundRows().find(r=>r.key===cfg.sourceKey);return r?r.label:null;}
// True when a START & GO asset-class bucket exists in the DARTA portfolio.
function hasDartaSG(){return !!(CD?.portfolios?.darta?.weights?.['START & GO']);}
function setEasySwitchEnabled(on){recordUndo();ensurePac();const cfg=CD.portfolios.darta.pac;cfg.enabled=on;
  if(on){if(!CD.portfolios.pac){if(CD._pacBackup&&Object.keys(CD._pacBackup).length){CD.portfolios.pac=CD._pacBackup;delete CD._pacBackup;}else CD.portfolios.pac=makeDefault('pac');}}
  else{if(CD.portfolios.pac){CD._pacBackup=JSON.parse(JSON.stringify(CD.portfolios.pac));delete CD.portfolios.pac;}}
  const e=ALL_PTFL.filter(t=>!isPacType(t)&&!!CD.portfolios[t]);renderTotals(e);renderTabs(e);
}
function setEasySwitchSource(key){ensurePac();CD.portfolios.darta.pac.sourceKey=key||null;buildTab('darta');const e=ALL_PTFL.filter(t=>!isPacType(t)&&!!CD.portfolios[t]);renderTotals(e);}
function setEasySwitchTotal(v){ensurePac();CD.portfolios.darta.pac.totalSwitch=parseEuro(v);updatePac('darta');buildTab('darta');}
// One-time migration: convert legacy START & GO bucket → Easy Switch source model.
function migrateDartaEasySwitch(){
  const pd=CD?.portfolios?.darta;if(!pd)return;
  // Runs once per portfolio. The `_esMigrated` flag stops it from stripping a
  // START & GO bucket the user has deliberately re-added (new Easy Switch model).
  if(pd._esMigrated)return;
  const w=pd.weights||{};
  if(w['START & GO']){
    const sgVal=Math.round((pd.total||0)*((w['START & GO'].pct||0)/100));
    if(!pd.pac)pd.pac={};
    pd.pac.enabled=true;
    if(!(pd.pac.totalSwitch>0))pd.pac.totalSwitch=sgVal;
    if(pd.pac.mesi==null)pd.pac.mesi=15;
    delete w['START & GO'];
    if(pd._order)pd._order=pd._order.filter(a=>a!=='START & GO');
    if(!CD.portfolios.pac){if(CD._pacBackup&&Object.keys(CD._pacBackup).length){CD.portfolios.pac=CD._pacBackup;delete CD._pacBackup;}else CD.portfolios.pac=makeDefault('pac');}
  }
  pd._esMigrated=true;
}
// One-time migration: detach the old SICAV START & GO → PAC derivation into a
// standalone sicavPac (own total + months). Preserves the previous monthly rata:
// old rata = sicav.total × START&GO% ÷ mesi, so we set total = sicav.total × START&GO%.
function migrateSicavPacStandalone(){
  const s=CD?.portfolios?.sicav;const sw=s?.weights||{};
  const hadSG=!!sw['START & GO'];const oldMesi=s?.pac?.mesi||null;
  const sgVal=hadSG?Math.round((s.total||0)*((sw['START & GO'].pct||0)/100)):0;
  let sp=CD.portfolios.sicavPac;
  // Only migrate legacy data (old sicavPac had no own `mesi` field).
  if(sp&&sp.mesi===undefined){
    sp.mesi=oldMesi;
    if(!(sp.total>0))sp.total=sgVal;
  } else if(!sp&&hadSG){
    sp=makeDefault('sicavPac');sp.mesi=oldMesi;sp.total=sgVal;CD.portfolios.sicavPac=sp;
  }
  // Strip the now-obsolete START & GO bucket and pac config from SICAV.
  if(hadSG){delete sw['START & GO'];if(s._order)s._order=s._order.filter(a=>a!=='START & GO');}
  if(s&&s.pac)delete s.pac;
}
function easySwitchPanelHTML(){
  const cfg=CD?.portfolios?.darta?.pac||{};const on=!!cfg.enabled;
  if(!on){
    return `<div class="es-panel"><div class="es-head"><label class="es-toggle"><input type="checkbox" onchange="setEasySwitchEnabled(this.checked)"/> <span>Easy Switch</span></label><span class="es-sub">Converti parte di un fondo esistente in un piano Easy Switch</span></div></div>`;
  }
  const rows=dartaFundRows();
  const isSG=cfg.sourceKey==='__SG__';
  const dpd=CD?.portfolios?.darta||{};
  const sgCtrv=Math.round((dpd.total||0)*(((dpd.weights||{})['START & GO']?.pct||0)/100));
  const sgOpt=hasDartaSG()?`<option value="__SG__" ${isSG?'selected':''}>START & GO${sgCtrv>0?' — '+fmt(sgCtrv):''}</option>`:'';
  const opts=`<option value="">— scegli un fondo —</option>`+sgOpt+rows.map(r=>`<option value="${r.key}" ${cfg.sourceKey===r.key?'selected':''}>${r.label}${r.ctrv>0?' — '+fmt(r.ctrv):''}</option>`).join('');
  const ms=cfg.mesi||null;const tot=isSG?sgCtrv:(cfg.totalSwitch||0);const rata=(ms&&tot)?tot/ms:0;
  const src=rows.find(r=>r.key===cfg.sourceKey);
  const warn=(!isSG&&src&&src.ctrv>0&&tot>src.ctrv)?`<div class="es-warn">⚠ L'importo supera il controvalore del fondo di origine (${fmt(src.ctrv)})</div>`:'';
  const noFunds=(isSG||rows.length)?'':`<div class="es-warn">Aggiungi prima dei fondi nel portafoglio DARTA, poi scegline uno come origine.</div>`;
  // START & GO is the residual bucket → amount is its live controvalore (read-only);
  // a real fund source keeps the typed amount.
  const amountField=isSG
    ?`<div class="es-field"><label>Importo totale da convertire</label><div class="es-inp-row"><span class="tc-pre">€</span><input id="es-amount-inp" type="text" value="${fmtEuroInp(tot)}" disabled title="Definito dalla quota residua START & GO"/></div><span class="es-sub" style="margin-top:4px;display:block">= quota residua START &amp; GO del totale DARTA</span></div>`
    :`<div class="es-field"><label>Importo totale da convertire</label><div class="es-inp-row"><span class="tc-pre">€</span><input id="es-amount-inp" type="text" inputmode="numeric" value="${fmtEuroInp(tot)}" onfocus="onEuroFocus(this)" onchange="setEasySwitchTotal(this.value);onEuroBlur(this)"/></div></div>`;
  return `<div class="es-panel on">
    <div class="es-head"><label class="es-toggle"><input type="checkbox" checked onchange="setEasySwitchEnabled(this.checked)"/> <span>Easy Switch</span></label><span class="es-sub">Un fondo DARTA alimenta mensilmente il piano Easy Switch</span></div>
    <div class="es-body">
      <div class="es-field"><label>Fondo di origine</label><select onchange="setEasySwitchSource(this.value)">${opts}</select></div>
      ${amountField}
      <div class="es-field"><label>Durata (mesi)</label><input type="number" min="1" step="1" value="${ms||''}" onchange="setPacMesi('darta',this.value)"/></div>
      <div class="es-calc"><span class="es-calc-lbl">Rata mensile</span><strong class="es-calc-val">${fmt(rata)}/mese</strong><span class="es-calc-arr">→ alimenta il tab <strong>Easy Switch</strong></span></div>
      ${noFunds}${warn}
    </div>
  </div>`;
}

// ── TABS ──────────────────────────────────────────────────────────────────────
function renderTabs(enabled){
  const bar=document.getElementById('tab-bar');if(!bar)return;
  const show=ALL_PTFL.filter(t=>{
    if(enabled.includes(t))return true;
    if(isPacType(t)){const par=PAC_PARENT[t];if(!enabled.includes(par))return false;return t==='pac'?!!CD.portfolios?.darta?.pac?.enabled:true;}
    return false;});
  bar.innerHTML=show.map((t,i)=>`<button class="tab-btn${i===0?' active':''}" id="tb-${t}" onclick="switchTab('${t}')">${TL[t]}</button>`).join('');
  ALL_PTFL.forEach(t=>{const el=document.getElementById('tp-'+t);if(el)el.classList.remove('active');});
  if(show.length){document.getElementById('tp-'+show[0])?.classList.add('active');show.forEach(t=>buildTab(t));}
}
function switchTab(t){document.querySelectorAll('.tab-btn').forEach(b=>b.classList.remove('active'));document.querySelectorAll('.tab-pane').forEach(p=>p.classList.remove('active'));document.getElementById('tb-'+t)?.classList.add('active');document.getElementById('tp-'+t)?.classList.add('active');}

// ── FX RATES (EUR base · live, no API key) ────────────────────────────────────
let _fxRates={EUR:1},_fxDate=null;
async function loadFx(){
  try{const r=await fetch('/api/fx');const d=await r.json();if(d&&d.rates&&Object.keys(d.rates).length){_fxRates=d.rates;_fxDate=d.date||null;}}catch(e){}
}
// Convert an amount entered in EUR into an instrument's quote currency.
function eurTo(amount,cur){const rate=_fxRates[(cur||'EUR').toUpperCase()];return rate?amount*rate:amount;}

// ── STOCK UNIVERSE CACHE ──────────────────────────────────────────────────────
let _stockUniverse=[];
async function loadStockUniverse(){
  try{const r=await fetch('/api/stock-universe');_stockUniverse=await r.json();}catch(e){_stockUniverse=[];}
}
function stockByCompany(company){return _stockUniverse.find(s=>s.company===company)||null;}
function stockSearch(q){
  const u=q.toUpperCase();
  return _stockUniverse.filter(s=>(s.company||'').toUpperCase().includes(u)||(s.ticker||'').toUpperCase().includes(u));
}

// ── ETF UNIVERSE CACHE ────────────────────────────────────────────────────────
let _etfUniverse=[];
async function loadEtfUniverse(){
  try{const r=await fetch('/api/etf-universe');_etfUniverse=await r.json();}catch(e){_etfUniverse=[];}
}
function etfByTicker(ticker){return _etfUniverse.find(e=>e.ticker===ticker)||null;}
function etfSearch(q){
  const u=q.toUpperCase();
  return _etfUniverse.filter(e=>(e.ticker||'').toUpperCase().includes(u)||(e.etf_name||'').toUpperCase().includes(u)||(e.sector||'').toUpperCase().includes(u));
}

