// ── STATE ─────────────────────────────────────────────────────────────────────
let clients=[], AID=null, CD=null;
let consultants=[], CONS=null; // CONS = id of the consultant whose book is open
let delId=null, dupId=null, addAcPtfl=null, addSubPtfl=null, addSubAc=null;
let _dSub=null, _dPtfl=null, _dAc=null;
let isinCache={}, dtimers={}, autoTimer=null;
let cgData={pro:{isin:{},name:{}},chpv:{isin:{},name:{}}};
let dartaCats={equity:[],bond:[],bilanciato:[],flessibili:[],commodity:[],altri:[]};
let catTree={};
let _catTreeSelected=null;
let dartaFundsByCat={pro:{},chpv:{}};
let _fdbData={pro:{headers:[],rows:[]},chpv:{headers:[],rows:[]},sicav:{headers:[],rows:[]}};
let _fdbCurrentTab='pro';
let _fdbNextId=0;
let sicavFundsByCat={};
let dashCharts={}, sumCharts={};
// ── undo / autosave state ──
let _undo=[], _redo=[], _undoBaseline=null;
let _lastSavedJSON=null, _lastSavedAt=null, _saving=false;

const CLSC={EQUITY:'eq',BOND:'bd',COMMODITY:'cm','START & GO':'sg'};
const RESIDUAL={amm:'ETC',darta:'START & GO',pac:'BOND',sicav:'COMMODITY',sicavPac:'BOND'};
function getResidual(ptfl){if(ptfl==='darta')return getW(ptfl)?.['START & GO']?'START & GO':null;const r=RESIDUAL[ptfl];return r&&getW(ptfl)?.[r]?r:null;}
const TL={amm:'Amministrato ETF',ammAz:'Amministrato Azioni',ammStock:'Amministrato Azioni',ammEtf:'Amministrato ETF',darta:'DARTA',pac:'Easy Switch',sicav:'SICAV & Fondi',sicavPac:'PAC'};
function acLabel(ac,ptfl){return ac;}
const ALL_PTFL=['amm','ammAz','ammStock','ammEtf','darta','pac','sicav','sicavPac'];
// Easy Switch (pac) is still a child derived from DARTA. PAC (sicavPac) is now a
// standalone top-level type: the user enters its total + months directly.
const PAC_PARENT={pac:'darta'};
const PAC_CHILD={darta:'pac'};
const isPacType=p=>!!PAC_PARENT[p];
// "pac-like" = monthly-accumulation plan (Easy Switch + PAC): shows Costo/m + Costo Tot.
const isPacLike=p=>p==='pac'||p==='sicavPac';
// Months for a pac-like plan: Easy Switch reads DARTA's config; PAC stores its own.
function pacMesi(p){return (p==='sicavPac'?CD?.portfolios?.sicavPac?.mesi:CD?.portfolios?.darta?.pac?.mesi)||0;}
const AC_COLORS=['#2eaa5f','#2e72d2','#d97f00','#c03060','#7c5cbf','#1db8b0','#e05038','#a09018','#3090d0','#b05820','#208050','#8838b8'];
const AC_FIXED={EQUITY:'#2eaa5f',BOND:'#2e72d2',COMMODITY:'#d97f00','START & GO':'#c03060'};
const CUSTOM_AC_COLORS=['#c9a84c','#1db8b0','#e05038','#a09018','#3090d0','#b05820','#208050','#8838b8','#d06040','#6050d0','#30a870','#c08830'];
let _customAcIdx=0;const _customAcMap={};
function acColor(ac,ci){if(AC_FIXED[ac])return AC_FIXED[ac];if(!_customAcMap[ac]){_customAcMap[ac]=CUSTOM_AC_COLORS[_customAcIdx%CUSTOM_AC_COLORS.length];_customAcIdx++;}return _customAcMap[ac];}
const SUB_COLORS=['#3dc96e','#4d8ce8','#f0920c','#d94080','#9060d0','#1bc8b8','#e05838','#a0b018','#5090e0','#d068b8','#58c040','#c09028','#4050d8','#e84870','#78c888','#b040a0'];

// ── HELPERS ───────────────────────────────────────────────────────────────────
const fmt=n=>(!n||isNaN(n))?'—':'€'+Math.round(n).toLocaleString('it-IT',{useGrouping:'always'});
const _esc=s=>(s||'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const fmtC=n=>(!n||isNaN(n))?'—':'€'+n.toLocaleString('it-IT',{minimumFractionDigits:2,maximumFractionDigits:2,useGrouping:'always'});
const fmtP=n=>(n==null||isNaN(n))?'—':(n*100).toFixed(2)+'%';
const fmtQ=n=>(!n||isNaN(n))?'—':parseFloat(n.toFixed(3)).toLocaleString('it-IT');
// ── Icone SVG a tratto (stile line, colore = currentColor) ──
const ICONS={
  menu:'<path d="M4 6h16M4 12h16M4 18h16"/>',
  plus:'<path d="M12 5v14M5 12h14"/>',
  database:'<ellipse cx="12" cy="5" rx="8" ry="3"/><path d="M4 5v14c0 1.7 3.6 3 8 3s8-1.3 8-3V5"/><path d="M4 12c0 1.7 3.6 3 8 3s8-1.3 8-3"/>',
  trend:'<path d="M3 17l6-6 4 4 8-8"/><path d="M14 7h7v7"/>',
  bars:'<path d="M3 21h18"/><path d="M6 17v-5M11 17V6M16 17v-8M21 17v-3"/>',
  upload:'<path d="M12 15V3"/><path d="M7 8l5-5 5 5"/><path d="M4 15v4a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-4"/>',
  download:'<path d="M12 3v12"/><path d="M7 10l5 5 5-5"/><path d="M4 17v2a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-2"/>',
  home:'<path d="M3 10.5 12 3l9 7.5"/><path d="M5 9.5V20a1 1 0 0 0 1 1h4v-6h4v6h4a1 1 0 0 0 1-1V9.5"/>',
  edit:'<path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4Z"/>',
  copy:'<rect x="8" y="8" width="13" height="13" rx="2"/><path d="M16 8V5a2 2 0 0 0-2-2H5a2 2 0 0 0-2 2v9a2 2 0 0 0 2 2h3"/>',
  save:'<path d="M5 3h11l5 5v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2Z"/><path d="M7 3v5h8"/><path d="M7 21v-7h10v7"/>',
  trash:'<path d="M3 6h18"/><path d="M8 6V4h8v2"/><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/>',
  refresh:'<path d="M21 12a9 9 0 1 1-2.6-6.4L21 8"/><path d="M21 3v5h-5"/>',
  file:'<path d="M14 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9Z"/><path d="M14 3v6h6"/>',
  slides:'<rect x="3" y="4" width="18" height="12" rx="1.5"/><path d="M12 16v4M8 21h8"/>',
  sheet:'<rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M3 15h18M9 3v18"/>',
  clipboard:'<rect x="5" y="4" width="14" height="18" rx="2"/><path d="M9 4V2.5h6V4"/><path d="M9 11h6M9 15h4"/>',
  bank:'<path d="M3 10 12 4l9 6"/><path d="M5 10v8M9.5 10v8M14.5 10v8M19 10v8"/><path d="M3 21h18"/>',
  coins:'<circle cx="9" cy="9" r="6"/><path d="M15.4 9.6a6 6 0 1 1-5.8 5.8"/>',
  folder:'<path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v9a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2Z"/>',
  pie:'<path d="M21 12A9 9 0 1 1 12 3v9Z"/><path d="M15 3.3A9 9 0 0 1 20.7 9H15Z"/>',
  scale:'<path d="M12 3v18M7 21h10M5 7h14"/><path d="M5 7l-3 7a3 3 0 0 0 6 0Z"/><path d="M19 7l-3 7a3 3 0 0 0 6 0Z"/>',
  layers:'<path d="m12 3 9 5-9 5-9-5Z"/><path d="m3 13 9 5 9-5"/>',
  gem:'<path d="M6 3h12l4 6-10 12L2 9Z"/><path d="M2 9h20M12 21 8 9l4-6 4 6"/>',
};
function hydrateIcons(root){(root||document).querySelectorAll('i[data-ic]').forEach(i=>{i.outerHTML=ic(i.dataset.ic);});}
function ic(name,cls){return `<svg class="ic${cls?' '+cls:''}" viewBox="0 0 24 24" aria-hidden="true">${ICONS[name]||''}</svg>`;}
// ── Conteggio animato dei totali (0 → valore) per gli elementi con data-cu ──
const _reduceMotion=()=>window.matchMedia&&matchMedia('(prefers-reduced-motion: reduce)').matches;
function countUp(el,target,f=fmt,ms=900){
  if(!el)return;
  if(!target||isNaN(target)||_reduceMotion()){el.textContent=f(target);return;}
  const t0=performance.now();el._cuTok=t0;
  const step=now=>{
    if(el._cuTok!==t0)return; // un nuovo render ha preso il controllo
    const p=Math.min(1,(now-t0)/ms),e=1-Math.pow(1-p,4);
    const v=target*e;
    el.textContent=(p<1&&Math.abs(v)<1)?f(target).replace(/[\d.,]+/,'0'):f(p<1?v:target);
    if(p<1)requestAnimationFrame(step);
  };
  requestAnimationFrame(step);
}
function countUps(root){(root||document).querySelectorAll('[data-cu]').forEach(el=>{const v=parseFloat(el.dataset.cu);el.removeAttribute('data-cu');countUp(el,v);});}
// Animazione donut: crescita + rotazione al primo disegno, rapida nei drill-down
const donutAnim=level=>level==='top'?{animateRotate:true,animateScale:true,duration:1000,easing:'easeOutQuart'}:{animateRotate:true,duration:450,easing:'easeOutCubic'};
// ── Euro input formatting: raw digits while editing, raggruppato (1.250.000) a riposo ──
function unfmtEuro(s){return String(s==null?'':s).replace(/\./g,'').replace(/,/g,'.').replace(/[^\d.\-]/g,'');}
function parseEuro(s){const n=parseFloat(unfmtEuro(s));return isNaN(n)?0:n;}
function fmtEuroInp(n){const v=parseFloat(n);return(!v||isNaN(v))?'':Math.round(v).toLocaleString('it-IT',{useGrouping:'always'});}
function onEuroFocus(el){el.value=unfmtEuro(el.value);el.select&&el.select();}
function onEuroBlur(el){el.value=fmtEuroInp(parseEuro(el.value));}
// ── Etichette leggibili dei portafogli (per avvisi export) ──
const PTFL_LBL={amm:'Amministrato ETF',ammAz:'Amministrato Azioni',ammStock:'Amministrato Azioni',ammEtf:'Amministrato ETF',darta:'DARTA + Easy Switch',pac:'Easy Switch (PAC)',sicav:'SICAV & Fondi',sicavPac:'PAC'};
// ── Controllo allocazioni: ritorna gli scostamenti da 100% (asset class e sotto-settori) ──
function getAllocWarnings(){
  const out=[];if(!CD)return out;
  for(const ptfl of Object.keys(CD.portfolios)){
    const pd=CD.portfolios[ptfl];const w=pd&&pd.weights;
    if(!w||typeof w!=='object'||!Object.keys(w).length)continue;
    const label=PTFL_LBL[ptfl]||ptfl;
    const acSum=Math.round(Object.values(w).reduce((s,d)=>s+(parseFloat(d.pct)||0),0)*100)/100;
    if(Math.abs(acSum-100)>=0.06)out.push({label,scope:'Asset class',sum:acSum});
    for(const[ac,d]of Object.entries(w)){
      const subs=d.subs||{};const keys=Object.keys(subs);if(!keys.length)continue;
      const ss=Math.round(keys.reduce((s,k)=>s+(parseFloat(subs[k])||0),0)*100)/100;
      if(Math.abs(ss-100)>=0.06)out.push({label,scope:acLabel(ac,ptfl),sum:ss});
    }
  }
  return out;
}
// ── UNDO / REDO ──
function _cloneP(){return JSON.parse(JSON.stringify(CD.portfolios));}
function _pushUndo(state){_undo.push(state);if(_undo.length>80)_undo.shift();_redo=[];updateUndoBtns();}
function _flushUndo(){if(_undoBaseline!=null){if(_undoBaseline!==JSON.stringify(CD.portfolios))_pushUndo(JSON.parse(_undoBaseline));_undoBaseline=null;}}
function recordUndo(){if(!CD)return;_flushUndo();_pushUndo(_cloneP());}
function _applyHistory(snapshot,msg){CD.portfolios=snapshot;const enabled=ALL_PTFL.filter(t=>!isPacType(t)&&!!CD.portfolios[t]);renderPtflCards(enabled);renderTotals(enabled);renderTabs(enabled);updateUndoBtns();updateSaveInfo();toast(msg,'info');}
function doUndo(){if(!CD)return;_flushUndo();if(!_undo.length){toast('Niente da annullare','info');return;}_redo.push(_cloneP());_applyHistory(_undo.pop(),'Modifica annullata');}
function doRedo(){if(!CD)return;_flushUndo();if(!_redo.length){toast('Niente da ripristinare','info');return;}_undo.push(_cloneP());_applyHistory(_redo.pop(),'Modifica ripristinata');}
function updateUndoBtns(){const u=document.getElementById('undo-btn');if(u)u.disabled=!_undo.length;const r=document.getElementById('redo-btn');if(r)r.disabled=!_redo.length;}
// ── Indicatore di salvataggio nella save-bar ──
function isDirty(){return CD&&_lastSavedJSON!==JSON.stringify(CD.portfolios);}
function updateSaveInfo(){const el=document.querySelector('.save-info');if(!el||!CD)return;let status='';if(isDirty())status=` · <span style="color:#e8c97a;font-weight:700">● Modifiche non salvate</span>`;else if(_lastSavedAt)status=` · <span style="color:#7fd99f;font-weight:700">✓ Salvato alle ${_lastSavedAt.toLocaleTimeString('it-IT',{hour:'2-digit',minute:'2-digit'})}</span>`;el.innerHTML=`Cliente: <strong>${_esc(CD.name)}</strong>${status}`;}
// ── Tastiera nella tabella fondi: Invio avanza al campo/riga successiva, Ctrl+Invio aggiunge una riga ──
const FUND_FIELD_SEL='.isin-inp,.name-inp-manual,.pset-inp,.price-inp,.ctrv-inp,.cg-inp,.sc-inp';
function fundRowKeydown(e){
  if(e.key!=='Enter')return;
  const t=e.target;if(!t||!t.matches('input'))return;
  const row=t.closest('.fund-row');if(!row)return;
  const sub=row.parentElement;
  const addBtn=sub&&sub.querySelector('.add-row-btn');
  // Ctrl/Cmd+Invio → aggiungi sempre una riga in questo sotto-settore
  if(e.metaKey||e.ctrlKey){if(addBtn){e.preventDefault();addBtn.click();}return;}
  // se è aperto il menu suggerimenti (ISIN/nome) lascia il comportamento esistente
  if(row.querySelector('.fund-dd'))return;
  // % Port.: Invio salva e passa al % Port. del fondo successivo (anche nel sotto-settore dopo)
  if(t.classList.contains('pset-inp')){
    e.preventDefault();
    const pane=t.closest('.tab-pane')||document;
    const ids=[...pane.querySelectorAll('.pset-inp')].filter(i=>!i.disabled).map(i=>i.id);
    const next=ids[ids.indexOf(t.id)+1];
    t.blur(); // fa partire onchange (che può ridisegnare la tabella)
    setTimeout(()=>{const n=next&&document.getElementById(next);if(n){n.focus();n.select&&n.select();}},60);
    return;
  }
  const inputs=[...row.querySelectorAll(FUND_FIELD_SEL)].filter(i=>!i.disabled&&i.offsetParent!==null);
  const i=inputs.indexOf(t);
  e.preventDefault();
  if(i>-1&&i<inputs.length-1){const n=inputs[i+1];n.focus();n.select&&n.select();return;}
  // ultimo campo della riga → prossima riga, oppure crea una nuova riga
  let nx=row.nextElementSibling;while(nx&&!nx.classList.contains('fund-row')&&!nx.classList.contains('add-row-btn'))nx=nx.nextElementSibling;
  if(nx&&nx.classList.contains('fund-row')){nx.querySelector('.name-inp-manual,.isin-inp')?.focus();}
  else if(addBtn){addBtn.click();}
}
// ── Scorciatoie globali undo/redo nella vista Modifica ──
document.addEventListener('keydown',e=>{
  if(!(e.metaKey||e.ctrlKey)||e.key.toLowerCase()!=='z')return;
  if(!document.getElementById('view-edit')?.classList.contains('active'))return;
  e.preventDefault();e.shiftKey?doRedo():doUndo();
});
// Short date/time for the purchase baseline: "dd/MM HH:mm" for timestamps, "dd/MM/yy" for date-only.
function fmtShortDt(s){
  if(!s)return '';
  if(/^\d{4}-\d{2}-\d{2}$/.test(s)){const[y,m,d]=s.split('-');return `${d}/${m}/${y.slice(2)}`;}
  const dt=new Date(s);if(isNaN(dt))return s;
  const p=n=>String(n).padStart(2,'0');
  return `${p(dt.getDate())}/${p(dt.getMonth()+1)} ${p(dt.getHours())}:${p(dt.getMinutes())}`;
}
const rk=(p,a,s,t,i)=>`${p}|${a}|${s}|${t}|${i}`;
const sid=k=>'r'+k.replace(/[|]/g,'_');
const tcls=t=>!t?'az':/etf/i.test(t)?'etf':/fondo|mutual/i.test(t)?'fondo':/obbl|bond/i.test(t)?'ob':'az';

function getPD(ptfl){return ptfl==='pac'?CD?.portfolios?.pac:CD?.portfolios?.[ptfl];}
function getW(ptfl){return getPD(ptfl)?.weights||{};}
function getRows(ptfl){return getPD(ptfl)?.rows||{};}
function getRow(k,ptfl){return getRows(ptfl)[k]||{isin:'',name:'',price:null,type:'',pctSettore:0,cg:'',sc:''};}
function setRow(k,ptfl,d){const pd=getPD(ptfl);if(!pd)return;if(!pd.rows)pd.rows={};pd.rows[k]={...(pd.rows[k]||{isin:'',name:'',price:null,type:'',pctSettore:0,cg:'',sc:''}),...d};}

const CHPV_CG_EXCEPTIONS={
  'PIMCO OBBLIGAZIONARIO PRUDENTE':0.62,
  'ALLIANZGI TESORERIA':0.35,
};
function parsePctStr(s){if(!s&&s!==0)return 0;const v=parseFloat(String(s).replace(',','.').replace('%','').trim());return isNaN(v)?0:v/100;}
function getAutoCommission(k,ptfl){
  const rd=getRow(k,ptfl);const plan=(ptfl==='pac'?getPD('darta'):getPD(ptfl))?.plan||'pro';
  const iK=(rd.isin||'').trim().toUpperCase(),nK=(rd.name||'').trim().toUpperCase();
  let cg='',sc='';
  if(plan==='pro'){
    const m=(iK&&cgData.pro.isin[iK])||(nK&&cgData.pro.name[nK])||{};
    cg=m.cg||'';sc=m.sc||'';
  }else{
    const ms=(iK&&cgData.chpv.isin[iK])||(nK&&cgData.chpv.name[nK])||{};
    sc=ms.sc||'';
    const exc=nK?CHPV_CG_EXCEPTIONS[nK]:undefined;
    if(exc!=null){cg=exc.toFixed(2)+'%';}
    else{cg=getDartaPrivateCG(getTotal(ptfl)).toFixed(2)+'%';}
  }
  return{cg,sc};
}
// Easy Switch (pac): monthly rata = source total ÷ months, derived from DARTA.
// When the source is the START & GO residual bucket, the amount is its live
// controvalore (DARTA total × residual %); otherwise it's the typed totalSwitch.
function easySwitchAmount(){const ppd=CD?.portfolios?.darta;const cfg=ppd?.pac;if(!cfg)return 0;
  if(cfg.sourceKey==='__SG__'){const sg=(ppd.weights||{})['START & GO'];return sg?Math.round((ppd.total||0)*((sg.pct||0)/100)):0;}
  return cfg.totalSwitch||0;}
function pacActual(p='pac'){const ppd=CD?.portfolios?.darta;const ms=ppd?.pac?.mesi||15;const amt=easySwitchAmount();return ms>0?amt/ms:0;}
function pacTotal(p='pac'){const actual=pacActual(p);const pac=CD?.portfolios?.darta?.pac;if(pac?.useRounded&&pac?.roundedTotal>0)return pac.roundedTotal;return actual;}
// PAC (sicavPac): monthly rata = user-entered total ÷ user-entered months.
function sicavPacRata(){const pd=CD?.portfolios?.sicavPac;const ms=pd?.mesi||0;return ms>0?(pd.total||0)/ms:0;}
function getTotal(ptfl){if(ptfl==='pac')return pacTotal('pac');if(ptfl==='sicavPac')return sicavPacRata();return CD?.portfolios?.[ptfl]?.total||0;}
function derive(rd,sv){const c=rd.ctrvAbs!=null?rd.ctrvAbs:sv*(rd.pctSettore||0);return{ctrv:c,qty:(rd.price&&rd.price>0)?c/rd.price:null};}
// Sum of the actual controvalore of every individual asset in a portfolio.
function getCtrvSum(ptfl){
  if(ptfl==='ammStock'||ptfl==='ammEtf'){return (getPD(ptfl)?.rows||[]).reduce((s,r)=>s+(r.ctrv||0),0);}
  const pd=getPD(ptfl);if(!pd)return 0;
  const total=getTotal(ptfl);
  const weights=pd.weights||{};const rows=getRows(ptfl);let sum=0;
  Object.keys(rows).forEach(k=>{
    const parts=k.split('|');if(parts[0]!==ptfl)return;
    const ac=parts[1],sub=parts[2];const acd=weights[ac];if(!acd)return;
    const sp100=(acd.subs||{})[sub]||0;const sv=total*(acd.pct/100)*(sp100/100);
    const {ctrv}=derive(getRow(k,ptfl),sv);sum+=ctrv>0?Math.round(ctrv):0;
  });
  return sum;
}
// Progress bar: yellow (far) → green (matched) → red (overshoot) for ctrv-sum vs total.
function ctrvBar(ptfl,total){
  if(!(total>0))return '';
  const s=getCtrvSum(ptfl);if(!(s>0))return '';
  const ratio=s/total,over=ratio>1.0001;
  let color,fillW,label;
  if(over){color='#e0344b';fillW=100;label=`${fmt(s)} · +${Math.round((ratio-1)*100)}% oltre il totale`;}
  else{const hue=60+60*ratio;color=`hsl(${hue},72%,46%)`;fillW=ratio*100;label=`${fmt(s)} · ${Math.round(ratio*100)}% del totale`;}
  return `<div class="tc-bar" title="Somma controvalori asset: ${fmt(s)} / ${fmt(total)}"><div class="tc-bar-fill" style="width:${fillW}%;background:${color}"></div></div><div class="tc-bar-lbl"${over?' style="color:#e0344b"':''}>${label}</div>`;
}
function refreshCtrvBar(ptfl){const el=document.getElementById('tcbar-'+ptfl);if(el)el.innerHTML=ctrvBar(ptfl,getTotal(ptfl));}
function getWOrder(ptfl){const pd=getPD(ptfl);const w=getW(ptfl);if(pd?._order)return pd._order.filter(k=>w[k]);return Object.keys(w);}
function setWOrder(ptfl,o){const pd=getPD(ptfl);if(pd)pd._order=o;}
function getSubOrder(ptfl,ac){const acd=getW(ptfl)[ac];if(!acd)return[];if(acd._subOrder)return acd._subOrder.filter(k=>Object.prototype.hasOwnProperty.call(acd.subs,k));return Object.keys(acd.subs||{});}
function setSubOrder(ptfl,ac,o){const acd=getW(ptfl)[ac];if(acd)acd._subOrder=o;}

// ── TOAST ─────────────────────────────────────────────────────────────────────
function toast(msg,type='info'){const el=document.getElementById('toast');el.textContent=msg;el.className=`toast ${type} show`;clearTimeout(el._t);el._t=setTimeout(()=>el.classList.remove('show'),3000);}

// ── MODALS ────────────────────────────────────────────────────────────────────
function openM(id){document.getElementById('m-'+id).classList.add('open');const inp=document.getElementById('inp-'+id);if(inp){inp.value='';setTimeout(()=>inp.focus(),80);}}
function closeM(id){document.getElementById('m-'+id).classList.remove('open');}

// ── Zoom di pagina (scala fissa sui computer): Chart.js calcola la posizione del mouse
// senza considerare il CSS zoom; la ricalcoliamo dalle misure reali a schermo del canvas.
// Rettangolo a schermo (px visivi, come clientX) di un elemento, qualunque sia il modo
// in cui il browser applica il CSS zoom a getBoundingClientRect.
function visRect(el){
  const z=window.PAGE_ZOOM||1,r=el.getBoundingClientRect();if(z===1)return r;
  const lw=el.offsetWidth||0; // larghezza in px CSS della pagina (non scalata)
  if(!lw||Math.abs(r.width-lw*z)<=Math.abs(r.width-lw))return r; // già visivo (Chrome/Firefox)
  return{left:r.left*z,top:r.top*z,right:r.right*z,bottom:r.bottom*z,width:r.width*z,height:r.height*z};
}
if(window.Chart&&Chart.register)Chart.register({id:'pageZoomFix',beforeEvent(chart,args){
  const z=window.PAGE_ZOOM||1;if(z===1)return;
  const e=args.event,n=e&&e.native;if(!n||n.clientX==null)return;
  const r=visRect(chart.canvas);if(!r.width||!r.height)return;
  e.x=(n.clientX-r.left)/r.width*chart.width;e.y=(n.clientY-r.top)/r.height*chart.height;
}});
// Chart.js misura il contenitore in px a schermo (già scalati dallo zoom) e disegnerebbe
// grafici troppo piccoli/grandi: si riporta la misura in px CSS della pagina.
if(window.Chart&&Chart.DomPlatform){
  const _gms=Chart.DomPlatform.prototype.getMaximumSize;
  Chart.DomPlatform.prototype.getMaximumSize=function(canvas,...a){
    const s=_gms.call(this,canvas,...a),z=window.PAGE_ZOOM||1,par=canvas&&canvas.parentNode;
    if(z===1||!par||!s||!s.width)return s;
    const lw=par.clientWidth;
    if(lw&&Math.abs(s.width-lw*z)<Math.abs(s.width-lw)){s.width=Math.round(s.width/z);s.height=Math.round(s.height/z);}
    return s;
  };
}
