// Build a standard {total,weights,_order} donut structure from an Amministrato
// Azioni/ETF row table so it drills like every other portfolio:
//  · Azioni → one EQUITY class, split by stock sector (industry).
//  · ETF    → EQUITY/BOND/COMMODITY classes (from ETF type; ETC = commodities),
//             each split by ETF sector.
function ammDataFromRows(ptfl,pd){
  const rows=(pd&&pd.rows)||[];
  const total=rows.reduce((s,r)=>s+(r.ctrv||0),0);
  if(total<=0)return {total:0,weights:{},_order:[]};
  const groups={}; // acKey -> {sectors:{name:ctrv}, total}
  rows.forEach(r=>{
    const ctrv=r.ctrv||0;if(ctrv<=0)return;
    let acKey,sector;
    if(ptfl==='ammStock'){
      acKey='EQUITY';
      const st=r.company?stockByCompany(r.company):null;
      sector=(st&&st.industry)||'Altro';
    }else{
      const etf=r.ticker?etfByTicker(r.ticker):null;
      const t=((etf&&etf.type)||'').trim().toLowerCase();
      acKey=t==='bond'?'BOND':t==='etc'?'COMMODITY':'EQUITY';
      sector=(etf&&etf.sector)||'Altro';
    }
    if(!groups[acKey])groups[acKey]={sectors:{},total:0};
    groups[acKey].sectors[sector]=(groups[acKey].sectors[sector]||0)+ctrv;
    groups[acKey].total+=ctrv;
  });
  const _order=['EQUITY','BOND','COMMODITY'].filter(k=>groups[k]);
  const weights={};
  _order.forEach(ac=>{
    const g=groups[ac];const subs={};
    Object.entries(g.sectors).sort((a,b)=>b[1]-a[1]).forEach(([sec,cv])=>{
      subs[sec]=Math.round(cv/g.total*1000)/10; // % within this asset class
    });
    weights[ac]={pct:Math.round(g.total/total*1000)/10,subs,_type:ac.toLowerCase()};
  });
  return {total,weights,_order};
}
// Resolve the donut data for a saved portfolio entry, expanding Amministrato rows.
function dashPd(pt,pi){
  if(!pi||!pi.data)return null;
  return (pt==='ammStock'||pt==='ammEtf')?ammDataFromRows(pt,pi.data):pi.data;
}

// ── AMMINISTRATO AZIONI (flat stock table) ────────────────────────────────────
function buildAmmStockTab(container,pd){
  const rows=pd.rows||[];

  const tb=document.createElement('div');
  tb.style.cssText='display:flex;align-items:center;gap:10px;margin-bottom:14px;';
  tb.innerHTML=`<button onclick="syncStockUniverse()" style="padding:5px 14px;background:var(--navy);color:var(--gold-l);border:1.5px solid rgba(201,168,76,.35);border-radius:6px;font-size:11px;font-weight:700;cursor:pointer;">${ic('refresh')} Sincronizza da Excel</button><span id="ammstock-sync-msg" style="font-size:10px;color:var(--txt3)"></span><span title="La quantità converte il CTRV in EUR nella valuta del titolo" style="margin-left:auto;font-size:10px;color:var(--txt3)">Cambio EUR/USD: <strong>${_fxRates.USD?_fxRates.USD.toFixed(4):'—'}</strong>${_fxDate?' · '+_fxDate:''}</span>`;
  container.appendChild(tb);

  const tbl=document.createElement('div');tbl.id='ammstock-table';tbl.className='tbl';

  const hdr=document.createElement('div');
  hdr.className='tbl-ch g8-stock';
  [['Società',''],['Settore',''],['Acquisto','col-r'],['Prezzo','col-r'],['CTRV (€)','col-r'],['Quantità','col-r'],['P/L (€)','col-r'],['% Port.','col-r'],['YTD %','col-r'],['','']].forEach(([h,cls])=>{
    const s=document.createElement('span');if(cls)s.className=cls;s.textContent=h;hdr.appendChild(s);
  });
  tbl.appendChild(hdr);

  const rowsEl=document.createElement('div');rowsEl.id='ammstock-rows';
  rows.forEach((r,i)=>rowsEl.appendChild(makeStockRow(r,i)));
  tbl.appendChild(rowsEl);

  const foot=document.createElement('div');foot.id='ammstock-foot';foot.className='grand-row g8-stock';
  [['TOTALE PORTAFOGLIO',''],['',''],['',''],['',''],['—','col-r grand-val'],['',''],['—','col-r grand-val'],['',''],['—','col-r'],['','']].forEach(([h,cls],i)=>{
    const s=document.createElement('span');if(cls)s.className=cls;
    if(i===4)s.id='ammstock-total-ctrv';
    if(i===6)s.id='ammstock-total-pl';
    if(i===8)s.id='ammstock-wytd';
    s.textContent=h;foot.appendChild(s);
  });
  tbl.appendChild(foot);
  container.appendChild(tbl);

  const addBtn=document.createElement('button');
  addBtn.className='add-row-btn';addBtn.style.marginTop='8px';
  addBtn.textContent='+ Aggiungi titolo';addBtn.onclick=()=>addStockRow();
  container.appendChild(addBtn);

  const summ=document.createElement('div');summ.className='amm-stats';
  summ.innerHTML=`
    <div class="amm-stat"><div class="amm-stat-lbl">Capitale investito</div><div class="amm-stat-val" id="ammstock-sum-put">—</div></div>
    <div class="amm-stat"><div class="amm-stat-lbl">Valore attuale</div><div class="amm-stat-val" id="ammstock-sum-now">—</div></div>
    <div class="amm-stat"><div class="amm-stat-lbl">Plus/Minusvalenze nette</div><div class="amm-stat-val" id="ammstock-sum-pl">—</div></div>
    <div class="amm-stat"><div class="amm-stat-lbl">Rendimento</div><div class="amm-stat-val" id="ammstock-sum-ret">—</div></div>`;
  container.appendChild(summ);

  recalcAmmStock();
}

function makeStockRow(r,idx){
  const row=document.createElement('div');
  row.className='fund-row g8-stock ammstock-row';
  row.dataset.idx=idx;

  const stock=r.company?stockByCompany(r.company):null;
  const ctrvBlocked=getTotal('ammStock')<=0;

  const tickerWrap=document.createElement('div');tickerWrap.style.cssText='position:relative;padding:2px 6px;display:flex;align-items:center;';
  const tickerInp=document.createElement('input');
  tickerInp.type='text';tickerInp.placeholder='Cerca titolo…';tickerInp.value=(r.company||r.ticker||'');
  tickerInp.className='flat-ticker flat-name';
  const dropdown=document.createElement('div');
  dropdown.className='fund-dd';dropdown.style.display='none';
  tickerInp.oninput=()=>{
    const q=tickerInp.value.trim();
    if(!q){dropdown.style.display='none';return;}
    const hits=stockSearch(q).slice(0,10);
    if(!hits.length){dropdown.style.display='none';return;}
    dropdown.innerHTML='';
    hits.forEach(s=>{
      const opt=document.createElement('div');opt.className='fund-dd-item';
      opt.innerHTML=`<span class="fund-dd-isin">${s.ticker||'—'}</span><span class="fund-dd-name">${s.company}</span>`;
      opt.onmousedown=()=>{tickerInp.value=s.company||s.ticker||'';dropdown.style.display='none';applyStockToRow(idx,s);};
      dropdown.appendChild(opt);
    });
    dropdown.style.display='block';
  };
  tickerInp.onblur=()=>setTimeout(()=>{dropdown.style.display='none';},150);
  tickerWrap.appendChild(tickerInp);tickerWrap.appendChild(dropdown);

  const mkText=(text,role)=>{const d=document.createElement('div');d.className='flat-cell r';if(role)d.dataset.role=role;d.textContent=text||'—';return d;};
  const mkLeft=(text)=>{const d=document.createElement('div');d.className='flat-cell';d.style.fontSize='11px';d.textContent=text||'—';return d;};
  const mkEditNum=(val,onchange,blocked)=>{
    const d=document.createElement('div');d.className='flat-cell r';
    const inp=document.createElement('input');inp.type='number';inp.min='0';inp.step='100';inp.value=val||'';inp.className='flat-inp';
    if(blocked){inp.disabled=true;inp.placeholder='Imposta totale';inp.title='Inserisci prima il totale';}
    inp.onchange=()=>{if(onchange)onchange(parseFloat(inp.value)||0);};
    d.appendChild(inp);return d;
  };

  const sectorSp =mkLeft(stock?.industry||'');
  const buySp    =document.createElement('div');buySp.className='flat-cell r buy-cell';buySp.dataset.role='buy';buySp.textContent='—';
  buySp.title='Clicca per impostare prezzo e data/ora di acquisto';buySp.onclick=()=>openBuyEditor('ammStock',idx,buySp);
  const priceSp  =mkText(stock?.price!=null?stock.price.toLocaleString('it-IT',{minimumFractionDigits:2,maximumFractionDigits:2}):'—');
  const ctrvCell =mkEditNum(r.ctrv||'',(v)=>{const pd2=getPD('ammStock');if(pd2){pd2.rows[idx].ctrv=v;recalcAmmStock();}},ctrvBlocked);
  const qtySp    =document.createElement('div');qtySp.className='flat-cell r';qtySp.dataset.role='qty';qtySp.textContent='—';
  const plSp     =document.createElement('div');plSp.className='flat-cell r pl-cell';plSp.dataset.role='pl';plSp.title='Plus/minusvalenza dal prezzo e cambio di acquisto (in EUR e %)';plSp.textContent='—';
  const pctSp    =document.createElement('div');pctSp.className='flat-cell r';pctSp.dataset.role='pct';pctSp.style.fontFamily='inherit';pctSp.textContent='—';
  const ytdSp    =document.createElement('div');ytdSp.className='flat-cell r';ytdSp.dataset.role='ytd';
  if(stock?.ytd_pct!=null){const v=stock.ytd_pct*100;ytdSp.textContent=(v>=0?'+':'')+v.toFixed(2)+'%';ytdSp.className+=(v>=0?' flat-ytd-p':' flat-ytd-n');}
  else ytdSp.textContent='—';

  const delBtn=document.createElement('button');delBtn.className='row-del-btn';delBtn.textContent='✕';delBtn.title='Rimuovi';
  delBtn.onclick=()=>deleteStockRow(idx);

  row.appendChild(tickerWrap);row.appendChild(sectorSp);row.appendChild(buySp);
  row.appendChild(priceSp);row.appendChild(ctrvCell);row.appendChild(qtySp);row.appendChild(plSp);
  row.appendChild(pctSp);row.appendChild(ytdSp);row.appendChild(delBtn);
  return row;
}

function applyStockToRow(idx,stock){
  const pd=getPD('ammStock');if(!pd)return;
  const row=pd.rows[idx];
  if(row.company!==stock.company){delete row.basePrice;delete row.baseRate;delete row.baseDate;} // new instrument → new baseline
  row.company=stock.company;
  row.ticker=stock.ticker||'';
  buildTab('ammStock');
}

function addStockRow(){
  const pd=getPD('ammStock');if(!pd)return;
  if(!pd.rows)pd.rows=[];
  pd.rows.push({ticker:'',company:'',ctrv:0});
  buildTab('ammStock');
}

function deleteStockRow(idx){
  const pd=getPD('ammStock');if(!pd||!pd.rows)return;
  pd.rows.splice(idx,1);
  buildTab('ammStock');
}

// ── Purchase-baseline editor (manual "bought at" price + date/time, or auto) ───
function toLocalInput(iso){
  const d=iso?new Date(iso):new Date();if(isNaN(d))return '';
  const p=n=>String(n).padStart(2,'0');
  return `${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())}T${p(d.getHours())}:${p(d.getMinutes())}`;
}
function ammInstrument(ptfl,r){return ptfl==='ammEtf'?(r.ticker?etfByTicker(r.ticker):null):(r.company?stockByCompany(r.company):null);}
let _buyEditEl=null;
function _buyOutside(e){if(_buyEditEl&&!_buyEditEl.contains(e.target))closeBuyEditor();}
function closeBuyEditor(){if(_buyEditEl){_buyEditEl.remove();_buyEditEl=null;document.removeEventListener('mousedown',_buyOutside,true);}}
function openBuyEditor(ptfl,idx,anchor){
  const pd=getPD(ptfl);if(!pd||!pd.rows[idx])return;
  const r=pd.rows[idx];const inst=ammInstrument(ptfl,r);
  if(!inst){toast('Seleziona prima un titolo','err');return;}
  closeBuyEditor();
  const box=document.createElement('div');box.className='buy-edit';_buyEditEl=box;
  const defPrice=r.basePrice!=null?r.basePrice:(inst.price!=null?inst.price:'');
  box.innerHTML=`
    <label>Prezzo di acquisto (${inst.currency||'EUR'})</label>
    <input type="number" step="0.0001" id="be-price" value="${defPrice}">
    <label>Data e ora di acquisto</label>
    <input type="datetime-local" id="be-dt" value="${toLocalInput(r.baseDate)}">
    <div class="buy-edit-row"><button class="be-save" id="be-save">Salva</button><button class="be-auto" id="be-auto" title="Usa prezzo/ora correnti e mantieni l'aggiornamento automatico">Automatico</button></div>
    <div class="buy-edit-row" style="margin-top:6px"><button class="be-cancel" id="be-cancel">Annulla</button></div>`;
  document.body.appendChild(box);
  // le misure a schermo sono già scalate dal page zoom: box largo 220px CSS = 220×zoom a schermo
  const rc=visRect(anchor),pz=window.PAGE_ZOOM||1;
  let top=rc.bottom+4,left=rc.right-220*pz;if(left<8)left=8;if(top+210*pz>window.innerHeight)top=Math.max(8,rc.top-210*pz);
  box.style.top=(top/pz)+'px';box.style.left=(left/pz)+'px';
  box.querySelector('#be-cancel').onclick=closeBuyEditor;
  box.querySelector('#be-auto').onclick=()=>{
    delete r.basePrice;delete r.baseRate;delete r.baseDate;delete r.basePriceDate;delete r.baseManual;
    closeBuyEditor();buildTab(ptfl);if(AID)savePtfls(true);
  };
  box.querySelector('#be-save').onclick=async()=>{
    const price=parseFloat(box.querySelector('#be-price').value);
    const dtv=box.querySelector('#be-dt').value;
    if(!(price>0)){toast('Inserisci un prezzo valido','err');return;}
    r.basePrice=price;r.baseManual=true;
    if(dtv){const d=new Date(dtv);r.baseDate=isNaN(d)?new Date().toISOString():d.toISOString();r.basePriceDate=dtv.slice(0,10);}
    else{r.baseDate=new Date().toISOString();r.basePriceDate=null;}
    const cur=(inst.currency||'EUR').toUpperCase();let rate=_fxRates[cur]||1;
    const dOnly=(dtv||'').slice(0,10);
    if(dOnly){try{const j=await(await fetch('/api/fx?date='+dOnly)).json();if(j&&j.rates&&j.rates[cur])rate=j.rates[cur];}catch(e){}}
    r.baseRate=rate;
    closeBuyEditor();buildTab(ptfl);if(AID)savePtfls(true);
  };
  setTimeout(()=>document.addEventListener('mousedown',_buyOutside,true),0);
}

function recalcAmmStock(){
  const pd=getPD('ammStock');if(!pd||!pd.rows)return;
  const rows=pd.rows;
  const totalCtrv=rows.reduce((s,r)=>s+(r.ctrv||0),0);
  let wytdNum=0,wytdDen=0,totalValueNow=0,capturedNew=false;
  const signEur=v=>(v>=0?'+':'−')+'€'+Math.abs(Math.round(v)).toLocaleString('it-IT',{useGrouping:'always'});
  const plColor=v=>v>=0?'rgba(38,160,90,.95)':'rgba(214,60,48,.95)';

  document.querySelectorAll('.ammstock-row').forEach((rowEl,i)=>{
    const r=rows[i];if(!r)return;
    const stock=r.company?stockByCompany(r.company):null;
    const price=stock?.price||0;
    const ctrv=r.ctrv||0;
    const curRate=_fxRates[(stock?.currency||'EUR').toUpperCase()]||1; // EUR→quote currency, now
    const pct=totalCtrv>0?ctrv/totalCtrv:0;
    const ytd=stock?.ytd_pct;

    // Baseline = price + EUR/quote-currency rate captured when the position was first valued.
    // Shares held are locked at that baseline; P/L (in EUR) tracks price + FX moves since.
    let qty=null,plEur=null;
    if(ctrv>0&&price>0){
      if(!r.basePrice){r.basePrice=price;r.baseRate=curRate;r.baseDate=new Date().toISOString();r.basePriceDate=stock?.price_date||null;capturedNew=true;}
      const baseRate=r.baseRate||curRate;
      qty=ctrv*baseRate/r.basePrice;          // shares bought at the baseline
      const valEur=qty*price/curRate;          // current value back in EUR (live FX)
      plEur=valEur-ctrv;
      totalValueNow+=valEur;
    }else{
      totalValueNow+=ctrv; // unpriced row → assume flat (no gain/loss)
    }

    const bEl=rowEl.querySelector('[data-role="buy"]');
    if(bEl){
      if(r.basePrice!=null){
        bEl.innerHTML=`<span>${r.basePrice.toLocaleString('it-IT',{minimumFractionDigits:2,maximumFractionDigits:2})}</span><span class="buy-dt">${fmtShortDt(r.baseDate)}</span>`;
        bEl.title=r.basePriceDate?('Prezzo Refinitiv del '+r.basePriceDate):'';
      }else{bEl.textContent='—';bEl.title='';}
    }
    const qEl=rowEl.querySelector('[data-role="qty"]');
    const lEl=rowEl.querySelector('[data-role="pl"]');
    const pEl=rowEl.querySelector('[data-role="pct"]');
    if(qEl)qEl.textContent=qty!=null?qty.toLocaleString('it-IT',{minimumFractionDigits:3,maximumFractionDigits:3}):'—';
    if(lEl){
      if(plEur!=null){
        const pp=ctrv>0?plEur/ctrv*100:0;
        lEl.innerHTML=`<span>${signEur(plEur)}</span><span class="pl-sub">${(plEur>=0?'+':'−')+Math.abs(pp).toFixed(2)}%</span>`;
        lEl.style.color=plColor(plEur);
      }else{lEl.innerHTML='—';lEl.style.color='';}
    }
    if(pEl)pEl.textContent=totalCtrv>0?(pct*100).toFixed(2)+'%':'—';

    if(ytd!=null&&ctrv>0){wytdNum+=ctrv*ytd;wytdDen+=ctrv;}
  });

  const netPl=totalValueNow-totalCtrv;
  const retPct=totalCtrv>0?netPl/totalCtrv*100:0;

  const tc=document.getElementById('ammstock-total-ctrv');
  const tp=document.getElementById('ammstock-total-pl');
  const wy=document.getElementById('ammstock-wytd');
  if(tc)tc.textContent=totalCtrv>0?'€'+totalCtrv.toLocaleString('it-IT',{minimumFractionDigits:0,maximumFractionDigits:0,useGrouping:'always'}):'—';
  if(tp){if(totalCtrv>0){tp.textContent=signEur(netPl);tp.style.color=plColor(netPl);}else{tp.textContent='—';tp.style.color='';}}
  if(wy){
    if(wytdDen>0){
      const v=(wytdNum/wytdDen)*100;
      wy.textContent=(v>=0?'+':'')+v.toFixed(2)+'%';
      wy.style.color=v>=0?'rgba(100,220,150,.9)':'rgba(255,120,100,.9)';
    }else{wy.textContent='—';}
  }

  const put=document.getElementById('ammstock-sum-put');
  const now=document.getElementById('ammstock-sum-now');
  const plE=document.getElementById('ammstock-sum-pl');
  const ret=document.getElementById('ammstock-sum-ret');
  if(put)put.textContent=fmt(totalCtrv);
  if(now)now.textContent=fmt(totalValueNow);
  if(plE){if(totalCtrv>0){plE.textContent=signEur(netPl);plE.style.color=plColor(netPl);}else{plE.textContent='—';plE.style.color='';}}
  if(ret){if(totalCtrv>0){ret.textContent=(netPl>=0?'+':'−')+Math.abs(retPct).toFixed(2)+'%';ret.style.color=plColor(netPl);}else{ret.textContent='—';ret.style.color='';}}

  refreshCtrvBar('ammStock');
  updateSaveInfo();
  // Persist freshly captured baselines so P/L survives reloads/price updates.
  if(capturedNew&&AID)savePtfls(true);
}

async function syncStockUniverse(){
  const msg=document.getElementById('ammstock-sync-msg');
  if(msg)msg.textContent='Sincronizzazione in corso…';
  try{
    const r=await fetch('/api/stock-universe/sync',{method:'POST'});
    const d=await r.json();
    await loadStockUniverse();
    if(msg)msg.textContent=`✓ ${d.synced} titoli sincronizzati`;
    buildTab('ammStock');
  }catch(e){
    if(msg)msg.textContent='Errore durante la sincronizzazione';
  }
}

// ── AMMINISTRATO ETF (flat etf table) ─────────────────────────────────────────
function buildAmmEtfTab(container,pd){
  const rows=pd.rows||[];
  const tb=document.createElement('div');
  tb.style.cssText='display:flex;align-items:center;gap:10px;margin-bottom:14px;';
  tb.innerHTML=`<button onclick="syncEtfUniverse()" style="padding:5px 14px;background:var(--navy);color:var(--gold-l);border:1.5px solid rgba(201,168,76,.35);border-radius:6px;font-size:11px;font-weight:700;cursor:pointer;">${ic('refresh')} Sincronizza da Excel</button><span id="ammetf-sync-msg" style="font-size:10px;color:var(--txt3)"></span><span title="La quantità converte il CTRV in EUR nella valuta dell'ETF" style="margin-left:auto;font-size:10px;color:var(--txt3)">Cambio EUR/USD: <strong>${_fxRates.USD?_fxRates.USD.toFixed(4):'—'}</strong>${_fxDate?' · '+_fxDate:''}</span>`;
  container.appendChild(tb);

  const tbl=document.createElement('div');tbl.id='ammetf-table';tbl.className='tbl';

  const hdr=document.createElement('div');hdr.className='tbl-ch g8-etf';
  [['Nome ETF',''],['Settore',''],['Tipo',''],['Acquisto','col-r'],['Prezzo','col-r'],['CTRV (€)','col-r'],['TER %','col-r'],['Quantità','col-r'],['P/L (€)','col-r'],['% Port.','col-r'],['YTD %','col-r'],['','']].forEach(([h,cls])=>{
    const s=document.createElement('span');if(cls)s.className=cls;s.textContent=h;hdr.appendChild(s);
  });
  tbl.appendChild(hdr);

  const rowsEl=document.createElement('div');rowsEl.id='ammetf-rows';
  rows.forEach((r,i)=>rowsEl.appendChild(makeEtfRow(r,i)));
  tbl.appendChild(rowsEl);

  const foot=document.createElement('div');foot.id='ammetf-foot';foot.className='grand-row g8-etf';
  [['TOTALE PORTAFOGLIO',''],['',''],['',''],['',''],['',''],['—','col-r grand-val'],['—','col-r'],['',''],['—','col-r grand-val'],['',''],['—','col-r'],['','']].forEach(([h,cls],i)=>{
    const s=document.createElement('span');if(cls)s.className=cls;
    if(i===5)s.id='ammetf-total-ctrv';
    if(i===6)s.id='ammetf-wter';
    if(i===8)s.id='ammetf-total-pl';
    if(i===10)s.id='ammetf-wytd';
    s.textContent=h;foot.appendChild(s);
  });
  tbl.appendChild(foot);
  container.appendChild(tbl);

  const addBtn=document.createElement('button');
  addBtn.className='add-row-btn';addBtn.style.marginTop='8px';
  addBtn.textContent='+ Aggiungi ETF';addBtn.onclick=()=>addEtfRow();
  container.appendChild(addBtn);

  const summ=document.createElement('div');summ.className='amm-stats';
  summ.innerHTML=`
    <div class="amm-stat"><div class="amm-stat-lbl">Capitale investito</div><div class="amm-stat-val" id="ammetf-sum-put">—</div></div>
    <div class="amm-stat"><div class="amm-stat-lbl">Valore attuale</div><div class="amm-stat-val" id="ammetf-sum-now">—</div></div>
    <div class="amm-stat"><div class="amm-stat-lbl">Plus/Minusvalenze nette</div><div class="amm-stat-val" id="ammetf-sum-pl">—</div></div>
    <div class="amm-stat"><div class="amm-stat-lbl">Rendimento</div><div class="amm-stat-val" id="ammetf-sum-ret">—</div></div>`;
  container.appendChild(summ);

  recalcAmmEtf();
}

function makeEtfRow(r,idx){
  const row=document.createElement('div');
  row.className='fund-row g8-etf ammetf-row';row.dataset.idx=idx;
  const etf=r.ticker?etfByTicker(r.ticker):null;
  const ctrvBlockedEtf=getTotal('ammEtf')<=0;

  const tickerWrap=document.createElement('div');tickerWrap.style.cssText='position:relative;padding:2px 6px;display:flex;align-items:center;';
  const tickerInp=document.createElement('input');
  tickerInp.type='text';tickerInp.placeholder='Cerca ETF…';tickerInp.value=(etf?.etf_name||r.ticker||'');tickerInp.className='flat-ticker flat-name';
  const dropdown=document.createElement('div');dropdown.className='fund-dd';dropdown.style.display='none';
  tickerInp.oninput=()=>{
    const q=tickerInp.value.trim();
    if(!q){dropdown.style.display='none';return;}
    const hits=etfSearch(q).slice(0,10);
    if(!hits.length){dropdown.style.display='none';return;}
    dropdown.innerHTML='';
    hits.forEach(e=>{
      const opt=document.createElement('div');opt.className='fund-dd-item';
      opt.innerHTML=`<span class="fund-dd-isin">${e.ticker||'—'}</span><span class="fund-dd-name">${e.etf_name||''}</span><span style="color:var(--txt3);font-size:10px;flex-shrink:0">${e.sector||''}</span>`;
      opt.onmousedown=()=>{tickerInp.value=e.etf_name||e.ticker||'';dropdown.style.display='none';applyEtfToRow(idx,e);};
      dropdown.appendChild(opt);
    });
    dropdown.style.display='block';
  };
  tickerInp.onblur=()=>setTimeout(()=>{dropdown.style.display='none';},150);
  tickerWrap.appendChild(tickerInp);tickerWrap.appendChild(dropdown);

  const mkLeft=(text)=>{const d=document.createElement('div');d.className='flat-cell';d.textContent=text||'—';return d;};
  const mkText=(text,role)=>{const d=document.createElement('div');d.className='flat-cell r';if(role)d.dataset.role=role;d.textContent=text||'—';return d;};
  const mkEditNum=(val,onchange,blocked)=>{
    const d=document.createElement('div');d.className='flat-cell r';
    const inp=document.createElement('input');inp.type='number';inp.min='0';inp.step='100';inp.value=val||'';inp.className='flat-inp';
    if(blocked){inp.disabled=true;inp.placeholder='Imposta totale';inp.title='Inserisci prima il totale';}
    inp.onchange=()=>{if(onchange)onchange(parseFloat(inp.value)||0);};
    d.appendChild(inp);return d;
  };

  const sectorSp=mkLeft(etf?.sector||'');
  const typeSp  =mkLeft(etf?.type||'');
  const buySp   =document.createElement('div');buySp.className='flat-cell r buy-cell';buySp.dataset.role='buy';buySp.textContent='—';
  buySp.title='Clicca per impostare prezzo e data/ora di acquisto';buySp.onclick=()=>openBuyEditor('ammEtf',idx,buySp);
  const priceSp =mkText(etf?.price!=null?etf.price.toLocaleString('it-IT',{minimumFractionDigits:2,maximumFractionDigits:2}):'—');
  const ctrvCell=mkEditNum(r.ctrv||'',(v)=>{const pd2=getPD('ammEtf');if(pd2){pd2.rows[idx].ctrv=v;recalcAmmEtf();}},ctrvBlockedEtf);

  const terVal=r.ter!=null?r.ter:(etf?.ter!=null?etf.ter:null);
  const terCell=document.createElement('div');terCell.className='flat-cell r';
  const terInp=document.createElement('input');terInp.type='number';terInp.min='0';terInp.step='0.01';terInp.placeholder='0.20';
  terInp.value=terVal!=null?(terVal*100).toFixed(2):'';terInp.className='flat-inp';
  terInp.style.cssText='border:1px solid var(--bdr2);border-radius:4px;padding:3px 5px;background:var(--surf2);width:100%;text-align:right;font-size:11px;font-family:"JetBrains Mono",monospace;color:var(--txt);';
  terInp.onchange=()=>{const pd2=getPD('ammEtf');if(!pd2)return;const v=terInp.value===''?null:parseFloat(terInp.value)/100;pd2.rows[idx].ter=v;recalcAmmEtf();};
  terCell.appendChild(terInp);

  const qtySp=document.createElement('div');qtySp.className='flat-cell r';qtySp.dataset.role='qty';qtySp.textContent='—';
  const plSp=document.createElement('div');plSp.className='flat-cell r pl-cell';plSp.dataset.role='pl';plSp.title='Plus/minusvalenza dal prezzo e cambio di acquisto (in EUR e %)';plSp.textContent='—';
  const pctSp=document.createElement('div');pctSp.className='flat-cell r';pctSp.dataset.role='pct';pctSp.style.fontFamily='inherit';pctSp.textContent='—';
  const ytdSp=document.createElement('div');ytdSp.className='flat-cell r';ytdSp.dataset.role='ytd';
  if(etf?.ytd_pct!=null){const v=etf.ytd_pct*100;ytdSp.textContent=(v>=0?'+':'')+v.toFixed(2)+'%';ytdSp.className+=(v>=0?' flat-ytd-p':' flat-ytd-n');}
  else ytdSp.textContent='—';

  const delBtn=document.createElement('button');delBtn.className='row-del-btn';delBtn.textContent='✕';delBtn.title='Rimuovi';
  delBtn.onclick=()=>deleteEtfRow(idx);

  row.appendChild(tickerWrap);row.appendChild(sectorSp);row.appendChild(typeSp);row.appendChild(buySp);
  row.appendChild(priceSp);row.appendChild(ctrvCell);row.appendChild(terCell);row.appendChild(qtySp);row.appendChild(plSp);
  row.appendChild(pctSp);row.appendChild(ytdSp);row.appendChild(delBtn);
  return row;
}

function applyEtfToRow(idx,etf){const pd=getPD('ammEtf');if(!pd)return;const row=pd.rows[idx];if(row.ticker!==etf.ticker){delete row.basePrice;delete row.baseRate;delete row.baseDate;}row.ticker=etf.ticker||'';if(etf.ter!=null&&row.ter==null)row.ter=etf.ter;buildTab('ammEtf');}
function addEtfRow(){const pd=getPD('ammEtf');if(!pd)return;if(!pd.rows)pd.rows=[];pd.rows.push({ticker:'',ctrv:0});buildTab('ammEtf');}
function deleteEtfRow(idx){const pd=getPD('ammEtf');if(!pd||!pd.rows)return;pd.rows.splice(idx,1);buildTab('ammEtf');}
function recalcAmmEtf(){
  const pd=getPD('ammEtf');if(!pd||!pd.rows)return;
  const rows=pd.rows;const totalCtrv=rows.reduce((s,r)=>s+(r.ctrv||0),0);
  let wytdNum=0,wytdDen=0,wterNum=0,wterDen=0,totalValueNow=0,capturedNew=false;
  const signEur=v=>(v>=0?'+':'−')+'€'+Math.abs(Math.round(v)).toLocaleString('it-IT',{useGrouping:'always'});
  const plColor=v=>v>=0?'rgba(38,160,90,.95)':'rgba(214,60,48,.95)';
  document.querySelectorAll('.ammetf-row').forEach((rowEl,i)=>{
    const r=rows[i];if(!r)return;
    const etf=r.ticker?etfByTicker(r.ticker):null;
    const price=etf?.price||0;const ctrv=r.ctrv||0;
    const curRate=_fxRates[(etf?.currency||'EUR').toUpperCase()]||1; // EUR→quote currency, now
    const pct=totalCtrv>0?ctrv/totalCtrv:0;const ytd=etf?.ytd_pct;
    const ter=r.ter!=null?r.ter:(etf?.ter!=null?etf.ter:null);

    // Baseline = price + EUR/quote-currency rate captured when the position was first valued.
    let qty=null,plEur=null;
    if(ctrv>0&&price>0){
      if(!r.basePrice){r.basePrice=price;r.baseRate=curRate;r.baseDate=new Date().toISOString();r.basePriceDate=etf?.price_date||null;capturedNew=true;}
      const baseRate=r.baseRate||curRate;
      qty=ctrv*baseRate/r.basePrice;          // shares bought at the baseline
      const valEur=qty*price/curRate;          // current value back in EUR (live FX)
      plEur=valEur-ctrv;
      totalValueNow+=valEur;
    }else{
      totalValueNow+=ctrv; // unpriced row → assume flat
    }

    const bEl=rowEl.querySelector('[data-role="buy"]');
    if(bEl){
      if(r.basePrice!=null){
        bEl.innerHTML=`<span>${r.basePrice.toLocaleString('it-IT',{minimumFractionDigits:2,maximumFractionDigits:2})}</span><span class="buy-dt">${fmtShortDt(r.baseDate)}</span>`;
        bEl.title=r.basePriceDate?('Prezzo Refinitiv del '+r.basePriceDate):'';
      }else{bEl.textContent='—';bEl.title='';}
    }
    const qEl=rowEl.querySelector('[data-role="qty"]');const lEl=rowEl.querySelector('[data-role="pl"]');const pEl=rowEl.querySelector('[data-role="pct"]');
    if(qEl)qEl.textContent=qty!=null?qty.toLocaleString('it-IT',{minimumFractionDigits:3,maximumFractionDigits:3}):'—';
    if(lEl){if(plEur!=null){const pp=ctrv>0?plEur/ctrv*100:0;lEl.innerHTML=`<span>${signEur(plEur)}</span><span class="pl-sub">${(plEur>=0?'+':'−')+Math.abs(pp).toFixed(2)}%</span>`;lEl.style.color=plColor(plEur);}else{lEl.innerHTML='—';lEl.style.color='';}}
    if(pEl)pEl.textContent=totalCtrv>0?(pct*100).toFixed(2)+'%':'—';
    if(ytd!=null&&ctrv>0){wytdNum+=ctrv*ytd;wytdDen+=ctrv;}
    if(ter!=null&&ctrv>0){wterNum+=ctrv*ter;wterDen+=ctrv;}
  });
  const netPl=totalValueNow-totalCtrv;
  const retPct=totalCtrv>0?netPl/totalCtrv*100:0;

  const tc=document.getElementById('ammetf-total-ctrv');const tp=document.getElementById('ammetf-total-pl');const wy=document.getElementById('ammetf-wytd');const wt=document.getElementById('ammetf-wter');
  if(tc)tc.textContent=totalCtrv>0?'€'+totalCtrv.toLocaleString('it-IT',{minimumFractionDigits:0,maximumFractionDigits:0,useGrouping:'always'}):'—';
  if(tp){if(totalCtrv>0){tp.textContent=signEur(netPl);tp.style.color=plColor(netPl);}else{tp.textContent='—';tp.style.color='';}}
  if(wy){if(wytdDen>0){const v=(wytdNum/wytdDen)*100;wy.textContent=(v>=0?'+':'')+v.toFixed(2)+'%';wy.style.color=v>=0?'rgba(100,220,150,.9)':'rgba(255,120,100,.9)';}else{wy.textContent='—';}}
  if(wt){if(wterDen>0){wt.textContent=(wterNum/wterDen*100).toFixed(3)+'%';}else{wt.textContent='—';}}

  const put=document.getElementById('ammetf-sum-put');const now=document.getElementById('ammetf-sum-now');const plE=document.getElementById('ammetf-sum-pl');const ret=document.getElementById('ammetf-sum-ret');
  if(put)put.textContent=fmt(totalCtrv);
  if(now)now.textContent=fmt(totalValueNow);
  if(plE){if(totalCtrv>0){plE.textContent=signEur(netPl);plE.style.color=plColor(netPl);}else{plE.textContent='—';plE.style.color='';}}
  if(ret){if(totalCtrv>0){ret.textContent=(netPl>=0?'+':'−')+Math.abs(retPct).toFixed(2)+'%';ret.style.color=plColor(netPl);}else{ret.textContent='—';ret.style.color='';}}

  refreshCtrvBar('ammEtf');
  updateSaveInfo();
  if(capturedNew&&AID)savePtfls(true);
}
async function syncEtfUniverse(){
  const msg=document.getElementById('ammetf-sync-msg');const btn=event?.target;
  if(btn){btn.textContent='…';btn.disabled=true;}
  try{
    const r=await fetch('/api/etf-universe/sync',{method:'POST'});const d=await r.json();
    await loadEtfUniverse();if(msg)msg.textContent=`✓ ${d.synced} ETF sincronizzati`;buildTab('ammEtf');
  }catch(e){if(msg)msg.textContent='Errore';}
  finally{if(btn){btn.innerHTML=ic('refresh')+' Sincronizza da Excel';btn.disabled=false;}}
}

// ── BUILD TAB ─────────────────────────────────────────────────────────────────
function buildTab(ptfl){
  const container=document.getElementById('tp-'+ptfl);if(!container)return;
  const pd=getPD(ptfl);if(!pd&&!isPacType(ptfl)){container.innerHTML='';return;}
  if(ptfl==='ammStock'){container.innerHTML='';buildAmmStockTab(container,pd);return;}
  if(ptfl==='ammEtf'){container.innerHTML='';buildAmmEtfTab(container,pd);return;}
  if(ptfl==='darta')migrateDartaEasySwitch();
  const total=getTotal(ptfl);const weights=getW(ptfl);const order=getWOrder(ptfl);
  container.innerHTML='';

  // ── DARTA: plan selector (before everything) ─────────────────────────────
  if(ptfl==='darta'){
    const plan=pd?.plan||'pro';
    const planSel=document.createElement('div');planSel.className='darta-plan-sel';
    planSel.innerHTML=`<button class="dps-btn${plan==='pro'?' dps-active':''}" onclick="setDartaPlan('pro')"><div class="dps-btn-title">${ic('layers')} Challenge Pro</div><div class="dps-btn-sub">Caricamento 2.00% (≤€100k) / 1.50% (>€100k) · Commissioni per fondo</div></button><button class="dps-btn${plan==='private'?' dps-active':''}" onclick="setDartaPlan('private')"><div class="dps-btn-title">${ic('gem')} Challenge Private</div><div class="dps-btn-sub">Nessun caricamento · Commissioni di gestione basate sul capitale investito</div></button>`;
    container.appendChild(planSel);
    const copertura=getDartaCopertura(total);
    let cgVal,cgNote,caricVal,caricNote;
    if(plan==='private'){
      const cg=getDartaPrivateCG(total);
      const cls=getDartaPrivateClass(total);
      cgVal=cg.toFixed(2)+'%';cgNote='Classe '+cls+' · varia con capitale netto versato';
      caricVal='0.00%';caricNote='Nessun caricamento ingresso';
    }else{
      cgVal='0.35% – 2.00%';cgNote='Varia per fondo interno selezionato';
      caricVal='2.00% / 1.50%';caricNote='≤€100k · >€100k cumulo premi';
    }
    const copNote=total<=500000?'Soglia ≤€500k (flat 0.25%)':'Scalata: 0.25% × €500k / investito';
    const planInfo=document.createElement('div');planInfo.className='darta-plan-info';
    planInfo.innerHTML=`<div class="dpi-card"><div class="dpi-lbl">Piano selezionato</div><div class="dpi-val">${plan==='pro'?'Challenge Pro':'Challenge Private'}</div><div class="dpi-note">Tariffa ${plan==='pro'?'USL4S.03':'USL4S.05'}</div></div><div class="dpi-card"><div class="dpi-lbl">Caricamento</div><div class="dpi-val" id="dpi-caric-val">${caricVal}</div><div class="dpi-note" id="dpi-caric-note">${caricNote}</div></div><div class="dpi-card"><div class="dpi-lbl">Comm. gestione (CG)</div><div class="dpi-val" id="dpi-cg-val">${cgVal}</div><div class="dpi-note" id="dpi-cg-note">${cgNote}</div></div><div class="dpi-card"><div class="dpi-lbl">Copertura assicurativa</div><div class="dpi-val" id="dpi-cop-val">${(copertura*100).toFixed(4)}%</div><div class="dpi-note" id="dpi-cop-note">${copNote}</div></div>`;
    container.appendChild(planInfo);
    const rsRow=document.createElement('div');rsRow.style.cssText='margin-bottom:.875rem;display:flex;gap:.625rem;align-items:center;';
    rsRow.innerHTML=`<button style="padding:5px 14px;background:var(--navy);color:var(--gold-l);border:1.5px solid rgba(201,168,76,.35);border-radius:6px;font-size:11px;font-weight:700;cursor:pointer;" onclick="resyncDartaCommissions()">${ic('refresh')} Ri-sincronizza CG/SC da Excel</button><span style="font-size:10px;color:var(--txt3)">Aggiorna CG e SC di tutti i fondi già salvati usando il file DARTA CHP+CHPV.xlsx</span>`;
    container.appendChild(rsRow);
  }

  const mode=pd?._mode||'weights';
  const mb=document.createElement('div');mb.className='mode-bar';
  mb.innerHTML=`<span class="mode-lbl">Modalità</span><div class="mode-toggle"><button class="mt-btn pesi${mode==='weights'?' mt-active':''}" onclick="setPortfolioMode('${ptfl}','weights')">${ic('scale')} Pesi</button><button class="mt-btn ctrv${mode==='ctrv'?' mt-active':''}" onclick="setPortfolioMode('${ptfl}','ctrv')">€ CTRV</button></div><span class="mode-hint">${mode==='weights'?'Pesi → CTRV derivato':'CTRV fisso → Pesi derivati'}</span>`;
  container.appendChild(mb);
  const wt=document.createElement('div');wt.className='sec-title';wt.textContent='Pesi Asset — '+TL[ptfl];container.appendChild(wt);
  const ww=document.createElement('div');buildWeights(ptfl,ww,weights,order,mode);container.appendChild(ww);

  if(isPacType(ptfl)){
    const par=PAC_PARENT[ptfl];const ppd=CD.portfolios?.[par];
    const rata=pacTotal(ptfl);const ms=ppd?.pac?.mesi||15;
    const isES=ptfl==='pac';
    const srcLabel=isES?(easySwitchSourceLabel()||'(nessun fondo)'):'START & GO';
    const switchTotal=isES?easySwitchAmount():((ppd?.total||0)*((ppd?.weights?.['START & GO']?.pct||0)/100));
    const f=document.createElement('div');f.className='pac-flow';
    f.innerHTML=`<span class="flow-chip chip-sg">${isES?srcLabel:'START &amp; GO'}</span><span class="flow-arr">→</span><span style="font-size:12px;color:var(--txt2)">Rata mensile: <strong>${fmt(rata)}</strong></span><span class="flow-arr">→</span><span class="flow-chip chip-pac">${TL[ptfl]}</span>`;
    container.appendChild(f);
    const stats=document.createElement('div');stats.className='pac-stats';
    stats.innerHTML=`<div class="pac-stat"><div class="psl">${isES?'Totale da convertire':'Totale START &amp; GO'}</div><div class="psv">${fmt(switchTotal)}</div></div><div class="pac-stat"><div class="psl">Rata effettiva</div><div class="psv"><span>${fmt(pacActual(ptfl))}</span></div></div><div class="pac-stat"><div class="psl">Durata</div><div class="psv">${ms} <span>mesi</span></div></div>`;
    container.appendChild(stats);
    const pac=ppd?.pac||{};
    const useRnd=pac.useRounded||false;
    const roundedVal=pac.roundedTotal||(Math.floor(pacActual(ptfl)/100)*100)||Math.round(pacActual(ptfl));
    const residual=pacActual(ptfl)-roundedVal;
    const roundBar=document.createElement('div');roundBar.className='pac-round-bar';
    roundBar.innerHTML=`<label class="pac-round-lbl"><input type="checkbox" id="pac-use-rounded-${ptfl}" ${useRnd?'checked':''} onchange="setPacUseRounded('${ptfl}',this.checked)"/> Arrotonda rata mensile</label>${useRnd?`<div class="pac-round-inp-row"><span class="tc-pre">€</span><input type="text" inputmode="numeric" class="tc-inp w130" value="${fmtEuroInp(roundedVal)}" onfocus="onEuroFocus(this)" onchange="setPacRoundedTotal('${ptfl}',this.value);onEuroBlur(this)"/><span class="pac-residual">(residuo: ${fmt(residual)})</span></div>`:''}`;
    container.appendChild(roundBar);
  }
  if(ptfl==='sicavPac'){
    const pd2=CD.portfolios?.sicavPac||{};const tot=pd2.total||0;const ms=pd2.mesi||0;const rata=sicavPacRata();
    const f=document.createElement('div');f.className='pac-stats';
    f.innerHTML=`<div class="pac-stat"><div class="psl">Totale da investire</div><div class="psv">${fmt(tot)}</div></div><div class="pac-stat"><div class="psl">Durata</div><div class="psv">${ms||'—'} <span>mesi</span></div></div><div class="pac-stat"><div class="psl">Rata mensile</div><div class="psv"><span>${fmt(rata)}</span></div></div>`;
    container.appendChild(f);
    if(!(rata>0)){const warn=document.createElement('div');warn.className='es-warn';warn.style.marginBottom='.75rem';warn.textContent='Imposta il totale e i mesi del PAC nella barra in alto per calcolare la rata mensile.';container.appendChild(warn);}
  }
  if(ptfl==='darta'){
    // START & GO stays the residual asset class: its % is the leftover of the
    // DARTA total after the real funds. Mirror that controvalore into the Easy
    // Switch amount so the rata and the child tab stay in sync (display only).
    const cfg=pd?.pac||{};
    if(cfg.sourceKey==='__SG__')cfg.totalSwitch=easySwitchAmount();
    const es=document.createElement('div');es.innerHTML=easySwitchPanelHTML();container.appendChild(es.firstElementChild);
  }

  const pl=document.createElement('div');pl.className='ptfl-lbl';
  pl.innerHTML=`<span class="ptfl-title">${TL[ptfl]}</span><span class="ptfl-total">${fmt(total)}</span>`;
  container.appendChild(pl);

  order.forEach(ac=>{
    const acd=weights[ac];if(!acd)return;
    if(ac==='START & GO'){
      const bl=document.createElement('div');bl.className='tbl';bl.style.marginBottom='.875rem';
      const sh=document.createElement('div');sh.className=`ac-hdr ${gcls(ptfl)} sg`;
      const sgVal=(pd?.total||0)*(acd.pct/100);
      if(ptfl==='darta'){
        sh.innerHTML=`<span>START &amp; GO</span><span></span><span></span><span></span><span class="col-r">${fmt(sgVal)}</span><span class="col-r">${fmtP(acd.pct/100)}</span><span></span><span></span><span></span><span></span><span></span><span></span>`;
      }else{
        sh.innerHTML=`<span>${acLabel(ac,ptfl)}</span><span></span><span></span><span></span><span class="col-r">${fmt(sgVal)}</span><span class="col-r">${fmtP(acd.pct/100)}</span><span></span><span></span><span></span>`;
      }
      bl.appendChild(sh);
      const n=document.createElement('div');n.className='sg-note-row';n.innerHTML=`↗ ${fmt(sgVal)} allocati al ${TL[PAC_CHILD[ptfl]]||'PAC'}`;
      bl.appendChild(n);container.appendChild(bl);return;
    }
    const acPct=acd.pct/100,acVal=total*acPct;
    const clsName=CLSC[ac]||(acd._type==='bond'?'bd':'eq');
    const acId=`ac-${ptfl}-${ac.replace(/[^a-z0-9]/gi,'_')}`;
    const bl=document.createElement('div');bl.className='tbl';bl.style.marginBottom='.875rem';
    const sh=document.createElement('div');sh.className=`ac-hdr ${gcls(ptfl)} ${clsName}`;
    sh.setAttribute('onclick',`toggleAc('${acId}')`);
    sh.innerHTML=gridCells(ptfl,{name:`<span>${ac} <span class="ac-toggle open" id="tog-${acId}">▶</span></span>`,
      pport:_cr(`achp-${ptfl}-${ac}`,fmtP(acPct)),ctrv:_cr(`ach-${ptfl}-${ac}`,fmt(acVal)),
      ct:_cr(`ach-wcm-${ptfl}-${ac}`),costo:_cr(`ach-costo-${ptfl}-${ac}`),costo2:_cr(`ach-costo2-${ptfl}-${ac}`)});
    bl.appendChild(sh);
    const acBody=document.createElement('div');acBody.className='ac-body';acBody.id=acId;
    acBody.appendChild(makeColHdr(ptfl));
    const subOrder=getSubOrder(ptfl,ac);
    subOrder.forEach(sub=>appendSub(acBody,ptfl,ac,acd,sub,acd.subs[sub],total));
    const acTr=document.createElement('div');acTr.className=`ac-total-row ${gcls(ptfl)}`;
    acTr.innerHTML=gridCells(ptfl,{name:`<span>Totale ${ac}</span>`,
      pport:_cr(`acp-${ptfl}-${ac}`),ctrv:_cr(`act-${ptfl}-${ac}`),
      ct:_cr(`act-wcm-${ptfl}-${ac}`),costo:_cr(`act-costo-${ptfl}-${ac}`),costo2:_cr(`act-costo2-${ptfl}-${ac}`)});
    acBody.appendChild(acTr);
    bl.appendChild(acBody);container.appendChild(bl);
  });

  const gr=document.createElement('div');gr.className=`grand-row ${gcls(ptfl)}`;
  const isAmmLegacy=ptfl==='amm'||ptfl==='ammAz';
  gr.innerHTML=gridCells(ptfl,{name:'<span>TOTALE PORTAFOGLIO</span>',
    ctrv:`<span class="col-r grand-val" id="grand-${ptfl}">—</span>`,
    ct:`<span class="col-r grand-val" id="grand-cop-${ptfl}">—</span>`,
    costo:`<span class="col-r grand-val" id="grand-costo-${ptfl}">—</span>`,
    costo2:`<span class="col-r grand-val" id="grand-costo2-${ptfl}">—</span>`,
    cg:isAmmLegacy?_cr(`wter-pct-${ptfl}`,'—',"color:rgba(201,168,76,.65);font-family:'JetBrains Mono',monospace;font-size:12px;font-weight:600"):null});
  container.appendChild(gr);
  // Quota del totale non ancora assegnata ai fondi (o assegnata in eccesso)
  const ua=document.createElement('div');ua.className='unalloc-row';ua.id=`unalloc-${ptfl}`;ua.style.display='none';
  container.appendChild(ua);
  if(isAmmLegacy){
    const tr=document.createElement('div');tr.className=`grand-row ${gcls(ptfl)}`;
    tr.style.cssText='background:var(--navy3);border-top:1px solid rgba(201,168,76,.3)';
    tr.innerHTML=gridCells(ptfl,{name:'<span style="color:rgba(255,255,255,.55);font-size:11px;font-weight:600">COSTO ANNUALE TER</span>',
      ctrv:_cr(`wter-euro-${ptfl}`,'—',"color:var(--gold-l);font-family:'JetBrains Mono',monospace;font-size:13px;font-weight:700")});
    container.appendChild(tr);
  }
  recalc(ptfl,total,weights,order);
}

function toggleAc(id){const body=document.getElementById(id);const tog=document.getElementById('tog-'+id);if(!body)return;const c=body.classList.toggle('collapsed');if(tog)tog.classList.toggle('open',!c);}

function calcCopertura(){const total=getTotal('darta');if(!total||total<=0)return null;return getDartaCopertura(total);}
function updateDartaInfoCards(total){
  const plan=getPD('darta')?.plan||'pro';
  const cop=getDartaCopertura(total)||0;
  const caricEl=document.getElementById('dpi-caric-val');const caricNoteEl=document.getElementById('dpi-caric-note');
  const cgEl=document.getElementById('dpi-cg-val');const cgNoteEl=document.getElementById('dpi-cg-note');
  const copEl=document.getElementById('dpi-cop-val');const copNoteEl=document.getElementById('dpi-cop-note');
  if(plan==='private'){
    if(caricEl)caricEl.textContent='0.00%';
    if(caricNoteEl)caricNoteEl.textContent='Nessun caricamento ingresso';
    const cg=getDartaPrivateCG(total);const cls=getDartaPrivateClass(total);
    if(cgEl)cgEl.textContent=cg.toFixed(2)+'%';
    if(cgNoteEl)cgNoteEl.textContent='Classe '+cls+' · varia con capitale netto versato';
  }else{
    const rate=total>100000?'1.50%':'2.00%';
    if(caricEl)caricEl.textContent=rate;
    if(caricNoteEl)caricNoteEl.textContent=total>100000?'Caricamento ridotto (>€100k)':'Caricamento standard (≤€100k)';
    if(cgEl)cgEl.textContent='0.35% – 2.00%';
    if(cgNoteEl)cgNoteEl.textContent='Varia per fondo interno selezionato';
  }
  if(copEl)copEl.textContent=(cop*100).toFixed(4)+'%';
  if(copNoteEl)copNoteEl.textContent=total<=500000?'Soglia ≤€500k (flat 0.25%)':'Scalata: 0.25% × €500k / investito';
}
function fmtCop(r){if(r==null)return'—';const v=r*100;return(v%1===0?v.toFixed(0):parseFloat(v.toFixed(4)).toString())+'%';}
const gcls=ptfl=>ptfl==='darta'?'g8-darta':ptfl==='pac'?'g8-pac':'g8';
// Colonne delle tabelle fondi. "% Port." è il campo che si compila (% del portafoglio);
// "% Sett." resta calcolato perché alimenta i donut e i sotto-settori.
const GCOLS={
  darta:['name','m','pport','ctrv','cg','sc','cop','ct','costo','psett','del'],
  pac:['name','m','pport','ctrv','cg','sc','cop','ct','costo','costo2','psett','del'],
  g8:['name','m','pport','ctrv','cg','psett','del'],
};
const gcols=ptfl=>ptfl==='darta'?GCOLS.darta:ptfl==='pac'?GCOLS.pac:GCOLS.g8;
// Una riga della griglia: celle per chiave di colonna, le mancanti restano vuote.
function gridCells(ptfl,cells){return gcols(ptfl).map(c=>cells[c]!=null?cells[c]:'<span></span>').join('');}
const _cr=(id,txt='—',st='')=>`<span class="col-r"${id?` id="${id}"`:''}${st?` style="${st}"`:''}>${txt}</span>`;

function makeColHdr(ptfl){
  const d=document.createElement('div');d.className=`tbl-ch ${gcls(ptfl)}`;
  const ms=CD?.portfolios?.darta?.pac?.mesi||15;
  const h=(t,c)=>`<span class="col-r"${c?` style="color:${c}"`:''}>${t}</span>`;
  d.innerHTML=gridCells(ptfl,{
    name:`<span>${ptfl==='darta'||ptfl==='pac'?'Azione / ISIN / Nome':'ISIN / Nome'}</span>`,
    pport:h('% Port. ✎','rgba(160,240,160,.9)'),
    ctrv:h('CTRV ✎','rgba(220,190,100,.9)'),
    cg:h((ptfl==='amm'||ptfl==='ammAz')?'TER ✎':'CG ✎','rgba(160,200,255,.9)'),
    sc:h('SC ✎','rgba(100,210,240,.9)'),
    cop:h('Cop.','rgba(255,200,120,.9)'),
    ct:h('Comm.Tot','rgba(255,180,80,.9)'),
    costo:h(ptfl==='pac'?'Costo/m':'Costo','rgba(255,160,60,.9)'),
    costo2:h(`Tot.×${ms}m`,'rgba(255,120,40,.9)'),
    psett:h('% Sett.','rgba(160,200,255,.9)'),
  });
  return d;
}

function appendSub(acBody,ptfl,ac,acd,sub,sp100,total){
  const subPct=sp100/100,objPct=(acd.pct/100)*subPct,sv=total*(acd.pct/100)*subPct;
  const subId=`sub-${ptfl}-${ac}-${sub}`.replace(/[^a-z0-9-]/gi,'_');
  const sh=document.createElement('div');sh.className='sub-hdr';
  sh.setAttribute('onclick',`toggleSub('${subId}')`);
  sh.innerHTML=`<div class="sub-hdr-l"><span class="sub-toggle open" id="stog-${subId}">▶</span><span class="sub-name">${sub}</span><span class="sub-val" id="svr-${ptfl}-${ac}-${sub}">${fmt(sv)}</span></div><div style="display:flex;align-items:center;gap:8px"><span id="ss-${ptfl}-${ac}-${sub}" class="sub-sum ss-err">∑ 0%</span></div>`;
  acBody.appendChild(sh);
  const subBody=document.createElement('div');subBody.className='sub-body';subBody.id=subId;
  const isAmmType=ptfl==='amm'||ptfl==='ammAz';
  const types=ptfl==='amm'?['ETF']:ptfl==='ammAz'?['Azione']:ac==='BOND'||(acd._type==='bond')?['Obbligazione']:['Azione'];
  types.forEach(type=>{
    // il tipo strumento resta nella chiave di riga ma non viene più mostrato:
    // l'asset class è già indicata nell'intestazione sopra
    const rows=getRows(ptfl);
    const existingKeys=Object.keys(rows).filter(k=>k.startsWith(`${ptfl}|${ac}|${sub}|${type}|`)).sort((a,b)=>parseInt(a.split('|')[4])-parseInt(b.split('|')[4]));
    existingKeys.forEach(k=>subBody.appendChild(buildFundRow(k,objPct,sv,ptfl,ac,sub)));
    const addBtn=document.createElement('button');addBtn.className='add-row-btn';addBtn.innerHTML='+ Aggiungi fondo';
    addBtn.onclick=()=>addFundRow(subBody,ptfl,ac,sub,type,objPct,sv,addBtn);
    subBody.appendChild(addBtn);
  });
  const tr=document.createElement('div');tr.className=`tot-row ${gcls(ptfl)}`;
  tr.innerHTML=gridCells(ptfl,{name:`<span style="color:var(--txt3)">Totale ${sub}</span>`,
    pport:_cr(`tp-${ptfl}-${ac}-${sub}`),ctrv:_cr(`tc-${ptfl}-${ac}-${sub}`),
    costo:_cr(`tcc-${ptfl}-${ac}-${sub}`),costo2:_cr(`tcc2-${ptfl}-${ac}-${sub}`)});
  subBody.appendChild(tr);acBody.appendChild(subBody);
}

function toggleSub(id){const body=document.getElementById(id);const tog=document.getElementById('stog-'+id);if(!body)return;const c=body.classList.toggle('collapsed');if(tog)tog.classList.toggle('open',!c);}

function addFundRow(subBody,ptfl,ac,sub,type,objPct,sv,addBtn){recordUndo();
  const rows=getRows(ptfl);const existing=Object.keys(rows).filter(k=>k.startsWith(`${ptfl}|${ac}|${sub}|${type}|`));
  const used=new Set(existing.map(k=>parseInt(k.split('|')[4])));let next=0;while(used.has(next))next++;
  const k=rk(ptfl,ac,sub,type,next);const newRow=buildFundRow(k,objPct,sv,ptfl,ac,sub);
  subBody.insertBefore(newRow,addBtn);newRow.querySelector('.name-inp-manual')?.focus();
}

function msUrl(isin){
  if(!isin) return null;
  return `https://www.morningstar.com/search?query=${encodeURIComponent(isin)}`;
}

function buildFundRow(k,objPct,sv,ptfl,ac,sub){
  const rd=getRow(k,ptfl);const sk=sid(k);
  const{ctrv,qty}=derive(rd,sv);const ps=rd.pctSettore||0;
  const price=rd.price||null;
  const url=msUrl(rd.isin);
  const total=getTotal(ptfl);const pctPort=total>0&&ctrv>0?fmtP(ctrv/total):'—';
  const ctrvBlocked=total<=0;
  const isDartaLike=ptfl==='darta'||ptfl==='pac';
  if(isDartaLike){
    const auto=getAutoCommission(k,ptfl);
    if(!rd.cg&&auto.cg){rd.cg=auto.cg;setRow(k,ptfl,{cg:auto.cg});}
    if(!rd.sc&&auto.sc){rd.sc=auto.sc;setRow(k,ptfl,{sc:auto.sc});}
  }
  const copBase=ptfl==='pac'?(CD?.portfolios?.darta?.total||total):total;
  const copRate=isDartaLike?getDartaCopertura(copBase):0;
  const commTot=isDartaLike?parsePctStr(rd.cg)+parsePctStr(rd.sc)+copRate:0;
  const rctrv=ctrv>0?Math.round(ctrv):0;
  const costoVal=isDartaLike&&rctrv>0&&commTot>0?rctrv*commTot:0;
  const esCfg=CD?.portfolios?.darta?.pac;
  const isEsSource=ptfl==='darta'&&!!esCfg?.enabled&&esCfg.sourceKey===k;
  const esBadge=isEsSource?`<span class="es-src-badge" title="Fondo di origine Easy Switch">↗ Easy Switch: ${fmt(pacTotal('pac'))}/mese</span>`:'';
  const row=document.createElement('div');row.className=`fund-row ${gcls(ptfl)}${isEsSource?' es-source':''}`;
  const pp=total>0&&rctrv>0?rctrv/total*100:0;
  row.innerHTML=gridCells(ptfl,{
    name:`<div style="display:flex;flex-direction:column;gap:3px;min-width:0;position:relative">
      <input class="name-inp-manual" placeholder="Nome fondo / strumento…" value="${rd.name||''}"
        data-k="${k}" data-ptfl="${ptfl}" data-sub="${sub}"
        oninput="fundTypeahead(this)" onfocus="fundTypeahead(this)" onblur="hideFundTypeahead(this)"
        onchange="onNameManual('${k}','${ptfl}',this.value)"/>
      <input class="isin-inp" placeholder="ISIN…" value="${rd.isin||''}"
        oninput="onIsinInput(this,'${k}','${ptfl}')"
        onchange="onIsinManual('${k}','${ptfl}',this.value)"/>
      ${esBadge}
    </div>`,
    m:`<div style="display:flex;align-items:center;justify-content:center">
      ${rd.isin?`<a class="ms-link" href="${url}" target="_blank" title="Cerca su Morningstar">M</a>`:'<span class="ms-link disabled" title="Inserisci ISIN prima">M</span>'}
    </div>`,
    pport:`<div class="pport-cell"><input class="pset-inp" id="pv-${sk}" type="number" min="0" max="100" step="0.01"
      value="${pp>0?pp.toFixed(2):''}" placeholder="${ctrvBlocked?'—':'0.00'}"
      onchange="onPortSet('${k}','${ptfl}',this.value)"
      ${ctrvBlocked?'disabled title="Inserisci prima il totale del portafoglio"':'title="% del portafoglio totale · Invio: fondo successivo"'}/><button class="fill-rest" onclick="fillRest('${k}','${ptfl}')" title="Assegna il residuo a questo fondo" ${ctrvBlocked?'disabled':''}>⇥</button></div>`,
    ctrv:`<input class="ctrv-inp" id="cv-${sk}" type="number" min="0" step="1"
      value="${rctrv>0?rctrv:''}" placeholder="${ctrvBlocked?'Imposta totale':'CTRV…'}"
      onchange="onCtrvSet('${k}','${ptfl}',this.value)"
      ${ctrvBlocked?'disabled title="Inserisci prima il totale del portafoglio"':''}/>`,
    cg:`<input class="cg-inp" id="cg-${sk}" placeholder="${(ptfl==='amm'||ptfl==='ammAz')?'TER…':'CG…'}"
      value="${rd.cg||''}"
      onchange="onCgManual('${k}','${ptfl}',this.value)"/>`,
    sc:isDartaLike?`<input class="sc-inp" id="sc-${sk}" placeholder="SC…"
      value="${rd.sc||''}"
      onchange="onScManual('${k}','${ptfl}',this.value)"/>`:null,
    cop:isDartaLike?`<span class="computed muted" id="cop-${sk}" style="text-align:right">${fmtCop(copRate||null)}</span>`:null,
    ct:isDartaLike?`<span class="computed muted" id="ct-${sk}" style="text-align:right">${commTot>0?fmtP(commTot):'—'}</span>`:null,
    costo:isDartaLike?`<span class="computed muted" id="co-${sk}" style="text-align:right">${costoVal>0?fmtC(costoVal):'—'}</span>`:null,
    costo2:ptfl==='pac'?`<span class="computed muted" id="ct2-${sk}" style="text-align:right;color:rgba(255,120,40,.9);font-weight:700">${costoVal>0?fmtC(costoVal*(CD?.portfolios?.darta?.pac?.mesi||15)):'—'}</span>`:null,
    psett:`<span class="computed muted psett-val" id="ps-${sk}" title="% all'interno del sotto-settore">${ps>0?(ps*100).toFixed(2)+'%':'—'}</span>`,
    del:`<button class="row-del-btn" onclick="delFundRow('${k}','${ptfl}')" title="Rimuovi riga">×</button>`,
  });
  return row;
}

// ── WEIGHTS PANEL ─────────────────────────────────────────────────────────────
function buildWeights(ptfl,wrap,weights,order,mode){
  const isCtrv=(mode||getPD(ptfl)?._mode||'weights')==='ctrv';
  const res=getResidual(ptfl);wrap.innerHTML='';
  const grid=document.createElement('div');grid.className='w-grid';
  order.forEach(ac=>{
    const d=weights[ac];if(!d)return;
    const isRes=ac===res;
    const clsName=CLSC[ac]||(d._type==='bond'?'bd':d._type==='equity'?'eq':'custom');
    const bl=document.createElement('div');bl.className='w-block';
    const hdr=document.createElement('div');hdr.className=`w-hdr ${clsName}`;
    hdr.innerHTML=`<span>${acLabel(ac,ptfl)}${isRes?` <span style="font-size:9px;opacity:.55">(auto)</span>`:''}</span><div style="display:flex;align-items:center;gap:6px"><span id="wsum-${ptfl}-${ac}" class="sum-chip sum-ok">∑ 100%</span><button onclick="delAc('${ptfl}','${ac}')" title="Rimuovi asset class" style="background:transparent;border:none;color:rgba(255,255,255,.45);font-size:13px;cursor:pointer;padding:0 2px;line-height:1" onmouseover="this.style.color='#ff6b6b'" onmouseout="this.style.color='rgba(255,255,255,.45)'">×</button></div>`;
    bl.appendChild(hdr);
    if(isCtrv)bl.classList.add('w-locked');
    const acRow=document.createElement('div');acRow.className='w-ac-row';
    acRow.innerHTML=`<span class="w-ac-lbl">% portafoglio</span><input type="number" min="0" max="100" step="1" value="${d.pct.toFixed(2)}" id="wac-${ptfl}-${ac}" class="w-ac-inp${isRes?' auto':''}" ${isCtrv?'disabled':''} onchange="setAcPct('${ptfl}','${ac}',this.value)"/><span class="w-suf">%</span>`;
    bl.appendChild(acRow);
    if(ac==='START & GO'){const n=document.createElement('div');n.className='sg-note';n.textContent='↗ Alimenta il '+(TL[PAC_CHILD[ptfl]]||'PAC');bl.appendChild(n);}
    getSubOrder(ptfl,ac).forEach(sub=>{const v=d.subs[sub];if(v===undefined)return;
      const row=document.createElement('div');row.className='w-row';
      row.draggable=true;
      row.innerHTML=`<span class="w-drag" title="Trascina per riordinare">⠿</span><label title="${sub}">${sub}</label><input type="number" min="0" max="100" step="0.01" value="${v}" id="wsub-${ptfl}-${ac}-${sub}" class="w-inp" onchange="setSubPct('${ptfl}','${ac}','${sub}',this.value)"/><span class="w-suf">%</span><button class="w-del-sub" onclick="delSub('${ptfl}','${ac}','${sub}')" title="Rimuovi">×</button>`;
      row.addEventListener('dragstart',e=>{_dSub=sub;_dPtfl=ptfl;_dAc=ac;row.classList.add('is-dragging');e.dataTransfer.effectAllowed='move';});
      row.addEventListener('dragend',()=>{row.classList.remove('is-dragging');document.querySelectorAll('.drag-over-top,.drag-over-bot').forEach(el=>{el.classList.remove('drag-over-top','drag-over-bot');});});
      row.addEventListener('dragover',e=>{if(_dSub===null||_dSub===sub||_dPtfl!==ptfl||_dAc!==ac)return;e.preventDefault();const mid=row.getBoundingClientRect().top+row.getBoundingClientRect().height/2;document.querySelectorAll('.drag-over-top,.drag-over-bot').forEach(el=>{el.classList.remove('drag-over-top','drag-over-bot');});row.classList.add(e.clientY<mid?'drag-over-top':'drag-over-bot');});
      row.addEventListener('dragleave',()=>{row.classList.remove('drag-over-top','drag-over-bot');});
      row.addEventListener('drop',e=>{e.preventDefault();row.classList.remove('drag-over-top','drag-over-bot');if(!_dSub||_dSub===sub||_dPtfl!==ptfl||_dAc!==ac)return;const mid=row.getBoundingClientRect().top+row.getBoundingClientRect().height/2;const before=e.clientY<mid;const acd=getW(ptfl)[ac];if(!acd)return;const keys=getSubOrder(ptfl,ac);const fi=keys.indexOf(_dSub),ti=keys.indexOf(sub);if(fi===-1||ti===-1)return;keys.splice(fi,1);keys.splice(before?ti:ti+1-(fi<ti?1:0),0,_dSub);const ns={};keys.forEach(k=>{ns[k]=acd.subs[k];});acd.subs=ns;acd._subOrder=keys;_dSub=null;buildTab(ptfl);});
      bl.appendChild(row);
    });
    if(ac!=='START & GO'){const addSubBtn=document.createElement('button');addSubBtn.className='w-add-sub';addSubBtn.innerHTML='+ Sotto-settore';addSubBtn.onclick=()=>openAddSub(ptfl,ac);bl.appendChild(addSubBtn);}
    grid.appendChild(bl);
  });
  wrap.appendChild(grid);
  const addAcBtn=document.createElement('button');addAcBtn.className='w-add-ac';addAcBtn.innerHTML='+ Aggiungi asset class personalizzata';addAcBtn.onclick=()=>openAddAc(ptfl);wrap.appendChild(addAcBtn);
  const total=order.reduce((s,ac)=>s+(weights[ac]?.pct||0),0);
  const col=total===100?'var(--ok)':total>100?'var(--err)':'var(--warn)';
  const bar=document.createElement('div');bar.className='gbar-row';
  bar.innerHTML=`<span style="font-size:10px;font-weight:600;color:var(--txt2)">Totale pesi</span><div class="gbar"><div class="gbar-fill" id="gbf-${ptfl}" style="width:${Math.min(100,total)}%;background:${col}"></div></div><span class="gbar-lbl" id="gbv-${ptfl}" style="color:${col}">${total}%</span><span style="font-size:9px;color:var(--txt3);margin-left:4px">residuo auto → ${res||'—'}</span>`;
  wrap.appendChild(bar);
}

function delAc(ptfl,ac){recordUndo();
  const w=getW(ptfl);if(!w[ac])return;
  const pd=getPD(ptfl);
  if(pd?.rows)Object.keys(pd.rows).filter(k=>k.startsWith(`${ptfl}|${ac}|`)).forEach(k=>delete pd.rows[k]);
  delete w[ac];
  setWOrder(ptfl,getWOrder(ptfl).filter(a=>a!==ac));
  if(ac==='START & GO'&&PAC_CHILD[ptfl]){
    const child=PAC_CHILD[ptfl];
    CD['_'+child+'Backup']=JSON.parse(JSON.stringify(CD.portfolios[child]||{}));
    delete CD.portfolios[child];
  }
  buildTab(ptfl);
}
function delSub(ptfl,ac,sub){const w=getW(ptfl);if(!w[ac])return;recordUndo();delete w[ac].subs[sub];if(w[ac]._subOrder)w[ac]._subOrder=w[ac]._subOrder.filter(s=>s!==sub);const pd=getPD(ptfl);if(pd?.rows)Object.keys(pd.rows).filter(k=>k.includes(`|${ac}|${sub}|`)).forEach(k=>delete pd.rows[k]);buildTab(ptfl);}
function setAcPct(ptfl,ac,v){const w=getW(ptfl);if(!w[ac])return;w[ac].pct=Math.min(100,Math.max(0,parseFloat(v)||0));autoAdj(ptfl,ac,w);buildTab(ptfl);if(PAC_CHILD[ptfl])updatePac(ptfl);}
function setSubPct(ptfl,ac,sub,v){const w=getW(ptfl);if(!w[ac])return;w[ac].subs[sub]=Math.min(100,Math.max(0,parseFloat(v)||0));const s=Math.round(Object.values(w[ac].subs).reduce((a,b)=>a+b,0)*100)/100;const ok=Math.abs(s-100)<0.06;const el=document.getElementById(`wsum-${ptfl}-${ac}`);if(el){el.textContent=ok?'✓ 100%':'⚠ '+s.toFixed(2)+'%';el.className=`sum-chip ${ok?'sum-ok':'sum-err'}`;}buildTab(ptfl);}
function autoAdj(ptfl,changed,w){const res=getResidual(ptfl);if(!res||changed===res||!w[res])return;const used=Object.entries(w).filter(([a])=>a!==res).reduce((s,[,d])=>s+(d.pct||0),0);w[res].pct=Math.max(0,100-used);const inp=document.getElementById(`wac-${ptfl}-${res}`);if(inp){inp.value=w[res].pct;inp.classList.add('auto');}const tot=Object.values(w).reduce((s,d)=>s+(d.pct||0),0);const col=tot===100?'var(--ok)':tot>100?'var(--err)':'var(--warn)';const f=document.getElementById(`gbf-${ptfl}`);const vl=document.getElementById(`gbv-${ptfl}`);if(f){f.style.width=Math.min(100,tot)+'%';f.style.background=col;}if(vl){vl.textContent=tot.toFixed(2)+'%';vl.style.color=col;}}

// ── ISIN ──────────────────────────────────────────────────────────────────────
function getDartaEntry(isinKey,nameKey){
  const plan=getPD('darta')?.plan||'pro';
  const src=plan==='pro'?cgData.pro:cgData.chpv;
  return (isinKey&&src.isin[isinKey])||(nameKey&&src.name[nameKey])||null;
}
function _findFundByIsin(isin,ptfl){
  if(ptfl==='darta'||ptfl==='pac'){
    const plan=getPD(ptfl==='pac'?'darta':ptfl)?.plan||'pro';
    const src=plan==='pro'?cgData?.pro:cgData?.chpv;
    if(src?.isin?.[isin])return src.isin[isin];
    const catSrc=plan==='pro'?dartaFundsByCat?.pro:dartaFundsByCat?.chpv;
    for(const funds of Object.values(catSrc||{})){const f=funds.find(f=>f.isin===isin);if(f)return f;}
  }else{
    for(const funds of Object.values(sicavFundsByCat||{})){const f=funds.find(f=>f.isin===isin);if(f)return f;}
  }
  return null;
}
function onIsinInput(inp,k,ptfl){
  const isin=inp.value.trim().toUpperCase();
  if(isin.length<12)return;
  if(ptfl==='amm'||ptfl==='ammAz')return;
  const fund=_findFundByIsin(isin,ptfl);
  if(!fund)return;
  const fundRow=inp.closest('.fund-row');
  const sk=sid(k);
  if(fund.name){const nameInp=fundRow?.querySelector('.name-inp-manual');if(nameInp&&!nameInp.value){nameInp.value=fund.name;setRow(k,ptfl,{name:fund.name});}}
  if(fund.cg){const el=document.getElementById('cg-'+sk);if(el&&!el.value){el.value=fund.cg;setRow(k,ptfl,{cg:fund.cg});}}
  if(fund.sc){const el=document.getElementById('sc-'+sk);if(el&&!el.value){el.value=fund.sc;setRow(k,ptfl,{sc:fund.sc});}}
  if(fund.fee&&!fund.cg){const el=document.getElementById('cg-'+sk);if(el&&!el.value){el.value=fund.fee;setRow(k,ptfl,{cg:fund.fee});}}
}
function onIsinManual(k,ptfl,v){
  const isin=v.trim().toUpperCase();
  setRow(k,ptfl,{isin});
  const fundRow=document.querySelector(`input[onchange*="${k}"][class="isin-inp"]`)?.closest('.fund-row');
  if(fundRow){
    const msWrap=fundRow.children[1];
    if(msWrap){
      if(isin){const url=msUrl(isin);msWrap.innerHTML=`<a class="ms-link" href="${url}" target="_blank" title="Cerca su Morningstar">M</a>`;}
      else{msWrap.innerHTML='<span class="ms-link disabled" title="Inserisci ISIN prima">M</span>';}
    }
    if(ptfl!=='amm'&&ptfl!=='ammAz'){
      const fund=_findFundByIsin(isin,ptfl);
      if(fund){
        const sk=sid(k);
        if(fund.name){const nameInp=fundRow.querySelector('.name-inp-manual');if(nameInp&&!nameInp.value){nameInp.value=fund.name;setRow(k,ptfl,{name:fund.name});}}
        if(fund.cg){const el=document.getElementById('cg-'+sk);if(el&&!el.value){el.value=fund.cg;setRow(k,ptfl,{cg:fund.cg});}}
        if(fund.sc){const el=document.getElementById('sc-'+sk);if(el&&!el.value){el.value=fund.sc;setRow(k,ptfl,{sc:fund.sc});}}
        if(fund.fee&&!fund.cg){const el=document.getElementById('cg-'+sk);if(el&&!el.value){el.value=fund.fee;setRow(k,ptfl,{cg:fund.fee});}}
      }
    }
  }
  if(ptfl==='darta'||ptfl==='pac'){
    const auto=getAutoCommission(k,ptfl);const sk=sid(k);
    const cgEl=document.getElementById('cg-'+sk);const scEl=document.getElementById('sc-'+sk);
    if(cgEl&&!cgEl.value){cgEl.value=auto.cg;setRow(k,ptfl,{cg:auto.cg});}
    if(scEl&&!scEl.value){scEl.value=auto.sc;setRow(k,ptfl,{sc:auto.sc});}
  }
  trigRecalc(k,ptfl);
}
function onNameManual(k,ptfl,v){
  const name=v.trim();
  setRow(k,ptfl,{name});
  if((ptfl==='darta'||ptfl==='pac')&&name){
    const entry=getDartaEntry(null,name.toUpperCase());
    if(entry?.isin){
      const fundRow=document.querySelector(`input[onchange*="${k}"][class="name-inp-manual"]`)?.closest('.fund-row');
      if(fundRow){
        const isinInp=fundRow.querySelector('.isin-inp');
        if(isinInp){
          isinInp.value=entry.isin;setRow(k,ptfl,{isin:entry.isin});
          const msWrap=fundRow.children[1];
          if(msWrap){const url=msUrl(entry.isin);msWrap.innerHTML=`<a class="ms-link" href="${url}" target="_blank" title="Cerca su Morningstar">M</a>`;}
        }
      }
    }
    const auto=getAutoCommission(k,ptfl);const sk=sid(k);
    const cgEl=document.getElementById('cg-'+sk);const scEl=document.getElementById('sc-'+sk);
    if(cgEl){cgEl.value=auto.cg;setRow(k,ptfl,{cg:auto.cg});}
    if(scEl){scEl.value=auto.sc;setRow(k,ptfl,{sc:auto.sc});}
  }
  trigRecalc(k,ptfl);
}
function onCgManual(k,ptfl,v){setRow(k,ptfl,{cg:v.trim()});trigRecalc(k,ptfl);}
function onScManual(k,ptfl,v){setRow(k,ptfl,{sc:v.trim()});trigRecalc(k,ptfl);}

// ── Fund typeahead ─────────────────────────────────────────────────────────────
function fundTypeahead(inp){
  const q=inp.value.trim();
  const k=inp.dataset.k, ptfl=inp.dataset.ptfl, sub=inp.dataset.sub;
  const wrap=inp.parentElement;
  const existing=wrap.querySelector('.fund-dd');
  if(existing)existing.remove();
  if(ptfl==='amm'||ptfl==='ammAz')return;
  let src;
  if(ptfl==='sicav'||ptfl==='sicavPac'){src=sicavFundsByCat||{};}
  else{const plan=(ptfl==='pac'?getPD('darta'):getPD(ptfl))?.plan||'pro';src=(plan==='pro'?dartaFundsByCat.pro:dartaFundsByCat.chpv)||{};}
  const funds=src[sub]||[];
  const qu=q.toUpperCase();
  const matches=q?funds.filter(f=>f.name.toUpperCase().includes(qu)||(f.isin&&f.isin.includes(qu))).slice(0,15):funds.slice(0,20);
  const dd=document.createElement('div');dd.className='fund-dd';
  if(!matches.length){const empty=document.createElement('div');empty.className='fund-dd-item';empty.style.color='var(--txt3)';empty.style.fontStyle='italic';empty.style.cursor='default';empty.textContent='—';dd.appendChild(empty);wrap.appendChild(dd);return;}
  matches.forEach(f=>{
    const item=document.createElement('div');item.className='fund-dd-item';
    const nm=document.createElement('span');nm.className='fund-dd-name';nm.textContent=f.name;item.appendChild(nm);
    if(f.isin){const is=document.createElement('span');is.className='fund-dd-isin';is.textContent=f.isin;item.appendChild(is);}
    item.addEventListener('mousedown',e=>{e.preventDefault();selectFundSuggestion(inp,k,ptfl,f);});
    dd.appendChild(item);
  });
  wrap.appendChild(dd);
}
function hideFundTypeahead(inp){
  setTimeout(()=>{inp.parentElement?.querySelector('.fund-dd')?.remove();},150);
}
function selectFundSuggestion(inp,k,ptfl,fund){
  inp.value=fund.name;setRow(k,ptfl,{name:fund.name});
  const fundRow=inp.closest('.fund-row');
  if(fundRow&&fund.isin){
    const isinInp=fundRow.querySelector('.isin-inp');
    if(isinInp){isinInp.value=fund.isin;setRow(k,ptfl,{isin:fund.isin});}
    const msWrap=fundRow.children[1];
    if(msWrap){msWrap.innerHTML=`<a class="ms-link" href="${msUrl(fund.isin)}" target="_blank" title="Cerca su Morningstar">M</a>`;}
  }
  const sk=sid(k);
  if(fund.cg){const el=document.getElementById('cg-'+sk);if(el){el.value=fund.cg;setRow(k,ptfl,{cg:fund.cg});}}
  if(fund.sc){const el=document.getElementById('sc-'+sk);if(el){el.value=fund.sc;setRow(k,ptfl,{sc:fund.sc});}}
  if(fund.fee){const el=document.getElementById('cg-'+sk);if(el&&!el.value){el.value=fund.fee;setRow(k,ptfl,{cg:fund.fee});}}
  inp.parentElement?.querySelector('.fund-dd')?.remove();
  trigRecalc(k,ptfl);
}
function resyncDartaCommissions(){
  const rows=getRows('darta');
  Object.keys(rows).forEach(k=>{
    const auto=getAutoCommission(k,'darta');
    const sk=sid(k);
    const cgEl=document.getElementById('cg-'+sk);const scEl=document.getElementById('sc-'+sk);
    if(cgEl){cgEl.value=auto.cg;setRow(k,'darta',{cg:auto.cg});}
    if(scEl){scEl.value=auto.sc;setRow(k,'darta',{sc:auto.sc});}
  });
  recalc('darta',getTotal('darta'),getW('darta'),getWOrder('darta'));
  toast('Commissioni aggiornate da Excel','ok');
}
function delFundRow(k,ptfl){recordUndo();
  // Remove from state
  const pd=getPD(ptfl);
  if(pd?.rows) delete pd.rows[k];
  // Remove from DOM
  const sk=sid(k);
  const el=document.getElementById('cv-'+sk)?.closest('.fund-row');
  if(el) el.remove();
  // Recalc
  trigRecalc(k,ptfl);
}
function onPriceManual(k,ptfl,v){
  const price=parseFloat(v)||null;
  setRow(k,ptfl,{price});
  trigRecalc(k,ptfl);
}
function onPctSet(k,ptfl,v){
  const ps=(parseFloat(v)||0)/100;
  const mode=getPD(ptfl)?._mode||'weights';
  if(mode==='ctrv'){
    const parts=k.split('|');const ac=parts[1],sub=parts[2];
    const total=getTotal(ptfl);const w=getW(ptfl);
    const sv=total*((w[ac]?.pct||0)/100)*((w[ac]?.subs?.[sub]||0)/100);
    setRow(k,ptfl,{pctSettore:ps,ctrvAbs:sv*ps,_ctrvAbs:null});
    deriveWeightsFromRows(ptfl);
  } else {
    setRow(k,ptfl,{pctSettore:ps,ctrvAbs:null,_ctrvAbs:null});
    trigRecalc(k,ptfl);
  }
}
// Ultimo stato di allocazione per portafoglio (aggiornato da recalc): {total, allocated}.
const _alloc={};
// ⇥ Assegna il residuo: in modalità CTRV porta il fondo a coprire il non allocato del portafoglio;
// in modalità Pesi completa il 100% del suo sotto-settore.
function fillRest(k,ptfl){
  const mode=getPD(ptfl)?._mode||'weights';
  const parts=k.split('|');const ac=parts[1],sub=parts[2];
  const total=getTotal(ptfl);const w=getW(ptfl);
  const sv=total*((w[ac]?.pct||0)/100)*((w[ac]?.subs?.[sub]||0)/100);
  const cur=Math.round(derive(getRow(k,ptfl),sv).ctrv||0);
  if(mode==='ctrv'){
    const a=_alloc[ptfl];const rem=a?a.total-a.allocated:0;
    if(rem<1){toast('Niente da allocare: il portafoglio è già al 100%','info');return;}
    recordUndo();onCtrvSet(k,ptfl,cur+Math.round(rem));
  }else{
    const others=Object.keys(getRows(ptfl)).filter(x=>x!==k&&x.startsWith(`${ptfl}|${ac}|${sub}|`)).reduce((s,x)=>s+(getRow(x,ptfl).pctSettore||0),0);
    const ps=Math.max(0,1-others);
    if(ps*sv-cur<1){toast('Niente da allocare in questo sotto-settore','info');return;}
    recordUndo();onPctSet(k,ptfl,(ps*100).toFixed(4));
  }
}
// % del portafoglio → controvalore (totale × %), poi stesso percorso del CTRV.
function onPortSet(k,ptfl,v){
  const total=getTotal(ptfl);if(!(total>0))return;
  const pp=parseFloat(String(v).replace(',','.'))||0;
  onCtrvSet(k,ptfl,Math.round(total*pp/100));
}
function onCtrvSet(k,ptfl,v){
  const ctrv=parseFloat(v)||0;
  const mode=getPD(ptfl)?._mode||'weights';
  if(mode==='ctrv'){
    setRow(k,ptfl,{ctrvAbs:ctrv,_ctrvAbs:null});
    deriveWeightsFromRows(ptfl);
  } else {
    const parts=k.split('|');const ac=parts[1],sub=parts[2];
    const total=getTotal(ptfl);const w=getW(ptfl);
    const sv=total*((w[ac]?.pct||0)/100)*((w[ac]?.subs?.[sub]||0)/100);
    setRow(k,ptfl,{pctSettore:sv>0?ctrv/sv:0,ctrvAbs:null,_ctrvAbs:sv>0?null:ctrv});
    if(sv<=0)deriveWeightsFromRows(ptfl);else trigRecalc(k,ptfl);
  }
}
function trigRecalc(k,ptfl){recalc(ptfl,getTotal(ptfl),getW(ptfl),getWOrder(ptfl));}

function setPortfolioMode(ptfl,mode){
  const pd=getPD(ptfl);if(!pd)return;
  const old=pd._mode||'weights';if(old===mode)return;
  recordUndo();
  pd._mode=mode;
  if(mode==='ctrv'){
    // Snapshot ctrvAbs for all existing rows so they become sticky
    const w=getW(ptfl);const rows=getRows(ptfl);const order=getWOrder(ptfl);
    const total=getTotal(ptfl);
    order.forEach(ac=>{const acd=w[ac];if(!acd||ac==='START & GO')return;
      const sv_ac=total*(acd.pct/100);
      Object.entries(acd.subs||{}).forEach(([sub,sp100])=>{
        const sv=sv_ac*(sp100/100);
        const types=(ac==='BOND'||acd._type==='bond')?['Obbligazione']:['Azione'];
        const ks=Object.keys(rows).filter(k=>k.startsWith(`${ptfl}|${ac}|${sub}|`));
        ks.forEach(k=>{const rd=getRow(k,ptfl);const c=sv*(rd.pctSettore||0);if(c>0)setRow(k,ptfl,{ctrvAbs:c});});
      });
    });
  } else {
    // Clear ctrvAbs — pctSettore is already in sync, weights take over
    const rows=getRows(ptfl);Object.keys(rows).forEach(k=>{if(rows[k])rows[k].ctrvAbs=null;});
  }
  buildTab(ptfl);
}

function deriveWeightsFromRows(ptfl){
  const total=getTotal(ptfl);
  const weights=getW(ptfl);const rows=getRows(ptfl);const order=getWOrder(ptfl);
  if(!total){recalc(ptfl,total,weights,order);return;}

  // Step 1: snapshot absolute CTRVs using current sv
  const absC={};
  order.forEach(ac=>{
    const acd=weights[ac];if(!acd||ac==='START & GO')return;
    const sv_ac=total*(acd.pct/100);
    Object.entries(acd.subs||{}).forEach(([sub,sp100])=>{
      const sv=sv_ac*(sp100/100);
      const types=(ac==='BOND'||acd._type==='bond')?['Obbligazione']:['Azione'];
      const ks=Object.keys(rows).filter(k=>k.startsWith(`${ptfl}|${ac}|${sub}|`));
      ks.forEach(k=>{const rd=getRow(k,ptfl);absC[k]=rd.ctrvAbs!=null?rd.ctrvAbs:rd._ctrvAbs!=null?rd._ctrvAbs:sv*(rd.pctSettore||0);});
    });
  });

  // Step 2: derive new weights and re-sync pctSettore
  let nonSGSum=0;
  order.forEach(ac=>{
    const acd=weights[ac];if(!acd||ac==='START & GO')return;
    const subSums={};let acSum=0;
    Object.keys(acd.subs||{}).forEach(sub=>{
      const types=(ac==='BOND'||acd._type==='bond')?['Obbligazione']:['Azione'];
      const ks=Object.keys(rows).filter(k=>k.startsWith(`${ptfl}|${ac}|${sub}|`));
      const s=ks.reduce((sum,k)=>sum+(absC[k]||0),0);
      subSums[sub]=s;acSum+=s;
    });
    const newAcPct=acSum/total*100;
    acd.pct=newAcPct;nonSGSum+=newAcPct;
    const acInp=document.getElementById(`wac-${ptfl}-${ac}`);if(acInp)acInp.value=parseFloat(newAcPct.toFixed(2));

    let sPctSum=0;
    Object.keys(acd.subs).forEach(sub=>{
      const newSP=acSum>0?subSums[sub]/acSum*100:acd.subs[sub];
      acd.subs[sub]=newSP;sPctSum+=newSP;
      const si=document.getElementById(`wsub-${ptfl}-${ac}-${sub}`);if(si)si.value=parseFloat(newSP.toFixed(2));
    });
    const chip=document.getElementById(`wsum-${ptfl}-${ac}`);
    if(chip){const ok=Math.abs(sPctSum-100)<0.06;chip.textContent=ok?'✓ 100%':'⚠ '+sPctSum.toFixed(2)+'%';chip.className=`sum-chip ${ok?'sum-ok':'sum-err'}`;}

    // Re-derive pctSettore using exact subSums to guarantee sum = 1.0 (avoids >100% display)
    Object.entries(acd.subs).forEach(([sub])=>{
      const exactSv=subSums[sub]||0;
      const types=(ac==='BOND'||acd._type==='bond')?['Obbligazione']:['Azione'];
      const ks=Object.keys(rows).filter(k=>k.startsWith(`${ptfl}|${ac}|${sub}|`));
      ks.forEach(k=>{
        const np=exactSv>0?(absC[k]||0)/exactSv:0;
        setRow(k,ptfl,{pctSettore:np,ctrvAbs:absC[k]||0,_ctrvAbs:null});
        const psel=document.getElementById('ps-'+sid(k));if(psel)psel.textContent=np>0?(np*100).toFixed(2)+'%':'—';
      });
    });
  });

  // Step 3: keep START & GO as residual for DARTA
  if(weights['START & GO']){
    const sgPct=Math.max(0,100-nonSGSum);
    weights['START & GO'].pct=sgPct;
    const sgInp=document.getElementById(`wac-${ptfl}-START & GO`);
    if(sgInp){sgInp.value=parseFloat(sgPct.toFixed(2));sgInp.classList.add('auto');}
  }

  // Step 4: update global weight bar
  const tot=order.reduce((s,ac)=>s+(weights[ac]?.pct||0),0);
  const col=Math.abs(tot-100)<0.6?'var(--ok)':tot>100?'var(--err)':'var(--warn)';
  const f=document.getElementById(`gbf-${ptfl}`);const vl=document.getElementById(`gbv-${ptfl}`);
  if(f){f.style.width=Math.min(100,tot)+'%';f.style.background=col;}
  if(vl){vl.textContent=tot.toFixed(2)+'%';vl.style.color=col;}

  recalc(ptfl,total,weights,order);
  if(PAC_CHILD[ptfl])updatePac(ptfl);
}

// ── RECALC ────────────────────────────────────────────────────────────────────
function recalc(ptfl,total,weights,order){
  let grand=0,grandCosto=0,wterNum=0;
  const isDarta=ptfl==='darta';
  const isDartaLike=isDarta||ptfl==='pac';
  const isPac=ptfl==='pac';
  const isSicav=ptfl==='sicav';
  const pacMesi=isPac?(CD?.portfolios?.darta?.pac?.mesi||15):0;
  const copBase=ptfl==='pac'?(CD?.portfolios?.darta?.total||total):total;
  const copRate=isDartaLike?getDartaCopertura(copBase):0;
  // Pre-compute effective total for darta/sicav: sum of actual ctrvAbs values
  let effectiveTotal=total;
  if(isDarta||isSicav){
    let ctrvSum=0;
    (order||Object.keys(weights)).forEach(ac=>{
      const acd=weights[ac];if(!acd||ac==='START & GO')return;
      Object.entries(acd.subs||{}).forEach(([sub,sp100])=>{
        const sv=total*(acd.pct/100)*(sp100/100);
        const rows=getRows(ptfl);
        Object.keys(rows).filter(k=>k.startsWith(`${ptfl}|${ac}|${sub}|`)).forEach(k=>{
          const rd=getRow(k,ptfl);const{ctrv}=derive(rd,sv);ctrvSum+=ctrv>0?Math.round(ctrv):0;
        });
      });
    });
    if(isDarta||isSicav){const sgPct=(weights['START & GO']?.pct||0)/100;ctrvSum+=Math.round(total*sgPct);}
    effectiveTotal=ctrvSum>0?ctrvSum:total;
  }
  (order||Object.keys(weights)).forEach(ac=>{
    const acd=weights[ac];if(!acd||ac==='START & GO')return;
    const acPct=acd.pct/100;
    let acCtrv=0,acCosto=0;
    Object.entries(acd.subs||{}).forEach(([sub,sp100])=>{
      const subPct=sp100/100,objPct=acPct*subPct,sv=total*acPct*subPct;
      let subCtrv=0,subCosto=0,sumPS=0;
      const rows=getRows(ptfl);
      const allKeys=Object.keys(rows).filter(k=>k.startsWith(`${ptfl}|${ac}|${sub}|`));
      const isAmmType=ptfl==='amm'||ptfl==='ammAz';
      let subCtrvPre=0;
      if(isAmmType)allKeys.forEach(k=>{const rd2=getRow(k,ptfl);const{ctrv:c2}=derive(rd2,sv);subCtrvPre+=c2>0?Math.round(c2):0;});
      allKeys.forEach(k=>{
        const rd=getRow(k,ptfl);const{ctrv,qty}=derive(rd,sv);
        const rctrv=ctrv>0?Math.round(ctrv):0;
        const effPS=isAmmType&&subCtrvPre>0?rctrv/subCtrvPre:rd.pctSettore||0;
        sumPS+=effPS;
        subCtrv+=rctrv;grand+=rctrv;
        if(isAmmType&&rctrv>0){const ter=parsePctStr(rd.cg);if(ter>0)wterNum+=rctrv*ter;}
        const sk=sid(k);
        const cv=document.getElementById('cv-'+sk);const qv=document.getElementById('qv-'+sk);const pv=document.getElementById('pv-'+sk);const pe=document.getElementById('pe-'+sk);const psEl=document.getElementById('ps-'+sk);
        if(cv)cv.value=rctrv>0?rctrv:'';if(qv)qv.textContent=qty&&qty>0?fmtQ(qty):'—';
        // % Port. si riferisce al totale impostato del portafoglio (stesso valore usato in inserimento)
        if(pv&&document.activeElement!==pv)pv.value=rctrv>0&&total>0?(rctrv/total*100).toFixed(2):'';
        if(psEl)psEl.textContent=effPS>0?(effPS*100).toFixed(2)+'%':'—';
        if(pe){if(rctrv>0){const d=(rctrv/effectiveTotal)-objPct;const c=Math.abs(d)<0.005?'db-z':d>0?'db-p':'db-n';pe.innerHTML=`<span class="dbadge ${c}">${fmtP(rctrv/effectiveTotal)}</span>`;}else pe.textContent='—';}
        if(isDartaLike){
          const commTot=parsePctStr(rd.cg)+parsePctStr(rd.sc)+copRate;
          const costo=rctrv>0&&commTot>0?rctrv*commTot:0;
          subCosto+=costo;grandCosto+=costo;
          const copEl=document.getElementById('cop-'+sk);const ctEl=document.getElementById('ct-'+sk);const coEl=document.getElementById('co-'+sk);
          if(copEl)copEl.textContent=fmtCop(copRate);
          if(ctEl)ctEl.textContent=commTot>0?fmtP(commTot):'—';
          if(coEl)coEl.textContent=costo>0?fmtC(costo):'—';
          if(isPac){const ct2El=document.getElementById('ct2-'+sk);if(ct2El)ct2El.textContent=costo>0?fmtC(costo*pacMesi):'—';}
        }
      });
      acCtrv+=subCtrv;acCosto+=subCosto;
      const ssi=document.getElementById(`ss-${ptfl}-${ac}-${sub}`);if(ssi){const p=(sumPS*100).toFixed(2);ssi.textContent='∑ '+p+'%';ssi.className=`sub-sum ${Math.abs(sumPS-1)<0.0001?'ss-ok':'ss-err'}`;}
      const tc=document.getElementById(`tc-${ptfl}-${ac}-${sub}`);const tp=document.getElementById(`tp-${ptfl}-${ac}-${sub}`);const svr=document.getElementById(`svr-${ptfl}-${ac}-${sub}`);
      if(tc)tc.textContent=subCtrv>0?fmt(subCtrv):'—';if(tp)tp.textContent=subCtrv>0?fmtP(subCtrv/(total||effectiveTotal)):'—';
      if(svr&&subCtrv>0)svr.textContent=fmt(subCtrv);
      if(isDartaLike){const tcc=document.getElementById(`tcc-${ptfl}-${ac}-${sub}`);if(tcc)tcc.textContent=subCosto>0?fmtC(subCosto):'—';}
      if(isPac){const tcc2=document.getElementById(`tcc2-${ptfl}-${ac}-${sub}`);if(tcc2)tcc2.textContent=subCosto>0?fmtC(subCosto*pacMesi):'—';}
    });
    const actEl=document.getElementById(`act-${ptfl}-${ac}`);const acpEl=document.getElementById(`acp-${ptfl}-${ac}`);
    if(actEl)actEl.textContent=acCtrv>0?fmt(acCtrv):'—';if(acpEl)acpEl.textContent=acCtrv>0?fmtP(acCtrv/(total||effectiveTotal)):'—';
    const achEl=document.getElementById(`ach-${ptfl}-${ac}`);const achpEl=document.getElementById(`achp-${ptfl}-${ac}`);
    if(achEl)achEl.textContent=acCtrv>0?fmt(acCtrv):'—';if(achpEl)achpEl.textContent=acCtrv>0?fmtP(acCtrv/(total||effectiveTotal)):'—';
    if(isDartaLike){
      const actCosto=document.getElementById(`act-costo-${ptfl}-${ac}`);if(actCosto)actCosto.textContent=acCosto>0?fmtC(acCosto):'—';
      const achCosto=document.getElementById(`ach-costo-${ptfl}-${ac}`);if(achCosto)achCosto.textContent=acCosto>0?fmtC(acCosto):'—';
      const wcm=acCtrv>0&&acCosto>0?fmtP(acCosto/acCtrv):'—';
      const actWcm=document.getElementById(`act-wcm-${ptfl}-${ac}`);if(actWcm)actWcm.textContent=wcm;
      const achWcm=document.getElementById(`ach-wcm-${ptfl}-${ac}`);if(achWcm)achWcm.textContent=wcm;
      if(isPac){
        const actC2=document.getElementById(`act-costo2-${ptfl}-${ac}`);if(actC2)actC2.textContent=acCosto>0?fmtC(acCosto*pacMesi):'—';
        const achC2=document.getElementById(`ach-costo2-${ptfl}-${ac}`);if(achC2)achC2.textContent=acCosto>0?fmtC(acCosto*pacMesi):'—';
      }
    }
  });
  const grandDisplay=(isDarta||isSicav)?effectiveTotal:grand;
  {const sgA=(isDarta||isSicav)?Math.round(total*((weights['START & GO']?.pct||0)/100)):0;
   const allocated=grand+sgA;_alloc[ptfl]={total,allocated};
   const ua=document.getElementById('unalloc-'+ptfl);
   if(ua){const gap=total-allocated;
     if(total>0&&grand>0&&Math.abs(gap)>=1){ua.style.display='flex';ua.classList.toggle('over',gap<0);
       ua.innerHTML=gap>0?`<span>Non allocato</span><b>${fmt(gap)}</b><span class="ua-p">${fmtP(gap/total)} del totale</span><span class="ua-hint">Usa ⇥ su un fondo per assegnare il residuo</span>`
         :`<span>Oltre il totale</span><b>${fmt(-gap)}</b><span class="ua-p">${fmtP(-gap/total)} in più</span><span class="ua-hint">Riduci i controvalori o aumenta il totale</span>`;}
     else ua.style.display='none';}}
  const ge=document.getElementById('grand-'+ptfl);if(ge)ge.textContent=grandDisplay>0?fmt(grandDisplay):'—';
  if(ptfl==='amm'||ptfl==='ammAz'){
    const wterPctEl=document.getElementById('wter-pct-'+ptfl);if(wterPctEl)wterPctEl.textContent=grand>0&&wterNum>0?fmtP(wterNum/grand):'—';
    const wterEuroEl=document.getElementById('wter-euro-'+ptfl);if(wterEuroEl)wterEuroEl.textContent=wterNum>0?fmtC(wterNum):'—';
  }
  if(isDartaLike){
    const gc=document.getElementById('grand-costo-'+ptfl);if(gc)gc.textContent=grandCosto>0?fmtC(grandCosto):'—';
    const gcop=document.getElementById('grand-cop-'+ptfl);if(gcop)gcop.textContent=grandCosto>0&&grandDisplay>0?fmtP(grandCosto/grandDisplay):'—';
    if(isPac){const gc2=document.getElementById('grand-costo2-'+ptfl);if(gc2)gc2.textContent=grandCosto>0?fmtC(grandCosto*pacMesi):'—';}
    if(isDarta)updateDartaInfoCards(total);
  }
  refreshCtrvBar(ptfl);
  updateSaveInfo();
}

