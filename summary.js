// ── SUMMARY ───────────────────────────────────────────────────────────────────
// Tipi di portafoglio aggiungibili dal riepilogo (stessi delle card in Modifica).
const PTFL_TYPES=[
  {id:'ammStock',icon:'trend',name:'Amministrato Azioni',desc:'Titoli azionari da Excel'},
  {id:'ammEtf',icon:'bars',name:'Amministrato ETF',desc:'ETF da Excel'},
  {id:'darta',icon:'clipboard',name:'DARTA + Easy Switch',desc:'Incl. START & GO → Easy Switch'},
  {id:'sicav',icon:'bank',name:'SICAV & Fondi',desc:'Fondi e SICAV di terzi'},
  {id:'sicavPac',icon:'coins',name:'PAC',desc:'Piano di accumulo · totale ÷ mesi'},
];
function addPtflMenuHtml(){
  const free=PTFL_TYPES.filter(t=>!CD?.portfolios?.[t.id]);
  if(!free.length)return '';
  return `<div class="add-ptfl"><button class="btn btn-add" onclick="toggleAddPtfl(event)">${ic('plus')} Aggiungi portafoglio</button>
    <div class="add-ptfl-menu">${free.map(t=>`<button class="apm-item" onclick="addPtflFromSummary('${t.id}')"><span class="apm-ic">${ic(t.icon)}</span><span><span class="apm-name">${t.name}</span><span class="apm-desc">${t.desc}</span></span></button>`).join('')}</div></div>`;
}
function toggleAddPtfl(e){
  e.stopPropagation();
  const w=e.currentTarget.closest('.add-ptfl');const open=!w.classList.contains('open');
  document.querySelectorAll('.add-ptfl.open').forEach(x=>x.classList.remove('open'));
  if(open)w.classList.add('open');
}
document.addEventListener('click',e=>{if(!e.target.closest('.add-ptfl'))document.querySelectorAll('.add-ptfl.open').forEach(x=>x.classList.remove('open'));});
// Aggiunge il portafoglio e apre direttamente la sua scheda in Modifica.
function addPtflFromSummary(type){goEdit();if(!CD.portfolios[type])togglePtfl(type);switchTab(type);}

// Tutte le posizioni (nome → controvalore) raccolte dai drill-down monetari.
function collectHoldings(subs){
  const m=new Map();
  const walk=a=>(a||[]).forEach(x=>{if(!x)return;if(x.monetary)x.labels.forEach((l,i)=>m.set(l,(m.get(l)||0)+(x.vals[i]||0)));else walk(x.subs);});
  walk(subs);return [...m].sort((a,b)=>b[1]-a[1]);
}
function sumStat(lbl,val,sub){return `<div class="sum-stat"><div class="sum-stat-l">${lbl}</div><div class="sum-stat-v">${val}</div>${sub?`<div class="sum-stat-s">${sub}</div>`:''}</div>`;}

function renderSummary(client){
  const el=document.getElementById('view-summary');el.innerHTML='';
  const head=document.createElement('div');head.className='summary-header';
  head.innerHTML=`<div><div style="font-size:11px;color:var(--txt3);margin-bottom:2px">Cliente</div><div class="summary-name">${_esc(client.name)}</div></div>
    <div class="summary-actions">
      ${addPtflMenuHtml()}
      <button class="btn btn-ghost" onclick="openExportModal('excel')">${ic('sheet')} Excel</button>
      <button class="btn btn-ghost" onclick="openExportModal('pdf')">${ic('file')} PDF</button>
      <button class="btn btn-ghost" onclick="openExportModal('pptx')">${ic('slides')} PowerPoint</button>
      <button class="btn btn-ghost" onclick="openDup(${client.id},CD.name)">${ic('copy')} Duplica</button>
      <button class="btn btn-del" onclick="openDel(${client.id},CD.name)">${ic('trash')} Elimina</button>
    </div>`;
  el.appendChild(head);
  if(!CD||!Object.keys(CD.portfolios).length){
    const em=document.createElement('div');em.className='dash-empty';
    em.innerHTML=`<div class="empty-ic">${ic('folder')}</div><div style="font-size:15px;font-weight:700;color:var(--txt);margin-bottom:.5rem">Nessun portafoglio configurato</div><div style="margin-bottom:1.25rem;color:var(--txt3)">Aggiungi il primo portafoglio per visualizzare le allocazioni.</div><div style="display:flex;justify-content:center">${addPtflMenuHtml()}</div>`;
    el.appendChild(em);return;
  }

  Object.values(sumCharts).forEach(c=>{try{c.destroy();}catch(e){}});sumCharts={};

  const grid=document.createElement('div');grid.className='summary-grid sum-rows';
  for(const[pt,pdRaw] of Object.entries(CD.portfolios)){
    if(!pdRaw)continue;
    const pd=(pt==='ammStock'||pt==='ammEtf')?ammDataFromRows(pt,pdRaw):pdRaw;
    const label=TL[pt]||pt;
    const total=getTotal(pt);
    const weights=pd.weights||{};
    const order=pd._order||Object.keys(weights);
    // For darta/sicav use actual ctrv sums; otherwise use weight %
    const _useCtrv=(pt==='darta'||pt==='sicav');
    const _acCtrvM=_useCtrv?computeAcCtrvMap(pt,pd,total):null;
    const _grandAcCtrv=_acCtrvM?Object.values(_acCtrvM).reduce((s,v)=>s+v,0):0;
    const _sgVal=(pt==='darta'||pt==='sicav')?Math.round(total*(weights['START & GO']?.pct||0)/100):0;
    const _effectiveTotal=_useCtrv?((_grandAcCtrv+_sgVal)>0?(_grandAcCtrv+_sgVal):total):total;
    const topL=[],topV=[],topC=[],topSubs=[];let ci=0;
    const r2=x=>Math.round(x*100)/100;
    order.forEach(ac=>{
      const acd=weights[ac];if(!acd)return;
      if(ac==='START & GO'){
        if(pt==='darta'||pt==='sicav'){topL.push('START & GO');topV.push(_effectiveTotal>0?r2(_sgVal/_effectiveTotal*100):r2(acd.pct||0));topC.push('#c03060');topSubs.push(null);ci++;}
        return;
      }
      const _cv=_acCtrvM?(_acCtrvM[ac]||0):null;
      const _pctV=_cv!=null&&_effectiveTotal>0?_cv/_effectiveTotal*100:acd.pct||0;
      topL.push(ac);topV.push(r2(_pctV));topC.push(acColor(ac,ci));
      const subL=[],subV=[],subC=[],subFunds=[];let si=0;
      Object.entries(acd.subs||{}).forEach(([s,sp])=>{
        subL.push(s);subV.push(sp);subC.push(SUB_COLORS[si%SUB_COLORS.length]);
        const sv=total*(acd.pct||0)/100*(sp||0)/100;
        subFunds.push((pt==='ammStock'||pt==='ammEtf')?ammFundDrill(pt,pdRaw,ac,s):fundDrill(pt,pdRaw,ac,s,sv));
        si++;
      });
      topSubs.push({labels:subL,vals:subV,colors:subC,subs:subFunds,name:ac,base:_effectiveTotal*_pctV/100});ci++;
    });
    if(pt==='ammStock'&&topSubs[0]){
      // Amministrato Azioni: show sectors directly (the single Equity class is implicit)
      const s=topSubs[0];
      topL.length=topV.length=topC.length=topSubs.length=0;
      s.labels.forEach((lbl,i)=>{topL.push(lbl);topV.push(s.vals[i]);topC.push(s.colors[i]);topSubs.push(s.subs?s.subs[i]:null);});
    }

    const _child=PAC_CHILD[pt];
    const onSpecial=_child?{'START & GO':()=>{const el=document.getElementById('sum-pac-reveal-'+_child);if(el){el.style.display='block';el.scrollIntoView({behavior:'smooth',block:'nearest'});el.style.outline='2px solid #c03060';setTimeout(()=>el.style.outline='',1200);}}}:null;

    const card=document.createElement('div');card.className='sum-card';
    if(isPacType(pt)){card.id='sum-pac-reveal-'+pt;if(!(pt==='pac'&&CD.portfolios?.darta?.pac?.enabled))card.style.display='none';}
    const ch=document.createElement('div');ch.className='sum-card-head';
    const editTarget=isPacType(pt)?PAC_PARENT[pt]:pt;
    // Totale impostato ≠ somma dei fondi (+ START & GO): segnala la differenza
    const uaGap=_useCtrv&&_grandAcCtrv>0?total-(_grandAcCtrv+_sgVal):0;
    const uaChip=Math.abs(uaGap)>=1?`<span class="sum-unalloc${uaGap<0?' over':''}" title="Totale impostato ${fmt(total)}">${uaGap>0?'Non allocato '+fmt(uaGap):'Oltre il totale di '+fmt(-uaGap)}</span>`:'';
    ch.innerHTML=`<div class="sch-l"><span class="sum-card-title">${label}</span><button class="sum-edit-btn" onclick="goEditPtfl('${editTarget}')">${ic('edit')} Modifica</button></div><div class="sch-r">${uaChip}<span class="sum-card-total" data-cu="${_effectiveTotal}">${fmt(_effectiveTotal)}</span></div>`;
    card.appendChild(ch);

    const body=document.createElement('div');body.className='sum-card-body';
    const cWrap=document.createElement('div');cWrap.className='sum-chart-wrap';
    const canvas=document.createElement('canvas');const cid=`sc-${pt}-${AID}`;canvas.id=cid;
    cWrap.innerHTML=`<div class="sum-chart-center"><div class="sum-chart-center-lbl" id="scl-${cid}">Alloc.</div></div>`;
    cWrap.appendChild(canvas);body.appendChild(cWrap);

    const legWrap=document.createElement('div');legWrap.className='sum-leg-wrap';
    const backBtn=document.createElement('button');backBtn.className='sum-back-btn';backBtn.id=`sback-${cid}`;backBtn.textContent='← Indietro';backBtn.onclick=()=>sumDrillBack(cid);
    const legEl=document.createElement('div');legEl.className='sum-legend';legEl.id=`sleg-${cid}`;
    legWrap.appendChild(backBtn);legWrap.appendChild(legEl);body.appendChild(legWrap);

    // Pannello dettagli: KPI + principali posizioni
    const side=document.createElement('div');side.className='sum-side';
    const hold=collectHoldings(topSubs);
    // Easy Switch salva total=0 (la rata deriva da DARTA): si passa la rata reale per il calcolo costi.
    const costPct=portfolioCostPct(pt,pt==='pac'?{...pdRaw,total}:pdRaw,{darta:{data:CD.portfolios.darta}});
    const stats=[];
    if(pt==='sicavPac'){
      const ms=pdRaw.mesi||0;
      stats.push(sumStat('Totale piano',fmt(pdRaw.total)),sumStat('Rata mensile',fmt(total),ms?`${ms} mesi`:'Durata non impostata'));
    }else if(pt==='pac'){
      // getTotal('pac') è già la rata mensile (START & GO ÷ mesi)
      const ms=pacMesi('pac');
      stats.push(sumStat('Rata mensile',fmt(total),ms?`${ms} mesi`:'Durata non impostata'));
      const esC=CD.portfolios.darta?.pac||{};
      const srcLbl=esC.sourceKey==='__SG__'?'START & GO':(CD.portfolios.darta?.rows?.[esC.sourceKey]?.name||'DARTA');
      if(ms)stats.push(sumStat('Totale switch',fmt(total*ms),'da '+_esc(srcLbl)));
    }
    if(costPct!=null){
      // Easy Switch: il costo a regime si applica a tutto l'importo switchato (rata × mesi).
      const msPac=pt==='pac'?pacMesi('pac'):0;
      const base=pt==='pac'&&msPac?total*msPac:_effectiveTotal;
      stats.push(sumStat('Costo annuo',fmtP(costPct),`≈ ${fmt(costPct*base)} / anno${pt==='pac'&&msPac?' a regime':''}`));
    }
    else if(pt==='ammStock')stats.push(sumStat('Costo annuo','—','Nessun costo ricorrente'));
    stats.push(sumStat('Strumenti',hold.length?hold.length.toLocaleString('it-IT'):'—',`${topL.length} ${pt==='ammStock'?'settori':'asset class'}`));
    // START & GO è un'asset class senza fondi: compare comunque tra le posizioni con il suo peso.
    // Fondo (o START & GO) da cui parte l'Easy Switch: etichetta nelle posizioni DARTA
    const esCfg=pt==='darta'&&pdRaw.pac?.enabled?pdRaw.pac:null;
    const esSrc=esCfg?(esCfg.sourceKey==='__SG__'?'START & GO':(pdRaw.rows?.[esCfg.sourceKey]?.name||null)):null;
    const topHold=_sgVal>0?[...hold,['START & GO',_sgVal]].sort((a,b)=>b[1]-a[1]):hold;
    const top=topHold.slice(0,5),denom=_effectiveTotal||topHold.reduce((a,h)=>a+h[1],0)||1;
    side.innerHTML=`<div class="sum-stats">${stats.join('')}</div>`+(top.length?`<div class="sum-top"><div class="sum-top-h">Principali posizioni</div>${top.map(([n,v])=>{const w=v/denom*100;return `<div class="sum-top-row" title="${_esc(n)} · ${fmt(v)}"><span class="sum-top-n">${n==='START & GO'?'<i class="sum-top-sg"></i>':''}<span class="sum-top-t">${_esc(n)}</span>${n===esSrc?'<span class="sum-es-tag" title="Fondo di origine dell\'Easy Switch">↗ Easy Switch</span>':''}</span><span class="sum-top-p">${w.toFixed(1)}%</span><div class="sum-top-bar"><i style="width:${Math.min(100,w).toFixed(1)}%${n==='START & GO'?';background:#c03060':''}"></i></div></div>`;}).join('')}${topHold.length>5?`<div class="sum-top-more">+ altre ${topHold.length-5} posizioni</div>`:''}</div>`:'');
    body.appendChild(side);
    card.appendChild(body);grid.appendChild(card);

    setTimeout(()=>{
      if(!topV.length)return;
      buildSumChart(cid,canvas,topL,topV,topC,topSubs,'top',onSpecial,null,false,_effectiveTotal);
    },60);
  }
  el.appendChild(grid);
  countUps(el);
}

function buildSumChart(cid,canvas,labels,vals,colors,subs,level,onSpecial,name,monetary,base){
  if(sumCharts[cid]){try{sumCharts[cid].destroy();}catch(e){}}
  const chart=new Chart(canvas,{
    type:'doughnut',
    data:{labels,datasets:[{data:vals,backgroundColor:colors,borderWidth:2,borderColor:'#13223f',hoverBorderWidth:3,hoverOffset:8}]},
    options:{responsive:true,maintainAspectRatio:true,cutout:'62%',layout:{padding:8},animation:donutAnim(level),
      plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>`${ctx.label}: ${monetary?fmt(ctx.parsed):(+ctx.parsed).toFixed(2)+'%'}`}}},
      onClick:(e,els)=>{
        if(!els.length)return;
        const idx=els[0].index;
        if(onSpecial&&onSpecial[labels[idx]]){onSpecial[labels[idx]]();return;}
        if(subs)sumDrill(cid,subs[idx]);
      }
    }
  });
  sumCharts[cid]=chart;
  canvas._cur={labels,vals,colors,subs,level,onSpecial,name,monetary,base};
  if(level==='top')canvas._topData=canvas._cur;
  renderSumLegend(cid,labels,vals,colors,subs,level,onSpecial,monetary,base);
  const lbl=document.getElementById(`scl-${cid}`);if(lbl)lbl.textContent=name||'Alloc.';
}

// Drill one level down on a summary donut, remembering the path for ← Indietro.
function sumDrill(cid,sub){
  if(!sub||!sub.vals||!sub.vals.length)return;
  const canvas=document.getElementById(cid);if(!canvas||!canvas._cur)return;
  const back=document.getElementById(`sback-${cid}`);
  if(back){(back._stack=back._stack||[]).push(canvas._cur);back.style.display='block';}
  buildSumChart(cid,canvas,sub.labels,sub.vals,sub.colors,sub.subs||null,'sub',null,sub.name,!!sub.monetary,sub.base);
}

function renderSumLegend(cid,labels,vals,colors,subs,level,onSpecial,monetary,base){
  const leg=document.getElementById(`sleg-${cid}`);if(!leg)return;
  leg.innerHTML='';
  const total=vals.reduce((a,b)=>a+b,0)||1;
  labels.forEach((lbl,i)=>{
    const pct=((vals[i]/total)*100).toFixed(2);
    const eur=monetary?vals[i]:(base>0?base*vals[i]/total:null);
    const isSpecial=level==='top'&&onSpecial&&onSpecial[lbl];
    const canDrill=!isSpecial&&subs&&subs[i]&&subs[i].vals&&subs[i].vals.length>0;
    const item=document.createElement('div');item.className='sum-legend-item';
    item.innerHTML=`<div class="sum-legend-dot" style="background:${colors[i]}"></div><span class="sum-legend-name" title="${_esc(lbl)}">${_esc(lbl)}</span><span class="sum-legend-eur">${eur!=null?fmt(eur):''}</span><span class="sum-legend-pct">${pct}%</span><span class="sum-legend-go">${(canDrill||isSpecial)?'›':''}</span>`;
    if(isSpecial){
      item.style.cursor='pointer';
      item.onclick=onSpecial[lbl];
    } else if(canDrill){
      item.style.cursor='pointer';
      item.onclick=()=>sumDrill(cid,subs[i]);
    }
    leg.appendChild(item);
  });
}

function sumDrillBack(cid){
  const canvas=document.getElementById(cid);
  const back=document.getElementById(`sback-${cid}`);
  const parent=(back&&back._stack&&back._stack.pop())||(canvas&&canvas._topData);
  if(!parent)return;
  buildSumChart(cid,canvas,parent.labels,parent.vals,parent.colors,parent.subs,parent.level,parent.onSpecial,parent.name,parent.monetary,parent.base);
  if(back&&(!back._stack||!back._stack.length))back.style.display='none';
}

function goEdit(){showView('edit');document.getElementById('nav-crumb').innerHTML=consCrumb()+`<span>${_esc(CD?.name)}</span> › Modifica`;renderEdit();}
// Open the editor directly on one portfolio's tab (click on a summary card head).
function goEditPtfl(pt){goEdit();switchTab(pt);}
function goSummary(){const client=clients.find(c=>c.id===AID);if(!client)return;renderSummary(client);showView('summary');document.getElementById('nav-crumb').innerHTML=consCrumb()+`<span>${_esc(CD?.name)}</span>`;}

