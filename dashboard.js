// ── VIEWS ─────────────────────────────────────────────────────────────────────
function showView(v){document.querySelectorAll('.view').forEach(el=>el.classList.remove('active'));document.getElementById('view-'+v)?.classList.add('active');closeSb();}
function toggleSb(){document.body.classList.toggle('sb-open');}
function closeSb(){document.body.classList.remove('sb-open');}

// ── CLIENT LIST ───────────────────────────────────────────────────────────────
function clientColor(name){
  const hues=[200,145,25,340,270,180,55,310];
  let h=0;for(let i=0;i<name.length;i++)h=(h*31+name.charCodeAt(i))%hues.length;
  return `hsl(${hues[h]},50%,36%)`;
}
async function loadClients(){const r=await fetch('/api/clients');clients=await r.json();renderCliList(clients);}
function filterCli(q){renderCliList(clients.filter(c=>c.name.toLowerCase().includes(q.toLowerCase())));}
function renderCliList(list){
  list=list.filter(c=>c.consultant_id===CONS);
  const el=document.getElementById('cli-list');
  if(!list.length){el.innerHTML='<div class="sb-empty">Nessun cliente.<br><strong>+ Nuovo</strong> per iniziare.</div>';return;}
  el.innerHTML=list.map(c=>{
    const ini=c.name.split(' ').map(w=>w[0]).join('').toUpperCase().slice(0,2);
    const pts=c.portfolios.map(p=>TL[p]||p).join(' · ')||'—';
    return `<div class="cli${c.id===AID?' active':''}" onclick="selectClient(${c.id})"><div class="cli-av" style="background:${clientColor(c.name)}">${_esc(ini)}</div><div class="cli-info"><div class="cli-name">${_esc(c.name)}</div><div class="cli-sub">${pts}</div></div></div>`;
  }).join('');
}
// Portfolio cache: render instantly on click, refresh from the server in background.
const _ptfCache={};
async function selectClient(id){
  AID=id;const client=clients.find(c=>c.id===id);
  const apply=saved=>{
    CD={id,name:client.name,portfolios:{}};
    for(const[pt,pi] of Object.entries(saved))CD.portfolios[pt]=pi.data;
    migrateDartaEasySwitch();
    migrateSicavPacStandalone();
    renderCliList(clients);renderSummary(client);showView('summary');
    document.getElementById('nav-crumb').innerHTML=consCrumb()+`<span>${_esc(client.name)}</span>`;
    startAutoSave();
  };
  const cached=_ptfCache[id];
  if(cached){
    apply(JSON.parse(cached));
    fetch(`/api/clients/${id}/portfolios`).then(r=>r.json()).then(saved=>{
      const s=JSON.stringify(saved);
      if(s===_ptfCache[id])return;
      _ptfCache[id]=s;
      // re-apply only if the user is still on this client's summary with no local edits
      if(AID===id&&!isDirty()&&document.getElementById('view-summary')?.classList.contains('active'))apply(saved);
    }).catch(()=>{});
    return;
  }
  const r=await fetch(`/api/clients/${id}/portfolios`);const saved=await r.json();
  _ptfCache[id]=JSON.stringify(saved);
  apply(saved);
}

// ── DASHBOARD ─────────────────────────────────────────────────────────────────
// Total AUM of a client from the /api/dashboard payload (same types as the card).
function clientAum(saved){
  let t=0;
  for(const[pt,pi] of Object.entries(saved||{})){
    if(isPacType(pt))continue;
    const pd=dashPd(pt,pi);
    if(pd)t+=Number(pd.total)||0;
  }
  return t;
}

// Weighted annual cost of a portfolio as a fraction of its value — same math as
// recalc(): CG+SC+copertura for DARTA/Easy Switch, CG for SICAV/PAC, TER for gli
// Amministrato. Returns null when there's nothing to compute.
function portfolioCostPct(pt,pd,saved){
  if(!pd)return null;
  if(pt==='ammStock')return null; // niente costi ricorrenti sulle azioni
  if(pt==='ammEtf'){
    const rows=pd.rows||[];let num=0,den=0;
    rows.forEach(r=>{const c=Number(r.ctrv)||0;if(c>0){den+=c;const t=Number(r.ter);if(t>0)num+=c*t;}});
    return den>0&&num>0?num/den:null;
  }
  const total=Number(pd.total)||0;if(!(total>0))return null;
  const isDartaLike=pt==='darta'||pt==='pac';
  const copRate=isDartaLike?getDartaCopertura(pt==='pac'?(Number(saved?.darta?.data?.total)||total):total):0;
  const weights=pd.weights||{},rows=pd.rows||{};
  let ctrvSum=0,costo=0;
  Object.entries(weights).forEach(([ac,acd])=>{
    if(!acd||ac==='START & GO')return;
    Object.entries(acd.subs||{}).forEach(([sub,sp100])=>{
      const sv=total*(acd.pct||0)/100*(sp100||0)/100;
      Object.keys(rows).filter(k=>k.startsWith(`${pt}|${ac}|${sub}|`)).forEach(k=>{
        const rd=rows[k]||{};
        const c=rd.ctrvAbs!=null?rd.ctrvAbs:sv*(rd.pctSettore||0);
        const rc=c>0?Math.round(c):0;if(!rc)return;
        ctrvSum+=rc;
        const comm=parsePctStr(rd.cg)+(isDartaLike?parsePctStr(rd.sc)+copRate:0);
        if(comm>0)costo+=rc*comm;
      });
    });
  });
  let eff=ctrvSum;
  if(pt==='darta'||pt==='sicav')eff+=Math.round(total*((weights['START & GO']?.pct||0)/100));
  if(!(eff>0))eff=total;
  return costo>0?costo/eff:null;
}

const DASH_SORTS=[['name','Nome'],['aum','AUM'],['updated','Aggiornato']];
function setDashSort(s){try{localStorage.setItem('pm_dash_sort',s);}catch(e){}renderDash();}

// Dashboard cache: paint the last known state immediately, then refresh behind it.
let _dashCache=null;
async function renderDash(){
  if(_dashCache){
    _buildDash(_dashCache);
    fetch('/api/dashboard').then(r=>r.json()).then(all=>{
      const changed=JSON.stringify(all)!==JSON.stringify(_dashCache);
      _dashCache=all;
      clients=all.map(({portfolios_data,...c})=>c);
      renderCliList(clients);
      if(changed&&document.getElementById('view-dash')?.classList.contains('active'))_buildDash(all);
      if(changed&&document.getElementById('view-cons')?.classList.contains('active'))renderConsultants();
    }).catch(()=>{});
    return;
  }
  // single round-trip: clients + all portfolio data together
  let all=clients.map(c=>({...c,portfolios_data:{}}));
  try{
    const r=await fetch('/api/dashboard');
    all=await r.json();
    _dashCache=all;
    clients=all.map(({portfolios_data,...c})=>c);
    renderCliList(clients);
  }catch(e){}
  _buildDash(all);
}
function _buildDash(all){
  all=all.filter(c=>c.consultant_id===CONS);
  const el=document.getElementById('view-dash');
  if(!all.length){
    el.innerHTML=`<div class="dash-empty"><div class="empty-ic">${ic('pie')}</div><div style="font-size:17px;font-weight:700;color:var(--txt);margin-bottom:.5rem">Benvenuto nel Portfolio Manager</div><div style="margin-bottom:1.5rem">Crea il tuo primo cliente per iniziare.</div><button class="btn btn-primary" onclick="openM('add')">+ Crea primo cliente</button></div>`;
    return;
  }
  let sort='name';try{sort=localStorage.getItem('pm_dash_sort')||'name';}catch(e){}
  if(sort==='aum')all.sort((a,b)=>clientAum(b.portfolios_data)-clientAum(a.portfolios_data));
  else if(sort==='updated')all.sort((a,b)=>new Date(b.updated_at)-new Date(a.updated_at));
  else all.sort((a,b)=>a.name.localeCompare(b.name,'it'));
  const sortBtns=DASH_SORTS.map(([k,l])=>`<button class="ds-btn${k===sort?' on':''}" onclick="setDashSort('${k}')">${l}</button>`).join('');
  el.innerHTML=`<div class="dash-sec"><span>Clienti di ${_esc(consName())} — ${all.length}</span><div class="dash-sort">${sortBtns}</div></div><div class="dash-grid" id="dash-grid"></div>`;
  Object.values(dashCharts).forEach(c=>{try{c.destroy();}catch(e){}});dashCharts={};
  all.forEach(c=>buildDashCard(c));
}

function buildDashCard(c){
  const grid=document.getElementById('dash-grid');if(!grid)return;
  const pts=c.portfolios.map(p=>TL[p]||p).join(' · ')||'—';
  const date=new Date(c.updated_at).toLocaleDateString('it-IT');
  const card=document.createElement('div');card.className='dash-card';
  card.innerHTML=`
    <div class="dash-card-top">
      <div class="dash-card-info">
        <div class="dash-card-name">${_esc(c.name)}</div>
        <div class="dash-card-ptfls">${pts}</div>
        <div class="dash-card-aum" id="daum-${c.id}">—</div>
        <div class="dash-card-date">Aggiornato: ${date}</div>
      </div>
    </div>
    <div class="dash-chart-area" id="dca-${c.id}">
      <div class="dash-chart-col">
        <div class="dash-chart-col-title">Per Portafoglio</div>
        <div class="dash-chart-wrap">
          <canvas id="dc-${c.id}"></canvas>
          <div class="dash-chart-center"><div class="dash-chart-center-label" id="dcl-${c.id}">—</div></div>
        </div>
        <button class="dash-back-btn" id="dback-${c.id}" onclick="drillBack(${c.id})">← Indietro</button>
        <div class="dash-legend" id="dleg-${c.id}" style="width:100%"></div>
      </div>
      <div class="dash-chart-divider"></div>
      <div class="dash-chart-col">
        <div class="dash-chart-col-title">Asset Class</div>
        <div class="dash-chart-wrap">
          <canvas id="dc2-${c.id}"></canvas>
          <div class="dash-chart-center"><div class="dash-chart-center-label" id="dcl2-${c.id}">—</div></div>
        </div>
        <button class="dash-back-btn" id="dback2-${c.id}" onclick="drillBack2(${c.id})">← Indietro</button>
        <div class="dash-legend" id="dleg2-${c.id}" style="width:100%"></div>
      </div>
    </div>
    <div class="dash-cost-row" id="dcost-${c.id}"></div>`;
  // whole card navigates; the chart area keeps its own drill-down interactions
  card.onclick=()=>selectClient(c.id);
  card.querySelector('.dash-chart-area').onclick=e=>e.stopPropagation();
  grid.appendChild(card);

  {
    const saved=c.portfolios_data||{};
    let totalAum=0;
    // TOP LEVEL: split by portfolio type (Amministrato, DARTA, PAC, SICAV)
    const ptflColors={amm:'#2eaa5f',ammStock:'#7c5cbf',ammEtf:'#1db8b0',darta:'#2e72d2',pac:'#c03060',sicav:'#d97f00',sicavPac:'#c03060'};
    const topLabels=[],topVals=[],topColors=[],topSubs=[];
    for(const[pt,pi] of Object.entries(saved)){
      const pd=dashPd(pt,pi);if(!pd)continue;
      if(isPacType(pt)) continue; // Easy Switch/PAC are inside the parent's drill
      const ptTotal=pd.total||0;
      totalAum+=ptTotal;
      if(ptTotal<=0) continue;
      topLabels.push(TL[pt]||pt);
      topVals.push(Math.round(ptTotal));
      topColors.push(ptflColors[pt]||AC_COLORS[Object.keys(ptflColors).indexOf(pt)%AC_COLORS.length]);
      // DRILL LEVEL 1: asset classes within this portfolio
      const w=pd.weights||{};
      const order=pd._order||Object.keys(w);
      const acL=[],acV=[],acC=[],acSubs=[];let ci=0;
      order.forEach(ac=>{
        const acd=w[ac];if(!acd)return;
        const v=ptTotal*(acd.pct||0)/100;
        if(v<=0)return;
        acL.push(ac);acV.push(Math.round(v));acC.push(acColor(ac,ci));
        if(ac==='START & GO'){
          // Drill into the linked Easy Switch / PAC allocation
          const child=PAC_CHILD[pt];
          const pacPd=child?saved[child]?.data:null;
          if(pacPd){
            const pacW=pacPd.weights||{};
            const pacL=[],pacV=[],pacC=[];let pi2=0;
            Object.entries(pacW).forEach(([pac_ac,pac_acd])=>{
              const pv=v*(pac_acd.pct||0)/100;
              if(pv>0){pacL.push(pac_ac);pacV.push(Math.round(pv));pacC.push(SUB_COLORS[pi2%SUB_COLORS.length]);pi2++;}
            });
            acSubs.push({labels:pacL,vals:pacV,colors:pacC,name:'START & GO → '+(TL[child]||'PAC')});
          } else {
            acSubs.push({labels:[],vals:[],colors:[],name:'START & GO'});
          }
        } else {
          const subL=[],subV=[],subC=[],subFunds=[];let si=0;
          Object.entries(acd.subs||{}).forEach(([sub,sp100])=>{
            const sv=v*(sp100||0)/100;
            subL.push(sub);subV.push(Math.round(sv));subC.push(SUB_COLORS[si%SUB_COLORS.length]);
            subFunds.push((pt==='ammStock'||pt==='ammEtf')?ammFundDrill(pt,pi&&pi.data,ac,sub):fundDrill(pt,pd,ac,sub,sv));
            si++;
          });
          acSubs.push({labels:subL,vals:subV,colors:subC,subs:subFunds,name:ac});
        }
        ci++;
      });
      if(pt==='ammStock'&&acSubs[0]){
        // Amministrato Azioni: drill straight into sectors (skip the single Equity class)
        const s=acSubs[0];
        topSubs.push({labels:s.labels,vals:s.vals,colors:s.colors,subs:s.subs,name:TL[pt]||pt});
      } else {
        topSubs.push({labels:acL,vals:acV,colors:acC,subs:acSubs,name:TL[pt]||pt});
      }
    }
    const aumEl=document.getElementById(`daum-${c.id}`);
    if(aumEl)countUp(aumEl,totalAum);

    // Left chart: by portfolio type
    const canvas=document.getElementById(`dc-${c.id}`);
    if(canvas&&topVals.length) buildDrillChart(c.id,canvas,topLabels,topVals,topColors,topSubs,'top');

    // Right chart: by asset class (combined across all portfolios)
    // Each AC shows % of total AUM; clicking shows which portfolio contributes what
    const acMap={}; // ac -> {total, byPortfolio: {ptfl: val}}
    for(const[pt,pi] of Object.entries(saved)){
      if(isPacType(pt)) continue;
      const pd=dashPd(pt,pi); if(!pd) continue;
      const ptTotal=pd.total||0;
      const w=pd.weights||{};
      const order=pd._order||Object.keys(w);
      order.forEach(ac=>{
        const acd=w[ac]; if(!acd||ac==='START & GO') return;
        const v=ptTotal*(acd.pct||0)/100;
        if(!acMap[ac]) acMap[ac]={total:0,byPortfolio:{}};
        acMap[ac].total+=v;
        acMap[ac].byPortfolio[pt]=(acMap[ac].byPortfolio[pt]||0)+v;
      });
    }
    const ac2L=[],ac2V=[],ac2C=[],ac2Subs=[];let ci2=0;
    Object.entries(acMap).forEach(([ac,data])=>{
      if(data.total<=0) return;
      ac2L.push(ac);
      ac2V.push(Math.round(data.total));
      ac2C.push(acColor(ac,ci2));
      // drill: shows per-portfolio breakdown of this AC
      const subL=[],subV=[],subC=[];let si=0;
      Object.entries(data.byPortfolio).forEach(([pt,v])=>{
        if(v>0){subL.push(TL[pt]||pt);subV.push(Math.round(v));subC.push(AC_COLORS[(ci2+si+3)%AC_COLORS.length]);si++;}
      });
      ac2Subs.push({labels:subL,vals:subV,colors:subC,name:ac});
      ci2++;
    });
    const canvas2=document.getElementById(`dc2-${c.id}`);
    if(canvas2&&ac2V.length) buildDrillChart2(c.id,canvas2,ac2L,ac2V,ac2C,ac2Subs,'top');

    // Cost row: annual cost % per portfolio type (e.g. DARTA 1.52%)
    const CL={amm:'Amm. ETF',ammAz:'Amm. Azioni',ammEtf:'Amm. ETF',darta:'DARTA',sicav:'SICAV',sicavPac:'PAC'};
    const costParts=[];
    for(const[pt,pi] of Object.entries(saved)){
      if(isPacType(pt))continue;
      const cp=portfolioCostPct(pt,pi&&pi.data,saved);
      if(cp!=null)costParts.push(`<span class="dash-cost-item">${CL[pt]||TL[pt]||pt} <b>${fmtP(cp)}</b></span>`);
    }
    const costEl=document.getElementById(`dcost-${c.id}`);
    if(costEl&&costParts.length){costEl.style.display='flex';costEl.innerHTML=`<span class="dash-cost-lbl">Costi</span>`+costParts.join('');}
  }
}

// Fund-level drill: the holdings inside one sub-sector (name + controvalore).
function fundDrill(pt,pd,ac,sub,sv){
  const rows=pd.rows||{};const L=[],V=[],C=[];let fi=0;
  Object.keys(rows).filter(k=>k.startsWith(`${pt}|${ac}|${sub}|`)).forEach(k=>{
    const rd=rows[k]||{};
    const c=rd.ctrvAbs!=null?rd.ctrvAbs:sv*(rd.pctSettore||0);
    if(!(c>0))return;
    L.push(rd.name||rd.isin||'—');V.push(Math.round(c));C.push(SUB_COLORS[fi%SUB_COLORS.length]);fi++;
  });
  return L.length?{labels:L,vals:V,colors:C,name:sub,monetary:true}:null;
}
// Same for Amministrato Azioni/ETF: the holdings of one sector, from the flat row table.
function ammFundDrill(pt,raw,acKey,sector){
  const rows=(raw&&raw.rows)||[];const L=[],V=[],C=[];let fi=0;
  rows.forEach(r=>{
    const ctrv=r.ctrv||0;if(ctrv<=0)return;
    let ac,sec,name;
    if(pt==='ammStock'){
      ac='EQUITY';const st=r.company?stockByCompany(r.company):null;
      sec=(st&&st.industry)||'Altro';name=r.company||'—';
    }else{
      const etf=r.ticker?etfByTicker(r.ticker):null;
      const t=((etf&&etf.type)||'').trim().toLowerCase();
      ac=t==='bond'?'BOND':t==='etc'?'COMMODITY':'EQUITY';
      sec=(etf&&etf.sector)||'Altro';name=(etf&&etf.name)||r.ticker||'—';
    }
    if(ac!==acKey||sec!==sector)return;
    L.push(name);V.push(Math.round(ctrv));C.push(SUB_COLORS[fi%SUB_COLORS.length]);fi++;
  });
  return L.length?{labels:L,vals:V,colors:C,name:sector,monetary:true}:null;
}

function buildDrillChart(cid,canvas,labels,vals,colors,subs,level,name){
  if(dashCharts[cid]){try{dashCharts[cid].destroy();}catch(e){}}
  const chart=new Chart(canvas,{
    type:'doughnut',
    data:{labels,datasets:[{data:vals,backgroundColor:colors,borderWidth:0,hoverBorderWidth:3,hoverBorderColor:'#fff'}]},
    options:{responsive:true,maintainAspectRatio:true,cutout:'60%',animation:donutAnim(level),
      plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>`${ctx.label}: ${fmt(ctx.parsed)}`},backgroundColor:'rgba(15,30,53,.95)',padding:8}},
      onClick:(e,els)=>{
        if(!els.length||!subs)return;
        dashDrill(cid,subs[els[0].index]);
      }
    }
  });
  dashCharts[cid]=chart;
  canvas._cur={labels,vals,colors,subs,level,name};
  if(level==='top')canvas._topData=canvas._cur;
  renderDashLegend(cid,labels,vals,colors,subs,level);
  const lbl=document.getElementById(`dcl-${cid}`);if(lbl)lbl.textContent=name||'AUM';
}

// Drill one level down on the left chart, remembering the path for ← Indietro.
function dashDrill(cid,sub){
  if(!sub||!sub.vals||!sub.vals.length)return;
  const canvas=document.getElementById(`dc-${cid}`);if(!canvas||!canvas._cur)return;
  const back=document.getElementById(`dback-${cid}`);
  if(back){(back._stack=back._stack||[]).push(canvas._cur);back.style.display='block';}
  buildDrillChart(cid,canvas,sub.labels,sub.vals,sub.colors,sub.subs||null,'sub',sub.name);
}

function renderDashLegend(cid,labels,vals,colors,subs,level){
  const leg=document.getElementById(`dleg-${cid}`);if(!leg)return;
  leg.innerHTML='';
  const total=vals.reduce((a,b)=>a+b,0)||1;
  labels.forEach((lbl,i)=>{
    const pct=((vals[i]/total)*100).toFixed(2)+'%';
    const item=document.createElement('div');item.className='dash-legend-item';
    const canDrill=subs&&subs[i]&&subs[i].vals&&subs[i].vals.length>0;
    item.innerHTML=`<div class="dash-legend-dot" style="background:${colors[i]}"></div><span class="dash-legend-name" title="${_esc(lbl)}">${_esc(lbl)}</span><span class="dash-legend-pct">${pct}${canDrill?' ›':''}</span>`;
    if(canDrill){
      item.title='Clicca per il dettaglio'; item.style.cursor='pointer';
      item.onclick=()=>dashDrill(cid,subs[i]);
    }
    leg.appendChild(item);
  });
}

function buildDrillChart2(cid,canvas,labels,vals,colors,subs,level){
  if(dashCharts['2_'+cid]){try{dashCharts['2_'+cid].destroy();}catch(e){}}
  const chart=new Chart(canvas,{
    type:'doughnut',
    data:{labels,datasets:[{data:vals,backgroundColor:colors,borderWidth:0,hoverBorderWidth:3,hoverBorderColor:'#fff'}]},
    options:{responsive:true,maintainAspectRatio:true,cutout:'60%',animation:donutAnim(level),
      plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>`${ctx.label}: ${fmt(ctx.parsed)}`},backgroundColor:'rgba(15,30,53,.95)',padding:8}},
      onClick:(e,els)=>{
        if(!els.length||!subs)return;
        const idx=els[0].index;const sub=subs[idx];
        if(sub&&sub.vals&&sub.vals.length){
          buildDrillChart2(cid,canvas,sub.labels,sub.vals,sub.colors,null,'sub');
          renderDashLegend2(cid,sub.labels,sub.vals,sub.colors,null,'sub');
          const lbl=document.getElementById(`dcl2-${cid}`);if(lbl)lbl.textContent=sub.name;
          const back=document.getElementById(`dback2-${cid}`);
          if(back){back.style.display='block';back._parent={labels,vals,colors,subs,level};}
        }
      }
    }
  });
  dashCharts['2_'+cid]=chart;
  if(level==='top')canvas._topData2={labels,vals,colors,subs,level:'top'};
  renderDashLegend2(cid,labels,vals,colors,subs,level);
  const lbl=document.getElementById(`dcl2-${cid}`);if(lbl)lbl.textContent=level==='top'?'AC':'%';
}
function renderDashLegend2(cid,labels,vals,colors,subs,level){
  const leg=document.getElementById(`dleg2-${cid}`);if(!leg)return;
  leg.innerHTML='';
  const total=vals.reduce((a,b)=>a+b,0)||1;
  labels.forEach((lbl,i)=>{
    const pct=((vals[i]/total)*100).toFixed(2);
    const canDrill=subs&&subs[i]&&subs[i].vals&&subs[i].vals.length>0;
    const item=document.createElement('div');item.className='dash-legend-item';
    item.innerHTML=`<div class="dash-legend-dot" style="background:${colors[i]}"></div><span class="dash-legend-name" title="${_esc(lbl)}">${_esc(lbl)}</span><span class="dash-legend-pct">${pct}%${canDrill?' ›':''}</span>`;
    if(canDrill){
      item.style.cursor='pointer';
      item.onclick=()=>{
        const canvas=document.getElementById(`dc2-${cid}`);
        const sub=subs[i];const parentData={labels,vals,colors,subs,level};
        buildDrillChart2(cid,canvas,sub.labels,sub.vals,sub.colors,null,'sub');
        renderDashLegend2(cid,sub.labels,sub.vals,sub.colors,null,'sub');
        const lbl2=document.getElementById(`dcl2-${cid}`);if(lbl2)lbl2.textContent=sub.name;
        const back=document.getElementById(`dback2-${cid}`);if(back){back.style.display='block';back._parent=parentData;}
      };
    }
    leg.appendChild(item);
  });
}
function drillBack2(cid){
  const back=document.getElementById(`dback2-${cid}`);
  const canvas=document.getElementById(`dc2-${cid}`);
  const parent=back?._parent;
  if(parent){
    buildDrillChart2(cid,canvas,parent.labels,parent.vals,parent.colors,parent.subs,parent.level);
    back._parent=null;
    if(parent.level==='top')back.style.display='none';
  } else {
    if(!canvas||!canvas._topData2)return;
    const{labels,vals,colors,subs}=canvas._topData2;
    buildDrillChart2(cid,canvas,labels,vals,colors,subs,'top');
    if(back)back.style.display='none';
  }
  const lbl=document.getElementById(`dcl2-${cid}`);if(lbl&&(!back||back.style.display==='none'))lbl.textContent='AC';
}

function drillBack(cid){
  const canvas=document.getElementById(`dc-${cid}`);
  const back=document.getElementById(`dback-${cid}`);
  const parent=(back&&back._stack&&back._stack.pop())||(canvas&&canvas._topData);
  if(!parent)return;
  buildDrillChart(cid,canvas,parent.labels,parent.vals,parent.colors,parent.subs,parent.level,parent.name);
  if(back&&(!back._stack||!back._stack.length))back.style.display='none';
}

