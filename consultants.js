// ── FINANCIAL CONSULTANTS ─────────────────────────────────────────────────────
// Top level of the app: pick a consultant, then work inside their client book.
let _consEditId=null, _consDelId=null;

async function loadConsultants(){const r=await fetch('/api/consultants');consultants=await r.json();}
function consName(){return consultants.find(f=>f.id===CONS)?.name||'';}
// Breadcrumb prefix that links back to the open consultant's dashboard.
function consCrumb(){return CONS?`<a class="crumb-link" onclick="openConsultant(${CONS})">${_esc(consName())}</a> › `:'';}

function openConsultant(id){
  CONS=id;AID=null;CD=null;
  document.body.classList.remove('at-cons');
  document.querySelector('.sb-srch input').value='';
  showView('dash');renderDash();renderCliList(clients);
  document.getElementById('nav-crumb').innerHTML=`<span>${_esc(consName())}</span>`;
}

async function renderConsultants(){
  const el=document.getElementById('view-cons');
  let all=_dashCache;
  if(!all){
    try{const r=await fetch('/api/dashboard');all=await r.json();_dashCache=all;clients=all.map(({portfolios_data,...c})=>c);}
    catch(e){all=clients.map(c=>({...c,portfolios_data:{}}));}
  }
  // Ordine: dal consulente creato per primo (a parità di data, id più basso)
  const byAge=[...consultants].sort((a,b)=>String(a.created_at||'').localeCompare(String(b.created_at||''))||a.id-b.id);
  const cards=byAge.map(f=>{
    const mine=all.filter(c=>c.consultant_id===f.id);
    const aum=mine.reduce((s,c)=>s+clientAum(c.portfolios_data),0);
    const ini=f.name.split(' ').filter(Boolean).map(w=>w[0]).join('').toUpperCase().slice(0,2);
    const delBtn=mine.length
      ?`<button class="cons-act" disabled title="Sposta o elimina prima i clienti">${ic('trash')}</button>`
      :`<button class="cons-act" title="Elimina" onclick="event.stopPropagation();openDelCons(${f.id})">${ic('trash')}</button>`;
    return `<div class="dash-card cons-card cons-row" onclick="openConsultant(${f.id})">
      <div class="cons-top">
        <div class="cons-av" style="background:${clientColor(f.name)}">${_esc(ini)}</div>
        <div class="dash-card-info">
          <div class="cons-role">Consulente Finanziario</div>
          <div class="dash-card-name">${_esc(f.name)}</div>
        </div>
      </div>
      <div class="cons-stat"><div class="cons-stat-v">${mine.length.toLocaleString('it-IT',{useGrouping:'always'})}</div><div class="cons-stat-l">Clienti</div></div>
      <div class="cons-stat"><div class="cons-stat-v" data-cu="${aum}">${fmt(aum)}</div><div class="cons-stat-l">Patrimonio</div></div>
      <div class="cons-acts">
        <button class="cons-act" title="Rinomina" onclick="event.stopPropagation();openConsModal(${f.id})">${ic('edit')}</button>
        <button class="cons-act" title="Esporta tutti i clienti (Excel)" onclick="event.stopPropagation();exportConsultant(${f.id})" ${mine.length?'':'disabled'}>${ic('download')}</button>${delBtn}
      </div>
    </div>`;
  }).join('');
  el.innerHTML=`<div class="dash-sec"><span>Consulenti Finanziari — ${consultants.length}</span></div>
    <div class="cons-list">${cards}
      <div class="cons-new cons-new-row" onclick="openConsModal()"><div class="cons-new-plus">+</div><div>Nuovo consulente</div></div>
    </div>`;
  countUps(el);
}

function openConsModal(id){
  _consEditId=id||null;
  const f=consultants.find(x=>x.id===id);
  document.getElementById('cons-m-title').textContent=f?'Rinomina Consulente':'Nuovo Consulente';
  document.getElementById('cons-m-btn').textContent=f?'Salva':'Crea';
  openM('cons');
  document.getElementById('inp-cons').value=f?f.name:'';
}
async function confirmCons(){
  const name=document.getElementById('inp-cons').value.trim();
  if(!name){toast('Inserisci un nome','err');return;}
  const id=_consEditId;
  const r=await fetch(id?`/api/consultants/${id}`:'/api/consultants',{method:id?'PUT':'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({name})});
  if(!r.ok){toast('Errore di salvataggio','err');return;}
  closeM('cons');await loadConsultants();renderConsultants();
  toast(id?'Consulente rinominato':`Consulente "${name}" creato`,'ok');
}
function openDelCons(id){
  _consDelId=id;
  document.getElementById('delcons-txt').textContent=`Elimina il consulente "${consultants.find(f=>f.id===id)?.name||''}"?`;
  openM('delcons');
}
async function confirmDelCons(){
  if(!_consDelId)return;
  const r=await fetch(`/api/consultants/${_consDelId}`,{method:'DELETE'});const d=await r.json();
  closeM('delcons');_consDelId=null;
  if(!r.ok){toast(d.error||'Errore','err');return;}
  await loadConsultants();renderConsultants();toast('Consulente eliminato','info');
}

// Excel con tutti i clienti del consulente: un foglio per cliente e portafoglio
function exportConsultant(id){
  const f=consultants.find(x=>x.id===id);
  window.location.href=`/api/consultants/${id}/export?_=${Date.now()}`;
  toast(`Download Excel — clienti di ${f?f.name:'consulente'}…`,'info');
}
