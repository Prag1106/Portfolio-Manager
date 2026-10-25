async function confirmAdd(){const name=document.getElementById('inp-add').value.trim();if(!name){toast('Inserisci un nome','err');return;}const r=await fetch('/api/clients',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({name,consultant_id:CONS})});const d=await r.json();closeM('add');await loadClients();selectClient(d.id);toast(`Cliente "${name}" creato`,'ok');}
function openDel(id,name){delId=id;document.getElementById('del-txt').textContent=`Elimina "${name}"? Tutti i portafogli verranno eliminati permanentemente.`;openM('del');}
async function confirmDel(){if(!delId)return;await fetch(`/api/clients/${delId}`,{method:'DELETE'});closeM('del');delete _ptfCache[delId];_dashCache=null;if(AID===delId){AID=null;CD=null;showView('dash');}await loadClients();renderDash();toast('Cliente eliminato','info');delId=null;}
function openDup(id,name){dupId=id;document.getElementById('inp-dup').value=name+' (copia)';document.getElementById('m-dup').classList.add('open');setTimeout(()=>document.getElementById('inp-dup').focus(),80);}
async function confirmDup(){const name=document.getElementById('inp-dup').value.trim();if(!name){toast('Inserisci un nome','err');return;}const r=await fetch(`/api/clients/${dupId}/duplicate`,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({name})});const d=await r.json();closeM('dup');await loadClients();selectClient(d.id);toast(`Duplicato come "${name}"`,'ok');dupId=null;}
function openAddAc(ptfl){
  addAcPtfl=ptfl;
  document.getElementById('inp-ac-name').value='';
  document.getElementById('inp-ac-sub1').value='';
  const existing=Object.keys(getW(ptfl)||{});
  const sel=document.getElementById('sel-ac-type');
  sel.innerHTML='';
  // Preset options for fixed ACs not currently in portfolio
  const PRESETS=[
    {value:'preset_equity',label:'Azionario',name:'EQUITY'},
    {value:'preset_bond',label:'Obbligazionario',name:'BOND'},
    {value:'preset_bilanciati',label:'Bilanciati',name:'BILANCIATI'},
    {value:'preset_flessibili',label:'Flessibili',name:'FLESSIBILI'},
    {value:'preset_pa',label:'Performance Assoluta',name:'PERFORMANCE ASSOLUTA'},
    {value:'preset_commodity',label:'Commodity',name:'COMMODITY'},
  ];
  // START & GO is DARTA-only: it feeds the Easy Switch plan.
  if(ptfl==='darta')PRESETS.push({value:'preset_sg',label:'START & GO → Easy Switch',name:'START & GO'});
  PRESETS.filter(p=>!existing.includes(p.name)).forEach(p=>{
    const o=document.createElement('option');o.value=p.value;o.textContent=p.label;sel.appendChild(o);
  });
  // Separator + custom types
  if(sel.options.length){const sep=document.createElement('option');sep.disabled=true;sep.textContent='──────────';sel.appendChild(sep);}
  [{value:'custom_bilanciato',label:'Personalizzato – Bilanciato'},
   {value:'custom_flessibili',label:'Personalizzato – Flessibili'},
   {value:'custom_altri',label:'Personalizzato – Perf. Assoluta'}].forEach(({value,label})=>{
    const o=document.createElement('option');o.value=value;o.textContent=label;sel.appendChild(o);
  });
  _syncAddAcInputs();
  sel.onchange=_syncAddAcInputs;
  document.getElementById('m-addac').classList.add('open');
  setTimeout(()=>sel.focus(),80);
}
function _syncAddAcInputs(){
  const v=document.getElementById('sel-ac-type').value;
  const isCustom=v.startsWith('custom_');
  document.getElementById('inp-ac-name').style.display=isCustom?'':'none';
  document.getElementById('inp-ac-sub1').style.display=isCustom?'':'none';
}
function addAcPreset(ptfl,name){
  const w=getW(ptfl);
  if(name==='EQUITY') w['EQUITY']={pct:0,subs:{},_type:'equity'};
  else if(name==='BOND') w['BOND']={pct:0,subs:{},_type:'bond'};
  else if(name==='COMMODITY') w['COMMODITY']={pct:0,subs:{}};
  else if(name==='BILANCIATI') w['BILANCIATI']={pct:0,subs:{},_type:'bilanciato'};
  else if(name==='FLESSIBILI') w['FLESSIBILI']={pct:0,subs:{},_type:'flessibili'};
  else if(name==='PERFORMANCE ASSOLUTA') w['PERFORMANCE ASSOLUTA']={pct:0,subs:{},_type:'altri'};
  else if(name==='START & GO'){
    // Residual bucket: default to the leftover of the DARTA total (100% − real funds).
    const others=getWOrder(ptfl).filter(a=>a!=='START & GO').reduce((s,a)=>s+(w[a]?.pct||0),0);
    w['START & GO']={pct:Math.max(0,100-others),subs:{}};
    const child=PAC_CHILD[ptfl];
    if(child){
      const bk='_'+child+'Backup';
      if(CD[bk]&&Object.keys(CD[bk]).length){CD.portfolios[child]=CD[bk];delete CD[bk];}
      else if(!CD.portfolios[child]){CD.portfolios[child]=makeDefault(child);}
    }
    // Turn on the Easy Switch plan and make START & GO its source.
    const pd=CD.portfolios[ptfl];
    if(!pd.pac)pd.pac={};
    pd.pac.enabled=true;
    pd.pac.sourceKey='__SG__';
    if(pd.pac.mesi==null)pd.pac.mesi=15;
    if(child)buildTab(child);
    // Reveal the Easy Switch tab, stay on this tab, then scroll it into view.
    const en=ALL_PTFL.filter(t=>!isPacType(t)&&!!CD.portfolios[t]);renderTotals(en);renderTabs(en);switchTab(ptfl);
    setTimeout(()=>document.getElementById('es-amount-inp')?.scrollIntoView({block:'center',behavior:'smooth'}),60);
  }
  const order=getWOrder(ptfl);if(!order.includes(name))order.push(name);setWOrder(ptfl,order);
  buildTab(ptfl);toast(`"${name}" ripristinata`,'ok');
}
function confirmAddAc(){
  const type=document.getElementById('sel-ac-type').value;
  if(type.startsWith('preset_')){
    const map={preset_equity:'EQUITY',preset_bond:'BOND',preset_bilanciati:'BILANCIATI',preset_flessibili:'FLESSIBILI',preset_pa:'PERFORMANCE ASSOLUTA',preset_commodity:'COMMODITY',preset_sg:'START & GO'};
    recordUndo();addAcPreset(addAcPtfl,map[type]);closeM('addac');return;
  }
  const name=document.getElementById('inp-ac-name').value.trim().toUpperCase();
  if(!name){toast('Inserisci un nome','err');return;}
  const sub1=document.getElementById('inp-ac-sub1').value.trim();
  const w=getW(addAcPtfl);
  const acType=type.replace('custom_','');
  const subs=sub1?{[sub1]:100}:{};
  recordUndo();w[name]={pct:0,subs,_type:acType};
  const order=getWOrder(addAcPtfl);order.push(name);setWOrder(addAcPtfl,order);
  closeM('addac');buildTab(addAcPtfl);toast(`"${name}" aggiunta`,'ok');
}
function _getAcCatGroup(ptfl,ac){
  const acd=getW(ptfl)[ac];const type=acd?._type||'';
  if(ac==='EQUITY'||type==='equity')return 'equity';
  if(ac==='BOND'||type==='bond')return 'bond';
  if(ac==='COMMODITY'||type==='commodity')return 'commodity';
  if(type==='bilanciato')return 'bilanciato';
  if(type==='flessibili')return 'flessibili';
  if(type==='altri')return 'altri';
  return null;
}
function _buildCatTreePicker(group,existing){
  const picker=document.getElementById('cat-tree-picker');
  picker.innerHTML='';
  _catTreeSelected=null;
  const nodes=(catTree[group]||[]);
  if(!nodes.length)return false;
  nodes.forEach(node=>{
    const used=existing.includes(node.name);
    const children=(node.children||[]).filter(c=>!existing.includes(c));
    const hasChildren=children.length>0;
    // Build childWrap first so the expand listener can close over it
    let childWrap=null;
    if(hasChildren){
      childWrap=document.createElement('div');
      childWrap.style.display='none';
      children.forEach(c=>{
        const crow=document.createElement('div');crow.className='cat-tree-item child-item';
        const ph=document.createElement('span');ph.className='cat-tree-placeholder';crow.appendChild(ph);
        const cl=document.createElement('span');cl.className='cat-tree-label';cl.textContent=c;
        crow.appendChild(cl);
        crow.addEventListener('click',()=>_selectCatTreeItem(c,crow));
        childWrap.appendChild(crow);
      });
    }
    // Parent row
    const row=document.createElement('div');
    row.className='cat-tree-item'+(used?' cat-tree-used':'');
    row.style.opacity=used?'0.4':'1';
    row.style.pointerEvents=used?'none':'';
    if(hasChildren){
      const exp=document.createElement('span');
      exp.className='cat-tree-expand';exp.textContent='▶';
      exp.addEventListener('click',e=>{
        e.stopPropagation();
        const open=exp.classList.toggle('open');
        childWrap.style.display=open?'':'none';
      });
      row.appendChild(exp);
    } else {
      const ph=document.createElement('span');ph.className='cat-tree-placeholder';row.appendChild(ph);
    }
    const lbl=document.createElement('span');lbl.className='cat-tree-label';lbl.textContent=node.name;
    row.appendChild(lbl);
    row.addEventListener('click',()=>_selectCatTreeItem(node.name,row));
    picker.appendChild(row);
    if(childWrap)picker.appendChild(childWrap);
  });
  // custom entry at bottom
  const customRow=document.createElement('div');customRow.className='cat-tree-item';
  const ph=document.createElement('span');ph.className='cat-tree-placeholder';customRow.appendChild(ph);
  const cl=document.createElement('span');cl.className='cat-tree-label';cl.style.color='rgba(255,255,255,.45)';cl.textContent='✎ Personalizzato…';
  customRow.appendChild(cl);
  customRow.addEventListener('click',()=>{
    _catTreeSelected=null;
    picker.querySelectorAll('.cat-tree-item').forEach(r=>r.classList.remove('selected'));
    customRow.classList.add('selected');
    const inp=document.getElementById('inp-sub-name');
    inp.style.display='';inp.value='';setTimeout(()=>inp.focus(),60);
  });
  picker.appendChild(customRow);
  return true;
}
function _selectCatTreeItem(name,rowEl){
  document.getElementById('cat-tree-picker').querySelectorAll('.cat-tree-item').forEach(r=>r.classList.remove('selected'));
  rowEl.classList.add('selected');
  _catTreeSelected=name;
  document.getElementById('inp-sub-name').style.display='none';
}
function openAddSub(ptfl,ac){
  addSubPtfl=ptfl;addSubAc=ac;
  document.getElementById('addsub-txt').textContent=`Nuovo sotto-settore in ${ac}:`;
  document.getElementById('m-addsub').classList.add('open');
  const group=_getAcCatGroup(ptfl,ac);
  const picker=document.getElementById('cat-tree-picker');
  const inp=document.getElementById('inp-sub-name');
  inp.value='';
  const existing=Object.keys(getW(ptfl)[ac]?.subs||{});
  if(ptfl!=='amm'&&ptfl!=='ammAz'&&group&&catTree[group]?.length&&_buildCatTreePicker(group,existing)){
    picker.style.display='';inp.style.display='none';
  } else {
    picker.style.display='none';inp.style.display='';
    setTimeout(()=>inp.focus(),80);
  }
}
function confirmAddSub(){
  const inp=document.getElementById('inp-sub-name');
  const name=_catTreeSelected||(inp.style.display!=='none'?inp.value.trim():'');
  if(!name){toast('Seleziona o inserisci un nome','err');return;}
  const w=getW(addSubPtfl);if(!w[addSubAc]){closeM('addsub');return;}
  recordUndo();w[addSubAc].subs[name]=0;
  const so=getSubOrder(addSubPtfl,addSubAc);
  if(!so.includes(name)){so.push(name);setSubOrder(addSubPtfl,addSubAc,so);}
  closeM('addsub');buildTab(addSubPtfl);toast(`Settore "${name}" aggiunto`,'ok');
}

