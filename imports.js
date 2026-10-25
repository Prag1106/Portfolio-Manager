/* ---- aggiornamento file Excel ---- */
const EXL_SYNC={stock:['/api/stock-universe/sync','/api/etf-universe/sync']};
async function exlLoad(){
  const box=document.getElementById('exl-rows');
  box.innerHTML='<p style="font-size:12px;color:var(--txt3)">Caricamento…</p>';
  try{
    const files=await (await fetch('/api/excel-files')).json();
    box.innerHTML='';
    files.forEach(f=>{
      const row=document.createElement('div');
      row.style.cssText='display:flex;align-items:center;gap:10px;padding:10px 0;border-bottom:1px solid rgba(128,128,128,.15)';
      row.innerHTML=`<div style="flex:1;min-width:0">
        <div style="font-size:12px;font-weight:700">${f.label}</div>
        <div style="font-size:10px;color:var(--txt3)">${f.filename} — aggiornato: ${f.updated||'mai'}</div>
        <div style="font-size:10px;color:var(--txt3)" id="exl-st-${f.key}"></div></div>
      <input type="file" accept=".xlsx" id="exl-inp-${f.key}" style="display:none" onchange="exlUpload('${f.key}')"/>
      <button class="btn btn-primary" style="padding:6px 14px;font-size:11px" onclick="document.getElementById('exl-inp-${f.key}').click()">Carica nuovo…</button>`;
      box.appendChild(row);
    });
  }catch(e){box.innerHTML='<p style="font-size:12px;color:#ef6a5a">Errore caricamento elenco file</p>';}
}
async function exlUpload(key){
  const inp=document.getElementById('exl-inp-'+key);
  const st=document.getElementById('exl-st-'+key);
  if(!inp.files.length)return;
  const fd=new FormData();fd.append('file',inp.files[0]);
  st.textContent='Caricamento in corso…';
  try{
    const r=await fetch('/api/excel-files/'+key,{method:'POST',body:fd});
    const d=await r.json();
    if(!r.ok)throw new Error(d.error||'errore server');
    for(const u of (EXL_SYNC[key]||[])){st.textContent='Sincronizzazione database…';await fetch(u,{method:'POST'});}
    if(key==='stock'){await loadStockUniverse();await loadEtfUniverse();}
    st.textContent='';
    toast('✓ File Excel aggiornato','ok');
    exlLoad();
  }catch(e){st.textContent='';toast('Errore: '+e.message,'err');}
  finally{inp.value='';}
}
/* ---- import fondi da PDF Provvigionale ---- */
let PDFIMP={categories:[],new:[],updates:[]};
const _pct=v=>(v===null||v===undefined||v==='')?'—':(v*100).toFixed(3).replace(/\.?0+$/,'')+'%';
async function pdfImportAnalyze(){
  const inp=document.getElementById('pdfimp-input');
  if(!inp.files.length)return;
  const ld=document.getElementById('pdfimp-loading');
  const fd=new FormData();fd.append('file',inp.files[0]);
  ld.textContent='Analisi PDF in corso…';
  try{
    const r=await fetch('/api/sicav-import-pdf/analyze',{method:'POST',body:fd});
    const d=await r.json();
    if(!r.ok)throw new Error(d.error||'errore server');
    PDFIMP=d;
    closeM('excelup');
    document.getElementById('m-pdfimport').classList.add('open');
    renderPdfImport();
  }catch(e){toast('Errore: '+e.message,'err');}
  finally{ld.textContent='';inp.value='';}
}
function renderPdfImport(){
  const s=PDFIMP.stats;
  document.getElementById('pdfimp-stats').textContent=`Letti ${s.parsed} · Nuovi ${s.new} · Aggiornamenti costi ${s.updates} · Già presenti ${s.existing}`;
  const opts=PDFIMP.categories.map(c=>`<option value="${_esc(c)}">${_esc(c)}</option>`).join('');
  let h='';
  if(!PDFIMP.new.length && !PDFIMP.updates.length){
    h='<div style="padding:40px;text-align:center;color:#8fa0c0;font-size:13px">Nessun nuovo fondo e nessun costo variato: l\'Excel è già aggiornato rispetto a questo PDF.</div>';
    document.getElementById('pdfimp-commit').style.display='none';
  }else{
    document.getElementById('pdfimp-commit').style.display='';
  }
  if(PDFIMP.new.length){
    h+=`<div style="font-size:13px;font-weight:800;color:#e2e9f8;margin:4px 0 8px">Nuovi fondi da aggiungere (${PDFIMP.new.length}) <span style="font-weight:500;color:#8fa0c0">— controlla e correggi la Categoria Quantalys se necessario</span></div>`;
    h+='<table style="width:100%;border-collapse:collapse;font-size:11px;background:#142443;border-radius:8px;overflow:hidden"><thead><tr style="background:#12203c;color:#fff;text-align:left">'+
       ['ISIN','Fondo','Società','Cl.','Val.','Comm.','BS/CG','Spese BS','Categoria Quantalys','Fonte'].map(t=>`<th style="padding:6px 8px;font-weight:600">${t}</th>`).join('')+'</tr></thead><tbody>';
    PDFIMP.new.forEach((r,i)=>{
      const badge=r.method==='auto'?'<span style="color:#e8c060">auto</span>':`<span style="color:#6fd39c">${r.method}</span>`;
      h+=`<tr style="border-bottom:1px solid rgba(255,255,255,.07)">
        <td style="padding:5px 8px;font-family:monospace;color:#9aa6c0">${_esc(r.isin)}</td>
        <td style="padding:5px 8px">${_esc(r.name)}</td>
        <td style="padding:5px 8px;color:#8fa0c0">${_esc(r.societa)}</td>
        <td style="padding:5px 8px">${_esc(r.classe)}</td>
        <td style="padding:5px 8px">${_esc(r.currency)}</td>
        <td style="padding:5px 8px">${_pct(r.fee)}</td>
        <td style="padding:5px 8px">${_pct(r.bs_fee)}</td>
        <td style="padding:5px 8px">${_pct(r.bs)}</td>
        <td style="padding:4px 8px"><select id="pdfimp-q-${i}" style="width:100%;font-size:11px;padding:3px;border:1px solid #3a5080;border-radius:4px;background:#1a2c50;color:#eef2fa">${opts}</select></td>
        <td style="padding:5px 8px;font-size:10px">${badge}</td></tr>`;
    });
    h+='</tbody></table>';
  }
  if(PDFIMP.updates.length){
    h+=`<div style="font-size:13px;font-weight:800;color:#e2e9f8;margin:18px 0 8px">Costi variati su fondi già presenti (${PDFIMP.updates.length}) <span style="font-weight:500;color:#8fa0c0">— deseleziona quelli da non aggiornare</span></div>`;
    h+='<table style="width:100%;border-collapse:collapse;font-size:11px;background:#142443;border-radius:8px;overflow:hidden"><thead><tr style="background:#12203c;color:#fff;text-align:left">'+
       ['','ISIN','Fondo','Comm.','BS/CG','Spese BS'].map(t=>`<th style="padding:6px 8px;font-weight:600">${t}</th>`).join('')+'</tr></thead><tbody>';
    PDFIMP.updates.forEach((u,i)=>{
      const cell=(o,n)=>o===n?_pct(n):`<span style="color:#7f8aa6;text-decoration:line-through">${_pct(o)}</span> <span style="color:#ef6a5a;font-weight:700">${_pct(n)}</span>`;
      h+=`<tr style="border-bottom:1px solid rgba(255,255,255,.07)">
        <td style="padding:5px 8px"><input type="checkbox" id="pdfimp-u-${i}" checked></td>
        <td style="padding:5px 8px;font-family:monospace;color:#9aa6c0">${_esc(u.isin)}</td>
        <td style="padding:5px 8px">${_esc(u.name)}</td>
        <td style="padding:5px 8px">${cell(u.old.fee,u.new.fee)}</td>
        <td style="padding:5px 8px">${cell(u.old.bs_fee,u.new.bs_fee)}</td>
        <td style="padding:5px 8px">${cell(u.old.bs,u.new.bs)}</td></tr>`;
    });
    h+='</tbody></table>';
  }
  document.getElementById('pdfimp-body').innerHTML=h;
  PDFIMP.new.forEach((r,i)=>{const s=document.getElementById('pdfimp-q-'+i);if(s)s.value=r.cat_q;});
}
async function pdfImportCommit(){
  const btn=document.getElementById('pdfimp-commit');
  const newRows=PDFIMP.new.map((r,i)=>({...r,cat_q:document.getElementById('pdfimp-q-'+i).value}));
  const updates=PDFIMP.updates.filter((u,i)=>document.getElementById('pdfimp-u-'+i).checked);
  btn.disabled=true;btn.textContent='Salvataggio…';
  try{
    const r=await fetch('/api/sicav-import-pdf/commit',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({new:newRows,updates})});
    const d=await r.json();
    if(!r.ok)throw new Error(d.error||'errore server');
    document.getElementById('m-pdfimport').classList.remove('open');
    toast(`✓ ${d.added} fondi aggiunti · ${d.updated} costi aggiornati`,'ok');
  }catch(e){toast('Errore: '+e.message,'err');}
  finally{btn.disabled=false;btn.textContent='Conferma e salva';}
}
/* ---- import DARTA Challenge Pro/Private da circolare ---- */
let DARTA={};
async function dartaImportAnalyze(sheet){
  const inp=document.getElementById('darta-input-'+sheet);
  if(!inp.files.length)return;
  const ld=document.getElementById('darta-loading');
  const fd=new FormData();fd.append('file',inp.files[0]);fd.append('sheet',sheet);
  ld.textContent='Analisi circolare in corso…';
  try{
    const r=await fetch('/api/darta-import-pdf/analyze',{method:'POST',body:fd});
    const d=await r.json();
    if(!r.ok)throw new Error(d.error||'errore server');
    DARTA=d;
    closeM('excelup');
    document.getElementById('m-dartaimport').classList.add('open');
    renderDartaImport();
  }catch(e){toast('Errore: '+e.message,'err');}
  finally{ld.textContent='';inp.value='';}
}
function renderDartaImport(){
  const s=DARTA.stats;
  document.getElementById('darta-title').textContent='Import DARTA — '+DARTA.sheet_name;
  document.getElementById('darta-stats').textContent=`Letti ${s.parsed} · Nuovi ${s.new} · Già presenti ${s.matched}`;
  const opts=DARTA.categories.map(c=>`<option value="${_esc(c)}">${_esc(c)}</option>`).join('');
  const upd=DARTA.matched.filter(m=>Object.keys(m.changes).length||m.cat_was_empty);
  let h='';
  if(!DARTA.new.length && !upd.length){
    h='<div style="padding:40px;text-align:center;color:#8fa0c0;font-size:13px">Tutti i fondi della circolare sono già presenti e aggiornati: nessuna modifica necessaria.</div>';
    document.getElementById('darta-commit').style.display='none';
  }else{document.getElementById('darta-commit').style.display='';}
  if(DARTA.new.length){
    h+=`<div style="font-size:13px;font-weight:800;color:#e2e9f8;margin:4px 0 8px">Nuovi fondi da aggiungere (${DARTA.new.length}) <span style="font-weight:500;color:#8fa0c0">— controlla Asset Manager e Categoria Quantalys</span></div>`;
    h+='<table style="width:100%;border-collapse:collapse;font-size:11px;background:#142443;border-radius:8px;overflow:hidden"><thead><tr style="background:#12203c;color:#fff;text-align:left">'+
      ['','Asset Manager','Fondo Interno','CG','Spese Corr.','Comm. Perf.','Categoria Quantalys'].map(t=>`<th style="padding:6px 8px;font-weight:600">${t}</th>`).join('')+'</tr></thead><tbody>';
    DARTA.new.forEach((r,i)=>{
      h+=`<tr style="border-bottom:1px solid rgba(255,255,255,.07)">
        <td style="padding:5px 8px"><input type="checkbox" id="darta-n-chk-${i}" checked></td>
        <td style="padding:4px 8px"><input id="darta-n-am-${i}" value="${_esc(r.asset_manager)}" style="width:110px;font-size:11px;padding:3px;border:1px solid #3a5080;border-radius:4px;background:#1a2c50;color:#eef2fa"></td>
        <td style="padding:5px 8px">${_esc(r.name)}</td>
        <td style="padding:5px 8px">${r.has_cg?_esc(r.cg||'—'):'<span style=\"color:#5c6c8c\">n/d</span>'}</td>
        <td style="padding:5px 8px">${_pct(r.spese)}</td>
        <td style="padding:5px 8px">${r.has_perf?_esc(r.perf||'—'):'<span style=\"color:#5c6c8c\">n/d</span>'}</td>
        <td style="padding:4px 8px"><select id="darta-n-q-${i}" style="width:100%;font-size:11px;padding:3px;border:1px solid #3a5080;border-radius:4px;background:#1a2c50;color:#eef2fa">${opts}</select></td></tr>`;
    });
    h+='</tbody></table>';
  }
  if(upd.length){
    h+=`<div style="font-size:13px;font-weight:800;color:#e2e9f8;margin:18px 0 8px">Aggiornamenti a fondi esistenti (${upd.length})</div>`;
    h+='<table style="width:100%;border-collapse:collapse;font-size:11px;background:#142443;border-radius:8px;overflow:hidden"><thead><tr style="background:#12203c;color:#fff;text-align:left">'+
      ['','Fondo Interno','CG','Spese Corr.','Comm. Perf.','Categoria Quantalys'].map(t=>`<th style="padding:6px 8px;font-weight:600">${t}</th>`).join('')+'</tr></thead><tbody>';
    upd.forEach((m)=>{
      const gi=DARTA.matched.indexOf(m);
      const dc=(k,cur)=>{const c=m.changes[k];if(!c)return cur===undefined||cur===null?'—':(typeof cur==='number'?_pct(cur):_esc(cur));const o=typeof c.old==='number'?_pct(c.old):(c.old==null?'—':_esc(c.old));const nv=typeof c.new==='number'?_pct(c.new):_esc(c.new);return `<span style="color:#7f8aa6;text-decoration:line-through">${o}</span> <span style="color:#ef6a5a;font-weight:700">${nv}</span>`;};
      h+=`<tr style="border-bottom:1px solid rgba(255,255,255,.07)">
        <td style="padding:5px 8px"><input type="checkbox" id="darta-u-chk-${gi}" checked></td>
        <td style="padding:5px 8px">${_esc(m.name)}</td>
        <td style="padding:5px 8px">${m.has_cg?dc('cg',m.cg):'<span style=\"color:#5c6c8c\">n/d</span>'}</td>
        <td style="padding:5px 8px">${dc('spese',m.spese)}</td>
        <td style="padding:5px 8px">${m.has_perf?dc('perf',m.perf):'<span style=\"color:#5c6c8c\">n/d</span>'}</td>
        <td style="padding:4px 8px"><select id="darta-u-q-${gi}" style="width:100%;font-size:11px;padding:3px;border:1px solid ${m.cat_was_empty?'#e0a800':'#3a5080'};border-radius:4px;background:#1a2c50;color:#eef2fa">${opts}</select></td></tr>`;
    });
    h+='</tbody></table>';
  }
  document.getElementById('darta-body').innerHTML=h;
  DARTA.new.forEach((r,i)=>{const s=document.getElementById('darta-n-q-'+i);if(s)s.value=r.cat_q;});
  upd.forEach((m)=>{const gi=DARTA.matched.indexOf(m);const s=document.getElementById('darta-u-q-'+gi);if(s)s.value=m.cat_q;});
}
async function dartaImportCommit(){
  const btn=document.getElementById('darta-commit');
  const newRows=DARTA.new.map((r,i)=>({...r,asset_manager:document.getElementById('darta-n-am-'+i).value,cat_q:document.getElementById('darta-n-q-'+i).value}))
                         .filter((r,i)=>document.getElementById('darta-n-chk-'+i).checked);
  const rows=DARTA.matched.map((m,gi)=>{const chk=document.getElementById('darta-u-chk-'+gi);if(!chk||!chk.checked)return null;const q=document.getElementById('darta-u-q-'+gi);return {name:m.name,cg:m.cg,spese:m.spese,perf:m.perf,cat_q:q?q.value:m.cat_q};}).filter(Boolean);
  btn.disabled=true;btn.textContent='Salvataggio…';
  try{
    const r=await fetch('/api/darta-import-pdf/commit',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({sheet:DARTA.sheet,rows,new:newRows})});
    const d=await r.json();
    if(!r.ok)throw new Error(d.error||'errore server');
    document.getElementById('m-dartaimport').classList.remove('open');
    toast(`✓ ${d.added} aggiunti · ${d.updated} aggiornati`,'ok');
  }catch(e){toast('Errore: '+e.message,'err');}
  finally{btn.disabled=false;btn.textContent='Conferma e salva';}
}
