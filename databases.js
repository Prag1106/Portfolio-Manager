// ── FUND DATABASE / STOCK DATABASE ────────────────────────────────────────────
function _fdbCI(kw,headers){return headers.findIndex(h=>h&&h.toLowerCase().includes(kw));}

async function openFundDb(){
  // open the panel right away; data streams in behind the loading note
  _fdbCurrentTab='pro';
  document.getElementById('fdb-search').value='';
  ['pro','chpv','sicav'].forEach(t=>document.getElementById('fdb-tab-'+t)?.classList.toggle('active',t==='pro'));
  document.getElementById('m-funddb').classList.add('open');
  document.getElementById('fdb-body').innerHTML='<div style="padding:28px;text-align:center;color:var(--txt3);font-size:12px">Caricamento database…</div>';
  try{
    const [r,rs]=await Promise.all([fetch('/api/fund-db'),fetch('/api/sicav-fund-db')]);
    const raw=await r.json();
    ['pro','chpv'].forEach(tab=>{
      _fdbData[tab]={
        headers:raw[tab]?.headers||[],
        rows:(raw[tab]?.rows||[]).map(r=>({_raw:[...r],_id:_fdbNextId++,_new:false}))
      };
    });
    const rawS=await rs.json();
    _fdbData.sicav={headers:rawS.headers||[],rows:(rawS.rows||[]).map(r=>({_raw:[...r],_id:_fdbNextId++,_new:false}))};
  }catch(e){toast('Errore caricamento database','err');document.getElementById('m-funddb').classList.remove('open');return;}
  fdbRender();
}

function fdbSetTab(tab){
  _fdbCurrentTab=tab;
  ['pro','chpv','sicav'].forEach(t=>document.getElementById('fdb-tab-'+t)?.classList.toggle('active',t===tab));
  document.getElementById('fdb-search').value='';
  const cf=document.getElementById('fdb-cat-filter');if(cf)cf.value='';
  fdbRender();
}

function fdbRender(){
  const q=(document.getElementById('fdb-search')?.value||'').trim().toLowerCase();
  const isSicav=_fdbCurrentTab==='sicav';
  const td=_fdbData[_fdbCurrentTab];
  const hdrs=td.headers||[];

  // Column indices differ per sheet
  const ci=isSicav?{
    name:_fdbCI('name',hdrs),
    isin:_fdbCI('isin',hdrs),
    cat:_fdbCI('categoria quantalys',hdrs),
    fee:_fdbCI('fee',hdrs),
    cg:-1,sc:-1
  }:{
    name:_fdbCI('fondo interno',hdrs),
    isin:_fdbCI('isin under',hdrs),
    cat:_fdbCI('categoria quantalys',hdrs),
    cg:_fdbCI('cg (',hdrs),
    sc:_fdbCI('spese correnti',hdrs),
    fee:-1
  };
  const hasCg=!isSicav&&ci.cg>=0;

  const rows=td.rows||[];

  // Build categoria filter dropdown
  const catSel=document.getElementById('fdb-cat-filter');
  if(catSel){
    const cats=[...new Set(rows.map(r=>ci.cat>=0&&r._raw[ci.cat]?String(r._raw[ci.cat]).trim():'').filter(Boolean))].sort();
    const curCat=catSel.value;
    catSel.innerHTML=`<option value="">Tutte le categorie</option>`+cats.map(c=>`<option value="${c}"${c===curCat?' selected':''}>${c}</option>`).join('');
  }
  const activeCat=catSel?.value||'';

  const shown=rows.filter(r=>{
    const matchQ=!q||[ci.name,ci.isin,ci.cat].some(i=>i>=0&&r._raw[i]&&String(r._raw[i]).toLowerCase().includes(q));
    const matchCat=!activeCat||(ci.cat>=0&&String(r._raw[ci.cat]||'').trim()===activeCat);
    return matchQ&&matchCat;
  });
  document.getElementById('fdb-count').textContent=`${shown.length} / ${rows.length} fondi`;

  const tbl=document.createElement('table');tbl.className='fdb-tbl';
  tbl.innerHTML=`<thead><tr>
    <th style="width:35%">${isSicav?'Nome':'Fondo Interno'}</th>
    <th style="width:13%">ISIN</th>
    <th style="width:28%">Categoria Quantalys</th>
    ${hasCg?'<th style="width:8%">CG %</th>':''}
    ${!isSicav?'<th style="width:8%">SC %</th>':''}
    ${isSicav&&ci.fee>=0?'<th style="width:8%">Fee %</th>':''}
    <th style="width:36px"></th>
  </tr></thead>`;

  const tbody=document.createElement('tbody');
  shown.forEach(rowObj=>{
    const raw=rowObj._raw||[];
    const tr=document.createElement('tr');
    if(rowObj._new)tr.classList.add('fdb-new');

    const mkTd=(colIdx,placeholder,mono,pctMode)=>{
      const td=document.createElement('td');
      const inp=document.createElement('input');
      inp.className='fdb-inp';
      if(mono)inp.style.fontFamily="'JetBrains Mono',monospace";
      const rawVal=colIdx>=0&&raw[colIdx]!=null?raw[colIdx]:'';
      // SC / Fee stored as decimal (0.0059), display/edit as % ("0.59%")
      const num=parseFloat(String(rawVal).replace(',','.'));
      if(pctMode&&rawVal!==''&&!isNaN(num)){
        inp.value=String(rawVal).includes('%')?String(rawVal):(num*100).toFixed(2)+'%';
      }else{inp.value=rawVal;}
      inp.placeholder=placeholder||'';
      inp.addEventListener('change',e=>{
        if(!rowObj._raw)rowObj._raw=Array(hdrs.length).fill(null);
        if(colIdx>=0){
          const v=e.target.value;
          // Convert % back to decimal
          const vn=parseFloat(String(v).replace('%','').replace(',','.'));
          if(pctMode&&v!==''&&!isNaN(vn)){rowObj._raw[colIdx]=String(+(vn/100).toFixed(6));}
          else{rowObj._raw[colIdx]=v||null;}
        }
      });
      td.appendChild(inp);return td;
    };

    tr.appendChild(mkTd(ci.name,isSicav?'Nome fondo…':'Nome fondo…'));
    tr.appendChild(mkTd(ci.isin,'ISIN…',true));
    tr.appendChild(mkTd(ci.cat,'Categoria…'));
    if(hasCg)tr.appendChild(mkTd(ci.cg,'es. 0.8'));
    if(!isSicav)tr.appendChild(mkTd(ci.sc,'es. 1.20%',false,true));
    if(isSicav&&ci.fee>=0)tr.appendChild(mkTd(ci.fee,'es. 1.60%',false,true));

    const delTd=document.createElement('td');
    const delBtn=document.createElement('button');
    delBtn.className='fdb-del-btn';delBtn.textContent='×';delBtn.title='Rimuovi';
    delBtn.addEventListener('click',()=>{
      const lbl=(rowObj._raw||[]).find(x=>x&&String(x).trim())||'questa riga';
      if(!confirm(`⚠️ ELIMINAZIONE PERMANENTE\n\n"${lbl}" verrà rimosso dal database dei fondi.\nLa rimozione diventa definitiva al salvataggio del database.\n\nConfermi?`))return;
      _fdbData[_fdbCurrentTab].rows=_fdbData[_fdbCurrentTab].rows.filter(r=>r._id!==rowObj._id);
      fdbRender();
    });
    delTd.appendChild(delBtn);tr.appendChild(delTd);
    tbody.appendChild(tr);
  });

  tbl.appendChild(tbody);
  const body=document.getElementById('fdb-body');
  body.innerHTML='';body.appendChild(tbl);
}

function fdbAddRow(){
  const td=_fdbData[_fdbCurrentTab];
  const newRow={_raw:Array(td.headers.length).fill(null),_id:_fdbNextId++,_new:true};
  td.rows.unshift(newRow);
  document.getElementById('fdb-search').value='';
  fdbRender();
  document.getElementById('fdb-body')?.querySelector('.fdb-inp')?.focus();
}

async function fdbSave(){
  try{
    const payload={};
    ['pro','chpv'].forEach(tab=>{
      payload[tab]={headers:_fdbData[tab].headers,rows:_fdbData[tab].rows.map(r=>r._raw||[])};
    });
    const sicavPayload={headers:_fdbData.sicav.headers,rows:_fdbData.sicav.rows.map(r=>r._raw||[])};
    const [r,rs]=await Promise.all([
      fetch('/api/fund-db',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)}),
      fetch('/api/sicav-fund-db',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(sicavPayload)})
    ]);
    if(r.ok&&rs.ok){
      toast('Excel aggiornati','ok');
      loadDartaFundsByCat();loadDartaCats();loadSicavFundsByCat();loadCategoryTree();
    }else{toast('Errore nel salvataggio','err');}
  }catch(e){toast('Errore nel salvataggio','err');}
}

// ── STOCK/ETF DATABASE PANEL ──────────────────────────────────────────────────
let _sdbTab='stocks';
function openStockDb(){document.getElementById('m-stockdb').classList.add('open');sdbRender();}

function sdbSetTab(tab){
  _sdbTab=tab;
  document.getElementById('sdb-tab-stocks')?.classList.toggle('active',tab==='stocks');
  document.getElementById('sdb-tab-etfs')?.classList.toggle('active',tab==='etfs');
  const lbl=document.getElementById('sdb-add-label');
  const inp=document.getElementById('sdb-new-ticker');
  const note=document.getElementById('sdb-footer-note');
  if(tab==='etfs'){
    if(lbl)lbl.textContent='Aggiungi ETF ID:';
    if(inp)inp.placeholder='Es. IE00B4L5Y983 (ISIN) o NUCL…';
    if(note)note.innerHTML='Usa il <strong style="color:#5ed68a">ticker Refinitiv</strong> o l\'<strong style="color:#5ed68a">ISIN</strong> dell\'ETF. Dopo l\'aggiunta apri Excel, aggiorna Refinitiv, poi Sincronizza.';
  } else {
    if(lbl)lbl.textContent='Aggiungi ticker:';
    if(inp)inp.placeholder='Es. AAPL.O, MSFT.OQ, ENI.MI…';
    if(note)note.innerHTML='Ticker Refinitiv con suffisso exchange — es. <code style="color:#5ed68a">.O</code> Nasdaq · <code style="color:#5ed68a">.OQ</code> OTC · <code style="color:#5ed68a">.MI</code> Milano · <code style="color:#5ed68a">.AS</code> Amsterdam.';
  }
  // Reset filters when switching tabs
  const srch=document.getElementById('sdb-search');if(srch)srch.value='';
  const filEl=document.getElementById('sdb-filter-area');if(filEl)filEl.innerHTML='';
  sdbRender();
}

function sdbRender(){
  const q=(document.getElementById('sdb-search')?.value||'').toLowerCase();
  const isEtf=_sdbTab==='etfs';
  const data=isEtf?_etfUniverse:_stockUniverse;

  // Build filter dropdowns in the header bar
  const filterArea=document.getElementById('sdb-filter-area');
  const SEL_STYLE='padding:5px 12px;font-size:11px;border:1.5px solid rgba(255,255,255,.2);border-radius:999px;background:rgba(255,255,255,.09);color:#fff;outline:none;cursor:pointer;max-width:190px;';

  // Read current values before rebuilding
  const prevTipo=document.getElementById('sdb-f-tipo')?.value||'';
  const prevSector=document.getElementById('sdb-f-sector')?.value||'';

  if(filterArea){
    filterArea.innerHTML='';

    if(isEtf){
      // ── TIPO first (not cascaded — always shows all types) ────────────────
      const tipos=[...new Set(data.map(s=>s.type||'').filter(Boolean))].sort();
      const stTipo=document.createElement('select');stTipo.id='sdb-f-tipo';stTipo.style.cssText=SEL_STYLE;
      stTipo.innerHTML=`<option value="">Tutti i tipi</option>`+tipos.map(t=>`<option value="${t}"${t===prevTipo?' selected':''}>${t}</option>`).join('');
      // When tipo changes → repopulate sector AND re-render
      stTipo.onchange=()=>{
        // Cascade: reset sector if the current selection no longer exists for this tipo
        const newTipo=stTipo.value;
        const validSectors=new Set(data.filter(s=>!newTipo||s.type===newTipo).map(s=>s.sector||'').filter(Boolean));
        const ss=document.getElementById('sdb-f-sector');
        if(ss&&ss.value&&!validSectors.has(ss.value))ss.value='';
        sdbRender();
      };
      filterArea.appendChild(stTipo);

      // ── SETTORE — cascades off Tipo ────────────────────────────────────────
      const activeTipoNow=stTipo.value;
      const dataForSectors=activeTipoNow?data.filter(s=>s.type===activeTipoNow):data;
      const sectors=[...new Set(dataForSectors.map(s=>s.sector||'').filter(Boolean))].sort();
      const validSector=sectors.includes(prevSector)?prevSector:'';
      const stSector=document.createElement('select');stSector.id='sdb-f-sector';stSector.style.cssText=SEL_STYLE;
      stSector.innerHTML=`<option value="">Tutti i settori</option>`+sectors.map(s=>`<option value="${s}"${s===validSector?' selected':''}>${s}</option>`).join('');
      stSector.onchange=sdbRender;
      filterArea.appendChild(stSector);

    } else {
      // ── AZIONI: only Settore ──────────────────────────────────────────────
      const sectors=[...new Set(data.map(s=>s.industry||'').filter(Boolean))].sort();
      const stSector=document.createElement('select');stSector.id='sdb-f-sector';stSector.style.cssText=SEL_STYLE;
      stSector.innerHTML=`<option value="">Tutti i settori</option>`+sectors.map(s=>`<option value="${s}"${s===prevSector?' selected':''}>${s}</option>`).join('');
      stSector.onchange=sdbRender;
      filterArea.appendChild(stSector);
    }

    // Reset button — only visible when a filter is active
    const activeTipoNow=document.getElementById('sdb-f-tipo')?.value||'';
    const activeSectorNow=document.getElementById('sdb-f-sector')?.value||'';
    if(activeTipoNow||activeSectorNow){
      const clr=document.createElement('button');clr.textContent='✕';clr.title='Reset filtri';
      clr.style.cssText='padding:4px 11px;font-size:11px;background:rgba(255,80,80,.15);border:1.5px solid rgba(255,80,80,.35);border-radius:999px;color:rgba(255,180,180,.9);cursor:pointer;';
      clr.onclick=()=>{
        const st=document.getElementById('sdb-f-tipo');if(st)st.value='';
        const ss=document.getElementById('sdb-f-sector');if(ss)ss.value='';
        sdbRender();
      };
      filterArea.appendChild(clr);
    }
  }

  // Read final active filter values (after rebuilding selects)
  const activeSector=document.getElementById('sdb-f-sector')?.value||'';
  const activeTipo=document.getElementById('sdb-f-tipo')?.value||'';

  const filtered=data.filter(s=>{
    const matchQ=!q||(isEtf
      ?(s.etf_name||'').toLowerCase().includes(q)||(s.ticker||'').toLowerCase().includes(q)||(s.sector||'').toLowerCase().includes(q)
      :(s.company||'').toLowerCase().includes(q)||(s.ticker||'').toLowerCase().includes(q)||(s.industry||'').toLowerCase().includes(q));
    const matchSector=!activeSector||(isEtf?s.sector:s.industry)===activeSector;
    const matchTipo=!activeTipo||s.type===activeTipo;
    return matchQ&&matchSector&&matchTipo;
  });
  const cnt=document.getElementById('sdb-count');
  if(cnt)cnt.textContent=`${filtered.length} ${isEtf?'ETF':'titoli'}`;
  const body=document.getElementById('sdb-body');if(!body)return;

  let COLS,WIDTHS,grid,html='';
  if(isEtf){
    COLS=['Nome ETF','Settore','Tipo','Prezzo','YTD %','1Y %','Mkt Cap','Borsa','TER %',''];
    WIDTHS=['2fr','1fr','80px','90px','72px','72px','110px','100px','80px','46px'];
  } else {
    COLS=['Ticker','Società','Settore / Industry','Prezzo','YTD %','1Y %','Market Cap','Borsa',''];
    WIDTHS=['90px','1fr','1fr','90px','72px','72px','120px','110px','46px'];
  }
  grid=WIDTHS.join(' ');
  html=`<div style="display:grid;grid-template-columns:${grid};background:#0f1e35;padding:0 12px;position:sticky;top:0;z-index:10;border-bottom:2px solid #c9a84c;">`;
  COLS.forEach(h=>{html+=`<span style="padding:9px 6px;font-size:9px;font-weight:700;color:rgba(255,255,255,.65);text-transform:uppercase;letter-spacing:.6px;">${h}</span>`;});
  html+='</div>';

  if(!filtered.length){
    html+=`<div style="padding:40px;text-align:center;color:var(--txt3);font-size:13px;">Nessun ${isEtf?'ETF':'titolo'} nel database. Premi "Sincronizza da Excel" per importare.</div>`;
  } else {
    filtered.forEach((s,i)=>{
      const bg=i%2===0?'#12213f':'#152443';
      const ytd=s.ytd_pct!=null?(s.ytd_pct*100).toFixed(2)+'%':'—';
      const ytdColor=s.ytd_pct!=null?(s.ytd_pct>=0?'#7fd6a4':'#ef6a5a'):'var(--txt3)';
      const one_y=s.one_y_pct!=null?(s.one_y_pct*100).toFixed(2)+'%':'—';
      const price=s.price!=null?s.price.toLocaleString('it-IT',{minimumFractionDigits:2,maximumFractionDigits:2}):'—';
      if(isEtf){
        const mcap=s.mkt_cap!=null?'€'+(s.mkt_cap/1e9).toFixed(1)+'B':'—';
        const isPh=(s.etf_name||'').startsWith('[');
        const displayName=isPh?'⏳ In attesa Refinitiv…':(s.etf_name&&s.etf_name!==s.ticker?s.etf_name:s.ticker);
        html+=`<div style="display:grid;grid-template-columns:${grid};padding:0 12px;background:${bg};border-bottom:1px solid rgba(255,255,255,.06);align-items:center;">
          <span style="padding:8px 6px;font-size:11px;font-weight:600;color:${isPh?'#e8c060':'var(--txt)'};">${displayName}</span>
          <span style="padding:8px 6px;font-size:10px;color:var(--txt2);">${s.sector||'—'}</span>
          <span style="padding:8px 6px;font-size:10px;color:var(--txt3);">${s.type||'—'}</span>
          <span style="padding:8px 6px;font-size:11px;font-family:'JetBrains Mono',monospace;color:var(--txt);text-align:right;">${price}</span>
          <span style="padding:8px 6px;font-size:11px;font-weight:700;color:${ytdColor};text-align:right;">${ytd}</span>
          <span style="padding:8px 6px;font-size:11px;color:var(--txt2);text-align:right;">${one_y}</span>
          <span style="padding:8px 6px;font-size:10px;color:var(--txt3);text-align:right;">${mcap}</span>
          <span style="padding:8px 6px;font-size:10px;color:var(--txt3);">${s.exchange||'—'}</span>
          <span style="padding:8px 6px;font-size:11px;font-family:'JetBrains Mono',monospace;color:var(--txt2);text-align:right;">${s.ter!=null?(s.ter*100).toFixed(2)+'%':'—'}</span>
          <span style="padding:4px 6px;text-align:center;"><button onclick="sdbDeleteEtf(${s.id},'${(s.ticker||'').replace(/'/g,"\\'")}')" title="Rimuovi" style="background:none;border:none;color:rgba(200,50,50,.4);cursor:pointer;font-size:13px;padding:4px 6px;border-radius:4px;" onmouseover="this.style.color='rgba(200,50,50,.9)'" onmouseout="this.style.color='rgba(200,50,50,.4)'">✕</button></span>
        </div>`;
      } else {
        const mcap=s.market_cap!=null?'$'+(s.market_cap/1e9).toFixed(1)+'B':'—';
        const isPh=(s.company||'').startsWith('[');
        html+=`<div style="display:grid;grid-template-columns:${grid};padding:0 12px;background:${bg};border-bottom:1px solid rgba(255,255,255,.06);align-items:center;">
          <span style="padding:8px 6px;font-size:11px;font-family:'JetBrains Mono',monospace;font-weight:700;color:${isPh?'#e8c060':'#7fd6a4'};">${s.ticker||'—'}</span>
          <span style="padding:8px 6px;font-size:11px;color:${isPh?'#e8c060':'var(--txt)'};">${isPh?'⏳ In attesa Refinitiv…':s.company}</span>
          <span style="padding:8px 6px;font-size:10px;color:var(--txt2);">${s.industry||'—'}</span>
          <span style="padding:8px 6px;font-size:11px;font-family:'JetBrains Mono',monospace;color:var(--txt);text-align:right;">${price}</span>
          <span style="padding:8px 6px;font-size:11px;font-weight:700;color:${ytdColor};text-align:right;">${ytd}</span>
          <span style="padding:8px 6px;font-size:11px;color:var(--txt2);text-align:right;">${one_y}</span>
          <span style="padding:8px 6px;font-size:10px;color:var(--txt3);text-align:right;">${mcap}</span>
          <span style="padding:8px 6px;font-size:10px;color:var(--txt3);">${s.exchange||'—'}</span>
          <span style="padding:4px 6px;text-align:center;"><button onclick="sdbDelete(${s.id},'${(s.company||'').replace(/'/g,"\\'")}')" title="Rimuovi" style="background:none;border:none;color:rgba(200,50,50,.4);cursor:pointer;font-size:13px;padding:4px 6px;border-radius:4px;" onmouseover="this.style.color='rgba(200,50,50,.9)'" onmouseout="this.style.color='rgba(200,50,50,.4)'">✕</button></span>
        </div>`;
      }
    });
  }
  body.innerHTML=html;
}

async function sdbSync(){
  const btn=document.getElementById('sdb-sync-btn');
  if(btn){btn.textContent='…';btn.disabled=true;}
  const isEtf=_sdbTab==='etfs';
  const url=isEtf?'/api/etf-universe/sync':'/api/stock-universe/sync';
  try{
    const r=await fetch(url,{method:'POST'});const d=await r.json();
    if(isEtf){await loadEtfUniverse();if(getPD('ammEtf'))buildTab('ammEtf');}
    else{await loadStockUniverse();if(getPD('ammStock'))buildTab('ammStock');}
    sdbRender();
    if(typeof updateSyncDot==='function')updateSyncDot();
    toast(`✓ ${d.synced} ${isEtf?'ETF':'titoli'} sincronizzati`,'ok');
  }catch(e){toast('Errore sincronizzazione','err');}
  finally{if(btn){btn.innerHTML=ic('refresh')+' Sincronizza da Excel';btn.disabled=false;}}
}

async function sdbAddTicker(){
  const inp=document.getElementById('sdb-new-ticker');
  const msg=document.getElementById('sdb-add-msg');
  const val=(inp?.value||'').trim().toUpperCase();
  if(!val){if(msg)msg.textContent='Inserisci un valore';return;}
  if(msg)msg.textContent='Aggiunta in corso…';
  const isEtf=_sdbTab==='etfs';
  const url=isEtf?'/api/etf-universe':'/api/stock-universe';
  const body=isEtf?{etf_id:val}:{ticker:val};
  try{
    const r=await fetch(url,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)});
    const d=await r.json();if(!r.ok)throw new Error(d.error||'Errore');
    if(inp)inp.value='';
    if(isEtf){await loadEtfUniverse();}else{await loadStockUniverse();}
    sdbRender();
    if(msg)msg.textContent=`✓ ${val} aggiunto — apri Excel e lascia aggiornare Refinitiv`;
  }catch(e){if(msg)msg.textContent='Errore: '+(e.message||'');}
}

async function sdbDelete(id,company){
  if(!confirm(`⚠️ ELIMINAZIONE PERMANENTE\n\n"${company}" verrà rimosso definitivamente dal database e dal file Excel.\nL'operazione NON è reversibile.\n\nConfermi la rimozione?`))return;
  try{
    const r=await fetch(`/api/stock-universe/${id}`,{method:'DELETE'});
    const d=await r.json();if(!r.ok)throw new Error(d.error||'Errore');
    await loadStockUniverse();sdbRender();
    if(getPD('ammStock'))buildTab('ammStock');
    toast(`"${company}" rimosso`,'info');
  }catch(e){toast('Errore nella rimozione','err');}
}

async function sdbDeleteEtf(id,ticker){
  if(!confirm(`⚠️ ELIMINAZIONE PERMANENTE\n\n"${ticker}" verrà rimosso definitivamente dal database e dal file Excel.\nL'operazione NON è reversibile.\n\nConfermi la rimozione?`))return;
  try{
    const r=await fetch(`/api/etf-universe/${id}`,{method:'DELETE'});
    const d=await r.json();if(!r.ok)throw new Error(d.error||'Errore');
    await loadEtfUniverse();sdbRender();
    if(getPD('ammEtf'))buildTab('ammEtf');
    toast(`"${ticker}" rimosso`,'info');
  }catch(e){toast('Errore nella rimozione','err');}
}

