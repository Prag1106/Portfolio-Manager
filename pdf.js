// ── PDF EXPORT ────────────────────────────────────────────────────────────────
function hexToRgb(hex){return[parseInt(hex.slice(1,3),16),parseInt(hex.slice(3,5),16),parseInt(hex.slice(5,7),16)];}

function computeAcCtrvMap(pt,pd,total){
  const weights=pd.weights||{};const order=pd._order||Object.keys(weights);
  const map={};
  const allRows=getRows(pt);
  order.forEach(ac=>{
    const acd=weights[ac];if(!acd||ac==='START & GO')return;
    // Somma di ogni fondo: CTRV inserito se presente, altrimenti quota del sotto-settore
    // (stessa regola dell'editor). Prima i fondi senza CTRV venivano ignorati se un altro
    // fondo della classe ne aveva uno.
    let hasAbs=false,sum=0;
    Object.keys(allRows).filter(k=>k.startsWith(`${pt}|${ac}|`)).forEach(k=>{
      const rd=getRow(k,pt);if(!rd.isin&&!rd.name)return;
      if(rd.ctrvAbs!=null)hasAbs=true;
      const sub=k.split('|')[2];
      const sv=total*(acd.pct/100)*((acd.subs?.[sub]||0)/100);
      const c=derive(rd,sv).ctrv;if(c>0)sum+=c;
    });
    map[ac]=hasAbs?sum:total*(acd.pct/100);
  });
  return map;
}


// DARTA with Easy Switch on but no START & GO weight: the uninvested
// residual (total − funds) is what feeds Easy Switch — show it as a slice.
function pdfEsResidual(pt,total,acMap,sgV){
  if(pt!=='darta'||sgV>0||!CD?.portfolios?.darta?.pac?.enabled)return 0;
  const inv=Object.values(acMap).reduce((s,v)=>s+v,0);
  return Math.max(0,Math.round(total-inv));
}

// Allianz corporate palette (same values as the PowerPoint deck in pptx.js).
const PDF_C={blue:[0,55,129],navy:[31,50,87],azure:[0,123,181],cyan:'#26B5E3',body:[89,89,89],muted:[128,128,128],
  panel:[240,240,240],hair:[217,217,217],totRow:[220,227,238],white:[255,255,255]};
const PDF_AC_HEX={EQUITY:'#8A679C',BOND:'#007D8C',COMMODITY:'#FAB709','START & GO':'#26B5E3','EASY SWITCH':'#26B5E3'};
const PDF_EXTRA_HEX=['#EB6218','#6BBB87','#007BB5','#003781'];
const _pdfAcMap={};let _pdfAcIdx=0;
function pdfAcColor(ac){if(PDF_AC_HEX[ac])return PDF_AC_HEX[ac];if(!_pdfAcMap[ac])_pdfAcMap[ac]=PDF_EXTRA_HEX[(_pdfAcIdx++)%PDF_EXTRA_HEX.length];return _pdfAcMap[ac];}
const pdfAcLabel=ac=>ac==='START & GO'?'Start & Go':ac==='EASY SWITCH'?'Easy Switch':ac.charAt(0)+ac.slice(1).toLowerCase();
const _xE=v=>'€'+Math.round(v||0).toLocaleString('it-IT',{useGrouping:'always'});
const _xE2=v=>'€'+(v||0).toLocaleString('it-IT',{minimumFractionDigits:2,maximumFractionDigits:2,useGrouping:'always'});
const _xC=f=>((f||0)*100).toLocaleString('it-IT',{minimumFractionDigits:2,maximumFractionDigits:2})+'%';

// Doughnut in the deck style: white separators, white bold % on each segment.
async function makeDonutImage(labels,vals,colors,size=240){
  return new Promise(resolve=>{
    const cv=document.createElement('canvas');cv.width=size;cv.height=size;
    const tot=vals.reduce((s,v)=>s+(v||0),0)||1;
    const pctLabels={id:'pctLabels',afterDatasetsDraw(ch){
      const c=ch.ctx;c.save();c.fillStyle='#fff';c.textAlign='center';c.textBaseline='middle';
      c.font=`bold ${Math.round(size*0.055)}px Arial, Helvetica, sans-serif`;
      ch.getDatasetMeta(0).data.forEach((arc,i)=>{const p=(vals[i]||0)/tot;if(p<0.05)return;
        const{x,y}=arc.tooltipPosition();c.fillText(Math.round(p*100)+'%',x,y);});
      c.restore();}};
    const ch=new Chart(cv,{type:'doughnut',data:{labels,datasets:[{data:vals,backgroundColor:colors,borderWidth:Math.max(2,size*0.008),borderColor:'#fff'}]},
      options:{animation:false,responsive:false,devicePixelRatio:1,plugins:{legend:{display:false},tooltip:{enabled:false}},cutout:'62%'},plugins:[pctLabels]});
    setTimeout(()=>{resolve(cv.toDataURL('image/png'));ch.destroy();},60);
  });
}

async function _pdfLogo(){
  try{const r=await fetch('/static/img/allianz_bank_fa_horizontal_white.png');if(!r.ok)return null;const b=await r.blob();
    return await new Promise(res=>{const fr=new FileReader();fr.onload=()=>res(fr.result);fr.onerror=()=>res(null);fr.readAsDataURL(b);});}
  catch(e){return null;}
}

async function downloadPdf(opts={}){
  let _curPt=null;
  // Per-type column lookup (uses _curPt set inside the section loop)
  const o=k=>{const t=opts&&opts.types&&opts.types[_curPt];if(!t)return true;const c=t.cols||{};return (k in c)?c[k]!==false:true;};
  // Document-level option lookup (donuts etc.)
  const docO=k=>(opts&&opts.doc)?opts.doc[k]!==false:true;
  // Whether a portfolio type is included at all
  const inc=pt=>!(opts&&opts.types&&opts.types[pt]&&opts.types[pt].include===false);
  if(!AID||!CD){toast('Nessun cliente selezionato','err');return;}
  toast('Generando PDF…','info');
  await savePtfls(true);
  const logo=await _pdfLogo();
  const{jsPDF}=window.jspdf;
  const doc=new jsPDF({orientation:'p',unit:'mm',format:'a4',compress:true});
  const PW=210,PH=297,M=14,CW=PW-2*M;
  const FOOT=13;                 // navy footer band height
  const TOP=22,BOT=PH-FOOT-6;    // content area on every page
  const{blue:BLUE,navy:NAVY,azure:AZURE,body:BODY,muted:MUTED,panel:PANEL,hair:HAIR,totRow:TOTROW}=PDF_C;
  const today=new Date().toLocaleDateString('it-IT',{day:'2-digit',month:'long',year:'numeric'});
  const fundName=raw=>{if(typeof pptFundName!=='function')return raw||'—';const f=pptFundName(raw);return [f.line,f.name].filter(Boolean).join(' ')||'—';};

  // Page furniture is drawn at the end (header + footer on every page).
  function furniture(i,n){
    doc.setPage(i);
    doc.setFont('helvetica','normal');doc.setFontSize(8);doc.setTextColor(...BLUE);
    doc.text(`Portafoglio ${CD.name}`,M,10);
    doc.text(today,PW-M,10,{align:'right'});
    doc.setDrawColor(...BLUE);doc.setLineWidth(0.3);doc.line(M,13,PW-M,13);
    doc.setFillColor(...NAVY);doc.rect(0,PH-FOOT,PW,FOOT,'F');
    const lw=42,lh=lw*86/603;
    if(logo)doc.addImage(logo,'PNG',PW-M-lw,PH-FOOT+(FOOT-lh)/2,lw,lh,'allianz-logo','FAST');
    doc.setFontSize(7);doc.setTextColor(255,255,255);
    doc.text(`Documento riservato  ·  ${CD.name}`,M,PH-FOOT/2+1.2);
    doc.setTextColor(170,180,200);
    doc.text(`${i} / ${n}`,PW-M-lw-8,PH-FOOT/2+1.2,{align:'right'});
  }
  let y=TOP;
  const newPage=()=>{doc.addPage();y=TOP;};
  const fits=h=>{if(y+h>BOT)newPage();};

  // Cover title on page 1
  doc.setFont('helvetica','normal');doc.setTextColor(...BLUE);doc.setFontSize(24);
  doc.text('Portafoglio '+CD.name,M,y+8);
  doc.setFontSize(11);doc.setTextColor(...AZURE);
  doc.text('Riepilogo dei portafogli e struttura commissionale',M,y+15);
  y+=24;

  // Section title: blue heading, optional right-hand figure, blue rule.
  function sectionTitle(title,right){
    doc.setFont('helvetica','normal');doc.setFontSize(15);doc.setTextColor(...BLUE);
    doc.text(title,M,y+6);
    if(right){doc.setFont('helvetica','bold');doc.setFontSize(12);doc.text(right,PW-M,y+6,{align:'right'});}
    doc.setDrawColor(...BLUE);doc.setLineWidth(0.4);doc.line(M,y+9,PW-M,y+9);
    y+=13;
  }
  // Total + caption in the donut hole.
  function donutCentre(cx,cy,ds,value,caption){
    doc.setFont('helvetica','normal');doc.setTextColor(...BLUE);
    let fs=ds*0.42;doc.setFontSize(fs);
    while(fs>6&&doc.getTextWidth(value)>ds*0.52){fs-=0.5;doc.setFontSize(fs);}
    doc.text(value,cx,cy+(caption?0:fs*0.12),{align:'center'});
    if(caption){doc.setFontSize(Math.max(6,fs*0.55));doc.setTextColor(...MUTED);doc.text(caption,cx,cy+fs*0.45,{align:'center'});}
  }
  // Common autoTable styling (Allianz grid table).
  const tableBase=(fs,pad)=>({theme:'grid',margin:{left:M,right:M,top:TOP,bottom:PH-BOT},
    styles:{font:'helvetica',fontSize:fs,cellPadding:pad,textColor:BODY,lineColor:HAIR,lineWidth:0.15,overflow:'linebreak',valign:'middle'},
    headStyles:{fillColor:BLUE,textColor:[255,255,255],fontStyle:'bold',fontSize:fs,cellPadding:pad+0.4,lineColor:BLUE,valign:'bottom'},
    bodyStyles:{fillColor:[255,255,255]}});
  const totStyle=cell=>{cell.styles.fillColor=TOTROW;cell.styles.textColor=BLUE;cell.styles.fontStyle='bold';};

  // Annual cost (CG+SC+Cop. on CTRV) and invested CTRV of a DARTA-style
  // portfolio's funds — START & GO is a cash residual and carries no cost.
  function dartaFundCost(pt){
    const pd=CD.portfolios[pt];if(!pd)return{ctrv:0,costo:0};
    const total=getTotal(pt);if(!(total>0))return{ctrv:0,costo:0};
    const weights=pd.weights||{};const order=pd._order||Object.keys(weights);
    const copBase=pt==='pac'?(CD?.portfolios?.darta?.total||total):total;
    const copRate=getDartaCopertura(copBase);
    const allRows=getRows(pt);
    let ctrvSum=0,costoSum=0;
    order.forEach(ac=>{
      const acd=weights[ac];if(!acd||ac==='START & GO')return;
      Object.entries(acd.subs||{}).forEach(([sub,sp100])=>{
        const sv=total*(acd.pct/100)*(sp100/100);
        const type=(ac==='BOND'||acd._type==='bond')?'Obbligazione':'Azione';
        Object.keys(allRows).filter(k=>k.startsWith(`${pt}|${ac}|${sub}|${type}|`)).forEach(k=>{
          const rd=getRow(k,pt);if(!rd.isin&&!rd.name)return;
          const{ctrv}=derive(rd,sv);const rctrv=ctrv>0?ctrv:0;
          const commTot=parsePctStr(rd.cg)+parsePctStr(rd.sc)+copRate;
          ctrvSum+=rctrv;if(rctrv>0&&commTot>0)costoSum+=rctrv*commTot;
        });
      });
    });
    return{ctrv:ctrvSum,costo:costoSum};
  }

  // ── Overview donut row ────────────────────────────────────────────
  if(docO('overview')){
    const activePts=Object.entries(CD.portfolios).filter(([pt,pd])=>{
      if(!pd||!inc(pt))return false;
      if(pt==='ammEtf'||pt==='ammStock')return(pd.rows||[]).reduce((s,r)=>s+(r.ctrv||0),0)>0;
      return(getTotal(pt))>0;
    });
    if(activePts.length){
      sectionTitle('Riepilogo portafogli');
      const spacing=CW/activePts.length,DS=Math.min(44,spacing-8);
      const imgs=await Promise.all(activePts.map(async([pt,pd])=>{
        const isFlatAmm2=pt==='ammEtf'||pt==='ammStock';
        const total=isFlatAmm2?(pd.rows||[]).reduce((s,r)=>s+(r.ctrv||0),0):(getTotal(pt));
        if(isFlatAmm2){
          // For flat amm: single segment donut
          return{pt,total,img:await makeDonutImage([TL[pt]||pt],[100],['#007BB5'],500)};
        }
        const weights=pd.weights||{};const order=pd._order||Object.keys(weights);
        const acCtrvM=computeAcCtrvMap(pt,pd,total);
        const sgV=(pt==='darta'||pt==='sicav')?Math.round(total*((weights['START & GO']?.pct||0)/100)):0;
        const esV=pdfEsResidual(pt,total,acCtrvM,sgV);
        const grandAc=(Object.values(acCtrvM).reduce((s,v)=>s+v,0)+sgV+esV)||total;
        const L=[],V=[],C=[];
        order.forEach(ac=>{const d=weights[ac];if(!d)return;
          if(ac==='START & GO'){if(sgV>0){L.push(ac);V.push(grandAc>0?sgV/grandAc*100:(d.pct||0));C.push(pdfAcColor(ac));}return;}
          const cv=acCtrvM[ac]||0;L.push(ac);V.push(grandAc>0?cv/grandAc*100:(d.pct||0));C.push(pdfAcColor(ac));});
        if(esV>0){L.push('EASY SWITCH');V.push(esV/grandAc*100);C.push(pdfAcColor('EASY SWITCH'));}
        return{pt,total,img:await makeDonutImage(L,V,C,500)};
      }));
      imgs.forEach((d,i)=>{
        const cx=M+i*spacing+spacing/2;
        doc.addImage(d.img,'PNG',cx-DS/2,y,DS,DS,undefined,'FAST');
        donutCentre(cx,y+DS/2,DS,fmt(d.total),null);
        doc.setTextColor(...BLUE);doc.setFontSize(9);doc.setFont('helvetica','bold');
        doc.text(TL[d.pt]||d.pt,cx,y+DS+6,{align:'center'});
      });
      y+=DS+14;
    }
  }

  // helper to build column spec for autoTable
  function buildPdfColSpec(isDartaPdf,isPacPdf,isAmmPdf,isPacLikePdf){
    const cols=[];const nameCs={cellWidth:'auto',fontStyle:'bold',textColor:BLUE};
    if(isDartaPdf){
      // Narrow numeric/% columns so the 'auto' name column gets the leftover width.
      const sw=isPacPdf?23:26;
      cols.push({key:'settore',h:'Settore',cs:{cellWidth:sw}});
      if(o('isin')){const iw=isPacPdf?19:21;cols.push({key:'isin',h:'ISIN',cs:{cellWidth:iw,fontSize:isPacPdf?5.5:6}});}
      cols.push({key:'name',h:'Fondo',cs:nameCs});
      if(o('ctrv')){const ctw=isPacPdf?16:18;cols.push({key:'ctrv',h:isPacPdf?'CTRV / mese':'CTRV',cs:{cellWidth:ctw,halign:'right'}});}
      if(o('pctAc')){const paw=isPacPdf?11:13;cols.push({key:'pct_ac',h:'% classe',cs:{cellWidth:paw,halign:'right'}});}
      if(o('cg')){const cgw=isPacPdf?10:11;cols.push({key:'cg',h:'CG',cs:{cellWidth:cgw,halign:'right'}});}
      if(o('sc')){const scw=isPacPdf?10:11;cols.push({key:'sc',h:'SC',cs:{cellWidth:scw,halign:'right'}});}
      if(o('cop')){const cpw=isPacPdf?10:11;cols.push({key:'cop',h:'Cop.',cs:{cellWidth:cpw,halign:'right'}});}
      if(o('commTot')){const cmw=isPacPdf?12:13;cols.push({key:'comm_tot',h:'Comm. tot.',cs:{cellWidth:cmw,halign:'right'}});}
      if(o('costo')){const csw=isPacPdf?16:18;cols.push({key:'costo',h:isPacPdf?'Costo annuo / rata':'Costo annuo',cs:{cellWidth:csw,halign:'right'}});}
      if(isPacPdf&&o('costoTot')){cols.push({key:'costo_tot',h:'Costo annuo a regime',cs:{cellWidth:17,halign:'right'}});}
    }else{
      cols.push({key:'settore',h:'Settore',cs:{cellWidth:27}});
      if(o('isin')){cols.push({key:'isin',h:'ISIN',cs:{cellWidth:22,fontSize:6.5}});}
      cols.push({key:'name',h:'Fondo',cs:nameCs});
      if(o('ctrv')){cols.push({key:'ctrv',h:'CTRV',cs:{cellWidth:18,halign:'right'}});}
      if(o('pctAc')){cols.push({key:'pct_ac',h:'% classe',cs:{cellWidth:15,halign:'right'}});}
      if(isAmmPdf&&o('ter')){cols.push({key:'ter',h:'TER',cs:{cellWidth:12,halign:'right'}});}
      if(isAmmPdf&&o('costo')){cols.push({key:'costo',h:'Costo annuo',cs:{cellWidth:17,halign:'right'}});}
      if(!isAmmPdf&&o('cg')){cols.push({key:'cg',h:'CG',cs:{cellWidth:11,halign:'right'}});}
      if(!isAmmPdf&&o('costo')){cols.push({key:'costo',h:isPacLikePdf?'Costo annuo / rata':'Costo annuo',cs:{cellWidth:17,halign:'right'}});}
      if(!isAmmPdf&&isPacLikePdf&&o('costoTot')){cols.push({key:'costo_tot',h:'Costo annuo a regime',cs:{cellWidth:17,halign:'right'}});}
      if(o('pctPort')){cols.push({key:'pct_port',h:'% port.',cs:{cellWidth:14,halign:'right'}});}
      if(!isAmmPdf){cols.push({key:'price',h:'Prezzo',cs:{cellWidth:15,halign:'right'}});}
      if(o('qty')){cols.push({key:'qty',h:'Qtà',cs:{cellWidth:14,halign:'right'}});}
    }
    return cols;
  }

  const pctCell=s=>{const v=parsePctStr(s);return s!=null&&String(s).trim()!==''?_xC(v):'–';};
  function buildPdfDataRow(colSpec,rd,sv,total,isDartaPdf,isPacPdf,copRatePdf,pacMesiPdf,isAmmPdf,acCtrv){
    const{ctrv,qty}=derive(rd,sv);const rctrv=ctrv>0?ctrv:0;
    const commTotV=isDartaPdf?parsePctStr(rd.cg)+parsePctStr(rd.sc)+copRatePdf:0;
    const terRatePdf=isAmmPdf?parsePctStr(rd.cg):0;
    const cgRatePdf=(!isDartaPdf&&!isAmmPdf)?parsePctStr(rd.cg):0;
    const costoV=isDartaPdf?(rctrv>0&&commTotV>0?rctrv*commTotV:0):
                 isAmmPdf?(rctrv>0&&terRatePdf>0?rctrv*terRatePdf:0):
                 (rctrv>0&&cgRatePdf>0?rctrv*cgRatePdf:0);
    return colSpec.map(col=>{
      switch(col.key){
        case 'settore':return String(rd._sub||'—').replace(/^Obbligazioni/i,'Obbl.');
        case 'isin':return rd.isin||'–';
        case 'name':return fundName(rd.name);
        case 'ctrv':return rctrv>0?'€'+rctrv.toLocaleString('it-IT',{minimumFractionDigits:0,maximumFractionDigits:2,useGrouping:'always'}):'–';
        case 'cg':return pctCell(rd.cg);
        case 'sc':return pctCell(rd.sc);
        case 'cop':return _xC(copRatePdf);
        case 'comm_tot':return commTotV>0?_xC(commTotV):'–';
        case 'costo':return costoV>0?_xE2(costoV):'–';
        case 'costo_tot':return costoV>0?_xE2(costoV*pacMesiPdf):'–';
        case 'ter':return pctCell(rd.cg);
        case 'pct_port':return total>0&&rctrv>0?_xC(rctrv/total):'–';
        case 'pct_ac':return acCtrv>0&&rctrv>0?_xC(rctrv/acCtrv):'–';
        case 'price':return rd.price?'€'+rd.price.toLocaleString('it-IT',{minimumFractionDigits:2,maximumFractionDigits:2}):'–';
        case 'qty':return qty&&qty>0?parseFloat(qty.toFixed(2)).toLocaleString('it-IT'):'–';
        default:return '–';
      }
    });
  }

  // ── Per-portfolio sections ────────────────────────────────────────
  for(const[pt,pd] of Object.entries(CD.portfolios)){
    if(!pd||!inc(pt))continue;
    _curPt=pt;
    const isFlatAmm=pt==='ammEtf'||pt==='ammStock';
    const flatRows=isFlatAmm?(pd.rows||[]):null;
    const flatTotal=isFlatAmm?flatRows.reduce((s,r)=>s+(r.ctrv||0),0):0;
    const total=isFlatAmm?flatTotal:(getTotal(pt));
    if(total<=0)continue;
    const label=TL[pt]||pt;

    // ── Flat ammEtf / ammStock PDF section ───────────────────────
    if(isFlatAmm){
      const isEtfFlat=pt==='ammEtf';
      fits(30);
      sectionTitle(label,fmt(flatTotal));
      // Build table (column-driven, respects per-type toggles)
      const head=[];const body=[];
      // For ETFs the 'id' (Ticker) column already holds the descriptive name,
      // so the separate 'name' column is a duplicate — drop it to free space.
      const fcols=[{k:'id'}];
      if(!isEtfFlat)fcols.push({k:'name'});
      fcols.push({k:'sector'},{k:'price'});
      if(o('ctrv'))fcols.push({k:'ctrv'});
      if(isEtfFlat&&o('ter'))fcols.push({k:'ter'});
      if(isEtfFlat&&o('costo'))fcols.push({k:'costo'});
      if(o('pctPort'))fcols.push({k:'pctPort'});
      if(o('qty'))fcols.push({k:'qty'});
      const FHEAD={id:isEtfFlat?'ETF':'Ticker',name:'Società',sector:'Settore',price:'Prezzo',ctrv:'Controvalore',ter:'TER',costo:'Costo annuo',pctPort:'% port.',qty:'Quantità'};
      head.push(fcols.map(c=>FHEAD[c.k]));
      let totCosto=0;
      flatRows.forEach(r=>{
        const ref=isEtfFlat?etfByTicker(r.ticker||''):(r.company?stockByCompany(r.company):null);
        const price=ref?.price||0;const ctrv=r.ctrv||0;
        const ter=isEtfFlat?(r.ter!=null?r.ter:(ref?.ter!=null?ref.ter:null)):null;
        const costo=isEtfFlat&&ctrv>0&&ter>0?ctrv*ter:0;
        totCosto+=costo;
        const qty=price>0&&ctrv>0?ctrv/price:null;
        const vals={
          id:r.ticker||'—',
          name:(isEtfFlat?(ref?.etf_name||r.ticker):(r.company||r.ticker))||'—',
          sector:(isEtfFlat?ref?.sector:ref?.industry)||'—',
          price:price>0?'€'+price.toLocaleString('it-IT',{minimumFractionDigits:2,maximumFractionDigits:2}):'–',
          ctrv:ctrv>0?_xE(ctrv):'–',
          ter:ter!=null?_xC(ter):'–',
          costo:costo>0?_xE2(costo):'–',
          pctPort:flatTotal>0&&ctrv>0?_xC(ctrv/flatTotal):'–',
          qty:qty!=null?qty.toLocaleString('it-IT',{minimumFractionDigits:3,maximumFractionDigits:3}):'–',
        };
        body.push(fcols.map(c=>vals[c.k]));
      });
      // Totals row
      const avgTer=flatTotal>0&&totCosto>0?totCosto/flatTotal:null;
      const totVals={id:'Totale',name:'',sector:'',price:'',ctrv:_xE(flatTotal),ter:avgTer!=null?_xC(avgTer):'',
        costo:totCosto>0?_xE2(totCosto):'',pctPort:'100%',qty:''};
      body.push(fcols.map(c=>totVals[c.k]));
      const columnStyles={};const RIGHT=new Set(['price','ctrv','ter','costo','pctPort','qty']);
      fcols.forEach((c,i)=>{columnStyles[i]=RIGHT.has(c.k)?{halign:'right'}:(i===0||c.k==='name')?{fontStyle:'bold',textColor:BLUE}:{};});
      doc.autoTable(Object.assign(tableBase(7.5,2),{startY:y,head,body,columnStyles,
        didParseCell:d=>{if(d.section==='body'&&d.row.index===body.length-1)totStyle(d.cell);}}));
      y=doc.lastAutoTable.finalY+12;
      continue;
    }
    const weights=pd.weights||{};
    const order=pd._order||Object.keys(weights);

    // For darta/sicav, compute effective total from actual ctrv sums
    let effectiveTotalPdf=total;
    if(pt==='darta'||pt==='sicav'){
      let ctrvSumPdf=0;
      order.forEach(ac=>{
        const acd=weights[ac];if(!acd||ac==='START & GO')return;
        Object.entries(acd.subs||{}).forEach(([sub,sp100])=>{
          const sv=total*(acd.pct/100)*(sp100/100);
          const allRows=getRows(pt);
          Object.keys(allRows).filter(k=>k.startsWith(`${pt}|${ac}|${sub}|`)).forEach(k=>{
            const rd=getRow(k,pt);if(!rd.isin&&!rd.name)return;
            const{ctrv}=derive(rd,sv);ctrvSumPdf+=ctrv>0?Math.round(ctrv):0;
          });
        });
      });
      if(pt==='darta'||pt==='sicav'){const sgPct=(weights['START & GO']?.pct||0)/100;ctrvSumPdf+=Math.round(total*sgPct);}
      effectiveTotalPdf=ctrvSumPdf>0?ctrvSumPdf:total;
    }

    // Pre-build donut data using actual CTRV sums (not stored target weights).
    // START & GO is included as its residual controvalore (mirrors the summary view).
    const _acCtrvMap=computeAcCtrvMap(pt,pd,total);
    const _sgValDonut=(pt==='darta'||pt==='sicav')?Math.round(total*((weights['START & GO']?.pct||0)/100)):0;
    const _esValDonut=pdfEsResidual(pt,total,_acCtrvMap,_sgValDonut);
    effectiveTotalPdf+=_esValDonut;
    const _grandAcCtrv=(Object.values(_acCtrvMap).reduce((s,v)=>s+v,0)+_sgValDonut+_esValDonut)||total;
    const L=[],V=[],C=[],VAL=[];
    order.forEach(ac=>{const d=weights[ac];if(!d)return;
      if(ac==='START & GO'){if(_sgValDonut>0){L.push(ac);V.push(_grandAcCtrv>0?_sgValDonut/_grandAcCtrv*100:(d.pct||0));C.push(pdfAcColor(ac));VAL.push(_sgValDonut);}return;}
      const cv=_acCtrvMap[ac]||0;L.push(ac);V.push(_grandAcCtrv>0?cv/_grandAcCtrv*100:(d.pct||0));C.push(pdfAcColor(ac));VAL.push(cv);});
    if(_esValDonut>0){L.push('EASY SWITCH');V.push(_esValDonut/_grandAcCtrv*100);C.push(pdfAcColor('EASY SWITCH'));VAL.push(_esValDonut);}
    const DS=52;
    const neededH=15+(L.length&&docO('sectionDonut')?Math.max(DS,L.length*9+6)+8:0);
    fits(neededH);
    sectionTitle(label,fmt(effectiveTotalPdf));

    // Donut + legend
    if(docO('sectionDonut')&&L.length){
      const img=await makeDonutImage(L,V,C,600);
      const dx=M+4;
      doc.addImage(img,'PNG',dx,y,DS,DS,undefined,'FAST');
      donutCentre(dx+DS/2,y+DS/2,DS,fmt(effectiveTotalPdf),'controvalore');
      const legX=dx+DS+14,legR=Math.min(PW-M,legX+100);
      let legY=y+DS/2-(L.length-1)*9/2;
      L.forEach((lbl,i)=>{
        doc.setFillColor(...hexToRgb(C[i]));doc.circle(legX+1.6,legY-1.2,1.6,'F');
        doc.setTextColor(...BODY);doc.setFontSize(10);doc.setFont('helvetica','normal');
        doc.text(pdfAcLabel(lbl),legX+6,legY);
        const pct=_xC(V[i]/100);doc.setTextColor(...MUTED);doc.text(pct,legR,legY,{align:'right'});
        doc.setFont('helvetica','bold');doc.setTextColor(...BLUE);
        doc.text(fmt(VAL[i]||0),legR-doc.getTextWidth(pct)-4,legY,{align:'right'});
        doc.setDrawColor(...HAIR);doc.setLineWidth(0.15);doc.line(legX,legY+3.2,legR,legY+3.2);
        legY+=9;
      });
      y+=DS+8;
    }

    // Fund tables per asset class
    const isDartaPdf=pt==='darta'||pt==='pac';
    const isPacPdf=pt==='pac';
    const isPacLikePdf=isPacLike(pt);
    const isAmmPdf=pt==='amm'||pt==='ammAz'||pt==='ammEtf'||pt==='ammStock';
    const pacMesiPdf=isPacLikePdf?(pacMesi(pt)||15):0;
    const copBasePdf=pt==='pac'?(CD?.portfolios?.darta?.total||total):total;
    const copRatePdf=isDartaPdf?getDartaCopertura(copBasePdf):0;
    let grandCostoPdf=0,grandCtrvPdf=0;
    order.forEach(ac=>{
      const acd=weights[ac];if(!acd||ac==='START & GO')return;
      const acPct=acd.pct/100;
      // Pre-sum the asset class controvalore (denominator for the % Classe column),
      // so each fund's % within its asset class sums to 100%.
      let acCtrvTotalPdf=0;
      Object.entries(acd.subs||{}).forEach(([sub,sp100])=>{
        const sv=total*acPct*(sp100/100);
        const types=isAmmPdf?(pt==='amm'?['ETF']:['Azione']):(ac==='BOND'||(acd._type==='bond')?['Obbligazione']:['Azione']);
        types.forEach(type=>{
          const allRows=getRows(pt);
          Object.keys(allRows).filter(k=>k.startsWith(`${pt}|${ac}|${sub}|${type}|`)).forEach(k=>{
            const rd=getRow(k,pt);if(!rd.isin&&!rd.name)return;
            const{ctrv}=derive(rd,sv);if(ctrv>0)acCtrvTotalPdf+=ctrv;
          });
        });
      });
      const colSpec=buildPdfColSpec(isDartaPdf,isPacPdf,isAmmPdf,isPacLikePdf);
      const rows=[];let acCostoPdf=0,acCtrvPdf=0;
      Object.entries(acd.subs||{}).forEach(([sub,sp100])=>{
        const sv=total*acPct*(sp100/100);
        const types=isAmmPdf?(pt==='amm'?['ETF']:['Azione']):(ac==='BOND'||(acd._type==='bond')?['Obbligazione']:['Azione']);
        types.forEach(type=>{
          const allRows=getRows(pt);
          const keys=Object.keys(allRows).filter(k=>k.startsWith(`${pt}|${ac}|${sub}|${type}|`));
          keys.forEach(k=>{
            const rd=getRow(k,pt);if(!rd.isin&&!rd.name)return;
            const{ctrv}=derive(rd,sv);const rctrv=ctrv>0?ctrv:0;
            if(isDartaPdf){
              const commTot=parsePctStr(rd.cg)+parsePctStr(rd.sc)+copRatePdf;
              const costo=rctrv>0&&commTot>0?rctrv*commTot:0;
              grandCostoPdf+=costo;acCostoPdf+=costo;
            } else if(!isAmmPdf){
              const cgRate=parsePctStr(rd.cg);
              const costo=rctrv>0&&cgRate>0?rctrv*cgRate:0;
              grandCostoPdf+=costo;acCostoPdf+=costo;
            }
            if(isAmmPdf&&rctrv>0){const ter=parsePctStr(rd.cg);if(ter>0){grandCostoPdf+=rctrv*ter;acCostoPdf+=rctrv*ter;}}
            acCtrvPdf+=rctrv;grandCtrvPdf+=rctrv;
            const rdWithSub=Object.assign({},rd,{_sub:sub});
            rows.push(buildPdfDataRow(colSpec,rdWithSub,sv,effectiveTotalPdf,isDartaPdf,isPacPdf,copRatePdf,pacMesiPdf,isAmmPdf,acCtrvTotalPdf));
          });
        });
      });
      if(!rows.length)return;
      // Class totals row (light-blue, like "Totale switch" in the deck)
      const acComm=acCtrvPdf>0&&acCostoPdf>0?acCostoPdf/acCtrvPdf:0;
      const portW=effectiveTotalPdf>0&&acCtrvPdf>0?acCtrvPdf/effectiveTotalPdf:acPct;
      const totVals={settore:'Totale '+pdfAcLabel(ac).toLowerCase(),isin:'',name:_xC(portW).replace(/,00%$/,'%')+' del portafoglio',
        ctrv:_xE(acCtrvPdf),pct_ac:'100%',cg:'',sc:'',cop:'',comm_tot:acComm?_xC(acComm):'',ter:isAmmPdf&&acComm?_xC(acComm):'',
        costo:acCostoPdf>0?_xE2(acCostoPdf):'',costo_tot:acCostoPdf>0&&isPacLikePdf?_xE2(acCostoPdf*pacMesiPdf):'',
        pct_port:_xC(portW),price:'',qty:''};
      if(!isDartaPdf&&!isAmmPdf)totVals.cg=acComm?_xC(acComm):'';
      rows.push(colSpec.map(c=>totVals[c.key]??''));

      fits(30);
      // Asset-class band: grey panel, class colour dot, blue label
      doc.setFillColor(...PANEL);doc.rect(M,y,CW,8,'F');
      doc.setFillColor(...hexToRgb(pdfAcColor(ac)));doc.circle(M+4,y+4,1.6,'F');
      doc.setTextColor(...BLUE);doc.setFontSize(9);doc.setFont('helvetica','bold');
      doc.text(pdfAcLabel(ac),M+8,y+5.4);
      doc.setFont('helvetica','normal');doc.setFontSize(8);doc.setTextColor(...BODY);
      doc.text(fmt(acCtrvPdf)+'  ·  '+_xC(portW)+' del portafoglio',PW-M-3,y+5.4,{align:'right'});
      y+=8;
      const head=[colSpec.map(c=>c.h)];
      const columnStyles={};colSpec.forEach((c,i)=>{columnStyles[i]=c.cs;});
      const n=colSpec.length;
      const fs=isDartaPdf?(isPacPdf?(n<=4?8:n<=7?7:6.3):(n<=4?8:n<=7?7:6.6)):(n<=4?8:n<=7?7.5:7);
      doc.autoTable(Object.assign(tableBase(fs,isDartaPdf?1.8:2.2),{startY:y,head,body:rows,columnStyles,
        didParseCell:d=>{if(d.section==='head'&&columnStyles[d.column.index]?.halign==='right')d.cell.styles.halign='right';
          if(d.section==='body'&&d.row.index===rows.length-1)totStyle(d.cell);}}));
      y=doc.lastAutoTable.finalY+8;
    });
    if(o('commBar')&&grandCostoPdf>0){
      fits(16);
      doc.setFillColor(...TOTROW);doc.rect(M,y,CW,12,'F');
      doc.setDrawColor(...BLUE);doc.setLineWidth(0.4);doc.line(M,y,PW-M,y);
      doc.setTextColor(...BLUE);doc.setFont('helvetica','bold');doc.setFontSize(9);
      doc.text('Totale '+label,M+3,y+7.6);
      doc.setFontSize(11);doc.text(_xE(grandCtrvPdf),M+3+doc.getTextWidth('Totale '+label)+8,y+7.8);
      const avgPctPdf=grandCtrvPdf>0?'Comm. media '+_xC(grandCostoPdf/grandCtrvPdf)+'   ·   ':'';
      const costoTotTxt=isPacLikePdf?'   (×'+pacMesiPdf+' = '+_xE2(grandCostoPdf*pacMesiPdf)+')':'';
      doc.setFontSize(9);
      // Easy Switch is fed monthly from START & GO, so its total is a monthly cost.
      doc.text(avgPctPdf+(isPacPdf?'Costo mensile Start&Go ':'Costo annuo ')+_xE2(grandCostoPdf)+costoTotTxt,PW-M-3,y+7.6,{align:'right'});
      y+=16;
    }
    // Easy Switch: total DARTA cost once every monthly rata has been moved
    // out of START & GO (DARTA funds + Easy Switch funds × mesi).
    if(isPacPdf&&o('costoPost')&&grandCostoPdf>0&&CD.portfolios.darta){
      const dc=dartaFundCost('darta');
      const pd0=CD.portfolios.darta;
      const sgVal=getTotal('darta')*((pd0.weights?.['START & GO']?.pct||0)/100);
      const swCtrv=grandCtrvPdf*pacMesiPdf,swCosto=grandCostoPdf*pacMesiPdf;
      const postCtrv=dc.ctrv+Math.max(sgVal,swCtrv);
      const postCosto=dc.costo+swCosto;
      const pc=(c,v)=>v>0?_xC(c/v):'—';
      fits(26);
      doc.setFillColor(...PANEL);doc.rect(M,y,CW,22,'F');
      doc.setTextColor(...AZURE);doc.setFontSize(10);doc.setFont('helvetica','normal');
      doc.text('Costo medio ponderato del portafoglio a regime ('+pacMesiPdf+' mesi)',M+5,y+8);
      doc.setFontSize(8);doc.setTextColor(...MUTED);
      doc.text('DARTA '+pc(dc.costo,dc.ctrv)+'  ·  Easy Switch '+pc(swCosto,swCtrv),M+5,y+15);
      doc.setTextColor(...BLUE);doc.setFontSize(24);
      doc.text(pc(postCosto,postCtrv),PW-M-5,y+14.5,{align:'right'});
      y+=26;
    }
    y+=6;
  }

  // ── Header + footer on every page ─────────────────────────────────
  const pc=doc.getNumberOfPages();
  for(let i=1;i<=pc;i++)furniture(i,pc);

  doc.save(`${CD.name.replace(/\s+/g,'_')}_portfolio.pdf`);
  toast('PDF scaricato','ok');
}
