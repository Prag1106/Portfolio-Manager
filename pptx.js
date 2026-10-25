// ── POWERPOINT PROPOSAL (Allianz corporate identity) ─────────────────────────
// Builds the "Proposta di portafoglio" deck for the DARTA policy, mirroring the
// Hasbani reference deck: cover · asset allocation · strumenti · costi ·
// allocazione dopo lo switch · dettaglio DARTA · dettaglio Start & Go.
// Slides that depend on START & GO / Easy Switch are skipped when absent.
const PPTX_LIB='https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js';
const PX={blue:'003781',navy:'1F3257',azure:'007BB5',cyan:'26B5E3',teal:'007D8C',green:'6BBB87',
  yellow:'FAB709',orange:'EB6218',purple:'8A679C',body:'595959',muted:'808080',panel:'F0F0F0',
  hair:'D9D9D9',totRow:'DCE3EE',white:'FFFFFF'};
const PPT_AC_COLOR={EQUITY:PX.purple,BOND:PX.teal,COMMODITY:PX.yellow,'START & GO':PX.cyan};
const PPT_EXTRA_COLORS=[PX.orange,PX.green,PX.azure,PX.blue];

function _loadPptxLib(){
  if(window.PptxGenJS)return Promise.resolve();
  return new Promise((res,rej)=>{const s=document.createElement('script');s.src=PPTX_LIB;s.onload=res;
    s.onerror=()=>rej(new Error('Impossibile caricare la libreria PowerPoint'));document.head.appendChild(s);});
}

// ── Formatting ──
const _pE=v=>'€'+Math.round(v||0).toLocaleString('it-IT',{useGrouping:'always'});
const _pE2=v=>'€'+(v||0).toLocaleString('it-IT',{minimumFractionDigits:2,maximumFractionDigits:2,useGrouping:'always'});
const _pC=f=>((f||0)*100).toLocaleString('it-IT',{minimumFractionDigits:2,maximumFractionDigits:2})+'%';   // 2,20%
const _pW=f=>((f||0)*100).toLocaleString('it-IT',{maximumFractionDigits:2})+'%';                          // 7,2% / 50%
const _pMonth=off=>{const n=new Date();return new Date(n.getFullYear(),n.getMonth()+off,1).toLocaleDateString('it-IT',{month:'long',year:'numeric'});};
const _pMMYY=off=>{const n=new Date();const d=new Date(n.getFullYear(),n.getMonth()+off,1);return String(d.getMonth()+1).padStart(2,'0')+'/'+String(d.getFullYear()).slice(2);};
const _cap=s=>s.charAt(0).toUpperCase()+s.slice(1);

// Fund names as stored ("XCLASSIC T.ROWE PRICE GlbValue") → client-friendly
// ("T. Rowe Price Global Value") plus the product line tag ("XCLASSIC").
const _P_LINES=/^(MKT|XCLASSIC|XTREND|XPLUS|XCHANGE)$/i;
const _P_TOK={ALLIANZGI:'AllianzGI','T.ROWE':'T. Rowe',GLBVALUE:'Global Value',GLB:'Global',EQ:'Equity',MKTS:'Markets',SUST:'Sustainable',TECH:'Technology'};
const _P_KEEP=new Set(['PIMCO','MSCI','BNY','ESG','USA']);
function pptFundName(raw){
  const t=String(raw||'').trim().split(/\s+/).filter(Boolean);let line='';
  if(t.length>1&&_P_LINES.test(t[0]))line=t.shift().toUpperCase();
  const name=t.map(w=>{const u=w.toUpperCase();if(_P_TOK[u])return _P_TOK[u];
    if(/^[A-Z][A-Z.&'-]{3,}$/.test(w)&&!_P_KEEP.has(w))return w.charAt(0)+w.slice(1).toLowerCase();return w;}).join(' ');
  return{line,name:name||String(raw||'—')};
}
function pptSubShort(sub,name,isBond){
  let s=String(sub||'');
  if(/^Azioni\s+/i.test(s))return s.replace(/^Azioni\s+/i,'');
  s=s.replace(/^Obbligazioni\s+/i,'Obbl. ');
  if(isBond&&/ a Scadenza$/i.test(s)){const yrs=String(name||'').match(/20\d\d/g);return s.replace(/ a Scadenza$/i,'')+(yrs?' · Scadenza '+yrs[yrs.length-1]:' · a scadenza');}
  return s.replace(/^Obbl\. /,'');
}
const pptSettore=sub=>String(sub||'').replace(/^Obbligazioni/i,'Obbl.');
function pptAcLabel(ac){if(ac==='START & GO')return 'Start & Go';return ac.charAt(0)+ac.slice(1).toLowerCase();}

// Arial bold ≈ 0.55em per char: shrink a fund name to fit one line, then ellipsise.
function pptFit(text,wIn,fs){const per=c=>wIn*72/(c*0.55);let f=fs;const n=text.length;
  if(per(n)<f)f=Math.max(7,Math.floor(per(n)*2)/2);
  const max=Math.floor(wIn*72/(f*0.55));return{text:n>max?text.slice(0,Math.max(1,max-1)).trimEnd()+'…':text,fs:f};}

// ── Data model (mirrors the PDF/summary maths) ──
function pptCollect(pt,total,copRate){
  const pd=CD.portfolios[pt];const weights=pd.weights||{};const order=pd._order||Object.keys(weights);
  const allRows=getRows(pt);const classes=[];let extra=0;
  order.forEach(ac=>{
    const acd=weights[ac];if(!acd||ac==='START & GO')return;
    const isBond=ac==='BOND'||acd._type==='bond';const type=isBond?'Obbligazione':'Azione';
    const funds=[];
    Object.entries(acd.subs||{}).forEach(([sub,sp100])=>{
      const sv=total*(acd.pct/100)*(sp100/100);
      Object.keys(allRows).filter(k=>k.startsWith(`${pt}|${ac}|${sub}|${type}|`)).forEach(k=>{
        const rd=getRow(k,pt);if(!rd.isin&&!rd.name)return;
        const ctrv=derive(rd,sv).ctrv;if(!(ctrv>0))return;
        const cg=parsePctStr(rd.cg),sc=parsePctStr(rd.sc),comm=cg+sc+copRate;const pn=pptFundName(rd.name);
        funds.push({key:k,sub,isin:rd.isin||'',raw:rd.name||'',name:pn.name,line:pn.line,tag:[pn.line,pptSubShort(sub,rd.name,isBond)].filter(Boolean).join(' · '),
          cg:rd.cg?cg:null,sc:rd.sc?sc:null,cop:copRate,comm,ctrv,cost:ctrv*comm});
      });
    });
    if(!funds.length)return;
    const ctrv=funds.reduce((s,f)=>s+f.ctrv,0),cost=funds.reduce((s,f)=>s+f.cost,0);
    funds.forEach(f=>f.wIn=f.ctrv/ctrv);
    const byW=[...funds].sort((a,b)=>b.ctrv-a.ctrv);
    classes.push({ac,label:pptAcLabel(ac),color:PPT_AC_COLOR[ac]||PPT_EXTRA_COLORS[(extra++)%PPT_EXTRA_COLORS.length],funds,byW,ctrv,cost,comm:cost/ctrv});
  });
  return classes;
}

function pptModel(opts){
  const pd=CD.portfolios.darta;const total=pd.total||0;const copRate=getDartaCopertura(total);
  const classes=pptCollect('darta',total,copRate);
  const sgVal=Math.round(total*((pd.weights?.['START & GO']?.pct||0)/100));
  const cfg=pd.pac||{};const mesi=cfg.mesi||0;
  let pac=[];if(cfg.enabled&&mesi>0&&CD.portfolios.pac)pac=pptCollect('pac',pacTotal('pac'),copRate);
  const srcIsSG=cfg.sourceKey==='__SG__';
  let srcFund=null;
  if(pac.length&&!srcIsSG&&cfg.sourceKey){for(const c of classes){const f=c.funds.find(x=>x.key===cfg.sourceKey);if(f){srcFund={...f,ac:c.ac};break;}}}
  const hasSwitch=pac.length>0&&(srcIsSG?sgVal>0:!!srcFund);
  // Colours: a pac-only class (e.g. COMMODITY bought only via switch) keeps its own colour.
  const used=new Set(classes.map(c=>c.color));
  pac.forEach(c=>{const same=classes.find(x=>x.ac===c.ac);if(same)c.color=same.color;else if(!PPT_AC_COLOR[c.ac]){c.color=PPT_EXTRA_COLORS.find(x=>!used.has(x))||c.color;}used.add(c.color);});
  const rataSum=pac.reduce((s,c)=>s+c.ctrv,0);
  const rata=hasSwitch?pacTotal('pac'):0;
  const swCost=pac.reduce((s,c)=>s+c.cost,0);
  const swComm=rataSum>0?swCost/rataSum:0;
  const delay=hasSwitch?Math.max(0,parseInt(opts.sgDelay,10)||0):0;
  const rate=opts.sgRate!=null&&String(opts.sgRate).trim()!==''?parseFloat(String(opts.sgRate).replace(',','.')):null;
  const srcLabel=srcIsSG||!hasSwitch?'Start & Go':'Easy Switch';
  const srcAmount=hasSwitch?(srcIsSG?sgVal:(cfg.totalSwitch||0)):sgVal;
  const dartaCtrv=classes.reduce((s,c)=>s+c.ctrv,0);
  const denom=dartaCtrv+sgVal||total;
  const dartaCost=classes.reduce((s,c)=>s+c.cost,0);
  // Allocation (as € per bucket) after k switches.
  const order=[...classes.map(c=>c.ac),...pac.map(c=>c.ac).filter(a=>!classes.some(c=>c.ac===a))];
  const colorOf=ac=>(classes.find(c=>c.ac===ac)||pac.find(c=>c.ac===ac)||{}).color||PX.muted;
  function allocAt(k){
    const seg={};order.forEach(ac=>{seg[ac]=(classes.find(c=>c.ac===ac)||{}).ctrv||0;});
    if(sgVal>0)seg['START & GO']=sgVal;
    if(hasSwitch&&k>0){
      if(srcIsSG)seg['START & GO']=Math.max(0,sgVal-rata*k);else seg[srcFund.ac]-=rata*k;
      pac.forEach(c=>{seg[c.ac]=(seg[c.ac]||0)+(rataSum>0?rata*c.ctrv/rataSum:0)*k;});
    }
    return seg;
  }
  const endMonth=delay+mesi;
  const regime=allocAt(hasSwitch?mesi:0);
  const regimeCost=dartaCost+(hasSwitch?swCost*mesi-(srcFund?srcFund.comm*rata*mesi:0):0);
  // Per-class regime ctrv + cost (for "Equity 2,60% · peso 85%").
  const regimeClass={};order.forEach(ac=>{
    const d=classes.find(c=>c.ac===ac),p=pac.find(c=>c.ac===ac);
    let ctrv=(d?.ctrv||0),cost=(d?.cost||0);
    if(hasSwitch&&p){ctrv+=(rataSum>0?rata*p.ctrv/rataSum:0)*mesi;cost+=p.cost*mesi;}
    if(srcFund&&srcFund.ac===ac){ctrv-=rata*mesi;cost-=srcFund.comm*rata*mesi;}
    regimeClass[ac]={ctrv,cost,comm:ctrv>0?cost/ctrv:0,funds:new Set([...(d?.funds||[]),...(hasSwitch&&p?p.funds:[])].map(f=>f.raw)).size};
  });
  return{name:CD.name,total,denom,copRate,classes,pac,sgVal,hasSwitch,srcIsSG,srcFund,srcLabel,srcAmount,rata,mesi,delay,rate,
    swComm,dartaCost,regime,regimeCost,regimeClass,order,colorOf,allocAt,endMonth,
    initialComm:denom>0?dartaCost/denom:0,regimeComm:denom>0?regimeCost/denom:0};
}

// ── Deck ──
async function downloadPptx(opts={}){
  if(!AID||!CD){toast('Nessun cliente selezionato','err');return;}
  if(!CD.portfolios.darta||!(CD.portfolios.darta.total>0)){toast('La presentazione richiede una polizza DARTA','err');return;}
  toast('Generando PowerPoint…','info');
  await savePtfls(true);
  try{await _loadPptxLib();}catch(e){toast(e.message,'err');return;}
  const m=pptModel(opts);
  const pptx=new PptxGenJS();
  pptx.defineLayout({name:'A4L',width:11.69,height:8.27});pptx.layout='A4L';
  pptx.author='Portfolio Manager';pptx.company='Allianz Bank Financial Advisors';pptx.title='Proposta di portafoglio '+m.name;
  const S=pptx.ShapeType;const W=11.69,M=0.6,CW=10.49,F='Arial';
  const today=new Date().toLocaleDateString('it-IT');
  const fonteBase=`Fonte: Portfolio Manager, proposta di portafoglio ${m.name} del ${today}.`;
  const T=(sl,text,o)=>sl.addText(text,Object.assign({fontFace:F,margin:0,valign:'top',color:PX.body},o));
  const R=(sl,x,y,w,h,color)=>sl.addShape(S.rect,{x,y,w,h,fill:{color},line:{type:'none'}});
  const L=(sl,x,y,w,h,color,pt=0.75)=>sl.addShape(S.line,{x,y,w,h,line:{color,width:pt}});
  const sgName=m.srcLabel;
  const sgLegend=m.srcIsSG&&m.rate!=null?`Start & Go (${m.rate.toLocaleString('it-IT',{minimumFractionDigits:2,maximumFractionDigits:2})}%)`:'Start & Go';
  const firstSw=m.delay+1,lastSw=m.endMonth;

  // Section plan (numbers/pages depend on what exists)
  const secs=[{id:'alloc',name:'Asset allocation'},{id:'fondi',name:'Strumenti finanziari'},{id:'costi',name:'Costi'}];
  if(m.hasSwitch)secs.push({id:'regime',name:`Allocazione al mese ${lastSw}`});
  secs.push({id:'darta',name:'Dettaglio polizza DARTA'});
  if(m.hasSwitch)secs.push({id:'switch',name:`Dettaglio ${sgName}`});
  let page=1;
  function content(sec,title,sub,fonteExtra){
    const sl=pptx.addSlide();sl.background={color:PX.white};page++;
    const n=secs.findIndex(s=>s.id===sec.id)+1;
    T(sl,sec.name,{x:M,y:0.22,w:6,h:0.25,fontSize:9,color:PX.blue});
    T(sl,n+'.',{x:10.09,y:0.22,w:1,h:0.25,fontSize:9,color:PX.blue,align:'right'});
    L(sl,M,0.5,CW,0,PX.blue);L(sl,M,7.72,CW,0,PX.blue);
    T(sl,fonteBase+(fonteExtra?' '+fonteExtra:''),{x:M,y:7.8,w:9.8,h:0.25,fontSize:7,color:PX.muted});
    T(sl,String(page),{x:10.09,y:7.8,w:1,h:0.25,fontSize:7,color:PX.muted,align:'right'});
    T(sl,title,{x:M,y:0.75,w:CW,h:0.65,fontSize:32,color:PX.blue});
    if(sub)T(sl,sub,{x:M,y:1.42,w:CW,h:0.35,fontSize:14,color:PX.azure});
    return sl;
  }
  const secById=id=>secs.find(s=>s.id===id);
  const pageOf=id=>secs.findIndex(s=>s.id===id)+2; // cover is page 1; one page per section unless tables overflow

  // 1 · Cover
  {const sl=pptx.addSlide();sl.background={color:PX.navy};
    T(sl,_cap(new Date().toLocaleDateString('it-IT',{month:'long',year:'numeric'})),{x:M,y:0.7,w:5,h:0.3,fontSize:12,color:PX.white});
    L(sl,M,3.05,1.2,0,PX.white);
    T(sl,'Proposta di portafoglio',{x:M,y:3.25,w:10,h:0.9,fontSize:42,color:PX.white});
    T(sl,`Cliente ${m.name}  ·  Patrimonio ${_pE(m.total)}`,{x:M,y:4.2,w:9,h:0.45,fontSize:18,color:PX.cyan});
    T(sl,secs.filter(s=>!['darta','switch'].includes(s.id)).map(s=>s.name).join(' · '),{x:M,y:4.75,w:9,h:0.35,fontSize:12,color:PX.white});
    sl.addImage({path:'/static/img/allianz_bank_fa_horizontal_white.png',x:M,y:7.3,w:3.6,h:3.6*86/603});
    T(sl,'Documento riservato',{x:8.09,y:7.35,w:3,h:0.3,fontSize:9,color:PX.white,align:'right'});
  }

  // 2 · Asset allocation
  {const sec=secById('alloc');
    const sub=m.hasSwitch?`Polizza DARTA da ${_pE(m.total)}: allocazione iniziale e allocazione a regime dopo lo ${sgName}`:`Polizza DARTA da ${_pE(m.total)}: allocazione dell'investimento`;
    const extra=m.hasSwitch?(m.delay>0?`Lo ${sgName} si attiva ${m.delay} mesi dopo l'investimento iniziale e si completa in ${m.mesi} mesi (mese ${lastSw}).`:`Lo ${sgName} si completa in ${m.mesi} switch mensili (mese ${lastSw}).`):'';
    const sl=content(sec,'Asset allocation',sub,extra);
    const donut=(x,seg,center,centerSub,head,headSub,o={})=>{
      const items=Object.entries(seg).filter(([,v])=>v>0.5);const tot=items.reduce((s,[,v])=>s+v,0)||1;
      const sz=o.size||3.2,cy=o.y||2.8;
      T(sl,head,{x,y:2.0,w:4.4,h:0.4,fontSize:18,color:PX.azure});
      T(sl,headSub,{x,y:2.4,w:4.4,h:0.25,fontSize:10,color:PX.muted});
      sl.addChart(pptx.ChartType.doughnut,[{name:'Allocazione',labels:items.map(([a])=>a==='START & GO'?sgLegend:pptAcLabel(a)),values:items.map(([,v])=>v/tot)}],
        {x:x+0.3,y:cy,w:sz,h:sz,holeSize:62,showLegend:false,showTitle:false,showValue:false,showPercent:true,showLabel:false,
         dataLabelColor:PX.white,dataLabelFontSize:sz>3.5?13:11,dataLabelFontBold:true,dataLabelFontFace:F,
         chartColors:items.map(([a])=>a==='START & GO'?PX.cyan:m.colorOf(a)),dataBorder:{pt:1.5,color:PX.white}});
      const cw=sz*0.6;
      T(sl,[{text:_pE(center),options:{fontSize:sz>3.5?24:20,color:PX.blue,breakLine:true}},{text:centerSub,options:{fontSize:9,color:PX.muted}}],
        {x:x+0.3+(sz-cw)/2,y:cy+sz/2-0.42,w:cw,h:0.84,align:'center',valign:'middle'});
      if(o.legend===false)return;
      const step=items.length>3?0.3:0.38;
      items.forEach(([a,v],i)=>{const y=6.25+i*step;
        sl.addShape(S.ellipse,{x,y:y+0.07,w:0.17,h:0.17,fill:{color:a==='START & GO'?PX.cyan:m.colorOf(a)},line:{type:'none'}});
        T(sl,a==='START & GO'?sgLegend:pptAcLabel(a),{x:x+0.3,y,w:2.0,h:0.3,fontSize:12});
        T(sl,[{text:_pE(v)+'  ',options:{bold:true,color:PX.blue}},{text:_pW(v/tot),options:{color:PX.muted}}],{x:x+1.9,y,w:1.9,h:0.3,fontSize:12,align:'right'});
      });
    };
    const seg0=m.allocAt(0);
    if(m.hasSwitch){
      donut(M,seg0,m.denom,'mese 0','Oggi','Investimento iniziale');
      donut(7.29,m.regime,m.denom,'mese '+lastSw,'Portafoglio da '+_pMMYY(lastSw),'Allocazione a regime');
      T(sl,sgName,{x:4.85,y:3.35,w:2.29,h:0.3,fontSize:12,bold:true,color:PX.blue,align:'center'});
      sl.addShape(S.rightArrow,{x:5.1,y:4.15,w:1.79,h:0.5,fill:{color:PX.panel},line:{type:'none'}});
      const dest=m.pac.length===1?`verso la componente ${m.pac[0].label.toLowerCase()}`:'verso i fondi della componente '+m.pac.map(c=>c.label.toLowerCase()).join(' e ');
      T(sl,`${m.mesi} switch mensili da ${_pE(m.rata)}\n${dest}`,{x:4.65,y:4.85,w:2.69,h:0.5,fontSize:9.5,align:'center'});
      T(sl,`Avvio: ${_pMonth(firstSw)} · completamento: ${_pMonth(lastSw)}`,{x:4.25,y:5.4,w:3.49,h:0.4,fontSize:9,bold:true,color:PX.azure,align:'center'});
    }else{
      // Nessuno switch: donut grande a sinistra, a destra composizione per componente + numeri chiave.
      donut(M,seg0,m.denom,'mese 0','Oggi','Investimento iniziale',{size:4.3,y:2.85,legend:false});
      const px=5.85,pw=W-M-px;
      T(sl,'Composizione per componente',{x:px,y:2.0,w:pw,h:0.4,fontSize:18,color:PX.azure});
      T(sl,'Controvalore, peso e costo medio annuo di ciascuna componente',{x:px,y:2.4,w:pw,h:0.25,fontSize:10,color:PX.muted});
      const rows=m.classes.map(c=>({c:c.color,l:c.label,n:c.funds.length,v:c.ctrv,comm:c.comm}));
      if(m.sgVal>0)rows.push({c:PX.cyan,l:sgLegend,n:0,v:m.sgVal,comm:0});
      const rh=rows.length<=3?0.52:rows.length<=5?0.44:0.36,cx=[px,px+2.35,px+3.1,px+4.25];
      let y=2.85;
      [['Componente','left'],['Fondi','right'],['Controvalore','right'],['Peso · Costo','right']].forEach(([t,al],i)=>
        T(sl,t,{x:i?cx[i]-(i===1?0.55:i===2?0.15:0.25):px,y,w:i===0?2:i===1?0.9:i===2?1.2:pw-4.0,h:0.22,fontSize:8,color:PX.muted,align:al}));
      y+=0.28;L(sl,px,y,pw,0,PX.hair);
      rows.forEach(r=>{
        sl.addShape(S.ellipse,{x:px,y:y+rh/2-0.08,w:0.17,h:0.17,fill:{color:r.c},line:{type:'none'}});
        T(sl,r.l,{x:px+0.3,y,w:1.9,h:rh,fontSize:12.5,bold:true,color:PX.blue,valign:'middle'});
        T(sl,r.n?String(r.n):'–',{x:cx[1]-0.55,y,w:0.9,h:rh,fontSize:11,align:'right',valign:'middle'});
        T(sl,_pE(r.v),{x:cx[2]-0.15,y,w:1.2,h:rh,fontSize:11.5,bold:true,color:PX.blue,align:'right',valign:'middle'});
        T(sl,[{text:_pW(r.v/m.denom),options:{bold:true,color:PX.blue}},{text:'  '+_pC(r.comm),options:{color:PX.muted,fontSize:9.5}}],{x:cx[3]-0.25,y,w:pw-4.0,h:rh,fontSize:11.5,align:'right',valign:'middle'});
        y+=rh;L(sl,px,y,pw,0,PX.hair,0.5);
      });
      L(sl,px,y,pw,0,PX.blue,1);
      // Numeri chiave: 2×2 se c'è spazio, altrimenti su una riga.
      const nF=m.classes.reduce((s,c)=>s+c.funds.length,0);
      const kpi=[[_pE(m.denom),'Patrimonio investito'],[_pC(m.initialComm),'Costo medio ponderato'],[_pE(m.dartaCost),'Costo annuo stimato'],[String(nF),nF===1?'Fondo in portafoglio':'Fondi in portafoglio']];
      const ky=y+0.4;
      if(ky+2.1<=7.55){const kw=(pw-0.2)/2;kpi.forEach(([v,l],i)=>{const kx=px+(i%2)*(kw+0.2),kyy=ky+Math.floor(i/2)*1.1;
        R(sl,kx,kyy,kw,0.95,PX.panel);T(sl,v,{x:kx+0.18,y:kyy+0.12,w:kw-0.3,h:0.5,fontSize:22,color:PX.blue,valign:'middle'});
        T(sl,l,{x:kx+0.18,y:kyy+0.62,w:kw-0.3,h:0.25,fontSize:9,bold:true,color:PX.blue});});}
      else if(ky+0.95<=7.55){const kw=(pw-0.3)/4;kpi.forEach(([v,l],i)=>{const kx=px+i*(kw+0.1);
        R(sl,kx,ky,kw,0.85,PX.panel);T(sl,v,{x:kx+0.1,y:ky+0.08,w:kw-0.15,h:0.45,fontSize:14,color:PX.blue,valign:'middle'});
        T(sl,l,{x:kx+0.1,y:ky+0.53,w:kw-0.15,h:0.28,fontSize:7.5,bold:true,color:PX.blue});});}
    }
  }

  // Column grid shared by "Strumenti finanziari" and "Costi"
  const grid=(n,min=3)=>{const slots=Math.max(min,n);const gap=slots>3?0.3:0.4;const w=(CW-gap*(slots-1))/slots;return{slots,w,x:i=>M+i*(w+gap)};};

  // 3 · Strumenti finanziari
  {const sec=secById('fondi');
    const sl=content(sec,'Strumenti finanziari','Tutti i fondi della polizza DARTA, con il peso all\'interno di ciascuna componente','Pesi sul controvalore di ciascuna componente; barre sulla stessa scala.');
    const cols=m.classes.map(c=>({title:'DARTA '+c.label,color:c.color,sub:`${_pE(c.ctrv)} · ${_pW(c.ctrv/m.denom)} del portafoglio · ${c.funds.length} fond${c.funds.length===1?'o':'i'}`,funds:c.byW}));
    if(m.hasSwitch){const pf=m.pac.flatMap(c=>c.funds).sort((a,b)=>b.ctrv-a.ctrv);const rs=pf.reduce((s,f)=>s+f.ctrv,0)||1;
      cols.push({title:sgName,color:PX.azure,bar:PX.cyan,sub:`${_pE(m.srcAmount)} · ${_pW(m.srcAmount/m.denom)} · switch mensile da ${_pE(m.rata)}`,funds:pf.map(f=>({...f,wIn:f.ctrv/rs}))});}
    const hasInfo=m.hasSwitch||m.sgVal>0;
    // Pochi componenti: lo slot libero ospita il pannello "In sintesi".
    const hasSyn=!hasInfo&&cols.length<3;
    const g=grid(cols.length+((hasInfo||hasSyn)&&cols.length<3?1:0),2);
    const maxW=Math.max(...cols.flatMap(c=>c.funds.map(f=>f.wIn)),0.0001);
    const maxRows=Math.max(...cols.map(c=>c.funds.length),1);
    // Pochi fondi: righe più alte e testo più grande, così la slide non resta vuota.
    const roomy=maxRows*0.66<=4.4&&g.w>=2.6;
    let rowH=roomy?0.66:Math.min(0.5,4.85/maxRows);const small=!roomy&&(g.w<2.6||rowH<0.42);
    // Give rows extra height (for two-line names) when some name won't fit on one line.
    if(!roomy&&4.85/maxRows>=0.55&&cols.some(c=>c.funds.some(f=>pptFit(f.name,g.w-0.8,small?8.5:9.5).fs<(small?8.5:9.5))))rowH=Math.min(0.6,4.85/maxRows);
    const fsName=roomy?11:small?8.5:9.5,fsTag=roomy?8.5:small?7:8,fsW=roomy?13:small?10:11,barH=roomy?0.06:0.04;
    cols.forEach((c,i)=>{const x=g.x(i);
      const tt=pptFit(c.title,g.w,small?15:18);
      T(sl,tt.text,{x,y:1.95,w:g.w,h:0.4,fontSize:tt.fs,color:c.color});
      T(sl,c.sub,{x,y:2.33,w:g.w,h:0.25,fontSize:small?8:9,color:PX.muted});
      L(sl,x,2.65,g.w,0,PX.hair,0.75);
      c.funds.forEach((f,j)=>{const y=2.75+j*rowH;
        // Long names wrap onto two lines when the row has room, otherwise shrink/ellipsise.
        const fs0=fsName,two=rowH>=0.55&&pptFit(f.name,g.w-0.8,fs0).fs<fs0;
        const nm=two?{text:f.name,fs:Math.max(7.5,fs0-1)}:pptFit(f.name,g.w-0.8,fs0);
        const nh=roomy?(two?0.36:0.25):(two?0.3:0.21);
        T(sl,nm.text,{x,y,w:g.w-0.75,h:nh,fontSize:nm.fs,bold:true,color:PX.blue});
        T(sl,f.tag,{x,y:y+nh+0.01,w:g.w-0.75,h:0.19,fontSize:fsTag,color:PX.muted});
        T(sl,_pW(f.wIn),{x:x+g.w-0.85,y,w:0.85,h:roomy?0.44:0.37,fontSize:fsW,color:PX.blue,align:'right',valign:'middle'});
        R(sl,x,y+rowH-0.13,g.w,barH,PX.panel);
        R(sl,x,y+rowH-0.13,Math.max(0.02,g.w*f.wIn/maxW),barH,c.bar||c.color);
      });
    });
    if(hasInfo){
      let txt;
      if(m.hasSwitch){
        const a=m.srcIsSG?`I ${_pE(m.srcAmount)} `:`I ${_pE(m.srcAmount)} del fondo ${m.srcFund.name} `;
        const wait=m.delay>0?(m.rate!=null?`vengono remunerati per ${m.delay} mesi ad un tasso del ${m.rate.toLocaleString('it-IT',{minimumFractionDigits:2})}% p.a. `:`restano in attesa per ${m.delay} mesi. `):'';
        txt=a+wait+(m.delay>0?`Dal mese ${firstSw} al mese ${lastSw}`:`Dal mese 1 al mese ${lastSw}`)+` vengono investiti ${_pE(m.rata)} al mese nei fondi indicati (Easy Switch).`;
      }else txt=`I ${_pE(m.sgVal)} restano nello Start & Go in attesa di essere investiti.`;
      // Free slot to the right, otherwise under the shortest column.
      let x,y;if(cols.length<g.slots){x=g.x(cols.length);y=1.95;}
      else{const sh=cols.slice(0,-1).reduce((b,c,i)=>c.funds.length<cols[b].funds.length?i:b,0);x=g.x(sh);y=2.75+cols[sh].funds.length*rowH+0.35;}
      if(y+1.35<7.6){R(sl,x,y,g.w,1.35,PX.panel);
        T(sl,[{text:`Come funziona lo ${sgName}`,options:{bold:true,color:PX.blue,breakLine:true}},{text:' ',options:{fontSize:4,breakLine:true}},{text:txt,options:{color:PX.body}}],{x:x+0.2,y:y+0.1,w:g.w-0.4,h:1.15,fontSize:9});}
    }
    const allF=m.classes.flatMap(c=>c.funds.map(f=>({...f,color:c.color}))).sort((a,b)=>b.ctrv-a.ctrv);
    // Pannello "In sintesi": nello slot libero, oppure sotto la colonna più corta se c'è spazio.
    let synAt=null;const synH=0.55+(allF.length>3?3:2)*0.62+0.95;
    if(hasSyn)synAt={x:g.x(cols.length),y:1.95};
    else if(!hasInfo&&cols.length>=3){const sh=cols.reduce((b,c,i)=>c.funds.length<cols[b].funds.length?i:b,0);
      const yy=2.75+cols[sh].funds.length*rowH+0.35;if(yy+synH<=7.6)synAt={x:g.x(sh),y:yy};}
    if(synAt&&allF.length){
      const {x,y}=synAt,top3=allF.slice(0,3).reduce((s,f)=>s+f.ctrv,0)/m.denom;
      R(sl,x,y,g.w,synH,PX.panel);
      T(sl,'In sintesi',{x:x+0.2,y:y+0.15,w:g.w-0.4,h:0.3,fontSize:12,bold:true,color:PX.blue});
      const st=[[String(allF.length),allF.length===1?'fondo in portafoglio':'fondi in portafoglio'],[String(m.classes.length),m.classes.length===1?'componente':'componenti']];
      if(allF.length>3)st.push([_pW(top3),'peso dei primi 3 fondi']);
      st.forEach(([v,l],i)=>{const yy=y+0.55+i*0.62;T(sl,v,{x:x+0.2,y:yy,w:1.25,h:0.5,fontSize:pptFit(v,1.25,26).fs,color:PX.blue,valign:'middle'});T(sl,l,{x:x+1.45,y:yy,w:g.w-1.6,h:0.5,fontSize:10,color:PX.body,valign:'middle'});});
      const big=allF[0];T(sl,[{text:'Fondo principale',options:{fontSize:8,color:PX.muted,breakLine:true}},{text:big.name,options:{fontSize:10,bold:true,color:PX.blue,breakLine:true}},{text:`${_pW(big.ctrv/m.denom)} del portafoglio`,options:{fontSize:9,color:PX.body}}],
        {x:x+0.2,y:y+0.6+st.length*0.62,w:g.w-0.4,h:0.75});
    }
    // Poche righe: fascia in basso con il peso di ogni fondo sul portafoglio totale.
    const endY=2.75+maxRows*rowH;
    if(allF.length&&endY<=5.15){
      const by=Math.max(endY+0.45,5.55),bh=0.42;
      T(sl,'Peso di ciascun fondo sul portafoglio totale',{x:M,y:by-0.4,w:CW,h:0.28,fontSize:12,color:PX.azure});
      const segs=allF.map(f=>({w:f.ctrv/m.denom,c:f.color,n:f.name}));if(m.sgVal>0)segs.push({w:m.sgVal/m.denom,c:PX.cyan,n:'Start & Go'});
      const tw=segs.reduce((s,x)=>s+x.w,0)||1;let bx=M;
      segs.forEach(sg=>{const w=CW*sg.w/tw;R(sl,bx,by,Math.max(0.01,w-0.03),bh,sg.c);
        if(w>=0.55)T(sl,_pW(sg.w),{x:bx,y:by,w:w-0.03,h:bh,fontSize:10,bold:true,color:PX.white,align:'center',valign:'middle'});
        if(w>=1.0){const nm=pptFit(sg.n,w-0.08,8.5);T(sl,nm.text,{x:bx,y:by+bh+0.06,w:w-0.03,h:0.36,fontSize:nm.fs,color:PX.body,align:'center'});}
        bx+=w;});
      if(segs.some(sg=>CW*sg.w/tw<1.0))T(sl,'Fondi minori senza etichetta: dettaglio nelle colonne sopra.',{x:M,y:by+bh+0.45,w:CW,h:0.22,fontSize:8,color:PX.muted});
    }
  }

  // 4 · Costi
  {const sec=secById('costi');
    const sl=content(sec,'Costi del portafoglio','Commissione annua di ciascun fondo e costo medio ponderato all\'investimento iniziale',`Costo DARTA = commissione di gestione + spese correnti del fondo + costo di copertura (${_pC(m.copRate)}).`);
    const hasSG=m.sgVal>0||m.hasSwitch;
    const nCols=m.classes.length+(hasSG?1:0);
    // Meno di 3 colonne: si aggiunge una colonna libera per il riepilogo costi.
    const freeCol=nCols<3;
    const g=grid(freeCol?nCols+1:nCols,2);
    const maxRows=Math.max(...m.classes.map(c=>c.funds.length),1);
    const roomy=maxRows*0.56<=3.55&&g.w>=2.6;
    const rowH=roomy?0.56:Math.min(0.42,(6.95-2.75-0.65)/maxRows);const small=!roomy&&(g.w<2.6||rowH<0.36);
    m.classes.forEach((c,i)=>{const x=g.x(i);
      const tt=pptFit('DARTA '+c.label,g.w,small?15:18);
      T(sl,tt.text,{x,y:1.95,w:g.w,h:0.4,fontSize:tt.fs,color:c.color});
      T(sl,`${_pW(c.ctrv/m.denom)} del portafoglio · ${c.funds.length} fond${c.funds.length===1?'o':'i'}`,{x,y:2.33,w:g.w-0.9,h:0.25,fontSize:small?8:9,color:PX.muted});
      T(sl,'Comm. tot.',{x:x+g.w-1,y:2.36,w:1,h:0.22,fontSize:8,color:PX.muted,align:'right'});
      L(sl,x,2.65,g.w,0,PX.hair);
      c.byW.forEach((f,j)=>{const y=2.75+j*rowH;
        const nm=pptFit(f.name,g.w-0.8,roomy?11:small?8.5:9.5);T(sl,nm.text,{x,y,w:g.w-0.75,h:roomy?0.25:0.2,fontSize:nm.fs,bold:true,color:PX.blue});
        T(sl,f.tag,{x,y:y+(roomy?0.26:0.19),w:g.w-0.75,h:0.18,fontSize:roomy?8.5:small?7:7.5,color:PX.muted});
        T(sl,_pC(f.comm),{x:x+g.w-0.85,y,w:0.85,h:roomy?0.46:0.35,fontSize:roomy?13:small?10:11,color:PX.blue,align:'right',valign:'middle'});
        L(sl,x,y+rowH-0.03,g.w,0,PX.hair,0.5);
      });
      const yt=2.75+c.funds.length*rowH+0.05;L(sl,x,yt,g.w,0,PX.blue,1);
      T(sl,[{text:'Costo medio annuo',options:{fontSize:10,bold:true,color:PX.blue,breakLine:true}},{text:'Comm. ponderata',options:{fontSize:8,color:PX.muted}}],{x,y:yt+0.08,w:g.w-1.1,h:0.42});
      T(sl,_pC(c.comm),{x:x+g.w-1.3,y:yt+0.05,w:1.3,h:0.46,fontSize:22,color:PX.blue,align:'right',valign:'middle'});
    });
    const bx=g.x(m.classes.length);
    if(hasSG){
      T(sl,sgName,{x:bx,y:1.95,w:g.w,h:0.4,fontSize:small?15:18,color:PX.azure});
      T(sl,`${_pW(m.srcAmount/m.denom)} del portafoglio · in attesa di switch`,{x:bx,y:2.33,w:g.w,h:0.25,fontSize:small?8:9,color:PX.muted});
      L(sl,bx,2.65,g.w,0,PX.hair);
      const head=m.rate!=null&&m.srcIsSG?'Liquidità remunerata':'Liquidità in attesa di investimento';
      const body=m.hasSwitch?`Nessun costo di gestione finché i ${_pE(m.srcAmount)} restano in ${sgName}. Dal mese ${firstSw}, ogni switch paga la commissione dei fondi in cui viene investito (media ${_pC(m.swComm)}).`
        :`Nessun costo di gestione finché i ${_pE(m.sgVal)} restano in Start & Go.`;
      if(m.srcIsSG||!m.hasSwitch){
        T(sl,[{text:head,options:{bold:true,color:PX.blue,breakLine:true}},{text:body,options:{color:PX.body}}],{x:bx,y:2.8,w:g.w,h:1.3,fontSize:9.5});
        L(sl,bx,4.2,g.w,0,PX.blue,1);
        T(sl,[{text:'Costo attuale',options:{fontSize:10,bold:true,color:PX.blue,breakLine:true}},{text:m.hasSwitch?(m.delay>0?'Fino a '+_pMonth(m.delay):'Fino al primo switch'):'Liquidità',options:{fontSize:8,color:PX.muted}}],{x:bx,y:4.28,w:g.w-1.2,h:0.42});
        T(sl,'0,00%',{x:bx+g.w-1.3,y:4.25,w:1.3,h:0.46,fontSize:22,color:PX.blue,align:'right',valign:'middle'});
      }else T(sl,body,{x:bx,y:2.8,w:g.w,h:1.3,fontSize:9.5});
    }
    let sx=hasSG?bx:g.x(m.classes.length),sy=hasSG?5.0:4.95;
    // Colonna libera: il riepilogo sale in alto e mostra il dettaglio per componente.
    if(freeCol){sx=g.x(nCols);sy=1.95;}
    if(!hasSG&&nCols>=3){const sh=m.classes.reduce((b,c,i)=>c.funds.length<m.classes[b].funds.length?i:b,0);
      sx=g.x(sh);sy=Math.max(4.95,2.75+m.classes[sh].funds.length*rowH+0.75);if(sy+2.45>7.6)sy=7.6-2.45;}
    if(freeCol){
      const ph=Math.min(5.55,2.45+0.55+m.classes.length*0.5+0.9);
      R(sl,sx,sy,g.w,ph,PX.panel);
      let yy=sy+2.45;L(sl,sx+0.2,yy,g.w-0.4,0,PX.hair);yy+=0.15;
      T(sl,'Costo per componente',{x:sx+0.2,y:yy,w:g.w-0.4,h:0.25,fontSize:9,bold:true,color:PX.blue});yy+=0.32;
      m.classes.forEach(c=>{sl.addShape(S.ellipse,{x:sx+0.2,y:yy+0.1,w:0.14,h:0.14,fill:{color:c.color},line:{type:'none'}});
        T(sl,[{text:c.label,options:{bold:true,color:PX.blue,breakLine:true}},{text:`${_pE(c.cost)} / anno`,options:{fontSize:8,color:PX.muted}}],{x:sx+0.42,y:yy,w:g.w-1.6,h:0.44,fontSize:10});
        T(sl,_pC(c.comm),{x:sx+g.w-1.3,y:yy,w:1.1,h:0.3,fontSize:13,color:PX.blue,align:'right'});yy+=0.5;});
      L(sl,sx+0.2,yy+0.05,g.w-0.4,0,PX.hair);
      T(sl,[{text:'Costo annuo stimato',options:{fontSize:9,bold:true,color:PX.blue,breakLine:true}},{text:'All\'investimento iniziale',options:{fontSize:8,color:PX.muted}}],{x:sx+0.2,y:yy+0.15,w:g.w-1.7,h:0.45});
      T(sl,_pE(m.dartaCost),{x:sx+g.w-1.7,y:yy+0.12,w:1.5,h:0.45,fontSize:18,color:PX.blue,align:'right',valign:'middle'});
    }else R(sl,sx,sy,g.w,2.45,PX.panel);
    T(sl,'Costo medio ponderato del portafoglio',{x:sx+0.2,y:sy+0.18,w:g.w-0.4,h:0.28,fontSize:10.5,color:PX.azure});
    T(sl,_pC(m.initialComm),{x:sx+0.2,y:sy+0.5,w:g.w-0.4,h:0.75,fontSize:40,color:PX.blue,valign:'middle'});
    T(sl,'All\'investimento iniziale',{x:sx+0.2,y:sy+1.3,w:g.w-0.4,h:0.22,fontSize:9,bold:true,color:PX.blue});
    const pesi=m.classes.map(c=>`${c.label} ${_pW(c.ctrv/m.denom)}`).concat(m.sgVal>0?[`Start & Go ${_pW(m.sgVal/m.denom)}`]:[]).join(' · ');
    T(sl,'Pesi: '+pesi,{x:sx+0.2,y:sy+1.55,w:g.w-0.4,h:0.4,fontSize:8.5,color:PX.muted});
    if(m.hasSwitch)T(sl,`A regime (mese ${lastSw}): ${_pC(m.regimeComm)} · dettaglio a pag. ${pageOf('regime')}`,{x:sx+0.2,y:sy+1.95,w:g.w-0.4,h:0.4,fontSize:8.5});
  }

  // 5 · Allocazione dopo lo switch
  if(m.hasSwitch){const sec=secById('regime');
    const sl=content(sec,`Allocazione dopo lo ${sgName}`,`Evoluzione dell'allocazione dall'investimento iniziale al mese ${lastSw}`,
      (m.delay>0?`${sgName}: avvio ${m.delay} mesi dopo l'investimento iniziale, `:`${sgName}: `)+`${m.mesi} switch mensili da ${_pE(m.rata)} (mesi ${firstSw}-${lastSw}).`);
    const keys=[...m.order.filter(a=>a!=='START & GO')];if(m.srcIsSG)keys.push('START & GO');
    const col=a=>a==='START & GO'?PX.cyan:m.colorOf(a);
    keys.forEach((a,i)=>{const x=M+i*1.35;R(sl,x,1.98,0.14,0.14,col(a));T(sl,a==='START & GO'?'Start & Go':pptAcLabel(a),{x:x+0.22,y:1.93,w:1.1,h:0.24,fontSize:9.5});});
    // Bars: mese 0, avvio (if waiting), each switch — sampled down to 8 bars max.
    let sw=[];for(let k=1;k<=m.mesi;k++)sw.push(k);
    const room=8-1-(m.delay>0?1:0);
    if(sw.length>room){const pick=new Set([1,m.mesi]);for(let i=1;i<room-1;i++)pick.add(Math.round(1+i*(m.mesi-1)/(room-1)));sw=[...pick].sort((a,b)=>a-b);}
    const bars=[{mo:0,k:0,lab:'Investimento'}];
    if(m.delay>0)bars.push({mo:m.delay,k:0,lab:'Avvio '+(m.srcIsSG?'S&G':'switch')});
    sw.forEach(k=>bars.push({mo:m.delay+k,k,lab:'Switch '+k}));
    const top=2.35,hTot=4.3,bw=0.68,span=6.6/bars.length;
    bars.forEach((b,i)=>{const seg=m.allocAt(b.k);const tot=Object.values(seg).reduce((s,v)=>s+Math.max(0,v),0)||1;
      const x=M+span*(i+0.5)-bw/2;let y=top+hTot;
      keys.forEach(a=>{const v=Math.max(0,seg[a]||0);if(v<=0.5)return;const h=hTot*v/tot;y-=h;
        R(sl,x,y,bw,h,col(a));
        if(h>=0.2)T(sl,_pW(v/tot),{x,y,w:bw,h,fontSize:9,bold:true,color:PX.white,align:'center',valign:'middle'});});
      T(sl,'Mese '+b.mo,{x:x+bw/2-span/2,y:6.73,w:span,h:0.22,fontSize:9.5,bold:true,color:PX.blue,align:'center'});
      T(sl,b.lab,{x:x+bw/2-span/2,y:6.95,w:span,h:0.2,fontSize:8,color:PX.muted,align:'center'});
    });
    L(sl,7.5,2.0,0,5.4,PX.hair);
    T(sl,`Allocazione al mese ${lastSw}`,{x:7.8,y:1.95,w:3.29,h:0.35,fontSize:16,color:PX.azure});
    const fin=keys.filter(a=>a!=='START & GO'&&(m.regime[a]||0)>0.5).sort((a,b)=>m.regime[b]-m.regime[a]);
    const big=fin.length<=2,stepY=big?1.05:Math.min(0.75,2.2/Math.max(fin.length,1));
    fin.forEach((a,i)=>{const y=2.45+i*stepY;const rc=m.regimeClass[a];
      T(sl,_pW(m.regime[a]/m.denom),{x:7.8,y,w:1.6,h:big?0.8:stepY,fontSize:big?44:Math.min(30,stepY*40),color:col(a),valign:'middle'});
      T(sl,[{text:pptAcLabel(a),options:{fontSize:12,bold:true,color:PX.blue,breakLine:true}},{text:`${_pE(m.regime[a])} · ${rc.funds} fond${rc.funds===1?'o':'i'}`,options:{fontSize:9,color:PX.muted}}],
        {x:9.5,y:y+(big?0.12:0),w:1.59,h:big?0.56:stepY,valign:'middle'});
    });
    R(sl,7.8,4.75,3.29,2.7,PX.panel);
    T(sl,`Costo medio ponderato al mese ${lastSw}`,{x:8.0,y:4.93,w:2.89,h:0.28,fontSize:10.5,color:PX.azure});
    T(sl,_pC(m.regimeComm),{x:8.0,y:5.25,w:2.89,h:0.75,fontSize:40,color:PX.blue,valign:'middle'});
    const lines=fin.map(a=>({text:`${pptAcLabel(a)} ${_pC(m.regimeClass[a].comm)} · peso ${_pW(m.regime[a]/m.denom)}`,options:{color:PX.blue,breakLine:true}}));
    lines.push({text:`Oggi: ${_pC(m.initialComm)}`+(m.srcIsSG?` (${m.rate!=null?'Start & Go remunerato':'Start & Go senza costi'})`:''),options:{color:PX.muted}});
    T(sl,lines,{x:8.0,y:6.1,w:2.89,h:1.2,fontSize:9});
  }

  // Table helper (paginates long tables, repeating header + furniture)
  const tableStyle={fontFace:F,fontSize:8,color:PX.body,valign:'middle',border:{type:'solid',pt:0.5,color:PX.hair},margin:0.05};
  const hdrCell=(t,al)=>({text:t,options:{bold:true,color:PX.white,fill:{color:PX.blue},align:al||'left'}});
  function pagedTable(sec,title,sub,fonteExtra,head,body,colW,after){
    const fitH=5.45/(body.length+1);const short=body.length+1<=12;
    const rowH=short?0.34:fitH>=0.22?Math.min(0.27,fitH):0.27;
    const per=fitH>=0.22?body.length:19;let i=0,first=true,sl,y=2.0;
    do{const chunk=body.slice(i,i+per);i+=per;
      sl=content(sec,first?title:title+' (segue)',first?sub:'',fonteExtra);first=false;
      sl.addTable([head,...chunk],Object.assign({x:M,y:2.0,w:CW,colW,rowH},tableStyle,short?{fontSize:9}:{}));
      y=2.0+rowH*(chunk.length+1);
    }while(i<body.length);
    if(after)after(sl,y);
  }
  const num=(t,o={})=>({text:t,options:Object.assign({align:'right'},o)});

  // 6 · Dettaglio polizza DARTA
  {const sec=secById('darta');
    const head=['Settore','ISIN','Fondo','Controvalore','% classe','CG','SC','Cop.','Comm. tot.','Costo annuo'].map((t,i)=>hdrCell(t,i>=3?'right':'left'));
    const body=[];const grey={fill:{color:PX.panel},bold:true,color:PX.blue};
    m.classes.forEach(c=>{
      body.push([{text:c.label,options:grey},{text:'',options:grey},{text:`${_pW(c.ctrv/m.denom)} del portafoglio`,options:grey},num(_pE(c.ctrv),grey),num('100%',grey),
        num('',grey),num('',grey),num('',grey),num(_pC(c.comm),grey),num(_pE2(c.cost),grey)]);
      c.funds.forEach(f=>body.push([pptSettore(f.sub),f.isin||'–',{text:[f.line,f.name].filter(Boolean).join(' '),options:{bold:true,color:PX.blue}},
        num(_pE(f.ctrv)),num(_pC(f.wIn)),num(f.cg!=null?_pC(f.cg):'–'),num(f.sc!=null?_pC(f.sc):'–'),num(_pC(f.cop)),num(_pC(f.comm)),num(_pE2(f.cost))]));
    });
    const inv=m.classes.reduce((s,c)=>s+c.ctrv,0);const tot={fill:{color:PX.totRow},bold:true,color:PX.blue};
    body.push([{text:'Totale investito',options:tot},{text:'',options:tot},{text:`${_pW(inv/m.denom)} del portafoglio`,options:tot},num(_pE(inv),tot),num('',tot),num('',tot),num('',tot),num('',tot),
      num(_pC(inv>0?m.dartaCost/inv:0),tot),num(_pE2(m.dartaCost),tot)]);
    const comps=m.classes.map(c=>c.label.toLowerCase());
    pagedTable(sec,'Dettaglio polizza DARTA',`Componenti ${comps.length>1?comps.slice(0,-1).join(', ')+' e '+comps[comps.length-1]:comps[0]||''} investite: controvalore, peso nella classe e struttura commissionale`,
      'CG = commissione di gestione, SC = spese correnti del fondo, Cop. = costo di copertura.',head,body,[1.55,1.1,2.45,0.95,0.75,0.55,0.55,0.55,0.85,1.19],
      (sl,y)=>{if(m.sgVal>0)T(sl,`I restanti ${_pE(m.sgVal)} (${_pW(m.sgVal/m.denom)} del portafoglio) sono destinati allo Start & Go`+(m.hasSwitch?`: dettaglio degli switch mensili a pagina ${page+1}.`:'.'),{x:M,y:y+0.1,w:CW,h:0.3,fontSize:9.5});
        // Tabella corta: numeri chiave nello spazio libero sotto (come nel dettaglio switch).
        const sy=Math.max(y+0.7,5.55);if(sy>6.6)return;L(sl,M,sy,CW,0,PX.hair);
        const nF=m.classes.reduce((s,c)=>s+c.funds.length,0);
        const st=[[_pE(inv),'Controvalore investito'],[String(nF),nF===1?'Fondo':'Fondi'],[_pC(inv>0?m.dartaCost/inv:0),'Commissione media ponderata'],[_pE(m.dartaCost),'Costo annuo stimato']];
        st.forEach(([v,l],i)=>{const x=M+i*2.62;T(sl,v,{x,y:sy+0.25,w:2.62,h:0.7,fontSize:32,color:PX.blue});T(sl,l,{x,y:sy+0.95,w:2.62,h:0.3,fontSize:10,bold:true,color:PX.blue});});});
  }

  // 7 · Dettaglio Start & Go / Easy Switch
  if(m.hasSwitch){const sec=secById('switch');
    const head=['Settore','ISIN','Fondo','CTRV / mese','CG','SC','Cop.','Comm. tot.','Costo annuo / rata',`Costo annuo a regime`].map((t,i)=>hdrCell(t,i>=3?'right':'left'));
    const body=[];
    m.pac.forEach(c=>c.funds.forEach(f=>body.push([pptSettore(f.sub),f.isin||'–',{text:[f.line,f.name].filter(Boolean).join(' '),options:{bold:true,color:PX.blue}},
      num(_pE2(f.ctrv).replace(/,00$/,'')),num(f.cg!=null?_pC(f.cg):'–'),num(f.sc!=null?_pC(f.sc):'–'),num(_pC(f.cop)),num(_pC(f.comm)),num(_pE2(f.cost)),num(_pE2(f.cost*m.mesi))])));
    const rs=m.pac.reduce((s,c)=>s+c.ctrv,0),cs=m.pac.reduce((s,c)=>s+c.cost,0);const tot={fill:{color:PX.totRow},bold:true,color:PX.blue};
    const mix=m.pac.map(c=>`${_pW(c.ctrv/rs)} ${c.label.toLowerCase()}`).join(' · ');
    body.push([{text:'Totale switch',options:tot},{text:'',options:tot},{text:mix,options:tot},num(_pE(rs),tot),num('',tot),num('',tot),num('',tot),num(_pC(m.swComm),tot),num(_pE2(cs),tot),num(_pE2(cs*m.mesi),tot)]);
    pagedTable(sec,`Dettaglio ${sgName}`,`Ripartizione di ciascuno switch mensile da ${_pE(m.rata)} (Easy Switch) nei fondi`,
      `CG = commissione di gestione, SC = spese correnti del fondo, Cop. = costo di copertura. Costo annuo a regime = costo annuo di una rata × ${m.mesi} switch.`,
      head,body,[1.45,1.05,2.4,0.95,0.55,0.55,0.55,0.8,1.0,1.19],
      (sl,y)=>{const sy=Math.max(y+0.5,5.55);if(sy>6.6)return;L(sl,M,sy,CW,0,PX.hair);
        const st=[[_pE(m.rata),'Importo di ogni switch'],[String(m.mesi),'Switch mensili'],[_pE(m.rata*m.mesi),'Totale investito tramite switch'],[_pC(m.swComm),'Commissione media ponderata']];
        st.forEach(([v,l],i)=>{const x=M+i*2.62;T(sl,v,{x,y:sy+0.25,w:2.62,h:0.7,fontSize:32,color:PX.blue});T(sl,l,{x,y:sy+0.95,w:2.62,h:0.3,fontSize:10,bold:true,color:PX.blue});});});
  }

  await pptx.writeFile({fileName:`Proposta_portafoglio_${CD.name.replace(/\s+/g,'_')}.pptx`});
  toast('PowerPoint scaricato','ok');
}
