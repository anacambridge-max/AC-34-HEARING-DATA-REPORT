'use client';
import {useMemo,useState,useRef,useEffect} from 'react';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';

const n=v=>Number.isFinite(Number(v))?Number(v):0;
const pct=v=>n(v).toFixed(2)+'%';
const cleanOfficerName=s=>String(s??'').replace(/\s*\(\s*\d{10}\s*\)/g,'').replace(/\s*-?\s*\d{10}\b/g,'').replace(/\s+/g,' ').trim();
const rgb=(r,g,b)=>[r,g,b];
const excelBlue=rgb(31,56,100), excelHeaderText=rgb(255,255,255), stripe=rgb(242,242,242), totalYellow=rgb(255,192,0);
const red=rgb(248,105,107), yellow=rgb(255,235,132), green=rgb(99,190,123);

function scaleColor(v){
 const x=Math.max(0,Math.min(100,n(v)));
 if(x<=50){const t=x/50;return rgb(Math.round(248+(255-248)*t),Math.round(105+(235-105)*t),Math.round(107+(132-107)*t))}
 const t=(x-50)/50;return rgb(Math.round(255+(99-255)*t),Math.round(235+(190-235)*t),Math.round(132+(123-132)*t));
}
function hearingDateStyle(status){
 const s=String(status??'').toLowerCase();
 if(s.includes('pending')) return {bg:[255,199,206],text:[156,0,6]};
 if(s.includes('partial')) return {bg:[255,235,156],text:[156,101,0]};
 if(s.includes('held')) return {bg:[198,239,206],text:[0,97,0]};
 if(s.includes('no hearing')) return {bg:[231,230,230],text:[89,89,89]};
 return null;
}
function rowsFor(wb,name){
 const ws=wb.Sheets[name];
 return ws?XLSX.utils.sheet_to_json(ws,{header:1,defval:null}):[];
}
function findHeader(rows,tests){
 for(let i=0;i<Math.min(rows.length,20);i++){
   const row=rows[i]||[];
   const ok=tests.every(t=>row.some(v=>String(v??'').trim().toLowerCase()===t.toLowerCase()));
   if(ok)return {index:i,row};
 }
 return {index:-1,row:[]};
}
function parseReference(wb){
 const m=rowsFor(wb,'PS Mapping'),h=rowsFor(wb,'Hearing Dates');
 if(!m.length||!h.length) throw Error('Reference workbook must contain PS Mapping and Hearing Dates sheets.');
 const mr=m.slice(1).filter(r=>r?.[0]!=null).map(r=>[n(r[0]),String(r[1]??''),String(r[2]??''),String(r[3]??''),String(r[4]??'')]);
 const hr=h.slice(1).filter(r=>r?.[0]!=null).map(r=>[n(r[0]),String(r[4]??''),String(r[5]??'')]);
 // IMPORTANT: "Hearing Notice Scheduled NO MAPPING" is NOT the BLO Documents sheet's
 // "Hearing Notice Scheduled" field. It comes from the reference Officer Wise PS Detail.
 const pd=rowsFor(wb,'Officer Wise PS Detail');
 const si=pd.findIndex(r=>r?.some(v=>String(v??'').trim().toLowerCase()==='part no') && r?.some(v=>String(v??'').trim().toLowerCase()==='hearing notice scheduled no mapping'));
 const scheduledMap=new Map();
 if(si>=0){
   const ph=pd[si].map(v=>String(v??'').trim().toLowerCase());
   const pi=ph.findIndex(v=>v==='part no');
   const sci=ph.findIndex(v=>v==='hearing notice scheduled no mapping');
   pd.slice(si+1).forEach(r=>{const ps=n(r[pi]);if(ps)scheduledMap.set(ps,n(r[sci]))});
 }
 if(scheduledMap.size<400) throw Error('Reference workbook: Hearing Notice Scheduled NO MAPPING data for all 430 PS is required.');
 let order= [];
 const ow=rowsFor(wb,'Officer Wise Report');
 const oi=Math.max(0,ow.findIndex(r=>r?.some(v=>String(v??'').trim().toLowerCase()==='officer name')));
 if(oi>=0) order=ow.slice(oi+1).filter(r=>r?.[1]&&!String(r[1]).toUpperCase().includes('GRAND TOTAL')).map(r=>cleanOfficerName(r[1]));
 const names=[...new Set(mr.map(r=>cleanOfficerName(r[1])))];
 order=[...order,...names.filter(x=>!order.includes(x))];
 return {mapping:mr,hearing:hr,scheduledMap:Object.fromEntries(scheduledMap),officerOrder:order};
}
const AERO_GROUPS=[
  {name:"Sh. Ajay Kumar, Ad.AERO",mobile:"9718343424",ps:[146,147,148,149,150,151,152,153,154,155,156,157,158,159,160,161,162,163,170,171,172,175,176,177,178,179,180,185,187,188,189,190,191,192,193,276,277,286,287,288]},
  {name:"Smt. Ranjana Sharma, Ad.AERO",mobile:"9540474555",ps:[164,165,166,167,168,169,173,174,181,182,183,184,186,222,224,225,231,232,233,234,278,279,280,281,282,283,284,285]},
  {name:"Smt. Mamta Meena, Ad.AERO",mobile:"9318473015",ps:[194,195,196,197,198,199,200,201,202,203,204,205,206,207,208,209,210,211,212,213,214,215,216,217,218,219,220,221,223,226,227,228,229,230]},
  {name:"Sh. Rajesh Shriwastav, Ad.AERO",mobile:"9818148076",ps:[1,2,3,12,13,14,18,19,20,21,22,23,24,25,35,42,43,44,45,46,47,48,49,50,55,56]},
  {name:"Smt. Saroj Meena, Ad.AERO",mobile:"9818244615",ps:[4,5,6,7,8,9,10,11,15,16,17,26,27,28,29,30,31,32,33,34,57,58,59,60]},
  {name:"Smt. Shashi Bala, AERO",mobile:"9953312984",ps:[80,81,82,83,84,85,86,87,88,89,90,97,98,99,100,101,102,103,104,105,106,107,108,109,135,136,137,138,139,140,141,142,143,144,145]},
  {name:"Sh. Hemvir, Ad.AERO",mobile:"9910172161",ps:[36,37,38,39,40,41,51,52,53,54,61,62,63,64,65,66,67,68,69,70,71,72,73,74,75,76,77,78,79,91,92,93,94,95,96]},
  {name:"Sh. Rakesh Yadav, Ad.AERO",mobile:"7011971522",ps:[110,111,112,113,114,115,116,117,118]},
  {name:"Sh. Mohit, Ad.AERO",mobile:"9654467035",ps:[119,120,121,122,123,124,125,126,127,128,129,130,131,132,133,134,331,332,333,334,335,336,337,338,339,340,341,343,344,347,348,349]},
  {name:"Sh. Subhashish Boss, AERO",mobile:"9810598314",ps:[235,236,237,245,246,247,252,253,254,255,256,264,265,266,268,269,270,289,290,291,292,310,311,312,313,314,315,316,317,318,319,320,321,322]},
  {name:"Smt. Vandana Bansal, Ad.AERO",mobile:"8700311269",ps:[238,239,240,241,242,243,244,248,249,250,251,257,258,259,260,261,262,263,267,271,272,273,274,275,299,300,301,302,325,326,327,328,329,330]},
  {name:"Sh. Dharamvir Singh, Ad.AERO",mobile:"9871534579",ps:[293,294,295,296,297,298,303,304,305,306,307,308,309,323,324]},
  {name:"Sh. Virender Singh, AERO",mobile:"9868252144",ps:[342,345,346,350,351,352,353,354,355,356,357,358,359,360,361,362,363,364,365,366,367,368,369,370,371,372,373,374]},
  {name:"Sh. Manoj Kumar, Ad.AERO",mobile:"9811802528",ps:[375,376,377,378,379,380,381,382,383,384,385,386,387,388,389,390,391,392,393,394,395,396,397,398,399,400,401,402,403,404,405,406,407,408,409,410,411,412,413,414,415,416,417,418,419,420,421,422,423,424,425,426,427,428,429,430]}
];
const AERO_BY_PS=new Map(AERO_GROUPS.flatMap(g=>g.ps.map(ps=>[ps,{name:g.name,mobile:g.mobile}])));

function parseECI(wb){
 const names=['sirNoticeGenerate','ECI Raw Data','Part Wise Report'];
 let rows=[];
 for(const name of names){rows=rowsFor(wb,name);if(rows.length)break}
 if(!rows.length)throw Error('ECI workbook: sirNoticeGenerate / ECI Raw Data sheet not found.');
 const hi=rows.findIndex(r=>r?.some(v=>String(v??'').trim().toLowerCase()==='part no'));
 if(hi<0)throw Error('ECI workbook: Part No column not found.');
 const head=rows[hi].map(v=>String(v??'').trim());
 const ix=key=>head.findIndex(v=>v.toLowerCase()===key.toLowerCase());
 const p=ix('Part No'),gen=ix('Notice Generated'),pg=ix('Pending for Notice Generation'),del=ix('Notice Delivered'),pend=ix('Notice Pending Delivery'),held=ix('Hearings Held'),lapsed=ix('Hearing Date Lapsed'),park=ix('Parked for Final Publication');
 if([p,gen,pg,del,pend,held,lapsed].some(x=>x<0))throw Error('ECI workbook is missing one or more required columns.');
 const out=new Map();
 rows.slice(hi+1).forEach(r=>{
   const ps=n(r[p]);if(!ps)return;
   out.set(ps,{ps,generated:n(r[gen]),pendingGen:n(r[pg]),delivered:n(r[del]),pending:n(r[pend]),held:n(r[held]),lapsed:n(r[lapsed]),parked:park>=0?n(r[park]):0});
 });
 if(out.size<400)throw Error('ECI workbook appears incomplete: fewer than 400 PS records found.');
 return out;
}
function parseBLO(wb){
 const names=['PS Wise Report','BLO Documents Raw Data','BLO Documents','Documents Uploaded by BLO'];
 let rows=[];
 for(const name of names){rows=rowsFor(wb,name);if(rows.length)break}
 if(!rows.length)throw Error('BLO workbook: PS Wise Report / BLO Documents Raw Data sheet not found.');
 const hi=rows.findIndex(r=>r?.some(v=>String(v??'').trim().toLowerCase()==='ps no.'));
 if(hi<0)throw Error('BLO workbook: PS No. column not found.');
 const head=rows[hi].map(v=>String(v??'').trim().toLowerCase());
 const find=(...keys)=>head.findIndex(v=>keys.some(k=>v===k.toLowerCase()));
 const p=find('ps no.','ps no','part no'),sch=find('hearing notice scheduled'),docs=find('documents uploaded by blo');
 if(p<0||sch<0||docs<0)throw Error('BLO workbook must contain PS No., Hearing Notice Scheduled and Documents Uploaded by BLO.');
 const out=new Map();
 rows.slice(hi+1).forEach(r=>{const ps=n(r[p]);if(ps)out.set(ps,{ps,scheduled:n(r[sch]),docs:n(r[docs])})});
 if(out.size<400)throw Error('BLO workbook appears incomplete: fewer than 400 PS records found.');
 return out;
}
function buildData(ref,eci,blo,baseline){
 const hm=new Map(ref.hearing.map(r=>[n(r[0]),{dates:r[1]||'',status:r[2]||''}]));
 const mm=new Map(ref.mapping.map(r=>[n(r[0]),{officer:r[1]||'',blo:r[2]||'',supervisor:r[3]||'',centre:r[4]||''}]));
 const names=ref.officerOrder.length?ref.officerOrder:[...new Set(ref.mapping.map(r=>cleanOfficerName(r[1])))];
 const details=ref.mapping.map(r=>{
   const ps=n(r[0]),map=mm.get(ps)||{},e=eci.get(ps)||{},b=blo.get(ps)||{},h=hm.get(ps)||{};
   const delivered=n(e.delivered),scheduled=n(ref.scheduledMap?.[ps]),docs=n(b.docs),held=n(e.held),lapsed=n(e.lapsed);
   const y=baseline?.get(ps)||{delivered:0,pending:0,held:0,lapsed:0};
   const todayDelivered=delivered-n(y.delivered),todayPending=n(e.pending)-n(y.pending),todayHeld=held-n(y.held),todayLapsed=lapsed-n(y.lapsed);
   const ai=AERO_BY_PS.get(ps)||{name:'Unmapped AERO',mobile:''};
   return {ps,aero:ai.name,aeroMobile:ai.mobile,officer:map.officer,blo:map.blo,supervisor:map.supervisor,centre:map.centre,generated:n(e.generated),pendingGen:n(e.pendingGen),scheduled,delivered,pending:n(e.pending),deliveredPct:scheduled?delivered/scheduled*100:0,docs,docsPct:delivered?docs/delivered*100:0,dates:h.dates||'',status:h.status||'',heldLapsed:held+lapsed,disposal:held,disposalPct:(held+lapsed)?held/(held+lapsed)*100:0,yesterdayDelivered:n(y.delivered),todayDelivered,todayPending,yesterdayPending:n(y.pending),yesterdayHeld:n(y.held),todayHeld,yesterdayLapsed:n(y.lapsed),todayLapsed,totalDelivered:delivered,totalPending:n(e.pending),totalHeld:held,totalLapsed:lapsed,parked:n(e.parked)};
 });
 const wanted=['S No','Officer Name','No. of PS','Notice Generated (NO MAP + ANOMALY)','Hearing Notice Scheduled NO MAPPING','NO MAP NOTICE DELIVERED','% NO MAPPING DELIVERED','Documents Uploaded by BLO','% Docs Uploaded (of Notice Delivered)','Hearing Held + Date Lapsed','Total Disposal','% Total Disposal'];
 const officers=names.map((name,i)=>{
   const rr=details.filter(r=>cleanOfficerName(r.officer)===cleanOfficerName(name));
   const sum=k=>rr.reduce((a,r)=>a+n(r[k]),0);
   const scheduled=sum('scheduled'),delivered=sum('delivered'),docs=sum('docs'),heldLapsed=sum('heldLapsed'),disposal=sum('disposal');
   const yesterdayDisposal=rr.reduce((a,r)=>a+n(r.yesterdayHeld),0),todayDisposal=rr.reduce((a,r)=>a+n(r.todayHeld),0);
   const row=[i+1,name,rr.length,sum('generated'),scheduled,delivered,scheduled?delivered/scheduled*100:0,docs,delivered?docs/delivered*100:0,heldLapsed,yesterdayDisposal,todayDisposal,disposal,heldLapsed?disposal/heldLapsed*100:0];
   return {sno:i+1,name:cleanOfficerName(name),ps:rr.length,generated:sum('generated'),pendingGen:sum('pendingGen'),scheduled,delivered,deliveredPct:scheduled?delivered/scheduled*100:0,pendingDelivery:sum('pending'),held:disposal,lapsed:heldLapsed-disposal,docs,reportRow:row};
 });
 const reportHeaders=['S No','Officer Name','No. of PS','Notice Generated (NO MAP + ANOMALY)','Hearing Notice Scheduled NO MAPPING','NO MAP NOTICE DELIVERED','% NO MAPPING DELIVERED','Documents Uploaded by BLO','% Docs Uploaded (of Notice Delivered)','Hearing Held + Date Lapsed','Disposal Till Yesterday','Disposal Today','Total Disposal','% Total Disposal'];
 const aeroReports=AERO_GROUPS.map((g,gi)=>{
   const rr=details.filter(r=>AERO_BY_PS.get(r.ps)?.name===g.name);
   const supervisors=[...new Set(rr.map(r=>String(r.supervisor||'').trim()).filter(Boolean))];
   const generated=rr.reduce((s,r)=>s+n(r.generated),0),delivered=rr.reduce((s,r)=>s+n(r.delivered),0),lapse=rr.reduce((s,r)=>s+n(r.totalLapsed),0),held=rr.reduce((s,r)=>s+n(r.totalHeld),0),parked=rr.reduce((s,r)=>s+n(r.parked),0);
   return {sno:gi+1,name:g.name,mobile:g.mobile,supervisors,ps:rr.length,generated,delivered,deliveredPct:generated?delivered/generated*100:0,lapse,held,heldLapsed:held+lapse,heldLapsedPct:generated?(held+lapse)/generated*100:0,parked,rows:rr.sort((a,b)=>a.ps-b.ps)};
 });
 const reportRows=officers.map(o=>o.reportRow);
 const grandVals=(()=>{const sum=k=>details.reduce((a,r)=>a+n(r[k]),0);const scheduled=sum('scheduled'),delivered=sum('delivered'),docs=sum('docs'),heldLapsed=sum('heldLapsed'),disposal=sum('disposal'),yesterdayDisposal=sum('yesterdayHeld'),todayDisposal=sum('todayHeld');return ['GRAND TOTAL','',details.length,sum('generated'),scheduled,delivered,scheduled?delivered/scheduled*100:0,docs,delivered?docs/delivered*100:0,heldLapsed,yesterdayDisposal,todayDisposal,disposal,heldLapsed?disposal/heldLapsed*100:0]})();
 return {officers,details,aeroReports,reportHeaders,reportRows,grandRow:grandVals,sourceCounts:{eci:eci.size,blo:blo.size,reference:ref.mapping.length},hasBaseline:!!baseline};
}
function parseLegacy(wb){return parseReference(wb)}


function officerVisibleIndexes(headers){return headers.map((h,i)=>({h,i})).filter(x=>x.h!=='Documents Uploaded by BLO'&&x.h!=='% Docs Uploaded (of Notice Delivered)').map(x=>x.i)}
function reportPDF(title,headers,rows,grandRow,psRows=[]){
 const d=new jsPDF({orientation:'landscape',unit:'mm',format:'a3'});
 const generatedAt=new Date().toLocaleString('en-IN',{timeZone:'Asia/Kolkata',day:'2-digit',month:'2-digit',year:'numeric',hour:'2-digit',minute:'2-digit',second:'2-digit',hour12:true});
 d.setFont('helvetica','bold');d.setTextColor(...excelBlue);d.setFontSize(17);
 d.text('AC-34 MATIALA — OFFICER WISE REPORT',210,13,{align:'center'});
 d.setFontSize(11);d.text(title,210,21,{align:'center'});
 d.setFontSize(8);d.setTextColor(80,80,80);d.text('SIR-2026 • Officer-wise monitoring report',210,26,{align:'center'});
 d.setFontSize(7.5);d.text('Report Generated: '+generatedAt,291,26,{align:'right'});
 const isPctHeader=h=>['% NO MAPPING DELIVERED','% Docs Uploaded (of Notice Delivered)','% Total Disposal'].includes(String(h||'').trim());
 const fmt=(v,i)=>{const num=parseFloat(String(v??''));return isPctHeader(headers[i])&&Number.isFinite(num)?num.toFixed(2):String(v??'')};
 let body=rows.map(r=>headers.map((_,i)=>fmt(r[i],i)));
 if(grandRow){const gr=grandRow.map((v,i)=>fmt(v,i));gr[0]='GRAND TOTAL';gr[1]='';body.push(gr)}
 const totalIndex=grandRow?body.length-1:-1;
 const widths=[10,43,14,26,27,25,24,29,34,28,23,23,23,21];
 autoTable(d,{
   startY:31,head:[headers],body,theme:'grid',tableWidth:390,
   columnStyles:Object.fromEntries(widths.map((w,i)=>[i,{cellWidth:w}])),
   styles:{font:'helvetica',fontSize:9.5,fontStyle:'bold',cellPadding:{top:3,right:1.8,bottom:3,left:1.8},overflow:'linebreak',valign:'middle',halign:'center',lineColor:[0,0,0],lineWidth:.35,textColor:[20,20,20],minCellHeight:12},
   headStyles:{fillColor:excelBlue,textColor:excelHeaderText,font:'helvetica',fontStyle:'bold',fontSize:9.2,cellPadding:{top:4,right:1.5,bottom:4,left:1.5},halign:'center',valign:'middle',overflow:'linebreak',minCellHeight:25},
   alternateRowStyles:{fillColor:[242,242,242]},
   didParseCell:data=>{
     if(data.section==='body'&&data.row.index===totalIndex){
       data.cell.styles.fillColor=totalYellow;data.cell.styles.fontStyle='bold';data.cell.styles.fontSize=9.5;
       if(data.column.index===0){data.cell.colSpan=2;data.cell.styles.halign='center'}
       if(data.column.index===1)data.cell.styles.textColor=totalYellow;
     }
   },
   didDrawCell:data=>{
     if(data.section==='body'&&data.row.index===totalIndex&&data.column.index===0){
       d.setDrawColor(0,0,0);d.setLineWidth(.35);d.rect(data.cell.x,data.cell.y,data.cell.width,data.cell.height);
     }
     const h=headers[data.column.index];
     if(data.section==='body'&&isPctHeader(h)){
       const v=parseFloat(String(data.cell.raw));
       if(Number.isFinite(v)){d.setFillColor(...scaleColor(v));d.rect(data.cell.x,data.cell.y,data.cell.width,data.cell.height,'F');d.setTextColor(20,20,20);d.setFont('helvetica','bold');d.setFontSize(9.5);d.text(fmt(data.cell.raw,data.column.index),data.cell.x+data.cell.width/2,data.cell.y+data.cell.height/2+2,{align:'center'});}
     }
   },
   didDrawPage:()=>{d.setFont('helvetica','normal');d.setFontSize(10);d.setTextColor(100,100,100);d.text('AC-34 MATIALA • SIR-2026',8,285);d.text('Generated: '+generatedAt,210,285,{align:'center'});d.text('Page '+d.internal.getNumberOfPages(),412,285,{align:'right'})},
   margin:{left:15,right:15,top:31,bottom:12},rowPageBreak:'avoid',pageBreak:'auto'
 });
 
 if(psRows&&psRows.length){
   const psHeaders=['S No','Part No','Officer Name','BLO Name','BLO Supervisor Name','Notice Generated (NO MAP + ANOMALY)','Hearing Notice Scheduled NO MAPPING','Notice Delivered','Notice Pending Delivery','% NO MAPPING NOTICE DELIVERED','Documents Uploaded by BLO','% DOCS UPLOADED (of NOTICE DELIVERED)','Hearing Date(s)','Hearing Held + Date Lapsed','Disposal Till Yesterday','Disposal Today','Total Disposal','% Total Disposal'];
   const psBody=psRows.map((r,i)=>[i+1,r.ps,cleanOfficerName(r.officer||title),r.blo,r.supervisor,r.generated||'',r.scheduled,r.delivered,r.pending,pct(r.deliveredPct),r.docs,pct(r.docsPct),r.dates||'—',r.heldLapsed||0,r.yesterdayHeld||0,r.todayHeld||0,r.disposal||0,pct(r.disposalPct)]);
   psBody.push(['','', 'OFFICER TOTAL','','',psRows.reduce((a,r)=>a+n(r.generated),0),psRows.reduce((a,r)=>a+n(r.scheduled),0),psRows.reduce((a,r)=>a+n(r.delivered),0),psRows.reduce((a,r)=>a+n(r.pending),0),'',psRows.reduce((a,r)=>a+n(r.docs),0),'','',psRows.reduce((a,r)=>a+n(r.heldLapsed),0),psRows.reduce((a,r)=>a+n(r.yesterdayHeld),0),psRows.reduce((a,r)=>a+n(r.todayHeld),0),psRows.reduce((a,r)=>a+n(r.disposal),0),'']);
   d.addPage();
   d.setFont('helvetica','bold');d.setTextColor(...excelBlue);d.setFontSize(16);
   d.text('AC-34 MATIALA — OFFICER WISE PS DETAIL',210,14,{align:'center'});
   d.setFontSize(11);d.text(cleanOfficerName(title),210,21,{align:'center'});
   d.setFontSize(7.5);d.setTextColor(80,80,80);d.text('SIR-2026 • Complete PS-wise report',210,26,{align:'center'});
   d.text('Report Generated: '+generatedAt,291,26,{align:'right'});
   autoTable(d,{startY:31,head:[psHeaders],body:psBody,theme:'grid',tableWidth:'wrap',
     columnStyles:Object.fromEntries([9,10,26,31,31,24,23,23,21,25,23,26,27,25,24,23,24,18].map((w,i)=>[i,{cellWidth:w}])),
     styles:{font:'helvetica',fontSize:9.5,fontStyle:'bold',cellPadding:{top:3.2,right:1.6,bottom:3.2,left:1.6},overflow:'linebreak',valign:'middle',halign:'center',lineColor:[0,0,0],lineWidth:.35,textColor:[20,20,20],minCellHeight:14},
     headStyles:{fillColor:excelBlue,textColor:excelHeaderText,font:'helvetica',fontStyle:'bold',fontSize:9.2,cellPadding:{top:5,right:1.5,bottom:5,left:1.5},halign:'center',valign:'middle',overflow:'linebreak',minCellHeight:30},
     alternateRowStyles:{fillColor:stripe},
     didParseCell:data=>{if(data.section==='body'&&data.row.index===psBody.length-1){data.cell.styles.fillColor=totalYellow;data.cell.styles.fontStyle='bold'} if(data.section==='body'&&[9,11,17].includes(data.column.index)){const v=parseFloat(String(data.cell.raw));if(Number.isFinite(v))data.cell.styles.fillColor=scaleColor(v)}
     if(data.section==='body'&&data.column.index===12){
       const row=psRows[data.row.index];
       if(row && row.dates){
         const st=String(row.status||'').toLowerCase();
         const hc=hearingDateStyle(row.status);
         if(hc){data.cell.styles.fillColor=hc.bg;data.cell.styles.textColor=hc.text;}
       }
     }},
     didDrawPage:()=>{d.setFont('helvetica','normal');d.setFontSize(6.5);d.setTextColor(100,100,100);d.text('AC-34 MATIALA • SIR-2026',8,285);d.text('Generated: '+generatedAt,210,285,{align:'center'});d.text('Page '+d.internal.getNumberOfPages(),412,285,{align:'right'})},
     margin:{left:5,right:5,top:31,bottom:12},rowPageBreak:'avoid'
   });
 } d.save('AC34_'+title.replace(/[^A-Za-z0-9]+/g,'_')+'_Officer_Wise_Report.pdf');
}
function psPDF(o,rows){
 const d=new jsPDF({orientation:'landscape',unit:'mm',format:'a3'});
 const generatedAt=new Date().toLocaleString('en-IN',{timeZone:'Asia/Kolkata',day:'2-digit',month:'2-digit',year:'numeric',hour:'2-digit',minute:'2-digit',second:'2-digit',hour12:true});
 d.setFont('helvetica','bold');d.setTextColor(...excelBlue);d.setFontSize(16);
 d.text('AC-34 MATIALA — OFFICER WISE PS DETAIL',148,14,{align:'center'});
 d.setFontSize(11);d.text(cleanOfficerName(o.name),148,21,{align:'center'});
 d.setFontSize(7.5);d.setTextColor(80,80,80);d.text('SIR-2026 • Complete PS-wise report',148,26,{align:'center'});
 d.setFontSize(7.5);d.text('Report Generated: '+generatedAt,291,26,{align:'right'});
 const headers=['S No','Part No','Officer Name','BLO Name','BLO Supervisor Name','Notice Generated (NO MAP + ANOMALY)','Hearing Notice Scheduled NO MAPPING','Notice Delivered','Notice Pending Delivery','% NO MAPPING NOTICE DELIVERED','Documents Uploaded by BLO','% DOCS UPLOADED (of NOTICE DELIVERED)','Hearing Date(s)','Hearing Held + Date Lapsed','Total Disposal','% Total Disposal'];
 const body=rows.map((r,i)=>[i+1,r.ps,cleanOfficerName(r.officer||o.name),r.blo,r.supervisor,r.generated||'',r.scheduled,r.delivered,r.pending,pct(r.deliveredPct),r.docs,pct(r.docsPct),r.dates||'—',r.heldLapsed||0,r.disposal||0,pct(r.disposalPct)]);
 body.push(['','','OFFICER TOTAL','','',rows.reduce((a,r)=>a+n(r.generated),0),rows.reduce((a,r)=>a+n(r.scheduled),0),rows.reduce((a,r)=>a+n(r.delivered),0),rows.reduce((a,r)=>a+n(r.pending),0),'',rows.reduce((a,r)=>a+n(r.docs),0),'','',rows.reduce((a,r)=>a+n(r.heldLapsed),0),rows.reduce((a,r)=>a+n(r.disposal),0),'']);
 const totalIndex=body.length-1;
 const widths=[16,18,55,62,64,44,42,44,38,42,40,42,46,44,42,32];
 autoTable(d,{startY:42,head:[headers],body,theme:'grid',tableWidth:'wrap',
 columnStyles:Object.fromEntries(widths.map((w,i)=>[i,{cellWidth:w}])),
 styles:{font:'helvetica',fontSize:11,fontStyle:'bold',cellPadding:{top:2.5,right:1.5,bottom:2.5,left:1.5},overflow:'linebreak',valign:'middle',halign:'center',lineColor:[0,0,0],lineWidth:.35,textColor:[20,20,20],minCellHeight:12},
 headStyles:{fillColor:excelBlue,textColor:excelHeaderText,font:'helvetica',fontStyle:'bold',fontSize:10.5,cellPadding:{top:7,right:2,bottom:7,left:2},halign:'center',valign:'middle',overflow:'linebreak',minCellHeight:34},
 alternateRowStyles:{fillColor:stripe},
 didParseCell:data=>{
   if(data.section==='body'&&data.row.index===totalIndex){data.cell.styles.fillColor=totalYellow;data.cell.styles.fontStyle='bold'}
   if(data.section==='body'&&[9,11,15].includes(data.column.index)){const v=parseFloat(String(data.cell.raw));if(Number.isFinite(v))data.cell.styles.fillColor=scaleColor(v)}
 },
 didDrawPage:()=>{d.setFont('helvetica','normal');d.setFontSize(6.5);d.setTextColor(100,100,100);d.text('AC-34 MATIALA • SIR-2026',8,285);d.text('Generated: '+generatedAt,210,285,{align:'center'});d.text('Page '+d.internal.getNumberOfPages(),412,285,{align:'right'})},
 margin:{left:10,right:10,top:42,bottom:20},rowPageBreak:'avoid'
 });
 d.save('AC34_'+cleanOfficerName(o.name).replace(/[^A-Za-z0-9]+/g,'_')+'_PS_Detail_'+Date.now()+'.pdf');
}
export default function Page(){
 const[data,setData]=useState(null),[sel,setSel]=useState(''),[q,setQ]=useState(''),[file,setFile]=useState(''),[tab,setTab]=useState('dash'),[busy,setBusy]=useState(false),[reference,setReference]=useState(null),[eci,setEci]=useState(null),[blo,setBlo]=useState(null),[baseline,setBaseline]=useState(null),[refFile,setRefFile]=useState(''),[eciFile,setEciFile]=useState(''),[bloFile,setBloFile]=useState(''),[baselineFile,setBaselineFile]=useState(''),[sortKey,setSortKey]=useState('ps'),[sortDir,setSortDir]=useState('asc');
 const reportWrap=useRef(null);
 useEffect(()=>{if(tab==='report'&&reportWrap.current) reportWrap.current.scrollLeft=0},[tab,data]);
 const officer=data?.officers.find(x=>x.name===sel)||data?.officers[0];
 const idx=officer?data.officers.findIndex(x=>x.name===officer.name):-1;
 const normName=s=>cleanOfficerName(s).toLowerCase();
 const allOfficerRows=useMemo(()=>data?.details.filter(x=>normName(x.officer)===normName(officer?.name))||[],[data,officer]);
 const rows=useMemo(()=>{
 const filtered=allOfficerRows.filter(r=>String(r.ps).includes(q)||String(r.blo).toLowerCase().includes(q.toLowerCase())||String(r.supervisor).toLowerCase().includes(q.toLowerCase()));
 const numeric=new Set(['ps','generated','scheduled','delivered','pending','deliveredPct','docs','docsPct','heldLapsed','disposal','disposalPct']);
 return [...filtered].sort((a,b)=>{
   let av=a[sortKey],bv=b[sortKey];
   if(numeric.has(sortKey)){av=n(av);bv=n(bv);return sortDir==='asc'?av-bv:bv-av}
   av=String(av??'').toLowerCase();bv=String(bv??'').toLowerCase();
   return sortDir==='asc'?av.localeCompare(bv,undefined,{numeric:true}):bv.localeCompare(av,undefined,{numeric:true});
 });
},[allOfficerRows,q,sortKey,sortDir]);
 useEffect(()=>{try{const s=localStorage.getItem('ac34_reference_v1');if(s)setReference(JSON.parse(s))}catch{}},[]);
useEffect(()=>{try{const s=localStorage.getItem('ac34_baseline_2am_v1');if(s)setBaseline(new Map(JSON.parse(s)))}catch{}},[]);
useEffect(()=>{if(reference&&eci&&blo){const x=buildData(reference,eci,blo,baseline);setData(x);setSel(x.officers[0]?.name||'');setQ('');}},[reference,eci,blo,baseline]);
function readFile(f,kind){
 if(!f)return;setBusy(true);
 f.arrayBuffer().then(b=>XLSX.read(b,{type:'array',cellDates:true})).then(wb=>{
   if(kind==='reference'){const r=parseReference(wb);setReference(r);localStorage.setItem('ac34_reference_v1',JSON.stringify(r));setRefFile(f.name)}
   if(kind==='eci'){setEci(parseECI(wb));setEciFile(f.name)}
   if(kind==='baseline'){const m=parseECI(wb);const arr=[...m.entries()];setBaseline(new Map(arr));localStorage.setItem('ac34_baseline_2am_v1',JSON.stringify(arr));setBaselineFile(f.name)}
   if(kind==='blo'){setBlo(parseBLO(wb));setBloFile(f.name)}
 }).catch(err=>alert(err.message||'Excel could not be read')).finally(()=>{setBusy(false)});
}
function uploadLegacy(e){const f=e.target.files?.[0];if(!f)return;readFile(f,'reference');e.target.value=''}
function uploadSource(e,kind){const f=e.target.files?.[0];if(!f)return;readFile(f,kind);e.target.value=''}

 function downloadOfficerWise(o,i){reportPDF(cleanOfficerName(o.name),data.reportHeaders,[data.reportRows[i]],null,allOfficerRows)}
 function downloadAllOfficerWise(){data.officers.forEach((o,i)=>setTimeout(()=>{const rr=data.details.filter(x=>normName(x.officer)===normName(o.name));reportPDF(cleanOfficerName(o.name),data.reportHeaders,[data.reportRows[i]],null,rr)},i*500))}
 function downloadConsolidated(){reportPDF('ALL 6 OFFICERS',data.reportHeaders,data.reportRows,data.grandRow)}
 return <main>
 <header className="topbar"><div><div className="eyebrow">SIR-2026 • AC-34 MATIALA</div><h1>Officer Command Dashboard</h1><p>Upload the ECI report and BLO Documents report. Existing PS mapping and Hearing Dates are preserved.</p></div><div className="source-actions">
<label className="upload">{busy?'READING…':'1. REFERENCE / HEARING DATA'}<input type="file" accept=".xlsx,.xls" onChange={uploadLegacy}/></label>
<label className="upload">{busy?'READING…':'2. YESTERDAY 2 AM BASELINE'}<input type="file" accept=".xlsx,.xls" onChange={e=>uploadSource(e,'baseline')}/></label>
<label className="upload">{busy?'READING…':'3. UPLOAD ECI UPDATED'}<input type="file" accept=".xlsx,.xls" onChange={e=>uploadSource(e,'eci')}/></label>
<label className="upload">{busy?'READING…':'4. UPLOAD BLO DOCUMENTS'}<input type="file" accept=".xlsx,.xls" onChange={e=>uploadSource(e,'blo')}/></label>
</div></header>
 {!data?<section className="empty"><h2>LOAD THE 3 DATA SOURCES</h2><p>Reference is needed once. Upload the 2 AM ECI report as the baseline once; after that, upload the latest ECI and BLO files repeatedly. Today is calculated against the baseline; Total is the latest ECI value.</p><div className="setup-grid">
<label className="upload big">1. REFERENCE / HEARING DATA<input type="file" accept=".xlsx,.xls" onChange={uploadLegacy}/></label>
<label className="upload big">2. YESTERDAY 2 AM BASELINE<input type="file" accept=".xlsx,.xls" onChange={e=>uploadSource(e,'baseline')}/></label>
<label className="upload big">3. ECI UPDATED FILE<input type="file" accept=".xlsx,.xls" onChange={e=>uploadSource(e,'eci')}/></label>
<label className="upload big">4. BLO DOCUMENTS FILE<input type="file" accept=".xlsx,.xls" onChange={e=>uploadSource(e,'blo')}/></label>
</div><div className="status-box"><b>Reference:</b> {refFile|| (reference?'Saved in this browser':'Not loaded')} &nbsp; • &nbsp; <b>Baseline:</b> {baselineFile|| (baseline?'Saved in this browser':'Not loaded')} &nbsp; • &nbsp; <b>ECI:</b> {eciFile||'Not loaded'} &nbsp; • &nbsp; <b>BLO:</b> {bloFile||'Not loaded'}</div></section>:<>
 <div className="filebar">ECI: <b>{eciFile||'—'}</b> • BLO DOCUMENTS: <b>{bloFile||'—'}</b> • BASELINE: <b>{baselineFile|| (baseline?'saved 2 AM baseline':'—')}</b> • REFERENCE: <b>{refFile||'saved reference'}</b> • <b>{data.officers.length}</b> officers • <b>{data.details.length}</b> PS updated</div>
 <nav className="tabs"><button className={tab==='dash'?'active':''} onClick={()=>setTab('dash')}>6 OFFICER DASHBOARDS</button><button className={tab==='report'?'active':''} onClick={()=>setTab('report')}>OFFICER WISE REPORT</button></nav>
 {tab==='dash'?<>
 <section className="officer-grid">{data.officers.map((o,i)=><button key={o.name} className={'officer-card '+(officer.name===o.name?'active':'')} onClick={()=>{setSel(o.name);setQ('')}}><b>{o.name}</b><strong>{o.ps}</strong><small>PS</small></button>)}</section>
 <section className="summary">{[['Officer',officer.name],['PS',officer.ps],['Generated',officer.generated],['Scheduled',officer.scheduled],['Delivered',officer.delivered],['% Delivered',pct(officer.deliveredPct)],['Docs',officer.docs],['Hearings',officer.held]].map(x=><div key={x[0]}><small>{x[0]}</small><b>{x[1]}</b></div>)}<button className="download" onClick={()=>downloadOfficerWise(officer,idx)}>DOWNLOAD OFFICER + PS PDF</button></section>
 <section className="table-section report-excel officer-dashboard-report officer-selected-report"><div className="excel-report-title">AC-34 MATIALA — OFFICER WISE REPORT</div><div className="report-head"><div><h2>{cleanOfficerName(officer.name)}</h2><small>Documents columns hidden in dashboard • Disposal shown as Till Yesterday • Today • Total</small></div></div><div className="table-wrap report-scroll"><table><thead><tr>{officerVisibleIndexes(data.reportHeaders).map(j=><th key={j}>{data.reportHeaders[j]}</th>)}</tr></thead><tbody><tr>{officerVisibleIndexes(data.reportHeaders).map(j=>{const v=data.reportRows[idx][j],h=data.reportHeaders[j],numv=parseFloat(String(v??'')),isPct=['% NO MAPPING DELIVERED','% Total Disposal'].includes(h);return <td key={j} style={isPct&&Number.isFinite(numv)?{background:'rgb('+scaleColor(numv).join(',')+')'}:undefined}>{isPct&&Number.isFinite(numv)?numv.toFixed(2):String(v??'')}</td>})}</tr></tbody></table></div></section> <section className="table-section"><div className="table-head"><div><h2>{cleanOfficerName(officer.name)} — PS DETAIL</h2><small>{rows.length} PS • Sortable</small></div><div className="ps-controls"><input value={q} onChange={e=>setQ(e.target.value)} placeholder="Search PS / BLO / Supervisor"/><select value={sortKey} onChange={e=>setSortKey(e.target.value)}><option value="ps">Sort: Part No</option><option value="generated">Sort: Notice Generated</option><option value="scheduled">Sort: Hearing Scheduled</option><option value="delivered">Sort: Notice Delivered</option><option value="pending">Sort: Pending Delivery</option><option value="deliveredPct">Sort: % Delivered</option><option value="docs">Sort: Documents Uploaded</option><option value="docsPct">Sort: % Docs Uploaded</option><option value="heldLapsed">Sort: Hearing Held + Lapsed</option><option value="disposal">Sort: Total Disposal</option><option value="disposalPct">Sort: % Total Disposal</option><option value="blo">Sort: BLO Name</option><option value="supervisor">Sort: BLO Supervisor</option><option value="dates">Sort: Hearing Date</option></select><button className="sort-dir" onClick={()=>setSortDir(x=>x==='asc'?'desc':'asc')}>{sortDir==='asc'?'↑ ASC':'↓ DESC'}</button></div></div><div className="table-wrap report-scroll"><table><thead><tr><th>S No</th><th>Part No</th><th>Officer Name</th><th>BLO Name</th><th>BLO Supervisor Name</th><th>Notice Generated (NO MAP + ANOMALY)</th><th>Hearing Notice Scheduled NO MAPPING</th><th>Notice Delivered</th><th>Notice Pending Delivery</th><th>% NO MAPPING NOTICE DELIVERED</th><th>Documents Uploaded by BLO</th><th>% DOCS UPLOADED (of NOTICE DELIVERED)</th><th>Hearing Date(s)</th><th>Hearing Held + Date Lapsed</th><th>Disposal Till Yesterday</th><th>Disposal Today</th><th>Total Disposal</th><th>% Total Disposal</th></tr></thead><tbody>{rows.map((r,i)=><tr key={r.ps}><td>{i+1}</td><td>{r.ps}</td><td>{cleanOfficerName(r.officer||officer.name)}</td><td>{r.blo}</td><td>{r.supervisor}</td><td>{r.generated}</td><td>{r.scheduled}</td><td>{r.delivered}</td><td>{r.pending}</td><td style={{background:'rgb('+scaleColor(r.deliveredPct).join(',')+')',fontWeight:800}}>{pct(r.deliveredPct)}</td><td>{r.docs}</td><td style={{background:'rgb('+scaleColor(r.docsPct).join(',')+')',fontWeight:800}}>{pct(r.docsPct)}</td><td style={{background:(hearingDateStyle(r.status)?'rgb('+hearingDateStyle(r.status).bg.join(',')+')':undefined),color:(hearingDateStyle(r.status)?'rgb('+hearingDateStyle(r.status).text.join(',')+')':undefined),fontWeight:800}}>{r.dates||'—'}</td><td>{r.heldLapsed}</td><td>{r.yesterdayHeld??0}</td><td>{r.todayHeld??0}</td><td>{r.disposal}</td><td style={{background:'rgb('+scaleColor(r.disposalPct).join(',')+')',fontWeight:800}}>{pct(r.disposalPct)}</td></tr>)}</tbody></table></div></section>
 </>:<section className="table-section report-excel officer-dashboard-report"><div className="excel-report-title">AC-34 MATIALA — OFFICER WISE REPORT</div><div className="report-head"><div><h2>OFFICER WISE REPORT — ALL 6 OFFICERS</h2><small>Documents Uploaded by BLO and % Documents Uploaded are hidden in dashboard</small></div><div className="report-actions"><button className="download" onClick={downloadConsolidated}>DOWNLOAD CONSOLIDATED PDF</button><button className="download" onClick={downloadAllOfficerWise}>DOWNLOAD ALL OFFICER PDFs</button></div></div><div className="table-wrap report-scroll" ref={reportWrap}><table><thead><tr>{officerVisibleIndexes(data.reportHeaders).map(j=><th key={j}>{data.reportHeaders[j]}</th>)}</tr></thead><tbody>{data.reportRows.map((r,i)=><tr key={i}>{officerVisibleIndexes(data.reportHeaders).map(j=>{const v=r[j],h=data.reportHeaders[j];const val=String(v??'');const isPct=['% NO MAPPING DELIVERED','% Total Disposal'].includes(h);const numv=parseFloat(val);const display=isPct&&Number.isFinite(numv)?numv.toFixed(2):val;return <td key={j} style={isPct&&Number.isFinite(numv)?{background:'rgb('+scaleColor(numv).join(',')+')',fontWeight:700}:undefined}>{display}</td>})}</tr>)}{data.grandRow&&<tr className="grand"><td colSpan={2}>GRAND TOTAL</td>{officerVisibleIndexes(data.reportHeaders).filter(j=>j>=2).map(j=>{const v=data.grandRow[j],h=data.reportHeaders[j],val=String(v??''),isPct=['% NO MAPPING DELIVERED','% Total Disposal'].includes(h),numv=parseFloat(val);return <td key={j}>{isPct&&Number.isFinite(numv)?numv.toFixed(2):val}</td>})}</tr>}</tbody></table></div></section>}
 </>}
 <footer>Upload Excel → visible columns only → same report headings → same Excel-style colours → individual + consolidated PDFs.</footer>
 </main>
}
