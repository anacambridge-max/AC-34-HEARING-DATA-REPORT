'use client';
import {useMemo,useState,useRef,useEffect} from 'react';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';

const n=v=>Number.isFinite(Number(v))?Number(v):0;
const pct=v=>n(v).toFixed(2)+'%';
const rgb=(r,g,b)=>[r,g,b];
const excelBlue=rgb(31,56,100), excelHeaderText=rgb(255,255,255), stripe=rgb(242,242,242), totalYellow=rgb(255,192,0);
const red=rgb(248,105,107), yellow=rgb(255,235,132), green=rgb(99,190,123);

function scaleColor(v){
 const x=Math.max(0,Math.min(100,n(v)));
 if(x<=50){const t=x/50;return rgb(Math.round(248+(255-248)*t),Math.round(105+(235-105)*t),Math.round(107+(132-107)*t))}
 const t=(x-50)/50;return rgb(Math.round(255+(99-255)*t),Math.round(235+(190-235)*t),Math.round(132+(123-132)*t));
}
function parse(wb){
 const S=x=>wb.Sheets[x]?XLSX.utils.sheet_to_json(wb.Sheets[x],{header:1,defval:null}):[];
 const a=S('Officer Wise Report'),b=S('Officer Wise PS Detail'),m=S('PS Mapping'),h=S('Hearing Dates');
 const sheet=a, ws=wb.Sheets['Officer Wise Report'];
 const cols=ws?.['!cols']||[];
 const headerIndex=Math.max(0,a.findIndex(r=>r?.some(v=>String(v??'').trim().toLowerCase()==='officer name')));
 const headers=a[headerIndex]||[];
 const wantedHeaders=['S No','Officer Name','No. of PS','Notice Generated (NO MAP + ANOMALY)','Hearing Notice Scheduled NO MAPPING','NO MAP NOTICE DELIVERED','% NO MAPPING DELIVERED','Documents Uploaded by BLO','% Docs Uploaded (of Notice Delivered)','Hearing Held + Date Lapsed','Total Disposal','% Total Disposal'];
 const norm=s=>String(s??'').replace(/\\s+/g,' ').trim().toLowerCase();
 const visibleIndexes=wantedHeaders.map(w=>headers.findIndex(h=>norm(h)===norm(w))).filter(i=>i>=0);
 const reportHeaders=visibleIndexes.map(i=>String(headers[i]));
 const rawOfficerRows=a.slice(headerIndex+1).filter(r=>r?.some(v=>String(v??'').trim()!=='') );
 const isGrand=r=>r?.some(v=>String(v??'').trim().toUpperCase().replace(/\s+/g,' ')==='GRAND TOTAL');
 const officerRows=rawOfficerRows.filter(r=>!isGrand(r) && r?.[1]);
 const reportRows=officerRows.map(r=>visibleIndexes.map(i=>r[i]??''));
 const grand=rawOfficerRows.find(isGrand);
 const grandRow=grand?visibleIndexes.map(i=>grand[i]??''):null;
 const officers=officerRows.map(r=>({sno:r[0],name:String(r[1]),ps:n(r[2]),generated:n(r[3]),pendingGen:n(r[4]),scheduled:n(r[5]),delivered:n(r[6]),deliveredPct:n(r[7]),pendingDelivery:n(r[8]),held:n(r[9]),lapsed:n(r[10]),reschedule:n(r[11]),deoPending:n(r[12]),deoGt5:n(r[13]),deoVerified:n(r[14]),docs:n(r[15])}));
 const mm=new Map(m.slice(1).filter(r=>r?.[0]!=null).map(r=>[n(r[0]),{blo:r[2]||'',supervisor:r[3]||''}])),hh=new Map(h.slice(1).filter(r=>r?.[0]!=null).map(r=>[n(r[0]),{dates:r[4]||'',status:r[5]||''}]));
 let cur='',details=[];
 b.slice(2).forEach(r=>{if(!r||r[0]==null)return;if(r[1]==='OFFICER TOTAL'){cur=r[2]||'';return}if(typeof r[1]==='number'){const ps=n(r[1]);details.push({ps,officer:r[2]||cur,blo:r[3]||mm.get(ps)?.blo||'',supervisor:r[4]||mm.get(ps)?.supervisor||'',scheduled:n(r[8]),delivered:n(r[9]),pending:n(r[10]),deliveredPct:n(r[12]),docs:n(r[13]),docsPct:n(r[14]),...(hh.get(ps)||{})})}});
 return{officers,details,reportHeaders,reportRows,grandRow,visibleIndexes};
}

function reportPDF(title,headers,rows,grandRow){
 const d=new jsPDF({orientation:'landscape',unit:'mm',format:'a4'});
 d.setFont('helvetica','bold');d.setTextColor(...excelBlue);d.setFontSize(17);
 d.text('AC-34 MATIALA — OFFICER WISE REPORT',148,13,{align:'center'});
 d.setFontSize(11);d.text(title,148,21,{align:'center'});
 d.setFontSize(8);d.setTextColor(80,80,80);d.text('SIR-2026 • Officer-wise monitoring report',148,26,{align:'center'});
 const isPctHeader=h=>['% NO MAPPING DELIVERED','% Docs Uploaded (of Notice Delivered)','% Total Disposal'].includes(String(h||'').trim());
 const fmt=(v,i)=>{const num=parseFloat(String(v??''));return isPctHeader(headers[i])&&Number.isFinite(num)?num.toFixed(2):String(v??'')};
 let body=rows.map(r=>r.map((v,i)=>fmt(v,i)));
 if(grandRow){const gr=grandRow.map((v,i)=>fmt(v,i));gr[0]='GRAND TOTAL';gr[1]='';body.push(gr)}
 const totalIndex=grandRow?body.length-1:-1;
 const widths=[9,41,12,24,24,21,20,23,30,24,17,19];
 autoTable(d,{
   startY:31,head:[headers],body,theme:'grid',tableWidth:'wrap',
   columnStyles:Object.fromEntries(widths.map((w,i)=>[i,{cellWidth:w}])),
   styles:{font:'helvetica',fontSize:7,fontStyle:'bold',cellPadding:{top:3,right:1.8,bottom:3,left:1.8},overflow:'linebreak',valign:'middle',halign:'center',lineColor:[0,0,0],lineWidth:.35,textColor:[20,20,20],minCellHeight:12},
   headStyles:{fillColor:excelBlue,textColor:excelHeaderText,font:'helvetica',fontStyle:'bold',fontSize:7,cellPadding:{top:4,right:1.5,bottom:4,left:1.5},halign:'center',valign:'middle',overflow:'linebreak',minCellHeight:25},
   alternateRowStyles:{fillColor:[242,242,242]},
   didParseCell:data=>{
     if(data.section==='body'&&data.row.index===totalIndex){
       data.cell.styles.fillColor=totalYellow;data.cell.styles.fontStyle='bold';data.cell.styles.fontSize=7.2;
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
       if(Number.isFinite(v)){d.setFillColor(...scaleColor(v));d.rect(data.cell.x,data.cell.y,data.cell.width,data.cell.height,'F');d.setTextColor(20,20,20);d.setFont('helvetica','bold');d.setFontSize(7.2);d.text(fmt(data.cell.raw,data.column.index),data.cell.x+data.cell.width/2,data.cell.y+data.cell.height/2+2,{align:'center'});}
     }
   },
   didDrawPage:()=>{d.setFont('helvetica','normal');d.setFontSize(6.5);d.setTextColor(100,100,100);d.text('AC-34 MATIALA • SIR-2026',8,204);d.text('Page '+d.internal.getNumberOfPages(),289,204,{align:'right'})},
   margin:{left:5,right:5,top:31,bottom:12},rowPageBreak:'avoid'
 });
 d.save('AC34_'+title.replace(/[^A-Za-z0-9]+/g,'_')+'_Officer_Wise_Report.pdf');
}
function psPDF(o,rows){
 const d=new jsPDF({orientation:'landscape',unit:'mm',format:'a3'});
 const generatedAt=new Date().toLocaleString('en-IN',{timeZone:'Asia/Kolkata',day:'2-digit',month:'2-digit',year:'numeric',hour:'2-digit',minute:'2-digit',second:'2-digit',hour12:true});
 d.setFont('helvetica','bold');d.setTextColor(...excelBlue);d.setFontSize(18);
 d.text('AC-34 MATIALA — OFFICER WISE PS DETAIL',210,14,{align:'center'});
 d.setFontSize(12);d.text(o.name,210,22,{align:'center'});
 d.setFontSize(8);d.setTextColor(80,80,80);d.text('SIR-2026 • Complete PS-wise report',210,27,{align:'center'});
 d.setFontSize(7.5);d.text('Report Generated: '+generatedAt,405,27,{align:'right'});
 const headers=['S No','Part No','Officer Name','BLO Name','BLO Supervisor Name','Notice Generated (NO MAP + ANOMALY)','Pending for Notice Generation','Hearing Notice Scheduled NO MAPPING','Notice Delivered','Notice Pending Delivery','% NO MAPPING NOTICE DELIVERED','Documents Uploaded by BLO','% DOCS UPLOADED (of NOTICE DELIVERED)','Hearing Date(s)','Hearing Held + Date Lapsed','Total Disposal','% Total Disposal'];
 const body=rows.map((r,i)=>[i+1,r.ps,r.officer||o.name,r.blo,r.supervisor,r.generated||'',r.pendingGen||0,r.scheduled,r.delivered,r.pending,pct(r.deliveredPct),r.docs,pct(r.docsPct),r.dates||'—',r.status==='Hearing Held'?(r.dates||''):(r.status||'—'),r.disposal||0,pct(r.disposalPct)]);
 body.push(['','','OFFICER TOTAL','','',rows.reduce((a,r)=>a+n(r.generated),0),rows.reduce((a,r)=>a+n(r.pendingGen),0),rows.reduce((a,r)=>a+n(r.scheduled),0),rows.reduce((a,r)=>a+n(r.delivered),0),rows.reduce((a,r)=>a+n(r.pending),0),'',rows.reduce((a,r)=>a+n(r.docs),0),'','','',rows.reduce((a,r)=>a+n(r.heldLapsed),0),rows.reduce((a,r)=>a+n(r.disposal),0),'']);
 const totalIndex=body.length-1;
 const widths=[9,10,32,36,37,25,23,25,19,23,22,22,25,24,25,17,19];
 autoTable(d,{startY:32,head:[headers],body,theme:'grid',tableWidth:'wrap',
 columnStyles:Object.fromEntries(widths.map((w,i)=>[i,{cellWidth:w}])),
 styles:{font:'helvetica',fontSize:6.4,fontStyle:'bold',cellPadding:{top:2.5,right:1.5,bottom:2.5,left:1.5},overflow:'linebreak',valign:'middle',halign:'center',lineColor:[0,0,0],lineWidth:.35,textColor:[20,20,20],minCellHeight:11},
 headStyles:{fillColor:excelBlue,textColor:excelHeaderText,font:'helvetica',fontStyle:'bold',fontSize:6.4,cellPadding:{top:4,right:1.2,bottom:4,left:1.2},halign:'center',valign:'middle',overflow:'linebreak',minCellHeight:24},
 alternateRowStyles:{fillColor:stripe},
 didParseCell:data=>{
   if(data.section==='body'&&data.row.index===totalIndex){data.cell.styles.fillColor=totalYellow;data.cell.styles.fontStyle='bold'}
   if(data.section==='body'&&[10,12,16].includes(data.column.index)){const v=parseFloat(String(data.cell.raw));if(Number.isFinite(v))data.cell.styles.fillColor=scaleColor(v)}
 },
 didDrawPage:()=>{d.setFont('helvetica','normal');d.setFontSize(6.5);d.setTextColor(100,100,100);d.text('AC-34 MATIALA • SIR-2026',8,285);d.text('Generated: '+generatedAt,210,285,{align:'center'});d.text('Page '+d.internal.getNumberOfPages(),412,285,{align:'right'})},
 margin:{left:5,right:5,top:32,bottom:12},rowPageBreak:'avoid'
 });
 d.save('AC34_'+o.name.replace(/[^A-Za-z0-9]+/g,'_')+'_PS_Detail_'+Date.now()+'.pdf');
}
export default function Page(){
 const[data,setData]=useState(null),[sel,setSel]=useState(''),[q,setQ]=useState(''),[file,setFile]=useState(''),[tab,setTab]=useState('dash'),[busy,setBusy]=useState(false);
 const reportWrap=useRef(null);
 useEffect(()=>{if(tab==='report'&&reportWrap.current) reportWrap.current.scrollLeft=0},[tab,data]);
 const officer=data?.officers.find(x=>x.name===sel)||data?.officers[0];
 const idx=officer?data.officers.findIndex(x=>x.name===officer.name):-1;
 const normName=s=>String(s??'').replace(/\s+/g,' ').trim().toLowerCase();
 const allOfficerRows=useMemo(()=>data?.details.filter(x=>normName(x.officer)===normName(officer?.name))||[],[data,officer]);
 const rows=useMemo(()=>allOfficerRows.filter(r=>String(r.ps).includes(q)||String(r.blo).toLowerCase().includes(q.toLowerCase())||String(r.supervisor).toLowerCase().includes(q.toLowerCase())),[allOfficerRows,q]);
 function upload(e){const f=e.target.files?.[0];if(!f)return;setBusy(true);setFile(f.name);f.arrayBuffer().then(b=>{const x=parse(XLSX.read(b,{type:'array',cellDates:true}));if(!x.officers.length)throw Error('Officer Wise Report sheet not found');setData(x);setSel(x.officers[0].name);setQ('')}).catch(e=>alert(e.message||'Excel could not be read')).finally(()=>{setBusy(false);e.target.value=''})}
 function downloadOfficerWise(o,i){reportPDF(o.name,data.reportHeaders,[data.reportRows[i]],null)}
 function downloadAllOfficerWise(){data.officers.forEach((o,i)=>setTimeout(()=>downloadOfficerWise(o,i),i*500))}
 function downloadConsolidated(){reportPDF('ALL 6 OFFICERS',data.reportHeaders,data.reportRows,data.grandRow)}
 return <main>
 <header className="topbar"><div><div className="eyebrow">SIR-2026 • AC-34 MATIALA</div><h1>Officer Command Dashboard</h1><p>Excel format preserved: hidden columns stay hidden, visible headings stay exactly as in Excel.</p></div><label className="upload">{busy?'READING…':'UPLOAD UPDATED EXCEL'}<input type="file" accept=".xlsx,.xls" onChange={upload}/></label></header>
 {!data?<section className="empty"><h2>UPLOAD UPDATED EXCEL</h2><p>Upload the latest workbook. The app will use the workbook's hidden/visible columns automatically.</p><label className="upload big">SELECT EXCEL FILE<input type="file" accept=".xlsx,.xls" onChange={upload}/></label></section>:<>
 <div className="filebar">CURRENT FILE: <b>{file}</b> • <b>{data.officers.length}</b> officers loaded • <b>{data.reportHeaders.length}</b> visible report columns</div>
 <nav className="tabs"><button className={tab==='dash'?'active':''} onClick={()=>setTab('dash')}>6 OFFICER DASHBOARDS</button><button className={tab==='report'?'active':''} onClick={()=>setTab('report')}>OFFICER WISE REPORT</button><button className="all" onClick={downloadConsolidated}>DOWNLOAD CONSOLIDATED PDF</button><button className="all" onClick={()=>downloadOfficerWise(officer,idx)}>DOWNLOAD SELECTED OFFICER PDF</button></nav>
 {tab==='dash'?<>
 <section className="officer-grid">{data.officers.map((o,i)=><button key={o.name} className={'officer-card '+(officer.name===o.name?'active':'')} onClick={()=>{setSel(o.name);setQ('')}}><b>{o.name}</b><strong>{o.ps}</strong><small>PS</small><span>Delivered <em>{o.delivered}</em> • Docs <em>{o.docs}</em> • Held <em>{o.held}</em></span></button>)}</section>
 <section className="summary">{[['Officer',officer.name],['PS',officer.ps],['Generated',officer.generated],['Scheduled',officer.scheduled],['Delivered',officer.delivered],['% Delivered',pct(officer.deliveredPct)],['Docs',officer.docs],['Hearings',officer.held]].map(x=><div key={x[0]}><small>{x[0]}</small><b>{x[1]}</b></div>)}<button className="download" onClick={()=>{psPDF(officer,allOfficerRows);downloadOfficerWise(officer,idx)}}>DOWNLOAD OFFICER + PS PDFs</button></section>
 <section className="table-section report-excel officer-selected-report"><div className="excel-report-title">AC-34 MATIALA — OFFICER WISE REPORT</div><div className="report-head"><div><h2>{officer.name}</h2><small>Same 12 visible Excel headings</small></div></div><div className="table-wrap report-scroll"><table><thead><tr>{data.reportHeaders.map(x=><th key={x}>{x}</th>)}</tr></thead><tbody><tr>{data.reportRows[idx].map((v,j)=>{const h=data.reportHeaders[j],numv=parseFloat(String(v??'')),isPct=['% NO MAPPING DELIVERED','% Docs Uploaded (of Notice Delivered)','% Total Disposal'].includes(h);return <td key={j} style={isPct&&Number.isFinite(numv)?{background:'rgb('+scaleColor(numv).join(',')+')'}:undefined}>{isPct&&Number.isFinite(numv)?numv.toFixed(2):String(v??'')}</td>})}</tr></tbody></table></div></section> <section className="table-section"><div className="table-head"><div><h2>{officer.name} — PS DETAIL</h2><small>{rows.length} PS</small></div><input value={q} onChange={e=>setQ(e.target.value)} placeholder="Search PS / BLO / Supervisor"/></div><div className="table-wrap"><table><thead><tr><th>PS</th><th>BLO</th><th>Supervisor</th><th>Scheduled</th><th>Delivered</th><th>Pending</th><th>% Delivered</th><th>Docs Uploaded</th><th>% Docs</th><th>Hearing</th><th>Date(s)</th></tr></thead><tbody>{rows.map(r=><tr key={r.ps}><td><b>{r.ps}</b></td><td>{r.blo}</td><td>{r.supervisor}</td><td>{r.scheduled}</td><td>{r.delivered}</td><td>{r.pending}</td><td><i className={'pill '+(r.deliveredPct>=80?'good':r.deliveredPct>=60?'mid':'low')}>{pct(r.deliveredPct)}</i></td><td>{r.docs}</td><td>{pct(r.docsPct)}</td><td>{r.status||'—'}</td><td>{r.dates||'—'}</td></tr>)}</tbody></table></div></section>
 </>:<section className="table-section report-excel"><div className="excel-report-title">AC-34 MATIALA — OFFICER WISE REPORT</div><div className="report-head"><div><h2>OFFICER WISE REPORT — ALL 6 OFFICERS</h2><small>Same visible headings, hidden columns excluded, Excel-style colours preserved</small></div><div className="report-actions"><button className="download" onClick={downloadConsolidated}>DOWNLOAD CONSOLIDATED PDF</button><button className="download" onClick={()=>downloadOfficerWise(officer,idx)}>DOWNLOAD SELECTED OFFICER PDF</button></div></div><div className="table-wrap report-scroll" ref={reportWrap}><table><thead><tr>{data.reportHeaders.map(x=><th key={x}>{x}</th>)}</tr></thead><tbody>{data.reportRows.map((r,i)=><tr key={i}>{r.map((v,j)=>{const h=data.reportHeaders[j];const val=String(v??'');const isPct=['% NO MAPPING DELIVERED','% Docs Uploaded (of Notice Delivered)','% Total Disposal'].includes(h);const numv=parseFloat(val);const display=isPct&&Number.isFinite(numv)?numv.toFixed(2):val;return <td key={j} style={isPct&&Number.isFinite(numv)?{background:'rgb('+scaleColor(numv).join(',')+')',fontWeight:700}:undefined}>{display}</td>})}</tr>)}{data.grandRow&&<tr className="grand"><td colSpan={2}>GRAND TOTAL</td>{data.grandRow.slice(2).map((v,j)=>{const col=j+2,h=data.reportHeaders[col],val=String(v??''),isPct=['% NO MAPPING DELIVERED','% Docs Uploaded (of Notice Delivered)','% Total Disposal'].includes(h),numv=parseFloat(val);return <td key={col}>{isPct&&Number.isFinite(numv)?numv.toFixed(2):val}</td>})}</tr>}</tbody></table></div></section>}
 </>}
 <footer>Upload Excel → visible columns only → same report headings → same Excel-style colours → individual + consolidated PDFs.</footer>
 </main>
}
