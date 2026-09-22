'use client';
import {useMemo,useState} from 'react';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';

const n=v=>Number.isFinite(Number(v))?Number(v):0,pct=v=>n(v).toFixed(1)+'%';

function parse(wb){
 const S=x=>wb.Sheets[x]?XLSX.utils.sheet_to_json(wb.Sheets[x],{header:1,defval:null}):[];
 const a=S('Officer Wise Report'),b=S('Officer Wise PS Detail'),m=S('PS Mapping'),h=S('Hearing Dates');
 const cols=wb.Sheets['Officer Wise Report']?.['!cols']||[];
 const headerIndex=Math.max(0,a.findIndex(r=>r?.some(v=>String(v??'').trim().toLowerCase()==='officer name')));
 const headers=a[headerIndex]||[];
 const visibleIndexes=headers.map((_,i)=>i).filter(i=>!cols[i]?.hidden && String(headers[i]??'').trim()!=='');
 const reportHeaders=visibleIndexes.map(i=>String(headers[i]));
 const reportRows=a.slice(headerIndex+1).filter(r=>r?.[1]).map(r=>visibleIndexes.map(i=>r[i]??''));
 const officers=a.slice(headerIndex+1).filter(r=>r?.[1]).map(r=>({sno:r[0],name:String(r[1]),ps:n(r[2]),generated:n(r[3]),pendingGen:n(r[4]),scheduled:n(r[5]),delivered:n(r[6]),deliveredPct:n(r[7]),pendingDelivery:n(r[8]),held:n(r[9]),lapsed:n(r[10]),reschedule:n(r[11]),deoPending:n(r[12]),deoGt5:n(r[13]),deoVerified:n(r[14]),docs:n(r[15])}));
 const mm=new Map(m.slice(1).filter(r=>r?.[0]!=null).map(r=>[n(r[0]),{blo:r[2]||'',supervisor:r[3]||''}])),hh=new Map(h.slice(1).filter(r=>r?.[0]!=null).map(r=>[n(r[0]),{dates:r[4]||'',status:r[5]||''}]));
 let cur='',details=[];
 b.slice(2).forEach(r=>{if(!r||r[0]==null)return;if(r[1]==='OFFICER TOTAL'){cur=r[2]||'';return}if(typeof r[1]==='number'){const ps=n(r[1]);details.push({ps,officer:r[2]||cur,blo:r[3]||mm.get(ps)?.blo||'',supervisor:r[4]||mm.get(ps)?.supervisor||'',scheduled:n(r[8]),delivered:n(r[9]),pending:n(r[10]),deliveredPct:n(r[12]),docs:n(r[13]),docsPct:n(r[14]),...(hh.get(ps)||{})})}});
 return{officers,details,reportHeaders,reportRows};
}

function officerPDF(o,headers,row){
 const d=new jsPDF({orientation:'landscape',unit:'mm',format:'a4'});
 d.setFontSize(16);d.text('AC-34 MATIALA — SIR-2026',14,13);
 d.setFontSize(11);d.text('OFFICER WISE REPORT — '+o.name,14,20);
 autoTable(d,{startY:26,head:[headers],body:[row],styles:{fontSize:5.8,cellPadding:1.7,overflow:'linebreak'},headStyles:{fontSize:5.5},margin:{left:6,right:6}});
 d.save('AC34_Officer_'+o.name.replace(/[^A-Za-z0-9]+/g,'_')+'.pdf');
}
function psPDF(o,rows){
 const d=new jsPDF({orientation:'landscape',unit:'mm',format:'a4'});
 d.setFontSize(16);d.text('AC-34 MATIALA — SIR-2026',14,13);d.setFontSize(11);d.text('PS DETAIL — '+o.name,14,20);
 autoTable(d,{startY:26,head:[['PS','BLO','Supervisor','Scheduled','Delivered','Pending','% Delivered','Docs Uploaded','% Docs','Hearing','Date(s)']],body:rows.map(r=>[r.ps,r.blo,r.supervisor,r.scheduled,r.delivered,r.pending,pct(r.deliveredPct),r.docs,pct(r.docsPct),r.status||'—',r.dates||'—']),styles:{fontSize:6,cellPadding:1.4},headStyles:{fontSize:6.5},margin:{left:8,right:8}});
 d.save('AC34_'+o.name.replace(/[^A-Za-z0-9]+/g,'_')+'_PS.pdf');
}

export default function Page(){
 const[data,setData]=useState(null),[sel,setSel]=useState(''),[q,setQ]=useState(''),[file,setFile]=useState(''),[tab,setTab]=useState('dash'),[busy,setBusy]=useState(false);
 const officer=data?.officers.find(x=>x.name===sel)||data?.officers[0];
 const officerIndex=officer?data.officers.findIndex(x=>x.name===officer.name):-1;
 const rows=useMemo(()=>data?.details.filter(x=>x.officer===officer?.name).filter(r=>String(r.ps).includes(q)||String(r.blo).toLowerCase().includes(q.toLowerCase())||String(r.supervisor).toLowerCase().includes(q.toLowerCase()))||[],[data,officer,q]);
 function upload(e){const f=e.target.files?.[0];if(!f)return;setBusy(true);setFile(f.name);f.arrayBuffer().then(b=>{const x=parse(XLSX.read(b,{type:'array',cellDates:true}));if(!x.officers.length)throw Error('Officer Wise Report sheet not found');setData(x);setSel(x.officers[0].name)}).catch(e=>alert(e.message)).finally(()=>{setBusy(false);e.target.value=''})}
 return <main>
 <header className="topbar"><div><div className="eyebrow">SIR-2026 • AC-34 MATIALA</div><h1>Officer Command Dashboard</h1><p>Upload one updated Excel. Hidden Excel columns are automatically excluded.</p></div><label className="upload">{busy?'READING…':'UPLOAD UPDATED EXCEL'}<input type="file" accept=".xlsx,.xls" onChange={upload}/></label></header>
 {!data?<section className="empty"><h2>UPLOAD UPDATED EXCEL</h2><p>Start by uploading the latest report workbook.</p><label className="upload big">SELECT EXCEL FILE<input type="file" accept=".xlsx,.xls" onChange={upload}/></label></section>:<>
 <div className="filebar">CURRENT FILE: <b>{file}</b> • <b>{data.officers.length}</b> officers loaded</div>
 <nav className="tabs"><button className={tab==='dash'?'active':''} onClick={()=>setTab('dash')}>6 OFFICER DASHBOARDS</button><button className={tab==='report'?'active':''} onClick={()=>setTab('report')}>OFFICER WISE REPORT</button><button className="all" onClick={()=>data.officers.forEach((o,i)=>setTimeout(()=>{officerPDF(o,data.reportHeaders,data.reportRows[i]);psPDF(o,data.details.filter(r=>r.officer===o.name))},i*500))}>DOWNLOAD ALL PDF REPORTS</button></nav>
 {tab==='dash'?<><section className="officer-grid">{data.officers.map(o=><button key={o.name} className={'officer-card '+(officer.name===o.name?'active':'')} onClick={()=>{setSel(o.name);setQ('')}}><b>{o.name}</b><strong>{o.ps}</strong><small>PS</small><span>Delivered <em>{o.delivered}</em> • Docs <em>{o.docs}</em> • Held <em>{o.held}</em></span></button>)}</section><section className="summary">{[['Officer',officer.name],['PS',officer.ps],['Generated',officer.generated],['Scheduled',officer.scheduled],['Delivered',officer.delivered],['% Delivered',pct(officer.deliveredPct)],['Docs',officer.docs],['Hearings',officer.held]].map(x=><div key={x[0]}><small>{x[0]}</small><b>{x[1]}</b></div>)}<button className="download" onClick={()=>psPDF(officer,rows)}>DOWNLOAD PS PDF</button></section><section className="table-section"><div className="table-head"><div><h2>{officer.name} — PS DETAIL</h2><small>{rows.length} PS</small></div><input value={q} onChange={e=>setQ(e.target.value)} placeholder="Search PS / BLO / Supervisor"/></div><div className="table-wrap"><table><thead><tr><th>PS</th><th>BLO</th><th>Supervisor</th><th>Scheduled</th><th>Delivered</th><th>Pending</th><th>% Delivered</th><th>Docs Uploaded</th><th>% Docs</th><th>Hearing</th><th>Date(s)</th></tr></thead><tbody>{rows.map(r=><tr key={r.ps}><td><b>{r.ps}</b></td><td>{r.blo}</td><td>{r.supervisor}</td><td>{r.scheduled}</td><td>{r.delivered}</td><td>{r.pending}</td><td><i className={'pill '+(r.deliveredPct>=80?'good':r.deliveredPct>=60?'mid':'low')}>{pct(r.deliveredPct)}</i></td><td>{r.docs}</td><td>{pct(r.docsPct)}</td><td>{r.status||'—'}</td><td>{r.dates||'—'}</td></tr>)}</tbody></table></div></section></>:<section className="table-section"><div className="report-head"><div><h2>OFFICER WISE REPORT — ALL 6 OFFICERS</h2><small>Only columns visible in the uploaded Excel are shown</small></div><button className="download" onClick={()=>data.officers.forEach((o,i)=>setTimeout(()=>officerPDF(o,data.reportHeaders,data.reportRows[i]),i*500))}>DOWNLOAD OFFICER WISE PDFs</button></div><div className="table-wrap"><table><thead><tr>{data.reportHeaders.map(x=><th key={x}>{x}</th>)}</tr></thead><tbody>{data.reportRows.map((r,i)=><tr key={i}>{r.map((v,j)=><td key={j}>{String(v??'')}</td>)}</tr>)}</tbody></table></div></section>}
 </>}
 <footer>Upload Excel → hidden columns excluded → dashboards → Officer Wise Report → PDFs.</footer>
 </main>
}
