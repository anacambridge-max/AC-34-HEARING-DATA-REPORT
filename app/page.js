'use client';
import {useEffect,useMemo,useState} from 'react';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';

function n(v){const x=Number(v);return Number.isFinite(x)?x:0}
function pct(v){return n(v).toFixed(1)+'%'}
function parse(wb){
 const sheet=(name)=>wb.Sheets[name]?XLSX.utils.sheet_to_json(wb.Sheets[name],{header:1,defval:null}):[];
 const a=sheet('Officer Wise Report'), b=sheet('Officer Wise PS Detail'), c=sheet('PS Mapping'), d=sheet('Hearing Dates');
 const officers=a.slice(3).filter(r=>r&&r[1]).map(r=>({sno:r[0],name:r[1],ps:n(r[2]),generated:n(r[3]),scheduled:n(r[5]),delivered:n(r[6]),deliveredPct:n(r[7]),held:n(r[9]),docs:n(r[15])}));
 const details=[];let current='';
 b.slice(2).forEach(r=>{if(!r||r[0]==null)return;if(r[1]==='OFFICER TOTAL'){current=r[2]||'';return}if(typeof r[1]==='number')details.push({ps:n(r[1]),officer:r[2]||current,blo:r[3]||'',supervisor:r[4]||'',scheduled:n(r[8]),delivered:n(r[9]),pending:n(r[10]),deliveredPct:n(r[12]),docs:n(r[13]),docsPct:n(r[14])})});
 const hm=new Map(d.slice(1).filter(r=>r?.[0]!=null).map(r=>[n(r[0]),{dates:r[4]||'',hearingStatus:r[5]||''}]));
 const mm=new Map(c.slice(1).filter(r=>r?.[0]!=null).map(r=>[n(r[0]),{blo:r[2]||'',supervisor:r[3]||'',centre:r[4]||''}]));
 return {officers,details:details.map(x=>({...x,...(mm.get(x.ps)||{}),...(hm.get(x.ps)||{})}))};
}
function pdf(o,rows){
 const doc=new jsPDF({orientation:'landscape',unit:'mm',format:'a4'});
 doc.setFontSize(16);doc.text('AC-34 MATIALA — SIR-2026',14,13);
 doc.setFontSize(11);doc.text('Officer Report: '+o.name,14,20);
 doc.setFontSize(8);doc.text(`PS: ${o.ps} | Scheduled: ${o.scheduled} | Delivered: ${o.delivered} | Docs Uploaded: ${o.docs} | Hearings Held: ${o.held}`,14,26);
 autoTable(doc,{startY:31,head:[['PS','BLO','Supervisor','Scheduled','Delivered','Pending','% Delivered','Docs Uploaded','% Docs','Hearing','Date(s)']],body:rows.map(r=>[r.ps,r.blo,r.supervisor,r.scheduled,r.delivered,r.pending,pct(r.deliveredPct),r.docs,pct(r.docsPct),r.hearingStatus||'',r.dates||'']),styles:{fontSize:6,cellPadding:1.4},headStyles:{fontSize:6.5},margin:{left:8,right:8}});
 doc.save('AC34_'+o.name.replace(/[^A-Za-z0-9]+/g,'_')+'_Report.pdf');
}
export default function Page(){
 const [data,setData]=useState(null),[sel,setSel]=useState(''),[search,setSearch]=useState(''),[fileName,setFileName]=useState(''),[busy,setBusy]=useState(false);
 useEffect(()=>{fetch('/starter-data.json').then(r=>r.json()).then(raw=>setData(parse(XLSX.read(JSON.stringify(raw),{type:'string'})))) .catch(()=>{});},[]);
 // starter fallback is populated in a separate static parser below
 useEffect(()=>{fetch('/starter-data.json').then(r=>r.json()).then(raw=>{const officers=raw.officerReport.filter(r=>r?.[1]).map(r=>({sno:r[0],name:r[1],ps:n(r[2]),generated:n(r[3]),scheduled:n(r[5]),delivered:n(r[6]),deliveredPct:n(r[7]),held:n(r[9]),docs:n(r[15])}));let current='',details=[];raw.psDetail.forEach(r=>{if(!r||r[0]==null)return;if(r[1]==='OFFICER TOTAL'){current=r[2]||'';return}if(typeof r[1]==='number')details.push({ps:n(r[1]),officer:r[2]||current,blo:r[3]||'',supervisor:r[4]||'',scheduled:n(r[8]),delivered:n(r[9]),pending:n(r[10]),deliveredPct:n(r[12]),docs:n(r[13]),docsPct:n(r[14])})});const hm=new Map(raw.hearingDates.map(r=>[n(r[0]),{dates:r[4]||'',hearingStatus:r[5]||''}]));setData({officers,details:details.map(x=>({...x,...(hm.get(x.ps)||{})}))});}).catch(()=>{})},[]);
 const officer=data?.officers?.find(x=>x.name===sel)||data?.officers?.[0];
 const rows=useMemo(()=>data?.details?.filter(x=>x.officer===officer?.name).filter(r=>String(r.ps).includes(search)||r.blo.toLowerCase().includes(search.toLowerCase())||r.supervisor.toLowerCase().includes(search.toLowerCase()))||[ ],[data,officer,search]);
 function upload(e){const f=e.target.files?.[0];if(!f)return;setBusy(true);setFileName(f.name);f.arrayBuffer().then(buf=>{setData(parse(XLSX.read(buf,{type:'array',cellDates:true})));setSel('')}).catch(err=>alert(err.message)).finally(()=>setBusy(false));e.target.value=''}
 if(!data||!officer)return <main className="loading">Loading dashboard…</main>;
 return <main>
 <header className="topbar"><div><div className="eyebrow">SIR-2026 • AC-34 MATIALA</div><h1>Officer Command Dashboard</h1><p>Upload the latest Excel — everything below updates automatically.</p></div><div className="actions"><label className="upload">{busy?'Reading…':'Upload Updated Excel'}<input type="file" accept=".xlsx,.xls" onChange={upload}/></label><button className="secondary" onClick={()=>data.officers.forEach((o,i)=>setTimeout(()=>pdf(o,data.details.filter(r=>r.officer===o.name)),i*500))}>Download All 6 PDFs</button></div></header>
 {fileName&&<div className="filebar">Loaded: <b>{fileName}</b></div>}
 <section className="officer-grid">{data.officers.map(o=><button key={o.name} className={'officer-card '+(officer.name===o.name?'active':'')} onClick={()=>{setSel(o.name);setSearch('')}}><div className="card-title">{o.name}</div><div className="big">{o.ps}</div><div className="muted">PS</div><div className="mini"><span>Delivered <b>{o.delivered}</b></span><span>Docs <b>{o.docs}</b></span><span>Held <b>{o.held}</b></span></div></button>)}</section>
 <section className="summary"><div><span>Officer</span><b>{officer.name}</b></div><div><span>Total PS</span><b>{officer.ps}</b></div><div><span>Generated</span><b>{officer.generated}</b></div><div><span>Scheduled</span><b>{officer.scheduled}</b></div><div><span>Delivered</span><b>{officer.delivered}</b></div><div><span>% Delivered</span><b>{pct(officer.deliveredPct)}</b></div><div><span>Docs</span><b>{officer.docs}</b></div><div><span>Hearings</span><b>{officer.held}</b></div><button className="download" onClick={()=>pdf(officer,data.details.filter(r=>r.officer===officer.name))}>Download Officer PDF</button></section>
 <section className="table-section"><div className="table-head"><div><h2>{officer.name} — PS Detail</h2><p>{rows.length} PS shown</p></div><input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search PS / BLO / Supervisor"/></div><div className="table-wrap"><table><thead><tr><th>PS</th><th>BLO</th><th>Supervisor</th><th>Scheduled</th><th>Delivered</th><th>Pending</th><th>% Delivered</th><th>Docs Uploaded</th><th>% Docs</th><th>Hearing</th><th>Date(s)</th></tr></thead><tbody>{rows.map(r=><tr key={r.ps}><td><b>{r.ps}</b></td><td>{r.blo}</td><td>{r.supervisor}</td><td>{r.scheduled}</td><td>{r.delivered}</td><td>{r.pending}</td><td><span className={'pill '+(r.deliveredPct>=80?'good':r.deliveredPct>=60?'mid':'low')}>{pct(r.deliveredPct)}</span></td><td>{r.docs}</td><td>{pct(r.docsPct)}</td><td>{r.hearingStatus||'—'}</td><td>{r.dates||'—'}</td></tr>)}</tbody></table></div></section>
 <footer>Upload-only workflow: no manual selection or code changes are required for daily Excel updates.</footer>
 </main>
}