import { useState, useMemo, useRef, useEffect } from "react";
import * as XLSX from "xlsx";

// ─── Supabase ─────────────────────────────────────────────────────────────────
const SB_URL = "https://xljglqiifogyxefhszwa.supabase.co";
const SB_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InhsamdscWlpZm9neXhlZmhzendhIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODEwMTY2MTQsImV4cCI6MjA5NjU5MjYxNH0.asql85bUrgL5JuzqYoU0ZtizIWJ1yU6NYTt3yMUW5us";
const SB_H = { "Content-Type":"application/json","apikey":SB_KEY,"Authorization":`Bearer ${SB_KEY}`,"Prefer":"return=minimal" };
async function sbGet(t,f=""){const r=await fetch(`${SB_URL}/rest/v1/${t}?${f}`,{headers:SB_H});if(!r.ok)throw new Error(await r.text());return r.json();}
async function sbUpsert(t,d){const r=await fetch(`${SB_URL}/rest/v1/${t}`,{method:"POST",headers:{...SB_H,"Prefer":"resolution=merge-duplicates"},body:JSON.stringify(d)});if(!r.ok)throw new Error(await r.text());}
async function sbDelete(t,f){const r=await fetch(`${SB_URL}/rest/v1/${t}?${f}`,{method:"DELETE",headers:SB_H});if(!r.ok)throw new Error(await r.text());}

// ─── Constants ────────────────────────────────────────────────────────────────
const BASE_DAYS=["Mon","Tue","Wed","Thu","Fri"];
const WEEKEND_DAYS=["Sat","Sun"];
const ALL_DAYS=[...BASE_DAYS,...WEEKEND_DAYS];
const POSITIONS=["Welder","Fixer","Fitter","Semiskilled","Supervisor","Labourer","Manager","Driver"];
const COMPANIES=["Bright Matalwork","Dodi Metalwork","External"];
const DEFAULT_HOURS=9;
const PRESET_COLORS=["#3b82f6","#8b5cf6","#06b6d4","#10b981","#f59e0b","#ef4444","#f97316","#ec4899","#6366f1","#84cc16","#a78bfa","#14b8a6","#e11d48","#7c3aed","#0284c7","#d97706","#65a30d","#db2777"];

const CERTS=[
  {key:"cscs",label:"CSCS Card",hasExpiry:true},{key:"nvq2",label:"NVQ 2 Fenestration",hasExpiry:false},
  {key:"nvq3",label:"NVQ 3 Supervisor",hasExpiry:false},{key:"nvq4",label:"Level 4 NVQ",hasExpiry:false},
  {key:"schuco",label:"Schuco Skills Card",hasExpiry:true},{key:"healthSafety",label:"Health & Safety",hasExpiry:true},
  {key:"harness",label:"Harness & Leading Edge",hasExpiry:true},{key:"manualHandling",label:"Manual Handling",hasExpiry:true},
  {key:"ipaf3",label:"IPAF 3a/3b",hasExpiry:true},{key:"ipaf1b",label:"IPAF 1b",hasExpiry:true},
  {key:"ipafMast",label:"IPAF Mast Climber",hasExpiry:true},{key:"pasma",label:"PASMA",hasExpiry:true},
  {key:"abrasiveWheel",label:"Abrasive Wheel",hasExpiry:true},{key:"trafficMarshal",label:"Traffic Marshal",hasExpiry:true},
  {key:"firstAid",label:"First Aid",hasExpiry:true},{key:"fireSafety",label:"Fire Safety Marshall",hasExpiry:true},
  {key:"faceFit",label:"Face Fit Testing",hasExpiry:true},{key:"iosh",label:"IOSH Managing Safely",hasExpiry:false},
  {key:"smsts",label:"SMSTS Certificate",hasExpiry:true},{key:"sssts",label:"SSSTS Certificate",hasExpiry:true},
  {key:"asbestos",label:"Asbestos Awareness",hasExpiry:true},{key:"spiderCrane",label:"Spider Crane",hasExpiry:true},
  {key:"vacuumLifter",label:"Vacuum Lifter",hasExpiry:true},
];

// Built-in sites are now fully editable — stored in state, seeded from this default
const DEFAULT_BUILTIN_SITES=[
  {id:"b1",name:"003 - STF",color:"#f59e0b",clientId:null,builtin:true},
  {id:"b2",name:"0066 - UKTOP",color:"#3b82f6",clientId:null,builtin:true},
  {id:"b3",name:"JAUK - 42 Station Road",color:"#8b5cf6",clientId:null,builtin:true},
  {id:"b4",name:"JAUK - Pool Street",color:"#06b6d4",clientId:null,builtin:true},
  {id:"b5",name:"JAUK - Tower 42",color:"#10b981",clientId:null,builtin:true},
  {id:"b6",name:"BMW",color:"#ef4444",clientId:null,builtin:true},
  {id:"b7",name:"SB - Camden",color:"#f97316",clientId:null,builtin:true},
  {id:"b8",name:"DODI",color:"#ec4899",clientId:null,builtin:true},
  {id:"b9",name:"SS",color:"#6366f1",clientId:null,builtin:true},
  {id:"b10",name:"XX - OFF",color:"#6b7280",clientId:null,builtin:true},
  {id:"b11",name:"X - Holiday",color:"#84cc16",clientId:null,builtin:true},
  {id:"b12",name:"XX - Storage",color:"#a78bfa",clientId:null,builtin:true},
];

// ─── Helpers ──────────────────────────────────────────────────────────────────
function getSiteColor(site,allSites=[]){
  if(!site?.trim()) return "#374151";
  const c=site.trim();
  for(const s of allSites){if(c===s.name||c.toUpperCase().includes(s.name.toUpperCase())) return s.color;}
  let h=0; for(let i=0;i<c.length;i++) h=(h*31+c.charCodeAt(i))&0xffff;
  return `hsl(${[200,220,260,280,300,340,20,40,160,180][h%10]},60%,45%)`;
}
function isOff(s){if(!s) return true;const x=s.toLowerCase();return x.includes("off")||x.includes("holiday")||x.includes("storage")||!x.trim();}
function cSt(cert,w){
  const v=w.certs?.[cert.key];if(!v||!v.held) return "missing";
  if(!cert.hasExpiry||!v.expiry) return "valid";
  const d=(new Date(v.expiry)-new Date())/86400000;
  return d<0?"expired":d<30?"expiring":"valid";
}
function emptyCerts(){return Object.fromEntries(CERTS.map(c=>[c.key,{held:false,expiry:""}]));}
function emptyDays(){return Object.fromEntries(ALL_DAYS.map(d=>[d,""]));}
function emptyHrs(){return Object.fromEntries(ALL_DAYS.map(d=>[d,DEFAULT_HOURS]));}
function emptyOT(){return Object.fromEntries(ALL_DAYS.map(d=>[d,0]));}
function mkW(o={}){
  return {id:String(Date.now()+Math.random()),name:"",company:"",position:"",scope:"",
    days:emptyDays(),hoursPerDay:emptyHrs(),overtimeHours:emptyOT(),
    agreedRate:null,actualRate:null,taxRate:0,overtimeMultiplier:1.5,customOTRate:null,
    contact:"",email:"",dob:"",address:"",nino:"",utr:"",
    bankName:"",bankAccount:"",bankSort:"",nextOfKin:"",nextOfKinPhone:"",
    comments:"",certs:emptyCerts(),...o};
}
function calcPay(w,days,siteHours){
  const rate=w.agreedRate||0,tax=w.taxRate||0,otM=w.customOTRate||(w.overtimeMultiplier||1.5);
  let stdH=0,otH=0,gross=0;const bd={};
  days.forEach(d=>{
    const site=w.days[d];if(!site||isOff(site)) return;
    const sk=site.trim(),hrs=siteHours[sk]?.hours||w.hoursPerDay?.[d]||DEFAULT_HOURS,ot=w.overtimeHours?.[d]||0;
    const stdPay=hrs*rate,otPay=ot*rate*otM,g=stdPay+otPay;
    stdH+=hrs;otH+=ot;gross+=g;
    bd[d]={site:sk,hours:hrs,ot,stdPay,otPay,gross:g};
  });
  return {stdH,otH,gross,tax:gross*tax,net:gross-(gross*tax),bd};
}
function fmtDate(d){return d?new Date(d).toLocaleDateString("en-GB",{day:"2-digit",month:"short",year:"numeric"}):"—";}
function formatWeekLabel(d){return d.toLocaleDateString("en-GB",{day:"2-digit",month:"short",year:"numeric"});}
function addWeeks(label,n){
  try{const d=new Date(label);if(isNaN(d)){const nd=new Date();nd.setDate(nd.getDate()+n*7);return formatWeekLabel(nd);}
  d.setDate(d.getDate()+n*7);return formatWeekLabel(d);}catch(e){return label;}
}

// ─── Initial Data ─────────────────────────────────────────────────────────────
const INIT_W=[
  mkW({id:"w1",name:"Oleg Goraz",company:"Dodi Metalwork",position:"Welder",agreedRate:20,taxRate:0.20,contact:"07700000001",email:"oleg@example.com",days:{...emptyDays(),Mon:"003 - STF - Stair Formwork",Tue:"003 - STF - Stair Formwork",Wed:"003 - STF - Stair Formwork",Thu:"003 - STF - Stair Formwork",Fri:"003 - STF - Stair Formwork"},scope:"Welding Balustrade",certs:{...emptyCerts(),cscs:{held:true,expiry:"2026-12-01"},harness:{held:true,expiry:"2027-03-15"}}}),
  mkW({id:"w2",name:"Stefan Vrabiuta",company:"Bright Matalwork",position:"Fixer",agreedRate:18,taxRate:0.20,days:{...emptyDays(),Mon:"003 - STF - Stair Formwork",Tue:"003 - STF - Stair Formwork",Wed:"003 - STF - Stair Formwork",Thu:"003 - STF - Stair Formwork",Fri:"003 - STF - Stair Formwork"},scope:"Welding Balustrade",certs:{...emptyCerts(),cscs:{held:true,expiry:"2025-11-30"},nvq2:{held:true,expiry:""}}}),
  mkW({id:"w3",name:"Yonir Ordonez Molina",company:"Bright Matalwork",position:"Welder",agreedRate:18,taxRate:0,days:{...emptyDays(),Mon:"003 - STF - Stair Formwork",Tue:"003 - STF - Stair Formwork",Wed:"003 - STF - Stair Formwork",Thu:"003 - STF - Stair Formwork",Fri:"003 - STF - Stair Formwork"},scope:"Welding Balustrade"}),
  mkW({id:"w4",name:"Rodrigo Amigo Lopez",company:"Bright Matalwork",position:"Welder",agreedRate:25,taxRate:0.20,comments:"Supervisor",days:{...emptyDays(),Mon:"003 - STF - Stair Formwork",Tue:"003 - STF - Stair Formwork",Wed:"003 - STF - Stair Formwork",Thu:"003 - STF - Stair Formwork",Fri:"003 - STF - Stair Formwork"},scope:"Welding Balustrade",certs:{...emptyCerts(),cscs:{held:true,expiry:"2027-05-01"},nvq3:{held:true,expiry:""},smsts:{held:true,expiry:"2026-08-01"},firstAid:{held:true,expiry:"2026-06-15"}}}),
  mkW({id:"w5",name:"Adrian Bacescu",company:"Bright Matalwork",position:"Fitter",days:{...emptyDays(),Mon:"0066 - UKTOP - Fulham reach",Tue:"0066 - UKTOP - Fulham reach",Wed:"0066 - UKTOP - Fulham reach",Thu:"0066 - UKTOP - Fulham reach",Fri:"0066 - UKTOP - Fulham reach"},scope:"Mixt"}),
  mkW({id:"w6",name:"Vasile Cristinel Oprea",company:"Bright Matalwork",position:"Fitter",agreedRate:22,taxRate:0.20,days:{...emptyDays(),Mon:"0066 - UKTOP - Fulham reach",Tue:"0066 - UKTOP - Fulham reach",Wed:"0066 - UKTOP - Fulham reach",Thu:"0066 - UKTOP - Fulham reach",Fri:"0066 - UKTOP - Fulham reach"},scope:"Mixt",certs:{...emptyCerts(),cscs:{held:true,expiry:"2026-11-01"},ipaf3:{held:true,expiry:"2026-07-20"}}}),
  mkW({id:"w7",name:"Florin Badac",company:"Bright Matalwork",position:"Fitter",days:{...emptyDays(),Mon:"DODI",Tue:"DODI",Wed:"DODI",Thu:"DODI",Fri:"DODI"}}),
  mkW({id:"w8",name:"Sergiu Tugui",company:"Bright Matalwork",position:"Fitter",agreedRate:21,taxRate:0.20,comments:"Supervisor",days:{...emptyDays(),Mon:"JAUK - 42 Station Road",Tue:"JAUK - 42 Station Road",Wed:"JAUK - 42 Station Road",Thu:"JAUK - 42 Station Road",Fri:"JAUK - 42 Station Road"},scope:"Soffits",certs:{...emptyCerts(),cscs:{held:true,expiry:"2027-02-01"},nvq3:{held:true,expiry:""},firstAid:{held:true,expiry:"2027-01-01"}}}),
  mkW({id:"w9",name:"Vasile Gorbatii",company:"Bright Matalwork",position:"Semiskilled",days:{...emptyDays(),Mon:"JAUK - 42 Station Road",Tue:"JAUK - 42 Station Road",Wed:"JAUK - 42 Station Road",Thu:"JAUK - 42 Station Road",Fri:"JAUK - 42 Station Road"},scope:"Soffits"}),
  mkW({id:"w10",name:"Florin Stanciu",company:"Bright Matalwork",position:"Semiskilled",days:{...emptyDays(),Mon:"JAUK - 42 Station Road",Tue:"JAUK - 42 Station Road",Wed:"JAUK - 42 Station Road",Thu:"JAUK - 42 Station Road",Fri:"JAUK - 42 Station Road"},scope:"Soffits"}),
  mkW({id:"w11",name:"Davidel Nicolae",company:"Bright Matalwork",position:"Semiskilled",agreedRate:17,taxRate:0,days:{...emptyDays(),Mon:"JAUK - Pool Street",Tue:"JAUK - Pool Street",Wed:"JAUK - Pool Street",Thu:"JAUK - Pool Street",Fri:"JAUK - Pool Street"},scope:"Structural steel",certs:{...emptyCerts(),cscs:{held:true,expiry:"2026-06-25"}}}),
  mkW({id:"w12",name:"Luka Davitashvili",company:"Bright Matalwork",position:"Semiskilled",days:{...emptyDays(),Mon:"JAUK - Pool Street",Tue:"JAUK - Pool Street",Wed:"JAUK - Pool Street",Thu:"JAUK - Pool Street",Fri:"JAUK - Pool Street"},scope:"Structural steel"}),
  mkW({id:"w13",name:"Gurvinder Singh",company:"Bright Matalwork",position:"Semiskilled",agreedRate:16,taxRate:0,days:{...emptyDays(),Mon:"XX - Storage",Tue:"XX - Storage",Wed:"JAUK - Tower 42",Thu:"JAUK - Tower 42",Fri:"JAUK - Tower 42"}}),
  mkW({id:"w14",name:"Costel Clapa",company:"Bright Matalwork",position:"Fitter",agreedRate:21,taxRate:0.20,days:{...emptyDays(),Mon:"SS - Daniel House",Tue:"JAUK - Tower 42",Wed:"JAUK - Tower 42",Thu:"JAUK - Tower 42",Fri:"JAUK - Tower 42"},scope:"Balustrade",certs:{...emptyCerts(),cscs:{held:true,expiry:"2026-12-31"},ipaf3:{held:true,expiry:"2026-07-05"}}}),
  mkW({id:"w15",name:"Haroon Ahmed",company:"Bright Matalwork",position:"Fixer",days:{...emptyDays(),Mon:"XX - OFF",Tue:"JAUK - Tower 42",Wed:"JAUK - Tower 42",Thu:"JAUK - Tower 42",Fri:"JAUK - Tower 42"},scope:"Balustrade"}),
  mkW({id:"w16",name:"Florentin Firtat",company:"Bright Matalwork",position:"Fitter",days:{...emptyDays(),Mon:"X - Holiday",Tue:"X - Holiday",Wed:"X - Holiday",Thu:"X - Holiday",Fri:"X - Holiday"},scope:"Enjoy"}),
];
const INIT_CLIENTS=[
  {id:"c1",name:"JAUK Ltd",email:"info@jauk.com",phone:"02012345678",color:"#8b5cf6",notes:""},
  {id:"c2",name:"STF Projects",email:"info@stf.com",phone:"02087654321",color:"#f59e0b",notes:""}
];

// ─── Styles ───────────────────────────────────────────────────────────────────
const TH={padding:"9px 10px",textAlign:"left",fontSize:11,fontWeight:700,color:"#64748b",textTransform:"uppercase",letterSpacing:"0.07em",borderBottom:"1px solid #1e2535",background:"#0d1117",whiteSpace:"nowrap"};
const TD={padding:"6px 9px",borderBottom:"1px solid #1a2030",verticalAlign:"middle"};
const INP={width:"100%",background:"#0f1421",border:"1px solid #2d3555",borderRadius:6,padding:"7px 9px",color:"#e2e8f0",fontSize:13,outline:"none",boxSizing:"border-box"};
const LBL={display:"block",fontSize:11,color:"#64748b",marginBottom:3,fontWeight:700,textTransform:"uppercase",letterSpacing:"0.04em"};
const BP={padding:"8px 16px",background:"linear-gradient(135deg,#3b82f6,#6366f1)",border:"none",borderRadius:7,color:"#fff",cursor:"pointer",fontSize:13,fontWeight:700};
const BG={padding:"8px 16px",background:"linear-gradient(135deg,#059669,#10b981)",border:"none",borderRadius:7,color:"#fff",cursor:"pointer",fontSize:13,fontWeight:700};

// ─── Atoms ────────────────────────────────────────────────────────────────────
function Bdg({label,color}){return <span style={{display:"inline-block",padding:"2px 7px",borderRadius:4,fontSize:11,fontWeight:600,color:"#fff",background:color,whiteSpace:"nowrap",maxWidth:145,overflow:"hidden",textOverflow:"ellipsis"}} title={label}>{label||"—"}</span>;}
function CDot({status,label}){const c={valid:"#34d399",expiring:"#fbbf24",expired:"#f87171",missing:"#2d3555"}[status];return <span title={`${label}: ${status}`} style={{display:"inline-block",width:10,height:10,borderRadius:"50%",background:c,border:`1px solid ${c}`,margin:1}}/>;}
function Overlay({onClose,children,wide}){
  return <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.78)",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center",padding:12}} onClick={onClose}>
    <div style={{background:"#1a1f2e",border:"1px solid #2d3555",borderRadius:12,padding:24,width:"100%",maxWidth:wide?960:780,maxHeight:"93vh",overflowY:"auto",boxShadow:"0 24px 80px rgba(0,0,0,0.9)"}} onClick={e=>e.stopPropagation()}>
      {children}
    </div>
  </div>;
}
function MH({title,onClose}){return <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:18}}><h3 style={{margin:0,color:"#e2e8f0",fontSize:17,fontWeight:700}}>{title}</h3><button onClick={onClose} style={{background:"none",border:"none",color:"#94a3b8",cursor:"pointer",fontSize:24,lineHeight:1}}>×</button></div>;}
function FI({label,value,onChange,type="text",placeholder=""}){return <div style={{marginBottom:11}}><label style={LBL}>{label}</label><input type={type} value={value??""} onChange={e=>onChange(e.target.value)} placeholder={placeholder} style={INP}/></div>;}
function FSel({label,value,onChange,options}){return <div style={{marginBottom:11}}><label style={LBL}>{label}</label><select value={value??""} onChange={e=>onChange(e.target.value)} style={{...INP,cursor:"pointer"}}><option value="">— Select —</option>{options.map(o=>typeof o==="string"?<option key={o} value={o}>{o}</option>:<option key={o.value} value={o.value}>{o.label}</option>)}</select></div>;}
function Sec({title,color="#64748b",children}){return <div style={{background:"#0f1421",borderRadius:10,padding:14,marginBottom:14,border:"1px solid #1e2535"}}><div style={{fontSize:11,color,fontWeight:700,textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:12}}>{title}</div>{children}</div>;}
function TabBar({tabs,active,onChange}){return <div style={{display:"flex",gap:3,background:"#0d1117",borderRadius:8,padding:3,marginBottom:18}}>{tabs.map(([v,l])=><button key={v} onClick={()=>onChange(v)} style={{flex:1,padding:"6px 8px",background:active===v?"#1e3a5f":"transparent",border:active===v?"1px solid #3b82f6":"1px solid transparent",borderRadius:6,color:active===v?"#60a5fa":"#64748b",cursor:"pointer",fontSize:12,fontWeight:active===v?700:400}}>{l}</button>)}</div>;}

// ─── Inline Cell ──────────────────────────────────────────────────────────────
function InlineCell({value,workerId,day,allSiteNames,allSites,onUpdate}){
  const [editing,setEditing]=useState(false);
  const [val,setVal]=useState(value||"");
  const ref=useRef(null);
  const uid=`ic-${workerId}-${day}`;
  useEffect(()=>{setVal(value||"");},[value]);
  useEffect(()=>{if(editing&&ref.current){ref.current.focus();ref.current.select();}},[editing]);
  const commit=()=>{onUpdate(workerId,day,val.trim());setEditing(false);};
  const cancel=()=>{setVal(value||"");setEditing(false);};
  if(editing) return <div style={{minWidth:130}}>
    <input ref={ref} list={uid} value={val} onChange={e=>setVal(e.target.value)}
      onKeyDown={e=>{if(e.key==="Enter")commit();if(e.key==="Escape")cancel();}} onBlur={commit}
      style={{width:"100%",background:"#0d1117",border:`2px solid ${getSiteColor(val,allSites)||"#3b82f6"}`,borderRadius:6,padding:"5px 8px",color:"#e2e8f0",fontSize:12,outline:"none",boxSizing:"border-box"}}/>
    <datalist id={uid}>{allSiteNames.map(s=><option key={s} value={s}/>)}</datalist>
  </div>;
  return <div onClick={()=>setEditing(true)} title="Click to edit" style={{cursor:"text",minWidth:110,padding:"3px 4px",borderRadius:5,border:"1px solid transparent",transition:"border-color 0.15s"}}
    onMouseEnter={e=>e.currentTarget.style.borderColor="#2d3555"} onMouseLeave={e=>e.currentTarget.style.borderColor="transparent"}>
    {value?<Bdg label={value.trim()} color={getSiteColor(value,allSites)}/>:<span style={{color:"#374151",fontSize:11}}>— click —</span>}
  </div>;
}

// ─── Worker Profile PDF ───────────────────────────────────────────────────────
function exportWorkerProfile(w,allSites,weekLabel){
  const heldCerts=CERTS.filter(c=>w.certs?.[c.key]?.held);
  const SC={valid:"#22c55e",expiring:"#f59e0b",expired:"#ef4444",missing:"#6b7280"};
  const html=`<!DOCTYPE html><html><head><meta charset="utf-8"/><title>Worker Profile — ${w.name}</title>
<style>*{margin:0;padding:0;box-sizing:border-box;}body{background:#0d1117;color:#e2e8f0;font-family:'Segoe UI',Arial,sans-serif;font-size:12px;padding:24px;}
.hdr{display:flex;align-items:center;justify-content:space-between;margin-bottom:22px;padding-bottom:14px;border-bottom:2px solid #1e2535;}
.avatar{width:56px;height:56px;background:linear-gradient(135deg,#3b82f6,#6366f1);border-radius:14px;display:flex;align-items:center;justify-content:center;font-size:22px;color:#fff;margin-right:14px;flex-shrink:0;}
.name{font-size:22px;font-weight:800;color:#f1f5f9;letter-spacing:-0.02em;}
.pos{font-size:13px;color:#64748b;margin-top:3px;}
.company{font-size:12px;color:#60a5fa;margin-top:2px;}
.grid{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:20px;}
.card{background:#1a1f2e;border:1px solid #2d3555;border-radius:10px;padding:16px;}
.card-title{font-size:10px;color:#64748b;font-weight:700;text-transform:uppercase;letter-spacing:0.07em;margin-bottom:12px;}
.field{margin-bottom:9px;display:flex;align-items:flex-start;gap:8px;}
.field-label{font-size:10px;color:#64748b;font-weight:600;min-width:100px;flex-shrink:0;text-transform:uppercase;margin-top:1px;}
.field-value{font-size:12px;color:#e2e8f0;font-weight:500;}
.week-table{width:100%;border-collapse:collapse;margin-bottom:20px;}
.week-table th{padding:8px 10px;text-align:left;font-size:10px;font-weight:700;color:#64748b;text-transform:uppercase;border-bottom:2px solid #1e2535;background:#0d1117;}
.week-table td{padding:7px 10px;border-bottom:1px solid #1a2030;font-size:11px;}
.week-table tr:nth-child(even) td{background:#111827;}
.site-badge{display:inline-block;padding:2px 8px;border-radius:4px;font-size:11px;font-weight:600;color:#fff;}
.cert-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin-bottom:20px;}
.cert-item{background:#1a1f2e;border-radius:8px;padding:9px 11px;border-left:3px solid #2d3555;}
.cert-name{font-size:11px;color:#e2e8f0;font-weight:600;margin-bottom:3px;}
.cert-exp{font-size:10px;color:#64748b;}
.cert-status{font-size:9px;font-weight:700;text-transform:uppercase;}
.scope-badge{display:inline-block;background:#1e3a5f;border:1px solid #3b82f6;border-radius:6px;padding:4px 10px;font-size:12px;color:#60a5fa;font-weight:600;}
.ft{margin-top:20px;padding-top:12px;border-top:1px solid #1e2535;display:flex;justify-content:space-between;font-size:9px;color:#374151;}
@media print{body{padding:12px;}@page{margin:8mm;size:A4;}}</style></head><body>
<div class="hdr">
  <div style="display:flex;align-items:center">
    <div class="avatar">👷</div>
    <div>
      <div class="name">${w.name||"—"}</div>
      <div class="pos">${w.position||"—"} ${w.comments?`· <span style="color:#fbbf24">⚑ ${w.comments}</span>`:""}
      </div>
      <div class="company">${w.company||"—"}</div>
    </div>
  </div>
  <div style="text-align:right">
    <div style="font-size:10px;color:#64748b">Week Commencing</div>
    <div style="font-size:15px;font-weight:700;color:#60a5fa">${weekLabel}</div>
    <div style="font-size:10px;color:#374151;margin-top:4px">Generated: ${new Date().toLocaleDateString("en-GB")}</div>
  </div>
</div>

<div class="grid">
  <div class="card">
    <div class="card-title">Contact Details</div>
    ${w.contact?`<div class="field"><span class="field-label">Phone</span><span class="field-value">${w.contact}</span></div>`:""}
    ${w.email?`<div class="field"><span class="field-label">Email</span><span class="field-value">${w.email}</span></div>`:""}
    ${w.dob?`<div class="field"><span class="field-label">Date of Birth</span><span class="field-value">${fmtDate(w.dob)}</span></div>`:""}
    ${w.address?`<div class="field"><span class="field-label">Address</span><span class="field-value">${w.address}</span></div>`:""}
    ${w.nextOfKin?`<div class="field"><span class="field-label">Next of Kin</span><span class="field-value">${w.nextOfKin}${w.nextOfKinPhone?` · ${w.nextOfKinPhone}`:""}</span></div>`:""}
  </div>
  <div class="card">
    <div class="card-title">Role & Scope</div>
    <div class="field"><span class="field-label">Position</span><span class="field-value">${w.position||"—"}</span></div>
    <div class="field"><span class="field-label">Company</span><span class="field-value">${w.company||"—"}</span></div>
    ${w.scope?`<div class="field"><span class="field-label">Scope</span><span class="field-value"><span class="scope-badge">${w.scope}</span></span></div>`:""}
  </div>
</div>

<div class="card" style="margin-bottom:20px">
  <div class="card-title">Weekly Site Allocation — WC ${weekLabel}</div>
  <table class="week-table">
    <thead><tr><th>Day</th><th>Site Allocated</th></tr></thead>
    <tbody>
      ${ALL_DAYS.map(d=>{const site=w.days[d];const col=site?getSiteColor(site,allSites):"#374151";
        return `<tr><td style="font-weight:700;color:#94a3b8">${d}</td><td>${site?`<span class="site-badge" style="background:${col}22;color:${col};border:1px solid ${col}44">${site}</span>`:`<span style="color:#374151;font-style:italic">—</span>`}</td></tr>`;
      }).join("")}
    </tbody>
  </table>
</div>

${heldCerts.length>0?`
<div style="margin-bottom:20px">
  <div style="font-size:11px;color:#64748b;font-weight:700;text-transform:uppercase;letter-spacing:0.07em;margin-bottom:10px">Certifications & Training (${heldCerts.length} held)</div>
  <div class="cert-grid">
    ${heldCerts.map(cert=>{const status=cSt(cert,w);const sc=SC[status];const val=w.certs[cert.key];
      return `<div class="cert-item" style="border-left-color:${sc}">
        <div class="cert-name">${cert.label}</div>
        ${cert.hasExpiry&&val.expiry?`<div class="cert-exp">Expiry: ${fmtDate(val.expiry)}</div>`:""}
        <div class="cert-status" style="color:${sc};margin-top:4px">${status.toUpperCase()}</div>
      </div>`;
    }).join("")}
  </div>
</div>`:""}

<div class="ft">
  <span>Worker Profile — ${w.name}</span>
  <span>Bright Metalwork Ltd — Confidential</span>
  <span>Generated ${new Date().toLocaleDateString("en-GB")}</span>
</div>
<script>window.onload=function(){window.print();}</script>
</body></html>`;
  const b=new Blob([html],{type:"text/html"});
  const u=URL.createObjectURL(b);
  const win=window.open(u,"_blank","width=900,height=800");
  if(!win){const a=document.createElement("a");a.href=u;a.download=`Worker_Profile_${w.name.replace(/\s+/g,"_")}.html`;a.click();}
  setTimeout(()=>URL.revokeObjectURL(u),5000);
}

// ─── Payslip PDF ──────────────────────────────────────────────────────────────
function exportPayslip(w,activeDays,weekLabel,siteHours){
  const rate=w.agreedRate||0;
  const tax=w.taxRate||0;
  const otM=w.customOTRate||(w.overtimeMultiplier||1.5);
  const {stdH,otH,gross,tax:taxAmt,net,bd}=calcPay(w,activeDays,siteHours);
  const stdPay=Object.values(bd).reduce((a,r)=>a+r.stdPay,0);
  const otPay=Object.values(bd).reduce((a,r)=>a+r.otPay,0);
  const taxPct=Math.round(tax*100);

  const html=`<!DOCTYPE html><html><head><meta charset="utf-8"/><title>Payslip — ${w.name} — WC ${weekLabel}</title>
<style>*{margin:0;padding:0;box-sizing:border-box;}body{background:#0d1117;color:#e2e8f0;font-family:'Segoe UI',Arial,sans-serif;font-size:12px;padding:24px;}
.hdr{display:flex;align-items:center;justify-content:space-between;margin-bottom:20px;padding-bottom:14px;border-bottom:2px solid #1e2535;}
.logo{width:44px;height:44px;background:linear-gradient(135deg,#3b82f6,#6366f1);border-radius:11px;display:flex;align-items:center;justify-content:center;font-size:20px;color:#fff;margin-right:12px;flex-shrink:0;}
.title{font-size:18px;font-weight:800;color:#f1f5f9;}
.sub{font-size:11px;color:#64748b;margin-top:2px;}
.worker-box{background:#1a1f2e;border:1px solid #2d3555;border-radius:10px;padding:14px;margin-bottom:18px;display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;}
.wf-label{font-size:9px;color:#64748b;font-weight:700;text-transform:uppercase;letter-spacing:0.06em;}
.wf-value{font-size:13px;font-weight:700;color:#e2e8f0;margin-top:2px;}
table{width:100%;border-collapse:collapse;margin-bottom:16px;}
th{padding:8px 10px;text-align:left;font-size:10px;font-weight:700;color:#64748b;text-transform:uppercase;border-bottom:2px solid #1e2535;background:#0d1117;}
td{padding:8px 10px;border-bottom:1px solid #1a2030;font-size:12px;}
tr:nth-child(even) td{background:#111827;}tr:nth-child(odd) td{background:#0f1421;}
.site-badge{display:inline-block;padding:2px 8px;border-radius:4px;font-size:11px;font-weight:600;color:#fff;}
.summary{background:#1a1f2e;border:1px solid #2d3555;border-radius:10px;padding:16px;margin-bottom:18px;}
.s-row{display:flex;justify-content:space-between;align-items:center;padding:8px 0;border-bottom:1px solid #1e2535;}
.s-row:last-child{border-bottom:none;}
.s-label{font-size:12px;color:#94a3b8;}
.s-value{font-size:14px;font-weight:700;}
.s-row.total{background:#0d2218;border-radius:8px;padding:12px;margin-top:8px;border:1px solid #065f46;}
.net-box{background:linear-gradient(135deg,#0d2218,#1a3a28);border:2px solid #10b981;border-radius:12px;padding:20px;text-align:center;margin-bottom:18px;}
.net-label{font-size:11px;color:#34d399;font-weight:700;text-transform:uppercase;letter-spacing:0.08em;margin-bottom:6px;}
.net-amount{font-size:36px;font-weight:900;color:#34d399;letter-spacing:-0.02em;}
.net-sub{font-size:11px;color:#64748b;margin-top:4px;}
.notice{background:#1a1f2e;border:1px solid #2d3555;border-radius:8px;padding:12px;font-size:10px;color:#64748b;line-height:1.6;}
.ft{margin-top:16px;padding-top:10px;border-top:1px solid #1e2535;display:flex;justify-content:space-between;font-size:9px;color:#374151;}
@media print{body{padding:12px;}@page{margin:8mm;size:A4;}}</style></head><body>

<div class="hdr">
  <div style="display:flex;align-items:center">
    <div class="logo">🏗</div>
    <div>
      <div class="title">Payslip</div>
      <div class="sub">Bright Metalwork Ltd · Week Commencing ${weekLabel}</div>
    </div>
  </div>
  <div style="text-align:right">
    <div style="font-size:11px;font-weight:700;color:#60a5fa">PAY PERIOD</div>
    <div style="font-size:14px;font-weight:800;color:#f1f5f9">${weekLabel}</div>
    <div style="font-size:9px;color:#374151;margin-top:4px">Issued: ${new Date().toLocaleDateString("en-GB")}</div>
  </div>
</div>

<div class="worker-box">
  <div><div class="wf-label">Employee</div><div class="wf-value">${w.name||"—"}</div></div>
  <div><div class="wf-label">Position</div><div class="wf-value">${w.position||"—"}</div></div>
  <div><div class="wf-label">Company</div><div class="wf-value">${w.company||"—"}</div></div>
  <div><div class="wf-label">Hourly Rate</div><div class="wf-value" style="color:#34d399">${rate?`£${rate}/hr`:"Not set"}</div></div>
  <div><div class="wf-label">OT Rate</div><div class="wf-value" style="color:#fbbf24">${w.customOTRate?`£${w.customOTRate}/hr (custom)`:`×${otM} standard`}</div></div>
  <div><div class="wf-label">Tax Rate</div><div class="wf-value" style="color:${taxPct===30?"#f87171":taxPct===20?"#fbbf24":"#34d399"}">${taxPct}%</div></div>
</div>

<div style="font-size:11px;color:#64748b;font-weight:700;text-transform:uppercase;letter-spacing:0.07em;margin-bottom:10px">Daily Breakdown</div>
<table>
  <thead><tr><th>Day</th><th>Site</th><th>Std Hours</th><th>Std Pay</th><th>OT Hours</th><th>OT Pay (×${otM})</th><th>Day Total</th></tr></thead>
  <tbody>
    ${activeDays.map(d=>{const b=bd[d];const site=w.days[d];
      if(!b&&(!site||isOff(site))) return `<tr><td style="color:#94a3b8;font-weight:700">${d}</td><td style="color:#374151;font-style:italic">${site||"—"}</td><td colspan="5" style="color:#374151;text-align:center;font-style:italic">Not working</td></tr>`;
      if(!b) return `<tr><td style="color:#94a3b8;font-weight:700">${d}</td><td style="color:#374151;font-style:italic">${site||"—"}</td><td colspan="5" style="color:#374151;text-align:center;font-style:italic">No rate set</td></tr>`;
      return `<tr>
        <td style="font-weight:700;color:#94a3b8">${d}</td>
        <td><span class="site-badge" style="background:#1e3a5f;color:#60a5fa;border:1px solid #3b82f6">${b.site}</span></td>
        <td style="color:#60a5fa;font-weight:600">${b.hours}h</td>
        <td style="color:#e2e8f0">£${b.stdPay.toFixed(2)}</td>
        <td style="color:${b.ot>0?"#fbbf24":"#374151"}">${b.ot>0?b.ot+"h":"—"}</td>
        <td style="color:${b.ot>0?"#fbbf24":"#374151"}">${b.ot>0?"£"+b.otPay.toFixed(2):"—"}</td>
        <td style="color:#34d399;font-weight:700">£${b.gross.toFixed(2)}</td>
      </tr>`;
    }).join("")}
    <tr style="background:#0d1117;border-top:2px solid #2d3555">
      <td colspan="2" style="font-weight:700;color:#94a3b8;font-size:11px">TOTALS</td>
      <td style="color:#60a5fa;font-weight:800">${stdH}h</td>
      <td style="color:#e2e8f0;font-weight:700">£${stdPay.toFixed(2)}</td>
      <td style="color:#fbbf24;font-weight:800">${otH>0?otH+"h":"—"}</td>
      <td style="color:#fbbf24;font-weight:700">${otH>0?"£"+otPay.toFixed(2):"—"}</td>
      <td style="color:#34d399;font-weight:800">£${gross.toFixed(2)}</td>
    </tr>
  </tbody>
</table>

<div class="summary">
  <div style="font-size:11px;color:#64748b;font-weight:700;text-transform:uppercase;letter-spacing:0.07em;margin-bottom:10px">Pay Summary</div>
  <div class="s-row"><span class="s-label">Standard Pay (${stdH}h × £${rate}/hr)</span><span class="s-value" style="color:#e2e8f0">£${stdPay.toFixed(2)}</span></div>
  ${otH>0?`<div class="s-row"><span class="s-label">Overtime Pay (${otH}h × £${rate} × ${otM})</span><span class="s-value" style="color:#fbbf24">£${otPay.toFixed(2)}</span></div>`:""}
  <div class="s-row"><span class="s-label" style="font-weight:700;color:#e2e8f0">Gross Pay</span><span class="s-value" style="color:#34d399;font-size:16px">£${gross.toFixed(2)}</span></div>
  <div class="s-row"><span class="s-label">Tax Deduction (${taxPct}%)</span><span class="s-value" style="color:#f87171">-£${taxAmt.toFixed(2)}</span></div>
</div>

<div class="net-box">
  <div class="net-label">💷 Net Pay to Account</div>
  <div class="net-amount">£${net.toFixed(2)}</div>
  <div class="net-sub">After ${taxPct}% tax deduction · Week Commencing ${weekLabel}</div>
</div>

<div class="notice">
  <strong style="color:#94a3b8">PAYSLIP NOTICE:</strong> This payslip is for internal reference only. Standard hours calculated at £${rate}/hr. 
  Overtime calculated at ×${otM} (£${(rate*otM).toFixed(2)}/hr). Tax deducted at ${taxPct}% flat rate. 
  ${w.nino?`NI Number on file.`:""} Please retain for your records.
</div>

<div class="ft">
  <span>Payslip — ${w.name} — WC ${weekLabel}</span>
  <span>Bright Metalwork Ltd</span>
  <span>Net Pay: £${net.toFixed(2)}</span>
</div>
<script>window.onload=function(){window.print();}</script>
</body></html>`;

  const b=new Blob([html],{type:"text/html"});
  const u=URL.createObjectURL(b);
  const win=window.open(u,"_blank","width=900,height=800");
  if(!win){const a=document.createElement("a");a.href=u;a.download=`Payslip_${w.name.replace(/\s+/g,"_")}_${weekLabel.replace(/\s+/g,"_")}.html`;a.click();}
  setTimeout(()=>URL.revokeObjectURL(u),5000);
}

// ─── Schedule PDF ─────────────────────────────────────────────────────────────
function exportSchedulePDF(workers,activeDays,weekLabel,allSites){
  const getSC=s=>getSiteColor(s,allSites);
  const usedSites=[...new Set(workers.flatMap(w=>activeDays.map(d=>w.days[d]||"")).filter(s=>s&&!isOff(s)))];
  const html=`<!DOCTYPE html><html><head><meta charset="utf-8"/><title>Labour Schedule WC ${weekLabel}</title>
<style>*{margin:0;padding:0;box-sizing:border-box;}body{background:#0d1117;color:#e2e8f0;font-family:'Segoe UI',Arial,sans-serif;font-size:11px;padding:20px;}
.hdr{display:flex;align-items:center;justify-content:space-between;margin-bottom:18px;padding-bottom:12px;border-bottom:2px solid #1e2535;}
.icon{width:38px;height:38px;background:linear-gradient(135deg,#3b82f6,#6366f1);border-radius:9px;display:flex;align-items:center;justify-content:center;font-size:18px;color:#fff;margin-right:10px;}
.title{font-size:19px;font-weight:800;color:#f1f5f9;letter-spacing:-0.02em;}
.stats{display:flex;gap:10px;margin-bottom:14px;flex-wrap:wrap;}
.stat{background:#1a1f2e;border:1px solid #1e2535;border-radius:8px;padding:7px 13px;}
.sl{font-size:9px;color:#64748b;font-weight:700;text-transform:uppercase;}
.sv{font-size:15px;font-weight:800;margin-top:2px;}
table{width:100%;border-collapse:collapse;margin-bottom:16px;}
th{padding:7px 9px;text-align:left;font-size:10px;font-weight:700;color:#64748b;text-transform:uppercase;letter-spacing:0.07em;border-bottom:2px solid #1e2535;background:#0d1117;white-space:nowrap;}
td{padding:6px 8px;border-bottom:1px solid #1a2030;vertical-align:middle;}
tr:nth-child(even) td{background:#111827;}tr:nth-child(odd) td{background:#0f1421;}
.wn{font-weight:700;color:#f1f5f9;font-size:12px;}.wp{font-size:10px;color:#64748b;margin-top:1px;}.wc{font-size:10px;color:#fbbf24;}
.sb{display:inline-block;padding:2px 7px;border-radius:4px;font-size:10px;font-weight:700;white-space:nowrap;}
.leg{margin-top:14px;border-top:1px solid #1e2535;padding-top:10px;}
.li{display:inline-flex;align-items:center;padding:2px 7px;border-radius:20px;background:#111827;font-size:10px;color:#94a3b8;margin:2px;}
.dot{width:7px;height:7px;border-radius:50%;display:inline-block;margin-right:4px;}
.ft{margin-top:14px;padding-top:10px;border-top:1px solid #1e2535;display:flex;justify-content:space-between;font-size:9px;color:#374151;}
@media print{body{padding:8px;}@page{margin:8mm;size:A3 landscape;}}</style></head><body>
<div class="hdr"><div style="display:flex;align-items:center"><div class="icon">🏗</div><div><div class="title">Labour Schedule</div><div style="font-size:11px;color:#64748b">Week Commencing: <strong style="color:#60a5fa">${weekLabel}</strong></div></div></div>
<div style="text-align:right;font-size:10px;color:#64748b">Generated: ${new Date().toLocaleDateString("en-GB")}<br/><span style="color:#374151;font-size:9px">Bright Metalwork — Confidential</span></div></div>
<div class="stats">
<div class="stat"><div class="sl">Total Operatives</div><div class="sv" style="color:#60a5fa">${workers.length}</div></div>
<div class="stat"><div class="sl">Active Sites</div><div class="sv" style="color:#34d399">${usedSites.length}</div></div>
<div class="stat"><div class="sl">On Holiday</div><div class="sv" style="color:#fbbf24">${workers.filter(w=>activeDays.some(d=>w.days[d]?.includes("Holiday"))).length}</div></div>
<div class="stat"><div class="sl">Off</div><div class="sv" style="color:#94a3b8">${workers.filter(w=>activeDays.every(d=>isOff(w.days[d]))).length}</div></div>
</div>
<table><thead><tr><th style="min-width:150px">Operative</th><th>Company</th><th>Position</th>
${activeDays.map(d=>`<th style="min-width:120px;color:${WEEKEND_DAYS.includes(d)?"#fbbf24":"#64748b"}">${d}${WEEKEND_DAYS.includes(d)?" 🟡":""}</th>`).join("")}</tr></thead>
<tbody>${workers.map(w=>`<tr><td><div class="wn">${w.name||"—"}</div><div class="wp">${w.position||""}</div>${w.comments?`<div class="wc">⚑ ${w.comments}</div>`:""}</td><td style="color:#94a3b8;font-size:10px">${w.company||"—"}</td><td style="color:#94a3b8;font-size:10px">${w.position||"—"}</td>
${activeDays.map(d=>{const s=w.days[d];const c=getSC(s);return s&&s.trim()?`<td><span class="sb" style="background:${c}22;color:${c};border:1px solid ${c}44">${s.trim()}</span></td>`:`<td style="color:#374151;font-style:italic;font-size:10px">—</td>`;}).join("")}</tr>`).join("")}
</tbody></table>
<div class="leg"><div style="font-size:10px;color:#374151;font-weight:700;text-transform:uppercase;margin-bottom:7px">Site Legend</div><div>${usedSites.map(s=>`<span class="li" style="border:1px solid ${getSC(s)}44"><span class="dot" style="background:${getSC(s)}"></span>${s}</span>`).join("")}</div></div>
<div class="ft"><span>Labour Schedule — WC ${weekLabel}</span><span>Total: ${workers.length} operatives</span><span>Bright Metalwork Ltd — Confidential</span></div>
<script>window.onload=function(){window.print();}</script></body></html>`;
  const b=new Blob([html],{type:"text/html"});
  const u=URL.createObjectURL(b);
  const win=window.open(u,"_blank","width=1200,height=800");
  if(!win){const a=document.createElement("a");a.href=u;a.download=`Schedule_WC_${weekLabel.replace(/\s+/g,"_")}.html`;a.click();}
  setTimeout(()=>URL.revokeObjectURL(u),5000);
}

// ─── Excel Export ─────────────────────────────────────────────────────────────
function doExcel(workers,weekLabel,activeDays,siteHours,clients,allSites){
  const wb=XLSX.utils.book_new();
  const ws1=XLSX.utils.aoa_to_sheet([[`SCHEDULE — WC: ${weekLabel}`],["Name","Company","Position",...activeDays,"Scope","Rate","OT","Tax%","Comments"],...workers.map(w=>[w.name,w.company,w.position,...activeDays.map(d=>w.days[d]||""),w.scope,w.agreedRate||"",w.customOTRate?`£${w.customOTRate}/hr`:`×${w.overtimeMultiplier||1.5}`,Math.round((w.taxRate||0)*100)+"%",w.comments])]);
  XLSX.utils.book_append_sheet(wb,ws1,"Schedule");
  const pr=workers.map(w=>({w,...calcPay(w,activeDays,siteHours)}));
  const ws2=XLSX.utils.aoa_to_sheet([[`PAYROLL — WC: ${weekLabel}`],["Name","Company","Rate","OT Rate","Tax%","Std Hrs","OT Hrs","Std Pay","OT Pay","Gross","Tax","Net"],...pr.map(({w,stdH,otH,gross,tax,net,bd})=>{const stdP=Object.values(bd).reduce((a,r)=>a+r.stdPay,0);const otP=Object.values(bd).reduce((a,r)=>a+r.otPay,0);return [w.name,w.company,w.agreedRate||"",w.customOTRate?`£${w.customOTRate}/hr`:`×${w.overtimeMultiplier||1.5}`,Math.round((w.taxRate||0)*100)+"%",stdH,otH,+stdP.toFixed(2),+otP.toFixed(2),+gross.toFixed(2),+tax.toFixed(2),+net.toFixed(2)];}),["","","","TOTALS","",pr.reduce((a,r)=>a+r.stdH,0),pr.reduce((a,r)=>a+r.otH,0),"","",+pr.reduce((a,r)=>a+r.gross,0).toFixed(2),+pr.reduce((a,r)=>a+r.tax,0).toFixed(2),+pr.reduce((a,r)=>a+r.net,0).toFixed(2)]]);
  XLSX.utils.book_append_sheet(wb,ws2,"Payroll");
  const ws3=XLSX.utils.aoa_to_sheet([[`TRAINING MATRIX — WC: ${weekLabel}`],["Name","Company","Position","DOB","Contact","Email","NINO",...CERTS.map(c=>c.label),...CERTS.filter(c=>c.hasExpiry).map(c=>`${c.label} Expiry`)],...workers.map(w=>[w.name,w.company,w.position,w.dob||"",w.contact||"",w.email||"",w.nino||"",...CERTS.map(c=>{const s=cSt(c,w);return w.certs?.[c.key]?.held?s.toUpperCase():""}),...CERTS.filter(c=>c.hasExpiry).map(c=>w.certs?.[c.key]?.expiry||"")])]);
  XLSX.utils.book_append_sheet(wb,ws3,"Training Matrix");
  const ws4=XLSX.utils.aoa_to_sheet([[`WORKER DIRECTORY — WC: ${weekLabel}`],["Name","Company","Position","DOB","Contact","Email","Address","NINO","UTR","Bank","Account","Sort","NOK","NOK Phone"],...workers.map(w=>[w.name,w.company,w.position,w.dob||"",w.contact||"",w.email||"",w.address||"",w.nino||"",w.utr||"",w.bankName||"",w.bankAccount||"",w.bankSort||"",w.nextOfKin||"",w.nextOfKinPhone||""])]);
  XLSX.utils.book_append_sheet(wb,ws4,"Worker Directory");
  XLSX.writeFile(wb,`LabourSchedule_WC_${weekLabel.replace(/\s+/g,"_")}.xlsx`);
}

// ─── Manage Sites Modal ───────────────────────────────────────────────────────
function SitesModal({allSites,clients,onSave,onClose}){
  const [sites,setSites]=useState(allSites.map(s=>({...s})));
  const [nn,setNn]=useState(""),[nc,setNc]=useState(PRESET_COLORS[0]),[ncl,setNcl]=useState("");
  const add=()=>{const n=nn.trim();if(!n||sites.find(s=>s.name.toLowerCase()===n.toLowerCase()))return;setSites(s=>[...s,{id:"s"+Date.now(),name:n,color:nc,clientId:ncl||null,builtin:false}]);setNn("");};
  const rm=id=>{if(window.confirm("Delete this site?"))setSites(s=>s.filter(x=>x.id!==id));};
  const up=(id,k,v)=>setSites(s=>s.map(x=>x.id===id?{...x,[k]:v}:x));
  const builtins=sites.filter(s=>s.builtin);
  const custom=sites.filter(s=>!s.builtin);
  return <Overlay onClose={onClose} wide>
    <MH title="🏗 Manage Sites" onClose={onClose}/>
    <Sec title="Add New Site">
      <div style={{display:"flex",gap:10,alignItems:"flex-end",flexWrap:"wrap"}}>
        <div style={{flex:2,minWidth:160}}><label style={LBL}>Site Name</label><input value={nn} onChange={e=>setNn(e.target.value)} onKeyDown={e=>e.key==="Enter"&&add()} placeholder="e.g. JAUK - New Road" style={INP}/></div>
        <div style={{flex:1,minWidth:130}}><label style={LBL}>Client</label><select value={ncl} onChange={e=>setNcl(e.target.value)} style={{...INP,cursor:"pointer"}}><option value="">No client</option>{clients.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}</select></div>
        <div><label style={LBL}>Colour</label><div style={{display:"flex",gap:3,flexWrap:"wrap",width:120}}>{PRESET_COLORS.map(c=><div key={c} onClick={()=>setNc(c)} style={{width:18,height:18,borderRadius:3,background:c,cursor:"pointer",border:nc===c?"3px solid #fff":"2px solid transparent",boxSizing:"border-box"}}/>)}</div></div>
        <button onClick={add} style={{...BP,whiteSpace:"nowrap"}}>+ Add</button>
      </div>
      {nn&&<div style={{marginTop:8}}><span style={{fontSize:12,color:"#64748b",marginRight:8}}>Preview:</span><span style={{display:"inline-block",padding:"2px 8px",borderRadius:4,fontSize:12,fontWeight:600,color:"#fff",background:nc}}>{nn}</span></div>}
    </Sec>

    {/* Built-in sites — now fully editable */}
    <div style={{marginBottom:18}}>
      <div style={{fontSize:11,color:"#fbbf24",fontWeight:700,textTransform:"uppercase",marginBottom:9}}>Built-in Sites — Now Editable ✏️</div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
        {builtins.map(s=><div key={s.id} style={{padding:"10px 12px",background:"#0f1421",borderRadius:8,border:`1px solid ${s.color}55`}}>
          <div style={{display:"flex",alignItems:"center",gap:7,marginBottom:7}}>
            <input value={s.name} onChange={e=>up(s.id,"name",e.target.value)} style={{flex:1,background:"#1a1f2e",border:`1px solid ${s.color}`,borderRadius:5,padding:"5px 8px",color:"#e2e8f0",fontSize:13,fontWeight:600,outline:"none"}}/>
          </div>
          <div style={{display:"flex",gap:3,flexWrap:"wrap",marginBottom:7}}>{PRESET_COLORS.map(c=><div key={c} onClick={()=>up(s.id,"color",c)} style={{width:16,height:16,borderRadius:3,background:c,cursor:"pointer",border:s.color===c?"3px solid #fff":"1px solid transparent",boxSizing:"border-box"}}/>)}</div>
          <select value={s.clientId||""} onChange={e=>up(s.id,"clientId",e.target.value||null)} style={{...INP,fontSize:12,padding:"4px 7px",cursor:"pointer"}}><option value="">No client</option>{clients.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}</select>
        </div>)}
      </div>
    </div>

    {custom.length>0&&<div>
      <div style={{fontSize:11,color:"#64748b",fontWeight:700,textTransform:"uppercase",marginBottom:9}}>Custom Sites ({custom.length})</div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
        {custom.map(s=><div key={s.id} style={{padding:"10px 12px",background:"#0f1421",borderRadius:8,border:`1px solid ${s.color}55`}}>
          <div style={{display:"flex",alignItems:"center",gap:7,marginBottom:7}}>
            <input value={s.name} onChange={e=>up(s.id,"name",e.target.value)} style={{flex:1,background:"#1a1f2e",border:`1px solid ${s.color}`,borderRadius:5,padding:"5px 8px",color:"#e2e8f0",fontSize:13,fontWeight:600,outline:"none"}}/>
            <button onClick={()=>rm(s.id)} style={{background:"#2d1515",border:"1px solid #ef4444",borderRadius:5,color:"#f87171",cursor:"pointer",fontSize:11,padding:"4px 8px",fontWeight:700}}>Delete</button>
          </div>
          <div style={{display:"flex",gap:3,flexWrap:"wrap",marginBottom:7}}>{PRESET_COLORS.map(c=><div key={c} onClick={()=>up(s.id,"color",c)} style={{width:16,height:16,borderRadius:3,background:c,cursor:"pointer",border:s.color===c?"3px solid #fff":"1px solid transparent",boxSizing:"border-box"}}/>)}</div>
          <select value={s.clientId||""} onChange={e=>up(s.id,"clientId",e.target.value||null)} style={{...INP,fontSize:12,padding:"4px 7px",cursor:"pointer"}}><option value="">No client</option>{clients.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}</select>
        </div>)}
      </div>
    </div>}

    <div style={{display:"flex",gap:10,justifyContent:"flex-end",marginTop:20,borderTop:"1px solid #1e2535",paddingTop:16}}>
      <button onClick={onClose} style={{padding:"8px 18px",background:"#1e2535",border:"1px solid #2d3555",borderRadius:7,color:"#94a3b8",cursor:"pointer"}}>Cancel</button>
      <button onClick={()=>onSave(sites)} style={BG}>Save All Sites</button>
    </div>
  </Overlay>;
}

// ─── Clients Modal ────────────────────────────────────────────────────────────
function ClientsModal({clients,onSave,onClose}){
  const [list,setList]=useState(clients.map(c=>({...c})));
  const [nn,setNn]=useState(""),[ne,setNe]=useState(""),[np,setNp]=useState(""),[nc,setNc]=useState(PRESET_COLORS[2]);
  const add=()=>{if(!nn.trim())return;setList(l=>[...l,{id:"c"+Date.now(),name:nn.trim(),email:ne,phone:np,color:nc,notes:""}]);setNn("");setNe("");setNp("");};
  const up=(id,k,v)=>setList(l=>l.map(x=>x.id===id?{...x,[k]:v}:x));
  const rm=id=>{if(window.confirm("Delete?"))setList(l=>l.filter(x=>x.id!==id));};
  return <Overlay onClose={onClose}>
    <MH title="👔 Manage Clients" onClose={onClose}/>
    <Sec title="Add New Client">
      <div style={{display:"flex",gap:9,alignItems:"flex-end",flexWrap:"wrap"}}>
        <div style={{flex:2,minWidth:140}}><label style={LBL}>Name</label><input value={nn} onChange={e=>setNn(e.target.value)} style={INP}/></div>
        <div style={{flex:1,minWidth:130}}><label style={LBL}>Email</label><input value={ne} onChange={e=>setNe(e.target.value)} type="email" style={INP}/></div>
        <div style={{flex:1,minWidth:110}}><label style={LBL}>Phone</label><input value={np} onChange={e=>setNp(e.target.value)} style={INP}/></div>
        <div><label style={LBL}>Colour</label><div style={{display:"flex",gap:3,flexWrap:"wrap",width:100}}>{PRESET_COLORS.slice(0,9).map(c=><div key={c} onClick={()=>setNc(c)} style={{width:18,height:18,borderRadius:3,background:c,cursor:"pointer",border:nc===c?"3px solid #fff":"2px solid transparent",boxSizing:"border-box"}}/>)}</div></div>
        <button onClick={add} style={{...BP,whiteSpace:"nowrap"}}>+ Add</button>
      </div>
    </Sec>
    {list.length===0&&<div style={{textAlign:"center",padding:28,color:"#374151",fontSize:13}}>No clients yet.</div>}
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
      {list.map(c=><div key={c.id} style={{background:"#0f1421",borderRadius:10,padding:14,border:`1px solid ${c.color}44`}}>
        <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:9}}>
          <span style={{width:11,height:11,borderRadius:"50%",background:c.color,flexShrink:0}}/>
          <input value={c.name} onChange={e=>up(c.id,"name",e.target.value)} style={{flex:1,background:"#1a1f2e",border:"1px solid #2d3555",borderRadius:5,padding:"5px 8px",color:"#e2e8f0",fontSize:13,fontWeight:700,outline:"none"}}/>
          <button onClick={()=>rm(c.id)} style={{background:"#2d1515",border:"1px solid #ef4444",borderRadius:5,color:"#f87171",cursor:"pointer",fontSize:11,padding:"3px 7px",fontWeight:700}}>Del</button>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"0 9px"}}>
          <div style={{marginBottom:8}}><label style={LBL}>Email</label><input value={c.email||""} onChange={e=>up(c.id,"email",e.target.value)} style={{...INP,fontSize:12,padding:"5px 7px"}}/></div>
          <div style={{marginBottom:8}}><label style={LBL}>Phone</label><input value={c.phone||""} onChange={e=>up(c.id,"phone",e.target.value)} style={{...INP,fontSize:12,padding:"5px 7px"}}/></div>
        </div>
        <div style={{marginBottom:8}}><label style={LBL}>Notes</label><input value={c.notes||""} onChange={e=>up(c.id,"notes",e.target.value)} style={{...INP,fontSize:12,padding:"5px 7px"}}/></div>
        <div style={{display:"flex",gap:3,flexWrap:"wrap"}}>{PRESET_COLORS.slice(0,12).map(col=><div key={col} onClick={()=>up(c.id,"color",col)} style={{width:14,height:14,borderRadius:3,background:col,cursor:"pointer",border:c.color===col?"2px solid #fff":"1px solid transparent"}}/>)}</div>
      </div>)}
    </div>
    <div style={{display:"flex",gap:10,justifyContent:"flex-end",marginTop:20,borderTop:"1px solid #1e2535",paddingTop:16}}>
      <button onClick={onClose} style={{padding:"8px 18px",background:"#1e2535",border:"1px solid #2d3555",borderRadius:7,color:"#94a3b8",cursor:"pointer"}}>Cancel</button>
      <button onClick={()=>onSave(list)} style={BG}>Save Clients</button>
    </div>
  </Overlay>;
}

// ─── Worker Modal ─────────────────────────────────────────────────────────────
function WorkerModal({worker,onSave,onClose,allSiteNames,allSites,activeDays}){
  const [f,setF]=useState({...worker,days:{...worker.days},hoursPerDay:{...worker.hoursPerDay},overtimeHours:{...worker.overtimeHours},certs:{...worker.certs}});
  const [tab,setTab]=useState("personal");
  const set=(k,v)=>setF(x=>({...x,[k]:v}));
  const setD=(d,v)=>setF(x=>({...x,days:{...x.days,[d]:v}}));
  const setH=(d,v)=>setF(x=>({...x,hoursPerDay:{...x.hoursPerDay,[d]:Number(v)||0}}));
  const setOT=(d,v)=>setF(x=>({...x,overtimeHours:{...x.overtimeHours,[d]:Number(v)||0}}));
  const setC=(k,v)=>setF(x=>({...x,certs:{...x.certs,[k]:v}}));
  const held=CERTS.filter(c=>f.certs?.[c.key]?.held).length;
  const alerts=CERTS.filter(c=>{const s=cSt(c,f);return s==="expired"||s==="expiring";}).length;
  const SC={valid:"#34d399",expiring:"#fbbf24",expired:"#f87171",missing:"#374151"};
  return <Overlay onClose={onClose} wide>
    <MH title={worker.name?`Edit: ${worker.name}`:"Add New Worker"} onClose={onClose}/>
    <TabBar tabs={[["personal","👤 Personal"],["schedule","📅 Schedule"],["pay","💷 Pay & OT"],["certs",`🛡 Certs ${alerts>0?`⚠${alerts}`:`(${held})`}`]]} active={tab} onChange={setTab}/>
    {tab==="personal"&&<div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"0 14px"}}>
        <FI label="Full Name" value={f.name} onChange={v=>set("name",v)}/><FSel label="Company" value={f.company} onChange={v=>set("company",v)} options={COMPANIES}/>
        <FSel label="Position" value={f.position} onChange={v=>set("position",v)} options={POSITIONS}/><FI label="Scope" value={f.scope} onChange={v=>set("scope",v)}/>
        <FI label="Date of Birth" value={f.dob} onChange={v=>set("dob",v)} type="date"/><FI label="Contact Number" value={f.contact} onChange={v=>set("contact",v)}/>
        <FI label="Email" value={f.email} onChange={v=>set("email",v)} type="email"/><FI label="Comments" value={f.comments} onChange={v=>set("comments",v)}/>
      </div>
      <Sec title="Address"><FI label="Full Address" value={f.address} onChange={v=>set("address",v)}/></Sec>
      <Sec title="Bank Details">
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:"0 12px"}}>
          <FI label="Bank Name" value={f.bankName} onChange={v=>set("bankName",v)}/><FI label="Account Number" value={f.bankAccount} onChange={v=>set("bankAccount",v)}/><FI label="Sort Code" value={f.bankSort} onChange={v=>set("bankSort",v)} placeholder="00-00-00"/>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"0 12px"}}>
          <FI label="NINO" value={f.nino} onChange={v=>set("nino",v)} placeholder="AB 12 34 56 C"/><FI label="UTR Number" value={f.utr} onChange={v=>set("utr",v)}/>
        </div>
      </Sec>
      <Sec title="Next of Kin">
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"0 12px"}}>
          <FI label="Name" value={f.nextOfKin} onChange={v=>set("nextOfKin",v)}/><FI label="Phone" value={f.nextOfKinPhone} onChange={v=>set("nextOfKinPhone",v)}/>
        </div>
      </Sec>
    </div>}
    {tab==="schedule"&&<div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
        <label style={LBL}>Site Allocation</label>
        <div style={{display:"flex",gap:6}}>
          <input id="fill-all" list="fill-l" placeholder="Fill all days…" style={{...INP,width:180,padding:"5px 8px",fontSize:12}}/>
          <datalist id="fill-l">{allSiteNames.map(s=><option key={s} value={s}/>)}</datalist>
          <button onClick={()=>{const v=document.getElementById("fill-all")?.value;if(v){const nd={};activeDays.forEach(d=>nd[d]=v);setF(x=>({...x,days:{...x.days,...nd}}));}}} style={{...BP,padding:"5px 11px",fontSize:12}}>Apply All</button>
        </div>
      </div>
      <div style={{display:"grid",gridTemplateColumns:`repeat(${activeDays.length},1fr)`,gap:7}}>
        {activeDays.map(d=><div key={d}>
          <div style={{fontSize:11,color:WEEKEND_DAYS.includes(d)?"#fbbf24":"#94a3b8",marginBottom:3,textAlign:"center",fontWeight:700}}>{d}</div>
          <div style={{height:3,borderRadius:2,background:getSiteColor(f.days[d],allSites),marginBottom:4}}/>
          <input list="sites-l" value={f.days[d]??""} onChange={e=>setD(d,e.target.value)} style={{...INP,border:`1px solid ${getSiteColor(f.days[d],allSites)||"#2d3555"}`,padding:"5px 6px",fontSize:11}}/>
        </div>)}
      </div>
      <datalist id="sites-l">{allSiteNames.map(s=><option key={s} value={s}/>)}</datalist>
    </div>}
    {tab==="pay"&&<div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:"0 14px"}}>
        <FI label="Agreed Rate £/hr" value={f.agreedRate} onChange={v=>set("agreedRate",v?Number(v):null)} type="number"/>
        <FI label="Actual Rate £/hr" value={f.actualRate} onChange={v=>set("actualRate",v?Number(v):null)} type="number"/>
        <div style={{marginBottom:11}}><label style={LBL}>Tax Rate</label>
          <select value={f.taxRate??0} onChange={e=>set("taxRate",Number(e.target.value))} style={{...INP,cursor:"pointer",border:`1px solid ${f.taxRate===0.30?"#f87171":f.taxRate===0.20?"#fbbf24":"#34d399"}`}}>
            <option value={0}>0% — No Tax</option><option value={0.20}>20% — Basic Rate</option><option value={0.30}>30% — Higher Rate</option>
          </select>
        </div>
      </div>
      <Sec title="Overtime" color="#fbbf24">
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"0 14px",marginBottom:12}}>
          <div style={{marginBottom:11}}><label style={LBL}>OT Multiplier</label>
            <select value={f.overtimeMultiplier??1.5} onChange={e=>set("overtimeMultiplier",Number(e.target.value))} style={{...INP,cursor:"pointer"}}>
              <option value={1.25}>×1.25</option><option value={1.5}>×1.5 (Standard)</option><option value={2}>×2.0 (Double Time)</option>
            </select>
          </div>
          <FI label="Custom OT Rate £/hr" value={f.customOTRate} onChange={v=>set("customOTRate",v?Number(v):null)} type="number" placeholder="Overrides multiplier"/>
        </div>
        <label style={LBL}>OT Hours Per Day</label>
        <div style={{display:"grid",gridTemplateColumns:`repeat(${activeDays.length},1fr)`,gap:7}}>
          {activeDays.map(d=>{const w=f.days[d]&&!isOff(f.days[d]);return <div key={d} style={{opacity:w?1:0.3}}>
            <div style={{fontSize:11,color:WEEKEND_DAYS.includes(d)?"#fbbf24":"#94a3b8",marginBottom:3,textAlign:"center",fontWeight:700}}>{d}</div>
            <input type="number" min="0" max="12" value={f.overtimeHours?.[d]??0} onChange={e=>setOT(d,e.target.value)} disabled={!w} style={{...INP,textAlign:"center",padding:"5px 6px",fontSize:12,color:"#fbbf24"}}/>
          </div>;})}
        </div>
      </Sec>
      {f.agreedRate&&(()=>{const {stdH,otH,gross,tax,net}=calcPay(f,activeDays,{});return <div style={{background:"#0d2218",border:"1px solid #065f46",borderRadius:10,padding:14}}>
        <div style={{fontSize:11,color:"#34d399",fontWeight:700,textTransform:"uppercase",marginBottom:10}}>Live Payroll Preview</div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:8}}>
          {[["Std h",`${stdH}h`,"#60a5fa"],["OT h",`${otH}h`,"#fbbf24"],["Gross",`£${gross.toFixed(2)}`,"#34d399"],["Tax",`-£${tax.toFixed(2)}`,"#f87171"],["Net",`£${net.toFixed(2)}`,"#a78bfa"]].map(([l,v,c])=>(
            <div key={l} style={{textAlign:"center"}}><div style={{fontSize:10,color:"#64748b"}}>{l}</div><div style={{fontSize:15,fontWeight:800,color:c}}>{v}</div></div>
          ))}
        </div>
      </div>;})()} 
    </div>}
    {tab==="certs"&&<div>
      <div style={{fontSize:12,color:"#64748b",marginBottom:14}}>Tick each held certification and set expiry dates.</div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"0 16px"}}>
        {CERTS.map(cert=>{const val=f.certs[cert.key]||{held:false,expiry:""};const status=cSt(cert,f);
          return <div key={cert.key} style={{marginBottom:9,padding:"9px 11px",background:"#0f1421",borderRadius:8,border:`1px solid ${val.held?SC[status]+"66":"#1e2535"}`}}>
            <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:cert.hasExpiry&&val.held?7:0}}>
              <input type="checkbox" checked={!!val.held} onChange={e=>setC(cert.key,{...val,held:e.target.checked})} style={{width:15,height:15,cursor:"pointer",accentColor:"#3b82f6"}}/>
              <span style={{fontSize:12,color:val.held?"#e2e8f0":"#64748b",fontWeight:val.held?600:400,flex:1}}>{cert.label}</span>
              {val.held&&<span style={{fontSize:10,color:SC[status],fontWeight:700,textTransform:"uppercase"}}>{status}</span>}
            </div>
            {cert.hasExpiry&&val.held&&<div style={{display:"flex",alignItems:"center",gap:6}}>
              <span style={{fontSize:11,color:"#64748b",minWidth:48}}>Expiry:</span>
              <input type="date" value={val.expiry||""} onChange={e=>setC(cert.key,{...val,expiry:e.target.value})} style={{flex:1,background:"#1a1f2e",border:"1px solid #2d3555",borderRadius:5,padding:"3px 6px",color:"#e2e8f0",fontSize:12,outline:"none"}}/>
            </div>}
          </div>;
        })}
      </div>
    </div>}
    <div style={{display:"flex",gap:10,justifyContent:"flex-end",marginTop:20,borderTop:"1px solid #1e2535",paddingTop:16}}>
      <button onClick={onClose} style={{padding:"8px 18px",background:"#1e2535",border:"1px solid #2d3555",borderRadius:7,color:"#94a3b8",cursor:"pointer"}}>Cancel</button>
      <button onClick={()=>onSave(f)} style={BP}>Save Worker</button>
    </div>
  </Overlay>;
}

// ─── Supporting Views ─────────────────────────────────────────────────────────
function CertView({workers}){
  const [fs,setFs]=useState("all");
  const rows=useMemo(()=>workers.map(w=>{const st=CERTS.map(c=>cSt(c,w));return {...w,expired:st.filter(s=>s==="expired").length,expiring:st.filter(s=>s==="expiring").length,valid:st.filter(s=>s==="valid").length};}),[workers]);
  const fil=fs==="all"?rows:fs==="expired"?rows.filter(r=>r.expired>0):fs==="expiring"?rows.filter(r=>r.expiring>0):rows.filter(r=>r.valid===0&&r.expired===0&&r.expiring===0);
  return <div style={{padding:"14px 18px"}}>
    <div style={{display:"flex",gap:7,marginBottom:14,flexWrap:"wrap"}}>
      {[["all","All","#64748b"],["expired","Expired","#f87171"],["expiring","Expiring","#fbbf24"],["none","No Certs","#374151"]].map(([v,l,c])=>(
        <button key={v} onClick={()=>setFs(v)} style={{padding:"5px 12px",background:fs===v?c+"22":"#1a1f2e",border:`1px solid ${fs===v?c:"#2d3555"}`,borderRadius:7,color:fs===v?c:"#64748b",cursor:"pointer",fontSize:11,fontWeight:fs===v?700:400}}>{l}</button>
      ))}
    </div>
    <div style={{overflowX:"auto"}}>
      <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
        <thead><tr><th style={{...TH,minWidth:145}}>Worker</th><th style={TH}>Summary</th>
          {CERTS.slice(0,13).map(c=><th key={c.key} style={{...TH,minWidth:22,padding:"8px 3px",fontSize:9,textAlign:"center"}} title={c.label}>{c.label.split(" ").map(w=>w[0]).join("").slice(0,5)}</th>)}
          <th style={TH}>More</th></tr></thead>
        <tbody>{fil.map((w,i)=><tr key={w.id} style={{background:i%2===0?"#111827":"#0f1421"}}>
          <td style={{...TD,fontWeight:600,color:"#f1f5f9"}}><div>{w.name}</div><div style={{fontSize:10,color:"#64748b"}}>{w.position}</div></td>
          <td style={TD}><div style={{display:"flex",gap:5}}>
            {w.valid>0&&<span style={{fontSize:11,color:"#34d399",fontWeight:700}}>✓{w.valid}</span>}
            {w.expiring>0&&<span style={{fontSize:11,color:"#fbbf24",fontWeight:700}}>⚠{w.expiring}</span>}
            {w.expired>0&&<span style={{fontSize:11,color:"#f87171",fontWeight:700}}>✗{w.expired}</span>}
            {w.valid===0&&w.expiring===0&&w.expired===0&&<span style={{color:"#374151"}}>None</span>}
          </div></td>
          {CERTS.slice(0,13).map(c=><td key={c.key} style={{...TD,textAlign:"center",padding:"6px 3px"}}><CDot status={cSt(c,w)} label={c.label}/></td>)}
          <td style={TD}><div style={{display:"flex",flexWrap:"wrap",gap:2}}>{CERTS.slice(13).map(c=><CDot key={c.key} status={cSt(c,w)} label={c.label}/>)}</div></td>
        </tr>)}</tbody>
      </table>
    </div>
    <div style={{marginTop:10,display:"flex",gap:12,flexWrap:"wrap"}}>
      {[["valid","#34d399"],["expiring","#fbbf24"],["expired","#f87171"],["missing","#2d3555"]].map(([s,c])=>(
        <div key={s} style={{display:"flex",alignItems:"center",gap:5,fontSize:11,color:"#94a3b8"}}><span style={{width:9,height:9,borderRadius:"50%",background:c,border:`1px solid ${c}`,display:"inline-block"}}/>{s}</div>
      ))}
    </div>
  </div>;
}

function PayrollView({workers,activeDays,siteHours,allSites,weekLabel}){
  const rows=useMemo(()=>workers.map(w=>({...w,...calcPay(w,activeDays,siteHours)})),[workers,activeDays,siteHours]);
  const tot=rows.reduce((a,r)=>({h:a.h+r.stdH,ot:a.ot+r.otH,g:a.g+r.gross,t:a.t+r.tax,n:a.n+r.net}),{h:0,ot:0,g:0,t:0,n:0});
  return <div style={{padding:"14px 18px"}}>
    <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:9,marginBottom:16}}>
      {[["Std Hrs",tot.h+"h","#60a5fa"],["OT Hrs",tot.ot+"h","#fbbf24"],["Gross","£"+tot.g.toFixed(2),"#34d399"],["Tax","£"+tot.t.toFixed(2),"#f87171"],["Net Pay","£"+tot.n.toFixed(2),"#a78bfa"]].map(([l,v,c])=>(
        <div key={l} style={{background:"#1a1f2e",border:`1px solid ${c}44`,borderRadius:10,padding:"10px 12px"}}>
          <div style={{fontSize:10,color:"#64748b",fontWeight:700,textTransform:"uppercase"}}>{l}</div>
          <div style={{fontSize:18,fontWeight:800,color:c}}>{v}</div>
        </div>
      ))}
    </div>
    <div style={{overflowX:"auto"}}>
      <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
        <thead><tr><th style={{...TH,minWidth:140}}>Worker</th>
          {activeDays.map(d=><th key={d} style={{...TH,minWidth:100}}>{d}</th>)}
          <th style={TH}>Rate</th><th style={TH}>OT</th><th style={TH}>Tax%</th><th style={TH}>Std h</th><th style={TH}>OT h</th><th style={TH}>Gross</th><th style={TH}>-Tax</th><th style={{...TH,color:"#a78bfa"}}>Net</th><th style={TH}>Payslip</th>
        </tr></thead>
        <tbody>{rows.map((w,i)=><tr key={w.id} style={{background:i%2===0?"#111827":"#0f1421"}}>
          <td style={{...TD,fontWeight:600,color:"#f1f5f9"}}><div>{w.name}</div><div style={{fontSize:10,color:"#64748b"}}>{w.position}</div></td>
          {activeDays.map(d=>{const b=w.bd[d];return <td key={d} style={TD}>{b?<div><span style={{display:"inline-block",padding:"1px 6px",borderRadius:4,fontSize:10,fontWeight:600,color:"#fff",background:getSiteColor(b.site,allSites),marginBottom:2}}>{b.site.split("-")[0].trim()}</span><div style={{fontSize:10,color:"#60a5fa"}}>{b.hours}h{b.ot>0?<span style={{color:"#fbbf24"}}>+{b.ot}OT</span>:""}</div></div>:<span style={{color:"#374151"}}>—</span>}</td>;})}
          <td style={{...TD,color:"#34d399",fontWeight:600}}>{w.agreedRate?`£${w.agreedRate}`:"—"}</td>
          <td style={{...TD,color:"#fbbf24",fontSize:11,fontWeight:700}}>{w.customOTRate?`£${w.customOTRate}`:w.otH>0?`×${w.overtimeMultiplier||1.5}`:"—"}</td>
          <td style={TD}><span style={{fontSize:11,fontWeight:700,color:w.taxRate===0.30?"#f87171":w.taxRate===0.20?"#fbbf24":"#34d399"}}>{Math.round((w.taxRate||0)*100)}%</span></td>
          <td style={{...TD,color:"#60a5fa",fontWeight:700}}>{w.stdH}h</td><td style={{...TD,color:"#fbbf24",fontWeight:700}}>{w.otH>0?w.otH+"h":"—"}</td>
          <td style={{...TD,color:"#34d399",fontWeight:700}}>£{w.gross.toFixed(2)}</td><td style={{...TD,color:"#f87171"}}>£{w.tax.toFixed(2)}</td>
          <td style={{...TD,color:"#a78bfa",fontWeight:800,fontSize:13}}>£{w.net.toFixed(2)}</td>
          <td style={TD}>
            <button onClick={()=>exportPayslip(w,activeDays,weekLabel,siteHours)}
              style={{padding:"4px 9px",background:"#1a2535",border:"1px solid #10b981",borderRadius:5,color:"#34d399",cursor:"pointer",fontSize:10,fontWeight:700,whiteSpace:"nowrap"}}>
              💷 Payslip
            </button>
          </td>
        </tr>)}</tbody>
        <tfoot><tr style={{background:"#0d1117",borderTop:"2px solid #2d3555"}}>
          <td colSpan={1+activeDays.length} style={{...TD,fontWeight:700,color:"#94a3b8"}}>TOTALS</td>
          <td style={TD}/><td style={TD}/><td style={TD}/>
          <td style={{...TD,color:"#60a5fa",fontWeight:800}}>{tot.h}h</td><td style={{...TD,color:"#fbbf24",fontWeight:800}}>{tot.ot>0?tot.ot+"h":"—"}</td>
          <td style={{...TD,color:"#34d399",fontWeight:800}}>£{tot.g.toFixed(2)}</td><td style={{...TD,color:"#f87171",fontWeight:800}}>£{tot.t.toFixed(2)}</td>
          <td style={{...TD,color:"#a78bfa",fontWeight:800,fontSize:13}}>£{tot.n.toFixed(2)}</td>
          <td style={TD}/>
        </tr></tfoot>
      </table>
    </div>
  </div>;
}

function ClientCostView({workers,clients,allSites,activeDays,siteHours}){
  const data=useMemo(()=>{
    const sc={};
    workers.forEach(w=>{const {bd}=calcPay(w,activeDays,siteHours);Object.values(bd).forEach(({site,gross})=>{if(!sc[site])sc[site]={gross:0,wIds:new Set()};sc[site].gross+=gross;sc[site].wIds.add(w.id);});});
    const byC={};clients.forEach(c=>{byC[c.id]={client:c,sites:{},total:0,wIds:new Set()};});
    byC["__none"]={client:{id:"__none",name:"Unassigned",color:"#374151"},sites:{},total:0,wIds:new Set()};
    Object.entries(sc).forEach(([site,{gross,wIds}])=>{
      const cs=allSites.find(s=>site===s.name||site.includes(s.name));
      const cid=cs?.clientId||"__none";const bkt=byC[cid]||byC["__none"];
      if(!bkt.sites[site])bkt.sites[site]={gross:0,wIds:new Set()};
      bkt.sites[site].gross+=gross;wIds.forEach(id=>bkt.sites[site].wIds.add(id));
      bkt.total+=gross;wIds.forEach(id=>bkt.wIds.add(id));
    });
    return Object.values(byC).filter(d=>d.total>0);
  },[workers,clients,allSites,activeDays,siteHours]);
  const grand=data.reduce((a,d)=>a+d.total,0);
  return <div style={{padding:"14px 18px"}}>
    <div style={{background:"#1a1f2e",border:"1px solid #a78bfa55",borderRadius:10,padding:"11px 16px",marginBottom:16,display:"inline-block"}}>
      <div style={{fontSize:10,color:"#64748b",fontWeight:700,textTransform:"uppercase"}}>Total Labour Cost</div>
      <div style={{fontSize:24,fontWeight:800,color:"#a78bfa"}}>£{grand.toFixed(2)}</div>
    </div>
    {data.length===0&&<div style={{textAlign:"center",padding:40,color:"#374151",fontSize:13}}>No cost data yet.</div>}
    {data.map(({client,sites,total,wIds})=>(
      <div key={client.id} style={{background:"#111827",border:`1px solid ${client.color}44`,borderRadius:11,padding:15,marginBottom:12}}>
        <div style={{display:"flex",alignItems:"center",gap:9,marginBottom:11}}>
          <span style={{width:11,height:11,borderRadius:"50%",background:client.color,flexShrink:0}}/>
          <span style={{fontWeight:800,color:"#f1f5f9",fontSize:14}}>{client.name}</span>
          <span style={{marginLeft:"auto",fontWeight:800,color:"#34d399",fontSize:17}}>£{total.toFixed(2)}</span>
          <span style={{fontSize:11,color:"#64748b"}}>{wIds.size} workers</span>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(170px,1fr))",gap:8}}>
          {Object.entries(sites).map(([site,{gross,wIds:sw}])=>(
            <div key={site} style={{background:"#1a1f2e",border:`1px solid ${getSiteColor(site,allSites)}33`,borderRadius:8,padding:"9px 12px"}}>
              <div style={{display:"flex",alignItems:"center",gap:5,marginBottom:4}}>
                <span style={{width:7,height:7,borderRadius:"50%",background:getSiteColor(site,allSites),flexShrink:0}}/>
                <span style={{fontSize:11,color:"#cbd5e1",fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{site}</span>
              </div>
              <div style={{fontSize:18,fontWeight:800,color:"#34d399"}}>£{gross.toFixed(2)}</div>
              <div style={{fontSize:10,color:"#64748b"}}>{sw.size} worker{sw.size!==1?"s":""}</div>
            </div>
          ))}
        </div>
      </div>
    ))}
  </div>;
}

// ─── Main App ──────────────────────────────────────────────────────────────────
export default function App(){
  const [workers,setWorkers]=useState([]);
  const [weekLabel,setWeekLabel]=useState(formatWeekLabel(new Date()));
  const [showWeekend,setShowWeekend]=useState(false);
  const [allSites,setAllSites]=useState(DEFAULT_BUILTIN_SITES); // unified site list
  const [clients,setClients]=useState(INIT_CLIENTS);
  const [siteHours,setSiteHours]=useState({});
  const [filter,setFilter]=useState({name:"",position:"",site:""});
  const [view,setView]=useState("schedule");
  const [modal,setModal]=useState(null);
  const [loading,setLoading]=useState(true);
  const [syncStatus,setSyncStatus]=useState("saved");
  const saveTimer=useRef(null);

  useEffect(()=>{
    async function loadAll(){
      try {
        setLoading(true);
        const wRows=await sbGet("workers","select=id,data&order=data->name");
        if(wRows.length>0) setWorkers(wRows.map(r=>({...mkW(),...r.data,id:r.id})));
        else { setWorkers(INIT_W); await Promise.all(INIT_W.map(w=>sbUpsert("workers",[{id:w.id,data:w}]))); }
        const cfgRows=await sbGet("app_config","select=key,value");
        const cfg=Object.fromEntries(cfgRows.map(r=>[r.key,r.value]));
        if(cfg.week_label) setWeekLabel(cfg.week_label);
        if(cfg.show_weekend!==undefined) setShowWeekend(cfg.show_weekend);
        if(cfg.all_sites&&cfg.all_sites.length>0) setAllSites(cfg.all_sites);
        if(cfg.clients) setClients(cfg.clients);
        if(cfg.site_hours) setSiteHours(cfg.site_hours);
      } catch(e){ console.error("Load error:",e); setSyncStatus("error"); setWorkers(INIT_W); }
      finally { setLoading(false); }
    }
    loadAll();
  },[]);

  useEffect(()=>{
    if(loading) return;
    setSyncStatus("saving");
    if(saveTimer.current) clearTimeout(saveTimer.current);
    saveTimer.current=setTimeout(async()=>{
      try {
        await sbUpsert("app_config",[
          {key:"week_label",value:weekLabel},
          {key:"show_weekend",value:showWeekend},
          {key:"all_sites",value:allSites},
          {key:"clients",value:clients},
          {key:"site_hours",value:siteHours},
        ]);
        setSyncStatus("saved");
      } catch(e){ setSyncStatus("error"); }
    },800);
  },[weekLabel,showWeekend,allSites,clients,siteHours,loading]);

  const saveWorker=async(w)=>{
    const isNew=!workers.find(x=>x.id===w.id);
    if(isNew) setWorkers(ws=>[...ws,w]); else setWorkers(ws=>ws.map(x=>x.id===w.id?w:x));
    setModal(null); setSyncStatus("saving");
    try { await sbUpsert("workers",[{id:w.id,data:w}]); setSyncStatus("saved"); }
    catch(e){ setSyncStatus("error"); }
  };
  const delWorker=async(id)=>{
    if(!window.confirm("Remove this worker?")) return;
    setWorkers(ws=>ws.filter(w=>w.id!==id)); setSyncStatus("saving");
    try { await sbDelete("workers",`id=eq.${id}`); setSyncStatus("saved"); }
    catch(e){ setSyncStatus("error"); }
  };
  const updateCell=async(wId,day,val)=>{
    const updated=workers.map(w=>w.id===wId?{...w,days:{...w.days,[day]:val}}:w);
    setWorkers(updated); setSyncStatus("saving");
    try { const w=updated.find(x=>x.id===wId); await sbUpsert("workers",[{id:wId,data:w}]); setSyncStatus("saved"); }
    catch(e){ setSyncStatus("error"); }
  };

  const activeDays=useMemo(()=>showWeekend?ALL_DAYS:BASE_DAYS,[showWeekend]);
  const allSiteNames=useMemo(()=>{
    const s=new Set(allSites.map(x=>x.name));
    workers.forEach(w=>ALL_DAYS.forEach(d=>{if(w.days[d])s.add(w.days[d].trim());}));
    return Array.from(s).filter(Boolean).sort();
  },[allSites,workers]);

  const filtered=useMemo(()=>workers.filter(w=>{
    if(filter.name&&!w.name.toLowerCase().includes(filter.name.toLowerCase()))return false;
    if(filter.position&&w.position!==filter.position)return false;
    if(filter.site&&!Object.values(w.days).some(d=>d&&d.toLowerCase().includes(filter.site.toLowerCase())))return false;
    return true;
  }),[workers,filter]);

  const stats=useMemo(()=>{
    const onHol=workers.filter(w=>activeDays.some(d=>w.days[d]?.includes("Holiday"))).length;
    const off=workers.filter(w=>activeDays.every(d=>isOff(w.days[d]))).length;
    const alerts=workers.reduce((n,w)=>n+CERTS.filter(c=>{const s=cSt(c,w);return s==="expired"||s==="expiring";}).length,0);
    const pay=workers.reduce((a,w)=>{const {gross,net}=calcPay(w,activeDays,siteHours);return{g:a.g+gross,n:a.n+net};},{g:0,n:0});
    return{total:workers.length,onHol,off,alerts,...pay};
  },[workers,activeDays,siteHours]);

  const VIEWS=[["schedule","📋 Schedule"],["site","📍 By Site"],["certs","🛡 Certs"],["payroll","💷 Payroll"],["costs","👔 Client Costs"],["stats","📊 Stats"]];

  if(loading) return <div style={{minHeight:"100vh",background:"#0d1117",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",gap:16}}>
    <div style={{width:48,height:48,background:"linear-gradient(135deg,#3b82f6,#6366f1)",borderRadius:12,display:"flex",alignItems:"center",justifyContent:"center",fontSize:24}}>🏗</div>
    <div style={{color:"#60a5fa",fontSize:16,fontWeight:700}}>Loading Labour Schedule…</div>
    <div style={{color:"#374151",fontSize:13}}>Connecting to cloud database…</div>
    <div style={{width:200,height:3,background:"#1e2535",borderRadius:3,overflow:"hidden"}}>
      <div style={{height:"100%",background:"linear-gradient(90deg,#3b82f6,#6366f1)",borderRadius:3,animation:"slide 1.5s ease-in-out infinite"}}/>
    </div>
    <style>{`@keyframes slide{0%{width:0%;margin-left:0%}50%{width:60%;margin-left:20%}100%{width:0%;margin-left:100%}}`}</style>
  </div>;

  return <div style={{minHeight:"100vh",background:"#0d1117",fontFamily:"system-ui,'Segoe UI',sans-serif",color:"#e2e8f0",fontSize:13}}>

    {/* Header */}
    <div style={{background:"linear-gradient(135deg,#0f172a,#1a1f2e)",borderBottom:"1px solid #1e2535",padding:"13px 18px"}}>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:9,marginBottom:12}}>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <div style={{width:36,height:36,background:"linear-gradient(135deg,#3b82f6,#6366f1)",borderRadius:9,display:"flex",alignItems:"center",justifyContent:"center",fontSize:17}}>🏗</div>
          <div>
            <div style={{fontSize:17,fontWeight:800,color:"#f1f5f9",letterSpacing:"-0.02em"}}>Labour Schedule</div>
            <div style={{display:"flex",alignItems:"center",gap:6,marginTop:2}}>
              <button onClick={()=>setWeekLabel(addWeeks(weekLabel,-1))} style={{background:"#1e2535",border:"1px solid #2d3555",borderRadius:5,color:"#94a3b8",cursor:"pointer",fontSize:13,padding:"1px 7px",fontWeight:700,lineHeight:1.4}}>‹</button>
              <span style={{fontSize:10,color:"#64748b"}}>WC:</span>
              <input value={weekLabel} onChange={e=>setWeekLabel(e.target.value)} style={{background:"none",border:"none",borderBottom:"1px solid #2d3555",color:"#60a5fa",fontWeight:600,fontSize:12,outline:"none",width:115}}/>
              <button onClick={()=>setWeekLabel(addWeeks(weekLabel,1))} style={{background:"#1e2535",border:"1px solid #2d3555",borderRadius:5,color:"#94a3b8",cursor:"pointer",fontSize:13,padding:"1px 7px",fontWeight:700,lineHeight:1.4}}>›</button>
              <button onClick={()=>setWeekLabel(formatWeekLabel(new Date()))} style={{background:"#1e2535",border:"1px solid #2d3555",borderRadius:5,color:"#64748b",cursor:"pointer",fontSize:10,padding:"2px 7px",fontWeight:700}}>Today</button>
            </div>
          </div>
        </div>
        <div style={{display:"flex",gap:6,flexWrap:"wrap",alignItems:"center"}}>
          <div style={{display:"flex",alignItems:"center",gap:5,fontSize:11,color:syncStatus==="saved"?"#34d399":syncStatus==="error"?"#f87171":"#fbbf24"}}>
            <span style={{width:7,height:7,borderRadius:"50%",background:syncStatus==="saved"?"#34d399":syncStatus==="error"?"#f87171":"#fbbf24",display:"inline-block"}}/>
            {syncStatus==="saved"?"☁ Saved":syncStatus==="error"?"⚠ Error":"Saving…"}
          </div>
          <button onClick={()=>setShowWeekend(s=>!s)} style={{padding:"6px 11px",background:showWeekend?"#1a3020":"#1a1f2e",border:`1px solid ${showWeekend?"#10b981":"#2d3555"}`,borderRadius:7,color:showWeekend?"#34d399":"#64748b",cursor:"pointer",fontSize:11,fontWeight:700}}>{showWeekend?"✓ Weekend":"+ Weekend"}</button>
          <button onClick={()=>setModal({type:"sites"})} style={{padding:"6px 11px",background:"#1a1f2e",border:"1px solid #f59e0b",borderRadius:7,color:"#fbbf24",cursor:"pointer",fontSize:11,fontWeight:700}}>🏗 Sites</button>
          <button onClick={()=>setModal({type:"clients"})} style={{padding:"6px 11px",background:"#1a1f2e",border:"1px solid #8b5cf6",borderRadius:7,color:"#a78bfa",cursor:"pointer",fontSize:11,fontWeight:700}}>👔 Clients</button>
          {view==="schedule"&&<button onClick={()=>exportSchedulePDF(filtered,activeDays,weekLabel,allSites)} style={{padding:"6px 11px",background:"#1a1f2e",border:"1px solid #ef4444",borderRadius:7,color:"#f87171",cursor:"pointer",fontSize:11,fontWeight:700}}>📄 PDF</button>}
          <button onClick={()=>doExcel(workers,weekLabel,activeDays,siteHours,clients,allSites)} style={{...BG,padding:"6px 13px",fontSize:11}}>⬇ Excel</button>
          <button onClick={()=>setModal({type:"worker",worker:mkW()})} style={{...BP,padding:"6px 13px",fontSize:11}}>+ Worker</button>
        </div>
      </div>
      <div style={{display:"flex",gap:7,flexWrap:"wrap"}}>
        {[{l:"Workers",v:stats.total,c:"#60a5fa"},{l:"On Holiday",v:stats.onHol,c:"#fbbf24"},{l:"Off",v:stats.off,c:"#94a3b8"},{l:"Cert Alerts",v:stats.alerts,c:stats.alerts>0?"#fbbf24":"#34d399"},{l:"Gross",v:stats.g>0?`£${stats.g.toFixed(0)}`:"—",c:"#34d399"},{l:"Net",v:stats.g>0?`£${stats.n.toFixed(0)}`:"—",c:"#a78bfa"},{l:"Clients",v:clients.length,c:"#8b5cf6"}].map(s=>(
          <div key={s.l} style={{background:"#111827",border:"1px solid #1e2535",borderRadius:9,padding:"6px 12px"}}>
            <div style={{fontSize:10,color:"#64748b",fontWeight:700,textTransform:"uppercase"}}>{s.l}</div>
            <div style={{fontSize:15,fontWeight:800,color:s.c}}>{s.v}</div>
          </div>
        ))}
      </div>
    </div>

    {/* Tabs + Filters */}
    <div style={{padding:"9px 18px",background:"#111827",borderBottom:"1px solid #1e2535",display:"flex",gap:8,flexWrap:"wrap",alignItems:"center",justifyContent:"space-between"}}>
      <div style={{display:"flex",gap:3,background:"#0d1117",borderRadius:8,padding:3}}>
        {VIEWS.map(([v,l])=><button key={v} onClick={()=>setView(v)} style={{padding:"5px 10px",background:view===v?"#1e3a5f":"transparent",border:view===v?"1px solid #3b82f6":"1px solid transparent",borderRadius:6,color:view===v?"#60a5fa":"#64748b",cursor:"pointer",fontSize:11,fontWeight:view===v?700:400}}>{l}</button>)}
      </div>
      <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
        <input value={filter.name} onChange={e=>setFilter(f=>({...f,name:e.target.value}))} placeholder="🔍 Name…" style={{background:"#1a1f2e",border:"1px solid #2d3555",borderRadius:7,padding:"5px 9px",color:"#e2e8f0",fontSize:11,outline:"none",width:120}}/>
        <select value={filter.position} onChange={e=>setFilter(f=>({...f,position:e.target.value}))} style={{background:"#1a1f2e",border:"1px solid #2d3555",borderRadius:7,padding:"5px 9px",color:filter.position?"#e2e8f0":"#64748b",fontSize:11,outline:"none",cursor:"pointer"}}>
          <option value="">All Positions</option>{POSITIONS.map(p=><option key={p} value={p}>{p}</option>)}
        </select>
        <input value={filter.site} onChange={e=>setFilter(f=>({...f,site:e.target.value}))} placeholder="📍 Site…" style={{background:"#1a1f2e",border:"1px solid #2d3555",borderRadius:7,padding:"5px 9px",color:"#e2e8f0",fontSize:11,outline:"none",width:110}}/>
        {Object.values(filter).some(Boolean)&&<button onClick={()=>setFilter({name:"",position:"",site:""})} style={{padding:"5px 9px",background:"#1e2535",border:"1px solid #f87171",borderRadius:7,color:"#f87171",cursor:"pointer",fontSize:11,fontWeight:700}}>✕</button>}
      </div>
    </div>

    {/* Views */}
    <div style={{paddingBottom:40}}>
      {view==="site"&&<div style={{padding:"12px 18px"}}>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
            <thead><tr><th style={{...TH,minWidth:180}}>Site</th>
              {activeDays.map(d=><th key={d} style={{...TH,minWidth:120,color:WEEKEND_DAYS.includes(d)?"#fbbf24":"#64748b"}}>{d}{WEEKEND_DAYS.includes(d)?" 🟡":""}</th>)}
              <th style={TH}>Total</th></tr></thead>
            <tbody>{(()=>{const sm={};workers.forEach(w=>activeDays.forEach(d=>{const s=(w.days[d]||"").trim();if(s){if(!sm[s])sm[s]={};if(!sm[s][d])sm[s][d]=[];sm[s][d].push(w);}}));
              return Object.keys(sm).sort().map((site,i)=>{const color=getSiteColor(site,allSites);const all=new Set();activeDays.forEach(d=>(sm[site][d]||[]).forEach(w=>all.add(w.id)));
                return <tr key={site} style={{background:i%2===0?"#111827":"#0f1421"}}>
                  <td style={{...TD,borderLeft:`3px solid ${color}`,paddingLeft:10}}><span style={{fontWeight:700,color}}>{site}</span></td>
                  {activeDays.map(d=><td key={d} style={TD}>{(sm[site][d]||[]).map(w=><div key={w.id} style={{fontSize:11,color:"#cbd5e1"}}>{w.name} <span style={{color:"#64748b"}}>({w.position||"—"})</span></div>)}</td>)}
                  <td style={{...TD,textAlign:"center",fontWeight:700,color:"#60a5fa"}}>{all.size}</td>
                </tr>;}
              );})()} 
            </tbody>
          </table>
        </div>
      </div>}
      {view==="certs"&&<CertView workers={filtered}/>}
      {view==="payroll"&&<PayrollView workers={filtered} activeDays={activeDays} siteHours={siteHours} allSites={allSites} weekLabel={weekLabel}/>}
      {view==="costs"&&<ClientCostView workers={workers} clients={clients} allSites={allSites} activeDays={activeDays} siteHours={siteHours}/>}
      {view==="stats"&&<div style={{padding:"14px 18px",display:"grid",gridTemplateColumns:"1fr 1fr",gap:20}}>
        <div>
          <div style={{fontWeight:700,color:"#94a3b8",marginBottom:11,fontSize:13}}>Workers per Site</div>
          <div style={{display:"flex",flexWrap:"wrap",gap:8}}>
            {(()=>{const m={};workers.forEach(w=>activeDays.forEach(d=>{const s=(w.days[d]||"").trim();if(s&&!isOff(s))m[s]=(m[s]||0)+1;}));return Object.entries(m).sort((a,b)=>b[1]-a[1]).map(([site,cnt])=>(
              <div key={site} style={{background:"#1a1f2e",border:`1px solid ${getSiteColor(site,allSites)}`,borderRadius:9,padding:"8px 12px"}}>
                <div style={{fontSize:11,color:getSiteColor(site,allSites),fontWeight:700}}>{site}</div>
                <div style={{fontSize:20,fontWeight:800,color:"#f1f5f9"}}>{cnt}</div>
              </div>
            ));})()}
          </div>
        </div>
        <div>
          <div style={{fontWeight:700,color:"#94a3b8",marginBottom:11,fontSize:13}}>Cert Compliance</div>
          {CERTS.slice(0,12).map(c=>{const held=workers.filter(w=>w.certs?.[c.key]?.held).length;const pct=workers.length>0?Math.round((held/workers.length)*100):0;return <div key={c.key} style={{marginBottom:7}}>
            <div style={{display:"flex",justifyContent:"space-between",fontSize:11,color:"#94a3b8",marginBottom:2}}><span>{c.label}</span><span style={{color:pct>50?"#34d399":"#64748b"}}>{held}/{workers.length}</span></div>
            <div style={{height:4,background:"#1e2535",borderRadius:3}}><div style={{height:"100%",borderRadius:3,background:pct>70?"#34d399":pct>30?"#fbbf24":"#f87171",width:`${pct}%`,transition:"width 0.3s"}}/></div>
          </div>;})}
        </div>
      </div>}

      {view==="schedule"&&<div>
        <div style={{padding:"5px 18px",background:"#0f1421",borderBottom:"1px solid #1e2535",fontSize:11,color:"#64748b"}}>
          💡 <strong style={{color:"#60a5fa"}}>Click any site cell</strong> to edit inline · <strong style={{color:"#60a5fa"}}>Edit</strong> opens full profile · <strong style={{color:"#f87171"}}>📋 Profile</strong> exports worker card · <strong style={{color:"#34d399"}}>💷 Payslip</strong> in Payroll tab
        </div>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
            <thead><tr>
              <th style={{...TH,minWidth:155,position:"sticky",left:0,zIndex:2}}>Worker</th>
              <th style={TH}>Company</th><th style={TH}>Position</th>
              {activeDays.map(d=><th key={d} style={{...TH,minWidth:145,color:WEEKEND_DAYS.includes(d)?"#fbbf24":"#64748b"}}>{d}{WEEKEND_DAYS.includes(d)?" 🟡":""}</th>)}
              <th style={TH}>Rate</th><th style={TH}>Tax</th><th style={TH}>Certs</th><th style={TH}>Actions</th>
            </tr></thead>
            <tbody>
              {filtered.map((w,i)=>{
                const exp=CERTS.filter(c=>cSt(c,w)==="expired").length;
                const expg=CERTS.filter(c=>cSt(c,w)==="expiring").length;
                return <tr key={w.id} style={{background:i%2===0?"#111827":"#0f1421"}}>
                  <td style={{...TD,fontWeight:600,color:"#f1f5f9",position:"sticky",left:0,background:i%2===0?"#111827":"#0f1421",zIndex:1}}>
                    <div>{w.name}</div>
                    {w.comments&&<div style={{fontSize:10,color:"#fbbf24"}}>⚑ {w.comments}</div>}
                  </td>
                  <td style={{...TD,color:"#94a3b8",fontSize:11}}>{w.company||"—"}</td>
                  <td style={{...TD,color:"#94a3b8",fontSize:11}}>{w.position||"—"}</td>
                  {activeDays.map(d=>(
                    <td key={d} style={{...TD,background:WEEKEND_DAYS.includes(d)?"rgba(251,191,36,0.03)":undefined,padding:"4px 7px"}}>
                      <InlineCell value={w.days[d]} workerId={w.id} day={d} allSiteNames={allSiteNames} allSites={allSites} onUpdate={updateCell}/>
                    </td>
                  ))}
                  <td style={{...TD,color:"#34d399",fontWeight:600}}>{w.agreedRate?`£${w.agreedRate}`:<span style={{color:"#374151"}}>—</span>}</td>
                  <td style={TD}><span style={{fontSize:11,fontWeight:700,color:w.taxRate===0.30?"#f87171":w.taxRate===0.20?"#fbbf24":"#34d399"}}>{Math.round((w.taxRate||0)*100)}%</span></td>
                  <td style={TD}><div style={{display:"flex",gap:3}}>
                    {exp>0&&<span style={{color:"#f87171",fontSize:11,fontWeight:700}}>✗{exp}</span>}
                    {expg>0&&<span style={{color:"#fbbf24",fontSize:11,fontWeight:700}}>⚠{expg}</span>}
                    {exp===0&&expg===0&&<span style={{color:"#374151"}}>—</span>}
                  </div></td>
                  <td style={TD}>
                    <div style={{display:"flex",gap:4,flexWrap:"nowrap"}}>
                      <button onClick={()=>setModal({type:"worker",worker:w})} style={{padding:"4px 8px",background:"#1e3a5f",border:"1px solid #3b82f6",borderRadius:5,color:"#60a5fa",cursor:"pointer",fontSize:11,fontWeight:600}}>Edit</button>
                      <button onClick={()=>exportWorkerProfile(w,allSites,weekLabel)} title="Export worker profile PDF" style={{padding:"4px 8px",background:"#1a2535",border:"1px solid #8b5cf6",borderRadius:5,color:"#a78bfa",cursor:"pointer",fontSize:11,fontWeight:600}}>📋</button>
                      <button onClick={()=>exportPayslip(w,activeDays,weekLabel,siteHours)} title="Generate payslip PDF" style={{padding:"4px 8px",background:"#0d2218",border:"1px solid #10b981",borderRadius:5,color:"#34d399",cursor:"pointer",fontSize:11,fontWeight:600}}>💷</button>
                      <button onClick={()=>delWorker(w.id)} style={{padding:"4px 8px",background:"#2d1515",border:"1px solid #ef4444",borderRadius:5,color:"#f87171",cursor:"pointer",fontSize:11,fontWeight:600}}>✕</button>
                    </div>
                  </td>
                </tr>;
              })}
            </tbody>
          </table>
          {filtered.length===0&&<div style={{textAlign:"center",padding:50,color:"#374151"}}>No workers match filters.</div>}
        </div>
      </div>}
    </div>

    {/* Legend */}
    <div style={{padding:"10px 18px",borderTop:"1px solid #1e2535",background:"#0d1117"}}>
      <div style={{fontSize:10,color:"#374151",marginBottom:5,fontWeight:700,textTransform:"uppercase"}}>Site Legend</div>
      <div style={{display:"flex",flexWrap:"wrap",gap:4}}>
        {allSites.filter(s=>!isOff(s.name)).map(s=><span key={s.id} style={{display:"inline-flex",alignItems:"center",gap:3,padding:"2px 7px",borderRadius:20,background:"#111827",border:`1px solid ${s.color}`,fontSize:10,color:"#94a3b8"}}>
          <span style={{width:6,height:6,borderRadius:"50%",background:s.color}}/>{s.name}
        </span>)}
      </div>
    </div>

    {/* Modals */}
    {modal?.type==="worker"&&<WorkerModal worker={modal.worker} onSave={saveWorker} onClose={()=>setModal(null)} allSiteNames={allSiteNames} allSites={allSites} activeDays={activeDays}/>}
    {modal?.type==="sites"&&<SitesModal allSites={allSites} clients={clients} onSave={s=>{setAllSites(s);setModal(null);}} onClose={()=>setModal(null)}/>}
    {modal?.type==="clients"&&<ClientsModal clients={clients} onSave={l=>{setClients(l);setModal(null);}} onClose={()=>setModal(null)}/>}
  </div>;
}
