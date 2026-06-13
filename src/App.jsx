import { useState, useMemo, useRef, useEffect } from "react";
import * as XLSX from "xlsx";

// ─── Supabase ─────────────────────────────────────────────────────────────────
const SB_URL = "https://xljglqiifogyxefhszwa.supabase.co";
const SB_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InhsamdscWlpZm9neXhlZmhzendhIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODEwMTY2MTQsImV4cCI6MjA5NjU5MjYxNH0.asql85bUrgL5JuzqYoU0ZtizIWJ1yU6NYTt3yMUW5us";
const SB_H = { "Content-Type":"application/json","apikey":SB_KEY,"Authorization":`Bearer ${SB_KEY}` };
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
const PRESET_COLORS=[
  "#3b82f6","#6366f1","#8b5cf6","#a855f7","#d946ef","#ec4899","#f43f5e","#ef4444",
  "#f97316","#f59e0b","#eab308","#84cc16","#22c55e","#10b981","#14b8a6","#06b6d4",
  "#0ea5e9","#0284c7","#1d4ed8","#4f46e5","#7c3aed","#9333ea","#c026d3","#db2777",
  "#e11d48","#dc2626","#ea580c","#d97706","#ca8a04","#65a30d","#16a34a","#059669",
  "#0d9488","#0891b2","#0369a1","#1e40af","#374151","#6b7280","#94a3b8","#a78bfa",
  "#34d399","#60a5fa","#fbbf24","#f87171","#c084fc",
];

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
  {id:"b1",scopes:[],variations:[],name:"003 - STF",color:"#f59e0b",clientId:null,builtin:true},
  {id:"b2",scopes:[],variations:[],name:"0066 - UKTOP",color:"#3b82f6",clientId:null,builtin:true},
  {id:"b3",scopes:[],variations:[],name:"JAUK - 42 Station Road",color:"#8b5cf6",clientId:null,builtin:true},
  {id:"b4",scopes:[],variations:[],name:"JAUK - Pool Street",color:"#06b6d4",clientId:null,builtin:true},
  {id:"b5",scopes:[],variations:[],name:"JAUK - Tower 42",color:"#10b981",clientId:null,builtin:true},
  {id:"b6",scopes:[],variations:[],name:"BMW",color:"#ef4444",clientId:null,builtin:true},
  {id:"b7",scopes:[],variations:[],name:"SB - Camden",color:"#f97316",clientId:null,builtin:true},
  {id:"b8",scopes:[],variations:[],name:"DODI",color:"#ec4899",clientId:null,builtin:true},
  {id:"b9",scopes:[],variations:[],name:"SS",color:"#6366f1",clientId:null,builtin:true},
  {id:"b10",scopes:[],variations:[],name:"XX - OFF",color:"#6b7280",clientId:null,builtin:true},
  {id:"b11",scopes:[],variations:[],name:"X - Holiday",color:"#84cc16",clientId:null,builtin:true},
  {id:"b12",scopes:[],variations:[],name:"XX - Storage",color:"#a78bfa",clientId:null,builtin:true},
];

// ─── Team Types for Client Day Rates ─────────────────────────────────────────
const TEAM_TYPES=[
  {key:"welding",label:"Welding Team (2 Skilled + Semiskilled)"},
  {key:"erectors",label:"Erectors Team (2 Skilled + Semiskilled)"},
  {key:"architectural",label:"Architectural Team (2 Skilled + Semiskilled)"},
  {key:"cladding",label:"Cladding Team (2 Skilled + Semiskilled)"},
  {key:"plant_op",label:"Plant Operator"},
  {key:"supervisor",label:"Supervisor"},
  {key:"manager",label:"Manager"},
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
function emptyCerts(){return Object.fromEntries(CERTS.map(c=>[c.key,{held:false,expiry:"",regNo:"",fileUrl:""}]));}
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
  {id:"c1",name:"JAUK Ltd",email:"info@jauk.com",phone:"02012345678",color:"#8b5cf6",notes:"",rates:[]},
  {id:"c2",name:"STF Projects",email:"info@stf.com",phone:"02087654321",color:"#f59e0b",notes:"",rates:[]}
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


// ─── Color Picker ─────────────────────────────────────────────────────────────
function ColorPicker({value,onChange}){
  return <div style={{display:"flex",flexWrap:"wrap",gap:4,marginTop:4}}>
    {PRESET_COLORS.map(c=><div key={c} onClick={()=>onChange(c)}
      style={{width:22,height:22,borderRadius:4,background:c,cursor:"pointer",flexShrink:0,
        border:value===c?"3px solid #fff":"2px solid transparent",boxSizing:"border-box",
        boxShadow:value===c?"0 0 0 1px "+c:"none"}}
      title={c}/>)}
  </div>;
}

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
function SitesModal({allSites,clients,onSave,onClose,onOpenDetail}){
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
        <div style={{minWidth:260}}><label style={LBL}>Colour</label><ColorPicker value={nc} onChange={setNc}/></div>
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
            <button onClick={()=>{onSave(sites);onOpenDetail&&onOpenDetail(s);}} title="Open site detail (scopes & variations)" style={{padding:"4px 9px",background:"#1a3a5f",border:"1px solid #3b82f6",borderRadius:5,color:"#60a5fa",cursor:"pointer",fontSize:12}}>📂</button>
          </div>
          <div style={{marginBottom:7}}><ColorPicker value={s.color} onChange={(c)=>{up(s.id,"color",c);}}/></div>
          <select value={s.clientId||""} onChange={(e)=>{up(s.id,"clientId",e.target.value||null);}} style={{...INP,fontSize:12,padding:"4px 7px",cursor:"pointer"}}><option value="">No client</option>{clients.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}</select>
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
          <div style={{marginBottom:7}}><ColorPicker value={s.color} onChange={(c)=>{up(s.id,"color",c);}}/></div>
          <select value={s.clientId||""} onChange={(e)=>{up(s.id,"clientId",e.target.value||null);}} style={{...INP,fontSize:12,padding:"4px 7px",cursor:"pointer"}}><option value="">No client</option>{clients.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}</select>
        </div>)}
      </div>
    </div>}

    <div style={{display:"flex",gap:10,justifyContent:"flex-end",marginTop:20,borderTop:"1px solid #1e2535",paddingTop:16}}>
      <button onClick={onClose} style={{padding:"8px 18px",background:"#1e2535",border:"1px solid #2d3555",borderRadius:7,color:"#94a3b8",cursor:"pointer"}}>Cancel</button>
      <button onClick={()=>onSave(sites)} style={BG}>Save All Sites</button>
    </div>
  </Overlay>;
}

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
    <TabBar tabs={[["personal","👤 Personal"],["schedule","📅 Schedule"],["pay","💷 Pay & OT"],["certs","🛡 Certs "+(alerts>0?"⚠"+alerts:"("+held+")")]]} active={tab} onChange={setTab}/>
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
            {val.held&&<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"6px 10px",marginTop:7}}>
              <div><label style={{...LBL,marginBottom:2}}>Reg / Card No</label>
                <input value={val.regNo||""} onChange={e=>setC(cert.key,{...val,regNo:e.target.value})} placeholder="e.g. SK123456" style={{...INP,fontSize:11,padding:"4px 7px"}}/>
              </div>
              {cert.hasExpiry?<div><label style={{...LBL,marginBottom:2}}>Expiry Date</label>
                <input type="date" value={val.expiry||""} onChange={e=>setC(cert.key,{...val,expiry:e.target.value})} style={{...INP,fontSize:11,padding:"4px 7px"}}/>
              </div>:<div/>}
              <div style={{gridColumn:"1/-1"}}><label style={{...LBL,marginBottom:2}}>Certificate URL / File Link</label>
                <input value={val.fileUrl||""} onChange={e=>setC(cert.key,{...val,fileUrl:e.target.value})} placeholder="https://drive.google.com/… or any URL" style={{...INP,fontSize:11,padding:"4px 7px"}}/>
                {val.fileUrl&&<a href={val.fileUrl} target="_blank" rel="noreferrer" style={{fontSize:10,color:"#60a5fa",marginTop:3,display:"block"}}>📎 Open Certificate →</a>}
              </div>
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


// ─── Open-In-New-Window helpers (works for both interfaces) ─────────────────
function openWorkerWindow(w, allSites, weekLabel, activeDays, siteHours) {
  const {gross,net,stdH,otH,bd}=calcPay(w,activeDays||BASE_DAYS,siteHours||{});
  const held=Object.entries(w.certs||{}).filter(([,v])=>v.held).map(([k,v])=>({...v,key:k,label:CERTS.find(c=>c.key===k)?.label||k}));
  const html=`<!DOCTYPE html><html><head><meta charset="utf-8"/><title>Worker — ${w.name}</title>
<style>*{margin:0;padding:0;box-sizing:border-box;}body{background:#0d1117;color:#e2e8f0;font-family:'Segoe UI',Arial,sans-serif;font-size:13px;padding:24px;}
.hdr{display:flex;align-items:center;gap:14px;padding-bottom:16px;border-bottom:2px solid #1e2535;margin-bottom:20px;}
.av{width:52px;height:52px;background:linear-gradient(135deg,#3b82f6,#6366f1);border-radius:12px;display:flex;align-items:center;justify-content:center;font-size:22px;font-weight:800;flex-shrink:0;}
.grid{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:18px;}
.card{background:#1a1f2e;border:1px solid #2d3555;border-radius:10px;padding:14px;}
.cl{font-size:10px;color:#64748b;font-weight:700;text-transform:uppercase;letter-spacing:0.06em;margin-bottom:10px;}
.row{display:flex;gap:8px;margin-bottom:7px;} .rl{font-size:10px;color:#64748b;font-weight:600;min-width:90px;flex-shrink:0;text-transform:uppercase;} .rv{font-size:12px;}
.stat{background:#0f1421;border-radius:8px;padding:10px 12px;text-align:center;} .sl{font-size:9px;color:#64748b;text-transform:uppercase;} .sv{font-size:18px;font-weight:800;margin-top:3px;}
table{width:100%;border-collapse:collapse;margin-bottom:16px;} th{padding:7px 10px;text-align:left;font-size:10px;font-weight:700;color:#64748b;text-transform:uppercase;border-bottom:2px solid #1e2535;background:#0a0e17;}
td{padding:6px 10px;border-bottom:1px solid #1a2030;font-size:12px;} tr:nth-child(even) td{background:#111827;} tr:nth-child(odd) td{background:#0f1421;}
.cert{background:#1a1f2e;border-radius:8px;padding:9px 11px;border-left:3px solid #2d3555;margin-bottom:6px;}
.ft{margin-top:20px;padding-top:12px;border-top:1px solid #1e2535;display:flex;justify-content:space-between;font-size:10px;color:#374151;}
@media print{body{padding:12px;}@page{margin:8mm;size:A4;}}</style></head><body>
<div class="hdr"><div class="av">${(w.name||"?")[0]}</div><div>
  <div style="font-size:20px;font-weight:800;color:#f1f5f9">${w.name||"—"}</div>
  <div style="font-size:13px;color:#64748b">${w.position||"—"} · ${w.company||"—"}</div>
  ${w.comments?`<div style="font-size:11px;color:#fbbf24;margin-top:3px">⚑ ${w.comments}</div>`:""}
</div>
<div style="margin-left:auto;display:grid;grid-template-columns:repeat(4,1fr);gap:8px">
  ${[["Rate",w.agreedRate?`£${w.agreedRate}/hr`:"—","#34d399"],["Tax",Math.round((w.taxRate||0)*100)+"%","#fbbf24"],["Gross",`£${gross.toFixed(0)}`,"#60a5fa"],["Net",`£${net.toFixed(0)}`,"#a78bfa"]].map(([l,v,c])=>`<div class="stat"><div class="sl">${l}</div><div class="sv" style="color:${c}">${v}</div></div>`).join("")}
</div></div>
<div class="grid"><div class="card"><div class="cl">Contact</div>
${[["Phone",w.contact],["Email",w.email],["DOB",w.dob?fmtDate(w.dob):""],["Address",w.address],["NINO",w.nino],["UTR",w.utr]].filter(([,v])=>v).map(([l,v])=>`<div class="row"><span class="rl">${l}</span><span class="rv">${v}</span></div>`).join("")}
</div><div class="card"><div class="cl">Bank & Emergency</div>
${[["Bank",w.bankName],["Account",w.bankAccount],["Sort",w.bankSort],["NOK",w.nextOfKin],["NOK Phone",w.nextOfKinPhone]].filter(([,v])=>v).map(([l,v])=>`<div class="row"><span class="rl">${l}</span><span class="rv">${v}</span></div>`).join("")}
</div></div>
<div class="card" style="margin-bottom:16px"><div class="cl">Week ${weekLabel} — Daily Allocation</div>
<table><thead><tr><th>Day</th><th>Site</th><th>Std Hrs</th><th>OT Hrs</th><th>Day Pay</th></tr></thead><tbody>
${(activeDays||BASE_DAYS).map(d=>{const b=bd[d];const site=w.days?.[d];return `<tr><td style="font-weight:700;color:#94a3b8">${d}</td><td>${b?`<span style="color:#60a5fa">${b.site}</span>`:`<span style="color:#374151">${site||"—"}</span>`}</td><td>${b?b.hours+"h":"—"}</td><td>${b&&b.ot>0?b.ot+"h":"—"}</td><td>${b?`£${b.gross.toFixed(2)}`:"—"}</td></tr>`;}).join("")}
</tbody></table></div>
${held.length>0?`<div class="card" style="margin-bottom:16px"><div class="cl">Certificates (${held.length} held)</div>
${held.map(c=>{const exp=c.expiry?new Date(c.expiry):null;const days=exp?(exp-new Date())/86400000:null;const st=!exp?"valid":days<0?"expired":days<30?"expiring":"valid";const sc={valid:"#34d399",expiring:"#fbbf24",expired:"#f87171"}[st];return `<div class="cert" style="border-left-color:${sc}"><div style="font-weight:600;color:#e2e8f0;font-size:12px">${c.label}</div>${c.regNo?`<div style="font-size:11px;color:#60a5fa">Reg: ${c.regNo}</div>`:""}${c.expiry?`<div style="font-size:10px;color:#64748b">Expiry: ${c.expiry}</div>`:""}${c.fileUrl?`<div><a href="${c.fileUrl}" target="_blank" style="font-size:10px;color:#60a5fa">📎 View Certificate</a></div>`:""}<div style="font-size:10px;font-weight:700;color:${sc};text-transform:uppercase;margin-top:3px">${st}</div></div>`;}).join("")}
</div>`:""}
<div class="ft"><span>Worker Profile — ${w.name}</span><span>WC: ${weekLabel}</span><span>Bright Metalwork Ltd · Confidential</span><span>${new Date().toLocaleDateString("en-GB")}</span></div>
<script>window.onload=function(){window.print();}<\/script></body></html>`;
  const blob=new Blob([html],{type:"text/html"});const url=URL.createObjectURL(blob);
  const win=window.open(url,"_blank","width=960,height=820");
  if(!win){const a=document.createElement("a");a.href=url;a.download="Worker_"+w.name.replace(/\s+/g,"_")+".html";a.click();}
  setTimeout(()=>URL.revokeObjectURL(url),6000);
}

function openSiteWindow(site, clients, workers, activeDays, siteHours) {
  const client=clients.find(c=>c.id===site.clientId);
  const sc=(site.scopes||[]),vr=(site.variations||[]);
  const scopeT=sc.reduce((a,s)=>a+(s.qty*s.rate),0);
  const varT=vr.reduce((a,v)=>a+(v.type==="addition"?v.value:-v.value),0);
  let labourT=0;workers.forEach(w=>{const{bd}=calcPay(w,activeDays||BASE_DAYS,siteHours||{});Object.values(bd).forEach(b=>{if(b.site===site.name||b.site.includes(site.name))labourT+=b.gross;});});
  const siteWorkers=workers.filter(w=>(activeDays||BASE_DAYS).some(d=>(w.days?.[d]||"").includes(site.name)));
  const html=`<!DOCTYPE html><html><head><meta charset="utf-8"/><title>Site — ${site.name}</title>
<style>*{margin:0;padding:0;box-sizing:border-box;}body{background:#0d1117;color:#e2e8f0;font-family:'Segoe UI',Arial,sans-serif;font-size:13px;padding:24px;}
.hdr{padding-bottom:16px;border-bottom:2px solid ${site.color};margin-bottom:20px;}
.stat{background:#1a1f2e;border-radius:9px;padding:11px 14px;} .sl{font-size:9px;color:#64748b;text-transform:uppercase;font-weight:700;} .sv{font-size:20px;font-weight:800;margin-top:3px;}
table{width:100%;border-collapse:collapse;margin-bottom:16px;} th{padding:7px 10px;text-align:left;font-size:10px;font-weight:700;color:#64748b;text-transform:uppercase;border-bottom:2px solid #1e2535;background:#0a0e17;}
td{padding:7px 10px;border-bottom:1px solid #1a2030;font-size:12px;} tr:nth-child(even) td{background:#111827;} tr:nth-child(odd) td{background:#0f1421;}
.section{margin-bottom:20px;} .sh{font-size:11px;color:#64748b;font-weight:700;text-transform:uppercase;margin-bottom:10px;}
.ft{margin-top:20px;padding-top:12px;border-top:1px solid #1e2535;display:flex;justify-content:space-between;font-size:10px;color:#374151;}
@media print{@page{margin:8mm;size:A3 landscape;}}</style></head><body>
<div class="hdr">
  <div style="display:flex;align-items:center;gap:10;margin-bottom:12px">
    <span style="width:14px;height:14px;border-radius:50%;background:${site.color};display:inline-block;margin-right:6px"></span>
    <span style="font-size:22px;font-weight:800;color:#f1f5f9">${site.name}</span>
    ${client?`<span style="margin-left:12px;font-size:13px;color:${client.color};font-weight:600">${client.name}</span>`:""}
  </div>
  <div style="display:grid;grid-template-columns:repeat(5,1fr);gap:10px">
    ${[["Contract","£"+(scopeT+varT).toLocaleString(),"#60a5fa"],["Scope Value","£"+scopeT.toLocaleString(),"#34d399"],["Variations","£"+Math.abs(varT).toLocaleString(),"#fbbf24"],["Labour (wk)","£"+labourT.toFixed(0),"#f87171"],["Workers",siteWorkers.length,"#a78bfa"]].map(([l,v,c])=>`<div class="stat"><div class="sl">${l}</div><div class="sv" style="color:${c}">${v}</div></div>`).join("")}
  </div>
</div>
<div class="section"><div class="sh">Scopes of Work (${sc.length})</div>
<table><thead><tr><th>Description</th><th>Unit</th><th>Qty</th><th>Rate £</th><th>Total £</th></tr></thead><tbody>
${sc.map(s=>`<tr><td>${s.description||s.desc||""}</td><td>${s.unit||""}</td><td style="text-align:right;color:#60a5fa">${s.qty}</td><td style="text-align:right">£${(s.rate||s.unitIncome||0).toLocaleString()}</td><td style="text-align:right;color:#34d399;font-weight:700">£${((s.qty||0)*(s.rate||s.unitIncome||0)).toLocaleString()}</td></tr>`).join("")}
${sc.length===0?"<tr><td colspan='5' style='text-align:center;color:#374151'>No scopes defined</td></tr>":""}
</tbody></table></div>
<div class="section"><div class="sh">Variations (${vr.length})</div>
<table><thead><tr><th>Description</th><th>Type</th><th>Value £</th><th>Status</th></tr></thead><tbody>
${vr.map(v=>`<tr><td>${v.description||v.desc||""}</td><td style="color:${v.type==="addition"?"#34d399":"#f87171"}">${v.type}</td><td style="color:${v.type==="addition"?"#34d399":"#f87171"};font-weight:700">${v.type==="addition"?"+":"-"}£${(v.value||0).toLocaleString()}</td><td style="color:${v.approved?"#34d399":"#fbbf24"}">${v.approved?"Approved":"Pending"}</td></tr>`).join("")}
${vr.length===0?"<tr><td colspan='4' style='text-align:center;color:#374151'>No variations</td></tr>":""}
</tbody></table></div>
<div class="section"><div class="sh">Workers on Site (${siteWorkers.length})</div>
<table><thead><tr><th>Name</th><th>Position</th><th>Rate</th><th>Tax</th></tr></thead><tbody>
${siteWorkers.map(w=>`<tr><td style="font-weight:600">${w.name}</td><td>${w.position||"—"}</td><td style="color:#34d399">${w.agreedRate?"£"+w.agreedRate+"/hr":"—"}</td><td>${Math.round((w.taxRate||0)*100)}%</td></tr>`).join("")}
${siteWorkers.length===0?"<tr><td colspan='4' style='text-align:center;color:#374151'>No workers allocated</td></tr>":""}
</tbody></table></div>
<div class="ft"><span>Site Report — ${site.name}</span><span>${client?client.name:""}</span><span>Bright Metalwork Ltd</span><span>${new Date().toLocaleDateString("en-GB")}</span></div>
<script>window.onload=function(){window.print();}<\/script></body></html>`;
  const blob=new Blob([html],{type:"text/html"});const url=URL.createObjectURL(blob);
  const win=window.open(url,"_blank","width=1100,height=820");
  if(!win){const a=document.createElement("a");a.href=url;a.download="Site_"+site.name.replace(/\s+/g,"_")+".html";a.click();}
  setTimeout(()=>URL.revokeObjectURL(url),6000);
}

function openClientWindow(client, allSites, invoices, workers, activeDays, siteHours) {
  const sites=allSites.filter(s=>s.clientId===client.id);
  const invs=(invoices||[]).filter(i=>sites.find(s=>s.id===i.siteId));
  const totalInv=invs.reduce((a,i)=>a+(i.amount||0),0);
  const paid=invs.filter(i=>i.status==="paid").reduce((a,i)=>a+i.amount,0);
  let labourT=0;workers.forEach(w=>{const{bd}=calcPay(w,activeDays||BASE_DAYS,siteHours||{});Object.values(bd).forEach(b=>{if(sites.find(s=>b.site===s.name||b.site.includes(s.name)))labourT+=b.gross;});});
  const html=`<!DOCTYPE html><html><head><meta charset="utf-8"/><title>Client — ${client.name}</title>
<style>*{margin:0;padding:0;box-sizing:border-box;}body{background:#0d1117;color:#e2e8f0;font-family:'Segoe UI',Arial,sans-serif;font-size:13px;padding:24px;}
.hdr{padding-bottom:16px;border-bottom:2px solid ${client.color};margin-bottom:20px;display:flex;align-items:center;gap:14px;}
.av{width:52px;height:52px;border-radius:12px;background:${client.color}22;border:1px solid ${client.color}44;display:flex;align-items:center;justify-content:center;font-size:22px;font-weight:800;color:${client.color};}
.stat{background:#1a1f2e;border-radius:9px;padding:11px 14px;} .sl{font-size:9px;color:#64748b;text-transform:uppercase;font-weight:700;} .sv{font-size:20px;font-weight:800;margin-top:3px;}
table{width:100%;border-collapse:collapse;margin-bottom:16px;} th{padding:7px 10px;text-align:left;font-size:10px;font-weight:700;color:#64748b;text-transform:uppercase;border-bottom:2px solid #1e2535;background:#0a0e17;}
td{padding:7px 10px;border-bottom:1px solid #1a2030;font-size:12px;} tr:nth-child(even) td{background:#111827;} tr:nth-child(odd) td{background:#0f1421;}
.sh{font-size:11px;color:#64748b;font-weight:700;text-transform:uppercase;margin-bottom:10px;}
.ft{margin-top:20px;padding-top:12px;border-top:1px solid #1e2535;display:flex;justify-content:space-between;font-size:10px;color:#374151;}
@media print{@page{margin:8mm;size:A4;}}</style></head><body>
<div class="hdr"><div class="av">${client.name[0]}</div><div>
  <div style="font-size:22px;font-weight:800;color:#f1f5f9">${client.name}</div>
  <div style="font-size:13px;color:#64748b">${client.email||""} · ${client.phone||""}</div>
  ${client.notes?`<div style="font-size:11px;color:#94a3b8;margin-top:3px">${client.notes}</div>`:""}
</div>
<div style="margin-left:auto;display:grid;grid-template-columns:repeat(4,1fr);gap:8px">
  ${[["Sites",sites.length,"#60a5fa"],["Labour","£"+labourT.toFixed(0),"#f87171"],["Invoiced","£"+totalInv.toLocaleString(),"#34d399"],["Collected","£"+paid.toLocaleString(),"#a78bfa"]].map(([l,v,c])=>`<div class="stat"><div class="sl">${l}</div><div class="sv" style="color:${c}">${v}</div></div>`).join("")}
</div></div>
<div style="margin-bottom:20px"><div class="sh">Agreed Day Rates</div>
<table><thead><tr><th>Team / Role Type</th><th>Day Rate</th><th>Notes</th></tr></thead><tbody>
${(client.rates||[]).map(r=>`<tr><td style="font-weight:600">${TEAM_TYPES.find(t=>t.key===r.teamType)?.label||r.teamType||""}</td><td style="color:#34d399;font-weight:700">£${r.dayRate||r.dayrate||0}/day</td><td style="color:#64748b">${r.notes||r.description||""}</td></tr>`).join("")}
${(client.rates||[]).length===0?"<tr><td colspan='3' style='text-align:center;color:#374151'>No rates configured</td></tr>":""}
</tbody></table></div>
<div style="margin-bottom:20px"><div class="sh">Sites (${sites.length})</div>
<table><thead><tr><th>Site</th><th>Scopes</th><th>Variations</th><th>Contract Value</th></tr></thead><tbody>
${sites.map(s=>{const sc=(s.scopes||[]).reduce((a,x)=>a+(x.qty||0)*(x.rate||x.unitIncome||0),0);const vt=(s.variations||[]).reduce((a,v)=>a+(v.type==="addition"?v.value:-v.value),0);return `<tr><td style="font-weight:600;color:${s.color}">${s.name}</td><td>${(s.scopes||[]).length}</td><td>${(s.variations||[]).length}</td><td style="color:#34d399;font-weight:700">£${(sc+vt).toLocaleString()}</td></tr>`;}).join("")}
</tbody></table></div>
<div class="sh">Invoices (${invs.length})</div>
<table><thead><tr><th>Invoice</th><th>Date</th><th>Amount</th><th>Status</th></tr></thead><tbody>
${invs.map(i=>{const sc={paid:"#34d399",pending:"#fbbf24",overdue:"#f87171",draft:"#64748b"}[i.status]||"#64748b";return `<tr><td style="font-weight:600;color:#60a5fa">${i.number||i.num||"—"}</td><td>${i.date||""}</td><td style="font-weight:700;color:#34d399">£${(i.amount||i.total||0).toLocaleString()}</td><td style="color:${sc};font-weight:700;text-transform:capitalize">${i.status||""}</td></tr>`;}).join("")}
${invs.length===0?"<tr><td colspan='4' style='text-align:center;color:#374151'>No invoices</td></tr>":""}
</tbody></table>
<div class="ft"><span>Client Report — ${client.name}</span><span>Bright Metalwork Ltd</span><span>${new Date().toLocaleDateString("en-GB")}</span></div>
<script>window.onload=function(){window.print();}<\/script></body></html>`;
  const blob=new Blob([html],{type:"text/html"});const url=URL.createObjectURL(blob);
  const win=window.open(url,"_blank","width=960,height=820");
  if(!win){const a=document.createElement("a");a.href=url;a.download="Client_"+client.name.replace(/\s+/g,"_")+".html";a.click();}
  setTimeout(()=>URL.revokeObjectURL(url),6000);
}

function openInvoiceWindow(inv, client, site) {
  const {subtotal,vat,total}=calcInvoiceTotals(inv);
  const html=`<!DOCTYPE html><html><head><meta charset="utf-8"/><title>Invoice ${inv.number||inv.num}</title>
<style>*{margin:0;padding:0;box-sizing:border-box;}body{background:#0d1117;color:#e2e8f0;font-family:'Segoe UI',Arial,sans-serif;font-size:13px;padding:24px;}
.hdr{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:20px;padding-bottom:16px;border-bottom:2px solid #1e2535;}
table{width:100%;border-collapse:collapse;margin-bottom:16px;} th{padding:8px 10px;text-align:left;font-size:10px;font-weight:700;color:#64748b;text-transform:uppercase;border-bottom:2px solid #1e2535;background:#0a0e17;}
td{padding:8px 10px;border-bottom:1px solid #1a2030;font-size:12px;} tr:nth-child(even) td{background:#111827;} tr:nth-child(odd) td{background:#0f1421;}
.total{background:#0d2218;border:2px solid #10b981;border-radius:10px;padding:16px 20px;text-align:right;margin-top:10px;}
.ft{margin-top:20px;padding-top:12px;border-top:1px solid #1e2535;display:flex;justify-content:space-between;font-size:10px;color:#374151;}
@media print{@page{margin:10mm;size:A4;}}</style></head><body>
<div class="hdr">
  <div><div style="font-size:24px;font-weight:900;color:#f1f5f9">${inv.number||inv.num||"INVOICE"}</div>
  <div style="font-size:12px;color:#64748b;margin-top:4px">Date: ${inv.date||"—"} · ${site?site.name:""}</div>
  <div style="margin-top:6px;padding:3px 9px;display:inline-block;border-radius:5px;font-size:11px;font-weight:700;text-transform:capitalize;color:${inv.status==="paid"?"#34d399":inv.status==="pending"?"#fbbf24":"#f87171"};background:${inv.status==="paid"?"#0d2218":inv.status==="pending"?"#1a1500":"#2d1515"}">${inv.status||"draft"}</div>
  </div>
  <div style="text-align:right"><div style="font-size:13px;font-weight:700;color:#f1f5f9">${client?client.name:"—"}</div><div style="font-size:11px;color:#64748b">${client?.email||""}</div></div>
</div>
<table><thead><tr><th style="width:50%">Description</th><th>Qty</th><th>Unit Price</th><th>VAT</th><th>Amount</th></tr></thead><tbody>
${(inv.lineItems||inv.items||[]).map(li=>`<tr><td>${li.description||li.desc||""}</td><td style="text-align:right">${li.qty||1}</td><td style="text-align:right">£${(li.unitPrice||li.rate||0).toFixed(2)}</td><td style="text-align:right;color:#64748b">${li.vatRate||0}%</td><td style="text-align:right;font-weight:700;color:#34d399">£${(li.lineTotal||li.amount||0).toFixed(2)}</td></tr>`).join("")}
</tbody></table>
<div class="total">
  <div style="display:flex;justify-content:flex-end;gap:20px;margin-bottom:6px"><span style="color:#94a3b8">Subtotal</span><span>£${subtotal.toFixed(2)}</span></div>
  ${vat>0?`<div style="display:flex;justify-content:flex-end;gap:20px;margin-bottom:6px"><span style="color:#94a3b8">VAT</span><span style="color:#fbbf24">£${vat.toFixed(2)}</span></div>`:""}
  <div style="display:flex;justify-content:flex-end;gap:20px"><span style="font-size:14px;font-weight:700;color:#e2e8f0">TOTAL</span><span style="font-size:22px;font-weight:900;color:#34d399">£${total.toFixed(2)}</span></div>
</div>
<div class="ft"><span>Invoice ${inv.number||inv.num} — ${client?client.name:""}</span><span>Bright Metalwork Ltd</span><span>${new Date().toLocaleDateString("en-GB")}</span></div>
<script>window.onload=function(){window.print();}<\/script></body></html>`;
  const blob=new Blob([html],{type:"text/html"});const url=URL.createObjectURL(blob);
  const win=window.open(url,"_blank","width=900,height=780");
  if(!win){const a=document.createElement("a");a.href=url;a.download="Invoice_"+(inv.number||inv.num||"").replace(/\//g,"-")+".html";a.click();}
  setTimeout(()=>URL.revokeObjectURL(url),6000);
}


// ─── Site Detail Modal ────────────────────────────────────────────────────────
function SiteDetailModal({site,clients,workers,activeDays,siteHours,onSave,onClose}){
  const [s,setS]=useState({...site,scopes:[...(site.scopes||[])],variations:[...(site.variations||[])]});
  const [tab,setTab]=useState("scopes");
  const uid=()=>Date.now().toString(36)+Math.random().toString(36).slice(2);

  const addScope=()=>setS(x=>({...x,scopes:[...x.scopes,{id:uid(),description:"",unit:"",qty:0,rate:0}]}));
  const updScope=(id,k,v)=>setS(x=>({...x,scopes:x.scopes.map(sc=>sc.id===id?{...sc,[k]:v}:sc)}));
  const delScope=id=>setS(x=>({...x,scopes:x.scopes.filter(sc=>sc.id!==id)}));

  const addVar=()=>setS(x=>({...x,variations:[...x.variations,{id:uid(),description:"",value:0,type:"addition",approved:false}]}));
  const updVar=(id,k,v)=>setS(x=>({...x,variations:x.variations.map(vr=>vr.id===id?{...vr,[k]:v}:vr)}));
  const delVar=id=>setS(x=>({...x,variations:x.variations.filter(vr=>vr.id!==id)}));

  const labourCost=useMemo(()=>{let t=0;workers.forEach(w=>{const{bd}=calcPay(w,activeDays,siteHours);Object.values(bd).forEach(b=>{if(b.site===site.name||b.site.toUpperCase().includes(site.name.toUpperCase()))t+=b.gross;});});return t;},[workers,activeDays,siteHours,site.name]);
  const scopeTotal=s.scopes.reduce((a,sc)=>a+(Number(sc.qty||0)*Number(sc.rate||0)),0);
  const varTotal=s.variations.reduce((a,vr)=>a+(vr.type==="addition"?Number(vr.value||0):-Number(vr.value||0)),0);
  const contractValue=scopeTotal+varTotal;
  const profit=contractValue-labourCost;

  return <Overlay onClose={onClose} wide>
    <MH title={<span style={{display:"flex",alignItems:"center",gap:10}}>
      <span style={{width:14,height:14,borderRadius:"50%",background:s.color,display:"inline-block"}}/>{s.name} — Site Detail
    </span>} onClose={onClose}/>

    {/* Summary cards */}
    <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10,marginBottom:18}}>
      {[["Contract Value",`£${contractValue.toFixed(2)}`,"#60a5fa"],["Labour Cost",`£${labourCost.toFixed(2)}`,"#f87171"],["Variations",`${varTotal>=0?"+":""}£${varTotal.toFixed(2)}`,"#fbbf24"],["Profit / Loss",`£${profit.toFixed(2)}`,profit>=0?"#34d399":"#f87171"]].map(([l,v,c])=>(
        <div key={l} style={{background:"#0f1421",border:`1px solid ${c}33`,borderRadius:10,padding:"11px 14px"}}>
          <div style={{fontSize:10,color:"#64748b",fontWeight:700,textTransform:"uppercase"}}>{l}</div>
          <div style={{fontSize:18,fontWeight:800,color:c,marginTop:3}}>{v}</div>
        </div>
      ))}
    </div>

    <TabBar tabs={[["scopes","📋 Scopes of Work"],["variations","⚡ Variations"],["workers","👷 Workers"],["costs","💷 Cost Breakdown"]]} active={tab} onChange={setTab}/>

    {tab==="scopes"&&<div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
        <div style={{fontSize:13,color:"#94a3b8"}}>Agreed scope line items for this site</div>
        <button onClick={addScope} style={{...BP,padding:"6px 14px",fontSize:12}}>+ Add Scope Item</button>
      </div>
      {s.scopes.length===0&&<div style={{textAlign:"center",padding:32,color:"#374151",fontSize:13,border:"1px dashed #1e2535",borderRadius:8}}>No scope items yet. Click "+ Add Scope Item" to start.</div>}
      {s.scopes.map((sc,i)=>(
        <div key={sc.id} style={{display:"grid",gridTemplateColumns:"3fr 80px 80px 90px 90px 40px",gap:8,alignItems:"flex-end",padding:"10px 12px",background:i%2===0?"#0f1421":"#111827",borderRadius:8,marginBottom:6}}>
          <div><label style={LBL}>Description</label><input value={sc.description} onChange={e=>updScope(sc.id,"description",e.target.value)} placeholder="Scope item description…" style={INP}/></div>
          <div><label style={LBL}>Unit</label><input value={sc.unit} onChange={e=>updScope(sc.id,"unit",e.target.value)} placeholder="m², nr…" style={INP}/></div>
          <div><label style={LBL}>Qty</label><input type="number" value={sc.qty} onChange={e=>updScope(sc.id,"qty",e.target.value)} style={{...INP,textAlign:"right"}}/></div>
          <div><label style={LBL}>Rate £</label><input type="number" value={sc.rate} onChange={e=>updScope(sc.id,"rate",e.target.value)} style={{...INP,textAlign:"right"}}/></div>
          <div><label style={LBL}>Total</label><div style={{...INP,background:"#1a1f2e",color:"#34d399",fontWeight:700,textAlign:"right",padding:"7px 9px"}}>£{(Number(sc.qty||0)*Number(sc.rate||0)).toFixed(2)}</div></div>
          <button onClick={()=>delScope(sc.id)} style={{padding:"6px 10px",background:"#2d1515",border:"1px solid #ef4444",borderRadius:5,color:"#f87171",cursor:"pointer",fontSize:12,fontWeight:700,alignSelf:"flex-end"}}>✕</button>
        </div>
      ))}
      {s.scopes.length>0&&<div style={{display:"flex",justifyContent:"flex-end",marginTop:10,padding:"10px 14px",background:"#0d2218",border:"1px solid #065f46",borderRadius:8}}>
        <span style={{color:"#94a3b8",marginRight:14}}>Scope Total:</span>
        <span style={{color:"#34d399",fontSize:17,fontWeight:800}}>£{scopeTotal.toFixed(2)}</span>
      </div>}
    </div>}

    {tab==="variations"&&<div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
        <div style={{fontSize:13,color:"#94a3b8"}}>Variations and change orders</div>
        <button onClick={addVar} style={{...BP,padding:"6px 14px",fontSize:12}}>+ Add Variation</button>
      </div>
      {s.variations.length===0&&<div style={{textAlign:"center",padding:32,color:"#374151",fontSize:13,border:"1px dashed #1e2535",borderRadius:8}}>No variations yet.</div>}
      {s.variations.map((vr,i)=>(
        <div key={vr.id} style={{display:"grid",gridTemplateColumns:"3fr 110px 110px 110px 40px",gap:8,alignItems:"flex-end",padding:"10px 12px",background:i%2===0?"#0f1421":"#111827",borderRadius:8,marginBottom:6}}>
          <div><label style={LBL}>Description</label><input value={vr.description} onChange={e=>updVar(vr.id,"description",e.target.value)} placeholder="Variation description…" style={INP}/></div>
          <div><label style={LBL}>Type</label>
            <select value={vr.type} onChange={e=>updVar(vr.id,"type",e.target.value)} style={{...INP,cursor:"pointer"}}>
              <option value="addition">Addition (+)</option>
              <option value="omission">Omission (−)</option>
            </select>
          </div>
          <div><label style={LBL}>Value £</label><input type="number" value={vr.value} onChange={e=>updVar(vr.id,"value",e.target.value)} style={{...INP,textAlign:"right"}}/></div>
          <div><label style={LBL}>Approved</label>
            <select value={vr.approved?"yes":"no"} onChange={e=>updVar(vr.id,"approved",e.target.value==="yes")} style={{...INP,cursor:"pointer",color:vr.approved?"#34d399":"#fbbf24"}}>
              <option value="no">⏳ Pending</option><option value="yes">✓ Approved</option>
            </select>
          </div>
          <button onClick={()=>delVar(vr.id)} style={{padding:"6px 10px",background:"#2d1515",border:"1px solid #ef4444",borderRadius:5,color:"#f87171",cursor:"pointer",fontSize:12,fontWeight:700,alignSelf:"flex-end"}}>✕</button>
        </div>
      ))}
      {s.variations.length>0&&<div style={{display:"flex",justifyContent:"flex-end",marginTop:10,padding:"10px 14px",background:"#0d2218",border:"1px solid #065f46",borderRadius:8}}>
        <span style={{color:"#94a3b8",marginRight:14}}>Variations Net:</span>
        <span style={{color:varTotal>=0?"#34d399":"#f87171",fontSize:17,fontWeight:800}}>{varTotal>=0?"+":""}£{varTotal.toFixed(2)}</span>
      </div>}
    </div>}

    {tab==="workers"&&<div style={{overflowX:"auto"}}>
      <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
        <thead><tr><th style={TH}>Worker</th><th style={TH}>Position</th>{activeDays.map(d=><th key={d} style={TH}>{d}</th>)}<th style={TH}>Hrs</th><th style={TH}>Cost</th></tr></thead>
        <tbody>
          {workers.filter(w=>activeDays.some(d=>(w.days[d]||"").trim()===site.name||(w.days[d]||"").toUpperCase().includes(site.name.toUpperCase()))).map((w,i)=>{
            const{bd}=calcPay(w,activeDays,siteHours);
            const onSiteDays=activeDays.filter(d=>bd[d]?.site===site.name||bd[d]?.site?.toUpperCase().includes(site.name.toUpperCase()));
            const hrs=onSiteDays.reduce((a,d)=>a+(bd[d]?.hours||0),0);
            const cost=onSiteDays.reduce((a,d)=>a+(bd[d]?.gross||0),0);
            return <tr key={w.id} style={{background:i%2===0?"#111827":"#0f1421"}}>
              <td style={{...TD,fontWeight:600,color:"#f1f5f9"}}>{w.name}</td>
              <td style={{...TD,color:"#94a3b8",fontSize:11}}>{w.position}</td>
              {activeDays.map(d=><td key={d} style={TD}>{bd[d]&&(bd[d].site===site.name||bd[d].site?.toUpperCase().includes(site.name.toUpperCase()))?<span style={{color:"#34d399",fontSize:11}}>✓ {bd[d].hours}h</span>:<span style={{color:"#374151"}}>—</span>}</td>)}
              <td style={{...TD,color:"#60a5fa",fontWeight:700}}>{hrs}h</td>
              <td style={{...TD,color:"#f87171",fontWeight:700}}>£{cost.toFixed(2)}</td>
            </tr>;
          })}
        </tbody>
      </table>
      {workers.filter(w=>activeDays.some(d=>(w.days[d]||"").includes(site.name))).length===0&&<div style={{textAlign:"center",padding:28,color:"#374151"}}>No workers allocated to this site this week.</div>}
    </div>}

    {tab==="costs"&&<div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
        <div style={{background:"#0f1421",borderRadius:10,padding:16,border:"1px solid #1e2535"}}>
          <div style={{fontSize:11,color:"#64748b",fontWeight:700,textTransform:"uppercase",marginBottom:12}}>Income</div>
          <div style={{display:"flex",justifyContent:"space-between",padding:"8px 0",borderBottom:"1px solid #1e2535"}}><span style={{color:"#94a3b8"}}>Agreed Scope</span><span style={{fontWeight:700,color:"#60a5fa"}}>£{scopeTotal.toFixed(2)}</span></div>
          <div style={{display:"flex",justifyContent:"space-between",padding:"8px 0",borderBottom:"1px solid #1e2535"}}><span style={{color:"#94a3b8"}}>Variations</span><span style={{fontWeight:700,color:varTotal>=0?"#34d399":"#f87171"}}>{varTotal>=0?"+":""}£{varTotal.toFixed(2)}</span></div>
          <div style={{display:"flex",justifyContent:"space-between",padding:"10px 0",marginTop:4}}><span style={{color:"#e2e8f0",fontWeight:700}}>Total Contract Value</span><span style={{fontWeight:800,color:"#34d399",fontSize:16}}>£{contractValue.toFixed(2)}</span></div>
        </div>
        <div style={{background:"#0f1421",borderRadius:10,padding:16,border:"1px solid #1e2535"}}>
          <div style={{fontSize:11,color:"#64748b",fontWeight:700,textTransform:"uppercase",marginBottom:12}}>Internal Costs</div>
          <div style={{display:"flex",justifyContent:"space-between",padding:"8px 0",borderBottom:"1px solid #1e2535"}}><span style={{color:"#94a3b8"}}>Labour (this week)</span><span style={{fontWeight:700,color:"#f87171"}}>£{labourCost.toFixed(2)}</span></div>
          <div style={{display:"flex",justifyContent:"space-between",padding:"10px 0",marginTop:4}}>
            <span style={{color:"#e2e8f0",fontWeight:700}}>{profit>=0?"Profit":"Loss"}</span>
            <span style={{fontWeight:800,color:profit>=0?"#34d399":"#f87171",fontSize:16}}>£{Math.abs(profit).toFixed(2)}</span>
          </div>
        </div>
      </div>
    </div>}

    <div style={{display:"flex",gap:10,justifyContent:"flex-end",marginTop:20,borderTop:"1px solid #1e2535",paddingTop:16}}>
      <button onClick={onClose} style={{padding:"8px 18px",background:"#1e2535",border:"1px solid #2d3555",borderRadius:7,color:"#94a3b8",cursor:"pointer"}}>Cancel</button>
      <button onClick={()=>onSave(s)} style={BG}>Save Site</button>
    </div>
  </Overlay>;
}

// ─── Bank Import Modal ─────────────────────────────────────────────────────────
function BankImportModal({allSites,clients,onClose}){
  const [txns,setTxns]=useState([]);
  const [fileName,setFileName]=useState("");
  const INCOME_CATS=["Client Payment","Contract Payment","Variation Payment","Retention Release","Other Income"];
  const EXPENSE_CATS=["Materials","Plant Hire","Subcontractor","Labour (External)","Transport","Insurance","Tools & Equipment","Professional Fees","Utilities","Office","Other Expense"];

  const handleFile=e=>{
    const f=e.target.files[0];if(!f)return;setFileName(f.name);
    const reader=new FileReader();
    reader.onload=ev=>{
      try{
        const wb=XLSX.read(ev.target.result,{type:"binary",cellDates:false});
        const ws=wb.Sheets[wb.SheetNames[0]];
        const data=XLSX.utils.sheet_to_json(ws,{header:1,raw:true});
        const startRow=data.findIndex(r=>r.some(c=>c!==undefined&&c!==""));
        const header=(data[startRow]||[]).map(h=>String(h||"").toLowerCase().trim());
        const dataRows=data.slice(startRow+1).filter(r=>r.some(c=>c!==undefined&&c!==""));
        let dateCol=-1,descCol=-1,amtCol=-1,creditCol=-1,debitCol=-1;
        header.forEach((h,i)=>{
          if(h.includes("date")) dateCol=i;
          if(h.includes("desc")||h.includes("narr")||h.includes("detail")||h.includes("memo")||h.includes("ref")) descCol=i;
          if(h.includes("amount")&&!h.includes("credit")&&!h.includes("debit")) amtCol=i;
          if(h.includes("credit")) creditCol=i;
          if(h.includes("debit")) debitCol=i;
        });
        if(dateCol===-1) dateCol=0;
        if(descCol===-1) descCol=1;
        if(amtCol===-1&&creditCol===-1) amtCol=2;
        const toDate=v=>{
          if(!v&&v!==0) return "";
          if(typeof v==="string"&&(v.includes("-")||v.includes("/"))) return v;
          if(typeof v==="number"&&v>1000&&v<100000){const d=new Date(Math.round((v-25569)*86400*1000));return d.toLocaleDateString("en-GB",{day:"2-digit",month:"2-digit",year:"numeric"});}
          return String(v||"");
        };
        setTxns(dataRows.map(r=>{
          const desc=String(r[descCol]||"").trim();
          let amount=0;
          if(creditCol>-1||debitCol>-1){
            const cr=parseFloat(String(r[creditCol]||"0").replace(/[£,]/g,""))||0;
            const dr=parseFloat(String(r[debitCol]||"0").replace(/[£,]/g,""))||0;
            amount=cr>0?cr:-dr;
          } else {
            amount=parseFloat(String(r[amtCol]||"0").replace(/[£,]/g,""))||0;
          }
          return {id:Date.now().toString(36)+Math.random().toString(36).slice(2),date:toDate(r[dateCol]),description:desc,amount,type:amount>=0?"income":"expense",category:"",siteId:"",clientId:"",notes:""};
        }).filter(t=>t.description||t.amount!==0));
      }catch(err){alert("Error reading file: "+err.message);}
    };
    reader.readAsBinaryString(f);
  };

  const upT=(id,k,v)=>setTxns(t=>t.map(x=>x.id===id?{...x,[k]:v}:x));
  const income=txns.filter(t=>t.type==="income").reduce((a,t)=>a+Math.abs(t.amount),0);
  const expense=txns.filter(t=>t.type==="expense").reduce((a,t)=>a+Math.abs(t.amount),0);
  const categorised=txns.filter(t=>t.category).length;

  const exportCat=()=>{
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([
      ["Date","Description","Amount","Type","Category","Site","Client","Notes"],
      ...txns.map(t=>[t.date,t.description,t.amount,t.type,t.category,
        allSites.find(s=>s.id===t.siteId)?.name||"",
        clients.find(c=>c.id===t.clientId)?.name||"",t.notes])
    ]),"Categorised Transactions");
    XLSX.writeFile(wb,"Bank_Categorised_"+new Date().toLocaleDateString("en-GB").replace(/\//g,"-")+".xlsx");
  };

  return <Overlay onClose={onClose} wide>
    <MH title="🏦 Bank Import & Transaction Categorisation" onClose={onClose}/>
    <div style={{display:"grid",gridTemplateColumns:"1fr 2fr",gap:14,marginBottom:18}}>
      <div style={{background:"#0f1421",border:"1px solid #1e2535",borderRadius:10,padding:14}}>
        <div style={{fontSize:11,color:"#64748b",fontWeight:700,textTransform:"uppercase",marginBottom:8}}>Import Bank Statement</div>
        <p style={{fontSize:12,color:"#64748b",marginBottom:12,lineHeight:1.5}}>Upload your bank export Excel or CSV.<br/><strong style={{color:"#94a3b8"}}>Expected columns:</strong><br/>Date · Description · Amount</p>
        <label style={{display:"block",padding:"12px 14px",background:"#1e3a5f",border:"2px dashed #3b82f6",borderRadius:8,cursor:"pointer",textAlign:"center",color:"#60a5fa",fontSize:12,fontWeight:600}}>
          📁 {fileName||"Click to upload Excel / CSV"}
          <input type="file" accept=".xlsx,.xls,.csv" onChange={handleFile} style={{display:"none"}}/>
        </label>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:9,alignContent:"start"}}>
        {[["Transactions",txns.length,"#60a5fa"],["Income","£"+income.toFixed(2),"#34d399"],["Expenses","£"+expense.toFixed(2),"#f87171"],["Categorised",categorised+"/"+txns.length,categorised===txns.length&&txns.length>0?"#34d399":"#fbbf24"]].map(([l,v,c])=>(
          <div key={l} style={{background:"#0f1421",border:`1px solid ${c}33`,borderRadius:9,padding:12}}>
            <div style={{fontSize:10,color:"#64748b",fontWeight:700,textTransform:"uppercase"}}>{l}</div>
            <div style={{fontSize:17,fontWeight:800,color:c,marginTop:4}}>{v}</div>
          </div>
        ))}
      </div>
    </div>

    {txns.length===0&&<div style={{textAlign:"center",padding:40,color:"#374151",fontSize:13,border:"1px dashed #1e2535",borderRadius:10}}>Upload a bank statement to categorise transactions.</div>}

    {txns.length>0&&<>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
        <div style={{fontSize:13,color:"#94a3b8"}}>{txns.length} transactions · {categorised} categorised</div>
        <button onClick={exportCat} style={{...BG,padding:"7px 14px",fontSize:12}}>⬇ Export Categorised Excel</button>
      </div>
      <div style={{maxHeight:420,overflowY:"auto",border:"1px solid #1e2535",borderRadius:9}}>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
          <thead style={{position:"sticky",top:0,zIndex:1}}>
            <tr>
              <th style={{...TH,minWidth:85}}>Date</th>
              <th style={{...TH,minWidth:180}}>Description</th>
              <th style={{...TH,minWidth:85}}>Amount</th>
              <th style={{...TH,minWidth:100}}>Type</th>
              <th style={{...TH,minWidth:155}}>Category</th>
              <th style={{...TH,minWidth:125}}>Site</th>
              <th style={{...TH,minWidth:115}}>Client</th>
              <th style={{...TH,minWidth:100}}>Notes</th>
            </tr>
          </thead>
          <tbody>
            {txns.map((t,i)=>(
              <tr key={t.id} style={{background:i%2===0?"#111827":"#0f1421"}}>
                <td style={{...TD,color:"#94a3b8",whiteSpace:"nowrap"}}>{t.date}</td>
                <td style={{...TD,maxWidth:180}}><div style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",color:"#e2e8f0"}} title={t.description}>{t.description}</div></td>
                <td style={{...TD,fontWeight:700,color:t.amount>=0?"#34d399":"#f87171",whiteSpace:"nowrap"}}>£{Math.abs(t.amount).toFixed(2)}</td>
                <td style={TD}><select value={t.type} onChange={e=>upT(t.id,"type",e.target.value)} style={{...INP,fontSize:10,padding:"2px 5px",cursor:"pointer",color:t.type==="income"?"#34d399":"#f87171",width:"auto",minWidth:90}}>
                  <option value="income">Income</option><option value="expense">Expense</option>
                </select></td>
                <td style={TD}><select value={t.category} onChange={e=>upT(t.id,"category",e.target.value)} style={{...INP,fontSize:10,padding:"2px 5px",cursor:"pointer",width:"auto",minWidth:150}}>
                  <option value="">— Category —</option>
                  <optgroup label="Income">{INCOME_CATS.map(c=><option key={c} value={c}>{c}</option>)}</optgroup>
                  <optgroup label="Expenses">{EXPENSE_CATS.map(c=><option key={c} value={c}>{c}</option>)}</optgroup>
                </select></td>
                <td style={TD}><select value={t.siteId} onChange={e=>upT(t.id,"siteId",e.target.value)} style={{...INP,fontSize:10,padding:"2px 5px",cursor:"pointer",width:"auto",minWidth:120}}>
                  <option value="">— Site —</option>
                  {allSites.filter(s=>!isOff(s.name)).map(s=><option key={s.id} value={s.id}>{s.name}</option>)}
                </select></td>
                <td style={TD}><select value={t.clientId} onChange={e=>upT(t.id,"clientId",e.target.value)} style={{...INP,fontSize:10,padding:"2px 5px",cursor:"pointer",width:"auto",minWidth:110}}>
                  <option value="">— Client —</option>
                  {clients.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}
                </select></td>
                <td style={TD}><input value={t.notes} onChange={e=>upT(t.id,"notes",e.target.value)} placeholder="Notes…" style={{...INP,fontSize:10,padding:"2px 5px",minWidth:90}}/></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </>}

    <div style={{display:"flex",gap:10,justifyContent:"flex-end",marginTop:18,borderTop:"1px solid #1e2535",paddingTop:14}}>
      <button onClick={onClose} style={{padding:"8px 18px",background:"#1e2535",border:"1px solid #2d3555",borderRadius:7,color:"#94a3b8",cursor:"pointer"}}>Close</button>
    </div>
  </Overlay>;
}


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
            <div key={site} style={{background:"#1a1f2e",border:`1px solid ${getSiteColor(site,allSites)}33`,borderRadius:8,padding:"9px 12px",cursor:"pointer",transition:"border-color 0.15s"}}
              onClick={()=>{const so=allSites.find(s=>site===s.name||site.toUpperCase().includes(s.name.toUpperCase()));if(so)setModal({type:"siteDetail",site:so});}}
              onMouseEnter={e=>e.currentTarget.style.borderColor=getSiteColor(site,allSites)+"88"}
              onMouseLeave={e=>e.currentTarget.style.borderColor=getSiteColor(site,allSites)+"33"}>
              <div style={{display:"flex",alignItems:"center",gap:5,marginBottom:4}}>
                <span style={{width:7,height:7,borderRadius:"50%",background:getSiteColor(site,allSites),flexShrink:0}}/>
                <span style={{fontSize:11,color:"#cbd5e1",fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{site}</span>
              </div>
              <div style={{fontSize:18,fontWeight:800,color:"#34d399"}}>£{gross.toFixed(2)}</div>
              <div style={{fontSize:10,color:"#64748b"}}>{sw.size} worker{sw.size!==1?"s":""}</div>
              <div style={{display:"flex",gap:5,marginTop:4}}>
                <div style={{fontSize:10,color:"#60a5fa"}}>📂 Open detail</div>
                <div onClick={e=>{e.stopPropagation();const s=allSites.find(x=>site===x.name||site.toUpperCase().includes(x.name.toUpperCase()));if(s)openSiteWindow(s,clients,workers,activeDays,siteHours);}} style={{fontSize:10,color:"#34d399",cursor:"pointer"}}>🔗 New window</div>
              </div>
            </div>
          ))}
        </div>
      </div>
    ))}
  </div>;
}

// ─── Company Constants (from real invoice) ───────────────────────────────────
const OUR_COMPANY={
  name:"Bright Metalwork Ltd.",
  address:"Suite 608 Crown House, North Circular Road, London, NW10 7PN",
  crn:"12020937",utr:"8672212219",vatNo:"330921430",
  bankName:"HSBC Bank",sortCode:"40-25-02",accountNo:"03869261",
  email:"lucian@bright-group.org",phone:"+44 (0)771 078 3500",
  contactName:"Lucian Ciocoiu",
  logo:"", // set via settings
};

// ─── Clients Modal — with day rates per trade ─────────────────────────────────
function ClientsModal({clients,onSave,onClose}){
  const [list,setList]=useState(clients.map(c=>({...c,dayRates:c.dayRates||{},address:c.address||"",crn:c.crn||""})));
  const [nn,setNn]=useState(""),[ne,setNe]=useState(""),[np,setNp]=useState(""),[nc,setNc]=useState(PRESET_COLORS[2]);
  const add=()=>{if(!nn.trim())return;setList(l=>[...l,{id:"c"+Date.now(),name:nn.trim(),email:ne,phone:np,color:nc,notes:"",dayRates:{},address:"",crn:""}]);setNn("");setNe("");setNp("");};
  const up=(id,k,v)=>setList(l=>l.map(x=>x.id===id?{...x,[k]:v}:x));
  const setDayRate=(id,pos,val)=>setList(l=>l.map(x=>x.id===id?{...x,dayRates:{...x.dayRates,[pos]:val}}:x));
  const rm=id=>{if(window.confirm("Delete?"))setList(l=>l.filter(x=>x.id!==id));};
  return <Overlay onClose={onClose} wide>
    <MH title="👔 Manage Clients" onClose={onClose}/>
    <Sec title="Add New Client">
      <div style={{display:"flex",gap:9,alignItems:"flex-end",flexWrap:"wrap"}}>
        <div style={{flex:2,minWidth:140}}><label style={LBL}>Name</label><input value={nn} onChange={e=>setNn(e.target.value)} style={INP}/></div>
        <div style={{flex:1,minWidth:130}}><label style={LBL}>Email</label><input value={ne} onChange={e=>setNe(e.target.value)} type="email" style={INP}/></div>
        <div style={{flex:1,minWidth:110}}><label style={LBL}>Phone</label><input value={np} onChange={e=>setNp(e.target.value)} style={INP}/></div>
        <div style={{minWidth:260}}><label style={LBL}>Colour</label><ColorPicker value={nc} onChange={setNc}/></div>
        <button onClick={add} style={{...BP,whiteSpace:"nowrap"}}>+ Add</button>
      </div>
    </Sec>
    {list.length===0&&<div style={{textAlign:"center",padding:28,color:"#374151",fontSize:13}}>No clients yet.</div>}
    <div style={{display:"flex",flexDirection:"column",gap:12}}>
      {list.map(c=><div key={c.id} style={{background:"#0f1421",borderRadius:10,padding:14,border:`1px solid ${c.color}44`}}>
        <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:9}}>
          <span style={{width:11,height:11,borderRadius:"50%",background:c.color,flexShrink:0}}/>
          <input value={c.name} onChange={e=>up(c.id,"name",e.target.value)} style={{flex:1,background:"#1a1f2e",border:"1px solid #2d3555",borderRadius:5,padding:"5px 8px",color:"#e2e8f0",fontSize:13,fontWeight:700,outline:"none"}}/>
          <button onClick={()=>rm(c.id)} style={{background:"#2d1515",border:"1px solid #ef4444",borderRadius:5,color:"#f87171",cursor:"pointer",fontSize:11,padding:"3px 7px",fontWeight:700}}>Del</button>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:"0 9px",marginBottom:8}}>
          <div><label style={LBL}>Email</label><input value={c.email||""} onChange={e=>up(c.id,"email",e.target.value)} style={{...INP,fontSize:12,padding:"5px 7px"}}/></div>
          <div><label style={LBL}>Phone</label><input value={c.phone||""} onChange={e=>up(c.id,"phone",e.target.value)} style={{...INP,fontSize:12,padding:"5px 7px"}}/></div>
          <div><label style={LBL}>CRN</label><input value={c.crn||""} onChange={e=>up(c.id,"crn",e.target.value)} style={{...INP,fontSize:12,padding:"5px 7px"}}/></div>
        </div>
        <div style={{marginBottom:8}}><label style={LBL}>Billing Address</label><input value={c.address||""} onChange={e=>up(c.id,"address",e.target.value)} style={{...INP,fontSize:12,padding:"5px 7px"}}/></div>
        <div style={{marginBottom:8}}><label style={LBL}>Notes</label><input value={c.notes||""} onChange={e=>up(c.id,"notes",e.target.value)} style={{...INP,fontSize:12,padding:"5px 7px"}}/></div>
        {/* Agreed Day Rates by Team Type */}
        <div style={{background:"#1a2035",borderRadius:8,padding:12,border:"1px solid #2d3555",marginTop:4}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:9}}>
            <div style={{fontSize:11,color:"#60a5fa",fontWeight:700,textTransform:"uppercase",letterSpacing:"0.06em"}}>Agreed Day Rates</div>
            <button onClick={()=>up(c.id,"rates",[...(c.rates||[]),{id:Date.now().toString(36),teamType:"welding",dayRate:0,notes:""}])} style={{padding:"3px 10px",background:"#1e3a5f",border:"1px solid #3b82f6",borderRadius:5,color:"#60a5fa",cursor:"pointer",fontSize:11,fontWeight:700}}>+ Add Rate</button>
          </div>
          {(c.rates||[]).length===0&&<div style={{color:"#374151",fontSize:11,textAlign:"center",padding:"6px 0"}}>No rates set. Click "+ Add Rate" to start.</div>}
          {(c.rates||[]).map(r=>(
            <div key={r.id} style={{display:"grid",gridTemplateColumns:"2fr 1fr 1fr auto",gap:7,marginBottom:7,alignItems:"flex-end"}}>
              <div><label style={LBL}>Team / Role Type</label>
                <select value={r.teamType} onChange={e=>{const v=e.target.value;up(c.id,"rates",(c.rates||[]).map(x=>x.id===r.id?{...x,teamType:v}:x));}} style={{...INP,cursor:"pointer",fontSize:11,padding:"4px 7px"}}>
                  {TEAM_TYPES.map(t=><option key={t.key} value={t.key}>{t.label}</option>)}
                </select>
              </div>
              <div><label style={LBL}>Day Rate £</label>
                <input type="number" value={r.dayRate||0} onChange={e=>{const v=Number(e.target.value);up(c.id,"rates",(c.rates||[]).map(x=>x.id===r.id?{...x,dayRate:v}:x));}} style={{...INP,fontSize:11,padding:"4px 7px",textAlign:"right"}}/>
              </div>
              <div><label style={LBL}>Notes</label>
                <input value={r.notes||""} onChange={e=>{const v=e.target.value;up(c.id,"rates",(c.rates||[]).map(x=>x.id===r.id?{...x,notes:v}:x));}} placeholder="Notes…" style={{...INP,fontSize:11,padding:"4px 7px"}}/>
              </div>
              <button onClick={()=>{up(c.id,"rates",(c.rates||[]).filter(x=>x.id!==r.id));}} style={{padding:"5px 9px",background:"#2d1515",border:"1px solid #ef4444",borderRadius:5,color:"#f87171",cursor:"pointer",fontSize:11,fontWeight:700,alignSelf:"flex-end"}}>✕</button>
            </div>
          ))}
        </div>
        <div style={{marginTop:8}}><label style={LBL}>Colour</label><ColorPicker value={c.color} onChange={(col)=>{up(c.id,"color",col);}}/></div>
      </div>)}
    </div>
    <div style={{display:"flex",gap:10,justifyContent:"flex-end",marginTop:20,borderTop:"1px solid #1e2535",paddingTop:16}}>
      <button onClick={onClose} style={{padding:"8px 18px",background:"#1e2535",border:"1px solid #2d3555",borderRadius:7,color:"#94a3b8",cursor:"pointer"}}>Cancel</button>
      <button onClick={()=>onSave(list)} style={BG}>Save Clients</button>
    </div>
  </Overlay>;
}

// ─── Scope Modal (unchanged from v2) ─────────────────────────────────────────
function emptyScopeItem(siteId){return {id:"si"+Date.now()+Math.random(),siteId,description:"",unit:"item",qty:1,unitIncome:0,overheadPct:15};}
function ScopeModal({site,scopeItems,onSave,onClose}){
  const [items,setItems]=useState((scopeItems||[]).map(x=>({...x})));
  const addItem=()=>setItems(s=>[...s,emptyScopeItem(site.id)]);
  const upItem=(id,k,v)=>setItems(s=>s.map(x=>x.id===id?{...x,[k]:v}:x));
  const rmItem=id=>setItems(s=>s.filter(x=>x.id!==id));
  const UNITS=["item","m","m²","m³","day","week","nr","set","lot","tonne","kg"];
  const totals=items.reduce((a,it)=>{const inc=(it.qty||0)*(it.unitIncome||0);const bud=inc*(1+(it.overheadPct||0)/100);return{income:a.income+inc,budget:a.budget+bud};},{income:0,budget:0});
  return <Overlay onClose={onClose} wide>
    <MH title={`📐 Scope of Works — ${site.name}`} onClose={onClose}/>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10,marginBottom:16}}>
      {[["Agreed Income","£"+totals.income.toFixed(2),"#34d399"],["Budget (incl. OH)","£"+totals.budget.toFixed(2),"#60a5fa"],["Overhead Value","£"+(totals.budget-totals.income).toFixed(2),"#a78bfa"]].map(([l,v,c])=>
        <div key={l} style={{background:"#0f1421",border:`1px solid ${c}44`,borderRadius:10,padding:"10px 14px"}}><div style={{fontSize:10,color:"#64748b",fontWeight:700,textTransform:"uppercase"}}>{l}</div><div style={{fontSize:20,fontWeight:800,color:c}}>{v}</div></div>
      )}
    </div>
    <div style={{marginBottom:12}}>
      <div style={{display:"grid",gridTemplateColumns:"2fr 80px 70px 90px 70px 32px",gap:6,marginBottom:7}}>
        {["Description","Unit","Qty","Income/Unit","OH %",""].map((h,i)=><div key={i} style={{fontSize:10,color:"#64748b",fontWeight:700,textTransform:"uppercase"}}>{h}</div>)}
      </div>
      {items.map(it=>{
        const lineInc=(it.qty||0)*(it.unitIncome||0);const lineBud=lineInc*(1+(it.overheadPct||0)/100);
        return <div key={it.id} style={{display:"grid",gridTemplateColumns:"2fr 80px 70px 90px 70px 32px",gap:6,marginBottom:7,alignItems:"center",background:"#0f1421",borderRadius:7,padding:"8px 10px",border:"1px solid #1e2535"}}>
          <input value={it.description} onChange={e=>upItem(it.id,"description",e.target.value)} placeholder="e.g. Install curtain walling panel" style={{...INP,padding:"5px 8px",fontSize:12}}/>
          <select value={it.unit} onChange={e=>upItem(it.id,"unit",e.target.value)} style={{...INP,padding:"5px 6px",fontSize:11,cursor:"pointer"}}>{UNITS.map(u=><option key={u} value={u}>{u}</option>)}</select>
          <input type="number" min="0" value={it.qty} onChange={e=>upItem(it.id,"qty",+e.target.value||0)} style={{...INP,padding:"5px 6px",fontSize:12,textAlign:"right"}}/>
          <div style={{position:"relative"}}><span style={{position:"absolute",left:8,top:"50%",transform:"translateY(-50%)",color:"#64748b",fontSize:12}}>£</span><input type="number" min="0" value={it.unitIncome} onChange={e=>upItem(it.id,"unitIncome",+e.target.value||0)} style={{...INP,padding:"5px 6px 5px 20px",fontSize:12,textAlign:"right"}}/></div>
          <div style={{position:"relative"}}><input type="number" min="0" max="100" value={it.overheadPct} onChange={e=>upItem(it.id,"overheadPct",+e.target.value||0)} style={{...INP,padding:"5px 18px 5px 6px",fontSize:12,textAlign:"right"}}/><span style={{position:"absolute",right:7,top:"50%",transform:"translateY(-50%)",color:"#64748b",fontSize:11}}>%</span></div>
          <button onClick={()=>rmItem(it.id)} style={{background:"#2d1515",border:"1px solid #ef4444",borderRadius:5,color:"#f87171",cursor:"pointer",fontSize:13,padding:"4px 7px",fontWeight:700}}>✕</button>
          <div style={{gridColumn:"1/-1",display:"flex",gap:14,paddingTop:3,paddingLeft:2}}>
            <span style={{fontSize:10,color:"#64748b"}}>Income: <span style={{color:"#34d399",fontWeight:700}}>£{lineInc.toFixed(2)}</span></span>
            <span style={{fontSize:10,color:"#64748b"}}>Budget: <span style={{color:"#60a5fa",fontWeight:700}}>£{lineBud.toFixed(2)}</span></span>
          </div>
        </div>;
      })}
    </div>
    <button onClick={addItem} style={{...BP,padding:"6px 14px",fontSize:12,marginBottom:16}}>+ Add Line Item</button>
    <div style={{display:"flex",gap:10,justifyContent:"flex-end",borderTop:"1px solid #1e2535",paddingTop:16}}>
      <button onClick={onClose} style={{padding:"8px 18px",background:"#1e2535",border:"1px solid #2d3555",borderRadius:7,color:"#94a3b8",cursor:"pointer"}}>Cancel</button>
      <button onClick={()=>onSave(items)} style={BG}>Save Scope</button>
    </div>
  </Overlay>;
}

// ─── Budget View (unchanged from v2) ─────────────────────────────────────────
function BudgetView({workers,clients,allSites,activeDays,siteHours,scopeData,onEditScope}){
  const labourBySite=useMemo(()=>{const m={};workers.forEach(w=>{const {bd}=calcPay(w,activeDays,siteHours);Object.values(bd).forEach(({site,gross})=>{if(!m[site])m[site]={labourCost:0,wIds:new Set()};m[site].labourCost+=gross;m[site].wIds.add(w.id);});});return m;},[workers,activeDays,siteHours]);
  const rows=useMemo(()=>allSites.filter(s=>!isOff(s.name)).map(site=>{
    const items=(scopeData[site.id]||[]);
    const income=items.reduce((a,it)=>a+(it.qty||0)*(it.unitIncome||0),0);
    const budget=items.reduce((a,it)=>{const inc=(it.qty||0)*(it.unitIncome||0);return a+inc*(1+(it.overheadPct||0)/100);},0);
    const labour=labourBySite[site.name]?.labourCost||0;
    const profit=income-labour;
    const margin=income>0?((profit/income)*100):null;
    const client=clients.find(c=>c.id===site.clientId);
    return{site,items,income,budget,labour,profit,margin,client,workers:labourBySite[site.name]?.wIds.size||0};
  }).filter(r=>r.income>0||r.labour>0),[allSites,scopeData,labourBySite,clients]);
  const grand=rows.reduce((a,r)=>({income:a.income+r.income,budget:a.budget+r.budget,labour:a.labour+r.labour}),{income:0,budget:0,labour:0});
  return <div style={{padding:"14px 18px"}}>
    <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10,marginBottom:18}}>
      {[["Total Income","£"+grand.income.toFixed(2),"#34d399"],["Total Budget","£"+grand.budget.toFixed(2),"#60a5fa"],["Labour Cost","£"+grand.labour.toFixed(2),"#f87171"],["Margin","£"+(grand.income-grand.labour).toFixed(2),grand.income-grand.labour>=0?"#a78bfa":"#f87171"]].map(([l,v,c])=>
        <div key={l} style={{background:"#1a1f2e",border:`1px solid ${c}44`,borderRadius:10,padding:"10px 14px"}}><div style={{fontSize:10,color:"#64748b",fontWeight:700,textTransform:"uppercase"}}>{l}</div><div style={{fontSize:20,fontWeight:800,color:c}}>{v}</div></div>
      )}
    </div>
    {rows.length===0&&<div style={{textAlign:"center",padding:40,color:"#374151",fontSize:13}}>No scope defined. Open Budget tab on a site to add scope of works.</div>}
    {rows.map(r=>{
      const overBudget=r.labour>r.budget&&r.budget>0;const pct=r.budget>0?Math.min(100,(r.labour/r.budget)*100):0;
      return <div key={r.site.id} style={{background:"#111827",border:`1px solid ${r.site.color}44`,borderRadius:11,padding:15,marginBottom:12}}>
        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:12,flexWrap:"wrap"}}>
          <span style={{width:11,height:11,borderRadius:"50%",background:r.site.color,flexShrink:0}}/>
          <span style={{fontWeight:800,color:"#f1f5f9",fontSize:14,flex:1}}>{r.site.name}</span>
          {r.client&&<span style={{fontSize:11,color:r.client.color,fontWeight:600}}>👔 {r.client.name}</span>}
          <button onClick={()=>onEditScope(r.site)} style={{padding:"4px 10px",background:"#1e3a5f",border:"1px solid #3b82f6",borderRadius:6,color:"#60a5fa",cursor:"pointer",fontSize:11,fontWeight:700}}>📐 Edit Scope</button>
        </div>
        <div style={{marginBottom:12}}>
          <div style={{display:"flex",justifyContent:"space-between",fontSize:11,color:"#64748b",marginBottom:4}}><span>Labour cost vs Budget</span><span style={{color:overBudget?"#f87171":"#34d399",fontWeight:700}}>{pct.toFixed(0)}%{overBudget?" ⚠️ OVER BUDGET":""}</span></div>
          <div style={{height:8,background:"#1e2535",borderRadius:4,overflow:"hidden"}}><div style={{height:"100%",borderRadius:4,background:overBudget?"#f87171":pct>75?"#fbbf24":"#34d399",width:pct+"%",transition:"width 0.4s"}}/></div>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(130px,1fr))",gap:8}}>
          {[["Agreed Income","£"+r.income.toFixed(2),"#34d399"],["Budget (w/OH)","£"+r.budget.toFixed(2),"#60a5fa"],["Labour Cost","£"+r.labour.toFixed(2),"#f87171"],["Profit","£"+r.profit.toFixed(2),r.profit>=0?"#a78bfa":"#f87171"],r.margin!==null&&["Margin",r.margin.toFixed(1)+"%",r.margin>=10?"#34d399":r.margin>=0?"#fbbf24":"#f87171"]].filter(Boolean).map(([l,v,c])=>
            <div key={l} style={{background:"#0f1421",borderRadius:8,padding:"8px 11px",border:`1px solid ${c}22`}}><div style={{fontSize:9,color:"#64748b",fontWeight:700,textTransform:"uppercase"}}>{l}</div><div style={{fontSize:15,fontWeight:800,color:c}}>{v}</div></div>
          )}
        </div>
      </div>;
    })}
  </div>;
}

// ─── Invoice helpers ──────────────────────────────────────────────────────────
function calcInvoiceTotals(inv){
  const subtotal=inv.lines.reduce((a,l)=>a+(l.qty||0)*(l.rate||0),0);
  // VAT Reverse Charge: show VAT amount but total stays as subtotal
  const isRC=(inv.vatType||"standard")==="reverse_charge";
  const vat=subtotal*(inv.vatRate||0)/100;
  return{subtotal,vat,total:isRC?subtotal:subtotal+vat,isRC};
}
function getNextInvoiceNumber(invoices){
  // Find the highest INV number; from PDF we know last was INV1484
  let max=1484;
  invoices.forEach(inv=>{const m=(inv.invoiceNumber||"").match(/INV(\d+)/i);if(m)max=Math.max(max,+m[1]);});
  return "INV"+(max+1);
}
function emptyInvoice(clients,allSites,invoices=[]){
  return {
    id:"inv"+Date.now(),
    invoiceNumber:getNextInvoiceNumber(invoices),
    issueDate:new Date().toISOString().slice(0,10),
    dueDate:"",
    paymentTerms:"On Receipt",
    clientId:clients[0]?.id||"",
    siteId:allSites.filter(s=>!isOff(s.name))[0]?.id||"",
    poNumber:"",
    weekCommencing:"",
    vatRate:20,
    vatType:"reverse_charge", // matches the PDF invoice format
    notes:"",
    status:"draft",
    lines:[],
    lineMode:"dayworks", // "dayworks" or "scope"
  };
}

// ─── Invoice Modal — redesigned to match actual invoice layout ────────────────
function InvoiceModal({invoice,clients,allSites,scopeData,workers,invoices,onSave,onClose}){
  const [f,setF]=useState({...invoice,lines:(invoice.lines||[]).map(l=>({...l}))});
  const [tab,setTab]=useState("details");
  const [lineModeChosen,setLineModeChosen]=useState(!!invoice.lines?.length);
  const set=(k,v)=>setF(x=>({...x,[k]:v}));
  const {subtotal,vat,total,isRC}=calcInvoiceTotals(f);

  // Auto-fill client when site changes
  const handleSiteChange=siteId=>{
    set("siteId",siteId);
    const site=allSites.find(s=>s.id===siteId);
    if(site?.clientId&&!f.clientId) set("clientId",site.clientId);
    // Also update weekCommencing hint
  };
  const handleClientChange=clientId=>{
    set("clientId",clientId);
  };

  const client=clients.find(c=>c.id===f.clientId);
  const site=allSites.find(s=>s.id===f.siteId);
  const scopeItems=site?scopeData[site.id]||[]:[];

  // Day-rate line items from client day rates
  const addDayworkLine=()=>{
    const newLine={id:"il"+Date.now()+Math.random(),description:"",qty:1,rate:0,unit:"day",type:"daywork"};
    setF(x=>({...x,lines:[...x.lines,newLine]}));
  };
  const importScopeItem=it=>{
    const alreadyAdded=f.lines.some(l=>l.scopeItemId===it.id);
    if(alreadyAdded)return;
    const lineInc=(it.qty||0)*(it.unitIncome||0);
    const newLine={id:"il"+Date.now()+Math.random(),description:it.description,qty:it.qty,rate:it.unitIncome,unit:it.unit,type:"scope",scopeItemId:it.id};
    setF(x=>({...x,lines:[...x.lines,newLine]}));
  };
  const upLine=(id,k,v)=>setF(x=>({...x,lines:x.lines.map(l=>l.id===id?{...l,[k]:v}:l)}));
  const rmLine=id=>setF(x=>({...x,lines:x.lines.filter(l=>l.id!==id)}));
  const UNITS=["day","week","nr","m","m²","m³","item","lot","set","hour","%"];

  // If no lines yet, show mode chooser
  const showModeChooser=!lineModeChosen&&f.lines.length===0;

  return <Overlay onClose={onClose} wide>
    <MH title={f.invoiceNumber?`🧾 ${f.invoiceNumber}`:"🧾 New Invoice"} onClose={onClose}/>
    <TabBar tabs={[["details","🏢 Details"],["lines","📋 Line Items"],["preview","👁 Preview"]]} active={tab} onChange={setTab}/>

    {tab==="details"&&<div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:"0 14px"}}>
        <FI label="Invoice Number" value={f.invoiceNumber} onChange={v=>set("invoiceNumber",v)}/>
        <FI label="Issue Date" value={f.issueDate} onChange={v=>set("issueDate",v)} type="date"/>
        <FSel label="Status" value={f.status} onChange={v=>set("status",v)} options={[{value:"draft",label:"Draft"},{value:"sent",label:"Sent"},{value:"paid",label:"Paid"}]}/>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"0 14px"}}>
        <FI label="PO Number" value={f.poNumber||""} onChange={v=>set("poNumber",v)} placeholder="Client PO number"/>
        <FI label="Week Commencing" value={f.weekCommencing||""} onChange={v=>set("weekCommencing",v)} placeholder="e.g. W.C.01.06.2026"/>
        <FI label="Payment Terms" value={f.paymentTerms||"On Receipt"} onChange={v=>set("paymentTerms",v)}/>
        <FI label="Due Date" value={f.dueDate||""} onChange={v=>set("dueDate",v)} type="date"/>
      </div>
      {/* Client — auto-fills from site */}
      <Sec title="Client & Site">
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"0 14px"}}>
          <div style={{marginBottom:11}}>
            <label style={LBL}>Site (auto-fills client)</label>
            <select value={f.siteId||""} onChange={e=>handleSiteChange(e.target.value)} style={{...INP,cursor:"pointer"}}>
              <option value="">— Select Site —</option>
              {allSites.filter(s=>!isOff(s.name)).map(s=><option key={s.id} value={s.id}>{s.name}</option>)}
            </select>
          </div>
          <div style={{marginBottom:11}}>
            <label style={LBL}>Client (Bill To)</label>
            <select value={f.clientId||""} onChange={e=>handleClientChange(e.target.value)} style={{...INP,cursor:"pointer"}}>
              <option value="">— Select Client —</option>
              {clients.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}
            </select>
          </div>
        </div>
        {client&&<div style={{background:"#0f1421",borderRadius:8,padding:"10px 12px",fontSize:12,border:`1px solid ${client.color}33`}}>
          <div style={{fontWeight:700,color:client.color,marginBottom:3}}>{client.name}</div>
          {client.address&&<div style={{color:"#94a3b8"}}>{client.address}</div>}
          {client.crn&&<div style={{color:"#64748b"}}>CRN: {client.crn}</div>}
        </div>}
      </Sec>
      <Sec title="VAT Settings">
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"0 14px"}}>
          <div style={{marginBottom:11}}><label style={LBL}>VAT Rate %</label>
            <select value={f.vatRate} onChange={e=>set("vatRate",+e.target.value)} style={{...INP,cursor:"pointer"}}>
              <option value={0}>0%</option><option value={5}>5%</option><option value={20}>20%</option>
            </select>
          </div>
          <div style={{marginBottom:11}}><label style={LBL}>VAT Treatment</label>
            <select value={f.vatType||"reverse_charge"} onChange={e=>set("vatType",e.target.value)} style={{...INP,cursor:"pointer"}}>
              <option value="reverse_charge">Reverse Charge (VAT shown, not added)</option>
              <option value="standard">Standard (VAT added to total)</option>
              <option value="exempt">Exempt / Zero Rated</option>
            </select>
          </div>
        </div>
      </Sec>
      <FI label="Notes / Payment Reference" value={f.notes||""} onChange={v=>set("notes",v)} placeholder="e.g. Payment terms, reference notes"/>
    </div>}

    {tab==="lines"&&<div>
      {/* Mode chooser if no lines yet */}
      {showModeChooser&&<div style={{background:"#0f1421",borderRadius:12,padding:24,marginBottom:16,textAlign:"center",border:"1px solid #2d3555"}}>
        <div style={{fontSize:14,color:"#e2e8f0",fontWeight:700,marginBottom:8}}>How do you want to build this invoice?</div>
        <div style={{fontSize:12,color:"#64748b",marginBottom:20}}>Choose the line item type — you can mix both after</div>
        <div style={{display:"flex",gap:14,justifyContent:"center",flexWrap:"wrap"}}>
          <button onClick={()=>{setLineModeChosen(true);set("lineMode","dayworks");addDayworkLine();}} style={{padding:"16px 22px",background:"#1e3a5f",border:"2px solid #3b82f6",borderRadius:10,color:"#60a5fa",cursor:"pointer",fontSize:13,fontWeight:700,minWidth:180}}>
            <div style={{fontSize:22,marginBottom:6}}>🔧</div>
            <div style={{fontWeight:800}}>Day Works</div>
            <div style={{fontSize:11,color:"#64748b",marginTop:4}}>Daily/weekly teams, individual dates, agreed day rates</div>
          </button>
          {scopeItems.length>0&&<button onClick={()=>{setLineModeChosen(true);set("lineMode","scope");scopeItems.forEach(importScopeItem);}} style={{padding:"16px 22px",background:"#1a2d20",border:"2px solid #10b981",borderRadius:10,color:"#34d399",cursor:"pointer",fontSize:13,fontWeight:700,minWidth:180}}>
            <div style={{fontSize:22,marginBottom:6}}>📐</div>
            <div style={{fontWeight:800}}>Scope of Works</div>
            <div style={{fontSize:11,color:"#64748b",marginTop:4}}>Import agreed scope items ({scopeItems.length} available)</div>
          </button>}
          <button onClick={()=>{setLineModeChosen(true);}} style={{padding:"16px 22px",background:"#1e2535",border:"2px solid #374151",borderRadius:10,color:"#94a3b8",cursor:"pointer",fontSize:13,fontWeight:700,minWidth:180}}>
            <div style={{fontSize:22,marginBottom:6}}>✏️</div>
            <div style={{fontWeight:800}}>Manual / Mixed</div>
            <div style={{fontSize:11,color:"#64748b",marginTop:4}}>Start blank, add lines manually</div>
          </button>
        </div>
      </div>}

      {/* Scope import button */}
      {lineModeChosen&&scopeItems.length>0&&<Sec title="📐 Import from Scope of Works" color="#60a5fa">
        <div style={{display:"flex",flexWrap:"wrap",gap:7}}>
          {scopeItems.map(it=>{const alreadyAdded=f.lines.some(l=>l.scopeItemId===it.id);const lineInc=(it.qty||0)*(it.unitIncome||0);return(
            <button key={it.id} onClick={()=>importScopeItem(it)} style={{padding:"6px 11px",background:alreadyAdded?"#1a3020":"#1e3a5f",border:`1px solid ${alreadyAdded?"#10b981":"#3b82f6"}`,borderRadius:7,color:alreadyAdded?"#34d399":"#60a5fa",cursor:alreadyAdded?"default":"pointer",fontSize:11,fontWeight:600}}>
              {alreadyAdded?"✓ ":""}{it.description||"Item"} · £{lineInc.toFixed(2)}
            </button>
          );})}
        </div>
      </Sec>}

      {/* Client day rates quick-fill */}
      {lineModeChosen&&client?.dayRates&&Object.keys(client.dayRates).some(k=>client.dayRates[k])&&<Sec title="⚡ Quick Add — Client Day Rates" color="#fbbf24">
        <div style={{display:"flex",flexWrap:"wrap",gap:7}}>
          {POSITIONS.filter(p=>client.dayRates[p]).map(pos=><button key={pos} onClick={()=>{const newLine={id:"il"+Date.now()+Math.random(),description:`Team - ${pos}`,qty:1,rate:client.dayRates[pos],unit:"day",type:"daywork"};setF(x=>({...x,lines:[...x.lines,newLine]}));}} style={{padding:"6px 11px",background:"#2d2008",border:"1px solid #f59e0b",borderRadius:7,color:"#fbbf24",cursor:"pointer",fontSize:11,fontWeight:600}}>
            + {pos} · £{client.dayRates[pos]}/day
          </button>)}
        </div>
      </Sec>}

      {/* Line items */}
      <div style={{marginBottom:12}}>
        <div style={{display:"grid",gridTemplateColumns:"3fr 70px 80px 80px 100px 32px",gap:6,marginBottom:7}}>
          {["Description","Unit","Qty","Rate/Unit","Line Total",""].map((h,i)=><div key={i} style={{fontSize:10,color:"#64748b",fontWeight:700,textTransform:"uppercase"}}>{h}</div>)}
        </div>
        {f.lines.map(l=>{
          const lineTotal=(l.qty||0)*(l.rate||0);
          return <div key={l.id} style={{display:"grid",gridTemplateColumns:"3fr 70px 80px 80px 100px 32px",gap:6,marginBottom:6,alignItems:"center"}}>
            <input value={l.description} onChange={e=>upLine(l.id,"description",e.target.value)} placeholder="e.g. 5 June 2026 - Team of 2 - Welding" style={{...INP,padding:"5px 8px",fontSize:12}}/>
            <select value={l.unit} onChange={e=>upLine(l.id,"unit",e.target.value)} style={{...INP,padding:"5px 4px",fontSize:11,cursor:"pointer"}}>{UNITS.map(u=><option key={u} value={u}>{u}</option>)}</select>
            <input type="number" min="0" value={l.qty} onChange={e=>upLine(l.id,"qty",+e.target.value||0)} style={{...INP,padding:"5px 6px",fontSize:12,textAlign:"right"}}/>
            <div style={{position:"relative"}}><span style={{position:"absolute",left:7,top:"50%",transform:"translateY(-50%)",color:"#64748b",fontSize:12}}>£</span><input type="number" min="0" value={l.rate} onChange={e=>upLine(l.id,"rate",+e.target.value||0)} style={{...INP,padding:"5px 6px 5px 18px",fontSize:12,textAlign:"right"}}/></div>
            <div style={{fontSize:13,fontWeight:700,color:"#34d399",textAlign:"right"}}>£{lineTotal.toFixed(2)}</div>
            <button onClick={()=>rmLine(l.id)} style={{background:"#2d1515",border:"1px solid #ef4444",borderRadius:5,color:"#f87171",cursor:"pointer",fontSize:13,padding:"4px 7px",fontWeight:700}}>✕</button>
          </div>;
        })}
        {f.lines.length===0&&lineModeChosen&&<div style={{textAlign:"center",padding:24,color:"#374151",fontSize:12}}>No line items yet.</div>}
      </div>
      {lineModeChosen&&<button onClick={addDayworkLine} style={{...BP,padding:"6px 14px",fontSize:12,marginBottom:16}}>+ Add Line</button>}
      {/* Totals */}
      <div style={{background:"#0f1421",borderRadius:10,padding:"12px 16px",border:"1px solid #2d3555",maxWidth:360,marginLeft:"auto"}}>
        <div style={{display:"flex",justifyContent:"space-between",padding:"5px 0",borderBottom:"1px solid #1e2535"}}><span style={{fontSize:12,color:"#64748b"}}>Subtotal</span><span style={{fontWeight:700,color:"#e2e8f0"}}>£{subtotal.toFixed(2)}</span></div>
        {f.vatRate>0&&<div style={{display:"flex",justifyContent:"space-between",padding:"5px 0",borderBottom:"1px solid #1e2535"}}><span style={{fontSize:12,color:"#64748b"}}>VAT {isRC?"Reverse Charge":""} ({f.vatRate}%)</span><span style={{fontWeight:700,color:"#fbbf24"}}>£{vat.toFixed(2)}</span></div>}
        <div style={{display:"flex",justifyContent:"space-between",padding:"7px 0"}}><span style={{fontSize:13,color:"#e2e8f0",fontWeight:700}}>TOTAL DUE</span><span style={{fontSize:18,fontWeight:800,color:"#34d399"}}>£{total.toFixed(2)}</span></div>
        {isRC&&<div style={{fontSize:10,color:"#fbbf24",marginTop:4}}>* VAT Reverse Charge — customer to account for VAT. Total due is net amount.</div>}
      </div>
    </div>}

    {tab==="preview"&&<BrightInvoicePreview invoice={f} client={client} site={site} subtotal={subtotal} vat={vat} total={total} isRC={isRC}/>}

    <div style={{display:"flex",gap:10,justifyContent:"flex-end",marginTop:20,borderTop:"1px solid #1e2535",paddingTop:16}}>
      <button onClick={onClose} style={{padding:"8px 18px",background:"#1e2535",border:"1px solid #2d3555",borderRadius:7,color:"#94a3b8",cursor:"pointer"}}>Cancel</button>
      {tab==="preview"&&<button onClick={()=>exportBrightInvoicePDF(f,client,site,subtotal,vat,total,isRC)} style={{padding:"8px 16px",background:"#1a2535",border:"1px solid #ef4444",borderRadius:7,color:"#f87171",cursor:"pointer",fontWeight:700,fontSize:13}}>📄 Export PDF</button>}
      <button onClick={()=>onSave(f)} style={BP}>Save Invoice</button>
    </div>
  </Overlay>;
}

// ─── Invoice Preview — matches Bright Metalwork brand ────────────────────────
function BrightInvoicePreview({invoice:f,client,site,subtotal,vat,total,isRC}){
  const statusColor={draft:"#64748b",sent:"#1a56db",paid:"#166534"}[f.status]||"#64748b";
  return <div style={{background:"#fff",borderRadius:12,padding:28,color:"#1a1a2e",fontFamily:"Arial,sans-serif",fontSize:12,maxHeight:"70vh",overflowY:"auto"}}>
    {/* Header: logo left, company right */}
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:24}}>
      <div>
        <div style={{fontSize:26,fontWeight:900,color:"#111",marginBottom:6}}>INVOICE</div>
        <div style={{fontSize:11,color:"#444",lineHeight:1.6}}>
          <div style={{fontWeight:600}}>{OUR_COMPANY.name}</div>
          <div>{OUR_COMPANY.address}</div>
          <div>CRN: {OUR_COMPANY.crn}</div>
          <div>UTR: {OUR_COMPANY.utr}</div>
          <div>VAT No: {OUR_COMPANY.vatNo}</div>
        </div>
      </div>
      <div style={{textAlign:"right"}}>
        {/* Logo placeholder */}
        <div style={{background:"#1a3a5f",borderRadius:8,padding:"8px 16px",display:"inline-block",marginBottom:4}}>
          <div style={{fontSize:13,fontWeight:900,color:"#fff",letterSpacing:"0.1em"}}>BRIGHT METALWORK</div>
          <div style={{fontSize:8,color:"#60a5fa",letterSpacing:"0.15em"}}>PASSION SHAPED INTO PERFECTION</div>
        </div>
      </div>
    </div>
    {/* Grid: Bill To + Invoice Details */}
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginBottom:18}}>
      <table style={{width:"100%",borderCollapse:"collapse"}}>
        <tbody>
          <tr><td style={{background:"#1a3a5f",color:"#fff",fontWeight:700,padding:"6px 10px",fontSize:11}}>BILL TO</td></tr>
          <tr><td style={{padding:"8px 10px",fontSize:11,lineHeight:1.7,border:"1px solid #e2e8f0"}}>
            <div style={{fontWeight:600}}>{client?.name||"—"}</div>
            {client?.address&&<div style={{color:"#555",whiteSpace:"pre-wrap"}}>{client.address}</div>}
            {client?.crn&&<div style={{color:"#555"}}>CRN: {client.crn}</div>}
          </td></tr>
        </tbody>
      </table>
      <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
        <tbody>
          {[["INVOICE NO",f.invoiceNumber||"—"],["DATE",f.issueDate?new Date(f.issueDate).toLocaleDateString("en-GB"):"—"],["PO NUMBER",f.poNumber||"—"],["SITE NAME / NUMBER",(site?.name||"—")+(f.weekCommencing?"\n"+f.weekCommencing:"")],["PAYMENT TERMS",f.paymentTerms||"On Receipt"],["DUE DATE",f.dueDate?new Date(f.dueDate).toLocaleDateString("en-GB"):f.paymentTerms||"On Receipt"]].map(([l,v],i)=>
            <tr key={l}><td style={{background:"#1a3a5f",color:"#fff",fontWeight:700,padding:"5px 8px",border:"1px solid #fff"}}>{l}</td><td style={{padding:"5px 8px",border:"1px solid #e2e8f0",verticalAlign:"top",whiteSpace:"pre-wrap"}}>{v}</td></tr>
          )}
        </tbody>
      </table>
    </div>
    {/* Line items */}
    <table style={{width:"100%",borderCollapse:"collapse",marginBottom:14,fontSize:11}}>
      <thead><tr style={{background:"#1a3a5f"}}>
        {["DESCRIPTION","QTY","UNIT PRICE","AMOUNT"].map(h=><th key={h} style={{padding:"8px 10px",color:"#fff",fontWeight:700,textAlign:h!=="DESCRIPTION"?"right":"left"}}>{h}</th>)}
      </tr></thead>
      <tbody>
        {f.lines.map((l,i)=>{const lt=(l.qty||0)*(l.rate||0);return(
          <tr key={l.id} style={{background:i%2===0?"#fff":"#f8fafc"}}>
            <td style={{padding:"8px 10px",borderBottom:"1px solid #e2e8f0"}}>{l.description||"—"}</td>
            <td style={{padding:"8px 10px",textAlign:"right",borderBottom:"1px solid #e2e8f0"}}>{l.qty}.0</td>
            <td style={{padding:"8px 10px",textAlign:"right",borderBottom:"1px solid #e2e8f0"}}>£{(l.rate||0).toFixed(2)}</td>
            <td style={{padding:"8px 10px",textAlign:"right",borderBottom:"1px solid #e2e8f0",fontWeight:600}}>£{lt.toFixed(2)}</td>
          </tr>
        );})}
      </tbody>
    </table>
    {/* Totals */}
    <div style={{display:"flex",justifyContent:"flex-end",marginBottom:20}}>
      <div style={{minWidth:220}}>
        <div style={{display:"flex",justifyContent:"space-between",padding:"5px 0",borderBottom:"1px solid #e2e8f0"}}><span>Subtotal</span><span style={{fontWeight:600}}>£{subtotal.toFixed(2)}</span></div>
        {f.vatRate>0&&<div style={{display:"flex",justifyContent:"space-between",padding:"5px 0",borderBottom:"1px solid #e2e8f0"}}><span>VAT {isRC?"Reverse Charge":""}<br/><span style={{fontSize:10,color:"#64748b"}}>({f.vatRate}%)</span></span><span style={{fontWeight:600}}>£{vat.toFixed(2)}</span></div>}
        <div style={{display:"flex",justifyContent:"space-between",padding:"7px 0",borderTop:"2px solid #111"}}><span style={{fontWeight:800,fontSize:13}}>TOTAL DUE</span><span style={{fontWeight:900,fontSize:16}}>£{total.toFixed(2)}</span></div>
      </div>
    </div>
    {/* Footer */}
    <div style={{textAlign:"center",borderTop:"2px solid #e2e8f0",paddingTop:14,fontSize:11,color:"#444",lineHeight:1.9}}>
      <div style={{fontStyle:"italic",fontWeight:600,fontSize:13}}>Thank you for your business!</div>
      <div><strong>Payments to be made to Bright Metalwork LTD</strong></div>
      <div>{OUR_COMPANY.bankName}, Sort Code: {OUR_COMPANY.sortCode}, Account Number: {OUR_COMPANY.accountNo}</div>
      <div>Please use "{f.invoiceNumber}" as your payment reference</div>
      <div style={{marginTop:6}}>If you have any questions about this invoice please contact {OUR_COMPANY.contactName} at<br/><span style={{color:"#1a56db"}}>{OUR_COMPANY.email}</span><br/><span style={{color:"#1a56db"}}>{OUR_COMPANY.phone}</span></div>
    </div>
  </div>;
}

// ─── Invoice PDF Export — matches Bright Metalwork layout ────────────────────
function exportBrightInvoicePDF(inv,client,site,subtotal,vat,total,isRC){
  const html=`<!DOCTYPE html><html><head><meta charset="utf-8"/><title>Invoice ${inv.invoiceNumber||""}</title>
<style>*{margin:0;padding:0;box-sizing:border-box;}body{background:#fff;color:#111;font-family:Arial,sans-serif;font-size:11px;padding:20px 28px;}
.hdr{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:24px;}
.logo-box{background:#1a3a5f;border-radius:6px;padding:7px 14px;display:inline-block;}
.logo-name{font-size:12px;font-weight:900;color:#fff;letter-spacing:0.1em;}
.logo-sub{font-size:7px;color:#93c5fd;letter-spacing:0.15em;}
.co-details{font-size:10px;color:#444;line-height:1.65;}
.inv-title{font-size:28px;font-weight:900;color:#111;margin-bottom:8px;}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:16px;}
table.meta td{border:1px solid #e2e8f0;padding:5px 8px;vertical-align:top;white-space:pre-wrap;font-size:10px;}
table.meta td.hd{background:#1a3a5f;color:#fff;font-weight:700;border:1px solid #fff;}
table.lines{width:100%;border-collapse:collapse;margin-bottom:12px;font-size:10px;}
table.lines th{background:#1a3a5f;color:#fff;padding:7px 9px;font-weight:700;text-align:left;}
table.lines th.r{text-align:right;}
table.lines td{padding:7px 9px;border-bottom:1px solid #e2e8f0;}
table.lines td.r{text-align:right;}
table.lines tr:nth-child(even) td{background:#f8fafc;}
.totals{display:flex;justify-content:flex-end;margin-bottom:16px;}
.tot-box{width:230px;}
.tot-row{display:flex;justify-content:space-between;padding:4px 0;border-bottom:1px solid #e2e8f0;font-size:10px;}
.tot-final{border-top:2px solid #111;border-bottom:none;padding-top:7px;margin-top:4px;}
.footer{text-align:center;border-top:2px solid #e2e8f0;padding-top:12px;line-height:2;font-size:10px;color:#444;}
.footer-title{font-style:italic;font-size:12px;font-weight:700;margin-bottom:2px;}
.ft{margin-top:12px;padding-top:8px;border-top:1px solid #e2e8f0;display:flex;justify-content:space-between;font-size:8px;color:#94a3b8;}
@media print{body{padding:10px 16px;}@page{margin:8mm;size:A4;}}</style></head><body>
<div class="hdr">
  <div>
    <div class="inv-title">INVOICE</div>
    <div class="co-details">
      <strong>${OUR_COMPANY.name}</strong><br/>${OUR_COMPANY.address}<br/>
      CRN: ${OUR_COMPANY.crn}<br/>UTR: ${OUR_COMPANY.utr}<br/>VAT No: ${OUR_COMPANY.vatNo}
    </div>
  </div>
  <div class="logo-box"><div class="logo-name">BRIGHT METALWORK</div><div class="logo-sub">PASSION SHAPED INTO PERFECTION</div></div>
</div>
<div class="grid2">
  <table class="meta" style="width:100%;border-collapse:collapse">
    <tr><td class="hd">BILL TO</td></tr>
    <tr><td style="font-size:10px;line-height:1.7;border:1px solid #e2e8f0;padding:8px 10px">${client?.name||"—"}<br/>${client?.address?"<br/>"+client.address.replace(/,\s*/g,"<br/>"):""}${client?.crn?"<br/>CRN: "+client.crn:""}</td></tr>
  </table>
  <table class="meta" style="width:100%;border-collapse:collapse">
    ${[["INVOICE NO",inv.invoiceNumber||"—"],["DATE",inv.issueDate?new Date(inv.issueDate).toLocaleDateString("en-GB"):"—"],["PO NUMBER",inv.poNumber||"—"],["SITE NAME / NUMBER",(site?.name||"—")+(inv.weekCommencing?"\n"+inv.weekCommencing:"")],["PAYMENT TERMS",inv.paymentTerms||"On Receipt"],["DUE DATE",inv.dueDate?new Date(inv.dueDate).toLocaleDateString("en-GB"):inv.paymentTerms||"On Receipt"]].map(([l,v])=>`<tr><td class="hd">${l}</td><td style="border:1px solid #e2e8f0;padding:5px 8px;font-size:10px">${v}</td></tr>`).join("")}
  </table>
</div>
<table class="lines">
  <thead><tr><th>DESCRIPTION</th><th class="r">QTY</th><th class="r">UNIT PRICE</th><th class="r">AMOUNT</th></tr></thead>
  <tbody>
    ${inv.lines.map((l,i)=>{const lt=(l.qty||0)*(l.rate||0);return`<tr><td>${l.description||"—"}</td><td class="r">${l.qty}.0</td><td class="r">£${(l.rate||0).toFixed(2)}</td><td class="r" style="font-weight:600">£${lt.toFixed(2)}</td></tr>`;}).join("")}
  </tbody>
</table>
<div class="totals"><div class="tot-box">
  <div class="tot-row"><span>Subtotal</span><span>£${subtotal.toFixed(2)}</span></div>
  ${inv.vatRate>0?`<div class="tot-row"><span>VAT ${isRC?"Reverse Charge":""}<br/><small>(${inv.vatRate}%)</small></span><span>£${vat.toFixed(2)}</span></div>`:""}
  <div class="tot-row tot-final"><span style="font-weight:800;font-size:12px">TOTAL DUE</span><span style="font-weight:900;font-size:15px">£${total.toFixed(2)}</span></div>
</div></div>
<div class="footer">
  <div class="footer-title">Thank you for your business!</div>
  <strong>Payments to be made to Bright Metalwork LTD</strong><br/>
  ${OUR_COMPANY.bankName}, Sort Code: ${OUR_COMPANY.sortCode}, Account Number: ${OUR_COMPANY.accountNo}<br/>
  Please use "${inv.invoiceNumber}" as your payment reference<br/><br/>
  If you have any questions about this invoice please contact ${OUR_COMPANY.contactName} at<br/>
  <span style="color:#1a56db">${OUR_COMPANY.email}</span><br/>
  <span style="color:#1a56db">${OUR_COMPANY.phone}</span>
</div>
<div class="ft"><span>${OUR_COMPANY.name}</span><span>Invoice: ${inv.invoiceNumber||"—"}</span><span>Total: £${total.toFixed(2)}</span></div>
<script>window.onload=function(){window.print();}</script></body></html>`;
  const b=new Blob([html],{type:"text/html"});const u=URL.createObjectURL(b);
  const win=window.open(u,"_blank","width=900,height=900");
  if(!win){const a=document.createElement("a");a.href=u;a.download=`Invoice_${inv.invoiceNumber||"draft"}.html`;a.click();}
  setTimeout(()=>URL.revokeObjectURL(u),5000);
}

// ─── Invoices List View ───────────────────────────────────────────────────────
function InvoicesView({invoices,clients,allSites,scopeData,workers,onNew,onEdit,onDelete}){
  const totals={draft:0,sent:0,paid:0};
  invoices.forEach(inv=>{const{total}=calcInvoiceTotals(inv);totals[inv.status]=(totals[inv.status]||0)+total;});
  const STATUS_COLOR={draft:"#64748b",sent:"#60a5fa",paid:"#34d399"};
  return <div style={{padding:"14px 18px"}}>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16,flexWrap:"wrap",gap:10}}>
      <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10,flex:1}}>
        {[["Draft","£"+totals.draft.toFixed(2),"#64748b"],["Sent","£"+totals.sent.toFixed(2),"#60a5fa"],["Paid","£"+totals.paid.toFixed(2),"#34d399"]].map(([l,v,c])=>
          <div key={l} style={{background:"#1a1f2e",border:`1px solid ${c}44`,borderRadius:10,padding:"9px 13px"}}><div style={{fontSize:10,color:"#64748b",fontWeight:700,textTransform:"uppercase"}}>{l}</div><div style={{fontSize:18,fontWeight:800,color:c}}>{v}</div></div>
        )}
      </div>
      <button onClick={onNew} style={{...BP,padding:"8px 16px",fontSize:12,whiteSpace:"nowrap"}}>+ New Invoice</button>
    </div>
    {invoices.length===0&&<div style={{textAlign:"center",padding:40,color:"#374151",fontSize:13}}>No invoices yet.</div>}
    <div style={{display:"flex",flexDirection:"column",gap:9}}>
      {invoices.map(inv=>{
        const{subtotal,vat,total,isRC}=calcInvoiceTotals(inv);
        const client=clients.find(c=>c.id===inv.clientId);
        const site=allSites.find(s=>s.id===inv.siteId);
        const sc=STATUS_COLOR[inv.status]||"#64748b";
        return <div key={inv.id} style={{background:"#111827",border:`1px solid ${sc}33`,borderRadius:10,padding:"12px 15px",display:"flex",alignItems:"center",gap:12,flexWrap:"wrap"}}>
          <div style={{flex:1,minWidth:160}}>
            <div style={{fontWeight:700,color:"#f1f5f9",fontSize:13}}>{inv.invoiceNumber||"No number"}</div>
            <div style={{fontSize:11,color:"#64748b",marginTop:2}}>{inv.issueDate?new Date(inv.issueDate).toLocaleDateString("en-GB"):"—"}{inv.weekCommencing?` · WC: ${inv.weekCommencing}`:""}</div>
          </div>
          <div style={{fontSize:12,color:client?.color||"#64748b",fontWeight:600}}>{client?.name||"No client"}</div>
          <div style={{fontSize:11,color:"#94a3b8"}}>{site?.name||"No site"}</div>
          <div style={{marginLeft:"auto",textAlign:"right"}}>
            <div style={{fontSize:16,fontWeight:800,color:"#34d399"}}>£{total.toFixed(2)}</div>
            {isRC&&<div style={{fontSize:9,color:"#fbbf24"}}>VAT RC £{vat.toFixed(2)}</div>}
          </div>
          <span style={{display:"inline-block",padding:"2px 10px",borderRadius:20,background:sc+"22",color:sc,fontWeight:700,fontSize:11,textTransform:"uppercase",border:`1px solid ${sc}`}}>{inv.status}</span>
          <div style={{display:"flex",gap:5}}>
            <button onClick={()=>onEdit(inv)} style={{padding:"5px 10px",background:"#1e3a5f",border:"1px solid #3b82f6",borderRadius:6,color:"#60a5fa",cursor:"pointer",fontSize:11,fontWeight:700}}>Edit</button>
                    <button onClick={()=>{const c=clients.find(x=>x.id===inv.clientId);const s=allSites.find(x=>x.id===inv.siteId);openInvoiceWindow(inv,c,s);}} title="Open in new window" style={{padding:"5px 10px",background:"#1a1f2e",border:"1px solid #60a5fa",borderRadius:6,color:"#60a5fa",cursor:"pointer",fontSize:11}}>🔗 Open</button>
            <button onClick={()=>{const{subtotal:s,vat:v,total:t,isRC:r}=calcInvoiceTotals(inv);const cl=clients.find(c=>c.id===inv.clientId);const si=allSites.find(ss=>ss.id===inv.siteId);exportBrightInvoicePDF(inv,cl,si,s,v,t,r);}} style={{padding:"5px 10px",background:"#1a2535",border:"1px solid #ef4444",borderRadius:6,color:"#f87171",cursor:"pointer",fontSize:11,fontWeight:700}}>📄 PDF</button>
            <button onClick={()=>onDelete(inv.id)} style={{padding:"5px 9px",background:"#2d1515",border:"1px solid #ef4444",borderRadius:6,color:"#f87171",cursor:"pointer",fontSize:11,fontWeight:700}}>✕</button>
          </div>
        </div>;
      })}
    </div>
  </div>;
}

// ─── Training Matrix PDF Export ───────────────────────────────────────────────
function exportTrainingMatrix(workers,label,allSites,filterClient=null,filterSite=null,clients=[]){
  const cSt2=(cert,w)=>{const v=w.certs?.[cert.key];if(!v||!v.held)return"missing";if(!cert.hasExpiry||!v.expiry)return"valid";const d=(new Date(v.expiry)-new Date())/86400000;return d<0?"expired":d<30?"expiring":"valid";};
  const SC2={valid:"#22c55e",expiring:"#f59e0b",expired:"#ef4444",missing:"#d1d5db"};
  const html=`<!DOCTYPE html><html><head><meta charset="utf-8"/><title>Training Matrix — ${label}</title>
<style>*{margin:0;padding:0;box-sizing:border-box;}body{background:#fff;color:#111;font-family:Arial,sans-serif;font-size:9px;padding:16px;}
.hdr{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:16px;padding-bottom:10px;border-bottom:3px solid #1a3a5f;}
.title{font-size:20px;font-weight:900;color:#1a3a5f;}
.sub{font-size:11px;color:#555;margin-top:3px;}
.logo-box{background:#1a3a5f;border-radius:5px;padding:6px 12px;display:inline-block;}
.logo-name{font-size:11px;font-weight:900;color:#fff;letter-spacing:0.08em;}
.logo-sub{font-size:7px;color:#93c5fd;letter-spacing:0.12em;}
table{width:100%;border-collapse:collapse;font-size:8px;}
th{background:#1a3a5f;color:#fff;padding:5px 4px;text-align:center;font-size:7px;font-weight:700;border:1px solid #fff;}
th.name{text-align:left;min-width:110px;padding-left:8px;}
td{padding:4px 3px;border:1px solid #e5e7eb;text-align:center;vertical-align:middle;}
td.wname{text-align:left;padding-left:7px;font-weight:600;color:#111;min-width:110px;}
td.pos{text-align:left;padding-left:5px;color:#555;}
tr:nth-child(even) td{background:#f9fafb;}
.dot{width:10px;height:10px;border-radius:50%;display:inline-block;}
.legend{display:flex;gap:16px;margin-top:12px;font-size:9px;color:#555;}
.leg-item{display:flex;align-items:center;gap:5px;}
.ft{margin-top:10px;display:flex;justify-content:space-between;font-size:8px;color:#9ca3af;border-top:1px solid #e5e7eb;padding-top:6px;}
</style></head><body>
<div class="hdr">
  <div>
    <div class="title">Training Matrix</div>
    <div class="sub">${label}${filterClient?" · "+filterClient:""}${filterSite?" · "+filterSite:""}</div>
    <div class="sub" style="margin-top:4px;color:#9ca3af">Generated: ${new Date().toLocaleDateString("en-GB")} · ${workers.length} operative${workers.length!==1?"s":""}</div>
  </div>
  <div class="logo-box"><div class="logo-name">BRIGHT METALWORK</div><div class="logo-sub">PASSION SHAPED INTO PERFECTION</div></div>
</div>
<table>
  <thead><tr>
    <th class="name">Operative</th><th style="text-align:left;padding-left:5px">Position</th>
    ${CERTS.map(c=>`<th title="${c.label}">${c.label.split(" ").map(w=>w[0]).join("").slice(0,5)}</th>`).join("")}
    <th>Valid</th><th>Exp.Soon</th><th>Expired</th>
  </tr></thead>
  <tbody>
    ${workers.map(w=>{
      const stats=CERTS.map(c=>cSt2(c,w));
      const valid=stats.filter(s=>s==="valid").length;
      const expiring=stats.filter(s=>s==="expiring").length;
      const expired=stats.filter(s=>s==="expired").length;
      return `<tr>
        <td class="wname">${w.name||"—"}</td>
        <td class="pos">${w.position||"—"}</td>
        ${CERTS.map((c,i)=>{const s=stats[i];const col=SC2[s];const v=w.certs?.[c.key];const exp=v?.expiry&&c.hasExpiry?` title="${c.label}: ${s}${v.expiry?" · Exp: "+new Date(v.expiry).toLocaleDateString("en-GB"):""}"`:"";return `<td${exp}><span class="dot" style="background:${col}"></span></td>`;}).join("")}
        <td style="color:#22c55e;font-weight:700">${valid}</td>
        <td style="color:#f59e0b;font-weight:700">${expiring>0?expiring:""}</td>
        <td style="color:#ef4444;font-weight:700">${expired>0?expired:""}</td>
      </tr>`;
    }).join("")}
  </tbody>
</table>
<div class="legend">
  <div class="leg-item"><span class="dot" style="background:#22c55e"></span>Valid</div>
  <div class="leg-item"><span class="dot" style="background:#f59e0b"></span>Expiring &lt;30 days</div>
  <div class="leg-item"><span class="dot" style="background:#ef4444"></span>Expired</div>
  <div class="leg-item"><span class="dot" style="background:#d1d5db"></span>Not held</div>
</div>
<div class="ft"><span>Bright Metalwork Ltd · Training Matrix · ${label}</span><span>${workers.length} operative${workers.length!==1?"s":""}</span></div>
<script>window.onload=function(){window.print();}</script></body></html>`;
  const b=new Blob([html],{type:"text/html"});const u=URL.createObjectURL(b);
  const win=window.open(u,"_blank","width:1400,height:900");
  if(!win){const a=document.createElement("a");a.href=u;a.download=`TrainingMatrix_${label.replace(/\s+/g,"_")}.html`;a.click();}
  setTimeout(()=>URL.revokeObjectURL(u),5000);
}

// ─── Training Matrix Modal ────────────────────────────────────────────────────
function TrainingMatrixModal({workers,clients,allSites,activeDays,weekLabel,onClose}){
  const [mode,setMode]=useState("all"); // all | client | site
  const [clientId,setClientId]=useState(clients[0]?.id||"");
  const [siteId,setSiteId]=useState(allSites.filter(s=>!isOff(s.name))[0]?.id||"");

  const getWorkers=()=>{
    if(mode==="all") return workers;
    if(mode==="client"){
      const clientSites=allSites.filter(s=>s.clientId===clientId).map(s=>s.name);
      return workers.filter(w=>activeDays.some(d=>clientSites.some(sn=>(w.days[d]||"").includes(sn))));
    }
    if(mode==="site"){
      const site=allSites.find(s=>s.id===siteId);
      if(!site) return workers;
      return workers.filter(w=>activeDays.some(d=>(w.days[d]||"").includes(site.name)));
    }
    return workers;
  };
  const filtered=getWorkers();
  const clientName=clients.find(c=>c.id===clientId)?.name;
  const siteName=allSites.find(s=>s.id===siteId)?.name;

  return <Overlay onClose={onClose}>
    <MH title="🛡 Export Training Matrix PDF" onClose={onClose}/>
    <div style={{marginBottom:16}}>
      <label style={LBL}>Select scope</label>
      <div style={{display:"flex",gap:9,flexWrap:"wrap"}}>
        {[["all","🌐 All Active Workers"],["client","👔 By Client"],["site","📍 By Site"]].map(([v,l])=>
          <button key={v} onClick={()=>setMode(v)} style={{padding:"9px 16px",background:mode===v?"#1e3a5f":"#1a1f2e",border:`1px solid ${mode===v?"#3b82f6":"#2d3555"}`,borderRadius:8,color:mode===v?"#60a5fa":"#64748b",cursor:"pointer",fontSize:12,fontWeight:mode===v?700:400}}>{l}</button>
        )}
      </div>
    </div>
    {mode==="client"&&<FSel label="Choose Client" value={clientId} onChange={setClientId} options={clients.map(c=>({value:c.id,label:c.name}))}/>}
    {mode==="site"&&<FSel label="Choose Site" value={siteId} onChange={setSiteId} options={allSites.filter(s=>!isOff(s.name)).map(s=>({value:s.id,label:s.name}))}/>}
    <div style={{background:"#0f1421",borderRadius:8,padding:"11px 14px",marginBottom:16,border:"1px solid #1e2535"}}>
      <div style={{fontSize:11,color:"#64748b",marginBottom:4}}>Workers included in this matrix:</div>
      <div style={{fontSize:18,fontWeight:800,color:"#60a5fa"}}>{filtered.length}</div>
      <div style={{display:"flex",flexWrap:"wrap",gap:4,marginTop:7}}>
        {filtered.map(w=><span key={w.id} style={{fontSize:10,color:"#94a3b8",background:"#1a1f2e",borderRadius:4,padding:"2px 6px"}}>{w.name}</span>)}
        {filtered.length===0&&<span style={{color:"#374151",fontSize:12}}>No workers found for this selection.</span>}
      </div>
    </div>
    <div style={{display:"flex",gap:10,justifyContent:"flex-end",borderTop:"1px solid #1e2535",paddingTop:16}}>
      <button onClick={onClose} style={{padding:"8px 18px",background:"#1e2535",border:"1px solid #2d3555",borderRadius:7,color:"#94a3b8",cursor:"pointer"}}>Cancel</button>
      <button onClick={()=>{exportTrainingMatrix(filtered,`${weekLabel}`,allSites,mode==="client"?clientName:null,mode==="site"?siteName:null,clients);onClose();}} disabled={filtered.length===0} style={{...BP,opacity:filtered.length===0?0.4:1}}>📄 Export PDF ({filtered.length} workers)</button>
    </div>
  </Overlay>;
}


// ═══════════════════════════════════════════════════════════════════════════════
// FINANCIAL OVERVIEW PANEL — matches screenshot design
// ═══════════════════════════════════════════════════════════════════════════════
function FinancialOverviewPanel({invoices, clients, allSites, workers, activeDays, siteHours}){
  const [expanded, setExpanded] = useState(false);
  const allInv = invoices||[];
  const totalInvoiced = allInv.reduce((a,i)=>a+(i.amount||calcInvoiceTotals(i).total||0),0);
  const collected     = allInv.filter(i=>i.status==="paid").reduce((a,i)=>a+(i.amount||calcInvoiceTotals(i).total||0),0);
  const pending       = allInv.filter(i=>i.status==="pending").reduce((a,i)=>a+(i.amount||calcInvoiceTotals(i).total||0),0);
  const overdue       = allInv.filter(i=>i.status==="overdue").reduce((a,i)=>a+(i.amount||calcInvoiceTotals(i).total||0),0);
  const weekLabour    = useMemo(()=>workers.reduce((a,w)=>{const{gross}=calcPay(w,activeDays,siteHours);return a+gross;},0),[workers,activeDays,siteHours]);

  const fmtK = v => v>=1000?`£${(v/1000).toFixed(1)}k`:`£${Math.round(v).toLocaleString()}`;

  const KPI = [
    {label:"TOTAL INVOICED", value:totalInvoiced, color:"#3b82f6",  bar:1.0},
    {label:"COLLECTED",      value:collected,      color:"#22c55e",  bar:totalInvoiced>0?collected/totalInvoiced:0},
    {label:"PENDING",        value:pending,        color:"#f59e0b",  bar:totalInvoiced>0?pending/totalInvoiced:0},
    {label:"OVERDUE",        value:overdue,        color:"#ef4444",  bar:totalInvoiced>0?overdue/totalInvoiced:0},
  ];

  return (
    <div style={{background:"linear-gradient(135deg,#0a1a0a,#0d1f0d)",border:"1px solid #16a34a44",borderRadius:12,marginBottom:16,overflow:"hidden"}}>
      {/* Header row */}
      <div style={{display:"flex",alignItems:"center",gap:12,padding:"14px 18px",borderBottom:expanded?"1px solid #16a34a22":"none"}}>
        <div style={{width:32,height:32,background:"linear-gradient(135deg,#16a34a,#059669)",borderRadius:8,display:"flex",alignItems:"center",justifyContent:"center",fontSize:15,flexShrink:0}}>📊</div>
        <div>
          <div style={{fontSize:14,fontWeight:800,color:"#34d399",letterSpacing:"-0.01em"}}>Financial Overview</div>
          <div style={{fontSize:11,color:"#374151",marginTop:1}}>Invoice pipeline &amp; cash position</div>
        </div>
        <button onClick={()=>setExpanded(e=>!e)}
          style={{marginLeft:"auto",padding:"4px 12px",background:"#0d2218",border:"1px solid #16a34a66",borderRadius:6,color:"#34d399",cursor:"pointer",fontSize:11,fontWeight:700}}>
          {expanded?"Collapse ▲":"Expand ▼"}
        </button>
      </div>

      {/* KPI strip — always visible, matches screenshot exactly */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:0,borderBottom:expanded?"1px solid #16a34a22":"none"}}>
        {KPI.map((k,i)=>(
          <div key={k.label} style={{padding:"14px 20px",borderRight:i<3?"1px solid #16a34a18":"none"}}>
            <div style={{fontSize:10,color:"#6b7280",fontWeight:700,textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:6}}>{k.label}</div>
            <div style={{fontSize:22,fontWeight:900,color:k.color,letterSpacing:"-0.02em",marginBottom:8}}>£{Math.round(k.value).toLocaleString()}</div>
            {/* Progress bar — exactly as in screenshot */}
            <div style={{height:3,background:"#1a2e1a",borderRadius:2,overflow:"hidden"}}>
              <div style={{height:"100%",width:`${Math.min(100,Math.round(k.bar*100))}%`,background:k.color,borderRadius:2,transition:"width 0.5s"}}/>
            </div>
          </div>
        ))}
      </div>

      {/* Expanded detail */}
      {expanded&&<div style={{padding:"16px 18px"}}>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14,marginBottom:14}}>
          {/* Weekly labour */}
          <div style={{background:"#0a0e17",border:"1px solid #f8717122",borderRadius:9,padding:"12px 14px"}}>
            <div style={{fontSize:10,color:"#6b7280",fontWeight:700,textTransform:"uppercase",marginBottom:6}}>Labour Cost This Week</div>
            <div style={{fontSize:22,fontWeight:900,color:"#f87171"}}>£{Math.round(weekLabour).toLocaleString()}</div>
            <div style={{fontSize:11,color:"#374151",marginTop:3}}>{workers.length} operatives · {activeDays.length} days</div>
          </div>
          {/* Net position */}
          <div style={{background:"#0a0e17",border:`1px solid ${(collected-weekLabour)>=0?"#34d39922":"#f8717122"}`,borderRadius:9,padding:"12px 14px"}}>
            <div style={{fontSize:10,color:"#6b7280",fontWeight:700,textTransform:"uppercase",marginBottom:6}}>Collected vs Labour (Net)</div>
            <div style={{fontSize:22,fontWeight:900,color:(collected-weekLabour)>=0?"#34d399":"#f87171"}}>£{Math.abs(Math.round(collected-weekLabour)).toLocaleString()}</div>
            <div style={{fontSize:11,color:"#374151",marginTop:3}}>{(collected-weekLabour)>=0?"surplus":"deficit"}</div>
          </div>
        </div>

        {/* Per-client statement */}
        <div style={{fontSize:10,color:"#6b7280",fontWeight:700,textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:10}}>Per Client Statement</div>
        {clients.map(c=>{
          const cSites=allSites.filter(s=>s.clientId===c.id);
          const cInv=allInv.filter(i=>cSites.find(s=>s.id===i.siteId));
          const cTot=cInv.reduce((a,i)=>a+(i.amount||0),0);
          const cPaid=cInv.filter(i=>i.status==="paid").reduce((a,i)=>a+i.amount,0);
          const cPend=cInv.filter(i=>i.status==="pending").reduce((a,i)=>a+i.amount,0);
          const cOver=cInv.filter(i=>i.status==="overdue").reduce((a,i)=>a+i.amount,0);
          if(cTot===0&&cInv.length===0) return null;
          return <div key={c.id} style={{background:"#0a0e17",border:`1px solid ${c.color}22`,borderRadius:9,padding:"11px 14px",marginBottom:8}}>
            <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
              <span style={{width:8,height:8,borderRadius:"50%",background:c.color}}/>
              <span style={{fontWeight:700,color:"#f1f5f9",fontSize:13,flex:1}}>{c.name}</span>
              <span style={{fontWeight:800,color:"#34d399",fontSize:13}}>£{cTot.toLocaleString()}</span>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:6}}>
              {[["Paid",cPaid,"#34d399"],["Pending",cPend,"#fbbf24"],["Overdue",cOver,"#f87171"],["Invoices",cInv.length,"#60a5fa"]].map(([l,v,col])=>(
                <div key={l} style={{textAlign:"center"}}>
                  <div style={{fontSize:9,color:"#6b7280",textTransform:"uppercase",fontWeight:700}}>{l}</div>
                  <div style={{fontSize:13,fontWeight:800,color:col,marginTop:2}}>{l==="Invoices"?v:"£"+Math.round(v).toLocaleString()}</div>
                </div>
              ))}
            </div>
          </div>;
        })}
      </div>}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// WEEKLY RECORDS SYSTEM — saved snapshots per week
// ═══════════════════════════════════════════════════════════════════════════════

// ── Weekly Records List Page ──────────────────────────────────────────────────
function DWeeklyRecords({weeklyRecords,setWeeklyRecords,workers,allSites,clients,siteHours,activeDays,weekLabel,showWeekend,invoices,setPage,setDetailId}){
  const sorted=[...(weeklyRecords||[])].sort((a,b)=>new Date(b.savedAt)-new Date(a.savedAt));

  function saveCurrentWeek(){
    if(!window.confirm(`Save snapshot of WC ${weekLabel} to weekly records?\n\nThis creates a permanent saved record of the current week.`)) return;
    const snap={
      id:"wk_"+Date.now(),
      weekLabel,
      savedAt:new Date().toISOString(),
      workers:JSON.parse(JSON.stringify(workers)),
      allSites:JSON.parse(JSON.stringify(allSites)),
      siteHours:JSON.parse(JSON.stringify(siteHours)),
      activeDays:showWeekend?[...BASE_DAYS,...WEEKEND_DAYS]:BASE_DAYS,
      invoices:JSON.parse(JSON.stringify(invoices)),
      status:"closed",
    };
    setWeeklyRecords(recs=>{
      const existing=recs.find(r=>r.weekLabel===weekLabel);
      if(existing) return recs.map(r=>r.weekLabel===weekLabel?snap:r);
      return [...recs,snap];
    });
  }

  function deleteRecord(id){
    if(!window.confirm("Delete this weekly record permanently?")) return;
    setWeeklyRecords(recs=>recs.filter(r=>r.id!==id));
  }

  return <div>
    <DPageHdr title="📅 Weekly Records" sub={`${weeklyRecords.length} saved weeks · click any week to open full detail`}
      actions={<>
        <button onClick={saveCurrentWeek} style={{padding:"7px 14px",background:"linear-gradient(135deg,#059669,#10b981)",border:"none",borderRadius:7,color:"#fff",cursor:"pointer",fontSize:12,fontWeight:700}}>💾 Save Current Week ({weekLabel})</button>
      </>}/>
    <div style={DS.body}>
      {weeklyRecords.length===0&&<div style={{textAlign:"center",padding:"60px 24px",border:"1px dashed #1e2535",borderRadius:12}}>
        <div style={{fontSize:40,marginBottom:14}}>📅</div>
        <div style={{fontSize:16,fontWeight:700,color:"#f1f5f9",marginBottom:8}}>No Weekly Records Yet</div>
        <div style={{fontSize:13,color:"#64748b",marginBottom:20}}>Close a week to save a permanent snapshot of workers, schedule, payroll and sites.</div>
        <button onClick={saveCurrentWeek} style={{padding:"10px 24px",background:"linear-gradient(135deg,#059669,#10b981)",border:"none",borderRadius:8,color:"#fff",cursor:"pointer",fontSize:13,fontWeight:700}}>💾 Save Current Week ({weekLabel})</button>
      </div>}
      <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:12}}>
        {sorted.map(rec=>{
          const labourTotal=rec.workers.reduce((a,w)=>{const{gross}=calcPay(w,rec.activeDays||BASE_DAYS,rec.siteHours||{});return a+gross;},0);
          const workingWorkers=rec.workers.filter(w=>(rec.activeDays||BASE_DAYS).some(d=>w.days?.[d]&&!isOff(w.days[d]))).length;
          const usedSites=[...new Set(rec.workers.flatMap(w=>(rec.activeDays||BASE_DAYS).map(d=>w.days?.[d]||"").filter(s=>s&&!isOff(s))))];
          return <div key={rec.id}
            style={{background:"#111827",border:"1px solid #1e2535",borderRadius:12,padding:16,cursor:"pointer",transition:"all 0.15s",position:"relative"}}
            onClick={()=>{setDetailId(rec.id);setPage("weekly_record_detail");}}
            onMouseEnter={e=>{e.currentTarget.style.borderColor="#3b82f6";e.currentTarget.style.transform="translateY(-2px)";}}
            onMouseLeave={e=>{e.currentTarget.style.borderColor="#1e2535";e.currentTarget.style.transform="";}}>
            <div style={{position:"absolute",top:0,left:0,right:0,height:2,background:"linear-gradient(90deg,#3b82f6,#6366f1)",borderRadius:"12px 12px 0 0"}}/>
            <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",marginBottom:12}}>
              <div>
                <div style={{fontSize:15,fontWeight:800,color:"#f1f5f9"}}>WC: {rec.weekLabel}</div>
                <div style={{fontSize:11,color:"#64748b",marginTop:2}}>Saved: {new Date(rec.savedAt).toLocaleDateString("en-GB",{day:"2-digit",month:"short",year:"numeric",hour:"2-digit",minute:"2-digit"})}</div>
              </div>
              <div style={{display:"flex",gap:5}}>
                <span style={{padding:"2px 9px",background:"#0d2218",border:"1px solid #10b981",borderRadius:5,fontSize:10,fontWeight:700,color:"#34d399",textTransform:"uppercase"}}>✓ Closed</span>
                <button onClick={e=>{e.stopPropagation();deleteRecord(rec.id);}} style={{padding:"2px 7px",background:"#2d1515",border:"1px solid #ef4444",borderRadius:5,color:"#f87171",cursor:"pointer",fontSize:10,marginLeft:4}}>✕</button>
              </div>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:8,marginBottom:10}}>
              {[["Workers",rec.workers.length,"#60a5fa"],["Working",workingWorkers,"#34d399"],["Sites",usedSites.length,"#f59e0b"],["Labour","£"+Math.round(labourTotal).toLocaleString(),"#a78bfa"]].map(([l,v,c])=>(
                <div key={l} style={{background:"#0f1421",borderRadius:7,padding:"7px 9px"}}>
                  <div style={{fontSize:9,color:"#64748b",textTransform:"uppercase",fontWeight:700}}>{l}</div>
                  <div style={{fontSize:13,fontWeight:800,color:c,marginTop:2}}>{v}</div>
                </div>
              ))}
            </div>
            <div style={{display:"flex",flexWrap:"wrap",gap:4}}>
              {usedSites.slice(0,5).map(s=>{const site=rec.allSites?.find(x=>s===x.name||s.includes(x.name));const col=site?.color||getSiteColor(s,rec.allSites||[]);return<span key={s} style={{fontSize:9,padding:"2px 7px",borderRadius:4,border:`1px solid ${col}44`,background:`${col}12`,color:col}}>{s}</span>;})}
              {usedSites.length>5&&<span style={{fontSize:9,color:"#374151"}}>+{usedSites.length-5} more</span>}
            </div>
          </div>;
        })}
      </div>
    </div>
  </div>;
}

// ── Weekly Record Detail Page ─────────────────────────────────────────────────
function DWeeklyRecordDetail({weeklyRecords,recordId,setPage,allSites}){
  const [tab,setTab]=useState("schedule");
  const rec=weeklyRecords.find(r=>r.id===recordId);
  if(!rec) return <div style={DS.body}><div style={{color:"#374151",padding:40,textAlign:"center"}}>Record not found.</div></div>;
  
  const days=rec.activeDays||BASE_DAYS;
  const payRows=rec.workers.map(w=>({...w,...calcPay(w,days,rec.siteHours||{})}));
  const labourTotal=payRows.reduce((a,r)=>a+r.gross,0);
  const labourNet=payRows.reduce((a,r)=>a+r.net,0);
  const workingW=payRows.filter(r=>r.stdH>0||r.otH>0);
  const usedSites=[...new Set(rec.workers.flatMap(w=>days.map(d=>w.days?.[d]||"").filter(s=>s&&!isOff(s))))];

  function exportWeekExcel(){
    const wb=XLSX.utils.book_new();
    // Schedule sheet
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([
      [`SCHEDULE — WC: ${rec.weekLabel} (ARCHIVED)`],
      ["Name","Company","Position",...days,"Scope","Rate"],
      ...rec.workers.map(w=>[w.name,w.company,w.position,...days.map(d=>w.days?.[d]||""),w.scope,w.agreedRate||""])
    ]),"Schedule");
    // Payroll sheet
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([
      [`PAYROLL — WC: ${rec.weekLabel} (ARCHIVED)`],
      ["Name","Rate","Tax%","Std h","OT h","Gross","Tax","Net"],
      ...payRows.map(r=>[r.name,r.agreedRate||"",Math.round((r.taxRate||0)*100)+"%",r.stdH,r.otH,+r.gross.toFixed(2),+r.tax.toFixed(2),+r.net.toFixed(2)]),
      ["","TOTALS","",workingW.reduce((a,r)=>a+r.stdH,0),workingW.reduce((a,r)=>a+r.otH,0),+labourTotal.toFixed(2),"",+labourNet.toFixed(2)]
    ]),"Payroll");
    XLSX.writeFile(wb,`WeeklyRecord_WC_${rec.weekLabel.replace(/\s+/g,"_")}.xlsx`);
  }

  return <div>
    <DPageHdr title={`📅 WC: ${rec.weekLabel}`}
      sub={`Archived · Saved ${new Date(rec.savedAt).toLocaleDateString("en-GB",{day:"2-digit",month:"short",year:"numeric",hour:"2-digit",minute:"2-digit"})}`}
      back="Weekly Records" onBack={()=>setPage("weekly_records")}
      actions={<>
        <button onClick={exportWeekExcel} style={{padding:"6px 13px",background:"linear-gradient(135deg,#059669,#10b981)",border:"none",borderRadius:7,color:"#fff",cursor:"pointer",fontSize:12,fontWeight:700}}>⬇ Export Excel</button>
        <button onClick={()=>exportSchedulePDF(rec.workers,days,rec.weekLabel,rec.allSites||allSites)} style={{padding:"6px 11px",background:"#1a1f2e",border:"1px solid #ef4444",borderRadius:7,color:"#f87171",cursor:"pointer",fontSize:11,fontWeight:700}}>📄 PDF</button>
      </>}/>
    <div style={DS.body}>
      {/* Stats */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:10,marginBottom:18}}>
        <DStat label="Workers" value={rec.workers.length} color="#60a5fa"/>
        <DStat label="Working" value={workingW.length} color="#34d399"/>
        <DStat label="Sites" value={usedSites.length} color="#f59e0b"/>
        <DStat label="Gross Labour" value={"£"+Math.round(labourTotal).toLocaleString()} color="#f87171"/>
        <DStat label="Net Pay" value={"£"+Math.round(labourNet).toLocaleString()} color="#a78bfa"/>
      </div>
      {/* Tabs */}
      <div style={{display:"flex",gap:3,background:"#0a0e17",borderRadius:8,padding:3,marginBottom:16,width:"fit-content"}}>
        {[["schedule","📋 Schedule"],["timesheets","⏱ Timesheets"],["payslips","💷 Payslips"]].map(([v,l])=>(
          <button key={v} onClick={()=>setTab(v)} style={{padding:"6px 14px",background:tab===v?"#1e3a5f":"transparent",border:tab===v?"1px solid #3b82f6":"1px solid transparent",borderRadius:6,color:tab===v?"#60a5fa":"#64748b",cursor:"pointer",fontSize:12,fontWeight:tab===v?700:400}}>{l}</button>
        ))}
      </div>

      {/* ── Schedule tab ── */}
      {tab==="schedule"&&<div>
        <div style={{fontSize:11,color:"#64748b",fontWeight:700,textTransform:"uppercase",marginBottom:10}}>Worker Allocation — WC {rec.weekLabel}</div>
        {(()=>{
          const grp={};
          rec.workers.forEach(w=>{
            const cnts={};days.forEach(d=>{const s=w.days?.[d];if(s&&!isOff(s))cnts[s]=(cnts[s]||0)+1;});
            const primary=Object.entries(cnts).sort((a,b)=>b[1]-a[1])[0]?.[0]||"Unassigned";
            if(!grp[primary])grp[primary]=[];grp[primary].push(w);
          });
          const sitesList=rec.allSites||allSites;
          return Object.keys(grp).sort().map(siteName=>{
            const col=getSiteColor(siteName,sitesList);
            return <div key={siteName} style={{marginBottom:12}}>
              <div style={{background:`${col}15`,borderLeft:`3px solid ${col}`,padding:"6px 12px",display:"flex",alignItems:"center",gap:8,marginBottom:0,borderRadius:"6px 6px 0 0"}}>
                <span style={{width:8,height:8,borderRadius:"50%",background:col}}/><span style={{fontWeight:700,color:col,fontSize:12}}>{siteName}</span>
                <span style={{fontSize:10,color:"#64748b"}}>{grp[siteName].length} operative{grp[siteName].length!==1?"s":""}</span>
              </div>
              <div style={{border:`1px solid ${col}33`,borderTop:"none",borderRadius:"0 0 6px 6px",overflow:"hidden"}}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
                  <thead><tr><th style={{...DS.th,paddingLeft:16}}>Worker</th><th style={DS.th}>Position</th>{days.map(d=><th key={d} style={{...DS.th,minWidth:100}}>{d}</th>)}<th style={DS.th}>Rate</th></tr></thead>
                  <tbody>{grp[siteName].map((w,i)=><tr key={w.id} style={{background:i%2===0?"#111827":"#0f1421"}}>
                    <td style={{...DS.td,fontWeight:600,color:"#f1f5f9",paddingLeft:16}}>{w.name}</td>
                    <td style={{...DS.td,color:"#94a3b8"}}>{w.position}</td>
                    {days.map(d=>{const s=w.days?.[d];const sc=s?getSiteColor(s,sitesList):"#374151";return<td key={d} style={DS.td}>{s&&!isOff(s)?<span style={{display:"inline-block",padding:"1px 6px",borderRadius:3,fontSize:9,fontWeight:600,background:`${sc}22`,color:sc,border:`1px solid ${sc}44`}}>{s}</span>:<span style={{color:"#374151",fontSize:10}}>—</span>}</td>;})}
                    <td style={{...DS.td,color:"#34d399",fontWeight:600}}>{w.agreedRate?`£${w.agreedRate}/hr`:"—"}</td>
                  </tr>)}</tbody>
                </table>
              </div>
            </div>;
          });
        })()}
      </div>}

      {/* ── Timesheets tab ── */}
      {tab==="timesheets"&&<div>
        <div style={{fontSize:11,color:"#64748b",fontWeight:700,textTransform:"uppercase",marginBottom:10}}>Timesheet Record — WC {rec.weekLabel}</div>
        <DTable cols={[
          {key:"name",label:"Worker",w:180,r:(v,r)=><div style={{display:"flex",alignItems:"center",gap:8}}><div style={{width:24,height:24,borderRadius:5,background:"#3b82f622",display:"flex",alignItems:"center",justifyContent:"center",fontSize:10,fontWeight:700,color:"#60a5fa"}}>{v[0]}</div><div><div style={{fontWeight:600,color:"#f1f5f9",fontSize:12}}>{v}</div><div style={{fontSize:10,color:"#64748b"}}>{r.position}</div></div></div>},
          {key:"position",label:"Position",r:v=><span style={{fontSize:11,color:"#94a3b8"}}>{v}</span>},
          {key:"id",label:"Days Worked",r:(_,r)=>{
            const daysWorked=days.filter(d=>r.days?.[d]&&!isOff(r.days[d])).length;
            return <span style={{color:"#60a5fa",fontWeight:700}}>{daysWorked}/{days.length}d</span>;
          }},
          {key:"stdH",label:"Std Hours",r:(v,r)=>{const{stdH}=calcPay(r,days,rec.siteHours||{});return<span style={{color:"#34d399",fontWeight:700}}>{stdH}h</span>;}},
          {key:"otH",label:"OT Hours",r:(v,r)=>{const{otH}=calcPay(r,days,rec.siteHours||{});return otH>0?<span style={{color:"#fbbf24",fontWeight:700}}>{otH}h</span>:<span style={{color:"#374151"}}>—</span>;}},
          {key:"agreedRate",label:"Rate",r:v=>v?<span style={{color:"#34d399",fontWeight:600}}>£{v}/hr</span>:<span style={{color:"#374151"}}>—</span>},
          {key:"id2",label:"Site(s)",r:(_,r)=>{
            const sites=[...new Set(days.map(d=>r.days?.[d]||"").filter(s=>s&&!isOff(s)))];
            return <div style={{display:"flex",flexWrap:"wrap",gap:3}}>{sites.slice(0,2).map(s=>{const col=getSiteColor(s,rec.allSites||allSites);return<span key={s} style={{fontSize:9,padding:"1px 5px",borderRadius:3,background:`${col}18`,color:col,border:`1px solid ${col}33`}}>{s}</span>;})}</div>;
          }},
        ]} rows={rec.workers}/>
      </div>}

      {/* ── Payslips tab ── */}
      {tab==="payslips"&&<div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10,marginBottom:14}}>
          <DStat label="Total Gross" value={"£"+Math.round(labourTotal).toLocaleString()} color="#34d399"/>
          <DStat label="Total Tax" value={"£"+Math.round(payRows.reduce((a,r)=>a+r.tax,0)).toLocaleString()} color="#f87171"/>
          <DStat label="Total Net" value={"£"+Math.round(labourNet).toLocaleString()} color="#a78bfa"/>
          <DStat label="Payslips" value={workingW.length} color="#60a5fa"/>
        </div>
        <DTable cols={[
          {key:"name",label:"Worker",w:180,r:(v,r)=><div style={{display:"flex",alignItems:"center",gap:8}}><div style={{width:24,height:24,borderRadius:5,background:"#3b82f622",display:"flex",alignItems:"center",justifyContent:"center",fontSize:10,fontWeight:700,color:"#60a5fa"}}>{v[0]}</div><span style={{fontWeight:600,color:"#f1f5f9",fontSize:12}}>{v}</span></div>},
          {key:"agreedRate",label:"Rate",r:(v,r)=>{const pr=payRows.find(x=>x.id===r.id);return<span style={{color:"#34d399",fontWeight:600}}>{pr?.agreedRate?`£${pr.agreedRate}/hr`:"—"}</span>;}},
          {key:"stdH_v",label:"Std h",r:(_,r)=>{const pr=payRows.find(x=>x.id===r.id);return<span style={{color:"#60a5fa",fontWeight:600}}>{pr?.stdH||0}h</span>;}},
          {key:"otH_v",label:"OT h",r:(_,r)=>{const pr=payRows.find(x=>x.id===r.id);return pr?.otH>0?<span style={{color:"#fbbf24",fontWeight:600}}>{pr.otH}h</span>:<span style={{color:"#374151"}}>—</span>;}},
          {key:"gross",label:"Gross",r:(_,r)=>{const pr=payRows.find(x=>x.id===r.id);return<span style={{color:"#34d399",fontWeight:700}}>£{(pr?.gross||0).toFixed(2)}</span>;}},
          {key:"tax",label:"Tax",r:(_,r)=>{const pr=payRows.find(x=>x.id===r.id);return<span style={{color:"#f87171"}}>£{(pr?.tax||0).toFixed(2)}</span>;}},
          {key:"net",label:"Net Pay",r:(_,r)=>{const pr=payRows.find(x=>x.id===r.id);return<span style={{color:"#a78bfa",fontWeight:800,fontSize:13}}>£{(pr?.net||0).toFixed(2)}</span>;}},
          {key:"id_act",label:"Export",r:(_,r)=><button onClick={()=>exportPayslip(r,days,rec.weekLabel,rec.siteHours||{})} style={{padding:"3px 8px",background:"#0d2218",border:"1px solid #10b981",borderRadius:5,color:"#34d399",cursor:"pointer",fontSize:10,fontWeight:700}}>💷 Payslip</button>},
        ]} rows={rec.workers.filter(w=>payRows.find(pr=>pr.id===w.id&&(pr.stdH>0||pr.otH>0)))}/>
      </div>}
    </div>
  </div>;
}


// ─── Financial Dashboard ──────────────────────────────────────────────────────
function FinancialDashboard({workers,clients,allSites,activeDays,siteHours,scopeData,invoices}){
  // Compute per-site: income, budget, labour cost
  const labourBySite=useMemo(()=>{
    const m={};
    workers.forEach(w=>{const{bd}=calcPay(w,activeDays,siteHours);Object.values(bd).forEach(({site,gross})=>{if(!m[site])m[site]={labour:0,wIds:new Set()};m[site].labour+=gross;m[site].wIds.add(w.id);});});
    return m;
  },[workers,activeDays,siteHours]);

  const siteRows=useMemo(()=>allSites.filter(s=>!isOff(s.name)).map(site=>{
    const items=scopeData[site.id]||[];
    const income=items.reduce((a,it)=>a+(it.qty||0)*(it.unitIncome||0),0);
    const budget=items.reduce((a,it)=>{const i=(it.qty||0)*(it.unitIncome||0);return a+i*(1+(it.overheadPct||0)/100);},0);
    const labour=labourBySite[site.name]?.labour||0;
    const client=clients.find(c=>c.id===site.clientId);
    const invoiced=invoices.filter(inv=>inv.siteId===site.id).reduce((a,inv)=>{const{total}=calcInvoiceTotals(inv);return a+total;},0);
    const paidInvoices=invoices.filter(inv=>inv.siteId===site.id&&inv.status==="paid").reduce((a,inv)=>{const{total}=calcInvoiceTotals(inv);return a+total;},0);
    return{site,income,budget,labour,client,invoiced,paidInvoices,margin:income>0?((income-labour)/income*100):null,profitGross:income-labour};
  }),[allSites,scopeData,labourBySite,clients,invoices]);

  const active=siteRows.filter(r=>r.income>0||r.labour>0);
  const grand=active.reduce((a,r)=>({income:a.income+r.income,budget:a.budget+r.budget,labour:a.labour+r.labour,invoiced:a.invoiced+r.invoiced,paid:a.paid+r.paidInvoices}),{income:0,budget:0,labour:0,invoiced:0,paid:0});

  // Client-level aggregation
  const clientRollup=useMemo(()=>{
    const m={};
    active.forEach(r=>{const cid=r.client?.id||"__none";
      if(!m[cid])m[cid]={client:r.client||{id:"__none",name:"Unassigned",color:"#374151"},income:0,labour:0,invoiced:0,paid:0,sites:[]};
      m[cid].income+=r.income;m[cid].labour+=r.labour;m[cid].invoiced+=r.invoiced;m[cid].paid+=r.paidInvoices;m[cid].sites.push(r);
    });
    return Object.values(m).filter(c=>c.income>0||c.labour>0).sort((a,b)=>b.income-a.income);
  },[active]);

  const C={income:"#34d399",budget:"#60a5fa",labour:"#f87171",margin:"#a78bfa",inv:"#fbbf24",paid:"#10b981"};
  const bar=(val,max,color,height=8)=><div style={{height,background:"#1e2535",borderRadius:4,overflow:"hidden",flex:1}}><div style={{height:"100%",borderRadius:4,background:color,width:max>0?Math.min(100,(val/max*100))+"%":"0%",transition:"width 0.4s"}}/></div>;

  return <div style={{padding:"14px 18px"}}>
    {/* Top KPI strip */}
    <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:10,marginBottom:20}}>
      {[["Total Income",grand.income,C.income,"£"],["Total Budget",grand.budget,C.budget,"£"],["Labour Cost",grand.labour,C.labour,"£"],["Invoiced",grand.invoiced,C.inv,"£"],["Collected",grand.paid,C.paid,"£"]].map(([l,v,c,pre])=>
        <div key={l} style={{background:"#1a1f2e",border:`1px solid ${c}44`,borderRadius:10,padding:"11px 14px"}}>
          <div style={{fontSize:9,color:"#64748b",fontWeight:700,textTransform:"uppercase",marginBottom:4}}>{l}</div>
          <div style={{fontSize:18,fontWeight:800,color:c}}>{pre}£{(v||0).toFixed(0).replace(/\B(?=(\d{3})+(?!\d))/g,",")}</div>
          {bar(v,grand.income||1,c,5)}
        </div>
      )}
    </div>

    {/* Client-level P&L */}
    <div style={{fontWeight:700,color:"#94a3b8",fontSize:12,marginBottom:10,textTransform:"uppercase",letterSpacing:"0.06em"}}>Client P&L</div>
    {clientRollup.map(cr=>{
      const margin=cr.income>0?((cr.income-cr.labour)/cr.income*100):null;
      const over=cr.labour>cr.income&&cr.income>0;
      return <div key={cr.client.id} style={{background:"#111827",border:`1px solid ${cr.client.color}44`,borderRadius:11,padding:14,marginBottom:12}}>
        <div style={{display:"flex",alignItems:"center",gap:9,marginBottom:10,flexWrap:"wrap"}}>
          <span style={{width:11,height:11,borderRadius:"50%",background:cr.client.color}}/>
          <span style={{fontWeight:800,color:"#f1f5f9",fontSize:14,flex:1}}>{cr.client.name}</span>
          {margin!==null&&<span style={{fontSize:12,fontWeight:700,color:over?"#f87171":margin>=20?"#34d399":"#fbbf24",padding:"2px 8px",borderRadius:5,background:over?"#2d1515":margin>=20?"#0d2218":"#2d2008"}}>{over?"⚠️ Over Budget":`${margin.toFixed(1)}% margin`}</span>}
        </div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:8,marginBottom:10}}>
          {[["Income",cr.income,C.income],["Labour Cost",cr.labour,C.labour],["Invoiced",cr.invoiced,C.inv],["Paid",cr.paid,C.paid]].map(([l,v,c])=>
            <div key={l} style={{background:"#0f1421",borderRadius:7,padding:"7px 10px",border:`1px solid ${c}22`}}>
              <div style={{fontSize:9,color:"#64748b",fontWeight:700,textTransform:"uppercase"}}>{l}</div>
              <div style={{fontSize:14,fontWeight:800,color:c}}>£{(v||0).toFixed(0)}</div>
            </div>
          )}
        </div>
        {/* Income vs Labour bar chart per site */}
        <div style={{display:"flex",flexDirection:"column",gap:5}}>
          {cr.sites.map(r=>{
            const siteMax=Math.max(r.income,r.labour)||1;
            const sOver=r.labour>r.income&&r.income>0;
            return <div key={r.site.id} style={{display:"flex",alignItems:"center",gap:8}}>
              <span style={{fontSize:10,color:r.site.color,fontWeight:600,minWidth:130,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}} title={r.site.name}>{r.site.name}</span>
              <div style={{flex:1,display:"flex",flexDirection:"column",gap:2}}>
                <div style={{display:"flex",alignItems:"center",gap:5}}>
                  {bar(r.income,siteMax,C.income,6)}
                  <span style={{fontSize:9,color:C.income,minWidth:55,textAlign:"right"}}>£{r.income.toFixed(0)}</span>
                </div>
                <div style={{display:"flex",alignItems:"center",gap:5}}>
                  {bar(r.labour,siteMax,sOver?"#f87171":C.labour,6)}
                  <span style={{fontSize:9,color:sOver?"#f87171":C.labour,minWidth:55,textAlign:"right"}}>£{r.labour.toFixed(0)}</span>
                </div>
              </div>
              {sOver&&<span style={{fontSize:10,color:"#f87171",fontWeight:700}}>⚠️</span>}
            </div>;
          })}
        </div>
        <div style={{display:"flex",gap:10,marginTop:5}}>
          <div style={{display:"flex",alignItems:"center",gap:4,fontSize:9,color:"#64748b"}}><div style={{width:10,height:4,borderRadius:2,background:C.income}}/> Income</div>
          <div style={{display:"flex",alignItems:"center",gap:4,fontSize:9,color:"#64748b"}}><div style={{width:10,height:4,borderRadius:2,background:C.labour}}/> Labour</div>
        </div>
      </div>;
    })}
    {active.length===0&&<div style={{textAlign:"center",padding:40,color:"#374151"}}>No financial data yet. Add scope of works and workers to sites.</div>}
  </div>;
}

// ═══════════════════════════════════════════════════════════════════════════════
// DASHBOARD VIEW — reads from the same live state as the schedule app
// No separate data, no duplicate saving. Pure read + navigate layer.
// ═══════════════════════════════════════════════════════════════════════════════

// ── Shared dashboard styles ───────────────────────────────────────────────────
const DS={
  sidebar:{width:210,background:"#0a0e17",borderRight:"1px solid #1e2535",height:"calc(100vh - 60px)",position:"sticky",top:60,flexShrink:0,overflowY:"auto",display:"flex",flexDirection:"column"},
  card:(color)=>({background:"#111827",border:`1px solid ${color||"#1e2535"}33`,borderRadius:12,padding:18,cursor:"pointer",transition:"all 0.15s",position:"relative",overflow:"hidden"}),
  th:{padding:"8px 12px",textAlign:"left",fontSize:11,fontWeight:700,color:"#64748b",textTransform:"uppercase",letterSpacing:"0.07em",borderBottom:"1px solid #1e2535",background:"#0a0e17",whiteSpace:"nowrap"},
  td:{padding:"8px 12px",borderBottom:"1px solid #1a2030",verticalAlign:"middle",fontSize:13},
  pill:(color)=>({display:"inline-flex",alignItems:"center",gap:4,padding:"2px 9px",borderRadius:20,border:`1px solid ${color}44`,background:`${color}15`,fontSize:11,color:color,fontWeight:600}),
  badge:(c,bg)=>({display:"inline-block",padding:"2px 8px",borderRadius:4,fontSize:11,fontWeight:600,color:c||"#fff",background:bg||"#1e2535",whiteSpace:"nowrap"}),
  hdr:{padding:"18px 24px",borderBottom:"1px solid #1e2535",background:"#0d1117",display:"flex",alignItems:"center",justifyContent:"space-between",minHeight:60},
  body:{padding:"22px 24px"},
};

const DASH_NAV=[
  // ── Overview
  {id:"home",           icon:"🏠", label:"Overview",         group:"main"},
  // ── People & Labour
  {id:"workers",        icon:"👷", label:"Workers",           group:"labour"},
  {id:"schedule",       icon:"📋", label:"Labour Schedule",   group:"labour"},
  {id:"site_by_site",   icon:"📍", label:"By Site",           group:"labour"},
  {id:"payslips",       icon:"💷", label:"Payroll & Payslips",group:"labour"},
  {id:"timesheets",     icon:"⏱", label:"Timesheets",         group:"labour"},
  {id:"weekly_records", icon:"📅", label:"Weekly Records",    group:"labour"},
  // ── Projects
  {id:"sites",          icon:"🏗", label:"Sites",             group:"projects"},
  {id:"clients",        icon:"👔", label:"Clients",           group:"projects"},
  {id:"invoices",       icon:"🧾", label:"Invoices",          group:"projects"},
  {id:"payapps",        icon:"📐", label:"Payment Apps",      group:"projects"},
  {id:"budget",         icon:"💰", label:"Budget",            group:"projects"},
  // ── Analysis
  {id:"certs",          icon:"🛡", label:"Certificates",      group:"analysis"},
  {id:"finance",        icon:"📊", label:"Finance",           group:"analysis"},
  {id:"stats",          icon:"🔢", label:"Stats",             group:"analysis"},
  {id:"bank",           icon:"🏦", label:"Bank Import",       group:"analysis"},
];

function DStat({label,value,color,sub}){
  return <div style={{background:"#111827",border:`1px solid ${color||"#1e2535"}22`,borderRadius:10,padding:"12px 15px"}}>
    <div style={{fontSize:10,color:"#64748b",fontWeight:700,textTransform:"uppercase",letterSpacing:"0.06em"}}>{label}</div>
    <div style={{fontSize:20,fontWeight:800,color:color||"#60a5fa",marginTop:4,lineHeight:1}}>{value}</div>
    {sub&&<div style={{fontSize:10,color:"#374151",marginTop:3}}>{sub}</div>}
  </div>;
}

function DTable({cols,rows,onRow}){
  return <div style={{border:"1px solid #1e2535",borderRadius:10,overflow:"hidden"}}>
    <table style={{width:"100%",borderCollapse:"collapse"}}>
      <thead><tr>{cols.map(c=><th key={c.key} style={{...DS.th,minWidth:c.w||90}}>{c.label}</th>)}</tr></thead>
      <tbody>{rows.map((r,i)=>(
        <tr key={r.id||i} onClick={()=>onRow&&onRow(r)}
          style={{background:i%2===0?"#111827":"#0f1421",cursor:onRow?"pointer":"default"}}
          onMouseEnter={e=>{if(onRow)e.currentTarget.style.background="#1a2535";}}
          onMouseLeave={e=>{e.currentTarget.style.background=i%2===0?"#111827":"#0f1421";}}>
          {cols.map(c=><td key={c.key} style={{...DS.td,...(c.style||{})}}>{c.r?c.r(r[c.key],r):r[c.key]}</td>)}
        </tr>
      ))}
      {rows.length===0&&<tr><td colSpan={cols.length} style={{...DS.td,textAlign:"center",color:"#374151",padding:28}}>No records.</td></tr>}
      </tbody>
    </table>
  </div>;
}

function DPageHdr({title,sub,back,onBack,actions}){
  return <div style={DS.hdr}>
    <div>
      {back&&<div onClick={onBack} style={{display:"flex",alignItems:"center",gap:6,color:"#64748b",cursor:"pointer",fontSize:12,marginBottom:6,userSelect:"none"}}>
        <span style={{fontSize:14}}>←</span><span>Back to {back}</span>
      </div>}
      <div style={{fontSize:18,fontWeight:800,color:"#f1f5f9",letterSpacing:"-0.02em"}}>{title}</div>
      {sub&&<div style={{fontSize:12,color:"#64748b",marginTop:2}}>{sub}</div>}
    </div>
    {actions&&<div style={{display:"flex",gap:8}}>{actions}</div>}
  </div>;
}

function DStatusBadge({status}){
  const m={paid:["#34d399","#0d2218"],pending:["#fbbf24","#1a1500"],draft:["#94a3b8","#1e2535"],submitted:["#60a5fa","#0d1a2e"],approved:["#34d399","#0d2218"],rejected:["#f87171","#2d1515"],issued:["#a78bfa","#1a0d2e"],outstanding:["#fbbf24","#1a1500"],addition:["#34d399","#0d2218"],omission:["#f87171","#2d1515"]};
  const[c,bg]=m[status]||["#94a3b8","#1e2535"];
  return <span style={{...DS.badge(c,bg),textTransform:"capitalize"}}>{status}</span>;
}

// ── Dashboard Sidebar ─────────────────────────────────────────────────────────
function DashSidebar({page,setPage,workers,allSites,clients,invoices,setModal,activeDays,siteHours,weekLabel}){
  const expiring=workers.flatMap(w=>Object.values(w.certs||{}).filter(c=>{if(!c.held||!c.expiry)return false;const d=(new Date(c.expiry)-new Date())/86400000;return d>=0&&d<30;})).length;
  const badges={certs:expiring,invoices:invoices.filter(i=>i.status==="pending").length};
  const isActive=(id)=>page===id||page.startsWith(id+"_");

  return <div style={DS.sidebar}>
    <div style={{padding:"12px 10px",flex:1}}>
      {DASH_NAV.map(item=>{
        const active=isActive(item.id);
        const cnt=badges[item.id];
        return <div key={item.id} onClick={()=>setPage(item.id)}
          style={{display:"flex",alignItems:"center",gap:9,padding:"8px 10px",borderRadius:7,marginBottom:2,cursor:"pointer",background:active?"#1e3a5f":"transparent",border:active?"1px solid #3b82f6":"1px solid transparent",transition:"all 0.12s"}}
          onMouseEnter={e=>{if(!active)e.currentTarget.style.background="#1a1f2e";}}
          onMouseLeave={e=>{if(!active)e.currentTarget.style.background="transparent";}}>
          <span style={{fontSize:15,width:18,textAlign:"center",flexShrink:0}}>{item.icon}</span>
          <span style={{flex:1,fontSize:12,fontWeight:active?700:400,color:active?"#60a5fa":"#94a3b8"}}>{item.label}</span>
          {cnt>0&&<span style={{fontSize:10,fontWeight:700,color:"#fbbf24",background:"#1a1500",padding:"1px 5px",borderRadius:9,minWidth:18,textAlign:"center"}}>{cnt}</span>}
        </div>;
      })}
    </div>
    <div style={{padding:"10px 12px",borderTop:"1px solid #1e2535"}}>
      <button onClick={()=>setModal({type:"worker",worker:mkW()})} style={{width:"100%",padding:"7px 10px",background:"linear-gradient(135deg,#3b82f6,#6366f1)",border:"none",borderRadius:7,color:"#fff",cursor:"pointer",fontSize:12,fontWeight:700,marginBottom:6}}>+ Add Worker</button>
      <button onClick={()=>setModal({type:"sites"})} style={{width:"100%",padding:"7px 10px",background:"#1e2535",border:"1px solid #2d3555",borderRadius:7,color:"#94a3b8",cursor:"pointer",fontSize:12,marginBottom:6}}>🏗 Manage Sites</button>
      <button onClick={()=>setModal({type:"clients"})} style={{width:"100%",padding:"7px 10px",background:"#1e2535",border:"1px solid #2d3555",borderRadius:7,color:"#94a3b8",cursor:"pointer",fontSize:12,marginBottom:6}}>👔 Manage Clients</button>
      <button onClick={()=>setModal({type:"trainingMatrix"})} style={{width:"100%",padding:"7px 10px",background:"#1e2535",border:"1px solid #10b981",borderRadius:7,color:"#34d399",cursor:"pointer",fontSize:12,marginBottom:6}}>🛡 Training Matrix PDF</button>
      <button onClick={()=>exportSchedulePDF(workers,activeDays,weekLabel,allSites)} style={{width:"100%",padding:"7px 10px",background:"#1e2535",border:"1px solid #ef4444",borderRadius:7,color:"#f87171",cursor:"pointer",fontSize:12,marginBottom:6}}>📄 Schedule PDF</button>
      <button onClick={()=>doExcel(workers,weekLabel,activeDays,siteHours,clients,allSites)} style={{width:"100%",padding:"7px 10px",background:"linear-gradient(135deg,#059669,#10b981)",border:"none",borderRadius:7,color:"#fff",cursor:"pointer",fontSize:12,fontWeight:700}}>⬇ Export Excel</button>
    </div>
  </div>;
}


// ── AI Business Insights Panel ────────────────────────────────────────────────
function AIInsightsPanel({workers,allSites,clients,invoices,activeDays,siteHours}){
  const [loading,setLoading]=useState(false);
  const [data,setData]=useState(null);
  const [hidden,setHidden]=useState(false);

  const buildCtx=()=>{
    const gross=workers.reduce((a,w)=>{const rate=w.agreedRate||0,otM=w.customOTRate||(w.overtimeMultiplier||1.5);let g=0;activeDays.forEach(d=>{const site=w.days?.[d];if(!site||isOff(site))return;const hrs=siteHours?.[site.trim()]?.hours||w.hoursPerDay?.[d]||9,ot=w.overtimeHours?.[d]||0;g+=(hrs*rate)+(ot*rate*otM);});return a+g;},0);
    const invTotal=invoices.reduce((a,inv)=>{const s=(inv.lines||[]).reduce((x,l)=>x+(l.qty||0)*(l.rate||0),0);return a+s;},0);
    const invPaid=invoices.filter(i=>i.status==="paid").reduce((a,inv)=>{const s=(inv.lines||[]).reduce((x,l)=>x+(l.qty||0)*(l.rate||0),0);return a+s;},0);
    const expiredCerts=workers.reduce((n,w)=>n+CERTS.filter(c=>cSt(c,w)==="expired").length,0);
    const expiringCerts=workers.reduce((n,w)=>n+CERTS.filter(c=>cSt(c,w)==="expiring").length,0);
    const activeSites=new Set(workers.flatMap(w=>activeDays.map(d=>(w.days?.[d]||"").trim()).filter(s=>s&&!isOff(s)))).size;
    const noRate=workers.filter(w=>!w.agreedRate&&activeDays.some(d=>!isOff(w.days?.[d]))).length;
    const avgRate=workers.filter(w=>w.agreedRate).reduce((a,w,_,arr)=>a+(w.agreedRate/arr.length),0);
    return {gross,invTotal,invPaid,expiredCerts,expiringCerts,activeSites,noRate,avgRate,workers:workers.length,clients:clients.length,invoices:invoices.length};
  };

  const analyse=async()=>{
    setLoading(true);setData(null);setHidden(false);
    const ctx=buildCtx();
    const prompt=`You are a business analyst for Bright Metalwork Ltd, a London metalwork subcontractor (steel balustrades, curtain walling, cladding).

Current week data: ${workers.length} operatives, ${ctx.activeSites} active sites, ${ctx.clients} clients.
Weekly gross labour: £${Math.round(ctx.gross)}. 
Invoiced total: £${Math.round(ctx.invTotal)}, Collected: £${Math.round(ctx.invPaid)}, Outstanding: £${Math.round(ctx.invTotal-ctx.invPaid)}.
Workers without pay rate: ${ctx.noRate}. Average rate: £${ctx.avgRate.toFixed(2)}/hr.
Cert alerts: ${ctx.expiredCerts} expired, ${ctx.expiringCerts} expiring.

Respond ONLY with valid JSON (no markdown, no backticks):
{"insights":[{"type":"warning","icon":"⚠️","title":"Title max 4 words","body":"2 sentence actionable insight"},{"type":"opportunity","icon":"💡","title":"Title max 4 words","body":"2 sentence actionable insight"},{"type":"risk","icon":"🔴","title":"Title max 4 words","body":"2 sentence insight"},{"type":"positive","icon":"✅","title":"Title max 4 words","body":"2 sentence insight"}]}`;

    try{
      const res=await fetch("https://api.anthropic.com/v1/messages",{method:"POST",headers:{"Content-Type":"application/json"},
        body:JSON.stringify({model:"claude-sonnet-4-6",max_tokens:700,messages:[{role:"user",content:prompt}]})});
      const d=await res.json();
      const text=(d.content||[]).map(b=>b.text||"").join("").trim().replace(/```json|```/g,"").trim();
      setData(JSON.parse(text));
    }catch(e){
      setData({insights:[{type:"warning",icon:"⚠️",title:"Connection Error",body:"Could not reach AI. Check your internet connection and try again."}]});
    }
    setLoading(false);
  };

  const TYPE_STYLE={
    warning:{bg:"#1a1200",border:"#92400e",color:"#fbbf24",leftBorder:"#f59e0b"},
    opportunity:{bg:"#0c1a2e",border:"#1e3a5f",color:"#60a5fa",leftBorder:"#3b82f6"},
    risk:{bg:"#1c0808",border:"#7f1d1d",color:"#f87171",leftBorder:"#ef4444"},
    positive:{bg:"#0a1c12",border:"#065f46",color:"#34d399",leftBorder:"#10b981"},
  };

  return(
    <div style={{marginTop:20,background:"linear-gradient(135deg,#0d1117,#111827)",border:"1px solid #1e2535",borderRadius:12,padding:"18px 20px"}}>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:data&&!hidden?16:0}}>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <div style={{width:36,height:36,borderRadius:10,background:"linear-gradient(135deg,#6366f1,#8b5cf6)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:17,flexShrink:0}}>🤖</div>
          <div>
            <div style={{fontSize:14,fontWeight:700,color:"#f1f5f9"}}>AI Business Insights</div>
            <div style={{fontSize:10,color:"#64748b"}}>{"Week analysis · "+new Date().toLocaleDateString("en-GB")}</div>
          </div>
        </div>
        <div style={{display:"flex",gap:7,alignItems:"center"}}>
          {data&&<button onClick={()=>setHidden(h=>!h)} style={{padding:"5px 11px",background:"#1e2535",border:"1px solid #2d3555",borderRadius:7,color:"#64748b",cursor:"pointer",fontSize:11}}>{hidden?"Show ▼":"Hide ▲"}</button>}
          <button onClick={analyse} disabled={loading}
            style={{padding:"7px 16px",background:loading?"#1e2535":"linear-gradient(135deg,#6366f1,#8b5cf6)",border:"none",borderRadius:8,color:loading?"#64748b":"#fff",cursor:loading?"default":"pointer",fontSize:12,fontWeight:700,display:"flex",alignItems:"center",gap:6,opacity:loading?0.7:1}}>
            {loading?<><span style={{width:11,height:11,border:"2px solid #64748b",borderTopColor:"#94a3b8",borderRadius:"50%",display:"inline-block",animation:"spin 0.8s linear infinite"}}/> Analysing…</>:"🔍 Analyse My Business"}
          </button>
        </div>
      </div>
      <style>{".ai-spin{animation:spin 0.8s linear infinite}@keyframes spin{to{transform:rotate(360deg)}}"}</style>
      {loading&&<div style={{marginTop:14,display:"flex",flexDirection:"column",gap:8}}>
        {[100,80,60].map((w,i)=><div key={i} style={{height:14,background:"#1e2535",borderRadius:5,width:w+"%",animation:"pulse 1.4s ease-in-out "+i*0.15+"s infinite"}}/>)}
        <style>{"@keyframes pulse{0%,100%{opacity:0.3}50%{opacity:0.8}}"}</style>
      </div>}
      {data&&!hidden&&!loading&&<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
        {(data.insights||[]).map((ins,i)=>{
          const st=TYPE_STYLE[ins.type]||TYPE_STYLE.warning;
          return <div key={i} style={{background:st.bg,border:"1px solid "+st.border,borderRadius:10,padding:"12px 14px",borderLeft:"3px solid "+st.leftBorder}}>
            <div style={{display:"flex",alignItems:"center",gap:7,marginBottom:6}}>
              <span style={{fontSize:14}}>{ins.icon}</span>
              <span style={{fontSize:11,fontWeight:700,color:st.color,textTransform:"uppercase",letterSpacing:"0.04em"}}>{ins.title}</span>
            </div>
            <div style={{fontSize:12,color:"#94a3b8",lineHeight:1.6}}>{ins.body}</div>
          </div>;
        })}
      </div>}
    </div>
  );
}

// ── Dashboard Home ────────────────────────────────────────────────────────────
function DHome({workers,allSites,clients,invoices,scopeData,activeDays,siteHours,weeklyRecords,setPage}){
  const totalLabour=useMemo(()=>workers.reduce((a,w)=>{const{gross}=calcPay(w,activeDays,siteHours);return a+gross;},0),[workers,activeDays,siteHours]);
  const totalInvoiced=invoices.reduce((a,i)=>a+i.amount,0);
  const expiring=workers.flatMap(w=>Object.values(w.certs||{}).filter(c=>{if(!c.held||!c.expiry)return false;const d=(new Date(c.expiry)-new Date())/86400000;return d>=0&&d<30;})).length;
  const expired=workers.flatMap(w=>Object.values(w.certs||{}).filter(c=>{if(!c.held||!c.expiry)return false;return new Date(c.expiry)<new Date();})).length;

  const objects=[
    {id:"workers",icon:"👷",label:"Workers",count:workers.length,color:"#3b82f6",sub:`${workers.filter(w=>Object.values(w.days||{}).some(d=>d&&!isOff(d))).length} active this week`},
    {id:"sites",icon:"🏗",label:"Sites",count:allSites.filter(s=>!isOff(s.name)).length,color:"#f59e0b",sub:`${allSites.filter(s=>s.scopes&&s.scopes.length>0).length} with scope`},
    {id:"clients",icon:"👔",label:"Clients",count:clients.length,color:"#8b5cf6",sub:`${clients.filter(c=>(c.rates||[]).length>0).length} with day rates`},
    {id:"schedule",icon:"📋",label:"Labour Schedule",count:"WC",color:"#06b6d4",sub:"Weekly worker allocation"},
    {id:"timesheets",icon:"⏱",label:"Timesheets",count:0,color:"#10b981",sub:"Coming soon"},
    {id:"payslips",icon:"💷",label:"Payslips",count:"£"+totalLabour.toFixed(0),color:"#34d399",sub:"Weekly labour gross"},
    {id:"invoices",icon:"🧾",label:"Invoices",count:invoices.length,color:"#fbbf24",sub:`£${totalInvoiced.toLocaleString()} total · ${invoices.filter(i=>i.status==="pending").length} pending`},
    {id:"certs",icon:"🛡",label:"Certificates",count:workers.reduce((a,w)=>a+Object.values(w.certs||{}).filter(c=>c.held).length,0),color:expiring+expired>0?"#fbbf24":"#34d399",sub:`${expiring} expiring · ${expired} expired`},
    {id:"payapps",icon:"📐",label:"Payment Apps",count:0,color:"#a78bfa",sub:"Valuation applications"},
  ];

  return <div>
    <DPageHdr title="🏗 Bright Metalwork" sub="Project Management Overview"/>
    <div style={DS.body}>
      {/* ── Financial Overview Panel (matches screenshot) ── */}
      <FinancialOverviewPanel invoices={invoices} clients={clients} allSites={allSites} workers={workers} activeDays={activeDays} siteHours={siteHours}/>
      <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:10,marginBottom:24}}>
        <DStat label="Operatives" value={workers.length} color="#60a5fa"/>
        <DStat label="Active Sites" value={allSites.filter(s=>!isOff(s.name)).length} color="#f59e0b"/>
        <DStat label="Weekly Labour" value={"£"+totalLabour.toFixed(0)} color="#34d399"/>
        <DStat label="Total Invoiced" value={"£"+totalInvoiced.toLocaleString()} color="#a78bfa"/>
        <DStat label="Cert Alerts" value={expiring+expired} color={expiring+expired>0?"#fbbf24":"#34d399"} sub={`${expiring} expiring · ${expired} expired`}/>
      </div>
      <div style={{fontSize:11,color:"#64748b",fontWeight:700,textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:14}}>All Objects — click to open</div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:12}}>
        {objects.map(obj=>(
          <div key={obj.id} onClick={()=>setPage(obj.id)}
            style={{...DS.card(obj.color),borderColor:`${obj.color}33`}}
            onMouseEnter={e=>{e.currentTarget.style.borderColor=obj.color;e.currentTarget.style.transform="translateY(-2px)";e.currentTarget.style.boxShadow=`0 8px 30px ${obj.color}22`;}}
            onMouseLeave={e=>{e.currentTarget.style.borderColor=`${obj.color}33`;e.currentTarget.style.transform="";e.currentTarget.style.boxShadow="";}}>
            <div style={{position:"absolute",top:0,left:0,right:0,height:2,background:`linear-gradient(90deg,${obj.color},${obj.color}44)`}}/>
            <div style={{display:"flex",alignItems:"flex-start",gap:12}}>
              <div style={{width:42,height:42,borderRadius:10,background:`${obj.color}18`,border:`1px solid ${obj.color}33`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:19,flexShrink:0}}>{obj.icon}</div>
              <div style={{flex:1,minWidth:0}}>
                <div style={{fontSize:14,fontWeight:700,color:"#f1f5f9",marginBottom:2}}>{obj.label}</div>
                <div style={{fontSize:11,color:"#64748b",marginBottom:8}}>{obj.sub}</div>
                <div style={{fontSize:24,fontWeight:900,color:obj.color,lineHeight:1}}>{obj.count}</div>
              </div>
              <span style={{color:`${obj.color}66`,fontSize:16}}>→</span>
            </div>
          </div>
        ))}
      </div>
      <AIInsightsPanel workers={workers} allSites={allSites} clients={clients} invoices={invoices} activeDays={activeDays} siteHours={siteHours}/>
    </div>
  </div>;
}

// ── Dashboard Workers Page ────────────────────────────────────────────────────
function DWorkers({workers,allSites,clients,activeDays,siteHours,setPage,setDetailId,setModal}){
  const[search,setSearch]=useState("");
  const shown=workers.filter(w=>!search||w.name.toLowerCase().includes(search.toLowerCase())||w.position.toLowerCase().includes(search.toLowerCase()));
  const {gross}=useMemo(()=>workers.reduce((a,w)=>{const r=calcPay(w,activeDays,siteHours);return{gross:a.gross+r.gross};},{gross:0}),[workers,activeDays,siteHours]);

  return <div>
    <DPageHdr title="👷 Workers" sub={`${workers.length} operatives · £${gross.toFixed(0)} gross this week`}
      actions={<button onClick={()=>setModal({type:"worker",worker:mkW()})} style={{padding:"7px 14px",background:"linear-gradient(135deg,#3b82f6,#6366f1)",border:"none",borderRadius:7,color:"#fff",cursor:"pointer",fontSize:12,fontWeight:700}}>+ Add Worker</button>}/>
    <div style={DS.body}>
      <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 Search by name or position…"
        style={{...{width:"100%",background:"#0f1421",border:"1px solid #2d3555",borderRadius:6,padding:"7px 10px",color:"#e2e8f0",fontSize:13,outline:"none",boxSizing:"border-box"},maxWidth:320,marginBottom:16}}/>
      <DTable cols={[
        {key:"name",label:"Name",w:200,r:(v,r)=><div style={{display:"flex",alignItems:"center",gap:9}}>
          <div style={{width:30,height:30,borderRadius:7,background:r.color+"22",border:`1px solid ${r.color}44`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:12,fontWeight:800,color:r.color,flexShrink:0}}>{(v||"?")[0]}</div>
          <div><div style={{fontWeight:600,color:"#f1f5f9"}}>{v||"Unnamed"}</div><div style={{fontSize:10,color:"#64748b"}}>{r.company}</div></div>
        </div>},
        {key:"position",label:"Position",r:v=><span style={DS.badge("#60a5fa","#0d1a2e")}>{v||"—"}</span>},
        {key:"agreedRate",label:"Rate",r:v=>v?<span style={{color:"#34d399",fontWeight:600}}>£{v}/hr</span>:<span style={{color:"#374151"}}>—</span>},
        {key:"taxRate",label:"Tax",r:v=><span style={{color:v===0.30?"#f87171":v===0.20?"#fbbf24":"#34d399",fontWeight:600}}>{Math.round((v||0)*100)}%</span>},
        {key:"certs",label:"Certs",r:(v,r)=>{
          const held=Object.values(r.certs||{}).filter(c=>c.held);
          const exp=held.filter(c=>{if(!c.expiry)return false;return new Date(c.expiry)<new Date();});
          const warn=held.filter(c=>{if(!c.expiry)return false;const d=(new Date(c.expiry)-new Date())/86400000;return d>=0&&d<30;});
          return <div style={{display:"flex",gap:5,alignItems:"center"}}>
            <span style={{color:"#34d399",fontSize:11,fontWeight:700}}>✓{held.length}</span>
            {warn.length>0&&<span style={{color:"#fbbf24",fontSize:11,fontWeight:700}}>⚠{warn.length}</span>}
            {exp.length>0&&<span style={{color:"#f87171",fontSize:11,fontWeight:700}}>✗{exp.length}</span>}
          </div>;
        }},
        {key:"days",label:"This Week",r:(v,r)=>{
          const site=Object.values(v||{}).filter(d=>d&&!isOff(d));
          const primary=[...new Set(site)][0];
          const s=primary&&allSites.find(x=>primary.includes(x.name));
          return s?<span style={DS.pill(s.color)}>{s.name}</span>:<span style={{color:"#374151",fontSize:11}}>—</span>;
        }},
              {key:"id",label:"",w:80,r:(_,r)=><div style={{display:"flex",gap:4}}>
          <button onClick={e=>{e.stopPropagation();openWorkerWindow(r,allSites,formatWeekLabel(new Date()),activeDays,siteHours);}} title="Open in new window" style={{padding:"3px 7px",background:"#1a1f2e",border:"1px solid #60a5fa",borderRadius:5,color:"#60a5fa",cursor:"pointer",fontSize:10}}>🔗</button>
        </div>},
      ]} rows={shown} onRow={r=>{setDetailId(r.id);setPage("worker_detail");}}/>
    </div>
  </div>;
}

// ── Dashboard Worker Detail ───────────────────────────────────────────────────
function DWorkerDetail({workers,allSites,clients,activeDays,siteHours,workerId,setPage,setModal}){
  const w=workers.find(x=>x.id===workerId);
  if(!w) return <div style={DS.body}><div style={{color:"#374151"}}>Worker not found.</div></div>;
  const {gross,net,stdH,otH}=calcPay(w,activeDays,siteHours);
  const heldCerts=Object.entries(w.certs||{}).filter(([,v])=>v.held).map(([k,v])=>({...v,key:k,label:CERTS.find(c=>c.key===k)?.label||k}));

  return <div>
    <DPageHdr title={w.name} sub={`${w.position} · ${w.company}`} back="Workers" onBack={()=>setPage("workers")}
      actions={<>
        <button onClick={()=>setModal({type:"worker",worker:w})} style={{padding:"6px 12px",background:"#1e3a5f",border:"1px solid #3b82f6",borderRadius:6,color:"#60a5fa",cursor:"pointer",fontSize:12,fontWeight:600}}>✏️ Edit</button>
        <button onClick={()=>exportWorkerProfile(w,allSites,formatWeekLabel(new Date()))} style={{padding:"6px 12px",background:"#1e2535",border:"1px solid #2d3555",borderRadius:6,color:"#94a3b8",cursor:"pointer",fontSize:12}}>📋 Profile PDF</button>
        <button onClick={()=>exportPayslip(w,activeDays,formatWeekLabel(new Date()),siteHours)} style={{padding:"6px 12px",background:"#0d2218",border:"1px solid #10b981",borderRadius:6,color:"#34d399",cursor:"pointer",fontSize:12,fontWeight:600}}>💷 Payslip</button>
      </>}/>
    <div style={DS.body}>
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10,marginBottom:20}}>
        <DStat label="Hourly Rate" value={w.agreedRate?`£${w.agreedRate}/hr`:"Not set"} color="#34d399"/>
        <DStat label="Tax Rate" value={Math.round((w.taxRate||0)*100)+"%"} color={w.taxRate===0.30?"#f87171":w.taxRate===0.20?"#fbbf24":"#34d399"}/>
        <DStat label="This Week Gross" value={gross>0?`£${gross.toFixed(0)}`:"£0"} color="#60a5fa"/>
        <DStat label="This Week Net" value={net>0?`£${net.toFixed(0)}`:"£0"} color="#a78bfa"/>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginBottom:16}}>
        <div style={{background:"#111827",border:"1px solid #1e2535",borderRadius:10,padding:16}}>
          <div style={{fontSize:11,color:"#64748b",fontWeight:700,textTransform:"uppercase",marginBottom:12}}>Contact Details</div>
          {[["Name",w.name],["Position",w.position],["Company",w.company],["Phone",w.contact||"—"],["Email",w.email||"—"]].map(([l,v])=>(
            <div key={l} style={{display:"flex",gap:8,marginBottom:8}}>
              <span style={{fontSize:10,color:"#64748b",fontWeight:600,minWidth:75,textTransform:"uppercase",flexShrink:0}}>{l}</span>
              <span style={{fontSize:12,color:"#e2e8f0"}}>{v}</span>
            </div>
          ))}
        </div>
        <div style={{background:"#111827",border:"1px solid #1e2535",borderRadius:10,padding:16}}>
          <div style={{fontSize:11,color:"#64748b",fontWeight:700,textTransform:"uppercase",marginBottom:12}}>This Week Allocation</div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:4,marginBottom:10}}>
            {(w.days?BASE_DAYS:[]).map(d=>{const site=w.days[d];const s=site&&allSites.find(x=>site.includes(x.name));const c=s?.color||"#374151";return <div key={d} style={{textAlign:"center"}}>
              <div style={{fontSize:9,color:"#64748b",fontWeight:700,textTransform:"uppercase",marginBottom:3}}>{d}</div>
              <div style={{height:3,borderRadius:2,background:c,marginBottom:3}}/>
              <div style={{fontSize:8,color:c,fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{site?site.split("-")[0].trim():"—"}</div>
            </div>;})}
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:6,marginTop:10}}>
            {[["Std Hrs",stdH+"h","#60a5fa"],["OT Hrs",otH>0?otH+"h":"—","#fbbf24"],["Gross",`£${gross.toFixed(0)}`,"#34d399"]].map(([l,v,c])=>(
              <div key={l} style={{background:"#0f1421",borderRadius:7,padding:8,textAlign:"center"}}>
                <div style={{fontSize:9,color:"#64748b",textTransform:"uppercase"}}>{l}</div>
                <div style={{fontSize:14,fontWeight:800,color:c,marginTop:2}}>{v}</div>
              </div>
            ))}
          </div>
        </div>
      </div>
      <div style={{fontSize:11,color:"#64748b",fontWeight:700,textTransform:"uppercase",marginBottom:10}}>Certificates ({heldCerts.length} held)</div>
      {heldCerts.length===0&&<div style={{color:"#374151",fontSize:12,textAlign:"center",padding:"20px 0",border:"1px dashed #1e2535",borderRadius:8}}>No certificates on file. Edit worker to add certificates.</div>}
      <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:8}}>
        {heldCerts.map(c=>{
          const exp=c.expiry?new Date(c.expiry):null;const now=new Date();
          const days=exp?(exp-now)/86400000:null;
          const st=!exp?"valid":days<0?"expired":days<30?"expiring":"valid";
          const stColor={valid:"#34d399",expiring:"#fbbf24",expired:"#f87171"}[st];
          return <div key={c.key} style={{background:"#0f1421",borderRadius:8,padding:"10px 12px",border:`1px solid ${stColor}44`}}>
            <div style={{fontWeight:600,color:"#e2e8f0",fontSize:12,marginBottom:4}}>{c.label}</div>
            {c.regNo&&<div style={{fontSize:11,color:"#60a5fa",marginBottom:2}}>Reg: {c.regNo}</div>}
            {c.expiry&&<div style={{fontSize:10,color:"#64748b",marginBottom:3}}>Exp: {c.expiry} {days!==null&&days<30&&<span style={{color:stColor}}>({days<0?"EXPIRED":`${Math.ceil(days)}d`})</span>}</div>}
            {c.fileUrl&&<a href={c.fileUrl} target="_blank" rel="noreferrer" style={{fontSize:10,color:"#60a5fa"}}>📎 View</a>}
            <div style={{fontSize:10,color:stColor,fontWeight:700,textTransform:"uppercase",marginTop:4}}>{st}</div>
          </div>;
        })}
      </div>
    </div>
  </div>;
}

// ── Dashboard Sites Page ──────────────────────────────────────────────────────
function DSites({allSites,clients,workers,activeDays,siteHours,setPage,setDetailId,setModal}){
  const activeSites=allSites.filter(s=>!isOff(s.name));
  return <div>
    <DPageHdr title="🏗 Sites" sub={`${activeSites.length} sites`}
      actions={<button onClick={()=>setModal({type:"sites"})} style={{padding:"7px 14px",background:"#1e2535",border:"1px solid #f59e0b",borderRadius:7,color:"#fbbf24",cursor:"pointer",fontSize:12,fontWeight:700}}>🏗 Manage Sites</button>}/>
    <div style={DS.body}>
      <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:12}}>
        {activeSites.map(site=>{
          const wk=workers.filter(w=>BASE_DAYS.some(d=>(w.days[d]||"").includes(site.name)));
          const client=clients.find(c=>c.id===site.clientId);
          const sc=site.scopes||[],vr=site.variations||[];
          const scopeT=sc.reduce((a,s)=>a+(s.qty*s.rate),0);
          const varT=vr.reduce((a,v)=>a+(v.type==="addition"?v.value:-v.value),0);
          return <div key={site.id} onClick={()=>{setDetailId(site.id);setPage("site_detail");}}
            style={{...DS.card(site.color),borderColor:`${site.color}33`}}
            onMouseEnter={e=>{e.currentTarget.style.borderColor=site.color;e.currentTarget.style.transform="translateY(-2px)";}}
            onMouseLeave={e=>{e.currentTarget.style.borderColor=`${site.color}33`;e.currentTarget.style.transform="";}}>
            <div style={{position:"absolute",top:0,left:0,right:0,height:2,background:site.color}}/>
            <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:12}}>
              <span style={{width:11,height:11,borderRadius:"50%",background:site.color,flexShrink:0}}/>
              <span style={{fontSize:15,fontWeight:800,color:"#f1f5f9",flex:1}}>{site.name}</span>
              {client&&<span style={DS.pill(client.color)}>{client.name}</span>}
            </div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:6}}>
              {[["Contract","£"+(scopeT+varT).toLocaleString(),"#60a5fa"],["Scope","£"+scopeT.toLocaleString(),"#34d399"],["Workers",wk.length,"#a78bfa"],["Variations",vr.length,"#fbbf24"]].map(([l,v,c])=>(
                <div key={l} style={{background:"#0a0e17",borderRadius:6,padding:"6px 8px"}}>
                  <div style={{fontSize:9,color:"#64748b",textTransform:"uppercase"}}>{l}</div>
                  <div style={{fontSize:13,fontWeight:800,color:c,marginTop:2}}>{v}</div>
                </div>
              ))}
            </div>
          </div>;
        })}
      </div>
    </div>
  </div>;
}

// ── Dashboard Site Detail ─────────────────────────────────────────────────────
function DSiteDetail({allSites,clients,workers,activeDays,siteHours,siteId,setPage,setDetailId,setModal,saveSiteDetail}){
  const site=allSites.find(s=>s.id===siteId);
  if(!site) return <div style={DS.body}><div style={{color:"#374151"}}>Site not found.</div></div>;
  const client=clients.find(c=>c.id===site.clientId);
  const siteWorkers=workers.filter(w=>BASE_DAYS.some(d=>(w.days[d]||"").includes(site.name)));
  const sc=site.scopes||[],vr=site.variations||[];
  const scopeT=sc.reduce((a,s)=>a+(s.qty*s.rate),0);
  const varT=vr.reduce((a,v)=>a+(v.type==="addition"?v.value:-v.value),0);
  const contract=scopeT+varT;
  const labourCost=useMemo(()=>{let t=0;workers.forEach(w=>{const{bd}=calcPay(w,activeDays,siteHours);Object.values(bd).forEach(b=>{if(b.site===site.name||b.site.includes(site.name))t+=b.gross;});});return t;},[workers,activeDays,siteHours,site.name]);
  const profit=contract-labourCost;
  const [tab,setTab]=useState("scopes");

  return <div>
    <DPageHdr title={<span style={{display:"flex",alignItems:"center",gap:9}}><span style={{width:12,height:12,borderRadius:"50%",background:site.color}}/>{site.name}</span>}
      sub={client?`Client: ${client.name}`:"No client assigned"} back="Sites" onBack={()=>setPage("sites")}
      actions={<button onClick={()=>setModal({type:"siteDetail",site})} style={{padding:"6px 12px",background:"#1e3a5f",border:"1px solid #3b82f6",borderRadius:6,color:"#60a5fa",cursor:"pointer",fontSize:12,fontWeight:600}}>✏️ Edit Site</button>}/>
    <div style={DS.body}>
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10,marginBottom:18}}>
        <DStat label="Contract" value={"£"+contract.toLocaleString()} color="#60a5fa"/>
        <DStat label="Labour Cost" value={"£"+labourCost.toFixed(0)} color="#f87171"/>
        <DStat label="Profit / Loss" value={"£"+Math.abs(profit).toFixed(0)} color={profit>=0?"#34d399":"#f87171"} sub={profit>=0?"Profit":"Loss"}/>
        <DStat label="Workers This Week" value={siteWorkers.length} color="#a78bfa"/>
      </div>
      <div style={{display:"flex",gap:3,background:"#0d1117",borderRadius:8,padding:3,marginBottom:16,width:"fit-content"}}>
        {[["scopes","📋 Scopes"],["variations","⚡ Variations"],["workers","👷 Workers"],["costs","💷 Costs"]].map(([v,l])=>(
          <button key={v} onClick={()=>setTab(v)} style={{padding:"5px 12px",background:tab===v?"#1e3a5f":"transparent",border:tab===v?"1px solid #3b82f6":"1px solid transparent",borderRadius:6,color:tab===v?"#60a5fa":"#64748b",cursor:"pointer",fontSize:12,fontWeight:tab===v?700:400}}>{l}</button>
        ))}
      </div>
      {tab==="scopes"&&<DTable cols={[
        {key:"description",label:"Description",w:280},{key:"unit",label:"Unit",r:v=><span style={DS.badge("#94a3b8","#1e2535")}>{v}</span>},
        {key:"qty",label:"Qty",r:v=><span style={{color:"#60a5fa",fontWeight:600}}>{v}</span>},
        {key:"rate",label:"Rate",r:v=>`£${v.toLocaleString()}`},
        {key:"total",label:"Total",r:(_,r)=><span style={{color:"#34d399",fontWeight:700}}>£{(r.qty*r.rate).toLocaleString()}</span>},
      ]} rows={sc}/>}
      {tab==="variations"&&<DTable cols={[
        {key:"description",label:"Description",w:280},
        {key:"type",label:"Type",r:v=><DStatusBadge status={v}/>},
        {key:"value",label:"Value",r:(v,r)=><span style={{color:r.type==="addition"?"#34d399":"#f87171",fontWeight:700}}>{r.type==="addition"?"+":"-"}£{v.toLocaleString()}</span>},
        {key:"approved",label:"Status",r:v=><DStatusBadge status={v?"approved":"pending"}/>},
      ]} rows={vr}/>}
      {tab==="workers"&&<DTable cols={[
        {key:"name",label:"Worker",r:(v,r)=><div style={{display:"flex",alignItems:"center",gap:8}}><div style={{width:26,height:26,borderRadius:6,background:r.color+"22",display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,fontWeight:700,color:r.color}}>{v[0]}</div><div><div style={{fontWeight:600,color:"#f1f5f9"}}>{v}</div><div style={{fontSize:10,color:"#64748b"}}>{r.position}</div></div></div>},
        {key:"agreedRate",label:"Rate",r:v=>v?`£${v}/hr`:"—"},
        {key:"certs",label:"Certs",r:(v,r)=><span style={{color:"#a78bfa",fontWeight:600}}>{Object.values(r.certs||{}).filter(c=>c.held).length} held</span>},
      ]} rows={siteWorkers} onRow={r=>{setDetailId(r.id);setPage("worker_detail");}}/>}
      {tab==="costs"&&<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
        <div style={{background:"#0f1421",borderRadius:10,padding:16,border:"1px solid #1e2535"}}>
          <div style={{fontSize:11,color:"#64748b",fontWeight:700,textTransform:"uppercase",marginBottom:12}}>Income</div>
          {[["Agreed Scope","£"+scopeT.toLocaleString(),"#60a5fa"],["Variations",(varT>=0?"+":"-")+"£"+Math.abs(varT).toLocaleString(),"#fbbf24"],["Total Contract","£"+contract.toLocaleString(),"#34d399"]].map(([l,v,c])=>(
            <div key={l} style={{display:"flex",justifyContent:"space-between",padding:"8px 0",borderBottom:"1px solid #1e2535"}}><span style={{color:"#94a3b8"}}>{l}</span><span style={{fontWeight:700,color:c}}>{v}</span></div>
          ))}
        </div>
        <div style={{background:"#0f1421",borderRadius:10,padding:16,border:"1px solid #1e2535"}}>
          <div style={{fontSize:11,color:"#64748b",fontWeight:700,textTransform:"uppercase",marginBottom:12}}>Costs & Margin</div>
          {[["Labour (this week)","£"+labourCost.toFixed(0),"#f87171"],[profit>=0?"Profit":"Loss","£"+Math.abs(profit).toFixed(0),profit>=0?"#34d399":"#f87171"]].map(([l,v,c])=>(
            <div key={l} style={{display:"flex",justifyContent:"space-between",padding:"8px 0",borderBottom:"1px solid #1e2535"}}><span style={{color:"#94a3b8"}}>{l}</span><span style={{fontWeight:700,color:c}}>{v}</span></div>
          ))}
        </div>
      </div>}
    </div>
  </div>;
}

// ── Dashboard Clients Page ────────────────────────────────────────────────────
function DClients({clients,allSites,invoices,workers,activeDays,siteHours,setPage,setDetailId,setModal}){
  return <div>
    <DPageHdr title="👔 Clients" sub={`${clients.length} accounts`}
      actions={<button onClick={()=>setModal({type:"clients"})} style={{padding:"7px 14px",background:"#1e2535",border:"1px solid #8b5cf6",borderRadius:7,color:"#a78bfa",cursor:"pointer",fontSize:12,fontWeight:700}}>👔 Manage Clients</button>}/>
    <div style={DS.body}>
      {clients.map(c=>{
        const sites=allSites.filter(s=>s.clientId===c.id);
        const invs=invoices.filter(i=>i.siteId&&sites.find(s=>s.id===i.siteId));
        const totalInv=invs.reduce((a,i)=>a+i.amount,0);
        return <div key={c.id} onClick={()=>{setDetailId(c.id);setPage("client_detail");}}
          style={{...DS.card(c.color),borderColor:`${c.color}33`,marginBottom:12}}
          onMouseEnter={e=>{e.currentTarget.style.borderColor=c.color;e.currentTarget.style.transform="translateY(-2px)";}}
          onMouseLeave={e=>{e.currentTarget.style.borderColor=`${c.color}33`;e.currentTarget.style.transform="";}}>
          <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:12}}>
            <div style={{width:40,height:40,borderRadius:10,background:`${c.color}18`,border:`1px solid ${c.color}33`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:17,fontWeight:800,color:c.color}}>{c.name[0]}</div>
            <div style={{flex:1}}><div style={{fontSize:15,fontWeight:800,color:"#f1f5f9"}}>{c.name}</div><div style={{fontSize:11,color:"#64748b"}}>{c.contact}</div></div>
            <div style={{display:"flex",gap:4,alignItems:"center"}}>
              <button onClick={e=>{e.stopPropagation();openClientWindow(c,allSites,invoices,workers,activeDays,siteHours);}} title="Open client in new window" style={{padding:"3px 8px",background:"#1a1f2e",border:"1px solid #60a5fa",borderRadius:5,color:"#60a5fa",cursor:"pointer",fontSize:10}}>🔗</button>
              <span style={{color:`${c.color}66`,fontSize:16}}>→</span>
            </div>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:10}}>
            {[["Sites",sites.length,"#60a5fa"],["Invoiced","£"+totalInv.toLocaleString(),"#34d399"],["Rates",(c.rates||[]).length,"#a78bfa"]].map(([l,v,col])=>(
              <div key={l} style={{background:"#0a0e17",borderRadius:7,padding:"7px 9px"}}><div style={{fontSize:9,color:"#64748b",textTransform:"uppercase"}}>{l}</div><div style={{fontSize:14,fontWeight:800,color:col,marginTop:2}}>{v}</div></div>
            ))}
          </div>
          <div style={{display:"flex",flexWrap:"wrap",gap:5}}>
            {(c.rates||[]).map(r=><span key={r.id} style={DS.pill(c.color)}>{TEAM_TYPES.find(t=>t.key===r.teamType)?.label.split("(")[0].trim()||r.teamType} · £{r.dayRate}/day</span>)}
            {(c.rates||[]).length===0&&<span style={{color:"#374151",fontSize:11}}>No day rates configured</span>}
          </div>
        </div>;
      })}
    </div>
  </div>;
}

// ── Dashboard Certs Page ──────────────────────────────────────────────────────
function DCerts({workers,setPage,setDetailId}){
  const allCerts=useMemo(()=>workers.flatMap(w=>
    Object.entries(w.certs||{}).filter(([,v])=>v.held).map(([k,v])=>({
      ...v,key:k,workerId:w.id,workerName:w.name,workerColor:w.color,
      label:CERTS.find(c=>c.key===k)?.label||k,
    }))
  ),[workers]);
  const expired=allCerts.filter(c=>{if(!c.expiry)return false;return new Date(c.expiry)<new Date();});
  const expiring=allCerts.filter(c=>{if(!c.expiry)return false;const d=(new Date(c.expiry)-new Date())/86400000;return d>=0&&d<30;});
  const[filter,setFilter]=useState("all");
  const shown=filter==="all"?allCerts:filter==="expiring"?expiring:expired;

  return <div>
    <DPageHdr title="🛡 Certificates" sub={`${allCerts.length} held · ${expiring.length} expiring · ${expired.length} expired`}/>
    <div style={DS.body}>
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10,marginBottom:18}}>
        <DStat label="Total" value={allCerts.length} color="#a78bfa"/>
        <DStat label="Valid" value={allCerts.length-expiring.length-expired.length} color="#34d399"/>
        <DStat label="Expiring Soon" value={expiring.length} color="#fbbf24" sub="within 30 days"/>
        <DStat label="Expired" value={expired.length} color="#f87171"/>
      </div>
      <div style={{display:"flex",gap:7,marginBottom:14}}>
        {[["all","All"],["expiring","Expiring Soon"],["expired","Expired"]].map(([v,l])=>(
          <button key={v} onClick={()=>setFilter(v)} style={{padding:"5px 12px",background:filter===v?"#1e3a5f":"#1a1f2e",border:`1px solid ${filter===v?"#3b82f6":"#2d3555"}`,borderRadius:7,color:filter===v?"#60a5fa":"#64748b",cursor:"pointer",fontSize:12,fontWeight:filter===v?700:400}}>{l}</button>
        ))}
      </div>
      <DTable cols={[
        {key:"label",label:"Certificate",w:200},
        {key:"workerName",label:"Worker",r:(v,r)=><div style={{display:"flex",alignItems:"center",gap:7}}><div style={{width:24,height:24,borderRadius:5,background:r.workerColor+"22",display:"flex",alignItems:"center",justifyContent:"center",fontSize:10,fontWeight:700,color:r.workerColor}}>{v[0]}</div><span style={{fontWeight:500}}>{v}</span></div>},
        {key:"regNo",label:"Reg No",r:v=>v?<span style={{color:"#60a5fa",fontFamily:"monospace",fontSize:12}}>{v}</span>:<span style={{color:"#374151"}}>—</span>},
        {key:"expiry",label:"Expiry",r:v=>{if(!v)return <span style={{color:"#374151"}}>No expiry</span>;const d=(new Date(v)-new Date())/86400000;const col=d<0?"#f87171":d<30?"#fbbf24":"#34d399";return <span style={{color:col,fontWeight:600}}>{v} {d<0?"(EXPIRED)":d<30?`(${Math.ceil(d)}d)`:"✓"}</span>;}},
        {key:"fileUrl",label:"File",r:v=>v?<a href={v} target="_blank" rel="noreferrer" style={{color:"#60a5fa",fontSize:11}}>📎 View</a>:<span style={{color:"#374151",fontSize:11}}>—</span>},
      ]} rows={shown} onRow={r=>{setDetailId(r.workerId);setPage("worker_detail");}}/>
    </div>
  </div>;
}

// ── Dashboard Invoices Page ───────────────────────────────────────────────────
function DInvoices({invoices,allSites,clients,setModal}){
  const total=invoices.reduce((a,i)=>a+i.amount,0);
  const paid=invoices.filter(i=>i.status==="paid").reduce((a,i)=>a+i.amount,0);
  return <div>
    <DPageHdr title="🧾 Invoices" sub={`${invoices.length} invoices · £${total.toLocaleString()} total`}
      actions={<button onClick={()=>setModal({type:"invoice",invoice:{id:"inv"+Date.now(),number:"INV-00"+(invoices.length+1),siteId:"",clientId:"",date:new Date().toISOString().slice(0,10),status:"draft",amount:0,items:[]}})} style={{padding:"7px 14px",background:"linear-gradient(135deg,#3b82f6,#6366f1)",border:"none",borderRadius:7,color:"#fff",cursor:"pointer",fontSize:12,fontWeight:700}}>+ New Invoice</button>}/>
    <div style={DS.body}>
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10,marginBottom:18}}>
        <DStat label="Total" value={"£"+total.toLocaleString()} color="#34d399"/>
        <DStat label="Paid" value={"£"+paid.toLocaleString()} color="#a78bfa"/>
        <DStat label="Outstanding" value={"£"+(total-paid).toLocaleString()} color="#fbbf24"/>
        <DStat label="Draft" value={invoices.filter(i=>i.status==="draft").length} color="#94a3b8"/>
      </div>
      <DTable cols={[
        {key:"number",label:"Invoice",r:v=><span style={{color:"#60a5fa",fontWeight:700}}>{v}</span>},
        {key:"date",label:"Date"},
        {key:"siteId",label:"Site",r:v=>{const s=allSites.find(x=>x.id===v);return s?<span style={DS.pill(s.color)}>{s.name}</span>:<span style={{color:"#374151"}}>—</span>;}},
        {key:"clientId",label:"Client",r:v=>{const c=clients.find(x=>x.id===v);return c?<span style={DS.pill(c.color)}>{c.name}</span>:<span style={{color:"#374151"}}>—</span>;}},
        {key:"amount",label:"Amount",r:v=><span style={{color:"#34d399",fontWeight:700}}>£{v.toLocaleString()}</span>},
        {key:"status",label:"Status",r:v=><DStatusBadge status={v}/>},
              {key:"id",label:"",w:100,r:(_,r)=><div style={{display:"flex",gap:4}}>
          <button onClick={e=>{e.stopPropagation();setModal({type:"invoice",invoice:r});}} style={{padding:"3px 7px",background:"#1e3a5f",border:"1px solid #3b82f6",borderRadius:5,color:"#60a5fa",cursor:"pointer",fontSize:10}}>✏️ Edit</button>
          <button onClick={e=>{e.stopPropagation();const c=allSites&&clients.find(x=>x.id===r.clientId);const s=allSites&&allSites.find(x=>x.id===r.siteId);openInvoiceWindow(r,c,s);}} title="Open in new window" style={{padding:"3px 7px",background:"1a1f2e",border:"1px solid #60a5fa",borderRadius:5,color:"#60a5fa",cursor:"pointer",fontSize:10}}>🔗</button>
        </div>},
      ]} rows={invoices}/>
    </div>
  </div>;
}

// ── Coming Soon placeholder ───────────────────────────────────────────────────

// ── Dashboard Payroll Page ────────────────────────────────────────────────────
function DPayroll({workers,allSites,activeDays,siteHours,weekLabel,setModal}){
  const rows=workers.map(w=>({...w,...calcPay(w,activeDays,siteHours)}));
  const tot=rows.reduce((a,r)=>({g:a.g+r.gross,t:a.t+r.tax,n:a.n+r.net}),{g:0,t:0,n:0});
  return <div>
    <DPageHdr title="💷 Payroll" sub={`WC: ${weekLabel} · ${workers.length} workers`}
      actions={<>
        <button onClick={()=>doExcel(workers,weekLabel,activeDays,siteHours,[],allSites)} style={{padding:"6px 12px",background:"linear-gradient(135deg,#059669,#10b981)",border:"none",borderRadius:6,color:"#fff",cursor:"pointer",fontSize:12,fontWeight:700}}>⬇ Excel</button>
      </>}/>
    <div style={DS.body}>
      <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10,marginBottom:18}}>
        <DStat label="Gross" value={"£"+tot.g.toFixed(0)} color="#34d399"/>
        <DStat label="Tax" value={"£"+tot.t.toFixed(0)} color="#f87171"/>
        <DStat label="Net Pay" value={"£"+tot.n.toFixed(0)} color="#a78bfa"/>
      </div>
      <DTable cols={[
        {key:"name",label:"Worker",w:180,r:(v,r)=><div style={{display:"flex",alignItems:"center",gap:8}}><div style={{width:26,height:26,borderRadius:6,background:(r.color||"#3b82f6")+"22",display:"flex",alignItems:"center",justifyContent:"center",fontSize:10,fontWeight:700,color:r.color||"#3b82f6"}}>{v[0]}</div><div><div style={{fontWeight:600,color:"#f1f5f9"}}>{v}</div><div style={{fontSize:10,color:"#64748b"}}>{r.position}</div></div></div>},
        {key:"agreedRate",label:"Rate",r:v=>v?<span style={{color:"#34d399",fontWeight:600}}>£{v}/hr</span>:<span style={{color:"#374151"}}>—</span>},
        {key:"stdH",label:"Std h",r:v=><span style={{color:"#60a5fa",fontWeight:600}}>{v}h</span>},
        {key:"otH",label:"OT h",r:v=>v>0?<span style={{color:"#fbbf24",fontWeight:600}}>{v}h</span>:<span style={{color:"#374151"}}>—</span>},
        {key:"gross",label:"Gross",r:v=><span style={{color:"#34d399",fontWeight:700}}>£{v.toFixed(2)}</span>},
        {key:"tax",label:"Tax",r:v=><span style={{color:"#f87171"}}>£{v.toFixed(2)}</span>},
        {key:"net",label:"Net Pay",r:v=><span style={{color:"#a78bfa",fontWeight:800,fontSize:13}}>£{v.toFixed(2)}</span>},
        {key:"id",label:"Actions",r:(_,r)=><div style={{display:"flex",gap:5}}>
          <button onClick={()=>exportPayslip(r,activeDays,weekLabel,siteHours)} style={{padding:"4px 8px",background:"#0d2218",border:"1px solid #10b981",borderRadius:5,color:"#34d399",cursor:"pointer",fontSize:10,fontWeight:700}}>💷 Payslip</button>
          <button onClick={()=>openWorkerWindow(r,allSites,weekLabel,activeDays,siteHours)} style={{padding:"4px 8px",background:"#1a1f2e",border:"1px solid #60a5fa",borderRadius:5,color:"#60a5fa",cursor:"pointer",fontSize:10}}>🔗</button>
        </div>},
      ]} rows={rows}/>
    </div>
  </div>;
}

// ── Dashboard Stats Page ──────────────────────────────────────────────────────
function DStats({workers,allSites,activeDays}){
  const siteMap={};
  workers.forEach(w=>activeDays.forEach(d=>{const s=(w.days?.[d]||"").trim();if(s&&!isOff(s))siteMap[s]=(siteMap[s]||0)+1;}));
  return <div>
    <DPageHdr title="🔢 Stats" sub="Workers per site · Certificate compliance"/>
    <div style={DS.body}>
      <div style={{marginBottom:24}}>
        <div style={{fontSize:12,fontWeight:700,color:"#94a3b8",marginBottom:12}}>Workers per Site This Week</div>
        <div style={{display:"flex",flexWrap:"wrap",gap:10}}>
          {Object.entries(siteMap).sort((a,b)=>b[1]-a[1]).map(([site,cnt])=>(
            <div key={site} style={{background:"#1a1f2e",border:`1px solid ${getSiteColor(site,allSites)}`,borderRadius:10,padding:"10px 14px"}}>
              <div style={{fontSize:11,color:getSiteColor(site,allSites),fontWeight:700}}>{site}</div>
              <div style={{fontSize:22,fontWeight:900,color:"#f1f5f9"}}>{cnt}</div>
            </div>
          ))}
        </div>
      </div>
      <div style={{fontSize:12,fontWeight:700,color:"#94a3b8",marginBottom:12}}>Certificate Compliance</div>
      {CERTS.slice(0,14).map(c=>{const held=workers.filter(w=>w.certs?.[c.key]?.held).length;const pct=workers.length>0?Math.round((held/workers.length)*100):0;return <div key={c.key} style={{marginBottom:8}}>
        <div style={{display:"flex",justifyContent:"space-between",fontSize:11,color:"#94a3b8",marginBottom:3}}><span>{c.label}</span><span style={{color:pct>50?"#34d399":"#64748b"}}>{held}/{workers.length} ({pct}%)</span></div>
        <div style={{height:5,background:"#1e2535",borderRadius:3}}><div style={{height:"100%",borderRadius:3,background:pct>70?"#34d399":pct>30?"#fbbf24":"#f87171",width:`${pct}%`,transition:"width 0.4s"}}/></div>
      </div>;})}
    </div>
  </div>;
}

// ── Dashboard Bank Import Page ────────────────────────────────────────────────
function DBank({allSites,clients,setModal}){
  return <div>
    <DPageHdr title="🏦 Bank" sub="Import bank statement and categorise transactions"/>
    <div style={DS.body}>
      <div style={{background:"#111827",border:"1px solid #1e2535",borderRadius:12,padding:32,textAlign:"center"}}>
        <div style={{fontSize:40,marginBottom:14}}>🏦</div>
        <div style={{fontSize:16,fontWeight:700,color:"#f1f5f9",marginBottom:8}}>Bank Import & Categorisation</div>
        <div style={{fontSize:13,color:"#64748b",marginBottom:20,lineHeight:1.6}}>Upload your bank Excel or CSV · categorise each transaction as income or expense · allocate to sites and clients</div>
        <button onClick={()=>setModal({type:"bank"})} style={{padding:"10px 24px",background:"linear-gradient(135deg,#3b82f6,#6366f1)",border:"none",borderRadius:8,color:"#fff",cursor:"pointer",fontSize:14,fontWeight:700}}>📂 Open Bank Import Tool</button>
      </div>
    </div>
  </div>;
}

// ── Dashboard Budget Page ─────────────────────────────────────────────────────
function DBudget({workers,clients,allSites,activeDays,siteHours,scopeData,setModal}){
  return <div>
    <DPageHdr title="📐 Budget" sub="Site budgets, scopes and financial overview"/>
    <div style={DS.body}>
      {allSites.filter(s=>!isOff(s.name)).map(site=>{
        const sc=(site.scopes||[]);const vr=(site.variations||[]);
        const scopeT=sc.reduce((a,s)=>a+(s.qty*s.rate),0);
        const varT=vr.reduce((a,v)=>a+(v.type==="addition"?v.value:-v.value),0);
        let labourT=0;workers.forEach(w=>{const{bd}=calcPay(w,activeDays,siteHours);Object.values(bd).forEach(b=>{if(b.site===site.name||b.site.includes(site.name))labourT+=b.gross;});});
        const profit=scopeT+varT-labourT;
        const client=clients.find(c=>c.id===site.clientId);
        return <div key={site.id} style={{background:"#111827",border:`1px solid ${site.color}33`,borderRadius:10,padding:"13px 16px",marginBottom:10}}>
          <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:10}}>
            <span style={{width:9,height:9,borderRadius:"50%",background:site.color,flexShrink:0}}/>
            <span style={{fontWeight:700,color:site.color,fontSize:14,flex:1}}>{site.name}</span>
            {client&&<span style={DS.pill(client.color)}>{client.name}</span>}
            <button onClick={()=>setModal({type:"siteDetail",site})} style={{padding:"4px 10px",background:"#1e3a5f",border:"1px solid #3b82f6",borderRadius:5,color:"#60a5fa",cursor:"pointer",fontSize:11}}>✏️ Edit</button>
            <button onClick={()=>openSiteWindow(site,clients,workers,activeDays,siteHours)} style={{padding:"4px 10px",background:"#1a1f2e",border:"1px solid #60a5fa",borderRadius:5,color:"#60a5fa",cursor:"pointer",fontSize:11}}>🔗 Open</button>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:8}}>
            {[["Scope",`£${scopeT.toLocaleString()}`,"#60a5fa"],["Variations",(varT>=0?"+":"")+"£"+Math.abs(varT).toLocaleString(),"#fbbf24"],["Contract","£"+(scopeT+varT).toLocaleString(),"#34d399"],["Labour","£"+labourT.toFixed(0),"#f87171"],[profit>=0?"Profit":"Loss","£"+Math.abs(profit).toFixed(0),profit>=0?"#34d399":"#f87171"]].map(([l,v,c])=>(
              <div key={l} style={{background:"#0f1421",borderRadius:7,padding:"7px 9px"}}>
                <div style={{fontSize:9,color:"#64748b",textTransform:"uppercase",fontWeight:700}}>{l}</div>
                <div style={{fontSize:14,fontWeight:800,color:c,marginTop:2}}>{v}</div>
              </div>
            ))}
          </div>
        </div>;
      })}
    </div>
  </div>;
}

// ── Dashboard Finance Page ────────────────────────────────────────────────────  
function DFinance({workers,clients,allSites,activeDays,siteHours,scopeData,invoices}){
  return <div>
    <DPageHdr title="📊 Finance" sub="Full financial overview — mirroring the Schedule Finance tab"/>
    <div style={DS.body}>
      <FinancialDashboard workers={workers} clients={clients} allSites={allSites} activeDays={activeDays} siteHours={siteHours} scopeData={scopeData} invoices={invoices}/>
    </div>
  </div>;
}



// ═══════════════════════════════════════════════════════════════════════════
// TIMESHEETS — individual records, saved per worker per week
// ═══════════════════════════════════════════════════════════════════════════
function DTimesheets({workers,allSites,activeDays,siteHours,weekLabel,timesheetRecords,setTimesheetRecords,generateTimesheets,generatePayslips,payslipRecords,setPayslipRecords,setPage}){
  const weekSheets=timesheetRecords.filter(t=>t.weekLabel===weekLabel);
  const allWeeks=[...new Set(timesheetRecords.map(t=>t.weekLabel))].sort((a,b)=>new Date(b)-new Date(a));
  const [selWeek,setSelWeek]=useState(weekLabel);
  const shown=timesheetRecords.filter(t=>t.weekLabel===selWeek);

  function createTimesheets(){
    if(!window.confirm(`Generate timesheets for WC ${weekLabel} from current worker data?\n\nThis will create individual timesheet records for each working operative.`)) return;
    const existing=timesheetRecords.filter(t=>t.weekLabel!==weekLabel);
    const newSheets=generateTimesheets(weekLabel);
    if(newSheets.length===0){alert("No working operatives found for this week.");return;}
    setTimesheetRecords([...existing,...newSheets]);
    alert(`✓ Created ${newSheets.length} timesheets for WC ${weekLabel}`);
  }

  function approveAll(){
    setTimesheetRecords(recs=>recs.map(t=>t.weekLabel===selWeek?{...t,approved:true}:t));
  }

  function generatePaysFromTimesheets(){
    if(!window.confirm(`Generate payslips from approved timesheets for WC ${selWeek}?`)) return;
    const approved=shown.filter(t=>t.approved);
    if(approved.length===0){alert("No approved timesheets found.");return;}
    const existing=payslipRecords.filter(p=>p.weekLabel!==selWeek);
    const newPays=generatePayslips(approved);
    setPayslipRecords([...existing,...newPays]);
    alert(`✓ Generated ${newPays.length} payslips for WC ${selWeek}`);
    setPage("payslips");
  }

  function updateSheet(id,key,val){
    setTimesheetRecords(recs=>recs.map(t=>t.id===id?{...t,[key]:val}:t));
  }

  const totGross=shown.reduce((a,t)=>a+t.gross,0);
  const totStd=shown.reduce((a,t)=>a+t.stdHours,0);
  const totOT=shown.reduce((a,t)=>a+t.otHours,0);

  return <div>
    <DPageHdr title="⏱ Timesheets" sub={`${timesheetRecords.length} total records`}
      actions={<div style={{display:"flex",gap:7}}>
        <button onClick={createTimesheets} style={{padding:"6px 13px",background:"linear-gradient(135deg,#3b82f6,#6366f1)",border:"none",borderRadius:7,color:"#fff",cursor:"pointer",fontSize:12,fontWeight:700}}>+ Generate WC {weekLabel}</button>
        {shown.length>0&&<button onClick={approveAll} style={{padding:"6px 13px",background:"#0d2218",border:"1px solid #10b981",borderRadius:7,color:"#34d399",cursor:"pointer",fontSize:12,fontWeight:700}}>✓ Approve All</button>}
        {shown.some(t=>t.approved)&&<button onClick={generatePaysFromTimesheets} style={{padding:"6px 13px",background:"linear-gradient(135deg,#059669,#10b981)",border:"none",borderRadius:7,color:"#fff",cursor:"pointer",fontSize:12,fontWeight:700}}>💷 → Generate Payslips</button>}
      </div>}/>
    <div style={DS.body}>
      {/* Week selector */}
      <div style={{display:"flex",gap:6,marginBottom:16,flexWrap:"wrap"}}>
        {allWeeks.length===0&&<div style={{color:"#374151",fontSize:12}}>No timesheets yet. Click "+ Generate" to create timesheets from current worker data.</div>}
        {allWeeks.map(wk=><button key={wk} onClick={()=>setSelWeek(wk)} style={{padding:"5px 12px",background:selWeek===wk?"#1e3a5f":"#1a1f2e",border:`1px solid ${selWeek===wk?"#3b82f6":"#2d3555"}`,borderRadius:7,color:selWeek===wk?"#60a5fa":"#64748b",cursor:"pointer",fontSize:11,fontWeight:selWeek===wk?700:400}}>WC {wk} ({timesheetRecords.filter(t=>t.weekLabel===wk).length})</button>)}
      </div>

      {shown.length>0&&<>
        {/* Summary */}
        <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10,marginBottom:16}}>
          <DStat label="Timesheets" value={shown.length} color="#60a5fa"/>
          <DStat label="Std Hours" value={totStd+"h"} color="#34d399"/>
          <DStat label="OT Hours" value={totOT>0?totOT+"h":"—"} color="#fbbf24"/>
          <DStat label="Gross" value={"£"+totGross.toFixed(0)} color="#a78bfa"/>
        </div>

        {/* Timesheet table */}
        <div style={{border:"1px solid #1e2535",borderRadius:10,overflow:"hidden"}}>
          <table style={{width:"100%",borderCollapse:"collapse"}}>
            <thead><tr>
              <th style={DS.th}>Worker</th><th style={DS.th}>Position</th>
              <th style={DS.th}>Std h</th><th style={DS.th}>OT h</th>
              <th style={DS.th}>Rate</th><th style={DS.th}>Gross</th>
              <th style={DS.th}>Tax</th><th style={DS.th}>Net</th>
              <th style={DS.th}>Status</th><th style={DS.th}>Actions</th>
            </tr></thead>
            <tbody>{shown.map((t,i)=>(
              <tr key={t.id} style={{background:i%2===0?"#111827":"#0f1421"}}>
                <td style={{...DS.td,fontWeight:600,color:"#f1f5f9"}}>{t.workerName}</td>
                <td style={{...DS.td,color:"#94a3b8",fontSize:11}}>{t.position}</td>
                <td style={DS.td}><input type="number" value={t.stdHours} onChange={e=>updateSheet(t.id,"stdHours",Number(e.target.value))} style={{width:55,background:"#0f1421",border:"1px solid #2d3555",borderRadius:4,padding:"2px 5px",color:"#34d399",fontSize:11,textAlign:"right",outline:"none"}}/></td>
                <td style={DS.td}><input type="number" value={t.otHours} onChange={e=>updateSheet(t.id,"otHours",Number(e.target.value))} style={{width:55,background:"#0f1421",border:"1px solid #2d3555",borderRadius:4,padding:"2px 5px",color:"#fbbf24",fontSize:11,textAlign:"right",outline:"none"}}/></td>
                <td style={{...DS.td,color:"#34d399",fontWeight:600}}>{t.rate?`£${t.rate}/hr`:"—"}</td>
                <td style={{...DS.td,color:"#34d399",fontWeight:700}}>£{t.gross.toFixed(2)}</td>
                <td style={{...DS.td,color:"#f87171"}}>£{t.tax.toFixed(2)}</td>
                <td style={{...DS.td,color:"#a78bfa",fontWeight:700}}>£{t.net.toFixed(2)}</td>
                <td style={DS.td}>
                  <button onClick={()=>updateSheet(t.id,"approved",!t.approved)}
                    style={{padding:"3px 9px",background:t.approved?"#0d2218":"#1a1f2e",border:`1px solid ${t.approved?"#10b981":"#2d3555"}`,borderRadius:5,color:t.approved?"#34d399":"#64748b",cursor:"pointer",fontSize:10,fontWeight:700}}>
                    {t.approved?"✓ Approved":"Pending"}
                  </button>
                </td>
                <td style={DS.td}>
                  <div style={{display:"flex",gap:4}}>
                    <button onClick={()=>exportPayslip({...t,agreedRate:t.rate,taxRate:t.taxRate,days:t.days||{},overtimeHours:{},hoursPerDay:{}},activeDays,t.weekLabel,siteHours||{})} style={{padding:"3px 7px",background:"#0d2218",border:"1px solid #10b981",borderRadius:4,color:"#34d399",cursor:"pointer",fontSize:10}}>💷</button>
                    <button onClick={()=>setTimesheetRecords(recs=>recs.filter(x=>x.id!==t.id))} style={{padding:"3px 7px",background:"#2d1515",border:"1px solid #ef4444",borderRadius:4,color:"#f87171",cursor:"pointer",fontSize:10}}>✕</button>
                  </div>
                </td>
              </tr>
            ))}</tbody>
          </table>
        </div>
      </>}
    </div>
  </div>;
}

// ═══════════════════════════════════════════════════════════════════════════
// PAYSLIPS — individual records, permanent in system
// ═══════════════════════════════════════════════════════════════════════════
function DPayslips({workers,allSites,activeDays,siteHours,weekLabel,payslipRecords,setPayslipRecords,timesheetRecords,generatePayslips,setPage}){
  const allWeeks=[...new Set(payslipRecords.map(p=>p.weekLabel))].sort((a,b)=>new Date(b)-new Date(a));
  const [selWeek,setSelWeek]=useState(payslipRecords[0]?.weekLabel||weekLabel);
  const shown=payslipRecords.filter(p=>p.weekLabel===selWeek);
  const totGross=shown.reduce((a,p)=>a+p.gross,0);
  const totNet=shown.reduce((a,p)=>a+p.net,0);
  const totTax=shown.reduce((a,p)=>a+p.tax,0);

  function markIssued(){
    setPayslipRecords(recs=>recs.map(p=>p.weekLabel===selWeek?{...p,issued:true}:p));
  }

  function genFromTimesheets(){
    const sheets=timesheetRecords.filter(t=>t.weekLabel===weekLabel&&t.approved);
    if(sheets.length===0){alert("No approved timesheets for current week. Go to Timesheets, approve them, then return.");return;}
    const ex=payslipRecords.filter(p=>p.weekLabel!==weekLabel);
    const newP=generatePayslips(sheets);
    setPayslipRecords([...ex,...newP]);
    alert(`✓ Generated ${newP.length} payslips for WC ${weekLabel}`);
  }

  return <div>
    <DPageHdr title="💷 Payslips" sub={`${payslipRecords.length} total · ${payslipRecords.filter(p=>p.issued).length} issued`}
      actions={<div style={{display:"flex",gap:7}}>
        <button onClick={genFromTimesheets} style={{padding:"6px 13px",background:"linear-gradient(135deg,#3b82f6,#6366f1)",border:"none",borderRadius:7,color:"#fff",cursor:"pointer",fontSize:12,fontWeight:700}}>+ From Timesheets WC {weekLabel}</button>
        {shown.length>0&&!shown.every(p=>p.issued)&&<button onClick={markIssued} style={{padding:"6px 13px",background:"#0d2218",border:"1px solid #10b981",borderRadius:7,color:"#34d399",cursor:"pointer",fontSize:12,fontWeight:700}}>✓ Mark All Issued</button>}
      </div>}/>
    <div style={DS.body}>
      {/* Week tabs */}
      <div style={{display:"flex",gap:6,marginBottom:16,flexWrap:"wrap"}}>
        {payslipRecords.length===0&&<div style={{color:"#374151",fontSize:12}}>No payslips yet. Generate from approved timesheets.</div>}
        {allWeeks.map(wk=>{const wkPays=payslipRecords.filter(p=>p.weekLabel===wk);return<button key={wk} onClick={()=>setSelWeek(wk)} style={{padding:"5px 12px",background:selWeek===wk?"#1e3a5f":"#1a1f2e",border:`1px solid ${selWeek===wk?"#3b82f6":"#2d3555"}`,borderRadius:7,color:selWeek===wk?"#60a5fa":"#64748b",cursor:"pointer",fontSize:11,fontWeight:selWeek===wk?700:400}}>WC {wk} ({wkPays.length}) {wkPays.every(p=>p.issued)?"✓":""}</button>;})}
      </div>

      {shown.length>0&&<>
        <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10,marginBottom:16}}>
          <DStat label="Payslips" value={shown.length} color="#60a5fa"/>
          <DStat label="Gross" value={"£"+totGross.toFixed(0)} color="#34d399"/>
          <DStat label="Tax" value={"£"+totTax.toFixed(0)} color="#f87171"/>
          <DStat label="Net Pay" value={"£"+totNet.toFixed(0)} color="#a78bfa"/>
        </div>

        <div style={{border:"1px solid #1e2535",borderRadius:10,overflow:"hidden"}}>
          <table style={{width:"100%",borderCollapse:"collapse"}}>
            <thead><tr>
              <th style={DS.th}>Worker</th><th style={DS.th}>Week</th>
              <th style={DS.th}>Std h</th><th style={DS.th}>OT h</th>
              <th style={DS.th}>Gross</th><th style={DS.th}>Tax</th>
              <th style={DS.th}>Net Pay</th><th style={DS.th}>Status</th><th style={DS.th}>Export</th>
            </tr></thead>
            <tbody>{shown.map((p,i)=>(
              <tr key={p.id} style={{background:i%2===0?"#111827":"#0f1421"}}>
                <td style={{...DS.td,fontWeight:600,color:"#f1f5f9"}}>{p.workerName}<div style={{fontSize:10,color:"#64748b"}}>{p.position}</div></td>
                <td style={{...DS.td,color:"#94a3b8",fontSize:11}}>{p.weekLabel}</td>
                <td style={{...DS.td,color:"#60a5fa",fontWeight:600}}>{p.stdHours}h</td>
                <td style={{...DS.td,color:"#fbbf24"}}>{p.otHours>0?p.otHours+"h":"—"}</td>
                <td style={{...DS.td,color:"#34d399",fontWeight:700}}>£{p.gross.toFixed(2)}</td>
                <td style={{...DS.td,color:"#f87171"}}>£{p.tax.toFixed(2)}</td>
                <td style={{...DS.td,color:"#a78bfa",fontWeight:800,fontSize:13}}>£{p.net.toFixed(2)}</td>
                <td style={DS.td}><span style={{padding:"2px 8px",borderRadius:4,fontSize:10,fontWeight:700,color:p.issued?"#34d399":"#fbbf24",background:p.issued?"#0d2218":"#1a1500"}}>{p.issued?"✓ Issued":"Pending"}</span></td>
                <td style={DS.td}>
                  <div style={{display:"flex",gap:4}}>
                    <button onClick={()=>exportPayslip({id:p.workerId,name:p.workerName,position:p.position,agreedRate:p.rate,taxRate:p.taxRate,overtimeMultiplier:1.5,customOTRate:null,days:{},overtimeHours:{},hoursPerDay:{}},activeDays,p.weekLabel,siteHours||{})} style={{padding:"3px 7px",background:"#0d2218",border:"1px solid #10b981",borderRadius:4,color:"#34d399",cursor:"pointer",fontSize:10,fontWeight:700}}>💷 PDF</button>
                    <button onClick={()=>setPayslipRecords(recs=>recs.filter(x=>x.id!==p.id))} style={{padding:"3px 7px",background:"#2d1515",border:"1px solid #ef4444",borderRadius:4,color:"#f87171",cursor:"pointer",fontSize:10}}>✕</button>
                  </div>
                </td>
              </tr>
            ))}</tbody>
            <tfoot><tr style={{background:"#0d1117",borderTop:"2px solid #2d3555"}}>
              <td colSpan={4} style={{...DS.td,fontWeight:700,color:"#94a3b8"}}>TOTALS</td>
              <td style={{...DS.td,color:"#34d399",fontWeight:800}}>£{totGross.toFixed(2)}</td>
              <td style={{...DS.td,color:"#f87171",fontWeight:800}}>£{totTax.toFixed(2)}</td>
              <td style={{...DS.td,color:"#a78bfa",fontWeight:900,fontSize:14}}>£{totNet.toFixed(2)}</td>
              <td colSpan={2} style={DS.td}/>
            </tr></tfoot>
          </table>
        </div>
      </>}
    </div>
  </div>;
}

// ═══════════════════════════════════════════════════════════════════════════
// BANK IMPORT — save transactions to system
// ═══════════════════════════════════════════════════════════════════════════
function DBankFull({allSites,clients,bankTransactions,setBankTransactions,setModal}){
  const [txns,setTxns]=useState([]);
  const [fileName,setFileName]=useState("");
  const INCOME_CATS=["Client Payment","Contract Payment","Variation Payment","Retention Release","Other Income"];
  const EXPENSE_CATS=["Materials","Plant Hire","Subcontractor","Labour (External)","Transport","Insurance","Tools & Equipment","Professional Fees","Utilities","Office","Other Expense"];

  // Convert Excel serial date number to readable string
  function excelDateToString(v){
    if(!v&&v!==0) return "";
    // If already a date string, return as-is
    if(typeof v==="string"&&v.includes("-")||typeof v==="string"&&v.includes("/")) return v;
    // Excel serial date: days since 1900-01-01
    if(typeof v==="number"&&v>1000&&v<100000){
      const d=new Date(Math.round((v-25569)*86400*1000));
      return d.toLocaleDateString("en-GB",{day:"2-digit",month:"2-digit",year:"numeric"});
    }
    return String(v||"");
  }

  const handleFile=e=>{
    const f=e.target.files[0];if(!f)return;setFileName(f.name);
    const reader=new FileReader();
    reader.onload=ev=>{
      try{
        const wb=XLSX.read(ev.target.result,{type:"binary",cellDates:false});
        const ws=wb.Sheets[wb.SheetNames[0]];
        const data=XLSX.utils.sheet_to_json(ws,{header:1,raw:true});
        // Auto-detect header row (skip empty rows at top)
        const startRow=data.findIndex(r=>r.some(c=>c!==undefined&&c!==""));
        const header=(data[startRow]||[]).map(h=>String(h||"").toLowerCase().trim());
        const dataRows=data.slice(startRow+1).filter(r=>r.some(c=>c!==undefined&&c!==""));

        // Smart column detection by header name
        let dateCol=-1,descCol=-1,amtCol=-1,creditCol=-1,debitCol=-1;
        header.forEach((h,i)=>{
          if(h.includes("date")) dateCol=i;
          if(h.includes("desc")||h.includes("narr")||h.includes("detail")||h.includes("memo")||h.includes("ref")) descCol=i;
          if(h.includes("amount")&&!h.includes("credit")&&!h.includes("debit")) amtCol=i;
          if(h.includes("credit")) creditCol=i;
          if(h.includes("debit")) debitCol=i;
        });
        // Fallback: if no header found, guess by column position
        if(dateCol===-1) dateCol=0;
        if(descCol===-1) descCol=1;
        if(amtCol===-1&&creditCol===-1) amtCol=2;

        setTxns(dataRows.map(r=>{
          const rawDate=r[dateCol];
          const desc=String(r[descCol]||r[descCol+1]||"").trim();
          let amount=0;
          if(creditCol>-1||debitCol>-1){
            // Separate credit/debit columns
            const cr=parseFloat(String(r[creditCol]||"0").replace(/[£,]/g,""))||0;
            const dr=parseFloat(String(r[debitCol]||"0").replace(/[£,]/g,""))||0;
            amount=cr>0?cr:-dr;
          } else {
            amount=parseFloat(String(r[amtCol]||"0").replace(/[£,]/g,""))||0;
          }
          return {
            id:"bt_"+Date.now()+Math.random().toString(36).slice(2),
            date:excelDateToString(rawDate),
            description:desc,
            amount:amount,
            type:amount>=0?"income":"expense",
            category:"",siteId:"",clientId:"",notes:"",saved:false,
          };
        }).filter(t=>t.description||t.amount!==0));
      }catch(err){alert("Error reading file: "+err.message+"\n\nMake sure the file has columns: Date, Description, Amount");}
    };
    reader.readAsBinaryString(f);
  };

  const upT=(id,k,v)=>setTxns(t=>t.map(x=>x.id===id?{...x,[k]:v}:x));

  function saveToSystem(){
    const toSave=txns.filter(t=>t.category);
    if(toSave.length===0){alert("Please categorise at least one transaction before saving.");return;}
    const existing=bankTransactions.filter(bt=>!txns.find(t=>t.id===bt.id));
    setBankTransactions([...existing,...txns.map(t=>({...t,saved:true}))]);
    alert(`✓ ${toSave.length} transactions saved to system.`);
  }

  function exportCat(){
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([
      ["Date","Description","Amount","Type","Category","Site","Client","Notes"],
      ...txns.map(t=>[t.date,t.description,t.amount,t.type,t.category,
        allSites.find(s=>s.id===t.siteId)?.name||"",
        clients.find(c=>c.id===t.clientId)?.name||"",t.notes])
    ]),"Transactions");
    XLSX.writeFile(wb,"Bank_Categorised_"+new Date().toLocaleDateString("en-GB").replace(/\//g,"-")+".xlsx");
  }

  const saved=bankTransactions.length;
  const income=txns.filter(t=>t.type==="income").reduce((a,t)=>a+Math.abs(t.amount),0);
  const expense=txns.filter(t=>t.type==="expense").reduce((a,t)=>a+Math.abs(t.amount),0);
  const allSavedIncome=bankTransactions.filter(t=>t.type==="income").reduce((a,t)=>a+Math.abs(t.amount),0);
  const allSavedExpense=bankTransactions.filter(t=>t.type==="expense").reduce((a,t)=>a+Math.abs(t.amount),0);

  return <div>
    <DPageHdr title="🏦 Bank Import" sub={`${saved} transactions saved in system`}
      actions={<div style={{display:"flex",gap:7}}>
        {txns.length>0&&<button onClick={saveToSystem} style={{padding:"6px 13px",background:"linear-gradient(135deg,#059669,#10b981)",border:"none",borderRadius:7,color:"#fff",cursor:"pointer",fontSize:12,fontWeight:700}}>💾 Save to System ({txns.filter(t=>t.category).length} categorised)</button>}
        {txns.length>0&&<button onClick={exportCat} style={{padding:"6px 13px",background:"#1a1f2e",border:"1px solid #3b82f6",borderRadius:7,color:"#60a5fa",cursor:"pointer",fontSize:12}}>⬇ Export Excel</button>}
      </div>}/>
    <div style={DS.body}>
      {/* Saved transactions summary */}
      {saved>0&&<div style={{background:"#0a1a0a",border:"1px solid #16a34a33",borderRadius:10,padding:"12px 16px",marginBottom:16}}>
        <div style={{fontSize:11,color:"#34d399",fontWeight:700,textTransform:"uppercase",marginBottom:8}}>Saved in System ({saved} transactions)</div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10}}>
          <DStat label="Total Income" value={"£"+allSavedIncome.toLocaleString()} color="#34d399"/>
          <DStat label="Total Expenses" value={"£"+allSavedExpense.toLocaleString()} color="#f87171"/>
          <DStat label="Net Position" value={"£"+(allSavedIncome-allSavedExpense).toLocaleString()} color={(allSavedIncome-allSavedExpense)>=0?"#34d399":"#f87171"}/>
        </div>
      </div>}

      {/* Import area */}
      <div style={{background:"#111827",border:"1px solid #1e2535",borderRadius:10,padding:16,marginBottom:16}}>
        <div style={{fontSize:11,color:"#64748b",fontWeight:700,textTransform:"uppercase",marginBottom:10}}>Import New Statement</div>
        <label style={{display:"block",padding:"12px 16px",background:"#1e3a5f",border:"2px dashed #3b82f6",borderRadius:8,cursor:"pointer",textAlign:"center",color:"#60a5fa",fontSize:12,fontWeight:600}}>
          📁 {fileName||"Click to upload Excel / CSV · Date | Description | Amount"}
          <input type="file" accept=".xlsx,.xls,.csv" onChange={handleFile} style={{display:"none"}}/>
        </label>
        {txns.length>0&&<div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:8,marginTop:10}}>
          <DStat label="Transactions" value={txns.length} color="#60a5fa"/>
          <DStat label="Income" value={"£"+income.toLocaleString()} color="#34d399"/>
          <DStat label="Expenses" value={"£"+expense.toLocaleString()} color="#f87171"/>
          <DStat label="Categorised" value={txns.filter(t=>t.category).length+"/"+txns.length} color="#fbbf24"/>
        </div>}
      </div>

      {txns.length>0&&<div style={{border:"1px solid #1e2535",borderRadius:10,overflow:"hidden"}}>
        <div style={{maxHeight:480,overflowY:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse"}}>
            <thead style={{position:"sticky",top:0,zIndex:1}}><tr>
              <th style={{...DS.th,minWidth:85}}>Date</th>
              <th style={{...DS.th,minWidth:180}}>Description</th>
              <th style={{...DS.th,minWidth:80}}>Amount</th>
              <th style={{...DS.th,minWidth:90}}>Type</th>
              <th style={{...DS.th,minWidth:150}}>Category</th>
              <th style={{...DS.th,minWidth:120}}>Site</th>
              <th style={{...DS.th,minWidth:110}}>Client</th>
              <th style={{...DS.th,minWidth:100}}>Notes</th>
            </tr></thead>
            <tbody>{txns.map((t,i)=>(
              <tr key={t.id} style={{background:i%2===0?"#111827":"#0f1421"}}>
                <td style={{...DS.td,color:"#94a3b8",fontSize:11}}>{t.date}</td>
                <td style={{...DS.td,maxWidth:180}}><div style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",color:"#e2e8f0",fontSize:12}} title={t.description}>{t.description}</div></td>
                <td style={{...DS.td,fontWeight:700,color:t.amount>=0?"#34d399":"#f87171",whiteSpace:"nowrap"}}>£{Math.abs(t.amount).toFixed(2)}</td>
                <td style={DS.td}><select value={t.type} onChange={e=>upT(t.id,"type",e.target.value)} style={{...{width:"100%",background:"#0f1421",border:"1px solid #2d3555",borderRadius:4,padding:"3px 5px",color:t.type==="income"?"#34d399":"#f87171",fontSize:10,outline:"none"},cursor:"pointer"}}>
                  <option value="income">Income</option><option value="expense">Expense</option>
                </select></td>
                <td style={DS.td}><select value={t.category} onChange={e=>upT(t.id,"category",e.target.value)} style={{width:"100%",background:"#0f1421",border:"1px solid #2d3555",borderRadius:4,padding:"3px 5px",color:t.category?"#e2e8f0":"#64748b",fontSize:10,outline:"none",cursor:"pointer"}}>
                  <option value="">— Category —</option>
                  <optgroup label="Income">{INCOME_CATS.map(c=><option key={c} value={c}>{c}</option>)}</optgroup>
                  <optgroup label="Expenses">{EXPENSE_CATS.map(c=><option key={c} value={c}>{c}</option>)}</optgroup>
                </select></td>
                <td style={DS.td}><select value={t.siteId} onChange={e=>upT(t.id,"siteId",e.target.value)} style={{width:"100%",background:"#0f1421",border:"1px solid #2d3555",borderRadius:4,padding:"3px 5px",color:"#e2e8f0",fontSize:10,outline:"none",cursor:"pointer"}}>
                  <option value="">— Site —</option>
                  {allSites.filter(s=>!isOff(s.name)).map(s=><option key={s.id} value={s.id}>{s.name}</option>)}
                </select></td>
                <td style={DS.td}><select value={t.clientId} onChange={e=>upT(t.id,"clientId",e.target.value)} style={{width:"100%",background:"#0f1421",border:"1px solid #2d3555",borderRadius:4,padding:"3px 5px",color:"#e2e8f0",fontSize:10,outline:"none",cursor:"pointer"}}>
                  <option value="">— Client —</option>
                  {clients.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}
                </select></td>
                <td style={DS.td}><input value={t.notes} onChange={e=>upT(t.id,"notes",e.target.value)} placeholder="Notes…" style={{width:"100%",background:"#0f1421",border:"1px solid #2d3555",borderRadius:4,padding:"3px 5px",color:"#e2e8f0",fontSize:10,outline:"none"}}/></td>
              </tr>
            ))}</tbody>
          </table>
        </div>
      </div>}
    </div>
  </div>;
}

// ═══════════════════════════════════════════════════════════════════════════
// PAYMENT APPLICATIONS — per site, scopes+variations+dayworks+prelims
// ═══════════════════════════════════════════════════════════════════════════
const PA_ITEM_TYPES=["Scope of Work","Variation","Dayworks","Preliminaries","Other"];
const PA_UNITS=["l/m","m²","m³","nr","kg","tonne","day","week","%","sum","hrs"];

function DPayApps({allSites,clients,workers,activeDays,siteHours,scopeData,payApplications,setPayApplications,setPage,setDetailId}){
  const [selSite,setSelSite]=useState("");
  const activeSites=allSites.filter(s=>!isOff(s.name));

  function newPayApp(){
    if(!selSite){alert("Please select a site first.");return;}
    const site=allSites.find(s=>s.id===selSite);
    if(!site){return;}
    // Build items from site scopes and variations
    const items=[
      ...(site.scopes||[]).map(sc=>({
        id:"pai_"+Date.now()+Math.random().toString(36).slice(2),
        type:"Scope of Work",description:sc.description||sc.desc||"",
        unit:sc.unit||"sum",contractQty:sc.qty||0,contractRate:sc.rate||sc.unitIncome||0,
        claimedQtyToDate:0,claimedPctToDate:0,useQty:true,previousQty:0,
      })),
      ...(site.variations||[]).map(vr=>({
        id:"pai_"+Date.now()+Math.random().toString(36).slice(2),
        type:"Variation",description:vr.description||vr.desc||"",
        unit:"sum",contractQty:1,contractRate:vr.value||0,
        claimedQtyToDate:0,claimedPctToDate:0,useQty:false,previousQty:0,
      })),
    ];
    const pa={
      id:"pa_"+Date.now(),
      siteId:selSite,siteName:site.name,
      clientId:site.clientId||"",
      number:"PA-"+(payApplications.filter(p=>p.siteId===selSite).length+1).toString().padStart(3,"0"),
      date:new Date().toISOString().slice(0,10),
      status:"draft",items,
      createdAt:new Date().toISOString(),
    };
    setPayApplications(pas=>[...pas,pa]);
    setDetailId(pa.id);
    setPage("payapp_detail");
  }

  return <div>
    <DPageHdr title="📐 Payment Applications" sub={`${payApplications.length} applications across ${[...new Set(payApplications.map(p=>p.siteId))].length} sites`}
      actions={<div style={{display:"flex",gap:8,alignItems:"flex-end"}}>
        <div>
          <label style={{...DS.th,display:"block",marginBottom:3}}>Select Site</label>
          <select value={selSite} onChange={e=>setSelSite(e.target.value)} style={{background:"#0f1421",border:"1px solid #2d3555",borderRadius:7,padding:"6px 10px",color:"#e2e8f0",fontSize:12,outline:"none",cursor:"pointer",minWidth:200}}>
            <option value="">— Choose site —</option>
            {activeSites.map(s=><option key={s.id} value={s.id}>{s.name}</option>)}
          </select>
        </div>
        <button onClick={newPayApp} style={{padding:"6px 14px",background:"linear-gradient(135deg,#3b82f6,#6366f1)",border:"none",borderRadius:7,color:"#fff",cursor:"pointer",fontSize:12,fontWeight:700}}>+ New Application</button>
      </div>}/>
    <div style={DS.body}>
      {payApplications.length===0&&<div style={{textAlign:"center",padding:"60px 24px",border:"1px dashed #1e2535",borderRadius:12}}>
        <div style={{fontSize:40,marginBottom:14}}>📐</div>
        <div style={{fontSize:16,fontWeight:700,color:"#f1f5f9",marginBottom:8}}>No Payment Applications Yet</div>
        <div style={{fontSize:13,color:"#64748b"}}>Select a site above and click "+ New Application" to create a payment application from the site's scopes, variations, dayworks and preliminaries.</div>
      </div>}
      {/* Group by site */}
      {activeSites.filter(s=>payApplications.some(p=>p.siteId===s.id)).map(site=>{
        const sitePAs=payApplications.filter(p=>p.siteId===site.id).sort((a,b)=>new Date(b.createdAt)-new Date(a.createdAt));
        return <div key={site.id} style={{marginBottom:20}}>
          <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:10}}>
            <span style={{width:10,height:10,borderRadius:"50%",background:site.color}}/><span style={{fontWeight:700,color:site.color,fontSize:14}}>{site.name}</span>
            <span style={{fontSize:11,color:"#64748b"}}>{sitePAs.length} application{sitePAs.length!==1?"s":""}</span>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10}}>
            {sitePAs.map(pa=>{
              const total=pa.items.reduce((a,it)=>{const v=it.useQty?(it.claimedQtyToDate||0)*it.contractRate:((it.claimedPctToDate||0)/100)*(it.contractQty*it.contractRate);return a+v;},0);
              const contract=pa.items.reduce((a,it)=>a+it.contractQty*it.contractRate,0);
              return <div key={pa.id} onClick={()=>{setDetailId(pa.id);setPage("payapp_detail");}}
                style={{background:"#111827",border:"1px solid #1e2535",borderRadius:10,padding:14,cursor:"pointer",transition:"all 0.15s"}}
                onMouseEnter={e=>{e.currentTarget.style.borderColor=site.color;e.currentTarget.style.transform="translateY(-2px)";}}
                onMouseLeave={e=>{e.currentTarget.style.borderColor="#1e2535";e.currentTarget.style.transform="";}}>
                <div style={{display:"flex",justifyContent:"space-between",marginBottom:8}}>
                  <span style={{fontWeight:700,color:"#a78bfa",fontSize:13}}>{pa.number}</span>
                  <span style={{padding:"2px 8px",borderRadius:4,fontSize:10,fontWeight:700,color:pa.status==="submitted"?"#60a5fa":pa.status==="certified"?"#34d399":"#94a3b8",background:pa.status==="submitted"?"#0d1a2e":pa.status==="certified"?"#0d2218":"#1e2535",textTransform:"capitalize"}}>{pa.status}</span>
                </div>
                <div style={{fontSize:11,color:"#64748b",marginBottom:8}}>{pa.date} · {pa.items.length} items</div>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:6}}>
                  <div style={{background:"#0f1421",borderRadius:6,padding:"6px 8px"}}><div style={{fontSize:9,color:"#64748b",textTransform:"uppercase"}}>Claimed</div><div style={{fontSize:14,fontWeight:800,color:"#34d399"}}>£{total.toLocaleString(undefined,{maximumFractionDigits:0})}</div></div>
                  <div style={{background:"#0f1421",borderRadius:6,padding:"6px 8px"}}><div style={{fontSize:9,color:"#64748b",textTransform:"uppercase"}}>Contract</div><div style={{fontSize:14,fontWeight:800,color:"#60a5fa"}}>£{contract.toLocaleString(undefined,{maximumFractionDigits:0})}</div></div>
                </div>
              </div>;
            })}
          </div>
        </div>;
      })}
    </div>
  </div>;
}

// ── Payment Application Detail ─────────────────────────────────────────────────
function DPayAppDetail({payApplications,setPayApplications,payappId,allSites,clients,setPage}){
  const pa=payApplications.find(p=>p.id===payappId);
  const [items,setItems]=useState(pa?JSON.parse(JSON.stringify(pa.items)):[]);
  const [addType,setAddType]=useState("Dayworks");
  if(!pa) return <div style={DS.body}><div style={{color:"#374151",textAlign:"center",padding:40}}>Application not found.</div></div>;
  const site=allSites.find(s=>s.id===pa.siteId);
  const client=clients.find(c=>c.id===pa.clientId);

  const calcClaimed=(it)=>{
    if(it.useQty) return (it.claimedQtyToDate||0)*it.contractRate;
    return ((it.claimedPctToDate||0)/100)*(it.contractQty*it.contractRate);
  };
  const calcPrev=(it)=>(it.previousQty||0)*it.contractRate;
  const calcThisPeriod=(it)=>calcClaimed(it)-calcPrev(it);

  const totalContract=items.reduce((a,it)=>a+it.contractQty*it.contractRate,0);
  const totalClaimed=items.reduce((a,it)=>a+calcClaimed(it),0);
  const totalPrev=items.reduce((a,it)=>a+calcPrev(it),0);
  const totalThis=totalClaimed-totalPrev;
  const pctComplete=totalContract>0?Math.round((totalClaimed/totalContract)*100):0;

  function updateItem(id,key,val){setItems(its=>its.map(it=>it.id===id?{...it,[key]:val}:it));}
  function toggleMode(id){setItems(its=>its.map(it=>it.id===id?{...it,useQty:!it.useQty}:it));}
  function addItem(){
    setItems(its=>[...its,{id:"pai_"+Date.now(),type:addType,description:"",unit:"sum",contractQty:1,contractRate:0,claimedQtyToDate:0,claimedPctToDate:0,useQty:addType!=="Preliminaries",previousQty:0}]);
  }
  function deleteItem(id){setItems(its=>its.filter(it=>it.id!==id));}

  function savePA(){
    setPayApplications(pas=>pas.map(p=>p.id===payappId?{...p,items,status:"draft"}:p));
    alert("✓ Saved");
  }
  function submitPA(){
    setPayApplications(pas=>pas.map(p=>p.id===payappId?{...p,items,status:"submitted"}:p));
    alert("✓ Submitted");
    setPage("payapps");
  }

  function exportPDF(){
    const html=`<!DOCTYPE html><html><head><meta charset="utf-8"/><title>${pa.number} — ${site?.name||""}</title>
<style>*{margin:0;padding:0;box-sizing:border-box;}body{background:#0d1117;color:#e2e8f0;font-family:'Segoe UI',Arial,sans-serif;font-size:12px;padding:24px;}
.hdr{display:flex;justify-content:space-between;margin-bottom:20px;padding-bottom:14px;border-bottom:2px solid #1e2535;}
table{width:100%;border-collapse:collapse;margin-bottom:16px;}
th{padding:7px 10px;text-align:left;font-size:10px;font-weight:700;color:#64748b;text-transform:uppercase;border-bottom:2px solid #1e2535;background:#0a0e17;}
td{padding:6px 10px;border-bottom:1px solid #1a2030;font-size:11px;}tr:nth-child(even) td{background:#111827;}tr:nth-child(odd) td{background:#0f1421;}
.total{background:#0d2218;border:2px solid #10b981;border-radius:10px;padding:16px 20px;margin-top:10px;}
.ft{margin-top:16px;padding-top:10px;border-top:1px solid #1e2535;display:flex;justify-content:space-between;font-size:9px;color:#374151;}
@media print{@page{margin:8mm;size:A3 landscape;}}</style></head><body>
<div class="hdr">
  <div><div style="font-size:22px;font-weight:800;color:#f1f5f9">${pa.number}</div>
  <div style="font-size:13px;color:#64748b">${site?.name||"—"} · ${client?.name||"—"}</div>
  <div style="font-size:11px;color:#64748b;margin-top:4px">Date: ${pa.date} · Status: ${pa.status}</div></div>
  <div style="text-align:right"><div style="font-size:11px;color:#64748b">Contract Value</div><div style="font-size:24px;font-weight:900;color:#60a5fa">£${totalContract.toLocaleString(undefined,{maximumFractionDigits:0})}</div></div>
</div>
<table><thead><tr><th>Type</th><th style="width:35%">Description</th><th>Unit</th><th>Qty</th><th>Rate £</th><th>Contract £</th><th>Prev Qty</th><th>Claimed ToDate</th><th>This Period £</th><th>%</th></tr></thead><tbody>
${items.map(it=>{const cl=calcClaimed(it);const pr=calcPrev(it);const th=cl-pr;const pct=it.contractQty*it.contractRate>0?Math.round((cl/(it.contractQty*it.contractRate))*100):0;return `<tr><td style="color:#a78bfa">${it.type}</td><td style="font-weight:600">${it.description||"—"}</td><td>${it.unit}</td><td style="text-align:right">${it.contractQty}</td><td style="text-align:right">£${it.contractRate.toLocaleString()}</td><td style="text-align:right;font-weight:700">£${(it.contractQty*it.contractRate).toLocaleString()}</td><td style="text-align:right;color:#64748b">${it.previousQty||0}</td><td style="text-align:right;color:#34d399;font-weight:700">£${cl.toLocaleString(undefined,{maximumFractionDigits:0})}</td><td style="text-align:right;color:${th>=0?"#34d399":"#f87171"};font-weight:700">£${th.toLocaleString(undefined,{maximumFractionDigits:0})}</td><td style="text-align:right">${pct}%</td></tr>`;}).join("")}
</tbody></table>
<div class="total" style="display:grid;grid-template-columns:repeat(4,1fr);gap:16px">
  <div><div style="font-size:10px;color:#64748b;text-transform:uppercase">Contract</div><div style="font-size:20px;font-weight:900;color:#60a5fa">£${totalContract.toLocaleString(undefined,{maximumFractionDigits:0})}</div></div>
  <div><div style="font-size:10px;color:#64748b;text-transform:uppercase">Claimed To Date</div><div style="font-size:20px;font-weight:900;color:#34d399">£${totalClaimed.toLocaleString(undefined,{maximumFractionDigits:0})}</div></div>
  <div><div style="font-size:10px;color:#64748b;text-transform:uppercase">This Period</div><div style="font-size:20px;font-weight:900;color:#fbbf24">£${totalThis.toLocaleString(undefined,{maximumFractionDigits:0})}</div></div>
  <div><div style="font-size:10px;color:#64748b;text-transform:uppercase">% Complete</div><div style="font-size:20px;font-weight:900;color:#a78bfa">${pctComplete}%</div></div>
</div>
<div class="ft"><span>${pa.number} — ${site?.name||""}</span><span>${client?.name||""}</span><span>Bright Metalwork Ltd</span><span>${new Date().toLocaleDateString("en-GB")}</span></div>
<script>window.onload=function(){window.print();}<\/script></body></html>`;
    const blob=new Blob([html],{type:"text/html"});const url=URL.createObjectURL(blob);
    const win=window.open(url,"_blank","width=1200,height=820");
    if(!win){const a=document.createElement("a");a.href=url;a.download=`${pa.number}.html`;a.click();}
    setTimeout(()=>URL.revokeObjectURL(url),6000);
  }

  const typeColors={"Scope of Work":"#60a5fa","Variation":"#fbbf24","Dayworks":"#f97316","Preliminaries":"#a78bfa","Other":"#94a3b8"};

  return <div>
    <DPageHdr title={`📐 ${pa.number}`} sub={`${site?.name||"—"} · ${client?.name||"—"} · ${pa.date}`}
      back="Payment Applications" onBack={()=>setPage("payapps")}
      actions={<div style={{display:"flex",gap:7}}>
        <button onClick={savePA} style={{padding:"6px 13px",background:"#1e2535",border:"1px solid #2d3555",borderRadius:7,color:"#94a3b8",cursor:"pointer",fontSize:12}}>💾 Save</button>
        <button onClick={exportPDF} style={{padding:"6px 13px",background:"#1a1f2e",border:"1px solid #ef4444",borderRadius:7,color:"#f87171",cursor:"pointer",fontSize:12,fontWeight:700}}>📄 Export PDF</button>
        <button onClick={submitPA} style={{padding:"6px 13px",background:"linear-gradient(135deg,#059669,#10b981)",border:"none",borderRadius:7,color:"#fff",cursor:"pointer",fontSize:12,fontWeight:700}}>✓ Submit</button>
      </div>}/>
    <div style={DS.body}>
      {/* Summary cards */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10,marginBottom:20}}>
        <DStat label="Contract Value" value={"£"+totalContract.toLocaleString(undefined,{maximumFractionDigits:0})} color="#60a5fa"/>
        <DStat label="Claimed To Date" value={"£"+totalClaimed.toLocaleString(undefined,{maximumFractionDigits:0})} color="#34d399"/>
        <DStat label="This Period" value={"£"+totalThis.toLocaleString(undefined,{maximumFractionDigits:0})} color="#fbbf24"/>
        <DStat label="% Complete" value={pctComplete+"%"} color="#a78bfa"/>
      </div>

      {/* Add item */}
      <div style={{display:"flex",gap:8,marginBottom:16,alignItems:"flex-end"}}>
        <div><label style={{...DS.th,display:"block",marginBottom:3}}>Item Type</label>
          <select value={addType} onChange={e=>setAddType(e.target.value)} style={{background:"#0f1421",border:"1px solid #2d3555",borderRadius:6,padding:"6px 10px",color:"#e2e8f0",fontSize:12,outline:"none",cursor:"pointer"}}>
            {PA_ITEM_TYPES.map(t=><option key={t} value={t}>{t}</option>)}
          </select>
        </div>
        <button onClick={addItem} style={{padding:"6px 14px",background:"linear-gradient(135deg,#3b82f6,#6366f1)",border:"none",borderRadius:7,color:"#fff",cursor:"pointer",fontSize:12,fontWeight:700}}>+ Add {addType} Item</button>
      </div>

      {/* Items table */}
      <div style={{border:"1px solid #1e2535",borderRadius:10,overflow:"hidden",marginBottom:16}}>
        <table style={{width:"100%",borderCollapse:"collapse"}}>
          <thead><tr>
            <th style={{...DS.th,minWidth:100}}>Type</th>
            <th style={{...DS.th,minWidth:220}}>Description</th>
            <th style={{...DS.th,minWidth:60}}>Unit</th>
            <th style={{...DS.th,minWidth:70}}>Qty</th>
            <th style={{...DS.th,minWidth:80}}>Rate £</th>
            <th style={{...DS.th,minWidth:90}}>Contract £</th>
            <th style={{...DS.th,background:"#0d1a2e",color:"#60a5fa",minWidth:80}}>Prev Qty</th>
            <th style={{...DS.th,background:"#0d1a2e",color:"#60a5fa",minWidth:80}}>Mode</th>
            <th style={{...DS.th,background:"#0d1a2e",color:"#60a5fa",minWidth:90}}>Claimed ToDate</th>
            <th style={{...DS.th,background:"#0d2218",color:"#34d399",minWidth:90}}>This Period £</th>
            <th style={{...DS.th,minWidth:40}}></th>
          </tr></thead>
          <tbody>{items.map((it,i)=>{
            const cl=calcClaimed(it);const pr=calcPrev(it);const thisPeriod=cl-pr;
            const typeCol=typeColors[it.type]||"#94a3b8";
            return <tr key={it.id} style={{background:i%2===0?"#111827":"#0f1421"}}>
              <td style={DS.td}><span style={{fontSize:10,fontWeight:600,color:typeCol}}>{it.type}</span></td>
              <td style={DS.td}><input value={it.description} onChange={e=>updateItem(it.id,"description",e.target.value)} style={{width:"100%",background:"#0f1421",border:"1px solid #2d3555",borderRadius:4,padding:"3px 6px",color:"#e2e8f0",fontSize:11,outline:"none"}}/></td>
              <td style={DS.td}><select value={it.unit} onChange={e=>updateItem(it.id,"unit",e.target.value)} style={{background:"#0f1421",border:"1px solid #2d3555",borderRadius:4,padding:"2px 4px",color:"#94a3b8",fontSize:10,outline:"none",cursor:"pointer"}}>
                {PA_UNITS.map(u=><option key={u} value={u}>{u}</option>)}
              </select></td>
              <td style={DS.td}><input type="number" value={it.contractQty} onChange={e=>updateItem(it.id,"contractQty",Number(e.target.value))} style={{width:60,background:"#0f1421",border:"1px solid #2d3555",borderRadius:4,padding:"2px 4px",color:"#60a5fa",fontSize:11,textAlign:"right",outline:"none"}}/></td>
              <td style={DS.td}><input type="number" value={it.contractRate} onChange={e=>updateItem(it.id,"contractRate",Number(e.target.value))} style={{width:70,background:"#0f1421",border:"1px solid #2d3555",borderRadius:4,padding:"2px 4px",color:"#e2e8f0",fontSize:11,textAlign:"right",outline:"none"}}/></td>
              <td style={{...DS.td,fontWeight:700,color:"#e2e8f0",textAlign:"right"}}>£{(it.contractQty*it.contractRate).toLocaleString(undefined,{maximumFractionDigits:0})}</td>
              <td style={{...DS.td,background:"#080d14"}}><input type="number" value={it.previousQty||0} onChange={e=>updateItem(it.id,"previousQty",Number(e.target.value))} style={{width:65,background:"#0f1421",border:"1px solid #2d3555",borderRadius:4,padding:"2px 4px",color:"#64748b",fontSize:11,textAlign:"right",outline:"none"}}/></td>
              <td style={{...DS.td,background:"#080d14",textAlign:"center"}}>
                <button onClick={()=>toggleMode(it.id)} style={{padding:"2px 8px",background:it.useQty?"#0d1a2e":"#1a0d2e",border:`1px solid ${it.useQty?"#3b82f6":"#a855f7"}`,borderRadius:4,color:it.useQty?"#60a5fa":"#c084fc",cursor:"pointer",fontSize:10,fontWeight:700}}>
                  {it.useQty?"Qty":"Pct%"}
                </button>
              </td>
              <td style={{...DS.td,background:"#080d14"}}>
                {it.useQty
                  ?<input type="number" value={it.claimedQtyToDate||0} onChange={e=>updateItem(it.id,"claimedQtyToDate",Number(e.target.value))} style={{width:70,background:"#0f1421",border:"1px solid #2d3555",borderRadius:4,padding:"2px 4px",color:"#34d399",fontSize:11,textAlign:"right",outline:"none"}} placeholder={it.unit}/>
                  :<div style={{display:"flex",alignItems:"center",gap:3}}><input type="number" min="0" max="100" value={it.claimedPctToDate||0} onChange={e=>updateItem(it.id,"claimedPctToDate",Number(e.target.value))} style={{width:50,background:"#0f1421",border:"1px solid #2d3555",borderRadius:4,padding:"2px 4px",color:"#34d399",fontSize:11,textAlign:"right",outline:"none"}}/><span style={{color:"#64748b",fontSize:11}}>%</span></div>
                }
              </td>
              <td style={{...DS.td,fontWeight:800,fontSize:13,color:thisPeriod>=0?"#34d399":"#f87171",background:"#080d14",textAlign:"right"}}>
                £{thisPeriod.toLocaleString(undefined,{maximumFractionDigits:0})}
                <div style={{fontSize:9,color:"#64748b"}}>{Math.round((cl/(it.contractQty*it.contractRate||1))*100)}%</div>
              </td>
              <td style={DS.td}><button onClick={()=>deleteItem(it.id)} style={{padding:"3px 6px",background:"#2d1515",border:"1px solid #ef4444",borderRadius:4,color:"#f87171",cursor:"pointer",fontSize:10}}>✕</button></td>
            </tr>;
          })}</tbody>
          <tfoot><tr style={{background:"#0d1117",borderTop:"2px solid #2d3555"}}>
            <td colSpan={5} style={{...DS.td,fontWeight:700,color:"#94a3b8"}}>TOTALS</td>
            <td style={{...DS.td,fontWeight:800,color:"#e2e8f0",textAlign:"right"}}>£{totalContract.toLocaleString(undefined,{maximumFractionDigits:0})}</td>
            <td colSpan={2} style={DS.td}/>
            <td style={{...DS.td,background:"#080d14",fontWeight:800,color:"#34d399",textAlign:"right"}}>£{totalClaimed.toLocaleString(undefined,{maximumFractionDigits:0})}</td>
            <td style={{...DS.td,background:"#080d14",fontWeight:900,fontSize:15,color:"#34d399",textAlign:"right"}}>£{totalThis.toLocaleString(undefined,{maximumFractionDigits:0})}</td>
            <td style={DS.td}/>
          </tr></tfoot>
        </table>
      </div>
    </div>
  </div>;
}


function DComingSoon({icon,title,sub}){
  return <div>
    <DPageHdr title={`${icon} ${title}`} sub={sub}/>
    <div style={{...DS.body,textAlign:"center",padding:60}}>
      <div style={{fontSize:48,marginBottom:16}}>{icon}</div>
      <div style={{fontSize:16,fontWeight:700,color:"#f1f5f9",marginBottom:8}}>{title}</div>
      <div style={{fontSize:13,color:"#64748b",marginBottom:20}}>{sub}</div>
      <div style={{fontSize:12,color:"#374151"}}>This section uses the same data as the Schedule view.<br/>Switch to 📋 Schedule to manage {title.toLowerCase()}.</div>
    </div>
  </div>;
}

// ─── DashboardView — the complete dashboard shell ─────────────────────────────

// ── Embedded Schedule View (full schedule table, inside dashboard) ─────────────
function DScheduleView({workers=[],allSites=[],activeDays=[],siteHours={},weekLabel,allSiteNames,filter={},setFilter,showWeekend,setShowWeekend,updateCell,setModal,delWorker,scheduleHistory,saveScheduleSnapshot,weeklyRecords,setWeeklyRecords}){
  // Build siteNames from allSiteNames prop or compute inline
  const siteNames=useMemo(()=>{
    if(allSiteNames&&allSiteNames.length>0) return allSiteNames;
    const s=new Set((allSites||[]).map(x=>x.name));
    (workers||[]).forEach(w=>ALL_DAYS.forEach(d=>{if(w.days?.[d])s.add(w.days[d].trim());}));
    return Array.from(s).filter(Boolean).sort();
  },[allSiteNames,allSites,workers]);

  // Auto-snapshot this week
  useEffect(()=>{
    if(workers.length>0&&scheduleHistory&&!scheduleHistory[weekLabel]&&saveScheduleSnapshot){
      saveScheduleSnapshot(weekLabel,workers);
    }
  },[weekLabel]);

  const nm=filter?.name||"";const pos=filter?.position||"";const si=filter?.site||"";
  const displayed=useMemo(()=>workers.filter(w=>{
    if(nm&&!w.name?.toLowerCase().includes(nm.toLowerCase())) return false;
    if(pos&&w.position!==pos) return false;
    if(si&&!Object.values(w.days||{}).some(d=>d&&d.toLowerCase().includes(si.toLowerCase()))) return false;
    return true;
  }),[workers,nm,pos,si]);

  // Group by primary site
  const groups=useMemo(()=>{
    const g={};
    displayed.forEach(w=>{
      const cnts={};
      (activeDays||BASE_DAYS).forEach(d=>{const s=w.days?.[d];if(s&&!isOff(s))cnts[s]=(cnts[s]||0)+1;});
      const primary=Object.entries(cnts).sort((a,b)=>b[1]-a[1])[0]?.[0]||"Unassigned / Off";
      if(!g[primary])g[primary]=[];
      g[primary].push(w);
    });
    return g;
  },[displayed,activeDays]);

  const siteOrder=Object.keys(groups).sort((a,b)=>{
    if(a==="Unassigned / Off") return 1;
    if(b==="Unassigned / Off") return -1;
    return a.localeCompare(b);
  });

  const TH2={padding:"7px 10px",textAlign:"left",fontSize:10,fontWeight:700,color:"#64748b",textTransform:"uppercase",borderBottom:"1px solid #1e2535",background:"#0a0e17",whiteSpace:"nowrap"};
  const TD2={padding:"5px 8px",borderBottom:"1px solid #1a2030",verticalAlign:"middle"};

  return <div>
    <DPageHdr title="📋 Labour Schedule" sub={`WC: ${weekLabel} · ${displayed.length} operatives · grouped by site`}
      actions={<div style={{display:"flex",gap:6,flexWrap:"wrap",alignItems:"center"}}>
        <input value={nm} onChange={e=>setFilter&&setFilter(f=>({...f,name:e.target.value}))} placeholder="🔍 Name…"
          style={{background:"#1a1f2e",border:"1px solid #2d3555",borderRadius:6,padding:"5px 9px",color:"#e2e8f0",fontSize:11,outline:"none",width:100}}/>
        <button onClick={()=>setShowWeekend&&setShowWeekend(s=>!s)}
          style={{padding:"5px 10px",background:showWeekend?"#1a3020":"#1a1f2e",border:`1px solid ${showWeekend?"#10b981":"#2d3555"}`,borderRadius:6,color:showWeekend?"#34d399":"#64748b",cursor:"pointer",fontSize:10,fontWeight:700}}>
          {showWeekend?"✓ Weekend":"+ Weekend"}
        </button>
        <button onClick={()=>exportSchedulePDF(displayed,activeDays||BASE_DAYS,weekLabel,allSites)}
          style={{padding:"5px 10px",background:"#1a1f2e",border:"1px solid #ef4444",borderRadius:6,color:"#f87171",cursor:"pointer",fontSize:10,fontWeight:700}}>📄 PDF</button>
        <button onClick={()=>doExcel(workers,weekLabel,activeDays||BASE_DAYS,siteHours||{},{},allSites)}
          style={{padding:"5px 10px",background:"linear-gradient(135deg,#059669,#10b981)",border:"none",borderRadius:6,color:"#fff",cursor:"pointer",fontSize:10,fontWeight:700}}>⬇ Excel</button>
        <button onClick={()=>setModal&&setModal({type:"worker",worker:mkW()})}
          style={{padding:"5px 10px",background:"linear-gradient(135deg,#3b82f6,#6366f1)",border:"none",borderRadius:6,color:"#fff",cursor:"pointer",fontSize:10,fontWeight:700}}>+ Worker</button>
      </div>}/>

    <div style={{padding:"4px 20px",background:"#0f1421",borderBottom:"1px solid #1e2535",fontSize:11,color:"#64748b"}}>
      💡 <strong style={{color:"#60a5fa"}}>Click any site cell</strong> to edit inline · <strong style={{color:"#a78bfa"}}>📋</strong> profile PDF · <strong style={{color:"#34d399"}}>💷</strong> payslip PDF · <strong style={{color:"#60a5fa"}}>🔗</strong> open in new window
    </div>

    <div style={{overflowX:"auto"}}>
      {siteOrder.length===0&&<div style={{textAlign:"center",padding:50,color:"#374151"}}>No workers found. Add workers using the + Worker button.</div>}
      {siteOrder.map(siteName=>{
        const siteColor=getSiteColor(siteName,allSites);
        const grpWorkers=groups[siteName]||[];
        return <div key={siteName}>
          <div style={{background:`${siteColor}15`,borderLeft:`4px solid ${siteColor}`,padding:"6px 18px",display:"flex",alignItems:"center",gap:10,borderTop:"1px solid #1e2535",borderBottom:`1px solid ${siteColor}33`}}>
            <span style={{width:9,height:9,borderRadius:"50%",background:siteColor,flexShrink:0}}/>
            <span style={{fontWeight:800,color:siteColor,fontSize:13,flex:1}}>{siteName}</span>
            <span style={{fontSize:11,color:"#64748b"}}>{grpWorkers.length} operative{grpWorkers.length!==1?"s":""}</span>
          </div>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
            <thead><tr>
              <th style={{...TH2,minWidth:140,paddingLeft:20}}>Worker</th>
              <th style={TH2}>Co.</th>
              <th style={TH2}>Position</th>
              {(activeDays||BASE_DAYS).map(d=><th key={d} style={{...TH2,minWidth:130,color:WEEKEND_DAYS.includes(d)?"#fbbf24":"#64748b"}}>{d}{WEEKEND_DAYS.includes(d)?" 🟡":""}</th>)}
              <th style={TH2}>Rate</th>
              <th style={TH2}>Tax</th>
              <th style={TH2}>Certs</th>
              <th style={TH2}>Actions</th>
            </tr></thead>
            <tbody>
              {grpWorkers.map((w,i)=>{
                const exp=CERTS.filter(c=>cSt(c,w)==="expired").length;
                const expg=CERTS.filter(c=>cSt(c,w)==="expiring").length;
                return <tr key={w.id} style={{background:i%2===0?"#111827":"#0f1421"}}>
                  <td style={{...TD2,fontWeight:600,color:"#f1f5f9",paddingLeft:20}}>
                    {w.name}
                    {w.comments&&<div style={{fontSize:9,color:"#fbbf24"}}>⚑ {w.comments}</div>}
                  </td>
                  <td style={{...TD2,color:"#94a3b8",fontSize:10}}>{(w.company||"").split(" ")[0]||"—"}</td>
                  <td style={{...TD2,color:"#94a3b8",fontSize:10}}>{w.position||"—"}</td>
                  {(activeDays||BASE_DAYS).map(d=>(
                    <td key={d} style={{...TD2,background:WEEKEND_DAYS.includes(d)?"rgba(251,191,36,0.03)":undefined,padding:"3px 6px"}}>
                      <InlineCell value={w.days?.[d]||""} workerId={w.id} day={d} allSiteNames={siteNames} allSites={allSites} onUpdate={updateCell||((id,day,val)=>{})}/>
                    </td>
                  ))}
                  <td style={{...TD2,color:"#34d399",fontWeight:600,fontSize:11}}>{w.agreedRate?`£${w.agreedRate}/hr`:"—"}</td>
                  <td style={TD2}><span style={{fontSize:10,fontWeight:700,color:w.taxRate===0.30?"#f87171":w.taxRate===0.20?"#fbbf24":"#34d399"}}>{Math.round((w.taxRate||0)*100)}%</span></td>
                  <td style={TD2}><div style={{display:"flex",gap:3}}>
                    {exp>0&&<span style={{color:"#f87171",fontSize:10,fontWeight:700}}>✗{exp}</span>}
                    {expg>0&&<span style={{color:"#fbbf24",fontSize:10,fontWeight:700}}>⚠{expg}</span>}
                    {exp===0&&expg===0&&<span style={{color:"#374151",fontSize:10}}>—</span>}
                  </div></td>
                  <td style={TD2}><div style={{display:"flex",gap:3,flexWrap:"nowrap"}}>
                    <button onClick={()=>setModal&&setModal({type:"worker",worker:w})} style={{padding:"3px 7px",background:"#1e3a5f",border:"1px solid #3b82f6",borderRadius:4,color:"#60a5fa",cursor:"pointer",fontSize:10,fontWeight:600}}>Edit</button>
                    <button onClick={()=>openWorkerWindow(w,allSites,weekLabel,activeDays||BASE_DAYS,siteHours||{})} style={{padding:"3px 6px",background:"#1a1f2e",border:"1px solid #60a5fa",borderRadius:4,color:"#60a5fa",cursor:"pointer",fontSize:10}}>🔗</button>
                    <button onClick={()=>exportWorkerProfile(w,allSites,weekLabel)} style={{padding:"3px 6px",background:"#1a2535",border:"1px solid #8b5cf6",borderRadius:4,color:"#a78bfa",cursor:"pointer",fontSize:10}}>📋</button>
                    <button onClick={()=>exportPayslip(w,activeDays||BASE_DAYS,weekLabel,siteHours||{})} style={{padding:"3px 6px",background:"#0d2218",border:"1px solid #10b981",borderRadius:4,color:"#34d399",cursor:"pointer",fontSize:10}}>💷</button>
                    <button onClick={()=>delWorker&&delWorker(w.id)} style={{padding:"3px 6px",background:"#2d1515",border:"1px solid #ef4444",borderRadius:4,color:"#f87171",cursor:"pointer",fontSize:10}}>✕</button>
                  </div></td>
                </tr>;
              })}
            </tbody>
          </table>
        </div>;
      })}
    </div>

    {/* Site legend */}
    <div style={{padding:"8px 20px",borderTop:"1px solid #1e2535",background:"#0a0e17",display:"flex",flexWrap:"wrap",gap:4}}>
      {allSites.filter(s=>!isOff(s.name)).map(s=>(
        <span key={s.id} style={{display:"inline-flex",alignItems:"center",gap:3,padding:"2px 7px",borderRadius:20,background:"#111827",border:`1px solid ${s.color}`,fontSize:9,color:"#94a3b8"}}>
          <span style={{width:5,height:5,borderRadius:"50%",background:s.color}}/>{s.name}
        </span>
      ))}
    </div>
  </div>;
}

function DSiteBySite({workers,allSites,activeDays}){
  const sm={};
  workers.forEach(w=>activeDays.forEach(d=>{const s=(w.days[d]||"").trim();if(s){if(!sm[s])sm[s]={};if(!sm[s][d])sm[s][d]=[];sm[s][d].push(w);}}));
  return <div>
    <DPageHdr title="📍 By Site" sub="All workers grouped by site per day"/>
    <div style={{padding:"0 24px 24px",overflowX:"auto"}}>
      <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,marginTop:16}}>
        <thead><tr>
          <th style={{...DS.th,minWidth:180}}>Site</th>
          {activeDays.map(d=><th key={d} style={{...DS.th,minWidth:120,color:WEEKEND_DAYS.includes(d)?"#fbbf24":"#64748b"}}>{d}{WEEKEND_DAYS.includes(d)?" 🟡":""}</th>)}
          <th style={DS.th}>Total</th>
        </tr></thead>
        <tbody>{Object.keys(sm).sort().map((site,i)=>{
          const color=getSiteColor(site,allSites);
          const all=new Set();activeDays.forEach(d=>(sm[site][d]||[]).forEach(w=>all.add(w.id)));
          return <tr key={site} style={{background:i%2===0?"#111827":"#0f1421"}}>
            <td style={{...DS.td,borderLeft:`3px solid ${color}`,paddingLeft:12}}><span style={{fontWeight:700,color}}>{site}</span></td>
            {activeDays.map(d=><td key={d} style={DS.td}>{(sm[site][d]||[]).map(w=><div key={w.id} style={{fontSize:11,color:"#cbd5e1"}}>{w.name} <span style={{color:"#64748b",fontSize:10}}>({w.position||"—"})</span></div>)}</td>)}
            <td style={{...DS.td,textAlign:"center",fontWeight:700,color:"#60a5fa"}}>{all.size}</td>
          </tr>;
        })}</tbody>
      </table>
    </div>
  </div>;
}


function DashboardView({workers,allSites,clients,weekLabel,activeDays,siteHours,scopeData,invoices,saveWorker,delWorker,setAllSites,setClients,setModal,weeklyRecords,setWeeklyRecords,scheduleHistory,saveScheduleSnapshot,bankTransactions,setBankTransactions,timesheetRecords,setTimesheetRecords,payslipRecords,setPayslipRecords,payApplications,setPayApplications,generateTimesheets,generatePayslips,showWeekend,filter,setFilter,allSiteNames,updateCell,dashPage,setDashPage,dashDetailId,setDashDetailId}){
  const nav=(page,id)=>{setDashPage(page);if(id!==undefined)setDashDetailId(id);};
  const goBack=(page)=>setDashPage(page);

    // Shared props passed to all dashboard pages
  const SP={
    workers,allSites,clients,invoices,activeDays,siteHours,weekLabel,setModal,scopeData,
    filter,setFilter,allSiteNames,updateCell,delWorker,showWeekend,
    weeklyRecords,setWeeklyRecords,scheduleHistory,saveScheduleSnapshot,
    bankTransactions,setBankTransactions,
    timesheetRecords,setTimesheetRecords,
    payslipRecords,setPayslipRecords,
    payApplications,setPayApplications,
    generateTimesheets,generatePayslips,
  };
  const renderPage=()=>{
    switch(dashPage){
      // ── Overview
      case "home":          return <DHome {...SP} weeklyRecords={weeklyRecords} setPage={setDashPage}/>;
      // ── Labour & People
      case "workers":       return <DWorkers {...SP} setPage={nav} setDetailId={setDashDetailId}/>;
      case "worker_detail": return <DWorkerDetail {...SP} workerId={dashDetailId} setPage={setDashPage}/>;
      // ── Labour Schedule (auto-snapshots each week)
      case "schedule":      return <DScheduleView {...SP} setShowWeekend={setShowWeekend}/>;
      case "site_by_site":  return <DSiteBySite {...SP}/>;
      case "payroll":       return <DPayroll {...SP}/>;
      case "payslips":      return <DPayslips {...SP} setPage={setDashPage}/>;
      case "timesheets":    return <DTimesheets {...SP} setPage={setDashPage}/>;
      case "weekly_records":return <DWeeklyRecords weeklyRecords={weeklyRecords} setWeeklyRecords={setWeeklyRecords} workers={workers} allSites={allSites} clients={clients} siteHours={siteHours} activeDays={activeDays} weekLabel={weekLabel} showWeekend={showWeekend} invoices={invoices} setPage={setDashPage} setDetailId={setDashDetailId}/>;
      case "weekly_record_detail": return <DWeeklyRecordDetail weeklyRecords={weeklyRecords} recordId={dashDetailId} setPage={setDashPage} allSites={allSites}/>;
      // ── Projects & Finance
      case "sites":         return <DSites {...SP} setPage={nav} setDetailId={setDashDetailId}/>;
      case "site_detail":   return <DSiteDetail {...SP} siteId={dashDetailId} setPage={setDashPage} setDetailId={setDashDetailId}/>;
      case "clients":       return <DClients {...SP} setPage={nav} setDetailId={setDashDetailId}/>;
      case "client_detail": return <DComingSoon icon="👔" title="Client Detail" sub="Select a client from the Clients list"/>;
      case "invoices":      return <DInvoices {...SP}/>;
      case "payapps":       return <DPayApps {...SP} setPage={setDashPage} setDetailId={setDashDetailId}/>;
      case "payapp_detail": return <DPayAppDetail payApplications={payApplications} setPayApplications={setPayApplications} payappId={dashDetailId} allSites={allSites} clients={clients} setPage={setDashPage}/>;
      case "budget":        return <DBudget {...SP}/>;
      // ── Analysis
      case "certs":         return <DCerts workers={workers} setPage={nav} setDetailId={setDashDetailId}/>;
      case "finance":       return <DFinance {...SP}/>;
      case "stats":         return <DStats workers={workers} allSites={allSites} activeDays={activeDays}/>;
      case "bank":          return <DBankFull {...SP}/>;
      default:              return <DHome {...SP} weeklyRecords={weeklyRecords} setPage={setDashPage}/>;
    }
  };


  // DashboardView only renders the page content — sidebar is in App
  return <div style={{minHeight:"100%",background:"#080d14"}}>{renderPage()}</div>;
}


// ─── Main App ──────────────────────────────────────────────────────────────────
export default function App(){
  const [workers,setWorkers]=useState([]);
  const [weekLabel,setWeekLabel]=useState(formatWeekLabel(new Date()));
  const [showWeekend,setShowWeekend]=useState(false);
  const [allSites,setAllSites]=useState(DEFAULT_BUILTIN_SITES);
  const [clients,setClients]=useState(INIT_CLIENTS);
  const [siteHours,setSiteHours]=useState({});
  const [filter,setFilter]=useState({name:"",position:"",site:""});
  const [view,setView]=useState("schedule");
  const [modal,setModal]=useState(null);
  const [loading,setLoading]=useState(true);
  const [syncStatus,setSyncStatus]=useState("saved");
  const saveTimer=useRef(null);
  const [scopeData,setScopeData]=useState({});
  const [invoices,setInvoices]=useState([]);
  const [weeklyRecords,setWeeklyRecords]=useState([]);    // closed week snapshots
  const [scheduleHistory,setScheduleHistory]=useState({}); // {weekLabel: {workers,sites,...}}
  const [bankTransactions,setBankTransactions]=useState([]); // saved bank statement lines
  const [timesheetRecords,setTimesheetRecords]=useState([]); // individual timesheet entries
  const [payslipRecords,setPayslipRecords]=useState([]);     // individual payslip entries
  const [payApplications,setPayApplications]=useState([]);   // payment applications per site
  const [dashPage,setDashPage]=useState("home");
  const [dashDetailId,setDashDetailId]=useState(null);

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
        if(cfg.scope_data) setScopeData(cfg.scope_data);
        if(cfg.invoices) setInvoices(cfg.invoices);
        if(cfg.weekly_records) setWeeklyRecords(cfg.weekly_records);
        if(cfg.schedule_history) setScheduleHistory(cfg.schedule_history);
        if(cfg.bank_transactions) setBankTransactions(cfg.bank_transactions);
        if(cfg.timesheet_records) setTimesheetRecords(cfg.timesheet_records);
        if(cfg.payslip_records) setPayslipRecords(cfg.payslip_records);
        if(cfg.pay_applications) setPayApplications(cfg.pay_applications);
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
          {key:"week_label",value:weekLabel},{key:"show_weekend",value:showWeekend},
          {key:"all_sites",value:allSites},{key:"clients",value:clients},
          {key:"site_hours",value:siteHours},{key:"scope_data",value:scopeData},
          {key:"invoices",value:invoices},
          {key:"weekly_records",value:weeklyRecords},
          {key:"schedule_history",value:scheduleHistory},
          {key:"bank_transactions",value:bankTransactions},
          {key:"timesheet_records",value:timesheetRecords},
          {key:"payslip_records",value:payslipRecords},
          {key:"pay_applications",value:payApplications},
        ]);
        setSyncStatus("saved");
      } catch(e){ setSyncStatus("error"); }
    },800);
  },[weekLabel,showWeekend,allSites,clients,siteHours,scopeData,invoices,loading]);

  const saveSiteDetail=(updatedSite)=>{
    setAllSites(sites=>sites.map(s=>s.id===updatedSite.id?updatedSite:s));
    setModal(null);
  };
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
  const saveScopeForSite=(siteId,items)=>{setScopeData(d=>({...d,[siteId]:items}));setModal(null);};
  // Auto-save schedule snapshot whenever weekLabel changes (creates one per week automatically)
  const saveScheduleSnapshot=(label,workerData)=>{
    const snap={
      weekLabel:label||weekLabel,
      savedAt:new Date().toISOString(),
      workers:JSON.parse(JSON.stringify(workerData||workers)),
      allSites:JSON.parse(JSON.stringify(allSites)),
      siteHours:JSON.parse(JSON.stringify(siteHours)),
      activeDays:showWeekend?[...BASE_DAYS,...WEEKEND_DAYS]:BASE_DAYS,
    };
    setScheduleHistory(h=>({...h,[label||weekLabel]:snap}));
  };
  // Generate timesheets for current week from worker data
  const generateTimesheets=(wkLabel)=>{
    const days=showWeekend?[...BASE_DAYS,...WEEKEND_DAYS]:BASE_DAYS;
    const sheets=workers.map(w=>{
      const {stdH,otH,gross,net,tax}=calcPay(w,days,siteHours);
      return {
        id:"ts_"+w.id+"_"+Date.now(),
        workerId:w.id, workerName:w.name, position:w.position,
        weekLabel:wkLabel||weekLabel, stdHours:stdH, otHours:otH,
        days:JSON.parse(JSON.stringify(w.days||{})),
        rate:w.agreedRate||0, gross, net, tax, taxRate:w.taxRate||0,
        approved:false, createdAt:new Date().toISOString(),
      };
    }).filter(t=>t.stdHours>0||t.otHours>0);
    return sheets;
  };
  // Generate payslips from timesheets
  const generatePayslips=(timesheets)=>{
    return timesheets.map(t=>({
      id:"ps_"+t.workerId+"_"+Date.now()+Math.random().toString(36).slice(2),
      workerId:t.workerId, workerName:t.workerName, position:t.position,
      weekLabel:t.weekLabel, stdHours:t.stdHours, otHours:t.otHours,
      rate:t.rate, gross:t.gross, net:t.net, tax:t.tax, taxRate:t.taxRate,
      issued:false, createdAt:new Date().toISOString(),
    }));
  };
  const saveInvoice=inv=>{setInvoices(list=>{const exists=list.find(x=>x.id===inv.id);return exists?list.map(x=>x.id===inv.id?inv:x):[...list,inv];});setModal(null);};
  const delInvoice=id=>{if(window.confirm("Delete invoice?"))setInvoices(list=>list.filter(x=>x.id!==id));};

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

  const VIEWS=[["schedule","📋 Schedule"],["site","📍 By Site"],["certs","🛡 Certs"],["payroll","💷 Payroll"],["costs","👔 Costs"],["bank","🏦 Bank"],["budget","📐 Budget"],["invoices","🧾 Invoices"],["finance","📊 Finance"],["stats","🔢 Stats"]];

  if(loading) return <div style={{minHeight:"100vh",background:"#0d1117",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",gap:16}}>
    <div style={{width:48,height:48,background:"linear-gradient(135deg,#3b82f6,#6366f1)",borderRadius:12,display:"flex",alignItems:"center",justifyContent:"center",fontSize:24}}>🏗</div>
    <div style={{color:"#60a5fa",fontSize:16,fontWeight:700}}>Loading…</div>
    <div style={{width:200,height:3,background:"#1e2535",borderRadius:3,overflow:"hidden"}}><div style={{height:"100%",background:"linear-gradient(90deg,#3b82f6,#6366f1)",borderRadius:3,animation:"slide 1.5s ease-in-out infinite"}}/></div>
    <style>{`@keyframes slide{0%{width:0%;margin-left:0%}50%{width:60%;margin-left:20%}100%{width:0%;margin-left:100%}}`}</style>
  </div>;

  return (
    <div style={{minHeight:"100vh",background:"#0d1117",fontFamily:"system-ui,'Segoe UI',sans-serif",color:"#e2e8f0",fontSize:13,display:"flex",flexDirection:"column"}}>

      {/* ── TOP BAR: always visible, always exactly one ── */}
      <div style={{background:"linear-gradient(135deg,#0f172a,#1a1f2e)",borderBottom:"1px solid #1e2535",padding:"9px 16px",position:"sticky",top:0,zIndex:300,flexShrink:0}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",gap:8}}>

          {/* Logo + week nav */}
          <div style={{display:"flex",alignItems:"center",gap:10}}>
            <div style={{width:30,height:30,background:"linear-gradient(135deg,#3b82f6,#6366f1)",borderRadius:7,display:"flex",alignItems:"center",justifyContent:"center",fontSize:14}}>🏗</div>
            <div>
              <div style={{fontSize:14,fontWeight:800,color:"#f1f5f9",lineHeight:1}}>Bright Metalwork</div>
              <div style={{display:"flex",alignItems:"center",gap:5,marginTop:3}}>
                <button onClick={()=>setWeekLabel(addWeeks(weekLabel,-1))} style={{background:"#1e2535",border:"1px solid #2d3555",borderRadius:4,color:"#94a3b8",cursor:"pointer",fontSize:12,padding:"0 6px",fontWeight:700,lineHeight:1.5}}>‹</button>
                <span style={{fontSize:10,color:"#64748b"}}>WC:</span>
                <input value={weekLabel} onChange={e=>setWeekLabel(e.target.value)} style={{background:"none",border:"none",borderBottom:"1px solid #2d3555",color:"#60a5fa",fontWeight:600,fontSize:11,outline:"none",width:105}}/>
                <button onClick={()=>setWeekLabel(addWeeks(weekLabel,1))} style={{background:"#1e2535",border:"1px solid #2d3555",borderRadius:4,color:"#94a3b8",cursor:"pointer",fontSize:12,padding:"0 6px",fontWeight:700,lineHeight:1.5}}>›</button>
                <button onClick={()=>setWeekLabel(formatWeekLabel(new Date()))} style={{background:"#1e2535",border:"1px solid #2d3555",borderRadius:4,color:"#64748b",cursor:"pointer",fontSize:9,padding:"1px 6px",fontWeight:700}}>Today</button>
              </div>
            </div>
          </div>

          {/* Live stats strip — clickable */}
          <div style={{display:"flex",gap:5,flexWrap:"wrap"}}>
            {[
              {l:"Workers",  v:stats.total,                              c:"#60a5fa",  pg:"workers"},
              {l:"Holiday",  v:stats.onHol,                             c:"#fbbf24",  pg:"workers"},
              {l:"Off",      v:stats.off,                                c:"#94a3b8",  pg:"workers"},
              {l:"Cert ⚠",  v:stats.alerts, c:stats.alerts>0?"#f87171":"#34d399",    pg:"certs"},
              {l:"Gross",    v:stats.g>0?`£${stats.g.toFixed(0)}`:"—",  c:"#34d399",  pg:"payslips"},
              {l:"Net",      v:stats.g>0?`£${stats.n.toFixed(0)}`:"—",  c:"#a78bfa",  pg:"payslips"},
              {l:"Clients",  v:clients.length,                           c:"#8b5cf6",  pg:"clients"},
              {l:"Invoices", v:invoices.length,                          c:"#10b981",  pg:"invoices"},
            ].map(s=>(
              <div key={s.l} onClick={()=>setDashPage(s.pg)}
                style={{background:"#111827",border:"1px solid #1e2535",borderRadius:7,padding:"3px 10px",cursor:"pointer",transition:"border-color 0.15s"}}
                onMouseEnter={e=>e.currentTarget.style.borderColor=s.c} onMouseLeave={e=>e.currentTarget.style.borderColor="#1e2535"}>
                <div style={{fontSize:9,color:"#64748b",fontWeight:700,textTransform:"uppercase"}}>{s.l}</div>
                <div style={{fontSize:13,fontWeight:800,color:s.c,lineHeight:1.2}}>{s.v}</div>
              </div>
            ))}
          </div>

          {/* Right side controls */}
          <div style={{display:"flex",gap:5,alignItems:"center",flexShrink:0}}>
            <span style={{display:"flex",alignItems:"center",gap:4,fontSize:11,color:syncStatus==="saved"?"#34d399":syncStatus==="error"?"#f87171":"#fbbf24"}}>
              <span style={{width:6,height:6,borderRadius:"50%",background:syncStatus==="saved"?"#34d399":syncStatus==="error"?"#f87171":"#fbbf24",display:"inline-block"}}/>
              {syncStatus==="saved"?"☁ Saved":syncStatus==="error"?"⚠ Error":"Saving…"}
            </span>
            <button onClick={()=>setShowWeekend(s=>!s)} style={{padding:"4px 9px",background:showWeekend?"#1a3020":"#1a1f2e",border:`1px solid ${showWeekend?"#10b981":"#2d3555"}`,borderRadius:6,color:showWeekend?"#34d399":"#64748b",cursor:"pointer",fontSize:10,fontWeight:700}}>{showWeekend?"✓ Wknd":"+ Wknd"}</button>
            <button onClick={()=>{
              if(!window.confirm(`💾 Save WC ${weekLabel} as a weekly record?`)) return;
              const snap={id:"wk_"+Date.now(),weekLabel,savedAt:new Date().toISOString(),
                workers:JSON.parse(JSON.stringify(workers)),
                allSites:JSON.parse(JSON.stringify(allSites)),
                siteHours:JSON.parse(JSON.stringify(siteHours)),
                activeDays:showWeekend?[...BASE_DAYS,...WEEKEND_DAYS]:BASE_DAYS,
                invoices:JSON.parse(JSON.stringify(invoices)),status:"closed"};
              setWeeklyRecords(recs=>{const ex=recs.find(r=>r.weekLabel===weekLabel);return ex?recs.map(r=>r.weekLabel===weekLabel?snap:r):[...recs,snap];});
              alert(`✓ WC ${weekLabel} saved to Weekly Records.`);
            }} style={{padding:"4px 9px",background:"#0d2218",border:"1px solid #10b981",borderRadius:6,color:"#34d399",cursor:"pointer",fontSize:10,fontWeight:700}}>💾 Close Week</button>
            <button onClick={()=>setModal({type:"worker",worker:mkW()})} style={{padding:"4px 10px",background:"linear-gradient(135deg,#3b82f6,#6366f1)",border:"none",borderRadius:6,color:"#fff",cursor:"pointer",fontSize:10,fontWeight:700}}>+ Worker</button>
          </div>
        </div>
      </div>

      {/* ── BODY: ONE sidebar + content — no nesting, no duplication ── */}
      <div style={{display:"flex",flex:1,minHeight:0}}>

        {/* THE ONLY SIDEBAR IN THE ENTIRE APP */}
        <DashSidebar
          page={dashPage}
          setPage={p=>{setDashPage(p);setDashDetailId(null);}}
          workers={workers} allSites={allSites} clients={clients}
          invoices={invoices} setModal={setModal}
          activeDays={activeDays} siteHours={siteHours} weekLabel={weekLabel}/>

        {/* Main content — DashboardView renders ONLY page content, never a sidebar */}
        <div style={{flex:1,overflowY:"auto",height:"calc(100vh - 56px)"}}>
          <DashboardView
            workers={workers} allSites={allSites} clients={clients}
            weekLabel={weekLabel} activeDays={activeDays} siteHours={siteHours}
            scopeData={scopeData} invoices={invoices}
            saveWorker={saveWorker} delWorker={delWorker}
            setAllSites={setAllSites} setClients={setClients} setModal={setModal}
            weeklyRecords={weeklyRecords} setWeeklyRecords={setWeeklyRecords}
            scheduleHistory={scheduleHistory} saveScheduleSnapshot={saveScheduleSnapshot}
            bankTransactions={bankTransactions} setBankTransactions={setBankTransactions}
            timesheetRecords={timesheetRecords} setTimesheetRecords={setTimesheetRecords}
            payslipRecords={payslipRecords} setPayslipRecords={setPayslipRecords}
            payApplications={payApplications} setPayApplications={setPayApplications}
            generateTimesheets={generateTimesheets} generatePayslips={generatePayslips}
            showWeekend={showWeekend}
            filter={filter} setFilter={setFilter}
            allSiteNames={allSiteNames} updateCell={updateCell}
            dashPage={dashPage} setDashPage={setDashPage}
            dashDetailId={dashDetailId} setDashDetailId={setDashDetailId}/>
        </div>
      </div>

      {/* ── MODALS — always available regardless of current page ── */}
      {modal?.type==="worker"&&<WorkerModal worker={modal.worker} onSave={saveWorker} onClose={()=>setModal(null)} allSiteNames={allSiteNames} allSites={allSites} activeDays={activeDays}/>}
      {modal?.type==="sites"&&<SitesModal allSites={allSites} clients={clients} onSave={s=>{setAllSites(s);}} onClose={()=>setModal(null)} onOpenDetail={site=>setModal({type:"siteDetail",site})}/>}
      {modal?.type==="clients"&&<ClientsModal clients={clients} onSave={l=>{setClients(l);setModal(null);}} onClose={()=>setModal(null)}/>}
      {modal?.type==="siteDetail"&&<SiteDetailModal site={modal.site} clients={clients} workers={workers} activeDays={activeDays} siteHours={siteHours} onSave={saveSiteDetail} onClose={()=>setModal(null)}/>}
      {modal?.type==="bank"&&<BankImportModal allSites={allSites} clients={clients} onClose={()=>setModal(null)}/>}
      {modal?.type==="scope"&&<ScopeModal site={modal.site} scopeItems={scopeData[modal.site?.id]||[]} onSave={items=>saveScopeForSite(modal.site.id,items)} onClose={()=>setModal(null)}/>}
      {modal?.type==="invoice"&&<InvoiceModal invoice={modal.invoice} clients={clients} allSites={allSites} scopeData={scopeData} workers={workers} invoices={invoices} onSave={saveInvoice} onClose={()=>setModal(null)}/>}
      {modal?.type==="trainingMatrix"&&<TrainingMatrixModal workers={workers} clients={clients} allSites={allSites} activeDays={activeDays} weekLabel={weekLabel} onClose={()=>setModal(null)}/>}
    </div>
  );
}
{
  "sites": [
    {
      "id": "s1",
      "name": "Paddington Station",
      "address": "Praed St, London W2 1HQ",
      "scope": "External façade works – brick replacement",
      "supervisor": {
        "name": "Marcus Webb",
        "phone": "07700 900123"
      },
      "workers": [
        { "id": "w1", "name": "Tom Bradley", "role": "Scaffolder" },
        { "id": "w2", "name": "Sarah Chen",  "role": "Bricklayer" }
      ]
    }
  ]
}
