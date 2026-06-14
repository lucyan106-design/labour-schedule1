import { useState, useMemo, useRef, useEffect, useCallback } from "react";
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
const DEFAULT_POSITIONS=[
  "Welder","Steel Erector","Fixer","Fitter","Semiskilled","Supervisor",
  "Labourer","Manager","Driver","Slinger Signaller","Crane Driver",
  "Crane Supervisor","Magic Man","Fire Watcher","Cladding Operative",
  "Architectural Metalworker","Plant Operator","Site Engineer",
];
// Mutable array — custom trades get pushed here at runtime
const POSITIONS=DEFAULT_POSITIONS.slice();
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
    contact:"",email:"",dob:"",address:"",nationality:"",
    nino:"",utr:"",
    bankName:"",bankAccount:"",bankSort:"",nextOfKin:"",nextOfKinPhone:"",
    shareCode:"",shareCodeDate:"",shareCodeExpiry:"",
    workerFiles:[],
    holidayRequests:[],
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
function InlineCell({value,workerId,day,allSiteNames,allSites,onUpdate,confirmed=false}){
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
  return <div onClick={()=>setEditing(true)} title={confirmed?"✅ GPS confirmed — click to edit forecast":"📋 Forecast — click to edit"}
    style={{cursor:"text",minWidth:110,padding:"3px 4px",borderRadius:5,
      border:`1px solid ${confirmed?"#34d39933":"transparent"}`,
      background:confirmed?"#0d221811":"transparent",transition:"border-color 0.15s"}}
    onMouseEnter={e=>{if(!confirmed)e.currentTarget.style.borderColor="#2d3555";}}
    onMouseLeave={e=>{if(!confirmed)e.currentTarget.style.borderColor="transparent";}}>
    {value
      ?<div style={{display:"flex",alignItems:"center",gap:3}}>
          <Bdg label={value.trim()} color={getSiteColor(value,allSites)}/>
          {confirmed
            ?<span style={{fontSize:9,color:"#34d399",fontWeight:700}}>✓</span>
            :<span style={{fontSize:9,color:"#64748b"}}>📋</span>}
        </div>
      :<span style={{color:"#374151",fontSize:11}}>— click —</span>}
  </div>;
}

// ─── Worker Profile PDF ───────────────────────────────────────────────────────


// ─── Export Payslip PDF ───────────────────────────────────────────────────────
function exportPayslipPDF(ps, worker) {
  const w = worker;
  const or  = Math.round((w.agreedRate||0)*(w.overtimeMultiplier||1.5)*100)/100;
  const customOT = w.customOTRate ? Number(w.customOTRate) : or;
  const stdAmt  = Math.round((ps.stdH||0)*(w.agreedRate||0)*100)/100;
  const otAmt   = Math.round((ps.otH||0)*customOT*100)/100;
  const gross   = Math.round((stdAmt+otAmt)*100)/100;
  const taxRate = w.taxRate||0;
  const taxAmt  = Math.round(gross*taxRate*100)/100;
  const net     = Math.round((gross-taxAmt)*100)/100;
  const taxPct  = Math.round(taxRate*100);
  const hasCIS  = taxPct > 0;
  const issued  = new Date().toLocaleDateString("en-GB",{day:"2-digit",month:"short",year:"numeric"});
  const otRate  = w.customOTRate ? "£"+customOT.toFixed(2)+"/hr (custom)" : "£"+customOT.toFixed(2)+"/hr (×"+( w.overtimeMultiplier||1.5)+")";

  const html = `<!DOCTYPE html><html><head><meta charset="utf-8"/>
  <title>Payslip — ${w.name} — WC ${ps.weekLabel}</title>
  <style>
    @media print{body{margin:0}@page{size:A4 portrait;margin:14mm 16mm}}
    *{box-sizing:border-box;margin:0;padding:0}
    body{font-family:Arial,Helvetica,sans-serif;font-size:10px;color:#1a1a1a;background:#fff}
    .hdr{display:flex;justify-content:space-between;align-items:flex-start;padding-bottom:9px;border-bottom:2.5px solid #1a1a2e;margin-bottom:12px}
    .co{font-size:15px;font-weight:700;color:#1a1a2e}.co-s{font-size:8px;color:#999;margin-top:2px}
    .dt{font-size:13px;font-weight:700;color:#1a1a2e;text-align:right}.dt-s{font-size:8px;color:#999;text-align:right;margin-top:1px}
    .strip{display:grid;grid-template-columns:1fr 1fr;border:1px solid #e5e5e5;border-radius:4px;overflow:hidden;margin-bottom:10px}
    .sc{padding:7px 11px;background:#f7f7f7}.sc:first-child{border-right:1px solid #e5e5e5}
    .sl{font-size:7px;font-weight:700;color:#bbb;text-transform:uppercase;letter-spacing:.07em;margin-bottom:2px}
    .sv{font-size:12px;font-weight:700;color:#1a1a2e}.ss{font-size:8px;color:#777;margin-top:1px}
    .id{display:grid;grid-template-columns:${hasCIS?"1fr 1fr 1fr":"1fr 1fr"};gap:6px;margin-bottom:10px}
    .ic{border:1px solid #ebebeb;border-radius:4px;padding:6px 9px;background:#fafafa}
    .il{font-size:7px;font-weight:700;color:#bbb;text-transform:uppercase;letter-spacing:.07em;margin-bottom:2px}
    .iv{font-size:11px;font-weight:700;color:#1a1a2e}
    table{width:100%;border-collapse:collapse;margin-bottom:10px}
    thead tr{background:#1a1a2e}
    thead th{padding:6px 9px;color:#fff;font-size:7.5px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;text-align:left}
    thead th:last-child{text-align:right}
    tbody td{padding:8px 9px;border-bottom:1px solid #f3f3f3;font-size:10px;color:#1a1a2e}
    td.r{text-align:right}td.m{color:#999;font-size:9px}td.ot{color:#92400e;font-weight:700}
    tr.sub td{background:#f5f5f5;font-weight:700}tr.sub td.r{font-size:13px}
    tr.tax td{color:#991b1b}tr.tax td.r{font-weight:700;color:#991b1b}
    tr.net{background:#1a1a2e}tr.net td{padding:10px 9px;color:#fff;font-weight:700;font-size:11px;border:none}
    tr.net td.r{font-size:18px;color:${hasCIS?"#fbbf24":"#34d399"};text-align:right}
    .pay{border:1px solid #e5e5e5;border-radius:4px;overflow:hidden;margin-bottom:10px}
    .ph{padding:6px 11px;background:#1a1a2e}
    .pht{font-size:7.5px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.06em}
    .pr{display:flex;align-items:center;padding:7px 11px;border-bottom:1px solid #f3f3f3}
    .pr:last-child{border-bottom:none}
    .pl{font-size:8px;font-weight:700;color:#bbb;text-transform:uppercase;letter-spacing:.06em;min-width:88px}
    .pv{font-size:10px;color:#1a1a2e;font-weight:600}
    .pc{background:#f0fdf4}.pc .pl{color:#166534}.pc .pv{color:#166534}
    .note{font-size:8px;color:#aaa;line-height:1.55;padding:7px 10px;background:#fafafa;border-left:2px solid #e0e0e0;border-radius:2px;margin-bottom:8px}
    .foot{display:flex;justify-content:space-between;font-size:7.5px;color:#ccc;padding-top:6px;border-top:1px solid #eee}
  </style></head><body>

  <div class="hdr">
    <div><div class="co">Bright Metalwork Ltd</div>
    <div class="co-s">lucian@bright-group.org · +44 (0)771 078 3500 · Unit 4, Empire Way, Aldershot GU11 1PJ</div></div>
    <div><div class="dt">Payslip</div>
    <div class="dt-s">WC ${ps.weekLabel} · Issued ${issued}</div></div>
  </div>

  <div class="strip">
    <div class="sc">
      <div class="sl">Operative</div>
      <div class="sv">${w.name}</div>
      <div class="ss">${w.position||""}${w.company?" · "+w.company:""}</div>
      <div class="ss">${w.address||""}</div>
    </div>
    <div class="sc">
      <div class="sl">Payment period</div>
      <div class="sv">WC ${ps.weekLabel}</div>
      <div class="ss" style="color:${hasCIS?"#854d0e":"#166534"};font-weight:700;margin-top:4px">CIS Subcontractor${hasCIS?" · "+taxPct+"% deduction":""}</div>
    </div>
  </div>

  <div class="id">
    <div class="ic"><div class="il">NINO</div><div class="iv">${w.nino||"—"}</div></div>
    <div class="ic"><div class="il">UTR</div><div class="iv">${w.utr||"—"}</div></div>
    ${hasCIS?`<div class="ic"><div class="il">CIS Rate</div><div class="iv" style="color:${taxPct===30?"#991b1b":"#854d0e"};font-size:14px">${taxPct}%</div></div>`:""}
  </div>

  <table>
    <thead><tr>
      <th style="width:38%">Description</th>
      <th style="width:15%;text-align:right">Hours</th>
      <th style="width:22%;text-align:right">Rate</th>
      <th style="width:25%;text-align:right">Amount</th>
    </tr></thead>
    <tbody>
      <tr>
        <td>Normal hours</td>
        <td class="r">${(ps.stdH||0).toFixed(2)} h</td>
        <td class="r m">£${(w.agreedRate||0).toFixed(2)} / hr</td>
        <td class="r">£${stdAmt.toFixed(2)}</td>
      </tr>
      ${(ps.otH||0)>0?`<tr>
        <td>Overtime hours</td>
        <td class="r">${(ps.otH||0).toFixed(2)} h</td>
        <td class="r m">${otRate}</td>
        <td class="r ot">£${otAmt.toFixed(2)}</td>
      </tr>`:""}
      <tr class="sub">
        <td colspan="3" style="font-size:8.5px;color:#777;text-transform:uppercase;letter-spacing:.04em">Gross pay</td>
        <td class="r" style="font-size:13px">£${gross.toFixed(2)}</td>
      </tr>
      ${hasCIS?`<tr class="tax">
        <td>CIS tax deduction (${taxPct}%)</td>
        <td colspan="2"></td>
        <td class="r">− £${taxAmt.toFixed(2)}</td>
      </tr>`:""}
      <tr class="net">
        <td>Net pay</td>
        <td colspan="2" style="font-size:8px;color:#9ca3af">${hasCIS?"After "+taxPct+"% CIS deduction of £"+taxAmt.toFixed(2):"Paid to operative's bank account"}</td>
        <td class="r">£${net.toFixed(2)}</td>
      </tr>
    </tbody>
  </table>

  <div class="pay">
    <div class="ph"><div class="pht">Payment details</div></div>
    <div class="pr"><span class="pl">Bank</span><span class="pv">${w.bankName||"—"}</span></div>
    <div class="pr"><span class="pl">Sort code</span><span class="pv">${w.bankSort||"—"}</span></div>
    <div class="pr"><span class="pl">Account no</span><span class="pv">${w.bankAccount?"••••"+String(w.bankAccount).slice(-4):"—"}</span></div>
    <div class="pr pc"><span class="pl">CIS paid to</span><span class="pv">Bright Metalwork Ltd · Unit 4, Empire Way, Aldershot GU11 1PJ</span></div>
  </div>

  <div class="note">This payslip is issued under the Construction Industry Scheme (CIS). CIS deductions may be offset against your tax liability. Retain this document for your records. Queries: lucian@bright-group.org · +44 (0)771 078 3500</div>

  <div class="foot">
    <span>Bright Metalwork Ltd · Confidential</span>
    <span>Payslip · ${w.name} · WC ${ps.weekLabel}</span>
    <span>Page 1 of 1</span>
  </div>
  <script>window.onload=()=>window.print();<\/script>
  </body></html>`;

  const win = window.open("","_blank");
  if(win){win.document.write(html);win.document.close();}
  else{const a=document.createElement("a");a.href="data:text/html,"+encodeURIComponent(html);a.download="Payslip_"+w.name.replace(/ /g,"_")+"_WC_"+ps.weekLabel.replace(/ /g,"_")+".html";a.click();}
}

// ─── Export Timesheet PDF ─────────────────────────────────────────────────────
function exportTimesheetPDF(ts, worker) {
  const w = worker;
  const entries = ts.workEntries || [];
  const DAYS = ["Mon","Tue","Wed","Thu","Fri","Sat","Sun"];
  const dayMap = {};
  entries.forEach(l => {
    const hrs = Math.round(((new Date(l.signOut) - new Date(l.signIn)) / 3600000) * 100) / 100;
    const otThr = 9;
    if (!dayMap[l.day]) dayMap[l.day] = { total: 0, std: 0, ot: 0 };
    dayMap[l.day].total = Math.round((dayMap[l.day].total + hrs) * 100) / 100;
    dayMap[l.day].std   = Math.min(dayMap[l.day].total, otThr);
    dayMap[l.day].ot    = Math.max(0, Math.round((dayMap[l.day].total - otThr) * 100) / 100);
  });
  const stdHrs = Object.values(dayMap).reduce((a, d) => a + d.std, 0);
  const otHrs  = Object.values(dayMap).reduce((a, d) => a + d.ot,  0);
  const totHrs = Math.round((stdHrs + otHrs) * 100) / 100;
  const statusLabel = { open: "Open", pending: "Pending review", approved: "Approved" }[ts.status] || ts.status;
  const statusColor = { open: "#1d4ed8", pending: "#854d0e", approved: "#166534" }[ts.status] || "#555";
  const statusBg    = { open: "#eff6ff", pending: "#fef9c3", approved: "#f0fdf4" }[ts.status] || "#f5f5f5";

  const rowsHTML = entries.map((l, i) => {
    const hrs = Math.round(((new Date(l.signOut) - new Date(l.signIn)) / 3600000) * 100) / 100;
    const otThr = 9;
    const std = Math.min(hrs, otThr);
    const ot  = Math.max(0, Math.round((hrs - otThr) * 100) / 100);
    const tIn  = l.signIn  ? new Date(l.signIn).toLocaleTimeString("en-GB",  { hour: "2-digit", minute: "2-digit" }) : "—";
    const tOut = l.signOut ? new Date(l.signOut).toLocaleTimeString("en-GB", { hour: "2-digit", minute: "2-digit" }) : "—";
    const dateStr = l.signIn ? new Date(l.signIn).toLocaleDateString("en-GB", { weekday: "short", day: "2-digit", month: "short" }) : "—";
    const rowBg = i % 2 === 0 ? "#fff" : "#f9f9f9";
    const rep = l.workReport;
    const repCell = rep ? (rep.type === "positive" ? '<td style="padding:9px 12px;text-align:center;color:#166534;font-weight:700">✓</td>' : '<td style="padding:9px 12px;text-align:center;color:#991b1b;font-weight:700">⚠</td>') : '<td style="padding:9px 12px;text-align:center;color:#ccc">·</td>';
    return `<tr style="background:${rep?.type==="negative"?"#fff5f5":rowBg};border-bottom:1px solid #f0f0f0">
      <td style="padding:9px 12px;font-weight:600;color:#1a1a2e">${dateStr}</td>
      <td style="padding:9px 12px;color:#444">${l.siteName || "—"}</td>
      <td style="padding:9px 12px;text-align:center;color:#166534;font-weight:600">${tIn}</td>
      <td style="padding:9px 12px;text-align:center;color:#991b1b;font-weight:600">${tOut}</td>
      <td style="padding:9px 12px;text-align:right;font-weight:600;color:#1a1a2e">${std.toFixed(2)}</td>
      <td style="padding:9px 12px;text-align:right;color:${ot>0?"#92400e":"#ccc"};font-weight:${ot>0?700:400}">${ot > 0 ? ot.toFixed(2) : "—"}</td>
    </tr>`;
  }).join("");

  const negReports = entries.filter(l => l.workReport?.type === "negative");
  const negBlock = negReports.length > 0 ? negReports.map(l => {
    const dateStr = l.signIn ? new Date(l.signIn).toLocaleDateString("en-GB", { weekday: "short", day: "2-digit", month: "short" }) : "—";
    const acts = (l.workReport.activities || []).map(a => `<div style="margin-bottom:8px;padding:8px 10px;background:#fff;border-left:3px solid #991b1b;border-radius:3px"><div style="font-size:10px;color:#991b1b;font-weight:700;margin-bottom:3px">${a.duration || ""}</div><div style="font-size:11px;color:#7f1d1d;line-height:1.5">${a.text || ""}</div></div>`).join("");
    return `<div style="margin-bottom:8px;padding:10px 13px;background:#fff1f2;border:1px solid #fca5a5;border-radius:5px"><div style="font-size:8.5px;font-weight:700;color:#991b1b;text-transform:uppercase;letter-spacing:.05em;margin-bottom:6px">⚠ ${dateStr} — ${l.siteName || ""}</div>${acts}</div>`;
  }).join("") : "";

  const html = `<!DOCTYPE html><html><head><meta charset="utf-8"/>
  <title>Timesheet — ${w.name} — WC ${ts.weekLabel}</title>
  <style>
    @media print { body { margin: 0; } @page { size: A4 landscape; margin: 18mm 16mm; } }
    body { font-family: Arial, Helvetica, sans-serif; font-size: 11px; color: #1a1a1a; background: #fff; }
    table { border-collapse: collapse; width: 100%; }
    th, td { padding: 0; }
  </style></head><body>
  <div style="display:flex;justify-content:space-between;align-items:flex-start;padding-bottom:10px;border-bottom:2.5px solid #1a1a2e;margin-bottom:14px">
    <div><div style="font-size:18px;font-weight:700;color:#1a1a2e">Bright Metalwork Ltd</div><div style="font-size:9px;color:#888;margin-top:2px">lucian@bright-group.org · +44 (0)771 078 3500 · Aldershot, Hampshire</div></div>
    <div style="text-align:right"><div style="font-size:15px;font-weight:700;color:#1a1a2e">Weekly Timesheet</div><div style="font-size:9px;color:#888;margin-top:2px">Week Commencing: ${ts.weekLabel}</div></div>
  </div>
  <div style="display:grid;grid-template-columns:2.5fr 1.2fr 1fr;gap:0;border:1px solid #e0e0e0;border-radius:5px;overflow:hidden;margin-bottom:14px">
    <div style="padding:8px 12px;background:#f5f5f5;border-right:1px solid #e0e0e0"><div style="font-size:7.5px;font-weight:700;color:#aaa;text-transform:uppercase;letter-spacing:.07em;margin-bottom:3px">Worker</div><div style="font-size:14px;font-weight:700;color:#1a1a2e">${w.name}</div><div style="font-size:9px;color:#666;margin-top:1px">${w.position || ""}${w.company ? " · " + w.company : ""}</div></div>
    <div style="padding:8px 12px;background:#f5f5f5;border-right:1px solid #e0e0e0"><div style="font-size:7.5px;font-weight:700;color:#aaa;text-transform:uppercase;letter-spacing:.07em;margin-bottom:3px">Week commencing</div><div style="font-size:13px;font-weight:700;color:#1a1a2e">${ts.weekLabel}</div></div>
    <div style="padding:8px 12px;background:#f5f5f5"><div style="font-size:7.5px;font-weight:700;color:#aaa;text-transform:uppercase;letter-spacing:.07em;margin-bottom:4px">Status</div><div style="display:inline-flex;align-items:center;gap:5px;padding:3px 9px;border-radius:3px;background:${statusBg};border:1px solid ${statusColor}44"><div style="width:6px;height:6px;border-radius:50%;background:${statusColor}"></div><span style="font-size:10px;font-weight:700;color:${statusColor}">${statusLabel}</span></div></div>
  </div>
  <table>
    <thead><tr style="background:#1a1a2e">
      <th style="padding:8px 12px;text-align:left;color:#fff;font-size:8.5px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;width:16%">Date</th>
      <th style="padding:8px 12px;text-align:left;color:#fff;font-size:8.5px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;width:32%">Site</th>
      <th style="padding:8px 12px;text-align:center;color:#fff;font-size:8.5px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;width:13%">Sign in</th>
      <th style="padding:8px 12px;text-align:center;color:#fff;font-size:8.5px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;width:13%">Sign out</th>
      <th style="padding:8px 12px;text-align:right;color:#fff;font-size:8.5px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;width:13%">Normal hrs</th>
      <th style="padding:8px 12px;text-align:right;color:#fff;font-size:8.5px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;width:13%">OT hrs</th>
    </tr></thead>
    <tbody>${rowsHTML}</tbody>
    <tfoot><tr style="background:#1a1a2e">
      <td colspan="4" style="padding:9px 12px;font-weight:700;font-size:9px;text-transform:uppercase;letter-spacing:.05em;color:#9ca3af">Weekly totals</td>
      <td style="padding:9px 12px;text-align:right;font-weight:700;font-size:15px;color:#fff">${stdHrs.toFixed(2)}</td>
      <td style="padding:9px 12px;text-align:right;font-weight:700;font-size:15px;color:#fcd34d">${otHrs > 0 ? otHrs.toFixed(2) : "—"}</td>
    </tr></tfoot>
  </table>
  <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;margin-top:14px;margin-bottom:${negBlock?"12px":"14px"}">
    <div style="border:1px solid #e5e5e5;border-radius:5px;padding:9px 12px;background:#f5f5f5"><div style="font-size:8px;font-weight:700;color:#aaa;text-transform:uppercase;letter-spacing:.06em;margin-bottom:4px">Normal hours</div><div style="font-size:20px;font-weight:700;color:#1a1a2e">${stdHrs.toFixed(2)} h</div></div>
    <div style="border:1px solid #fde68a;border-radius:5px;padding:9px 12px;background:#fffbeb"><div style="font-size:8px;font-weight:700;color:#aaa;text-transform:uppercase;letter-spacing:.06em;margin-bottom:4px">Overtime hours</div><div style="font-size:20px;font-weight:700;color:#92400e">${otHrs.toFixed(2)} h</div></div>
    <div style="border:1px solid #1a1a2e;border-radius:5px;padding:9px 12px;background:#1a1a2e"><div style="font-size:8px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.06em;margin-bottom:4px">Total hours</div><div style="font-size:20px;font-weight:700;color:#fff">${totHrs.toFixed(2)} h</div></div>
  </div>
  ${negBlock}
  <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:24px;padding-top:14px;border-top:1.5px solid #e0e0e0;margin-top:${negBlock?"0":"4px"}">
    <div><div style="font-size:8px;font-weight:700;color:#aaa;text-transform:uppercase;letter-spacing:.06em;margin-bottom:8px">Worker signature</div><div style="height:38px;border-bottom:1.5px solid #1a1a2e;margin-bottom:4px"></div><div style="font-size:9px;color:#aaa">${w.name}</div><div style="font-size:9px;color:#ccc;margin-top:2px">Date: ________________________</div></div>
    <div><div style="font-size:8px;font-weight:700;color:#aaa;text-transform:uppercase;letter-spacing:.06em;margin-bottom:8px">Site Manager sign-off</div><div style="height:38px;border-bottom:1.5px solid #1a1a2e;margin-bottom:4px"></div><div style="font-size:9px;color:#aaa">Site Manager</div><div style="font-size:9px;color:#ccc;margin-top:2px">Date: ________________________</div></div>
    <div><div style="font-size:8px;font-weight:700;color:#aaa;text-transform:uppercase;letter-spacing:.06em;margin-bottom:8px">Director / Approval</div><div style="height:38px;border-bottom:1.5px solid #1a1a2e;margin-bottom:4px"></div><div style="font-size:9px;color:#aaa">Lucian Ciocoiu</div><div style="font-size:9px;color:#ccc;margin-top:2px">Date: ________________________</div></div>
  </div>
  <div style="margin-top:12px;padding-top:7px;border-top:1px solid #eee;display:flex;justify-content:space-between;font-size:8px;color:#ccc"><span>Bright Metalwork Ltd · Confidential</span><span>Timesheet · ${w.name} · WC ${ts.weekLabel}</span><span>Page 1 of 1</span></div>
  <script>window.onload=()=>window.print();<\/script>
  </body></html>`;
  const win = window.open("", "_blank");
  if (win) { win.document.write(html); win.document.close(); }
  else { const a = document.createElement("a"); a.href = "data:text/html," + encodeURIComponent(html); a.download = `Timesheet_${w.name.replace(/ /g,"_")}_WC_${ts.weekLabel.replace(/ /g,"_")}.html`; a.click(); }
}

// ─── Export Work Report PDF ───────────────────────────────────────────────────
function exportWorkReportPDF(entry, worker) {
  const w = worker;
  const rep = entry.workReport;
  if (!rep) return;
  const isNeg = rep.type === "negative";
  const dateStr = entry.signIn ? new Date(entry.signIn).toLocaleDateString("en-GB", { weekday: "long", day: "2-digit", month: "long", year: "numeric" }) : "—";
  const tIn  = entry.signIn  ? new Date(entry.signIn).toLocaleTimeString("en-GB",  { hour: "2-digit", minute: "2-digit" }) : "—";
  const tOut = entry.signOut ? new Date(entry.signOut).toLocaleTimeString("en-GB", { hour: "2-digit", minute: "2-digit" }) : "—";

  const actsHTML = (rep.activities || []).map((a, i) => `
    <div style="margin-bottom:${i < (rep.activities.length - 1) ? "18px" : "8px"};padding-bottom:${i < (rep.activities.length - 1) ? "18px" : "0"};border-bottom:${i < (rep.activities.length - 1) ? "1px solid #f0f0f0" : "none"}">
      <div style="display:flex;align-items:center;gap:10px;margin-bottom:7px">
        <div style="font-size:8.5px;font-weight:700;color:#aaa;text-transform:uppercase;letter-spacing:.06em">Activity ${i + 1}</div>
        ${a.duration ? `<span style="font-size:10px;font-weight:700;color:#1e3a8a;padding:2px 8px;background:#eff6ff;border:1px solid #bfdbfe;border-radius:3px">${a.duration}</span>` : ""}
      </div>
      ${a.text ? `<div style="font-size:12px;color:#1a1a2e;line-height:1.65;padding:9px 11px;background:#f8f8f8;border-radius:4px;border-left:3px solid #1a1a2e;margin-bottom:${(a.photos && a.photos.length) ? "10px" : "0"}">${a.text}</div>` : ""}
      ${a.photos && a.photos.length ? `<div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-top:${a.text ? "0" : "0"}">${a.photos.map(p => `<div style="border-radius:4px;overflow:hidden;border:1px solid #e5e5e5"><img src="${p.src || ""}" alt="${p.name}" style="width:100%;height:140px;object-fit:cover;display:block"><div style="padding:3px 7px;font-size:8px;color:#aaa;background:#fafafa;border-top:1px solid #f0f0f0">${p.name}</div></div>`).join("")}</div>` : ""}
    </div>`).join("");

  const html = `<!DOCTYPE html><html><head><meta charset="utf-8"/>
  <title>Work Report — ${w.name} — ${entry.date || dateStr}</title>
  <style>
    @media print { body { margin: 0; } @page { size: A4 portrait; margin: 16mm 16mm; } }
    body { font-family: Arial, Helvetica, sans-serif; font-size: 11px; color: #1a1a1a; background: #fff; }
  </style></head><body>
  <div style="display:flex;justify-content:space-between;align-items:flex-start;padding-bottom:11px;border-bottom:2.5px solid #1a1a2e;margin-bottom:16px">
    <div><div style="font-size:17px;font-weight:700;color:#1a1a2e">Bright Metalwork Ltd</div><div style="font-size:9px;color:#888;margin-top:2px">lucian@bright-group.org · +44 (0)771 078 3500</div></div>
    <div style="text-align:right"><div style="font-size:14px;font-weight:700;color:#1a1a2e">Daily Work Report</div><div style="font-size:9px;color:#888;margin-top:1px">${new Date().toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" })}</div></div>
  </div>
  <table style="width:100%;border-collapse:collapse;margin-bottom:14px;font-size:11px">
    <tr>
      <td style="padding:8px 12px;background:#f8f8f8;border:1px solid #e8e8e8;width:28%;vertical-align:top"><div style="font-size:8px;font-weight:700;color:#aaa;text-transform:uppercase;letter-spacing:.06em;margin-bottom:2px">Worker</div><div style="font-weight:700;color:#1a1a2e">${w.name}</div><div style="color:#666;font-size:10px">${w.position || ""}</div></td>
      <td style="padding:8px 12px;background:#f8f8f8;border:1px solid #e8e8e8;border-left:none;width:22%;vertical-align:top"><div style="font-size:8px;font-weight:700;color:#aaa;text-transform:uppercase;letter-spacing:.06em;margin-bottom:2px">Date</div><div style="font-weight:700;color:#1a1a2e">${dateStr}</div></td>
      <td style="padding:8px 12px;background:#f8f8f8;border:1px solid #e8e8e8;border-left:none;width:28%;vertical-align:top"><div style="font-size:8px;font-weight:700;color:#aaa;text-transform:uppercase;letter-spacing:.06em;margin-bottom:2px">Site</div><div style="font-weight:700;color:#1a1a2e">${entry.siteName || "—"}</div></td>
      <td style="padding:8px 12px;background:#f8f8f8;border:1px solid #e8e8e8;border-left:none;width:22%;vertical-align:top"><div style="font-size:8px;font-weight:700;color:#aaa;text-transform:uppercase;letter-spacing:.06em;margin-bottom:2px">Scope</div><div style="font-weight:700;color:#1a1a2e">${entry.scopeName || "—"}</div></td>
    </tr>
  </table>
  <div style="margin-bottom:16px">
    <div style="display:inline-flex;align-items:center;gap:7px;padding:6px 14px;border-radius:4px;background:${isNeg ? "#fff1f2" : "#f0fdf4"};border:1px solid ${isNeg ? "#fca5a5" : "#86efac"}">
      <div style="width:9px;height:9px;border-radius:50%;background:${isNeg ? "#dc2626" : "#16a34a"}"></div>
      <span style="font-size:12px;font-weight:700;color:${isNeg ? "#991b1b" : "#166534"}">${isNeg ? "⚠ Negative — Issues reported" : "✓ Positive — Work completed as planned"}</span>
    </div>
  </div>
  ${actsHTML}
  <div style="display:grid;grid-template-columns:1fr 1fr;gap:20px;padding-top:14px;border-top:1.5px solid #e0e0e0;margin-top:16px">
    <div><div style="font-size:8px;font-weight:700;color:#aaa;text-transform:uppercase;letter-spacing:.06em;margin-bottom:8px">Worker signature</div><div style="height:38px;border-bottom:1.5px solid #1a1a2e;margin-bottom:4px"></div><div style="font-size:9px;color:#aaa">${w.name}</div><div style="font-size:9px;color:#ccc;margin-top:2px">Date: ________________________</div></div>
    <div><div style="font-size:8px;font-weight:700;color:#aaa;text-transform:uppercase;letter-spacing:.06em;margin-bottom:8px">Supervisor sign-off</div><div style="height:38px;border-bottom:1.5px solid #1a1a2e;margin-bottom:4px"></div><div style="font-size:9px;color:#aaa">Site Manager</div><div style="font-size:9px;color:#ccc;margin-top:2px">Date: ________________________</div></div>
  </div>
  <div style="margin-top:12px;padding-top:7px;border-top:1px solid #eee;display:flex;justify-content:space-between;font-size:8px;color:#ccc"><span>Bright Metalwork Ltd · Confidential</span><span>Daily Work Report · ${w.name} · ${entry.date || dateStr}</span><span>Page 1 of 1</span></div>
  <script>window.onload=()=>window.print();<\/script>
  </body></html>`;
  const win = window.open("", "_blank");
  if (win) { win.document.write(html); win.document.close(); }
  else { const a = document.createElement("a"); a.href = "data:text/html," + encodeURIComponent(html); a.download = `WorkReport_${w.name.replace(/ /g,"_")}_${(entry.date||"").replace(/ /g,"_")}.html`; a.click(); }
}



// ─── Work Report Modal ────────────────────────────────────────────────────────
const DURATIONS=["15 min","30 min","45 min","1h","1h 15","1h 30","1h 45","2h","2h 30","3h","3h 30","4h","4h 30","5h","5h 30","6h","6h 30","7h","7h 30","8h","9h"];
function WorkReportModal({entry,worker,onSave,onClose}){
  const existing=entry.workReport||null;
  const [type,setType]=useState(existing?.type||"positive");
  const [activities,setActivities]=useState(existing?.activities||[{id:1,text:"",duration:"",photos:[]}]);
  const [actCount,setActCount]=useState(existing?.activities?.length||1);
  const w=worker;
  const dateStr=entry.signIn?new Date(entry.signIn).toLocaleDateString("en-GB",{weekday:"short",day:"2-digit",month:"short"}):"—";
  const tIn=entry.signIn?new Date(entry.signIn).toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit"}):"—";
  const tOut=entry.signOut?new Date(entry.signOut).toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit"}):"—";
  const isNeg=type==="negative";
  const addAct=()=>{const id=actCount+1;setActCount(id);setActivities(a=>[...a,{id,text:"",duration:"",photos:[]}]);};
  const removeAct=(id)=>{if(activities.length===1)return;setActivities(a=>a.filter(x=>x.id!==id));};
  const updateAct=(id,k,v)=>setActivities(a=>a.map(x=>x.id===id?{...x,[k]:v}:x));
  const addPhoto=(id,file)=>{if(!file)return;const r=new FileReader();r.onload=e=>updateAct(id,"photos",[...(activities.find(a=>a.id===id)?.photos||[]),{name:file.name,src:e.target.result}]);r.readAsDataURL(file);};
  const removePhoto=(id,pi)=>updateAct(id,"photos",activities.find(a=>a.id===id)?.photos.filter((_,i)=>i!==pi)||[]);
  const save=()=>{
    const report={type,activities,createdAt:new Date().toISOString()};
    if(isNeg){
      // Notification logic would fire here
      if(window.confirm("This is a negative report. Manager and Director will be notified automatically. Save?"))
        onSave({...entry,workReport:report});
    } else {
      onSave({...entry,workReport:report});
    }
  };
  const INP={width:"100%",background:"#0d1117",border:"1px solid #1e2535",borderRadius:7,padding:"9px 11px",color:"#f1f5f9",fontSize:13,outline:"none",resize:"vertical",fontFamily:"inherit",boxSizing:"border-box"};
  return <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.75)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:300,padding:16}}>
    <div style={{background:"#111827",borderRadius:14,width:"100%",maxWidth:540,maxHeight:"90vh",overflowY:"auto",border:"1px solid #1e2535",boxShadow:"0 24px 60px rgba(0,0,0,.6)"}}>
      {/* Header */}
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"16px 18px",borderBottom:"1px solid #1e2535",position:"sticky",top:0,background:"#111827",zIndex:1}}>
        <div>
          <div style={{fontSize:14,fontWeight:700,color:"#f1f5f9"}}>Daily Work Report</div>
          <div style={{fontSize:11,color:"#64748b",marginTop:2}}>{dateStr} · {w?.name} · {entry.siteName} · {tIn}–{tOut}</div>
        </div>
        <button onClick={onClose} style={{background:"none",border:"none",color:"#64748b",cursor:"pointer",fontSize:20,padding:"0 4px",lineHeight:1}}>×</button>
      </div>
      <div style={{padding:18}}>
        {/* Type */}
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:16}}>
          {[["positive","✓","Positive","Work completed as planned","#34d399","#0d2218","#34d39944"],
            ["negative","⚠","Negative","Issues, delays or problems","#f87171","#2d1515","#f8717144"]].map(([t,ic,lbl,sub,col,bg,bd])=>(
            <button key={t} onClick={()=>setType(t)} style={{padding:"11px 14px",borderRadius:10,border:`2px solid ${type===t?col:bd}`,background:type===t?bg:"#0d1117",cursor:"pointer",textAlign:"left",transition:"all .15s"}}>
              <div style={{fontSize:20,marginBottom:3}}>{ic}</div>
              <div style={{fontSize:13,fontWeight:700,color:type===t?col:"#94a3b8"}}>{lbl}</div>
              <div style={{fontSize:11,color:type===t?col:"#374151",marginTop:1}}>{sub}</div>
            </button>
          ))}
        </div>
        {isNeg&&<div style={{background:"#2d1515",border:"1px solid #f8717144",borderRadius:9,padding:"10px 13px",marginBottom:14,fontSize:12,color:"#f87171",display:"flex",gap:8,alignItems:"center"}}>
          <span style={{fontSize:16}}>📣</span>
          <span>Saving a negative report will automatically notify the <b style={{fontWeight:700}}>Site Manager</b> and <b style={{fontWeight:700}}>Director</b>.</span>
        </div>}
        {/* Activity blocks */}
        {activities.map((a,i)=>(
          <div key={a.id} style={{background:"#0d1117",borderRadius:10,border:"1px solid #1e2535",padding:14,marginBottom:10}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
              <div style={{fontSize:10,fontWeight:700,color:"#64748b",textTransform:"uppercase",letterSpacing:".05em"}}>Activity {i+1}</div>
              {activities.length>1&&<button onClick={()=>removeAct(a.id)} style={{background:"none",border:"none",color:"#374151",cursor:"pointer",fontSize:13}}>✕</button>}
            </div>
            {/* Text + Duration side by side */}
            <div style={{display:"grid",gridTemplateColumns:"1fr 130px",gap:10,marginBottom:12,alignItems:"start"}}>
              <div>
                <div style={{fontSize:10,color:"#64748b",fontWeight:700,textTransform:"uppercase",letterSpacing:".05em",marginBottom:5}}>Description</div>
                <textarea value={a.text} onChange={e=>updateAct(a.id,"text",e.target.value)} placeholder="Describe what was worked on — location, method, drawing reference…" rows={4} style={{...INP}}/>
              </div>
              <div>
                <div style={{fontSize:10,color:"#64748b",fontWeight:700,textTransform:"uppercase",letterSpacing:".05em",marginBottom:5}}>Duration</div>
                <select size={8} value={a.duration} onChange={e=>updateAct(a.id,"duration",e.target.value)}
                  style={{...INP,padding:"5px 8px",height:"auto",cursor:"pointer",fontSize:12}}>
                  <option value="">—</option>
                  {DURATIONS.map(d=><option key={d} value={d}>{d}</option>)}
                </select>
              </div>
            </div>
            {/* Photos */}
            <div style={{fontSize:10,color:"#64748b",fontWeight:700,textTransform:"uppercase",letterSpacing:".05em",marginBottom:8}}>Photos</div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
              {a.photos.map((p,pi)=>(
                <div key={pi} style={{position:"relative",borderRadius:8,overflow:"hidden",border:"1px solid #1e2535",aspectRatio:"4/3"}}>
                  <img src={p.src} alt={p.name} style={{width:"100%",height:"100%",objectFit:"cover",display:"block"}}/>
                  <button onClick={()=>removePhoto(a.id,pi)} style={{position:"absolute",top:5,right:5,width:22,height:22,borderRadius:"50%",background:"rgba(0,0,0,.6)",color:"#fff",border:"none",cursor:"pointer",fontSize:13,display:"flex",alignItems:"center",justifyContent:"center",lineHeight:1}}>×</button>
                  <div style={{position:"absolute",bottom:0,left:0,right:0,padding:"3px 7px",background:"rgba(0,0,0,.5)",fontSize:9,color:"#e2e8f0"}}>{p.name}</div>
                </div>
              ))}
              <label style={{display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",gap:6,border:"1.5px dashed #2d3555",borderRadius:8,cursor:"pointer",aspectRatio:"4/3",color:"#64748b",fontSize:12,background:"#111827"}}>
                <span style={{fontSize:24}}>📷</span>
                <span>Add photo</span>
                <input type="file" accept="image/*" style={{display:"none"}} onChange={e=>{addPhoto(a.id,e.target.files[0]);e.target.value="";}}/>
              </label>
            </div>
          </div>
        ))}
        <button onClick={addAct} style={{width:"100%",padding:"9px",border:"1.5px dashed #2d3555",borderRadius:9,background:"none",color:"#64748b",cursor:"pointer",fontSize:13,display:"flex",alignItems:"center",justifyContent:"center",gap:7,marginBottom:14}}>
          + Add another activity
        </button>
        <div style={{display:"flex",gap:8}}>
          <button onClick={onClose} style={{flex:1,padding:"10px",background:"#1e2535",border:"1px solid #374151",borderRadius:8,color:"#94a3b8",cursor:"pointer",fontSize:13,fontWeight:600}}>Cancel</button>
          <button onClick={save} style={{flex:2,padding:"10px",background:isNeg?"#2d1515":"linear-gradient(135deg,#14532d,#16a34a)",border:isNeg?"1px solid #f87171":"none",borderRadius:8,color:isNeg?"#f87171":"#fff",cursor:"pointer",fontSize:13,fontWeight:700}}>
            {isNeg?"Save & Notify Manager + Director":"Save Report"}
          </button>
        </div>
      </div>
    </div>
  </div>;
}


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
    ${w.contact?"<div class=\"field\"><span class=\"field-label\">Phone</span><span class=\"field-value\">"+w.contact+"</span></div>":""}
    ${w.email?"<div class=\"field\"><span class=\"field-label\">Email</span><span class=\"field-value\">"+w.email+"</span></div>":""}
    ${w.dob?"<div class=\"field\"><span class=\"field-label\">Date of Birth</span><span class=\"field-value\">"+fmtDate(w.dob)+"</span></div>":""}
    ${w.nationality?"<div class=\"field\"><span class=\"field-label\">Nationality</span><span class=\"field-value\">"+w.nationality+"</span></div>":""}
    ${w.address?"<div class=\"field\"><span class=\"field-label\">Address</span><span class=\"field-value\">"+w.address+"</span></div>":""}
    ${w.nextOfKin?"<div class=\"field\"><span class=\"field-label\">Next of Kin</span><span class=\"field-value\">"+(w.nextOfKin+(w.nextOfKinPhone?" · "+w.nextOfKinPhone:""))+"</span></div>":""}
  </div>
  <div class="card">
    <div class="card-title">Role & Qualifications</div>
    <div class="field"><span class="field-label">Position</span><span class="field-value">${w.position||"—"}</span></div>
    <div class="field"><span class="field-label">Company</span><span class="field-value">${w.company||"—"}</span></div>
    ${w.scope?"<div class=\"field\"><span class=\"field-label\">Scope</span><span class=\"field-value\"><span class=\"scope-badge\">"+w.scope+"</span></span></div>":""}
    ${w.shareCode?"<div class=\"field\"><span class=\"field-label\">Right to Work</span><span class=\"field-value\" style=\"color:#34d399;font-weight:700\">Share code verified</span></div>":""}
  </div>
</div>

<!-- site allocation removed from worker profile PDF -->

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
  if(!win){const a=document.createElement("a");a.href=u;a.download="Worker_Profile_"+w.name.split(" ").join("_")+".html";a.click();}
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
  if(!win){const a=document.createElement("a");a.href=u;a.download="Payslip_"+w.name.split(" ").join("_")+"_"+weekLabel.split(" ").join("_")+".html";a.click();}
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
  if(!win){const a=document.createElement("a");a.href=u;a.download="Schedule_WC_"+weekLabel.split(" ").join("_")+".html";a.click();}
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
  XLSX.writeFile(wb,"LabourSchedule_WC_"+weekLabel.split(" ").join("_")+".xlsx");
}

// ─── Manage Sites Modal ───────────────────────────────────────────────────────
function SitesModal({allSites,clients,onSave,onClose,onOpenDetail}){
  const [sites,setSites]=useState(allSites.map(s=>({...s})));
  const [nn,setNn]=useState(""),[nc,setNc]=useState(PRESET_COLORS[0]),[ncl,setNcl]=useState("");
  const add=()=>{const n=nn.trim();if(!n||sites.find(s=>s.name.toLowerCase()===n.toLowerCase()))return;setSites(s=>[...s,{id:"s"+Date.now(),name:n,color:nc,clientId:ncl||null,builtin:false,lat:null,lng:null,radius:100,stdHours:9,startTime:"07:30",otThreshold:9,contractType:"dayrate",pohPct:0,retentionPct:0}]);setNn("");};
  const [locating,setLocating]=useState({});
  const useMyLocation=(id)=>{setLocating(l=>({...l,[id]:true}));navigator.geolocation.getCurrentPosition(pos=>{up(id,"lat",+pos.coords.latitude.toFixed(6));up(id,"lng",+pos.coords.longitude.toFixed(6));setLocating(l=>({...l,[id]:false}));},()=>{alert("Could not get location. Please type coordinates manually.");setLocating(l=>({...l,[id]:false}));},{enableHighAccuracy:true,timeout:10000});};
  const GpsFields=({s})=><div style={{marginTop:8,background:"#0a0e1a",borderRadius:7,padding:"8px 10px",border:"1px solid #1e2535"}}><div style={{fontSize:9,color:"#10b981",fontWeight:700,textTransform:"uppercase",marginBottom:6}}>📍 GPS — Worker Sign In/Out</div><div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:6,marginBottom:6}}><div><div style={{fontSize:9,color:"#64748b",fontWeight:700,textTransform:"uppercase",marginBottom:2}}>Latitude</div><input type="number" step="0.000001" value={s.lat||""} onChange={e=>up(s.id,"lat",+e.target.value||null)} placeholder="51.509865" style={{width:"100%",background:"#1a1f2e",border:"1px solid #2d3555",borderRadius:5,padding:"5px 7px",color:"#e2e8f0",fontSize:11,outline:"none",boxSizing:"border-box"}}/></div><div><div style={{fontSize:9,color:"#64748b",fontWeight:700,textTransform:"uppercase",marginBottom:2}}>Longitude</div><input type="number" step="0.000001" value={s.lng||""} onChange={e=>up(s.id,"lng",+e.target.value||null)} placeholder="-0.118092" style={{width:"100%",background:"#1a1f2e",border:"1px solid #2d3555",borderRadius:5,padding:"5px 7px",color:"#e2e8f0",fontSize:11,outline:"none",boxSizing:"border-box"}}/></div><div><div style={{fontSize:9,color:"#64748b",fontWeight:700,textTransform:"uppercase",marginBottom:2}}>Radius (m)</div><input type="number" min="10" max="2000" step="10" value={s.radius||100} onChange={e=>up(s.id,"radius",+e.target.value||100)} style={{width:"100%",background:"#1a1f2e",border:"1px solid #2d3555",borderRadius:5,padding:"5px 7px",color:"#e2e8f0",fontSize:11,outline:"none",boxSizing:"border-box"}}/></div></div><div style={{marginBottom:8}}><div style={{fontSize:9,color:"#60a5fa",fontWeight:700,textTransform:"uppercase",marginBottom:4}}>⏱ Hours & Overtime</div><div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:6}}><div><div style={{fontSize:9,color:"#64748b",fontWeight:700,textTransform:"uppercase",marginBottom:2}}>Start Time</div><input type="time" value={s.startTime||"07:30"} onChange={e=>up(s.id,"startTime",e.target.value)} style={{width:"100%",background:"#1a1f2e",border:"1px solid #2d3555",borderRadius:5,padding:"5px 7px",color:"#e2e8f0",fontSize:11,outline:"none",boxSizing:"border-box"}}/></div><div><div style={{fontSize:9,color:"#64748b",fontWeight:700,textTransform:"uppercase",marginBottom:2}}>Std Hours/Day</div><input type="number" min="1" max="24" step="0.5" value={s.stdHours||9} onChange={e=>up(s.id,"stdHours",+e.target.value||9)} style={{width:"100%",background:"#1a1f2e",border:"1px solid #2d3555",borderRadius:5,padding:"5px 7px",color:"#e2e8f0",fontSize:11,outline:"none",boxSizing:"border-box"}}/></div><div><div style={{fontSize:9,color:"#64748b",fontWeight:700,textTransform:"uppercase",marginBottom:2}}>OT After (hrs)</div><input type="number" min="1" max="24" step="0.5" value={s.otThreshold||s.stdHours||9} onChange={e=>up(s.id,"otThreshold",+e.target.value||9)} style={{width:"100%",background:"#1a1f2e",border:"1px solid #2d3555",borderRadius:5,padding:"5px 7px",color:"#e2e8f0",fontSize:11,outline:"none",boxSizing:"border-box"}}/></div></div></div><div style={{display:"flex",alignItems:"center",gap:8}}><button onClick={()=>useMyLocation(s.id)} disabled={locating[s.id]} style={{padding:"4px 10px",background:"#0d2218",border:"1px solid #10b981",borderRadius:5,color:"#34d399",cursor:"pointer",fontSize:10,fontWeight:700,opacity:locating[s.id]?0.6:1}}>{locating[s.id]?"Getting…":"📍 Use My Current Location"}</button>{s.lat&&s.lng?<span style={{fontSize:10,color:"#34d399",fontWeight:600}}>✓ GPS set · {s.radius||100}m</span>:<span style={{fontSize:10,color:"#f87171",fontWeight:700}}>🔒 No GPS — workers cannot sign in</span>}</div></div>;
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
          <GpsFields s={s}/>
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
          <GpsFields s={s}/>
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
  const [allPositions,setAllPositions]=useState([...POSITIONS]);
  const [showCustomInput,setShowCustomInput]=useState(!POSITIONS.includes(worker.position)&&!!worker.position);
  const set=(k,v)=>{
    if(["agreedRate","taxRate","overtimeMultiplier","customOTRate"].includes(k)&&worker.id&&worker[k]!==undefined&&worker[k]!==null&&worker[k]!==v){
      const histEntry={changedAt:new Date().toISOString(),[k]:worker[k],changedBy:"Admin"};
      setF(x=>({...x,[k]:v,payRateHistory:[...(x.payRateHistory||[]),histEntry]}));
    } else {
      setF(x=>({...x,[k]:v}));
    }
  };
  const setD=(d,v)=>setF(x=>({...x,days:{...x.days,[d]:v}}));
  const setH=(d,v)=>setF(x=>({...x,hoursPerDay:{...x.hoursPerDay,[d]:Number(v)||0}}));
  const setOT=(d,v)=>setF(x=>({...x,overtimeHours:{...x.overtimeHours,[d]:Number(v)||0}}));
  const setC=(k,v)=>setF(x=>({...x,certs:{...x.certs,[k]:v}}));
  const held=CERTS.filter(c=>f.certs?.[c.key]?.held).length;
  const alerts=CERTS.filter(c=>{const s=cSt(c,f);return s==="expired"||s==="expiring";}).length;
  const SC={valid:"#34d399",expiring:"#fbbf24",expired:"#f87171",missing:"#374151"};
  function saveCustomTrade(){
    if(!f.position||allPositions.includes(f.position)) return;
    const updated=[...allPositions,f.position];
    setAllPositions(updated);
    POSITIONS.splice(0,POSITIONS.length,...updated);
  }
  return <Overlay onClose={onClose} wide>
    <MH title={worker.name?`Edit: ${worker.name}`:"Add New Worker"} onClose={onClose}/>
    <TabBar tabs={[["personal","👤 Personal"],["schedule","📅 Schedule"],["pay","💷 Pay & OT"],["certs","🛡 Certs "+(alerts>0?"⚠"+alerts:"("+held+")")]]} active={tab} onChange={setTab}/>
    {tab==="personal"&&<div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"0 14px"}}>
        <FI label="Full Name" value={f.name} onChange={v=>set("name",v)}/><FSel label="Company" value={f.company} onChange={v=>set("company",v)} options={COMPANIES}/>
        <div style={{marginBottom:11}}>
          <label style={LBL}>Position / Trade</label>
          <div style={{display:"flex",gap:6}}>
            <select value={allPositions.includes(f.position)?f.position:"__custom"}
              onChange={e=>{
                if(e.target.value==="__custom"){setShowCustomInput(true);set("position","");}
                else{setShowCustomInput(false);set("position",e.target.value);}
              }}
              style={{...INP,cursor:"pointer",flex:1}}>
              <option value="">— Select position —</option>
              {allPositions.map(p=><option key={p} value={p}>{p}</option>)}
              <option value="__custom">+ Add custom trade…</option>
            </select>
            {showCustomInput&&<input value={f.position} onChange={e=>set("position",e.target.value)}
              placeholder="Type custom trade name…" style={{...INP,flex:1}} autoFocus/>}
          </div>
          {showCustomInput&&f.position&&!allPositions.includes(f.position)&&<div style={{marginTop:6,display:"flex",alignItems:"center",gap:8,padding:"5px 10px",background:"#1a1500",borderRadius:6,border:"1px solid #92400e"}}>
            <span style={{fontSize:10,color:"#fbbf24",flex:1}}>New trade: <strong>{f.position}</strong> — save to add permanently to all dropdowns</span>
            <button onClick={saveCustomTrade} style={{padding:"3px 10px",background:"#1e3a5f",border:"1px solid #3b82f6",borderRadius:4,color:"#60a5fa",cursor:"pointer",fontSize:10,fontWeight:700,whiteSpace:"nowrap"}}>+ Save to list</button>
          </div>}
        </div>
        <FI label="Scope" value={f.scope} onChange={v=>set("scope",v)}/>
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
      <Sec title="Personal Details" color="#60a5fa">
        <FI label="Nationality" value={f.nationality||""} onChange={v=>set("nationality",v)} placeholder="e.g. Romanian, British…"/>
      </Sec>
      <Sec title="Right to Work in UK" color="#34d399">
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:"0 12px"}}>
          <FI label="Share Code" value={f.shareCode||""} onChange={v=>set("shareCode",v)} placeholder="e.g. W4P-B7C-XY3"/>
          <FI label="Date Added / Checked" value={f.shareCodeDate||""} onChange={v=>set("shareCodeDate",v)} type="date"/>
          <FI label="Expiry Date" value={f.shareCodeExpiry||""} onChange={v=>set("shareCodeExpiry",v)} type="date"/>
        </div>
        <div style={{marginTop:10}}>
          <label style={LBL}>Attach Documents (ID, RTW proof, Driving Licence, Signed Agreement)</label>
          <div style={{display:"flex",flexWrap:"wrap",gap:8,marginBottom:8}}>
            {(f.workerFiles||[]).map((wf,i)=>{
              const isImg=wf.type&&wf.type.startsWith("image/");
              const icon=isImg?"🖼":wf.type==="application/pdf"?"📄":wf.type&&wf.type.includes("word")?"📝":"📎";
              return <div key={i} style={{display:"flex",alignItems:"center",gap:6,padding:"5px 10px",background:"#1a2035",borderRadius:7,border:"1px solid #2d3555",maxWidth:200}}>
                <span>{icon}</span>
                <a href={wf.url} target="_blank" rel="noreferrer" style={{fontSize:10,color:"#60a5fa",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",maxWidth:120,textDecoration:"none"}} title={wf.name}>{wf.name}</a>
                <button onClick={()=>set("workerFiles",(f.workerFiles||[]).filter((_,j)=>j!==i))} style={{background:"none",border:"none",color:"#f87171",cursor:"pointer",fontSize:13,lineHeight:1,padding:0}}>×</button>
              </div>;
            })}
          </div>
          <label style={{display:"inline-flex",alignItems:"center",gap:7,padding:"7px 14px",background:"#1e3a5f",border:"1px solid #3b82f6",borderRadius:7,cursor:"pointer",color:"#60a5fa",fontSize:11,fontWeight:700}}>
            📎 Attach File
            <input type="file" multiple accept="image/*,.pdf,.doc,.docx" style={{display:"none"}} onChange={e=>{
              const files=Array.from(e.target.files||[]);
              files.forEach(file=>{
                const reader=new FileReader();
                reader.onload=ev=>set("workerFiles",[...(f.workerFiles||[]),{name:file.name,type:file.type,size:file.size,url:ev.target.result,addedAt:new Date().toISOString()}]);
                reader.readAsDataURL(file);
              });
              e.target.value="";
            }}/>
          </label>
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
      {/* ── App Role — controls what this user sees when they log in ── */}
      <div style={{marginBottom:14,padding:"10px 13px",background:"#0d1117",borderRadius:9,border:"1px solid #3b82f644"}}>
        <label style={{...LBL,color:"#60a5fa"}}>App Role — Login Access</label>
        <select value={f.userRole||"worker"} onChange={e=>set("userRole",e.target.value)}
          style={{...INP,cursor:"pointer",border:"1px solid #3b82f6",fontWeight:700,
            color:{admin:"#34d399",director:"#a78bfa",accountant:"#fbbf24",manager:"#60a5fa",worker:"#94a3b8",logistics:"#f97316"}[f.userRole||"worker"]}}>
          <option value="admin">🔑 Admin — sees everything</option>
          <option value="director">👔 Director — sees everything</option>
          <option value="accountant">📊 Accountant — sees everything</option>
          <option value="manager">🏗 Manager — sees their sites only</option>
          <option value="worker">👷 Worker — sees their own data only</option>
          <option value="logistics">🚛 Logistics — asset management</option>
        </select>
        <div style={{fontSize:10,color:"#64748b",marginTop:5}}>This determines what the user sees when they log into the admin app. Workers also use the separate Worker Portal app.</div>
      </div>
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
  if(!win){const a=document.createElement("a");a.href=u;a.download="TrainingMatrix_"+label.split(" ").join("_")+".html";a.click();}
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
    XLSX.writeFile(wb,"WeeklyRecord_WC_"+rec.weekLabel.split(" ").join("_")+".xlsx");
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
  sidebar:{width:210,minWidth:210,background:"#0a0e17",borderRight:"1px solid #1e2535",height:"100%",flexShrink:0,overflowY:"auto",display:"flex",flexDirection:"column"},
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
  {id:"expenses",       icon:"💸", label:"Expenses",           group:"analysis"},
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
function DashSidebar({page,setPage,workers,allSites,clients,invoices,bankTransactions,setModal,activeDays,siteHours,weekLabel,role,setAuthState}){
  const expiring=workers.flatMap(w=>Object.values(w.certs||{}).filter(c=>{if(!c.held||!c.expiry)return false;const d=(new Date(c.expiry)-new Date())/86400000;return d>=0&&d<30;})).length;
  const expenseCount=(bankTransactions||[]).filter(t=>t.type==="expense").length;
  const badges={certs:expiring,invoices:invoices.filter(i=>i.status==="pending").length,expenses:expenseCount};
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
      <button onClick={()=>doExcel(workers,weekLabel,activeDays,siteHours,clients,allSites)} style={{width:"100%",padding:"7px 10px",background:"linear-gradient(135deg,#059669,#10b981)",border:"none",borderRadius:7,color:"#fff",cursor:"pointer",fontSize:12,fontWeight:700,marginBottom:8}}>⬇ Export Excel</button>
        <div style={{borderTop:"1px solid #1e2535",paddingTop:8}}>
        <div style={{fontSize:9,color:"#374151",fontWeight:700,textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:6,paddingLeft:2}}>Integrated Tools</div>
        <button onClick={()=>setPage("app_sitemanager")}
          style={{width:"100%",padding:"9px 10px",background:page==="app_sitemanager"?"#0c1a2e":"#080f1a",border:`1px solid ${page==="app_sitemanager"?"#3b82f6":"#1e3a5f"}`,borderRadius:8,color:page==="app_sitemanager"?"#93c5fd":"#3b82f6",cursor:"pointer",fontSize:12,fontWeight:700,marginBottom:5,display:"flex",alignItems:"center",gap:8,transition:"all 0.15s"}}>
          <span style={{fontSize:17}}>📊</span><span>Site Manager</span>
        </button>
        <button onClick={()=>setPage("app_sitedocs")}
          style={{width:"100%",padding:"9px 10px",background:page==="app_sitedocs"?"#0d0a1e":"#0a0814",border:`1px solid ${page==="app_sitedocs"?"#7c3aed":"#2d1f4e"}`,borderRadius:8,color:page==="app_sitedocs"?"#c4b5fd":"#7c3aed",cursor:"pointer",fontSize:12,fontWeight:700,marginBottom:5,display:"flex",alignItems:"center",gap:8,transition:"all 0.15s"}}>
          <span style={{fontSize:17}}>📋</span><span>Site Documents</span>
        </button>
        <button onClick={()=>setPage("app_assets")}
          style={{width:"100%",padding:"9px 10px",background:page==="app_assets"?"#1a1200":"#0d0900",border:`1px solid ${page==="app_assets"?"#ffb000":"#3a2a00"}`,borderRadius:8,color:page==="app_assets"?"#ffb000":"#a87000",cursor:"pointer",fontSize:12,fontWeight:700,display:"flex",alignItems:"center",gap:8,transition:"all 0.15s"}}>
          <span style={{fontSize:17}}>⚙</span><span>Asset Register</span>
        </button>
      </div>
      {/* Role + logout */}
      <div style={{borderTop:"1px solid #1e2535",paddingTop:10,marginTop:6}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:7}}>
          <span style={{fontSize:10,padding:"2px 8px",borderRadius:20,background:"#34d39922",color:"#34d399",border:"1px solid #34d39944",fontWeight:700,textTransform:"uppercase"}}>
            {role==="director"?"👔 Director":role==="accountant"?"📊 Accountant":"🔑 Admin"}
          </span>
        </div>
        <button onClick={()=>setAuthState(null)}
          style={{width:"100%",padding:"7px 10px",background:"#1e2535",border:"1px solid #374151",borderRadius:7,color:"#64748b",cursor:"pointer",fontSize:11,fontWeight:600,display:"flex",alignItems:"center",justifyContent:"center",gap:6}}>
          <span>🚪</span> Sign Out
        </button>
      </div>
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
// ─── Invite Worker Modal ──────────────────────────────────────────────────────
function InviteWorkerModal({onClose,allSites}){
  const [name,setName]=useState("");
  const [email,setEmail]=useState("");
  const [role,setRole]=useState("worker");
  const [site,setSite]=useState("");
  const [sent,setSent]=useState(false);
  const [link,setLink]=useState("");
  const [copying,setCopying]=useState(false);

  const APP_URL=window.location.origin||"https://labour-schedule1.netlify.app";
  const ROLE_LABELS={admin:"Admin — full access",director:"Director — full access",accountant:"Accountant — full access",manager:"Manager — their sites only",worker:"Worker — own data only",logistics:"Logistics — asset management"};
  const ROLE_C={admin:"#34d399",director:"#a78bfa",accountant:"#fbbf24",manager:"#60a5fa",worker:"#94a3b8",logistics:"#f97316"};

  const generateInvite=async()=>{
    if(!name.trim()||!email.trim()){alert("Name and email are required.");return;}
    // Create invite token
    const token=btoa(JSON.stringify({name:name.trim(),email:email.trim(),role,site,createdAt:Date.now()})).replace(/=/g,"");
    const inviteUrl=`${APP_URL}?invite=${token}`;
    setLink(inviteUrl);

    // Try to save invite to Supabase
    try{
      await sbUpsert("invites",[{id:"inv_"+Date.now(),data:{name:name.trim(),email:email.trim(),role,site,token,inviteUrl,createdAt:new Date().toISOString(),status:"pending"}}]);
    }catch(_){}

    // Try to send email via Supabase Edge Function
    try{
      const body=`Hi ${name.trim()},\n\nYou have been invited to join Bright Metalwork's construction management platform as: ${ROLE_LABELS[role]||role}.\n\nClick the link below to complete your registration:\n${inviteUrl}\n\nThe link is pre-filled with your details. You just need to set your password and fill in any remaining information.\n\nBright Metalwork Ltd\n+44 (0)771 078 3500\nlucian@bright-group.org`;
      await fetch(`${SB_URL}/functions/v1/send-email`,{
        method:"POST",headers:{...SB_H,"Content-Type":"application/json"},
        body:JSON.stringify({to:email.trim(),subject:`Invitation to Bright Metalwork Portal — ${name.trim()}`,text:body})
      });
    }catch(_){}

    setSent(true);
  };

  const copyLink=()=>{
    navigator.clipboard.writeText(link).then(()=>{setCopying(true);setTimeout(()=>setCopying(false),2000);});
  };

  const OL={width:"100%",background:"#0d1117",border:"1px solid #1e2535",borderRadius:8,padding:"9px 12px",color:"#f1f5f9",fontSize:13,outline:"none",boxSizing:"border-box",marginTop:5};

  return <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.7)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:200,padding:16}}>
    <div style={{background:"#111827",borderRadius:16,padding:24,width:"100%",maxWidth:480,border:"1px solid #1e2535",boxShadow:"0 24px 64px rgba(0,0,0,.6)"}}>
      {/* Header */}
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20}}>
        <div>
          <div style={{fontSize:16,fontWeight:800,color:"#f1f5f9"}}>✉ Invite New Worker</div>
          <div style={{fontSize:11,color:"#64748b",marginTop:2}}>An invitation link will be generated and optionally emailed</div>
        </div>
        <button onClick={onClose} style={{background:"none",border:"none",color:"#64748b",cursor:"pointer",fontSize:20,lineHeight:1}}>×</button>
      </div>

      {!sent?<div>
        {/* Name */}
        <div style={{marginBottom:12}}>
          <label style={{fontSize:10,color:"#64748b",fontWeight:700,textTransform:"uppercase",letterSpacing:"0.06em"}}>Full Name *</label>
          <input value={name} onChange={e=>setName(e.target.value)} placeholder="e.g. John Smith" style={OL}/>
        </div>
        {/* Email */}
        <div style={{marginBottom:12}}>
          <label style={{fontSize:10,color:"#64748b",fontWeight:700,textTransform:"uppercase",letterSpacing:"0.06em"}}>Email Address *</label>
          <input type="email" value={email} onChange={e=>setEmail(e.target.value)} placeholder="john@example.com" style={OL}/>
        </div>
        {/* Role */}
        <div style={{marginBottom:12}}>
          <label style={{fontSize:10,color:"#64748b",fontWeight:700,textTransform:"uppercase",letterSpacing:"0.06em"}}>Access Level</label>
          <select value={role} onChange={e=>setRole(e.target.value)}
            style={{...OL,color:ROLE_C[role]||"#f1f5f9",fontWeight:700,border:`1px solid ${ROLE_C[role]||"#1e2535"}44`,cursor:"pointer"}}>
            <option value="worker">👷 Worker — own data only</option>
            <option value="manager">🏗 Manager — their sites only</option>
            <option value="admin">🔑 Admin — full access</option>
            <option value="director">👔 Director — full access</option>
            <option value="accountant">📊 Accountant — full access</option>
            <option value="logistics">🚛 Logistics — asset management</option>
          </select>
        </div>
        {/* Site (for managers) */}
        {role==="manager"&&<div style={{marginBottom:12}}>
          <label style={{fontSize:10,color:"#64748b",fontWeight:700,textTransform:"uppercase",letterSpacing:"0.06em"}}>Assigned Site</label>
          <select value={site} onChange={e=>setSite(e.target.value)} style={{...OL,cursor:"pointer"}}>
            <option value="">— All sites (set later) —</option>
            {allSites.filter(s=>!s.builtin).map(s=><option key={s.id} value={s.name}>{s.name}</option>)}
          </select>
        </div>}
        {/* Info box */}
        <div style={{background:"#0d1117",borderRadius:8,padding:"10px 13px",marginBottom:16,fontSize:11,color:"#64748b",lineHeight:1.6}}>
          <div style={{fontWeight:700,color:ROLE_C[role],marginBottom:4}}>
            {role==="worker"?"👷 Worker"
             :role==="manager"?"🏗 Manager"
             :"🔑 Admin/Director"} access
          </div>
          {role==="worker"&&"They will see their own schedule, timesheets, payslips and certificates. Cannot edit or delete anything."}
          {role==="manager"&&"They will see the Site Manager app — their assigned sites, workers, expenses, variations and payment applications."}
          {["admin","director","accountant"].includes(role)&&"Full access to all features, all workers, financials and settings."}
        </div>
        <div style={{display:"flex",gap:9}}>
          <button onClick={onClose} style={{flex:1,padding:"10px",background:"#1e2535",border:"1px solid #2d3555",borderRadius:8,color:"#94a3b8",cursor:"pointer",fontSize:13,fontWeight:600}}>Cancel</button>
          <button onClick={generateInvite} style={{flex:2,padding:"10px",background:"linear-gradient(135deg,#1e3a5f,#3b82f6)",border:"none",borderRadius:8,color:"#fff",cursor:"pointer",fontSize:13,fontWeight:700}}>Generate Invitation Link</button>
        </div>
      </div>:<div>
        {/* Success state */}
        <div style={{background:"#0d2218",border:"1px solid #34d39944",borderRadius:10,padding:16,marginBottom:16}}>
          <div style={{fontSize:14,fontWeight:800,color:"#34d399",marginBottom:4}}>✓ Invitation created for {name}</div>
          <div style={{fontSize:11,color:"#64748b"}}>An email was attempted to {email}. Whether it arrived depends on your Supabase email setup. Share the link below directly to be sure.</div>
        </div>
        {/* Link */}
        <div style={{marginBottom:14}}>
          <div style={{fontSize:10,color:"#64748b",fontWeight:700,textTransform:"uppercase",marginBottom:5}}>Invitation Link</div>
          <div style={{background:"#0d1117",borderRadius:8,padding:"10px 12px",border:"1px solid #1e2535",fontSize:11,color:"#60a5fa",wordBreak:"break-all",marginBottom:8,fontFamily:"monospace"}}>{link}</div>
          <button onClick={copyLink}
            style={{width:"100%",padding:"9px",background:copying?"#0d2218":"#1e2535",border:`1px solid ${copying?"#34d399":"#2d3555"}`,borderRadius:7,color:copying?"#34d399":"#94a3b8",cursor:"pointer",fontSize:12,fontWeight:600,transition:"all .2s"}}>
            {copying?"✓ Copied!":"📋 Copy Link"}
          </button>
        </div>
        <div style={{fontSize:11,color:"#64748b",background:"#0d1117",borderRadius:8,padding:"10px 12px",border:"1px solid #1e2535",marginBottom:14,lineHeight:1.6}}>
          When {name} opens the link they will see a pre-filled registration form with their name, email and <b style={{color:ROLE_C[role]}}>{role}</b> access already set. They just need to fill in their details and set a password.
        </div>
        <div style={{display:"flex",gap:9}}>
          <button onClick={()=>{setSent(false);setName("");setEmail("");setRole("worker");setSite("");setLink("");}}
            style={{flex:1,padding:"9px",background:"#1e2535",border:"1px solid #2d3555",borderRadius:8,color:"#94a3b8",cursor:"pointer",fontSize:12,fontWeight:600}}>
            + Another Invite
          </button>
          <button onClick={onClose}
            style={{flex:1,padding:"9px",background:"linear-gradient(135deg,#1e3a5f,#3b82f6)",border:"none",borderRadius:8,color:"#fff",cursor:"pointer",fontSize:12,fontWeight:700}}>
            Done
          </button>
        </div>
      </div>}
    </div>
  </div>;
}

function DWorkers({workers,allSites,clients,activeDays,siteHours,setPage,setDetailId,setModal}){\
  const[search,setSearch]=useState("");
  const[showInvite,setShowInvite]=useState(false);
  const shown=workers.filter(w=>!search||w.name.toLowerCase().includes(search.toLowerCase())||w.position.toLowerCase().includes(search.toLowerCase()));
  const {gross}=useMemo(()=>workers.reduce((a,w)=>{const r=calcPay(w,activeDays,siteHours);return{gross:a.gross+r.gross};},{gross:0}),[workers,activeDays,siteHours]);

  return <div>
    <DPageHdr title="👷 Workers" sub={`${workers.length} operatives · £${gross.toFixed(0)} gross this week`}
      actions={<div style={{display:"flex",gap:7}}>
        <button onClick={()=>setShowInvite(true)} style={{padding:"7px 14px",background:"#0d1a2e",border:"1px solid #3b82f6",borderRadius:7,color:"#60a5fa",cursor:"pointer",fontSize:12,fontWeight:700}}>✉ Invite</button>
        <button onClick={()=>setModal({type:"worker",worker:mkW()})} style={{padding:"7px 14px",background:"linear-gradient(135deg,#3b82f6,#6366f1)",border:"none",borderRadius:7,color:"#fff",cursor:"pointer",fontSize:12,fontWeight:700}}>+ Add Worker</button>
      </div>}/>
    {showInvite&&<InviteWorkerModal onClose={()=>setShowInvite(false)} allSites={allSites}/>}
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
function DWorkerDetail({workers,allSites,clients,activeDays,siteHours,workerId,timesheetRecords,payslipRecords,setTimesheetRecords,setPayslipRecords,setPage,setModal,weekLabel}){
  const w=workers.find(x=>x.id===workerId);
  if(!w)return <div style={DS.body}><div style={{color:"#374151",textAlign:"center",padding:40}}>Worker not found.</div></div>;

  const [tab,setTab]=useState("profile");
  const [reportEntry,setReportEntry]=useState(null); // entry being edited for work report

  // ── Derived data ────────────────────────────────────────────────────────────
  const {gross,net,stdH,otH}=calcPay(w,activeDays,siteHours);
  const heldCerts=CERTS.filter(c=>w.certs?.[c.key]?.held);
  const certAlerts=heldCerts.filter(c=>{const s=cSt(c,w);return s==="expired"||s==="expiring";});
  const CERT_STATUS_C={valid:"#34d399",expiring:"#fbbf24",expired:"#f87171",missing:"#374151"};

  // RTW
  const rtwExp=w.shareCodeExpiry?new Date(w.shareCodeExpiry):null;
  const rtwDays=rtwExp?(rtwExp-new Date())/86400000:null;
  const rtwStatus=!w.shareCode?"missing":!rtwExp?"valid":rtwDays<0?"expired":rtwDays<30?"expiring":"valid";
  const rtwColor={missing:"#f87171",valid:"#34d399",expiring:"#fbbf24",expired:"#f87171"}[rtwStatus];
  const rtwUploaded=!!(w.workerFiles||[]).length;

  // Pay rate history
  const payHistory=w.payRateHistory||[];

  // Timesheets & payslips for this worker
  const wTimesheets=(timesheetRecords||[]).filter(t=>t.workerId===w.id)
    .sort((a,b)=>new Date(b.weekLabel)-new Date(a.weekLabel));
  const wPayslips=(payslipRecords||[]).filter(p=>p.workerId===w.id)
    .sort((a,b)=>new Date(b.weekLabel)-new Date(a.weekLabel));

  // GPS confirmed work entries
  const wLogs=(w.attendanceLogs||[]).filter(l=>l.signIn&&l.signOut)
    .sort((a,b)=>new Date(b.signIn)-new Date(a.signIn));

  // This week's confirmed entries
  const thisWeekLogs=wLogs.filter(l=>l.weekLabel===weekLabel);
  const confirmedHrsThisWeek=thisWeekLogs.reduce((a,l)=>a+Math.round(((new Date(l.signOut)-new Date(l.signIn))/3600000)*100)/100,0);

  const initials=(w.name||"?").split(" ").map(n=>n[0]).join("").slice(0,2).toUpperCase();
  const fmtD=d=>d?new Date(d).toLocaleDateString("en-GB",{day:"2-digit",month:"short",year:"numeric"}):"—";
  const fmtT=iso=>iso?new Date(iso).toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit"}):"—";

  // ── Action buttons ─────────────────────────────────────────────────────────
  const ActionButtons=()=><div style={{display:"flex",gap:6,flexWrap:"wrap",marginBottom:0}}>
    <button onClick={()=>setModal({type:"worker",worker:w})}
      style={{padding:"7px 13px",background:"#1e3a5f",border:"1px solid #3b82f6",borderRadius:7,color:"#60a5fa",cursor:"pointer",fontSize:12,fontWeight:700}}>
      ✏️ Edit
    </button>
    <button onClick={()=>setTab("payslips")}
      style={{padding:"7px 13px",background:tab==="payslips"?"#1a0d2e":"#0d1117",border:`1px solid ${tab==="payslips"?"#a78bfa":"#1e2535"}`,borderRadius:7,color:tab==="payslips"?"#a78bfa":"#94a3b8",cursor:"pointer",fontSize:12,fontWeight:600}}>
      💷 Payslips{wPayslips.length>0?` (${wPayslips.length})`:""}
    </button>
    <button onClick={()=>exportWorkerProfile(w,allSites,weekLabel||formatWeekLabel(new Date()))}
      style={{padding:"7px 13px",background:"#0d1117",border:"1px solid #1e2535",borderRadius:7,color:"#64748b",cursor:"pointer",fontSize:12}}>
      📋 PDF
    </button>
  </div>;

  const handleSaveReport=(updatedEntry)=>{
    // Update the log in worker data
    const updatedLogs=(w.attendanceLogs||[]).map(l=>l.id===updatedEntry.id?updatedEntry:l);
    // Trigger save via setModal pattern — update worker attendanceLogs
    if(setModal) setModal({type:"_updateWorkerLogs",workerId:w.id,logs:updatedLogs});
    setReportEntry(null);
  };

  return <div>
    {reportEntry&&<WorkReportModal entry={reportEntry} worker={w} onSave={handleSaveReport} onClose={()=>setReportEntry(null)}/>}
    <DPageHdr
      title={<span style={{display:"flex",alignItems:"center",gap:10}}>
        <div style={{width:38,height:38,borderRadius:9,background:"#3b82f622",border:"1px solid #3b82f644",display:"flex",alignItems:"center",justifyContent:"center",fontSize:15,fontWeight:800,color:"#3b82f6"}}>{initials}</div>
        <div>
          <div>{w.name}</div>
          <div style={{fontSize:11,color:"#64748b",fontWeight:400,marginTop:1}}>{w.position} · {w.company}</div>
        </div>
      </span>}
      sub="" back="Workers" onBack={()=>setPage("workers")}
      actions={<ActionButtons/>}
    />

    <div style={DS.body}>
      {/* ── KPI STRIP — 4 tiles (no Certs Held) ── */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:8,marginBottom:16}}>
        <DStat label="Hourly Rate" value={w.agreedRate?"£"+w.agreedRate+"/hr":"Not set"} color="#34d399"/>
        <DStat label="CIS Tax" value={Math.round((w.taxRate||0)*100)+"%"} color={w.taxRate===0.30?"#f87171":w.taxRate===0.20?"#fbbf24":"#34d399"}/>
        <DStat label="Wk Confirmed" value={confirmedHrsThisWeek>0?confirmedHrsThisWeek.toFixed(1)+"h":"—"} color="#60a5fa" sub="GPS sign-ins"/>
        <DStat label="Right to Work" value={rtwStatus.toUpperCase()} color={rtwColor} sub={rtwDays!==null&&rtwDays<30?rtwDays<0?"EXPIRED":Math.ceil(rtwDays)+"d left":""}/>
      </div>

      {/* ── TAB NAV ── */}
      <div style={{display:"flex",gap:3,background:"#0d1117",borderRadius:8,padding:3,marginBottom:16,overflowX:"auto",flexWrap:"wrap"}}>
        {[
          ["profile","👤 Profile"],
          ["pay","💷 Pay Rate"],
          ["workentries","✅ Work Entries"],
          ["timesheets","⏱ Timesheets"+(wTimesheets.length?` (${wTimesheets.length})`:"")],
          ["payslips","💰 Payslips"+(wPayslips.length?` (${wPayslips.length})`:"")],
          ["certs","🏅 Certs"+(certAlerts.length?` ⚠${certAlerts.length}`:`(${heldCerts.length})`)],
          ["rtw","🛡 RTW"],
          ["schedule","📅 Schedule"],
          ["holidays","🏖 Holidays"],
        ].map(([v,l])=>(
          <button key={v} onClick={()=>setTab(v)}
            style={{padding:"6px 11px",background:tab===v?"#1e3a5f":"transparent",
              border:tab===v?"1px solid #3b82f6":"1px solid transparent",
              borderRadius:6,color:tab===v?"#60a5fa":"#64748b",cursor:"pointer",
              fontSize:11,fontWeight:tab===v?700:400,whiteSpace:"nowrap"}}>
            {l}
          </button>
        ))}
      </div>

      {/* ══ PROFILE TAB ══════════════════════════════════════════════════════ */}
      {tab==="profile"&&<div style={{display:"flex",flexDirection:"column",gap:14}}>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
          {/* Left — personal */}
          <div style={{background:"#111827",border:"1px solid #1e2535",borderRadius:11,padding:16}}>
            <div style={{fontSize:11,color:"#60a5fa",fontWeight:700,textTransform:"uppercase",marginBottom:12}}>Personal Details</div>
            {[["Full Name",w.name],["Position",w.position],["Company",w.company],["Nationality",w.nationality||"—"],["Date of Birth",w.dob?fmtD(w.dob):"—"],["Contact",w.contact||"—"],["Email",w.email||"—"],["Address",w.address||"—"],["Next of Kin",w.nextOfKin?(w.nextOfKin+(w.nextOfKinPhone?" · "+w.nextOfKinPhone:"")):"—"]].map(([l,v])=>(
              <div key={l} style={{display:"flex",gap:8,padding:"6px 0",borderBottom:"1px solid #1e2535"}}>
                <span style={{fontSize:10,color:"#64748b",fontWeight:700,minWidth:90,textTransform:"uppercase",flexShrink:0}}>{l}</span>
                <span style={{fontSize:12,color:"#e2e8f0",wordBreak:"break-word"}}>{v||"—"}</span>
              </div>
            ))}
          </div>
          <div style={{display:"flex",flexDirection:"column",gap:12}}>
            <div style={{background:"#111827",border:"1px solid #1e2535",borderRadius:11,padding:16}}>
              <div style={{fontSize:11,color:"#fbbf24",fontWeight:700,textTransform:"uppercase",marginBottom:12}}>Pay Details</div>
              {[["Hourly Rate",w.agreedRate?"£"+w.agreedRate+"/hr":"Not set"],["CIS Tax",Math.round((w.taxRate||0)*100)+"%"],["OT Rate",w.customOTRate?"£"+w.customOTRate+"/hr":"×"+(w.overtimeMultiplier||1.5)+" standard"],["NINO",w.nino||"—"],["UTR",w.utr||"—"],["Bank",w.bankName?(w.bankName+" · "+(w.bankAccount||"")+" · "+(w.bankSort||"")):"—"]].map(([l,v])=>(
                <div key={l} style={{display:"flex",gap:8,padding:"5px 0",borderBottom:"1px solid #1e2535"}}>
                  <span style={{fontSize:10,color:"#64748b",fontWeight:700,minWidth:80,textTransform:"uppercase",flexShrink:0}}>{l}</span>
                  <span style={{fontSize:12,color:"#e2e8f0"}}>{v}</span>
                </div>
              ))}
            </div>
            <div style={{background:"#111827",border:"1px solid #1e2535",borderRadius:11,padding:16}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
                <div style={{fontSize:11,color:"#a78bfa",fontWeight:700,textTransform:"uppercase"}}>🏅 Certifications</div>
                <button onClick={()=>setTab("certs")} style={{fontSize:10,color:"#3b82f6",background:"none",border:"none",cursor:"pointer",fontWeight:600}}>View all →</button>
              </div>
              {heldCerts.length===0
                ?<div style={{fontSize:12,color:"#374151",fontStyle:"italic"}}>No certifications recorded.</div>
                :<div style={{display:"flex",flexDirection:"column",gap:4}}>
                  {heldCerts.map(cert=>{const s=cSt(cert,w);const val=w.certs[cert.key];
                    return <div key={cert.key} style={{display:"flex",alignItems:"center",gap:8,padding:"4px 7px",background:"#0f1421",borderRadius:6,border:`1px solid ${CERT_STATUS_C[s]}33`}}>
                      <div style={{width:6,height:6,borderRadius:"50%",background:CERT_STATUS_C[s],flexShrink:0}}/>
                      <span style={{fontSize:11,color:"#e2e8f0",flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{cert.label}</span>
                      {cert.hasExpiry&&val?.expiry&&<span style={{fontSize:10,color:CERT_STATUS_C[s],fontWeight:600,flexShrink:0}}>{fmtD(val.expiry)}</span>}
                      <span style={{fontSize:9,color:CERT_STATUS_C[s],fontWeight:700,textTransform:"uppercase",flexShrink:0}}>{s}</span>
                    </div>;})}
                </div>}
            </div>
          </div>
        </div>
        {/* ── Bottom: Week Forecast — full width ── */}
        <div style={{background:"#111827",border:"1px solid #1e2535",borderRadius:11,padding:16}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
            <div style={{fontSize:11,color:"#3b82f6",fontWeight:700,textTransform:"uppercase"}}>📅 Week Forecast — WC {weekLabel||"—"}</div>
            <span style={{fontSize:10,color:"#374151",fontStyle:"italic"}}>Route only · ✅ = GPS confirmed</span>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:6}}>
            {["Mon","Tue","Wed","Thu","Fri","Sat","Sun"].map(d=>{
              const site=w.days?.[d];const off=!site||isOff(site);
              const confirmed=thisWeekLogs.some(l=>l.day===d);
              const activeNow=thisWeekLogs.some(l=>l.day===d&&l.signIn&&!l.signOut);
              const siteObj=allSites.find(s=>site&&(site.includes(s.name)||s.name.includes((site||"").split("-")[0]?.trim())));
              const col=siteObj?.color||"#3b82f6";
              const isToday=["Sun","Mon","Tue","Wed","Thu","Fri","Sat"][new Date().getDay()]===d;
              return <div key={d} style={{textAlign:"center",background:confirmed?"#0d2218":activeNow?"#0d1e2a":isToday?"#1a1f2e":"#0f1421",borderRadius:9,padding:"9px 4px",border:`2px solid ${activeNow?"#fbbf24":confirmed?"#34d399":isToday?col+"88":"#1e2535"}`,transition:"all 0.2s"}}>
                <div style={{fontSize:9,fontWeight:700,textTransform:"uppercase",marginBottom:4,color:activeNow?"#fbbf24":confirmed?"#34d399":isToday?col:"#64748b"}}>{d}</div>
                {activeNow&&<div style={{fontSize:9,color:"#fbbf24",marginBottom:2}}>● Live</div>}
                {!activeNow&&confirmed&&<div style={{fontSize:11,marginBottom:2}}>✅</div>}
                <div style={{fontSize:9,color:off?"#374151":col,fontWeight:off?400:600,lineHeight:1.3,wordBreak:"break-word",minHeight:24}}>{off?"Off":(site||"").trim()}</div>
              </div>;})}
          </div>
          <div style={{display:"flex",gap:16,marginTop:8,fontSize:10,color:"#374151"}}>
            <span>✅ GPS confirmed</span><span style={{color:"#fbbf24"}}>● On site now</span><span>📋 Forecast only</span>
          </div>
        </div>
      </div>}

      {/* ══ PAY RATE TAB ═════════════════════════════════════════════════════ */}
      {tab==="pay"&&<div>
        {/* Current rates */}
        <div style={{background:"#111827",border:"1px solid #1e2535",borderRadius:11,padding:16,marginBottom:14}}>
          <div style={{fontSize:11,color:"#fbbf24",fontWeight:700,textTransform:"uppercase",marginBottom:12}}>Current Pay Rates</div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10,marginBottom:14}}>
            {[["Hourly Rate",w.agreedRate?"£"+w.agreedRate+"/hr":"Not set","#34d399"],["CIS Tax Rate",Math.round((w.taxRate||0)*100)+"%",w.taxRate===0.30?"#f87171":w.taxRate===0.20?"#fbbf24":"#34d399"],["OT Multiplier","×"+(w.overtimeMultiplier||1.5),"#fbbf24"],["Custom OT Rate",w.customOTRate?"£"+w.customOTRate+"/hr":"—","#f97316"]].map(([l,v,c])=>(
              <div key={l} style={{background:"#0d1117",borderRadius:9,padding:"10px 13px",border:`1px solid ${c}22`}}>
                <div style={{fontSize:9,color:"#64748b",fontWeight:700,textTransform:"uppercase",marginBottom:3}}>{l}</div>
                <div style={{fontSize:16,fontWeight:800,color:c}}>{v}</div>
              </div>
            ))}
          </div>
          <div style={{display:"flex",gap:8,padding:"8px 0",borderTop:"1px solid #1e2535",fontSize:11,color:"#64748b"}}>
            {w.payRateValidFrom&&<span>Valid from: <b style={{color:"#94a3b8"}}>{fmtD(w.payRateValidFrom)}</b></span>}
            {w.payRateValidTo&&<span>· Until: <b style={{color:"#94a3b8"}}>{fmtD(w.payRateValidTo)}</b></span>}
            <span style={{marginLeft:"auto",color:"#374151"}}>Edit rates via the ✏️ Edit button above</span>
          </div>
        </div>

        {/* Pay rate history */}
        <div style={{background:"#111827",border:"1px solid #1e2535",borderRadius:11,padding:16}}>
          <div style={{fontSize:11,color:"#64748b",fontWeight:700,textTransform:"uppercase",marginBottom:12}}>
            Pay Rate History — {payHistory.length} change{payHistory.length!==1?"s":""}
          </div>
          {payHistory.length===0
            ?<div style={{fontSize:12,color:"#374151",fontStyle:"italic"}}>No rate changes recorded yet. When you update pay rates, previous values are automatically saved here with a timestamp.</div>
            :<div style={{display:"flex",flexDirection:"column",gap:6}}>
              {[...payHistory].reverse().map((h,i)=>(
                <div key={i} style={{background:"#0f1421",borderRadius:8,padding:"10px 13px",border:"1px solid #1e2535"}}>
                  <div style={{display:"flex",justifyContent:"space-between",marginBottom:6}}>
                    <div style={{fontSize:10,color:"#64748b",fontWeight:700,textTransform:"uppercase"}}>Changed on {fmtD(h.changedAt)}</div>
                    {h.changedBy&&<div style={{fontSize:10,color:"#64748b"}}>by {h.changedBy}</div>}
                  </div>
                  <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:8}}>
                    {h.agreedRate!==undefined&&<div><span style={{fontSize:10,color:"#64748b"}}>Rate: </span><span style={{fontSize:12,fontWeight:700,color:"#94a3b8",textDecoration:"line-through"}}>£{h.agreedRate}/hr</span></div>}
                    {h.taxRate!==undefined&&<div><span style={{fontSize:10,color:"#64748b"}}>Tax: </span><span style={{fontSize:12,fontWeight:700,color:"#94a3b8",textDecoration:"line-through"}}>{Math.round(h.taxRate*100)}%</span></div>}
                    {h.overtimeMultiplier!==undefined&&<div><span style={{fontSize:10,color:"#64748b"}}>OT: </span><span style={{fontSize:12,fontWeight:700,color:"#94a3b8",textDecoration:"line-through"}}>×{h.overtimeMultiplier}</span></div>}
                  </div>
                  {h.note&&<div style={{fontSize:11,color:"#64748b",marginTop:5,fontStyle:"italic"}}>{h.note}</div>}
                </div>
              ))}
            </div>}
        </div>
      </div>}

      {/* ══ WORK ENTRIES TAB ════════════════════════════════════════════════ */}
      {tab==="workentries"&&<div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:8,marginBottom:14}}>
          <DStat label="Total Entries" value={wLogs.length} color="#60a5fa"/>
          <DStat label="This Week" value={thisWeekLogs.length} color="#34d399"/>
          <DStat label="Conf. Hours Wk" value={confirmedHrsThisWeek>0?confirmedHrsThisWeek.toFixed(1)+"h":"—"} color="#fbbf24"/>
          <DStat label="Sites Visited" value={[...new Set(wLogs.map(l=>l.siteName))].length} color="#a78bfa"/>
        </div>
        {wLogs.length===0
          ?<div style={{textAlign:"center",padding:40,color:"#374151",fontSize:13}}>
              <div style={{fontSize:32,marginBottom:10}}>📍</div>
              No GPS work entries yet. Entries are created automatically when the worker signs in via the Worker Portal.
            </div>
          :<div style={{border:"1px solid #1e2535",borderRadius:10,overflow:"hidden"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
              <thead><tr>
                <th style={{...DS.th,textAlign:"left"}}>Date</th>
                <th style={{...DS.th,textAlign:"left"}}>Site</th>
                <th style={{...DS.th,textAlign:"center"}}>Sign In</th>
                <th style={{...DS.th,textAlign:"center"}}>Sign Out</th>
                <th style={{...DS.th,textAlign:"right"}}>Normal</th>
                <th style={{...DS.th,textAlign:"right"}}>OT</th>
                <th style={{...DS.th,textAlign:"center"}}>Report</th>
                <th style={{...DS.th,textAlign:"center"}}></th>
              </tr></thead>
              <tbody>
                {wLogs.map((l,i)=>{
                  const hrs=Math.round(((new Date(l.signOut)-new Date(l.signIn))/3600000)*100)/100;
                  const site=allSites.find(s=>s.id===l.siteId||s.name===l.siteName);
                  const otThreshold=site?.otThreshold||site?.stdHours||9;
                  const stdHrs=Math.min(hrs,otThreshold);
                  const otHrs=Math.max(0,Math.round((hrs-otThreshold)*100)/100);
                  const isThisWeek=l.weekLabel===weekLabel;
                  return <tr key={l.id} style={{background:isThisWeek?"#0d1e2a":i%2===0?"#111827":"#0f1421"}}>
                    <td style={{...DS.td,color:"#94a3b8",fontSize:11}}>{fmtD(l.signIn?.split("T")[0])}</td>
                    <td style={{...DS.td,fontWeight:600,color:"#f1f5f9"}}>{l.siteName||"—"}</td>
                    <td style={{...DS.td,textAlign:"center",color:"#34d399",fontWeight:600}}>{fmtT(l.signIn)}</td>
                    <td style={{...DS.td,textAlign:"center",color:"#f87171",fontWeight:600}}>{l.signOut?fmtT(l.signOut):<span style={{color:"#fbbf24"}}>On site</span>}</td>
                    <td style={{...DS.td,textAlign:"right",color:"#60a5fa",fontWeight:700}}>{hrs.toFixed(2)}h</td>
                    <td style={{...DS.td,textAlign:"right",color:otHrs>0?"#fbbf24":"#374151",fontWeight:otHrs>0?700:400}}>{otHrs>0?"+"+otHrs.toFixed(1)+"h":"—"}</td>
                    <td style={{...DS.td,textAlign:"center",fontSize:13}}>
                      {l.workReport?.type==="positive"?<span style={{color:"#34d399",fontWeight:700}}>✓</span>:l.workReport?.type==="negative"?<span style={{color:"#f87171",fontWeight:700}}>⚠</span>:<span style={{color:"#374151"}}>·</span>}
                    </td>
                    <td style={{...DS.td,textAlign:"center"}}>
                      <button onClick={e=>{e.stopPropagation();setReportEntry(l);}}
                        style={{padding:"3px 9px",fontSize:10,fontWeight:700,borderRadius:5,border:`1px solid ${l.workReport?"#3b82f6":"#1e2535"}`,background:l.workReport?"#1e3a5f":"#0d1117",color:l.workReport?"#60a5fa":"#64748b",cursor:"pointer",whiteSpace:"nowrap"}}>
                        {l.workReport?"Edit":"+ Report"}
                      </button>
                    </td>
                  </tr>;
                })}
              </tbody>
            </table>
          </div>}
      </div>}

      {/* ══ TIMESHEETS TAB ════════════════════════════════════════════════════ */}
      {tab==="timesheets"&&<div>
        <div style={{fontSize:11,color:"#64748b",marginBottom:12,lineHeight:1.6}}>
          Timesheets are created automatically from GPS sign-ins. Each sign-in/sign-out creates a work entry. Normal and overtime hours are calculated per the site's threshold.
        </div>
        {wTimesheets.length===0
          ?<div style={{textAlign:"center",padding:40,color:"#374151",border:"1px dashed #1e2535",borderRadius:11}}>
              <div style={{fontSize:32,marginBottom:10}}>⏱</div>
              <div style={{fontWeight:600,marginBottom:4}}>No timesheets yet</div>
              <div style={{fontSize:11}}>Auto-created on first GPS sign-in each week.</div>
            </div>
          :wTimesheets.map(ts=>{
            const tsLogs=(ts.workEntries||wLogs.filter(l=>l.weekLabel===ts.weekLabel));
            const DAYS=["Mon","Tue","Wed","Thu","Fri","Sat","Sun"];
            const dayMap={};
            tsLogs.forEach(l=>{
              const hrs=Math.round(((new Date(l.signOut)-new Date(l.signIn))/3600000)*100)/100;
              const site=allSites.find(s=>s.id===l.siteId||s.name===l.siteName);
              const otThr=site?.otThreshold||site?.stdHours||9;
              if(!dayMap[l.day]) dayMap[l.day]={total:0,std:0,ot:0};
              dayMap[l.day].total=Math.round((dayMap[l.day].total+hrs)*100)/100;
              dayMap[l.day].std=Math.min(dayMap[l.day].total,otThr);
              dayMap[l.day].ot=Math.max(0,Math.round((dayMap[l.day].total-otThr)*100)/100);
            });
            const totalHrs=Object.values(dayMap).reduce((a,d)=>a+d.total,0);
            const stdHrs=ts.stdHours||Object.values(dayMap).reduce((a,d)=>a+d.std,0);
            const otHrs=ts.otHours||Object.values(dayMap).reduce((a,d)=>a+d.ot,0);
            const statusColor={open:"#3b82f6",pending:"#fbbf24",approved:"#34d399"}[ts.status]||"#64748b";
            return <div key={ts.id} style={{background:"#111827",border:`1px solid ${statusColor}44`,borderRadius:11,padding:16,marginBottom:14}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:12}}>
                <div>
                  <div style={{fontSize:14,fontWeight:800,color:"#f1f5f9"}}>WC {ts.weekLabel}</div>
                  <div style={{fontSize:11,color:"#64748b",marginTop:2}}>{w.name} · {tsLogs.length} work entr{tsLogs.length===1?"y":"ies"}</div>
                </div>
                <span style={{padding:"3px 10px",borderRadius:20,fontSize:11,fontWeight:700,background:statusColor+"22",color:statusColor,border:`1px solid ${statusColor}44`}}>
                  {{open:"● Open",pending:"⏳ Pending",approved:"✓ Approved"}[ts.status]||ts.status}
                </span>
              </div>
              {/* Work entries table */}
              <div style={{border:"1px solid #1e2535",borderRadius:8,overflow:"hidden",marginBottom:12}}>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 80px 80px 70px 60px 60px",background:"#0d1117",padding:"6px 10px",fontSize:9,color:"#64748b",fontWeight:700,textTransform:"uppercase",gap:6}}>
                  <span>Date</span><span>Site</span><span>Sign In</span><span>Sign Out</span><span style={{textAlign:"right"}}>Hours</span><span style={{textAlign:"right"}}>Normal</span><span style={{textAlign:"right"}}>OT</span>
                </div>
                {tsLogs.map((l,i)=>{
                  const hrs=Math.round(((new Date(l.signOut)-new Date(l.signIn))/3600000)*100)/100;
                  const site=allSites.find(s=>s.id===l.siteId||s.name===l.siteName);
                  const otThr=site?.otThreshold||9;
                  const std=Math.min(hrs,otThr);const ot=Math.max(0,Math.round((hrs-otThr)*100)/100);
                  const dateStr=l.signIn?new Date(l.signIn).toLocaleDateString("en-GB",{day:"2-digit",month:"short",weekday:"short"}):"—";
                  const tIn=l.signIn?new Date(l.signIn).toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit"}):"—";
                  const tOut=l.signOut?new Date(l.signOut).toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit"}):"—";
                  return <div key={l.id||i} style={{display:"grid",gridTemplateColumns:"1fr 1fr 80px 80px 70px 60px 60px",padding:"8px 10px",borderTop:"1px solid #1e2535",background:i%2===0?"#111827":"#0f1421",fontSize:11,gap:6,alignItems:"center"}}>
                    <span style={{color:"#94a3b8"}}>{dateStr}</span>
                    <span style={{color:"#f1f5f9",fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{l.siteName||"—"}</span>
                    <span style={{color:"#34d399",fontWeight:600}}>{tIn}</span>
                    <span style={{color:"#f87171",fontWeight:600}}>{tOut}</span>
                    <span style={{textAlign:"right",color:"#60a5fa",fontWeight:700}}>{hrs.toFixed(2)}h</span>
                    <span style={{textAlign:"right",color:"#60a5fa"}}>{std.toFixed(2)}h</span>
                    <span style={{textAlign:"right",color:ot>0?"#fbbf24":"#374151",fontWeight:ot>0?700:400}}>{ot>0?"+"+ot.toFixed(2):"—"}</span>
                  </div>;
                })}
              </div>
              {/* Day grid */}
              <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:5,marginBottom:12}}>
                {DAYS.map(d=><div key={d} style={{textAlign:"center",background:"#0f1421",borderRadius:6,padding:"6px 3px",border:`1px solid ${dayMap[d]?"#3b82f633":"#1e2535"}`}}>
                  <div style={{fontSize:9,color:"#64748b",fontWeight:700,textTransform:"uppercase",marginBottom:2}}>{d}</div>
                  {dayMap[d]?<>
                    <div style={{fontSize:11,fontWeight:700,color:"#60a5fa"}}>{dayMap[d].total.toFixed(1)}h</div>
                    {dayMap[d].ot>0&&<div style={{fontSize:9,color:"#fbbf24"}}>+{dayMap[d].ot.toFixed(1)}OT</div>}
                  </>:<div style={{fontSize:11,color:"#374151"}}>—</div>}
                </div>)}
              </div>
              {/* Totals */}
              <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:8,marginBottom:10}}>
                {[["Total Hours",totalHrs.toFixed(2)+"h","#f1f5f9"],["Normal Hours",stdHrs.toFixed(2)+"h","#60a5fa"],["Overtime Hours",otHrs.toFixed(2)+"h","#fbbf24"],["Source",ts.source==="gps"?"📍 GPS":"📋 Schedule","#64748b"]].map(([l,v,c])=>(
                  <div key={l} style={{background:"#0d1117",borderRadius:7,padding:"8px 10px"}}>
                    <div style={{fontSize:9,color:"#64748b",fontWeight:700,textTransform:"uppercase",marginBottom:2}}>{l}</div>
                    <div style={{fontSize:13,fontWeight:700,color:c}}>{v}</div>
                  </div>
                ))}
              </div>
              <div style={{fontSize:10,color:"#374151",fontStyle:"italic",marginBottom:10}}>Pay amounts are not calculated here — see the Payslips tab</div>
              <div style={{display:"flex",gap:8,marginTop:2}}>
                <button onClick={()=>exportTimesheetPDF(ts,w)}
                  style={{flex:1,padding:"9px",background:"#0d1117",border:"1px solid #1e2535",borderRadius:8,color:"#64748b",cursor:"pointer",fontSize:12,fontWeight:600,display:"flex",alignItems:"center",justifyContent:"center",gap:6}}>
                  📄 Export PDF
                </button>
                {ts.status!=="approved"&&<button
                  onClick={()=>{
                    // 1 — Approve timesheet
                    const updated=timesheetRecords.map(t=>t.id===ts.id?{...t,status:"approved",approvedAt:new Date().toISOString(),approvedBy:"Admin"}:t);
                    setTimesheetRecords(updated);
                    // 2 — Auto-generate payslip record
                    const otR=w.customOTRate||Math.round((w.agreedRate||0)*(w.overtimeMultiplier||1.5)*100)/100;
                    const stdAmt=Math.round((ts.stdHours||0)*(w.agreedRate||0)*100)/100;
                    const otAmt=Math.round((ts.otHours||0)*otR*100)/100;
                    const gross=Math.round((stdAmt+otAmt)*100)/100;
                    const taxAmt=Math.round(gross*(w.taxRate||0)*100)/100;
                    const net=Math.round((gross-taxAmt)*100)/100;
                    const newPS={
                      id:"ps_"+w.id+"_"+(ts.weekLabel||"").replace(/\s+/g,""),
                      workerId:w.id,workerName:w.name,position:w.position,
                      weekLabel:ts.weekLabel,
                      stdH:ts.stdHours||0,otH:ts.otHours||0,
                      rate:w.agreedRate||0,otRate:otR,taxRate:w.taxRate||0,
                      gross,taxAmt,net,
                      createdAt:new Date().toISOString(),source:"auto",
                    };
                    setPayslipRecords(prev=>{
                      const i=prev.findIndex(p=>p.id===newPS.id);
                      return i>=0?prev.map(p=>p.id===newPS.id?newPS:p):[...prev,newPS];
                    });
                    // 3 — Send email via Supabase Edge Function
                    if(w.email){
                      try{fetch(`${SB_URL}/functions/v1/send-email`,{method:"POST",headers:{...SB_H,"Content-Type":"application/json"},body:JSON.stringify({to:w.email,subject:"Bright Metalwork — Payslip WC "+ts.weekLabel,text:"Hi "+w.name+",\n\nYour payslip for the week commencing "+ts.weekLabel+" has been issued.\nNet pay: £"+net.toFixed(2)+"\n\nPlease log in to the worker portal to view and download your payslip.\n\nBright Metalwork Ltd\nlucian@bright-group.org\n+44 (0)771 078 3500"})});}catch(_){}
                      alert("✓ Timesheet approved. Payslip generated and sent to "+w.email+".");
                    }
                  }}
                  style={{flex:2,padding:"9px",background:"linear-gradient(135deg,#14532d,#16a34a)",border:"none",borderRadius:8,color:"#fff",cursor:"pointer",fontSize:12,fontWeight:700}}>
                  ✓ Approve — Generate &amp; Send Payslip
                </button>}
                {ts.status==="approved"&&<div style={{flex:2,padding:"9px",background:"#0d2218",border:"1px solid #34d39944",borderRadius:8,color:"#34d399",fontSize:12,fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center",gap:6}}>
                  ✓ Approved · Payslip sent
                </div>}
              </div>
            </div>;
          })}
      </div>}

      {/* ══ PAYSLIPS TAB ══════════════════════════════════════════════════════ */}
      {tab==="payslips"&&<div>
        <div style={{fontSize:11,color:"#64748b",marginBottom:12,display:"flex",alignItems:"center",gap:8}}>
          <span style={{fontSize:14}}>📧</span>
          <span>Payslip PDF is sent automatically to the worker's email when the timesheet is approved.</span>
        </div>
        {wPayslips.length===0
          ?<div style={{textAlign:"center",padding:40,color:"#374151",fontSize:13,border:"1px dashed #1e2535",borderRadius:11}}>
              <div style={{fontSize:32,marginBottom:10}}>💰</div>
              <div style={{fontWeight:600,marginBottom:4}}>No payslips yet</div>
              <div style={{fontSize:11}}>Auto-generated when timesheets are approved.</div>
            </div>
          :wPayslips.map(ps=>{
            const otRate=w.customOTRate||Math.round((w.agreedRate||0)*(w.overtimeMultiplier||1.5)*100)/100;
            const stdAmt=Math.round((ps.stdH||0)*(w.agreedRate||0)*100)/100;
            const otAmt=Math.round((ps.otH||0)*otRate*100)/100;
            const gross=Math.round((stdAmt+otAmt)*100)/100;
            const taxAmt=Math.round(gross*(w.taxRate||0)*100)/100;
            const net=Math.round((gross-taxAmt)*100)/100;
            const taxPct=Math.round((w.taxRate||0)*100);
            return <div key={ps.id} style={{background:"#111827",border:"1px solid #1e2535",borderRadius:11,overflow:"hidden",marginBottom:10}}>
              <div style={{padding:"12px 14px",borderBottom:"1px solid #1e2535",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                <div>
                  <div style={{fontSize:13,fontWeight:700,color:"#f1f5f9"}}>WC {ps.weekLabel}</div>
                  <div style={{fontSize:11,color:"#64748b",marginTop:2}}>{w.name} · {w.position}</div>
                </div>
                <div style={{fontSize:20,fontWeight:700,color:"#34d399"}}>£{net.toFixed(2)}</div>
              </div>
              <div style={{background:"#0d1117"}}>
                <div style={{display:"grid",gridTemplateColumns:"1.8fr 0.8fr 1fr 0.8fr",padding:"7px 14px",borderBottom:"1px solid #1e2535"}}>
                  {["Description","Hours","Rate","Amount"].map(h=><div key={h} style={{fontSize:9,fontWeight:700,color:"#64748b",textTransform:"uppercase",letterSpacing:".05em"}}>{h}</div>)}
                </div>
                <div style={{display:"grid",gridTemplateColumns:"1.8fr 0.8fr 1fr 0.8fr",padding:"9px 14px",borderBottom:"1px solid #1e2535",alignItems:"center"}}>
                  <div style={{fontSize:12,color:"#f1f5f9"}}>Normal hours</div>
                  <div style={{fontSize:12,color:"#94a3b8"}}>{(ps.stdH||0).toFixed(2)} h</div>
                  <div style={{fontSize:11,color:"#64748b"}}>£{(w.agreedRate||0).toFixed(2)}/hr</div>
                  <div style={{fontSize:12,fontWeight:600,color:"#f1f5f9",textAlign:"right"}}>£{stdAmt.toFixed(2)}</div>
                </div>
                {(ps.otH||0)>0&&<div style={{display:"grid",gridTemplateColumns:"1.8fr 0.8fr 1fr 0.8fr",padding:"9px 14px",borderBottom:"1px solid #1e2535",alignItems:"center"}}>
                  <div style={{fontSize:12,color:"#f1f5f9"}}>Overtime hours</div>
                  <div style={{fontSize:12,color:"#94a3b8"}}>{(ps.otH||0).toFixed(2)} h</div>
                  <div style={{fontSize:11,color:"#64748b"}}>£{otRate.toFixed(2)}/hr (×{w.overtimeMultiplier||1.5})</div>
                  <div style={{fontSize:12,fontWeight:700,color:"#fbbf24",textAlign:"right"}}>£{otAmt.toFixed(2)}</div>
                </div>}
                <div style={{display:"grid",gridTemplateColumns:"1.8fr 0.8fr 1fr 0.8fr",padding:"9px 14px",borderBottom:"1px solid #1e2535",alignItems:"center",background:"#111827"}}>
                  <div style={{fontSize:12,fontWeight:600,color:"#f1f5f9"}}>Gross pay</div>
                  <div></div><div></div>
                  <div style={{fontSize:14,fontWeight:700,color:"#f1f5f9",textAlign:"right"}}>£{gross.toFixed(2)}</div>
                </div>
                {taxPct>0&&<div style={{display:"grid",gridTemplateColumns:"1.8fr 0.8fr 1fr 0.8fr",padding:"9px 14px",borderBottom:"1px solid #1e2535",alignItems:"center"}}>
                  <div style={{fontSize:12,color:"#94a3b8"}}>CIS tax ({taxPct}%)</div>
                  <div></div><div></div>
                  <div style={{fontSize:12,fontWeight:700,color:"#f87171",textAlign:"right"}}>− £{taxAmt.toFixed(2)}</div>
                </div>}
                <div style={{display:"grid",gridTemplateColumns:"1.8fr 0.8fr 1fr 0.8fr",padding:"12px 14px",alignItems:"center"}}>
                  <div style={{fontSize:14,fontWeight:700,color:"#34d399"}}>Net pay</div>
                  <div></div><div></div>
                  <div style={{fontSize:22,fontWeight:700,color:"#34d399",textAlign:"right"}}>£{net.toFixed(2)}</div>
                </div>
              </div>
              <div style={{padding:"10px 14px",borderTop:"1px solid #1e2535",display:"flex",gap:8}}>
                <button onClick={()=>exportPayslipPDF(ps,w)}
                  style={{flex:1,padding:"8px",background:"#0d1117",border:"1px solid #2d3555",borderRadius:8,color:"#64748b",cursor:"pointer",fontSize:12,fontWeight:600,display:"flex",alignItems:"center",justifyContent:"center",gap:6}}>
                  📄 Export PDF
                </button>
                <button onClick={()=>alert("Payslip resent to "+w.email)}
                  style={{flex:1,padding:"8px",background:"#0d1117",border:"1px solid #3b82f644",borderRadius:8,color:"#60a5fa",cursor:"pointer",fontSize:12,fontWeight:600,display:"flex",alignItems:"center",justifyContent:"center",gap:6}}>
                  📧 Resend email
                </button>
              </div>
            </div>;
          })}
      </div>}

      {/* ══ CERTS TAB ════════════════════════════════════════════════════════ */}
      {tab==="certs"&&<div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:8,marginBottom:14}}>
          {[["Held",heldCerts.length,"#60a5fa"],["Valid",heldCerts.filter(c=>cSt(c,w)==="valid").length,"#34d399"],["Expiring Soon",heldCerts.filter(c=>cSt(c,w)==="expiring").length,"#fbbf24"],["Expired",heldCerts.filter(c=>cSt(c,w)==="expired").length,"#f87171"]].map(([l,v,c])=>(
            <DStat key={l} label={l} value={v} color={c}/>
          ))}
        </div>
        {/* Full cert list — ALL certs, held ones highlighted */}
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
          {CERTS.map(cert=>{
            const val=w.certs?.[cert.key];
            const held=val?.held;
            const s=cSt(cert,w);
            if(!held)return null;
            return <div key={cert.key} style={{background:"#0f1421",borderRadius:9,padding:"10px 12px",border:`1px solid ${CERT_STATUS_C[s]}44`}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:4}}>
                <span style={{fontSize:12,fontWeight:700,color:"#f1f5f9"}}>{cert.label}</span>
                <span style={{fontSize:10,fontWeight:700,color:CERT_STATUS_C[s],textTransform:"uppercase"}}>{s}</span>
              </div>
              {cert.hasExpiry&&<div style={{fontSize:10,color:"#64748b"}}>Expires: <span style={{color:CERT_STATUS_C[s],fontWeight:600}}>{val?.expiry?fmtD(val.expiry):"—"}</span></div>}
              {val?.photoUrl&&<img src={val.photoUrl} alt={cert.label} style={{marginTop:7,width:"100%",maxHeight:80,objectFit:"cover",borderRadius:5,border:"1px solid #1e2535",cursor:"pointer"}} onClick={()=>window.open(val.photoUrl,"_blank")}/>}
            </div>;
          }).filter(Boolean)}
        </div>
        {heldCerts.length===0&&<div style={{textAlign:"center",padding:32,color:"#374151"}}>No certifications recorded. Edit the worker to add certificates.</div>}
      </div>}

      {/* ══ RTW TAB ══════════════════════════════════════════════════════════ */}
      {tab==="rtw"&&<div>
        <div style={{background:"#111827",border:`1px solid ${rtwColor}44`,borderRadius:11,padding:16,marginBottom:14}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
            <div style={{fontSize:13,fontWeight:800,color:"#f1f5f9"}}>Right to Work Status</div>
            <span style={{padding:"4px 12px",borderRadius:20,fontSize:12,fontWeight:700,background:rtwColor+"22",color:rtwColor,border:`1px solid ${rtwColor}44`}}>{rtwStatus.toUpperCase()}</span>
          </div>
          {[["Share Code",w.shareCode||"Not provided"],["Date Checked",w.shareCodeDate?fmtD(w.shareCodeDate):"—"],["Expiry",w.shareCodeExpiry?fmtD(w.shareCodeExpiry):"—"],["Documents",(w.workerFiles||[]).length+" file(s) attached"]].map(([l,v])=>(
            <div key={l} style={{display:"flex",gap:8,padding:"7px 0",borderBottom:"1px solid #1e2535"}}>
              <span style={{fontSize:10,color:"#64748b",fontWeight:700,minWidth:100,textTransform:"uppercase",flexShrink:0}}>{l}</span>
              <span style={{fontSize:12,color:"#e2e8f0"}}>{v}</span>
            </div>
          ))}
          {(w.workerFiles||[]).length>0&&<div style={{marginTop:10,display:"flex",gap:6,flexWrap:"wrap"}}>
            {(w.workerFiles||[]).map((wf,i)=>(
              <a key={i} href={wf.url} target="_blank" rel="noreferrer"
                style={{display:"inline-flex",alignItems:"center",gap:5,padding:"4px 10px",background:"#1a2035",borderRadius:6,border:"1px solid #2d3555",color:"#60a5fa",fontSize:11,textDecoration:"none"}}>
                {wf.type?.startsWith("image/")?"🖼":wf.type==="application/pdf"?"📄":"📎"} {wf.name}
              </a>
            ))}
          </div>}
        </div>
        {!rtwUploaded&&<div style={{background:"#2d1515",border:"1px solid #ef444444",borderRadius:9,padding:12,fontSize:12,color:"#f87171"}}>
          ⚠ No RTW documents uploaded. Please attach ID / share code proof via the Edit button.
        </div>}
      </div>}

      {/* ══ SCHEDULE TAB ══════════════════════════════════════════════════════ */}
      {tab==="schedule"&&<div>
        <div style={{fontSize:11,color:"#64748b",marginBottom:12}}>Current week forecast allocation. ✅ = GPS confirmed on site.</div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:8}}>
          {["Mon","Tue","Wed","Thu","Fri","Sat","Sun"].map(d=>{
            const site=w.days?.[d];
            const off=!site||isOff(site);
            const confirmed=wLogs.some(l=>l.day===d&&l.weekLabel===weekLabel);
            const s=allSites.find(x=>site&&(site.includes(x.name)||x.name.includes(site)));
            const c=s?.color||"#374151";
            return <div key={d} style={{textAlign:"center",background:"#0f1421",borderRadius:8,padding:"10px 6px",border:`1px solid ${confirmed?"#34d39944":off?"#1e2535":c+"44"}`}}>
              <div style={{fontSize:10,color:confirmed?"#34d399":off?"#374151":c,fontWeight:700,textTransform:"uppercase",marginBottom:5}}>{d}</div>
              {confirmed&&<div style={{fontSize:10,color:"#34d399",marginBottom:3}}>✅</div>}
              <div style={{fontSize:10,color:off?"#374151":c,fontWeight:off?400:600,wordBreak:"break-word",lineHeight:1.3}}>{site||"—"}</div>
            </div>;
          })}
        </div>
      </div>}

      {/* ══ HOLIDAYS TAB ══════════════════════════════════════════════════════ */}
      {tab==="holidays"&&(()=>{
        const [holidays,setHolidays]=useState(w.holidayRequests||[]);
        const [from,setFrom]=useState("");const [to,setTo]=useState("");const [note,setNote]=useState("");
        return <div>
          <div style={{background:"#111827",border:"1px solid #1e2535",borderRadius:11,padding:16,marginBottom:14}}>
            <div style={{fontSize:11,color:"#64748b",fontWeight:700,textTransform:"uppercase",marginBottom:10}}>Request Holiday</div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 2fr",gap:8,marginBottom:10}}>
              <div><div style={{fontSize:9,color:"#64748b",fontWeight:700,textTransform:"uppercase",marginBottom:4}}>From</div><input type="date" value={from} onChange={e=>setFrom(e.target.value)} style={{...INP}}/></div>
              <div><div style={{fontSize:9,color:"#64748b",fontWeight:700,textTransform:"uppercase",marginBottom:4}}>To</div><input type="date" value={to} onChange={e=>setTo(e.target.value)} style={{...INP}}/></div>
              <div><div style={{fontSize:9,color:"#64748b",fontWeight:700,textTransform:"uppercase",marginBottom:4}}>Note</div><input value={note} onChange={e=>setNote(e.target.value)} placeholder="Reason…" style={{...INP}}/></div>
            </div>
            <button onClick={()=>{if(!from)return;setHolidays(h=>[...h,{id:"hol_"+Date.now(),from,to:to||from,note,status:"pending",requestedAt:new Date().toISOString()}]);setFrom("");setTo("");setNote("");}}
              style={{...BP,padding:"7px 16px",fontSize:12}}>+ Add Holiday</button>
          </div>
          {holidays.map(h=>(
            <div key={h.id} style={{background:"#111827",border:`1px solid ${h.status==="approved"?"#34d39944":h.status==="declined"?"#ef444444":"#fbbf2444"}`,borderRadius:9,padding:"11px 14px",marginBottom:8,display:"flex",alignItems:"center",gap:10}}>
              <div style={{flex:1}}>
                <div style={{fontSize:12,fontWeight:700,color:"#f1f5f9"}}>{fmtD(h.from)}{h.to&&h.to!==h.from?" — "+fmtD(h.to):""}</div>
                {h.note&&<div style={{fontSize:11,color:"#64748b",marginTop:2}}>{h.note}</div>}
              </div>
              <span style={{padding:"2px 9px",borderRadius:20,fontSize:10,fontWeight:700,background:h.status==="approved"?"#34d39922":h.status==="declined"?"#ef444422":"#fbbf2422",color:h.status==="approved"?"#34d399":h.status==="declined"?"#f87171":"#fbbf24"}}>{h.status}</span>
              {h.status==="pending"&&<div style={{display:"flex",gap:5}}>
                <button onClick={()=>setHolidays(hs=>hs.map(x=>x.id===h.id?{...x,status:"approved"}:x))} style={{padding:"4px 9px",background:"#0d2218",border:"1px solid #34d399",borderRadius:5,color:"#34d399",cursor:"pointer",fontSize:11}}>✓</button>
                <button onClick={()=>setHolidays(hs=>hs.map(x=>x.id===h.id?{...x,status:"declined"}:x))} style={{padding:"4px 9px",background:"#2d1515",border:"1px solid #ef4444",borderRadius:5,color:"#f87171",cursor:"pointer",fontSize:11}}>✕</button>
              </div>}
            </div>
          ))}
        </div>;
      })()}
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

// ─── SiteWorkersPanel — GPS-confirmed workers with scope allocation ────────────

// ─── Pending Registrations View ───────────────────────────────────────────────
function PendingWorkersView({workers,onApprove}){
  const [pending,setPending]=useState([]);
  const [loading,setLoading]=useState(true);
  const [actioning,setActioning]=useState({});
  const [expanded,setExpanded]=useState(null);
  const [rejectNote,setRejectNote]=useState({});
  const CERT_LABELS=Object.fromEntries(CERTS.map(c=>[c.key,c.label]));
  const load=async()=>{setLoading(true);try{const rows=await sbGet("pending_workers","select=id,created_at,status,data&order=created_at.desc");setPending(rows);}catch(e){console.error(e);}setLoading(false);};
  useEffect(()=>{load();},[]);
  const approve=async(row)=>{
    if(!window.confirm(`Approve ${row.data.name} and add them as an active worker?`))return;
    setActioning(a=>({...a,[row.id]:"approving"}));
    try{
      await sbUpsert("workers",[{id:row.data.id,data:{...row.data,approvedAt:new Date().toISOString()}}]);
      await fetch(`${SB_URL}/rest/v1/pending_workers?id=eq.${row.id}`,{method:"PATCH",headers:{...SB_H,"Prefer":"return=minimal"},body:JSON.stringify({status:"approved"})});
      setPending(p=>p.map(x=>x.id===row.id?{...x,status:"approved"}:x));
      if(onApprove)onApprove();
    }catch(e){alert("Approval failed: "+e.message);}
    setActioning(a=>({...a,[row.id]:null}));
  };
  const reject=async(row)=>{
    if(!window.confirm(`Reject ${row.data.name}?`))return;
    setActioning(a=>({...a,[row.id]:"rejecting"}));
    try{
      await fetch(`${SB_URL}/rest/v1/pending_workers?id=eq.${row.id}`,{method:"PATCH",headers:{...SB_H,"Prefer":"return=minimal"},body:JSON.stringify({status:"rejected",data:{...row.data,rejectedAt:new Date().toISOString(),rejectionNote:rejectNote[row.id]||""}})});
      setPending(p=>p.map(x=>x.id===row.id?{...x,status:"rejected"}:x));
    }catch(e){alert("Rejection failed: "+e.message);}
    setActioning(a=>({...a,[row.id]:null}));
  };
  const pendingRows=pending.filter(r=>r.status==="pending");
  const doneRows=pending.filter(r=>r.status!=="pending");
  const WorkerCard=({row})=>{
    const d=row.data||{};const isOpen=expanded===row.id;const act=actioning[row.id];const isPend=row.status==="pending";const isApp=row.status==="approved";
    const heldCerts=Object.entries(d.certs||{}).filter(([,v])=>v?.held);
    return <div style={{background:"#111827",border:`1px solid ${isPend?"#f59e0b44":isApp?"#34d39944":"#ef444444"}`,borderRadius:12,marginBottom:10,overflow:"hidden"}}>
      <div style={{display:"flex",alignItems:"center",gap:12,padding:"12px 15px",cursor:"pointer"}} onClick={()=>setExpanded(isOpen?null:row.id)}>
        <div style={{width:38,height:38,borderRadius:"50%",background:"linear-gradient(135deg,#1e3a5f,#3b82f6)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:14,fontWeight:900,color:"#fff",flexShrink:0}}>{d.name?.split(" ").map(n=>n[0]).join("").slice(0,2).toUpperCase()||"?"}</div>
        <div style={{flex:1}}><div style={{fontSize:14,fontWeight:800,color:"#f1f5f9"}}>{d.name||"Unknown"}</div><div style={{fontSize:11,color:"#64748b",marginTop:1}}>{d.position||"—"} · {d.company||"—"} · {d.email||"—"}</div></div>
        <div style={{display:"flex",alignItems:"center",gap:8}}>
          <span style={{fontSize:11,color:"#64748b"}}>{new Date(row.created_at).toLocaleDateString("en-GB")}</span>
          <span style={{display:"inline-block",padding:"3px 10px",borderRadius:20,fontSize:11,fontWeight:700,background:isPend?"#f59e0b22":isApp?"#34d39922":"#ef444422",color:isPend?"#fbbf24":isApp?"#34d399":"#f87171",border:`1px solid ${isPend?"#f59e0b44":isApp?"#34d39944":"#ef444444"}`}}>{isPend?"⏳ Pending":isApp?"✓ Approved":"✕ Rejected"}</span>
          <span style={{color:"#64748b",fontSize:13}}>{isOpen?"▲":"▼"}</span>
        </div>
      </div>
      {isOpen&&<div style={{borderTop:"1px solid #1e2535",padding:"14px 15px"}}>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"0 20px",marginBottom:14}}>
          {[["Full Name",d.name],["Position",d.position],["Company",d.company],["Date of Birth",d.dob?new Date(d.dob).toLocaleDateString("en-GB"):"—"],["Phone",d.phone],["NI Number",d.niNumber],["Email",d.email],["Address",d.address],["Emergency Contact",d.emergencyName],["Emergency Phone",d.emergencyPhone],["Bank Name",d.bankName],["Sort Code",d.sortCode],["Account No",d.accountNo?"••••"+d.accountNo.slice(-4):"—"],["T&Cs Signed",d.termsSignedAt?new Date(d.termsSignedAt).toLocaleString("en-GB"):"—"]].map(([l,v])=>
            <div key={l} style={{padding:"4px 0",borderBottom:"1px solid #1e2535"}}><div style={{fontSize:9,color:"#64748b",fontWeight:700,textTransform:"uppercase"}}>{l}</div><div style={{fontSize:12,color:v?"#f1f5f9":"#374151",marginTop:1}}>{v||"—"}</div></div>
          )}
        </div>
        {d.termsAccepted&&<div style={{background:"#0d2218",border:"1px solid #34d39944",borderRadius:8,padding:"8px 12px",marginBottom:12,fontSize:12,color:"#34d399"}}>✅ T&Cs signed {d.termsSignedAt?new Date(d.termsSignedAt).toLocaleString("en-GB"):""}</div>}
        {heldCerts.length>0&&<div style={{marginBottom:12}}>
          <div style={{fontSize:10,color:"#64748b",fontWeight:700,textTransform:"uppercase",marginBottom:7}}>Certifications ({heldCerts.length})</div>
          {heldCerts.map(([key,val])=><div key={key} style={{background:"#0f1421",borderRadius:7,padding:"8px 12px",marginBottom:6,border:"1px solid #1e2535"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:val.photoUrl?6:0}}>
              <span style={{fontSize:12,fontWeight:600,color:"#e2e8f0"}}>{CERT_LABELS[key]||key}</span>
              <div style={{display:"flex",gap:8}}>{val.expiry&&<span style={{fontSize:11,color:"#64748b"}}>Exp: {new Date(val.expiry).toLocaleDateString("en-GB")}</span>}<span style={{fontSize:11,fontWeight:700,color:"#34d399"}}>✓</span></div>
            </div>
            {val.photoUrl&&<img src={val.photoUrl} alt={key} style={{width:"100%",maxHeight:120,objectFit:"cover",borderRadius:6,border:"1px solid #2d3555",cursor:"pointer"}} onClick={()=>window.open(val.photoUrl,"_blank")}/>}
          </div>)}
        </div>}
        {isPend&&<div style={{borderTop:"1px solid #1e2535",paddingTop:12,marginTop:4}}>
          <div style={{marginBottom:9}}><div style={{fontSize:10,color:"#64748b",fontWeight:700,textTransform:"uppercase",marginBottom:4}}>Rejection Note (optional)</div>
            <input value={rejectNote[row.id]||""} onChange={e=>setRejectNote(n=>({...n,[row.id]:e.target.value}))} placeholder="Reason for rejection…" style={{width:"100%",background:"#0f1421",border:"1px solid #2d3555",borderRadius:7,padding:"8px 11px",color:"#e2e8f0",fontSize:12,outline:"none",boxSizing:"border-box"}}/>
          </div>
          <div style={{display:"flex",gap:9}}>
            <button onClick={()=>reject(row)} disabled={!!act} style={{flex:1,padding:"9px",background:"#2d1515",border:"1px solid #ef4444",borderRadius:8,color:"#f87171",cursor:"pointer",fontSize:12,fontWeight:700,opacity:act?0.6:1}}>{act==="rejecting"?"Rejecting…":"✕ Reject"}</button>
            <button onClick={()=>approve(row)} disabled={!!act} style={{flex:2,padding:"9px",background:"linear-gradient(135deg,#14532d,#16a34a)",border:"1px solid #34d399",borderRadius:8,color:"#fff",cursor:"pointer",fontSize:13,fontWeight:800,opacity:act?0.6:1}}>{act==="approving"?"Approving…":"✓ Approve & Add to Workers"}</button>
          </div>
        </div>}
      </div>}
    </div>;
  };
  return <div style={{padding:"16px 20px"}}>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10,marginBottom:18}}>
      {[["Pending",pendingRows.length,"#f59e0b"],["Approved",pending.filter(r=>r.status==="approved").length,"#34d399"],["Rejected",pending.filter(r=>r.status==="rejected").length,"#f87171"]].map(([l,v,c])=>
        <div key={l} style={{background:"#1a1f2e",border:`1px solid ${c}44`,borderRadius:10,padding:"10px 14px"}}><div style={{fontSize:10,color:"#64748b",fontWeight:700,textTransform:"uppercase"}}>{l}</div><div style={{fontSize:22,fontWeight:900,color:c}}>{v}</div></div>
      )}
    </div>
    <div style={{display:"flex",justifyContent:"flex-end",marginBottom:12}}><button onClick={load} style={{padding:"5px 13px",background:"#1e2535",border:"1px solid #2d3555",borderRadius:7,color:"#64748b",cursor:"pointer",fontSize:12,fontWeight:600}}>↻ Refresh</button></div>
    {loading&&<div style={{textAlign:"center",padding:40,color:"#64748b"}}>Loading registrations…</div>}
    {!loading&&pendingRows.length===0&&<div style={{textAlign:"center",padding:40,color:"#374151",fontSize:13}}><div style={{fontSize:32,marginBottom:10}}>✅</div>No pending registrations.</div>}
    {pendingRows.map(row=><WorkerCard key={row.id} row={row}/>)}
    {doneRows.length>0&&<div style={{marginTop:20}}><div style={{fontSize:11,color:"#64748b",fontWeight:700,textTransform:"uppercase",marginBottom:10}}>Previously Processed ({doneRows.length})</div>{doneRows.map(row=><WorkerCard key={row.id} row={row}/>)}</div>}
  </div>;
}

function SiteWorkersPanel({site,allWorkers,weekLabel,scopes,setDetailId,setPage}){
  const [scopeAlloc,setScopeAlloc]=useState({});
  // ── Fresh worker data fetched from Supabase on every render of this tab ──────
  // allWorkers is cached at startup; attendanceLogs are written by the Worker Portal
  // so we need a live fetch to see sign-ins that happened after the admin app loaded
  const [liveWorkers,setLiveWorkers]=useState(allWorkers);
  const [fetching,setFetching]=useState(false);
  const [lastFetch,setLastFetch]=useState(null);

  const refreshFromDB=async()=>{
    setFetching(true);
    try{
      const rows=await sbGet("workers","select=id,data&order=data->name");
      if(rows.length>0) setLiveWorkers(rows.map(r=>({...r.data,id:r.id})));
    }catch(e){console.warn("Workers refresh failed:",e.message);}
    setFetching(false);
    setLastFetch(new Date());
  };

  // Auto-refresh when panel first mounts and every 30 seconds while open
  useEffect(()=>{
    refreshFromDB();
    const t=setInterval(refreshFromDB,30000);
    return()=>clearInterval(t);
  },[site.id]);

  // ── FIX: include workers currently ON SITE (signIn exists, signOut may not yet) ──
  const confirmedWorkers=useMemo(()=>{
    return liveWorkers
      .map(w=>{
        // Include logs that have signIn — regardless of whether signOut exists yet
        const logs=(w.attendanceLogs||[]).filter(l=>
          (l.siteId===site.id||l.siteName===site.name)&&l.signIn
        );
        if(logs.length===0)return null;
        // Hours: for active sessions (no signOut), calculate up to now
        const totalHrs=logs.reduce((a,l)=>{
          const end=l.signOut?new Date(l.signOut):new Date();
          return a+Math.max(0,Math.round(((end-new Date(l.signIn))/3600000)*100)/100);
        },0);
        const activeLog=logs.find(l=>l.signIn&&!l.signOut); // currently on site
        return{...w,confirmedLogs:logs,confirmedHours:totalHrs,
               labourCost:totalHrs*(w.agreedRate||0),activeLog};
      })
      .filter(Boolean);
  },[liveWorkers,site.id,site.name]);

  // Forecast workers (in schedule but not yet signed in to this site)
  const forecastWorkers=useMemo(()=>{
    const confirmedIds=new Set(confirmedWorkers.map(w=>w.id));
    return liveWorkers.filter(w=>
      !confirmedIds.has(w.id)&&
      Object.values(w.days||{}).some(d=>(d||"").includes(site.name))
    );
  },[liveWorkers,confirmedWorkers,site.name]);

  // Cost per scope (from confirmed hours)
  const costPerScope=useMemo(()=>{
    const m={};
    confirmedWorkers.forEach(w=>{
      const sid=scopeAlloc[w.id]||"unassigned";
      m[sid]=(m[sid]||0)+w.labourCost;
    });
    return m;
  },[confirmedWorkers,scopeAlloc]);

  const totalLabour=confirmedWorkers.reduce((a,w)=>a+w.labourCost,0);
  const currentlyOnSite=confirmedWorkers.filter(w=>w.activeLog).length;
  const fmtH=h=>h>0?h.toFixed(1)+"h":"—";
  const fmtM=v=>"£"+v.toFixed(2);

  return <div>
    {/* Header with refresh button */}
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
      <div style={{fontSize:11,color:"#64748b"}}>
        {fetching?"⟳ Refreshing…":lastFetch?`Last updated: ${lastFetch.toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit",second:"2-digit"})}`:""}
      </div>
      <button onClick={refreshFromDB} disabled={fetching}
        style={{padding:"5px 13px",background:"#1e2535",border:"1px solid #3b82f6",borderRadius:7,color:"#60a5fa",cursor:"pointer",fontSize:12,fontWeight:600,display:"flex",alignItems:"center",gap:6,opacity:fetching?0.6:1}}>
        <span style={{fontSize:13,display:"inline-block",animation:fetching?"spin 1s linear infinite":""}}>{fetching?"⟳":"↻"}</span>
        Refresh Sign-ins
      </button>
    </div>
    <style>{`@keyframes spin{from{transform:rotate(0deg)}to{transform:rotate(360deg)}}`}</style>

    {/* Summary bar */}
    <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10,marginBottom:16}}>
      {[
        ["Currently On Site",currentlyOnSite+" workers","#34d399"],
        ["Forecast Only",forecastWorkers.length+" workers","#fbbf24"],
        ["Total Confirmed Hrs",confirmedWorkers.reduce((a,w)=>a+w.confirmedHours,0).toFixed(1)+"h","#60a5fa"],
        ["Total Labour Cost","£"+totalLabour.toFixed(2),"#f87171"],
      ].map(([l,v,c])=><div key={l} style={{background:"#0f1421",border:`1px solid ${c}33`,borderRadius:9,padding:"10px 13px"}}>
        <div style={{fontSize:9,color:"#64748b",fontWeight:700,textTransform:"uppercase"}}>{l}</div>
        <div style={{fontSize:16,fontWeight:800,color:c,marginTop:3}}>{v}</div>
      </div>)}
    </div>

    {/* Cost per scope */}
    {scopes.length>0&&<div style={{marginBottom:16,background:"#0f1421",borderRadius:10,padding:"12px 14px",border:"1px solid #1e2535"}}>
      <div style={{fontSize:10,color:"#64748b",fontWeight:700,textTransform:"uppercase",marginBottom:9}}>Labour Cost per Scope</div>
      {scopes.map(sc=>{
        const cost=costPerScope[sc.id]||0;
        const agreed=(Number(sc.qty)||0)*(Number(sc.rate)||0);
        const pct=agreed>0?Math.min(100,(cost/agreed*100)):0;
        return <div key={sc.id} style={{display:"flex",alignItems:"center",gap:10,marginBottom:7}}>
          <div style={{fontSize:11,color:"#94a3b8",flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={sc.description}>{sc.description||"—"}</div>
          <div style={{flex:1,height:5,background:"#1e2535",borderRadius:3,overflow:"hidden"}}>
            <div style={{height:"100%",background:pct>80?"#f87171":pct>50?"#fbbf24":"#34d399",width:pct+"%",borderRadius:3}}/>
          </div>
          <div style={{fontSize:11,fontWeight:700,color:"#f87171",minWidth:65,textAlign:"right"}}>£{cost.toFixed(2)}</div>
          <div style={{fontSize:10,color:"#64748b",minWidth:38,textAlign:"right"}}>{pct.toFixed(0)}%</div>
        </div>;
      })}
      {(costPerScope["unassigned"]||0)>0&&<div style={{display:"flex",justifyContent:"space-between",paddingTop:6,borderTop:"1px solid #1e2535",fontSize:11,color:"#64748b"}}>
        <span>Unassigned labour</span><span style={{color:"#64748b",fontWeight:700}}>£{(costPerScope["unassigned"]||0).toFixed(2)}</span>
      </div>}
      <div style={{display:"flex",justifyContent:"space-between",paddingTop:7,borderTop:"1px solid #1e2535",marginTop:4}}>
        <span style={{fontSize:11,color:"#94a3b8",fontWeight:700}}>Total</span>
        <span style={{fontSize:14,fontWeight:900,color:"#f87171"}}>£{totalLabour.toFixed(2)}</span>
      </div>
    </div>}


    {/* GPS-confirmed workers */}
    {confirmedWorkers.length>0&&<div style={{marginBottom:14}}>
      <div style={{fontSize:10,color:"#34d399",fontWeight:700,textTransform:"uppercase",marginBottom:8,display:"flex",alignItems:"center",justifyContent:"space-between"}}>
        <div style={{display:"flex",alignItems:"center",gap:6}}>
          <span style={{width:8,height:8,borderRadius:"50%",background:"#34d399",display:"inline-block"}}/>
          On Site / GPS Confirmed ({confirmedWorkers.length})
        </div>
        {confirmedWorkers.filter(w=>w.activeLog).length>0&&
          <span style={{fontSize:10,color:"#34d399",background:"#0d221844",padding:"2px 9px",borderRadius:20,border:"1px solid #34d39944"}}>
            🟢 {confirmedWorkers.filter(w=>w.activeLog).length} currently on site
          </span>}
      </div>
      <div style={{border:"1px solid #1e2535",borderRadius:10,overflow:"hidden",overflowX:"auto"}}>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,minWidth:640}}>
          <thead><tr>
            <th style={{...DS.th,textAlign:"left"}}>Worker</th>
            <th style={{...DS.th,textAlign:"center"}}>Status</th>
            <th style={{...DS.th,textAlign:"center"}}>Sessions</th>
            <th style={{...DS.th,textAlign:"right"}}>Hours</th>
            <th style={{...DS.th,textAlign:"right"}}>Cost</th>
            <th style={{...DS.th,textAlign:"left",minWidth:190}}>Scope Allocation</th>
          </tr></thead>
          <tbody>
            {confirmedWorkers.map((w,i)=>(
              <tr key={w.id} style={{background:i%2===0?"#111827":"#0f1421",cursor:"pointer",
                borderLeft:w.activeLog?"3px solid #34d399":"3px solid transparent"}}
                onClick={e=>{if(e.target.tagName==="SELECT")return;setDetailId(w.id);setPage("worker_detail");}}>
                <td style={DS.td}>
                  <div style={{display:"flex",alignItems:"center",gap:8}}>
                    <div style={{width:30,height:30,borderRadius:7,
                      background:w.activeLog?"#34d39922":"#3b82f622",
                      border:`2px solid ${w.activeLog?"#34d399":"#3b82f644"}`,
                      display:"flex",alignItems:"center",justifyContent:"center",
                      fontSize:12,fontWeight:700,color:w.activeLog?"#34d399":"#60a5fa",
                      position:"relative"}}>
                      {(w.name||"?")[0]}
                      {w.activeLog&&<span style={{position:"absolute",top:-3,right:-3,width:8,height:8,
                        borderRadius:"50%",background:"#34d399",
                        boxShadow:"0 0 5px #34d399"}}/>}
                    </div>
                    <div>
                      <div style={{fontWeight:600,color:"#f1f5f9",fontSize:12}}>{w.name}</div>
                      <div style={{fontSize:10,color:"#64748b"}}>{w.position}</div>
                    </div>
                  </div>
                </td>
                <td style={{...DS.td,textAlign:"center"}}>
                  {w.activeLog
                    ?<span style={{padding:"3px 9px",borderRadius:20,fontSize:10,fontWeight:700,
                        background:"#0d221844",color:"#34d399",border:"1px solid #34d39944",
                        display:"inline-flex",alignItems:"center",gap:4,whiteSpace:"nowrap"}}>
                        <span style={{width:5,height:5,borderRadius:"50%",background:"#34d399",display:"inline-block"}}/>
                        On Site Now
                      </span>
                    :<span style={{padding:"3px 9px",borderRadius:20,fontSize:10,fontWeight:600,
                        background:"#1e253544",color:"#64748b",border:"1px solid #2d355544",whiteSpace:"nowrap"}}>
                        Signed Out
                      </span>}
                </td>
                <td style={{...DS.td,textAlign:"center",color:"#60a5fa",fontWeight:600}}>{w.confirmedLogs.length}</td>
                <td style={{...DS.td,textAlign:"right",fontWeight:700,
                  color:w.activeLog?"#fbbf24":"#60a5fa"}}>
                  {fmtH(w.confirmedHours)}
                  {w.activeLog&&<div style={{fontSize:9,color:"#fbbf24"}}>in progress</div>}
                </td>
                <td style={{...DS.td,textAlign:"right",color:"#f87171",fontWeight:700}}>
                  {w.labourCost>0?fmtM(w.labourCost):"—"}
                  {w.activeLog&&<div style={{fontSize:9,color:"#fbbf24"}}>running</div>}
                </td>
                <td style={{...DS.td,minWidth:190}} onClick={e=>e.stopPropagation()}>
                  {scopes.length>0
                    ?<select value={scopeAlloc[w.id]||""} onChange={e=>setScopeAlloc(a=>({...a,[w.id]:e.target.value}))}
                        style={{width:"100%",background:scopeAlloc[w.id]?"#0d1e0d":"#0f1421",
                          border:`1px solid ${scopeAlloc[w.id]?"#34d399":"#2d3555"}`,
                          borderRadius:7,padding:"6px 9px",
                          color:scopeAlloc[w.id]?"#34d399":"#64748b",
                          fontSize:11,outline:"none",cursor:"pointer",fontWeight:scopeAlloc[w.id]?600:400}}>
                        <option value="">— Select scope —</option>
                        {scopes.map(sc=><option key={sc.id} value={sc.id}>{sc.description||"Scope item"}</option>)}
                      </select>
                    :<span style={{color:"#374151",fontSize:11,fontStyle:"italic"}}>No scopes on this site</span>}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <style>{`@keyframes pulse{0%,100%{opacity:1}50%{opacity:0.4}}`}</style>
    </div>}

    {/* Forecast workers (not yet signed in) */}
    {forecastWorkers.length>0&&<div>
      <div style={{fontSize:10,color:"#fbbf24",fontWeight:700,textTransform:"uppercase",marginBottom:8,display:"flex",alignItems:"center",gap:6}}>
        <span style={{width:8,height:8,borderRadius:"50%",background:"#fbbf24",display:"inline-block"}}/>
        Forecast — Allocated but not yet GPS-confirmed ({forecastWorkers.length})
      </div>
      <div style={{border:"1px solid #1e2535",borderRadius:10,overflow:"hidden"}}>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
          <thead><tr>
            <th style={{...DS.th,textAlign:"left"}}>Worker</th>
            <th style={{...DS.th,textAlign:"left"}}>Position</th>
            <th style={{...DS.th,textAlign:"left"}}>Allocated days</th>
          </tr></thead>
          <tbody>
            {forecastWorkers.map((w,i)=>{
              const allocDays=Object.entries(w.days||{}).filter(([,v])=>(v||"").includes(site.name)).map(([d])=>d);
              return <tr key={w.id} style={{background:i%2===0?"#111827":"#0f1421",cursor:"pointer"}} onClick={()=>{setDetailId(w.id);setPage("worker_detail");}}>
                <td style={DS.td}>
                  <div style={{display:"flex",alignItems:"center",gap:8}}>
                    <div style={{width:28,height:28,borderRadius:7,background:"#fbbf2422",border:"1px solid #fbbf2444",display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,fontWeight:700,color:"#fbbf24"}}>{(w.name||"?")[0]}</div>
                    <div><div style={{fontWeight:600,color:"#e2e8f0",fontSize:12}}>{w.name}</div><div style={{fontSize:10,color:"#64748b"}}>{w.company}</div></div>
                  </div>
                </td>
                <td style={{...DS.td,color:"#94a3b8"}}>{w.position||"—"}</td>
                <td style={{...DS.td}}>
                  <div style={{display:"flex",gap:4,flexWrap:"wrap"}}>
                    {allocDays.map(d=><span key={d} style={{padding:"2px 7px",borderRadius:5,background:"#fbbf2422",color:"#fbbf24",fontSize:10,fontWeight:600,border:"1px solid #fbbf2444"}}>{d}</span>)}
                  </div>
                </td>
              </tr>;
            })}
          </tbody>
        </table>
      </div>
      <div style={{fontSize:11,color:"#64748b",marginTop:7,fontStyle:"italic"}}>
        📋 These workers are in the forecast but have not GPS-signed in to this site. They will move to "GPS Confirmed" once they sign in via the Worker Portal.
      </div>
    </div>}

    {confirmedWorkers.length===0&&forecastWorkers.length===0&&<div style={{textAlign:"center",padding:36,color:"#374151",fontSize:12}}>
      <div style={{fontSize:32,marginBottom:10}}>👷</div>
      No workers allocated or GPS-confirmed on this site yet.
    </div>}
  </div>;
}

function DSiteDetail({allSites,clients,workers,activeDays,siteHours,siteId,invoices,payApplications,setPage,setDetailId,setModal,weekLabel}){
  const site=allSites.find(s=>s.id===siteId);
  if(!site) return <div style={DS.body}><div style={{color:"#374151",textAlign:"center",padding:40}}>Site not found.</div></div>;
  const client=clients.find(c=>c.id===site.clientId);
  const siteWorkers=workers.filter(w=>activeDays.some(d=>(w.days?.[d]||"").includes(site.name)));
  const sc=site.scopes||[], vr=site.variations||[];
  const isPW=site.contractType==="pricework";
  const pohPct=Number(site.pohPct)||0;
  const retPct=Number(site.retentionPct)||0;
  const scopeT=sc.reduce((a,s)=>a+(Number(s.qty)||0)*(Number(s.rate)||0),0);
  const pohTotal=isPW?scopeT*(pohPct/100):0;
  const scopeNetToBM=scopeT-pohTotal;            // budget for manager to execute
  const varT=vr.reduce((a,v)=>a+(v.type==="addition"?Number(v.value||0):-Number(v.value||0)),0);
  const contract=scopeT+varT;                    // original agreed price (for invoicing/payment apps)
  const retTotal=isPW?scopeNetToBM*(retPct/100):0;
  const netCertified=scopeNetToBM-retTotal;
  const labourCost=useMemo(()=>{let t=0;workers.forEach(w=>{const{bd}=calcPay(w,activeDays,siteHours);Object.values(bd).forEach(b=>{if(b.site===site.name||b.site.includes(site.name))t+=b.gross;});});return t;},[workers,activeDays,siteHours,site.name]);
  const siteInvs=(invoices||[]).filter(i=>i.siteId===site.id);
  const totalInvoiced=siteInvs.reduce((a,i)=>{const s=(i.lines||[]).reduce((x,l)=>x+(l.qty||0)*(l.rate||0),0);return a+s;},0);
  const totalPaid=siteInvs.filter(i=>i.status==="paid").reduce((a,i)=>{const s=(i.lines||[]).reduce((x,l)=>x+(l.qty||0)*(l.rate||0),0);return a+s;},0);
  const totalDue=totalInvoiced-totalPaid;
  const profit=contract-labourCost;
  const margin=contract>0?((contract-labourCost)/contract*100):0;
  const [tab,setTab]=useState("scopes");
  // Variation file attachments state — stored locally per session
  const [varFiles,setVarFiles]=useState({});
  const [siteFiles,setSiteFiles]=useState([]);
  const [newLinkUrl,setNewLinkUrl]=useState("");
  const [newLinkName,setNewLinkName]=useState("");

  function addSiteFile(file){
    const reader=new FileReader();
    reader.onload=ev=>{
      setSiteFiles(f=>[...f,{id:Date.now().toString(),name:file.name,size:file.size,type:file.type,url:ev.target.result,kind:"file",addedAt:new Date().toISOString()}]);
    };
    reader.readAsDataURL(file);
  }
  function addSiteLink(){
    const url=newLinkUrl.trim();
    if(!url) return;
    setSiteFiles(f=>[...f,{id:Date.now().toString(),name:newLinkName.trim()||url,url,kind:"link",type:"link",addedAt:new Date().toISOString()}]);
    setNewLinkUrl(""); setNewLinkName("");
  }
  function removeSiteFile(id){setSiteFiles(f=>f.filter(x=>x.id!==id));}

  function addVarFile(varId,file){
    const reader=new FileReader();
    reader.onload=ev=>{
      const url=ev.target.result;
      setVarFiles(f=>({...f,[varId]:[...(f[varId]||[]),{name:file.name,size:file.size,type:file.type,url,addedAt:new Date().toISOString()}]}));
    };
    reader.readAsDataURL(file);
  }
  function removeVarFile(varId,idx){
    setVarFiles(f=>({...f,[varId]:(f[varId]||[]).filter((_,i)=>i!==idx)}));
  }
  function fileIcon(type){
    if(type.startsWith("image/")) return "🖼";
    if(type.startsWith("video/")) return "🎬";
    if(type==="application/pdf") return "📄";
    if(type.includes("word")||type.includes("document")) return "📝";
    if(type.includes("email")||type.includes("message")) return "📧";
    if(type.includes("sheet")||type.includes("excel")) return "📊";
    return "📎";
  }

  // 6 financial cards
  const financials=[
    {label:"Contract Value",value:"£"+Math.round(contract).toLocaleString(),color:"#60a5fa",sub:"Scopes + variations"},
    {label:"Labour Cost (Gross)",value:"£"+Math.round(labourCost).toLocaleString(),color:"#f87171",sub:"Total weekly labour"},
    {label:"Invoiced to Date",value:"£"+Math.round(totalInvoiced).toLocaleString(),color:"#fbbf24",sub:siteInvs.length+" invoice"+(siteInvs.length!==1?"s":"")},
    {label:"Paid to Date",value:"£"+Math.round(totalPaid).toLocaleString(),color:"#34d399",sub:siteInvs.filter(i=>i.status==="paid").length+" paid"},
    {label:"Total Due",value:"£"+Math.round(totalDue).toLocaleString(),color:totalDue>0?"#f97316":"#34d399",sub:totalDue>0?"Outstanding":"Fully collected"},
    {label:"Profit / Loss",value:"£"+Math.round(Math.abs(profit)).toLocaleString(),color:profit>=0?"#a78bfa":"#f87171",sub:(profit>=0?"Profit ":"Loss ")+Math.abs(margin).toFixed(1)+"%"},
  ];

  return <div>
    <DPageHdr
      title={<span style={{display:"flex",alignItems:"center",gap:9}}>
        <span style={{width:13,height:13,borderRadius:"50%",background:site.color,boxShadow:"0 0 8px "+site.color+"88"}}/>
        {site.name}
      </span>}
      sub={client?<span>Client: <span style={{color:client.color,fontWeight:700}}>{client.name}</span></span>:"No client assigned"}
      back="Sites" onBack={()=>setPage("sites")}
      actions={<div style={{display:"flex",gap:7}}>
        <button onClick={()=>setModal({type:"siteDetail",site})} style={{padding:"6px 12px",background:"#1e3a5f",border:"1px solid #3b82f6",borderRadius:6,color:"#60a5fa",cursor:"pointer",fontSize:12,fontWeight:600}}>✏️ Edit Site</button>
      </div>}/>

    <div style={DS.body}>
      {/* 6-card financial header */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(6,1fr)",gap:10,marginBottom:20}}>
        {financials.map(({label,value,color,sub})=>(
          <div key={label} style={{background:"linear-gradient(145deg,#141924,#1a2035)",border:"1px solid "+color+"33",borderRadius:12,padding:"13px 14px",position:"relative",overflow:"hidden"}}>
            <div style={{position:"absolute",top:0,left:0,right:0,height:2,background:color}}/>
            <div style={{fontSize:9,color:"#64748b",fontWeight:700,textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:4}}>{label}</div>
            <div style={{fontSize:18,fontWeight:900,color,lineHeight:1,marginBottom:4}}>{value}</div>
            <div style={{fontSize:10,color:"#374151"}}>{sub}</div>
          </div>
        ))}
      </div>

      {/* Profit bar */}
      {contract>0&&<div style={{background:"#111827",borderRadius:10,padding:"11px 16px",marginBottom:18,border:"1px solid #1e2535",display:"flex",alignItems:"center",gap:14}}>
        <span style={{fontSize:11,color:"#64748b",fontWeight:700,minWidth:110}}>Cost vs Contract</span>
        <div style={{flex:1,height:8,background:"#1e2535",borderRadius:4,overflow:"hidden",position:"relative"}}>
          <div style={{position:"absolute",left:0,top:0,height:"100%",borderRadius:4,background:labourCost>contract?"#ef4444":"linear-gradient(90deg,#34d399,#10b981)",width:Math.min(100,labourCost/contract*100)+"%",transition:"width 0.6s ease"}}/>
        </div>
        <span style={{fontSize:12,fontWeight:700,color:labourCost>contract?"#f87171":"#34d399",minWidth:80,textAlign:"right"}}>{Math.round(labourCost/contract*100)}% used</span>
        <span style={{fontSize:11,color:"#64748b"}}>Labour £{Math.round(labourCost).toLocaleString()} of £{Math.round(contract).toLocaleString()}</span>
      </div>}

      {/* Tab bar */}
      <div style={{display:"flex",gap:3,background:"#0d1117",borderRadius:8,padding:3,marginBottom:18,width:"fit-content",flexWrap:"wrap"}}>
        {[["scopes","📋 Scopes ("+(sc.length)+")"],["variations","⚡ Variations ("+(vr.length)+")"],["workers","👷 Workers ("+(siteWorkers.length)+")"],["costs","💷 Full Costs"],["invoices","🧾 Invoices ("+(siteInvs.length)+")"],["docs","📁 Documents"]].map(([v,l])=>(
          <button key={v} onClick={()=>setTab(v)} style={{padding:"6px 14px",background:tab===v?"#1e3a5f":"transparent",border:tab===v?"1px solid #3b82f6":"1px solid transparent",borderRadius:6,color:tab===v?"#60a5fa":"#64748b",cursor:"pointer",fontSize:12,fontWeight:tab===v?700:400}}>{l}</button>
        ))}
      </div>

      {/* Scopes */}
      {tab==="scopes"&&<div>
        {isPW&&(pohPct>0||retPct>0)&&<div style={{display:"flex",gap:14,marginBottom:12,padding:"8px 14px",background:"#0f1421",borderRadius:8,border:"1px solid #1e2535",fontSize:11,flexWrap:"wrap",alignItems:"center"}}>
          <span style={{color:"#64748b",fontWeight:700}}>📐 Price Work</span>
          {pohPct>0&&<span style={{color:"#fbbf24"}}>P&OH deducted: <b>{pohPct}%</b></span>}
          {retPct>0&&<span style={{color:"#f97316"}}>Retention: <b>{retPct}%</b></span>}
          <span style={{color:"#64748b",marginLeft:"auto",fontStyle:"italic"}}>Agreed price shown in full · Budget = agreed − P&OH</span>
        </div>}
        {sc.length===0?<div style={{textAlign:"center",padding:40,border:"1px dashed #1e2535",borderRadius:10,color:"#374151"}}>
          No scopes yet. Click "✏️ Edit Site" to add scope line items.
        </div>:
        <div style={{border:"1px solid #1e2535",borderRadius:10,overflow:"hidden",overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse",minWidth:isPW?780:500}}>
            <thead><tr>
              <th style={DS.th}>Description</th>
              <th style={{...DS.th,width:60,textAlign:"center"}}>Unit</th>
              <th style={{...DS.th,width:55,textAlign:"right"}}>Qty</th>
              <th style={{...DS.th,width:90,textAlign:"right"}}>Rate £</th>
              <th style={{...DS.th,width:100,textAlign:"right",color:"#60a5fa"}}>Agreed £</th>
              {isPW&&pohPct>0&&<th style={{...DS.th,width:95,textAlign:"right",color:"#fbbf24"}}>P&OH ({pohPct}%)</th>}
              {isPW&&<th style={{...DS.th,width:100,textAlign:"right",color:"#34d399"}}>Budget £</th>}
              {isPW&&retPct>0&&<th style={{...DS.th,width:100,textAlign:"right",color:"#f97316"}}>Retention</th>}
              {isPW&&retPct>0&&<th style={{...DS.th,width:105,textAlign:"right",color:"#a78bfa"}}>Net Certified</th>}
              {!isPW&&<th style={{...DS.th,width:105,textAlign:"right",color:"#34d399"}}>Total £</th>}
            </tr></thead>
            <tbody>
              {sc.map((s,i)=>{
                const agreed=(Number(s.qty)||0)*(Number(s.rate)||0);
                const poh=isPW?agreed*(pohPct/100):0;
                const budget=agreed-poh;
                const ret=isPW?budget*(retPct/100):0;
                const cert=budget-ret;
                return <tr key={s.id||i} style={{background:i%2===0?"#111827":"#0f1421"}}>
                  <td style={{...DS.td,fontWeight:600,color:"#f1f5f9"}}>{s.description||"—"}</td>
                  <td style={{...DS.td,textAlign:"center"}}><span style={DS.badge("#94a3b8","#1e2535")}>{s.unit}</span></td>
                  <td style={{...DS.td,textAlign:"right",color:"#60a5fa",fontWeight:600}}>{s.qty}</td>
                  <td style={{...DS.td,textAlign:"right"}}>£{Number(s.rate).toLocaleString()}</td>
                  <td style={{...DS.td,textAlign:"right",color:"#60a5fa",fontWeight:600}}>£{agreed.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2})}</td>
                  {isPW&&pohPct>0&&<td style={{...DS.td,textAlign:"right",color:"#fbbf24"}}>-£{poh.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2})}</td>}
                  {isPW&&<td style={{...DS.td,textAlign:"right",color:"#34d399",fontWeight:700}}>£{budget.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2})}</td>}
                  {isPW&&retPct>0&&<td style={{...DS.td,textAlign:"right",color:"#f97316"}}>-£{ret.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2})}</td>}
                  {isPW&&retPct>0&&<td style={{...DS.td,textAlign:"right",color:"#a78bfa",fontWeight:700}}>£{cert.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2})}</td>}
                  {!isPW&&<td style={{...DS.td,textAlign:"right",color:"#34d399",fontWeight:700}}>£{agreed.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2})}</td>}
                </tr>;
              })}
            </tbody>
            <tfoot>
              <tr style={{background:"#0d1117",borderTop:"2px solid #2d3555"}}>
                <td colSpan={4} style={{...DS.td,fontWeight:700,color:"#94a3b8"}}>TOTAL SCOPE</td>
                <td style={{...DS.td,textAlign:"right",color:"#60a5fa",fontWeight:800}}>£{scopeT.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2})}</td>
                {isPW&&pohPct>0&&<td style={{...DS.td,textAlign:"right",color:"#fbbf24",fontWeight:800}}>-£{pohTotal.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2})}</td>}
                {isPW&&<td style={{...DS.td,textAlign:"right",color:"#34d399",fontWeight:900,fontSize:14}}>£{scopeNetToBM.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2})}</td>}
                {isPW&&retPct>0&&<td style={{...DS.td,textAlign:"right",color:"#f97316",fontWeight:800}}>-£{retTotal.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2})}</td>}
                {isPW&&retPct>0&&<td style={{...DS.td,textAlign:"right",color:"#a78bfa",fontWeight:900,fontSize:14}}>£{netCertified.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2})}</td>}
                {!isPW&&<td style={{...DS.td,textAlign:"right",color:"#34d399",fontWeight:900,fontSize:14}}>£{scopeT.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2})}</td>}
              </tr>
              {isPW&&<tr style={{background:"#0a1a0d",borderTop:"1px solid #1e2535"}}>
                <td colSpan={4} style={{...DS.td,fontSize:10,color:"#64748b",fontStyle:"italic"}}>
                  Agreed price = full contract value invoiced to client · Budget = manager's allocation after P&OH deduction
                </td>
                {pohPct>0&&<td style={{...DS.td}}/>}
                <td style={{...DS.td,textAlign:"right"}}>
                  <span style={{fontSize:10,color:"#34d399",fontWeight:700}}>Manager budget: £{scopeNetToBM.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2})}</span>
                </td>
                {retPct>0&&<td style={{...DS.td}}/>}
                {retPct>0&&<td style={{...DS.td}}/>}
              </tr>}
            </tfoot>
          </table>
        </div>}
      </div>}

      {/* Variations — with file attachments + inline add */}
      {tab==="variations"&&<div>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:14}}>
          <div style={{fontSize:11,color:"#64748b",fontWeight:700,textTransform:"uppercase",letterSpacing:"0.06em"}}>
            {vr.length} Variation{vr.length!==1?"s":""} · Net: <span style={{color:varT>=0?"#34d399":"#f87171"}}>{varT>=0?"+":"-"}£{Math.abs(Math.round(varT)).toLocaleString()}</span>
          </div>
          <button onClick={()=>setModal({type:"siteDetail",site,defaultTab:"variations"})}
            style={{padding:"6px 14px",background:"linear-gradient(135deg,#f59e0b,#d97706)",border:"none",borderRadius:7,color:"#fff",cursor:"pointer",fontSize:12,fontWeight:700}}>
            ⚡ + Add / Edit Variations
          </button>
        </div>
        {vr.length===0?<div style={{textAlign:"center",padding:40,border:"1px dashed #1e2535",borderRadius:10,color:"#374151"}}>
          No variations yet. Click "⚡ + Add / Edit Variations" above.
        </div>:vr.map((v,i)=>{
          const files=varFiles[v.id]||[];
          const isAdd=v.type==="addition";
          const val=Number(v.value||0);
          const bc=isAdd?"#34d399":"#f87171";
          return <div key={v.id||i} style={{background:"linear-gradient(145deg,#141924,#1a2035)",border:"1px solid "+(isAdd?"#065f4666":"#7f1d1d66"),borderRadius:12,padding:16,marginBottom:12,borderLeft:"4px solid "+bc}}>
            <div style={{display:"flex",alignItems:"flex-start",gap:12,marginBottom:12}}>
              <div style={{flex:1}}>
                <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:5}}>
                  <span style={{fontSize:13,fontWeight:800,color:"#f1f5f9"}}>{v.description||"—"}</span>
                  <DStatusBadge status={v.type==="addition"?"addition":"omission"}/>
                  <DStatusBadge status={v.approved?"approved":"pending"}/>
                </div>
                {v.notes&&<div style={{fontSize:11,color:"#64748b",marginBottom:4}}>{v.notes}</div>}
              </div>
              <div style={{textAlign:"right",flexShrink:0}}>
                <div style={{fontSize:9,color:"#64748b",textTransform:"uppercase"}}>Value</div>
                <div style={{fontSize:20,fontWeight:900,color:bc}}>{isAdd?"+":"-"}£{val.toLocaleString()}</div>
              </div>
            </div>

            {/* File attachments */}
            <div style={{background:"#0f1421",borderRadius:9,padding:"10px 12px",border:"1px solid #1e2535"}}>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:files.length>0?10:0}}>
                <span style={{fontSize:10,color:"#64748b",fontWeight:700,textTransform:"uppercase",letterSpacing:"0.06em"}}>📎 Attachments ({files.length})</span>
                <label style={{padding:"4px 11px",background:"#1e3a5f",border:"1px solid #3b82f6",borderRadius:6,color:"#60a5fa",cursor:"pointer",fontSize:10,fontWeight:700,display:"flex",alignItems:"center",gap:5}}>
                  + Attach File
                  <input type="file" multiple accept="image/*,video/*,.pdf,.doc,.docx,.xls,.xlsx,.eml,.msg" style={{display:"none"}} onChange={e=>{Array.from(e.target.files||[]).forEach(f=>addVarFile(v.id,f));e.target.value="";}}/>
                </label>
              </div>
              {files.length>0&&<div style={{display:"flex",flexWrap:"wrap",gap:7}}>
                {files.map((f,fi)=>(
                  <div key={fi} style={{display:"flex",alignItems:"center",gap:6,padding:"5px 9px",background:"#1a2035",borderRadius:7,border:"1px solid #2d3555",maxWidth:220}}>
                    <span style={{fontSize:15,flexShrink:0}}>{fileIcon(f.type)}</span>
                    {f.type.startsWith("image/")?
                      <a href={f.url} target="_blank" rel="noreferrer" style={{textDecoration:"none"}}>
                        <img src={f.url} alt={f.name} style={{width:40,height:40,objectFit:"cover",borderRadius:4,border:"1px solid #2d3555"}}/>
                      </a>:
                      <a href={f.url} download={f.name} style={{fontSize:10,color:"#60a5fa",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",maxWidth:120,textDecoration:"none"}} title={f.name}>{f.name}</a>
                    }
                    <span style={{fontSize:9,color:"#374151",flexShrink:0}}>{f.size>1048576?(f.size/1048576).toFixed(1)+"MB":(f.size/1024).toFixed(0)+"KB"}</span>
                    <button onClick={()=>removeVarFile(v.id,fi)} style={{background:"none",border:"none",color:"#f87171",cursor:"pointer",fontSize:13,lineHeight:1,flexShrink:0,padding:"0 2px"}}>×</button>
                  </div>
                ))}
              </div>}
              {files.length===0&&<div style={{fontSize:10,color:"#374151",marginTop:4}}>Attach photos, videos, PDFs, emails or documents as evidence for this variation.</div>}
            </div>
          </div>;
        })}
        {vr.length>0&&<div style={{textAlign:"right",padding:"10px 0",fontSize:14,fontWeight:800,color:varT>=0?"#34d399":"#f87171"}}>
          Net Variations: {varT>=0?"+":"-"}£{Math.abs(varT).toLocaleString()}
        </div>}
      </div>}

      {/* Workers */}
      {tab==="workers"&&<SiteWorkersPanel site={site} allWorkers={workers} weekLabel={weekLabel} scopes={sc} setDetailId={setDetailId} setPage={setPage}/>}

      {/* Full Costs */}
      {tab==="costs"&&<div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginBottom:16}}>
          <div style={{background:"#0f1421",borderRadius:11,padding:16,border:"1px solid #1e2535"}}>
            <div style={{fontSize:11,color:"#34d399",fontWeight:700,textTransform:"uppercase",marginBottom:12}}>Income</div>
            {[["Agreed Scope","£"+Math.round(scopeT).toLocaleString(),"#60a5fa"],["Approved Variations",(varT>=0?"+":"-")+"£"+Math.round(Math.abs(varT)).toLocaleString(),"#fbbf24"],["Total Contract Value","£"+Math.round(contract).toLocaleString(),"#34d399"],["Invoiced to Date","£"+Math.round(totalInvoiced).toLocaleString(),"#fbbf24"],["Paid to Date","£"+Math.round(totalPaid).toLocaleString(),"#34d399"],["Outstanding Due","£"+Math.round(totalDue).toLocaleString(),totalDue>0?"#f97316":"#34d399"]].map(([l,v,c])=>(
              <div key={l} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"9px 0",borderBottom:"1px solid #1e2535"}}>
                <span style={{color:"#94a3b8",fontSize:12}}>{l}</span><span style={{fontWeight:700,color:c,fontSize:13}}>{v}</span>
              </div>
            ))}
          </div>
          <div style={{background:"#0f1421",borderRadius:11,padding:16,border:"1px solid #1e2535"}}>
            <div style={{fontSize:11,color:"#f87171",fontWeight:700,textTransform:"uppercase",marginBottom:12}}>Costs</div>
            {[["Labour Cost (Gross)","£"+Math.round(labourCost).toLocaleString(),"#f87171"],["Labour Cost (Net)","£"+Math.round(workers.reduce((a,w)=>{const{bd}=calcPay(w,activeDays,siteHours);return a+Object.values(bd).filter(b=>b.site===site.name||b.site.includes(site.name)).reduce((x,b)=>x+b.gross,0);},0)*(1-(workers.find(w=>activeDays.some(d=>(w.days?.[d]||"").includes(site.name)))?.taxRate||0))).toLocaleString(),"#ef4444"]].map(([l,v,c])=>(
              <div key={l} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"9px 0",borderBottom:"1px solid #1e2535"}}>
                <span style={{color:"#94a3b8",fontSize:12}}>{l}</span><span style={{fontWeight:700,color:c,fontSize:13}}>{v}</span>
              </div>
            ))}
            <div style={{background:profit>=0?"#0d2218":"#2d1515",border:"2px solid "+(profit>=0?"#10b981":"#ef4444"),borderRadius:10,padding:14,textAlign:"center",marginTop:14}}>
              <div style={{fontSize:10,color:profit>=0?"#34d399":"#f87171",fontWeight:700,textTransform:"uppercase",marginBottom:4}}>{profit>=0?"Profit":"Loss"}</div>
              <div style={{fontSize:28,fontWeight:900,color:profit>=0?"#34d399":"#f87171"}}>£{Math.round(Math.abs(profit)).toLocaleString()}</div>
              {contract>0&&<div style={{fontSize:11,color:"#64748b",marginTop:4}}>Margin: {Math.abs(margin).toFixed(1)}%</div>}
            </div>
          </div>
        </div>
      </div>}

      {/* Invoices */}
      {tab==="invoices"&&<div>
        {/* Latest Payment Application banner */}
        {(()=>{
          const sitePAs=(payApplications||[]).filter(p=>p.siteId===site.id).sort((a,b)=>new Date(b.createdAt)-new Date(a.createdAt));
          const latestPA=sitePAs[0];
          if(!latestPA) return null;
          const totalClaimed=latestPA.items.reduce((a,it)=>a+(it.useQty?(it.claimedQtyToDate||0)*it.contractRate:((it.claimedPctToDate||0)/100)*(it.contractQty*it.contractRate)),0);
          const totalContract=latestPA.items.reduce((a,it)=>a+it.contractQty*it.contractRate,0);
          const pct=totalContract>0?Math.round(totalClaimed/totalContract*100):0;
          const statusColor={draft:"#64748b",submitted:"#60a5fa",certified:"#34d399",paid:"#a78bfa"}[latestPA.status]||"#64748b";
          return <div style={{background:"linear-gradient(145deg,#0c1a2e,#111827)",border:"1px solid #1e3a5f",borderRadius:11,padding:14,marginBottom:16}}>
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:10}}>
              <div style={{display:"flex",alignItems:"center",gap:9}}>
                <span style={{fontSize:13}}>📐</span>
                <div>
                  <div style={{fontSize:12,fontWeight:700,color:"#f1f5f9"}}>Latest Payment Application: <span style={{color:"#a78bfa"}}>{latestPA.number}</span></div>
                  <div style={{fontSize:10,color:"#64748b"}}>{latestPA.date} · {latestPA.items.length} items</div>
                </div>
              </div>
              <div style={{display:"flex",alignItems:"center",gap:10}}>
                <span style={{padding:"2px 9px",borderRadius:5,fontSize:10,fontWeight:700,color:statusColor,background:statusColor+"18",border:"1px solid "+statusColor+"44",textTransform:"capitalize"}}>{latestPA.status}</span>
                <button onClick={()=>setPage("payapps")} style={{padding:"4px 11px",background:"#1e3a5f",border:"1px solid #3b82f6",borderRadius:5,color:"#60a5fa",cursor:"pointer",fontSize:10,fontWeight:700}}>View All →</button>
              </div>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:9}}>
              {[["Contract",totalContract,"#60a5fa"],["Claimed to Date",totalClaimed,"#34d399"],["This Period",totalClaimed-latestPA.items.reduce((a,it)=>a+(it.previousQty||0)*it.contractRate,0),"#fbbf24"],["% Complete",pct+"%","#a78bfa"]].map(([l,v,c])=>(
                <div key={l} style={{background:"#0d1117",borderRadius:7,padding:"7px 10px",border:"1px solid "+c+"22"}}>
                  <div style={{fontSize:9,color:"#64748b",textTransform:"uppercase",fontWeight:700}}>{l}</div>
                  <div style={{fontSize:14,fontWeight:800,color:c,marginTop:3}}>{typeof v==="number"?"£"+Math.round(v).toLocaleString():v}</div>
                </div>
              ))}
            </div>
            {sitePAs.length>1&&<div style={{marginTop:8,fontSize:10,color:"#64748b"}}>{sitePAs.length} total applications for this site</div>}
          </div>;
        })()}
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
          <div style={{fontSize:11,color:"#64748b",fontWeight:700,textTransform:"uppercase",letterSpacing:"0.06em"}}>{siteInvs.length} Invoice{siteInvs.length!==1?"s":""}</div>
          <button onClick={()=>{
            const inv=emptyInvoice(clients,allSites,invoices);
            inv.siteId=site.id;
            inv.clientId=site.clientId||"";
            // pre-fill site name and client
            setModal({type:"invoice",invoice:inv});
          }} style={{padding:"6px 16px",background:"linear-gradient(135deg,#3b82f6,#6366f1)",border:"none",borderRadius:7,color:"#fff",cursor:"pointer",fontSize:12,fontWeight:700}}>
            🧾 + New Invoice
          </button>
        </div>
        {siteInvs.length===0?<div style={{textAlign:"center",padding:40,border:"1px dashed #1e2535",borderRadius:10,color:"#374151"}}>No invoices for this site yet. Click "+ New Invoice" above.</div>:
        <DTable cols={[
          {key:"invoiceNumber",label:"Invoice No.",r:v=><span style={{color:"#60a5fa",fontWeight:700}}>{v||"Draft"}</span>},
          {key:"issueDate",label:"Date",r:v=>v?new Date(v).toLocaleDateString("en-GB",{day:"2-digit",month:"short",year:"numeric"}):"—"},
          {key:"lines",label:"Amount",r:v=>{const t=(v||[]).reduce((a,l)=>a+(l.qty||0)*(l.rate||0),0);return <span style={{color:"#34d399",fontWeight:700}}>£{Math.round(t).toLocaleString()}</span>;}},
          {key:"status",label:"Status",r:v=><DStatusBadge status={v||"draft"}/>},
          {key:"id",label:"",r:(_,r)=><button onClick={e=>{e.stopPropagation();setModal({type:"invoice",invoice:r});}} style={{padding:"3px 9px",background:"#1e3a5f",border:"1px solid #3b82f6",borderRadius:4,color:"#60a5fa",cursor:"pointer",fontSize:10,fontWeight:700}}>✏️ Edit</button>},
        ]} rows={siteInvs}/>}
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"10px 0",fontSize:13,fontWeight:700,color:"#fbbf24"}}>
          <span>Total Invoiced: £{Math.round(totalInvoiced).toLocaleString()} · Paid: £{Math.round(totalPaid).toLocaleString()}</span>
          <span style={{color:totalDue>0?"#f97316":"#34d399"}}>Due: £{Math.round(totalDue).toLocaleString()}</span>
        </div>
      </div>}

      {/* Documents & Links directory */}
      {tab==="docs"&&<div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14,marginBottom:16}}>
          {/* Upload files */}
          <div style={{background:"#111827",border:"1px solid #1e2535",borderRadius:11,padding:14}}>
            <div style={{fontSize:11,color:"#64748b",fontWeight:700,textTransform:"uppercase",marginBottom:10}}>Upload Files</div>
            <label style={{display:"flex",alignItems:"center",justifyContent:"center",gap:8,padding:"14px",background:"#1e3a5f",border:"2px dashed #3b82f6",borderRadius:8,cursor:"pointer",color:"#60a5fa",fontSize:12,fontWeight:600}}>
              📁 Click to upload files
              <input type="file" multiple accept="image/*,video/*,.pdf,.doc,.docx,.xls,.xlsx,.eml,.msg,.zip,.dwg,.dxf" style={{display:"none"}}
                onChange={e=>{Array.from(e.target.files||[]).forEach(f=>addSiteFile(f));e.target.value="";}}/>
            </label>
            <div style={{fontSize:10,color:"#374151",marginTop:6}}>Images, PDFs, Word, Excel, emails, CAD files, videos, ZIP</div>
          </div>
          {/* Add shared link */}
          <div style={{background:"#111827",border:"1px solid #1e2535",borderRadius:11,padding:14}}>
            <div style={{fontSize:11,color:"#64748b",fontWeight:700,textTransform:"uppercase",marginBottom:10}}>Add Shared Link</div>
            <input value={newLinkName} onChange={e=>setNewLinkName(e.target.value)} placeholder="Label (e.g. JAUK Drawing Pack)" style={{width:"100%",background:"#0f1421",border:"1px solid #2d3555",borderRadius:6,padding:"6px 9px",color:"#e2e8f0",fontSize:12,outline:"none",marginBottom:7,boxSizing:"border-box"}}/>
            <div style={{display:"flex",gap:7}}>
              <input value={newLinkUrl} onChange={e=>setNewLinkUrl(e.target.value)} onKeyDown={e=>e.key==="Enter"&&addSiteLink()} placeholder="https://drive.google.com/… or any URL" style={{flex:1,background:"#0f1421",border:"1px solid #2d3555",borderRadius:6,padding:"6px 9px",color:"#e2e8f0",fontSize:12,outline:"none",boxSizing:"border-box"}}/>
              <button onClick={addSiteLink} style={{padding:"6px 13px",background:"linear-gradient(135deg,#3b82f6,#6366f1)",border:"none",borderRadius:6,color:"#fff",cursor:"pointer",fontSize:12,fontWeight:700,whiteSpace:"nowrap"}}>+ Add</button>
            </div>
          </div>
        </div>

        {/* Files & links grid */}
        {siteFiles.length===0?<div style={{textAlign:"center",padding:40,border:"1px dashed #1e2535",borderRadius:10,color:"#374151"}}>
          <div style={{fontSize:28,marginBottom:8}}>📁</div>
          No files or links yet. Upload documents or add shared links above.
        </div>:
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(200px,1fr))",gap:10}}>
          {siteFiles.map(f=>{
            const isImg=f.type&&f.type.startsWith("image/");
            const isLink=f.kind==="link";
            const icon=isLink?"🔗":f.type?.startsWith("image/")?"🖼":f.type?.startsWith("video/")?"🎬":f.type==="application/pdf"?"📄":f.type?.includes("word")?"📝":f.type?.includes("sheet")||f.type?.includes("excel")?"📊":f.type?.includes("email")||f.type?.includes("message")?"📧":"📎";
            return <div key={f.id} style={{background:"linear-gradient(145deg,#141924,#1a2035)",border:"1px solid #2d3555",borderRadius:10,padding:12,position:"relative"}}>
              <button onClick={()=>removeSiteFile(f.id)} style={{position:"absolute",top:7,right:8,background:"none",border:"none",color:"#f87171",cursor:"pointer",fontSize:15,lineHeight:1,padding:0}}>×</button>
              {isImg&&!isLink?
                <a href={f.url} target="_blank" rel="noreferrer">
                  <img src={f.url} alt={f.name} style={{width:"100%",height:90,objectFit:"cover",borderRadius:6,display:"block",marginBottom:7,border:"1px solid #2d3555"}}/>
                </a>:
                <div style={{height:60,display:"flex",alignItems:"center",justifyContent:"center",fontSize:28,marginBottom:7}}>{icon}</div>
              }
              <div style={{fontSize:11,color:"#f1f5f9",fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",marginBottom:3}} title={f.name}>{f.name}</div>
              {!isLink&&<div style={{fontSize:9,color:"#374151",marginBottom:6}}>{f.size>1048576?(f.size/1048576).toFixed(1)+"MB":(f.size/1024).toFixed(0)+"KB"}</div>}
              <div style={{display:"flex",gap:6}}>
                <a href={f.url} target="_blank" rel="noreferrer" style={{flex:1,padding:"4px 0",background:"#1e3a5f",border:"1px solid #3b82f6",borderRadius:5,color:"#60a5fa",cursor:"pointer",fontSize:10,fontWeight:700,textAlign:"center",textDecoration:"none",display:"block"}}>
                  {isLink?"🔗 Open":"👁 View"}
                </a>
                {!isLink&&<a href={f.url} download={f.name} style={{padding:"4px 8px",background:"#0f1421",border:"1px solid #2d3555",borderRadius:5,color:"#94a3b8",cursor:"pointer",fontSize:10,textDecoration:"none",display:"block"}}>⬇</a>}
              </div>
              <div style={{fontSize:8,color:"#374151",marginTop:5}}>{new Date(f.addedAt).toLocaleDateString("en-GB",{day:"2-digit",month:"short",year:"numeric"})}</div>
            </div>;
          })}
        </div>}
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

// ═══════════════════════════════════════════════════════════════════════════
// TIMESHEETS — auto-generated from schedule, week-locked, finance approval
// ═══════════════════════════════════════════════════════════════════════════
function DTimesheets({workers,allSites,activeDays,siteHours,weekLabel,timesheetRecords,setTimesheetRecords,generateTimesheets,generatePayslips,payslipRecords,setPayslipRecords,setPage}){
  const [selWeek,setSelWeek]=useState(weekLabel);
  const [editId,setEditId]=useState(null);

  // Auto-sync current week timesheets from schedule whenever weekLabel or workers change
  useEffect(()=>{
    if(!workers||workers.length===0) return;
    const already=timesheetRecords.some(t=>t.weekLabel===weekLabel&&t.source==="auto");
    if(already) return; // don't overwrite if already generated
    const sheets=buildSheets(weekLabel,workers,activeDays,siteHours);
    if(sheets.length===0) return;
    setTimesheetRecords(prev=>{
      const other=prev.filter(t=>t.weekLabel!==weekLabel);
      return [...other,...sheets];
    });
  },[weekLabel,workers.length]);

  function buildSheets(wkLabel,wkrs,days,shrs){
    return wkrs.map(w=>{
      const dayBreakdown={};
      let stdH=0,otH=0,gross=0;
      (days||BASE_DAYS).forEach(d=>{
        const site=(w.days?.[d]||"").trim();
        if(!site||isOff(site)) return;
        const hrs=shrs?.[site]?.hours||w.hoursPerDay?.[d]||9;
        const ot=w.overtimeHours?.[d]||0;
        const rate=w.agreedRate||0;
        const otM=w.customOTRate||(w.overtimeMultiplier||1.5);
        const stdPay=hrs*rate, otPay=ot*rate*otM;
        stdH+=hrs; otH+=ot; gross+=stdPay+otPay;
        dayBreakdown[d]={site,hours:hrs,ot,stdPay,otPay,total:stdPay+otPay};
      });
      const tax=gross*(w.taxRate||0);
      return {
        id:"ts_"+w.id+"_"+wkLabel.replace(/\s+/g,""),
        workerId:w.id, workerName:w.name, position:w.position,
        company:w.company||"", weekLabel:wkLabel,
        stdHours:stdH, otHours:otH,
        rate:w.agreedRate||0, taxRate:w.taxRate||0,
        gross, tax, net:gross-tax,
        dayBreakdown, days:JSON.parse(JSON.stringify(w.days||{})),
        status:"draft",     // draft → submitted → approved → payslip_generated
        source:"auto", notes:"",
        lockedAt:null, approvedAt:null, approvedBy:"",
        createdAt:new Date().toISOString(),
      };
    }).filter(t=>t.stdHours>0||t.otHours>0);
  }

  // Recalculate a single timesheet from its day data
  function recalcSheet(t,w){
    let stdH=0,otH=0,gross=0;
    const bd={};
    (activeDays||BASE_DAYS).forEach(d=>{
      const site=(t.days?.[d]||"").trim();
      if(!site||isOff(site)) return;
      const hrs=t.dayBreakdown?.[d]?.hours||9;
      const ot=t.dayBreakdown?.[d]?.ot||0;
      const rate=t.rate||0;
      const otM=w?.customOTRate||(w?.overtimeMultiplier||1.5);
      const stdPay=hrs*rate, otPay=ot*rate*otM;
      stdH+=hrs; otH+=ot; gross+=stdPay+otPay;
      bd[d]={site,hours:hrs,ot,stdPay,otPay,total:stdPay+otPay};
    });
    const tax=gross*(t.taxRate||0);
    return {...t,stdHours:stdH,otHours:otH,gross,tax,net:gross-tax,dayBreakdown:bd};
  }

  function regenWeek(){
    if(!window.confirm("Re-sync WC "+weekLabel+" timesheets from current schedule data?\n\nThis will update DRAFT timesheets only. Approved timesheets are unchanged.")) return;
    const sheets=buildSheets(weekLabel,workers,activeDays,siteHours);
    setTimesheetRecords(prev=>{
      const other=prev.filter(t=>t.weekLabel!==weekLabel);
      const approved=prev.filter(t=>t.weekLabel===weekLabel&&t.status==="approved");
      const approvedIds=new Set(approved.map(t=>t.workerId));
      const merged=sheets.map(s=>approvedIds.has(s.workerId)?approved.find(a=>a.workerId===s.workerId):s);
      return [...other,...merged];
    });
  }

  function lockWeek(){
    if(!window.confirm("Lock WC "+selWeek+" and submit all timesheets for finance approval?\n\nThis action cannot be undone.")) return;
    setTimesheetRecords(prev=>prev.map(t=>
      t.weekLabel===selWeek&&t.status==="draft"
        ?{...t,status:"submitted",lockedAt:new Date().toISOString()}:t
    ));
  }

  function approveSheet(id){
    setTimesheetRecords(prev=>prev.map(t=>
      t.id===id?{...t,status:"approved",approvedAt:new Date().toISOString(),approvedBy:"Finance"}:t
    ));
  }
  function approveAll(){
    setTimesheetRecords(prev=>prev.map(t=>
      t.weekLabel===selWeek&&t.status==="submitted"
        ?{...t,status:"approved",approvedAt:new Date().toISOString(),approvedBy:"Finance"}:t
    ));
  }

  function updateDay(tsId,day,key,val){
    setTimesheetRecords(prev=>prev.map(t=>{
      if(t.id!==tsId) return t;
      const w=workers.find(x=>x.id===t.workerId);
      const newBd={...t.dayBreakdown,[day]:{...(t.dayBreakdown?.[day]||{}),[key]:Number(val)||0}};
      return recalcSheet({...t,dayBreakdown:newBd},w);
    }));
  }

  function createPayslips(){
    const approved=shown.filter(t=>t.status==="approved");
    if(approved.length===0){alert("No approved timesheets for this week.");return;}
    const existing=payslipRecords.filter(p=>p.weekLabel!==selWeek);
    const newPays=approved.map(t=>({
      id:"ps_"+t.workerId+"_"+selWeek.replace(/\s+/g,""),
      workerId:t.workerId, workerName:t.workerName,
      position:t.position, company:t.company||"",
      weekLabel:t.weekLabel, stdHours:t.stdHours, otHours:t.otHours,
      rate:t.rate, taxRate:t.taxRate, gross:t.gross, tax:t.tax, net:t.net,
      dayBreakdown:t.dayBreakdown, days:t.days,
      status:"pending",   // pending → issued
      timesheetId:t.id,
      approvedAt:t.approvedAt, approvedBy:t.approvedBy,
      issuedAt:null, createdAt:new Date().toISOString(),
    }));
    setPayslipRecords([...existing,...newPays]);
    // Mark timesheets as payslip generated
    setTimesheetRecords(prev=>prev.map(t=>approved.find(a=>a.id===t.id)?{...t,status:"payslip_generated"}:t));
    alert("✓ "+newPays.length+" payslips created for WC "+selWeek);
    setPage("payslips");
  }

  const allWeeks=[...new Set(timesheetRecords.map(t=>t.weekLabel))].sort((a,b)=>new Date(b)-new Date(a));
  if(!allWeeks.includes(weekLabel)) allWeeks.unshift(weekLabel);
  const shown=timesheetRecords.filter(t=>t.weekLabel===selWeek);
  const totGross=shown.reduce((a,t)=>a+t.gross,0);
  const totStd=shown.reduce((a,t)=>a+t.stdHours,0);
  const totOT=shown.reduce((a,t)=>a+t.otHours,0);

  // Week status
  const wkStatus=shown.length===0?"empty":
    shown.every(t=>t.status==="payslip_generated")?"payslip_generated":
    shown.every(t=>t.status==="approved"||(t.status==="payslip_generated"))?"approved":
    shown.some(t=>t.status==="submitted")?"submitted":"draft";

  const STATUS_STYLE={
    draft:   {color:"#64748b",bg:"#1e2535",border:"#2d3555",label:"Draft"},
    submitted:{color:"#fbbf24",bg:"#1a1500",border:"#92400e",label:"Submitted"},
    approved: {color:"#34d399",bg:"#0d2218",border:"#065f46",label:"Approved"},
    payslip_generated:{color:"#a78bfa",bg:"#1a0d2e",border:"#5b21b6",label:"Payslip Issued"},
  };

  return <div>
    <DPageHdr title="⏱ Timesheets" sub="Auto-generated from labour schedule · approve for payroll"
      actions={<div style={{display:"flex",gap:7,alignItems:"center"}}>
        {selWeek===weekLabel&&<button onClick={regenWeek} style={{padding:"6px 13px",background:"#1e2535",border:"1px solid #3b82f6",borderRadius:7,color:"#60a5fa",cursor:"pointer",fontSize:12,fontWeight:700}}>🔄 Re-sync from Schedule</button>}
        {wkStatus==="draft"&&shown.length>0&&<button onClick={lockWeek} style={{padding:"6px 13px",background:"#2d2008",border:"1px solid #f59e0b",borderRadius:7,color:"#fbbf24",cursor:"pointer",fontSize:12,fontWeight:700}}>🔒 Submit for Approval</button>}
        {wkStatus==="submitted"&&<button onClick={approveAll} style={{padding:"6px 13px",background:"#0d2218",border:"1px solid #10b981",borderRadius:7,color:"#34d399",cursor:"pointer",fontSize:12,fontWeight:700}}>✓ Approve All</button>}
        {wkStatus==="approved"&&<button onClick={createPayslips} style={{padding:"6px 13px",background:"linear-gradient(135deg,#7c3aed,#8b5cf6)",border:"none",borderRadius:7,color:"#fff",cursor:"pointer",fontSize:12,fontWeight:700}}>💷 Create Payslips →</button>}
      </div>}/>

    <div style={DS.body}>
      {/* Workflow banner */}
      <div style={{display:"flex",alignItems:"center",gap:0,marginBottom:18,background:"#111827",border:"1px solid #1e2535",borderRadius:10,overflow:"hidden"}}>
        {[
          ["1","Auto-generate","Schedule → Timesheets","draft","#64748b"],
          ["2","Finance Review","Check hours & pay","submitted","#fbbf24"],
          ["3","Approve","Finance signs off","approved","#34d399"],
          ["4","Issue Payslips","Sent to workers","payslip_generated","#a78bfa"],
        ].map(([n,label,sub,st,c],i,arr)=>{
          const active=wkStatus===st;
          const past=["draft","submitted","approved","payslip_generated"].indexOf(wkStatus)>i;
          return <div key={n} style={{flex:1,padding:"11px 14px",background:active?c+"18":past?c+"08":"transparent",borderRight:i<arr.length-1?"1px solid #1e2535":"none",textAlign:"center"}}>
            <div style={{width:24,height:24,borderRadius:6,background:active?c:past?c+"33":"#1e2535",display:"flex",alignItems:"center",justifyContent:"center",margin:"0 auto 5px",fontSize:11,fontWeight:800,color:active?"#fff":past?c:"#374151"}}>{past?"✓":n}</div>
            <div style={{fontSize:11,fontWeight:700,color:active?c:past?c:'"#374151"'}}>{label}</div>
            <div style={{fontSize:10,color:"#64748b",marginTop:2}}>{sub}</div>
          </div>;
        })}
      </div>

      {/* Week selector */}
      <div style={{display:"flex",gap:6,marginBottom:16,flexWrap:"wrap",alignItems:"center"}}>
        <span style={{fontSize:11,color:"#64748b",fontWeight:700,marginRight:4}}>Week:</span>
        {allWeeks.map(wk=>{
          const wkRecs=timesheetRecords.filter(t=>t.weekLabel===wk);
          const wkSt=wkRecs.length===0?"empty":wkRecs.every(t=>t.status==="payslip_generated")?"payslip_generated":wkRecs.every(t=>t.status==="approved"||(t.status==="payslip_generated"))?"approved":wkRecs.some(t=>t.status==="submitted")?"submitted":"draft";
          const sc=STATUS_STYLE[wkSt]||STATUS_STYLE.draft;
          return <button key={wk} onClick={()=>setSelWeek(wk)}
            style={{padding:"5px 12px",background:selWeek===wk?"#1e3a5f":"#1a1f2e",border:"1px solid "+(selWeek===wk?"#3b82f6":"#2d3555"),borderRadius:7,color:selWeek===wk?"#60a5fa":"#64748b",cursor:"pointer",fontSize:11,fontWeight:selWeek===wk?700:400,display:"flex",alignItems:"center",gap:5}}>
            WC {wk}
            {wkRecs.length>0&&<span style={{fontSize:9,padding:"1px 5px",borderRadius:4,background:sc.bg,color:sc.color,border:"1px solid "+sc.border,fontWeight:700}}>{sc.label}</span>}
          </button>;
        })}
      </div>

      {/* Empty state */}
      {shown.length===0&&<div style={{textAlign:"center",padding:48,border:"1px dashed #1e2535",borderRadius:12}}>
        <div style={{fontSize:32,marginBottom:12}}>⏱</div>
        <div style={{fontSize:15,fontWeight:700,color:"#94a3b8",marginBottom:6}}>No timesheets for WC {selWeek}</div>
        <div style={{fontSize:12,color:"#374151",marginBottom:16}}>Timesheets are auto-generated as soon as workers are assigned to sites in the Labour Schedule.</div>
        {selWeek===weekLabel&&<button onClick={regenWeek} style={{padding:"8px 18px",background:"linear-gradient(135deg,#3b82f6,#6366f1)",border:"none",borderRadius:8,color:"#fff",cursor:"pointer",fontSize:13,fontWeight:700}}>⚡ Generate Now from Current Schedule</button>}
      </div>}

      {shown.length>0&&<>
        {/* Summary cards */}
        <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:10,marginBottom:18}}>
          <DStat label="Operatives" value={shown.length} color="#60a5fa"/>
          <DStat label="Std Hours" value={totStd+"h"} color="#34d399"/>
          <DStat label="OT Hours" value={totOT>0?totOT+"h":"—"} color="#fbbf24"/>
          <DStat label="Gross Pay" value={"£"+totGross.toFixed(0)} color="#a78bfa"/>
          <DStat label="Week Status" value={(STATUS_STYLE[wkStatus]||STATUS_STYLE.draft).label} color={(STATUS_STYLE[wkStatus]||STATUS_STYLE.draft).color}/>
        </div>

        {/* Main table */}
        <div style={{border:"1px solid #1e2535",borderRadius:10,overflow:"hidden"}}>
          <table style={{width:"100%",borderCollapse:"collapse"}}>
            <thead><tr>
              <th style={DS.th}>Worker</th>
              <th style={DS.th}>Position</th>
              {(activeDays||BASE_DAYS).map(d=><th key={d} style={{...DS.th,textAlign:"center",color:d==="Sat"||d==="Sun"?"#fbbf24":"#64748b"}}>{d}</th>)}
              <th style={DS.th}>Std h</th>
              <th style={DS.th}>OT h</th>
              <th style={DS.th}>Rate</th>
              <th style={DS.th}>Gross</th>
              <th style={DS.th}>Tax</th>
              <th style={DS.th}>Net</th>
              <th style={DS.th}>Status</th>
              <th style={DS.th}>Actions</th>
            </tr></thead>
            <tbody>
              {shown.map((t,i)=>{
                const st=STATUS_STYLE[t.status]||STATUS_STYLE.draft;
                const isEditing=editId===t.id;
                const canEdit=t.status==="draft";
                return <tr key={t.id} style={{background:i%2===0?"#111827":"#0f1421"}}>
                  <td style={{...DS.td,fontWeight:600,color:"#f1f5f9"}}>
                    {t.workerName}
                    <div style={{fontSize:9,color:"#64748b"}}>{t.company}</div>
                  </td>
                  <td style={{...DS.td,color:"#94a3b8",fontSize:11}}>{t.position}</td>
                  {(activeDays||BASE_DAYS).map(d=>{
                    const bd=t.dayBreakdown?.[d];
                    const site=t.days?.[d]||"";
                    const sc=getSiteColor(site,allSites);
                    return <td key={d} style={{...DS.td,textAlign:"center",padding:"4px 3px"}}>
                      {bd?<div>
                        <div style={{display:"inline-block",padding:"1px 5px",borderRadius:3,fontSize:9,fontWeight:700,color:"#fff",background:sc,maxWidth:80,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={bd.site}>{bd.site.split("-")[0].trim()}</div>
                        {isEditing&&canEdit?<div style={{display:"flex",gap:1,marginTop:2,justifyContent:"center"}}>
                          <input type="number" value={bd.hours||0} onChange={e=>updateDay(t.id,d,"hours",e.target.value)} style={{width:28,background:"#0f1421",border:"1px solid #2d3555",borderRadius:3,padding:"1px 2px",color:"#34d399",fontSize:9,textAlign:"center",outline:"none"}} title="Std hrs"/>
                          <span style={{fontSize:9,color:"#374151"}}>+</span>
                          <input type="number" value={bd.ot||0} onChange={e=>updateDay(t.id,d,"ot",e.target.value)} style={{width:24,background:"#0f1421",border:"1px solid #2d3555",borderRadius:3,padding:"1px 2px",color:"#fbbf24",fontSize:9,textAlign:"center",outline:"none"}} title="OT hrs"/>
                        </div>:<div style={{fontSize:9,color:"#64748b",marginTop:1}}>{bd.hours}h{bd.ot>0?<span style={{color:"#fbbf24"}}>+{bd.ot}</span>:""}</div>}
                      </div>:<span style={{color:"#2d3555",fontSize:10}}>—</span>}
                    </td>;
                  })}
                  <td style={{...DS.td,color:"#34d399",fontWeight:700,textAlign:"center"}}>{t.stdHours}h</td>
                  <td style={{...DS.td,color:"#fbbf24",fontWeight:700,textAlign:"center"}}>{t.otHours>0?t.otHours+"h":"—"}</td>
                  <td style={{...DS.td,color:"#34d399"}}>{t.rate?`£${t.rate}/hr`:"—"}</td>
                  <td style={{...DS.td,color:"#34d399",fontWeight:700}}>£{t.gross.toFixed(2)}</td>
                  <td style={{...DS.td,color:"#f87171",fontSize:11}}>£{t.tax.toFixed(2)}</td>
                  <td style={{...DS.td,color:"#a78bfa",fontWeight:800}}>£{t.net.toFixed(2)}</td>
                  <td style={DS.td}>
                    <span style={{padding:"2px 8px",borderRadius:5,fontSize:10,fontWeight:700,color:st.color,background:st.bg,border:"1px solid "+st.border,whiteSpace:"nowrap"}}>{st.label}</span>
                  </td>
                  <td style={DS.td}>
                    <div style={{display:"flex",gap:3,flexWrap:"nowrap"}}>
                      {canEdit&&<button onClick={()=>setEditId(isEditing?null:t.id)}
                        style={{padding:"3px 7px",background:isEditing?"#1e3a5f":"#1a1f2e",border:"1px solid "+(isEditing?"#3b82f6":"#2d3555"),borderRadius:4,color:isEditing?"#60a5fa":"#64748b",cursor:"pointer",fontSize:10}}>
                        {isEditing?"✓":"Edit"}
                      </button>}
                      {t.status==="submitted"&&<button onClick={()=>approveSheet(t.id)}
                        style={{padding:"3px 7px",background:"#0d2218",border:"1px solid #10b981",borderRadius:4,color:"#34d399",cursor:"pointer",fontSize:10,fontWeight:700}}>✓ Approve</button>}
                      <button onClick={()=>exportPayslip({id:t.workerId,name:t.workerName,position:t.position,agreedRate:t.rate,taxRate:t.taxRate,overtimeMultiplier:1.5,customOTRate:null,days:t.days||{},overtimeHours:{},hoursPerDay:{}},activeDays,t.weekLabel,siteHours||{})}
                        style={{padding:"3px 6px",background:"#0d2218",border:"1px solid #10b981",borderRadius:4,color:"#34d399",cursor:"pointer",fontSize:10}}>💷</button>
                    </div>
                  </td>
                </tr>;
              })}
            </tbody>
            <tfoot><tr style={{background:"#0d1117",borderTop:"2px solid #2d3555"}}>
              <td colSpan={2+(activeDays||BASE_DAYS).length} style={{...DS.td,fontWeight:700,color:"#94a3b8"}}>TOTALS — {shown.length} operatives</td>
              <td style={{...DS.td,color:"#34d399",fontWeight:800,textAlign:"center"}}>{totStd}h</td>
              <td style={{...DS.td,color:"#fbbf24",fontWeight:800,textAlign:"center"}}>{totOT>0?totOT+"h":"—"}</td>
              <td style={DS.td}/>
              <td style={{...DS.td,color:"#34d399",fontWeight:800}}>£{totGross.toFixed(2)}</td>
              <td style={{...DS.td,color:"#f87171",fontWeight:700}}>£{shown.reduce((a,t)=>a+t.tax,0).toFixed(2)}</td>
              <td style={{...DS.td,color:"#a78bfa",fontWeight:900,fontSize:13}}>£{shown.reduce((a,t)=>a+t.net,0).toFixed(2)}</td>
              <td colSpan={2} style={DS.td}/>
            </tr></tfoot>
          </table>
        </div>
      </>}
    </div>
  </div>;
}

// ═══════════════════════════════════════════════════════════════════════════
// PAYSLIPS — created from approved timesheets, issued to worker portal
// ═══════════════════════════════════════════════════════════════════════════
function DPayslips({workers,allSites,activeDays,siteHours,weekLabel,payslipRecords,setPayslipRecords,timesheetRecords,generatePayslips,setPage}){
  const allWeeks=[...new Set(payslipRecords.map(p=>p.weekLabel))].sort((a,b)=>new Date(b)-new Date(a));
  const [selWeek,setSelWeek]=useState(payslipRecords[0]?.weekLabel||weekLabel);
  const [viewPayslip,setViewPayslip]=useState(null);
  const shown=payslipRecords.filter(p=>p.weekLabel===selWeek);
  const totGross=shown.reduce((a,p)=>a+p.gross,0);
  const totNet=shown.reduce((a,p)=>a+p.net,0);
  const totTax=shown.reduce((a,p)=>a+p.tax,0);

  function issueAll(){
    if(!window.confirm("Mark all payslips for WC "+selWeek+" as issued to workers?\n\nThis simulates sending to the worker portal.")) return;
    setPayslipRecords(recs=>recs.map(p=>
      p.weekLabel===selWeek&&p.status==="pending"
        ?{...p,status:"issued",issuedAt:new Date().toISOString()}:p
    ));
    alert("✓ "+shown.filter(p=>p.status==="pending").length+" payslips marked as issued.");
  }

  function issueOne(id){
    setPayslipRecords(recs=>recs.map(p=>
      p.id===id?{...p,status:"issued",issuedAt:new Date().toISOString()}:p
    ));
  }

  // Worker Portal preview - opens a printable payslip with portal styling
  function openPortal(p){
    const w=workers.find(x=>x.id===p.workerId)||{name:p.workerName,position:p.position,agreedRate:p.rate,taxRate:p.taxRate,days:p.days||{},overtimeHours:{},hoursPerDay:{},overtimeMultiplier:1.5,customOTRate:null};
    exportPayslip(w,activeDays,p.weekLabel,siteHours||{});
  }

  const allPending=shown.filter(p=>p.status==="pending").length;

  return <div>
    <DPageHdr title="💷 Payroll & Payslips" sub={"Total issued: "+payslipRecords.filter(p=>p.status==="issued").length+" · Pending: "+payslipRecords.filter(p=>p.status==="pending").length}
      actions={<div style={{display:"flex",gap:7}}>
        {allPending>0&&<button onClick={issueAll} style={{padding:"6px 13px",background:"linear-gradient(135deg,#059669,#10b981)",border:"none",borderRadius:7,color:"#fff",cursor:"pointer",fontSize:12,fontWeight:700}}>📤 Issue All to Workers ({allPending})</button>}
      </div>}/>

    <div style={DS.body}>
      {/* Workflow reminder */}
      <div style={{background:"#0c1a0c",border:"1px solid #065f46",borderRadius:9,padding:"10px 16px",marginBottom:16,fontSize:11,color:"#34d399",display:"flex",alignItems:"center",gap:9}}>
        <span style={{fontSize:16}}>ℹ️</span>
        <span><strong>Payslip workflow:</strong> Timesheets are auto-generated → approved by finance → payslips created automatically → issued to worker portal. Go to <span onClick={()=>setPage("timesheets")} style={{color:"#60a5fa",cursor:"pointer",textDecoration:"underline"}}>Timesheets</span> to manage approval.</span>
      </div>

      {/* Week tabs */}
      <div style={{display:"flex",gap:6,marginBottom:16,flexWrap:"wrap",alignItems:"center"}}>
        <span style={{fontSize:11,color:"#64748b",fontWeight:700}}>Week:</span>
        {payslipRecords.length===0&&<span style={{color:"#374151",fontSize:12}}>No payslips yet — approve timesheets first.</span>}
        {allWeeks.map(wk=>{
          const wkP=payslipRecords.filter(p=>p.weekLabel===wk);
          const allIssued=wkP.length>0&&wkP.every(p=>p.status==="issued");
          return <button key={wk} onClick={()=>setSelWeek(wk)}
            style={{padding:"5px 12px",background:selWeek===wk?"#1e3a5f":"#1a1f2e",border:"1px solid "+(selWeek===wk?"#3b82f6":"#2d3555"),borderRadius:7,color:selWeek===wk?"#60a5fa":"#64748b",cursor:"pointer",fontSize:11,fontWeight:selWeek===wk?700:400,display:"flex",alignItems:"center",gap:5}}>
            WC {wk}
            <span style={{fontSize:9,padding:"1px 5px",borderRadius:4,background:allIssued?"#0d2218":"#1a1500",color:allIssued?"#34d399":"#fbbf24",border:"1px solid "+(allIssued?"#065f46":"#92400e"),fontWeight:700}}>
              {allIssued?"✓ Issued":wkP.filter(p=>p.status==="pending").length+" pending"}
            </span>
          </button>;
        })}
      </div>

      {shown.length===0&&selWeek&&<div style={{textAlign:"center",padding:48,border:"1px dashed #1e2535",borderRadius:12}}>
        <div style={{fontSize:32,marginBottom:12}}>💷</div>
        <div style={{fontSize:15,fontWeight:700,color:"#94a3b8",marginBottom:6}}>No payslips for WC {selWeek}</div>
        <div style={{fontSize:12,color:"#374151",marginBottom:14}}>Approve timesheets for this week to generate payslips automatically.</div>
        <button onClick={()=>setPage("timesheets")} style={{padding:"8px 18px",background:"linear-gradient(135deg,#3b82f6,#6366f1)",border:"none",borderRadius:8,color:"#fff",cursor:"pointer",fontSize:13,fontWeight:700}}>⏱ Go to Timesheets →</button>
      </div>}

      {shown.length>0&&<>
        {/* Summary */}
        <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:10,marginBottom:18}}>
          <DStat label="Payslips" value={shown.length} color="#60a5fa"/>
          <DStat label="Gross Pay" value={"£"+totGross.toFixed(0)} color="#34d399"/>
          <DStat label="Tax" value={"£"+totTax.toFixed(0)} color="#f87171"/>
          <DStat label="Net Pay" value={"£"+totNet.toFixed(0)} color="#a78bfa"/>
          <DStat label="Issued" value={shown.filter(p=>p.status==="issued").length+"/"+shown.length} color={shown.every(p=>p.status==="issued")?"#34d399":"#fbbf24"}/>
        </div>

        {/* Payslip cards grid */}
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(280px,1fr))",gap:12,marginBottom:20}}>
          {shown.map(p=>{
            const issued=p.status==="issued";
            const w=workers.find(x=>x.id===p.workerId);
            const initials=(p.workerName||"?").split(" ").map(n=>n[0]).join("").slice(0,2).toUpperCase();
            const color=w?.color||"#3b82f6";
            return <div key={p.id} style={{background:"linear-gradient(145deg,#141924,#1a2035)",border:"1px solid "+(issued?"#065f4666":"#1e2535"),borderRadius:12,padding:16,position:"relative",overflow:"hidden"}}>
              {/* Issued banner */}
              {issued&&<div style={{position:"absolute",top:10,right:-20,background:"#059669",color:"#fff",fontSize:9,fontWeight:800,padding:"3px 24px",transform:"rotate(30deg)",letterSpacing:"0.08em"}}>ISSUED</div>}
              <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:12}}>
                <div style={{width:38,height:38,borderRadius:10,background:color+"22",border:"1px solid "+color+"44",display:"flex",alignItems:"center",justifyContent:"center",fontSize:14,fontWeight:800,color,flexShrink:0}}>{initials}</div>
                <div style={{flex:1,minWidth:0}}>
                  <div style={{fontWeight:700,color:"#f1f5f9",fontSize:13,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.workerName}</div>
                  <div style={{fontSize:10,color:"#64748b"}}>{p.position} · {p.company}</div>
                </div>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:7,marginBottom:12}}>
                {[["Gross","£"+p.gross.toFixed(0),"#34d399"],["Tax","-£"+p.tax.toFixed(0),"#f87171"],["Net","£"+p.net.toFixed(0),"#a78bfa"]].map(([l,v,c])=>(
                  <div key={l} style={{background:"#0f1421",borderRadius:7,padding:"6px 8px",textAlign:"center"}}>
                    <div style={{fontSize:9,color:"#64748b",textTransform:"uppercase"}}>{l}</div>
                    <div style={{fontSize:13,fontWeight:800,color:c,marginTop:2}}>{v}</div>
                  </div>
                ))}
              </div>
              <div style={{fontSize:10,color:"#64748b",marginBottom:10,display:"flex",gap:8}}>
                <span>⏱ {p.stdHours}h std</span>
                {p.otHours>0&&<span style={{color:"#fbbf24"}}>+{p.otHours}h OT</span>}
                <span>£{p.rate||0}/hr</span>
              </div>
              <div style={{display:"flex",gap:5}}>
                <button onClick={()=>openPortal(p)} style={{flex:1,padding:"6px 0",background:"#1e3a5f",border:"1px solid #3b82f6",borderRadius:6,color:"#60a5fa",cursor:"pointer",fontSize:11,fontWeight:700}}>💷 View Payslip</button>
                {!issued&&<button onClick={()=>issueOne(p.id)} style={{flex:1,padding:"6px 0",background:"#0d2218",border:"1px solid #10b981",borderRadius:6,color:"#34d399",cursor:"pointer",fontSize:11,fontWeight:700}}>📤 Issue</button>}
                {issued&&<div style={{flex:1,padding:"6px 0",textAlign:"center",fontSize:10,color:"#34d399",display:"flex",alignItems:"center",justifyContent:"center",gap:3}}>✓ {p.issuedAt?new Date(p.issuedAt).toLocaleDateString("en-GB",{day:"2-digit",month:"short"}):""}</div>}
              </div>
              {issued&&p.issuedAt&&<div style={{marginTop:6,fontSize:9,color:"#374151",textAlign:"center"}}>Issued {new Date(p.issuedAt).toLocaleString("en-GB",{day:"2-digit",month:"short",hour:"2-digit",minute:"2-digit"})}</div>}
            </div>;
          })}
        </div>

        {/* Payslip table */}
        <div style={{border:"1px solid #1e2535",borderRadius:10,overflow:"hidden"}}>
          <div style={{padding:"10px 14px",background:"#0d1117",fontSize:11,color:"#64748b",fontWeight:700,textTransform:"uppercase",letterSpacing:"0.06em",borderBottom:"1px solid #1e2535"}}>Payroll Summary Table</div>
          <table style={{width:"100%",borderCollapse:"collapse"}}>
            <thead><tr>
              <th style={DS.th}>Worker</th><th style={DS.th}>WC</th>
              <th style={DS.th}>Std h</th><th style={DS.th}>OT h</th>
              <th style={DS.th}>Rate</th><th style={DS.th}>Gross</th>
              <th style={DS.th}>Tax</th><th style={DS.th}>Net Pay</th>
              <th style={DS.th}>Status</th><th style={DS.th}>Actions</th>
            </tr></thead>
            <tbody>{shown.map((p,i)=>(
              <tr key={p.id} style={{background:i%2===0?"#111827":"#0f1421"}}>
                <td style={{...DS.td,fontWeight:600,color:"#f1f5f9"}}>{p.workerName}<div style={{fontSize:10,color:"#64748b"}}>{p.position}</div></td>
                <td style={{...DS.td,color:"#94a3b8",fontSize:11}}>{p.weekLabel}</td>
                <td style={{...DS.td,color:"#60a5fa",fontWeight:600,textAlign:"center"}}>{p.stdHours}h</td>
                <td style={{...DS.td,color:"#fbbf24",textAlign:"center"}}>{p.otHours>0?p.otHours+"h":"—"}</td>
                <td style={{...DS.td,color:"#34d399"}}>{p.rate?`£${p.rate}/hr`:"—"}</td>
                <td style={{...DS.td,color:"#34d399",fontWeight:700}}>£{p.gross.toFixed(2)}</td>
                <td style={{...DS.td,color:"#f87171"}}>£{p.tax.toFixed(2)}</td>
                <td style={{...DS.td,color:"#a78bfa",fontWeight:800,fontSize:13}}>£{p.net.toFixed(2)}</td>
                <td style={DS.td}><span style={{padding:"2px 8px",borderRadius:4,fontSize:10,fontWeight:700,color:p.status==="issued"?"#34d399":"#fbbf24",background:p.status==="issued"?"#0d2218":"#1a1500",border:"1px solid "+(p.status==="issued"?"#065f46":"#92400e")}}>{p.status==="issued"?"✓ Issued":"⏳ Pending"}</span></td>
                <td style={DS.td}><div style={{display:"flex",gap:4}}>
                  <button onClick={()=>openPortal(p)} style={{padding:"3px 7px",background:"#0d2218",border:"1px solid #10b981",borderRadius:4,color:"#34d399",cursor:"pointer",fontSize:10,fontWeight:700}}>💷 PDF</button>
                  {p.status==="pending"&&<button onClick={()=>issueOne(p.id)} style={{padding:"3px 7px",background:"#1e3a5f",border:"1px solid #3b82f6",borderRadius:4,color:"#60a5fa",cursor:"pointer",fontSize:10}}>📤</button>}
                  <button onClick={()=>setPayslipRecords(recs=>recs.filter(x=>x.id!==p.id))} style={{padding:"3px 7px",background:"#2d1515",border:"1px solid #ef4444",borderRadius:4,color:"#f87171",cursor:"pointer",fontSize:10}}>✕</button>
                </div></td>
              </tr>
            ))}</tbody>
            <tfoot><tr style={{background:"#0d1117",borderTop:"2px solid #2d3555"}}>
              <td colSpan={4} style={{...DS.td,fontWeight:700,color:"#94a3b8"}}>TOTALS — {shown.length} payslips</td>
              <td style={DS.td}/>
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
// EXPENSES — categorised transactions list + site cost summary
// ═══════════════════════════════════════════════════════════════════════════
const EXPENSE_CATS_LIST=["Materials","Plant Hire","Subcontractor","Labour (External)","Transport","Insurance","Tools & Equipment","Professional Fees","Utilities","Office","Other Expense"];
const INCOME_CATS_LIST=["Client Payment","Contract Payment","Variation Payment","Retention Release","Other Income"];

function DExpenses({bankTransactions,allSites,clients,workers,activeDays,siteHours,setPage}){
  const [typeF,setTypeF]=useState("all");
  const [catF,setCatF]=useState("");
  const [siteF,setSiteF]=useState("");
  const [srch,setSrch]=useState("");
  const txns=bankTransactions||[];
  const shown=txns.filter(t=>{
    if(typeF!=="all"&&t.type!==typeF) return false;
    if(catF&&t.category!==catF) return false;
    if(siteF&&t.siteId!==siteF) return false;
    if(srch&&!(t.description||"").toLowerCase().includes(srch.toLowerCase())&&!(t.category||"").toLowerCase().includes(srch.toLowerCase())) return false;
    return true;
  });
  const totInc=txns.filter(t=>t.type==="income").reduce((a,t)=>a+Math.abs(t.amount),0);
  const totExp=txns.filter(t=>t.type==="expense").reduce((a,t)=>a+Math.abs(t.amount),0);
  const totShown=shown.reduce((a,t)=>a+(t.type==="income"?1:-1)*Math.abs(t.amount),0);
  // Category breakdown
  const byCat={};txns.filter(t=>t.type==="expense").forEach(t=>{const c=t.category||"Uncategorised";byCat[c]=(byCat[c]||0)+Math.abs(t.amount);});
  const catRows=Object.entries(byCat).sort((a,b)=>b[1]-a[1]);
  const maxC=catRows[0]?.[1]||1;
  // Site expense totals (bank + labour)
  const bankBySite={};txns.filter(t=>t.type==="expense"&&t.siteId).forEach(t=>{bankBySite[t.siteId]=(bankBySite[t.siteId]||0)+Math.abs(t.amount);});
  const labourBySite={};workers.forEach(w=>{const{bd}=calcPay(w,activeDays,siteHours);Object.values(bd).forEach(b=>{labourBySite[b.site]=(labourBySite[b.site]||0)+b.gross;});});
  const actSites=allSites.filter(s=>!isOff(s.name));

  return <div>
    <DPageHdr title="💸 Expenses & Costs" sub={txns.length+" transactions · £"+Math.round(totExp).toLocaleString()+" expenses"}
      actions={<button onClick={()=>setPage("bank")} style={{padding:"6px 14px",background:"#1e3a5f",border:"1px solid #3b82f6",borderRadius:7,color:"#60a5fa",cursor:"pointer",fontSize:12,fontWeight:700}}>🏦 Import Bank Statement</button>}/>
    <div style={DS.body}>
      {txns.length===0?<div style={{textAlign:"center",padding:60,border:"1px dashed #1e2535",borderRadius:12}}>
        <div style={{fontSize:40,marginBottom:12}}>💸</div>
        <div style={{fontSize:16,fontWeight:700,color:"#f1f5f9",marginBottom:6}}>No transactions yet</div>
        <div style={{fontSize:12,color:"#64748b",marginBottom:16}}>Import your bank statement to categorise income and expenses.</div>
        <button onClick={()=>setPage("bank")} style={{padding:"9px 22px",background:"linear-gradient(135deg,#3b82f6,#6366f1)",border:"none",borderRadius:8,color:"#fff",cursor:"pointer",fontSize:13,fontWeight:700}}>🏦 Import Bank Statement</button>
      </div>:<>
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12,marginBottom:22}}>
        <DStat label="Total Income" value={"£"+Math.round(totInc).toLocaleString()} color="#34d399" sub={txns.filter(t=>t.type==="income").length+" txns"}/>
        <DStat label="Total Expenses" value={"£"+Math.round(totExp).toLocaleString()} color="#f87171" sub={txns.filter(t=>t.type==="expense").length+" txns"}/>
        <DStat label="Net Position" value={"£"+Math.round(totInc-totExp).toLocaleString()} color={(totInc-totExp)>=0?"#34d399":"#f87171"} sub="Income minus expenses"/>
        <DStat label="Filtered Total" value={"£"+Math.abs(Math.round(totShown)).toLocaleString()} color="#fbbf24" sub={shown.length+" records"}/>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginBottom:22}}>
        <div style={{background:"#111827",border:"1px solid #1e2535",borderRadius:11,padding:16}}>
          <div style={{fontSize:11,color:"#f87171",fontWeight:700,textTransform:"uppercase",marginBottom:14}}>Expenses by Category</div>
          {catRows.length===0?<div style={{color:"#374151",fontSize:12}}>No categorised expenses yet.</div>:
          catRows.map(([cat,amt])=>(
            <div key={cat} style={{marginBottom:9}}>
              <div style={{display:"flex",justifyContent:"space-between",marginBottom:3}}>
                <span style={{fontSize:11,color:"#94a3b8"}}>{cat}</span>
                <span style={{fontSize:11,color:"#f87171",fontWeight:700}}>£{Math.round(amt).toLocaleString()}</span>
              </div>
              <div style={{height:6,background:"#1e2535",borderRadius:3,overflow:"hidden"}}>
                <div style={{height:"100%",borderRadius:3,background:"linear-gradient(90deg,#dc2626,#f87171)",width:(amt/maxC*100)+"%"}}/>
              </div>
            </div>
          ))}
        </div>
        <div style={{background:"#111827",border:"1px solid #1e2535",borderRadius:11,padding:16}}>
          <div style={{fontSize:11,color:"#fbbf24",fontWeight:700,textTransform:"uppercase",marginBottom:14}}>Total Cost per Site (Bank + Labour)</div>
          {actSites.map(site=>{
            const bk=bankBySite[site.id]||0;
            const lb=labourBySite[site.name]||0;
            if(bk+lb===0) return null;
            return <div key={site.id} style={{marginBottom:10,padding:"9px 11px",background:"#0f1421",borderRadius:8,border:"1px solid "+site.color+"33"}}>
              <div style={{display:"flex",alignItems:"center",gap:7,marginBottom:4}}>
                <span style={{width:8,height:8,borderRadius:"50%",background:site.color,flexShrink:0}}/>
                <span style={{fontSize:12,fontWeight:700,color:site.color,flex:1}}>{site.name}</span>
                <span style={{fontSize:14,fontWeight:800,color:"#f87171"}}>£{Math.round(bk+lb).toLocaleString()}</span>
              </div>
              <div style={{display:"flex",gap:14,fontSize:10,color:"#64748b"}}>
                {lb>0&&<span>Labour: <span style={{color:"#f87171",fontWeight:600}}>£{Math.round(lb).toLocaleString()}</span></span>}
                {bk>0&&<span>Other: <span style={{color:"#fbbf24",fontWeight:600}}>£{Math.round(bk).toLocaleString()}</span></span>}
              </div>
            </div>;
          })}
        </div>
      </div>
      <div style={{display:"flex",gap:9,marginBottom:14,flexWrap:"wrap",alignItems:"center"}}>
        {[["all","All"],["income","Income"],["expense","Expense"]].map(([v,l])=>(
          <button key={v} onClick={()=>setTypeF(v)} style={{padding:"5px 12px",background:typeF===v?"#1e3a5f":"#1a1f2e",border:"1px solid "+(typeF===v?"#3b82f6":"#2d3555"),borderRadius:7,color:typeF===v?"#60a5fa":"#64748b",cursor:"pointer",fontSize:11,fontWeight:typeF===v?700:400}}>{l}</button>
        ))}
        <select value={catF} onChange={e=>setCatF(e.target.value)} style={{background:"#1a1f2e",border:"1px solid #2d3555",borderRadius:7,padding:"5px 9px",color:catF?"#e2e8f0":"#64748b",fontSize:11,outline:"none",cursor:"pointer"}}>
          <option value="">All Categories</option>
          <optgroup label="Income">{INCOME_CATS_LIST.map(c=><option key={c} value={c}>{c}</option>)}</optgroup>
          <optgroup label="Expenses">{EXPENSE_CATS_LIST.map(c=><option key={c} value={c}>{c}</option>)}</optgroup>
        </select>
        <select value={siteF} onChange={e=>setSiteF(e.target.value)} style={{background:"#1a1f2e",border:"1px solid #2d3555",borderRadius:7,padding:"5px 9px",color:siteF?"#e2e8f0":"#64748b",fontSize:11,outline:"none",cursor:"pointer"}}>
          <option value="">All Sites</option>
          {actSites.map(s=><option key={s.id} value={s.id}>{s.name}</option>)}
        </select>
        <input value={srch} onChange={e=>setSrch(e.target.value)} placeholder="🔍 Search…" style={{background:"#1a1f2e",border:"1px solid #2d3555",borderRadius:7,padding:"5px 9px",color:"#e2e8f0",fontSize:11,outline:"none",width:160}}/>
        {(typeF!=="all"||catF||siteF||srch)&&<button onClick={()=>{setTypeF("all");setCatF("");setSiteF("");setSrch("");}} style={{padding:"5px 9px",background:"#2d1515",border:"1px solid #ef4444",borderRadius:7,color:"#f87171",cursor:"pointer",fontSize:11,fontWeight:700}}>✕</button>}
      </div>
      <div style={{border:"1px solid #1e2535",borderRadius:10,overflow:"hidden"}}>
        <table style={{width:"100%",borderCollapse:"collapse"}}>
          <thead><tr>
            <th style={{...DS.th,minWidth:80}}>Date</th>
            <th style={{...DS.th,minWidth:200}}>Description</th>
            <th style={{...DS.th,minWidth:90}}>Amount</th>
            <th style={{...DS.th,minWidth:80}}>Type</th>
            <th style={{...DS.th,minWidth:140}}>Category</th>
            <th style={{...DS.th,minWidth:110}}>Site</th>
            <th style={{...DS.th,minWidth:100}}>Client</th>
          </tr></thead>
          <tbody>
            {shown.length===0&&<tr><td colSpan={7} style={{...DS.td,textAlign:"center",color:"#374151",padding:28}}>No records match filters.</td></tr>}
            {shown.map((t,i)=>{
              const site=allSites.find(s=>s.id===t.siteId);
              const client=clients.find(c=>c.id===t.clientId);
              const inc=t.type==="income";
              return <tr key={t.id||i} style={{background:i%2===0?"#111827":"#0f1421"}}>
                <td style={{...DS.td,color:"#94a3b8",fontSize:11,whiteSpace:"nowrap"}}>{t.date}</td>
                <td style={{...DS.td,maxWidth:200}}><div style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",fontSize:12,color:"#e2e8f0"}} title={t.description}>{t.description||"—"}</div>{t.notes&&<div style={{fontSize:9,color:"#64748b"}}>{t.notes}</div>}</td>
                <td style={{...DS.td,fontWeight:700,color:inc?"#34d399":"#f87171",whiteSpace:"nowrap"}}>{inc?"+":"-"}£{Math.abs(t.amount).toFixed(2)}</td>
                <td style={DS.td}><span style={{padding:"2px 8px",borderRadius:4,fontSize:10,fontWeight:700,color:inc?"#34d399":"#f87171",background:inc?"#0d2218":"#2d1515"}}>{inc?"Income":"Expense"}</span></td>
                <td style={{...DS.td,fontSize:11,color:"#fbbf24"}}>{t.category||<span style={{color:"#374151"}}>—</span>}</td>
                <td style={DS.td}>{site?<span style={{padding:"1px 7px",borderRadius:4,fontSize:10,fontWeight:600,color:"#fff",background:site.color}}>{site.name.split("-")[0].trim()}</span>:<span style={{color:"#374151",fontSize:11}}>—</span>}</td>
                <td style={{...DS.td,fontSize:11,color:client?.color||"#64748b"}}>{client?.name||"—"}</td>
              </tr>;
            })}
          </tbody>
          {shown.length>0&&<tfoot><tr style={{background:"#0d1117",borderTop:"2px solid #2d3555"}}>
            <td colSpan={2} style={{...DS.td,fontWeight:700,color:"#94a3b8"}}>{shown.length} RECORDS</td>
            <td style={{...DS.td,fontWeight:800,color:totShown>=0?"#34d399":"#f87171"}}>{totShown>=0?"+":"-"}£{Math.abs(Math.round(totShown)).toLocaleString()}</td>
            <td colSpan={4} style={DS.td}/>
          </tr></tfoot>}
        </table>
      </div>
    </>}
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
    const html="<!DOCTYPE html><html><head><meta charset='utf-8'/><title>"+pa.number+" — "+(site?.name||"")+"<\/title>"+
"<style>*{margin:0;padding:0;box-sizing:border-box;}body{background:#0d1117;color:#e2e8f0;font-family:'Segoe UI',Arial,sans-serif;font-size:12px;padding:24px;}"+
".hdr{display:flex;justify-content:space-between;margin-bottom:20px;padding-bottom:14px;border-bottom:2px solid #1e2535;}"+
"table{width:100%;border-collapse:collapse;margin-bottom:16px;}"+
"th{padding:7px 10px;text-align:left;font-size:10px;font-weight:700;color:#64748b;text-transform:uppercase;border-bottom:2px solid #1e2535;background:#0a0e17;}"+
"td{padding:6px 10px;border-bottom:1px solid #1a2030;font-size:11px;}"+
"tr:nth-child(even) td{background:#111827;}tr:nth-child(odd) td{background:#0f1421;}"+
".total{background:#0d2218;border:2px solid #10b981;border-radius:10px;padding:16px 20px;margin-top:10px;}"+
".ft{margin-top:16px;padding-top:10px;border-top:1px solid #1e2535;display:flex;justify-content:space-between;font-size:9px;color:#374151;}"+
"@media print{@page{margin:8mm;size:A3 landscape;}}<\/style><\/head><body>"+
"<div class='hdr'>"+
"  <div><div style='font-size:22px;font-weight:800;color:#f1f5f9'>"+pa.number+"<\/div>"+
"  <div style='font-size:13px;color:#64748b'>"+(site?.name||"—")+" · "+(client?.name||"—")+"<\/div>"+
"  <div style='font-size:11px;color:#64748b;margin-top:4px'>Date: "+pa.date+" · Status: "+pa.status+"<\/div><\/div>"+
"  <div style='text-align:right'><div style='font-size:11px;color:#64748b'>Contract Value<\/div><div style='font-size:24px;font-weight:900;color:#60a5fa'>£"+totalContract.toLocaleString(undefined,{maximumFractionDigits:0})+"<\/div><\/div>"+
"<\/div>"+
"<table><thead><tr><th>Type<\/th><th style='width:35%'>Description<\/th><th>Unit<\/th><th>Qty<\/th><th>Rate £<\/th><th>Contract £<\/th><th>Prev Qty<\/th><th>Claimed ToDate<\/th><th>This Period £<\/th><th>%<\/th><\/tr><\/thead><tbody>"+
items.map(it=>{const cl=calcClaimed(it);const pr=calcPrev(it);const th=cl-pr;const pct=it.contractQty*it.contractRate>0?Math.round((cl/(it.contractQty*it.contractRate))*100):0;
  return "<tr><td style='color:#a78bfa'>"+it.type+"<\/td><td style='font-weight:600'>"+(it.description||"—")+"<\/td><td>"+it.unit+"<\/td><td style='text-align:right'>"+it.contractQty+"<\/td><td style='text-align:right'>£"+it.contractRate.toLocaleString()+"<\/td><td style='text-align:right;font-weight:700'>£"+(it.contractQty*it.contractRate).toLocaleString()+"<\/td><td style='text-align:right;color:#64748b'>"+(it.previousQty||0)+"<\/td><td style='text-align:right;color:#34d399;font-weight:700'>£"+cl.toLocaleString(undefined,{maximumFractionDigits:0})+"<\/td><td style='text-align:right;color:"+(th>=0?"#34d399":"#f87171")+";font-weight:700'>£"+th.toLocaleString(undefined,{maximumFractionDigits:0})+"<\/td><td style='text-align:right'>"+pct+"%<\/td><\/tr>";
}).join("")+
"<\/tbody><\/table>"+
"<div class='total' style='display:grid;grid-template-columns:repeat(4,1fr);gap:16px'>"+
"  <div><div style='font-size:10px;color:#64748b;text-transform:uppercase'>Contract<\/div><div style='font-size:20px;font-weight:900;color:#60a5fa'>£"+totalContract.toLocaleString(undefined,{maximumFractionDigits:0})+"<\/div><\/div>"+
"  <div><div style='font-size:10px;color:#64748b;text-transform:uppercase'>Claimed To Date<\/div><div style='font-size:20px;font-weight:900;color:#34d399'>£"+totalClaimed.toLocaleString(undefined,{maximumFractionDigits:0})+"<\/div><\/div>"+
"  <div><div style='font-size:10px;color:#64748b;text-transform:uppercase'>This Period<\/div><div style='font-size:20px;font-weight:900;color:#fbbf24'>£"+totalThis.toLocaleString(undefined,{maximumFractionDigits:0})+"<\/div><\/div>"+
"  <div><div style='font-size:10px;color:#64748b;text-transform:uppercase'>% Complete<\/div><div style='font-size:20px;font-weight:900;color:#a78bfa'>"+pctComplete+"%<\/div><\/div>"+
"<\/div>"+
"<div class='ft'><span>"+pa.number+" — "+(site?.name||"")+"<\/span><span>"+(client?.name||"")+"<\/span><span>Bright Metalwork Ltd<\/span><span>"+new Date().toLocaleDateString("en-GB")+"<\/span><\/div>"+
"<script>window.onload=function(){window.print();}<\/script><\/body><\/html>";
    const blob=new Blob([html],{type:"text/html"});const url=URL.createObjectURL(blob);
    const win=window.open(url,"_blank","width=1200,height=820");
    if(!win){const a=document.createElement("a");a.href=url;a.download=pa.number+".html";a.click();}
    setTimeout(()=>URL.revokeObjectURL(url),6000);
  }

  function exportExcel(){
    // Colour helpers (ARGB)
    const COL={
      headerBg:"FF0A0E17", headerFg:"FF94A3B8",
      scopeBg:"FF0D1A2E",  scopeFg:"FF60A5FA",
      varBg:"FF1A1200",    varFg:"FFFBBF24",
      dayBg:"FF1A0D00",    dayFg:"FFF97316",
      prelBg:"FF1A0D2E",   prelFg:"FFA78BFA",
      totalBg:"FF0D2218",  totalFg:"FF34D399",
      contractFg:"FF34D399",
      prevFg:"FF94A3B8",
      thisFg:"FFFBBF24",
      pctFg:"FFA78BFA",
      white:"FFF1F5F9",    dark:"FF0D1117",
      border:"FF2D3555",
    };
    const typeColMap={"Scope of Work":{bg:COL.scopeBg,fg:COL.scopeFg},"Variation":{bg:COL.varBg,fg:COL.varFg},"Dayworks":{bg:COL.dayBg,fg:COL.dayFg},"Preliminaries":{bg:COL.prelBg,fg:COL.prelFg},"Other":{bg:"FF1E2535",fg:COL.white}};

    function cell(v,opts={}){
      const c={v};
      if(opts.t) c.t=opts.t;
      if(opts.f) c.f=opts.f;
      if(opts.s){
        c.s={};
        if(opts.s.fill) c.s.fill={patternType:"solid",fgColor:{argb:opts.s.fill}};
        if(opts.s.fgColor) c.s.font={...c.s.font,color:{argb:opts.s.fgColor},bold:!!opts.s.bold,sz:opts.s.sz||11};
        if(opts.s.bold) c.s.font={...c.s.font,bold:true};
        if(opts.s.sz) c.s.font={...c.s.font,sz:opts.s.sz};
        if(opts.s.align) c.s.alignment={horizontal:opts.s.align};
        if(opts.s.border) c.s.border={top:{style:"thin",color:{argb:COL.border}},bottom:{style:"thin",color:{argb:COL.border}},left:{style:"thin",color:{argb:COL.border}},right:{style:"thin",color:{argb:COL.border}}};
      }
      return c;
    }

    const wb=XLSX.utils.book_new();
    const dataRows=[];

    // ── Row 1: Title
    dataRows.push([
      cell(pa.number+" — "+(site?.name||""),{s:{fill:COL.dark,fgColor:COL.white,bold:true,sz:14}}),
      cell("",{s:{fill:COL.dark}}),cell("",{s:{fill:COL.dark}}),cell("",{s:{fill:COL.dark}}),cell("",{s:{fill:COL.dark}}),cell("",{s:{fill:COL.dark}}),cell("",{s:{fill:COL.dark}}),cell("",{s:{fill:COL.dark}}),cell("",{s:{fill:COL.dark}}),cell("",{s:{fill:COL.dark}}),
    ]);
    // ── Row 2: Sub-title
    dataRows.push([
      cell("Client: "+(client?.name||"—")+" · Date: "+pa.date+" · Status: "+pa.status,{s:{fill:COL.dark,fgColor:COL.prevFg}}),
      cell(""),cell(""),cell(""),cell(""),cell(""),cell(""),cell(""),cell(""),cell(""),
    ]);
    dataRows.push(Array(10).fill(cell("",{s:{fill:COL.dark}})));

    // ── Row 4: Headers
    const hdrs=["Type","Description","Unit","Contract Qty","Rate £","Contract £","Prev. Qty","Claimed to Date £","This Period £","% Complete"];
    dataRows.push(hdrs.map(h=>cell(h,{s:{fill:COL.headerBg,fgColor:COL.headerFg,bold:true,border:true,align:"center"}})));

    // ── Data rows with formulas
    const DATA_START=5; // row index where data begins (1-based for formula)
    items.forEach((it,idx)=>{
      const row=idx+DATA_START;
      const tc=typeColMap[it.type]||typeColMap["Other"];
      const cl=calcClaimed(it);
      const pr=calcPrev(it);
      const th=cl-pr;
      const pct=it.contractQty*it.contractRate>0?Math.round((cl/(it.contractQty*it.contractRate))*100):0;
      const contractVal=it.contractQty*it.contractRate;
      // Use actual formulas for Contract £, This Period, %
      dataRows.push([
        cell(it.type,                            {s:{fill:tc.bg,fgColor:tc.fg,bold:true,border:true}}),
        cell(it.description||"",                 {s:{fill:tc.bg,fgColor:COL.white,border:true}}),
        cell(it.unit,                            {s:{fill:tc.bg,fgColor:COL.prevFg,border:true,align:"center"}}),
        cell(it.contractQty,                     {t:"n",s:{fill:tc.bg,fgColor:COL.white,border:true,align:"right"}}),
        cell(it.contractRate,                    {t:"n",s:{fill:tc.bg,fgColor:COL.contractFg,border:true,align:"right"}}),
        {v:contractVal,t:"n",f:"D"+row+"*E"+row,s:{fill:{patternType:"solid",fgColor:{argb:tc.bg}},font:{color:{argb:COL.contractFg},bold:true},alignment:{horizontal:"right"},border:{top:{style:"thin",color:{argb:COL.border}},bottom:{style:"thin",color:{argb:COL.border}},left:{style:"thin",color:{argb:COL.border}},right:{style:"thin",color:{argb:COL.border}}}}},
        cell(it.previousQty||0,                  {t:"n",s:{fill:tc.bg,fgColor:COL.prevFg,border:true,align:"right"}}),
        cell(cl,                                 {t:"n",s:{fill:tc.bg,fgColor:COL.contractFg,bold:true,border:true,align:"right"}}),
        {v:th,t:"n",f:"H"+row+"-G"+row+"*E"+row,s:{fill:{patternType:"solid",fgColor:{argb:tc.bg}},font:{color:{argb:th>=0?"FF34D399":"FFF87171"},bold:true},alignment:{horizontal:"right"},border:{top:{style:"thin",color:{argb:COL.border}},bottom:{style:"thin",color:{argb:COL.border}},left:{style:"thin",color:{argb:COL.border}},right:{style:"thin",color:{argb:COL.border}}}}},
        {v:pct/100,t:"n",f:"IF(F"+row+">0,H"+row+"/F"+row+",0)",z:"0%",s:{fill:{patternType:"solid",fgColor:{argb:tc.bg}},font:{color:{argb:COL.pctFg},bold:true},alignment:{horizontal:"center"},border:{top:{style:"thin",color:{argb:COL.border}},bottom:{style:"thin",color:{argb:COL.border}},left:{style:"thin",color:{argb:COL.border}},right:{style:"thin",color:{argb:COL.border}}}}},
      ]);
    });

    // ── Totals row
    const lastDataRow=DATA_START+items.length-1;
    const totRow=lastDataRow+1;
    dataRows.push([
      cell("TOTALS",{s:{fill:COL.totalBg,fgColor:COL.totalFg,bold:true,sz:12,border:true}}),
      cell("",{s:{fill:COL.totalBg,border:true}}),
      cell("",{s:{fill:COL.totalBg,border:true}}),
      cell("",{s:{fill:COL.totalBg,border:true}}),
      cell("",{s:{fill:COL.totalBg,border:true}}),
      {v:totalContract,t:"n",f:"SUM(F"+DATA_START+":F"+lastDataRow+")",s:{fill:{patternType:"solid",fgColor:{argb:COL.totalBg}},font:{color:{argb:COL.contractFg},bold:true,sz:12},alignment:{horizontal:"right"},border:{top:{style:"medium",color:{argb:"FF10B981"}},bottom:{style:"medium",color:{argb:"FF10B981"}},left:{style:"thin"},right:{style:"thin"}}}},
      cell("",{s:{fill:COL.totalBg,border:true}}),
      {v:totalClaimed,t:"n",f:"SUM(H"+DATA_START+":H"+lastDataRow+")",s:{fill:{patternType:"solid",fgColor:{argb:COL.totalBg}},font:{color:{argb:COL.contractFg},bold:true,sz:12},alignment:{horizontal:"right"},border:{top:{style:"medium",color:{argb:"FF10B981"}},bottom:{style:"medium",color:{argb:"FF10B981"}},left:{style:"thin"},right:{style:"thin"}}}},
      {v:totalThis,t:"n",f:"SUM(I"+DATA_START+":I"+lastDataRow+")",s:{fill:{patternType:"solid",fgColor:{argb:COL.totalBg}},font:{color:{argb:totalThis>=0?"FF34D399":"FFF87171"},bold:true,sz:12},alignment:{horizontal:"right"},border:{top:{style:"medium",color:{argb:"FF10B981"}},bottom:{style:"medium",color:{argb:"FF10B981"}},left:{style:"thin"},right:{style:"thin"}}},},
      {v:pctComplete/100,t:"n",f:"IF(F"+totRow+">0,H"+totRow+"/F"+totRow+",0)",z:"0%",s:{fill:{patternType:"solid",fgColor:{argb:COL.totalBg}},font:{color:{argb:COL.pctFg},bold:true,sz:12},alignment:{horizontal:"center"},border:{top:{style:"medium",color:{argb:"FF10B981"}},bottom:{style:"medium",color:{argb:"FF10B981"}},left:{style:"thin"},right:{style:"thin"}}}},
    ]);

    // ── Build sheet
    const ws=XLSX.utils.aoa_to_sheet(dataRows);

    // Column widths
    ws["!cols"]=[{wch:18},{wch:38},{wch:8},{wch:14},{wch:12},{wch:14},{wch:12},{wch:18},{wch:16},{wch:12}];

    // Merge title cells A1:J1
    ws["!merges"]=[{s:{r:0,c:0},e:{r:0,c:9}},{s:{r:1,c:0},e:{r:1,c:9}},{s:{r:2,c:0},e:{r:2,c:9}}];

    XLSX.utils.book_append_sheet(wb,ws,"Payment Application");

    // ── Summary sheet
    const sumRows=[
      [cell("PAYMENT APPLICATION SUMMARY",{s:{fill:COL.dark,fgColor:COL.white,bold:true,sz:14}})],
      [cell(pa.number+" · "+(site?.name||"")+" · "+(client?.name||""),{s:{fill:COL.dark,fgColor:COL.prevFg}})],
      [cell("")],
      [cell("Contract Value",{s:{fill:"FF1E3A5F",fgColor:"FF60A5FA",bold:true}}),cell("£"+totalContract.toLocaleString(undefined,{maximumFractionDigits:0}),{t:"n",s:{fill:"FF1E3A5F",fgColor:"FF34D399",bold:true,sz:13}})],
      [cell("Claimed to Date",{s:{fill:"FF0D2218",fgColor:"FF34D399",bold:true}}),cell("£"+totalClaimed.toLocaleString(undefined,{maximumFractionDigits:0}),{t:"n",s:{fill:"FF0D2218",fgColor:"FF34D399",bold:true,sz:13}})],
      [cell("Previous",{s:{fill:"FF111827",fgColor:COL.prevFg,bold:true}}),cell("£"+totalPrev.toLocaleString(undefined,{maximumFractionDigits:0}),{t:"n",s:{fill:"FF111827",fgColor:COL.prevFg,bold:true,sz:13}})],
      [cell("This Period",{s:{fill:"FF1A1500",fgColor:COL.thisFg,bold:true}}),cell("£"+totalThis.toLocaleString(undefined,{maximumFractionDigits:0}),{t:"n",s:{fill:"FF1A1500",fgColor:COL.thisFg,bold:true,sz:13}})],
      [cell("% Complete",{s:{fill:"FF1A0D2E",fgColor:COL.pctFg,bold:true}}),cell(pctComplete+"%",{s:{fill:"FF1A0D2E",fgColor:COL.pctFg,bold:true,sz:13}})],
    ];
    const ws2=XLSX.utils.aoa_to_sheet(sumRows);
    ws2["!cols"]=[{wch:22},{wch:20}];
    XLSX.utils.book_append_sheet(wb,ws2,"Summary");

    const safeNum=(pa.number||"PA").split("/").join("-").split("\\").join("-");
    const safeSite=(site?.name||"site").split(" ").join("_");
    XLSX.writeFile(wb,safeNum+"_"+safeSite+".xlsx");
  }


  const typeColors={"Scope of Work":"#60a5fa","Variation":"#fbbf24","Dayworks":"#f97316","Preliminaries":"#a78bfa","Other":"#94a3b8"};

  return <div>
    <DPageHdr title={`📐 ${pa.number}`} sub={`${site?.name||"—"} · ${client?.name||"—"} · ${pa.date}`}
      back="Payment Applications" onBack={()=>setPage("payapps")}
      actions={<div style={{display:"flex",gap:7}}>
        <button onClick={savePA} style={{padding:"6px 13px",background:"#1e2535",border:"1px solid #2d3555",borderRadius:7,color:"#94a3b8",cursor:"pointer",fontSize:12}}>💾 Save</button>
        <button onClick={exportPDF} style={{padding:"6px 13px",background:"#1a1f2e",border:"1px solid #ef4444",borderRadius:7,color:"#f87171",cursor:"pointer",fontSize:12,fontWeight:700}}>📄 PDF</button>
        <button onClick={exportExcel} style={{padding:"6px 13px",background:"#0a1a0a",border:"1px solid #10b981",borderRadius:7,color:"#34d399",cursor:"pointer",fontSize:12,fontWeight:700}}>📊 Excel</button>
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
    <DPageHdr title="📋 Route / Forecast" sub={`WC: ${weekLabel} · ${displayed.length} operatives · Forecast only — confirmed by GPS sign-in`}
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

    <div style={{padding:"6px 20px",background:"#0f1421",borderBottom:"1px solid #1e2535",fontSize:11,color:"#64748b",display:"flex",alignItems:"center",gap:16,flexWrap:"wrap"}}>
      <span style={{color:"#fbbf24",fontWeight:700}}>🗺 This is a Route / Forecast only</span>
      <span style={{color:"#374151"}}>|</span>
      <span>Timesheet entries are created <strong style={{color:"#34d399"}}>only</strong> when a worker GPS signs in via the Worker Portal</span>
      <span style={{color:"#374151"}}>|</span>
      <span>✅ <span style={{color:"#34d399",fontWeight:600}}>Green cell</span> = GPS confirmed · 📋 Click cell to edit forecast</span>
      <span style={{color:"#374151"}}>|</span>
      <span>Route changes notify workers in-app and by email</span>
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
                      {(()=>{const confirmedLog=w.attendanceLogs?.find(l=>l.day===d&&l.weekLabel===weekLabel&&l.signIn&&l.signOut);return <InlineCell value={w.days?.[d]||""} workerId={w.id} day={d} allSiteNames={siteNames} allSites={allSites} onUpdate={updateCell||((id,day,val)=>{})} confirmed={!!confirmedLog}/>;})()}
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


function DashboardView({workers,allSites,clients,weekLabel,activeDays,siteHours,scopeData,invoices,saveWorker,delWorker,setAllSites,setClients,setModal,weeklyRecords,setWeeklyRecords,scheduleHistory,saveScheduleSnapshot,bankTransactions,setBankTransactions,timesheetRecords,setTimesheetRecords,payslipRecords,setPayslipRecords,payApplications,setPayApplications,generateTimesheets,generatePayslips,showWeekend,setShowWeekend,filter,setFilter,allSiteNames,updateCell,dashPage,setDashPage,dashDetailId,setDashDetailId}){
  const nav=(page,id)=>{setDashPage(page);if(id!==undefined)setDashDetailId(id);};
  const goBack=(page)=>setDashPage(page);

    // Shared props passed to all dashboard pages
  const SP={
    workers,allSites,clients,invoices,activeDays,siteHours,weekLabel,setModal,scopeData,
    filter,setFilter,allSiteNames,updateCell,delWorker,showWeekend,setShowWeekend,
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
      case "worker_detail": return <DWorkerDetail {...SP} workerId={dashDetailId} timesheetRecords={timesheetRecords} payslipRecords={payslipRecords} setTimesheetRecords={setTimesheetRecords} setPayslipRecords={setPayslipRecords} setPage={setDashPage}/>;
      // ── Labour Schedule (auto-snapshots each week)
      case "schedule":      return <DScheduleView {...SP}/>;
      case "site_by_site":  return <DSiteBySite {...SP}/>;
      case "payroll":       return <DPayroll {...SP}/>;
      case "payslips":      return <DPayslips {...SP} setPage={setDashPage}/>;
      case "timesheets":    return <DTimesheets {...SP} setPage={setDashPage}/>;
      case "weekly_records":return <DWeeklyRecords weeklyRecords={weeklyRecords} setWeeklyRecords={setWeeklyRecords} workers={workers} allSites={allSites} clients={clients} siteHours={siteHours} activeDays={activeDays} weekLabel={weekLabel} showWeekend={showWeekend} invoices={invoices} setPage={setDashPage} setDetailId={setDashDetailId}/>;
      case "weekly_record_detail": return <DWeeklyRecordDetail weeklyRecords={weeklyRecords} recordId={dashDetailId} setPage={setDashPage} allSites={allSites}/>;
      // ── Projects & Finance
      case "sites":         return <DSites {...SP} setPage={nav} setDetailId={setDashDetailId}/>;
      case "site_detail":   return <DSiteDetail {...SP} siteId={dashDetailId} setPage={setDashPage} setDetailId={setDashDetailId} invoices={invoices} payApplications={payApplications} weekLabel={weekLabel}/>;
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
      case "expenses":      return <DExpenses bankTransactions={bankTransactions} allSites={allSites} clients={clients} workers={workers} activeDays={activeDays} siteHours={siteHours} setPage={setDashPage}/>;
      case "pending_reg":   return <PendingWorkersView workers={workers} onApprove={()=>window.location.reload()}/>;
      case "app_sitemanager":  return <SiteManagerView/>;
      case "app_sitedocs":    return <SiteDocGeneratorView/>;
      case "app_assets":      return <AssetRegisterView/>;
      default:              return <DHome {...SP} weeklyRecords={weeklyRecords} setPage={setDashPage}/>;
    }
  };


  // DashboardView only renders the page content — sidebar is in App
  return <div style={{height:"100%",background:"#080d14",overflowY:"auto"}}>{renderPage()}</div>;
}


// ─── Site Manager App — dark theme matching Labour Schedule ───────────────────
const SiteManagerView = (()=>{

// ─── Helpers ──────────────────────────────────────────────────────────────────
const smFmt    = (n)  => `£${Math.round(n).toLocaleString("en-GB")}`;
const allocVal = (s)  => Math.round(s.total * (1 - (s.profitPct + s.overheadPct) / 100));
const earnedVal= (s)  => Math.round(allocVal(s) * s.completed / 100);
const pctTag   = (p)  => p < 70 ? "green" : p < 90 ? "amber" : "red";
const ragDot   = (p)  => p < 70 ? "🟢" : p < 90 ? "🟡" : "🔴";
const fileIcon = (name) => {
  const ext = (name||"").split(".").pop().toLowerCase();
  if(["jpg","jpeg","png","webp"].includes(ext)) return "🖼️";
  if(["mp4","mov","avi"].includes(ext))         return "🎬";
  if(["pdf"].includes(ext))                     return "📄";
  if(["eml","msg"].includes(ext))               return "📧";
  return "📎";
};

const BASE_MON = new Date(2026,5,9);
const DAY_SHORT = ["Mon","Tue","Wed","Thu","Fri","Sat"];
const MONTHS    = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
function weekMonday(off=0){ const d=new Date(BASE_MON); d.setDate(d.getDate()+off*7); return d; }
function weekDates(off=0){ const m=weekMonday(off); return Array.from({length:6},(_,i)=>{ const d=new Date(m); d.setDate(d.getDate()+i); return d; }); }
function weekKey(off=0){ const d=weekMonday(off); return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`; }
function smFmtDate(d){ return `${d.getDate()} ${MONTHS[d.getMonth()]}`; }
function dateToWK(ds){ const d=new Date(ds); const di=(d.getDay()+6)%7; const m=new Date(d); m.setDate(d.getDate()-di); return `${m.getFullYear()}-${String(m.getMonth()+1).padStart(2,"0")}-${String(m.getDate()).padStart(2,"0")}`; }
function dateToDI(ds){ return (new Date(ds).getDay()+6)%7; }

// ─── Colour palette — matches Labour Schedule dark theme ──────────────────────
const C = {
  bg:     "#0a0e1a",   // page background
  card:   "#111827",   // card background
  card2:  "#1a1f2e",   // elevated card
  border: "#1e2535",   // borders
  border2:"#2d3555",   // lighter borders
  accent: "#3b82f6",   // blue
  green:  "#34d399",   // green
  yellow: "#fbbf24",   // amber
  red:    "#f87171",   // red
  purple: "#a78bfa",   // purple
  orange: "#f97316",   // orange
  text:   "#f1f5f9",   // primary text
  sub:    "#94a3b8",   // secondary text
  muted:  "#64748b",   // muted text
  dim:    "#374151",   // very dim
};

const RAG_C = {
  green: {bg:"#0d221844",bd:"#34d39944",col:"#34d399"},
  amber: {bg:"#1a150044",bd:"#fbbf2444",col:"#fbbf24"},
  red:   {bg:"#2d151544",bd:"#f8717144",col:"#f87171"},
};
const TRADE_COL={
  "Management":"#3b82f6","Structural":"#f97316","General":"#34d399",
  "Plant":"#fbbf24","Mechanical":"#a78bfa","Electrical":"#818cf8",
  "Masonry":"#94a3b8","Carpentry":"#2dd4bf"
};

// ─── Data ─────────────────────────────────────────────────────────────────────
const MANAGER={name:"James Mitchell",initials:"JM"};
const SM_CATS=["Materials","Labour","Plant & Equipment","Subcontract","Preliminaries","Other"];
const TODAY="2026-06-13";

const SITES=[
  {id:1,name:"Highfield Business Park",    client:"Nexus Developments Ltd", address:"Highfield Rd, Aldershot GU11",status:"active",  appStatus:"draft",    start:"Jan 2025",end:"Aug 2025"},
  {id:2,name:"Riverside Apartments — Block C",client:"Riverside Living PLC",address:"Riverside Dr, Farnham GU9",  status:"active",  appStatus:"submitted",start:"Mar 2025",end:"Nov 2025"},
  {id:3,name:"Sainsbury's Refurb — Fleet",client:"Sainsbury's Supermarkets",address:"Fleet Rd, Fleet GU51",       status:"complete",appStatus:"finalised",start:"Nov 2024",end:"Apr 2025"},
];

const INIT_SCOPES={
  1:[
    {id:1,name:"Groundworks & Drainage",  total:95000, profitPct:12,overheadPct:8, completed:100},
    {id:2,name:"Structural Steel Frame",  total:142000,profitPct:15,overheadPct:7, completed:85},
    {id:3,name:"Cladding & Curtain Wall", total:118000,profitPct:10,overheadPct:9, completed:40},
    {id:4,name:"Electrical First Fix",    total:67000, profitPct:12,overheadPct:8, completed:60},
    {id:5,name:"Mechanical Services",     total:63000, profitPct:11,overheadPct:9, completed:20},
  ],
  2:[
    {id:6,name:"Demolition & Strip Out",  total:45000, profitPct:10,overheadPct:8, completed:100},
    {id:7,name:"RC Frame Works",          total:120000,profitPct:14,overheadPct:8, completed:55},
    {id:8,name:"Brickwork & Masonry",     total:88000, profitPct:12,overheadPct:7, completed:20},
    {id:9,name:"Roof & Waterproofing",    total:67000, profitPct:11,overheadPct:9, completed:0},
  ],
  3:[
    {id:10,name:"Strip Out & Demolition", total:32000, profitPct:8, overheadPct:7, completed:100},
    {id:11,name:"M&E Services",           total:78000, profitPct:10,overheadPct:8, completed:100},
    {id:12,name:"Finishes & Fittings",    total:68000, profitPct:9, overheadPct:8, completed:100},
  ],
};

const INIT_EXP=[
  {id:1,siteId:1,scopeId:2,desc:"Steel delivery — Phase 1",   amount:14800,date:"12 May 2025",category:"Materials",        files:["invoice_steel.pdf"], status:"approved"},
  {id:2,siteId:1,scopeId:4,desc:"Cable drum & conduit supply", amount:3400, date:"20 May 2025",category:"Materials",        files:["receipt_elec.jpg"],  status:"pending"},
  {id:3,siteId:2,scopeId:7,desc:"Pump hire — concrete pour",   amount:1650, date:"30 Apr 2025",category:"Plant & Equipment",files:["hire_invoice.pdf"],  status:"approved"},
];

const INIT_VAR=[
  {id:1,siteId:1,ref:"VAR-001",title:"Additional waterproof membrane to basement",desc:"Client requested upgraded spec following high water table survey.",amount:8500,status:"pending", files:["survey_report.pdf","site_photo_001.jpg"]},
  {id:2,siteId:1,ref:"VAR-002",title:"Revised escape route — Level 2",desc:"Building control requirement following L2 design change.",amount:3200,status:"approved",files:["bcf_email.eml","revised_plan.pdf"]},
];

const LABOUR_DATA={
  1:{workers:[{id:"w1",name:"Dave Hartley",role:"Site Foreman",trade:"Management",initials:"DH",dailyRate:380},{id:"w2",name:"Mike Patel",role:"Steel Fixer",trade:"Structural",initials:"MP",dailyRate:280},{id:"w3",name:"John Walsh",role:"Steel Fixer",trade:"Structural",initials:"JW",dailyRate:270},{id:"w4",name:"Steve Clarke",role:"Labourer",trade:"General",initials:"SC",dailyRate:200},{id:"w5",name:"Ryan Brooks",role:"Plant Operator",trade:"Plant",initials:"RB",dailyRate:320},{id:"w6",name:"Lucy Chen",role:"M&E Engineer",trade:"Mechanical",initials:"LC",dailyRate:350}],
    schedule:{"2026-05-19":{w1:[0,1,2,3,4],w2:[0,1,2,3,4],w3:[0,1,2,3,4],w4:[0,1,2,3,4],w5:[0,1,2],w6:[]},"2026-06-09":{w1:[0,1,2,3,4],w2:[0,1,2,3,4],w3:[0,1,2,3],w4:[0,1,2,3,4,5],w5:[0,2,4],w6:[2,3,4]},"2026-06-16":{w1:[0,1,2,3,4],w2:[0,1,3,4],w3:[0,1,2,3,4],w4:[1,2,3,4],w5:[0,1,2],w6:[0,1,2,3,4]}}},
  2:{workers:[{id:"w7",name:"Tom Bradley",role:"Site Foreman",trade:"Management",initials:"TB",dailyRate:360},{id:"w8",name:"Sean Murphy",role:"Bricklayer",trade:"Masonry",initials:"SM",dailyRate:260},{id:"w9",name:"Ali Hassan",role:"Carpenter",trade:"Carpentry",initials:"AH",dailyRate:250},{id:"w10",name:"Chris Ford",role:"Labourer",trade:"General",initials:"CF",dailyRate:190}],
    schedule:{"2026-06-09":{w7:[0,1,2,3,4],w8:[0,1,2,3,4],w9:[1,2,3],w10:[0,1,3,4]},"2026-06-16":{w7:[0,1,2,3,4],w8:[0,2,3,4],w9:[0,1,2,3,4],w10:[0,1,2]}}},
  3:{workers:[{id:"w11",name:"Paul Green",role:"Site Foreman",trade:"Management",initials:"PG",dailyRate:360},{id:"w12",name:"Neil Carter",role:"M&E Engineer",trade:"Mechanical",initials:"NC",dailyRate:340},{id:"w13",name:"Sam Osei",role:"Electrician",trade:"Electrical",initials:"SO",dailyRate:290}],
    schedule:{"2026-04-07":{w11:[0,1,2,3,4],w12:[0,1,2,3,4],w13:[0,1,2,3,4]}}},
};

const INIT_SIGN_INS={
  1:{"2026-06-09":{w1:2,w2:2,w3:2,w4:3,w5:3,w6:null},"2026-06-10":{w1:2,w2:2,w3:2,w4:3,w5:null,w6:5},"2026-06-11":{w1:2,w2:2,w3:2,w4:3,w5:3,w6:5},"2026-06-12":{w1:2,w2:2,w3:2,w4:3,w5:null,w6:5},"2026-06-13":{w1:null,w2:null,w3:null,w4:3,w5:3,w6:null}},
  2:{"2026-06-09":{w7:7,w8:8,w9:null,w10:7},"2026-06-10":{w7:7,w8:8,w9:7,w10:7},"2026-06-11":{w7:7,w8:8,w9:7,w10:7},"2026-06-12":{w7:7,w8:8,w9:null,w10:7}},
  3:{},
};

// ─── Micro components ──────────────────────────────────────────────────────────
function SmBadge({s}){
  const M={
    approved:[C.green+"22",C.green,"Approved"],
    pending: [C.yellow+"22",C.yellow,"Pending"],
    rejected:[C.red+"22",C.red,"Rejected"],
    draft:   [C.border,C.muted,"Draft"],
    submitted:[C.accent+"22",C.accent,"Submitted"],
    finalised:[C.green+"22",C.green,"Finalised"],
    active:  [C.accent+"22",C.accent,"Active"],
    complete:[C.green+"22",C.green,"Complete"],
  };
  const [bg,col,lbl]=M[s]||[C.border,C.muted,s];
  return <span style={{background:bg,color:col,padding:"2px 10px",borderRadius:999,fontSize:11,fontWeight:700,border:`1px solid ${col}44`}}>{lbl}</span>;
}

function SmBar({pct,thin,color}){
  const c=color||(pct===100?C.green:pct>80?C.red:pct>60?C.yellow:C.accent);
  return <div style={{background:C.border,borderRadius:999,height:thin?3:6,overflow:"hidden"}}>
    <div style={{width:`${Math.min(pct,100)}%`,background:c,borderRadius:999,height:"100%",transition:"width .4s ease"}}/>
  </div>;
}

function SmCrumb({crumbs}){
  return <div style={{display:"flex",alignItems:"center",gap:5,marginBottom:14,flexWrap:"wrap"}}>
    {crumbs.map((c,i)=>(
      <span key={i} style={{display:"flex",alignItems:"center",gap:5}}>
        {i>0&&<span style={{color:C.muted,fontSize:13}}>›</span>}
        <span onClick={c.fn} style={{fontSize:12,color:c.fn?C.accent:C.sub,cursor:c.fn?"pointer":"default",fontWeight:c.fn?500:700}}>{c.label}</span>
      </span>
    ))}
  </div>;
}

function SmPill({name,onX}){
  return <span style={{background:C.card2,border:`1px solid ${C.border2}`,borderRadius:6,padding:"4px 10px",fontSize:11,color:C.sub,display:"inline-flex",alignItems:"center",gap:4}}>
    {fileIcon(name)}{name}
    {onX&&<span style={{cursor:"pointer",color:C.muted,marginLeft:2}} onClick={onX}>✕</span>}
  </span>;
}

function SmDrop({fRef,onFiles,hint}){
  return <>
    <div onClick={()=>fRef.current?.click()} style={{border:`2px dashed ${C.border2}`,borderRadius:10,padding:"22px 16px",textAlign:"center",cursor:"pointer",background:C.card}}>
      <div style={{fontSize:28,marginBottom:5}}>📎</div>
      <div style={{color:C.sub,fontSize:13,fontWeight:500}}>Click to attach files</div>
      <div style={{color:C.muted,fontSize:11,marginTop:2}}>{hint}</div>
    </div>
    <input ref={fRef} type="file" multiple style={{display:"none"}} onChange={e=>{onFiles(Array.from(e.target.files).map(f=>f.name));e.target.value="";}}/>
  </>;
}

// ─── Financial Widget ──────────────────────────────────────────────────────────
function FinancialWidget({fin,scopes,open,onToggle}){
  const tag=pctTag(fin.pctUsed);
  const rag=RAG_C[tag];
  return <div style={{background:rag.bg,border:`1px solid ${rag.bd}`,borderRadius:12,padding:"14px 16px",marginBottom:16}}>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:10,flexWrap:"wrap",gap:8}}>
      <div style={{display:"flex",gap:20,flexWrap:"wrap"}}>
        {[["Allocated Budget",smFmt(fin.allocated),C.text],["Total Spent",smFmt(fin.totalSpent),C.sub],["Remaining",fin.remaining>=0?smFmt(fin.remaining):`${smFmt(Math.abs(fin.remaining))} OVER`,fin.remaining>=0?C.green:C.red]].map(([l,v,c])=>(
          <div key={l}>
            <div style={{fontSize:9,color:C.muted,fontWeight:700,textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:2}}>{l}</div>
            <div style={{fontSize:18,fontWeight:800,color:c}}>{v}</div>
          </div>
        ))}
      </div>
      <div style={{display:"flex",alignItems:"center",gap:8}}>
        <span style={{fontSize:22}}>{ragDot(fin.pctUsed)}</span>
        <button onClick={onToggle} style={{background:C.card,border:`1px solid ${rag.bd}`,borderRadius:7,padding:"5px 11px",cursor:"pointer",fontSize:11,color:C.sub,fontWeight:700}}>
          {open?"▲ Hide":"▼ Scopes"}
        </button>
      </div>
    </div>
    <SmBar pct={fin.pctUsed} thin color={rag.col}/>
    <div style={{display:"flex",justifyContent:"space-between",fontSize:10,color:C.muted,marginTop:4}}>
      <span>Labour {smFmt(fin.labour)} · Expenses {smFmt(fin.expenses)}</span>
      <span style={{fontWeight:700,color:rag.col}}>{fin.pctUsed.toFixed(1)}% of budget used</span>
    </div>
    {open&&<div style={{marginTop:14,borderTop:`1px solid ${rag.bd}`,paddingTop:12,display:"flex",flexDirection:"column",gap:10}}>
      {scopes.map(s=>{
        const av=allocVal(s),lc=fin.labourByScope[s.id]||0,ec=fin.expByScope[s.id]||0,sp=lc+ec,rem=av-sp,p=av>0?sp/av*100:0,st=pctTag(p);
        return <div key={s.id}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"baseline",marginBottom:4}}>
            <span style={{fontSize:12,fontWeight:700,color:C.text}}>{s.name}</span>
            <div style={{display:"flex",alignItems:"center",gap:8}}>
              <span style={{fontSize:10,color:C.muted}}>{smFmt(sp)} of {smFmt(av)}</span>
              <span style={{fontSize:13}}>{ragDot(p)}</span>
            </div>
          </div>
          <SmBar pct={p} thin color={RAG_C[st].col}/>
          <div style={{fontSize:10,color:C.muted,marginTop:2,display:"flex",gap:10}}>
            <span>Labour {smFmt(lc)}</span><span>Expenses {smFmt(ec)}</span>
            <span style={{color:rem>=0?C.green:C.red,fontWeight:700}}>{rem>=0?`${smFmt(rem)} left`:`${smFmt(Math.abs(rem))} OVER`}</span>
            <span style={{marginLeft:"auto",color:C.muted}}>{s.completed}% done</span>
          </div>
        </div>;
      })}
    </div>}
  </div>;
}

// ─── Shell ─────────────────────────────────────────────────────────────────────
function Shell({view,siteId,go,toast,fin,scopes,widgetOpen,onWidgetToggle,children}){
  const hasSite=siteId!==null;
  const navItems=[
    {id:"dashboard",icon:"🏗️",label:"My Sites"},
    ...(hasSite?[
      {id:"site",       icon:"📍",label:"Site"},
      {id:"scopes",     icon:"📁",label:"Scopes"},
      {id:"labour",     icon:"👷",label:"Labour"},
      {id:"signin",     icon:"✅",label:"Sign-In"},
      {id:"expenses",   icon:"🧾",label:"Expenses"},
      {id:"variations", icon:"📐",label:"Variations"},
      {id:"application",icon:"📊",label:"Application"},
    ]:[]),
  ];
  const isActive=(id)=>{
    if(id==="expenses")  return view==="expenses"||view==="expense-new";
    if(id==="variations")return view==="variations"||view==="variation-new";
    return view===id;
  };
  return <div style={{minHeight:"100vh",background:C.bg,fontFamily:"'Inter',system-ui,sans-serif",display:"flex",flexDirection:"column",color:C.text}}>
    {/* Header */}
    <header style={{background:"#0f172a",borderBottom:`1px solid ${C.border}`,height:56,display:"flex",alignItems:"center",justifyContent:"space-between",padding:"0 20px",position:"sticky",top:0,zIndex:100,boxShadow:"0 2px 12px rgba(0,0,0,.4)"}}>
      <div style={{display:"flex",alignItems:"center",gap:10}}>
        <div style={{width:34,height:34,background:"linear-gradient(135deg,#1e3a5f,#3b82f6)",borderRadius:9,display:"flex",alignItems:"center",justifyContent:"center",fontSize:18}}>🏗️</div>
        <div>
          <div style={{color:C.text,fontSize:14,fontWeight:800,letterSpacing:"-0.01em"}}>Site Manager</div>
          <div style={{color:C.muted,fontSize:10}}>Bright Metalwork</div>
        </div>
      </div>
      <div style={{display:"flex",alignItems:"center",gap:8}}>
        <div style={{textAlign:"right"}}>
          <div style={{color:C.text,fontSize:12,fontWeight:600}}>{MANAGER.name}</div>
          <div style={{color:C.muted,fontSize:10}}>Site Manager</div>
        </div>
        <div style={{width:34,height:34,background:"linear-gradient(135deg,#1e3a5f,#3b82f6)",borderRadius:"50%",display:"flex",alignItems:"center",justifyContent:"center",color:"#fff",fontWeight:800,fontSize:12}}>{MANAGER.initials}</div>
      </div>
    </header>

    {/* Nav tabs */}
    <nav style={{background:"#111827",borderBottom:`1px solid ${C.border}`,display:"flex",overflowX:"auto"}}>
      {navItems.map(n=>(
        <button key={n.id} onClick={()=>go(n.id)}
          style={{padding:"11px 13px",background:"none",border:"none",
            borderBottom:`2px solid ${isActive(n.id)?C.accent:"transparent"}`,
            cursor:"pointer",fontSize:11,fontWeight:700,
            color:isActive(n.id)?C.accent:C.muted,
            whiteSpace:"nowrap",transition:"color .15s"}}>
          <span style={{marginRight:3}}>{n.icon}</span>{n.label}
        </button>
      ))}
    </nav>

    <main style={{flex:1,maxWidth:860,width:"100%",margin:"0 auto",padding:"16px 16px 48px",boxSizing:"border-box"}}>
      {hasSite&&fin&&<FinancialWidget fin={fin} scopes={scopes} open={widgetOpen} onToggle={onWidgetToggle}/>}
      {children}
    </main>

    {toast&&<div style={{position:"fixed",bottom:24,left:"50%",transform:"translateX(-50%)",background:C.green,color:"#0a1a10",padding:"11px 24px",borderRadius:10,fontSize:13,fontWeight:700,boxShadow:"0 4px 20px rgba(52,211,153,.4)",zIndex:9999,whiteSpace:"nowrap"}}>{toast}</div>}
  </div>;
}

// ─── Shared card / section styles ─────────────────────────────────────────────
const smCard={background:C.card2,border:`1px solid ${C.border}`,borderRadius:12,padding:"14px 16px",marginBottom:14};
const smBtn=(pri)=>({padding:"8px 16px",borderRadius:8,border:pri?"none":`1px solid ${C.border2}`,background:pri?"linear-gradient(135deg,#1e3a5f,#3b82f6)":C.card,color:pri?C.text:C.sub,cursor:"pointer",fontSize:12,fontWeight:700});
const smInp={width:"100%",background:C.card,border:`1px solid ${C.border2}`,borderRadius:8,padding:"9px 12px",color:C.text,fontSize:13,outline:"none",boxSizing:"border-box"};
const smSel={...smInp,cursor:"pointer"};
const smTH={padding:"8px 12px",background:"#0d1117",color:C.muted,fontSize:10,fontWeight:700,textTransform:"uppercase",letterSpacing:"0.05em",textAlign:"left",borderBottom:`1px solid ${C.border}`};
const smTD={padding:"10px 12px",borderBottom:`1px solid ${C.border}`,fontSize:13,color:C.text};

function pop(setToast,msg){setToast(msg);setTimeout(()=>setToast(null),2400);}
function SmSec({title,action,children}){
  return <div style={smCard}>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
      <div style={{fontSize:12,color:C.muted,fontWeight:700,textTransform:"uppercase",letterSpacing:"0.06em"}}>{title}</div>
      {action}
    </div>
    {children}
  </div>;
}

// ─── Main App ─────────────────────────────────────────────────────────────────
return function SiteManagerView(){
  const [view,setView]       =useState("dashboard");
  const [siteId,setSiteId]   =useState(null);
  const [scopeId,setScopeId] =useState(null);
  const [scopes,setScopes]   =useState(INIT_SCOPES);
  const [expenses,setExp]    =useState(INIT_EXP);
  const [variations,setVar]  =useState(INIT_VAR);
  const [signIns,setSignIns] =useState(INIT_SIGN_INS);
  const [scopePct,setScopePct]=useState(null);
  const [weekOffset,setWeekOff]=useState(0);
  const [widgetOpen,setWidgetOpen]=useState(false);
  const [signInDate,setSignInDate]=useState(TODAY);
  const [draftSI,setDraftSI] =useState(null);
  const [expForm,setEF]=useState({desc:"",amount:"",category:"Materials",scopeId:"",date:"",files:[]});
  const [varForm,setVF]=useState({title:"",desc:"",amount:"",files:[]});
  const [toast,setToast]=useState(null);
  const expRef=useRef(); const varRef=useRef();

  const go=(v,opts={})=>{
    setView(v);
    if(opts.siteId!==undefined) setSiteId(opts.siteId);
    if(opts.scopeId!==undefined) setScopeId(opts.scopeId);
    if(opts.pct!==undefined) setScopePct(opts.pct);
  };
  const site=SITES.find(s=>s.id===siteId);
  const siteSc=siteId?scopes[siteId]||[]:[];
  const siteLD=siteId?LABOUR_DATA[siteId]:null;
  const siteSI=siteId?signIns[siteId]||{}:{};

  const fin=siteId?(()=>{
    const ss=scopes[siteId]||[];
    const allocated=ss.reduce((a,s)=>a+allocVal(s),0);
    const labourByScope={};
    const expByScope={};
    Object.entries(siteSI).forEach(([date,day])=>{
      Object.entries(day).forEach(([wid,sid])=>{
        if(!sid) return;
        const w=siteLD?.workers.find(x=>x.id===wid);
        if(!w) return;
        labourByScope[sid]=(labourByScope[sid]||0)+w.dailyRate;
      });
    });
    (expenses.filter(e=>e.siteId===siteId)).forEach(e=>{
      expByScope[e.scopeId]=(expByScope[e.scopeId]||0)+e.amount;
    });
    const labour=Object.values(labourByScope).reduce((a,v)=>a+v,0);
    const exps=Object.values(expByScope).reduce((a,v)=>a+v,0);
    const totalSpent=labour+exps;
    const remaining=allocated-totalSpent;
    const pctUsed=allocated>0?totalSpent/allocated*100:0;
    return{allocated,labour,expenses:exps,totalSpent,remaining,pctUsed,labourByScope,expByScope};
  })():null;

  // ── DASHBOARD ─────────────────────────────────────────────────────────────
  if(view==="dashboard") return <Shell view={view} siteId={null} go={(v)=>go(v)} toast={toast} fin={null} scopes={[]} widgetOpen={false} onWidgetToggle={()=>{}}>
    <div style={{marginBottom:20}}>
      <div style={{fontSize:22,fontWeight:800,color:C.text,marginBottom:2}}>My Sites</div>
      <div style={{fontSize:12,color:C.muted}}>Welcome back, {MANAGER.name}</div>
    </div>
    {SITES.map(s=>{
      const ss=scopes[s.id]||[];
      const totalVal=ss.reduce((a,x)=>a+x.total,0);
      const avgDone=ss.length?ss.reduce((a,x)=>a+x.completed,0)/ss.length:0;
      return <div key={s.id} onClick={()=>go("site",{siteId:s.id})}
        style={{...smCard,cursor:"pointer",borderLeft:`3px solid ${s.status==="complete"?C.green:C.accent}`,
          transition:"border-color .15s"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:10}}>
          <div>
            <div style={{fontSize:15,fontWeight:800,color:C.text,marginBottom:2}}>{s.name}</div>
            <div style={{fontSize:11,color:C.muted}}>{s.client} · {s.address}</div>
          </div>
          <SmBadge s={s.status}/>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10,marginBottom:10}}>
          {[["Contract Value",smFmt(totalVal),C.accent],["Avg Progress",avgDone.toFixed(0)+"%",avgDone===100?C.green:C.yellow],["Scopes",ss.length+" items",C.sub]].map(([l,v,col])=>(
            <div key={l} style={{background:C.bg,borderRadius:8,padding:"8px 10px"}}>
              <div style={{fontSize:9,color:C.muted,fontWeight:700,textTransform:"uppercase"}}>{l}</div>
              <div style={{fontSize:14,fontWeight:800,color:col,marginTop:2}}>{v}</div>
            </div>
          ))}
        </div>
        <SmBar pct={avgDone}/>
        <div style={{display:"flex",justifyContent:"space-between",marginTop:4,fontSize:10,color:C.muted}}>
          <span>{s.start} — {s.end}</span>
          <span style={{color:C.accent,fontWeight:600}}>Open site →</span>
        </div>
      </div>;
    })}
  </Shell>;

  // ── SITE OVERVIEW ─────────────────────────────────────────────────────────
  if(view==="site"&&site) return <Shell {...{view,siteId,go:(v)=>go(v,{siteId}),toast,fin,scopes:siteSc,widgetOpen,onWidgetToggle:()=>setWidgetOpen(o=>!o)}}>
    <SmCrumb crumbs={[{label:"My Sites",fn:()=>go("dashboard")},{label:site.name}]}/>
    <div style={{...smCard}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:12}}>
        <div><div style={{fontSize:18,fontWeight:800,color:C.text}}>{site.name}</div><div style={{fontSize:11,color:C.muted,marginTop:2}}>{site.client} · {site.address}</div></div>
        <SmBadge s={site.status}/>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
        {[["Start",site.start],["End",site.end],["PA Status",site.appStatus],["Scopes",siteSc.length+" items"]].map(([l,v])=>(
          <div key={l} style={{background:C.bg,borderRadius:8,padding:"8px 10px"}}>
            <div style={{fontSize:9,color:C.muted,fontWeight:700,textTransform:"uppercase",marginBottom:2}}>{l}</div>
            <div style={{fontSize:13,fontWeight:600,color:C.text}}>{v}</div>
          </div>
        ))}
      </div>
    </div>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
      {[["📁 Scopes","scopes",C.accent],["👷 Labour","labour",C.purple],["✅ Sign-In","signin",C.green],["🧾 Expenses","expenses",C.yellow],["📐 Variations","variations",C.orange],["📊 Application","application",C.red]].map(([l,v,col])=>(
        <button key={v} onClick={()=>go(v,{siteId})}
          style={{padding:"14px 16px",background:C.card2,border:`1px solid ${col}33`,borderRadius:10,cursor:"pointer",textAlign:"left",transition:"border-color .15s"}}>
          <div style={{fontSize:16,marginBottom:4}}>{l.split(" ")[0]}</div>
          <div style={{fontSize:13,fontWeight:700,color:col}}>{l.split(" ").slice(1).join(" ")}</div>
        </button>
      ))}
    </div>
  </Shell>;

  // ── SCOPES ─────────────────────────────────────────────────────────────────
  if(view==="scopes"&&site) return <Shell {...{view,siteId,go:(v)=>go(v,{siteId}),toast,fin,scopes:siteSc,widgetOpen,onWidgetToggle:()=>setWidgetOpen(o=>!o)}}>
    <SmCrumb crumbs={[{label:"My Sites",fn:()=>go("dashboard")},{label:site.name,fn:()=>go("site",{siteId})},{label:"Scopes"}]}/>
    {siteSc.map(s=>{
      const av=allocVal(s),lab=fin?.labourByScope[s.id]||0,exp=fin?.expByScope[s.id]||0,sp=lab+exp,rem=av-sp,p=av>0?sp/av*100:0,st=pctTag(p);
      return <div key={s.id} style={{...smCard,borderLeft:`3px solid ${RAG_C[st].col}`}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:8}}>
          <div style={{fontSize:14,fontWeight:700,color:C.text}}>{s.name}</div>
          <div style={{display:"flex",alignItems:"center",gap:8}}>
            <span style={{fontSize:18}}>{ragDot(p)}</span>
            <SmBadge s={s.completed===100?"complete":"active"}/>
          </div>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:8,marginBottom:10}}>
          {[["Contract",smFmt(s.total),C.accent],["Budget",smFmt(av),C.text],["Spent",smFmt(sp),sp>av?C.red:C.sub],["Remaining",rem>=0?smFmt(rem):`${smFmt(Math.abs(rem))} OVER`,rem>=0?C.green:C.red]].map(([l,v,c])=>(
            <div key={l} style={{background:C.bg,borderRadius:7,padding:"7px 9px"}}>
              <div style={{fontSize:9,color:C.muted,fontWeight:700,textTransform:"uppercase"}}>{l}</div>
              <div style={{fontSize:13,fontWeight:700,color:c,marginTop:1}}>{v}</div>
            </div>
          ))}
        </div>
        <SmBar pct={p} color={RAG_C[st].col}/>
        <div style={{display:"flex",justifyContent:"space-between",marginTop:5,fontSize:10,color:C.muted}}>
          <span>Work done: <b style={{color:C.text}}>{s.completed}%</b></span>
          <span>{p.toFixed(0)}% of budget used</span>
        </div>
        <div style={{marginTop:10}}>
          <div style={{fontSize:9,color:C.muted,fontWeight:700,textTransform:"uppercase",marginBottom:5}}>Completion Progress</div>
          <SmBar pct={s.completed} color={s.completed===100?C.green:C.accent}/>
        </div>
      </div>;
    })}
  </Shell>;

  // ── LABOUR ─────────────────────────────────────────────────────────────────
  if(view==="labour"&&site&&siteLD) return <Shell {...{view,siteId,go:(v)=>go(v,{siteId}),toast,fin,scopes:siteSc,widgetOpen,onWidgetToggle:()=>setWidgetOpen(o=>!o)}}>
    <SmCrumb crumbs={[{label:"My Sites",fn:()=>go("dashboard")},{label:site.name,fn:()=>go("site",{siteId})},{label:"Labour"}]}/>
    <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:14}}>
      <button onClick={()=>setWeekOff(o=>o-1)} style={{...smBtn(false),padding:"6px 12px"}}>←</button>
      <div style={{flex:1,textAlign:"center"}}>
        <div style={{fontWeight:700,color:C.text,fontSize:14}}>WC {smFmtDate(weekMonday(weekOffset))}</div>
        <div style={{fontSize:10,color:C.muted}}>Week {weekOffset===0?"(Current)":weekOffset>0?`+${weekOffset}`:`${weekOffset}`}</div>
      </div>
      <button onClick={()=>setWeekOff(o=>o+1)} style={{...smBtn(false),padding:"6px 12px"}}>→</button>
    </div>
    <div style={{overflowX:"auto"}}>
      <table style={{width:"100%",borderCollapse:"collapse",minWidth:520}}>
        <thead><tr>
          <th style={smTH}>Worker</th>
          {weekDates(weekOffset).map(d=><th key={d} style={{...smTH,textAlign:"center",minWidth:48}}>{smFmtDate(d)}</th>)}
          <th style={{...smTH,textAlign:"right"}}>Days</th>
          <th style={{...smTH,textAlign:"right"}}>Cost</th>
        </tr></thead>
        <tbody>
          {siteLD.workers.map((w,i)=>{
            const wk=weekKey(weekOffset);
            const sched=siteLD.schedule[wk]?.[w.id]||[];
            const days=sched.length;
            return <tr key={w.id} style={{background:i%2===0?C.card:"#0f1421"}}>
              <td style={smTD}>
                <div style={{display:"flex",alignItems:"center",gap:8}}>
                  <div style={{width:30,height:30,borderRadius:7,background:(TRADE_COL[w.trade]||C.accent)+"22",border:`1px solid ${(TRADE_COL[w.trade]||C.accent)}44`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,fontWeight:700,color:TRADE_COL[w.trade]||C.accent}}>{w.initials}</div>
                  <div><div style={{fontSize:12,fontWeight:600,color:C.text}}>{w.name}</div><div style={{fontSize:10,color:C.muted}}>{w.role}</div></div>
                </div>
              </td>
              {weekDates(weekOffset).map((d,di)=>{
                const on=sched.includes(di);
                return <td key={d} style={{...smTD,textAlign:"center"}}>
                  <div style={{width:24,height:24,borderRadius:5,background:on?(TRADE_COL[w.trade]||C.accent)+"33":"transparent",border:`1px solid ${on?(TRADE_COL[w.trade]||C.accent)+"66":C.border}`,display:"flex",alignItems:"center",justifyContent:"center",margin:"0 auto",fontSize:12}}>
                    {on?"✓":""}
                  </div>
                </td>;
              })}
              <td style={{...smTD,textAlign:"right",color:C.accent,fontWeight:700}}>{days}</td>
              <td style={{...smTD,textAlign:"right",color:C.green,fontWeight:700}}>{smFmt(days*w.dailyRate)}</td>
            </tr>;
          })}
        </tbody>
      </table>
    </div>
  </Shell>;

  // ── SIGN-IN ────────────────────────────────────────────────────────────────
  if(view==="signin"&&site&&siteLD){
    const dayKey=signInDate;
    const dayData=siteSI[dayKey]||{};
    const draft=draftSI||Object.fromEntries(siteLD.workers.map(w=>[w.id,dayData[w.id]!==undefined?dayData[w.id]:null]));
    const save=()=>{setSignIns(si=>({...si,[siteId]:{...si[siteId],[dayKey]:draft}}));setDraftSI(null);pop(setToast,"Sign-in saved ✓");};
    return <Shell {...{view,siteId,go:(v)=>go(v,{siteId}),toast,fin,scopes:siteSc,widgetOpen,onWidgetToggle:()=>setWidgetOpen(o=>!o)}}>
      <SmCrumb crumbs={[{label:"My Sites",fn:()=>go("dashboard")},{label:site.name,fn:()=>go("site",{siteId})},{label:"Daily Sign-In"}]}/>
      <div style={{...smCard,marginBottom:12}}>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <div style={{flex:1}}>
            <div style={{fontSize:9,color:C.muted,fontWeight:700,textTransform:"uppercase",marginBottom:4}}>Date</div>
            <input type="date" value={signInDate} onChange={e=>{setSignInDate(e.target.value);setDraftSI(null);}} style={{...smInp}}/>
          </div>
          <button onClick={save} style={{...smBtn(true),marginTop:16,whiteSpace:"nowrap"}}>Save Sign-In</button>
        </div>
      </div>
      <div style={{border:`1px solid ${C.border}`,borderRadius:10,overflow:"hidden"}}>
        <table style={{width:"100%",borderCollapse:"collapse"}}>
          <thead><tr>
            <th style={smTH}>Worker</th>
            <th style={smTH}>Trade</th>
            <th style={{...smTH,textAlign:"center"}}>Present</th>
            <th style={smTH}>Scope</th>
          </tr></thead>
          <tbody>
            {siteLD.workers.map((w,i)=>{
              const present=draft[w.id]!==null&&draft[w.id]!==undefined&&draft[w.id]!==false;
              const sc=present?draft[w.id]:null;
              return <tr key={w.id} style={{background:i%2===0?C.card:"#0f1421"}}>
                <td style={smTD}>
                  <div style={{display:"flex",alignItems:"center",gap:8}}>
                    <div style={{width:28,height:28,borderRadius:7,background:(TRADE_COL[w.trade]||C.accent)+"22",border:`1px solid ${(TRADE_COL[w.trade]||C.accent)}44`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,fontWeight:700,color:TRADE_COL[w.trade]||C.accent}}>{w.initials}</div>
                    <div style={{fontSize:12,fontWeight:600,color:C.text}}>{w.name}<div style={{fontSize:10,color:C.muted}}>{w.role}</div></div>
                  </div>
                </td>
                <td style={smTD}><span style={{color:TRADE_COL[w.trade]||C.accent,fontSize:11,fontWeight:600}}>{w.trade}</span></td>
                <td style={{...smTD,textAlign:"center"}}>
                  <div onClick={()=>setDraftSI(d=>({...d,[w.id]:present?null:(siteSc[0]?.id||null)}))}
                    style={{width:28,height:28,borderRadius:7,background:present?C.green+"22":C.card,border:`2px solid ${present?C.green:C.border}`,display:"flex",alignItems:"center",justifyContent:"center",cursor:"pointer",margin:"0 auto",fontSize:14,fontWeight:700,color:present?C.green:C.dim}}>
                    {present?"✓":""}
                  </div>
                </td>
                <td style={smTD}>
                  {present
                    ?<select value={sc||""} onChange={e=>setDraftSI(d=>({...d,[w.id]:Number(e.target.value)||null}))} style={{...smSel,width:"100%"}}>
                        <option value="">— Scope —</option>
                        {siteSc.map(s=><option key={s.id} value={s.id}>{s.name}</option>)}
                      </select>
                    :<span style={{fontSize:11,color:C.dim}}>—</span>}
                </td>
              </tr>;
            })}
          </tbody>
        </table>
      </div>
    </Shell>;
  }

  // ── EXPENSES ───────────────────────────────────────────────────────────────
  if((view==="expenses"||view==="expense-new")&&site){
    const siteExp=expenses.filter(e=>e.siteId===siteId);
    if(view==="expense-new") return <Shell {...{view,siteId,go:(v)=>go(v,{siteId}),toast,fin,scopes:siteSc,widgetOpen,onWidgetToggle:()=>setWidgetOpen(o=>!o)}}>
      <SmCrumb crumbs={[{label:"My Sites",fn:()=>go("dashboard")},{label:site.name,fn:()=>go("site",{siteId})},{label:"Expenses",fn:()=>go("expenses",{siteId})},{label:"New Expense"}]}/>
      <div style={smCard}>
        {[["Description","text","desc","Steel delivery…"],["Amount £","number","amount","0.00"],["Date","date","date",""]].map(([l,t,k,ph])=>(
          <div key={k} style={{marginBottom:12}}>
            <div style={{fontSize:9,color:C.muted,fontWeight:700,textTransform:"uppercase",marginBottom:4}}>{l}</div>
            <input type={t} value={expForm[k]} onChange={e=>setEF(f=>({...f,[k]:e.target.value}))} placeholder={ph} style={smInp}/>
          </div>
        ))}
        <div style={{marginBottom:12}}>
          <div style={{fontSize:9,color:C.muted,fontWeight:700,textTransform:"uppercase",marginBottom:4}}>Category</div>
          <select value={expForm.category} onChange={e=>setEF(f=>({...f,category:e.target.value}))} style={smSel}>
            {SM_CATS.map(c=><option key={c} value={c}>{c}</option>)}
          </select>
        </div>
        <div style={{marginBottom:12}}>
          <div style={{fontSize:9,color:C.muted,fontWeight:700,textTransform:"uppercase",marginBottom:4}}>Scope</div>
          <select value={expForm.scopeId} onChange={e=>setEF(f=>({...f,scopeId:e.target.value}))} style={smSel}>
            <option value="">— Select scope —</option>
            {siteSc.map(s=><option key={s.id} value={s.id}>{s.name}</option>)}
          </select>
        </div>
        <div style={{marginBottom:14}}>
          <div style={{fontSize:9,color:C.muted,fontWeight:700,textTransform:"uppercase",marginBottom:6}}>Attachments</div>
          <SmDrop fRef={expRef} onFiles={fs=>setEF(f=>({...f,files:[...f.files,...fs]}))} hint="Invoices, receipts, photos"/>
          {expForm.files.length>0&&<div style={{display:"flex",gap:6,flexWrap:"wrap",marginTop:8}}>{expForm.files.map((n,i)=><SmPill key={i} name={n} onX={()=>setEF(f=>({...f,files:f.files.filter((_,j)=>j!==i)}))}/>)}</div>}
        </div>
        <div style={{display:"flex",gap:10}}>
          <button onClick={()=>go("expenses",{siteId})} style={{...smBtn(false),flex:1}}>Cancel</button>
          <button onClick={()=>{if(!expForm.desc||!expForm.amount)return;setExp(ex=>[...ex,{id:Date.now(),siteId,scopeId:Number(expForm.scopeId)||null,desc:expForm.desc,amount:Number(expForm.amount),date:expForm.date||TODAY,category:expForm.category,files:expForm.files,status:"pending"}]);setEF({desc:"",amount:"",category:"Materials",scopeId:"",date:"",files:[]});go("expenses",{siteId});pop(setToast,"Expense submitted ✓");}} style={{...smBtn(true),flex:2}}>Submit Expense</button>
        </div>
      </div>
    </Shell>;

    return <Shell {...{view,siteId,go:(v)=>go(v,{siteId}),toast,fin,scopes:siteSc,widgetOpen,onWidgetToggle:()=>setWidgetOpen(o=>!o)}}>
      <SmCrumb crumbs={[{label:"My Sites",fn:()=>go("dashboard")},{label:site.name,fn:()=>go("site",{siteId})},{label:"Expenses"}]}/>
      <div style={{display:"flex",justifyContent:"flex-end",marginBottom:12}}>
        <button onClick={()=>go("expense-new",{siteId})} style={smBtn(true)}>+ New Expense</button>
      </div>
      {siteExp.length===0&&<div style={{...smCard,textAlign:"center",color:C.muted,padding:32}}>No expenses yet.</div>}
      {siteExp.map(e=>(
        <div key={e.id} style={{...smCard,borderLeft:`3px solid ${e.status==="approved"?C.green:e.status==="pending"?C.yellow:C.red}`}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:6}}>
            <div style={{fontSize:13,fontWeight:700,color:C.text,flex:1,marginRight:10}}>{e.desc}</div>
            <SmBadge s={e.status}/>
          </div>
          <div style={{display:"flex",gap:14,flexWrap:"wrap",fontSize:11,color:C.muted}}>
            <span style={{color:C.green,fontWeight:700,fontSize:14}}>{smFmt(e.amount)}</span>
            <span>{e.category}</span><span>{e.date}</span>
            {e.scopeId&&<span style={{color:C.accent}}>{siteSc.find(s=>s.id===e.scopeId)?.name||"—"}</span>}
          </div>
          {e.files.length>0&&<div style={{display:"flex",gap:6,flexWrap:"wrap",marginTop:8}}>{e.files.map((n,i)=><SmPill key={i} name={n}/>)}</div>}
        </div>
      ))}
    </Shell>;
  }

  // ── VARIATIONS ─────────────────────────────────────────────────────────────
  if((view==="variations"||view==="variation-new")&&site){
    const siteVar=variations.filter(v=>v.siteId===siteId);
    if(view==="variation-new") return <Shell {...{view,siteId,go:(v)=>go(v,{siteId}),toast,fin,scopes:siteSc,widgetOpen,onWidgetToggle:()=>setWidgetOpen(o=>!o)}}>
      <SmCrumb crumbs={[{label:"My Sites",fn:()=>go("dashboard")},{label:site.name,fn:()=>go("site",{siteId})},{label:"Variations",fn:()=>go("variations",{siteId})},{label:"New Variation"}]}/>
      <div style={smCard}>
        {[["Title","text","title","e.g. Additional steelwork to level 3"],["Description","text","desc","Reason / instruction reference…"],["Amount £","number","amount","0.00"]].map(([l,t,k,ph])=>(
          <div key={k} style={{marginBottom:12}}>
            <div style={{fontSize:9,color:C.muted,fontWeight:700,textTransform:"uppercase",marginBottom:4}}>{l}</div>
            <input type={t} value={varForm[k]} onChange={e=>setVF(f=>({...f,[k]:e.target.value}))} placeholder={ph} style={smInp}/>
          </div>
        ))}
        <div style={{marginBottom:14}}>
          <div style={{fontSize:9,color:C.muted,fontWeight:700,textTransform:"uppercase",marginBottom:6}}>Supporting Files</div>
          <SmDrop fRef={varRef} onFiles={fs=>setVF(f=>({...f,files:[...f.files,...fs]}))} hint="Emails, drawings, photos"/>
          {varForm.files.length>0&&<div style={{display:"flex",gap:6,flexWrap:"wrap",marginTop:8}}>{varForm.files.map((n,i)=><SmPill key={i} name={n} onX={()=>setVF(f=>({...f,files:f.files.filter((_,j)=>j!==i)}))}/>)}</div>}
        </div>
        <div style={{display:"flex",gap:10}}>
          <button onClick={()=>go("variations",{siteId})} style={{...smBtn(false),flex:1}}>Cancel</button>
          <button onClick={()=>{if(!varForm.title||!varForm.amount)return;const n=siteVar.length+1;setVar(vs=>[...vs,{id:Date.now(),siteId,ref:`VAR-${String(n).padStart(3,"0")}`,title:varForm.title,desc:varForm.desc,amount:Number(varForm.amount),status:"pending",files:varForm.files}]);setVF({title:"",desc:"",amount:"",files:[]});go("variations",{siteId});pop(setToast,"Variation submitted ✓");}} style={{...smBtn(true),flex:2}}>Submit Variation</button>
        </div>
      </div>
    </Shell>;

    return <Shell {...{view,siteId,go:(v)=>go(v,{siteId}),toast,fin,scopes:siteSc,widgetOpen,onWidgetToggle:()=>setWidgetOpen(o=>!o)}}>
      <SmCrumb crumbs={[{label:"My Sites",fn:()=>go("dashboard")},{label:site.name,fn:()=>go("site",{siteId})},{label:"Variations"}]}/>
      <div style={{display:"flex",justifyContent:"flex-end",marginBottom:12}}>
        <button onClick={()=>go("variation-new",{siteId})} style={smBtn(true)}>+ New Variation</button>
      </div>
      {siteVar.map(v=>(
        <div key={v.id} style={{...smCard,borderLeft:`3px solid ${v.status==="approved"?C.green:v.status==="pending"?C.yellow:C.red}`}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:6}}>
            <div>
              <div style={{fontSize:10,color:C.muted,fontWeight:700,fontFamily:"monospace",marginBottom:2}}>{v.ref}</div>
              <div style={{fontSize:13,fontWeight:700,color:C.text}}>{v.title}</div>
            </div>
            <SmBadge s={v.status}/>
          </div>
          {v.desc&&<div style={{fontSize:12,color:C.sub,marginBottom:8,lineHeight:1.5}}>{v.desc}</div>}
          <div style={{fontSize:14,fontWeight:800,color:C.orange,marginBottom:8}}>{smFmt(v.amount)}</div>
          {v.files.length>0&&<div style={{display:"flex",gap:6,flexWrap:"wrap"}}>{v.files.map((n,i)=><SmPill key={i} name={n}/>)}</div>}
        </div>
      ))}
    </Shell>;
  }

  // ── PAYMENT APPLICATION ────────────────────────────────────────────────────
  if(view==="application"&&site){
    const appStat=site.appStatus;
    const approvedVars=variations.filter(v=>v.siteId===siteId&&v.status==="approved");
    const varTotal=approvedVars.reduce((a,v)=>a+v.amount,0);
    const scopeTotal=siteSc.reduce((a,s)=>a+earnedVal(s),0);
    return <Shell {...{view,siteId,go:(v)=>go(v,{siteId}),toast,fin,scopes:siteSc,widgetOpen,onWidgetToggle:()=>setWidgetOpen(o=>!o)}}>
      <SmCrumb crumbs={[{label:"My Sites",fn:()=>go("dashboard")},{label:site.name,fn:()=>go("site",{siteId})},{label:"Payment Application"}]}/>
      <div style={{...smCard,display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
        <div style={{fontSize:13,fontWeight:700,color:C.sub}}>Application Status</div>
        <SmBadge s={appStat}/>
      </div>
      <div style={{border:`1px solid ${C.border}`,borderRadius:10,overflow:"hidden",marginBottom:14}}>
        <div style={{padding:"9px 14px",background:"#0d1117",borderBottom:`1px solid ${C.border}`}}>
          <div style={{fontSize:10,color:C.muted,fontWeight:700,textTransform:"uppercase"}}>Scope Breakdown — Earned Value</div>
        </div>
        {siteSc.map((s,i)=>(
          <div key={s.id} style={{display:"flex",justifyContent:"space-between",padding:"10px 14px",borderBottom:`1px solid ${C.border}`,background:i%2===0?C.card:"#0f1421",alignItems:"center"}}>
            <div style={{fontSize:13,fontWeight:600,color:C.text,flex:1}}>{s.name}</div>
            <div style={{textAlign:"right"}}>
              <div style={{fontSize:13,fontWeight:700,color:C.accent}}>{smFmt(earnedVal(s))}</div>
              <div style={{fontSize:10,color:C.muted}}>{s.completed}% of {smFmt(allocVal(s))}</div>
            </div>
          </div>
        ))}
        <div style={{padding:"10px 14px",background:"#0d1117",display:"flex",justifyContent:"space-between"}}>
          <span style={{fontWeight:700,fontSize:12,color:C.sub}}>Scopes Subtotal</span>
          <span style={{fontWeight:800,fontSize:14,color:C.green}}>{smFmt(scopeTotal)}</span>
        </div>
      </div>
      {approvedVars.length>0&&<div style={{border:`1px solid ${C.border}`,borderRadius:10,overflow:"hidden",marginBottom:14}}>
        <div style={{padding:"9px 14px",background:"#0d1117",borderBottom:`1px solid ${C.border}`}}>
          <div style={{fontSize:10,color:C.muted,fontWeight:700,textTransform:"uppercase"}}>Approved Variations</div>
        </div>
        {approvedVars.map(v=><div key={v.id} style={{padding:"10px 14px",borderBottom:`1px solid ${C.border}`,display:"flex",justifyContent:"space-between",background:C.card}}>
          <div style={{fontSize:13,fontWeight:600,color:C.text}}>{v.ref} — {v.title}</div>
          <div style={{fontWeight:800,color:C.orange}}>{smFmt(v.amount)}</div>
        </div>)}
        <div style={{padding:"10px 14px",background:"#0d1117",display:"flex",justifyContent:"space-between"}}>
          <span style={{fontWeight:700,fontSize:12,color:C.sub}}>Variations Subtotal</span>
          <span style={{fontWeight:800,fontSize:14,color:C.orange}}>{smFmt(varTotal)}</span>
        </div>
      </div>}
      <div style={{background:"linear-gradient(135deg,#1e3a5f,#1a1f2e)",border:`1px solid ${C.accent}44`,borderRadius:12,padding:"16px 20px",marginBottom:16,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <div>
          <div style={{color:C.muted,fontSize:10,fontWeight:700,textTransform:"uppercase",letterSpacing:"0.06em"}}>Manager's Total (excl. margin)</div>
          <div style={{color:C.sub,fontSize:11,marginTop:1}}>Director will add margin values before issuing to client</div>
        </div>
        <div style={{color:C.green,fontSize:28,fontWeight:900}}>{smFmt(scopeTotal+varTotal)}</div>
      </div>
      {appStat==="draft"&&<div style={{...smCard,border:`1px solid ${C.yellow}44`,background:C.yellow+"11"}}>
        <div style={{fontWeight:800,color:C.yellow,fontSize:13,marginBottom:5}}>⚠️ Submit for Director Review</div>
        <div style={{color:C.sub,fontSize:12,marginBottom:12,lineHeight:1.5}}>Once submitted the director and finance team will review, add full margin values, and issue to the client.</div>
        <button style={smBtn(true)} onClick={()=>pop(setToast,"Submitted for director review ✓")}>Submit for Review</button>
      </div>}
      {appStat==="submitted"&&<div style={{...smCard,border:`1px solid ${C.accent}44`,background:C.accent+"11"}}>
        <div style={{fontWeight:800,color:C.accent,fontSize:13,marginBottom:5}}>⏳ Awaiting Finalisation</div>
        <div style={{color:C.sub,fontSize:12}}>With the director team. You'll be notified once finalised and issued to the client.</div>
      </div>}
      {appStat==="finalised"&&<div style={{...smCard,border:`1px solid ${C.green}44`,background:C.green+"11"}}>
        <div style={{fontWeight:800,color:C.green,fontSize:13,marginBottom:5}}>✅ Finalised & Issued</div>
        <div style={{color:C.sub,fontSize:12}}>Finalised by the director team and issued to the client.</div>
      </div>}
    </Shell>;
  }

  return null;
};

})();
// ─── Site Doc Generator (self-contained IIFE — zero scope conflicts) ─────────
const SiteDocGeneratorView = (()=>{

// ═══════════════════════════════════════════════════════
// DATA
// ═══════════════════════════════════════════════════════

const DEMO_SITES = [
  {
    id: "s1", name: "Paddington Station", address: "Praed St, London W2 1HQ",
    scope: "External façade works – structural repointing and brick replacement",
    supervisor: { name: "Marcus Webb", phone: "07700 900123" },
    workers: [
      { id: "w1", name: "Tom Bradley", role: "Scaffolder" },
      { id: "w2", name: "Sarah Chen", role: "Bricklayer" },
      { id: "w3", name: "Dev Patel", role: "Labourer" },
      { id: "w4", name: "James O'Brien", role: "Foreman" },
    ],
  },
  {
    id: "s2", name: "Canary Wharf – Tower B", address: "Bank St, London E14 5JP",
    scope: "Internal fit-out – MEP first fix and structural steelwork",
    supervisor: { name: "Lisa Park", phone: "07700 900456" },
    workers: [
      { id: "w5", name: "Ahmed Khalil", role: "Steel Erector" },
      { id: "w6", name: "Emma Thornton", role: "Electrician" },
      { id: "w7", name: "Carlos Mendez", role: "Pipefitter" },
      { id: "w8", name: "Priya Singh", role: "Site Manager" },
      { id: "w9", name: "Noel Byrne", role: "Labourer" },
    ],
  },
];

const CATEGORIES = [
  { id: "safety",     label: "Safety",               color: "#F5A623" },
  { id: "inspection", label: "Pre-Use Inspections",  color: "#3B82F6" },
  { id: "lifting",    label: "Lifting",              color: "#7C3AED" },
  { id: "qa",         label: "Quality Assurance",    color: "#0D9488" },
  { id: "legal",      label: "Legal",                color: "#DC2626" },
];

const DOC_TYPES = [
  { id: "daily_safe_start", label: "Daily Safe Start",       cat: "safety",     desc: "Pre-work briefing, hazard ID & team sign-off" },
  { id: "toolbox_talk",     label: "Toolbox Talk",           cat: "safety",     desc: "Focused safety topic with attendance record" },
  { id: "puwer",            label: "PUWER Assessment",       cat: "safety",     desc: "Provision & Use of Work Equipment Regulations" },
  { id: "loler",            label: "LOLER Checklist",        cat: "safety",     desc: "Lifting Operations & Lifting Equipment Regs" },
  { id: "coshh",            label: "COSHH Assessment",       cat: "safety",     desc: "Control of Substances Hazardous to Health" },
  { id: "task_briefing",    label: "Task Briefing",          cat: "safety",     desc: "Method statement briefing & team sign-off" },
  { id: "harness",          label: "Harness Inspection",     cat: "inspection", desc: "Full-body harness pre-use checklist" },
  { id: "scissor_lift",     label: "Scissor Lift",           cat: "inspection", desc: "Aerial work platform pre-use check" },
  { id: "cherry_picker",    label: "Cherry Picker",          cat: "inspection", desc: "MEWP / boom lift pre-use inspection" },
  { id: "spider_crane",     label: "Spider Crane",           cat: "inspection", desc: "Mini crawler crane pre-use check" },
  { id: "pat_testing",      label: "PAT Testing Record",     cat: "lifting",    desc: "Portable appliance test record" },
  { id: "thorough_exam",    label: "Thorough Examination",   cat: "lifting",    desc: "LOLER thorough examination certificate" },
  { id: "qa_handover",      label: "QA Handover Form",       cat: "qa",         desc: "Quality assurance sign-off for completed works" },
  { id: "ncr",              label: "NCR Form",               cat: "qa",         desc: "Non-Conformance Report" },
  { id: "early_delay",      label: "Early Delay Notice",     cat: "legal",      desc: "Early warning of potential delay event" },
  { id: "delay_notice",     label: "Delay Notice",           cat: "legal",      desc: "Formal notice of delay and compensation claim" },
];

const CL = {
  harness: {
    s: null, sc: null,
    items: [
      "Webbing – cuts, fraying, abrasion, heat or chemical damage",
      "Stitching – all intact, no broken threads",
      "Back D-ring – no deformation, cracks or corrosion",
      "Front / chest D-ring – secure, no deformation",
      "Buckles & adjusters – undamaged, function correctly",
      "Leg straps – no twisting, correct routing",
      "Shoulder straps – no deformation, correct routing",
      "Connectors / karabiners – gate closes and locks",
      "Lanyard / energy absorber – no deployment triggered",
      "ID label – serial number and manufacture date legible",
    ],
  },
  scissor_lift: {
    s: ["External", "Ground controls", "Platform", "Platform controls", "Safety devices"],
    sc: [3, 3, 4, 4, 2],
    items: [
      "Body panels – no significant damage or missing sections",
      "Tyres / wheels – correct pressure, no damage",
      "Hydraulic lines – no visible leaks or damage",
      "Emergency lowering function – operates correctly",
      "Ground control panel – labelling legible, switches functional",
      "Battery / fuel level – adequate for task",
      "Platform floor – anti-slip surface intact",
      "Guardrails – all in place, no deformation",
      "Mid-rails and toe-boards – present and secure",
      "Platform gate – closes and latches correctly",
      "Control panel – all labels legible",
      "Drive / steer controls – smooth, self-centring",
      "Lift / lower controls – smooth, correct speed",
      "Emergency stop – tested and functional",
      "Pothole protection / outriggers – deploy and lock",
      "Tilt sensor / alarm – tested and operational",
    ],
  },
  cherry_picker: {
    s: ["External", "Boom", "Outriggers", "Platform / basket", "Ground controls", "Platform controls", "Safety devices"],
    sc: [3, 3, 3, 3, 3, 3, 2],
    items: [
      "Body – no significant damage, leaks or missing guards",
      "Engine / battery – fluid levels correct, no leaks",
      "Drive system – tyres / tracks in good condition",
      "Boom sections – no cracks, buckles or deformation",
      "Boom seals – no hydraulic leaks at cylinders",
      "Rotation system – slewing smooth, locks correctly",
      "Outrigger legs – extend and retract fully",
      "Outrigger pads – present, in good condition",
      "Machine levelling – achieves level within tolerance",
      "Basket floor and sides – no cracks or missing sections",
      "Guardrails – all present, height correct",
      "Anchor point – rated, no deformation",
      "Ground control panel – all functions operate correctly",
      "Emergency descent – tested and operational",
      "Key switch / access control – functioning",
      "Control panel – legible, all switches functional",
      "Joystick / controls – smooth, self-centring",
      "Emergency stop on platform – tested",
      "Overload indicator – present and functional",
      "Envelope / height limiter – set and tested",
    ],
  },
  spider_crane: {
    s: ["Outrigger legs", "Boom & structure", "Wire rope & hoist", "Hook & lifting gear", "Controls & drive", "Safety & hydraulics"],
    sc: [3, 3, 3, 4, 3, 4],
    items: [
      "All four legs extend fully, locking pins secure",
      "Outrigger pads in place, adequate bearing capacity",
      "Machine levels within tolerance before lift",
      "Main boom sections – no cracks, welds sound",
      "Hinge pins fully inserted with retaining clips",
      "Hydraulic cylinders – no leaks at seals",
      "Wire rope – no broken strands, kinks or corrosion",
      "Rope anchoring on drum correct, min 3 turns remaining",
      "Sheaves / rollers – no cracking, rotate freely",
      "Hook – no deformation, latch closes and holds",
      "Hook swivel rotates freely",
      "Sling / chain certificate – in date",
      "Shackles / slings in use – rated, no damage",
      "Radio / pendant controls – all functions tested",
      "E-stop on remote and machine – tested",
      "Slew / boom controls – smooth, no drift",
      "SLI (safe load indicator) – calibrated, alarm sounds",
      "Hydraulic fluid level – within acceptable range",
      "All hydraulic connections tight, no weeping",
      "Anti-two-block device – present and functional",
    ],
  },
};

const QA_ITEMS = [
  "Workmanship meets specification",
  "Materials correct type and grade as specified",
  "Dimensions within tolerance per drawings",
  "Surface finish meets required standard",
  "Fixings / connections correctly installed",
  "Waterproofing / sealing correctly applied",
  "Adequate protection applied post-installation",
  "As-built record / marked-up drawing available",
];

// ═══════════════════════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════════════════════

const genId = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
const today = () => new Date().toLocaleDateString("en-GB");
const CAT_COLOR = { safety: "#F5A623", inspection: "#3B82F6", lifting: "#7C3AED", qa: "#0D9488", legal: "#DC2626" };
const RESP_ROLES = ["foreman", "supervisor", "site manager", "manager"];

function itemSection(docId, idx) {
  const cl = CL[docId];
  if (!cl?.s) return null;
  let c = 0;
  for (let si = 0; si < cl.s.length; si++) {
    c += cl.sc[si];
    if (idx < c) return cl.s[si];
  }
  return null;
}

function initChecks(docId) {
  const cl = CL[docId];
  if (!cl) return {};
  return cl.items.reduce((a, _, i) => { a[i] = { r: "", n: "" }; return a; }, {});
}

function initFd(dt, site) {
  return {
    date: today(), site: site?.name || "", address: site?.address || "",
    scope: site?.scope || "", supervisor: site?.supervisor?.name || "",
    supPhone: site?.supervisor?.phone || "", preparedBy: site?.supervisor?.name || "",
    topic: "", hazards: "", controls: "", ppe: "", emergencyProc: "",
    methodRef: "", taskDesc: "", resources: "", taskRisks: "", taskControls: "",
    equipment: "", equipmentId: "", make: "", model: "",
    riskLevel: "", frequency: "", maintenance: "", training: "",
    substance: "", supplier: "", hazardType: "", exposureLimits: "",
    healthEffects: "", exposureRoutes: "", cohhControls: "", cohhPpe: "", emergencyActions: "",
    liftingEquip: "", swl: "", liftPlan: "", examDate: "", nextExam: "", examRef: "",
    applianceDesc: "", assetNo: "", location: "", testedBy: "",
    visualPass: "", earthPass: "", insulationPass: "", polarityPass: "", patResult: "", nextTestDate: "",
    examEquipment: "", examId: "", examSWL: "", examiner: "", examBody: "",
    conditionRating: "", defectsDesc: "", safeForUse: "", nextExamDate: "",
    contractRef: "", workPackage: "", deficiencies: "",
    qaItems: QA_ITEMS.map(item => ({ item, result: "", notes: "" })),
    ncrRef: "", ncrDesc: "", rootCause: "", correctiveAction: "", targetDate: "", closedBy: "",
    contractNo: "", employer: "", contractor: "", programmeRef: "",
    eventDate: today(), eventDesc: "", delayWeeks: "", costImpact: "", delayImpact: "", replyBy: "",
    checks: initChecks(dt.id), result: "", notes: "",
    attendees: (site?.workers || []).map(w => ({
      ...w, present: true,
      responsible: RESP_ROLES.some(r => w.role?.toLowerCase().includes(r)),
    })),
  };
}

async function loadDocs() {
  try {
    const raw = localStorage.getItem("bm_site_documents");
    if (raw) return JSON.parse(raw);
  } catch (_) {}
  return [];
}

async function saveDoc(doc, existing) {
  try {
    const updated = [...(existing || []).filter(d => d.id !== doc.id), doc];
    localStorage.setItem("bm_site_documents", JSON.stringify(updated));
    return updated;
  } catch (_) { return existing || []; }
}

// ═══════════════════════════════════════════════════════
// TOKENS
// ═══════════════════════════════════════════════════════

const BG = "#F5F4F0", DARK = "#111318", HDR = "#0F1117", AMB = "#F5A623",
  GY1 = "#F3F4F6", GY2 = "#E5E7EB", GY3 = "#D1D5DB",
  GY5 = "#6B7280", GY6 = "#4B5563", GY7 = "#374151", GY8 = "#1F2937",
  GRN = "#15803D", GBG = "#DCFCE7", GRL = "#86EFAC",
  RED = "#DC2626", RBG = "#FEE2E2",
  BLU = "#1D4ED8", BBG = "#DBEAFE";

const inp = {
  width: "100%", padding: "7px 10px", border: `1px solid ${GY3}`,
  borderRadius: 6, fontSize: 13, color: GY8, background: "#fff",
  boxSizing: "border-box", fontFamily: "inherit",
};
const ta = { ...inp, resize: "vertical", minHeight: 70, lineHeight: 1.5 };
const lbl = {
  fontSize: 10, fontWeight: 600, color: GY5,
  textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 3, display: "block",
};

// ═══════════════════════════════════════════════════════
// SHARED COMPONENTS
// ═══════════════════════════════════════════════════════

function Field({ label, value, onChange, multi, span }) {
  return (
    <div style={{ gridColumn: span ? "1 / -1" : undefined }}>
      {label && <label style={lbl}>{label}</label>}
      {multi
        ? <textarea style={ta} value={value || ""} onChange={e => onChange?.(e.target.value)} />
        : <input style={inp} value={value || ""} onChange={e => onChange?.(e.target.value)} />}
    </div>
  );
}

function Sec({ title, children }) {
  return (
    <div style={{ background: "#fff", border: `1px solid ${GY2}`, borderRadius: 10, padding: "14px 17px", marginBottom: 12 }}>
      <p style={{ margin: "0 0 11px", fontSize: 10, fontWeight: 700, color: GY5, textTransform: "uppercase", letterSpacing: "0.08em" }}>{title}</p>
      {children}
    </div>
  );
}

function Grid({ cols = 2, children }) {
  return (
    <div style={{ display: "grid", gridTemplateColumns: `repeat(${cols}, 1fr)`, gap: 11 }}>
      {children}
    </div>
  );
}

function DocCard({ dt, onClick, count }) {
  const cc = CAT_COLOR[dt.cat];
  const [hov, setHov] = useState(false);
  return (
    <button onClick={onClick}
      onMouseEnter={() => setHov(true)} onMouseLeave={() => setHov(false)}
      style={{
        background: hov ? GY1 : "#fff", border: `1px solid ${GY2}`,
        borderLeft: `3px solid ${cc}`, borderRadius: 0,
        padding: "12px 14px", cursor: "pointer", textAlign: "left", transition: "background 0.1s",
      }}>
      <p style={{ margin: 0, fontSize: 13, fontWeight: 600, color: GY8 }}>{dt.label}</p>
      <p style={{ margin: "3px 0 0", fontSize: 11, color: GY5 }}>{dt.desc}</p>
      {count > 0 && <p style={{ margin: "5px 0 0", fontSize: 10, fontWeight: 700, color: cc }}>{count} saved</p>}
    </button>
  );
}

function IBlock({ label, value }) {
  return (
    <div style={{ flex: 1, minWidth: 140 }}>
      <p style={{ margin: 0, fontSize: 9, fontWeight: 700, color: "#9CA3AF", textTransform: "uppercase", letterSpacing: "0.08em" }}>{label}</p>
      <p style={{ margin: "2px 0 0", fontSize: 12, color: GY7 }}>{value || "—"}</p>
    </div>
  );
}

// ═══════════════════════════════════════════════════════
// SIDEBAR
// ═══════════════════════════════════════════════════════

function Sidebar({ sites, sel, setSel, sync, syncSt, onRep, docsCount }) {
  return (
    <aside style={{ width: 210, background: DARK, display: "flex", flexDirection: "column", flexShrink: 0, minHeight: "100vh" }}>
      <div style={{ padding: "16px 13px 11px", borderBottom: "1px solid rgba(255,255,255,0.07)" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 7, marginBottom: 9 }}>
          <div style={{ width: 26, height: 26, background: AMB, borderRadius: 5, display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 800, fontSize: 11, color: DARK }}>SD</div>
          <span style={{ color: "#fff", fontWeight: 700, fontSize: 13 }}>Site Docs</span>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 5 }}>
          <div style={{ width: 6, height: 6, borderRadius: "50%", background: syncSt === "live" ? "#22C55E" : AMB }} />
          <span style={{ fontSize: 10, color: syncSt === "live" ? "#86EFAC" : "#FCD34D" }}>
            {syncSt === "live" ? "Live – Labour Schedule" : "Demo mode"}
          </span>
        </div>
      </div>
      <div style={{ padding: "10px 7px", flex: 1, overflowY: "auto" }}>
        <p style={{ margin: "0 0 5px", color: "rgba(255,255,255,0.28)", fontSize: 9, textTransform: "uppercase", letterSpacing: "0.1em", padding: "0 6px" }}>Sites</p>
        {sites.map(s => (
          <button key={s.id} onClick={() => setSel(s)} style={{
            width: "100%", textAlign: "left", padding: "7px 7px", borderRadius: 6, border: "none", cursor: "pointer",
            background: sel?.id === s.id ? "rgba(245,166,35,0.12)" : "transparent", marginBottom: 2,
          }}>
            <p style={{ margin: 0, fontSize: 12, fontWeight: 500, color: sel?.id === s.id ? AMB : "rgba(255,255,255,0.85)" }}>{s.name}</p>
            <p style={{ margin: 0, fontSize: 10, color: "rgba(255,255,255,0.3)" }}>{s.workers?.length || 0} workers</p>
          </button>
        ))}
      </div>
      <div style={{ padding: "10px 11px 13px", borderTop: "1px solid rgba(255,255,255,0.07)" }}>
        <button onClick={onRep} style={{ width: "100%", padding: "7px 10px", background: "rgba(255,255,255,0.05)", border: "1px solid rgba(255,255,255,0.11)", borderRadius: 6, color: "rgba(255,255,255,0.7)", fontSize: 12, cursor: "pointer", marginBottom: 6, textAlign: "left" }}>
          📋 Reports ({docsCount})
        </button>
        <button onClick={sync} style={{ width: "100%", padding: "5px 10px", background: "none", border: "1px solid rgba(255,255,255,0.09)", borderRadius: 6, color: "rgba(255,255,255,0.35)", fontSize: 11, cursor: "pointer" }}>↻ Sync now</button>
      </div>
    </aside>
  );
}

// ═══════════════════════════════════════════════════════
// MAIN APP
// ═══════════════════════════════════════════════════════

return function SiteDocGeneratorView() {
  const [sites, setSites] = useState([]);
  const [sel, setSel] = useState(null);
  const [syncSt, setSyncSt] = useState("demo");
  const [view, setView] = useState("dash");
  const [docType, setDocType] = useState(null);
  const [fd, setFd] = useState({});
  const [docs, setDocs] = useState([]);
  const [activeCat, setActiveCat] = useState("safety");

  const sync = useCallback(async () => {
    // Sites data comes from the labour schedule app via props
    // Falls back to built-in demo sites when no live data available
    setSites(DEMO_SITES); setSyncSt("demo");
  }, []);

  useEffect(() => {
    sync();
    loadDocs().then(setDocs);
    const iv = setInterval(sync, 30000);
    return () => clearInterval(iv);
  }, [sync]);

  useEffect(() => { if (sites.length && !sel) setSel(sites[0]); }, [sites, sel]);

  const start = dt => { setDocType(dt); setFd(initFd(dt, sel)); setView("form"); };

  const onGen = async formData => {
    const doc = {
      id: genId(), docType: docType.id, docLabel: docType.label, category: docType.cat,
      site: formData.site, siteId: sel?.id, generatedBy: formData.preparedBy,
      generatedAt: new Date().toISOString(), date: formData.date,
      responsible: (formData.attendees || []).filter(a => a.responsible).map(a => a.name),
      team: (formData.attendees || []).filter(a => a.present).map(a => `${a.name} (${a.role})`),
      formData,
    };
    const updated = await saveDoc(doc, docs);
    setDocs(updated); setFd(formData); setView("preview");
  };

  if (view === "form") return <DocForm dt={docType} initFd={fd} onBack={() => setView("dash")} onGen={onGen} />;
  if (view === "preview") return <Preview dt={docType} fd={fd} onBack={() => setView("form")} onNew={() => setView("dash")} onRep={() => setView("reports")} />;
  if (view === "reports") return (
    <Reports docs={docs} sites={sites} onBack={() => setView("dash")}
      onView={doc => { setDocType(DOC_TYPES.find(d => d.id === doc.docType)); setFd(doc.formData); setView("preview"); }} />
  );

  return (
    <div style={{ display: "flex", minHeight: "100vh", fontFamily: "system-ui, -apple-system, sans-serif", background: BG }}>
      <Sidebar sites={sites} sel={sel} setSel={setSel} sync={sync} syncSt={syncSt} onRep={() => setView("reports")} docsCount={docs.length} />
      <main style={{ flex: 1, padding: "22px 26px", overflowY: "auto" }}>
        <div style={{ marginBottom: 18 }}>
          <h1 style={{ margin: 0, fontSize: 21, fontWeight: 700, color: GY8 }}>Site Document Generator</h1>
          {sel && <p style={{ margin: "3px 0 0", fontSize: 13, color: GY5 }}>{sel.name} · {sel.address}</p>}
        </div>
        {sel && (
          <div style={{ background: "#fff", border: `1px solid ${GY2}`, borderRadius: 10, padding: "12px 15px", marginBottom: 18 }}>
            <div style={{ display: "flex", gap: 20, flexWrap: "wrap" }}>
              <IBlock label="Scope" value={sel.scope} />
              <IBlock label="Supervisor" value={`${sel.supervisor?.name} · ${sel.supervisor?.phone}`} />
              <IBlock label="Workers" value={`${sel.workers?.length || 0} allocated`} />
            </div>
          </div>
        )}
        <div style={{ display: "flex", gap: 6, marginBottom: 16, flexWrap: "wrap" }}>
          {CATEGORIES.map(cat => (
            <button key={cat.id} onClick={() => setActiveCat(cat.id)} style={{
              padding: "6px 13px", borderRadius: 20, border: `1px solid ${activeCat === cat.id ? cat.color : GY2}`,
              background: activeCat === cat.id ? cat.color : "#fff", fontSize: 12, cursor: "pointer",
              color: activeCat === cat.id ? (cat.id === "safety" ? DARK : "#fff") : GY6,
              fontWeight: activeCat === cat.id ? 700 : 400,
            }}>{cat.label}</button>
          ))}
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(195px, 1fr))", gap: 10 }}>
          {DOC_TYPES.filter(d => d.cat === activeCat).map(dt => (
            <DocCard key={dt.id} dt={dt} onClick={() => start(dt)}
              count={docs.filter(d => d.docType === dt.id && d.siteId === sel?.id).length} />
          ))}
        </div>
        {syncSt === "demo" && (
          <div style={{ marginTop: 22, padding: "10px 14px", background: "#FFFBEB", border: "1px solid #FCD34D", borderRadius: 8, fontSize: 12, color: "#92400E" }}>
            <b>Demo mode.</b> Connect your Labour Schedule app — save to storage key <code style={{ background: "rgba(0,0,0,0.06)", padding: "1px 5px", borderRadius: 3 }}>labour_schedule_data</code>.
          </div>
        )}
      </main>
    </div>
  );
}

// ═══════════════════════════════════════════════════════
// DOC FORM
// ═══════════════════════════════════════════════════════

function DocForm({ dt, initFd: init, onBack, onGen }) {
  const [fd, setFd] = useState(init);
  const set = (k, v) => setFd(f => ({ ...f, [k]: v }));
  const chk = (i, f, v) => setFd(d => ({ ...d, checks: { ...d.checks, [i]: { ...d.checks[i], [f]: v } } }));
  const qa = (i, f, v) => setFd(d => ({ ...d, qaItems: d.qaItems.map((it, idx) => idx === i ? { ...it, [f]: v } : it) }));
  const tog = (id, fld) => setFd(d => ({ ...d, attendees: d.attendees.map(a => a.id === id ? { ...a, [fld]: !a[fld] } : a) }));

  const isCl = dt.cat === "inspection";
  const cl = CL[dt.id];
  const noAtt = dt.cat === "legal";

  const RadioRow = ({ label, field }) => (
    <tr style={{ borderBottom: `1px solid ${GY2}` }}>
      <td style={{ padding: "7px 10px", color: GY8, fontSize: 13 }}>{label}</td>
      {["pass", "fail", "na"].map(r => (
        <td key={r} style={{ padding: "7px 10px", textAlign: "center" }}>
          <input type="radio" name={field} checked={fd[field] === r} onChange={() => set(field, r)}
            style={{ accentColor: r === "pass" ? GRN : r === "fail" ? RED : "#9CA3AF" }} />
        </td>
      ))}
    </tr>
  );

  return (
    <div style={{ fontFamily: "system-ui, -apple-system, sans-serif", background: BG, minHeight: "100vh" }}>
      <div style={{ background: HDR, padding: "0 20px", height: 52, display: "flex", alignItems: "center", gap: 13, position: "sticky", top: 0, zIndex: 10 }}>
        <button onClick={onBack} style={{ background: "none", border: "none", color: "rgba(255,255,255,0.45)", cursor: "pointer", fontSize: 18, padding: 0 }}>←</button>
        <div style={{ flex: 1 }}>
          <p style={{ margin: 0, fontSize: 10, color: "rgba(255,255,255,0.38)" }}>New document</p>
          <p style={{ margin: 0, fontSize: 14, fontWeight: 700, color: "#fff" }}>{dt.label}</p>
        </div>
        <button onClick={() => onGen(fd)} style={{ padding: "7px 18px", background: AMB, border: "none", borderRadius: 6, fontWeight: 700, fontSize: 13, cursor: "pointer", color: DARK }}>Generate →</button>
      </div>

      <div style={{ maxWidth: 760, margin: "0 auto", padding: "18px 16px 40px" }}>
        {/* Always: site details */}
        <Sec title="Site details">
          <Grid>
            <Field label="Site name" value={fd.site} onChange={v => set("site", v)} />
            <Field label="Date" value={fd.date} onChange={v => set("date", v)} />
            <Field label="Address" value={fd.address} onChange={v => set("address", v)} span />
            <Field label="Scope of works" value={fd.scope} onChange={v => set("scope", v)} span multi />
            <Field label="Supervisor" value={fd.supervisor} onChange={v => set("supervisor", v)} />
            <Field label="Contact number" value={fd.supPhone} onChange={v => set("supPhone", v)} />
            <Field label="Prepared by" value={fd.preparedBy} onChange={v => set("preparedBy", v)} />
          </Grid>
        </Sec>

        {/* Inspection: equipment */}
        {isCl && (
          <Sec title="Equipment details">
            <Grid cols={3}>
              <Field label="Equipment ID / serial" value={fd.equipmentId} onChange={v => set("equipmentId", v)} />
              <Field label="Make" value={fd.make} onChange={v => set("make", v)} />
              <Field label="Model" value={fd.model} onChange={v => set("model", v)} />
            </Grid>
          </Sec>
        )}

        {/* Daily Safe Start */}
        {dt.id === "daily_safe_start" && (
          <Sec title="Briefing content">
            <Grid cols={1}>
              <Field label="Hazards identified" value={fd.hazards} onChange={v => set("hazards", v)} multi />
              <Field label="Control measures in place" value={fd.controls} onChange={v => set("controls", v)} multi />
              <Field label="PPE required" value={fd.ppe} onChange={v => set("ppe", v)} />
              <Field label="Emergency procedure / muster point" value={fd.emergencyProc} onChange={v => set("emergencyProc", v)} />
            </Grid>
          </Sec>
        )}

        {/* Toolbox Talk */}
        {dt.id === "toolbox_talk" && (
          <Sec title="Talk content">
            <Grid cols={1}>
              <Field label="Topic / title" value={fd.topic} onChange={v => set("topic", v)} />
              <Field label="Key points covered" value={fd.hazards} onChange={v => set("hazards", v)} multi />
              <Field label="Actions arising / follow-up" value={fd.controls} onChange={v => set("controls", v)} multi />
            </Grid>
          </Sec>
        )}

        {/* PUWER */}
        {dt.id === "puwer" && (<>
          <Sec title="Equipment">
            <Grid>
              <Field label="Equipment / machine" value={fd.equipment} onChange={v => set("equipment", v)} />
              <Field label="Equipment ID / ref" value={fd.equipmentId} onChange={v => set("equipmentId", v)} />
              <Field label="Make" value={fd.make} onChange={v => set("make", v)} />
              <Field label="Model" value={fd.model} onChange={v => set("model", v)} />
            </Grid>
          </Sec>
          <Sec title="Risk assessment & controls">
            <Grid cols={1}>
              <Field label="Hazards associated with equipment" value={fd.hazards} onChange={v => set("hazards", v)} multi />
              <Field label="Risk level (High / Medium / Low)" value={fd.riskLevel} onChange={v => set("riskLevel", v)} />
              <Field label="Control measures" value={fd.controls} onChange={v => set("controls", v)} multi />
              <Field label="Training required / provided" value={fd.training} onChange={v => set("training", v)} multi />
              <Field label="Maintenance schedule" value={fd.maintenance} onChange={v => set("maintenance", v)} />
              <Field label="Inspection frequency" value={fd.frequency} onChange={v => set("frequency", v)} />
            </Grid>
          </Sec>
        </>)}

        {/* LOLER */}
        {dt.id === "loler" && (<>
          <Sec title="Lifting equipment">
            <Grid>
              <Field label="Equipment type" value={fd.liftingEquip} onChange={v => set("liftingEquip", v)} />
              <Field label="SWL / WLL" value={fd.swl} onChange={v => set("swl", v)} />
              <Field label="Equipment ID" value={fd.equipmentId} onChange={v => set("equipmentId", v)} />
              <Field label="Certificate / exam ref" value={fd.examRef} onChange={v => set("examRef", v)} />
              <Field label="Last examination date" value={fd.examDate} onChange={v => set("examDate", v)} />
              <Field label="Next examination due" value={fd.nextExam} onChange={v => set("nextExam", v)} />
            </Grid>
          </Sec>
          <Sec title="Lift plan">
            <Grid cols={1}>
              <Field label="Lift description / method" value={fd.liftPlan} onChange={v => set("liftPlan", v)} multi />
              <Field label="Lifting equipment to be used" value={fd.resources} onChange={v => set("resources", v)} multi />
              <Field label="Hazards / special precautions" value={fd.hazards} onChange={v => set("hazards", v)} multi />
            </Grid>
          </Sec>
        </>)}

        {/* COSHH */}
        {dt.id === "coshh" && (<>
          <Sec title="Substance">
            <Grid>
              <Field label="Substance name" value={fd.substance} onChange={v => set("substance", v)} />
              <Field label="Supplier / manufacturer" value={fd.supplier} onChange={v => set("supplier", v)} />
              <Field label="Hazard type (GHS classification)" value={fd.hazardType} onChange={v => set("hazardType", v)} />
              <Field label="WEL (workplace exposure limit)" value={fd.exposureLimits} onChange={v => set("exposureLimits", v)} />
              <Field label="Health effects of exposure" value={fd.healthEffects} onChange={v => set("healthEffects", v)} span multi />
            </Grid>
          </Sec>
          <Sec title="Controls">
            <Grid cols={1}>
              <Field label="Exposure routes (inhalation / dermal / ingestion)" value={fd.exposureRoutes} onChange={v => set("exposureRoutes", v)} />
              <Field label="Engineering / substitution controls" value={fd.cohhControls} onChange={v => set("cohhControls", v)} multi />
              <Field label="PPE required" value={fd.cohhPpe} onChange={v => set("cohhPpe", v)} />
              <Field label="Emergency actions (spillage / fire / first aid)" value={fd.emergencyActions} onChange={v => set("emergencyActions", v)} multi />
            </Grid>
          </Sec>
        </>)}

        {/* Task Briefing */}
        {dt.id === "task_briefing" && (
          <Sec title="Task details">
            <Grid cols={1}>
              <Field label="Task description" value={fd.taskDesc} onChange={v => set("taskDesc", v)} multi />
              <Field label="Method statement reference" value={fd.methodRef} onChange={v => set("methodRef", v)} />
              <Field label="Plant / equipment / materials required" value={fd.resources} onChange={v => set("resources", v)} multi />
              <Field label="Task-specific risks" value={fd.taskRisks} onChange={v => set("taskRisks", v)} multi />
              <Field label="Controls / safe working methods" value={fd.taskControls} onChange={v => set("taskControls", v)} multi />
            </Grid>
          </Sec>
        )}

        {/* PAT Testing */}
        {dt.id === "pat_testing" && (<>
          <Sec title="Appliance details">
            <Grid>
              <Field label="Appliance description" value={fd.applianceDesc} onChange={v => set("applianceDesc", v)} />
              <Field label="Asset / inventory number" value={fd.assetNo} onChange={v => set("assetNo", v)} />
              <Field label="Location on site" value={fd.location} onChange={v => set("location", v)} />
              <Field label="Tested by" value={fd.testedBy} onChange={v => set("testedBy", v)} />
            </Grid>
          </Sec>
          <Sec title="Test results">
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
              <thead><tr style={{ background: GY1 }}>
                <th style={{ padding: "7px 10px", textAlign: "left", fontWeight: 600, color: GY7 }}>Test</th>
                <th style={{ padding: "7px 10px", width: 65, textAlign: "center" }}>Pass</th>
                <th style={{ padding: "7px 10px", width: 65, textAlign: "center" }}>Fail</th>
                <th style={{ padding: "7px 10px", width: 65, textAlign: "center" }}>N/A</th>
              </tr></thead>
              <tbody>
                <RadioRow label="Visual inspection" field="visualPass" />
                <RadioRow label="Earth continuity" field="earthPass" />
                <RadioRow label="Insulation resistance" field="insulationPass" />
                <RadioRow label="Polarity check" field="polarityPass" />
              </tbody>
            </table>
            <div style={{ marginTop: 12, display: "flex", gap: 9 }}>
              {[{ v: "pass", l: "✓ Pass", bg: GBG, c: GRN }, { v: "fail", l: "✗ Fail", bg: RBG, c: RED }].map(o => (
                <button key={o.v} onClick={() => set("patResult", o.v)} style={{ padding: "7px 16px", borderRadius: 6, border: `1.5px solid ${fd.patResult === o.v ? o.c : GY2}`, background: fd.patResult === o.v ? o.bg : "#fff", color: fd.patResult === o.v ? o.c : GY5, fontWeight: fd.patResult === o.v ? 700 : 400, cursor: "pointer", fontSize: 13 }}>{o.l}</button>
              ))}
            </div>
            <div style={{ marginTop: 10 }}><Field label="Next test date" value={fd.nextTestDate} onChange={v => set("nextTestDate", v)} /></div>
          </Sec>
        </>)}

        {/* Thorough Examination */}
        {dt.id === "thorough_exam" && (<>
          <Sec title="Equipment">
            <Grid>
              <Field label="Equipment type" value={fd.examEquipment} onChange={v => set("examEquipment", v)} />
              <Field label="Equipment ID / serial" value={fd.examId} onChange={v => set("examId", v)} />
              <Field label="SWL / WLL" value={fd.examSWL} onChange={v => set("examSWL", v)} />
              <Field label="Examination reference" value={fd.examRef} onChange={v => set("examRef", v)} />
            </Grid>
          </Sec>
          <Sec title="Examination details">
            <Grid>
              <Field label="Examiner name" value={fd.examiner} onChange={v => set("examiner", v)} />
              <Field label="Examination body / company" value={fd.examBody} onChange={v => set("examBody", v)} />
              <Field label="Date of examination" value={fd.examDate} onChange={v => set("examDate", v)} />
              <Field label="Next examination due" value={fd.nextExamDate} onChange={v => set("nextExamDate", v)} />
              <Field label="Condition rating (1–4)" value={fd.conditionRating} onChange={v => set("conditionRating", v)} />
            </Grid>
            <div style={{ marginTop: 10 }}><Field label="Defects found (or 'None')" value={fd.defectsDesc} onChange={v => set("defectsDesc", v)} multi /></div>
            <div style={{ display: "flex", gap: 8, marginTop: 10, flexWrap: "wrap" }}>
              {[{ v: "yes", l: "✓ Safe for use" }, { v: "no", l: "✗ Not safe – remove" }, { v: "conditions", l: "⚠ Conditional" }].map(o => (
                <button key={o.v} onClick={() => set("safeForUse", o.v)} style={{ padding: "6px 12px", borderRadius: 6, border: `1.5px solid ${fd.safeForUse === o.v ? BLU : GY2}`, background: fd.safeForUse === o.v ? BBG : "#fff", color: fd.safeForUse === o.v ? BLU : GY5, fontWeight: fd.safeForUse === o.v ? 600 : 400, cursor: "pointer", fontSize: 12 }}>{o.l}</button>
              ))}
            </div>
          </Sec>
        </>)}

        {/* QA Handover */}
        {dt.id === "qa_handover" && (<>
          <Sec title="Contract details">
            <Grid>
              <Field label="Contract reference" value={fd.contractRef} onChange={v => set("contractRef", v)} />
              <Field label="Work package" value={fd.workPackage} onChange={v => set("workPackage", v)} />
            </Grid>
          </Sec>
          <Sec title="Quality checklist">
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
              <thead><tr style={{ background: GY1 }}>
                <th style={{ padding: "7px 10px", textAlign: "left", fontWeight: 600, color: GY7 }}>Item</th>
                <th style={{ padding: "7px 10px", width: 55, textAlign: "center" }}>Pass</th>
                <th style={{ padding: "7px 10px", width: 55, textAlign: "center" }}>Fail</th>
                <th style={{ padding: "7px 10px", width: 55, textAlign: "center" }}>N/A</th>
                <th style={{ padding: "7px 10px", width: 140, textAlign: "left" }}>Notes</th>
              </tr></thead>
              <tbody>
                {fd.qaItems?.map((it, i) => (
                  <tr key={i} style={{ borderBottom: `1px solid ${GY2}` }}>
                    <td style={{ padding: "7px 10px", color: GY8 }}>{it.item}</td>
                    {["pass", "fail", "na"].map(r => (
                      <td key={r} style={{ padding: "7px 10px", textAlign: "center" }}>
                        <input type="radio" name={`qa${i}`} checked={it.result === r} onChange={() => qa(i, "result", r)} />
                      </td>
                    ))}
                    <td style={{ padding: "7px 6px" }}>
                      <input type="text" value={it.notes || ""} onChange={e => qa(i, "notes", e.target.value)} placeholder="Notes…" style={{ ...inp, fontSize: 11, padding: "3px 6px" }} />
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
            <div style={{ marginTop: 10 }}><Field label="Deficiencies / outstanding items" value={fd.deficiencies} onChange={v => set("deficiencies", v)} multi /></div>
          </Sec>
        </>)}

        {/* NCR */}
        {dt.id === "ncr" && (
          <Sec title="Non-conformance details">
            <Grid>
              <Field label="NCR reference number" value={fd.ncrRef} onChange={v => set("ncrRef", v)} />
              <Field label="Target closure date" value={fd.targetDate} onChange={v => set("targetDate", v)} />
              <Field label="Description of non-conformance" value={fd.ncrDesc} onChange={v => set("ncrDesc", v)} span multi />
              <Field label="Root cause analysis" value={fd.rootCause} onChange={v => set("rootCause", v)} span multi />
              <Field label="Corrective action required" value={fd.correctiveAction} onChange={v => set("correctiveAction", v)} span multi />
              <Field label="Closed by" value={fd.closedBy} onChange={v => set("closedBy", v)} />
            </Grid>
          </Sec>
        )}

        {/* Early Delay / Delay Notice */}
        {(dt.id === "early_delay" || dt.id === "delay_notice") && (<>
          <Sec title="Contract details">
            <Grid>
              <Field label="Contract number" value={fd.contractNo} onChange={v => set("contractNo", v)} />
              <Field label="Programme reference" value={fd.programmeRef} onChange={v => set("programmeRef", v)} />
              <Field label="Employer" value={fd.employer} onChange={v => set("employer", v)} />
              <Field label="Contractor" value={fd.contractor} onChange={v => set("contractor", v)} />
            </Grid>
          </Sec>
          <Sec title="Event details">
            <Grid>
              <Field label="Date of event" value={fd.eventDate} onChange={v => set("eventDate", v)} />
              <Field label="Reply required by" value={fd.replyBy} onChange={v => set("replyBy", v)} />
              <Field label="Description of delay event" value={fd.eventDesc} onChange={v => set("eventDesc", v)} span multi />
              <Field label="Delay (weeks)" value={fd.delayWeeks} onChange={v => set("delayWeeks", v)} />
              {dt.id === "delay_notice" && <Field label="Cost impact (£)" value={fd.costImpact} onChange={v => set("costImpact", v)} />}
              <Field label="Programme impact / consequences" value={fd.delayImpact} onChange={v => set("delayImpact", v)} span multi />
            </Grid>
          </Sec>
        </>)}

        {/* Inspection checklist */}
        {isCl && cl && (<>
          <Sec title="Inspection checklist">
            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                <thead><tr style={{ background: GY1 }}>
                  {cl.s && <th style={{ padding: "6px 8px", textAlign: "left", fontWeight: 600, color: GY7, width: 100 }}>Section</th>}
                  <th style={{ padding: "6px 8px", textAlign: "left", fontWeight: 600, color: GY7 }}>Item</th>
                  <th style={{ padding: "6px 8px", width: 44, textAlign: "center" }}>Pass</th>
                  <th style={{ padding: "6px 8px", width: 44, textAlign: "center" }}>Fail</th>
                  <th style={{ padding: "6px 8px", width: 44, textAlign: "center" }}>N/A</th>
                  <th style={{ padding: "6px 8px", width: 130, textAlign: "left" }}>Notes</th>
                </tr></thead>
                <tbody>
                  {cl.items.map((item, i) => (
                    <tr key={i} style={{ borderBottom: `1px solid ${GY2}` }}>
                      {cl.s && <td style={{ padding: "6px 8px", fontSize: 10, color: GY5, fontWeight: 500, verticalAlign: "top" }}>{itemSection(dt.id, i)}</td>}
                      <td style={{ padding: "6px 8px", color: GY8, verticalAlign: "top" }}>{item}</td>
                      {["pass", "fail", "na"].map(r => (
                        <td key={r} style={{ padding: "6px 8px", textAlign: "center", verticalAlign: "top" }}>
                          <input type="radio" name={`c${i}`} checked={fd.checks?.[i]?.r === r} onChange={() => chk(i, "r", r)}
                            style={{ accentColor: r === "pass" ? GRN : r === "fail" ? RED : "#9CA3AF" }} />
                        </td>
                      ))}
                      <td style={{ padding: "6px 5px" }}>
                        <input type="text" value={fd.checks?.[i]?.n || ""} onChange={e => chk(i, "n", e.target.value)}
                          placeholder="Notes…" style={{ width: "100%", fontSize: 11, border: `1px solid ${GY2}`, borderRadius: 4, padding: "3px 5px", background: "#fff" }} />
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </Sec>
          <Sec title="Overall result">
            <div style={{ display: "flex", gap: 9, flexWrap: "wrap" }}>
              {[
                { v: "safe",   l: "✓ Safe to use",         bg: GBG,      c: GRN },
                { v: "remove", l: "✗ Remove from service", bg: RBG,      c: RED },
                { v: "monitor",l: "⚠ Monitor",             bg: "#FFF7ED", c: "#C2410C" },
              ].map(o => (
                <button key={o.v} onClick={() => set("result", o.v)} style={{ padding: "7px 13px", borderRadius: 6, border: `1.5px solid ${fd.result === o.v ? o.c : GY2}`, background: fd.result === o.v ? o.bg : "#fff", color: fd.result === o.v ? o.c : GY6, fontWeight: fd.result === o.v ? 700 : 400, cursor: "pointer", fontSize: 12 }}>{o.l}</button>
              ))}
            </div>
            <div style={{ marginTop: 10 }}><Field label="Additional notes" value={fd.notes} onChange={v => set("notes", v)} multi /></div>
          </Sec>
        </>)}

        {/* Attendees & responsible signatories */}
        {!noAtt && fd.attendees?.length > 0 && (
          <Sec title="Team & responsible signatories">
            <p style={{ margin: "0 0 9px", fontSize: 11, color: GY5 }}>Tick who is present. Mark as responsible the person(s) who sign off this document.</p>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(185px, 1fr))", gap: 8 }}>
              {fd.attendees.map(a => (
                <div key={a.id} style={{ padding: "9px 11px", background: a.present ? GBG : GY1, border: `1px solid ${a.present ? GRL : GY2}`, borderRadius: 7 }}>
                  <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 6 }}>
                    <div>
                      <p style={{ margin: 0, fontSize: 12, fontWeight: 600, color: a.present ? GRN : GY6 }}>{a.name}</p>
                      <p style={{ margin: 0, fontSize: 10, color: a.present ? "#16A34A" : GY5 }}>{a.role}</p>
                    </div>
                    <label style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 1, cursor: "pointer" }}>
                      <input type="checkbox" checked={a.present} onChange={() => tog(a.id, "present")} style={{ accentColor: GRN }} />
                      <span style={{ fontSize: 9, color: "#9CA3AF" }}>Present</span>
                    </label>
                  </div>
                  {a.present && (
                    <label style={{ display: "flex", alignItems: "center", gap: 5, marginTop: 6, cursor: "pointer" }}>
                      <input type="checkbox" checked={a.responsible || false} onChange={() => tog(a.id, "responsible")} style={{ accentColor: AMB }} />
                      <span style={{ fontSize: 10, color: GY6 }}>Responsible signatory</span>
                    </label>
                  )}
                </div>
              ))}
            </div>
          </Sec>
        )}

        <div style={{ display: "flex", justifyContent: "flex-end", gap: 9, paddingTop: 4 }}>
          <button onClick={onBack} style={{ padding: "8px 18px", border: `1px solid ${GY3}`, background: "#fff", borderRadius: 6, cursor: "pointer", fontSize: 13, color: GY7 }}>Cancel</button>
          <button onClick={() => onGen(fd)} style={{ padding: "8px 22px", background: AMB, border: "none", borderRadius: 6, cursor: "pointer", fontSize: 13, fontWeight: 700, color: DARK }}>Generate document →</button>
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════
// PREVIEW
// ═══════════════════════════════════════════════════════

function Preview({ dt, fd, onBack, onNew, onRep }) {
  const isCl = dt.cat === "inspection";
  const cl = CL[dt.id];
  const resp = (fd.attendees || []).filter(a => a.responsible);
  const team = (fd.attendees || []).filter(a => a.present);
  const noAtt = dt.cat === "legal";
  const cc = CAT_COLOR[dt.cat];

  useEffect(() => {
    const el = document.createElement("style");
    el.id = "piso";
    el.textContent = `@media print{body *{visibility:hidden}#pdoc,#pdoc *{visibility:visible}#pdoc{position:fixed;top:0;left:0;width:100%;padding:16px}@page{margin:12mm}}`;
    document.head.appendChild(el);
    return () => document.getElementById("piso")?.remove();
  }, []);

  const tb = { width: "100%", borderCollapse: "collapse", border: `1px solid ${GY2}` };

  function PR({ l, v }) {
    return (
      <tr>
        <td style={{ padding: "5px 10px", fontWeight: 600, color: GY7, fontSize: 12, width: "32%", verticalAlign: "top" }}>{l}</td>
        <td style={{ padding: "5px 10px", color: GY8, fontSize: 12, whiteSpace: "pre-wrap" }}>{v || "—"}</td>
      </tr>
    );
  }

  function PS({ title, children }) {
    return (
      <div style={{ marginBottom: 18 }}>
        <p style={{ margin: "0 0 6px", fontSize: 11, fontWeight: 700, color: GY5, textTransform: "uppercase", letterSpacing: "0.07em" }}>{title}</p>
        {children}
      </div>
    );
  }

  function Badge({ val, pass = "pass", fail = "fail" }) {
    const isP = val === pass, isF = val === fail;
    return <span style={{ padding: "2px 8px", borderRadius: 4, fontSize: 11, fontWeight: 700, background: isP ? GBG : isF ? RBG : GY1, color: isP ? GRN : isF ? RED : GY5 }}>{val?.toUpperCase() || "—"}</span>;
  }

  return (
    <div style={{ fontFamily: "system-ui, -apple-system, sans-serif", background: BG, minHeight: "100vh" }}>
      <div style={{ background: HDR, padding: "0 20px", height: 52, display: "flex", alignItems: "center", gap: 12, position: "sticky", top: 0, zIndex: 10 }}>
        <button onClick={onBack} style={{ background: "none", border: "none", color: "rgba(255,255,255,0.45)", cursor: "pointer", fontSize: 18, padding: 0 }}>←</button>
        <span style={{ color: "#fff", fontSize: 14, fontWeight: 700, flex: 1 }}>{dt.label}</span>
        <div style={{ display: "flex", gap: 7 }}>
          <button onClick={onRep} style={{ padding: "6px 12px", background: "none", border: "1px solid rgba(255,255,255,0.22)", borderRadius: 6, color: "rgba(255,255,255,0.75)", cursor: "pointer", fontSize: 11 }}>Reports</button>
          <button onClick={onNew} style={{ padding: "6px 12px", background: "none", border: "1px solid rgba(255,255,255,0.22)", borderRadius: 6, color: "rgba(255,255,255,0.75)", cursor: "pointer", fontSize: 11 }}>New doc</button>
          <button onClick={() => window.print()} style={{ padding: "6px 14px", background: AMB, border: "none", borderRadius: 6, fontWeight: 700, fontSize: 12, cursor: "pointer", color: DARK }}>🖨 Print / PDF</button>
        </div>
      </div>

      <div style={{ maxWidth: 820, margin: "18px auto 40px", padding: "0 14px" }}>
        <div id="pdoc" style={{ background: "#fff", border: `1px solid ${GY2}`, borderRadius: 10, overflow: "hidden" }}>
          {/* Document header */}
          <div style={{ borderTop: `4px solid ${cc}`, padding: "18px 22px 13px", borderBottom: `1px solid ${GY2}` }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
              <div>
                <p style={{ margin: 0, fontSize: 10, color: "#9CA3AF", textTransform: "uppercase", letterSpacing: "0.09em" }}>Construction site document</p>
                <h1 style={{ margin: "3px 0 0", fontSize: 21, fontWeight: 800, color: GY8 }}>{dt.label}</h1>
              </div>
              <div style={{ textAlign: "right" }}>
                <p style={{ margin: 0, fontSize: 10, color: "#9CA3AF" }}>Date issued</p>
                <p style={{ margin: 0, fontSize: 15, fontWeight: 700, color: GY8 }}>{fd.date}</p>
              </div>
            </div>
          </div>

          <div style={{ padding: "18px 22px" }}>
            {/* Site details — always */}
            <PS title="Site details">
              <table style={tb}><tbody>
                <PR l="Site" v={fd.site} />
                <PR l="Address" v={fd.address} />
                <PR l="Scope of works" v={fd.scope} />
                <PR l="Supervisor" v={`${fd.supervisor}${fd.supPhone ? ` · ${fd.supPhone}` : ""}`} />
                <PR l="Prepared by" v={fd.preparedBy} />
              </tbody></table>
            </PS>

            {/* Doc-specific content */}
            {(dt.id === "daily_safe_start" || dt.id === "toolbox_talk") && (
              <PS title={dt.id === "toolbox_talk" ? "Talk content" : "Briefing content"}>
                <table style={tb}><tbody>
                  {dt.id === "toolbox_talk" && <PR l="Topic" v={fd.topic} />}
                  <PR l={dt.id === "toolbox_talk" ? "Key points" : "Hazards identified"} v={fd.hazards} />
                  <PR l={dt.id === "toolbox_talk" ? "Actions arising" : "Control measures"} v={fd.controls} />
                  {dt.id === "daily_safe_start" && <><PR l="PPE required" v={fd.ppe} /><PR l="Emergency procedure" v={fd.emergencyProc} /></>}
                </tbody></table>
              </PS>
            )}
            {dt.id === "puwer" && (
              <PS title="PUWER assessment">
                <table style={tb}><tbody>
                  <PR l="Equipment" v={fd.equipment} /><PR l="ID / ref" v={fd.equipmentId} />
                  <PR l="Make / model" v={`${fd.make} ${fd.model}`.trim()} /><PR l="Hazards" v={fd.hazards} />
                  <PR l="Risk level" v={fd.riskLevel} /><PR l="Controls" v={fd.controls} />
                  <PR l="Training" v={fd.training} /><PR l="Maintenance" v={fd.maintenance} />
                  <PR l="Inspection frequency" v={fd.frequency} />
                </tbody></table>
              </PS>
            )}
            {dt.id === "loler" && (
              <PS title="LOLER details">
                <table style={tb}><tbody>
                  <PR l="Lifting equipment" v={fd.liftingEquip} /><PR l="SWL / WLL" v={fd.swl} />
                  <PR l="Equipment ID" v={fd.equipmentId} /><PR l="Certificate ref" v={fd.examRef} />
                  <PR l="Last examination" v={fd.examDate} /><PR l="Next examination" v={fd.nextExam} />
                  <PR l="Lift plan" v={fd.liftPlan} /><PR l="Equipment to be used" v={fd.resources} />
                  <PR l="Hazards / precautions" v={fd.hazards} />
                </tbody></table>
              </PS>
            )}
            {dt.id === "coshh" && (
              <PS title="COSHH assessment">
                <table style={tb}><tbody>
                  <PR l="Substance" v={fd.substance} /><PR l="Supplier" v={fd.supplier} />
                  <PR l="Hazard type" v={fd.hazardType} /><PR l="WEL" v={fd.exposureLimits} />
                  <PR l="Health effects" v={fd.healthEffects} /><PR l="Exposure routes" v={fd.exposureRoutes} />
                  <PR l="Controls" v={fd.cohhControls} /><PR l="PPE required" v={fd.cohhPpe} />
                  <PR l="Emergency actions" v={fd.emergencyActions} />
                </tbody></table>
              </PS>
            )}
            {dt.id === "task_briefing" && (
              <PS title="Task briefing">
                <table style={tb}><tbody>
                  <PR l="Task description" v={fd.taskDesc} /><PR l="Method statement ref" v={fd.methodRef} />
                  <PR l="Resources required" v={fd.resources} /><PR l="Task risks" v={fd.taskRisks} />
                  <PR l="Safe working methods" v={fd.taskControls} />
                </tbody></table>
              </PS>
            )}
            {dt.id === "pat_testing" && (
              <PS title="PAT testing record">
                <table style={tb}><tbody>
                  <PR l="Appliance" v={fd.applianceDesc} /><PR l="Asset number" v={fd.assetNo} />
                  <PR l="Location" v={fd.location} /><PR l="Tested by" v={fd.testedBy} />
                  {[["Visual inspection", fd.visualPass], ["Earth continuity", fd.earthPass], ["Insulation resistance", fd.insulationPass], ["Polarity check", fd.polarityPass]].map(([l, v]) => (
                    <tr key={l}>
                      <td style={{ padding: "5px 10px", fontWeight: 600, color: GY7, fontSize: 12, width: "32%" }}>{l}</td>
                      <td style={{ padding: "5px 10px" }}><Badge val={v} /></td>
                    </tr>
                  ))}
                  <tr>
                    <td style={{ padding: "5px 10px", fontWeight: 600, color: GY7, fontSize: 12 }}>Overall result</td>
                    <td style={{ padding: "5px 10px" }}><span style={{ padding: "3px 11px", borderRadius: 4, fontSize: 12, fontWeight: 800, background: fd.patResult === "pass" ? GBG : RBG, color: fd.patResult === "pass" ? GRN : RED }}>{fd.patResult?.toUpperCase() || "—"}</span></td>
                  </tr>
                  <PR l="Next test date" v={fd.nextTestDate} />
                </tbody></table>
              </PS>
            )}
            {dt.id === "thorough_exam" && (
              <PS title="Thorough examination">
                <table style={tb}><tbody>
                  <PR l="Equipment" v={fd.examEquipment} /><PR l="ID / serial" v={fd.examId} />
                  <PR l="SWL" v={fd.examSWL} /><PR l="Exam reference" v={fd.examRef} />
                  <PR l="Examiner" v={fd.examiner} /><PR l="Examination body" v={fd.examBody} />
                  <PR l="Date of examination" v={fd.examDate} /><PR l="Condition rating" v={fd.conditionRating} />
                  <PR l="Defects found" v={fd.defectsDesc || "None"} />
                  <tr>
                    <td style={{ padding: "5px 10px", fontWeight: 600, color: GY7, fontSize: 12 }}>Safe for use</td>
                    <td style={{ padding: "5px 10px" }}>
                      <span style={{ padding: "3px 10px", borderRadius: 4, fontSize: 11, fontWeight: 700, background: fd.safeForUse === "yes" ? GBG : fd.safeForUse === "conditions" ? "#FFF7ED" : RBG, color: fd.safeForUse === "yes" ? GRN : fd.safeForUse === "conditions" ? "#C2410C" : RED }}>
                        {fd.safeForUse === "yes" ? "YES – SAFE" : fd.safeForUse === "conditions" ? "CONDITIONAL" : "NO – REMOVE"}
                      </span>
                    </td>
                  </tr>
                  <PR l="Next examination" v={fd.nextExamDate} />
                </tbody></table>
              </PS>
            )}
            {dt.id === "qa_handover" && (
              <PS title="QA handover checklist">
                <p style={{ margin: "0 0 6px", fontSize: 12, color: GY6 }}>Contract: {fd.contractRef || "—"} · Work package: {fd.workPackage || "—"}</p>
                <table style={{ ...tb, fontSize: 12 }}><thead><tr style={{ background: GY1 }}>
                  <th style={{ padding: "6px 10px", textAlign: "left", fontWeight: 600 }}>Item</th>
                  <th style={{ padding: "6px 10px", width: 60, textAlign: "center", fontWeight: 600 }}>Result</th>
                  <th style={{ padding: "6px 10px", width: 160, textAlign: "left", fontWeight: 600 }}>Notes</th>
                </tr></thead><tbody>
                  {(fd.qaItems || []).map((it, i) => (
                    <tr key={i} style={{ borderBottom: `1px solid ${GY2}` }}>
                      <td style={{ padding: "5px 10px" }}>{it.item}</td>
                      <td style={{ padding: "5px 10px", textAlign: "center" }}><Badge val={it.result} /></td>
                      <td style={{ padding: "5px 10px", color: GY5, fontSize: 11 }}>{it.notes || "—"}</td>
                    </tr>
                  ))}
                </tbody></table>
                {fd.deficiencies && <p style={{ marginTop: 7, fontSize: 12, color: GY7 }}><b>Deficiencies:</b> {fd.deficiencies}</p>}
              </PS>
            )}
            {dt.id === "ncr" && (
              <PS title="Non-conformance report">
                <table style={tb}><tbody>
                  <PR l="NCR reference" v={fd.ncrRef} /><PR l="Description" v={fd.ncrDesc} />
                  <PR l="Root cause" v={fd.rootCause} /><PR l="Corrective action" v={fd.correctiveAction} />
                  <PR l="Target date" v={fd.targetDate} /><PR l="Closed by" v={fd.closedBy} />
                </tbody></table>
              </PS>
            )}
            {(dt.id === "early_delay" || dt.id === "delay_notice") && (
              <PS title={dt.label}>
                <table style={tb}><tbody>
                  <PR l="Contract number" v={fd.contractNo} /><PR l="Programme ref" v={fd.programmeRef} />
                  <PR l="Employer" v={fd.employer} /><PR l="Contractor" v={fd.contractor} />
                  <PR l="Date of event" v={fd.eventDate} /><PR l="Description" v={fd.eventDesc} />
                  <PR l="Delay (weeks)" v={fd.delayWeeks} />
                  {dt.id === "delay_notice" && <PR l="Cost impact (£)" v={fd.costImpact} />}
                  <PR l="Programme impact" v={fd.delayImpact} /><PR l="Reply by" v={fd.replyBy} />
                </tbody></table>
              </PS>
            )}

            {/* Inspection checklist */}
            {isCl && cl && (
              <PS title="Inspection checklist">
                <table style={{ ...tb, fontSize: 11 }}>
                  <thead><tr style={{ background: GY1 }}>
                    {cl.s && <th style={{ padding: "5px 8px", textAlign: "left", fontWeight: 600, width: 95 }}>Section</th>}
                    <th style={{ padding: "5px 8px", textAlign: "left", fontWeight: 600 }}>Item</th>
                    <th style={{ padding: "5px 8px", width: 52, textAlign: "center" }}>Result</th>
                    <th style={{ padding: "5px 8px", width: 140, textAlign: "left" }}>Notes</th>
                  </tr></thead>
                  <tbody>
                    {cl.items.map((item, i) => {
                      const r = fd.checks?.[i]?.r;
                      return (
                        <tr key={i} style={{ borderBottom: `1px solid ${GY2}` }}>
                          {cl.s && <td style={{ padding: "4px 8px", fontSize: 10, color: GY5, fontWeight: 500 }}>{itemSection(dt.id, i)}</td>}
                          <td style={{ padding: "4px 8px", color: GY7 }}>{item}</td>
                          <td style={{ padding: "4px 8px", textAlign: "center" }}>
                            <span style={{ padding: "1px 6px", borderRadius: 3, fontSize: 10, fontWeight: 700, background: r === "pass" ? GBG : r === "fail" ? RBG : GY1, color: r === "pass" ? GRN : r === "fail" ? RED : GY5 }}>{r?.toUpperCase() || "—"}</span>
                          </td>
                          <td style={{ padding: "4px 8px", fontSize: 10, color: GY5 }}>{fd.checks?.[i]?.n || ""}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
                {fd.result && (
                  <div style={{ marginTop: 9, padding: "7px 12px", display: "inline-block", borderRadius: 6, background: fd.result === "safe" ? GBG : fd.result === "remove" ? RBG : "#FFF7ED" }}>
                    <span style={{ fontWeight: 800, fontSize: 12, color: fd.result === "safe" ? GRN : fd.result === "remove" ? RED : "#C2410C" }}>
                      Overall: {fd.result === "safe" ? "SAFE TO USE" : fd.result === "remove" ? "REMOVE FROM SERVICE" : "MONITOR / CONDITIONAL"}
                    </span>
                  </div>
                )}
                {fd.notes && <p style={{ marginTop: 8, fontSize: 12, color: GY7 }}><b>Notes:</b> {fd.notes}</p>}
              </PS>
            )}

            {/* Attendance register */}
            {!noAtt && team.length > 0 && (
              <PS title="Attendance register">
                <table style={{ ...tb, fontSize: 12 }}>
                  <thead><tr style={{ background: GY1 }}>
                    <th style={{ padding: "6px 10px", textAlign: "left", fontWeight: 600 }}>Name</th>
                    <th style={{ padding: "6px 10px", textAlign: "left", fontWeight: 600 }}>Role</th>
                    <th style={{ padding: "6px 10px", textAlign: "center", fontWeight: 600, width: 80 }}>Responsible</th>
                    <th style={{ padding: "6px 10px", textAlign: "left", fontWeight: 600, width: 180 }}>Signature</th>
                  </tr></thead>
                  <tbody>
                    {team.map((a, i) => (
                      <tr key={i} style={{ borderBottom: `1px solid ${GY2}` }}>
                        <td style={{ padding: "8px 10px", fontWeight: 500 }}>{a.name}</td>
                        <td style={{ padding: "8px 10px", color: GY5 }}>{a.role}</td>
                        <td style={{ padding: "8px 10px", textAlign: "center" }}>{a.responsible ? <span style={{ color: AMB, fontWeight: 800, fontSize: 14 }}>✓</span> : ""}</td>
                        <td style={{ padding: "8px 10px", borderLeft: `2px solid ${GY2}` }}></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </PS>
            )}

            {/* Responsible person sign-off boxes */}
            {!noAtt && resp.length > 0 && (
              <PS title="Responsible person sign-off">
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(190px, 1fr))", gap: 12 }}>
                  {[...resp, { name: fd.preparedBy || "—", role: "Document prepared by" }].map((a, i) => (
                    <div key={i} style={{ border: `1px solid ${GY2}`, borderRadius: 7, padding: "12px 14px" }}>
                      <p style={{ margin: 0, fontSize: 12, fontWeight: 700, color: GY8 }}>{a.name}</p>
                      <p style={{ margin: "2px 0 12px", fontSize: 11, color: GY5 }}>{a.role}</p>
                      <div style={{ borderBottom: `1.5px solid ${GY7}`, height: 32, marginBottom: 4 }}></div>
                      <p style={{ margin: 0, fontSize: 10, color: "#9CA3AF" }}>Signature · Date ___________</p>
                    </div>
                  ))}
                </div>
              </PS>
            )}

            <div style={{ borderTop: `1px solid ${GY2}`, paddingTop: 10, display: "flex", justifyContent: "space-between", fontSize: 10, color: "#9CA3AF" }}>
              <span>Generated by {fd.preparedBy || "—"} · {fd.date}</span>
              <span>Site Document Generator · {fd.site || "—"}</span>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════
// REPORTS
// ═══════════════════════════════════════════════════════

function Reports({ docs, sites, onBack, onView }) {
  const [fSite, setFSite] = useState("all");
  const [fCat, setFCat] = useState("all");

  const filtered = docs
    .filter(d => (fSite === "all" || d.siteId === fSite) && (fCat === "all" || d.category === fCat))
    .sort((a, b) => new Date(b.generatedAt) - new Date(a.generatedAt));

  return (
    <div style={{ fontFamily: "system-ui, -apple-system, sans-serif", background: BG, minHeight: "100vh" }}>
      <div style={{ background: HDR, padding: "0 20px", height: 52, display: "flex", alignItems: "center", gap: 12, position: "sticky", top: 0, zIndex: 10 }}>
        <button onClick={onBack} style={{ background: "none", border: "none", color: "rgba(255,255,255,0.45)", cursor: "pointer", fontSize: 18, padding: 0 }}>←</button>
        <span style={{ color: "#fff", fontSize: 15, fontWeight: 700, flex: 1 }}>Document reports</span>
        <span style={{ padding: "2px 10px", background: "rgba(245,166,35,0.2)", borderRadius: 12, fontSize: 11, color: AMB, fontWeight: 700 }}>{filtered.length} total</span>
      </div>
      <div style={{ maxWidth: 900, margin: "0 auto", padding: "18px 16px" }}>
        <div style={{ display: "flex", gap: 9, marginBottom: 18, flexWrap: "wrap" }}>
          <select value={fSite} onChange={e => setFSite(e.target.value)} style={{ ...inp, width: "auto", minWidth: 160 }}>
            <option value="all">All sites</option>
            {sites.map(s => <option key={s.id} value={s.id}>{s.name}</option>)}
          </select>
          <select value={fCat} onChange={e => setFCat(e.target.value)} style={{ ...inp, width: "auto", minWidth: 160 }}>
            <option value="all">All categories</option>
            {CATEGORIES.map(c => <option key={c.id} value={c.id}>{c.label}</option>)}
          </select>
        </div>

        {filtered.length === 0 ? (
          <div style={{ textAlign: "center", padding: "60px 20px", color: GY5 }}>
            <p style={{ fontSize: 36, marginBottom: 12 }}>📋</p>
            <p style={{ fontSize: 16, fontWeight: 600, color: GY7 }}>No documents yet</p>
            <p style={{ fontSize: 13, marginTop: 6 }}>Generate your first document from the dashboard — it will appear here, saved with the team and signatories.</p>
          </div>
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: 7 }}>
            {filtered.map(doc => {
              const cc = CAT_COLOR[doc.category];
              const genAt = new Date(doc.generatedAt).toLocaleString("en-GB", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" });
              return (
                <div key={doc.id} onClick={() => onView(doc)}
                  style={{ background: "#fff", border: `1px solid ${GY2}`, borderLeft: `3px solid ${cc}`, borderRadius: 0, padding: "11px 15px", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, transition: "background 0.1s" }}
                  onMouseEnter={e => e.currentTarget.style.background = GY1}
                  onMouseLeave={e => e.currentTarget.style.background = "#fff"}>
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                      <p style={{ margin: 0, fontSize: 13, fontWeight: 700, color: GY8 }}>{doc.docLabel}</p>
                      <span style={{ padding: "1px 7px", borderRadius: 10, fontSize: 10, fontWeight: 700, background: cc + "22", color: cc }}>
                        {CATEGORIES.find(c => c.id === doc.category)?.label}
                      </span>
                    </div>
                    <p style={{ margin: "3px 0 0", fontSize: 12, color: GY5 }}>{doc.site} · {genAt} · by {doc.generatedBy || "—"}</p>
                    {doc.team?.length > 0 && (
                      <p style={{ margin: "2px 0 0", fontSize: 11, color: GY5 }}>
                        Team: {doc.team.slice(0, 3).join(", ")}{doc.team.length > 3 ? ` +${doc.team.length - 3} more` : ""}
                      </p>
                    )}
                    {doc.responsible?.length > 0 && (
                      <p style={{ margin: "2px 0 0", fontSize: 11, color: AMB, fontWeight: 600 }}>
                        Responsible: {doc.responsible.join(", ")}
                      </p>
                    )}
                  </div>
                  <span style={{ fontSize: 12, color: GY5, flexShrink: 0 }}>View →</span>
                </div>
              );
            })}
          </div>
        )}
      </div>
    </div>
  );
}

})();
// ─── Asset Register — embedded (localStorage, zero conflicts) ─────────────────
function AssetRegisterView(){
  const htmlDoc=`<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>SiteKit — Asset Register</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@500;600;700&family=Barlow:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>
  :root{
    --steel:#1f2933; --steel-2:#323f4b; --steel-3:#3e4c59;
    --paper:#f3f2ee; --card:#ffffff;
    --ink:#1f2933; --ink-soft:#52606d; --ink-faint:#9aa5b1; --line:#e2e0da;
    --hivis:#ffb000; --hivis-ink:#3a2a00;
    --ok:#2e9e5b; --ok-bg:#e7f4ec;
    --due:#ef7d18; --due-bg:#fdeede;
    --over:#d64545; --over-bg:#fbe6e6;
    --none:#7b8794; --none-bg:#eceef0;
    --r:14px;
    --shadow:0 1px 2px rgba(31,41,51,.06), 0 6px 18px rgba(31,41,51,.06);
    --sheet-shadow:0 -8px 40px rgba(31,41,51,.22);
    --safe-b:env(safe-area-inset-bottom,0px);
  }
  *{box-sizing:border-box}
  html,body{margin:0;padding:0}
  body{font-family:"Barlow",system-ui,-apple-system,Segoe UI,Roboto,sans-serif;background:var(--paper);
    color:var(--ink);-webkit-font-smoothing:antialiased;overscroll-behavior-y:none}
  button{font-family:inherit;cursor:pointer;border:none;background:none;color:inherit}
  input,select,textarea{font-family:inherit;font-size:16px}
  .mono{font-family:"IBM Plex Mono",ui-monospace,SFMono-Regular,Menlo,monospace}

  .app{max-width:520px;margin:0 auto;min-height:100vh;min-height:100dvh;position:relative;background:var(--paper)}
  .topbar{position:sticky;top:0;z-index:20;background:var(--steel);color:#fff;
    padding:calc(env(safe-area-inset-top,0px) + 14px) 18px 14px;display:flex;align-items:center;justify-content:space-between;gap:12px}
  .brand{display:flex;align-items:center;gap:10px;min-width:0}
  .brand .mark{width:30px;height:30px;border-radius:7px;flex:none;background:var(--hivis);color:var(--hivis-ink);
    display:grid;place-items:center;font-weight:700;box-shadow:inset 0 0 0 2px rgba(0,0,0,.12)}
  .brand .mark svg{width:18px;height:18px}
  .brand h1{font-family:"Barlow Condensed";font-weight:700;font-size:21px;letter-spacing:.06em;text-transform:uppercase;margin:0;line-height:1}
  .brand small{display:block;font-family:"IBM Plex Mono";font-size:9.5px;letter-spacing:.14em;text-transform:uppercase;color:#9aa5b1;margin-top:3px}
  .icon-btn{width:40px;height:40px;border-radius:10px;flex:none;background:var(--steel-2);color:#cdd5dd;display:grid;place-items:center;transition:background .15s,color .15s}
  .icon-btn:hover{background:var(--steel-3);color:#fff}
  .icon-btn svg{width:20px;height:20px}

  .scroll{padding:18px 16px 130px;animation:fade .25s ease}
  @keyframes fade{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:none}}
  .eyebrow{font-family:"Barlow Condensed";font-weight:600;font-size:12px;letter-spacing:.16em;text-transform:uppercase;color:var(--ink-faint);margin:0 2px 10px}
  .h-row{display:flex;align-items:baseline;justify-content:space-between;margin:0 2px 12px}
  .h-row h2{font-family:"Barlow Condensed";font-weight:700;font-size:24px;letter-spacing:.02em;margin:0;text-transform:uppercase}
  .muted{color:var(--ink-soft);font-size:14px;line-height:1.5;padding:4px 2px}

  .tiles{display:grid;grid-template-columns:1fr 1fr;gap:11px;margin-bottom:22px}
  .tile{background:var(--card);border-radius:var(--r);padding:15px 15px 13px;box-shadow:var(--shadow);border:1px solid var(--line);position:relative;overflow:hidden}
  .tile .n{font-family:"Barlow Condensed";font-weight:700;font-size:38px;line-height:.95}
  .tile .l{font-size:12.5px;color:var(--ink-soft);margin-top:5px;font-weight:500}
  .tile .bar{position:absolute;left:0;top:0;bottom:0;width:4px;background:var(--steel-3)}
  .tile.acc .bar{background:var(--hivis)}
  .tile.due .bar{background:var(--due)} .tile.due .n{color:var(--due)}
  .tile.over .bar{background:var(--over)} .tile.over .n{color:var(--over)}

  /* inspection tag */
  .tag{display:inline-flex;align-items:center;gap:7px;padding:4px 9px 4px 6px;border-radius:5px;
    font-family:"IBM Plex Mono";font-size:11px;font-weight:500;line-height:1;white-space:nowrap;border:1px solid transparent}
  .tag .hole{width:7px;height:7px;border-radius:50%;flex:none;background:var(--paper);box-shadow:inset 0 0 0 1.5px rgba(0,0,0,.18)}
  .tag .lab{font-weight:600;letter-spacing:.04em}
  .tag.ok{background:var(--ok-bg);color:#1c6b3c;border-color:#bfe3cd} .tag.ok .hole{box-shadow:inset 0 0 0 1.5px var(--ok)}
  .tag.due{background:var(--due-bg);color:#9a4e08;border-color:#f6cda3} .tag.due .hole{box-shadow:inset 0 0 0 1.5px var(--due)}
  .tag.over{background:var(--over-bg);color:#a32525;border-color:#f0bcbc} .tag.over .hole{box-shadow:inset 0 0 0 1.5px var(--over)}
  .tag.none{background:var(--none-bg);color:#52606d;border-color:#d5dadf}

  .alert{display:flex;align-items:center;gap:12px;background:var(--card);border:1px solid var(--line);border-radius:12px;padding:11px 13px;margin-bottom:9px;box-shadow:var(--shadow);transition:transform .12s}
  .alert:active{transform:scale(.99)}
  .alert .thumb{width:42px;height:42px;border-radius:9px;flex:none;background:var(--paper);background-size:cover;background-position:center;display:grid;place-items:center;color:var(--ink-faint)}
  .alert .thumb svg{width:20px;height:20px}
  .alert .mid{min-width:0;flex:1}
  .alert .nm{font-weight:600;font-size:15px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
  .alert .sub{font-size:12.5px;color:var(--ink-soft);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}

  .search{display:flex;align-items:center;gap:9px;background:var(--card);border:1px solid var(--line);border-radius:11px;padding:10px 13px;margin-bottom:13px;box-shadow:var(--shadow)}
  .search svg{width:18px;height:18px;color:var(--ink-faint);flex:none}
  .search input{border:none;outline:none;background:none;width:100%;color:var(--ink)}
  .chips{display:flex;gap:8px;overflow-x:auto;padding:2px 2px 12px;-webkit-overflow-scrolling:touch;scrollbar-width:none}
  .chips::-webkit-scrollbar{display:none}
  .chip{flex:none;padding:7px 13px;border-radius:999px;background:var(--card);border:1px solid var(--line);font-family:"Barlow Condensed";font-weight:600;font-size:13.5px;letter-spacing:.04em;text-transform:uppercase;color:var(--ink-soft);transition:.15s}
  .chip.on{background:var(--steel);color:#fff;border-color:var(--steel)}

  .card{display:flex;gap:13px;background:var(--card);border:1px solid var(--line);border-radius:var(--r);padding:12px;margin-bottom:11px;box-shadow:var(--shadow);transition:transform .12s;position:relative}
  .card:active{transform:scale(.992)}
  .card.oos{border-color:#f0bcbc}
  .card.oos::before{content:"";position:absolute;left:0;top:0;bottom:0;width:4px;border-radius:var(--r) 0 0 var(--r);background:var(--over)}
  .card .photo{width:74px;height:74px;border-radius:11px;flex:none;background:var(--paper);background-size:cover;background-position:center;display:grid;place-items:center;color:var(--ink-faint);border:1px solid var(--line);position:relative;overflow:hidden}
  .card .photo svg{width:26px;height:26px}
  .card .photo .oosdot{position:absolute;inset:auto 0 0 0;background:rgba(214,69,69,.92);color:#fff;font-family:"Barlow Condensed";font-weight:700;font-size:9px;letter-spacing:.08em;text-align:center;padding:2px 0;text-transform:uppercase}
  .card .body{min-width:0;flex:1;display:flex;flex-direction:column}
  .card .name{font-weight:600;font-size:16.5px;line-height:1.15;overflow:hidden;text-overflow:ellipsis;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical}
  .card .cat{font-family:"IBM Plex Mono";font-size:10.5px;letter-spacing:.06em;text-transform:uppercase;color:var(--ink-faint);margin-top:3px}
  .card .tagrow{display:flex;gap:6px;flex-wrap:wrap;margin-top:9px}
  .card .meta{display:flex;align-items:center;gap:6px;margin-top:auto;padding-top:9px;font-size:12px;color:var(--ink-soft)}
  .card .meta svg{width:13px;height:13px;flex:none;color:var(--ink-faint)}

  .site{background:var(--card);border:1px solid var(--line);border-radius:var(--r);padding:14px 15px;margin-bottom:11px;box-shadow:var(--shadow);transition:transform .12s}
  .site:active{transform:scale(.992)}
  .site .st-top{display:flex;align-items:center;justify-content:space-between;gap:10px}
  .site .st-name{display:flex;align-items:center;gap:9px;min-width:0}
  .site .pin{width:34px;height:34px;border-radius:9px;flex:none;background:var(--paper);display:grid;place-items:center;color:var(--steel);border:1px solid var(--line)}
  .site .pin svg{width:18px;height:18px}
  .site .st-name b{font-family:"Barlow Condensed";font-weight:700;font-size:19px;letter-spacing:.02em;text-transform:uppercase;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
  .site .count{font-family:"IBM Plex Mono";font-size:12px;color:var(--ink-soft);flex:none}
  .site .breakdown{display:flex;gap:7px;margin-top:12px;flex-wrap:wrap}
  .pillstat{display:inline-flex;align-items:center;gap:6px;font-family:"IBM Plex Mono";font-size:11px;padding:4px 8px;border-radius:6px;background:var(--paper)}
  .pillstat i{width:8px;height:8px;border-radius:50%;display:inline-block}

  .empty{text-align:center;padding:46px 24px;color:var(--ink-soft)}
  .empty .ic{width:62px;height:62px;border-radius:16px;background:var(--card);border:1px solid var(--line);display:grid;place-items:center;margin:0 auto 16px;color:var(--ink-faint);box-shadow:var(--shadow)}
  .empty .ic svg{width:30px;height:30px}
  .empty h3{font-family:"Barlow Condensed";font-weight:700;font-size:21px;letter-spacing:.02em;text-transform:uppercase;margin:0 0 6px;color:var(--ink)}
  .empty p{margin:0 auto 18px;max-width:300px;font-size:14.5px;line-height:1.5}

  .btn{display:inline-flex;align-items:center;justify-content:center;gap:8px;font-family:"Barlow Condensed";font-weight:700;font-size:15px;letter-spacing:.05em;text-transform:uppercase;padding:12px 20px;border-radius:11px;transition:transform .1s,filter .15s}
  .btn:active{transform:scale(.97)}
  .btn svg{width:17px;height:17px}
  .btn-pri{background:var(--hivis);color:var(--hivis-ink);box-shadow:0 3px 0 #c98a00}
  .btn-pri:active{box-shadow:0 1px 0 #c98a00;transform:translateY(2px)}
  .btn-ghost{background:var(--card);border:1px solid var(--line);color:var(--ink)}
  .btn-block{width:100%}

  .navwrap{position:fixed;left:0;right:0;bottom:0;z-index:30;pointer-events:none}
  .navwrap .inner{max-width:520px;margin:0 auto;position:relative}
  .nav{pointer-events:auto;background:var(--steel);padding:8px 12px calc(8px + var(--safe-b));display:grid;grid-template-columns:1fr 1fr 1.4fr 1fr 1fr;align-items:center;box-shadow:0 -2px 20px rgba(31,41,51,.18)}
  .nav button{display:flex;flex-direction:column;align-items:center;gap:3px;padding:6px 0;color:#8895a2;font-family:"Barlow Condensed";font-weight:600;font-size:11px;letter-spacing:.07em;text-transform:uppercase;transition:color .15s}
  .nav button svg{width:22px;height:22px}
  .nav button.on{color:#fff}
  .fab{pointer-events:auto;position:absolute;left:50%;transform:translateX(-50%);bottom:calc(20px + var(--safe-b));width:60px;height:60px;border-radius:50%;background:var(--hivis);color:var(--hivis-ink);display:grid;place-items:center;box-shadow:0 6px 18px rgba(255,176,0,.45),0 0 0 5px var(--paper);transition:transform .12s}
  .fab:active{transform:translateX(-50%) scale(.93)}
  .fab svg{width:30px;height:30px}

  /* sheet + modal share base */
  .scrim{position:fixed;inset:0;background:rgba(31,41,51,.5);z-index:40;opacity:0;transition:opacity .2s;pointer-events:none;backdrop-filter:blur(2px)}
  .scrim.show{opacity:1;pointer-events:auto}
  .sheet{position:fixed;left:0;right:0;bottom:0;z-index:50;max-width:520px;margin:0 auto;background:var(--paper);border-radius:22px 22px 0 0;box-shadow:var(--sheet-shadow);max-height:94vh;max-height:94dvh;display:flex;flex-direction:column;transform:translateY(100%);transition:transform .26s cubic-bezier(.32,.72,0,1)}
  .sheet.show{transform:none}
  .modal-scrim{z-index:60}
  .modal{z-index:70}
  .grab{width:38px;height:4px;border-radius:99px;background:#cfd4d9;margin:9px auto 2px}
  .sh-head{display:flex;align-items:center;justify-content:space-between;padding:6px 16px 12px;border-bottom:1px solid var(--line)}
  .sh-head h3{font-family:"Barlow Condensed";font-weight:700;font-size:22px;letter-spacing:.03em;text-transform:uppercase;margin:0}
  .sh-x{width:34px;height:34px;border-radius:9px;background:var(--card);border:1px solid var(--line);display:grid;place-items:center;color:var(--ink-soft)}
  .sh-x svg{width:18px;height:18px}
  .sh-body{overflow-y:auto;padding:16px 16px 14px;-webkit-overflow-scrolling:touch}
  .sh-foot{padding:13px 16px calc(14px + var(--safe-b));border-top:1px solid var(--line);display:flex;gap:11px;background:var(--paper)}
  .sh-foot .btn{flex:1}

  .photozone{position:relative;width:100%;aspect-ratio:16/10;border-radius:14px;overflow:hidden;background:var(--card);border:1.5px dashed var(--line);display:grid;place-items:center;margin-bottom:18px;background-size:cover;background-position:center}
  .photozone .ph-empty{display:flex;flex-direction:column;align-items:center;gap:8px;color:var(--ink-faint)}
  .photozone .ph-empty svg{width:30px;height:30px}
  .photozone .ph-empty span{font-family:"Barlow Condensed";font-weight:600;font-size:14px;letter-spacing:.06em;text-transform:uppercase}
  .photozone.has{border-style:solid}
  .ph-actions{position:absolute;right:10px;bottom:10px;display:flex;gap:8px}
  .ph-actions button{width:38px;height:38px;border-radius:10px;background:rgba(31,41,51,.82);color:#fff;display:grid;place-items:center;backdrop-filter:blur(4px)}
  .ph-actions button svg{width:18px;height:18px}

  .field{margin-bottom:15px}
  .field label{display:block;font-family:"Barlow Condensed";font-weight:600;font-size:13px;letter-spacing:.08em;text-transform:uppercase;color:var(--ink-soft);margin-bottom:6px}
  .field .req{color:var(--over)}
  .inp{width:100%;padding:12px 13px;border-radius:11px;border:1px solid var(--line);background:var(--card);color:var(--ink);outline:none;transition:border-color .15s,box-shadow .15s}
  .inp:focus{border-color:var(--hivis);box-shadow:0 0 0 3px rgba(255,176,0,.18)}
  textarea.inp{resize:vertical;min-height:64px}
  .two{display:grid;grid-template-columns:1fr 1fr;gap:11px}

  .seg{display:flex;gap:8px}
  .seg button{flex:1;padding:11px 0;border-radius:11px;border:1px solid var(--line);background:var(--card);font-family:"Barlow Condensed";font-weight:700;font-size:15px;letter-spacing:.03em;color:var(--ink-soft);transition:.15s}
  .seg button.on{background:var(--steel);color:#fff;border-color:var(--steel)}
  .seg.use button.on.yes{background:var(--ok);border-color:var(--ok);color:#fff}
  .seg.use button.on.no{background:var(--over);border-color:var(--over);color:#fff}
  .seg.res button.on.pass{background:var(--ok);border-color:var(--ok);color:#fff}
  .seg.res button.on.fail{background:var(--over);border-color:var(--over);color:#fff}

  .nextline{font-family:"IBM Plex Mono";font-size:12.5px;margin-top:9px;padding:9px 11px;border-radius:9px;background:var(--card);border:1px solid var(--line);display:flex;align-items:center;gap:8px;color:var(--ink-soft)}
  .nextline b{color:var(--ink)}
  .danger{color:var(--over);font-family:"Barlow Condensed";font-weight:700;font-size:14px;letter-spacing:.05em;text-transform:uppercase;display:flex;align-items:center;justify-content:center;gap:7px;padding:11px;margin-top:4px;border-radius:11px;border:1px solid #f0c9c9;background:#fdf2f2}
  .danger svg{width:16px;height:16px}

  /* detail */
  .det-photo{position:relative;width:100%;aspect-ratio:16/10;border-radius:14px;background:var(--card);border:1px solid var(--line);background-size:cover;background-position:center;margin-bottom:14px;display:grid;place-items:center;color:var(--ink-faint);overflow:hidden}
  .det-photo svg{width:40px;height:40px}
  .oos-ribbon{position:absolute;inset:auto 0 0 0;background:rgba(214,69,69,.94);color:#fff;font-family:"Barlow Condensed";font-weight:700;font-size:13px;letter-spacing:.12em;text-transform:uppercase;text-align:center;padding:7px 0}
  .det-head{display:flex;align-items:flex-start;justify-content:space-between;gap:10px;margin-bottom:10px}
  .det-head .dh-name{font-family:"Barlow Condensed";font-weight:700;font-size:24px;letter-spacing:.02em;line-height:1.05}
  .det-head .dh-cat{font-family:"IBM Plex Mono";font-size:11px;letter-spacing:.06em;color:var(--ink-faint);margin-top:4px;text-transform:uppercase}
  .det-chips{display:flex;flex-direction:column;gap:6px;align-items:flex-end;flex:none}
  .qa{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin:14px 0 6px}
  .qa button{display:flex;flex-direction:column;align-items:center;gap:6px;padding:11px 4px;border-radius:12px;background:var(--card);border:1px solid var(--line);transition:transform .1s}
  .qa button:active{transform:scale(.95)}
  .qa button .qi{width:30px;height:30px;border-radius:9px;display:grid;place-items:center;background:var(--paper);color:var(--steel)}
  .qa button .qi svg{width:17px;height:17px}
  .qa button span{font-family:"Barlow Condensed";font-weight:600;font-size:11px;letter-spacing:.04em;text-transform:uppercase;color:var(--ink-soft);text-align:center;line-height:1.05}
  .qa button.accent .qi{background:var(--hivis);color:var(--hivis-ink)}

  .sec{display:flex;align-items:center;justify-content:space-between;margin:22px 2px 11px}
  .sec span{font-family:"Barlow Condensed";font-weight:700;font-size:16px;letter-spacing:.06em;text-transform:uppercase}
  .addlink{font-family:"Barlow Condensed";font-weight:700;font-size:13px;letter-spacing:.05em;text-transform:uppercase;color:#9a6a00;display:inline-flex;align-items:center;gap:4px;padding:6px 10px;border-radius:9px;background:var(--due-bg);border:1px solid #f6cda3}
  .addlink svg{width:14px;height:14px}

  .det-row{display:flex;justify-content:space-between;gap:14px;padding:12px 2px;border-bottom:1px solid var(--line)}
  .det-row .k{font-family:"Barlow Condensed";font-weight:600;font-size:13px;letter-spacing:.07em;text-transform:uppercase;color:var(--ink-faint);flex:none}
  .det-row .v{text-align:right;font-size:15px}
  .det-row .v.mono{font-size:13.5px}

  .patrec{display:flex;gap:11px;align-items:center;background:var(--card);border:1px solid var(--line);border-radius:12px;padding:12px;margin-bottom:9px;box-shadow:var(--shadow)}
  .patrec .pr-main{flex:1;min-width:0}
  .patrec .pr-date{font-family:"IBM Plex Mono";font-weight:500;font-size:14px}
  .patrec .pr-meta{font-size:12.5px;color:var(--ink-soft);margin-top:3px}
  .patrec .pr-meta b{color:var(--ink)}
  .patrec .pr-tester{font-size:12px;color:var(--ink-faint);margin-top:2px}
  .patrec .pr-right{display:flex;flex-direction:column;gap:7px;align-items:flex-end;flex:none}
  .cert-btn{display:inline-flex;align-items:center;gap:5px;font-family:"Barlow Condensed";font-weight:700;font-size:11.5px;letter-spacing:.04em;text-transform:uppercase;color:var(--ink);background:var(--paper);border:1px solid var(--line);padding:6px 9px;border-radius:8px}
  .cert-btn svg{width:13px;height:13px}

  .tl{position:relative;margin-left:6px;padding-left:20px}
  .tl::before{content:"";position:absolute;left:5px;top:6px;bottom:6px;width:2px;background:var(--line)}
  .tl-item{position:relative;padding:0 0 16px}
  .tl-item .dot{position:absolute;left:-19px;top:2px;width:14px;height:14px;border-radius:50%;background:var(--card);border:2px solid var(--steel-3);display:grid;place-items:center}
  .tl-item .dot.move{border-color:var(--steel)}
  .tl-item .dot.cond{border-color:var(--due)}
  .tl-item .dot.pat{border-color:var(--ok)}
  .tl-item .dot.oos{border-color:var(--over)}
  .tl-item .ti-top{display:flex;align-items:baseline;justify-content:space-between;gap:10px}
  .tl-item .ti-title{font-weight:600;font-size:14.5px}
  .tl-item .ti-date{font-family:"IBM Plex Mono";font-size:10.5px;color:var(--ink-faint);flex:none}
  .tl-item .ti-sub{font-size:13px;color:var(--ink-soft);margin-top:2px;line-height:1.4}
  .tl-item .ti-photo{margin-top:8px;width:100%;max-width:200px;aspect-ratio:4/3;border-radius:10px;background-size:cover;background-position:center;border:1px solid var(--line)}
  .move-arrow{display:inline-flex;align-items:center;gap:6px;font-family:"IBM Plex Mono";font-size:12.5px;margin-top:3px}
  .move-arrow svg{width:13px;height:13px;color:var(--ink-faint)}

  /* mini photo control (modals) */
  .mphoto{display:flex;gap:11px;align-items:center;margin-bottom:4px}
  .mphoto .prev{width:74px;height:74px;border-radius:11px;background:var(--card);border:1px solid var(--line);background-size:cover;background-position:center;display:grid;place-items:center;color:var(--ink-faint);flex:none}
  .mphoto .prev svg{width:24px;height:24px}
  .mphoto .addbtn{flex:1;padding:13px;border-radius:11px;border:1.5px dashed var(--line);background:var(--card);font-family:"Barlow Condensed";font-weight:600;font-size:14px;letter-spacing:.05em;text-transform:uppercase;color:var(--ink-soft);display:flex;align-items:center;justify-content:center;gap:8px}
  .mphoto .addbtn svg{width:18px;height:18px}
  .mphoto .rm{font-family:"Barlow Condensed";font-weight:600;font-size:12px;letter-spacing:.05em;text-transform:uppercase;color:var(--over);background:none;padding:6px}

  .toast{position:fixed;left:50%;transform:translateX(-50%) translateY(20px);bottom:calc(96px + var(--safe-b));z-index:95;background:var(--steel);color:#fff;padding:11px 18px;border-radius:11px;font-weight:500;font-size:14px;box-shadow:var(--shadow);opacity:0;transition:.25s;pointer-events:none;display:flex;align-items:center;gap:9px;max-width:90%}
  .toast.show{opacity:1;transform:translateX(-50%) translateY(0)}
  .toast svg{width:18px;height:18px;color:var(--hivis);flex:none}

  .export-card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:14px;margin-bottom:14px}
  .export-card h4{font-family:"Barlow Condensed";font-weight:700;font-size:16px;letter-spacing:.04em;text-transform:uppercase;margin:0 0 4px}
  .export-card p{margin:0 0 11px;font-size:13.5px;color:var(--ink-soft);line-height:1.45}
  .export-card .row{display:flex;gap:9px}
  .export-card .btn{flex:1;font-size:13.5px;padding:11px 10px}
  textarea.exp{width:100%;height:140px;border:1px solid var(--line);border-radius:11px;padding:11px;font-family:"IBM Plex Mono";font-size:11px;line-height:1.4;background:var(--paper);color:var(--ink-soft);resize:none}

  @media (prefers-reduced-motion: reduce){ *{animation:none!important;transition:none!important} }
</style>
</head>
<body>
<div class="app">
  <header class="topbar">
    <div class="brand">
      <div class="mark" aria-hidden="true"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 18h18"/><path d="M5 18V11a7 7 0 0 1 14 0v7"/><path d="M12 4v2"/></svg></div>
      <div><h1>SiteKit</h1><small>Asset register</small></div>
    </div>
    <button class="icon-btn" id="exportBtn" aria-label="Share register with head office" title="Share with head office">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 12v6a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-6"/><path d="M16 6l-4-4-4 4"/><path d="M12 2v14"/></svg>
    </button>
  </header>
  <main id="view" class="scroll"></main>
</div>

<div class="navwrap"><div class="inner">
  <nav class="nav" id="nav">
    <button data-view="overview" class="on">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="7" height="9" rx="1"/><rect x="14" y="3" width="7" height="5" rx="1"/><rect x="14" y="12" width="7" height="9" rx="1"/><rect x="3" y="16" width="7" height="5" rx="1"/></svg>Overview
    </button>
    <button data-view="assets">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z"/><path d="m3.27 6.96 8.73 5.04 8.73-5.04"/><path d="M12 22.08V12"/></svg>Assets
    </button>
    <span></span>
    <button data-view="sites">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 10c0 6-8 12-8 12s-8-6-8-12a8 8 0 0 1 16 0z"/><circle cx="12" cy="10" r="3"/></svg>Sites
    </button>
    <button data-view="assets" style="visibility:hidden">x</button>
  </nav>
  <button class="fab" id="fab" aria-label="Add asset"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round"><path d="M12 5v14M5 12h14"/></svg></button>
</div></div>

<!-- primary sheet (detail / create / edit / export) -->
<div class="scrim" id="scrim"></div>
<section class="sheet" id="sheet" aria-hidden="true">
  <div class="grab"></div>
  <div class="sh-head"><h3 id="sheetTitle">Add asset</h3>
    <button class="sh-x" id="sheetClose" aria-label="Close"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><path d="M18 6 6 18M6 6l12 12"/></svg></button></div>
  <div class="sh-body" id="sheetBody"></div>
</section>

<!-- secondary modal (move / update / pat) layered above the detail sheet -->
<div class="scrim modal-scrim" id="modalScrim"></div>
<section class="sheet modal" id="modal" aria-hidden="true">
  <div class="grab"></div>
  <div class="sh-head"><h3 id="modalTitle"></h3>
    <button class="sh-x" id="modalClose" aria-label="Close"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><path d="M18 6 6 18M6 6l12 12"/></svg></button></div>
  <div class="sh-body" id="modalBody"></div>
</section>

<div class="toast" id="toast"><span id="toastMsg"></span></div>
<input type="file" id="fileInput" accept="image/*" capture="environment" style="display:none">

<script>
"use strict";
/* ============================================================
   SiteKit v2 — asset register with PAT records, moves,
   condition updates, serviceability and certificate output.
   ============================================================ */

/* ---------- storage ---------- */
// Storage: localStorage with in-memory fallback
const mem = new Map();
const Store = {
  async list(){
    try{
      const keys=Object.keys(localStorage).filter(k=>k.startsWith("bm_asset:"));
      return keys.map(k=>k.replace(/^bm_/,""));
    }catch(e){ return [...mem.keys()].filter(k=>k.startsWith("asset:")); }
  },
  async get(k){
    try{ const v=localStorage.getItem("bm_"+k); return v||null; }
    catch(e){ return mem.get(k)||null; }
  },
  async set(k,v){
    try{ localStorage.setItem("bm_"+k,v); }
    catch(e){ mem.set(k,v); }
  },
  async del(k){
    try{ localStorage.removeItem("bm_"+k); }
    catch(e){ mem.delete(k); }
  }
};

/* ---------- state ---------- */
const state = { assets:[], view:"overview", filterSite:"__all", search:"" };
let editingId=null, editFromDetail=false;
let draftPhoto=null;
let photoSink=null;           // callback for the shared file input
const CATS=["Power tools","Plant & machinery","Access equipment","Site lighting","Leads & cables","Welfare unit","Survey equipment","Hand tools","Other"];

/* ---------- helpers ---------- */
const $=s=>document.querySelector(s);
const el=(t,c,h)=>{const e=document.createElement(t);if(c)e.className=c;if(h!=null)e.innerHTML=h;return e;};
const esc=s=>(s==null?"":String(s)).replace(/[&<>"]/g,m=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[m]));
const uid=()=>"a"+Date.now().toString(36)+Math.random().toString(36).slice(2,7);
const MON=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
function parseD(s){ if(!s) return null; const [y,m,d]=s.split("-").map(Number); return new Date(y,m-1,d); }
function fmtD(s){ const d=(s instanceof Date)?s:parseD(s); if(!d) return "—"; return d.getDate()+" "+MON[d.getMonth()]+" "+d.getFullYear(); }
function fmtShort(d){ const x=(d instanceof Date)?d:parseD(d); return x?x.getDate()+" "+MON[x.getMonth()]:"—"; }
function isoLocal(d){ const m=String(d.getMonth()+1).padStart(2,"0"); const day=String(d.getDate()).padStart(2,"0"); return d.getFullYear()+"-"+m+"-"+day; }
function addMonths(d,n){ const x=new Date(d.getTime()); const day=x.getDate(); x.setMonth(x.getMonth()+n); if(x.getDate()<day) x.setDate(0); return x; }
function today0(){ const t=new Date(); return new Date(t.getFullYear(),t.getMonth(),t.getDate()); }
function tsDate(ts){ return fmtD(isoLocal(new Date(ts))); }
const boxIcon=\`<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z"/><path d="m3.27 6.96 8.73 5.04 8.73-5.04"/><path d="M12 22.08V12"/></svg>\`;
const pinIcon=\`<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 10c0 6-8 12-8 12s-8-6-8-12a8 8 0 0 1 16 0z"/><circle cx="12" cy="10" r="3"/></svg>\`;

/* ---------- PAT engine ---------- */
function latestPat(a){ if(!a.patTests||!a.patTests.length) return null; return [...a.patTests].sort((x,y)=>(y.date||"").localeCompare(x.date||""))[0]; }
function patExpiry(t,a){ if(!t||!t.date) return null; return addMonths(parseD(t.date), t.interval||(a&&a.interval)||3); }
function patStatus(a){
  const lp=latestPat(a);
  if(!lp) return {key:"none",due:null,days:null};
  const due=patExpiry(lp,a);
  if(lp.result==="fail") return {key:"fail",due,days:null};
  const days=Math.round((due-today0())/86400000);
  return {key: days<0?"over":(days<=30?"due":"ok"), due, days};
}
function tagHTML(a){
  const s=patStatus(a); let txt;
  if(s.key==="none") txt="NO PAT TEST";
  else if(s.key==="fail") txt="PAT FAILED";
  else if(s.key==="over") txt="OVERDUE "+fmtShort(s.due);
  else if(s.key==="due") txt="DUE "+fmtShort(s.due);
  else txt="PAT OK · "+fmtShort(s.due);
  const cls=(s.key==="fail")?"over":s.key;
  return \`<span class="tag \${cls}"><span class="hole"></span><span class="lab">\${txt}</span></span>\`;
}
const oosTag=\`<span class="tag over"><span class="hole"></span><span class="lab">OUT OF SERVICE</span></span>\`;

/* an asset needs attention if out of service or PAT over/due/failed */
function attention(a){
  if(a.usable===false) return {sev:0,label:"Out of service",reason:"oos"};
  const s=patStatus(a);
  if(s.key==="fail") return {sev:0,label:"PAT failed",reason:"fail"};
  if(s.key==="over") return {sev:0,label:\`PAT overdue · \${Math.abs(s.days)}d\`,reason:"over"};
  if(s.key==="due")  return {sev:1,label:\`PAT due \${fmtShort(s.due)} · \${s.days}d\`,reason:"due"};
  return null;
}

/* ---------- load + migrate ---------- */
async function loadAll(){
  const keys=await Store.list();
  const out=[];
  for(const raw of keys){
    const k=String(raw).startsWith("asset:")?String(raw):"asset:"+raw;
    const v=await Store.get(k);
    if(v){ try{ out.push(migrate(JSON.parse(v))); }catch(e){} }
  }
  out.sort((a,b)=>(a.name||"").localeCompare(b.name||""));
  state.assets=out;
}
function migrate(a){
  if(a.usable===undefined) a.usable=true;
  if(a.interval===undefined) a.interval=3;
  if(!Array.isArray(a.activity)) a.activity=[];
  if(!Array.isArray(a.patTests)){
    a.patTests=[];
    if(a.lastPat) a.patTests.push({id:uid(),date:a.lastPat,interval:a.interval||3,result:"pass",tester:"",notes:""});
  }
  return a;
}
async function saveAsset(a){ await Store.set("asset:"+a.id, JSON.stringify(a)); }
async function deleteAsset(id){ await Store.del("asset:"+id); }
function findAsset(id){ return state.assets.find(x=>x.id===id); }
function logActivity(a,entry){ entry.id=uid(); entry.ts=entry.ts||Date.now(); a.activity=a.activity||[]; a.activity.unshift(entry); }

/* ---------- sites + filter ---------- */
function sites(){ return [...new Set(state.assets.map(a=>a.site).filter(Boolean))].sort((a,b)=>a.localeCompare(b)); }
function filteredAssets(){
  const q=state.search.trim().toLowerCase();
  return state.assets.filter(a=>{
    if(state.filterSite!=="__all" && a.site!==state.filterSite) return false;
    if(q){ const hay=[a.name,a.category,a.site,a.tag,a.notes].join(" ").toLowerCase(); if(!hay.includes(q)) return false; }
    return true;
  });
}

/* ============================================================
   RENDER VIEWS
   ============================================================ */
function render(){
  document.querySelectorAll(".nav button[data-view]").forEach(b=>{ if(b.style.visibility==="hidden") return; b.classList.toggle("on", b.dataset.view===state.view); });
  const v=$("#view");
  v.style.animation="none"; void v.offsetWidth; v.style.animation="";
  if(state.view==="overview") v.innerHTML=renderOverview();
  else if(state.view==="assets") v.innerHTML=renderAssets();
  else if(state.view==="sites") v.innerHTML=renderSites();
  wire();
}

function renderOverview(){
  const total=state.assets.length;
  if(total===0) return emptyState();
  const nSites=sites().length;
  let patAtt=0, oos=0;
  state.assets.forEach(a=>{ const s=patStatus(a); if(s.key==="due"||s.key==="over"||s.key==="fail") patAtt++; if(a.usable===false) oos++; });

  const alerts=state.assets.map(a=>({a,att:attention(a)})).filter(x=>x.att).sort((x,y)=> x.att.sev-y.att.sev);

  let html=\`<p class="eyebrow">Site overview</p>
    <div class="tiles">
      <div class="tile acc"><span class="bar"></span><div class="n">\${total}</div><div class="l">Assets tracked</div></div>
      <div class="tile"><span class="bar"></span><div class="n">\${nSites}</div><div class="l">Active site\${nSites===1?"":"s"}</div></div>
      <div class="tile due"><span class="bar"></span><div class="n">\${patAtt}</div><div class="l">Need PAT attention</div></div>
      <div class="tile over"><span class="bar"></span><div class="n">\${oos}</div><div class="l">Out of service</div></div>
    </div>
    <div class="h-row"><h2>Action needed</h2></div>\`;

  if(!alerts.length){
    html+=\`<div class="alert" style="color:var(--ink-soft)">
      <div class="thumb" style="background:var(--ok-bg);color:var(--ok)"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6 9 17l-5-5"/></svg></div>
      <div class="mid"><div class="nm">All clear</div><div class="sub">No PAT due and nothing out of service.</div></div></div>\`;
  } else {
    alerts.forEach(({a,att})=>{
      const ph=a.photo?\`style="background-image:url('\${a.photo}')"\`:"";
      const t = att.reason==="oos" ? oosTag : tagHTML(a);
      html+=\`<div class="alert" data-open="\${a.id}">
        <div class="thumb" \${ph}>\${a.photo?"":boxIcon}</div>
        <div class="mid"><div class="nm">\${esc(a.name)}</div><div class="sub">\${esc(a.site||"No site")} · \${esc(att.label)}</div></div>
        \${t}</div>\`;
    });
  }
  return html;
}

function renderAssets(){
  if(state.assets.length===0) return emptyState();
  const list=filteredAssets(), ss=sites();
  let chips=\`<button class="chip \${state.filterSite==="__all"?"on":""}" data-site="__all">All sites</button>\`;
  ss.forEach(s=> chips+=\`<button class="chip \${state.filterSite===s?"on":""}" data-site="\${esc(s)}">\${esc(s)}</button>\`);

  let html=\`<div class="h-row"><h2>Assets</h2><span class="mono" style="font-size:12px;color:var(--ink-faint)">\${list.length} item\${list.length===1?"":"s"}</span></div>
    <div class="search"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="7"/><path d="m21 21-4.3-4.3"/></svg>
      <input id="searchInput" type="search" placeholder="Search name, tag, category…" value="\${esc(state.search)}"></div>
    <div class="chips">\${chips}</div>\`;

  if(!list.length){ html+=\`<div class="empty"><div class="ic">\${boxIcon}</div><h3>Nothing here</h3><p>No assets match your search or this site.</p></div>\`; return html; }

  list.forEach(a=>{
    const ph=a.photo?\`style="background-image:url('\${a.photo}')"\`:"";
    const oos=a.usable===false;
    html+=\`<div class="card \${oos?"oos":""}" data-open="\${a.id}">
      <div class="photo" \${ph}>\${a.photo?"":boxIcon}\${oos?'<div class="oosdot">Out of service</div>':""}</div>
      <div class="body">
        <div class="name">\${esc(a.name)}</div>
        <div class="cat">\${esc(a.category||"Uncategorised")}\${a.tag?" · "+esc(a.tag):""}</div>
        <div class="tagrow">\${tagHTML(a)}\${oos?oosTag:""}</div>
        <div class="meta">\${pinIcon}\${esc(a.site||"No site")}\${a.delivered?\`&nbsp;·&nbsp;<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="18" rx="2"/><path d="M16 2v4M8 2v4M3 10h18"/></svg> \${fmtShort(a.delivered)}\`:""}</div>
      </div></div>\`;
  });
  return html;
}

function renderSites(){
  if(state.assets.length===0) return emptyState();
  const ss=sites();
  let html=\`<div class="h-row"><h2>Sites</h2><span class="mono" style="font-size:12px;color:var(--ink-faint)">\${ss.length} active</span></div>\`;
  if(!ss.length){ html+=\`<div class="empty"><div class="ic">\${pinIcon}</div><h3>No sites yet</h3><p>Add a site when you create an asset and it will appear here.</p></div>\`; return html; }
  ss.forEach(site=>{
    const items=state.assets.filter(a=>a.site===site);
    let ok=0,due=0,over=0,oos=0;
    items.forEach(a=>{ if(a.usable===false){oos++;return;} const s=patStatus(a); if(s.key==="ok")ok++; else if(s.key==="due")due++; else over++; });
    const seg=(n,c,l)=> n>0?\`<span class="pillstat"><i style="background:\${c}"></i>\${n} \${l}</span>\`:"";
    html+=\`<div class="site" data-site-open="\${esc(site)}">
      <div class="st-top"><div class="st-name"><div class="pin">\${pinIcon}</div><b>\${esc(site)}</b></div>
        <span class="count">\${items.length} asset\${items.length===1?"":"s"}</span></div>
      <div class="breakdown">\${seg(oos,"var(--over)","out of service")}\${seg(over,"var(--over)","PAT issue")}\${seg(due,"var(--due)","PAT due")}\${seg(ok,"var(--ok)","ready")}</div>
    </div>\`;
  });
  return html;
}

function emptyState(){
  return \`<div class="empty"><div class="ic">\${boxIcon}</div><h3>No assets yet</h3>
    <p>Add the tools, plant and equipment on your sites. Then move them, log condition and track PAT tests over time.</p>
    <div style="display:flex;flex-direction:column;gap:10px;align-items:center">
      <button class="btn btn-pri" id="emptyAdd"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round"><path d="M12 5v14M5 12h14"/></svg> Add your first asset</button>
      <button class="btn btn-ghost" id="emptySample">Load sample assets</button>
    </div></div>\`;
}

/* ---------- wiring ---------- */
function wire(){
  document.querySelectorAll("[data-open]").forEach(n=> n.onclick=()=>openDetail(n.dataset.open));
  document.querySelectorAll(".chip[data-site]").forEach(n=> n.onclick=()=>{state.filterSite=n.dataset.site;render();});
  document.querySelectorAll("[data-site-open]").forEach(n=> n.onclick=()=>{state.filterSite=n.dataset.siteOpen;state.view="assets";render();});
  const si=$("#searchInput"); if(si) si.oninput=()=>{ state.search=si.value; renderAssetsInPlace(); };
  const ea=$("#emptyAdd"); if(ea) ea.onclick=()=>openSheet(null);
  const es=$("#emptySample"); if(es) es.onclick=loadSamples;
}
function renderAssetsInPlace(){
  const v=$("#view"); const val=$("#searchInput")?$("#searchInput").value:null;
  v.innerHTML=renderAssets(); wire();
  const si=$("#searchInput"); if(si&&val!=null){ si.focus(); si.setSelectionRange(val.length,val.length); }
}

$("#nav").addEventListener("click",e=>{ const b=e.target.closest("button[data-view]"); if(!b||b.style.visibility==="hidden") return; state.view=b.dataset.view; render(); });
$("#fab").onclick=()=>openSheet(null);

/* ============================================================
   DETAIL  (primary sheet)
   ============================================================ */
function openDetail(id){
  const a=findAsset(id); if(!a) return;
  const s=patStatus(a);
  const ph=a.photo?\`style="background-image:url('\${a.photo}')"\`:"";
  const oos=a.usable===false;
  const useChip = oos
    ? \`<span class="tag over"><span class="hole"></span><span class="lab">NOT USABLE</span></span>\`
    : \`<span class="tag ok"><span class="hole"></span><span class="lab">USABLE</span></span>\`;

  $("#sheetTitle").textContent="Asset details";
  $("#sheetBody").innerHTML=\`
    <div class="det-photo" \${ph}>\${a.photo?"":boxIcon}\${oos?'<div class="oos-ribbon">Out of service</div>':""}</div>
    <div class="det-head">
      <div><div class="dh-name">\${esc(a.name)}</div><div class="dh-cat">\${esc(a.category||"Uncategorised")}\${a.tag?" · "+esc(a.tag):""}</div></div>
      <div class="det-chips">\${tagHTML(a)}\${useChip}</div>
    </div>

    <div class="qa">
      <button id="detMove"><span class="qi">\${pinIcon}</span><span>Move</span></button>
      <button id="detUpdate" class="accent"><span class="qi"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14.5 4h-5L7 7H4a2 2 0 0 0-2 2v9a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2V9a2 2 0 0 0-2-2h-3l-2.5-3Z"/><circle cx="12" cy="13" r="3.5"/></svg></span><span>Update</span></button>
      <button id="detPat"><span class="qi"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M9 12l2 2 4-4"/><path d="M21 12c0 1.66-4 3-9 3s-9-1.34-9-3"/><path d="M3 5v14c0 1.66 4 3 9 3s9-1.34 9-3V5"/><path d="M3 5c0 1.66 4 3 9 3s9-1.34 9-3-4-3-9-3-9 1.34-9 3"/></svg></span><span>PAT test</span></button>
      <button id="detEdit"><span class="qi"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 20h9"/><path d="M16.5 3.5a2.12 2.12 0 0 1 3 3L7 19l-4 1 1-4Z"/></svg></span><span>Edit</span></button>
    </div>

    <div class="sec"><span>Details</span></div>
    <div class="det-row"><span class="k">Current site</span><span class="v">\${esc(a.site||"—")}</span></div>
    <div class="det-row"><span class="k">Asset tag</span><span class="v mono">\${esc(a.tag||"—")}</span></div>
    <div class="det-row"><span class="k">Delivered to site</span><span class="v mono">\${a.delivered?fmtD(a.delivered):"—"}</span></div>
    <div class="det-row"><span class="k">Next PAT due</span><span class="v mono">\${s.due?fmtD(s.due):"—"}</span></div>
    \${a.notes?\`<div class="det-row" style="flex-direction:column;align-items:flex-start"><span class="k" style="margin-bottom:6px">Notes</span><span class="v" style="text-align:left;font-size:14px;color:var(--ink-soft);line-height:1.5">\${esc(a.notes)}</span></div>\`:""}

    <div class="sec"><span>PAT tests</span><button class="addlink" id="patAdd"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round"><path d="M12 5v14M5 12h14"/></svg> Add test</button></div>
    <div id="patList">\${patListHTML(a)}</div>

    <div class="sec"><span>Activity</span></div>
    <div id="actList">\${activityHTML(a)}</div>
    <div style="height:8px"></div>\`;
  showSheet();

  $("#detMove").onclick=()=>openMove(id);
  $("#detUpdate").onclick=()=>openUpdate(id);
  $("#detPat").onclick=()=>openPat(id);
  $("#patAdd").onclick=()=>openPat(id);
  $("#detEdit").onclick=()=>{ editFromDetail=true; openSheet(id); };
  wireCertButtons(a);
}

function patListHTML(a){
  if(!a.patTests||!a.patTests.length) return \`<p class="muted">No PAT tests recorded. Add one to start tracking the expiry date.</p>\`;
  return [...a.patTests].sort((x,y)=>(y.date||"").localeCompare(x.date||"")).map(t=>{
    const exp=patExpiry(t,a);
    const pass=t.result!=="fail";
    return \`<div class="patrec">
      <div class="pr-main">
        <div class="pr-date">\${fmtD(t.date)}</div>
        <div class="pr-meta">Retest every \${t.interval||a.interval||3} mo · expires <b>\${fmtD(exp)}</b></div>
        \${t.tester?\`<div class="pr-tester">Tested by \${esc(t.tester)}</div>\`:""}
      </div>
      <div class="pr-right">
        <span class="tag \${pass?"ok":"over"}"><span class="hole"></span><span class="lab">\${pass?"PASS":"FAIL"}</span></span>
        <button class="cert-btn" data-cert="\${t.id}"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/><path d="M9 15l2 2 4-4"/></svg> Certificate</button>
      </div></div>\`;
  }).join("");
}
function wireCertButtons(a){ document.querySelectorAll("[data-cert]").forEach(b=> b.onclick=()=>generateCert(a.id,b.dataset.cert)); }

function activityHTML(a){
  if(!a.activity||!a.activity.length) return \`<p class="muted">No activity yet. Moves, condition photos and comments will appear here.</p>\`;
  return \`<div class="tl">\`+a.activity.map(e=>{
    let dotCls="", title="", sub="";
    if(e.type==="create"){ title="Added to register"; sub=e.to?\`Initial location: \${esc(e.to)}\`:""; }
    else if(e.type==="move"){ dotCls="move"; title="Moved"; sub=\`<span class="move-arrow">\${esc(e.from||"—")} <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12h14M13 6l6 6-6 6"/></svg> \${esc(e.to||"—")}</span>\${e.text?\`<br>\${esc(e.text)}\`:""}\`; }
    else if(e.type==="condition"){ dotCls=e.usable===false?"oos":"cond"; title="Condition update"; const bits=[]; if(e.usable===true)bits.push("Marked usable"); if(e.usable===false)bits.push("Marked not usable"); if(e.text)bits.push(esc(e.text)); sub=bits.join(" · "); }
    else if(e.type==="pat"){ dotCls=e.result==="fail"?"oos":"pat"; title=\`PAT test — \${e.result==="fail"?"failed":"passed"}\`; sub=e.expiry?\`Expires \${fmtD(e.expiry)}\`:""; }
    const date = e.effDate?fmtD(e.effDate):tsDate(e.ts);
    const photo = e.photo?\`<div class="ti-photo" style="background-image:url('\${e.photo}')"></div>\`:"";
    return \`<div class="tl-item"><div class="dot \${dotCls}"></div>
      <div class="ti-top"><div class="ti-title">\${title}</div><div class="ti-date">\${date}</div></div>
      \${sub?\`<div class="ti-sub">\${sub}</div>\`:""}\${photo}</div>\`;
  }).join("")+\`</div>\`;
}

/* ============================================================
   CREATE / EDIT  (primary sheet)
   ============================================================ */
let formUsable=true;
function openSheet(id){
  editingId=id;
  const a=id?findAsset(id):null;
  draftPhoto=a?(a.photo||null):null;
  formUsable=a?(a.usable!==false):true;
  $("#sheetTitle").textContent=a?"Edit asset":"Add asset";
  const siteOpts=sites().map(s=>\`<option value="\${esc(s)}">\`).join("");
  const catOpts=CATS.map(c=>\`<option value="\${esc(c)}">\`).join("");
  const todayISO=isoLocal(today0());

  $("#sheetBody").innerHTML=\`
    <div class="photozone \${draftPhoto?"has":""}" id="photozone" \${draftPhoto?\`style="background-image:url('\${draftPhoto}')"\`:""}>
      \${draftPhoto?"":\`<div class="ph-empty"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M14.5 4h-5L7 7H4a2 2 0 0 0-2 2v9a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2V9a2 2 0 0 0-2-2h-3l-2.5-3Z"/><circle cx="12" cy="13" r="3.5"/></svg><span>Tap to add photo</span></div>\`}
      <div class="ph-actions"><button type="button" id="photoBtn"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14.5 4h-5L7 7H4a2 2 0 0 0-2 2v9a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2V9a2 2 0 0 0-2-2h-3l-2.5-3Z"/><circle cx="12" cy="13" r="3.5"/></svg></button>\${draftPhoto?\`<button type="button" id="photoDel"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 6h18M8 6V4h8v2M19 6l-1 14H6L5 6"/></svg></button>\`:""}</div>
    </div>
    <div class="field"><label for="f_name">Asset name <span class="req">*</span></label><input class="inp" id="f_name" placeholder="e.g. 110V transformer" value="\${a?esc(a.name):""}"></div>
    <div class="two">
      <div class="field"><label for="f_cat">Category</label><input class="inp" id="f_cat" list="catList" placeholder="Choose / type" value="\${a?esc(a.category||""):""}"><datalist id="catList">\${catOpts}</datalist></div>
      <div class="field"><label for="f_tag">Asset tag / serial</label><input class="inp mono" id="f_tag" placeholder="e.g. TRF-0142" value="\${a?esc(a.tag||""):""}"></div>
    </div>
    <div class="field"><label for="f_site">\${a?"Current site":"Site"} <span class="req">*</span></label><input class="inp" id="f_site" list="siteList" placeholder="e.g. Riverside Plot 4" value="\${a?esc(a.site||""):""}"><datalist id="siteList">\${siteOpts}</datalist>
      \${a?\`<div class="nextline"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="width:14px;height:14px"><circle cx="12" cy="12" r="9"/><path d="M12 8v4M12 16h.01"/></svg><span>To record a relocation with history, use <b>Move</b> instead of editing this field.</span></div>\`:""}
    </div>
    <div class="field"><label for="f_delivered">Delivered to site</label><input class="inp mono" id="f_delivered" type="date" max="\${todayISO}" value="\${a?esc(a.delivered||""):""}"></div>
    <div class="field"><label>Serviceability</label>
      <div class="seg use" id="useSel"><button type="button" class="yes" data-u="1">Usable</button><button type="button" class="no" data-u="0">Not usable</button></div>
    </div>
    <div class="field"><label for="f_notes">Notes</label><textarea class="inp" id="f_notes" placeholder="Condition, accessories, who delivered it…">\${a?esc(a.notes||""):""}</textarea></div>
    \${id?\`<button class="danger" id="deleteBtn"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 6h18M8 6V4h8v2M19 6l-1 14H6L5 6M10 11v6M14 11v6"/></svg> Delete this asset</button>\`:""}\`;

  // footer
  let foot=$("#sheetFoot"); if(foot) foot.remove();
  foot=el("div","sh-foot"); foot.id="sheetFoot";
  foot.innerHTML=\`<button class="btn btn-ghost" id="cancelBtn">\${editFromDetail?"Back":"Cancel"}</button><button class="btn btn-pri" id="saveBtn">\${id?"Save changes":"Add asset"}</button>\`;
  $("#sheet").appendChild(foot);

  function paintUse(){ document.querySelectorAll("#useSel button").forEach(b=> b.classList.toggle("on", (b.dataset.u==="1")===formUsable)); }
  document.querySelectorAll("#useSel button").forEach(b=> b.onclick=()=>{ formUsable=(b.dataset.u==="1"); paintUse(); });
  paintUse();

  $("#photoBtn").onclick=()=>pickPhoto(u=>{draftPhoto=u;refreshPhotoZone();});
  $("#photozone").onclick=(e)=>{ if(e.target.closest(".ph-actions")) return; pickPhoto(u=>{draftPhoto=u;refreshPhotoZone();}); };
  const pd=$("#photoDel"); if(pd) pd.onclick=()=>{draftPhoto=null;refreshPhotoZone();};

  const db=$("#deleteBtn");
  if(db) db.onclick=async()=>{ if(!confirm("Delete this asset? This cannot be undone.")) return; await deleteAsset(id); await loadAll(); closeSheet(); render(); toast("Asset deleted"); };

  $("#cancelBtn").onclick=()=>{ if(editFromDetail){ editFromDetail=false; openDetail(id); } else closeSheet(); };
  $("#saveBtn").onclick=saveFromForm;
  showSheet();
  setTimeout(()=>{ const n=$("#f_name"); if(n&&!id) n.focus(); },320);
}
function refreshPhotoZone(){
  const z=$("#photozone"); if(!z) return;
  z.classList.toggle("has",!!draftPhoto);
  z.style.backgroundImage=draftPhoto?\`url('\${draftPhoto}')\`:"";
  z.innerHTML = (draftPhoto?"":\`<div class="ph-empty"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M14.5 4h-5L7 7H4a2 2 0 0 0-2 2v9a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2V9a2 2 0 0 0-2-2h-3l-2.5-3Z"/><circle cx="12" cy="13" r="3.5"/></svg><span>Tap to add photo</span></div>\`)
    + \`<div class="ph-actions"><button type="button" id="photoBtn"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14.5 4h-5L7 7H4a2 2 0 0 0-2 2v9a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2V9a2 2 0 0 0-2-2h-3l-2.5-3Z"/><circle cx="12" cy="13" r="3.5"/></svg></button>\${draftPhoto?\`<button type="button" id="photoDel"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 6h18M8 6V4h8v2M19 6l-1 14H6L5 6"/></svg></button>\`:""}</div>\`;
  $("#photoBtn").onclick=()=>pickPhoto(u=>{draftPhoto=u;refreshPhotoZone();});
  const pd=$("#photoDel"); if(pd) pd.onclick=()=>{draftPhoto=null;refreshPhotoZone();};
}

async function saveFromForm(){
  const name=$("#f_name").value.trim(), site=$("#f_site").value.trim();
  if(!name){ toast("Please enter an asset name"); $("#f_name").focus(); return; }
  if(!site){ toast("Please enter a site"); $("#f_site").focus(); return; }
  const isNew=!editingId;
  let a = isNew ? {id:uid(),patTests:[],activity:[],interval:3,created:Date.now()} : Object.assign({},findAsset(editingId));
  const prevUsable = a.usable;
  a.name=name; a.category=$("#f_cat").value.trim(); a.tag=$("#f_tag").value.trim();
  a.site=site; a.delivered=$("#f_delivered").value||""; a.notes=$("#f_notes").value.trim();
  a.photo=draftPhoto||null; a.usable=formUsable;
  if(isNew){ logActivity(a,{type:"create",to:site,effDate:a.delivered||isoLocal(today0())}); }
  else if(prevUsable!==formUsable){ logActivity(a,{type:"condition",usable:formUsable,text:"Updated from edit screen"}); }
  await saveAsset(a);
  await loadAll();
  const back=editFromDetail; editFromDetail=false;
  if(back){ openDetail(a.id); } else { closeSheet(); }
  render();
  toast(isNew?"Asset added":"Asset updated");
}

/* ---------- primary sheet open/close ---------- */
function showSheet(){ $("#scrim").classList.add("show"); $("#sheet").classList.add("show"); $("#sheet").setAttribute("aria-hidden","false"); document.body.style.overflow="hidden"; }
function closeSheet(){ $("#scrim").classList.remove("show"); $("#sheet").classList.remove("show"); $("#sheet").setAttribute("aria-hidden","true"); document.body.style.overflow=""; editingId=null; editFromDetail=false; draftPhoto=null; }
$("#scrim").onclick=closeSheet; $("#sheetClose").onclick=()=>{ if(editFromDetail){ const id=editingId; editFromDetail=false; openDetail(id);} else closeSheet(); };

/* ============================================================
   MODAL LAYER  (move / update / pat)  — above the detail sheet
   ============================================================ */
let modalPhoto=null;
function showModal(title){ $("#modalTitle").textContent=title; $("#modalScrim").classList.add("show"); $("#modal").classList.add("show"); $("#modal").setAttribute("aria-hidden","false"); }
function closeModal(){ $("#modalScrim").classList.remove("show"); $("#modal").classList.remove("show"); $("#modal").setAttribute("aria-hidden","true"); modalPhoto=null; }
$("#modalScrim").onclick=closeModal; $("#modalClose").onclick=closeModal;
function modalFoot(saveLabel){
  let f=$("#modalFoot"); if(f) f.remove();
  f=el("div","sh-foot"); f.id="modalFoot";
  f.innerHTML=\`<button class="btn btn-ghost" id="mCancel">Cancel</button><button class="btn btn-pri" id="mSave">\${saveLabel}</button>\`;
  $("#modal").appendChild(f); $("#mCancel").onclick=closeModal;
}
function refreshDetailSections(a){ const p=$("#patList"); if(p)p.innerHTML=patListHTML(a); const ac=$("#actList"); if(ac)ac.innerHTML=activityHTML(a); wireCertButtons(a); }

/* MOVE */
function openMove(id){
  const a=findAsset(id); if(!a) return;
  const siteOpts=sites().filter(s=>s!==a.site).map(s=>\`<option value="\${esc(s)}">\`).join("");
  const todayISO=isoLocal(today0());
  $("#modalBody").innerHTML=\`
    <div class="nextline" style="margin:0 0 16px"><span>Currently at <b>\${esc(a.site||"—")}</b></span></div>
    <div class="field"><label for="m_site">New location <span class="req">*</span></label><input class="inp" id="m_site" list="moveSites" placeholder="e.g. Eastgate Phase 2" autocomplete="off"><datalist id="moveSites">\${siteOpts}</datalist></div>
    <div class="field"><label for="m_date">Date moved</label><input class="inp mono" id="m_date" type="date" max="\${todayISO}" value="\${todayISO}"></div>
    <div class="field"><label for="m_note">Note (optional)</label><input class="inp" id="m_note" placeholder="e.g. transferred with the breaker set"></div>\`;
  modalFoot("Move asset"); showModal("Move asset");
  setTimeout(()=>$("#m_site").focus(),320);
  $("#mSave").onclick=async()=>{
    const to=$("#m_site").value.trim();
    if(!to){ toast("Enter the new location"); $("#m_site").focus(); return; }
    const from=a.site||""; const date=$("#m_date").value||todayISO;
    a.site=to; logActivity(a,{type:"move",from,to,text:$("#m_note").value.trim(),effDate:date});
    await saveAsset(a); closeModal(); openDetail(id); render(); toast("Moved to "+to);
  };
}

/* UPDATE CONDITION (photo + comment + usable) */
function openUpdate(id){
  const a=findAsset(id); if(!a) return;
  modalPhoto=null;
  let mUse=a.usable!==false;
  $("#modalBody").innerHTML=\`
    <div class="field"><label>Current condition photo</label><div class="mphoto" id="mPhotoWrap"></div></div>
    <div class="field"><label for="u_comment">Comment</label><textarea class="inp" id="u_comment" placeholder="e.g. casing cracked, guard missing, serviced and cleaned…"></textarea></div>
    <div class="field"><label>Serviceability</label><div class="seg use" id="uUse"><button type="button" class="yes" data-u="1">Usable</button><button type="button" class="no" data-u="0">Not usable</button></div></div>\`;
  modalFoot("Save update"); showModal("Update condition");
  function paintU(){ document.querySelectorAll("#uUse button").forEach(b=> b.classList.toggle("on",(b.dataset.u==="1")===mUse)); }
  document.querySelectorAll("#uUse button").forEach(b=> b.onclick=()=>{ mUse=(b.dataset.u==="1"); paintU(); });
  paintU();
  function paintPhoto(){
    const w=$("#mPhotoWrap"); if(!w) return;
    if(modalPhoto){ w.innerHTML=\`<div class="prev" style="background-image:url('\${modalPhoto}')"></div><button type="button" class="rm" id="mRm">Remove</button>\`; $("#mRm").onclick=()=>{modalPhoto=null;paintPhoto();}; }
    else { w.innerHTML=\`<div class="prev">\${boxIcon}</div><button type="button" class="addbtn" id="mAdd"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14.5 4h-5L7 7H4a2 2 0 0 0-2 2v9a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2V9a2 2 0 0 0-2-2h-3l-2.5-3Z"/><circle cx="12" cy="13" r="3.5"/></svg> Take / choose photo</button>\`; $("#mAdd").onclick=()=>pickPhoto(u=>{modalPhoto=u;paintPhoto();}); }
  }
  paintPhoto();
  $("#mSave").onclick=async()=>{
    const comment=$("#u_comment").value.trim();
    const usableChanged = (a.usable!==false)!==mUse;
    if(!modalPhoto && !comment && !usableChanged){ toast("Add a photo, comment or change status"); return; }
    if(modalPhoto) a.photo=modalPhoto;            // newest photo = current condition
    a.usable=mUse;
    logActivity(a,{type:"condition",photo:modalPhoto||null,text:comment,usable:mUse});
    await saveAsset(a); closeModal(); openDetail(id); render(); toast("Condition updated");
  };
}

/* ADD PAT TEST */
function openPat(id){
  const a=findAsset(id); if(!a) return;
  let pInt=a.interval||3, pRes="pass";
  const todayISO=isoLocal(today0());
  $("#modalBody").innerHTML=\`
    <div class="field"><label for="p_date">Test date <span class="req">*</span></label><input class="inp mono" id="p_date" type="date" max="\${todayISO}" value="\${todayISO}"></div>
    <div class="field"><label>Retest every</label><div class="seg" id="pInt"><button type="button" data-m="3">3 mo</button><button type="button" data-m="6">6 mo</button><button type="button" data-m="12">12 mo</button></div>
      <div class="nextline" id="pNext"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="width:15px;height:15px"><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/></svg><span id="pNextTxt"></span></div></div>
    <div class="field"><label>Result</label><div class="seg res" id="pRes"><button type="button" class="pass" data-r="pass">Pass</button><button type="button" class="fail" data-r="fail">Fail</button></div></div>
    <div class="field"><label for="p_tester">Tested by (optional)</label><input class="inp" id="p_tester" placeholder="e.g. J. Doe / Acme PAT Ltd"></div>
    <div class="field"><label for="p_notes">Notes (optional)</label><input class="inp" id="p_notes" placeholder="e.g. replaced plug, retest passed"></div>\`;
  modalFoot("Save test"); showModal("Add PAT test");
  function paintInt(){ document.querySelectorAll("#pInt button").forEach(b=> b.classList.toggle("on",Number(b.dataset.m)===pInt)); updNext(); }
  function paintRes(){ document.querySelectorAll("#pRes button").forEach(b=> b.classList.toggle("on",b.dataset.r===pRes)); }
  function updNext(){
    const d=$("#p_date").value; const t=$("#pNextTxt");
    if(!d){ t.textContent="Pick a test date to see the expiry."; return; }
    const due=addMonths(parseD(d),pInt); const days=Math.round((due-today0())/86400000);
    t.innerHTML=\`Expires <b>\${fmtD(due)}</b>\${days>=0?\` · in \${days} day\${days===1?"":"s"}\`:\` · \${Math.abs(days)} day\${Math.abs(days)===1?"":"s"} ago\`}\`;
  }
  document.querySelectorAll("#pInt button").forEach(b=> b.onclick=()=>{pInt=Number(b.dataset.m);paintInt();});
  document.querySelectorAll("#pRes button").forEach(b=> b.onclick=()=>{pRes=b.dataset.r;paintRes();});
  $("#p_date").addEventListener("change",updNext);
  paintInt(); paintRes();
  $("#mSave").onclick=async()=>{
    const date=$("#p_date").value;
    if(!date){ toast("Pick a test date"); return; }
    const test={id:uid(),date,interval:pInt,result:pRes,tester:$("#p_tester").value.trim(),notes:$("#p_notes").value.trim()};
    a.patTests=a.patTests||[]; a.patTests.push(test); a.interval=pInt;
    const exp=patExpiry(test,a);
    logActivity(a,{type:"pat",result:pRes,expiry:isoLocal(exp),effDate:date});
    let msg="PAT test added";
    if(pRes==="fail"){ a.usable=false; logActivity(a,{type:"condition",usable:false,text:"Failed PAT test"}); msg="Recorded as failed — asset set out of service"; }
    await saveAsset(a); closeModal(); openDetail(id); render(); toast(msg);
  };
}

/* ---------- shared photo capture + compress ---------- */
function pickPhoto(cb){ photoSink=cb; const fi=$("#fileInput"); fi.value=""; fi.click(); }
$("#fileInput").addEventListener("change",function(){
  const f=this.files&&this.files[0]; if(!f) return;
  const reader=new FileReader();
  reader.onload=e=>{ const img=new Image(); img.onload=()=>{
    const max=1024; let w=img.width,h=img.height;
    if(w>max||h>max){ const r=Math.min(max/w,max/h); w=Math.round(w*r); h=Math.round(h*r); }
    const c=document.createElement("canvas"); c.width=w; c.height=h; c.getContext("2d").drawImage(img,0,0,w,h);
    const data=c.toDataURL("image/jpeg",0.72); if(photoSink) photoSink(data); photoSink=null;
  }; img.src=e.target.result; };
  reader.readAsDataURL(f);
});

/* ---------- toast ---------- */
let toastT=null;
function toast(msg){
  const t=$("#toast"); $("#toastMsg").textContent=msg;
  const old=t.querySelector("svg"); if(old) old.remove();
  t.insertAdjacentHTML("afterbegin",\`<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6 9 17l-5-5"/></svg>\`);
  t.classList.add("show"); clearTimeout(toastT); toastT=setTimeout(()=>t.classList.remove("show"),2400);
}

/* ============================================================
   PAT CERTIFICATE  (document generator)
   ============================================================ */
function generateCert(assetId,testId){
  const a=findAsset(assetId); if(!a) return;
  const t=(a.patTests||[]).find(x=>x.id===testId); if(!t) return;
  const exp=patExpiry(t,a);
  const pass=t.result!=="fail";
  const html=certHTML(a,t,exp,pass);
  const fname=\`PAT-cert-\${(a.tag||a.name||"asset").replace(/[^\\w-]+/g,"_")}-\${t.date}.html\`;
  const ok=download(fname,html,"text/html");
  toast(ok?"Certificate downloaded — open & print to PDF":"Use the on-screen copy instead");
  if(!ok){ // fallback: show in modal to copy
    $("#modalBody").innerHTML=\`<p class="muted">Couldn't download automatically. Copy this certificate file and save it as <b>\${esc(fname)}</b>, then open it in a browser to print.</p><textarea class="exp" readonly id="certArea" style="height:240px">\${esc(html)}</textarea>\`;
    modalFoot("Copy"); showModal("PAT certificate");
    $("#mSave").textContent="Copy"; $("#mSave").onclick=async()=>{ const ta=$("#certArea"); try{ await navigator.clipboard.writeText(ta.value); toast("Certificate copied"); }catch(e){ ta.focus(); ta.select(); toast("Selected — copy manually"); } };
  }
}
function certHTML(a,t,exp,pass){
  const E=esc;
  return \`<!doctype html><html lang="en"><head><meta charset="utf-8"><title>PAT Certificate — \${E(a.name)}</title>
<style>
@page{size:A4;margin:18mm}
*{box-sizing:border-box;font-family:-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif}
body{margin:0;color:#1f2933;background:#f3f2ee;padding:24px}
.sheet{max-width:760px;margin:0 auto;background:#fff;border:1px solid #e2e0da;border-radius:10px;overflow:hidden}
.head{background:#1f2933;color:#fff;padding:22px 26px;display:flex;justify-content:space-between;align-items:flex-start}
.head h1{font-size:20px;letter-spacing:.06em;text-transform:uppercase;margin:0}
.head .sub{font-size:11px;letter-spacing:.18em;text-transform:uppercase;color:#9aa5b1;margin-top:6px}
.mark{background:#ffb000;color:#3a2a00;font-weight:800;border-radius:8px;padding:8px 12px;font-size:13px;letter-spacing:.12em}
.body{padding:26px}
.result{display:flex;align-items:center;gap:16px;margin-bottom:22px}
.stamp{font-size:26px;font-weight:800;letter-spacing:.08em;padding:10px 22px;border-radius:8px;border:3px solid}
.pass{color:#1c6b3c;border-color:#2e9e5b;background:#e7f4ec}
.fail{color:#a32525;border-color:#d64545;background:#fbe6e6}
.result .meta{font-size:13px;color:#52606d}
table{width:100%;border-collapse:collapse;margin-bottom:8px}
td{padding:11px 6px;border-bottom:1px solid #eceae4;font-size:14px;vertical-align:top}
td.k{width:42%;color:#7b8794;text-transform:uppercase;letter-spacing:.06em;font-size:11px;font-weight:700;padding-top:13px}
.expiry{margin-top:18px;padding:14px 16px;border-radius:8px;background:#f3f2ee;border:1px solid #e2e0da;display:flex;justify-content:space-between;font-size:14px}
.expiry b{font-size:16px}
.sign{display:flex;gap:40px;margin-top:34px}
.sign div{flex:1}
.line{border-top:1px solid #1f2933;margin-top:34px;padding-top:6px;font-size:11px;color:#7b8794;text-transform:uppercase;letter-spacing:.06em}
.foot{padding:16px 26px;border-top:1px solid #eceae4;font-size:11px;color:#9aa5b1;display:flex;justify-content:space-between}
.print{position:fixed;right:18px;bottom:18px;background:#ffb000;color:#3a2a00;border:none;border-radius:10px;padding:13px 20px;font-weight:800;letter-spacing:.05em;text-transform:uppercase;font-size:13px;cursor:pointer;box-shadow:0 4px 14px rgba(0,0,0,.2)}
@media print{.print{display:none}body{background:#fff;padding:0}.sheet{border:none}}
</style></head><body>
<div class="sheet">
  <div class="head"><div><h1>Portable Appliance Test</h1><div class="sub">Test certificate</div></div><div class="mark">SITEKIT</div></div>
  <div class="body">
    <div class="result"><span class="stamp \${pass?"pass":"fail"}">\${pass?"PASS":"FAIL"}</span><div class="meta">Test performed on <b>\${fmtD(t.date)}</b>\${t.tester?\`<br>by \${E(t.tester)}\`:""}</div></div>
    <table>
      <tr><td class="k">Asset</td><td>\${E(a.name)}</td></tr>
      <tr><td class="k">Category</td><td>\${E(a.category||"—")}</td></tr>
      <tr><td class="k">Asset tag / serial</td><td>\${E(a.tag||"—")}</td></tr>
      <tr><td class="k">Location at test</td><td>\${E(a.site||"—")}</td></tr>
      <tr><td class="k">Test result</td><td>\${pass?"Pass — safe for use":"Fail — removed from service"}</td></tr>
      <tr><td class="k">Retest interval</td><td>Every \${t.interval||a.interval||3} months</td></tr>
      \${t.notes?\`<tr><td class="k">Notes</td><td>\${E(t.notes)}</td></tr>\`:""}
    </table>
    <div class="expiry"><span>Next test due</span><b>\${fmtD(exp)}</b></div>
    <div class="sign"><div><div class="line">Tester signature</div></div><div><div class="line">Date</div></div></div>
  </div>
  <div class="foot"><span>Generated by SiteKit asset register</span><span>\${fmtD(isoLocal(new Date()))}</span></div>
</div>
<button class="print" onclick="window.print()">Print / Save as PDF</button>
</body></html>\`;
}

/* ============================================================
   EXPORT / SHARE WITH HEAD OFFICE
   ============================================================ */
function buildCSV(){
  const head=["Name","Category","Asset tag","Current site","Delivered","Usable","Last PAT","Last result","Next PAT due","PAT status","PAT tests (count)","Notes"];
  const rows=state.assets.map(a=>{
    const s=patStatus(a), lp=latestPat(a);
    const st=s.key==="none"?"No test":s.key==="fail"?"Failed":s.key==="over"?"Overdue":s.key==="due"?"Due soon":"Valid";
    const cells=[a.name,a.category,a.tag,a.site,a.delivered,(a.usable===false?"No":"Yes"),lp?lp.date:"",lp?(lp.result==="fail"?"Fail":"Pass"):"",s.due?isoLocal(s.due):"",st,(a.patTests||[]).length,(a.notes||"").replace(/\\n/g," ")];
    return cells.map(v=>{ v=v==null?"":String(v); return /[",\\n]/.test(v)?'"'+v.replace(/"/g,'""')+'"':v; }).join(",");
  });
  return [head.join(","),...rows].join("\\n");
}
function buildJSON(){
  return JSON.stringify({
    exported:new Date().toISOString(), site_count:sites().length, asset_count:state.assets.length,
    assets:state.assets.map(a=>{ const {photo,activity,patTests,...rest}=a; return {...rest, patTests:patTests||[], activity:(activity||[]).map(({photo,...e})=>e)}; })
  },null,2);
}
function download(filename,text,mime){
  try{ const blob=new Blob([text],{type:mime}); const url=URL.createObjectURL(blob); const a=document.createElement("a"); a.href=url; a.download=filename; document.body.appendChild(a); a.click(); setTimeout(()=>{URL.revokeObjectURL(url);a.remove();},500); return true; }catch(e){ return false; }
}
async function copyText(t){ try{ await navigator.clipboard.writeText(t); return true; }catch(e){ return false; } }

$("#exportBtn").onclick=()=>{
  if(state.assets.length===0){ toast("Add some assets first"); return; }
  closeModal();
  $("#sheetTitle").textContent="Share with head office";
  const csv=buildCSV();
  $("#sheetBody").innerHTML=\`
    <p style="font-size:14px;color:var(--ink-soft);line-height:1.5;margin:2px 2px 16px">Export the full register to send or upload to the parent company. In a live system this would sync automatically.</p>
    <div class="export-card"><h4>Spreadsheet (CSV)</h4><p>Opens in Excel or Sheets — every asset, location, PAT date and status.</p>
      <div class="row"><button class="btn btn-pri" id="csvDl"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><path d="M7 10l5 5 5-5"/><path d="M12 15V3"/></svg> Download CSV</button><button class="btn btn-ghost" id="csvCopy">Copy</button></div></div>
    <div class="export-card"><h4>Data backup (JSON)</h4><p>Full structured backup including PAT history for re-import or integration.</p>
      <div class="row"><button class="btn btn-pri" id="jsonDl"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><path d="M7 10l5 5 5-5"/><path d="M12 15V3"/></svg> Download JSON</button><button class="btn btn-ghost" id="jsonCopy">Copy</button></div></div>
    <label class="field" style="margin-top:4px"><span style="font-family:'Barlow Condensed';font-weight:600;font-size:13px;letter-spacing:.08em;text-transform:uppercase;color:var(--ink-soft);display:block;margin-bottom:6px">Preview (CSV)</span></label>
    <textarea class="exp" id="expArea" readonly>\${esc(csv)}</textarea>\`;
  let foot=$("#sheetFoot"); if(foot) foot.remove();
  foot=el("div","sh-foot"); foot.id="sheetFoot"; foot.innerHTML=\`<button class="btn btn-ghost btn-block" id="expClose">Done</button>\`;
  $("#sheet").appendChild(foot); $("#expClose").onclick=closeSheet;
  const stamp=isoLocal(new Date());
  $("#csvDl").onclick=()=>{ download(\`site-assets-\${stamp}.csv\`,csv,"text/csv")?toast("CSV downloaded"):toast("Use Copy instead"); };
  $("#jsonDl").onclick=()=>{ download(\`site-assets-\${stamp}.json\`,buildJSON(),"application/json")?toast("JSON downloaded"):toast("Use Copy instead"); };
  $("#csvCopy").onclick=async()=>{ (await copyText(csv))?toast("CSV copied"):selArea(); };
  $("#jsonCopy").onclick=async()=>{ const ar=$("#expArea"); ar.value=buildJSON(); (await copyText(ar.value))?toast("JSON copied"):selArea(); };
  showSheet();
};
function selArea(){ const a=$("#expArea"); a.focus(); a.select(); toast("Selected — copy manually"); }

/* ============================================================
   SAMPLE DATA  (new shape)
   ============================================================ */
async function loadSamples(){
  const t=today0(); const ago=d=>{ const x=new Date(t); x.setDate(x.getDate()-d); return isoLocal(x); };
  const now=Date.now(), DAY=86400000;
  const mk=(o)=>{ o.id=uid(); o.created=now; if(o.usable===undefined)o.usable=true; return o; };
  const samples=[
    mk({name:"110V site transformer 3.3kVA",category:"Power tools",tag:"TRF-0142",site:"Riverside Plot 4",delivered:ago(40),notes:"Yellow casing, twin outlet.",interval:3,
        patTests:[{id:uid(),date:ago(98),interval:3,result:"pass",tester:"Acme PAT Ltd",notes:""}],
        activity:[{id:uid(),ts:now-40*DAY,type:"create",to:"Riverside Plot 4",effDate:ago(40)}]}),
    mk({name:"Kango breaker K900",category:"Power tools",tag:"BRK-0098",site:"Riverside Plot 4",delivered:ago(22),notes:"In flight case with chisel set.",interval:3,
        patTests:[{id:uid(),date:ago(20),interval:3,result:"pass",tester:"J. Mara",notes:""}],
        activity:[{id:uid(),ts:now-20*DAY,type:"pat",result:"pass",expiry:ago(-70),effDate:ago(20)},{id:uid(),ts:now-22*DAY,type:"create",to:"Riverside Plot 4",effDate:ago(22)}]}),
    mk({name:"LED tower light",category:"Site lighting",tag:"LGT-0031",site:"Riverside Plot 4",delivered:ago(60),notes:"",interval:6,usable:false,
        patTests:[{id:uid(),date:ago(170),interval:6,result:"pass",tester:"Acme PAT Ltd",notes:""}],
        activity:[{id:uid(),ts:now-3*DAY,type:"condition",usable:false,text:"Mast won't extend — sent for repair",photo:null},{id:uid(),ts:now-60*DAY,type:"create",to:"Riverside Plot 4",effDate:ago(60)}]}),
    mk({name:"Belle cement mixer 150L",category:"Plant & machinery",tag:"MIX-0007",site:"Eastgate Phase 2",delivered:ago(12),notes:"",interval:6,
        patTests:[{id:uid(),date:ago(5),interval:6,result:"pass",tester:"",notes:""}],
        activity:[{id:uid(),ts:now-9*DAY,type:"move",from:"Riverside Plot 4",to:"Eastgate Phase 2",text:"Moved with the gang",effDate:ago(9)},{id:uid(),ts:now-12*DAY,type:"create",to:"Riverside Plot 4",effDate:ago(40)}]}),
    mk({name:"Extension lead 25m 110V",category:"Leads & cables",tag:"LD-0204",site:"Eastgate Phase 2",delivered:ago(12),notes:"Awaiting first PAT test.",interval:3,
        patTests:[],activity:[{id:uid(),ts:now-12*DAY,type:"create",to:"Eastgate Phase 2",effDate:ago(12)}]}),
    mk({name:"Podium step access platform",category:"Access equipment",tag:"ACC-0019",site:"Eastgate Phase 2",delivered:ago(30),notes:"",interval:6,
        patTests:[{id:uid(),date:ago(85),interval:6,result:"pass",tester:"",notes:""}],
        activity:[{id:uid(),ts:now-30*DAY,type:"create",to:"Eastgate Phase 2",effDate:ago(30)}]}),
  ];
  for(const s of samples){ s.photo=null; await saveAsset(s); }
  await loadAll(); render(); toast("Sample assets loaded");
}

/* ---------- boot ---------- */
(async function init(){ await loadAll(); render(); })();
</script>
</body>
</html>
`;
  return(
    <div style={{position:"fixed",top:0,left:200,right:0,bottom:0,background:"#f3f2ee",zIndex:5,display:"flex",flexDirection:"column"}}>
      <iframe
        srcDoc={htmlDoc}
        style={{flex:1,width:"100%",border:"none",display:"block"}}
        title="SiteKit Asset Register"
        sandbox="allow-scripts allow-same-origin allow-forms"
      />
    </div>
  );
}
// ─── LOGIN GATE ───────────────────────────────────────────────────────────────
// ─── Hardcoded credentials (work without Supabase — for preview & first login)
const HARDCODED_USERS=[
  {email:"lucian@bright-group.org", password:"BrightAdmin2025", role:"admin",  name:"Lucian Ciocoiu",    note:"Director"},
  {email:"admin@bright-group.org",  password:"BrightAdmin2025", role:"admin",  name:"Admin",             note:"Admin"},
  {email:"demo@bright.app",         password:"demo1234",         role:"admin",  name:"Demo Admin",        note:"Preview / Demo"},
  {email:"manager@bright.app",      password:"manager1234",      role:"manager",name:"Demo Manager",      note:"Manager preview"},
  {email:"worker@bright.app",       password:"worker1234",       role:"worker", name:"Demo Worker",       note:"Worker preview"},
];

function LoginGate({onLogin}){
  const [email,setEmail]=useState("");
  const [password,setPassword]=useState("");
  const [loading,setLoading]=useState(false);
  const [error,setError]=useState("");
  const [showPw,setShowPw]=useState(false);
  const [showCreds,setShowCreds]=useState(false);

  const handleLogin=async(e)=>{
    e&&e.preventDefault();
    if(!email||!password){setError("Please enter email and password.");return;}
    setLoading(true);setError("");
    try{
      // ── Step 1: Check hardcoded credentials first (always works, no Supabase needed)
      const hardcoded=HARDCODED_USERS.find(
        u=>u.email.toLowerCase()===email.toLowerCase().trim()&&u.password===password
      );
      if(hardcoded){
        // Try to fetch real worker data from Supabase if available
        let worker=null;
        try{
          const rows=await sbGet("workers",`select=id,data&data->>email=eq.${encodeURIComponent(hardcoded.email)}`);
          if(rows.length>0) worker={...mkW(),...rows[0].data,id:rows[0].id};
        }catch(_){}
        // If no Supabase worker record, create minimal profile from hardcoded data
        if(!worker) worker={...mkW(),name:hardcoded.name,email:hardcoded.email,userRole:hardcoded.role,id:"local_"+hardcoded.email};
        onLogin({worker,role:hardcoded.role,email:hardcoded.email});
        return;
      }

      // ── Step 2: Try Supabase Auth for real registered users
      let authToken=null;
      try{
        const authRes=await fetch(`${SB_URL}/auth/v1/token?grant_type=password`,{
          method:"POST",
          headers:{"apikey":SB_KEY,"Content-Type":"application/json"},
          body:JSON.stringify({email:email.trim(),password})
        });
        const authData=await authRes.json();
        if(!authData.error) authToken=authData.access_token;
      }catch(_){}

      // ── Step 3: Look up worker in Supabase by email
      let worker=null;
      try{
        const rows=await sbGet("workers",`select=id,data&data->>email=eq.${encodeURIComponent(email.trim())}`);
        if(rows.length>0) worker={...mkW(),...rows[0].data,id:rows[0].id};
        if(!worker){
          const rows2=await sbGet("workers",`select=id,data&data->>authEmail=eq.${encodeURIComponent(email.trim())}`);
          if(rows2.length>0) worker={...mkW(),...rows2[0].data,id:rows2[0].id};
        }
      }catch(_){}

      if(!worker&&!authToken){
        setError("No account found. Check your email and password, or use the credentials shown below.");
        setShowCreds(true);
        setLoading(false);return;
      }

      const role=worker?.userRole||"worker";
      onLogin({worker,role,email:email.trim()});
    }catch(e){
      setError("Login error: "+e.message);
    }
    setLoading(false);
  };

  const fillCreds=(u)=>{setEmail(u.email);setPassword(u.password);setShowCreds(false);};

  return(
    <div style={{minHeight:"100vh",background:"#0a0e1a",display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"'Inter',system-ui,sans-serif",padding:16}}>
      <div style={{width:"100%",maxWidth:400}}>
        {/* Logo */}
        <div style={{textAlign:"center",marginBottom:36}}>
          <div style={{width:64,height:64,background:"linear-gradient(135deg,#1e3a5f,#3b82f6)",borderRadius:18,display:"flex",alignItems:"center",justifyContent:"center",fontSize:30,margin:"0 auto 16px",boxShadow:"0 8px 32px rgba(59,130,246,.3)"}}>🏗</div>
          <div style={{fontSize:24,fontWeight:800,color:"#f1f5f9",letterSpacing:"-0.02em"}}>Bright Metalwork</div>
          <div style={{fontSize:13,color:"#64748b",marginTop:4}}>Construction Management Platform</div>
        </div>

        {/* Login card */}
        <form onSubmit={handleLogin} style={{background:"#111827",borderRadius:16,padding:28,border:"1px solid #1e2535",boxShadow:"0 24px 64px rgba(0,0,0,.5)"}}>
          <div style={{fontSize:16,fontWeight:700,color:"#f1f5f9",marginBottom:20}}>Sign in to your account</div>

          {error&&<div style={{background:"#2d1515",border:"1px solid #f8717144",borderRadius:9,padding:"10px 13px",fontSize:12,color:"#f87171",marginBottom:14}}>{error}</div>}

          <div style={{marginBottom:14}}>
            <label style={{display:"block",fontSize:11,color:"#64748b",fontWeight:700,textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:5}}>Email</label>
            <input type="email" value={email} onChange={e=>setEmail(e.target.value)} placeholder="your@email.com" autoComplete="email"
              style={{width:"100%",background:"#0d1117",border:"1px solid #1e2535",borderRadius:9,padding:"11px 13px",color:"#f1f5f9",fontSize:14,outline:"none",boxSizing:"border-box"}}
              onFocus={e=>e.target.style.borderColor="#3b82f6"} onBlur={e=>e.target.style.borderColor="#1e2535"}/>
          </div>

          <div style={{marginBottom:20}}>
            <label style={{display:"block",fontSize:11,color:"#64748b",fontWeight:700,textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:5}}>Password</label>
            <div style={{position:"relative"}}>
              <input type={showPw?"text":"password"} value={password} onChange={e=>setPassword(e.target.value)} placeholder="••••••••" autoComplete="current-password"
                style={{width:"100%",background:"#0d1117",border:"1px solid #1e2535",borderRadius:9,padding:"11px 40px 11px 13px",color:"#f1f5f9",fontSize:14,outline:"none",boxSizing:"border-box"}}
                onFocus={e=>e.target.style.borderColor="#3b82f6"} onBlur={e=>e.target.style.borderColor="#1e2535"}/>
              <button type="button" onClick={()=>setShowPw(s=>!s)}
                style={{position:"absolute",right:11,top:"50%",transform:"translateY(-50%)",background:"none",border:"none",cursor:"pointer",fontSize:14,color:"#64748b"}}>
                {showPw?"🙈":"👁"}
              </button>
            </div>
          </div>

          <button type="submit" disabled={loading}
            style={{width:"100%",padding:"12px",background:loading?"#1e3a5f":"linear-gradient(135deg,#1e3a5f,#3b82f6)",border:"none",borderRadius:9,color:"#fff",fontSize:14,fontWeight:700,cursor:loading?"not-allowed":"pointer",transition:"opacity .15s",opacity:loading?0.7:1}}>
            {loading?"Signing in…":"Sign In"}
          </button>

          {/* ── Quick credentials panel ── */}
          <div style={{marginTop:14,background:"#0d1117",borderRadius:9,border:"1px solid #1e2535",overflow:"hidden"}}>
            <button type="button" onClick={()=>setShowCreds(s=>!s)}
              style={{width:"100%",padding:"9px 13px",background:"none",border:"none",cursor:"pointer",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <span style={{fontSize:11,color:"#64748b",fontWeight:600}}>🔑 Login credentials</span>
              <span style={{fontSize:10,color:"#374151"}}>{showCreds?"▲ Hide":"▼ Show"}</span>
            </button>
            {showCreds&&<div style={{padding:"0 13px 12px"}}>
              {HARDCODED_USERS.map(u=>(
                <button type="button" key={u.email} onClick={()=>fillCreds(u)}
                  style={{width:"100%",textAlign:"left",padding:"8px 10px",marginBottom:5,background:"#111827",border:"1px solid #1e2535",borderRadius:7,cursor:"pointer",transition:"border-color .15s"}}
                  onMouseEnter={e=>e.currentTarget.style.borderColor="#3b82f6"}
                  onMouseLeave={e=>e.currentTarget.style.borderColor="#1e2535"}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                    <div>
                      <div style={{fontSize:12,fontWeight:700,color:"#f1f5f9"}}>{u.name} <span style={{fontSize:10,color:"#64748b",fontWeight:400}}>— {u.note}</span></div>
                      <div style={{fontSize:10,color:"#64748b",marginTop:2}}>{u.email} · <span style={{color:"#374151",fontFamily:"monospace"}}>{u.password}</span></div>
                    </div>
                    <span style={{fontSize:10,padding:"2px 7px",borderRadius:20,fontWeight:700,
                      background:{admin:"#34d39922",manager:"#60a5fa22",worker:"#94a3b822"}[u.role],
                      color:{admin:"#34d399",manager:"#60a5fa",worker:"#94a3b8"}[u.role]}}>
                      {u.role}
                    </span>
                  </div>
                </button>
              ))}
              <div style={{fontSize:10,color:"#374151",marginTop:4,fontStyle:"italic"}}>Click any row to fill in the credentials automatically</div>
            </div>}
          </div>
        </form>
      </div>
    </div>
  );
}

// ─── MANAGER VIEW ─────────────────────────────────────────────────────────────
// Manager sees the SiteManagerView (already built) as their full interface
function ManagerView({authState,onLogout}){
  return(
    <div style={{minHeight:"100vh",background:"#0a0e1a",fontFamily:"'Inter',system-ui,sans-serif"}}>
      {/* Manager top bar */}
      <div style={{background:"#0f172a",borderBottom:"1px solid #1e2535",padding:"10px 20px",display:"flex",alignItems:"center",justifyContent:"space-between",position:"sticky",top:0,zIndex:100}}>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <div style={{width:32,height:32,background:"linear-gradient(135deg,#1e3a5f,#3b82f6)",borderRadius:8,display:"flex",alignItems:"center",justifyContent:"center",fontSize:16}}>🏗</div>
          <div>
            <div style={{color:"#f1f5f9",fontSize:13,fontWeight:700}}>Bright Metalwork — Manager Portal</div>
            <div style={{color:"#64748b",fontSize:10}}>{authState.worker?.name||authState.email}</div>
          </div>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <span style={{fontSize:11,padding:"3px 10px",borderRadius:20,background:"#60a5fa22",color:"#60a5fa",border:"1px solid #60a5fa44",fontWeight:700}}>🏗 Manager</span>
          <button onClick={onLogout} style={{padding:"5px 12px",background:"#1e2535",border:"1px solid #2d3555",borderRadius:7,color:"#94a3b8",cursor:"pointer",fontSize:11,fontWeight:600}}>Sign Out</button>
        </div>
      </div>
      {/* Render the Site Manager App — already built */}
      <SiteManagerView/>
    </div>
  );
}

// ─── WORKER SELF VIEW ─────────────────────────────────────────────────────────
// Worker sees their own data — read-only except they can add holidays + sign-in
function WorkerSelfView({authState,allSites,weekLabel,timesheetRecords,payslipRecords,onLogout}){
  const w=authState.worker;
  if(!w) return(
    <div style={{minHeight:"100vh",background:"#0a0e1a",display:"flex",alignItems:"center",justifyContent:"center",color:"#f87171",fontFamily:"system-ui"}}>
      Worker profile not found. Contact your administrator.
    </div>
  );

  const [tab,setTab]=useState("profile");
  const [holidays,setHolidays]=useState(w.holidayRequests||[]);
  const [holFrom,setHolFrom]=useState("");
  const [holTo,setHolTo]=useState("");
  const [holNote,setHolNote]=useState("");

  const heldCerts=CERTS.filter(c=>w.certs?.[c.key]?.held);
  const wLogs=(w.attendanceLogs||[]).filter(l=>l.signIn&&l.signOut).sort((a,b)=>new Date(b.signIn)-new Date(a.signIn));
  const thisWeekLogs=wLogs.filter(l=>l.weekLabel===weekLabel);
  const confirmedHrs=thisWeekLogs.reduce((a,l)=>a+Math.round(((new Date(l.signOut)-new Date(l.signIn))/3600000)*100)/100,0);
  const wTimesheets=(timesheetRecords||[]).filter(t=>t.workerId===w.id).sort((a,b)=>new Date(b.weekLabel)-new Date(a.weekLabel));
  const wPayslips=(payslipRecords||[]).filter(p=>p.workerId===w.id).sort((a,b)=>new Date(b.weekLabel)-new Date(a.weekLabel));
  const initials=(w.name||"?").split(" ").map(n=>n[0]).join("").slice(0,2).toUpperCase();
  const fmtD=d=>d?new Date(d).toLocaleDateString("en-GB",{day:"2-digit",month:"short",year:"numeric"}):"—";
  const fmtT=iso=>iso?new Date(iso).toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit"}):"—";
  const CERT_C={valid:"#34d399",expiring:"#fbbf24",expired:"#f87171",missing:"#374151"};

  const TABS=[["profile","👤 Profile"],["schedule","📅 Schedule"],["timesheets","⏱ Timesheets"],["payslips","💰 Payslips"],["certs","🏅 Certs"],["holidays","🏖 Holidays"]];

  return(
    <div style={{minHeight:"100vh",background:"#0a0e1a",fontFamily:"'Inter',system-ui,sans-serif"}}>
      {/* Header */}
      <div style={{background:"#0f172a",borderBottom:"1px solid #1e2535",padding:"12px 20px",display:"flex",alignItems:"center",justifyContent:"space-between",position:"sticky",top:0,zIndex:100}}>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <div style={{width:38,height:38,borderRadius:10,background:"#3b82f622",border:"1px solid #3b82f644",display:"flex",alignItems:"center",justifyContent:"center",fontSize:15,fontWeight:800,color:"#3b82f6"}}>{initials}</div>
          <div>
            <div style={{color:"#f1f5f9",fontWeight:700,fontSize:14}}>{w.name}</div>
            <div style={{color:"#64748b",fontSize:11}}>{w.position} · {w.company}</div>
          </div>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:8}}>
          <span style={{fontSize:11,padding:"3px 10px",borderRadius:20,background:"#94a3b822",color:"#94a3b8",border:"1px solid #94a3b844",fontWeight:700}}>👷 Worker</span>
          <button onClick={onLogout} style={{padding:"5px 12px",background:"#1e2535",border:"1px solid #2d3555",borderRadius:7,color:"#94a3b8",cursor:"pointer",fontSize:11,fontWeight:600}}>Sign Out</button>
        </div>
      </div>

      {/* KPI strip */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:8,padding:"12px 16px 0"}}>
        {[["Hourly Rate",w.agreedRate?"£"+w.agreedRate+"/hr":"Not set","#34d399"],["CIS Tax",Math.round((w.taxRate||0)*100)+"%",w.taxRate===0.30?"#f87171":"#34d399"],["This Week",confirmedHrs>0?confirmedHrs.toFixed(1)+"h":"—","#60a5fa"],["Timesheets",wTimesheets.length,"#a78bfa"]].map(([l,v,c])=>(
          <div key={l} style={{background:"#111827",border:`1px solid ${c}22`,borderRadius:9,padding:"9px 11px"}}>
            <div style={{fontSize:9,color:"#64748b",fontWeight:700,textTransform:"uppercase"}}>{l}</div>
            <div style={{fontSize:15,fontWeight:800,color:c,marginTop:2}}>{v}</div>
          </div>
        ))}
      </div>

      {/* Tab nav */}
      <div style={{display:"flex",gap:2,padding:"12px 16px 0",overflowX:"auto"}}>
        {TABS.map(([v,l])=>(
          <button key={v} onClick={()=>setTab(v)}
            style={{padding:"6px 12px",borderRadius:7,border:`1px solid ${tab===v?"#3b82f6":"transparent"}`,background:tab===v?"#1e3a5f":"transparent",color:tab===v?"#60a5fa":"#64748b",cursor:"pointer",fontSize:11,fontWeight:tab===v?700:400,whiteSpace:"nowrap"}}>
            {l}
          </button>
        ))}
      </div>

      {/* Tab content */}
      <div style={{padding:16}}>

        {/* PROFILE */}
        {tab==="profile"&&<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
          <div style={{background:"#111827",border:"1px solid #1e2535",borderRadius:11,padding:14}}>
            <div style={{fontSize:11,color:"#60a5fa",fontWeight:700,textTransform:"uppercase",marginBottom:10}}>Personal Details</div>
            {[["Name",w.name],["Position",w.position],["Company",w.company],["Email",w.email||"—"],["Phone",w.contact||"—"],["Address",w.address||"—"]].map(([l,v])=>(
              <div key={l} style={{display:"flex",gap:8,padding:"5px 0",borderBottom:"1px solid #1e2535"}}>
                <span style={{fontSize:10,color:"#64748b",fontWeight:700,minWidth:72,textTransform:"uppercase",flexShrink:0}}>{l}</span>
                <span style={{fontSize:12,color:"#e2e8f0"}}>{v||"—"}</span>
              </div>
            ))}
          </div>
          <div style={{display:"flex",flexDirection:"column",gap:10}}>
            <div style={{background:"#111827",border:"1px solid #1e2535",borderRadius:11,padding:14}}>
              <div style={{fontSize:11,color:"#fbbf24",fontWeight:700,textTransform:"uppercase",marginBottom:10}}>Pay Details</div>
              {[["Rate","£"+(w.agreedRate||0)+"/hr"],["CIS Tax",Math.round((w.taxRate||0)*100)+"%"],["OT Rate","×"+(w.overtimeMultiplier||1.5)],["NI No",w.nino||"—"],["UTR",w.utr||"—"]].map(([l,v])=>(
                <div key={l} style={{display:"flex",gap:8,padding:"5px 0",borderBottom:"1px solid #1e2535"}}>
                  <span style={{fontSize:10,color:"#64748b",fontWeight:700,minWidth:60,textTransform:"uppercase",flexShrink:0}}>{l}</span>
                  <span style={{fontSize:12,color:"#e2e8f0"}}>{v}</span>
                </div>
              ))}
            </div>
            {heldCerts.length>0&&<div style={{background:"#111827",border:"1px solid #1e2535",borderRadius:11,padding:14}}>
              <div style={{fontSize:11,color:"#a78bfa",fontWeight:700,textTransform:"uppercase",marginBottom:8}}>🏅 My Certificates</div>
              {heldCerts.map(cert=>{const s=cSt(cert,w);return(
                <div key={cert.key} style={{display:"flex",alignItems:"center",gap:6,padding:"4px 0",borderBottom:"1px solid #1e2535"}}>
                  <div style={{width:6,height:6,borderRadius:"50%",background:CERT_C[s],flexShrink:0}}/>
                  <span style={{fontSize:11,color:"#e2e8f0",flex:1}}>{cert.label}</span>
                  <span style={{fontSize:9,color:CERT_C[s],fontWeight:700,textTransform:"uppercase"}}>{s}</span>
                </div>
              );})}
            </div>}
          </div>
        </div>}

        {/* SCHEDULE */}
        {tab==="schedule"&&<div>
          <div style={{fontSize:11,color:"#64748b",marginBottom:10}}>Your forecast route for WC {weekLabel}. ✅ = GPS confirmed.</div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:8}}>
            {["Mon","Tue","Wed","Thu","Fri","Sat","Sun"].map(d=>{
              const site=w.days?.[d];const off=!site||isOff(site);
              const confirmed=thisWeekLogs.some(l=>l.day===d);
              const s=allSites.find(x=>site&&(site.includes(x.name)||x.name.includes((site||"").split("-")[0]?.trim())));
              const col=s?.color||"#3b82f6";
              return<div key={d} style={{textAlign:"center",background:"#111827",borderRadius:9,padding:"10px 4px",border:`2px solid ${confirmed?"#34d39944":off?"#1e2535":col+"44"}`}}>
                <div style={{fontSize:9,fontWeight:700,textTransform:"uppercase",color:confirmed?"#34d399":off?"#374151":col,marginBottom:4}}>{d}</div>
                {confirmed&&<div style={{fontSize:11,marginBottom:3}}>✅</div>}
                <div style={{fontSize:9,color:off?"#374151":col,fontWeight:600,lineHeight:1.3,wordBreak:"break-word"}}>{site||"Off"}</div>
              </div>;
            })}
          </div>
        </div>}

        {/* TIMESHEETS — read only */}
        {tab==="timesheets"&&<div>
          {wTimesheets.length===0&&<div style={{textAlign:"center",padding:40,color:"#374151"}}>No timesheets yet.</div>}
          {wTimesheets.map(ts=>{
            const tsLogs=wLogs.filter(l=>l.weekLabel===ts.weekLabel);
            const days=["Mon","Tue","Wed","Thu","Fri","Sat","Sun"];
            const dayTotals={};
            tsLogs.forEach(l=>{const hrs=Math.round(((new Date(l.signOut)-new Date(l.signIn))/3600000)*100)/100;dayTotals[l.day]=(dayTotals[l.day]||0)+hrs;});
            const total=Object.values(dayTotals).reduce((a,v)=>a+v,0);
            return<div key={ts.id} style={{background:"#111827",border:`1px solid ${ts.status==="approved"?"#34d39944":"#1e2535"}`,borderRadius:11,padding:14,marginBottom:10}}>
              <div style={{display:"flex",justifyContent:"space-between",marginBottom:10}}>
                <div style={{fontWeight:700,color:"#f1f5f9"}}>WC {ts.weekLabel}</div>
                <span style={{padding:"2px 9px",borderRadius:20,fontSize:10,fontWeight:700,background:ts.status==="approved"?"#34d39922":"#fbbf2422",color:ts.status==="approved"?"#34d399":"#fbbf24",border:`1px solid ${ts.status==="approved"?"#34d39944":"#fbbf2444"}`}}>{ts.status==="approved"?"✓ Approved":"⏳ Pending"}</span>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:4,marginBottom:8}}>
                {days.map(d=><div key={d} style={{textAlign:"center",background:"#0f1421",borderRadius:5,padding:"5px 2px",border:`1px solid ${dayTotals[d]?"#3b82f633":"#1e2535"}`}}>
                  <div style={{fontSize:8,color:"#64748b",textTransform:"uppercase"}}>{d}</div>
                  <div style={{fontSize:11,fontWeight:700,color:dayTotals[d]?"#60a5fa":"#374151"}}>{dayTotals[d]?dayTotals[d].toFixed(1):"—"}</div>
                </div>)}
              </div>
              <div style={{display:"flex",gap:12,fontSize:11,color:"#64748b"}}>
                <span>Total: <b style={{color:"#34d399"}}>{total.toFixed(1)}h</b></span>
                <span style={{color:"#374151",fontStyle:"italic"}}>Pay amounts shown in payslips</span>
              </div>
            </div>;
          })}
        </div>}

        {/* PAYSLIPS — read only */}
        {tab==="payslips"&&<div>
          {wPayslips.length===0&&<div style={{textAlign:"center",padding:40,color:"#374151"}}>No payslips yet.</div>}
          {wPayslips.map(ps=><div key={ps.id} style={{background:"#111827",border:`1px solid ${ps.status==="paid"?"#34d39944":"#1e2535"}`,borderRadius:11,padding:14,marginBottom:10}}>
            <div style={{display:"flex",justifyContent:"space-between",marginBottom:10}}>
              <div style={{fontWeight:700,color:"#f1f5f9"}}>WC {ps.weekLabel}</div>
              <span style={{padding:"2px 9px",borderRadius:20,fontSize:10,fontWeight:700,background:ps.status==="paid"?"#34d39922":"#fbbf2422",color:ps.status==="paid"?"#34d399":"#fbbf24",border:`1px solid ${ps.status==="paid"?"#34d39944":"#fbbf2444"}`}}>{ps.status==="paid"?"💷 Paid":"⏳ Processing"}</span>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:6}}>
              {[["Std Hours",(ps.stdH||0).toFixed(1)+"h","#60a5fa"],["OT Hours",(ps.otH||0).toFixed(1)+"h","#fbbf24"],["Gross Pay","£"+(ps.gross||0).toFixed(2),"#34d399"],["CIS Tax","-£"+(ps.taxAmt||0).toFixed(2),"#f87171"],["Net Pay","£"+(ps.net||0).toFixed(2),"#a78bfa"],["To Account",w.bankAccount?"••••"+w.bankAccount.slice(-4):"—","#64748b"]].map(([l,v,c])=>(
                <div key={l} style={{background:"#0d1117",borderRadius:7,padding:"8px 10px"}}>
                  <div style={{fontSize:9,color:"#64748b",fontWeight:700,textTransform:"uppercase"}}>{l}</div>
                  <div style={{fontSize:13,fontWeight:700,color:c}}>{v}</div>
                </div>
              ))}
            </div>
          </div>)}
        </div>}

        {/* CERTS — read only */}
        {tab==="certs"&&<div>
          {heldCerts.length===0&&<div style={{textAlign:"center",padding:40,color:"#374151"}}>No certificates recorded. Contact your administrator.</div>}
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
            {heldCerts.map(cert=>{const val=w.certs?.[cert.key];const s=cSt(cert,w);return(
              <div key={cert.key} style={{background:"#111827",borderRadius:9,padding:"11px 13px",border:`1px solid ${CERT_C[s]}44`}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:4}}>
                  <span style={{fontSize:12,fontWeight:700,color:"#f1f5f9"}}>{cert.label}</span>
                  <span style={{fontSize:10,fontWeight:700,color:CERT_C[s],textTransform:"uppercase"}}>{s}</span>
                </div>
                {cert.hasExpiry&&<div style={{fontSize:10,color:"#64748b"}}>Expires: <span style={{color:CERT_C[s],fontWeight:600}}>{val?.expiry?fmtD(val.expiry):"—"}</span></div>}
              </div>
            );})}
          </div>
        </div>}

        {/* HOLIDAYS — can ADD */}
        {tab==="holidays"&&<div>
          <div style={{background:"#111827",border:"1px solid #1e2535",borderRadius:11,padding:14,marginBottom:14}}>
            <div style={{fontSize:11,color:"#34d399",fontWeight:700,textTransform:"uppercase",marginBottom:10}}>Request Holiday</div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 2fr",gap:8,marginBottom:10}}>
              <div><div style={{fontSize:9,color:"#64748b",fontWeight:700,textTransform:"uppercase",marginBottom:4}}>From</div><input type="date" value={holFrom} onChange={e=>setHolFrom(e.target.value)} style={{width:"100%",background:"#0d1117",border:"1px solid #1e2535",borderRadius:7,padding:"8px 10px",color:"#f1f5f9",fontSize:12,outline:"none",boxSizing:"border-box"}}/></div>
              <div><div style={{fontSize:9,color:"#64748b",fontWeight:700,textTransform:"uppercase",marginBottom:4}}>To</div><input type="date" value={holTo} onChange={e=>setHolTo(e.target.value)} style={{width:"100%",background:"#0d1117",border:"1px solid #1e2535",borderRadius:7,padding:"8px 10px",color:"#f1f5f9",fontSize:12,outline:"none",boxSizing:"border-box"}}/></div>
              <div><div style={{fontSize:9,color:"#64748b",fontWeight:700,textTransform:"uppercase",marginBottom:4}}>Note</div><input value={holNote} onChange={e=>setHolNote(e.target.value)} placeholder="e.g. Annual leave" style={{width:"100%",background:"#0d1117",border:"1px solid #1e2535",borderRadius:7,padding:"8px 10px",color:"#f1f5f9",fontSize:12,outline:"none",boxSizing:"border-box"}}/></div>
            </div>
            <button onClick={()=>{if(!holFrom)return;setHolidays(h=>[...h,{id:"h"+Date.now(),from:holFrom,to:holTo||holFrom,note:holNote,status:"pending",requestedAt:new Date().toISOString()}]);setHolFrom("");setHolTo("");setHolNote("");}}
              style={{padding:"8px 16px",background:"linear-gradient(135deg,#1e3a5f,#3b82f6)",border:"none",borderRadius:8,color:"#fff",fontSize:12,fontWeight:700,cursor:"pointer"}}>
              + Submit Request
            </button>
          </div>
          {holidays.length===0&&<div style={{textAlign:"center",padding:32,color:"#374151"}}>No holiday requests yet.</div>}
          {holidays.map(h=><div key={h.id} style={{background:"#111827",border:`1px solid ${h.status==="approved"?"#34d39944":h.status==="declined"?"#f8717144":"#fbbf2444"}`,borderRadius:9,padding:"11px 14px",marginBottom:8,display:"flex",alignItems:"center",gap:10}}>
            <div style={{flex:1}}>
              <div style={{fontSize:12,fontWeight:700,color:"#f1f5f9"}}>{fmtD(h.from)}{h.to&&h.to!==h.from?" → "+fmtD(h.to):""}</div>
              {h.note&&<div style={{fontSize:11,color:"#64748b",marginTop:2}}>{h.note}</div>}
            </div>
            <span style={{padding:"2px 9px",borderRadius:20,fontSize:10,fontWeight:700,background:h.status==="approved"?"#34d39922":h.status==="declined"?"#f8717122":"#fbbf2422",color:h.status==="approved"?"#34d399":h.status==="declined"?"#f87171":"#fbbf24"}}>{h.status}</span>
          </div>)}
        </div>}

      </div>
    </div>
  );
}

// ─── Main App ──────────────────────────────────────────────────────────────────
export default function App(){
  // ── Auth gate ────────────────────────────────────────────────────────────
  const [authState,setAuthState]=useState(null); // {worker, role, email}
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
    const prevWorker=workers.find(w=>w.id===wId);
    const prevVal=prevWorker?.days?.[day]||"";
    const updated=workers.map(w=>w.id===wId?{...w,days:{...w.days,[day]:val}}:w);
    setWorkers(updated); setSyncStatus("saving");
    try {
      const w=updated.find(x=>x.id===wId);
      // ── Route change detection ──
      const hasChange=prevVal.trim()!==val.trim()&&!(isOff(prevVal)&&isOff(val));
      if(hasChange){
        const DAY_FULL={Mon:"Monday",Tue:"Tuesday",Wed:"Wednesday",Thu:"Thursday",Fri:"Friday",Sat:"Saturday",Sun:"Sunday"};
        const changeEntry={id:"rc_"+Date.now(),changedAt:new Date().toISOString(),day,weekLabel,from:prevVal||"(unset)",to:val||"(unset)",seen:false};
        w.routeNotifications=[...(w.routeNotifications||[]).slice(-19),changeEntry];
        w.routeHistory=[...(w.routeHistory||[]).slice(-49),{...changeEntry,seen:undefined}];
        // ── Email notification ──
        const toEmail=w.email||w.authEmail;
        if(toEmail){
          const subject=`Route Update WC ${weekLabel} — ${DAY_FULL[day]||day} changed`;
          const body=`Hi ${w.name},

Your route for the week commencing ${weekLabel} has been updated.

${DAY_FULL[day]||day}: changed from "${prevVal||"(unset)"}" to "${val||"(unset)"}".

Please check the Bright Metalwork Worker Portal for your full updated week ahead.

Bright Metalwork Ltd
+44 (0)771 078 3500
lucian@bright-group.org`;
          try{
            await fetch(`${SB_URL}/functions/v1/send-email`,{
              method:"POST",
              headers:{...SB_H,"Content-Type":"application/json"},
              body:JSON.stringify({to:toEmail,subject,text:body})
            });
          }catch(emailErr){console.warn("Email notification failed:",emailErr.message);}
        }
      }
      await sbUpsert("workers",[{id:wId,data:w}]); setSyncStatus("saved");
    }
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
  // generateTimesheets — kept for backward compat but DTimesheets now builds its own
  const buildTimesheetFromLogs=(worker,wkLabel)=>{
    const wkLogs=(worker.attendanceLogs||[]).filter(l=>l.weekLabel===wkLabel&&l.signIn&&l.signOut);
    if(wkLogs.length===0) return null;
    const dayMap={};
    wkLogs.forEach(l=>{
      const hrs=Math.round(((new Date(l.signOut)-new Date(l.signIn))/3600000)*100)/100;
      const site=allSites.find(s=>s.name===l.siteName||s.id===l.siteId);
      const otThr=site?.otThreshold||site?.stdHours||9;
      if(!dayMap[l.day]) dayMap[l.day]={total:0,std:0,ot:0,site:l.siteName};
      dayMap[l.day].total=Math.round((dayMap[l.day].total+hrs)*100)/100;
      dayMap[l.day].std=Math.min(dayMap[l.day].total,otThr);
      dayMap[l.day].ot=Math.max(0,Math.round((dayMap[l.day].total-otThr)*100)/100);
    });
    const totalStd=Object.values(dayMap).reduce((a,d)=>a+d.std,0);
    const totalOT=Object.values(dayMap).reduce((a,d)=>a+d.ot,0);
    const rate=worker.agreedRate||0;
    const otRate=worker.customOTRate||rate*(worker.overtimeMultiplier||1.5);
    const gross=Math.round((totalStd*rate+totalOT*otRate)*100)/100;
    const tax=Math.round(gross*(worker.taxRate||0)*100)/100;
    return {
      id:"ts_"+worker.id+"_"+wkLabel.replace(/\s+/g,""),
      workerId:worker.id, workerName:worker.name, position:worker.position, company:worker.company||"",
      weekLabel:wkLabel, workEntries:wkLogs, dayTotals:dayMap,
      stdHours:Math.round(totalStd*100)/100, otHours:Math.round(totalOT*100)/100,
      rate, otRate, gross, tax, net:Math.round((gross-tax)*100)/100, taxRate:worker.taxRate||0,
      status:"open", source:"gps", notes:"",
      createdAt:new Date().toISOString(), updatedAt:new Date().toISOString(),
      approvedAt:null, approvedBy:"",
    };
  };

  const generateTimesheets=(wkLabel)=>{
    const days=showWeekend?[...BASE_DAYS,...WEEKEND_DAYS]:BASE_DAYS;
    return workers.map(w=>{
      const gpsTs=buildTimesheetFromLogs(w,wkLabel||weekLabel);
      if(gpsTs) return gpsTs;
      const {stdH,otH,gross,net,tax}=calcPay(w,days,siteHours);
      return {
        id:"ts_"+w.id+"_"+(wkLabel||weekLabel).replace(/\s+/g,""),
        workerId:w.id, workerName:w.name, position:w.position,
        weekLabel:wkLabel||weekLabel, stdHours:stdH, otHours:otH,
        days:JSON.parse(JSON.stringify(w.days||{})),
        rate:w.agreedRate||0, gross, net, tax, taxRate:w.taxRate||0,
        status:"draft", source:"schedule", notes:"",
        lockedAt:null, approvedAt:null, approvedBy:"",
        createdAt:new Date().toISOString(),
      };
    }).filter(t=>t.stdHours>0||t.otHours>0);
  };
  // Generate payslips from timesheets
  const generatePayslips=(timesheets)=>{
    return timesheets.map(t=>({
      id:"ps_"+t.workerId+"_"+(t.weekLabel||weekLabel).replace(/\s+/g,""),
      workerId:t.workerId, workerName:t.workerName, position:t.position,
      company:t.company||"",
      weekLabel:t.weekLabel, stdHours:t.stdHours, otHours:t.otHours,
      rate:t.rate, gross:t.gross, net:t.net, tax:t.tax, taxRate:t.taxRate,
      dayBreakdown:t.dayBreakdown||{}, days:t.days||{},
      status:"pending", timesheetId:t.id,
      approvedAt:t.approvedAt||null, approvedBy:t.approvedBy||"",
      issuedAt:null, createdAt:new Date().toISOString(),
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

  if(!authState) return <LoginGate onLogin={setAuthState}/>;
  const {role}=authState;
  const handleLogout=()=>{setAuthState(null);};

  // Manager → Site Manager app only
  if(role==="manager") return <ManagerView authState={authState} onLogout={handleLogout}/>;

  // Worker → their own read-only view
  if(role==="worker") return <WorkerSelfView authState={authState} allSites={allSites} weekLabel={weekLabel} timesheetRecords={timesheetRecords} payslipRecords={payslipRecords} onLogout={handleLogout}/>;

  // Admin / Director / Accountant / Logistics → full app continues below
  if(loading) return <div style={{height:"100vh",width:"100vw",background:"#0d1117",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",gap:16,position:"fixed",top:0,left:0}}>
    <div style={{width:48,height:48,background:"linear-gradient(135deg,#3b82f6,#6366f1)",borderRadius:12,display:"flex",alignItems:"center",justifyContent:"center",fontSize:24}}>🏗</div>
    <div style={{color:"#60a5fa",fontSize:16,fontWeight:700}}>Loading…</div>
    <div style={{width:200,height:3,background:"#1e2535",borderRadius:3,overflow:"hidden"}}><div style={{height:"100%",background:"linear-gradient(90deg,#3b82f6,#6366f1)",borderRadius:3,animation:"slide 1.5s ease-in-out infinite"}}/></div>
    <style>{`@keyframes slide{0%{width:0%;margin-left:0%}50%{width:60%;margin-left:20%}100%{width:0%;margin-left:100%}}`}</style>
  </div>;

  return (
    <div style={{height:"100vh",width:"100vw",background:"#0d1117",fontFamily:"system-ui,'Segoe UI',sans-serif",color:"#e2e8f0",fontSize:13,display:"flex",flexDirection:"column",overflow:"hidden",position:"fixed",top:0,left:0}}>
      <style>{`*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}html,body{height:100%;width:100%;overflow:hidden}body{background:#0d1117}::-webkit-scrollbar{width:5px;height:5px}::-webkit-scrollbar-track{background:#0a0e17}::-webkit-scrollbar-thumb{background:#2d3555;border-radius:3px}::-webkit-scrollbar-thumb:hover{background:#3b82f6}`}</style>

      {/* ── TOP BAR: always visible, always exactly one ── */}
      <div style={{background:"linear-gradient(135deg,#0f172a,#1a1f2e)",borderBottom:"1px solid #1e2535",padding:"9px 16px",flexShrink:0,zIndex:300}}>
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
      <div style={{display:"flex",flex:1,minHeight:0,overflow:"hidden"}}>

        {/* THE ONLY SIDEBAR IN THE ENTIRE APP */}
        <DashSidebar
          page={dashPage}
          setPage={p=>{setDashPage(p);setDashDetailId(null);}}
          workers={workers} allSites={allSites} clients={clients}
          invoices={invoices} bankTransactions={bankTransactions} setModal={setModal}
          activeDays={activeDays} siteHours={siteHours} weekLabel={weekLabel}
          role={role} setAuthState={setAuthState}/>

        {/* Main content — DashboardView renders ONLY page content, never a sidebar */}
        <div style={{flex:1,overflowY:"auto",minWidth:0}}>
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
            setShowWeekend={setShowWeekend}
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
