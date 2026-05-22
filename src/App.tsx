import { useState, useEffect, useCallback, useRef, Component } from "react";

/* ─── SheetJS via CDN ────────────────────── */
function useXLSX() {
  const [ready, setReady] = useState(!!window.XLSX);
  useEffect(() => {
    if (window.XLSX) { setReady(true); return; }
    const s = document.createElement("script");
    s.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
    s.onload = () => setReady(true);
    document.head.appendChild(s);
  }, []);
  return ready;
}

/* ─── Google Drive upload ───────────────── */
const DRIVE_FOLDER       = "1xbxhkgi_8AhcelI2cGJVnJOWDTAR9fkw";
const DEFAULT_FIREBASE   = "https://gpd-archive-default-rtdb.firebaseio.com";
const DEFAULT_CLIENT_ID  = "926347832107-9u08816ppn9sgkmueae0i7mgbcbuuv5i.apps.googleusercontent.com";
// drive scope completo — garante acesso à pasta compartilhada
const DRIVE_SCOPE  = "https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/drive.appdata";

function loadGIS() {
  return new Promise((resolve, reject) => {
    if (window.google?.accounts?.oauth2) { resolve(); return; }
    const s = document.createElement("script");
    s.src = "https://accounts.google.com/gsi/client";
    s.onload = resolve;
    s.onerror = () => reject(new Error("Falha ao carregar GIS"));
    document.head.appendChild(s);
  });
}

function getGoogleToken(clientId) {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error("Timeout — feche o popup e tente novamente")), 90000);
    const client = window.google.accounts.oauth2.initTokenClient({
      client_id: clientId,
      scope: DRIVE_SCOPE,
      callback: resp => {
        clearTimeout(timer);
        if (resp.error) {
          const msg = resp.error === "access_denied"
            ? "Acesso negado — clique em Avançado > Acessar no popup do Google"
            : (resp.error_description || resp.error);
          reject(new Error(msg));
        } else {
          resolve(resp.access_token);
        }
      },
      error_callback: err => {
        clearTimeout(timer);
        reject(new Error(err?.message || "Falha ao abrir popup do Google"));
      },
    });
    client.requestAccessToken({ prompt: "consent" });
  });
}

async function fetchWithTimeout(url, opts, ms = 30000) {
  const ctrl = new AbortController();
  const id    = setTimeout(() => ctrl.abort(), ms);
  try {
    const r = await fetch(url, { ...opts, signal: ctrl.signal });
    clearTimeout(id);
    return r;
  } catch(e) {
    clearTimeout(id);
    throw e.name === "AbortError" ? new Error("Timeout: upload demorou mais de 30s") : e;
  }
}

async function uploadFileToDrive(token, file, gpdDate) {
  // Renomeia para "GPD 2026-05-21.xlsx" para fácil localização
  const ext      = file.name.split(".").pop();
  const safeName = gpdDate ? `GPD ${gpdDate}.${ext}` : file.name;
  const meta     = JSON.stringify({ name: safeName, parents: [DRIVE_FOLDER] });
  const form = new FormData();
  form.append("metadata", new Blob([meta], { type: "application/json" }));
  form.append("file", file);
  const r = await fetchWithTimeout(
    "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart",
    { method: "POST", headers: { Authorization: `Bearer ${token}` }, body: form }
  );
  if (!r.ok) {
    const txt = await r.text().catch(() => "");
    throw new Error(`Drive ${r.status}: ${txt.slice(0,120)}`);
  }
  return r.json();
}

/* ─── Firebase REST API (sem SDK) ───────── */
async function fbGet(url) {
  const r = await fetch(`${url}/gpd-archive.json`);
  if (!r.ok) throw new Error(`Firebase GET ${r.status}`);
  const d = await r.json();
  return Array.isArray(d) ? d : [];
}
async function fbSet(url, data) {
  const r = await fetch(`${url}/gpd-archive.json`, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(data),
  });
  if (!r.ok) throw new Error(`Firebase PUT ${r.status}`);
  return r.json();
}

/* ─── localStorage helper ────────────────── */
const ls = {
  get: k => { try { return localStorage.getItem(k); } catch { return null; } },
  set: (k, v) => { try { localStorage.setItem(k, v); } catch {} },
};

/* ─── Constants ──────────────────────────── */
const MONTHS_FULL = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"];
const MONTHS_SH   = ["JAN","FEV","MAR","ABR","MAI","JUN","JUL","AGO","SET","OUT","NOV","DEZ"];
const DAYS_SH     = ["Dom","Seg","Ter","Qua","Qui","Sex","Sáb"];
const DAYS_MIN    = ["D","S","T","Q","Q","S","S"];
const VIEWS       = { ANNUAL:"annual", MONTHLY:"monthly", WEEKLY:"weekly" };

/* ─── Date helpers ───────────────────────── */
const localDateStr = d => `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
const TODAY        = localDateStr(new Date());
const addDays      = (d,n) => { const r=new Date(d); r.setDate(r.getDate()+n); return r; };
const wkStart      = d => { const r=new Date(d); r.setDate(r.getDate()-r.getDay()); return r; };
const fmtMonth     = (y,m) => `${MONTHS_FULL[m].toUpperCase()} ${y}`;
const fmtWeek      = (ws,we) => {
  const sm=MONTHS_SH[ws.getMonth()],em=MONTHS_SH[we.getMonth()],sy=ws.getFullYear(),ey=we.getFullYear();
  if(sm===em&&sy===ey) return `${ws.getDate()}–${we.getDate()} ${sm} ${sy}`;
  if(sy===ey) return `${ws.getDate()} ${sm} – ${we.getDate()} ${em} ${sy}`;
  return `${ws.getDate()} ${sm} ${sy} – ${we.getDate()} ${em} ${ey}`;
};
const pct = (n=0,t=0) => t?Math.round((n/t)*100):0;

/* ─── Theme ──────────────────────────────── */
const T = light => ({
  bg:       light?"#f8fafc":"#030712",
  bgHdr:    light?"#ffffff":"#040a17",
  bgSide:   light?"#ffffff":"#040a17",
  bgCard:   light?"#ffffff":"rgba(255,255,255,0.02)",
  bgStat:   light?"#f1f5f9":"rgba(255,255,255,0.025)",
  bgAccLt:  light?"#e0f2fe":"rgba(0,212,255,0.07)",
  border:   light?"rgba(0,0,0,0.08)":"rgba(255,255,255,0.06)",
  borderAcc:light?"rgba(3,105,161,0.3)":"rgba(0,212,255,0.2)",
  txt:      light?"#0f172a":"#f1f5f9",
  txtSec:   light?"#1e293b":"#cbd5e1",   // mais forte
  txtMuted: light?"#475569":"#94a3b8",   // antes era quase invisível
  accent:   light?"#0369a1":"#00d4ff",
  accentInv:light?"#ffffff":"#030712",
  monthTxt: light?"#0369a1":"#00d4ff",
  tgBg:     light?"#0f172a":"rgba(255,255,255,0.06)",
  tgTxt:    light?"#ffffff":"#e2e8f0",
  scrollBar:light?"rgba(3,105,161,0.2)":"rgba(0,212,255,0.18)",
  blue:     light?"#1d4ed8":"#60a5fa",
  green:    light?"#15803d":"#4ade80",
  blueBg:   light?"#eff6ff":"#0f2744",
  greenBg:  light?"#f0fdf4":"#0a2c1a",
  red:      "#ef4444", yellow:"#f59e0b", greenVal:"#22c55e",
});

/* ─── Parser ─────────────────────────────── */
function parseGPD(wb, filename) {
  try {
    const sn  = wb.SheetNames.find(n=>n.includes("DASHBOARD"))||wb.SheetNames[0];
    const sh  = wb.Sheets[sn];
    if(!sh) return null;
    const rows= window.XLSX.utils.sheet_to_json(sh,{header:1,defval:null});

    let referencia="", geradoEm="", calendarDate="";
    for(const row of rows) {
      for(const cell of row) {
        const str = typeof cell==="string"?cell:String(cell??"");
        if(!str) continue;
        const rM = str.match(/Refer[êe]ncia[:\s]+([A-Za-zÀ-ÿ]{3}\/\d{2,4})/i);
        const dM = str.match(/Gerado\s+em[:\s]+(\d{2}\/\d{2}\/\d{4})/i);
        if(rM&&!referencia) referencia=rM[1];
        if(dM&&!calendarDate) {
          geradoEm=dM[1];
          const [d,m,y]=dM[1].split("/");
          calendarDate=`${y}-${m.padStart(2,"0")}-${d.padStart(2,"0")}`;
        }
      }
    }
    if(!calendarDate) {
      const fn=filename.match(/(\d{2})[-_](\d{2})[-_](\d{4})/);
      if(fn) calendarDate=`${fn[3]}-${fn[2].padStart(2,"0")}-${fn[1].padStart(2,"0")}`;
    }

    const blocks=[];
    for(const row of rows) {
      if(blocks.length>=2) break;
      const nums=row.filter(c=>typeof c==="number"&&!isNaN(c)&&c>0&&Number.isInteger(c));
      if(nums.length>=4) blocks.push(nums.slice(0,4));
    }
    const [totalMetas=0,comRealizado=0,semRealizado=0,emVermelho=0]=blocks[0]||[];
    const [emAmarelo=0,emVerde=0,vermelhoSemCM=0,amareloSemCM=0]=blocks[1]||[];

    const areas=[];
    let cap=false;
    for(const row of rows){
      if(!cap&&row.some(c=>typeof c==="string"&&c.includes("RESULTADO POR DIRETORIA"))){cap=true;continue;}
      if(cap&&typeof row[1]==="string"&&typeof row[2]==="number"){
        const rp=row[3];
        const p=typeof rp==="number"?(rp<=1?rp*100:rp):parseFloat(String(rp||0).replace("%",""))||0;
        areas.push({nome:row[1].trim(),total:row[2]||0,pct:Math.round(p),
          verde:row[4]||0,amarelo:row[5]||0,vermelho:row[6]||0,critico:row[7]||0,semReal:row[8]||0});
      }
    }
    return {
      id:`${Date.now()}-${Math.random().toString(36).slice(2,7)}`,
      filename,referencia,geradoEm,
      calendarDate:calendarDate||TODAY,
      uploadedAt:new Date().toISOString(),
      stats:{totalMetas,comRealizado,semRealizado,emVermelho,emAmarelo,emVerde,vermelhoSemCM,amareloSemCM},
      areas,
    };
  } catch(e){console.error("parseGPD:",e);return null;}
}

/* ─── Helpers ────────────────────────────── */
function healthColor(stats){
  const{emVermelho:r=0,emAmarelo:a=0,emVerde:g=0}=stats||{};
  const tot=r+a+g; if(!tot) return null;
  if(r/tot>0.3) return "#ef4444";
  if((r+a)/tot>0.4) return "#f59e0b";
  return "#22c55e";
}
function exportJSON(gpds){
  const a=document.createElement("a");
  a.href=URL.createObjectURL(new Blob([JSON.stringify(gpds,null,2)],{type:"application/json"}));
  a.download=`gpd-archive-${TODAY}.json`; a.click(); URL.revokeObjectURL(a.href);
}

/* ─── ErrorBoundary ──────────────────────── */
class ErrorBoundary extends Component {
  constructor(p){super(p);this.state={err:null};}
  static getDerivedStateFromError(e){return{err:e};}
  render(){
    if(this.state.err) return(
      <div style={{padding:40,textAlign:"center",background:"#030712",color:"#f1f5f9",minHeight:"100vh",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",fontFamily:"monospace"}}>
        <div style={{fontSize:36,marginBottom:12}}>⚠️</div>
        <div style={{color:"#ef4444",marginBottom:8}}>Algo deu errado</div>
        <div style={{fontSize:11,color:"#64748b",marginBottom:20,maxWidth:400}}>{String(this.state.err?.message||this.state.err)}</div>
        <button onClick={()=>this.setState({err:null})} style={{background:"#0f2744",color:"#60a5fa",border:"1px solid #3b82f644",borderRadius:8,padding:"8px 20px",cursor:"pointer",fontFamily:"monospace"}}>↺ Tentar novamente</button>
      </div>
    );
    return this.props.children;
  }
}

/* ─── StatCard ───────────────────────────── */
function StatCard({label,value,color,sub,t}){
  return(
    <div style={{background:t.bgStat,border:`1px solid ${t.border}`,borderRadius:10,padding:"12px 16px"}}>
      <div style={{fontSize:9,letterSpacing:"0.1em",color:t.txtMuted,marginBottom:8}}>{label}</div>
      <div style={{fontSize:28,fontWeight:700,color,lineHeight:1}}>{value??"—"}</div>
      {sub&&<div style={{fontSize:9,color:t.txtSec,marginTop:6}}>{sub}</div>}
    </div>
  );
}

/* ─── FarolBar ───────────────────────────── */
function FarolBar({stats,t}){
  const{emVermelho:r=0,emAmarelo:a=0,emVerde:g=0}=stats||{};
  const tot=r+a+g; if(!tot) return null;
  return(
    <div style={{marginBottom:20}}>
      <div style={{fontSize:10,color:t.txtMuted,letterSpacing:"0.1em",marginBottom:8}}>DISTRIBUIÇÃO DE FAROL</div>
      <div style={{display:"flex",height:6,borderRadius:3,overflow:"hidden",gap:1}}>
        <div style={{flex:r/tot,background:"#ef4444"}}/><div style={{flex:a/tot,background:"#f59e0b"}}/><div style={{flex:g/tot,background:"#22c55e"}}/>
      </div>
      <div style={{display:"flex",gap:20,marginTop:6,fontSize:10,color:t.txtSec}}>
        <span>🔴 {pct(r,tot)}%</span><span>🟡 {pct(a,tot)}%</span><span>🟢 {pct(g,tot)}%</span>
      </div>
    </div>
  );
}

/* ─── AreaTable ──────────────────────────── */
function AreaTable({areas,t}){
  if(!areas?.length) return null;
  return(
    <div>
      <div style={{fontSize:10,color:t.txtMuted,letterSpacing:"0.1em",marginBottom:10}}>RESULTADO POR DIRETORIA</div>
      <div style={{overflowX:"auto"}}>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
          <thead><tr style={{borderBottom:`1px solid ${t.border}`}}>
            {["Diretoria","Total","% Atual.","🟢","🟡","🔴","🚨","⏳"].map(h=>(
              <th key={h} style={{padding:"6px 10px",textAlign:h==="Diretoria"?"left":"center",fontSize:9,color:t.txtMuted,letterSpacing:"0.08em",fontWeight:600,whiteSpace:"nowrap"}}>{h}</th>
            ))}
          </tr></thead>
          <tbody>
            {areas.map((a,i)=>{
              const pc=a.pct>=80?"#22c55e":a.pct>=50?"#f59e0b":"#ef4444";
              const nm=a.nome.replace("DIRETORIA DE ","").replace("DIRETORIA ","").replace("GERENTE EXECUTIVA ","GE ");
              return(
                <tr key={i} style={{borderBottom:`1px solid ${t.border}`}}>
                  <td style={{padding:"7px 10px",color:t.txtSec,fontSize:10,maxWidth:180,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={a.nome}>{nm}</td>
                  <td style={{padding:"7px 10px",textAlign:"center",color:t.txt,fontWeight:700}}>{a.total}</td>
                  <td style={{padding:"7px 10px",textAlign:"center",color:pc,fontWeight:700}}>{a.pct}%</td>
                  <td style={{padding:"7px 10px",textAlign:"center",color:"#22c55e"}}>{a.verde}</td>
                  <td style={{padding:"7px 10px",textAlign:"center",color:"#f59e0b"}}>{a.amarelo}</td>
                  <td style={{padding:"7px 10px",textAlign:"center",color:"#ef4444"}}>{a.vermelho}</td>
                  <td style={{padding:"7px 10px",textAlign:"center",color:"#ef4444",fontWeight:700}}>{a.critico}</td>
                  <td style={{padding:"7px 10px",textAlign:"center",color:t.txtMuted}}>{a.semReal}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ─── DetailPanel ────────────────────────── */
function DetailPanel({gpd,onClose,onRemove,t}){
  if(!gpd) return null;
  return(
    <div style={{background:t.bgAccLt,border:`1px solid ${t.borderAcc}`,borderRadius:14,padding:26,marginTop:24}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:22}}>
        <div>
          <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:24,letterSpacing:"0.1em",color:t.txt}}>GPD — {gpd.referencia}</div>
          <div style={{fontSize:10,color:t.txtSec,marginTop:5}}>Gerado em {gpd.geradoEm} · {gpd.filename}</div>
        </div>
        <button onClick={onClose} style={{background:"none",border:"none",color:t.txtMuted,fontSize:22,cursor:"pointer",padding:"0 4px",lineHeight:1}}>✕</button>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10,marginBottom:22}}>
        <StatCard t={t} label="TOTAL METAS" value={gpd.stats?.totalMetas} color={t.txt}/>
        <StatCard t={t} label="COM REALIZADO" value={gpd.stats?.comRealizado} color={t.blue} sub={`${pct(gpd.stats?.comRealizado,gpd.stats?.totalMetas)}% atualizado`}/>
        <StatCard t={t} label="SEM REALIZADO" value={gpd.stats?.semRealizado} color={t.txtSec}/>
        <StatCard t={t} label="EM VERMELHO" value={gpd.stats?.emVermelho} color="#ef4444" sub={`${pct(gpd.stats?.emVermelho,gpd.stats?.totalMetas)}% do total`}/>
        <StatCard t={t} label="EM AMARELO" value={gpd.stats?.emAmarelo} color="#f59e0b"/>
        <StatCard t={t} label="EM VERDE" value={gpd.stats?.emVerde} color="#22c55e" sub={`${pct(gpd.stats?.emVerde,gpd.stats?.comRealizado)}% no alvo`}/>
        <StatCard t={t} label="🚨 VERM. s/CM" value={gpd.stats?.vermelhoSemCM} color="#ef4444"/>
        <StatCard t={t} label="⚡ AMAR. s/CM" value={gpd.stats?.amareloSemCM} color="#f59e0b"/>
      </div>
      <FarolBar stats={gpd.stats} t={t}/>
      <AreaTable areas={gpd.areas} t={t}/>
      <div style={{marginTop:22,paddingTop:16,borderTop:`1px solid ${t.border}`,display:"flex",justifyContent:"flex-end"}}>
        <button onClick={onRemove} style={{background:"rgba(239,68,68,0.08)",color:"#ef4444",border:"1px solid rgba(239,68,68,0.3)",borderRadius:7,padding:"6px 16px",fontSize:11,cursor:"pointer",fontFamily:"'IBM Plex Mono',monospace"}}>✕ Remover este GPD</button>
      </div>
    </div>
  );
}

/* ─── Calendar cells ─────────────────────── */
function CalDay({ds,gpd,selected,isToday,onSelect,size="sm",t}){
  const hc=gpd?healthColor(gpd.stats):null;
  const sel=selected===ds;
  if(size==="sm") return(
    <div onClick={()=>gpd&&onSelect(ds)} title={gpd?`GPD ${gpd.referencia}`:""}
      style={{height:22,display:"flex",alignItems:"center",justifyContent:"center",borderRadius:3,fontSize:9,cursor:gpd?"pointer":"default",position:"relative",
        background:sel?t.accent:gpd?`${hc}18`:isToday?`${t.accent}10`:"transparent",
        border:sel?`1px solid ${t.accent}`:gpd?`1px solid ${hc}55`:isToday?`1px solid ${t.accent}30`:"none",
        color:sel?t.accentInv:gpd?hc:isToday?t.accent:t.txtMuted,fontWeight:gpd||isToday?700:400,transition:"all 0.1s"}}>
      {ds.slice(8)}
      {gpd&&!sel&&<span style={{position:"absolute",bottom:1,left:"50%",transform:"translateX(-50%)",width:3,height:3,borderRadius:"50%",background:hc}}/>}
    </div>
  );
  return(
    <div onClick={()=>gpd&&onSelect(ds)}
      style={{minHeight:90,borderRadius:10,padding:"8px 10px",cursor:gpd?"pointer":"default",
        border:sel?`2px solid ${t.accent}`:isToday?`1.5px solid ${t.accent}55`:`1px solid ${t.border}`,
        background:sel?t.bgAccLt:gpd?`${hc}0e`:t.bgCard,transition:"all 0.12s"}}>
      <div style={{fontSize:14,fontWeight:isToday||gpd?700:400,color:isToday?t.accent:sel?t.accent:t.txt,marginBottom:5,display:"flex",alignItems:"center",gap:4}}>
        {parseInt(ds.slice(8))}
        {isToday&&<span style={{fontSize:8,background:t.accent,color:t.accentInv,borderRadius:4,padding:"1px 5px"}}>HOJE</span>}
      </div>
      {gpd&&(
        <>
          <div style={{fontSize:10,fontWeight:700,color:hc,marginBottom:3}}>{gpd.referencia}</div>
          <div style={{fontSize:9,display:"flex",gap:4,flexWrap:"wrap",marginBottom:4}}>
            <span style={{color:"#ef4444"}}>🔴{gpd.stats?.emVermelho||0}</span>
            <span style={{color:"#f59e0b"}}>🟡{gpd.stats?.emAmarelo||0}</span>
            <span style={{color:"#22c55e"}}>🟢{gpd.stats?.emVerde||0}</span>
          </div>
          {(()=>{const{emVermelho:r=0,emAmarelo:a=0,emVerde:g=0}=gpd.stats||{};const tot=r+a+g;return tot?(<div style={{display:"flex",height:3,borderRadius:2,overflow:"hidden"}}><div style={{flex:r/tot,background:"#ef4444"}}/><div style={{flex:a/tot,background:"#f59e0b"}}/><div style={{flex:g/tot,background:"#22c55e"}}/></div>):null;})()}
        </>
      )}
    </div>
  );
}

function AnnualView({year,byDate,selected,onSelect,t}){
  return(
    <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12}}>
      {Array.from({length:12},(_,m)=>{
        const days=new Date(year,m+1,0).getDate(),fd=new Date(year,m,1).getDay();
        return(
          <div key={m} style={{background:t.bgCard,border:`1px solid ${t.border}`,borderRadius:10,padding:"12px 14px"}}>
            <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:14,letterSpacing:"0.18em",color:t.monthTxt,marginBottom:10}}>{MONTHS_SH[m]}</div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:1}}>
              {DAYS_MIN.map((d,i)=><div key={i} style={{height:17,display:"flex",alignItems:"center",justifyContent:"center",fontSize:8,color:t.txtMuted,fontWeight:600}}>{d}</div>)}
              {Array.from({length:fd},(_,i)=><div key={`e${i}`}/>)}
              {Array.from({length:days},(_,idx)=>{const d=idx+1;const ds=`${year}-${String(m+1).padStart(2,"0")}-${String(d).padStart(2,"0")}`;return <CalDay key={d} ds={ds} gpd={byDate[ds]} selected={selected} isToday={ds===TODAY} onSelect={onSelect} size="sm" t={t}/>;})}
            </div>
          </div>
        );
      })}
    </div>
  );
}

function MonthlyView({year,month,byDate,selected,onSelect,t}){
  const days=new Date(year,month+1,0).getDate(),fd=new Date(year,month,1).getDay();
  const cells=[...Array(fd).fill(null),...Array.from({length:days},(_,i)=>i+1)];
  return(
    <div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:2,marginBottom:4}}>
        {DAYS_SH.map(d=><div key={d} style={{textAlign:"center",padding:"8px 0",fontSize:11,color:t.txtMuted,fontWeight:600,letterSpacing:"0.05em"}}>{d}</div>)}
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:6}}>
        {cells.map((d,i)=>{if(!d) return <div key={`e${i}`} style={{minHeight:90}}/>;const ds=`${year}-${String(month+1).padStart(2,"0")}-${String(d).padStart(2,"0")}`;return <CalDay key={d} ds={ds} gpd={byDate[ds]} selected={selected} isToday={ds===TODAY} onSelect={onSelect} size="md" t={t}/>;})}
      </div>
    </div>
  );
}

function WeeklyView({navDate,byDate,selected,onSelect,t}){
  const ws=wkStart(navDate);
  const days=Array.from({length:7},(_,i)=>addDays(ws,i));
  return(
    <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gridAutoRows:"1fr",gap:8}}>
      {days.map((day,i)=>{
        const ds=localDateStr(day),gpd=byDate[ds],hc=gpd?healthColor(gpd.stats):null,sel=selected===ds,isTd=ds===TODAY;
        return(
          <div key={i} onClick={()=>gpd&&onSelect(ds)}
            style={{borderRadius:12,padding:14,cursor:gpd?"pointer":"default",display:"flex",flexDirection:"column",
              border:sel?`2px solid ${t.accent}`:isTd?`1.5px solid ${t.accent}55`:`1px solid ${t.border}`,
              background:sel?t.bgAccLt:gpd?`${hc}0a`:t.bgCard,transition:"all 0.12s"}}>
            {/* Header com altura fixa para alinhar todos os dias */}
            <div style={{height:72,marginBottom:12,paddingBottom:10,borderBottom:`1px solid ${t.border}`,flexShrink:0}}>
              <div style={{fontSize:10,color:isTd?t.accent:t.txtMuted,fontWeight:700,letterSpacing:"0.06em",textTransform:"uppercase"}}>{DAYS_SH[i]}</div>
              <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:36,letterSpacing:"0.04em",color:isTd?t.accent:sel?t.accent:t.txt,lineHeight:1}}>{day.getDate()}</div>
              <div style={{fontSize:10,color:t.txtSec,fontWeight:600}}>{MONTHS_SH[day.getMonth()]} {day.getFullYear()}</div>
            </div>
            {gpd?(
              <div>
                <div style={{fontSize:11,fontWeight:700,color:hc,background:`${hc}18`,borderRadius:6,padding:"3px 8px",display:"inline-block",marginBottom:10}}>{gpd.referencia}</div>
                {(()=>{const{emVermelho:r=0,emAmarelo:a=0,emVerde:g=0}=gpd.stats||{};const tot=r+a+g;return tot?(<div style={{marginBottom:10}}><div style={{display:"flex",height:5,borderRadius:3,overflow:"hidden",gap:1}}><div style={{flex:r/tot,background:"#ef4444"}}/><div style={{flex:a/tot,background:"#f59e0b"}}/><div style={{flex:g/tot,background:"#22c55e"}}/></div><div style={{display:"flex",justifyContent:"space-between",marginTop:4,fontSize:9,color:t.txtSec}}><span>🔴{pct(r,tot)}%</span><span>🟡{pct(a,tot)}%</span><span>🟢{pct(g,tot)}%</span></div></div>):null;})()}
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:4}}>
                  {[
                    {l:"Total",   v:gpd.stats?.totalMetas,      c:t.txt},
                    {l:"C Real.", v:gpd.stats?.comRealizado,    c:t.blue},
                    {l:"S Real.", v:gpd.stats?.semRealizado,    c:t.txtSec},
                    {l:"Vermelho",v:gpd.stats?.emVermelho,      c:"#ef4444"},
                    {l:"Amarelo", v:gpd.stats?.emAmarelo,       c:"#f59e0b"},
                    {l:"Verde",   v:gpd.stats?.emVerde,         c:"#22c55e"},
                  ].map(({l,v,c})=>(
                    <div key={l} style={{background:t.bgStat,borderRadius:6,padding:"6px 7px"}}>
                      <div style={{fontSize:8,color:t.txtSec,marginBottom:3,letterSpacing:"0.03em",fontWeight:600}}>{l}</div>
                      <div style={{fontSize:17,fontWeight:700,color:c,lineHeight:1}}>{v??0}</div>
                    </div>
                  ))}
                </div>
              </div>
            ):(
              <div style={{display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",flex:1,paddingTop:40,color:t.txtMuted}}>
                <div style={{fontSize:24,opacity:0.2,marginBottom:8}}>—</div>
                <div style={{fontSize:10,letterSpacing:"0.06em",fontWeight:600}}>sem GPD</div>
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}

/* ─── Setup Screen ───────────────────────── */
function SetupScreen({onSave,t}){
  const [url,setUrl]=useState("");
  return(
    <div style={{minHeight:"100vh",background:t.bg,display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"'IBM Plex Mono',monospace",padding:20}}>
      <div style={{width:"100%",maxWidth:480,background:t.bgHdr,border:`1px solid ${t.borderAcc}`,borderRadius:16,padding:36}}>
        <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:28,letterSpacing:"0.14em",color:t.accent,marginBottom:6}}>GPD ARCHIVE</div>
        <div style={{fontSize:11,color:t.txtMuted,marginBottom:28,lineHeight:1.7}}>Configure o banco de dados para sincronizar com todos os usuários.</div>

        <div style={{background:t.bgStat,borderRadius:10,padding:20,marginBottom:20,border:`1px solid ${t.border}`}}>
          <div style={{fontSize:10,fontWeight:700,color:t.txt,marginBottom:8,letterSpacing:"0.07em"}}>🔥 FIREBASE — CONFIGURAÇÃO (grátis, 3 min)</div>
          <ol style={{fontSize:11,color:t.txtSec,lineHeight:2,paddingLeft:18,marginBottom:16}}>
            <li>Acesse <a href="https://console.firebase.google.com" target="_blank" style={{color:t.accent}}>console.firebase.google.com</a></li>
            <li>Criar projeto → nome qualquer → continuar</li>
            <li>Menu lateral: <b style={{color:t.txt}}>Build → Realtime Database</b></li>
            <li>Criar banco → <b style={{color:t.txt}}>Modo de teste</b> → Concluir</li>
            <li>Copie a URL do banco (ex: <code style={{fontSize:9,color:t.accent}}>https://SEU-PROJETO-default-rtdb.firebaseio.com</code>)</li>
          </ol>
          <input value={url} onChange={e=>setUrl(e.target.value)}
            placeholder="https://SEU-PROJETO-default-rtdb.firebaseio.com"
            style={{width:"100%",background:t.bg,color:t.txt,border:`1px solid ${t.borderAcc}`,borderRadius:8,padding:"10px 12px",fontSize:11,fontFamily:"'IBM Plex Mono',monospace",outline:"none",marginBottom:12,boxSizing:"border-box"}}/>
          <button onClick={()=>{if(url.trim()){ls.set("gpd-firebase-url",url.trim());onSave(url.trim());}}}
            disabled={!url.trim()}
            style={{width:"100%",background:url.trim()?t.accent:"#1e293b",color:url.trim()?t.accentInv:t.txtMuted,border:"none",borderRadius:8,padding:"12px 0",fontSize:13,fontWeight:700,cursor:url.trim()?"pointer":"not-allowed",fontFamily:"'IBM Plex Mono',monospace",letterSpacing:"0.06em",transition:"all 0.2s"}}>
            CONECTAR E ABRIR →
          </button>
        </div>

        <div style={{fontSize:10,color:t.txtMuted,textAlign:"center",lineHeight:1.7}}>
          O Firebase gratuito suporta até 1GB de dados e 100 conexões simultâneas.<br/>Mais que suficiente para uso interno.
        </div>
      </div>
    </div>
  );
}

/* ─── Main App ───────────────────────────── */
function GPDArchiveInner(){
  const [firebaseUrl, setFirebaseUrl] = useState(()=>{
    const params = new URLSearchParams(window.location.search);
    const fbParam = params.get("fb");
    if(fbParam){ ls.set("gpd-firebase-url", fbParam); return fbParam; }
    const saved = ls.get("gpd-firebase-url");
    if(saved) return saved;
    ls.set("gpd-firebase-url", DEFAULT_FIREBASE);
    return DEFAULT_FIREBASE;
  });
  const [light,    setLight]    = useState(false);
  const [gpds,     setGpds]     = useState([]);
  const [view,     setView]     = useState(VIEWS.ANNUAL);
  const [navDate,  setNavDate]  = useState(new Date());
  const [selected, setSelected] = useState(null);
  const [dragging,    setDragging]    = useState(false);
  const [showConfig,  setShowConfig]  = useState(false);
  const [syncStatus,setSyncStatus]=useState("idle"); // idle | syncing | ok | err
  const [syncMsg,  setSyncMsg]  = useState("");
  const [uploading,  setUploading]   = useState(false);
  const [uploadErr,  setUploadErr]   = useState(null);
  const [driveStatus,setDriveStatus] = useState(null);
  const [gToken,     setGToken]      = useState(null);
  const [clientId,   setClientId]    = useState(()=>{
    const saved = ls.get("gpd-client-id");
    if(saved) return saved;
    ls.set("gpd-client-id", DEFAULT_CLIENT_ID);
    return DEFAULT_CLIENT_ID;
  });
  const clientIdRef  = useRef(ls.get("gpd-client-id")||DEFAULT_CLIENT_ID);
  const gTokenRef    = useRef(null);
  const xlsxReady = useXLSX();
  const t = T(light);

  // Mantém clientId ref sincronizado (gTokenRef é atualizado diretamente)
  useEffect(()=>{ clientIdRef.current = clientId; }, [clientId]);

  /* ── Sync from Firebase ── */
  const syncFromFirebase = useCallback(async (url=firebaseUrl) => {
    if(!url) return;
    try {
      const data = await fbGet(url);
      setGpds(data);
    } catch(e){ console.error("sync:",e); }
  },[firebaseUrl]);

  /* ── Save to Firebase ── */
  const saveToFirebase = useCallback(async (list, url=firebaseUrl) => {
    if(!url) return;
    try { await fbSet(url, list); } catch(e){ console.error("save:",e); }
  },[firebaseUrl]);

  /* ── Initial load + polling ── */
  useEffect(()=>{
    if(!firebaseUrl) return;
    syncFromFirebase();
    const iv = setInterval(syncFromFirebase, 8000);
    return ()=>clearInterval(iv);
  },[firebaseUrl,syncFromFirebase]);

  /* ── Upload handler ── */
  const addGPD = useCallback(file=>{
    if(!window.XLSX){ setUploadErr("Aguarde o leitor XLSX carregar."); return; }
    setUploading(true); setUploadErr(null);
    const reader=new FileReader();
    reader.onerror=()=>{ setUploading(false); setUploadErr("Erro ao ler o arquivo."); };
    reader.onload=e=>{
      setTimeout(()=>{
        try{
          const wb=window.XLSX.read(new Uint8Array(e.target.result),{type:"array"});
          const parsed=parseGPD(wb,file.name);
          if(!parsed){ setUploading(false); setUploadErr("Não foi possível ler o GPD."); return; }
          setGpds(prev=>{
            const next=[parsed,...prev.filter(g=>g.calendarDate!==parsed.calendarDate)];
            saveToFirebase(next);
            return next;
          });
          setNavDate(new Date(parsed.calendarDate+"T12:00:00"));
          setSelected(parsed.calendarDate);
          setUploading(false);
          // Upload do arquivo original para o Google Drive (em background)
          // Upload para o Drive com nome "GPD YYYY-MM-DD.xlsx"
          uploadToGDrive(file, parsed.calendarDate);
        } catch(err){
          console.error(err); setUploading(false); setUploadErr("Erro: "+err?.message);
        }
      },0);
    };
    reader.readAsArrayBuffer(file);
  },[saveToFirebase]);

  /* ── Google Drive upload (background) ── */
  const uploadToGDrive = useCallback(async (file, fileDate) => {
    const cid = clientIdRef.current;
    if(!cid) return;
    try {
      await loadGIS();
      // Reutiliza token existente; se expirado (401) pede novo
      let token = gTokenRef.current;
      if(!token) {
        setDriveStatus("auth");
        token = await getGoogleToken(cid);
        gTokenRef.current = token;  // atualiza imediatamente
        setGToken(token);
      }
      setDriveStatus("uploading");
      let res;
      try {
        res = await uploadFileToDrive(token, file, fileDate);
      } catch(uploadErr) {
        // Se 401: token expirado → pede novo e tenta de novo
        if(String(uploadErr.message).includes("401")) {
          gTokenRef.current = null; setGToken(null);
          setDriveStatus("auth");
          token = await getGoogleToken(cid);
          gTokenRef.current = token; setGToken(token);
          setDriveStatus("uploading");
          res = await uploadFileToDrive(token, file, fileDate);
        } else { throw uploadErr; }
      }
      console.log("Drive upload OK:", res?.id);
      setDriveStatus("ok");
      setTimeout(()=>setDriveStatus(null), 5000);
    } catch(e) {
      console.error("Drive upload error:", e);
      const msg = String(e?.message || "Erro desconhecido");
      if(msg.includes("401")||msg.includes("403")||msg.includes("invalid_token")) {
        gTokenRef.current = null; setGToken(null);
      }
      setDriveStatus("err:" + msg.slice(0, 100));
      setTimeout(()=>setDriveStatus(null), 15000);
    }
  }, []);

  /* ── Navigation ── */
  const navigate=delta=>setNavDate(d=>{
    const r=new Date(d);
    if(view===VIEWS.ANNUAL) r.setFullYear(r.getFullYear()+delta);
    if(view===VIEWS.MONTHLY) r.setMonth(r.getMonth()+delta);
    if(view===VIEWS.WEEKLY) r.setDate(r.getDate()+delta*7);
    return r;
  });

  const Btn=(bg,color)=>({background:bg,color,border:`1px solid ${color}44`,borderRadius:7,padding:"5px 14px",fontSize:11,cursor:"pointer",fontFamily:"'IBM Plex Mono',monospace",letterSpacing:"0.05em",transition:"opacity 0.15s"});

  // Config panel overlay
  const ConfigPanel = showConfig ? (
    <div style={{position:"fixed",inset:0,zIndex:1000,display:"flex",alignItems:"flex-start",justifyContent:"flex-end",padding:"60px 20px 0"}}>
      {/* backdrop */}
      <div onClick={()=>setShowConfig(false)} style={{position:"absolute",inset:0,background:"rgba(0,0,0,0.4)"}}/>
      <div style={{position:"relative",width:400,background:t.bgHdr,border:`1px solid ${t.borderAcc}`,borderRadius:14,padding:24,boxShadow:"0 8px 40px rgba(0,0,0,0.5)",maxHeight:"85vh",overflowY:"auto"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20}}>
          <span style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:18,letterSpacing:"0.12em",color:t.accent}}>CONFIGURAÇÕES</span>
          <button onClick={()=>setShowConfig(false)} style={{background:"none",border:"none",color:t.txtMuted,fontSize:20,cursor:"pointer",lineHeight:1}}>✕</button>
        </div>

        {/* Firebase URL */}
        <div style={{background:t.bgStat,borderRadius:10,padding:16,marginBottom:14,border:`1px solid ${t.border}`}}>
          <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
            <div style={{fontSize:10,fontWeight:700,color:t.txt,letterSpacing:"0.07em"}}>🔥 FIREBASE DATABASE URL</div>
            {firebaseUrl===DEFAULT_FIREBASE&&<span style={{fontSize:8,background:"rgba(34,197,94,0.12)",color:"#22c55e",border:"1px solid rgba(34,197,94,0.3)",borderRadius:4,padding:"1px 6px",letterSpacing:"0.06em"}}>PADRÃO</span>}
          </div>
          <input id="cfg-fb" defaultValue={firebaseUrl}
            placeholder="https://projeto-default-rtdb.firebaseio.com"
            style={{width:"100%",background:t.bg,color:t.txt,border:`1px solid ${t.borderAcc}`,borderRadius:7,padding:"8px 10px",fontSize:10,fontFamily:"'IBM Plex Mono',monospace",outline:"none",marginBottom:10,boxSizing:"border-box"}}/>
          <button className="ba" onClick={()=>{
            const v=document.getElementById("cfg-fb")?.value?.trim();
            if(v){ls.set("gpd-firebase-url",v);setFirebaseUrl(v);}
          }} style={{...Btn(t.greenBg,t.green),width:"100%",textAlign:"center",padding:"8px 0"}}>Salvar URL</button>
          <div style={{marginTop:12,padding:"10px 12px",background:`${t.accent}0d`,borderRadius:8,border:`1px solid ${t.accent}22`}}>
            <div style={{fontSize:9,color:t.accent,fontWeight:700,marginBottom:6,letterSpacing:"0.06em"}}>🔗 LINK PARA COMPARTILHAR</div>
            <div style={{fontSize:10,color:t.txtSec,marginBottom:8,lineHeight:1.7}}>Envie esse link para sua dupla — ela abre e já está configurado:</div>
            <div style={{background:t.bg,borderRadius:6,padding:"6px 10px",fontSize:9,color:t.txtMuted,wordBreak:"break-all",letterSpacing:"0.02em",marginBottom:8,fontFamily:"monospace"}}>
              {window.location.origin}/?fb={encodeURIComponent(firebaseUrl)}
            </div>
            <button className="ba" onClick={()=>{
              navigator.clipboard.writeText(`${window.location.origin}/?fb=${encodeURIComponent(firebaseUrl)}`);
            }} style={{...Btn(t.bgStat,t.accent),width:"100%",textAlign:"center",padding:"7px 0",fontSize:10}}>📋 Copiar link</button>
          </div>
        </div>

        {/* Google Client ID */}
        <div style={{background:t.bgStat,borderRadius:10,padding:16,marginBottom:14,border:`1px solid ${t.border}`}}>
          <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:6}}>
            <div style={{fontSize:10,fontWeight:700,color:t.txt,letterSpacing:"0.07em"}}>☁ GOOGLE CLIENT ID (backup Drive)</div>
            {clientId===DEFAULT_CLIENT_ID&&<span style={{fontSize:8,background:"rgba(34,197,94,0.12)",color:"#22c55e",border:"1px solid rgba(34,197,94,0.3)",borderRadius:4,padding:"1px 6px",letterSpacing:"0.06em"}}>PADRÃO</span>}
          </div>
          <div style={{fontSize:10,color:t.txtSec,marginBottom:8,lineHeight:1.6}}>
            Cada usuário usa o seu próprio login Google — basta ter o mesmo Client ID configurado.
          </div>
          <input id="cfg-cid" defaultValue={clientId}
            placeholder="926...apps.googleusercontent.com"
            style={{width:"100%",background:t.bg,color:t.txt,border:`1px solid ${clientId?"#22c55e44":t.borderAcc}`,borderRadius:7,padding:"8px 10px",fontSize:10,fontFamily:"'IBM Plex Mono',monospace",outline:"none",marginBottom:10,boxSizing:"border-box"}}/>
          <button className="ba" onClick={()=>{
            const v=document.getElementById("cfg-cid")?.value?.trim();
            if(v){setClientId(v);ls.set("gpd-client-id",v);clientIdRef.current=v;}
          }} style={{...Btn(t.greenBg,t.green),width:"100%",textAlign:"center",padding:"8px 0"}}>Salvar Client ID</button>
          {clientId && <div style={{marginTop:8,fontSize:10,color:"#22c55e"}}>✓ Configurado — backup automático ativo</div>}
        </div>

        {/* Export JSON */}
        <div style={{background:t.bgStat,borderRadius:10,padding:16,marginBottom:14,border:`1px solid ${t.border}`}}>
          <div style={{fontSize:10,fontWeight:700,color:t.txt,marginBottom:6,letterSpacing:"0.07em"}}>📄 EXPORTAR BACKUP JSON</div>
          <div style={{fontSize:10,color:t.txtSec,marginBottom:10,lineHeight:1.6}}>Baixa todos os GPDs do histórico como arquivo .json.</div>
          <button className="ba" onClick={()=>exportJSON(gpds)} disabled={!gpds.length}
            style={{...Btn(t.blueBg,t.blue),width:"100%",textAlign:"center",padding:"8px 0",opacity:gpds.length?1:0.4}}>
            ↓ Exportar JSON ({gpds.length} GPD{gpds.length!==1?"s":""})
          </button>
        </div>

        {/* Danger zone */}
        <div style={{background:"rgba(239,68,68,0.05)",borderRadius:10,padding:16,border:"1px solid rgba(239,68,68,0.2)"}}>
          <div style={{fontSize:10,fontWeight:700,color:"#f87171",marginBottom:6,letterSpacing:"0.07em"}}>⚠ ZONA DE RISCO</div>
          <button className="ba" onClick={()=>{
            if(window.confirm("Redefinir para as configurações padrão?")){
              ls.set("gpd-firebase-url",DEFAULT_FIREBASE);setFirebaseUrl(DEFAULT_FIREBASE);
              ls.set("gpd-client-id",DEFAULT_CLIENT_ID);setClientId(DEFAULT_CLIENT_ID);clientIdRef.current=DEFAULT_CLIENT_ID;
              setGpds([]);setShowConfig(false);
            }
          }} style={{...Btn("rgba(239,68,68,0.08)","#f87171"),width:"100%",textAlign:"center",padding:"8px 0",fontSize:10}}>
            Redefinir para configurações padrão
          </button>
        </div>
      </div>
    </div>
  ) : null;

  const navYear=navDate.getFullYear(),navMonth=navDate.getMonth();
  const ws=wkStart(navDate),we=addDays(ws,6);
  const byDate=Object.fromEntries(gpds.map(g=>[g.calendarDate,g]));
  const sorted=[...gpds].sort((a,b)=>b.calendarDate.localeCompare(a.calendarDate));
  const selGpd=selected?gpds.find(g=>g.calendarDate===selected):null;
  const navLabel=view===VIEWS.ANNUAL?`${navYear}`:view===VIEWS.MONTHLY?fmtMonth(navYear,navMonth):fmtWeek(ws,we);
  const gpdsInPeriod=view===VIEWS.ANNUAL?gpds.filter(g=>g.calendarDate?.startsWith(`${navYear}`)).length:view===VIEWS.MONTHLY?gpds.filter(g=>g.calendarDate?.startsWith(`${navYear}-${String(navMonth+1).padStart(2,"0")}`)).length:gpds.filter(g=>g.calendarDate>=localDateStr(ws)&&g.calendarDate<=localDateStr(we)).length;

  return(
    <>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@300;400;600;700&family=Bebas+Neue&display=swap');
        *,*::before,*::after{box-sizing:border-box;margin:0;padding:0;}
        ::-webkit-scrollbar{width:4px;height:4px;}
        ::-webkit-scrollbar-track{background:transparent;}
        ::-webkit-scrollbar-thumb{background:${t.scrollBar};border-radius:2px;}
        .ba{transition:opacity 0.15s,transform 0.1s;cursor:pointer;}
        .ba:hover{opacity:0.8;}
        .ba:active{transform:scale(0.97);}
        .gi:hover{border-color:${t.accent}66 !important;background:${t.bgAccLt} !important;}
        @keyframes pulse{0%,100%{opacity:1}50%{opacity:0.4}}
      `}</style>
      <div style={{fontFamily:"'IBM Plex Mono',monospace",background:t.bg,minHeight:"100vh",color:t.txt,display:"flex",flexDirection:"column",fontSize:13}}>

        {/* HEADER */}
        <header style={{height:52,flexShrink:0,borderBottom:`1px solid ${t.border}`,display:"flex",alignItems:"center",padding:"0 20px",gap:14,background:t.bgHdr}}>
          <span style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:22,letterSpacing:"0.16em",color:t.accent,whiteSpace:"nowrap"}}>GPD ARCHIVE</span>
          <span style={{fontSize:9,color:t.txtMuted,letterSpacing:"0.1em",whiteSpace:"nowrap",display:"none"}}>GESTÃO DE PERFORMANCE DURADOURA</span>
          <div style={{display:"flex",alignItems:"center",gap:5,padding:"3px 10px",background:"rgba(34,197,94,0.08)",borderRadius:20,border:"1px solid rgba(34,197,94,0.2)"}}>
            <span style={{width:6,height:6,borderRadius:"50%",background:"#22c55e",display:"inline-block",animation:"pulse 2s infinite"}}/>
            <span style={{fontSize:9,color:"#22c55e",letterSpacing:"0.06em"}}>TEMPO REAL</span>
          </div>
          <div style={{flex:1}}/>
          <button className="ba" onClick={()=>syncFromFirebase()} style={{...Btn(t.bgStat,t.accent),fontSize:10}}>⟳ Sincronizar</button>
          {/* Drive status */}
          {driveStatus && (
            <span style={{fontSize:10,letterSpacing:"0.04em",maxWidth:320,
              color:driveStatus==="ok"?"#22c55e":driveStatus==="err"?"#ef4444":"#f59e0b"}}>
              {driveStatus==="auth"?"⟳ Autenticando Google...":driveStatus==="uploading"?"⟳ Salvando no Drive...":driveStatus==="ok"?"✓ Salvo no Drive":driveStatus?.startsWith("err:")?`✗ ${driveStatus.slice(4)}`:"✗ Erro no Drive"}
            </span>
          )}
          <button className="ba" onClick={()=>setShowConfig(c=>!c)} style={{...Btn(t.bgStat,t.txtSec),fontSize:10}}>⚙ Config</button>
        </header>

        <div style={{display:"flex",flex:1,overflow:"hidden",minHeight:0}}>

          {/* SIDEBAR */}
          <aside style={{width:252,flexShrink:0,borderRight:`1px solid ${t.border}`,display:"flex",flexDirection:"column",background:t.bgSide}}>

            {/* Theme */}
            <div style={{padding:"12px 14px 0"}}>
              <button className="ba" onClick={()=>setLight(l=>!l)} style={{width:"100%",background:t.tgBg,color:t.tgTxt,border:`1px solid ${t.border}`,borderRadius:8,padding:"8px 12px",fontSize:11,cursor:"pointer",fontFamily:"'IBM Plex Mono',monospace",letterSpacing:"0.05em",display:"flex",alignItems:"center",justifyContent:"center",gap:8}}>
                <span style={{display:"inline-block",width:28,height:16,borderRadius:8,flexShrink:0,background:light?t.accent:"rgba(255,255,255,0.12)",position:"relative",transition:"background 0.2s"}}>
                  <span style={{position:"absolute",top:2,left:light?"calc(100% - 14px)":"2px",width:12,height:12,borderRadius:"50%",background:"#fff",transition:"left 0.2s",boxShadow:"0 1px 3px rgba(0,0,0,0.3)"}}/>
                </span>
                {light?"🌙 Modo Escuro":"☀️ Modo Claro"}
              </button>
            </div>

            {/* Upload */}
            <div onDragOver={e=>{e.preventDefault();e.stopPropagation();setDragging(true);}}
              onDragLeave={()=>setDragging(false)}
              onDrop={e=>{e.preventDefault();e.stopPropagation();setDragging(false);const f=e.dataTransfer.files?.[0];if(f?.name.endsWith(".xlsx")) addGPD(f);}}
              style={{position:"relative",margin:14,border:`1.5px dashed ${dragging?t.accent:t.borderAcc}`,borderRadius:10,padding:"18px 12px",textAlign:"center",background:dragging?t.bgAccLt:t.bg,transition:"all 0.2s",overflow:"hidden"}}>
              <div style={{fontSize:28,marginBottom:8,pointerEvents:"none"}}>{uploading?"⟳":uploadErr?"⚠️":xlsxReady?"📊":"⟳"}</div>
              <div style={{fontSize:11,lineHeight:1.8,pointerEvents:"none",color:uploadErr?"#ef4444":t.txtSec}}>
                {uploading?<span style={{color:t.accent}}>Processando...</span>
                  :uploadErr?<>{uploadErr}<br/><span style={{fontSize:10,color:t.txtMuted}}>Tente novamente</span></>
                  :xlsxReady?<>Arraste o GPD aqui ou<br/><span style={{color:t.accent,fontWeight:700}}>clique para selecionar .xlsx</span></>
                  :<>Carregando leitor...</>}
              </div>
              {xlsxReady&&!uploading&&(
                <input type="file" accept=".xlsx"
                  style={{position:"absolute",inset:0,width:"100%",height:"100%",opacity:0,cursor:"pointer",zIndex:2}}
                  onChange={e=>{const f=e.target.files?.[0];if(f){addGPD(f);e.target.value="";}}}/>
              )}
            </div>

            {/* List */}
            <div style={{padding:"0 14px 8px",fontSize:9,color:t.txtMuted,letterSpacing:"0.1em",fontWeight:600}}>HISTÓRICO — {gpds.length} ARQUIVO{gpds.length!==1?"S":""}</div>
            <div style={{flex:1,overflowY:"auto",padding:"0 14px 14px"}}>
              {gpds.length===0&&<div style={{textAlign:"center",padding:"32px 0",fontSize:11,color:t.txtMuted,lineHeight:1.9}}>Nenhum GPD carregado.<br/>Faça upload do seu<br/>primeiro arquivo.</div>}
              {sorted.map(g=>{
                const hc=healthColor(g.stats),isAct=selected===g.calendarDate;
                return(
                  <div key={g.id} className="gi"
                    onClick={()=>{setSelected(g.calendarDate);setNavDate(new Date(g.calendarDate+"T12:00:00"));}}
                    style={{padding:"10px 12px",borderRadius:8,marginBottom:6,cursor:"pointer",border:`1px solid ${isAct?t.accent+"88":t.border}`,background:isAct?t.bgAccLt:t.bgCard,transition:"all 0.15s"}}>
                    <div style={{display:"flex",justifyContent:"space-between",marginBottom:4}}>
                      <span style={{fontWeight:700,fontSize:13,color:t.txt}}>{g.referencia||g.filename?.slice(0,12)}</span>
                      {hc&&<span style={{width:8,height:8,borderRadius:"50%",background:hc,display:"inline-block",marginTop:3,flexShrink:0}}/>}
                    </div>
                    <div style={{fontSize:10,color:t.txtSec,marginBottom:5}}>{g.geradoEm||g.calendarDate}</div>
                    <div style={{fontSize:10,display:"flex",gap:10}}>
                      <span style={{color:"#ef4444"}}>🔴{g.stats?.emVermelho||0}</span>
                      <span style={{color:"#f59e0b"}}>🟡{g.stats?.emAmarelo||0}</span>
                      <span style={{color:"#22c55e"}}>🟢{g.stats?.emVerde||0}</span>
                    </div>
                  </div>
                );
              })}
            </div>
          </aside>

          {/* MAIN */}
          <main style={{flex:1,overflowY:"auto",padding:"20px 24px"}}>

            {/* Nav */}
            <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:22,flexWrap:"wrap"}}>
              <div style={{display:"flex",gap:2,background:t.bgStat,borderRadius:9,padding:3,border:`1px solid ${t.border}`}}>
                {[{k:VIEWS.ANNUAL,l:"Anual"},{k:VIEWS.MONTHLY,l:"Mensal"},{k:VIEWS.WEEKLY,l:"Semanal"}].map(({k,l})=>(
                  <button key={k} className="ba" onClick={()=>setView(k)}
                    style={{background:view===k?t.accent:"transparent",color:view===k?t.accentInv:t.txtSec,border:"none",borderRadius:7,padding:"6px 16px",fontSize:11,cursor:"pointer",fontFamily:"'IBM Plex Mono',monospace",fontWeight:view===k?700:400,letterSpacing:"0.04em",transition:"all 0.15s"}}>{l}</button>
                ))}
              </div>
              <button className="ba" onClick={()=>navigate(-1)} style={{background:t.bgStat,color:t.txt,border:`1px solid ${t.border}`,borderRadius:7,padding:"5px 16px",fontSize:20,cursor:"pointer",lineHeight:1.3}}>‹</button>
              <span style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:24,letterSpacing:"0.1em",color:t.txt,minWidth:view===VIEWS.ANNUAL?70:180,textAlign:"center"}}>{navLabel}</span>
              <button className="ba" onClick={()=>navigate(1)} style={{background:t.bgStat,color:t.txt,border:`1px solid ${t.border}`,borderRadius:7,padding:"5px 16px",fontSize:20,cursor:"pointer",lineHeight:1.3}}>›</button>
              <span style={{fontSize:10,color:t.txtMuted}}>{gpdsInPeriod} GPD{gpdsInPeriod!==1?"s":""} no período</span>
              <button className="ba" onClick={()=>setNavDate(new Date())} style={{marginLeft:"auto",background:t.bgStat,color:t.accent,border:`1px solid ${t.accent}44`,borderRadius:7,padding:"5px 14px",fontSize:10,cursor:"pointer",fontFamily:"'IBM Plex Mono',monospace"}}>↻ Hoje</button>
            </div>

            {view===VIEWS.ANNUAL&&<AnnualView year={navYear} byDate={byDate} selected={selected} onSelect={ds=>{setSelected(ds);setNavDate(new Date(ds+"T12:00:00"));}} t={t}/>}
            {view===VIEWS.MONTHLY&&<MonthlyView year={navYear} month={navMonth} byDate={byDate} selected={selected} onSelect={ds=>{setSelected(ds);setNavDate(new Date(ds+"T12:00:00"));}} t={t}/>}
            {view===VIEWS.WEEKLY&&<WeeklyView navDate={navDate} byDate={byDate} selected={selected} onSelect={ds=>{setSelected(ds);setNavDate(new Date(ds+"T12:00:00"));}} t={t}/>}

            {selGpd&&<DetailPanel gpd={selGpd} t={t}
              onClose={()=>setSelected(null)}
              onRemove={async()=>{const next=gpds.filter(g=>g.id!==selGpd.id);setGpds(next);await saveToFirebase(next);setSelected(null);}}/>}
          </main>
        </div>
      </div>
      {ConfigPanel}
    </>
  );
}

export default function GPDArchive(){
  const t=T(false);
  return <ErrorBoundary><GPDArchiveInner/></ErrorBoundary>;
}
