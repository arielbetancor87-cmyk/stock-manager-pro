import { useState, useEffect } from "react";
import * as XLSX from "xlsx";

// ─── PERSISTENCIA ─────────────────────────────────────────────────────────────
const DB = {
  get: function(k, fb) {
    try { var v = localStorage.getItem("smp_" + k); return v ? JSON.parse(v) : fb; }
    catch(e) { return fb; }
  },
  set: function(k, v) {
    try { localStorage.setItem("smp_" + k, JSON.stringify(v)); } catch(e) {}
  }
};

// ─── CONSTANTES ───────────────────────────────────────────────────────────────
var ROLES  = ["Distribuidora Mayorista", "Revendedora Exclusiva", "Vendedora", "Otro"];
var CATS   = ["General", "Linea Kaloe", "Pocket x 10ml", "Beauty Collagen", "Premium Femenino", "Exhibidores", "Joyas", "Cosmetica"];
var EMOJIS = [
  {v:"✨",l:"Brillo"}, {v:"🧴",l:"Crema"}, {v:"💼",l:"Maletin"},
  {v:"💄",l:"Maquillaje"}, {v:"💎",l:"Joya"}, {v:"📿",l:"Cadena"},
  {v:"🌸",l:"Floral"}, {v:"🎀",l:"Accesorio"}, {v:"🍃",l:"Natural"}, {v:"⭐",l:"Destacado"}
];
var COLORS = ["#6366f1","#10b981","#f59e0b","#ef4444","#0ea5e9","#8b5cf6","#ec4899","#14b8a6"];

function uid() { return "id_" + Date.now() + "_" + Math.random().toString(36).slice(2,7); }
function fmtARS(n) { return "$ " + Number(n||0).toLocaleString("es-AR", {minimumFractionDigits:2}); }
function initials(name) { return (name||"?").trim().split(" ").slice(0,2).map(function(w){return w[0];}).join("").toUpperCase(); }

// ─── ICONOS ───────────────────────────────────────────────────────────────────
function Ic(props) {
  var n = props.n; var s = props.s || 18;
  var p = {width:s,height:s,viewBox:"0 0 24 24",fill:"none",stroke:"currentColor",strokeWidth:"2",strokeLinecap:"round",strokeLinejoin:"round"};
  switch(n) {
    case "box":    return <svg {...p}><polyline points="21,8 21,21 3,21 3,8"/><rect x="1" y="3" width="22" height="5"/></svg>;
    case "plus":   return <svg {...p}><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>;
    case "send":   return <svg {...p}><line x1="22" y1="2" x2="11" y2="13"/><polygon points="22,2 15,22 11,13 2,9 22,2"/></svg>;
    case "upload": return <svg {...p}><polyline points="16,16 12,12 8,16"/><line x1="12" y1="12" x2="12" y2="21"/><path d="M20.39 18.39A5 5 0 0 0 18 9h-1.26A8 8 0 1 0 3 16.3"/></svg>;
    case "list":   return <svg {...p}><line x1="8" y1="6" x2="21" y2="6"/><line x1="8" y1="12" x2="21" y2="12"/><line x1="8" y1="18" x2="21" y2="18"/><line x1="3" y1="6" x2="3.01" y2="6"/><line x1="3" y1="12" x2="3.01" y2="12"/><line x1="3" y1="18" x2="3.01" y2="18"/></svg>;
    case "users":  return <svg {...p}><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>;
    case "clock":  return <svg {...p}><circle cx="12" cy="12" r="10"/><polyline points="12,6 12,12 16,14"/></svg>;
    case "check":  return <svg {...p} strokeWidth="2.5"><polyline points="20,6 9,17 4,12"/></svg>;
    case "x":      return <svg {...p} strokeWidth="2.5"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>;
    case "edit":   return <svg {...p}><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4Z"/></svg>;
    case "trash":  return <svg {...p}><polyline points="3,6 5,6 21,6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>;
    case "eye":    return <svg {...p}><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg>;
    case "eyeoff": return <svg {...p}><path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94"/><path d="M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19"/><line x1="1" y1="1" x2="23" y2="23"/></svg>;
    case "search": return <svg {...p}><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>;
    case "minus":  return <svg {...p}><line x1="5" y1="12" x2="19" y2="12"/></svg>;
    case "logout": return <svg {...p}><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/><polyline points="16,17 21,12 16,7"/><line x1="21" y1="12" x2="9" y2="12"/></svg>;
    case "undo":   return <svg {...p}><polyline points="1,4 1,10 7,10"/><path d="M3.51 15a9 9 0 1 0 2.13-9.36L1 10"/></svg>;
    case "img":    return <svg {...p}><rect x="3" y="3" width="18" height="18" rx="2"/><circle cx="8.5" cy="8.5" r="1.5"/><polyline points="21,15 16,10 5,21"/></svg>;
    case "link":   return <svg {...p}><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/></svg>;
    case "wa":     return <svg {...p} fill="currentColor" stroke="none" viewBox="0 0 24 24"><path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347zm-5.421 7.403h-.004a9.87 9.87 0 0 1-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 0 1-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 0 1 2.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884zm8.413-18.297A11.815 11.815 0 0 0 12.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 0 0 5.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 0 0-3.48-8.413z"/></svg>;
    default:       return null;
  }
}

// ─── SUBCOMPONENTES ───────────────────────────────────────────────────────────
function Avatar(props) {
  var name = props.name; var color = props.color; var size = props.size || 38; var style = props.style || {};
  return (
    <div style={Object.assign({
      width:size, height:size, borderRadius:"50%", background:color||"#6366f1",
      display:"flex", alignItems:"center", justifyContent:"center",
      fontSize:Math.round(size*0.36), fontWeight:800, color:"#fff", flexShrink:0,
      fontFamily:"var(--hf)"
    }, style)}>
      {initials(name)}
    </div>
  );
}

function ProdThumb(props) {
  var prod = props.prod; var size = props.size || 40;
  if (!prod) return <div style={{width:size,height:size,borderRadius:9,background:"var(--bg2)",flexShrink:0}}/>;
  if (prod.photo) return <img src={prod.photo} alt="" style={{width:size,height:size,borderRadius:9,objectFit:"cover",flexShrink:0,border:"1px solid var(--brd)"}}/>;
  return (
    <div style={{width:size,height:size,borderRadius:9,background:"var(--bg2)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:Math.round(size*0.5),flexShrink:0}}>
      {prod.emoji||"📦"}
    </div>
  );
}

function SearchBar(props) {
  return (
    <div style={{position:"relative",marginBottom:12}}>
      <span style={{position:"absolute",left:12,top:"50%",transform:"translateY(-50%)",color:"var(--t3)",pointerEvents:"none"}}><Ic n="search" s={16}/></span>
      <input
        style={{width:"100%",padding:"11px 38px",borderRadius:10,border:"1.5px solid var(--brd)",background:"var(--card)",color:"var(--t1)",fontFamily:"var(--hf)",fontSize:14,outline:"none"}}
        placeholder={props.placeholder||"Buscar..."}
        value={props.value}
        onChange={function(e){props.onChange(e.target.value);}}
      />
      {props.value ? (
        <button onClick={function(){props.onChange("");}} style={{position:"absolute",right:10,top:"50%",transform:"translateY(-50%)",background:"var(--bg2)",border:"none",borderRadius:6,padding:4,cursor:"pointer",color:"var(--t3)",display:"flex"}}>
          <Ic n="x" s={14}/>
        </button>
      ) : null}
    </div>
  );
}

function Toast(props) {
  var t = props.t;
  var styles = {s:{bg:"#ecfdf5",brd:"#10b981",ico:"#10b981"}, e:{bg:"#fef2f2",brd:"#ef4444",ico:"#ef4444"}, i:{bg:"#eff6ff",brd:"#6366f1",ico:"#6366f1"}};
  var st = styles[t.type]||styles.s;
  return (
    <div style={{background:st.bg,border:"1px solid "+st.brd,borderLeft:"4px solid "+st.brd,borderRadius:12,padding:"12px 14px",display:"flex",alignItems:"flex-start",gap:10,boxShadow:"0 4px 16px rgba(0,0,0,.1)"}}>
      <div style={{color:st.ico,marginTop:1}}>{t.type==="s"?<Ic n="check" s={16}/>:<Ic n="x" s={16}/>}</div>
      <div style={{flex:1}}>
        <div style={{fontWeight:700,fontSize:13,color:"#1a1d2e"}}>{t.title}</div>
        {t.body?<div style={{fontSize:12,color:"#4b5068",marginTop:2}}>{t.body}</div>:null}
      </div>
      <button onClick={function(){props.remove(t.id);}} style={{background:"none",border:"none",cursor:"pointer",color:"#8b90a8",padding:2}}><Ic n="x" s={14}/></button>
    </div>
  );
}

function QtyControl(props) {
  var val = props.val; var set = props.set; var min = props.min||1; var max = props.max||9999;
  var big = props.big;
  if (big) {
    return (
      <div style={{display:"flex",border:"1.5px solid var(--brd)",borderRadius:12,overflow:"hidden",background:"var(--card)"}}>
        <button onClick={function(){set(Math.max(min,val-1));}} style={{flex:1,height:56,border:"none",background:"var(--bg2)",cursor:"pointer",color:"var(--t2)",display:"flex",alignItems:"center",justifyContent:"center"}}><Ic n="minus" s={20}/></button>
        <div style={{flex:2,textAlign:"center",fontFamily:"var(--mf)",fontSize:28,fontWeight:800,color:"var(--in)",borderLeft:"1.5px solid var(--brd)",borderRight:"1.5px solid var(--brd)",background:"var(--card)",lineHeight:"56px"}}>{val}</div>
        <button onClick={function(){set(Math.min(max,val+1));}} style={{flex:1,height:56,border:"none",background:"var(--bg2)",cursor:"pointer",color:"var(--t2)",display:"flex",alignItems:"center",justifyContent:"center"}}><Ic n="plus" s={20}/></button>
      </div>
    );
  }
  return (
    <div style={{display:"flex",border:"1.5px solid var(--brd)",borderRadius:9,overflow:"hidden",background:"var(--card)",width:"fit-content"}}>
      <button onClick={function(){set(Math.max(min,val-1));}} style={{width:32,height:32,border:"none",background:"none",cursor:"pointer",color:"var(--t3)",display:"flex",alignItems:"center",justifyContent:"center"}}><Ic n="minus" s={14}/></button>
      <div style={{minWidth:36,textAlign:"center",fontFamily:"var(--mf)",fontSize:15,fontWeight:700,color:"var(--in)",borderLeft:"1.5px solid var(--brd)",borderRight:"1.5px solid var(--brd)",background:"var(--card)",lineHeight:"32px"}}>{val}</div>
      <button onClick={function(){set(Math.min(max,val+1));}} style={{width:32,height:32,border:"none",background:"none",cursor:"pointer",color:"var(--t3)",display:"flex",alignItems:"center",justifyContent:"center"}}><Ic n="plus" s={14}/></button>
    </div>
  );
}

// ─── CSS ──────────────────────────────────────────────────────────────────────
const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;600;700&display=swap');
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{
  --bg:#f4f5fa; --bg2:#eceef5; --card:#fff; --card2:#f7f8fc;
  --brd:#e2e5ef; --brd2:#d0d4e4;
  --t1:#1a1d2e; --t2:#4b5068; --t3:#8b90a8; --t4:#b8bdd0;
  --in:#6366f1; --in-d:#4f46e5; --in-l:#eef2ff; --in-t:rgba(99,102,241,.1);
  --em:#10b981; --em-d:#059669; --em-l:#ecfdf5; --em-t:rgba(16,185,129,.1);
  --am:#f59e0b; --am-d:#d97706; --am-l:#fffbeb; --am-t:rgba(245,158,11,.1);
  --cr:#ef4444; --cr-d:#dc2626; --cr-l:#fef2f2; --cr-t:rgba(239,68,68,.08);
  --wa:#25d366; --wa-d:#128c7e; --wa-l:#dcfce7; --wa-t:rgba(37,211,102,.1);
  --hf:'Plus Jakarta Sans',sans-serif; --mf:'JetBrains Mono',monospace;
  --r:10px; --r2:14px; --r3:18px;
  --sh:0 1px 3px rgba(26,29,46,.06),0 2px 8px rgba(26,29,46,.06);
  --sh2:0 4px 20px rgba(26,29,46,.1);
  --tab:64px;
}
html,body{height:100%;background:var(--bg);color:var(--t1);font-family:var(--hf);font-size:14px;-webkit-font-smoothing:antialiased}
@keyframes fadeUp{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:translateY(0)}}
@keyframes zoomIn{from{opacity:0;transform:scale(.96)}to{opacity:1;transform:scale(1)}}
@keyframes toastIn{from{opacity:0;transform:translateY(14px) scale(.97)}to{opacity:1;transform:translateY(0) scale(1)}}
@keyframes shake{0%,100%{transform:translateX(0)}20%,60%{transform:translateX(-7px)}40%,80%{transform:translateX(7px)}}
@keyframes pulse{0%,100%{opacity:1}50%{opacity:.5}}

/* LAYOUT */
.app{display:flex;flex-direction:column;height:100vh;overflow:hidden}
.main{flex:1;overflow-y:auto;padding-bottom:var(--tab)}
.tabbar{position:fixed;bottom:0;left:0;right:0;height:var(--tab);background:var(--card);border-top:1px solid var(--brd);display:flex;z-index:50;box-shadow:0 -4px 16px rgba(26,29,46,.06)}
.tab{flex:1;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:3px;cursor:pointer;color:var(--t3);transition:color .15s;position:relative;padding:4px 2px 2px}
.tab:active{background:var(--bg2)}
.tab.on{color:var(--in)}
.tab-bubble{width:36px;height:26px;border-radius:13px;display:flex;align-items:center;justify-content:center;transition:background .15s}
.tab.on .tab-bubble{background:var(--in-t)}
.tab-lbl{font-size:10px;font-weight:600}
.tab-dot{position:absolute;top:5px;right:calc(50% - 20px);width:7px;height:7px;border-radius:50%;background:var(--cr);border:2px solid var(--card)}

/* HEADER */
.hdr{display:flex;align-items:center;gap:10px;padding:11px 14px;background:var(--card);border-bottom:1px solid var(--brd);flex-shrink:0}
.hdr-name{font-size:15px;font-weight:800;color:var(--t1);letter-spacing:-.02em}
.hdr-role{font-size:10px;color:var(--t3);margin-top:1px}
.saved-dot{display:flex;align-items:center;gap:4px;background:var(--em-l);border:1px solid var(--em-t);border-radius:20px;padding:3px 9px;font-size:10px;font-weight:700;color:var(--em-d)}

/* PAGE */
.ph{padding:16px 14px 10px;display:flex;align-items:flex-end;justify-content:space-between}
.ph-h{font-size:22px;font-weight:800;color:var(--t1);letter-spacing:-.03em}
.ph-s{font-size:12px;color:var(--t3);margin-top:2px}
.pc{padding:0 12px 20px;animation:fadeUp .28s ease both}

/* CARD */
.card{background:var(--card);border-radius:var(--r2);border:1px solid var(--brd);overflow:hidden;margin-bottom:12px;box-shadow:var(--sh)}
.card-h{display:flex;align-items:center;justify-content:space-between;padding:12px 14px;border-bottom:1px solid var(--brd)}
.card-title{font-size:14px;font-weight:700;display:flex;align-items:center;gap:8px}
.card-ico{width:28px;height:28px;border-radius:8px;display:flex;align-items:center;justify-content:center;flex-shrink:0}
.badge{display:inline-flex;align-items:center;padding:2px 8px;border-radius:20px;font-size:11px;font-weight:700}
.b-in{background:var(--in-t);color:var(--in-d)} .b-em{background:var(--em-t);color:var(--em-d)}
.b-am{background:var(--am-t);color:var(--am-d)} .b-cr{background:var(--cr-t);color:var(--cr)}
.b-wa{background:var(--wa-t);color:var(--wa-d)} .b-sl{background:var(--bg2);color:var(--t2)}

/* TABLE */
.tw{overflow-x:auto}
table{width:100%;border-collapse:collapse}
th{padding:8px 10px;text-align:left;font-size:9px;text-transform:uppercase;letter-spacing:.09em;color:var(--t3);border-bottom:1px solid var(--brd);background:var(--bg2);font-weight:700;white-space:nowrap}
td{padding:10px 10px;border-bottom:1px solid var(--brd);font-size:12px;vertical-align:middle}
tr:last-child td{border-bottom:none}
.tr:hover td{background:var(--bg2)}

/* FORMS */
.fld{margin-bottom:12px}
.fl{display:block;font-size:11px;font-weight:600;color:var(--t2);margin-bottom:5px}
.fi{width:100%;padding:11px 13px;border-radius:var(--r);border:1.5px solid var(--brd);background:var(--card);color:var(--t1);font-family:var(--hf);font-size:14px;outline:none;transition:border .13s,box-shadow .13s}
.fi:focus{border-color:var(--in);box-shadow:0 0 0 3px var(--in-t)}
.fi-sel{appearance:none;cursor:pointer}
.pw-wrap{position:relative}
.pw-eye{position:absolute;right:11px;top:50%;transform:translateY(-50%);background:none;border:none;cursor:pointer;color:var(--t3);display:flex;padding:4px}

/* BUTTONS */
.btn{display:inline-flex;align-items:center;gap:6px;padding:8px 14px;border-radius:var(--r);border:none;cursor:pointer;font-family:var(--hf);font-size:12px;font-weight:700;transition:all .14s;white-space:nowrap;flex-shrink:0}
.btn:disabled{opacity:.3;cursor:not-allowed}
.btn:active:not(:disabled){transform:scale(.97)}
.b-pri{background:linear-gradient(135deg,var(--in),var(--in-d));color:#fff;box-shadow:0 2px 8px rgba(99,102,241,.3)}
.b-pri:hover:not(:disabled){box-shadow:0 4px 14px rgba(99,102,241,.4)}
.b-em{background:var(--em-t);color:var(--em-d);border:1px solid rgba(16,185,129,.2)}
.b-em:hover:not(:disabled){background:var(--em);color:#fff}
.b-am{background:var(--am-t);color:var(--am-d);border:1px solid rgba(245,158,11,.2)}
.b-am:hover:not(:disabled){background:var(--am);color:#fff}
.b-cr{background:var(--cr-t);color:var(--cr);border:1px solid rgba(239,68,68,.2)}
.b-cr:hover:not(:disabled){background:var(--cr);color:#fff}
.b-in{background:var(--in-t);color:var(--in-d);border:1px solid rgba(99,102,241,.2)}
.b-in:hover:not(:disabled){background:var(--in);color:#fff}
.b-ghost{background:var(--card);border:1.5px solid var(--brd2);color:var(--t2)}
.b-ghost:hover{background:var(--bg2);color:var(--t1)}
.b-wa{background:var(--wa-t);color:var(--wa-d);border:1px solid rgba(37,211,102,.25)}
.b-wa:hover:not(:disabled){background:var(--wa);color:#fff}
.b-wa-full{background:linear-gradient(135deg,var(--wa),var(--wa-d));color:#fff;border:none;box-shadow:0 3px 12px rgba(37,211,102,.35)}
.btn-xs{padding:5px 9px;font-size:11px;border-radius:8px}
.cta{width:100%;padding:17px;border-radius:var(--r2);border:none;font-family:var(--hf);font-size:16px;font-weight:800;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:8px;margin-top:14px;transition:all .15s}
.cta:active:not(:disabled){transform:scale(.98)}
.cta:disabled{opacity:.35;cursor:not-allowed}
.cta-am{background:linear-gradient(135deg,var(--am),var(--am-d));color:#fff;box-shadow:0 4px 14px rgba(245,158,11,.35)}
.cta-em{background:linear-gradient(135deg,var(--em),var(--em-d));color:#fff;box-shadow:0 4px 14px rgba(16,185,129,.35)}
.cta-wa{background:linear-gradient(135deg,var(--wa),var(--wa-d));color:#fff;box-shadow:0 4px 14px rgba(37,211,102,.35)}
.cta-in{background:linear-gradient(135deg,var(--in),var(--in-d));color:#fff;box-shadow:0 4px 14px rgba(99,102,241,.35)}

/* MODAL */
.ovl{position:fixed;inset:0;background:rgba(26,29,46,.45);backdrop-filter:blur(8px);z-index:200;display:flex;align-items:flex-end;justify-content:center;animation:fadeUp .2s ease}
.mbox{background:var(--card);border-radius:var(--r3) var(--r3) 0 0;width:100%;max-width:600px;max-height:90vh;overflow-y:auto;box-shadow:0 -8px 40px rgba(26,29,46,.15);animation:fadeUp .3s cubic-bezier(.16,1,.3,1)}
.mhd{display:flex;align-items:center;justify-content:space-between;padding:18px 18px 12px;border-bottom:1px solid var(--brd)}
.mhd-t{font-size:16px;font-weight:800;color:var(--t1)}
.mbd{padding:16px 18px}
.mft{padding:12px 18px;border-top:1px solid var(--brd);display:flex;justify-content:flex-end;gap:8px}
.ic-btn{width:36px;height:36px;border-radius:10px;border:none;background:var(--bg2);cursor:pointer;color:var(--t3);display:flex;align-items:center;justify-content:center;transition:all .13s}
.ic-btn:hover{background:var(--brd);color:var(--t1)}

/* LOGIN */
.login-bg{min-height:100vh;display:flex;align-items:center;justify-content:center;background:var(--bg);padding:16px;background-image:radial-gradient(ellipse 500px 400px at 70% 15%,rgba(99,102,241,.07) 0%,transparent 60%),radial-gradient(ellipse 400px 350px at 20% 80%,rgba(16,185,129,.06) 0%,transparent 60%)}
.lbox{background:var(--card);border-radius:var(--r3);width:100%;max-width:420px;padding:32px 26px;box-shadow:var(--sh2);border:1px solid var(--brd)}
.lbox.shake{animation:shake .38s ease}
.ltabs{display:flex;border-radius:var(--r);overflow:hidden;border:1.5px solid var(--brd);padding:3px;background:var(--bg2);margin-bottom:20px}
.ltab{flex:1;padding:9px;text-align:center;font-size:13px;font-weight:700;cursor:pointer;border-radius:8px;color:var(--t3);transition:all .15s}
.ltab.on{background:var(--card);color:var(--t1);box-shadow:var(--sh)}
.err-box{background:var(--cr-l);border:1px solid rgba(239,68,68,.2);border-radius:var(--r);padding:10px 13px;font-size:13px;color:var(--cr);margin-bottom:14px;display:flex;align-items:center;gap:7px}

/* QUICK LOAD */
.ql-tabs{display:flex;border-radius:var(--r);overflow:hidden;border:1.5px solid var(--brd);background:var(--bg2);padding:3px;margin-bottom:16px}
.ql-tab{flex:1;padding:10px;text-align:center;font-size:13px;font-weight:700;cursor:pointer;color:var(--t3);border-radius:7px;transition:all .15s}
.ql-tab.on{background:var(--am);color:#fff;box-shadow:0 2px 8px rgba(245,158,11,.3)}
.prod-card{display:flex;align-items:center;gap:12px;padding:12px 14px;border-radius:var(--r2);border:2px solid var(--brd);background:var(--card);cursor:pointer;margin-bottom:8px;transition:all .15s}
.prod-card:active{transform:scale(.98)}
.prod-card.sel{border-color:var(--am);background:var(--am-l)}

/* SEND */
.step-row{display:flex;align-items:flex-start;margin-bottom:18px}
.step-dot{width:30px;height:30px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:13px;font-weight:700;flex-shrink:0;border:2px solid var(--brd);color:var(--t3);background:var(--card);transition:all .2s}
.step-dot.on{background:var(--in);border-color:var(--in);color:#fff;box-shadow:0 0 0 4px var(--in-t)}
.step-dot.done{background:var(--em);border-color:var(--em);color:#fff}
.step-line{flex:1;height:2px;background:var(--brd);margin:14px 6px 0;transition:background .2s}
.step-line.done{background:var(--em)}
.step-lbl{font-size:10px;color:var(--t3);margin-top:4px;font-weight:600;text-align:center}
.ct-card{display:flex;align-items:center;gap:12px;padding:14px 16px;border-radius:var(--r2);border:2px solid var(--brd);background:var(--card);cursor:pointer;margin-bottom:10px;transition:all .15s}
.ct-card:active{transform:scale(.98)}
.ct-card.sel{border-color:var(--em);background:var(--em-l)}
.sic{display:flex;align-items:center;gap:10px;padding:11px 13px;border-radius:var(--r);border:1px solid var(--brd);background:var(--card);margin-bottom:8px}
.cart-box{background:var(--em-l);border:1px solid var(--em-t);border-radius:var(--r);padding:12px 14px;margin:12px 0}

/* IMPORT */
.drop-z{border:2px dashed var(--brd2);border-radius:var(--r2);padding:40px 20px;text-align:center;cursor:pointer;transition:all .18s;background:var(--bg2);position:relative;overflow:hidden}
.drop-z:hover,.drop-z.drag{border-color:var(--in);background:var(--in-l)}
.drop-z input{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%}
.prev-row{display:flex;align-items:center;gap:8px;padding:7px 10px;border-radius:8px;margin-bottom:5px;font-size:12px}
.prev-ok{background:var(--em-l);border:1px solid var(--em-t)} .prev-err{background:var(--cr-l);border:1px solid var(--cr-t)}

/* PHOTO */
.photo-zone{border:2px dashed var(--brd2);border-radius:var(--r);padding:18px;text-align:center;cursor:pointer;transition:all .15s;background:var(--bg2);position:relative;overflow:hidden;margin-bottom:10px}
.photo-zone:hover{border-color:var(--in);background:var(--in-l)}
.photo-zone input{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%}

/* SHARE BANNER */
.wa-banner{width:100%;display:flex;align-items:center;justify-content:space-between;padding:16px 18px;border-radius:var(--r2);background:linear-gradient(135deg,var(--wa),var(--wa-d));border:none;cursor:pointer;margin-bottom:14px;box-shadow:0 4px 16px rgba(37,211,102,.3);transition:all .15s}
.wa-banner:active{transform:scale(.98)}

/* EMPTY */
.empty{text-align:center;padding:36px 16px;color:var(--t3);font-size:13px;line-height:1.7}

/* TOASTS */
.toast-wrap{position:fixed;bottom:calc(var(--tab) + 10px);left:12px;right:12px;z-index:500;display:flex;flex-direction:column;gap:7px;pointer-events:none}
.toast-wrap > *{pointer-events:all;animation:toastIn .3s cubic-bezier(.16,1,.3,1)}

/* MISC */
.row{display:flex;align-items:center} .jb{justify-content:space-between} .g8{gap:8px} .g12{gap:12px}
.or-div{display:flex;align-items:center;gap:8px;color:var(--t4);font-size:11px;margin:10px 0}
.or-div::before,.or-div::after{content:"";flex:1;height:1px;background:var(--brd)}
::-webkit-scrollbar{width:4px;height:4px} ::-webkit-scrollbar-thumb{background:var(--brd2);border-radius:3px}
`;

// ─── APP ──────────────────────────────────────────────────────────────────────
export default function App() {
  // ── PERSISTENCIA ────────────────────────────────────────────────────────────
  var [users,    setUsers]    = useState(function(){ return DB.get("users",    []); });
  var [allProds, setAllProds] = useState(function(){ return DB.get("prods",    []); });
  var [allInv,   setAllInv]   = useState(function(){ return DB.get("inv",      {}); });
  var [allCons,  setAllCons]  = useState(function(){ return DB.get("contacts", {}); });
  var [allLogs,  setAllLogs]  = useState(function(){ return DB.get("logs",     []); });
  var [transfers,setTransfers]= useState(function(){ return DB.get("transfers",[]); });
  var [notifs,   setNotifs]   = useState(function(){ return DB.get("notifs",   []); });

  useEffect(function(){ DB.set("users",    users);    }, [users]);
  useEffect(function(){ DB.set("prods",    allProds); }, [allProds]);
  useEffect(function(){ DB.set("inv",      allInv);   }, [allInv]);
  useEffect(function(){ DB.set("contacts", allCons);  }, [allCons]);
  useEffect(function(){ DB.set("logs",     allLogs);  }, [allLogs]);
  useEffect(function(){ DB.set("transfers",transfers); }, [transfers]);
  useEffect(function(){ DB.set("notifs",    notifs);    }, [notifs]);

  // ── SESSION ─────────────────────────────────────────────────────────────────
  var [me, setMe] = useState(null);
  var [tab, setTab] = useState("stock");
  var [toasts, setToasts] = useState([]);

  // ── AUTH FORM ────────────────────────────────────────────────────────────────
  var [authMode,   setAuthMode]   = useState("login");
  var [aEmail,     setAEmail]     = useState("");
  var [aPass,      setAPass]      = useState("");
  var [aName,      setAName]      = useState("");
  var [aRole,      setARole]      = useState("Revendedora Exclusiva");
  var [aPass2,     setAPass2]     = useState("");
  var [showPass,   setShowPass]   = useState(false);
  var [authErr,    setAuthErr]    = useState("");
  var [authShake,  setAuthShake]  = useState(false);

  // ── STOCK/SEND SEARCH ────────────────────────────────────────────────────────
  var [srchStock,  setSrchStock]  = useState("");
  var [srchCat,    setSrchCat]    = useState("");
  var [srchCon,    setSrchCon]    = useState("");
  var [srchLog,    setSrchLog]    = useState("");

  // ── CATALOG ABM ─────────────────────────────────────────────────────────────
  var [editP,    setEditP]    = useState(null);
  var [fSku,     setFSku]     = useState("");
  var [fName,    setFName]    = useState("");
  var [fPrice,   setFPrice]   = useState("");
  var [fCat,     setFCat]     = useState("General");
  var [fEmoji,   setFEmoji]   = useState("✨");
  var [fPhoto,   setFPhoto]   = useState(null);
  var [fStock,   setFStock]   = useState("0");
  var [delConf,  setDelConf]  = useState(null);

  // ── CARGAR STOCK ─────────────────────────────────────────────────────────────
  var [qlMode,   setQlMode]   = useState("add");
  var [qlPid,    setQlPid]    = useState("");
  var [qlQty,    setQlQty]    = useState(1);
  var [qlSku,    setQlSku]    = useState("");
  var [qlName,   setQlName]   = useState("");
  var [qlPrice,  setQlPrice]  = useState("");
  var [qlCat,    setQlCat]    = useState("General");
  var [qlEmoji,  setQlEmoji]  = useState("✨");
  var [qlPhoto,  setQlPhoto]  = useState(null);
  var [qlSrch,   setQlSrch]   = useState("");

  // ── ENVIAR ───────────────────────────────────────────────────────────────────
  var [sendStep, setSendStep] = useState(1);
  var [sendTo,   setSendTo]   = useState("");
  var [sendCart, setSendCart] = useState({});
  var [sendSrch, setSendSrch] = useState("");

  // ── IMPORTAR ─────────────────────────────────────────────────────────────────
  var [impTab,   setImpTab]   = useState("file");
  var [impTxt,   setImpTxt]   = useState("");
  var [impRows,  setImpRows]  = useState([]);
  var [impQty,   setImpQty]   = useState(1);
  var [impFile,  setImpFile]  = useState(null);

  // ── CONTACTOS ────────────────────────────────────────────────────────────────
  var [ctQ,      setCtQ]      = useState("");

  // ── MODALS ───────────────────────────────────────────────────────────────────
  var [txModal,  setTxModal]  = useState(null);
  var [txQty,    setTxQty]    = useState(1);
  var [txTo,     setTxTo]     = useState("");
  var [shareM,   setShareM]   = useState(false);
  var [shareSel, setShareSel] = useState({});

  // ── HELPERS ──────────────────────────────────────────────────────────────────
  function toast(title, body, type) {
    var id = Date.now() + Math.random();
    setToasts(function(p){ return [...p, {id:id,title:title,body:body||"",type:type||"s"}]; });
    setTimeout(function(){ setToasts(function(p){ return p.filter(function(t){ return t.id!==id; }); }); }, 3800);
  }
  function rmToast(id){ setToasts(function(p){ return p.filter(function(t){ return t.id!==id; }); }); }

  function addLog(act, txt) {
    var entry = {id:uid(), uid:me.id, who:me.name, act:act, txt:txt, t:new Date().toLocaleTimeString("es-AR")};
    setAllLogs(function(p){ return [entry, ...p.slice(0,199)]; });
  }

  function addNotif(toId, msg, type) {
    var n = {id:uid(), toId:toId, msg:msg, type:type||"info", read:false, t:new Date().toLocaleTimeString("es-AR")};
    setNotifs(function(p){ return [n,...p.slice(0,99)]; });
  }

  function myInv()      { return (me && allInv[me.id]) || {stock:[], consign:[]}; }
  function myContacts() { return me ? ((allCons[me.id]||[]).map(function(id){ return users.find(function(u){ return u.id===id; }); }).filter(Boolean)) : []; }
  function myStock()    { return myInv().stock.filter(function(i){ return i.avail>0; }); }
  function myCons()     { return myInv().consign.filter(function(i){ return i.avail>0; }); }

  // ── AUTH ─────────────────────────────────────────────────────────────────────
  function shake() { setAuthShake(true); setTimeout(function(){ setAuthShake(false); }, 450); }

  function doLogin() {
    var em = aEmail.trim().toLowerCase();
    var u  = users.find(function(x){ return x.email===em; });
    if (!u)           { setAuthErr("Email no registrado."); shake(); return; }
    if (u.pass!==aPass){ setAuthErr("Contraseña incorrecta."); shake(); return; }
    setAuthErr(""); setMe(u); toast("Bienvenida, "+u.name+"!", u.role, "s");
  }

  function doRegister() {
    var em = aEmail.trim().toLowerCase();
    if (!em||!aName.trim()||!aPass) { setAuthErr("Completa todos los campos."); shake(); return; }
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(em)) { setAuthErr("Email inválido."); shake(); return; }
    if (aPass!==aPass2)              { setAuthErr("Las contraseñas no coinciden."); shake(); return; }
    if (aPass.length<4)              { setAuthErr("Contraseña mínimo 4 caracteres."); shake(); return; }
    if (users.some(function(u){ return u.email===em; })) { setAuthErr("Ese email ya está registrado."); shake(); return; }
    var color = COLORS[users.length % COLORS.length];
    var nu = {id:uid(), email:em, pass:aPass, name:aName.trim(), role:aRole, color:color};
    setUsers(function(p){ return [...p, nu]; });
    setAllInv(function(p){ return Object.assign({}, p, {[nu.id]:{stock:[],consign:[]}}); });
    setAllCons(function(p){ return Object.assign({}, p, {[nu.id]:[]}); });
    toast("Cuenta creada!", "Ya puedes iniciar sesión.", "s");
    setAuthMode("login"); setAEmail(em); setAPass(""); setAPass2(""); setAName("");
  }

  function doLogout() {
    var n = me.name;
    setMe(null); setTab("stock"); setAEmail(""); setAPass("");
    toast("Hasta luego, "+n+"!", "Tus datos están guardados.", "i");
  }

  // ── VENTAS ───────────────────────────────────────────────────────────────────
  function doSell(itemId, section) {
    var prod = null;
    var consignFrom = null;
    setAllInv(function(prev){
      var sec = prev[me.id][section].map(function(it){
        if (it.id!==itemId || it.avail<=0) return it;
        prod = allProds.find(function(p){ return p.id===it.pid; });
        if (section==="consign") consignFrom = it.from;
        return Object.assign({}, it, {avail:it.avail-1, sold:it.sold+1});
      });
      var u = Object.assign({}, prev[me.id], {[section]:sec});
      return Object.assign({}, prev, {[me.id]:u});
    });
    setTimeout(function(){
      if (prod) {
        addLog("Venta", "1x "+prod.name+(consignFrom?" (consigna de "+consignFrom+")":""));
        toast("Venta registrada", prod.name, "s");
        // Si era consignación, notificar al dueño original
        if (consignFrom) {
          var owner = users.find(function(u){ return u.name===consignFrom; });
          if (owner) addNotif(owner.id, me.name+" vendió 1x "+prod.name+" de tu consignación!", "sale");
        }
      }
    }, 0);
  }

  function doReturn(itemId) {
    var prod = null;
    setAllInv(function(prev){
      var sec = prev[me.id].consign.map(function(it){
        if (it.id!==itemId || it.avail<=0) return it;
        prod = allProds.find(function(p){ return p.id===it.pid; });
        return Object.assign({}, it, {avail:it.avail-1});
      });
      var u = Object.assign({}, prev[me.id], {consign:sec});
      return Object.assign({}, prev, {[me.id]:u});
    });
    setTimeout(function(){
      if (prod) { addLog("Devolución", prod.name); toast("Devolución guardada", "", "i"); }
    }, 0);
  }

  // ── CARGAR STOCK ─────────────────────────────────────────────────────────────
  function doQuickLoad() {
    if (qlMode==="add") {
      if (!qlPid) { toast("Selecciona un producto","","e"); return; }
      setAllInv(function(prev){
        var ui = prev[me.id]||{stock:[],consign:[]};
        var idx = ui.stock.findIndex(function(s){ return s.pid===qlPid; });
        var ns;
        if (idx>-1) {
          ns = ui.stock.map(function(s,i){ return i===idx ? Object.assign({},s,{avail:s.avail+Number(qlQty)}) : s; });
        } else {
          ns = [...ui.stock, {id:uid(), pid:qlPid, avail:Number(qlQty), sold:0}];
        }
        return Object.assign({}, prev, {[me.id]:Object.assign({},ui,{stock:ns})});
      });
      var p = allProds.find(function(x){ return x.id===qlPid; });
      addLog("Carga", "+"+qlQty+"u. de "+(p?p.name:""));
      toast("Stock cargado!", "+"+ qlQty+" unidades", "s");
      setQlQty(1); setQlPid("");
    } else {
      if (!qlSku.trim()||!qlName.trim()||!qlPrice) { toast("Completa SKU, nombre y precio","","e"); return; }
      var skuC = qlSku.trim().toUpperCase();
      if (allProds.some(function(p){ return p.sku===skuC; })) { toast("SKU ya existe","","e"); return; }
      var newId = uid();
      var newP = {id:newId, sku:skuC, name:qlName.trim(), price:parseFloat(qlPrice)||0, emoji:qlEmoji, photo:qlPhoto||null, cat:qlCat};
      setAllProds(function(p){ return [...p, newP]; });
      if (qlQty>0) {
        setAllInv(function(prev){
          var ui = prev[me.id]||{stock:[],consign:[]};
          return Object.assign({},prev,{[me.id]:Object.assign({},ui,{stock:[...ui.stock,{id:uid(),pid:newId,avail:Number(qlQty),sold:0}]})});
        });
      }
      addLog("Alta+Carga", "["+skuC+"] +"+qlQty+"u.");
      toast("Producto creado!", qlName.trim()+" con "+qlQty+" u.", "s");
      setQlSku(""); setQlName(""); setQlPrice(""); setQlQty(1); setQlPhoto(null); setQlMode("add");
    }
  }

  // ── TRANSFERENCIA SIMPLE (modal) ─────────────────────────────────────────────
  function openTx(item) {
    var cts = myContacts();
    if (!cts.length) { toast("Sin contactos","Agrega uno en Contactos","i"); return; }
    setTxModal(item); setTxQty(1); setTxTo(cts[0].id);
  }
  function doTx() {
    if (txQty>txModal.avail) { toast("Stock insuficiente","","e"); return; }
    var tUser = users.find(function(u){ return u.id===txTo; });
    var prod  = allProds.find(function(p){ return p.id===txModal.pid; });
    // Descontar del stock del emisor, crear transferencia pendiente
    setAllInv(function(prev){
      var src = Object.assign({},prev[me.id]);
      src.stock = src.stock.map(function(it){ return it.id===txModal.id ? Object.assign({},it,{avail:it.avail-txQty}) : it; });
      return Object.assign({},prev,{[me.id]:src});
    });
    var txId = uid();
    setTransfers(function(p){ return [{id:txId, from:me.id, fromName:me.name, to:txTo, pid:txModal.pid, qty:txQty, status:"pending", t:new Date().toLocaleTimeString("es-AR")},...p]; });
    addNotif(txTo, me.name+" te envió "+txQty+"x "+prod.name+". Confirmá la recepción!", "transfer");
    addLog("Enviado", txQty+"x "+prod.name+" → "+tUser.name+" (pendiente confirmación)");
    toast("Enviado!", "Esperando confirmación de "+tUser.name, "s");
    setTxModal(null);
  }

  // Confirmar recepción de transferencia
  function confirmTransfer(txId) {
    var tx = transfers.find(function(t){ return t.id===txId; });
    if (!tx) return;
    var prod = allProds.find(function(p){ return p.id===tx.pid; });
    // Sumar al stock del receptor
    setAllInv(function(prev){
      var dst = prev[me.id]||{stock:[],consign:[]};
      var idx = dst.stock.findIndex(function(s){ return s.pid===tx.pid; });
      var dstock = idx>-1
        ? dst.stock.map(function(s,i){ return i===idx?Object.assign({},s,{avail:s.avail+Number(tx.qty)}):s; })
        : [...dst.stock,{id:uid(),pid:tx.pid,avail:Number(tx.qty),sold:0}];
      return Object.assign({},prev,{[me.id]:Object.assign({},dst,{stock:dstock})});
    });
    setTransfers(function(p){ return p.map(function(t){ return t.id===txId?Object.assign({},t,{status:"confirmed"}):t; }); });
    // Notificar al emisor
    addNotif(tx.from, me.name+" confirmó la recepción de "+tx.qty+"x "+(prod?prod.name:"producto")+"!", "confirm");
    // Marcar notif como leída
    setNotifs(function(p){ return p.map(function(n){ return n.toId===me.id&&n.type==="transfer"?Object.assign({},n,{read:true}):n; }); });
    addLog("Recepción", "Confirmó "+tx.qty+"x "+(prod?prod.name:"")+" de "+tx.fromName);
    toast("Recepción confirmada!", (prod?prod.name:"")+" añadido a tu stock", "s");
  }

  // ── ENVIO MULTI ───────────────────────────────────────────────────────────────
  function doMultiSend() {
    var entries = Object.entries(sendCart).filter(function(e){ return Number(e[1])>0; });
    if (!entries.length) { toast("Agrega al menos un producto","","e"); return; }
    var tUser = users.find(function(u){ return u.id===sendTo; });
    setAllInv(function(prev){
      var src = Object.assign({},prev[me.id]);
      src.stock = src.stock.map(function(it){
        var q = Number(sendCart[it.id]||0);
        return q>0 ? Object.assign({},it,{avail:it.avail-q}) : it;
      });
      var dst = prev[sendTo]||{stock:[],consign:[]};
      var dstock = [...dst.stock];
      entries.forEach(function(e){
        var srcItem = prev[me.id].stock.find(function(i){ return i.id===e[0]; });
        if (!srcItem) return;
        var idx = dstock.findIndex(function(s){ return s.pid===srcItem.pid; });
        if (idx>-1) { dstock[idx]=Object.assign({},dstock[idx],{avail:dstock[idx].avail+Number(e[1])}); }
        else { dstock.push({id:uid(),pid:srcItem.pid,avail:Number(e[1]),sold:0}); }
      });
      return Object.assign({},prev,{[me.id]:src,[sendTo]:Object.assign({},dst,{stock:dstock})});
    });
    addLog("Envío multi", entries.length+" producto(s) → @"+tUser.name);
    toast("Envío completado!", entries.length+" producto(s) a "+tUser.name, "s");
    setSendCart({}); setSendStep(1);
  }

  // ── CATALOG ABM ───────────────────────────────────────────────────────────────
  function doSaveProd(e) {
    e.preventDefault();
    if (!fSku.trim()||!fName.trim()||!fPrice) { toast("Completa SKU, nombre y precio","","e"); return; }
    var skuC = fSku.trim().toUpperCase();
    var pVal = parseFloat(fPrice)||0;
    if (editP) {
      setAllProds(function(p){ return p.map(function(x){ return x.id===editP.id ? Object.assign({},x,{sku:skuC,name:fName.trim(),price:pVal,cat:fCat,emoji:fEmoji,photo:fPhoto!==undefined?fPhoto:x.photo}) : x; }); });
      addLog("Edición", "["+skuC+"] "+fName.trim());
      toast("Producto actualizado","","s"); setEditP(null);
    } else {
      if (allProds.some(function(p){ return p.sku===skuC; })) { toast("SKU ya existe","","e"); return; }
      var newId = uid();
      var newP  = {id:newId, sku:skuC, name:fName.trim(), price:pVal, cat:fCat, emoji:fEmoji, photo:fPhoto||null};
      setAllProds(function(p){ return [...p, newP]; });
      var qty = parseInt(fStock)||0;
      if (qty>0) {
        setAllInv(function(prev){
          var ui = prev[me.id]||{stock:[],consign:[]};
          return Object.assign({},prev,{[me.id]:Object.assign({},ui,{stock:[...ui.stock,{id:uid(),pid:newId,avail:qty,sold:0}]})});
        });
      }
      addLog("Alta", "["+skuC+"] con "+qty+"u.");
      toast("Producto creado!", fName.trim(), "s");
    }
    setFSku(""); setFName(""); setFPrice(""); setFEmoji("✨"); setFStock("0"); setFPhoto(null);
  }
  function startEdit(p) { setEditP(p); setFSku(p.sku); setFName(p.name); setFPrice(String(p.price)); setFCat(p.cat||"General"); setFEmoji(p.emoji||"✨"); setFPhoto(p.photo||null); setTab("catalog"); }
  function cancelEdit() { setEditP(null); setFSku(""); setFName(""); setFPrice(""); setFEmoji("✨"); setFStock("0"); setFPhoto(null); }
  function doDelProd(prodId, sku) {
    setAllProds(function(p){ return p.filter(function(x){ return x.id!==prodId; }); });
    setAllInv(function(prev){
      var u = Object.assign({},prev);
      Object.keys(u).forEach(function(uid2){
        u[uid2] = Object.assign({},u[uid2],{
          stock:  u[uid2].stock.filter(function(s){ return s.pid!==prodId; }),
          consign:u[uid2].consign.filter(function(c){ return c.pid!==prodId; })
        });
      });
      return u;
    });
    addLog("Baja", "["+sku+"]"); toast("Producto eliminado","","i");
    setDelConf(null); if (editP&&editP.id===prodId) cancelEdit();
  }

  // ── FOTO UPLOAD ───────────────────────────────────────────────────────────────
  function handlePhoto(e, setter) {
    var file = e.target.files && e.target.files[0];
    if (!file) return;
    if (file.size>1600000) { toast("Foto muy grande","Usa una imagen menor a 1.6MB","e"); return; }
    var reader = new FileReader();
    reader.onloadend = function() { setter(reader.result); toast("Foto cargada","","s"); };
    reader.readAsDataURL(file);
  }

  // ── IMPORTAR ─────────────────────────────────────────────────────────────────
  function parseLines(text, qty) {
    var lines = text.split("\n");
    var rows = []; var curCat = "Importado";
    lines.forEach(function(line){
      var row = line.trim();
      if (!row||row.toLowerCase().includes("mostrar")) return;
      if (!row.includes("-")&&!row.includes(",")&&!row.includes("\t")) {
        if (row.length>2&&row.length<40) curCat=row; return;
      }
      var cols = row.includes("\t")?row.split("\t"):row.split(",");
      if (cols.length<2) { rows.push({ok:false,raw:row,err:"Menos de 2 columnas"}); return; }
      var c0=cols[0].trim(), c1=cols[1].trim();
      if (!c0.includes("-")) { rows.push({ok:false,raw:row,err:"Sin guion separador"}); return; }
      var dash=c0.indexOf("-");
      var sku=c0.substring(0,dash).trim().toUpperCase();
      var name=c0.substring(dash+1).trim();
      if (!sku||!name) { rows.push({ok:false,raw:row,err:"SKU o nombre vacío"}); return; }
      var pc=c1.replace(/[^\d.,]/g,"");
      if (pc.includes(",")) { pc=pc.replace(/\./g,"").replace(",","."); }
      rows.push({ok:true, sku:sku, name:name, cat:curCat, price:parseFloat(pc)||0, qty:qty});
    });
    return rows;
  }

  function doParseFile(e) {
    var file = e.target.files&&e.target.files[0];
    if (!file) return;
    setImpFile(file.name);
    var ext = file.name.split(".").pop().toLowerCase();
    var reader = new FileReader();
    reader.onload = function(ev) {
      try {
        var rows = [];
        if (ext==="xlsx"||ext==="xls") {
          var wb = XLSX.read(ev.target.result, {type:"binary"});
          var ws = wb.Sheets[wb.SheetNames[0]];
          var data = XLSX.utils.sheet_to_json(ws, {header:1, raw:false});
          var startRow=0, colSku=-1, colName=-1, colCat=-1, colPrice=-1;
          if (data.length>0) {
            var hdr=data[0].map(function(h){ return (h||"").toString().toLowerCase(); });
            hdr.forEach(function(h,i){
              if (/sku|cod/.test(h)) colSku=i;
              else if (/nom|desc|prod|det/.test(h)) colName=i;
              else if (/cat|rub|lin|tipo/.test(h)) colCat=i;
              else if (/prec|price|val|cost/.test(h)) colPrice=i;
            });
            if (colSku>=0&&colName>=0) startRow=1;
            else { colSku=0; colName=1; colCat=2; colPrice=3; startRow=1; }
          }
          data.slice(startRow).forEach(function(row){
            if (!row||!row.length) return;
            var sku=(row[colSku]||"").toString().trim().toUpperCase();
            var name=(row[colName]||"").toString().trim();
            var cat=colCat>=0?(row[colCat]||"Importado").toString().trim():"Importado";
            var priceRaw=colPrice>=0?(row[colPrice]||"0").toString():"0";
            var pc=priceRaw.replace(/[^\d.,]/g,"");
            if (pc.includes(",")){ pc=pc.replace(/\./g,"").replace(",","."); }
            var price=parseFloat(pc)||0;
            if (!sku||!name){ rows.push({ok:false,raw:(row||[]).join(","),err:"Fila incompleta"}); return; }
            rows.push({ok:true,sku:sku,name:name,cat:cat,price:price,qty:impQty});
          });
        } else {
          rows = parseLines(ev.target.result, impQty);
        }
        setImpRows(rows);
        toast(rows.filter(function(r){return r.ok;}).length+" productos detectados","Revisa y confirma","s");
      } catch(err) { toast("Error al leer el archivo","","e"); }
    };
    if (ext==="xlsx"||ext==="xls") reader.readAsBinaryString(file);
    else reader.readAsText(file);
  }

  function doParseTxt() {
    if (!impTxt.trim()) { toast("Pega texto primero","","e"); return; }
    var rows = parseLines(impTxt, impQty);
    setImpRows(rows);
    toast(rows.filter(function(r){return r.ok;}).length+" detectados","Revisa y confirma","s");
  }

  function doConfirmImport() {
    var valid = impRows.filter(function(r){ return r.ok; });
    if (!valid.length) { toast("Sin filas válidas","","e"); return; }
    var curP = [...allProds];
    var items = [];
    valid.forEach(function(row){
      var found = curP.find(function(p){ return p.sku===row.sku; });
      if (!found) { found={id:uid(),sku:row.sku,name:row.name,price:row.price,emoji:"✨",photo:null,cat:row.cat}; curP=[...curP,found]; }
      else { found.price=row.price; }
      items.push({pid:found.id, qty:row.qty});
    });
    setAllProds(curP);
    setAllInv(function(prev){
      var ui=prev[me.id]||{stock:[],consign:[]};
      var ns=[...ui.stock];
      items.forEach(function(item){
        var idx=ns.findIndex(function(s){ return s.pid===item.pid; });
        if (idx>-1) { ns=ns.map(function(s,i){ return i===idx?Object.assign({},s,{avail:s.avail+item.qty}):s; }); }
        else { ns=[...ns,{id:uid(),pid:item.pid,avail:item.qty,sold:0}]; }
      });
      return Object.assign({},prev,{[me.id]:Object.assign({},ui,{stock:ns})});
    });
    addLog("Importación", valid.length+" productos desde planilla");
    toast("Importación exitosa!", valid.length+" productos cargados", "s");
    setImpRows([]); setImpTxt(""); setImpFile(null); setTab("stock");
  }

  // ── CONTACTOS ────────────────────────────────────────────────────────────────
  function doAddContact() {
    var em = ctQ.trim().toLowerCase();
    var target = users.find(function(u){ return u.email===em; });
    if (!target)                    { toast("Email no encontrado","","e"); return; }
    if (target.id===me.id)          { toast("No puedes agregarte a ti mismo","","e"); return; }
    var cur = allCons[me.id]||[];
    if (cur.indexOf(target.id)>-1)  { toast("Ya está en tu red","","i"); return; }
    setAllCons(function(p){ return Object.assign({},p,{[me.id]:[...cur,target.id]}); });
    toast("Contacto agregado", target.name, "s"); setCtQ("");
  }

  // ── WHATSAPP SHARE ────────────────────────────────────────────────────────────
  function shareOne(prod) {
    var si = myInv().stock.find(function(s){ return s.pid===prod.id; });
    var avail = si ? si.avail : 0;
    var nl = "\n";
    var msg = prod.name + nl + "Precio: " + fmtARS(prod.price) + nl +
      (avail>0 ? "Stock: "+avail+" u." + nl : "") + "Escribime para pedirlo!";
    if (prod.photo && navigator.share) {
      try {
        var arr=prod.photo.split(","), mime=arr[0].match(/:(.*?);/)[1], bstr=atob(arr[1]);
        var n=bstr.length, u8=new Uint8Array(n);
        while(n--) u8[n]=bstr.charCodeAt(n);
        var f = new File([u8], (prod.sku||"p")+".jpg", {type:mime});
        navigator.share({title:prod.name, text:msg, files:[f]}).catch(function(){
          navigator.share({title:prod.name, text:msg}).catch(function(){
            window.open("https://wa.me/?text="+encodeURIComponent(msg),"_blank");
          });
        });
        return;
      } catch(e) {}
    }
    if (navigator.share) { navigator.share({title:prod.name,text:msg}).catch(function(){ window.open("https://wa.me/?text="+encodeURIComponent(msg),"_blank"); }); return; }
    window.open("https://wa.me/?text="+encodeURIComponent(msg),"_blank");
  }

  function doShareCatalog() {
    var sel = myStock().filter(function(i){ return shareSel[i.id]; });
    if (!sel.length) { toast("Selecciona al menos un producto","","e"); return; }
    var nl = "\n";
    var lines = sel.map(function(item){
      var p=allProds.find(function(x){ return x.id===item.pid; });
      if (!p) return "";
      return (p.photo?"":p.emoji||"")+" *"+p.name+"*"+nl+"   "+fmtARS(p.price)+" · "+item.avail+" u.";
    }).filter(Boolean);
    var msg = "Mis productos disponibles hoy:"+nl+nl+lines.join(nl+nl)+nl+nl+"Escribime para hacer tu pedido!";
    var photos = [];
    sel.forEach(function(item){
      var p=allProds.find(function(x){ return x.id===item.pid; });
      if (p&&p.photo) {
        try {
          var arr=p.photo.split(","), mime=arr[0].match(/:(.*?);/)[1], bstr=atob(arr[1]);
          var n=bstr.length, u8=new Uint8Array(n);
          while(n--) u8[n]=bstr.charCodeAt(n);
          photos.push(new File([u8],(p.sku||"p")+".jpg",{type:mime}));
        } catch(e) {}
      }
    });
    setShareM(false); setShareSel({});
    if (navigator.share&&photos.length>0) {
      navigator.share({title:"Mis productos",text:msg,files:photos}).catch(function(){
        navigator.share({title:"Mis productos",text:msg}).catch(function(){
          window.open("https://wa.me/?text="+encodeURIComponent(msg),"_blank");
        });
      }); return;
    }
    if (navigator.share) { navigator.share({title:"Mis productos",text:msg}).catch(function(){ window.open("https://wa.me/?text="+encodeURIComponent(msg),"_blank"); }); return; }
    window.open("https://wa.me/?text="+encodeURIComponent(msg),"_blank");
  }

  // ── DERIVED ──────────────────────────────────────────────────────────────────
  var stockFiltered = myStock().filter(function(i){
    var p=allProds.find(function(x){return x.id===i.pid;});
    if (!p) return false;
    var q=srchStock.toLowerCase();
    return !q||p.name.toLowerCase().includes(q)||p.sku.toLowerCase().includes(q)||(p.cat||"").toLowerCase().includes(q);
  });

  var catFiltered = allProds.filter(function(p){
    var q=srchCat.toLowerCase();
    return !q||p.name.toLowerCase().includes(q)||p.sku.toLowerCase().includes(q)||(p.cat||"").toLowerCase().includes(q);
  });

  var conFiltered = myContacts().filter(function(c){
    var q=srchCon.toLowerCase();
    return !q||c.name.toLowerCase().includes(q)||c.email.toLowerCase().includes(q);
  });

  var sendFiltered = myStock().filter(function(i){
    var p=allProds.find(function(x){return x.id===i.pid;});
    if (!p) return false;
    var q=sendSrch.toLowerCase();
    return !q||p.name.toLowerCase().includes(q)||p.sku.toLowerCase().includes(q);
  });

  var qlFiltered = allProds.filter(function(p){
    var q=qlSrch.toLowerCase();
    return !q||p.name.toLowerCase().includes(q)||p.sku.toLowerCase().includes(q);
  });

  var myLogs = allLogs.filter(function(l){ return !l.uid || l.uid===me.id; });
  var logFiltered = myLogs.filter(function(l){
    var q=srchLog.toLowerCase();
    return !q||l.txt.toLowerCase().includes(q)||l.act.toLowerCase().includes(q)||l.who.toLowerCase().includes(q);
  });

  // ─────────────────────────── LOGIN ────────────────────────────────────────
  if (!me) {
    return (
      <div>
        <style>{CSS}</style>
        <div className="login-bg">
          <div className={"lbox"+(authShake?" shake":"")}>
            <div style={{textAlign:"center",marginBottom:16}}>
              <img src="data:image/png;base64,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" alt="StockPro" style={{width:"100%",maxWidth:260,margin:"0 auto 4px",display:"block"}}/>
            </div>

            <div className="ltabs">
              <div className={"ltab"+(authMode==="login"?" on":"")} onClick={function(){setAuthMode("login");setAuthErr("");}}>Iniciar sesión</div>
              <div className={"ltab"+(authMode==="reg"?" on":"")} onClick={function(){setAuthMode("reg");setAuthErr("");}}>Registrarme</div>
            </div>

            {authErr ? <div className="err-box"><Ic n="x" s={14}/>{authErr}</div> : null}

            {authMode==="login" ? (
              <div>
                <div className="fld">
                  <label className="fl">Email</label>
                  <input className="fi" type="email" placeholder="tu@email.com" value={aEmail} onChange={function(e){setAEmail(e.target.value);setAuthErr("");}} onKeyDown={function(e){if(e.key==="Enter")doLogin();}}/>
                </div>
                <div className="fld">
                  <label className="fl">Contraseña</label>
                  <div className="pw-wrap">
                    <input className="fi" style={{paddingRight:42}} type={showPass?"text":"password"} placeholder="••••••" value={aPass} onChange={function(e){setAPass(e.target.value);setAuthErr("");}} onKeyDown={function(e){if(e.key==="Enter")doLogin();}}/>
                    <button className="pw-eye" type="button" onClick={function(){setShowPass(function(v){return !v;});}}><Ic n={showPass?"eyeoff":"eye"} s={16}/></button>
                  </div>
                </div>
                <button className="cta cta-in" onClick={doLogin}><Ic n="check" s={18}/>Entrar</button>
              </div>
            ) : (
              <div>
                <div className="fld"><label className="fl">Nombre completo o nombre del negocio</label><input className="fi" placeholder="Ej: Laura Cosméticos" value={aName} onChange={function(e){setAName(e.target.value);setAuthErr("");}}/></div>
                <div className="fld"><label className="fl">Email</label><input className="fi" type="email" placeholder="tu@email.com" value={aEmail} onChange={function(e){setAEmail(e.target.value);setAuthErr("");}}/></div>
                <div className="fld">
                  <label className="fl">Contraseña</label>
                  <div className="pw-wrap">
                    <input className="fi" style={{paddingRight:42}} type={showPass?"text":"password"} placeholder="Mínimo 4 caracteres" value={aPass} onChange={function(e){setAPass(e.target.value);setAuthErr("");}}/>
                    <button className="pw-eye" type="button" onClick={function(){setShowPass(function(v){return !v;});}}><Ic n={showPass?"eyeoff":"eye"} s={16}/></button>
                  </div>
                </div>
                <div className="fld"><label className="fl">Repetir contraseña</label><input className="fi" type="password" placeholder="Repetir contraseña" value={aPass2} onChange={function(e){setAPass2(e.target.value);setAuthErr("");}}/></div>
                <div className="fld">
                  <label className="fl">Rol</label>
                  <select className="fi fi-sel" value={aRole} onChange={function(e){setARole(e.target.value);}}>
                    {ROLES.map(function(r){ return <option key={r}>{r}</option>; })}
                  </select>
                </div>
                <button className="cta cta-em" onClick={doRegister}><Ic n="plus" s={18}/>Crear cuenta gratis</button>
              </div>
            )}
          </div>
        </div>
        <div className="toast-wrap">{toasts.map(function(t){ return <Toast key={t.id} t={t} remove={rmToast}/>; })}</div>
      </div>
    );
  }

  // ─────────────────────────── APP PRINCIPAL ────────────────────────────────
  var TABS = [
    {id:"stock",   lbl:"Stock",    ico:"box",    dot: myStock().length>0 },
    {id:"cargar",  lbl:"Cargar",   ico:"plus",   dot: false },
    {id:"enviar",  lbl:"Enviar",   ico:"send",   dot: false },
    {id:"importar",lbl:"Importar", ico:"upload", dot: false },
    {id:"catalog", lbl:"Catálogo", ico:"list",   dot: allProds.length>0 },
    {id:"contacts",lbl:"Contactos",ico:"users",  dot: myContacts().length>0 },
  ];

  return (
    <div>
      <style>{CSS}</style>
      <div className="app">

        {/* HEADER */}
        <div className="hdr">
          <Avatar name={me.name} color={me.color} size={38}/>
          <div style={{flex:1,minWidth:0}}>
            <div className="hdr-name" style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{me.name}</div>
            <div className="hdr-role">{me.role}</div>
          </div>
          <div className="saved-dot"><div style={{width:6,height:6,borderRadius:"50%",background:"var(--em)",flexShrink:0}}/> guardado</div>
          {(function(){
            var unread = notifs.filter(function(n){ return n.toId===me.id&&!n.read; }).length;
            var pending = transfers.filter(function(t){ return t.to===me.id&&t.status==="pending"; }).length;
            var total = unread + pending;
            return total>0 ? (
              <div onClick={function(){setTab("contacts");}} style={{background:"var(--cr)",color:"#fff",borderRadius:20,padding:"3px 10px",fontSize:11,fontWeight:800,cursor:"pointer",flexShrink:0,display:"flex",alignItems:"center",gap:4}}>
                🔔 {total}
              </div>
            ) : null;
          })()}
          <button className="ic-btn" style={{marginLeft:6}} onClick={doLogout} title="Salir"><Ic n="logout" s={17}/></button>
        </div>

        {/* MAIN CONTENT */}
        <div className="main">

          {/* ════ STOCK ════ */}
          {tab==="stock" && (
            <div>
              <div className="ph"><div><div className="ph-h">Mi Stock</div><div style={{fontSize:12,color:"var(--t3)",marginTop:2}}>Solo artículos con unidades disponibles</div></div></div>
              <div className="pc">
                {/* WhatsApp banner */}
                <button className="wa-banner" onClick={function(){setShareM(true);setShareSel({});}}>
                  <div className="row g12">
                    <div style={{width:46,height:46,borderRadius:"50%",background:"rgba(255,255,255,.2)",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}><Ic n="wa" s={24}/></div>
                    <div style={{textAlign:"left"}}>
                      <div style={{fontSize:15,fontWeight:800,color:"#fff"}}>Compartir disponible por WhatsApp</div>
                      <div style={{fontSize:11,color:"rgba(255,255,255,.8)",marginTop:2}}>Selecciona productos y envía con fotos</div>
                    </div>
                  </div>
                  <div style={{color:"rgba(255,255,255,.7)"}}><Ic n="send" s={18}/></div>
                </button>

                <SearchBar value={srchStock} onChange={setSrchStock} placeholder="Buscar por nombre, SKU o categoría..."/>

                {/* Stock propio */}
                <div className="card">
                  <div className="card-h">
                    <div className="card-title"><div className="card-ico" style={{background:"var(--am-l)",color:"var(--am-d)"}}><Ic n="box" s={14}/></div>Stock Propio</div>
                    <span className="badge b-am">{stockFiltered.length}</span>
                  </div>
                  {stockFiltered.length===0 ? (
                    <div className="empty">Sin existencias.{srchStock?" Intenta otra búsqueda.":" Carga productos en la tab Cargar."}</div>
                  ) : (
                    <div className="tw"><table>
                      <thead><tr><th></th><th>SKU</th><th>Producto</th><th>Precio</th><th>Disp.</th><th></th></tr></thead>
                      <tbody>
                        {stockFiltered.map(function(item){
                          var p=allProds.find(function(x){return x.id===item.pid;}); if (!p) return null;
                          return (
                            <tr key={item.id} className="tr">
                              <td><ProdThumb prod={p} size={36}/></td>
                              <td><span style={{color:"var(--in-d)",fontFamily:"var(--mf)",fontSize:11,background:"var(--in-l)",padding:"2px 6px",borderRadius:5,fontWeight:600}}>{p.sku}</span></td>
                              <td><div style={{fontWeight:600,fontSize:12}}>{p.name}</div><div style={{fontSize:10,color:"var(--t3)"}}>{p.cat}</div></td>
                              <td><span style={{fontFamily:"var(--mf)",fontWeight:700,fontSize:12}}>{fmtARS(p.price)}</span></td>
                              <td><span style={{fontFamily:"var(--mf)",fontWeight:800,color:"var(--em-d)",fontSize:14}}>{item.avail}</span></td>
                              <td>
                                <div className="row g8" style={{justifyContent:"flex-end",flexWrap:"wrap"}}>
                                  <button className="btn btn-xs b-wa" onClick={function(){shareOne(p);}}><Ic n="wa" s={12}/></button>
                                  <button className="btn btn-xs b-am" onClick={function(){openTx(item);}} disabled={item.avail===0}><Ic n="send" s={11}/>Pasar</button>
                                  <button className="btn btn-xs b-em" onClick={function(){doSell(item.id,"stock");}} disabled={item.avail===0}><Ic n="check" s={11}/>Venta</button>
                                </div>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table></div>
                  )}
                </div>

                {/* Transferencias pendientes de confirmar (recibidas) */}
                {(function(){
                  var pending = transfers.filter(function(t){ return t.to===me.id&&t.status==="pending"; });
                  if (!pending.length) return null;
                  return (
                    <div className="card" style={{border:"2px solid var(--am)",background:"var(--am-l)"}}>
                      <div className="card-h" style={{background:"var(--am-l)"}}>
                        <div className="card-title" style={{color:"var(--am-d)"}}>🔔 Productos para confirmar ({pending.length})</div>
                      </div>
                      <div style={{padding:"8px 12px"}}>
                        {pending.map(function(tx){
                          var p=allProds.find(function(x){return x.id===tx.pid;});
                          return (
                            <div key={tx.id} style={{display:"flex",alignItems:"center",gap:10,padding:"10px 12px",background:"var(--card)",borderRadius:10,marginBottom:8,border:"1px solid var(--brd)"}}>
                              <div style={{fontSize:28}}>{p?p.emoji:"📦"}</div>
                              <div style={{flex:1}}>
                                <div style={{fontWeight:700,fontSize:13}}>{p?p.name:"Producto"}</div>
                                <div style={{fontSize:11,color:"var(--t3)"}}>{tx.qty} u. de {tx.fromName} · {tx.t}</div>
                              </div>
                              <button className="btn b-em" style={{fontSize:12,padding:"8px 12px"}} onClick={function(){confirmTransfer(tx.id);}}>
                                ✓ Confirmar
                              </button>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  );
                })()}

                {/* Enviados en consigna (pendientes) */}
                {(function(){
                  var sent = transfers.filter(function(t){ return t.from===me.id&&t.status==="pending"; });
                  if (!sent.length) return null;
                  return (
                    <div className="card">
                      <div className="card-h">
                        <div className="card-title"><div className="card-ico" style={{background:"var(--am-l)",color:"var(--am-d)"}}>📤</div>Enviados (esperando confirmación)</div>
                        <span className="badge b-am">{sent.length}</span>
                      </div>
                      <div className="tw"><table>
                        <thead><tr><th></th><th>Producto</th><th>Cant.</th><th>Para</th><th>Estado</th></tr></thead>
                        <tbody>
                          {sent.map(function(tx){
                            var p=allProds.find(function(x){return x.id===tx.pid;});
                            var tUser=users.find(function(u){return u.id===tx.to;});
                            return (
                              <tr key={tx.id} className="tr">
                                <td>{p?<div style={{fontSize:22}}>{p.emoji}</div>:null}</td>
                                <td><div style={{fontWeight:600,fontSize:12}}>{p?p.name:""}</div></td>
                                <td><span style={{fontFamily:"var(--mf)",fontWeight:700,color:"var(--am-d)"}}>{tx.qty}</span></td>
                                <td><span style={{fontSize:11}}>{tUser?tUser.name:""}</span></td>
                                <td><span style={{background:"var(--am-t)",color:"var(--am-d)",borderRadius:5,padding:"2px 8px",fontSize:10,fontWeight:700}}>Pendiente</span></td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table></div>
                    </div>
                  );
                })()}

                {/* Consignación */}
                {myCons().length>0 && (
                  <div className="card">
                    <div className="card-h">
                      <div className="card-title"><div className="card-ico" style={{background:"var(--in-l)",color:"var(--in-d)"}}><Ic n="users" s={14}/></div>En Consignación</div>
                      <span className="badge b-in">{myCons().length}</span>
                    </div>
                    <div className="tw"><table>
                      <thead><tr><th></th><th>Producto</th><th>Precio</th><th>Disp.</th><th></th></tr></thead>
                      <tbody>
                        {myCons().map(function(item){
                          var p=allProds.find(function(x){return x.id===item.pid;}); if (!p) return null;
                          return (
                            <tr key={item.id} className="tr">
                              <td><ProdThumb prod={p} size={36}/></td>
                              <td><div style={{fontWeight:600,fontSize:12}}>{p.name}</div><div style={{fontSize:10,color:"var(--t3)"}}>{item.from||"—"}</div></td>
                              <td><span style={{fontFamily:"var(--mf)",fontWeight:700,fontSize:12}}>{fmtARS(p.price)}</span></td>
                              <td><span style={{fontFamily:"var(--mf)",fontWeight:800,color:"var(--am-d)",fontSize:14}}>{item.avail}</span></td>
                              <td>
                                <div className="row g8" style={{justifyContent:"flex-end",flexWrap:"wrap"}}>
                                  <button className="btn btn-xs b-wa" onClick={function(){shareOne(p);}}><Ic n="wa" s={12}/></button>
                                  <button className="btn btn-xs b-cr" onClick={function(){doReturn(item.id);}} disabled={item.avail===0}><Ic n="undo" s={11}/>Dev.</button>
                                  <button className="btn btn-xs b-in" onClick={function(){doSell(item.id,"consign");}} disabled={item.avail===0}><Ic n="check" s={11}/>Venta</button>
                                </div>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table></div>
                  </div>
                )}
              </div>
            </div>
          )}

          {/* ════ CARGAR STOCK ════ */}
          {tab==="cargar" && (
            <div>
              <div className="ph"><div><div className="ph-h">Cargar Stock</div><div style={{fontSize:12,color:"var(--t3)",marginTop:2}}>Agrega unidades o crea productos nuevos</div></div></div>
              <div className="pc">
                <div className="card">
                  <div className="card-h"><div className="card-title">Modo de carga</div></div>
                  <div style={{padding:"14px 14px 0"}}>
                    <div className="ql-tabs">
                      <div className={"ql-tab"+(qlMode==="add"?" on":"")} onClick={function(){setQlMode("add");}}>Agregar a existente</div>
                      <div className={"ql-tab"+(qlMode==="new"?" on":"")} onClick={function(){setQlMode("new");}}>Crear producto nuevo</div>
                    </div>

                    {qlMode==="add" ? (
                      <div>
                        <div style={{fontSize:11,fontWeight:700,color:"var(--t3)",textTransform:"uppercase",letterSpacing:".08em",marginBottom:10}}>Seleccionar producto</div>
                        <SearchBar value={qlSrch} onChange={setQlSrch} placeholder="Buscar producto..."/>
                        <div style={{maxHeight:300,overflowY:"auto",marginBottom:14}}>
                          {qlFiltered.length===0 ? <div className="empty" style={{padding:"20px 0"}}>Sin resultados</div> :
                            qlFiltered.map(function(p){
                              var si=myInv().stock.find(function(s){return s.pid===p.id;});
                              return (
                                <div key={p.id} className={"prod-card"+(qlPid===p.id?" sel":"")} onClick={function(){setQlPid(p.id);}}>
                                  <ProdThumb prod={p} size={40}/>
                                  <div style={{flex:1,minWidth:0}}>
                                    <div style={{fontWeight:700,fontSize:13,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.name}</div>
                                    <div style={{fontSize:11,color:"var(--t3)",marginTop:2}}>Stock: {si?si.avail:0} u. · {fmtARS(p.price)}</div>
                                  </div>
                                  {qlPid===p.id && <div style={{width:22,height:22,borderRadius:"50%",background:"var(--am)",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}><Ic n="check" s={12}/></div>}
                                </div>
                              );
                            })
                          }
                        </div>
                        <div style={{fontSize:11,fontWeight:700,color:"var(--t3)",textTransform:"uppercase",letterSpacing:".08em",marginBottom:8}}>Cantidad a agregar</div>
                        <QtyControl val={qlQty} set={setQlQty} big={true}/>
                        <button className="cta cta-am" onClick={doQuickLoad} disabled={!qlPid}><Ic n="plus" s={18}/>Agregar {qlQty} unidades</button>
                      </div>
                    ) : (
                      <div style={{paddingBottom:16}}>
                        <div className="fld"><label className="fl">Código SKU</label><input className="fi" placeholder="Ej: PERF01" value={qlSku} onChange={function(e){setQlSku(e.target.value);}}/></div>
                        <div className="fld"><label className="fl">Nombre del producto</label><input className="fi" placeholder="Ej: Perfume Lattafa 100ml" value={qlName} onChange={function(e){setQlName(e.target.value);}}/></div>
                        <div className="row g8" style={{marginBottom:12}}>
                          <div style={{flex:1}}><label className="fl">Precio ($)</label><input className="fi" type="number" placeholder="0.00" value={qlPrice} onChange={function(e){setQlPrice(e.target.value);}}/></div>
                          <div style={{flex:1}}><label className="fl">Icono</label><select className="fi fi-sel" value={qlEmoji} onChange={function(e){setQlEmoji(e.target.value);}}>{EMOJIS.map(function(o){return <option key={o.v} value={o.v}>{o.v} {o.l}</option>;})}</select></div>
                        </div>
                        <div className="fld"><label className="fl">Categoría</label><select className="fi fi-sel" value={qlCat} onChange={function(e){setQlCat(e.target.value);}}>{CATS.map(function(c){return <option key={c}>{c}</option>;})}</select></div>
                        <div className="fld">
                          <label className="fl">Foto del producto (opcional)</label>
                          <div className="photo-zone">
                            <input type="file" accept="image/*" onChange={function(e){handlePhoto(e,setQlPhoto);}}/>
                            {qlPhoto ? <img src={qlPhoto} alt="" style={{width:60,height:60,borderRadius:10,objectFit:"cover",margin:"0 auto"}}/> : <div><div style={{fontSize:28,marginBottom:6}}>📸</div><div style={{fontSize:12,fontWeight:700,color:"var(--in)"}}>Tocar para subir foto</div></div>}
                          </div>
                          {qlPhoto && <button onClick={function(){setQlPhoto(null);}} style={{fontSize:11,color:"var(--cr)",background:"none",border:"none",cursor:"pointer",fontWeight:700}}>Quitar foto</button>}
                        </div>
                        <div style={{fontSize:11,fontWeight:700,color:"var(--t3)",textTransform:"uppercase",letterSpacing:".08em",marginBottom:8}}>Stock inicial</div>
                        <QtyControl val={qlQty} set={setQlQty} min={0} big={true}/>
                        <button className="cta cta-am" onClick={doQuickLoad}><Ic n="check" s={18}/>Crear y cargar al stock</button>
                      </div>
                    )}
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* ════ ENVIAR ════ */}
          {tab==="enviar" && (
            <div>
              <div className="ph"><div><div className="ph-h">Enviar Productos</div><div style={{fontSize:12,color:"var(--t3)",marginTop:2}}>Transferencia de stock a contactos</div></div></div>
              <div className="pc">
                {/* Steps */}
                <div className="step-row">
                  <div style={{textAlign:"center",flex:1}}>
                    <div className={"step-dot"+(sendStep>=1?" on":"")} style={sendStep>1?{background:"var(--em)",borderColor:"var(--em)"}:{}}>{sendStep>1?<Ic n="check" s={14}/>:"1"}</div>
                    <div className="step-lbl">Destinatario</div>
                  </div>
                  <div className={"step-line"+(sendStep>=2?" done":"")}/>
                  <div style={{textAlign:"center",flex:1}}>
                    <div className={"step-dot"+(sendStep>=2?" on":"")}>2</div>
                    <div className="step-lbl">Productos</div>
                  </div>
                </div>

                {sendStep===1 && (
                  <div className="card">
                    <div className="card-h"><div className="card-title">¿A quién le envías?</div></div>
                    <div style={{padding:14}}>
                      {myContacts().length===0 ? (
                        <div className="empty">
                          <div style={{fontSize:32,marginBottom:8}}>👥</div>
                          <div style={{fontWeight:700,marginBottom:6}}>Sin contactos</div>
                          <div style={{fontSize:12,marginBottom:14}}>Agrega contactos en la tab Contactos.</div>
                          <button className="btn b-am" onClick={function(){setTab("contacts");}}>Ir a Contactos</button>
                        </div>
                      ) : myContacts().map(function(c){
                        return (
                          <div key={c.id} className={"ct-card"+(sendTo===c.id?" sel":"")} onClick={function(){setSendTo(c.id);setSendStep(2);}}>
                            <Avatar name={c.name} color={c.color} size={46}/>
                            <div style={{flex:1}}>
                              <div style={{fontWeight:700,fontSize:14}}>{c.name}</div>
                              <div style={{fontSize:11,color:"var(--t3)",marginTop:2}}>{c.role}</div>
                            </div>
                            <Ic n="send" s={18}/>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                )}

                {sendStep===2 && (function(){
                  var tUser = users.find(function(u){return u.id===sendTo;});
                  var cartTotal = Object.values(sendCart).reduce(function(s,v){return s+Number(v);},0);
                  return (
                    <div>
                      <div className="row g12" style={{padding:"10px 14px",background:"var(--card)",borderRadius:14,border:"1px solid var(--brd)",marginBottom:14}}>
                        <Avatar name={tUser?tUser.name:"?"} color={tUser?tUser.color:"#999"} size={36}/>
                        <div style={{flex:1}}><div style={{fontWeight:700,fontSize:13}}>{tUser?tUser.name:""}</div><div style={{fontSize:11,color:"var(--t3)"}}>Destinatario</div></div>
                        <button className="btn btn-xs b-ghost" onClick={function(){setSendStep(1);setSendCart({});}}>Cambiar</button>
                      </div>
                      <SearchBar value={sendSrch} onChange={setSendSrch} placeholder="Buscar producto..."/>
                      <div style={{fontSize:11,fontWeight:700,color:"var(--t3)",textTransform:"uppercase",letterSpacing:".08em",marginBottom:8}}>Productos disponibles</div>
                      {sendFiltered.length===0 ? <div className="empty">Sin stock disponible</div> :
                        sendFiltered.map(function(item){
                          var p=allProds.find(function(x){return x.id===item.pid;}); if (!p) return null;
                          var qty=Number(sendCart[item.id]||0);
                          return (
                            <div key={item.id} className="sic">
                              <ProdThumb prod={p} size={38}/>
                              <div style={{flex:1,minWidth:0}}>
                                <div style={{fontWeight:600,fontSize:13,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.name}</div>
                                <div style={{fontSize:11,color:"var(--t3)",marginTop:2}}>{item.avail} disp. · {fmtARS(p.price)}</div>
                              </div>
                              <QtyControl val={qty} set={function(v){setSendCart(function(c){var n=Object.assign({},c);n[item.id]=Math.min(item.avail,v);return n;});}} min={0} max={item.avail}/>
                            </div>
                          );
                        })
                      }
                      {cartTotal>0 && (
                        <div className="cart-box">
                          <div style={{fontSize:11,fontWeight:700,color:"var(--em-d)",marginBottom:6,textTransform:"uppercase",letterSpacing:".08em"}}>Resumen</div>
                          {Object.entries(sendCart).filter(function(e){return Number(e[1])>0;}).map(function(e){
                            var si=myInv().stock.find(function(i){return i.id===e[0];}); var p=si?allProds.find(function(x){return x.id===si.pid;}):null;
                            return p?<div key={e[0]} className="row jb" style={{fontSize:12,color:"var(--t2)",marginBottom:3}}><span>{p.emoji} {p.name}</span><span style={{fontWeight:700}}>{e[1]} u.</span></div>:null;
                          })}
                          <div className="row jb" style={{fontWeight:800,fontSize:14,color:"var(--em-d)",marginTop:8,paddingTop:8,borderTop:"1px solid var(--em-t)"}}><span>Total</span><span>{cartTotal} u.</span></div>
                        </div>
                      )}
                      <button className="cta cta-em" onClick={doMultiSend} disabled={cartTotal===0}><Ic n="send" s={18}/>Confirmar envío ({cartTotal} u.)</button>
                    </div>
                  );
                })()}
              </div>
            </div>
          )}

          {/* ════ IMPORTAR ════ */}
          {tab==="importar" && (
            <div>
              <div className="ph">
                <div><div className="ph-h">Importar</div><div style={{fontSize:12,color:"var(--t3)",marginTop:2}}>Carga tu lista de precios desde Excel</div></div>
                {impRows.some(function(r){return r.ok;}) && <button className="btn b-em" style={{fontSize:12,padding:"8px 14px"}} onClick={doConfirmImport}><Ic n="check" s={14}/>Confirmar ({impRows.filter(function(r){return r.ok;}).length})</button>}
              </div>
              <div className="pc">
                <div className="card">
                  <div className="card-h"><div className="card-title">Modo de importación</div></div>
                  <div style={{padding:"12px 14px 0"}}>
                    <div className="ql-tabs">
                      <div className={"ql-tab"+(impTab==="file"?" on":"")} onClick={function(){setImpTab("file");}}>Subir archivo</div>
                      <div className={"ql-tab"+(impTab==="txt"?" on":"")} onClick={function(){setImpTab("txt");}}>Pegar texto</div>
                    </div>

                    <div className="fld">
                      <label className="fl">Unidades a cargar por producto</label>
                      <QtyControl val={impQty} set={setImpQty} min={1}/>
                    </div>

                    {impTab==="file" ? (
                      <div>
                        <div className="drop-z" style={{marginBottom:14}}>
                          <input type="file" accept=".csv,.txt,.xls,.xlsx,.xlsm" onChange={doParseFile}/>
                          <div style={{fontSize:40,marginBottom:8}}>📊</div>
                          <div style={{fontSize:14,fontWeight:700,color:"var(--in)"}}>Tocar para subir archivo</div>
                          <div style={{fontSize:12,color:"var(--t3)",marginTop:4}}>.xlsx · .xls · .csv · .txt</div>
                          {impFile && <div style={{marginTop:10,background:"var(--in-l)",color:"var(--in-d)",borderRadius:20,padding:"4px 12px",display:"inline-flex",alignItems:"center",gap:6,fontSize:12,fontWeight:700}}><Ic n="check" s={12}/>{impFile}</div>}
                        </div>
                        <div style={{background:"var(--in-l)",border:"1px solid var(--in-t)",borderRadius:10,padding:"12px 14px",marginBottom:14,fontSize:12,color:"var(--in-d)"}}>
                          <div style={{fontWeight:700,marginBottom:8}}>Columnas que detecta automáticamente:</div>
                          <div className="row g8" style={{flexWrap:"wrap"}}>
                            {[["SKU / Código","Obligatorio"],["Nombre","Obligatorio"],["Categoría","Opcional"],["Precio","Opcional"]].map(function(c){
                              return <div key={c[0]} style={{background:"var(--card)",borderRadius:7,padding:"5px 10px",border:"1px solid var(--brd)",fontSize:11}}>
                                <div style={{fontWeight:700}}>{c[0]}</div>
                                <div style={{color:c[1]==="Obligatorio"?"var(--em-d)":"var(--t4)",fontWeight:600,fontSize:10}}>{c[1]}</div>
                              </div>;
                            })}
                          </div>
                        </div>
                      </div>
                    ) : (
                      <div>
                        <div style={{fontSize:12,color:"var(--t3)",marginBottom:10,lineHeight:1.5}}>Formato: <strong>SKU - Nombre TAB Precio</strong> (un producto por línea)</div>
                        <div style={{fontFamily:"var(--mf)",fontSize:11,background:"var(--bg2)",border:"1px solid var(--brd)",borderRadius:8,padding:"10px 12px",marginBottom:10,color:"var(--t2)",lineHeight:1.9}}>
                          {"KCG1 - KALOE Crema Gel x 50gr.\t$ 11.900,00"}<br/>
                          {"BF100 - 212 Rosa x 80 ml.\t$ 40.990,00"}
                        </div>
                        <textarea
                          style={{width:"100%",height:160,padding:"12px 14px",borderRadius:10,border:"1.5px solid var(--brd)",background:"var(--card)",color:"var(--t1)",fontFamily:"var(--mf)",fontSize:12,resize:"none",outline:"none",lineHeight:1.7}}
                          placeholder={"KCG1 - KALOE Crema Gel x 50gr.\t$ 11.900,00"}
                          value={impTxt} onChange={function(e){setImpTxt(e.target.value);setImpRows([]);}}
                        />
                        <button className="btn b-in" style={{marginTop:8,width:"100%",justifyContent:"center",padding:12}} onClick={doParseTxt} disabled={!impTxt.trim()}><Ic n="search" s={14}/>Previsualizar</button>
                      </div>
                    )}

                    {impRows.length>0 && (
                      <div style={{marginTop:14,marginBottom:14}}>
                        <div style={{fontSize:11,fontWeight:700,color:"var(--t3)",textTransform:"uppercase",marginBottom:8}}>{impRows.filter(function(r){return r.ok;}).length} OK · {impRows.filter(function(r){return !r.ok;}).length} errores</div>
                        {impRows.slice(0,20).map(function(row,i){
                          return (
                            <div key={i} className={"prev-row "+(row.ok?"prev-ok":"prev-err")}>
                              <div style={{width:7,height:7,borderRadius:"50%",background:row.ok?"var(--em)":"var(--cr)",flexShrink:0}}/>
                              {row.ok ? (
                                <div className="row g8" style={{flex:1,flexWrap:"wrap"}}>
                                  <span style={{fontFamily:"var(--mf)",fontSize:11,fontWeight:700,color:"var(--in-d)"}}>{row.sku}</span>
                                  <span style={{flex:1,fontSize:12}}>{row.name}</span>
                                  <span style={{fontFamily:"var(--mf)",fontSize:11,fontWeight:700,color:"var(--am-d)"}}>{fmtARS(row.price)}</span>
                                </div>
                              ) : (
                                <div className="row g8" style={{flex:1}}>
                                  <span style={{flex:1,fontFamily:"var(--mf)",fontSize:11,color:"var(--t3)",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{row.raw}</span>
                                  <span style={{fontSize:11,color:"var(--cr)",flexShrink:0}}>{row.err}</span>
                                </div>
                              )}
                            </div>
                          );
                        })}
                        {impRows.length>20 && <div style={{textAlign:"center",fontSize:11,color:"var(--t3)",padding:"8px 0"}}>...y {impRows.length-20} más</div>}
                        <button className="cta cta-em" onClick={doConfirmImport} disabled={!impRows.some(function(r){return r.ok;})}><Ic n="check" s={18}/>Importar {impRows.filter(function(r){return r.ok;}).length} productos</button>
                      </div>
                    )}
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* ════ CATÁLOGO ════ */}
          {tab==="catalog" && (
            <div>
              <div className="ph">
                <div><div className="ph-h">Catálogo</div><div style={{fontSize:12,color:"var(--t3)",marginTop:2}}>{allProds.length} productos · ABM completo</div></div>
                <button className="btn btn-xs b-em" onClick={function(){setTab("importar");}}><Ic n="upload" s={13}/>Importar</button>
              </div>
              <div className="pc">
                <SearchBar value={srchCat} onChange={setSrchCat} placeholder="Buscar por nombre, SKU o categoría..."/>
                <div className="card">
                  <div className="card-h">
                    <div className="card-title">{editP?"✏️ Modificar":"✨ Nuevo producto"}</div>
                    {editP && <button className="btn btn-xs b-ghost" onClick={cancelEdit}>Cancelar</button>}
                  </div>
                  <div style={{padding:"12px 14px"}}>
                    <form onSubmit={doSaveProd}>
                      <div className="row g8" style={{marginBottom:12}}>
                        <div style={{flex:1}}><label className="fl">SKU</label><input className="fi" placeholder="Ej: KCG1" value={fSku} onChange={function(e){setFSku(e.target.value);}}/></div>
                        <div style={{flex:2}}><label className="fl">Nombre</label><input className="fi" placeholder="Ej: KALOE Crema Gel..." value={fName} onChange={function(e){setFName(e.target.value);}}/></div>
                      </div>
                      <div className="row g8" style={{marginBottom:12}}>
                        <div style={{flex:1}}><label className="fl">Precio ($)</label><input className="fi" type="number" step="0.01" placeholder="0.00" value={fPrice} onChange={function(e){setFPrice(e.target.value);}}/></div>
                        <div style={{flex:1}}><label className="fl">Icono</label><select className="fi fi-sel" value={fEmoji} onChange={function(e){setFEmoji(e.target.value);}}>{EMOJIS.map(function(o){return <option key={o.v} value={o.v}>{o.v} {o.l}</option>;})}</select></div>
                      </div>
                      <div className="fld"><label className="fl">Categoría</label><select className="fi fi-sel" value={fCat} onChange={function(e){setFCat(e.target.value);}}>{CATS.map(function(c){return <option key={c}>{c}</option>;})}</select></div>
                      <div className="fld">
                        <label className="fl">Foto real (opcional)</label>
                        <div className="photo-zone">
                          <input type="file" accept="image/*" onChange={function(e){handlePhoto(e,setFPhoto);}}/>
                          {fPhoto ? <img src={fPhoto} alt="" style={{width:56,height:56,borderRadius:10,objectFit:"cover",margin:"0 auto"}}/> : <div><div style={{fontSize:24,marginBottom:4}}>📸</div><div style={{fontSize:12,fontWeight:700,color:"var(--in)"}}>Tocar para subir foto</div></div>}
                        </div>
                        {fPhoto && <button type="button" onClick={function(){setFPhoto(null);}} style={{fontSize:11,color:"var(--cr)",background:"none",border:"none",cursor:"pointer",fontWeight:700}}>Quitar foto</button>}
                      </div>
                      {!editP && <div className="fld"><label className="fl">Stock inicial</label><input className="fi" type="number" min="0" value={fStock} onChange={function(e){setFStock(e.target.value);}}/></div>}
                      <button type="submit" className="cta cta-in" style={{marginTop:10}}><Ic n="check" s={18}/>{editP?"Guardar cambios":"Dar de alta"}</button>
                    </form>
                  </div>
                </div>

                <div className="card">
                  <div className="card-h"><div className="card-title">Catálogo indexado ({catFiltered.length})</div></div>
                  {catFiltered.length===0 ? <div className="empty">Sin productos.{srchCat?" Intenta otra búsqueda.":" Crea el primero arriba o importa desde Excel."}</div> : (
                    <div className="tw"><table>
                      <thead><tr><th></th><th>SKU</th><th>Producto</th><th>Precio</th><th></th></tr></thead>
                      <tbody>
                        {catFiltered.map(function(p){
                          return (
                            <tr key={p.id} className="tr">
                              <td><ProdThumb prod={p} size={34}/></td>
                              <td><span style={{color:"var(--in-d)",fontFamily:"var(--mf)",fontSize:11,background:"var(--in-l)",padding:"2px 6px",borderRadius:5,fontWeight:600}}>{p.sku}</span></td>
                              <td><div style={{fontWeight:600,fontSize:12}}>{p.name}</div><div style={{fontSize:10,color:"var(--t3)"}}>{p.cat}</div></td>
                              <td style={{fontFamily:"var(--mf)",fontSize:12,fontWeight:700}}>{fmtARS(p.price)}</td>
                              <td>
                                <div className="row g8" style={{justifyContent:"flex-end"}}>
                                  <button className="btn btn-xs b-in" onClick={function(){startEdit(p);}}><Ic n="edit" s={12}/></button>
                                  {delConf===p.id
                                    ? <div className="row g8"><button className="btn btn-xs b-cr" onClick={function(){doDelProd(p.id,p.sku);}}>Sí</button><button className="btn btn-xs b-ghost" onClick={function(){setDelConf(null);}}>No</button></div>
                                    : <button className="btn btn-xs b-cr" onClick={function(){setDelConf(p.id);}}><Ic n="trash" s={12}/></button>
                                  }
                                </div>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table></div>
                  )}
                </div>
              </div>
            </div>
          )}

          {/* ════ CONTACTOS ════ */}
          {tab==="contacts" && (
            <div>
              <div className="ph"><div><div className="ph-h">Contactos</div><div style={{fontSize:12,color:"var(--t3)",marginTop:2}}>Tu red de distribución</div></div></div>
              <div className="pc">
                <div className="card">
                  <div className="card-h"><div className="card-title"><div className="card-ico" style={{background:"var(--em-l)",color:"var(--em-d)"}}><Ic n="users" s={14}/></div>Agregar contacto</div></div>
                  <div style={{padding:"12px 14px"}}>
                    <div style={{fontSize:12,color:"var(--t3)",marginBottom:10,lineHeight:1.5}}>Ingresa el <strong>email exacto</strong> de la persona que quieres agregar. Esa persona debe estar registrada en la app.</div>
                    <div className="row g8">
                      <input className="fi" style={{flex:1}} type="email" placeholder="email@ejemplo.com" value={ctQ} onChange={function(e){setCtQ(e.target.value);}} onKeyDown={function(e){if(e.key==="Enter")doAddContact();}}/>
                      <button className="btn b-pri" onClick={doAddContact}><Ic n="plus" s={14}/>Agregar</button>
                    </div>
                  </div>
                </div>

                <SearchBar value={srchCon} onChange={setSrchCon} placeholder="Buscar contacto..."/>

                <div className="card">
                  <div className="card-h"><div className="card-title">Tu red ({conFiltered.length})</div></div>
                  {conFiltered.length===0 ? <div className="empty">{srchCon?"Sin resultados.":"Aún no tienes contactos. Agrega el email de alguien registrado."}</div> : (
                    <div style={{padding:"10px 14px"}}>
                      {conFiltered.map(function(c){
                        return (
                          <div key={c.id} className="ct-card" style={{cursor:"default"}}>
                            <Avatar name={c.name} color={c.color} size={44}/>
                            <div style={{flex:1}}>
                              <div style={{fontWeight:700,fontSize:14}}>{c.name}</div>
                              <div style={{fontSize:11,color:"var(--t3)",marginTop:2}}>{c.email} · {c.role}</div>
                            </div>
                            <div style={{background:"var(--em-l)",color:"var(--em-d)",borderRadius:20,padding:"3px 9px",fontSize:11,fontWeight:700,border:"1px solid var(--em-t)",display:"flex",alignItems:"center",gap:4}}><Ic n="check" s={11}/>OK</div>
                          </div>
                        );
                      })}
                    </div>
                  )}
                </div>

                {/* Notificaciones */}
                {(function(){
                  var myNotifs = notifs.filter(function(n){ return n.toId===me.id; }).slice(0,20);
                  if (!myNotifs.length) return null;
                  return (
                    <div className="card" style={{marginTop:4,border: myNotifs.some(function(n){return !n.read;}) ? "2px solid var(--am)":"1px solid var(--brd)"}}>
                      <div className="card-h">
                        <div className="card-title"><div className="card-ico" style={{background:"var(--am-l)",color:"var(--am-d)"}}>🔔</div>Notificaciones</div>
                        <div style={{display:"flex",gap:8,alignItems:"center"}}>
                          {myNotifs.some(function(n){return !n.read;}) && <span className="badge b-am">{myNotifs.filter(function(n){return !n.read;}).length} nuevas</span>}
                          <button className="btn btn-xs b-ghost" onClick={function(){setNotifs(function(p){ return p.map(function(n){ return n.toId===me.id?Object.assign({},n,{read:true}):n; }); });}}>Marcar leídas</button>
                        </div>
                      </div>
                      <div style={{padding:"8px 14px"}}>
                        {myNotifs.map(function(n){
                          var ico = n.type==="transfer"?"📦":n.type==="confirm"?"✅":n.type==="sale"?"💰":"🔔";
                          return (
                            <div key={n.id} style={{display:"flex",gap:10,padding:"10px 0",borderBottom:"1px solid var(--brd)",opacity:n.read?0.6:1}}>
                              <div style={{fontSize:20,flexShrink:0}}>{ico}</div>
                              <div style={{flex:1}}>
                                <div style={{fontSize:13,fontWeight:n.read?400:700,color:"var(--t1)"}}>{n.msg}</div>
                                <div style={{fontSize:10,color:"var(--t3)",marginTop:2}}>{n.t}</div>
                              </div>
                              {!n.read && <div style={{width:8,height:8,borderRadius:"50%",background:"var(--am)",flexShrink:0,marginTop:4}}/>}
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  );
                })()}

                {/* Historial aquí abajo */}
                <div className="card" style={{marginTop:4}}>
                  <div className="card-h"><div className="card-title"><div className="card-ico" style={{background:"var(--in-l)",color:"var(--in-d)"}}><Ic n="clock" s={14}/></div>Mi historial</div><span className="badge b-in">{myLogs.length}</span></div>
                  <div style={{padding:"0 0 4px"}}>
                    <div style={{padding:"8px 14px 4px"}}><SearchBar value={srchLog} onChange={setSrchLog} placeholder="Buscar en historial..."/></div>
                    {logFiltered.length===0 ? <div className="empty">Sin movimientos</div> : (
                      <div className="tw"><table>
                        <thead><tr><th>Hora</th><th>Quién</th><th>Acción</th><th>Detalle</th></tr></thead>
                        <tbody>
                          {logFiltered.slice(0,50).map(function(l){
                            return (
                              <tr key={l.id} className="tr">
                                <td style={{fontFamily:"var(--mf)",fontSize:10,color:"var(--t3)",whiteSpace:"nowrap"}}>{l.t}</td>
                                <td style={{fontWeight:600,fontSize:12,whiteSpace:"nowrap"}}>{l.who}</td>
                                <td><span style={{background:"var(--in-l)",color:"var(--in-d)",borderRadius:5,padding:"2px 7px",fontSize:10,fontWeight:700,fontFamily:"var(--mf)"}}>{l.act}</span></td>
                                <td style={{fontSize:12,color:"var(--t2)"}}>{l.txt}</td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table></div>
                    )}
                  </div>
                </div>
              </div>
            </div>
          )}

        </div>

        {/* TAB BAR */}
        <nav className="tabbar">
          {TABS.map(function(t){
            return (
              <div key={t.id} className={"tab"+(tab===t.id?" on":"")} onClick={function(){setTab(t.id);}}>
                {t.dot && tab!==t.id ? <span className="tab-dot"/> : null}
                <div className="tab-bubble"><Ic n={t.ico} s={20}/></div>
                <span className="tab-lbl">{t.lbl}</span>
              </div>
            );
          })}
        </nav>

      </div>

      {/* MODAL: TRANSFERENCIA RÁPIDA */}
      {txModal && (
        <div className="ovl" onClick={function(e){if(e.target===e.currentTarget)setTxModal(null);}}>
          <div className="mbox">
            <div className="mhd"><div className="mhd-t">Transferir Existencias</div><button className="ic-btn" onClick={function(){setTxModal(null);}}><Ic n="x" s={16}/></button></div>
            <div className="mbd">
              {(function(){
                var prod=allProds.find(function(p){return p.id===txModal.pid;});
                return (
                  <div>
                    <div className="row g12" style={{padding:"12px 14px",background:"var(--bg2)",borderRadius:10,border:"1px solid var(--brd)",marginBottom:14}}>
                      <ProdThumb prod={prod} size={44}/>
                      <div style={{flex:1}}><div style={{fontWeight:700,fontSize:14}}>{prod?prod.name:""}</div><div style={{fontSize:11,color:"var(--t3)",marginTop:2}}>{prod?prod.sku:""} · {fmtARS(prod?prod.price:0)}</div></div>
                      <div style={{fontFamily:"var(--mf)",fontWeight:800,fontSize:15,color:"var(--in)"}}>{fmtARS(prod?(prod.price*txQty):0)}</div>
                    </div>
                    <div className="fld"><label className="fl">Destinatario</label>
                      <select className="fi fi-sel" value={txTo} onChange={function(e){setTxTo(e.target.value);}}>
                        {myContacts().map(function(c){return <option key={c.id} value={c.id}>{c.name}</option>;})}
                      </select>
                    </div>
                    <div className="fld"><label className="fl">Cantidad ({txModal.avail} disponibles)</label><QtyControl val={txQty} set={setTxQty} max={txModal.avail} big={true}/></div>
                  </div>
                );
              })()}
            </div>
            <div className="mft"><button className="btn b-ghost" onClick={function(){setTxModal(null);}}>Cancelar</button><button className="btn b-pri" onClick={doTx}><Ic n="send" s={14}/>Confirmar</button></div>
          </div>
        </div>
      )}

      {/* MODAL: COMPARTIR CATÁLOGO */}
      {shareM && (
        <div className="ovl" onClick={function(e){if(e.target===e.currentTarget){setShareM(false);setShareSel({});}}}>
          <div className="mbox">
            <div className="mhd">
              <div className="mhd-t row g8"><Ic n="wa" s={20}/>Compartir por WhatsApp</div>
              <button className="ic-btn" onClick={function(){setShareM(false);setShareSel({});}}><Ic n="x" s={16}/></button>
            </div>
            <div className="mbd">
              <div style={{fontSize:13,color:"var(--t2)",marginBottom:12,lineHeight:1.5}}>Selecciona los productos para incluir en el mensaje. Si tienen foto, se enviarán con imagen.</div>
              <div className="row jb" style={{marginBottom:12}}>
                <span style={{fontSize:12,fontWeight:700,color:"var(--t3)"}}>{Object.keys(shareSel).length} seleccionados</span>
                <button className="btn btn-xs b-in" onClick={function(){var all={};myStock().forEach(function(i){all[i.id]=true;});setShareSel(all);}}>Seleccionar todo</button>
              </div>
              <div style={{maxHeight:360,overflowY:"auto",display:"flex",flexDirection:"column",gap:8}}>
                {myStock().map(function(item){
                  var p=allProds.find(function(x){return x.id===item.pid;}); if (!p) return null;
                  var sel=!!shareSel[item.id];
                  return (
                    <div key={item.id} onClick={function(){setShareSel(function(prev){var n=Object.assign({},prev);if(n[item.id])delete n[item.id];else n[item.id]=true;return n;});}}
                      style={{display:"flex",alignItems:"center",gap:12,padding:"12px 14px",borderRadius:12,border:"2px solid "+(sel?"var(--wa)":"var(--brd)"),background:sel?"var(--wa-l)":"var(--card)",cursor:"pointer",transition:"all .15s"}}>
                      <ProdThumb prod={p} size={44}/>
                      <div style={{flex:1,minWidth:0}}>
                        <div style={{fontWeight:700,fontSize:13,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.name}</div>
                        <div style={{fontSize:11,color:"var(--t3)",marginTop:2}}>{fmtARS(p.price)} · {item.avail} u.{p.photo?" · 📸":""}</div>
                      </div>
                      <div style={{width:24,height:24,borderRadius:"50%",background:sel?"var(--wa)":"var(--bg2)",border:"2px solid "+(sel?"var(--wa)":"var(--brd2)"),display:"flex",alignItems:"center",justifyContent:"center",color:"#fff",flexShrink:0}}>
                        {sel?<Ic n="check" s={13}/>:null}
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
            <div className="mft">
              <button className="btn b-ghost" onClick={function(){setShareM(false);setShareSel({});}}>Cancelar</button>
              <button style={{background:"linear-gradient(135deg,var(--wa),var(--wa-d))",color:"#fff",border:"none",borderRadius:10,padding:"12px 18px",fontWeight:800,fontSize:14,cursor:"pointer",display:"flex",alignItems:"center",gap:8,boxShadow:"0 4px 14px rgba(37,211,102,.35)",opacity:Object.keys(shareSel).length===0?0.4:1}} onClick={doShareCatalog} disabled={Object.keys(shareSel).length===0}>
                <Ic n="wa" s={16}/>Compartir con fotos
              </button>
            </div>
          </div>
        </div>
      )}

      {/* TOASTS */}
      <div className="toast-wrap">{toasts.map(function(t){ return <Toast key={t.id} t={t} remove={rmToast}/>; })}</div>
    </div>
  );
}

