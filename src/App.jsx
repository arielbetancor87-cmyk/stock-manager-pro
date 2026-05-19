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

  useEffect(function(){ DB.set("users",    users);    }, [users]);
  useEffect(function(){ DB.set("prods",    allProds); }, [allProds]);
  useEffect(function(){ DB.set("inv",      allInv);   }, [allInv]);
  useEffect(function(){ DB.set("contacts", allCons);  }, [allCons]);
  useEffect(function(){ DB.set("logs",     allLogs);  }, [allLogs]);

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
    var entry = {id:uid(), who:me.name, act:act, txt:txt, t:new Date().toLocaleTimeString("es-AR")};
    setAllLogs(function(p){ return [entry, ...p.slice(0,199)]; });
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
    setAllInv(function(prev){
      var sec = prev[me.id][section].map(function(it){
        if (it.id!==itemId || it.avail<=0) return it;
        prod = allProds.find(function(p){ return p.id===it.pid; });
        return Object.assign({}, it, {avail:it.avail-1, sold:it.sold+1});
      });
      var u = Object.assign({}, prev[me.id], {[section]:sec});
      return Object.assign({}, prev, {[me.id]:u});
    });
    setTimeout(function(){
      if (prod) { addLog("Venta", "1x "+prod.name); toast("Venta registrada", prod.name, "s"); }
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
    setAllInv(function(prev){
      var src = Object.assign({},prev[me.id]);
      src.stock = src.stock.map(function(it){ return it.id===txModal.id ? Object.assign({},it,{avail:it.avail-txQty}) : it; });
      var dst = prev[txTo]||{stock:[],consign:[]};
      var idx = dst.stock.findIndex(function(s){ return s.pid===txModal.pid; });
      var dstock = idx>-1
        ? dst.stock.map(function(s,i){ return i===idx?Object.assign({},s,{avail:s.avail+Number(txQty)}):s; })
        : [...dst.stock,{id:uid(),pid:txModal.pid,avail:Number(txQty),sold:0}];
      return Object.assign({},prev,{[me.id]:src,[txTo]:Object.assign({},dst,{stock:dstock})});
    });
    addLog("Transferencia", txQty+"x "+prod.name+" → @"+tUser.name);
    toast("Enviado!", txQty+"x "+prod.name+" a "+tUser.name, "s");
    setTxModal(null);
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

  var logFiltered = allLogs.filter(function(l){
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
              <div style={{fontSize:36,marginBottom:6}}>💎</div>
              <div style={{fontSize:22,fontWeight:800,color:"var(--in)",letterSpacing:"-.02em"}}>Stock Manager Pro</div>
              <div style={{fontSize:12,color:"var(--t3)",marginTop:3}}>Gestión de stock y consignaciones</div>
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

                {/* Historial aquí abajo */}
                <div className="card" style={{marginTop:4}}>
                  <div className="card-h"><div className="card-title"><div className="card-ico" style={{background:"var(--in-l)",color:"var(--in-d)"}}><Ic n="clock" s={14}/></div>Historial</div><span className="badge b-in">{allLogs.length}</span></div>
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

