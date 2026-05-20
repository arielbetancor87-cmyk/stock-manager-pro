import { useState, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";
import { createClient } from "@supabase/supabase-js";

// ─── SUPABASE ─────────────────────────────────────────────────────────────────
const SUPA_URL  = "https://cgqrgyouunxfluujhwey.supabase.co";
const SUPA_ANON = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImNncXJneW91dW54Zmx1dWpod2V5Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzkxNzk1NTgsImV4cCI6MjA5NDc1NTU1OH0.4n0zTCGUau02Xz677RwtaqcOUi_7HwvogerqNf3rWkY";
const sb = createClient(SUPA_URL, SUPA_ANON);

// ─── CONSTANTES ───────────────────────────────────────────────────────────────
var ROLES  = ["user","distributor","reseller"];
var CATS   = ["General","Linea Kaloe","Pocket x 10ml","Beauty Collagen","Premium Femenino","Exhibidores","Joyas","Cosmetica"];
var EMOJIS = [{v:"✨",l:"Brillo"},{v:"🧴",l:"Crema"},{v:"💼",l:"Maletin"},{v:"💄",l:"Maquillaje"},{v:"💎",l:"Joya"},{v:"📿",l:"Cadena"},{v:"🌸",l:"Floral"},{v:"🎀",l:"Accesorio"},{v:"🍃",l:"Natural"},{v:"⭐",l:"Destacado"}];

function uid() { return Date.now().toString(36) + Math.random().toString(36).slice(2); }
function fmtARS(n) { return "$ " + Number(n||0).toLocaleString("es-AR",{minimumFractionDigits:2}); }
function ini(name) { return (name||"?").trim().split(" ").slice(0,2).map(function(w){return w[0];}).join("").toUpperCase(); }

// ─── LOGO ─────────────────────────────────────────────────────────────────────
var LOGO_URL = "https://cgqrgyouunxfluujhwey.supabase.co/storage/v1/object/public/assets/logo.png";

// ─── ICONS ───────────────────────────────────────────────────────────────────
function Ic(props) {
  var n=props.n, s=props.s||18;
  var p={width:s,height:s,viewBox:"0 0 24 24",fill:"none",stroke:"currentColor",strokeWidth:"2",strokeLinecap:"round",strokeLinejoin:"round"};
  switch(n) {
    case "box":     return <svg {...p}><polyline points="21,8 21,21 3,21 3,8"/><rect x="1" y="3" width="22" height="5"/></svg>;
    case "plus":    return <svg {...p}><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>;
    case "send":    return <svg {...p}><line x1="22" y1="2" x2="11" y2="13"/><polygon points="22,2 15,22 11,13 2,9 22,2"/></svg>;
    case "upload":  return <svg {...p}><polyline points="16,16 12,12 8,16"/><line x1="12" y1="12" x2="12" y2="21"/><path d="M20.39 18.39A5 5 0 0 0 18 9h-1.26A8 8 0 1 0 3 16.3"/></svg>;
    case "list":    return <svg {...p}><line x1="8" y1="6" x2="21" y2="6"/><line x1="8" y1="12" x2="21" y2="12"/><line x1="8" y1="18" x2="21" y2="18"/><line x1="3" y1="6" x2="3.01" y2="6"/><line x1="3" y1="12" x2="3.01" y2="12"/><line x1="3" y1="18" x2="3.01" y2="18"/></svg>;
    case "users":   return <svg {...p}><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>;
    case "clock":   return <svg {...p}><circle cx="12" cy="12" r="10"/><polyline points="12,6 12,12 16,14"/></svg>;
    case "check":   return <svg {...p} strokeWidth="2.5"><polyline points="20,6 9,17 4,12"/></svg>;
    case "x":       return <svg {...p} strokeWidth="2.5"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>;
    case "edit":    return <svg {...p}><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4Z"/></svg>;
    case "trash":   return <svg {...p}><polyline points="3,6 5,6 21,6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>;
    case "eye":     return <svg {...p}><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg>;
    case "eyeoff":  return <svg {...p}><path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94"/><path d="M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19"/><line x1="1" y1="1" x2="23" y2="23"/></svg>;
    case "search":  return <svg {...p}><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>;
    case "minus":   return <svg {...p}><line x1="5" y1="12" x2="19" y2="12"/></svg>;
    case "logout":  return <svg {...p}><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/><polyline points="16,17 21,12 16,7"/><line x1="21" y1="12" x2="9" y2="12"/></svg>;
    case "undo":    return <svg {...p}><polyline points="1,4 1,10 7,10"/><path d="M3.51 15a9 9 0 1 0 2.13-9.36L1 10"/></svg>;
    case "chart":   return <svg {...p}><line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/><line x1="2" y1="20" x2="22" y2="20"/></svg>;
    case "shield":  return <svg {...p}><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/></svg>;
    case "bell":    return <svg {...p}><path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"/><path d="M13.73 21a2 2 0 0 1-3.46 0"/></svg>;
    case "img":     return <svg {...p}><rect x="3" y="3" width="18" height="18" rx="2"/><circle cx="8.5" cy="8.5" r="1.5"/><polyline points="21,15 16,10 5,21"/></svg>;
    case "wa":      return <svg {...p} fill="currentColor" stroke="none" viewBox="0 0 24 24"><path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347zm-5.421 7.403h-.004a9.87 9.87 0 0 1-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 0 1-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 0 1 2.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884zm8.413-18.297A11.815 11.815 0 0 0 12.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 0 0 5.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 0 0-3.48-8.413z"/></svg>;
    default:        return null;
  }
}

// ─── CSS ─────────────────────────────────────────────────────────────────────
const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;600;700&display=swap');
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{
  --bg:#f4f5fa;--bg2:#eceef5;--card:#fff;--card2:#f7f8fc;
  --brd:#e2e5ef;--brd2:#d0d4e4;
  --t1:#1a1d2e;--t2:#4b5068;--t3:#8b90a8;--t4:#b8bdd0;
  --in:#6366f1;--in-d:#4f46e5;--in-l:#eef2ff;--in-t:rgba(99,102,241,.1);
  --em:#10b981;--em-d:#059669;--em-l:#ecfdf5;--em-t:rgba(16,185,129,.1);
  --am:#f59e0b;--am-d:#d97706;--am-l:#fffbeb;--am-t:rgba(245,158,11,.1);
  --cr:#ef4444;--cr-d:#dc2626;--cr-l:#fef2f2;--cr-t:rgba(239,68,68,.08);
  --wa:#25d366;--wa-d:#128c7e;--wa-l:#dcfce7;--wa-t:rgba(37,211,102,.1);
  --navy:#1e3a5f;
  --hf:'Plus Jakarta Sans',sans-serif;--mf:'JetBrains Mono',monospace;
  --r:10px;--r2:14px;--r3:18px;
  --sh:0 1px 3px rgba(26,29,46,.06),0 2px 8px rgba(26,29,46,.06);
  --sh2:0 4px 20px rgba(26,29,46,.1);--sh3:0 8px 32px rgba(26,29,46,.14);
  --tab:64px;
}
html,body{height:100%;background:var(--bg);color:var(--t1);font-family:var(--hf);font-size:14px;-webkit-font-smoothing:antialiased}
@keyframes fadeUp{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:translateY(0)}}
@keyframes zoomIn{from{opacity:0;transform:scale(.96)}to{opacity:1;transform:scale(1)}}
@keyframes toastIn{from{opacity:0;transform:translateY(14px) scale(.97)}to{opacity:1;transform:translateY(0) scale(1)}}
@keyframes shake{0%,100%{transform:translateX(0)}20%,60%{transform:translateX(-7px)}40%,80%{transform:translateX(7px)}}
@keyframes spin{to{transform:rotate(360deg)}}
@keyframes pulse{0%,100%{opacity:1}50%{opacity:.4}}

.app{display:flex;flex-direction:column;height:100vh;overflow:hidden}
.spin{animation:spin 1s linear infinite}
.pulse{animation:pulse 1.5s ease infinite}

/* TABBAR */
.main{flex:1;overflow-y:auto;padding-bottom:var(--tab)}
.tabbar{position:fixed;bottom:0;left:0;right:0;height:var(--tab);background:var(--card);border-top:1px solid var(--brd);display:flex;z-index:50;box-shadow:0 -4px 16px rgba(26,29,46,.06)}
.tab{flex:1;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:3px;cursor:pointer;color:var(--t3);transition:color .15s;position:relative;padding:4px 2px 2px}
.tab:active{background:var(--bg2)}
.tab.on{color:var(--in)}
.tab-bub{width:36px;height:26px;border-radius:13px;display:flex;align-items:center;justify-content:center;transition:background .15s}
.tab.on .tab-bub{background:var(--in-t)}
.tab-lbl{font-size:10px;font-weight:600}
.tab-dot{position:absolute;top:5px;right:calc(50% - 20px);width:7px;height:7px;border-radius:50%;background:var(--cr);border:2px solid var(--card)}

/* HEADER */
.hdr{display:flex;align-items:center;gap:10px;padding:11px 14px;background:var(--card);border-bottom:1px solid var(--brd);flex-shrink:0}
.hdr-name{font-size:15px;font-weight:800;color:var(--t1);letter-spacing:-.02em}
.hdr-role{font-size:10px;color:var(--t3);margin-top:1px}
.saved-dot{display:flex;align-items:center;gap:4px;background:var(--em-l);border:1px solid var(--em-t);border-radius:20px;padding:3px 9px;font-size:10px;font-weight:700;color:var(--em-d)}
.notif-badge{background:var(--cr);color:#fff;border-radius:20px;padding:3px 10px;font-size:11px;font-weight:800;cursor:pointer;display:flex;align-items:center;gap:4px;flex-shrink:0}

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
.b-in{background:var(--in-t);color:var(--in-d)}.b-em{background:var(--em-t);color:var(--em-d)}
.b-am{background:var(--am-t);color:var(--am-d)}.b-cr{background:var(--cr-t);color:var(--cr)}.b-sl{background:var(--bg2);color:var(--t2)}

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
.fi{width:100%;padding:11px 13px;border-radius:var(--r);border:1.5px solid var(--brd);background:var(--card);color:var(--t1);font-family:var(--hf);font-size:14px;outline:none;transition:all .13s}
.fi:focus{border-color:var(--in);box-shadow:0 0 0 3px var(--in-t)}
.fi-sel{appearance:none;cursor:pointer}
.pw-wrap{position:relative}
.pw-eye{position:absolute;right:11px;top:50%;transform:translateY(-50%);background:none;border:none;cursor:pointer;color:var(--t3);display:flex;padding:4px}

/* BUTTONS */
.btn{display:inline-flex;align-items:center;gap:6px;padding:8px 14px;border-radius:var(--r);border:none;cursor:pointer;font-family:var(--hf);font-size:12px;font-weight:700;transition:all .14s;white-space:nowrap;flex-shrink:0}
.btn:disabled{opacity:.3;cursor:not-allowed}
.btn:active:not(:disabled){transform:scale(.97)}
.b-pri{background:linear-gradient(135deg,var(--in),var(--in-d));color:#fff;box-shadow:0 2px 8px rgba(99,102,241,.3)}
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
.mbox{background:var(--card);border-radius:var(--r3) var(--r3) 0 0;width:100%;max-width:600px;max-height:92vh;overflow-y:auto;box-shadow:var(--sh3);animation:fadeUp .3s cubic-bezier(.16,1,.3,1)}
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
.ct-card{display:flex;align-items:center;gap:12px;padding:14px 16px;border-radius:var(--r2);border:2px solid var(--brd);background:var(--card);cursor:pointer;margin-bottom:10px;transition:all .15s}
.ct-card:active{transform:scale(.98)}
.ct-card.sel{border-color:var(--em);background:var(--em-l)}
.sic{display:flex;align-items:center;gap:10px;padding:11px 13px;border-radius:var(--r);border:1px solid var(--brd);background:var(--card);margin-bottom:8px}
.cart-box{background:var(--em-l);border:1px solid var(--em-t);border-radius:var(--r);padding:12px 14px;margin:12px 0}

/* IMPORT */
.drop-z{border:2px dashed var(--brd2);border-radius:var(--r2);padding:40px 20px;text-align:center;cursor:pointer;transition:all .18s;background:var(--bg2);position:relative;overflow:hidden}
.drop-z:hover{border-color:var(--in);background:var(--in-l)}
.drop-z input{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%}

/* PHOTO */
.photo-zone{border:2px dashed var(--brd2);border-radius:var(--r);padding:18px;text-align:center;cursor:pointer;transition:all .15s;background:var(--bg2);position:relative;overflow:hidden;margin-bottom:10px}
.photo-zone:hover{border-color:var(--in);background:var(--in-l)}
.photo-zone input{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%}

/* WA BANNER */
.wa-banner{width:100%;display:flex;align-items:center;justify-content:space-between;padding:16px 18px;border-radius:var(--r2);background:linear-gradient(135deg,var(--wa),var(--wa-d));border:none;cursor:pointer;margin-bottom:14px;box-shadow:0 4px 16px rgba(37,211,102,.3);transition:all .15s}
.wa-banner:active{transform:scale(.98)}

/* ADMIN DASHBOARD */
.metric-grid{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:14px}
.metric-card{background:var(--card);border-radius:var(--r2);border:1px solid var(--brd);padding:16px;box-shadow:var(--sh);text-align:center}
.metric-val{font-size:32px;font-weight:800;font-family:var(--mf);color:var(--in);line-height:1}
.metric-lbl{font-size:11px;color:var(--t3);margin-top:4px;font-weight:600}

/* LOADING */
.loading-screen{display:flex;flex-direction:column;align-items:center;justify-content:center;height:100vh;background:var(--bg);gap:14px}
.loading-spinner{width:40px;height:40px;border:3px solid var(--brd);border-top-color:var(--in);border-radius:50%;animation:spin 1s linear infinite}

/* PREV ROWS */
.prev-row{display:flex;align-items:center;gap:9px;padding:8px 12px;border-radius:8px;margin-bottom:5px;font-size:12px}
.prev-ok{background:var(--em-l);border:1px solid var(--em-t)}.prev-err{background:var(--cr-l);border:1px solid var(--cr-t)}

/* EMPTY */
.empty{text-align:center;padding:36px 16px;color:var(--t3);font-size:13px;line-height:1.7}

/* TOAST */
.toast-wrap{position:fixed;bottom:calc(var(--tab) + 10px);left:12px;right:12px;z-index:500;display:flex;flex-direction:column;gap:7px;pointer-events:none}
.toast-wrap > *{pointer-events:all;animation:toastIn .3s cubic-bezier(.16,1,.3,1)}

/* MISC */
.row{display:flex;align-items:center}.jb{justify-content:space-between}.g8{gap:8px}.g12{gap:12px}
.or-div{display:flex;align-items:center;gap:8px;color:var(--t4);font-size:11px;margin:10px 0}
.or-div::before,.or-div::after{content:"";flex:1;height:1px;background:var(--brd)}
::-webkit-scrollbar{width:4px;height:4px}::-webkit-scrollbar-thumb{background:var(--brd2);border-radius:3px}
`;

// ─── SUBCOMPONENTES ───────────────────────────────────────────────────────────
function Avatar(props) {
  var name=props.name, color=props.color, size=props.size||38, style=props.style||{};
  return (
    <div style={Object.assign({width:size,height:size,borderRadius:"50%",background:color||"#6366f1",display:"flex",alignItems:"center",justifyContent:"center",fontSize:Math.round(size*.36),fontWeight:800,color:"#fff",flexShrink:0,fontFamily:"var(--hf)"},style)}>
      {ini(name)}
    </div>
  );
}

function ProdThumb(props) {
  var prod=props.prod, size=props.size||40;
  if (!prod) return <div style={{width:size,height:size,borderRadius:9,background:"var(--bg2)",flexShrink:0}}/>;
  if (prod.photo_url) return <img src={prod.photo_url} alt="" style={{width:size,height:size,borderRadius:9,objectFit:"cover",flexShrink:0,border:"1px solid var(--brd)"}}/>;
  return <div style={{width:size,height:size,borderRadius:9,background:"var(--bg2)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:Math.round(size*.5),flexShrink:0}}>{prod.emoji||"📦"}</div>;
}

function SearchBar(props) {
  return (
    <div style={{position:"relative",marginBottom:12}}>
      <span style={{position:"absolute",left:12,top:"50%",transform:"translateY(-50%)",color:"var(--t3)",pointerEvents:"none"}}><Ic n="search" s={16}/></span>
      <input style={{width:"100%",padding:"11px 38px",borderRadius:10,border:"1.5px solid var(--brd)",background:"var(--card)",color:"var(--t1)",fontFamily:"var(--hf)",fontSize:14,outline:"none"}} placeholder={props.placeholder||"Buscar..."} value={props.value} onChange={function(e){props.onChange(e.target.value);}}/>
      {props.value?<button onClick={function(){props.onChange("");}} style={{position:"absolute",right:10,top:"50%",transform:"translateY(-50%)",background:"var(--bg2)",border:"none",borderRadius:6,padding:4,cursor:"pointer",color:"var(--t3)",display:"flex"}}><Ic n="x" s={14}/></button>:null}
    </div>
  );
}

function Toast(props) {
  var t=props.t;
  var st={s:{bg:"#ecfdf5",brd:"#10b981",ico:"#10b981"},e:{bg:"#fef2f2",brd:"#ef4444",ico:"#ef4444"},i:{bg:"#eff6ff",brd:"#6366f1",ico:"#6366f1"}}[t.type]||{bg:"#ecfdf5",brd:"#10b981",ico:"#10b981"};
  return (
    <div style={{background:st.bg,border:"1px solid "+st.brd,borderLeft:"4px solid "+st.brd,borderRadius:12,padding:"12px 14px",display:"flex",alignItems:"flex-start",gap:10,boxShadow:"0 4px 16px rgba(0,0,0,.1)"}}>
      <div style={{color:st.ico,marginTop:1}}><Ic n={t.type==="s"?"check":"x"} s={16}/></div>
      <div style={{flex:1}}><div style={{fontWeight:700,fontSize:13,color:"#1a1d2e"}}>{t.title}</div>{t.body?<div style={{fontSize:12,color:"#4b5068",marginTop:2}}>{t.body}</div>:null}</div>
      <button onClick={function(){props.remove(t.id);}} style={{background:"none",border:"none",cursor:"pointer",color:"#8b90a8",padding:2}}><Ic n="x" s={14}/></button>
    </div>
  );
}

function QtyControl(props) {
  var val=props.val, set=props.set, min=props.min!==undefined?props.min:1, max=props.max||9999, big=props.big;
  if (big) return (
    <div style={{display:"flex",border:"1.5px solid var(--brd)",borderRadius:12,overflow:"hidden",background:"var(--card)"}}>
      <button onClick={function(){set(Math.max(min,val-1));}} style={{flex:1,height:56,border:"none",background:"var(--bg2)",cursor:"pointer",color:"var(--t2)",display:"flex",alignItems:"center",justifyContent:"center"}}><Ic n="minus" s={20}/></button>
      <div style={{flex:2,textAlign:"center",fontFamily:"var(--mf)",fontSize:28,fontWeight:800,color:"var(--in)",borderLeft:"1.5px solid var(--brd)",borderRight:"1.5px solid var(--brd)",background:"var(--card)",lineHeight:"56px"}}>{val}</div>
      <button onClick={function(){set(Math.min(max,val+1));}} style={{flex:1,height:56,border:"none",background:"var(--bg2)",cursor:"pointer",color:"var(--t2)",display:"flex",alignItems:"center",justifyContent:"center"}}><Ic n="plus" s={20}/></button>
    </div>
  );
  return (
    <div style={{display:"flex",border:"1.5px solid var(--brd)",borderRadius:9,overflow:"hidden",background:"var(--card)",width:"fit-content"}}>
      <button onClick={function(){set(Math.max(min,val-1));}} style={{width:32,height:32,border:"none",background:"none",cursor:"pointer",color:"var(--t3)",display:"flex",alignItems:"center",justifyContent:"center"}}><Ic n="minus" s={14}/></button>
      <div style={{minWidth:36,textAlign:"center",fontFamily:"var(--mf)",fontSize:15,fontWeight:700,color:"var(--in)",borderLeft:"1.5px solid var(--brd)",borderRight:"1.5px solid var(--brd)",background:"var(--card)",lineHeight:"32px"}}>{val}</div>
      <button onClick={function(){set(Math.min(max,val+1));}} style={{width:32,height:32,border:"none",background:"none",cursor:"pointer",color:"var(--t3)",display:"flex",alignItems:"center",justifyContent:"center"}}><Ic n="plus" s={14}/></button>
    </div>
  );
}

// ─── APP ──────────────────────────────────────────────────────────────────────
export default function App() {
  // ── SESSION ─────────────────────────────────────────────────────────────────
  var [loading,   setLoading]   = useState(true);
  var [me,        setMe]        = useState(null);   // users row
  var [authMode,  setAuthMode]  = useState("login");
  var [aEmail,    setAEmail]    = useState("");
  var [aPass,     setAPass]     = useState("");
  var [aName,     setAName]     = useState("");
  var [aRole,     setARole]     = useState("Revendedora Exclusiva");
  var [aPass2,    setAPass2]    = useState("");
  var [showPass,  setShowPass]  = useState(false);
  var [authErr,   setAuthErr]   = useState("");
  var [authShake, setAuthShake] = useState(false);

  // ── DATA ────────────────────────────────────────────────────────────────────
  var [products,   setProducts]   = useState([]);
  var [inventory,  setInventory]  = useState([]); // my rows from inventory table
  var [contacts,   setContacts]   = useState([]); // user objects
  var [allUsers,   setAllUsers]   = useState([]); // for superadmin
  var [transfers,  setTransfers]  = useState([]);
  var [notifs,     setNotifs]     = useState([]);
  var [logs,       setLogs]       = useState([]);
  var [adminStats, setAdminStats] = useState(null);

  // ── UI ──────────────────────────────────────────────────────────────────────
  var [tab,       setTab]       = useState("stock");
  var [toasts,    setToasts]    = useState([]);
  var [srchStock, setSrchStock] = useState("");
  var [srchCat,   setSrchCat]   = useState("");
  var [srchCon,   setSrchCon]   = useState("");
  var [srchLog,   setSrchLog]   = useState("");

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

  // ── QUICK LOAD ──────────────────────────────────────────────────────────────
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

  // ── SEND ────────────────────────────────────────────────────────────────────
  var [sendStep, setSendStep] = useState(1);
  var [sendTo,   setSendTo]   = useState("");
  var [sendCart, setSendCart] = useState({});
  var [sendSrch, setSendSrch] = useState("");

  // ── IMPORT ──────────────────────────────────────────────────────────────────
  var [impTab,  setImpTab]  = useState("file");
  var [impTxt,  setImpTxt]  = useState("");
  var [impRows, setImpRows] = useState([]);
  var [impQty,  setImpQty]  = useState(0);
  var [impFile, setImpFile] = useState(null);

  // ── CONTACTS ────────────────────────────────────────────────────────────────
  var [ctQ,     setCtQ]     = useState("");

  // ── MODALS ──────────────────────────────────────────────────────────────────
  var [txModal,  setTxModal]  = useState(null);
  var [txQty,    setTxQty]    = useState(1);
  var [txTo,     setTxTo]     = useState("");
  var [shareM,   setShareM]   = useState(false);
  var [shareSel, setShareSel] = useState({});

  // ── HELPERS ──────────────────────────────────────────────────────────────────
  function toast(title, body, type) {
    var id = Date.now() + Math.random();
    setToasts(function(p){ return [...p,{id:id,title:title,body:body||"",type:type||"s"}]; });
    setTimeout(function(){ setToasts(function(p){ return p.filter(function(t){ return t.id!==id; }); }); }, 3800);
  }
  function rmToast(id){ setToasts(function(p){ return p.filter(function(t){ return t.id!==id; }); }); }
  function shake(){ setAuthShake(true); setTimeout(function(){ setAuthShake(false); },450); }

  var isAdmin = me && me.role === "superadmin";
  var myStock = inventory.filter(function(i){ return i.qty_available>0; });
  var unreadNotifs = notifs.filter(function(n){ return !n.read; }).length;
  var pendingTransfers = transfers.filter(function(t){ return t.to_user_id===me?.id && t.status==="pending"; }).length;
  var totalBadge = unreadNotifs + pendingTransfers;

  // ── SUPABASE DATA LOADERS ────────────────────────────────────────────────────
  var loadData = useCallback(async function(userId, userRole) {
    try {
      // Products (all users)
      var pr = await sb.from("products").select("*").eq("is_active",true).order("name");
      if (pr.data) setProducts(pr.data);

      // My inventory
      var inv = await sb.from("inventory").select("*, products(*)").eq("user_id",userId);
      if (inv.data) setInventory(inv.data);

      // My contacts
      var cts = await sb.from("contacts").select("*, contact:contact_id(id,name,email,color,role)").eq("user_id",userId);
      if (cts.data) setContacts(cts.data.map(function(c){ return c.contact; }).filter(Boolean));

      // Transfers (sent and received)
      var tx = await sb.from("transfers").select("*, product:product_id(name,emoji,photo_url,sku), from_user:from_user_id(name,email), to_user:to_user_id(name,email)").or("from_user_id.eq."+userId+",to_user_id.eq."+userId).order("created_at",{ascending:false});
      if (tx.data) setTransfers(tx.data);

      // Notifications
      var nf = await sb.from("notifications").select("*").eq("to_user_id",userId).order("created_at",{ascending:false}).limit(50);
      if (nf.data) setNotifs(nf.data);

      // My logs
      var lg = await sb.from("sale_logs").select("*, product:product_id(name,sku,emoji)").eq("user_id",userId).order("created_at",{ascending:false}).limit(100);
      if (lg.data) setLogs(lg.data);

      // Admin only
      if (userRole === "superadmin") {
        var us = await sb.from("users").select("*").order("created_at",{ascending:false});
        if (us.data) setAllUsers(us.data);
        var stats = await sb.from("admin_dashboard").select("*").single();
        if (stats.data) setAdminStats(stats.data);
      }
    } catch(e) { console.error("loadData error:", e); }
  }, []);

  // ── AUTH INIT ────────────────────────────────────────────────────────────────
  useEffect(function(){
    sb.auth.getSession().then(async function(res){
      var session = res.data.session;
      if (session) {
        var pr = await sb.from("users").select("*").eq("id",session.user.id).single();
        if (pr.data) {
          setMe(pr.data);
          await loadData(pr.data.id, pr.data.role);
        }
      }
      setLoading(false);
    });

    var sub = sb.auth.onAuthStateChange(async function(event, session){
      if (event === "SIGNED_OUT") { setMe(null); setProducts([]); setInventory([]); setContacts([]); setLogs([]); }
    });
    return function(){ sub.data.subscription.unsubscribe(); };
  }, [loadData]);

  // ── REALTIME ────────────────────────────────────────────────────────────────
  useEffect(function(){
    if (!me) return;
    var ch = sb.channel("stockpro_"+me.id)
      .on("postgres_changes",{event:"*",schema:"public",table:"notifications",filter:"to_user_id=eq."+me.id},function(){
        sb.from("notifications").select("*").eq("to_user_id",me.id).order("created_at",{ascending:false}).limit(50).then(function(r){ if(r.data) setNotifs(r.data); });
      })
      .on("postgres_changes",{event:"*",schema:"public",table:"transfers",filter:"to_user_id=eq."+me.id},function(){
        sb.from("transfers").select("*, product:product_id(name,emoji,photo_url,sku), from_user:from_user_id(name,email), to_user:to_user_id(name,email)").or("from_user_id.eq."+me.id+",to_user_id.eq."+me.id).order("created_at",{ascending:false}).then(function(r){ if(r.data) setTransfers(r.data); });
      })
      .subscribe();
    return function(){ sb.removeChannel(ch); };
  }, [me]);

  // ── LOGIN ────────────────────────────────────────────────────────────────────
  async function doLogin() {
    setAuthErr("");
    var res = await sb.auth.signInWithPassword({email:aEmail.trim().toLowerCase(), password:aPass});
    if (res.error) { setAuthErr(res.error.message==="Invalid login credentials"?"Email o contraseña incorrectos.":res.error.message); shake(); return; }
    // Retry loop: wait for RLS session to settle
    var pr = null;
    for (var attempt=0; attempt<3; attempt++) {
      await new Promise(function(r){ setTimeout(r, 600); });
      pr = await sb.from("users").select("*").eq("id",res.data.user.id).single();
      if (pr.data) break;
    }
    if (pr && pr.data) {
      setMe(pr.data);
      await loadData(pr.data.id, pr.data.role);
      toast("Bienvenido, "+pr.data.name+"!","","s");
    } else {
      setAuthErr("No se encontró tu perfil. Intentá de nuevo en unos segundos.");
      shake();
    }
  }

  // ── REGISTER ────────────────────────────────────────────────────────────────
  async function doRegister() {
    var em = aEmail.trim().toLowerCase();
    if (!em||!aName.trim()||!aPass){ setAuthErr("Completa todos los campos."); shake(); return; }
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(em)){ setAuthErr("Email inválido."); shake(); return; }
    if (aPass!==aPass2){ setAuthErr("Las contraseñas no coinciden."); shake(); return; }
    if (aPass.length<6){ setAuthErr("Contraseña mínimo 6 caracteres."); shake(); return; }
    var res = await sb.auth.signUp({email:em, password:aPass, options:{data:{name:aName.trim()}}});
    if (res.error){ setAuthErr(res.error.message); shake(); return; }
    // Update name in users table (trigger creates the row)
    if (res.data.user) {
      await new Promise(function(r){ setTimeout(r,1000); }); // wait for trigger
      await sb.from("users").update({name:aName.trim()}).eq("id",res.data.user.id);
      var pr = await sb.from("users").select("*").eq("id",res.data.user.id).single();
      if (pr.data){ setMe(pr.data); await loadData(pr.data.id, pr.data.role); toast("Cuenta creada!","Bienvenida "+aName,"s"); }
    }
  }

  // ── LOGOUT ──────────────────────────────────────────────────────────────────
  async function doLogout() {
    var n = me.name;
    await sb.auth.signOut();
    setMe(null); setTab("stock"); setAEmail(""); setAPass(""); toast("Hasta luego, "+n+"!","","i");
  }

  // ── SELL ────────────────────────────────────────────────────────────────────
  async function doSell(invRow) {
    if (invRow.qty_available<=0) return;
    var prod = invRow.products || products.find(function(p){ return p.id===invRow.product_id; });
    var r = await sb.from("inventory").update({qty_available:invRow.qty_available-1,qty_sold:invRow.qty_sold+1}).eq("id",invRow.id).select().single();
    if (r.error){ toast("Error",""+r.error.message,"e"); return; }
    setInventory(function(p){ return p.map(function(i){ return i.id===invRow.id?r.data:i; }); });
    await sb.from("sale_logs").insert({user_id:me.id,product_id:invRow.product_id,qty:1,sale_price:prod?prod.price:0,source:"own_stock"});
    toast("Venta registrada",prod?prod.name:"","s");
  }

  // ── TRANSFER (send to contact, pending confirmation) ─────────────────────────
  async function doTx() {
    if (txQty>txModal.qty_available){ toast("Stock insuficiente","","e"); return; }
    var prod = txModal.products || products.find(function(p){ return p.id===txModal.product_id; });
    var tUser = contacts.find(function(c){ return c.id===txTo; });
    // Deduct from sender
    var upd = await sb.from("inventory").update({qty_available:txModal.qty_available-txQty}).eq("id",txModal.id).select().single();
    if (upd.error){ toast("Error",""+upd.error.message,"e"); return; }
    setInventory(function(p){ return p.map(function(i){ return i.id===txModal.id?upd.data:i; }); });
    // Create pending transfer
    await sb.from("transfers").insert({from_user_id:me.id,to_user_id:txTo,product_id:txModal.product_id,qty:txQty,status:"pending"});
    // Notify recipient
    await sb.from("notifications").insert({to_user_id:txTo,from_name:me.name,type:"transfer",message:me.name+" te envió "+txQty+"x "+(prod?prod.name:"producto")+". Confirmá la recepción!"});
    toast("Enviado!","Esperando confirmación de "+(tUser?tUser.name:""),"s");
    setTxModal(null);
    await loadData(me.id,me.role);
  }

  // ── CONFIRM TRANSFER ─────────────────────────────────────────────────────────
  async function confirmTransfer(tx) {
    var prod = tx.product || products.find(function(p){ return p.id===tx.product_id; });
    // Add to receiver inventory
    var existing = inventory.find(function(i){ return i.product_id===tx.product_id; });
    if (existing) {
      await sb.from("inventory").update({qty_available:existing.qty_available+tx.qty}).eq("id",existing.id);
    } else {
      await sb.from("inventory").insert({user_id:me.id,product_id:tx.product_id,qty_available:tx.qty,qty_sold:0});
    }
    // Mark transfer confirmed
    await sb.from("transfers").update({status:"confirmed",confirmed_at:new Date().toISOString()}).eq("id",tx.id);
    // Notify sender
    await sb.from("notifications").insert({to_user_id:tx.from_user_id,from_name:me.name,type:"confirm",message:me.name+" confirmó la recepción de "+tx.qty+"x "+(prod?prod.name:"producto")+"!"});
    // Mark my notification read
    await sb.from("notifications").update({read:true}).eq("to_user_id",me.id).eq("type","transfer");
    toast("Recepción confirmada!",(prod?prod.name:"")+" en tu stock","s");
    await loadData(me.id,me.role);
  }

  // ── MULTI SEND ──────────────────────────────────────────────────────────────
  async function doMultiSend() {
    var entries = Object.entries(sendCart).filter(function(e){ return Number(e[1])>0; });
    if (!entries.length){ toast("Agrega al menos un producto","","e"); return; }
    var tUser = contacts.find(function(c){ return c.id===sendTo; });
    for (var i=0; i<entries.length; i++) {
      var invId = entries[i][0], qty = Number(entries[i][1]);
      var invRow = inventory.find(function(iv){ return iv.id===invId; });
      if (!invRow) continue;
      var prod = invRow.products || products.find(function(p){ return p.id===invRow.product_id; });
      await sb.from("inventory").update({qty_available:invRow.qty_available-qty}).eq("id",invId);
      await sb.from("transfers").insert({from_user_id:me.id,to_user_id:sendTo,product_id:invRow.product_id,qty:qty,status:"pending"});
    }
    await sb.from("notifications").insert({to_user_id:sendTo,from_name:me.name,type:"transfer",message:me.name+" te envió "+entries.length+" producto(s). Confirmá la recepción!"});
    toast("Envío completado!",entries.length+" producto(s) a "+(tUser?tUser.name:""),"s");
    setSendCart({}); setSendStep(1);
    await loadData(me.id,me.role);
  }

  // ── CATALOG ABM (superadmin only) ────────────────────────────────────────────
  async function doSaveProd(e) {
    e.preventDefault();
    if (!fSku.trim()||!fName.trim()||!fPrice){ toast("Completa SKU, nombre y precio","","e"); return; }
    var skuC=fSku.trim().toUpperCase(), pVal=parseFloat(fPrice)||0;
    if (editP) {
      var upd = await sb.from("products").update({sku:skuC,name:fName.trim(),price:pVal,category:fCat,emoji:fEmoji,photo_url:fPhoto||editP.photo_url||null,updated_at:new Date().toISOString()}).eq("id",editP.id).select().single();
      if (upd.error){ toast("Error",""+upd.error.message,"e"); return; }
      setProducts(function(p){ return p.map(function(x){ return x.id===editP.id?upd.data:x; }); });
      toast("Producto actualizado","","s"); setEditP(null);
    } else {
      if (products.some(function(p){ return p.sku===skuC; })){ toast("SKU ya existe","","e"); return; }
      var ins = await sb.from("products").insert({sku:skuC,name:fName.trim(),price:pVal,category:fCat,emoji:fEmoji,photo_url:fPhoto||null,created_by:me.id}).select().single();
      if (ins.error){ toast("Error",""+ins.error.message,"e"); return; }
      var newProd = ins.data;
      setProducts(function(p){ return [...p,newProd]; });
      var qty=parseInt(fStock)||0;
      if (qty>0){ await sb.from("inventory").insert({user_id:me.id,product_id:newProd.id,qty_available:qty,qty_sold:0}); }
      toast("Producto creado!",fName.trim(),"s");
    }
    setFSku(""); setFName(""); setFPrice(""); setFEmoji("✨"); setFStock("0"); setFPhoto(null);
    await loadData(me.id,me.role);
  }

  function startEdit(p){ setEditP(p); setFSku(p.sku); setFName(p.name); setFPrice(String(p.price)); setFCat(p.category||"General"); setFEmoji(p.emoji||"✨"); setFPhoto(null); setTab("catalog"); }
  function cancelEdit(){ setEditP(null); setFSku(""); setFName(""); setFPrice(""); setFEmoji("✨"); setFStock("0"); setFPhoto(null); }

  async function doDelProd(prodId, sku) {
    await sb.from("products").update({is_active:false}).eq("id",prodId);
    setProducts(function(p){ return p.filter(function(x){ return x.id!==prodId; }); });
    setDelConf(null); if (editP&&editP.id===prodId) cancelEdit();
    toast("Producto eliminado","","i");
  }

  // ── QUICK LOAD ────────────────────────────────────────────────────────────────
  async function doQuickLoad() {
    if (qlMode==="add") {
      if (!qlPid){ toast("Selecciona un producto","","e"); return; }
      var existing = inventory.find(function(i){ return i.product_id===qlPid; });
      if (existing) {
        var upd = await sb.from("inventory").update({qty_available:existing.qty_available+qlQty}).eq("id",existing.id).select("*, products(*)").single();
        if (upd.data) setInventory(function(p){ return p.map(function(i){ return i.id===existing.id?upd.data:i; }); });
      } else {
        var ins = await sb.from("inventory").insert({user_id:me.id,product_id:qlPid,qty_available:qlQty,qty_sold:0}).select("*, products(*)").single();
        if (ins.data) setInventory(function(p){ return [...p,ins.data]; });
      }
      var p=products.find(function(x){ return x.id===qlPid; });
      toast("Stock cargado!","+"+ qlQty+" u. de "+(p?p.name:""),"s");
      setQlQty(1); setQlPid("");
    } else {
      if (!qlSku.trim()||!qlName.trim()||!qlPrice){ toast("Completa todos los campos","","e"); return; }
      var skuC=qlSku.trim().toUpperCase();
      if (products.some(function(p){ return p.sku===skuC; })){ toast("SKU ya existe","","e"); return; }
      var pr = await sb.from("products").insert({sku:skuC,name:qlName.trim(),price:parseFloat(qlPrice)||0,emoji:qlEmoji,photo_url:qlPhoto||null,category:qlCat,created_by:me.id}).select().single();
      if (pr.error){ toast("Error",""+pr.error.message,"e"); return; }
      setProducts(function(p){ return [...p,pr.data]; });
      if (qlQty>0){
        var ii = await sb.from("inventory").insert({user_id:me.id,product_id:pr.data.id,qty_available:qlQty,qty_sold:0}).select("*, products(*)").single();
        if (ii.data) setInventory(function(p){ return [...p,ii.data]; });
      }
      toast("Producto creado!",qlName.trim()+" con "+qlQty+" u.","s");
      setQlSku(""); setQlName(""); setQlPrice(""); setQlQty(1); setQlPhoto(null); setQlMode("add");
    }
  }

  // ── PHOTO UPLOAD ─────────────────────────────────────────────────────────────
  async function handlePhoto(e, setter) {
    var file=e.target.files&&e.target.files[0];
    if (!file) return;
    if (file.size>1600000){ toast("Foto muy grande","Usa una imagen menor a 1.6MB","e"); return; }
    var reader=new FileReader();
    reader.onloadend=function(){ setter(reader.result); toast("Foto cargada","","s"); };
    reader.readAsDataURL(file);
  }

  // ── IMPORT ────────────────────────────────────────────────────────────────────
  function parseLines(text, qty) {
    var lines=text.split("\n"), rows=[], curCat="Importado";
    lines.forEach(function(line){
      var row=line.trim();
      if (!row||row.toLowerCase().includes("mostrar")) return;
      if (!row.includes("-")&&!row.includes(",")&&!row.includes("\t")){ if(row.length>2&&row.length<40) curCat=row; return; }
      var cols=row.includes("\t")?row.split("\t"):row.split(",");
      if (cols.length<2){ rows.push({ok:false,raw:row,err:"Menos de 2 columnas"}); return; }
      var c0=cols[0].trim(), c1=cols[1].trim();
      if (!c0.includes("-")){ rows.push({ok:false,raw:row,err:"Sin guion separador"}); return; }
      var dash=c0.indexOf("-"), sku=c0.substring(0,dash).trim().toUpperCase(), name=c0.substring(dash+1).trim();
      if (!sku||!name){ rows.push({ok:false,raw:row,err:"SKU o nombre vacío"}); return; }
      var pc=c1.replace(/[^\d.,]/g,"");
      if (pc.includes(",")){ pc=pc.replace(/\./g,"").replace(",","."); }
      rows.push({ok:true,sku:sku,name:name,cat:curCat,price:parseFloat(pc)||0,qty:qty});
    });
    return rows;
  }

  function doParseFile(e) {
    var file=e.target.files&&e.target.files[0]; if (!file) return;
    setImpFile(file.name);
    var ext=file.name.split(".").pop().toLowerCase();
    var reader=new FileReader();
    reader.onload=function(ev){
      try {
        var rows=[];
        if (ext==="xlsx"||ext==="xls"){
          var wb=XLSX.read(ev.target.result,{type:"binary"});
          var ws=wb.Sheets[wb.SheetNames[0]];
          var data=XLSX.utils.sheet_to_json(ws,{header:1,raw:false});
          var sr=0,cs=-1,cn=-1,cc=-1,cp=-1;
          if (data.length>0){ var h=data[0].map(function(x){ return (x||"").toString().toLowerCase(); }); h.forEach(function(x,i){ if(/sku|cod/.test(x)) cs=i; else if(/nom|desc|prod|det/.test(x)) cn=i; else if(/cat|rub/.test(x)) cc=i; else if(/prec|price|val/.test(x)) cp=i; }); if(cs>=0&&cn>=0) sr=1; else{ cs=0;cn=1;cc=2;cp=3;sr=1; } }
          data.slice(sr).forEach(function(row){ if(!row||!row.length) return; var sku=(row[cs]||"").toString().trim().toUpperCase(); var name=(row[cn]||"").toString().trim(); var cat=cc>=0?(row[cc]||"Importado").toString().trim():"Importado"; var pc=cp>=0?(row[cp]||"0").toString():"0"; pc=pc.replace(/[^\d.,]/g,""); if(pc.includes(",")){ pc=pc.replace(/\./g,"").replace(",","."); } if(!sku||!name){ rows.push({ok:false,raw:(row||[]).join(","),err:"Fila incompleta"}); return; } rows.push({ok:true,sku:sku,name:name,cat:cat,price:parseFloat(pc)||0,qty:impQty}); });
        } else { rows=parseLines(ev.target.result,impQty); }
        setImpRows(rows);
        toast(rows.filter(function(r){return r.ok;}).length+" detectados","Revisa y confirma","s");
      } catch(err){ toast("Error al leer archivo","","e"); }
    };
    if (ext==="xlsx"||ext==="xls") reader.readAsBinaryString(file); else reader.readAsText(file);
  }

  function doParseTxt(){ if (!impTxt.trim()){ toast("Pega texto primero","","e"); return; } var rows=parseLines(impTxt,impQty); setImpRows(rows); toast(rows.filter(function(r){return r.ok;}).length+" detectados","Revisa y confirma","s"); }

  async function doConfirmImport() {
    var valid=impRows.filter(function(r){return r.ok;});
    if (!valid.length){ toast("Sin filas válidas","","e"); return; }
    var count=0;
    var newProds = [];
    for (var i=0;i<valid.length;i++){
      var row=valid[i];
      var existing=products.find(function(p){ return p.sku===row.sku; });
      var pid;
      if (!existing){
        var ins=await sb.from("products").insert({sku:row.sku,name:row.name,price:row.price,emoji:"✨",category:row.cat,created_by:me.id}).select().single();
        if (ins.data){ newProds.push(ins.data); pid=ins.data.id; }
      } else {
        // Update price
        await sb.from("products").update({price:row.price,name:row.name,category:row.cat}).eq("id",existing.id);
        pid=existing.id;
      }
      if (pid){
        // Only touch inventory if qty > 0
        if (row.qty > 0) {
          var inv=inventory.find(function(ii){ return ii.product_id===pid; });
          if (inv){ await sb.from("inventory").update({qty_available:inv.qty_available+row.qty}).eq("id",inv.id); }
          else { await sb.from("inventory").insert({user_id:me.id,product_id:pid,qty_available:row.qty,qty_sold:0}); }
        }
        count++;
      }
    }
    if (newProds.length>0) setProducts(function(p){ return [...p,...newProds]; });
    toast("Importación exitosa!",count+" productos al catálogo","s");
    setImpRows([]); setImpTxt(""); setImpFile(null); setTab("catalog");
    await loadData(me.id,me.role);
  }

  // ── CONTACTS ──────────────────────────────────────────────────────────────────
  async function doAddContact() {
    var em=ctQ.trim().toLowerCase();
    var target=await sb.from("users").select("*").eq("email",em).single();
    if (target.error||!target.data){ toast("Email no encontrado","","e"); return; }
    if (target.data.id===me.id){ toast("No puedes agregarte","","e"); return; }
    if (contacts.some(function(c){ return c.id===target.data.id; })){ toast("Ya está en tu red","","i"); return; }
    var ins=await sb.from("contacts").insert({user_id:me.id,contact_id:target.data.id});
    if (ins.error){ toast("Error",""+ins.error.message,"e"); return; }
    setContacts(function(p){ return [...p,target.data]; });
    toast("Contacto agregado",target.data.name,"s"); setCtQ("");
  }

  // ── WHATSAPP SHARE ────────────────────────────────────────────────────────────
  function shareOne(prod) {
    var inv=inventory.find(function(i){ return i.product_id===prod.id; });
    var avail=inv?inv.qty_available:0;
    var nl="\n";
    var msg=prod.name+nl+"Precio: "+fmtARS(prod.price)+nl+(avail>0?"Stock: "+avail+" u."+nl:"")+"Escribime para pedirlo!";
    if (prod.photo_url&&navigator.share){ try{ navigator.share({title:prod.name,text:msg,url:prod.photo_url}).catch(function(){}); return; }catch(e){} }
    if (navigator.share){ navigator.share({title:prod.name,text:msg}).catch(function(){}); return; }
    window.open("https://wa.me/?text="+encodeURIComponent(msg),"_blank");
  }

  function doShareCatalog() {
    var sel=myStock.filter(function(i){ return shareSel[i.id]; });
    if (!sel.length){ toast("Selecciona al menos un producto","","e"); return; }
    var nl="\n";
    var lines=sel.map(function(item){ var p=item.products||products.find(function(x){ return x.id===item.product_id; }); if(!p) return ""; return (p.emoji||"")+" *"+p.name+"*"+nl+"   "+fmtARS(p.price)+" · "+item.qty_available+" u."; }).filter(Boolean);
    var msg="Mis productos disponibles hoy:"+nl+nl+lines.join(nl+nl)+nl+nl+"Escribime para hacer tu pedido!";
    setShareM(false); setShareSel({});
    if (navigator.share){ navigator.share({title:"Mis productos",text:msg}).catch(function(){ window.open("https://wa.me/?text="+encodeURIComponent(msg),"_blank"); }); return; }
    window.open("https://wa.me/?text="+encodeURIComponent(msg),"_blank");
  }

  // ── DERIVED ───────────────────────────────────────────────────────────────────
  var stockFiltered = myStock.filter(function(i){
    var p=i.products||products.find(function(x){ return x.id===i.product_id; });
    if (!p) return false;
    var q=srchStock.toLowerCase();
    return !q||p.name.toLowerCase().includes(q)||p.sku.toLowerCase().includes(q)||(p.category||"").toLowerCase().includes(q);
  });

  var catFiltered = products.filter(function(p){ var q=srchCat.toLowerCase(); return !q||p.name.toLowerCase().includes(q)||p.sku.toLowerCase().includes(q); });
  var conFiltered = contacts.filter(function(c){ var q=srchCon.toLowerCase(); return !q||c.name.toLowerCase().includes(q)||c.email.toLowerCase().includes(q); });
  var logFiltered = logs.filter(function(l){ var q=srchLog.toLowerCase(); var p=l.product; return !q||(p&&p.name.toLowerCase().includes(q))||l.source.includes(q); });
  var sendFiltered = myStock.filter(function(i){ var p=i.products||products.find(function(x){ return x.id===i.product_id; }); if(!p) return false; var q=sendSrch.toLowerCase(); return !q||p.name.toLowerCase().includes(q)||p.sku.toLowerCase().includes(q); });
  var qlFiltered = products.filter(function(p){ var q=qlSrch.toLowerCase(); return !q||p.name.toLowerCase().includes(q)||p.sku.toLowerCase().includes(q); });
  var pendingTx = transfers.filter(function(t){ return t.to_user_id===me?.id&&t.status==="pending"; });
  var sentTx = transfers.filter(function(t){ return t.from_user_id===me?.id&&t.status==="pending"; });
  var myNotifs = notifs.filter(function(n){ return n.to_user_id===me?.id; });

  // ── LOADING SCREEN ────────────────────────────────────────────────────────────
  if (loading) return (
    <div>
      <style>{CSS}</style>
      <div className="loading-screen">
        <img src="https://cgqrgyouunxfluujhwey.supabase.co/storage/v1/object/public/assets/logo.png" alt="StockPro" style={{width:200,marginBottom:8}} onError={function(e){e.target.style.display="none";}}/>
        <div style={{fontSize:22,fontWeight:800,color:"var(--navy)"}}>STOCKPRO</div>
        <div className="loading-spinner"/>
        <div style={{fontSize:12,color:"var(--t3)"}}>Conectando...</div>
      </div>
    </div>
  );

  // ── LOGIN SCREEN ──────────────────────────────────────────────────────────────
  if (!me) return (
    <div>
      <style>{CSS}</style>
      <div className="login-bg">
        <div className={"lbox"+(authShake?" shake":"")}>
          <div style={{textAlign:"center",marginBottom:20}}>
            <img src="https://cgqrgyouunxfluujhwey.supabase.co/storage/v1/object/public/assets/logo.png" alt="StockPro" style={{width:"100%",maxWidth:260,margin:"0 auto",display:"block"}}
              onError={function(e){ e.target.style.display="none"; e.target.nextSibling.style.display="block"; }}
            />
            <div style={{display:"none",fontSize:22,fontWeight:800,color:"var(--navy)"}}>STOCKPRO</div>
          </div>
          <div className="ltabs">
            <div className={"ltab"+(authMode==="login"?" on":"")} onClick={function(){setAuthMode("login");setAuthErr("");}}>Iniciar sesión</div>
            <div className={"ltab"+(authMode==="reg"?" on":"")} onClick={function(){setAuthMode("reg");setAuthErr("");}}>Registrarme</div>
          </div>
          {authErr?<div className="err-box"><Ic n="x" s={14}/>{authErr}</div>:null}
          {authMode==="login"?(
            <div>
              <div className="fld"><label className="fl">Email</label><input className="fi" type="email" placeholder="tu@email.com" value={aEmail} onChange={function(e){setAEmail(e.target.value);setAuthErr("");}} onKeyDown={function(e){if(e.key==="Enter")doLogin();}}/></div>
              <div className="fld"><label className="fl">Contraseña</label><div className="pw-wrap"><input className="fi" style={{paddingRight:42}} type={showPass?"text":"password"} placeholder="••••••" value={aPass} onChange={function(e){setAPass(e.target.value);setAuthErr("");}} onKeyDown={function(e){if(e.key==="Enter")doLogin();}}/><button className="pw-eye" type="button" onClick={function(){setShowPass(function(v){return !v;});}}><Ic n={showPass?"eyeoff":"eye"} s={16}/></button></div></div>
              <button className="cta cta-in" onClick={doLogin}><Ic n="check" s={18}/>Entrar</button>
            </div>
          ):(
            <div>
              <div className="fld"><label className="fl">Nombre completo o negocio</label><input className="fi" placeholder="Ej: Laura Cosméticos" value={aName} onChange={function(e){setAName(e.target.value);setAuthErr("");}}/></div>
              <div className="fld"><label className="fl">Email</label><input className="fi" type="email" placeholder="tu@email.com" value={aEmail} onChange={function(e){setAEmail(e.target.value);setAuthErr("");}}/></div>
              <div className="fld"><label className="fl">Contraseña (mín. 6 caracteres)</label><div className="pw-wrap"><input className="fi" style={{paddingRight:42}} type={showPass?"text":"password"} placeholder="Mínimo 6 caracteres" value={aPass} onChange={function(e){setAPass(e.target.value);setAuthErr("");}}/><button className="pw-eye" type="button" onClick={function(){setShowPass(function(v){return !v;});}}><Ic n={showPass?"eyeoff":"eye"} s={16}/></button></div></div>
              <div className="fld"><label className="fl">Repetir contraseña</label><input className="fi" type="password" placeholder="Repetir contraseña" value={aPass2} onChange={function(e){setAPass2(e.target.value);setAuthErr("");}}/></div>
              <button className="cta cta-em" onClick={doRegister}><Ic n="plus" s={18}/>Crear cuenta gratis</button>
            </div>
          )}
        </div>
      </div>
      <div className="toast-wrap">{toasts.map(function(t){ return <Toast key={t.id} t={t} remove={rmToast}/>; })}</div>
    </div>
  );

  // ── TABS ─────────────────────────────────────────────────────────────────────
  var TABS = [
    {id:"stock",    lbl:"Stock",    ico:"box"},
    {id:"cargar",   lbl:"Cargar",   ico:"plus"},
    {id:"enviar",   lbl:"Enviar",   ico:"send"},
    {id:"importar", lbl:"Importar", ico:"upload"},
    {id:"catalog",  lbl:"Catálogo", ico:"list"},
    {id:"contacts", lbl:"Red",      ico:"users"},
  ];
  if (isAdmin) TABS.push({id:"admin",lbl:"Admin",ico:"shield"});

  // ── MAIN APP ──────────────────────────────────────────────────────────────────
  return (
    <div>
      <style>{CSS}</style>
      <div className="app">

        {/* HEADER */}
        <div className="hdr">
          <Avatar name={me.name} color={me.color||"#6366f1"} size={38}/>
          <div style={{flex:1,minWidth:0}}>
            <div className="hdr-name" style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{me.name}</div>
            <div className="hdr-role">{me.role==="superadmin"?"👑 Administrador":me.role}</div>
          </div>
          <div className="saved-dot"><div style={{width:6,height:6,borderRadius:"50%",background:"var(--em)",flexShrink:0}}/> online</div>
          {totalBadge>0?<div className="notif-badge" onClick={function(){setTab("contacts");}}><Ic n="bell" s={14}/>{totalBadge}</div>:null}
          <button className="ic-btn" style={{marginLeft:6}} onClick={doLogout}><Ic n="logout" s={17}/></button>
        </div>

        <div className="main">

          {/* ══ STOCK ══ */}
          {tab==="stock"&&(
            <div>
              <div className="ph"><div><div className="ph-h">Mi Stock</div><div className="ph-s">Solo artículos con unidades disponibles</div></div></div>
              <div className="pc">
                <button className="wa-banner" onClick={function(){setShareM(true);setShareSel({});}}>
                  <div className="row g12"><div style={{width:46,height:46,borderRadius:"50%",background:"rgba(255,255,255,.2)",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}><Ic n="wa" s={24}/></div><div style={{textAlign:"left"}}><div style={{fontSize:15,fontWeight:800,color:"#fff"}}>Compartir disponible por WhatsApp</div><div style={{fontSize:11,color:"rgba(255,255,255,.8)",marginTop:2}}>Selecciona productos y envía con fotos</div></div></div>
                  <div style={{color:"rgba(255,255,255,.7)"}}><Ic n="send" s={18}/></div>
                </button>
                <SearchBar value={srchStock} onChange={setSrchStock} placeholder="Buscar por nombre, SKU o categoría..."/>

                {/* Pending transfers to confirm */}
                {pendingTx.length>0&&(
                  <div className="card" style={{border:"2px solid var(--am)",background:"var(--am-l)"}}>
                    <div className="card-h" style={{background:"var(--am-l)"}}><div className="card-title" style={{color:"var(--am-d)"}}>🔔 Confirmar recepción ({pendingTx.length})</div></div>
                    <div style={{padding:"8px 12px"}}>
                      {pendingTx.map(function(tx){
                        var p=tx.product;
                        return (<div key={tx.id} style={{display:"flex",alignItems:"center",gap:10,padding:"10px 12px",background:"var(--card)",borderRadius:10,marginBottom:8,border:"1px solid var(--brd)"}}>
                          <div style={{fontSize:26}}>{p?p.emoji:"📦"}</div>
                          <div style={{flex:1}}><div style={{fontWeight:700,fontSize:13}}>{p?p.name:"Producto"}</div><div style={{fontSize:11,color:"var(--t3)"}}>{tx.qty} u. de {tx.from_user?tx.from_user.name:""}</div></div>
                          <button className="btn b-em" style={{fontSize:12,padding:"8px 12px"}} onClick={function(){confirmTransfer(tx);}}>✓ Confirmar</button>
                        </div>);
                      })}
                    </div>
                  </div>
                )}

                {/* Sent pending */}
                {sentTx.length>0&&(
                  <div className="card">
                    <div className="card-h"><div className="card-title">📤 Enviados (esperando confirmación)</div><span className="badge b-am">{sentTx.length}</span></div>
                    <div className="tw"><table>
                      <thead><tr><th>Producto</th><th>Cant.</th><th>Para</th><th>Estado</th></tr></thead>
                      <tbody>{sentTx.map(function(tx){ var p=tx.product; var tu=tx.to_user; return (<tr key={tx.id} className="tr"><td><div style={{fontWeight:600,fontSize:12}}>{p?p.name:""}</div></td><td><span style={{fontFamily:"var(--mf)",fontWeight:700,color:"var(--am-d)"}}>{tx.qty}</span></td><td><span style={{fontSize:11}}>{tu?tu.name:""}</span></td><td><span style={{background:"var(--am-t)",color:"var(--am-d)",borderRadius:5,padding:"2px 8px",fontSize:10,fontWeight:700}}>Pendiente</span></td></tr>); })}</tbody>
                    </table></div>
                  </div>
                )}

                <div className="card">
                  <div className="card-h"><div className="card-title"><div className="card-ico" style={{background:"var(--am-l)",color:"var(--am-d)"}}><Ic n="box" s={14}/></div>Stock Propio</div><span className="badge b-am">{stockFiltered.length}</span></div>
                  {stockFiltered.length===0?<div className="empty">Sin existencias.{srchStock?" Intenta otra búsqueda.":" Carga productos en Cargar."}</div>:(
                    <div className="tw"><table>
                      <thead><tr><th></th><th>SKU</th><th>Producto</th><th>Precio</th><th>Disp.</th><th></th></tr></thead>
                      <tbody>{stockFiltered.map(function(item){ var p=item.products||products.find(function(x){ return x.id===item.product_id; }); if(!p) return null; return (<tr key={item.id} className="tr"><td><ProdThumb prod={p} size={36}/></td><td><span style={{color:"var(--in-d)",fontFamily:"var(--mf)",fontSize:11,background:"var(--in-l)",padding:"2px 6px",borderRadius:5,fontWeight:600}}>{p.sku}</span></td><td><div style={{fontWeight:600,fontSize:12}}>{p.name}</div><div style={{fontSize:10,color:"var(--t3)"}}>{p.category}</div></td><td><span style={{fontFamily:"var(--mf)",fontWeight:700,fontSize:12}}>{fmtARS(p.price)}</span></td><td><span style={{fontFamily:"var(--mf)",fontWeight:800,color:"var(--em-d)",fontSize:14}}>{item.qty_available}</span></td><td><div className="row g8" style={{justifyContent:"flex-end",flexWrap:"wrap"}}><button className="btn btn-xs b-wa" onClick={function(){shareOne(p);}}><Ic n="wa" s={12}/></button><button className="btn btn-xs b-am" onClick={function(){setTxModal(item);setTxQty(1);setTxTo(contacts[0]?contacts[0].id:"");}} disabled={item.qty_available===0}><Ic n="send" s={11}/>Pasar</button><button className="btn btn-xs b-em" onClick={function(){doSell(item);}} disabled={item.qty_available===0}><Ic n="check" s={11}/>Venta</button></div></td></tr>); })}</tbody>
                    </table></div>
                  )}
                </div>
              </div>
            </div>
          )}

          {/* ══ CARGAR ══ */}
          {tab==="cargar"&&(
            <div>
              <div className="ph"><div><div className="ph-h">Cargar Stock</div><div className="ph-s">Agrega unidades o crea productos</div></div></div>
              <div className="pc">
                <div className="card">
                  <div className="card-h"><div className="card-title">Modo de carga</div></div>
                  <div style={{padding:"14px 14px 0"}}>
                    <div className="ql-tabs">
                      <div className={"ql-tab"+(qlMode==="add"?" on":"")} onClick={function(){setQlMode("add");}}>Agregar a existente</div>
                      <div className={"ql-tab"+(qlMode==="new"?" on":"")} onClick={function(){setQlMode("new");}}>Crear producto nuevo</div>
                    </div>
                    {qlMode==="add"?(
                      <div>
                        <div style={{fontSize:11,fontWeight:700,color:"var(--t3)",textTransform:"uppercase",letterSpacing:".08em",marginBottom:10}}>Seleccionar producto</div>
                        <SearchBar value={qlSrch} onChange={setQlSrch} placeholder="Buscar producto..."/>
                        <div style={{maxHeight:300,overflowY:"auto",marginBottom:14}}>
                          {qlFiltered.length===0?<div className="empty" style={{padding:"20px 0"}}>Sin resultados</div>:qlFiltered.map(function(p){ var si=inventory.find(function(s){ return s.product_id===p.id; }); return (<div key={p.id} className={"prod-card"+(qlPid===p.id?" sel":"")} onClick={function(){setQlPid(p.id);}}><ProdThumb prod={p} size={40}/><div style={{flex:1,minWidth:0}}><div style={{fontWeight:700,fontSize:13,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.name}</div><div style={{fontSize:11,color:"var(--t3)",marginTop:2}}>Stock: {si?si.qty_available:0} u. · {fmtARS(p.price)}</div></div>{qlPid===p.id&&<div style={{width:22,height:22,borderRadius:"50%",background:"var(--am)",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}><Ic n="check" s={12}/></div>}</div>); })}
                        </div>
                        <div style={{fontSize:11,fontWeight:700,color:"var(--t3)",textTransform:"uppercase",letterSpacing:".08em",marginBottom:8}}>Cantidad a agregar</div>
                        <QtyControl val={qlQty} set={setQlQty} big={true}/>
                        <button className="cta cta-am" onClick={doQuickLoad} disabled={!qlPid}><Ic n="plus" s={18}/>Agregar {qlQty} unidades</button>
                      </div>
                    ):(
                      <div style={{paddingBottom:16}}>
                        <div className="fld"><label className="fl">SKU</label><input className="fi" placeholder="Ej: PERF01" value={qlSku} onChange={function(e){setQlSku(e.target.value);}}/></div>
                        <div className="fld"><label className="fl">Nombre</label><input className="fi" placeholder="Ej: Perfume Lattafa 100ml" value={qlName} onChange={function(e){setQlName(e.target.value);}}/></div>
                        <div className="row g8" style={{marginBottom:12}}><div style={{flex:1}}><label className="fl">Precio ($)</label><input className="fi" type="number" placeholder="0.00" value={qlPrice} onChange={function(e){setQlPrice(e.target.value);}}/></div><div style={{flex:1}}><label className="fl">Icono</label><select className="fi fi-sel" value={qlEmoji} onChange={function(e){setQlEmoji(e.target.value);}}>{EMOJIS.map(function(o){ return <option key={o.v} value={o.v}>{o.v} {o.l}</option>; })}</select></div></div>
                        <div className="fld"><label className="fl">Categoría</label><select className="fi fi-sel" value={qlCat} onChange={function(e){setQlCat(e.target.value);}}>{CATS.map(function(c){ return <option key={c}>{c}</option>; })}</select></div>
                        <div className="fld"><label className="fl">Foto (opcional)</label><div className="photo-zone"><input type="file" accept="image/*" onChange={function(e){handlePhoto(e,setQlPhoto);}}/>{qlPhoto?<img src={qlPhoto} alt="" style={{width:60,height:60,borderRadius:10,objectFit:"cover",margin:"0 auto"}}/>:<div><div style={{fontSize:28,marginBottom:6}}>📸</div><div style={{fontSize:12,fontWeight:700,color:"var(--in)"}}>Tocar para subir foto</div></div>}</div>{qlPhoto&&<button onClick={function(){setQlPhoto(null);}} style={{fontSize:11,color:"var(--cr)",background:"none",border:"none",cursor:"pointer",fontWeight:700}}>Quitar foto</button>}</div>
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

          {/* ══ ENVIAR ══ */}
          {tab==="enviar"&&(
            <div>
              <div className="ph"><div><div className="ph-h">Enviar</div><div className="ph-s">Transferencia a contactos</div></div></div>
              <div className="pc">
                <div className="step-row">
                  <div style={{textAlign:"center",flex:1}}><div className={"step-dot"+(sendStep>=1?" on":"")} style={sendStep>1?{background:"var(--em)",borderColor:"var(--em)"}:{}}>{sendStep>1?<Ic n="check" s={14}/>:"1"}</div><div style={{fontSize:10,color:"var(--t3)",marginTop:4,fontWeight:600}}>Destinatario</div></div>
                  <div className={"step-line"+(sendStep>=2?" done":"")}/>
                  <div style={{textAlign:"center",flex:1}}><div className={"step-dot"+(sendStep>=2?" on":"")}>2</div><div style={{fontSize:10,color:"var(--t3)",marginTop:4,fontWeight:600}}>Productos</div></div>
                </div>
                {sendStep===1&&(
                  <div className="card">
                    <div className="card-h"><div className="card-title">¿A quién le enviás?</div></div>
                    <div style={{padding:14}}>
                      {contacts.length===0?(<div className="empty"><div style={{fontSize:32,marginBottom:8}}>👥</div><div style={{fontWeight:700,marginBottom:6}}>Sin contactos</div><button className="btn b-am" onClick={function(){setTab("contacts");}}>Ir a Contactos</button></div>):contacts.map(function(c){ return (<div key={c.id} className={"ct-card"+(sendTo===c.id?" sel":"")} onClick={function(){setSendTo(c.id);setSendStep(2);}}><Avatar name={c.name} color={c.color} size={46}/><div style={{flex:1}}><div style={{fontWeight:700,fontSize:14}}>{c.name}</div><div style={{fontSize:11,color:"var(--t3)",marginTop:2}}>{c.email}</div></div><Ic n="send" s={18}/></div>); })}
                    </div>
                  </div>
                )}
                {sendStep===2&&(function(){ var tUser=contacts.find(function(c){ return c.id===sendTo; }); var cartTotal=Object.values(sendCart).reduce(function(s,v){ return s+Number(v); },0); return (<div>
                  <div className="row g12" style={{padding:"10px 14px",background:"var(--card)",borderRadius:14,border:"1px solid var(--brd)",marginBottom:14}}><Avatar name={tUser?tUser.name:"?"} color={tUser?tUser.color:"#999"} size={36}/><div style={{flex:1}}><div style={{fontWeight:700,fontSize:13}}>{tUser?tUser.name:""}</div></div><button className="btn btn-xs b-ghost" onClick={function(){setSendStep(1);setSendCart({});}}>Cambiar</button></div>
                  <SearchBar value={sendSrch} onChange={setSendSrch} placeholder="Buscar producto..."/>
                  {sendFiltered.length===0?<div className="empty">Sin stock</div>:sendFiltered.map(function(item){ var p=item.products||products.find(function(x){ return x.id===item.product_id; }); if(!p) return null; var qty=Number(sendCart[item.id]||0); return (<div key={item.id} className="sic"><ProdThumb prod={p} size={38}/><div style={{flex:1,minWidth:0}}><div style={{fontWeight:600,fontSize:13,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.name}</div><div style={{fontSize:11,color:"var(--t3)",marginTop:2}}>{item.qty_available} disp. · {fmtARS(p.price)}</div></div><QtyControl val={qty} set={function(v){setSendCart(function(c){ var n=Object.assign({},c);n[item.id]=Math.min(item.qty_available,v);return n; });}} min={0} max={item.qty_available}/></div>); })}
                  {cartTotal>0&&(<div className="cart-box"><div style={{fontSize:11,fontWeight:700,color:"var(--em-d)",marginBottom:6,textTransform:"uppercase",letterSpacing:".08em"}}>Resumen</div>{Object.entries(sendCart).filter(function(e){ return Number(e[1])>0; }).map(function(e){ var ii=myStock.find(function(i){ return i.id===e[0]; }); var p=ii?(ii.products||products.find(function(x){ return x.id===ii.product_id; })):null; return p?<div key={e[0]} className="row jb" style={{fontSize:12,color:"var(--t2)",marginBottom:3}}><span>{p.emoji} {p.name}</span><span style={{fontWeight:700}}>{e[1]} u.</span></div>:null; })}<div className="row jb" style={{fontWeight:800,fontSize:14,color:"var(--em-d)",marginTop:8,paddingTop:8,borderTop:"1px solid var(--em-t)"}}><span>Total</span><span>{cartTotal} u.</span></div></div>)}
                  <button className="cta cta-em" onClick={doMultiSend} disabled={cartTotal===0}><Ic n="send" s={18}/>Confirmar envío ({cartTotal} u.)</button>
                </div>); })()}
              </div>
            </div>
          )}

          {/* ══ IMPORTAR ══ */}
          {tab==="importar"&&(
            <div>
              <div className="ph"><div><div className="ph-h">Importar</div><div className="ph-s">Carga tu lista desde Excel</div></div>{impRows.some(function(r){return r.ok;})&&<button className="btn b-em" style={{fontSize:12,padding:"8px 14px"}} onClick={doConfirmImport}><Ic n="check" s={14}/>Confirmar ({impRows.filter(function(r){return r.ok;}).length})</button>}</div>
              <div className="pc">
                <div className="card">
                  <div className="card-h"><div className="card-title">Modo de importación</div></div>
                  <div style={{padding:"12px 14px 0"}}>
                    <div className="ql-tabs"><div className={"ql-tab"+(impTab==="file"?" on":"")} onClick={function(){setImpTab("file");}}>Subir archivo</div><div className={"ql-tab"+(impTab==="txt"?" on":"")} onClick={function(){setImpTab("txt");}}>Pegar texto</div></div>
                    <div className="fld"><label className="fl">Unidades por producto</label><QtyControl val={impQty} set={setImpQty} min={1}/></div>
                    {impTab==="file"?(
                      <div><div className="drop-z" style={{marginBottom:14}}><input type="file" accept=".csv,.txt,.xls,.xlsx" onChange={doParseFile}/><div style={{fontSize:40,marginBottom:8}}>📊</div><div style={{fontSize:14,fontWeight:700,color:"var(--in)"}}>Tocar para subir archivo</div><div style={{fontSize:12,color:"var(--t3)",marginTop:4}}>.xlsx · .xls · .csv · .txt</div>{impFile&&<div style={{marginTop:10,background:"var(--in-l)",color:"var(--in-d)",borderRadius:20,padding:"4px 12px",display:"inline-flex",alignItems:"center",gap:6,fontSize:12,fontWeight:700}}><Ic n="check" s={12}/>{impFile}</div>}</div></div>
                    ):(
                      <div><div style={{fontFamily:"var(--mf)",fontSize:11,background:"var(--bg2)",border:"1px solid var(--brd)",borderRadius:8,padding:"10px 12px",marginBottom:10,color:"var(--t2)",lineHeight:1.9}}>{"KCG1 - KALOE Crema Gel x 50gr.\t$ 11.900,00"}<br/>{"BF100 - 212 Rosa x 80 ml.\t$ 40.990,00"}</div><textarea style={{width:"100%",height:160,padding:"12px 14px",borderRadius:10,border:"1.5px solid var(--brd)",background:"var(--card)",color:"var(--t1)",fontFamily:"var(--mf)",fontSize:12,resize:"none",outline:"none",lineHeight:1.7}} placeholder={"KCG1 - Producto\t11900"} value={impTxt} onChange={function(e){setImpTxt(e.target.value);setImpRows([]);}} /><button className="btn b-in" style={{marginTop:8,width:"100%",justifyContent:"center",padding:12}} onClick={doParseTxt} disabled={!impTxt.trim()}><Ic n="search" s={14}/>Previsualizar</button></div>
                    )}
                    {impRows.length>0&&(<div style={{marginTop:14,marginBottom:14}}><div style={{fontSize:11,fontWeight:700,color:"var(--t3)",textTransform:"uppercase",marginBottom:8}}>{impRows.filter(function(r){return r.ok;}).length} OK · {impRows.filter(function(r){return !r.ok;}).length} errores</div>{impRows.slice(0,15).map(function(row,i){ return (<div key={i} className={"prev-row "+(row.ok?"prev-ok":"prev-err")}><div style={{width:7,height:7,borderRadius:"50%",background:row.ok?"var(--em)":"var(--cr)",flexShrink:0}}/>{row.ok?(<div className="row g8" style={{flex:1,flexWrap:"wrap"}}><span style={{fontFamily:"var(--mf)",fontSize:11,fontWeight:700,color:"var(--in-d)"}}>{row.sku}</span><span style={{flex:1,fontSize:12}}>{row.name}</span><span style={{fontFamily:"var(--mf)",fontSize:11,fontWeight:700,color:"var(--am-d)"}}>{fmtARS(row.price)}</span><span style={{fontSize:10,color:"var(--t3)",background:"var(--bg2)",borderRadius:5,padding:"2px 6px"}}>{row.qty>0?row.qty+" u.":"solo catálogo"}</span></div>):(<div className="row g8" style={{flex:1}}><span style={{flex:1,fontFamily:"var(--mf)",fontSize:11,color:"var(--t3)",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{row.raw}</span><span style={{fontSize:11,color:"var(--cr)",flexShrink:0}}>{row.err}</span></div>)}</div>); })}<button className="cta cta-em" onClick={doConfirmImport} disabled={!impRows.some(function(r){return r.ok;})}><Ic n="check" s={18}/>Importar {impRows.filter(function(r){return r.ok;}).length}</button></div>)}
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* ══ CATÁLOGO ══ */}
          {tab==="catalog"&&(
            <div>
              <div className="ph"><div><div className="ph-h">Catálogo</div><div className="ph-s">{allProds.length||products.length} productos · {isAdmin?"ABM completo":"solo lectura"}</div></div>{isAdmin&&<button className="btn btn-xs b-em" onClick={function(){setTab("importar");}}><Ic n="upload" s={13}/>Importar</button>}</div>
              <div className="pc">
                <SearchBar value={srchCat} onChange={setSrchCat} placeholder="Buscar por nombre, SKU o categoría..."/>
                {isAdmin&&(
                  <div className="card">
                    <div className="card-h"><div className="card-title">{editP?"✏️ Modificar":"✨ Nuevo producto"}</div>{editP&&<button className="btn btn-xs b-ghost" onClick={cancelEdit}>Cancelar</button>}</div>
                    <div style={{padding:"12px 14px"}}>
                      <form onSubmit={doSaveProd}>
                        <div className="row g8" style={{marginBottom:12}}><div style={{flex:1}}><label className="fl">SKU</label><input className="fi" placeholder="Ej: KCG1" value={fSku} onChange={function(e){setFSku(e.target.value);}}/></div><div style={{flex:2}}><label className="fl">Nombre</label><input className="fi" placeholder="Nombre del producto" value={fName} onChange={function(e){setFName(e.target.value);}}/></div></div>
                        <div className="row g8" style={{marginBottom:12}}><div style={{flex:1}}><label className="fl">Precio ($)</label><input className="fi" type="number" step="0.01" placeholder="0.00" value={fPrice} onChange={function(e){setFPrice(e.target.value);}}/></div><div style={{flex:1}}><label className="fl">Icono</label><select className="fi fi-sel" value={fEmoji} onChange={function(e){setFEmoji(e.target.value);}}>{EMOJIS.map(function(o){ return <option key={o.v} value={o.v}>{o.v} {o.l}</option>; })}</select></div></div>
                        <div className="fld"><label className="fl">Categoría</label><select className="fi fi-sel" value={fCat} onChange={function(e){setFCat(e.target.value);}}>{CATS.map(function(c){ return <option key={c}>{c}</option>; })}</select></div>
                        <div className="fld"><label className="fl">Foto (opcional)</label><div className="photo-zone"><input type="file" accept="image/*" onChange={function(e){handlePhoto(e,setFPhoto);}}/>{fPhoto?<img src={fPhoto} alt="" style={{width:56,height:56,borderRadius:10,objectFit:"cover",margin:"0 auto"}}/>:<div><div style={{fontSize:24,marginBottom:4}}>📸</div><div style={{fontSize:12,fontWeight:700,color:"var(--in)"}}>Tocar para subir foto</div></div>}</div>{fPhoto&&<button type="button" onClick={function(){setFPhoto(null);}} style={{fontSize:11,color:"var(--cr)",background:"none",border:"none",cursor:"pointer",fontWeight:700}}>Quitar foto</button>}</div>
                        {!editP&&<div className="fld"><label className="fl">Mi stock inicial</label><input className="fi" type="number" min="0" value={fStock} onChange={function(e){setFStock(e.target.value);}}/></div>}
                        <button type="submit" className="cta cta-in" style={{marginTop:10}}><Ic n="check" s={18}/>{editP?"Guardar cambios":"Dar de alta"}</button>
                      </form>
                    </div>
                  </div>
                )}
                <div className="card">
                  <div className="card-h"><div className="card-title">Catálogo ({catFiltered.length})</div></div>
                  {catFiltered.length===0?<div className="empty">Sin productos.{srchCat?" Intenta otra búsqueda.":" El administrador aún no cargó productos."}</div>:(
                    <div className="tw"><table><thead><tr><th></th><th>SKU</th><th>Producto</th><th>Precio</th>{isAdmin&&<th></th>}</tr></thead><tbody>{catFiltered.map(function(p){ return (<tr key={p.id} className="tr"><td><ProdThumb prod={p} size={34}/></td><td><span style={{color:"var(--in-d)",fontFamily:"var(--mf)",fontSize:11,background:"var(--in-l)",padding:"2px 6px",borderRadius:5,fontWeight:600}}>{p.sku}</span></td><td><div style={{fontWeight:600,fontSize:12}}>{p.name}</div><div style={{fontSize:10,color:"var(--t3)"}}>{p.category}</div></td><td style={{fontFamily:"var(--mf)",fontSize:12,fontWeight:700}}>{fmtARS(p.price)}</td>{isAdmin&&<td><div className="row g8" style={{justifyContent:"flex-end"}}><button className="btn btn-xs b-in" onClick={function(){startEdit(p);}}><Ic n="edit" s={12}/></button>{delConf===p.id?<div className="row g8"><button className="btn btn-xs b-cr" onClick={function(){doDelProd(p.id,p.sku);}}>Sí</button><button className="btn btn-xs b-ghost" onClick={function(){setDelConf(null);}}>No</button></div>:<button className="btn btn-xs b-cr" onClick={function(){setDelConf(p.id);}}><Ic n="trash" s={12}/></button>}</div></td>}</tr>); })}</tbody></table></div>
                  )}
                </div>
              </div>
            </div>
          )}

          {/* ══ CONTACTOS ══ */}
          {tab==="contacts"&&(
            <div>
              <div className="ph"><div><div className="ph-h">Mi Red</div><div className="ph-s">Contactos y notificaciones</div></div></div>
              <div className="pc">
                {/* Notificaciones */}
                {myNotifs.length>0&&(
                  <div className="card" style={{border:myNotifs.some(function(n){return !n.read;})?"2px solid var(--am)":"1px solid var(--brd)"}}>
                    <div className="card-h"><div className="card-title">🔔 Notificaciones</div><div className="row g8">{myNotifs.some(function(n){return !n.read;})&&<span className="badge b-am">{myNotifs.filter(function(n){return !n.read;}).length} nuevas</span>}<button className="btn btn-xs b-ghost" onClick={async function(){await sb.from("notifications").update({read:true}).eq("to_user_id",me.id);setNotifs(function(p){return p.map(function(n){return Object.assign({},n,{read:true});});});}}>Leídas</button></div></div>
                    <div style={{padding:"8px 14px"}}>{myNotifs.map(function(n){ var ico=n.type==="transfer"?"📦":n.type==="confirm"?"✅":n.type==="sale"?"💰":"🔔"; return (<div key={n.id} style={{display:"flex",gap:10,padding:"10px 0",borderBottom:"1px solid var(--brd)",opacity:n.read?0.6:1}}><div style={{fontSize:20,flexShrink:0}}>{ico}</div><div style={{flex:1}}><div style={{fontSize:13,fontWeight:n.read?400:700,color:"var(--t1)"}}>{n.message}</div><div style={{fontSize:10,color:"var(--t3)",marginTop:2}}>{new Date(n.created_at).toLocaleString("es-AR")}</div></div>{!n.read&&<div style={{width:8,height:8,borderRadius:"50%",background:"var(--am)",flexShrink:0,marginTop:4}}/>}</div>); })}</div>
                  </div>
                )}

                <div className="card">
                  <div className="card-h"><div className="card-title"><div className="card-ico" style={{background:"var(--em-l)",color:"var(--em-d)"}}><Ic n="plus" s={14}/></div>Agregar contacto</div></div>
                  <div style={{padding:"12px 14px"}}>
                    <div style={{fontSize:12,color:"var(--t3)",marginBottom:10,lineHeight:1.5}}>Ingresa el <strong>email exacto</strong> de la persona registrada en StockPro.</div>
                    <div className="row g8"><input className="fi" style={{flex:1}} type="email" placeholder="email@ejemplo.com" value={ctQ} onChange={function(e){setCtQ(e.target.value);}} onKeyDown={function(e){if(e.key==="Enter")doAddContact();}}/><button className="btn b-pri" onClick={doAddContact}><Ic n="plus" s={14}/>Agregar</button></div>
                  </div>
                </div>

                <SearchBar value={srchCon} onChange={setSrchCon} placeholder="Buscar contacto..."/>
                <div className="card">
                  <div className="card-h"><div className="card-title">Tu red ({conFiltered.length})</div></div>
                  {conFiltered.length===0?<div className="empty">{srchCon?"Sin resultados.":"Aún no tienes contactos."}</div>:(
                    <div style={{padding:"10px 14px"}}>{conFiltered.map(function(c){ return (<div key={c.id} className="ct-card" style={{cursor:"default"}}><Avatar name={c.name} color={c.color} size={44}/><div style={{flex:1}}><div style={{fontWeight:700,fontSize:14}}>{c.name}</div><div style={{fontSize:11,color:"var(--t3)",marginTop:2}}>{c.email}</div></div><div style={{background:"var(--em-l)",color:"var(--em-d)",borderRadius:20,padding:"3px 9px",fontSize:11,fontWeight:700,border:"1px solid var(--em-t)",display:"flex",alignItems:"center",gap:4}}><Ic n="check" s={11}/>OK</div></div>); })}</div>
                  )}
                </div>

                <div className="card">
                  <div className="card-h"><div className="card-title"><div className="card-ico" style={{background:"var(--in-l)",color:"var(--in-d)"}}><Ic n="clock" s={14}/></div>Mi historial</div><span className="badge b-in">{logFiltered.length}</span></div>
                  <div style={{padding:"8px 14px 4px"}}><SearchBar value={srchLog} onChange={setSrchLog} placeholder="Buscar en historial..."/></div>
                  {logFiltered.length===0?<div className="empty">Sin movimientos</div>:(
                    <div className="tw"><table><thead><tr><th>Fecha</th><th>Producto</th><th>Cant.</th><th>Tipo</th></tr></thead><tbody>{logFiltered.slice(0,50).map(function(l){ var p=l.product; return (<tr key={l.id} className="tr"><td style={{fontFamily:"var(--mf)",fontSize:10,color:"var(--t3)",whiteSpace:"nowrap"}}>{new Date(l.created_at).toLocaleString("es-AR",{day:"2-digit",month:"2-digit",hour:"2-digit",minute:"2-digit"})}</td><td style={{fontWeight:600,fontSize:12}}>{p?p.name:""}</td><td style={{fontFamily:"var(--mf)",fontWeight:700,color:"var(--am-d)"}}>{l.qty}</td><td><span style={{background:"var(--in-l)",color:"var(--in-d)",borderRadius:5,padding:"2px 7px",fontSize:10,fontWeight:700,fontFamily:"var(--mf)"}}>{l.source}</span></td></tr>); })}</tbody></table></div>
                  )}
                </div>
              </div>
            </div>
          )}

          {/* ══ ADMIN DASHBOARD ══ */}
          {tab==="admin"&&isAdmin&&(
            <div>
              <div className="ph"><div><div className="ph-h">👑 Panel Admin</div><div className="ph-s">Métricas globales de StockPro</div></div><button className="btn btn-xs b-ghost" onClick={function(){loadData(me.id,me.role);}}><Ic n="undo" s={13}/>Actualizar</button></div>
              <div className="pc">
                {adminStats&&(
                  <div className="metric-grid">
                    {[["Usuarios registrados",adminStats.total_users,"var(--in)"],["Nuevos (30 días)",adminStats.new_users_30d,"var(--em)"],["Productos activos",adminStats.total_products,"var(--am)"],["Ventas este mes",adminStats.units_sold_30d,"var(--cr)"],["Envíos pendientes",adminStats.pending_transfers,"var(--in)"],["Stock total sistema",adminStats.total_stock_system,"var(--em)"]].map(function(m){ return (<div key={m[0]} className="metric-card"><div className="metric-val" style={{color:m[2]}}>{m[1]||0}</div><div className="metric-lbl">{m[0]}</div></div>); })}
                  </div>
                )}
                <div className="card">
                  <div className="card-h"><div className="card-title"><Ic n="users" s={14}/>Usuarios registrados ({allUsers.length})</div></div>
                  {allUsers.length===0?<div className="empty">Sin usuarios aún</div>:(
                    <div className="tw"><table><thead><tr><th>Nombre</th><th>Email</th><th>Rol</th><th>Registro</th></tr></thead><tbody>{allUsers.map(function(u){ return (<tr key={u.id} className="tr"><td><div className="row g8"><Avatar name={u.name} color={u.color} size={28}/><span style={{fontWeight:600,fontSize:12}}>{u.name}</span></div></td><td style={{fontSize:11,color:"var(--t3)"}}>{u.email}</td><td><span style={{background:u.role==="superadmin"?"var(--am-t)":"var(--in-t)",color:u.role==="superadmin"?"var(--am-d)":"var(--in-d)",borderRadius:5,padding:"2px 7px",fontSize:10,fontWeight:700}}>{u.role==="superadmin"?"👑 Admin":"Usuario"}</span></td><td style={{fontSize:10,color:"var(--t3)"}}>{new Date(u.created_at).toLocaleDateString("es-AR")}</td></tr>); })}</tbody></table></div>
                  )}
                </div>
              </div>
            </div>
          )}

        </div>

        {/* TAB BAR */}
        <nav className="tabbar">
          {TABS.map(function(t){ return (<div key={t.id} className={"tab"+(tab===t.id?" on":"")} onClick={function(){setTab(t.id);}}><div className="tab-bub"><Ic n={t.ico} s={20}/></div><span className="tab-lbl">{t.lbl}</span></div>); })}
        </nav>
      </div>

      {/* TRANSFER MODAL */}
      {txModal&&(
        <div className="ovl" onClick={function(e){if(e.target===e.currentTarget)setTxModal(null);}}>
          <div className="mbox">
            <div className="mhd"><div className="mhd-t">Transferir</div><button className="ic-btn" onClick={function(){setTxModal(null);}}><Ic n="x" s={16}/></button></div>
            <div className="mbd">
              {(function(){ var prod=txModal.products||products.find(function(p){return p.id===txModal.product_id;}); return (<div>
                <div className="row g12" style={{padding:"12px 14px",background:"var(--bg2)",borderRadius:10,border:"1px solid var(--brd)",marginBottom:14}}><ProdThumb prod={prod} size={44}/><div style={{flex:1}}><div style={{fontWeight:700,fontSize:14}}>{prod?prod.name:""}</div><div style={{fontSize:11,color:"var(--t3)",marginTop:2}}>{fmtARS(prod?prod.price:0)}</div></div><div style={{fontFamily:"var(--mf)",fontWeight:800,fontSize:15,color:"var(--in)"}}>{fmtARS(prod?(prod.price*txQty):0)}</div></div>
                <div className="fld"><label className="fl">Destinatario</label><select className="fi fi-sel" value={txTo} onChange={function(e){setTxTo(e.target.value);}}>{contacts.map(function(c){ return <option key={c.id} value={c.id}>{c.name}</option>; })}</select></div>
                <div className="fld"><label className="fl">Cantidad ({txModal.qty_available} disponibles)</label><QtyControl val={txQty} set={setTxQty} max={txModal.qty_available} big={true}/></div>
              </div>); })()}
            </div>
            <div className="mft"><button className="btn b-ghost" onClick={function(){setTxModal(null);}}>Cancelar</button><button className="btn b-pri" onClick={doTx} disabled={contacts.length===0}><Ic n="send" s={14}/>Confirmar</button></div>
          </div>
        </div>
      )}

      {/* SHARE CATALOG MODAL */}
      {shareM&&(
        <div className="ovl" onClick={function(e){if(e.target===e.currentTarget){setShareM(false);setShareSel({});}}}>
          <div className="mbox">
            <div className="mhd"><div className="mhd-t row g8"><Ic n="wa" s={20}/>Compartir por WhatsApp</div><button className="ic-btn" onClick={function(){setShareM(false);setShareSel({});}}><Ic n="x" s={16}/></button></div>
            <div className="mbd">
              <div style={{fontSize:13,color:"var(--t2)",marginBottom:12}}>Selecciona los productos para incluir en el mensaje.</div>
              <div className="row jb" style={{marginBottom:12}}><span style={{fontSize:12,fontWeight:700,color:"var(--t3)"}}>{Object.keys(shareSel).length} seleccionados</span><button className="btn btn-xs b-in" onClick={function(){var all={};myStock.forEach(function(i){all[i.id]=true;});setShareSel(all);}}>Seleccionar todo</button></div>
              <div style={{maxHeight:360,overflowY:"auto",display:"flex",flexDirection:"column",gap:8}}>
                {myStock.map(function(item){ var p=item.products||products.find(function(x){return x.id===item.product_id;}); if(!p) return null; var sel=!!shareSel[item.id]; return (<div key={item.id} onClick={function(){setShareSel(function(prev){var n=Object.assign({},prev);if(n[item.id])delete n[item.id];else n[item.id]=true;return n;});}} style={{display:"flex",alignItems:"center",gap:12,padding:"12px 14px",borderRadius:12,border:"2px solid "+(sel?"var(--wa)":"var(--brd)"),background:sel?"var(--wa-l)":"var(--card)",cursor:"pointer",transition:"all .15s"}}><ProdThumb prod={p} size={44}/><div style={{flex:1,minWidth:0}}><div style={{fontWeight:700,fontSize:13,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.name}</div><div style={{fontSize:11,color:"var(--t3)",marginTop:2}}>{fmtARS(p.price)} · {item.qty_available} u.{p.photo_url?" · 📸":""}</div></div><div style={{width:24,height:24,borderRadius:"50%",background:sel?"var(--wa)":"var(--bg2)",border:"2px solid "+(sel?"var(--wa)":"var(--brd2)"),display:"flex",alignItems:"center",justifyContent:"center",color:"#fff",flexShrink:0}}>{sel?<Ic n="check" s={13}/>:null}</div></div>); })}
              </div>
            </div>
            <div className="mft"><button className="btn b-ghost" onClick={function(){setShareM(false);setShareSel({});}}>Cancelar</button><button style={{background:"linear-gradient(135deg,var(--wa),var(--wa-d))",color:"#fff",border:"none",borderRadius:10,padding:"12px 18px",fontWeight:800,fontSize:14,cursor:"pointer",display:"flex",alignItems:"center",gap:8,boxShadow:"0 4px 14px rgba(37,211,102,.35)",opacity:Object.keys(shareSel).length===0?0.4:1}} onClick={doShareCatalog} disabled={Object.keys(shareSel).length===0}><Ic n="wa" s={16}/>Compartir</button></div>
          </div>
        </div>
      )}

      <div className="toast-wrap">{toasts.map(function(t){ return <Toast key={t.id} t={t} remove={rmToast}/>; })}</div>
    </div>
  );
}
