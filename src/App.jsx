import React, { useState, useEffect, useCallback, memo, useMemo } from "react";
import ConsignacionModule from "./ConsignacionModule";
import * as XLSX from "xlsx";
import { createClient } from "@supabase/supabase-js";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";

// ─── SUPABASE ─────────────────────────────────────────────────────────────────
// Requiere archivo .env en la raíz:
//   VITE_SUPABASE_URL=https://xxxx.supabase.co
//   VITE_SUPABASE_ANON_KEY=eyJhbGci...
if (!import.meta.env.VITE_SUPABASE_URL || !import.meta.env.VITE_SUPABASE_ANON_KEY) {
  console.error("⚠️  Faltan variables de entorno VITE_SUPABASE_URL y/o VITE_SUPABASE_ANON_KEY");
}
const sb = createClient(
  import.meta.env.VITE_SUPABASE_URL,
  import.meta.env.VITE_SUPABASE_ANON_KEY
);

// ─── CONSTANTES ───────────────────────────────────────────────────────────────
var ROLES  = ["user","distributor","reseller"];
const CATS   = ["General","Linea Kaloe","Pocket x 10ml","Beauty Collagen","Premium Femenino","Exhibidores","Joyas","Cosmetica"];
const EMOJIS = [{v:"✨",l:"Brillo"},{v:"🧴",l:"Crema"},{v:"💼",l:"Maletin"},{v:"💄",l:"Maquillaje"},{v:"💎",l:"Joya"},{v:"📿",l:"Cadena"},{v:"🌸",l:"Floral"},{v:"🎀",l:"Accesorio"},{v:"🍃",l:"Natural"},{v:"⭐",l:"Destacado"}];

// uid() is for ephemeral client IDs only (toasts, local keys). DB entities use Supabase UUIDs.
function uid() { return Date.now().toString(36) + Math.random().toString(36).slice(2); }
function fmtARS(n) { return "$ " + Number(n||0).toLocaleString("es-AR",{minimumFractionDigits:2}); }
function ini(name) { return (name||"?").trim().split(" ").slice(0,2).map(function(w){return w[0];}).join("").toUpperCase(); }

// ─── LOGO ─────────────────────────────────────────────────────────────────────
const LOGO_URL = "https://cgqrgyouunxfluujhwey.supabase.co/storage/v1/object/public/assets/logo.png";

// ─── ICONS (memoized above before CSS) ───────────────────────────────────────

// Memoize Ic so it doesn't re-render when parent re-renders with same props
const IcBase = function Ic(props) {
  var n=props.n, s=props.s||18;
  var p={width:s,height:s,viewBox:"0 0 24 24",fill:"none",stroke:"currentColor",strokeWidth:"2.2",strokeLinecap:"round",strokeLinejoin:"round"};
  switch(n) {
    case "box":     return <svg {...p}><polyline points="21,8 21,21 3,21 3,8"/><rect x="1" y="3" width="22" height="5"/></svg>;
    case "plus":    return <svg {...p}><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>;
    case "home":    return <svg {...p}><path d="M3 9.5L12 3l9 6.5"/><path d="M5 10v10h14V10"/><path d="M9 20v-6h6v6"/></svg>;
    case "dots":    return <svg {...p}><circle cx="5" cy="12" r="1.6"/><circle cx="12" cy="12" r="1.6"/><circle cx="19" cy="12" r="1.6"/></svg>;
    case "send":    return <svg {...p}><line x1="22" y1="2" x2="11" y2="13"/><polygon points="22,2 15,22 11,13 2,9 22,2"/></svg>;
    case "upload":  return <svg {...p}><polyline points="16,16 12,12 8,16"/><line x1="12" y1="12" x2="12" y2="21"/><path d="M20.39 18.39A5 5 0 0 0 18 9h-1.26A8 8 0 1 0 3 16.3"/></svg>;
    case "list":    return <svg {...p}><line x1="8" y1="6" x2="21" y2="6"/><line x1="8" y1="12" x2="21" y2="12"/><line x1="8" y1="18" x2="21" y2="18"/><line x1="3" y1="6" x2="3.01" y2="6"/><line x1="3" y1="12" x2="3.01" y2="12"/><line x1="3" y1="18" x2="3.01" y2="18"/></svg>;
    case "users":   return <svg {...p}><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>;
    case "user":    return <svg {...p}><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/></svg>;
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
};
const Ic = memo(IcBase);

// ─── CSS ─────────────────────────────────────────────────────────────────────
const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Nunito:wght@400;500;600;700;800;900&family=JetBrains+Mono:wght@400;600;700&display=swap');
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{
  /* ── Fondos: blanco limpio con gris muy claro para cards ── */
  --bg:#ffffff;--bg2:#f3f3f6;--card:#ffffff;
  --brd:#ececf0;--brd2:#dcdce2;

  /* ── Tipografía ── */
  --t1:#1a1a2e;--t2:#44444f;--t3:#8a8a96;--t4:#b4b4be;

  /* ── Rosa magenta vibrante como color de marca (estilo delivery) ── */
  --in:#e0224e;--in-d:#c11743;--in-l:#fff0f3;--in-t:rgba(224,34,78,.10);
  --in-m:#fa1e5a;

  /* ── Esmeralda para éxito ── */
  --em:#10b981;--em-d:#059669;--em-l:#ecfdf5;--em-t:rgba(16,185,129,.09);

  /* ── Ámbar para alertas ── */
  --am:#f59e0b;--am-d:#d97706;--am-l:#fffbeb;--am-t:rgba(245,158,11,.09);

  /* ── Rojo para errores ── */
  --cr:#ef4444;--cr-d:#dc2626;--cr-l:#fef2f2;--cr-t:rgba(239,68,68,.08);

  /* ── Azul cielo para info ── */
  --bl:#0ea5e9;--bl-d:#0284c7;--bl-l:#f0f9ff;--bl-t:rgba(14,165,233,.09);

  /* ── WhatsApp ── */
  --wa:#16a34a;--wa-d:#15803d;--wa-l:#f0fdf4;

  /* ── Violeta ── */
  --pu:#8b5cf6;--pu-d:#7c3aed;--pu-l:#f3eafe;

  /* ── Primario = rosa magenta ── */
  --pri:#e0224e;--pri-d:#c11743;--pri-l:#fff0f3;

  /* ── Gradiente magenta vibrante ── */
  --grad:linear-gradient(145deg,#e0224e 0%,#e0224e 50%,#fa1e5a 100%);

  /* ── Tipografías ── */
  --hf:'Nunito',sans-serif;--mf:'JetBrains Mono',monospace;

  /* ── Radio de bordes: muy redondeados (estilo delivery) ── */
  --r:16px;--r2:22px;--r3:30px;

  /* ── Sombras muy suaves y difusas ── */
  --sh:0 1px 4px rgba(0,0,0,.04),0 4px 16px rgba(0,0,0,.04);
  --sh2:0 4px 20px rgba(0,0,0,.08),0 1px 4px rgba(0,0,0,.04);
  --sh3:0 20px 60px rgba(0,0,0,.12),0 4px 16px rgba(0,0,0,.05);
  --tab:66px;
}
html,body{height:100%;background:var(--bg);color:var(--t1);font-family:var(--hf);font-size:14px;-webkit-font-smoothing:antialiased}
@keyframes fadeUp{from{opacity:0;transform:translateY(18px)}to{opacity:1;transform:translateY(0)}}
@keyframes toastIn{from{opacity:0;transform:translateY(20px) scale(.94)}to{opacity:1;transform:translateY(0) scale(1)}}
@keyframes shake{0%,100%{transform:translateX(0)}20%,60%{transform:translateX(-8px)}40%,80%{transform:translateX(8px)}}
@keyframes spin{to{transform:rotate(360deg)}}
@keyframes pulse{0%,100%{opacity:1}50%{opacity:.6}}

.app{display:flex;flex-direction:column;height:100vh;overflow:hidden;background:var(--bg)}
.main{flex:1;overflow-y:auto;padding-bottom:var(--tab)}
.spin{animation:spin 1s linear infinite}

/* HEADER */
.hdr{background:var(--card);padding:0;flex-shrink:0;position:sticky;top:0;z-index:60;box-shadow:0 1px 0 var(--brd)}
.hdr-btn{width:44px;height:44px;border-radius:14px;border:1.5px solid var(--brd);background:var(--card);cursor:pointer;color:var(--t2);display:flex;align-items:center;justify-content:center;transition:all .18s;flex-shrink:0}
.hdr-btn:active{background:var(--bg2);transform:scale(.94)}

.hdr-top{display:flex;align-items:center;justify-content:space-between;padding:16px 16px 10px;gap:12px}
.hdr-search{margin:0 16px 14px;position:relative}
.hdr-search input{width:100%;padding:14px 18px 14px 46px;border-radius:16px;border:1.5px solid var(--brd);background:var(--card);color:var(--t1);font-family:var(--hf);font-size:14px;outline:none;font-weight:600}
.hdr-search input:focus{border-color:var(--in-m)}
.hdr-search-ico{position:absolute;left:16px;top:50%;transform:translateY(-50%);color:var(--t3);pointer-events:none}
.hdr-avatar{width:46px;height:46px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-weight:900;font-size:16px;color:#fff;flex-shrink:0;background:var(--grad)}
.hdr-hi{font-size:17px;font-weight:900;color:var(--t1);letter-spacing:-.02em;line-height:1.15}
.hdr-sub{font-size:12px;color:var(--t3);font-weight:600;margin-top:2px;line-height:1.3}
/* TABBAR */
.tabbar{position:fixed;bottom:0;left:0;right:0;height:var(--tab);background:rgba(255,255,255,.96);backdrop-filter:blur(20px);-webkit-backdrop-filter:blur(20px);display:flex;z-index:50;box-shadow:0 -1px 0 var(--brd),0 -10px 30px rgba(0,0,0,.08);border-radius:24px 24px 0 0;padding-bottom:env(safe-area-inset-bottom,0px)}
.tab{flex:1;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:2px;cursor:pointer;color:var(--t4);transition:all .22s cubic-bezier(.34,1.56,.64,1);position:relative;padding:8px 2px 5px;min-width:0;-webkit-tap-highlight-color:transparent;touch-action:manipulation}
.tab.on{color:var(--pri)}
.tab-bub{width:46px;height:28px;border-radius:14px;display:flex;align-items:center;justify-content:center;transition:all .22s cubic-bezier(.34,1.56,.64,1)}
.tab.on .tab-bub{background:var(--in-l);transform:scale(1.05)}
.tab-lbl{font-size:9px;font-weight:700;letter-spacing:.01em;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:60px;transition:all .2s}
.tab.on .tab-lbl{color:var(--in-m)}
.tab-dot{position:absolute;top:5px;right:calc(50% - 22px);width:7px;height:7px;border-radius:50%;background:var(--cr);border:2px solid white;box-shadow:0 0 0 1px var(--cr-l)}
.tab-fab-wrap{flex:1;display:flex;align-items:center;justify-content:center;cursor:pointer;position:relative}
.tab-fab{width:56px;height:56px;border-radius:50%;background:var(--grad);display:flex;align-items:center;justify-content:center;color:#fff;box-shadow:0 6px 20px rgba(224,34,78,.42);margin-top:-26px;border:4px solid var(--card);transition:transform .18s cubic-bezier(.34,1.56,.64,1)}
.tab-fab-wrap:active .tab-fab{transform:scale(.9)}

/* PAGE */
.ph{padding:20px 16px 10px;display:flex;align-items:flex-end;justify-content:space-between}
.ph-h{font-size:24px;font-weight:900;color:var(--t1);letter-spacing:-.04em}
.ph-s{font-size:12px;color:var(--t3);margin-top:3px;font-weight:500;letter-spacing:.01em}
.pc{padding:0 14px 28px;animation:fadeUp .28s cubic-bezier(.16,1,.3,1) both}

/* CARDS */
.card{background:var(--card);border-radius:var(--r2);border:1px solid rgba(0,0,0,.055);overflow:hidden;margin-bottom:12px;box-shadow:var(--sh)}
.card-h{display:flex;align-items:center;justify-content:space-between;padding:15px 17px;border-bottom:1px solid var(--brd)}
.card-title{font-size:14px;font-weight:800;color:var(--t1);display:flex;align-items:center;gap:9px;letter-spacing:-.015em}
.card-ico{width:32px;height:32px;border-radius:11px;display:flex;align-items:center;justify-content:center;flex-shrink:0}
.badge{display:inline-flex;align-items:center;padding:3px 9px;border-radius:20px;font-size:11px;font-weight:800}
.b-in{background:var(--in-t);color:var(--in-d)}.b-em{background:var(--em-t);color:var(--em-d)}
.b-am{background:var(--am-t);color:var(--am-d)}.b-cr{background:var(--cr-t);color:var(--cr)}.b-sl{background:var(--bg2);color:var(--t2)}

/* TABLE */
.tw{overflow-x:auto}
table{width:100%;border-collapse:collapse}
th{padding:10px 13px;text-align:left;font-size:9px;text-transform:uppercase;letter-spacing:.12em;color:var(--t4);border-bottom:1px solid var(--brd);background:var(--bg);font-weight:800;white-space:nowrap}
td{padding:12px 13px;border-bottom:1px solid rgba(0,0,0,.04);font-size:12.5px;vertical-align:middle}
tr:last-child td{border-bottom:none}
.tr{cursor:default;transition:background .12s}.tr:active td{background:var(--bg)}

/* FORMS */
.fld{margin-bottom:16px}
.fl{display:block;font-size:11px;font-weight:700;color:var(--t3);margin-bottom:6px;text-transform:uppercase;letter-spacing:.06em}
.fi{width:100%;padding:13px 15px;border-radius:var(--r);border:1.5px solid var(--brd);background:var(--card);color:var(--t1);font-family:var(--hf);font-size:14px;outline:none;transition:all .16s;font-weight:500}
.fi:focus{border-color:var(--in-m);box-shadow:0 0 0 3px var(--in-t),0 2px 8px rgba(155,28,28,.08)}
.fi-sel{appearance:none;cursor:pointer}
.pw-wrap{position:relative}
.pw-eye{position:absolute;right:12px;top:50%;transform:translateY(-50%);background:none;border:none;cursor:pointer;color:var(--t3);display:flex;padding:4px}

/* BUTTONS */
.btn{display:inline-flex;align-items:center;gap:7px;padding:9px 18px;border-radius:var(--r);border:none;cursor:pointer;font-family:var(--hf);font-size:12px;font-weight:800;transition:all .16s cubic-bezier(.34,1.56,.64,1);white-space:nowrap;flex-shrink:0;letter-spacing:.01em}
.btn:disabled{opacity:.3;cursor:not-allowed}.btn:active:not(:disabled){transform:scale(.95)}
.b-pri{background:linear-gradient(140deg,var(--pri),var(--in-m));color:#fff;box-shadow:0 4px 16px rgba(155,28,28,.32)}
.b-em{background:var(--em-l);color:var(--em-d);border:1px solid rgba(0,184,122,.2)}.b-em:hover:not(:disabled){background:var(--em);color:#fff}
.b-am{background:var(--am-l);color:var(--am-d);border:1px solid rgba(255,122,0,.2)}.b-am:hover:not(:disabled){background:var(--am);color:#fff}
.b-cr{background:var(--cr-l);color:var(--cr);border:1px solid rgba(230,57,70,.2)}.b-cr:hover:not(:disabled){background:var(--cr);color:#fff}
.b-in{background:var(--in-l);color:var(--in-d);border:1px solid rgba(124,58,237,.2)}.b-in:hover:not(:disabled){background:var(--in);color:#fff}
.b-ghost{background:var(--card);border:1.5px solid var(--brd2);color:var(--t2)}.b-ghost:hover{background:var(--bg);color:var(--t1)}
.b-wa{background:var(--wa-l);color:var(--wa-d);border:1px solid rgba(37,211,102,.2)}.b-wa:hover:not(:disabled){background:var(--wa);color:#fff}
.btn-xs{padding:5px 10px;font-size:11px;border-radius:9px}
.cta{width:100%;padding:16px;border-radius:var(--r2);border:none;font-family:var(--hf);font-size:15px;font-weight:900;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:9px;margin-top:16px;transition:all .18s cubic-bezier(.34,1.56,.64,1);letter-spacing:.01em}
.cta:active:not(:disabled){transform:scale(.97)}.cta:disabled{opacity:.35;cursor:not-allowed}
.cta-am{background:linear-gradient(140deg,#f59e0b,#d97706);color:#fff;box-shadow:0 6px 22px rgba(217,119,6,.30)}
.cta-em{background:linear-gradient(140deg,#10b981,#059669);color:#fff;box-shadow:0 6px 22px rgba(5,150,105,.28)}
.cta-wa{background:linear-gradient(140deg,#22c55e,#16a34a);color:#fff;box-shadow:0 6px 22px rgba(22,163,74,.28)}
.cta-in{background:linear-gradient(140deg,var(--in-m),var(--pri));color:#fff;box-shadow:0 6px 22px rgba(155,28,28,.30)}

/* MODAL */
.ovl{position:fixed;inset:0;background:rgba(9,9,11,.55);backdrop-filter:blur(14px);-webkit-backdrop-filter:blur(14px);z-index:200;display:flex;align-items:flex-end;justify-content:center;animation:fadeUp .22s ease}
.mbox{background:rgba(255,255,255,.98);border-radius:var(--r3) var(--r3) 0 0;width:100%;max-width:600px;max-height:93vh;overflow-y:auto;box-shadow:0 -4px 40px rgba(0,0,0,.18),var(--sh3);animation:fadeUp .32s cubic-bezier(.16,1,.3,1)}
.mhd{display:flex;align-items:center;justify-content:space-between;padding:22px 20px 14px;border-bottom:1px solid var(--brd)}
.mhd-t{font-size:17px;font-weight:900;color:var(--t1);letter-spacing:-.02em}
.mbd{padding:18px 20px}.mft{padding:14px 20px;border-top:1px solid var(--brd);display:flex;justify-content:flex-end;gap:10px}
.ic-btn{width:38px;height:38px;border-radius:12px;border:none;background:var(--bg);cursor:pointer;color:var(--t3);display:flex;align-items:center;justify-content:center;transition:all .15s}
.ic-btn:active{background:var(--brd2);color:var(--t1);transform:scale(.93)}

/* LOGIN */
.login-bg{min-height:100vh;display:flex;align-items:center;justify-content:center;background:linear-gradient(155deg,#450a0a 0%,#7f1d1d 40%,#9b1c1c 70%,#dc2626 100%);padding:20px}
.lbox{background:rgba(255,255,255,.97);backdrop-filter:blur(20px);border-radius:var(--r3);width:100%;max-width:420px;padding:38px 28px;box-shadow:0 30px 80px rgba(0,0,0,.25),0 4px 20px rgba(0,0,0,.12)}
.lbox.shake{animation:shake .38s ease}
.ltabs{display:flex;border-radius:var(--r);overflow:hidden;border:1.5px solid var(--brd);padding:3px;background:var(--bg);margin-bottom:22px}
.ltab{flex:1;padding:10px;text-align:center;font-size:13px;font-weight:800;cursor:pointer;border-radius:10px;color:var(--t3);transition:all .15s}
.ltab.on{background:var(--card);color:var(--in-d);box-shadow:var(--sh)}
.err-box{background:var(--cr-l);border:1.5px solid rgba(230,57,70,.2);border-radius:var(--r);padding:11px 14px;font-size:13px;color:var(--cr);margin-bottom:14px;display:flex;align-items:center;gap:8px;font-weight:600}

/* METRIC CARDS */
.metric-grid{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:14px}
.metric-card{background:var(--card);border-radius:var(--r2);border:1px solid rgba(0,0,0,.05);padding:18px 14px;box-shadow:var(--sh);text-align:center;transition:transform .16s}
.metric-card:active{transform:scale(.97)}
.metric-val{font-size:32px;font-weight:900;font-family:var(--mf);background:linear-gradient(140deg,var(--in-d),var(--in-m));-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text;line-height:1}
.metric-lbl{font-size:10.5px;color:var(--t3);margin-top:5px;font-weight:700;letter-spacing:.04em;text-transform:uppercase}

/* QUICK TABS */
.ql-tabs{display:flex;border-radius:var(--r);overflow:hidden;border:1.5px solid var(--brd);background:var(--bg);padding:3px;margin-bottom:16px}
.ql-tab{flex:1;padding:10px;text-align:center;font-size:13px;font-weight:800;cursor:pointer;color:var(--t3);border-radius:10px;transition:all .15s}
.ql-tab.on{background:linear-gradient(135deg,var(--in),var(--in-d));color:#fff;box-shadow:0 4px 12px rgba(124,58,237,.25)}
.prod-card{display:flex;align-items:center;gap:13px;padding:14px 15px;border-radius:var(--r2);border:1.5px solid var(--brd);background:var(--card);cursor:pointer;margin-bottom:9px;transition:all .18s cubic-bezier(.34,1.56,.64,1);box-shadow:var(--sh)}
.prod-card:active{transform:scale(.97)}.prod-card.sel{border-color:var(--in-m);background:var(--in-l)}

/* SEND */
.step-dot{width:32px;height:32px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:13px;font-weight:900;flex-shrink:0;border:2px solid var(--brd);color:var(--t3);background:var(--card);transition:all .2s}
.step-dot.on{background:linear-gradient(135deg,var(--in),var(--in-d));border-color:var(--in);color:#fff;box-shadow:0 0 0 4px var(--in-t)}
.step-dot.done{background:var(--em);border-color:var(--em);color:#fff}
.step-line{flex:1;height:2px;background:var(--brd);margin:15px 8px 0;transition:background .2s}.step-line.done{background:var(--em)}
.ct-card{display:flex;align-items:center;gap:13px;padding:15px 17px;border-radius:var(--r2);border:1.5px solid var(--brd);background:var(--card);cursor:pointer;margin-bottom:11px;transition:all .16s;box-shadow:var(--sh)}
.ct-card:active{transform:scale(.98)}.ct-card.sel{border-color:var(--em);background:var(--em-l)}
.sic{display:flex;align-items:center;gap:11px;padding:12px 14px;border-radius:var(--r);border:1.5px solid var(--brd);background:var(--card);margin-bottom:9px;box-shadow:var(--sh)}
.cart-box{background:var(--em-l);border:1.5px solid rgba(0,184,122,.2);border-radius:var(--r);padding:14px 16px;margin:14px 0}

/* IMPORT / PHOTO */
.drop-z{border:2px dashed var(--brd2);border-radius:var(--r2);padding:44px 24px;text-align:center;cursor:pointer;transition:all .18s;background:var(--bg);position:relative;overflow:hidden}
.drop-z:hover{border-color:var(--in);background:var(--in-l)}.drop-z input{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%}
.prev-row{display:flex;align-items:center;gap:9px;padding:9px 12px;border-radius:10px;margin-bottom:5px;font-size:12px}
.prev-ok{background:var(--em-l);border:1px solid rgba(0,184,122,.2)}.prev-err{background:var(--cr-l);border:1px solid rgba(230,57,70,.2)}
.photo-zone{border:2px dashed var(--brd2);border-radius:var(--r);padding:20px;text-align:center;cursor:pointer;transition:all .15s;background:var(--bg);position:relative;overflow:hidden;margin-bottom:10px}
.photo-zone:hover{border-color:var(--in);background:var(--in-l)}.photo-zone input{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%}

/* WA BANNER */
.wa-banner{width:100%;display:flex;align-items:center;justify-content:space-between;padding:16px 18px;border-radius:var(--r2);background:linear-gradient(135deg,#25d366,#128c7e);border:none;cursor:pointer;margin-bottom:14px;box-shadow:0 4px 16px rgba(37,211,102,.2);transition:all .15s}
.wa-banner:active{transform:scale(.98)}

/* LOADING */
.loading-screen{display:flex;flex-direction:column;align-items:center;justify-content:center;height:100vh;background:linear-gradient(155deg,#450a0a 0%,#7f1d1d 50%,#dc2626 100%);gap:18px}
.loading-spinner{width:46px;height:46px;border:3px solid rgba(255,255,255,.20);border-top-color:rgba(255,255,255,.9);border-radius:50%;animation:spin .9s cubic-bezier(.5,0,.5,1) infinite}

/* EMPTY */
.empty{text-align:center;padding:40px 20px;color:var(--t3);font-size:13px;line-height:1.8;font-weight:500}

/* TOAST */
.toast-wrap{position:fixed;bottom:calc(var(--tab) + 14px);left:14px;right:14px;z-index:500;display:flex;flex-direction:column;gap:8px;pointer-events:none}
.toast-wrap > *{pointer-events:all;animation:toastIn .32s cubic-bezier(.16,1,.3,1)}

/* CARRUSEL */
.carousel-wrap{overflow-x:auto;scroll-snap-type:x mandatory;display:flex;gap:10px;padding:0 14px 4px;-webkit-overflow-scrolling:touch;scrollbar-width:none}
.carousel-wrap::-webkit-scrollbar{display:none}
.carousel-slide{flex-shrink:0;scroll-snap-align:start;border-radius:18px;overflow:hidden;cursor:pointer;box-shadow:0 4px 20px rgba(0,0,0,.12)}
.carousel-slide img{display:block;width:100%;height:100%;object-fit:cover}
.carousel-slide-placeholder{display:flex;flex-direction:column;align-items:center;justify-content:center;height:100%;font-weight:800;font-size:13px}
.carousel-dots{display:flex;justify-content:center;gap:5px;margin-top:6px}
.carousel-dot{width:5px;height:5px;border-radius:3px;background:var(--brd2);transition:all .28s cubic-bezier(.34,1.56,.64,1);cursor:pointer}
.carousel-dot.on{width:20px;background:var(--in-m)}
/* SECCIONES FAVORITAS */
.fav-grid{display:grid;grid-template-columns:1fr 1fr;gap:12px;padding:0 14px;margin-bottom:18px}
.fav-card{background:var(--bg2);border-radius:22px;border:none;padding:18px;display:flex;align-items:center;gap:13px;cursor:pointer;box-shadow:none;transition:transform .14s cubic-bezier(.34,1.56,.64,1)}
.fav-card:active{transform:scale(.96)}
.fav-card-ico{width:50px;height:50px;border-radius:16px;display:flex;align-items:center;justify-content:center;font-size:25px;flex-shrink:0;box-shadow:0 2px 8px rgba(0,0,0,.06)}
.fav-card-lbl{font-size:13px;font-weight:800;color:var(--t1);line-height:1.3}
.fav-card-sub{font-size:10.5px;color:var(--t3);font-weight:600;margin-top:1px}
/* MÉTRICAS SUPERIORES (estilo dashboard) */
.mtop-grid{display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;padding:0 14px;margin-bottom:18px}
.mtop-card{background:var(--card);border:1px solid var(--brd);border-radius:18px;padding:14px 12px;box-shadow:var(--sh);display:flex;flex-direction:column;gap:8px}
.mtop-ico{width:38px;height:38px;border-radius:12px;display:flex;align-items:center;justify-content:center;flex-shrink:0}
.mtop-lbl{font-size:10.5px;color:var(--t3);font-weight:700;line-height:1.25}
.mtop-val{font-size:20px;font-weight:900;color:var(--t1);letter-spacing:-.03em;line-height:1}
.mtop-foot{font-size:10px;font-weight:700;margin-top:1px}
/* CATEGORÍAS */
.cats-wrap{overflow-x:auto;display:flex;gap:10px;padding:0 14px 2px;scrollbar-width:none;-webkit-overflow-scrolling:touch}
.cats-wrap::-webkit-scrollbar{display:none}
.cat-item{flex-shrink:0;display:flex;flex-direction:column;align-items:center;gap:6px;cursor:pointer}
.cat-ico{width:60px;height:60px;border-radius:50%;background:var(--bg2);display:flex;align-items:center;justify-content:center;font-size:28px;transition:transform .12s}
.cat-item:active .cat-ico{transform:scale(.92)}
.cat-lbl{font-size:10px;font-weight:700;color:var(--t2);text-align:center;max-width:64px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
/* SECTION HEADERS */
.sec-hdr{display:flex;justify-content:space-between;align-items:center;padding:0 14px;margin-bottom:10px}
.sec-hdr-t{font-size:16px;font-weight:900;color:var(--t1)}
.sec-hdr-a{font-size:12px;font-weight:700;color:var(--in);cursor:pointer}
/* MISC */
.row{display:flex;align-items:center}.jb{justify-content:space-between}.g8{gap:8px}.g12{gap:12px}
::-webkit-scrollbar{width:4px;height:4px}::-webkit-scrollbar-thumb{background:var(--brd2);border-radius:4px}
`

// ─── SUBCOMPONENTES ───────────────────────────────────────────────────────────
const Avatar = memo(function Avatar(props) {
  var name=props.name, color=props.color, size=props.size||38, style=props.style||{};
  return (
    <div style={Object.assign({width:size,height:size,borderRadius:"50%",background:color||"#6366f1",display:"flex",alignItems:"center",justifyContent:"center",fontSize:Math.round(size*.36),fontWeight:800,color:"#fff",flexShrink:0,fontFamily:"var(--hf)"},style)}>
      {ini(name)}
    </div>
  );
});

const ProdThumb = memo(function ProdThumb(props) {
  var prod=props.prod, size=props.size||40, imgs=props.images;
  if (!prod) return <div style={{width:size,height:size,borderRadius:9,background:"var(--bg2)",flexShrink:0}}/>;
  var mainImg = imgs&&imgs.length>0 ? (imgs.find(function(i){return i.es_principal;})||imgs[0]) : null;
  var url = mainImg ? mainImg.url : prod.photo_url;
  if (url) return <img src={url} alt="" style={{width:size,height:size,borderRadius:9,objectFit:"cover",flexShrink:0,border:"1px solid var(--brd)"}}/>;
  return <div style={{width:size,height:size,borderRadius:9,background:"var(--bg2)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:Math.round(size*.5),flexShrink:0}}>{prod.emoji||"📦"}</div>;
});

// SearchBar with internal debounced state to avoid parent re-renders on keystroke
const SearchBar = memo(function SearchBar(props) {
  const [local, setLocal] = useState(props.value||"");
  useEffect(function(){ setLocal(props.value||""); }, [props.value]);
  useEffect(function(){
    var t = setTimeout(function(){ props.onChange(local); }, 200);
    return function(){ clearTimeout(t); };
  }, [local]);
  return (
    <div style={{position:"relative",marginBottom:12}}>
      <span style={{position:"absolute",left:12,top:"50%",transform:"translateY(-50%)",color:"var(--t3)",pointerEvents:"none"}}><Ic n="search" s={16}/></span>
      <input style={{width:"100%",padding:"11px 38px",borderRadius:10,border:"1.5px solid var(--brd)",background:"var(--card)",color:"var(--t1)",fontFamily:"var(--hf)",fontSize:14,outline:"none"}} placeholder={props.placeholder||"Buscar..."} value={local} onChange={function(e){setLocal(e.target.value);}}/>
      {local?<button onClick={function(){setLocal("");props.onChange("");}} style={{position:"absolute",right:10,top:"50%",transform:"translateY(-50%)",background:"var(--bg2)",border:"none",borderRadius:6,padding:4,cursor:"pointer",color:"var(--t3)",display:"flex"}}><Ic n="x" s={14}/></button>:null}
    </div>
  );
});

// Header search: pill style + debounced (no re-render del padre por tecla)
const HeaderSearch = memo(function HeaderSearch(props) {
  const [local, setLocal] = useState(props.value||"");
  useEffect(function(){ setLocal(props.value||""); }, [props.value]);
  useEffect(function(){
    var t = setTimeout(function(){ props.onChange(local); }, 220);
    return function(){ clearTimeout(t); };
  }, [local]);
  return (
    <div className="hdr-search">
      <span className="hdr-search-ico"><Ic n="search" s={18}/></span>
      <input placeholder={props.placeholder||"Buscar..."} value={local} onChange={function(e){setLocal(e.target.value);}}/>
    </div>
  );
});

function Toast(props) {
  var t=props.t;
  var cfg={
    s:{bg:"linear-gradient(135deg,#059669,#047857)",ico:"✅"},
    e:{bg:"linear-gradient(135deg,#dc2626,#b91c1c)",ico:"❌"},
    i:{bg:"linear-gradient(135deg,#7c3aed,#6d28d9)",ico:"ℹ️"}
  }[t.type]||{bg:"linear-gradient(135deg,#059669,#047857)",ico:"✅"};
  return (
    <div style={{background:cfg.bg,borderRadius:16,padding:"14px 16px",display:"flex",alignItems:"flex-start",gap:10,boxShadow:"0 8px 24px rgba(0,0,0,.2)"}}>
      <div style={{fontSize:18,flexShrink:0,marginTop:1}}>{cfg.ico}</div>
      <div style={{flex:1}}>
        <div style={{fontWeight:800,fontSize:13,color:"#fff"}}>{t.title}</div>
        {t.body?<div style={{fontSize:12,color:"rgba(255,255,255,.8)",marginTop:3}}>{t.body}</div>:null}
      </div>
      <button onClick={function(){props.remove(t.id);}} style={{background:"rgba(255,255,255,.2)",border:"none",cursor:"pointer",color:"#fff",borderRadius:8,padding:"3px 6px",fontSize:11,fontWeight:700}}>✕</button>
    </div>
  );
}

const QtyControl = memo(function QtyControl(props) {
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
});

// ─── APP ──────────────────────────────────────────────────────────────────────
// ══════════ TIENDA PÚBLICA (sin login) ══════════
function TiendaPublica({ tiendaId }) {
  const [info, setInfo]       = useState(null);
  const [prods, setProds]     = useState([]);
  const [vista, setVista]     = useState("stock"); // stock | catalogo
  const [busq, setBusq]       = useState("");
  const [sel, setSel]         = useState(null);    // producto seleccionado
  const [qty, setQty]         = useState(1);
  const [nom, setNom]         = useState("");
  const [tel, setTel]         = useState("");
  const [nota, setNota]       = useState("");
  const [enviando, setEnviando] = useState(false);
  const [exito, setExito]     = useState(false);
  const [err, setErr]         = useState("");

  useEffect(function() {
    Promise.all([
      sb.rpc("rpc_tienda_info", { p_user_id: tiendaId }),
      sb.rpc("rpc_tienda_publica", { p_user_id: tiendaId })
    ]).then(function(rs) {
      if (rs[0].data && rs[0].data.length) setInfo(rs[0].data[0]);
      if (rs[1].data) setProds(rs[1].data);
    }).catch(function(){ setErr("No se pudo cargar la tienda"); });
  }, [tiendaId]);

  async function enviarPedido() {
    if (!nom.trim() || nom.trim().length < 2) { setErr("Poné tu nombre"); return; }
    if (!sel) return;
    setEnviando(true); setErr("");
    try {
      var r = await sb.rpc("rpc_pedido_publico", {
        p_user_id: tiendaId, p_nombre: nom.trim(), p_telefono: tel.trim(),
        p_product_id: sel.product_id, p_qty: qty, p_nota: nota.trim()
      });
      if (r.error) { setErr(r.error.message); setEnviando(false); return; }
      setExito(true);
    } catch(e) { setErr("Error de conexión. Reintentá."); }
    setEnviando(false);
  }

  var col = (info && info.color) || "#e0224e";
  var q = busq.toLowerCase().trim();
  var lista = prods.filter(function(p) {
    if (vista === "stock" && p.qty_disponible <= 0) return false;
    if (q && !p.name.toLowerCase().includes(q) && !(p.sku||"").toLowerCase().includes(q)) return false;
    return true;
  });

  if (exito) return (
    <div style={{minHeight:"100vh",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",background:"#fdf6f7",fontFamily:"system-ui",padding:20,textAlign:"center"}}>
      <div style={{fontSize:60,marginBottom:16}}>🎉</div>
      <div style={{fontSize:22,fontWeight:800,color:col,marginBottom:8}}>¡Pedido enviado!</div>
      <div style={{fontSize:14,color:"#777",maxWidth:300}}>{info?info.name:"La vendedora"} recibió tu pedido de <b>{sel?sel.name:""}</b> y se va a contactar con vos para coordinar la entrega.</div>
      <button onClick={function(){setExito(false);setSel(null);setQty(1);setNota("");}} style={{marginTop:24,background:col,color:"#fff",border:"none",borderRadius:12,padding:"12px 28px",fontSize:14,fontWeight:700,cursor:"pointer"}}>Hacer otro pedido</button>
    </div>
  );

  return (
    <div style={{minHeight:"100vh",background:"#fdf6f7",fontFamily:"system-ui",paddingBottom:40}}>
      {/* Header */}
      <div style={{background:col,padding:"26px 18px 20px",color:"#fff"}}>
        <div style={{fontSize:12,opacity:0.85,fontWeight:600,letterSpacing:1}}>TIENDA DE</div>
        <div style={{fontSize:24,fontWeight:800}}>{info?info.name:"..."}</div>
      </div>

      {/* Tabs */}
      <div style={{display:"flex",gap:8,padding:"14px 16px 6px"}}>
        <button onClick={function(){setVista("stock");}} style={{flex:1,padding:"10px",borderRadius:12,border:"none",fontWeight:700,fontSize:13,cursor:"pointer",background:vista==="stock"?col:"#fff",color:vista==="stock"?"#fff":"#999",boxShadow:"0 1px 4px rgba(0,0,0,0.08)"}}>📦 Disponible ya</button>
        <button onClick={function(){setVista("catalogo");}} style={{flex:1,padding:"10px",borderRadius:12,border:"none",fontWeight:700,fontSize:13,cursor:"pointer",background:vista==="catalogo"?col:"#fff",color:vista==="catalogo"?"#fff":"#999",boxShadow:"0 1px 4px rgba(0,0,0,0.08)"}}>📖 Catálogo completo</button>
      </div>

      {/* Buscador */}
      <div style={{padding:"8px 16px"}}>
        <input value={busq} onChange={function(e){setBusq(e.target.value);}} placeholder="Buscar producto..."
          style={{width:"100%",boxSizing:"border-box",border:"1.5px solid #eee",borderRadius:12,padding:"11px 14px",fontSize:14,outline:"none"}}/>
      </div>

      {/* Lista */}
      <div style={{padding:"4px 16px"}}>
        {lista.length===0&&<div style={{textAlign:"center",color:"#aaa",fontSize:13,padding:"30px 0"}}>{vista==="stock"?"Sin stock disponible en este momento":"Sin resultados"}</div>}
        {lista.map(function(p) {
          return (
            <div key={p.product_id} onClick={function(){setSel(p);setQty(1);setErr("");}}
              style={{display:"flex",alignItems:"center",gap:12,background:"#fff",borderRadius:14,padding:"12px 14px",marginBottom:8,boxShadow:"0 1px 4px rgba(0,0,0,0.06)",cursor:"pointer",border:sel&&sel.product_id===p.product_id?"2px solid "+col:"2px solid transparent"}}>
              {p.photo_url
                ? <img src={p.photo_url} style={{width:46,height:46,borderRadius:10,objectFit:"cover"}}/>
                : <div style={{width:46,height:46,borderRadius:10,background:"#fdf0f2",display:"flex",alignItems:"center",justifyContent:"center",fontSize:22}}>{p.emoji||"📦"}</div>}
              <div style={{flex:1,minWidth:0}}>
                <div style={{fontSize:14,fontWeight:700,color:"#333"}}>{p.name}</div>
                <div style={{fontSize:11,color:"#999"}}>{p.sku}{p.qty_disponible>0?" · "+p.qty_disponible+" disponibles":" · a pedido"}</div>
              </div>
              <div style={{fontWeight:800,fontSize:15,color:col,whiteSpace:"nowrap"}}>${parseFloat(p.price||0).toLocaleString("es-AR")}</div>
            </div>
          );
        })}
      </div>

      {/* Formulario de pedido */}
      {sel&&(
        <div style={{position:"fixed",bottom:0,left:0,right:0,background:"#fff",borderRadius:"22px 22px 0 0",boxShadow:"0 -4px 24px rgba(0,0,0,0.15)",padding:"18px 18px 22px",zIndex:50}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
            <div style={{fontSize:15,fontWeight:800,color:"#333"}}>{sel.name}</div>
            <button onClick={function(){setSel(null);}} style={{background:"none",border:"none",fontSize:18,cursor:"pointer",color:"#bbb"}}>✕</button>
          </div>
          <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:12}}>
            <span style={{fontSize:13,color:"#777"}}>Cantidad</span>
            <button onClick={function(){setQty(Math.max(1,qty-1));}} style={{width:32,height:32,borderRadius:9,border:"1.5px solid #eee",background:"#fff",fontSize:16,cursor:"pointer"}}>−</button>
            <span style={{fontWeight:800,fontSize:16,minWidth:24,textAlign:"center"}}>{qty}</span>
            <button onClick={function(){setQty(Math.min(50,qty+1));}} style={{width:32,height:32,borderRadius:9,border:"1.5px solid #eee",background:"#fff",fontSize:16,cursor:"pointer"}}>+</button>
            <div style={{marginLeft:"auto",fontWeight:800,fontSize:17,color:col}}>${(parseFloat(sel.price||0)*qty).toLocaleString("es-AR")}</div>
          </div>
          <input value={nom} onChange={function(e){setNom(e.target.value);}} placeholder="Tu nombre *"
            style={{width:"100%",boxSizing:"border-box",border:"1.5px solid #eee",borderRadius:11,padding:"11px 13px",fontSize:14,outline:"none",marginBottom:8}}/>
          <input value={tel} onChange={function(e){setTel(e.target.value);}} placeholder="Tu teléfono (opcional)"
            style={{width:"100%",boxSizing:"border-box",border:"1.5px solid #eee",borderRadius:11,padding:"11px 13px",fontSize:14,outline:"none",marginBottom:8}}/>
          <input value={nota} onChange={function(e){setNota(e.target.value);}} placeholder="Nota: talle, color, etc. (opcional)"
            style={{width:"100%",boxSizing:"border-box",border:"1.5px solid #eee",borderRadius:11,padding:"11px 13px",fontSize:14,outline:"none",marginBottom:10}}/>
          {err&&<div style={{color:"#d32",fontSize:12,fontWeight:700,marginBottom:8}}>{err}</div>}
          <button onClick={enviarPedido} disabled={enviando}
            style={{width:"100%",background:col,color:"#fff",border:"none",borderRadius:13,padding:"14px",fontSize:15,fontWeight:800,cursor:"pointer",opacity:enviando?0.6:1}}>
            {enviando?"Enviando...":"✅ Confirmar pedido"}
          </button>
        </div>
      )}
    </div>
  );
}

export default function App() {
  // Modo tienda pública: ?tienda=USER_ID — sin login
  var tiendaParam = new URLSearchParams(window.location.search).get("tienda");
  if (tiendaParam) return <TiendaPublica tiendaId={tiendaParam}/>;

  // ── SESSION ─────────────────────────────────────────────────────────────────
  const [loading,   setLoading]   = useState(true);
  const [me,        setMe]        = useState(null);   // users row
  const [authMode,  setAuthMode]  = useState("login");
  const [aEmail,    setAEmail]    = useState("");
  const [aPass,     setAPass]     = useState("");
  const [aName,     setAName]     = useState("");
  const [aRole,     setARole]     = useState("Revendedora Exclusiva");
  const [aPass2,    setAPass2]    = useState("");
  const [showPass,  setShowPass]  = useState(false);
  const [authErr,   setAuthErr]   = useState("");
  const [loggingIn, setLoggingIn] = useState(false);
  const [authOk,    setAuthOk]    = useState(false);  // login exitoso, cargando perfil

  // Seguridad: si el perfil no carga en 10s, volver al login con mensaje
  useEffect(function() {
    if (!authOk || me) return;
    var t = setTimeout(function() {
      setAuthOk(false);
      setAuthErr("No se pudo cargar tu cuenta. Intentá de nuevo.");
    }, 10000);
    return function(){ clearTimeout(t); };
  }, [authOk, me]);
  const [pSearch,   setPSearch]   = useState("");     // buscador rápido de precios
  const [authShake, setAuthShake] = useState(false);

  // ── DATA ────────────────────────────────────────────────────────────────────
  const [products,   setProducts]   = useState([]);
  const [inventory,  setInventory]  = useState([]); // my rows from inventory table
  const [contacts,   setContacts]   = useState([]); // user objects
  const [allUsers,   setAllUsers]   = useState([]); // for superadmin
  const [transfers,  setTransfers]  = useState([]);
  const [notifs,     setNotifs]     = useState([]);
  const [logs,       setLogs]       = useState([]);
  const [adminStats, setAdminStats] = useState(null);

  // ── UI ──────────────────────────────────────────────────────────────────────
  const [tab,       setTab]       = useState("inicio");
  const [masMenu,   setMasMenu]   = useState(false);

  // ── FASE 1: Roles y jerarquía ──────────────────────────────────────────
  const [misDatos,      setMisDatos]      = useState(null); // {telefono, direccion, avatar_url} de "me" (recargado aparte por si acaso)
  const [ctaName,       setCtaName]       = useState("");
  const [ctaTel,        setCtaTel]        = useState("");
  const [ctaDir,        setCtaDir]        = useState("");
  const [ctaDni,        setCtaDni]        = useState("");
  const [ctaCodVend,    setCtaCodVend]    = useState("");
  const [ctaLocalidad,  setCtaLocalidad]  = useState("");
  const [ctaPass,       setCtaPass]       = useState("");
  const [ctaPass2,      setCtaPass2]      = useState("");
  const [ctaSaving,     setCtaSaving]     = useState(false);
  const [ctaJerInfo,    setCtaJerInfo]    = useState(null); // {lider:{name}, empresa:{name}, comision}
  const [miEquipo,      setMiEquipo]      = useState([]);   // usuarios bajo mi jerarquía
  const [misInvites,    setMisInvites]    = useState([]);   // invitaciones pendientes creadas por mí
  const [invName,       setInvName]       = useState("");
  const [invEmail,      setInvEmail]      = useState("");
  const [invRole,       setInvRole]       = useState("reseller");
  const [invLiderId,    setInvLiderId]    = useState("");
  const [invComis,      setInvComis]      = useState("10");
  const [invSaving,     setInvSaving]     = useState(false);
  const [equipoEditComis, setEquipoEditComis] = useState({}); // {userId: valor en edición}

  // ── FASE 2: Pedidos en cascada (productos sin stock) ───────────────────
  const [pedEspList,   setPedEspList]   = useState([]);
  const [pedEspLoading,setPedEspLoading]= useState(false);
  const [peShowForm,   setPeShowForm]   = useState(false);
  const [peProdSrch,   setPeProdSrch]   = useState("");
  const [peProdId,     setPeProdId]     = useState("");
  const [peQty,        setPeQty]        = useState(1);
  const [peColor,      setPeColor]      = useState("");
  const [peTalle,      setPeTalle]      = useState("");
  const [peCarrito,    setPeCarrito]    = useState([]); // [{product_id,name,sku,emoji,photo_url,price,qty,color,talle}]
  const [peCliNom,     setPeCliNom]     = useState("");
  const [peCliTel,     setPeCliTel]     = useState("");
  const [peNota,       setPeNota]       = useState("");
  const [peFiltroEmp,  setPeFiltroEmp]  = useState("");
  const [peFiltroLid,  setPeFiltroLid]  = useState("");
  const [peFiltroVend, setPeFiltroVend] = useState("");
  const [peFiltroCamp, setPeFiltroCamp] = useState("");
  const [peFiltroEst,  setPeFiltroEst]  = useState("");

  // ── FASE 5: Campañas ─────────────────────────────────────────────────────
  const [campanias,       setCampanias]       = useState([]);
  const [campNombre,      setCampNombre]      = useState("");
  const [campSaving,      setCampSaving]      = useState(false);
  const [peCampaniaId,    setPeCampaniaId]    = useState("");
  const [peSaving,     setPeSaving]     = useState(false);
  const [peObserv,     setPeObserv]     = useState({}); // {pedidoId: texto observación}
  const [peBusy,       setPeBusy]       = useState({}); // {pedidoId: true mientras procesa acción}
  const [peHistOpen,   setPeHistOpen]   = useState(null); // id del pedido con historial abierto
  const [peHist,       setPeHist]       = useState([]);

  // ── FASE 3: Resúmenes (pedidos + ventas + comisiones) ───────────────────
  const [resumen,       setResumen]       = useState(null);
  const [resumenLoading,setResumenLoading]= useState(false);

  useEffect(function() {
    if (!me) return;
    setCtaName(me.name||""); setCtaTel(me.telefono||""); setCtaDir(me.direccion||"");
    setCtaDni(me.dni||""); setCtaCodVend(me.codigo_vendedora||""); setCtaLocalidad(me.localidad||"");
    loadJerarquia();
    loadPedidosEspeciales();
    loadResumen();
    loadCampanias();
  }, [me && me.id]);

  // Re-sincronizar mi propio perfil: al volver a la app, al navegar a
  // secciones sensibles al rol, y cada 30s de respaldo — así si un
  // superadmin/empresaria me cambia el rol, se refleja solo sin relogin.
  async function resyncPerfil() {
    if (!me) return;
    try {
      var res = await sb.from("users").select("*").eq("id", me.id).maybeSingle();
      if (!res.data) return;
      var d = res.data;
      if (d.role !== me.role || d.empresa_id !== me.empresa_id || d.lider_id !== me.lider_id || d.comision_lider_pct !== me.comision_lider_pct) {
        setMe(d);
        toast("Tu cuenta fue actualizada", "Tu rol o equipo cambió", "i");
      }
    } catch(e) { /* noop */ }
  }

  useEffect(function() {
    if (!me) return;
    function onFocusResync() { if (document.visibilityState === "visible") resyncPerfil(); }
    document.addEventListener("visibilitychange", onFocusResync);
    window.addEventListener("focus", onFocusResync);
    var poll = setInterval(resyncPerfil, 180000);
    return function() {
      document.removeEventListener("visibilitychange", onFocusResync);
      window.removeEventListener("focus", onFocusResync);
      clearInterval(poll);
    };
  }, [me && me.id, me && me.role, me && me.empresa_id, me && me.lider_id, me && me.comision_lider_pct]);
  const [toasts,    setToasts]    = useState([]);
  const [srchStock, setSrchStock] = useState("");
  const [srchCat,   setSrchCat]   = useState("");
  const [srchCon,   setSrchCon]   = useState("");
  const [srchLog,   setSrchLog]   = useState("");
  // Filtro de mes en la pestaña Ventas (offset: 0 = mes actual, -1 = mes anterior, etc.)
  const [ventasMesOffset, setVentasMesOffset] = useState(0);
  const [ventasVerTodo,   setVentasVerTodo]   = useState(false);
  // ── PEDIDOS ──────────────────────────────────────────────────────────────
  const [pedidos,    setPedidos]   = useState([]);
  const [pedNombre,  setPedNombre] = useState("");
  const [showPedForm, setShowPedForm] = useState(false);
  const [pedWA,      setPedWA]     = useState("");
  const [pedProdId,  setPedProdId] = useState("");
  const [pedQty,     setPedQty]    = useState(1);
  const [pedNota,    setPedNota]   = useState("");
  const [pedSena,    setPedSena]   = useState("");
  const [pedSrch,    setPedSrch]   = useState("");
  const [pedPSrch,   setPedPSrch]  = useState("");
  const [pedFilter,  setPedFilter] = useState("todos"); // todos|pendiente|entregado
  const [pedEdit,    setPedEdit]   = useState(null);
  const [ganancia,  setGanancia]  = useState(30);
  const [editGan,   setEditGan]   = useState(false);
  const [ganInput,  setGanInput]  = useState("30");
  const [srchPrice, setSrchPrice] = useState("");
  const [priceMin,  setPriceMin]  = useState("");
  const [priceMax,  setPriceMax]  = useState("");

  // ── CATALOG ABM ─────────────────────────────────────────────────────────────
  const [editP,    setEditP]    = useState(null);
  const [fSku,     setFSku]     = useState("");
  const [fName,    setFName]    = useState("");
  // Carrusel de imágenes por producto
  const [prodImages,   setProdImages]   = useState({});  // { product_id: [img,...] }
  const [imgUploading, setImgUploading] = useState(false);
  const [imgEditPid,   setImgEditPid]   = useState(null);
  // Carga rápida inline en card de stock
  const [quickLoadId,  setQuickLoadId]  = useState(null);  // inventory item.id abierto
  const [quickLoadQty, setQuickLoadQty] = useState(1);
  const [fPrice,   setFPrice]   = useState("");
  const [fCat,     setFCat]     = useState("General");
  const [fEmoji,   setFEmoji]   = useState("✨");
  const [fPhoto,   setFPhoto]   = useState(null);
  const [fStock,   setFStock]   = useState("0");
  const [delConf,  setDelConf]  = useState(null);
  const [showInactive, setShowInactive] = useState(false);
  const [bulkSel, setBulkSel] = useState({});  // {prodId: true}

  // ── QUICK LOAD ──────────────────────────────────────────────────────────────
  const [qlMode,      setQlMode]      = useState("add");
  const [qlStockType, setQlStockType] = useState("PROPIO");       // 'PROPIO' | 'CONSIGNACION'
  const [qlSupplierId,setQlSupplierId]= useState("");              // contact id when CONSIGNACION
  const [qlPid,    setQlPid]    = useState("");
  const [qlQty,    setQlQty]    = useState(1);
  const [qlSku,    setQlSku]    = useState("");
  const [qlName,   setQlName]   = useState("");
  const [qlPrice,  setQlPrice]  = useState("");
  const [qlCat,    setQlCat]    = useState("General");
  const [qlEmoji,  setQlEmoji]  = useState("✨");
  const [qlPhoto,  setQlPhoto]  = useState(null);
  const [qlSrch,   setQlSrch]   = useState("");

  // ── SEND ────────────────────────────────────────────────────────────────────
  const [sendStep, setSendStep] = useState(1);
  const [sendTo,   setSendTo]   = useState("");
  const [sendCart, setSendCart] = useState({});
  const [sendSrch, setSendSrch] = useState("");

  // ── IMPORT ──────────────────────────────────────────────────────────────────
  const [impTab,  setImpTab]  = useState("file");
  const [impTxt,  setImpTxt]  = useState("");
  const [impRows, setImpRows] = useState([]);
  const [impQty,  setImpQty]  = useState(0);
  const [impFile, setImpFile] = useState(null);

  // ── CONSIGNACIONES (conteo para badge) ──────────────────────────────────────
  const [consignActivas,  setConsignActivas]  = useState([]);
  const [cancelConfirm,   setCancelConfirm]   = useState(null); // tx.id esperando confirmación de cancelación

  // ── CAROUSEL ─────────────────────────────────────────────────────────────────
  const [carousels,    setCarousels]    = useState([]);  // rows from offer_carousels table
  const [offerSlide,   setOfferSlide]   = useState(0);   // active offer carousel index

  // Rotación automática del carrusel cada 4 segundos
  useEffect(function() {
    var activos = carousels.filter(function(c){ return c.active; });
    var total = activos.length > 0 ? activos.length : 1;
    if (total <= 1) return;
    var timer = setInterval(function() {
      setOfferSlide(function(prev) {
        var next = (prev + 1) % total;
        var el = document.getElementById("main-carousel");
        if (el) {
          var w = el.scrollWidth / total;
          el.scrollTo({ left: w * next, behavior: "smooth" });
        }
        return next;
      });
    }, 4000);
    return function() { clearInterval(timer); };
  }, [carousels]);
  const [carouselEdit, setCarouselEdit] = useState(false);
  const [cTitle,       setCTitle]       = useState("");
  const [cSubtitle,    setCSubtitle]    = useState("");
  const [cBg,          setCBg]          = useState("#e0224e");
  const [cEmoji,       setCEmoji]       = useState("🔥");
  const [cLink,        setCLink]        = useState("");
  const [cEditId,      setCEditId]      = useState(null);
  const [cImg,         setCImg]         = useState("");      // URL de imagen del slide
  const [cImgUp,       setCImgUp]       = useState(false);
  const [cDelConf,     setCDelConf]     = useState(null);
  const [cSaving,      setCSaving]      = useState(false);

  // ── CONTACTS ────────────────────────────────────────────────────────────────
  const [ctQ,     setCtQ]     = useState("");

  // ── MODALS ──────────────────────────────────────────────────────────────────
  const [txModal,  setTxModal]  = useState(null);
  const [txQty,    setTxQty]    = useState(1);
  const [txTo,     setTxTo]     = useState("");
  const [shareM,   setShareM]   = useState(false);
  const [shareSel,  setShareSel]  = useState({});
  const [editStock,  setEditStock]  = useState(null);
  const [editQty,    setEditQty]    = useState(0);
  const [movModal,   setMovModal]   = useState(null);  // invRow for movement
  const [movType,    setMovType]    = useState("entrada"); // "entrada"|"salida"|"ajuste"
  const [movQty,     setMovQty]     = useState(1);
  const [movNote,    setMovNote]    = useState("");
  const [movHistory,   setMovHistory]   = useState([]);
  // Deudas de consignación: {id, supplier_id, supplier_name, product_name, qty, amount, date}
  const [consignDebts, setConsignDebts] = useState([]);
  const [retModal,    setRetModal]    = useState(null);
  const [retQty,      setRetQty]      = useState(1);
  const [consignSrch, setConsignSrch] = useState("");
  const [consignTab,     setConsignTab]     = useState("enviado");
  const [consignView,    setConsignView]    = useState("list");
  const [selRevend,      setSelRevend]      = useState(null);
  const [prodSearch,     setProdSearch]     = useState("");
  const [recvInv,        setRecvInv]        = useState([]);   // recipient inventory rows

  // ── HELPERS ──────────────────────────────────────────────────────────────────
  function toast(title, body, type) {
    var id = Date.now() + Math.random();
    setToasts(function(p){ return [...p,{id:id,title:title,body:body||"",type:type||"s"}]; });
    setTimeout(function(){ setToasts(function(p){ return p.filter(function(t){ return t.id!==id; }); }); }, 3800);
  }
  function rmToast(id){ setToasts(function(p){ return p.filter(function(t){ return t.id!==id; }); }); }
  function shake(){ setAuthShake(true); setTimeout(function(){ setAuthShake(false); },450); }

  var isAdmin = me && me.role === "superadmin";
  var myStock = useMemo(function(){ return inventory.filter(function(i){ return i.qty_available>0; }); }, [inventory]);
  var unreadNotifs = useMemo(function(){ return notifs.filter(function(n){ return !n.read; }).length; }, [notifs]);
  var pendingTransfers = useMemo(function(){ return transfers.filter(function(t){ return t.to_user_id===me?.id && t.status==="pending"; }).length; }, [transfers, me]);
  var totalBadge = unreadNotifs + pendingTransfers;

  // ── DEVOLVER CONSIGNA (receptor devuelve al emisor) ─────────────────────────
  async function doReturnConsigna() {
    var tx = retModal;
    var prod = tx.product || products.find(function(p){ return p.id===tx.product_id; });
    if (retQty<=0){ toast("Cantidad inválida","","e"); return; }

    // Quien devuelve: puede ser el receptor (me.id===tx.to_user_id)
    // o el admin forzando desde enviado (me.id===tx.from_user_id)
    var receiverId = tx.to_user_id;
    var senderId   = tx.from_user_id;

    // 1. Restar del inventario del receptor
    var rcvInv = await sb.from("inventory").select("*").eq("user_id", receiverId).eq("product_id", tx.product_id).single();
    if (rcvInv.data && rcvInv.data.qty_available >= retQty) {
      await sb.from("inventory").update({qty_available: rcvInv.data.qty_available - retQty}).eq("id", rcvInv.data.id);
    } else {
      toast("El receptor no tiene suficiente stock disponible","","e"); return;
    }

    // 2. Restituir al emisor — query DB directly
    var srcInvDB = await sb.from("inventory").select("*").eq("user_id", senderId).eq("product_id", tx.product_id).single();
    if (srcInvDB.data) {
      await sb.from("inventory").update({qty_available: srcInvDB.data.qty_available + retQty}).eq("id", srcInvDB.data.id);
    } else {
      await sb.from("inventory").insert({user_id: senderId, product_id: tx.product_id, qty_available: retQty, qty_sold:0});
    }

    // 3. Actualizar cantidad del transfer
    var newQty = tx.qty - retQty;
    if (newQty <= 0) {
      await sb.from("transfers").update({status:"rechazado"}).eq("id", tx.id);
    } else {
      await sb.from("transfers").update({qty: newQty}).eq("id", tx.id);
    }

    // 4. Notificar al emisor si quien devuelve es el receptor
    if (me.id === receiverId) {
      await sb.from("notifications").insert({
        to_user_id: senderId,
        from_name: me.name,
        type: "confirm",
        message: "↩️ "+me.name+" devolvió "+retQty+"x "+(prod?prod.name:"producto")+". Stock restituido a tu inventario."
      });
    }

    toast("Devolución registrada", retQty+"x "+(prod?prod.name:"")+" restituido", "s");
    setRetModal(null);
    await loadData(me.id, me.role);
    await loadRecvInventory();
  }

  // ── LOAD RECIPIENT INVENTORY FOR CONSIGNA ─────────────────────────────────────
  async function loadRecvInventory() {
    var sentTxs = transfers.filter(function(t){ return t.from_user_id===me.id&&t.status==="confirmed"&&t.qty>0; });
    if (!sentTxs.length) { setRecvInv([]); return; }
    // Get unique recipient IDs
    var uids = [...new Set(sentTxs.map(function(t){return t.to_user_id;}))];
    var allRecvInv = [];
    for (var i=0;i<uids.length;i++) {
      var r = await sb.from("inventory").select("*").eq("user_id", uids[i]);
      if (r.data) allRecvInv = allRecvInv.concat(r.data);
    }
    setRecvInv(allRecvInv);
  }

  // ── DELETE USER (admin only) ─────────────────────────────────────────────────
  const [delUserConf, setDelUserConf] = useState(null);

  // ── IMPORTAR PRECIOS ──────────────────────────────────────────────────────
  const [pxFile,    setPxFile]    = useState(null);   // nombre del archivo
  const [pxRows,    setPxRows]    = useState([]);     // [{sku, name, oldPrice, newPrice, ok, err}]
  const [pxSaving,  setPxSaving]  = useState(false);
  const [pxDone,    setPxDone]    = useState(false);
  const [resetConf, setResetConf] = useState(false);
  const [resetBusy, setResetBusy] = useState(false);

  async function doDeleteUser(userId, userName) {
    var r = await sb.rpc("rpc_eliminar_usuario", { p_user_id: userId });
    if (r.error) { toast("Error", ""+r.error.message, "e"); return; }
    setAllUsers(function(p){ return p.filter(function(u){ return u.id!==userId; }); });
    toast("Usuario eliminado", userName, "i");
    setDelUserConf(null);
    await loadData(me.id, me.role);
  }


  // ── IMPORTAR PRECIOS POR EXCEL ────────────────────────────────────────────
  function pxHandleFile(file) {
    if (!file) return;
    setPxFile(file.name); setPxRows([]); setPxDone(false);
    var reader = new FileReader();
    reader.onload = async function(ev) {
      var raw;
      try {
        var wb = XLSX.read(ev.target.result, {type:"array"});
        var ws = wb.Sheets[wb.SheetNames[0]];
        raw = XLSX.utils.sheet_to_json(ws, {header:1, defval:""});
      } catch(e) { toast("Error al leer el archivo","","e"); return; }
      var hdr = raw[0] ? raw[0].map(function(c){ return String(c).toLowerCase().trim(); }) : [];
      var iSku   = hdr.findIndex(function(h){ return h.includes("sku") || h.includes("digo"); });
      var iName  = hdr.findIndex(function(h){ return h.includes("nombre") || h.includes("name") || h.includes("producto"); });
      var iPrice = hdr.findIndex(function(h){ return h.includes("precio") || h.includes("price"); });
      if (iSku < 0 && iName < 0) { toast("No se encontró columna SKU o Nombre","Revisá el encabezado","e"); return; }
      if (iPrice < 0) { toast("No se encontró columna Precio","El Excel debe tener Precio","e"); return; }
      var { data: prods } = await sb.from("products").select("id,sku,name,price");
      var prodBySku = {};
      var prodByName = {};
      (prods||[]).forEach(function(p){
        if (p.sku) prodBySku[String(p.sku).toLowerCase().trim()] = p;
        prodByName[p.name.toLowerCase().trim()] = p;
      });
      var rows = [];
      for (var i = 1; i < raw.length; i++) {
        var r = raw[i];
        if (!r || r.every(function(c){ return !c; })) continue;
        var sku   = iSku >= 0  ? String(r[iSku]||"").trim()  : "";
        var name  = iName >= 0 ? String(r[iName]||"").trim() : "";
        var priceRaw = String(r[iPrice]||"").replace(",",".").replace(/[^0-9.]/g,"");
        var newPrice = parseFloat(priceRaw);
        var match = (sku ? prodBySku[sku.toLowerCase()] : null) || (name ? prodByName[name.toLowerCase()] : null);
        rows.push({
          sku, name: name || (match ? match.name : ""),
          oldPrice: match ? match.price : null,
          newPrice: isNaN(newPrice) ? null : newPrice,
          prodId:   match ? match.id : null,
          ok:       !!match && !isNaN(newPrice) && newPrice > 0,
          err:      !match ? "Sin coincidencia" : (isNaN(newPrice)||newPrice<=0 ? "Precio invalido" : "")
        });
      }
      setPxRows(rows);
      var validos = rows.filter(function(r){ return r.ok; }).length;
      toast(validos + " precios listos", rows.length - validos + " sin coincidencia", validos > 0 ? "s" : "e");
    };
    reader.readAsArrayBuffer(file);
  }

  async function doPxImport() {
    var validos = pxRows.filter(function(r){ return r.ok; });
    if (!validos.length) return;
    setPxSaving(true);
    var errores = 0;
    for (var i = 0; i < validos.length; i++) {
      var r = validos[i];
      var res = await sb.from("products").update({price: r.newPrice}).eq("id", r.prodId);
      if (res.error) errores++;
    }
    setPxSaving(false);
    if (errores === 0) {
      toast("Precios actualizados", validos.length + " productos", "s");
      setPxDone(true);
    } else {
      toast("Actualizados " + (validos.length - errores) + "/" + validos.length, errores + " errores", "e");
    }
  }

  async function doResetData() {
    setResetBusy(true);
    try {
      var NEQ = "00000000-0000-0000-0000-000000000000";
      await sb.from("consignacion_deudas").delete().neq("id", NEQ);
      await sb.from("consignacion_items").delete().neq("id", NEQ);
      await sb.from("consignaciones").delete().neq("id", NEQ);
      await sb.from("sale_logs").delete().neq("id", NEQ);
      await sb.from("transfers").delete().neq("id", NEQ);
      await sb.from("stock_movements").delete().neq("id", NEQ);
      await sb.from("pedidos").delete().neq("id", NEQ);
      await sb.from("inventory").update({qty_available:0, qty_sold:0, stock_recibido:0}).neq("id", NEQ);
      toast("Datos reseteados", "Stock y movimientos en cero", "s");
      setResetConf(false);
      await loadData(me.id, me.role);
    } catch(e) { toast("Error en reset", ""+e.message, "e"); }
    setResetBusy(false);
  }

  // ── PERSISTENCIA LOCAL ─────────────────────────────────────────────────────
  // [localStorage eliminado — pedidos y estado en Supabase]

  // ── SUPABASE DATA LOADERS ────────────────────────────────────────────────────
  async function loadCarousels() {
    try {
      const { data } = await sb.from("offer_carousels")
        .select("*")
        .order("sort_order", { ascending: true });
      setCarousels(data || []);
    } catch(e) { /* no carousel table yet = ok */ }
  }

  async function uploadCarouselImg(file) {
    if (!file) return;
    if (file.size > 3000000) { toast("Imagen muy grande","Máx 3MB","e"); return; }
    setCImgUp(true);
    try {
      var ext = file.name.split(".").pop();
      var path = "carousel/" + Date.now() + "." + ext;
      // Timeout: si la subida no resuelve en 20s, cortar
      var uploadPromise = sb.storage.from("product-images").upload(path, file, { upsert: true });
      var timeout = new Promise(function(_, rej){ setTimeout(function(){ rej(new Error("La subida tardó demasiado. Revisá los permisos del bucket en Supabase.")); }, 20000); });
      var up = await Promise.race([uploadPromise, timeout]);
      if (up.error) throw up.error;
      var pub = sb.storage.from("product-images").getPublicUrl(path);
      setCImg(pub.data.publicUrl);
      toast("✅ Imagen lista","","s");
    } catch(e) { toast("Error al subir", e.message||"No se pudo subir", "e"); }
    finally { setCImgUp(false); }
  }

  async function saveCarousel() {
    if (!cTitle.trim()) return;
    setCSaving(true);
    try {
      const row = { title: cTitle.trim(), subtitle: cSubtitle.trim(), bg_color: cBg, emoji: cEmoji, image_url: cImg||null, link_tab: cLink.trim(), active: true, sort_order: carousels.length };
      var res;
      if (cEditId) {
        res = await sb.from("offer_carousels").update(row).eq("id", cEditId);
      } else {
        res = await sb.from("offer_carousels").insert(row);
      }
      if (res.error) { toast("Error al guardar", res.error.message, "e"); return; }
      await loadCarousels();
      setCTitle(""); setCSubtitle(""); setCBg("#e0224e"); setCEmoji("🔥"); setCLink(""); setCEditId(null); setCImg("");
      toast(cEditId?"Carrusel actualizado":"Carrusel creado","","s");
    } catch(e) { toast("Error",e.message,"e"); }
    finally { setCSaving(false); }
  }

  async function deleteCarousel(id) {
    try {
      await sb.from("offer_carousels").delete().eq("id", id);
      setCDelConf(null);
      await loadCarousels();
      toast("Eliminado","","s");
    } catch(e) { toast("Error",e.message,"e"); }
  }

  async function moveCarousel(row, dir) {
    // dir: -1 = subir, +1 = bajar
    var idx = carousels.findIndex(function(c){ return c.id === row.id; });
    var swapIdx = idx + dir;
    if (swapIdx < 0 || swapIdx >= carousels.length) return;
    var other = carousels[swapIdx];
    // Intercambiar sort_order en DB
    await Promise.all([
      sb.from("offer_carousels").update({ sort_order: swapIdx }).eq("id", row.id),
      sb.from("offer_carousels").update({ sort_order: idx }).eq("id", other.id)
    ]);
    await loadCarousels();
  }

  async function toggleCarouselActive(row) {
    await sb.from("offer_carousels").update({ active: !row.active }).eq("id", row.id);
    await loadCarousels();
  }

  const loadData = useCallback(async function(userId, userRole) {
    try {
      // Products - paginate to get ALL (bypass 1000 row limit)
      var allProds = [];
      var PAGE = 800;
      var from = 0;
      while (true) {
        var pr = await sb.from("products").select("id,sku,name,price,category,emoji,photo_url,is_active,created_by,updated_at").eq("is_active",true).order("name").range(from, from+PAGE-1);
        if (pr.data && pr.data.length > 0) {
          allProds = allProds.concat(pr.data);
          if (pr.data.length < PAGE) break;
          from += PAGE;
        } else { break; }
      }
      setProducts(allProds);

      // Todas las demás consultas en paralelo
      var queries = [
        sb.from("inventory").select("id,user_id,product_id,qty_available,qty_sold,stock_recibido,products(id,sku,name,price,emoji,photo_url,category)").eq("user_id",userId),
        sb.from("contacts").select("*, contact:contact_id(id,name,email,color,role)").eq("user_id",userId),
        sb.from("transfers").select("id,product_id,from_user_id,to_user_id,qty,status,created_at,product:product_id(id,name,emoji,photo_url,sku,price,category),from_user:from_user_id(id,name,color),to_user:to_user_id(id,name,color)").or("from_user_id.eq."+userId+",to_user_id.eq."+userId).order("created_at",{ascending:false}).limit(100),
        sb.from("notifications").select("*").eq("to_user_id",userId).order("created_at",{ascending:false}).limit(50),
        sb.from("sale_logs").select("*, product:product_id(name,sku,emoji)").eq("user_id",userId).order("created_at",{ascending:false}).limit(100),
        sb.from("pedidos").select("*, product:product_id(id,name,sku,price,emoji)").eq("user_id",userId).order("created_at",{ascending:false}).limit(200),
        sb.from("consignaciones").select("id,status,owner_id,vendedora_id,comision_pct,created_at").eq("owner_id",userId).neq("status","cancelada"),
        sb.from("consignaciones").select("id,status,owner_id,vendedora_id,comision_pct,created_at").eq("vendedora_id",userId).neq("status","cancelada"),
      ];

      if (userRole === "superadmin") {
        queries.push(sb.from("users").select("*").order("created_at",{ascending:false}));
        queries.push(sb.from("admin_dashboard").select("*").single());
      }

      var results = await Promise.all(queries);
      var inv = results[0], cts = results[1], tx = results[2], nf = results[3];
      var lg = results[4], ped = results[5], envC = results[6], recC = results[7];

      if (inv.data)   setInventory(inv.data);
      if (cts.data)   setContacts(cts.data.map(function(c){ return c.contact; }).filter(Boolean));
      if (tx.data)    setTransfers(tx.data);
      if (nf.data)    setNotifs(nf.data);
      if (lg.data)    setLogs(lg.data);
      if (!ped.error) setPedidos(ped.data||[]);
      setConsignActivas([...(envC.data||[]),...(recC.data||[])]);

      if (userRole === "superadmin") {
        var us = results[8]; var stats = results[9];
        if (us.data)    setAllUsers(us.data);
        if (stats.data) setAdminStats(stats.data);
      }
    } catch(e) { console.error("loadData error:", e); }
  }, []);

  // ── AUTH INIT ────────────────────────────────────────────────────────────────
  useEffect(function(){
    // Carga (o espera) el perfil del usuario tras autenticarse.
    // El trigger de Supabase crea la fila en `users`; reintentamos por si tarda.
    async function ensureProfile(session, intentos){
      intentos = intentos || 0;
      var pr = await sb.from("users").select("*").eq("id", session.user.id).maybeSingle();
      if (pr.data) {
        // Completar nombre/foto desde Google si vinieron y faltan
        var meta = session.user.user_metadata || {};
        var nombreGoogle = meta.full_name || meta.name || "";
        var fotoGoogle   = meta.avatar_url || meta.picture || "";
        var patch = {};
        if (nombreGoogle && (!pr.data.name || pr.data.name===session.user.email)) patch.name = nombreGoogle;
        if (fotoGoogle && !pr.data.avatar_url) patch.avatar_url = fotoGoogle;
        if (Object.keys(patch).length) {
          var upd = await sb.from("users").update(patch).eq("id", session.user.id).select("*").maybeSingle();
          if (upd.data) pr.data = upd.data;
        }
        setMe(pr.data);
        loadData(pr.data.id, pr.data.role);  // no await — no bloquear
        loadCarousels();
        return true;
      }
      // No existe la fila todavía. Reintentar varias veces (la sesión RLS tarda).
      if (intentos < 2) {
        await new Promise(function(r){ setTimeout(r, 200); });
        return ensureProfile(session, intentos + 1);
      }
      // Tras varios intentos sin encontrarla → crearla como revendedora.
      // upsert evita romper si la fila en realidad ya existía (timing RLS).
      var meta2 = session.user.user_metadata || {};
      var nuevoNombre = meta2.full_name || meta2.name || (session.user.email||"").split("@")[0];
      var nuevoColor  = ["#e0224e","#8b5cf6","#0ea5e9","#10b981","#f59e0b","#ec4899"][Math.floor(Math.random()*6)];
      var ins = await sb.from("users").upsert({
        id:         session.user.id,
        email:      session.user.email,
        name:       nuevoNombre,
        role:       "reseller",
        color:      nuevoColor,
        avatar_url: meta2.avatar_url || meta2.picture || null
      }, { onConflict: "id", ignoreDuplicates: false }).select("*").maybeSingle();
      if (ins.data) {
        var perfilFinal = ins.data;
        try {
          var invitado = await aplicarInvitacionSiExiste(ins.data.id, session.user.email);
          if (invitado) perfilFinal = invitado;
        } catch(e) { /* noop */ }
        setMe(perfilFinal);
        loadData(perfilFinal.id, perfilFinal.role);
        loadCarousels();
        return true;
      }
      return false;
    }

    // Red de seguridad: nunca quedar trabado en "Conectando..."
    var bootSafety = setTimeout(function(){ setLoading(false); }, 6000);

    // Renovar sesión SIEMPRE que la app vuelve a primer plano (fix cliente congelado en móvil)
    function onVisible() {
      if (document.visibilityState === "visible") {
        try { sb.auth.startAutoRefresh && sb.auth.startAutoRefresh(); } catch(e){}
        sb.auth.refreshSession().catch(function(){});
      } else {
        try { sb.auth.stopAutoRefresh && sb.auth.stopAutoRefresh(); } catch(e){}
      }
    }
    document.addEventListener("visibilitychange", onVisible);
    window.addEventListener("focus", onVisible);
    window.addEventListener("online", function(){ sb.auth.refreshSession().catch(function(){}); });

    sb.auth.getSession().then(async function(res){
      var session = res.data.session;
      // Apagar el splash apenas sabemos si hay sesión o no; el perfil carga aparte
      clearTimeout(bootSafety);
      setLoading(false);
      if (session) { try { await ensureProfile(session); } catch(e){ /* no bloquear */ } }
    }).catch(function(){
      clearTimeout(bootSafety);
      setLoading(false);
    });

    var sub = sb.auth.onAuthStateChange(async function(event, session){
      if (event === "SIGNED_OUT") { setMe(null); setAuthOk(false); setProducts([]); setInventory([]); setContacts([]); setLogs([]); }
      // Al volver de Google (o cualquier login nuevo), asegurar perfil
      if (event === "SIGNED_IN" && session) { try { await ensureProfile(session); } catch(e){ /* no bloquear */ } }
    });
    return function(){ clearTimeout(bootSafety); sub.data.subscription.unsubscribe(); };
  }, [loadData]);

  // ── REALTIME ────────────────────────────────────────────────────────────────
  useEffect(function(){
    if (!me) return;
    var ch = sb.channel("stockpro_"+me.id)
      .on("postgres_changes",{event:"*",schema:"public",table:"notifications",filter:"to_user_id=eq."+me.id},function(){
        sb.from("notifications").select("*").eq("to_user_id",me.id).order("created_at",{ascending:false}).limit(50).then(function(r){ if(r.data) setNotifs(r.data); });
      })
      .on("postgres_changes",{event:"*",schema:"public",table:"transfers",filter:"to_user_id=eq."+me.id},function(){
        sb.from("transfers").select("*, product:product_id(id,name,emoji,photo_url,sku,price,category), from_user:from_user_id(id,name,email,color), to_user:to_user_id(id,name,email,color)").or("from_user_id.eq."+me.id+",to_user_id.eq."+me.id).order("created_at",{ascending:false}).then(function(r){ if(r.data) setTransfers(r.data); });
      })
      .subscribe();
    return function(){ sb.removeChannel(ch); };
  }, [me]);

  // ── LOGIN ────────────────────────────────────────────────────────────────────
  async function doLoginGoogle() {
    setAuthErr("");
    var res = await sb.auth.signInWithOAuth({
      provider: "google",
      options: { redirectTo: window.location.origin }
    });
    if (res.error) { setAuthErr("No se pudo iniciar con Google: " + res.error.message); shake(); }
    // El redirect a Google ocurre solo; al volver, onAuthStateChange → ensureProfile
  }

  async function doLogin() {
    if (loggingIn) return;
    setAuthErr("");
    setLoggingIn(true);
    try {
      // Timeout de 12 segundos — si Supabase no responde, mostrar error claro
      var timeout = new Promise(function(_, rej){
        setTimeout(function(){ rej(new Error("Tiempo de espera agotado. Verificá tu conexión a internet e intentá de nuevo."));  }, 12000);
      });
      var res = await Promise.race([
        sb.auth.signInWithPassword({email:aEmail.trim().toLowerCase(), password:aPass}),
        timeout
      ]);
      if (res.error) {
        setAuthErr(res.error.message==="Invalid login credentials"?"Email o contraseña incorrectos.":res.error.message);
        shake();
      } else {
        setAuthOk(true);  // cambiar pantalla de inmediato
      }
    } catch(e) {
      setAuthErr((e&&e.message)||"No se pudo conectar. Revisá tu internet.");
      shake();
    } finally {
      setLoggingIn(false);
    }
  }

  // Si hay una invitación pendiente para este email, aplica el rol/jerarquía
  // definidos por quien invitó, y borra la invitación (se consume una sola vez).
  async function aplicarInvitacionSiExiste(userId, email) {
    try {
      var res = await sb.rpc("rpc_consumir_invitacion", { p_email: email });
      return res.data || null;
    } catch(e) { return null; }
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
      var invitada = await aplicarInvitacionSiExiste(res.data.user.id, em);
      var pr = invitada ? {data:invitada} : await sb.from("users").select("*").eq("id",res.data.user.id).single();
      if (pr.data){ setMe(pr.data); await loadData(pr.data.id, pr.data.role); toast("Cuenta creada!","Bienvenida "+aName,"s"); }
    }
  }

  // ── LOGOUT ──────────────────────────────────────────────────────────────────
  // ── FASE 1: Roles y jerarquía ──────────────────────────────────────────

  // Cargar info de jerarquía propia (líder, empresa, comisión) + mi equipo
  async function loadJerarquia() {
    if (!me) return;
    try {
      var tareas = [
        me.lider_id   ? sb.from("users").select("id,name").eq("id", me.lider_id).maybeSingle()   : Promise.resolve({data:null}),
        me.empresa_id ? sb.from("users").select("id,name").eq("id", me.empresa_id).maybeSingle()  : Promise.resolve({data:null}),
        (me.role==="empresaria"||me.role==="lider"||me.role==="superadmin")
          ? sb.from("users").select("id,name,email,role,lider_id,empresa_id,comision_lider_pct,color,avatar_url").neq("id", me.id).order("role").order("name")
          : Promise.resolve({data:null}),
        (me.role==="empresaria"||me.role==="superadmin")
          ? sb.from("pending_invites").select("*").eq("invited_by", me.id).order("created_at", {ascending:false})
          : Promise.resolve({data:null}),
      ];
      var res = await Promise.all(tareas);
      setCtaJerInfo({ lider: res[0].data, empresa: res[1].data });
      if (res[2].data) setMiEquipo(res[2].data);
      if (res[3].data) setMisInvites(res[3].data);
    } catch(e) { /* noop */ }
  }

  // Guardar cambios de "Mi Cuenta"
  async function doSaveMiCuenta() {
    if (!ctaName.trim()) { toast("Falta el nombre", "", "e"); return; }
    if (ctaPass && ctaPass.length < 6) { toast("La contraseña debe tener al menos 6 caracteres", "", "e"); return; }
    if (ctaPass && ctaPass !== ctaPass2) { toast("Las contraseñas no coinciden", "", "e"); return; }
    setCtaSaving(true);
    try {
      var upd = await sb.from("users").update({
        name: ctaName.trim(), telefono: ctaTel.trim(), direccion: ctaDir.trim(),
        dni: ctaDni.trim(), codigo_vendedora: ctaCodVend.trim(), localidad: ctaLocalidad.trim()
      }).eq("id", me.id).select().single();
      if (upd.error) { toast("Error al guardar", upd.error.message, "e"); setCtaSaving(false); return; }
      if (ctaPass) {
        var pr = await sb.auth.updateUser({ password: ctaPass });
        if (pr.error) { toast("Datos guardados, pero la contraseña falló", pr.error.message, "e"); setCtaSaving(false); return; }
      }
      setMe(upd.data);
      setCtaPass(""); setCtaPass2("");
      toast("Datos actualizados", "", "s");
    } catch(e) { toast("Error", e.message, "e"); }
    setCtaSaving(false);
  }

  // Crear invitación (superadmin invita empresaria; empresaria invita líder/vendedora)
  async function doInvitar() {
    if (!invName.trim() || !invEmail.trim()) { toast("Completá nombre y email", "", "e"); return; }
    setInvSaving(true);
    try {
      var row = {
        name: invName.trim(),
        email: invEmail.trim().toLowerCase(),
        role: me.role === "superadmin" ? "empresaria" : invRole,
        invited_by: me.id,
        empresa_id: me.role === "empresaria" ? me.id : null,
        lider_id: (me.role === "empresaria" && invRole === "reseller" && invLiderId) ? invLiderId : null,
        comision_lider_pct: (me.role === "empresaria" && invRole === "lider") ? parseFloat(invComis||0) : 0
      };
      var res = await sb.from("pending_invites").insert(row).select().single();
      if (res.error) { toast("Error al invitar", res.error.message, "e"); setInvSaving(false); return; }
      setMisInvites(function(prev){ return [res.data, ...prev]; });
      setInvName(""); setInvEmail(""); setInvLiderId(""); setInvComis("10");
      toast("Invitación creada", "Compartí el link de registro con "+row.name, "s");
    } catch(e) { toast("Error", e.message, "e"); }
    setInvSaving(false);
  }

  async function doCancelarInvite(id) {
    await sb.from("pending_invites").delete().eq("id", id);
    setMisInvites(function(prev){ return prev.filter(function(i){ return i.id !== id; }); });
  }

  // Asignar/cambiar líder de una vendedora (empresaria)
  async function doAsignarLider(userId, liderId) {
    var res = await sb.rpc("rpc_editar_jerarquia_usuario", { p_user_id: userId, p_set_lider: true, p_lider_id: liderId || null });
    if (res.error) { toast("No se pudo asignar", res.error.message, "e"); return; }
    setMiEquipo(function(prev){ return prev.map(function(u){ return u.id===userId ? Object.assign({},u,{lider_id:liderId||null}) : u; }); });
    toast("Asignación actualizada", "", "s");
  }

  // Actualizar % de comisión de un líder (empresaria)
  async function doGuardarComisionLider(userId) {
    var val = parseFloat(equipoEditComis[userId]);
    if (isNaN(val) || val < 0 || val > 100) { toast("Porcentaje inválido", "", "e"); return; }
    var res = await sb.rpc("rpc_editar_jerarquia_usuario", { p_user_id: userId, p_comision: val });
    if (res.error) { toast("Error", res.error.message, "e"); return; }
    setMiEquipo(function(prev){ return prev.map(function(u){ return u.id===userId ? Object.assign({},u,{comision_lider_pct:val}) : u; }); });
    toast("Comisión actualizada", "", "s");
  }

  // Cambiar el rol de un usuario del equipo (superadmin: cualquiera / empresaria: los suyos)
  async function doCambiarRol(userId, nuevoRol) {
    var res = await sb.rpc("rpc_editar_jerarquia_usuario", { p_user_id: userId, p_nuevo_role: nuevoRol });
    if (res.error) { toast("No se pudo cambiar el rol", res.error.message, "e"); return; }
    var patch = { role: nuevoRol };
    if (nuevoRol === "empresaria") { patch.empresa_id = null; patch.lider_id = null; }
    else if (nuevoRol === "lider") { patch.lider_id = null; }
    setMiEquipo(function(prev){ return prev.map(function(u){ return u.id===userId ? Object.assign({},u,patch) : u; }); });
    toast("Rol actualizado", "", "s");
  }

  // ── FASE 2: Pedidos en cascada ──────────────────────────────────────────

  async function loadPedidosEspeciales() {
    if (!me) return;
    setPedEspLoading(true);
    try {
      var res = await sb.from("pedidos_especiales")
        .select("*, items:pedidos_especiales_items(id,product_id,qty,precio_unit,subtotal,color,talle,product:product_id(id,name,sku,emoji,photo_url)), vendedor:vendedor_id(id,name,color,codigo_vendedora,dni,telefono,localidad), lider:lider_id(id,name), empresa:empresa_id(id,name), campania_rel:campania_id(id,nombre,estado)")
        .order("created_at", {ascending:false})
        .limit(200);
      if (res.data) setPedEspList(res.data);
    } catch(e) { /* noop */ }
    setPedEspLoading(false);
  }

  function doAgregarAlCarrito() {
    if (!peProdId) { toast("Elegí un producto", "", "e"); return; }
    var p = products.find(function(x){ return x.id===peProdId; });
    if (!p) return;
    setPeCarrito(function(prev){
      return prev.concat([{
        product_id: p.id, name: p.name, sku: p.sku, emoji: p.emoji, photo_url: p.photo_url,
        price: p.price, qty: peQty, color: peColor.trim(), talle: peTalle.trim()
      }]);
    });
    setPeProdId(""); setPeProdSrch(""); setPeQty(1); setPeColor(""); setPeTalle("");
  }

  function doQuitarDelCarrito(idx) {
    setPeCarrito(function(prev){ return prev.filter(function(_,i){ return i!==idx; }); });
  }

  async function doCrearPedidoEspecial() {
    if (peCarrito.length===0) { toast("Agregá al menos un producto al carrito", "", "e"); return; }
    if (!peCliNom.trim()) { toast("Falta el nombre del cliente", "", "e"); return; }
    setPeSaving(true);
    try {
      var itemsPayload = peCarrito.map(function(it){
        return { product_id: it.product_id, qty: it.qty, color: it.color, talle: it.talle };
      });
      var res = await sb.rpc("rpc_crear_pedido_especial", {
        p_items: itemsPayload,
        p_cliente_nombre: peCliNom.trim(), p_cliente_telefono: peCliTel.trim(), p_nota: peNota.trim(),
        p_campania_id: peCampaniaId || null
      });
      if (res.error) { toast("Error al crear pedido", res.error.message, "e"); setPeSaving(false); return; }
      toast("📝 Pedido guardado", "Enviálo cuando quieras desde la lista", "s");
      setPeShowForm(false); setPeCarrito([]); setPeCliNom(""); setPeCliTel(""); setPeNota(""); setPeCampaniaId("");
      await loadPedidosEspeciales();
    } catch(e) { toast("Error", e.message, "e"); }
    setPeSaving(false);
  }

  async function doEnviarPedidoEspecial(id) {
    setPeBusy(function(prev){ return Object.assign({},prev,{[id]:true}); });
    try {
      var res = await sb.rpc("rpc_enviar_pedido_especial", { p_pedido_id: id });
      if (res.error) { toast("Error al enviar", res.error.message, "e"); setPeBusy(function(prev){ return Object.assign({},prev,{[id]:false}); }); return; }
      toast("✅ Pedido enviado", "Queda pendiente de aprobación", "s");
      await loadPedidosEspeciales();
    } catch(e) { toast("Error", e.message, "e"); }
    setPeBusy(function(prev){ return Object.assign({},prev,{[id]:false}); });
  }

  async function doAccionPedidoEsp(id, rpcName, extraParams) {
    setPeBusy(function(prev){ return Object.assign({},prev,{[id]:true}); });
    try {
      var params = Object.assign({ p_pedido_id: id, p_observ: peObserv[id]||"" }, extraParams||{});
      var res = await sb.rpc(rpcName, params);
      if (res.error) { toast("Error", res.error.message, "e"); setPeBusy(function(prev){ return Object.assign({},prev,{[id]:false}); }); return; }
      toast("Listo", "", "s");
      setPeObserv(function(prev){ return Object.assign({},prev,{[id]:""}); });
      await loadPedidosEspeciales();

      // Al quedar "Aprobado" (aprobó la empresaria), generar el PDF automáticamente
      if (rpcName === "rpc_empresaria_decidir_pedido" && extraParams && extraParams.p_aprobar) {
        var pedido = pedEspList.find(function(x){ return x.id===id; });
        if (pedido) {
          toast("Generando PDF...", "", "i");
          var url = await generarYSubirPdfPedido(Object.assign({}, pedido, {estado:"aprobado"}));
          if (url) { toast("📄 PDF generado", "Disponible para el administrador", "s"); await loadPedidosEspeciales(); }
        }
      }
    } catch(e) { toast("Error", e.message, "e"); }
    setPeBusy(function(prev){ return Object.assign({},prev,{[id]:false}); });
  }

  async function verHistorialPedidoEsp(id) {
    if (peHistOpen === id) { setPeHistOpen(null); return; }
    setPeHistOpen(id);
    var res = await sb.from("pedidos_especiales_historial").select("*, usuario:usuario_id(name)").eq("pedido_id", id).order("created_at",{ascending:true});
    setPeHist(res.data || []);
  }

  function peEstadoInfo(estado) {
    var map = {
      borrador:              {lbl:"📝 Borrador",              bg:"#f1f1f1", col:"#666"},
      pendiente_lider:      {lbl:"Pendiente de líder",      bg:"#fff3e0", col:"#e07800"},
      rechazado_lider:      {lbl:"Rechazado por líder",     bg:"#ffe0e5", col:"#d32"},
      pendiente_empresaria: {lbl:"Pendiente de empresaria", bg:"#fff3e0", col:"#e07800"},
      rechazado_empresaria: {lbl:"Rechazado por empresaria",bg:"#ffe0e5", col:"#d32"},
      aprobado:             {lbl:"Aprobado",                bg:"#e0f2ff", col:"#0369a1"},
      enviado_proveedor:    {lbl:"Enviado a proveedor",     bg:"#f3e8ff", col:"#7c3aed"},
      recibido:             {lbl:"Recibido",                bg:"#e7f9ee", col:"#0a8f4d"},
      listo_entregar:       {lbl:"Listo para entregar",     bg:"#dcfce7", col:"#15803d"},
      entregado:            {lbl:"Entregado",               bg:"#dcfce7", col:"#15803d"},
      cancelado:            {lbl:"Cancelado",               bg:"#f1f1f1", col:"#888"},
    };
    return map[estado] || {lbl:estado, bg:"#eee", col:"#666"};
  }

  // ── FASE 4: Generación automática de PDF del pedido ─────────────────────
  async function generarYSubirPdfPedido(p) {
    try {
      var doc = new jsPDF({ unit: "mm", format: "a4" });
      var pageW = doc.internal.pageSize.getWidth();
      var margin = 14;

      // Encabezado
      doc.setFillColor(224, 34, 78);
      doc.rect(0, 0, pageW, 26, "F");
      doc.setTextColor(255,255,255);
      doc.setFontSize(18);
      doc.setFont(undefined, "bold");
      doc.text("Venta Directa", margin, 12);
      doc.setFontSize(10);
      doc.setFont(undefined, "normal");
      doc.text("Pedido Especial — Orden de Compra", margin, 19);
      doc.setFontSize(11);
      doc.setFont(undefined, "bold");
      doc.text("Pedido #" + (p.numero_pedido||"-"), pageW-margin, 12, {align:"right"});
      doc.setFont(undefined, "normal");
      doc.setFontSize(9);
      doc.text(new Date(p.created_at).toLocaleString("es-AR"), pageW-margin, 19, {align:"right"});

      var y = 34;
      doc.setTextColor(30,30,30);

      // Datos del pedido
      doc.setFontSize(11); doc.setFont(undefined,"bold");
      doc.text("Datos del pedido", margin, y); y += 6;
      doc.setFontSize(9.5); doc.setFont(undefined,"normal");
      var ei = peEstadoInfo(p.estado);
      doc.text("Estado: " + ei.lbl, margin, y);
      doc.text("Campaña: " + (p.campania_rel?p.campania_rel.nombre:(p.campania||"-")), margin+90, y); y += 5;
      if (p.fecha_entrega_estimada) { doc.text("Entrega estimada: " + p.fecha_entrega_estimada, margin, y); y += 5; }
      y += 3;

      // Datos de la vendedora
      doc.setFontSize(11); doc.setFont(undefined,"bold");
      doc.text("Vendedora", margin, y); y += 6;
      doc.setFontSize(9.5); doc.setFont(undefined,"normal");
      var v = p.vendedor||{};
      doc.text("Nombre: " + (v.name||"-"), margin, y);
      doc.text("Código: " + (v.codigo_vendedora||"-"), margin+90, y); y += 5;
      doc.text("DNI: " + (v.dni||"-"), margin, y);
      doc.text("Teléfono: " + (v.telefono||"-"), margin+90, y); y += 5;
      doc.text("Localidad: " + (v.localidad||"-"), margin, y); y += 8;

      // Estructura comercial
      doc.setFontSize(11); doc.setFont(undefined,"bold");
      doc.text("Estructura comercial", margin, y); y += 6;
      doc.setFontSize(9.5); doc.setFont(undefined,"normal");
      doc.text("Líder: " + (p.lider?p.lider.name:"— (directa)"), margin, y);
      doc.text("Empresaria: " + (p.empresa?p.empresa.name:"-"), margin+90, y); y += 8;

      // Tabla de productos (todos los ítems del pedido)
      var itemsPdf = p.items||[];
      autoTable(doc, {
        startY: y,
        head: [["Código","Descripción","Color","Talle","Cant.","P. Unit.","Subtotal"]],
        body: itemsPdf.map(function(it){
          return [
            it.product?it.product.sku||"-":"-",
            it.product?it.product.name:"-",
            it.color||"-",
            it.talle||"-",
            String(it.qty),
            "$"+Number(it.precio_unit).toLocaleString("es-AR"),
            "$"+Number(it.subtotal).toLocaleString("es-AR")
          ];
        }),
        theme: "grid",
        headStyles: { fillColor: [224,34,78], textColor: 255, fontSize: 9 },
        bodyStyles: { fontSize: 9 },
        margin: { left: margin, right: margin }
      });

      var finalY = doc.lastAutoTable.finalY + 8;

      // Resumen
      doc.setFillColor(245,245,245);
      doc.rect(margin, finalY, pageW-margin*2, 20, "F");
      doc.setFontSize(10); doc.setFont(undefined,"bold");
      doc.text("Cantidad total: " + p.qty + " unidad(es)", margin+4, finalY+8);
      doc.setFontSize(13);
      doc.setTextColor(224,34,78);
      doc.text("Total: $" + Number(p.total).toLocaleString("es-AR"), pageW-margin-4, finalY+12, {align:"right"});
      doc.setTextColor(30,30,30);
      finalY += 26;

      if (p.nota) {
        doc.setFontSize(9.5); doc.setFont(undefined,"bold");
        doc.text("Observaciones:", margin, finalY); finalY += 5;
        doc.setFont(undefined,"normal");
        var notaLines = doc.splitTextToSize(p.nota, pageW-margin*2);
        doc.text(notaLines, margin, finalY);
      }

      // Pie de página
      var pageCount = doc.internal.getNumberOfPages();
      for (var i=1; i<=pageCount; i++) {
        doc.setPage(i);
        doc.setFontSize(8); doc.setTextColor(150,150,150);
        doc.text("Emitido: " + new Date().toLocaleDateString("es-AR"), margin, 290);
        doc.text("Página " + i + " de " + pageCount, pageW-margin, 290, {align:"right"});
      }

      var blob = doc.output("blob");
      var fileName = "pedido-" + (p.numero_pedido||p.id) + "-" + Date.now() + ".pdf";
      var up = await sb.storage.from("pedidos-pdf").upload(fileName, blob, { contentType: "application/pdf", upsert: true });
      if (up.error) { console.error("Error subiendo PDF:", up.error); return null; }

      var pub = sb.storage.from("pedidos-pdf").getPublicUrl(fileName);
      var url = pub.data.publicUrl;

      await sb.rpc("rpc_guardar_pdf_pedido", { p_pedido_id: p.id, p_pdf_url: url });
      return url;
    } catch(e) { console.error("Error generando PDF:", e); return null; }
  }

  function compartirPdfPorEmail(p) {
    if (!p.pdf_url) { toast("Todavía no hay PDF generado", "", "e"); return; }
    var asunto = "Pedido #" + (p.numero_pedido||"") + " — " + (p.vendedor?p.vendedor.name:"");
    var listaProductos = (p.items||[]).map(function(it){ return "- "+it.qty+"x "+(it.product?it.product.name:"-"); }).join("\n");
    var cuerpo = "Adjunto el pedido especial.\n\nDescargar PDF: " + p.pdf_url + "\n\nVendedora: " + (p.vendedor?p.vendedor.name:"") +
      "\nProductos:\n" + listaProductos + "\n\nCantidad total: " + p.qty + " un.\nTotal: $" + Number(p.total).toLocaleString("es-AR");
    window.location.href = "mailto:?subject=" + encodeURIComponent(asunto) + "&body=" + encodeURIComponent(cuerpo);
  }

  // ── FASE 5: Campañas ─────────────────────────────────────────────────────
  async function loadCampanias() {
    if (!me) return;
    try {
      var res = await sb.from("campanias").select("*").order("created_at",{ascending:false});
      if (res.data) setCampanias(res.data);
    } catch(e) { /* noop */ }
  }

  async function doCrearCampania() {
    if (!campNombre.trim()) { toast("Falta el nombre de la campaña", "", "e"); return; }
    setCampSaving(true);
    try {
      var res = await sb.rpc("rpc_crear_campania", { p_nombre: campNombre.trim() });
      if (res.error) { toast("Error", res.error.message, "e"); setCampSaving(false); return; }
      setCampNombre("");
      toast("Campaña creada", "Dura 30 días", "s");
      await loadCampanias();
    } catch(e) { toast("Error", e.message, "e"); }
    setCampSaving(false);
  }

  async function doCambiarEstadoCampania(id, nuevoEstado) {
    var res = await sb.rpc("rpc_cambiar_estado_campania", { p_campania_id: id, p_estado: nuevoEstado });
    if (res.error) { toast("Error", res.error.message, "e"); return; }
    setCampanias(function(prev){ return prev.map(function(c){ return c.id===id ? Object.assign({},c,{estado:nuevoEstado}) : c; }); });
    toast(nuevoEstado==="abierta"?"Campaña reabierta":"Campaña cerrada", "", "s");
  }

  // ── FASE 3: Resúmenes ────────────────────────────────────────────────────
  async function loadResumen() {
    if (!me || (me.role!=="lider" && me.role!=="empresaria")) return;
    setResumenLoading(true);
    try {
      var fn = me.role==="lider" ? "rpc_resumen_lider" : "rpc_resumen_empresaria";
      var param = me.role==="lider" ? {p_lider_id: me.id} : {p_empresa_id: me.id};
      var res = await sb.rpc(fn, param);
      if (res.data && res.data.length) setResumen(res.data[0]);
    } catch(e) { /* noop */ }
    setResumenLoading(false);
  }

  function compartirResumenWhatsapp() {
    if (!resumen) return;
    var msg;
    if (me.role==="lider") {
      msg = "📋 Resumen de "+me.name+"\n\n"+
        "🔸 Pedidos pendientes de mi aprobación: "+resumen.pedidos_pendientes+"\n"+
        "🔸 Pedidos en curso: "+resumen.pedidos_en_curso+"\n"+
        "💰 Ventas del mes: $"+Number(resumen.ventas_mes||0).toLocaleString("es-AR")+"\n"+
        "📦 Unidades vendidas: "+resumen.unidades_mes+"\n"+
        "💵 Mi comisión del mes: $"+Number(resumen.comision_mes||0).toLocaleString("es-AR")+"\n"+
        "👥 Vendedoras a mi cargo: "+resumen.equipo_count;
    } else {
      msg = "📊 Resumen de "+me.name+"\n\n"+
        "🔸 Pedidos pendientes de mi aprobación: "+resumen.pedidos_pendientes+"\n"+
        "🔸 Pedidos en curso: "+resumen.pedidos_en_curso+"\n"+
        "💰 Ventas del mes: $"+Number(resumen.ventas_mes||0).toLocaleString("es-AR")+"\n"+
        "📦 Unidades vendidas: "+resumen.unidades_mes+"\n"+
        "💵 Comisiones a líderes: $"+Number(resumen.comision_lideres_mes||0).toLocaleString("es-AR")+"\n"+
        "💵 Comisiones a vendedoras: $"+Number(resumen.comision_vendedoras_mes||0).toLocaleString("es-AR")+"\n"+
        "📈 Rentabilidad del mes: $"+Number(resumen.rentabilidad_mes||0).toLocaleString("es-AR")+"\n"+
        "👥 Líderes: "+resumen.lideres_count+" · Vendedoras: "+resumen.vendedoras_count;
    }
    window.open("https://wa.me/?text="+encodeURIComponent(msg), "_blank");
  }

  async function doLogout() {
    var n = me.name;
    // Resetear UI de inmediato — no esperar a Supabase
    setMe(null); setAuthOk(false); setTab("stock"); setAEmail(""); setAPass("");
    toast("Hasta luego, "+n+"!","","i");
    try { await sb.auth.signOut(); } catch(e) { /* la sesión local ya se limpió */ }
  }

  // ── SELL ────────────────────────────────────────────────────────────────────
  async function doSell(invRow) {
    if (invRow.qty_available<=0) return;
    var prod = invRow.products || products.find(function(p){ return p.id===invRow.product_id; });

    // ¿Este producto vino por una consignación confirmada?
    var originTransfer = transfers.find(function(t){
      return t.to_user_id===me.id &&
             t.product_id===invRow.product_id &&
             t.status==="confirmed";
    });

    // Venta atómica en el servidor (evita condiciones de carrera)
    var r = await sb.rpc("rpc_vender_unidad", { p_inv_id: invRow.id });
    if (r.error){ toast("No se pudo vender", r.error.message, "e"); return; }
    // La RPC decrementa stock_propio; si era recibido, ajustamos el canal
    var updated = r.data;
    if ((originTransfer || invRow.source==="consigna") && updated) {
      var fix = await sb.from("inventory").update({
        stock_recibido: Math.max(0, (invRow.stock_recibido||0) - 1),
        stock_propio:   invRow.stock_propio||0
      }).eq("id", invRow.id).select().single();
      if (fix.data) updated = fix.data;
    }
    if (updated) setInventory(function(p){ return p.map(function(i){ return i.id===invRow.id?updated:i; }); });

    var source = originTransfer ? "consignment" : "own_stock";

    // Registrar venta
    var logInsert = await sb.from("sale_logs").insert({
      user_id: me.id,
      product_id: invRow.product_id,
      qty: 1,
      sale_price: prod ? prod.price : 0,
      source: source
    }).select("*, product:product_id(name,sku,emoji)").single();

    if (logInsert.data) {
      setLogs(function(prev){ return [logInsert.data, ...prev]; });
    }

    // Ledger
    await logMovimiento({
      product_id: invRow.product_id, qty: -1,
      estado_anterior: originTransfer ? "en_consigna" : "stock_central",
      estado_nuevo: "vendido", referencia_tipo: "venta",
      nota: "Venta directa de " + (prod?prod.name:"producto")
    });

    // Si era consignación, avisar al remitente original
    if (originTransfer) {
      await sb.from("notifications").insert({
        to_user_id: originTransfer.from_user_id,
        from_name: me.name,
        type: "sale",
        message: "💰 "+me.name+" vendió 1x "+(prod?prod.name:"producto")+" que le enviaste en consignación!"
      });
    }

    toast("Venta registrada", prod?prod.name:"", "s");
  }

  // ── TRANSFER (send to contact, pending confirmation) ─────────────────────────
  async function doTx() {
    if (txQty>txModal.qty_available){ toast("Stock insuficiente","","e"); return; }
    var prod  = txModal.products || products.find(function(p){ return p.id===txModal.product_id; });
    var tUser = contacts.find(function(c){ return c.id===txTo; });
    // 1+2. Descontar stock y crear transfer de forma atómica
    var r = await sb.rpc("rpc_transferir", { p_inv_id: txModal.id, p_to_user: txTo, p_qty: txQty });
    if (r.error){ toast("No se pudo enviar", r.error.message, "e"); return; }
    var txInsert = r.data;
    // Reflejar el descuento localmente
    setInventory(function(p){ return p.map(function(i){
      return i.id===txModal.id
        ? Object.assign({}, i, {
            qty_available: i.qty_available - txQty,
            stock_propio: Math.max(0, (i.stock_propio!=null?i.stock_propio:i.qty_available) - txQty)
          })
        : i;
    }); });
    // 3. Ledger
    await logMovimiento({
      product_id: txModal.product_id, qty: -txQty,
      estado_anterior:"stock_central", estado_nuevo:"en_transito",
      related_user_id: txTo,
      referencia_tipo:"transfer", referencia_id: txInsert?txInsert.id:null,
      nota:"Enviado a "+(tUser?tUser.name:"")
    });
    // 4. Notificar
    await sb.from("notifications").insert({
      to_user_id:txTo, from_name:me.name, type:"transfer",
      message:me.name+" te envió "+txQty+"× "+(prod?prod.name:"producto")+". Confirmá la recepción!"
    });
    toast("Enviado!","Esperando confirmación de "+(tUser?tUser.name:""),"s");
    setTxModal(null);
    await loadData(me.id,me.role);
  }

  // ── PEDIDOS FUNCTIONS ────────────────────────────────────────────────────────
  async function doAddPedido() {
    if (!pedNombre.trim()) {
      toast("Falta el nombre del cliente", "Completá el campo Nombre arriba del formulario", "e");
      return;
    }
    var prod = products.find(function(p){ return p.id===pedProdId; });
    var totalPed = prod ? (parseFloat(prod.price||0) * pedQty) : 0;
    var senaPed  = parseFloat(pedSena||0) || 0;
    var row = {
      user_id:       me.id,
      nombre:        pedNombre.trim(),
      wa:            pedWA.trim()||null,
      product_id:    pedProdId||null,
      product_name:  prod ? prod.name  : null,
      product_sku:   prod ? prod.sku   : null,
      product_price: prod ? prod.price : 0,
      qty:           pedQty,
      total:         totalPed,
      sena:          senaPed,
      pagado:        false,
      nota:          pedNota.trim()||null,
      estado:        "pendiente",
    };
    var insPromise = sb.from("pedidos").insert(row).select("*, product:product_id(id,name,sku,price,emoji)").single();
    var timeoutP = new Promise(function(_, rej){ setTimeout(function(){ rej(new Error("La conexión tardó demasiado. Reintentá.")); }, 10000); });
    var data, error;
    try {
      var r = await Promise.race([insPromise, timeoutP]);
      data = r.data; error = r.error;
    } catch(e) { toast("Error al guardar pedido", e.message, "e"); return; }
    if (error) { toast("Error al guardar pedido", error.message, "e"); return; }
    setPedidos(function(prev){ return [data, ...prev]; });
    setPedNombre(""); setPedWA(""); setPedProdId(""); setPedQty(1); setPedNota(""); setPedPSrch(""); setPedSena("");
    setShowPedForm(false);
    toast("Pedido guardado!", pedNombre.trim()+(prod?" — "+prod.name:""), "s");
  }

  // Helper: promesa con timeout para evitar cuelgues silenciosos en móvil
  function conTimeout(promesa, ms) {
    return Promise.race([
      promesa,
      new Promise(function(_, rej){ setTimeout(function(){ rej(new Error("Sin respuesta del servidor. Reintentá.")); }, ms||8000); })
    ]);
  }

  async function doMarcarPagado(id) {
    var error;
    try {
      var r = await conTimeout(sb.from("pedidos").update({ pagado: true }).eq("id", id));
      error = r.error;
    } catch(e) {
      // Reintentar una vez tras renovar sesión
      try {
        await sb.auth.refreshSession();
        var r2 = await conTimeout(sb.from("pedidos").update({ pagado: true }).eq("id", id));
        error = r2.error;
      } catch(e2) { toast("Error", e2.message, "e"); return; }
    }
    if (error) { toast("Error", error.message, "e"); return; }
    setPedidos(function(prev){ return prev.map(function(p){ return p.id===id ? Object.assign({},p,{pagado:true}) : p; }); });
    toast("✅ Pedido pagado", "", "s");
  }

  function waPedidoGracias(ped) {
    var prod = ped.product || products.find(function(p){ return p.id===ped.product_id; });
    var total = ped.total || (ped.product_price||0) * (ped.qty||1);
    var sena  = ped.sena || 0;
    var saldo = Math.max(0, total - sena);
    var msg = "¡Hola " + ped.nombre + "! 🌸\n\n";
    msg += "¡Gracias por tu compra! 💕\n\n";
    if (prod) msg += "🛍️ " + (ped.qty||1) + "× " + prod.name + "\n";
    if (total>0) msg += "💰 Total: " + fmtARS(total) + "\n";
    if (sena>0) {
      msg += "✅ Seña recibida: " + fmtARS(sena) + "\n";
      msg += "📌 Saldo a pagar: " + fmtARS(saldo) + "\n";
    }
    msg += "\n¡Cualquier cosa me avisás! 😊";
    var tel = (ped.wa||"").replace(/\D/g,"");
    var url = "https://wa.me/" + tel + "?text=" + encodeURIComponent(msg);
    window.open(url, "_blank");
  }

  async function doEntregarPedido(id) {
    var ped = pedidos.find(function(p){ return p.id===id; });
    if (!ped) return;
    var now = new Date().toISOString();
    var error;
    try {
      var r = await conTimeout(sb.from("pedidos").update({ estado: "entregado", entregado_at: now }).eq("id", id));
      error = r.error;
    } catch(e) {
      try {
        await sb.auth.refreshSession();
        var r2 = await conTimeout(sb.from("pedidos").update({ estado: "entregado", entregado_at: now }).eq("id", id));
        error = r2.error;
      } catch(e2) { toast("Error", e2.message, "e"); return; }
    }
    if (error) { toast("Error", error.message, "e"); return; }

    // SIEMPRE registrar la venta al entregar (independiente del inventario)
    if (ped.product_id) {
      var precioVenta = parseFloat(ped.total || ped.product_price || 0);
      var slRes = await sb.from("sale_logs").insert({
        user_id: me.id, product_id: ped.product_id,
        qty: ped.qty, sale_price: precioVenta / (ped.qty||1), source: "pedido"
      });
      if (slRes.error) toast("Aviso", "Venta no registrada: "+slRes.error.message, "e");

      // Descontar inventario solo si existe y alcanza
      var invRow = inventory.find(function(i){ return i.product_id===ped.product_id; });
      if (invRow && invRow.qty_available >= ped.qty) {
        await sb.from("inventory")
          .update({ qty_available: invRow.qty_available - ped.qty, qty_sold: (invRow.qty_sold||0) + ped.qty })
          .eq("id", invRow.id);
        var estadoAnt = invRow.source==="consigna" ? "en_consigna" : "stock_central";
        await logMovimiento({
          product_id: ped.product_id, qty: -(ped.qty),
          estado_anterior: estadoAnt, estado_nuevo: "vendido",
          referencia_tipo: "pedido", referencia_id: id,
          nota: "Pedido entregado a " + ped.nombre
        });
      }
    }

    setPedidos(function(prev){ return prev.map(function(p){ return p.id===id ? Object.assign({},p,{estado:"entregado",entregado_at:now}) : p; }); });
    toast("✅ Pedido entregado!", ped.nombre, "s");
    await loadData(me.id, me.role);
  }

  async function doCancelarPedido(id) {
    var { error } = await sb.from("pedidos").update({ estado: "cancelado" }).eq("id", id);
    if (error) { toast("Error", error.message, "e"); return; }
    setPedidos(function(prev){ return prev.map(function(p){ return p.id===id ? Object.assign({},p,{estado:"cancelado"}) : p; }); });
    toast("Pedido cancelado", "", "i");
  }

  async function doDeletePedido(id) {
    var { error } = await sb.from("pedidos").delete().eq("id", id);
    if (error) { toast("Error", error.message, "e"); return; }
    setPedidos(function(prev){ return prev.filter(function(p){ return p.id!==id; }); });
  }

  function doWAPedido(ped) {
    var prod = ped.product_name ? ped.product_name : "";
    var msg = "Hola " + ped.nombre + "! 👋\n";
    if (prod) msg += "Te confirmo tu pedido:\n📦 " + prod + (ped.qty>1?" x"+ped.qty:"") + "\n";
    if (ped.product_price>0) msg += "💰 " + fmtARS(ped.product_price * ped.qty) + "\n";
    if (ped.nota) msg += "📝 " + ped.nota + "\n";
    msg += "\n¡Avisame cuando podés coordinar la entrega!";
    var wa = ped.wa.replace(/\D/g,"");
    var url = wa ? "https://wa.me/"+wa+"?text="+encodeURIComponent(msg) : "https://wa.me/?text="+encodeURIComponent(msg);
    window.open(url, "_blank");
  }


  // ── LEDGER: registrar movimiento de stock ─────────────────────────────────
  async function logMovimiento(opts) {
    // opts: { product_id, user_id, related_user_id, qty, estado_anterior, estado_nuevo, referencia_tipo, referencia_id, nota }
    try {
      await sb.from("stock_movements").insert({
        product_id:      opts.product_id,
        user_id:         opts.user_id || me.id,
        related_user_id: opts.related_user_id || null,
        qty:             opts.qty,
        estado_anterior: opts.estado_anterior || null,
        estado_nuevo:    opts.estado_nuevo,
        referencia_tipo: opts.referencia_tipo || null,
        referencia_id:   opts.referencia_id   || null,
        nota:            opts.nota || null,
      });
    } catch(e) { console.warn("logMovimiento error:", e.message); }
  }

  // ── LOAD RECV INV ON TAB CHANGE ──────────────────────────────────────────────
  useEffect(function(){
    if ((tab==="enviados" || tab==="recibidos") && me) loadRecvInventory();
  }, [tab, transfers]);

  // ── ANULAR ENVÍO (solo admin o emisor, antes de confirmación) ───────────────
  async function anularTransfer(tx) {
    var prod = tx.product || products.find(function(p){ return p.id===tx.product_id; });
    // Restituir stock al emisor
    var srcInvDB2 = await sb.from("inventory").select("*").eq("user_id",tx.from_user_id).eq("product_id",tx.product_id).single();
    if (srcInvDB2.data) {
      await sb.from("inventory").update({qty_available:srcInvDB2.data.qty_available+tx.qty}).eq("id",srcInvDB2.data.id);
    } else {
      await sb.from("inventory").insert({user_id:tx.from_user_id,product_id:tx.product_id,qty_available:tx.qty,qty_sold:0});
    }
    // Cambiar estado a "anulado"
    await sb.from("transfers").update({status:"cancelled"}).eq("id",tx.id);
    // Notificar al destinatario si era pendiente
    if (tx.status==="pending") {
      await sb.from("notifications").insert({
        to_user_id: tx.to_user_id,
        from_name: me.name,
        type: "info",
        message: "❌ "+me.name+" anuló el envío de "+tx.qty+"x "+(prod?prod.name:"producto")+". Stock devuelto al remitente."
      });
    }
    toast("Envío anulado","Stock restituido correctamente","i");
    await loadData(me.id, me.role);
  }

  // ── CONFIRM TRANSFER ─────────────────────────────────────────────────────────
  async function confirmTransfer(tx) {
    var prod = tx.product || products.find(function(p){ return p.id===tx.product_id; });
    try {
      // Confirmación atómica en el servidor (suma al receptor + marca confirmed)
      var r = await sb.rpc("rpc_confirmar_transfer", { p_tx_id: tx.id });
      if (r.error) throw r.error;
      // Actualización optimista
      setTransfers(function(prev){ return prev.map(function(t){ return t.id===tx.id ? Object.assign({},t,{status:"confirmed"}) : t; }); });
      await sb.from("notifications").insert({
        to_user_id: tx.from_user_id, from_name: me.name, type: "confirm",
        message: me.name + " confirmó la recepción de " + tx.qty + "× " + (prod ? prod.name : "producto") + " ✅"
      });
      await sb.from("notifications").update({ read: true }).eq("to_user_id", me.id).eq("type", "transfer");
      await logMovimiento({
        product_id: tx.product_id, qty: tx.qty,
        estado_anterior: "en_transito", estado_nuevo: "stock_central",
        related_user_id: tx.from_user_id,
        referencia_tipo: "transfer", referencia_id: tx.id,
        nota: "Recibido de " + (tx.from_user ? tx.from_user.name : "")
      });
      toast("✅ Recepción confirmada!", (prod ? prod.name : "") + " en tu stock", "s");
      await loadData(me.id, me.role);
    } catch(e) {
      toast("Error al confirmar", e.message, "e");
    }
  }

  // ── CANCEL TRANSFER ─────────────────────────────────────────────────────────
  async function cancelTransfer(tx) {
    var prod = tx.product || products.find(function(p){ return p.id===tx.product_id; });
    try {
      // Cancelación atómica (restituye stock al emisor + marca cancelled)
      var r = await sb.rpc("rpc_cancelar_transfer", { p_tx_id: tx.id });
      if (r.error) throw r.error;

      // Sacar de la lista local INMEDIATAMENTE (antes de loadData)
      setTransfers(function(prev){ return prev.filter(function(t){ return t.id !== tx.id; }); });
      setCancelConfirm(null);
      toast("✅ Envío cancelado", (prod ? prod.name + " volvió a tu stock" : ""), "s");

      // Notificar y ledger en background (no bloquean la UI)
      sb.from("notifications").insert({
        to_user_id: tx.to_user_id, from_name: me.name, type: "confirm",
        message: me.name + " canceló el envío de " + tx.qty + "× " + (prod ? prod.name : "producto") + ". El envío fue cancelado."
      });
      logMovimiento({
        product_id: tx.product_id, qty: tx.qty,
        estado_anterior: "en_transito", estado_nuevo: "stock_central",
        related_user_id: tx.to_user_id,
        referencia_tipo: "cancelacion", referencia_id: tx.id,
        nota: "Envío cancelado — stock restituido"
      });

      // Reload en background para sincronizar inventario
      loadData(me.id, me.role);
    } catch(e) {
      setCancelConfirm(null);
      toast("Error al cancelar", e.message, "e");
    }
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
      if (invRow.qty_available < qty) { toast("Stock insuficiente", (invRow.products||{}).name||"", "e"); continue; }
      // 1. Descontar del stock del remitente inmediatamente
      var nuevoPropio = Math.max(0, (invRow.stock_propio != null ? invRow.stock_propio : invRow.qty_available) - qty);
      await sb.from("inventory").update({
        qty_available: invRow.qty_available - qty,
        stock_propio:  nuevoPropio
      }).eq("id", invId);
      // 2. Guardar transfer como pending
      var { data: txData } = await sb.from("transfers").insert({
        from_user_id: me.id, to_user_id: sendTo,
        product_id: invRow.product_id, qty: qty,
        status: "pending", inventory_id: invId
      }).select().single();
      await logMovimiento({
        product_id: invRow.product_id, qty: -qty,
        estado_anterior: "stock_central", estado_nuevo: "en_transito",
        related_user_id: sendTo,
        referencia_tipo: "transfer", referencia_id: txData ? txData.id : null,
        nota: "Enviado a " + (contacts.find(function(c){return c.id===sendTo;})||{}).name
      });
    }
    await sb.from("notifications").insert({to_user_id:sendTo,from_name:me.name,type:"transfer",message:me.name+" te envió "+entries.length+" producto(s). Confirmá la recepción!"});
    toast("Envío completado!",entries.length+" producto(s) a "+(tUser?tUser.name:""),"s");
    setSendCart({}); setSendStep(1);
    await loadData(me.id,me.role);
  }

  // ── CATALOG ABM (superadmin only) ────────────────────────────────────────────
  // ── CARGA RÁPIDA INLINE ──────────────────────────────────────────────────────
  async function doQuickLoadInline(item, qty) {
    if (qty < 1) return;
    try {
      var r = await sb.rpc("rpc_cargar_stock", { p_inv_id: item.id, p_qty: qty });
      if (r.error) throw r.error;
      if (r.data) setInventory(function(prev){ return prev.map(function(i){ return i.id===item.id ? r.data : i; }); });
      // Ledger
      await logMovimiento({
        product_id: item.product_id, qty: qty,
        estado_anterior: "stock_central", estado_nuevo: "stock_central",
        referencia_tipo: "carga",
        nota: "Carga rápida +" + qty + " u."
      });
      var p = item.products || products.find(function(x){ return x.id===item.product_id; });
      toast("✅ Stock cargado", "+" + qty + " u. de " + (p ? p.name : ""), "s");
      setQuickLoadId(null);
      setQuickLoadQty(1);
    } catch(e) { toast("Error", e.message, "e"); }
  }

  // ── IMAGE HELPERS ──────────────────────────────────────────────────────────────
  async function loadProdImages(pid) {
    try {
      const { data } = await sb.from("product_images").select("*").eq("product_id",pid).order("orden",{ascending:true});
      setProdImages(function(prev){ return Object.assign({},prev,{[pid]:data||[]}); });
      return data||[];
    } catch(e) { return []; }
  }
  async function uploadProdImage(pid, file) {
    if (!file) return;
    if (file.size>3000000){ toast("Imagen muy grande","Máx 3MB","e"); return; }
    setImgUploading(true);
    try {
      const ext  = file.name.split(".").pop();
      const path = "products/"+pid+"/"+Date.now()+"."+ext;
      var upPromise = sb.storage.from("product-images").upload(path, file, {upsert:true});
      var upTimeout = new Promise(function(_, rej){ setTimeout(function(){ rej(new Error("La subida tardó demasiado. Revisá los permisos del bucket.")); }, 20000); });
      const { error:upErr } = await Promise.race([upPromise, upTimeout]);
      if (upErr) throw upErr;
      const { data:urlData } = sb.storage.from("product-images").getPublicUrl(path);
      const existentes = prodImages[pid]||[];
      await sb.from("product_images").insert({
        product_id:pid, url:urlData.publicUrl, storage_path:path,
        orden:existentes.length, es_principal:existentes.length===0
      });
      if (existentes.length===0){
        await sb.from("products").update({photo_url:urlData.publicUrl}).eq("id",pid);
        setProducts(function(prev){ return prev.map(function(p){ return p.id===pid?Object.assign({},p,{photo_url:urlData.publicUrl}):p; }); });
      }
      await loadProdImages(pid);
      toast("✅ Imagen cargada","","s");
    } catch(e){ toast("Error al subir",e.message,"e"); }
    finally{ setImgUploading(false); }
  }
  async function deleteProdImage(img, pid) {
    try {
      if (img.storage_path) await sb.storage.from("product-images").remove([img.storage_path]);
      await sb.from("product_images").delete().eq("id",img.id);
      const resto=(prodImages[pid]||[]).filter(function(x){ return x.id!==img.id; });
      if (img.es_principal && resto.length>0){
        await sb.from("product_images").update({es_principal:true}).eq("id",resto[0].id);
        await sb.from("products").update({photo_url:resto[0].url}).eq("id",pid);
        setProducts(function(prev){ return prev.map(function(p){ return p.id===pid?Object.assign({},p,{photo_url:resto[0].url}):p; }); });
      } else if (img.es_principal){
        await sb.from("products").update({photo_url:null}).eq("id",pid);
        setProducts(function(prev){ return prev.map(function(p){ return p.id===pid?Object.assign({},p,{photo_url:null}):p; }); });
      }
      await loadProdImages(pid);
    } catch(e){ toast("Error",e.message,"e"); }
  }
  async function setPrincipalImage(img, pid) {
    try {
      await sb.from("product_images").update({es_principal:false}).eq("product_id",pid);
      await sb.from("product_images").update({es_principal:true}).eq("id",img.id);
      await sb.from("products").update({photo_url:img.url}).eq("id",pid);
      setProducts(function(prev){ return prev.map(function(p){ return p.id===pid?Object.assign({},p,{photo_url:img.url}):p; }); });
      await loadProdImages(pid);
      toast("✅ Principal actualizada","","s");
    } catch(e){ toast("Error",e.message,"e"); }
  }
  async function moverImagen(pid, fromIdx, toIdx) {
    const imgs=[...(prodImages[pid]||[])];
    if (fromIdx<0||toIdx<0||fromIdx>=imgs.length||toIdx>=imgs.length) return;
    const [m]=imgs.splice(fromIdx,1); imgs.splice(toIdx,0,m);
    try {
      await Promise.all(imgs.map(function(img,i){ return sb.from("product_images").update({orden:i}).eq("id",img.id); }));
      await loadProdImages(pid);
    } catch(e){ toast("Error",e.message,"e"); }
  }

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
      if (qty>0){ await sb.from("inventory").insert({user_id:me.id,product_id:newProd.id,qty_available:qty,qty_sold:0,source:'own'}); }
      toast("Producto creado!",fName.trim(),"s");
    }
    setFSku(""); setFName(""); setFPrice(""); setFEmoji("✨"); setFStock("0"); setFPhoto(null);
    await loadData(me.id,me.role);
  }

  function startEdit(p){ setEditP(p); setFSku(p.sku); setFName(p.name); setFPrice(String(p.price)); setFCat(p.category||"General"); setFEmoji(p.emoji||"✨"); setFPhoto(null); setTab("catalog"); loadProdImages(p.id); }
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
      // Always upsert into single inventory row (UNIQUE constraint user_id+product_id)
      const existing = inventory.find(function(i){ return i.product_id===qlPid; });
      if (existing) {
        await logMovimiento({
          product_id: qlPid, qty: qlQty,
          estado_anterior: "stock_central", estado_nuevo: "stock_central",
          referencia_tipo: "carga",
          nota: "Carga de stock — +" + qlQty + " unidades"
        });
        const upd = await sb.from("inventory").update({
          qty_available: existing.qty_available + qlQty,
          stock_propio:  (existing.stock_propio || 0) + qlQty,
          source: "own"
        }).eq("id", existing.id).select("*, products(*)").single();
        if (upd.data) setInventory(function(p){ return p.map(function(i){ return i.id===existing.id?upd.data:i; }); });
        if (upd.error){ toast("Error",""+upd.error.message,"e"); return; }
      } else {
        await logMovimiento({
          product_id: qlPid, qty: qlQty,
          estado_anterior: null, estado_nuevo: "stock_central",
          referencia_tipo: "carga",
          nota: "Stock inicial — " + qlQty + " unidades"
        });
        const ins = await sb.from("inventory").insert({
          user_id:me.id, product_id:qlPid,
          qty_available:qlQty, qty_sold:0,
          stock_propio:qlQty, stock_recibido:0, source:"own"
        }).select("*, products(*)").single();
        if (ins.error){ toast("Error",""+ins.error.message,"e"); return; }
        if (ins.data) setInventory(function(p){ return [...p, ins.data]; });
      }
      const p = products.find(function(x){ return x.id===qlPid; });
      toast("Stock cargado!", "+" + qlQty + " u. de " + (p?p.name:""), "s");
      setQlQty(1); setQlPid(""); setQlSrch("");
    } else {
      if (!qlSku.trim()||!qlName.trim()||!qlPrice){ toast("Completa todos los campos","","e"); return; }
      const skuC = qlSku.trim().toUpperCase();
      if (products.some(function(p){ return p.sku===skuC; })){ toast("SKU ya existe","","e"); return; }
      const pr = await sb.from("products")
        .insert({sku:skuC, name:qlName.trim(), price:parseFloat(qlPrice)||0, emoji:qlEmoji, photo_url:qlPhoto||null, category:qlCat, created_by:me.id})
        .select().single();
      if (pr.error){ toast("Error",""+pr.error.message,"e"); return; }
      setProducts(function(p){ return [...p, pr.data]; });
      if (qlQty>0){
        const ii = await sb.from("inventory").insert({
          user_id:me.id, product_id:pr.data.id,
          qty_available:qlQty, qty_sold:0,
          stock_propio:qlQty, stock_recibido:0, source:"own"
        }).select("*, products(*)").single();
        if (ii.data) setInventory(function(p){ return [...p, ii.data]; });
      }
      toast("Producto creado!", qlName.trim()+" con "+qlQty+" u.", "s");
      setQlSku(""); setQlName(""); setQlPrice(""); setQlQty(1); setQlPhoto(null); setQlMode("add");
    }
  }

  // ── BULK ACTIONS ─────────────────────────────────────────────────────────────
  var bulkCount = Object.keys(bulkSel).length;

  function toggleBulkSel(id) {
    setBulkSel(function(p){ var n=Object.assign({},p); if(n[id]) delete n[id]; else n[id]=true; return n; });
  }

  function selectAllBulk() {
    var all={};
    catFiltered.forEach(function(p){ all[p.id]=true; });
    setBulkSel(all);
  }

  async function doBulkDeactivate() {
    var ids = Object.keys(bulkSel);
    for (var i=0;i<ids.length;i++) {
      await sb.from("products").update({is_active:false}).eq("id",ids[i]);
    }
    setProducts(function(p){ return p.map(function(x){ return bulkSel[x.id]?Object.assign({},x,{is_active:false}):x; }); });
    toast("Desactivados",ids.length+" productos desactivados","i");
    setBulkSel({});
  }

  async function doBulkActivate() {
    var ids = Object.keys(bulkSel);
    for (var i=0;i<ids.length;i++) {
      await sb.from("products").update({is_active:true}).eq("id",ids[i]);
    }
    setProducts(function(p){ return p.map(function(x){ return bulkSel[x.id]?Object.assign({},x,{is_active:true}):x; }); });
    toast("Activados",ids.length+" productos activados","s");
    setBulkSel({});
  }

  // ── BULK PRICE UPDATE (admin) ────────────────────────────────────────────────
  async function doBulkPriceUpdate(e) {
    var file = e.target.files && e.target.files[0];
    if (!file) return;
    var ext = file.name.split(".").pop().toLowerCase();
    var reader = new FileReader();
    reader.onload = async function(ev) {
      try {
        var rows = [];
        if (ext==="xlsx"||ext==="xls") {
          var wb = XLSX.read(ev.target.result, {type:"binary"});
          var ws = wb.Sheets[wb.SheetNames[0]];
          var data = XLSX.utils.sheet_to_json(ws, {header:1, raw:false});
          var cs=-1, cp=-1, sr=0;
          if (data.length>0) {
            var h=data[0].map(function(x){ return (x||"").toString().toLowerCase(); });
            h.forEach(function(x,i){ if(/sku|cod/.test(x)) cs=i; else if(/prec|price|val/.test(x)) cp=i; });
            if (cs>=0&&cp>=0) sr=1; else { cs=0; cp=1; sr=1; }
          }
          data.slice(sr).forEach(function(row){
            if (!row||!row.length) return;
            var sku=(row[cs]||"").toString().trim().toUpperCase();
            var priceRaw=(row[cp]||"0").toString().replace(/[^\d.,]/g,"");
            if (priceRaw.includes(",")) priceRaw=priceRaw.replace(/\./g,"").replace(",",".");
            var price=parseFloat(priceRaw)||0;
            if (sku&&price>0) rows.push({sku:sku,price:price});
          });
        } else {
          ev.target.result.split("\n").forEach(function(line){
            var p=line.trim().split(/[,\t]/);
            if (p.length>=2) {
              var sku=p[0].trim().toUpperCase();
              var priceRaw=p[1].replace(/[^\d.,]/g,"");
              if (priceRaw.includes(",")) priceRaw=priceRaw.replace(/\./g,"").replace(",",".");
              var price=parseFloat(priceRaw)||0;
              if (sku&&price>0) rows.push({sku:sku,price:price});
            }
          });
        }
        var updated=0;
        for (var i=0;i<rows.length;i++) {
          var r=rows[i];
          var res=await sb.from("products").update({price:r.price,updated_at:new Date().toISOString()}).eq("sku",r.sku);
          if (!res.error) updated++;
        }
        toast("Precios actualizados!",updated+" productos actualizados","s");
        await loadData(me.id,me.role);
      } catch(err){ toast("Error al procesar archivo",""+err.message,"e"); }
    };
    if (ext==="xlsx"||ext==="xls") reader.readAsBinaryString(file); else reader.readAsText(file);
  }

  // ── TOGGLE PRODUCT ACTIVE/INACTIVE ──────────────────────────────────────────
  async function toggleProduct(prod) {
    var newVal = !prod.is_active;
    await sb.from("products").update({is_active:newVal}).eq("id",prod.id);
    setProducts(function(p){ return p.map(function(x){ return x.id===prod.id?Object.assign({},x,{is_active:newVal}):x; }); });
    toast(newVal?"Producto activado":"Producto desactivado","","i");
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
    var valid = impRows.filter(function(r){ return r.ok; });
    if (!valid.length){ toast("Sin filas válidas","","e"); return; }

    toast("Importando...","Procesando "+valid.length+" productos. No cierres la página.","i");

    var BATCH = 50; // insert 50 at a time
    var totalInserted = 0;
    var totalUpdated  = 0;

    // Separate new vs existing
    var toInsert = [];
    var toUpdate = [];
    valid.forEach(function(row){
      var ex = products.find(function(p){ return p.sku===row.sku; });
      if (ex) { toUpdate.push({id:ex.id, price:row.price, name:row.name, category:row.cat}); }
      else     { toInsert.push({sku:row.sku, name:row.name, price:row.price, emoji:"✨", category:row.cat, created_by:me.id, is_active:true}); }
    });

    // Batch INSERT new products
    var insertedProds = [];
    for (var i=0; i<toInsert.length; i+=BATCH) {
      var batch = toInsert.slice(i, i+BATCH);
      var res = await sb.from("products").insert(batch).select();
      if (res.data) { insertedProds = insertedProds.concat(res.data); totalInserted += res.data.length; }
      if (res.error) { console.error("Insert batch error:", res.error.message); }
      // Small delay to avoid rate limiting
      await new Promise(function(r){ setTimeout(r, 100); });
    }

    // Batch UPDATE existing prices (upsert by sku)
    for (var i=0; i<toUpdate.length; i+=BATCH) {
      var batch = toUpdate.slice(i, i+BATCH);
      for (var j=0; j<batch.length; j++) {
        await sb.from("products").update({price:batch[j].price, name:batch[j].name, category:batch[j].category, updated_at:new Date().toISOString()}).eq("id",batch[j].id);
      }
      totalUpdated += batch.length;
      await new Promise(function(r){ setTimeout(r, 80); });
    }

    if (insertedProds.length>0) setProducts(function(p){ return [...p,...insertedProds]; });

    var msg = "";
    if (totalInserted>0) msg += totalInserted+" nuevos";
    if (totalUpdated>0)  msg += (msg?", ":"")+totalUpdated+" actualizados";
    toast("Importación completada!", msg, "s");
    setImpRows([]); setImpTxt(""); setImpFile(null); setTab("catalog");
    await loadData(me.id, me.role);
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

  // ── MOVIMIENTO DE STOCK ──────────────────────────────────────────────────────
  async function doMovimiento() {
    var item=movModal, prod=item.products||products.find(function(p){return p.id===item.product_id;});
    var newQty=item.qty_available;
    if(movType==="entrada") newQty=item.qty_available+movQty;
    if(movType==="salida")  newQty=Math.max(0,item.qty_available-movQty);
    if(movType==="ajuste")  newQty=movQty;
    if(newQty<0){toast("Stock no puede ser negativo","","e");return;}
    var r=await sb.from("inventory").update({qty_available:newQty}).eq("id",item.id).select("*, products(*)").single();
    if(r.error){toast("Error",""+r.error.message,"e");return;}
    setInventory(function(p){return p.map(function(i){return i.id===item.id?r.data:i;});});
    var entry={t:new Date().toLocaleString("es-AR",{day:"2-digit",month:"2-digit",hour:"2-digit",minute:"2-digit"}),prod:prod?prod.name:"",sku:prod?prod.sku:"",type:movType,before:item.qty_available,after:newQty,diff:newQty-item.qty_available,note:movNote||""};
    setMovHistory(function(p){return [entry,...p.slice(0,99)];});
    if (movType==="salida") {
      await sb.from("sale_logs").insert({user_id:me.id,product_id:item.product_id,qty:Math.abs(newQty-item.qty_available),sale_price:prod?parseFloat(prod.price||0):0,source:"own_stock"});
    }
    var icons={entrada:"📥",salida:"📤",ajuste:"🔧"};
    toast(icons[movType]+" "+movType.charAt(0).toUpperCase()+movType.slice(1),(prod?prod.name:"")+": "+item.qty_available+" → "+newQty+" u.","s");
    setMovModal(null);setMovNote("");setMovQty(1);
  }

  // ── EDITAR STOCK MANUALMENTE ─────────────────────────────────────────────────
  async function doEditStock() {
    if (editQty < 0) { toast("La cantidad no puede ser negativa","","e"); return; }
    var r = await sb.from("inventory").update({qty_available: editQty}).eq("id", editStock.id).select("*, products(*)").single();
    if (r.error) { toast("Error",""+r.error.message,"e"); return; }
    setInventory(function(p){ return p.map(function(i){ return i.id===editStock.id ? r.data : i; }); });
    var prod = editStock.products || products.find(function(x){ return x.id===editStock.product_id; });
    toast("Stock actualizado", (prod?prod.name:"")+" → "+editQty+" u.", "s");
    setEditStock(null);
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

  var catAll = isAdmin && showInactive ? products.concat([]) : products.filter(function(p){ return p.is_active!==false; });
  var catFiltered = catAll.filter(function(p){ var q=srchCat.toLowerCase(); return !q||p.name.toLowerCase().includes(q)||p.sku.toLowerCase().includes(q)||(p.category||"").toLowerCase().includes(q); });
  var conFiltered = contacts.filter(function(c){ var q=srchCon.toLowerCase(); return !q||c.name.toLowerCase().includes(q)||c.email.toLowerCase().includes(q); });
  var logFiltered = logs.filter(function(l){ var q=srchLog.toLowerCase(); var p=l.product; return !q||(p&&p.name.toLowerCase().includes(q))||l.source.includes(q); });
  var sendFiltered = myStock.filter(function(i){ var p=i.products||products.find(function(x){ return x.id===i.product_id; }); if(!p) return false; var q=sendSrch.toLowerCase(); return !q||p.name.toLowerCase().includes(q)||p.sku.toLowerCase().includes(q); });
  var qlFiltered = products.filter(function(p){ var q=qlSrch.toLowerCase(); return !q||p.name.toLowerCase().includes(q)||p.sku.toLowerCase().includes(q); });
  var pendingTx = transfers.filter(function(t){ return t.to_user_id===me?.id&&t.status==="pending"; });
  var sentTx = transfers.filter(function(t){ return t.from_user_id===me?.id&&t.status==="pending"; }).slice(0,10);
  var myNotifs = notifs.filter(function(n){ return n.to_user_id===me?.id; });

  // ── LOADING SCREEN ────────────────────────────────────────────────────────────
  if (loading) return (
    <div>
      <style>{CSS}</style>
      <div className="loading-screen">
        <img src={LOGO_URL} alt="Venta Directa" onError={function(e){e.target.style.display="none";e.target.nextSibling.style.display="block";}} style={{width:220,marginBottom:4,filter:"drop-shadow(0 8px 24px rgba(0,0,0,.45)) drop-shadow(0 2px 8px rgba(0,0,0,.35))"}} /><div style={{display:"none",fontSize:32,fontWeight:900,color:"#fff",letterSpacing:"-.02em"}}>Venta Directa</div>
        <div className="loading-spinner"/>
        <div style={{fontSize:12,color:"rgba(255,255,255,.7)",fontWeight:600,letterSpacing:".05em"}}>Conectando con Supabase...</div>
      </div>
    </div>
  );

  // ── LOGIN SCREEN ──────────────────────────────────────────────────────────────
  if (!me && authOk) return (
    <div style={{display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",height:"100vh",background:"var(--bg,#f5f5f5)"}}>
      <style>{CSS}</style>
      <div style={{fontSize:32,marginBottom:16}}>⏳</div>
      <div style={{fontSize:16,fontWeight:700,color:"#e0224e"}}>Cargando tu cuenta...</div>
    </div>
  );

  if (!me) return (
    <div>
      <style>{CSS}</style>
      <div className="login-bg">
        <div className={"lbox"+(authShake?" shake":"")}>
          <div style={{textAlign:"center",marginBottom:20}}>
            <img src={LOGO_URL} alt="Venta Directa" onError={function(e){e.target.style.display="none";e.target.nextSibling.style.display="block";}} style={{width:200,marginBottom:4}} /><div style={{display:"none",fontSize:28,fontWeight:900,color:"var(--in)",letterSpacing:"-.02em"}}>Venta Directa</div>
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
              <button className="cta cta-in" onClick={doLogin} disabled={loggingIn}><Ic n="check" s={18}/>{loggingIn?"Entrando...":"Entrar"}</button>
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
          {/* ── Login con Google ── */}
          <div style={{display:"flex",alignItems:"center",gap:10,margin:"18px 0 14px"}}>
            <div style={{flex:1,height:1,background:"var(--brd)"}}/>
            <span style={{fontSize:11,color:"var(--t3)",fontWeight:700}}>o</span>
            <div style={{flex:1,height:1,background:"var(--brd)"}}/>
          </div>
          <button onClick={doLoginGoogle}
            style={{width:"100%",padding:"14px",borderRadius:"var(--r2)",border:"1.5px solid var(--brd)",background:"#fff",color:"var(--t1)",fontFamily:"var(--hf)",fontWeight:800,fontSize:14,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",gap:10}}>
            <svg width="20" height="20" viewBox="0 0 48 48"><path fill="#FFC107" d="M43.6 20.5h-1.9V20H24v8h11.3c-1.6 4.7-6.1 8-11.3 8-6.6 0-12-5.4-12-12s5.4-12 12-12c3.1 0 5.9 1.2 8 3.1l5.7-5.7C34.6 6.1 29.6 4 24 4 12.9 4 4 12.9 4 24s8.9 20 20 20 20-8.9 20-20c0-1.3-.1-2.3-.4-3.5z"/><path fill="#FF3D00" d="M6.3 14.7l6.6 4.8C14.7 16 19 13 24 13c3.1 0 5.9 1.2 8 3.1l5.7-5.7C34.6 6.1 29.6 4 24 4 16.3 4 9.7 8.3 6.3 14.7z"/><path fill="#4CAF50" d="M24 44c5.5 0 10.5-2.1 14.3-5.6l-6.6-5.6C29.7 34.6 27 36 24 36c-5.2 0-9.6-3.3-11.3-7.9l-6.6 5.1C9.6 39.6 16.2 44 24 44z"/><path fill="#1976D2" d="M43.6 20.5H24v8h11.3c-.8 2.2-2.2 4.1-4 5.4l6.6 5.6C41.4 36.9 44 31 44 24c0-1.3-.1-2.3-.4-3.5z"/></svg>
            Continuar con Google
          </button>
        </div>
      </div>
      <div className="toast-wrap">{toasts.map(function(t){ return <Toast key={t.id} t={t} remove={rmToast}/>; })}</div>
    </div>
  );

  // ── TABS ─────────────────────────────────────────────────────────────────────
  var pedPendCount = pedidos.filter(function(p){return p.estado==="pendiente";}).length;
  var TABS = [
    {id:"stock",     lbl:"Stock",     ico:"box"},
    {id:"cargar",    lbl:"Cargar",    ico:"plus"},
    {id:"pedidos",   lbl:"Pedidos",   ico:"list"},
    {id:"enviados",  lbl:"Enviados",  ico:"send"},
    {id:"recibidos", lbl:"Recibidos", ico:"users"},
    {id:"ventas",    lbl:"Ventas",    ico:"chart"},
    {id:"contacts",  lbl:"Red",       ico:"clock"},
  ];
  if (isAdmin) {
    TABS.splice(4, 0, {id:"catalog",  lbl:"Catálogo", ico:"list"});
    TABS.splice(5, 0, {id:"importar", lbl:"Importar", ico:"upload"});
    TABS.push({id:"admin", lbl:"Admin", ico:"shield"});
  }

  // ── MAIN APP ──────────────────────────────────────────────────────────────────
  return (
    <div>
      <style>{CSS}</style>
      <div className="app">

        {/* HEADER */}
        <div className="hdr">
          <div className="hdr-top">
            <div className="hdr-avatar" style={me.color?{background:me.color}:{}}>{(me.name||"?").trim().charAt(0).toUpperCase()}</div>
            <div style={{flex:1,minWidth:0,cursor:"pointer"}} onClick={function(){setTab("stock");}}>
              <div className="hdr-hi">{me.role==="superadmin"?"👑 Administrador":("¡Hola, "+me.name.split(" ")[0]+"!")}</div>
              <div className="hdr-sub">{me.role==="superadmin"?"Panel de control":(tab==="recibidos"?"Acá ves lo que te entregaron para vender":"¿Lista para vender hoy?")}</div>
            </div>
            <div style={{display:"flex",alignItems:"center",gap:8,flexShrink:0}}>
              <div style={{position:"relative",cursor:"pointer"}} onClick={function(){setTab("contacts");}}>
                <button className="hdr-btn"><Ic n="bell" s={18}/></button>
                {totalBadge>0&&<div style={{position:"absolute",top:-4,right:-4,width:18,height:18,background:"var(--in-m)",borderRadius:"50%",fontSize:9,fontWeight:900,color:"#fff",display:"flex",alignItems:"center",justifyContent:"center",border:"2px solid var(--card)"}}>{totalBadge}</div>}
              </div>
              <button className="hdr-btn" onClick={doLogout}><Ic n="logout" s={18}/></button>
            </div>
          </div>
          {(tab==="stock"||tab==="inicio")&&(
          <HeaderSearch value={srchStock} onChange={setSrchStock} placeholder="Buscar productos, marcas o categorías..."/>
          )}
        </div>

        <div className="main">

          {/* ══ STOCK ══ */}
          {(tab==="stock"||tab==="inicio")&&(function(){
            // Separar por source: 'own' = stock propio, 'consigna' = recibido de otro usuario
            const ownStock      = inventory.filter(function(i){
              return i.user_id===me.id && (i.source==="own" || !i.source) && i.qty_available>0;
            });
            const consignStock  = inventory.filter(function(i){
              return i.user_id===me.id && i.source==="consigna" && i.qty_available>0;
            });

            const q = srchStock.toLowerCase();
            function filtrar(list){ return list.filter(function(i){
              const p=i.products||products.find(function(x){return x.id===i.product_id;});
              if(!p) return false;
              return !q||p.name.toLowerCase().includes(q)||p.sku.toLowerCase().includes(q)||(p.category||"").toLowerCase().includes(q);
            });}
            const ownFilt     = filtrar(ownStock);
            const consignFilt  = filtrar(consignStock);

            const totalVal = ownStock.reduce(function(s,i){
              const p=i.products||products.find(function(x){return x.id===i.product_id;});
              return s+(p?parseFloat(p.price||0)*i.qty_available:0);
            },0);
            const consignaEnv = consignActivas.filter(function(c){ return c.owner_id===me.id || c.vendedora_id===me.id; }).length;
            // Ventas del mes actual (suma de sale_logs del mes)
            var _now = new Date();
            const ventasMes = logs.filter(function(l){
              var d=new Date(l.created_at);
              return d.getMonth()===_now.getMonth() && d.getFullYear()===_now.getFullYear();
            }).reduce(function(s,l){ return s + (parseFloat(l.sale_price||0) * (l.qty||1)); }, 0);

            function StockTable(list, isConsigna){
              if(list.length===0) return <div className="empty" style={{padding:"16px"}}>Sin productos.</div>;
              return (
                <div>
                  {list.map(function(item){
                    const p=item.products||products.find(function(x){return x.id===item.product_id;});
                    if(!p) return null;
                    const senderTx = isConsigna ? transfers.find(function(t){
                      return t.to_user_id===me.id && t.product_id===item.product_id && t.status==="confirmed";
                    }) : null;
                    const sender = senderTx ? senderTx.from_user : null;
                    // Colores por tipo
                    var accentBg  = isConsigna ? "var(--am-l)"  : "var(--em-l)";
                    var accentCol = isConsigna ? "var(--am-d)"  : "var(--em-d)";
                    var tagBg     = isConsigna ? "#fff3e0"      : "#e8f5e9";
                    var tagCol    = isConsigna ? "#e65100"      : "#2e7d32";
                    var tagTxt    = isConsigna ? ("📨 De "+( sender?sender.name:"otro vendedor")) : "📦 Stock propio";
                    var borderCol = isConsigna ? "rgba(255,122,0,.3)" : "rgba(0,184,122,.2)";
                    // Calcular propio vs recibido
                    var stockPropio   = item.stock_propio   != null ? item.stock_propio   : (isConsigna ? 0 : item.qty_available);
                    var stockRecibido = item.stock_recibido != null ? item.stock_recibido : (isConsigna ? item.qty_available : 0);
                    var stockTotal    = item.qty_available;
                    return (
                      <div key={item.id} style={{
                        background:"var(--card)",
                        borderRadius:14,
                        border:"1.5px solid "+borderCol,
                        marginBottom:10,
                        overflow:"hidden",
                        boxShadow:"var(--sh)"
                      }}>
                        {/* Etiqueta de tipo */}
                        <div style={{background:tagBg,padding:"5px 12px",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
                          <span style={{fontSize:10,fontWeight:800,color:tagCol,letterSpacing:".02em"}}>{tagTxt}</span>
                          <span style={{fontFamily:"var(--mf)",fontSize:10,fontWeight:700,color:"var(--t3)"}}>{p.sku}</span>
                        </div>
                        {/* Cuerpo */}
                        <div style={{display:"flex",alignItems:"center",gap:12,padding:"11px 13px"}}>
                          <ProdThumb prod={p} size={46}/>
                          <div style={{flex:1,minWidth:0}}>
                            <div style={{fontWeight:800,fontSize:13,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",color:"var(--t1)"}}>{p.name}</div>
                            <div style={{fontSize:11,color:"var(--t3)",marginTop:2}}>{p.category}</div>
                            <div style={{fontSize:13,fontWeight:800,color:accentCol,marginTop:3,fontFamily:"var(--mf)"}}>{fmtARS(p.price)}</div>
                          </div>
                          {/* Total disponible */}
                          <div style={{textAlign:"center",flexShrink:0,background:accentBg,borderRadius:12,padding:"8px 14px",minWidth:52}}>
                            <div style={{fontFamily:"var(--mf)",fontWeight:900,fontSize:24,color:accentCol,lineHeight:1}}>{stockTotal}</div>
                            <div style={{fontSize:9,fontWeight:700,color:accentCol,textTransform:"uppercase",letterSpacing:".05em",marginTop:2}}>total</div>
                          </div>
                        </div>
                        {/* Desglose propio / recibido */}
                        {(stockPropio>0||stockRecibido>0)&&(
                          <div style={{display:"flex",background:"var(--bg)",margin:"0 12px 10px",borderRadius:10,overflow:"hidden",border:"1px solid var(--brd)"}}>
                            <div style={{flex:1,textAlign:"center",padding:"8px 6px",borderRight:"1px solid var(--brd)"}}>
                              <div style={{fontFamily:"var(--mf)",fontWeight:800,fontSize:16,color:"var(--em-d)"}}>{stockPropio}</div>
                              <div style={{fontSize:10,color:"var(--t3)",fontWeight:700,marginTop:1}}>📦 Propio</div>
                            </div>
                            <div style={{flex:1,textAlign:"center",padding:"8px 6px"}}>
                              <div style={{fontFamily:"var(--mf)",fontWeight:800,fontSize:16,color:"var(--bl-d)"}}>{stockRecibido}</div>
                              <div style={{fontSize:10,color:"var(--t3)",fontWeight:700,marginTop:1}}>📩 Recibido</div>
                            </div>
                          </div>
                        )}
                        {/* Acciones */}
                        <div style={{display:"flex",borderTop:"1px solid var(--brd)"}}>
                          <button className="btn btn-xs b-wa" style={{flex:1,justifyContent:"center",borderRadius:0,padding:"10px 0",border:"none",borderRight:"1px solid var(--brd)"}} onClick={function(){shareOne(p);}}><Ic n="wa" s={13}/></button>
                          <button className="btn btn-xs b-em"
                            style={{flex:1,justifyContent:"center",borderRadius:0,padding:"10px 0",border:"none",borderRight:"1px solid var(--brd)",color:"#00b87a",background: quickLoadId===item.id ? "#00b87a" : "#e8faf4"}}
                            onClick={function(){
                              if (quickLoadId===item.id) { setQuickLoadId(null); }
                              else { setQuickLoadId(item.id); setQuickLoadQty(1); }
                            }}>
                            <Ic n="plus" s={13}/><span style={{fontSize:11,color: quickLoadId===item.id ? "#fff" : "#00b87a"}}>Cargar</span>
                          </button>
                          <button className="btn btn-xs b-in" style={{flex:1,justifyContent:"center",borderRadius:0,padding:"10px 0",border:"none",borderRight:"1px solid var(--brd)",color:"var(--pri)",background:"var(--pri-l)"}} onClick={function(){setMovModal(item);setMovType("entrada");setMovQty(1);setMovNote("");}}>
                            <Ic n="plus" s={13}/><span style={{fontSize:11}}>Movimiento</span>
                          </button>
                          <button className="btn btn-xs b-em" style={{flex:1,justifyContent:"center",borderRadius:0,padding:"10px 0",border:"none"}} onClick={function(){doSell(item);}} disabled={item.qty_available===0}>
                            <Ic n="check" s={13}/><span style={{fontSize:11}}>Venta</span>
                          </button>
                        </div>
                        {/* Panel carga rápida inline */}
                        {quickLoadId===item.id&&(
                          <div style={{
                            background:"#f0fff8",
                            border:"1.5px solid #00b87a",
                            borderTop:"none",
                            borderRadius:"0 0 12px 12px",
                            padding:"12px 14px",
                            display:"flex",
                            alignItems:"center",
                            gap:10
                          }}>
                            <div style={{fontSize:12,fontWeight:700,color:"#00b87a",flexShrink:0}}>+Stock</div>
                            <div style={{display:"flex",alignItems:"center",border:"1.5px solid #b2dfdb",borderRadius:10,overflow:"hidden",flexShrink:0}}>
                              <button onClick={function(){setQuickLoadQty(function(q){return Math.max(1,q-1);});}}
                                style={{width:32,height:32,border:"none",background:"#e8faf4",cursor:"pointer",fontSize:18,fontWeight:700,color:"#00b87a",display:"flex",alignItems:"center",justifyContent:"center"}}>−</button>
                              <div style={{minWidth:36,textAlign:"center",fontFamily:"var(--mf)",fontWeight:900,fontSize:16,color:"#00b87a",padding:"0 4px"}}>{quickLoadQty}</div>
                              <button onClick={function(){setQuickLoadQty(function(q){return q+1;});}}
                                style={{width:32,height:32,border:"none",background:"#e8faf4",cursor:"pointer",fontSize:18,fontWeight:700,color:"#00b87a",display:"flex",alignItems:"center",justifyContent:"center"}}>+</button>
                            </div>
                            <button onClick={function(){doQuickLoadInline(item,quickLoadQty);}}
                              style={{flex:1,padding:"8px",borderRadius:10,border:"none",background:"#00b87a",color:"#fff",fontFamily:"var(--hf)",fontWeight:800,fontSize:13,cursor:"pointer"}}>
                              ✅ Confirmar +{quickLoadQty}
                            </button>
                            <button onClick={function(){setQuickLoadId(null);}}
                              style={{width:28,height:28,borderRadius:"50%",border:"none",background:"#e0e0e0",color:"#888",cursor:"pointer",fontSize:14,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>✕</button>
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>
              );
            }

            // ── Carousel slides (real + default) ────────────────────────────
            var activos = carousels.filter(function(c){ return c.active; });
            var slides = activos.length>0 ? activos : [
              {id:"d1",title:"¡Bienvenida!",subtitle:"Tu catálogo digital",bg_color:"#e0224e",emoji:"🛍️",link_tab:"catalog"},
              {id:"d2",title:"Compartí tus productos",subtitle:"Enviá lista por WhatsApp",bg_color:"#10b981",emoji:"📲",link_tab:""},
            ];

            const pq = pSearch.toLowerCase().trim();
            const pResults = pq.length >= 2 ? products.filter(function(p){
              return p.name.toLowerCase().includes(pq) || (p.sku||"").toLowerCase().includes(pq);
            }).slice(0, 8) : [];

            return (
              <div style={{paddingBottom:24}}>

                {/* ─ BUSCADOR RÁPIDO DE PRECIOS ─ */}
                <div className="card" style={{marginTop:10,marginBottom:14}}>
                  <div style={{padding:"12px 14px"}}>
                    <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:pResults.length>0||pq.length>=2?10:0}}>
                      <span style={{fontSize:18}}>💲</span>
                      <input
                        value={pSearch}
                        onChange={function(e){setPSearch(e.target.value);}}
                        placeholder="Consultar precio... (nombre o SKU)"
                        style={{flex:1,border:"1.5px solid var(--brd)",borderRadius:10,padding:"10px 12px",fontSize:14,outline:"none",fontFamily:"inherit"}}
                      />
                      {pSearch&&<button onClick={function(){setPSearch("");}} style={{background:"none",border:"none",fontSize:16,cursor:"pointer",color:"var(--t3)"}}>✕</button>}
                    </div>
                    {pq.length>=2&&pResults.length===0&&(
                      <div style={{fontSize:12,color:"var(--t3)",textAlign:"center",padding:"8px 0"}}>Sin resultados para "{pSearch}"</div>
                    )}
                    {pResults.map(function(p){
                      return (
                        <div key={p.id} style={{display:"flex",alignItems:"center",gap:10,padding:"8px 4px",borderTop:"1px solid var(--brd)"}}>
                          <span style={{fontSize:20,width:28,textAlign:"center"}}>{p.photo_url?<img src={p.photo_url} style={{width:28,height:28,borderRadius:6,objectFit:"cover"}}/>:p.emoji||"📦"}</span>
                          <div style={{flex:1,minWidth:0}}>
                            <div style={{fontSize:13,fontWeight:700,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.name}</div>
                            <div style={{fontSize:10,color:"var(--t3)"}}><span style={{fontFamily:"var(--mf)",fontWeight:800,color:"var(--pri)",background:"var(--pri-l)",borderRadius:4,padding:"1px 5px",marginRight:5}}>{p.sku||"s/c"}</span>{p.category||""}</div>
                          </div>
                          <div style={{fontFamily:"var(--mf)",fontWeight:800,fontSize:16,color:"var(--em-d,#00875a)",whiteSpace:"nowrap"}}>${parseFloat(p.price||0).toLocaleString("es-AR")}</div>
                        </div>
                      );
                    })}
                  </div>
                </div>

                {/* ─ CARRUSEL DE OFERTAS ─ */}
                {slides.length>0&&(
                  <div style={{marginBottom:18,paddingTop:10}}>
                    <div className="carousel-wrap" id="main-carousel"
                      onScroll={function(e){
                        var w=e.target.scrollWidth/slides.length;
                        setOfferSlide(Math.round(e.target.scrollLeft/w));
                      }}>
                      {slides.map(function(sl,i){
                        return (
                          <div key={sl.id} className="carousel-slide"
                            style={{width:"calc(100vw - 44px)",height:150,background:sl.bg_color||"#e0224e",cursor:sl.link_tab?"pointer":"default",position:"relative",overflow:"hidden"}}
                            onClick={function(){if(sl.link_tab)setTab(sl.link_tab);}}>
                            <div style={{display:"flex",alignItems:"center",height:"100%",padding:"16px 20px",gap:16,position:"relative",zIndex:2}}>
                              {!sl.image_url&&<div style={{fontSize:52,flexShrink:0,filter:"drop-shadow(0 4px 8px rgba(0,0,0,.2))"}}>{sl.emoji||"🔥"}</div>}
                              <div style={{flex:1,minWidth:0}}>
                                <div style={{fontSize:20,fontWeight:900,color:"#fff",lineHeight:1.15,marginBottom:5,textShadow:sl.image_url?"0 1px 6px rgba(0,0,0,.4)":"none"}}>{sl.title}</div>
                                {sl.subtitle&&<div style={{fontSize:12,color:"rgba(255,255,255,.92)",fontWeight:600,textShadow:sl.image_url?"0 1px 4px rgba(0,0,0,.4)":"none"}}>{sl.subtitle}</div>}
                                {sl.link_tab&&<div style={{marginTop:8,display:"inline-flex",alignItems:"center",gap:5,background:"rgba(255,255,255,.95)",borderRadius:20,padding:"5px 14px",fontSize:11,color:sl.bg_color||"#e0224e",fontWeight:800}}>Ver producto →</div>}
                              </div>
                              {sl.image_url&&<img src={sl.image_url} alt="" style={{height:"118%",maxWidth:"46%",objectFit:"contain",flexShrink:0,filter:"drop-shadow(0 8px 16px rgba(0,0,0,.3))",marginRight:-6}}/>}
                            </div>
                            {sl.image_url&&<div style={{position:"absolute",inset:0,background:"linear-gradient(90deg,"+(sl.bg_color||"#e0224e")+" 35%,transparent 100%)",zIndex:1}}/>}
                          </div>
                        );
                      })}
                    </div>
                    {slides.length>1&&(
                      <div className="carousel-dots">
                        {slides.map(function(_,i){return <div key={i} className={"carousel-dot"+(i===offerSlide?" on":"")} onClick={function(){setOfferSlide(i);var el=document.getElementById("main-carousel");if(el){var w=el.scrollWidth/slides.length;el.scrollTo({left:w*i,behavior:"smooth"});}}}/>;})}
                      </div>
                    )}
                  </div>
                )}

                {/* ─ MÉTRICAS SUPERIORES ─ */}
                <div className="mtop-grid">
                  <div className="mtop-card" onClick={function(){setTab("ventas");}} style={{cursor:"pointer"}}>
                    <div className="mtop-ico" style={{background:"var(--in-l)",color:"var(--in)"}}><Ic n="chart" s={18}/></div>
                    <div>
                      <div className="mtop-lbl">Ventas del mes</div>
                      <div className="mtop-val">{fmtARS(ventasMes).replace("$ ","$").replace(",00","")}</div>
                    </div>
                  </div>
                  <div className="mtop-card" onClick={function(){if(isAdmin)setTab("catalog");}} style={{cursor:isAdmin?"pointer":"default"}}>
                    <div className="mtop-ico" style={{background:"var(--pu-l,#f3eafe)",color:"var(--pu)"}}><Ic n="box" s={18}/></div>
                    <div>
                      <div className="mtop-lbl">Productos activos</div>
                      <div className="mtop-val">{products.filter(function(p){return p.is_active!==false;}).length}</div>
                    </div>
                  </div>
                  <div className="mtop-card" onClick={function(){setTab("pedidos");}} style={{cursor:"pointer"}}>
                    <div className="mtop-ico" style={{background:"var(--em-l)",color:"var(--em-d)"}}><Ic n="list" s={18}/></div>
                    <div>
                      <div className="mtop-lbl">Pedidos pendientes</div>
                      <div className="mtop-val">{pedPendCount}</div>
                      {pedPendCount>0&&<div className="mtop-foot" style={{color:"var(--in)"}}>Ver pedidos →</div>}
                    </div>
                  </div>
                </div>

                {/* ─ SECCIONES FAVORITAS ─ */}
                <div style={{marginBottom:18}}>
                  <div className="sec-hdr"><div className="sec-hdr-t">Tus secciones</div><span className="sec-hdr-a" onClick={function(){setShareM(true);setShareSel({});}}>📲 WhatsApp</span></div>
                  <div className="fav-grid">
                    {(function(){
                      var favs = [
                        {ico:"📦",bg:"var(--in-l)",col:"var(--in)",lbl:"Mi Stock",sub:ownStock.reduce(function(s,i){return s+i.qty_available;},0)+" productos",tab:"stock",scroll:true},
                        {ico:"🤝",bg:"var(--bl-l)",col:"var(--bl)",lbl:"Consigna",sub:consignaEnv>0?consignaEnv+" activos":"Enviados/Recibidos",tab:"consigna",badge:consignaEnv||null},
                        {ico:"📋",bg:"#fff3e6",col:"#e06a00",lbl:"Pedidos",sub:pedPendCount>0?pedPendCount+" pendientes":"Ver pedidos",tab:"pedidos",badge:pedPendCount||null},
                        {ico:"💰",bg:"var(--em-l)",col:"var(--em-d)",lbl:"Ventas",sub:fmtARS(totalVal).replace("$ ","$"),tab:"ventas"},
                      ];
                      if (me.role==="reseller"||me.role==="lider"||me.role==="empresaria") {
                        var peProp = pedEspList.filter(function(p){
                          if (me.role==="reseller")   return p.vendedor_id===me.id && ["borrador","listo_entregar"].includes(p.estado);
                          if (me.role==="lider")      return p.lider_id===me.id && p.estado==="pendiente_lider";
                          if (me.role==="empresaria") return p.empresa_id===me.id && ["pendiente_empresaria","aprobado","enviado_proveedor","recibido"].includes(p.estado);
                          return false;
                        }).length;
                        favs.push({ico:"🛍️",bg:"#fdf0f7",col:"#c2185b",lbl:"Pedido Especial",sub:peProp>0?peProp+" por revisar":"Productos por encargo",tab:"pedesp",badge:peProp||null});
                      }
                      return favs;
                    })().map(function(f,i){
                      return (
                        <div key={i} className="fav-card" onClick={function(){setTab(f.tab);}} style={{position:"relative"}}>
                          <div className="fav-card-ico" style={{background:f.bg,color:f.col}}>{f.ico}</div>
                          <div>
                            <div className="fav-card-lbl">{f.lbl}</div>
                            <div className="fav-card-sub">{f.sub}</div>
                          </div>
                          {f.badge&&<div style={{position:"absolute",top:8,right:8,background:"var(--cr)",color:"#fff",borderRadius:20,minWidth:18,height:18,display:"flex",alignItems:"center",justifyContent:"center",fontSize:9,fontWeight:900,padding:"0 4px"}}>{f.badge}</div>}
                        </div>
                      );
                    })}
                  </div>
                </div>

                {/* ─ ENVÍOS PENDIENTES: LOS QUE YO MANDÉ ─ */}
                {sentTx.length>0&&(
                  <div style={{margin:"0 14px 16px"}}>
                    <div style={{background:"var(--bl-l)",border:"1.5px solid rgba(0,150,199,.25)",borderRadius:14,padding:"12px 14px"}}>
                      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:8}}>
                        <div style={{fontWeight:800,fontSize:13,color:"var(--bl-d)"}}>
                          📤 {sentTx.length} en camino — sin confirmar
                        </div>
                        <div style={{fontSize:10,color:"var(--bl-d)",fontWeight:600,opacity:.7}}>
                          El stock vuelve si cancelás
                        </div>
                      </div>
                      {sentTx.map(function(tx){
                        const p=tx.product;
                        const dest=tx.to_user;
                        const confirmingCancel = cancelConfirm===tx.id;
                        return (
                          <div key={tx.id} style={{background:"var(--card)",borderRadius:12,marginBottom:6,border:"1px solid var(--brd)",overflow:"hidden"}}>
                            <div style={{display:"flex",alignItems:"center",gap:10,padding:"10px 12px"}}>
                              <div style={{width:38,height:38,borderRadius:10,background:"var(--bl-l)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:20,flexShrink:0}}>{p?p.emoji:"📦"}</div>
                              <div style={{flex:1,minWidth:0}}>
                                <div style={{fontWeight:700,fontSize:12,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p?p.name:"Producto"}</div>
                                <div style={{fontSize:11,color:"var(--t3)",marginTop:1}}>{tx.qty} u. → <strong>{dest?dest.name:"?"}</strong></div>
                              </div>
                              <div style={{display:"flex",alignItems:"center",gap:6,flexShrink:0}}>
                                <span style={{background:"var(--bl-l)",color:"var(--bl-d)",borderRadius:20,padding:"3px 8px",fontSize:10,fontWeight:800}}>⏳ Pendiente</span>
                                <button onClick={function(){setCancelConfirm(confirmingCancel?null:tx.id);}}
                                  style={{background:"var(--cr-l)",color:"var(--cr-d)",border:"none",borderRadius:8,padding:"4px 8px",fontSize:11,fontWeight:700,cursor:"pointer"}}>
                                  ✕
                                </button>
                              </div>
                            </div>
                            {confirmingCancel&&(
                              <div style={{padding:"10px 12px",background:"#fff3f3",borderTop:"1px solid var(--brd)",display:"flex",alignItems:"center",gap:10}}>
                                <div style={{flex:1,fontSize:11,color:"var(--cr-d)",fontWeight:700}}>
                                  ¿Cancelar el envío? El stock vuelve a tu inventario.
                                </div>
                                <button onClick={function(){setCancelConfirm(null);cancelTransfer(tx);}}
                                  style={{background:"var(--cr)",color:"#fff",border:"none",borderRadius:8,padding:"6px 12px",fontSize:11,fontWeight:800,cursor:"pointer",flexShrink:0}}>
                                  Sí, cancelar
                                </button>
                                <button onClick={function(){setCancelConfirm(null);}}
                                  style={{background:"#eee",color:"#666",border:"none",borderRadius:8,padding:"6px 10px",fontSize:11,fontWeight:700,cursor:"pointer",flexShrink:0}}>
                                  No
                                </button>
                              </div>
                            )}
                          </div>
                        );
                      })}
                    </div>
                  </div>
                )}

                {/* ─ ENVÍOS PENDIENTES: LOS QUE ME MANDARON A MÍ ─ */}
                {pendingTx.length>0&&(
                  <div style={{margin:"0 14px 16px"}}>
                    <div style={{background:"var(--am-l)",border:"1.5px solid rgba(255,122,0,.25)",borderRadius:14,padding:"12px 14px"}}>
                      <div style={{fontWeight:800,fontSize:13,color:"var(--am-d)",marginBottom:8}}>🔔 {pendingTx.length} envío{pendingTx.length!==1?"s":""} esperando tu confirmación</div>
                      {pendingTx.map(function(tx){
                        const p=tx.product;
                        return (
                          <div key={tx.id} style={{display:"flex",alignItems:"center",gap:10,padding:"10px 12px",background:"var(--card)",borderRadius:12,marginBottom:6,border:"1px solid var(--brd)"}}>
                            <div style={{width:38,height:38,borderRadius:10,background:"var(--am-l)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:20,flexShrink:0}}>{p?p.emoji:"📦"}</div>
                            <div style={{flex:1,minWidth:0}}>
                              <div style={{fontWeight:700,fontSize:12,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p?p.name:"Producto"}</div>
                              <div style={{fontSize:11,color:"var(--t3)",marginTop:1}}>{tx.qty} u. de <strong>{tx.from_user?tx.from_user.name:""}</strong></div>
                            </div>
                            <button className="btn btn-xs b-em" onClick={function(){confirmTransfer(tx);}}>✓ Recibí</button>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                )}

                {/* ─ STOCK ─ */}
                <div style={{padding:"0 14px 8px"}}>

                  {/* Stock Propio */}
                  <div className="sec-hdr" style={{padding:"0 0 10px"}}>
                    <div style={{display:"flex",alignItems:"center",gap:8}}>
                      <div style={{width:10,height:10,borderRadius:3,background:"var(--em)"}}/>
                      <div className="sec-hdr-t">Stock Propio</div>
                    </div>
                    <div style={{display:"flex",alignItems:"center",gap:6}}>
                      {sentTx.length>0&&<span style={{background:"var(--bl-l)",color:"var(--bl-d)",borderRadius:20,padding:"2px 8px",fontSize:10,fontWeight:800}}>📤 {sentTx.length} en tránsito</span>}
                      <span className="badge" style={{background:"var(--em-l)",color:"var(--em-d)"}}>{ownFilt.length}</span>
                    </div>
                  </div>
                  {ownFilt.length===0
                    ?<div style={{textAlign:"center",padding:"24px 20px",color:"var(--t3)",fontSize:13,background:"var(--card)",borderRadius:14,marginBottom:16,border:"1px solid var(--brd)"}}>
                      {srchStock?("Sin resultados para \""+srchStock+"\""):<span>Sin stock propio.<br/>Usá <strong>Cargar Stock</strong> para agregar.</span>}
                    </div>
                    :StockTable(ownFilt, false)
                  }

                  {/* Recibido de otro vendedor */}
                  {(consignFilt.length>0||consignStock.length>0)&&(
                    <div style={{marginTop:8}}>
                      <div className="sec-hdr" style={{padding:"0 0 10px"}}>
                        <div style={{display:"flex",alignItems:"center",gap:8}}>
                          <div style={{width:10,height:10,borderRadius:3,background:"var(--am)"}}/>
                          <div className="sec-hdr-t">Recibido de otro vendedor</div>
                        </div>
                        <span className="badge" style={{background:"var(--am-l)",color:"var(--am-d)"}}>{consignFilt.length}</span>
                      </div>
                      <div style={{background:"var(--am-l)",border:"1.5px solid rgba(255,122,0,.2)",borderRadius:12,padding:"10px 14px",marginBottom:12,fontSize:11,color:"var(--am-d)",fontWeight:700}}>
                        📨 Estos productos te los enviaron en consignación. Podés venderlos o devolverlos desde la tab <strong>Recibidos</strong>.
                      </div>
                      {consignFilt.length===0
                        ?<div className="empty" style={{padding:"16px"}}>Sin resultados.</div>
                        :StockTable(consignFilt, true)
                      }
                    </div>
                  )}
                </div>
              </div>
            );
          })()}
          {/* ══ CARGAR ══ */}
          {tab==="cargar"&&(
            <div>
              <div className="ph"><div><div className="ph-h">Cargar Stock</div><div className="ph-s">Agrega unidades o crea productos</div></div></div>
              <div className="pc">
                <div className="card">
                  <div className="card-h"><div className="card-title">Modo de carga</div></div>
                  <div style={{padding:"14px 14px 0"}}>

                    <div style={{background:"var(--in-l)",borderRadius:14,padding:"10px 14px",marginBottom:12,border:"1.5px solid rgba(124,58,237,.15)"}}>
                      <div style={{fontSize:12,fontWeight:700,color:"var(--in-d)"}}>📦 Cargando Stock Propio</div>
                      <div style={{fontSize:11,color:"var(--t3)",marginTop:2}}>Para consignaciones usá la tab Consigna</div>
                    </div>

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

          {/* ══ PEDIDOS ══ */}
          {tab==="pedidos"&&(function(){
            var pedFilt = pedidos.filter(function(p){
              var q = (srchStock||"").toLowerCase();
              if (!q) return true;
              return (p.nombre||"").toLowerCase().includes(q)||(p.product_name||"").toLowerCase().includes(q);
            });
            var pend = pedFilt.filter(function(p){return p.estado==="pendiente";});
            var entregados = pedFilt.filter(function(p){return p.estado==="entregado";});
            var cancelados = pedFilt.filter(function(p){return p.estado==="cancelado";});
            function PedCard({ped}) {
              var prod = ped.product || products.find(function(p){return p.id===ped.product_id;});
              var statusCfg = {
                pendiente:  {bg:"#fff8e1",col:"#e06a00",lbl:"⏳ Pendiente"},
                entregado:  {bg:"#e8faf4",col:"#009a66",lbl:"✅ Entregado"},
                cancelado:  {bg:"#fde8ea",col:"#c1121f",lbl:"❌ Cancelado"},
              };
              var sc = statusCfg[ped.estado]||statusCfg.pendiente;
              var pTotal = ped.total || (ped.product_price||0)*(ped.qty||1);
              var pSena  = ped.sena || 0;
              var pSaldo = Math.max(0, pTotal - pSena);
              var conDeuda = pSaldo>0 && !ped.pagado && ped.estado!=="cancelado";
              return (
                <div style={{background:"var(--card)",borderRadius:16,border:"1.5px solid var(--brd)",marginBottom:10,overflow:"hidden",boxShadow:"var(--sh)"}}>
                  <div style={{display:"flex",alignItems:"center",gap:12,padding:"13px 14px"}}>
                    <div style={{width:44,height:44,borderRadius:12,background:"#f5f5f5",display:"flex",alignItems:"center",justifyContent:"center",fontSize:22,flexShrink:0}}>
                      {prod?<ProdThumb prod={prod} size={44}/>:<span>📦</span>}
                    </div>
                    <div style={{flex:1,minWidth:0}}>
                      <div style={{fontWeight:800,fontSize:14,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{ped.nombre}</div>
                      {prod&&<div style={{fontSize:12,color:"var(--t3)",marginTop:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{prod.name} · {fmtARS(prod.price||0)} × {ped.qty||1}</div>}
                      {ped.wa&&<a href={"https://wa.me/"+ped.wa.replace(/\D/g,"")} target="_blank" rel="noreferrer" style={{fontSize:11,color:"var(--wa-d)",marginTop:2,display:"block",fontWeight:700}}>📱 {ped.wa}</a>}
                      {ped.nota&&<div style={{fontSize:11,color:"var(--t3)",marginTop:2,fontStyle:"italic"}}>{ped.nota}</div>}
                    </div>
                    <div style={{textAlign:"right",flexShrink:0}}>
                      <div style={{background:sc.bg,color:sc.col,borderRadius:20,padding:"4px 10px",fontSize:11,fontWeight:800,marginBottom:4}}>{sc.lbl}</div>
                      <div style={{fontSize:10,color:"var(--t3)"}}>{new Date(ped.created_at).toLocaleDateString("es-AR")}</div>
                    </div>
                  </div>
                  {/* Desglose financiero */}
                  {pTotal>0&&ped.estado!=="cancelado"&&(
                    <div style={{padding:"0 14px 10px"}}>
                      <div style={{background:conDeuda?"#fff7ed":"#ecfdf5",borderRadius:12,padding:"10px 12px"}}>
                        <div style={{display:"flex",justifyContent:"space-between",fontSize:12,marginBottom:pSena>0?4:0}}>
                          <span style={{color:"var(--t2)",fontWeight:700}}>Total</span>
                          <span style={{fontFamily:"var(--mf)",fontWeight:800,color:"var(--t1)"}}>{fmtARS(pTotal)}</span>
                        </div>
                        {pSena>0&&(
                          <div style={{display:"flex",justifyContent:"space-between",fontSize:12,marginBottom:4}}>
                            <span style={{color:"#059669",fontWeight:700}}>Seña pagada</span>
                            <span style={{fontFamily:"var(--mf)",fontWeight:800,color:"#059669"}}>− {fmtARS(pSena)}</span>
                          </div>
                        )}
                        <div style={{display:"flex",justifyContent:"space-between",fontSize:13,paddingTop:pSena>0?6:0,borderTop:pSena>0?"1px solid rgba(0,0,0,.06)":"none"}}>
                          <span style={{fontWeight:800,color:conDeuda?"#d97706":"#059669"}}>{ped.pagado?"✅ Pagado":(conDeuda?"🔴 Saldo pendiente":"Saldo")}</span>
                          <span style={{fontFamily:"var(--mf)",fontWeight:900,fontSize:15,color:conDeuda?"#d97706":"#059669"}}>{ped.pagado?fmtARS(0):fmtARS(pSaldo)}</span>
                        </div>
                      </div>
                    </div>
                  )}
                  {/* Acciones */}
                  {ped.estado!=="cancelado"&&(
                    <div style={{display:"flex",borderTop:"1px solid var(--brd)"}}>
                      {ped.wa&&(
                        <button onClick={function(){waPedidoGracias(ped);}} style={{flex:1,padding:"10px",border:"none",borderRight:"1px solid var(--brd)",background:"#f0fdf4",color:"#16a34a",fontFamily:"var(--hf)",fontWeight:800,fontSize:12.5,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",gap:5}}>
                          <Ic n="wa" s={14}/>Gracias
                        </button>
                      )}
                      {conDeuda&&(
                        <button onClick={function(){doMarcarPagado(ped.id);}} style={{flex:1,padding:"10px",border:"none",borderRight:"1px solid var(--brd)",background:"#eff6ff",color:"#0284c7",fontFamily:"var(--hf)",fontWeight:800,fontSize:12.5,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",gap:5}}>
                          <Ic n="check" s={14}/>Pagado
                        </button>
                      )}
                      {ped.estado==="pendiente"&&(
                        <button onClick={function(){doEntregarPedido(ped.id);}} style={{flex:1,padding:"10px",border:"none",borderRight:"1px solid var(--brd)",background:"#e8faf4",color:"#009a66",fontFamily:"var(--hf)",fontWeight:800,fontSize:12.5,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",gap:5}}>
                          <Ic n="check" s={14}/>Entregar
                        </button>
                      )}
                      {ped.estado==="pendiente"&&(
                        <button onClick={function(){doCancelarPedido(ped.id);}} style={{flex:1,padding:"10px",border:"none",background:"#fde8ea",color:"#c1121f",fontFamily:"var(--hf)",fontWeight:800,fontSize:12.5,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",gap:5}}>
                          <Ic n="x" s={14}/>Cancelar
                        </button>
                      )}
                    </div>
                  )}
                </div>
              );
            }
            return (
              <div>
                <div className="ph">
                  <div>
                    <div className="ph-h">Pedidos</div>
                    <div className="ph-s">{pedPendCount>0?pedPendCount+" pendientes de entrega":"Todos al día"}</div>
                  </div>
                  <div style={{display:"flex",gap:8}}>
                    <button className="btn" style={{background:"#e7f9ee",color:"#0a8f4d",border:"1px solid #bfe9d2",borderRadius:12,fontWeight:700,fontSize:12,padding:"9px 12px",cursor:"pointer"}} onClick={function(){
                      var link = window.location.origin + "/?tienda=" + me.id;
                      var msg = "¡Hola! 👋 Mirá mi tienda y hacé tu pedido acá: " + link;
                      window.open("https://wa.me/?text=" + encodeURIComponent(msg), "_blank");
                    }}>📲 Compartir tienda</button>
                    <button className="btn b-pri" onClick={function(){setShowPedForm(function(v){return !v;});}}>
                      <Ic n="plus" s={15}/>{showPedForm?"Cancelar":"Nuevo"}
                    </button>
                  </div>
                </div>
                <div className="pc">
                  {/* Formulario nuevo pedido */}
                  {showPedForm&&(
                    <div style={{background:"var(--card)",borderRadius:16,border:"1.5px solid var(--brd)",marginBottom:14,padding:"14px 14px",boxShadow:"var(--sh)"}}>
                      <div style={{fontWeight:800,fontSize:14,marginBottom:12,color:"var(--t1)"}}>Nuevo pedido</div>
                      <div className="fld"><label className="fl">Nombre del cliente *</label><input className="fi" placeholder="Ej: María López" value={pedNombre} onChange={function(e){setPedNombre(e.target.value);}}/></div>
                      <div className="fld"><label className="fl">WhatsApp</label><input className="fi" placeholder="+54 9 11..." type="tel" value={pedWA} onChange={function(e){setPedWA(e.target.value);}}/></div>
                      <div className="fld">
                        <label className="fl">Producto (opcional)</label>
                        <SearchBar value={pedPSrch} onChange={setPedPSrch} placeholder="Buscar producto..."/>
                        {pedPSrch&&(
                          <div style={{maxHeight:180,overflowY:"auto",marginTop:6,borderRadius:10,border:"1px solid var(--brd)"}}>
                            {products.filter(function(p){var q=pedPSrch.toLowerCase();return p.name.toLowerCase().includes(q)||(p.sku||"").toLowerCase().includes(q);}).slice(0,8).map(function(p){
                              return (
                                <div key={p.id} onClick={function(){setPedProdId(p.id);setPedPSrch(p.name+" ["+(p.sku||"")+"]");}} style={{padding:"10px 12px",cursor:"pointer",borderBottom:"1px solid var(--brd)",display:"flex",alignItems:"center",gap:8,background:pedProdId===p.id?"var(--pri-l)":"var(--card)"}}>
                                  <ProdThumb prod={p} size={32}/>
                                  <div>
                                    <div style={{fontSize:13,fontWeight:700}}>{p.name}</div>
                                    <div style={{fontSize:11,color:"var(--t3)"}}>
                                      <span style={{fontFamily:"var(--mf)",fontWeight:800,color:"var(--pri)",background:"var(--pri-l)",borderRadius:5,padding:"1px 6px",marginRight:6}}>{p.sku||"sin código"}</span>
                                      {fmtARS(p.price)}
                                    </div>
                                  </div>
                                </div>
                              );
                            })}
                          </div>
                        )}
                      </div>
                      {pedProdId&&(
                        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:12}}>
                          <label style={{fontSize:12,fontWeight:700,color:"var(--t2)"}}>Cantidad</label>
                          <QtyControl val={pedQty} set={setPedQty}/>
                        </div>
                      )}
                      {pedProdId&&(
                        <div style={{background:"var(--bg2)",borderRadius:12,padding:"10px 12px",marginBottom:12,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                          <span style={{fontSize:12,color:"var(--t2)",fontWeight:700}}>Total del pedido</span>
                          <span style={{fontFamily:"var(--mf)",fontWeight:900,fontSize:15,color:"var(--in)"}}>{fmtARS((products.find(function(p){return p.id===pedProdId;})||{}).price*pedQty||0)}</span>
                        </div>
                      )}
                      <div className="fld"><label className="fl">Seña / pago adelantado (opcional)</label><input className="fi" type="number" placeholder="0" value={pedSena} onChange={function(e){setPedSena(e.target.value);}}/></div>
                      <div className="fld"><label className="fl">Nota</label><input className="fi" placeholder="Ej: talla, color, etc." value={pedNota} onChange={function(e){setPedNota(e.target.value);}}/></div>
                      <button className="cta cta-am" onClick={doAddPedido}>
                        <Ic n="check" s={16}/>Guardar pedido
                      </button>
                    </div>
                  )}
                  {/* Pendientes */}
                  {pend.length>0&&(
                    <div style={{marginBottom:16}}>
                      <div style={{fontSize:11,fontWeight:800,color:"#e06a00",marginBottom:8,textTransform:"uppercase",letterSpacing:".07em"}}>⏳ Pendientes ({pend.length})</div>
                      {pend.map(function(p){return <PedCard key={p.id} ped={p}/>;}) }
                    </div>
                  )}
                  {/* Entregados */}
                  {entregados.length>0&&(
                    <div style={{marginBottom:16}}>
                      <div style={{fontSize:11,fontWeight:800,color:"#009a66",marginBottom:8,textTransform:"uppercase",letterSpacing:".07em"}}>✅ Entregados ({entregados.length})</div>
                      {entregados.map(function(p){return <PedCard key={p.id} ped={p}/>;}) }
                    </div>
                  )}
                  {/* Cancelados */}
                  {cancelados.length>0&&(
                    <div style={{marginBottom:16}}>
                      <div style={{fontSize:11,fontWeight:800,color:"#c1121f",marginBottom:8,textTransform:"uppercase",letterSpacing:".07em"}}>❌ Cancelados ({cancelados.length})</div>
                      {cancelados.map(function(p){return <PedCard key={p.id} ped={p}/>;}) }
                    </div>
                  )}
                  {pedFilt.length===0&&(
                    <div style={{textAlign:"center",padding:"60px 20px",color:"var(--t3)"}}>
                      <div style={{fontSize:48,marginBottom:12}}>📋</div>
                      <div style={{fontSize:14,fontWeight:700}}>No hay pedidos</div>
                      <div style={{fontSize:12,marginTop:4}}>Tocá "Nuevo" para crear el primero</div>
                    </div>
                  )}
                </div>
              </div>
            );
          })()}

          {/* ══ CATÁLOGO ══ */}
          {tab==="catalog"&&isAdmin&&(function(){
            var catGroups = {};
            catFiltered.forEach(function(p){ var c=p.category||"General"; if(!catGroups[c]) catGroups[c]=[]; catGroups[c].push(p); });
            return (
              <div>
                <div className="ph">
                  <div><div className="ph-h">Catálogo</div><div className="ph-s">{catFiltered.length} productos</div></div>
                  <div style={{display:"flex",gap:8}}>
                    {isAdmin&&<button className="btn b-ghost" onClick={function(){setShowInactive(function(v){return !v;});}}>
                      {showInactive?"Ver activos":"Ver inactivos"}
                    </button>}
                    <button className="btn b-pri" onClick={function(){cancelEdit();setTab("catalog");}}>
                      <Ic n="plus" s={15}/>{editP?"Cancelar":"Nuevo"}
                    </button>
                  </div>
                </div>
                <div className="pc">
                  <SearchBar value={srchCat} onChange={setSrchCat} placeholder="Buscar SKU, nombre, categoría..."/>
                  {/* Formulario crear/editar */}
                  {(editP!==null||editP===null&&false)&&false}
                  <div className="card" style={{marginBottom:14}}>
                    <div className="card-h"><div className="card-title">{editP?"✏️ Modificar":"✨ Nuevo producto"}</div>{editP&&<button className="btn btn-xs b-ghost" onClick={cancelEdit}>Cancelar</button>}</div>
                    <form onSubmit={doSaveProd} style={{padding:"14px 16px 6px"}}>
                      <div className="row g8" style={{marginBottom:12}}>
                        <div style={{flex:1}}><label className="fl">SKU *</label><input className="fi" placeholder="PERF01" value={fSku} onChange={function(e){setFSku(e.target.value);}}/></div>
                        <div style={{flex:1}}><label className="fl">Ícono</label><select className="fi fi-sel" value={fEmoji} onChange={function(e){setFEmoji(e.target.value);}}>{EMOJIS.map(function(o){return <option key={o.v} value={o.v}>{o.v} {o.l}</option>;})}</select></div>
                      </div>
                      <div className="fld"><label className="fl">Nombre *</label><input className="fi" placeholder="Perfume Lattafa 100ml" value={fName} onChange={function(e){setFName(e.target.value);}}/></div>
                      <div className="row g8" style={{marginBottom:12}}>
                        <div style={{flex:1}}><label className="fl">Precio ($) *</label><input className="fi" type="number" placeholder="0.00" value={fPrice} onChange={function(e){setFPrice(e.target.value);}}/></div>
                        <div style={{flex:1}}><label className="fl">Categoría</label><select className="fi fi-sel" value={fCat} onChange={function(e){setFCat(e.target.value);}}>{CATS.map(function(c){return <option key={c}>{c}</option>;})}</select></div>
                      </div>
                      {/* Carrusel de imágenes */}
                      <div className="fld">
                        <label className="fl">Fotos del producto</label>
                        {editP&&(prodImages[editP.id]||[]).length>0&&(
                          <div style={{marginBottom:10}}>
                            <div style={{display:"flex",gap:8,overflowX:"auto",paddingBottom:6}}>
                              {(prodImages[editP.id]||[]).map(function(img,i){
                                return (
                                  <div key={img.id} style={{position:"relative",flexShrink:0}}>
                                    <img src={img.url} alt="" style={{width:72,height:72,borderRadius:10,objectFit:"cover",border:img.es_principal?"2.5px solid var(--pri)":"1.5px solid #ddd"}}/>
                                    {img.es_principal&&<div style={{position:"absolute",top:2,left:2,background:"var(--pri)",borderRadius:6,padding:"1px 5px",fontSize:9,color:"#fff",fontWeight:800}}>★</div>}
                                    <div style={{position:"absolute",top:2,right:2,display:"flex",flexDirection:"column",gap:2}}>
                                      <button type="button" onClick={function(){setPrincipalImage(img,editP.id);}} style={{width:18,height:18,borderRadius:4,background:"#ffcc00",border:"none",cursor:"pointer",fontSize:9}}>★</button>
                                      <button type="button" onClick={function(){deleteProdImage(img,editP.id);}} style={{width:18,height:18,borderRadius:4,background:"var(--cr)",border:"none",cursor:"pointer",fontSize:9,color:"#fff"}}>✕</button>
                                    </div>
                                  </div>
                                );
                              })}
                            </div>
                          </div>
                        )}
                        {editP?(
                          <label style={{display:"block",background:"var(--in-l)",border:"1.5px dashed var(--in-m)",borderRadius:10,padding:"10px",textAlign:"center",cursor:"pointer"}}>
                            <input type="file" accept="image/*" multiple style={{display:"none"}} onChange={function(e){Array.from(e.target.files||[]).forEach(function(f){uploadProdImage(editP.id,f);});e.target.value="";}}/>
                            {imgUploading?<div style={{fontSize:12,color:"var(--in-d)",fontWeight:700}}>Subiendo...</div>
                              :<div><div style={{fontSize:20,marginBottom:2}}>📸</div><div style={{fontSize:12,fontWeight:700,color:"var(--in-d)"}}>Subir imágenes</div></div>}
                          </label>
                        ):(
                          <div className="photo-zone"><input type="file" accept="image/*" onChange={function(e){handlePhoto(e,setFPhoto);}}/>{fPhoto?<img src={fPhoto} alt="" style={{width:56,height:56,borderRadius:10,objectFit:"cover",margin:"0 auto"}}/>:<div><div style={{fontSize:24,marginBottom:4}}>📸</div><div style={{fontSize:12,fontWeight:700,color:"var(--in)"}}>Tocar para subir</div></div>}</div>
                        )}
                        {!editP&&fPhoto&&<button type="button" onClick={function(){setFPhoto(null);}} style={{fontSize:11,color:"var(--cr)",background:"none",border:"none",cursor:"pointer",fontWeight:700}}>Quitar foto</button>}
                      </div>
                      <button type="submit" className="cta cta-in"><Ic n="check" s={18}/>{editP?"Guardar cambios":"Crear producto"}</button>
                    </form>
                  </div>
                  {/* Lista de productos por categoría */}
                  {Object.keys(catGroups).sort().map(function(cat){
                    return (
                      <div key={cat} style={{marginBottom:8}}>
                        <div style={{fontSize:10,fontWeight:800,color:"var(--t3)",textTransform:"uppercase",letterSpacing:".08em",marginBottom:8,paddingLeft:2}}>{cat} · {catGroups[cat].length}</div>
                        {catGroups[cat].map(function(p){
                          var imgs=prodImages[p.id]||[];
                          var inv=inventory.find(function(i){return i.product_id===p.id&&i.user_id===me.id;});
                          return (
                            <div key={p.id} style={{background:"var(--card)",borderRadius:"var(--r2)",border:"1px solid rgba(0,0,0,.055)",marginBottom:10,overflow:"hidden",boxShadow:"var(--sh)"}}>
                              <div style={{display:"flex",alignItems:"center",gap:12,padding:"12px 14px"}}>
                                <ProdThumb prod={p} images={imgs} size={52}/>
                                <div style={{flex:1,minWidth:0}}>
                                  <div style={{fontWeight:800,fontSize:13,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.emoji} {p.name}</div>
                                  <div style={{fontSize:11,color:"var(--t3)",marginTop:1}}>{p.sku}</div>
                                  <div style={{fontSize:13,fontWeight:800,color:"var(--pri)",marginTop:2,fontFamily:"var(--mf)"}}>{fmtARS(p.price)}</div>
                                </div>
                                <div style={{display:"flex",flexDirection:"column",gap:6,alignItems:"flex-end"}}>
                                  {inv&&<div style={{fontFamily:"var(--mf)",fontWeight:800,fontSize:16,color:inv.qty_available>0?"var(--em-d)":"var(--t4)"}}>{inv.qty_available}</div>}
                                  <button className="btn btn-xs b-ghost" onClick={function(){startEdit(p);}}><Ic n="edit" s={12}/>Editar</button>
                                </div>
                              </div>
                              {p.is_active===false&&<div style={{background:"var(--cr-l)",padding:"4px 14px",fontSize:11,color:"var(--cr)",fontWeight:700}}>⚠️ Inactivo</div>}
                            </div>
                          );
                        })}
                      </div>
                    );
                  })}
                  {catFiltered.length===0&&<div className="empty"><div style={{fontSize:40,marginBottom:8}}>📭</div>Sin productos</div>}
                </div>
              </div>
            );
          })()}

          {/* ══ IMPORTAR ══ */}
          {tab==="importar"&&isAdmin&&(
            <div>
              <div className="ph"><div><div className="ph-h">Importar</div><div className="ph-s">Carga masiva de productos</div></div></div>
              <div className="pc">
                <div className="card">
                  <div className="card-h"><div className="card-title">📂 Subir archivo Excel / CSV</div></div>
                  <div style={{padding:"14px 16px"}}>
                    <div className="drop-z" onClick={function(){document.getElementById("imp-file").click();}}>
                      <input id="imp-file" type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={function(e){
                        var f=e.target.files[0]; if(!f) return; setImpFile(f.name);
                        var reader=new FileReader();
                        reader.onload=function(ev){
                          var data=new Uint8Array(ev.target.result);
                          var wb=XLSX.read(data,{type:"array"});
                          var ws=wb.Sheets[wb.SheetNames[0]];
                          var rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:""});
                          var parsed=rows.slice(1).filter(function(r){return r[0]||r[1];}).map(function(r){
                            var sku=String(r[0]||"X"+Math.random().toString(36).slice(2,6)).toUpperCase().trim();
                            var name=String(r[1]||"Sin nombre").trim();
                            var price=parseFloat(String(r[2]||0).replace(/[^0-9.]/g,""))||0;
                            var cat=String(r[3]||"General").trim();
                            return {sku,name,price,cat,ok:!!(sku&&name&&price)};
                          });
                          setImpRows(parsed);
                          toast(parsed.filter(function(r){return r.ok;}).length+" filas válidas","Revisa y confirma","s");
                        };
                        reader.readAsArrayBuffer(f);
                        e.target.value="";
                      }}/>
                      <div style={{fontSize:36,marginBottom:8}}>📊</div>
                      <div style={{fontWeight:700,color:"var(--t2)"}}>{impFile?impFile:"Elegir Excel o CSV"}</div>
                      <div style={{fontSize:11,color:"var(--t3)",marginTop:4}}>Columnas: SKU · Nombre · Precio · Categoría</div>
                    </div>
                  </div>
                </div>
                {impRows.length>0&&(
                  <div>
                    <div className="card">
                      <div className="card-h"><div className="card-title">Vista previa ({impRows.filter(function(r){return r.ok;}).length} válidos)</div></div>
                      <div style={{padding:"10px 14px"}}>
                        {impRows.slice(0,20).map(function(r,i){
                          return (
                            <div key={i} className={"prev-row "+(r.ok?"prev-ok":"prev-err")}>
                              <div style={{flex:1,minWidth:0}}>
                                <div style={{fontWeight:700,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.sku} · {r.name}</div>
                                <div style={{fontSize:11,opacity:.7}}>{fmtARS(r.price)} · {r.cat}</div>
                              </div>
                              <span style={{marginLeft:8}}>{r.ok?"✅":"❌"}</span>
                            </div>
                          );
                        })}
                        {impRows.length>20&&<div style={{textAlign:"center",padding:8,fontSize:12,color:"var(--t3)"}}>...y {impRows.length-20} más</div>}
                      </div>
                    </div>
                    <button className="cta cta-in" onClick={doImport} disabled={impRows.filter(function(r){return r.ok;}).length===0}>
                      <Ic n="upload" s={18}/>Importar {impRows.filter(function(r){return r.ok;}).length} productos
                    </button>
                  </div>
                )}
              </div>
            </div>
          )}

          {/* ══ CONSIGNA (hub: Enviados / Recibidos / Mi Red) ══ */}
          {tab==="consigna"&&(
            <div>
              <div className="ph"><div><div className="ph-h">🤝 Consigna</div><div className="ph-s">Elegí qué querés ver</div></div></div>
              <div className="pc">
                <div onClick={function(){setTab("enviados");}} className="card" style={{marginBottom:12,cursor:"pointer"}}>
                  <div style={{padding:"16px",display:"flex",alignItems:"center",gap:14}}>
                    <div style={{width:46,height:46,borderRadius:14,background:"var(--bl-l)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:22}}>📤</div>
                    <div style={{flex:1}}>
                      <div style={{fontSize:15,fontWeight:800}}>Enviados</div>
                      <div style={{fontSize:12,color:"var(--t3)"}}>Productos que entregaste a otras vendedoras</div>
                    </div>
                    {consignaEnv>0&&<span style={{background:"var(--cr,#e0224e)",color:"#fff",borderRadius:9,minWidth:22,height:22,fontSize:11,fontWeight:800,display:"flex",alignItems:"center",justifyContent:"center",padding:"0 6px"}}>{consignaEnv}</span>}
                  </div>
                </div>
                <div onClick={function(){setTab("recibidos");}} className="card" style={{marginBottom:12,cursor:"pointer"}}>
                  <div style={{padding:"16px",display:"flex",alignItems:"center",gap:14}}>
                    <div style={{width:46,height:46,borderRadius:14,background:"#e0f2fe",display:"flex",alignItems:"center",justifyContent:"center",fontSize:22}}>📥</div>
                    <div style={{flex:1}}>
                      <div style={{fontSize:15,fontWeight:800}}>Recibidos</div>
                      <div style={{fontSize:12,color:"var(--t3)"}}>Productos que te entregaron para vender</div>
                    </div>
                  </div>
                </div>
                <div onClick={function(){setTab("contacts");}} className="card" style={{cursor:"pointer"}}>
                  <div style={{padding:"16px",display:"flex",alignItems:"center",gap:14}}>
                    <div style={{width:46,height:46,borderRadius:14,background:"var(--pu-l,#f3eafe)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:22}}>🌐</div>
                    <div style={{flex:1}}>
                      <div style={{fontSize:15,fontWeight:800}}>Mi Red</div>
                      <div style={{fontSize:12,color:"var(--t3)"}}>Personas con las que trabajás en consigna</div>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* ══ ENVIADOS ══ */}
          {tab==="enviados"&&(
            <ConsignacionModule
              sb={sb}
              me={me}
              products={products}
              inventory={inventory}
              contacts={contacts}
              onRefresh={function(){ loadData(me.id, me.role); }}
              toast={toast}
              fmtARS={fmtARS}
              vistaInicial="enviados"
            />
          )}

          {/* ══ RECIBIDOS ══ */}
          {tab==="recibidos"&&(
            <ConsignacionModule
              sb={sb}
              me={me}
              products={products}
              inventory={inventory}
              contacts={contacts}
              onRefresh={function(){ loadData(me.id, me.role); }}
              toast={toast}
              fmtARS={fmtARS}
              vistaInicial="recibidos"
            />
          )}

          {/* ══ VENTAS ══ */}
          {tab==="ventas"&&(function(){
            var meses = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];
            var now = new Date();
            // Mes objetivo según el offset elegido
            var target = new Date(now.getFullYear(), now.getMonth() + ventasMesOffset, 1);
            var mesActual  = target.getMonth();
            var anioActual = target.getFullYear();

            // Filtrar logs por el mes elegido (o todos si ventasVerTodo)
            var logsDelMes = logs.filter(function(l){
              if (ventasVerTodo) return true;
              var d = new Date(l.created_at);
              return d.getMonth()===mesActual && d.getFullYear()===anioActual;
            });

            // Group by product
            var byProd = {};
            logsDelMes.forEach(function(l){
              var pid = l.product_id;
              var p = l.product || products.find(function(x){ return x.id===pid; });
              if (!byProd[pid]) byProd[pid] = {prod:p, qty:0, total:0, isConsigna: l.source==="consignment"};
              byProd[pid].qty   += l.qty || 1;
              byProd[pid].total += (l.sale_price || (p?p.price:0)) * (l.qty||1);
            });
            var ventaItems = Object.values(byProd).sort(function(a,b){ return b.total-a.total; });

            var totalVentas   = ventaItems.reduce(function(s,v){ return s+v.total; }, 0);
            var totalUnidades = ventaItems.reduce(function(s,v){ return s+v.qty;   }, 0);
            var gananciaTotal = totalVentas * (ganancia/100);

            // Reload logs on demand
            async function recargarVentas() {
              var lg = await sb.from("sale_logs")
                .select("*, product:product_id(name,sku,emoji)")
                .eq("user_id", me.id)
                .order("created_at",{ascending:false})
                .limit(500);
              if (lg.data) setLogs(lg.data);
              toast("Ventas actualizadas","","s");
            }

            return (
              <div>
                <div className="ph">
                  <div><div className="ph-h">Mis Ventas</div><div className="ph-s">{ventasVerTodo?"Historial completo":(meses[mesActual]+" "+anioActual)}</div></div>
                  <button className="btn btn-xs b-ghost" onClick={recargarVentas}><Ic n="undo" s={13}/>Actualizar</button>
                </div>
                <div className="pc">

                  {/* Selector de mes */}
                  <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:14}}>
                    <button onClick={function(){ setVentasVerTodo(false); setVentasMesOffset(function(o){return o-1;}); }}
                      style={{width:40,height:40,borderRadius:12,border:"1.5px solid var(--brd)",background:"var(--card)",color:"var(--t2)",cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0,fontSize:18,fontWeight:800}}>‹</button>
                    <div onClick={function(){ setVentasVerTodo(function(v){return !v;}); }}
                      style={{flex:1,textAlign:"center",padding:"11px",borderRadius:12,border:"1.5px solid "+(ventasVerTodo?"var(--in-m)":"var(--brd)"),background:ventasVerTodo?"var(--in-l)":"var(--card)",cursor:"pointer"}}>
                      <div style={{fontWeight:800,fontSize:14,color:ventasVerTodo?"var(--in-d)":"var(--t1)"}}>
                        {ventasVerTodo?"📅 Ver todo el historial":(meses[mesActual]+" "+anioActual)}
                      </div>
                      <div style={{fontSize:10,color:"var(--t3)",marginTop:1}}>{ventasVerTodo?"Tocá para volver al mes":"Tocá para ver todo"}</div>
                    </div>
                    <button onClick={function(){ setVentasVerTodo(false); setVentasMesOffset(function(o){return Math.min(0,o+1);}); }} disabled={ventasMesOffset>=0&&!ventasVerTodo}
                      style={{width:40,height:40,borderRadius:12,border:"1.5px solid var(--brd)",background:"var(--card)",color:(ventasMesOffset>=0&&!ventasVerTodo)?"var(--t4)":"var(--t2)",cursor:(ventasMesOffset>=0&&!ventasVerTodo)?"not-allowed":"pointer",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0,fontSize:18,fontWeight:800,opacity:(ventasMesOffset>=0&&!ventasVerTodo)?.4:1}}>›</button>
                  </div>

                  {/* Summary cards */}
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:16}}>
                    <div className="metric-card">
                      <div className="metric-val" style={{color:"var(--em-d)",fontSize:24}}>{fmtARS(totalVentas)}</div>
                      <div className="metric-lbl">Total vendido</div>
                    </div>
                    <div className="metric-card">
                      <div className="metric-val" style={{color:"var(--in-d)",fontSize:28}}>{totalUnidades}</div>
                      <div className="metric-lbl">Unidades</div>
                    </div>
                  </div>

                  {/* Consignment debts */}
                  {consignDebts.filter(function(d){return !d.paid;}).length>0&&(
                    <div className="card" style={{marginBottom:16,border:"2px solid var(--am)"}}>
                      <div className="card-h" style={{background:"var(--am-l)"}}>
                        <div className="card-title" style={{color:"var(--am-d)"}}>🤝 Deudas de Consignación</div>
                        <span className="badge b-am">{fmtARS(consignDebts.filter(function(d){return !d.paid;}).reduce(function(s,d){return s+d.amount;},0))}</span>
                      </div>
                      <div style={{padding:"8px 0"}}>
                        {consignDebts.filter(function(d){return !d.paid;}).map(function(debt){
                          return (
                            <div key={debt.id} style={{display:"flex",alignItems:"center",gap:12,padding:"12px 16px",borderBottom:"1px solid var(--brd)"}}>
                              <div style={{width:36,height:36,borderRadius:10,background:"var(--am-l)",display:"flex",alignItems:"center",justifyContent:"center",fontWeight:900,fontSize:14,color:"var(--am-d)",flexShrink:0}}>
                                {debt.supplier_name.charAt(0).toUpperCase()}
                              </div>
                              <div style={{flex:1,minWidth:0}}>
                                <div style={{fontWeight:700,fontSize:13}}>{debt.supplier_name}</div>
                                <div style={{fontSize:11,color:"var(--t3)",marginTop:1}}>{debt.product_name} · {debt.qty} u. · {new Date(debt.date).toLocaleDateString("es-AR")}</div>
                              </div>
                              <div style={{textAlign:"right",flexShrink:0}}>
                                <div style={{fontFamily:"var(--mf)",fontWeight:900,color:"var(--am-d)",fontSize:14}}>{fmtARS(debt.amount)}</div>
                                <button className="btn btn-xs b-em" style={{marginTop:4}} onClick={function(){
                                  setConsignDebts(function(prev){ return prev.map(function(d){ return d.id===debt.id?Object.assign({},d,{paid:true}):d; }); });
                                  toast("Deuda liquidada",""+debt.supplier_name,"s");
                                }}>✓ Liquidar</button>
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  )}

                  {/* Ganancia card */}
                  <div className="card" style={{marginBottom:16,border:"2px solid var(--in-t)"}}>
                    <div className="card-h" style={{background:"var(--in-l)"}}>
                      <div className="card-title" style={{color:"var(--in-d)"}}>💰 Mi Ganancia Estimada</div>
                      <button className="btn btn-xs b-in" onClick={function(){ setGanInput(String(ganancia)); setEditGan(true); }}>Editar %</button>
                    </div>
                    <div style={{padding:"18px 16px"}}>
                      {editGan?(
                        <div>
                          <label className="fl">Porcentaje de ganancia (%)</label>
                          <div className="row g8" style={{marginBottom:8}}>
                            <input className="fi" type="number" min="0" max="100" step="0.5" value={ganInput} onChange={function(e){setGanInput(e.target.value);}} style={{flex:1}}/>
                            <span style={{fontSize:22,fontWeight:800,color:"var(--in-d)"}}>%</span>
                          </div>
                          <div className="row g8">
                            <button className="btn b-ghost" onClick={function(){setEditGan(false);}}>Cancelar</button>
                            <button className="btn b-pri" onClick={function(){ const val=parseFloat(ganInput)||0; setGanancia(val); setEditGan(false); }}>Guardar</button>
                          </div>
                        </div>
                      ):(
                        <div>
                          <div style={{fontFamily:"var(--mf)",fontWeight:900,fontSize:36,color:"var(--in-d)",lineHeight:1,marginBottom:8}}>{fmtARS(gananciaTotal)}</div>
                          <div style={{fontSize:12,color:"var(--t3)"}}>Basado en <strong>{ganancia}%</strong> sobre {fmtARS(totalVentas)} vendidos</div>
                          <div style={{marginTop:12,background:"var(--bg)",borderRadius:8,overflow:"hidden",height:8}}>
                            <div style={{height:"100%",width:Math.min(100,ganancia)+"%",background:"linear-gradient(90deg,var(--in),var(--pu))",borderRadius:8}}/>
                          </div>
                        </div>
                      )}
                    </div>
                  </div>

                  {/* Detalle por producto */}
                  <div className="card">
                    <div className="card-h">
                      <div className="card-title"><div className="card-ico" style={{background:"var(--em-l)",color:"var(--em-d)"}}><Ic n="clock" s={14}/></div>{ventasVerTodo?"Detalle (todo)":"Detalle de "+meses[mesActual]}</div>
                      <span className="badge b-em">{ventaItems.length} productos</span>
                    </div>
                    {ventaItems.length===0
                      ?<div className="empty">Sin ventas este mes.<br/>Registrá ventas con el botón "Venta" en tu Stock.</div>
                      :<div className="tw"><table>
                        <thead><tr><th></th><th>Producto</th><th>Tipo</th><th>Unid.</th><th>Total</th><th>Ganancia</th></tr></thead>
                        <tbody>
                          {ventaItems.map(function(item,idx){
                            var p=item.prod;
                            var gan=item.total*(ganancia/100);
                            return (
                              <tr key={idx} className="tr">
                                <td><ProdThumb prod={p} size={32}/></td>
                                <td><div style={{fontWeight:600,fontSize:12,maxWidth:120,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p?p.name:"—"}</div></td>
                                <td><span style={{fontSize:10,fontWeight:700,background:item.isConsigna?"var(--am-l)":"var(--in-l)",color:item.isConsigna?"var(--am-d)":"var(--in-d)",borderRadius:20,padding:"2px 7px"}}>{item.isConsigna?"🤝 Consigna":"📦 Propio"}</span></td>
                                <td><span style={{fontFamily:"var(--mf)",fontWeight:800,fontSize:14,color:"var(--in-d)"}}>{item.qty}</span></td>
                                <td><span style={{fontFamily:"var(--mf)",fontWeight:700,fontSize:12,color:"var(--em-d)"}}>{fmtARS(item.total)}</span></td>
                                <td><span style={{fontFamily:"var(--mf)",fontWeight:700,fontSize:12,color:"var(--in-d)"}}>{fmtARS(gan)}</span></td>
                              </tr>
                            );
                          })}
                          <tr style={{background:"var(--in-l)"}}>
                            <td colSpan={3} style={{fontWeight:800,fontSize:12,color:"var(--in-d)",padding:"12px"}}>TOTAL</td>
                            <td><span style={{fontFamily:"var(--mf)",fontWeight:900,fontSize:14,color:"var(--in-d)"}}>{totalUnidades}</span></td>
                            <td><span style={{fontFamily:"var(--mf)",fontWeight:900,fontSize:12,color:"var(--em-d)"}}>{fmtARS(totalVentas)}</span></td>
                            <td><span style={{fontFamily:"var(--mf)",fontWeight:900,fontSize:12,color:"var(--in-d)"}}>{fmtARS(gananciaTotal)}</span></td>
                          </tr>
                        </tbody>
                      </table></div>
                    }
                  </div>
                </div>
              </div>
            );
          })()}

          {/* ══ PRECIOS ══ */}
          {tab==="precios"&&(
            <div>
              <div className="ph"><div><div className="ph-h">Catálogo</div><div className="ph-s">Lista de precios con filtro</div></div></div>
              <div className="pc">
                {/* Search + price filter */}
                <SearchBar value={srchPrice} onChange={setSrchPrice} placeholder="Buscar por nombre, SKU o categoría..."/>
                <div className="card" style={{marginBottom:14}}>
                  <div className="card-h"><div className="card-title">💲 Filtrar por precio</div></div>
                  <div style={{padding:"12px 14px"}}>
                    <div className="row g8">
                      <div style={{flex:1}}>
                        <label className="fl">Precio mínimo ($)</label>
                        <input className="fi" type="number" placeholder="0" value={priceMin} onChange={function(e){setPriceMin(e.target.value);}}/>
                      </div>
                      <div style={{flex:1}}>
                        <label className="fl">Precio máximo ($)</label>
                        <input className="fi" type="number" placeholder="Sin límite" value={priceMax} onChange={function(e){setPriceMax(e.target.value);}}/>
                      </div>
                    </div>
                    {(priceMin||priceMax||srchPrice)&&(
                      <button className="btn btn-xs b-ghost" style={{marginTop:10}} onClick={function(){setPriceMin("");setPriceMax("");setSrchPrice("");}}>
                        Limpiar filtros
                      </button>
                    )}
                  </div>
                </div>

                {(function(){
                  var filtered = products.filter(function(p){
                    if (p.is_active===false) return false;
                    var q = srchPrice.toLowerCase();
                    if (q && !p.name.toLowerCase().includes(q) && !p.sku.toLowerCase().includes(q) && !(p.category||"").toLowerCase().includes(q)) return false;
                    if (priceMin && p.price < parseFloat(priceMin)) return false;
                    if (priceMax && p.price > parseFloat(priceMax)) return false;
                    return true;
                  });
                  return (
                    <div className="card">
                      <div className="card-h">
                        <div className="card-title">Productos ({filtered.length})</div>
                        {filtered.length<products.length&&<span className="badge b-in">filtrado</span>}
                      </div>
                      {filtered.length===0
                        ?<div className="empty">Sin productos con ese criterio.</div>
                        :(
                        <div className="tw"><table>
                          <thead><tr><th></th><th>SKU</th><th>Producto</th><th>Categoría</th><th>Precio</th></tr></thead>
                          <tbody>
                            {filtered.map(function(p){
                              var myInvRow = inventory.find(function(i){ return i.product_id===p.id; });
                              return (
                                <tr key={p.id} className="tr">
                                  <td><ProdThumb prod={p} size={34}/></td>
                                  <td><span style={{color:"var(--in-d)",fontFamily:"var(--mf)",fontSize:11,background:"var(--in-l)",padding:"2px 6px",borderRadius:5,fontWeight:600}}>{p.sku}</span></td>
                                  <td><div style={{fontWeight:600,fontSize:12}}>{p.name}</div></td>
                                  <td><span style={{fontSize:11,color:"var(--t3)"}}>{p.category}</span></td>
                                  <td>
                                    <div style={{fontFamily:"var(--mf)",fontWeight:800,fontSize:13,color:"var(--in-d)"}}>{fmtARS(p.price)}</div>
                                    {myInvRow&&myInvRow.qty_available>0&&(
                                      <div style={{fontSize:10,color:"var(--em-d)",fontWeight:700,marginTop:2}}>📦 {myInvRow.qty_available} en stock</div>
                                    )}
                                  </td>
                                </tr>
                              );
                            })}
                          </tbody>
                        </table></div>
                      )}
                    </div>
                  );
                })()}
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
                    <div className="card-h"><div className="card-title">🔔 Notificaciones</div><div className="row g8">{myNotifs.some(function(n){return !n.read;})&&<span className="badge b-am">{myNotifs.filter(function(n){return !n.read;}).length} nuevas</span>}<button className="btn btn-xs b-cr" onClick={async function(){ await sb.from("notifications").delete().eq("to_user_id",me.id); setNotifs([]); toast("Borradas","","i"); }}>🗑️ Borrar</button></div></div>
                    <div style={{padding:"8px 14px"}}>{myNotifs.map(function(n){
                      var ico=n.type==="transfer"?"📦":n.type==="confirm"?"✅":n.type==="sale"?"💰":"🔔";
                      var linkedTx=n.type==="transfer"?transfers.find(function(t){return t.to_user_id===me.id&&t.status==="pending";}):null;
                      return (
                        <div key={n.id} style={{display:"flex",gap:10,padding:"12px 0",borderBottom:"1px solid var(--brd)",opacity:n.read?0.6:1}}>
                          <div style={{fontSize:20,flexShrink:0}}>{ico}</div>
                          <div style={{flex:1}}>
                            <div style={{fontSize:13,fontWeight:n.read?400:700,color:"var(--t1)",lineHeight:1.4}}>{n.message}</div>
                            <div style={{fontSize:10,color:"var(--t3)",marginTop:3}}>{new Date(n.created_at).toLocaleString("es-AR")}</div>
                            {linkedTx&&(
                              <button style={{marginTop:8,padding:"8px 14px",borderRadius:10,border:"none",background:"var(--em)",color:"#fff",fontFamily:"var(--hf)",fontWeight:800,fontSize:12,cursor:"pointer",display:"flex",alignItems:"center",gap:6}} onClick={async function(){ await confirmTransfer(linkedTx); setNotifs(function(p){return p.map(function(x){return x.id===n.id?Object.assign({},x,{read:true}):x;});}); }}>
                                ✓ Confirmar recepción
                              </button>
                            )}
                          </div>
                          {!n.read&&<div style={{width:8,height:8,borderRadius:"50%",background:"var(--am)",flexShrink:0,marginTop:4}}/>}
                        </div>
                      );
                    })}</div>
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

          {/* ══ MI CUENTA ══ */}
          {tab==="cuenta"&&(
            <div>
              <div className="ph"><div><div className="ph-h">👤 Mi Cuenta</div><div className="ph-s">Tus datos personales</div></div></div>
              <div className="pc">
                <div className="card" style={{marginBottom:14}}>
                  <div style={{padding:"18px 16px",display:"flex",alignItems:"center",gap:14}}>
                    <Avatar name={me.name} color={me.color} size={56}/>
                    <div>
                      <div style={{fontSize:16,fontWeight:800,color:"var(--t1)"}}>{me.name}</div>
                      <div style={{fontSize:12,color:"var(--t3)"}}>{me.email}</div>
                      <div style={{marginTop:4}}>
                        <span style={{background:"var(--pri-l)",color:"var(--pri)",borderRadius:6,padding:"2px 9px",fontSize:10,fontWeight:800,textTransform:"uppercase"}}>
                          {me.role==="superadmin"?"👑 Superadmin":me.role==="empresaria"?"🏢 Empresaria":me.role==="lider"?"⭐ Líder":"🛍️ Vendedora"}
                        </span>
                      </div>
                    </div>
                  </div>
                </div>

                {(ctaJerInfo&&(ctaJerInfo.lider||ctaJerInfo.empresa))&&(
                  <div className="card" style={{marginBottom:14}}>
                    <div style={{padding:"14px 16px"}}>
                      <div style={{fontSize:11,fontWeight:800,color:"var(--t3)",textTransform:"uppercase",marginBottom:8}}>Mi jerarquía</div>
                      {ctaJerInfo.empresa&&<div style={{fontSize:13,marginBottom:4}}>🏢 Empresa: <b>{ctaJerInfo.empresa.name}</b></div>}
                      {ctaJerInfo.lider&&<div style={{fontSize:13,marginBottom:4}}>⭐ Líder: <b>{ctaJerInfo.lider.name}</b></div>}
                      {me.role==="reseller"&&<div style={{fontSize:13}}>💰 Comisión: <b>30% fijo</b></div>}
                      {me.role==="lider"&&<div style={{fontSize:13}}>💰 Mi comisión: <b>{me.comision_lider_pct||0}%</b></div>}
                    </div>
                  </div>
                )}

                <div className="card">
                  <div style={{padding:"16px"}}>
                    <div style={{fontSize:11,fontWeight:800,color:"var(--t3)",textTransform:"uppercase",marginBottom:10}}>Editar datos</div>
                    <label style={{fontSize:11,fontWeight:700,color:"var(--t3)"}}>Nombre</label>
                    <input value={ctaName} onChange={function(e){setCtaName(e.target.value);}} style={{width:"100%",boxSizing:"border-box",border:"1.5px solid var(--brd)",borderRadius:10,padding:"10px 12px",fontSize:14,marginTop:4,marginBottom:12,fontFamily:"inherit"}}/>
                    <label style={{fontSize:11,fontWeight:700,color:"var(--t3)"}}>Teléfono</label>
                    <input value={ctaTel} onChange={function(e){setCtaTel(e.target.value);}} placeholder="Opcional" style={{width:"100%",boxSizing:"border-box",border:"1.5px solid var(--brd)",borderRadius:10,padding:"10px 12px",fontSize:14,marginTop:4,marginBottom:12,fontFamily:"inherit"}}/>
                    <label style={{fontSize:11,fontWeight:700,color:"var(--t3)"}}>Dirección</label>
                    <input value={ctaDir} onChange={function(e){setCtaDir(e.target.value);}} placeholder="Opcional" style={{width:"100%",boxSizing:"border-box",border:"1.5px solid var(--brd)",borderRadius:10,padding:"10px 12px",fontSize:14,marginTop:4,marginBottom:16,fontFamily:"inherit"}}/>

                    {me.role==="reseller"&&(<>
                      <label style={{fontSize:11,fontWeight:700,color:"var(--t3)"}}>DNI</label>
                      <input value={ctaDni} onChange={function(e){setCtaDni(e.target.value);}} placeholder="Opcional" style={{width:"100%",boxSizing:"border-box",border:"1.5px solid var(--brd)",borderRadius:10,padding:"10px 12px",fontSize:14,marginTop:4,marginBottom:12,fontFamily:"inherit"}}/>
                      <label style={{fontSize:11,fontWeight:700,color:"var(--t3)"}}>Código de vendedora</label>
                      <input value={ctaCodVend} onChange={function(e){setCtaCodVend(e.target.value);}} placeholder="Opcional" style={{width:"100%",boxSizing:"border-box",border:"1.5px solid var(--brd)",borderRadius:10,padding:"10px 12px",fontSize:14,marginTop:4,marginBottom:12,fontFamily:"inherit"}}/>
                      <label style={{fontSize:11,fontWeight:700,color:"var(--t3)"}}>Localidad</label>
                      <input value={ctaLocalidad} onChange={function(e){setCtaLocalidad(e.target.value);}} placeholder="Opcional" style={{width:"100%",boxSizing:"border-box",border:"1.5px solid var(--brd)",borderRadius:10,padding:"10px 12px",fontSize:14,marginTop:4,marginBottom:16,fontFamily:"inherit"}}/>
                    </>)}

                    <div style={{fontSize:11,fontWeight:800,color:"var(--t3)",textTransform:"uppercase",marginBottom:10,borderTop:"1px solid var(--brd)",paddingTop:14}}>Cambiar contraseña (opcional)</div>
                    <label style={{fontSize:11,fontWeight:700,color:"var(--t3)"}}>Nueva contraseña</label>
                    <input type="password" value={ctaPass} onChange={function(e){setCtaPass(e.target.value);}} placeholder="Dejar vacío para no cambiar" style={{width:"100%",boxSizing:"border-box",border:"1.5px solid var(--brd)",borderRadius:10,padding:"10px 12px",fontSize:14,marginTop:4,marginBottom:12,fontFamily:"inherit"}}/>
                    {ctaPass&&(
                      <div style={{marginBottom:14}}>
                        <label style={{fontSize:11,fontWeight:700,color:"var(--t3)"}}>Repetir contraseña</label>
                        <input type="password" value={ctaPass2} onChange={function(e){setCtaPass2(e.target.value);}} style={{width:"100%",boxSizing:"border-box",border:"1.5px solid var(--brd)",borderRadius:10,padding:"10px 12px",fontSize:14,marginTop:4,fontFamily:"inherit"}}/>
                      </div>
                    )}
                    <button className="cta cta-in" onClick={doSaveMiCuenta} disabled={ctaSaving}>{ctaSaving?"Guardando...":"💾 Guardar cambios"}</button>
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* ══ PEDIDOS ESPECIALES (flujo en cascada) ══ */}
          {tab==="pedesp"&&(me.role==="reseller"||me.role==="lider"||me.role==="empresaria"||isAdmin)&&(
            <div>
              <div className="ph">
                <div><div className="ph-h">📦 Pedido Especial</div><div className="ph-s">Genera la orden de compra hacia la empresa</div></div>
                <div style={{display:"flex",gap:8}}>
                  <button className="btn btn-xs b-ghost" onClick={loadPedidosEspeciales}><Ic n="undo" s={13}/></button>
                  {me.role==="reseller"&&<button className="btn b-pri" onClick={function(){setPeShowForm(function(v){var nv=!v; if(!nv){setPeCarrito([]);setPeProdId("");setPeProdSrch("");}return nv;});}}><Ic n="plus" s={15}/>{peShowForm?"Cancelar":"Nuevo"}</button>}
                </div>
              </div>
              <div className="pc">

                {/* Formulario nuevo pedido especial (vendedora) — carrito de varios productos */}
                {peShowForm&&me.role==="reseller"&&(
                  <div className="card" style={{marginBottom:14}}>
                    <div style={{padding:"14px 16px"}}>
                      <div style={{fontSize:13,fontWeight:800,marginBottom:10}}>Nuevo pedido especial</div>

                      {/* Buscador + agregar producto al carrito */}
                      <SearchBar value={peProdSrch} onChange={setPeProdSrch} placeholder="Buscar producto..."/>
                      {peProdSrch&&(
                        <div style={{maxHeight:180,overflowY:"auto",border:"1px solid var(--brd)",borderRadius:10,marginTop:8}}>
                          {products.filter(function(p){var q=peProdSrch.toLowerCase();return p.name.toLowerCase().includes(q)||(p.sku||"").toLowerCase().includes(q);}).slice(0,8).map(function(p){
                            return (
                              <div key={p.id} onClick={function(){setPeProdId(p.id);setPeProdSrch(p.name+" ["+(p.sku||"")+"]");}} style={{padding:"10px 12px",cursor:"pointer",borderBottom:"1px solid var(--brd)",display:"flex",alignItems:"center",gap:8,background:peProdId===p.id?"var(--pri-l)":"var(--card)"}}>
                                <ProdThumb prod={p} size={32}/>
                                <div>
                                  <div style={{fontSize:13,fontWeight:700}}>{p.name}</div>
                                  <div style={{fontSize:11,color:"var(--t3)"}}><span style={{fontFamily:"var(--mf)",fontWeight:800,color:"var(--pri)",background:"var(--pri-l)",borderRadius:5,padding:"1px 6px",marginRight:6}}>{p.sku||"s/c"}</span>{fmtARS(p.price)}</div>
                                </div>
                              </div>
                            );
                          })}
                        </div>
                      )}
                      {peProdId&&(
                        <div style={{background:"var(--bg2)",borderRadius:10,padding:10,marginTop:8}}>
                          <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:8}}>
                            <span style={{fontSize:12,color:"var(--t3)"}}>Cantidad</span>
                            <button className="btn btn-xs b-ghost" onClick={function(){setPeQty(Math.max(1,peQty-1));}}>−</button>
                            <span style={{fontWeight:800,minWidth:20,textAlign:"center"}}>{peQty}</span>
                            <button className="btn btn-xs b-ghost" onClick={function(){setPeQty(peQty+1);}}>+</button>
                          </div>
                          <div style={{display:"flex",gap:8,marginBottom:8}}>
                            <input value={peColor} onChange={function(e){setPeColor(e.target.value);}} placeholder="Color (opcional)" style={{flex:1,boxSizing:"border-box",border:"1.5px solid var(--brd)",borderRadius:9,padding:"8px 10px",fontSize:12,fontFamily:"inherit"}}/>
                            <input value={peTalle} onChange={function(e){setPeTalle(e.target.value);}} placeholder="Talle (opcional)" style={{flex:1,boxSizing:"border-box",border:"1.5px solid var(--brd)",borderRadius:9,padding:"8px 10px",fontSize:12,fontFamily:"inherit"}}/>
                          </div>
                          <button className="btn btn-xs b-pri" style={{width:"100%",padding:"8px"}} onClick={doAgregarAlCarrito}>+ Agregar al pedido</button>
                        </div>
                      )}

                      {/* Carrito */}
                      {peCarrito.length>0&&(
                        <div style={{marginTop:12,marginBottom:4}}>
                          <div style={{fontSize:11,fontWeight:800,color:"var(--t3)",textTransform:"uppercase",marginBottom:6}}>🛒 {peCarrito.length} producto{peCarrito.length!==1?"s":""} en el pedido</div>
                          {peCarrito.map(function(it,idx){
                            return (
                              <div key={idx} style={{display:"flex",alignItems:"center",gap:8,padding:"7px 0",borderTop:"1px solid var(--brd)"}}>
                                <div style={{flex:1,minWidth:0}}>
                                  <div style={{fontSize:12,fontWeight:700}}>{it.qty}x {it.name}</div>
                                  <div style={{fontSize:10,color:"var(--t3)"}}>{[it.color,it.talle].filter(Boolean).join(" · ")||it.sku}</div>
                                </div>
                                <div style={{fontSize:12,fontWeight:700,color:"var(--pri)"}}>{fmtARS(it.price*it.qty)}</div>
                                <button onClick={function(){doQuitarDelCarrito(idx);}} style={{background:"none",border:"none",color:"var(--t3)",fontSize:14,cursor:"pointer",padding:"0 2px"}}>✕</button>
                              </div>
                            );
                          })}
                          <div style={{display:"flex",justifyContent:"space-between",fontWeight:800,fontSize:13,paddingTop:8,borderTop:"1.5px solid var(--brd)",marginTop:4}}>
                            <span>Total</span>
                            <span style={{color:"var(--pri)"}}>{fmtARS(peCarrito.reduce(function(s,it){return s+it.price*it.qty;},0))}</span>
                          </div>
                        </div>
                      )}

                      <div style={{marginTop:12}}>
                        <input value={peCliNom} onChange={function(e){setPeCliNom(e.target.value);}} placeholder="Nombre del cliente" style={{width:"100%",boxSizing:"border-box",border:"1.5px solid var(--brd)",borderRadius:10,padding:"10px 12px",fontSize:14,marginBottom:8,fontFamily:"inherit"}}/>
                        <input value={peCliTel} onChange={function(e){setPeCliTel(e.target.value);}} placeholder="Teléfono (opcional)" style={{width:"100%",boxSizing:"border-box",border:"1.5px solid var(--brd)",borderRadius:10,padding:"10px 12px",fontSize:14,marginBottom:8,fontFamily:"inherit"}}/>
                        <select value={peCampaniaId} onChange={function(e){setPeCampaniaId(e.target.value);}} style={{width:"100%",boxSizing:"border-box",border:"1.5px solid var(--brd)",borderRadius:10,padding:"10px 12px",fontSize:14,marginBottom:8,fontFamily:"inherit"}}>
                          <option value="">Sin campaña</option>
                          {campanias.filter(function(c){return c.empresa_id===me.empresa_id && c.estado==="abierta";}).map(function(c){
                            return <option key={c.id} value={c.id}>{c.nombre}</option>;
                          })}
                        </select>
                        <input value={peNota} onChange={function(e){setPeNota(e.target.value);}} placeholder="Nota general (opcional)" style={{width:"100%",boxSizing:"border-box",border:"1.5px solid var(--brd)",borderRadius:10,padding:"10px 12px",fontSize:14,marginBottom:12,fontFamily:"inherit"}}/>
                        <button className="cta cta-am" onClick={doCrearPedidoEspecial} disabled={peSaving||peCarrito.length===0}><Ic n="check" s={16}/>{peSaving?"Guardando...":"Guardar pedido ("+peCarrito.length+")"}</button>
                      </div>
                    </div>
                  </div>
                )}

                {/* Filtros (solo superadmin — agrupar por estructura comercial) */}
                {isAdmin&&pedEspList.length>0&&(function(){
                  var empresas = Array.from(new Set(pedEspList.map(function(p){return p.empresa?p.empresa.name:null;}).filter(Boolean)));
                  var lideres  = Array.from(new Set(pedEspList.map(function(p){return p.lider?p.lider.name:null;}).filter(Boolean)));
                  var vendedoras = Array.from(new Set(pedEspList.map(function(p){return p.vendedor?p.vendedor.name:null;}).filter(Boolean)));
                  var campaniasNombres = Array.from(new Set(pedEspList.map(function(p){return p.campania_rel?p.campania_rel.nombre:p.campania;}).filter(Boolean)));
                  return (
                    <div className="card" style={{marginBottom:14}}>
                      <div style={{padding:"12px 14px"}}>
                        <div style={{fontSize:11,fontWeight:800,color:"var(--t3)",textTransform:"uppercase",marginBottom:8}}>🔍 Filtrar / agrupar</div>
                        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
                          <select value={peFiltroEmp} onChange={function(e){setPeFiltroEmp(e.target.value);}} style={{fontSize:12,border:"1.5px solid var(--brd)",borderRadius:8,padding:"7px 8px",fontFamily:"inherit"}}>
                            <option value="">Todas las empresarias</option>
                            {empresas.map(function(n){return <option key={n} value={n}>{n}</option>;})}
                          </select>
                          <select value={peFiltroLid} onChange={function(e){setPeFiltroLid(e.target.value);}} style={{fontSize:12,border:"1.5px solid var(--brd)",borderRadius:8,padding:"7px 8px",fontFamily:"inherit"}}>
                            <option value="">Todos los líderes</option>
                            {lideres.map(function(n){return <option key={n} value={n}>{n}</option>;})}
                          </select>
                          <select value={peFiltroVend} onChange={function(e){setPeFiltroVend(e.target.value);}} style={{fontSize:12,border:"1.5px solid var(--brd)",borderRadius:8,padding:"7px 8px",fontFamily:"inherit"}}>
                            <option value="">Todas las vendedoras</option>
                            {vendedoras.map(function(n){return <option key={n} value={n}>{n}</option>;})}
                          </select>
                          <select value={peFiltroCamp} onChange={function(e){setPeFiltroCamp(e.target.value);}} style={{fontSize:12,border:"1.5px solid var(--brd)",borderRadius:8,padding:"7px 8px",fontFamily:"inherit"}}>
                            <option value="">Todas las campañas</option>
                            {campaniasNombres.map(function(n){return <option key={n} value={n}>{n}</option>;})}
                          </select>
                          <select value={peFiltroEst} onChange={function(e){setPeFiltroEst(e.target.value);}} style={{fontSize:12,border:"1.5px solid var(--brd)",borderRadius:8,padding:"7px 8px",fontFamily:"inherit",gridColumn:"span 2"}}>
                            <option value="">Todos los estados</option>
                            {["pendiente_lider","pendiente_empresaria","aprobado","enviado_proveedor","recibido","listo_entregar","entregado","rechazado_lider","rechazado_empresaria","cancelado"].map(function(s){return <option key={s} value={s}>{peEstadoInfo(s).lbl}</option>;})}
                          </select>
                        </div>
                        {(peFiltroEmp||peFiltroLid||peFiltroVend||peFiltroCamp||peFiltroEst)&&
                          <button className="btn btn-xs b-ghost" style={{marginTop:8}} onClick={function(){setPeFiltroEmp("");setPeFiltroLid("");setPeFiltroVend("");setPeFiltroCamp("");setPeFiltroEst("");}}>✕ Limpiar filtros</button>}
                      </div>
                    </div>
                  );
                })()}

                {/* Lista de pedidos */}
                {pedEspLoading&&<div className="empty">Cargando...</div>}
                {!pedEspLoading&&pedEspList.length===0&&<div className="empty">No hay pedidos especiales todavía</div>}
                {pedEspList.filter(function(p){
                  // Los borradores son privados: solo los ve la vendedora que los creó
                  if (p.estado==="borrador" && p.vendedor_id!==me.id) return false;
                  if (peFiltroEmp && (!p.empresa||p.empresa.name!==peFiltroEmp)) return false;
                  if (peFiltroLid && (!p.lider||p.lider.name!==peFiltroLid)) return false;
                  if (peFiltroVend && (!p.vendedor||p.vendedor.name!==peFiltroVend)) return false;
                  if (peFiltroCamp && (p.campania_rel?p.campania_rel.nombre:p.campania)!==peFiltroCamp) return false;
                  if (peFiltroEst && p.estado!==peFiltroEst) return false;
                  return true;
                }).map(function(p){
                  var ei = peEstadoInfo(p.estado);
                  var busy = !!peBusy[p.id];
                  var puedeVendEnviar = me.role==="reseller" && p.vendedor_id===me.id && p.estado==="borrador";
                  var puedeLider = me.role==="lider" && p.lider_id===me.id && p.estado==="pendiente_lider";
                  var puedeEmpAprobar   = me.role==="empresaria" && p.empresa_id===me.id && p.estado==="pendiente_empresaria";
                  var puedeEmpEnviar    = me.role==="empresaria" && p.empresa_id===me.id && p.estado==="aprobado";
                  var puedeEmpRecibir   = me.role==="empresaria" && p.empresa_id===me.id && p.estado==="enviado_proveedor";
                  var puedeEmpListo     = me.role==="empresaria" && p.empresa_id===me.id && p.estado==="recibido";
                  var puedeVendEntregar = me.role==="reseller" && p.vendedor_id===me.id && p.estado==="listo_entregar";
                  var puedeCancelar = (p.vendedor_id===me.id||p.lider_id===me.id||p.empresa_id===me.id) && !["entregado","cancelado"].includes(p.estado);
                  var necesitaAccion = puedeVendEnviar||puedeLider||puedeEmpAprobar||puedeEmpEnviar||puedeEmpRecibir||puedeEmpListo||puedeVendEntregar;
                  var items = p.items||[];
                  var primerItem = items[0];
                  return (
                    <div key={p.id} className="card" style={{marginBottom:10,border:necesitaAccion?"1.5px solid var(--am-d)":undefined}}>
                      <div style={{padding:"12px 14px"}}>
                        <div style={{display:"flex",alignItems:"center",gap:10}}>
                          <ProdThumb prod={primerItem?primerItem.product:null} size={38}/>
                          <div style={{flex:1,minWidth:0}}>
                            <div style={{fontSize:13,fontWeight:700}}>
                              {primerItem?primerItem.product.name:"Pedido"}{items.length>1?" +"+(items.length-1)+" más":""}
                            </div>
                            <div style={{fontSize:11,color:"var(--t3)"}}>{p.qty} un. en {items.length} producto{items.length!==1?"s":""} · {fmtARS(p.total)} · Cliente: {p.cliente_nombre}</div>
                            <div style={{fontSize:10,color:"var(--t3)"}}>Vendedora: {p.vendedor?p.vendedor.name:"-"}{p.lider?" · Líder: "+p.lider.name:""}</div>
                          </div>
                          <span style={{background:ei.bg,color:ei.col,borderRadius:8,padding:"4px 9px",fontSize:10,fontWeight:800,whiteSpace:"nowrap"}}>{ei.lbl}</span>
                        </div>

                        {/* Lista de productos del pedido */}
                        {items.length>0&&(
                          <div style={{marginTop:8,background:"var(--bg2)",borderRadius:9,padding:"8px 10px"}}>
                            {items.map(function(it){
                              return (
                                <div key={it.id} style={{display:"flex",justifyContent:"space-between",fontSize:11,color:"var(--t2)",padding:"2px 0"}}>
                                  <span>{it.qty}x {it.product?it.product.name:"-"}{[it.color,it.talle].filter(Boolean).length>0?" ("+[it.color,it.talle].filter(Boolean).join(", ")+")":""}</span>
                                  <span style={{fontWeight:700}}>{fmtARS(it.subtotal)}</span>
                                </div>
                              );
                            })}
                          </div>
                        )}

                        {p.nota&&<div style={{fontSize:11,color:"var(--t3)",marginTop:6,fontStyle:"italic"}}>📝 {p.nota}</div>}

                        {necesitaAccion&&(
                          <div style={{marginTop:10,paddingTop:10,borderTop:"1px solid var(--brd)"}}>
                            <input value={peObserv[p.id]||""} onChange={function(e){setPeObserv(function(prev){return Object.assign({},prev,{[p.id]:e.target.value});});}}
                              placeholder="Observación (opcional)" style={{width:"100%",boxSizing:"border-box",border:"1.5px solid var(--brd)",borderRadius:9,padding:"8px 10px",fontSize:12,marginBottom:8,fontFamily:"inherit"}}/>
                            <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                              {puedeVendEnviar&&<button className="btn btn-xs b-pri" style={{padding:"7px 12px"}} disabled={busy} onClick={function(){doEnviarPedidoEspecial(p.id);}}>📤 Enviar a aprobación</button>}
                              {puedeLider&&(<>
                                <button className="btn btn-xs" style={{background:"#e7f9ee",color:"#0a8f4d",border:"1px solid #bfe9d2",borderRadius:8,fontWeight:700,padding:"7px 12px"}} disabled={busy} onClick={function(){doAccionPedidoEsp(p.id,"rpc_lider_decidir_pedido",{p_aprobar:true});}}>✅ Aprobar</button>
                                <button className="btn btn-xs b-cr" style={{padding:"7px 12px"}} disabled={busy} onClick={function(){doAccionPedidoEsp(p.id,"rpc_lider_decidir_pedido",{p_aprobar:false});}}>✕ Rechazar</button>
                              </>)}
                              {puedeEmpAprobar&&(<>
                                <button className="btn btn-xs" style={{background:"#e7f9ee",color:"#0a8f4d",border:"1px solid #bfe9d2",borderRadius:8,fontWeight:700,padding:"7px 12px"}} disabled={busy} onClick={function(){doAccionPedidoEsp(p.id,"rpc_empresaria_decidir_pedido",{p_aprobar:true});}}>✅ Aprobar</button>
                                <button className="btn btn-xs b-cr" style={{padding:"7px 12px"}} disabled={busy} onClick={function(){doAccionPedidoEsp(p.id,"rpc_empresaria_decidir_pedido",{p_aprobar:false});}}>✕ Rechazar</button>
                              </>)}
                              {puedeEmpEnviar&&<button className="btn btn-xs b-pri" style={{padding:"7px 12px"}} disabled={busy} onClick={function(){doAccionPedidoEsp(p.id,"rpc_pedido_enviado_proveedor");}}>📤 Enviar a proveedor</button>}
                              {puedeEmpRecibir&&<button className="btn btn-xs b-pri" style={{padding:"7px 12px"}} disabled={busy} onClick={function(){doAccionPedidoEsp(p.id,"rpc_pedido_recibido");}}>📦 Marcar recibido</button>}
                              {puedeEmpListo&&<button className="btn btn-xs b-pri" style={{padding:"7px 12px"}} disabled={busy} onClick={function(){doAccionPedidoEsp(p.id,"rpc_pedido_listo_entregar");}}>✅ Listo para entregar</button>}
                              {puedeVendEntregar&&<button className="btn btn-xs" style={{background:"#e7f9ee",color:"#0a8f4d",border:"1px solid #bfe9d2",borderRadius:8,fontWeight:700,padding:"7px 12px"}} disabled={busy} onClick={function(){doAccionPedidoEsp(p.id,"rpc_pedido_entregado");}}>🎉 Entregar al cliente</button>}
                            </div>
                          </div>
                        )}

                        <div style={{display:"flex",gap:12,marginTop:8,flexWrap:"wrap",alignItems:"center"}}>
                          <button onClick={function(){verHistorialPedidoEsp(p.id);}} style={{background:"none",border:"none",color:"var(--t3)",fontSize:11,fontWeight:700,cursor:"pointer",padding:0}}>{peHistOpen===p.id?"▲ Ocultar historial":"▼ Ver historial"}</button>
                          {puedeCancelar&&<button onClick={function(){doAccionPedidoEsp(p.id,"rpc_pedido_cancelar");}} disabled={busy} style={{background:"none",border:"none",color:"var(--cr,#d32)",fontSize:11,fontWeight:700,cursor:"pointer",padding:0}}>Cancelar pedido</button>}
                          {p.pdf_url&&<a href={p.pdf_url} target="_blank" rel="noreferrer" style={{color:"var(--pri)",fontSize:11,fontWeight:700,textDecoration:"none"}}>📄 Descargar PDF</a>}
                          {p.pdf_url&&isAdmin&&<button onClick={function(){compartirPdfPorEmail(p);}} style={{background:"none",border:"none",color:"var(--bl-d,#0369a1)",fontSize:11,fontWeight:700,cursor:"pointer",padding:0}}>✉️ Enviar por email</button>}
                        </div>

                        {peHistOpen===p.id&&(
                          <div style={{marginTop:8,background:"var(--bg2)",borderRadius:10,padding:"10px 12px"}}>
                            {peHist.length===0&&<div style={{fontSize:11,color:"var(--t3)"}}>Sin historial</div>}
                            {peHist.map(function(h){
                              return (
                                <div key={h.id} style={{fontSize:11,color:"var(--t2)",marginBottom:6}}>
                                  <b>{peEstadoInfo(h.estado_nuevo).lbl}</b> — {h.usuario?h.usuario.name:"Sistema"} · {new Date(h.created_at).toLocaleString("es-AR",{day:"2-digit",month:"2-digit",hour:"2-digit",minute:"2-digit"})}
                                  {h.observaciones&&<div style={{fontStyle:"italic",color:"var(--t3)"}}>"{h.observaciones}"</div>}
                                </div>
                              );
                            })}
                          </div>
                        )}
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          )}

          {/* ══ RESUMEN (pedidos + ventas + comisiones) ══ */}
          {tab==="resumen"&&(me.role==="lider"||me.role==="empresaria")&&(
            <div>
              <div className="ph">
                <div><div className="ph-h">{me.role==="lider"?"📋 Mi Resumen":"📊 Resumen General"}</div><div className="ph-s">Este mes</div></div>
                <button className="btn btn-xs b-ghost" onClick={loadResumen}><Ic n="undo" s={13}/>Actualizar</button>
              </div>
              <div className="pc">
                {resumenLoading&&<div className="empty">Cargando...</div>}
                {!resumenLoading&&resumen&&(
                  <>
                    {/* Pedidos pendientes de mi acción */}
                    <div className="card" style={{marginBottom:12,background:resumen.pedidos_pendientes>0?"var(--am-l,#fff3e0)":undefined}}>
                      <div style={{padding:"16px",display:"flex",alignItems:"center",gap:14}}>
                        <div style={{fontSize:32}}>{resumen.pedidos_pendientes>0?"🔔":"✅"}</div>
                        <div style={{flex:1}}>
                          <div style={{fontSize:24,fontWeight:800,color:resumen.pedidos_pendientes>0?"var(--am-d,#e07800)":"var(--em-d,#0a8f4d)"}}>{resumen.pedidos_pendientes}</div>
                          <div style={{fontSize:12,color:"var(--t3)"}}>Pedidos pendientes de tu aprobación</div>
                        </div>
                        {resumen.pedidos_pendientes>0&&<button className="btn btn-xs b-pri" onClick={function(){setTab("pedesp");}}>Ver →</button>}
                      </div>
                    </div>

                    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:12}}>
                      <div className="card"><div style={{padding:"14px"}}>
                        <div style={{fontSize:20,fontWeight:800,color:"var(--pri)"}}>{resumen.pedidos_en_curso}</div>
                        <div style={{fontSize:11,color:"var(--t3)"}}>Pedidos en curso</div>
                      </div></div>
                      <div className="card"><div style={{padding:"14px"}}>
                        <div style={{fontSize:20,fontWeight:800,color:"var(--bl-d)"}}>{resumen.unidades_mes}</div>
                        <div style={{fontSize:11,color:"var(--t3)"}}>Unidades vendidas</div>
                      </div></div>
                    </div>

                    <div className="card" style={{marginBottom:12}}>
                      <div style={{padding:"16px"}}>
                        <div style={{fontSize:11,fontWeight:800,color:"var(--t3)",textTransform:"uppercase",marginBottom:6}}>💰 Ventas del mes</div>
                        <div style={{fontSize:26,fontWeight:800,color:"var(--em-d,#0a8f4d)"}}>${Number(resumen.ventas_mes||0).toLocaleString("es-AR")}</div>
                      </div>
                    </div>

                    {me.role==="lider"&&(
                      <div className="card" style={{marginBottom:12}}>
                        <div style={{padding:"16px"}}>
                          <div style={{fontSize:11,fontWeight:800,color:"var(--t3)",textTransform:"uppercase",marginBottom:6}}>💵 Mi comisión del mes</div>
                          <div style={{fontSize:26,fontWeight:800,color:"var(--pri)"}}>${Number(resumen.comision_mes||0).toLocaleString("es-AR")}</div>
                          <div style={{fontSize:11,color:"var(--t3)",marginTop:4}}>👥 {resumen.equipo_count} vendedora(s) a tu cargo</div>
                        </div>
                      </div>
                    )}

                    {me.role==="empresaria"&&(
                      <>
                        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:12}}>
                          <div className="card"><div style={{padding:"14px"}}>
                            <div style={{fontSize:16,fontWeight:800,color:"var(--pri)"}}>${Number(resumen.comision_lideres_mes||0).toLocaleString("es-AR")}</div>
                            <div style={{fontSize:10,color:"var(--t3)"}}>Comisiones a líderes</div>
                          </div></div>
                          <div className="card"><div style={{padding:"14px"}}>
                            <div style={{fontSize:16,fontWeight:800,color:"var(--pri)"}}>${Number(resumen.comision_vendedoras_mes||0).toLocaleString("es-AR")}</div>
                            <div style={{fontSize:10,color:"var(--t3)"}}>Comisiones a vendedoras</div>
                          </div></div>
                        </div>
                        <div className="card" style={{marginBottom:12}}>
                          <div style={{padding:"16px"}}>
                            <div style={{fontSize:11,fontWeight:800,color:"var(--t3)",textTransform:"uppercase",marginBottom:6}}>📈 Rentabilidad del mes</div>
                            <div style={{fontSize:26,fontWeight:800,color:resumen.rentabilidad_mes>=0?"var(--em-d,#0a8f4d)":"var(--cr,#d32)"}}>${Number(resumen.rentabilidad_mes||0).toLocaleString("es-AR")}</div>
                            <div style={{fontSize:11,color:"var(--t3)",marginTop:4}}>⭐ {resumen.lideres_count} líder(es) · 🛍️ {resumen.vendedoras_count} vendedora(s)</div>
                          </div>
                        </div>
                      </>
                    )}

                    <button className="cta" style={{background:"#25D366",color:"#fff"}} onClick={compartirResumenWhatsapp}>
                      📲 Compartir resumen por WhatsApp
                    </button>
                  </>
                )}
              </div>
            </div>
          )}

          {/* ══ MI EQUIPO ══ */}
          {tab==="equipo"&&(me.role==="empresaria"||me.role==="lider"||isAdmin)&&(
            <div>
              <div className="ph"><div><div className="ph-h">{isAdmin?"🏢 Empresas":"👥 Mi Estructura"}</div><div className="ph-s">{me.role==="empresaria"?"Tus líderes y vendedoras":me.role==="lider"?"Tus vendedoras":"Empresarias del sistema"}</div></div><button className="btn btn-xs b-ghost" onClick={loadJerarquia}><Ic n="undo" s={13}/>Actualizar</button></div>
              <div className="pc">

                {/* Campañas (empresaria: crear, abrir/cerrar — duran 30 días) */}
                {me.role==="empresaria"&&(
                  <div className="card" style={{marginBottom:14}}>
                    <div style={{padding:"14px 16px"}}>
                      <div style={{fontSize:13,fontWeight:800,marginBottom:10}}>📅 Campañas <span style={{fontWeight:400,color:"var(--t3)",fontSize:11}}>(duran 30 días)</span></div>
                      <div style={{display:"flex",gap:8,marginBottom:12}}>
                        <input value={campNombre} onChange={function(e){setCampNombre(e.target.value);}} placeholder="Nombre de la campaña" style={{flex:1,boxSizing:"border-box",border:"1.5px solid var(--brd)",borderRadius:10,padding:"9px 12px",fontSize:13,fontFamily:"inherit"}}/>
                        <button className="btn btn-xs b-pri" style={{padding:"0 14px"}} onClick={doCrearCampania} disabled={campSaving}>{campSaving?"...":"+ Crear"}</button>
                      </div>
                      {campanias.filter(function(c){return c.empresa_id===me.id;}).length===0&&<div style={{fontSize:12,color:"var(--t3)"}}>Sin campañas todavía</div>}
                      {campanias.filter(function(c){return c.empresa_id===me.id;}).map(function(c){
                        return (
                          <div key={c.id} style={{display:"flex",alignItems:"center",gap:8,padding:"8px 0",borderTop:"1px solid var(--brd)"}}>
                            <div style={{flex:1,minWidth:0}}>
                              <div style={{fontSize:13,fontWeight:700}}>{c.nombre}</div>
                              <div style={{fontSize:10,color:"var(--t3)"}}>{c.fecha_inicio} → {c.fecha_fin}</div>
                            </div>
                            <span style={{background:c.estado==="abierta"?"#dcfce7":"#f1f1f1",color:c.estado==="abierta"?"#15803d":"#888",borderRadius:6,padding:"2px 8px",fontSize:9,fontWeight:800,textTransform:"uppercase"}}>{c.estado}</span>
                            <button className="btn btn-xs b-ghost" style={{padding:"5px 9px",fontSize:10}} onClick={function(){doCambiarEstadoCampania(c.id, c.estado==="abierta"?"cerrada":"abierta");}}>
                              {c.estado==="abierta"?"Cerrar":"Reabrir"}
                            </button>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                )}

                {/* Invitar (empresaria invita líder/vendedora; superadmin invita empresaria) */}
                {(me.role==="empresaria"||isAdmin)&&(
                  <div className="card" style={{marginBottom:14}}>
                    <div style={{padding:"14px 16px"}}>
                      <div style={{fontSize:13,fontWeight:800,marginBottom:10}}>➕ Invitar {isAdmin?"empresaria":"a mi equipo"}</div>
                      <input value={invName} onChange={function(e){setInvName(e.target.value);}} placeholder="Nombre completo" style={{width:"100%",boxSizing:"border-box",border:"1.5px solid var(--brd)",borderRadius:10,padding:"10px 12px",fontSize:14,marginBottom:8,fontFamily:"inherit"}}/>
                      <input value={invEmail} onChange={function(e){setInvEmail(e.target.value);}} placeholder="Email" style={{width:"100%",boxSizing:"border-box",border:"1.5px solid var(--brd)",borderRadius:10,padding:"10px 12px",fontSize:14,marginBottom:8,fontFamily:"inherit"}}/>
                      {!isAdmin&&(
                        <div style={{display:"flex",gap:8,marginBottom:8}}>
                          <button onClick={function(){setInvRole("reseller");}} style={{flex:1,padding:"9px",borderRadius:9,border:"1.5px solid "+(invRole==="reseller"?"var(--pri)":"var(--brd)"),background:invRole==="reseller"?"var(--pri-l)":"#fff",fontWeight:700,fontSize:12,cursor:"pointer",color:invRole==="reseller"?"var(--pri)":"var(--t2)"}}>🛍️ Vendedora</button>
                          <button onClick={function(){setInvRole("lider");}} style={{flex:1,padding:"9px",borderRadius:9,border:"1.5px solid "+(invRole==="lider"?"var(--pri)":"var(--brd)"),background:invRole==="lider"?"var(--pri-l)":"#fff",fontWeight:700,fontSize:12,cursor:"pointer",color:invRole==="lider"?"var(--pri)":"var(--t2)"}}>⭐ Líder</button>
                        </div>
                      )}
                      {(!isAdmin&&invRole==="reseller"&&miEquipo.filter(function(u){return u.role==="lider";}).length>0)&&(
                        <select value={invLiderId} onChange={function(e){setInvLiderId(e.target.value);}} style={{width:"100%",boxSizing:"border-box",border:"1.5px solid var(--brd)",borderRadius:10,padding:"10px 12px",fontSize:14,marginBottom:8,fontFamily:"inherit"}}>
                          <option value="">Sin líder (directo)</option>
                          {miEquipo.filter(function(u){return u.role==="lider";}).map(function(l){return <option key={l.id} value={l.id}>{l.name}</option>;})}
                        </select>
                      )}
                      {(!isAdmin&&invRole==="lider")&&(
                        <div style={{marginBottom:8}}>
                          <label style={{fontSize:11,fontWeight:700,color:"var(--t3)"}}>% de comisión para este líder</label>
                          <input type="number" value={invComis} onChange={function(e){setInvComis(e.target.value);}} style={{width:"100%",boxSizing:"border-box",border:"1.5px solid var(--brd)",borderRadius:10,padding:"10px 12px",fontSize:14,marginTop:4,fontFamily:"inherit"}}/>
                        </div>
                      )}
                      <button className="cta cta-in" onClick={doInvitar} disabled={invSaving}>{invSaving?"Creando...":"✉️ Crear invitación"}</button>
                    </div>
                  </div>
                )}

                {/* Invitaciones pendientes */}
                {misInvites.length>0&&(
                  <div className="card" style={{marginBottom:14}}>
                    <div style={{padding:"14px 16px"}}>
                      <div style={{fontSize:13,fontWeight:800,marginBottom:10}}>⏳ Invitaciones pendientes</div>
                      {misInvites.map(function(inv){
                        var link = window.location.origin + "?invite=" + encodeURIComponent(inv.email);
                        return (
                          <div key={inv.id} style={{display:"flex",alignItems:"center",gap:8,padding:"8px 0",borderTop:"1px solid var(--brd)"}}>
                            <div style={{flex:1,minWidth:0}}>
                              <div style={{fontSize:13,fontWeight:700}}>{inv.name}</div>
                              <div style={{fontSize:11,color:"var(--t3)"}}>{inv.email} · {inv.role}</div>
                            </div>
                            <button className="btn btn-xs" style={{background:"#e7f9ee",color:"#0a8f4d",border:"1px solid #bfe9d2",borderRadius:8,fontSize:10,fontWeight:700,padding:"6px 8px"}} onClick={function(){
                              var msg = "¡Hola "+inv.name+"! Te invito a StockPro. Registrate con este email ("+inv.email+") en: "+window.location.origin;
                              window.open("https://wa.me/?text="+encodeURIComponent(msg),"_blank");
                            }}>📲</button>
                            <button className="btn btn-xs b-cr" style={{padding:"6px 8px"}} onClick={function(){doCancelarInvite(inv.id);}}><Ic n="trash" s={11}/></button>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                )}

                {/* Lista de equipo */}
                <div className="card">
                  <div style={{padding:"14px 16px"}}>
                    <div style={{fontSize:13,fontWeight:800,marginBottom:10}}>{isAdmin?"Empresarias":"Miembros"} ({miEquipo.length})</div>
                    {miEquipo.length===0&&<div className="empty">Todavía no hay nadie en tu equipo</div>}
                    {miEquipo.map(function(u){
                      return (
                        <div key={u.id} style={{padding:"10px 0",borderTop:"1px solid var(--brd)"}}>
                          <div style={{display:"flex",alignItems:"center",gap:10}}>
                            <Avatar name={u.name} color={u.color} size={34}/>
                            <div style={{flex:1,minWidth:0}}>
                              <div style={{fontSize:13,fontWeight:700}}>{u.name}</div>
                              <div style={{fontSize:11,color:"var(--t3)"}}>{u.email}</div>
                            </div>
                            {(isAdmin || (me.role==="empresaria" && u.role!=="empresaria")) ? (
                              <select value={u.role} onChange={function(e){doCambiarRol(u.id, e.target.value);}}
                                style={{fontSize:10,fontWeight:800,textTransform:"uppercase",border:"1.5px solid var(--pri-l)",background:"var(--pri-l)",color:"var(--pri)",borderRadius:6,padding:"3px 6px",fontFamily:"inherit",cursor:"pointer"}}>
                                {isAdmin&&<option value="empresaria">🏢 Empresaria</option>}
                                <option value="lider">⭐ Líder</option>
                                <option value="reseller">🛍️ Vendedora</option>
                              </select>
                            ) : (
                              <span style={{background:"var(--pri-l)",color:"var(--pri)",borderRadius:6,padding:"2px 8px",fontSize:9,fontWeight:800,textTransform:"uppercase",whiteSpace:"nowrap"}}>
                                {u.role==="empresaria"?"🏢 Empresaria":u.role==="lider"?"⭐ Líder":"🛍️ Vendedora"}
                              </span>
                            )}
                          </div>
                          {/* Asignar líder a una vendedora (solo empresaria) */}
                          {(me.role==="empresaria"&&u.role==="reseller")&&(
                            <div style={{marginTop:8,marginLeft:44}}>
                              <select value={u.lider_id||""} onChange={function(e){doAsignarLider(u.id, e.target.value);}} style={{fontSize:12,border:"1.5px solid var(--brd)",borderRadius:8,padding:"6px 10px",fontFamily:"inherit"}}>
                                <option value="">Sin líder (directo)</option>
                                {miEquipo.filter(function(l){return l.role==="lider";}).map(function(l){return <option key={l.id} value={l.id}>{l.name}</option>;})}
                              </select>
                            </div>
                          )}
                          {/* Editar % comisión de un líder (solo empresaria) */}
                          {(me.role==="empresaria"&&u.role==="lider")&&(
                            <div style={{marginTop:8,marginLeft:44,display:"flex",gap:6,alignItems:"center"}}>
                              <input type="number" value={equipoEditComis[u.id]!==undefined?equipoEditComis[u.id]:(u.comision_lider_pct||0)}
                                onChange={function(e){setEquipoEditComis(function(prev){return Object.assign({},prev,{[u.id]:e.target.value});});}}
                                style={{width:64,fontSize:12,border:"1.5px solid var(--brd)",borderRadius:8,padding:"6px 8px",fontFamily:"inherit"}}/>
                              <span style={{fontSize:12,color:"var(--t3)"}}>% comisión</span>
                              <button className="btn btn-xs b-pri" style={{padding:"5px 10px",fontSize:10}} onClick={function(){doGuardarComisionLider(u.id);}}>Guardar</button>
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
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
                {/* ── GESTOR DE CARRUSELES ── */}
                <div className="card" style={{marginBottom:12}}>
                  <div className="card-h">
                    <div className="card-title"><div className="card-ico" style={{background:"var(--in-l)",color:"var(--in)"}}><Ic n="img" s={14}/></div>🎠 Carruseles de Ofertas</div>
                    <button className="btn btn-xs b-in" onClick={function(){setCarouselEdit(function(v){return !v;});setCEditId(null);setCTitle("");setCSubtitle("");setCBg("#e0224e");setCEmoji("🔥");setCLink("");setCImg("");}}>
                      {carouselEdit?"✕ Cerrar":"+ Nuevo"}
                    </button>
                  </div>
                  {carouselEdit&&(
                    <div style={{padding:"14px 16px",borderBottom:"1px solid var(--brd)",background:"var(--bg)"}}>
                      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}>
                        <div>
                          <label className="fl">Emoji</label>
                          <select className="fi fi-sel" value={cEmoji} onChange={function(e){setCEmoji(e.target.value);}}>
                            {["🔥","⭐","💥","🎉","🛍️","💄","🧴","💎","🌸","✨","📦","🏷️","💰","🎀","📲","🆕"].map(function(e){ return <option key={e}>{e}</option>; })}
                          </select>
                        </div>
                        <div>
                          <label className="fl">Color de fondo</label>
                          <input type="color" value={cBg} onChange={function(e){setCBg(e.target.value);}} style={{width:"100%",height:44,border:"1.5px solid var(--brd)",borderRadius:12,cursor:"pointer",padding:"2px 4px"}}/>
                        </div>
                      </div>
                      <div className="fld">
                        <label className="fl">Título *</label>
                        <input className="fi" placeholder="Ej: ¡25% de descuento esta semana!" value={cTitle} onChange={function(e){setCTitle(e.target.value);}}/>
                      </div>
                      <div className="fld">
                        <label className="fl">Subtítulo</label>
                        <input className="fi" placeholder="Ej: En productos seleccionados" value={cSubtitle} onChange={function(e){setCSubtitle(e.target.value);}}/>
                      </div>
                      <div className="fld">
                        <label className="fl">Tab al tocar (opcional)</label>
                        <select className="fi fi-sel" value={cLink} onChange={function(e){setCLink(e.target.value);}}>
                          <option value="">Sin acción</option>
                          <option value="catalog">Catálogo</option>
                          <option value="enviados">Enviados</option>
                          <option value="recibidos">Recibidos</option>
                          <option value="ventas">Ventas</option>
                          <option value="pedidos">Pedidos</option>
                        </select>
                      </div>
                      {/* Imagen del slide */}
                      <div className="fld">
                        <label className="fl">Imagen del producto (opcional)</label>
                        {cImg?(
                          <div style={{display:"flex",alignItems:"center",gap:10}}>
                            <img src={cImg} alt="" style={{width:60,height:60,borderRadius:12,objectFit:"cover",border:"1.5px solid var(--brd)"}}/>
                            <button type="button" onClick={function(){setCImg("");}} style={{fontSize:12,color:"var(--cr)",background:"none",border:"none",cursor:"pointer",fontWeight:700}}>Quitar imagen</button>
                          </div>
                        ):(
                          <label style={{display:"block",background:"var(--in-l)",border:"1.5px dashed var(--in-m)",borderRadius:12,padding:"12px",textAlign:"center",cursor:"pointer"}}>
                            <input type="file" accept="image/*" style={{display:"none"}} onChange={function(e){var f=e.target.files&&e.target.files[0];uploadCarouselImg(f);e.target.value="";}}/>
                            {cImgUp?<div style={{fontSize:12,color:"var(--in-d)",fontWeight:700}}>Subiendo...</div>
                              :<div><div style={{fontSize:20,marginBottom:2}}>📸</div><div style={{fontSize:12,fontWeight:700,color:"var(--in-d)"}}>Subir imagen</div></div>}
                          </label>
                        )}
                      </div>
                      {/* Preview */}
                      <div style={{borderRadius:14,overflow:"hidden",marginBottom:12,height:120,background:cBg,display:"flex",alignItems:"center",padding:"14px 18px",gap:14,position:"relative"}}>
                        {!cImg&&<div style={{fontSize:44}}>{cEmoji}</div>}
                        <div style={{flex:1,minWidth:0,position:"relative",zIndex:2}}>
                          <div style={{fontSize:17,fontWeight:900,color:"#fff",lineHeight:1.15,textShadow:cImg?"0 1px 6px rgba(0,0,0,.4)":"none"}}>{cTitle||"Título del carrusel"}</div>
                          {cSubtitle&&<div style={{fontSize:11,color:"rgba(255,255,255,.9)",marginTop:4,fontWeight:600,textShadow:cImg?"0 1px 4px rgba(0,0,0,.4)":"none"}}>{cSubtitle}</div>}
                        </div>
                        {cImg&&<img src={cImg} alt="" style={{height:"120%",maxWidth:"44%",objectFit:"contain",zIndex:2,filter:"drop-shadow(0 6px 12px rgba(0,0,0,.3))"}}/>}
                        {cImg&&<div style={{position:"absolute",inset:0,background:"linear-gradient(90deg,"+cBg+" 35%,transparent 100%)",zIndex:1}}/>}
                      </div>
                      <button className="cta cta-in" style={{marginTop:0}} onClick={saveCarousel} disabled={cSaving||!cTitle.trim()}>
                        {cSaving?"Guardando...":cEditId?"💾 Guardar cambios":"✅ Crear carrusel"}
                      </button>
                    </div>
                  )}
                  {carousels.length===0&&!carouselEdit&&(
                    <div className="empty">Sin carruseles. Creá el primero con el botón "Nuevo".</div>
                  )}
                  {carousels.map(function(sl){
                    return (
                      <div key={sl.id} style={{display:"flex",alignItems:"center",gap:10,padding:"12px 14px",borderBottom:"1px solid var(--brd)"}}>
                        <div style={{width:52,height:52,borderRadius:12,background:sl.bg_color||"#cc0000",display:"flex",alignItems:"center",justifyContent:"center",fontSize:24,flexShrink:0}}>
                          {sl.emoji||"🔥"}
                        </div>
                        <div style={{flex:1,minWidth:0}}>
                          <div style={{fontWeight:800,fontSize:13,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{sl.title}</div>
                          {sl.subtitle&&<div style={{fontSize:11,color:"var(--t3)",marginTop:1}}>{sl.subtitle}</div>}
                          <div style={{fontSize:10,color:sl.active?"var(--em-d)":"var(--t4)",fontWeight:700,marginTop:2}}>{sl.active?"● Activo":"○ Inactivo"}</div>
                        </div>
                        <div style={{display:"flex",gap:6,flexShrink:0}}>
                          <button className="btn btn-xs b-ghost" style={{padding:"5px 7px"}} onClick={function(){moveCarousel(sl,-1);}} title="Subir">▲</button>
                          <button className="btn btn-xs b-ghost" style={{padding:"5px 7px"}} onClick={function(){moveCarousel(sl,1);}} title="Bajar">▼</button>
                          <button className="btn btn-xs b-ghost" onClick={function(){
                            setCEditId(sl.id);setCTitle(sl.title);setCSubtitle(sl.subtitle||"");setCBg(sl.bg_color||"#e0224e");setCEmoji(sl.emoji||"🔥");setCLink(sl.link_tab||"");setCImg(sl.image_url||"");setCarouselEdit(true);
                          }}><Ic n="edit" s={11}/></button>
                          <button className="btn btn-xs" style={{background:sl.active?"var(--am-l)":"var(--em-l)",color:sl.active?"var(--am-d)":"var(--em-d)",border:"1px solid var(--brd)",borderRadius:8,padding:"5px 9px",fontSize:10,fontWeight:700,cursor:"pointer"}} onClick={function(){toggleCarouselActive(sl);}}>
                            {sl.active?"Pausar":"Activar"}
                          </button>
                          {cDelConf===sl.id
                            ?<div style={{display:"flex",gap:5}}>
                              <button className="btn btn-xs b-cr" onClick={function(){deleteCarousel(sl.id);}}>Sí</button>
                              <button className="btn btn-xs b-ghost" onClick={function(){setCDelConf(null);}}>No</button>
                            </div>
                            :<button className="btn btn-xs b-cr" onClick={function(){setCDelConf(sl.id);}}><Ic n="trash" s={11}/></button>
                          }
                        </div>
                      </div>
                    );
                  })}
                </div>

                <div className="card">
                  <div className="card-h"><div className="card-title"><Ic n="users" s={14}/>Usuarios registrados ({allUsers.length})</div></div>
                  {allUsers.length===0?<div className="empty">Sin usuarios aún</div>:(
                    <div className="tw"><table><thead><tr><th>Nombre</th><th>Email</th><th>Rol</th><th>Registro</th><th></th></tr></thead><tbody>{allUsers.map(function(u){ return (<tr key={u.id} className="tr"><td><div className="row g8"><Avatar name={u.name} color={u.color} size={28}/><span style={{fontWeight:600,fontSize:12}}>{u.name}</span></div></td><td style={{fontSize:11,color:"var(--t3)"}}>{u.email}</td><td><span style={{background:u.role==="superadmin"?"var(--am-t)":"var(--in-t)",color:u.role==="superadmin"?"var(--am-d)":"var(--in-d)",borderRadius:5,padding:"2px 7px",fontSize:10,fontWeight:700}}>{u.role==="superadmin"?"👑 Admin":"Usuario"}</span></td><td style={{fontSize:10,color:"var(--t3)"}}>{new Date(u.created_at).toLocaleDateString("es-AR")}</td>
                              <td>
                                {u.id!==me.id&&u.role!=="superadmin"&&(
                                  delUserConf===u.id
                                  ?<div className="row g8">
                                    <button className="btn btn-xs b-cr" onClick={function(){doDeleteUser(u.id,u.name);}}>Confirmar</button>
                                    <button className="btn btn-xs b-ghost" onClick={function(){setDelUserConf(null);}}>No</button>
                                  </div>
                                  :<button className="btn btn-xs b-cr" onClick={function(){setDelUserConf(u.id);}}><Ic n="trash" s={11}/>Eliminar</button>
                                )}
                              </td>
                            </tr>); })}</tbody></table></div>
                  )}
                </div>

                {/* ── IMPORTAR PRECIOS POR EXCEL ── */}
                <div className="card" style={{marginBottom:12}}>
                  <div className="card-h">
                    <div className="card-title"><div className="card-ico" style={{background:"var(--am-l)",color:"var(--am-d)"}}><Ic n="upload" s={14}/></div>📊 Actualizar precios por Excel</div>
                    {pxDone&&<button className="btn btn-xs b-ghost" onClick={function(){setPxFile(null);setPxRows([]);setPxDone(false);}}>Nuevo</button>}
                  </div>
                  <div style={{padding:"14px 16px"}}>
                    <div style={{fontSize:12,color:"var(--t3)",marginBottom:10}}>Subí un Excel con columnas <b>SKU</b> (o <b>Nombre</b>) y <b>Precio</b>. Solo se actualizan los precios, sin tocar el stock.</div>
                    {!pxFile&&(
                      <label style={{display:"block",background:"var(--am-l)",border:"1.5px dashed var(--am-m,#f0b429)",borderRadius:12,padding:"16px",textAlign:"center",cursor:"pointer"}}>
                        <input type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={function(e){var f=e.target.files&&e.target.files[0];if(f)pxHandleFile(f);e.target.value="";}}/>
                        <div style={{fontSize:24,marginBottom:4}}>📂</div>
                        <div style={{fontSize:13,fontWeight:700,color:"var(--am-d)"}}>Elegir Excel o CSV</div>
                        <div style={{fontSize:11,color:"var(--t3)",marginTop:4}}>Columnas requeridas: SKU o Nombre + Precio</div>
                      </label>
                    )}
                    {pxFile&&<div style={{fontSize:12,color:"var(--t2)",fontWeight:700,marginBottom:8}}>📎 {pxFile}</div>}
                    {pxRows.length>0&&(
                      <div style={{marginTop:8}}>
                        <div style={{display:"flex",gap:8,marginBottom:8,flexWrap:"wrap"}}>
                          <span style={{background:"var(--em-l)",color:"var(--em-d)",borderRadius:8,padding:"3px 10px",fontSize:11,fontWeight:700}}>✅ {pxRows.filter(function(r){return r.ok;}).length} válidos</span>
                          {pxRows.filter(function(r){return !r.ok;}).length>0&&<span style={{background:"var(--cr-l,#ffe0e5)",color:"var(--cr)",borderRadius:8,padding:"3px 10px",fontSize:11,fontWeight:700}}>❌ {pxRows.filter(function(r){return !r.ok;}).length} sin coincidencia</span>}
                        </div>
                        <div style={{maxHeight:220,overflowY:"auto",border:"1px solid var(--brd)",borderRadius:10,marginBottom:12}}>
                          <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
                            <thead><tr style={{background:"var(--bg2)"}}><th style={{padding:"6px 10px",textAlign:"left",fontWeight:700}}>Producto</th><th style={{padding:"6px 10px",textAlign:"right",fontWeight:700}}>Precio actual</th><th style={{padding:"6px 10px",textAlign:"right",fontWeight:700}}>Precio nuevo</th><th style={{padding:"6px 10px",textAlign:"center",fontWeight:700}}></th></tr></thead>
                            <tbody>
                              {pxRows.map(function(r,i){
                                return (
                                  <tr key={i} style={{borderTop:"1px solid var(--brd)",opacity:r.ok?1:0.5}}>
                                    <td style={{padding:"6px 10px",fontWeight:600,maxWidth:160,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.name||r.sku||"-"}</td>
                                    <td style={{padding:"6px 10px",textAlign:"right",color:"var(--t3)"}}>{r.oldPrice!=null?"$"+r.oldPrice:"-"}</td>
                                    <td style={{padding:"6px 10px",textAlign:"right",fontWeight:700,color:r.ok?(r.newPrice>r.oldPrice?"var(--cr)":"var(--em-d)"):"var(--t4)"}}>{r.newPrice!=null?"$"+r.newPrice:"-"}</td>
                                    <td style={{padding:"6px 10px",textAlign:"center",fontSize:10}}>{r.ok?"✅":<span style={{color:"var(--cr)",fontWeight:700}}>{r.err}</span>}</td>
                                  </tr>
                                );
                              })}
                            </tbody>
                          </table>
                        </div>
                        {!pxDone&&<button className="cta cta-in" onClick={doPxImport} disabled={pxSaving||pxRows.filter(function(r){return r.ok;}).length===0}>{pxSaving?"Actualizando...":"💾 Actualizar "+pxRows.filter(function(r){return r.ok;}).length+" precios"}</button>}
                        {pxDone&&<div style={{textAlign:"center",padding:12,fontSize:14,fontWeight:700,color:"var(--em-d)"}}>✅ Precios actualizados correctamente</div>}
                      </div>
                    )}
                  </div>
                </div>

                {/* ── ZONA PELIGROSA: RESET ── */}
                <div className="card" style={{marginBottom:12,border:"1.5px solid var(--cr)"}}>
                  <div className="card-h"><div className="card-title"><div className="card-ico" style={{background:"var(--cr-l,#ffe0e5)",color:"var(--cr)"}}><Ic n="trash" s={14}/></div>⚠️ Zona peligrosa</div></div>
                  <div style={{padding:"14px 16px"}}>
                    <div style={{fontSize:12,color:"var(--t3)",marginBottom:12}}>Borra <b>todo el stock, movimientos y consignaciones</b> para empezar desde cero. <b>No se eliminan productos ni usuarios.</b></div>
                    {!resetConf&&(
                      <button className="btn b-cr" style={{width:"100%",padding:"12px",fontWeight:700,borderRadius:12}} onClick={function(){setResetConf(true);}}>🗑️ Resetear todos los datos</button>
                    )}
                    {resetConf&&(
                      <div style={{background:"var(--cr-l,#ffe0e5)",borderRadius:12,padding:"14px",textAlign:"center"}}>
                        <div style={{fontSize:13,fontWeight:700,color:"var(--cr)",marginBottom:12}}>¿Estás seguro? Esta acción no se puede deshacer.</div>
                        <div style={{display:"flex",gap:10,justifyContent:"center"}}>
                          <button className="btn b-cr" style={{padding:"10px 24px",fontWeight:700}} onClick={doResetData} disabled={resetBusy}>{resetBusy?"Borrando...":"Sí, resetear"}</button>
                          <button className="btn b-ghost" style={{padding:"10px 24px"}} onClick={function(){setResetConf(false);}} disabled={resetBusy}>Cancelar</button>
                        </div>
                      </div>
                    )}
                  </div>
                </div>

              </div>
            </div>
          )}

        </div>

        {/* TAB BAR */}
        <nav className="tabbar">
          {(function(){
            var izq = [
              {id:"inicio", lbl:"Inicio", ico:"home"},
              {id:"ventas", lbl:"Ventas", ico:"chart"},
            ];
            var der = [
              {id:"pedidos", lbl:"Pedidos", ico:"list", dot: pedPendCount>0},
              {id:"__mas",   lbl:"Más",     ico:"dots"},
            ];
            function TabItem(t){
              return (
                <div key={t.id} className={"tab"+((tab===t.id||(t.id==="__mas"&&masMenu))?" on":"")}
                  onClick={function(){ if(t.id==="__mas"){ setMasMenu(true); } else { setMasMenu(false); setTab(t.id); } }}>
                  <div className="tab-bub" style={{position:"relative"}}>
                    <Ic n={t.ico} s={22}/>
                    {t.dot&&<div style={{position:"absolute",top:-2,right:-2,width:8,height:8,borderRadius:"50%",background:"var(--cr)",border:"2px solid var(--card)"}}/>}
                  </div>
                  <span className="tab-lbl">{t.lbl}</span>
                </div>
              );
            }
            return (
              <React.Fragment>
                {izq.map(TabItem)}
                {/* Botón central [+] */}
                <div className="tab-fab-wrap" onClick={function(){ setMasMenu(false); setTab("cargar"); }}>
                  <div className="tab-fab"><Ic n="plus" s={26}/></div>
                </div>
                {der.map(TabItem)}
              </React.Fragment>
            );
          })()}
        </nav>

        {/* ── MENÚ "MÁS" ── */}
        {masMenu&&(
          <div className="ovl" style={{alignItems:"flex-end",zIndex:120}} onClick={function(e){if(e.target===e.currentTarget)setMasMenu(false);}}>
            <div style={{background:"var(--card)",borderRadius:"28px 28px 0 0",width:"100%",maxWidth:600,padding:"10px 16px calc(env(safe-area-inset-bottom,0px) + 20px)",boxShadow:"var(--sh3)",animation:"fadeUp .3s cubic-bezier(.16,1,.3,1)"}}>
              <div style={{width:40,height:4,borderRadius:2,background:"var(--brd2)",margin:"6px auto 14px"}}/>
              <div style={{fontSize:16,fontWeight:900,color:"var(--t1)",marginBottom:14,paddingLeft:4}}>Más opciones</div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10}}>
                {(function(){
                  var items = [
                    {id:"stock",     lbl:"Stock",     ico:"box",   col:"var(--in-d)"},
                    {id:"consigna",  lbl:"Consigna",  ico:"send",  col:"var(--in)"},
                    {id:"cuenta",    lbl:"Mi Cuenta", ico:"user",  col:"var(--pri)"},
                  ];
                  if (me.role==="reseller"||me.role==="lider"||me.role==="empresaria") {
                    var pePend = pedEspList.filter(function(p){
                      if (me.role==="reseller")   return p.vendedor_id===me.id && p.estado==="listo_entregar";
                      if (me.role==="lider")      return p.lider_id===me.id && p.estado==="pendiente_lider";
                      if (me.role==="empresaria") return p.empresa_id===me.id && ["pendiente_empresaria","aprobado","enviado_proveedor","recibido"].includes(p.estado);
                      return false;
                    }).length;
                    items.push({id:"pedesp", lbl:"Pedido Especial", ico:"send", col:"var(--am-d)", badge:pePend});
                  }
                  if (me.role==="empresaria" || me.role==="lider") {
                    items.push({id:"resumen", lbl:"Resumen", ico:"chart", col:"var(--em-d,#0a8f4d)"});
                    items.push({id:"equipo", lbl:"Mi Estructura", ico:"users", col:"var(--pu)"});
                  }
                  if (isAdmin) {
                    items.push({id:"catalog", lbl:"Catálogo", ico:"list",   col:"var(--am-d)"});
                    items.push({id:"importar",lbl:"Importar", ico:"upload", col:"var(--t2)"});
                    items.push({id:"admin",   lbl:"Admin",    ico:"shield", col:"var(--in-d)"});
                    items.push({id:"equipo",  lbl:"Empresas", ico:"users",  col:"var(--pu)"});
                  }
                  return items.map(function(it){
                    return (
                      <div key={it.id} onClick={function(){ setTab(it.id); setMasMenu(false); if(["cuenta","equipo","resumen","pedesp"].includes(it.id)) resyncPerfil(); }}
                        style={{background:"var(--bg2)",borderRadius:18,padding:"16px 8px",display:"flex",flexDirection:"column",alignItems:"center",gap:8,cursor:"pointer"}}>
                        <div style={{width:46,height:46,borderRadius:14,background:"var(--card)",display:"flex",alignItems:"center",justifyContent:"center",color:it.col,boxShadow:"var(--sh)",position:"relative"}}>
                          <Ic n={it.ico} s={22}/>
                          {!!it.badge&&<span style={{position:"absolute",top:-4,right:-4,background:"var(--cr,#e0224e)",color:"#fff",borderRadius:9,minWidth:18,height:18,fontSize:10,fontWeight:800,display:"flex",alignItems:"center",justifyContent:"center",padding:"0 4px"}}>{it.badge}</span>}
                        </div>
                        <span style={{fontSize:12,fontWeight:700,color:"var(--t1)"}}>{it.lbl}</span>
                      </div>
                    );
                  });
                })()}
              </div>
            </div>
          </div>
        )}
      </div>

      {/* RETURN CONSIGNA MODAL */}
      {retModal&&(
        <div className="ovl" onClick={function(e){if(e.target===e.currentTarget)setRetModal(null);}}>
          <div className="mbox">
            <div className="mhd">
              <div className="mhd-t">↩️ Registrar devolución</div>
              <button className="ic-btn" onClick={function(){setRetModal(null);}}><Ic n="x" s={16}/></button>
            </div>
            <div className="mbd">
              {(function(){
                var prod = retModal.product||products.find(function(p){return p.id===retModal.product_id;});
                var toUser = retModal.to_user;
                return (
                  <div>
                    <div className="row g12" style={{padding:"12px 14px",background:"var(--bg)",borderRadius:12,border:"1px solid var(--brd)",marginBottom:16}}>
                      <ProdThumb prod={prod} size={44}/>
                      <div style={{flex:1}}>
                        <div style={{fontWeight:700,fontSize:14}}>{prod?prod.name:"Producto"}</div>
                        <div style={{fontSize:11,color:"var(--t3)",marginTop:2}}>{retModal.to_user_id===me.id?"De:":"Para:"} <strong>{retModal.to_user_id===me.id?(retModal.from_user?retModal.from_user.name:"?"):(retModal.to_user?retModal.to_user.name:"?")}</strong> · {retModal.qty} u. pendientes</div>
                      </div>
                    </div>
                    <div className="fld">
                      {(function(){
                        var myI2 = retModal.to_user_id===me.id ? inventory.find(function(i){return i.product_id===retModal.product_id;}) : null;
                        var maxRet = myI2 ? myI2.qty_available : retModal.qty;
                        return (<div>
                          <label className="fl">¿Cuántas unidades? (máx. {maxRet} disponibles)</label>
                          <QtyControl val={Math.min(retQty,maxRet)} set={setRetQty} min={1} max={maxRet} big={true}/>
                        </div>);
                      })()}
                    </div>
                    <div style={{background:"var(--em-l)",border:"1px solid var(--em-t)",borderRadius:10,padding:"10px 14px",fontSize:12,color:"var(--em-d)",marginTop:8}}>
                      ✅ Se restituirán <strong>{retQty} u.</strong> a tu stock y se notificará al remitente.
                    </div>
                  </div>
                );
              })()}
            </div>
            <div className="mft">
              <button className="btn b-ghost" onClick={function(){setRetModal(null);}}>Cancelar</button>
              <button className="btn b-em" onClick={doReturnConsigna}><Ic n="undo" s={14}/>Confirmar devolución</button>
            </div>
          </div>
        </div>
      )}

      {/* MOVIMIENTO DE STOCK MODAL */}
      {movModal&&(
        <div className="ovl" onClick={function(e){if(e.target===e.currentTarget)setMovModal(null);}}>
          <div className="mbox">
            <div className="mhd">
              <div className="mhd-t">📦 Movimiento de stock</div>
              <button className="ic-btn" onClick={function(){setMovModal(null);}}><Ic n="x" s={16}/></button>
            </div>
            <div className="mbd">
              {(function(){
                var prod=movModal.products||products.find(function(p){return p.id===movModal.product_id;});
                return (
                  <div>
                    {/* Product info */}
                    <div className="row g12" style={{padding:"12px 14px",background:"var(--bg)",borderRadius:12,border:"1px solid var(--brd)",marginBottom:16}}>
                      <ProdThumb prod={prod} size={44}/>
                      <div style={{flex:1}}>
                        <div style={{fontWeight:700,fontSize:14}}>{prod?prod.name:""}</div>
                        <div style={{fontSize:11,color:"var(--t3)",marginTop:2}}>{prod?prod.sku:""} · Stock actual: <strong style={{color:"var(--in-d)"}}>{movModal.qty_available} u.</strong></div>
                      </div>
                    </div>

                    {/* Type selector */}
                    <div className="fld">
                      <label className="fl">Tipo de movimiento</label>
                      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8}}>
                        {[
                          {v:"entrada", lbl:"📥 Entrada",  bg:"var(--em-l)",  brd:"var(--em)",   col:"var(--em-d)"},
                          {v:"salida",  lbl:"📤 Salida",   bg:"var(--cr-l)",  brd:"var(--cr)",   col:"var(--cr)"},
                          {v:"ajuste",  lbl:"🔧 Ajuste",   bg:"var(--am-l)",  brd:"var(--am)",   col:"var(--am-d)"},
                        ].map(function(opt){
                          var sel=movType===opt.v;
                          return (
                            <div key={opt.v} onClick={function(){setMovType(opt.v);setMovQty(1);}} style={{padding:"10px 8px",borderRadius:12,border:"2px solid "+(sel?opt.brd:"var(--brd)"),background:sel?opt.bg:"var(--card)",textAlign:"center",cursor:"pointer",fontWeight:700,fontSize:12,color:sel?opt.col:"var(--t3)",transition:"all .15s"}}>
                              {opt.lbl}
                            </div>
                          );
                        })}
                      </div>
                    </div>

                    {/* Quantity */}
                    <div className="fld">
                      <label className="fl">
                        {movType==="ajuste"?"Nueva cantidad total":"Cantidad a "+(movType==="entrada"?"agregar":"retirar")}
                      </label>
                      <QtyControl val={movQty} set={setMovQty} min={movType==="ajuste"?0:1} max={movType==="salida"?movModal.qty_available:9999} big={true}/>
                    </div>

                    {/* Preview */}
                    <div style={{background:movType==="entrada"?"var(--em-l)":movType==="salida"?"var(--cr-l)":"var(--am-l)",border:"1.5px solid "+(movType==="entrada"?"var(--em-t)":movType==="salida"?"var(--cr-t)":"var(--am-t)"),borderRadius:12,padding:"12px 14px",marginBottom:12}}>
                      <div className="row jb" style={{fontSize:13}}>
                        <span style={{color:"var(--t2)"}}>Antes:</span>
                        <span style={{fontFamily:"var(--mf)",fontWeight:700}}>{movModal.qty_available} u.</span>
                      </div>
                      <div className="row jb" style={{fontSize:13,marginTop:4}}>
                        <span style={{color:"var(--t2)"}}>Después:</span>
                        <span style={{fontFamily:"var(--mf)",fontWeight:900,fontSize:16,color:movType==="entrada"?"var(--em-d)":movType==="salida"?"var(--cr)":"var(--am-d)"}}>
                          {movType==="entrada"?movModal.qty_available+movQty:movType==="salida"?Math.max(0,movModal.qty_available-movQty):movQty} u.
                        </span>
                      </div>
                      {movType!=="ajuste"&&(
                        <div className="row jb" style={{fontSize:12,marginTop:4,paddingTop:4,borderTop:"1px solid rgba(0,0,0,.06)"}}>
                          <span style={{color:"var(--t3)"}}>Diferencia:</span>
                          <span style={{fontFamily:"var(--mf)",fontWeight:700,color:movType==="entrada"?"var(--em-d)":"var(--cr)"}}>
                            {movType==="entrada"?"+":"-"}{movQty} u.
                          </span>
                        </div>
                      )}
                    </div>

                    {/* Note */}
                    <div className="fld">
                      <label className="fl">Nota (opcional)</label>
                      <input className="fi" placeholder="Ej: Compra al proveedor, merma, corrección..." value={movNote} onChange={function(e){setMovNote(e.target.value);}}/>
                    </div>

                    {/* History */}
                    {movHistory.filter(function(m){var p=movModal.products||products.find(function(x){return x.id===movModal.product_id;});return p&&m.sku===p.sku;}).slice(0,5).length>0&&(
                      <div style={{marginTop:8}}>
                        <div style={{fontSize:11,fontWeight:700,color:"var(--t3)",textTransform:"uppercase",letterSpacing:".08em",marginBottom:8}}>Últimos movimientos</div>
                        {movHistory.filter(function(m){var p=movModal.products||products.find(function(x){return x.id===movModal.product_id;});return p&&m.sku===p.sku;}).slice(0,5).map(function(m,i){
                          var ico=m.type==="entrada"?"📥":m.type==="salida"?"📤":"🔧";
                          var col=m.type==="entrada"?"var(--em-d)":m.type==="salida"?"var(--cr)":"var(--am-d)";
                          return (
                            <div key={i} className="row" style={{gap:10,padding:"7px 10px",borderRadius:8,background:"var(--bg)",marginBottom:5,fontSize:11}}>
                              <span style={{flexShrink:0}}>{ico}</span>
                              <span style={{color:"var(--t3)",flexShrink:0}}>{m.t}</span>
                              <span style={{fontFamily:"var(--mf)",fontWeight:700,color:col,flexShrink:0}}>{m.diff>=0?"+":""}{m.diff}</span>
                              <span style={{color:"var(--t2)",flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{m.note||m.type}</span>
                              <span style={{fontFamily:"var(--mf)",color:"var(--t3)",flexShrink:0}}>{m.before}→{m.after}</span>
                            </div>
                          );
                        })}
                      </div>
                    )}
                  </div>
                );
              })()}
            </div>
            <div className="mft">
              <button className="btn b-ghost" onClick={function(){setMovModal(null);}}>Cancelar</button>
              <button className="btn b-pri" onClick={doMovimiento}><Ic n="check" s={14}/>Confirmar movimiento</button>
            </div>
          </div>
        </div>
      )}

      {/* EDIT STOCK MODAL */}
      {editStock&&(
        <div className="ovl" onClick={function(e){if(e.target===e.currentTarget)setEditStock(null);}}>
          <div className="mbox">
            <div className="mhd">
              <div className="mhd-t">✏️ Corregir cantidad</div>
              <button className="ic-btn" onClick={function(){setEditStock(null);}}><Ic n="x" s={16}/></button>
            </div>
            <div className="mbd">
              {(function(){
                var prod = editStock.products||products.find(function(p){return p.id===editStock.product_id;});
                return (
                  <div>
                    <div className="row g12" style={{padding:"12px 14px",background:"var(--bg)",borderRadius:12,border:"1px solid var(--brd)",marginBottom:18}}>
                      <ProdThumb prod={prod} size={44}/>
                      <div style={{flex:1}}>
                        <div style={{fontWeight:700,fontSize:14}}>{prod?prod.name:""}</div>
                        <div style={{fontSize:11,color:"var(--t3)",marginTop:2}}>{prod?prod.sku:""} · Stock actual: <strong>{editStock.qty_available} u.</strong></div>
                      </div>
                    </div>
                    <div className="fld">
                      <label className="fl">Nueva cantidad disponible</label>
                      <QtyControl val={editQty} set={setEditQty} min={0} big={true}/>
                    </div>
                    <div style={{background:"var(--am-l)",border:"1px solid var(--am-t)",borderRadius:10,padding:"10px 14px",fontSize:12,color:"var(--am-d)",marginTop:8}}>
                      ⚠️ Esta acción sobreescribe directamente el stock. Usá solo para corregir errores de carga.
                    </div>
                  </div>
                );
              })()}
            </div>
            <div className="mft">
              <button className="btn b-ghost" onClick={function(){setEditStock(null);}}>Cancelar</button>
              <button className="btn b-pri" onClick={doEditStock}><Ic n="check" s={14}/>Guardar corrección</button>
            </div>
          </div>
        </div>
      )}

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
