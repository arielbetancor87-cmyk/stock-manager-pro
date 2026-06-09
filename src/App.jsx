import { useState, useEffect, useCallback, memo, useMemo } from "react";
import ConsignacionModule from "./ConsignacionModule";
import * as XLSX from "xlsx";
import { createClient } from "@supabase/supabase-js";

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
export default function App() {
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
  const [tab,       setTab]       = useState("stock");
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
  const [carouselIdx,  setCarouselIdx]  = useState({});  // { product_id: activeIdx }
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
  var myStock = inventory.filter(function(i){ return i.qty_available>0; });
  var unreadNotifs = notifs.filter(function(n){ return !n.read; }).length;
  var pendingTransfers = transfers.filter(function(t){ return t.to_user_id===me?.id && t.status==="pending"; }).length;
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

  async function doDeleteUser(userId, userName) {
    // Delete from public.users (CASCADE will handle related data)
    var r = await sb.from("users").delete().eq("id", userId);
    if (r.error) { toast("Error",""+r.error.message,"e"); return; }
    setAllUsers(function(p){ return p.filter(function(u){ return u.id!==userId; }); });
    // Also remove from contacts
    await sb.from("contacts").delete().eq("contact_id", userId);
    await sb.from("contacts").delete().eq("user_id", userId);
    toast("Usuario eliminado", userName, "i");
    setDelUserConf(null);
    await loadData(me.id, me.role);
  }

  // ── PERSISTENCIA LOCAL ─────────────────────────────────────────────────────
  // [localStorage eliminado — pedidos y estado en Supabase]

  // ── SUPABASE DATA LOADERS ────────────────────────────────────────────────────
  async function loadCarousels() {
    try {
      const { data } = await sb.from("offer_carousels")
        .select("*").eq("active", true)
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
      var up = await sb.storage.from("product-images").upload(path, file, { upsert: true });
      if (up.error) throw up.error;
      var pub = sb.storage.from("product-images").getPublicUrl(path);
      setCImg(pub.data.publicUrl);
      toast("✅ Imagen lista","","s");
    } catch(e) { toast("Error al subir", e.message, "e"); }
    finally { setCImgUp(false); }
  }

  async function saveCarousel() {
    if (!cTitle.trim()) return;
    setCSaving(true);
    try {
      const row = { title: cTitle.trim(), subtitle: cSubtitle.trim(), bg_color: cBg, emoji: cEmoji, image_url: cImg||null, link_tab: cLink.trim(), active: true, sort_order: carousels.length };
      if (cEditId) {
        await sb.from("offer_carousels").update(row).eq("id", cEditId);
      } else {
        await sb.from("offer_carousels").insert(row);
      }
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
        var pr = await sb.from("products").select("*").eq("is_active",true).order("name").range(from, from+PAGE-1);
        if (pr.data && pr.data.length > 0) {
          allProds = allProds.concat(pr.data);
          if (pr.data.length < PAGE) break; // last page
          from += PAGE;
        } else { break; }
      }
      setProducts(allProds);

      // My inventory
      var inv = await sb.from("inventory").select("*, products(*)").eq("user_id",userId);
      if (inv.data) setInventory(inv.data);

      // My contacts
      var cts = await sb.from("contacts").select("*, contact:contact_id(id,name,email,color,role)").eq("user_id",userId);
      if (cts.data) setContacts(cts.data.map(function(c){ return c.contact; }).filter(Boolean));

      // Transfers (sent and received)
      var tx = await sb.from("transfers").select("*, product:product_id(id,name,emoji,photo_url,sku,price,category), from_user:from_user_id(id,name,email,color), to_user:to_user_id(id,name,email,color)").or("from_user_id.eq."+userId+",to_user_id.eq."+userId).order("created_at",{ascending:false});
      if (tx.data) setTransfers(tx.data);

      // Notifications
      var nf = await sb.from("notifications").select("*").eq("to_user_id",userId).order("created_at",{ascending:false}).limit(50);
      if (nf.data) setNotifs(nf.data);

      // My logs
      var lg = await sb.from("sale_logs").select("*, product:product_id(name,sku,emoji)").eq("user_id",userId).order("created_at",{ascending:false}).limit(100);
      var ped = await sb.from("pedidos").select("*, product:product_id(id,name,sku,price,emoji)").eq("user_id",userId).order("created_at",{ascending:false}).limit(200);
      if (!ped.error) setPedidos(ped.data||[]);
      if (lg.data) setLogs(lg.data);

      // Admin only
      if (userRole === "superadmin") {
        var us = await sb.from("users").select("*").order("created_at",{ascending:false});
        if (us.data) setAllUsers(us.data);
        var stats = await sb.from("admin_dashboard").select("*").single();
        if (stats.data) setAdminStats(stats.data);
      }
      // Consignaciones activas (enviadas + recibidas)
      try {
        const [envC, recC] = await Promise.all([
          sb.from("consignaciones").select("id,status,owner_id,vendedora_id,comision_pct,created_at")
            .eq("owner_id",userId).neq("status","cancelada"),
          sb.from("consignaciones").select("id,status,owner_id,vendedora_id,comision_pct,created_at")
            .eq("vendedora_id",userId).neq("status","cancelada"),
        ]);
        setConsignActivas([...(envC.data||[]),...(recC.data||[])]);
      } catch(e) { /* tablas pueden no existir aún */ }
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
        await loadData(pr.data.id, pr.data.role);
        loadCarousels();
        return true;
      }
      // No existe la fila todavía. Reintentar por si un trigger la está creando.
      if (intentos < 3) {
        await new Promise(function(r){ setTimeout(r, 400); });
        return ensureProfile(session, intentos + 1);
      }
      // Sigue sin existir → crearla nosotros como revendedora (entra directo)
      var meta2 = session.user.user_metadata || {};
      var nuevoNombre = meta2.full_name || meta2.name || (session.user.email||"").split("@")[0];
      var nuevoColor  = ["#e0224e","#8b5cf6","#0ea5e9","#10b981","#f59e0b","#ec4899"][Math.floor(Math.random()*6)];
      var ins = await sb.from("users").insert({
        id:         session.user.id,
        email:      session.user.email,
        name:       nuevoNombre,
        role:       "reseller",
        color:      nuevoColor,
        avatar_url: meta2.avatar_url || meta2.picture || null
      }).select("*").maybeSingle();
      if (ins.data) {
        setMe(ins.data);
        await loadData(ins.data.id, ins.data.role);
        loadCarousels();
        return true;
      }
      return false;
    }

    sb.auth.getSession().then(async function(res){
      var session = res.data.session;
      if (session) { await ensureProfile(session); }
      setLoading(false);
    });

    var sub = sb.auth.onAuthStateChange(async function(event, session){
      if (event === "SIGNED_OUT") { setMe(null); setProducts([]); setInventory([]); setContacts([]); setLogs([]); }
      // Al volver de Google (o cualquier login nuevo), asegurar perfil
      if (event === "SIGNED_IN" && session) { await ensureProfile(session); }
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
    if (!pedNombre.trim()) return;
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
    var { data, error } = await sb.from("pedidos").insert(row).select("*, product:product_id(id,name,sku,price,emoji)").single();
    if (error) { toast("Error al guardar pedido", error.message, "e"); return; }
    setPedidos(function(prev){ return [data, ...prev]; });
    setPedNombre(""); setPedWA(""); setPedProdId(""); setPedQty(1); setPedNota(""); setPedPSrch(""); setPedSena("");
    setShowPedForm(false);
    toast("Pedido guardado!", pedNombre.trim()+(prod?" — "+prod.name:""), "s");
  }

  async function doMarcarPagado(id) {
    var { error } = await sb.from("pedidos").update({ pagado: true }).eq("id", id);
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
    var { error } = await sb.from("pedidos")
      .update({ estado: "entregado", entregado_at: now })
      .eq("id", id);
    if (error) { toast("Error", error.message, "e"); return; }

    // Actualizar inventario si tiene producto asociado
    if (ped.product_id) {
      var invRow = inventory.find(function(i){ return i.product_id===ped.product_id; });
      if (invRow && invRow.qty_available >= ped.qty) {
        await sb.from("inventory")
          .update({ qty_available: invRow.qty_available - ped.qty, qty_sold: (invRow.qty_sold||0) + ped.qty })
          .eq("id", invRow.id);
        // Registrar venta en sale_logs
        await sb.from("sale_logs").insert({
          user_id: me.id, product_id: ped.product_id,
          qty: ped.qty, sale_price: ped.product_price||0, source: "pedido"
        });
        // Ledger
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
      const { error:upErr } = await sb.storage.from("product-images").upload(path, file, {upsert:true});
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
    await sb.from("sale_logs").insert({user_id:me.id,product_id:item.product_id,qty:Math.abs(newQty-item.qty_available),sale_price:0,source:movType==="salida"?"own_stock":movType});
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
        <img src="data:image/png;base64,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" alt="Venta Directa" style={{width:220,marginBottom:4,filter:"drop-shadow(0 8px 24px rgba(0,0,0,.45)) drop-shadow(0 2px 8px rgba(0,0,0,.35))"}} /><div style={{display:"none"}}>VD</div>
        <div className="loading-spinner"/>
        <div style={{fontSize:12,color:"rgba(255,255,255,.7)",fontWeight:600,letterSpacing:".05em"}}>Conectando con Supabase...</div>
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
            <img src="data:image/png;base64,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" alt="Venta Directa" style={{width:200,marginBottom:4}} /><div style={{display:"none"}}>VD</div>
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
          {tab==="stock"&&(
          <div className="hdr-search">
            <span className="hdr-search-ico"><Ic n="search" s={18}/></span>
            <input placeholder="Buscar productos, marcas o categorías..." value={srchStock} onChange={function(e){setSrchStock(e.target.value);}}/>
          </div>
          )}
        </div>

        <div className="main">

          {/* ══ STOCK ══ */}
          {tab==="stock"&&(function(){
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
            var slides = carousels.length>0 ? carousels : [
              {id:"d1",title:"¡Bienvenida!",subtitle:"Tu catálogo digital",bg_color:"#e0224e",emoji:"🛍️",link_tab:"catalog"},
              {id:"d2",title:"Compartí tus productos",subtitle:"Enviá lista por WhatsApp",bg_color:"#10b981",emoji:"📲",link_tab:""},
            ];

            return (
              <div style={{paddingBottom:24}}>

                {/* ─ CARRUSEL DE OFERTAS ─ */}
                {slides.length>0&&(
                  <div style={{marginBottom:18,paddingTop:10}}>
                    <div className="carousel-wrap" id="main-carousel"
                      onScroll={function(e){
                        var w=e.target.scrollWidth/slides.length;
                        setCarouselIdx(Math.round(e.target.scrollLeft/w));
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
                    {[
                      {ico:"📦",bg:"var(--in-l)",col:"var(--in)",lbl:"Mi Stock",sub:ownStock.reduce(function(s,i){return s+i.qty_available;},0)+" productos",tab:"stock",scroll:true},
                      {ico:"📤",bg:"var(--bl-l)",col:"var(--bl)",lbl:"Enviados",sub:consignaEnv>0?consignaEnv+" activos":"Ver enviados",tab:"enviados",badge:consignaEnv||null},
                      {ico:"📋",bg:"#fff3e6",col:"#e06a00",lbl:"Pedidos",sub:pedPendCount>0?pedPendCount+" pendientes":"Ver pedidos",tab:"pedidos",badge:pedPendCount||null},
                      {ico:"💰",bg:"var(--em-l)",col:"var(--em-d)",lbl:"Ventas",sub:fmtARS(totalVal).replace("$ ","$"),tab:"ventas"},
                    ].map(function(f,i){
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
                  <button className="btn b-pri" onClick={function(){setShowPedForm(function(v){return !v;});}}>
                    <Ic n="plus" s={15}/>{showPedForm?"Cancelar":"Nuevo"}
                  </button>
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
                            {products.filter(function(p){return p.name.toLowerCase().includes(pedPSrch.toLowerCase());}).slice(0,8).map(function(p){
                              return (
                                <div key={p.id} onClick={function(){setPedProdId(p.id);setPedPSrch(p.name);}} style={{padding:"10px 12px",cursor:"pointer",borderBottom:"1px solid var(--brd)",display:"flex",alignItems:"center",gap:8,background:pedProdId===p.id?"var(--pri-l)":"var(--card)"}}>
                                  <ProdThumb prod={p} size={32}/>
                                  <div><div style={{fontSize:13,fontWeight:700}}>{p.name}</div><div style={{fontSize:11,color:"var(--t3)"}}>{fmtARS(p.price)}</div></div>
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
                      <button className="cta cta-am" onClick={doAddPedido} disabled={!pedNombre.trim()}>
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
              </div>
            </div>
          )}

        </div>

        {/* TAB BAR */}
        <nav className="tabbar">
          {TABS.map(function(t){
            var hasDot = (t.id==="contacts"&&totalBadge>0)||(t.id==="pedidos"&&pedPendCount>0);
            return (
              <div key={t.id} className={"tab"+(tab===t.id?" on":"")} onClick={function(){setTab(t.id);}}>
                <div className="tab-bub" style={{position:"relative"}}>
                  <Ic n={t.ico} s={22}/>
                  {hasDot&&<div style={{position:"absolute",top:-2,right:-2,width:8,height:8,borderRadius:"50%",background:"var(--cr)",border:"2px solid var(--card)"}}/>}
                </div>
                <span className="tab-lbl">{t.lbl}</span>
              </div>
            );
          })}
        </nav>
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
