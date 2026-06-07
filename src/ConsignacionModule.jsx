/**
 * ConsignacionModule.jsx — v4
 * Rediseño completo para uso cotidiano:
 *  - Wizard 3 pasos para crear entregas
 *  - Barra de progreso por producto
 *  - Botón grande "¡Vendí 1!" con animación
 *  - Liquidación con efectivo o transferencia + comprobante
 *  - Pago parcial o total
 *  - Deudas agrupadas por persona
 */
import { useState, useEffect, useRef, memo, useMemo, useCallback } from "react";

const ProdThumb = memo(({ prod, size = 40 }) => {
  if (!prod) return <div style={{ width:size,height:size,borderRadius:10,background:"#f0f0f0",flexShrink:0 }}/>;
  if (prod.photo_url) return <img src={prod.photo_url} alt="" style={{ width:size,height:size,borderRadius:10,objectFit:"cover",flexShrink:0,border:"1px solid #eee" }}/>;
  return <div style={{ width:size,height:size,borderRadius:10,background:"#f0f0f0",display:"flex",alignItems:"center",justifyContent:"center",fontSize:Math.round(size*.48),flexShrink:0 }}>{prod.emoji||"📦"}</div>;
});

const Spin = () => <div style={{ display:"inline-block",width:16,height:16,border:"2.5px solid rgba(255,255,255,.3)",borderTopColor:"#fff",borderRadius:"50%",animation:"spin .7s linear infinite",flexShrink:0 }}/>;

function Avatar({ name, color, size=38 }) {
  const ini = (name||"?").trim().split(" ").slice(0,2).map(w=>w[0]).join("").toUpperCase();
  return <div style={{ width:size,height:size,borderRadius:"50%",background:color||"#e0224e",display:"flex",alignItems:"center",justifyContent:"center",fontWeight:900,fontSize:Math.round(size*.38),color:"#fff",flexShrink:0 }}>{ini}</div>;
}

function fmt(n) { return "$ " + Number(n||0).toLocaleString("es-AR",{minimumFractionDigits:2}); }

async function logMov(sb, me, opts) {
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
  } catch(e) { console.warn("logMov:", e.message); }
}

function ProgBar({ vendidas, devueltas, total }) {
  const pV = total>0?(vendidas/total)*100:0;
  const pD = total>0?(devueltas/total)*100:0;
  return (
    <div style={{ height:7,borderRadius:4,background:"#eee",overflow:"hidden",display:"flex" }}>
      <div style={{ width:`${pV}%`,background:"#10b981",transition:"width .3s" }}/>
      <div style={{ width:`${pD}%`,background:"#bbb",transition:"width .3s" }}/>
    </div>
  );
}

export default function ConsignacionModule({ sb, me, products, inventory, contacts, onRefresh, toast, fmtARS, vistaInicial }) {
  const COMISION_DEFAULT = 30;

  // vistaInicial: "enviados" muestra solo entregas, "recibidos" muestra solo recibidas
  const vistaDefault = vistaInicial === "enviados" ? "main_env" : vistaInicial === "recibidos" ? "main_rec" : "main";
  const [view,    setView]    = useState(vistaDefault);
  const [loading, setLoading] = useState(true);
  const [saving,  setSaving]  = useState(false);

  // Carga manual de stock propio
  const [cargaPid,   setCargaPid]   = useState("");
  const [cargaQty,   setCargaQty]   = useState(1);
  const [cargaShow,  setCargaShow]  = useState(false);
  const [cargaSaving,setCargaSaving]= useState(false);

  // Historial de transfers
  const [transfers,    setTransfers]    = useState([]);
  const [txLoading,    setTxLoading]    = useState(false);
  const [txExpanded,   setTxExpanded]   = useState(false);

  const [enviadas,  setEnviadas]  = useState([]);
  const [recibidas, setRecibidas] = useState([]);
  const [deudas,    setDeudas]    = useState([]);

  // wizard
  const [paso,        setPaso]        = useState(1);
  const [vendedoraId, setVendedoraId] = useState("");
  const [comision,    setComision]    = useState(COMISION_DEFAULT);
  const [notas,       setNotas]       = useState("");
  const [carrito,     setCarrito]     = useState({});
  const [prodSrch,    setProdSrch]    = useState("");

  // detalle
  const [detEnv,      setDetEnv]      = useState(null);
  const [detEnvItems, setDetEnvItems] = useState([]);
  const [detRec,      setDetRec]      = useState(null);
  const [detRecItems, setDetRecItems] = useState([]);
  const [sellAnim,    setSellAnim]    = useState(null);

  // liquidar
  const [liqTarget, setLiqTarget] = useState(null);
  const [liqModo,   setLiqModo]   = useState("efectivo");
  const [liqMonto,  setLiqMonto]  = useState("");
  const [liqFoto,   setLiqFoto]   = useState(null);
  const [liqNota,   setLiqNota]   = useState("");
  const fileRef = useRef();

  const [dbError, setDbError] = useState(null);
  useEffect(() => { if (me) loadAll(); }, [me]);

  async function loadAll() {
    setLoading(true);
    setDbError(null);
    try {
      const envRes = await sb.from("consignaciones")
        .select("*, vendedora:vendedora_id(id,name,color), items:consignacion_items(*, product:product_id(id,name,sku,price,emoji,photo_url))")
        .eq("owner_id",me.id).neq("status","cancelada").order("created_at",{ascending:false});
      if (envRes.error) {
        if (envRes.error.code==="42P01"||envRes.error.message?.includes("does not exist")) {
          setDbError("falta_tabla");
          setLoading(false);
          return;
        }
        throw envRes.error;
      }
      const recRes = await sb.from("consignaciones")
        .select("*, owner:owner_id(id,name,color), items:consignacion_items(*, product:product_id(id,name,sku,price,emoji,photo_url))")
        .eq("vendedora_id",me.id).neq("status","cancelada").order("created_at",{ascending:false});
      const deuRes = await sb.from("consignacion_deudas")
        .select("*, vendedora:vendedora_id(id,name,color), item:item_id(id,product_id,consignacion_id,precio_venta), product:item_id(product_id(id,name,sku,emoji,price))")
        .eq("owner_id",me.id).order("created_at",{ascending:false});
      setEnviadas(envRes.data||[]);
      setRecibidas(recRes.data||[]);
      setDeudas(deuRes.data||[]);
      // Cargar historial de transfers
      try {
        const txRes = await sb.from("transfers")
          .select("*, product:product_id(id,name,sku,emoji,photo_url), from_user:from_user_id(id,name,color), to_user:to_user_id(id,name,color)")
          .or("from_user_id.eq."+me.id+",to_user_id.eq."+me.id)
          .order("created_at",{ascending:false}).limit(50);
        setTransfers(txRes.data||[]);
      } catch(e) { /* ignore */ }
    } catch(e) { toast("Error cargando",""+e.message,"e"); }
    finally { setLoading(false); }
  }

  // ── CARGA MANUAL DE STOCK PROPIO ─────────────────────────────────────────────
  async function cargarStockPropio() {
    if (!cargaPid || cargaQty < 1) { toast("Elegí un producto y cantidad","","e"); return; }
    setCargaSaving(true);
    try {
      const prod = products.find(p=>p.id===cargaPid);
      const { data:inv } = await sb.from("inventory")
        .select("*").eq("user_id",me.id).eq("product_id",cargaPid).maybeSingle();
      if (inv) {
        await sb.from("inventory").update({
          qty_available: inv.qty_available + cargaQty,
          stock_propio:  (inv.stock_propio||0) + cargaQty,
          source: "own"
        }).eq("id", inv.id);
      } else {
        await sb.from("inventory").insert({
          user_id:me.id, product_id:cargaPid,
          qty_available:cargaQty, qty_sold:0,
          stock_propio:cargaQty, stock_recibido:0, source:"own"
        });
      }
      try {
        await sb.from("stock_movements").insert({
          product_id:cargaPid, user_id:me.id, qty:cargaQty,
          estado_anterior:null, estado_nuevo:"stock_central",
          referencia_tipo:"carga",
          nota:"Carga manual — "+(prod?.name||"")
        });
      } catch(e) {}
      toast("✅ Stock cargado","+"+ cargaQty +" u. de "+(prod?.name||""),"s");
      setCargaPid(""); setCargaQty(1); setCargaShow(false);
      onRefresh();
    } catch(e) { toast("Error",e.message,"e"); }
    finally { setCargaSaving(false); }
  }

  async function crearConsignacion() {
    const items = Object.entries(carrito).filter(([,q])=>q>0);
    if (!vendedoraId) { toast("Elegí una persona","","e"); return; }
    if (!items.length) { toast("Agregá productos","","e"); return; }
    setSaving(true);
    try {
      for (const [pid,qty] of items) {
        const inv=inventory.find(i=>i.product_id===pid);
        const p=products.find(x=>x.id===pid);
        if (!inv||inv.qty_available<qty) throw new Error(`Stock insuficiente: "${p?.name}" — tenés ${inv?.qty_available??0} u.`);
      }
      const { data:consig, error:cErr } = await sb.from("consignaciones").insert({
        owner_id:me.id, vendedora_id:vendedoraId, comision_pct:comision,
        notas:notas.trim()||null, status:"activa"
      }).select().single();
      if (cErr) throw cErr;
      for (const [pid,qty] of items) {
        const p=products.find(x=>x.id===pid);
        const inv=inventory.find(i=>i.product_id===pid);
        await sb.from("consignacion_items").insert({ consignacion_id:consig.id, product_id:pid, qty_enviada:qty, qty_vendida:0, qty_devuelta:0, precio_venta:p?.price??0 });
        await sb.from("inventory").update({ qty_available:inv.qty_available-qty }).eq("id",inv.id);
        await logMov(sb, me, {
          product_id: pid, qty: -qty,
          estado_anterior: "stock_central", estado_nuevo: "en_consigna",
          related_user_id: vendedoraId,
          referencia_tipo: "consignacion", referencia_id: consig.id,
          nota: "Entregado en consigna a " + (contacts.find(c=>c.id===vendedoraId)?.name||"")
        });
        const { data:vInv } = await sb.from("inventory").select("*").eq("user_id",vendedoraId).eq("product_id",pid).maybeSingle();
        if (vInv) { await sb.from("inventory").update({ qty_available:vInv.qty_available+qty }).eq("id",vInv.id); }
        else { await sb.from("inventory").insert({ user_id:vendedoraId, product_id:pid, qty_available:qty, qty_sold:0, source:"consigna", supplier_id:me.id }); }
      }
      const totalU=items.reduce((s,[,q])=>s+q,0);
      await sb.from("notifications").insert({ to_user_id:vendedoraId, from_name:me.name, type:"transfer", message:`📦 ${me.name} te entregó ${items.length} producto${items.length!==1?"s":""} en consignación (${totalU} u. en total)` });
      const vName=contacts.find(c=>c.id===vendedoraId)?.name??"la vendedora";
      toast("✅ Entrega registrada",`${totalU} u. a ${vName}`,"s");
      setCarrito({}); setVendedoraId(""); setNotas(""); setPaso(1); setView("main");
      loadAll(); onRefresh();
    } catch(err) { toast("Error",err.message,"e"); }
    finally { setSaving(false); }
  }

  async function vender(item, consig) {
    const disp = item.qty_enviada-item.qty_vendida-item.qty_devuelta;
    if (disp<1) { toast("Sin unidades disponibles","","e"); return; }
    setSellAnim(item.id); setSaving(true);
    try {
      const p=item.product||products.find(x=>x.id===item.product_id);
      const precio=item.precio_venta;
      // Regla de negocio: la vendedora se queda con comision_pct% del precio.
      // La deuda al propietario es el resto. Ej: 30% comisión → deuda = 70% del precio.
      const pct     = (consig.comision_pct!=null ? consig.comision_pct : 30) / 100;
      const comis   = Math.round(precio * pct * 100) / 100;       // ganancia vendedora
      const aPagar  = Math.round((precio - comis) * 100) / 100;   // deuda al propietario
      await sb.from("consignacion_items").update({ qty_vendida:item.qty_vendida+1 }).eq("id",item.id);
      const { data:vInv } = await sb.from("inventory").select("*").eq("user_id",consig.vendedora_id).eq("product_id",item.product_id).maybeSingle();
      if (vInv) await sb.from("inventory").update({ qty_available:Math.max(0,vInv.qty_available-1), qty_sold:(vInv.qty_sold||0)+1, stock_recibido:Math.max(0,(vInv.stock_recibido||0)-1) }).eq("id",vInv.id);
      var ahora = new Date().toISOString();
      await sb.from("consignacion_deudas").insert({
        consignacion_id: consig.id,
        item_id:         item.id,
        owner_id:        consig.owner_id,
        vendedora_id:    consig.vendedora_id,
        qty:             1,
        monto_total:     precio,
        comision:        comis,
        monto_a_pagar:   aPagar,
        pagada:          false,
        fecha_venta:     ahora,
        fecha_envio_consigna: consig.created_at,
        notas_auditoria: "Venta registrada por " + me.name + " — Precio: " + fmt(precio) + " | Deuda propietario: " + fmt(aPagar) + " | Ganancia vendedora: " + fmt(comis)
      });
      await sb.from("sale_logs").insert({ user_id:consig.vendedora_id, product_id:item.product_id, qty:1, sale_price:precio, source:"consignment" });
      await logMov(sb, me, {
        product_id: item.product_id, qty: -1,
        estado_anterior: "en_consigna", estado_nuevo: "vendido",
        related_user_id: consig.owner_id,
        referencia_tipo: "venta_consigna", referencia_id: item.id,
        nota: "Vendido en consigna — comisión: " + fmt(comis)
      });
      await sb.from("notifications").insert({ to_user_id:consig.owner_id, from_name:me.name, type:"sale", message:`💰 ${me.name} vendió 1× "${p?.name}". Rendición: ${fmt(aPagar)} (tu comisión: ${fmt(comis)})` });
      toast(`💰 ¡Vendido!`,`+${fmt(comis)} para vos`,"s");
      await reloadDetRec(consig); loadAll();
    } catch(err) { toast("Error",err.message,"e"); }
    finally { setSaving(false); setTimeout(()=>setSellAnim(null),700); }
  }

  async function devolver(item, consig) {
    const disp=item.qty_enviada-item.qty_vendida-item.qty_devuelta;
    if (disp<1) { toast("Sin unidades para devolver","","e"); return; }
    setSaving(true);
    try {
      const p=item.product||products.find(x=>x.id===item.product_id);
      // Devolución atómica en el servidor (suma al propietario + descuenta a la vendedora)
      const r = await sb.rpc("rpc_devolver_unidad", {
        p_item_id:      item.id,
        p_product_id:   item.product_id,
        p_owner_id:     consig.owner_id,
        p_vendedora_id: consig.vendedora_id
      });
      if (r.error) throw r.error;
      await sb.from("notifications").insert({ to_user_id:consig.owner_id, from_name:me.name, type:"confirm", message:`↩️ ${me.name} devolvió 1× "${p?.name}". El stock volvió a vos.` });
      await logMov(sb, me, {
        product_id: item.product_id, qty: 1,
        estado_anterior: "en_consigna", estado_nuevo: "stock_central",
        related_user_id: consig.owner_id,
        referencia_tipo: "devolucion", referencia_id: item.id,
        nota: "Devolución a propietario — " + (p?.name||"")
      });
      toast("↩️ Devuelto",`"${p?.name}" vuelve al propietario`,"s");
      await reloadDetRec(consig); loadAll();
    } catch(err) { toast("Error",err.message,"e"); }
    finally { setSaving(false); }
  }

  async function reloadDetRec(consig) {
    const { data } = await sb.from("consignacion_items").select("*, product:product_id(id,name,sku,price,emoji,photo_url)").eq("consignacion_id",consig.id);
    setDetRecItems(data||[]); setDetRec({...consig});
  }

  async function abrirDetEnv(consig) {
    const { data } = await sb.from("consignacion_items").select("*, product:product_id(id,name,sku,price,emoji,photo_url)").eq("consignacion_id",consig.id);
    setDetEnv(consig); setDetEnvItems(data||[]); setView("detalle_env");
  }

  async function abrirDetRec(consig) {
    const { data } = await sb.from("consignacion_items").select("*, product:product_id(id,name,sku,price,emoji,photo_url)").eq("consignacion_id",consig.id);
    setDetRec(consig); setDetRecItems(data||[]); setView("detalle_rec");
  }

  function abrirLiquidar(vendedoraObj) {
    const dp=deudas.filter(d=>d.vendedora_id===vendedoraObj.id&&!d.pagada);
    const total=dp.reduce((s,d)=>s+d.monto_a_pagar,0);
    setLiqTarget({ vendedora:vendedoraObj, deudas:dp, total });
    setLiqMonto(total.toFixed(2)); setLiqModo("efectivo"); setLiqFoto(null); setLiqNota("");
    setView("liquidar");
  }

  async function confirmarLiquidacion() {
    if (!liqTarget) return;
    setSaving(true);
    try {
      const monto=parseFloat(liqMonto)||0;
      const esParcial=monto<liqTarget.total-0.01;
      if (esParcial) {
        let resto=monto;
        for (const d of liqTarget.deudas) {
          if (resto<=0) break;
          if (d.monto_a_pagar<=resto+0.01) {
            await sb.from("consignacion_deudas").update({ pagada:true, pagada_at:new Date().toISOString(), notas_pago:liqNota||null, modo_pago:liqModo }).eq("id",d.id);
            resto-=d.monto_a_pagar;
          }
        }
        toast("💵 Pago parcial registrado",`${fmt(monto)} de ${fmt(liqTarget.total)}`,"s");
      } else {
        for (const d of liqTarget.deudas) {
          await sb.from("consignacion_deudas").update({ pagada:true, pagada_at:new Date().toISOString(), notas_pago:liqNota||null, modo_pago:liqModo }).eq("id",d.id);
        }
        toast("✅ Deuda liquidada completamente",`${fmt(liqTarget.total)} cobrados`,"s");
      }
      await sb.from("notifications").insert({ to_user_id:liqTarget.vendedora.id, from_name:me.name, type:"confirm", message:`✅ ${me.name} registró un pago de ${fmt(monto)} (${liqModo}). ${esParcial?"Quedan deudas pendientes.":"¡Deuda saldada!"}` });
      // Ledger: marcar unidades vendidas → pagadas
      for (const d of liqTarget.deudas.filter(x => !x.pagada)) {
        await logMov(sb, me, {
          product_id: d.item?.product_id || null, qty: 1,
          estado_anterior: "vendido", estado_nuevo: "pagado",
          related_user_id: liqTarget.vendedora.id,
          referencia_tipo: "liquidacion", referencia_id: d.id,
          nota: "Liquidación " + liqModo + " — " + fmt(monto)
        });
      }
      setView("main"); loadAll();
    } catch(err) { toast("Error",err.message,"e"); }
    finally { setSaving(false); }
  }

  function handleFoto(e) {
    const file=e.target.files[0]; if(!file) return;
    const r=new FileReader(); r.onload=ev=>setLiqFoto(ev.target.result); r.readAsDataURL(file);
  }

  // derivados — con useMemo para evitar recálculos en cada render
  const deudasPend     = useMemo(()=>deudas.filter(d=>!d.pagada), [deudas]);
  const totalPendiente = useMemo(()=>deudasPend.reduce((s,d)=>s+d.monto_a_pagar,0), [deudasPend]);
  const cartItems      = useMemo(()=>Object.entries(carrito).filter(([,q])=>q>0), [carrito]);
  const cartUnits      = useMemo(()=>cartItems.reduce((s,[,q])=>s+q,0), [cartItems]);
  const cartTotal      = useMemo(()=>cartItems.reduce((s,[pid,q])=>{ const p=products.find(x=>x.id===pid); return s+(p?p.price*q:0); },0), [cartItems,products]);
  const deudasPorVend  = useMemo(()=>deudasPend.reduce((acc,d)=>{ const id=d.vendedora_id; if(!acc[id]) acc[id]={vendedora:d.vendedora,items:[],total:0}; acc[id].items.push(d); acc[id].total+=d.monto_a_pagar; return acc; },{}), [deudasPend]);
  // Productos propios del usuario con stock disponible
  const miInventario   = useMemo(()=>inventory.filter(i=>i.user_id===me.id && i.qty_available>0), [inventory, me.id]);
  const prodsFilt      = useMemo(()=>{
    const q = prodSrch.toLowerCase();
    const pidsConStock = new Set(miInventario.map(i=>i.product_id));
    return products.filter(p=>{
      if (!pidsConStock.has(p.id)) return false;
      if (!q) return true;
      return p.name.toLowerCase().includes(q) || p.sku.toLowerCase().includes(q);
    });
  }, [products, miInventario, prodSrch]);
  // Filtrar activas usando neq cancelada (coincide con lo que carga Supabase)
  const enviActivas    = useMemo(()=>enviadas.filter(c=>c.status!=="cancelada"), [enviadas]);
  const recibActivas   = useMemo(()=>recibidas.filter(c=>c.status!=="cancelada"), [recibidas]);

  const vistaOrigen = vistaInicial === "enviados" ? "main_env" : vistaInicial === "recibidos" ? "main_rec" : "main_env";
  const BtnBack = () => (
    <button onClick={()=>setView(vistaOrigen)} style={{ background:"rgba(255,255,255,.2)",border:"none",borderRadius:10,padding:"7px 14px",color:"#fff",fontFamily:"var(--hf)",fontWeight:700,fontSize:12,cursor:"pointer",marginBottom:12,display:"inline-flex",alignItems:"center",gap:6 }}>← Volver</button>
  );

  if (loading) return (
    <div style={{ display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:"60px 20px",gap:12 }}>
      <div style={{ width:32,height:32,border:"3px solid #eee",borderTopColor:"#e0224e",borderRadius:"50%",animation:"spin .7s linear infinite" }}/>
      <div style={{ fontSize:13,color:"#999" }}>Cargando...</div>
    </div>
  );

  if (dbError==="falta_tabla") return (
    <div style={{ padding:"30px 20px",maxWidth:480,margin:"0 auto" }}>
      <div style={{ background:"#fff",borderRadius:20,border:"2px solid #ffcc00",overflow:"hidden",boxShadow:"0 4px 20px rgba(0,0,0,.1)" }}>
        <div style={{ background:"linear-gradient(135deg,#e0224e,#fa1e5a)",padding:"24px 24px 20px",textAlign:"center" }}>
          <div style={{ fontSize:44,marginBottom:8 }}>🗄️</div>
          <div style={{ fontSize:18,fontWeight:900,color:"#fff" }}>Falta configurar Supabase</div>
          <div style={{ fontSize:12,color:"rgba(255,255,255,.8)",marginTop:4 }}>Las tablas de consignación no existen todavía</div>
        </div>
        <div style={{ padding:"20px 24px 24px" }}>
          <div style={{ fontSize:13,fontWeight:700,color:"#333",marginBottom:12 }}>Seguí estos pasos para activar el módulo:</div>
          {[
            {n:1, txt:"Abrí tu proyecto en supabase.com"},
            {n:2, txt:'Andá a SQL Editor → "New query"'},
            {n:3, txt:"Pegá y ejecutá el SQL que te mandé (consignacion_schema_v2.sql)"},
            {n:4, txt:"Volvé acá y recargá la página"},
          ].map(s=>(
            <div key={s.n} style={{ display:"flex",alignItems:"flex-start",gap:12,marginBottom:12 }}>
              <div style={{ width:28,height:28,borderRadius:"50%",background:"#e0224e",color:"#fff",display:"flex",alignItems:"center",justifyContent:"center",fontWeight:900,fontSize:13,flexShrink:0,marginTop:1 }}>{s.n}</div>
              <div style={{ fontSize:13,color:"#444",fontWeight:600,lineHeight:1.5,paddingTop:4 }}>{s.txt}</div>
            </div>
          ))}
          <button onClick={loadAll}
            style={{ width:"100%",marginTop:8,padding:"14px",borderRadius:14,border:"none",background:"#e0224e",color:"#fff",fontFamily:"var(--hf)",fontWeight:900,fontSize:14,cursor:"pointer" }}>
            🔄 Reintentar conexión
          </button>
        </div>
      </div>
    </div>
  );

  // ══════════════════ MAIN (ENVIADOS + RECIBIDOS) ══════════════════
  if (view==="main_env" || view==="main" || view==="main_rec") return (
    <div style={{ paddingBottom:32 }}>
      {/* Header */}
      <div style={{ background: view==="main_rec" ? "linear-gradient(135deg,#0ea5e9,#0284c7)" : "linear-gradient(135deg,#e0224e,#fa1e5a)", padding:"20px 16px 24px" }}>
        <div style={{ fontSize:22,fontWeight:900,color:"#fff",marginBottom:4 }}>
          {view==="main_rec" ? "📥 Recibidos" : "📤 Enviados"}
        </div>
        <div style={{ fontSize:12,color:"rgba(255,255,255,.8)",marginBottom:16 }}>
          {view==="main_rec" ? "Productos que te entregaron para vender" : "Productos que entregaste a otras vendedoras"}
        </div>
        <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:10 }}>
          {(view==="main_rec" ? [
            {ico:"📥",val:recibActivas.length,lbl:"Recibidas activas"},
            {ico:"⏳",val:recibActivas.reduce(function(s,c){return s+(c.items||[]).reduce(function(s2,i){return s2+i.qty_enviada-i.qty_vendida-i.qty_devuelta;},0);},0),lbl:"Unidades pendientes"},
            {ico:"💵",val:fmt(recibActivas.reduce(function(s,c){return s+(c.items||[]).reduce(function(s2,i){return s2+(i.qty_vendida*i.precio_venta*(c.comision_pct||30)/100);},0);},0)),lbl:"Ganancia estimada",sm:true},
            {ico:"✅",val:recibActivas.reduce(function(s,c){return s+(c.items||[]).reduce(function(s2,i){return s2+i.qty_vendida;},0);},0),lbl:"Unidades vendidas"},
          ] : [
            {ico:"📤",val:enviActivas.length,lbl:"Entregas activas"},
            {ico:"⏳",val:enviActivas.reduce(function(s,c){return s+(c.items||[]).reduce(function(s2,i){return s2+i.qty_enviada-i.qty_vendida-i.qty_devuelta;},0);},0),lbl:"Unidades pendientes"},
            {ico:"💰",val:fmt(totalPendiente),lbl:"Pendiente de cobro",sm:true},
            {ico:"✅",val:deudas.filter(function(d){return d.pagada;}).length,lbl:"Pagos cobrados"},
          ]).map(function(m,i){return (
            <div key={i} style={{ background:"rgba(255,255,255,.15)",borderRadius:14,padding:"12px",border:"1px solid rgba(255,255,255,.2)" }}>
              <div style={{ fontSize:20,marginBottom:2 }}>{m.ico}</div>
              <div style={{ fontFamily:"var(--mf)",fontWeight:900,fontSize:m.sm?11:20,color:"#fff",lineHeight:1 }}>{m.val}</div>
              <div style={{ fontSize:10,color:"rgba(255,255,255,.7)",marginTop:3,fontWeight:600 }}>{m.lbl}</div>
            </div>
          );})}
        </div>
      </div>

      <div style={{ padding:"14px 14px 0" }}>
        {/* CTA nueva entrega — solo en vista enviados */}
        {view!=="main_rec"&&<button onClick={()=>{ setView("nueva"); setCarrito({}); setVendedoraId(""); setNotas(""); setPaso(1); }}
          style={{ width:"100%",display:"flex",alignItems:"center",gap:14,padding:"16px",borderRadius:16,background:"linear-gradient(135deg,#ff7a00,#d97706)",border:"none",cursor:"pointer",boxShadow:"0 4px 16px rgba(255,122,0,.3)",marginBottom:20 }}>
          <div style={{ width:46,height:46,borderRadius:14,background:"rgba(255,255,255,.2)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:24,flexShrink:0 }}>📦</div>
          <div style={{ textAlign:"left",flex:1 }}>
            <div style={{ fontSize:15,fontWeight:900,color:"#fff" }}>Nueva entrega en consignación</div>
            <div style={{ fontSize:11,color:"rgba(255,255,255,.8)",marginTop:2 }}>Elegí persona, productos y cantidades</div>
          </div>
          <div style={{ color:"rgba(255,255,255,.7)",fontSize:22 }}>›</div>
        </button>}

        {/* Deudas agrupadas — solo en vista enviados */}
        {view!=="main_rec"&&Object.values(deudasPorVend).length>0&&(
          <div style={{ marginBottom:22 }}>
            <div style={{ fontSize:11,fontWeight:800,color:"#888",marginBottom:10,textTransform:"uppercase",letterSpacing:".07em" }}>💰 Pendiente de cobro</div>
            {Object.values(deudasPorVend).map(g=>(
              <div key={g.vendedora?.id} style={{ background:"#fff",borderRadius:16,border:"2px solid #ffcc00",marginBottom:10,overflow:"hidden",boxShadow:"0 2px 10px rgba(0,0,0,.07)" }}>
                <div style={{ display:"flex",alignItems:"center",gap:12,padding:"14px 16px" }}>
                  <Avatar name={g.vendedora?.name} color={g.vendedora?.color} size={44}/>
                  <div style={{ flex:1,minWidth:0 }}>
                    <div style={{ fontWeight:800,fontSize:14 }}>{g.vendedora?.name}</div>
                    <div style={{ fontSize:11,color:"#999",marginTop:2 }}>{g.items.length} venta{g.items.length!==1?"s":""} sin cobrar</div>
                  </div>
                  <div style={{ textAlign:"right",flexShrink:0 }}>
                    <div style={{ fontFamily:"var(--mf)",fontWeight:900,fontSize:18,color:"#d97706" }}>{fmt(g.total)}</div>
                    <button onClick={()=>abrirLiquidar(g.vendedora)}
                      style={{ marginTop:6,background:"#e0224e",border:"none",borderRadius:10,padding:"7px 14px",color:"#fff",fontFamily:"var(--hf)",fontWeight:800,fontSize:12,cursor:"pointer" }}>
                      Cobrar
                    </button>
                  </div>
                </div>
                <div style={{ borderTop:"1px solid #f5f5f5",padding:"8px 16px 12px" }}>
                  {g.items.slice(0,3).map(d=>{ const pid=d.product?.product_id?.id||d.item?.product_id; const p=(d.product?.product_id)||products.find(x=>x.id===pid); return (
                    <div key={d.id} style={{ display:"flex",justifyContent:"space-between",fontSize:11,color:"#666",padding:"2px 0" }}>
                      <span>📦 {p?.name??"Producto"}</span><span style={{fontWeight:700}}>{fmt(d.monto_a_pagar)}</span>
                    </div>
                  );})}
                  {g.items.length>3&&<div style={{ fontSize:10,color:"#bbb",marginTop:4 }}>+{g.items.length-3} más...</div>}
                </div>
              </div>
            ))}
          </div>
        )}

        {/* Recibidas — solo en vista recibidos */}
        {(view==="main_rec"||view==="main")&&recibActivas.length>0&&(
          <div style={{ marginBottom:22 }}>
            <div style={{ fontSize:11,fontWeight:800,color:"#888",marginBottom:10,textTransform:"uppercase",letterSpacing:".07em" }}>📥 Productos que tengo para vender</div>
            {recibActivas.map(consig=>{
              const tot=( consig.items||[]).reduce((s,i)=>s+i.qty_enviada,0);
              const vend=(consig.items||[]).reduce((s,i)=>s+i.qty_vendida,0);
              const dev= (consig.items||[]).reduce((s,i)=>s+i.qty_devuelta,0);
              const pend=tot-vend-dev;
              const ganEst=(consig.items||[]).reduce((s,i)=>s+(i.qty_vendida*i.precio_venta*consig.comision_pct/100),0);
              return (
                <div key={consig.id} style={{ background:"#fff",borderRadius:18,border:"1.5px solid #ececf0",marginBottom:12,overflow:"hidden",boxShadow:"0 2px 12px rgba(0,0,0,.05)" }}>
                  <div style={{ display:"flex",alignItems:"center",gap:12,padding:"14px 16px",background:"#fafafb",borderBottom:"1px solid #f0f0f3" }}>
                    <Avatar name={consig.owner?.name} color={consig.owner?.color} size={42}/>
                    <div style={{ flex:1,minWidth:0 }}>
                      <div style={{ fontWeight:800,fontSize:14 }}>{consig.owner?.name}</div>
                      <div style={{ fontSize:11,color:"#999",marginTop:1 }}>Te dejó <strong>{tot} u.</strong> · Ganás el <strong style={{color:"#10b981"}}>{consig.comision_pct}%</strong> de cada venta</div>
                    </div>
                    {pend>0&&<div style={{ background:"#fff7ed",color:"#d97706",borderRadius:20,padding:"4px 11px",fontSize:11,fontWeight:800,flexShrink:0 }}>Quedan {pend}</div>}
                  </div>
                  <div style={{ padding:"12px 16px 6px" }}>
                    <ProgBar vendidas={vend} devueltas={dev} total={tot}/>
                    <div style={{ display:"flex",justifyContent:"space-between",marginTop:6,fontSize:10.5,color:"#999",fontWeight:600 }}>
                      <span>✅ {vend} vendidos</span><span>⏳ {pend} por vender</span><span>↩️ {dev} devueltos</span>
                    </div>
                  </div>
                  {/* Lista de productos de esta entrega */}
                  <div style={{ padding:"6px 14px 2px" }}>
                    {(consig.items||[]).map(it=>{
                      const pr=it.product||products.find(x=>x.id===it.product_id);
                      const restan=it.qty_enviada-it.qty_vendida-it.qty_devuelta;
                      return (
                        <div key={it.id} style={{ display:"flex",alignItems:"center",gap:10,padding:"7px 0",borderBottom:"1px solid #f4f4f6" }}>
                          <ProdThumb prod={pr} size={38}/>
                          <div style={{ flex:1,minWidth:0 }}>
                            <div style={{ fontSize:12.5,fontWeight:800,color:"#222",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" }}>{pr?.name||"Producto"}</div>
                            <div style={{ fontSize:10.5,color:"#999",marginTop:1 }}>{fmt(it.precio_venta)} c/u</div>
                          </div>
                          <div style={{ flexShrink:0,textAlign:"right" }}>
                            <div style={{ fontFamily:"var(--mf)",fontWeight:900,fontSize:15,color:restan>0?"#d97706":"#bbb" }}>{restan}</div>
                            <div style={{ fontSize:9,color:"#bbb",fontWeight:700,textTransform:"uppercase" }}>quedan</div>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                  {ganEst>0&&(
                    <div style={{ margin:"8px 14px 8px",background:"#ecfdf5",borderRadius:12,padding:"10px 14px",display:"flex",justifyContent:"space-between",alignItems:"center" }}>
                      <span style={{ fontSize:11.5,color:"#10b981",fontWeight:700 }}>💵 Ya ganaste</span>
                      <span style={{ fontFamily:"var(--mf)",fontWeight:900,fontSize:15,color:"#059669" }}>{fmt(ganEst)}</span>
                    </div>
                  )}
                  <div style={{ padding:"0 14px 14px" }}>
                    <button onClick={()=>abrirDetRec(consig)}
                      style={{ width:"100%",padding:"14px",borderRadius:14,background:pend>0?"linear-gradient(145deg,#e0224e,#fa1e5a)":"#f3f3f6",border:"none",color:pend>0?"#fff":"#888",fontFamily:"var(--hf)",fontWeight:800,fontSize:14,cursor:"pointer",boxShadow:pend>0?"0 4px 14px rgba(224,34,78,.28)":"none" }}>
                      {pend>0 ? "💵 Marcar lo que vendí" : "Ver detalle"}
                    </button>
                  </div>
                </div>
              );
            })}
          </div>
        )}

        {/* Entregas enviadas — solo en vista enviados */}
        {view!=="main_rec"&&enviActivas.length>0&&(
          <div style={{ marginBottom:20 }}>
            <div style={{ fontSize:11,fontWeight:800,color:"#888",marginBottom:10,textTransform:"uppercase",letterSpacing:".07em" }}>📤 Lo que entregué</div>
            {enviActivas.map(consig=>{
              const tot=(consig.items||[]).reduce((s,i)=>s+i.qty_enviada,0);
              const vend=(consig.items||[]).reduce((s,i)=>s+i.qty_vendida,0);
              const dev=(consig.items||[]).reduce((s,i)=>s+i.qty_devuelta,0);
              const pend=tot-vend-dev;
              // Deudas de esta consignación
              const mis_deudas = deudas.filter(d=>d.consignacion_id===consig.id);
              const deudas_pend = mis_deudas.filter(d=>!d.pagada);
              const deudas_pagas = mis_deudas.filter(d=>d.pagada);
              const total_vendido = mis_deudas.reduce((s,d)=>s+d.monto_total,0);
              const total_comision = mis_deudas.reduce((s,d)=>s+d.comision,0);
              const total_a_cobrar = deudas_pend.reduce((s,d)=>s+d.monto_a_pagar,0);
              const total_cobrado  = deudas_pagas.reduce((s,d)=>s+d.monto_a_pagar,0);
              const comision_pct = consig.comision_pct||30;
              return (
                <div key={consig.id}
                  style={{ background:"#fff",borderRadius:16,border:"1.5px solid #e0e0e0",marginBottom:12,overflow:"hidden",boxShadow:"0 2px 10px rgba(0,0,0,.07)" }}>
                  {/* Header con vendedora */}
                  <div style={{ display:"flex",alignItems:"center",gap:12,padding:"13px 16px",background:"#fafafa",borderBottom:"1px solid #eee",cursor:"pointer" }}
                    onClick={()=>abrirDetEnv(consig)}>
                    <Avatar name={consig.vendedora?.name} color={consig.vendedora?.color} size={44}/>
                    <div style={{ flex:1,minWidth:0 }}>
                      <div style={{ fontWeight:800,fontSize:14 }}>{consig.vendedora?.name}</div>
                      <div style={{ fontSize:11,color:"#999",marginTop:1 }}>{tot} u. entregadas · {new Date(consig.created_at).toLocaleDateString("es-AR")}</div>
                    </div>
                    <div style={{ textAlign:"right",flexShrink:0 }}>
                      <div style={{ fontSize:10,color:"#999" }}>Pend.</div>
                      <div style={{ fontFamily:"var(--mf)",fontWeight:900,fontSize:22,color:pend>0?"#d97706":"#10b981" }}>{pend}</div>
                    </div>
                    <div style={{ color:"#ccc",fontSize:20 }}>›</div>
                  </div>
                  <div style={{ padding:"4px 16px 10px" }}>
                    <ProgBar vendidas={vend} devueltas={dev} total={tot}/>
                    <div style={{ display:"flex",justifyContent:"space-between",marginTop:5,fontSize:10,color:"#999",fontWeight:600 }}>
                      <span>✅ {vend} vendidos</span><span>⏳ {pend} pendientes</span><span>↩️ {dev} devueltos</span>
                    </div>
                  </div>
                  {/* ── Desglose financiero ── */}
                  {vend>0&&(
                    <div style={{ margin:"0 12px 12px",borderRadius:12,overflow:"hidden",border:"1px solid #e0e0e0" }}>
                      {/* Total vendido */}
                      <div style={{ display:"flex",justifyContent:"space-between",alignItems:"center",padding:"9px 14px",background:"#f8f8f8",borderBottom:"1px solid #eee" }}>
                        <span style={{ fontSize:12,color:"#555",fontWeight:700 }}>💰 Total vendido por ella</span>
                        <span style={{ fontFamily:"var(--mf)",fontWeight:800,fontSize:13,color:"#222" }}>{fmt(total_vendido)}</span>
                      </div>
                      {/* Comisión vendedora */}
                      <div style={{ display:"flex",justifyContent:"space-between",alignItems:"center",padding:"9px 14px",background:"#fff8e1",borderBottom:"1px solid #eee" }}>
                        <div>
                          <div style={{ fontSize:12,color:"#d97706",fontWeight:700 }}>🤝 Comisión vendedora ({comision_pct}%)</div>
                          <div style={{ fontSize:10,color:"#bbb",marginTop:1 }}>Se descuenta de lo que te debe</div>
                        </div>
                        <span style={{ fontFamily:"var(--mf)",fontWeight:800,fontSize:13,color:"#d97706" }}>− {fmt(total_comision)}</span>
                      </div>
                      {/* Deuda */}
                      <div style={{ display:"flex",justifyContent:"space-between",alignItems:"center",padding:"10px 14px",background:total_a_cobrar>0?"#fde8ea":"#e8faf4" }}>
                        <div>
                          <div style={{ fontSize:12,fontWeight:800,color:total_a_cobrar>0?"#dc2626":"#059669" }}>
                            {total_a_cobrar>0?"🔴 Deuda pendiente":"✅ Todo cobrado"}
                          </div>
                          {total_cobrado>0&&<div style={{ fontSize:10,color:"#888",marginTop:1 }}>Ya cobrado: {fmt(total_cobrado)}</div>}
                        </div>
                        {total_a_cobrar>0&&(
                          <div style={{ textAlign:"right" }}>
                            <div style={{ fontFamily:"var(--mf)",fontWeight:900,fontSize:16,color:"#dc2626" }}>{fmt(total_a_cobrar)}</div>
                            <button onClick={e=>{ e.stopPropagation(); abrirLiquidar(consig.vendedora); }}
                              style={{ marginTop:6,background:"#e0224e",border:"none",borderRadius:10,padding:"6px 14px",color:"#fff",fontFamily:"var(--hf)",fontWeight:800,fontSize:12,cursor:"pointer" }}>
                              Cobrar
                            </button>
                          </div>
                        )}
                      </div>
                      {/* Detalle ventas pendientes */}
                      {deudas_pend.length>0&&(
                        <div style={{ padding:"8px 14px 10px",background:"#fafafa",borderTop:"1px solid #eee" }}>
                          <div style={{ fontSize:10,fontWeight:800,color:"#aaa",marginBottom:6,textTransform:"uppercase",letterSpacing:".05em" }}>Ventas sin cobrar</div>
                          {deudas_pend.map((d,idx)=>{
                            // Buscar producto: desde join anidado, desde item.product_id, o desde lista
                            const pid = d.product?.product_id?.id || d.item?.product_id;
                            const p = (d.product?.product_id) || products.find(x=>x.id===pid);
                            return (
                              <div key={d.id} style={{ display:"flex",justifyContent:"space-between",alignItems:"center",padding:"6px 0",borderBottom:idx<deudas_pend.length-1?"1px solid #f0f0f0":"none" }}>
                                <div style={{ display:"flex",alignItems:"center",gap:8 }}>
                                  <div style={{ width:36,height:36,borderRadius:9,background:"#f5f5f5",display:"flex",alignItems:"center",justifyContent:"center",fontSize:18,flexShrink:0 }}>{p?.emoji||"📦"}</div>
                                  <div>
                                    <div style={{ fontSize:12,fontWeight:800,color:"#222" }}>{p?.name||"Producto"}</div>
                                    <div style={{ fontSize:10,color:"#aaa",marginTop:1 }}>Precio: {fmt(d.monto_total)} · Com: {fmt(d.comision)}</div>
                                  </div>
                                </div>
                                <div style={{ fontFamily:"var(--mf)",fontWeight:800,fontSize:13,color:"#9b1c1c",flexShrink:0 }}>{fmt(d.monto_a_pagar)}</div>
                              </div>
                            );
                          })}
                        </div>
                      )}
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        )}

        {view==="main_rec" && recibActivas.length===0 && (
          <div style={{ textAlign:"center",padding:"48px 28px",color:"#999" }}>
            <div style={{ fontSize:52,marginBottom:12 }}>📭</div>
            <div style={{ fontSize:15,fontWeight:800,color:"#444",marginBottom:6 }}>Todavía no tenés productos para vender</div>
            <div style={{ fontSize:13,lineHeight:1.5 }}>Cuando alguien te deje productos en consignación, te van a aparecer acá y te avisamos con una notificación 🔔</div>
          </div>
        )}
        {view!=="main_rec" && enviActivas.length===0 && (
          <div style={{ textAlign:"center",padding:"32px 20px 16px",color:"#bbb",fontSize:13 }}>
            <div style={{ fontSize:44,marginBottom:8 }}>📭</div>
            No entregaste productos todavía.<br/>Tocá el botón de arriba para crear la primera entrega.
          </div>
        )}

        {/* ── CARGA MANUAL DE STOCK PROPIO ─────────────────────────── */}
        <div style={{ marginBottom:16 }}>
          <button onClick={()=>setCargaShow(v=>!v)}
            style={{ width:"100%",display:"flex",alignItems:"center",gap:12,padding:"14px 16px",borderRadius:14,background:"#fff",border:"1.5px solid #e0e0e0",cursor:"pointer",boxShadow:"0 2px 8px rgba(0,0,0,.05)" }}>
            <div style={{ width:40,height:40,borderRadius:12,background:"#e8faf4",display:"flex",alignItems:"center",justifyContent:"center",fontSize:20,flexShrink:0 }}>📦</div>
            <div style={{ textAlign:"left",flex:1 }}>
              <div style={{ fontSize:14,fontWeight:800,color:"#222" }}>Cargar stock propio</div>
              <div style={{ fontSize:11,color:"#999",marginTop:1 }}>Sumar unidades a mi inventario</div>
            </div>
            <div style={{ color:"#ccc",fontSize:18 }}>{cargaShow?"▲":"▼"}</div>
          </button>
          {cargaShow&&(
            <div style={{ background:"#f8fffe",border:"1.5px solid rgba(0,184,122,.2)",borderRadius:14,padding:"14px 16px",marginTop:8 }}>
              <div style={{ fontSize:12,fontWeight:800,color:"#555",marginBottom:10 }}>Elegí producto y cantidad</div>
              <select value={cargaPid} onChange={e=>setCargaPid(e.target.value)}
                style={{ width:"100%",padding:"11px 12px",borderRadius:10,border:"1.5px solid #e0e0e0",fontFamily:"var(--hf)",fontSize:14,marginBottom:10,outline:"none",color:"#222",background:"#fff" }}>
                <option value="">— Seleccionar producto —</option>
                {products.filter(p=>p.is_active!==false).map(p=>(
                  <option key={p.id} value={p.id}>{p.emoji||"📦"} {p.name} ({p.sku})</option>
                ))}
              </select>
              <div style={{ display:"flex",gap:10,alignItems:"center" }}>
                <div style={{ display:"flex",alignItems:"center",border:"1.5px solid #e0e0e0",borderRadius:10,overflow:"hidden",flexShrink:0 }}>
                  <button onClick={()=>setCargaQty(q=>Math.max(1,q-1))}
                    style={{ width:36,height:36,border:"none",background:"#f5f5f5",cursor:"pointer",fontSize:20,display:"flex",alignItems:"center",justifyContent:"center",color:"#444" }}>−</button>
                  <div style={{ minWidth:40,textAlign:"center",fontFamily:"var(--mf)",fontWeight:700,fontSize:16,color:"#e0224e" }}>{cargaQty}</div>
                  <button onClick={()=>setCargaQty(q=>q+1)}
                    style={{ width:36,height:36,border:"none",background:"#ffeaea",cursor:"pointer",fontSize:20,display:"flex",alignItems:"center",justifyContent:"center",color:"#e0224e" }}>+</button>
                </div>
                <button onClick={cargarStockPropio} disabled={!cargaPid||cargaSaving}
                  style={{ flex:1,padding:"11px",borderRadius:10,border:"none",background:"#10b981",color:"#fff",fontFamily:"var(--hf)",fontWeight:800,fontSize:14,cursor:"pointer",opacity:(!cargaPid||cargaSaving)?.5:1 }}>
                  {cargaSaving?"Guardando...":"✅ Confirmar carga"}
                </button>
              </div>
            </div>
          )}
        </div>

        {/* ── HISTORIAL DE TRANSFERS ────────────────────────────────── */}
        {(()=>{
          const txFilt = view==="main_rec"
            ? transfers.filter(t=>t.to_user_id===me.id)
            : transfers.filter(t=>t.from_user_id===me.id);
          if (txFilt.length===0) return null;
          const visible = txExpanded ? txFilt : txFilt.slice(0,5);
          const badgeCfg = {
            pending:    {bg:"#fff8e1",col:"#d97706",txt:"⏳ Pendiente"},
            pendiente:  {bg:"#fff8e1",col:"#d97706",txt:"⏳ Pendiente"},
            confirmed:  {bg:"#e8faf4",col:"#059669",txt:"✅ Confirmado"},
            aceptado:   {bg:"#e8faf4",col:"#059669",txt:"✅ Aceptado"},
            cancelled:  {bg:"#fde8ea",col:"#dc2626",txt:"❌ Cancelado"},
            rechazado:  {bg:"#fde8ea",col:"#dc2626",txt:"❌ Rechazado"},
          };
          return (
            <div style={{ marginBottom:20 }}>
              <div style={{ fontSize:11,fontWeight:800,color:"#888",marginBottom:10,textTransform:"uppercase",letterSpacing:".07em" }}>
                {view==="main_rec" ? "📩 Envíos que recibí" : "📤 Historial de envíos"}
                <span style={{ marginLeft:8,background:"#e0e0e0",color:"#555",borderRadius:20,padding:"2px 8px",fontSize:10,fontWeight:700 }}>{txFilt.length}</span>
              </div>
              <div style={{ background:"#fff",borderRadius:16,border:"1.5px solid #e0e0e0",overflow:"hidden",boxShadow:"0 2px 10px rgba(0,0,0,.07)" }}>
                {visible.map((tx,i)=>{
                  const p   = tx.product;
                  const usr = view==="main_rec" ? tx.from_user : tx.to_user;
                  const badg = badgeCfg[tx.status] || {bg:"#f0f0f0",col:"#888",txt:tx.status};
                  return (
                    <div key={tx.id} style={{ display:"flex",alignItems:"center",gap:12,padding:"12px 14px",borderBottom:i<visible.length-1?"1px solid #f5f5f5":"none" }}>
                      {/* Producto */}
                      <div style={{ width:40,height:40,borderRadius:10,background:"#f0f0f0",display:"flex",alignItems:"center",justifyContent:"center",fontSize:22,flexShrink:0 }}>
                        {p?.emoji||"📦"}
                      </div>
                      {/* Info */}
                      <div style={{ flex:1,minWidth:0 }}>
                        <div style={{ fontWeight:700,fontSize:13,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",color:"#222" }}>
                          {p?.name||"Producto"}
                        </div>
                        <div style={{ fontSize:11,color:"#999",marginTop:1,display:"flex",gap:6,alignItems:"center" }}>
                          <span>{view==="main_rec"?"De:":"Para:"} <strong style={{color:"#444"}}>{usr?.name||"?"}</strong></span>
                          <span>·</span>
                          <span style={{ fontFamily:"var(--mf)",fontWeight:700 }}>{tx.qty} u.</span>
                          <span>·</span>
                          <span>{new Date(tx.created_at).toLocaleDateString("es-AR")}</span>
                        </div>
                      </div>
                      {/* Badge estado */}
                      <div style={{ background:badg.bg,color:badg.col,borderRadius:20,padding:"4px 10px",fontSize:10,fontWeight:800,flexShrink:0,whiteSpace:"nowrap" }}>
                        {badg.txt}
                      </div>
                    </div>
                  );
                })}
                {txFilt.length>5&&(
                  <div onClick={()=>setTxExpanded(v=>!v)}
                    style={{ padding:"12px",textAlign:"center",fontSize:12,color:"#e0224e",fontWeight:700,cursor:"pointer",borderTop:"1px solid #f5f5f5" }}>
                    {txExpanded?"▲ Ver menos":"▼ Ver todos ("+txFilt.length+")"}
                  </div>
                )}
              </div>
            </div>
          );
        })()}

      </div>
    </div>
  );

  // ══════════════════ NUEVA (wizard) ══════════════════
  if (view==="nueva") {
    const vendedora=contacts.find(c=>c.id===vendedoraId);
    const pasos=["Persona","Productos","Confirmar"];
    return (
      <div style={{ paddingBottom:32 }}>
        <div style={{ background:"linear-gradient(135deg,#e0224e,#fa1e5a)",padding:"16px 16px 20px" }}>
          <BtnBack/>
          <div style={{ fontSize:18,fontWeight:900,color:"#fff",marginBottom:12 }}>📦 Nueva entrega</div>
          <div style={{ display:"flex",gap:6 }}>
            {pasos.map((p,i)=>(
              <div key={i} style={{ flex:1,display:"flex",flexDirection:"column",alignItems:"center",gap:4 }}>
                <div style={{ width:28,height:28,borderRadius:"50%",background:paso>i+1?"rgba(255,255,255,.9)":paso===i+1?"#fff":"rgba(255,255,255,.3)",display:"flex",alignItems:"center",justifyContent:"center",fontWeight:900,fontSize:13,color:paso===i+1?"#e0224e":"rgba(255,255,255,.7)",transition:"all .2s" }}>
                  {paso>i+1?"✓":(i+1)}
                </div>
                <div style={{ fontSize:9,color:paso===i+1?"#fff":"rgba(255,255,255,.5)",fontWeight:700,textTransform:"uppercase",letterSpacing:".04em" }}>{p}</div>
              </div>
            ))}
          </div>
        </div>
        <div style={{ padding:"16px 14px 0" }}>

          {/* PASO 1 */}
          {paso===1&&(
            <div>
              <div style={{ fontSize:13,fontWeight:800,color:"#444",marginBottom:12 }}>¿A quién le entregás?</div>
              {contacts.length===0&&<div style={{ textAlign:"center",padding:"30px",color:"#bbb",fontSize:13 }}>Sin contactos en tu Red.</div>}
              {contacts.map(c=>(
                <div key={c.id} onClick={()=>setVendedoraId(c.id)}
                  style={{ display:"flex",alignItems:"center",gap:12,padding:"14px",borderRadius:14,border:`2px solid ${vendedoraId===c.id?"#e0224e":"#e0e0e0"}`,background:vendedoraId===c.id?"#ffeaea":"#fff",marginBottom:10,cursor:"pointer",transition:"all .15s" }}>
                  <Avatar name={c.name} color={c.color} size={46}/>
                  <div style={{ flex:1 }}>
                    <div style={{ fontWeight:800,fontSize:14 }}>{c.name}</div>
                    <div style={{ fontSize:11,color:"#999",marginTop:1 }}>{c.email}</div>
                  </div>
                  {vendedoraId===c.id&&<div style={{ width:24,height:24,borderRadius:"50%",background:"#e0224e",display:"flex",alignItems:"center",justifyContent:"center",color:"#fff",fontWeight:900,fontSize:14,flexShrink:0 }}>✓</div>}
                </div>
              ))}
              <div style={{ background:"#fff3e0",borderRadius:14,padding:"14px",marginTop:4,marginBottom:16 }}>
                <div style={{ fontSize:12,fontWeight:800,color:"#d97706",marginBottom:10 }}>💰 Comisión de la vendedora</div>
                <div style={{ display:"flex",alignItems:"center",gap:12 }}>
                  <input type="range" min="0" max="50" step="5" value={comision} onChange={e=>setComision(Number(e.target.value))} style={{ flex:1,accentColor:"#e0224e" }}/>
                  <div style={{ fontFamily:"var(--mf)",fontWeight:900,fontSize:26,color:"#e0224e",minWidth:52,textAlign:"right" }}>{comision}%</div>
                </div>
                <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginTop:10 }}>
                  <div style={{ background:"#fff",borderRadius:10,padding:"10px",textAlign:"center" }}>
                    <div style={{ fontSize:10,color:"#999",fontWeight:700 }}>Ella cobra</div>
                    <div style={{ fontWeight:900,fontSize:15,color:"#d97706",marginTop:2 }}>{comision}%</div>
                  </div>
                  <div style={{ background:"#fff",borderRadius:10,padding:"10px",textAlign:"center" }}>
                    <div style={{ fontSize:10,color:"#999",fontWeight:700 }}>Vos cobrás</div>
                    <div style={{ fontWeight:900,fontSize:15,color:"#10b981",marginTop:2 }}>{100-comision}%</div>
                  </div>
                </div>
              </div>
              <button onClick={()=>vendedoraId&&setPaso(2)} disabled={!vendedoraId}
                style={{ width:"100%",padding:"16px",borderRadius:14,border:"none",background:"#e0224e",color:"#fff",fontFamily:"var(--hf)",fontWeight:900,fontSize:15,cursor:vendedoraId?"pointer":"not-allowed",opacity:vendedoraId?1:.4 }}>
                Siguiente → Elegir productos
              </button>
            </div>
          )}

          {/* PASO 2 */}
          {paso===2&&(
            <div>
              <div style={{ display:"flex",alignItems:"center",gap:10,marginBottom:14,padding:"10px 12px",background:"#ffeaea",borderRadius:12,border:"1px solid rgba(204,0,0,.2)" }}>
                <Avatar name={vendedora?.name} color={vendedora?.color} size={32}/>
                <div style={{ fontSize:13,fontWeight:700 }}>Entregando a <strong>{vendedora?.name}</strong></div>
                <button onClick={()=>setPaso(1)} style={{ marginLeft:"auto",background:"none",border:"none",color:"#e0224e",fontWeight:700,fontSize:12,cursor:"pointer" }}>Cambiar</button>
              </div>
              <div style={{ position:"relative",marginBottom:12 }}>
                <span style={{ position:"absolute",left:12,top:"50%",transform:"translateY(-50%)",fontSize:16,color:"#999" }}>🔍</span>
                <input style={{ width:"100%",padding:"11px 12px 11px 38px",borderRadius:12,border:"1.5px solid #e0e0e0",fontFamily:"var(--hf)",fontSize:14,outline:"none",color:"#333" }}
                  placeholder="Buscar producto..." value={prodSrch} onChange={e=>setProdSrch(e.target.value)}/>
              </div>
              <div style={{ marginBottom:16 }}>
                {prodsFilt.map(p=>{
                  const inv=inventory.find(i=>i.product_id===p.id && i.user_id===me.id);
                  const avail=inv?.qty_available??0;
                  const inC=carrito[p.id]??0;
                  // No mostrar si no hay stock (ya filtrado en prodsFilt, pero por si acaso)
                  if (avail===0&&inC===0) return null;
                  return (
                    <div key={p.id} style={{ display:"flex",alignItems:"center",gap:12,padding:"11px 13px",borderRadius:14,border:`1.5px solid ${inC>0?"#e0224e":"#e0e0e0"}`,background:inC>0?"#ffeaea":"#fff",marginBottom:8,boxShadow:"0 1px 6px rgba(0,0,0,.05)" }}>
                      <ProdThumb prod={p} size={40}/>
                      <div style={{ flex:1,minWidth:0 }}>
                        <div style={{ fontWeight:700,fontSize:13,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" }}>{p.name}</div>
                        <div style={{ fontSize:10,color:"#999",marginTop:1 }}>{fmt(p.price)} · {avail} disp.</div>
                      </div>
                      <div style={{ display:"flex",alignItems:"center",border:"1.5px solid #e0e0e0",borderRadius:10,overflow:"hidden",flexShrink:0 }}>
                        <button onClick={()=>setCarrito(c=>{ const n={...c}; if((n[p.id]??0)<=1){delete n[p.id];}else{n[p.id]--;} return n; })}
                          style={{ width:32,height:32,border:"none",background:"#f5f5f5",cursor:"pointer",fontSize:18,display:"flex",alignItems:"center",justifyContent:"center",color:"#444" }}>−</button>
                        <div style={{ minWidth:30,textAlign:"center",fontFamily:"var(--mf)",fontWeight:700,fontSize:14,color:"#e0224e" }}>{inC||0}</div>
                        <button onClick={()=>{ if(inC>=avail)return; setCarrito(c=>({...c,[p.id]:(c[p.id]??0)+1})); }} disabled={inC>=avail}
                          style={{ width:32,height:32,border:"none",background:inC>=avail?"#f5f5f5":"#ffeaea",cursor:inC>=avail?"not-allowed":"pointer",fontSize:18,display:"flex",alignItems:"center",justifyContent:"center",color:"#e0224e",opacity:inC>=avail?.3:1 }}>+</button>
                      </div>
                    </div>
                  );
                })}
              </div>
              {cartUnits>0&&(
                <div style={{ background:"#e8faf4",border:"1.5px solid rgba(0,184,122,.2)",borderRadius:14,padding:"12px 14px",marginBottom:14 }}>
                  <div style={{ fontWeight:800,fontSize:12,color:"#059669",marginBottom:8 }}>✅ {cartUnits} unidades seleccionadas</div>
                  {cartItems.map(([pid,qty])=>{ const p=products.find(x=>x.id===pid); return (
                    <div key={pid} style={{ display:"flex",justifyContent:"space-between",fontSize:12,color:"#555",marginBottom:3 }}>
                      <span>{p?.emoji} {p?.name}</span><span style={{fontWeight:700}}>{qty} u.</span>
                    </div>
                  );})}
                </div>
              )}
              <div style={{ display:"flex",gap:10 }}>
                <button onClick={()=>setPaso(1)} style={{ flex:1,padding:"14px",borderRadius:14,border:"1.5px solid #e0e0e0",background:"#fff",fontFamily:"var(--hf)",fontWeight:700,fontSize:14,cursor:"pointer",color:"#666" }}>← Atrás</button>
                <button onClick={()=>cartUnits>0&&setPaso(3)} disabled={cartUnits===0}
                  style={{ flex:2,padding:"14px",borderRadius:14,border:"none",background:"#e0224e",color:"#fff",fontFamily:"var(--hf)",fontWeight:900,fontSize:14,cursor:cartUnits>0?"pointer":"not-allowed",opacity:cartUnits>0?1:.4 }}>
                  Ver resumen →
                </button>
              </div>
            </div>
          )}

          {/* PASO 3 */}
          {paso===3&&(
            <div>
              <div style={{ fontSize:14,fontWeight:800,color:"#222",marginBottom:14 }}>Resumen de la entrega</div>
              <div style={{ display:"flex",alignItems:"center",gap:12,padding:"14px",background:"#fff",borderRadius:14,border:"1.5px solid #e0e0e0",marginBottom:12,boxShadow:"0 1px 6px rgba(0,0,0,.05)" }}>
                <Avatar name={vendedora?.name} color={vendedora?.color} size={46}/>
                <div>
                  <div style={{ fontSize:11,color:"#999",fontWeight:600 }}>Le entregás a</div>
                  <div style={{ fontWeight:900,fontSize:16 }}>{vendedora?.name}</div>
                  <div style={{ fontSize:11,color:"#10b981",fontWeight:700,marginTop:1 }}>Su comisión: {comision}% · La tuya: {100-comision}%</div>
                </div>
              </div>
              <div style={{ background:"#fff",borderRadius:14,border:"1.5px solid #e0e0e0",marginBottom:12,overflow:"hidden" }}>
                {cartItems.map(([pid,qty],i)=>{ const p=products.find(x=>x.id===pid); return (
                  <div key={pid} style={{ display:"flex",alignItems:"center",gap:12,padding:"12px 14px",borderBottom:i<cartItems.length-1?"1px solid #f5f5f5":"none" }}>
                    <ProdThumb prod={p} size={36}/>
                    <div style={{ flex:1 }}>
                      <div style={{ fontWeight:700,fontSize:13 }}>{p?.name}</div>
                      <div style={{ fontSize:11,color:"#999" }}>{fmt(p?.price??0)} c/u</div>
                    </div>
                    <div style={{ textAlign:"right" }}>
                      <div style={{ fontFamily:"var(--mf)",fontWeight:900,fontSize:16,color:"#e0224e" }}>{qty}u</div>
                      <div style={{ fontSize:10,color:"#999" }}>{fmt((p?.price??0)*qty)}</div>
                    </div>
                  </div>
                );})}
                <div style={{ padding:"12px 14px",background:"#f8f8f8",display:"grid",gridTemplateColumns:"1fr 1fr",gap:8 }}>
                  <div style={{ textAlign:"center" }}>
                    <div style={{ fontSize:10,color:"#999",fontWeight:700 }}>Total mercadería</div>
                    <div style={{ fontFamily:"var(--mf)",fontWeight:900,fontSize:15,marginTop:2 }}>{fmt(cartTotal)}</div>
                  </div>
                  <div style={{ textAlign:"center" }}>
                    <div style={{ fontSize:10,color:"#999",fontWeight:700 }}>Vos cobrás ({100-comision}%)</div>
                    <div style={{ fontFamily:"var(--mf)",fontWeight:900,fontSize:15,color:"#10b981",marginTop:2 }}>{fmt(cartTotal*(100-comision)/100)}</div>
                  </div>
                </div>
              </div>
              <textarea style={{ width:"100%",padding:"11px 14px",borderRadius:12,border:"1.5px solid #e0e0e0",fontFamily:"var(--hf)",fontSize:14,outline:"none",resize:"vertical",minHeight:68,color:"#333",marginBottom:14 }}
                placeholder="Nota opcional (ej: productos para la feria del 15/06)" value={notas} onChange={e=>setNotas(e.target.value)}/>
              <div style={{ display:"flex",gap:10 }}>
                <button onClick={()=>setPaso(2)} style={{ flex:1,padding:"14px",borderRadius:14,border:"1.5px solid #e0e0e0",background:"#fff",fontFamily:"var(--hf)",fontWeight:700,fontSize:14,cursor:"pointer",color:"#666" }}>← Atrás</button>
                <button onClick={crearConsignacion} disabled={saving}
                  style={{ flex:2,padding:"14px",borderRadius:14,border:"none",background:"#10b981",color:"#fff",fontFamily:"var(--hf)",fontWeight:900,fontSize:14,cursor:saving?"not-allowed":"pointer",display:"flex",alignItems:"center",justifyContent:"center",gap:8,boxShadow:"0 4px 14px rgba(0,184,122,.3)" }}>
                  {saving?<Spin/>:"✅"} {saving?"Creando...":"Confirmar entrega"}
                </button>
              </div>
            </div>
          )}
        </div>
      </div>
    );
  }

  // ══════════════════ DETALLE RECIBIDA ══════════════════
  if (view==="detalle_rec"&&detRec) {
    const tot=detRecItems.reduce((s,i)=>s+i.qty_enviada,0);
    const vend=detRecItems.reduce((s,i)=>s+i.qty_vendida,0);
    const dev=detRecItems.reduce((s,i)=>s+i.qty_devuelta,0);
    const pend=tot-vend-dev;
    const gan=detRecItems.reduce((s,i)=>s+(i.qty_vendida*i.precio_venta*detRec.comision_pct/100),0);
    return (
      <div style={{ paddingBottom:32 }}>
        <div style={{ background:"linear-gradient(135deg,#0ea5e9,#0284c7)",padding:"16px 16px 20px" }}>
          <BtnBack/>
          <div style={{ display:"flex",alignItems:"center",gap:12,marginBottom:12 }}>
            <Avatar name={detRec.owner?.name} color={detRec.owner?.color} size={46}/>
            <div>
              <div style={{ fontSize:17,fontWeight:900,color:"#fff" }}>{detRec.owner?.name}</div>
              <div style={{ fontSize:11,color:"rgba(255,255,255,.85)",marginTop:1 }}>Te dejó estos productos · Ganás el <strong style={{color:"#7df3c0"}}>{detRec.comision_pct}%</strong></div>
            </div>
          </div>
          <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8 }}>
            {[{val:pend,lbl:"Pendientes",col:"#fff"},{val:vend,lbl:"Vendidos",col:"#7df3c0"},{val:dev,lbl:"Devueltos",col:"rgba(255,255,255,.6)"}].map((m,i)=>(
              <div key={i} style={{ background:"rgba(255,255,255,.15)",borderRadius:12,padding:"10px",textAlign:"center" }}>
                <div style={{ fontFamily:"var(--mf)",fontWeight:900,fontSize:22,color:m.col }}>{m.val}</div>
                <div style={{ fontSize:9,color:"rgba(255,255,255,.7)",fontWeight:700,textTransform:"uppercase",letterSpacing:".05em",marginTop:2 }}>{m.lbl}</div>
              </div>
            ))}
          </div>
          {gan>0&&(
            <div style={{ marginTop:12,background:"rgba(125,243,192,.2)",borderRadius:12,padding:"10px 14px",display:"flex",justifyContent:"space-between",alignItems:"center" }}>
              <span style={{ fontSize:12,color:"#7df3c0",fontWeight:700 }}>💵 Tu ganancia acumulada</span>
              <span style={{ fontFamily:"var(--mf)",fontWeight:900,fontSize:16,color:"#fff" }}>{fmt(gan)}</span>
            </div>
          )}
        </div>
        <div style={{ padding:"14px 14px 0" }}>
          <div style={{ fontSize:12,fontWeight:800,color:"#888",marginBottom:12,textTransform:"uppercase",letterSpacing:".06em" }}>Cada vez que vendas, tocá "¡Vendí 1!"</div>
          {detRecItems.map(item=>{
            const p=item.product||products.find(x=>x.id===item.product_id);
            const disp=item.qty_enviada-item.qty_vendida-item.qty_devuelta;
            const miComis=item.precio_venta*detRec.comision_pct/100;
            const isAnim=sellAnim===item.id;
            return (
              <div key={item.id} style={{ background:"#fff",borderRadius:16,border:`1.5px solid ${disp>0?"#e0e0e0":"#f0f0f0"}`,marginBottom:14,overflow:"hidden",boxShadow:"0 2px 10px rgba(0,0,0,.07)",opacity:disp===0?.7:1 }}>
                <div style={{ display:"flex",alignItems:"center",gap:12,padding:"13px 14px",borderBottom:"1px solid #f5f5f5" }}>
                  <ProdThumb prod={p} size={46}/>
                  <div style={{ flex:1,minWidth:0 }}>
                    <div style={{ fontWeight:800,fontSize:13,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" }}>{p?.name}</div>
                    <div style={{ fontSize:11,color:"#999",marginTop:2 }}>Precio: <strong>{fmt(item.precio_venta)}</strong> · Tu comisión: <strong style={{color:"#10b981"}}>{fmt(miComis)}</strong></div>
                  </div>
                  <div style={{ textAlign:"center",background:disp>0?"#fff8e1":"#f5f5f5",borderRadius:12,padding:"8px 12px",flexShrink:0 }}>
                    <div style={{ fontFamily:"var(--mf)",fontWeight:900,fontSize:24,color:disp>0?"#d97706":"#bbb",lineHeight:1 }}>{disp}</div>
                    <div style={{ fontSize:9,color:"#999",fontWeight:700,textTransform:"uppercase",marginTop:2 }}>disp.</div>
                  </div>
                </div>
                <div style={{ padding:"8px 14px 4px" }}>
                  <ProgBar vendidas={item.qty_vendida} devueltas={item.qty_devuelta} total={item.qty_enviada}/>
                  <div style={{ display:"flex",gap:14,marginTop:5,fontSize:10,color:"#999",fontWeight:600 }}>
                    <span>✅ {item.qty_vendida} vendidos</span><span>↩️ {item.qty_devuelta} devueltos</span><span>📦 {item.qty_enviada} enviados</span>
                  </div>
                </div>
                {disp>0?(
                  <div style={{ display:"grid",gridTemplateColumns:"2fr 1fr",gap:0,borderTop:"1px solid #f5f5f5",marginTop:8 }}>
                    <button disabled={saving} onClick={()=>vender(item,detRec)}
                      style={{ padding:"18px 12px",border:"none",borderRight:"1px solid #f5f5f5",background:isAnim?"#10b981":"#fff",cursor:"pointer",display:"flex",flexDirection:"column",alignItems:"center",gap:5,transition:"background .2s" }}>
                      <span style={{ fontSize:30 }}>{isAnim?"🎉":"💵"}</span>
                      <div style={{ fontWeight:900,fontSize:14,color:isAnim?"#fff":"#10b981" }}>¡Vendí 1!</div>
                      <div style={{ fontSize:10,color:isAnim?"rgba(255,255,255,.8)":"#999" }}>+{fmt(miComis)} para vos</div>
                    </button>
                    <button disabled={saving} onClick={()=>devolver(item,detRec)}
                      style={{ padding:"18px 12px",border:"none",background:"#fff",cursor:"pointer",display:"flex",flexDirection:"column",alignItems:"center",gap:5 }}>
                      <span style={{ fontSize:30 }}>↩️</span>
                      <div style={{ fontWeight:900,fontSize:12,color:"#e63946" }}>Devolver</div>
                      <div style={{ fontSize:10,color:"#999" }}>Al dueño</div>
                    </button>
                  </div>
                ):(
                  <div style={{ padding:"12px 14px",background:item.qty_vendida===item.qty_enviada?"#e8faf4":"#f5f5f5",display:"flex",alignItems:"center",gap:8,borderTop:"1px solid #f0f0f0" }}>
                    <span style={{ fontSize:18 }}>{item.qty_vendida===item.qty_enviada?"🎉":"✓"}</span>
                    <div style={{ fontSize:12,color:item.qty_vendida===item.qty_enviada?"#059669":"#888",fontWeight:700 }}>{item.qty_vendida===item.qty_enviada?"¡Todo vendido!":"Todo devuelto"}</div>
                    {item.qty_vendida>0&&<div style={{ marginLeft:"auto",fontFamily:"var(--mf)",fontWeight:900,fontSize:13,color:"#10b981" }}>{fmt(item.qty_vendida*miComis)}</div>}
                  </div>
                )}
              </div>
            );
          })}
        </div>
      </div>
    );
  }

  // ══════════════════ DETALLE ENVIADA ══════════════════
  if (view==="detalle_env"&&detEnv) {
    const env=detEnvItems.reduce((s,i)=>s+i.qty_enviada,0);
    const vend=detEnvItems.reduce((s,i)=>s+i.qty_vendida,0);
    const dev=detEnvItems.reduce((s,i)=>s+i.qty_devuelta,0);
    const pend=env-vend-dev;
    const deudaC=deudas.filter(d=>d.item?.consignacion_id===detEnv.id&&!d.pagada).reduce((s,d)=>s+d.monto_a_pagar,0);
    return (
      <div style={{ paddingBottom:32 }}>
        <div style={{ background:"linear-gradient(135deg,#ff7a00,#d97706)",padding:"16px 16px 20px" }}>
          <BtnBack/>
          <div style={{ display:"flex",alignItems:"center",gap:12,marginBottom:12 }}>
            <Avatar name={detEnv.vendedora?.name} color={detEnv.vendedora?.color} size={46}/>
            <div>
              <div style={{ fontSize:17,fontWeight:900,color:"#fff" }}>{detEnv.vendedora?.name}</div>
              <div style={{ fontSize:11,color:"rgba(255,255,255,.75)",marginTop:1 }}>{detEnvItems.length} productos · {new Date(detEnv.created_at).toLocaleDateString("es-AR")}</div>
            </div>
          </div>
          {deudaC>0&&(
            <div style={{ background:"rgba(255,255,255,.2)",borderRadius:12,padding:"12px 14px",border:"1px solid rgba(255,255,255,.3)" }}>
              <div style={{ fontSize:10,color:"rgba(255,255,255,.8)",fontWeight:700,textTransform:"uppercase" }}>💰 Pendiente de cobro</div>
              <div style={{ fontFamily:"var(--mf)",fontWeight:900,fontSize:26,color:"#fff",marginTop:2 }}>{fmt(deudaC)}</div>
              <button onClick={()=>{ abrirLiquidar(detEnv.vendedora); }}
                style={{ marginTop:8,background:"#fff",border:"none",borderRadius:10,padding:"8px 16px",color:"#e0224e",fontFamily:"var(--hf)",fontWeight:900,fontSize:12,cursor:"pointer" }}>
                💰 Cobrar ahora
              </button>
            </div>
          )}
        </div>
        <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr 1fr",background:"#fff",borderBottom:"1px solid #eee" }}>
          {[{lbl:"Enviadas",val:env,col:"#555",bg:"#f8f8f8"},{lbl:"Vendidas",val:vend,col:"#059669",bg:"#e8faf4"},{lbl:"Pendientes",val:pend,col:"#d97706",bg:"#fff8e1"}].map((m,i)=>(
            <div key={i} style={{ padding:"12px 8px",background:m.bg,textAlign:"center",borderRight:i<2?"1px solid #eee":"none" }}>
              <div style={{ fontFamily:"var(--mf)",fontWeight:900,fontSize:22,color:m.col }}>{m.val}</div>
              <div style={{ fontSize:9,fontWeight:700,color:m.col,textTransform:"uppercase",letterSpacing:".06em",marginTop:2,opacity:.8 }}>{m.lbl}</div>
            </div>
          ))}
        </div>
        <div style={{ padding:"14px 14px 0" }}>
          {detEnvItems.map(item=>{
            const p=item.product||products.find(x=>x.id===item.product_id);
            const disp=item.qty_enviada-item.qty_vendida-item.qty_devuelta;
            const dItem=deudas.filter(d=>d.item_id===item.id&&!d.pagada).reduce((s,d)=>s+d.monto_a_pagar,0);
            return (
              <div key={item.id} style={{ background:"#fff",borderRadius:16,border:"1.5px solid #e0e0e0",marginBottom:12,overflow:"hidden",boxShadow:"0 2px 10px rgba(0,0,0,.07)" }}>
                <div style={{ display:"flex",alignItems:"center",gap:12,padding:"13px 14px",borderBottom:"1px solid #f5f5f5" }}>
                  <ProdThumb prod={p} size={44}/>
                  <div style={{ flex:1,minWidth:0 }}>
                    <div style={{ fontWeight:800,fontSize:13,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" }}>{p?.name}</div>
                    <div style={{ fontSize:11,color:"#999",marginTop:1 }}>{fmt(p?.price??0)} c/u</div>
                  </div>
                  <div style={{ textAlign:"center",flexShrink:0 }}>
                    <div style={{ fontSize:9,color:"#999",textTransform:"uppercase",fontWeight:700 }}>Enviadas</div>
                    <div style={{ fontFamily:"var(--mf)",fontWeight:900,fontSize:20,color:"#555" }}>{item.qty_enviada}</div>
                  </div>
                </div>
                <div style={{ padding:"8px 14px 6px" }}>
                  <ProgBar vendidas={item.qty_vendida} devueltas={item.qty_devuelta} total={item.qty_enviada}/>
                  <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:4,marginTop:8,textAlign:"center" }}>
                    {[{val:item.qty_vendida,lbl:"Vendidas",col:"#059669",bg:"#e8faf4"},{val:disp,lbl:"Disponibles",col:"#d97706",bg:"#fff8e1"},{val:item.qty_devuelta,lbl:"Devueltas",col:"#888",bg:"#f5f5f5"}].map((m,i)=>(
                      <div key={i} style={{ background:m.bg,borderRadius:8,padding:"7px 4px" }}>
                        <div style={{ fontFamily:"var(--mf)",fontWeight:900,fontSize:16,color:m.col }}>{m.val}</div>
                        <div style={{ fontSize:8,fontWeight:700,color:m.col,textTransform:"uppercase",opacity:.8,marginTop:1 }}>{m.lbl}</div>
                      </div>
                    ))}
                  </div>
                </div>
                {dItem>0&&(
                  <div style={{ padding:"8px 14px 10px",background:"#fff8e1",borderTop:"1px solid #f0e0b0",display:"flex",justifyContent:"space-between",alignItems:"center" }}>
                    <span style={{ fontSize:11,color:"#d97706",fontWeight:700 }}>💰 Pendiente este producto</span>
                    <span style={{ fontFamily:"var(--mf)",fontWeight:900,fontSize:14,color:"#d97706" }}>{fmt(dItem)}</span>
                  </div>
                )}
              </div>
            );
          })}
        </div>
      </div>
    );
  }

  // ══════════════════ LIQUIDAR ══════════════════
  if (view==="liquidar"&&liqTarget) {
    const montoNum=parseFloat(liqMonto)||0;
    const esParcial=montoNum<liqTarget.total-0.01;
    const esValido=montoNum>0&&montoNum<=liqTarget.total+0.01;
    return (
      <div style={{ paddingBottom:32 }}>
        <div style={{ background:"linear-gradient(135deg,#e0224e,#fa1e5a)",padding:"16px 16px 24px" }}>
          <BtnBack/>
          <div style={{ display:"flex",alignItems:"center",gap:12 }}>
            <Avatar name={liqTarget.vendedora?.name} color={liqTarget.vendedora?.color} size={48}/>
            <div>
              <div style={{ fontSize:18,fontWeight:900,color:"#fff" }}>Cobrar a {liqTarget.vendedora?.name}</div>
              <div style={{ fontSize:12,color:"rgba(255,255,255,.8)",marginTop:2 }}>{liqTarget.deudas.length} ventas pendientes</div>
            </div>
          </div>
          <div style={{ background:"rgba(255,255,255,.2)",borderRadius:14,padding:"14px 16px",marginTop:14 }}>
            <div style={{ fontSize:10,color:"rgba(255,255,255,.8)",fontWeight:700,textTransform:"uppercase" }}>Total a cobrar</div>
            <div style={{ fontFamily:"var(--mf)",fontWeight:900,fontSize:32,color:"#fff",marginTop:2 }}>{fmt(liqTarget.total)}</div>
          </div>
        </div>
        <div style={{ padding:"16px 14px 0" }}>
          {/* Desglose */}
          <div style={{ background:"#fff",borderRadius:14,border:"1.5px solid #e0e0e0",marginBottom:16,overflow:"hidden" }}>
            <div style={{ padding:"12px 14px",borderBottom:"1px solid #f5f5f5",fontSize:12,fontWeight:800,color:"#555" }}>Detalle de ventas</div>
            {liqTarget.deudas.slice(0,8).map(d=>{ const pid=d.product?.product_id?.id||d.item?.product_id; const p=(d.product?.product_id)||products.find(x=>x.id===pid); return (
              <div key={d.id} style={{ display:"flex",justifyContent:"space-between",alignItems:"center",padding:"10px 14px",borderBottom:"1px solid #fafafa",fontSize:12 }}>
                <div>
                  <div style={{ fontWeight:600,color:"#222" }}>{p?.name??"Producto"}</div>
                  <div style={{ fontSize:10,color:"#bbb",marginTop:1 }}>{new Date(d.created_at).toLocaleDateString("es-AR")}</div>
                </div>
                <div style={{ textAlign:"right" }}>
                  <div style={{ fontFamily:"var(--mf)",fontWeight:700,color:"#d97706" }}>{fmt(d.monto_a_pagar)}</div>
                  <div style={{ fontSize:9,color:"#bbb" }}>de {fmt(d.monto_total)}</div>
                </div>
              </div>
            );})}
            {liqTarget.deudas.length>8&&<div style={{ padding:"10px 14px",fontSize:11,color:"#bbb",textAlign:"center" }}>+{liqTarget.deudas.length-8} ventas más</div>}
          </div>

          {/* Monto */}
          <div style={{ marginBottom:16 }}>
            <div style={{ fontSize:12,fontWeight:800,color:"#555",marginBottom:8 }}>¿Cuánto te pagó?</div>
            <div style={{ position:"relative" }}>
              <span style={{ position:"absolute",left:14,top:"50%",transform:"translateY(-50%)",fontSize:20,fontWeight:900,color:"#e0224e",fontFamily:"var(--mf)" }}>$</span>
              <input type="number" inputMode="decimal" value={liqMonto} onChange={e=>setLiqMonto(e.target.value)}
                style={{ width:"100%",padding:"16px 14px 16px 36px",borderRadius:14,border:`2px solid ${esValido?"#10b981":"#e0e0e0"}`,fontFamily:"var(--mf)",fontSize:22,fontWeight:900,outline:"none",color:"#222" }}/>
            </div>
            {esParcial&&montoNum>0&&(
              <div style={{ marginTop:8,padding:"8px 12px",background:"#fff8e1",borderRadius:10,fontSize:11,color:"#d97706",fontWeight:700 }}>
                ⚠️ Pago parcial · quedan {fmt(liqTarget.total-montoNum)} pendientes
              </div>
            )}
            <button onClick={()=>setLiqMonto(liqTarget.total.toFixed(2))}
              style={{ marginTop:8,background:"#e8faf4",border:"1.5px solid rgba(0,184,122,.3)",borderRadius:10,padding:"8px 16px",color:"#059669",fontFamily:"var(--hf)",fontWeight:800,fontSize:12,cursor:"pointer" }}>
              ✅ Pago total ({fmt(liqTarget.total)})
            </button>
          </div>

          {/* Modo pago */}
          <div style={{ marginBottom:16 }}>
            <div style={{ fontSize:12,fontWeight:800,color:"#555",marginBottom:10 }}>¿Cómo pagó?</div>
            <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:10 }}>
              {[{id:"efectivo",ico:"💵",lbl:"Efectivo"},{id:"transferencia",ico:"📲",lbl:"Transferencia"}].map(m=>(
                <div key={m.id} onClick={()=>setLiqModo(m.id)}
                  style={{ display:"flex",alignItems:"center",gap:10,padding:"14px",borderRadius:14,border:`2px solid ${liqModo===m.id?"#e0224e":"#e0e0e0"}`,background:liqModo===m.id?"#ffeaea":"#fff",cursor:"pointer",transition:"all .15s" }}>
                  <span style={{ fontSize:22 }}>{m.ico}</span>
                  <span style={{ fontWeight:800,fontSize:13,color:liqModo===m.id?"#e0224e":"#333" }}>{m.lbl}</span>
                  {liqModo===m.id&&<div style={{ marginLeft:"auto",width:20,height:20,borderRadius:"50%",background:"#e0224e",display:"flex",alignItems:"center",justifyContent:"center",color:"#fff",fontSize:12 }}>✓</div>}
                </div>
              ))}
            </div>
          </div>

          {/* Comprobante */}
          {liqModo==="transferencia"&&(
            <div style={{ marginBottom:16 }}>
              <div style={{ fontSize:12,fontWeight:800,color:"#555",marginBottom:8 }}>📸 Comprobante (opcional)</div>
              <div onClick={()=>fileRef.current?.click()}
                style={{ border:"2px dashed #e0e0e0",borderRadius:14,padding:"20px",textAlign:"center",cursor:"pointer",background:"#fafafa" }}>
                {liqFoto?(
                  <div>
                    <img src={liqFoto} alt="" style={{ maxWidth:"100%",borderRadius:10,maxHeight:180,objectFit:"cover" }}/>
                    <div style={{ fontSize:11,color:"#10b981",fontWeight:700,marginTop:8 }}>✅ Comprobante cargado</div>
                  </div>
                ):(
                  <div>
                    <div style={{ fontSize:32,marginBottom:6 }}>📷</div>
                    <div style={{ fontSize:13,fontWeight:700,color:"#666" }}>Tocar para subir comprobante</div>
                    <div style={{ fontSize:11,color:"#bbb",marginTop:2 }}>Foto del comprobante de transferencia</div>
                  </div>
                )}
              </div>
              <input ref={fileRef} type="file" accept="image/*" style={{ display:"none" }} onChange={handleFoto}/>
            </div>
          )}

          {/* Nota */}
          <input style={{ width:"100%",padding:"12px 14px",borderRadius:12,border:"1.5px solid #e0e0e0",fontFamily:"var(--hf)",fontSize:14,outline:"none",color:"#333",marginBottom:20 }}
            placeholder="Nota opcional (ej: pagó el martes en efectivo)" value={liqNota} onChange={e=>setLiqNota(e.target.value)}/>

          <button onClick={confirmarLiquidacion} disabled={saving||!esValido}
            style={{ width:"100%",padding:"18px",borderRadius:16,border:"none",background:esValido?"linear-gradient(135deg,#10b981,#059669)":"#e0e0e0",color:"#fff",fontFamily:"var(--hf)",fontWeight:900,fontSize:16,cursor:!esValido||saving?"not-allowed":"pointer",display:"flex",alignItems:"center",justifyContent:"center",gap:10,boxShadow:esValido?"0 6px 20px rgba(0,184,122,.3)":"none",transition:"all .2s" }}>
            {saving?<Spin/>:<span style={{fontSize:20}}>✅</span>}
            {saving?"Registrando...":(esParcial?`Registrar pago parcial (${fmt(montoNum)})`:"Confirmar cobro completo")}
          </button>
        </div>
      </div>
    );
  }

  return null;
}
