/**
 * ConsignacionModule.jsx — v3
 * Sistema completo de consignación con:
 *  - Agrupación de productos en una consignación
 *  - Comisión configurable (default 30% vendedora / 70% propietario)
 *  - Deudas automáticas al vender
 *  - Pago de deudas con historial
 *  - Devolución de productos no vendidos
 *  - Vista dual: propietario y vendedora
 *
 * Tablas Supabase: consignaciones, consignacion_items, consignacion_deudas
 */
import { useState, useEffect, memo } from "react";

// ── Subcomponentes ────────────────────────────────────────────────────────────
const ProdThumb = memo(({ prod, size = 38 }) => {
  if (!prod) return <div style={{ width:size,height:size,borderRadius:10,background:"var(--bg2)",flexShrink:0 }}/>;
  if (prod.photo_url) return <img src={prod.photo_url} alt="" style={{ width:size,height:size,borderRadius:10,objectFit:"cover",flexShrink:0,border:"1px solid var(--brd)" }}/>;
  return <div style={{ width:size,height:size,borderRadius:10,background:"var(--bg2)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:Math.round(size*.48),flexShrink:0 }}>{prod.emoji||"📦"}</div>;
});

const Spin = () => <div style={{ display:"inline-block",width:15,height:15,border:"2px solid rgba(255,255,255,.3)",borderTopColor:"#fff",borderRadius:"50%",animation:"spin .7s linear infinite" }}/>;

const Badge = ({ children, color = "in" }) => (
  <span style={{ background:`var(--${color}-l)`,color:`var(--${color}-d)`,borderRadius:20,padding:"2px 9px",fontSize:10,fontWeight:800 }}>{children}</span>
);

function fmt(n) { return "$ " + Number(n||0).toLocaleString("es-AR",{minimumFractionDigits:2}); }

// ── Módulo principal ──────────────────────────────────────────────────────────
export default function ConsignacionModule({ sb, me, products, inventory, contacts, onRefresh, toast, fmtARS }) {

  const COMISION_DEFAULT = 30; // % para la vendedora

  // ── Estados ─────────────────────────────────────────────────────────────────
  const [view,        setView]        = useState("main");  // main|nueva|detalle|recibidas
  const [loadingMain, setLoadingMain] = useState(true);
  const [saving,      setSaving]      = useState(false);

  // Datos
  const [consignEnviadas,  setConsignEnviadas]  = useState([]);
  const [consignRecibidas, setConsignRecibidas] = useState([]);
  const [deudas,           setDeudas]           = useState([]);

  // Nueva consignación
  const [vendedoraId, setVendedoraId] = useState("");
  const [comisionPct, setComisionPct] = useState(COMISION_DEFAULT);
  const [notas,       setNotas]       = useState("");
  const [carrito,     setCarrito]     = useState({});   // { product_id: qty }
  const [prodSrch,    setProdSrch]    = useState("");

  // Detalle
  const [detalle,     setDetalle]     = useState(null);  // consignacion row + items
  const [detalleItems,setDetalleItems]= useState([]);

  // ── Carga de datos ───────────────────────────────────────────────────────────
  useEffect(() => { if (me) loadAll(); }, [me]);

  async function loadAll() {
    setLoadingMain(true);
    try {
      // Enviadas por mí
      const { data: env } = await sb.from("consignaciones")
        .select("*, vendedora:vendedora_id(id,name,color), items:consignacion_items(*, product:product_id(id,name,sku,price,emoji,photo_url))")
        .eq("owner_id", me.id)
        .neq("status","cancelada")
        .order("created_at", { ascending: false });
      setConsignEnviadas(env || []);

      // Recibidas por mí
      const { data: rec } = await sb.from("consignaciones")
        .select("*, owner:owner_id(id,name,color), items:consignacion_items(*, product:product_id(id,name,sku,price,emoji,photo_url))")
        .eq("vendedora_id", me.id)
        .neq("status","cancelada")
        .order("created_at", { ascending: false });
      setConsignRecibidas(rec || []);

      // Deudas donde soy propietario (me deben)
      const { data: deu } = await sb.from("consignacion_deudas")
        .select("*, vendedora:vendedora_id(id,name), item:item_id(product_id)")
        .eq("owner_id", me.id)
        .order("created_at", { ascending: false });
      setDeudas(deu || []);
    } catch(e) {
      toast("Error cargando consignaciones", e.message, "e");
    } finally {
      setLoadingMain(false);
    }
  }

  // ── Derivados ────────────────────────────────────────────────────────────────
  const totalDeudaPendiente = deudas.filter(d=>!d.pagada).reduce((s,d)=>s+d.monto_a_pagar,0);
  const cartUnits  = Object.values(carrito).reduce((s,q)=>s+q,0);
  const prodsFilt  = products.filter(p => {
    const q = prodSrch.toLowerCase();
    return !q || p.name.toLowerCase().includes(q) || p.sku.toLowerCase().includes(q);
  });

  // ── FLUJO 1: Crear consignación ──────────────────────────────────────────────
  async function crearConsignacion() {
    if (!vendedoraId) { toast("Seleccioná una vendedora","","e"); return; }
    const items = Object.entries(carrito).filter(([,qty])=>qty>0);
    if (!items.length) { toast("Agregá al menos un producto","","e"); return; }
    setSaving(true);
    try {
      // Verificar stock propio para cada item
      for (const [pid, qty] of items) {
        const inv = inventory.find(i=>i.product_id===pid);
        const prod = products.find(p=>p.id===pid);
        if (!inv || inv.qty_available < qty)
          throw new Error(`Stock insuficiente: "${prod?.name}" — tenés ${inv?.qty_available??0} u.`);
      }

      // Crear la consignación
      const { data: consig, error: cErr } = await sb.from("consignaciones").insert({
        owner_id:     me.id,
        vendedora_id: vendedoraId,
        comision_pct: comisionPct,
        notas:        notas.trim() || null,
        status:       "activa"
      }).select().single();
      if (cErr) throw cErr;

      // Insertar items y ajustar inventario
      for (const [pid, qty] of items) {
        const prod = products.find(p=>p.id===pid);
        const inv  = inventory.find(i=>i.product_id===pid);

        // Crear item
        const { error: iErr } = await sb.from("consignacion_items").insert({
          consignacion_id: consig.id,
          product_id:      pid,
          qty_enviada:     qty,
          qty_vendida:     0,
          qty_devuelta:    0,
          precio_venta:    prod?.price ?? 0,
        });
        if (iErr) throw iErr;

        // Restar de MI inventario
        const { error: invErr } = await sb.from("inventory")
          .update({ qty_available: inv.qty_available - qty })
          .eq("id", inv.id);
        if (invErr) throw invErr;

        // Sumar al inventario de la vendedora
        const { data: vInv } = await sb.from("inventory").select("*")
          .eq("user_id", vendedoraId).eq("product_id", pid).maybeSingle();
        if (vInv) {
          await sb.from("inventory").update({ qty_available: vInv.qty_available + qty }).eq("id", vInv.id);
        } else {
          await sb.from("inventory").insert({ user_id:vendedoraId, product_id:pid, qty_available:qty, qty_sold:0, source:"consigna", supplier_id:me.id });
        }
      }

      // Notificar a la vendedora
      const vName = contacts.find(c=>c.id===vendedoraId)?.name ?? "Vendedora";
      await sb.from("notifications").insert({
        to_user_id: vendedoraId, from_name: me.name, type:"transfer",
        message: `📦 ${me.name} te envió una consignación con ${items.length} producto${items.length!==1?"s":""} (${cartUnits} u. totales)`
      });

      toast("✅ Consignación creada", `${cartUnits} u. enviadas a ${vName}`, "s");
      setCarrito({}); setVendedoraId(""); setNotas(""); setView("main");
      loadAll(); onRefresh();
    } catch(err) {
      toast("Error al crear consignación", err.message, "e");
    } finally {
      setSaving(false);
    }
  }

  // ── FLUJO 2: Registrar venta (vendedora vende 1 unidad) ──────────────────────
  async function registrarVenta(item, consig) {
    const qtyDisp = item.qty_enviada - item.qty_vendida - item.qty_devuelta;
    if (qtyDisp < 1) { toast("Sin unidades disponibles","","e"); return; }
    setSaving(true);
    try {
      const prod = item.product || products.find(p=>p.id===item.product_id);
      const precio    = item.precio_venta;
      const comision  = Math.round(precio * consig.comision_pct / 100 * 100) / 100;
      const aPagar    = precio - comision;

      // 1. Actualizar item
      const { error: iErr } = await sb.from("consignacion_items")
        .update({ qty_vendida: item.qty_vendida + 1 })
        .eq("id", item.id);
      if (iErr) throw iErr;

      // 2. Restar inventario de la vendedora
      const { data: vInv } = await sb.from("inventory").select("*")
        .eq("user_id", consig.vendedora_id).eq("product_id", item.product_id).maybeSingle();
      if (vInv) {
        await sb.from("inventory").update({ qty_available: Math.max(0,vInv.qty_available-1), qty_sold:(vInv.qty_sold||0)+1 }).eq("id",vInv.id);
      }

      // 3. Crear deuda
      const { error: dErr } = await sb.from("consignacion_deudas").insert({
        consignacion_id: consig.id,
        item_id:         item.id,
        owner_id:        consig.owner_id,
        vendedora_id:    consig.vendedora_id,
        qty:             1,
        monto_total:     precio,
        comision:        comision,
        monto_a_pagar:   aPagar,
        pagada:          false,
      });
      if (dErr) throw dErr;

      // 4. sale_log
      await sb.from("sale_logs").insert({ user_id:consig.vendedora_id, product_id:item.product_id, qty:1, sale_price:precio, source:"consignment" });

      // 5. Notificar al propietario
      await sb.from("notifications").insert({
        to_user_id: consig.owner_id, from_name: me.name, type:"sale",
        message: `💰 ${me.name} vendió 1× "${prod?.name}". Te debe rendir ${fmtARS(aPagar)} (comisión 30%: ${fmtARS(comision)})`
      });

      toast("💰 Venta registrada", `${prod?.name} — Propietario: ${fmtARS(aPagar)}`, "s");
      loadAll(); onRefresh();
    } catch(err) {
      toast("Error al registrar venta", err.message, "e");
    } finally {
      setSaving(false);
    }
  }

  // ── FLUJO 3: Devolver producto no vendido ─────────────────────────────────────
  async function registrarDevolucion(item, consig) {
    const qtyDisp = item.qty_enviada - item.qty_vendida - item.qty_devuelta;
    if (qtyDisp < 1) { toast("Sin unidades para devolver","","e"); return; }
    setSaving(true);
    try {
      const prod = item.product || products.find(p=>p.id===item.product_id);

      // 1. Actualizar item
      await sb.from("consignacion_items").update({ qty_devuelta: item.qty_devuelta + 1 }).eq("id", item.id);

      // 2. Restar inventario vendedora
      const { data: vInv } = await sb.from("inventory").select("*")
        .eq("user_id", consig.vendedora_id).eq("product_id", item.product_id).maybeSingle();
      if (vInv) await sb.from("inventory").update({ qty_available: Math.max(0,vInv.qty_available-1) }).eq("id",vInv.id);

      // 3. Sumar a inventario propietario
      const { data: oInv } = await sb.from("inventory").select("*")
        .eq("user_id", consig.owner_id).eq("product_id", item.product_id).maybeSingle();
      if (oInv) {
        await sb.from("inventory").update({ qty_available: oInv.qty_available+1 }).eq("id",oInv.id);
      } else {
        await sb.from("inventory").insert({ user_id:consig.owner_id, product_id:item.product_id, qty_available:1, qty_sold:0, source:"own" });
      }

      // 4. Notificar propietario
      await sb.from("notifications").insert({
        to_user_id: consig.owner_id, from_name: me.name, type:"confirm",
        message: `↩️ ${me.name} devolvió 1× "${prod?.name}". Stock restituido.`
      });

      toast("↩️ Devolución registrada", `+1 "${prod?.name}" en stock del propietario`, "s");
      loadAll(); onRefresh();
    } catch(err) {
      toast("Error al registrar devolución", err.message, "e");
    } finally {
      setSaving(false);
    }
  }

  // ── FLUJO 4: Marcar deuda como pagada ────────────────────────────────────────
  async function marcarPagada(deuda) {
    setSaving(true);
    try {
      await sb.from("consignacion_deudas")
        .update({ pagada: true, pagada_at: new Date().toISOString() })
        .eq("id", deuda.id);
      const vName = deuda.vendedora?.name ?? "Vendedora";
      toast("✅ Deuda liquidada", `${fmtARS(deuda.monto_a_pagar)} de ${vName}`, "s");
      loadAll();
    } catch(err) {
      toast("Error", err.message, "e");
    } finally {
      setSaving(false);
    }
  }

  // ── Abrir detalle ─────────────────────────────────────────────────────────────
  async function verDetalle(consig) {
    const { data: items } = await sb.from("consignacion_items")
      .select("*, product:product_id(id,name,sku,price,emoji,photo_url)")
      .eq("consignacion_id", consig.id);
    setDetalle(consig);
    setDetalleItems(items || []);
    setView("detalle");
  }

  // ─────────────────────────────────────────────────────────────────────────────
  // RENDER
  // ─────────────────────────────────────────────────────────────────────────────

  if (loadingMain) return (
    <div style={{ display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:"60px 20px",gap:12 }}>
      <Spin/>
      <div style={{ fontSize:13,color:"var(--t3)" }}>Cargando consignaciones...</div>
    </div>
  );

  // ── MAIN ─────────────────────────────────────────────────────────────────────
  if (view === "main") {
    const deudaPend = deudas.filter(d=>!d.pagada);
    const deudaTotal = deudaPend.reduce((s,d)=>s+d.monto_a_pagar,0);

    return (
      <div style={{ paddingBottom:32 }}>
        {/* Header */}
        <div style={{ background:"linear-gradient(135deg,#6d28d9,#a855f7)",padding:"20px 16px 24px" }}>
          <div style={{ fontSize:21,fontWeight:900,color:"#fff",marginBottom:4 }}>🤝 Consignaciones</div>
          <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginTop:14 }}>
            {[
              { ico:"📤",val:consignEnviadas.filter(c=>c.status==="activa").length,lbl:"Activas enviadas" },
              { ico:"📥",val:consignRecibidas.filter(c=>c.status==="activa").length,lbl:"Recibidas activas" },
              { ico:"💰",val:fmtARS(deudaTotal),lbl:"Me deben",sm:true },
              { ico:"✅",val:deudas.filter(d=>d.pagada).length,lbl:"Deudas cobradas" },
            ].map((m,i)=>(
              <div key={i} style={{ background:"rgba(255,255,255,.15)",borderRadius:16,padding:"13px 12px",border:"1px solid rgba(255,255,255,.2)" }}>
                <div style={{ fontSize:18,marginBottom:3 }}>{m.ico}</div>
                <div style={{ fontFamily:"var(--mf)",fontWeight:900,fontSize:m.sm?12:20,color:"#fff",lineHeight:1 }}>{m.val}</div>
                <div style={{ fontSize:10,color:"rgba(255,255,255,.7)",marginTop:3,fontWeight:600 }}>{m.lbl}</div>
              </div>
            ))}
          </div>
        </div>

        <div style={{ padding:"14px 14px 0" }}>
          {/* Acciones principales */}
          <div style={{ display:"flex",flexDirection:"column",gap:9,marginBottom:16 }}>
            <button onClick={()=>{setView("nueva");setCarrito({});setVendedoraId("");setNotas("");}}
              style={{ display:"flex",alignItems:"center",gap:14,padding:"14px 16px",borderRadius:16,background:"linear-gradient(135deg,#ff7a00,#e06a00)",border:"none",cursor:"pointer",boxShadow:"0 4px 16px rgba(255,122,0,.3)" }}>
              <div style={{ width:42,height:42,borderRadius:13,background:"rgba(255,255,255,.2)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:20,flexShrink:0 }}>📤</div>
              <div style={{ textAlign:"left" }}>
                <div style={{ fontSize:14,fontWeight:900,color:"#fff" }}>Nueva entrega en consignación</div>
                <div style={{ fontSize:11,color:"rgba(255,255,255,.8)",marginTop:1 }}>Agrupa productos y asignálos a una vendedora</div>
              </div>
            </button>
            {consignRecibidas.filter(c=>c.status==="activa").length>0&&(
              <button onClick={()=>setView("recibidas")}
                style={{ display:"flex",alignItems:"center",gap:14,padding:"14px 16px",borderRadius:16,background:"linear-gradient(135deg,#0096c7,#0077a3)",border:"none",cursor:"pointer",boxShadow:"0 4px 16px rgba(0,150,199,.25)" }}>
                <div style={{ width:42,height:42,borderRadius:13,background:"rgba(255,255,255,.2)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:20,flexShrink:0 }}>📥</div>
                <div style={{ textAlign:"left" }}>
                  <div style={{ fontSize:14,fontWeight:900,color:"#fff" }}>Mis consignaciones recibidas</div>
                  <div style={{ fontSize:11,color:"rgba(255,255,255,.8)",marginTop:1 }}>{consignRecibidas.filter(c=>c.status==="activa").length} activas — registrar ventas y devoluciones</div>
                </div>
              </button>
            )}
          </div>

          {/* Deudas pendientes */}
          {deudaPend.length>0&&(
            <div style={{ marginBottom:16 }}>
              <div style={{ fontSize:12,fontWeight:800,color:"var(--t2)",marginBottom:8,textTransform:"uppercase",letterSpacing:".07em" }}>💰 Deudas pendientes de cobro</div>
              {deudaPend.map(d=>{
                const prod = products.find(p=>p.id===d.item?.product_id);
                return (
                  <div key={d.id} style={{ background:"#fff",borderRadius:16,border:"1.5px solid var(--am)",marginBottom:8,overflow:"hidden",boxShadow:"var(--sh)" }}>
                    <div style={{ display:"flex",alignItems:"center",gap:12,padding:"13px 14px" }}>
                      <div style={{ width:40,height:40,borderRadius:12,background:"var(--am-l)",display:"flex",alignItems:"center",justifyContent:"center",fontWeight:900,fontSize:16,color:"var(--am-d)",flexShrink:0 }}>
                        {d.vendedora?.name?.charAt(0).toUpperCase()??"?"}
                      </div>
                      <div style={{ flex:1,minWidth:0 }}>
                        <div style={{ fontWeight:700,fontSize:13 }}>{d.vendedora?.name??"-"}</div>
                        <div style={{ fontSize:11,color:"var(--t3)",marginTop:1 }}>
                          {prod?.name??"-"} · vendió {d.qty} u. · {new Date(d.created_at).toLocaleDateString("es-AR")}
                        </div>
                        <div style={{ fontSize:10,color:"var(--t3)",marginTop:1 }}>
                          Total: {fmtARS(d.monto_total)} · Comisión 30%: {fmtARS(d.comision)} · <strong style={{color:"var(--am-d)"}}>A cobrar: {fmtARS(d.monto_a_pagar)}</strong>
                        </div>
                      </div>
                      <div style={{ textAlign:"right",flexShrink:0 }}>
                        <div style={{ fontFamily:"var(--mf)",fontWeight:900,fontSize:15,color:"var(--am-d)" }}>{fmtARS(d.monto_a_pagar)}</div>
                        <button disabled={saving} onClick={()=>marcarPagada(d)}
                          style={{ marginTop:5,background:"var(--em)",border:"none",borderRadius:9,padding:"5px 10px",cursor:"pointer",color:"#fff",fontFamily:"var(--hf)",fontWeight:800,fontSize:11 }}>
                          ✓ Cobrado
                        </button>
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>
          )}

          {/* Historial cobradas */}
          {deudas.filter(d=>d.pagada).length>0&&(
            <div style={{ marginBottom:16 }}>
              <div style={{ fontSize:12,fontWeight:800,color:"var(--t3)",marginBottom:8,textTransform:"uppercase",letterSpacing:".07em" }}>✅ Historial de pagos</div>
              {deudas.filter(d=>d.pagada).slice(0,5).map(d=>{
                const prod = products.find(p=>p.id===d.item?.product_id);
                return (
                  <div key={d.id} style={{ display:"flex",alignItems:"center",gap:10,padding:"10px 12px",borderRadius:12,background:"var(--em-l)",border:"1px solid rgba(0,184,122,.2)",marginBottom:6 }}>
                    <div style={{ fontSize:18 }}>✅</div>
                    <div style={{ flex:1 }}>
                      <div style={{ fontWeight:600,fontSize:12 }}>{d.vendedora?.name} · {prod?.name??"-"}</div>
                      <div style={{ fontSize:10,color:"var(--t3)" }}>Cobrado {new Date(d.pagada_at).toLocaleDateString("es-AR")}</div>
                    </div>
                    <div style={{ fontFamily:"var(--mf)",fontWeight:700,fontSize:13,color:"var(--em-d)" }}>{fmtARS(d.monto_a_pagar)}</div>
                  </div>
                );
              })}
            </div>
          )}

          {/* Consignaciones activas enviadas */}
          {consignEnviadas.filter(c=>c.status==="activa").length>0&&(
            <div>
              <div style={{ fontSize:12,fontWeight:800,color:"var(--t2)",marginBottom:8,textTransform:"uppercase",letterSpacing:".07em" }}>📤 Enviadas activas</div>
              {consignEnviadas.filter(c=>c.status==="activa").map(c=>{
                const totalItems = (c.items||[]).reduce((s,i)=>s+i.qty_enviada,0);
                const vendidas   = (c.items||[]).reduce((s,i)=>s+i.qty_vendida,0);
                const devueltas  = (c.items||[]).reduce((s,i)=>s+i.qty_devuelta,0);
                const pendientes = totalItems - vendidas - devueltas;
                return (
                  <div key={c.id} style={{ background:"#fff",borderRadius:16,border:"1.5px solid var(--brd)",marginBottom:10,overflow:"hidden",boxShadow:"var(--sh)" }}>
                    <div style={{ display:"flex",alignItems:"center",gap:12,padding:"13px 14px",borderBottom:"1px solid var(--brd)" }}>
                      <div style={{ width:42,height:42,borderRadius:13,background:c.vendedora?.color||"var(--in)",display:"flex",alignItems:"center",justifyContent:"center",fontWeight:900,fontSize:17,color:"#fff",flexShrink:0 }}>
                        {c.vendedora?.name?.charAt(0).toUpperCase()??"?"}
                      </div>
                      <div style={{ flex:1 }}>
                        <div style={{ fontWeight:800,fontSize:14 }}>{c.vendedora?.name??"-"}</div>
                        <div style={{ fontSize:11,color:"var(--t3)",marginTop:1 }}>{(c.items||[]).length} productos · {new Date(c.created_at).toLocaleDateString("es-AR")}</div>
                      </div>
                      <button onClick={()=>verDetalle(c)}
                        style={{ background:"var(--in-l)",border:"1.5px solid rgba(124,58,237,.2)",borderRadius:10,padding:"7px 12px",cursor:"pointer",fontFamily:"var(--hf)",fontWeight:800,fontSize:11,color:"var(--in-d)" }}>
                        Ver detalle
                      </button>
                    </div>
                    <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr 1fr",textAlign:"center" }}>
                      {[{lbl:"Enviadas",val:totalItems,col:"var(--t2)",bg:"var(--bg)"},{lbl:"Vendidas",val:vendidas,col:"var(--em-d)",bg:"var(--em-l)"},{lbl:"Pendientes",val:pendientes,col:"var(--am-d)",bg:"var(--am-l)"}].map((m,i)=>(
                        <div key={i} style={{ padding:"10px 8px",background:m.bg,borderRight:i<2?"1px solid var(--brd)":"none" }}>
                          <div style={{ fontFamily:"var(--mf)",fontWeight:900,fontSize:20,color:m.col }}>{m.val}</div>
                          <div style={{ fontSize:9,fontWeight:700,color:m.col,opacity:.75,textTransform:"uppercase",letterSpacing:".06em" }}>{m.lbl}</div>
                        </div>
                      ))}
                    </div>
                    {c.notas&&<div style={{ padding:"8px 14px",fontSize:11,color:"var(--t3)",background:"var(--bg)",borderTop:"1px solid var(--brd)" }}>📝 {c.notas}</div>}
                  </div>
                );
              })}
            </div>
          )}

          {consignEnviadas.length===0&&consignRecibidas.length===0&&(
            <div style={{ textAlign:"center",padding:"40px 20px",color:"var(--t3)",fontSize:13 }}>
              <div style={{ fontSize:42,marginBottom:8 }}>📭</div>
              Sin consignaciones aún.<br/>Usá el botón de arriba para crear la primera.
            </div>
          )}
        </div>
      </div>
    );
  }

  // ── NUEVA CONSIGNACIÓN ───────────────────────────────────────────────────────
  if (view === "nueva") {
    const cartTotal = Object.entries(carrito).reduce((s,[pid,qty])=>{
      const p=products.find(x=>x.id===pid); return s+(p?p.price*qty:0);
    },0);
    const aPagarPropietario = cartTotal * (1 - comisionPct/100);

    return (
      <div style={{ paddingBottom:32 }}>
        <div style={{ background:"linear-gradient(135deg,#6d28d9,#a855f7)",padding:"16px 16px 20px" }}>
          <button onClick={()=>setView("main")} style={{ background:"rgba(255,255,255,.2)",border:"none",borderRadius:10,padding:"6px 12px",color:"#fff",fontFamily:"var(--hf)",fontWeight:700,fontSize:12,cursor:"pointer",marginBottom:10 }}>← Volver</button>
          <div style={{ fontSize:18,fontWeight:900,color:"#fff" }}>📤 Nueva Consignación</div>
          <div style={{ fontSize:11,color:"rgba(255,255,255,.7)",marginTop:2 }}>Agrupá productos y asignálos a una vendedora</div>
        </div>
        <div style={{ padding:"16px 14px 0" }}>

          {/* Paso 1: Vendedora */}
          <div style={{ fontSize:12,fontWeight:800,color:"var(--t2)",marginBottom:10,textTransform:"uppercase",letterSpacing:".07em" }}>1. Vendedora</div>
          {contacts.map(c=>(
            <div key={c.id} onClick={()=>setVendedoraId(c.id)}
              style={{ display:"flex",alignItems:"center",gap:12,padding:"12px 14px",borderRadius:14,border:`2px solid ${vendedoraId===c.id?"var(--in)":"var(--brd)"}`,background:vendedoraId===c.id?"var(--in-l)":"#fff",marginBottom:8,cursor:"pointer",boxShadow:"var(--sh)" }}>
              <div style={{ width:38,height:38,borderRadius:11,background:c.color||"var(--in)",display:"flex",alignItems:"center",justifyContent:"center",fontWeight:900,fontSize:15,color:"#fff",flexShrink:0 }}>
                {c.name.charAt(0).toUpperCase()}
              </div>
              <div style={{ flex:1 }}><div style={{ fontWeight:700,fontSize:13 }}>{c.name}</div><div style={{ fontSize:11,color:"var(--t3)" }}>{c.email}</div></div>
              {vendedoraId===c.id&&<span style={{ fontSize:18,color:"var(--in-d)" }}>✓</span>}
            </div>
          ))}
          {contacts.length===0&&<div style={{ color:"var(--t3)",fontSize:13,padding:"12px 0" }}>Sin contactos. Agregá uno en la tab Red.</div>}

          {/* Paso 2: Comisión */}
          <div style={{ marginTop:16,marginBottom:16 }}>
            <div style={{ fontSize:12,fontWeight:800,color:"var(--t2)",marginBottom:10,textTransform:"uppercase",letterSpacing:".07em" }}>2. Comisión de la vendedora</div>
            <div style={{ background:"var(--am-l)",border:"1.5px solid rgba(255,122,0,.2)",borderRadius:14,padding:"14px" }}>
              <div style={{ display:"flex",alignItems:"center",gap:12,marginBottom:10 }}>
                <input type="range" min="0" max="50" step="5" value={comisionPct} onChange={e=>setComisionPct(Number(e.target.value))} style={{ flex:1,accentColor:"var(--am)" }}/>
                <div style={{ fontFamily:"var(--mf)",fontWeight:900,fontSize:22,color:"var(--am-d)",minWidth:48,textAlign:"right" }}>{comisionPct}%</div>
              </div>
              {cartUnits>0&&(
                <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:8 }}>
                  <div style={{ background:"#fff",borderRadius:10,padding:"10px",textAlign:"center" }}>
                    <div style={{ fontSize:10,color:"var(--t3)",fontWeight:700 }}>Vendedora cobra</div>
                    <div style={{ fontFamily:"var(--mf)",fontWeight:900,fontSize:14,color:"var(--am-d)",marginTop:3 }}>{fmtARS(cartTotal*comisionPct/100)}</div>
                  </div>
                  <div style={{ background:"#fff",borderRadius:10,padding:"10px",textAlign:"center" }}>
                    <div style={{ fontSize:10,color:"var(--t3)",fontWeight:700 }}>Vos cobrás</div>
                    <div style={{ fontFamily:"var(--mf)",fontWeight:900,fontSize:14,color:"var(--em-d)",marginTop:3 }}>{fmtARS(aPagarPropietario)}</div>
                  </div>
                </div>
              )}
            </div>
          </div>

          {/* Paso 3: Productos */}
          <div style={{ fontSize:12,fontWeight:800,color:"var(--t2)",marginBottom:10,textTransform:"uppercase",letterSpacing:".07em" }}>3. Productos</div>
          <div style={{ position:"relative",marginBottom:10 }}>
            <span style={{ position:"absolute",left:12,top:"50%",transform:"translateY(-50%)",color:"var(--t3)",fontSize:16 }}>🔍</span>
            <input style={{ width:"100%",padding:"11px 12px 11px 38px",borderRadius:12,border:"1.5px solid var(--brd)",background:"#fff",fontFamily:"var(--hf)",fontSize:14,outline:"none",color:"var(--t1)" }}
              placeholder="Buscar producto..." value={prodSrch} onChange={e=>setProdSrch(e.target.value)}/>
          </div>
          {prodsFilt.slice(0,15).map(p=>{
            const inv   = inventory.find(i=>i.product_id===p.id);
            const avail = inv?.qty_available??0;
            const inC   = carrito[p.id]??0;
            return (
              <div key={p.id} style={{ display:"flex",alignItems:"center",gap:12,padding:"11px 13px",borderRadius:14,border:`1.5px solid ${inC>0?"var(--in)":"var(--brd)"}`,background:inC>0?"var(--in-l)":"#fff",marginBottom:8,boxShadow:"var(--sh)" }}>
                <ProdThumb prod={p} size={38}/>
                <div style={{ flex:1,minWidth:0 }}>
                  <div style={{ fontWeight:700,fontSize:12,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" }}>{p.name}</div>
                  <div style={{ fontSize:10,color:"var(--t3)",marginTop:1 }}>{p.sku} · {fmtARS(p.price)} · {avail} disp.</div>
                </div>
                <div style={{ display:"flex",alignItems:"center",border:"1.5px solid var(--brd)",borderRadius:9,overflow:"hidden",flexShrink:0 }}>
                  <button onClick={()=>setCarrito(c=>{const n={...c};if((n[p.id]??0)<=1){delete n[p.id];}else{n[p.id]--;}return n;})}
                    style={{ width:28,height:28,border:"none",background:"var(--bg)",cursor:"pointer",fontSize:16,display:"flex",alignItems:"center",justifyContent:"center" }}>−</button>
                  <div style={{ minWidth:26,textAlign:"center",fontFamily:"var(--mf)",fontWeight:700,fontSize:13,color:"var(--in)" }}>{inC}</div>
                  <button onClick={()=>{if(inC>=avail)return;setCarrito(c=>({...c,[p.id]:(c[p.id]??0)+1}));}}
                    disabled={inC>=avail}
                    style={{ width:28,height:28,border:"none",background:inC>=avail?"var(--bg2)":"var(--in-l)",cursor:inC>=avail?"not-allowed":"pointer",fontSize:16,display:"flex",alignItems:"center",justifyContent:"center",opacity:inC>=avail?.4:1 }}>+</button>
                </div>
              </div>
            );
          })}

          {/* Paso 4: Notas */}
          <div style={{ marginTop:14,marginBottom:14 }}>
            <div style={{ fontSize:12,fontWeight:800,color:"var(--t2)",marginBottom:8,textTransform:"uppercase",letterSpacing:".07em" }}>4. Notas (opcional)</div>
            <textarea style={{ width:"100%",padding:"11px 14px",borderRadius:12,border:"1.5px solid var(--brd)",fontFamily:"var(--hf)",fontSize:14,outline:"none",resize:"vertical",minHeight:72,color:"var(--t1)" }}
              placeholder="Ej: Productos para la feria del 15/06..." value={notas} onChange={e=>setNotas(e.target.value)}/>
          </div>

          {/* Resumen carrito */}
          {cartUnits>0&&(
            <div style={{ background:"var(--em-l)",border:"1.5px solid rgba(0,184,122,.2)",borderRadius:14,padding:"14px",marginBottom:14 }}>
              <div style={{ fontSize:11,fontWeight:800,color:"var(--em-d)",textTransform:"uppercase",letterSpacing:".06em",marginBottom:8 }}>Resumen de la consignación</div>
              {Object.entries(carrito).filter(([,qty])=>qty>0).map(([pid,qty])=>{
                const p=products.find(x=>x.id===pid);
                return <div key={pid} style={{ display:"flex",justifyContent:"space-between",fontSize:12,color:"var(--t2)",marginBottom:4 }}>
                  <span>{p?.emoji} {p?.name}</span><span style={{ fontWeight:700 }}>{qty} u. · {fmtARS((p?.price??0)*qty)}</span>
                </div>;
              })}
              <div style={{ borderTop:"1px solid rgba(0,184,122,.2)",marginTop:10,paddingTop:10,display:"grid",gridTemplateColumns:"1fr 1fr" }}>
                <div>
                  <div style={{ fontSize:10,color:"var(--t3)",fontWeight:700 }}>Total en mercadería</div>
                  <div style={{ fontFamily:"var(--mf)",fontWeight:900,fontSize:15,color:"var(--em-d)" }}>{fmtARS(cartTotal)}</div>
                </div>
                <div style={{ textAlign:"right" }}>
                  <div style={{ fontSize:10,color:"var(--t3)",fontWeight:700 }}>Vos cobrás ({100-comisionPct}%)</div>
                  <div style={{ fontFamily:"var(--mf)",fontWeight:900,fontSize:15,color:"var(--em-d)" }}>{fmtARS(aPagarPropietario)}</div>
                </div>
              </div>
            </div>
          )}

          <button onClick={crearConsignacion} disabled={saving||!vendedoraId||cartUnits===0}
            style={{ width:"100%",padding:"16px",borderRadius:16,border:"none",background:"linear-gradient(135deg,#00b87a,#009a66)",color:"#fff",fontFamily:"var(--hf)",fontWeight:900,fontSize:15,cursor:saving||!vendedoraId||cartUnits===0?"not-allowed":"pointer",opacity:!vendedoraId||cartUnits===0?.45:1,display:"flex",alignItems:"center",justifyContent:"center",gap:10,boxShadow:"0 6px 20px rgba(0,184,122,.3)" }}>
            {saving?<Spin/>:"✅"} {saving?"Creando consignación...":`Crear consignación (${cartUnits} u.)`}
          </button>
        </div>
      </div>
    );
  }

  // ── DETALLE DE CONSIGNACIÓN (propietario) ────────────────────────────────────
  if (view === "detalle" && detalle) {
    const vendida   = detalleItems.reduce((s,i)=>s+i.qty_vendida,0);
    const devuelta  = detalleItems.reduce((s,i)=>s+i.qty_devuelta,0);
    const enviada   = detalleItems.reduce((s,i)=>s+i.qty_enviada,0);
    const pendiente = enviada - vendida - devuelta;
    const deudaConsig = deudas.filter(d=>d.consignacion_id===detalle.id&&!d.pagada).reduce((s,d)=>s+d.monto_a_pagar,0);

    return (
      <div style={{ paddingBottom:32 }}>
        <div style={{ background:"linear-gradient(135deg,#ff7a00,#e06a00)",padding:"16px 16px 20px" }}>
          <button onClick={()=>setView("main")} style={{ background:"rgba(255,255,255,.2)",border:"none",borderRadius:10,padding:"6px 12px",color:"#fff",fontFamily:"var(--hf)",fontWeight:700,fontSize:12,cursor:"pointer",marginBottom:10 }}>← Volver</button>
          <div style={{ display:"flex",alignItems:"center",gap:12 }}>
            <div style={{ width:44,height:44,borderRadius:13,background:detalle.vendedora?.color||"rgba(255,255,255,.25)",display:"flex",alignItems:"center",justifyContent:"center",fontWeight:900,fontSize:18,color:"#fff",flexShrink:0 }}>
              {detalle.vendedora?.name?.charAt(0).toUpperCase()??"?"}
            </div>
            <div>
              <div style={{ fontSize:17,fontWeight:900,color:"#fff" }}>{detalle.vendedora?.name}</div>
              <div style={{ fontSize:11,color:"rgba(255,255,255,.75)",marginTop:1 }}>{detalleItems.length} productos · {new Date(detalle.created_at).toLocaleDateString("es-AR")}</div>
            </div>
          </div>
          {deudaConsig>0&&(
            <div style={{ background:"rgba(255,255,255,.2)",borderRadius:13,padding:"11px 13px",marginTop:12,border:"1px solid rgba(255,255,255,.3)" }}>
              <div style={{ fontSize:10,color:"rgba(255,255,255,.8)",fontWeight:700 }}>💰 PENDIENTE DE COBRO</div>
              <div style={{ fontFamily:"var(--mf)",fontWeight:900,fontSize:24,color:"#fff" }}>{fmtARS(deudaConsig)}</div>
            </div>
          )}
        </div>
        {/* Stats */}
        <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr 1fr",borderBottom:"1px solid var(--brd)",background:"#fff" }}>
          {[{lbl:"Enviadas",val:enviada,col:"var(--t2)",bg:"var(--bg)"},{lbl:"Vendidas",val:vendida,col:"var(--em-d)",bg:"var(--em-l)"},{lbl:"Pendientes",val:pendiente,col:"var(--am-d)",bg:"var(--am-l)"}].map((m,i)=>(
            <div key={i} style={{ padding:"12px 8px",background:m.bg,textAlign:"center",borderRight:i<2?"1px solid var(--brd)":"none" }}>
              <div style={{ fontFamily:"var(--mf)",fontWeight:900,fontSize:22,color:m.col }}>{m.val}</div>
              <div style={{ fontSize:9,fontWeight:800,color:m.col,opacity:.75,textTransform:"uppercase",letterSpacing:".06em",marginTop:2 }}>{m.lbl}</div>
            </div>
          ))}
        </div>

        <div style={{ padding:"14px 14px 0" }}>
          {detalleItems.map(item=>{
            const p      = item.product || products.find(x=>x.id===item.product_id);
            const qDisp  = item.qty_enviada - item.qty_vendida - item.qty_devuelta;
            const deudaI = deudas.filter(d=>d.item_id===item.id&&!d.pagada).reduce((s,d)=>s+d.monto_a_pagar,0);
            return (
              <div key={item.id} style={{ background:"#fff",borderRadius:16,border:"1.5px solid var(--brd)",marginBottom:12,overflow:"hidden",boxShadow:"var(--sh)" }}>
                <div style={{ display:"flex",alignItems:"center",gap:12,padding:"13px 14px",borderBottom:"1px solid var(--brd)" }}>
                  <ProdThumb prod={p} size={44}/>
                  <div style={{ flex:1,minWidth:0 }}>
                    <div style={{ fontWeight:800,fontSize:13,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" }}>{p?.name??"-"}</div>
                    <div style={{ fontSize:11,color:"var(--t3)",marginTop:1 }}>{p?.sku} · {fmtARS(p?.price??0)} c/u</div>
                  </div>
                  <div style={{ textAlign:"right",flexShrink:0 }}>
                    <div style={{ fontSize:9,color:"var(--t3)",textTransform:"uppercase" }}>Enviadas</div>
                    <div style={{ fontFamily:"var(--mf)",fontWeight:900,fontSize:20,color:"var(--t2)" }}>{item.qty_enviada}</div>
                  </div>
                </div>
                <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr 1fr",borderBottom:"1px solid var(--brd)" }}>
                  {[{lbl:"Vendidas",val:item.qty_vendida,col:"var(--em-d)",bg:"var(--em-l)"},{lbl:"Disponibles",val:qDisp,col:"var(--am-d)",bg:"var(--am-l)"},{lbl:"Devueltas",val:item.qty_devuelta,col:"var(--in-d)",bg:"var(--in-l)"}].map((m,i)=>(
                    <div key={i} style={{ padding:"9px 6px",background:m.bg,textAlign:"center",borderRight:i<2?"1px solid var(--brd)":"none" }}>
                      <div style={{ fontFamily:"var(--mf)",fontWeight:900,fontSize:18,color:m.col }}>{m.val}</div>
                      <div style={{ fontSize:8,fontWeight:800,color:m.col,opacity:.75,textTransform:"uppercase",letterSpacing:".05em",marginTop:2 }}>{m.lbl}</div>
                    </div>
                  ))}
                </div>
                {item.qty_vendida>0&&(
                  <div style={{ padding:"10px 13px",background:"var(--am-l)",borderBottom:"1px solid var(--brd)",display:"flex",justifyContent:"space-between",alignItems:"center" }}>
                    <span style={{ fontSize:12,color:"var(--am-d)",fontWeight:700 }}>💰 Pendiente de cobro por ventas</span>
                    <span style={{ fontFamily:"var(--mf)",fontWeight:900,fontSize:14,color:"var(--am-d)" }}>{fmtARS(deudaI)}</span>
                  </div>
                )}
              </div>
            );
          })}
        </div>
      </div>
    );
  }

  // ── CONSIGNACIONES RECIBIDAS (vendedora registra ventas/devoluciones) ─────────
  if (view === "recibidas") {
    return (
      <div style={{ paddingBottom:32 }}>
        <div style={{ background:"linear-gradient(135deg,#0096c7,#0077a3)",padding:"16px 16px 20px" }}>
          <button onClick={()=>setView("main")} style={{ background:"rgba(255,255,255,.2)",border:"none",borderRadius:10,padding:"6px 12px",color:"#fff",fontFamily:"var(--hf)",fontWeight:700,fontSize:12,cursor:"pointer",marginBottom:10 }}>← Volver</button>
          <div style={{ fontSize:18,fontWeight:900,color:"#fff" }}>📥 Consignaciones recibidas</div>
          <div style={{ fontSize:11,color:"rgba(255,255,255,.75)",marginTop:2 }}>Registrá ventas y devoluciones</div>
        </div>
        <div style={{ padding:"14px 14px 0" }}>
          {consignRecibidas.filter(c=>c.status==="activa").map(consig=>(
            <div key={consig.id} style={{ background:"#fff",borderRadius:16,border:"1.5px solid var(--brd)",marginBottom:14,overflow:"hidden",boxShadow:"var(--sh)" }}>
              <div style={{ display:"flex",alignItems:"center",gap:12,padding:"13px 14px",borderBottom:"1px solid var(--brd)",background:"var(--in-l)" }}>
                <div style={{ width:40,height:40,borderRadius:12,background:consig.owner?.color||"var(--in)",display:"flex",alignItems:"center",justifyContent:"center",fontWeight:900,fontSize:16,color:"#fff",flexShrink:0 }}>
                  {consig.owner?.name?.charAt(0).toUpperCase()??"?"}
                </div>
                <div style={{ flex:1 }}>
                  <div style={{ fontWeight:700,fontSize:13 }}>De: {consig.owner?.name??"-"}</div>
                  <div style={{ fontSize:10,color:"var(--in-d)",marginTop:1 }}>Tu comisión: {consig.comision_pct}% · {new Date(consig.created_at).toLocaleDateString("es-AR")}</div>
                </div>
              </div>
              {(consig.items||[]).map(item=>{
                const p     = item.product || products.find(x=>x.id===item.product_id);
                const qDisp = item.qty_enviada - item.qty_vendida - item.qty_devuelta;
                const miComision = item.precio_venta * consig.comision_pct / 100;
                return (
                  <div key={item.id} style={{ borderBottom:"1px solid var(--brd)" }}>
                    <div style={{ display:"flex",alignItems:"center",gap:10,padding:"11px 13px" }}>
                      <ProdThumb prod={p} size={38}/>
                      <div style={{ flex:1,minWidth:0 }}>
                        <div style={{ fontWeight:700,fontSize:12,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" }}>{p?.name??"-"}</div>
                        <div style={{ fontSize:10,color:"var(--t3)",marginTop:1 }}>{fmtARS(item.precio_venta)} · tu comisión: {fmtARS(miComision)}</div>
                      </div>
                      <div style={{ textAlign:"right",flexShrink:0 }}>
                        <div style={{ fontSize:9,color:"var(--t3)",textTransform:"uppercase" }}>Disponibles</div>
                        <div style={{ fontFamily:"var(--mf)",fontWeight:900,fontSize:20,color:qDisp>0?"var(--am-d)":"var(--t4)" }}>{qDisp}</div>
                      </div>
                    </div>
                    {qDisp>0&&(
                      <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:0,borderTop:"1px solid var(--brd)" }}>
                        <button disabled={saving} onClick={()=>registrarVenta(item,consig)}
                          style={{ padding:"13px 10px",border:"none",borderRight:"1px solid var(--brd)",background:saving?"var(--em-l)":"#fff",cursor:"pointer",display:"flex",flexDirection:"column",alignItems:"center",gap:4 }}>
                          <span style={{ fontSize:22 }}>{saving?<Spin/>:"💵"}</span>
                          <div style={{ fontWeight:800,fontSize:12,color:"var(--em-d)" }}>Vendí 1</div>
                          <div style={{ fontSize:10,color:"var(--t3)" }}>Tu ganancia: {fmtARS(miComision)}</div>
                        </button>
                        <button disabled={saving} onClick={()=>registrarDevolucion(item,consig)}
                          style={{ padding:"13px 10px",border:"none",background:saving?"var(--cr-l)":"#fff",cursor:"pointer",display:"flex",flexDirection:"column",alignItems:"center",gap:4 }}>
                          <span style={{ fontSize:22 }}>{saving?<Spin/>:"↩️"}</span>
                          <div style={{ fontWeight:800,fontSize:12,color:"var(--cr)" }}>Devolver 1</div>
                          <div style={{ fontSize:10,color:"var(--t3)" }}>Vuelve al propietario</div>
                        </button>
                      </div>
                    )}
                    {qDisp===0&&item.qty_vendida>0&&(
                      <div style={{ padding:"11px 13px",background:"var(--em-l)",display:"flex",alignItems:"center",gap:8,borderTop:"1px solid var(--brd)" }}>
                        <span style={{ fontSize:18 }}>✅</span>
                        <div style={{ fontSize:12,color:"var(--em-d)",fontWeight:700 }}>Todo vendido · {fmtARS(item.qty_vendida*miComision)} fue tu comisión</div>
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          ))}
        </div>
      </div>
    );
  }

  return null;
}
