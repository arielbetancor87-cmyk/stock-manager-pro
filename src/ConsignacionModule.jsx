/**
 * ConsignacionModule.jsx
 * Módulo completo de consignación con lógica real de Supabase.
 * Flujo 1: Confirmar entrega en consignación a una revendedora.
 * Flujo 2: Rendición — Vendió 1 / Devolvió 1.
 *
 * Props:
 *   sb          — instancia de Supabase Client
 *   me          — objeto del usuario logueado { id, name, role }
 *   products    — array de productos del catálogo
 *   inventory   — array de filas de inventory del usuario actual
 *   contacts    — array de contactos/revendedoras
 *   transfers   — array de transfers del usuario actual
 *   onRefresh   — callback para recargar datos desde el padre
 *   toast       — función toast(title, body, type)
 *   fmtARS      — función de formateo de moneda
 */

import { useState, useEffect, useCallback, memo } from "react";

// ─── SUBCOMPONENTES INTERNOS ──────────────────────────────────────────────────

const ProdThumb = memo(function ProdThumb({ prod, size = 40 }) {
  if (!prod) return <div style={{ width: size, height: size, borderRadius: 10, background: "var(--bg2)", flexShrink: 0 }} />;
  if (prod.photo_url) return <img src={prod.photo_url} alt="" style={{ width: size, height: size, borderRadius: 10, objectFit: "cover", flexShrink: 0, border: "1px solid var(--brd)" }} />;
  return <div style={{ width: size, height: size, borderRadius: 10, background: "var(--bg2)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: Math.round(size * 0.5), flexShrink: 0 }}>{prod.emoji || "📦"}</div>;
});

const Spinner = () => (
  <div style={{ display: "inline-block", width: 16, height: 16, border: "2px solid rgba(255,255,255,.4)", borderTopColor: "#fff", borderRadius: "50%", animation: "spin 0.8s linear infinite" }} />
);

// ─── MÓDULO PRINCIPAL ─────────────────────────────────────────────────────────

export default function ConsignacionModule({ sb, me, products, inventory, contacts, transfers, onRefresh, toast, fmtARS }) {

  // ── VISTAS ──────────────────────────────────────────────────────────────────
  const [view,         setView]         = useState("main");       // "main" | "nueva" | "rendicion"
  const [loading,      setLoading]      = useState(false);

  // ── FLUJO 1: Confirmar nueva entrega ────────────────────────────────────────
  const [supplierId,   setSupplierId]   = useState("");           // revendedora seleccionada
  const [consigCart,   setConsigCart]   = useState({});          // { product_id: qty }
  const [prodSearch,   setProdSearch]   = useState("");

  // ── FLUJO 2: Rendición ─────────────────────────────────────────────────────
  const [rendSupplier, setRendSupplier] = useState(null);         // revendedora en rendición
  const [rendItems,    setRendItems]    = useState([]);           // consignas activas de esa rev.
  const [rendLoading,  setRendLoading]  = useState({});          // { txId_accion: bool }

  // ── DERIVADOS ───────────────────────────────────────────────────────────────

  // Consignas activas enviadas por mí: transfers confirmados con qty > 0
  const activeSent = transfers.filter(t =>
    t.from_user_id === me.id && t.status === "confirmed" && t.qty > 0
  );

  // Agrupar por revendedora
  const byRevend = activeSent.reduce((acc, tx) => {
    if (!acc[tx.to_user_id]) acc[tx.to_user_id] = { user: tx.to_user, items: [] };
    acc[tx.to_user_id].items.push(tx);
    return acc;
  }, {});
  const revendedoras = Object.values(byRevend);

  // Productos filtrados para el carrito
  const prodsFilt = products.filter(p => {
    const q = prodSearch.toLowerCase();
    if (!q) return true;
    return p.name.toLowerCase().includes(q) || p.sku.toLowerCase().includes(q);
  });

  // Total del carrito actual
  const cartTotal = Object.entries(consigCart).reduce((s, [pid, qty]) => {
    const p = products.find(x => x.id === pid);
    return s + (p ? p.price * qty : 0);
  }, 0);
  const cartUnits = Object.values(consigCart).reduce((s, q) => s + q, 0);

  // ────────────────────────────────────────────────────────────────────────────
  // FLUJO 1 — CONFIRMAR ENTREGA EN CONSIGNACIÓN
  // ────────────────────────────────────────────────────────────────────────────
  async function confirmarEntrega() {
    if (!supplierId) { toast("Seleccioná una revendedora", "", "e"); return; }
    const entries = Object.entries(consigCart).filter(([, qty]) => qty > 0);
    if (!entries.length) { toast("Agregá al menos un producto", "", "e"); return; }

    setLoading(true);
    try {
      for (const [product_id, qty] of entries) {
        const prod = products.find(p => p.id === product_id);

        // a) Restar del stock propio del emisor (me)
        const { data: myInv, error: myInvErr } = await sb
          .from("inventory")
          .select("*")
          .eq("user_id", me.id)
          .eq("product_id", product_id)
          .maybeSingle();

        if (myInvErr) throw new Error(`Error leyendo mi stock: ${myInvErr.message}`);
        if (!myInv || myInv.qty_available < qty) {
          throw new Error(`Stock insuficiente de "${prod?.name}". Tenés ${myInv?.qty_available ?? 0} u.`);
        }

        const { error: deductErr } = await sb
          .from("inventory")
          .update({ qty_available: myInv.qty_available - qty })
          .eq("id", myInv.id);

        if (deductErr) throw new Error(`Error restando stock: ${deductErr.message}`);

        // b) Insertar o actualizar en inventory de la revendedora con source = 'consigna'
        const { data: revInv } = await sb
          .from("inventory")
          .select("*")
          .eq("user_id", supplierId)
          .eq("product_id", product_id)
          .eq("source", "consigna")
          .maybeSingle();

        if (revInv) {
          const { error: updErr } = await sb
            .from("inventory")
            .update({ qty_available: revInv.qty_available + qty })
            .eq("id", revInv.id);
          if (updErr) throw new Error(`Error actualizando consigna: ${updErr.message}`);
        } else {
          const { error: insErr } = await sb
            .from("inventory")
            .insert({
              user_id:       supplierId,
              product_id,
              qty_available: qty,
              qty_sold:      0,
              source:        "consigna",
              supplier_id:   me.id,
            });
          if (insErr) throw new Error(`Error creando consigna: ${insErr.message}`);
        }

        // c) Crear transfer pendiente → confirmado directamente
        const { error: txErr } = await sb
          .from("transfers")
          .insert({
            from_user_id:  me.id,
            to_user_id:    supplierId,
            product_id,
            qty,
            status:        "confirmed",
            confirmed_at:  new Date().toISOString(),
          });
        if (txErr) throw new Error(`Error registrando transfer: ${txErr.message}`);

        // d) Notificar a la revendedora
        await sb.from("notifications").insert({
          to_user_id: supplierId,
          from_name:  me.name,
          type:       "transfer",
          message:    `${me.name} te asignó ${qty}x "${prod?.name}" en consignación.`,
        });
      }

      toast("✅ Consignación registrada", `${cartUnits} u. entregadas a ${contacts.find(c => c.id === supplierId)?.name ?? "revendedora"}`, "s");
      setConsigCart({});
      setSupplierId("");
      setView("main");
      onRefresh();
    } catch (err) {
      toast("Error al confirmar", err.message, "e");
    } finally {
      setLoading(false);
    }
  }

  // ────────────────────────────────────────────────────────────────────────────
  // FLUJO 2 — RENDICIÓN: cargar items de una revendedora
  // ────────────────────────────────────────────────────────────────────────────
  async function abrirRendicion(revUser) {
    setRendSupplier(revUser);
    setRendLoading({});
    // Buscar todos los transfers confirmados de esta revendedora con qty > 0
    const items = transfers.filter(t =>
      t.from_user_id === me.id &&
      t.to_user_id   === revUser.id &&
      t.status       === "confirmed" &&
      t.qty > 0
    );
    setRendItems(items);
    setView("rendicion");
  }

  // ── Vendió 1 ────────────────────────────────────────────────────────────────
  async function vendioUno(tx) {
    const key = `${tx.id}_venta`;
    setRendLoading(p => ({ ...p, [key]: true }));
    const prod = tx.product || products.find(p => p.id === tx.product_id);

    try {
      // 1. Restar 1 del inventario de consigna de la revendedora
      const { data: revInv, error: revErr } = await sb
        .from("inventory")
        .select("*")
        .eq("user_id", tx.to_user_id)
        .eq("product_id", tx.product_id)
        .eq("source", "consigna")
        .maybeSingle();

      if (revErr) throw new Error(revErr.message);
      if (!revInv || revInv.qty_available < 1) throw new Error("La revendedora no tiene unidades disponibles.");

      const { error: updInvErr } = await sb
        .from("inventory")
        .update({
          qty_available: revInv.qty_available - 1,
          qty_sold:      (revInv.qty_sold ?? 0) + 1,
        })
        .eq("id", revInv.id);
      if (updInvErr) throw new Error(updInvErr.message);

      // 2. Reducir qty del transfer
      const newQty = tx.qty - 1;
      const { error: txUpdErr } = await sb
        .from("transfers")
        .update(newQty > 0 ? { qty: newQty } : { qty: 0, status: "rejected" })
        .eq("id", tx.id);
      if (txUpdErr) throw new Error(txUpdErr.message);

      // 3. Registrar venta en sale_logs (ingreso para mí como emisor)
      const { error: logErr } = await sb
        .from("sale_logs")
        .insert({
          user_id:    me.id,
          product_id: tx.product_id,
          qty:        1,
          sale_price: prod?.price ?? 0,
          source:     "consignment",
        });
      if (logErr) console.warn("sale_log error:", logErr.message); // no-fatal

      // 4. Notificar al emisor (yo) desde el punto de vista de la revendedora
      await sb.from("notifications").insert({
        to_user_id: me.id,
        from_name:  rendSupplier?.name ?? "Revendedora",
        type:       "sale",
        message:    `💰 ${rendSupplier?.name} vendió 1x "${prod?.name}". Ganancia: ${fmtARS(prod?.price ?? 0)}`,
      });

      // Actualizar estado local
      setRendItems(prev => prev.map(t2 =>
        t2.id === tx.id
          ? { ...t2, qty: newQty, status: newQty > 0 ? "confirmed" : "rejected" }
          : t2
      ).filter(t2 => t2.qty > 0));

      toast("💰 Venta registrada", `${prod?.name} — ${fmtARS(prod?.price ?? 0)}`, "s");
      onRefresh();
    } catch (err) {
      toast("Error al registrar venta", err.message, "e");
    } finally {
      setRendLoading(p => ({ ...p, [key]: false }));
    }
  }

  // ── Devolvió 1 ──────────────────────────────────────────────────────────────
  async function devolvioUno(tx) {
    const key = `${tx.id}_devol`;
    setRendLoading(p => ({ ...p, [key]: true }));
    const prod = tx.product || products.find(p => p.id === tx.product_id);

    try {
      // 1. Restar 1 del inventario de consigna de la revendedora
      const { data: revInv, error: revErr } = await sb
        .from("inventory")
        .select("*")
        .eq("user_id", tx.to_user_id)
        .eq("product_id", tx.product_id)
        .eq("source", "consigna")
        .maybeSingle();

      if (revErr) throw new Error(revErr.message);
      if (!revInv || revInv.qty_available < 1) throw new Error("Sin unidades para devolver.");

      const { error: deductErr } = await sb
        .from("inventory")
        .update({ qty_available: revInv.qty_available - 1 })
        .eq("id", revInv.id);
      if (deductErr) throw new Error(deductErr.message);

      // 2. Sumar 1 de vuelta al stock propio del emisor (me)
      const { data: myInv } = await sb
        .from("inventory")
        .select("*")
        .eq("user_id", me.id)
        .eq("product_id", tx.product_id)
        .maybeSingle();

      if (myInv) {
        await sb.from("inventory")
          .update({ qty_available: myInv.qty_available + 1 })
          .eq("id", myInv.id);
      } else {
        await sb.from("inventory")
          .insert({ user_id: me.id, product_id: tx.product_id, qty_available: 1, qty_sold: 0, source: "own" });
      }

      // 3. Actualizar transfer
      const newQty = tx.qty - 1;
      await sb.from("transfers")
        .update(newQty > 0 ? { qty: newQty } : { qty: 0, status: "rejected" })
        .eq("id", tx.id);

      // 4. Notificar
      await sb.from("notifications").insert({
        to_user_id: me.id,
        from_name:  rendSupplier?.name ?? "Revendedora",
        type:       "confirm",
        message:    `↩️ ${rendSupplier?.name} devolvió 1x "${prod?.name}". Stock restituido.`,
      });

      // Actualizar estado local
      setRendItems(prev => prev.map(t2 =>
        t2.id === tx.id
          ? { ...t2, qty: newQty, status: newQty > 0 ? "confirmed" : "rejected" }
          : t2
      ).filter(t2 => t2.qty > 0));

      toast("↩️ Devolución registrada", `+1 "${prod?.name}" en tu stock`, "s");
      onRefresh();
    } catch (err) {
      toast("Error al registrar devolución", err.message, "e");
    } finally {
      setRendLoading(p => ({ ...p, [key]: false }));
    }
  }

  // ────────────────────────────────────────────────────────────────────────────
  // RENDER
  // ────────────────────────────────────────────────────────────────────────────

  // ── MAIN VIEW ───────────────────────────────────────────────────────────────
  if (view === "main") {
    const totalConsignado = activeSent.reduce((s, t) => s + t.qty, 0);
    const totalValor      = activeSent.reduce((s, t) => {
      const p = t.product || products.find(x => x.id === t.product_id);
      return s + (p ? p.price * t.qty : 0);
    }, 0);

    return (
      <div style={{ padding: "0 0 32px" }}>
        {/* Header gradient */}
        <div style={{ background: "linear-gradient(135deg,#6d28d9,#a855f7)", padding: "20px 16px 24px" }}>
          <div style={{ fontSize: 22, fontWeight: 900, color: "#fff", marginBottom: 4 }}>🤝 Consignaciones</div>
          <div style={{ fontSize: 12, color: "rgba(255,255,255,.75)" }}>Gestión de mercadería entregada</div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginTop: 16 }}>
            {[
              { ico: "👥", val: revendedoras.length, lbl: "Revendedoras activas" },
              { ico: "📦", val: totalConsignado,      lbl: "Unidades en consigna" },
              { ico: "💰", val: fmtARS(totalValor),   lbl: "Valor total",         small: true },
              { ico: "🔄", val: activeSent.length,    lbl: "Ítems activos" },
            ].map((m, i) => (
              <div key={i} style={{ background: "rgba(255,255,255,.15)", borderRadius: 16, padding: "14px 12px", border: "1px solid rgba(255,255,255,.2)" }}>
                <div style={{ fontSize: 20, marginBottom: 4 }}>{m.ico}</div>
                <div style={{ fontFamily: "var(--mf)", fontWeight: 900, fontSize: m.small ? 13 : 22, color: "#fff", lineHeight: 1 }}>{m.val}</div>
                <div style={{ fontSize: 10, color: "rgba(255,255,255,.7)", marginTop: 4, fontWeight: 600 }}>{m.lbl}</div>
              </div>
            ))}
          </div>
        </div>

        <div style={{ padding: "16px 14px 0" }}>
          {/* CTA nueva consignación */}
          <button
            onClick={() => { setView("nueva"); setConsigCart({}); setSupplierId(""); }}
            style={{ width: "100%", padding: "16px", borderRadius: 18, background: "linear-gradient(135deg,#ff7a00,#e06a00)", border: "none", cursor: "pointer", display: "flex", alignItems: "center", gap: 14, marginBottom: 16, boxShadow: "0 6px 20px rgba(255,122,0,.3)" }}
          >
            <div style={{ width: 44, height: 44, borderRadius: 14, background: "rgba(255,255,255,.2)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 22, flexShrink: 0 }}>📤</div>
            <div style={{ textAlign: "left" }}>
              <div style={{ fontSize: 15, fontWeight: 900, color: "#fff" }}>Nueva entrega en consigna</div>
              <div style={{ fontSize: 11, color: "rgba(255,255,255,.8)", marginTop: 2 }}>Asignar productos a una revendedora</div>
            </div>
          </button>

          {/* Lista de revendedoras activas */}
          {revendedoras.length === 0 ? (
            <div style={{ textAlign: "center", padding: "40px 20px", color: "var(--t3)", fontSize: 13 }}>
              <div style={{ fontSize: 40, marginBottom: 8 }}>📭</div>
              Sin consignaciones activas.<br />Usá el botón de arriba para asignar mercadería.
            </div>
          ) : (
            revendedoras.map(({ user: u, items }) => {
              const totalU = items.reduce((s, t) => s + t.qty, 0);
              const totalV = items.reduce((s, t) => {
                const p = t.product || products.find(x => x.id === t.product_id);
                return s + (p ? p.price * t.qty : 0);
              }, 0);
              return (
                <div key={u?.id ?? "?"} style={{ background: "#fff", borderRadius: 18, border: "1.5px solid var(--brd)", marginBottom: 12, overflow: "hidden", boxShadow: "var(--sh)" }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 12, padding: "14px 16px", borderBottom: "1px solid var(--brd)" }}>
                    <div style={{ width: 44, height: 44, borderRadius: 14, background: u?.color || "var(--in)", display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 900, fontSize: 18, color: "#fff", flexShrink: 0 }}>
                      {u?.name?.charAt(0).toUpperCase() ?? "?"}
                    </div>
                    <div style={{ flex: 1 }}>
                      <div style={{ fontWeight: 800, fontSize: 15 }}>{u?.name ?? "—"}</div>
                      <div style={{ fontSize: 11, color: "var(--t3)", marginTop: 2 }}>{items.length} producto{items.length !== 1 ? "s" : ""} · {totalU} u. · {fmtARS(totalV)}</div>
                    </div>
                    <button
                      onClick={() => abrirRendicion(u)}
                      style={{ background: "var(--am-l)", border: "1.5px solid rgba(255,122,0,.25)", borderRadius: 12, padding: "8px 14px", cursor: "pointer", fontFamily: "var(--hf)", fontWeight: 800, fontSize: 12, color: "var(--am-d)" }}
                    >
                      Rendir 📊
                    </button>
                  </div>
                  {/* Mini-lista de productos */}
                  {items.slice(0, 3).map(tx => {
                    const p = tx.product || products.find(x => x.id === tx.product_id);
                    return (
                      <div key={tx.id} style={{ display: "flex", alignItems: "center", gap: 10, padding: "10px 16px", borderBottom: "1px solid var(--brd)" }}>
                        <ProdThumb prod={p} size={32} />
                        <div style={{ flex: 1, fontSize: 12, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", fontWeight: 600 }}>{p?.name ?? "—"}</div>
                        <span style={{ fontFamily: "var(--mf)", fontWeight: 900, color: "var(--am-d)", fontSize: 14 }}>{tx.qty} u.</span>
                      </div>
                    );
                  })}
                  {items.length > 3 && (
                    <div style={{ padding: "8px 16px", fontSize: 11, color: "var(--t3)", textAlign: "center" }}>+{items.length - 3} más</div>
                  )}
                </div>
              );
            })
          )}
        </div>
      </div>
    );
  }

  // ── NUEVA CONSIGNACIÓN ───────────────────────────────────────────────────────
  if (view === "nueva") {
    const selSupplier = contacts.find(c => c.id === supplierId);
    return (
      <div style={{ padding: "0 0 32px" }}>
        <div style={{ background: "linear-gradient(135deg,#6d28d9,#a855f7)", padding: "16px 16px 20px" }}>
          <button onClick={() => setView("main")} style={{ background: "rgba(255,255,255,.2)", border: "none", borderRadius: 10, padding: "6px 12px", color: "#fff", fontFamily: "var(--hf)", fontWeight: 700, fontSize: 12, cursor: "pointer", marginBottom: 10 }}>← Volver</button>
          <div style={{ fontSize: 18, fontWeight: 900, color: "#fff" }}>📤 Nueva Consignación</div>
          <div style={{ fontSize: 11, color: "rgba(255,255,255,.7)", marginTop: 2 }}>Seleccioná revendedora y productos</div>
        </div>

        <div style={{ padding: "16px 14px 0" }}>

          {/* Paso 1: Revendedora */}
          <div style={{ marginBottom: 16 }}>
            <div style={{ fontSize: 11, fontWeight: 800, color: "var(--t3)", textTransform: "uppercase", letterSpacing: ".08em", marginBottom: 10 }}>1. Revendedora</div>
            {contacts.length === 0
              ? <div style={{ textAlign: "center", color: "var(--t3)", fontSize: 13, padding: "16px" }}>Sin contactos. Agregá uno en la tab Red.</div>
              : contacts.map(c => (
                <div key={c.id} onClick={() => setSupplierId(c.id)}
                  style={{ display: "flex", alignItems: "center", gap: 12, padding: "13px 15px", borderRadius: 16, border: `2px solid ${supplierId === c.id ? "var(--am)" : "var(--brd)"}`, background: supplierId === c.id ? "var(--am-l)" : "#fff", marginBottom: 8, cursor: "pointer", transition: "all .14s", boxShadow: "var(--sh)" }}>
                  <div style={{ width: 40, height: 40, borderRadius: 12, background: c.color || "var(--in)", display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 900, fontSize: 16, color: "#fff", flexShrink: 0 }}>
                    {c.name.charAt(0).toUpperCase()}
                  </div>
                  <div style={{ flex: 1 }}>
                    <div style={{ fontWeight: 700, fontSize: 14 }}>{c.name}</div>
                    <div style={{ fontSize: 11, color: "var(--t3)", marginTop: 1 }}>{c.email}</div>
                  </div>
                  {supplierId === c.id && <div style={{ fontSize: 18, color: "var(--am-d)" }}>✓</div>}
                </div>
              ))
            }
          </div>

          {/* Paso 2: Productos */}
          <div style={{ marginBottom: 16 }}>
            <div style={{ fontSize: 11, fontWeight: 800, color: "var(--t3)", textTransform: "uppercase", letterSpacing: ".08em", marginBottom: 10 }}>2. Productos</div>
            <div style={{ position: "relative", marginBottom: 10 }}>
              <span style={{ position: "absolute", left: 12, top: "50%", transform: "translateY(-50%)", color: "var(--t3)", fontSize: 16 }}>🔍</span>
              <input
                style={{ width: "100%", padding: "11px 12px 11px 38px", borderRadius: 12, border: "1.5px solid var(--brd)", background: "#fff", fontFamily: "var(--hf)", fontSize: 14, outline: "none", color: "var(--t1)" }}
                placeholder="Buscar producto..."
                value={prodSearch}
                onChange={e => setProdSearch(e.target.value)}
              />
            </div>
            {prodsFilt.slice(0, 20).map(p => {
              const myInvRow = inventory.find(i => i.product_id === p.id && (!i.source || i.source === "own"));
              const avail = myInvRow?.qty_available ?? 0;
              const inCart = consigCart[p.id] ?? 0;
              return (
                <div key={p.id} style={{ display: "flex", alignItems: "center", gap: 12, padding: "12px 14px", borderRadius: 14, border: `1.5px solid ${inCart > 0 ? "var(--in)" : "var(--brd)"}`, background: inCart > 0 ? "var(--in-l)" : "#fff", marginBottom: 8, boxShadow: "var(--sh)" }}>
                  <ProdThumb prod={p} size={40} />
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ fontWeight: 700, fontSize: 13, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{p.name}</div>
                    <div style={{ fontSize: 11, color: "var(--t3)", marginTop: 2 }}>{p.sku} · {fmtARS(p.price)} · {avail} disp.</div>
                  </div>
                  {/* QtyControl */}
                  <div style={{ display: "flex", alignItems: "center", border: "1.5px solid var(--brd)", borderRadius: 10, overflow: "hidden", flexShrink: 0 }}>
                    <button
                      onClick={() => setConsigCart(c => { const n = { ...c }; if ((n[p.id] ?? 0) <= 1) { delete n[p.id]; } else { n[p.id] = n[p.id] - 1; } return n; })}
                      style={{ width: 30, height: 30, border: "none", background: "var(--bg)", cursor: "pointer", fontSize: 16, display: "flex", alignItems: "center", justifyContent: "center" }}
                    >−</button>
                    <div style={{ minWidth: 28, textAlign: "center", fontFamily: "var(--mf)", fontWeight: 700, fontSize: 14, color: "var(--in)" }}>{inCart}</div>
                    <button
                      onClick={() => {
                        if (inCart >= avail) { return; }
                        setConsigCart(c => ({ ...c, [p.id]: (c[p.id] ?? 0) + 1 }));
                      }}
                      disabled={inCart >= avail}
                      style={{ width: 30, height: 30, border: "none", background: inCart >= avail ? "var(--bg2)" : "var(--in-l)", cursor: inCart >= avail ? "not-allowed" : "pointer", fontSize: 16, display: "flex", alignItems: "center", justifyContent: "center", opacity: inCart >= avail ? 0.4 : 1 }}
                    >+</button>
                  </div>
                </div>
              );
            })}
          </div>

          {/* Resumen del carrito */}
          {cartUnits > 0 && (
            <div style={{ background: "var(--em-l)", border: "1.5px solid rgba(0,184,122,.2)", borderRadius: 14, padding: "14px 16px", marginBottom: 16 }}>
              <div style={{ fontSize: 12, fontWeight: 700, color: "var(--em-d)", marginBottom: 8, textTransform: "uppercase", letterSpacing: ".06em" }}>Resumen</div>
              {Object.entries(consigCart).filter(([, qty]) => qty > 0).map(([pid, qty]) => {
                const p = products.find(x => x.id === pid);
                return (
                  <div key={pid} style={{ display: "flex", justifyContent: "space-between", fontSize: 13, color: "var(--t2)", marginBottom: 4 }}>
                    <span>{p?.emoji} {p?.name}</span>
                    <span style={{ fontWeight: 700 }}>{qty} u. · {fmtARS((p?.price ?? 0) * qty)}</span>
                  </div>
                );
              })}
              <div style={{ display: "flex", justifyContent: "space-between", fontWeight: 900, fontSize: 15, color: "var(--em-d)", marginTop: 10, paddingTop: 10, borderTop: "1px solid rgba(0,184,122,.2)" }}>
                <span>Total consignado</span>
                <span style={{ fontFamily: "var(--mf)" }}>{fmtARS(cartTotal)}</span>
              </div>
            </div>
          )}

          {/* Botón confirmar */}
          <button
            onClick={confirmarEntrega}
            disabled={loading || !supplierId || cartUnits === 0}
            style={{ width: "100%", padding: "17px", borderRadius: 18, border: "none", background: "linear-gradient(135deg,#00b87a,#009a66)", color: "#fff", fontFamily: "var(--hf)", fontWeight: 900, fontSize: 16, cursor: loading || !supplierId || cartUnits === 0 ? "not-allowed" : "pointer", opacity: !supplierId || cartUnits === 0 ? 0.4 : 1, display: "flex", alignItems: "center", justifyContent: "center", gap: 10, boxShadow: "0 6px 20px rgba(0,184,122,.3)", transition: "all .15s" }}
          >
            {loading ? <Spinner /> : "✅"}
            {loading ? "Registrando..." : `Confirmar entrega (${cartUnits} u.)`}
          </button>
        </div>
      </div>
    );
  }

  // ── RENDICIÓN ────────────────────────────────────────────────────────────────
  if (view === "rendicion") {
    return (
      <div style={{ padding: "0 0 32px" }}>
        <div style={{ background: "linear-gradient(135deg,#ff7a00,#e06a00)", padding: "16px 16px 20px" }}>
          <button onClick={() => setView("main")} style={{ background: "rgba(255,255,255,.2)", border: "none", borderRadius: 10, padding: "6px 12px", color: "#fff", fontFamily: "var(--hf)", fontWeight: 700, fontSize: 12, cursor: "pointer", marginBottom: 10 }}>← Volver</button>
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <div style={{ width: 46, height: 46, borderRadius: 14, background: rendSupplier?.color || "rgba(255,255,255,.25)", display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 900, fontSize: 20, color: "#fff" }}>
              {rendSupplier?.name?.charAt(0).toUpperCase() ?? "?"}
            </div>
            <div>
              <div style={{ fontSize: 18, fontWeight: 900, color: "#fff" }}>{rendSupplier?.name}</div>
              <div style={{ fontSize: 11, color: "rgba(255,255,255,.75)", marginTop: 2 }}>{rendItems.length} producto{rendItems.length !== 1 ? "s" : ""} en consigna</div>
            </div>
          </div>
        </div>

        <div style={{ padding: "16px 14px 0" }}>
          {rendItems.length === 0 ? (
            <div style={{ textAlign: "center", padding: "40px 20px", color: "var(--t3)", fontSize: 13 }}>
              <div style={{ fontSize: 40, marginBottom: 8 }}>🎉</div>
              Sin productos pendientes.<br />Toda la consigna fue rendida.
            </div>
          ) : (
            rendItems.map(tx => {
              const p = tx.product || products.find(x => x.id === tx.product_id);
              const keyV = `${tx.id}_venta`;
              const keyD = `${tx.id}_devol`;
              return (
                <div key={tx.id} style={{ background: "#fff", borderRadius: 18, border: "1.5px solid var(--brd)", marginBottom: 14, overflow: "hidden", boxShadow: "var(--sh)" }}>
                  {/* Producto info */}
                  <div style={{ display: "flex", alignItems: "center", gap: 12, padding: "14px 16px", borderBottom: "1px solid var(--brd)" }}>
                    <ProdThumb prod={p} size={48} />
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontWeight: 800, fontSize: 14, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{p?.name ?? "—"}</div>
                      <div style={{ fontSize: 11, color: "var(--t3)", marginTop: 2 }}>{p?.sku} · {fmtARS(p?.price ?? 0)}</div>
                    </div>
                    <div style={{ textAlign: "right", flexShrink: 0 }}>
                      <div style={{ fontSize: 9, color: "var(--t3)", textTransform: "uppercase", letterSpacing: ".06em" }}>Pendiente</div>
                      <div style={{ fontFamily: "var(--mf)", fontWeight: 900, fontSize: 26, color: "var(--am-d)", lineHeight: 1 }}>{tx.qty}</div>
                    </div>
                  </div>

                  {/* Acciones rápidas */}
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 0 }}>
                    {/* Vendió 1 */}
                    <button
                      onClick={() => vendioUno(tx)}
                      disabled={!!rendLoading[keyV] || !!rendLoading[keyD]}
                      style={{ padding: "16px 12px", border: "none", borderRight: "1px solid var(--brd)", background: rendLoading[keyV] ? "var(--em-l)" : "#fff", cursor: "pointer", display: "flex", flexDirection: "column", alignItems: "center", gap: 6, transition: "background .13s" }}
                    >
                      {rendLoading[keyV] ? <Spinner /> : <span style={{ fontSize: 26 }}>💵</span>}
                      <div style={{ fontFamily: "var(--hf)", fontWeight: 800, fontSize: 13, color: "var(--em-d)" }}>
                        {rendLoading[keyV] ? "..." : "Vendió 1"}
                      </div>
                      <div style={{ fontSize: 10, color: "var(--t3)" }}>+{fmtARS(p?.price ?? 0)} ganancia</div>
                    </button>

                    {/* Devolvió 1 */}
                    <button
                      onClick={() => devolvioUno(tx)}
                      disabled={!!rendLoading[keyV] || !!rendLoading[keyD]}
                      style={{ padding: "16px 12px", border: "none", background: rendLoading[keyD] ? "var(--cr-l)" : "#fff", cursor: "pointer", display: "flex", flexDirection: "column", alignItems: "center", gap: 6, transition: "background .13s" }}
                    >
                      {rendLoading[keyD] ? <Spinner /> : <span style={{ fontSize: 26 }}>↩️</span>}
                      <div style={{ fontFamily: "var(--hf)", fontWeight: 800, fontSize: 13, color: "var(--cr)" }}>
                        {rendLoading[keyD] ? "..." : "Devolvió 1"}
                      </div>
                      <div style={{ fontSize: 10, color: "var(--t3)" }}>Vuelve a tu stock</div>
                    </button>
                  </div>
                </div>
              );
            })
          )}
        </div>
      </div>
    );
  }

  return null;
}
