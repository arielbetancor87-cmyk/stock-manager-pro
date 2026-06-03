-- ═══════════════════════════════════════════════════════════
--  SCHEMA DE CONSIGNACIÓN — ejecutar en Supabase SQL Editor
-- ═══════════════════════════════════════════════════════════

-- 1. Tabla principal de consignaciones (agrupa items)
CREATE TABLE IF NOT EXISTS public.consignaciones (
  id            UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  owner_id      UUID NOT NULL REFERENCES public.users(id),   -- quien envía
  vendedora_id  UUID NOT NULL REFERENCES public.users(id),   -- quien recibe
  status        TEXT NOT NULL DEFAULT 'activa'
                CHECK (status IN ('activa','cerrada','cancelada')),
  comision_pct  NUMERIC(5,2) NOT NULL DEFAULT 30,            -- % vendedora
  notas         TEXT,
  created_at    TIMESTAMPTZ DEFAULT NOW(),
  updated_at    TIMESTAMPTZ DEFAULT NOW()
);

-- 2. Items dentro de cada consignación
CREATE TABLE IF NOT EXISTS public.consignacion_items (
  id                UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  consignacion_id   UUID NOT NULL REFERENCES public.consignaciones(id) ON DELETE CASCADE,
  product_id        UUID NOT NULL REFERENCES public.products(id),
  qty_enviada       INT NOT NULL CHECK (qty_enviada > 0),
  qty_vendida       INT NOT NULL DEFAULT 0,
  qty_devuelta      INT NOT NULL DEFAULT 0,
  precio_venta      NUMERIC(12,2) NOT NULL DEFAULT 0,
  created_at        TIMESTAMPTZ DEFAULT NOW(),
  -- qty_disponible = qty_enviada - qty_vendida - qty_devuelta
  CONSTRAINT no_overflow CHECK (qty_vendida + qty_devuelta <= qty_enviada)
);

-- 3. Deudas generadas por ventas
CREATE TABLE IF NOT EXISTS public.consignacion_deudas (
  id                UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  consignacion_id   UUID NOT NULL REFERENCES public.consignaciones(id),
  item_id           UUID NOT NULL REFERENCES public.consignacion_items(id),
  owner_id          UUID NOT NULL REFERENCES public.users(id),
  vendedora_id      UUID NOT NULL REFERENCES public.users(id),
  qty               INT NOT NULL DEFAULT 1,
  monto_total       NUMERIC(12,2) NOT NULL,   -- precio × qty
  comision          NUMERIC(12,2) NOT NULL,   -- monto_total × comision_pct / 100
  monto_a_pagar     NUMERIC(12,2) NOT NULL,   -- monto_total - comision
  pagada            BOOLEAN NOT NULL DEFAULT FALSE,
  pagada_at         TIMESTAMPTZ,
  created_at        TIMESTAMPTZ DEFAULT NOW()
);

-- 4. RLS — habilitar para todas las tablas nuevas
ALTER TABLE public.consignaciones        ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.consignacion_items    ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.consignacion_deudas   ENABLE ROW LEVEL SECURITY;

-- Políticas: owner o vendedora pueden ver/editar sus consignaciones
CREATE POLICY "consig_select"  ON public.consignaciones        FOR SELECT  USING (owner_id = auth.uid() OR vendedora_id = auth.uid());
CREATE POLICY "consig_insert"  ON public.consignaciones        FOR INSERT  WITH CHECK (owner_id = auth.uid());
CREATE POLICY "consig_update"  ON public.consignaciones        FOR UPDATE  USING (owner_id = auth.uid() OR vendedora_id = auth.uid());

CREATE POLICY "items_select"   ON public.consignacion_items    FOR SELECT  USING (EXISTS (SELECT 1 FROM public.consignaciones c WHERE c.id = consignacion_id AND (c.owner_id = auth.uid() OR c.vendedora_id = auth.uid())));
CREATE POLICY "items_insert"   ON public.consignacion_items    FOR INSERT  WITH CHECK (EXISTS (SELECT 1 FROM public.consignaciones c WHERE c.id = consignacion_id AND c.owner_id = auth.uid()));
CREATE POLICY "items_update"   ON public.consignacion_items    FOR UPDATE  USING (EXISTS (SELECT 1 FROM public.consignaciones c WHERE c.id = consignacion_id AND (c.owner_id = auth.uid() OR c.vendedora_id = auth.uid())));

CREATE POLICY "deudas_select"  ON public.consignacion_deudas   FOR SELECT  USING (owner_id = auth.uid() OR vendedora_id = auth.uid());
CREATE POLICY "deudas_update"  ON public.consignacion_deudas   FOR UPDATE  USING (owner_id = auth.uid());
CREATE POLICY "deudas_insert"  ON public.consignacion_deudas   FOR INSERT  WITH CHECK (owner_id = auth.uid() OR vendedora_id = auth.uid());

-- 5. Índices de performance
CREATE INDEX IF NOT EXISTS idx_consig_owner     ON public.consignaciones(owner_id);
CREATE INDEX IF NOT EXISTS idx_consig_vendedora ON public.consignaciones(vendedora_id);
CREATE INDEX IF NOT EXISTS idx_items_consig     ON public.consignacion_items(consignacion_id);
CREATE INDEX IF NOT EXISTS idx_deudas_owner     ON public.consignacion_deudas(owner_id);
CREATE INDEX IF NOT EXISTS idx_deudas_vendedora ON public.consignacion_deudas(vendedora_id);
