import openpyxl
import json
from collections import defaultdict
from datetime import datetime

def safe_float(v):
    try: return float(v) if v not in (None, 'None', '') else 0.0
    except: return 0.0

def safe_date(v):
    if v is None or v == 'None': return None
    if isinstance(v, datetime): return v
    try:
        s = str(v).strip()
        for fmt in ('%Y-%m-%d %H:%M:%S.%f', '%Y-%m-%d %H:%M:%S', '%Y-%m-%d'):
            try: return datetime.strptime(s[:19], fmt)
            except: pass
    except: pass
    return None

base = "C:/Users/Montolivo/Dropbox/Asesorias PYMES/1. Proyectos Actuales/2. MONTOLIVO/2. Proyecto/Indicadores Montolivo/Compras/Bases para actualizar Claude/"

months_all = ["2025-01","2025-02","2025-03","2025-04","2025-05","2025-06","2025-07","2025-08",
              "2025-09","2025-10","2025-11","2025-12","2026-01","2026-02","2026-03","2026-04","2026-05","2026-06","2026-07","2026-08","2026-09"]
month_idx = {m:i for i,m in enumerate(months_all)}

def abbrev_provider(name):
    n = name
    for suf in [" SOCIEDAD POR ACCIONES SIMPLIFICADAS", " S.A.S.", " S.A.S", " S.A.", " S.A", " LTDA.", " LTDA", " SAS"]:
        if n.upper().endswith(suf.upper()):
            n = n[: -len(suf)]
            break
    n = n.strip()
    LIMIT = 30
    if len(n) <= LIMIT:
        return n
    cut = n[:LIMIT]
    if ' ' in cut:
        cut = cut[:cut.rfind(' ')]
    return cut.strip()

# ════════════════════════════════════════════════════════════════
# PASS 1: Historico de compras por item.xlsx
#   -> LT_DATA, full supplier_monthly (for prov_var), full item_facturas
#      (for item_count top6), categoria proxy (Referencia prefix x Cantidad)
# ════════════════════════════════════════════════════════════════
print("PASS 1: Historico de compras...")
wb2 = openpyxl.load_workbook(base + "Histórico de compras por item.xlsx", data_only=True, read_only=True)
ws2 = wb2.active

REF_CAT = {
    "MP": "MATERIA PRIMA", "PC": "PTO TERMINADO", "AC": "INV ASEO Y CAFETERIA",
    "MJ": "MENAJE", "EE": "EMPAQUES", "PP": "PAPELERIA", "RP": "REPUESTOS",
}

lt_real, lt_prom, lt_late = [], [], []  # parallel arrays per linked row
lt_rows = []  # (fecha_oc, real, prom, on_time, late_days, prov, item)
supplier_monthly_full = defaultdict(lambda: [0.0]*len(months_all))
item_facturas_full = defaultdict(int)
cant_by_refcat = defaultdict(float)
cant_by_um_hist = defaultdict(float)

headers2 = None
n_total = 0
n_linked = 0
for i, row in enumerate(ws2.iter_rows(values_only=True)):
    if i == 0:
        headers2 = [str(c).strip() if c else '' for c in row]
        continue
    r = dict(zip(headers2, row))
    n_total += 1
    fecha = safe_date(r.get('Fecha'))
    neto = safe_float(r.get('Valor neto local'))
    prov = str(r.get('Razon social proveedor') or r.get('Razón social proveedor') or '').strip()
    item = str(r.get('Desc. item') or '').strip()
    cant = safe_float(r.get('Cantidad'))
    ref = str(r.get('Referencia') or '').strip()
    um = str(r.get('U.M.') or '').strip()

    if fecha:
        key = f"{fecha.year}-{fecha.month:02d}"
        if key in month_idx and prov:
            supplier_monthly_full[prov][month_idx[key]] += neto
    if item:
        item_facturas_full[item] += 1
    if um:
        cant_by_um_hist[um] += cant

    prefix = ref[:2].upper()
    cat = REF_CAT.get(prefix, "OTROS")
    cant_by_refcat[cat] += cant

    docto_orden = r.get('Docto. orden')
    if docto_orden not in (None, '', 'None'):
        f_oc = safe_date(r.get('Fecha documento OC'))
        f_prom = safe_date(r.get('Fecha de entrega item'))
        if fecha and f_oc:
            real = (fecha - f_oc).days
            if 0 <= real <= 365:
                prom = (f_prom - f_oc).days if f_prom else None
                if prom is not None and not (0 <= prom <= 365):
                    prom = None
                on_time = (f_prom is not None and fecha <= f_prom)
                late_days = (fecha - f_prom).days if f_prom else None
                n_linked += 1
                lt_rows.append((f_oc, real, prom, on_time, late_days, prov, item))

wb2.close()
print(f"  Filas totales: {n_total}  |  Filas con Docto.orden vinculado: {n_linked}")

# ---- LT_DATA ----
total = len(lt_rows)
avg_real = sum(r[1] for r in lt_rows) / total if total else 0
prom_vals = [r[2] for r in lt_rows if r[2] is not None]
avg_prom = sum(prom_vals) / len(prom_vals) if prom_vals else 0
n_ontime = sum(1 for r in lt_rows if r[3])
pct_ontime = n_ontime / total * 100 if total else 0

def stats_subset(rows):
    t = len(rows)
    if t == 0:
        return {"avgReal": 0, "avgProm": 0, "pctOn": 0, "count": 0}
    ar = sum(r[1] for r in rows) / t
    pv = [r[2] for r in rows if r[2] is not None]
    ap = sum(pv) / len(pv) if pv else 0
    pon = sum(1 for r in rows if r[3]) / t * 100
    return {"avgReal": round(ar, 1), "avgProm": round(ap, 1), "pctOn": round(pon, 1), "count": t}

rows_2025 = [r for r in lt_rows if r[0].year == 2025]
rows_2026 = [r for r in lt_rows if r[0].year == 2026]
por2025 = stats_subset(rows_2025)
por2026 = stats_subset(rows_2026)

monthly_rows = defaultdict(list)
for r in lt_rows:
    key = f"{r[0].year}-{r[0].month:02d}"
    if key in month_idx:
        monthly_rows[key].append(r)

monthlyReal, monthlyProm, monthlyPctOn, monthlyCount = [], [], [], []
for m in months_all:
    rs = monthly_rows.get(m, [])
    s = stats_subset(rs)
    monthlyReal.append(s["avgReal"])
    monthlyProm.append(s["avgProm"])
    monthlyPctOn.append(s["pctOn"])
    monthlyCount.append(s["count"])

dist = {"ontime": 0, "late1_3": 0, "late4_7": 0, "late8plus": 0}
for r in lt_rows:
    ld = r[4]
    if ld is None:
        continue
    if ld <= 0:
        dist["ontime"] += 1
    elif ld <= 3:
        dist["late1_3"] += 1
    elif ld <= 7:
        dist["late4_7"] += 1
    else:
        dist["late8plus"] += 1

def top_mayor_lt(group_key_idx, min_count=3, top_n=8):
    g = defaultdict(list)
    for r in lt_rows:
        k = r[group_key_idx]
        if k:
            g[k].append(r)
    out = []
    for k, rs in g.items():
        if len(rs) < min_count:
            continue
        ar = sum(x[1] for x in rs) / len(rs)
        late_n = sum(1 for x in rs if x[4] is not None and x[4] > 0)
        pct_late = late_n / len(rs) * 100
        out.append([k, round(ar, 1), round(pct_late, 1), len(rs)])
    out.sort(key=lambda x: -x[1])
    return out[:top_n]

top_prov_lt = top_mayor_lt(5)
top_item_lt = top_mayor_lt(6)

LT_DATA = {
    "total": total, "avgReal": round(avg_real, 1), "avgProm": round(avg_prom, 1), "pctOnTime": round(pct_ontime, 1),
    "por2025": por2025, "por2026": por2026,
    "monthlyReal": monthlyReal, "monthlyProm": monthlyProm, "monthlyPctOn": monthlyPctOn, "monthlyCount": monthlyCount,
    "dist": dist,
    "topProvMayorLt": top_prov_lt,
    "topItemsMayorLt": top_item_lt,
}
print(f"  LT_DATA total={total} avgReal={avg_real:.1f} pctOnTime={pct_ontime:.1f}")

item_count_top6 = sorted(item_facturas_full.items(), key=lambda x: -x[1])[:6]
categoria_full = sorted(cant_by_refcat.items(), key=lambda x: -x[1])

# ════════════════════════════════════════════════════════════════
# PASS 2: Orden de compra por item.xlsx
#   -> D.fulfillByMonth/fulfillByYear, devoluciones (Anulado), DN.histOC,
#      DN.ocDetail
# ════════════════════════════════════════════════════════════════
print("PASS 2: Orden de compra...")
wb3 = openpyxl.load_workbook(base + "Orden de compra por item.xlsx", data_only=True, read_only=True)
ws3 = wb3.active

monthly_ord = defaultdict(float)
monthly_ent = defaultdict(float)
year_ord = defaultdict(float)
year_ent = defaultdict(float)

prov_row_count = defaultdict(int)   # all rows, by provider (for aprobadas_prov)
prov_set = set()

anulado_count = defaultdict(int)    # by provider
anulado_costo = defaultdict(float)  # by provider
anulado_items = defaultdict(int)    # by item
anulado_items_um = {}
anulado_porMes = defaultdict(int)   # by calendar month (1-12), collapsed across years
anulado_total = 0
anulado_monto = 0.0

um_cant_oc = defaultdict(float)
monthly_valor_neto_gr = defaultdict(float)
monthly_cant_ord_gr = defaultdict(float)

aprobadas_mes18 = [0]*len(months_all)
anuladas_mes18 = [0]*len(months_all)
aprobadas_yr = defaultdict(int)
anuladas_yr = defaultdict(int)

rows_for_table = []  # (year, month, prov, item, um, cant, precio, neto)

headers3 = None
n3 = 0
for i, row in enumerate(ws3.iter_rows(values_only=True)):
    if i == 0:
        headers3 = [str(c).strip() if c else '' for c in row]
        continue
    r = dict(zip(headers3, row))
    n3 += 1
    cant_ord = safe_float(r.get('Cant. ordenada'))
    cant_ent = safe_float(r.get('Cant. entrada'))
    valor = safe_float(r.get('Valor neto'))
    prov = str(r.get('Razón social proveedor') or r.get('Razon social proveedor') or '').strip()
    estado = str(r.get('Estado') or '').strip()
    item = str(r.get('Desc. item') or '').strip()
    um = str(r.get('U.M.') or '').strip()
    precio = safe_float(r.get('Precio unit.'))
    f_orden = safe_date(r.get('Fecha orden'))

    if prov:
        prov_row_count[prov] += 1
        prov_set.add(prov)

    if f_orden:
        key = f"{f_orden.year}-{f_orden.month:02d}"
        if key in month_idx:
            monthly_ord[key] += cant_ord
            monthly_ent[key] += cant_ent
            year_ord[f_orden.year] += cant_ord
            year_ent[f_orden.year] += cant_ent
            idx = month_idx[key]
            aprobadas_mes18[idx] += 1 if estado != 'Anulado' else 0
            anuladas_mes18[idx] += 1 if estado == 'Anulado' else 0
            aprobadas_yr[f_orden.year] += 1 if estado != 'Anulado' else 0
            anuladas_yr[f_orden.year] += 1 if estado == 'Anulado' else 0
            um_cant_oc[um] += cant_ord
            if um == 'GR':
                monthly_valor_neto_gr[key] += valor
                monthly_cant_ord_gr[key] += cant_ord
            rows_for_table.append((f_orden.year, key, prov, item, um, cant_ord, precio, valor))

    if estado == 'Anulado':
        anulado_total += 1
        anulado_monto += valor
        if prov:
            anulado_count[prov] += 1
            anulado_costo[prov] += valor
        if item:
            anulado_items[item] += 1
            anulado_items_um[item] = um
        if f_orden:
            anulado_porMes[f_orden.month] += 1

wb3.close()
print(f"  Filas OC: {n3}  Anuladas: {anulado_total}")

fulfillByMonth = []
for m in months_all:
    o = monthly_ord.get(m, 0)
    e = monthly_ent.get(m, 0)
    fulfillByMonth.append(round(e/o*100, 1) if o else 0.0)
fulfillByYear = {}
for y in (2025, 2026):
    o = year_ord.get(y, 0)
    e = year_ent.get(y, 0)
    fulfillByYear[str(y)] = round(e/o*100, 1) if o else 0.0

devoluciones_porSitio = sorted(anulado_count.items(), key=lambda x: -x[1])[:6]
devoluciones_costoProv = sorted(anulado_costo.items(), key=lambda x: -x[1])[:8]
devoluciones_items = sorted(anulado_items.items(), key=lambda x: -x[1])[:8]
MESES_ES = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"]
devoluciones_porMes = [[MESES_ES[mi-1], anulado_porMes.get(mi, 0)] for mi in range(1, 13)]

anuladas_prov_top6 = sorted(anulado_count.items(), key=lambda x: -x[1])[:6]
aprobadas_prov_top6 = sorted(prov_row_count.items(), key=lambda x: -x[1])[:6]

um_cant_oc_top = sorted(um_cant_oc.items(), key=lambda x: -x[1])[:6]

_MESES_LBL = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"]
costo_prom_lbl = [f"{_MESES_LBL[int(m[5:7])-1]}-{m[2:4]}" for m in months_all]
costo_prom_vals = []
for m in months_all:
    vneto = monthly_valor_neto_gr.get(m, 0)
    vcant = monthly_cant_ord_gr.get(m, 0)
    costo_prom_vals.append(round(vneto/vcant, 1) if vcant else 0.0)

rows_for_table.sort(key=lambda x: -x[7])
tabla = []
for (yr, mk, prov, item, um, cant, precio, neto) in rows_for_table[:15]:
    tabla.append({
        "a": str(yr), "m": mk, "p": abbrev_provider(prov), "it": item, "um": um,
        "cant": int(round(cant)), "pr": int(round(precio)), "neto": int(round(neto))
    })

n_providers_oc = len(prov_set)

wb_oc_stats = {
    "n3": n3, "anulado_total": anulado_total, "anulado_monto": anulado_monto,
    "n_providers_oc": n_providers_oc,
}

print("PASS 2 done.")

# ════════════════════════════════════════════════════════════════
# PASS 3: Salidas de inventario.xlsx -> DA.ajustes, DA.mermas
# ════════════════════════════════════════════════════════════════
print("PASS 3: Salidas de inventario...")
wb4 = openpyxl.load_workbook(base + "Salidas de inventario.xlsx", data_only=True, read_only=True)
ws4 = wb4.active

aj_items = set()
aj_bodegas = set()
aj_costo_total = 0.0
aj_count = 0
aj_costo_por_tienda = defaultdict(float)

me_count = 0
me_bodegas_origen = set()
me_bodegas_destino = set()
me_costo_item = defaultdict(float)
me_cant_item = defaultdict(float)
me_costo_destino = defaultdict(float)
me_hist2025 = [0]*12
me_hist2026 = [0]*12
me_detalle_demo = []

headers4 = None
for i, row in enumerate(ws4.iter_rows(values_only=True)):
    if i == 0:
        headers4 = [str(c).strip() if c else '' for c in row]
        continue
    r = dict(zip(headers4, row))
    fecha = safe_date(r.get('Fecha'))
    motivo = str(r.get('Desc. motivo') or '').strip().upper()
    bodega = str(r.get('Desc. bodega') or '').strip()
    co_movto = str(r.get('Desc. C.O. movto.') or '').strip()
    item = str(r.get('Desc. item') or '').strip()
    um = str(r.get('U.M. inv.') or '').strip()
    cant = safe_float(r.get('Salidas (inv.)'))
    costo = safe_float(r.get('Costo salidas (prom.)'))

    if motivo == 'AJUSTE DE INVENTARIO':
        aj_count += 1
        if item: aj_items.add(item)
        if bodega: aj_bodegas.add(bodega)
        aj_costo_total += costo
        if bodega:
            label = bodega[4:] if bodega.upper().startswith('PDV ') else bodega
            aj_costo_por_tienda[label] += costo

    if motivo == 'MERMA':
        me_count += 1
        if bodega: me_bodegas_origen.add(bodega)
        if co_movto: me_bodegas_destino.add(co_movto)
        if item:
            me_costo_item[item] += costo
            me_cant_item[item] += cant
        if co_movto:
            me_costo_destino[co_movto] += costo
        if fecha:
            if fecha.year == 2025 and 1 <= fecha.month <= 12:
                me_hist2025[fecha.month-1] += 1
            elif fecha.year == 2026 and 1 <= fecha.month <= 12:
                me_hist2026[fecha.month-1] += 1
        if len(me_detalle_demo) < 8 and fecha:
            me_detalle_demo.append({
                "f": fecha.strftime("%Y-%m-%d"), "item": item, "bo": bodega, "bd": co_movto,
                "um": um.strip(), "cant": int(round(cant)), "costo": int(round(costo)), "mot": "MERMA"
            })

wb4.close()
print(f"  Ajustes: {aj_count}  Mermas: {me_count}")

ajustes_porTiendaNeg = sorted(aj_costo_por_tienda.items(), key=lambda x: -x[1])[:8]
ajustes_porTiendaNeg = [[k, int(round(v))] for k, v in ajustes_porTiendaNeg]

mermas_costoItem = [[k, int(round(v))] for k, v in sorted(me_costo_item.items(), key=lambda x: -x[1])[:10]]
mermas_cantItem = [[k, int(round(v))] for k, v in sorted(me_cant_item.items(), key=lambda x: -x[1])[:10]]
mermas_bodegaDest = [[k, int(round(v))] for k, v in sorted(me_costo_destino.items(), key=lambda x: -x[1])[:8]]

DA = {
    "ajustes": {
        "itemsDiferentes": len(aj_items),
        "pvDiferentes": len(aj_bodegas),
        "sumaCostoEntradas": 0,
        "sumaCostoSalidas": int(round(aj_costo_total)),
        "porTiendaPos": [],
        "porTiendaNeg": ajustes_porTiendaNeg,
        "motivoCant": [["AJUSTE DE INVENTARIO", aj_count]],
        "tiposAjuste": [
            {"m": "AJUSTE DE INVENTARIO", "e": 0, "s": int(round(aj_costo_total)),
             "n": -aj_count, "t": -int(round(aj_costo_total))}
        ]
    },
    "devoluciones": {
        "total": anulado_total,
        "montoTotal": int(round(anulado_monto)),
        "motivos": [["OC Anuladas", anulado_total]],
        "porSitio": [[abbrev_provider(k), v] for k, v in devoluciones_porSitio],
        "costoProv": [[abbrev_provider(k), int(round(v))] for k, v in devoluciones_costoProv],
        "porMes": devoluciones_porMes,
        "items": [{"item": k, "um": anulado_items_um.get(k, ''), "cant": v} for k, v in devoluciones_items],
    },
    "mermas": {
        "totalMermas": me_count,
        "bodegasOrigen": len(me_bodegas_origen),
        "bodegasDestino": len(me_bodegas_destino),
        "costoItem": mermas_costoItem,
        "cantItem": mermas_cantItem,
        "bodegaDest": mermas_bodegaDest,
        "hist2025": me_hist2025,
        "hist2026": me_hist2026,
        "detalleDemo": me_detalle_demo,
    }
}

# ════════════════════════════════════════════════════════════════
# Load already-refreshed JSON outputs (dashboard_data.json / filter_data.json)
# to build D-extras and DN (reuse, do not re-derive what's already computed)
# ════════════════════════════════════════════════════════════════
with open(base + "dashboard_data.json", encoding="utf-8") as f:
    DD = json.load(f)

M = 1_000_000
G = 1_000_000_000

monthlySpend_m = {k: round(v/M, 1) for k, v in DD['monthly_spend'].items()}
hist_s2025 = [monthlySpend_m.get(f"2025-{mm:02d}", 0.0) for mm in range(1,13)]
hist_s2026 = [monthlySpend_m.get(f"2026-{mm:02d}", 0.0) for mm in range(1,13)]

monthlyOC_m = {k: round(v/M, 1) for k, v in DD['monthly_oc'].items()}
ocdetail_s2025 = [monthlyOC_m.get(f"2025-{mm:02d}", 0.0) for mm in range(1,13)]
ocdetail_s2026 = [monthlyOC_m.get(f"2026-{mm:02d}", 0.0) for mm in range(1,13)]

oc_estado_total = sum(DD['oc_por_estado'].values())
oc_anulado = DD['oc_por_estado'].get('Anulado', 0)

# prom3m: avg of the last 3 complete months of monthlyOC (Apr,May,Jun 2026) -- redefinition,
# see report notes (previous version was frozen on Oct-Dec 2025 window).
last3_keys = months_all[-3:]
prom3m = round(sum(monthlyOC_m.get(k, 0.0) for k in last3_keys) / 3, 2)

monto_prov_m_top6 = DD['top_suppliers_hist'][:6]
monto_prov_m_top6 = [[abbrev_provider(k), round(v/M, 2)] for k, v in monto_prov_m_top6]

categoria_top6 = [[k, int(round(v))] for k, v in categoria_full[:6]]
cant_um_hist_top7 = [[k, int(round(v))] for k, v in sorted(cant_by_um_hist.items(), key=lambda x: -x[1])[:7]]

# prov_var: % change last month vs previous month of monthly spend per supplier (Historico),
# top 10 by absolute magnitude, min base 5M COP to avoid noisy ratios -- judgment-call definition,
# see report notes.
prov_var = []
for prov, monthly in supplier_monthly_full.items():
    prev_v = monthly[-2]
    last_v = monthly[-1]
    if prev_v and abs(prev_v) >= 5_000_000:
        pct = (last_v - prev_v) / prev_v * 100
        prov_var.append((prov, pct))
prov_var.sort(key=lambda x: -abs(x[1]))
prov_var_top10 = [{"p": abbrev_provider(p), "v": round(v, 1)} for p, v in prov_var[:10]]

D_EXTRA = {
    "fulfillByMonth": fulfillByMonth,
    "fulfillByYear": fulfillByYear,
}

DN = {
    "histOC": {
        "kpi": {"total": oc_estado_total, "aprobadas": oc_estado_total - oc_anulado, "anuladas": oc_anulado,
                "monto_bn": round(DD['total_neto']/G, 2)},
        "anuladas_prov": [[abbrev_provider(k), v] for k, v in anuladas_prov_top6],
        "aprobadas_prov": [[abbrev_provider(k), v] for k, v in aprobadas_prov_top6],
        "cant_um": cant_um_hist_top7,
        "monto_prov_m": monto_prov_m_top6,
        "categoria": categoria_top6,
        "item_count": [[k, v] for k, v in item_count_top6],
        "hist_s2025": hist_s2025,
        "hist_s2026": hist_s2026,
        "aprobadas_mes": aprobadas_mes18,
        "anuladas_mes": anuladas_mes18,
        "aprobadas_yr": {"2025": aprobadas_yr.get(2025,0), "2026": aprobadas_yr.get(2026,0)},
        "anuladas_yr": {"2025": anuladas_yr.get(2025,0), "2026": anuladas_yr.get(2026,0)},
    },
    "ocDetail": {
        "kpi": {"monto_bn": round(DD['oc_valor_total']/G, 2), "cantidad": oc_estado_total,
                "proveedores": n_providers_oc, "prom3m": prom3m},
        "prov_var": prov_var_top10,
        "hist_s2025": ocdetail_s2025,
        "hist_s2026": ocdetail_s2026,
        "um_cant": [[k, int(round(v))] for k, v in um_cant_oc_top],
        "costo_prom_lbl": costo_prom_lbl,
        "costo_prom_vals": costo_prom_vals,
        "tabla": tabla,
    }
}

OUT = {"LT_DATA": LT_DATA, "DA": DA, "DN": DN, "D_EXTRA": D_EXTRA}

with open(base + "extract_lt_da_dn_output.json", "w", encoding="utf-8") as f:
    json.dump(OUT, f, ensure_ascii=False, indent=2)

print("\nGuardado extract_lt_da_dn_output.json")
print("fulfillByYear:", D_EXTRA["fulfillByYear"])
print("LT_DATA.total:", LT_DATA["total"], "avgReal:", LT_DATA["avgReal"], "pctOnTime:", LT_DATA["pctOnTime"])
print("DA.devoluciones.total:", DA["devoluciones"]["total"])
print("DA.ajustes.sumaCostoSalidas:", DA["ajustes"]["sumaCostoSalidas"])
print("DA.mermas.totalMermas:", DA["mermas"]["totalMermas"])
print("DN.histOC.kpi:", DN["histOC"]["kpi"])
print("DN.ocDetail.kpi:", DN["ocDetail"]["kpi"])
