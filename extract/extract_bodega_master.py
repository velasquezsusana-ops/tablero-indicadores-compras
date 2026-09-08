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
month_idx = {m: i for i, m in enumerate(months_all)}
NM = len(months_all)

BODEGA_NAMES = {
    "CDP97": "CDP CENTRO DE PRODUCCION", "DOT01": "DOTACION USADA", "DOTAC": "DOTACION",
    "DV101": "DEVOLUCION PRINCIPAL", "FD099": "FINCA DIRECTA", "LOCAL": "LOCAL",
    "MONTE": "BODEGA MONTERREY", "PTERC": "OLIVENZA", "PV102": "PDV AVENIDA POBLADO",
    "PV103": "PDV SAN DIEGO", "PV104": "PDV PREMIUM", "PV105": "PDV MAYORCA ET 2",
    "PV106": "PDV OVIEDO", "PV107": "PDV VEGAS PLAZA", "PV108": "PDV SANTAFE",
    "PV109": "PDV UNICENTRO MEDELLIN", "PV110": "PDV AVENTURA", "PV111": "PDV GRAN PLAZA",
    "PV112": "PDV FLORIDA", "PV114": "PDV MAYORCA ET 1", "PV115": "PDV PUERTA DEL NORTE",
    "PV116": "PDV JARDIN PLAZA", "PV117": "PDV PALMETTO PLAZA", "PV118": "PDV GRAN ESTACION",
    "PV119": "PDV VIVA ENVIGADO", "PV120": "PDV DE MODA", "PV121": "PDV VIVA LAURELES",
    "PV122": "PDV TITAN PLAZA", "PV124": "PDV LOS MOLINOS", "PV125": "PDV SAN NICOLAS",
    "PV127": "PDV SANTAFE BOGOTA", "PV128": "PDV CENTRO MAYOR", "PV129": "PDV ARKADIA",
    "PV131": "PDV PARQUE ARBOLEDA", "PV132": "PDV EL TESORO", "PV133": "PDV FUNDADORES",
    "PV134": "PDV UNICO OUTLET II", "PV135": "PDV LA CENTRAL", "PV136": "PDV PARQUE FABRICATO",
    "PV137": "PDV MALL PLAZA", "PV138": "PDV VIVA BARRANQUILLA", "PV139": "PDV LA COLINA",
    "PV141": "PDV DISTRITO VERA", "PV142": "MALL PLAZA CALI", "PV143": "FLORIDA ET 2",
    "PV144": "BUENAVISTA", "PV146": "UNICENTRO BOGOTA", "PV147": "UNICENTRO CALI",
    "PV148": "PLAZA DE LAS AMERICAS", "PV149": "PLAZA CENTRAL", "PV151": "VICTORIA",
    "RP001": "REPUESTOS", "PV102": "PDV AVENIDA POBLADO",
}

with open(base + "filter_data.json", encoding="utf-8") as f:
    fd = json.load(f)
top_provs = set(s["name"] for s in fd["topSuppliers"])
top_items = set(it["name"] for it in fd["topItems"])

# ════════════════════════════════════════════════════════════════
# PASS 1: Historico de compras por item.xlsx -> spend by bodega x month x (prov/item)
# ════════════════════════════════════════════════════════════════
print("PASS 1: Historico de compras (bodega x mes x prov/item)...")
wb2 = openpyxl.load_workbook(base + "Histórico de compras por item.xlsx", data_only=True, read_only=True)
ws2 = wb2.active

all_bodegas = set()
hist_monthly = defaultdict(lambda: [0.0]*NM)
hist_facturas = defaultdict(lambda: [0]*NM)
hist_prov_monthly = defaultdict(lambda: defaultdict(lambda: [0.0]*NM))
hist_item_monthly = defaultdict(lambda: defaultdict(lambda: [0.0]*NM))

# Lead time por bodega x mes (mismo criterio que extract_lt_da_dn.py: requiere Docto. orden vinculado)
# on_time (igual que LT_DATA.pctOnTime) = f_prom existe Y late_days <= 0. late_days=None (sin fecha
# prometida) NO cuenta como on-time.
lt_sum_by_bodega_month = defaultdict(lambda: [[0.0,0,0,0] for _ in range(NM)])  # [sumReal, count, lateCount, onTimeCount]
def empty_dist(): return {"ontime":0,"late1_3":0,"late4_7":0,"late8plus":0}
lt_dist_by_bodega_month = defaultdict(lambda: [empty_dist() for _ in range(NM)])
def bucket(ld):
    if ld is None: return None
    if ld <= 0: return 'ontime'
    if ld <= 3: return 'late1_3'
    if ld <= 7: return 'late4_7'
    return 'late8plus'

headers2 = None
n2 = 0
for i, row in enumerate(ws2.iter_rows(values_only=True)):
    if i == 0:
        headers2 = [str(c).strip() if c else '' for c in row]
        continue
    r = dict(zip(headers2, row))
    n2 += 1
    bodega = str(r.get('Bodega') or '').strip()
    if not bodega:
        continue
    all_bodegas.add(bodega)
    fecha = safe_date(r.get('Fecha'))
    neto = safe_float(r.get('Valor neto local'))
    prov = str(r.get('Razon social proveedor') or r.get('Razón social proveedor') or '').strip()
    item = str(r.get('Desc. item') or '').strip()

    docto_orden = r.get('Docto. orden')
    if docto_orden not in (None, '', 'None'):
        f_oc = safe_date(r.get('Fecha documento OC'))
        f_prom = safe_date(r.get('Fecha de entrega item'))
        if fecha and f_oc:
            real = (fecha - f_oc).days
            if 0 <= real <= 365:
                late_days = (fecha - f_prom).days if f_prom else None
                key_oc = f"{f_oc.year}-{f_oc.month:02d}"
                idx_oc = month_idx.get(key_oc)
                if idx_oc is not None:
                    row_acc = lt_sum_by_bodega_month[bodega][idx_oc]
                    row_acc[0] += real; row_acc[1] += 1
                    if late_days is not None and late_days > 0: row_acc[2] += 1
                    if f_prom is not None and fecha <= f_prom: row_acc[3] += 1
                    b = bucket(late_days)
                    if b: lt_dist_by_bodega_month[bodega][idx_oc][b] += 1

    if not fecha:
        continue
    key = f"{fecha.year}-{fecha.month:02d}"
    idx = month_idx.get(key)
    if idx is None:
        continue
    hist_monthly[bodega][idx] += neto
    hist_facturas[bodega][idx] += 1
    if prov in top_provs:
        hist_prov_monthly[bodega][prov][idx] += neto
    if item in top_items:
        hist_item_monthly[bodega][item][idx] += neto

wb2.close()
print(f"  Filas: {n2}  Bodegas distintas: {len(all_bodegas)}")

# ════════════════════════════════════════════════════════════════
# PASS 2: Orden de compra por item.xlsx -> fulfillment/vencidas/pendiente by bodega x mes x (prov/item)
# ════════════════════════════════════════════════════════════════
print("PASS 2: Orden de compra (bodega x mes x prov/item)...")
wb3 = openpyxl.load_workbook(base + "Orden de compra por item.xlsx", data_only=True, read_only=True)
ws3 = wb3.active

oc_ord_monthly = defaultdict(lambda: [0.0]*NM)
oc_ent_monthly = defaultdict(lambda: [0.0]*NM)
oc_venc_monthly = defaultdict(lambda: [0]*NM)
oc_venc_prov_monthly = defaultdict(lambda: defaultdict(lambda: [0]*NM))
oc_venc_item_monthly = defaultdict(lambda: defaultdict(lambda: [0]*NM))
oc_pend_monthly = defaultdict(lambda: [0.0]*NM)
oc_pend_prov_monthly = defaultdict(lambda: defaultdict(lambda: [0.0]*NM))
oc_pend_item_monthly = defaultdict(lambda: defaultdict(lambda: [0.0]*NM))
oc_bodegas = set()

headers3 = None
n3 = 0
for i, row in enumerate(ws3.iter_rows(values_only=True)):
    if i == 0:
        headers3 = [str(c).strip() if c else '' for c in row]
        continue
    r = dict(zip(headers3, row))
    n3 += 1
    bodega = str(r.get('Bodega') or '').strip()
    if not bodega:
        continue
    f_orden = safe_date(r.get('Fecha orden'))
    if not f_orden:
        continue
    key = f"{f_orden.year}-{f_orden.month:02d}"
    idx = month_idx.get(key)
    if idx is None:
        continue

    oc_bodegas.add(bodega)
    all_bodegas.add(bodega)
    cant_ord = safe_float(r.get('Cant. ordenada'))
    cant_ent = safe_float(r.get('Cant. entrada'))
    cant_pend = safe_float(r.get('Cant. pendiente'))
    valor = safe_float(r.get('Valor neto'))
    dias_vto = safe_float(r.get('Dias vcto.'))
    prov = str(r.get('Razón social proveedor') or r.get('Razon social proveedor') or '').strip()
    item = str(r.get('Desc. item') or '').strip()

    oc_ord_monthly[bodega][idx] += cant_ord
    oc_ent_monthly[bodega][idx] += cant_ent

    if dias_vto > 0 and cant_pend > 0:
        oc_venc_monthly[bodega][idx] += 1
        if prov in top_provs:
            oc_venc_prov_monthly[bodega][prov][idx] += 1
        if item in top_items:
            oc_venc_item_monthly[bodega][item][idx] += 1

    if cant_pend > 0 and cant_ord > 0:
        valor_pend = valor * (cant_pend / cant_ord)
        oc_pend_monthly[bodega][idx] += valor_pend
        if prov in top_provs:
            oc_pend_prov_monthly[bodega][prov][idx] += valor_pend
        if item in top_items:
            oc_pend_item_monthly[bodega][item][idx] += valor_pend

wb3.close()
print(f"  Filas: {n3}  Bodegas OC distintas: {len(oc_bodegas)} -> {sorted(oc_bodegas)}")

# ════════════════════════════════════════════════════════════════
# Ensamblar salida
# ════════════════════════════════════════════════════════════════
M = 1_000_000
bodega_totals = {b: sum(v) for b, v in hist_monthly.items()}
bodegas_sorted = sorted(all_bodegas, key=lambda b: -bodega_totals.get(b, 0))

out = {
    "bodegas": bodegas_sorted,
    "bodegaNames": {b: BODEGA_NAMES.get(b, b) for b in bodegas_sorted},
    "ocBodegas": sorted(oc_bodegas),
    "hist": {
        "monthly": {b: [round(v/M,1) for v in hist_monthly[b]] for b in hist_monthly},
        "facturasMonthly": {b: hist_facturas[b] for b in hist_facturas},
        "provMonthly": {b: {p: [round(v/M,1) for v in arr] for p,arr in provs.items()} for b,provs in hist_prov_monthly.items()},
        "itemMonthly": {b: {it: [round(v/M,1) for v in arr] for it,arr in items.items()} for b,items in hist_item_monthly.items()},
    },
    "lt": {
        "sumByMonth": {b: [[round(s,1),c,l,o] for s,c,l,o in rows] for b,rows in lt_sum_by_bodega_month.items()},
        "distByMonth": {b: [dict(d) for d in rows] for b,rows in lt_dist_by_bodega_month.items()},
    },
    "oc": {
        "ordMonthly": {b: oc_ord_monthly[b] for b in oc_ord_monthly},
        "entMonthly": {b: oc_ent_monthly[b] for b in oc_ent_monthly},
        "vencMonthly": {b: oc_venc_monthly[b] for b in oc_venc_monthly},
        "vencProvMonthly": {b: {p: arr for p,arr in provs.items()} for b,provs in oc_venc_prov_monthly.items()},
        "vencItemMonthly": {b: {it: arr for it,arr in items.items()} for b,items in oc_venc_item_monthly.items()},
        "pendMonthly": {b: [round(v/M,1) for v in oc_pend_monthly[b]] for b in oc_pend_monthly},
        "pendProvMonthly": {b: {p: [round(v/M,1) for v in arr] for p,arr in provs.items()} for b,provs in oc_pend_prov_monthly.items()},
        "pendItemMonthly": {b: {it: [round(v/M,1) for v in arr] for it,arr in items.items()} for b,items in oc_pend_item_monthly.items()},
    },
}

with open(base + "bodega_master.json", "w", encoding="utf-8") as f:
    json.dump(out, f, ensure_ascii=False, separators=(',', ':'))

print("Guardado bodega_master.json")
print("Total bodegas:", len(bodegas_sorted))
print("Check total gasto (suma bodegas, M COP):", round(sum(bodega_totals.values())/M, 1))
print("Check total vencidas (suma bodegas):", sum(sum(v) for v in oc_venc_monthly.values()))
print("Check total pendiente (suma bodegas, M COP):", round(sum(sum(v) for v in oc_pend_monthly.values()), 1))
