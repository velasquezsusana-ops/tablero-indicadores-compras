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

with open(base + "filter_data.json", encoding="utf-8") as f:
    fd = json.load(f)
top_provs = set(s["name"] for s in fd["topSuppliers"])
top_items = set(it["name"] for it in fd["topItems"])

print("Leyendo historico de compras (lead time por mes/proveedor/item)...")
wb = openpyxl.load_workbook(base + "Histórico de compras por item.xlsx", data_only=True, read_only=True)
ws = wb.active

# Mismo criterio que LT_DATA original: solo filas con Docto. orden vinculado
lt_rows = []  # (f_oc, real, late_days, prov, item)

headers = None
n_total = 0
for i, row in enumerate(ws.iter_rows(values_only=True)):
    if i == 0:
        headers = [str(c).strip() if c else '' for c in row]
        continue
    r = dict(zip(headers, row))
    n_total += 1
    fecha = safe_date(r.get('Fecha'))
    prov = str(r.get('Razon social proveedor') or r.get('Razón social proveedor') or '').strip()
    item = str(r.get('Desc. item') or '').strip()

    docto_orden = r.get('Docto. orden')
    if docto_orden in (None, '', 'None'):
        continue
    f_oc = safe_date(r.get('Fecha documento OC'))
    f_prom = safe_date(r.get('Fecha de entrega item'))
    if not (fecha and f_oc):
        continue
    real = (fecha - f_oc).days
    if not (0 <= real <= 365):
        continue
    late_days = (fecha - f_prom).days if f_prom else None
    lt_rows.append((f_oc, real, late_days, prov, item))

wb.close()
print(f"  Filas con lead time válido: {len(lt_rows)}")

def bucket(ld):
    if ld is None: return None
    if ld <= 0: return 'ontime'
    if ld <= 3: return 'late1_3'
    if ld <= 7: return 'late4_7'
    return 'late8plus'

def empty_dist():
    return {"ontime": 0, "late1_3": 0, "late4_7": 0, "late8plus": 0}

dist_by_month = [empty_dist() for _ in months_all]
dist_by_prov = defaultdict(empty_dist)
dist_by_item = defaultdict(empty_dist)
# provMonthly[prov] = list of 18 [sumReal, count, lateCount]
prov_monthly = defaultdict(lambda: [[0.0, 0, 0] for _ in months_all])
prov_overall = defaultdict(lambda: [0.0, 0, 0])  # sumReal, count, lateCount

for f_oc, real, late_days, prov, item in lt_rows:
    key = f"{f_oc.year}-{f_oc.month:02d}"
    idx = month_idx.get(key)
    b = bucket(late_days)
    if idx is not None and b:
        dist_by_month[idx][b] += 1
    if prov in top_provs and b:
        dist_by_prov[prov][b] += 1
    if item in top_items and b:
        dist_by_item[item][b] += 1
    # provMonthly/provOverall cubren TODOS los proveedores (no solo top 30 por gasto)
    # para poder recalcular el top-8 por lead time dentro de un período filtrado.
    if prov:
        po = prov_overall[prov]
        po[0] += real; po[1] += 1
        if late_days is not None and late_days > 0: po[2] += 1
        if idx is not None:
            pm = prov_monthly[prov][idx]
            pm[0] += real; pm[1] += 1
            if late_days is not None and late_days > 0: pm[2] += 1

out = {
    "distByMonth": [{k: v for k, v in d.items()} for d in dist_by_month],
    "distByProv": {p: dict(d) for p, d in dist_by_prov.items()},
    "distByItem": {it: dict(d) for it, d in dist_by_item.items()},
    "provMonthly": {p: [[round(s, 1), c, l] for s, c, l in rows] for p, rows in prov_monthly.items()},
    "provOverall": {p: [round(s / c, 1) if c else 0, round(l / c * 100, 1) if c else 0, c] for p, (s, c, l) in prov_overall.items()},
}

with open(base + "lt_detail.json", "w", encoding="utf-8") as f:
    json.dump(out, f, ensure_ascii=False, separators=(',', ':'))

print("Guardado lt_detail.json")
