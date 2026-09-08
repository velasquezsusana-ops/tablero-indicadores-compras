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

print("Leyendo ordenes de compra (OC vencidas por mes/proveedor/item)...")
wb = openpyxl.load_workbook(base + "Orden de compra por item.xlsx", data_only=True, read_only=True)
ws = wb.active

monthly = [0] * len(months_all)
yearly = defaultdict(int)
prov_monthly = defaultdict(lambda: [0] * len(months_all))
prov_overall = defaultdict(int)
item_monthly = defaultdict(lambda: [0] * len(months_all))
item_overall = defaultdict(int)

headers = None
n = 0
n_venc = 0
for i, row in enumerate(ws.iter_rows(values_only=True)):
    if i == 0:
        headers = [str(c).strip() if c else '' for c in row]
        continue
    r = dict(zip(headers, row))
    n += 1
    cant_pend = safe_float(r.get('Cant. pendiente'))
    dias_vto = safe_float(r.get('Dias vcto.'))
    if not (dias_vto > 0 and cant_pend > 0):
        continue
    n_venc += 1
    prov = str(r.get('Razón social proveedor') or r.get('Razon social proveedor') or '').strip()
    item = str(r.get('Desc. item') or '').strip()
    f_orden = safe_date(r.get('Fecha orden'))

    if prov in top_provs:
        prov_overall[prov] += 1
    if item in top_items:
        item_overall[item] += 1

    if f_orden:
        key = f"{f_orden.year}-{f_orden.month:02d}"
        yearly[f_orden.year] += 1
        idx = month_idx.get(key)
        if idx is not None:
            monthly[idx] += 1
            if prov in top_provs:
                prov_monthly[prov][idx] += 1
            if item in top_items:
                item_monthly[item][idx] += 1

wb.close()
print(f"  Filas OC: {n}  Vencidas: {n_venc}")

out = {
    "monthly": monthly,
    "yearly": {str(y): c for y, c in yearly.items()},
    "provMonthly": {p: v for p, v in prov_monthly.items()},
    "provOverall": {p: c for p, c in prov_overall.items()},
    "itemMonthly": {it: v for it, v in item_monthly.items()},
    "itemOverall": {it: c for it, c in item_overall.items()},
}

with open(base + "ocvenc_detail.json", "w", encoding="utf-8") as f:
    json.dump(out, f, ensure_ascii=False, separators=(',', ':'))

print("Guardado ocvenc_detail.json")
print("Total vencidas (monthly sum):", sum(monthly))
print("Total vencidas (yearly sum):", sum(yearly.values()))
