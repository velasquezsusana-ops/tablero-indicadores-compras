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

months = ["2025-01","2025-02","2025-03","2025-04","2025-05","2025-06","2025-07","2025-08",
          "2025-09","2025-10","2025-11","2025-12","2026-01","2026-02","2026-03","2026-04","2026-05","2026-06","2026-07","2026-08","2026-09"]
month_idx = {m: i for i, m in enumerate(months)}

# ════════════════════════════════════════════════════════════════
# Bodegas top 10 (deben coincidir con D.bodegaSpend / dashboard_data.json)
# ════════════════════════════════════════════════════════════════
with open(base + "dashboard_data.json", encoding="utf-8") as f:
    dash = json.load(f)
bodegas_top = [b for b, _ in dash["bodega_spend"][:10]]
bodega_set = set(bodegas_top)

with open(base + "filter_data.json", encoding="utf-8") as f:
    fd = json.load(f)
top_provs = set(s["name"] for s in fd["topSuppliers"])
top_items = set(it["name"] for it in fd["topItems"])

print("Leyendo historico de compras (desglose bodega x mes/proveedor/item)...")
wb = openpyxl.load_workbook(base + "Histórico de compras por item.xlsx", data_only=True, read_only=True)
ws = wb.active

bodega_monthly = defaultdict(lambda: [0.0] * len(months))
bodega_by_prov = defaultdict(lambda: defaultdict(float))
bodega_by_item = defaultdict(lambda: defaultdict(float))

headers = None
n = 0
for i, row in enumerate(ws.iter_rows(values_only=True)):
    if i == 0:
        headers = [str(c).strip() if c else '' for c in row]
        continue
    r = dict(zip(headers, row))
    bodega = str(r.get('Bodega') or '').strip()
    if bodega not in bodega_set:
        continue
    neto = safe_float(r.get('Valor neto local'))
    fecha = safe_date(r.get('Fecha'))
    prov = str(r.get('Razón social proveedor') or '').strip()
    item = str(r.get('Desc. item') or '').strip()
    if fecha:
        key = f"{fecha.year}-{fecha.month:02d}"
        if key in month_idx:
            bodega_monthly[bodega][month_idx[key]] += neto
    if prov in top_provs:
        bodega_by_prov[prov][bodega] += neto
    if item in top_items:
        bodega_by_item[item][bodega] += neto
    n += 1

wb.close()
print(f"  Filas bodega top10: {n}")

M = 1_000_000
out = {
    "bodegas": bodegas_top,
    "monthly": {b: [round(v / M, 1) for v in bodega_monthly[b]] for b in bodegas_top},
    "byProv": {p: [round(bodega_by_prov[p].get(b, 0) / M, 1) for b in bodegas_top] for p in top_provs},
    "byItem": {it: [round(bodega_by_item[it].get(b, 0) / M, 1) for b in bodegas_top] for it in top_items},
}

with open(base + "bodega_detail.json", "w", encoding="utf-8") as f:
    json.dump(out, f, ensure_ascii=False, separators=(',', ':'))

print("Guardado bodega_detail.json")
