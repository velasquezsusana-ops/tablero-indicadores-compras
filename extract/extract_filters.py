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

print("Leyendo historico de compras (detalle proveedor/item por mes)...")
wb2 = openpyxl.load_workbook(base + "Histórico de compras por item.xlsx", data_only=True, read_only=True)
ws2 = wb2.active

months = ["2025-01","2025-02","2025-03","2025-04","2025-05","2025-06","2025-07","2025-08",
          "2025-09","2025-10","2025-11","2025-12","2026-01","2026-02","2026-03","2026-04","2026-05","2026-06","2026-07","2026-08","2026-09"]
month_idx = {m:i for i,m in enumerate(months)}

supplier_spend = defaultdict(float)
item_spend = defaultdict(float)
supplier_monthly = defaultdict(lambda: [0.0]*len(months))
item_monthly = defaultdict(lambda: [0.0]*len(months))
supplier_facturas = defaultdict(int)
item_facturas = defaultdict(int)

headers2 = None
n = 0
for i, row in enumerate(ws2.iter_rows(values_only=True)):
    if i == 0:
        headers2 = [str(c).strip() if c else '' for c in row]
        continue
    r = dict(zip(headers2, row))
    fecha = safe_date(r.get('Fecha'))
    neto = safe_float(r.get('Valor neto local'))
    prov = str(r.get('Razon social proveedor') or r.get('Razón social proveedor') or '').strip()
    item = str(r.get('Desc. item') or '').strip()

    if prov:
        supplier_spend[prov] += neto
        supplier_facturas[prov] += 1
    if item:
        item_spend[item] += neto
        item_facturas[item] += 1
    if fecha:
        key = f"{fecha.year}-{fecha.month:02d}"
        if key in month_idx:
            idx = month_idx[key]
            if prov: supplier_monthly[prov][idx] += neto
            if item: item_monthly[item][idx] += neto
    n += 1

wb2.close()
print(f"  Filas procesadas: {n}")
print(f"  Proveedores distintos: {len(supplier_spend)}")
print(f"  Items distintos: {len(item_spend)}")

TOP_N = 100000  # sin limite real: se incluyen todos los proveedores/items (215/636 en la base actual)

def top_n_names(d, n=TOP_N):
    return [k for k,_ in sorted(d.items(), key=lambda x: -x[1])[:n]]

top_suppliers = top_n_names(supplier_spend)
top_items = top_n_names(item_spend)

M = 1_000_000  # convertir a millones COP para coherencia con el resto del tablero

out = {
    "months": months,
    "totalProveedores": len(supplier_spend),
    "totalItems": len(item_spend),
    "topSuppliers": [
        {"name": s, "total": round(supplier_spend[s]/M,1), "facturas": supplier_facturas[s],
         "monthly": [round(v/M,1) for v in supplier_monthly[s]]}
        for s in top_suppliers
    ],
    "topItems": [
        {"name": it, "total": round(item_spend[it]/M,1), "facturas": item_facturas[it],
         "monthly": [round(v/M,1) for v in item_monthly[it]]}
        for it in top_items
    ],
}

with open(base + "filter_data.json", "w", encoding="utf-8") as f:
    json.dump(out, f, ensure_ascii=False, indent=2)

print("Guardado filter_data.json")
print("Top 5 proveedores:", top_suppliers[:5])
print("Top 5 items:", top_items[:5])
