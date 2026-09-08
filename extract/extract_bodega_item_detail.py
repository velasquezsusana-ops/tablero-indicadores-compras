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

print("Leyendo historico de compras (detalle bodega x item, TODOS los items, con desglose mensual)...")
wb = openpyxl.load_workbook(base + "Histórico de compras por item.xlsx", data_only=True, read_only=True)
ws = wb.active

combo_gasto = defaultdict(float)
combo_fact = defaultdict(int)
combo_gasto_m = defaultdict(lambda: [0.0]*NM)
combo_fact_m = defaultdict(lambda: [0]*NM)
bodega_gasto = defaultdict(float)
bodega_fact = defaultdict(int)
bodega_gasto_m = defaultdict(lambda: [0.0]*NM)
bodega_fact_m = defaultdict(lambda: [0]*NM)

headers = None
n = 0
n_sin_fecha = 0
for i, row in enumerate(ws.iter_rows(values_only=True)):
    if i == 0:
        headers = [str(c).strip() if c else '' for c in row]
        continue
    r = dict(zip(headers, row))
    n += 1
    bodega = str(r.get('Bodega') or '').strip()
    item = str(r.get('Desc. item') or '').strip()
    neto = safe_float(r.get('Valor neto local'))
    if not (bodega and item):
        continue
    combo_gasto[(bodega, item)] += neto
    combo_fact[(bodega, item)] += 1
    bodega_gasto[bodega] += neto
    bodega_fact[bodega] += 1

    fecha = safe_date(r.get('Fecha'))
    if not fecha:
        n_sin_fecha += 1
        continue
    idx = month_idx.get(f"{fecha.year}-{fecha.month:02d}")
    if idx is None:
        continue
    combo_gasto_m[(bodega, item)][idx] += neto
    combo_fact_m[(bodega, item)][idx] += 1
    bodega_gasto_m[bodega][idx] += neto
    bodega_fact_m[bodega][idx] += 1

wb.close()
print(f"  Filas: {n}  Sin fecha (no cuentan en desglose mensual): {n_sin_fecha}  Bodegas: {len(bodega_gasto)}  Combos bodega-item: {len(combo_gasto)}")

M = 1_000_000

bodegaTotales = [
    [b, round(v / M, 1), bodega_fact[b],
     [round(x / M, 2) for x in bodega_gasto_m[b]], bodega_fact_m[b]]
    for b, v in sorted(bodega_gasto.items(), key=lambda x: -x[1]) if v > 0
]

bodegaItemDetail = [
    [b, it, round(v / M, 1), combo_fact[(b, it)],
     [round(x / M, 2) for x in combo_gasto_m[(b, it)]], combo_fact_m[(b, it)]]
    for (b, it), v in sorted(combo_gasto.items(), key=lambda x: (-bodega_gasto[x[0][0]], -x[1]))
    if v > 0
]

out = {
    "months": months_all,
    "bodegaTotales": bodegaTotales,    # [bodega, gastoM_total, facturas_total, gastoMensual[19], facturasMensual[19]]
    "bodegaItemDetail": bodegaItemDetail,  # [bodega, item, gastoM_total, facturas_total, gastoMensual[19], facturasMensual[19]]
}

with open(base + "bodega_item_detail.json", "w", encoding="utf-8") as f:
    json.dump(out, f, ensure_ascii=False, separators=(',', ':'))

print("Guardado bodega_item_detail.json")
print("Total gasto bodegas (M COP):", round(sum(x[1] for x in bodegaTotales), 1))
print("Total gasto combos (M COP):", round(sum(x[2] for x in bodegaItemDetail), 1))
print("Ejemplo bodega:", bodegaTotales[0][:3])
print("Ejemplo combo:", bodegaItemDetail[0][:4])
