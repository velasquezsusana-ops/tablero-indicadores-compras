import openpyxl
import json
from collections import defaultdict
from datetime import datetime

def safe_float(v):
    try: return float(v) if v not in (None, 'None', '') else 0.0
    except: return 0.0

base = "C:/Users/Montolivo/Dropbox/Asesorias PYMES/1. Proyectos Actuales/2. MONTOLIVO/2. Proyecto/Indicadores Montolivo/Compras/Bases para actualizar Claude/"

months_all = ["2025-01","2025-02","2025-03","2025-04","2025-05","2025-06","2025-07","2025-08",
              "2025-09","2025-10","2025-11","2025-12","2026-01","2026-02","2026-03","2026-04","2026-05","2026-06","2026-07","2026-08","2026-09"]
month_idx = {m: i for i, m in enumerate(months_all)}
NM = len(months_all)

# ════════════════════════════════════════════════════════════════
# Ajustes de Inventario, desde "Salidas de inventario.xlsx" (Desc. motivo ==
# 'AJUSTE DE INVENTARIO'). Verificado: el 100% de estas filas son salidas
# (Entradas (inv.) siempre 0) -- no existe la contraparte "entrada" en esta
# fuente, ni las 5 categorias del mockup de referencia.
#
# Todo se guarda en arreglos de 18 meses (n, costoSalida, cantSalida) por
# bodega (clave = codigo PV1xx, igual que BODEGA_MASTER) y por item (clave =
# Desc. item, igual que FILTER.topItems), para que el tablero pueda sumar
# sobre los meses/tiendas/items activos segun los filtros de Año/Mes/Ítem/
# Bodega -- igual que OTIF_DATA/VENTAS_DATA.
# ════════════════════════════════════════════════════════════════

def new_bucket():
    return {"n": [0]*NM, "costoSalida": [0.0]*NM, "cantSalida": [0.0]*NM}

print("Leyendo Salidas de inventario.xlsx (motivo = AJUSTE DE INVENTARIO)...")
wb = openpyxl.load_workbook(base + "Salidas de inventario.xlsx", data_only=True, read_only=True)
ws = wb.active

overall = new_bucket()
byBodega = defaultdict(new_bucket)
byItem = defaultdict(new_bucket)
um_por_item = {}
items = set()
bodegas = set()

n = 0
n_fuera_periodo = 0
headers = None
for i, row in enumerate(ws.iter_rows(values_only=True)):
    if i == 0:
        headers = [str(c).strip() if c else '' for c in row]
        continue
    r = dict(zip(headers, row))
    if r.get('Desc. motivo') != 'AJUSTE DE INVENTARIO':
        continue
    n += 1

    fecha = r.get('Fecha')
    if not isinstance(fecha, datetime):
        continue
    key = f"{fecha.year}-{fecha.month:02d}"
    if key not in month_idx:
        n_fuera_periodo += 1
        continue
    mi = month_idx[key]

    bodega_cod = str(r.get('Bodega') or '').strip()
    item = str(r.get('Desc. item') or '').strip()
    um = str(r.get('U.M. inv.') or '').strip()
    cant_sal = safe_float(r.get('Salidas (inv.)'))
    costo_sal = safe_float(r.get('Costo salidas (prom.)'))

    targets = [overall]
    if bodega_cod:
        bodegas.add(bodega_cod)
        targets.append(byBodega[bodega_cod])
    if item:
        items.add(item)
        um_por_item[item] = um
        targets.append(byItem[item])

    for b in targets:
        b["n"][mi] += 1
        b["costoSalida"][mi] += costo_sal
        b["cantSalida"][mi] += cant_sal

wb.close()
print(f"  Filas AJUSTE DE INVENTARIO: {n}  (fuera de ventana Ene25-Jun26: {n_fuera_periodo})")
print(f"  Items distintos: {len(items)}  Bodegas distintas: {len(bodegas)}")


def round_bucket(b):
    return {k: [round(v, 4) if isinstance(v, float) else v for v in arr] for k, arr in b.items()}


AJUSTES_DATA = {
    "monthly": round_bucket(overall),
    "byBodega": {b: round_bucket(v) for b, v in byBodega.items()},
    "byItem": {it: round_bucket(v) for it, v in byItem.items()},
    "umPorItem": um_por_item,
}

tot = sum(overall["n"])
tot_costo = sum(overall["costoSalida"])
print(f"  Total movimientos (en ventana): {tot}  Costo salida total: {tot_costo:,.0f}")

with open(base + "extract/ajustes_data.json", "w", encoding="utf-8") as f:
    json.dump(AJUSTES_DATA, f, ensure_ascii=False, separators=(',', ':'))

print("\nGuardado ajustes_data.json")
