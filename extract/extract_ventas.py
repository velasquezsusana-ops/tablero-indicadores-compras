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
M = 1_000_000  # COP -> Millones, misma convencion que D.monthlySpend

print("Leyendo ventas 2025-2026.xlsx (hoja 'ventas')...")
wb = openpyxl.load_workbook(base + "ventas 2025-2026.xlsx", data_only=True, read_only=True)
ws = wb['ventas']

monthly = [0.0] * NM
by_bodega = defaultdict(lambda: [0.0] * NM)

headers = None
n = 0
n_used = 0
for i, row in enumerate(ws.iter_rows(values_only=True)):
    if i == 0:
        headers = [str(c).strip() if c else '' for c in row]
        continue
    r = dict(zip(headers, row))
    n += 1
    fecha = r.get('Fecha')
    if not isinstance(fecha, datetime):
        continue
    key = f"{fecha.year}-{fecha.month:02d}"
    if key not in month_idx:
        continue  # fuera de la ventana Ene 2025 - Jun 2026 del tablero
    mi = month_idx[key]
    valor = safe_float(r.get('Valor subtotal'))
    bodega = str(r.get('Bodega') or '').strip()

    monthly[mi] += valor
    if bodega:
        by_bodega[bodega][mi] += valor
    n_used += 1

wb.close()
print(f"  Filas leidas: {n}  Usadas (dentro de la ventana Ene25-Jun26): {n_used}")
print(f"  Bodegas distintas: {len(by_bodega)}")

VENTAS_DATA = {
    "monthly": [round(v / M, 2) for v in monthly],
    "byBodega": {b: [round(v / M, 2) for v in arr] for b, arr in by_bodega.items()},
}

total_ventas = sum(monthly)
print(f"  Ventas totales (Ene25-Jun26): ${total_ventas:,.0f}  (~${total_ventas/M:,.1f}M)")

with open(base + "extract/ventas_data.json", "w", encoding="utf-8") as f:
    json.dump(VENTAS_DATA, f, ensure_ascii=False, separators=(',', ':'))

print("Guardado ventas_data.json")
