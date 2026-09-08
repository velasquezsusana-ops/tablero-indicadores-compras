import openpyxl
import json
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

print("Leyendo historico de compras (detalle linea a linea de lead time: fecha solicitada/esperada/llegada)...")
wb = openpyxl.load_workbook(base + "Histórico de compras por item.xlsx", data_only=True, read_only=True)
ws = wb.active

headers = None
n_total = 0
rows_out = []
for i, row in enumerate(ws.iter_rows(values_only=True)):
    if i == 0:
        headers = [str(c).strip() if c else '' for c in row]
        continue
    r = dict(zip(headers, row))
    n_total += 1

    docto_orden = r.get('Docto. orden')
    if docto_orden in (None, '', 'None'):
        continue

    fecha = safe_date(r.get('Fecha'))              # fecha real de llegada
    f_oc = safe_date(r.get('Fecha documento OC'))   # fecha solicitada (creacion OC)
    if not (fecha and f_oc):
        continue
    real = (fecha - f_oc).days
    if not (0 <= real <= 365):
        continue
    f_prom = safe_date(r.get('Fecha de entrega item'))  # fecha esperada (prometida)
    retraso = (fecha - f_prom).days if f_prom else None

    prov = str(r.get('Razon social proveedor') or r.get('Razón social proveedor') or '').strip()
    item = str(r.get('Desc. item') or '').strip()
    bodega = str(r.get('Bodega') or '').strip()
    um = str(r.get('U.M.') or '').strip()
    cant = safe_float(r.get('Cantidad'))

    rows_out.append({
        "nro": str(docto_orden).strip(),
        "prov": prov,
        "item": item,
        "bodega": bodega,
        "um": um,
        "cant": round(cant, 2),
        "fSol": f_oc.strftime('%Y-%m-%d'),
        "fEsp": f_prom.strftime('%Y-%m-%d') if f_prom else None,
        "fLleg": fecha.strftime('%Y-%m-%d'),
        "lt": real,
        "retraso": retraso,
    })

wb.close()
print(f"  Filas totales: {n_total}  Con lead time valido (docto. orden + fechas): {len(rows_out)}")

# Mas reciente primero
rows_out.sort(key=lambda x: x['fSol'], reverse=True)

with open(base + "lt_line_detail.json", "w", encoding="utf-8") as f:
    json.dump(rows_out, f, ensure_ascii=False, separators=(',', ':'))

print("Guardado lt_line_detail.json")
