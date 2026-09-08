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

print("Leyendo ordenes de compra (detalle linea a linea de OC pendientes/vencidas)...")
wb = openpyxl.load_workbook(base + "Orden de compra por item.xlsx", data_only=True, read_only=True)
ws = wb.active

headers = None
n = 0
rows_out = []
for i, row in enumerate(ws.iter_rows(values_only=True)):
    if i == 0:
        headers = [str(c).strip() if c else '' for c in row]
        continue
    r = dict(zip(headers, row))
    n += 1
    cant_ord = safe_float(r.get('Cant. ordenada'))
    cant_pend = safe_float(r.get('Cant. pendiente'))
    if not (cant_pend > 0 and cant_ord > 0):
        continue
    dias_vto = safe_float(r.get('Dias vcto.'))
    valor_neto = safe_float(r.get('Valor neto'))
    valor_pend = valor_neto * (cant_pend / cant_ord)
    f_orden = safe_date(r.get('Fecha orden'))
    f_entrega = safe_date(r.get('Fecha entrega'))

    rows_out.append({
        "nro": str(r.get('Nro orden') or '').strip(),
        "prov": str(r.get('Razón social proveedor') or r.get('Razon social proveedor') or '').strip(),
        "item": str(r.get('Desc. item') or '').strip(),
        "bodega": str(r.get('Bodega') or '').strip(),
        "um": str(r.get('U.M.') or '').strip(),
        "cantPend": round(cant_pend, 2),
        "cantOrd": round(cant_ord, 2),
        "fechaOrden": f_orden.strftime('%Y-%m-%d') if f_orden else None,
        "fechaEntrega": f_entrega.strftime('%Y-%m-%d') if f_entrega else None,
        "diasVcto": int(round(dias_vto)),
        "valorPend": int(round(valor_pend)),
        "comprador": str(r.get('comprador') or '').strip(),
    })

wb.close()
print(f"  Filas OC: {n}  Pendientes (en transito): {len(rows_out)}  Vencidas: {sum(1 for x in rows_out if x['diasVcto'] > 0)}")

# Mas urgente primero: mas dias vencido, luego mayor valor pendiente
rows_out.sort(key=lambda x: (-x['diasVcto'], -x['valorPend']))

with open(base + "oc_pendientes_detalle.json", "w", encoding="utf-8") as f:
    json.dump(rows_out, f, ensure_ascii=False, separators=(',', ':'))

print("Guardado oc_pendientes_detalle.json")
