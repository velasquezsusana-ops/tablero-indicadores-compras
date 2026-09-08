import openpyxl, json
from datetime import date, datetime
from common import to_date, mes_nombre

CUTOFF = date(2026, 9, 30)

wb = openpyxl.load_workbook("dif picking.xlsx", read_only=True, data_only=True)
ws = wb["Hoja1"]

def date_str(v):
    if v is None:
        return None
    if isinstance(v, str):
        return v.strip()
    d = to_date(v)
    return d.isoformat() if d else None

ENVIOS_DATA = []
maxdate = None
for r in ws.iter_rows(min_row=7, values_only=True):
    picking = r[1]
    pedido = r[2]
    doctoERP = r[3]
    eancontenido = r[4]
    referencia = r[5]
    descripcion = r[7]
    cant_pedida = r[8]
    qty_comprometida = r[9]
    can_picking = r[10]
    pendientes = r[11]
    fecha_picking = date_str(r[12])
    fecha_pedido = date_str(r[13])
    observacion = r[14]

    if doctoERP is None or fecha_pedido is None:
        continue
    d_pedido = date.fromisoformat(fecha_pedido)
    if d_pedido > CUTOFF:
        continue

    pdv = str(doctoERP).split("-")[0]
    anio = d_pedido.year
    mes = mes_nombre(d_pedido.month)

    ENVIOS_DATA.append([pdv, anio, mes, fecha_pedido, referencia, descripcion, cant_pedida, can_picking,
                         pendientes, qty_comprometida, picking, pedido, fecha_picking, observacion,
                         eancontenido, doctoERP])
    if maxdate is None or d_pedido > maxdate:
        maxdate = d_pedido

print("ENVIOS_DATA", len(ENVIOS_DATA), "max fechaPedido:", maxdate)
json.dump(ENVIOS_DATA, open("scripts/envios_extract.json", "w", encoding="utf-8"), ensure_ascii=False)
print("saved")
