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

# Ventana fija del tablero (18 meses): igual que en el resto de los scripts de extract/.
# Sin este limite, meses parciales fuera de rango (ej. julio 2026 recien empezado) se
# cuelan como una clave extra en monthly_spend/monthly_oc/monthly_salidas y rompen la
# alineacion de 18 meses que usa todo el tablero (FILTER.months, BODEGA_MASTER, etc).
MESES_VALIDOS = {f"2025-{m:02d}" for m in range(1,13)} | {f"2026-{m:02d}" for m in range(1,10)}

# FILE 2 - Historico
print("Leyendo historico de compras...")
wb2 = openpyxl.load_workbook(base + "Histórico de compras por item.xlsx", data_only=True, read_only=True)
ws2 = wb2.active

monthly_spend = defaultdict(float)
supplier_spend = defaultdict(float)
item_spend = defaultdict(float)
bodega_spend = defaultdict(float)
total_bruto = 0.0
total_dsctos = 0.0
total_neto = 0.0
total_imptos = 0.0
n_compras = 0
estado_counts = defaultdict(int)

headers2 = None
for i, row in enumerate(ws2.iter_rows(values_only=True)):
    if i == 0:
        headers2 = [str(c).strip() if c else '' for c in row]
        continue
    r = dict(zip(headers2, row))
    fecha = safe_date(r.get('Fecha'))
    neto = safe_float(r.get('Valor neto local'))
    bruto = safe_float(r.get('Valor bruto local'))
    dscto = safe_float(r.get('Valor dsctos local'))
    impto = safe_float(r.get('Valor imptos local'))
    prov = str(r.get('Razon social proveedor') or r.get('Razón social proveedor') or '').strip()
    item = str(r.get('Desc. item') or '').strip()
    bodega = str(r.get('Bodega') or '').strip()
    estado = str(r.get('Estado') or '').strip()

    if fecha:
        key = f"{fecha.year}-{fecha.month:02d}"
        if key in MESES_VALIDOS:
            monthly_spend[key] += neto
    if prov: supplier_spend[prov] += neto
    if item: item_spend[item] += neto
    if bodega: bodega_spend[bodega] += neto
    total_bruto += bruto
    total_dsctos += dscto
    total_neto += neto
    total_imptos += impto
    n_compras += 1
    if estado: estado_counts[estado] += 1

wb2.close()
print(f"  Compras: {n_compras}")

# FILE 3 - Ordenes
print("Leyendo ordenes de compra...")
wb3 = openpyxl.load_workbook(base + "Orden de compra por item.xlsx", data_only=True, read_only=True)
ws3 = wb3.active

oc_total_ordenado = 0.0
oc_total_recibido = 0.0
oc_total_pendiente = 0.0
oc_valor_pendiente = 0.0
oc_valor_total = 0.0
comprador_counts = defaultdict(int)
comprador_valor = defaultdict(float)
prov_oc = defaultdict(float)
lead_times = []
oc_vencidas = 0
oc_por_estado = defaultdict(int)
monthly_oc = defaultdict(float)

headers3 = None
today = datetime(2026, 6, 11)
for i, row in enumerate(ws3.iter_rows(values_only=True)):
    if i == 0:
        headers3 = [str(c).strip() if c else '' for c in row]
        continue
    r = dict(zip(headers3, row))
    cant_ord = safe_float(r.get('Cant. ordenada'))
    cant_ent = safe_float(r.get('Cant. entrada'))
    cant_pend = safe_float(r.get('Cant. pendiente'))
    valor = safe_float(r.get('Valor neto'))
    dias_vto = safe_float(r.get('Dias vcto.'))
    comprador = str(r.get('comprador') or '').strip()
    prov = str(r.get('Razón social proveedor') or r.get('Razon social proveedor') or '').strip()
    estado = str(r.get('Estado') or '').strip()
    f_orden = safe_date(r.get('Fecha orden'))
    f_entrega = safe_date(r.get('Fecha entrega'))

    oc_total_ordenado += cant_ord
    oc_total_recibido += cant_ent
    oc_total_pendiente += cant_pend
    oc_valor_total += valor
    if cant_pend > 0 and cant_ord > 0:
        oc_valor_pendiente += valor * (cant_pend / cant_ord)

    if comprador:
        comprador_counts[comprador] += 1
        comprador_valor[comprador] += valor
    if prov: prov_oc[prov] += valor
    if estado: oc_por_estado[estado] += 1

    if f_orden and f_entrega:
        lt = (f_entrega - f_orden).days
        if 0 <= lt <= 365:
            lead_times.append(lt)

    if dias_vto > 0 and cant_pend > 0:
        oc_vencidas += 1

    if f_orden:
        key = f"{f_orden.year}-{f_orden.month:02d}"
        if key in MESES_VALIDOS:
            monthly_oc[key] += valor

wb3.close()
fulfillment_rate = (oc_total_recibido / oc_total_ordenado * 100) if oc_total_ordenado else 0
avg_lead_time = sum(lead_times) / len(lead_times) if lead_times else 0
print(f"  OC: {sum(oc_por_estado.values())}")

# FILE 4 - Salidas
print("Leyendo salidas de inventario...")
wb4 = openpyxl.load_workbook(base + "Salidas de inventario.xlsx", data_only=True, read_only=True)
ws4 = wb4.active

motivo_costo = defaultdict(float)
motivo_salidas_cant = defaultdict(float)
bodega_salidas = defaultdict(float)
item_salidas = defaultdict(float)
monthly_salidas = defaultdict(float)
total_costo_salidas = 0.0
n_salidas = 0

headers4 = None
for i, row in enumerate(ws4.iter_rows(values_only=True)):
    if i == 0:
        headers4 = [str(c).strip() if c else '' for c in row]
        continue
    r = dict(zip(headers4, row))
    fecha = safe_date(r.get('Fecha'))
    motivo = str(r.get('Desc. motivo') or '').strip()[:45]
    bodega = str(r.get('Desc. bodega') or '').strip()
    item = str(r.get('Desc. item') or '').strip()
    cant = safe_float(r.get('Salidas (inv.)'))
    costo = safe_float(r.get('Costo salidas (prom.)'))

    motivo_costo[motivo] += costo
    motivo_salidas_cant[motivo] += cant
    bodega_salidas[bodega] += costo
    if item: item_salidas[item] += costo
    total_costo_salidas += costo
    n_salidas += 1
    if fecha:
        key = f"{fecha.year}-{fecha.month:02d}"
        if key in MESES_VALIDOS:
            monthly_salidas[key] += costo

wb4.close()
print(f"  Salidas: {n_salidas}")

def top_n(d, n=10):
    return sorted(d.items(), key=lambda x: -x[1])[:n]

def sorted_monthly(d):
    return dict(sorted(d.items()))

data = {
    "total_neto": total_neto,
    "total_bruto": total_bruto,
    "total_dsctos": total_dsctos,
    "total_imptos": total_imptos,
    "n_compras": n_compras,
    "tasa_descuento": (total_dsctos / total_bruto * 100) if total_bruto else 0,
    "monthly_spend": sorted_monthly(monthly_spend),
    "monthly_oc": sorted_monthly(monthly_oc),
    "monthly_salidas": sorted_monthly(monthly_salidas),
    "top_suppliers_hist": top_n(supplier_spend, 10),
    "top_suppliers_oc": top_n(prov_oc, 10),
    "top_items_hist": top_n(item_spend, 10),
    "top_items_salidas": top_n(item_salidas, 10),
    "bodega_spend": top_n(bodega_spend, 10),
    "bodega_salidas": top_n(bodega_salidas, 10),
    "oc_total_ordenado": oc_total_ordenado,
    "oc_total_recibido": oc_total_recibido,
    "oc_total_pendiente": oc_total_pendiente,
    "oc_valor_pendiente": oc_valor_pendiente,
    "oc_valor_total": oc_valor_total,
    "fulfillment_rate": fulfillment_rate,
    "avg_lead_time": avg_lead_time,
    "oc_vencidas": oc_vencidas,
    "oc_por_estado": dict(oc_por_estado),
    "comprador_counts": top_n(comprador_counts, 10),
    "comprador_valor": top_n(comprador_valor, 10),
    "total_costo_salidas": total_costo_salidas,
    "motivo_costo": top_n(motivo_costo, 12),
    "motivo_salidas_cant": top_n(motivo_salidas_cant, 8),
    "estado_counts": dict(estado_counts),
}

with open(base + "dashboard_data.json", "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=2)

print("\n=== RESUMEN ===")
print(f"Gasto neto total: ${total_neto:,.0f}")
print(f"Descuento total:  ${total_dsctos:,.0f} ({data['tasa_descuento']:.2f}%)")
print(f"Fulfillment rate: {fulfillment_rate:.1f}%")
print(f"Lead time prom:   {avg_lead_time:.1f} dias")
print(f"Costo salidas:    ${total_costo_salidas:,.0f}")
print(f"OC pendientes $:  ${oc_valor_pendiente:,.0f}")
print("JSON guardado.")
