import openpyxl
import json
from collections import defaultdict

base = "C:/Users/Montolivo/Dropbox/Asesorias PYMES/1. Proyectos Actuales/2. MONTOLIVO/2. Proyecto/Indicadores Montolivo/Compras/Bases para actualizar Claude/"
sub = base + "OTIF Montolivo/"

# ════════════════════════════════════════════════════════════════
# Fuente: DATOS OTIFF MONTOLIVO (1).xlsx, hoja "BD REAL" (166 proveedores reales;
# se excluye a proposito la hoja "BD EJEMPLO", que es un ejemplo de otra empresa).
#
# Formula de calificacion (escala 1-5) tomada como referencia de
# "Ficha Proveedores Konkretus v7 (1).xlsm", hoja BENEFICIOS (formulas Excel reales):
#   Credito (dias):          >89->5, >30->4, >10->3, si no->1
#   Descuentos Especiales %: >9->5,  >5->4,  >2->3,  si no->1
#   Calidad/Precio:          EXCELENTE=5, BUENO=4, REGULAR=3, MALO=1
#   Promociones:              SI=5, "a veces"=3, NO=1
#   Campos SI/NO:              SI=5, NO=1
#
# Dos adaptaciones respecto a Konkretus (documentadas, no son copia literal):
#   - "Descuentos x Pronto pago": en Konkretus la formula es *100 (no calibra a 1-5,
#     luce como error de esa plantilla ya que rompe el promedio final). Aqui se le
#     aplica el mismo umbral que "Descuentos Especiales" para que sea comparable.
#   - "Tiempo de Entrega": en Konkretus es categorico (A TIEMPO/REGULAR/DEMORADA);
#     en Montolivo BD REAL es numerico (dias, mismo valor que LEAD TIME). Se usa un
#     umbral por dias en su lugar: <=1->5, <=3->4, <=7->3, <=15->2, si no->1.
# ════════════════════════════════════════════════════════════════

def score_credito(dias):
    if dias is None: return None
    if dias > 89: return 5
    if dias > 30: return 4
    if dias > 10: return 3
    return 1

def score_pct_umbral(pct):
    if pct is None: return None
    v = pct * 100 if pct <= 1 else pct  # tolera 0.06 o 6 indistintamente
    if v > 9: return 5
    if v > 5: return 4
    if v > 2: return 3
    return 1

def score_tiempo_entrega_dias(d):
    if d is None: return None
    if d <= 1: return 5
    if d <= 3: return 4
    if d <= 7: return 3
    if d <= 15: return 2
    return 1

def score_calidad_precio(v):
    v = str(v or '').strip().upper()
    return {'EXCELENTE': 5, 'BUENO': 4, 'REGULAR': 3, 'MALO': 1}.get(v)

def score_promociones(v):
    v = str(v or '').strip().upper()
    if v == 'SI': return 5
    if v == 'A VECES': return 3
    return 1

def score_sino(v):
    v = str(v or '').strip().upper()
    return 5 if v == 'SI' else 1

wb = openpyxl.load_workbook(sub + "DATOS OTIFF MONTOLIVO (1).xlsx", data_only=True, read_only=True)
ws = wb['BD REAL']

headers = None
rows = []
for i, row in enumerate(ws.iter_rows(values_only=True)):
    if i < 2:
        continue
    if i == 2:
        headers = [str(c).strip() if c else '' for c in row]
        continue
    r = dict(zip(headers, row))
    if not r.get('NOMBRE PROVEEDOR'):
        continue
    rows.append(r)
wb.close()
print(f"Proveedores leidos de BD REAL: {len(rows)}")

providers = []
for r in rows:
    nombre = str(r['NOMBRE PROVEEDOR']).strip()
    categoria = str(r.get('CATEGORIA') or '').strip()
    tipo = str(r.get('TIPO') or '').strip()

    admin = {
        'Crédito':                  {'um': 'DÍAS', 'resultado': r.get('Crédito'),                  'calif': score_credito(r.get('Crédito'))},
        'Descuentos x Pronto pago': {'um': '%',    'resultado': r.get('Descuentos x Pronto pago'),  'calif': score_pct_umbral(r.get('Descuentos x Pronto pago'))},
        'Descuentos Especiales':    {'um': '%',    'resultado': r.get('Descuentos Especiales'),     'calif': score_pct_umbral(r.get('Descuentos Especiales'))},
        'Tiempo de Entrega':        {'um': 'DÍAS', 'resultado': r.get('Tiempo de Entrega'),          'calif': score_tiempo_entrega_dias(r.get('Tiempo de Entrega'))},
    }
    producto = {
        'Calidad del Producto':               {'um': 'CALIDAD', 'resultado': r.get('Calidad del Producto'),               'calif': score_calidad_precio(r.get('Calidad del Producto'))},
        'Precio':                             {'um': 'NIVEL',   'resultado': r.get('Precio'),                             'calif': score_calidad_precio(r.get('Precio'))},
        'Promociones/Productos Agregados':    {'um': 'SI/NO',   'resultado': r.get('Promociones / Productos Agregados'),  'calif': score_promociones(r.get('Promociones / Productos Agregados'))},
    }
    servicio = {
        'Cambio mercancía dañada':  {'um': 'SI/NO', 'resultado': r.get('Cambio merc. Dañada'),  'calif': score_sino(r.get('Cambio merc. Dañada'))},
        'Cambio mercancía vencida': {'um': 'SI/NO', 'resultado': r.get('Cambio merc. vencida'), 'calif': score_sino(r.get('Cambio merc. vencida'))},
        'Entrega Puerta a Puerta':  {'um': 'SI/NO', 'resultado': r.get('Ent. Puerta a Puerta'),  'calif': score_sino(r.get('Ent. Puerta a Puerta'))},
        'Cuenta con SST y SGA':     {'um': 'SI/NO', 'resultado': r.get('Cuenta con SST y SGA'),  'calif': score_sino(r.get('Cuenta con SST y SGA'))},
    }

    all_calif = [v['calif'] for grp in (admin, producto, servicio) for v in grp.values() if v['calif'] is not None]
    total = round(sum(all_calif) / len(all_calif), 2) if all_calif else None

    providers.append({
        'nombre': nombre, 'categoria': categoria, 'tipo': tipo,
        'admin': admin, 'producto': producto, 'servicio': servicio,
        'total': total,
    })

# Promedio de la categoria (para comparacion vs. competidores), excluyendo al propio proveedor
by_cat = defaultdict(list)
for p in providers:
    if p['categoria']:
        by_cat[p['categoria']].append(p)

def campo_avg(lista, grp, campo, excl_nombre):
    vals = [p[grp][campo]['calif'] for p in lista if p['nombre'] != excl_nombre and p[grp][campo]['calif'] is not None]
    return round(sum(vals) / len(vals), 2) if vals else None

for p in providers:
    peers = by_cat.get(p['categoria'], [])
    p['proveedoresMismaCategoria'] = max(0, len(peers) - 1)
    p['peerAvg'] = {
        grp: {campo: campo_avg(peers, grp, campo, p['nombre']) for campo in p[grp]}
        for grp in ('admin', 'producto', 'servicio')
    }
    peer_totals = [q['total'] for q in peers if q['nombre'] != p['nombre'] and q['total'] is not None]
    p['peerTotalAvg'] = round(sum(peer_totals) / len(peer_totals), 2) if peer_totals else None

BENEFICIOS_DATA = {'providers': providers}

with open(base + "extract/beneficios_data.json", "w", encoding="utf-8") as f:
    json.dump(BENEFICIOS_DATA, f, ensure_ascii=False, separators=(',', ':'))

print(f"Guardado beneficios_data.json — {len(providers)} proveedores, {len(by_cat)} categorias")
