import openpyxl
import json
from collections import defaultdict

def safe_float(v):
    try: return float(v) if v not in (None, 'None', '') else 0.0
    except: return 0.0

base = "C:/Users/Montolivo/Dropbox/Asesorias PYMES/1. Proyectos Actuales/2. MONTOLIVO/2. Proyecto/Indicadores Montolivo/Compras/Bases para actualizar Claude/"

# Unidades de peso inequívocas — se convierten directo sin importar el ítem.
CLEAN_GRAMOS = {'GR': 1.0, 'KG': 1000.0, 'LB': 453.592}

print("Leyendo hoja 'Und de medida' (homologación oficial por Referencia)...")
wb = openpyxl.load_workbook(base + "Histórico de compras por item.xlsx", data_only=True, read_only=True)
ws_um = wb['Und de medida']

# um_lookup: Referencia -> (U.M. inventario, factor de la U.M. de orden -> U.M. inventario, U.M. de orden)
# Ej: ACEITEDEOLIVAORUJO -> ('GR', 3000, 'BT')  => 1 BT de ese ítem = 3000 GR
um_lookup = {}
for i, r in enumerate(ws_um.iter_rows(values_only=True)):
    if i == 0:
        continue
    ref = str(r[0]).strip() if r[0] else ''
    if not ref or ref in um_lookup:
        continue
    um_invent = (r[2] or '').strip()
    factor = r[3]
    um_orden = str(r[4]).strip() if r[4] is not None else ''
    if isinstance(factor, (int, float)) and factor > 0:
        um_lookup[ref] = (um_invent, float(factor), um_orden)
print(f"  Ítems con homologación oficial: {len(um_lookup)}")

print("Leyendo historico de compras (agregados proveedor/item con homologación de peso)...")
ws = wb.active

def new_acc():
    return {"gasto": 0.0, "facturas": 0, "gastoGramos": 0.0, "gramos": 0.0, "gastoOtras": 0.0, "otrasUM": defaultdict(float)}

prov_acc = defaultdict(new_acc)
item_acc = defaultdict(new_acc)

headers = None
n = 0
n_exact = n_clean = n_sin_factor = n_unid = n_sin_ref = 0
for i, row in enumerate(ws.iter_rows(values_only=True)):
    if i == 0:
        headers = [str(c).strip() if c else '' for c in row]
        continue
    r = dict(zip(headers, row))
    n += 1
    prov = str(r.get('Razon social proveedor') or r.get('Razón social proveedor') or '').strip()
    item = str(r.get('Desc. item') or '').strip()
    ref = str(r.get('Referencia') or '').strip()
    neto = safe_float(r.get('Valor neto local'))
    cant = safe_float(r.get('Cantidad'))
    um = str(r.get('U.M.') or '').strip()

    info = um_lookup.get(ref)
    gramos = 0.0
    homologado = False
    if info:
        um_invent, factor_orden, um_orden = info
        if um_invent == 'GR':
            if um == um_orden:
                gramos = cant * factor_orden
                homologado = True
                n_exact += 1
            elif um in CLEAN_GRAMOS:
                gramos = cant * CLEAN_GRAMOS[um]
                homologado = True
                n_clean += 1
            else:
                n_sin_factor += 1
        else:
            n_unid += 1
    else:
        n_sin_ref += 1
        if um in CLEAN_GRAMOS:
            gramos = cant * CLEAN_GRAMOS[um]
            homologado = True

    for name, acc_dict in ((prov, prov_acc), (item, item_acc)):
        if not name:
            continue
        a = acc_dict[name]
        a["gasto"] += neto
        a["facturas"] += 1
        if homologado:
            a["gastoGramos"] += neto
            a["gramos"] += gramos
        else:
            a["gastoOtras"] += neto
            if um: a["otrasUM"][um] += cant

wb.close()
print(f"  Filas: {n}  Proveedores: {len(prov_acc)}  Items: {len(item_acc)}")
print(f"  Homologadas exacto (Referencia+U.M.orden): {n_exact}")
print(f"  Homologadas por unidad de peso limpia (GR/KG/LB): {n_clean}")
print(f"  Item de peso pero sin factor para esa U.M. puntual: {n_sin_factor}")
print(f"  Item fundamentalmente por unidad/conteo (no homologable): {n_unid}")
print(f"  Sin referencia en hoja 'Und de medida': {n_sin_ref}")

M = 1_000_000

def build_out(acc, order):
    out = []
    for name in order:
        a = acc.get(name)
        if not a or a["gasto"] <= 0:
            continue
        pctPeso = round(a["gastoGramos"] / a["gasto"] * 100, 1) if a["gasto"] else 0.0
        out.append({
            "name": name,
            "gasto": round(a["gasto"] / M, 1),
            "facturas": a["facturas"],
            "ticketProm": round(a["gasto"] / a["facturas"] / 1000, 0) if a["facturas"] else 0,  # miles COP
            "gramos": round(a["gramos"], 1),
            "kg": round(a["gramos"] / 1000, 2),
            "pctPesoHomologado": pctPeso,
            "precioPromGramo": round(a["gastoGramos"] / a["gramos"], 2) if a["gramos"] > 0 else None,
            "otrasUM": {k: round(v, 1) for k, v in sorted(a["otrasUM"].items(), key=lambda x: -x[1])[:3]},
        })
    out.sort(key=lambda x: -x["gasto"])
    return out

todos_provs = sorted(prov_acc.keys(), key=lambda p: -prov_acc[p]["gasto"])
todos_items = sorted(item_acc.keys(), key=lambda it: -item_acc[it]["gasto"])

out = {
    "porProveedor": build_out(prov_acc, todos_provs),
    "porItem": build_out(item_acc, todos_items),
}

with open(base + "historico_prov_item.json", "w", encoding="utf-8") as f:
    json.dump(out, f, ensure_ascii=False, separators=(',', ':'))

print("Guardado historico_prov_item.json")
print("Total gasto proveedores (M COP):", round(sum(x["gasto"] for x in out["porProveedor"]), 1))
print("Total gasto items (M COP):", round(sum(x["gasto"] for x in out["porItem"]), 1))
gasto_total = sum(x["gasto"] for x in out["porItem"])
gasto_homolog = sum(x["gasto"]*x["pctPesoHomologado"]/100 for x in out["porItem"])
print("Cobertura global homologación (items):", round(gasto_homolog/gasto_total*100, 1), "%")
print("Ejemplo proveedor:", out["porProveedor"][0])
print("Ejemplo item:", out["porItem"][0])
