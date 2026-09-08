import json

base = "C:/Users/Montolivo/Dropbox/Asesorias PYMES/1. Proyectos Actuales/2. MONTOLIVO/2. Proyecto/Indicadores Montolivo/Compras/Bases para actualizar Claude/"
scratch = "C:/Users/MONTOL~1/AppData/Local/Temp/claude/C--Users-Montolivo/27afce9d-8783-40f9-967e-828b75f5dde0/scratchpad/"

with open(base + "dashboard_data.json", encoding="utf-8") as f:
    DD = json.load(f)
with open(base + "filter_data.json", encoding="utf-8") as f:
    FILTER_RAW = json.load(f)
with open(base + "filter_data_v2.json", encoding="utf-8") as f:
    PRICE_RAW = json.load(f)
with open(base + "extract_lt_da_dn_output.json", encoding="utf-8") as f:
    EXTRA = json.load(f)

M = 1_000_000

def js_arr_tuples(pairs, vround=1):
    # [["name", val], ...] -> JS array-of-tuples text
    parts = []
    for k, v in pairs:
        if isinstance(v, float):
            v = round(v, vround)
        parts.append(f'["{js_str(k)}",{v}]')
    return "[" + ",".join(parts) + "]"

def js_str(s):
    return str(s).replace('\\', '\\\\').replace('"', '\\"')

def fmt_num(v):
    if isinstance(v, float):
        if v == int(v):
            return f"{v:.1f}"
        return repr(round(v, 4)).rstrip('0').rstrip('.') if False else str(round(v,4))
    return str(v)

# ── D ──
monthly_spend_m = {k: round(v/M, 1) for k, v in sorted(DD['monthly_spend'].items())}
monthly_oc_m = {k: round(v/M, 1) for k, v in sorted(DD['monthly_oc'].items())}
monthly_salidas_m = {k: round(v/M, 1) for k, v in sorted(DD['monthly_salidas'].items())}

def obj_month(d):
    return "{" + ",".join(f'"{k}":{v}' for k, v in d.items()) + "}"

top_suppliers_hist = [[k, round(v/M,1)] for k,v in DD['top_suppliers_hist'][:10]]
top_items_hist = [[k, round(v/M,1)] for k,v in DD['top_items_hist'][:10]]
top_items_salidas = [[k, round(v/M,1)] for k,v in DD['top_items_salidas'][:10]]
bodega_spend = [[k, round(v/M,1)] for k,v in DD['bodega_spend'][:10]]
bodega_salidas = [[k, round(v/M,1)] for k,v in DD['bodega_salidas'][:10]]
motivo_costo = [[k, round(v/M,1)] for k,v in DD['motivo_costo'] if k][:12]
comprador_counts = [[k, v] for k,v in DD['comprador_counts'][:10]]

estadoOC = {}
for k, v in DD['oc_por_estado'].items():
    if k == 'En proceso de aprobación':
        k = 'En aprobación'
    estadoOC[k] = v
estadoHist = DD['estado_counts']

def obj_counts(d):
    parts = []
    for k, v in d.items():
        key = k if (k.isalpha() and ' ' not in k) else f'"{js_str(k)}"'
        parts.append(f'{key}:{v}')
    return "{" + ", ".join(parts) + "}"

LT = EXTRA['LT_DATA']
fulfillByMonth = EXTRA['D_EXTRA']['fulfillByMonth']
fulfillByYear = EXTRA['D_EXTRA']['fulfillByYear']

D_lines = []
D_lines.append("const D = {")
D_lines.append(f"  monthlySpend: {obj_month(monthly_spend_m)},")
D_lines.append(f"  monthlyOC:    {obj_month(monthly_oc_m)},")
D_lines.append(f"  monthlySalidas:{obj_month(monthly_salidas_m)},")
D_lines.append(f"  topSuppliersHist: {js_arr_tuples(top_suppliers_hist)},")
D_lines.append(f"  topItemsHist: {js_arr_tuples(top_items_hist)},")
D_lines.append(f"  topItemsSalidas: {js_arr_tuples(top_items_salidas)},")
D_lines.append(f"  bodegaSpend: {js_arr_tuples(bodega_spend)},")
D_lines.append(f"  bodegaSalidas: {js_arr_tuples(bodega_salidas)},")
D_lines.append(f"  estadoOC:    {obj_counts(estadoOC)},")
D_lines.append(f"  estadoHist:  {obj_counts(estadoHist)},")
D_lines.append(f"  motivoCosto: {js_arr_tuples(motivo_costo)},")
D_lines.append(f"  compradorCounts: {js_arr_tuples(comprador_counts, vround=0)},")
D_lines.append(f"  fulfillmentRate: {round(DD['fulfillment_rate'],1)},")
D_lines.append(f"  totalSalidas: {round(DD['total_costo_salidas']/M,2)},")
merma_v = next((v for k,v in DD['motivo_costo'] if k=='MERMA'), 0)
accid_v = next((v for k,v in DD['motivo_costo'] if k=='ACCIDENTES - ERROR PEDIDO'), 0)
calidad_v = next((v for k,v in DD['motivo_costo'] if k=='CALIDAD DE MATERIA PRIMA'), 0)
D_lines.append(f"  mermaCosto: {round(merma_v/M,2)},")
D_lines.append(f"  accidentesCosto: {round(accid_v/M,2)},")
D_lines.append(f"  calidadCosto: {round(calidad_v/M,2)},")
D_lines.append(f"  totalFacturas: {DD['n_compras']},")
D_lines.append(f"  totalProveedores: {FILTER_RAW['totalProveedores']},")
D_lines.append(f"  totalItems: {FILTER_RAW['totalItems']},")
D_lines.append(f"  totalIva: {round(DD['total_imptos']/M,1)},")
D_lines.append(f"  ocVencidas: {DD['oc_vencidas']},")
D_lines.append(f"  ocValorPendiente: {round(DD['oc_valor_pendiente']/M,1)},")
D_lines.append(f"  avgLeadTime: {LT['avgReal']},")
D_lines.append(f"  pctOnTime: {LT['pctOnTime']},")
D_lines.append(f"  fulfillByMonth: [{','.join(str(v) for v in fulfillByMonth)}],")
D_lines.append(f"  fulfillByYear:  {{\"2025\":{fulfillByYear['2025']},\"2026\":{fulfillByYear['2026']}}}")
D_lines.append("};")
D_block = "\n".join(D_lines)

with open(scratch + "D_block.txt", "w", encoding="utf-8") as f:
    f.write(D_block)

print("D block written. Lines:", len(D_lines))
print(D_block[:500])

# ── FILTER (near-direct dump of filter_data.json) ──
filter_json = json.dumps(FILTER_RAW, ensure_ascii=False, separators=(',', ':'))
FILTER_block = "const FILTER = " + filter_json + ";"
with open(scratch + "FILTER_block.txt", "w", encoding="utf-8") as f:
    f.write(FILTER_block)
print("\nFILTER block written. Chars:", len(FILTER_block))

# ── PRICE_DATA (near-direct dump of filter_data_v2.json) ──
price_json = json.dumps(PRICE_RAW, ensure_ascii=False, separators=(',', ':'))
PRICE_block = "const PRICE_DATA = " + price_json + ";"
with open(scratch + "PRICE_DATA_block.txt", "w", encoding="utf-8") as f:
    f.write(PRICE_block)
print("PRICE_DATA block written. Chars:", len(PRICE_block))

# ── LT_DATA ──
def obj_stats(s):
    return f'{{ avgReal:{s["avgReal"]}, avgProm:{s["avgProm"]}, pctOn:{s["pctOn"]}, count:{s["count"]} }}'

def js_arr_num(a):
    return "[" + ",".join(str(x) for x in a) + "]"

def js_arr_trip(a):
    parts = []
    for item in a:
        k, v1, v2, v3 = item
        parts.append(f'["{js_str(k)}",{v1},{v2},{v3}]')
    return "[" + ",".join(parts) + "]"

LT_lines = []
LT_lines.append("const LT_DATA = {")
LT_lines.append(f'  total: {LT["total"]}, avgReal: {LT["avgReal"]}, avgProm: {LT["avgProm"]}, pctOnTime: {LT["pctOnTime"]},')
LT_lines.append(f'  por2025: {obj_stats(LT["por2025"])},')
LT_lines.append(f'  por2026: {obj_stats(LT["por2026"])},')
LT_lines.append(f'  monthlyReal: {js_arr_num(LT["monthlyReal"])},')
LT_lines.append(f'  monthlyProm: {js_arr_num(LT["monthlyProm"])},')
LT_lines.append(f'  monthlyPctOn:{js_arr_num(LT["monthlyPctOn"])},')
LT_lines.append(f'  monthlyCount:{js_arr_num(LT["monthlyCount"])},')
LT_lines.append(f'  dist: {{ ontime:{LT["dist"]["ontime"]}, late1_3:{LT["dist"]["late1_3"]}, late4_7:{LT["dist"]["late4_7"]}, late8plus:{LT["dist"]["late8plus"]} }},')
LT_lines.append(f'  topProvMayorLt: {js_arr_trip(LT["topProvMayorLt"])},')
LT_lines.append(f'  topItemsMayorLt:{js_arr_trip(LT["topItemsMayorLt"])}')
LT_lines.append("};")
LT_block = "\n".join(LT_lines)
with open(scratch + "LT_DATA_block.txt", "w", encoding="utf-8") as f:
    f.write(LT_block)
print("\nLT_DATA block written.")
print(LT_block)

# ── DA ──
DA = EXTRA['DA']

def js_arr_pairs_int(pairs):
    return "[" + ",".join(f'["{js_str(k)}",{v}]' for k, v in pairs) + "]"

aj = DA['ajustes']
dev = DA['devoluciones']
me = DA['mermas']

DA_lines = []
DA_lines.append("const DA = {")
DA_lines.append("  ajustes: {")
DA_lines.append(f"    itemsDiferentes: {aj['itemsDiferentes']},")
DA_lines.append(f"    pvDiferentes: {aj['pvDiferentes']},")
DA_lines.append(f"    sumaCostoEntradas: {aj['sumaCostoEntradas']},")
DA_lines.append(f"    sumaCostoSalidas: {aj['sumaCostoSalidas']},")
DA_lines.append(f"    porTiendaPos: [],")
DA_lines.append(f"    porTiendaNeg: {js_arr_pairs_int(aj['porTiendaNeg'])},")
DA_lines.append(f"    motivoCant: {js_arr_pairs_int(aj['motivoCant'])},")
t = aj['tiposAjuste'][0]
DA_lines.append(f'    tiposAjuste: [\n      {{m:"{js_str(t["m"])}", e:{t["e"]}, s:{t["s"]}, n:{t["n"]}, t:{t["t"]}}}\n    ]')
DA_lines.append("  },")
DA_lines.append("  devoluciones: {")
DA_lines.append(f"    total: {dev['total']},")
DA_lines.append(f"    montoTotal: {dev['montoTotal']},")
DA_lines.append(f"    motivos: {js_arr_pairs_int(dev['motivos'])},")
DA_lines.append(f"    porSitio: {js_arr_pairs_int(dev['porSitio'])},")
DA_lines.append(f"    costoProv: {js_arr_pairs_int(dev['costoProv'])},")
DA_lines.append(f"    porMes: {js_arr_pairs_int(dev['porMes'])},")
items_js = ",\n      ".join(
    f'{{item:"{js_str(it["item"])}",um:"{js_str(it["um"])}",cant:{it["cant"]}}}' for it in dev['items']
)
DA_lines.append(f"    items: [\n      {items_js}\n    ]")
DA_lines.append("  },")
DA_lines.append("  mermas: {")
DA_lines.append(f"    totalMermas: {me['totalMermas']},")
DA_lines.append(f"    bodegasOrigen: {me['bodegasOrigen']},")
DA_lines.append(f"    bodegasDestino: {me['bodegasDestino']},")
DA_lines.append(f"    costoItem: {js_arr_pairs_int(me['costoItem'])},")
DA_lines.append(f"    cantItem: {js_arr_pairs_int(me['cantItem'])},")
DA_lines.append(f"    bodegaDest: {js_arr_pairs_int(me['bodegaDest'])},")
DA_lines.append(f"    hist2025: {js_arr_num(me['hist2025'])},")
DA_lines.append(f"    hist2026: {js_arr_num(me['hist2026'])},")
demo_js = ",\n      ".join(
    f'{{f:"{d["f"]}",item:"{js_str(d["item"])}",bo:"{js_str(d["bo"])}",bd:"{js_str(d["bd"])}",um:"{js_str(d["um"])}",cant:{d["cant"]},costo:{d["costo"]},mot:"{d["mot"]}"}}'
    for d in me['detalleDemo']
)
DA_lines.append(f"    detalleDemo: [\n      {demo_js}\n    ]")
DA_lines.append("  }")
DA_lines.append("};")
DA_block = "\n".join(DA_lines)
with open(scratch + "DA_block.txt", "w", encoding="utf-8") as f:
    f.write(DA_block)
print("\nDA block written.")

# ── DN ──
DN = EXTRA['DN']
hoc = DN['histOC']
ocd = DN['ocDetail']

DN_lines = []
DN_lines.append("const DN = {")
DN_lines.append("  histOC: {")
k = hoc['kpi']
DN_lines.append(f"    kpi: {{ total: {k['total']}, aprobadas: {k['aprobadas']}, anuladas: {k['anuladas']}, monto_bn: {k['monto_bn']} }},")
DN_lines.append(f"    anuladas_prov: {js_arr_pairs_int(hoc['anuladas_prov'])},")
DN_lines.append(f"    aprobadas_prov: {js_arr_pairs_int(hoc['aprobadas_prov'])},")
DN_lines.append(f"    cant_um: {js_arr_pairs_int(hoc['cant_um'])},")
DN_lines.append(f"    monto_prov_m: {js_arr_pairs_int(hoc['monto_prov_m'])},")
DN_lines.append(f"    categoria: {js_arr_pairs_int(hoc['categoria'])},")
DN_lines.append(f"    item_count: {js_arr_pairs_int(hoc['item_count'])},")
DN_lines.append(f"    hist_s2025: {js_arr_num(hoc['hist_s2025'])},")
DN_lines.append(f"    hist_s2026: {js_arr_num(hoc['hist_s2026'])},")
DN_lines.append(f"    aprobadas_mes: {js_arr_num(hoc['aprobadas_mes'])},")
DN_lines.append(f"    anuladas_mes:  {js_arr_num(hoc['anuladas_mes'])},")
DN_lines.append(f"    aprobadas_yr:  {{\"2025\":{hoc['aprobadas_yr']['2025']},\"2026\":{hoc['aprobadas_yr']['2026']}}},")
DN_lines.append(f"    anuladas_yr:   {{\"2025\":{hoc['anuladas_yr']['2025']},\"2026\":{hoc['anuladas_yr']['2026']}}}")
DN_lines.append("  },")
DN_lines.append("  ocDetail: {")
kk = ocd['kpi']
DN_lines.append(f"    kpi: {{ monto_bn: {kk['monto_bn']}, cantidad: {kk['cantidad']}, proveedores: {kk['proveedores']}, prom3m: {kk['prom3m']} }},")
prov_var_js = ",".join(f'{{p:"{js_str(pv["p"])}",v:{pv["v"]}}}' for pv in ocd['prov_var'])
DN_lines.append(f"    prov_var: [\n      {prov_var_js}\n    ],")
DN_lines.append(f"    hist_s2025: {js_arr_num(ocd['hist_s2025'])},")
DN_lines.append(f"    hist_s2026: {js_arr_num(ocd['hist_s2026'])},")
DN_lines.append(f"    um_cant: {js_arr_pairs_int(ocd['um_cant'])},")
lbl_js = ",".join(f'"{l}"' for l in ocd['costo_prom_lbl'])
DN_lines.append(f"    costo_prom_lbl: [{lbl_js}],")
DN_lines.append(f"    costo_prom_vals: {js_arr_num(ocd['costo_prom_vals'])},")
tabla_js = ",\n      ".join(
    f'{{a:"{t["a"]}",m:"{t["m"]}",p:"{js_str(t["p"])}",it:"{js_str(t["it"])}",um:"{js_str(t["um"])}",cant:{t["cant"]},pr:{t["pr"]},neto:{t["neto"]}}}'
    for t in ocd['tabla']
)
DN_lines.append(f"    tabla: [\n      {tabla_js}\n    ]")
DN_lines.append("  }")
DN_lines.append("};")
DN_block = "\n".join(DN_lines)
with open(scratch + "DN_block.txt", "w", encoding="utf-8") as f:
    f.write(DN_block)
print("\nDN block written.")

print("\n\nALL BLOCKS WRITTEN OK")
