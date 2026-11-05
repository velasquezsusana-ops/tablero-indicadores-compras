"""Detalle de Transferencias / Mermas, desde "Salidas de inventario.xlsx" (Desc. motivo == 'MERMA').

Genera extract/mermas_detalle.json con TODOS los movimientos de merma en formato compacto
(diccionarios + filas de indices) para el const MERMAS_DETALLE del tablero:
  it: nombres de item   bo: [codigo bodega, desc. bodega]   bd: bodega destino (Desc. C.O. movto.)
  um: unidades          mo: motivos
  r:  [fecha AAMMDD, i_item, i_bodega, i_destino, i_um, i_motivo, cantidad, costo]  (mas reciente primero)
Mismo calculo que actualizarComputeMermasDetalle() en el HTML (boton "Actualizar datos").
"""
import openpyxl, json, sys
from datetime import datetime

base = "C:/Users/Montolivo/Dropbox/Asesorias PYMES/1. Proyectos Actuales/2. MONTOLIVO/2. Proyecto/Indicadores Montolivo/Compras/Bases para actualizar Claude/"
MOTIVOS = {'MERMA'}

def num(v):
    try: return float(v) if v not in (None, '') else 0.0
    except Exception: return 0.0

class Dic:
    def __init__(self): self.lst, self.idx = [], {}
    def get(self, k):
        if k not in self.idx: self.idx[k] = len(self.lst); self.lst.append(k)
        return self.idx[k]

it, bo, bd, um, mo = Dic(), Dic(), Dic(), Dic(), Dic()
rows = []
wb = openpyxl.load_workbook(base + "Salidas de inventario.xlsx", data_only=True, read_only=True)
hdr = None
for i, row in enumerate(wb.active.iter_rows(values_only=True)):
    if i == 0: hdr = [str(c).strip() if c else '' for c in row]; continue
    r = dict(zip(hdr, row))
    motivo = str(r.get('Desc. motivo') or '').strip().upper()
    if motivo not in MOTIVOS: continue
    f = r.get('Fecha')
    if not isinstance(f, datetime): continue
    ymd = (f.year % 100) * 10000 + f.month * 100 + f.day
    cod = str(r.get('Bodega') or '').strip()
    desc = str(r.get('Desc. bodega') or '').strip()
    rows.append([ymd, it.get(str(r.get('Desc. item') or '').strip()), bo.get((cod, desc)),
                 bd.get(str(r.get('Desc. C.O. movto.') or '').strip()), um.get(str(r.get('U.M. inv.') or '').strip()),
                 mo.get(motivo), round(num(r.get('Salidas (inv.)')), 2), round(num(r.get('Costo salidas (prom.)')))])
wb.close()
rows.sort(key=lambda x: (-x[0], -x[7]))
for x in rows:
    if x[6] == int(x[6]): x[6] = int(x[6])
out = {"it": it.lst, "bo": [list(k) for k in bo.lst], "bd": bd.lst, "um": um.lst, "mo": mo.lst, "r": rows}
txt = json.dumps(out, ensure_ascii=False, separators=(',', ':'))
open(base + "extract/mermas_detalle.json", "w", encoding="utf-8").write(txt)
print(f"Filas: {len(rows)}  items: {len(it.lst)}  bodegas: {len(bo.lst)}  destinos: {len(bd.lst)}  bytes: {len(txt):,}")
print("Rango:", rows[-1][0], "->", rows[0][0])
