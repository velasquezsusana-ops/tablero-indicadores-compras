"""Devoluciones a proveedor, desde "Histórico de compras por item.xlsx".

Toma las filas con Cantidad < 0 (notas de devolución, motivos 402-xx). Las filas con cantidad 0
(notas que solo ajustan precio) NO se incluyen. Genera extract/devoluciones_proveedor.json en el
formato compacto del const DEVPROV_DATA del tablero:
  pv: proveedores   it: items   um: unidades   bo: bodegas   mo: motivos   doc: documentos
  r:  [fecha AAMMDD, i_prov, i_item, i_um, i_bodega, i_motivo, i_doc, cantidad, valor neto]
      (cantidad y valor tal como vienen: negativos; mas reciente primero)
Mismo calculo que actualizarComputeDevProv() en el HTML (boton "Actualizar datos").
"""
import openpyxl, json
from datetime import datetime

base = "C:/Users/Montolivo/Dropbox/Asesorias PYMES/1. Proyectos Actuales/2. MONTOLIVO/2. Proyecto/Indicadores Montolivo/Compras/Bases para actualizar Claude/"

def num(v):
    try: return float(v) if v not in (None, '') else 0.0
    except Exception: return 0.0

class Dic:
    def __init__(self): self.lst, self.idx = [], {}
    def get(self, k):
        if k not in self.idx: self.idx[k] = len(self.lst); self.lst.append(k)
        return self.idx[k]

pv, it, um, bo, mo, doc = Dic(), Dic(), Dic(), Dic(), Dic(), Dic()
rows = []
wb = openpyxl.load_workbook(base + "Histórico de compras por item.xlsx", data_only=True, read_only=True)
ws = wb[wb.sheetnames[0]]
hdr = None
for i, row in enumerate(ws.iter_rows(values_only=True)):
    if i == 0: hdr = [str(c).strip() if c else '' for c in row]; continue
    r = dict(zip(hdr, row))
    cant = num(r.get('Cantidad'))
    if cant >= 0: continue
    f = r.get('Fecha')
    if not isinstance(f, datetime): continue
    prov = str(r.get('Razón social proveedor') or r.get('Razon social proveedor') or '').strip()
    rows.append([(f.year % 100) * 10000 + f.month * 100 + f.day, pv.get(prov),
                 it.get(str(r.get('Desc. item') or '').strip()), um.get(str(r.get('U.M.') or '').strip()),
                 bo.get(str(r.get('Bodega') or '').strip()), mo.get(str(r.get('Motivo') or '').strip()),
                 doc.get(str(r.get('Documento') or '').strip()),
                 round(cant, 2), round(num(r.get('Valor neto local')))])
wb.close()
rows.sort(key=lambda x: (-x[0], x[8]))
for x in rows:
    if x[7] == int(x[7]): x[7] = int(x[7])
out = {"pv": pv.lst, "it": it.lst, "um": um.lst, "bo": bo.lst, "mo": mo.lst, "doc": doc.lst, "r": rows}
txt = json.dumps(out, ensure_ascii=False, separators=(',', ':'))
open(base + "extract/devoluciones_proveedor.json", "w", encoding="utf-8").write(txt)
print(f"Filas: {len(rows)}  proveedores: {len(pv.lst)}  items: {len(it.lst)}  valor: {sum(x[8] for x in rows):,.0f}  bytes: {len(txt):,}")
