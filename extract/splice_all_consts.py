"""
Reemplaza en Tablero_Indicadores_Compras.html los 16 const generados por el pipeline
(6 armados por build_js_blocks.py + 10 volcados casi directos de su JSON), usando un
escaner de llaves/corchetes balanceado para encontrar el cierre exacto de cada uno
(no depende de que el const quepa en una sola linea ni de su longitud).

Correr DESPUES de: los extract_*.py correspondientes y build_js_blocks.py.
No toca OC_PEND_LIST (ese tiene su propio script: update_oc_pendientes.py).
"""
import json
from pathlib import Path

BASE = Path("C:/Users/Montolivo/Dropbox/Asesorias PYMES/1. Proyectos Actuales/2. MONTOLIVO/2. Proyecto/Indicadores Montolivo/Compras/Bases para actualizar Claude")
SCRATCH = Path("C:/Users/MONTOL~1/AppData/Local/Temp/claude/C--Users-Montolivo/27afce9d-8783-40f9-967e-828b75f5dde0/scratchpad")
HTML_PATH = BASE / "Tablero_Indicadores_Compras.html"

BLOCK_FILES = {
    "D": "D_block.txt",
    "FILTER": "FILTER_block.txt",
    "PRICE_DATA": "PRICE_DATA_block.txt",
    "LT_DATA": "LT_DATA_block.txt",
    "DA": "DA_block.txt",
    "DN": "DN_block.txt",
}

JSON_CONSTS = {
    "BODEGA_DETAIL": "bodega_detail.json",
    "LT_DETAIL": "lt_detail.json",
    "OC_VENC_DETAIL": "ocvenc_detail.json",
    "OC_PEND_DETAIL": "ocpend_detail.json",
    "BODEGA_MASTER": "bodega_master.json",
    "HIST_PROV_ITEM": "historico_prov_item.json",
    "BODEGA_ITEM_DETAIL": "bodega_item_detail.json",
    "OTIF_DATA": "extract/otif_data.json",
    "VENTAS_DATA": "extract/ventas_data.json",
    "AJUSTES_DATA": "extract/ajustes_data.json",
}


def find_balanced_end(text, start):
    depth = 0
    i = start
    in_string = False
    n = len(text)
    while i < n:
        c = text[i]
        if in_string:
            if c == '\\':
                i += 2
                continue
            if c == '"':
                in_string = False
        else:
            if c == '"':
                in_string = True
            elif c in '[{':
                depth += 1
            elif c in ']}':
                depth -= 1
                if depth == 0:
                    j = i + 1
                    while j < n and text[j] in ' \t':
                        j += 1
                    if j < n and text[j] == ';':
                        return j + 1
                    raise ValueError("No se encontro ';' despues del cierre del bloque")
        i += 1
    raise ValueError("No se encontro el cierre balanceado del bloque")


def splice_full_statement(html_text, const_name, new_full_statement):
    marker = f"const {const_name} = "
    start = html_text.find(marker)
    if start == -1:
        raise ValueError(f"No se encontro 'const {const_name} = ' en el HTML")
    value_start = start + len(marker)
    end = find_balanced_end(html_text, value_start)
    return html_text[:start] + new_full_statement + html_text[end:]


def main():
    html_text = HTML_PATH.read_text(encoding="utf-8")

    for name, fname in BLOCK_FILES.items():
        block_text = (SCRATCH / fname).read_text(encoding="utf-8")
        html_text = splice_full_statement(html_text, name, block_text)
        print(f"  Actualizado const {name} (desde {fname})")

    for name, fname in JSON_CONSTS.items():
        with open(BASE / fname, encoding="utf-8") as f:
            data = json.load(f)
        new_json = json.dumps(data, separators=(",", ":"), ensure_ascii=False)
        statement = f"const {name} = {new_json};"
        html_text = splice_full_statement(html_text, name, statement)
        print(f"  Actualizado const {name} (desde {fname})")

    HTML_PATH.write_text(html_text, encoding="utf-8", newline="\n")
    print("HTML actualizado correctamente (16 consts).")


if __name__ == "__main__":
    main()
