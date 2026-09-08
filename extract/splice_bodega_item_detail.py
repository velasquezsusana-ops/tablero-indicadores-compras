"""
Splice puntual de BODEGA_ITEM_DETAIL (ya con desglose mensual, agregado 2026-07-31)
en Tablero_Indicadores_Compras.html. Correr despues de extract_bodega_item_detail.py.
"""
import json
from pathlib import Path

BASE = Path("C:/Users/Montolivo/Dropbox/Asesorias PYMES/1. Proyectos Actuales/2. MONTOLIVO/2. Proyecto/Indicadores Montolivo/Compras/Bases para actualizar Claude")
HTML_PATH = BASE / "Tablero_Indicadores_Compras.html"
JSON_PATH = BASE / "bodega_item_detail.json"
CONST_NAME = "BODEGA_ITEM_DETAIL"


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


def main():
    with open(JSON_PATH, encoding="utf-8") as f:
        data = json.load(f)
    new_json = json.dumps(data, separators=(",", ":"), ensure_ascii=False)

    html_text = HTML_PATH.read_text(encoding="utf-8")
    marker = f"const {CONST_NAME} = "
    start = html_text.find(marker)
    if start == -1:
        raise ValueError(f"No se encontro 'const {CONST_NAME} = ' en el HTML")
    value_start = start + len(marker)
    end = find_balanced_end(html_text, value_start)
    new_block = marker + new_json + ";"
    updated = html_text[:start] + new_block + html_text[end:]
    HTML_PATH.write_text(updated, encoding="utf-8", newline="\n")
    print(f"Listo. {CONST_NAME} actualizado ({len(new_json)} bytes de JSON).")


if __name__ == "__main__":
    main()
