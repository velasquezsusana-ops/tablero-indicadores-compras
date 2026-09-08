"""
Insercion UNICA (primera vez) del nuevo const LT_LINE_DETAIL en el HTML, justo
despues de OC_PEND_LIST. Para refrescos futuros usar update_lt_detail.py (splice,
no insert) una vez que el const ya existe en el archivo.
"""
import json
from pathlib import Path

BASE = Path("C:/Users/Montolivo/Dropbox/Asesorias PYMES/1. Proyectos Actuales/2. MONTOLIVO/2. Proyecto/Indicadores Montolivo/Compras/Bases para actualizar Claude")
HTML_PATH = BASE / "Tablero_Indicadores_Compras.html"
JSON_PATH = BASE / "lt_line_detail.json"


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

    marker = "const OC_PEND_LIST = "
    start = html_text.find(marker)
    if start == -1:
        raise ValueError("No se encontro 'const OC_PEND_LIST = ' en el HTML")
    value_start = start + len(marker)
    end = find_balanced_end(html_text, value_start)

    if "const LT_LINE_DETAIL = " in html_text:
        raise ValueError("LT_LINE_DETAIL ya existe en el HTML — usar update_lt_detail.py en vez de este script")

    insertion = f"\nconst LT_LINE_DETAIL = {new_json};"
    updated = html_text[:end] + insertion + html_text[end:]
    HTML_PATH.write_text(updated, encoding="utf-8", newline="\n")
    print(f"Insertado const LT_LINE_DETAIL con {len(data)} filas, despues de OC_PEND_LIST.")


if __name__ == "__main__":
    main()
