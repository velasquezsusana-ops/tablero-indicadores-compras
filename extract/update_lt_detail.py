"""
Actualiza el detalle linea a linea de Lead Time (const LT_LINE_DETAIL) en
Tablero_Indicadores_Compras.html a partir de "Historico de compras por item.xlsx".

Correr este script SOLO despues de reemplazar el excel "Historico de compras por item.xlsx"
con datos nuevos (y solo si LT_LINE_DETAIL ya existe en el HTML — para la primera insercion
se uso insert_lt_line_detail.py, que ya no hace falta volver a correr). Hace las dos cosas
en un solo paso:
  1. Re-extrae lt_line_detail.json desde el excel (extract_lt_line_detail.py).
  2. Reemplaza el bloque "const LT_LINE_DETAIL = [...];" dentro del HTML por los datos nuevos,
     sin tocar el resto del archivo.
"""
import json
import subprocess
import sys
from pathlib import Path

BASE = Path("C:/Users/Montolivo/Dropbox/Asesorias PYMES/1. Proyectos Actuales/2. MONTOLIVO/2. Proyecto/Indicadores Montolivo/Compras/Bases para actualizar Claude")
HTML_PATH = BASE / "Tablero_Indicadores_Compras.html"
JSON_PATH = BASE / "lt_line_detail.json"
EXTRACT_SCRIPT = BASE / "extract" / "extract_lt_line_detail.py"
CONST_NAME = "LT_LINE_DETAIL"


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


def splice_const(html_text, const_name, new_value_json):
    marker = f"const {const_name} = "
    start = html_text.find(marker)
    if start == -1:
        raise ValueError(f"No se encontro 'const {const_name} = ' en el HTML")
    value_start = start + len(marker)
    end = find_balanced_end(html_text, value_start)
    new_block = marker + new_value_json + ";"
    return html_text[:start] + new_block + html_text[end:]


def main():
    print(f"1) Re-extrayendo datos desde el excel (corriendo {EXTRACT_SCRIPT.name})...")
    result = subprocess.run([sys.executable, str(EXTRACT_SCRIPT)], capture_output=True, text=True)
    print(result.stdout)
    if result.returncode != 0:
        print(result.stderr, file=sys.stderr)
        sys.exit(1)

    print(f"2) Leyendo {JSON_PATH.name}...")
    with open(JSON_PATH, encoding="utf-8") as f:
        data = json.load(f)
    new_json = json.dumps(data, separators=(",", ":"), ensure_ascii=False)

    print(f"3) Insertando {len(data)} filas en el HTML como const {CONST_NAME}...")
    html_text = HTML_PATH.read_text(encoding="utf-8")
    updated = splice_const(html_text, CONST_NAME, new_json)
    HTML_PATH.write_text(updated, encoding="utf-8", newline="\n")

    print(f"Listo. {CONST_NAME} actualizado con {len(data)} filas en {HTML_PATH.name}.")


if __name__ == "__main__":
    main()
