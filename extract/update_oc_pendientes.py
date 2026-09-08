"""
Actualiza el detalle de Ordenes Pendientes/Vencidas (const OC_PEND_LIST) en
Tablero_Indicadores_Compras.html a partir de "Orden de compra por item.xlsx".

Correr este script SOLO despues de reemplazar el excel "Orden de compra por item.xlsx"
con datos nuevos. Hace las dos cosas en un solo paso:
  1. Re-extrae oc_pendientes_detalle.json desde el excel (extract_oc_pendientes_detalle.py).
  2. Reemplaza el bloque "const OC_PEND_LIST = [...];" dentro del HTML por los datos nuevos,
     sin tocar el resto del archivo.

No requiere copiar/pegar nada a mano ni recordar donde va el bloque.
"""
import json
import subprocess
import sys
from pathlib import Path

BASE = Path("C:/Users/Montolivo/Dropbox/Asesorias PYMES/1. Proyectos Actuales/2. MONTOLIVO/2. Proyecto/Indicadores Montolivo/Compras/Bases para actualizar Claude")
HTML_PATH = BASE / "Tablero_Indicadores_Compras.html"
JSON_PATH = BASE / "oc_pendientes_detalle.json"
EXTRACT_SCRIPT = BASE / "extract" / "extract_oc_pendientes_detalle.py"
CONST_NAME = "OC_PEND_LIST"


def find_balanced_end(text, start):
    """A partir del indice del '[' de apertura, devuelve el indice justo despues
    del ';' que cierra la sentencia `const NAME = [...];`. Respeta strings con
    comillas escapadas para no confundir corchetes/llaves dentro de texto."""
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
                    # buscar el ';' inmediatamente despues (puede haber espacios)
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
