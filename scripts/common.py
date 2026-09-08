import unicodedata, re
from datetime import date, datetime

MESES = ["enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre"]
MESES_IDX = {m: i+1 for i, m in enumerate(MESES)}

def mes_nombre(m):
    return MESES[m-1]

def to_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, str):
        return parse_spanish_date_str(v)
    return None

def _strip_accents(s):
    nfkd = unicodedata.normalize("NFD", s)
    return "".join(c for c in nfkd if unicodedata.category(c) != "Mn")

def parse_spanish_date_str(s):
    s2 = _strip_accents(s.lower())
    m = re.search(r"(\d{1,2})\s*de\s+([a-z]+)\s*de\s+(\d{4})", s2)
    if not m:
        return None
    dia = int(m.group(1))
    mes_txt = m.group(2)
    anio = int(m.group(3))
    meses_norm = [_strip_accents(x) for x in MESES]
    if mes_txt not in meses_norm:
        return None
    mes_num = meses_norm.index(mes_txt) + 1
    try:
        return date(anio, mes_num, dia)
    except ValueError:
        return None

def weeknum_type1(iso_date_str):
    d = date.fromisoformat(iso_date_str)
    jan1 = date(d.year, 1, 1)
    jan1_dow = (jan1.weekday() + 1) % 7  # 0=Sunday
    days_since_jan1 = (d - jan1).days
    return (days_since_jan1 + jan1_dow) // 7 + 1
