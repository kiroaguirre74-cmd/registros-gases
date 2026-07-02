#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Genera libros de Excel semanales de registros de gases usando como
plantilla el archivo de ejemplo, con datos aleatorios realistas y
fechas actualizadas a cada semana. Solo usa la biblioteca estandar.

Parametros por variables de entorno:
  SERIAL     -> numero de serie (columna A)          [def: MA215-012525]
  SHEET_NAME -> nombre de la hoja                     [def: mantener plantilla]
  OUTDIR     -> carpeta de salida                     [def: salida]
  VARY_ROWS  -> "1" para variar filas 15-30           [def: 0 = 29 fijas]
  ROW_MIN/ROW_MAX -> rango de filas si VARY_ROWS      [def: 15 / 30]
  MES_INI/MES_FIN -> rango de meses a generar         [def: 1 / 12]
"""
import zipfile, os, re, random, calendar
from datetime import datetime, timedelta, date

TEMPLATE = os.environ.get("TEMPLATE", "plantilla/ENE_09-01-2025 al 15-01-2025.xlsx")
EPOCH = datetime(1899, 12, 30)
M = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

MESES = {1: "ENE", 2: "FEB", 3: "MAR", 4: "ABR", 5: "MAY", 6: "JUN",
         7: "JUL", 8: "AGO", 9: "SEP", 10: "OCT", 11: "NOV", 12: "DIC"}

READ_COLS = {"E", "F", "G", "H", "L", "M", "N", "R", "V"}

SERIAL = os.environ.get("SERIAL", "MA215-012525")
SHEET_NAME = os.environ.get("SHEET_NAME", "").strip()
OUTDIR = os.environ.get("OUTDIR", "salida")
VARY_ROWS = os.environ.get("VARY_ROWS", "0") == "1"
ROW_MIN = int(os.environ.get("ROW_MIN", 15))
ROW_MAX = int(os.environ.get("ROW_MAX", 30))


def col_of(ref):
    return re.match(r"[A-Z]+", ref).group(0)


def serial_to_dt(s):
    return EPOCH + timedelta(days=float(s))


def dt_to_serial(dt):
    return (dt - EPOCH).total_seconds() / 86400.0


def esc(t):
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


# ---------------- semanas (lun-dom recortadas; fusiona <=2 dias) --------------
def semanas_del_mes(year, month):
    ndays = calendar.monthrange(year, month)[1]
    dias = [date(year, month, d) for d in range(1, ndays + 1)]
    grupos, actual = [], []
    for d in dias:
        if actual and d.weekday() == 0:
            grupos.append(actual)
            actual = []
        actual.append(d)
    if actual:
        grupos.append(actual)
    cambiado = True
    while cambiado and len(grupos) > 1:
        cambiado = False
        for i, g in enumerate(grupos):
            if len(g) <= 2:
                if i == 0:
                    grupos[1] = grupos[0] + grupos[1]; del grupos[0]
                else:
                    grupos[i - 1] = grupos[i - 1] + grupos[i]; del grupos[i]
                cambiado = True
                break
    return [(g[0], g[-1]) for g in grupos]


# ---------------- lectura de la plantilla ------------------------------------
def leer_plantilla():
    z = zipfile.ZipFile(TEMPLATE)
    xml = z.read("xl/worksheets/sheet1.xml").decode("utf-8")
    ss = []
    root = __import__("xml.etree.ElementTree", fromlist=["ET"]).fromstring(
        z.read("xl/sharedStrings.xml"))
    for si in root.findall(M + "si"):
        ss.append("".join(t.text or "" for t in si.iter(M + "t")))
    z.close()
    ini = xml.index("<sheetData>") + len("<sheetData>")
    fin = xml.index("</sheetData>")
    cabecera = xml[:ini]
    cola = xml[fin:]
    cuerpo = xml[ini:fin]
    filas_xml = re.findall(r"<row\b.*?</row>", cuerpo, re.S)
    return cabecera, cola, filas_xml, ss


def parse_row(row_xml):
    m = re.match(r"<row\b([^>]*)>(.*)</row>", row_xml, re.S)
    attrs, body = m.group(1), m.group(2)
    celdas = []
    for cm in re.finditer(r'<c r="([A-Z]+\d+)"([^>]*?)(?:/>|>(.*?)</c>)', body, re.S):
        ref, catrs, inner = cm.group(1), cm.group(2), cm.group(3)
        sm = re.search(r's="(\d+)"', catrs)
        tm = re.search(r't="(\w+)"', catrs)
        celdas.append([ref, sm.group(1) if sm else None,
                       tm.group(1) if tm else None, inner])
    return attrs, celdas


# ---------------- valores aleatorios de gas ----------------------------------
def gen_lecturas(rnd):
    elevado = rnd.random() < 0.30
    if elevado:
        h2s = round(rnd.uniform(10.0, 14.5), 1)
        stel = round(rnd.uniform(0.0, 0.3), 1)
    else:
        h2s = 0.0
        stel = 0.0
    vals = {
        "F": f"{h2s:.1f}", "G": f"{stel:.1f}", "H": "0.0",
        "L": "0.0", "M": "0.0", "N": "0.0",
        "R": rnd.choice(["20.9", "20.9", "20.9", "20.8", "21.0"]), "V": "0.0",
    }
    vals["_alarma"] = elevado
    return vals


def celda_xml(ref, style, tipo, inner):
    s = f' s="{style}"' if style else ""
    if inner is None:
        return f'<c r="{ref}"{s}/>' if not tipo else f'<c r="{ref}"{s} t="{tipo}"/>'
    t = f' t="{tipo}"' if tipo else ""
    return f'<c r="{ref}"{s}{t}>{inner}</c>'


def celda_inline(ref, style, texto):
    s = f' s="{style}"' if style else ""
    return f'<c r="{ref}"{s} t="inlineStr"><is><t>{esc(texto)}</t></is></c>'


def celda_num(ref, style, valor):
    s = f' s="{style}"' if style else ""
    return f'<c r="{ref}"{s}><v>{valor}</v></c>'


# ---------------- seleccion de filas (cantidad variable) ---------------------
def clasificar(filas, ss):
    """Devuelve lista de (indice, tipo) para las filas de datos (1..n)."""
    cats = []
    for idx in range(1, len(filas)):
        _, celdas = filas[idx]
        C = D = None
        for ref, style, tipo, inner in celdas:
            c = col_of(ref)
            if c in ("C", "D") and inner is not None:
                m = re.search(r"<v>(\d+)</v>", inner)
                val = ss[int(m.group(1))] if (m and tipo == "s") else None
                if c == "C":
                    C = val
                else:
                    D = val
        if C == "Lecturas":
            t = "reading"
        elif D == "Encendido":
            t = "start"
        elif D == "Apagado manual":
            t = "stop"
        else:
            t = "sensor"
        cats.append((idx, t))
    return cats


def seleccionar_indices(cats, n_obj, rnd):
    idxs = [i for i, _ in cats]
    tipos = {i: t for i, t in cats}
    if n_obj >= len(idxs):
        sel = list(idxs)
        # anadir duplicados de lecturas si hace falta (n_obj>total)
        readings = [i for i in idxs if tipos[i] == "reading"]
        while len(sel) < n_obj and readings:
            dup = rnd.choice(readings)
            pos = sel.index(dup)
            sel.insert(pos + 1, dup)
        return sel
    # hay que quitar (len-n_obj) filas: primero sensores, luego lecturas extra
    # conservar: todos start/stop y la 1a lectura de cada grupo contiguo
    esencial = set()
    prev_t = None
    for i, t in cats:
        if t in ("start", "stop"):
            esencial.add(i)
        elif t == "reading" and prev_t != "reading":
            esencial.add(i)  # primera lectura del grupo
        prev_t = t
    droppable = [i for i in idxs if i not in esencial]
    rnd.shuffle(droppable)
    quitar = set(droppable[: len(idxs) - n_obj])
    return [i for i in idxs if i not in quitar]


# ---------------- construir sheet1.xml de una semana -------------------------
def construir_sheet(cabecera, cola, filas_xml, ss, d_ini, d_fin, seed):
    rnd = random.Random(seed)
    filas = [parse_row(fx) for fx in filas_xml]

    # dt original por fila
    dts = [None]
    for idx in range(1, len(filas)):
        _, celdas = filas[idx]
        dtval = None
        for ref, style, tipo, inner in celdas:
            if col_of(ref) == "B" and inner is not None and tipo != "s":
                m = re.search(r"<v>([^<]+)</v>", inner)
                if m:
                    dtval = serial_to_dt(m.group(1))
                break
        dts.append(dtval)
    ultimo = None
    for i in range(1, len(dts)):
        if dts[i] is None:
            dts[i] = (ultimo or serial_to_dt("45666")) + timedelta(minutes=1)
        ultimo = dts[i]

    # seleccion de filas (cantidad variable opcional)
    cats = clasificar(filas, ss)
    if VARY_ROWS:
        n_obj = rnd.randint(ROW_MIN, ROW_MAX)
    else:
        n_obj = len(filas) - 1
    seleccion = seleccionar_indices(cats, n_obj, rnd)

    # mapeo de dias de plantilla -> dias de la semana destino
    dias_orig = sorted(set(dts[i].date() for i in seleccion))
    dias_semana = [d_ini + timedelta(days=i) for i in range((d_fin - d_ini).days + 1)]
    n = len(dias_orig)
    if n == 1:
        elegidos = [dias_semana[len(dias_semana) // 2]]
    else:
        elegidos = [dias_semana[round(k * (len(dias_semana) - 1) / (n - 1))]
                    for k in range(n)]
    mapa = dict(zip(dias_orig, elegidos))

    partes = [cabecera]
    # cabecera de columnas (fila 1) intacta
    partes.append(filas_xml[0])

    prev_dt = None
    nueva_fila = 1
    for idx in seleccion:
        nueva_fila += 1
        attrs, celdas = filas[idx]
        dt_o = dts[idx]
        nuevo_dia = mapa[dt_o.date()]
        jitter = timedelta(seconds=rnd.randint(0, 120))
        nuevo_dt = datetime.combine(nuevo_dia, dt_o.time()) + jitter
        if prev_dt is not None and nuevo_dt <= prev_dt:
            nuevo_dt = prev_dt + timedelta(seconds=rnd.randint(1, 45))
        if nuevo_dt.date() > d_fin:
            nuevo_dt = datetime.combine(d_fin, dt_o.time())
            if prev_dt is not None and nuevo_dt <= prev_dt:
                nuevo_dt = prev_dt + timedelta(seconds=rnd.randint(1, 45))
        prev_dt = nuevo_dt
        serial = f"{dt_to_serial(nuevo_dt):.11f}"

        tiene_lecturas = any(col_of(r) == "F" and inn is not None
                             for r, _, _, inn in celdas)
        vals = gen_lecturas(rnd) if tiene_lecturas else None

        # atributos de fila: renumerar r
        nattrs = re.sub(r'\s*r="\d+"', "", attrs)
        row_open = f'<row r="{nueva_fila}"{nattrs}>'

        celdas_xml = []
        for ref, style, tipo, inner in celdas:
            c = col_of(ref)
            nref = f"{c}{nueva_fila}"
            if c == "A":
                celdas_xml.append(celda_inline(nref, style, SERIAL))
            elif c == "B":
                celdas_xml.append(celda_num(nref, style or "4", serial))
            elif vals and c in READ_COLS and c != "E" and inner is not None:
                celdas_xml.append(celda_inline(nref, style, vals[c]))
            elif c == "E":
                if vals and vals.get("_alarma") and inner is not None:
                    celdas_xml.append(celda_inline(nref, style, "Alarma de nivel bajo"))
                elif inner is not None:
                    celdas_xml.append(celda_xml(nref, style, None, None))
                else:
                    celdas_xml.append(celda_xml(nref, style, tipo, inner))
            else:
                celdas_xml.append(celda_xml(nref, style, tipo, inner))
        partes.append(row_open + "".join(celdas_xml) + "</row>")

    # ajustar dimension al numero real de filas
    cab = re.sub(r'<dimension ref="A1:BC\d+"/>',
                 f'<dimension ref="A1:BC{nueva_fila}"/>', partes[0])
    partes[0] = cab
    partes.append(cola)
    return "".join(partes)


# ---------------- escribir xlsx ----------------------------------------------
def escribir_xlsx(destino, nuevo_sheet):
    src = zipfile.ZipFile(TEMPLATE, "r")
    with zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                data = nuevo_sheet.encode("utf-8")
            elif item.filename == "xl/workbook.xml" and SHEET_NAME:
                txt = data.decode("utf-8")
                txt = re.sub(r'(<sheet name=")[^"]*(")',
                             r"\1" + esc(SHEET_NAME) + r"\2", txt, count=1)
                data = txt.encode("utf-8")
            dst.writestr(item, data)
    src.close()


def main():
    os.makedirs(OUTDIR, exist_ok=True)
    cabecera, cola, filas_xml, ss = leer_plantilla()
    resumen = []
    ini = int(os.environ.get("MES_INI", 1))
    fin = int(os.environ.get("MES_FIN", 12))
    year = int(os.environ.get("YEAR", 2026))
    for mes in range(ini, fin + 1):
        for (d_ini, d_fin) in semanas_del_mes(year, mes):
            nombre = (f"{MESES[mes]}_{d_ini.strftime('%d-%m-%Y')} al "
                      f"{d_fin.strftime('%d-%m-%Y')}.xlsx")
            seed = int(d_ini.strftime("%Y%m%d")) + (0 if SERIAL.startswith("MA215") else 700000)
            sheet = construir_sheet(cabecera, cola, filas_xml, ss, d_ini, d_fin, seed)
            escribir_xlsx(os.path.join(OUTDIR, nombre), sheet)
            resumen.append(nombre)
    print(f"Generados {len(resumen)} archivos en '{OUTDIR}/' "
          f"(serie={SERIAL}, hoja={SHEET_NAME or 'plantilla'}, "
          f"filas={'15-30' if VARY_ROWS else 'fijas'})")
    for n in resumen:
        print("  -", n)


if __name__ == "__main__":
    main()
