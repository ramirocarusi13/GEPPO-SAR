"""
GEPPO automatico: toma los Excel que cargan los departamentos y los pasa al
MonthlyReport de Juan Salinas, como lo hace el a mano.

Uso:
  py geppo_cargar.py                 -> carga el mes vencido (una vez) y termina
  py geppo_cargar.py --mes 202610    -> carga un mes puntual
  py geppo_cargar.py --vigilar       -> queda revisando cada X minutos y carga cada depto ni bien cambia su archivo
  py geppo_cargar.py --forzar        -> recarga todos los deptos aunque no hayan cambiado
  py geppo_cargar.py --produccion    -> usa las rutas reales de Z: (ver config.json)

Nunca modifica los archivos de los departamentos ni el MonthlyReport original de Juan:
escribe en "MonthlyReport_SAR_YYYYMM_AUTO.xlsx". Por cada depto agrega una hoja
"AUTO <depto>" con lo que cargo, lo que hay que revisar, los textos y las fotos del depto.
Esas hojas se borran antes de enviar el informe.
"""
import argparse, datetime as dt, fnmatch, glob, json, os, re, shutil, time, traceback, warnings
from decimal import Decimal, ROUND_HALF_UP

import openpyxl
from openpyxl.utils import column_index_from_string as col_num, get_column_letter as col_letra, range_boundaries

warnings.filterwarnings("ignore")  # avisos de openpyxl por encabezados/imagenes que no lee (no afectan)

AQUI = os.path.dirname(os.path.abspath(__file__))
MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
         "septiembre", "octubre", "noviembre", "diciembre"]
MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August",
          "September", "October", "November", "December"]
SEGUNDOS_ESTABLE = 90  # no tocar un archivo de depto que se guardo hace menos de esto (puede estar a medio copiar)


# ---------------------------------------------------------------- utilidades
def log(msg):
    linea = f"{dt.datetime.now():%Y-%m-%d %H:%M:%S}  {msg}"
    print(linea, flush=True)
    with open(os.path.join(AQUI, "registro.log"), "a", encoding="utf-8") as f:
        f.write(linea + "\n")


def ruta(p):
    return os.path.normpath(p if os.path.isabs(p) else os.path.join(AQUI, p))


def cargar_config(produccion=False, archivo=None):
    with open(archivo or os.path.join(AQUI, "config.json"), encoding="utf-8") as f:
        cfg = json.load(f)
    if archivo:  # rutas relativas al archivo de config
        base = os.path.dirname(os.path.abspath(archivo))
        for k in ("carpeta_deptos", "carpeta_format", "carpeta_salida"):
            if not os.path.isabs(cfg[k]):
                cfg[k] = os.path.normpath(os.path.join(base, cfg[k]))
    if produccion:
        cfg.update(cfg["produccion"])
    return cfg


def mes_vencido(hoy=None):
    hoy = hoy or dt.date.today()
    return f"{hoy.replace(day=1) - dt.timedelta(days=1):%Y%m}"


def mes_anterior(mes):
    y, m = int(mes[:4]), int(mes[4:])
    return f"{y - 1}12" if m == 1 else f"{y}{m - 1:02d}"


def meses_entre(desde_yyyy_mm, mes):
    a, m = map(int, desde_yyyy_mm.split("-"))
    return (int(mes[:4]) - a) * 12 + (int(mes[4:]) - m)


def normal(s):
    return s.strip().replace("　", " ").replace("．", ".").lower()


def nombre_hoja(nombres, prefijo):
    """Busca la hoja por prefijo (los nombres traen espacios japoneses y caracteres raros)."""
    p = normal(prefijo)
    for n in nombres:
        if normal(n).startswith(p):
            return n
    raise KeyError(f"no encontre la hoja '{prefijo}' (hay: {', '.join(nombres)})")


def redondear(v, dec):
    return float(Decimal(str(v)).quantize(Decimal(1).scaleb(-dec), rounding=ROUND_HALF_UP))


def a_numero(v):
    """'0,,4%' -> 0.004 ; '3,51' -> 3.51 ; numeros quedan igual."""
    if v is None or isinstance(v, (int, float)):
        return v
    s = str(v).strip().replace(" ", "")
    if not s:
        return None
    pct = "%" in s
    s = re.sub(",+", ",", s.replace("%", "")).replace(",", ".")
    try:
        x = float(s)
    except ValueError:
        return v
    return x / 100 if pct else x


def transformar(v, t):
    if t in (None, "", "copiar") or v is None:
        return v
    if t == "numero":
        return a_numero(v)
    v = a_numero(v)
    if not isinstance(v, (int, float)):
        return v
    if t == "pct100":            # 0.048 -> 4.8 (si da 0.0 y no es 0, 2 decimales)
        r = redondear(v * 100, 1)
        return redondear(v * 100, 2) if r == 0 and v > 0 else r
    if t == "pct100_2":          # 0.02168 -> 2.17 (si ya viene x100, se deja)
        return redondear(v, 2) if v > 1 else redondear(v * 100, 2)
    if t == "pct_a_fraccion":    # 10.69 -> 0.1069
        return redondear(v / 100, 4) if v > 1 else redondear(v, 4)
    if t == "red2":
        return redondear(v, 2)
    if t == "red4":
        return redondear(v, 4)
    if t == "x100_red2":
        return redondear(v * 100, 2)
    if t == "miles_coma":        # 123400 -> '123,4'
        return f"{redondear(v / 1000, 1):.1f}".replace(".", ",")
    raise ValueError(f"transformacion desconocida: {t}")


# ------------------------------------------------------- ubicar los archivos
def carpeta_del_mes(cfg, mes):
    base = ruta(cfg["carpeta_deptos"].format(anio=mes[:4]))
    nombre_mes = MESES[int(mes[4:]) - 1]
    candidatas = []
    for d in glob.glob(os.path.join(base, "*")):
        n = os.path.basename(d).lower()
        if os.path.isdir(d) and re.search(rf"\b{nombre_mes}\b", n):
            anios = re.findall(r"20\d\d", n)
            if not anios or mes[:4] in anios:
                candidatas.append(d)
    if not candidatas:
        return None
    return max(candidatas, key=os.path.getmtime)


def archivo_depto(carpeta, mapeo):
    def buscar(patron):
        hits = []
        for f in glob.glob(os.path.join(carpeta, "*")):
            n = os.path.basename(f).lower()
            if n.startswith("~$") or not fnmatch.fnmatch(n, patron.lower()):
                continue
            if fnmatch.fnmatch(n, mapeo.get("excluir", "~$*").lower()):
                continue
            hits.append(f)
        return max(hits, key=os.path.getmtime) if hits else None
    return buscar(mapeo["archivo"]) or (buscar(mapeo["archivo_alternativo"]) if mapeo.get("archivo_alternativo") else None)


def monthly_report_base(cfg, mes):
    """El MonthlyReport de Juan de ese mes si existe; si no, el del mes anterior."""
    fmt = ruta(cfg["carpeta_format"])
    def buscar(m):
        hits = [f for f in glob.glob(os.path.join(fmt, f"MonthlyReport_SAR_{m}*.xlsx"))
                if "auto" not in os.path.basename(f).lower() and "copia" not in f.lower()
                and not os.path.basename(f).startswith("~$")]
        return max(hits, key=os.path.getmtime) if hits else None
    propio = buscar(mes)
    if propio:
        return propio, False
    previo = buscar(mes_anterior(mes))
    if previo:
        return previo, True
    raise FileNotFoundError(f"No encontre MonthlyReport_SAR_{mes} ni {mes_anterior(mes)} en {fmt}")


# -------------------------------------------------- ubicar la columna del mes
def col_depto(ws, regla, mes):
    modo = regla["modo"]
    if modo == "fecha_en_fila":
        y, m = int(mes[:4]), int(mes[4:])
        hallado = None
        for c in range(1, ws.max_column + 1):
            v = ws.cell(regla["fila"], c).value
            if isinstance(v, (dt.datetime, dt.date)) and v.year == y and v.month == m:
                hallado = c
        return hallado
    if modo == "ultima_llena":
        ult = None
        for c in range(1, ws.max_column + 1):
            if ws.cell(regla["fila"], c).value not in (None, ""):
                ult = c
        return ult
    if modo == "mes_desde":
        return col_num(regla["col"]) + meses_entre(regla["mes"], mes) * regla.get("paso", 1)
    raise ValueError(f"modo de columna desconocido: {modo}")


def col_mr(regla, mes):
    if regla["modo"] != "mes_desde":
        raise ValueError(f"modo de columna MR desconocido: {regla['modo']}")
    return col_num(regla["col"]) + meses_entre(regla["mes"], mes) * regla.get("paso", 1)


def depto_cargo(wb, hoja_depto, cond, mes):
    """True si el depto ya tiene dato en la fila/columna del mes (sirve para saber si 'ya esta')."""
    ws = wb[nombre_hoja(wb.sheetnames, cond.get("hoja_depto", hoja_depto))]
    c = col_depto(ws, cond["col_depto"], mes)
    return bool(c) and ws.cell(cond["fila"], c).value not in (None, "")


# -------------------------------------------------- leer lo que hay que escribir
class Plan:
    def __init__(self):
        self.escrituras = []   # (hoja_mr, celda, valor, origen, pisar_formula)
        self.revisar = []      # (hoja_mr, celda, motivo)
        self.avisos = []
        self.ops = []          # operaciones especiales a correr despues
        self.bloques_cargados = 0


def armar_plan(mapeo, archivo, mes):
    wb = openpyxl.load_workbook(archivo, data_only=True)
    p = Plan()
    try:
        for b in mapeo.get("bloques", []):
            ws = wb[nombre_hoja(wb.sheetnames, b["hoja_depto"])]
            if b.get("solo_si") and not depto_cargo(wb, b["hoja_depto"], b["solo_si"], mes):
                p.avisos.append(f"{b['nombre']}: el depto todavia no cargo {MESES[int(mes[4:]) - 1]}")
                continue
            cd = col_depto(ws, b["col_depto"], mes)
            if not cd:
                p.avisos.append(f"{b['nombre']}: el depto todavia no tiene la columna de {MESES[int(mes[4:]) - 1]}")
                continue
            cm = col_mr(b["col_mr"], mes)
            n = 0
            for fila in b["filas"]:
                fd, fm = fila[0], fila[1]
                t = fila[2] if len(fila) > 2 else None
                v = ws.cell(fd, cd).value
                if v in (None, ""):
                    continue
                celda = f"{col_letra(cm)}{fm}"
                p.escrituras.append((b["hoja_mr"], celda, transformar(v, t),
                                     f"{ws.title}!{col_letra(cd)}{fd}", b.get("pisar_formulas", False)))
                motivo = b.get("revisar", {}).get(str(fm))
                if motivo:
                    p.revisar.append((b["hoja_mr"], celda, motivo))
                n += 1
            if n == 0:
                p.avisos.append(f"{b['nombre']}: la columna de {MESES[int(mes[4:]) - 1]} esta vacia")
            else:
                p.bloques_cargados += 1
                p.ops.extend(b.get("ops_si_carga", []))

        cond = mapeo.get("celdas_solo_si")
        if mapeo.get("celdas") and (not cond or depto_cargo(wb, cond["hoja_depto"], cond, mes)):
            for c in mapeo["celdas"]:
                ws = wb[nombre_hoja(wb.sheetnames, c["hoja_depto"])]
                v = ws[c["depto"]].value
                if v in (None, ""):
                    continue
                p.escrituras.append((c["hoja_mr"], c["mr"], transformar(v, c.get("transf")),
                                     f"{ws.title}!{c['depto']}", c.get("pisar_formulas", False)))

        for pa in mapeo.get("paretos", []):
            if pa.get("solo_si") and not depto_cargo(wb, pa["hoja_depto"], pa["solo_si"], mes):
                continue
            ws = wb[nombre_hoja(wb.sheetnames, pa["hoja_depto"])]
            c1, r1, c2, r2 = range_boundaries(pa["rango_depto"])
            items = []
            for r in range(r1, r2 + 1):
                nom, val = ws.cell(r, c1).value, ws.cell(r, c1 + 1).value
                if nom not in (None, "") and val not in (None, "", 0):
                    items.append((str(nom).strip(), val, f"{ws.title}!{col_letra(c1)}{r}"))
            mc, mr0, _, _ = range_boundaries(pa["celda_mr"] + ":" + pa["celda_mr"])
            acum = 0
            for i in range(pa["filas_mr"]):
                fila = mr0 + i
                if i < len(items):
                    nom, val, org = items[i]
                    if pa.get("traducir"):
                        trad = mapeo.get("glosario", {}).get(nom)
                        if trad is None:
                            p.revisar.append((pa["hoja_mr"], f"{col_letra(mc)}{fila}",
                                              f"Defecto nuevo, falta traducir al ingles: '{nom}'"))
                        nom = trad or nom
                    acum += a_numero(val) or 0
                    vals = [nom, val] + ([acum] if pa.get("acumulado") else [])
                else:
                    org = f"{pa['nombre']} (fila vacia)"
                    vals = [None, 0] + ([None] if pa.get("acumulado") else [])  # Juan deja 0 en las filas vacias
                for j, v in enumerate(vals):
                    p.escrituras.append((pa["hoja_mr"], f"{col_letra(mc + j)}{fila}", v, org, False))
    finally:
        wb.close()
    return p


# --------------------------------------------------- Excel (COM)
class Excel:
    def __init__(self):
        import pythoncom, win32com.client
        pythoncom.CoInitialize()
        self.app = win32com.client.DispatchEx("Excel.Application")
        self.app.Visible = False
        self.app.DisplayAlerts = False
        self.app.AskToUpdateLinks = False
        self.app.ScreenUpdating = False
        self.app.EnableEvents = False

    def abrir(self, archivo, solo_lectura=False):
        if solo_lectura:
            # se abre una copia en una carpeta corta: no bloquea el archivo del depto
            # y evita el limite de 218 caracteres de ruta de Excel
            tmp = os.path.join(os.environ.get("TEMP", AQUI), "geppo_auto")
            os.makedirs(tmp, exist_ok=True)
            copia = os.path.join(tmp, f"depto_{int(time.time() * 1000)}{os.path.splitext(archivo)[1]}")
            shutil.copy2(archivo, copia)
            archivo = copia
        return self.app.Workbooks.Open(archivo, UpdateLinks=0, ReadOnly=solo_lectura)

    def cerrar(self):
        try:
            for wb in list(self.app.Workbooks):
                wb.Close(SaveChanges=False)
        finally:
            self.app.Quit()


def hojas_com(wb):
    return [wb.Worksheets(i).Name for i in range(1, wb.Worksheets.Count + 1)]


def hoja_com(wb, prefijo):
    return wb.Worksheets(nombre_hoja(hojas_com(wb), prefijo))


def a_com(v):
    if isinstance(v, dt.date) and not isinstance(v, dt.datetime):
        return dt.datetime(v.year, v.month, v.day)
    return v


REF = re.compile(r"(?<![A-Za-z_\d!'\"])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\w(])")


def correr_refs(formula, desde, hasta, delta=None, de_col=None, a_col=None):
    """Corre referencias de columna dentro de una formula.
    Con delta: corre todas las columnas entre 'desde' y 'hasta'. Con de_col/a_col: reemplaza una columna por otra."""
    if not formula.startswith("="):
        return formula
    def rep(m):
        c = col_num(m.group(2))
        if de_col is not None:
            if c != de_col:
                return m.group(0)
            nueva = a_col
        else:
            if not (desde <= c <= hasta):
                return m.group(0)
            nueva = c + delta
        return f"{m.group(1)}{col_letra(nueva)}{m.group(3)}{m.group(4)}"
    return REF.sub(rep, formula)


def op_ventana_graficos_produccion(wb, mes):
    """Tabla de los graficos de arriba de 'Production': Juan corre las referencias 1 columna por mes.
    Se calcula a partir de la celda ancla F18 (siempre apunta a mes-7 de la serie BG=abr-2025)."""
    ws = hoja_com(wb, "Production")
    desde, hasta = col_num("BG"), col_num("CD")
    destino = desde + meses_entre("2025-04", mes) - 7
    f18 = ws.Range("F18").Formula
    m = REF.search(f18) if f18.startswith("=") else None
    if not m:
        return "ventana de graficos: F18 no tiene formula, no se toco"
    delta = destino - col_num(m.group(2))
    if delta == 0:
        return "ventana de graficos: ya estaba en el mes"
    n = 0
    for rango in ["F18:AX18", "F20:AJ22", "F23:AL23", "F34:AH38"]:  # la fila 39 Juan no la corre
        for cel in ws.Range(rango):
            f = cel.Formula
            if isinstance(f, str) and f.startswith("="):
                nf = correr_refs(f, desde, hasta, delta=delta)
                if nf != f:
                    cel.Formula = nf
                    n += 1
    return f"ventana de graficos corrida {delta:+d} columna(s), {n} formulas"


def op_cuadritos_mes_actual_produccion(wb, mes):
    """Cuadritos 'Current month / Cumulative' de Production: apuntan a la columna del mes anterior -> mes actual."""
    ws = hoja_com(wb, "Production")
    actual = col_num("BG") + meses_entre("2026-04", mes)
    n = 0
    for cel in ws.Range("BN44:BS180"):
        f = cel.Formula
        if isinstance(f, str) and f.startswith("="):
            nf = correr_refs(f, 0, 0, de_col=actual - 1, a_col=actual)
            if nf != f:
                cel.Formula = nf
                n += 1
    return f"cuadritos 'Current month' actualizados: {n} formulas"


OPS = {"ventana_graficos_produccion": op_ventana_graficos_produccion,
       "cuadritos_mes_actual_produccion": op_cuadritos_mes_actual_produccion}


def aplicar(wb_mr, mapeo, plan, mes):
    hechas, saltadas = [], 0
    for hpref, celda, v, origen, pisar in plan.escrituras:
        ws = hoja_com(wb_mr, hpref)
        rg = ws.Range(celda)
        if rg.HasFormula and not pisar:
            saltadas += 1
            calc = rg.Value
            if isinstance(v, (int, float)) and isinstance(calc, (int, float)) and abs(calc - v) > max(0.01, 0.005 * abs(v)):
                plan.revisar.append((ws.Name, celda, f"El informe tiene una formula que da {round(calc, 4)}, "
                                                     f"pero el depto dice {v} ({origen}). No se toco."))
            continue
        rg.Value = a_com(v)
        hechas.append((ws.Name, celda, v, origen))

    if plan.bloques_cargados:
        for regla in mapeo.get("formulas_de_col_anterior", []):
            ws = hoja_com(wb_mr, regla["hoja_mr"])
            c = col_mr(regla["col_mr"], mes)
            for f in regla["filas"]:
                dst, src = ws.Cells(f, c), ws.Cells(f, c - 1)
                if dst.Formula in ("", None) and src.HasFormula:
                    dst.FormulaR1C1 = src.FormulaR1C1
                    hechas.append((ws.Name, dst.Address.replace("$", ""), dst.Formula, "formula copiada de la columna anterior"))
        for regla in mapeo.get("valores_de_col_anterior", []):
            ws = hoja_com(wb_mr, regla["hoja_mr"])
            c = col_mr(regla["col_mr"], mes)
            for f in regla["filas"]:
                dst, src = ws.Cells(f, c), ws.Cells(f, c - 1)
                if dst.Formula in ("", None) and src.Value is not None:
                    dst.Value = src.Value
                    hechas.append((ws.Name, dst.Address.replace("$", ""), src.Value, "meta copiada de la columna anterior"))
    for op in dict.fromkeys(plan.ops):
        plan.avisos.append(OPS[op](wb_mr, mes))
    return hechas, saltadas


def titulos_del_mes(wb_mr, mes):
    """Lo que Juan cambia en todos los meses: titulo de Comments y de A&F."""
    mn = MONTHS[int(mes[4:]) - 1]
    cambios = []
    try:
        ws = hoja_com(wb_mr, "Comments")
        v = ws.Range("L1").Value
        if isinstance(v, str) and re.search(r"SAR\s+\w+\s+\d{4}", v):
            nv = re.sub(r"SAR\s+\w+\s+\d{4}", f"SAR {mn} {mes[:4]}", v)
            if nv != v:
                ws.Range("L1").Value = nv
                cambios.append(("Comments", "L1", nv))
    except KeyError:
        pass
    try:
        ws = hoja_com(wb_mr, "JPN.Ⅴ")
        v = ws.Range("C4").Value
        if isinstance(v, str) and "for" in v:
            nv = re.sub(r"for\s+\w+", f"for {mn}", v, count=1)
            if nv != v:
                ws.Range("C4").Value = nv
                cambios.append(("JPN.Ⅴ A&F", "C4", nv))
    except KeyError:
        pass
    return cambios


# --------------------------------------------------- hoja AUTO de cada depto
def hoja_auto(xl, wb_mr, mapeo, archivo, mes, hechas, saltadas, plan):
    nombre = f"AUTO {mapeo['nombre']}"[:31]
    for n in hojas_com(wb_mr):
        if n == nombre:
            wb_mr.Worksheets(n).Delete()
    ws = wb_mr.Worksheets.Add(After=wb_mr.Worksheets(wb_mr.Worksheets.Count))
    ws.Name = nombre
    ws.Tab.Color = 0x00C0FF
    ws.Columns("A").ColumnWidth = 30
    ws.Columns("B").ColumnWidth = 12
    ws.Columns("C").ColumnWidth = 60
    ws.Columns("D").ColumnWidth = 40

    fila = [1]
    def esc(*vals, negrita=False, color=None):
        for i, v in enumerate(vals):
            c = ws.Cells(fila[0], i + 1)
            c.Value = a_com(v) if not isinstance(v, str) else ("'" + v if v.startswith("=") else v)
            if negrita:
                c.Font.Bold = True
            if color:
                c.Interior.Color = color
        fila[0] += 1

    mod = dt.datetime.fromtimestamp(os.path.getmtime(archivo))
    esc(f"Carga automatica {mapeo['nombre']} - {MESES[int(mes[4:]) - 1].upper()} {mes[:4]}", negrita=True)
    esc("Archivo del depto:", os.path.basename(archivo))
    esc("Guardado por el depto:", f"{mod:%d/%m/%Y %H:%M}")
    esc("Cargado al MonthlyReport:", f"{dt.datetime.now():%d/%m/%Y %H:%M}")
    esc("BORRAR ESTA HOJA ANTES DE ENVIAR EL INFORME", color=0x99FFFF, negrita=True)
    fila[0] += 1

    if plan.revisar:
        esc("PARA REVISAR", negrita=True, color=0x99FFFF)
        for h, c, motivo in plan.revisar:
            esc(h, c, motivo)
        fila[0] += 1
    if plan.avisos:
        esc("AVISOS", negrita=True)
        for a in plan.avisos:
            esc("", "", a)
        fila[0] += 1

    # textos del depto para que Juan los redacte en el informe
    textos = []
    wb_d = None
    try:
        if mapeo.get("textos") or mapeo.get("fotos") or mapeo.get("textos_de_cuadros"):
            wb_d = xl.abrir(archivo, solo_lectura=True)
        for t in mapeo.get("textos", []):
            try:
                v = hoja_com(wb_d, t["hoja_depto"]).Range(t["celda"]).Value
            except Exception:
                v = None
            if isinstance(v, str) and v.strip() and len(v.strip()) > 3:
                textos.append((t["seccion"], v.strip()))
        for t in mapeo.get("textos_de_cuadros", []):
            try:
                wsd = hoja_com(wb_d, t["hoja_depto"])
                partes = []
                for i in range(1, wsd.Shapes.Count + 1):
                    s = wsd.Shapes(i)
                    try:
                        if s.HasTextFrame and s.TextFrame2.HasText:
                            partes.append(s.TextFrame2.TextRange.Text.strip())
                    except Exception:
                        pass
                if partes:
                    textos.append((t["seccion"], "\n---\n".join(partes)))
            except KeyError:
                pass
        if textos:
            esc("TEXTOS DEL DEPTO (para redactar los comentarios del informe)", negrita=True)
            for sec, txt in textos:
                ws.Cells(fila[0], 1).Value = sec
                c = ws.Range(ws.Cells(fila[0], 3), ws.Cells(fila[0], 4))
                c.Merge()
                c.WrapText = True
                c.VerticalAlignment = -4160  # arriba
                ws.Cells(fila[0], 3).Value = txt.replace("\r", "")
                ws.Rows(fila[0]).RowHeight = min(409, 15 * (txt.count("\n") + 1 + len(txt) // 90))
                fila[0] += 1
            fila[0] += 1

        # fotos del depto
        total_fotos = 0
        for fr in mapeo.get("fotos", []):
            try:
                wsd = hoja_com(wb_d, fr["hoja_depto"])
            except KeyError:
                continue
            area = wsd.Range(fr["area_depto"])
            fotos = []
            for i in range(1, wsd.Shapes.Count + 1):
                s = wsd.Shapes(i)
                try:
                    if s.Type in (13, 11) and xl.app.Intersect(s.TopLeftCell, area) is not None:
                        fotos.append(s)
                except Exception:
                    pass
            if not fotos:
                continue
            esc(fr["seccion"], negrita=True)
            top = ws.Cells(fila[0], 1).Top
            left = ws.Cells(fila[0], 1).Left
            alto_max = 0
            for s in sorted(fotos, key=lambda s: (s.Top, s.Left)):
                s.Copy()
                ws.Paste(ws.Cells(fila[0], 1))
                nueva = ws.Shapes(ws.Shapes.Count)
                if nueva.Width > 380:
                    nueva.LockAspectRatio = True
                    nueva.Width = 380
                if left + nueva.Width > 1200:  # nueva fila de fotos
                    top += alto_max + 10
                    left, alto_max = ws.Cells(1, 1).Left, 0
                nueva.Top, nueva.Left = top, left
                left += nueva.Width + 10
                alto_max = max(alto_max, nueva.Height)
                total_fotos += 1
            # bajar las filas hasta debajo de las fotos
            while ws.Cells(fila[0], 1).Top < top + alto_max + 10:
                fila[0] += 1
        if total_fotos:
            plan.avisos.append(f"{total_fotos} fotos del depto copiadas en la hoja '{nombre}'")
    finally:
        if wb_d is not None:
            copia = wb_d.FullName
            wb_d.Close(SaveChanges=False)
            try:
                os.remove(copia)
            except OSError:
                pass

    fila[0] += 1
    esc("CELDAS CARGADAS EN EL INFORME", negrita=True)
    esc("Hoja", "Celda", "Valor", "De donde salio", negrita=True)
    for h, c, v, org in hechas:
        esc(h, c, v if not isinstance(v, float) else round(v, 6), org)
    if saltadas:
        esc("", "", f"{saltadas} celdas no se tocaron porque en el informe tienen formula")
    return len(textos)


# ------------------------------------------------------------- estado
def leer_estado():
    p = os.path.join(AQUI, "estado.json")
    if os.path.exists(p):
        with open(p, encoding="utf-8") as f:
            return json.load(f)
    return {}


def guardar_estado(e):
    with open(os.path.join(AQUI, "estado.json"), "w", encoding="utf-8") as f:
        json.dump(e, f, indent=1, ensure_ascii=False)


def mapeos():
    out = []
    for p in sorted(glob.glob(os.path.join(AQUI, "mapeos", "*.json"))):
        with open(p, encoding="utf-8") as f:
            out.append(json.load(f))
    return out


# ------------------------------------------------------------- proceso
def procesar(cfg, mes, forzar=False):
    estado = leer_estado()
    salida = os.path.join(ruta(cfg["carpeta_salida"]), cfg["nombre_salida"].format(mes=mes))
    em = estado.setdefault(salida, {})
    carpeta = carpeta_del_mes(cfg, mes)
    if not carpeta:
        log(f"{mes}: todavia no existe la carpeta de {MESES[int(mes[4:]) - 1].upper()} de los departamentos")
        return False

    pendientes = []
    for m in mapeos():
        f = archivo_depto(carpeta, m)
        if not f:
            continue
        mt = os.path.getmtime(f)
        firma = f"{os.path.basename(f)}|{mt:.0f}|{os.path.getsize(f)}"
        if not forzar and em.get(m["key"]) == firma:
            continue
        if time.time() - mt < SEGUNDOS_ESTABLE:
            log(f"[{m['nombre']}] el archivo se esta guardando, lo tomo en la proxima vuelta")
            continue
        pendientes.append((m, f, firma))
    if not pendientes:
        return False

    os.makedirs(os.path.dirname(salida), exist_ok=True)
    nuevo = not os.path.exists(salida)
    if nuevo:
        base, es_anterior = monthly_report_base(cfg, mes)
        shutil.copy2(base, salida)
        log(f"Cree {os.path.basename(salida)} a partir de {os.path.basename(base)}")

    xl = Excel()
    try:
        try:
            wb_mr = xl.abrir(salida)
        except Exception as e:
            log(f"No pude abrir {os.path.basename(salida)} ({e}). Reintento en la proxima vuelta.")
            return False
        if wb_mr.ReadOnly:
            log(f"{os.path.basename(salida)} lo tiene abierto alguien. Reintento en la proxima vuelta.")
            return False
        if nuevo:
            for h, c, v in titulos_del_mes(wb_mr, mes):
                log(f"titulo actualizado {h}!{c}: {v}")
        hubo = False
        for m, f, firma in pendientes:
            try:
                plan = armar_plan(m, f, mes)
                if not plan.escrituras:
                    log(f"[{m['nombre']}] todavia no cargo {MESES[int(mes[4:]) - 1]} ({os.path.basename(f)})")
                    em[m["key"]] = firma
                    continue
                hechas, saltadas = aplicar(wb_mr, m, plan, mes)
                try:
                    ntx = hoja_auto(xl, wb_mr, m, f, mes, hechas, saltadas, plan)
                except Exception as e:
                    ntx = 0
                    log(f"[{m['nombre']}] no pude armar la hoja AUTO (los numeros si se cargaron): {e}")
                log(f"[{m['nombre']}] cargue {len(hechas)} celdas, {len(plan.revisar)} para revisar, "
                    f"{ntx} textos; desde {os.path.basename(f)}")
                for a in plan.avisos:
                    log(f"[{m['nombre']}]   {a}")
                em[m["key"]] = firma
                hubo = True
            except Exception as e:
                log(f"[{m['nombre']}] ERROR: {e}\n{traceback.format_exc()}")
        if hubo or nuevo:
            wb_mr.Worksheets(1).Activate()
            wb_mr.Save()
    finally:
        xl.cerrar()
    guardar_estado(estado)
    return True


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--mes", help="YYYYMM (por defecto el mes vencido)")
    ap.add_argument("--vigilar", action="store_true")
    ap.add_argument("--forzar", action="store_true")
    ap.add_argument("--produccion", action="store_true", help="usar las rutas reales de Z:")
    ap.add_argument("--config", help="otro archivo de configuracion (para pruebas)")
    a = ap.parse_args()
    cfg = cargar_config(a.produccion, a.config)

    if not a.vigilar:
        mes = a.mes or mes_vencido()
        if not procesar(cfg, mes, a.forzar):
            log(f"{mes}: nada nuevo para cargar")
        return

    log("Vigilando las carpetas de los departamentos (Ctrl+C para salir)")
    while True:
        try:
            procesar(cfg, a.mes or mes_vencido(), a.forzar)
            a.forzar = False
        except Exception as e:
            log(f"ERROR: {e}\n{traceback.format_exc()}")
        time.sleep(60 * cfg.get("minutos_entre_revisiones", 5))


if __name__ == "__main__":
    main()
