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
escribe en "MonthlyReport_SAR_YYYYMM_AUTO.xlsx", completando las mismas hojas del informe
(numeros, cuadros de texto y fotos). Lo que hay que revisar queda en
"MonthlyReport_SAR_YYYYMM_AUTO - revisar.txt", al lado del informe.
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
    with open(archivo or os.path.join(AQUI, "config.json"), encoding="utf-8-sig") as f:
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
                    # Juan a veces deja cuentas tipo =2998+349: se reemplazan por el valor del depto
                    p.escrituras.append((pa["hoja_mr"], f"{col_letra(mc + j)}{fila}", v, org, True))
    finally:
        wb.close()
    # celdas que no salen del archivo del depto: se avisan para que Juan las complete
    if p.bloques_cargados:
        for a in mapeo.get("a_mano", []):
            celda = a["celda"] if "celda" in a else f"{col_letra(col_mr(a['col_mr'], mes))}{a['fila']}"
            p.revisar.append((a["hoja_mr"], celda, "Completar a mano: " + a["que"]))
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
    for rango in ["F18:AX18", "F20:AJ22", "F23:AL23", "F34:AH39"]:
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


# --------------------------------------------------- textos y fotos en el informe
ABR = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sept", "Oct", "Nov", "Dec"]
_MES_RE = r"(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?"
PERIODO_RE = re.compile(_MES_RE + r"\s*\d{1,2}(?:st|nd|rd|th)?\s*[-~]\s*" + _MES_RE + r"\s*\d{1,2}(?:st|nd|rd|th)?", re.I)


def periodo(mes):
    """Periodo que usa Juan: del 21 del mes anterior al 20 del mes. Ej: 'Sept 21st - Oct 20th'."""
    m = int(mes[4:])
    return f"{ABR[(m - 2) % 12]} 21st - {ABR[m - 1]} 20th"


def texto_de(shape):
    for leer in (lambda: shape.TextFrame2.TextRange.Text, lambda: shape.TextFrame.Characters().Text):
        try:
            t = leer()
            if t:
                return t
        except Exception:
            pass
    return ""


def buscar_cuadro(ws, regla):
    """El cuadro de texto del informe: por la celda donde arranca (cuadro_en) y/o por su nombre."""
    mejor = None
    for i in range(1, ws.Shapes.Count + 1):
        s = ws.Shapes(i)
        if s.Type in (3, 13, 11):
            continue
        try:
            if regla.get("cuadro_en") and s.TopLeftCell.Address.replace("$", "") != regla["cuadro_en"]:
                continue
            if regla.get("cuadro_nombre") and s.Name != regla["cuadro_nombre"]:
                continue
            s.TextFrame2.TextRange  # si la forma no tiene texto, esto falla
        except Exception:
            continue
        if mejor is None or s.Width * s.Height > mejor.Width * mejor.Height:
            mejor = s
    if mejor is not None or not regla.get("empieza_con") or not regla.get("cuadro_en"):
        return mejor
    # Juan a veces mueve los cuadros unas filas: se busca el mas cercano que empiece con el rotulo esperado
    ancla = ws.Range(regla["cuadro_en"])
    mejor_dist = None
    for i in range(1, ws.Shapes.Count + 1):
        s = ws.Shapes(i)
        if s.Type in (3, 13, 11):
            continue
        t = texto_de(s).lstrip()
        if not t or not re.match(regla["empieza_con"], t, re.I):
            continue
        try:
            celda = s.TopLeftCell
            dist = abs(celda.Row - ancla.Row) + abs(celda.Column - ancla.Column) / 4
        except Exception:
            continue
        if dist <= 40 and (mejor_dist is None or dist < mejor_dist):
            mejor, mejor_dist = s, dist
    return mejor


def limpiar(t, regla):
    t = (t or "").replace("\r\n", "\n").replace("\r", "\n")
    t = "\n".join(l for l in t.split("\n") if not re.match(r"^\s*\*?\s*Periodo de consideraci", l, re.I))
    if regla.get("despues_de"):
        m = re.search(regla["despues_de"], t, re.I)
        t = t[m.end():] if m else ""
    if regla.get("antes_de"):
        m = re.search(regla["antes_de"], t, re.I)
        t = t[:m.start()] if m else t
    for a, b in regla.get("reemplazos", {}).items():
        t = re.sub(a, b, t)
    t = re.sub(r"[ \t]*\t[ \t]*", " ", t)  # las tabulaciones del depto dejan huecos grandes en el cuadro
    t = re.sub(r"[ \t]+\n", "\n", t)
    t = re.sub(r"\n{3,}", "\n\n", t)
    return t.strip()


def poner_textos(wb_d, wb_mr, mapeo, mes, plan):
    """Pasa los textos del depto a los cuadros de comentarios del informe."""
    n = 0
    vars_ = {"periodo": periodo(mes), "mes_en": MONTHS[int(mes[4:]) - 1], "mes_abr": ABR[int(mes[4:]) - 1]}
    for r in mapeo.get("cuadros", []):
        try:
            ws = hoja_com(wb_mr, r["hoja_mr"])
        except KeyError:
            continue
        caja = buscar_cuadro(ws, r)
        if caja is None:
            plan.avisos.append(f"no encontre el cuadro de texto {r.get('cuadro_nombre', '')} en {r['hoja_mr']}!{r.get('cuadro_en', '')}")
            continue
        modo = r.get("modo", "desde_celda")
        if modo == "vaciar":  # el depto no manda ese texto: queda solo el rotulo para que Juan lo complete
            nuevo = r.get("encabezado", "").format(**vars_)
        elif modo == "solo_periodo":
            viejo = texto_de(caja).replace("\r", "\n")
            nuevo = PERIODO_RE.sub(vars_["periodo"], viejo, count=1)
        else:
            wsd = hoja_com(wb_d, r["hoja_depto"])
            if modo == "desde_cuadro":
                try:
                    fuente = texto_de(wsd.Shapes(r["cuadro_depto"]))
                except Exception:
                    fuente = ""
            else:
                celdas = r.get("celdas") or [r["celda"]]
                fuente = "\n".join(str(wsd.Range(c).Value) for c in celdas if wsd.Range(c).Value not in (None, ""))
            cuerpo = limpiar(fuente, r)
            if r.get("plantilla"):
                m = re.search(r["extraer"], cuerpo)
                if not m:
                    continue
                nuevo = r["plantilla"].format(x=m.group(1).strip(), **vars_)
            else:
                if not cuerpo and not r.get("aunque_vacio"):
                    continue
                nuevo = r.get("encabezado", "").format(**vars_) + cuerpo
        if nuevo.strip() and nuevo != texto_de(caja).replace("\r", "\n"):
            # el texto nuevo toma el formato de la primera letra real del cuadro
            # (si el cuadro arranca con una linea vacia en letra chica, no se usa esa)
            viejo = texto_de(caja)
            idx = next((k for k, ch in enumerate(viejo) if not ch.isspace()), 0)
            fmt = None
            try:
                f0 = caja.TextFrame.Characters(idx + 1, 1).Font
                fmt = (f0.Size, f0.Name, f0.Bold, f0.Color)
            except Exception:
                pass
            caja.TextFrame2.TextRange.Text = nuevo.replace("\n", "\r")
            if fmt:
                try:
                    f1 = caja.TextFrame.Characters().Font
                    f1.Size, f1.Name, f1.Bold, f1.Color = fmt
                except Exception:
                    pass
            n += 1
            if r.get("traducir"):
                plan.revisar.append((ws.Name, f"cuadro en {r.get('cuadro_en', r.get('cuadro_nombre'))}",
                                     "Texto del depto en castellano: pasarlo a ingles."))
    return n


def copiar_fotos(xl, wb_d, wb_mr, mapeo, plan):
    """Borra las fotos del mes anterior en su lugar del informe y pone las del depto, acomodadas en ese espacio."""
    total = 0
    for r in mapeo.get("fotos", []):
        try:
            wsd = hoja_com(wb_d, r["hoja_depto"])
            wsm = hoja_com(wb_mr, r["hoja_mr"])
        except KeyError:
            continue
        area_d = wsd.Range(r["area_depto"])
        zona = wsm.Range(r.get("zona_mr", r["area_mr"]))
        fotos = []
        for i in range(1, wsd.Shapes.Count + 1):
            s = wsd.Shapes(i)
            try:
                es_foto = s.Type in (13, 11) or (s.Type == 6 and any(s.GroupItems(k).Type in (13, 11)
                                                                     for k in range(1, s.GroupItems.Count + 1)))
                if es_foto and xl.app.Intersect(s.TopLeftCell, area_d) is not None:
                    fotos.append(s)
            except Exception:
                pass
        # Las fotos nuevas van donde estaban las del mes anterior (Juan a veces mueve el diseno):
        # se toma el rectangulo que ocupaban las fotos viejas dentro de la zona.
        viejas = []
        for i in range(1, wsm.Shapes.Count + 1):
            s = wsm.Shapes(i)
            try:
                es_foto = s.Type in (13, 11) or (s.Type == 6 and any(s.GroupItems(k).Type in (13, 11)
                                                                     for k in range(1, s.GroupItems.Count + 1)))
                if es_foto and xl.app.Intersect(s.TopLeftCell, zona) is not None:
                    viejas.append(s)
            except Exception:
                pass
        if viejas:
            x0 = min(s.Left for s in viejas); y0 = min(s.Top for s in viejas)
            x1 = max(s.Left + s.Width for s in viejas); y1 = max(s.Top + s.Height for s in viejas)
        else:
            a = wsm.Range(r["area_mr"])
            x0, y0, x1, y1 = a.Left, a.Top, a.Left + a.Width, a.Top + a.Height
        # se borran las fotos viejas y sus marcas chicas (flechas, circulos, rotulos) alrededor;
        # los cuadros de comentarios grandes no se tocan
        m = 20
        claves_viejas = {(v.Name, round(v.Left), round(v.Top)) for v in viejas}
        for i in range(wsm.Shapes.Count, 0, -1):
            s = wsm.Shapes(i)
            try:
                if s.Type in (3, 4):
                    continue
                dentro = (s.Left >= x0 - m and s.Top >= y0 - m and s.Left + s.Width <= x1 + m
                          and s.Top + s.Height <= y1 + m)
                chica = s.Width * s.Height < 0.35 * (x1 - x0) * (y1 - y0) and len(texto_de(s)) < 120
                toda = r.get("limpiar_toda_la_zona") and xl.app.Intersect(s.TopLeftCell, zona) is not None
                if (s.Name, round(s.Left), round(s.Top)) in claves_viejas or (dentro and chica) or toda:
                    s.Delete()
            except Exception:
                pass
        if not fotos:
            continue
        fotos.sort(key=lambda s: (round(s.Top / 20), s.Left))
        n = len(fotos)
        cols = n if n <= 3 else -(-n // 2)
        filas = -(-n // cols)
        gap = 4
        cw = ((x1 - x0) - gap * (cols - 1)) / cols
        ch = ((y1 - y0) - gap * (filas - 1)) / filas
        wsm.Activate()
        for i, s in enumerate(fotos):
            antes = {wsm.Shapes(k).ID for k in range(1, wsm.Shapes.Count + 1)}
            # el portapapeles de Windows a veces tarda: se reintenta el copiar/pegar
            for intento in range(4):
                try:
                    s.Copy()
                    time.sleep(0.3)
                    wsm.Paste()
                    break
                except Exception:
                    time.sleep(1.5)
            # la foto pegada es la forma que no estaba antes (no siempre es la ultima de la lista)
            nuevas = [wsm.Shapes(k) for k in range(1, wsm.Shapes.Count + 1) if wsm.Shapes(k).ID not in antes]
            if not nuevas:
                plan.avisos.append(f"no se pudo pegar una foto en {wsm.Name}")
                continue
            nueva = nuevas[0]
            nueva.LockAspectRatio = True
            esc = min(cw / nueva.Width, ch / nueva.Height)
            nueva.Width = nueva.Width * esc
            f, c = divmod(i, cols)
            nueva.Left = x0 + c * (cw + gap) + (cw - nueva.Width) / 2
            nueva.Top = y0 + f * (ch + gap) + (ch - nueva.Height) / 2
            total += 1
    return total


def textos_y_fotos(xl, wb_mr, mapeo, archivo, mes, plan):
    if not (mapeo.get("cuadros") or mapeo.get("fotos")):
        return 0, 0
    wb_d = xl.abrir(archivo, solo_lectura=True)
    nt = nf = 0
    try:
        try:
            nt = poner_textos(wb_d, wb_mr, mapeo, mes, plan)
        except Exception as e:
            plan.avisos.append(f"no pude pasar los textos: {e}")
        try:
            nf = copiar_fotos(xl, wb_d, wb_mr, mapeo, plan)
        except Exception as e:
            plan.avisos.append(f"no pude pasar las fotos: {e}")
    finally:
        copia = wb_d.FullName
        wb_d.Close(SaveChanges=False)
        try:
            os.remove(copia)
        except OSError:
            pass
    return nt, nf


def escribir_revisar(salida, notas, mes):
    """Lista de cosas para revisar, en un .txt al lado del informe (no se agregan hojas al Excel)."""
    p = os.path.splitext(salida)[0] + " - revisar.txt"
    lineas = [f"GEPPO {MESES[int(mes[4:]) - 1].upper()} {mes[:4]} - cosas para revisar en {os.path.basename(salida)}",
              f"Actualizado: {dt.datetime.now():%d/%m/%Y %H:%M}", ""]
    for depto, items in notas.items():
        lineas.append(f"== {depto}")
        lineas += [f"   - {x}" for x in items] or ["   (nada)"]
        lineas.append("")
    with open(p, "w", encoding="utf-8-sig") as f:
        f.write("\n".join(lineas))


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
        with open(p, encoding="utf-8-sig") as f:  # acepta archivos guardados con el Bloc de notas
            out.append(json.load(f))
    return out


# ------------------------------------------------------------- proceso
def mes_siguiente(mes):
    y, m = int(mes[:4]), int(mes[4:])
    return f"{y + 1}01" if m == 12 else f"{y}{m + 1:02d}"


def procesar(cfg, mes, forzar=False):
    estado = leer_estado()
    salida = os.path.join(ruta(cfg["carpeta_salida"]), cfg["nombre_salida"].format(mes=mes))
    em = estado.setdefault(salida, {})
    completo = em.setdefault("_completo", {})
    carpeta = carpeta_del_mes(cfg, mes)
    if not carpeta:
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

    # primero se lee lo que trae cada depto; el informe se abre solo si alguno ya cargo el mes
    con_datos = []
    for m, f, firma in pendientes:
        try:
            plan = armar_plan(m, f, mes)
        except Exception as e:
            log(f"[{m['nombre']}] ERROR leyendo {os.path.basename(f)}: {e}")
            continue
        falta = [a for a in plan.avisos if "todavia no" in a or "esta vacia" in a]
        completo[m["key"]] = bool(plan.bloques_cargados) and not falta
        if not plan.escrituras:
            log(f"[{m['nombre']}] todavia no cargo {MESES[int(mes[4:]) - 1]} ({os.path.basename(f)})")
            em[m["key"]] = firma
        else:
            con_datos.append((m, f, firma, plan))
    if not con_datos:
        guardar_estado(estado)
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
        # si el informe se guardo con varias hojas seleccionadas (agrupadas), Excel no deja
        # agregar ni pegar: se deja seleccionada solo la primera hoja visible
        for i in range(1, wb_mr.Worksheets.Count + 1):
            if wb_mr.Worksheets(i).Visible == -1:
                wb_mr.Worksheets(i).Select(True)
                break
        # hojas AUTO de versiones anteriores del programa: se sacan
        for n in [n for n in hojas_com(wb_mr) if n.startswith("AUTO ")]:
            wb_mr.Worksheets(n).Delete()
        notas = em.setdefault("_revisar", {})
        if nuevo:
            for h, c, v in titulos_del_mes(wb_mr, mes):
                log(f"titulo actualizado {h}!{c}: {v}")
        hubo = False
        for m, f, firma, plan in con_datos:
            try:
                hechas, saltadas = aplicar(wb_mr, m, plan, mes)
                ntx = nfo = 0
                if plan.bloques_cargados:
                    try:
                        ntx, nfo = textos_y_fotos(xl, wb_mr, m, f, mes, plan)
                    except Exception as e:
                        log(f"[{m['nombre']}] no pude pasar textos/fotos (los numeros si se cargaron): {e}")
                log(f"[{m['nombre']}] cargue {len(hechas)} celdas, {ntx} textos, {nfo} fotos, "
                    f"{len(plan.revisar)} para revisar; desde {os.path.basename(f)}")
                notas[m["nombre"]] = [f"{h}!{c}: {motivo}" for h, c, motivo in plan.revisar] + plan.avisos
                for a in plan.avisos:
                    log(f"[{m['nombre']}]   {a}")
                em[m["key"]] = firma
                hubo = True
            except Exception as e:
                log(f"[{m['nombre']}] ERROR: {e}\n{traceback.format_exc()}")
        if hubo or nuevo:
            wb_mr.Worksheets(1).Activate()
            wb_mr.Save()
            escribir_revisar(salida, notas, mes)
    finally:
        xl.cerrar()
    guardar_estado(estado)
    return True


def preparar_mes_siguiente(cfg, mes, hoy=None):
    """Lo que hace Juan: crea la carpeta 'GEPPO DE <Mes> <Anio>' del mes siguiente con los archivos del mes
    que se acaba de cerrar, para que cada depto abra el suyo y agregue el mes.
    Se crea apenas todos los deptos terminaron el mes, o el dia 'dia_limite_carpeta' aunque falte alguno."""
    hoy = hoy or dt.date.today()
    sig = mes_siguiente(mes)
    if carpeta_del_mes(cfg, sig):
        return None
    actual = carpeta_del_mes(cfg, mes)
    if not actual:
        return None
    estado = leer_estado()
    salida = os.path.join(ruta(cfg["carpeta_salida"]), cfg["nombre_salida"].format(mes=mes))
    completo = estado.get(salida, {}).get("_completo", {})
    deptos = [m for m in mapeos() if archivo_depto(actual, m)]
    faltan = [m["nombre"] for m in deptos if not completo.get(m["key"])]
    # fecha tope: el dia_limite del mes siguiente al cerrado (ej. octubre cerrado -> 10 de noviembre)
    tope = dt.date(int(sig[:4]), int(sig[4:]), int(cfg.get("dia_limite_carpeta", 10)))
    if faltan and hoy < tope:
        return None
    base = ruta(cfg["carpeta_deptos"].format(anio=sig[:4]))
    nueva = os.path.join(base, f"GEPPO DE {MESES[int(sig[4:]) - 1].capitalize()} {sig[:4]}")
    os.makedirs(nueva, exist_ok=True)
    copiados = []
    for f in sorted(glob.glob(os.path.join(actual, "*"))):
        n = os.path.basename(f)
        if os.path.isfile(f) and not n.startswith("~$"):
            shutil.copy2(f, os.path.join(nueva, n))
            copiados.append(n)
    log(f"Cree la carpeta {nueva} con {len(copiados)} archivos de {MESES[int(mes[4:]) - 1]}"
        + (f" (todavia faltaba terminar: {', '.join(faltan)})" if faltan else " (todos los deptos terminaron)"))
    return nueva, copiados, faltan


def una_vuelta(cfg, mes_fijo=None, forzar=False):
    meses = [mes_fijo] if mes_fijo else [mes_vencido(), f"{dt.date.today():%Y%m}"]
    hubo = False
    for mes in meses:
        if procesar(cfg, mes, forzar):
            hubo = True
    if not mes_fijo:
        preparar_mes_siguiente(cfg, mes_vencido())
    return hubo


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--mes", help="YYYYMM (por defecto el mes vencido y el mes en curso)")
    ap.add_argument("--vigilar", action="store_true")
    ap.add_argument("--forzar", action="store_true")
    ap.add_argument("--produccion", action="store_true", help="usar las rutas reales de Z:")
    ap.add_argument("--config", help="otro archivo de configuracion (para pruebas)")
    a = ap.parse_args()
    cfg = cargar_config(a.produccion, a.config)

    if not a.vigilar:
        if not una_vuelta(cfg, a.mes, a.forzar):
            log("nada nuevo para cargar")
        return

    log("Vigilando las carpetas de los departamentos (Ctrl+C para salir)")
    while True:
        try:
            una_vuelta(cfg, a.mes, a.forzar)
            a.forzar = False
        except Exception as e:
            log(f"ERROR: {e}\n{traceback.format_exc()}")
        time.sleep(60 * cfg.get("minutos_entre_revisiones", 10))


if __name__ == "__main__":
    main()
