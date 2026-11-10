"""
actualizar_certificaciones.py
Genera data/certificaciones_<obra>.json para certificaciones.html.

Lee, por cada obra:
  - Certificado_<OBRA>.xlsm   -> resúmenes ya calculados (Resumen General, Resumen Ajuste,
                                 Saldo de Obra, C I Comparativa, Metodología del último cert.,
                                 REDETERMINACIONES solo para el mes de salto)
  - 00 - Seguimiento certificacion.xlsx -> TABLA MASTER (fechas del circuito y pagos)
                                           y SALDO (fin de obra según plan vigente)

No recalcula nada del certificado: toma los valores tal como quedaron guardados en el Excel.
Las tablas se ubican por sus títulos, no por celdas fijas, para que sirva en todas las obras.

Uso:
    python actualizar_certificaciones.py            -> todas las obras configuradas
    python actualizar_certificaciones.py mechita    -> solo una
"""
import sys
import re
import json
import math
import statistics
import unicodedata
from datetime import datetime, date, timedelta
from pathlib import Path

from openpyxl import load_workbook

# ============================================================
# CONFIGURACIÓN — ajustar rutas a tu PC
# ============================================================
# Carpeta de trabajo (pruebas). Los Excel y la salida quedan acá, fuera del repositorio.
CARPETA_TRABAJO = Path(r"C:\Users\drago\BUENOS AIRES GAS S.A\TÉCNICA - Documents\Gerencia Tecnica\Andrés BG\prueba\pr Mechita")
CARPETA_CERTIFICADOS = CARPETA_TRABAJO                # donde están los Certificado_<OBRA>.xlsm
ARCHIVO_SEGUIMIENTO = CARPETA_TRABAJO / "00 - Seguimiento certificacion.xlsx"
CARPETA_SALIDA = CARPETA_TRABAJO / "data"             # los JSON se generan acá; copiarlos al repo es un paso manual

# clave de URL -> archivo de certificación y nombre de la obra en la TABLA MASTER / SALDO
# Opcional: "url" = dirección completa del panel de esa obra, si no es este mismo certificaciones.html
OBRAS = {
    "mechita":   {"archivo": "Certificado_MECHITA.xlsm",   "master": "MECHITA",
                  "planOriginal": "Plan_de_trabajos_e_inversion_v1.xls"},   # plan de la oferta (opcional)
    "anchorena": {"archivo": "Certificado_ANCHORENA.xlsm", "master": "ANCHORENA"},
    "chascomus": {"archivo": "Certificado_CHASCOMUS.xlsm", "master": "CHASCOMUS"},
}

# días por defecto de cada etapa hasta el pago, si no hay historia suficiente
DIAS_DEFAULT = {"certificado": 50, "acta": 46, "envio": 43, "entrada": 37, "tesoreria": 33}

# ============================================================
# Utilidades
# ============================================================
ROMANOS = {"I": 1, "V": 5, "X": 10, "L": 50}
ROMANOS_TXT = ["", "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII", "XIII", "XIV", "XV",
               "XVI", "XVII", "XVIII", "XIX", "XX"]


def romano(s):
    s = (s or "").strip().upper()
    if not s or any(ch not in ROMANOS for ch in s):
        return None
    total = 0
    for i, ch in enumerate(s):
        v = ROMANOS[ch]
        total += -v if i + 1 < len(s) and ROMANOS[s[i + 1]] > v else v
    return total


def norm(v):
    """Texto normalizado: sin tildes, mayúsculas, espacios simples."""
    s = unicodedata.normalize("NFKD", str(v))
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", s).strip().upper()


def num(v):
    if isinstance(v, bool) or v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v) if not math.isnan(v) else None
    return None


def fecha(v):
    if isinstance(v, datetime):
        v = v.date()
    if isinstance(v, date) and v.year >= 2000:
        return v
    return None


def iso(d):
    return d.isoformat() if d else None


def r2(x):
    return None if x is None else round(x, 2)


def r4(x):
    return None if x is None else round(x, 6)


def mes_key(d):
    return f"{d.year:04d}-{d.month:02d}" if d else None


def sumar_meses(d, n):
    m = d.month - 1 + n
    return date(d.year + m // 12, m % 12 + 1, 1)


def meses_entre(a, b):
    """Meses de a (exclusivo) a b (inclusivo), por mes calendario."""
    return (b.year - a.year) * 12 + (b.month - a.month)


CORTES_DESCRIPCION = ("Por ", "Elaboraci", "El ", "La ", "Los ", "Las ", "Incluye", "Se ")


def texto_corto(s, max_len=120, solo_primera=False):
    """Título de una descripción larga. Toma la primera línea y suma las siguientes solo si
    la continúan (empiezan en minúscula o son un fragmento corto como 'TGS,' o 'Tramo A-B')."""
    if not s:
        return ""
    lineas = [l.strip() for l in str(s).split("\n") if l.strip()]
    if not lineas:
        return ""
    txt = lineas[0]
    if not solo_primera:
        for l in lineas[1:]:
            continua = l[:1].islower() or len(l) <= 25
            if not continua or l.startswith(CORTES_DESCRIPCION) or len(txt) > max_len:
                break
            txt = txt + l if txt.endswith("-") else txt + " " + l
    txt = re.sub(r"\s+", " ", txt).strip(" ,.-")
    for corte in (" Por la ", " Por el ", " Por los "):
        if corte in txt:
            txt = txt.split(corte)[0].strip(" ,.-")
    return txt if len(txt) <= max_len else txt[: max_len - 1].rstrip() + "…"


class Hoja:
    """Copia en memoria de una hoja: {(fila, col): valor}."""

    def __init__(self, ws):
        self.c = {}
        for r, row in enumerate(ws.iter_rows(values_only=True), start=1):
            for col, v in enumerate(row, start=1):
                if v is None or v == "":
                    continue
                if isinstance(v, str) and v.strip().startswith("#"):  # #N/A, #REF!, #VALUE!
                    continue
                self.c[(r, col)] = v
        self.max_row = max((r for r, _ in self.c), default=0)

    def v(self, r, c):
        return self.c.get((r, c))

    def buscar(self, texto, col=None, filas=None, exacto=False):
        t = norm(texto)
        out = []
        for (r, c), v in self.c.items():
            if not isinstance(v, str):
                continue
            if col is not None and c != col:
                continue
            if filas is not None and not (filas[0] <= r <= filas[1]):
                continue
            n = norm(v)
            if (n == t) if exacto else n.startswith(t):
                out.append((r, c))
        return sorted(out)

    def primero(self, texto, **kw):
        res = self.buscar(texto, **kw)
        return res[0] if res else None

    def fila(self, r):
        return {c: v for (rr, c), v in self.c.items() if rr == r}


def abrir_hoja(wb, nombre):
    objetivo = norm(nombre)
    for ws in wb.worksheets:
        if norm(ws.title) == objetivo:
            return Hoja(ws)
    return None


def valor_a_la_derecha(h, etiqueta, filas=(1, 15)):
    pos = h.primero(etiqueta, filas=filas)
    if not pos:
        return None
    r, c = pos
    for cc in range(c + 1, c + 8):
        v = h.v(r, cc)
        if v is not None:
            return v
    return None


def columnas_por_titulo(h, fila, titulos):
    """titulos: {clave: texto_inicial}. Devuelve {clave: col} (primera coincidencia)."""
    celdas = sorted(h.fila(fila).items())
    out = {}
    for clave, t in titulos.items():
        tn = norm(t)
        for c, v in celdas:
            if isinstance(v, str) and norm(v).startswith(tn) and c not in out.values():
                out[clave] = c
                break
    return out


# ============================================================
# Lectura del Excel de certificación
# ============================================================
def leer_ficha(pres):
    plazo_txt = valor_a_la_derecha(pres, "Plazo de ejecución")
    m = re.search(r"\d+", str(plazo_txt or ""))
    return {
        "descripcion": texto_corto(valor_a_la_derecha(pres, "Obra:"), 200),
        "comitente": valor_a_la_derecha(pres, "Comitente:"),
        "contratista": valor_a_la_derecha(pres, "Contratista:"),
        "localidad": valor_a_la_derecha(pres, "Localidad:"),
        "montoContrato": num(valor_a_la_derecha(pres, "Monto Contrato:")),
        "plazoDias": int(m.group()) if m else None,
        "fechaContrato": fecha(valor_a_la_derecha(pres, "Fecha de contrato:")),
        "anticipoPct": num(valor_a_la_derecha(pres, "Anticipo financiero:")),
    }


def leer_resumen_general(h, avisos):
    pos = h.primero("Nº DE CERTIFICADO")
    if not pos:
        avisos.append("Resumen General: no se encontró el encabezado 'Nº DE CERTIFICADO'.")
        return [], [], {}
    fila_h = pos[0]
    col = columnas_por_titulo(h, fila_h, {
        "n": "Nº DE CERTIFICADO",
        "redet": "Nº DE REDETERMINACION",
        "fecha": "FECHA DEL CERTIFICADO",
        "mes": "MES DE EJECUCION",
        "base": "MONTO DEL CERTIFICADO A PRECIOS DE CONTRATO",
        "conRedet": "MONTO DEL CERTIFICADO A PRECIOS REDET",
        "dif": "DIFERENCIA",
        "ajuste": "MONTO DE LOS AJUSTES",
        "anticipo": "ANTICIPO FINANCIERO DESCONTADO",
        "fondo": "MONTO DE LA POLIZA",
        "neto": "NETO A PAGAR",
        "acum": "IMPORTE ACUMULADO CERTIFICADOS A LA FECHA",
        "pct": "PORCENTAJE MENSUAL",
        "pctAcum": "PORCENTAJE ACUMULADO",
    })
    fin = h.primero("SUBTOTALES")
    fila_fin = fin[0] if fin else h.max_row

    certs, anticipos = [], []
    for r in range(fila_h + 1, fila_fin):
        a = num(h.v(r, col.get("n", 1)))
        if a is None or a <= 0:
            continue
        f = fecha(h.v(r, col.get("fecha")))
        base = num(h.v(r, col.get("base")))
        if a < 1:  # filas de anticipo (0,2 / 0,1)
            if base:
                anticipos.append({
                    "pct": r4(a), "fecha": iso(f), "monto": r2(base),
                    "neto": r2(num(h.v(r, col.get("neto"))) or base),
                })
            continue
        if not f or not base:
            continue  # filas preparadas para certificados futuros
        redet_txt = str(h.v(r, col.get("redet")) or "")
        mr = re.search(r"REDET\.?\s*([IVXL]+)", norm(redet_txt))
        certs.append({
            "n": int(a),
            "redet": romano(mr.group(1)) if mr else 0,
            "fecha": iso(f),
            "mes": mes_key(fecha(h.v(r, col.get("mes")))),
            "base": r2(base),
            "conRedet": r2(num(h.v(r, col.get("conRedet")))),
            "difRedet": r2(num(h.v(r, col.get("dif"))) or 0),
            "anticipoDesc": r2(num(h.v(r, col.get("anticipo"))) or 0),
            "polizaCaucion": r2(num(h.v(r, col.get("fondo"))) or 0),
            "neto": r2(num(h.v(r, col.get("neto")))),
            "acumConRedet": r2(num(h.v(r, col.get("acum")))),
            "pct": r4(num(h.v(r, col.get("pct")))),
            "pctAcum": r4(num(h.v(r, col.get("pctAcum")))),
        })

    # Bloque de totales (debajo de la tabla)
    tot = {}
    pos_acum = h.primero("ACUMULADO", exacto=True, filas=(fila_fin, fila_fin + 20))
    pos_falt = h.primero("FALTANTE", exacto=True, filas=(fila_fin, fila_fin + 20))
    c_acum = pos_acum[1] if pos_acum else None
    c_falt = pos_falt[1] if pos_falt else None

    def tomar(etiqueta, columna):
        p = h.primero(etiqueta, filas=(fila_fin, fila_fin + 30))
        return num(h.v(p[0], columna)) if p and columna else None

    tot["certificadoPreciosContrato"] = r2(tomar("CERTIFICADO A LA FECHA PRECIOS DE CONTRATO", c_acum))
    tot["certificadoPreciosVigentes"] = r2(tomar("CERTIFICADO A LA FECHA PRECIOS ULTIMA", c_acum))
    tot["montoVigente"] = r2(tomar("MONTO DE CONTRATO MAS LOS IMPORTES REDETERMINADOS", c_falt))
    tot["pctAvance"] = r4(tomar("PORCENTAJE DE OBRA SEGUN PRECIOS DE CONTRATO", c_acum))
    tot["anticipoDescontado"] = r2(tomar("ANTICIPO DESCONTADO", c_falt))
    tot["anticipoAprobado"] = r2(tomar("ANTICIPO APROBADO ACTUALIZADO", c_falt))
    return certs, anticipos, tot


def leer_resumen_ajuste(h):
    if not h:
        return []
    pos = h.primero("Nº DE REDETERMINACION")
    if not pos:
        return []
    fila_h = pos[0]
    col = columnas_por_titulo(h, fila_h, {
        "redet": "Nº DE REDETERMINACION",
        "desde": "Nº DEL CERTIFICADO A PARTIR",
        "fecha": "FECHA DEL CERTIFICADO DE AJUSTE",
        "mesRedet": "MES DE LA REDETERMIN",
        "monto": "MONTO DEL CERTIFICADO",
        "anticipo": "ANTICIPO FINANCIERO DESCONTADO",
        "fondo": "MONTO DE LA POLIZA",
        "neto": "NETO A PAGAR",
        "ajusteAnticipo": "AJUSTE DEL ANTICIPO",
    })
    out = []
    for r in range(fila_h + 1, h.max_row + 1):
        k = num(h.v(r, col.get("redet", 1)))
        monto = num(h.v(r, col.get("monto")))
        if not k or k < 1 or not monto:
            continue
        out.append({
            "redet": int(k),
            "desdeCert": int(num(h.v(r, col.get("desde"))) or 0) or None,
            "fecha": iso(fecha(h.v(r, col.get("fecha")))),
            "monto": r2(monto),
            "anticipoDesc": r2(num(h.v(r, col.get("anticipo"))) or 0),
            "polizaCaucion": r2(num(h.v(r, col.get("fondo"))) or 0),
            "neto": r2(num(h.v(r, col.get("neto")))),
            # el ajuste del anticipo financiero se cobra junto con el certificado de ajuste
            "ajusteAnticipo": r2(num(h.v(r, col.get("ajusteAnticipo"))) or 0),
        })
    for a in out:
        a["netoConAnticipo"] = r2((a["neto"] or 0) + (a["ajusteAnticipo"] or 0))
    return out


def leer_rango_ajuste(h):
    """Hoja 'AJUSTE n': bloque 'CERTIFICADOS A AJUSTAR' con filas 'Nº <desde> al <hasta>'."""
    if not h:
        return []
    p = h.primero("CERTIFICADOS A AJUSTAR")
    if not p:
        return []
    out = []
    for r in range(p[0] + 1, p[0] + 8):
        for c, v in sorted(h.fila(r).items()):
            if isinstance(v, str) and norm(v) == "AL":
                a, b = num(h.v(r, c - 1)), num(h.v(r, c + 1))
                if a and b:
                    out.append([int(a), int(b)])
    return out


def leer_saldo_obra(h, avisos):
    if not h:
        avisos.append("No se encontró la hoja 'Saldo de Obra'.")
        return [], {}
    pos = h.primero("DESCRIPCIÓN", exacto=True)
    if not pos:
        avisos.append("Saldo de Obra: no se encontró el encabezado 'DESCRIPCIÓN'.")
        return [], {}
    fh, c_desc = pos
    c_cod = c_desc - 1
    col = columnas_por_titulo(h, fh, {
        "unidad": "Unidad", "cantidad": "Cantidad", "pu": "Precio Unitario",
        "total": "TOTAL", "peso": "Porcentaje del item", "acum": "Acumulado hasta la fecha",
    })
    sub = {}
    for c, v in h.fila(fh + 1).items():
        if isinstance(v, str):
            sub.setdefault(norm(v), c)
    c_acum = col.get("acum")
    c_saldo_cant = sub.get("SALDO DE OBRA EN CANTIDAD")
    c_no_cant = sub.get("NO SE EJECUTA CANTIDAD")
    c_saldo = sub.get("SALDO DE OBRA")
    c_no = sub.get("NO SE EJECUTA")

    fin = h.primero("TOTAL", col=c_desc, filas=(fh + 2, h.max_row))
    fila_fin = fin[0] if fin else h.max_row

    items, rubro, padre_cod, padre_desc = [], None, None, None
    for r in range(fh + 2, fila_fin):
        cod = h.v(r, c_cod)
        desc = h.v(r, c_desc)
        unidad = h.v(r, col.get("unidad"))
        cant = num(h.v(r, col.get("cantidad")))
        if isinstance(cod, str) and norm(cod).startswith("RUBRO"):
            rubro = texto_corto(desc)
            continue
        if not desc:
            continue
        if not unidad or not cant:
            # encabezado de ítem (p. ej. "1 Provisión e instalación de cañería") o subtítulo
            if isinstance(cod, (int, float)) or (isinstance(cod, str) and cod.strip().isdigit()):
                padre_cod = str(int(float(cod)))
                padre_desc = texto_corto(desc, 90, solo_primera=True)
            continue
        if isinstance(cod, (int, float)) or (isinstance(cod, str) and cod.strip().isdigit()):
            codigo, padre = str(int(float(cod))), None
            padre_cod = codigo
        else:
            codigo = f"{padre_cod or ''}{str(cod or '').strip()}"
            padre = padre_desc
        items.append({
            "codigo": codigo,
            "rubro": rubro,
            "grupo": padre,
            "descripcion": texto_corto(desc, 140),
            "unidad": str(unidad).strip(),
            "cantidad": r4(cant),
            "precioVigente": r2(num(h.v(r, col.get("pu")))),
            "totalVigente": r2(num(h.v(r, col.get("total")))),
            "peso": r4(num(h.v(r, col.get("peso")))),
            "pctAcum": r4(num(h.v(r, c_acum)) if c_acum else None),
            "cantAcum": r4(num(h.v(r, c_acum + 1)) if c_acum else None),
            "saldoCant": r4(num(h.v(r, c_saldo_cant)) if c_saldo_cant else None),
            "noSeEjecutaCant": r4(num(h.v(r, c_no_cant)) if c_no_cant else None),
            "saldoMonto": r2(num(h.v(r, c_saldo)) if c_saldo else None),
            "noSeEjecutaMonto": r2(num(h.v(r, c_no)) if c_no else None),
        })

    pie = {}
    p = h.primero("TOTAL GENERAL", col=c_desc, filas=(fila_fin, fila_fin + 10))
    if p and c_saldo:
        pie["saldoTotalConIva"] = r2(num(h.v(p[0], c_saldo)))
    p = h.primero("A PAGAR", filas=(fila_fin, fila_fin + 10))
    if p and c_saldo:
        pie["saldoAPagar"] = r2(num(h.v(p[0], c_saldo)))
    return items, pie


def leer_metodologia(h, items):
    """Avance de los ítems globales (ERP, etc.) desglosado por disciplina."""
    if not h:
        return []
    desc_por_cod = {i["codigo"]: i["descripcion"] for i in items}
    encabezados = [p for p in h.buscar("EJECUTADO ACUMULADO")]
    out = []
    for idx, (fh, c_acum) in enumerate(encabezados):
        fila_sig = encabezados[idx + 1][0] if idx + 1 < len(encabezados) else fh + 30
        c_tarea = c_acum - 2
        # número de ítem: primera celda numérica de la fila siguiente, columnas B
        item = num(h.v(fh + 1, 2))
        if not item:
            continue
        cod = str(int(item))
        bloque = {"codigo": cod, "descripcion": desc_por_cod.get(cod) or texto_corto(h.v(fh + 1, 3), 90),
                  "disciplinas": [], "total": None}
        for r in range(fh + 1, fila_sig):
            for c in range(1, c_tarea):
                v = h.v(r, c)
                if isinstance(v, str) and norm(v) in ("CIVIL", "MECANICA", "ELECTRICA"):
                    nombre = norm(v)
                    peso = num(h.v(r, c + 1))
                    if not peso:
                        continue
                    etiqueta = {"CIVIL": "% DEL TOTAL CIVIL", "MECANICA": "% DEL TOTAL MECANICO",
                                "ELECTRICA": "% DEL TOTAL ELECTRICO"}[nombre]
                    p = h.primero(etiqueta, col=c_tarea, filas=(fh + 1, fila_sig - 1))
                    avance = num(h.v(p[0], c_acum + 1)) if p else None
                    tareas = []
                    if nombre == "CIVIL":
                        fin_civil = p[0] if p else r
                        for rr in range(r, fin_civil):
                            t = h.v(rr, c_tarea)
                            if isinstance(t, str) and not norm(t).startswith("%"):
                                tareas.append({"tarea": texto_corto(t, 60).capitalize(),
                                               "peso": r4(num(h.v(rr, c_tarea + 1))),
                                               "avance": r4(num(h.v(rr, c_acum)))})
                    bloque["disciplinas"].append({
                        "nombre": {"CIVIL": "Civil", "MECANICA": "Mecánica", "ELECTRICA": "Eléctrica"}[nombre],
                        "peso": r4(peso), "avance": r4(avance), "tareas": tareas})
        p = h.primero("% DEL TOTAL DE LA ERP", col=c_acum, filas=(fh + 1, fila_sig - 1))
        if p:
            bloque["total"] = r4(num(h.v(p[0], c_acum + 2)))
        if bloque["disciplinas"]:
            out.append(bloque)
    return out


def leer_meses_salto(h):
    if not h:
        return {}
    p_mes = h.primero("MES DEL SALTO")
    p_num = h.primero("NUMERO DE REDETERMINACION")
    if not p_mes or not p_num:
        return {}
    out = {}
    for c in range(p_mes[1] + 1, p_mes[1] + 80):
        k = num(h.v(p_num[0], c))
        f = fecha(h.v(p_mes[0], c))
        if k is not None and f:
            out.setdefault(int(k), f)
    return out


class _FilasXLS:
    """Adaptador mínimo para leer una hoja .xls (xlrd) con la misma interfaz que openpyxl."""

    def __init__(self, hoja, datemode):
        self.h, self.datemode = hoja, datemode

    def iter_rows(self, values_only=True):
        import xlrd
        for r in range(self.h.nrows):
            fila = []
            for c in range(self.h.ncols):
                t, v = self.h.cell_type(r, c), self.h.cell_value(r, c)
                if t in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                    v = None
                elif t == xlrd.XL_CELL_DATE:
                    v = xlrd.xldate.xldate_as_datetime(v, self.datemode)
                elif t == xlrd.XL_CELL_ERROR:
                    v = "#ERROR"
                fila.append(v)
            yield tuple(fila)


def leer_plan_externo(ruta, hoja="PT - INVERS."):
    """Plan de trabajo e inversión en un archivo aparte (.xls, .xlsx o .xlsm).
    Devuelve la serie [{mes, pctAcum, pct}] tomada de las filas 'AVANCE ACUMULADO' y 'AVANCE PORCENTUAL'."""
    ruta = Path(ruta)
    if ruta.suffix.lower() == ".xls":
        try:
            import xlrd
        except ImportError:
            raise RuntimeError("Para leer un .xls hace falta xlrd (pip install xlrd) o guardar el archivo como .xlsx.")
        libro = xlrd.open_workbook(str(ruta))
        h = None
        for nombre in libro.sheet_names():
            if norm(nombre) == norm(hoja):
                h = Hoja(_FilasXLS(libro.sheet_by_name(nombre), libro.datemode))
    else:
        wbp = load_workbook(ruta, read_only=True, data_only=True)
        h = abrir_hoja(wbp, hoja)
        wbp.close()
    if not h:
        raise RuntimeError(f"No se encontró la hoja '{hoja}' en {ruta.name}.")
    fila_fechas = next((r for r in range(1, 30) if sum(1 for v in h.fila(r).values() if fecha(v)) >= 3), None)
    p_acum, p_mens = h.primero("AVANCE ACUMULADO"), h.primero("AVANCE PORCENTUAL")
    if not (fila_fechas and p_acum):
        raise RuntimeError(f"{ruta.name}: no se encontraron las fechas o la fila 'AVANCE ACUMULADO'.")
    out = []
    for c, v in sorted(h.fila(fila_fechas).items()):
        f = fecha(v)
        if not f:
            continue
        # las celdas del mes pueden estar combinadas: el valor queda en la columna de la derecha
        a = num(h.v(p_acum[0], c))
        a = num(h.v(p_acum[0], c + 1)) if a is None else a
        m = None
        if p_mens:
            m = num(h.v(p_mens[0], c))
            m = num(h.v(p_mens[0], c + 1)) if m is None else m
        if a is not None:
            out.append({"mes": mes_key(f), "pctAcum": r4(a), "pct": r4(m)})
    return out


def leer_nombre_plan(h):
    """Rótulo del plan en la hoja 'P.T.Propuesta %' (p. ej. 'PLAN DE TRABAJO - R IV')."""
    if not h:
        return None
    for (r, c), v in sorted(h.c.items()):
        if r <= 20 and isinstance(v, str) and norm(v).startswith("PLAN DE TRABAJO") and norm(v) != "PLAN DE TRABAJO PROPUESTO":
            return re.sub(r"\s+", " ", v).strip()
    return None


def leer_plan_original(h):
    if not h:
        return []
    p_acum = h.primero("% AVANCE ACUMULADO PROPUESTO")
    p_mens = h.primero("% AVANCE MENSUAL PROPUESTO")
    if not p_acum or not p_mens:
        return []
    # fila de fechas: la más cercana por encima con fechas
    fila_fechas = None
    for r in range(p_mens[0] - 1, p_mens[0] - 6, -1):
        if sum(1 for c, v in h.fila(r).items() if fecha(v)) >= 3:
            fila_fechas = r
            break
    if not fila_fechas:
        return []
    out = []
    for c, v in sorted(h.fila(fila_fechas).items()):
        f = fecha(v)
        a = num(h.v(p_acum[0], c))
        if f and a is not None:
            out.append({"mes": mes_key(f), "pctAcum": r4(a), "pct": r4(num(h.v(p_mens[0], c)))})
    # cortar repeticiones finales al 100 %
    while len(out) > 1 and out[-1]["pctAcum"] == out[-2]["pctAcum"]:
        out.pop()
    return out


# ============================================================
# Lectura del seguimiento (TABLA MASTER y SALDO)
# ============================================================
def parsear_tema(tema):
    t = norm(tema)
    if "ANTICIPO" in t:
        m = re.search(r"(\d+)\s*%", t)
        orden = re.match(r"\s*(\d+)", t)
        return {"tipo": "anticipo", "pct": int(m.group(1)) / 100 if m else None,
                "orden": int(orden.group(1)) if orden else None}
    m = re.search(r"\bAJUSTE\s+([IVXL]+)\b", t)
    if m:
        return {"tipo": "ajuste", "redet": romano(m.group(1))}
    m = re.search(r"CERTIFICADO\s*N\S*\s*(\d+)", t)
    if m:
        mr = re.search(r"REDET\.?\s*([IVXL]+)\b", t)
        return {"tipo": "certificado", "n": int(m.group(1)), "redet": romano(mr.group(1)) if mr else 0}
    return {"tipo": "otro"}


def leer_master(wb_seg, incluir_sin_obra=False):
    ws = None
    for w in wb_seg.worksheets:
        if norm(w.title) == "TABLA MASTER":
            ws = w
    if ws is None:
        return []
    filas = list(ws.iter_rows(values_only=True))
    if not filas:
        return []
    enc = {norm(v): i for i, v in enumerate(filas[0]) if v}

    def idx(nombre):
        return enc.get(norm(nombre))

    campos = {
        "obra": idx("Obra de Referencia"), "empresa": idx("Empresa"), "tema": idx("Tema"),
        "entrada": idx("Fecha de Entrada"), "tesoreria": idx("Entregado a Tesoreria"),
        "pago": idx("Fecha de Pago"), "monto": idx("Monto a pagar"),
        "montoActualizado": idx("Monto Actualizado"), "factura": idx("Adicional 1"),
        "poliza": idx("Adicional 2"), "acta": idx("Acta de medicion"),
        "envio": idx("Envio a contratista"), "obs": idx("observaciones"),
    }
    out = []
    for i, f in enumerate(filas[1:], start=2):
        def g(k):
            j = campos.get(k)
            return f[j] if j is not None and j < len(f) else None
        if not g("tema") or (not g("obra") and not incluir_sin_obra):
            continue
        reg = {
            "fila": i, "obra": norm(g("obra")) if g("obra") else "(SIN OBRA)",
            "obraTexto": str(g("obra") or "").strip(), "empresa": str(g("empresa") or "").strip(),
            "tema": re.sub(r"\s+", " ", str(g("tema"))).strip(),
            "acta": fecha(g("acta")), "envio": fecha(g("envio")), "entrada": fecha(g("entrada")),
            "tesoreria": fecha(g("tesoreria")), "pago": fecha(g("pago")),
            "monto": num(g("monto")), "montoActualizado": num(g("montoActualizado")),
            "factura": texto_corto(g("factura"), 120) if isinstance(g("factura"), str) else None,
            "poliza": texto_corto(g("poliza"), 160) if isinstance(g("poliza"), str) else None,
            "obs": re.sub(r"\s+", " ", g("obs")).strip() if isinstance(g("obs"), str) else None,
        }
        reg.update(parsear_tema(reg["tema"]))
        # textos sueltos en columnas de fecha (distintos de "-"), para el control de datos
        reg["textosFecha"] = {k: str(g(k)).strip() for k in ("acta", "envio", "entrada", "tesoreria", "pago")
                              if isinstance(g(k), str) and str(g(k)).strip(" -_") != ""}
        out.append(reg)
    return out


def leer_saldo_seguimiento(wb_seg, clave):
    ws = None
    for w in wb_seg.worksheets:
        if norm(w.title) == "SALDO":
            ws = w
    if ws is None:
        return {}
    h = Hoja(ws)
    enc = h.primero("Localidad de Referencia")
    if not enc:
        return {}
    col = columnas_por_titulo(h, enc[0], {
        "fin": "MES DE FINALIZACION", "monto": "Monto de obra actualizado", "saldo": "SALDO DE OBRA A PAGAR",
    })
    p = h.primero(clave, col=enc[1], exacto=True)
    if not p:
        return {}
    r = p[0]
    return {
        "finPlanVigente": iso(fecha(h.v(r, col.get("fin")))),
        "montoActualizado": r2(num(h.v(r, col.get("monto")))),
        "saldoAPagar": r2(num(h.v(r, col.get("saldo")))),
    }


# ============================================================
# Armado del JSON
# ============================================================
ETAPAS = [("acta", "Acta de medición"), ("envio", "Enviado al contratista"),
          ("entrada", "Factura recibida"), ("tesoreria", "En Tesorería"), ("pago", "Pagado")]


def medianas_a_pago(registros):
    out = {}
    for etapa, _ in ETAPAS[:-1]:
        dias = [(r["pago"] - r[etapa]).days for r in registros
                if r.get("pago") and r.get(etapa) and 0 <= (r["pago"] - r[etapa]).days <= 200]
        if len(dias) >= 3:
            out[etapa] = statistics.median(dias)
    dias = [(r["pago"] - r["fechaCert"]).days for r in registros
            if r.get("pago") and r.get("fechaCert") and 0 <= (r["pago"] - r["fechaCert"]).days <= 200]
    if len(dias) >= 3:
        out["certificado"] = statistics.median(dias)
    return out


def procesar_obra(clave, conf, master, wb_seg):
    avisos = []
    ruta = CARPETA_CERTIFICADOS / conf["archivo"]
    wb = load_workbook(ruta, read_only=True, data_only=True, keep_vba=False)

    pres = abrir_hoja(wb, "Presupuesto")
    ficha = leer_ficha(pres) if pres else {}
    certs, anticipos, tot = leer_resumen_general(abrir_hoja(wb, "Resumen General"), avisos)
    ajustes = leer_resumen_ajuste(abrir_hoja(wb, "Resumen Ajuste"))
    for a in ajustes:
        a["certRangos"] = leer_rango_ajuste(abrir_hoja(wb, f"AJUSTE {a['redet']}"))
    items, pie_saldo = leer_saldo_obra(abrir_hoja(wb, "Saldo de Obra"), avisos)
    ultimo = certs[-1] if certs else None
    metod = leer_metodologia(abrir_hoja(wb, f"Metodologia de medicion Nº {ultimo['n']}"), items) if ultimo else []
    meses_salto = leer_meses_salto(abrir_hoja(wb, "REDETERMINACIONES"))
    plan = leer_plan_original(abrir_hoja(wb, "C I Comparativa"))
    plan_nombre = leer_nombre_plan(abrir_hoja(wb, "P.T.Propuesta %"))
    plan_fin = next((x["mes"] for x in plan if (x["pctAcum"] or 0) >= 0.9999), None)
    plan_orig, plan_orig_fin = None, None
    if conf.get("planOriginal"):
        ruta_plan = CARPETA_CERTIFICADOS / conf["planOriginal"]
        if ruta_plan.exists():
            try:
                plan_orig = leer_plan_externo(ruta_plan)
                plan_orig_fin = next((x["mes"] for x in plan_orig if (x["pctAcum"] or 0) >= 0.9999), None)
                if not plan_orig_fin:
                    avisos.append(f"El plan original ({ruta_plan.name}) no llega al 100 %.")
            except Exception as e:  # el panel sigue funcionando sin el plan original
                avisos.append(f"No se pudo leer el plan original: {e}")
        else:
            avisos.append(f"No se encontró el plan original: {ruta_plan}")
    wb.close()

    saldo_seg = leer_saldo_seguimiento(wb_seg, conf["master"]) if wb_seg else {}

    # --- TABLA MASTER de esta obra, cruzada con el certificado
    regs = [r for r in master if r["obra"] == norm(conf["master"])]
    cert_por_n = {c["n"]: c for c in certs}
    aj_por_k = {a["redet"]: a for a in ajustes}
    ant_ordenados = sorted(anticipos, key=lambda a: a["fecha"] or "")
    for r in regs:
        r["fechaCert"] = None
        if r["tipo"] == "certificado":
            c = cert_por_n.get(r["n"])
            if c:
                r["fechaCert"] = date.fromisoformat(c["fecha"])
                if r["monto"] is not None and c["neto"] is not None and abs(r["monto"] - c["neto"]) > 1:
                    avisos.append(f"Cert. {r['n']}: el monto de la TABLA MASTER ({r['monto']:,.2f}) "
                                  f"no coincide con el neto del Resumen General ({c['neto']:,.2f}).")
            else:
                avisos.append(f"TABLA MASTER fila {r['fila']}: '{r['tema']}' no figura en el Resumen General.")
        elif r["tipo"] == "ajuste":
            a = aj_por_k.get(r["redet"])
            if a and a["fecha"]:
                r["fechaCert"] = date.fromisoformat(a["fecha"])
            if a and r["monto"] is not None and abs(r["monto"] - a["netoConAnticipo"]) > 1:
                avisos.append(f"Ajuste Redet. {r['redet']}: el monto de la TABLA MASTER ({r['monto']:,.2f}) no coincide "
                              f"con el neto + ajuste de anticipo del Resumen Ajuste ({a['netoConAnticipo']:,.2f}).")
            elif not a:
                avisos.append(f"TABLA MASTER fila {r['fila']}: '{r['tema']}' no figura en el Resumen Ajuste.")
        elif r["tipo"] == "anticipo":
            cand = [a for a in ant_ordenados if r.get("pct") and abs((a["pct"] or 0) - r["pct"]) < 1e-6]
            if cand and cand[0]["fecha"]:
                r["fechaCert"] = date.fromisoformat(cand[0]["fecha"])
        # orden lógico de fechas
        previas = [(k, r[k]) for k, _ in ETAPAS if r.get(k)]
        for (k1, f1), (k2, f2) in zip(previas, previas[1:]):
            if f2 < f1:
                avisos.append(f"TABLA MASTER fila {r['fila']} ('{r['tema']}'): "
                              f"{dict(ETAPAS)[k2].lower()} ({f2:%d/%m/%Y}) es anterior a "
                              f"{dict(ETAPAS)[k1].lower()} ({f1:%d/%m/%Y}).")
    certs_master_n = {r["n"] for r in regs if r["tipo"] == "certificado"}
    for c in certs:
        if c["n"] not in certs_master_n:
            avisos.append(f"Cert. {c['n']} está en el Resumen General pero no en la TABLA MASTER.")

    contratista = [r for r in regs if r["tipo"] != "otro"]
    otros = [r for r in regs if r["tipo"] == "otro"]

    med = medianas_a_pago(contratista)
    med_global = medianas_a_pago([r for r in master if r["tipo"] != "otro"])
    hoy = date.today()

    def dias_a_pago(etapa):
        return med.get(etapa) or med_global.get(etapa) or DIAS_DEFAULT[etapa]

    tramites = []
    for r in sorted(contratista, key=lambda x: (x.get("fechaCert") or x.get("entrada") or x.get("acta") or date.max)):
        etapa_actual, base = "certificado", r.get("fechaCert")
        for k, _ in ETAPAS:
            if r.get(k):
                etapa_actual, base = k, r[k]
        estado = dict(ETAPAS).get(etapa_actual, "Certificado emitido")
        pago_est = None
        if etapa_actual != "pago" and base:
            pago_est = base + timedelta(days=round(dias_a_pago(etapa_actual)))
        if r["tipo"] == "certificado":
            nombre = f"Certificado Nº {r['n']}"
        elif r["tipo"] == "ajuste":
            nombre = f"Ajuste por Redet. {ROMANOS_TXT[r['redet']] if r['redet'] and r['redet'] < len(ROMANOS_TXT) else r['redet']}"
        else:
            nombre = f"Anticipo {round((r.get('pct') or 0) * 100)} %"
        tramites.append({
            "nombre": nombre, "tipo": r["tipo"], "n": r.get("n"), "redet": r.get("redet"),
            "tema": r["tema"], "monto": r2(r["monto"]),
            "fechaCert": iso(r.get("fechaCert")),
            **{k: iso(r.get(k)) for k, _ in ETAPAS},
            "estado": estado, "pagoEstimado": iso(pago_est),
            "diasCertAPago": (r["pago"] - r["fechaCert"]).days if r.get("pago") and r.get("fechaCert") else None,
            "factura": r["factura"], "poliza": r["poliza"], "obs": r["obs"],
        })

    # --- Escalera de redeterminaciones
    redet_max = max([c["redet"] for c in certs] + [a["redet"] for a in ajustes] + [0])
    monto_por_redet = {}
    for r in contratista:
        k = r.get("redet") if r["tipo"] in ("certificado", "ajuste") else None
        if k is not None and r["montoActualizado"]:
            monto_por_redet.setdefault(k, r["montoActualizado"])
    escalera = []
    for k in range(0, redet_max + 1):
        primer_cert = min([c["n"] for c in certs if c["redet"] == k], default=None)
        monto = ficha.get("montoContrato") if k == 0 else monto_por_redet.get(k)
        escalera.append({
            "redet": k, "mesSalto": iso(meses_salto.get(k)) if k else iso(ficha.get("fechaContrato")),
            "monto": r2(monto), "primerCert": primer_cert,
            "ajuste": aj_por_k.get(k, {}).get("monto"),
        })
    if tot.get("montoVigente") and escalera and escalera[-1]["monto"] and abs(escalera[-1]["monto"] - tot["montoVigente"]) > 1:
        avisos.append("El monto actualizado de la última redeterminación en la TABLA MASTER no coincide "
                      "con el 'monto de contrato más importes redeterminados' del Resumen General.")

    # --- Indicadores y proyección
    pagado = sum(r["monto"] or 0 for r in contratista if r.get("pago") and r["pago"] <= hoy)
    pendientes = [t for t in tramites if t["estado"] != "Pagado"]
    por_cobrar = sum(t["monto"] or 0 for t in pendientes)
    neto_certificado = (sum(c["neto"] or 0 for c in certs) + sum(a["netoConAnticipo"] or 0 for a in ajustes)
                        + sum(a["neto"] or 0 for a in anticipos))

    proy = {}
    if ultimo:
        pct_acum = ultimo["pctAcum"] or 0
        saldo_pct = max(0.0, 1 - pct_acum)
        recientes = [c["pct"] for c in certs[-3:] if c["pct"] is not None]
        ritmo = sum(recientes) / len(recientes) if recientes else None
        ult_mes = date.fromisoformat(ultimo["mes"] + "-01")
        fin_contr = (ficha["fechaContrato"] + timedelta(days=ficha["plazoDias"])) \
            if ficha.get("fechaContrato") and ficha.get("plazoDias") else None
        fin_vig = date.fromisoformat(saldo_seg["finPlanVigente"]) if saldo_seg.get("finPlanVigente") else None
        meses_vig = meses_entre(ult_mes, fin_vig) if fin_vig else None
        proy = {
            "saldoPct": r4(saldo_pct),
            "ritmoReciente": r4(ritmo),
            "ritmoRecienteCerts": [c["n"] for c in certs[-3:]],
            "ultimoMes": ultimo["mes"],
            "finContractual": iso(fin_contr),
            "finPlanVigente": iso(fin_vig),
            "mesesHastaFinVigente": meses_vig,
            "ritmoNecesario": r4(saldo_pct / meses_vig) if meses_vig and meses_vig > 0 else None,
            "finEstimadoRitmo": iso(sumar_meses(ult_mes, math.ceil(saldo_pct / ritmo)))
            if ritmo and ritmo > 0 and saldo_pct > 0 else None,
            "saldoObraConIva": pie_saldo.get("saldoTotalConIva"),
            "saldoObraAPagar": pie_saldo.get("saldoAPagar"),
            "diasMedianosAPago": {k: round(v) for k, v in med.items()},
        }

    data = {
        "generatedAt": datetime.now().strftime("%d/%m/%Y"),
        "generatedTime": datetime.now().strftime("%H:%M"),
        "obra": {
            "clave": clave,
            "nombre": (ficha.get("localidad") or conf["master"]).title(),
            **{k: (iso(v) if isinstance(v, date) else v) for k, v in ficha.items()},
            "montoVigente": tot.get("montoVigente") or (escalera[-1]["monto"] if escalera else None),
            "redetVigente": redet_max,
        },
        "kpis": {
            "pctAvance": (ultimo or {}).get("pctAcum"),
            "ultimoCert": (ultimo or {}).get("n"),
            "ultimoCertFecha": (ultimo or {}).get("fecha"),
            "certificadoPreciosVigentes": tot.get("certificadoPreciosVigentes"),
            "certificadoPreciosContrato": tot.get("certificadoPreciosContrato"),
            "netoCertificado": r2(neto_certificado),
            "pagado": r2(pagado),
            "porCobrar": r2(por_cobrar),
            "anticipoDescontado": tot.get("anticipoDescontado"),
            "anticipoAprobado": tot.get("anticipoAprobado"),
        },
        "certificados": certs,
        "anticipos": anticipos,
        "ajustes": ajustes,
        "redeterminaciones": escalera,
        "curva": {"real": [{"n": c["n"], "mes": c["mes"], "pct": c["pct"], "pctAcum": c["pctAcum"]} for c in certs],
                  "plan": plan, "planNombre": plan_nombre, "planFin": plan_fin,
                  "planOriginal": plan_orig, "planOriginalFin": plan_orig_fin,
                  "planOriginalArchivo": conf.get("planOriginal") if plan_orig else None},
        "items": items,
        "globales": metod,
        "tramites": tramites,
        "otrosPagos": [{"tema": r["tema"], "empresa": r["empresa"], "monto": r2(r["monto"]),
                        "entrada": iso(r["entrada"]), "pago": iso(r["pago"])} for r in otros],
        "proyeccion": proy,
        "seguimientoSaldo": saldo_seg,
        "avisos": avisos,
    }
    return data


NOMBRES_OBRA = {"CHASCOMUS": "Chascomús"}   # correcciones de nombre para mostrar


def nombre_obra(txt):
    """'MONES CAZON - SALAZAR' -> 'Mones Cazon - Salazar'; respeta siglas cortas (PLA, ERP, PY)."""
    t = str(txt or "").strip()
    if not t:
        return "Sin obra asignada"
    if t.upper() in NOMBRES_OBRA:
        return NOMBRES_OBRA[t.upper()]
    if not t.isupper():
        return t
    chicas = {"DEL", "DE", "LA", "LAS", "LOS", "Y", "EL"}

    def palabra(m):
        w = m.group(0)
        if w in chicas:
            return w.lower() if m.start() > 0 else w.capitalize()
        if w.endswith("."):
            return w.capitalize()
        return w if len(w) <= 3 else w.capitalize()

    return re.sub(r"[A-ZÁÉÍÓÚÜÑ]+\.?", palabra, t)


def etiqueta_comprobante(r):
    if r["tipo"] == "certificado" and r.get("n") is not None:
        return f"Cert. {r['n']}"
    if r["tipo"] == "ajuste" and r.get("redet"):
        return f"Ajuste {ROMANOS_TXT[r['redet']] if r['redet'] < len(ROMANOS_TXT) else r['redet']}"
    if r["tipo"] == "anticipo":
        return f"Anticipo {round(r['pct'] * 100)} %" if r.get("pct") else "Anticipo"
    return texto_corto(r["tema"], 48)


def generar_indice(master_all, wb_seg):
    """Índice general: todos los comprobantes de la TABLA MASTER, con datos crudos.
    El estado, las demoras y las fechas de pago estimadas se calculan en el navegador con la fecha del día."""
    hoy = date.today()
    ordenadas = ["acta", "envio", "entrada", "tesoreria", "pago"]
    # medianas desde cada etapa hasta el pago, con todas las obras (sin TGS, que cobra por acuerdo)
    med = medianas_a_pago([r for r in master_all if r["tipo"] != "otro"])
    medianas = {k: round(med[k]) if k in med else DIAS_DEFAULT[k] for k in ("acta", "envio", "entrada", "tesoreria")}

    paneles, urls = {}, {}
    for clave, conf in OBRAS.items():
        if conf.get("url") or (CARPETA_CERTIFICADOS / conf["archivo"]).exists():
            paneles[norm(conf["master"])] = clave
            urls[norm(conf["master"])] = conf.get("url") or None

    comprobantes, avisos, obras = [], [], {}
    for r in master_all:
        etiqueta = etiqueta_comprobante(r)
        item = {
            "id": r["fila"], "obraKey": r["obra"], "obra": nombre_obra(r["obraTexto"]), "empresa": r["empresa"],
            "tema": r["tema"], "label": etiqueta, "tipo": r["tipo"], "n": r.get("n"), "redet": r.get("redet"),
            "pct": r.get("pct"), "monto": r2(r["monto"]),
            **{k: iso(r.get(k)) for k in ordenadas},
            "factura": r.get("factura"), "poliza": r.get("poliza"), "obs": r.get("obs"),
        }
        comprobantes.append(item)
        o = obras.setdefault(r["obra"], {"key": r["obra"], "nombre": item["obra"], "empresas": [], "panel": paneles.get(r["obra"]), "url": urls.get(r["obra"])})
        if r["empresa"] and r["empresa"] not in o["empresas"]:
            o["empresas"].append(r["empresa"])

        # control de datos
        if r["obra"] == "(SIN OBRA)":
            avisos.append({"tipo": "sin_obra", "fila": r["fila"], "texto": f"«{r['tema']}» no tiene obra asignada."})
        fechas = [(k, r[k]) for k in ordenadas if r.get(k)]
        for (k1, f1), (k2, f2) in zip(fechas, fechas[1:]):
            if f2 < f1:
                avisos.append({"tipo": "desorden", "fila": r["fila"], "leve": (f1 - f2).days <= 3,
                               "texto": f"{etiqueta} ({item['obra']}): {k2} ({f2:%d/%m/%Y}) es anterior a {k1} ({f1:%d/%m/%Y})."})
                break
        for k, f in fechas:
            if f < date(2020, 1, 1) or f > hoy + timedelta(days=90):
                avisos.append({"tipo": "fecha_rara", "fila": r["fila"], "texto": f"{etiqueta} ({item['obra']}): la fecha de {k} ({f:%d/%m/%Y}) parece mal cargada."})
        for k, t in (r.get("textosFecha") or {}).items():
            avisos.append({"tipo": "texto_en_fecha", "fila": r["fila"], "texto": f"{etiqueta} ({item['obra']}): en la columna de {k} dice «{t}» en lugar de una fecha."})

    return {
        "generatedAt": datetime.now().strftime("%d/%m/%Y"),
        "generatedTime": datetime.now().strftime("%H:%M"),
        "reglas": {"antiguoDias": 120, "demoraFactor": 1.5},
        "medianas": medianas,
        "obras": sorted(obras.values(), key=lambda o: o["nombre"].lower()),
        "empresas": sorted({c["empresa"] for c in comprobantes if c["empresa"]}),
        "comprobantes": comprobantes,
        "avisos": avisos,
    }


def main():
    elegidas = [a.lower() for a in sys.argv[1:]] or list(OBRAS)
    CARPETA_SALIDA.mkdir(parents=True, exist_ok=True)

    wb_seg, master = None, []
    if ARCHIVO_SEGUIMIENTO.exists():
        wb_seg = load_workbook(ARCHIVO_SEGUIMIENTO, read_only=True, data_only=True)
        master = leer_master(wb_seg)
    else:
        print(f"[aviso] No se encontró {ARCHIVO_SEGUIMIENTO}: el panel saldrá sin fechas de pago.")

    if wb_seg:
        indice = generar_indice(leer_master(wb_seg, incluir_sin_obra=True), wb_seg)
        salida_idx = CARPETA_SALIDA / "certificaciones_indice.json"
        salida_idx.write_text(json.dumps(indice, ensure_ascii=False, indent=1), encoding="utf-8")
        print(f"[ok] índice general: {len(indice['comprobantes'])} comprobantes, {len(indice['obras'])} obras, "
              f"{len(indice['avisos'])} avisos de datos -> {salida_idx}")

    for clave in elegidas:
        conf = OBRAS.get(clave)
        if not conf:
            print(f"[saltada] '{clave}' no está en OBRAS.")
            continue
        if not (CARPETA_CERTIFICADOS / conf["archivo"]).exists():
            print(f"[saltada] {clave}: no existe {CARPETA_CERTIFICADOS / conf['archivo']}")
            continue
        data = procesar_obra(clave, conf, master, wb_seg)
        salida = CARPETA_SALIDA / f"certificaciones_{clave}.json"
        salida.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
        print(f"[ok] {clave}: {len(data['certificados'])} certificados, "
              f"{len(data['items'])} ítems, {len(data['avisos'])} avisos -> {salida}")
        for a in data["avisos"]:
            print("     ·", a)

    if wb_seg:
        wb_seg.close()


if __name__ == "__main__":
    main()
