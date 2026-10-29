"""Planilla de trazabilidad: de dónde sale cada gráfico y tabla del PPT.

Mientras se arma el informe, cada diapositiva registra aquí los datos con que se
construyó. Al final se escribe un Excel con:

* **LEEME**: parámetros de la corrida, definiciones y cómo leer cada hoja.
* **Índice**: una fila por diapositiva, con vínculo a su hoja de detalle.
* **Una hoja por gráfico/tabla** (``D05 …``), con tres tablas:
    A. Detalle por entidad, en la unidad original (CLP o TRP), con la columna
       "Se muestra como" (la entidad o OTRAS).
    B. Datos del gráfico: fórmulas ``SUMAR.SI`` sobre la tabla A, de modo que
       la agrupación en OTRAS y los totales se pueden auditar en Excel.
    C. Participación (SOS / SOV / mix): fórmulas sobre la tabla B.
* **Clasificación**: cada marca × área × empresa con su compañía, categoría,
  marca del informe, BLYS y la regla que la ubicó ahí.
* **Diagnóstico**: marcas que no calzan con la hoja Agrupaciones.
* **Base** (opcional): la base de Megatime con las columnas de clasificación,
  como tabla de Excel para filtrar o hacer tablas dinámicas.
"""
from __future__ import annotations

import io
import re
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any

import pandas as pd
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name, xl_rowcol_to_cell

from . import config

VERDE = "#27B088"
GRIS = "#D9D9D9"
GRIS_CLARO = "#F2F2F2"


# --------------------------------------------------------------------------- #
# Estructuras que registran las diapositivas
# --------------------------------------------------------------------------- #
@dataclass
class Bloque:
    """Un conjunto de datos de un gráfico/tabla (p. ej. un año o un periodo)."""

    titulo: str
    detalle: pd.DataFrame           # filas = todas las entidades, columnas = periodos/meses/medios
    grupos: list[str]               # series del gráfico en orden (puede incluir OTRAS)
    asignacion: dict[str, str]      # entidad -> serie del gráfico
    unidad: str                     # "CLP" o "TRP"
    participacion: str | None       # "columna", "fila", "total" o None
    etiqueta_total: str = "TOTAL"   # nombre de la fila total en la tabla B
    etiquetas_columnas: list[str] | None = None
    nota: str = ""


@dataclass
class Variacion:
    """Insumos de la fila "Var.% año anterior" de las tablas mensuales."""

    componentes: pd.DataFrame       # index mes; actual, base, tipo_base, aplica
    total_actual: float
    total_base: float
    anio_actual: int
    anio_anterior: int
    unidad: str


@dataclass
class Entrada:
    n: int
    seccion: str
    tipo: str
    titulo: str
    subtitulo: str = ""
    universo: str = ""
    metrica: str = ""
    unidad_ppt: str = ""
    calculo: str = ""
    reproducir: str = ""
    bloques: list[Bloque] = field(default_factory=list)
    variacion: Variacion | None = None
    notas: list[str] = field(default_factory=list)
    hoja: str | None = None


@dataclass
class Trazabilidad:
    entradas: list[Entrada] = field(default_factory=list)
    omitidas: list[str] = field(default_factory=list)

    def registrar(self, entrada: Entrada) -> None:
        self.entradas.append(entrada)


# --------------------------------------------------------------------------- #
# Escritura
# --------------------------------------------------------------------------- #
ABREV_TIPO = {
    "Inversión y SOS": "Inv", "Inversión mensual": "InvMes", "SOS mensual": "SOSmes",
    "Mix de medios": "Mix", "TRP y SOV": "TRP", "TRP mensual": "TRPmes",
    "SOS vs SOV": "SOSvsSOV",
}
ABREV_SECCION = {
    "COMPAÑÍA": "Compañía", "CATEGORÍA": "Categoría", "MASIVOS": "Masivos",
    "BOTELLAS FINAS": "BFinas", "ESPUMANTES": "Espum", "COOLERS Y SANGRÍA": "Coolers",
    "BLYS VINOS": "BLYSvinos", "BLYS TOTAL": "BLYStotal",
}


class _Formatos:
    def __init__(self, wb):
        f = wb.add_format
        self.titulo = f({"bold": True, "font_size": 14, "font_color": VERDE})
        self.subtitulo = f({"bold": True, "font_size": 11, "bottom": 1})
        self.etiqueta = f({"bold": True, "font_color": "#595959", "valign": "top"})
        self.texto = f({"text_wrap": False})
        self.texto_wrap = f({"text_wrap": True, "valign": "top"})
        self.nota = f({"italic": True, "font_color": "#7F7F7F"})
        self.enc = f({"bold": True, "bg_color": GRIS, "border": 1, "align": "center",
                      "valign": "vcenter", "text_wrap": True})
        self.enc_izq = f({"bold": True, "bg_color": GRIS, "border": 1, "valign": "vcenter"})
        self.celda = f({"border": 1})
        self.grupo = f({"border": 1, "font_color": "#595959"})
        # Los ceros se muestran vacíos (meses sin actividad), como en el PPT.
        self.clp = f({"num_format": "#,##0;-#,##0;", "border": 1})
        self.trp = f({"num_format": "#,##0.0;-#,##0.0;", "border": 1})
        self.pct = f({"num_format": "0.0%", "border": 1})
        self.clp_tot = f({"num_format": "#,##0;-#,##0;", "border": 1, "bold": True,
                          "bg_color": GRIS_CLARO})
        self.trp_tot = f({"num_format": "#,##0.0;-#,##0.0;", "border": 1, "bold": True,
                          "bg_color": GRIS_CLARO})
        self.pct_tot = f({"num_format": "0.0%", "border": 1, "bold": True, "bg_color": GRIS_CLARO})
        self.tot_lbl = f({"border": 1, "bold": True, "bg_color": GRIS_CLARO})
        self.vinculo = f({"font_color": "#0563C1", "underline": 1})

    def num(self, unidad: str, total: bool = False):
        if unidad == "TRP":
            return self.trp_tot if total else self.trp
        return self.clp_tot if total else self.clp


def _nombre_hoja(e: Entrada, usados: set[str]) -> str:
    base = f"D{e.n:02d} {ABREV_SECCION.get(e.seccion, e.seccion[:10])} {ABREV_TIPO.get(e.tipo, e.tipo)}"
    base = re.sub(r"[\[\]:*?/\\]", "", base)[:31]
    nombre, k = base, 2
    while nombre in usados:
        nombre = f"{base[:28]}_{k}"
        k += 1
    usados.add(nombre)
    return nombre


def _valor(v) -> float:
    try:
        f = float(v)
    except (TypeError, ValueError):
        return 0.0
    return 0.0 if pd.isna(f) else f


def _escribir_bloque(ws, f: _Formatos, r: int, b: Bloque) -> int:
    """Escribe tablas A, B y C de un bloque a partir de la fila ``r``. Devuelve la
    siguiente fila libre."""
    det = b.detalle.copy()
    det = det.loc[:, [c for c in det.columns]].apply(pd.to_numeric, errors="coerce").fillna(0)
    det = det[det.abs().sum(axis=1) > 0]
    cols = list(det.columns)
    etiquetas = b.etiquetas_columnas or [str(c) for c in cols]
    k = len(cols)
    ws.write(r, 0, b.titulo, f.subtitulo)
    for c in range(1, k + 3):
        ws.write_blank(r, c, None, f.subtitulo)
    r += 1
    if b.nota:
        ws.write(r, 0, b.nota, f.nota)
        r += 1
    if det.empty:
        ws.write(r, 0, "Sin datos en este bloque.", f.nota)
        return r + 2

    # Orden: por serie del gráfico y luego por total descendente.
    orden_grupo = {g: i for i, g in enumerate(b.grupos)}
    det["_g"] = [b.asignacion.get(e, config.OTRAS) for e in det.index]
    det["_o"] = det["_g"].map(lambda g: orden_grupo.get(g, len(orden_grupo)))
    det["_t"] = det[cols].sum(axis=1)
    det = det.sort_values(["_o", "_t"], ascending=[True, False])

    # --- A. Detalle -------------------------------------------------------- #
    ws.write(r, 0, "A. Detalle por entidad (base del cálculo)", f.etiqueta)
    r += 1
    ws.write(r, 0, "Entidad", f.enc_izq)
    ws.write(r, 1, "Se muestra como", f.enc)
    for j, lbl in enumerate(etiquetas):
        ws.write(r, 2 + j, lbl, f.enc)
    # Total por fila solo cuando suma cosas comparables (meses de un año, medios de
    # un periodo). Entre periodos TT/YTD se solaparía, así que no se muestra.
    con_total = b.participacion in ("fila", "total")
    if con_total:
        ws.write(r, 2 + k, "Total", f.enc)
    a_ini = r + 1
    for i, (ent, fila) in enumerate(det.iterrows()):
        rr = a_ini + i
        ws.write(rr, 0, str(ent), f.celda)
        ws.write(rr, 1, fila["_g"], f.grupo)
        for j, c in enumerate(cols):
            v = _valor(fila[c])
            if v:
                ws.write_number(rr, 2 + j, v, f.num(b.unidad))
            else:
                ws.write_blank(rr, 2 + j, None, f.num(b.unidad))
        if con_total:
            ws.write_formula(rr, 2 + k, f"=SUM({xl_rowcol_to_cell(rr, 2)}:"
                                        f"{xl_rowcol_to_cell(rr, 1 + k)})",
                             f.num(b.unidad, True), float(fila["_t"]))
    a_fin = a_ini + len(det) - 1
    r = a_fin + 2

    # --- B. Datos del gráfico (SUMAR.SI sobre A) --------------------------- #
    ws.write(r, 0, "B. Datos del gráfico (fórmulas sobre la tabla A)", f.etiqueta)
    r += 1
    ws.write(r, 0, "Serie en el gráfico", f.enc_izq)
    for j, lbl in enumerate(etiquetas):
        ws.write(r, 1 + j, lbl, f.enc)
    if con_total:
        ws.write(r, 1 + k, "Total", f.enc)
    if b.participacion == "total":
        ws.write(r, 2 + k, "Participación", f.enc)
    b_ini = r + 1
    rango_g = f"$B${a_ini + 1}:$B${a_fin + 1}"
    grupos = [g for g in b.grupos if (det["_g"] == g).any()]
    agregados = det.groupby("_g")[cols].sum()
    for i, g in enumerate(grupos):
        rr = b_ini + i
        ws.write(rr, 0, g, f.celda)
        for j, c in enumerate(cols):
            col_a = xl_col_to_name(2 + j)
            ws.write_formula(
                rr, 1 + j,
                f"=SUMIF({rango_g},{xl_rowcol_to_cell(rr, 0, col_abs=True)},"
                f"{col_a}${a_ini + 1}:{col_a}${a_fin + 1})",
                f.num(b.unidad), _valor(agregados.at[g, c]))
        if con_total:
            ws.write_formula(rr, 1 + k,
                             f"=SUM({xl_rowcol_to_cell(rr, 1)}:{xl_rowcol_to_cell(rr, k)})",
                             f.num(b.unidad, True), float(agregados.loc[g].sum()))
    b_fin = b_ini + len(grupos) - 1
    rt = b_fin + 1
    ws.write(rt, 0, b.etiqueta_total, f.tot_lbl)
    for j in range(k + 1 if con_total else k):
        col = xl_col_to_name(1 + j)
        cache = float(agregados.loc[grupos].sum().sum()) if j == k else \
            float(agregados.loc[grupos, cols[j]].sum())
        ws.write_formula(rt, 1 + j, f"=SUM({col}{b_ini + 1}:{col}{b_fin + 1})",
                         f.num(b.unidad, True), cache)
    if b.participacion == "total":
        tot_ref = xl_rowcol_to_cell(rt, 1 + k, row_abs=True, col_abs=True)
        gran = float(agregados.loc[grupos].sum().sum())
        for i, g in enumerate(grupos + [b.etiqueta_total]):
            rr = b_ini + i
            cache = (float(agregados.loc[grupos].sum().sum()) if g == b.etiqueta_total
                     else float(agregados.loc[g].sum()))
            ws.write_formula(rr, 2 + k, f"=IFERROR({xl_rowcol_to_cell(rr, 1 + k)}/{tot_ref},\"\")",
                             f.pct_tot if g == b.etiqueta_total else f.pct,
                             cache / gran if gran else 0)
    r = rt + 2

    # --- C. Participación --------------------------------------------------- #
    if b.participacion in ("columna", "fila"):
        texto = ("C. Participación de cada serie dentro de cada columna (= B / TOTAL de la columna)"
                 if b.participacion == "columna" else
                 "C. Participación de cada medio dentro de cada fila (= B / Total de la fila)")
        ws.write(r, 0, texto, f.etiqueta)
        r += 1
        ws.write(r, 0, "Serie en el gráfico", f.enc_izq)
        for j, lbl in enumerate(etiquetas):
            ws.write(r, 1 + j, lbl, f.enc)
        c_ini = r + 1
        filas_c = grupos + ([b.etiqueta_total] if b.participacion == "fila" else [])
        for i, g in enumerate(filas_c):
            rr = c_ini + i
            rb = b_ini + i  # misma posición en la tabla B
            es_tot = g == b.etiqueta_total
            ws.write(rr, 0, g, f.tot_lbl if es_tot else f.celda)
            for j in range(k):
                num_ref = xl_rowcol_to_cell(rb, 1 + j)
                if b.participacion == "columna":
                    den_ref = xl_rowcol_to_cell(rt, 1 + j, row_abs=True)
                    num_v = _valor(agregados.at[g, cols[j]])
                    den_v = float(agregados.loc[grupos, cols[j]].sum())
                else:
                    den_ref = xl_rowcol_to_cell(rb, 1 + k, col_abs=True)
                    if es_tot:
                        num_v = float(agregados.loc[grupos, cols[j]].sum())
                        den_v = float(agregados.loc[grupos].sum().sum())
                    else:
                        num_v = _valor(agregados.at[g, cols[j]])
                        den_v = float(agregados.loc[g].sum())
                ws.write_formula(rr, 1 + j, f"=IFERROR({num_ref}/{den_ref},\"\")",
                                 f.pct_tot if es_tot else f.pct, num_v / den_v if den_v else 0)
        r = c_ini + len(filas_c) + 1
    return r + 1


def _escribir_variacion(ws, f: _Formatos, r: int, v: Variacion) -> int:
    ws.write(r, 0, "Fila «Var.% año anterior» de la tabla", f.subtitulo)
    for c in range(1, 6):
        ws.write_blank(r, c, None, f.subtitulo)
    r += 1
    ws.write(r, 0, "El mes en curso se compara contra el mismo tramo del año anterior "
                   "(Periodo = COMPARABLE); el total, YTD contra YTD comparable.", f.nota)
    r += 1
    enc = ["Mes", f"Año {v.anio_actual}", f"Base año {v.anio_anterior}", "Tipo de base", "Var. %"]
    for j, h in enumerate(enc):
        ws.write(r, j, h, f.enc)
    r0 = r + 1
    for i, (mes, fila) in enumerate(v.componentes.iterrows()):
        rr = r0 + i
        ws.write(rr, 0, config.MESES_CORTOS[int(mes) - 1], f.celda)
        ws.write_number(rr, 1, _valor(fila["actual"]), f.num(v.unidad))
        ws.write_number(rr, 2, _valor(fila["base"]), f.num(v.unidad))
        ws.write(rr, 3, fila["tipo_base"], f.celda)
        if fila["aplica"] and fila["base"] > 0:
            ws.write_formula(rr, 4, f"={xl_rowcol_to_cell(rr, 1)}/{xl_rowcol_to_cell(rr, 2)}-1",
                             f.pct, fila["actual"] / fila["base"] - 1)
        else:
            ws.write_blank(rr, 4, None, f.pct)
    rt = r0 + len(v.componentes)
    ws.write(rt, 0, "Total", f.tot_lbl)
    ws.write_number(rt, 1, v.total_actual, f.num(v.unidad, True))
    ws.write_number(rt, 2, v.total_base, f.num(v.unidad, True))
    ws.write(rt, 3, f"YTD {v.anio_anterior} comparable", f.tot_lbl)
    if v.total_base > 0:
        ws.write_formula(rt, 4, f"={xl_rowcol_to_cell(rt, 1)}/{xl_rowcol_to_cell(rt, 2)}-1",
                         f.pct_tot, v.total_actual / v.total_base - 1)
    else:
        ws.write_blank(rt, 4, None, f.pct_tot)
    return rt + 3


def _escribir_entrada(ws, f: _Formatos, e: Entrada) -> None:
    ws.set_column(0, 0, 34)
    ws.set_column(1, 1, 22)
    ws.set_column(2, 40, 14)
    ws.freeze_panes(0, 1)
    ws.write(0, 0, f"Diapositiva {e.n}: {e.titulo}", f.titulo)
    ws.write_url(0, 6, "internal:'Índice'!A1", f.vinculo, string="← Volver al índice")
    filas = [
        ("Sección", e.seccion), ("Subtítulo en el PPT", e.subtitulo),
        ("Universo (filtro)", e.universo), ("Métrica (columna de la base)", e.metrica),
        ("Unidad en el PPT", e.unidad_ppt), ("Cálculo", e.calculo),
        ("Cómo reproducirlo con la hoja «Base»", e.reproducir),
    ] + [("Nota", n) for n in e.notas]
    r = 2
    for lbl, txt in filas:
        if not txt:
            continue
        ws.write(r, 0, lbl, f.etiqueta)
        ws.write(r, 1, txt, f.texto)
        r += 1
    r += 1
    for b in e.bloques:
        r = _escribir_bloque(ws, f, r, b)
    if e.variacion is not None:
        _escribir_variacion(ws, f, r, e.variacion)


def _escribir_indice(ws, f: _Formatos, traza: Trazabilidad) -> None:
    ws.write(0, 0, "Índice de diapositivas", f.titulo)
    enc = ["N° diap.", "Sección", "Tipo", "Título", "Subtítulo en el PPT", "Métrica",
           "Unidad en el PPT", "Universo (filtro)", "Hoja de detalle"]
    anchos = [9, 18, 18, 34, 52, 30, 30, 50, 24]
    for j, (h, a) in enumerate(zip(enc, anchos)):
        ws.write(2, j, h, f.enc)
        ws.set_column(j, j, a)
    for i, e in enumerate(sorted(traza.entradas, key=lambda x: x.n)):
        r = 3 + i
        vals = [e.n, e.seccion, e.tipo, e.titulo, e.subtitulo, e.metrica, e.unidad_ppt,
                e.universo]
        for j, v in enumerate(vals):
            ws.write(r, j, v, f.celda)
        if e.hoja:
            ws.write_url(r, 8, f"internal:'{e.hoja}'!A1", f.vinculo, string=e.hoja)
        else:
            ws.write(r, 8, "—", f.celda)
    ws.freeze_panes(3, 0)
    ws.autofilter(2, 0, 2 + len(traza.entradas), len(enc) - 1)
    if traza.omitidas:
        r = 5 + len(traza.entradas)
        ws.write(r, 0, "Diapositivas omitidas por falta de datos", f.subtitulo)
        for i, a in enumerate(traza.omitidas, start=1):
            ws.write(r + i, 0, a)


def _escribir_leeme(ws, f: _Formatos, parametros: dict[str, Any]) -> None:
    ws.set_column(0, 0, 34)
    ws.set_column(1, 1, 110)
    ws.write(0, 0, "Trazabilidad del Informe de Competencia", f.titulo)
    r = 2
    ws.write(r, 0, "Parámetros de la corrida", f.subtitulo)
    ws.write_blank(r, 1, None, f.subtitulo)
    r += 1
    for k, v in parametros.items():
        ws.write(r, 0, k, f.etiqueta)
        ws.write(r, 1, str(v), f.texto_wrap)
        r += 1
    r += 1
    secciones = [
        ("Cómo usar esta planilla", [
            ("Índice", "Una fila por diapositiva del PPT, en el mismo orden y numeración. "
                       "La última columna vincula a la hoja con los datos de esa diapositiva."),
            ("Hojas «Dnn …»", "Una por gráfico o tabla. Arriba: universo, métrica, unidad, "
                              "cálculo y cómo reproducirlo. Abajo, por cada bloque:"),
            ("   A. Detalle por entidad", "Todas las entidades que suman en el gráfico, en la "
                                          "unidad original (pesos o TRPs), y con qué serie del "
                                          "gráfico se muestran (su nombre u OTRAS)."),
            ("   B. Datos del gráfico", "Fórmulas SUMAR.SI sobre la tabla A: son exactamente "
                                        "los valores dibujados. Si cambia un número en A, "
                                        "B y C se recalculan."),
            ("   C. Participación", "SOS / SOV / mix calculados con fórmulas sobre B."),
            ("Clasificación", "Cada marca × área × empresa de la base con su compañía, "
                              "categoría, marca del informe, BLYS y la regla que la ubicó."),
            ("Diagnóstico", "Marcas que no calzan con la hoja Agrupaciones y regla de respaldo."),
            ("Base", "La base de Megatime con las columnas de clasificación (si se incluyó). "
                     "Permite rehacer cualquier cifra con filtros o una tabla dinámica."),
        ]),
        ("Definiciones", [
            ("TT AAAA", "Año calendario completo (Año = AAAA)."),
            ("YTD AAAA", "Año a la fecha. Año actual: todo lo disponible. Año anterior: "
                         "avisos con Periodo = COMPARABLE (columna «ytd» = VERDADERO en Base)."),
            ("Inversión", "Suma de ValorPublico (pesos). En el PPT: millones en gráficos, "
                          "miles en tablas mensuales."),
            ("TRP", "Suma de la columna del target (ya ponderada a 30\"), solo TV abierta y "
                    "pagada (es_tv = VERDADERO)."),
            ("SOS", "Share of spend: inversión de la serie / inversión total del periodo."),
            ("SOV", "Share of voice: TRPs de la serie / TRPs totales del periodo."),
            ("SOS vs SOV", "Año actual, solo TV: SOS de inversión en TV vs SOV de TRPs."),
            ("OTRAS", "Suma de las entidades fuera del top-N. Gráficos por periodo: se rankea "
                      "por SOS sumado entre periodos (YTD actual y TT anterior pesan doble). "
                      "Tablas mensuales y mix: top-N por inversión del año/periodo. "
                      "SOS vs SOV: top-N por TRPs."),
            ("Medios", "TV Abierta; TV Paga; Radio (+ Radio Regional); Vía Pública (+ Metro); "
                       "Gráfica (Prensa + Revista); Internet."),
        ]),
    ]
    for titulo, filas in secciones:
        ws.write(r, 0, titulo, f.subtitulo)
        ws.write_blank(r, 1, None, f.subtitulo)
        r += 1
        for a, b in filas:
            ws.write(r, 0, a, f.etiqueta)
            ws.write(r, 1, b, f.texto_wrap)
            r += 1
        r += 1


def _escribir_df(ws, f: _Formatos, r: int, df: pd.DataFrame, titulo: str | None = None,
                 como_tabla: bool = False, nombre_tabla: str | None = None) -> int:
    if titulo:
        ws.write(r, 0, titulo, f.subtitulo)
        r += 1
    if df is None or df.empty:
        ws.write(r, 0, "Sin registros.", f.nota)
        return r + 2
    datos = df.astype(object).where(df.notna(), None).values.tolist()
    if como_tabla:
        ws.add_table(r, 0, r + len(datos), len(df.columns) - 1, {
            "data": datos, "name": nombre_tabla, "style": "Table Style Light 1",
            "columns": [{"header": str(c)} for c in df.columns]})
    else:
        for j, c in enumerate(df.columns):
            ws.write(r, j, str(c), f.enc)
        for i, fila in enumerate(datos, start=1):
            for j, v in enumerate(fila):
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    ws.write_number(r + i, j, v, f.clp)
                else:
                    ws.write(r + i, j, v, f.celda)
    return r + len(datos) + 3


COLUMNAS_BASE = [
    config.COL_AREA, config.COL_EMPRESA, config.COL_MARCA, config.COL_ANIO, config.COL_MES,
    config.COL_SEMANA, config.COL_MEDIO, config.COL_PERIODO, config.COL_VALOR,
    *config.TARGETS_TRP.keys(), "es_vino", "es_tv", "ytd", "medio_grp", "compania",
    "categoria", "marca", "marca_blys", "blys_vinos", "blys_total",
]


def escribir_excel(traza: Trazabilidad, *, parametros: dict[str, Any],
                   clasificacion: pd.DataFrame | None, diagnostico: dict[str, pd.DataFrame],
                   base: pd.DataFrame | None = None) -> bytes:
    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf, {"in_memory": True, "nan_inf_to_errors": True})
    f = _Formatos(wb)
    ws_leeme = wb.add_worksheet("LEEME")
    ws_indice = wb.add_worksheet("Índice")
    usados = {"LEEME", "Índice", "Clasificación", "Diagnóstico", "Base"}
    for e in sorted(traza.entradas, key=lambda x: x.n):
        if e.bloques:
            e.hoja = _nombre_hoja(e, usados)
            _escribir_entrada(wb.add_worksheet(e.hoja), f, e)
    _escribir_leeme(ws_leeme, f, parametros)
    _escribir_indice(ws_indice, f, traza)

    ws = wb.add_worksheet("Clasificación")
    ws.write(0, 0, "Clasificación de cada marca × área × empresa", f.titulo)
    ws.set_column(0, 2, 30)
    ws.set_column(3, 30, 22)
    _escribir_df(ws, f, 2, clasificacion, como_tabla=True, nombre_tabla="Clasificacion")
    ws.freeze_panes(3, 1)

    ws = wb.add_worksheet("Diagnóstico")
    ws.write(0, 0, "Diagnóstico de agrupaciones", f.titulo)
    ws.set_column(0, 4, 30)
    r = 2
    for titulo, tabla in diagnostico.items():
        r = _escribir_df(ws, f, r, tabla, titulo)

    if base is not None:
        cols = [c for c in COLUMNAS_BASE if c in base.columns]
        ws = wb.add_worksheet("Base")
        _escribir_df(ws, f, 0, base[cols], como_tabla=True, nombre_tabla="Base")
        ws.set_column(0, len(cols) - 1, 16)
        ws.freeze_panes(1, 0)
    wb.close()
    return buf.getvalue()


def parametros_corrida(op, per, targets: dict[str, str], fusiones: dict[str, str],
                       incluir: list[str] | None, nombres: dict[str, str],
                       origen: dict[str, str] | None = None) -> dict[str, Any]:
    p = {"Generado el": datetime.now().strftime("%d/%m/%Y %H:%M")}
    for k, v in (origen or {}).items():
        p[k] = v
    p.update({
        "Fecha de corte": op.fecha_corte.strftime("%d/%m/%Y"),
        "Periodos": ", ".join(per.etiquetas),
        "Secciones incluidas": ", ".join(nombres[s] for s in (incluir or nombres)),
        "Top-N (gráficos / tablas / mix / SOS-SOV)":
            f"{op.top_grafico} / {op.top_tabla} / {op.top_mix} / {op.top_sov}",
        "Target TRP por sección": "; ".join(
            f"{nombres[k]}: {config.TARGETS_TRP.get(v, v)}" for k, v in targets.items()
            if incluir is None or k in incluir),
        "Fusiones de marca": "; ".join(f"{k} → {v}" for k, v in fusiones.items()) or "—",
    })
    return p
