"""Generador del Informe de Competencia VSPT (Streamlit).

Ejecutar localmente:  streamlit run app.py
"""
from __future__ import annotations

import io
from datetime import date

import pandas as pd
import streamlit as st

from informe import config
from informe.agrupaciones import Agrupaciones, leer_agrupaciones
from informe.datos import fecha_corte_desde_nombre, leer_datos, preparar
from informe.diapositivas import Opciones
from informe.reporte import NOMBRES_SECCION, TARGETS_DEFECTO, generar

st.set_page_config(page_title="Informe de Competencia VSPT", page_icon="📊", layout="wide")


# --------------------------------------------------------------------------- #
# Carga (cacheada por contenido del archivo)
# --------------------------------------------------------------------------- #
@st.cache_data(show_spinner="Leyendo base de Megatime…", max_entries=3)
def cargar_base(contenido: bytes) -> pd.DataFrame:
    return leer_datos(io.BytesIO(contenido))


@st.cache_data(show_spinner=False, max_entries=5)
def cargar_agrupaciones(contenido: bytes | None) -> Agrupaciones:
    if contenido is None:
        return leer_agrupaciones(config.AGRUPACIONES_DEFECTO)
    return leer_agrupaciones(io.BytesIO(contenido))


def parsear_fusiones(texto: str) -> dict[str, str]:
    fusiones = {}
    for linea in texto.splitlines():
        if "->" in linea:
            origen, destino = (p.strip().upper() for p in linea.split("->", 1))
            if origen and destino:
                fusiones[origen] = destino
    return fusiones


# --------------------------------------------------------------------------- #
# Barra lateral
# --------------------------------------------------------------------------- #
st.title("Informe de Competencia – VSPT")
st.caption("Genera el PPT de competencia a partir del export de Megatime y la "
           "planilla de agrupaciones (Categorías Vinos).")

with st.sidebar:
    st.header("1 · Archivos")
    base_file = st.file_uploader("Base Megatime (hoja «datos»)", type=["xlsx"])
    agr_file = st.file_uploader(
        "Categorías Vinos (hoja «Agrupaciones»)", type=["xlsx"],
        help="Opcional. Si no se carga, se usa la versión incluida en el repositorio.")

if base_file is None:
    st.info("Carga la base de Megatime en la barra lateral para comenzar.")
    st.stop()

try:
    raw = cargar_base(base_file.getvalue())
    ag = cargar_agrupaciones(agr_file.getvalue() if agr_file else None)
except ValueError as e:
    st.error(str(e))
    st.stop()

anios = sorted(raw[config.COL_ANIO].unique())
fecha_sugerida = fecha_corte_desde_nombre(base_file.name) or date(int(max(anios)), 12, 31)

with st.sidebar:
    st.header("2 · Periodo")
    fecha_corte = st.date_input(
        "Fecha de corte", value=fecha_sugerida, format="DD/MM/YYYY",
        help="Se lee del nombre del archivo (p. ej. «02oct2026»). Define el año actual "
             "y el texto de los subtítulos; el tramo YTD del año anterior se toma de la "
             "columna «Periodo» = COMPARABLE.")
    anio_inicio = st.selectbox("Primer año de la serie TT",
                               [a for a in anios if a < fecha_corte.year] or anios, index=0)

    st.header("3 · Contenido")
    titulo = st.text_input("Título de portada", "INFORME DE COMPETENCIA – VSPT")
    incluir = st.multiselect("Secciones", list(NOMBRES_SECCION),
                             default=list(NOMBRES_SECCION),
                             format_func=NOMBRES_SECCION.get)
    c1, c2 = st.columns(2)
    top_grafico = c1.number_input("Top marcas gráficos", 3, 20, 10)
    top_tabla = c2.number_input("Top marcas tablas", 3, 20, 10)
    top_mix = c1.number_input("Top marcas mix", 3, 20, 10)
    top_sov = c2.number_input("Top marcas SOS/SOV", 2, 15, 8)

    with st.expander("Target de TRP por sección"):
        opciones_target = [t for t in config.TARGETS_TRP if t in raw.columns]
        targets = {}
        for clave, nombre in NOMBRES_SECCION.items():
            defecto = TARGETS_DEFECTO[clave]
            targets[clave] = st.selectbox(
                nombre, opciones_target,
                index=opciones_target.index(defecto) if defecto in opciones_target else 0,
                format_func=config.TARGETS_TRP.get, key=f"t_{clave}")

    with st.expander("Fusiones de marca"):
        texto_fus = st.text_area(
            "Una por línea: ORIGEN -> DESTINO",
            "\n".join(f"{k} -> {v}" for k, v in config.FUSIONES_DEFECTO.items()),
            help="Se aplican a las secciones por marca de cada categoría (Masivos, "
                 "Botellas finas, …). BLYS mantiene las marcas separadas.")
        fusiones = parsear_fusiones(texto_fus)

    st.header("4 · Trazabilidad")
    con_traza = st.checkbox(
        "Generar planilla de trazabilidad (Excel)", value=True,
        help="Una hoja por gráfico/tabla con los datos y fórmulas que lo originan, más la "
             "clasificación de cada marca.")
    con_base = st.checkbox(
        "Incluir base clasificada en la planilla", value=True, disabled=not con_traza,
        help="Agrega la base de Megatime con las columnas de clasificación para rehacer "
             "cualquier cifra con filtros o tablas dinámicas (≈ +4 MB, unos segundos más).")

# --------------------------------------------------------------------------- #
# Resumen y diagnóstico
# --------------------------------------------------------------------------- #
df, per, diag = preparar(raw, ag, fecha_corte, anio_inicio, fusiones)

m1, m2, m3, m4 = st.columns(4)
m1.metric("Avisos en la base", f"{len(raw):,}".replace(",", "."))
m2.metric("Periodos", f"{per.etiquetas[0]} → {per.etiquetas[-1]}")
inv_ytd = df.loc[df["es_vino"] & (df[config.COL_ANIO] == per.anio_actual), "inv"].sum()
m3.metric(f"Inversión vinos YTD {per.anio_actual}", f"$ {inv_ytd / 1e6:,.0f} MM".replace(",", "."))
m4.metric("Fecha de corte", fecha_corte.strftime("%d/%m/%Y"))

with st.expander("Diagnóstico de agrupaciones", expanded=bool(len(diag.sin_categoria))):
    st.markdown("Marcas y empresas que **no** calzan con la hoja *Agrupaciones*. "
                "Se clasifican con la regla de respaldo indicada; para cambiarlo, "
                "agrégalas a la hoja y vuelve a cargarla.")
    d1, d2 = st.columns(2)
    d1.markdown("**Marcas de vino sin categoría en la hoja** (se usa la categoría de su área)")
    d1.dataframe(diag.sin_categoria, hide_index=True, width="stretch")
    d2.markdown("**Marcas no listadas en su compañía** (asignadas por empresa)")
    d2.dataframe(diag.compania_por_respaldo, hide_index=True, width="stretch")
    d1.markdown("**Marcas BLYS no listadas** (incluidas por pertenecer a un área BLYS)")
    d1.dataframe(diag.blys_no_listadas, hide_index=True, width="stretch")
    d2.markdown("**Listas leídas de la hoja**")
    d2.dataframe(pd.Series(ag.resumen(), name="N° elementos"), width="stretch")
    if ag.comentarios:
        d2.markdown("**Comentarios de la hoja**\n\n" + "\n".join(f"- {c.lstrip('* ')}" for c in ag.comentarios))
    if diag.medios_sin_mapeo:
        st.warning(f"Medios sin mapeo (se tratan como Internet): {diag.medios_sin_mapeo}")

# --------------------------------------------------------------------------- #
# Generación
# --------------------------------------------------------------------------- #
if st.button("Generar PPT", type="primary", disabled=not incluir):
    barra = st.progress(0.0, "Preparando…")
    op = Opciones(fecha_corte=fecha_corte, anio_inicio=anio_inicio,
                  top_grafico=int(top_grafico), top_tabla=int(top_tabla),
                  top_mix=int(top_mix), top_sov=int(top_sov), titulo=titulo)
    origen = {"Base Megatime": base_file.name,
              "Agrupaciones": agr_file.name if agr_file else
              f"{config.AGRUPACIONES_DEFECTO.name} (incluida en el repositorio)"}
    res = generar(raw, ag, op, incluir=incluir, targets=targets, fusiones=fusiones,
                  trazabilidad=con_traza, incluir_base=con_base, origen=origen,
                  progreso=lambda p, m: barra.progress(p, m))
    prefijo = f"{fecha_corte:%Y-%m-%d}_Informe_de_Competencia_-_VSPT"
    st.session_state["pptx"] = res.pptx
    st.session_state["xlsx"] = res.xlsx
    st.session_state["avisos"] = res.avisos
    st.session_state["prefijo"] = prefijo

if "pptx" in st.session_state:
    st.success("Informe generado.")
    b1, b2, _ = st.columns([1, 1.4, 2])
    b1.download_button("Descargar PPT", st.session_state["pptx"],
                       file_name=f"{st.session_state['prefijo']}.pptx", type="primary",
                       mime="application/vnd.openxmlformats-officedocument."
                            "presentationml.presentation")
    if st.session_state.get("xlsx"):
        b2.download_button("Descargar planilla de trazabilidad", st.session_state["xlsx"],
                           file_name=f"{st.session_state['prefijo']}_trazabilidad.xlsx",
                           mime="application/vnd.openxmlformats-officedocument."
                                "spreadsheetml.sheet")
        st.caption("La planilla tiene una hoja por diapositiva (mismo número que en el PPT) "
                   "con los datos de origen, la agrupación en OTRAS y fórmulas que "
                   "reconstruyen cada cifra. Empieza por la hoja «LEEME».")
    for aviso in st.session_state.get("avisos", []):
        st.caption(f"ℹ️ {aviso}")
