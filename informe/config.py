"""Parámetros fijos del informe: rutas, nombres de columnas, mapeos y textos.

Todo lo que un analista podría querer ajustar sin tocar la lógica vive aquí o en
``assets/colores.json``. La aplicación Streamlit permite sobrescribir en tiempo de
ejecución los valores marcados como *configurables*.
"""
from __future__ import annotations

import json
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
ASSETS = RAIZ / "assets"
PLANTILLA = ASSETS / "plantilla.pptx"
LOGO_PORTADA = ASSETS / "logo_portada.jpg"
AGRUPACIONES_DEFECTO = ASSETS / "Categorias_Vinos.xlsx"
COLORES_JSON = ASSETS / "colores.json"

# --------------------------------------------------------------------------- #
# Datos de entrada (hoja "datos" del export de Megatime)
# --------------------------------------------------------------------------- #
HOJA_DATOS = "datos"
COL_AREA = "Área"
COL_EMPRESA = "Empresa_Clasif"
COL_MARCA = "Marcas reunidas"
COL_ANIO = "Año"
COL_MES = "Mes"
COL_SEMANA = "Semana"
COL_MEDIO = "Medio"
COL_VALOR = "ValorPublico"
COL_PERIODO = "Periodo"  # opcional: "COMPARABLE" marca el tramo YTD del año anterior

COLUMNAS_REQUERIDAS = [
    COL_AREA, COL_EMPRESA, COL_MARCA, COL_ANIO, COL_MES, COL_SEMANA,
    COL_MEDIO, COL_VALOR,
]

# Targets de TRP disponibles (columnas ya ponderadas a 30").
TARGETS_TRP = {
    "HM25_49MB_Pond": "HM 25-49 Medio Bajo",
    "HM35_49A_Pond": "HM 35-49 Alto",
    "HM18_49AMB_Pond": "HM 18-49 Alto, Medio y Bajo",
}
TARGET_DEFECTO = "HM25_49MB_Pond"

# Medios de la base -> grupo que se muestra en "Mix de medios".
MEDIOS = {
    "TV Abierta": "TV Abierta",
    "TV Paga": "TV Paga",
    "Radio": "Radio",
    "Radio Regional": "Radio",
    "Vía Pública": "Vía Pública",
    "Metro": "Vía Pública",
    "Prensa": "Gráfica",
    "Revista": "Gráfica",
    "Internet": "Internet",
}
ORDEN_MEDIOS = ["TV Abierta", "TV Paga", "Radio", "Vía Pública", "Gráfica", "Internet"]
MEDIOS_TV = {"TV Abierta", "TV Paga"}  # base de inversión para SOI vs SOV

# --------------------------------------------------------------------------- #
# Universos
# --------------------------------------------------------------------------- #
# Áreas que componen el universo "vinos" (compañía y categoría).
AREAS_VINO = ["VINO MASIVO", "VINO PREMIUM", "ESPUMANTES", "VINO-COOLER-SANGRÍA"]

# Categoría "natural" de cada área. Se usa para desempatar marcas presentes en más
# de una lista de la hoja Agrupaciones y como respaldo para marcas no listadas.
AREA_A_CATEGORIA = {
    "VINO MASIVO": "MASIVOS",
    "VINO PREMIUM": "BOTELLAS FINAS",
    "ESPUMANTES": "ESPUMANTES",
    "VINO-COOLER-SANGRÍA": "COOLERS Y SANGRÍA",
}
ORDEN_CATEGORIAS = ["MASIVOS", "BOTELLAS FINAS", "ESPUMANTES", "COOLERS Y SANGRÍA"]

# Empresas de la base que corresponden a cada compañía del informe (respaldo cuando
# la marca no está en las listas de la hoja Agrupaciones).
EMPRESA_A_COMPANIA = {
    "VIÑA SAN PEDRO - TARAPACÁ": "VSPT",
    "VIÑA CONCHA Y TORO": "CONCHA Y TORO",
    "VIÑA SANTA RITA": "SANTA RITA",
    "VIÑA RAVANAL": "RAVANAL",
}
COMPANIA_OTRAS = "OTRAS VIÑAS"

# Áreas cuyas marcas no listadas se consideran parte de BLYS.
AREAS_BLYS_VINOS = ["VINO-COOLER-SANGRÍA"]
AREAS_BLYS_TOTAL = ["VINO-COOLER-SANGRÍA", "CERVEZAS", "ICE", "BLYS"]

# Variantes de una marca que se distinguen por el área (p. ej. Gato en lata).
MARCA_POR_AREA = {("GATO", "VINO-COOLER-SANGRÍA"): "GATO LATA"}

# Equivalencias entre el nombre en la hoja Agrupaciones y el de la base.
ALIAS_MARCAS = {"CORONA TROPICAL": "CORONA"}

# Fusiones de marca (configurable desde la app). Según comentario de la hoja
# Agrupaciones: "Juntar 120 con 120 Spritz como una marca, con exportación".
FUSIONES_DEFECTO = {
    "120 SPRITZ": "120",
    "EXPORTACIÓN SELECTO ICE": "EXPORTACIÓN",
}

OTRAS = "OTRAS"

# --------------------------------------------------------------------------- #
# Textos
# --------------------------------------------------------------------------- #
MESES = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO",
         "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]
MESES_CORTOS = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct",
                "Nov", "Dic"]

FUENTE_MILLONES = "Fuente: Megatime, Valor Público en millones de pesos chilenos."
FUENTE_MILES = "Fuente: Megatime, Valor Público en miles de pesos chilenos."
FUENTE_TRP = ("Fuente: Megatime. Time Ibope, {target}, calidades distintas a "
              "“video color” ponderadas a 30 segundos.")

# Nombres de los diseños en assets/plantilla.pptx (ver tools/crear_plantilla.py).
DISENO_PORTADA = "VSPT Portada"
DISENO_SEPARADOR = "VSPT Separador"
DISENO_SUBSEPARADOR = "VSPT Subseparador"
DISENO_CONTENIDO = "VSPT Contenido"
DISENO_CONTENIDO_LATERAL = "VSPT Contenido lateral"
DISENO_TABLA = "VSPT Tabla"

COLOR_MARCA_TEXTO = "27B088"  # verde corporativo de títulos de separadores


def cargar_colores(ruta: Path = COLORES_JSON) -> dict[str, str]:
    """Colores fijos por entidad (marca, compañía, categoría, medio)."""
    with open(ruta, encoding="utf-8") as fh:
        datos = json.load(fh)
    colores: dict[str, str] = {}
    for grupo in datos.values():
        if isinstance(grupo, dict):
            colores.update({k.upper(): v.upper() for k, v in grupo.items()})
    return colores


def paleta_respaldo(ruta: Path = COLORES_JSON) -> list[str]:
    with open(ruta, encoding="utf-8") as fh:
        return [c.upper() for c in json.load(fh)["_respaldo"]]
