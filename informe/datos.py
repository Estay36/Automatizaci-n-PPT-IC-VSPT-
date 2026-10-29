"""Carga de la base de Megatime y clasificación de cada aviso.

Resultado: un ``DataFrame`` con las columnas originales más

* ``inv``          inversión (Valor Público, CLP)
* ``medio_grp``    medio agrupado para el mix (TV Abierta, Radio, Gráfica, ...)
* ``es_tv``        True si el aviso es de TV abierta o pagada
* ``es_vino``      True si pertenece al universo vinos (compañía / categoría)
* ``compania``     VSPT, CONCHA Y TORO, SANTA RITA, RAVANAL u OTRAS VIÑAS
* ``categoria``    MASIVOS, BOTELLAS FINAS, ESPUMANTES o COOLERS Y SANGRÍA
* ``marca``        marca con fusiones aplicadas (p. ej. 120 SPRITZ -> 120)
* ``marca_blys``   marca para BLYS (distingue GATO LATA de GATO)
* ``blys_vinos``, ``blys_total``  pertenencia a cada universo BLYS
* ``ytd``          True si el aviso cae en el tramo año-a-la-fecha comparable

Reglas de clasificación (documentadas también en el README):

Compañía
    1. Si la marca está en la lista de una compañía de la hoja Agrupaciones.
    2. Si no, por empresa (``config.EMPRESA_A_COMPANIA``).
    3. Si no, OTRAS VIÑAS.

Categoría
    1. Se buscan las listas de categoría que contienen la marca.
    2. Si la categoría natural del área (``config.AREA_A_CATEGORIA``) está entre
       ellas, se usa esa (desempata marcas como Undurraga o Valdivieso).
    3. Si la marca está en una sola lista, se usa esa (p. ej. Donnaluna figura
       en Botellas Finas aunque su área sea Vino-Cooler-Sangría).
    4. Si no está en ninguna lista, se usa la categoría natural del área y la
       marca se informa en el diagnóstico para completar la hoja.

BLYS
    Marca listada en BLYS VINOS / BLYS TOTAL, o marca no listada cuya área es
    propia de BLYS. Las marcas bajo "Tentativo incluir a futuro" quedan fuera.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from typing import BinaryIO

import pandas as pd

from . import config
from .agrupaciones import Agrupaciones, clave


# --------------------------------------------------------------------------- #
# Carga
# --------------------------------------------------------------------------- #
def leer_datos(fuente: str | BinaryIO) -> pd.DataFrame:
    """Lee la hoja ``datos``. Usa calamine si está disponible (mucho más rápido)."""
    try:
        df = pd.read_excel(fuente, sheet_name=config.HOJA_DATOS, engine="calamine")
    except (ImportError, ValueError):
        if hasattr(fuente, "seek"):
            fuente.seek(0)
        df = pd.read_excel(fuente, sheet_name=config.HOJA_DATOS)
    faltan = [c for c in config.COLUMNAS_REQUERIDAS if c not in df.columns]
    if faltan:
        raise ValueError(f"A la hoja '{config.HOJA_DATOS}' le faltan columnas: {faltan}")
    for c in (config.COL_ANIO, config.COL_MES, config.COL_SEMANA):
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).astype(int)
    df[config.COL_VALOR] = pd.to_numeric(df[config.COL_VALOR], errors="coerce").fillna(0.0)
    for t in config.TARGETS_TRP:
        if t in df.columns:
            df[t] = pd.to_numeric(df[t], errors="coerce").fillna(0.0)
    for c in (config.COL_AREA, config.COL_EMPRESA, config.COL_MARCA, config.COL_MEDIO):
        df[c] = df[c].fillna("SIN DATO").astype(str).str.strip()
    return df


# --------------------------------------------------------------------------- #
# Clasificación
# --------------------------------------------------------------------------- #
@dataclass
class Diagnostico:
    """Marcas/empresas que no calzaron con la hoja Agrupaciones."""

    sin_categoria: pd.DataFrame
    compania_por_respaldo: pd.DataFrame
    blys_no_listadas: pd.DataFrame
    medios_sin_mapeo: list[str]
    # Una fila por (marca, área, empresa) con el resultado y la regla aplicada.
    clasificacion: pd.DataFrame | None = None


def _tabla_diag(df: pd.DataFrame, mask: pd.Series, cols: list[str]) -> pd.DataFrame:
    if not mask.any():
        return pd.DataFrame(columns=cols + ["Inversión (MM$)"])
    t = (df.loc[mask].groupby(cols, as_index=False)["inv"].sum()
         .sort_values("inv", ascending=False))
    t["Inversión (MM$)"] = (t.pop("inv") / 1e6).round(1)
    return t.reset_index(drop=True)


def _tabla_clasificacion(df: pd.DataFrame, mapa: pd.DataFrame) -> pd.DataFrame:
    """Resultado de la clasificación por combinación, con inversión y TRPs totales."""
    claves = [config.COL_MARCA, config.COL_AREA, config.COL_EMPRESA]
    trp = [t for t in config.TARGETS_TRP if t in df.columns]
    tot = df.groupby(claves).agg(
        **{"Inversión total (CLP)": ("inv", "sum"), "Años": (config.COL_ANIO, "unique")},
        **{f"TRP {config.TARGETS_TRP[t]}": (t, "sum") for t in trp}).reset_index()
    tot["Años"] = tot["Años"].map(lambda a: ", ".join(str(x) for x in sorted(a)))
    t = mapa.merge(tot, on=claves, how="left")
    es_vino = t[config.COL_AREA].isin(config.AREAS_VINO)
    t.loc[~es_vino, ["compania", "categoria", "regla_compania", "regla_categoria"]] = (
        "— (fuera del universo vinos)")
    cols = claves + ["compania", "regla_compania", "categoria", "regla_categoria", "marca",
                     "marca_blys", "blys_vinos", "blys_total", "regla_blys",
                     "Inversión total (CLP)"] + [f"TRP {config.TARGETS_TRP[x]}" for x in trp] + ["Años"]
    return (t[cols].sort_values([config.COL_AREA, "Inversión total (CLP)"],
                                ascending=[True, False]).reset_index(drop=True))


def clasificar(df: pd.DataFrame, ag: Agrupaciones,
               fusiones: dict[str, str] | None = None) -> tuple[pd.DataFrame, Diagnostico]:
    fusiones = config.FUSIONES_DEFECTO if fusiones is None else fusiones
    df = df.copy()
    area, marca_raw = df[config.COL_AREA], df[config.COL_MARCA].str.upper()
    df["inv"] = df[config.COL_VALOR]
    df["medio_grp"] = df[config.COL_MEDIO].map(config.MEDIOS)
    df["es_tv"] = df[config.COL_MEDIO].isin(config.MEDIOS_TV)
    df["es_vino"] = area.isin(config.AREAS_VINO)

    # Trabajamos sobre combinaciones únicas (marca, área, empresa): son pocas.
    combos = df[[config.COL_MARCA, config.COL_AREA, config.COL_EMPRESA]].drop_duplicates()
    filas = []
    alias = {clave(k): clave(v) for k, v in config.ALIAS_MARCAS.items()}
    emp_a_comp = {clave(k): v for k, v in config.EMPRESA_A_COMPANIA.items()}
    fus = {clave(k): v.upper() for k, v in fusiones.items()}
    for marca, ar, emp in combos.itertuples(index=False):
        km = alias.get(clave(marca), clave(marca))
        # Compañía
        comp = ag.compania_de_marca(km)
        comp_resp = comp is None
        regla_comp = "Marca en lista de la compañía (hoja Agrupaciones)"
        if comp is None:
            comp = emp_a_comp.get(clave(emp), config.COMPANIA_OTRAS)
            if comp != config.COMPANIA_OTRAS:
                regla_comp = "Marca no listada: se asigna por empresa"
            elif clave(emp) in ag.empresas_otras:
                regla_comp = "Empresa en lista OTRAS VIÑAS"
            else:
                regla_comp = "Sin lista ni empresa conocida: OTRAS VIÑAS"
        # Categoría
        natural = config.AREA_A_CATEGORIA.get(ar)
        cands = ag.categorias_de(km)
        if natural in cands:
            cat = natural
            regla_cat = ("Marca en lista de la categoría" if len(cands) == 1 else
                         f"Marca en {len(cands)} listas: desempata el área")
        elif len(cands) == 1:
            cat = cands[0]
            regla_cat = "Marca en una sola lista (distinta a la de su área)"
        elif cands:
            cat = next(c for c in config.ORDEN_CATEGORIAS if c in cands)
            regla_cat = "Marca en varias listas, ninguna de su área: primera en orden"
        else:
            cat = natural
            regla_cat = "Marca no listada: categoría natural del área"
        # BLYS
        nombre_blys = config.MARCA_POR_AREA.get((marca.upper(), ar), marca.upper())
        kb = alias.get(clave(nombre_blys), clave(nombre_blys))
        # Una marca "mencionada" (listada o en futuro) se rige solo por las listas;
        # una no mencionada entra si su área es propia de BLYS.
        mencionada = kb in ag.blys_total or kb in ag.blys_futuro
        en_vinos = kb in ag.blys_vinos or (
            not mencionada and ar in config.AREAS_BLYS_VINOS)
        en_total = en_vinos or kb in ag.blys_total or (
            not mencionada and ar in config.AREAS_BLYS_TOTAL)
        listada = mencionada
        if kb in ag.blys_futuro:
            regla_blys = "En «Tentativo incluir a futuro»: excluida"
        elif mencionada:
            regla_blys = "Marca en lista BLYS de la hoja"
        elif en_total:
            regla_blys = "No listada; incluida por área BLYS"
        else:
            regla_blys = "No es BLYS"
        filas.append({
            config.COL_MARCA: marca, config.COL_AREA: ar, config.COL_EMPRESA: emp,
            "compania": comp, "_comp_resp": comp_resp,
            "categoria": cat, "_cat_resp": not cands,
            "marca": fus.get(km, marca.upper()),
            "marca_blys": nombre_blys,
            "blys_vinos": en_vinos, "blys_total": en_total,
            "_blys_no_listada": (en_total and not listada),
            "regla_compania": regla_comp, "regla_categoria": regla_cat,
            "regla_blys": regla_blys,
        })
    mapa = pd.DataFrame(filas)
    reglas = ["regla_compania", "regla_categoria", "regla_blys"]
    df = df.merge(mapa.drop(columns=reglas), on=[config.COL_MARCA, config.COL_AREA, config.COL_EMPRESA],
                  how="left", validate="many_to_one")
    df.loc[~df["es_vino"], ["categoria", "compania"]] = None

    diag = Diagnostico(
        sin_categoria=_tabla_diag(df, df["es_vino"] & df["_cat_resp"],
                                  [config.COL_MARCA, config.COL_AREA, "categoria"]),
        compania_por_respaldo=_tabla_diag(
            df, df["es_vino"] & df["_comp_resp"] & (df["compania"] != config.COMPANIA_OTRAS),
            [config.COL_MARCA, config.COL_EMPRESA, "compania"]),
        blys_no_listadas=_tabla_diag(df, df["_blys_no_listada"],
                                     ["marca_blys", config.COL_AREA]),
        medios_sin_mapeo=sorted(df.loc[df["medio_grp"].isna(), config.COL_MEDIO].unique()),
        clasificacion=_tabla_clasificacion(df, mapa),
    )
    df["medio_grp"] = df["medio_grp"].fillna("Internet")
    return df.drop(columns=[c for c in df.columns if c.startswith("_")]), diag


# --------------------------------------------------------------------------- #
# Periodos
# --------------------------------------------------------------------------- #
@dataclass
class Periodos:
    """Define TT (años completos) y YTD (año a la fecha comparable)."""

    anio_actual: int
    anio_inicio: int
    fecha_corte: date

    @property
    def anio_anterior(self) -> int:
        return self.anio_actual - 1

    @property
    def etiquetas(self) -> list[str]:
        tt = [f"TT {a}" for a in range(self.anio_inicio, self.anio_actual)]
        return tt + [f"YTD {self.anio_anterior}", f"YTD {self.anio_actual}"]


def marcar_ytd(df: pd.DataFrame, anio_actual: int) -> pd.Series:
    """True para avisos dentro del tramo año-a-la-fecha.

    Año actual: todo. Años anteriores: columna ``Periodo == "COMPARABLE"`` si
    existe; si no, semanas <= última semana con datos del año actual.
    """
    anio = df[config.COL_ANIO]
    if config.COL_PERIODO in df.columns:
        comp = df[config.COL_PERIODO].astype(str).str.upper().str.strip() == "COMPARABLE"
    else:
        ult_sem = df.loc[anio == anio_actual, config.COL_SEMANA].max()
        comp = df[config.COL_SEMANA] <= ult_sem
    return (anio == anio_actual) | ((anio < anio_actual) & comp)


def periodo_de(df: pd.DataFrame, per: Periodos) -> pd.DataFrame:
    """Expande cada aviso a los periodos en que participa (columna ``periodo``).

    Un aviso de 2025 comparable aparece en "TT 2025" y en "YTD 2025".
    """
    anio = df[config.COL_ANIO]
    partes = []
    tt = df[(anio >= per.anio_inicio) & (anio < per.anio_actual)]
    partes.append(tt.assign(periodo="TT " + tt[config.COL_ANIO].astype(str)))
    ytd_ant = df[(anio == per.anio_anterior) & df["ytd"]]
    partes.append(ytd_ant.assign(periodo=f"YTD {per.anio_anterior}"))
    partes.append(df[anio == per.anio_actual].assign(periodo=f"YTD {per.anio_actual}"))
    return pd.concat(partes, ignore_index=True)


def preparar(df_raw: pd.DataFrame, ag: Agrupaciones, fecha_corte: date,
             anio_inicio: int | None = None,
             fusiones: dict[str, str] | None = None) -> tuple[pd.DataFrame, Periodos, Diagnostico]:
    """Clasifica la base y define los periodos a partir de la fecha de corte."""
    df, diag = clasificar(df_raw, ag, fusiones)
    anio_actual = fecha_corte.year
    df = df[df[config.COL_ANIO] <= anio_actual].copy()
    inicio = anio_inicio or int(df[config.COL_ANIO].min())
    df["ytd"] = marcar_ytd(df, anio_actual)
    return df, Periodos(anio_actual, inicio, fecha_corte), diag


def fecha_corte_desde_nombre(nombre: str) -> date | None:
    """Extrae la fecha de nombres como ``Informe_VSPT_2022-02oct2026.xlsx``."""
    import re

    meses = {m[:3].lower(): i + 1 for i, m in enumerate(config.MESES)}
    meses.update({"sep": 9, "set": 9})
    m = re.search(r"(\d{1,2})([a-zA-Z]{3})(\d{4})", nombre)
    if not m or m.group(2).lower() not in meses:
        return None
    try:
        return date(int(m.group(3)), meses[m.group(2).lower()], int(m.group(1)))
    except ValueError:
        return None
