"""Agregaciones que alimentan cada tipo de diapositiva.

Convenciones
------------
* ``dim``: columna de la entidad (``compania``, ``categoria``, ``marca``, ``marca_blys``).
* ``valor``: ``"inv"`` para inversión o el nombre de la columna de TRP del target.
* Las tablas devueltas tienen entidades en columnas y periodos/meses en filas, salvo
  que se indique lo contrario. Los valores van en la unidad original (CLP o TRP);
  el escalado a millones/miles se hace al dibujar.
"""
from __future__ import annotations

from dataclasses import dataclass

import pandas as pd

from . import config
from .datos import Periodos, periodo_de


# --------------------------------------------------------------------------- #
# Selección de entidades
# --------------------------------------------------------------------------- #
def ranking(tabla: pd.DataFrame, n: int, filas_clave: list[str]) -> list[str]:
    """Top-``n`` entidades combinando varias filas de referencia.

    Se toma la unión de los top-``n`` de cada fila clave (p. ej. YTD actual y TT
    del año anterior) y se ordena por la suma de esas filas. Así una marca grande
    el año pasado que hoy no invierte no queda escondida en OTRAS.
    """
    filas = [f for f in filas_clave if f in tabla.index]
    if not filas or tabla.empty:
        return []
    ref = tabla.loc[filas].fillna(0)
    candidatas: set[str] = set()
    for f in filas:
        fila = ref.loc[f]
        candidatas |= set(fila[fila > 0].nlargest(n).index)
    orden = ref.sum().loc[list(candidatas)].sort_values(ascending=False)
    return [e for e in orden.index if e != config.OTRAS][:n]


def con_otras(tabla: pd.DataFrame, keep: list[str], orden_fijo: list[str] | None = None
              ) -> pd.DataFrame:
    """Deja las columnas ``keep`` y suma el resto en OTRAS (si aporta algo)."""
    if orden_fijo is not None:
        cols = [c for c in orden_fijo if c in tabla.columns]
        return tabla.reindex(columns=cols)
    resto = [c for c in tabla.columns if c not in keep]
    out = tabla.reindex(columns=keep).copy()
    if resto:
        otras = tabla[resto].sum(axis=1, min_count=1)
        if otras.fillna(0).abs().sum() > 0:
            out[config.OTRAS] = otras
    return out


# --------------------------------------------------------------------------- #
# Por periodo (TT / YTD)
# --------------------------------------------------------------------------- #
def por_periodo(df: pd.DataFrame, per: Periodos, dim: str, valor: str) -> pd.DataFrame:
    """Filas = periodos (TT 2022 ... YTD actual), columnas = entidades."""
    x = periodo_de(df, per)
    t = x.pivot_table(index="periodo", columns=dim, values=valor, aggfunc="sum")
    t = t.reindex(per.etiquetas)
    t.columns.name = None
    return t.where(t.abs() > 1e-9)


def ranking_ponderado(tabla: pd.DataFrame, n: int, pesos: dict[str, float]) -> list[str]:
    """Top-``n`` por participación ponderada entre periodos.

    El puntaje de una entidad es la suma de su SOS en cada periodo por el peso
    del periodo. Así una marca relevante en años anteriores (p. ej. con fuerte
    inversión en 2023) no queda escondida en OTRAS, pero los periodos recientes
    pesan más.
    """
    if tabla.empty:
        return []
    sh = participacion(tabla).fillna(0)
    w = pd.Series({f: pesos.get(f, 1.0) for f in sh.index})
    puntaje = sh.mul(w, axis=0).sum()
    puntaje = puntaje[(puntaje > 0) & (puntaje.index != config.OTRAS)]
    return list(puntaje.sort_values(ascending=False).index[:n])


def seleccionar(tabla: pd.DataFrame, per: Periodos, n: int,
                orden_fijo: list[str] | None = None) -> pd.DataFrame:
    """Top-``n`` entidades (+ OTRAS), con más peso para YTD actual y TT anterior."""
    pesos = {f"YTD {per.anio_actual}": 2.0, f"TT {per.anio_anterior}": 2.0}
    keep = ranking_ponderado(tabla, n, pesos) if orden_fijo is None else []
    return con_otras(tabla, keep, orden_fijo)


def participacion(tabla: pd.DataFrame) -> pd.DataFrame:
    tot = tabla.sum(axis=1, min_count=1)
    return tabla.div(tot.where(tot > 0), axis=0)


# --------------------------------------------------------------------------- #
# Mensual
# --------------------------------------------------------------------------- #
@dataclass
class BloqueMensual:
    anio: int
    valores: pd.DataFrame  # filas = entidades (+OTRAS), columnas = 1..12
    total_mes: pd.Series   # index 1..12
    total: float


def mensual(df: pd.DataFrame, anio: int, dim: str, valor: str, n: int,
            orden_fijo: list[str] | None = None) -> BloqueMensual:
    x = df[df[config.COL_ANIO] == anio]
    t = x.pivot_table(index=config.COL_MES, columns=dim, values=valor, aggfunc="sum")
    t = t.reindex(range(1, 13))
    t.columns.name = None
    if orden_fijo is None:
        tot = t.sum().sort_values(ascending=False)
        keep = [e for e in tot[tot > 0].index if e != config.OTRAS][:n]
    else:
        keep = []
    t = con_otras(t, keep, orden_fijo)
    t = t.where(t.abs() > 1e-9)
    return BloqueMensual(anio, t.T, t.sum(axis=1, min_count=1), float(t.sum().sum()))


def variacion_mensual(df: pd.DataFrame, per: Periodos, valor: str
                      ) -> tuple[pd.Series, float | None]:
    """Var % mensual año actual vs. anterior y var % YTD vs. YTD comparable.

    Para el mes en curso se compara contra el tramo comparable del año anterior
    (no contra el mes completo), igual que el total.
    """
    anio = df[config.COL_ANIO]
    act = df[anio == per.anio_actual].groupby(config.COL_MES)[valor].sum()
    ant_full = df[anio == per.anio_anterior].groupby(config.COL_MES)[valor].sum()
    ant_ytd = df[(anio == per.anio_anterior) & df["ytd"]].groupby(config.COL_MES)[valor].sum()
    ult_mes = int(act[act > 0].index.max()) if (act > 0).any() else 0
    var = {}
    for m in range(1, 13):
        if m > ult_mes:
            var[m] = None
            continue
        base = ant_ytd.get(m, 0.0) if m == ult_mes else ant_full.get(m, 0.0)
        a = act.get(m, 0.0)
        var[m] = (a / base - 1) if base > 0 else None
    base_tot = ant_ytd.sum()
    tot = (act.sum() / base_tot - 1) if base_tot > 0 else None
    return pd.Series(var), tot


def share_mensual(df: pd.DataFrame, per: Periodos, dim: str, valor: str, n: int,
                  orden_fijo: list[str] | None = None) -> pd.DataFrame:
    """Filas = meses del año anterior y actual (hasta el último con datos) + YTD.

    Devuelve valores absolutos; las filas YTD van al final.
    """
    anio = df[config.COL_ANIO]
    x = df[anio.isin([per.anio_anterior, per.anio_actual])].copy()
    x["mes_lbl"] = (x[config.COL_MES].map(lambda m: config.MESES_CORTOS[m - 1].lower())
                    + "-" + (x[config.COL_ANIO] % 100).astype(str))
    t = x.pivot_table(index="mes_lbl", columns=dim, values=valor, aggfunc="sum")
    ult_mes = int(x.loc[x[config.COL_ANIO] == per.anio_actual, config.COL_MES].max() or 12)
    filas = [f"{config.MESES_CORTOS[m - 1].lower()}-{per.anio_anterior % 100}" for m in range(1, 13)]
    filas += [f"{config.MESES_CORTOS[m - 1].lower()}-{per.anio_actual % 100}"
              for m in range(1, ult_mes + 1)]
    t = t.reindex(filas)
    ytd = por_periodo(df, per, dim, valor).loc[[f"YTD {per.anio_anterior}",
                                                f"YTD {per.anio_actual}"]]
    t = pd.concat([t, ytd])
    t.columns.name = None
    keep = [] if orden_fijo else ranking(t, n, list(ytd.index))
    t = con_otras(t, keep, orden_fijo)
    return t.where(t.abs() > 1e-9)


# --------------------------------------------------------------------------- #
# Mix de medios
# --------------------------------------------------------------------------- #
@dataclass
class BloqueMix:
    etiqueta: str            # "TT 2025" / "YTD 2026"
    shares: pd.DataFrame     # filas = entidades + total, columnas = medios
    totales: pd.Series       # inversión total por fila


def mix_medios(df: pd.DataFrame, per: Periodos, dim: str, n: int,
               orden_fijo: list[str] | None = None) -> list[BloqueMix]:
    x = periodo_de(df, per)
    bloques = []
    for etiqueta in (f"TT {per.anio_anterior}", f"YTD {per.anio_actual}"):
        p = x[x["periodo"] == etiqueta]
        if p.empty:
            continue
        t = p.pivot_table(index=dim, columns="medio_grp", values="inv", aggfunc="sum")
        t = t.reindex(columns=config.ORDEN_MEDIOS).fillna(0)
        t = t[t.sum(axis=1) > 0]
        if orden_fijo is not None:
            t = t.reindex([e for e in orden_fijo if e in t.index])
        else:
            orden = t.sum(axis=1).sort_values(ascending=False)
            keep = [e for e in orden.index if e != config.OTRAS][:n]
            resto = t.drop(index=keep)
            t = t.loc[keep]
            if not resto.empty and resto.values.sum() > 0:
                t.loc[config.OTRAS] = resto.sum()
        t.loc[etiqueta] = t.sum()
        tot = t.sum(axis=1)
        bloques.append(BloqueMix(etiqueta, t.div(tot.where(tot > 0), axis=0), tot))
    return bloques


# --------------------------------------------------------------------------- #
# SOS (TV) vs SOV
# --------------------------------------------------------------------------- #
@dataclass
class SosSov:
    entidades: list[str]
    sos: pd.Series   # participación de inversión en TV
    sov: pd.Series   # participación de TRPs
    trp_total: float
    inv_total: float


def sos_vs_sov(df: pd.DataFrame, per: Periodos, dim: str, target: str, n: int,
               orden_fijo: list[str] | None = None) -> SosSov | None:
    x = df[(df[config.COL_ANIO] == per.anio_actual) & df["es_tv"]]
    g = x.groupby(dim).agg(inv=("inv", "sum"), trp=(target, "sum"))
    if g.empty or g["trp"].sum() <= 0.5:
        return None
    if orden_fijo is not None:
        g = g.reindex([e for e in orden_fijo if e in g.index]).fillna(0)
    else:
        # Se nombran las entidades con TRPs; el resto de la inversión en TV va a OTRAS.
        g = g.sort_values(["trp", "inv"], ascending=False)
        keep = [e for e in g.index if e != config.OTRAS and g.at[e, "trp"] > 0.5][:n]
        resto = g.drop(index=keep)
        g = g.loc[keep]
        if not resto.empty and resto.values.sum() > 0:
            g.loc[config.OTRAS] = resto.sum()
    return SosSov(list(g.index), g["inv"] / g["inv"].sum(), g["trp"] / g["trp"].sum(),
                  float(g["trp"].sum()), float(g["inv"].sum()))
