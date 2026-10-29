"""Orquestación: arma el informe completo según la estructura de la hoja "Informe".

Estructura
----------
1. Compañía:   inversión y SOS, inversión mensual, mix de medios | TRP y SOV,
               TRP mensual | SOS vs SOV
2. Categoría:  inversión y SOS | TRP y SOV
3-6. Masivos, Botellas finas, Espumantes, Coolers y sangría (por marca):
               inversión y SOS, inversión mensual, SOS mensual, mix de medios |
               TRP y SOV, TRP mensual | SOS vs SOV
Adicional: BLYS vinos y BLYS total (por marca):
               inversión y SOS, inversión mensual | TRP y SOV, TRP mensual |
               SOS vs SOV
"""
from __future__ import annotations

import io
from dataclasses import dataclass, field
from typing import Callable

import pandas as pd

from . import config, diapositivas as D
from .agrupaciones import Agrupaciones
from .datos import Diagnostico, preparar
from .pptx_utils import Colores, Deck
from .trazabilidad import Entrada, Trazabilidad, escribir_excel, parametros_corrida

TODAS_MARCA = ("inv", "inv_mensual", "sos_mensual", "mix", "trp", "trp_mensual", "sos_sov")
BLYS = ("inv", "inv_mensual", "trp", "trp_mensual", "sos_sov")

TITULO_UNIVERSO = "MASIVOS + BOTELLAS FINAS + ESPUMANTES + COOLERS Y SANGRÍA"
UNIVERSO_VINO = ("es_vino = VERDADERO (áreas Vino masivo, Vino premium, Espumantes y "
                 "Vino-cooler-sangría)")

# Target por defecto de cada sección (en el informe de referencia, solo vino
# premium usa HM 35-49 Alto).
TARGETS_DEFECTO = {
    "compania": "HM25_49MB_Pond",
    "categoria": "HM25_49MB_Pond",
    "MASIVOS": "HM25_49MB_Pond",
    "BOTELLAS FINAS": "HM35_49A_Pond",
    "ESPUMANTES": "HM25_49MB_Pond",
    "COOLERS Y SANGRÍA": "HM25_49MB_Pond",
    "BLYS VINOS": "HM25_49MB_Pond",
    "BLYS TOTAL": "HM25_49MB_Pond",
}

NOMBRES_SECCION = {
    "compania": "1. Compañía",
    "categoria": "2. Categoría",
    "MASIVOS": "3. Masivos",
    "BOTELLAS FINAS": "4. Botellas finas",
    "ESPUMANTES": "5. Espumantes",
    "COOLERS Y SANGRÍA": "6. Coolers y sangría",
    "BLYS VINOS": "Adicional · BLYS vinos",
    "BLYS TOTAL": "Adicional · BLYS total",
}


def secciones(ag: Agrupaciones, targets: dict[str, str] | None = None) -> list[D.Seccion]:
    t = {**TARGETS_DEFECTO, **(targets or {})}
    companias = ag.companias + [config.COMPANIA_OTRAS]
    out = [
        D.Seccion("compania", "COMPAÑÍA", "COMPAÑÍA", "compania",
                  lambda d: d["es_vino"], ("inv", "inv_mensual", "mix", "trp",
                                           "trp_mensual", "sos_sov"),
                  t["compania"], companias,
                  separador=f"POR COMPAÑÍA:\n{TITULO_UNIVERSO}", universo=UNIVERSO_VINO),
        D.Seccion("categoria", "CATEGORÍA", "CATEGORÍA", "categoria",
                  lambda d: d["es_vino"], ("inv", "trp"), t["categoria"],
                  list(config.ORDEN_CATEGORIAS),
                  separador=f"POR CATEGORÍA:\n{TITULO_UNIVERSO}", universo=UNIVERSO_VINO),
    ]
    for cat in config.ORDEN_CATEGORIAS:
        out.append(D.Seccion(cat, cat, "MARCA", "marca",
                             (lambda c: lambda d: d["es_vino"] & (d["categoria"] == c))(cat),
                             TODAS_MARCA, t[cat],
                             universo=f"es_vino = VERDADERO y categoria = {cat}"))
    out += [
        D.Seccion("BLYS VINOS", "BLYS VINOS", "MARCA", "marca_blys",
                  lambda d: d["blys_vinos"], BLYS, t["BLYS VINOS"],
                  universo="blys_vinos = VERDADERO"),
        D.Seccion("BLYS TOTAL", "BLYS TOTAL", "MARCA", "marca_blys",
                  lambda d: d["blys_total"], BLYS, t["BLYS TOTAL"],
                  universo="blys_total = VERDADERO"),
    ]
    return out


def _registrar_simple(ctx: D.Contexto, deck: Deck, seccion: str, tipo: str, titulo: str):
    if ctx.traza is not None:
        ctx.traza.registrar(Entrada(n=len(deck.prs.slides), seccion=seccion, tipo=tipo,
                                    titulo=titulo.replace("\n", " ")))


def _construir_seccion(deck: Deck, ctx: D.Contexto, sec: D.Seccion) -> None:
    df = ctx.df[sec.filtro(ctx.df)]
    D.separador(deck, sec.separador or sec.nombre)
    _registrar_simple(ctx, deck, sec.nombre, "Separador", sec.separador or sec.nombre)
    pasos = sec.diapositivas
    if "inv" in pasos:
        D.por_periodo(deck, ctx, sec, df, "inv")
    if "inv_mensual" in pasos:
        D.tabla_mensual(deck, ctx, sec, df, "inv")
    if "sos_mensual" in pasos:
        D.sos_mensual(deck, ctx, sec, df)
    if "mix" in pasos:
        D.mix_medios(deck, ctx, sec, df)
    if {"trp", "trp_mensual", "sos_sov"} & set(pasos):
        if df.loc[df["es_tv"], sec.target].sum() <= 0.5:
            ctx.avisos.append(f"{sec.nombre}: sin TRPs ({sec.target}) en todo el periodo; "
                              "se omite el bloque de televisión.")
            return
        D.subseparador(deck, f"{sec.nombre} – TELEVISIÓN")
        _registrar_simple(ctx, deck, sec.nombre, "Separador", f"{sec.nombre} – TELEVISIÓN")
    if "trp" in pasos:
        D.por_periodo(deck, ctx, sec, df, "trp")
    if "trp_mensual" in pasos:
        D.tabla_mensual(deck, ctx, sec, df, "trp")
    if "sos_sov" in pasos:
        D.sos_vs_sov(deck, ctx, sec, df)


@dataclass
class Resultado:
    pptx: bytes
    diagnostico: Diagnostico
    avisos: list[str] = field(default_factory=list)
    xlsx: bytes | None = None  # planilla de trazabilidad


def generar(df_raw: pd.DataFrame, ag: Agrupaciones, op: D.Opciones, *,
            incluir: list[str] | None = None, targets: dict[str, str] | None = None,
            fusiones: dict[str, str] | None = None,
            trazabilidad: bool = True, incluir_base: bool = True,
            origen: dict[str, str] | None = None,
            progreso: Callable[[float, str], None] | None = None) -> Resultado:
    """Genera el PPTX y, opcionalmente, la planilla de trazabilidad.

    ``origen``: datos de la corrida a mostrar en la planilla (p. ej. nombres de
    archivos). ``incluir_base`` agrega la base clasificada a la planilla.
    """
    fusiones = config.FUSIONES_DEFECTO if fusiones is None else fusiones
    targets_ef = {**TARGETS_DEFECTO, **(targets or {})}
    df, per, diag = preparar(df_raw, ag, op.fecha_corte, op.anio_inicio, fusiones)
    colores = Colores(config.cargar_colores(), config.paleta_respaldo())
    ctx = D.Contexto(df, per, colores, op, traza=Trazabilidad() if trazabilidad else None)
    deck = Deck()
    D.portada(deck, ctx)
    _registrar_simple(ctx, deck, "PORTADA", "Portada", op.titulo)
    secs = [s for s in secciones(ag, targets) if incluir is None or s.clave in incluir]
    total = len(secs) + (1 if trazabilidad else 0)
    for i, sec in enumerate(secs):
        if progreso:
            progreso(i / total, f"Construyendo {NOMBRES_SECCION[sec.clave]}…")
        _construir_seccion(deck, ctx, sec)
    buf = io.BytesIO()
    deck.guardar(buf)
    res = Resultado(buf.getvalue(), diag, ctx.avisos)
    if trazabilidad:
        if progreso:
            progreso(len(secs) / total, "Escribiendo planilla de trazabilidad…")
        ctx.traza.omitidas = list(ctx.avisos)
        res.xlsx = escribir_excel(
            ctx.traza,
            parametros=parametros_corrida(op, per, targets_ef, fusiones, incluir,
                                          NOMBRES_SECCION, origen),
            clasificacion=diag.clasificacion,
            diagnostico={
                "Marcas de vino sin categoría en la hoja (se usa la de su área)":
                    diag.sin_categoria,
                "Marcas no listadas en su compañía (asignadas por empresa)":
                    diag.compania_por_respaldo,
                "Marcas BLYS no listadas (incluidas por su área)": diag.blys_no_listadas,
            },
            base=df if incluir_base else None)
    if progreso:
        progreso(1.0, "Listo")
    return res
