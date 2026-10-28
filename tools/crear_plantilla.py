"""Genera ``assets/plantilla.pptx`` a partir de un informe de competencia existente.

Uso
---
    python tools/crear_plantilla.py "Informe_de_Competencia_-_VSPT.pptx" assets/plantilla.pptx

El script:
1. Elimina todas las diapositivas del informe (y con ello sus gráficos y planillas
   incrustadas, que son la mayor parte del peso del archivo).
2. Conserva solo los diseños (layouts) que usa el generador y los renombra con
   nombres estables, para que la aplicación los encuentre por nombre.
3. Elimina los patrones (masters) que quedan sin diseños.

Si el formato corporativo cambia, basta con volver a correr este script sobre
el nuevo informe de referencia. Los diseños se identifican por el nombre de su
archivo interno (``slideLayoutN.xml``); si cambian, ajuste ``DISENOS``.
"""
from __future__ import annotations

import sys
from pathlib import Path

from pptx import Presentation

# Archivo interno del diseño en el informe de referencia -> nombre estable.
DISENOS = {
    "slideLayout36.xml": "VSPT Portada",
    "slideLayout28.xml": "VSPT Separador",
    "slideLayout29.xml": "VSPT Subseparador",
    "slideLayout18.xml": "VSPT Contenido",
    "slideLayout21.xml": "VSPT Contenido lateral",
    "slideLayout22.xml": "VSPT Tabla",
}


def _eliminar_diapositivas(prs) -> None:
    sld_id_lst = prs.slides._sldIdLst  # noqa: SLF001 (API interna de python-pptx)
    for sld_id in list(sld_id_lst):
        prs.part.drop_rel(sld_id.rId)
        sld_id_lst.remove(sld_id)


def _depurar_disenos(prs) -> None:
    for master in prs.slide_masters:
        for layout in list(master.slide_layouts):
            archivo = Path(str(layout.part.partname)).name
            if archivo in DISENOS:
                layout.name = DISENOS[archivo]
            else:
                master.slide_layouts.remove(layout)


def _eliminar_patrones_vacios(prs) -> None:
    lst = prs.slide_masters._sldMasterIdLst  # noqa: SLF001
    for sld_master_id in list(lst):
        master = prs.part.related_part(sld_master_id.rId).slide_master
        if len(master.slide_layouts) == 0:
            prs.part.drop_rel(sld_master_id.rId)
            lst.remove(sld_master_id)


def _eliminar_historial_revisiones(prs) -> None:
    """El historial de co-autoría referencia diapositivas que ya no existen."""
    for r_id, rel in list(prs.part.rels.items()):
        if "changesInfo" in rel.reltype or "revisionInfo" in rel.reltype:
            prs.part.drop_rel(r_id)


def crear_plantilla(origen: str | Path, destino: str | Path) -> Path:
    prs = Presentation(str(origen))
    _eliminar_historial_revisiones(prs)
    _eliminar_diapositivas(prs)
    _depurar_disenos(prs)
    _eliminar_patrones_vacios(prs)
    faltantes = set(DISENOS.values()) - {
        l.name for m in prs.slide_masters for l in m.slide_layouts
    }
    if faltantes:
        raise RuntimeError(f"No se encontraron los diseños: {sorted(faltantes)}")
    destino = Path(destino)
    destino.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(destino))
    return destino


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    ruta = crear_plantilla(sys.argv[1], sys.argv[2])
    print(f"Plantilla creada: {ruta} ({ruta.stat().st_size / 1e6:.1f} MB)")
