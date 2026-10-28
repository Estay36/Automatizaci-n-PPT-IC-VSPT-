# Informe de Competencia – VSPT

Aplicación Streamlit que genera el PPT de competencia de VSPT a partir del export de
Megatime (hoja `datos`) y de la planilla **Categorías Vinos** (hoja `Agrupaciones`).
El resultado usa la plantilla corporativa del informe de referencia y gráficos nativos
de PowerPoint (editables, con su planilla de datos).

## Uso

1. Cargar la base de Megatime (`Informe_VSPT_2022-02oct2026.xlsx` o similar).
2. (Opcional) Cargar una versión actualizada de *Categorías Vinos*. Si no se carga, se usa
   `assets/Categorias_Vinos.xlsx`.
3. Revisar la fecha de corte (se lee del nombre del archivo, p. ej. `02oct2026`), las
   secciones, el número de marcas por gráfico y el target de TRP de cada sección.
4. Revisar el **Diagnóstico de agrupaciones**: muestra las marcas que no calzan con la hoja.
5. *Generar PPT* → *Descargar PPT*.

## Estructura del informe (hoja «Informe»)

| Sección | I. Inversión | II. TRP (TV) | III. |
|---|---|---|---|
| 1. Compañía | Inversión y SOS · Inversión mensual (tabla) · Mix de medios | TRP y SOV · TRP mensual (tabla) | SOS vs SOV |
| 2. Categoría | Inversión y SOS | TRP y SOV | — |
| 3–6. Masivos, Botellas finas, Espumantes, Coolers y sangría | Inversión y SOS por marca · Inversión mensual · SOS mensual · Mix de medios | TRP y SOV · TRP mensual | SOS vs SOV |
| Adicional: BLYS vinos y BLYS total | Inversión y SOS por marca · Inversión mensual | TRP y SOV · TRP mensual | SOS vs SOV |

Periodos: `TT` = año calendario completo; `YTD` = año a la fecha. El YTD del año anterior se
toma de la columna `Periodo = COMPARABLE` (si no existe, se usan las mismas semanas del año
actual). La variación % de las tablas mensuales compara YTD contra YTD comparable, y el mes en
curso contra el mismo tramo del año anterior.

Si una sección no tiene TRPs, se omite su bloque de televisión y la app lo informa.

## Reglas de clasificación

**Compañía** (universo: Vino masivo, Vino premium, Espumantes y Vino-cooler-sangría)
1. Marca listada bajo VSPT, Concha y Toro, Santa Rita o Ravanal en la hoja → esa compañía.
2. Si no, según la empresa (`config.EMPRESA_A_COMPANIA`).
3. Si no, **OTRAS VIÑAS**.

**Categoría**
1. Listas de la hoja que contienen la marca (las celdas «OTRAS BOTELLAS FINAS\*» y
   «OTROS ESPUMANTES\*\*» remiten a sus columnas).
2. Si la marca está en varias listas, desempata el área (p. ej. Undurraga en *Vino masivo* →
   Masivos; en *Espumantes* → Espumantes).
3. Si está en una sola lista, se usa esa aunque el área diga otra cosa (Donnaluna → Botellas finas).
4. Si no está en ninguna, se usa la categoría natural de su área y aparece en el diagnóstico.

**Fusiones de marca** (editables en la app): `120 SPRITZ → 120` y
`EXPORTACIÓN SELECTO ICE → EXPORTACIÓN`, según el comentario de la hoja.

**BLYS**: marcas listadas en BLYS VINOS / BLYS TOTAL. `GATO` en área Vino-cooler-sangría se
trata como **GATO LATA**. Las marcas bajo «Tentativo incluir a futuro» quedan fuera. Las marcas
no mencionadas en la hoja pero cuya área es BLYS (Vino-cooler-sangría, Cervezas, ICE, BLYS) se
incluyen y se informan en el diagnóstico.

## Despliegue desde GitHub (Streamlit Community Cloud)

1. Subir este directorio a un repositorio **privado** de GitHub (la plantilla contiene logos
   corporativos; las bases de datos no se versionan, ver `.gitignore`).
2. En <https://share.streamlit.io> → *Create app* → elegir el repositorio, rama `main` y
   archivo principal `app.py`.
3. La app queda disponible en una URL propia; el acceso se controla desde la configuración de
   la app en Streamlit Cloud.

Ejecución local:

```bash
python -m venv .venv && source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

## Estructura del código

```
app.py                     Interfaz Streamlit
informe/config.py          Columnas, mapeos (medios, áreas, empresas), textos y rutas
informe/agrupaciones.py    Lectura de la hoja «Agrupaciones»
informe/datos.py           Carga de la base, clasificación y periodos TT/YTD
informe/metricas.py        Agregaciones (por periodo, mensual, mix, SOS vs SOV, top-N)
informe/pptx_utils.py      Plantilla, gráficos nativos y tablas (python-pptx)
informe/diapositivas.py    Un constructor por tipo de diapositiva
informe/reporte.py         Secciones del informe y orquestación
assets/plantilla.pptx      Diseños corporativos (sin diapositivas)
assets/colores.json        Color fijo por marca, compañía, categoría y medio
assets/Categorias_Vinos.xlsx  Agrupaciones por defecto
tools/crear_plantilla.py   Regenera la plantilla desde un informe de referencia
tests/                     Pruebas con datos sintéticos (python -m pytest -q)
```

### Ajustes frecuentes

- **Colores**: editar `assets/colores.json`. Las marcas sin color usan la paleta de respaldo y
  mantienen el mismo color en todo el informe.
- **Nuevas marcas**: agregarlas a la hoja *Agrupaciones* y volver a cargarla en la app.
- **Cambio de formato corporativo**: `python tools/crear_plantilla.py nuevo_informe.pptx assets/plantilla.pptx`.
- **Top-N**: las marcas de cada gráfico se eligen por su SOS ponderado entre periodos (YTD
  actual y TT del año anterior pesan el doble); el resto se agrupa en OTRAS.
