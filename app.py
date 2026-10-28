import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches
from io import BytesIO
import matplotlib.ticker as ticker
import numpy as np

st.set_page_config(page_title="Generador de Informe VSPT", layout="wide")
st.title("🍷 Generador Automático: Informe de Competencia VSPT completo")

col1, col2, col3 = st.columns(3)
with col1:
    file_datos = st.file_uploader("1. Informe VSPT (Excel datos)", type=['xlsx', 'xls'])
with col2:
    file_cat = st.file_uploader("2. Categorías Vinos", type=['xlsx', 'xls'])
with col3:
    file_ppt = st.file_uploader("3. Template Base (PPTX)", type=['pptx'])

def extraer_agrupaciones(archivo_categorias):
    df_agr = pd.read_excel(archivo_categorias, sheet_name='Agrupaciones', header=None)
    mapa_marcas = {}
    for col in df_agr.columns:
        company = df_agr.iloc[2, col]
        if pd.notna(company):
            marcas = df_agr.iloc[3:, col].dropna().tolist()
            for marca in marcas:
                marca_limpia = str(marca).strip().upper()
                mapa_marcas[marca_limpia] = str(company).strip().upper()
    return mapa_marcas

def borrar_diapositivas_existentes(prs):
    # Elimina las diapositivas de ejemplo del template para dejar solo el layout
    for i in range(len(prs.slides) - 1, -1, -1):
        rId = prs.slides._sldIdLst[i].rId
        prs.part.drop_rel(rId)
        del prs.slides._sldIdLst[i]

def add_slide_with_plot(prs, fig, slide_title):
    try:
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        slide.shapes.title.text = slide_title
    except:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        txBox = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(1))
        txBox.text_frame.text = slide_title
        
    img_stream = BytesIO()
    fig.savefig(img_stream, format='png', bbox_inches='tight', dpi=150)
    img_stream.seek(0)
    slide.shapes.add_picture(img_stream, Inches(1), Inches(1.5), width=Inches(8))
    plt.close(fig)

def generar_set_graficos(df, prs, nivel_agrupacion, titulo_seccion):
    """
    Genera el set estándar de gráficos (Inversión, SOS, Mensual, Medios, TRP) 
    para un dataframe filtrado, cumpliendo los ítems de la hoja 'Informe'.
    """
    # 1. Inversión y SOS
    if 'ValorPublico' in df.columns:
        inv = df.groupby(nivel_agrupacion)['ValorPublico'].sum().sort_values(ascending=False)
        if inv.sum() > 0:
            fig1, ax1 = plt.subplots(figsize=(10, 5))
            ax1.pie(inv, labels=inv.index, autopct='%1.1f%%', startangle=90, colors=plt.cm.tab20.colors)
            ax1.set_title(f"{titulo_seccion} - Inversión y SOS", fontsize=14, pad=20)
            add_slide_with_plot(prs, fig1, f"{titulo_seccion}: Inversión y SOS")

    # 2. Inversión Mensual 2025 vs 2026
    df_reciente = df[df['Año'].isin([2025, 2026])]
    if not df_reciente.empty and 'Mes' in df.columns:
        pivot_mensual = df_reciente.pivot_table(index='Mes', columns=[nivel_agrupacion, 'Año'], values='ValorPublico', aggfunc='sum').fillna(0)
        if not pivot_mensual.empty:
            fig2, ax2 = plt.subplots(figsize=(12, 6))
            pivot_mensual.plot(kind='line', marker='o', ax=ax2)
            ax2.set_title(f"{titulo_seccion} - Inversión Mensual (25-26)", fontsize=14)
            ax2.yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, pos: f'${x:,.0f}'))
            ax2.grid(True, linestyle='--', alpha=0.6)
            plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left', fontsize=8)
            add_slide_with_plot(prs, fig2, f"{titulo_seccion}: Inversión Mensual")

    # 3. Mix de Medios
    if 'Medio' in df.columns:
        mix = df.pivot_table(index=nivel_agrupacion, columns='Medio', values='ValorPublico', aggfunc='sum').fillna(0)
        if not mix.empty and mix.sum().sum() > 0:
            mix_pct = mix.div(mix.sum(axis=1), axis=0) * 100
            fig3, ax3 = plt.subplots(figsize=(12, 6))
            mix_pct.plot(kind='bar', stacked=True, ax=ax3, colormap='Spectral')
            ax3.set_title(f"{titulo_seccion} - Mix de Medios", fontsize=14)
            plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left', fontsize=8)
            plt.xticks(rotation=45, ha='right')
            add_slide_with_plot(prs, fig3, f"{titulo_seccion}: Mix de Medios")

    # 4. TRP y SOV (Usa la columna HM25_49MB_Pond como referencia principal de TRP)
    trp_col = 'HM25_49MB_Pond' if 'HM25_49MB_Pond' in df.columns else None
    if trp_col and trp_col in df.columns:
        df[trp_col] = pd.to_numeric(df[trp_col], errors='coerce').fillna(0)
        trp = df.groupby(nivel_agrupacion)[trp_col].sum().sort_values(ascending=False)
        trp = trp[trp > 0]
        if not trp.empty:
            fig4, ax4 = plt.subplots(figsize=(10, 5))
            trp.plot(kind='bar', ax=ax4, color='darkblue')
            ax4.set_title(f"{titulo_seccion} - TRP y SOV", fontsize=14)
            plt.xticks(rotation=45, ha='right')
            add_slide_with_plot(prs, fig4, f"{titulo_seccion}: TRP y SOV")

if st.button("🚀 Procesar Datos y Generar Presentación", type="primary"):
    if file_datos and file_cat and file_ppt:
        with st.spinner("1. Extrayendo agrupaciones y corrigiendo datos..."):
            mapa_marcas = extraer_agrupaciones(file_cat)
            df_datos = pd.read_excel(file_datos, sheet_name='datos')
            
            # CORRECCIÓN DEL ERROR DE SINTAXIS (agregado .str. antes de upper())
            df_datos['Marca_Clean'] = df_datos['Marcas reunidas'].astype(str).str.strip().str.upper()
            df_datos['Compañía_Agrupada'] = df_datos['Marca_Clean'].map(mapa_marcas).fillna('OTRAS VIÑAS')
            
            df_datos['Año'] = pd.to_numeric(df_datos['Año'], errors='coerce')
            df_datos['ValorPublico'] = pd.to_numeric(df_datos['ValorPublico'], errors='coerce').fillna(0)
            df_datos['Área'] = df_datos['Área'].astype(str).str.strip().str.upper()

        with st.spinner("2. Iterando sobre los 40 ítems estructurales e inyectando gráficos..."):
            prs = Presentation(file_ppt)
            borrar_diapositivas_existentes(prs) # Limpia el template
            
            # --- SECCIÓN 1: COMPAÑÍA ---
            generar_set_graficos(df_datos, prs, 'Compañía_Agrupada', "1. GLOBAL COMPAÑÍAS")
            
            # --- SECCIÓN 2: CATEGORÍA ---
            if 'Área' in df_datos.columns:
                generar_set_graficos(df_datos, prs, 'Área', "2. GLOBAL CATEGORÍAS")
            
            # --- SECCIONES ESPECÍFICAS (Mapeando columnas 7, 11, 15, 19 del informe) ---
            categorias_objetivo = {
                "3. VINO MASIVO": "VINO MASIVO",
                "4. VINO PREMIUM": "VINO PREMIUM",
                "5. ESPUMANTE": "ESPUMANTE",
                "6. COOLER Y SANGRÍA": "VINO_COOLER_SANGRÍA"
            }
            
            for titulo_seccion, area_filtro in categorias_objetivo.items():
                # Filtramos la data por el área específica (ej: VINO MASIVO)
                df_area = df_datos[df_datos['Área'].str.contains(area_filtro, na=False)]
                if not df_area.empty:
                    # Al entrar a un área, agrupamos por Marca_Clean o Compañía según detalle
                    generar_set_graficos(df_area, prs, 'Marca_Clean', titulo_seccion)

            output_pptx = BytesIO()
            prs.save(output_pptx)
            output_pptx.seek(0)
            
        st.success("¡Informe generado! Se han creado los cruces de Inversión, Mix de Medios y TRP por Compañía, Categoría y Marcas específicas.")
        st.download_button(
            label="📥 Descargar PPTX Completo",
            data=output_pptx,
            file_name="Informe_Competencia_VSPT_Completo.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
    else:
        st.warning("⚠️ Sube los 3 archivos para comenzar.")
