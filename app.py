import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from io import BytesIO
import matplotlib.ticker as ticker
import numpy as np

st.set_page_config(page_title="Generador de Informe VSPT", layout="wide")
st.title("🍷 Generador Automático: Informe de Competencia")
st.markdown("Sube los archivos para procesar los datos. El sistema limpiará las diapositivas del template original y generará gráficos reales basados en tus datos.")

col1, col2, col3 = st.columns(3)
with col1:
    file_datos = st.file_uploader("1. Informe VSPT (Excel datos)", type=['xlsx', 'xls'])
with col2:
    file_cat = st.file_uploader("2. Categorías Vinos (Agrupaciones)", type=['xlsx', 'xls'])
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
    """Elimina todas las diapositivas del template para dejarlo en blanco, manteniendo el diseño (Slide Master)"""
    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    for slide in slides:
        xml_slides.remove(slide)

def add_slide_with_plot(prs, fig, slide_title):
    # Usar layout en blanco (usualmente el 6)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # Agregar título de la diapositiva
    txBox = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(1))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = slide_title
    p.font.size = Pt(24)
    p.font.bold = True
    
    # Guardar gráfico e insertarlo
    img_stream = BytesIO()
    fig.savefig(img_stream, format='png', bbox_inches='tight', dpi=120)
    img_stream.seek(0)
    slide.shapes.add_picture(img_stream, Inches(0.5), Inches(1.2), width=Inches(9))
    plt.close(fig)

if st.button("🚀 Procesar Datos y Generar Presentación", type="primary"):
    if file_datos and file_cat and file_ppt:
        with st.spinner("Procesando datos y cruzando marcas..."):
            mapa_marcas = extraer_agrupaciones(file_cat)
            
            df_datos = pd.read_excel(file_datos, sheet_name='datos')
            df_datos['Marca_Clean'] = df_datos['Marcas reunidas'].astype(str).str.strip().upper()
            df_datos['Compañía_Agrupada'] = df_datos['Marca_Clean'].map(mapa_marcas).fillna('OTRAS VIÑAS')
            df_datos['Año'] = pd.to_numeric(df_datos['Año'], errors='coerce')
            df_datos['ValorPublico'] = pd.to_numeric(df_datos['ValorPublico'], errors='coerce').fillna(0)
            
            # Formateador de moneda para gráficos
            currency_formatter = ticker.FuncFormatter(lambda x, pos: f'${x:,.0f}')

        with st.spinner("Construyendo las diapositivas (limpiando template previo)..."):
            prs = Presentation(file_ppt)
            borrar_diapositivas_existentes(prs)
            
            # --- SECCIÓN 1: COMPAÑÍA GLOBAL ---
            # 1. Inversión Total por Compañía (Torta)
            inv_cia = df_datos.groupby('Compañía_Agrupada')['ValorPublico'].sum().sort_values(ascending=False)
            fig, ax = plt.subplots(figsize=(10, 5))
            ax.pie(inv_cia, labels=inv_cia.index, autopct='%1.1f%%', startangle=90, cmap='Set2')
            ax.set_title("SOS Inversión por Compañía (Total Histórico)")
            add_slide_with_plot(prs, fig, "Inversión y SOS por Compañía")

            # 2. Inversión Mensual por Compañía
            df_reciente = df_datos[df_datos['Año'].isin([2025, 2026])]
            if not df_reciente.empty:
                pivot_mensual = df_reciente.pivot_table(index='Mes', columns=['Compañía_Agrupada', 'Año'], values='ValorPublico', aggfunc='sum').fillna(0)
                fig, ax = plt.subplots(figsize=(12, 5))
                pivot_mensual.plot(kind='line', marker='o', ax=ax)
                ax.set_title("Inversión Mensual por Compañía (2025-2026)")
                ax.yaxis.set_major_formatter(currency_formatter)
                ax.legend(bbox_to_anchor=(1.01, 1), loc='upper left', fontsize=8)
                plt.tight_layout()
                add_slide_with_plot(prs, fig, "Inversión Mensual 2025 - 2026")

            # 3. Mix de Medios por Compañía
            mix_medios = df_datos.pivot_table(index='Compañía_Agrupada', columns='Medio', values='ValorPublico', aggfunc='sum').fillna(0)
            mix_medios_pct = mix_medios.div(mix_medios.sum(axis=1), axis=0) * 100
            fig, ax = plt.subplots(figsize=(12, 5))
            mix_medios_pct.plot(kind='bar', stacked=True, ax=ax, cmap='tab20')
            ax.set_title("Mix de Medios por Compañía (%)")
            ax.legend(bbox_to_anchor=(1.01, 1), loc='upper left', fontsize=8)
            plt.xticks(rotation=45)
            plt.tight_layout()
            add_slide_with_plot(prs, fig, "Mix de Medios por Compañía")

            # --- SECCIÓN 2: POR CATEGORÍA Y MARCAS ---
            categorias = df_datos['Área'].dropna().unique()
            
            for cat in categorias:
                df_cat = df_datos[df_datos['Área'] == cat]
                
                # Top Marcas Inversión
                inv_marcas = df_cat.groupby('Marca_Clean')['ValorPublico'].sum().sort_values(ascending=False).head(10) # Top 10 para que sea legible
                fig, ax = plt.subplots(figsize=(10, 5))
                inv_marcas.plot(kind='bar', color='#8E244D', ax=ax)
                ax.set_title(f"Top 10 Marcas en {cat} (Inversión)")
                ax.yaxis.set_major_formatter(currency_formatter)
                plt.xticks(rotation=45)
                plt.tight_layout()
                add_slide_with_plot(prs, fig, f"{cat}: Inversión y SOS por Marcas")

                # Mix Medios por Marca
                mix_marca = df_cat.pivot_table(index='Marca_Clean', columns='Medio', values='ValorPublico', aggfunc='sum').fillna(0)
                # Filtrar solo el top 10 de marcas para el gráfico apilado
                mix_marca = mix_marca.loc[inv_marcas.index] 
                mix_marca_pct = mix_marca.div(mix_marca.sum(axis=1), axis=0) * 100
                
                fig, ax = plt.subplots(figsize=(12, 5))
                mix_marca_pct.plot(kind='bar', stacked=True, ax=ax, cmap='tab20')
                ax.set_title(f"Mix de Medios por Marca - {cat} (%)")
                ax.legend(bbox_to_anchor=(1.01, 1), loc='upper left', fontsize=8)
                plt.xticks(rotation=45)
                plt.tight_layout()
                add_slide_with_plot(prs, fig, f"{cat}: Mix Medios por marca")

            # Guardar el resultado
            output_pptx = BytesIO()
            prs.save(output_pptx)
            output_pptx.seek(0)
            
        st.success("¡Informe generado! El template fue limpiado y rellenado con gráficos generados a partir de los datos.")
        
        st.download_button(
            label="📥 Descargar PPTX Limpio y Actualizado",
            data=output_pptx,
            file_name="Informe_VSPT_Automatizado.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
    else:
        st.warning("Sube los 3 archivos para comenzar.")
