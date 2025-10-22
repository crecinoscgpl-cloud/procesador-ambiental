import streamlit as st
import pandas as pd
import os
import tempfile
from pathlib import Path
import base64

# Importar nuestros procesadores
from processors.aire_processor import procesar_aire
from processors.ruido_processor import procesar_ruido
from processors.estres_termico_processor import procesar_estres_termico
from utils.excel_generator import generar_excel_consolidado

def main():
    st.set_page_config(
        page_title="Procesador de Datos Ambientales",
        page_icon="🏭",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    st.title("🏭 Procesador de Datos Ambientales")
    st.markdown("---")
    
    # Sidebar con información
    st.sidebar.header("Configuración")
    nombre_empresa = st.sidebar.text_input("Nombre de la Empresa", "Mi Empresa")
    
    st.sidebar.markdown("---")
    st.sidebar.info("""
    **Instrucciones:**
    1. Sube los archivos según el tipo de medición
    2. Configura los tiempos para Airthinx
    3. Procesa los datos
    4. Descarga el reporte consolidado
    """)
    
    # Pestañas para organizar
    tab1, tab2, tab3, tab4 = st.tabs([
        "📁 Carga de Archivos", 
        "⏰ Configuración Tiempos", 
        "📊 Resultados", 
        "📥 Descarga"
    ])
    
    with tab1:
        st.header("Carga de Archivos")
        cargar_archivos()
    
    with tab2:
        st.header("Configuración de Tiempos Airthinx")
        configurar_tiempos_airthinx()
    
    with tab3:
        st.header("Resultados y Gráficas")
        mostrar_resultados()
    
    with tab4:
        st.header("Descarga de Reporte")
        generar_descarga(nombre_empresa)

def cargar_archivos():
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.subheader("🌫️ Calidad de Aire")
        st.info("Archivos 3M (.xls) y Airthinx (.xlsx)")
        
        archivos_3m_aire = st.file_uploader(
            "Subir archivos 3M Aire",
            type=['xls'],
            accept_multiple_files=True,
            key="aire_3m"
        )
        
        archivo_airthinx = st.file_uploader(
            "Subir archivo Airthinx",
            type=['xlsx'],
            key="airthinx"
        )
        
        # Guardar en session state
        if archivos_3m_aire:
            st.session_state.archivos_3m_aire = archivos_3m_aire
            st.success(f"✅ {len(archivos_3m_aire)} archivos 3M cargados")
        
        if archivo_airthinx:
            st.session_state.archivo_airthinx = archivo_airthinx
            st.success("✅ Archivo Airthinx cargado")
    
    with col2:
        st.subheader("🔊 Niveles de Ruido")
        st.info("Archivos 3M Ruido (.xls)")
        
        archivos_ruido = st.file_uploader(
            "Subir archivos Ruido 3M",
            type=['xls'],
            accept_multiple_files=True,
            key="ruido"
        )
        
        if archivos_ruido:
            st.session_state.archivos_ruido = archivos_ruido
            st.success(f"✅ {len(archivos_ruido)} archivos de ruido cargados")
    
    with col3:
        st.subheader("🌡️ Estrés Térmico")
        st.info("Archivo de Estrés Térmico (.xls)")
        
        archivo_et = st.file_uploader(
            "Subir archivo Estrés Térmico",
            type=['xls'],
            key="et"
        )
        
        if archivo_et:
            st.session_state.archivo_et = archivo_et
            st.success("✅ Archivo de estrés térmico cargado")

def configurar_tiempos_airthinx():
    if 'archivos_3m_aire' in st.session_state and st.session_state.archivos_3m_aire:
        num_puntos = len(st.session_state.archivos_3m_aire)
        st.info(f"Se detectaron {num_puntos} puntos de medición. Configura los tiempos para cada punto:")
        
        tiempos_airthinx = {}
        for i in range(1, num_puntos + 1):
            st.markdown(f"**Punto {i}**")
            col1, col2 = st.columns(2)
            with col1:
                inicio = st.time_input(f"Inicio", value=pd.Timestamp("09:00").time(), key=f"inicio_{i}")
            with col2:
                fin = st.time_input(f"Fin", value=pd.Timestamp("17:00").time(), key=f"fin_{i}")
            
            tiempos_airthinx[i] = (inicio, fin)
        
        st.session_state.tiempos_airthinx = tiempos_airthinx
        st.success("✅ Tiempos configurados correctamente")
    else:
        st.warning("⚠️ Primero carga los archivos de aire en la pestaña 'Carga de Archivos'")

def mostrar_resultados():
    if st.button("🚀 Procesar Todos los Datos", type="primary"):
        with st.spinner("Procesando datos... Esto puede tomar unos segundos"):
            procesar_todos_datos()
    
    if 'resultados' in st.session_state:
        resultados = st.session_state.resultados
        
        # Mostrar resumen de datos procesados
        col1, col2, col3 = st.columns(3)
        with col1:
            if resultados['aire']:
                st.metric("Puntos de Aire", len(resultados['aire']))
        with col2:
            if resultados['ruido']:
                st.metric("Puntos de Ruido", len(resultados['ruido']))
        with col3:
            if resultados['estres_termico']:
                st.metric("Parámetros ET", len(resultados['estres_termico']))
        
        # Mostrar gráficas si existen
        if 'graficas' in st.session_state:
            mostrar_graficas_interactivas()

def procesar_todos_datos():
    resultados = {
        'aire': None,
        'ruido': None, 
        'estres_termico': None
    }
    
    try:
        # Procesar aire
        if 'archivos_3m_aire' in st.session_state:
            archivos_3m = st.session_state.archivos_3m_aire
            archivo_airthinx = st.session_state.get('archivo_airthinx', None)
            tiempos_airthinx = st.session_state.get('tiempos_airthinx', {})
            
            resultados['aire'] = procesar_aire(archivos_3m, archivo_airthinx, tiempos_airthinx)
        
        # Procesar ruido
        if 'archivos_ruido' in st.session_state:
            resultados['ruido'] = procesar_ruido(st.session_state.archivos_ruido)
        
        # Procesar estrés térmico
        if 'archivo_et' in st.session_state:
            resultados['estres_termico'] = procesar_estres_termico(st.session_state.archivo_et)
        
        st.session_state.resultados = resultados
        st.success("✅ ¡Datos procesados correctamente!")
        
    except Exception as e:
        st.error(f"❌ Error al procesar datos: {str(e)}")

def mostrar_graficas_interactivas():
    st.subheader("Gráficas Interactivas")
    # Aquí iría el código para mostrar gráficas con Plotly
    # (lo implementaremos después)

def generar_descarga(nombre_empresa):
    if 'resultados' in st.session_state and st.session_state.resultados:
        if st.button("📊 Generar Excel Consolidado", type="primary"):
            with st.spinner("Generando reporte Excel..."):
                try:
                    excel_buffer = generar_excel_consolidado(
                        st.session_state.resultados['aire'],
                        st.session_state.resultados['ruido'], 
                        st.session_state.resultados['estres_termico'],
                        nombre_empresa
                    )
                    
                    # Crear botón de descarga
                    st.download_button(
                        label="📥 Descargar Reporte Completo",
                        data=excel_buffer,
                        file_name=f"Reporte_Ambiental_{nombre_empresa}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    st.success("✅ Reporte generado correctamente")
                    
                except Exception as e:
                    st.error(f"❌ Error al generar reporte: {str(e)}")
    else:
        st.warning("⚠️ Primero procesa los datos en la pestaña 'Resultados'")

if __name__ == "__main__":
    main()