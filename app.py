import streamlit as st
import pandas as pd
import io
import base64
from datetime import datetime, time

# ===== CONFIGURACIÓN =====
st.set_page_config(
    page_title="Procesador de Datos Ambientales",
    page_icon="🏭",
    layout="wide"
)

# ===== FUNCIONES DE PROCESAMIENTO =====
def safe_min(series):
    try:
        return float(series.min())
    except:
        return 0.0

def safe_max(series):
    try:
        return float(series.max())
    except:
        return 0.0

def safe_mean(series):
    try:
        return float(series.mean())
    except:
        return 0.0

def procesar_3m_aire(archivos):
    """Procesa archivos 3M de calidad de aire"""
    resultados = {}
    
    for i, archivo in enumerate(archivos, 1):
        try:
            # Leer archivo .xls
            df = pd.read_excel(archivo, header=2, engine='xlrd')
            df.columns = [str(col).strip() for col in df.columns]
            
            if len(df.columns) >= 6:
                resultados[i] = {
                    'CO (ppm)': [safe_min(df.iloc[:, 1]), safe_mean(df.iloc[:, 1]), safe_max(df.iloc[:, 1])],
                    'Polvo (µg/m3)': [safe_min(df.iloc[:, 2]), safe_mean(df.iloc[:, 2]), safe_max(df.iloc[:, 2])],
                    'Humedad Relativa (%)': [safe_min(df.iloc[:, 3]), safe_mean(df.iloc[:, 3]), safe_max(df.iloc[:, 3])],
                    'Temperatura (°C)': [safe_min(df.iloc[:, 4]), safe_mean(df.iloc[:, 4]), safe_max(df.iloc[:, 4])]
                }
        except Exception as e:
            st.error(f"Error procesando {archivo.name}: {str(e)}")
    
    return resultados

def procesar_ruido_3m(archivos):
    """Procesa archivos 3M de ruido"""
    resultados = {}
    
    for i, archivo in enumerate(archivos, 1):
        try:
            df = pd.read_excel(archivo, header=2, engine='xlrd')
            df.columns = [str(col).strip() for col in df.columns]
            
            if len(df.columns) >= 3:
                resultados[i] = {
                    'Lapk-1': [safe_min(df.iloc[:, 1]), safe_mean(df.iloc[:, 1]), safe_max(df.iloc[:, 1])],
                    'Leq-1': [safe_min(df.iloc[:, 2]), safe_mean(df.iloc[:, 2]), safe_max(df.iloc[:, 2])]
                }
        except Exception as e:
            st.error(f"Error procesando ruido {archivo.name}: {str(e)}")
    
    return resultados

def generar_excel_consolidado(datos_aire, datos_ruido, nombre_empresa):
    """Genera Excel usando solo pandas (sin openpyxl)"""
    output = io.BytesIO()
    
    # Crear Excel con pandas
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Hoja Resumen Aire
        if datos_aire:
            df_aire = crear_resumen_aire(datos_aire)
            df_aire.to_excel(writer, sheet_name='Resumen Aire', index=False)
        
        # Hoja Resumen Ruido
        if datos_ruido:
            df_leq, df_lcpk = crear_resumen_ruido(datos_ruido)
            df_leq.to_excel(writer, sheet_name='Resumen Ruido Leq', index=False)
            df_lcpk.to_excel(writer, sheet_name='Resumen Ruido Lcpk', index=False)
    
    output.seek(0)
    return output.getvalue()

def crear_resumen_aire(datos_aire):
    """Crea DataFrame para resumen de aire"""
    filas = []
    
    # Agregar fila de límites
    filas.append({
        'Punto': 'Límite', 'Área': '', 'Nombre': '',
        'CO2 (ppm)': 5000, 'CO (ppm)': 100, 'Polvo (µg/m3)': 50,
        'COV (mg/m³)': 3, 'Temperatura (°C)': 27, 'Humedad Relativa (%)': 70
    })
    
    # Agregar datos de cada punto
    for punto, datos in datos_aire.items():
        filas.append({
            'Punto': punto,
            'Área': '',  # Para que complete el usuario
            'Nombre': '', # Para que complete el usuario
            'CO (ppm)': datos.get('CO (ppm)', [0, 0, 0])[1],
            'Polvo (µg/m3)': datos.get('Polvo (µg/m3)', [0, 0, 0])[1],
            'Temperatura (°C)': datos.get('Temperatura (°C)', [0, 0, 0])[1],
            'Humedad Relativa (%)': datos.get('Humedad Relativa (%)', [0, 0, 0])[1]
        })
    
    return pd.DataFrame(filas)

def crear_resumen_ruido(datos_ruido):
    """Crea DataFrames para resumen de ruido"""
    leq_filas = []
    lcpk_filas = []
    
    for punto, datos in datos_ruido.items():
        # Datos Leq
        leq_filas.append({
            'Punto': punto,
            'Área': '',
            'Mínimo dB(A)': datos.get('Leq-1', [0, 0, 0])[0],
            'Promedio dB (A)': datos.get('Leq-1', [0, 0, 0])[1],
            'Máximo dB (A)': datos.get('Leq-1', [0, 0, 0])[2],
            'Límite dB (A)': 85,
            'Detalle': ''
        })
        
        # Datos Lcpk
        lcpk_filas.append({
            'Punto': punto,
            'Área': '',
            'Mínimo dB(A)': datos.get('Lapk-1', [0, 0, 0])[0],
            'Promedio dB (A)': datos.get('Lapk-1', [0, 0, 0])[1],
            'Máximo dB (A)': datos.get('Lapk-1', [0, 0, 0])[2],
            'Límite dB (A)': 140,
            'Detalle': ''
        })
    
    return pd.DataFrame(leq_filas), pd.DataFrame(lcpk_filas)

# ===== INTERFAZ PRINCIPAL =====
def main():
    st.title("🏭 Procesador de Datos Ambientales")
    st.markdown("---")
    
    # Sidebar
    st.sidebar.header("Configuración")
    nombre_empresa = st.sidebar.text_input("Nombre de la Empresa", "Mi Empresa")
    
    st.sidebar.markdown("---")
    st.sidebar.info("""
    **Instrucciones:**
    1. Sube los archivos según el tipo
    2. Procesa los datos  
    3. Descarga el reporte consolidado
    """)
    
    # Pestañas
    tab1, tab2, tab3 = st.tabs(["📁 Carga de Archivos", "⚙️ Procesar", "📊 Resultados"])
    
    with tab1:
        st.header("Carga de Archivos")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("🌫️ Calidad de Aire (3M)")
            archivos_3m_aire = st.file_uploader(
                "Subir archivos 3M Aire (.xls)",
                type=['xls'],
                accept_multiple_files=True,
                key="aire_3m"
            )
            
        with col2:
            st.subheader("🔊 Niveles de Ruido")
            archivos_ruido = st.file_uploader(
                "Subir archivos Ruido 3M (.xls)",
                type=['xls'],
                accept_multiple_files=True,
                key="ruido"
            )
    
    with tab2:
        st.header("Procesar Datos")
        
        if archivos_3m_aire:
            st.success(f"✅ {len(archivos_3m_aire)} archivos de aire listos")
        
        if archivos_ruido:
            st.success(f"✅ {len(archivos_ruido)} archivos de ruido listos")
        
        if st.button("🚀 Procesar Todos los Datos", type="primary"):
            # Procesar aire
            if archivos_3m_aire:
                with st.spinner("Procesando datos de aire..."):
                    resultados_aire = procesar_3m_aire(archivos_3m_aire)
                    st.session_state.resultados_aire = resultados_aire
                    st.success(f"✅ {len(resultados_aire)} puntos de aire procesados")
            
            # Procesar ruido
            if archivos_ruido:
                with st.spinner("Procesando datos de ruido..."):
                    resultados_ruido = procesar_ruido_3m(archivos_ruido)
                    st.session_state.resultados_ruido = resultados_ruido
                    st.success(f"✅ {len(resultados_ruido)} puntos de ruido procesados")
    
    with tab3:
        st.header("Resultados y Descarga")
        
        # Mostrar resultados de aire
        if 'resultados_aire' in st.session_state:
            st.subheader("📊 Resumen Aire")
            df_aire = crear_resumen_aire(st.session_state.resultados_aire)
            st.dataframe(df_aire)
        
        # Mostrar resultados de ruido
        if 'resultados_ruido' in st.session_state:
            st.subheader("📊 Resumen Ruido - Leq")
            df_leq, df_lcpk = crear_resumen_ruido(st.session_state.resultados_ruido)
            st.dataframe(df_leq)
            
            st.subheader("📊 Resumen Ruido - Lcpk")
            st.dataframe(df_lcpk)
        
        # Botón de descarga
        if 'resultados_aire' in st.session_state or 'resultados_ruido' in st.session_state:
            st.markdown("---")
            if st.button("📥 Generar Excel Consolidado", type="primary"):
                with st.spinner("Generando reporte Excel..."):
                    excel_buffer = generar_excel_consolidado(
                        st.session_state.get('resultados_aire'),
                        st.session_state.get('resultados_ruido'),
                        nombre_empresa
                    )
                    
                    st.download_button(
                        label="💾 Descargar Reporte Completo",
                        data=excel_buffer,
                        file_name=f"Reporte_Ambiental_{nombre_empresa}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    st.success("✅ Reporte generado correctamente")

if __name__ == "__main__":
    main()
