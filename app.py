import streamlit as st
import pandas as pd
import io
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

def procesar_airthinx(archivo, tiempos_airthinx):
    """Procesa archivo Airthinx con configuración de tiempos"""
    resultados = {}
    
    try:
        df = pd.read_excel(archivo, header=0, engine='openpyxl')
        
        # Convertir columna de timestamp
        df['Timestamp'] = pd.to_datetime(df['Timestamp'])
        
        for punto, (inicio, fin) in tiempos_airthinx.items():
            # Combinar fecha del archivo con tiempos configurados
            fecha_base = df['Timestamp'].iloc[0].date()
            inicio_dt = datetime.combine(fecha_base, inicio)
            fin_dt = datetime.combine(fecha_base, fin)
            
            # Filtrar por rango
            mask = (df['Timestamp'] >= inicio_dt) & (df['Timestamp'] <= fin_dt)
            datos_punto = df[mask]
            
            if not datos_punto.empty:
                resultados[punto] = {
                    'CO2 (ppm)': [
                        safe_min(datos_punto.iloc[:, 1]),
                        safe_mean(datos_punto.iloc[:, 1]), 
                        safe_max(datos_punto.iloc[:, 1])
                    ],
                    'COV (mg/m³)': [
                        safe_min(datos_punto.iloc[:, 2]),
                        safe_mean(datos_punto.iloc[:, 2]),
                        safe_max(datos_punto.iloc[:, 2])
                    ]
                }
                
    except Exception as e:
        st.error(f"Error procesando Airthinx: {str(e)}")
    
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

def procesar_estres_termico(archivo):
    """Procesa archivo de estrés térmico"""
    try:
        df = pd.read_excel(archivo, header=0, engine='xlrd')
        
        parametros = {}
        
        # Procesar estructura de columnas múltiples
        for i in range(0, len(df.columns), 2):
            if i + 1 < len(df.columns):
                # Buscar nombre del parámetro en los datos
                nombre_parametro = None
                for j in range(len(df)):
                    unidad_val = df.iloc[j, i + 1]
                    if pd.notna(unidad_val) and isinstance(unidad_val, str):
                        nombre_parametro = str(unidad_val).strip()
                        break
                
                if nombre_parametro:
                    valores = df.iloc[1:, i].dropna()
                    if len(valores) > 0:
                        parametros[nombre_parametro] = [
                            safe_min(valores),
                            safe_mean(valores),
                            safe_max(valores)
                        ]
        
        return parametros
        
    except Exception as e:
        st.error(f"Error procesando estrés térmico: {str(e)}")
        return {}

def generar_excel_consolidado(datos_aire, datos_ruido, datos_et, nombre_empresa):
    """Genera Excel usando solo pandas"""
    output = io.BytesIO()
    
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
        
        # Hoja Resumen Estrés Térmico
        if datos_et:
            df_wbgt, df_otros = crear_resumen_et(datos_et)
            df_wbgt.to_excel(writer, sheet_name='Resumen ET WBGT', index=False)
            if not df_otros.empty:
                df_otros.to_excel(writer, sheet_name='Resumen ET Otros', index=False)
    
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
            'Área': '',
            'Nombre': '',
            'CO2 (ppm)': datos.get('CO2 (ppm)', [0, 0, 0])[1],
            'CO (ppm)': datos.get('CO (ppm)', [0, 0, 0])[1],
            'Polvo (µg/m3)': datos.get('Polvo (µg/m3)', [0, 0, 0])[1],
            'COV (mg/m³)': datos.get('COV (mg/m³)', [0, 0, 0])[1],
            'Temperatura (°C)': datos.get('Temperatura (°C)', [0, 0, 0])[1],
            'Humedad Relativa (%)': datos.get('Humedad Relativa (%)', [0, 0, 0])[1]
        })
    
    return pd.DataFrame(filas)

def crear_resumen_ruido(datos_ruido):
    """Crea DataFrames para resumen de ruido"""
    leq_filas = []
    lcpk_filas = []
    
    for punto, datos in datos_ruido.items():
        leq_filas.append({
            'Punto': punto,
            'Área': '',
            'Mínimo dB(A)': datos.get('Leq-1', [0, 0, 0])[0],
            'Promedio dB (A)': datos.get('Leq-1', [0, 0, 0])[1],
            'Máximo dB (A)': datos.get('Leq-1', [0, 0, 0])[2],
            'Límite dB (A)': 85,
            'Detalle': ''
        })
        
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

def crear_resumen_et(datos_et):
    """Crea DataFrames para resumen de estrés térmico"""
    wbgt_filas = []
    otros_filas = []
    
    for param_name, valores in datos_et.items():
        if 'WBGT' in param_name.upper():
            wbgt_filas.append({
                'Punto': 1,
                'Área': '',
                'Mínimo': valores[0],
                'Promedio': valores[1],
                'Máximo': valores[2],
                'Límite': 26.7,
                'Detalle': ''
            })
        else:
            otros_filas.append({
                'Parámetro': param_name,
                'Mínimo': valores[0],
                'Promedio': valores[1],
                'Máximo': valores[2],
                'Límite': ''
            })
    
    return pd.DataFrame(wbgt_filas), pd.DataFrame(otros_filas)

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
    2. Configura los tiempos para Airthinx
    3. Procesa los datos  
    4. Descarga el reporte consolidado
    """)
    
    # Pestañas
    tab1, tab2, tab3, tab4 = st.tabs(["📁 Carga de Archivos", "⏰ Tiempos Airthinx", "⚙️ Procesar", "📊 Resultados"])
    
    with tab1:
        st.header("Carga de Archivos")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.subheader("🌫️ Calidad de Aire")
            st.info("Archivos 3M (.xls) y Airthinx (.xlsx)")
            
            archivos_3m_aire = st.file_uploader(
                "Subir archivos 3M Aire (.xls)",
                type=['xls'],
                accept_multiple_files=True,
                key="aire_3m"
            )
            
            archivo_airthinx = st.file_uploader(
                "Subir archivo Airthinx (.xlsx)",
                type=['xlsx'],
                key="airthinx"
            )
            
        with col2:
            st.subheader("🔊 Niveles de Ruido")
            st.info("Archivos 3M Ruido (.xls)")
            
            archivos_ruido = st.file_uploader(
                "Subir archivos Ruido 3M (.xls)",
                type=['xls'],
                accept_multiple_files=True,
                key="ruido"
            )
            
        with col3:
            st.subheader("🌡️ Estrés Térmico")
            st.info("Archivo de Estrés Térmico (.xls)")
            
            archivo_et = st.file_uploader(
                "Subir archivo Estrés Térmico (.xls)",
                type=['xls'],
                key="et"
            )
    
    with tab2:
        st.header("Configuración de Tiempos Airthinx")
        
        if archivos_3m_aire:
            num_puntos = len(archivos_3m_aire)
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
            st.warning("⚠️ Primero carga los archivos de aire 3M en la pestaña 'Carga de Archivos'")
    
    with tab3:
        st.header("Procesar Datos")
        
        # Mostrar estado de archivos cargados
        if archivos_3m_aire:
            st.success(f"✅ {len(archivos_3m_aire)} archivos de aire 3M listos")
        
        if archivo_airthinx:
            st.success("✅ Archivo Airthinx listo")
        
        if archivos_ruido:
            st.success(f"✅ {len(archivos_ruido)} archivos de ruido listos")
        
        if archivo_et:
            st.success("✅ Archivo de estrés térmico listo")
        
        if st.button("🚀 Procesar Todos los Datos", type="primary"):
            resultados_totales = {}
            
            # Procesar aire 3M
            if archivos_3m_aire:
                with st.spinner("Procesando datos de aire 3M..."):
                    resultados_aire = procesar_3m_aire(archivos_3m_aire)
                    resultados_totales['aire'] = resultados_aire
                    st.success(f"✅ {len(resultados_aire)} puntos de aire 3M procesados")
            
            # Procesar Airthinx
            if archivo_airthinx and 'tiempos_airthinx' in st.session_state:
                with st.spinner("Procesando datos Airthinx..."):
                    resultados_airthinx = procesar_airthinx(archivo_airthinx, st.session_state.tiempos_airthinx)
                    
                    # Combinar con resultados de aire 3M si existen
                    if 'aire' in resultados_totales:
                        for punto, datos in resultados_airthinx.items():
                            if punto in resultados_totales['aire']:
                                resultados_totales['aire'][punto].update(datos)
                    else:
                        resultados_totales['aire'] = resultados_airthinx
                    
                    st.success(f"✅ {len(resultados_airthinx)} puntos de Airthinx procesados")
            
            # Procesar ruido
            if archivos_ruido:
                with st.spinner("Procesando datos de ruido..."):
                    resultados_ruido = procesar_ruido_3m(archivos_ruido)
                    resultados_totales['ruido'] = resultados_ruido
                    st.success(f"✅ {len(resultados_ruido)} puntos de ruido procesados")
            
            # Procesar estrés térmico
            if archivo_et:
                with st.spinner("Procesando datos de estrés térmico..."):
                    resultados_et = procesar_estres_termico(archivo_et)
                    resultados_totales['estres_termico'] = resultados_et
                    st.success(f"✅ {len(resultados_et)} parámetros de estrés térmico procesados")
            
            st.session_state.resultados = resultados_totales
            st.success("🎉 ¡Todos los datos han sido procesados!")
    
    with tab4:
        st.header("Resultados y Descarga")
        
        if 'resultados' in st.session_state:
            resultados = st.session_state.resultados
            
            # Mostrar resultados de aire
            if 'aire' in resultados:
                st.subheader("📊 Resumen Aire")
                df_aire = crear_resumen_aire(resultados['aire'])
                st.dataframe(df_aire)
            
            # Mostrar resultados de ruido
            if 'ruido' in resultados:
                st.subheader("📊 Resumen Ruido - Leq")
                df_leq, df_lcpk = crear_resumen_ruido(resultados['ruido'])
                st.dataframe(df_leq)
                
                st.subheader("📊 Resumen Ruido - Lcpk")
                st.dataframe(df_lcpk)
            
            # Mostrar resultados de estrés térmico
            if 'estres_termico' in resultados:
                st.subheader("📊 Resumen Estrés Térmico - WBGT")
                df_wbgt, df_otros = crear_resumen_et(resultados['estres_termico'])
                st.dataframe(df_wbgt)
                
                if not df_otros.empty:
                    st.subheader("📊 Resumen Estrés Térmico - Otros Parámetros")
                    st.dataframe(df_otros)
            
            # Botón de descarga
            st.markdown("---")
            if st.button("📥 Generar Excel Consolidado", type="primary"):
                with st.spinner("Generando reporte Excel..."):
                    excel_buffer = generar_excel_consolidado(
                        resultados.get('aire'),
                        resultados.get('ruido'),
                        resultados.get('estres_termico'),
                        nombre_empresa
                    )
                    
                    st.download_button(
                        label="💾 Descargar Reporte Completo",
                        data=excel_buffer,
                        file_name=f"Reporte_Ambiental_{nombre_empresa}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    st.success("✅ Reporte generado correctamente")
        else:
            st.info("👆 Primero procesa los datos en la pestaña 'Procesar'")

if __name__ == "__main__":
    main()
