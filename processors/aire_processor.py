import pandas as pd
import numpy as np
from datetime import datetime, time

def procesar_aire(archivos_3m, archivo_airthinx=None, tiempos_airthinx=None):
    """
    Procesa datos de aire de equipos 3M y Airthinx
    """
    resultados_3m = procesar_3m_aire(archivos_3m)
    
    if archivo_airthinx and tiempos_airthinx:
        resultados_airthinx = procesar_airthinx(archivo_airthinx, tiempos_airthinx)
        # Combinar resultados
        for punto, datos in resultados_airthinx.items():
            if punto in resultados_3m:
                resultados_3m[punto].update(datos)
    
    return resultados_3m

def procesar_3m_aire(archivos):
    """
    Procesa archivos 3M de calidad de aire
    """
    resultados = {}
    
    for i, archivo in enumerate(archivos, 1):
        try:
            # Leer archivo .xls
            df = pd.read_excel(archivo, header=2, engine='xlrd')
            
            # Limpiar nombres de columnas
            df.columns = [str(col).strip() for col in df.columns]
            
            # Asumir estructura estándar
            if len(df.columns) >= 6:
                resultados[i] = {
                    'CO (ppm)': [
                        safe_min(df.iloc[:, 1]), 
                        safe_mean(df.iloc[:, 1]), 
                        safe_max(df.iloc[:, 1])
                    ],
                    'Polvo (µg/m3)': [
                        safe_min(df.iloc[:, 2]), 
                        safe_mean(df.iloc[:, 2]), 
                        safe_max(df.iloc[:, 2])
                    ],
                    'Humedad Relativa (%)': [
                        safe_min(df.iloc[:, 3]), 
                        safe_mean(df.iloc[:, 3]), 
                        safe_max(df.iloc[:, 3])
                    ],
                    'Temperatura (°C)': [
                        safe_min(df.iloc[:, 4]), 
                        safe_mean(df.iloc[:, 4]), 
                        safe_max(df.iloc[:, 4])
                    ]
                }
                
        except Exception as e:
            print(f"Error procesando archivo {archivo.name}: {str(e)}")
    
    return resultados

def procesar_airthinx(archivo, tiempos_airthinx):
    """
    Procesa archivo Airthinx con configuración de tiempos
    """
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
        print(f"Error procesando Airthinx: {str(e)}")
    
    return resultados

def safe_min(series):
    """Calcula mínimo seguro ignorando NaN"""
    try:
        return float(series.min())
    except:
        return 0.0

def safe_max(series):
    """Calcula máximo seguro ignorando NaN"""
    try:
        return float(series.max())
    except:
        return 0.0

def safe_mean(series):
    """Calcula promedio seguro ignorando NaN"""
    try:
        return float(series.mean())
    except:
        return 0.0