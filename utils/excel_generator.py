import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
import io

def generar_excel_consolidado(datos_aire, datos_ruido, datos_et, nombre_empresa):
    """
    Genera el Excel consolidado con todas las hojas requeridas
    """
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Hoja Resumen Aire
        if datos_aire:
            df_aire = crear_resumen_aire(datos_aire)
            df_aire.to_excel(writer, sheet_name='Resumen Aire', index=False)
            
            # Hoja Gráficas Aire (datos para gráficas)
            df_graf_aire = crear_datos_graficas_aire(datos_aire)
            df_graf_aire.to_excel(writer, sheet_name='Graficas Aire', index=False)
        
        # Hoja Resumen Ruido
        if datos_ruido:
            df_leq, df_lcpk = crear_resumen_ruido(datos_ruido)
            df_leq.to_excel(writer, sheet_name='Resumen Ruido Leq', index=False)
            df_lcpk.to_excel(writer, sheet_name='Resumen Ruido Lcpk', index=False)
            
            # Hoja Gráficas Ruido
            df_graf_ruido = crear_datos_graficas_ruido(datos_ruido)
            df_graf_ruido.to_excel(writer, sheet_name='Graficas Ruido', index=False)
        
        # Hoja Resumen ET
        if datos_et:
            df_wbgt, df_otros = crear_resumen_et(datos_et)
            df_wbgt.to_excel(writer, sheet_name='Resumen ET WBGT', index=False)
            df_otros.to_excel(writer, sheet_name='Resumen ET Otros', index=False)
            
            # Hoja Gráficas ET
            df_graf_et = crear_datos_graficas_et(datos_et)
            df_graf_et.to_excel(writer, sheet_name='Graficas ET', index=False)
    
    output.seek(0)
    return output.getvalue()

def crear_resumen_aire(datos_aire):
    """
    Crea el resumen de aire con el formato específico
    """
    estructura = {
        'Punto': [], 'Área': [], 'Nombre': [],
        'CO2 (ppm)': [], 'CO (ppm)': [], 'Polvo (µg/m3)': [], 
        'COV (mg/m³)': [], 'Temperatura (°C)': [], 'Humedad Relativa (%)': []
    }
    
    # Agregar fila de límites
    estructura['Punto'].extend(['', '', 'Límite'])
    estructura['Área'].extend(['', '', ''])
    estructura['Nombre'].extend(['', '', ''])
    estructura['CO2 (ppm)'].extend(['', '', 5000])
    estructura['CO (ppm)'].extend(['', '', 100])
    estructura['Polvo (µg/m3)'].extend(['', '', 50])
    estructura['COV (mg/m³)'].extend(['', '', 3])
    estructura['Temperatura (°C)'].extend(['', '', 27])
    estructura['Humedad Relativa (%)'].extend(['', '', 70])
    
    # Agregar datos de cada punto
    for punto, datos in datos_aire.items():
        estructura['Punto'].append(punto)
        estructura['Área'].append('')  # Para que el usuario complete
        estructura['Nombre'].append('')  # Para que el usuario complete
        
        # Promedios para cada parámetro
        estructura['CO2 (ppm)'].append(datos.get('CO2 (ppm)', [0, 0, 0])[1])
        estructura['CO (ppm)'].append(datos.get('CO (ppm)', [0, 0, 0])[1])
        estructura['Polvo (µg/m3)'].append(datos.get('Polvo (µg/m3)', [0, 0, 0])[1])
        estructura['COV (mg/m³)'].append(datos.get('COV (mg/m³)', [0, 0, 0])[1])
        estructura['Temperatura (°C)'].append(datos.get('Temperatura (°C)', [0, 0, 0])[1])
        estructura['Humedad Relativa (%)'].append(datos.get('Humedad Relativa (%)', [0, 0, 0])[1])
    
    return pd.DataFrame(estructura)

def crear_resumen_ruido(datos_ruido):
    """
    Crea resumen de ruido para Leq y Lcpk
    """
    # Estructura Leq
    leq = {
        'Punto': [], 'Área': [], 'Mínimo dB(A)': [], 'Promedio dB (A)': [], 
        'Máximo dB (A)': [], 'Límite dB (A)': [], 'Detalle': []
    }
    
    # Estructura Lcpk
    lcpk = {
        'Punto': [], 'Área': [], 'Mínimo dB(A)': [], 'Promedio dB (A)': [], 
        'Máximo dB (A)': [], 'Límite dB (A)': [], 'Detalle': []
    }
    
    for punto, datos in datos_ruido.items():
        # Datos Leq
        leq['Punto'].append(punto)
        leq['Área'].append('')
        leq['Mínimo dB(A)'].append(datos.get('Leq-1', [0, 0, 0])[0])
        leq['Promedio dB (A)'].append(datos.get('Leq-1', [0, 0, 0])[1])
        leq['Máximo dB (A)'].append(datos.get('Leq-1', [0, 0, 0])[2])
        leq['Límite dB (A)'].append(85)
        leq['Detalle'].append('')
        
        # Datos Lcpk
        lcpk['Punto'].append(punto)
        lcpk['Área'].append('')
        lcpk['Mínimo dB(A)'].append(datos.get('Lapk-1', [0, 0, 0])[0])
        lcpk['Promedio dB (A)'].append(datos.get('Lapk-1', [0, 0, 0])[1])
        lcpk['Máximo dB (A)'].append(datos.get('Lapk-1', [0, 0, 0])[2])
        lcpk['Límite dB (A)'].append(140)
        lcpk['Detalle'].append('')
    
    return pd.DataFrame(leq), pd.DataFrame(lcpk)

def crear_resumen_et(datos_et):
    """
    Crea resumen de estrés térmico
    """
    wbgt = {
        'Punto': [], 'Área': [], 'Mínimo': [], 'Promedio': [], 
        'Máximo': [], 'Límite': [], 'Detalle': []
    }
    
    otros = {
        'Parámetro': [], 'Mínimo': [], 'Promedio': [], 'Máximo': [], 'Límite': []
    }
    
    # Buscar WBGT específicamente
    for param_name, valores in datos_et.items():
        if 'WBGT' in param_name.upper():
            wbgt['Punto'].append(1)  # Asumir primer punto
            wbgt['Área'].append('')
            wbgt['Mínimo'].append(valores[0])
            wbgt['Promedio'].append(valores[1])
            wbgt['Máximo'].append(valores[2])
            wbgt['Límite'].append(26.7)
            wbgt['Detalle'].append('')
        else:
            otros['Parámetro'].append(param_name)
            otros['Mínimo'].append(valores[0])
            otros['Promedio'].append(valores[1])
            otros['Máximo'].append(valores[2])
            otros['Límite'].append('')
    
    return pd.DataFrame(wbgt), pd.DataFrame(otros)

# Funciones para datos de gráficas (similares pero con min, prom, max)
def crear_datos_graficas_aire(datos_aire):
    # Implementar similar a resumen pero con min, prom, max por parámetro
    return pd.DataFrame()

def crear_datos_graficas_ruido(datos_ruido):
    return pd.DataFrame()

def crear_datos_graficas_et(datos_et):
    return pd.DataFrame()