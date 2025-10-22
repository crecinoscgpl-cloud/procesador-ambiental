import pandas as pd

def procesar_estres_termico(archivo):
    """
    Procesa archivo de estrés térmico con estructura especial
    """
    try:
        df = pd.read_excel(archivo, header=0, engine='xlrd')
        
        parametros = {}
        
        # Procesar estructura de columnas múltiples
        for i in range(0, len(df.columns), 2):
            if i + 1 < len(df.columns):
                col_valor = df.columns[i]
                col_unidad = df.columns[i + 1]
                
                # Obtener nombre del parámetro de la primera fila válida
                nombre_parametro = None
                for j in range(len(df)):
                    unidad_val = df.iloc[j, i + 1]
                    if pd.notna(unidad_val) and isinstance(unidad_val, str):
                        nombre_parametro = str(unidad_val).strip()
                        break
                
                if nombre_parametro:
                    valores = df.iloc[1:, i].dropna()  # Saltar primera fila (encabezados)
                    if len(valores) > 0:
                        parametros[nombre_parametro] = [
                            safe_min(valores),
                            safe_mean(valores),
                            safe_max(valores)
                        ]
        
        return parametros
        
    except Exception as e:
        print(f"Error procesando estrés térmico: {str(e)}")
        return {}

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