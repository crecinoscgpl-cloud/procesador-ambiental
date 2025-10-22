import pandas as pd

def procesar_ruido(archivos):
    """
    Procesa archivos de ruido 3M
    """
    resultados = {}
    
    for i, archivo in enumerate(archivos, 1):
        try:
            df = pd.read_excel(archivo, header=2, engine='xlrd')
            
            # Limpiar nombres de columnas
            df.columns = [str(col).strip() for col in df.columns]
            
            if len(df.columns) >= 3:
                resultados[i] = {
                    'Lapk-1': [
                        safe_min(df.iloc[:, 1]),
                        safe_mean(df.iloc[:, 1]), 
                        safe_max(df.iloc[:, 1])
                    ],
                    'Leq-1': [
                        safe_min(df.iloc[:, 2]),
                        safe_mean(df.iloc[:, 2]),
                        safe_max(df.iloc[:, 2])
                    ]
                }
                
        except Exception as e:
            print(f"Error procesando archivo de ruido {archivo.name}: {str(e)}")
    
    return resultados

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