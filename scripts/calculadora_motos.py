import pandas as pd
import numpy as np
from sklearn.linear_model import LinearRegression

def load_data(csv_path="Motos para calcular.csv"):
    try:
        data = pd.read_csv(csv_path, delimiter=';', encoding='utf-8')
        return data
    except Exception as e:
        raise RuntimeError(f"Error al cargar el archivo CSV: {str(e)}")

def calculate_price(data, marca, modelo, año, km):
    # Filtramos solo las filas que coincidan con la marca y el modelo
    subset = data[(data["MARCA"] == marca) & (data["MODELO"] == modelo)].copy()

    # Aseguramos que las columnas necesarias tengan valores numéricos
    # (PVP, Año, KM). Si hay problemas (texto, NaN, etc.), los convertimos
    # y eliminamos las filas no válidas.
    subset["PVP"] = pd.to_numeric(subset["PVP"], errors="coerce")
    subset["Año"] = pd.to_numeric(subset["Año"], errors="coerce")
    subset["KM"] = pd.to_numeric(subset["KM"], errors="coerce")
    subset.dropna(subset=["PVP","Año","KM"], inplace=True)

    # Si no hay datos suficientes tras filtrar, salimos
    if subset.empty:
        return None, None, None, None, None, None

    # Definimos las variables independientes (Año y KM) y la variable objetivo (PVP)
    X = subset[["Año","KM"]]
    y = subset["PVP"]

    # Entrenamos un modelo de regresión lineal
    model = LinearRegression()
    model.fit(X, y)

    # Obtenemos la predicción para el año y km que introduce el usuario
    precio_est = model.predict([[año, km]])[0]

    # Calculamos la desviación estándar de los residuos para estimar
    # qué tanta variación puede haber en nuestra predicción
    predicciones_entrenamiento = model.predict(X)
    residuos = y - predicciones_entrenamiento
    std_dev = residuos.std()

    # Extraemos la cantidad de motos consideradas y algunos valores de referencia
    num_motos = len(subset)
    min_año = subset["Año"].min()
    max_km  = subset["KM"].max()

    # Definimos la variación como la std. de los residuos (puedes ajustar si lo prefieres)
    variacion = std_dev

    # Devolvemos todos los valores esperados
    return precio_est, variacion, num_motos, min_año, max_km, subset
