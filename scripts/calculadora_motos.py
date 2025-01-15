import pandas as pd
import numpy as np

# Cargar los datos desde un archivo CSV
def load_data(csv_path="motos_data.csv"):
    try:
        data = pd.read_csv(csv_path, delimiter=';', encoding='utf-8')
        return data
    except Exception as e:
        raise RuntimeError(f"Error al cargar el archivo CSV: {str(e)}")

# Calcular precio estimado
def calculate_price(data, marca, modelo, año, km):
    subset = data[(data['MARCA'] == marca) & (data['MODELO'] == modelo)]
    if any(value is None for value in [avg_price, std_dev, num_motos]):
        return None, None, None, None, None
    print(data)
    subset['PVP'] = pd.to_numeric(subset['PVP'], errors='coerce')
    subset = subset.dropna(subset=['PVP'])
    avg_price = subset['PVP'].mean()

    avg_price = subset['PVP'].mean()
    std_dev = subset['PVP'].std()
    num_motos = len(subset)

    año_diferencia = año - subset['Año'].mean()
    km_diferencia = km - subset['KM'].mean()
    ajuste = -0.1 * año_diferencia + -0.05 * km_diferencia

    precio_estimado = avg_price + ajuste
    posible_variacion = std_dev / 2

    return precio_estimado, posible_variacion, num_motos, subset['Año'].min(), subset['KM'].max()
