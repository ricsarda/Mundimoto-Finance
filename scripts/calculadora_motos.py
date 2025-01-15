import pandas as pd
import numpy as np

# Cargar los datos desde un archivo CSV
def load_data(csv_path="motos_data.csv"):
    try:
        data = pd.read_csv(csv_path, delimiter=';', encoding='utf-8')
        return data
    except Exception as e:
        raise RuntimeError(f"Error al cargar el archivo CSV: {str(e)}")

def calculate_price(data, marca, modelo, año, km):
    subset = data[(data['MARCA'] == marca) & (data['MODELO'] == modelo)]


    subset['PVP'] = pd.to_numeric(subset['PVP'], errors='coerce')
    subset = subset.dropna(subset=['PVP'])  # Eliminar filas donde 'PVP' sea NaN


    avg_price = subset['PVP'].mean()
    std_dev = subset['PVP'].std()
    num_motos = len(subset)

    año_diferencia = año - subset['Año'].mean()
    km_diferencia = km - subset['KM'].mean()
    ajuste = -0.1 * año_diferencia + -0.05 * km_diferencia

    precio = avg_price + ajuste
    variacion = std_dev / 2

    min_año = subset['Año'].min() if not subset['Año'].isnull().all() else None
    max_km = subset['KM'].max() if not subset['KM'].isnull().all() else None

    return precio, variacion, num_motos, min_año, max_km, subset
