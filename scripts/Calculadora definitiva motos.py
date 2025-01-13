from io import BytesIO
import streamlit as st
import pandas as pd
from sklearn.preprocessing import OneHotEncoder, StandardScaler
from sklearn.compose import ColumnTransformer
from sklearn.linear_model import LinearRegression
from sklearn.pipeline import Pipeline

# Configuración inicial de la app
st.title("Calculadora PVP para Motos")

# Subida de archivo CSV
st.header("Sube el archivo de datos para calcular el precio")
uploaded_file = st.file_uploader("Sube el archivo CSV", type="csv")

if uploaded_file:
    # Cargar datos
    data = pd.read_csv(uploaded_file, delimiter=';', encoding='utf-8')

    # Verificar que el archivo contiene las columnas necesarias
    required_columns = ['MARCA', 'MODELO', 'Año', 'KM', 'PVP']
    if not all(col in data.columns for col in required_columns):
        st.error(f"El archivo debe contener las siguientes columnas: {', '.join(required_columns)}")
    else:
        # Configuración del modelo
        preprocessor = ColumnTransformer(
            transformers=[
                ('num', StandardScaler(), ['Año', 'KM']),
                ('cat', OneHotEncoder(drop='first', handle_unknown='ignore'), ['MARCA', 'MODELO'])
            ])

        model = Pipeline([
            ('preprocessor', preprocessor),
            ('regressor', LinearRegression())
        ])

        # Entrenamiento del modelo
        X = data[['MARCA', 'MODELO', 'Año', 'KM']]
        y = data['PVP']
        model.fit(X, y)

        # Interfaz de usuario para la predicción
        st.sidebar.header("Introduce los datos de la moto")

        marca = st.sidebar.selectbox("Selecciona la marca", options=data['MARCA'].unique())

        # Actualizar los modelos según la marca seleccionada
        modelos_disponibles = data[data['MARCA'] == marca]['MODELO'].unique()
        modelo = st.sidebar.selectbox("Selecciona el modelo", options=modelos_disponibles)

        año = st.sidebar.number_input(
            "Introduce el año de fabricación",
            min_value=int(data['Año'].min()),
            max_value=int(data['Año'].max()),
            value=int(data['Año'].mean()),
            step=1
        )

        kilometraje = st.sidebar.number_input(
            "Introduce el kilometraje",
            min_value=0,
            max_value=int(data['KM'].max()),
            value=int(data['KM'].median()),
            step=1000
        )

        # Botón para calcular el precio
        if st.sidebar.button("Calcular Precio"):
            # Realizar la predicción
            input_data = pd.DataFrame({
                'MARCA': [marca],
                'MODELO': [modelo],
                'Año': [año],
                'KM': [kilometraje]
            })

            predicted_price = model.predict(input_data)[0]

            # Mostrar resultados
            st.subheader("Resultados de la predicción")
            st.write(f"**Precio estimado:** {predicted_price:,.2f} €")

            # Datos adicionales
            subset_data = data[data['MODELO'] == modelo]
            num_motos = len(subset_data)
            std_dev = subset_data['PVP'].std()
            posible_precio = std_dev / 2
            min_año = int(subset_data['Año'].min())
            max_km = int(subset_data['KM'].max())

            st.write(f"**Variación estimada del precio:** +/- {posible_precio:,.2f} €")
            st.write(f"**Año más antiguo del modelo:** {min_año}")
            st.write(f"**Kilometraje máximo registrado:** {max_km} KM")
            st.write(f"**Número de motos en el análisis:** {num_motos}")
else:
    st.info("Sube un archivo CSV para comenzar.")
