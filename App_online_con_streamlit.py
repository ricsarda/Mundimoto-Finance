import streamlit as st
import pandas as pd
import subprocess
import os
from io import BytesIO

st.title("Mundimoto Finance")

# Definimos qué archivos necesita cada script
scripts_info = {
    "DAILY": {
        "script_path": "DAILY.py",
        "required_files": ["FC.xlsx", "AB.xlsx", "FT.xlsx", "Compras.xlsx"]  
    },
    "Credit Stock": {
        "script_path": "Credit stock.py",
        "required_files": ["credit_data.csv"]
    }
}

# Elegir el script
script_choice = st.selectbox("Selecciona funcionalidad", list(scripts_info.keys()))
script_info = scripts_info[script_choice]

st.write("Sube los archivos necesarios:")

# Contenedor para guardar los archivos subidos
uploaded_files = {}

# Por cada archivo requerido, ponemos un file_uploader
for required_file in script_info["required_files"]:
    # Determinar el tipo de archivo por la extensión
    file_ext = required_file.split('.')[-1]
    if file_ext in ["csv", "xlsx"]:
        file_type = [file_ext]
    else:
        file_type = None

    f = st.file_uploader(f"Sube el archivo {required_file}", type=file_type)
    if f is not None:
        # Guardar el archivo subido con el mismo nombre localmente
        with open(required_file, "wb") as out_file:
            out_file.write(f.getbuffer())
        uploaded_files[required_file] = required_file

# Solo permitimos ejecutar cuando todos los archivos están subidos
if len(uploaded_files) == len(script_info["required_files"]):
    if st.button("Ejecutar Script"):
        # Preparar los argumentos para ejecutar el script
        if script_choice == "DAILY":
            # Llamamos a DAILY.py con los 5 argumentos:
            args = ["python", script_info["script_path"],
                    "FC.xlsx", "AB.xlsx", "FT.xlsx", "Compras.xlsx"]
        else:
            # Para Credit Stock u otros, ajusta según sus necesidades
            args = ["python", script_info["script_path"]]

        try:
            subprocess.run(args, check=True)
            st.success(f"{script_choice} se ejecutó correctamente.")

            # Si es DAILY, leer el resultado como DataFrame y ofrecer descarga
            if script_choice == "DAILY" and os.path.exists("DAILY.xlsx"):
                # Leer el archivo resultante en un DataFrame
                Reportdaily = pd.read_excel("DAILY.xlsx")

                # Crear un buffer en memoria para exportar como Excel
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    Reportdaily.to_excel(writer, index=False, sheet_name='Resultados')
                output.seek(0)

                # Botón para descargar el archivo
                st.download_button(
                    label="Descargar Reportdaily.xlsx",
                    data=output,
                    file_name="Reportdaily.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("Error: No se generó el archivo DAILY.xlsx.")

        except Exception as e:
            st.error(f"Error al ejecutar {script_choice}: {e}")

