import streamlit as st
import pandas as pd
import subprocess
import os

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
    file_ext = required_file.split('.')[-1]
    f = st.file_uploader(f"Sube el archivo {required_file}", type=file_ext)
    if f is not None:
        with open(required_file, "wb") as out_file:
            out_file.write(f.getbuffer())
        uploaded_files[required_file] = required_file

# Solo permitimos ejecutar cuando todos los archivos están subidos
if len(uploaded_files) == len(script_info["required_files"]):
    if st.button("Ejecutar Script"):
        # Preparar los argumentos para ejecutar el script
        if script_choice == "DAILY":
            args = ["python", script_info["script_path"],
                    "FC.xlsx", "AB.xlsx", "FT.xlsx", "Compras.xlsx"]
        else:
            args = ["python", script_info["script_path"]]

        try:
            subprocess.run(args, check=True)
            st.success(f"{script_choice} se ejecutó correctamente.")
            st.info("Script finalizado. No hay archivo de salida para descargar.")

        except subprocess.CalledProcessError as e:
            st.error(f"Error al ejecutar {script_choice}: {e}")


