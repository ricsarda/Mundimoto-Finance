import streamlit as st
import pandas as pd
import subprocess
import os

st.title("Mundimoto Finance")

# Definimos qué archivos necesita cada script
scripts_info = {
    "DAILY": {
        "script_path": "DAILY.py",
        "required_files": ["FC.xlsx", "FT.xlsx", "AB.xlsx", "Compras.xlsx"]  
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
            # FC.xlsx AB.xlsx FT.xlsx Compras.xlsx DAILY.xlsx
            # Nota: Fíjate en el orden de argumentos dentro de DAILY.py y ajústalos si es necesario
            args = ["python", script_info["script_path"],
                    "FC.xlsx", "AB.xlsx", "FT.xlsx", "Compras.xlsx", "DAILY.xlsx"]
        else:
            # Para Credit Stock u otros, ajusta según sus necesidades
            args = ["python", script_info["script_path"]]

        try:
            subprocess.run(args, check=True)
            st.success(f"{script_choice} se ejecutó correctamente.")

            # Si es DAILY, ofrecer la descarga de DAILY.xlsx si existe
            if script_choice == "DAILY" and os.path.exists("DAILY.xlsx"):
                with open("DAILY.xlsx", "rb") as f:
                    data = f.read()
                st.download_button(
                    label="Descargar DAILY.xlsx",
                    data=data,
                    file_name="DAILY.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"Error al ejecutar {script_choice}: {e}")
