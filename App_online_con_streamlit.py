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
    },
    "Unnax": {
        "script_path": "Credit stock.py",
        "required_files": ["credit_data.csv"]
    },
    "Facturar ventas": {
        "script_path": "Credit stock.py",
        "required_files": ["credit_data.csv"]
    },
    "Facturar compras": {
        "script_path": "Credit stock.py",
        "required_files": ["credit_data.csv"]
    },
    "Abonos incoming": {
        "script_path": "Credit stock.py",
        "required_files": ["credit_data.csv"]
    },
    "Stripe": {
        "script_path": "Credit stock.py",
        "required_files": ["credit_data.csv"]
    }
}

script_choice = st.selectbox("Selecciona funcionalidad", list(scripts_info.keys()))
script_info = scripts_info[script_choice]

st.write("Sube los archivos necesarios:")


uploaded_files = {}

for required_file in script_info["required_files"]:
    file_ext = required_file.split('.')[-1]
    f = st.file_uploader(f"Sube el archivo {required_file}", type=file_ext)
    if f is not None:
        with open(required_file, "wb") as out_file:
            out_file.write(f.getbuffer())
        uploaded_files[required_file] = required_file


if len(uploaded_files) == len(script_info["required_files"]):
    if st.button("Ejecutar Script"):

        if script_choice == "DAILY":
            args = [script_info["script_path"],
                    "FC.xlsx", "AB.xlsx", "FT.xlsx", "Compras.xlsx"]
        else:
            args = [script_info["script_path"]]

        try:
            subprocess.run(args, check=True)
            st.success(f"{script_choice} se ejecutó correctamente.")

            output_filename = "Reportdaily.xlsx"
            if os.path.exists(output_filename):
                with open(output_filename, "rb") as f:
                    st.download_button(
                        label="Descargar Reportdaily.xlsx",
                        data=f,
                        file_name="Reportdaily.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            else:
                st.error("Error: No se generó el archivo de salida.")


        except subprocess.CalledProcessError as e:
            st.error(f"Error al ejecutar {script_choice}: {e}")
