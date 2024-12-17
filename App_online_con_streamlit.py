import streamlit as st
import pandas as pd
import subprocess
import os

st.title("Mundimoto Finance")

# Definimos qué archivos necesita cada script
scripts_info = {
    "DAILY": {
        "script_path": "DAILY.py",
        "required_files": ["FC.xlsx", "FT.xlsx","AB.xlsx","Compras.xlsx"]  # Ejemplo
    },
    "Credit Stock": {
        "script_path": "Credit stock.py",
        "required_files": ["credit_data.csv"]  # Ejemplo
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
    # Determinamos el tipo de archivo por la extensión (opcional, puedes simplificar)
    file_ext = required_file.split('.')[-1]
    if file_ext in ["csv", "xlsx"]:
        file_type = [file_ext]
    else:
        file_type = None  # cualquier tipo si no reconocemos la extensión

    f = st.file_uploader(f"Sube el archivo {required_file}", type=file_type)
    if f is not None:
        uploaded_files[required_file] = f

# Solo permitimos ejecutar cuando todos los archivos están subidos
if len(uploaded_files) == len(script_info["required_files"]):
    if st.button("Ejecutar Script"):
        # Guardamos los archivos subidos en el entorno local del servidor
        for fname, uploaded_f in uploaded_files.items():
            # Detectamos el tipo de archivo para leerlo y guardarlo en un formato estándar (por ejemplo CSV)
            if fname.endswith(".csv"):
                df = pd.read_csv(uploaded_f)
                df.to_csv(fname, index=False)
            elif fname.endswith(".xlsx"):
                df = pd.read_excel(uploaded_f)
                # Podrías guardarlo como CSV o dejarlo como xlsx.
                # Aquí lo guardamos como CSV para simplificar el tratamiento posterior:
                base_name = os.path.splitext(fname)[0]
                df.to_csv(base_name + ".csv", index=False)
                # Si tu script DAILY.py requiere el xlsx tal cual, en vez de convertir a csv:
                # with open(fname, "wb") as out:
                #     out.write(uploaded_f.getbuffer())

        # Ejecutamos el script
        try:
            # Ajusta la ruta a Python si es necesario
            # Aquí asumimos que el script se puede ejecutar directamente con "python"
            subprocess.run(["python", script_info["script_path"]], check=True)
            st.success(f"{script_choice} se ejecutó correctamente.")
        except Exception as e:
            st.error(f"Error al ejecutar {script_choice}: {e}")

