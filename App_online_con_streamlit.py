from io import BytesIO
import streamlit as st
import pandas as pd
import importlib.util
import sys
import os

# Configuración inicial de la app
st.title("Mundimoto Finance")
st.sidebar.header("Configuración")

# Selección del script
script_option = st.sidebar.selectbox(
    "Selecciona función para ejecutar:",
    ("DAILY", "Credit Stock")
)

st.write(f"Has seleccionado: {script_option}")

# Función para cargar y ejecutar un script externo
def load_and_execute_script(script_name, files):
    try:
        script_path = os.path.join("scripts", f"{script_name}.py")
        if not os.path.exists(script_path):
            raise FileNotFoundError(f"El script {script_name} no fue encontrado en {script_path}")
        
        spec = importlib.util.spec_from_file_location(script_name, script_path)
        module = importlib.util.module_from_spec(spec)
        sys.modules[script_name] = module
        spec.loader.exec_module(module)
        # Convertir archivos subidos a buffers y reiniciar el puntero
        processed_files = {}
        # Llama a la función principal del script con los archivos procesados
        result = module.main(processed_files)
    except FileNotFoundError as e:
        st.error(f"Error de archivo: {str(e)}")
    except AttributeError as e:
        st.error(f"Error en el script {script_name}: {str(e)}")
    except KeyError as e:
        st.error(f"Error con los archivos subidos: falta el archivo clave {str(e)}.")
    except Exception as e:
        st.error(f"Error inesperado al ejecutar el script {script_name}: {str(e)}")

# Subida de archivos según el script seleccionado
if script_option == "DAILY":
    st.header("Subida de archivos para DAILY")
    uploaded_files = {
        "FC": st.file_uploader("Sube el archivo FC", type=["xlsx"]),
        "AB": st.file_uploader("Sube el archivo AB", type=["xlsx"]),
        "FT": st.file_uploader("Sube el archivo FT", type=["xlsx"]),
        "Compras": st.file_uploader("Sube el archivo de Compras", type=["xlsx"])
    }

    if all(uploaded_files.values()):
        if st.button("Ejecutar DAILY"):
            load_and_execute_script("DAILY", uploaded_files)

elif script_option == "Credit Stock":
    st.header("Subida de archivos para Credit Stock")
    uploaded_files = {
        "Metabase": st.file_uploader("Sube el archivo Metabase", type=["xlsx"]),
        "Santander": st.file_uploader("Sube el archivo Santander", type=["xlsx"]),
        "Sabadell": st.file_uploader("Sube el archivo Sabadell", type=["xls"]),
        "Sofinco": st.file_uploader("Sube el archivo Sofinco", type=["xlsx"])
    }

    if all(uploaded_files.values()):
        if st.button("Ejecutar Script Credit Stock"):
            load_and_execute_script("Credit stock", uploaded_files)
