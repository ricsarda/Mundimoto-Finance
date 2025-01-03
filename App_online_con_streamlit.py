from io import BytesIO
import streamlit as st
import pandas as pd
import importlib.util
import sys
import os
from datetime import datetime

fecha_actual = datetime.now()
fecha = fecha_actual.strftime("%d-%m-%Y")

# Configuración inicial de la app
st.title("Mundimoto Finance")
st.sidebar.header("Configuración")

# Selección del script
script_option = st.sidebar.selectbox(
    "Selecciona función para ejecutar:",
    ("Credit Stock", "Daily Report", "Financiaciones Santander", "Performance Comerciales B2C")
)

st.write(f"Has seleccionado: {script_option}")

# Función para cargar y ejecutar un script externo
def load_and_execute_script(script_name, files, new_excel=None, month=None, year=None):
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
        for key, file in files.items():
            buffer = BytesIO(file.read())  # Convertir archivo a BytesIO
            buffer.seek(0)  # Reiniciar puntero
            processed_files[key] = buffer

        # Llamar a la función principal del script con los parámetros adicionales
        result = module.main(processed_files, new_excel ,month, year)
        return result
        
    except FileNotFoundError as e:
        st.error(f"Error de archivo: {str(e)}")
    except AttributeError as e:
        st.error(f"Error en el script {script_name}: {str(e)}")
    except KeyError as e:
        st.error(f"Error con los archivos subidos: falta el archivo clave {str(e)}.")
    except Exception as e:
        st.error(f"Error inesperado al ejecutar el script {script_name}: {str(e)}")

# Subida de archivos según el script seleccionado
if script_option == "Daily Report":
    st.header("Archivos")
    # Selección de Mes y Año
    st.subheader("Selecciona el Mes y Año:")
    uploaded_month = st.selectbox("Mes", range(1, 13), index=datetime.now().month - 1)
    uploaded_year = st.number_input("Año", min_value=2000, max_value=datetime.now().year, value=datetime.now().year)

    uploaded_FC = st.file_uploader("Sube el archivo FC", type=["xlsx"])
    uploaded_AB = st.file_uploader("Sube el archivo AB", type=["xlsx"])
    uploaded_FT = st.file_uploader("Sube el archivo FT", type=["xls"])
    uploaded_Compras = st.file_uploader("Sube el archivo de Compras", type=["xlsx"])
    
    uploaded_files = {
    "FC": uploaded_FC, #FC
    "AB": uploaded_AB, #AB
    "FT": uploaded_FT, #FT
    "Compras": uploaded_Compras, #Compras
    }

    if all(uploaded_files.values()):
        if st.button("Ejecutar"):
            load_and_execute_script("Daily Report", uploaded_files, uploaded_month, uploaded_year)

elif script_option == "Credit Stock":
    st.header("Archivos")
    uploaded_metabase = st.file_uploader("Sube el archivo Metabase", type=["xlsx"])
    uploaded_santander = st.file_uploader("Sube el archivo Santander", type=["xlsx"])
    uploaded_sabadell = st.file_uploader("Sube el archivo Sabadell", type=["xls"])
    uploaded_sofinco = st.file_uploader("Sube el archivo Sofinco", type=["xlsx"])
    
    uploaded_files = {
    "Metabase": uploaded_metabase,
    "Santander": uploaded_santander,
    "Sabadell": uploaded_sabadell,
    "Sofinco": uploaded_sofinco,
        }

    if all(uploaded_files.values()):
        if st.button("Ejecutar"):
            try:

                new_excel = BytesIO()

                excel_result = load_and_execute_script(
                    "Credit stock",
                    uploaded_files,
                    new_excel 
                )

                if excel_result is not None:
                    st.success("¡HECHO!")
                    st.download_button(
                        label="Descargar",
                        data=excel_result.getvalue(),
                        file_name=f"Credit Stock {fecha}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"Error al ejecutar el script: {str(e)}")

elif script_option == "Performance Comerciales B2C":

    st.subheader("Selecciona el Mes y Año:")
    uploaded_month = st.selectbox("Mes", range(1, 13), index=datetime.now().month - 1)
    uploaded_year = st.number_input("Año", min_value=2000, max_value=datetime.now().year, value=datetime.now().year)
    
    st.header("Archivos")
    uploaded_FC = st.file_uploader("Sube el archivo FC", type=["xlsx"])
    uploaded_AB = st.file_uploader("Sube el archivo AB", type=["xlsx"])
    uploaded_FT = st.file_uploader("Sube el archivo FT", type=["xlsx"])
    uploaded_ventas = st.file_uploader("Sube el archivo ventas", type=["xlsx"])
    uploaded_leads= st.file_uploader("Sube el archivo leads", type=["xlsx"])
    uploaded_anterior = st.file_uploader("Sube el archivo anterior", type=["xlsx"])
    uploaded_financiacion = st.file_uploader("Sube el archivo financiaciones", type=["xlsx"])
    
    uploaded_files = {
    "inf_usu_FC": uploaded_FC, #FC
    "inf_usu_AB": uploaded_AB, #AB
    "inf_usu_FT": uploaded_FT, #FT
    "archivo_ventas": uploaded_ventas, # Ruta del archivo del Report de comerciales Solo detalles
    "archivo_leads": uploaded_leads, # Ruta del archivo leads Solo detalles
    "sellers_anterior": uploaded_anterior,
    "archivo_financiacion": uploaded_financiacion, # Ruta del archivo de financiaciones
        }

    if all(uploaded_files.values()):
        if st.button("Ejecutar"):
            try:

                new_excel = BytesIO()

                excel_result = load_and_execute_script(
                    "Performance Comerciales B2C",
                    uploaded_files,
                    new_excel,
                    uploaded_month,
                    uploaded_year
                )

                if excel_result is not None:
                    st.success("¡HECHO!")
                    st.download_button(
                        label="Descargar",
                        data=excel_result.getvalue(),
                        file_name=f"Performance Comerciales B2C {fecha}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"Error al ejecutar el script: {str(e)}")
