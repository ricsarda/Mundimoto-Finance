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
    ("Credit Stock", "Calculadora Precios B2C", "Daily Report", "Facturación Ventas B2C", "Facturación Compras","Financiaciones Santander", "Performance Comerciales B2C", "Unnax CaixaBank", "Unnax Easy Payment", "Stripe")
)

st.write(f"Has seleccionado: {script_option}")

# Función para cargar y ejecutar un script externo
def load_and_execute_script(script_name, files, pdfs=None, new_excel=None, month=None, year=None):
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
        result = module.main(processed_files, processed_pdfs, new_excel ,month, year)
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

    uploaded_FC = st.file_uploader("FC", type=["xlsx"])
    uploaded_AB = st.file_uploader("AB", type=["xlsx"])
    uploaded_FT = st.file_uploader("FT", type=["xls"])
    uploaded_Compras = st.file_uploader("Compras", type=["xlsx"])
    
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
    uploaded_metabase = st.file_uploader("Metabase", type=["xlsx"])
    uploaded_santander = st.file_uploader("Santander", type=["xlsx"])
    uploaded_sabadell = st.file_uploader("Sabadell", type=["xls"])
    uploaded_sofinco = st.file_uploader("Sofinco", type=["xlsx"])
    
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
    uploaded_FC = st.file_uploader("FC", type=["xlsx"])
    uploaded_AB = st.file_uploader("AB", type=["xlsx"])
    uploaded_FT = st.file_uploader("FT", type=["xlsx"])
    uploaded_ventas = st.file_uploader("ventas", type=["xlsx"])
    uploaded_leads= st.file_uploader("leads", type=["xlsx"])
    uploaded_anterior = st.file_uploader("Anterior", type=["xlsx"])
    uploaded_financiacion = st.file_uploader("Financiaciones 2025", type=["xlsx"])
    
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

elif script_option == "Financiaciones Santander":
    st.header("Archivos y PDFs")
    # Pedimos que el usuario suba uno o varios PDFs
    uploaded_pdfs = st.file_uploader(
        "Sube los PDFs",
        type=["pdf"],
        accept_multiple_files=True
    )
    
    upload_Clientes = st.file_uploader("Clientes-Netsiut", type=["xlsx"])
    upload_ventas_SF = st.file_uploader("Ventas-SalesForce", type=["xlsx"])
    
    uploaded_files = {
    "Clientes": upload_Clientes,
    "Ventas": upload_ventas_SF,
    }
    
    if all(uploaded_files.values()):
        if st.button("Ejecutar"):
            try:

                new_excel = BytesIO()

                excel_result = load_and_execute_script(
                    "Financiaciones Santander",
                    uploaded_files,
                    uploaded_pdfs,
                    new_excel,
                    uploaded_month,
                    uploaded_year
                )
                if excel_result is not None:
                    st.success("¡HECHO!")
                    st.download_button(
                        label="Descargar",
                        data=excel_result.getvalue(),
                        file_name=f"Financiaciones Santander {fecha}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"Error al ejecutar el script: {str(e)}")

elif script_option == "Facturación Ventas B2C":
    st.header("Archivos")
    uploaded_clients = st.file_uploader("clients", type=["csv"])
    uploaded_mheaders = st.file_uploader("motorbike_headers", type=["csv"])
    uploaded_mlines = st.file_uploader("motorbike_lines", type=["csv"])
    uploaded_sheaders = st.file_uploader("services_headers", type=["csv"])
    uploaded_slines = st.file_uploader("services_lines", type=["csv"])
    uploaded_Netsuitclientes = st.file_uploader("Clientes de Netsuit", type=["xlsx"])
    uploaded_Netsuitarticulos = st.file_uploader("Artículos de Netsuit", type=["xlsx"])
    uploaded_salesforce = st.file_uploader("Salesforce", type=["xlsx"])
    
    uploaded_files = {
    "clients": uploaded_clients,
    "motorbike_headers":uploaded_mheaders,
    "motorbike_lines":uploaded_mlines,
    "services_headers":uploaded_sheaders,
    "services_lines":uploaded_slines,
    "Clientes de Netsuit":uploaded_Netsuitclientes,
    "Artículos de Netsuit":uploaded_Netsuitarticulos,
    "Salesforce":uploaded_salesforce,
    }

    if all(uploaded_files.values()):
        if st.button("Ejecutar"):
            try:

                new_excel = BytesIO()

                excel_result = load_and_execute_script(
                    "Facturacion ventas",
                    uploaded_files,
                    new_excel 
                )

                if excel_result is not None:
                    st.success("¡HECHO!")
                    st.download_button(
                        label="Descargar",
                        data=excel_result.getvalue(),
                        file_name=f"Ventas {fecha}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"Error al ejecutar el script: {str(e)}")

elif script_option == "Unnax CaixaBank":
    st.header("Archivos")
    uploaded_unnax = st.file_uploader("Unnax", type=["csv"])
    uploaded_compras = st.file_uploader("Compras Netsuit", type=["xlsx"])

    uploaded_files = {
    "Unnax": uploaded_unnax,
    "Compras": uploaded_compras,
    }

    if all(uploaded_files.values()):
        if st.button("Ejecutar"):
            try:

                new_excel = BytesIO()

                excel_result = load_and_execute_script(
                    "Unnax CB",
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
                        file_name=f"Carga Unnax CaixaBank {fecha}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"Error al ejecutar el script: {str(e)}")

elif script_option == "Unnax Easy Payment":
    st.header("Archivos")
    uploaded_unnax = st.file_uploader("Unnax", type=["csv"])
    uploaded_compras = st.file_uploader("Compras Netsuit", type=["xlsx"])

    uploaded_files = {
    "Unnax": uploaded_unnax,
    "Compras": uploaded_compras,
    }

    if all(uploaded_files.values()):
        if st.button("Ejecutar"):
            try:

                new_excel = BytesIO()

                excel_result = load_and_execute_script(
                    "Unnax EP",
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
                        file_name=f"Carga Unnax Easy Payment {fecha}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"Error al ejecutar el script: {str(e)}")
