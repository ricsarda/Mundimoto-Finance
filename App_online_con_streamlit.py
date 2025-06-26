from io import BytesIO
import streamlit as st
import pandas as pd
import importlib.util
import sys
import os
import zipfile
import re
from datetime import datetime

fecha_actual = datetime.now()
fecha = fecha_actual.strftime("%d-%m-%Y")

# Configuración inicial de la app
st.title("Mundimoto Finance")
st.sidebar.header("Configuration")


# Diferenciación entre Italia y España
pais = st.sidebar.radio("Country", ("Spain", "Italy"))

# Opciones específicas para cada país
if pais == "Spain":
    script_options = [
        "Credit Stock", "Calculadora Precios B2C",  "Daily Report", "DNI y Matrícula" ,"Financiaciones Santander",
        "Financiaciones Renting", "Performance Comerciales B2C", "Unnax CaixaBank",
        "Unnax Easy Payment", "Stripe"
    ]
elif pais == "Italy":
    script_options = [
        "Purchases","Sales" 
    ]

# Selección del script
script_option = st.sidebar.selectbox("Execute:", script_options)

st.write(f"{script_option}")

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
            
        # Asegúrate de que processed_pdfs tenga un valor predeterminado si no se pasa como parámetro
        if pdfs is None:
            processed_pdfs = [] 
        else:
            processed_pdfs = pdfs

        # Llamar a la función principal del script con los parámetros adicionales
        result = module.main(processed_files, processed_pdfs, new_excel , month, year)
        if isinstance(result, tuple):
            return result  # Se devuelve una tupla con múltiples archivos

        return result  # Caso normal (un solo archivo)

        
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
    uploaded_FT = st.file_uploader("FT", type=["xlsx"])
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
                    st.success("¡GAS!")
                    st.download_button(
                        label="Descargar",
                        data=excel_result.getvalue(),
                        file_name=f"Credit Stock {fecha}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"Error, Contacta con Ric {str(e)}")

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
                    st.success("¡GAS!")
                    st.download_button(
                        label="Descargar",
                        data=excel_result.getvalue(),
                        file_name=f"Performance Comerciales B2C {fecha}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"Error, Contacta con Ric {str(e)}")


elif script_option == "Financiaciones Renting":
    st.header("PDFs")

    uploaded_pdfs = st.file_uploader(
        "Sube los PDFs",
        type=["pdf"],
        accept_multiple_files=True
    )

    if uploaded_pdfs:
        if st.button("Ejecutar"):
            try:
                new_excel = BytesIO()

                # Renombra explícitamente los archivos eliminando caracteres conflictivos
                pdfs_dict = {}
                for f in uploaded_pdfs:
                    sanitized_name = re.sub(r'[^a-zA-Z0-9_.-]', '_', f.name)
                    pdfs_dict[sanitized_name] = BytesIO(f.getvalue())

                excel_result = load_and_execute_script(
                    "Financiaciones Renting",
                    files={},                    
                    pdfs=pdfs_dict,
                    new_excel=new_excel
                )

                if excel_result is not None:
                    st.success("¡GAS!")
                    st.download_button(
                        label="Descargar",
                        data=excel_result.getvalue(),
                        file_name=f"Financiaciones_Renting_{fecha}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"Error, Contacta con Ric {str(e)}")



elif script_option == "Financiaciones Santander":
    st.header("Subida de archivos")

    uploaded_pdfs = st.file_uploader("Sube PDFs Santander cartas de Pago", type=["pdf"], accept_multiple_files=True)
    uploaded_financiaciones = st.file_uploader("Sube Excel Financiaciones", type=["xlsx"])
    uploaded_invoices = st.file_uploader("Sube csv MM - Item Internal ID Transactions: Results", type=["csv"])
    
    uploaded_files = {
        "Financiaciones": uploaded_financiaciones,
        "Invoices": uploaded_invoices
    }
    
    if all(uploaded_files.values()) and uploaded_pdfs:
        if st.button("Ejecutar"):
            try:
                pdfs_dict = {f.name: f for f in uploaded_pdfs}
                # Llamamos al script
                resultados = load_and_execute_script(
                    "Financiaciones Santander",
                    files=uploaded_files,
                    pdfs=pdfs_dict
                )
                if resultados is not None:
                    # 'resultados' es una tupla (excel_final_ops, excel_rest)
                    excel_ops, excel_otros = resultados

                    st.success("¡GAS!")
                    # Botón para descargar final_operaciones
                    st.download_button(
                        label="Descargar Comisiones",
                        data=excel_ops.getvalue(),
                        file_name=f"Financiaciones Santander-Comisiones {fecha}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    # Botón para descargar el "resto"
                    st.download_button(
                        label="Descargar Pagos",
                        data=excel_otros.getvalue(),
                        file_name=f"Financiaciones Santander-Pagos {fecha}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"Error, Contacta con Ric: {str(e)}")


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
                    st.success("¡GAS!")
                    st.download_button(
                        label="Descargar",
                        data=excel_result.getvalue(),
                        file_name=f"Ventas {fecha}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"Error, Contacta con Ric {str(e)}")

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
                    st.success("¡GAS!")
                    st.download_button(
                        label="Descargar",
                        data=excel_result.getvalue(),
                        file_name=f"Carga Unnax CaixaBank {fecha}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"Error, Contacta con Ric {str(e)}")

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
                    st.success("¡GAS!")
                    st.download_button(
                        label="Descargar",
                        data=excel_result.getvalue(),
                        file_name=f"Carga Unnax Easy Payment {fecha}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"Error, Contacta con Ric {str(e)}")

elif script_option == "Calculadora Precios B2C":
    st.header("Calculadora")
    # Cargar datos desde el CSV en el repositorio
    try:
        from scripts.calculadora_motos import load_data, calculate_price
        data = load_data("Motos para calcular.csv")
    except Exception as e:
        st.error(f"Error al cargar los datos: {str(e)}")
        data = None

    if data is not None:
        # Entrada del usuario
        marca = st.selectbox("Selecciona la marca", options=data['MARCA'].unique())
        if marca:
            modelos_disponibles = data[data['MARCA'] == marca]['MODELO'].unique()
            modelo = st.selectbox("Selecciona el modelo", options=modelos_disponibles)

        año = st.number_input("Introduce el año", min_value=int(data['Año'].min()), max_value=int(data['Año'].max()), value=int(data['Año'].mean()))
        km = st.number_input("Introduce el kilometraje", min_value=0, value=int(data['KM'].median()))

        # Botón para calcular
        if st.button("Calcular precio"):
            if not modelo or not marca:
                st.error("Por favor selecciona una marca y un modelo válidos.")
            else:
                # Calcular precio
                precio, variacion, num_motos, min_año, max_km, subset = calculate_price(data, marca, modelo, año, km)
                if precio is None:
                    st.error("No se encontraron datos suficientes para calcular el precio.")
                else:
                    st.success(f"Precio estimado: {precio:,.2f} €")
                    st.write(f"Variación estimada: +/- {variacion:,.2f} €")
                    if min_año is not None and not pd.isna(min_año):
                        st.write(f"Mayor antigüedad encontrada: {int(min_año)}")
                    else:
                        st.write("Año no encontrado")
                    if max_km is not None and not pd.isna(max_km):
                        st.write(f"Mayor kilometraje encontrado: {int(max_km)} KM")
                    else:
                        st.write("KM no encontrado")
                    st.write(f"Número de unidades: {num_motos}")
                    st.write(f"Histórico {marca} {modelo}:")
                    st.dataframe(subset)

elif script_option == "Stripe":
    st.header("Subir csv")

    # Subida de un único archivo CSV
    uploaded_stripe = st.file_uploader("Archivo Conciliacin_detallada_de_transferencias", type=["csv"])

    # Construimos el diccionario con clave "Stripe"
    uploaded_files = {
        "Stripe": uploaded_stripe
    }

    # Verificamos si el usuario subió algo
    if uploaded_files["Stripe"] is not None:
        if st.button("Ejecutar"):
            try:
                # Llamamos a la función load_and_execute_script
                result = load_and_execute_script(
                    "Stripe",         # el nombre del script: stripe_data.py
                    files=uploaded_files   # pasamos el dict con "Stripe"
                )

                # 'result' será un BytesIO con el CSV final
                if result is not None:
                    st.success("¡GAS!")
                    st.download_button(
                        label="Descargar",
                        data=result.getvalue(),
                        file_name=f"Stripe_{fecha}.csv",
                        mime="text/csv"
                    )
            except Exception as e:
                st.error(f"Error al procesar CSV de Stripe, Contacta con Ric: {str(e)}")

elif script_option == "Purchases":
    st.header("File")

    # Subida del archivo requerido
    uploaded_purchases = st.file_uploader("Upload Purchases", type=["xlsx"])

    uploaded_files = {
        "PurchasesIT": uploaded_purchases
    }

    if uploaded_purchases:
        if st.button("Execute"):
            try:
                result_item, result_fornitore, result_purchase = load_and_execute_script(
                    "Purchases",
                    uploaded_files
                )

                if result_item is None or result_fornitore is None or result_purchase is None:
                    st.error("File upload error, contact with: ricardo.sarda@mundimoto.com")
                else:
                    st.success("¡GAS!")
                    # Guardar los archivos en `st.session_state`
                    st.session_state["PurchasesIT_Item"] = result_item
                    st.session_state["PurchasesIT_Fornitore"] = result_fornitore
                    st.session_state["PurchasesIT_Purchase"] = result_purchase

            except Exception as e:
                st.error(f"Error, contact with Ricardo Sarda via Slack or e-mail: ricardo.sarda@mundimoto.com -{str(e)}")
    # Crear un ZIP con los tres archivos si están disponibles en `st.session_state`
    if "PurchasesIT_Item" in st.session_state and "PurchasesIT_Fornitore" in st.session_state and "PurchasesIT_Purchase" in st.session_state:
        zip_buffer = BytesIO()

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
            zipf.writestr(f"Purchases_IT_Item_{fecha}.xlsx", st.session_state["PurchasesIT_Item"].getvalue())
            zipf.writestr(f"Purchases_IT_Fornitore_{fecha}.xlsx", st.session_state["PurchasesIT_Fornitore"].getvalue())
            zipf.writestr(f"Purchases_IT_Purchase_{fecha}.xlsx", st.session_state["PurchasesIT_Purchase"].getvalue())

        zip_buffer.seek(0)

        st.download_button(
            label="Download (ZIP)",
            data=zip_buffer.getvalue(),
            file_name=f"Purchases_{fecha}.zip",
            mime="application/zip"
        )

elif script_option == "DNI y Matrícula":
    uploaded_file = st.file_uploader("Sube el extracto de Santander", type=["xlsx"], key="santander")
    files = {"Extracto de Santander": uploaded_file}

    if uploaded_file is not None:
        st.success("¡GAS!")

        # Ejecutar el script
        df_resultado = load_and_execute_script(script_option, files)

        if df_resultado is not None:

            buffer = BytesIO()
            df_resultado.to_excel(buffer, index=False, engine='openpyxl')
            buffer.seek(0)

            # Botón de descarga
            st.download_button(
                label="Download",
                data=buffer,
                file_name=f"DNI_Matricula_{fecha_actual}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
elif script_option == "Sales":
    st.header("Files")
    # Subida de archivos requeridos
    uploaded_sales = st.file_uploader("Upload Sales", type=["xlsx"])
    uploaded_metabase = st.file_uploader("Upload Metabase: https://mundimoto.metabaseapp.com/dashboard/432-raw-data-purchases-stock?productive_status=", type=["xlsx"])

    uploaded_files = {
        "Sales": uploaded_sales,
        "Metabase": uploaded_metabase
    }

    if all(uploaded_files.values()) and st.button("Execute"):
        try:
            result_clienti, result_ordini = load_and_execute_script(
                "Sales",
                uploaded_files
            )

            if result_clienti is None or result_ordini is None:
                st.error("Error processing Sales. Check the input files.")
            else:
                st.success("¡GAS!")

                # Guardar los archivos en `st.session_state`
                st.session_state["Sales_Clienti"] = result_clienti
                st.session_state["Sales_Ordini"] = result_ordini

        except Exception as e:
            st.error(f"Error, contact Ricardo Sarda via Slack or e-mail: ricardo.sarda@mundimoto.com {str(e)}")

    # Crear un ZIP con los dos archivos si están disponibles en `st.session_state`
    if "Sales_Clienti" in st.session_state and "Sales_Ordini" in st.session_state:
        zip_buffer = BytesIO()

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
            zipf.writestr(f"MM IT - Importazione clienti {fecha}.csv", st.session_state["Sales_Clienti"].getvalue())
            zipf.writestr(f"MM IT - Importazione ordini di {fecha}.csv", st.session_state["Sales_Ordini"].getvalue())

        zip_buffer.seek(0)

        st.download_button(
            label="Download all files (ZIP)",
            data=zip_buffer.getvalue(),
            file_name=f"Sales_{fecha}.zip",
            mime="application/zip"
        )
