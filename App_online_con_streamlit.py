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


st.title("Mundimoto Finance")
st.sidebar.header("Configuration")

pais = st.sidebar.radio("Country", ("Spain", "Italy"))


if pais == "Spain":
    script_options = [
        "Calculadora Precios B2C", "Credit Stock", "DNI y Matrícula" , "Facilitea", "Financiaciones Renting","Revisión Pricing Web", "Sabadell Financiaciones",
        "Santander Financiaciones", "Sofinco Financiaciones", "Stripe"
    ]
elif pais == "Italy":
    script_options = [
        "Purchases","Sales" 
    ]

script_option = st.sidebar.selectbox("Execute:", script_options)

st.write(f"{script_option}")

def load_and_execute_script(script_name, files, pdfs=None, new_excel=None, month=None, year=None):
    try:
        script_path = os.path.join("scripts", f"{script_name}.py")
        if not os.path.exists(script_path):
            raise FileNotFoundError(f"El script {script_name} no fue encontrado en {script_path}")
        
        spec = importlib.util.spec_from_file_location(script_name, script_path)
        module = importlib.util.module_from_spec(spec)
        sys.modules[script_name] = module
        spec.loader.exec_module(module)

        processed_files = {}
        for key, file in files.items():
            buffer = BytesIO(file.read()) 
            buffer.seek(0)  
            processed_files[key] = buffer
            
        if pdfs is None:
            processed_pdfs = [] 
        else:
            processed_pdfs = pdfs


        result = module.main(processed_files, processed_pdfs, new_excel , month, year)
        if isinstance(result, tuple):
            return result  

        return result  

        
    except FileNotFoundError as e:
        st.error(f"Error de archivo: {str(e)}")
    except AttributeError as e:
        st.error(f"Error en el script {script_name}: {str(e)}")
    except KeyError as e:
        st.error(f"Error con los archivos subidos: falta el archivo clave {str(e)}.")
    except Exception as e:
        st.error(f"Error inesperado al ejecutar el script {script_name}: {str(e)}")

if script_option == "Daily Report":
    st.header("Archivos")

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
    uploaded_metabase = st.file_uploader("Metabase-[Raw data purchases - stock]", type=["xlsx"])
    uploaded_santander = st.file_uploader("Santander-[Consulta de Documentaciones]", type=["xlsx"])
    uploaded_sabadell = st.file_uploader("Sabadell-[MainServlet]", type=["xls"])
    uploaded_sofinco = st.file_uploader("Sofinco-[ExportListDraws]", type=["xlsx"])
    
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
                        label="Download",
                        data=excel_result.getvalue(),
                        file_name=f"Credit Stock {fecha}.xlsx",
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
                        label="Download",
                        data=excel_result.getvalue(),
                        file_name=f"Financiaciones_Renting_{fecha}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"Error, Contacta con Ric {str(e)}")


elif script_option == "Sofinco Financiaciones":
    st.header("Subida de archivos")
    
    uploaded_pdfs = st.file_uploader("PDFs de Sofinco, (financiaciones@mundimoto.com <Resumen de operaciones>)", type=["pdf"], accept_multiple_files=True)
    uploaded_invoices = st.file_uploader("Sube csv MM - Item Internal ID Transactions: Results", type=["csv"])
    uploaded_invoice = st.file_uploader("Sube csv de Invoices Marc", type=["csv"])

    uploaded_files = {
        "Invoices": uploaded_invoices,
        "invoice": uploaded_invoice
    }
    
    if all(uploaded_files.values()) and uploaded_pdfs:
        if st.button("Ejecutar"):
            try:
                pdfs_dict = {f.name: f for f in uploaded_pdfs}
                # Llamamos al script
                resultados = load_and_execute_script(
                    "Sofinco Financiaciones",
                    files=uploaded_files,
                    pdfs=pdfs_dict
                )
                if resultados is not None:

                    output_ops, output_rest = resultados

                    st.success("¡GAS!")

                    st.download_button(
                        label="Download Comisiones",
                        data=output_ops.getvalue(),
                        file_name=f"Sofinco Financiaciones-Comisiones {fecha}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                    st.download_button(
                        label="Download Pagos",
                        data=output_rest.getvalue(),
                        file_name=f"Sofinco Financiaciones-Pagos {fecha}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"Error, Contacta con Ric: {str(e)}")


elif script_option == "Santander Financiaciones":
    st.header("Subida de archivos")

    uploaded_pdfs = st.file_uploader("Sube PDFs Santander cartas de Pago", type=["pdf"], accept_multiple_files=True)
    uploaded_financiaciones = st.file_uploader("Sube csv de Financiaciones", type=["csv"])
    uploaded_invoices = st.file_uploader("Sube csv MM - Item Internal ID Transactions: Results", type=["csv"])
    uploaded_invoice = st.file_uploader("Sube csv de Invoices Marc", type=["csv"])
    
    uploaded_files = {
        "Financiaciones": uploaded_financiaciones,
        "Invoices": uploaded_invoices,
        "invoice": uploaded_invoice
    }
    
    if all(uploaded_files.values()) and uploaded_pdfs:
        if st.button("Ejecutar"):
            try:
                pdfs_dict = {f.name: f for f in uploaded_pdfs}
                # Llamamos al script
                resultados = load_and_execute_script(
                    "Santander Financiaciones",
                    files=uploaded_files,
                    pdfs=pdfs_dict
                )
                if resultados is not None:
  
                    excel_ops, excel_otros = resultados

                    st.success("¡GAS!")

                    st.download_button(
                        label="Download Comisiones",
                        data=excel_ops.getvalue(),
                        file_name=f"Santander Financiaciones-Comisiones {fecha}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                    st.download_button(
                        label="Download Pagos",
                        data=excel_otros.getvalue(),
                        file_name=f"Santander Financiaciones-Pagos {fecha}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"Error, Contacta con Ric: {str(e)}")

elif script_option == "Sabadell Financiaciones":
    st.header("Subida de archivos")

    uploaded_financiaciones = st.file_uploader("Sube Excel de Financiaciones de Sabadell", type=["xlsx"])
    uploaded_invoices = st.file_uploader("Sube csv MM - Item Internal ID Transactions: Results", type=["csv"])
    uploaded_invoice = st.file_uploader("Sube csv de Invoices Marc", type=["csv"])
    
    uploaded_files = {
        "Financiaciones": uploaded_financiaciones,
        "Invoices": uploaded_invoices,
        "invoice": uploaded_invoice
    }
    
    if all(uploaded_files.values()):
        if st.button("Ejecutar"):
            try:
                resultados = load_and_execute_script(
                    "Sabadell Financiaciones",
                    files=uploaded_files,
                )
                if resultados is not None:

                    output = resultados

                    st.success("¡GAS!")

                    st.download_button(
                        label="Download Pagos",
                        data=output.getvalue(),
                        file_name=f"Sabadell-Pago {fecha}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

            except Exception as e:
                st.error(f"Error, Contacta con Ric: {str(e)}")

elif script_option == "Calculadora Precios B2C":
    st.header("Calculadora")

    try:
        from scripts.calculadora_motos import load_data, calculate_price
        data = load_data("Motos para calcular.csv")
    except Exception as e:
        st.error(f"Error al cargar los datos: {str(e)}")
        data = None

    if data is not None:

        marca = st.selectbox("Marca", options=data['MARCA'].unique())
        if marca:
            modelos_disponibles = data[data['MARCA'] == marca]['MODELO'].unique()
            modelo = st.selectbox("Modelo", options=modelos_disponibles)

        año = st.number_input("Año", min_value=int(data['Año'].min()), max_value=int(data['Año'].max()), value=int(data['Año'].mean()))
        km = st.number_input("KM", min_value=0, value=int(data['KM'].median()))

        if st.button("Calcular precio"):
            if not modelo or not marca:
                st.error("Por favor selecciona una marca y un modelo válidos.")
            else:

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
    st.header("Conciliacin_pormenorizada(...).csv")

    uploaded_stripe = st.file_uploader("#stripe ([Informes],[Conciliación de transferencias] → En la fecha poner el mes a cerrar y seleccionar opción [download] en [Conciliación de transferencias] y seleccionar [Pormenorizado] y [Todas las columnas])", type=["csv"])

    uploaded_files = {
        "Stripe": uploaded_stripe
    }


    if uploaded_files["Stripe"] is not None:
        if st.button("Ejecutar"):
            try:
               
                result = load_and_execute_script(
                    "Stripe",         
                    files=uploaded_files  
                )

    
                if result is not None:
                    st.success("¡GAS!")
                    st.download_button(
                        label="Download",
                        data=result.getvalue(),
                        file_name=f"Stripe_{fecha}.csv",
                        mime="text/csv"
                    )
            except Exception as e:
                st.error(f"Error al procesar CSV de Stripe, Contacta con Ric: {str(e)}")

elif script_option == "Purchases":
    st.header("File")

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

                    st.session_state["PurchasesIT_Item"] = result_item
                    st.session_state["PurchasesIT_Fornitore"] = result_fornitore
                    st.session_state["PurchasesIT_Purchase"] = result_purchase

            except Exception as e:
                st.error(f"Error, contact with Ricardo Sarda via Slack or e-mail: ricardo.sarda@mundimoto.com -{str(e)}")
   
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

elif script_option == "Facilitea":
    st.header("Archivos necesarios")

    uploaded_zip = st.file_uploader("ZIP de Excels Facilitea (administracion@ajmotor.com <Liquidación Automática>)", type="zip")
    uploaded_invoices = st.file_uploader("Sube csv MM - Item Internal ID Transactions: Results", type=["csv"])
    uploaded_invoice = st.file_uploader("Sube csv de Invoices Marc", type=["csv"])

    uploaded_files = {
        "FaciliteaZIP": uploaded_zip,
        "InvoicesItem": uploaded_invoices,
        "Invoices": uploaded_invoice,
    }


    if all(uploaded_files.values()):
        if st.button("Ejecutar"):
            files = {k: BytesIO(v.read()) for k, v in uploaded_files.items()}
            result = load_and_execute_script("Facilitea", files)
            if result:
                st.download_button("Download", data=result, file_name=f"Facilitea {fecha}.xlsx")

elif script_option == "DNI y Matrícula":
    uploaded_file = st.file_uploader("Sube el extracto de Santander", type=["xlsx"], key="santander")
    files = {"Extracto de Santander": uploaded_file}

    if uploaded_file is not None:
        st.success("¡GAS!")

        df_resultado = load_and_execute_script(script_option, files)

        if df_resultado is not None:

            buffer = BytesIO()
            df_resultado.to_excel(buffer, index=False, engine='openpyxl')
            buffer.seek(0)


            st.download_button(
                label="Download",
                data=buffer,
                file_name=f"DNI_Matricula_{fecha_actual}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
elif script_option == "Revisión Pricing Web":
    st.header("Archivos necesarios")
    st.info("Solamente se sumarán los días que llevemos este mes al LEAD TIME. "
               "Si el cálculo es anterior no se tendrá en cuenta.")
    uploaded_retool = st.file_uploader("Retool CSV (table-data*.csv)", type=["csv"])
    uploaded_leadtime = st.file_uploader("Leadtime Excel", type=["xls", "xlsx"])

    uploaded_files = {
        "RetoolCSV": uploaded_retool,
        "LeadtimeExcel": uploaded_leadtime,
    }

    if all(uploaded_files.values()):
        if st.button("Ejecutar"):
            files = {k: BytesIO(v.read()) for k, v in uploaded_files.items()}
            result = load_and_execute_script("Revisión Pricing Web", files)
            if result:
                st.download_button(
                    "Download",
                    data=result,
                    file_name="Revisión Pricing Web.xlsx"
                )


elif script_option == "Sales":
    st.header("Files")

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


                st.session_state["Sales_Clienti"] = result_clienti
                st.session_state["Sales_Ordini"] = result_ordini

        except Exception as e:
            st.error(f"Error, contact Ricardo Sarda via Slack or e-mail: ricardo.sarda@mundimoto.com {str(e)}")

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





