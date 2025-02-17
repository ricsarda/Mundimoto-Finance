import pandas as pd
import streamlit as st
from io import BytesIO
from datetime import datetime

# Función para procesar los datos
def process_facturacion(clienti_file, sales_file, metabase_file, location_file):
    try:
        # Leer archivos subidos
        Sales = pd.read_excel(sales_file)
        metabase = pd.read_excel(metabase_file)
        locationIT = pd.read_csv(location_file, delimiter=';')

        # Procesamiento de datos
        clienti = Sales.copy()
        clienti['externalId'] = clienti['CF']
        clienti['companyName'] = clienti['CLIENTE']
        
        def check_length(x):
            return '' if len(x) > 13 else x
        
        clienti['vatregnumber'] = clienti['CF'].apply(check_length)
        clienti['[NExIL] Fiscal Code'] = clienti['CF']
        clienti['email'] = clienti['E-MAIL']

        def extract_via_cap(residenza):
            parts = residenza.split(',')
            if len(parts) < 3:
                via = residenza
                cap = 'Review'
            else:
                via = ','.join(parts[:2]) + ','
                cap = parts[2].strip()[:6]
            return via, cap

        clienti['ADDRESS'], clienti['Zip Code'] = zip(*clienti['RESIDENZA'].apply(extract_via_cap))
        clienti['ADDRESS'] = clienti['ADDRESS'].str.rstrip(',')
        clienti['Zip Code'] = clienti['Zip Code'].str.replace(' ', '')

        # Mapear ubicación
        cap_to_provincia = dict(zip(locationIT['cap'], locationIT['sigla_provincia']))
        clienti['Provincia'] = clienti['Zip Code'].map(cap_to_provincia)
        cap_to_citta = dict(zip(locationIT['cap'], locationIT['denominazione_ita_altra']))
        clienti['Ciudad'] = clienti['Zip Code'].map(cap_to_citta)
        clienti['Country'] = 'Italy'
        clienti['[NEXIL] ADDRESSEE PEC'] = clienti['E-MAIL']
        clienti['[NEXIL] CODICE DESTINATARIO PR'] = '0000000'

        Columnas_clienti = ['externalId','companyName', 'vatregnumber','[NExIL] Fiscal Code', 'email',
                            'ADDRESS','Ciudad','Provincia','Zip Code', 'Country',  
                            '[NEXIL] ADDRESSEE PEC', '[NEXIL] CODICE DESTINATARIO PR']
        clienti = clienti[Columnas_clienti]

        # Crear archivo CSV en memoria
        clienti_output = BytesIO()
        clienti.to_csv(clienti_output, sep=';', index=False, encoding='utf-8')
        clienti_output.seek(0)

        # Procesamiento de órdenes
        ordini = Sales.copy()
        ordini = ordini.merge(metabase[['license_plate', 'frame_number', 'brand', 'model']], 
                              left_on='TARGA', right_on='license_plate', how='left') 

        ordini['External ID'] = ordini['TARGA']
        ordini['Cliente'] = ordini['CF']
        ordini['Date'] = pd.to_datetime(ordini['PAYMENT DATE']).dt.strftime('%d/%m/%Y')
        ordini['Location'] = '13'
        ordini['itemLine_quantity'] = '1'
        ordini['Vendita moto - plate - vin - marca - modelo'] = ordini.apply(
            lambda row: f'Vendita moto - {row["TARGA"]} - {row["frame_number"]} - {row["brand"]} - {row["model"]}', axis=1)

        Columnas_ordini = ['External ID','Cliente', 'Date', 'itemLine_item', 'itemLine_salesPrice', 'Vendita moto - plate - vin - marca - modelo']
        
        # Generar archivo CSV de órdenes
        ordini_output = BytesIO()
        ordini.to_csv(ordini_output, sep=';', index=False, encoding='utf-8')
        ordini_output.seek(0)

        return clienti_output, ordini_output

    except Exception as e:
        st.error(f"Error al procesar los archivos: {str(e)}")
        return None, None

# --- INTERFAZ EN STREAMLIT ---
st.header("Carga de Archivos - Facturación IT")

# Cargar archivos
clienti_file = st.file_uploader("Subir archivo de Clientes (SalesIT.xlsx)", type=["xlsx"])
sales_file = st.file_uploader("Subir archivo de Ventas (Sales.xlsx)", type=["xlsx"])
metabase_file = st.file_uploader("Subir archivo de Metabase (metabase.xlsx)", type=["xlsx"])
location_file = st.file_uploader("Subir archivo de Ubicaciones (locationIT.csv)", type=["csv"])

if all([clienti_file, sales_file, metabase_file, location_file]):
    if st.button("Ejecutar Facturación"):
        clienti_output, ordini_output = process_facturacion(clienti_file, sales_file, metabase_file, location_file)

        if clienti_output and ordini_output:
            st.success("¡Procesamiento completado!")

            st.download_button(
                label="Descargar Clientes",
                data=clienti_output.getvalue(),
                file_name="MM IT - Importazione clienti.csv",
                mime="text/csv"
            )

            st.download_button(
                label="Descargar Órdenes",
                data=ordini_output.getvalue(),
                file_name="MM IT - Importazione ordini.csv",
                mime="text/csv"
            )
