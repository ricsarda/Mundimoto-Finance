import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO

def process_stripe_data(uploaded_stripe):
        try:
        # Verificar que el archivo se ha subido correctamente
        if uploaded_file is None:
            st.error("No se ha subido ningún archivo.")
            return None

        # Intentar leer el CSV
        try:
            stripe = pd.read_csv(uploaded_stripe, delimiter=',')
        except Exception as e:
            st.error(f"Error al leer el archivo CSV: {str(e)}")
            return None

        # Verificar que contiene las columnas necesarias
        required_columns = {'automatic_payout_effective_at', 'payment_metadata[origin]', 'gross', 'fee', 'net'}
        if not required_columns.issubset(stripe.columns):
            st.error(f"El archivo CSV no contiene las columnas necesarias. Se esperaban: {required_columns}")
            return None

        # Convertir la fecha a formato adecuado
        stripe['automatic_payout_effective_at'] = pd.to_datetime(stripe['automatic_payout_effective_at'], errors='coerce').dt.strftime('%d/%m/%Y')

        # Procesamiento de datos
        renting_blancks = stripe[stripe['payment_metadata[origin]'] != 'sales']
        renting_blancks = renting_blancks.groupby('automatic_payout_effective_at', as_index=False)[['gross', 'fee', 'net']].sum()
        renting_blancks = renting_blancks[['automatic_payout_effective_at', 'gross']]
        renting_blancks.rename(columns={'gross': 'Credit'}, inplace=True)
        renting_blancks['Account'] = '1841'

        ventas = stripe[stripe['payment_metadata[origin]'] == 'sales']
        ventas = ventas.groupby('automatic_payout_effective_at', as_index=False)[['gross', 'fee', 'net']].sum()
        ventas = ventas[['automatic_payout_effective_at', 'gross']]
        ventas.rename(columns={'gross': 'Credit'}, inplace=True)
        ventas['Account'] = '1866'

        stripegroup = stripe.groupby('automatic_payout_effective_at', as_index=False)[['gross', 'fee', 'net']].sum()
        fee = stripegroup[['automatic_payout_effective_at', 'fee']]
        fee.rename(columns={'fee': 'Debit'}, inplace=True)
        fee['Account'] = '1821'

        net = stripegroup[['automatic_payout_effective_at', 'net']]
        net.rename(columns={'net': 'Debit'}, inplace=True)
        net['Account'] = '2437'

        carga = pd.concat([renting_blancks, ventas, fee, net], axis=0)
        carga['Credit'] = pd.to_numeric(carga['Credit'], errors='coerce')
        carga['Debit'] = pd.to_numeric(carga['Debit'], errors='coerce')

        carga['ExternalID'] = carga.apply(lambda row: f"Stripe_{row['automatic_payout_effective_at']}", axis=1)
        carga['Memo'] = carga.apply(lambda row: f"Liquidación Stripe {row['automatic_payout_effective_at']}", axis=1)
        carga['Descripción linea'] = carga.apply(lambda row: f"Liquidación Stripe {row['automatic_payout_effective_at']}", axis=1)
        carga.rename(columns={'automatic_payout_effective_at': 'Fecha'}, inplace=True)
        carga['Clase'] = ''

        columnas_carga = ['ExternalID', 'Fecha', 'Memo', 'Account', 'Debit', 'Credit', 'Clase', 'Descripción linea']
        carga = carga[columnas_carga]
        carga = carga.sort_values(by='Fecha', ascending=True)

        # Guardar archivo en memoria
        output = BytesIO()
        carga.to_csv(output, sep=';', index=False, encoding='utf-8')
        output.seek(0)

        return output

    except Exception as e:
        st.error(f"Error inesperado al procesar el archivo: {str(e)}")
        return None
