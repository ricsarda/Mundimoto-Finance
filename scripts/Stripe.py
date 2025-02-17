import pandas as pd
import streamlit as st
from io import BytesIO
from datetime import datetime

def process_stripe_data(uploaded_file):
    try:
        # Leer el archivo CSV cargado
        stripe = pd.read_csv(uploaded_file, delimiter=',')
        
        # Formatear fecha
        stripe['automatic_payout_effective_at'] = pd.to_datetime(stripe['automatic_payout_effective_at']).dt.strftime('%d/%m/%Y')

        # Filtrar datos de renting
        renting_blancks = stripe[stripe['payment_metadata[origin]'] != 'sales']
        renting_blancks = renting_blancks.groupby('automatic_payout_effective_at', as_index=False)[['gross', 'fee', 'net']].sum()
        renting_blancks = renting_blancks[['automatic_payout_effective_at', 'gross']]
        renting_blancks.rename(columns={'gross': 'Credit'}, inplace=True)
        renting_blancks['Account'] = '1841'

        # Filtrar datos de ventas
        ventas = stripe[stripe['payment_metadata[origin]'] == 'sales']
        ventas = ventas.groupby('automatic_payout_effective_at', as_index=False)[['gross', 'fee', 'net']].sum()
        ventas = ventas[['automatic_payout_effective_at', 'gross']]
        ventas.rename(columns={'gross': 'Credit'}, inplace=True)
        ventas['Account'] = '1866'

        # Calcular fees
        stripegroup = stripe.groupby('automatic_payout_effective_at', as_index=False)[['gross', 'fee', 'net']].sum()
        fee = stripegroup[['automatic_payout_effective_at', 'fee']]
        fee.rename(columns={'fee': 'Debit'}, inplace=True)
        fee['Account'] = '1821'

        # Calcular neto
        net = stripegroup[['automatic_payout_effective_at', 'net']]
        net.rename(columns={'net': 'Debit'}, inplace=True)
        net['Account'] = '2437'

        # Concatenar todos los datos
        carga = pd.concat([renting_blancks, ventas, fee, net], axis=0)

        # Ajustar columnas y valores
        carga['Credit'] = carga['Credit'].astype(float)
        carga['Debit'] = carga['Debit'].astype(float)
        carga['ExternalID'] = carga.apply(lambda row: f"Stripe_{row['automatic_payout_effective_at']}", axis=1)
        carga['Memo'] = carga.apply(lambda row: f"Liquidación Stripe {row['automatic_payout_effective_at']}", axis=1)
        carga['Descripción linea'] = carga.apply(lambda row: f"Liquidación Stripe {row['automatic_payout_effective_at']}", axis=1)
        carga.rename(columns={'automatic_payout_effective_at': 'Fecha'}, inplace=True)
        carga['Clase'] = ''

        columnas_carga = ['ExternalID', 'Fecha', 'Memo', 'Account', 'Debit', 'Credit', 'Clase', 'Descripción linea']
        carga = carga[columnas_carga]
        carga = carga.sort_values(by='Fecha', ascending=True)

        # Guardar en memoria como CSV
        output = BytesIO()
        carga.to_csv(output, sep=';', index=False, encoding='utf-8')
        output.seek(0)

        return output

    except Exception as e:
        st.error(f"Error al procesar el archivo: {str(e)}")
        return None
