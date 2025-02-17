# scripts/stripe_data.py

import pandas as pd
from io import BytesIO
from datetime import datetime

def main(files):

    try:
        # 1) Recuperar el archivo 'Stripe' del diccionario files
        if "Stripe" not in files or files["Stripe"] is None:
            raise RuntimeError("Falta el archivo CSV de Stripe (clave 'Stripe').")
        
        # 'files["Stripe"]' es un UploadedFile, lo convertimos a un buffer
        uploaded_file = files["Stripe"]
        
        # 2) Leer el CSV
        stripe = pd.read_csv(uploaded_file, delimiter=',')

        # 3) Procesar los datos (misma lógica que tu script original)
        # ---------------------------------------------------------------------
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

        # Unir
        carga = pd.concat([renting_blancks, ventas, fee, net], axis=0)

        # Ajustar columnas
        carga['Credit'] = pd.to_numeric(carga['Credit'], errors='coerce')
        carga['Debit'] = pd.to_numeric(carga['Debit'], errors='coerce')
        carga['ExternalID'] = carga['automatic_payout_effective_at'].apply(lambda x: f"Stripe_{x}")
        carga['Memo'] = carga['automatic_payout_effective_at'].apply(lambda x: f"Liquidación Stripe {x}")
        carga['Descripción linea'] = carga['automatic_payout_effective_at'].apply(lambda x: f"Liquidación Stripe {x}")
        carga.rename(columns={'automatic_payout_effective_at': 'Fecha'}, inplace=True)
        carga['Clase'] = ''

        columnas_carga = ['ExternalID', 'Fecha', 'Memo', 'Account', 'Debit', 'Credit', 'Clase', 'Descripción linea']
        carga = carga[columnas_carga]
        carga = carga.sort_values(by='Fecha', ascending=True)

        # 4) Guardar en un CSV en memoria
        output = BytesIO()
        carga.to_csv(output, sep=';', index=False, encoding='utf-8')
        output.seek(0)

        # 5) Retornar el BytesIO
        return output

    except Exception as e:
        # Si algo falla, levantamos una RuntimeError para que Streamlit la capte
        raise RuntimeError(f"Error al procesar el archivo CSV de Stripe: {str(e)}")
