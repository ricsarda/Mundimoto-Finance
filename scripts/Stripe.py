import pandas as pd
from io import BytesIO
from datetime import datetime

def main(files, pdfs=None, new_excel=None, month=None, year=None):

    try:

        if "Stripe" not in files or files["Stripe"] is None:
            raise RuntimeError("Falta el archivo CSV de Stripe (clave 'Stripe').")
        
        uploaded_file = files["Stripe"]
        
        stripe = pd.read_csv(uploaded_file, delimiter=',')
        # ---------------------------------------------------------------------
        stripe['automatic_payout_effective_at'] = pd.to_datetime(stripe['automatic_payout_effective_at']).dt.strftime('%d/%m/%Y')
        renting_blancks = stripe[stripe['payment_metadata[origin]'] != 'sales']
        renting_blancks = renting_blancks.groupby('automatic_payout_effective_at', as_index=False)[['gross', 'fee', 'net']].sum()
        renting_blancks = renting_blancks[['automatic_payout_effective_at', 'gross']]
        renting_blancks.rename(columns={'gross':'Credit'}, inplace=True)
        renting_blancks['Account ID'] = '1841'
        ventas = stripe[stripe['payment_metadata[origin]'] == 'sales']
        ventas = ventas.groupby('automatic_payout_effective_at', as_index=False)[['gross', 'fee', 'net']].sum()
        ventas = ventas[['automatic_payout_effective_at', 'gross']]
        ventas.rename(columns={'gross':'Credit'}, inplace=True)
        ventas['Account ID'] = '1866'
        stripegroup = stripe.groupby('automatic_payout_effective_at', as_index=False)[['gross', 'fee', 'net']].sum()
        fee = stripegroup[['automatic_payout_effective_at', 'fee']]
        fee.rename(columns={'fee':'Debit'}, inplace=True)
        fee['Account ID'] = '1821'
        net = stripegroup[['automatic_payout_effective_at', 'net']]
        net.rename(columns={'net':'Debit'}, inplace=True)
        net['Account ID'] = '2437'
        carga = pd.concat([renting_blancks, ventas, fee, net], axis=0)
        carga['Credit'] = carga['Credit'].astype(float)
        carga['Debit'] = carga['Debit'].astype(float)
        carga['ExternalID'] = carga.apply(lambda row: f"Stripe_{row['automatic_payout_effective_at']}",axis=1)
        carga['Memo'] = carga.apply(lambda row: f"Stripe Settlement {row['automatic_payout_effective_at']}",axis=1)
        carga['Descripcion linea'] = carga.apply(lambda row: f"Stripe Settlement {row['automatic_payout_effective_at']}",axis=1)
        carga.rename(columns={'automatic_payout_effective_at':'Fecha'}, inplace=True)
        carga['Clase'] = ''
        columnas_carga = ['ExternalID','Fecha', 'Memo', 'Account ID', 'Debit', 'Credit',  'Clase', 'Descripcion linea']
        carga = carga[columnas_carga]
        carga = carga.sort_values(by='ExternalID', ascending=True)


        output = BytesIO()
        carga.to_csv(output, sep=';', index=False, encoding='utf-8')
        output.seek(0)


        return output

    except Exception as e:

        raise RuntimeError(f"Error al procesar el archivo CSV de Stripe: {str(e)}")
