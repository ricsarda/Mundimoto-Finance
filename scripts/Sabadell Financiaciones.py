import pandas as pd
from io import BytesIO

def main(files, pdfs, new_excel, month=None, year=None):
    try:
        # Leer archivos recibidos
        FinanciacionesSab = pd.read_excel(files["Financiaciones"], header=11)
        invoice = pd.read_csv(files["InvoicesItem"])
        invoices = pd.read_csv(files["Invoices"], sep=',')

        invoices = invoices[invoices['Status'] != 'Paid In Full']
        invoices1 = invoices.drop_duplicates(subset=['Tax Number'], keep='first')
        invoices2 = invoices[~invoices['Internal ID'].isin(invoices1['Internal ID'])]

        check = FinanciacionesSab.copy()

        # Lógica de Financiaciones Sabadell
        FinanciacionesSab['Date'] = FinanciacionesSab['Fecha Resolución']
        FinanciacionesSab['Importe'] = FinanciacionesSab['Cantidad A Financiar'].astype(str)
        FinanciacionesSab['Importe'] = FinanciacionesSab['Importe'].str.replace('.', '')
        FinanciacionesSab['Importe'] = FinanciacionesSab['Importe'].str.replace(',', '.')
        FinanciacionesSab['Importe'] = FinanciacionesSab['Importe'].astype(float)
        FinanciacionesSab['Importe'] = pd.to_numeric(FinanciacionesSab['Importe'], errors='coerce')
        FinanciacionesSab['Cliente_external ID'] = FinanciacionesSab['NIF Cliente']

        pago = FinanciacionesSab.merge(
            invoices1[['Tax Number','Amount (Gross)','Internal ID']],
            right_on='Tax Number',
            left_on='Cliente_external ID',
            how='left'
        )
        pago['Primera factura'] = -pago['Amount (Gross)'] + pago['Importe']

        def importecorrecto(row):
            if row['Primera factura'] > 0:
                return row['Amount (Gross)']
            elif row['Primera factura'] == 0:
                return row['Importe']
            else:
                return row['Importe']

        pago['Factura 1'] = pago.apply(importecorrecto, axis=1)
        pago = pago.merge(
            invoices2[['Tax Number','Amount (Gross)','Internal ID']],
            right_on='Tax Number',
            left_on='Cliente_external ID',
            how='left'
        )
        pago['Segunda factura'] = -pago['Amount (Gross)_y'] + pago['Primera factura']

        def importecorrecto2(row):
            if row['Primera factura'] < 0:
                return 0
            elif row['Primera factura'] == 0:
                return 0
            elif row['Segunda factura'] < 0:
                return row['Primera factura']
            elif row['Segunda factura'] == 0:
                return row['Primera factura']
            else:
                return row['Amount (Gross)_y']

        pago['Factura 2'] = pago.apply(importecorrecto2, axis=1)

        factura1 = pago[["Date","Cliente_external ID", "Tax Number_x" , "Amount (Gross)_x",  "Internal ID_x" , "Factura 1"]]
        factura2 = pago[["Date","Cliente_external ID", "Tax Number_y" , "Amount (Gross)_y",  "Internal ID_y" , "Factura 2"]]
        factura1 = factura1.rename(columns={"Tax Number_x": "Tax Number", "Amount (Gross)_x": "Amount (Gross)", "Internal ID_x": "Factura_INTERNAL ID","Factura 1": "Importe"} )
        factura2 = factura2.rename(columns={"Tax Number_y": "Tax Number", "Amount (Gross)_y": "Amount (Gross)", "Internal ID_y": "Factura_INTERNAL ID","Factura 2": "Importe"} )
        pago = pd.concat([factura1, factura2], ignore_index=True)
        pago = pago.dropna(subset=['Importe'])
        pago = pago[pago['Importe'] != 0]
        pago['External ID'] = pago.apply(
            lambda x: f'{int(x["Factura_INTERNAL ID"])}_PAY' if pd.notna(x["Factura_INTERNAL ID"]) else 'NaN_PAY', axis=1
        )
        pago['Cuenta Banco_EXTERNAL ID'] = 572000002
        pago = pago.drop(columns=['Tax Number','Amount (Gross)'])

        # Generar Excel en memoria
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            pago.to_excel(writer, sheet_name='Pago', index=False)
            check.to_excel(writer, sheet_name='Check', index=False)
            invoice.to_excel(writer, sheet_name='Item Internal ID', index=False)
            invoices.to_excel(writer, sheet_name='Invoices', index=False)
        output.seek(0)
        return output

    except Exception as e:
        raise RuntimeError(f"Error al procesar el script: {str(e)}")
