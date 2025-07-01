import zipfile
import pandas as pd
import io
from io import BytesIO

def main(files, pdfs, new_excel, month=None, year=None):
    try:
        # 1. Abrir el ZIP subido y extraer todos los excels .xls
        zip_buffer = files["FaciliteaZIP"]
        zip_buffer.seek(0)
        dataframes = []
        with zipfile.ZipFile(zip_buffer, 'r') as archivo_zip:
            archivos_xlsx = [nombre for nombre in archivo_zip.namelist() if nombre.endswith('.xls')]
            for nombre in archivos_xlsx:
                with archivo_zip.open(nombre) as file:
                    df = pd.read_excel(io.BytesIO(file.read()))
                    df['archivo_origen'] = nombre
                    dataframes.append(df)
        if dataframes:
            df_combined = pd.concat(dataframes, ignore_index=True)
        else:
            df_combined = pd.DataFrame()

        # 2. Limpieza y cruce
        empty_cols = [col for col in df_combined.columns if df_combined[col].isnull().all()]
        df_combined.drop(empty_cols, axis=1, inplace=True)
        df_combined = df_combined.dropna(subset=['Importe relevante liquida'])

        items = pd.read_csv(files["InvoicesItem"])
        items = items.loc[items['Type'].isin(['Invoice'])]
        df_combined = df_combined.merge(items[['Item', 'Customer External ID']], left_on='Referencia', right_on='Item', how='left')
        df_combined = df_combined[['Resultado', 'Pedido CCF', 'Documento Venta', 'Material', 'Denominación de posición', 'Fecha de Liquidación', 'Importe transferencia', 'Importe relevante liquida', 'Documento Liquidación', 'Importe Comisión a Facturar', 'Referencia', 'archivo_origen', 'Customer External ID']]
        pago = df_combined.loc[df_combined['Importe relevante liquida'] > 0]

        invoices = pd.read_csv(files["Invoices"])
        invoices = invoices[invoices['Status'] != 'Paid In Full']
        invoices1 = invoices.drop_duplicates(subset=['Tax Number'], keep='first')
        invoices2 = invoices[~invoices['Internal ID'].isin(invoices1['Internal ID'])]

        pago = pago.merge(invoices1[['Tax Number','Amount (Gross)', 'Internal ID']], right_on='Tax Number',left_on='Customer External ID', how='left')
        pago['Primera factura'] = -pago['Amount (Gross)']+pago['Importe transferencia']

        def importecorrecto(row):
            if row['Primera factura'] > 0:
                return row['Amount (Gross)']
            elif row['Primera factura'] == 0:
                return row['Importe transferencia']
            else:
                return row['Importe transferencia']

        pago['Factura 1'] = pago.apply(importecorrecto, axis=1)
        pago = pago.merge(invoices2[['Tax Number','Amount (Gross)','Internal ID']], right_on='Tax Number',left_on='Customer External ID', how='left')
        pago['Segunda factura'] = -pago['Amount (Gross)_y']+pago['Primera factura']

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

        factura1 = pago[["Fecha de Liquidación", "Customer External ID", "Importe transferencia", "Referencia", "Tax Number_x", "Amount (Gross)_x", "Internal ID_x", "Factura 1"]]
        factura2 = pago[["Fecha de Liquidación", "Customer External ID", "Importe transferencia", "Referencia", "Tax Number_y", "Amount (Gross)_y", "Internal ID_y", "Factura 2"]]
        factura1 = factura1.rename(columns={"Fecha de Liquidación":"Date","Customer External ID":"Cliente_external ID","Tax Number_x": "Tax Number", "Amount (Gross)_x": "Amount (Gross)", "Internal ID_x": "Factura_INTERNAL ID","Factura 1": "Importe"})
        factura2 = factura2.rename(columns={"Fecha de Liquidación":"Date","Customer External ID":"Cliente_external ID","Tax Number_y": "Tax Number", "Amount (Gross)_y": "Amount (Gross)", "Internal ID_y": "Factura_INTERNAL ID","Factura 2": "Importe"})
        pago = pd.concat([factura1, factura2], ignore_index=True)
        pago = pago.dropna(subset=['Importe'])
        pago = pago[pago['Importe'] != 0]
        pago = pago.drop(columns=['Tax Number','Amount (Gross)','Referencia','Importe transferencia'])
        pago['External ID'] = pago.apply(lambda x: f'{int(x["Factura_INTERNAL ID"])}_PAY' if pd.notna(x["Factura_INTERNAL ID"]) else 'NaN_PAY', axis=1)
        pago['Cuenta Banco_EXTERNAL ID'] = 572000005
        pago['Date'] = pd.to_datetime(pago['Date'], format='%d/%m/%Y', errors='coerce')

        # 3. Guardar en Excel (BytesIO)
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            pago.to_excel(writer, sheet_name='Pago', index=False)
            df_combined.to_excel(writer, sheet_name='Check', index=False)
            items.to_excel(writer, sheet_name='Item Internal ID', index=False)
            invoices.to_excel(writer, sheet_name='Invoices', index=False)
        output.seek(0)
        return output

    except Exception as e:
        raise RuntimeError(f"Error al procesar el script: {str(e)}")
