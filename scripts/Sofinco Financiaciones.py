import pdfplumber
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
from io import BytesIO

def main(files, pdfs, new_excel, month=None, year=None):
    try:
        # Procesar PDFs rotándolos en memoria y extrayendo texto línea a línea
        data = []
        date_by_pdf = {}

        for nombre_pdf, buffer in pdfs.items():
            buffer.seek(0)
            reader = PdfReader(buffer)
            writer = PdfWriter()

            for page in reader.pages:
                page.rotate(90)
                writer.add_page(page)

            rotated_pdf = BytesIO()
            writer.write(rotated_pdf)
            rotated_pdf.seek(0)

            with pdfplumber.open(rotated_pdf) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        for line in text.split('\n'):
                            data.append({
                                'archivo': nombre_pdf,
                                'linea_texto': line.strip()
                            })

        df_raw = pd.DataFrame(data)

        # Buscar la fecha de operación en cada PDF
        for pdf_file in df_raw['archivo'].unique():
            subset_pdf = df_raw[df_raw['archivo'] == pdf_file]
            fecha_operacion = None
            for idx, row in subset_pdf.iterrows():
                line = row['linea_texto']
                if line.startswith("Fecha Operación: ") or line.startswith("Fecha Operacion: "):
                    partes = line.split(':', 1)
                    if len(partes) == 2:
                        fecha_operacion = partes[1].strip()
                    break
            date_by_pdf[pdf_file] = fecha_operacion

        def parse_importe(importe_str):
            return float(importe_str.replace('.', '').replace(',', '.'))

        rows_limpias = []
        for idx, row in df_raw.iterrows():
            linea = row['linea_texto']
            archivo = row['archivo']
            if linea.startswith(('Nuevas','Anul.','Cancel.')):
                tokens = linea.split()
                if len(tokens) < 7:
                    continue
                importe_financiado_str = tokens[-3]
                comisiones_str         = tokens[-2]
                importe_neto_str       = tokens[-1]
                num_solicitud = tokens[-4]
                num_contrato  = tokens[-5]
                cliente_tokens = tokens[4:-5]
                cliente = " ".join(cliente_tokens)
                tipo_operacion = tokens[0]
                matricula      = tokens[1]
                bastidor       = tokens[2]
                nif_cliente    = tokens[3]
                try:
                    importe_financiado = parse_importe(importe_financiado_str)
                    comisiones         = parse_importe(comisiones_str)
                    importe_neto       = parse_importe(importe_neto_str)
                except ValueError:
                    continue
                rows_limpias.append({
                    'archivo'          : row['archivo'],
                    'Fecha'            : date_by_pdf.get(archivo, None),
                    'TipoOperacion'    : tipo_operacion,
                    'Matricula'        : matricula,
                    'Bastidor'         : bastidor,
                    'NIF'              : nif_cliente,
                    'Cliente'          : cliente,
                    'NumeroContrato'   : num_contrato,
                    'NumeroSolicitud'  : num_solicitud,
                    'ImporteFinanciado': importe_financiado,
                    'Comisiones'       : comisiones,
                    'ImporteNeto'      : importe_neto
                })

        df_final = pd.DataFrame(rows_limpias)

        # Reorganización y limpieza
        canyanul = df_final[df_final['TipoOperacion'].str.contains('Anul.|Cancel.')]
        nuevas = df_final[df_final['TipoOperacion'].str.contains('Nuevas')]
        canyanul = canyanul.rename(columns={'Matricula':'M','Bastidor':'Matricula','NIF':'Bastidor'})
        canyanul['NIF'] = canyanul['Cliente'].str[:10]
        canyanul['NIF'] = canyanul['NIF'].str.replace(' ', '')
        df_final = pd.concat([canyanul, nuevas], ignore_index=True)
        df_final['Bastidor'] = df_final['Bastidor'].str.replace('*', '', regex=False)
        if 'M' in df_final.columns:
            df_final = df_final.drop(columns=['M'])

        # Leer archivos adicionales desde files
        invoice = pd.read_csv(files["InvoicesItem"])
        invoice = invoice[invoice['Type'] == 'Invoice']
        invoices = pd.read_csv(files["Invoices"], sep=',')
        invoices = invoices[invoices['Status'] != 'Paid In Full']
        invoices1 = invoices.drop_duplicates(subset=['Tax Number'], keep='first')
        invoices2 = invoices[~invoices['Internal ID'].isin(invoices1['Internal ID'])]

        pagopre = df_final.drop(columns=['Comisiones','ImporteNeto'])
        pagopre['Operación'] = pagopre['NumeroSolicitud']
        pagopre['Importe'] = pagopre['ImporteFinanciado']
        pagopre['Cliente_external ID'] = pagopre['NIF']
        pagopre['Date'] = pagopre['Fecha']
        pago = pagopre.merge(invoices1[['Tax Number','Amount (Gross)','Internal ID']], right_on='Tax Number',left_on='Cliente_external ID', how='left')
        pago['Primera factura'] = -pago['Amount (Gross)']+pago['Importe']

        def importecorrecto(row):
            if row['Primera factura'] > 0:
                return row['Amount (Gross)']
            elif row['Primera factura'] == 0:
                return row['Importe']
            else:
                return row['Importe']

        pago['Factura 1'] = pago.apply(importecorrecto, axis=1)
        pago = pago.merge(invoices2[['Tax Number','Amount (Gross)','Internal ID']], right_on='Tax Number',left_on='Cliente_external ID', how='left')
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

        factura1 = pago[["Date","Cliente_external ID", "Tax Number_x" , "Amount (Gross)_x",  "Internal ID_x" , "Factura 1"]]
        factura2 = pago[["Date","Cliente_external ID", "Tax Number_y" , "Amount (Gross)_y",  "Internal ID_y" , "Factura 2"]]
        factura1 = factura1.rename(columns={"Tax Number_x": "Tax Number", "Amount (Gross)_x": "Amount (Gross)", "Internal ID_x": "Factura_INTERNAL ID","Factura 1": "Importe"} )
        factura2 = factura2.rename(columns={"Tax Number_y": "Tax Number", "Amount (Gross)_y": "Amount (Gross)", "Internal ID_y": "Factura_INTERNAL ID","Factura 2": "Importe"} )
        pago = pd.concat([factura1, factura2], ignore_index=True)
        pago = pago.dropna(subset=['Importe'])
        pago = pago[pago['Importe'] != 0]
        pago['External ID'] = pago.apply(lambda x: f'{int(x["Factura_INTERNAL ID"])}_PAY' if pd.notna(x["Factura_INTERNAL ID"]) else 'NaN_PAY', axis=1)
        pago['Cuenta Banco_EXTERNAL ID'] = 572000004
        pago = pago.drop(columns=['Tax Number','Amount (Gross)'])

        # Comisiones y compensaciones para asientos
        final_operaciones = df_final.drop(columns=['ImporteFinanciado','ImporteNeto'])

        def compensaciones(row):
            if row['TipoOperacion'] != 'Nuevas':
                return 'Compensaciones'
            else:
                return 'Comisiones'

        final_operaciones['TipoOperacion'] = final_operaciones.apply(compensaciones, axis=1)
        comisiones = final_operaciones[final_operaciones['TipoOperacion'] == 'Comisiones'].copy()
        comisiones['Credit'] = comisiones['Comisiones']
        compensacion = final_operaciones[final_operaciones['TipoOperacion'] == 'Compensaciones'].copy()
        compensacion['Debit'] = -(compensacion['Comisiones'])
        compensacion['Account ID'] = 2363
        final_operaciones = pd.concat([comisiones, compensacion], ignore_index=True)

        final_operaciones['Descripcion linea'] = final_operaciones['Cliente']
        final_operaciones['Memo'] = final_operaciones['Cliente']
        final_operaciones['ExternalID'] = final_operaciones['TipoOperacion'] +'_'+ final_operaciones['Fecha'].astype(str)
        operaciones_contraparte = final_operaciones.copy()
        operaciones_contraparte.rename(columns={'Credit': 'Debit', 'Debit': 'Credit'}, inplace=True)
        operaciones_contraparte['Account ID'] = 2437

        final_operaciones = pd.concat([final_operaciones, operaciones_contraparte], ignore_index=True)
        ordenfinalcolumnas = ['ExternalID', 'Fecha', 'Memo', 'Account ID','Credit','Debit', 'Descripcion linea']
        final_operaciones['Account ID'] = final_operaciones['Account ID'].fillna(2363)
        final_operaciones = final_operaciones[ordenfinalcolumnas]
        final_operaciones = final_operaciones.sort_values('ExternalID')

        # Guardar en dos excels (memoria)
        output_ops = BytesIO()
        output_rest = BytesIO()
        with pd.ExcelWriter(output_ops, engine='openpyxl') as writer:
            final_operaciones.to_excel(writer, sheet_name='Import', index=False)
        with pd.ExcelWriter(output_rest, engine='openpyxl') as writer:
            pago.to_excel(writer, sheet_name='Pago', index=False)
            df_final.to_excel(writer, sheet_name='Check', index=False)
            invoice.to_excel(writer, sheet_name='Item Internal ID', index=False)
            invoices.to_excel(writer, sheet_name='Invoices', index=False)
        output_ops.seek(0)
        output_rest.seek(0)
        return output_ops, output_rest

    except Exception as e:
        raise RuntimeError(f"Error al procesar el script: {str(e)}")
