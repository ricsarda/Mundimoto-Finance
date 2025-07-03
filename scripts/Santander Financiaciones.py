import pdfplumber
import pandas as pd
import numpy as np
from datetime import datetime
from io import BytesIO
import re

def convertir_fecha(fecha_str):
    meses = {
        "ENERO": "1", "FEBRERO": "2", "MARZO": "3", "ABRIL": "4",
        "MAYO": "5", "JUNIO": "6", "JULIO": "7", "AGOSTO": "8",
        "SEPTIEMBRE": "9", "OCTUBRE": "10", "NOVIEMBRE": "11", "DICIEMBRE": "12",
        "de": "/", ".": ""
    }
    for es, en in meses.items():
        fecha_str = fecha_str.replace(es, en)
    return ' '.join(fecha_str.split()[0:3])

def reorganizar_excepciones(df):
    if 'Extra 0' in df.columns and 'Extra 1' in df.columns:
        extra_rows = df[df['Tipo'] == '115']
        for idx, row in extra_rows.iterrows():
            if pd.notna(row['Extra 0']) and pd.notna(row['Extra 1']):
                df.at[idx, 'Tipo'] = 'Compensaciones'
                df.at[idx, 'Datos'] = f"{row['Extra 0']} {row['Extra 1']}"
        df = df.drop(columns=['Extra 0', 'Extra 1'], errors='ignore')
    return df

def procesar_pdf_en_memoria(pdf_buffer):
    try:
        with pdfplumber.open(pdf_buffer) as pdf:
            texto = "\n".join([
                pagina.extract_text() for pagina in pdf.pages if pagina.extract_text()
            ])

        lineas = texto.split('\n')
        # Tomar la línea que empieza con 'Madrid,'
        linea_fecha = next((l for l in lineas if l.startswith('Madrid,')), None)
        if linea_fecha:
            partes = linea_fecha.split('Madrid,', 1)
            fecha_str = partes[1].strip() if len(partes) > 1 else ''
        else:
            fecha_str = ""

        # Ubicar la línea donde aparece "OPERACION" para partir
        indice_inicio = next(
            (i for i, linea in enumerate(lineas) if "OPERACION" in linea),
            None
        )
        if indice_inicio is None:
            return pd.DataFrame()

        datos_operacion = [l.split() for l in lineas[indice_inicio:] if l.split()]

        max_cols = max(len(l) for l in datos_operacion)
        columnas = [
            "Tipo", "Datos", "Fecha", "Hora", "Importe", "Concepto"
        ] + [f"Extra {i}" for i in range(max_cols - 6)]

        _operacion = pd.DataFrame(
            datos_operacion,
            columns=columnas[:max_cols]
        ).replace([None, 'None'], np.nan).dropna(how='all')

        def unir_datos(row):
            valores = [row['Datos'], row['Fecha'], row['Hora'], row['Importe'], row['Concepto']]
            return ' '.join(str(v) for v in valores if pd.notna(v))
        _operacion['Datos'] = _operacion.apply(unir_datos, axis=1)
        _operacion.drop(columns=["Fecha", "Hora", "Importe", "Concepto"], inplace=True, errors='ignore')

        _operacion = _operacion[~_operacion['Tipo'].isin([
            'RELACION', '---------------------', 'Aténtamente,',
            'DEPARTAMENTOAUTOMOCION', 'SERVICIOPAGOAPRESCRIPTORES'
        ])]

        condiciones = {
            '001': '001 PAGO AL PROVEEDOR ',
            '002': '001 ENTREGA INICIAL ',
            '041': '001 COMISION A TERCEROS ',
            '115': '002 Compensación Operación',
            'TOTAL': 'OPERACION: '
        }
        def quitar_prefijo(row):
            if row['Tipo'] in condiciones:
                return row.replace(condiciones[row['Tipo']], '')
            return row
        _operacion = _operacion.apply(quitar_prefijo, axis=1)

        _operacion['Tipo'] = _operacion['Tipo'].replace({
            'OPERACION:': 'OPERACION',
            'TITULAR:': 'TITULAR',
            '001': 'PAGO AL PROVEEDOR',
            '002': 'ENTREGA INICIAL',
            '041': 'COMISION A TERCEROS'
        })

        _operacion = _operacion.replace(
            to_replace=[r'\.', r' EUROS', r'-'],
            value='',
            regex=True
        )
        _operacion['Datos'] = _operacion['Datos'].replace(',', '.', regex=True)

        _operacion = reorganizar_excepciones(_operacion)

        rows, current_row = [], {}

        for _, fila in _operacion.iterrows():
            if fila['Tipo'] == 'OPERACION':
                if current_row:
                    rows.append(current_row)
                current_row = {'OPERACION': fila['Datos']}
            elif fila['Tipo'] == 'TOTAL':
                current_row['TOTAL'] = fila['Datos']
                rows.append(current_row)
                current_row = {}
            elif fila['Tipo'] == 'Compensaciones':
                if 'Compensaciones' not in current_row:
                    current_row['Compensaciones'] = []
                current_row['Compensaciones'].append(fila['Datos'])
            else:
                current_row[fila['Tipo']] = fila['Datos']

        if current_row:
            rows.append(current_row)

        reorganized_df = pd.DataFrame(rows)
        reorganized_df['Fecha'] = convertir_fecha(fecha_str)

        if 'ENTREGA INICIAL' not in reorganized_df:
            reorganized_df['ENTREGA INICIAL'] = '001 ENTREGA INICIAL 0'
        reorganized_df['ENTREGA INICIAL'] = reorganized_df['ENTREGA INICIAL'].fillna('001 ENTREGA INICIAL 0')

        return reorganized_df

    except Exception as e:
        print(f"Error al procesar PDF en memoria: {e}")
        return pd.DataFrame()

def main(files, pdfs, new_excel, month=None, year=None):
    try:
        dataframes = []
        for pdf_name, pdf_buffer in pdfs.items():
            df_parcial = procesar_pdf_en_memoria(pdf_buffer)
            if not df_parcial.empty:
                dataframes.append(df_parcial)

        if dataframes:
            df_final = pd.concat(dataframes, ignore_index=True)
        else:
            df_final = pd.DataFrame()

        for col in list(df_final.columns):
            if any(s in col for s in ['Extra', 'Hemos', 'que']):
                df_final.drop(columns=col, inplace=True, errors='ignore')

        if 'PAGO AL PROVEEDOR' in df_final.columns:
            df_final['PAGO AL PROVEEDOR'] = df_final['PAGO AL PROVEEDOR'].str.replace('001 PAGO AL PROVEEDOR ', '', regex=False)
            df_final['PAGO AL PROVEEDOR'] = df_final['PAGO AL PROVEEDOR'].str.replace('002 PAGO AL PROVEEDOR ', '', regex=False)
        if 'COMISION A TERCEROS' in df_final.columns:
            df_final['COMISION A TERCEROS'] = df_final['COMISION A TERCEROS'].str.replace('001 COMISION A TERCEROS ', '', regex=False)
        if 'TOTAL' in df_final.columns:
            df_final['TOTAL'] = df_final['TOTAL'].str.replace('OPERACION: ', '', regex=False)
        if 'ENTREGA INICIAL' in df_final.columns:
            df_final['ENTREGA INICIAL'] = df_final['ENTREGA INICIAL'].str.replace('001 ENTREGA INICIAL ', '', regex=False)
            df_final['ENTREGA INICIAL'] = df_final['ENTREGA INICIAL'].str.replace('002 ENTREGA INICIAL ', '', regex=False)

        for c in ['PAGO AL PROVEEDOR','ENTREGA INICIAL','COMISION A TERCEROS','TOTAL']:
            if c in df_final.columns:
                df_final[c] = pd.to_numeric(df_final[c], errors='coerce')

        if 'TOTAL' in df_final.columns:
            df_final.dropna(subset=['TOTAL'], inplace=True)

        compensaciones_cols = [col for col in df_final.columns if col.startswith('Compensacion_')]
        rows_nuevas = []

        for i, row in df_final.iterrows():
            rows_nuevas.append({
                'Date': row['Fecha'],
                'CargoAbono': 'D',
                'CodigoCuenta': '572000004',
                'ImporteAsiento': row['TOTAL'],
                'Comentario': f'FINANC. SANTANDER - {row.get("OPERACION","")}',
                'Utilidad': 'Total'
            })
            if row.get('COMISION A TERCEROS', 0) > 0:
                rows_nuevas.append({
                    'Date': row['Fecha'],
                    'CargoAbono': 'H',
                    'CodigoCuenta': '754000000',
                    'ImporteAsiento': row['COMISION A TERCEROS'],
                    'Comentario': f'FINANC. SANTANDER - {row.get("OPERACION","")}',
                    'Utilidad': 'Comision Terceros'
                })
            pago_prov = row.get('PAGO AL PROVEEDOR', 0)
            ent_ini = row.get('ENTREGA INICIAL', 0)
            if (pago_prov > 0) or (ent_ini > 0):
                titular = row.get('TITULAR','99999999')
                rows_nuevas.append({
                    'Date': row['Fecha'],
                    'CargoAbono': 'H',
                    'CodigoCuenta': titular,
                    'ImporteAsiento': (pago_prov - ent_ini),
                    'Comentario': f'FINANC. SANTANDER - {row.get("OPERACION","")}',
                    'Utilidad': 'Pago Proveedor - Entrega Inicial'
                })
            for colc in compensaciones_cols:
                if pd.notna(row[colc]):
                    rows_nuevas.append({
                        'Date': row['Fecha'],
                        'CargoAbono': 'D',
                        'CodigoCuenta': '754000000',
                        'ImporteAsiento': row[colc],
                        'Comentario': str(row[colc]),
                        'Utilidad': 'Compensaciones'
                    })

        financiaciones = pd.read_csv(files["Financiaciones"])
        invoice = pd.read_csv(files["Invoices"], sep=',')
        invoices = pd.read_csv(files["invoice"], sep=',')
        invoice = invoice[invoice['Type'] == 'Invoice']
        invoices = invoices[invoices['Account'] != '430000001 Clientes - Renting']
        invoices1 = invoices.drop_duplicates(subset=['Tax Number'], keep='first')
        invoices2 = invoices[~invoices['Internal ID'].isin(invoices1['Internal ID'])]
        final_operaciones = pd.DataFrame(rows_nuevas)
        compensaciones = final_operaciones[final_operaciones['Utilidad'] == 'Compensaciones']
        compensaciones['Comentario'] = compensaciones['Comentario'].astype(str)
        pagopre = final_operaciones[final_operaciones['Utilidad'] == 'Pago Proveedor - Entrega Inicial']
        final_operaciones = final_operaciones[final_operaciones['Utilidad'] == 'Comision Terceros']

        pagopre['Operación'] = pagopre['Comentario'].astype(str).str.replace('FINANC. SANTANDER - ', '', regex=False)

        pagopre = pagopre.merge(financiaciones[['Operación', 'MATRÍCULA']], on='Operación', how='left')

        pagopre = pagopre.merge(invoice[['Item', 'Customer External ID']], right_on='Item',left_on='MATRÍCULA', how='left')
        pagopre['Cliente_external ID'] = pagopre['Customer External ID']

        def accountidcliente(row):
            if pd.isna(row['Cliente_external ID']) :
                return row['CodigoCuenta']
            else:
                return row['Cliente_external ID']


        pagopre['Cliente_external ID'] = pagopre.apply(accountidcliente, axis=1)
        pago = pagopre[['Date', 'Cliente_external ID', 'ImporteAsiento', 'Operación','Item']]
        pago = pago.merge(invoices1[['Tax Number','Amount (Gross)','Internal ID']], right_on='Tax Number',left_on='Cliente_external ID', how='left')
        pago['Primera factura'] = -pago['Amount (Gross)']+pago['ImporteAsiento']

        def importecorrecto(row):
            if row['Primera factura'] > 0:
                return row['Amount (Gross)']
            elif row['Primera factura'] == 0:
                return row['ImporteAsiento']
            else:
                return row['ImporteAsiento']

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

        factura1 = pago[["Date","Cliente_external ID", "ImporteAsiento","Operación","Item", "Tax Number_x" , "Amount (Gross)_x",  "Internal ID_x" , "Factura 1"]]
        factura2 = pago[["Date","Cliente_external ID", "ImporteAsiento","Operación","Item", "Tax Number_y" , "Amount (Gross)_y",  "Internal ID_y" , "Factura 2"]]
        factura1 = factura1.rename(columns={"Tax Number_x": "Tax Number", "Amount (Gross)_x": "Amount (Gross)", "Internal ID_x": "Factura_INTERNAL ID","Factura 1": "Importe"} )
        factura2 = factura2.rename(columns={"Tax Number_y": "Tax Number", "Amount (Gross)_y": "Amount (Gross)", "Internal ID_y": "Factura_INTERNAL ID","Factura 2": "Importe"} )
        pago = pd.concat([factura1, factura2], ignore_index=True)
        pago = pago.dropna(subset=['Importe'])
        pago = pago[pago['Importe'] != 0]
        pago = pago.drop(columns=['Tax Number','Amount (Gross)','Item','Operación','ImporteAsiento'])
        pago['External ID'] = pago.apply(lambda x: f'{int(x["Factura_INTERNAL ID"])}_PAY' if pd.notna(x["Factura_INTERNAL ID"]) else '_PAY', axis=1)
        pago = pago.drop_duplicates(subset=['Factura_INTERNAL ID'], keep='first')
        pago['Cuenta Banco_EXTERNAL ID'] = 572000004

        def reformatear_comentario(comentario):
            partes = comentario.split()
            if len(partes) == 2:
                codigo, importe = partes
                codigo_formateado = f'FINANC. SANTANDER - E {codigo[1]} {codigo[2:4]} {codigo[4:8]} {codigo[8:]}'
                return codigo_formateado, importe
            return comentario, None

        # Aplicar los cambios en el DataFrame
        for idx, row in compensaciones.iterrows():
            if row['Utilidad'] == 'Compensaciones':
                comentario = str(row.get('Comentario', ''))
                nuevo_comentario, nuevo_importe = reformatear_comentario(comentario)
                if nuevo_importe:
                    compensaciones.at[idx, 'ImporteAsiento'] = nuevo_importe
                    compensaciones.at[idx, 'Comentario'] = nuevo_comentario
        compensaciones = compensaciones.drop_duplicates()
        compensaciones['ImporteAsiento'] = compensaciones['ImporteAsiento'].str.replace(',', '.')
        compensaciones['ImporteAsiento'] = compensaciones['ImporteAsiento'].astype(float)
        compensaciones['Account ID'] = 2358
        final_operaciones = pd.concat([final_operaciones, compensaciones], ignore_index=True)

        final_operaciones['Fecha'] = final_operaciones['Date']
        final_operaciones['Descripcion linea'] = final_operaciones['Comentario']
        final_operaciones['Memo'] = final_operaciones['Comentario']

        def credit(row):
            if row['CargoAbono'] == 'H':
                return row['ImporteAsiento']
            else:
                return None
    
        def debit(row):
            if row['CargoAbono'] == 'D':
                return row['ImporteAsiento']
            else:
                return None

        final_operaciones['Credit'] = final_operaciones.apply(credit, axis=1)
        final_operaciones['Debit'] = final_operaciones.apply(debit, axis=1)
        final_operaciones['ExternalID'] = final_operaciones['Utilidad'] +'_'+ final_operaciones['Fecha'].astype(str)

        ordenfinalcolumnas = ['ExternalID', 'Fecha', 'Memo', 'Account ID', 'Credit', 'Debit' , 'Descripcion linea']
        final_operaciones = final_operaciones[ordenfinalcolumnas]

        operaciones_contraparte = final_operaciones.copy()
        operaciones_contraparte.rename(columns={'Credit': 'Debit', 'Debit': 'Credit'}, inplace=True)
        operaciones_contraparte['Account ID'] = 2437

        final_operaciones = pd.concat([final_operaciones, operaciones_contraparte], ignore_index=True)
        def cuenta_faltante(row):
            if pd.isna(row['Account ID']):
                return 2358
            else:
                return row['Account ID']
        final_operaciones['Account ID'] = final_operaciones.apply(cuenta_faltante, axis=1)
        final_operaciones = final_operaciones.sort_values('ExternalID')
        final_operaciones.to_excel(ruta_archivo_final_excel2, index=False)

        # Escribir a dos Excel en memoria
        excel_final_ops = BytesIO()  # asientos
        excel_rest = BytesIO()       # resto

        with pd.ExcelWriter(excel_final_ops, engine='openpyxl') as writer:
            final_operaciones.to_excel(writer, sheet_name='Import', index=False)

        with pd.ExcelWriter(excel_rest, engine='openpyxl') as writer:
            codigocliente.to_excel(writer, sheet_name='Pago', index=False)
            pagopre.to_excel(writer, sheet_name='Check', index=False)
            financiaciones.to_excel(writer, sheet_name='Financiaciones 2025', index=False)
            invoice.to_excel(writer, sheet_name='Item Internal ID', index=False)
            invoices.to_excel(writer, sheet_name='Invoices', index=False)

        excel_final_ops.seek(0)
        excel_rest.seek(0)

        return excel_final_ops, excel_rest

    except Exception as e:
        raise RuntimeError(f"Error al procesar el script: {str(e)}")
