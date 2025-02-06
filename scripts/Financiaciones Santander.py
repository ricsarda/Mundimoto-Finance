# scripts/financiaciones_renting.py

import pdfplumber
import pandas as pd
import numpy as np
from datetime import datetime
import re
from io import BytesIO

def convertir_fecha(fecha):
    meses = {
        "ENERO": "1", "FEBRERO": "2", "MARZO": "3", "ABRIL": "4",
        "MAYO": "5", "JUNIO": "6", "JULIO": "7", "AGOSTO": "8",
        "SEPTIEMBRE": "9", "OCTUBRE": "10", "NOVIEMBRE": "11", "DICIEMBRE": "12",
        "de": "/", ".": ""
    }
    for es, en in meses.items():
        fecha = fecha.replace(es, en)
    return ' '.join(fecha.split()[0:3])

def reorganizar_excepciones(df):
    if 'Extra 0' in df.columns and 'Extra 1' in df.columns:
        extra_rows = df[df['Tipo'] == '115']
        for idx, row in extra_rows.iterrows():
            if pd.notna(row['Extra 0']) and pd.notna(row['Extra 1']):
                df.at[idx, 'Tipo'] = 'Compensaciones'
                df.at[idx, 'Datos'] = f"{row['Extra 0']} {row['Extra 1']}"
        df = df.drop(columns=['Extra 0', 'Extra 1'])
    return df

def procesar_pdf(pdf_buffer):
    """
    Procesa el contenido de un PDF (BytesIO) y retorna un DataFrame.
    """
    try:
        with pdfplumber.open(pdf_buffer) as pdf:
            texto = "\n".join(
                pagina.extract_text()
                for pagina in pdf.pages
                if pagina.extract_text()
            )

        lineas = texto.split('\n')
        # Fecha
        linea_fecha = next((linea for linea in lineas if linea.startswith('Madrid,')), None)
        if not linea_fecha:
            # Si no hay 'Madrid,', forzamos vacío
            fecha_str = ''
        else:
            partes = linea_fecha.split('Madrid,', 1)
            fecha_str = partes[1].strip() if len(partes) > 1 else ''

        # Buscar "OPERACION"
        indice_inicio = None
        for i, linea in enumerate(lineas):
            if "OPERACION" in linea:
                indice_inicio = i
                break
        if indice_inicio is None:
            return pd.DataFrame()  # PDF no tiene "OPERACION"

        datos_operacion = [l.split() for l in lineas[indice_inicio:] if l.split()]
        max_cols = max(len(l) for l in datos_operacion)
        columnas = (
            ["Tipo", "Datos", "Fecha", "Hora", "Importe", "Concepto"]
            + [f"Extra {i}" for i in range(max_cols - 6)]
        )

        _operacion = pd.DataFrame(datos_operacion, columns=columnas[:max_cols])\
                     .replace([None, 'None'], np.nan)\
                     .dropna(how='all')

        # Unir a 'Datos' los valores de Fecha, Hora, Importe, Concepto
        _operacion['Datos'] = _operacion.apply(
            lambda x: ' '.join(str(v) for v in [x['Datos'], x['Fecha'], x['Hora'], x['Importe'], x['Concepto']] if pd.notna(v)),
            axis=1
        )
        _operacion.drop(columns=["Fecha", "Hora", "Importe", "Concepto"], inplace=True, errors='ignore')

        # Quitar filas irrelevantes
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

        def apply_condiciones(row):
            if row['Tipo'] in condiciones:
                return row.replace(condiciones[row['Tipo']], '')
            return row

        _operacion = _operacion.apply(apply_condiciones, axis=1)

        _operacion['Tipo'] = _operacion['Tipo'].replace({
            'OPERACION:': 'OPERACION',
            'TITULAR:': 'TITULAR',
            '001': 'PAGO AL PROVEEDOR',
            '002': 'ENTREGA INICIAL',
            '041': 'COMISION A TERCEROS'
        })

        # Remplazos en la columna 'Datos'
        _operacion = _operacion.replace(to_replace=[r'\.', r' EUROS', r'-'], value='', regex=True)
        _operacion['Datos'] = _operacion['Datos'].replace(',', '.', regex=True)

        _operacion = reorganizar_excepciones(_operacion)

        # Transformar el DataFrame en filas
        rows = []
        current_row = {}

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

        # Última fila
        if current_row:
            rows.append(current_row)

        # Convertir Compensaciones en columnas
        for r in rows:
            if 'Compensaciones' in r:
                comp_list = r.pop('Compensaciones')
                for i, comp in enumerate(comp_list):
                    r[f'Compensacion_{i+1}'] = comp

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
    """
    Función principal para integrarla en tu app (Streamlit u otra):
      - 'files': dict con BytesIO de los Excels ('financiaciones', 'ventas_SF'), si los necesitas.
      - 'pdfs': dict con BytesIO de los PDFs que quieras analizar.
      - 'new_excel': BytesIO donde se escribirá el Excel final.
      - 'month', 'year': opcionales, si tu app los requiere.

    Retorna un BytesIO con las hojas:
      1) Import (final_operaciones)
      2) Pago (codigocliente)
      3) Financiaciones 2025 (df financiaciones)
      4) Ventas SF (df ventas_SF)
    """
    try:
        # 1. Parsear todos los PDFs en memoria
        dataframes = []
        for nombre_pdf, buffer_pdf in pdfs.items():
            df_temp = procesar_pdf(buffer_pdf)
            if not df_temp.empty:
                dataframes.append(df_temp)

        if dataframes:
            df_final = pd.concat(dataframes, ignore_index=True)
        else:
            df_final = pd.DataFrame()

        # Limpieza de columnas Extra, Hemos, que
        for col in list(df_final.columns):
            if any(sub in col for sub in ['Extra', 'Hemos', 'que']):
                df_final.drop(columns=col, inplace=True, errors='ignore')

        # Reemplazos en PAGO AL PROVEEDOR, COMISION A TERCEROS, TOTAL, ENTREGA INICIAL
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
            df_final['ENTREGA INICIAL'] = pd.to_numeric(df_final['ENTREGA INICIAL'], errors='coerce')

        # Convertir a numérico
        for c in ['PAGO AL PROVEEDOR', 'ENTREGA INICIAL', 'COMISION A TERCEROS', 'TOTAL']:
            if c in df_final.columns:
                df_final[c] = pd.to_numeric(df_final[c], errors='coerce').fillna(0)

        # Quitar filas que no tengan 'TOTAL'
        if 'TOTAL' in df_final.columns:
            df_final = df_final[df_final['TOTAL'] != 0]

        # 2. Construir new_rows
        compensaciones_cols = [col for col in df_final.columns if col.startswith('Compensacion_')]
        new_rows = []

        for _, row in df_final.iterrows():
            new_rows.append({
                'FechaAsiento': row.get('Fecha', ''),
                'CargoAbono': 'D',
                'CodigoCuenta': '572000004',
                'ImporteAsiento': row.get('TOTAL', 0),
                'Comentario': f'FINANC. SANTANDER - {row.get("OPERACION", "")}',
                'Utilidad': 'Total',
            })
            if row.get('COMISION A TERCEROS', 0) > 0:
                new_rows.append({
                    'FechaAsiento': row['Fecha'],
                    'CargoAbono': 'H',
                    'CodigoCuenta': '754000000',
                    'ImporteAsiento': row['COMISION A TERCEROS'],
                    'Comentario': f'FINANC. SANTANDER - {row["OPERACION"]}',
                    'Utilidad': 'Comision Terceros',
                })
            if row.get('PAGO AL PROVEEDOR', 0) > 0 or row.get('ENTREGA INICIAL', 0) > 0:
                titular = row.get('TITULAR', '99999999')
                new_rows.append({
                    'FechaAsiento': row['Fecha'],
                    'CargoAbono': 'H',
                    'CodigoCuenta': titular,
                    'ImporteAsiento': row.get('PAGO AL PROVEEDOR', 0) - row.get('ENTREGA INICIAL', 0),
                    'Comentario': f'FINANC. SANTANDER - {row.get("OPERACION", "")}',
                    'Utilidad': 'Pago Proveedor - Entrega Inicial',
                })
            # Compensaciones
            for colcomp in compensaciones_cols:
                val = row.get(colcomp, None)
                if pd.notna(val):
                    new_rows.append({
                        'FechaAsiento': row['Fecha'],
                        'CargoAbono': 'D',
                        'CodigoCuenta': '754000000',
                        'ImporteAsiento': val,
                        'Comentario': val,
                        'Utilidad': 'Compensaciones',
                    })

        final_operaciones = pd.DataFrame(new_rows)

        # Leer financiaciones/ventas_SF desde 'files'
        # Si no existen, DF vacío
        financiaciones_df = pd.DataFrame()
        ventas_SF_df = pd.DataFrame()

        if 'financiaciones' in files:
            try:
                financiaciones_df = pd.read_excel(files["financiaciones"], engine='openpyxl')
            except Exception as e:
                print("Error leyendo financiaciones:", e)

        if 'ventas_SF' in files:
            try:
                ventas_SF_df = pd.read_excel(files["ventas_SF"], engine='openpyxl')
            except Exception as e:
                print("Error leyendo ventas_SF:", e)

        # Sacar compensaciones y codigocliente
        compensaciones = final_operaciones[final_operaciones['Utilidad'] == 'Compensaciones'].copy()
        codigocliente = final_operaciones[final_operaciones['Utilidad'] == 'Pago Proveedor - Entrega Inicial'].copy()

        # Dejamos en final_operaciones las otras 'Utilidad'
        final_operaciones = final_operaciones[~final_operaciones['Utilidad'].isin(['Pago Proveedor - Entrega Inicial','Compensaciones'])]

        # Merge en codigocliente con financiaciones y ventas_SF
        codigocliente['Operación'] = codigocliente['Comentario'].str.replace('FINANC. SANTANDER - ', '', regex=False)
        if not financiaciones_df.empty and 'Operación' in financiaciones_df.columns and 'MATRÍCULA' in financiaciones_df.columns:
            codigocliente = codigocliente.merge(financiaciones_df[['Operación','MATRÍCULA']], on='Operación', how='left')
        if not ventas_SF_df.empty and 'Moto' in ventas_SF_df.columns and 'DNI' in ventas_SF_df.columns:
            codigocliente = codigocliente.merge(ventas_SF_df[['Moto','DNI']], right_on='Moto', left_on='MATRÍCULA', how='left')
            codigocliente['External ID'] = codigocliente['DNI']

            def accountidcliente(r):
                if pd.isna(r['External ID']):
                    return r['CodigoCuenta']
                else:
                    return r['External ID']
            codigocliente['External ID'] = codigocliente.apply(accountidcliente, axis=1)

        codigocliente = codigocliente[['FechaAsiento','External ID','ImporteAsiento','Operación']]

        # Ajustar compensaciones
        def reformatear_comentario(comm):
            partes = str(comm).split()
            if len(partes) == 2:
                codigo, importe = partes
                codigo_formateado = f'FINANC. SANTANDER - E {codigo[1]} {codigo[2:4]} {codigo[4:8]} {codigo[8:]}'
                return codigo_formateado, importe
            return comm, None

        for idx, c_row in compensaciones.iterrows():
            if c_row['Utilidad'] == 'Compensaciones':
                nuevo_comentario, nuevo_importe = reformatear_comentario(c_row['Comentario'])
                if nuevo_importe:
                    compensaciones.at[idx, 'ImporteAsiento'] = nuevo_importe
                    compensaciones.at[idx, 'Comentario'] = nuevo_comentario

        compensaciones.drop_duplicates(inplace=True)
        compensaciones['ImporteAsiento'] = compensaciones['ImporteAsiento'].astype(str).str.replace(',', '.')
        compensaciones['ImporteAsiento'] = pd.to_numeric(compensaciones['ImporteAsiento'], errors='coerce').fillna(0)
        compensaciones['Account ID'] = 2358

        final_operaciones = pd.concat([final_operaciones, compensaciones], ignore_index=True)

        # Añadir columnas
        final_operaciones['Fecha'] = final_operaciones['FechaAsiento']
        final_operaciones['Descripcion linea'] = final_operaciones['Comentario']
        final_operaciones['Memo'] = final_operaciones['Comentario']

        def credit(r):
            return r['ImporteAsiento'] if r['CargoAbono'] == 'H' else None
        def debit(r):
            return r['ImporteAsiento'] if r['CargoAbono'] == 'D' else None

        final_operaciones['Credit'] = final_operaciones.apply(credit, axis=1)
        final_operaciones['Debit'] = final_operaciones.apply(debit, axis=1)
        final_operaciones['ExternalID'] = final_operaciones['Utilidad'] + '_' + final_operaciones['FechaAsiento'].astype(str)

        ordenfinalcolumnas = ['ExternalID', 'Fecha', 'Memo', 'CodigoCuenta', 'Credit', 'Debit', 'Descripcion linea']
        final_operaciones = final_operaciones[ordenfinalcolumnas]

        # Contraparte
        operaciones_contraparte = final_operaciones.copy()
        operaciones_contraparte.rename(columns={'Credit':'Debit','Debit':'Credit'}, inplace=True)
        final_operaciones = pd.concat([final_operaciones, operaciones_contraparte], ignore_index=True)

        # Generar Excel final en memoria
        with pd.ExcelWriter(new_excel, engine='openpyxl') as writer:
            final_operaciones.to_excel(writer, sheet_name='Import', index=False)
            codigocliente.to_excel(writer, sheet_name='Pago', index=False)

            if not financiaciones_df.empty:
                financiaciones_df.to_excel(writer, sheet_name='Financiaciones 2025', index=False)
            if not ventas_SF_df.empty:
                ventas_SF_df.to_excel(writer, sheet_name='Ventas SF', index=False)

        new_excel.seek(0)
        return new_excel

    except Exception as e:
        raise RuntimeError(f"Error al procesar el script: {str(e)}")
