import pdfplumber
import pandas as pd
import numpy as np
from datetime import datetime
from io import BytesIO
import os

def convertir_fecha(fecha_str):
    meses = {"ENERO": "1", "FEBRERO": "2", "MARZO": "3", "ABRIL": "4",
             "MAYO": "5", "JUNIO": "6", "JULIO": "7", "AGOSTO": "8",
             "SEPTIEMBRE": "9", "OCTUBRE": "10", "NOVIEMBRE": "11", "DICIEMBRE": "12",
             "de": "/", ".": ""}
    for es, num in meses.items():
        fecha_str = fecha_str.replace(es, num)
    return ' '.join(fecha_str.split()[0:3])

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
    try:
        # Abrir el PDF desde el buffer en memoria
        with pdfplumber.open(pdf_buffer) as pdf:
            texto = "\n".join([
                pagina.extract_text()
                for pagina in pdf.pages
                if pagina.extract_text()
            ])

        lineas = texto.split('\n')
        linea_fecha = next((linea for linea in lineas if linea.startswith('Madrid,')), "")
        partes = linea_fecha.split('Madrid,', 1)
        fecha_str = partes[1].strip() if len(partes) > 1 else ''
        indice_inicio = next(
            (i for i, linea in enumerate(lineas) if "OPERACION" in linea),
            None
        )
        if indice_inicio is None:
            return pd.DataFrame()

        datos_operacion = [linea.split() for linea in linees[indice_inicio:] if linea.split()]

        max_cols = max(len(linea) for linea in datos_operacion)
        columnas = ["Tipo", "Datos", "Fecha", "Hora", "Importe", "Concepto"] + [
            f"Extra {i}" for i in range(max_cols - 6)
        ]
        _operacion = pd.DataFrame(
            datos_operacion,
            columns=columnas[:max_cols]
        ).replace([None, 'None'], np.nan).dropna(how='all')

        _operacion['Datos'] = _operacion.apply(
            lambda x: ' '.join(str(v) for v in [x['Datos'], x['Fecha'], x['Hora'], x['Importe'], x['Concepto']] if pd.notna(v)),
            axis=1
        )

        _operacion = _operacion.drop(columns=["Fecha", "Hora", "Importe", "Concepto"])

        _operacion = _operacion[~_operacion['Tipo'].isin([
            'RELACION', '---------------------', 'Aténtamente,',
            'DEPARTAMENTOAUTOMOCION','SERVICIOPAGOAPRESCRIPTORES'
        ])]


        condiciones = {
            '001': '001 PAGO AL PROVEEDOR ',
            '002': '001 ENTREGA INICIAL ',
            '041': '001 COMISION A TERCEROS ',
            '115': '002 Compensación Operación',
            'TOTAL': 'OPERACION: '
        }
        _operacion = _operacion.apply(
            lambda x: x.replace(condiciones[x['Tipo']], '')
            if x['Tipo'] in condiciones else x,
            axis=1
        )

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

        # Estructurar
        rows = []
        current_row = {}

        for index, row in _operacion.iterrows():
            if row['Tipo'] == 'OPERACION':
                if current_row:
                    rows.append(current_row)
                current_row = {'OPERACION': row['Datos']}
            elif row['Tipo'] == 'TOTAL':
                current_row['TOTAL'] = row['Datos']
                rows.append(current_row)
                current_row = {}
            elif row['Tipo'] == 'Compensaciones':
                if 'Compensaciones' not in current_row:
                    current_row['Compensaciones'] = []
                current_row['Compensaciones'].append(row['Datos'])
            else:
                current_row[row['Tipo']] = row['Datos']

        if current_row:
            rows.append(current_row)

        for row in rows:
            if 'Compensaciones' in row:
                comp_list = row.pop('Compensaciones')
                for i, comp in enumerate(comp_list):
                    row[f'Compensacion_{i+1}'] = comp

        df_reorganizado = pd.DataFrame(rows)
        df_reorganizado['Fecha'] = convertir_fecha(fecha_str)
        if 'ENTREGA INICIAL' not in df_reorganizado:
            df_reorganizado['ENTREGA INICIAL'] = '001 ENTREGA INICIAL 0'
        df_reorganizado['ENTREGA INICIAL'] = df_reorganizado.get('ENTREGA INICIAL','0').fillna('001 ENTREGA INICIAL 0')
        return df_reorganizado

    except Exception as e:
        print(f"Error al procesar PDF en memoria: {e}")
        return pd.DataFrame()

def main(files, pdfs, new_excel, month=None, year=None):
    try:
        # Leer los PDFs desde 'pdfs'
        dataframes = []
        for pdf_name, pdf_buffer in pdfs.items():
            df_parcial = procesar_pdf(pdf_buffer)
            if not df_parcial.empty:
                dataframes.append(df_parcial)

        if not dataframes:
            with pd.ExcelWriter(new_excel, engine='openpyxl') as writer:
                pd.DataFrame({"Mensaje": ["No se encontraron datos en PDFs"]}).to_excel(
                    writer, sheet_name='SinDatos', index=False
                )
            new_excel.seek(0)
            return new_excel

        df_final = pd.concat(dataframes, ignore_index=True)

        for col in df_final.columns:
            if 'Extra' in col or 'Hemos' in col or 'que' in col:
                df_final.drop(columns=col, inplace=True, errors='ignore')

        if 'PAGO AL PROVEEDOR' in df_final.columns:
            df_final['PAGO AL PROVEEDOR'] = df_final['PAGO AL PROVEEDOR'].str.replace('001 PAGO AL PROVEEDOR ', '')
            df_final['PAGO AL PROVEEDOR'] = df_final['PAGO AL PROVEEDOR'].str.replace('002 PAGO AL PROVEEDOR ', '')
        if 'COMISION A TERCEROS' in df_final.columns:
            df_final['COMISION A TERCEROS'] = df_final['COMISION A TERCEROS'].str.replace('001 COMISION A TERCEROS ', '')
        if 'TOTAL' in df_final.columns:
            df_final['TOTAL'] = df_final['TOTAL'].str.replace('OPERACION: ', '')
        if 'ENTREGA INICIAL' in df_final.columns:
            df_final['ENTREGA INICIAL'] = df_final['ENTREGA INICIAL'].str.replace('001 ENTREGA INICIAL ', '')
            df_final['ENTREGA INICIAL'] = df_final['ENTREGA INICIAL'].str.replace('002 ENTREGA INICIAL ', '')
            df_final['ENTREGA INICIAL'] = df_final['ENTREGA INICIAL'].astype(float, errors='ignore')

        for columna in ['PAGO AL PROVEEDOR', 'ENTREGA INICIAL', 'COMISION A TERCEROS', 'TOTAL']:
            if columna in df_final.columns:
                df_final[columna] = pd.to_numeric(df_final[columna], errors='coerce')

        if 'TOTAL' in df_final.columns:
            df_final = df_final.dropna(subset=['TOTAL'])

        new_rows = []
        compensaciones_cols = [col for col in df_final.columns if col.startswith('Compensacion_')]
        
        for idx, row in df_final.iterrows():
            # Estructura "debito" por TOTAL
            new_rows.append({
                'FechaAsiento': row.get('Fecha', ''),
                'CargoAbono': 'D',
                'CodigoCuenta': '572000004',
                'ImporteAsiento': row.get('TOTAL', 0),
                'Comentario': f'FINANC. SANTANDER - {row.get("OPERACION", "")}',
                'Utilidad': 'Total',
            })
            # COMISION
            comision_val = row.get('COMISION A TERCEROS', 0)
            if comision_val > 0:
                new_rows.append({
                    'FechaAsiento': row.get('Fecha', ''),
                    'CargoAbono': 'H',
                    'CodigoCuenta': '754000000',
                    'ImporteAsiento': comision_val,
                    'Comentario': f'FINANC. SANTANDER - {row.get("OPERACION", "")}',
                    'Utilidad': 'Comision Terceros',
                })
            # Pago Proveedor - Entrega Inicial
            pago_prov = row.get('PAGO AL PROVEEDOR', 0)
            entrega_inicial = row.get('ENTREGA INICIAL', 0)
            titular = row.get('TITULAR', '99999999')
            if (pd.notna(pago_prov) and pd.notna(entrega_inicial)) and (pago_prov != 0 or entrega_inicial != 0):
                new_rows.append({
                    'FechaAsiento': row.get('Fecha', ''),
                    'CargoAbono': 'H',
                    'CodigoCuenta': titular,
                    'ImporteAsiento': (pago_prov - entrega_inicial),
                    'Comentario': f'FINANC. SANTANDER - {row.get("OPERACION", "")}',
                    'Utilidad': 'Pago Proveedor - Entrega Inicial',
                })
            
            # Compensaciones
            for colc in compensaciones_cols:
                val = row.get(colc, None)
                if pd.notna(val):
                    new_rows.append({
                        'FechaAsiento': row.get('Fecha', ''),
                        'CargoAbono': 'D',
                        'CodigoCuenta': '754000000',
                        'ImporteAsiento': val,
                        'Comentario': str(val),
                        'Utilidad': 'Compensaciones',
                    })

        final_operaciones = pd.DataFrame(new_rows)

        Sheet1 = final_operaciones[['FechaAsiento', 'CargoAbono', 'CodigoCuenta', 'ImporteAsiento', 'Comentario']].copy()
        Sheet1['CodigoEmpresa'] = 1
        Sheet1['Ejercicio'] = 2024  # o month/year si quieres
        Sheet1['MantenerAsiento'] = 0
        Sheet1['NumeroPeriodo'] = -1
        Sheet1['Asiento'] = ''
        Sheet1 = Sheet1[[
            'CodigoEmpresa','Ejercicio','MantenerAsiento','NumeroPeriodo','Asiento',
            'FechaAsiento','CargoAbono','CodigoCuenta','ImporteAsiento','Comentario'
        ]]

        if "Clientes" in files and files["Clientes"] is not None:
            clientes_df = pd.read_excel(files["Clientes"], engine='openpyxl')
        else:
            clientes_df = pd.DataFrame()  # o un DF vacío

        if "Ventas" in files and files["Ventas"] is not None:
            ventas_SF = pd.read_excel(files["Ventas"], engine='openpyxl')
        else:
            ventas_SF = pd.DataFrame()

        Sheet1['CodigoCuenta_lower'] = Sheet1['CodigoCuenta'].astype(str).str.lower()
        if not clientes_df.empty and 'Razón social' in clientes_df.columns:
            clientes_df['Razón social_lower'] = clientes_df['Razón social'].str.lower()
            Sheet1 = Sheet1.merge(
                clientes_df[['Razón social', 'Cód. contable', 'Razón social_lower']], 
                left_on='CodigoCuenta_lower', 
                right_on='Razón social_lower', 
                how='left'
            )
            Sheet1.drop(columns=[ 'CodigoCuenta_lower', 'Razón social_lower'], inplace=True, errors='ignore')

            def cod_contable(row):
                if pd.notna(row.get('Cód. contable')):
                    return row['Cód. contable']
                return row['CodigoCuenta']

            Sheet1['CodigoCuenta'] = Sheet1.apply(cod_contable, axis=1)
            Sheet1.drop(columns=['Cód. contable','Razón social'], inplace=True, errors='ignore')
        else:
            pass

        Sheet1['ImporteAsiento'] = pd.to_numeric(Sheet1['ImporteAsiento'], errors='coerce').round(2)
        control = Sheet1.groupby(['FechaAsiento','CargoAbono'], dropna=False).agg({'ImporteAsiento':'sum'}).reset_index()

        def debe(row):
            return row['ImporteAsiento'] if row['CargoAbono'] == 'D' else None
        def haber(row):
            return row['ImporteAsiento'] if row['CargoAbono'] == 'H' else None

        control['D'] = control.apply(debe, axis=1)
        control['H'] = control.apply(haber, axis=1)
        control = control.drop(columns=['CargoAbono','ImporteAsiento'])
        control = control.groupby('FechaAsiento').agg({'D':'sum','H':'sum'}).reset_index()
        control['Control'] = control['D'] - control['H']

        with pd.ExcelWriter(new_excel, engine='openpyxl') as writer:
            Sheet1.to_excel(writer, sheet_name='Sheet1', index=False)
            control.to_excel(writer, sheet_name='Control', index=False)
            final_operaciones.to_excel(writer, sheet_name='Financiaciones', index=False)

        # Mover el puntero al inicio y devolver
        new_excel.seek(0)
        return new_excel

    except Exception as e:
        raise RuntimeError(f"Error al procesar el script: {str(e)}")
