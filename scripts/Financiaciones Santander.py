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
    """
    Si hay columnas 'Extra 0' y 'Extra 1' con la fila Tipo=='115',
    reempaquetamos esos datos como 'Compensaciones'.
    """
    if 'Extra 0' in df.columns and 'Extra 1' in df.columns:
        extra_rows = df[df['Tipo'] == '115']
        for idx, row in extra_rows.iterrows():
            if pd.notna(row['Extra 0']) and pd.notna(row['Extra 1']):
                df.at[idx, 'Tipo'] = 'Compensaciones'
                df.at[idx, 'Datos'] = f"{row['Extra 0']} {row['Extra 1']}"
        df = df.drop(columns=['Extra 0', 'Extra 1'], errors='ignore')
    return df

def procesar_pdf_en_memoria(pdf_buffer):
    """
    Lee y procesa un PDF (contenido en un BytesIO).
    Devuelve un DataFrame reorganizado para la parte principal del script.
    """
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

        # Unir las columnas en 'Datos'
        def unir_datos(row):
            valores = [row['Datos'], row['Fecha'], row['Hora'], row['Importe'], row['Concepto']]
            return ' '.join(str(v) for v in valores if pd.notna(v))

        _operacion['Datos'] = _operacion.apply(unir_datos, axis=1)
        _operacion.drop(columns=["Fecha", "Hora", "Importe", "Concepto"], inplace=True, errors='ignore')

        # Excluir filas irrelevantes
        _operacion = _operacion[~_operacion['Tipo'].isin([
            'RELACION', '---------------------', 'Aténtamente,',
            'DEPARTAMENTOAUTOMOCION', 'SERVICIOPAGOAPRESCRIPTORES'
        ])]

        # Reemplazos
        condiciones = {
            '001': '001 PAGO AL PROVEEDOR ',
            '002': '001 ENTREGA INICIAL ',
            '041': '001 COMISION A TERCEROS ',
            '115': '002 Compensación Operación',
            'TOTAL': 'OPERACION: '
        }
        # Borramos prefijos de la columna 'Datos' según 'Tipo'
        def quitar_prefijo(row):
            if row['Tipo'] in condiciones:
                return row.replace(condiciones[row['Tipo']], '')
            return row
        _operacion = _operacion.apply(quitar_prefijo, axis=1)

        # Sustituciones en 'Tipo'
        _operacion['Tipo'] = _operacion['Tipo'].replace({
            'OPERACION:': 'OPERACION',
            'TITULAR:': 'TITULAR',
            '001': 'PAGO AL PROVEEDOR',
            '002': 'ENTREGA INICIAL',
            '041': 'COMISION A TERCEROS'
        })

        # Limpieza
        _operacion = _operacion.replace(
            to_replace=[r'\.', r' EUROS', r'-'],
            value='',
            regex=True
        )
        _operacion['Datos'] = _operacion['Datos'].replace(',', '.', regex=True)

        # Reorganización de excepciones
        _operacion = reorganizar_excepciones(_operacion)

        # Transformar el DataFrame en filas con (OPERACION, PAGO AL PROVEEDOR, TOTAL, etc.)
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

        # Última fila si queda
        if current_row:
            rows.append(current_row)

        # Convertir la lista de dict en DF
        reorganized_df = pd.DataFrame(rows)
        reorganized_df['Fecha'] = convertir_fecha(fecha_str)

        # Si falta 'ENTREGA INICIAL', creamos la col con '001 ENTREGA INICIAL 0'
        if 'ENTREGA INICIAL' not in reorganized_df:
            reorganized_df['ENTREGA INICIAL'] = '001 ENTREGA INICIAL 0'
        reorganized_df['ENTREGA INICIAL'] = reorganized_df['ENTREGA INICIAL'].fillna('001 ENTREGA INICIAL 0')

        return reorganized_df

    except Exception as e:
        print(f"Error al procesar PDF en memoria: {e}")
        return pd.DataFrame()

def main(files, pdfs, new_excel, month=None, year=None):
    """
    Función principal para “Financiaciones Santander”:
      - Lee PDFs desde el dict 'pdfs'.
      - Lee Excels “Ventas” y “Financiaciones” desde 'files' (keys: "Ventas", "Financiaciones").
      - Procesa y genera un Excel final en 'new_excel'.
      - Retorna el BytesIO con el Excel final.
    """
    try:
        # 1) Procesar todos los PDFs
        dataframes = []
        for pdf_name, pdf_buffer in pdfs.items():
            df_parcial = procesar_pdf_en_memoria(pdf_buffer)
            if not df_parcial.empty:
                dataframes.append(df_parcial)

        if dataframes:
            df_final = pd.concat(dataframes, ignore_index=True)
        else:
            df_final = pd.DataFrame()

        # Eliminar col “Extra” “Hemos” “que”, si existen
        for col in list(df_final.columns):
            if any(s in col for s in ['Extra', 'Hemos', 'que']):
                df_final.drop(columns=col, inplace=True, errors='ignore')

        # Limpieza final de strings
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

        # Convertir a float
        for c in ['PAGO AL PROVEEDOR','ENTREGA INICIAL','COMISION A TERCEROS','TOTAL']:
            if c in df_final.columns:
                df_final[c] = pd.to_numeric(df_final[c], errors='coerce')

        # Quitar filas donde 'TOTAL' sea NaN
        if 'TOTAL' in df_final.columns:
            df_final.dropna(subset=['TOTAL'], inplace=True)

        # 2) Construir "new_rows"
        compensaciones_cols = [col for col in df_final.columns if col.startswith('Compensacion_')]
        rows_nuevas = []

        for i, row in df_final.iterrows():
            # "Total"
            rows_nuevas.append({
                'FechaAsiento': row['Fecha'],
                'CargoAbono': 'D',
                'CodigoCuenta': '572000004',
                'ImporteAsiento': row['TOTAL'],
                'Comentario': f'FINANC. SANTANDER - {row.get("OPERACION","")}',
                'Utilidad': 'Total'
            })
            # Comisión a terceros
            if row.get('COMISION A TERCEROS', 0) > 0:
                rows_nuevas.append({
                    'FechaAsiento': row['Fecha'],
                    'CargoAbono': 'H',
                    'CodigoCuenta': '754000000',
                    'ImporteAsiento': row['COMISION A TERCEROS'],
                    'Comentario': f'FINANC. SANTANDER - {row.get("OPERACION","")}',
                    'Utilidad': 'Comision Terceros'
                })
            # Pago proveedor - entrega inicial
            pago_prov = row.get('PAGO AL PROVEEDOR', 0)
            ent_ini = row.get('ENTREGA INICIAL', 0)
            if (pago_prov > 0) or (ent_ini > 0):
                titular = row.get('TITULAR','99999999')  # si no existe titular, algo por defecto
                rows_nuevas.append({
                    'FechaAsiento': row['Fecha'],
                    'CargoAbono': 'H',
                    'CodigoCuenta': titular,
                    'ImporteAsiento': (pago_prov - ent_ini),
                    'Comentario': f'FINANC. SANTANDER - {row.get("OPERACION","")}',
                    'Utilidad': 'Pago Proveedor - Entrega Inicial'
                })

            # Compensaciones
            for colc in compensaciones_cols:
                if pd.notna(row[colc]):
                    rows_nuevas.append({
                        'FechaAsiento': row['Fecha'],
                        'CargoAbono': 'D',
                        'CodigoCuenta': '754000000',
                        'ImporteAsiento': row[colc],
                        'Comentario': str(row[colc]),
                        'Utilidad': 'Compensaciones'
                    })

        final_operaciones = pd.DataFrame(rows_nuevas)

        # 3) Leer "Financiaciones" y "Ventas" desde 'files'
        if "Financiaciones" in files and files["Financiaciones"] is not None:
            financiaciones = pd.read_excel(files["Financiaciones"], engine='openpyxl')
        else:
            financiaciones = pd.DataFrame()

        if "Ventas" in files and files["Ventas"] is not None:
            ventas_SF = pd.read_excel(files["Ventas"], engine='openpyxl')
        else:
            ventas_SF = pd.DataFrame()

        # (a) Separa “Compensaciones” y “Pago Proveedor - Entrega Inicial”
        compensaciones = final_operaciones[final_operaciones['Utilidad'] == 'Compensaciones']
        codigocliente = final_operaciones[final_operaciones['Utilidad'] == 'Pago Proveedor - Entrega Inicial']
        comisiones = final_operaciones[final_operaciones['Utilidad'] == 'Comision Terceros']
        # Merge con "Financiaciones" y "Ventas" para obtener info extra:
        codigocliente['Operación'] = codigocliente['Comentario'].str.replace('FINANC. SANTANDER - ', '', regex=False)
        codigocliente = codigocliente.merge(financiaciones[['Operación', 'MATRÍCULA']], on='Operación', how='left')
        codigocliente = codigocliente.merge(ventas_SF[['Moto', 'DNI']],right_on='Moto', left_on='MATRÍCULA', how='left')
        codigocliente['External ID'] = codigocliente['DNI']

        # Reemplazar 'CodigoCuenta' con 'External ID' si existe
        def accountidcliente(row):
            if pd.isna(row['External ID']):
                return row['CodigoCuenta']
            else:
                return row['External ID']
        codigocliente['External ID'] = codigocliente.apply(accountidcliente, axis=1)
        codigocliente = codigocliente[['FechaAsiento', 'External ID', 'ImporteAsiento', 'Operación']]

        # Reformatear comentarios de compensaciones
        def reformatear_comentario(comentario):
            partes = comentario.split()
            if len(partes) == 2:
                codigo, importe = partes
                codigo_formateado = f'FINANC. SANTANDER - E {codigo[1]} {codigo[2:4]} {codigo[4:8]} {codigo[8:]}'
                return codigo_formateado, importe
            return comentario, None

        for idx, rowc in compensaciones.iterrows():
            if rowc['Utilidad'] == 'Compensaciones':
                nuevo_coment, nuevo_importe = reformatear_comentario(rowc['Comentario'])
                if nuevo_importe:
                    compensaciones.at[idx, 'ImporteAsiento'] = float(nuevo_importe.replace(',', '.'))
                    compensaciones.at[idx, 'Comentario'] = nuevo_coment

        compensaciones['Account ID'] = 574000000
        # Concat final_operaciones + compensaciones
        final_operaciones = pd.concat([final_operaciones, compensaciones], ignore_index=True)

        # Añadir col “Fecha”, “Descripcion linea”, “Memo”
        final_operaciones['Fecha'] = final_operaciones['FechaAsiento']
        final_operaciones['Descripcion linea'] = final_operaciones['Comentario']
        final_operaciones['Memo'] = final_operaciones['Comentario']

        # Define credit/debit
        def credit(row):
            return row['ImporteAsiento'] if row['CargoAbono'] == 'H' else None
        def debit(row):
            return row['ImporteAsiento'] if row['CargoAbono'] == 'D' else None

        final_operaciones['Credit'] = final_operaciones.apply(credit, axis=1)
        final_operaciones['Debit'] = final_operaciones.apply(debit, axis=1)
        final_operaciones['ExternalID'] = final_operaciones['Utilidad'] + '_' + final_operaciones['Fecha'].astype(str)

        ordenfinal = ['ExternalID', 'Fecha', 'Memo', 'Account ID','Credit','Debit','Descripcion linea']
        final_operaciones = final_operaciones[ordenfinal]

        # Duplicar para contrapartida (si así lo haces en tu script)
        operaciones_contraparte = final_operaciones.copy()
        operaciones_contraparte.rename(columns={'Credit':'Debit','Debit':'Credit'}, inplace=True)
        operaciones_contraparte['Account ID'] = 572000004
        final_operaciones = pd.concat([final_operaciones, operaciones_contraparte], ignore_index=True)

        # Preparamos el Excel final en 'new_excel'
        with pd.ExcelWriter(new_excel, engine='openpyxl') as writer:
            final_operaciones.to_excel(writer, sheet_name='Import', index=False)
            codigocliente.to_excel(writer, sheet_name='Pago', index=False)
            financ_df.to_excel(writer, sheet_name='Financiaciones', index=False)
            ventas_df.to_excel(writer, sheet_name='Ventas SF', index=False)

        new_excel.seek(0)
        return new_excel

    except Exception as e:
        raise RuntimeError(f"Error al procesar el script: {str(e)}")

