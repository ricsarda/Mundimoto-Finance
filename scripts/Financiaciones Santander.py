# scripts/financiaciones_renting.py

import pdfplumber
import pandas as pd
import numpy as np
from datetime import datetime
import re
from io import BytesIO

def convertir_fecha(fecha_str):
    meses = {
        "ENERO": "1", "FEBRERO": "2", "MARZO": "3", "ABRIL": "4",
        "MAYO": "5", "JUNIO": "6", "JULIO": "7", "AGOSTO": "8",
        "SEPTIEMBRE": "9", "OCTUBRE": "10", "NOVIEMBRE": "11",
        "DICIEMBRE": "12", "de": "/", ".": ""
    }
    for es, en in meses.items():
        fecha_str = fecha_str.replace(es, en)
    return ' '.join(fecha_str.split()[0:3])

def reorganizar_excepciones(df):
    """
    Ajusta las filas marcadas como '115' para mover sus datos a filas 'Compensaciones'.
    """
    if 'Extra 0' in df.columns and 'Extra 1' in df.columns:
        extra_rows = df[df['Tipo'] == '115']
        for idx, row in extra_rows.iterrows():
            if pd.notna(row['Extra 0']) and pd.notna(row['Extra 1']):
                df.at[idx, 'Tipo'] = 'Compensaciones'
                df.at[idx, 'Datos'] = f"{row['Extra 0']} {row['Extra 1']}"
        df.drop(columns=['Extra 0', 'Extra 1'], inplace=True)
    return df

def procesar_pdf(pdf_buffer):
    """
    Toma un BytesIO de un PDF. Parsea su contenido y devuelve un DataFrame 'reorganizado'.
    """
    try:
        with pdfplumber.open(pdf_buffer) as pdf:
            texto = "\n".join(
                pagina.extract_text()
                for pagina in pdf.pages
                if pagina.extract_text()
            )

        lineas = texto.split('\n')
        # Buscar primera línea que empiece con "Madrid,"
        linea_fecha = next((l for l in lineas if l.startswith('Madrid,')), None)
        if not linea_fecha:
            # No se encontró 'Madrid,'; forzamos una cadena vacía
            fecha_str = ""
        else:
            partes = linea_fecha.split('Madrid,', 1)
            fecha_str = partes[1].strip() if len(partes) > 1 else ""

        # Buscar índice donde aparezca "OPERACION"
        indice_inicio = None
        for i, linea in enumerate(lineas):
            if "OPERACION" in linea:
                indice_inicio = i
                break
        if indice_inicio is None:
            # No se encontró "OPERACION" => no hay datos
            return pd.DataFrame()

        datos_operacion = [l.split() for l in lineas[indice_inicio:] if l.split()]
        max_cols = max(len(l) for l in datos_operacion)
        columnas = (
            ["Tipo", "Datos", "Fecha", "Hora", "Importe", "Concepto"]
            + [f"Extra {i}" for i in range(max_cols - 6)]
        )

        _operacion = (
            pd.DataFrame(datos_operacion, columns=columnas[:max_cols])
            .replace([None, 'None'], np.nan)
            .dropna(how='all')
        )
        # Unir varias columnas en 'Datos'
        _operacion['Datos'] = _operacion.apply(
            lambda x: ' '.join(
                str(v) for v in [
                    x['Datos'], x['Fecha'], x['Hora'], x['Importe'], x['Concepto']
                ] if pd.notna(v)
            ),
            axis=1
        )
        _operacion.drop(columns=["Fecha", "Hora", "Importe", "Concepto"], inplace=True, errors='ignore')

        # Remover filas irrelevantes
        _operacion = _operacion[
            ~_operacion['Tipo'].isin([
                'RELACION', '---------------------', 'Aténtamente,',
                'DEPARTAMENTOAUTOMOCION', 'SERVICIOPAGOAPRESCRIPTORES'
            ])
        ]

        # Reemplazos
        condiciones = {
            '001': '001 PAGO AL PROVEEDOR ',
            '002': '001 ENTREGA INICIAL ',
            '041': '001 COMISION A TERCEROS ',
            '115': '002 Compensación Operación',
            'TOTAL': 'OPERACION: '
        }
        def reemplazar_condiciones(row):
            if row['Tipo'] in condiciones:
                return row.replace(condiciones[row['Tipo']], '')
            return row

        _operacion = _operacion.apply(reemplazar_condiciones, axis=1)

        # Ajustar 'Tipo'
        _operacion['Tipo'] = _operacion['Tipo'].replace({
            'OPERACION:': 'OPERACION',
            'TITULAR:': 'TITULAR',
            '001': 'PAGO AL PROVEEDOR',
            '002': 'ENTREGA INICIAL',
            '041': 'COMISION A TERCEROS'
        })

        _operacion.replace(
            to_replace=[r'\.', r' EUROS', r'-'],
            value='',
            regex=True,
            inplace=True
        )
        _operacion['Datos'] = _operacion['Datos'].replace(',', '.', regex=True)

        _operacion = reorganizar_excepciones(_operacion)

        # Transformar el DataFrame
        rows = []
        current_row = {}
        for _, fila in _operacion.iterrows():
            if fila['Tipo'] == 'OPERACION':
                # Volcamos la fila actual si existe
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

        # Añadir la última si hubiera quedado algo sin volcar
        if current_row:
            rows.append(current_row)

        # Convertir 'Compensaciones' en columnas
        for r in rows:
            if 'Compensaciones' in r:
                lista_comp = r.pop('Compensaciones')
                for i, comp in enumerate(lista_comp):
                    r[f'Compensacion_{i+1}'] = comp

        reorganized_df = pd.DataFrame(rows)
        # Ajustar la fecha
        reorganized_df['Fecha'] = convertir_fecha(fecha_str)

        # Si no existe ENTREGA INICIAL, la creamos con '0'
        if 'ENTREGA INICIAL' not in reorganized_df.columns:
            reorganized_df['ENTREGA INICIAL'] = '001 ENTREGA INICIAL 0'
        reorganized_df['ENTREGA INICIAL'] = reorganized_df['ENTREGA INICIAL'].fillna('001 ENTREGA INICIAL 0')

        return reorganized_df

    except Exception as e:
        print(f"Error al procesar PDF en memoria: {e}")
        return pd.DataFrame()

def main(files, pdfs, new_excel, month=None, year=None):
    """
    Función principal para “Financiaciones Renting”:
    
    Recibe:
     - files: dict con {nombre: BytesIO} de Excels (“financiaciones”, “ventas_SF”, etc.)
     - pdfs: dict con {nombre: BytesIO} de PDFs
     - new_excel: BytesIO donde se guardará el resultado
     - month, year (opcionales)

    Devuelve:
     - new_excel (BytesIO) con las hojas:
         1) “Import” (final_operaciones)
         2) “Pago” (codigocliente)
         3) “Financiaciones 2025” (o el DF “financiaciones”)
         4) “Ventas SF” (o el DF “ventas_SF”)
    """
    try:
        # 1) Procesar todos los PDFs
        dataframes = []
        for pdf_name, pdf_buffer in pdfs.items():
            df_temp = procesar_pdf(pdf_buffer)
            if not df_temp.empty:
                dataframes.append(df_temp)

        if dataframes:
            df_final = pd.concat(dataframes, ignore_index=True)
        else:
            df_final = pd.DataFrame()

        # Eliminamos columnas indeseadas
        for col in list(df_final.columns):
            if 'Extra' in col or 'Hemos' in col or 'que' in col:
                df_final.drop(columns=col, inplace=True, errors='ignore')

        # Limpieza de cadenas
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
            df_final['ENTREGA INICIAL'] = pd.to_numeric(df_final['ENTREGA INICIAL'], errors='coerce').fillna(0)

        # Convertir a numérico las columnas relevantes
        for col in ['PAGO AL PROVEEDOR', 'ENTREGA INICIAL', 'COMISION A TERCEROS', 'TOTAL']:
            if col in df_final.columns:
                df_final[col] = pd.to_numeric(df_final[col], errors='coerce').fillna(0)

        # Quitar filas si 'TOTAL' es NaN o 0
        if 'TOTAL' in df_final.columns:
            df_final = df_final[df_final['TOTAL'] != 0]

        # 2) Construir la lista de dicts (new_rows)
        compensaciones_cols = [c for c in df_final.columns if c.startswith('Compensacion_')]
        new_rows = []

        for _, row in df_final.iterrows():
            # Fila principal: D 572000004
            new_rows.append({
                'FechaAsiento': row.get('Fecha', ''),
                'CargoAbono': 'D',
                'CodigoCuenta': '572000004',
                'ImporteAsiento': row.get('TOTAL', 0),
                'Comentario': f'FINANC. SANTANDER - {row.get("OPERACION", "")}',
                'Utilidad': 'Total',
            })
            # COMISION A TERCEROS
            if row.get('COMISION A TERCEROS', 0) > 0:
                new_rows.append({
                    'FechaAsiento': row['Fecha'],
                    'CargoAbono': 'H',
                    'CodigoCuenta': '754000000',
                    'ImporteAsiento': row['COMISION A TERCEROS'],
                    'Comentario': f'FINANC. SANTANDER - {row["OPERACION"]}',
                    'Utilidad': 'Comision Terceros',
                })
            # Pago Proveedor - Entrega Inicial
            pago_prov = row.get('PAGO AL PROVEEDOR', 0)
            entrega_inicial = row.get('ENTREGA INICIAL', 0)
            titular = row.get('TITULAR', '99999999')  # fallback si no existe
            if pago_prov > 0 or entrega_inicial > 0:
                new_rows.append({
                    'FechaAsiento': row['Fecha'],
                    'CargoAbono': 'H',
                    'CodigoCuenta': titular,
                    'ImporteAsiento': (pago_prov - entrega_inicial),
                    'Comentario': f'FINANC. SANTANDER - {row.get("OPERACION", "")}',
                    'Utilidad': 'Pago Proveedor - Entrega Inicial',
                })
            # Compensaciones
            for col_comp in compensaciones_cols:
                val = row.get(col_comp, None)
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

        # Capa: leer “financiaciones” y “ventas_SF” si existen en 'files'
        # (Si no existen, se crean DF vacíos)
        financiaciones_df = pd.DataFrame()
        ventas_SF_df = pd.DataFrame()

        if 'financiaciones' in files:
            try:
                financiaciones_df = pd.read_excel(files['financiaciones'], engine='openpyxl')
            except Exception as e:
                print(f"Error leyendo financiaciones: {e}")

        if 'ventas_SF' in files:
            try:
                ventas_SF_df = pd.read_excel(files['ventas_SF'], engine='openpyxl')
            except Exception as e:
                print(f"Error leyendo ventas_SF: {e}")

        # Creamos 'compensaciones' y 'codigocliente' a partir de final_operaciones
        compensaciones = final_operaciones[final_operaciones['Utilidad'] == 'Compensaciones'].copy()
        codigocliente = final_operaciones[final_operaciones['Utilidad'] == 'Pago Proveedor - Entrega Inicial'].copy()

        # Eliminamos 'Pago Proveedor - Entrega Inicial' de final_operaciones, 
        # y dejamos Comision Terceros + Total
        final_operaciones = final_operaciones[~final_operaciones['Utilidad'].isin(['Pago Proveedor - Entrega Inicial', 'Compensaciones'])]

        # Merge en codigocliente con financiaciones y ventas_SF
        codigocliente['Operación'] = codigocliente['Comentario'].str.replace('FINANC. SANTANDER - ', '', regex=False)
        if not financiaciones_df.empty:
            codigocliente = codigocliente.merge(financiaciones_df[['Operación', 'MATRÍCULA']], on='Operación', how='left')
        if not ventas_SF_df.empty:
            codigocliente = codigocliente.merge(ventas_SF_df[['Moto', 'DNI']], right_on='Moto', left_on='MATRÍCULA', how='left')
            codigocliente['External ID'] = codigocliente['DNI']

            # Si External ID es NaN, usar 'CodigoCuenta'
            def accountidcliente(r):
                if pd.isna(r['External ID']):
                    return r['CodigoCuenta']
                else:
                    return r['External ID']

            codigocliente['External ID'] = codigocliente.apply(accountidcliente, axis=1)

        codigocliente = codigocliente[['FechaAsiento', 'External ID', 'ImporteAsiento', 'Operación']]

        # Reformatear compensaciones
        def reformatear_comentario(comentario):
            partes = str(comentario).split()
            if len(partes) == 2:
                codigo, importe = partes
                codigo_formateado = f'FINANC. SANTANDER - E {codigo[1]} {codigo[2:4]} {codigo[4:8]} {codigo[8:]}'
                return codigo_formateado, importe
            return comentario, None

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

        # Concatenar compensaciones a final_operaciones
        final_operaciones = pd.concat([final_operaciones, compensaciones], ignore_index=True)

        # Añadir columnas finales (Fecha, Descripcion, etc.)
        final_operaciones['Fecha'] = final_operaciones['FechaAsiento']
        final_operaciones['Descripcion linea'] = final_operaciones['Comentario']
        final_operaciones['Memo'] = final_operaciones['Comentario']

        # Calcular Credit/Debit
        def credit(r):
            return r['ImporteAsiento'] if r['CargoAbono'] == 'H' else None

        def debit(r):
            return r['ImporteAsiento'] if r['CargoAbono'] == 'D' else None

        final_operaciones['Credit'] = final_operaciones.apply(credit, axis=1)
        final_operaciones['Debit'] = final_operaciones.apply(debit, axis=1)
        final_operaciones['ExternalID'] = final_operaciones['Utilidad'] + '_' + final_operaciones['FechaAsiento'].astype(str)

        # Columnas finales en orden
        ordenfinalcolumnas = [
            'ExternalID', 'Fecha', 'Memo', 'CodigoCuenta', 'Credit', 'Debit', 'Descripcion linea'
        ]
        final_operaciones = final_operaciones.reindex(columns=ordenfinalcolumnas, fill_value="")

        # Duplicamos con “contraparte”: invertimos Credit y Debit
        operaciones_contraparte = final_operaciones.copy()
        operaciones_contraparte.rename(columns={'Credit': 'Debit', 'Debit': 'Credit'}, inplace=True)
        # Account ID genérico? (Lo tenías en 2437, ajusta si procede)
        # Pero ojo que has usado 'CodigoCuenta' en la tabla, no 'Account ID' en la versión final.
        # Ajusta si lo necesitas.

        # Unir
        final_operaciones = pd.concat([final_operaciones, operaciones_contraparte], ignore_index=True)

        # Generar Excel final
        with pd.ExcelWriter(new_excel, engine='openpyxl') as writer:
            final_operaciones.to_excel(writer, sheet_name='Import', index=False)
            codigocliente.to_excel(writer, sheet_name='Pago', index=False)

            # Guardar financiaciones y ventas_SF si no están vacíos
            if not financiaciones_df.empty:
                financiaciones_df.to_excel(writer, sheet_name='Financiaciones 2025', index=False)
            if not ventas_SF_df.empty:
                ventas_SF_df.to_excel(writer, sheet_name='Ventas SF', index=False)

        new_excel.seek(0)
        return new_excel

    except Exception as e:
        raise RuntimeError(f"Error al procesar el script: {str(e)}")
