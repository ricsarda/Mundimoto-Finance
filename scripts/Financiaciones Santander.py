import pdfplumber
import pandas as pd
import numpy as np
from datetime import datetime
import os

#fecha actual
fecha_actual = datetime.now()
fecha_venci = datetime.now().date()
fecha = fecha_actual.strftime("%d-%m-%Y")

# Rutas
ruta_archivo_final_excel = f"C:/Users/Ricardo Sarda/Desktop/Python/Financiaciones santander/Financiaciones Santander {fecha}.xlsx"
ruta_carpeta_pdfs = "C:/Users/Ricardo Sarda/Desktop/Python/Financiaciones santander/"
ruta_Clientes = "C:/Users/Ricardo Sarda/Downloads/Libro_20241120_120403.xlsx" #clientes sage
ruta_ventas_SF = "C:/Users/Ricardo Sarda/Downloads/ES Sales - Invoiced last 10 weeks-2024-11-20-13-05-19.xlsx" #invoiced last 10 weeks solo detalles .xlsx

def convertir_fecha(fecha):
    meses = {"ENERO": "1", "FEBRERO": "2", "MARZO": "3", "ABRIL": "4", "MAYO": "5", "JUNIO": "6", "JULIO": "7", "AGOSTO": "8", "SEPTIEMBRE": "9", "OCTUBRE": "10", "NOVIEMBRE": "11", "DICIEMBRE": "12", "de": "/", ".": ""}
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

def procesar_pdf(ruta_PDF):
    try:
        with pdfplumber.open(ruta_PDF) as pdf:
            texto = "\n".join([pagina.extract_text() for pagina in pdf.pages if pagina.extract_text()])

        lineas = texto.split('\n')
        # Fecha
        linea_fecha = next((linea for linea in lineas if linea.startswith('Madrid,')), None)
        partes = linea_fecha.split('Madrid,', 1)
        fecha = partes[1].strip() if len(partes) > 1 else ''
        indice_inicio = next(i for i, linea in enumerate(lineas) if "OPERACION" in linea)
        datos_operacion = [linea.split() for linea in lineas[indice_inicio:] if linea.split()]

        max_cols = max(len(linea) for linea in datos_operacion)
        columnas = ["Tipo", "Datos", "Fecha", "Hora", "Importe", "Concepto"] + [f"Extra {i}" for i in range(max_cols - 6)]
        _operacion = pd.DataFrame(datos_operacion, columns=columnas[:max_cols]).replace([None, 'None'], np.nan).dropna(how='all')
        _operacion['Datos'] = _operacion.apply(lambda x: ' '.join(str(v) for v in [x['Datos'], x['Fecha'], x['Hora'], x['Importe'], x['Concepto']] if pd.notna(v)), axis=1)
        _operacion = _operacion.drop(columns=["Fecha", "Hora", "Importe", "Concepto"])
        _operacion = _operacion[~_operacion['Tipo'].isin(['RELACION', '---------------------', 'Aténtamente,', 'DEPARTAMENTOAUTOMOCION', 'SERVICIOPAGOAPRESCRIPTORES'])]

        condiciones = {'001': '001 PAGO AL PROVEEDOR ', '002': '001 ENTREGA INICIAL ', '041': '001 COMISION A TERCEROS ', '115': '002 Compensación Operación', 'TOTAL': 'OPERACION: '}
        _operacion = _operacion.apply(lambda x: x.replace(condiciones[x['Tipo']], '') if x['Tipo'] in condiciones else x, axis=1)

        _operacion['Tipo'] = _operacion['Tipo'].replace({'OPERACION:': 'OPERACION', 'TITULAR:': 'TITULAR', '001': 'PAGO AL PROVEEDOR', '002': 'ENTREGA INICIAL', '041': 'COMISION A TERCEROS'})
        
        _operacion = _operacion.replace(to_replace=[r'\.', r' EUROS', r'-'], value='', regex=True)
        _operacion['Datos'] = _operacion['Datos'].replace(',', '.', regex=True)

        # Reorganización de excepciones
        _operacion = reorganizar_excepciones(_operacion)

        # Transformar el DataFrame
        rows = []
        current_row = {}

        # Primero identificamos todas las posibles compensaciones
        compensaciones = [f'Compensacion_{i+1}' for i in range(len(_operacion[_operacion['Tipo'] == 'Compensaciones']))]

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

        # Asegurarse de añadir la última fila si está presente
        if current_row:
            rows.append(current_row)

        # Convertir las compensaciones en columnas
        for row in rows:
            if 'Compensaciones' in row:
                compensaciones_data = row.pop('Compensaciones')
                for i, comp in enumerate(compensaciones_data):
                    row[f'Compensacion_{i+1}'] = comp

        # Crear el DataFrame reorganizado
        reorganized_df = pd.DataFrame(rows)
        reorganized_df['Fecha'] = convertir_fecha(fecha)
        if 'ENTREGA INICIAL' not in reorganized_df:
            reorganized_df['ENTREGA INICIAL'] = '001 ENTREGA INICIAL 0'
        reorganized_df['ENTREGA INICIAL'] = reorganized_df.get('ENTREGA INICIAL', '0').fillna('001 ENTREGA INICIAL 0')
        return reorganized_df
    
    except Exception as e:
        print(f"Error al procesar {ruta_PDF}: {e}")
        return pd.DataFrame()

def extraer_datos_compensaciones(compensaciones):
    try:
        codigo, importe = compensaciones.split(' ', 1)
        codigo_formateado = ' '.join([codigo[:1], codigo[1:3], codigo[3:4], codigo[4:8], codigo[8:]])
        importe = float(importe.replace(',', '.'))
        return codigo_formateado, importe
    except Exception as e:
        print(f"Error al extraer datos de compensaciones: {e}")
        return None, None

archivos_pdf = [os.path.join(ruta_carpeta_pdfs, archivo) for archivo in os.listdir(ruta_carpeta_pdfs) if archivo.endswith('.PDF')]
dataframes = [procesar_pdf(archivo) for archivo in archivos_pdf if not procesar_pdf(archivo).empty]

# Concatenar todas las DataFrames reorganizadas
df_final = pd.concat(dataframes, ignore_index=True)
[df_final.drop(columns=col, inplace=True) for col in df_final.columns if 'Extra' in col]
[df_final.drop(columns=col, inplace=True) for col in df_final.columns if 'Hemos' in col]
[df_final.drop(columns=col, inplace=True) for col in df_final.columns if 'que' in col]
df_final['PAGO AL PROVEEDOR'] = df_final['PAGO AL PROVEEDOR'].str.replace('001 PAGO AL PROVEEDOR ', '')
df_final['PAGO AL PROVEEDOR'] = df_final['PAGO AL PROVEEDOR'].str.replace('002 PAGO AL PROVEEDOR ', '')
df_final['COMISION A TERCEROS'] = df_final['COMISION A TERCEROS'].str.replace('001 COMISION A TERCEROS ', '')
df_final['TOTAL'] = df_final['TOTAL'].str.replace('OPERACION: ', '')
df_final['ENTREGA INICIAL'] = df_final['ENTREGA INICIAL'].str.replace('001 ENTREGA INICIAL ', '')
df_final['ENTREGA INICIAL'] = df_final['ENTREGA INICIAL'].str.replace('002 ENTREGA INICIAL ', '')
print(df_final)
df_final['ENTREGA INICIAL'] = df_final['ENTREGA INICIAL'].astype(float)

# Rellenar las celdas vacías en la columna Operación con los datos de la celda superior
for columna in ['PAGO AL PROVEEDOR', 'ENTREGA INICIAL', 'COMISION A TERCEROS', 'TOTAL']:
     df_final[columna] = pd.to_numeric(df_final[columna], errors='coerce')

df_final = df_final.dropna(subset=['TOTAL'])
compensaciones_cols = [col for col in df_final.columns if col.startswith('Compensacion_')]
new_rows = []

for idx, row in df_final.iterrows():
        new_rows.append({
            'FechaAsiento': row['Fecha'],
            'CargoAbono': 'D',
            'CodigoCuenta': '572000004',
            'ImporteAsiento': row['TOTAL'],
            'Comentario': f'FINANC. SANTANDER - {row["OPERACION"]}',
            'Utilidad': 'Total',
        })
        if row['COMISION A TERCEROS'] > 0:
            new_rows.append({
                'FechaAsiento': row['Fecha'],
                'CargoAbono': 'H',
                'CodigoCuenta': '754000000',
                'ImporteAsiento': row['COMISION A TERCEROS'],
                'Comentario': f'FINANC. SANTANDER - {row["OPERACION"]}',
                'Utilidad': 'Comision Terceros',
            })
        if row['PAGO AL PROVEEDOR'] > 0 or row['ENTREGA INICIAL'] > 0:
            new_rows.append({
                'FechaAsiento': row['Fecha'],
                'CargoAbono': 'H',
                'CodigoCuenta': row['TITULAR'],
                'ImporteAsiento': row['PAGO AL PROVEEDOR'] - row['ENTREGA INICIAL'],
                'Comentario': f'FINANC. SANTANDER - {row["OPERACION"]}',
                'Utilidad': 'Pago Proveedor - Entrega Inicial',
            })
        # Iterar sobre cada fila del DataFrame
        for idx, row in df_final.iterrows():
            for col in compensaciones_cols:
                if pd.notna(row[col]):
                    new_rows.append({
                'FechaAsiento': row['Fecha'],
                'CargoAbono': 'D',
                'CodigoCuenta': '754000000',
                'ImporteAsiento': row[col],
                'Comentario': row[col],
                'Utilidad': 'Compensaciones',
                })

final_operaciones = pd.DataFrame(new_rows)
compensaciones = final_operaciones[final_operaciones['Utilidad'] == 'Compensaciones']
final_operaciones = final_operaciones[final_operaciones['Utilidad'] != 'Compensaciones']

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
        nuevo_comentario, nuevo_importe = reformatear_comentario(row['Comentario'])
        if nuevo_importe:
            compensaciones.at[idx, 'ImporteAsiento'] = nuevo_importe
            compensaciones.at[idx, 'Comentario'] = nuevo_comentario
compensaciones = compensaciones.drop_duplicates()
compensaciones['ImporteAsiento'] = compensaciones['ImporteAsiento'].str.replace(',', '.')
compensaciones['ImporteAsiento'] = compensaciones['ImporteAsiento'].astype(float)
final_operaciones = pd.concat([final_operaciones, compensaciones], ignore_index=True)

Sheet1 = final_operaciones[['FechaAsiento', 'CargoAbono', 'CodigoCuenta', 'ImporteAsiento', 'Comentario']]
Sheet1['CodigoEmpresa']= 1
Sheet1['Ejercicio'] = 2024
Sheet1['MantenerAsiento']= 0
Sheet1['NumeroPeriodo']=-1
Sheet1['Asiento']= ''
Sheet1 = Sheet1[[ 'CodigoEmpresa', 'Ejercicio', 'MantenerAsiento', 'NumeroPeriodo', 'Asiento', 'FechaAsiento', 'CargoAbono', 'CodigoCuenta', 'ImporteAsiento', 'Comentario']]

clientes = pd.read_excel(ruta_Clientes)
Sheet1['CodigoCuenta_lower'] = Sheet1['CodigoCuenta'].str.lower()
clientes['Razón social_lower'] = clientes['Razón social'].str.lower()

# Realizar el merge usando las columnas en minúsculas
Sheet1 = Sheet1.merge(clientes[['Razón social', 'Cód. contable', 'Razón social_lower']], 
                      left_on='CodigoCuenta_lower', 
                      right_on='Razón social_lower', 
                      how='left')

# Opcional: eliminar las columnas auxiliares después del merge
Sheet1.drop(columns=[ 'CodigoCuenta_lower' , 'Razón social_lower'], inplace=True)

def cod_contable(row):
    if pd.notna(row['Cód. contable']):
        return row['Cód. contable']
    return row['CodigoCuenta']

Sheet1['CodigoCuenta'] = Sheet1.apply(cod_contable, axis=1)
Sheet1.drop(columns=['Cód. contable','Razón social'], inplace=True)
Sheet1['ImporteAsiento'] = Sheet1['ImporteAsiento'].astype(float).round(2)

ventas_SF = pd.read_excel(ruta_ventas_SF)

control = Sheet1.groupby(['FechaAsiento', 'CargoAbono']).agg({'ImporteAsiento': 'sum'}).reset_index()
control = control.loc[control['CargoAbono'].isin(['D','H'])]

def debe(row):
    if row['CargoAbono'] == 'D':
        return row['ImporteAsiento']
    else:
        return None
    
def haber(row):
    if row['CargoAbono'] == 'H':
        return row['ImporteAsiento']
    else:
        return None
    
control['D'] = control.apply(debe, axis=1)
control['H'] = control.apply(haber, axis=1)

control = control.drop(columns=['CargoAbono', 'ImporteAsiento'])
control = control.groupby(['FechaAsiento']).agg({'D': 'sum', 'H': 'sum'}).reset_index()
control['Control'] = control['D'] - control['H']

# Formatear y actualizar la columna 'ImporteAsiento'
for index, row in Sheet1.iterrows():
    Sheet1.at[index, 'ImporteAsiento'] = f"{row['ImporteAsiento']} €"

new_row_2 = {'CodigoEmpresa': 'Datos básicos de movimiento','Ejercicio':'', 'MantenerAsiento':'', 'NumeroPeriodo':'', 'Asiento':'', 'FechaAsiento':'', 'CargoAbono':'', 'CodigoCuenta':'', 'ImporteAsiento':'', 'Comentario':''}
Sheet1 = pd.concat([Sheet1.iloc[:0], pd.DataFrame([new_row_2]), Sheet1.iloc[0:]]).reset_index(drop=True)

new_row_3 = {col: 'OBL' if col != 'Comentario' else '' for col in Sheet1.columns}
Sheet1 = pd.concat([Sheet1.iloc[:1], pd.DataFrame([new_row_3]), Sheet1.iloc[1:]]).reset_index(drop=True)

#archivo final excel
with pd.ExcelWriter(ruta_archivo_final_excel, engine='openpyxl') as writer:
     Sheet1.to_excel(writer, sheet_name='Sheet1', index=False)
     control.to_excel(writer, sheet_name='Control', index=False)
     final_operaciones.to_excel(writer, sheet_name='Financiaciones', index=False)
     ventas_SF.to_excel(writer, sheet_name='Ventas SF', index=False)
     clientes.to_excel(writer, sheet_name='Clientes SAGE', index=False)
