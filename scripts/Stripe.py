import pandas as pd
from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

#fecha actual
fecha_actual = datetime.now()
fecha_venci = datetime.now().date()
fecha = fecha_actual.strftime("%d-%m-%Y")

#Ruta archivo final excel
ruta_archivo_final_excel= f"C:/Users/Ricardo Sarda/Desktop/MM/Stripe/Stripe {fecha}.xlsx"

#archivos
anterior = "C:/Users/Ricardo Sarda/Desktop/MM/Stripe/Stripe 13-11-2024.xlsx" #archivo anterior
stripe = "C:/Users/Ricardo Sarda/Downloads/Conciliacin_detallada_de_transferencias_EUR_Del_2024-10-01_al_2024-10-31_Europe-Madrid.csv" #stripe (“Informes”,”Conciliación de transferencias” → En la fecha poner el mes a cerrar y seleccionar opción “download” en “Conciliación de transferencias”)
metabase = "C:/Users/Ricardo Sarda/Downloads/finance_raw_data_facturacion__flota_activa___2024-11-11T12_35_07.334766Z.xlsx" #metabase renting (Finance Raw Data Facturación "Flota Activa" )
sage = "C:/Users/Ricardo Sarda/Downloads/Libro_20241111_123348.xlsx" #clientes sage
salesforce = "C:/Users/Ricardo Sarda/Downloads/ES Sales - Invoiced last 10 weeks-2024-11-11-13-32-10.xlsx" #salesforce (ES Sales - Invoiced last 10 weeks) solo detalles
transacciones = "C:/Users/Ricardo Sarda/Downloads/payouts (5).csv" #saldos-transferencias

#igualar con la plantilla
datos1 = pd.read_csv(stripe, delimiter=',')
renting = pd.read_excel(metabase)
renting = renting.drop_duplicates(subset='mail')
clientes = pd.read_excel(sage)
ventas = pd.read_excel(salesforce)
ventas = ventas.drop_duplicates(subset='Correo')
comisiones = pd.read_csv(transacciones, delimiter=',')
anterior = pd.read_excel(anterior, sheet_name='DATOS T')
anterior = anterior.drop_duplicates(subset='customer_email')

#nombre
datos1['NOMBRE'] = datos1['customer_name']
def eliminar_filas_por_descripcion(datos):
    datos_filtrados = datos[~datos['description'].str.startswith(('Invoicing', 'Billing', 'Sigma'), na=False)]
    
    return datos_filtrados

datos = eliminar_filas_por_descripcion(datos1)

def motos_sin_matricula(row):
    if pd.isna(row['Moto']):
        return '-----'
    return row['Moto']
ventas['Moto'] = ventas.apply(motos_sin_matricula, axis=1)
#DNI
datos = datos.merge(renting[['mail', 'fiscalcode']], left_on='customer_email', right_on='mail', how='left')
datos = datos.merge(ventas[['Moto', 'DNI']], left_on='payment_metadata[license_plate]', right_on='Moto', how='left')
datos = datos.merge(anterior[['customer_email', 'DNI']], left_on='customer_email', right_on='customer_email', how='left')
rename = {'DNI_x': 'DNI'}
datos = datos.rename(columns=rename)
datos['DNI'] = datos['DNI'].fillna(datos['DNI_y'])
datos['DNI'] = datos['DNI'].fillna(datos['fiscalcode'])
datos['DNI'] = datos['DNI'].astype(str)
clientes['Cód. contable'] = clientes['Cód. contable'].str.replace(' ', '')
datos = datos.merge(clientes[['CIF/DNI', 'Cód. contable']], left_on='DNI', right_on='CIF/DNI', how='left')
datos['borrar repes'] = datos["customer_email"].astype(str) + " - " + datos["gross"].astype(str) + " - " + datos["automatic_payout_effective_at"].astype(str)
datos = datos.drop(columns=['fiscalcode', 'DNI_y','mail', 'CIF/DNI'])
datos = datos.drop_duplicates(subset='borrar repes')

def falta_dni(row):
    if pd.isna(row['DNI']):
        return "falta DNI"
    return row['Cód. contable']
datos['Cód. contable'] = datos.apply(falta_dni, axis=1)

#COMENTARIO
datos['Comentario'] = datos['customer_email']

#comisiones
comisiones['FechaAsiento'] = pd.to_datetime(comisiones['Created (UTC)'])
comisiones['FechaAsiento'] = comisiones['FechaAsiento'].dt.strftime('%d/%m/%Y')
comisiones['CodigoCuenta'] = 572000004
comisiones['ImporteAsiento'] = comisiones['Amount'].str.replace(',', '.').astype(float)
comisiones['Comentario'] = 'STRIPE'
comisiones['CargoAbono'] = 'D'
#sheet1
Sheet1 = datos[['automatic_payout_effective_at', 'Cód. contable', 'gross', 'Comentario']]

Sheet1['Asiento']= ''
Sheet1['FechaAsiento']= Sheet1['automatic_payout_effective_at']
Sheet1['FechaAsiento']= pd.to_datetime(Sheet1['FechaAsiento'])
Sheet1['FechaAsiento']= Sheet1['FechaAsiento'].dt.strftime('%d/%m/%Y')
Sheet1['CodigoCuenta']= Sheet1['Cód. contable']
Sheet1['ImporteAsiento']= Sheet1['gross']

def cargo_abono(row):
    if row['ImporteAsiento'] > 0:
        return 'H'
    elif row['ImporteAsiento'] < 0:
        return 'D'
    return ''
Sheet1['CargoAbono'] = Sheet1.apply(cargo_abono, axis=1)

#Ajuste de codigos contables
def reservas (row):
    if (row['ImporteAsiento'] == 250.00 or row['ImporteAsiento'] == -250.00) and pd.isna(row['CodigoCuenta']):
        return 555000003
    return row['CodigoCuenta']
Sheet1['CodigoCuenta'] = Sheet1.apply(reservas, axis=1)
Sheet1 = pd.concat([comisiones, Sheet1])

Sheet1['CodigoEmpresa']= 1
Sheet1['Ejercicio'] = 2024
Sheet1['MantenerAsiento']= 0
Sheet1['NumeroPeriodo']=-1
Sheet1 = Sheet1[[ 'CodigoEmpresa', 'Ejercicio', 'MantenerAsiento', 'NumeroPeriodo', 'Asiento', 'FechaAsiento', 'CargoAbono', 'CodigoCuenta', 'ImporteAsiento', 'Comentario']]
def faltantes(row):
    if pd.isna(row['CodigoCuenta']):
        return 555000004
    return row['CodigoCuenta']
Sheet1['CodigoCuenta'] = Sheet1.apply(faltantes, axis=1)
def menos(row):
    if row['ImporteAsiento'] < 0:
        return row['ImporteAsiento'] * -1
    return row['ImporteAsiento']
Sheet1['ImporteAsiento'] = Sheet1.apply(menos, axis=1)

#control y comisiones
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
control['Control'] = control['H'] - control['D']

#transformar comisiones
transformadas = control
transformadas['CodigoCuenta'] = 400025793
transformadas['ImporteAsiento'] = transformadas['Control']
transformadas['CargoAbono'] = 'D'
transformadas['Comentario'] = transformadas.apply(lambda row: f"COMISIONES {row['FechaAsiento']}", axis=1)

#finalizar sheet 1
Sheet1 = pd.concat([Sheet1, transformadas])
Sheet1['Asiento']= ''
Sheet1['Ejercicio'] = 2024
Sheet1['MantenerAsiento']= 0
Sheet1['NumeroPeriodo']=-1
Sheet1['CodigoEmpresa']= 1
Sheet1 = Sheet1[[ 'CodigoEmpresa', 'Ejercicio', 'MantenerAsiento', 'NumeroPeriodo', 'Asiento', 'FechaAsiento', 'CargoAbono', 'CodigoCuenta', 'ImporteAsiento', 'Comentario']]

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

Sheet1['ImporteAsiento'] = Sheet1['ImporteAsiento'].round(2)

new_row_2 = {'CodigoEmpresa': 'Datos básicos de movimiento','Ejercicio':'', 'MantenerAsiento':'', 'NumeroPeriodo':'', 'Asiento':'', 'FechaAsiento':'', 'CargoAbono':'', 'CodigoCuenta':'', 'ImporteAsiento':'', 'Comentario':''}
Sheet1 = pd.concat([Sheet1.iloc[:0], pd.DataFrame([new_row_2]), Sheet1.iloc[0:]]).reset_index(drop=True)

new_row_3 = {col: 'OBL' if col != 'Comentario' else '' for col in Sheet1.columns}
Sheet1 = pd.concat([Sheet1.iloc[:1], pd.DataFrame([new_row_3]), Sheet1.iloc[1:]]).reset_index(drop=True)

#archivo final excel
with pd.ExcelWriter(ruta_archivo_final_excel, engine='openpyxl') as writer:
     Sheet1.to_excel(writer, sheet_name='Sheet1', index=False)
     control.to_excel(writer, sheet_name='Control', index=False)
     datos1.to_excel(writer, sheet_name='DATOS', index=False)
     datos.to_excel(writer, sheet_name='DATOS T', index=False)
     clientes.to_excel(writer, sheet_name='CLIENTES', index=False)
     renting.to_excel(writer, sheet_name='RENTING', index=False)
     ventas.to_excel(writer, sheet_name='VENTAS', index=False)
     comisiones.to_excel(writer, sheet_name='COMISIONES', index=False)

