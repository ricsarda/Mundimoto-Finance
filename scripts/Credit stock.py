import pandas as pd
from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from io import BytesIO  # Para poder usar BytesIO
import traceback

def main(files, new_excel, pdfs=None, month=None, year=None):
    try:
        
        # fecha actual
        fecha_actual = datetime.now()
        fecha_venci = datetime.now().date()
        fecha = fecha_actual.strftime("%d-%m-%Y")

        # archivos
        metabase = pd.read_excel(files["Metabase"], engine="openpyxl")
        metabase['purchase_date'] = pd.to_datetime(metabase['purchase_date'], format='%d/%m/%Y', errors='coerce', dayfirst=True)
        metabase['registration_date'] = pd.to_datetime(metabase['registration_date'], format='%d/%m/%Y')

        Santander = pd.read_excel(files["Santander"], engine="openpyxl")
        Santander['Fecha Vencimiento'] = pd.to_datetime(Santander['Fecha Vencimiento'], format='%d/%m/%Y')
        Santander = Santander.sort_values(by='Fecha Vencimiento', ascending=True)
        Santander['license_plate'] = Santander['Matrícula']
        Santander['Bastiror '] = Santander['Bastidor '].str.upper()
        Santander['Importe Documentación'] = Santander['Importe Documentación'].astype(float)
        Santanderp = Santander.copy()

        Sabadell = pd.read_excel(files["Sabadell"], engine="xlrd")
        Sabadell['Fecha Vencimiento'] = pd.to_datetime(Sabadell['Fecha Vencimiento'], format='%d/%m/%Y')
        Sabadell['license_plate'] = Sabadell['Matrícula']
        Sabadell = Sabadell.loc[Sabadell['Estado'].isin(['Financed'])]
        Sabadellp = Sabadell.copy()

        Sofinco = pd.read_excel(files["Sofinco"], engine="openpyxl")
        Sofinco['End date'] = pd.to_datetime(Sofinco['End date'], format='%d/%m/%Y')
        Sofinco['Bastidor'] = Sofinco['VIN']
        Sofinco['Estado'] = Sofinco['Phase']
        Sofincop = Sofinco.copy()

        stock = metabase.loc[metabase['stock_status'].isin(['readyToMarket','onHold'])]
        Stock_disponible = stock['purchase_price'].sum()
        stock_libre = stock.loc[stock['santandersales'].isnull()]
        stock_libre = stock_libre[stock_libre['santandersales'] !='santanderSales']
        stock_libre = stock_libre[stock_libre['santanderrenting'] !='santanderSales']
        stock_libre = stock_libre[stock_libre['sabadellsales'] !='sabadellSales']
        stock_libre = stock_libre[stock_libre['sofincosales'] !='sofincoSales']
        stock_libre = stock_libre[stock_libre['wavi'] !='wavi']
        Stock_libre = stock_libre['purchase_price'].sum()


        #rescates Santander
        #seleccionar vencimientos
        mask = (Santander['Fecha Vencimiento'].dt.date == fecha_venci) | \
               (Santander['Fecha Vencimiento'].dt.date == (fecha_venci + timedelta(days=1))) | \
               (Santander['Fecha Vencimiento'].dt.date == (fecha_venci + timedelta(days=2))) | \
               (Santander['Fecha Vencimiento'].dt.date == (fecha_venci + timedelta(days=3))) | \
               (Santander['Fecha Vencimiento'].dt.date == (fecha_venci + timedelta(days=4))) | \
               (Santander['Fecha Vencimiento'].dt.date == (fecha_venci + timedelta(days=5)))

        vencimientos = Santander[mask]
        vencimiemtos_columnas = ['license_plate', 'Estado','Bastidor ','Fecha Vencimiento', 'Importe Documentación']
        vencimientos = vencimientos[vencimiemtos_columnas]

        #limpieza de pólizas normales
        Santanderp = Santanderp.merge(metabase[['license_plate','actual_credit_policy']],left_on='license_plate', right_on='license_plate', how='left')
        Sabadellp = Sabadellp.merge(metabase[['license_plate','actual_credit_policy']],left_on='license_plate', right_on='license_plate', how='left')
        Sofincop = Sofincop.merge(metabase[['frame_number','actual_credit_policy']],left_on='Bastidor', right_on='frame_number', how='left')
        columnas_santander = ['Nº póliza', 'Matrícula','Bastidor ','Marca/Modelo', 'Importe Documentación','Fecha Entrada Stock desde', 'Fecha Vencimiento', 'Estado', 'Estado de la Ficha Téc.','actual_credit_policy']
        columnas_Sabadell = ['Contrato', 'Bastidor', 'Matrícula','Marca', 'Modelo', 'Linea', 'Fecha Inicio', 'Fecha Vencimiento','Estado', 'Importe Financiado', 'Capital Pendiente', 'Contrato Recibido', 'actual_credit_policy']
        columnas_Sofinco = ['Contract', 'Phase', 'Financial plan', 'Asset type', 'Make', 'Invoice', 'VIN',  'Start date', 'End date','Amount', 'Estado', 'actual_credit_policy']
        Santanderp = Santanderp[columnas_santander]
        Sabadellp = Sabadellp[columnas_Sabadell]
        Sabadellp = Sabadellp.drop_duplicates(subset=['Bastidor'], keep='first')
        Sofincop = Sofincop[columnas_Sofinco]

        #rescate santander
        rescate_santander = metabase.loc[metabase['actual_credit_policy'].isin(['santanderSales'])]
        rescate_santander = rescate_santander.loc[rescate_santander['stock_status'].isin(['sold'])]
        rescate_santander = rescate_santander.loc[rescate_santander['productive_status'].isin(['delivered','readyToDeliver',])]
        rescate_santander = rescate_santander.merge(Santander[['license_plate', 'Bastidor ']], left_on='license_plate', right_on='license_plate', how='left')
        rescate_santander = rescate_santander.merge(Santander[['license_plate', 'Importe Documentación']], left_on='license_plate', right_on='license_plate', how='left')
        rescate_santander = rescate_santander.merge(Santander[['Matrícula', 'Fecha Vencimiento']], left_on='license_plate', right_on='Matrícula', how='left')
        rescate_santander = pd.concat([rescate_santander, vencimientos], axis=0)
        rescate_santander = rescate_santander.merge(Santander[['license_plate', 'Estado']],left_on='license_plate', right_on='license_plate', how='left')
        rescate_santander['Estado'] = rescate_santander['Estado_y']
        rescate_santander['Estado'] = rescate_santander['Estado'].fillna('Fuera de póliza')
        rescate_santander_columnas = ['license_plate', 'stock_status', 'productive_status','Bastidor ','Importe Documentación','Estado','Fecha Vencimiento']
        rescate_santander = rescate_santander[rescate_santander_columnas]
        rescate_santander['stockapp'] = rescate_santander.apply(lambda row: f"{row['license_plate']}#",axis=1)
        rescate_santander = rescate_santander.sort_values(by='Fecha Vencimiento',ascending=True)
        rescate_santander = rescate_santander.drop_duplicates(subset=['license_plate'], keep='first')

        #rescate Sabadell
        rescate_Sabadell = metabase.loc[metabase['actual_credit_policy'].isin(['SabadellSales'])]
        rescate_Sabadell = rescate_Sabadell.loc[rescate_Sabadell['stock_status'].isin(['sold'])]
        rescate_Sabadell = rescate_Sabadell.loc[rescate_Sabadell['productive_status'].isin(['delivered','readyToDeliver',])]
        rescate_Sabadell = rescate_Sabadell.merge(Sabadell[['license_plate', 'Estado']], left_on='license_plate', right_on='license_plate', how='left')
        rescate_Sabadell = rescate_Sabadell.merge(Sabadell[['license_plate', 'Bastidor']], left_on='license_plate', right_on='license_plate', how='left')
        rescate_Sabadell = rescate_Sabadell.merge(Sabadell[['license_plate', 'Importe Financiado']], left_on='license_plate', right_on='license_plate', how='left')
        rescate_Sabadell['stockapp'] = rescate_Sabadell.apply(lambda row: f"{row['license_plate']}#",axis=1)
        rescate_Sabadell_columnas = ['license_plate', 'stock_status', 'productive_status','Bastidor','Importe Financiado','Estado','stockapp']
        rescate_Sabadell = rescate_Sabadell[rescate_Sabadell_columnas]
        rescate_Sabadell['Estado'] = rescate_Sabadell['Estado'].fillna('Fuera de póliza')
        rescate_Sabadell = rescate_Sabadell.sort_values(by='Estado',ascending=True)

        #rescate Sofinco
        rescate_Sofinco = metabase.loc[metabase['actual_credit_policy'].isin(['SofincoSales'])]
        rescate_Sofinco = rescate_Sofinco.loc[rescate_Sofinco['stock_status'].isin(['sold'])]
        rescate_Sofinco = rescate_Sofinco.loc[rescate_Sofinco['productive_status'].isin(['delivered','readyToDeliver',])]
        rescate_Sofinco = rescate_Sofinco.merge(Sofinco[['Bastidor', 'Estado']], left_on='frame_number', right_on='Bastidor', how='left')
        rescate_Sofinco = rescate_Sofinco.merge(Sofinco[['Bastidor', 'Amount']], left_on='frame_number', right_on='Bastidor', how='left')
        rescate_Sofinco['stockapp'] = rescate_Sofinco.apply(lambda row: f"{row['license_plate']}#",axis=1)
        rescate_Sofinco_columnas = ['license_plate', 'stock_status', 'productive_status','frame_number','Amount','Estado','stockapp']
        rescate_Sofinco = rescate_Sofinco[rescate_Sofinco_columnas]
        rescate_Sofinco['Estado'] = rescate_Sofinco['Estado'].fillna('Fuera de póliza')
        rescate_Sofinco = rescate_Sofinco.sort_values(by='Estado',ascending=True)

        #motossinpoliza
        motosparadoc = metabase.loc[metabase['stock_status'].isin(['readyToMarket','onHold'])]
        motosparadoc = motosparadoc.loc[motosparadoc['actual_credit_policy'].isnull()]
        extrasdesab = motosparadoc.loc[motosparadoc['santandersales'].isin(['santanderSales'])]

        #motos libres para Santander
        motosparsantander = motosparadoc.loc[motosparadoc['santandersales'].isnull()]
        motosparsantander = motosparsantander.loc[motosparsantander['wavi'].isnull()]
        motosparsantander = motosparsantander.loc[motosparsantander['santanderrenting'].isnull()]
        motosparsantander = motosparsantander.loc[motosparsantander['sabadellsales'].isnull()]
        motosparsantander = motosparsantander.loc[motosparsantander['sofincosales'].isnull()]
        motosparsantander = motosparsantander.loc[motosparsantander['purchase_price'] > 1000]
        motosparsantander = motosparsantander.loc[motosparsantander['kilometers'] > 20]
        motosparsantander['CODIGO DEALER'] = 'B67377580'
        motosparsantander['NOMBRE DEALER'] = 'AJ MOTOR EUROPA, S.L.'
        motosparsantander['PRODUCTO'] = 'O'
        motosparsantander['NUM. OPERACION'] = 'ESET20225000402'
        motosparsantander['BASTIDOR'] = motosparsantander['frame_number']
        motosparsantander['IMPORTE'] = motosparsantander['purchase_price']
        motosparsantander['MONEDA'] = 'EUR'
        motosparsantander['MARCA'] = motosparsantander['brand']
        motosparsantander['MODELO'] = motosparsantander['model']
        motosparsantander['VERSION'] = motosparsantander.apply(lambda row: f"{row['MARCA']} {row['MODELO']}",axis=1)
        motosparsantander['MATRICULA'] = motosparsantander['license_plate']
        motosparsantander['FECHA MATRICULA'] = motosparsantander['registration_date']
        motosparsantander['FACTURA'] = motosparsantander['purchase_id']
        motosparsantander['FECHA FACTURA'] = motosparsantander['purchase_date']
        motosparsantander = motosparsantander.sort_values(by='FECHA FACTURA',ascending=False)

        #10 años
        def filtrar_antiguedadsan(df, columna_fecha, anos_antiguedad):
            # Obtener la fecha actual
            fecha_actual = datetime.now().date()
    
            # Calcular la fecha límite
            fecha_limite = fecha_actual - timedelta(days=anos_antiguedad * 365)
    
            # Filtrar los datos donde la fecha de registro sea superior a la fecha límite
            datos_filtrados = df[pd.to_datetime(df[columna_fecha]).dt.date > fecha_limite]
    
            return datos_filtrados

        motosparsantander = filtrar_antiguedadsan(motosparsantander, 'registration_date', 10)
        columnas_doc_santander =[ 'CODIGO DEALER', 'NOMBRE DEALER', 'PRODUCTO', 'NUM. OPERACION', 'BASTIDOR', 'IMPORTE', 'MONEDA', 'MARCA', 'MODELO', 'VERSION', 'MATRICULA', 'FECHA MATRICULA', 'FACTURA', 'FECHA FACTURA']
        motosparsantander = motosparsantander[columnas_doc_santander]
        motosparsantander['stockapp'] = motosparsantander.apply(lambda row: f"{row['MATRICULA']}#santanderSales",axis=1)
        motosparsantander = motosparsantander[~motosparsantander['MATRICULA'].isin(['7347MMT', '4624MNV', 'MAPK110', '2205LYR'])]
        maxsantander = motosparsantander['IMPORTE'].sum()

        #motoslibres para Sabadell
        motosparSabadell = motosparadoc.loc[motosparadoc['purchase_price'] > 2400]
        motosparSabadell = motosparSabadell.loc[motosparSabadell['kilometers'] > 50]
        motosparSabadell = motosparSabadell.sort_values(by='purchase_date',ascending=False)

        #Sabadell años
        def filtrar_antiguedadsab(df, columna_fecha, años_min, años_max):
            # Obtener la fecha actual
            fecha_actual = datetime.now().date()
    
            # Calcular las fechas límite
            fecha_limite_min = fecha_actual - timedelta(days=años_max * 365)
            fecha_limite_max = fecha_actual - timedelta(days=años_min * 365)
    
            # Filtrar los datos donde la fecha de registro esté entre las dos fechas límite
            datos_filtrados = df[(pd.to_datetime(df[columna_fecha]).dt.date > fecha_limite_min) & 
                                 (pd.to_datetime(df[columna_fecha]).dt.date <= fecha_limite_max)]

            return datos_filtrados

        #motos para Sabadell    
        motosparSabadell = filtrar_antiguedadsab(motosparSabadell, 'registration_date', 10, 17)
        motosparSabadell = pd.concat([motosparSabadell, extrasdesab], axis=0)
        motosparSabadell = motosparSabadell.sort_values(by='purchase_date',ascending=False)
        motosparSabadell = motosparSabadell.loc[motosparSabadell['sabadellsales'].isnull()]
        motosparSabadell = motosparSabadell.loc[motosparSabadell['sofincosales'].isnull()]
        motosparSabadell = motosparSabadell.drop_duplicates(subset=['license_plate'], keep='first')
        motosparSabadell = motosparSabadell.loc[motosparSabadell['purchase_price'] > 2490]
        motosparSabadell = motosparSabadell.loc[motosparSabadell['kilometers'] > 20]
        motosparSabadell['Marca'] = motosparSabadell['brand']
        motosparSabadell['Modelo'] = motosparSabadell['model']
        motosparSabadell['Name'] = motosparSabadell.apply(lambda row: f"{row['Marca']} {row['Modelo']}",axis=1)
        motosparSabadell['Matrícula'] = motosparSabadell['license_plate']
        motosparSabadell['Nº Bastidor'] = motosparSabadell['frame_number']
        motosparSabadell['kilometros'] = motosparSabadell['kilometers']
        motosparSabadell['Año'] = motosparSabadell['registration_date'].dt.year
        motosparSabadell['Precio compra'] = motosparSabadell['purchase_price']
        motosparSabadell = motosparSabadell.sort_values(by='Precio compra',ascending=False)
        columnas_doc_sabadeel = [ 'Name','Marca', 'Modelo', 'Matrícula', 'Nº Bastidor', 'kilometros', 'Año', 'Precio compra']
        motosparSabadell = motosparSabadell[columnas_doc_sabadeel]
        motosparSabadell['stockapp'] = motosparSabadell.apply(lambda row: f"{row['Matrícula']}#SabadellSales",axis=1)
        maxSabadell = motosparSabadell['Precio compra'].sum()


        #motos para wabi
        motosparadocwabi = metabase.loc[metabase['stock_status'].isin(['rented'])]
        motosparadocwabi = motosparadocwabi.loc[motosparadocwabi['actual_credit_policy'].isnull()]
        motosparwabi = motosparadocwabi.loc[motosparadocwabi['santandersales'].isnull()]
        motosparwabi = motosparwabi.loc[motosparwabi['wavi'].isnull()]
        motosparwabi = motosparwabi.loc[motosparwabi['sabadellsales'].isnull()]
        motosparwabi = motosparwabi.loc[motosparwabi['sofincosales'].isnull()]
        motosparwabi = motosparwabi.loc[motosparwabi['purchase_price'] > 1000]
        motosparwabi = motosparwabi.loc[motosparwabi['kilometers'] > 20]
        motosparwabi['CODIGO DEALER'] = 'B67377580'
        motosparwabi['NOMBRE DEALER'] = 'AJ MOTOR EUROPA, S.L.'
        motosparwabi['PRODUCTO'] = 'O'
        motosparwabi['NUM. OPERACION'] = 'ESET20235001800'
        motosparwabi['BASTIDOR'] = motosparwabi['frame_number']
        motosparwabi['IMPORTE'] = motosparwabi['purchase_price']
        motosparwabi['MONEDA'] = 'EUR'
        motosparwabi['MARCA'] = motosparwabi['brand']
        motosparwabi['MODELO'] = motosparwabi['model']
        motosparwabi['VERSION'] = motosparwabi.apply(lambda row: f"{row['MARCA']} {row['MODELO']}",axis=1)
        motosparwabi['MATRICULA'] = motosparwabi['license_plate']
        motosparwabi['FECHA MATRICULA'] = motosparwabi['registration_date']
        motosparwabi['FACTURA'] = motosparwabi['purchase_id']
        motosparwabi['FECHA FACTURA'] = motosparwabi['purchase_date']
        motosparwabi = motosparwabi.sort_values(by='FECHA FACTURA',ascending=False)
        motosparwabi = filtrar_antiguedadsan(motosparwabi, 'registration_date', 10)
        motosparwabi = motosparwabi[columnas_doc_santander]
        motosparwabi['stockapp'] = motosparwabi.apply(lambda row: f"{row['MATRICULA']}#wabi",axis=1)

        def filtrar_facturacion(df, columna_fecha, mes_antiguedad):
            # Obtener la fecha actual
            fecha_actual = datetime.now().date()
    
            # Calcular la fecha límite
            fecha_limite = fecha_actual - timedelta(days=mes_antiguedad * 30)
    
            # Filtrar los datos donde la fecha de registro sea superior a la fecha límite
            datos_filtrados = df[pd.to_datetime(df[columna_fecha]).dt.date > fecha_limite]
    
            return datos_filtrados

        #motos para Sofinco
        motosparSofinco = filtrar_facturacion(motosparsantander, 'FECHA FACTURA', 2)
        motosparSofinco = motosparSofinco.merge(metabase[['frame_number', 'kilometers']], left_on='BASTIDOR', right_on='frame_number', how='left')
        columnas_doc_Sofinco = ['MARCA', 'MODELO', 'BASTIDOR', 'kilometers', 'MATRICULA', 'FECHA MATRICULA', 'FACTURA', 'IMPORTE']
        motosparSofinco = motosparSofinco[columnas_doc_Sofinco]
        motosparSofinco['stockapp'] = motosparSofinco.apply(lambda row: f"{row['MATRICULA']}#SofincoSales",axis=1)
        maxSofinco = motosparSofinco['IMPORTE'].sum()

        #periodos
        periodos = [0, 30, 60, 90, 120, 150, 180]

        #Santander
        Santandertot = Santander[Santander['Estado de la Ficha Téc.'].isin(['Recibida fotocopia' , 'Recibida' , 'Solicitada'])]
        Santander = Santandertot.loc[Santandertot['Nº póliza'].isin([1019])]

        totalporperiodosan = {}

        for i in range(len(periodos) - 1):
            inicio_periodo = fecha_actual + timedelta(days=periodos[i])
            fin_periodo = fecha_actual + timedelta(days=periodos[i + 1])
    
            motos_en_periodo = Santander[
                (Santander['Fecha Vencimiento'] >= inicio_periodo) &
                (Santander['Fecha Vencimiento'] <= fin_periodo)
            ]
    
            total = motos_en_periodo['Importe Documentación'].sum()
    
            totalporperiodosan[f'Importe Docu ({periodos[i]}-{periodos[i+1]})'] = total

        Santander =  pd.DataFrame.from_dict(totalporperiodosan, orient='index', columns=['Total'])

        Santander = Santander.transpose()
        Santander['Disponible'] = maxsantander
        Santander['Póliza'] = 'Santander'

        #Wabi
        Wabi = Santandertot.loc[Santandertot['Nº póliza'].isin([1436])]

        totalporperiodowab = {}

        for i in range(len(periodos) - 1):
            inicio_periodo = fecha_actual + timedelta(days=periodos[i])
            fin_periodo = fecha_actual + timedelta(days=periodos[i + 1])
    
            motos_en_periodo = Wabi[
                (Wabi['Fecha Vencimiento'] >= inicio_periodo) &
                (Wabi['Fecha Vencimiento'] <= fin_periodo)
            ]
    
            total = motos_en_periodo['Importe Documentación'].sum()
    
            totalporperiodosan[f'Importe Docu ({periodos[i]}-{periodos[i+1]})'] = total

        Wabi =  pd.DataFrame.from_dict(totalporperiodosan, orient='index', columns=['Total'])

        Wabi = Wabi.transpose()
        Wabi['Póliza'] = 'Wabi'

        #Sabadell
        totalporperiodosab = {}

        for i in range(len(periodos) - 1):
            inicio_periodo = fecha_actual + timedelta(days=periodos[i])
            fin_periodo = fecha_actual + timedelta(days=periodos[i + 1])
    
            motos_en_periodo = Sabadell[
                (Sabadell['Fecha Vencimiento'] >= inicio_periodo) &
                (Sabadell['Fecha Vencimiento'] <= fin_periodo)
            ]
    
            total = motos_en_periodo['Importe Financiado'].sum()
    
            totalporperiodosab[f'Importe Docu ({periodos[i]}-{periodos[i+1]})'] = total

        Sabadell =  pd.DataFrame.from_dict(totalporperiodosab, orient='index', columns=['Total'])

        Sabadell = Sabadell.transpose()
        Sabadell['Disponible'] = maxSabadell
        Sabadell['Póliza'] = 'Sabadell'

        #Sofinco
        Sofinco = Sofinco[(Sofinco['Phase'] == 'Activo')]
        totalporperiodosof = {}

        for i in range(len(periodos) - 1):
            inicio_periodo = fecha_actual + timedelta(days=periodos[i])
            fin_periodo = fecha_actual + timedelta(days=periodos[i + 1])
    
            motos_en_periodo = Sofinco[
                (Sofinco['End date'] >= inicio_periodo) &
                (Sofinco['End date'] <= fin_periodo)
            ]
    
            total = motos_en_periodo['Amount'].sum()
    
            totalporperiodosof[f'Importe Docu ({periodos[i]}-{periodos[i+1]})'] = total

        Sofinco =  pd.DataFrame.from_dict(totalporperiodosof, orient='index', columns=['Total'])

        Sofinco = Sofinco.transpose()
        Sofinco['Disponible'] = maxSofinco
        Sofinco['Póliza'] = 'Sofinco'

        #control
        CreditStock = pd.concat([ Santander , Wabi, Sabadell , Sofinco], axis=0 , ignore_index=True)
        CreditStockcolumnas = ['Póliza', 'Importe Docu (0-30)', 'Importe Docu (30-60)', 'Importe Docu (60-90)', 'Importe Docu (90-120)', 'Importe Docu (120-150)', 'Importe Docu (150-180)', 'Disponible']
        CreditStock = CreditStock[CreditStockcolumnas]
        CreditStock['Stock Disponible'] = Stock_disponible
        CreditStock['Stock Libre'] = Stock_libre


        #cambio formato final de fecha
        metabase['purchase_date'] = pd.to_datetime(metabase['purchase_date']).dt.strftime('%d/%m/%Y')
        metabase['registration_date'] = pd.to_datetime(metabase['registration_date']).dt.strftime('%d/%m/%Y')
        Santanderp['Fecha Vencimiento'] = pd.to_datetime(Santanderp['Fecha Vencimiento']).dt.strftime('%d/%m/%Y')
        Sabadellp['Fecha Vencimiento'] = pd.to_datetime(Sabadellp['Fecha Vencimiento']).dt.strftime('%d/%m/%Y')
        Sofincop['End date'] = pd.to_datetime(Sofincop['End date']).dt.strftime('%d/%m/%Y')
        Sofincop['Start date'] = pd.to_datetime(Sofincop['Start date']).dt.strftime('%d/%m/%Y')
        rescate_santander['Fecha Vencimiento'] = pd.to_datetime(rescate_santander['Fecha Vencimiento']).dt.strftime('%d/%m/%Y')
        motosparsantander['FECHA MATRICULA'] = pd.to_datetime(motosparsantander['FECHA MATRICULA']).dt.strftime('%d/%m/%Y')
        motosparsantander['FECHA FACTURA'] = pd.to_datetime(motosparsantander['FECHA FACTURA']).dt.strftime('%d/%m/%Y')
        motosparwabi['FECHA MATRICULA'] = pd.to_datetime(motosparwabi['FECHA MATRICULA']).dt.strftime('%d/%m/%Y')
        motosparwabi['FECHA FACTURA'] = pd.to_datetime(motosparwabi['FECHA FACTURA']).dt.strftime('%d/%m/%Y')
        motosparSofinco['FECHA MATRICULA'] = pd.to_datetime(motosparSofinco['FECHA MATRICULA']).dt.strftime('%d/%m/%Y')

        
        new_excel = BytesIO()
        with pd.ExcelWriter(new_excel, engine='xlsxwriter')as writer:
            # Escribir cada DataFrame en una hoja diferente
            metabase.to_excel(writer, sheet_name='Metabase', index=False)
            Santanderp.to_excel(writer, sheet_name='Santander', index=False)
            rescate_santander.to_excel(writer, sheet_name='R.Santander', index=False)
            motosparsantander.to_excel(writer, sheet_name='Motos Santander', index=False)
            motosparwabi.to_excel(writer, sheet_name='Motos Wabi', index=False)
            Sabadellp.to_excel(writer, sheet_name='Sabadell', index=False)
            rescate_Sabadell.to_excel(writer, sheet_name='R.Sabadell', index=False)
            motosparSabadell.to_excel(writer, sheet_name='Motos Sabadell', index=False)
            Sofincop.to_excel(writer, sheet_name='Sofinco', index=False)
            rescate_Sofinco.to_excel(writer, sheet_name='R.Sofinco', index=False)
            motosparSofinco.to_excel(writer, sheet_name='Motos Sofinco', index=False)
            CreditStock.to_excel(writer, sheet_name='Control', index=False)

        new_excel.seek(0)  # Reiniciar el puntero del buffer
        return new_excel  # Devuelve el archivo generado como BytesIO

    except Exception as e:
        tb = traceback.format_exc()
        raise RuntimeError(f"Error al procesar el script:\n{str(e)}\n\nTraceback completo:\n{tb}")
