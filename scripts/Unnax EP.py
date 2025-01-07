import pandas as pd
import re

def main(files, new_excel, month=None, year=None):
    try:
        #Limpiar informes de usuario y unirlos
        Limpiar_EP = pd.read_csv(Unnax)

        #Entrada mo
        entrada_movimientos = pd.read_excel(Compras)

        entrada_movimientos = entrada_movimientos.sort_values(by ='Fecha', ascending=False)
        entrada_movimientos = entrada_movimientos.loc[entrada_movimientos['Serie'].isin(['CV'])]

        Ordenar_EP = Limpiar_EP

        Ordenar_EP['Banco'] = 'Easy Payment'
        Ordenar_EP['Cuenta'] = Ordenar_EP['Cuenta'].astype(str)
        Ordenar_EP['Importe (cents)'] = Ordenar_EP['Importe  (cents)'].astype(float)
        Ordenar_EP['Importe (cents)'] = Ordenar_EP['Importe (cents)']/100

        #Saldo
        #Balance
        Ordenar_EP['Concepto'] = Ordenar_EP['Concepto'].astype(str)
        Ordenar_EP['Saldo'] = ''
        Ordenar_EP['Balance'] = ''
        Ordenar_EP['Proveedor'] = Ordenar_EP['Beneficiario']

        def extraer_matricula(texto):
            # Usamos una expresión regular para buscar la matrícula
            match = re.search(r'Pago de (C?\d{4}[A-Z]{3})Beneficiario', texto)
            if match:
                return match.group(1)
            return None

        Ordenar_EP['Matrícula'] = Ordenar_EP['Concepto'].apply(extraer_matricula)

        def columna_matricula(row):
            if row['Matrícula'] == None:
                return row['Concepto']
            else:
                return row['Matrícula']
    
        Ordenar_EP['Matrícula'] = Ordenar_EP.apply(columna_matricula, axis=1)

        def agregar_cuenta_contable(df_ordenar_cb, df_entrada_movimientos):
            df_entrada_movimientos = df_entrada_movimientos.drop_duplicates(subset='Código artículo', keep='last')
            mapping_series = df_entrada_movimientos.set_index('Código artículo')['Código proveedor']

            df_ordenar_cb['Código proveedor'] = df_ordenar_cb['Matrícula'].map(mapping_series)
    
            return df_ordenar_cb

        # Ejemplo de uso:
        Cuentacontable = agregar_cuenta_contable(Ordenar_EP, entrada_movimientos)
        Cuentacontable['Cuentacontable'] = Cuentacontable['Código proveedor'] + 400000000

        #ordenar por fecha
        Cuentacontable = Cuentacontable.sort_values(by ='F. Creación', ascending=True)
        columnas_SUMMARY= ['Banco', 'Cuenta', 'Importe (cents)','Beneficiario',
                   'Saldo' ,'Balance', 
                   'Matrícula' , 'F. Creación' ,	'F. Deposito', 
                   'F. Transferencia' ,	'Código de orden' ,	'Código de orden del banco', 	
                   'Cuenta Unnax'	, 'Proveedor' ,	'Cuentacontable']

        Easypayment = Cuentacontable[columnas_SUMMARY]
        
        new_excel = BytesIO()
        with pd.ExcelWriter(new_excel, engine='xlsxwriter')as writer:
            Easypayment.to_excel(writer, sheet_name= "Sheet 1")

        new_excel.seek(0)  # Reiniciar el puntero del buffer
        return new_excel  # Devuelve el archivo generado como BytesIO

    except Exception as e:
        raise RuntimeError(f"Error al procesar el script: {str(e)}")
