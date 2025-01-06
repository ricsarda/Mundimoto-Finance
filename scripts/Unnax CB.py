import pandas as pd
import re

def main(files, new_excel, month=None, year=None):
    try:
        #Limpiar informes de usuario y unirlos
        Limpiar_CB = pd.read_csv(unnax_CB)

        #Entrada mo
        entrada_movimientos = pd.read_excel(entr_mov)

        entrada_movimientos = entrada_movimientos.sort_values(by ='Fecha', ascending=False)
        entrada_movimientos = entrada_movimientos.loc[entrada_movimientos['Serie'].isin(['CV'])]

        Ordenar_CB = Limpiar_CB

        Ordenar_CB['Banco'] = 'CaixaBank-Empresas'
        Ordenar_CB['Cuenta'] = 'ES2121003709822200121095'
        Ordenar_CB['Importe (cents)'] = Ordenar_CB['Importe  (cents)'].astype(float)
        Ordenar_CB['Importe (cents)'] = Ordenar_CB['Importe (cents)']/100

        #Saldo
        #Balance
        Ordenar_CB['Concepto'] = Ordenar_CB['Concepto'].astype(str)
        Ordenar_CB['Saldo'] = ''
        Ordenar_CB['Balance'] = ''
        Ordenar_CB['Proveedor'] = Ordenar_CB['Beneficiario']

        def extraer_matricula(texto):
            # Usamos una expresión regular para buscar la matrícula
            match = re.search(r'Pago de (C?\d{4}[A-Z]{3})', texto)
            if match:
                return match.group(1)
            return None

        Ordenar_CB['Matrícula'] = Ordenar_CB['Concepto'].apply(extraer_matricula)

        def columna_matricula(row):
            if row['Matrícula'] == None:
                return row['Concepto']
            else:
                return row['Matrícula']
    
        Ordenar_CB['Matrícula'] = Ordenar_CB.apply(columna_matricula, axis=1)

        def agregar_cuenta_contable(df_ordenar_cb, df_entrada_movimientos):

            return df_ordenar_cb.merge(
                df_entrada_movimientos[['Código artículo', 'Código proveedor']],
                left_on='Matrícula',
                right_on='Código artículo',
                how='left'
            )

        # Ejemplo de uso:
        Cuentacontable = agregar_cuenta_contable(Ordenar_CB, entrada_movimientos)

        Cuentacontable['Cuentacontable'] = Cuentacontable['Código proveedor'] + 400000000

        #ordenar por fecha
        Cuentacontable = Cuentacontable.sort_values(by ='F. Creación', ascending=True)
        columnas_SUMMARY= ['Banco', 'Cuenta', 'Importe (cents)','Beneficiario',
                   'Saldo' ,'Balance', 'Matrícula' ,	'F. Creación' ,	'F. Deposito', 
                   'F. Transferencia' ,	'Código de orden' ,	'Código de orden del banco', 	
                   'Cuenta Unnax'	, 'Proveedor' ,	'Cuentacontable']
        CaixaBank = Cuentacontable[columnas_SUMMARY]
        
        new_excel = BytesIO()
        with pd.ExcelWriter(new_excel, engine='xlsxwriter')as writer:
            SUMMARY.to_excel(writer, sheet_name= "Sheet 1")

        new_excel.seek(0)  # Reiniciar el puntero del buffer
        return new_excel  # Devuelve el archivo generado como BytesIO

    except Exception as e:
        raise RuntimeError(f"Error al procesar el script: {str(e)}")
