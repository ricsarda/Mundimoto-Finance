import pandas as pd
import re

def main(files, pdfs=None, new_excel=None, month=None, year=None):
    try:
        # Verifica que se haya subido el archivo
        if "Extracto de Santander" not in files or files["Extracto de Santander"] is None:
            raise RuntimeError("Falta el archivo Excel de Santander (clave 'Extracto de Santander').")

        uploaded_file = files["Extracto de Santander"]
        Limpiar = pd.read_excel(uploaded_file)

        Limpiar['Cuenta'] = Limpiar['Memo']

        def extraer_dni(texto):
            patron = r'(\d{8}[A-Za-z]|[A-Za-z]\d{8}|[A-Za-z]\d{7}[A-Za-z])'
            match = re.search(patron, str(texto))
            return match.group(1) if match else None

        def extraer_matricula(texto):
            patronmatricula = r'(\s*\d{4}\s*[A-Za-z]{3}\s*)'
            match = re.search(patronmatricula, str(texto))
            return match.group(1) if match else None

        Limpiar['Cuenta1'] = Limpiar['Cuenta']
        Limpiar['Cuenta1'] = Limpiar['Cuenta1'].str.replace(r'CONCEPTO','',regex=True)
        Limpiar['Cuenta1'] = Limpiar['Cuenta1'].str.replace(r'MATRICULA','',regex=True)
        Limpiar['Cuenta1'] = Limpiar['Cuenta1'].str.replace(r'A LAS','',regex=True)
        Limpiar['Cuenta1'] = Limpiar['Cuenta1'].str.replace(r'INGRESO Anonimo CONTRA CUENTA EN ATM','',regex=True)
        Limpiar['Cuenta1'] = Limpiar['Cuenta1'].str.replace(r'EL','',regex=True)
        Limpiar['Cuenta1'] = Limpiar['Cuenta1'].str.replace(r'Dni','',regex=True)
        Limpiar['Cuenta2'] = Limpiar['Cuenta1'].str.replace(r'\d{4}\s*[A-Za-z]{3}', '',regex=True)

        Limpiar['DNI'] = Limpiar['Cuenta1'].apply(extraer_dni)
        Limpiar['Matrícula'] = Limpiar['Cuenta1'].apply(extraer_matricula)
        Limpiar['Matrícula'] = Limpiar['Matrícula'].str.replace(r'\s+', '', regex=True)

        Limpiar.drop(columns=['Cuenta', 'Cuenta1', 'Cuenta2'], inplace=True)

        # Si se pasa una ruta para guardar el Excel, lo guarda
        if new_excel:
            Limpiar.to_excel(new_excel, index=False)

        return Limpiar

    except Exception as e:
        raise RuntimeError(f"Error al procesar el archivo: {str(e)}")
