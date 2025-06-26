import pandas as pd
from io import BytesIO
from datetime import datetime

def main(files, pdfs=None, new_excel=None, month=None, year=None):

    try:
        # 1) Recuperar el archivo 'Stripe' del diccionario files
        if "Santnader" not in files or files["Santnader"] is None:
            raise RuntimeError("Falta el archivo Excel de Santnader (clave 'Santnader').")
        
        # 'files["Stripe"]' es un UploadedFile, lo convertimos a un buffer
        uploaded_file = files["Santnader"]
        
        Limpiar = pd.read_excel(archivo)

        Limpiar['Cuenta'] = Limpiar['Memo']
        def extraer_dni(texto):
          # Busca cualquiera de estos patrones:
          # 1) \d{8}[A-Za-z]       (8 dígitos y 1 letra)
          # 2) [A-Za-z]\d{8}       (1 letra y 8 dígitos)
          # 3) [A-Za-z]\d{7}[A-Za-z] (1 letra, 7 dígitos y 1 letra)
          patron = r'(\d{8}[A-Za-z]|[A-Za-z]\d{8}|[A-Za-z]\d{7}[A-Za-z])'

          match = re.search(patron, texto)
          if match:
            return match.group(1)
          return None

        def extraer_matricula(texto):

          patronmatricula = r'(\s*\d{4}\s*[A-Za-z]{3}\s*)'
          match = re.search(patronmatricula, texto)
          if match:
            return match.group(1)
          return None

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


        Limpiar.drop(columns=['Cuenta'], inplace=True)
        Limpiar.drop(columns=['Cuenta1'], inplace=True)
        Limpiar.drop(columns=['Cuenta2'], inplace=True)

        Limpiar.to_excel(ruta, index=False)
        output.seek(0)

        # 5) Retornar el BytesIO
        return output

    except Exception as e:
        # Si algo falla, levantamos una RuntimeError para que Streamlit la capte
        raise RuntimeError(f"Error al procesar el archivo: {str(e)}")
