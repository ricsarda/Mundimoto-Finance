import pdfplumber
import pandas as pd
import re
from io import BytesIO
from datetime import datetime

def main(files, pdfs=None, new_excel=None, month=None, year=None):
    try:
        # Verificar si se han subido PDFs
        if not pdfs:
            raise RuntimeError("No se han subido archivos PDF.")

        # Obtener la fecha actual
        fecha_actual = datetime.now()
        fecha = fecha_actual.strftime("%d-%m-%Y")

        # Crear buffer para guardar el archivo Excel
        output = BytesIO()

        # Crear escritor de Excel en memoria
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            startrow = 0  # Fila inicial para empezar a escribir los datos en la hoja

            for pdf_name, pdf_file in pdfs.items():
                with pdfplumber.open(pdf_file) as pdf:
                    texto = "\n".join(page.extract_text() for page in pdf.pages if page.extract_text())

                lineas = texto.split('\n')

                # --- (1) Extraer FECHARECALCULO y código E31F...
                match_codigo = re.search(r'\bE\d{2}[A-Z]\d{8,}\b', texto)
                codigo_clave_extraido = match_codigo.group(0) if match_codigo else "COD NO ENCONTRADO"

                match_fecha = re.search(r"FECHARECALCULO\.\:\s+(\d{1,2}\/\d{1,2}\/\d{4})", texto)
                fecha_recal = match_fecha.group(1) if match_fecha else "FECHA NO ENCONTRADA"

                # --- (2) Recortar líneas de interés (ej. 6 a 36)
                lineas_deseadas = lineas[6:36]

                # --- (3) Convertir esas líneas en DataFrame
                filas_spliteadas = []
                for l in lineas_deseadas:
                    l = re.sub(r"^\*", "", l).strip()
                    columnas = l.split()
                    filas_spliteadas.append(columnas)

                df_resultado = pd.DataFrame(filas_spliteadas)

                # Evitar error si no hay filas suficientes
                if df_resultado.empty or df_resultado.shape[0] < 2:
                    continue

                # Asignar la 1ª fila como columnas
                df_resultado.columns = df_resultado.iloc[0]
                df_resultado = df_resultado.iloc[1:].reset_index(drop=True)

                # Comprobar columnas
                columnas_necesarias = ['FECHA', 'CAPITAL', 'PENDIENTE']
                if not set(columnas_necesarias).issubset(df_resultado.columns):
                    continue

                df_resultado = df_resultado[columnas_necesarias]

                # Renombrar columnas
                df_resultado.rename(columns={
                    'FECHA': 'Fecha',
                    'CAPITAL': 'Amort anticipada',
                    'PENDIENTE': 'Fee'
                }, inplace=True)

                # Indicar que la columna 'Fecha' sea índice y transponer
                df_resultado.set_index('Fecha', inplace=True)
                df_resultado = df_resultado.T

                # Elimina columnas 100% vacías
                df_resultado.dropna(axis=1, how='all', inplace=True)

                # --- (4) Escribir DataFrame en la misma hoja con separación
                sheet_name = "Datos Consolidados"

                # Escribir Código y Fecha Recalculo antes del DataFrame
                df_codigo_fecha = pd.DataFrame({
                    "A": [codigo_clave_extraido, fecha_recal]
                })
                df_codigo_fecha.to_excel(writer, sheet_name=sheet_name, startrow=startrow, startcol=0, header=False, index=False)

                # Escribir títulos de las columnas "Amort anticipada" y "Fee"
                df_titulos = pd.DataFrame({"A": ["Amort anticipada", "Fee"]})
                df_titulos.to_excel(writer, sheet_name=sheet_name, startrow=startrow + 4, startcol=0, header=False, index=False)

                # Escribir el DataFrame de la información del PDF
                df_resultado.to_excel(writer, sheet_name=sheet_name, startrow=startrow + 6, startcol=1, index=False)

                # Actualizar `startrow` para la próxima sección (dejando 5 filas de espacio entre PDFs)
                startrow += len(df_resultado) + 10

        # Guardar el archivo en el buffer
        output.seek(0)
        return output

    except Exception as e:
        raise RuntimeError(f"Error al procesar las financiaciones renting: {str(e)}")
