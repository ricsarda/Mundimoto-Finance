import pdfplumber
import pandas as pd
import numpy as np
from datetime import datetime
import os
import re

#fecha actual
fecha_actual = datetime.now()
fecha_venci = datetime.now().date()
fecha = fecha_actual.strftime("%d-%m-%Y")

# Rutas
ruta_archivo_final_excel = f"C:/Users/Ricardo Sarda/Desktop/Python/Financiaciones Renting/Financiaciones {fecha}.xlsx"
carpeta_pdfs = "C:/Users/Ricardo Sarda/Desktop/Python/Financiaciones Renting/"
ruta_Clientes = "C:/Users/Ricardo Sarda/Downloads/Libro_20241231_093303.xlsx" #clientes sage
ruta_ventas_SF = "C:/Users/Ricardo Sarda/Downloads/ES Sales - Invoiced this current year-2024-12-31-10-33-58.xlsx" #invoiced last 10 weeks solo detalles .xlsx

def procesar_pdfs_en_carpeta(carpeta_pdf):
    """Procesa todos los PDFs en una carpeta y convierte la información en un DataFrame."""
    dataframes = []  # Lista para almacenar los DataFrames de cada PDF

    # Patrones de interés en el texto del PDF (tal cual aparecen)
    patrones_interes = [
        "Tfno",
        "EntregaImporte",
        "InteresesDevengados",
        "Totalpara",
        "NuevoCapital"
    ]
    
    # Mapeo de patrones a los nombres de columna que queremos
    mapeo_nombres = {
        "Tfno": "Fecha",
        "EntregaImporte": "Importe",
        "InteresesDevengados": "Intereses",
        "Totalpara": "Total para Aplicar",
        "NuevoCapital": "Nuevo Capital Pendiente"
    }
    
    # Orden final que queremos en las filas/columnas tras renombrar
    orden_nombres_final = list(mapeo_nombres.values())  # ["Fecha", "Importe", "Intereses", "Total para Aplicar", "Nuevo Capital Pendiente"]

    # Iterar por todos los archivos PDF en la carpeta
    for archivo in os.listdir(carpeta_pdf):
        if archivo.endswith(".pdf"):
            ruta_pdf = os.path.join(carpeta_pdf, archivo)

            # Leer el contenido del PDF
            with pdfplumber.open(ruta_pdf) as pdf:
                texto = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])

            # Extraer el código clave
            codigo_clave = re.search(r'\bE\d{2}[A-Z]{1}\d{8,}\b', texto)
            codigo_clave_extraido = codigo_clave.group(0) if codigo_clave else None

            # Convertir el texto en un DataFrame por líneas
            lineas = texto.split("\n")
            df_lineas = pd.DataFrame(lineas, columns=["Contenido"])

            # Filtrar las filas cuyos inicios coincidan con alguno de los patrones
            df_filtrado = df_lineas[df_lineas["Contenido"].str.startswith(tuple(patrones_interes), na=False)].copy()

            # Extraer la parte numérica (Datos) al final de cada línea
            df_filtrado["Datos"] = df_filtrado["Contenido"].str.extract(r'([\d.,/]+)$')

            # Eliminar la parte numérica del final para quedarnos con el "encabezado" (Contenido)
            df_filtrado["Contenido"] = df_filtrado["Contenido"].str.replace(r'([\d.,/]+)$', '', regex=True).str.strip()

            # Función para renombrar el texto detectado al nombre deseado
            def cambiarnombres(texto_patron):
                # Si existe en el mapeo, devolvemos el valor nuevo, si no, lo dejamos tal cual
                for patron, nuevo_nombre in mapeo_nombres.items():
                    if texto_patron.startswith(patron):
                        return nuevo_nombre
                return texto_patron

            # Aplicamos el cambio de nombres (p.ej. "Tfno..." -> "Fecha")
            df_filtrado["Contenido"] = df_filtrado["Contenido"].apply(cambiarnombres)

            # Ahora queremos pivotar, poniendo el valor en "Datos" y el índice en "Contenido"
            # Pero antes forzamos a que existan todas las filas correspondientes a los patrones esperados
            # usando 'reindex' con fill_value=""
            
            # Si no hay coincidencias en absoluto (df_filtrado vacío), creamos un DF con NaN o blancos
            if df_filtrado.empty:
                df_transpuesto = pd.DataFrame(columns=orden_nombres_final)
                df_transpuesto.loc[0] = [""] * len(orden_nombres_final)  # fila en blanco
            else:
                df_pivote = df_filtrado.set_index("Contenido")["Datos"]
                df_pivote = df_pivote.reindex(orden_nombres_final, fill_value="")  # forzamos filas
                df_transpuesto = df_pivote.to_frame().T.reset_index(drop=True)

            # Agregamos la columna del código clave
            df_transpuesto["Operación"] = codigo_clave_extraido

            df_transpuesto["Nombre_Archivo"] = archivo
            # Agregar el DataFrame transpuesto a la lista
            dataframes.append(df_transpuesto)

    # Combinar todos los DataFrames en uno solo
    if dataframes:
        df_final = pd.concat(dataframes, ignore_index=True)
    else:
        df_final = pd.DataFrame()  # en caso de que no haya PDFs o no se extraiga nada

    return df_final


# Procesar los PDFs y generar el DataFrame final
df_resultado = procesar_pdfs_en_carpeta(carpeta_pdfs)
df_resultado.dropna(subset=['Fecha'], inplace=True)
df_resultado = df_resultado[df_resultado['Fecha'].str.strip() != '']
print(df_resultado)
# Guardar el resultado en un archivo Excel


