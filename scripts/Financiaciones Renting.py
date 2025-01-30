import pdfplumber
import pandas as pd
import numpy as np
from datetime import datetime
import os
import re
from io import BytesIO

def procesar_pdfs_en_memoria(pdfs_dict):
    """
    Recibe un diccionario {nombre_pdf: BytesIO} con los PDFs
    y retorna un DataFrame consolidado tras extraer la info.
    """
    dataframes = []
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
    orden_nombres_final = list(mapeo_nombres.values())

    # Iterar por los PDFs subidos
    for pdf_name, pdf_buffer in pdfs_dict.items():
        try:
            with pdfplumber.open(pdf_buffer) as pdf:
                texto = "\n".join(
                    page.extract_text()
                    for page in pdf.pages
                    if page.extract_text()
                )
            
            # Extraer el código clave con expresión regular
            codigo_clave = re.search(r'\bE\d{2}[A-Z]{1}\d{8,}\b', texto)
            codigo_clave_extraido = codigo_clave.group(0) if codigo_clave else None

            lineas = texto.split("\n")
            df_lineas = pd.DataFrame(lineas, columns=["Contenido"])

            # Filtrar las filas cuyos inicios coincidan con los patrones_interes
            df_filtrado = df_lineas[
                df_lineas["Contenido"].str.startswith(tuple(patrones_interes), na=False)
            ].copy()

            # Extraer parte numérica al final de cada línea
            df_filtrado["Datos"] = df_filtrado["Contenido"].str.extract(r'([\d.,/]+)$')

            # Eliminar la parte numérica de la columna "Contenido"
            df_filtrado["Contenido"] = df_filtrado["Contenido"].str.replace(
                r'([\d.,/]+)$', '', regex=True
            ).str.strip()

            def cambiarnombres(texto_patron):
                for patron, nuevo_nombre in mapeo_nombres.items():
                    if texto_patron.startswith(patron):
                        return nuevo_nombre
                return texto_patron

            # Cambiar nombres (p.ej. "Tfno" -> "Fecha")
            df_filtrado["Contenido"] = df_filtrado["Contenido"].apply(cambiarnombres)

            if df_filtrado.empty:
                # Si no se encontró nada, creamos una fila vacía
                df_transpuesto = pd.DataFrame(columns=orden_nombres_final)
                df_transpuesto.loc[0] = [""] * len(orden_nombres_final)
            else:
                df_pivote = df_filtrado.set_index("Contenido")["Datos"]
                df_pivote = df_pivote.reindex(orden_nombres_final, fill_value="")
                df_transpuesto = df_pivote.to_frame().T.reset_index(drop=True)

            # Agregar la columna de operación y nombre del archivo
            df_transpuesto["Operación"] = codigo_clave_extraido
            df_transpuesto["Nombre_Archivo"] = pdf_name

            dataframes.append(df_transpuesto)
        
        except Exception as e:
            print(f"Error procesando {pdf_name}: {e}")
            # Si hay error, puedes decidir agregar un DataFrame vacío o ignorarlo

    if dataframes:
        df_final = pd.concat(dataframes, ignore_index=True)
    else:
        df_final = pd.DataFrame()

    return df_final

def main(files, pdfs, new_excel, month=None, year=None):
    """
    Función principal para 'Financiaciones Renting':
    - Recibe 'files' (no usado en este ejemplo, pero se deja para uniformidad),
    - Recibe 'pdfs': dict con {nombre_pdf: BytesIO},
    - Recibe 'new_excel': BytesIO donde escribimos el Excel de salida,
    - Retorna el BytesIO con el Excel final.
    """
    try:
        # 1) Procesar PDFs en memoria
        df_resultado = procesar_pdfs_en_memoria(pdfs)

        # Filtrar nulos o vacíos de la columna Fecha
        if not df_resultado.empty and 'Fecha' in df_resultado.columns:
            df_resultado.dropna(subset=['Fecha'], inplace=True)
            df_resultado = df_resultado[df_resultado['Fecha'].str.strip() != '']

        # 2) Escribir el df_resultado en una hoja Excel
        #    (Podrías hacer más transformaciones si quisieras)
        with pd.ExcelWriter(new_excel, engine='openpyxl') as writer:
            if not df_resultado.empty:
                df_resultado.to_excel(writer, sheet_name='FinanciacionesRenting', index=False)
            else:
                # Si no hay nada, creamos una hoja vacía
                pd.DataFrame({"Mensaje":["No se encontraron datos en los PDFs"]}).to_excel(
                    writer, sheet_name='FinanciacionesRenting', index=False
                )

        # Retornar el buffer
        new_excel.seek(0)
        return new_excel

    except Exception as e:
        raise RuntimeError(f"Error al procesar el script: {str(e)}")
