import pdfplumber
import pandas as pd
import re
from io import BytesIO
from datetime import datetime

# === Función auxiliar para convertir texto con punto de miles y coma decimal a float
def convertir_a_float(valor_str):
    valor_str = valor_str.replace('.', '').replace(',', '.')
    try:
        return float(valor_str)
    except ValueError:
        return None

# === Extrae datos de "Financiaciones"
def extraer_financiaciones(texto, lineas, nombre_pdf):
    # Expresiones regulares
    pat_fecha                   = re.compile(r"Fecha:([0-9]{2}/[0-9]{2}/[0-9]{2})")
    pat_operacion               = re.compile(r"operación\s*nº\s*([A-Za-z0-9]+)")  
    pat_entrega_importe         = re.compile(r"EntregaImporte\s+([\d.,]+)")
    pat_intereses               = re.compile(r"InteresesDevengados.*?([\d.,]+)$")
    pat_total_para_aplicar      = re.compile(r"TotalparaAplicaraCapital\s+([\d.,]+)")
    pat_nuevo_capital_pendiente = re.compile(r"NuevoCapitalPendiente\s+([\d.,]+)")

    # Criterio mínimo para considerar Financiaciones
    if ('operación nº' not in texto) and ('EntregaImporte' not in texto):
        return None

    info_pdf = {
        'Fecha': None,
        'Importe': None,
        'Intereses': None,
        'Total para Aplicar': None,
        'Nuevo Capital Pendiente': None,
        'Operación': None,
        'Archivo': nombre_pdf  # nombre sin .pdf
    }

    for linea in lineas:
        # Fecha
        if 'Fecha:' in linea:
            m = pat_fecha.search(linea)
            if m:
                info_pdf['Fecha'] = m.group(1)

        # Operación
        if 'operación nº' in linea:
            m = pat_operacion.search(linea)
            if m:
                info_pdf['Operación'] = m.group(1)

        # Importe
        if 'EntregaImporte' in linea:
            m = pat_entrega_importe.search(linea)
            if m:
                info_pdf['Importe'] = convertir_a_float(m.group(1))

        # Intereses
        if 'InteresesDevengados' in linea:
            m = pat_intereses.search(linea)
            if m:
                info_pdf['Intereses'] = convertir_a_float(m.group(1))

        # Total para Aplicar
        if 'TotalparaAplicaraCapital' in linea:
            m = pat_total_para_aplicar.search(linea)
            if m:
                info_pdf['Total para Aplicar'] = convertir_a_float(m.group(1))

        # Nuevo Capital Pendiente
        if 'NuevoCapitalPendiente' in linea:
            m = pat_nuevo_capital_pendiente.search(linea)
            if m:
                info_pdf['Nuevo Capital Pendiente'] = convertir_a_float(m.group(1))

    # Si no ha extraído nada útil, devolvemos None
    if not info_pdf['Operación'] and not info_pdf['Importe']:
        return None

    return info_pdf

# === Extrae datos de "Amortizaciones"
def extraer_amortizaciones(texto, lineas):
    # Código clave
    match_codigo = re.search(r'\bE\d{2}[A-Z]\d{8,}\b', texto)
    codigo_clave_extraido = match_codigo.group(0) if match_codigo else None

    # Fecha de recalculo
    match_fecha = re.search(r"FECHARECALCULO\.\:\s+(\d{1,2}\/\d{2}\/\d{4})", texto)
    fecha_recal = match_fecha.group(1) if match_fecha else None

    # Si no hay nada, no consideramos amortizaciones
    if not codigo_clave_extraido and not fecha_recal:
        return None

    # Extraer líneas relevantes
    lineas_deseadas = lineas[6:36]
    if not lineas_deseadas:
        return None

    filas_spliteadas = []
    for l in lineas_deseadas:
        l = re.sub(r"^\*", "", l).strip()
        columnas = l.split()
        filas_spliteadas.append(columnas)

    df_resultado = pd.DataFrame(filas_spliteadas)
    if df_resultado.empty or df_resultado.shape[0] < 2:
        return None

    # La primera fila son los nombres de columna
    df_resultado.columns = df_resultado.iloc[0]
    df_resultado = df_resultado.iloc[1:].reset_index(drop=True)

    # Comprobamos si existen las columnas
    columnas_necesarias = ['FECHA', 'CAPITAL', 'PENDIENTE']
    if not set(columnas_necesarias).issubset(df_resultado.columns):
        return None

    # Renombramos
    df_resultado = df_resultado[columnas_necesarias]
    df_resultado.rename(columns={
        'FECHA': 'Fecha',
        'CAPITAL': 'Amort anticipada',
        'PENDIENTE': 'Fee'
    }, inplace=True)

    # Transponer
    df_resultado.set_index('Fecha', inplace=True)
    df_resultado = df_resultado.T
    df_resultado.dropna(axis=1, how='all', inplace=True)

    return {
        'codigo': codigo_clave_extraido,
        'fecha_recal': fecha_recal,
        'df': df_resultado
    }

# === Adaptación a Streamlit ===
def main(files, pdfs=None, new_excel=None, month=None, year=None):
    try:
        if not pdfs:
            raise RuntimeError("No se han subido archivos PDF.")

        output = BytesIO()
        datos_financiaciones = []
        amortizaciones_por_pdf = []

       with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            if datos_financiaciones:
                df_resumen = pd.DataFrame(datos_financiaciones)
                df_resumen.to_excel(writer, sheet_name="Resumen", index=False)

            pd.DataFrame().to_excel(writer, sheet_name="Amortizaciones", index=False)
    
            # 🔹 Ahora accedemos correctamente a la hoja de amortizaciones
            ws_amort = writer.sheets["Amortizaciones"]

            row_offset = 0
            for pdf_name, info_amort in amortizaciones_por_pdf:
                df_amort = info_amort['df']
        
                # 🔹 Corregimos la forma de escribir celdas en `xlsxwriter`
                ws_amort.write(row_offset, 0, info_amort['codigo'])
                ws_amort.write(row_offset + 1, 0, info_amort['fecha_recal'])

                df_amort.to_excel(writer, sheet_name="Amortizaciones", startrow=row_offset+3, startcol=1, index=False)

                ws_amort.write(row_offset + 5, 0, "Amort anticipada")
                ws_amort.write(row_offset + 6, 0, "Fee")

                row_offset += df_amort.shape[0] + 7

            for pdf_name, info_amort in amortizaciones_por_pdf:
                df_amort = info_amort['df']
                ws_amort[f"A{row_offset+1}"] = info_amort['codigo']
                ws_amort[f"A{row_offset+2}"] = info_amort['fecha_recal']
                df_amort.to_excel(writer, sheet_name="Amortizaciones", startrow=row_offset+3, startcol=1, index=False)
                ws_amort[f"A{row_offset+5}"] = "Amort anticipada"
                ws_amort[f"A{row_offset+6}"] = "Fee"
                row_offset += df_amort.shape[0] + 7

        output.seek(0)
        return output

    except Exception as e:
        raise RuntimeError(f"Error al procesar financiaciones y amortizaciones: {str(e)}")
