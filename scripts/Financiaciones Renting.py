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
    pat_fecha                   = re.compile(r"Fecha:([0-9]{2}/[0-9]{2}/[0-9]{2})")
    pat_operacion               = re.compile(r"operación\s*nº\s*([A-Za-z0-9]+)")  
    pat_entrega_importe         = re.compile(r"EntregaImporte\s+([\d.,]+)")
    pat_intereses               = re.compile(r"InteresesDevengados.*?([\d.,]+)$")
    pat_total_para_aplicar      = re.compile(r"TotalparaAplicaraCapital\s+([\d.,]+)")
    pat_nuevo_capital_pendiente = re.compile(r"NuevoCapitalPendiente\s+([\d.,]+)")

    if ('operación nº' not in texto) and ('EntregaImporte' not in texto):
        return None

    info_pdf = {
        'Fecha': None,
        'Importe': None,
        'Intereses': None,
        'Total para Aplicar': None,
        'Nuevo Capital Pendiente': None,
        'Operación': None,
        'Archivo': nombre_pdf
    }

    for linea in lineas:
        if 'Fecha:' in linea:
            m = pat_fecha.search(linea)
            if m:
                info_pdf['Fecha'] = m.group(1)

        if 'operación nº' in linea:
            m = pat_operacion.search(linea)
            if m:
                info_pdf['Operación'] = m.group(1)

        if 'EntregaImporte' in linea:
            m = pat_entrega_importe.search(linea)
            if m:
                info_pdf['Importe'] = convertir_a_float(m.group(1))

        if 'InteresesDevengados' in linea:
            m = pat_intereses.search(linea)
            if m:
                info_pdf['Intereses'] = convertir_a_float(m.group(1))

        if 'TotalparaAplicaraCapital' in linea:
            m = pat_total_para_aplicar.search(linea)
            if m:
                info_pdf['Total para Aplicar'] = convertir_a_float(m.group(1))

        if 'NuevoCapitalPendiente' in linea:
            m = pat_nuevo_capital_pendiente.search(linea)
            if m:
                info_pdf['Nuevo Capital Pendiente'] = convertir_a_float(m.group(1))

    if not info_pdf['Operación'] and not info_pdf['Importe']:
        return None

    return info_pdf

# === Extrae datos de "Amortizaciones"
def extraer_amortizaciones(texto, lineas):
    match_codigo = re.search(r'\bE\d{2}[A-Z]\d{8,}\b', texto)
    codigo_clave_extraido = match_codigo.group(0) if match_codigo else None

    match_fecha = re.search(r"FECHARECALCULO\.\:\s+(\d{1,2}\/\d{2}\/\d{4})", texto)
    fecha_recal = match_fecha.group(1) if match_fecha else None

    if not codigo_clave_extraido and not fecha_recal:
        return None

    lineas_deseadas = lineas[6:36]
    if not lineas_deseadas:
        return None

    filas_spliteadas = [re.sub(r"^\*", "", l).strip().split() for l in lineas_deseadas]
    df_resultado = pd.DataFrame(filas_spliteadas)

    if df_resultado.empty or df_resultado.shape[0] < 2:
        return None

    df_resultado.columns = df_resultado.iloc[0]
    df_resultado = df_resultado.iloc[1:].reset_index(drop=True)

    columnas_necesarias = ['FECHA', 'CAPITAL', 'PENDIENTE']
    if not set(columnas_necesarias).issubset(df_resultado.columns):
        return None

    df_resultado = df_resultado[columnas_necesarias]
    df_resultado.rename(columns={
        'FECHA': 'Fecha',
        'CAPITAL': 'Amort anticipada',
        'PENDIENTE': 'Fee'
    }, inplace=True)

    df_resultado.set_index('Fecha', inplace=True)
    df_resultado = df_resultado.T
    df_resultado.dropna(axis=1, how='all', inplace=True)

    return {'codigo': codigo_clave_extraido, 'fecha_recal': fecha_recal, 'df': df_resultado}

# === Función Principal ===
def main(files, pdfs=None, new_excel=None, month=None, year=None):
    try:
        if not pdfs:
            raise RuntimeError("No se han subido archivos PDF.")

        fecha_actual = datetime.now().strftime("%d-%m-%Y")
        output = BytesIO()

        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            datos_financiaciones = []
            amortizaciones_por_pdf = []

            for pdf_name, pdf_file in pdfs.items():
                with pdfplumber.open(pdf_file) as pdf:
                    texto = "\n".join(page.extract_text() for page in pdf.pages if page.extract_text())
                lineas = texto.split('\n')

                info_fin = extraer_financiaciones(texto, lineas, pdf_name)
                if info_fin is not None:
                    datos_financiaciones.append(info_fin)

                info_amort = extraer_amortizaciones(texto, lineas)
                if info_amort is not None:
                    amortizaciones_por_pdf.append((pdf_name, info_amort))

            if datos_financiaciones:
                df_resumen = pd.DataFrame(datos_financiaciones)
                columnas_orden = [
                    'Fecha', 'Importe', 'Intereses', 'Total para Aplicar',
                    'Nuevo Capital Pendiente', 'Operación', 'Archivo'
                ]
                df_resumen = df_resumen[[c for c in columnas_orden if c in df_resumen.columns]]
                df_resumen.to_excel(writer, sheet_name="Resumen", index=False)

            writer.book.create_sheet("Amortizaciones")
            ws_amort = writer.book["Amortizaciones"]
            row_offset = 0

            for (pdf_name, info_amort) in amortizaciones_por_pdf:
                df_amort = info_amort['df']
                ws_amort[f"A{row_offset+1}"] = info_amort['codigo'] or "COD NO ENCONTRADO"
                ws_amort[f"A{row_offset+2}"] = info_amort['fecha_recal'] or "FECHA NO ENCONTRADA"
                df_amort.to_excel(writer, sheet_name="Amortizaciones", startrow=row_offset+3, startcol=1, index=False)
                ws_amort[f"A{row_offset+5}"] = "Amort anticipada"
                ws_amort[f"A{row_offset+6}"] = "Fee"
                row_offset += df_amort.shape[0] + 6

        output.seek(0)
        return output

    except Exception as e:
        raise RuntimeError(f"Error al procesar las financiaciones y amortizaciones: {str(e)}")
