import pandas as pd
from datetime import datetime
from io import BytesIO

def main(files, pdfs, new_excel, month=None, year=None):
    try:

        #   - "RetoolCSV": CSV exportado de Retool (table-data*.csv)
        #   - "LeadtimeExcel": Excel con sheet 'Stock' (skiprows=1), columnas: Item, LEAD TIME POSTRENTING, Inv. Value
        retool = pd.read_csv(files["RetoolCSV"])
        leadtime = pd.read_excel(files["LeadtimeExcel"], sheet_name='Stock', skiprows=1)

        day = int(datetime.now().strftime("%d"))

        # === Selección y merge inicial ===
        columnas_retool = ['matrícula', 'frame_number', 'brand', 'model', 'Km', 'Año', 'Precio base', 'Oferta', 'model_id']
        retool = retool[columnas_retool].copy()

        retool = retool.merge(
            leadtime[['Item', 'LEAD TIME POSTRENTING', 'Inv. Value']],
            left_on='matrícula', right_on='Item', how='left'
        )

        # Tipos y ajustes
        retool['LEAD TIME POSTRENTING'] = pd.to_numeric(retool['LEAD TIME POSTRENTING'], errors='coerce').fillna(0)
        retool['LEAD TIME POSTRENTING'] = retool['LEAD TIME POSTRENTING'] + day

        # Precio web (si hay Oferta>0 usamos Oferta; si no, Precio base)
        retool['Oferta'] = pd.to_numeric(retool['Oferta'], errors='coerce').fillna(0)
        retool['Precio base'] = pd.to_numeric(retool['Precio base'], errors='coerce')
        retool['Precio web'] = retool.apply(
            lambda r: r['Oferta'] if r['Oferta'] > 0 else r['Precio base'],
            axis=1
        )

        # Márgenes
        retool['Inv. Value'] = pd.to_numeric(retool['Inv. Value'], errors='coerce')
        retool['Margen'] = retool['Precio web'] - retool['Inv. Value']
        retool['% Margen'] = (retool['Margen'] / retool['Precio web'] * 100).replace([pd.NA, pd.NaT], 0)

        # Limpieza de modelo
        retool['model'] = retool['model'].astype(str).str.replace(' A2', '', regex=False).str.replace(' ABS', '', regex=False)

        # Orden y tipos
        retool = retool.sort_values(by=['brand', 'model'], ignore_index=True)
        retool['Año'] = pd.to_numeric(retool['Año'], errors='coerce')
        retool['Km'] = pd.to_numeric(retool['Km'], errors='coerce').fillna(0).astype(int)
        retool['Año'] = retool['Año'].fillna(retool['Año'].median()).astype(int)

        # Coeficiente (mantengo la referencia a 2024 tal como tu script original)
        retool['Coeficiente'] = (2024 - retool['Año'] + retool['Km'] / 5000) * 100

        # --- Variaciones por modelo ---
        # Por precio
        min_precios = retool.groupby('model', as_index=False)['Precio web'].min().rename(columns={'Precio web': 'Precio mínimo'})
        retool = retool.merge(min_precios, on='model', how='left')
        retool['Variación precio'] = retool['Precio web'] - retool['Precio mínimo']

        # Por coeficiente
        max_coef = retool.groupby('model', as_index=False)['Coeficiente'].max().rename(columns={'Coeficiente': 'Coeficiente máximo'})
        retool = retool.merge(max_coef, on='model', how='left')
        retool['Variación coeficiente'] = retool['Coeficiente máximo'] - retool['Coeficiente']

        # Resultado Variación base
        retool['Resultado Variación'] = 0.0
        for idx in range(1, len(retool)):
            if retool.at[idx, 'model'] == retool.at[idx - 1, 'model']:
                va = retool.at[idx - 1, 'Variación coeficiente']
                vb = retool.at[idx, 'Variación coeficiente']
                if pd.notna(va) and pd.notna(vb):
                    retool.at[idx, 'Resultado Variación'] = (va - vb) / 10
            else:
                retool.at[idx, 'Resultado Variación'] = 0.0

        # Ajuste por modelo y año (más km y precio ≥ menor km + 2500)
        def actualizar_resultado_variacion(df):
            for (_, _ano), grupo in df.groupby(['model', 'Año']):
                if len(grupo) > 1:
                    g_idxmax = grupo['Km'].idxmax()
                    g_idxmin = grupo['Km'].idxmin()
                    moto_max = df.loc[g_idxmax]
                    moto_min = df.loc[g_idxmin]
                    if pd.notna(moto_max['Precio web']) and pd.notna(moto_min['Precio web']):
                        if moto_max['Precio web'] >= (moto_min['Precio web'] + 2500):
                            delta = (moto_max['Precio web'] - moto_min['Precio web'] + 2500) * 0.05 + 200
                            df.loc[g_idxmax, 'Resultado Variación'] += delta
            return df

        retool = actualizar_resultado_variacion(retool)

        # Ajuste por pares cercanos en Km dentro del mismo modelo
        def actualizar_resultado_variacion2(df):
            for _model, grupo in df.groupby('model'):
                if len(grupo) > 1:
                    grupo = grupo.sort_values(by='Km')
                    for i in range(len(grupo) - 1):
                        i1 = grupo.index[i]
                        i2 = grupo.index[i + 1]
                        m1 = df.loc[i1]
                        m2 = df.loc[i2]
                        if abs(m1['Km'] - m2['Km']) < 2500:
                            if (m1['Año'] < m2['Año']) and (m1['Precio web'] >= m2['Precio web']):
                                df.loc[i1, 'Resultado Variación'] += 200
                            elif (m2['Año'] < m1['Año']) and (m2['Precio web'] >= m1['Precio web']):
                                df.loc[i2, 'Resultado Variación'] += 200
            return df

        retool = actualizar_resultado_variacion2(retool)

        # Ordenar por resultado
        retool = retool.sort_values(by='Resultado Variación', ascending=False, ignore_index=True)
        retool['% Margen'] = retool['% Margen'].round(2)
        retool['Resultado Variación'] = retool['Resultado Variación'].round(2)
        # Salida final
        columnas_final = [
            'matrícula', 'frame_number', 'brand', 'model', 'Km', 'Año',
            'Inv. Value', 'Precio base', 'Oferta', 'Precio web', 'Margen',
            '% Margen', 'Resultado Variación', 'LEAD TIME POSTRENTING'
        ]
        retool = retool[columnas_final]

        # Exportar a Excel en memoria
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            retool.to_excel(writer, sheet_name='Revision pricing web', index=False)
        output.seek(0)
        return output

    except KeyError as e:
        raise RuntimeError(f"Falta el archivo clave 'files': {str(e)}")
    except Exception as e:
        raise RuntimeError(f"Error al procesar la revisión de pricing: {str(e)}")


