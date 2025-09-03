import pandas as pd
from sklearn.linear_model import LinearRegression
from sklearn.model_selection import train_test_split
from sklearn.metrics import mean_squared_error
from datetime import datetime

fecha = datetime.now()
day = fecha.strftime("%d")
day = int(day)

# Ruta archivo final excel
ruta_archivo_revisión = "C:/Users/Ricardo Sarda/Desktop/MM/Revisión pricing web.xlsx"
leadtime = "C:/Users/Ricardo Sarda/Downloads/Stock Enero 25 (1).xlsx" #leadtime
retoolcsv= "C:/Users/Ricardo Sarda/Downloads/table-data (2).csv" #retool

# archivo
retool = pd.read_csv(retoolcsv)
leadtime = pd.read_excel(leadtime, sheet_name='Stock', skiprows=1)

Columnas_retool = ['matrícula', 'frame_number', 'brand', 'model', 'Km', 'Año', 'Precio base', 'Oferta','model_id']
retool = retool[Columnas_retool]

retool = retool.merge(leadtime[['Item','LEADTIMEPOSTRENTING','P.Compra']], left_on='matrícula', right_on='Item', how='left')
retool['LEADTIMEPOSTRENTING'] = retool['LEADTIMEPOSTRENTING'].astype(float)
retool['LEADTIMEPOSTRENTING'] = retool['LEADTIMEPOSTRENTING'] + day

#Precio web
def Precio_web(row):
    if row['Oferta'] > 0:
        return row['Oferta']
    else:
        return row['Precio base']

retool['Precio web'] = retool.apply(Precio_web, axis=1)

def merge_retool(retool, articulossage):
    # Realizar el primer merge utilizando 'matrícula'
    merged_df = retool.merge(articulossage[['Artículo', 'P.Compra']], 
                             left_on='matrícula', 
                             right_on='Artículo', 
                             how='left', 
                             suffixes=('', '_matricula'))

    # Identificar filas con NaN en 'P.Compra'
    nan_rows = merged_df[merged_df['P.Compra'].isnull()]

    if not nan_rows.empty:
        # Realizar el segundo merge utilizando 'frame_number' para las filas con NaN
        nan_rows = nan_rows.drop(columns=['P.Compra', 'Artículo'])  # Eliminar columnas de merge anteriores
        nan_merged = nan_rows.merge(articulossage[['Artículo', 'P.Compra']], 
                                    left_on='frame_number', 
                                    right_on='Artículo', 
                                    how='left', 
                                    suffixes=('', '_frame'))

        # Combinar los resultados del segundo merge con el primero
        merged_df.update(nan_merged[['P.Compra', 'Artículo']])

    return merged_df


retool['Margen'] = retool['Precio web'] - retool['P.Compra']
retool['% Margen']= retool['Margen']/retool['Precio web']*100

retool['model'] = retool['model'].str.replace(' A2', '', regex=False)
retool['model'] = retool['model'].str.replace(' ABS', '', regex=False)

retool = retool.sort_values(by=['model'],ignore_index=True)
retool = retool.sort_values(by=['brand'],ignore_index=True)
retool['Año'] = retool['Año'].astype(int)
retool['Km'] = retool['Km'].astype(int)
retool['Coeficiente'] = (2024 - retool['Año'] + retool['Km']/5000)*100

#calcular variación precio
#Calcular el min precio por mpodelo
min_precios_por_modelo = retool.groupby('model')['Precio web'].min().reset_index()
min_precios_por_modelo.columns = ['model', 'Precio mínimo']

# Combinar el DataFrame original con los mínimos precios por modelo
retool = pd.merge(retool, min_precios_por_modelo, on='model')

# Calcular la diferencia y almacenarla en una nueva columna
retool['Variación precio'] = retool['Precio web'] - retool['Precio mínimo']

#calcular Varriacion coeficiente
# Calcular el máximo coeficiente por modelo
max_coeficientes_por_modelo = retool.groupby('model')['Coeficiente'].max().reset_index()
max_coeficientes_por_modelo.columns = ['model', 'Coeficiente máximo']

# Combinar el DataFrame original con los máximos coeficientes por modelo
retool = pd.merge(retool, max_coeficientes_por_modelo, on='model')

# Calcular la diferencia y almacenarla en una nueva columna
retool['Variación coeficiente'] = retool['Coeficiente máximo'] - retool['Coeficiente']

#resultado análisis
# Crea una nueva columna en el DataFrame para almacenar los resultados
retool['Resultado Variación'] = 0

# Aplica la lógica para cada fila, empezando desde la segunda fila
for index in range(1, len(retool)):
    if retool.at[index, 'model'] == retool.at[index - 1, 'model']:
        variacion_anterior = retool.at[index - 1, 'Variación coeficiente']
        variacion_actual = retool.at[index, 'Variación coeficiente']
        retool.at[index, 'Resultado Variación'] = (variacion_anterior - variacion_actual) / 10
    else:
        retool.at[index, 'Resultado Variación'] = 0

def actualizar_resultado_variacion(df):
    # Agrupar por modelo y año
    grupos = df.groupby(['model', 'Año'])
    
    # Iterar sobre cada grupo
    for _, grupo in grupos:
        if len(grupo) > 1:  # Solo si hay más de una moto del mismo modelo y año
            # Encontrar la moto con más km
            moto_con_mas_km = grupo.loc[grupo['Km'].idxmax()]
            # Encontrar la moto con menos km
            moto_con_menos_km = grupo.loc[grupo['Km'].idxmin()]

            # Verificar si la moto con más km tiene un precio más bajo
            if moto_con_mas_km['Precio web'] >= (moto_con_menos_km['Precio web']+2500):
                # Sumar 200 a 'Resultado variacion' para la moto con más km
                df.loc[moto_con_mas_km.name, 'Resultado Variación'] += ((moto_con_mas_km['Precio web'] - moto_con_menos_km['Precio web']+2500) * 0.05+200)

    return df

retool = actualizar_resultado_variacion(retool)

def actualizar_resultado_variacion2(df):
    # Agrupar por modelo
    grupos = df.groupby(['model'])
    
    # Iterar sobre cada grupo
    for _, grupo in grupos:
        if len(grupo) > 1:  # Solo si hay más de una moto del mismo modelo
            # Ordenar el grupo por kilometraje
            grupo = grupo.sort_values(by='Km')
            for i in range(len(grupo) - 1):
                moto1 = grupo.iloc[i]
                moto2 = grupo.iloc[i + 1]
                
                # Verificar si la diferencia en km es menor a 2500
                if abs(moto1['Km'] - moto2['Km']) < 2500:
                    # Verificar si la moto con el año más lejano (menor) tiene un precio menor
                    if moto1['Año'] < moto2['Año'] and moto1['Precio web'] >= moto2['Precio web']:
                        # Sumar 200 a 'Resultado variacion' para la moto con el año más lejano
                        df.loc[moto1.name, 'Resultado Variación'] += 200
                    elif moto2['Año'] < moto1['Año'] and moto2['Precio web'] >= moto1['Precio web']:
                        # Sumar 200 a 'Resultado variacion' para la moto con el año más lejano
                        df.loc[moto2.name, 'Resultado Variación'] += 200
                        
    return df

retool = actualizar_resultado_variacion2(retool)

def comparar_precios(df):
    # Agrupar por modelo
    grupos = df.groupby('model')
    
    # Iterar sobre cada grupo
    for _, grupo in grupos:
        if len(grupo) > 1:  # Solo si hay más de una moto del mismo modelo
            # Encontrar el precio más bajo
            precio_minimo = grupo['Precio web'].min()
            precio_maximo = grupo['Precio web'].max()
            # Iterar sobre cada fila del grupo
            for index, row in grupo.iterrows():
                if row['Precio web'] > precio_minimo:
                    # Calcular la diferencia de precio
                    diferencia_precio = row['Precio web'] - precio_minimo
                    # Actualizar el resultado de variación
                    df.loc[index, 'Resultado Variación'] += diferencia_precio * 0.1
    
    return df

retool = retool.sort_values(by='Resultado Variación',ascending=False, ignore_index=True)


Columnas_final = ['matrícula', 'frame_number', 'brand', 'model', 'Km', 'Año', 'P.Compra' , 'Precio base','Oferta', 'Precio web','Margen','% Margen','Resultado Variación', 'LEADTIMEPOSTRENTING']
retool = retool[Columnas_final]
retool.to_excel(ruta_archivo_revisión , index = False)