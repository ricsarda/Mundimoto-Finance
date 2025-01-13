import dash
from dash import dcc, html
from dash.dependencies import Input, Output
import dash_bootstrap_components as dbc
import pandas as pd
from sklearn.preprocessing import OneHotEncoder, StandardScaler
from sklearn.compose import ColumnTransformer
from sklearn.linear_model import LinearRegression
from sklearn.pipeline import Pipeline

# Datos de mundi
data = pd.read_csv("C:/Users/Ricardo Sarda/Desktop/MM/Motos/Motos para calcular.csv", delimiter=';', encoding='utf-8')

preprocessor = ColumnTransformer(
    transformers=[
        ('num', StandardScaler(), ['Año', 'KM']),
        ('cat', OneHotEncoder(drop='first', handle_unknown='ignore'), ['MARCA', 'MODELO'])
    ])

model = Pipeline([
    ('preprocessor', preprocessor),
    ('regressor', LinearRegression())
])

X = data[['MARCA', 'MODELO', 'Año', 'KM']]
y = data['PVP']
model.fit(X, y)

# La app
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])

app.layout = dbc.Container([
    html.H1("Calculadora PVP", style = {'font-size': '75px','font-weight': 'bold'}),
    dbc.Row([
        dbc.Col([
            html.Label("Marca" , style={'font-size': '18px','font-weight': 'bold'} ),
            dcc.Dropdown(id='marca-dropdown', 
                         options=[{'label': marca, 'value': marca} for marca in data['MARCA'].unique()],
                         value=data['MARCA'].iloc[0], className='mb-3')
        ]),
        
        dbc.Col([
            html.Label("Modelo" , style={'font-size': '18px','font-weight': 'bold'} ),
            dcc.Dropdown(id='modelo-dropdown', 
                         value=data['MODELO'].iloc[0], className='mb-3')
        ])
    ]),
    
    dbc.Row([
        dbc.Col([
            html.Label("Año" , style={'font-size': '18px','font-weight': 'bold'} ),
            dcc.Dropdown(
                id='ano-dropdown',
                options=[{'label': str(year), 'value': year} for year in range(int(data['Año'].min()), int(data['Año'].max()) + 1)],
                value=int(data['Año'].mean())
            )
        ],width=6),
       
        dbc.Col([
            html.Label("Kilometraje" , style={'font-size': '18px','font-weight': 'bold'} ),
            dbc.Input(id="km-input", type="number", placeholder="Introduce kilómetros", value=int(data['KM'].median()))
        ],width=6,),
    ]),

     dbc.Row([ 
        dbc.Col([
            dbc.Button("Calcular precio", id='calculate-button', color="primary", className="mr-1", style={'width': '400px' , 'font-weight': 'bold'}),
        ],width=4,  md=12, className="d-grid gap-2"),
    ]),
        html.Br(),
        html.Br(),
    
        html.Div(id='output')
    
])

@app.callback(
    Output('modelo-dropdown', 'options'),
    [Input('marca-dropdown', 'value')]
)
def update_modelos(marca_selected):
    modelos = data[data['MARCA'] == marca_selected]['MODELO'].unique()
    return [{'label': modelo, 'value': modelo} for modelo in modelos]


@app.callback(
    Output('output', 'children'),
    [Input('calculate-button', 'n_clicks')],
    [dash.dependencies.State('marca-dropdown', 'value'),
     dash.dependencies.State('modelo-dropdown', 'value'),
     dash.dependencies.State('ano-dropdown', 'value'),
     dash.dependencies.State('km-input', 'value')]
)
def calculate_price(n_clicks, marca, modelo, año, km):
    # Predecir el precio
    prediction = model.predict(pd.DataFrame({
        'MARCA': [marca],
        'MODELO': [modelo],
        'Año': [año],
        'KM': [km]
    }))
    
    # Filtrar los datos para el modelo seleccionado
    subset_data = data[data['MODELO'] == modelo]
    
    # Calcular el número de motos y la desviación estándar
    num_motos = len(subset_data)
    std_dev = subset_data['PVP'].std()
    posibleprecio = std_dev / 2
    min_ano = int(subset_data['Año'].min())
    max_km = int(subset_data['KM'].max())

    return [
        html.Div(f"Precio estimado: {prediction[0]:,.2f}€", style={'font-size': '22px', 'color': 'black'}),
        html.Div(f"Variación a tener en cuenta: +/-{posibleprecio:,.2f}€", style={'font-size': '15px', 'color': 'red'}),
        html.Div(f"Mayor antigüedad: {min_ano}", style={'font-size': '15px', 'color': 'black'}),
        html.Div(f"Mayor kilometraje: {max_km}KM", style={'font-size': '15px', 'color': 'black'}),
        html.Div(f"Número de motos en el análisis: {num_motos}", style={'font-size': '15px', 'color': 'black'})
    ]

if __name__ == '__main__':
    app.run_server(debug=True)