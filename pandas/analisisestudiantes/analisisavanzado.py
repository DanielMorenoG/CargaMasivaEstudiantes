import pandas as pd
import plotly.express as px
import os
import dash
from  dash  import html,Input,Output,dcc
from dash import dash_table

#cargar los datos
ruta = os.path.join(os.path.dirname(__file__), "notas_limpio.xlsx")
dataf = pd.read_excel(ruta)


#iniciar app
appnotas = dash.Dash(__name__)


appnotas.layout= html.Div([
       html.H1("TABLERO AVANZADO",style={"textAlign":"center",
                                         "backgroundColor":"#1E1BD2",
                                         "color":"white","padding":"20px"}),

        #crear los filtros
        html.Div([
            html.Label("Seleccionar carrera"),
            dcc.Dropdown(id="filtro_carrera",
                         options=[{"label":ca,"value":ca}for ca in sorted(dataf["Carrera"].unique())],
                         value=dataf["Carrera"].unique()[0]
                         ),
                         html.Br(),

            html.Label("Rango de edad"),
            dcc.RangeSlider(id="slider_edad",
            min = dataf["Edad"].min(),
            max = dataf["Edad"].max(),
            step=1,
            value=[dataf["Edad"].min(), dataf["Edad"].max()],
            tooltip={"placement":"bottom","always_visible":True}           
            
            ),
            html.Br(),
            html.Label("Rango promedio"),
            dcc.RangeSlider(id="slider_promedio",
                            min=0,
                            max=5,
                            step=1.5,
                            value=[0,5],
                            tooltip={"placement":"bottom","always_visible":True}
                            )

        ],style={"width":"80%","margin":"auto"}),
        html.Br(),
        #crearlos Kpis
        html.Div(id="kpis",
                 style={"display":"flex",
                        "justifyContent":"space-around"}),
                       html.Br(),

        #crear una tabla
        dcc.Loading(
            dash_table.DataTable(id="tabla",
                                         page_size=8,
                                         filter_action="native",
                                         sort_action="native",
                                         row_selectable="multi",
                                         selected_rows=[],
                                         style_table={"overflowX":"auto"},
                                         style_cell={"textAlign":"center"}
                                         ),
                                         type="circle"
                                         ),
                                        html.Br(),
        #crearel graficos interactivos
        dcc.Loading(
            dcc.Graph(id="gra_detallado"),
            type="default"
        ),
        html.Br(),
        #crear los tabs
        dcc.Tabs([
            dcc.Tab(label="Histograma",children=[dcc.Graph(id="histograma")]),
            dcc.Tab(label="Dispersion",children=[dcc.Graph(id="dispersion")]),
            dcc.Tab(label="Desempeño",children=[dcc.Graph(id="pie")]),
            dcc.Tab(label="Promedio por Carrera", children=[dcc.Graph(id="barras")])
        ])
])

#actualizacion de la tabla y grafico general

@appnotas.callback(
     Output("tabla","data"),
     Output("tabla","columns"),
     Output("kpis","children"),
     Output("histograma","figure"),
     Output("dispersion","figure"),
     Output("pie","figure"),
     Output("barras","figure"),
     Input("filtro_carrera","value"),
     Input("slider_edad","value"),
     Input("slider_promedio","value")

)
def actualizar_comp(carrera,rangoedad,rangoprome):

    filtro =  dataf[
        (dataf["Carrera"]==carrera)&
        (dataf["Edad"]>=rangoedad[0])&
        (dataf["Edad"]<=rangoedad[1])&
        (dataf["Promedio"]>=rangoprome[0])&
        (dataf["Promedio"]<=rangoprome[1])
    ]


    #creacion de kpis
    promedio = round(filtro["Promedio"].mean(),2)
    total = len(filtro)
    maximo = round(filtro["Promedio"].mean(),2)

    kpis = [
        html.Div([html.H4("Promedio"),html.H2(promedio)],
                 style={"backgroundColor":"#3498db","color":"white","padding":"15px","borderRadius":"10px"}),

        html.Div([html.H4("Total estudiantes"),html.H2(total)],
                 style={"backgroundColor":"#3498db","color":"white","padding":"15px","borderRadius":"10px"}),

        html.Div([html.H4("Máximo"),html.H2(maximo)],
                 style={"backgroundColor":"#3498db","color":"white","padding":"15px","borderRadius":"10px"}),
    ]
    # Gráficos originales mejorados
    histo = px.histogram(filtro, x="Promedio", nbins=10,
                         title="Distribución de Promedios")

    dispersion = px.scatter(filtro, x="Edad", y="Promedio",
                            color="Desempeño",
                            trendline="ols",
                            title="Edad vs Promedio")

    pie = px.pie(filtro, names="Desempeño",
                 title="Distribución por Desempeño")

    promedios = dataf.groupby("Carrera")["Promedio"].mean().reset_index()

    barras = px.bar(promedios, x="Carrera", y="Promedio",
                    color="Carrera",
                    title="Promedio General por Carrera")

    return (
        filtro.to_dict("records"),
        [{"name":i,"id":i} for i in filtro.columns],
        kpis,
        histo,
        dispersion,
        pie,
        barras
    )

@appnotas.callback(
          Output("gra_detallado","figure"),
          Input("tabla","derived_virtual_data"),
          Input("tabla","derived_virtual_selected_rows"))


def actualizartab(rows,selected_rows):
    if rows is None:
        return px.scatter(title="sin datos")

    dff = pd.DataFrame(rows)   

    if selected_rows:
            dff = dff.iloc[selected_rows]   

            fig = px.scatter(dff,
                             x="Edad",
                             y="Promedio",
                             color="Desempeño",
                             size="Promedio",
                             title="Analisis detallado (Seleccione filas de la tabla)",
                             trendline="ols")
            return fig



#ejecutar la app
if __name__ == '__main__':
    appnotas.run(debug=True)