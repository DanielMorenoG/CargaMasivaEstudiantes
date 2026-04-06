

import requests
import pandas as pd
from dash import Dash, dcc, html,Output,Input
import plotly.express as px

def obtener_datos_api():
    url = "https://restcountries.com/v3.1/all?fields=name,region,population,area"

    try:
        response = requests.get(url)
        response.raise_for_status()  # Si hay error, lanza excepción
        
        data = response.json()
        
        paises = []
        for pais in data:
            paises.append({
                "pais": pais.get("name", {}).get("common"),
                "continente": pais.get("region"),
                "poblacion": pais.get("population"),
                "area": pais.get("area")
            })

        df = pd.DataFrame(paises)
        return df

    except Exception as e:
        print("Error consultando API:", e)
        return None


# =========================
# Cargar datos
# =========================

df = obtener_datos_api()

if df is None or df.empty:
    print("No se pudieron cargar los datos. Verifica la API.")
    exit()


# =========================
# Crear App
# =========================

app = Dash(__name__)

app.layout = html.Div([
    html.H1("Dashboard Mundial 🌍"),

    dcc.Dropdown(
        id="continente-dropdown",
        options=[{"label": c, "value": c} for c in df["continente"].dropna().unique()],
        value=df["continente"].dropna().unique()[0]
    ),

    dcc.Graph(id="grafico")
])


@app.callback(
    Output("grafico", "figure"),
    Input("continente-dropdown", "value")
)
def actualizar_grafico(continente):
    df_filtrado = df[df["continente"] == continente]

    fig = px.bar(
        df_filtrado.sort_values("poblacion", ascending=False).head(10),
        x="pais",
        y="poblacion",
        title=f"Top 10 países en {continente}"
    )

    return fig


if __name__ == "__main__":
    app.run(debug=True)