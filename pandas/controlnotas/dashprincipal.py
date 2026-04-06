import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import unicodedata
import re as _re
import dash
from dash import html, Input, Output, State, dcc, callback_context
from dash import dash_table
from database import (obtenerestudiantes, insertar_estudiante, insertar_masivo,
                      obtener_claves_existentes, existe_estudiante)

# ── Paleta ─────────────────────────────────────────────────────────────────────
BG       = "#0c0e10"
SURFACE  = "#13161a"
SURFACE2 = "#181c21"
BORDER   = "#242830"
ACCENT   = "#c8a96e"
BLUE     = "#60a5fa"
GREEN    = "#4ade80"
RED      = "#e05c5c"
ORANGE   = "#fb923c"
TEXT     = "#e8e2d9"
MUTED    = "#6b7280"

C_NOMBRE   = "nombre"
C_CARRERA  = "carrera"
C_EDAD     = "edad"
C_PROMEDIO = "promedio"
C_DESEMPEN = "desempenio"
C_NOTA1    = "nota1"
C_NOTA2    = "nota2"
C_NOTA3    = "nota3"

MATERIAS = [
    {"label": "Todas las materias", "value": "todas"},
    {"label": "Materia 1 (Nota 1)",  "value": C_NOTA1},
    {"label": "Materia 2 (Nota 2)",  "value": C_NOTA2},
    {"label": "Materia 3 (Nota 3)",  "value": C_NOTA3},
]

ORDEN_PROMEDIO = [
    {"label": "Sin ordenar",         "value": "ninguno"},
    {"label": "Mejores promedios ↑", "value": "desc"},
    {"label": "Peores promedios ↓",  "value": "asc"},
]

PLOTLY_LAYOUT_BASE = dict(
    paper_bgcolor=SURFACE,
    plot_bgcolor=SURFACE2,
    font=dict(family="'DM Mono', monospace", color=TEXT, size=12),
    colorway=[ACCENT, BLUE, GREEN, RED, "#a78bfa", "#fb923c"],
    xaxis=dict(gridcolor=BORDER, linecolor=BORDER, tickcolor=BORDER, tickfont=dict(size=11)),
    yaxis=dict(gridcolor=BORDER, linecolor=BORDER, tickcolor=BORDER, tickfont=dict(size=11)),
    legend=dict(bgcolor=SURFACE, bordercolor=BORDER, borderwidth=1, font=dict(size=11)),
    margin=dict(l=48, r=24, t=56, b=48),
)

TITLE_FONT = dict(family="'DM Serif Display', serif", color=TEXT, size=17)

CARD = {
    "backgroundColor": SURFACE,
    "border": f"1px solid {BORDER}",
    "padding": "24px",
    "position": "relative",
}

# ── Utilidades de limpieza ─────────────────────────────────────────────────────

def normalizar_nombre(nombre):
    """
    'MARIA JOSE' → 'Maria Jose'
    '   juan  '  → 'Juan'
    Elimina espacios múltiples y aplica title case.
    """
    if not isinstance(nombre, str):
        return nombre
    nombre = nombre.strip()
    nombre = _re.sub(r"\s+", " ", nombre)          # espacios múltiples → uno
    return nombre.title()                            # Title Case

def limpiar_col(s):
    """Normaliza nombre de columna: minúsculas, sin tildes, sin símbolos."""
    s = str(s).strip().lower()
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = _re.sub(r"[\s\-_\.\(\)\[\]\/\\]+", "", s)
    return _re.sub(r"[^a-z0-9]", "", s)

def calc_desempenio(p):
    if p >= 4.5: return "Excelente"
    if p >= 3.5: return "Bueno"
    if p >= 3.0: return "Regular"
    return "Deficiente"

# ─────────────────────────────────────────────────────────────────────────────

def apply_template(fig, title=""):
    fig.update_layout(
        **PLOTLY_LAYOUT_BASE,
        title=dict(text=title, font=TITLE_FONT, pad=dict(b=16)),
    )
    return fig


def kpi_card(label, value, color=ACCENT):
    return html.Div([
        html.Div(style={
            "position": "absolute", "top": "-1px", "right": "-1px",
            "width": "20px", "height": "20px",
            "borderTop": f"2px solid {color}",
            "borderRight": f"2px solid {color}",
        }),
        html.P(label, style={
            "fontSize": "10px", "letterSpacing": ".2em",
            "textTransform": "uppercase", "color": MUTED, "marginBottom": "10px",
        }),
        html.H2(str(value), style={
            "fontFamily": "'DM Serif Display', serif",
            "fontSize": "40px", "fontWeight": "400",
            "color": color, "margin": "0", "lineHeight": "1",
        }),
    ], style={**CARD, "flex": "1", "minWidth": "150px"})


def build_index_string():
    css = (
        "*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }"
        "body { background:" + BG + "; color:" + TEXT + "; font-family:'DM Mono',monospace; font-size:13px; line-height:1.6; min-height:100vh; }"
        "body::before { content:''; position:fixed; inset:0; background-image:linear-gradient(rgba(200,169,110,.03) 1px,transparent 1px),linear-gradient(90deg,rgba(200,169,110,.03) 1px,transparent 1px); background-size:48px 48px; pointer-events:none; z-index:0; }"
        ".Select-control,.Select-menu-outer { background-color:" + SURFACE + " !important; border-color:" + BORDER + " !important; color:" + TEXT + " !important; }"
        ".Select-value-label,.Select-placeholder { color:" + TEXT + " !important; }"
        ".VirtualizedSelectOption { background:" + SURFACE + "; color:" + TEXT + "; }"
        ".VirtualizedSelectFocusedOption { background:" + BORDER + " !important; }"
        ".dash-dropdown .Select-control { background-color:" + SURFACE + " !important; border:1px solid " + BORDER + " !important; border-radius:0 !important; }"
        ".dash-dropdown .Select-value,.dash-dropdown .Select-placeholder { color:" + TEXT + " !important; }"
        ".dash-dropdown .Select-arrow { border-top-color:" + MUTED + " !important; }"
        ".dash-dropdown .Select-menu-outer { background-color:" + SURFACE + " !important; border:1px solid " + BORDER + " !important; border-radius:0 !important; z-index:9999 !important; }"
        ".dash-dropdown .Select-option { background-color:" + SURFACE + " !important; color:" + TEXT + " !important; font-family:'DM Mono',monospace !important; font-size:12px !important; }"
        ".dash-dropdown .Select-option.is-focused { background-color:" + BORDER + " !important; color:" + ACCENT + " !important; }"
        ".dash-dropdown .Select-option.is-selected { background-color:" + BORDER + " !important; color:" + ACCENT + " !important; }"
        ".dash-dropdown .Select-input > input { color:" + TEXT + " !important; }"
        ".dropdown { background-color:" + SURFACE + " !important; }"
        ".dash-table-container .dash-spreadsheet-container .dash-spreadsheet-inner td,"
        ".dash-table-container .dash-spreadsheet-container .dash-spreadsheet-inner th { background-color:" + SURFACE + " !important; color:" + TEXT + " !important; border-color:" + BORDER + " !important; font-family:'DM Mono',monospace !important; font-size:12px !important; }"
        ".dash-table-container .dash-spreadsheet-container .dash-spreadsheet-inner th { color:" + MUTED + " !important; letter-spacing:.12em; text-transform:uppercase; font-size:10px !important; }"
        ".dash-table-container .dash-spreadsheet-container .dash-spreadsheet-inner tr:hover td { background-color:" + BORDER + " !important; }"
        "#buscador_nombre { width:100%; background:" + SURFACE2 + "; border:1px solid " + BORDER + "; color:" + TEXT + "; font-family:'DM Mono',monospace; font-size:13px; padding:10px 14px; outline:none; transition:border-color .2s,box-shadow .2s; }"
        "#buscador_nombre:focus { border-color:" + ACCENT + "; box-shadow:0 0 0 3px rgba(200,169,110,.08); }"
        "#buscador_nombre::placeholder { color:#333740; }"
        ".dash-tab { background:" + SURFACE + " !important; border-color:" + BORDER + " !important; color:" + MUTED + " !important; font-family:'DM Mono',monospace !important; font-size:11px !important; letter-spacing:.14em; text-transform:uppercase; padding:10px 18px !important; }"
        ".dash-tab--selected { color:" + ACCENT + " !important; border-bottom-color:" + ACCENT + " !important; }"
        ".rc-slider-track { background-color:" + ACCENT + " !important; }"
        ".rc-slider-handle { border-color:" + ACCENT + " !important; background:" + ACCENT + " !important; }"
        ".rc-slider-rail { background-color:" + BORDER + " !important; }"
        "::-webkit-scrollbar { width:6px; height:6px; }"
        "::-webkit-scrollbar-track { background:" + BG + "; }"
        "::-webkit-scrollbar-thumb { background:" + BORDER + "; border-radius:3px; }"
        ".btn-logout:hover { background-color:" + RED + " !important; color:#fff !important; }"
        ".input-registro { width:100%; background:" + SURFACE2 + "; border:1px solid " + BORDER + "; color:" + TEXT + "; font-family:'DM Mono',monospace; font-size:13px; padding:10px 14px; outline:none; transition:border-color .2s,box-shadow .2s; box-sizing:border-box; }"
        ".input-registro:focus { border-color:" + ACCENT + "; box-shadow:0 0 0 3px rgba(200,169,110,.08); }"
        ".input-registro::placeholder { color:#333740; }"
        ".btn-registrar { background:" + ACCENT + "; color:#0c0e10; border:none; font-family:'DM Mono',monospace; font-size:11px; font-weight:500; letter-spacing:.2em; text-transform:uppercase; padding:14px 32px; cursor:pointer; transition:background .2s; }"
        ".btn-registrar:hover { background:#d4b87a; }"
        ".upload-zona { border:2px dashed " + BORDER + "; background:" + SURFACE2 + "; padding:40px 24px; text-align:center; cursor:pointer; transition:border-color .2s; }"
        ".upload-zona:hover { border-color:" + ACCENT + "; }"
        "#badge_busqueda { font-size:10px; letter-spacing:.14em; color:" + MUTED + "; margin-top:6px; min-height:16px; }"

        # ── Modal de confirmación ─────────────────────────────────────────────
        ".modal-overlay { position:fixed; inset:0; background:rgba(0,0,0,.72); z-index:9000; display:flex; align-items:center; justify-content:center; }"
        ".modal-box { background:" + SURFACE + "; border:1px solid " + BORDER + "; border-left:4px solid " + ORANGE + "; padding:32px 36px; max-width:620px; width:92%; max-height:85vh; overflow-y:auto; position:relative; animation:slideUp .25s ease both; }"
        ".modal-box::before { content:''; position:absolute; top:-1px; right:-1px; width:24px; height:24px; border-top:2px solid " + ORANGE + "; border-right:2px solid " + ORANGE + "; }"
        "@keyframes slideUp { from{opacity:0;transform:translateY(16px)} to{opacity:1;transform:translateY(0)} }"
        ".modal-btn-ok { background:" + ORANGE + "; color:#0c0e10; border:none; font-family:'DM Mono',monospace; font-size:11px; font-weight:500; letter-spacing:.2em; text-transform:uppercase; padding:12px 28px; cursor:pointer; transition:background .2s; margin-right:12px; }"
        ".modal-btn-ok:hover { background:#f59e4a; }"
        ".modal-btn-cancel { background:transparent; color:" + MUTED + "; border:1px solid " + BORDER + "; font-family:'DM Mono',monospace; font-size:11px; letter-spacing:.2em; text-transform:uppercase; padding:12px 28px; cursor:pointer; transition:border-color .2s,color .2s; }"
        ".modal-btn-cancel:hover { border-color:" + MUTED + "; color:" + TEXT + "; }"
        ".modal-btn-delete { background:transparent; color:" + RED + "; border:1px solid " + RED + "; font-family:'DM Mono',monospace; font-size:11px; letter-spacing:.2em; text-transform:uppercase; padding:12px 28px; cursor:pointer; transition:background .2s,color .2s; }"
        ".modal-btn-delete:hover { background:" + RED + "; color:#fff; }"
        ".modal-btn-fill { background:transparent; color:" + BLUE + "; border:1px solid " + BLUE + "; font-family:'DM Mono',monospace; font-size:11px; letter-spacing:.2em; text-transform:uppercase; padding:12px 28px; cursor:pointer; transition:background .2s,color .2s; }"
        ".modal-btn-fill:hover { background:" + BLUE + "; color:#0c0e10; }"
        ".badge-limpieza { display:inline-block; font-size:10px; letter-spacing:.12em; text-transform:uppercase; padding:3px 9px; border:1px solid; border-radius:2px; margin-right:6px; margin-bottom:4px; }"
        ".badge-verde { color:" + GREEN + "; border-color:" + GREEN + "; background:rgba(74,222,128,.06); }"
        ".badge-naranja { color:" + ORANGE + "; border-color:" + ORANGE + "; background:rgba(251,146,60,.06); }"
    )

    return (
        "<!DOCTYPE html><html><head>{%metas%}"
        "<title>Dashboard — Sistema de Estudiantes</title>"
        "{%favicon%}{%css%}"
        "<style>" + css + "</style>"
        "</head><body>"
        "{%app_entry%}"
        "<footer>{%config%}{%scripts%}{%renderer%}</footer>"
        "</body></html>"
    )


def creartablero(server):

    def cargar_datos():
        df = obtenerestudiantes()
        df.columns = [c.lower() for c in df.columns]
        rename_map = {col: "desempenio" for col in df.columns if "desempe" in col}
        if rename_map:
            df = df.rename(columns=rename_map)
        return df

    dataf = cargar_datos()

    appnotas = dash.Dash(
        __name__,
        server=server,
        url_base_pathname="/dashprincipal/",
        suppress_callback_exceptions=True,
        external_stylesheets=[
            "https://fonts.googleapis.com/css2?family=DM+Serif+Display:ital@0;1"
            "&family=DM+Mono:wght@300;400;500&display=swap"
        ],
    )

    appnotas.index_string = build_index_string()

    # ══════════════════════════════════════════════════════════════════════════
    # LAYOUT
    # ══════════════════════════════════════════════════════════════════════════
    appnotas.layout = html.Div(style={"position": "relative", "zIndex": "1"}, children=[

        # ── Modal de confirmación de notas fuera de rango ─────────────────────
        html.Div(
            id="modal_notas",
            style={"display": "none"},
            children=[
                html.Div([
                    html.Div([
                        html.Span("⚠", style={"fontSize": "22px", "color": ORANGE,
                                               "marginRight": "12px", "verticalAlign": "middle"}),
                        html.Span("NOTAS FUERA DE RANGO", style={
                            "fontSize": "11px", "letterSpacing": ".2em",
                            "color": ORANGE, "fontWeight": "500", "verticalAlign": "middle",
                        }),
                    ], style={"marginBottom": "20px"}),

                    html.P(id="modal_notas_msg", style={
                        "color": TEXT, "fontSize": "12px", "lineHeight": "1.7",
                        "marginBottom": "16px", "whiteSpace": "pre-line",
                        "maxHeight": "220px", "overflowY": "auto",
                        "background": SURFACE2, "border": f"1px solid {BORDER}",
                        "padding": "12px 14px",
                        "fontFamily": "'DM Mono', monospace",
                    }),

                    # Descripción de las 3 opciones
                    html.Div([
                        html.Div([
                            html.Span("✓ AJUSTAR", style={"color": ORANGE, "fontSize": "10px",
                                                           "letterSpacing": ".15em", "fontWeight": "500"}),
                            html.P("Notas >5 quedan en 5.0, notas <0 quedan en 0.0.",
                                   style={"color": MUTED, "fontSize": "11px", "marginTop": "4px"}),
                        ], style={"flex": "1", "padding": "10px 12px",
                                  "border": f"1px solid {BORDER}", "background": SURFACE2}),
                        html.Div([
                            html.Span("⬜ RELLENAR", style={"color": BLUE, "fontSize": "10px",
                                                             "letterSpacing": ".15em", "fontWeight": "500"}),
                            html.P("Notas vacías se completan con 0. Notas fuera de rango se ajustan.",
                                   style={"color": MUTED, "fontSize": "11px", "marginTop": "4px"}),
                        ], style={"flex": "1", "padding": "10px 12px",
                                  "border": f"1px solid {BORDER}", "background": SURFACE2}),
                        html.Div([
                            html.Span("✕ ELIMINAR", style={"color": RED, "fontSize": "10px",
                                                            "letterSpacing": ".15em", "fontWeight": "500"}),
                            html.P("Los estudiantes con notas inválidas o vacías son eliminados del lote.",
                                   style={"color": MUTED, "fontSize": "11px", "marginTop": "4px"}),
                        ], style={"flex": "1", "padding": "10px 12px",
                                  "border": f"1px solid {BORDER}", "background": SURFACE2}),
                    ], style={"display": "flex", "gap": "10px", "marginBottom": "24px",
                              "flexWrap": "wrap"}),

                    html.Div([
                        html.Button("✓ Ajustar al límite (0–5)",
                                    id="modal_btn_aceptar", className="modal-btn-ok"),
                        html.Button("⬜ Rellenar vacíos con 0",
                                    id="modal_btn_fill", className="modal-btn-fill"),
                        html.Button("✕ Eliminar inválidos",
                                    id="modal_btn_eliminar", className="modal-btn-delete"),
                        html.Button("↩ Cancelar",
                                    id="modal_btn_cancelar", className="modal-btn-cancel"),
                    ], style={"display": "flex", "gap": "10px", "flexWrap": "wrap"}),
                ], className="modal-box"),
            ],
            className="modal-overlay",
        ),

        # Header
        html.Div([
            html.Div([
                html.Div([
                    html.Span("●", style={"color": ACCENT, "marginRight": "10px",
                                          "fontSize": "8px", "verticalAlign": "middle"}),
                    html.Span("PORTAL ACADÉMICO", style={"fontSize": "10px",
                                                         "letterSpacing": ".22em", "color": MUTED}),
                ], style={"marginBottom": "8px"}),
                html.H1([
                    "Tablero de ",
                    html.Em("Estudiantes", style={"color": ACCENT, "fontStyle": "italic"}),
                ], style={"fontFamily": "'DM Serif Display', serif",
                          "fontSize": "30px", "fontWeight": "400", "color": TEXT}),
            ]),
            html.A("⏻  CERRAR SESIÓN", href="/logout", className="btn-logout",
                   style={"fontSize": "10px", "letterSpacing": ".18em",
                          "textTransform": "uppercase", "color": RED,
                          "border": f"1px solid {RED}", "padding": "10px 20px",
                          "textDecoration": "none", "alignSelf": "center",
                          "fontFamily": "'DM Mono', monospace",
                          "transition": "background-color 0.2s, color 0.2s",
                          "cursor": "pointer"}),
        ], style={**CARD, "borderLeft": f"3px solid {ACCENT}",
                  "margin": "24px 24px 0 24px",
                  "display": "flex", "justifyContent": "space-between", "alignItems": "center"}),

        # Búsqueda y filtros avanzados
        html.Div([
            html.P("BÚSQUEDA Y FILTROS AVANZADOS", style={
                "fontSize": "10px", "letterSpacing": ".2em", "color": ACCENT, "marginBottom": "20px",
            }),
            html.Div([
                html.Div([
                    html.Label("Buscar por nombre", style={
                        "fontSize": "10px", "letterSpacing": ".18em",
                        "textTransform": "uppercase", "color": MUTED,
                        "marginBottom": "10px", "display": "block",
                    }),
                    dcc.Input(id="buscador_nombre", type="text",
                              placeholder="Escribe el nombre del estudiante...",
                              debounce=True, value=""),
                    html.Div(id="badge_busqueda"),
                ], style={"flex": "2", "minWidth": "220px"}),
                html.Div([
                    html.Label("Enfocar en materia", style={
                        "fontSize": "10px", "letterSpacing": ".18em",
                        "textTransform": "uppercase", "color": MUTED,
                        "marginBottom": "10px", "display": "block",
                    }),
                    dcc.Dropdown(id="filtro_materia", options=MATERIAS, value="todas",
                                 clearable=False,
                                 style={"backgroundColor": SURFACE, "color": TEXT,
                                        "border": f"1px solid {BORDER}"}),
                ], style={"flex": "1", "minWidth": "200px"}),
                html.Div([
                    html.Label("Ordenar por promedio", style={
                        "fontSize": "10px", "letterSpacing": ".18em",
                        "textTransform": "uppercase", "color": MUTED,
                        "marginBottom": "10px", "display": "block",
                    }),
                    dcc.Dropdown(id="orden_promedio", options=ORDEN_PROMEDIO, value="ninguno",
                                 clearable=False,
                                 style={"backgroundColor": SURFACE, "color": TEXT,
                                        "border": f"1px solid {BORDER}"}),
                ], style={"flex": "1", "minWidth": "200px"}),
            ], style={"display": "flex", "gap": "32px", "flexWrap": "wrap", "alignItems": "flex-start"}),
        ], style={**CARD, "borderLeft": f"3px solid {BLUE}", "margin": "16px 24px 0 24px"}),

        # Filtros originales
        html.Div([
            html.Div([
                html.Label("Carrera", style={
                    "fontSize": "10px", "letterSpacing": ".18em",
                    "textTransform": "uppercase", "color": MUTED,
                    "marginBottom": "10px", "display": "block",
                }),
                dcc.Dropdown(
                    id="filtro_carrera",
                    options=[{"label": c, "value": c} for c in sorted(dataf[C_CARRERA].unique())],
                    value=dataf[C_CARRERA].unique()[0],
                    style={"backgroundColor": SURFACE, "color": TEXT, "border": f"1px solid {BORDER}"},
                ),
            ], style={"flex": "1", "minWidth": "200px"}),
            html.Div([
                html.Label("Rango de edad", style={
                    "fontSize": "10px", "letterSpacing": ".18em",
                    "textTransform": "uppercase", "color": MUTED,
                    "marginBottom": "16px", "display": "block",
                }),
                dcc.RangeSlider(
                    id="slider_edad",
                    min=int(dataf[C_EDAD].min()), max=int(dataf[C_EDAD].max()), step=1,
                    value=[int(dataf[C_EDAD].min()), int(dataf[C_EDAD].max())],
                    tooltip={"placement": "bottom", "always_visible": True},
                ),
            ], style={"flex": "2", "minWidth": "240px"}),
            html.Div([
                html.Label("Rango de promedio", style={
                    "fontSize": "10px", "letterSpacing": ".18em",
                    "textTransform": "uppercase", "color": MUTED,
                    "marginBottom": "16px", "display": "block",
                }),
                dcc.RangeSlider(id="slider_promedio", min=0, max=5, step=0.5, value=[0, 5],
                                tooltip={"placement": "bottom", "always_visible": True}),
            ], style={"flex": "2", "minWidth": "240px"}),
        ], style={**CARD, "margin": "16px 24px",
                  "display": "flex", "gap": "40px", "flexWrap": "wrap", "alignItems": "flex-start"}),

        # KPIs
        html.Div(id="kpis", style={
            "display": "flex", "gap": "16px", "flexWrap": "wrap",
            "margin": "0 24px 16px 24px",
        }),

        # Tabla
        html.Div([
            html.P("REGISTROS", style={
                "fontSize": "10px", "letterSpacing": ".2em", "color": MUTED, "marginBottom": "16px",
            }),
            dcc.Loading(
                dash_table.DataTable(
                    id="tabla", page_size=8,
                    filter_action="native", sort_action="native",
                    row_selectable="multi", selected_rows=[],
                    style_table={"overflowX": "auto"},
                    style_cell={"textAlign": "center", "padding": "10px 16px",
                                "fontFamily": "'DM Mono', monospace"},
                    style_header={"fontWeight": "400"},
                    style_data_conditional=[
                        {"if": {"filter_query": "{promedio} >= 4", "column_id": "promedio"},
                         "color": GREEN, "fontWeight": "500"},
                        {"if": {"filter_query": "{promedio} < 2.5", "column_id": "promedio"},
                         "color": RED, "fontWeight": "500"},
                    ],
                ),
                type="circle", color=ACCENT,
            ),
        ], style={**CARD, "margin": "0 24px 16px 24px"}),

        # Gráfico detallado
        html.Div([
            html.P("ANÁLISIS DETALLADO — selecciona filas de la tabla", style={
                "fontSize": "10px", "letterSpacing": ".2em", "color": MUTED, "marginBottom": "8px",
            }),
            dcc.Loading(dcc.Graph(id="gra_detallado"), type="default", color=ACCENT),
        ], style={**CARD, "margin": "0 24px 16px 24px"}),

        # Tabs
        html.Div([
            dcc.Tabs(children=[
                dcc.Tab(label="HISTOGRAMA",  children=[dcc.Graph(id="histograma")]),
                dcc.Tab(label="DISPERSIÓN",  children=[dcc.Graph(id="dispersion")]),
                dcc.Tab(label="DESEMPEÑO",   children=[dcc.Graph(id="pie")]),
                dcc.Tab(label="POR CARRERA", children=[dcc.Graph(id="barras")]),
                dcc.Tab(label="POR MATERIA", children=[dcc.Graph(id="barras_materia")]),

                # ── Tab Ranking (Punto 5) ─────────────────────────────────────
                dcc.Tab(label="🏆 RANKING", children=[
                    html.Div([
                        html.P("TOP 10 — MEJORES PROMEDIOS", style={
                            "fontSize": "10px", "letterSpacing": ".2em",
                            "color": ACCENT, "marginBottom": "6px",
                        }),
                        html.P(
                            "Los 10 estudiantes con mayor promedio entre todas las carreras.",
                            style={"fontSize": "11px", "color": MUTED, "marginBottom": "20px"}
                        ),
                        html.Div(id="tabla_ranking"),
                        dcc.Graph(id="grafico_ranking"),
                    ], style={**CARD, "margin": "24px", "borderLeft": f"3px solid {ACCENT}"}),
                ]),

                # ── Tab Alertas (Punto 6) ────────────────────────────────────
                dcc.Tab(label="⚠ RIESGO", children=[
                    html.Div([
                        html.P("ESTUDIANTES EN RIESGO ACADÉMICO", style={
                            "fontSize": "10px", "letterSpacing": ".2em",
                            "color": RED, "marginBottom": "6px",
                        }),
                        html.P(
                            "Estudiantes con promedio menor a 3.0 — requieren atención inmediata.",
                            style={"fontSize": "11px", "color": MUTED, "marginBottom": "20px"}
                        ),
                        html.Div(id="alerta_riesgo"),
                    ], style={**CARD, "margin": "24px", "borderLeft": f"3px solid {RED}"}),
                ]),

                # ── Tab carga masiva ──────────────────────────────────────────
                dcc.Tab(label="CARGA MASIVA", children=[
                    html.Div([
                        html.P("CARGA MASIVA DESDE EXCEL", style={
                            "fontSize": "10px", "letterSpacing": ".2em",
                            "color": ACCENT, "marginBottom": "6px",
                        }),
                        html.P(
                            "Sube un .xlsx con columnas: Nombre, Edad, Carrera, nota1, nota2, nota3. "
                            "El sistema limpiará nombres, corregirá mayúsculas y omitirá duplicados automáticamente.",
                            style={"fontSize": "11px", "color": MUTED, "marginBottom": "24px"}
                        ),
                        dcc.Upload(
                            id="upload_excel",
                            children=html.Div([
                                html.Span("📂", style={"fontSize": "32px", "display": "block", "marginBottom": "10px"}),
                                html.Span("Arrastra tu archivo aquí o ", style={"color": MUTED}),
                                html.Span("haz clic para buscar", style={"color": ACCENT}),
                                html.Br(),
                                html.Span(".xlsx — máx. 5 MB", style={"fontSize": "10px", "color": MUTED, "letterSpacing": ".1em"}),
                            ]),
                            className="upload-zona", accept=".xlsx",
                            max_size=5 * 1024 * 1024, multiple=False,
                        ),
                        html.Div(id="panel_carga_msg",     style={"marginTop": "20px"}),
                        html.Div(id="upload_preview",      style={"marginTop": "16px"}),
                        html.Div(id="panel_estadisticas_carga", style={"marginTop": "16px"}),
                        html.Div(id="panel_rechazados",    style={"marginTop": "16px"}),
                        html.Div(
                            html.Button("⬆  GUARDAR EN BASE DE DATOS",
                                        id="btn_confirmar_carga", className="btn-registrar",
                                        style={"display": "none"}),
                            style={"marginTop": "20px"}
                        ),
                        # Stores
                        dcc.Store(id="store_excel"),
                        dcc.Store(id="store_excel_pendiente"),       # datos previos al modal (notas fuera rango)
                        dcc.Store(id="store_excel_pendiente_vacios"), # datos previos al modal (notas vacías)
                        dcc.Store(id="store_carga_completada", data=0),  # trigger recarga dashboard
                    ], style={**CARD, "margin": "24px", "borderLeft": f"3px solid {BLUE}"}),
                ]),

                # ── Tab registrar ─────────────────────────────────────────────
                dcc.Tab(label="REGISTRAR", children=[
                    html.Div([
                        html.P("NUEVO ESTUDIANTE", style={
                            "fontSize": "10px", "letterSpacing": ".2em",
                            "color": ACCENT, "marginBottom": "24px",
                        }),
                        html.Div([
                            html.Div([
                                html.Label("Nombre completo", style={
                                    "fontSize": "10px", "letterSpacing": ".18em",
                                    "textTransform": "uppercase", "color": MUTED,
                                    "marginBottom": "8px", "display": "block",
                                }),
                                dcc.Input(id="reg_nombre", type="text",
                                          placeholder="Ej: María García",
                                          className="input-registro"),
                            ], style={"flex": "2", "minWidth": "200px"}),
                            html.Div([
                                html.Label("Edad", style={
                                    "fontSize": "10px", "letterSpacing": ".18em",
                                    "textTransform": "uppercase", "color": MUTED,
                                    "marginBottom": "8px", "display": "block",
                                }),
                                dcc.Input(id="reg_edad", type="number",
                                          placeholder="18", min=10, max=99,
                                          className="input-registro"),
                            ], style={"flex": "1", "minWidth": "120px"}),
                            html.Div([
                                html.Label("Carrera", style={
                                    "fontSize": "10px", "letterSpacing": ".18em",
                                    "textTransform": "uppercase", "color": MUTED,
                                    "marginBottom": "8px", "display": "block",
                                }),
                                dcc.Input(id="reg_carrera", type="text",
                                          placeholder="Ej: Ingeniería",
                                          className="input-registro"),
                            ], style={"flex": "2", "minWidth": "180px"}),
                        ], style={"display": "flex", "gap": "20px",
                                  "flexWrap": "wrap", "marginBottom": "20px"}),

                        html.Div([
                            html.Div([
                                html.Label("Materia 1 (Nota 1)", style={
                                    "fontSize": "10px", "letterSpacing": ".18em",
                                    "textTransform": "uppercase", "color": MUTED,
                                    "marginBottom": "8px", "display": "block",
                                }),
                                dcc.Input(id="reg_nota1", type="number",
                                          placeholder="0.0 – 5.0", step=0.1,
                                          className="input-registro"),
                            ], style={"flex": "1", "minWidth": "140px"}),
                            html.Div([
                                html.Label("Materia 2 (Nota 2)", style={
                                    "fontSize": "10px", "letterSpacing": ".18em",
                                    "textTransform": "uppercase", "color": MUTED,
                                    "marginBottom": "8px", "display": "block",
                                }),
                                dcc.Input(id="reg_nota2", type="number",
                                          placeholder="0.0 – 5.0", step=0.1,
                                          className="input-registro"),
                            ], style={"flex": "1", "minWidth": "140px"}),
                            html.Div([
                                html.Label("Materia 3 (Nota 3)", style={
                                    "fontSize": "10px", "letterSpacing": ".18em",
                                    "textTransform": "uppercase", "color": MUTED,
                                    "marginBottom": "8px", "display": "block",
                                }),
                                dcc.Input(id="reg_nota3", type="number",
                                          placeholder="0.0 – 5.0", step=0.1,
                                          className="input-registro"),
                            ], style={"flex": "1", "minWidth": "140px"}),
                            html.Div([
                                html.Label("Promedio (auto)", style={
                                    "fontSize": "10px", "letterSpacing": ".18em",
                                    "textTransform": "uppercase", "color": MUTED,
                                    "marginBottom": "8px", "display": "block",
                                }),
                                html.Div(id="reg_preview_promedio", children="—", style={
                                    "background": SURFACE2, "border": f"1px solid {BORDER}",
                                    "color": ACCENT, "fontFamily": "'DM Serif Display', serif",
                                    "fontSize": "22px", "padding": "8px 14px",
                                    "minHeight": "42px", "display": "flex", "alignItems": "center",
                                }),
                            ], style={"flex": "1", "minWidth": "140px"}),
                        ], style={"display": "flex", "gap": "20px",
                                  "flexWrap": "wrap", "marginBottom": "28px"}),

                        html.Div([
                            html.Button("+ REGISTRAR ESTUDIANTE",
                                        id="btn_registrar", className="btn-registrar"),
                            html.Div(id="msg_registro", style={
                                "marginLeft": "20px", "fontSize": "12px",
                                "alignSelf": "center", "letterSpacing": ".08em",
                            }),
                        ], style={"display": "flex", "alignItems": "center",
                                  "flexWrap": "wrap", "gap": "12px"}),
                    ], style={**CARD, "margin": "24px", "borderLeft": f"3px solid {GREEN}"}),
                ]),
            ],
            style={"borderBottom": f"1px solid {BORDER}"},
            colors={"border": BORDER, "primary": ACCENT, "background": SURFACE},
            ),
        ], style={**CARD, "margin": "0 24px 24px 24px"}),

        html.Div("© Sistema Académico", style={
            "textAlign": "center", "fontSize": "10px", "color": "#2a2d33",
            "letterSpacing": ".14em", "textTransform": "uppercase",
            "padding": "16px 0 32px 0",
        }),
    ])

    # ══════════════════════════════════════════════════════════════════════════
    # CALLBACKS
    # ══════════════════════════════════════════════════════════════════════════

    @appnotas.callback(
        Output("tabla", "data"),
        Output("tabla", "columns"),
        Output("kpis", "children"),
        Output("histograma", "figure"),
        Output("dispersion", "figure"),
        Output("pie", "figure"),
        Output("barras", "figure"),
        Output("barras_materia", "figure"),
        Output("badge_busqueda", "children"),
        Output("filtro_carrera", "value"),
        Output("buscador_nombre", "value"),
        Output("orden_promedio", "value"),
        Input("filtro_carrera", "value"),
        Input("slider_edad", "value"),
        Input("slider_promedio", "value"),
        Input("buscador_nombre", "value"),
        Input("filtro_materia", "value"),
        Input("orden_promedio", "value"),
        Input("store_carga_completada", "data"),
    )
    def actualizar_comp(carrera, rangoedad, rangoprome, nombre_busqueda, materia, orden, trigger_carga):
        from dash import callback_context as ctx
        dataf = cargar_datos()

        # Punto 2: si el trigger viene de carga masiva, resetear filtros
        reset_triggered = any(
            t["prop_id"] == "store_carga_completada.data"
            for t in ctx.triggered
        ) if ctx.triggered else False

        if reset_triggered and not dataf.empty:
            carrera = dataf[C_CARRERA].unique()[0]
            nombre_busqueda = ""
            orden = "ninguno"

        filtro = dataf[
            (dataf[C_CARRERA]  == carrera)      &
            (dataf[C_EDAD]     >= rangoedad[0])  &
            (dataf[C_EDAD]     <= rangoedad[1])  &
            (dataf[C_PROMEDIO] >= rangoprome[0]) &
            (dataf[C_PROMEDIO] <= rangoprome[1])
        ].copy()

        badge_msg = ""
        if nombre_busqueda and nombre_busqueda.strip():
            termino = nombre_busqueda.strip()
            mascara = filtro[C_NOMBRE].str.contains(termino, case=False, na=False)
            filtro  = filtro[mascara]
            encontrados = len(filtro)
            if encontrados == 0:
                badge_msg = f"⚠ Sin resultados para «{termino}»"
            elif encontrados == 1:
                badge_msg = "✓ 1 estudiante encontrado para «" + termino + "»"
            else:
                badge_msg = f"✓ {encontrados} estudiantes encontrados para «{termino}»"

        if orden == "desc":
            filtro = filtro.sort_values(C_PROMEDIO, ascending=False)
        elif orden == "asc":
            filtro = filtro.sort_values(C_PROMEDIO, ascending=True)

        n = len(filtro)
        if materia != "todas" and materia in filtro.columns:
            col_kpi = materia
            label_kpi = f"Promedio {materia.upper()}"
        else:
            col_kpi   = C_PROMEDIO
            label_kpi = "Promedio"

        promedio = round(filtro[col_kpi].mean(), 2) if n > 0 else 0
        maximo   = round(filtro[col_kpi].max(),  2) if n > 0 else 0
        minimo   = round(filtro[col_kpi].min(),  2) if n > 0 else 0

        kpis = [
            kpi_card(label_kpi,     promedio, ACCENT),
            kpi_card("Estudiantes", n,        BLUE),
            kpi_card("Máximo",      maximo,   GREEN),
            kpi_card("Mínimo",      minimo,   RED),
        ]

        col_hist   = col_kpi
        label_hist = "Nota" if materia != "todas" else "Promedio"
        histo = px.histogram(filtro, x=col_hist, nbins=10,
                             color_discrete_sequence=[ACCENT],
                             labels={col_hist: label_hist})
        histo.update_traces(marker_line_color=BG, marker_line_width=1)
        titulo_hist = (f"Distribución de {materia.upper()}" if materia != "todas"
                       else "Distribución de Promedios")
        apply_template(histo, titulo_hist)

        dispersion = px.scatter(filtro, x=C_EDAD, y=C_PROMEDIO, color=C_DESEMPEN,
                                labels={C_EDAD: "Edad", C_PROMEDIO: "Promedio", C_DESEMPEN: "Desempeño"},
                                color_discrete_sequence=[ACCENT, BLUE, GREEN, RED])
        apply_template(dispersion, "Edad vs Promedio")

        pie = px.pie(filtro, names=C_DESEMPEN,
                     color_discrete_sequence=[ACCENT, BLUE, GREEN, RED, "#a78bfa"], hole=0.45)
        pie.update_traces(textfont=dict(family="'DM Mono', monospace", size=12))
        apply_template(pie, "Distribución por Desempeño")

        promedios = dataf.groupby(C_CARRERA)[C_PROMEDIO].mean().reset_index()
        barras = px.bar(promedios, x=C_CARRERA, y=C_PROMEDIO, color=C_CARRERA,
                        labels={C_CARRERA: "Carrera", C_PROMEDIO: "Promedio"},
                        color_discrete_sequence=[ACCENT, BLUE, GREEN, RED, "#a78bfa"])
        barras.update_traces(marker_line_color=BG, marker_line_width=1)
        apply_template(barras, "Promedio General por Carrera")

        if n > 0 and all(c in filtro.columns for c in [C_NOTA1, C_NOTA2, C_NOTA3]):
            df_materias = pd.DataFrame({
                "Materia": ["Materia 1", "Materia 2", "Materia 3"],
                "Promedio": [round(filtro[C_NOTA1].mean(), 2),
                             round(filtro[C_NOTA2].mean(), 2),
                             round(filtro[C_NOTA3].mean(), 2)],
            })
            colores = [MUTED, MUTED, MUTED]
            if materia == C_NOTA1:   colores[0] = ACCENT
            elif materia == C_NOTA2: colores[1] = ACCENT
            elif materia == C_NOTA3: colores[2] = ACCENT
            else: colores = [ACCENT, BLUE, GREEN]

            barras_mat = px.bar(df_materias, x="Materia", y="Promedio", text="Promedio",
                                color="Materia", color_discrete_sequence=colores)
            barras_mat.update_traces(marker_line_color=BG, marker_line_width=1,
                                     textposition="outside", textfont=dict(color=TEXT, size=13))
            apply_template(barras_mat, "Comparación de Notas por Materia")
        else:
            barras_mat = go.Figure()
            apply_template(barras_mat, "Comparación de Notas por Materia")

        col_labels = {
            "id": "ID", "nombre": "Nombre", "edad": "Edad", "carrera": "Carrera",
            "nota1": "Materia 1", "nota2": "Materia 2", "nota3": "Materia 3",
            "promedio": "Promedio", "desempenio": "Desempeño",
        }
        cols_orden = list(filtro.columns)
        if materia != "todas" and materia in cols_orden:
            cols_orden.remove(materia)
            idx = cols_orden.index(C_NOMBRE) + 1 if C_NOMBRE in cols_orden else 1
            cols_orden.insert(idx, materia)

        columns = [{"name": col_labels.get(c, c.title()), "id": c} for c in cols_orden]
        return (filtro[cols_orden].to_dict("records"), columns, kpis,
                histo, dispersion, pie, barras, barras_mat, badge_msg,
                carrera, nombre_busqueda, orden)

    # ── Gráfico detallado ─────────────────────────────────────────────────────
    @appnotas.callback(
        Output("gra_detallado", "figure"),
        Input("tabla", "derived_virtual_data"),
        Input("tabla", "derived_virtual_selected_rows"),
        Input("filtro_materia", "value"),
    )
    def actualizartab(rows, selected_rows, materia):
        if not rows:
            fig = go.Figure()
            fig.update_layout(
                **PLOTLY_LAYOUT_BASE,
                title=dict(text="Sin datos", font=TITLE_FONT),
                annotations=[dict(
                    text="Selecciona filas de la tabla para el análisis detallado",
                    showarrow=False,
                    font=dict(color=MUTED, size=13, family="'DM Mono', monospace"),
                    xref="paper", yref="paper", x=0.5, y=0.5,
                )],
            )
            return fig

        dff = pd.DataFrame(rows)
        if selected_rows:
            dff = dff.iloc[selected_rows]

        if materia != "todas" and materia in dff.columns and C_NOMBRE in dff.columns:
            fig = px.bar(dff, x=C_NOMBRE, y=materia,
                         color=C_DESEMPEN if C_DESEMPEN in dff.columns else None,
                         labels={C_NOMBRE: "Estudiante", materia: f"Nota {materia.upper()}",
                                 C_DESEMPEN: "Desempeño"},
                         color_discrete_sequence=[ACCENT, BLUE, GREEN, RED], text=materia)
            fig.update_traces(marker_line_color=BG, marker_line_width=1,
                              textposition="outside", textfont=dict(color=TEXT))
            apply_template(fig, f"Notas en {materia.upper()} — estudiantes seleccionados")
        else:
            fig = px.scatter(dff, x=C_EDAD, y=C_PROMEDIO, color=C_DESEMPEN, size=C_PROMEDIO,
                             labels={C_EDAD: "Edad", C_PROMEDIO: "Promedio", C_DESEMPEN: "Desempeño"},
                             color_discrete_sequence=[ACCENT, BLUE, GREEN, RED])
            apply_template(fig, "Análisis detallado")
        return fig

    # ── Preview promedio ──────────────────────────────────────────────────────
    @appnotas.callback(
        Output("reg_preview_promedio", "children"),
        Output("reg_preview_promedio", "style"),
        Input("reg_nota1", "value"),
        Input("reg_nota2", "value"),
        Input("reg_nota3", "value"),
    )
    def preview_promedio(n1, n2, n3):
        base_style = {
            "background": SURFACE2, "border": f"1px solid {BORDER}",
            "fontFamily": "'DM Serif Display', serif", "fontSize": "22px",
            "padding": "8px 14px", "minHeight": "42px",
            "display": "flex", "alignItems": "center",
        }
        notas = [x for x in [n1, n2, n3] if x is not None]
        if not notas:
            return "—", {**base_style, "color": MUTED}
        prom  = round(sum(notas) / len(notas), 2)
        color = GREEN if prom >= 4 else (RED if prom < 3 else ACCENT)
        return str(prom), {**base_style, "color": color}

    # ── Registrar estudiante individual ───────────────────────────────────────
    @appnotas.callback(
        Output("msg_registro",  "children"),
        Output("msg_registro",  "style"),
        Output("reg_nombre",    "value"),
        Output("reg_edad",      "value"),
        Output("reg_carrera",   "value"),
        Output("reg_nota1",     "value"),
        Output("reg_nota2",     "value"),
        Output("reg_nota3",     "value"),
        Input("btn_registrar",  "n_clicks"),
        State("reg_nombre",     "value"),
        State("reg_edad",       "value"),
        State("reg_carrera",    "value"),
        State("reg_nota1",      "value"),
        State("reg_nota2",      "value"),
        State("reg_nota3",      "value"),
        prevent_initial_call=True,
    )
    def registrar_estudiante(n_clicks, nombre, edad, carrera, nota1, nota2, nota3):
        estilo_base = {
            "marginLeft": "20px", "fontSize": "12px", "alignSelf": "center",
            "letterSpacing": ".08em", "padding": "8px 14px", "border": "1px solid",
        }

        if not all([nombre, edad, carrera, nota1 is not None,
                    nota2 is not None, nota3 is not None]):
            return ("⚠ Completa todos los campos antes de registrar.",
                    {**estilo_base, "color": RED, "borderColor": RED,
                     "background": "rgba(224,92,92,.07)"},
                    nombre, edad, carrera, nota1, nota2, nota3)

        # ── Normalizar nombre ─────────────────────────────────────────────────
        nombre_limpio = normalizar_nombre(nombre)

        # ── Ajustar notas al rango 0–5 ────────────────────────────────────────
        nota1 = max(0.0, min(5.0, float(nota1)))
        nota2 = max(0.0, min(5.0, float(nota2)))
        nota3 = max(0.0, min(5.0, float(nota3)))

        # ── Verificar duplicado (Nombre + Carrera) ────────────────────────────
        if existe_estudiante(nombre_limpio, edad, carrera):
            return (
                f"⚠ Ya existe '{nombre_limpio}' en {carrera.strip()} — estudiante duplicado.",
                {**estilo_base, "color": ORANGE, "borderColor": ORANGE,
                 "background": "rgba(251,146,60,.07)"},
                nombre, edad, carrera, nota1, nota2, nota3,
            )

        promedio   = round((nota1 + nota2 + nota3) / 3, 2)
        desempenio = calc_desempenio(promedio)

        try:
            insertar_estudiante(nombre_limpio, int(edad), carrera.strip().title(),
                                nota1, nota2, nota3, promedio, desempenio)
            return (
                f"✓ {nombre_limpio} registrado — promedio {promedio} ({desempenio})",
                {**estilo_base, "color": GREEN, "borderColor": GREEN,
                 "background": "rgba(74,222,128,.07)"},
                "", None, "", None, None, None,
            )
        except Exception as e:
            return (f"✗ Error al guardar: {str(e)}",
                    {**estilo_base, "color": RED, "borderColor": RED,
                     "background": "rgba(224,92,92,.07)"},
                    nombre, edad, carrera, nota1, nota2, nota3)

    # ── Parsear Excel: limpiar, detectar, mostrar preview ─────────────────────
    # Este callback lee el archivo, hace TODA la limpieza y:
    #  a) si hay notas fuera de rango → muestra el modal (store_excel_pendiente)
    #  b) si todo está bien            → llena store_excel directamente
    @appnotas.callback(
        Output("panel_carga_msg",        "children"),
        Output("upload_preview",         "children"),
        Output("btn_confirmar_carga",    "style"),
        Output("store_excel",            "data"),
        Output("store_excel_pendiente",  "data"),
        Output("store_excel_pendiente_vacios", "data"),
        Output("modal_notas",            "style"),
        Output("modal_notas_msg",        "children"),
        Input("upload_excel",            "contents"),
        State("upload_excel",            "filename"),
        prevent_initial_call=True,
    )
    def previsualizar_excel(contents, filename):
        import base64, io

        BTN_OCULTO  = {"display": "none"}
        BTN_VISIBLE = {
            "display": "inline-block", "background": ACCENT, "color": "#0c0e10",
            "border": "none", "fontFamily": "'DM Mono', monospace",
            "fontSize": "11px", "fontWeight": "500", "letterSpacing": ".2em",
            "textTransform": "uppercase", "padding": "14px 32px", "cursor": "pointer",
        }
        MODAL_OCULTO  = {"display": "none"}
        MODAL_VISIBLE = {}   # usa la clase .modal-overlay que tiene display:flex

        def panel_err(titulo, detalle="", sugerencia=""):
            hijos = [
                html.Div(html.Span("✕  ERROR AL CARGAR", style={
                    "fontSize": "11px", "letterSpacing": ".2em", "color": RED, "fontWeight": "500",
                }), style={"marginBottom": "10px"}),
                html.P(titulo, style={"color": TEXT, "fontSize": "13px",
                                      "marginBottom": "6px", "lineHeight": "1.6"}),
            ]
            if detalle:
                hijos.append(html.P(detalle, style={"color": MUTED, "fontSize": "12px",
                                                     "marginBottom": "8px", "lineHeight": "1.5"}))
            if sugerencia:
                hijos.append(html.Div([html.Span("💡 "),
                    html.Span(sugerencia, style={"color": ACCENT, "fontSize": "11px"})],
                    style={"marginTop": "10px", "padding": "10px 14px",
                           "background": "rgba(200,169,110,.07)",
                           "border": "1px solid rgba(200,169,110,.2)"}))
            return html.Div(hijos, style={
                "background": "rgba(224,92,92,.08)", "border": f"1px solid {RED}",
                "borderLeft": f"4px solid {RED}", "padding": "18px 20px",
            })

        def panel_ok(titulo, detalle=""):
            return html.Div([
                html.Div(html.Span("✓  ARCHIVO VÁLIDO", style={
                    "fontSize": "11px", "letterSpacing": ".2em", "color": GREEN, "fontWeight": "500",
                }), style={"marginBottom": "8px"}),
                html.P(titulo, style={"color": TEXT, "fontSize": "13px"}),
                html.P(detalle, style={"color": MUTED, "fontSize": "11px", "marginTop": "4px"}),
            ], style={"background": "rgba(74,222,128,.07)", "border": f"1px solid {GREEN}",
                      "borderLeft": f"4px solid {GREEN}", "padding": "18px 20px"})

        def panel_warn(titulo, detalle=""):
            return html.Div([
                html.Div(html.Span("⚠  ATENCIÓN", style={
                    "fontSize": "11px", "letterSpacing": ".2em", "color": ORANGE, "fontWeight": "500",
                }), style={"marginBottom": "8px"}),
                html.P(titulo, style={"color": TEXT, "fontSize": "13px"}),
                html.P(detalle, style={"color": MUTED, "fontSize": "11px", "marginTop": "4px"}),
            ], style={"background": "rgba(251,146,60,.07)", "border": f"1px solid {ORANGE}",
                      "borderLeft": f"4px solid {ORANGE}", "padding": "18px 20px"})

        def err(titulo, detalle="", sugerencia=""):
            return (panel_err(titulo, detalle, sugerencia),
                    "", BTN_OCULTO, None, None, None, MODAL_OCULTO, "")

        if contents is None:
            return "", "", BTN_OCULTO, None, None, None, MODAL_OCULTO, ""

        if filename and not filename.lower().endswith(".xlsx"):
            ext = filename.rsplit(".", 1)[-1] if "." in filename else "desconocida"
            return err(f"El archivo '{filename}' no es compatible.",
                       f"Extensión '.{ext}' no aceptada. Solo .xlsx.",
                       "Guarda como 'Libro de Excel (.xlsx)' e inténtalo de nuevo.")

        try:
            _, content_string = contents.split(",", 1)
            decoded = base64.b64decode(content_string)
        except Exception:
            return err("El archivo está dañado o no se pudo leer.",
                       "No fue posible decodificar el contenido.",
                       "Ábrelo en Excel, guárdalo de nuevo y vuelve a subirlo.")

        try:
            df_up = pd.read_excel(io.BytesIO(decoded))
        except Exception as e:
            return err("No se pudo abrir el archivo Excel.", f"Detalle: {e}",
                       "Verifica que no esté abierto, protegido, o sea un .xlsx estándar.")

        if df_up.empty:
            return err("El archivo está vacío.", "No hay filas de datos.",
                       "Agrega al menos un estudiante y vuelve a subir.")

        if len(df_up.columns) < 3:
            return err(f"Solo {len(df_up.columns)} columna(s) detectada(s).",
                       f"Columnas: {', '.join(str(c) for c in df_up.columns)}.",
                       "Necesitas: Nombre, Edad, Carrera, nota1, nota2, nota3.")

        # ── Normalizar columnas ───────────────────────────────────────────────
        col_norm_map = {}
        for col in df_up.columns:
            norm = limpiar_col(col)
            if norm in col_norm_map:
                return err("Columnas duplicadas detectadas.",
                           f"'{col}' y '{col_norm_map[norm]}' se interpretan igual.",
                           "Renombra o elimina una de las columnas repetidas.")
            col_norm_map[norm] = col

        alias = {
            "nombre": C_NOMBRE, "name": C_NOMBRE, "estudiante": C_NOMBRE,
            "nombrecompleto": C_NOMBRE, "alumno": C_NOMBRE,
            "edad": C_EDAD, "age": C_EDAD, "anos": C_EDAD,
            "carrera": C_CARRERA, "career": C_CARRERA, "programa": C_CARRERA,
            "nota1": C_NOTA1, "note1": C_NOTA1, "n1": C_NOTA1,
            "materia1": C_NOTA1, "mat1": C_NOTA1, "p1": C_NOTA1,
            "nota2": C_NOTA2, "note2": C_NOTA2, "n2": C_NOTA2,
            "materia2": C_NOTA2, "mat2": C_NOTA2, "p2": C_NOTA2,
            "nota3": C_NOTA3, "note3": C_NOTA3, "n3": C_NOTA3,
            "materia3": C_NOTA3, "mat3": C_NOTA3, "p3": C_NOTA3,
        }

        rename = {}
        extras = []
        for norm, col_orig in col_norm_map.items():
            if norm in alias:
                rename[col_orig] = alias[norm]
            else:
                extras.append(col_orig)

        df_up = df_up.rename(columns=rename)

        requeridas  = [C_NOMBRE, C_EDAD, C_CARRERA, C_NOTA1, C_NOTA2, C_NOTA3]
        nombres_leg = {C_NOMBRE:"Nombre", C_EDAD:"Edad", C_CARRERA:"Carrera",
                       C_NOTA1:"nota1", C_NOTA2:"nota2", C_NOTA3:"nota3"}
        faltantes   = [c for c in requeridas if c not in df_up.columns]
        if faltantes:
            extra_txt = f" No reconocidas: {', '.join(extras)}." if extras else ""
            return err(
                f"Faltan columnas: {', '.join(nombres_leg[c] for c in faltantes)}.",
                f"Tu archivo tiene: {', '.join(str(c) for c in col_norm_map.values())}.{extra_txt}",
                "Columnas requeridas: Nombre, Edad, Carrera, nota1, nota2, nota3."
            )

        df_up = df_up[requeridas].copy()
        for col in [C_NOTA1, C_NOTA2, C_NOTA3, C_EDAD]:
            df_up[col] = pd.to_numeric(df_up[col], errors="coerce")
        df_up[C_NOMBRE]  = df_up[C_NOMBRE].astype(str).str.strip()
        df_up[C_CARRERA] = df_up[C_CARRERA].astype(str).str.strip()

        # Filas donde notas o edad no son numéricas (coerce las convirtió a NaN)
        mask_no_num = df_up[[C_NOTA1, C_NOTA2, C_NOTA3, C_EDAD]].isnull().any(axis=1)
        n_no_num    = mask_no_num.sum()
        nombres_no_num = df_up[mask_no_num][C_NOMBRE].tolist()
        if n_no_num > 0:
            df_up = df_up[~mask_no_num].copy()

        if df_up.empty:
            return err(
                "No quedan filas válidas tras eliminar registros con valores no numéricos.",
                f"Filas con texto en notas o edad ({n_no_num}): {', '.join(nombres_no_num)}.",
                "Las columnas nota1, nota2, nota3 y Edad deben contener solo números."
            )

        # ══════════════════════════════════════════════════════════════════════
        # PASO 1: LIMPIEZA AUTOMÁTICA DEL DOCUMENTO
        # ══════════════════════════════════════════════════════════════════════
        n_total_original = len(df_up)

        # 1a. Eliminar filas completamente vacías
        df_up = df_up.dropna(how="all")

        # 1b. Filas donde Nombre o Carrera están vacíos → eliminar siempre
        mask_sin_nombre = ((df_up[C_NOMBRE] == "") | (df_up[C_NOMBRE] == "nan") |
                           df_up[C_NOMBRE].isnull())
        n_invalidas     = int(mask_sin_nombre.sum())
        df_up           = df_up[~mask_sin_nombre].copy()

        if df_up.empty:
            return err("Ninguna fila tiene nombre válido.",
                       "Todas las filas tienen el campo Nombre vacío o inválido.",
                       "Completa la columna Nombre en el archivo.")

        # 1c. Normalizar nombres: Title Case + espacios limpios
        nombres_originales   = df_up[C_NOMBRE].copy()
        df_up[C_NOMBRE]      = df_up[C_NOMBRE].apply(normalizar_nombre)
        n_nombres_corregidos = int((nombres_originales != df_up[C_NOMBRE]).sum())

        # 1d. Normalizar carrera: Title Case
        carreras_originales   = df_up[C_CARRERA].copy()
        df_up[C_CARRERA]      = df_up[C_CARRERA].apply(
            lambda s: s.strip().title() if isinstance(s, str) else s
        )
        n_carreras_corregidas = int((carreras_originales != df_up[C_CARRERA]).sum())

        # 1e. Edades inválidas (negativas o > 99) → rechazar fila directamente
        mask_edad_inv    = df_up[C_EDAD].lt(0) | df_up[C_EDAD].gt(99) | df_up[C_EDAD].isnull()
        n_edad_inv       = int(mask_edad_inv.sum())
        nombres_edad_inv = df_up[mask_edad_inv][C_NOMBRE].tolist()
        if n_edad_inv > 0:
            df_up = df_up[~mask_edad_inv].copy()

        if df_up.empty:
            return err("Ninguna fila tiene edad válida.",
                       f"Filas con edad negativa o inválida ({n_edad_inv}): {', '.join(nombres_edad_inv)}.",
                       "La columna Edad debe contener valores entre 0 y 99.")

        # 1f. Ajustar notas fuera de rango (< 0 → 0, > 5 → 5) AUTOMÁTICAMENTE
        #     y registrar cuántos estudiantes fueron afectados
        mask_notas_fuera  = (
            df_up[C_NOTA1].lt(0) | df_up[C_NOTA1].gt(5) |
            df_up[C_NOTA2].lt(0) | df_up[C_NOTA2].gt(5) |
            df_up[C_NOTA3].lt(0) | df_up[C_NOTA3].gt(5)
        )
        mask_notas_vacias = df_up[[C_NOTA1, C_NOTA2, C_NOTA3]].isnull().any(axis=1)
        afectados_fuera   = df_up[mask_notas_fuera][C_NOMBRE].tolist()
        afectados_vacias  = df_up[mask_notas_vacias][C_NOMBRE].tolist()
        n_notas_ajustadas = int(mask_notas_fuera.sum())
        n_notas_vacias    = int(mask_notas_vacias.sum())

        # Ajustar automáticamente: vacíos → 0, fuera de rango → clip(0,5)
        for c in [C_NOTA1, C_NOTA2, C_NOTA3]:
            df_up[c] = df_up[c].fillna(0).clip(0, 5)

        # 1g. Calcular promedio y desempeño finales
        df_up[C_PROMEDIO] = ((df_up[C_NOTA1] + df_up[C_NOTA2] + df_up[C_NOTA3]) / 3).round(2)
        df_up[C_DESEMPEN] = df_up[C_PROMEDIO].apply(calc_desempenio)

        # ══════════════════════════════════════════════════════════════════════
        # PASO 2: DEDUPLICACIÓN INTERNA — clave: Nombre+Carrera+Edad
        # ══════════════════════════════════════════════════════════════════════
        antes_dup = len(df_up)
        clave = df_up[[C_NOMBRE, C_CARRERA]].apply(
            lambda r: (r[C_NOMBRE].lower(), r[C_CARRERA].lower()), axis=1
        )
        df_up = df_up[~clave.duplicated(keep="first")].copy()
        n_dup_internos = antes_dup - len(df_up)

        # ══════════════════════════════════════════════════════════════════════
        # PASO 3: DETECTAR DUPLICADOS CONTRA BD — clave: Nombre+Carrera+Edad
        # ══════════════════════════════════════════════════════════════════════
        existentes_bd = obtener_claves_existentes()

        def clave_fila(row):
            return (str(row[C_NOMBRE]).strip().lower(),
                    str(row[C_CARRERA]).strip().lower())

        claves_vistas = []
        estados = []
        for _, row in df_up.iterrows():
            ck = clave_fila(row)
            if ck in existentes_bd:
                estados.append("⚠ Ya existe en BD")
            else:
                estados.append("✓ Nuevo")
                claves_vistas.append(ck)

        df_up["_estado"] = estados
        df_nuevos = df_up[df_up["_estado"] == "✓ Nuevo"].copy()
        df_dup_bd = df_up[df_up["_estado"] == "⚠ Ya existe en BD"].copy()
        n_nuevos  = len(df_nuevos)
        n_dup_bd  = len(df_dup_bd)

        if n_nuevos == 0:
            msgs = []
            if n_dup_bd:
                msgs.append(f"{n_dup_bd} ya existe(n) en la BD")
            return (
                panel_warn("No hay registros nuevos para guardar.",
                           ". ".join(msgs) + ". Revisa y corrige el archivo."),
                "", BTN_OCULTO, None, None, None, MODAL_OCULTO, ""
            )

        # ══════════════════════════════════════════════════════════════════════
        # PASO 4: Construir resumen de limpieza (badges)
        # ══════════════════════════════════════════════════════════════════════
        badges_limpieza = []
        if n_invalidas > 0:
            badges_limpieza.append(
                html.Span(f"✂ {n_invalidas} fila(s) incompletas eliminadas",
                          className="badge-limpieza badge-naranja"))
        if n_no_num > 0:
            badges_limpieza.append(
                html.Span(
                    f"✂ {n_no_num} fila(s) con valores no numéricos eliminadas: "
                    f"{', '.join(nombres_no_num)}",
                    className="badge-limpieza badge-naranja"))
        if n_edad_inv > 0:
            badges_limpieza.append(
                html.Span(
                    f"✂ {n_edad_inv} fila(s) con edad inválida rechazadas: "
                    f"{', '.join(nombres_edad_inv)}",
                    className="badge-limpieza badge-naranja"))
        if n_nombres_corregidos > 0:
            badges_limpieza.append(
                html.Span(f"✎ {n_nombres_corregidos} nombre(s) normalizados a Title Case",
                          className="badge-limpieza badge-verde"))
        if n_carreras_corregidas > 0:
            badges_limpieza.append(
                html.Span(f"✎ {n_carreras_corregidas} carrera(s) normalizadas",
                          className="badge-limpieza badge-verde"))
        if n_dup_internos > 0:
            badges_limpieza.append(
                html.Span(f"✂ {n_dup_internos} duplicado(s) internos eliminados",
                          className="badge-limpieza badge-naranja"))
        if n_notas_ajustadas > 0:
            badges_limpieza.append(
                html.Span(
                    f"✎ {n_notas_ajustadas} estudiante(s) con notas ajustadas al rango 0–5: "
                    f"{', '.join(dict.fromkeys(afectados_fuera))}",
                    className="badge-limpieza badge-verde"))
        if n_notas_vacias > 0:
            badges_limpieza.append(
                html.Span(
                    f"✎ {n_notas_vacias} nota(s) vacía(s) rellenadas con 0: "
                    f"{', '.join(dict.fromkeys(afectados_vacias))}",
                    className="badge-limpieza badge-verde"))
        if n_dup_bd > 0:
            badges_limpieza.append(
                html.Span(f"⊘ {n_dup_bd} ya existe(n) en BD — omitido(s)",
                          className="badge-limpieza badge-naranja"))
        if extras:
            badges_limpieza.append(
                html.Span(f"ℹ Columnas ignoradas: {', '.join(extras)}",
                          className="badge-limpieza",
                          style={"color": BLUE, "border": f"1px solid {BLUE}",
                                 "background": "rgba(96,165,250,.06)"}))

        detalle_str = (
            f"Promedio grupo: {round(df_nuevos[C_PROMEDIO].mean(), 2)} | "
            f"Mejor: {df_nuevos[C_PROMEDIO].max()} | Menor: {df_nuevos[C_PROMEDIO].min()}"
        )

        panel = html.Div([
            html.Div(html.Span("✓  ARCHIVO VÁLIDO", style={
                "fontSize": "11px", "letterSpacing": ".2em", "color": GREEN, "fontWeight": "500",
            }), style={"marginBottom": "12px"}),
            html.P(f"{n_nuevos} estudiante(s) nuevos listos para cargar.",
                   style={"color": TEXT, "fontSize": "13px", "marginBottom": "4px"}),
            html.P(detalle_str, style={"color": MUTED, "fontSize": "11px", "marginBottom": "16px"}),
            html.Div([
                html.P("RESUMEN DE LIMPIEZA AUTOMÁTICA", style={
                    "fontSize": "10px", "letterSpacing": ".18em", "color": MUTED,
                    "marginBottom": "10px", "textTransform": "uppercase",
                }),
                html.Div(badges_limpieza if badges_limpieza
                         else [html.Span("✓ Sin cambios necesarios",
                                         className="badge-limpieza badge-verde")]),
            ], style={"padding": "14px 16px",
                      "background": SURFACE2, "border": f"1px solid {BORDER}"}),
        ], style={"background": "rgba(74,222,128,.07)", "border": f"1px solid {GREEN}",
                  "borderLeft": f"4px solid {GREEN}", "padding": "18px 20px"})

        # ── Tabla preview ─────────────────────────────────────────────────────
        col_labels_prev = {
            C_NOMBRE:"Nombre", C_EDAD:"Edad", C_CARRERA:"Carrera",
            C_NOTA1:"Mat. 1", C_NOTA2:"Mat. 2", C_NOTA3:"Mat. 3",
            C_PROMEDIO:"Promedio", C_DESEMPEN:"Desempeño", "_estado": "Estado",
        }
        cols_preview = [C_NOMBRE, C_EDAD, C_CARRERA, C_NOTA1, C_NOTA2,
                        C_NOTA3, C_PROMEDIO, C_DESEMPEN, "_estado"]

        preview = html.Div([
            html.P("PREVISUALIZACIÓN — filas ⚠ serán omitidas", style={
                "fontSize": "10px", "letterSpacing": ".18em",
                "textTransform": "uppercase", "color": MUTED, "marginBottom": "12px",
            }),
            dash_table.DataTable(
                data=df_up[cols_preview].to_dict("records"),
                columns=[{"name": col_labels_prev.get(c, c), "id": c} for c in cols_preview],
                page_size=10,
                style_table={"overflowX": "auto"},
                style_cell={"textAlign": "center", "padding": "8px 14px",
                            "fontFamily": "'DM Mono', monospace", "fontSize": "12px",
                            "backgroundColor": SURFACE, "color": TEXT,
                            "border": f"1px solid {BORDER}"},
                style_header={"backgroundColor": SURFACE2, "color": MUTED,
                              "fontWeight": "400", "fontSize": "10px",
                              "letterSpacing": ".12em", "textTransform": "uppercase",
                              "border": f"1px solid {BORDER}"},
                style_data_conditional=[
                    {"if": {"filter_query": '{_estado} contains "⚠"'},
                     "backgroundColor": "rgba(251,146,60,.08)",
                     "color": MUTED, "fontStyle": "italic"},
                    {"if": {"filter_query": '{_estado} = "✓ Nuevo"', "column_id": "_estado"},
                     "color": GREEN, "fontWeight": "500"},
                    {"if": {"filter_query": '{_estado} contains "⚠"', "column_id": "_estado"},
                     "color": ORANGE, "fontWeight": "500"},
                    {"if": {"filter_query": "{promedio} >= 4 && {_estado} = '✓ Nuevo'",
                            "column_id": "promedio"},
                     "color": GREEN, "fontWeight": "500"},
                    {"if": {"filter_query": "{promedio} < 3 && {_estado} = '✓ Nuevo'",
                            "column_id": "promedio"},
                     "color": RED, "fontWeight": "500"},
                ],
            ),
        ])

        # Sin modal: las notas ya fueron ajustadas automáticamente → directo al store
        store_data = df_nuevos.drop(columns=["_estado"], errors="ignore").to_dict("records")
        return panel, preview, BTN_VISIBLE, store_data, None, None, MODAL_OCULTO, ""


    # ── Modal: ajustar notas al límite 0–5 ───────────────────────────────────
    @appnotas.callback(
        Output("store_excel",                "data",  allow_duplicate=True),
        Output("store_excel_pendiente",      "data",  allow_duplicate=True),
        Output("store_excel_pendiente_vacios","data", allow_duplicate=True),
        Output("modal_notas",                "style", allow_duplicate=True),
        Output("btn_confirmar_carga",        "style", allow_duplicate=True),
        Input("modal_btn_aceptar",           "n_clicks"),
        State("store_excel_pendiente",       "data"),
        prevent_initial_call=True,
    )
    def modal_aceptar(n_clicks, pendiente):
        BTN_VISIBLE = {
            "display": "inline-block", "background": ACCENT, "color": "#0c0e10",
            "border": "none", "fontFamily": "'DM Mono', monospace",
            "fontSize": "11px", "fontWeight": "500", "letterSpacing": ".2em",
            "textTransform": "uppercase", "padding": "14px 32px", "cursor": "pointer",
        }
        if not pendiente:
            return dash.no_update, dash.no_update, dash.no_update, {"display": "none"}, dash.no_update

        for r in pendiente:
            r[C_NOTA1]    = max(0.0, min(5.0, float(r[C_NOTA1]) if r[C_NOTA1] is not None else 0.0))
            r[C_NOTA2]    = max(0.0, min(5.0, float(r[C_NOTA2]) if r[C_NOTA2] is not None else 0.0))
            r[C_NOTA3]    = max(0.0, min(5.0, float(r[C_NOTA3]) if r[C_NOTA3] is not None else 0.0))
            promedio      = round((r[C_NOTA1] + r[C_NOTA2] + r[C_NOTA3]) / 3, 2)
            r[C_PROMEDIO] = promedio
            r[C_DESEMPEN] = calc_desempenio(promedio)

        return pendiente, None, None, {"display": "none"}, BTN_VISIBLE


    # ── Modal: rellenar notas vacías con 0 y ajustar fuera de rango ───────────
    @appnotas.callback(
        Output("store_excel",                "data",  allow_duplicate=True),
        Output("store_excel_pendiente",      "data",  allow_duplicate=True),
        Output("store_excel_pendiente_vacios","data", allow_duplicate=True),
        Output("modal_notas",                "style", allow_duplicate=True),
        Output("btn_confirmar_carga",        "style", allow_duplicate=True),
        Input("modal_btn_fill",              "n_clicks"),
        State("store_excel_pendiente",       "data"),
        prevent_initial_call=True,
    )
    def modal_rellenar(n_clicks, pendiente):
        BTN_VISIBLE = {
            "display": "inline-block", "background": ACCENT, "color": "#0c0e10",
            "border": "none", "fontFamily": "'DM Mono', monospace",
            "fontSize": "11px", "fontWeight": "500", "letterSpacing": ".2em",
            "textTransform": "uppercase", "padding": "14px 32px", "cursor": "pointer",
        }
        if not pendiente:
            return dash.no_update, dash.no_update, dash.no_update, {"display": "none"}, dash.no_update

        for r in pendiente:
            # Rellenar vacíos con 0, luego ajustar al rango
            r[C_NOTA1]    = max(0.0, min(5.0, float(r[C_NOTA1]) if r[C_NOTA1] is not None else 0.0))
            r[C_NOTA2]    = max(0.0, min(5.0, float(r[C_NOTA2]) if r[C_NOTA2] is not None else 0.0))
            r[C_NOTA3]    = max(0.0, min(5.0, float(r[C_NOTA3]) if r[C_NOTA3] is not None else 0.0))
            promedio      = round((r[C_NOTA1] + r[C_NOTA2] + r[C_NOTA3]) / 3, 2)
            r[C_PROMEDIO] = promedio
            r[C_DESEMPEN] = calc_desempenio(promedio)

        return pendiente, None, None, {"display": "none"}, BTN_VISIBLE


    # ── Modal: eliminar estudiantes con notas inválidas o vacías ─────────────
    @appnotas.callback(
        Output("store_excel",                "data",  allow_duplicate=True),
        Output("store_excel_pendiente",      "data",  allow_duplicate=True),
        Output("store_excel_pendiente_vacios","data", allow_duplicate=True),
        Output("modal_notas",                "style", allow_duplicate=True),
        Output("btn_confirmar_carga",        "style", allow_duplicate=True),
        Output("panel_carga_msg",            "children", allow_duplicate=True),
        Input("modal_btn_eliminar",          "n_clicks"),
        State("store_excel_pendiente_vacios","data"),
        prevent_initial_call=True,
    )
    def modal_eliminar(n_clicks, pendiente):
        BTN_VISIBLE = {
            "display": "inline-block", "background": ACCENT, "color": "#0c0e10",
            "border": "none", "fontFamily": "'DM Mono', monospace",
            "fontSize": "11px", "fontWeight": "500", "letterSpacing": ".2em",
            "textTransform": "uppercase", "padding": "14px 32px", "cursor": "pointer",
        }
        BTN_OCULTO = {"display": "none"}

        if not pendiente:
            return (dash.no_update, dash.no_update, dash.no_update,
                    {"display": "none"}, dash.no_update, dash.no_update)

        # Filtrar solo los que tienen notas 100% válidas (no None, no fuera de rango)
        limpios   = []
        eliminados = []
        for r in pendiente:
            n1 = r.get(C_NOTA1)
            n2 = r.get(C_NOTA2)
            n3 = r.get(C_NOTA3)
            invalido = (
                n1 is None or n2 is None or n3 is None or
                float(n1) < 0 or float(n1) > 5 or
                float(n2) < 0 or float(n2) > 5 or
                float(n3) < 0 or float(n3) > 5
            )
            if invalido:
                eliminados.append(r[C_NOMBRE])
            else:
                limpios.append(r)

        if not limpios:
            msg_vacio = html.Div([
                html.Span("⚠  SIN REGISTROS VÁLIDOS", style={
                    "fontSize": "11px", "letterSpacing": ".2em",
                    "color": ORANGE, "fontWeight": "500",
                }),
                html.P(
                    f"Todos los estudiantes tenían notas inválidas o vacías "
                    f"({len(eliminados)} eliminados). No hay nada que guardar.",
                    style={"color": MUTED, "fontSize": "12px", "marginTop": "8px"}
                ),
            ], style={"background": "rgba(251,146,60,.08)", "border": f"1px solid {ORANGE}",
                      "borderLeft": f"4px solid {ORANGE}", "padding": "16px 20px"})
            return None, None, None, {"display": "none"}, BTN_OCULTO, msg_vacio

        # Recalcular promedio para los limpios (por si acaso)
        for r in limpios:
            promedio      = round((float(r[C_NOTA1]) + float(r[C_NOTA2]) + float(r[C_NOTA3])) / 3, 2)
            r[C_PROMEDIO] = promedio
            r[C_DESEMPEN] = calc_desempenio(promedio)

        resumen_elim = html.Div([
            html.Div(html.Span("✓  LOTE LIMPIADO", style={
                "fontSize": "11px", "letterSpacing": ".2em", "color": GREEN, "fontWeight": "500",
            }), style={"marginBottom": "10px"}),
            html.P(
                f"{len(limpios)} estudiante(s) válidos listos para guardar. "
                f"{len(eliminados)} eliminado(s) por notas inválidas o vacías.",
                style={"color": TEXT, "fontSize": "13px", "marginBottom": "6px"},
            ),
            html.P(
                f"Eliminados: {', '.join(eliminados)}",
                style={"color": MUTED, "fontSize": "11px"}
            ) if eliminados else None,
        ], style={"background": "rgba(74,222,128,.08)", "border": f"1px solid {GREEN}",
                  "borderLeft": f"4px solid {GREEN}", "padding": "16px 20px"})

        return limpios, None, None, {"display": "none"}, BTN_VISIBLE, resumen_elim


    # ── Modal: usuario cancela ────────────────────────────────────────────────
    @appnotas.callback(
        Output("modal_notas",                "style",    allow_duplicate=True),
        Output("store_excel_pendiente",      "data",     allow_duplicate=True),
        Output("store_excel_pendiente_vacios","data",    allow_duplicate=True),
        Output("panel_carga_msg",            "children", allow_duplicate=True),
        Output("upload_preview",             "children", allow_duplicate=True),
        Output("btn_confirmar_carga",        "style",    allow_duplicate=True),
        Input("modal_btn_cancelar",          "n_clicks"),
        prevent_initial_call=True,
    )
    def modal_cancelar(n_clicks):
        msg = html.Div([
            html.Span("↩  CARGA CANCELADA", style={
                "fontSize": "11px", "letterSpacing": ".2em", "color": MUTED, "fontWeight": "500",
            }),
            html.P("Corrige las notas en el archivo e inténtalo de nuevo.",
                   style={"color": MUTED, "fontSize": "12px", "marginTop": "8px"}),
        ], style={"background": SURFACE2, "border": f"1px solid {BORDER}",
                  "borderLeft": f"4px solid {MUTED}", "padding": "16px 20px"})
        return {"display": "none"}, None, None, msg, "", {"display": "none"}


    # ── Confirmar carga masiva a MySQL ────────────────────────────────────────
    @appnotas.callback(
        Output("panel_carga_msg",          "children", allow_duplicate=True),
        Output("upload_preview",           "children", allow_duplicate=True),
        Output("btn_confirmar_carga",      "style",    allow_duplicate=True),
        Output("store_excel",              "data",     allow_duplicate=True),
        Output("store_carga_completada",   "data",     allow_duplicate=True),
        Output("panel_estadisticas_carga", "children", allow_duplicate=True),
        Output("panel_rechazados",         "children", allow_duplicate=True),
        Input("btn_confirmar_carga",       "n_clicks"),
        State("store_excel",               "data"),
        State("store_carga_completada",    "data"),
        prevent_initial_call=True,
    )
    def confirmar_carga(n_clicks, datos, trigger_actual):
        import base64, io as _io
        BTN_OCULTO = {"display": "none"}

        if not datos:
            return (
                html.Div([
                    html.P("Sin datos para guardar.", style={"color": RED, "fontSize": "13px"}),
                    html.P("Sube un archivo primero.", style={"color": MUTED, "fontSize": "11px"}),
                ], style={"background": "rgba(224,92,92,.08)", "border": f"1px solid {RED}",
                          "borderLeft": f"4px solid {RED}", "padding": "16px 20px"}),
                "", BTN_OCULTO, None, dash.no_update, None, None,
            )

        # ── Separar los que tienen notas válidas de los rechazados ────────────
        rechazados_local = []
        validos = []
        for r in datos:
            n1 = r.get(C_NOTA1)
            n2 = r.get(C_NOTA2)
            n3 = r.get(C_NOTA3)
            edad = r.get(C_EDAD)
            nombre = r.get(C_NOMBRE, "")
            carrera = r.get(C_CARRERA, "")

            motivos = []
            if any(v is None for v in [n1, n2, n3, edad, nombre, carrera]):
                motivos.append("Datos faltantes")
            if nombre in (None, "", "nan"):
                motivos.append("Nombre vacío")
            if edad is not None and int(edad) < 0:
                motivos.append("Edad negativa")
            if n1 is not None and (float(n1) < 0 or float(n1) > 5):
                motivos.append(f"nota1={n1} inválida")
            if n2 is not None and (float(n2) < 0 or float(n2) > 5):
                motivos.append(f"nota2={n2} inválida")
            if n3 is not None and (float(n3) < 0 or float(n3) > 5):
                motivos.append(f"nota3={n3} inválida")

            if motivos:
                r_copy = dict(r)
                r_copy["motivo_rechazo"] = ", ".join(motivos)
                rechazados_local.append(r_copy)
            else:
                validos.append(r)

        try:
            filas = [
                (r[C_NOMBRE], int(r[C_EDAD]), r[C_CARRERA],
                 float(r[C_NOTA1]), float(r[C_NOTA2]), float(r[C_NOTA3]),
                 float(r[C_PROMEDIO]), r[C_DESEMPEN])
                for r in validos
            ]

            resultado = insertar_masivo(filas)
            n_ins  = resultado["insertados"]
            dups   = resultado["duplicados"]
            n_dups = len(dups)

            # Agregar duplicados de último momento a rechazados
            for nombre_dup in dups:
                rechazados_local.append({
                    C_NOMBRE: nombre_dup, "motivo_rechazo": "Duplicado (ya existe en BD)"
                })

            n_rechazados = len(rechazados_local)
            n_total = n_ins + n_rechazados

            # ── Punto 4: Panel de estadísticas ───────────────────────────────
            panel_stats = html.Div([
                html.P("ESTADÍSTICAS DEL CARGUE MASIVO", style={
                    "fontSize": "10px", "letterSpacing": ".2em",
                    "color": ACCENT, "marginBottom": "16px",
                    "textTransform": "uppercase",
                }),
                html.Div([
                    html.Div([
                        html.P("TOTAL PROCESADOS", style={
                            "fontSize": "10px", "letterSpacing": ".14em",
                            "color": MUTED, "marginBottom": "6px",
                        }),
                        html.H3(str(n_total), style={
                            "fontFamily": "'DM Serif Display', serif",
                            "fontSize": "36px", "color": TEXT,
                            "fontWeight": "400", "margin": "0",
                        }),
                    ], style={"flex": "1", "padding": "16px",
                              "background": SURFACE2, "border": f"1px solid {BORDER}",
                              "textAlign": "center"}),
                    html.Div([
                        html.P("INSERTADOS", style={
                            "fontSize": "10px", "letterSpacing": ".14em",
                            "color": MUTED, "marginBottom": "6px",
                        }),
                        html.H3(str(n_ins), style={
                            "fontFamily": "'DM Serif Display', serif",
                            "fontSize": "36px", "color": GREEN,
                            "fontWeight": "400", "margin": "0",
                        }),
                    ], style={"flex": "1", "padding": "16px",
                              "background": SURFACE2, "border": f"1px solid {GREEN}",
                              "textAlign": "center"}),
                    html.Div([
                        html.P("RECHAZADOS", style={
                            "fontSize": "10px", "letterSpacing": ".14em",
                            "color": MUTED, "marginBottom": "6px",
                        }),
                        html.H3(str(n_rechazados), style={
                            "fontFamily": "'DM Serif Display', serif",
                            "fontSize": "36px", "color": RED,
                            "fontWeight": "400", "margin": "0",
                        }),
                    ], style={"flex": "1", "padding": "16px",
                              "background": SURFACE2, "border": f"1px solid {RED}",
                              "textAlign": "center"}),
                    html.Div([
                        html.P("DUPLICADOS", style={
                            "fontSize": "10px", "letterSpacing": ".14em",
                            "color": MUTED, "marginBottom": "6px",
                        }),
                        html.H3(str(n_dups), style={
                            "fontFamily": "'DM Serif Display', serif",
                            "fontSize": "36px", "color": ORANGE,
                            "fontWeight": "400", "margin": "0",
                        }),
                    ], style={"flex": "1", "padding": "16px",
                              "background": SURFACE2, "border": f"1px solid {ORANGE}",
                              "textAlign": "center"}),
                ], style={"display": "flex", "gap": "12px", "flexWrap": "wrap"}),
            ], style={**CARD, "borderLeft": f"3px solid {ACCENT}", "marginTop": "16px"})

            # ── Punto 3: Generar Excel con rechazados ─────────────────────────
            panel_rechazados = None
            if rechazados_local:
                import pandas as pd_rej
                import openpyxl
                df_rej = pd_rej.DataFrame(rechazados_local)
                # Asegurar columna motivo_rechazo al final
                cols_rej = [c for c in df_rej.columns if c != "motivo_rechazo"] + ["motivo_rechazo"]
                df_rej = df_rej[[c for c in cols_rej if c in df_rej.columns]]

                buf = _io.BytesIO()
                df_rej.to_excel(buf, index=False)
                buf.seek(0)
                b64 = base64.b64encode(buf.read()).decode()
                href = f"data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}"

                panel_rechazados = html.Div([
                    html.P("REGISTROS RECHAZADOS", style={
                        "fontSize": "10px", "letterSpacing": ".2em",
                        "color": RED, "marginBottom": "12px",
                        "textTransform": "uppercase",
                    }),
                    html.P(
                        f"{n_rechazados} registro(s) no se insertaron. "
                        "Descarga el archivo para revisar los motivos.",
                        style={"color": TEXT, "fontSize": "12px", "marginBottom": "14px"}
                    ),
                    dash_table.DataTable(
                        data=df_rej.head(10).to_dict("records"),
                        columns=[{"name": c.title(), "id": c} for c in df_rej.columns],
                        page_size=5,
                        style_table={"overflowX": "auto"},
                        style_cell={"textAlign": "center", "padding": "8px 12px",
                                    "fontFamily": "'DM Mono', monospace", "fontSize": "11px",
                                    "backgroundColor": SURFACE, "color": TEXT,
                                    "border": f"1px solid {BORDER}"},
                        style_header={"backgroundColor": SURFACE2, "color": MUTED,
                                      "fontWeight": "400", "fontSize": "10px",
                                      "border": f"1px solid {BORDER}"},
                        style_data_conditional=[
                            {"if": {"column_id": "motivo_rechazo"},
                             "color": ORANGE, "fontWeight": "500"},
                        ],
                    ),
                    html.A(
                        "⬇  DESCARGAR RECHAZADOS (.xlsx)",
                        href=href,
                        download="rechazados.xlsx",
                        style={
                            "display": "inline-block", "marginTop": "14px",
                            "background": "transparent", "color": RED,
                            "border": f"1px solid {RED}", "padding": "10px 22px",
                            "fontFamily": "'DM Mono', monospace", "fontSize": "10px",
                            "letterSpacing": ".18em", "textTransform": "uppercase",
                            "textDecoration": "none", "transition": "background .2s",
                        }
                    ),
                ], style={**CARD, "borderLeft": f"3px solid {RED}", "marginTop": "16px"})

            # Punto 2: incrementar trigger para refrescar dashboard
            nuevo_trigger = (trigger_actual or 0) + 1

            avisos_dup = []
            if dups:
                avisos_dup.append(html.P(
                    f"ℹ {n_dups} omitido(s) por duplicado: {', '.join(dups)}.",
                    style={"color": ORANGE, "fontSize": "11px", "marginTop": "8px"}
                ))

            return (
                html.Div([
                    html.Div(html.Span("✓  CARGA EXITOSA", style={
                        "fontSize": "11px", "letterSpacing": ".2em",
                        "color": GREEN, "fontWeight": "500",
                    }), style={"marginBottom": "10px"}),
                    html.P(
                        f"{n_ins} estudiante{'s' if n_ins != 1 else ''} "
                        f"guardado{'s' if n_ins != 1 else ''} en la base de datos.",
                        style={"color": TEXT, "fontSize": "13px", "marginBottom": "4px"},
                    ),
                    html.P("Los gráficos y la tabla se han actualizado automáticamente.",
                           style={"color": MUTED, "fontSize": "11px"}),
                    *avisos_dup,
                ], style={"background": "rgba(74,222,128,.08)", "border": f"1px solid {GREEN}",
                          "borderLeft": f"4px solid {GREEN}", "padding": "18px 20px"}),
                "", BTN_OCULTO, None, nuevo_trigger, panel_stats, panel_rechazados,
            )

        except Exception as e:
            return (
                html.Div([
                    html.P("✕  ERROR AL GUARDAR", style={"color": RED, "fontSize": "11px",
                           "letterSpacing": ".2em", "marginBottom": "10px"}),
                    html.P("No se pudo guardar en la base de datos.",
                           style={"color": TEXT, "fontSize": "13px"}),
                    html.P(f"Detalle: {e}",
                           style={"color": MUTED, "fontSize": "11px", "marginTop": "6px"}),
                    html.Div([html.Span("💡 "),
                              html.Span("Verifica que MySQL esté activo.",
                                        style={"color": ACCENT, "fontSize": "11px"})],
                             style={"marginTop": "10px", "padding": "10px",
                                    "background": "rgba(200,169,110,.07)"}),
                ], style={"background": "rgba(224,92,92,.08)", "border": f"1px solid {RED}",
                          "borderLeft": f"4px solid {RED}", "padding": "18px 20px"}),
                "", BTN_OCULTO, None, dash.no_update, None, None,
            )

    # ── Punto 5: Ranking top 10 ───────────────────────────────────────────────
    @appnotas.callback(
        Output("tabla_ranking",   "children"),
        Output("grafico_ranking", "figure"),
        Input("store_carga_completada", "data"),
        Input("filtro_carrera",  "value"),   # cualquier cambio refresca
    )
    def actualizar_ranking(_trigger, _carrera):
        df_all = cargar_datos()
        if df_all.empty:
            return html.P("Sin datos.", style={"color": MUTED}), go.Figure()

        top10 = (df_all[[C_NOMBRE, C_CARRERA, C_PROMEDIO]]
                 .sort_values(C_PROMEDIO, ascending=False)
                 .head(10)
                 .reset_index(drop=True))
        top10.index += 1  # ranking 1-based

        tabla = dash_table.DataTable(
            data=top10.reset_index().rename(columns={"index": "Pos."}).to_dict("records"),
            columns=[
                {"name": "Pos.",     "id": "Pos."},
                {"name": "Nombre",   "id": C_NOMBRE},
                {"name": "Carrera",  "id": C_CARRERA},
                {"name": "Promedio", "id": C_PROMEDIO},
            ],
            style_table={"overflowX": "auto", "marginBottom": "20px"},
            style_cell={"textAlign": "center", "padding": "10px 16px",
                        "fontFamily": "'DM Mono', monospace", "fontSize": "12px",
                        "backgroundColor": SURFACE, "color": TEXT,
                        "border": f"1px solid {BORDER}"},
            style_header={"backgroundColor": SURFACE2, "color": MUTED,
                          "fontWeight": "400", "fontSize": "10px",
                          "letterSpacing": ".12em", "textTransform": "uppercase",
                          "border": f"1px solid {BORDER}"},
            style_data_conditional=[
                {"if": {"row_index": 0}, "color": ACCENT, "fontWeight": "600",
                 "borderLeft": f"3px solid {ACCENT}"},
                {"if": {"row_index": 1}, "color": TEXT, "fontWeight": "500"},
                {"if": {"row_index": 2}, "color": TEXT, "fontWeight": "500"},
                {"if": {"filter_query": f"{{{C_PROMEDIO}}} >= 4",
                        "column_id": C_PROMEDIO}, "color": GREEN},
                {"if": {"filter_query": f"{{{C_PROMEDIO}}} < 3",
                        "column_id": C_PROMEDIO}, "color": RED},
            ],
        )

        colores_rank = [ACCENT if i == 0 else (BLUE if i == 1 else (GREEN if i == 2 else MUTED))
                        for i in range(len(top10))]
        fig = px.bar(
            top10.reset_index(),
            x=C_NOMBRE, y=C_PROMEDIO,
            color=C_CARRERA,
            text=C_PROMEDIO,
            labels={C_NOMBRE: "Estudiante", C_PROMEDIO: "Promedio", C_CARRERA: "Carrera"},
            color_discrete_sequence=[ACCENT, BLUE, GREEN, RED, "#a78bfa", ORANGE],
        )
        fig.update_traces(textposition="outside", textfont=dict(color=TEXT, size=12),
                          marker_line_color=BG, marker_line_width=1)
        apply_template(fig, "Top 10 — Mejores Promedios")
        fig.update_yaxes(range=[0, 5.5])

        return tabla, fig

    # ── Punto 6: Alertas estudiantes en riesgo ────────────────────────────────
    @appnotas.callback(
        Output("alerta_riesgo", "children"),
        Input("store_carga_completada", "data"),
        Input("filtro_carrera", "value"),
    )
    def actualizar_alertas(_trigger, _carrera):
        df_all = cargar_datos()
        if df_all.empty:
            return html.P("Sin datos.", style={"color": MUTED})

        en_riesgo = df_all[df_all[C_PROMEDIO] < 3.0][[C_NOMBRE, C_CARRERA, C_PROMEDIO]].copy()
        en_riesgo = en_riesgo.sort_values(C_PROMEDIO).reset_index(drop=True)

        if en_riesgo.empty:
            return html.Div([
                html.Span("✓  SIN ESTUDIANTES EN RIESGO", style={
                    "fontSize": "11px", "letterSpacing": ".2em",
                    "color": GREEN, "fontWeight": "500",
                }),
                html.P("Todos los estudiantes tienen promedio ≥ 3.0.",
                       style={"color": MUTED, "fontSize": "12px", "marginTop": "8px"}),
            ], style={"background": "rgba(74,222,128,.07)", "border": f"1px solid {GREEN}",
                      "borderLeft": f"4px solid {GREEN}", "padding": "16px 20px"})

        n_riesgo = len(en_riesgo)

        alerta_header = html.Div([
            html.Div([
                html.Span("⚠", style={"fontSize": "20px", "color": RED, "marginRight": "10px",
                                       "verticalAlign": "middle"}),
                html.Span(f"{n_riesgo} ESTUDIANTE{'S' if n_riesgo != 1 else ''} EN RIESGO",
                          style={"fontSize": "11px", "letterSpacing": ".2em",
                                 "color": RED, "fontWeight": "500", "verticalAlign": "middle"}),
            ], style={"marginBottom": "16px"}),
            html.P("Promedio menor a 3.0 — se recomienda intervención académica urgente.",
                   style={"color": MUTED, "fontSize": "12px", "marginBottom": "16px"}),
        ])

        tabla_riesgo = dash_table.DataTable(
            data=en_riesgo.to_dict("records"),
            columns=[
                {"name": "Nombre",   "id": C_NOMBRE},
                {"name": "Carrera",  "id": C_CARRERA},
                {"name": "Promedio", "id": C_PROMEDIO},
            ],
            page_size=10,
            sort_action="native",
            style_table={"overflowX": "auto"},
            style_cell={"textAlign": "center", "padding": "10px 16px",
                        "fontFamily": "'DM Mono', monospace", "fontSize": "12px",
                        "backgroundColor": SURFACE, "color": TEXT,
                        "border": f"1px solid {BORDER}"},
            style_header={"backgroundColor": SURFACE2, "color": MUTED,
                          "fontWeight": "400", "fontSize": "10px",
                          "letterSpacing": ".12em", "textTransform": "uppercase",
                          "border": f"1px solid {BORDER}"},
            style_data_conditional=[
                {"if": {"filter_query": f"{{{C_PROMEDIO}}} < 2",
                        "column_id": C_PROMEDIO},
                 "color": RED, "fontWeight": "600"},
                {"if": {"filter_query": f"{{{C_PROMEDIO}}} >= 2",
                        "column_id": C_PROMEDIO},
                 "color": ORANGE, "fontWeight": "500"},
            ],
        )

        return html.Div([
            alerta_header,
            tabla_riesgo,
        ], style={"background": "rgba(224,92,92,.06)", "border": f"1px solid {RED}",
                  "borderLeft": f"4px solid {RED}", "padding": "20px"})

    return appnotas