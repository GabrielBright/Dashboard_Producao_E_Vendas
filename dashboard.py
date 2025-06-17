import pandas as pd
import plotly.graph_objects as go
import dash
from dash import dcc, html
from dash.dependencies import Input, Output
import dash_bootstrap_components as dbc

CORES = {
    "2023": "#D3D3D3",
    "2024": "#4A5E7D",
    "2025": "#F4A261",
    "hover": "#E76F51"
}

app = dash.Dash(__name__, external_stylesheets=[dbc.themes.FLATLY], suppress_callback_exceptions=True)
app.title = "Dashboard - Produção & Vendas LATAM"
server = app.server

DADOS_VENDAS = {}
DADOS_PRODUCAO = {}

def corrigir_sequencia(sequencia):
    if max(sequencia) > 1_000_000:
        return [x / 100_000 for x in sequencia]
    return sequencia


def carregar_vendas(caminho_arquivo, ano_coluna):
    try:
        df_raw = pd.read_excel(caminho_arquivo, sheet_name='I. Emplacamento', header=None)

        total_idx = df_raw.index[df_raw.apply(lambda row: row.astype(str).str.lower().str.contains("emplacamento total de autoveículos", case=False).any(), axis=1)].tolist()
        if not total_idx:
            raise ValueError("Tabela 'Emplacamento Total de Autoveículos' não encontrada!")

        header = total_idx[0] + 3
        df = pd.read_excel(caminho_arquivo, sheet_name='I. Emplacamento', header=[header, header+1])
        df.columns = [str(col[1]) if str(col[1]) != "nan" else str(col[0]) for col in df.columns]

        col_pesados = None
        col_leves = None
        for col in df.columns:
            if df[col].astype(str).str.lower().str.contains("caminhões|ônibus").any():
                col_pesados = col
            if df[col].astype(str).str.lower().str.contains("automóveis|comerciais leves").any():
                col_leves = col

        if not col_pesados:
            col_pesados = [c for c in df.columns if "Unnamed" in c][0]
        if not col_leves:
            col_leves = [c for c in df.columns if "Unnamed" in c][-1]

        meses_padrao = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']
        for mes in meses_padrao:
            try:
                if mes in df.columns:
                    df[mes] = pd.to_numeric(df[mes], errors='coerce').fillna(0)
                else:
                    df[mes] = 0
            except Exception as e:
                df[mes] = 0
                print(f"Erro não foi possivel processar mês {e}")

        automoveis = df[df[col_leves].str.strip() == "Automóveis"]
        comerciais = df[df[col_leves].str.strip() == "Comerciais leves"]
        
        soma_leves = []
        
        for mes in meses_padrao:
            soma = 0
            if mes in automoveis.columns:
                soma += automoveis[mes].sum()
            if mes in comerciais.columns:
                soma += comerciais[mes].sum()
            soma_leves.append(soma)
            
        soma_leves = corrigir_sequencia(soma_leves)
        
        pesados = df[df[col_pesados].isin(["Caminhões", "Ônibus"])]
        
        soma_pesados = [pesados[mes].sum() if mes in pesados.columns else 0 for mes in meses_padrao]
        soma_pesados = corrigir_sequencia(soma_pesados)
        
        total_leves = sum(soma_leves)
        total_pesados = sum(soma_pesados)

        return meses_padrao, soma_leves, soma_pesados, total_leves, total_pesados
    except Exception as e:
        raise Exception(f"Erro ao carregar vendas: {str(e)}")

def carregar_producao(caminho_arquivo, ano_coluna):
    try:
        df_raw = pd.read_excel(caminho_arquivo, sheet_name='VI. Produção', header=None)
        
        header_row = None
        for i, row in enumerate(df_raw.values):
            if "Unidades" in row:
                header_row = i
                break

        if header_row is None:
            raise ValueError("'Unidades' não encontrada na aba 'VI. Produção'!")

        df = pd.read_excel(caminho_arquivo, sheet_name='VI. Produção', header=header_row)
        df.columns = [str(col).replace("\n", "").replace("  ", " ").strip() for col in df.columns]

        meses_cols = [ano_coluna, 'Unnamed: 4', 'Unnamed: 5', 'Unnamed: 6', 'Unnamed: 7',
                      'Unnamed: 8', 'Unnamed: 9', 'Unnamed: 10', 'Unnamed: 11', 'Unnamed: 12',
                      'Unnamed: 13', 'Unnamed: 14']

        nomes_meses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']
        renomear = dict(zip(meses_cols, nomes_meses))
        df.rename(columns=renomear, inplace=True)

        if ano_coluna == "2025":
            df.rename(columns={'2025': 'Jan'}, inplace=True)

        for mes in nomes_meses:
            try:
                if mes in df.columns:
                    df[mes] = pd.to_numeric(df[mes], errors='coerce').fillna(0)
                else:
                    df[mes] = 0
            except Exception as e:
                df[mes] = 0
                print(f"Não foi possivel processar o mês: ")
                
        leves = df[df['Unidades'].isin(['Automóveis', 'Comerciais leves'])]
        pesados = df[df['Unidades'].isin(['Semileves', 'Leves', 'Médios', 'Semipesados', 'Pesados', 'Rodoviário', 'Urbano'])]

        soma_leves = leves[nomes_meses].sum().values.tolist()
        soma_leves = corrigir_sequencia(soma_leves)
        
        soma_pesados = pesados[nomes_meses].sum().values.tolist()
        soma_pesados = corrigir_sequencia(soma_pesados)
        
        total_leves = sum(soma_leves)
        total_pesados = sum(soma_pesados)

        return nomes_meses, soma_leves, soma_pesados, total_leves, total_pesados
    except Exception as e:
        raise Exception(f"Erro ao carregar produção: {str(e)}")

arquivos = {
    "2023": ["siteautoveiculos2023.xlsx", "2023"],
    "2024": ["siteautoveiculos2024.xlsx", "2024"],
    "2025": ["siteautoveiculos2025.xlsx", "2025"]
}

def inicializar_dados():
    global DADOS_VENDAS, DADOS_PRODUCAO
    for ano, (arquivo, ano_col) in arquivos.items():
        try:
            meses_v, leves_v, pesados_v, total_leves_v, total_pesados_v = carregar_vendas(arquivo, ano_col)
            DADOS_VENDAS[ano] = {"meses": meses_v, "leves": leves_v, "pesados": pesados_v}
        except Exception as e:
            print(f"Erro ao carregar vendas para {ano}: {e}")
            DADOS_VENDAS[ano] = {
                "meses": ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'],
                "leves": [0] * 12,
                "pesados": [0] * 12
            }

        try:
            meses_p, leves_p, pesados_p, total_leves_p, total_pesados_p = carregar_producao(arquivo, ano_col)
            DADOS_PRODUCAO[ano] = {"meses": meses_p, "leves": leves_p, "pesados": pesados_p}
        except Exception as e:
            print(f"Erro ao carregar produção para {ano}: {e}")
            DADOS_PRODUCAO[ano] = {
                "meses": ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'],
                "leves": [0] * 12,
                "pesados": [0] * 12
            }

inicializar_dados()

def carregar_dados(tipo):
    return DADOS_PRODUCAO if tipo == "Produção" else DADOS_VENDAS

def calcular_variacao(atual, anterior):
    if anterior == 0:
        return html.Span("0%", style={"color": "gray"})
    variacao = ((atual - anterior) / anterior) * 100
    cor = "green" if variacao > 0 else "red"
    simbolo = "▲" if variacao > 0 else "▼"
    return html.Span(f"{simbolo} {abs(variacao):.1f}%", style={"color": cor, "fontWeight": "bold"})

def criar_card(titulo, valor, variacao, cor):
    return dbc.Col(dbc.Card([
        dbc.CardBody([
            html.H6(titulo, className="text-center", style={"color": cor, "fontSize": "16px", "fontWeight": "600"}),
            html.H3(f"{valor:,.0f}".replace(",", "."), className="text-center fw-bold", style={"color": cor, "fontSize": "24px"}),
            html.Div(variacao, className="text-center", style={"fontSize": "14px"})
        ])
    ], className="shadow-sm rounded-3 hover-effect", style={"border": f"2px solid {cor}", "transition": "transform 0.2s", "margin": "10px"}),
    width=4)

app.layout = dbc.Container([
    html.Div([
        html.H1("Dashboard - Produção & Vendas LATAM", className="text-center", style={"color": "#4A5E7D", "marginBottom": "20px"}),
        html.Hr(style={"borderTop": "3px solid #F4A261", "width": "250px", "margin": "0 auto 20px"}),
        dcc.RadioItems(id="tipo-radio", options=["Produção", "Vendas"], value="Vendas",
                       inline=True, inputStyle={"margin-right": "8px", "margin-left": "15px"},
                       className="text-center", style={"color": "#4A5E7D", "marginBottom": "20px"})
    ], className="shadow rounded p-4 mb-4", style={"background": "linear-gradient(to right, #F9F9F9, #FFFFFF)"}),

    html.Div(id="error-message", style={"color": "red", "textAlign": "center", "marginBottom": "20px"}),

    dbc.Row(id="kpis", className="mb-4"),

    dbc.Row([
        dbc.Col([
            html.Div([
                html.H5("Veículos Leves – Mês a Mês", className="text-center", style={"color": "#4A5E7D", "marginBottom": "10px"}),
                dcc.Loading(dcc.Graph(id="grafico-linha-leves", style={"height": "400px"})),
                dcc.RangeSlider(id="slider-leves", min=0, max=11, value=[0, 11],
                                marks={i: m for i, m in enumerate(['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'])},
                                step=1),
                html.Div(id="slider-output-leves", style={"textAlign": "center", "color": "#4A5E7D", "marginTop": "10px"})
            ], className="shadow rounded p-3 bg-white", style={"height": "100%"})
        ], width=6, lg=6, md=12),
        
        dbc.Col([
            html.Div([
                html.H5("Veículos Leves – Total Anual", className="text-center", style={"color": "#4A5E7D", "marginBottom": "10px"}),
                dcc.Loading(dcc.Graph(id="grafico-barra-leves", style={"height": "400px"}))
            ], className="shadow rounded p-3 bg-white", style={"height": "100%"})
        ], width=6, lg=6, md=12)
    ], className="mb-4", style={"display": "flex", "alignItems": "stretch"}),

    dbc.Row([
        dbc.Col([
            html.Div([
                html.H5("Veículos Pesados – Mês a Mês", className="text-center", style={"color": "#4A5E7D", "marginBottom": "10px"}),
                dcc.Loading(dcc.Graph(id="grafico-linha-pesados", style={"height": "100%"})),
                dcc.RangeSlider(id="slider-pesados", min=0, max=11, value=[0, 11],
                                marks={i: m for i, m in enumerate(['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'])},
                                step=1),
                html.Div(id="slider-output-pesados", style={"textAlign": "center", "color": "#4A5E7D", "marginTop": "10px"})
            ], className="shadow rounded p-3 bg-white", style={"height": "100%"})
        ], width=6, lg=6, md=12),
        
        dbc.Col([
            html.Div([
                html.H5("Veículos Pesados – Total Anual", className="text-center", style={"color": "#4A5E7D", "marginBottom": "10px"}),
                dcc.Loading(dcc.Graph(id="grafico-barra-pesados", style={"height": "100%"}))
            ], className="shadow rounded p-3 bg-white", style={"height": "100%"})
        ], width=6, lg=6, md=12)
    ], className="mb-4", style={"display": "flex", "alignItems": "stretch"})
], fluid=True)

@app.callback(
    Output("kpis", "children"),
    Output("grafico-linha-leves", "figure"),
    Output("grafico-barra-leves", "figure"),
    Output("grafico-linha-pesados", "figure"),
    Output("grafico-barra-pesados", "figure"),
    Output("slider-output-leves", "children"),
    Output("slider-output-pesados", "children"),
    Output("error-message", "children"),
    Input("tipo-radio", "value"),
    Input("slider-leves", "value"),
    Input("slider-pesados", "value")
)
def atualizar(tipo, r_leves, r_pesados):
    try:
        dados = carregar_dados(tipo)
        error_message = ""

        fig_linha_leves = go.Figure()
        fig_linha_pesados = go.Figure()
        fig_barra_leves = go.Figure()
        fig_barra_pesados = go.Figure()

        total_25_leves = sum(dados["2025"]["leves"][:3])
        total_24_leves = sum(dados["2024"]["leves"][:3])
        total_25_pesados = sum(dados["2025"]["pesados"][:3])
        total_24_pesados = sum(dados["2024"]["pesados"][:3])

        cards = [
            criar_card("Total Vendas" if tipo == "Vendas" else "Total Produção", 
                       total_25_leves + total_25_pesados, 
                       calcular_variacao(total_25_leves + total_25_pesados, total_24_leves + total_24_pesados), 
                       "#4A5E7D"),
            criar_card("Total Leves", total_25_leves, calcular_variacao(total_25_leves, total_24_leves), "#F4A261"),
            criar_card("Total Pesados", total_25_pesados, calcular_variacao(total_25_pesados, total_24_pesados), "#D3D3D3"),
        ]

        for ano in ["2023", "2024", "2025"]:
            m = dados[ano]["meses"]
            l = dados[ano]["leves"]
            p = dados[ano]["pesados"]

            last_index_leves = max([i for i, v in enumerate(l) if v > 0], default=-1)
            last_index_pesados = max([i for i, v in enumerate(p) if v > 0], default=-1)

            meses_filtrados_leves = m[r_leves[0]:min(r_leves[1]+1, last_index_leves+1)]
            valores_filtrados_leves = l[r_leves[0]:min(r_leves[1]+1, last_index_leves+1)]

            meses_filtrados_pesados = m[r_pesados[0]:min(r_pesados[1]+1, last_index_pesados+1)]
            valores_filtrados_pesados = p[r_pesados[0]:min(r_pesados[1]+1, last_index_pesados+1)]

            fig_linha_leves.add_trace(go.Scatter(
                x=meses_filtrados_leves,
                y=valores_filtrados_leves,
                mode='lines+markers',
                name=ano,
                line=dict(color=CORES[ano], width=3),
                marker=dict(size=8),
                hovertemplate="%{y} unidades<br>%{x}"
            ))
            fig_linha_pesados.add_trace(go.Scatter(
                x=meses_filtrados_pesados,
                y=valores_filtrados_pesados,
                mode='lines+markers',
                name=ano,
                line=dict(color=CORES[ano], width=3),
                marker=dict(size=8),
                hovertemplate="%{y} unidades<br>%{x}"
            ))

            fig_barra_leves.add_trace(go.Bar(
                x=[ano],
                y=[sum(l)],
                name=ano,
                marker_color=CORES[ano],
                hovertemplate="%{y} unidades"
            ))
            fig_barra_pesados.add_trace(go.Bar(
                x=[ano],
                y=[sum(p)],
                name=ano,
                marker_color=CORES[ano],
                hovertemplate="%{y} unidades"
            ))

        for fig in [fig_linha_leves, fig_linha_pesados, fig_barra_leves, fig_barra_pesados]:
            fig.update_layout(
                plot_bgcolor='white',
                paper_bgcolor='white',
                margin=dict(t=20, b=40),
                showlegend=True,
                xaxis=dict(gridcolor='#EDEDED'),
                yaxis=dict(gridcolor='#EDEDED'),
                hovermode='x unified'
            )

        slider_leves_text = f"Intervalo: {r_leves[0] + 1} - {r_leves[1] + 1} meses"
        slider_pesados_text = f"Intervalo: {r_pesados[0] + 1} - {r_pesados[1] + 1} meses"

    except Exception as e:
        print(f"Erro no callback: {e}")  # Log no console para depuração
        error_message = "Ocorreu um erro. Tente novamente mais tarde."
        cards = [
            criar_card("Total Vendas" if tipo == "Vendas" else "Total Produção", 0, html.Span("N/A", style={"color": "gray"}), "#4A5E7D"),
            criar_card("Total Leves", 0, html.Span("N/A", style={"color": "gray"}), "#F4A261"),
            criar_card("Total Pesados", 0, html.Span("N/A", style={"color": "gray"}), "#D3D3D3"),
        ]
        fig_linha_leves = go.Figure()
        fig_linha_pesados = go.Figure()
        fig_barra_leves = go.Figure()
        fig_barra_pesados = go.Figure()
        slider_leves_text = "Intervalo: N/A"
        slider_pesados_text = "Intervalo: N/A"

    return cards, fig_linha_leves, fig_barra_leves, fig_linha_pesados, fig_barra_pesados, slider_leves_text, slider_pesados_text, error_message

if __name__ == "__main__":
    import os
    try:
        port = int(os.environ.get("PORT", 8050))  
        app.run_server(host="0.0.0.0", port=port)  
    except Exception as e:
        print(f"Erro ao iniciar o servidor Dash: {e}")