import math
import re
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

try:
    import gspread
    from google.oauth2.service_account import Credentials
except Exception:
    gspread = None
    Credentials = None

st.set_page_config(page_title="Dashboard Cubagem de Rotas", layout="wide")

PASTA_APP = Path(__file__).parent
ARQUIVO_PADRAO = PASTA_APP / "Cubagem mensal.xlsx"
SERVICE_ACCOUNT = PASTA_APP / "service_account.json"
ABA_HISTORICO = "base_historica"
ABA_ESPECIAIS = "cargas_especiais"
COLUNAS_ESPECIAIS = ["ID de carga", "Tipo especial", "Observação"]
COLUNAS_HISTORICO = [
    "ID de carga", "Data Carregamento", "Plano de Transporte", "Transportador", "Destino",
    "Modal", "Linha de Produção", "Tipo de Pedido", "Tipo do veículo",
    "Cubagem", "Capacidade", "Sobra", "Valor de Carga", "Ocupação %"
]

# -----------------------------
# Funções gerais
# -----------------------------
def localizar_planilha_padrao():
    if ARQUIVO_PADRAO.exists():
        return ARQUIVO_PADRAO
    arquivos_excel = sorted([p for p in PASTA_APP.glob("*.xlsx") if not p.name.startswith("~$")])
    return arquivos_excel[0] if arquivos_excel else None


def limpar_nome_colunas(df):
    df = df.copy()
    df.columns = [re.sub(r"\s+", " ", str(c).strip()) for c in df.columns]
    return df



def normalizar_chave_coluna(texto):
    texto = re.sub(r"\s+", " ", str(texto).strip().lower())
    mapa = str.maketrans("áàâãéêíóôõúüç", "aaaaeeiooouuc")
    return texto.translate(mapa)

def encontrar_coluna(df, candidatos=None, contem_todos=None):
    candidatos = candidatos or []
    mapa = {normalizar_chave_coluna(c): c for c in df.columns}
    for cand in candidatos:
        chave = normalizar_chave_coluna(cand)
        if chave in mapa:
            return mapa[chave]
    if contem_todos:
        termos = [normalizar_chave_coluna(t) for t in contem_todos]
        for col in df.columns:
            chave = normalizar_chave_coluna(col)
            if all(t in chave for t in termos):
                return col
    return None

def preencher_coluna_canonica(df, nome_padrao, candidatos=None, contem_todos=None):
    col = encontrar_coluna(df, candidatos=candidatos, contem_todos=contem_todos)
    if col is None:
        df[nome_padrao] = pd.NA
    elif col != nome_padrao:
        if nome_padrao not in df.columns:
            df[nome_padrao] = df[col]
        else:
            df[nome_padrao] = df[nome_padrao].combine_first(df[col])
    return df

def primeiro_valido(serie):
    for valor in serie:
        if pd.isna(valor):
            continue
        if isinstance(valor, str) and not valor.strip():
            continue
        return valor
    return pd.NA

def br_numero(valor, casas=0, moeda=False, percentual=False):
    if pd.isna(valor):
        return ""
    if percentual:
        return f"{valor * 100:,.1f}%".replace(",", "X").replace(".", ",").replace("X", ".")
    if moeda:
        return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{valor:,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def extrair_sheet_id(texto):
    texto = (texto or "").strip()
    if "/d/" in texto:
        return texto.split("/d/")[1].split("/")[0]
    return texto


@st.cache_data(show_spinner="Carregando e tratando a planilha...")
def carregar_dados_excel(arquivo):
    df = pd.read_excel(arquivo)
    df = limpar_nome_colunas(df)

    colunas_numericas = [
        "Capacidade do veículo", "Total de lotes da carga", "Total de pedidos da carga",
        "Peças da carga", "Cubagem da carga", "Peso da carga", "Valor da carga",
        "Total de pedidos do lote", "Peças do lote", "Cubagem do lote", "Peso do lote", "Valor do lote"
    ]
    for col in colunas_numericas:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    for col in ["Data de criação", "Data e Hora do carregamento", "Data de geração"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    coluna_carregamento = None
    for possivel in ["Data e Hora do carregamento", "Data Hora do carregamento", "Data do carregamento", "Data Carregamento"]:
        if possivel in df.columns:
            coluna_carregamento = possivel
            break

    if coluna_carregamento:
        df["Data Carregamento"] = pd.to_datetime(df[coluna_carregamento], errors="coerce").dt.normalize()

    df = preencher_coluna_canonica(df, "Tipo de pedido", candidatos=["Tipo de pedido", "Tipo do pedido", "Tipo Pedido", "Tipo pedido", "Tipo de Pedido"], contem_todos=["tipo", "pedido"])
    df = preencher_coluna_canonica(df, "Plano de transporte", candidatos=["Plano de transporte", "Plano Transporte", "Plano de Transporte"], contem_todos=["plano", "transporte"])
    df = preencher_coluna_canonica(df, "Linha de produção", candidatos=["Linha de produção", "Linha de producao", "Linha Produção", "Linha Producao"], contem_todos=["linha", "producao"])
    df = preencher_coluna_canonica(df, "Modalidade", candidatos=["Modalidade", "Modal"], contem_todos=["modal"])
    df = preencher_coluna_canonica(df, "Destino Filial", candidatos=["Destino Filial", "Destino", "Filial destino"], contem_todos=["destino"])
    df = preencher_coluna_canonica(df, "Tipo do veículo", candidatos=["Tipo do veículo", "Tipo do veiculo", "Veículo", "Veiculo"], contem_todos=["tipo", "veiculo"])

    colunas_filtro = ["Transportador", "Plano de transporte", "Tipo do veículo", "Modalidade", "Destino Filial", "Linha de produção", "Tipo de pedido"]
    for col in colunas_filtro:
        if col in df.columns:
            df[col] = df[col].astype("string").str.strip()
            df[col] = df[col].replace({"": pd.NA, "nan": pd.NA, "None": pd.NA, "<NA>": pd.NA})

    if "Tipo de pedido" in df.columns:
        df["Tipo de pedido"] = df["Tipo de pedido"].fillna("Não informado")

    if "ID de carga" not in df.columns:
        st.error("A planilha precisa ter a coluna 'ID de carga'.")
        st.stop()

    df["ID de carga"] = df["ID de carga"].astype(str).str.strip()
    base_cargas = df.groupby("ID de carga", as_index=False, dropna=False).agg(primeiro_valido).copy()
    cargas_antes = len(base_cargas)
    if "Cubagem da carga" in base_cargas.columns:
        base_cargas = base_cargas[base_cargas["Cubagem da carga"].fillna(0) > 0]
    if "Capacidade do veículo" in base_cargas.columns:
        base_cargas = base_cargas[base_cargas["Capacidade do veículo"].fillna(0) > 0]
    base_cargas.attrs["cargas_removidas_zeradas"] = cargas_antes - len(base_cargas)

    base_cargas = calcular_indicadores(base_cargas)
    return df, base_cargas


def calcular_indicadores(df):
    df = df.copy()
    if "Cubagem da carga" in df.columns and "Capacidade do veículo" in df.columns:
        df["Ocupação %"] = df["Cubagem da carga"] / df["Capacidade do veículo"]
        df["Sobra m³"] = df["Capacidade do veículo"] - df["Cubagem da carga"]
    else:
        df["Ocupação %"] = pd.NA
        df["Sobra m³"] = pd.NA

    def classificar(x):
        if pd.isna(x): return "Sem dado"
        if x < 0.20: return "Ofensor <20%"
        if x < 0.30: return "Crítica 20%-30%"
        if x < 0.50: return "Baixa 30%-50%"
        if x < 0.80: return "Boa 50%-80%"
        if x <= 1.00: return "Alta 80%-100%"
        return "Acima da capacidade"

    df["Faixa de ocupação"] = df["Ocupação %"].apply(classificar)
    return df


def preparar_para_historico(cargas):
    df = cargas.copy()
    saida = pd.DataFrame()
    saida["ID de carga"] = df.get("ID de carga", "").astype(str)
    saida["Data Carregamento"] = pd.to_datetime(df.get("Data Carregamento"), errors="coerce").dt.strftime("%d/%m/%Y")
    saida["Plano de Transporte"] = df.get("Plano de transporte", "")
    saida["Transportador"] = df.get("Transportador", "")
    saida["Destino"] = df.get("Destino Filial", "")
    saida["Modal"] = df.get("Modalidade", "")
    saida["Linha de Produção"] = df.get("Linha de produção", "")
    saida["Tipo de Pedido"] = df.get("Tipo de pedido", "Não informado")
    saida["Tipo de Pedido"] = saida["Tipo de Pedido"].astype("string").str.strip().replace({"": "Não informado", "nan": "Não informado", "None": "Não informado", "<NA>": "Não informado"}).fillna("Não informado")
    saida["Tipo do veículo"] = df.get("Tipo do veículo", "")
    saida["Cubagem"] = pd.to_numeric(df.get("Cubagem da carga", 0), errors="coerce").fillna(0)
    saida["Capacidade"] = pd.to_numeric(df.get("Capacidade do veículo", 0), errors="coerce").fillna(0)
    saida["Sobra"] = pd.to_numeric(df.get("Sobra m³", 0), errors="coerce").fillna(0)
    saida["Valor de Carga"] = pd.to_numeric(df.get("Valor da carga", 0), errors="coerce").fillna(0)
    saida["Ocupação %"] = pd.to_numeric(df.get("Ocupação %", 0), errors="coerce").fillna(0)
    return saida[COLUNAS_HISTORICO]


def historico_para_interno(hist):
    if hist.empty:
        return pd.DataFrame()
    h = limpar_nome_colunas(hist)
    for col in COLUNAS_HISTORICO:
        if col not in h.columns:
            h[col] = ""
    saida = pd.DataFrame()
    saida["ID de carga"] = h["ID de carga"].astype(str)
    saida["Data Carregamento"] = pd.to_datetime(h["Data Carregamento"], dayfirst=True, errors="coerce").dt.normalize()
    saida["Plano de transporte"] = h["Plano de Transporte"].astype("string").str.strip()
    saida["Transportador"] = h["Transportador"].astype("string").str.strip()
    saida["Destino Filial"] = h["Destino"].astype("string").str.strip()
    saida["Modalidade"] = h["Modal"].astype("string").str.strip()
    saida["Linha de produção"] = h["Linha de Produção"].astype("string").str.strip()
    saida["Tipo de pedido"] = h["Tipo de Pedido"].astype("string").str.strip().replace({"": "Não informado", "nan": "Não informado", "None": "Não informado", "<NA>": "Não informado"}).fillna("Não informado")
    saida["Tipo do veículo"] = h["Tipo do veículo"].astype("string").str.strip()
    saida["Cubagem da carga"] = pd.to_numeric(h["Cubagem"], errors="coerce")
    saida["Capacidade do veículo"] = pd.to_numeric(h["Capacidade"], errors="coerce")
    saida["Sobra m³"] = pd.to_numeric(h["Sobra"], errors="coerce")
    saida["Valor da carga"] = pd.to_numeric(h["Valor de Carga"], errors="coerce")
    saida = saida[(saida["Cubagem da carga"].fillna(0) > 0) & (saida["Capacidade do veículo"].fillna(0) > 0)]
    saida = calcular_indicadores(saida)
    return saida


@st.cache_resource(show_spinner=False)
def conectar_google_sheets():
    if gspread is None or Credentials is None:
        raise RuntimeError("Instale as bibliotecas gspread e google-auth: pip install -r requirements.txt")

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    # 1º tenta usar service_account.json local
    if SERVICE_ACCOUNT.exists():
        creds = Credentials.from_service_account_file(str(SERVICE_ACCOUNT), scopes=scopes)
        return gspread.authorize(creds)

    # 2º tenta usar secrets.toml quando estiver publicado no Streamlit Cloud
    try:
        if "gcp_service_account" in st.secrets:
            creds = Credentials.from_service_account_info(
                dict(st.secrets["gcp_service_account"]),
                scopes=scopes
            )
            return gspread.authorize(creds)
    except Exception:
        pass

    raise FileNotFoundError(
        "Não encontrei o service_account.json na pasta do app. "
        "Coloque o arquivo na mesma pasta do app.py."
    )


def abrir_worksheet_generico(sheet_id, nome_aba, colunas, linhas=1000):
    gc = conectar_google_sheets()
    sh = gc.open_by_key(sheet_id)
    try:
        ws = sh.worksheet(nome_aba)
    except Exception:
        ws = sh.add_worksheet(title=nome_aba, rows=linhas, cols=max(len(colunas), 3))
        ws.append_row(colunas)
    valores = ws.get_all_values()
    if not valores:
        ws.append_row(colunas)
    elif valores[0] != colunas:
        # Não apaga nada; só garante que o cabeçalho esteja padronizado na primeira linha.
        ws.update("A1", [colunas])
    return ws


def abrir_worksheet(sheet_id):
    return abrir_worksheet_generico(sheet_id, ABA_HISTORICO, COLUNAS_HISTORICO)


def ler_historico(sheet_id):
    if not sheet_id:
        return pd.DataFrame()
    ws = abrir_worksheet(sheet_id)
    registros = ws.get_all_records()
    if not registros:
        return pd.DataFrame(columns=COLUNAS_HISTORICO)
    return pd.DataFrame(registros)


def escrever_historico(sheet_id, df_hist):
    ws = abrir_worksheet(sheet_id)
    df = df_hist.copy()
    df = df[COLUNAS_HISTORICO]
    df = df.fillna("")
    ws.clear()
    ws.update("A1", [COLUNAS_HISTORICO] + df.astype(str).values.tolist())


def ler_especiais(sheet_id):
    if not sheet_id:
        return pd.DataFrame(columns=COLUNAS_ESPECIAIS)
    ws = abrir_worksheet_generico(sheet_id, ABA_ESPECIAIS, COLUNAS_ESPECIAIS, linhas=300)
    registros = ws.get_all_records()
    if not registros:
        return pd.DataFrame(columns=COLUNAS_ESPECIAIS)
    df = pd.DataFrame(registros)
    for col in COLUNAS_ESPECIAIS:
        if col not in df.columns:
            df[col] = ""
    df = df[COLUNAS_ESPECIAIS].copy()
    df["ID de carga"] = df["ID de carga"].astype(str).str.strip()
    df["Tipo especial"] = df["Tipo especial"].astype(str).str.strip()
    df["Observação"] = df["Observação"].astype(str).str.strip()
    df = df[df["ID de carga"].ne("")]
    return df.drop_duplicates(subset=["ID de carga"], keep="last")


def escrever_especiais(sheet_id, df_esp):
    ws = abrir_worksheet_generico(sheet_id, ABA_ESPECIAIS, COLUNAS_ESPECIAIS, linhas=300)
    df = df_esp.copy()
    for col in COLUNAS_ESPECIAIS:
        if col not in df.columns:
            df[col] = ""
    df = df[COLUNAS_ESPECIAIS].fillna("")
    df["ID de carga"] = df["ID de carga"].astype(str).str.strip()
    df = df[df["ID de carga"].ne("")].drop_duplicates(subset=["ID de carga"], keep="last")
    ws.clear()
    ws.update("A1", [COLUNAS_ESPECIAIS] + df.astype(str).values.tolist())


def preparar_visualizacao(df_view):
    view = df_view.copy()
    for col in ["Data Carregamento", "Data de criação", "Data e Hora do carregamento", "Data de geração"]:
        if col in view.columns:
            view[col] = pd.to_datetime(view[col], errors="coerce").dt.strftime("%d/%m/%Y").fillna("")
    if "Ocupação %" in view.columns:
        view["Ocupação %"] = view["Ocupação %"].map(lambda x: br_numero(x, percentual=True) if pd.notna(x) else "")
    for col in ["Cubagem da carga", "Capacidade do veículo", "Sobra m³"]:
        if col in view.columns:
            view[col] = view[col].map(lambda x: f"{math.ceil(x):,.0f}".replace(",", ".") if pd.notna(x) else "")
    if "Valor da carga" in view.columns:
        view["Valor da carga"] = view["Valor da carga"].map(lambda x: br_numero(x, moeda=True) if pd.notna(x) else "")
    return view


def resumo_por(df, coluna):
    if coluna not in df.columns or df.empty:
        return pd.DataFrame()
    return (df.groupby(coluna, dropna=False)
            .agg(Cargas=("ID de carga", "nunique"),
                 Cubagem=("Cubagem da carga", "sum"),
                 Capacidade=("Capacidade do veículo", "sum"),
                 Ocupacao_media=("Ocupação %", "mean"),
                 Sobra=("Sobra m³", "sum"),
                 Valor_carga=("Valor da carga", "sum"))
            .reset_index()
            .sort_values("Cargas", ascending=False))


def formatar_resumo_tabela(df_resumo):
    view = df_resumo.copy()
    if "Ocupacao_media" in view.columns:
        view = view.rename(columns={"Ocupacao_media": "Ocupação média"})
        view["Ocupação média"] = view["Ocupação média"].map(lambda x: br_numero(x, percentual=True) if pd.notna(x) else "")
    if "Valor_carga" in view.columns:
        view = view.rename(columns={"Valor_carga": "Valor de carga"})
        view["Valor de carga"] = view["Valor de carga"].map(lambda x: br_numero(x, moeda=True) if pd.notna(x) else "")
    for col in ["Cubagem", "Sobra", "Capacidade"]:
        if col in view.columns:
            view[col] = view[col].map(lambda x: f"{math.ceil(x):,.0f}".replace(",", ".") if pd.notna(x) else "")
    return view


def aplicar_rotulos_barra(fig, formato=None):
    for trace in fig.data:
        if getattr(trace, "type", None) == "bar":
            trace.textposition = "outside"
            if formato:
                trace.texttemplate = formato
    fig.update_layout(uniformtext_minsize=8, uniformtext_mode="hide")
    return fig


def gerar_sugestoes(df, limite_recorrente=0.40, limite_ofensor=0.20):
    sugestoes = []
    if df.empty or "Plano de transporte" not in df.columns:
        return pd.DataFrame(columns=["Tipo", "Plano de transporte", "Indicador", "Sugestão"])

    base = df.dropna(subset=["Plano de transporte"]).copy()
    if base.empty:
        return pd.DataFrame(columns=["Tipo", "Plano de transporte", "Indicador", "Sugestão"])

    grp = base.groupby("Plano de transporte").agg(
        Cargas=("ID de carga", "nunique"),
        Dias=("Data Carregamento", lambda x: pd.Series(x).dropna().nunique()),
        Ocupacao_media=("Ocupação %", "mean"),
        Ocupacao_std=("Ocupação %", "std"),
        Sobra_media=("Sobra m³", "mean"),
        Ofensores=("Ocupação %", lambda x: (x < limite_ofensor).sum()),
        Capacidade_media=("Capacidade do veículo", "mean"),
    ).reset_index()

    baixa = grp[(grp["Dias"] >= 2) & (grp["Ocupacao_media"] < limite_recorrente)].sort_values("Ocupacao_media").head(10)
    for _, r in baixa.iterrows():
        sugestoes.append({
            "Tipo": "Baixa ocupação recorrente",
            "Plano de transporte": r["Plano de transporte"],
            "Indicador": f"{r['Ocupacao_media']*100:.1f}% médio em {int(r['Dias'])} dia(s)",
            "Sugestão": "Possível ponto de melhoria: revisar frequência, veículo ou consolidação da rota."
        })

    superdim = grp[(grp["Ocupacao_media"] < 0.50) & (grp["Sobra_media"] > 0)].sort_values("Sobra_media", ascending=False).head(10)
    for _, r in superdim.iterrows():
        sugestoes.append({
            "Tipo": "Veículo superdimensionado",
            "Plano de transporte": r["Plano de transporte"],
            "Indicador": f"Sobra média {r['Sobra_media']:.0f} m³",
            "Sugestão": "Avaliar troca para perfil de veículo menor ou agrupamento com outra demanda."
        })

    variavel = grp[(grp["Cargas"] >= 4) & (grp["Ocupacao_std"].fillna(0) > 0.25)].sort_values("Ocupacao_std", ascending=False).head(10)
    for _, r in variavel.iterrows():
        sugestoes.append({
            "Tipo": "Alta variabilidade",
            "Plano de transporte": r["Plano de transporte"],
            "Indicador": f"Oscilação {r['Ocupacao_std']*100:.1f} p.p.",
            "Sugestão": "Verificar sazonalidade, janela de corte e aderência do plano à demanda real."
        })

    freq_of = grp[grp["Ofensores"] >= 2].sort_values("Ofensores", ascending=False).head(10)
    for _, r in freq_of.iterrows():
        sugestoes.append({
            "Tipo": f"Ofensores < {limite_ofensor*100:.0f}% frequentes",
            "Plano de transporte": r["Plano de transporte"],
            "Indicador": f"{int(r['Ofensores'])} carga(s) crítica(s)",
            "Sugestão": "Priorizar análise operacional dessas cargas antes da expedição."
        })

    if "Destino Filial" in base.columns:
        dest = base[base["Ocupação %"] < 0.50].groupby(["Data Carregamento", "Destino Filial"]).agg(
            Cargas=("ID de carga", "nunique"), Cubagem=("Cubagem da carga", "sum"), Capacidade=("Capacidade do veículo", "sum")
        ).reset_index()
        dest = dest[(dest["Cargas"] >= 2) & (dest["Capacidade"] > 0)]
        dest["Ocupacao"] = dest["Cubagem"] / dest["Capacidade"]
        dest = dest[dest["Ocupacao"] < 0.60].head(10)
        for _, r in dest.iterrows():
            sugestoes.append({
                "Tipo": "Oportunidade de consolidação",
                "Plano de transporte": str(r["Destino Filial"]),
                "Indicador": f"{int(r['Cargas'])} cargas no mesmo destino/dia",
                "Sugestão": "Verificar se é possível consolidar demanda do mesmo destino em uma única saída."
            })

    return pd.DataFrame(sugestoes).drop_duplicates().head(30)

# -----------------------------
# Interface
# -----------------------------
st.title("Dashboard de Cubagem de Rotas - CD 2900")
st.caption("Análise por ID de carga, com histórico em Google Sheets, simulação, cargas especiais, tipo de pedido e sugestões automáticas.")

with st.sidebar:
    st.header("Base atual")
    upload = st.file_uploader("Enviar base mensal ou diária", type=["xlsx"])
    planilha_local = localizar_planilha_padrao()
    if upload is not None:
        arquivo = upload
        st.caption("Usando a planilha enviada no upload.")
    elif planilha_local is not None:
        arquivo = planilha_local
        st.caption(f"Usando a planilha local: {planilha_local.name}")
    else:
        st.error("Envie uma planilha .xlsx ou coloque o arquivo na pasta do app.")
        st.stop()

try:
    base_original, cargas_atual = carregar_dados_excel(arquivo)
except FileNotFoundError:
    st.error("Não encontrei a planilha. Coloque o arquivo .xlsx na pasta do app.py ou use o upload.")
    st.stop()

with st.sidebar:
    st.header("Google Sheets / Histórico")
    sheet_id = "19KiVkhZnPpAap35V6uP9R1eNGAmOlH_WZKWEHUUBMUc"
usar_historico = True

historico_raw = pd.DataFrame()
historico_interno = pd.DataFrame()
if sheet_id and usar_historico:
    try:
        historico_raw = ler_historico(sheet_id)
        historico_interno = historico_para_interno(historico_raw)
        st.sidebar.success(f"Histórico conectado: {len(historico_interno):,.0f} carga(s).".replace(",", "."))
    except Exception as e:
        st.sidebar.error(f"Não foi possível ler o histórico: {e}")

especiais_raw = pd.DataFrame(columns=COLUNAS_ESPECIAIS)
if sheet_id and usar_historico:
    try:
        especiais_raw = ler_especiais(sheet_id)
        if not especiais_raw.empty:
            st.sidebar.info(f"Cargas especiais cadastradas: {len(especiais_raw):,.0f}".replace(",", "."))
    except Exception as e:
        st.sidebar.warning(f"Não foi possível ler cargas especiais: {e}")

with st.sidebar:
    fonte = st.radio("Fonte da análise", ["Base atual", "Histórico", "Atual + Histórico"], index=0)

if fonte == "Histórico" and not historico_interno.empty:
    cargas = historico_interno.copy()
elif fonte == "Atual + Histórico" and not historico_interno.empty:
    cargas = pd.concat([historico_interno, cargas_atual], ignore_index=True)
    cargas = cargas.drop_duplicates(subset=["ID de carga"], keep="last")
else:
    cargas = cargas_atual.copy()

cargas_com_especiais = cargas.copy()
ids_especiais = set(especiais_raw.get("ID de carga", pd.Series(dtype=str)).astype(str).str.strip().tolist()) if not especiais_raw.empty else set()
if ids_especiais and "ID de carga" in cargas_com_especiais.columns:
    mapa_tipo_especial = dict(zip(especiais_raw["ID de carga"].astype(str), especiais_raw["Tipo especial"].astype(str)))
    cargas_com_especiais["Tipo especial"] = cargas_com_especiais["ID de carga"].astype(str).map(mapa_tipo_especial)
else:
    cargas_com_especiais["Tipo especial"] = pd.NA

with st.sidebar:
    st.header("Cargas especiais")
    desconsiderar_especiais = st.checkbox("Desconsiderar cargas especiais nos KPIs", value=False, help="Remove os IDs cadastrados como Cofre, Cerveja etc. da análise principal, mas mantém na aba Cargas especiais.")

if desconsiderar_especiais and ids_especiais:
    cargas = cargas_com_especiais[~cargas_com_especiais["ID de carga"].astype(str).isin(ids_especiais)].copy()
else:
    cargas = cargas_com_especiais.copy()

with st.sidebar:
    st.header("Filtros")
    removidas = cargas_atual.attrs.get("cargas_removidas_zeradas", 0)
    if removidas:
        st.caption(f"Limpeza aplicada: {removidas} carga(s) zerada(s) removida(s) da base atual.")

    def multiselect_coluna(nome, label=None):
        if nome in cargas.columns:
            opcoes = sorted(cargas[nome].dropna().astype(str).unique().tolist())
            return st.multiselect(label or nome, opcoes)
        return []

    f_transportador = multiselect_coluna("Transportador")
    f_veiculo = multiselect_coluna("Tipo do veículo")
    f_modalidade = multiselect_coluna("Modalidade", "Modal")
    f_linha = multiselect_coluna("Linha de produção")
    f_tipo_pedido = multiselect_coluna("Tipo de pedido")
    f_destino = multiselect_coluna("Destino Filial", "Destino")

    if "Data Carregamento" in cargas.columns and cargas["Data Carregamento"].notna().any():
        data_min = min(cargas["Data Carregamento"].dropna()).date()
        data_max = max(cargas["Data Carregamento"].dropna()).date()
        f_data = st.date_input("Data carregamento", value=(data_min, data_max), min_value=data_min, max_value=data_max, format="DD/MM/YYYY")
    else:
        f_data = None

    ocup_min, ocup_max = st.slider("Ocupação %", 0, 150, (0, 150), step=5)
    limite_simulacao = st.slider("Simular sem cargas abaixo de (%)", 0, 50, 20, step=5)

filtrado = cargas.copy()
for col, filtro in [
    ("Transportador", f_transportador), ("Tipo do veículo", f_veiculo), ("Modalidade", f_modalidade),
    ("Linha de produção", f_linha), ("Tipo de pedido", f_tipo_pedido), ("Destino Filial", f_destino)
]:
    if filtro and col in filtrado.columns:
        filtrado = filtrado[filtrado[col].astype(str).isin(filtro)]

if f_data and "Data Carregamento" in filtrado.columns:
    if isinstance(f_data, tuple) and len(f_data) == 2:
        dt_ini, dt_fim = pd.to_datetime(f_data[0]), pd.to_datetime(f_data[1])
        filtrado = filtrado[(filtrado["Data Carregamento"] >= dt_ini) & (filtrado["Data Carregamento"] <= dt_fim)]

filtrado = filtrado[(filtrado["Ocupação %"].fillna(0) * 100 >= ocup_min) & (filtrado["Ocupação %"].fillna(0) * 100 <= ocup_max)]
simulado = filtrado[filtrado["Ocupação %"].fillna(0) * 100 >= limite_simulacao].copy()

col1, col2, col3, col4, col5, col6 = st.columns(6)
col1.metric("Cargas", f"{filtrado['ID de carga'].nunique():,.0f}".replace(",", "."))
col2.metric("Ocupação média", br_numero(filtrado["Ocupação %"].mean(), percentual=True))
col3.metric("Cubagem total", f"{filtrado['Cubagem da carga'].sum():,.0f} m³".replace(",", "."))
col4.metric("Capacidade total", f"{filtrado['Capacidade do veículo'].sum():,.0f} m³".replace(",", "."))
col5.metric("Sobra total", f"{filtrado['Sobra m³'].clip(lower=0).sum():,.0f} m³".replace(",", "."))
col6.metric("Cargas <50%", f"{(filtrado['Ocupação %'] < 0.5).sum():,.0f}".replace(",", "."))

st.divider()

aba1, aba2, aba3, aba4, aba5, aba6, aba7, aba8 = st.tabs([
    "Visão geral", "Transportadores", "Rotas/Destinos", "Base crítica", "Dia a dia", "Simulação", "Cargas especiais", "Histórico e sugestões"
])

with aba1:
    c1, c2 = st.columns(2)
    faixa = filtrado["Faixa de ocupação"].value_counts().reset_index()
    faixa.columns = ["Faixa", "Cargas"]
    fig = px.bar(faixa, x="Faixa", y="Cargas", text="Cargas", title="Distribuição por faixa de ocupação")
    aplicar_rotulos_barra(fig)
    c1.plotly_chart(fig, use_container_width=True)

    if "Tipo do veículo" in filtrado.columns:
        veic = resumo_por(filtrado, "Tipo do veículo")
        fig2 = px.bar(veic.head(20), x="Tipo do veículo", y="Ocupacao_media", text="Ocupacao_media", title="Ocupação média por tipo de veículo")
        fig2.update_yaxes(tickformat=".0%")
        aplicar_rotulos_barra(fig2, formato="%{y:.1%}")
        c2.plotly_chart(fig2, use_container_width=True)

    st.subheader("Cargas com maior sobra de cubagem")
    cols = [c for c in ["ID de carga", "Data Carregamento", "Plano de transporte", "Transportador", "Tipo do veículo", "Modalidade", "Linha de produção", "Destino Filial", "Cubagem da carga", "Capacidade do veículo", "Ocupação %", "Sobra m³", "Valor da carga"] if c in filtrado.columns]
    st.dataframe(preparar_visualizacao(filtrado.sort_values("Sobra m³", ascending=False)[cols].head(50)), use_container_width=True)

with aba2:
    trans = resumo_por(filtrado, "Transportador")
    if not trans.empty:
        c1, c2 = st.columns(2)
        fig = px.bar(trans.head(20), x="Transportador", y="Cargas", text="Cargas", title="Top transportadores por quantidade de cargas")
        aplicar_rotulos_barra(fig)
        c1.plotly_chart(fig, use_container_width=True)
        fig2 = px.bar(trans.head(20), x="Transportador", y="Ocupacao_media", text="Ocupacao_media", title="Ocupação média por transportador")
        fig2.update_yaxes(tickformat=".0%")
        aplicar_rotulos_barra(fig2, formato="%{y:.1%}")
        c2.plotly_chart(fig2, use_container_width=True)
        st.dataframe(formatar_resumo_tabela(trans), use_container_width=True, height=520)

with aba3:
    rota_col = "Plano de transporte" if "Plano de transporte" in filtrado.columns else "Destino Filial"
    destino = resumo_por(filtrado, rota_col)
    if not destino.empty:
        st.subheader("Resumo por plano de transporte")
        fig = px.bar(destino.head(30), x=rota_col, y="Sobra", text="Sobra", title="Planos de transporte com maior sobra de cubagem")
        aplicar_rotulos_barra(fig, formato="%{y:,.0f}")
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(formatar_resumo_tabela(destino), use_container_width=True, height=520)

with aba4:
    st.subheader("Base crítica")
    criticas = filtrado[filtrado["Ocupação %"] < 0.5].sort_values("Ocupação %")
    cols = [c for c in ["ID de carga", "Data Carregamento", "Plano de transporte", "Transportador", "Tipo de pedido", "Tipo do veículo", "Modalidade", "Linha de produção", "Cubagem da carga", "Capacidade do veículo", "Ocupação %", "Sobra m³", "Valor da carga"] if c in criticas.columns]
    st.dataframe(preparar_visualizacao(criticas[cols]), use_container_width=True, height=600)
    csv = preparar_visualizacao(criticas[cols]).to_csv(index=False, sep=";").encode("utf-8-sig")
    st.download_button("Baixar cargas críticas em CSV", csv, "cargas_criticas.csv", "text/csv")

with aba5:
    st.subheader("Dia a dia por data de carregamento")
    if "Data Carregamento" not in filtrado.columns or filtrado["Data Carregamento"].dropna().empty:
        st.info("Não encontrei data de carregamento válida.")
    else:
        base_dia = filtrado.dropna(subset=["Data Carregamento"]).copy()
        resumo_dia = (base_dia.groupby("Data Carregamento", dropna=False)
            .agg(Cargas=("ID de carga", "nunique"), Cubagem=("Cubagem da carga", "sum"), Capacidade=("Capacidade do veículo", "sum"), Sobra=("Sobra m³", "sum"))
            .reset_index().sort_values("Data Carregamento"))
        resumo_dia["Ocupacao_dia"] = resumo_dia["Cubagem"] / resumo_dia["Capacidade"]
        resumo_dia["Data_formatada"] = pd.to_datetime(resumo_dia["Data Carregamento"], errors="coerce").dt.strftime("%d/%m/%Y")

        mapa_datas = dict(zip(resumo_dia["Data_formatada"], resumo_dia["Data Carregamento"]))
        data_opcoes = resumo_dia["Data_formatada"].dropna().tolist()
        data_label = st.selectbox("Escolha o dia para ver ofensores/performance", data_opcoes, index=len(data_opcoes)-1 if data_opcoes else 0)
        data_escolhida = mapa_datas.get(data_label)
        dia_sel = base_dia[base_dia["Data Carregamento"] == data_escolhida].copy()

        k1, k2, k3, k4, k5 = st.columns(5)
        cap_dia = dia_sel["Capacidade do veículo"].sum()
        cub_dia = dia_sel["Cubagem da carga"].sum()
        k1.metric("Data selecionada", data_label)
        k2.metric("Ocupação do dia", br_numero(cub_dia / cap_dia if cap_dia else 0, percentual=True))
        k3.metric("Cargas do dia", f"{dia_sel['ID de carga'].nunique():,.0f}".replace(",", "."))
        k4.metric("Cubagem do dia", f"{cub_dia:,.0f} m³".replace(",", "."))
        k5.metric("Capacidade do dia", f"{cap_dia:,.0f} m³".replace(",", "."))

        fig_dia = make_subplots(specs=[[{"secondary_y": True}]])
        fig_dia.add_trace(go.Bar(x=resumo_dia["Data_formatada"], y=resumo_dia["Cubagem"], name="Cubagem diária (m³)", text=resumo_dia["Cubagem"].round(0), textposition="outside", texttemplate="%{text:,.0f}"), secondary_y=False)
        fig_dia.add_trace(go.Scatter(x=resumo_dia["Data_formatada"], y=resumo_dia["Ocupacao_dia"], name="Ocupação do dia", mode="lines+markers+text", text=[br_numero(x, percentual=True) for x in resumo_dia["Ocupacao_dia"]], textposition="top center"), secondary_y=True)
        fig_dia.update_layout(height=580, margin=dict(t=40, b=20), legend=dict(orientation="h"))
        fig_dia.update_yaxes(title_text="Cubagem diária (m³)", secondary_y=False)
        fig_dia.update_yaxes(title_text="Ocupação", tickformat=".0%", secondary_y=True)
        st.plotly_chart(fig_dia, use_container_width=True)

        plano_col = "Plano de transporte" if "Plano de transporte" in dia_sel.columns else "Transportador"
        cols_dia = [c for c in ["ID de carga", "Data Carregamento", plano_col, "Tipo de pedido", "Tipo do veículo", "Modalidade", "Linha de produção", "Destino Filial", "Cubagem da carga", "Capacidade do veículo", "Ocupação %", "Sobra m³", "Valor da carga"] if c in dia_sel.columns]
        st.markdown("#### 10 maiores ofensores do dia")
        ofensores = dia_sel.sort_values(["Ocupação %", "Sobra m³"], ascending=[True, False]).head(10)
        st.dataframe(preparar_visualizacao(ofensores[cols_dia]), use_container_width=True, height=430)
        st.markdown("#### 10 melhores performances do dia")
        melhores = dia_sel[dia_sel["Ocupação %"] <= 1].sort_values("Ocupação %", ascending=False).head(10)
        if melhores.empty:
            melhores = dia_sel.sort_values("Ocupação %", ascending=False).head(10)
        st.dataframe(preparar_visualizacao(melhores[cols_dia]), use_container_width=True, height=430)

with aba6:
    st.subheader("Simulação sem maiores ofensores")
    st.caption(f"Simulação removendo cargas com ocupação abaixo de {limite_simulacao}%.")
    a1, a2, a3, a4 = st.columns(4)
    a1.metric("Cargas removidas", f"{len(filtrado) - len(simulado):,.0f}".replace(",", "."))
    a2.metric("Ocupação atual", br_numero(filtrado["Ocupação %"].mean(), percentual=True))
    a3.metric("Ocupação simulada", br_numero(simulado["Ocupação %"].mean(), percentual=True))
    ganho = simulado["Ocupação %"].mean() - filtrado["Ocupação %"].mean()
    a4.metric("Ganho estimado", br_numero(ganho, percentual=True))

    comp = pd.DataFrame({
        "Cenário": ["Atual", f"Sem cargas < {limite_simulacao}%"],
        "Cargas": [filtrado["ID de carga"].nunique(), simulado["ID de carga"].nunique()],
        "Cubagem": [filtrado["Cubagem da carga"].sum(), simulado["Cubagem da carga"].sum()],
        "Capacidade": [filtrado["Capacidade do veículo"].sum(), simulado["Capacidade do veículo"].sum()],
    })
    comp["Ocupação"] = comp["Cubagem"] / comp["Capacidade"]
    fig_sim = px.bar(comp, x="Cenário", y="Ocupação", text="Ocupação", title="Comparativo de ocupação")
    fig_sim.update_yaxes(tickformat=".0%")
    aplicar_rotulos_barra(fig_sim, formato="%{y:.1%}")
    st.plotly_chart(fig_sim, use_container_width=True)
    st.dataframe(comp.assign(Ocupação=comp["Ocupação"].map(lambda x: br_numero(x, percentual=True))), use_container_width=True)

with aba7:
    st.subheader("Cargas especiais")
    st.caption("Cadastre IDs de carga que distorcem a cubagem, como Cofre ou Cerveja. Eles podem ser desconsiderados dos KPIs pela opção do menu lateral.")

    if not sheet_id:
        st.warning("Cole o link ou ID do Google Sheets no menu lateral para cadastrar e salvar cargas especiais.")
    else:
        col_a, col_b = st.columns([1, 1])
        with col_a:
            novo_id = st.text_input("ID de carga")
            novo_tipo = st.selectbox("Tipo especial", ["Cofre", "Cerveja", "Outro"], index=0)
            nova_obs = st.text_input("Observação", placeholder="Opcional")
            if st.button("Adicionar carga especial", type="primary"):
                try:
                    esp = ler_especiais(sheet_id)
                    nova_linha = pd.DataFrame([{"ID de carga": str(novo_id).strip(), "Tipo especial": novo_tipo, "Observação": nova_obs}])
                    if not str(novo_id).strip():
                        st.warning("Informe um ID de carga.")
                    else:
                        esp = pd.concat([esp, nova_linha], ignore_index=True)
                        escrever_especiais(sheet_id, esp)
                        st.success("Carga especial cadastrada.")
                        st.cache_data.clear()
                except Exception as e:
                    st.error(f"Erro ao cadastrar carga especial: {e}")
        with col_b:
            if not especiais_raw.empty:
                ids_remover = st.multiselect("Remover IDs cadastrados", sorted(especiais_raw["ID de carga"].astype(str).tolist()))
                if st.button("Excluir IDs selecionados"):
                    try:
                        esp = ler_especiais(sheet_id)
                        esp = esp[~esp["ID de carga"].astype(str).isin(ids_remover)].copy()
                        escrever_especiais(sheet_id, esp)
                        st.success("ID(s) removido(s) da lista de cargas especiais.")
                        st.cache_data.clear()
                    except Exception as e:
                        st.error(f"Erro ao excluir ID(s): {e}")
            else:
                st.info("Nenhuma carga especial cadastrada ainda.")

    if ids_especiais and "ID de carga" in cargas_com_especiais.columns:
        base_esp = cargas_com_especiais[cargas_com_especiais["ID de carga"].astype(str).isin(ids_especiais)].copy()
    else:
        base_esp = pd.DataFrame()

    if base_esp.empty:
        st.info("Nenhuma carga especial encontrada na base analisada.")
    else:
        e1, e2, e3, e4 = st.columns(4)
        e1.metric("Cargas especiais", f"{base_esp['ID de carga'].nunique():,.0f}".replace(",", "."))
        e2.metric("Cubagem especial", f"{base_esp['Cubagem da carga'].sum():,.0f} m³".replace(",", "."))
        e3.metric("Capacidade especial", f"{base_esp['Capacidade do veículo'].sum():,.0f} m³".replace(",", "."))
        e4.metric("Ocupação média especial", br_numero(base_esp["Ocupação %"].mean(), percentual=True))

        if not filtrado.empty:
            ocup_sem = filtrado["Ocupação %"].mean()
            base_completa_sem_filtro = cargas_com_especiais.copy()
            ocup_com = base_completa_sem_filtro["Ocupação %"].mean()
            st.info(f"Impacto estimado: ocupação média com especiais = {br_numero(ocup_com, percentual=True)} | cenário atual sem especiais = {br_numero(ocup_sem, percentual=True)}")

        resumo_esp = resumo_por(base_esp, "Tipo especial") if "Tipo especial" in base_esp.columns else pd.DataFrame()
        if not resumo_esp.empty:
            fig_esp = px.bar(resumo_esp, x="Tipo especial", y="Cargas", text="Cargas", title="Cargas especiais por tipo")
            aplicar_rotulos_barra(fig_esp)
            st.plotly_chart(fig_esp, use_container_width=True)
            st.dataframe(formatar_resumo_tabela(resumo_esp), use_container_width=True)

        cols_esp = [c for c in ["ID de carga", "Tipo especial", "Data Carregamento", "Plano de transporte", "Transportador", "Tipo de pedido", "Tipo do veículo", "Modalidade", "Linha de produção", "Destino Filial", "Cubagem da carga", "Capacidade do veículo", "Ocupação %", "Sobra m³", "Valor da carga"] if c in base_esp.columns]
        st.dataframe(preparar_visualizacao(base_esp[cols_esp]), use_container_width=True, height=420)

with aba8:
    st.subheader("Histórico Google Sheets")
    st.caption("O app salva somente a base limpa por carga, não a base bruta por lote.")
    if not sheet_id:
        st.warning("Cole o link ou ID do Google Sheets no menu lateral para usar o histórico.")
    else:
        c1, c2 = st.columns(2)
        with c1:
            if st.button("Adicionar base atual ao histórico", type="primary"):
                try:
                    hist_existente = ler_historico(sheet_id)
                    novo = preparar_para_historico(cargas_atual)
                    combinado = pd.concat([hist_existente, novo], ignore_index=True) if not hist_existente.empty else novo
                    combinado["ID de carga"] = combinado["ID de carga"].astype(str)
                    combinado = combinado.drop_duplicates(subset=["ID de carga"], keep="last")
                    escrever_historico(sheet_id, combinado)
                    st.success(f"Histórico atualizado com {len(combinado):,.0f} carga(s).".replace(",", "."))
                    st.cache_data.clear()
                except Exception as e:
                    st.error(f"Erro ao salvar no histórico: {e}")
        with c2:
            try:
                hist_atual = ler_historico(sheet_id)
                datas_hist = sorted(hist_atual.get("Data Carregamento", pd.Series(dtype=str)).dropna().astype(str).unique().tolist()) if not hist_atual.empty else []
                datas_excluir = st.multiselect("Excluir histórico por data", datas_hist)
                if st.button("Excluir datas selecionadas"):
                    if datas_excluir:
                        hist_novo = hist_atual[~hist_atual["Data Carregamento"].astype(str).isin(datas_excluir)].copy()
                        escrever_historico(sheet_id, hist_novo)
                        st.success("Datas excluídas do histórico.")
                        st.cache_data.clear()
                    else:
                        st.info("Selecione pelo menos uma data para excluir.")
            except Exception as e:
                st.error(f"Erro ao gerenciar histórico: {e}")

    st.subheader("Sugestões automáticas")
    sugestoes = gerar_sugestoes(filtrado, limite_recorrente=0.40, limite_ofensor=limite_simulacao / 100)
    if sugestoes.empty:
        st.success("Nenhum ponto crítico recorrente encontrado com os filtros atuais.")
    else:
        st.dataframe(sugestoes, use_container_width=True, height=520)
