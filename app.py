import math
import re
from pathlib import Path

import numpy as np
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

st.set_page_config(page_title="Painel Logístico CD 2900", layout="wide")

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

ABA_NS_1P = "NS_1P"
ABA_NS_FULL = "NS_FULL"
COLUNAS_NS_HISTORICO = [
    "Operação", "Data Prevista", "Modal", "Transportador", "Cidade",
    "Canal", "Situação", "Prazo Cliente", "Faixa Prazo", "Ofensor NS",
    "Total Pedidos", "Dentro Prazo", "Fora Prazo", "NS"
]

ABA_ECLUSA = "ECLUSA"
COLUNAS_ECLUSA_HISTORICO = [
    "Operação", "Data Nota", "Início Semana", "Semana", "Modal", "Transportador Grupo", "Cidade",
    "Status Saída", "Faixa Atraso", "Total Pedidos", "Dentro Prazo", "Fora Prazo",
    "NS", "Faltam Meta", "Atraso Médio h", "Maior Atraso h"
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




def semana_curta(valor):
    data = pd.to_datetime(valor, errors="coerce")
    if pd.isna(data):
        return ""
    return data.strftime("%d/%m")


def nome_mes_pt(valor):
    data = pd.to_datetime(valor, errors="coerce")
    if pd.isna(data):
        return ""
    meses = {
        1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
        5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
        9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
    }
    return meses.get(int(data.month), "")

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

    df = preencher_coluna_canonica(df, "Tipo de Pedido", candidatos=["Tipo de Pedido", "Tipo do Pedido", "Tipo Pedido", "Tipo Pedido", "Tipo de Pedido"], contem_todos=["tipo", "pedido"])
    df = preencher_coluna_canonica(df, "Plano de transporte", candidatos=["Plano de transporte", "Plano Transporte", "Plano de Transporte"], contem_todos=["plano", "transporte"])
    df = preencher_coluna_canonica(df, "Linha de produção", candidatos=["Linha de produção", "Linha de producao", "Linha Produção", "Linha Producao"], contem_todos=["linha", "producao"])
    df = preencher_coluna_canonica(df, "Modalidade", candidatos=["Modalidade", "Modal"], contem_todos=["modal"])
    df = preencher_coluna_canonica(df, "Destino Filial", candidatos=["Destino Filial", "Destino", "Filial destino"], contem_todos=["destino"])
    df = preencher_coluna_canonica(df, "Tipo do veículo", candidatos=["Tipo do veículo", "Tipo do veiculo", "Veículo", "Veiculo"], contem_todos=["tipo", "veiculo"])

    colunas_filtro = ["Transportador", "Plano de transporte", "Tipo do veículo", "Modalidade", "Destino Filial", "Linha de produção", "Tipo de Pedido"]
    for col in colunas_filtro:
        if col in df.columns:
            df[col] = df[col].astype("string").str.strip()
            df[col] = df[col].replace({"": pd.NA, "nan": pd.NA, "None": pd.NA, "<NA>": pd.NA})

    if "Tipo de Pedido" in df.columns:
        df["Tipo de Pedido"] = df["Tipo de Pedido"].fillna("Não informado")

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
    saida["Tipo de Pedido"] = df.get("Tipo de Pedido", "Não informado")
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
    saida["Tipo de Pedido"] = h["Tipo de Pedido"].astype("string").str.strip().replace({"": "Não informado", "nan": "Não informado", "None": "Não informado", "<NA>": "Não informado"}).fillna("Não informado")
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

    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

    # Local: usa service_account.json na pasta do app
    if SERVICE_ACCOUNT.exists():
        creds = Credentials.from_service_account_file(str(SERVICE_ACCOUNT), scopes=scopes)
        return gspread.authorize(creds)

    # Cloud: usa Secrets do Streamlit
    try:
        if "gcp_service_account" in st.secrets:
            creds = Credentials.from_service_account_info(dict(st.secrets["gcp_service_account"]), scopes=scopes)
            return gspread.authorize(creds)
    except Exception:
        pass

    raise FileNotFoundError(
        "Não encontrei credenciais. No Streamlit Cloud, cadastre gcp_service_account em Settings > Secrets. "
        "No computador, coloque service_account.json na pasta do app.py."
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


@st.cache_data(ttl=300, show_spinner=False)
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


@st.cache_data(ttl=300, show_spinner=False)
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
# Funções de Nível de Serviço 1P / Full
# -----------------------------
def localizar_arquivo_csv(padroes):
    for padrao in padroes:
        encontrados = sorted(PASTA_APP.glob(padrao))
        if encontrados:
            return encontrados[0]
    return None


@st.cache_data(show_spinner="Carregando nível de serviço...")
def carregar_ns_csv(arquivo, operacao):
    if arquivo is None:
        return pd.DataFrame()

    try:
        df = pd.read_csv(arquivo, sep=";", low_memory=False, encoding="utf-8-sig")
    except UnicodeDecodeError:
        df = pd.read_csv(arquivo, sep=";", low_memory=False, encoding="latin1")
    except Exception:
        try:
            df = pd.read_csv(arquivo, sep=";", low_memory=False, encoding="latin1")
        except Exception:
            df = pd.read_csv(arquivo, low_memory=False)

    df = limpar_nome_colunas(df)

    df = preencher_coluna_canonica(
        df, "Data Prevista",
        candidatos=["data entrega prevista cliente", "data_entrega_prevista_cliente", "data_entrega_prevista_cliente ", "Data Entrega Prevista Cliente"],
        contem_todos=["data", "entrega", "prevista", "cliente"]
    )
    df = preencher_coluna_canonica(df, "Modal", candidatos=["Modal Transp", "Modal_Transp", "Modal"], contem_todos=["modal"])
    df = preencher_coluna_canonica(df, "Cidade", candidatos=["cidade cliente", "Cidade Cliente"], contem_todos=["cidade", "cliente"])
    df = preencher_coluna_canonica(df, "Transportador", candidatos=["transportador (grupo)", "Transportador Grupo", "transportador", "Transportador"], contem_todos=["transportador"])
    df = preencher_coluna_canonica(df, "Pedido", candidatos=["pedido_pacote", "Pedido Pacote", "PEDIDO", "pedido_gemco"], contem_todos=["pedido"])
    df = preencher_coluna_canonica(df, "Situação", candidatos=["situacao", "situação", "status_cliente", "Status_Cliente"], contem_todos=["situacao"])
    df = preencher_coluna_canonica(df, "Prazo Cliente", candidatos=["prazo_cliente", "prazo_cliente (grupo)", "Faixa Prazo Cliente", "prazo cliente"], contem_todos=["prazo", "cliente"])
    df = preencher_coluna_canonica(df, "Ofensor NS", candidatos=["Ofensor NS", "Ofensor_NS", "ofensor ns", "ofensor_ns"], contem_todos=["ofensor", "ns"])

    if operacao == "1P":
        df = preencher_coluna_canonica(df, "Canal", candidatos=["ECC", "Ecom em casa", "canal_venda"], contem_todos=["ecc"])
        df = preencher_coluna_canonica(df, "Dentro Prazo", candidatos=["NS", "Qtd Dentro", "qtd dentro", "ns"], contem_todos=["ns"])
    else:
        df = preencher_coluna_canonica(df, "Canal", candidatos=["Ecom em casa", "ECC", "canal_venda"], contem_todos=["ecom", "casa"])
        df = preencher_coluna_canonica(df, "Dentro Prazo", candidatos=["ns", "qtd dentro", "Qtd Dentro", "NS"], contem_todos=["ns"])

    colunas = ["Data Prevista", "Modal", "Cidade", "Transportador", "Pedido", "Dentro Prazo", "Canal", "Situação", "Prazo Cliente", "Ofensor NS"]
    for col in colunas:
        if col not in df.columns:
            df[col] = pd.NA

    base = df[colunas].copy()
    base["Operação"] = "1P MAGALU" if operacao == "1P" else "3P FULL"
    base["Data Prevista"] = pd.to_datetime(base["Data Prevista"], dayfirst=True, errors="coerce").dt.normalize()
    base["Início Semana"] = base["Data Prevista"] - pd.to_timedelta(base["Data Prevista"].dt.weekday, unit="D")
    base["Semana Mês"] = base["Início Semana"].dt.strftime("%d/%m")

    for col in ["Modal", "Cidade", "Transportador", "Pedido", "Canal", "Situação", "Ofensor NS"]:
        base[col] = base[col].astype("string").str.strip().replace({"": pd.NA, "nan": pd.NA, "None": pd.NA, "<NA>": pd.NA})
        base[col] = base[col].fillna("Não informado")

    # Situação existe de fato no 1P. Valores vazios representam operação normal.
    # No Full, onde normalmente não há essa coluna, mantemos tudo como Operação Normal.
    if operacao == "1P":
        base["Situação"] = (base["Situação"]
            .astype("string")
            .str.strip()
            .replace({"": "Operação Normal", "nan": "Operação Normal", "None": "Operação Normal", "<NA>": "Operação Normal", "Não informado": "Operação Normal"})
            .fillna("Operação Normal"))
        base["Situação"] = base["Situação"].replace({"ROTY": "Roteiro Y", "ROTY ": "Roteiro Y", "Roty": "Roteiro Y"})
    else:
        base["Situação"] = "Operação Normal"

    base["Prazo Cliente"] = pd.to_numeric(base["Prazo Cliente"], errors="coerce")
    base["Faixa Prazo"] = base["Prazo Cliente"].apply(
        lambda x: "Maior D5" if pd.notna(x) and x > 5 else (f"D{int(x)}" if pd.notna(x) else "Não informado")
    )

    base["Dentro Prazo"] = pd.to_numeric(base["Dentro Prazo"], errors="coerce").fillna(0)
    base["Dentro Prazo"] = base["Dentro Prazo"].apply(lambda x: 1 if x >= 1 else 0)
    base = base.dropna(subset=["Pedido"])
    base = base[base["Pedido"].astype(str).str.strip().ne("")]

    base = (base.groupby(["Operação", "Pedido"], as_index=False, dropna=False)
            .agg({
                "Data Prevista": primeiro_valido,
                "Início Semana": primeiro_valido,
                "Semana Mês": primeiro_valido,
                "Modal": primeiro_valido,
                "Cidade": primeiro_valido,
                "Transportador": primeiro_valido,
                "Canal": primeiro_valido,
                "Situação": primeiro_valido,
                "Prazo Cliente": primeiro_valido,
                "Faixa Prazo": primeiro_valido,
                "Ofensor NS": primeiro_valido,
                "Dentro Prazo": "max"
            }))
    base["Fora Prazo"] = 1 - base["Dentro Prazo"]
    base["Total_pedidos"] = 1
    base["NS"] = base["Dentro Prazo"]
    return base


def preparar_ns_para_historico(ns):
    """Prepara histórico NS de forma agregada para não pesar o Google Sheets.

    A granularidade salva é: Operação + Data + Modal + Transportador + Cidade + Canal + Situação + Faixa Prazo.
    Assim mantemos análises por dia, modal, cidade, transportador, canal e prazo sem gravar pedido a pedido.
    """
    if ns.empty:
        return pd.DataFrame(columns=COLUNAS_NS_HISTORICO)

    df = ns.copy()
    for col in ["Operação", "Modal", "Transportador", "Cidade", "Canal", "Situação", "Faixa Prazo", "Ofensor NS"]:
        if col not in df.columns:
            df[col] = "Não informado"
        df[col] = df[col].astype("string").str.strip().replace({"": "Não informado", "nan": "Não informado", "None": "Não informado", "<NA>": "Não informado"}).fillna("Não informado")

    if "Data Prevista" not in df.columns:
        df["Data Prevista"] = pd.NaT
    df["Data Prevista"] = pd.to_datetime(df["Data Prevista"], errors="coerce").dt.normalize()

    if "Total_pedidos" not in df.columns:
        df["Total_pedidos"] = 1
    df["Total_pedidos"] = pd.to_numeric(df["Total_pedidos"], errors="coerce").fillna(1)
    df["Dentro Prazo"] = pd.to_numeric(df.get("Dentro Prazo", 0), errors="coerce").fillna(0)
    df["Fora Prazo"] = pd.to_numeric(df.get("Fora Prazo", 0), errors="coerce").fillna(0)
    df["Prazo Cliente"] = pd.to_numeric(df.get("Prazo Cliente", pd.NA), errors="coerce")

    dims = ["Operação", "Data Prevista", "Modal", "Transportador", "Cidade", "Canal", "Situação", "Faixa Prazo", "Ofensor NS"]
    resumo = (df.groupby(dims, dropna=False)
              .agg(**{
                  "Total Pedidos": ("Total_pedidos", "sum"),
                  "Dentro Prazo": ("Dentro Prazo", "sum"),
                  "Fora Prazo": ("Fora Prazo", "sum"),
                  "Prazo Cliente": ("Prazo Cliente", primeiro_valido),
              })
              .reset_index())
    resumo["NS"] = resumo["Dentro Prazo"] / resumo["Total Pedidos"].replace(0, np.nan)
    resumo["Data Prevista"] = pd.to_datetime(resumo["Data Prevista"], errors="coerce").dt.strftime("%d/%m/%Y")
    return resumo[COLUNAS_NS_HISTORICO]


def historico_ns_para_interno(hist):
    if hist.empty:
        return pd.DataFrame(columns=COLUNAS_NS_HISTORICO)
    h = limpar_nome_colunas(hist)

    for col in COLUNAS_NS_HISTORICO:
        if col not in h.columns:
            h[col] = ""

    df = h.copy()
    df["Data Prevista"] = pd.to_datetime(df["Data Prevista"], dayfirst=True, errors="coerce").dt.normalize()
    df["Início Semana"] = df["Data Prevista"] - pd.to_timedelta(df["Data Prevista"].dt.weekday, unit="D")
    df["Semana Mês"] = df["Início Semana"].dt.strftime("%d/%m")

    for col in ["Operação", "Modal", "Transportador", "Cidade", "Canal", "Situação", "Faixa Prazo", "Ofensor NS"]:
        df[col] = df[col].astype("string").str.strip().replace({"": "Não informado", "nan": "Não informado", "None": "Não informado", "<NA>": "Não informado"}).fillna("Não informado")

    if "Total Pedidos" in df.columns and df["Total Pedidos"].astype(str).str.strip().ne("").any():
        df["Total_pedidos"] = pd.to_numeric(df["Total Pedidos"], errors="coerce").fillna(0)
    elif "Pedido" in df.columns:
        df["Total_pedidos"] = 1
    else:
        df["Total_pedidos"] = 0

    df["Dentro Prazo"] = pd.to_numeric(df.get("Dentro Prazo", 0), errors="coerce").fillna(0)
    df["Fora Prazo"] = pd.to_numeric(df.get("Fora Prazo", 0), errors="coerce").fillna(df["Total_pedidos"] - df["Dentro Prazo"])
    df["Prazo Cliente"] = pd.to_numeric(df.get("Prazo Cliente", pd.NA), errors="coerce")
    df["NS"] = df["Dentro Prazo"] / df["Total_pedidos"].replace(0, np.nan)
    return df


@st.cache_data(ttl=300, show_spinner=False)
def ler_historico_ns(sheet_id, operacao):
    aba = ABA_NS_1P if operacao == "1P" else ABA_NS_FULL
    if not sheet_id:
        return pd.DataFrame(columns=COLUNAS_NS_HISTORICO)
    ws = abrir_worksheet_generico(sheet_id, aba, COLUNAS_NS_HISTORICO, linhas=1000)
    registros = ws.get_all_records()
    if not registros:
        return pd.DataFrame(columns=COLUNAS_NS_HISTORICO)
    return pd.DataFrame(registros)


def escrever_historico_ns(sheet_id, operacao, df_ns_hist):
    aba = ABA_NS_1P if operacao == "1P" else ABA_NS_FULL
    ws = abrir_worksheet_generico(sheet_id, aba, COLUNAS_NS_HISTORICO, linhas=max(len(df_ns_hist) + 10, 1000))
    df = df_ns_hist.copy()
    for col in COLUNAS_NS_HISTORICO:
        if col not in df.columns:
            df[col] = ""
    df = df[COLUNAS_NS_HISTORICO].fillna("")
    ws.clear()
    ws.update("A1", [COLUNAS_NS_HISTORICO] + df.astype(str).values.tolist())


def juntar_ns_sem_duplicar(hist_existente, novo):
    # Substitui no histórico as datas presentes no novo arquivo.
    # Se a base nova vier com desconsiderações/pedidos removidos, o histórico daquela data é regravado corretamente.
    combinado = juntar_historico_substituindo_datas(hist_existente, novo, "Data Prevista", COLUNAS_NS_HISTORICO)
    if combinado.empty:
        return combinado
    chave = ["Operação", "Data Prevista", "Modal", "Transportador", "Cidade", "Canal", "Situação", "Faixa Prazo", "Ofensor NS"]
    for col in chave:
        combinado[col] = combinado[col].astype(str).str.strip()
    return combinado.drop_duplicates(subset=chave, keep="last")[COLUNAS_NS_HISTORICO]


def consolidar_ns(df, dimensoes):
    if df.empty:
        return pd.DataFrame()
    if isinstance(dimensoes, str):
        dimensoes = [dimensoes]
    base = df.copy()
    if "Total_pedidos" not in base.columns:
        base["Total_pedidos"] = 1
    base["Total_pedidos"] = pd.to_numeric(base["Total_pedidos"], errors="coerce").fillna(1)
    base["Dentro Prazo"] = pd.to_numeric(base.get("Dentro Prazo", 0), errors="coerce").fillna(0)
    base["Fora Prazo"] = pd.to_numeric(base.get("Fora Prazo", 0), errors="coerce").fillna(base["Total_pedidos"] - base["Dentro Prazo"])

    resumo = (base.groupby(dimensoes, dropna=False)
              .agg(Total_pedidos=("Total_pedidos", "sum"),
                   Dentro_prazo=("Dentro Prazo", "sum"),
                   Fora_prazo=("Fora Prazo", "sum"))
              .reset_index())
    resumo["NS"] = resumo["Dentro_prazo"] / resumo["Total_pedidos"].replace(0, np.nan)
    return resumo




def consolidar_ofensor_ns_por_modal(df):
    if df.empty or "Ofensor NS" not in df.columns:
        return pd.DataFrame()

    base = df.copy()
    for col in ["Operação", "Modal", "Ofensor NS", "Transportador", "Cidade"]:
        if col not in base.columns:
            base[col] = "Não informado"
        base[col] = base[col].astype("string").str.strip().replace({
            "": "Não informado", "nan": "Não informado", "None": "Não informado", "<NA>": "Não informado"
        }).fillna("Não informado")

    base["Total_pedidos"] = pd.to_numeric(base.get("Total_pedidos", 1), errors="coerce").fillna(1)
    base["Dentro Prazo"] = pd.to_numeric(base.get("Dentro Prazo", 0), errors="coerce").fillna(0)
    base["Fora Prazo"] = pd.to_numeric(base.get("Fora Prazo", base["Total_pedidos"] - base["Dentro Prazo"]), errors="coerce").fillna(0)

    # Para o indicador de ofensor, consideramos somente registros que ficaram fora do prazo
    # e que possuem um motivo/ofensor válido preenchido na base.
    base = base[(base["Fora Prazo"] > 0) & (~base["Ofensor NS"].str.upper().isin(["NÃO INFORMADO", "SEM ATRASO", "SEM OFENSOR", "NAN", "NONE"]))]
    if base.empty:
        return pd.DataFrame()

    def maior_impacto(grupo, coluna):
        tmp = (grupo.groupby(coluna, dropna=False)["Fora Prazo"].sum()
               .sort_values(ascending=False))
        if tmp.empty:
            return "Não informado"
        nome = tmp.index[0]
        qtd = int(tmp.iloc[0])
        return f"{nome} ({qtd})"

    resumo = (base.groupby(["Operação", "Modal", "Ofensor NS"], dropna=False)
              .agg(**{
                  "Pedidos fora": ("Fora Prazo", "sum"),
                  "Total pedidos": ("Total_pedidos", "sum"),
                  "Dentro prazo": ("Dentro Prazo", "sum"),
                  "Transportador mais impactado": ("Transportador", lambda x: maior_impacto(base.loc[x.index], "Transportador")),
                  "Cidade mais impactada": ("Cidade", lambda x: maior_impacto(base.loc[x.index], "Cidade")),
              })
              .reset_index())
    total_modal = resumo.groupby(["Operação", "Modal"])["Pedidos fora"].transform("sum")
    resumo["% dos atrasos do modal"] = resumo["Pedidos fora"] / total_modal.replace(0, np.nan)
    resumo["NS do ofensor"] = resumo["Dentro prazo"] / resumo["Total pedidos"].replace(0, np.nan)

    colunas = [
        "Operação", "Modal", "Ofensor NS", "Pedidos fora", "% dos atrasos do modal",
        "Total pedidos", "Dentro prazo", "NS do ofensor",
        "Transportador mais impactado", "Cidade mais impactada"
    ]
    return resumo[colunas].sort_values(["Operação", "Modal", "Pedidos fora"], ascending=[True, True, False])


def formatar_ofensor_ns_tabela(df):
    view = df.copy()
    for col in ["Pedidos fora", "Total pedidos", "Dentro prazo"]:
        if col in view.columns:
            view[col] = view[col].map(lambda x: f"{int(x):,}".replace(",", ".") if pd.notna(x) else "")
    for col in ["% dos atrasos do modal", "NS do ofensor"]:
        if col in view.columns:
            view[col] = view[col].map(lambda x: br_numero(x, percentual=True) if pd.notna(x) else "")
    return view

def classificar_ns(valor, operacao):
    if pd.isna(valor):
        return "Sem dado", "#9CA3AF", 0
    if operacao == "1P MAGALU":
        if valor < 0.95:
            return "Zerado", "#6B21A8", 0
        if valor < 0.96:
            return "80%", "#DC2626", 80
        if valor < 0.97:
            return "100%", "#16A34A", 100
        return "120%", "#2563EB", 120
    else:
        if valor < 0.94:
            return "Zerado", "#6B21A8", 0
        if valor < 0.95:
            return "80%", "#DC2626", 80
        if valor < 0.96:
            return "100%", "#16A34A", 100
        return "120%", "#2563EB", 120


def meta_texto_ns(operacao):
    if operacao == "1P MAGALU":
        return "1P: <95% roxo | 95% vermelho/80% | 96% verde/100% | 97%+ azul/120%"
    return "Full: <94% roxo | 94% vermelho/80% | 95% verde/100% | 96%+ azul/120%"


def card_ns(titulo, ns, total, dentro, fora, operacao):
    faixa, cor, pontuacao = classificar_ns(ns, operacao)
    st.markdown(f"""
    <div style="border-left: 8px solid {cor}; padding: 14px 16px; border-radius: 12px; background: #F8FAFC; border-top: 1px solid #E5E7EB; border-right: 1px solid #E5E7EB; border-bottom: 1px solid #E5E7EB;">
        <div style="font-size: 0.9rem; color: #475569; font-weight: 600;">{titulo}</div>
        <div style="font-size: 2rem; font-weight: 800; color: #0F172A;">{br_numero(ns, percentual=True)}</div>
        <div style="font-size: 0.85rem; color: #334155;">{faixa} · Pontuação {pontuacao}%</div>
        <div style="font-size: 0.85rem; color: #334155; margin-top: 6px;">Total: {int(total):,} | Dentro: {int(dentro):,} | Fora: {int(fora):,}</div>
    </div>
    """.replace(",", "."), unsafe_allow_html=True)


def meta_ns_operacao(operacao):
    return 0.96 if operacao == "1P MAGALU" else 0.95


def card_ns_executivo(operacao, ns, total, dentro, fora):
    meta = meta_ns_operacao(operacao)
    faltam = pedidos_para_meta(total, dentro, meta)
    faixa, cor_meta, pontuacao = classificar_ns(ns, operacao)
    cor_alerta = "#DC2626" if faltam > 0 or fora > 0 else "#16A34A"
    alerta = f"Faltam {faltam:,.0f} pedido(s) para atingir a meta de {meta*100:.0f}%" if faltam > 0 else "Meta atingida"
    st.markdown(f"""
    <div style="border-left: 8px solid {cor_meta}; padding: 14px 16px; border-radius: 12px; background: #F8FAFC; border: 1px solid #E5E7EB; margin-bottom: 10px;">
        <div style="font-size: 0.9rem; color: #475569; font-weight: 700;">{operacao}</div>
        <div style="font-size: 1.8rem; font-weight: 850; color: #0F172A;">{br_numero(ns, percentual=True)}</div>
        <div style="font-size: 0.85rem; color: #334155;">{faixa} · Pontuação {pontuacao}%</div>
        <div style="font-size: 0.85rem; color: #334155; margin-top: 6px;">Total: {int(total):,} | Dentro: {int(dentro):,} | <span style="color:#DC2626; font-weight:800;">Fora: {int(fora):,}</span></div>
        <div style="font-size: 0.9rem; color: {cor_alerta}; margin-top: 8px; font-weight: 850;">{alerta}</div>
    </div>
    """.replace(",", "."), unsafe_allow_html=True)


def formatar_ns_tabela(df):
    view = df.copy()
    if "Data Prevista" in view.columns:
        view["Data Prevista"] = pd.to_datetime(view["Data Prevista"], errors="coerce").dt.strftime("%d/%m/%Y")
    if "NS" in view.columns:
        view["NS"] = view["NS"].map(lambda x: br_numero(x, percentual=True) if pd.notna(x) else "")
    for col in ["Total_pedidos", "Dentro_prazo", "Fora_prazo"]:
        if col in view.columns:
            view[col] = view[col].map(lambda x: f"{int(x):,}".replace(",", ".") if pd.notna(x) else "")
    return view


def gerar_insights_ns(df):
    insights = []
    if df.empty:
        return pd.DataFrame(columns=["Prioridade", "Tipo", "Detalhe", "Insight", "Ação sugerida"])

    for operacao in sorted(df["Operação"].dropna().unique()):
        base_op = df[df["Operação"] == operacao].copy()
        limite = 0.95 if operacao == "1P MAGALU" else 0.94

        # Transportadores com maior impacto fora do prazo
        transp = consolidar_ns(base_op, ["Transportador"])
        transp = transp[transp["Total_pedidos"] >= 20].copy()
        transp_crit = transp[transp["NS"] < limite].sort_values(["Fora_prazo", "NS"], ascending=[False, True]).head(8)
        for _, r in transp_crit.iterrows():
            insights.append({
                "Prioridade": "Alta",
                "Tipo": "Transportador abaixo da meta",
                "Detalhe": r["Transportador"],
                "Insight": f"{operacao}: {r['Transportador']} está com NS de {br_numero(r['NS'], percentual=True)} e {int(r['Fora_prazo'])} pedido(s) fora do prazo.",
                "Ação sugerida": "Abrir detalhe por modal/cidade e acionar transportador com plano de correção."
            })

        # Recorrência: transportador abaixo da meta em vários dias
        dia_transp = consolidar_ns(base_op, ["Transportador", "Data Prevista"])
        dia_transp = dia_transp[dia_transp["Total_pedidos"] >= 5].copy()
        recorr = (dia_transp[dia_transp["NS"] < limite]
                  .groupby("Transportador")
                  .agg(Dias_abaixo=("Data Prevista", "nunique"), Pedidos_fora=("Fora_prazo", "sum"), NS_medio=("NS", "mean"))
                  .reset_index()
                  .sort_values(["Dias_abaixo", "Pedidos_fora"], ascending=False)
                  .head(8))
        for _, r in recorr.iterrows():
            if r["Dias_abaixo"] >= 2:
                insights.append({
                    "Prioridade": "Alta",
                    "Tipo": "Recorrência",
                    "Detalhe": r["Transportador"],
                    "Insight": f"{operacao}: {r['Transportador']} ficou abaixo da meta em {int(r['Dias_abaixo'])} dia(s), com NS médio de {br_numero(r['NS_medio'], percentual=True)}.",
                    "Ação sugerida": "Priorizar reunião operacional; recorrência indica problema estrutural, não evento isolado."
                })

        # Cidades ofensoras
        cidade = consolidar_ns(base_op, ["Cidade"])
        cidade = cidade[cidade["Total_pedidos"] >= 20]
        cidade = cidade[cidade["NS"] < limite].sort_values(["Fora_prazo", "NS"], ascending=[False, True]).head(8)
        for _, r in cidade.iterrows():
            insights.append({
                "Prioridade": "Média",
                "Tipo": "Cidade abaixo da meta",
                "Detalhe": r["Cidade"],
                "Insight": f"{operacao}: cidade {r['Cidade']} está com NS de {br_numero(r['NS'], percentual=True)} e {int(r['Fora_prazo'])} pedido(s) fora do prazo.",
                "Ação sugerida": "Verificar concentração por transportador/modal e janela de entrega."
            })

        # Modal abaixo da meta
        modal = consolidar_ns(base_op, ["Modal"])
        modal = modal[modal["Total_pedidos"] >= 20]
        modal = modal[modal["NS"] < limite].sort_values("NS").head(5)
        for _, r in modal.iterrows():
            insights.append({
                "Prioridade": "Alta",
                "Tipo": "Modal crítico",
                "Detalhe": r["Modal"],
                "Insight": f"{operacao}: modal {r['Modal']} está abaixo da meta com NS de {br_numero(r['NS'], percentual=True)}.",
                "Ação sugerida": "Analisar capacidade, grade de saída e transportadores do modal."
            })


    return pd.DataFrame(insights).drop_duplicates().head(40)

# -----------------------------
# Funções de Saída Eclusa
# -----------------------------
META_ECLUSA = 0.992
ZERADO_ECLUSA = 0.982


def pedidos_para_meta(total, dentro, meta):
    try:
        return max(0, int(math.ceil((meta * float(total)) - float(dentro))))
    except Exception:
        return 0


def saldo_pedidos_meta(total, dentro, meta):
    """Retorna o saldo de pedidos em relação à meta.
    Negativo = pedidos abaixo da meta | Positivo = pedidos acima da meta.
    """
    try:
        total = float(total or 0)
        dentro = float(dentro or 0)
        minimo_meta = int(math.ceil(meta * total))
        return int(dentro - minimo_meta)
    except Exception:
        return 0


def card_executivo_meta(titulo, ns, total, dentro, meta, cor_lateral=None):
    saldo = saldo_pedidos_meta(total, dentro, meta)

    if saldo >= 0:
        cor = "#16A34A"
        fundo = "#F0FDF4"
        borda = "#BBF7D0"
        icone = "⬆️"
        texto = f"{saldo:,.0f} pedido(s) acima da meta".replace(",", ".")
    else:
        cor = "#DC2626"
        fundo = "#FEF2F2"
        borda = "#FECACA"
        icone = "⬇️"
        texto = f"{abs(saldo):,.0f} pedido(s) abaixo da meta".replace(",", ".")

    lateral = cor_lateral or cor
    st.markdown(f"""
    <div style="border-left: 8px solid {lateral}; padding: 14px 16px; border-radius: 12px; background: #F8FAFC; border: 1px solid #E5E7EB; margin-bottom: 10px;">
        <div style="font-size: 0.9rem; color: #475569; font-weight: 800;">{titulo}</div>
        <div style="font-size: 1.8rem; font-weight: 850; color: #0F172A;">{br_numero(ns, percentual=True)}</div>
        <div style="margin-top: 10px; padding: 10px 12px; border-radius: 10px; background: {fundo}; border: 1px solid {borda}; color: {cor}; font-size: 0.95rem; font-weight: 900;">
            {icone} {texto}
        </div>
    </div>
    """, unsafe_allow_html=True)


def status_meta_eclusa(ns):
    if pd.isna(ns):
        return "Sem dado", "#9CA3AF", 0
    if ns < ZERADO_ECLUSA:
        return "Zerado", "#6B21A8", 0
    if ns < META_ECLUSA:
        return "80%", "#DC2626", 80
    if ns < 0.998:
        return "100%", "#16A34A", 100
    return "120%", "#2563EB", 120


def card_indicador_meta(titulo, ns, total, dentro, fora, meta, tipo="NS"):
    if tipo == "Eclusa":
        faixa, cor, pontuacao = status_meta_eclusa(ns)
    else:
        faixa, cor, pontuacao = classificar_ns(ns, titulo if titulo in ["1P MAGALU", "3P FULL"] else "1P MAGALU")
    faltam = pedidos_para_meta(total, dentro, meta)
    alerta = f"Faltam {faltam:,} pedido(s) para atingir a meta".replace(",", ".") if faltam > 0 else "Meta atingida"
    st.markdown(f"""
    <div style="border-left: 8px solid {cor}; padding: 14px 16px; border-radius: 12px; background: #F8FAFC; border: 1px solid #E5E7EB;">
        <div style="font-size: 0.9rem; color: #475569; font-weight: 700;">{titulo}</div>
        <div style="font-size: 2rem; font-weight: 850; color: #0F172A;">{br_numero(ns, percentual=True)}</div>
        <div style="font-size: 0.85rem; color: #334155;">{faixa} · Pontuação {pontuacao}%</div>
        <div style="font-size: 0.85rem; color: #334155; margin-top: 6px;">Total: {int(total):,} | Dentro: {int(dentro):,} | Fora: {int(fora):,}</div>
        <div style="font-size: 0.9rem; color: {cor}; margin-top: 8px; font-weight: 800;">{alerta}</div>
    </div>
    """.replace(",", "."), unsafe_allow_html=True)


@st.cache_data(show_spinner="Carregando saída eclusa...")
def carregar_eclusa_csv(arquivo):
    if arquivo is None:
        return pd.DataFrame()
    try:
        df = pd.read_csv(arquivo, sep=";", low_memory=False, encoding="utf-8-sig")
    except UnicodeDecodeError:
        df = pd.read_csv(arquivo, sep=";", low_memory=False, encoding="latin1")
    except Exception:
        try:
            df = pd.read_csv(arquivo, sep=";", low_memory=False, encoding="latin1")
        except Exception:
            df = pd.read_csv(arquivo, low_memory=False)

    df = limpar_nome_colunas(df)
    df = preencher_coluna_canonica(df, "Data Nota", candidatos=["data_nota", "Data Nota"], contem_todos=["data", "nota"])
    df = preencher_coluna_canonica(df, "Data Saída CD", candidatos=["data_saida_cd_considerada", "Data Saida CD Considerada", "data saída cd considerada"], contem_todos=["data", "saida", "cd"])
    df = preencher_coluna_canonica(df, "Data Limite Origem", candidatos=["Data_limite_origem", "data_limite_origem", "Data limite origem"], contem_todos=["data", "limite", "origem"])
    df = preencher_coluna_canonica(df, "Modal", candidatos=["Modal Transp", "Modal_Transp", "Modal"], contem_todos=["modal"])
    df = preencher_coluna_canonica(df, "Operação", candidatos=["Tipo Pedido", "tipo pedido", "tipo_malha (grupo)", "tipo_malha"], contem_todos=["tipo"])
    df = preencher_coluna_canonica(df, "Transportador Grupo", candidatos=["transportador (grupo)", "Transportador Grupo"], contem_todos=["transportador", "grupo"])
    df = preencher_coluna_canonica(df, "Transportador", candidatos=["transportador", "Transportador"], contem_todos=["transportador"])
    df = preencher_coluna_canonica(df, "Cidade", candidatos=["cidade cliente", "Cidade Cliente"], contem_todos=["cidade", "cliente"])
    df = preencher_coluna_canonica(df, "Pedido", candidatos=["pedido_pacote", "Pedido Pacote", "pedido"], contem_todos=["pedido"])
    df = preencher_coluna_canonica(df, "Dentro Prazo", candidatos=["NS_CD", "ns_cd", "NS CD"], contem_todos=["ns", "cd"])
    df = preencher_coluna_canonica(df, "Status Saída", candidatos=["Status_saida", "status_saida", "Status Saida"], contem_todos=["status", "saida"])

    cols = ["Data Nota", "Data Saída CD", "Data Limite Origem", "Modal", "Operação", "Transportador Grupo", "Transportador", "Cidade", "Pedido", "Dentro Prazo", "Status Saída"]
    for col in cols:
        if col not in df.columns:
            df[col] = pd.NA
    base = df[cols].copy()

    base["Data Nota"] = pd.to_datetime(base["Data Nota"], dayfirst=True, errors="coerce").dt.normalize()
    base["Data Saída CD"] = pd.to_datetime(base["Data Saída CD"], dayfirst=True, errors="coerce")
    base["Data Limite Origem"] = pd.to_datetime(base["Data Limite Origem"], dayfirst=True, errors="coerce")

    for col in ["Modal", "Operação", "Transportador Grupo", "Transportador", "Cidade", "Pedido", "Status Saída"]:
        base[col] = base[col].astype("string").str.strip().replace({"": pd.NA, "nan": pd.NA, "None": pd.NA, "<NA>": pd.NA}).fillna("Não informado")

    base["Dentro Prazo"] = pd.to_numeric(base["Dentro Prazo"], errors="coerce").fillna(0).apply(lambda x: 1 if x >= 1 else 0)
    base = base[base["Pedido"].astype(str).str.strip().ne("")].copy()

    base["Atraso horas"] = ((base["Data Saída CD"] - base["Data Limite Origem"]).dt.total_seconds() / 3600).fillna(0)
    base["Atraso horas"] = base["Atraso horas"].clip(lower=0)
    base["Faixa Atraso"] = base["Atraso horas"].apply(lambda x: "No prazo" if x <= 0 else ("Até 2h" if x <= 2 else ("Até 6h" if x <= 6 else ("Até 12h" if x <= 12 else "Acima de 12h"))))
    base["Mês"] = base["Data Nota"].apply(nome_mes_pt)
    base["Início Semana"] = base["Data Nota"] - pd.to_timedelta(base["Data Nota"].dt.weekday, unit="D")
    base["Semana"] = base["Início Semana"].dt.strftime("%d/%m/%Y")
    base["Semana Mês"] = base["Início Semana"].dt.strftime("%d/%m")
    base["Fora Prazo"] = 1 - base["Dentro Prazo"]
    base["Total_pedidos"] = 1

    # Deduplicação por pedido, mantendo o pior cenário de NS e maior atraso.
    base = (base.groupby("Pedido", as_index=False, dropna=False)
            .agg({
                "Data Nota": primeiro_valido,
                "Data Saída CD": primeiro_valido,
                "Data Limite Origem": primeiro_valido,
                "Modal": primeiro_valido,
                "Operação": primeiro_valido,
                "Transportador Grupo": primeiro_valido,
                "Transportador": primeiro_valido,
                "Cidade": primeiro_valido,
                "Status Saída": primeiro_valido,
                "Dentro Prazo": "max",
                "Atraso horas": "max",
                "Faixa Atraso": primeiro_valido,
                "Mês": primeiro_valido,
                "Início Semana": primeiro_valido,
                "Semana": primeiro_valido,
                "Semana Mês": primeiro_valido,
                "Total_pedidos": "max"
            }))
    base["Fora Prazo"] = 1 - base["Dentro Prazo"]
    base["NS"] = base["Dentro Prazo"]
    return base


def consolidar_eclusa(df, dimensoes):
    if df.empty:
        return pd.DataFrame()
    if isinstance(dimensoes, str):
        dimensoes = [dimensoes]
    base = df.copy()
    resumo = (base.groupby(dimensoes, dropna=False)
              .agg(Total_pedidos=("Total_pedidos", "sum"),
                   Dentro_prazo=("Dentro Prazo", "sum"),
                   Fora_prazo=("Fora Prazo", "sum"),
                   Atraso_medio_h=("Atraso horas", lambda x: x[x > 0].mean() if (x > 0).any() else 0),
                   Atraso_max_h=("Atraso horas", "max"))
              .reset_index())
    resumo["NS"] = resumo["Dentro_prazo"] / resumo["Total_pedidos"].replace(0, np.nan)
    resumo["Faltam_meta"] = resumo.apply(lambda r: pedidos_para_meta(r["Total_pedidos"], r["Dentro_prazo"], META_ECLUSA), axis=1)
    return resumo


def formatar_eclusa_tabela(df):
    view = df.copy()
    if "Data Nota" in view.columns:
        view["Data Nota"] = pd.to_datetime(view["Data Nota"], errors="coerce").dt.strftime("%d/%m/%Y")
    if "NS" in view.columns:
        view["NS"] = view["NS"].map(lambda x: br_numero(x, percentual=True) if pd.notna(x) else "")
    for col in ["Total_pedidos", "Dentro_prazo", "Fora_prazo", "Faltam_meta"]:
        if col in view.columns:
            view[col] = view[col].map(lambda x: f"{int(x):,}".replace(",", ".") if pd.notna(x) else "")
    for col in ["Atraso_medio_h", "Atraso_max_h", "Atraso horas"]:
        if col in view.columns:
            view[col] = view[col].map(lambda x: f"{x:.1f}h" if pd.notna(x) else "")
    return view


def gerar_insights_eclusa(df):
    insights = []
    if df.empty:
        return pd.DataFrame(columns=["Prioridade", "Tipo", "Detalhe", "Insight", "Ação sugerida"])

    # Transportador grupo abaixo da meta com maior impacto
    transp = consolidar_eclusa(df, ["Transportador Grupo"])
    transp = transp[transp["Total_pedidos"] >= 20]
    crit = transp[transp["NS"] < META_ECLUSA].sort_values(["Fora_prazo", "NS"], ascending=[False, True]).head(8)
    for _, r in crit.iterrows():
        insights.append({
            "Prioridade": "Alta",
            "Tipo": "Transportador abaixo da meta",
            "Detalhe": r["Transportador Grupo"],
            "Insight": f"{r['Transportador Grupo']} está com saída eclusa em {br_numero(r['NS'], percentual=True)}, com {int(r['Fora_prazo'])} pedido(s) fora e faltam {int(r['Faltam_meta'])} pedido(s) para a meta.",
            "Ação sugerida": "Abrir detalhe por modal/cidade e tratar causa de atraso com o transportador."
        })

    # Recorrência por transportador/semana
    sem_transp = consolidar_eclusa(df, ["Transportador Grupo", "Semana"])
    sem_transp = sem_transp[sem_transp["Total_pedidos"] >= 10]
    recorr = (sem_transp[sem_transp["NS"] < META_ECLUSA]
              .groupby("Transportador Grupo")
              .agg(Semanas_abaixo=("Semana", "nunique"), Fora=("Fora_prazo", "sum"), NS_medio=("NS", "mean"))
              .reset_index()
              .sort_values(["Semanas_abaixo", "Fora"], ascending=False)
              .head(8))
    for _, r in recorr.iterrows():
        if r["Semanas_abaixo"] >= 2:
            insights.append({
                "Prioridade": "Alta",
                "Tipo": "Recorrência semanal",
                "Detalhe": r["Transportador Grupo"],
                "Insight": f"{r['Transportador Grupo']} ficou abaixo da meta em {int(r['Semanas_abaixo'])} semana(s), com NS médio de {br_numero(r['NS_medio'], percentual=True)}.",
                "Ação sugerida": "Tratar como ponto recorrente; avaliar janela, grade, capacidade e rotina da eclusa."
            })

    # Cidades ofensoras
    cidade = consolidar_eclusa(df, ["Cidade"])
    cidade = cidade[cidade["Total_pedidos"] >= 20]
    cidade = cidade[cidade["NS"] < META_ECLUSA].sort_values(["Fora_prazo", "NS"], ascending=[False, True]).head(8)
    for _, r in cidade.iterrows():
        insights.append({
            "Prioridade": "Média",
            "Tipo": "Cidade abaixo da meta",
            "Detalhe": r["Cidade"],
            "Insight": f"Cidade {r['Cidade']} está com eclusa em {br_numero(r['NS'], percentual=True)} e {int(r['Fora_prazo'])} pedido(s) fora.",
            "Ação sugerida": "Verificar concentração por transportador/modal e horário de saída."
        })

    # Atraso médio
    atraso = df[df["Atraso horas"] > 0]
    if not atraso.empty:
        atraso_transp = atraso.groupby("Transportador Grupo")["Atraso horas"].mean().sort_values(ascending=False).head(5)
        for nome, val in atraso_transp.items():
            if val >= 1:
                insights.append({
                    "Prioridade": "Média",
                    "Tipo": "Atraso médio elevado",
                    "Detalhe": nome,
                    "Insight": f"{nome} apresenta atraso médio de {val:.1f}h nas saídas fora do prazo.",
                    "Ação sugerida": "Avaliar etapa anterior à saída e tempo de permanência na eclusa."
                })
    return pd.DataFrame(insights).drop_duplicates().head(40)



def preparar_eclusa_para_historico(ecl):
    """Salva eclusa agregada para não pesar o Google Sheets."""
    if ecl.empty:
        return pd.DataFrame(columns=COLUNAS_ECLUSA_HISTORICO)
    df = ecl.copy()
    for col in ["Operação", "Modal", "Transportador Grupo", "Cidade", "Status Saída", "Faixa Atraso"]:
        if col not in df.columns:
            df[col] = "Não informado"
        df[col] = df[col].astype("string").str.strip().replace({"": "Não informado", "nan": "Não informado", "None": "Não informado", "<NA>": "Não informado"}).fillna("Não informado")
    df["Data Nota"] = pd.to_datetime(df.get("Data Nota"), errors="coerce").dt.normalize()
    if "Início Semana" not in df.columns:
        df["Início Semana"] = df["Data Nota"] - pd.to_timedelta(df["Data Nota"].dt.weekday, unit="D")
    df["Início Semana"] = pd.to_datetime(df["Início Semana"], errors="coerce").dt.normalize()
    if "Semana" not in df.columns:
        df["Semana"] = df["Início Semana"].dt.strftime("%d/%m/%Y")
    df["Total_pedidos"] = pd.to_numeric(df.get("Total_pedidos", 1), errors="coerce").fillna(1)
    df["Dentro Prazo"] = pd.to_numeric(df.get("Dentro Prazo", 0), errors="coerce").fillna(0)
    df["Fora Prazo"] = pd.to_numeric(df.get("Fora Prazo", df["Total_pedidos"] - df["Dentro Prazo"]), errors="coerce").fillna(0)
    df["Atraso horas"] = pd.to_numeric(df.get("Atraso horas", 0), errors="coerce").fillna(0)
    dims = ["Operação", "Data Nota", "Início Semana", "Semana", "Modal", "Transportador Grupo", "Cidade", "Status Saída", "Faixa Atraso"]
    resumo = (df.groupby(dims, dropna=False)
              .agg(**{
                  "Total Pedidos": ("Total_pedidos", "sum"),
                  "Dentro Prazo": ("Dentro Prazo", "sum"),
                  "Fora Prazo": ("Fora Prazo", "sum"),
                  "Atraso Médio h": ("Atraso horas", lambda x: x[x > 0].mean() if (x > 0).any() else 0),
                  "Maior Atraso h": ("Atraso horas", "max"),
              })
              .reset_index())
    resumo["NS"] = resumo["Dentro Prazo"] / resumo["Total Pedidos"].replace(0, np.nan)
    resumo["Faltam Meta"] = resumo.apply(lambda r: pedidos_para_meta(r["Total Pedidos"], r["Dentro Prazo"], META_ECLUSA), axis=1)
    resumo["Data Nota"] = pd.to_datetime(resumo["Data Nota"], errors="coerce").dt.strftime("%d/%m/%Y")
    resumo["Início Semana"] = pd.to_datetime(resumo["Início Semana"], errors="coerce").dt.strftime("%d/%m/%Y")
    return resumo[COLUNAS_ECLUSA_HISTORICO]


@st.cache_data(ttl=300, show_spinner=False)
def ler_historico_eclusa(sheet_id):
    if not sheet_id:
        return pd.DataFrame(columns=COLUNAS_ECLUSA_HISTORICO)
    ws = abrir_worksheet_generico(sheet_id, ABA_ECLUSA, COLUNAS_ECLUSA_HISTORICO, linhas=1000)
    registros = ws.get_all_records()
    if not registros:
        return pd.DataFrame(columns=COLUNAS_ECLUSA_HISTORICO)
    return pd.DataFrame(registros)


def escrever_historico_eclusa(sheet_id, df_hist):
    ws = abrir_worksheet_generico(sheet_id, ABA_ECLUSA, COLUNAS_ECLUSA_HISTORICO, linhas=max(len(df_hist) + 10, 1000))
    df = df_hist.copy()
    for col in COLUNAS_ECLUSA_HISTORICO:
        if col not in df.columns:
            df[col] = ""
    df = df[COLUNAS_ECLUSA_HISTORICO].fillna("")
    ws.clear()
    ws.update("A1", [COLUNAS_ECLUSA_HISTORICO] + df.astype(str).values.tolist())


def historico_eclusa_para_interno(hist):
    if hist.empty:
        return pd.DataFrame()
    h = limpar_nome_colunas(hist)
    for col in COLUNAS_ECLUSA_HISTORICO:
        if col not in h.columns:
            h[col] = ""
    df = pd.DataFrame()
    df["Operação"] = h["Operação"].astype("string").str.strip().replace({"": "Não informado"}).fillna("Não informado")
    df["Data Nota"] = pd.to_datetime(h["Data Nota"], dayfirst=True, errors="coerce").dt.normalize()
    df["Início Semana"] = pd.to_datetime(h["Início Semana"], dayfirst=True, errors="coerce").dt.normalize()
    df["Semana"] = h["Semana"].astype("string").str.strip()
    df["Semana Mês"] = df["Início Semana"].dt.strftime("%d/%m")
    df["Modal"] = h["Modal"].astype("string").str.strip().replace({"": "Não informado"}).fillna("Não informado")
    df["Transportador Grupo"] = h["Transportador Grupo"].astype("string").str.strip().replace({"": "Não informado"}).fillna("Não informado")
    df["Transportador"] = df["Transportador Grupo"]
    df["Cidade"] = h["Cidade"].astype("string").str.strip().replace({"": "Não informado"}).fillna("Não informado")
    df["Status Saída"] = h["Status Saída"].astype("string").str.strip().replace({"": "Não informado"}).fillna("Não informado")
    df["Faixa Atraso"] = h["Faixa Atraso"].astype("string").str.strip().replace({"": "Não informado"}).fillna("Não informado")
    df["Total_pedidos"] = pd.to_numeric(h["Total Pedidos"], errors="coerce").fillna(0)
    df["Dentro Prazo"] = pd.to_numeric(h["Dentro Prazo"], errors="coerce").fillna(0)
    df["Fora Prazo"] = pd.to_numeric(h["Fora Prazo"], errors="coerce").fillna(0)
    df["NS"] = pd.to_numeric(h["NS"], errors="coerce")
    df["Atraso horas"] = pd.to_numeric(h["Atraso Médio h"], errors="coerce").fillna(0)
    df["Mês"] = df["Data Nota"].apply(nome_mes_pt)
    df["Pedido"] = "Histórico agregado"
    return df


def juntar_historico_substituindo_datas(hist_existente, novo, data_col, colunas):
    if novo.empty:
        return hist_existente.copy() if not hist_existente.empty else pd.DataFrame(columns=colunas)
    hist = hist_existente.copy() if not hist_existente.empty else pd.DataFrame(columns=colunas)
    novo = novo.copy()
    for col in colunas:
        if col not in hist.columns:
            hist[col] = ""
        if col not in novo.columns:
            novo[col] = ""
    datas_novas = set(novo[data_col].astype(str).str.strip().dropna().tolist())
    if datas_novas and not hist.empty:
        hist = hist[~hist[data_col].astype(str).str.strip().isin(datas_novas)].copy()
    combinado = pd.concat([hist[colunas], novo[colunas]], ignore_index=True)
    return combinado.drop_duplicates(keep="last")[colunas]

# -----------------------------
# Interface
# -----------------------------
st.title("📊 Painel Transporte CD 2900")
st.caption("Cubagem + Nível de Serviço 1P/Full + Saída Eclusa, com histórico otimizado no Google Sheets e insights automáticos.")

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
        arquivo = None
        st.info("Nenhuma base local enviada. Usando histórico do Google Sheets.")


with st.sidebar:
    st.header("Bases NS 1P / Full")
    upload_1p = st.file_uploader("Base Nível de Serviço 1P (.csv)", type=["csv"], key="upload_ns_1p")
    upload_full = st.file_uploader("Base Nível de Serviço Full (.csv)", type=["csv"], key="upload_ns_full")

    arquivo_1p = upload_1p if upload_1p is not None else localizar_arquivo_csv(["1P.csv", "*1P*.csv"])
    arquivo_full = upload_full if upload_full is not None else localizar_arquivo_csv(["Full.csv", "*Full*.csv"])

with st.sidebar:
    st.header("Base Saída Eclusa")
    upload_eclusa = st.file_uploader("Base Saída Eclusa (.csv)", type=["csv"], key="upload_eclusa")
    arquivo_eclusa = upload_eclusa if upload_eclusa is not None else localizar_arquivo_csv(["Saida eclusa.csv", "Saida CD.csv", "*eclusa*.csv", "*Saida*.csv", "*saida*.csv"])

if arquivo is not None:
    try:
        base_original, cargas_atual = carregar_dados_excel(arquivo)
    except FileNotFoundError:
        base_original = pd.DataFrame()
        cargas_atual = pd.DataFrame()
else:
    base_original = pd.DataFrame()
    cargas_atual = pd.DataFrame()


try:
    ns_1p_atual = carregar_ns_csv(arquivo_1p, "1P") if arquivo_1p is not None else pd.DataFrame()
    ns_full_atual = carregar_ns_csv(arquivo_full, "Full") if arquivo_full is not None else pd.DataFrame()
except Exception as e:
    st.sidebar.error(f"Erro ao carregar bases de NS: {e}")
    ns_1p_atual = pd.DataFrame()
    ns_full_atual = pd.DataFrame()

try:
    eclusa_atual = carregar_eclusa_csv(arquivo_eclusa) if arquivo_eclusa is not None else pd.DataFrame()
except Exception as e:
    st.sidebar.error(f"Erro ao carregar base de saída eclusa: {e}")
    eclusa_atual = pd.DataFrame()

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

historico_ns_1p_raw = pd.DataFrame(columns=COLUNAS_NS_HISTORICO)
historico_ns_full_raw = pd.DataFrame(columns=COLUNAS_NS_HISTORICO)
historico_ns_1p = pd.DataFrame(columns=COLUNAS_NS_HISTORICO)
historico_ns_full = pd.DataFrame(columns=COLUNAS_NS_HISTORICO)
if sheet_id and usar_historico:
    try:
        historico_ns_1p_raw = ler_historico_ns(sheet_id, "1P")
        historico_ns_full_raw = ler_historico_ns(sheet_id, "Full")
        historico_ns_1p = historico_ns_para_interno(historico_ns_1p_raw)
        historico_ns_full = historico_ns_para_interno(historico_ns_full_raw)
        total_ns_hist = len(historico_ns_1p) + len(historico_ns_full)
        if total_ns_hist:
            st.sidebar.success(f"Histórico NS conectado: {total_ns_hist:,.0f} linha(s) agregada(s).".replace(",", "."))
    except Exception as e:
        st.sidebar.warning(f"Não foi possível ler histórico NS: {e}")

historico_eclusa_raw = pd.DataFrame(columns=COLUNAS_ECLUSA_HISTORICO)
historico_eclusa = pd.DataFrame()
if sheet_id and usar_historico:
    try:
        historico_eclusa_raw = ler_historico_eclusa(sheet_id)
        historico_eclusa = historico_eclusa_para_interno(historico_eclusa_raw)
        if not historico_eclusa.empty:
            st.sidebar.success(f"Histórico Eclusa conectado: {len(historico_eclusa):,.0f} linha(s) agregada(s).".replace(",", "."))
    except Exception as e:
        st.sidebar.warning(f"Não foi possível ler histórico Eclusa: {e}")

if eclusa_atual.empty and not historico_eclusa.empty:
    eclusa_atual = historico_eclusa.copy()

if not historico_ns_1p.empty and not ns_1p_atual.empty:
    ns_1p = pd.concat([historico_ns_1p, ns_1p_atual], ignore_index=True).drop_duplicates(subset=["Operação", "Pedido"], keep="last")
elif not ns_1p_atual.empty:
    ns_1p = ns_1p_atual.copy()
else:
    ns_1p = historico_ns_1p.copy()

if not historico_ns_full.empty and not ns_full_atual.empty:
    ns_full = pd.concat([historico_ns_full, ns_full_atual], ignore_index=True).drop_duplicates(subset=["Operação", "Pedido"], keep="last")
elif not ns_full_atual.empty:
    ns_full = ns_full_atual.copy()
else:
    ns_full = historico_ns_full.copy()

base_ns = pd.concat([ns_1p, ns_full], ignore_index=True) if (not ns_1p.empty or not ns_full.empty) else pd.DataFrame()

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
elif fonte == "Atual + Histórico" and not historico_interno.empty and not cargas_atual.empty:
    cargas = pd.concat([historico_interno, cargas_atual], ignore_index=True)
    cargas = cargas.drop_duplicates(subset=["ID de carga"], keep="last")
elif not cargas_atual.empty:
    cargas = cargas_atual.copy()
elif not historico_interno.empty:
    cargas = historico_interno.copy()
else:
    st.warning("Nenhuma base disponível. Verifique o histórico do Google Sheets ou envie uma planilha.")
    st.stop()

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
    f_tipo_pedido = multiselect_coluna("Tipo de Pedido")
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
    ("Linha de produção", f_linha), ("Tipo de Pedido", f_tipo_pedido), ("Destino Filial", f_destino)
]:
    if filtro and col in filtrado.columns:
        filtrado = filtrado[filtrado[col].astype(str).isin(filtro)]

if f_data and "Data Carregamento" in filtrado.columns:
    if isinstance(f_data, tuple) and len(f_data) == 2:
        dt_ini, dt_fim = pd.to_datetime(f_data[0]), pd.to_datetime(f_data[1])
        filtrado = filtrado[(filtrado["Data Carregamento"] >= dt_ini) & (filtrado["Data Carregamento"] <= dt_fim)]

filtrado = filtrado[(filtrado["Ocupação %"].fillna(0) * 100 >= ocup_min) & (filtrado["Ocupação %"].fillna(0) * 100 <= ocup_max)]
simulado = filtrado[filtrado["Ocupação %"].fillna(0) * 100 >= limite_simulacao].copy()



def calcular_ns_geral_por_operacao(base, operacao):
    if base.empty or "Operação" not in base.columns:
        return np.nan
    b = base[base["Operação"] == operacao].copy()
    if b.empty:
        return np.nan
    total = pd.to_numeric(b.get("Total_pedidos", pd.Series(dtype=float)), errors="coerce").fillna(1).sum()
    dentro = pd.to_numeric(b.get("Dentro Prazo", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()
    return dentro / total if total else np.nan


def render_resumo_compacto_superior():
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Cubagem geral", br_numero(filtrado["Ocupação %"].mean(), percentual=True))
    c2.metric("NS 1P geral", br_numero(calcular_ns_geral_por_operacao(base_ns, "1P MAGALU"), percentual=True))
    c3.metric("NS Full geral", br_numero(calcular_ns_geral_por_operacao(base_ns, "3P FULL"), percentual=True))
    if not eclusa_atual.empty:
        total_e = pd.to_numeric(eclusa_atual.get("Total_pedidos", pd.Series(dtype=float)), errors="coerce").fillna(1).sum()
        dentro_e = pd.to_numeric(eclusa_atual.get("Dentro Prazo", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()
        ns_e = dentro_e / total_e if total_e else np.nan
    else:
        ns_e = np.nan
    c4.metric("Saída Eclusa geral", br_numero(ns_e, percentual=True))
    st.divider()

aba_exec, aba_cubagem, aba_ns, aba_eclusa, aba_bases = st.tabs([
    "Visão Executiva", "Cubagem", "Nível de Serviço", "Saída Eclusa", "Bases / Histórico"
])

with aba_exec:
    st.subheader("Visão Executiva")
    st.caption("Resumo rápido dos principais indicadores do CD 2900.")

    c1, c2 = st.columns([1, 2])
    with c1:
        st.markdown("### Cubagem")
        st.metric("Ocupação média", br_numero(filtrado["Ocupação %"].mean(), percentual=True))
        st.metric("Cargas abaixo de 50%", f"{(filtrado['Ocupação %'] < 0.5).sum():,.0f}".replace(",", "."))
        st.metric("Sobra total", f"{filtrado['Sobra m³'].clip(lower=0).sum():,.0f} m³".replace(",", "."))

    with c2:
        st.markdown("### Nível de Serviço")
        if base_ns.empty:
            st.info("Envie bases NS 1P/Full ou salve histórico de NS para visualizar.")
        else:
            ns_exec = base_ns.copy()
            for operacao in ["1P MAGALU", "3P FULL"]:
                base_op = ns_exec[ns_exec["Operação"] == operacao]
                total = pd.to_numeric(base_op.get("Total_pedidos", pd.Series(dtype=float)), errors="coerce").fillna(1).sum()
                dentro = pd.to_numeric(base_op.get("Dentro Prazo", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()
                ns_val = dentro / total if total else np.nan
                faixa, cor_meta, _ = classificar_ns(ns_val, operacao)
                card_executivo_meta(operacao, ns_val, total, dentro, meta_ns_operacao(operacao), cor_lateral=cor_meta)

                modal_op = consolidar_ns(base_op, ["Modal"])
                modal_op = modal_op[modal_op["Modal"].astype(str).str.upper().isin(["MICRO", "RODO", "COURIER"])]
                if not modal_op.empty:
                    st.caption(f"{operacao} por modal")
                    cols_modal = st.columns(min(3, len(modal_op)))
                    for idx, (_, r) in enumerate(modal_op.sort_values("Modal").iterrows()):
                        with cols_modal[idx % len(cols_modal)]:
                            ns_modal = r["Dentro_prazo"] / r["Total_pedidos"] if r["Total_pedidos"] else np.nan
                            faixa_m, cor_m, _ = classificar_ns(ns_modal, operacao)
                            card_executivo_meta(str(r["Modal"]), ns_modal, r["Total_pedidos"], r["Dentro_prazo"], meta_ns_operacao(operacao), cor_lateral=cor_m)

    st.divider()
    st.markdown("### Saída Eclusa")
    if eclusa_atual.empty:
        st.info("Envie a base de Saída Eclusa no menu lateral para visualizar o indicador.")
    else:
        total_e = eclusa_atual["Total_pedidos"].sum()
        dentro_e = eclusa_atual["Dentro Prazo"].sum()
        ns_e = dentro_e / total_e if total_e else np.nan
        atraso_medio = eclusa_atual.loc[eclusa_atual["Atraso horas"] > 0, "Atraso horas"].mean()
        faixa_e, cor_e, _ = status_meta_eclusa(ns_e)

        e1, e2 = st.columns([1, 1])
        with e1:
            card_executivo_meta("Saída Eclusa Geral", ns_e, total_e, dentro_e, META_ECLUSA, cor_lateral=cor_e)
        with e2:
            st.metric("Atraso médio", f"{atraso_medio:.1f}h" if pd.notna(atraso_medio) else "0,0h")

        eclusa_modal = consolidar_eclusa(eclusa_atual, ["Modal"])
        eclusa_modal = eclusa_modal[eclusa_modal["Modal"].astype(str).str.upper().isin(["MICRO", "RODO", "COURIER"])]
        if not eclusa_modal.empty:
            st.caption("Saída Eclusa por modal")
            cols_eclusa = st.columns(min(3, len(eclusa_modal)))
            for idx, (_, r) in enumerate(eclusa_modal.sort_values("Modal").iterrows()):
                with cols_eclusa[idx % len(cols_eclusa)]:
                    ns_modal = r["Dentro_prazo"] / r["Total_pedidos"] if r["Total_pedidos"] else np.nan
                    faixa_m, cor_m, _ = status_meta_eclusa(ns_modal)
                    card_executivo_meta(str(r["Modal"]), ns_modal, r["Total_pedidos"], r["Dentro_prazo"], META_ECLUSA, cor_lateral=cor_m)

    st.divider()
    col_sug1, col_sug2 = st.columns(2)
    with col_sug1:
        st.markdown("#### Sugestões de Cubagem")
        sugestoes = gerar_sugestoes(filtrado, limite_recorrente=0.40, limite_ofensor=limite_simulacao / 100)
        if sugestoes.empty:
            st.success("Nenhum ponto crítico recorrente encontrado na cubagem.")
        else:
            st.dataframe(sugestoes, use_container_width=True, height=360)

    with col_sug2:
        st.markdown("#### Insights de NS")
        if base_ns.empty:
            st.info("Sem base de NS carregada/histórica.")
        else:
            insights_ns = gerar_insights_ns(base_ns)
            if insights_ns.empty:
                st.success("Nenhum insight crítico encontrado em NS.")
            else:
                st.dataframe(insights_ns, use_container_width=True, height=360)

with aba_cubagem:
    st.subheader("Cubagem")
    st.caption("Aba consolidada com KPIs, gráficos, ofensores, simulação e cargas especiais.")

    col1, col2, col3, col4, col5, col6 = st.columns(6)
    col1.metric("Cargas", f"{filtrado['ID de carga'].nunique():,.0f}".replace(",", "."))
    col2.metric("Ocupação média", br_numero(filtrado["Ocupação %"].mean(), percentual=True))
    col3.metric("Cubagem total", f"{filtrado['Cubagem da carga'].sum():,.0f} m³".replace(",", "."))
    col4.metric("Capacidade total", f"{filtrado['Capacidade do veículo'].sum():,.0f} m³".replace(",", "."))
    col5.metric("Sobra total", f"{filtrado['Sobra m³'].clip(lower=0).sum():,.0f} m³".replace(",", "."))
    col6.metric("Cargas <50%", f"{(filtrado['Ocupação %'] < 0.5).sum():,.0f}".replace(",", "."))
    st.divider()

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

    st.markdown("### Plano de transporte e dia a dia")
    rota_col = "Plano de transporte" if "Plano de transporte" in filtrado.columns else "Destino Filial"
    col_rota, col_dia = st.columns(2)
    with col_rota:
        destino = resumo_por(filtrado, rota_col)
        if not destino.empty:
            fig_rota = px.bar(destino.head(20), x=rota_col, y="Sobra", text="Sobra", title="Top planos com maior sobra de cubagem")
            aplicar_rotulos_barra(fig_rota, formato="%{y:,.0f}")
            st.plotly_chart(fig_rota, use_container_width=True)

    with col_dia:
        if "Data Carregamento" not in filtrado.columns or filtrado["Data Carregamento"].dropna().empty:
            st.info("Não encontrei data de carregamento válida.")
        else:
            base_dia = filtrado.dropna(subset=["Data Carregamento"]).copy()
            resumo_dia = (base_dia.groupby("Data Carregamento", dropna=False)
                .agg(Cargas=("ID de carga", "nunique"), Cubagem=("Cubagem da carga", "sum"), Capacidade=("Capacidade do veículo", "sum"), Sobra=("Sobra m³", "sum"))
                .reset_index().sort_values("Data Carregamento"))
            resumo_dia["Ocupacao_dia"] = resumo_dia["Cubagem"] / resumo_dia["Capacidade"]
            resumo_dia["Data_formatada"] = pd.to_datetime(resumo_dia["Data Carregamento"], errors="coerce").dt.strftime("%d/%m/%Y")
            fig_dia = make_subplots(specs=[[{"secondary_y": True}]])
            fig_dia.add_trace(go.Bar(x=resumo_dia["Data_formatada"], y=resumo_dia["Cubagem"], name="Cubagem diária (m³)", text=resumo_dia["Cubagem"].round(0), textposition="outside", texttemplate="%{text:,.0f}"), secondary_y=False)
            fig_dia.add_trace(go.Scatter(x=resumo_dia["Data_formatada"], y=resumo_dia["Ocupacao_dia"], name="Ocupação", mode="lines+markers+text", text=[br_numero(x, percentual=True) for x in resumo_dia["Ocupacao_dia"]], textposition="top center"), secondary_y=True)
            fig_dia.update_layout(height=420, margin=dict(t=40, b=20), legend=dict(orientation="h"), title="Cubagem diária x ocupação")
            fig_dia.update_yaxes(title_text="Cubagem diária (m³)", secondary_y=False)
            fig_dia.update_yaxes(title_text="Ocupação", tickformat=".0%", secondary_y=True)
            st.plotly_chart(fig_dia, use_container_width=True)

    st.markdown("### Ofensores e melhores cargas")
    col_of, col_mel = st.columns(2)
    cols_base = [c for c in ["ID de carga", "Data Carregamento", "Plano de transporte", "Transportador", "Tipo de Pedido", "Tipo do veículo", "Modalidade", "Linha de produção", "Destino Filial", "Cubagem da carga", "Capacidade do veículo", "Ocupação %", "Sobra m³", "Valor da carga"] if c in filtrado.columns]
    with col_of:
        st.markdown("#### 10 maiores ofensores")
        ofensores = filtrado.sort_values(["Ocupação %", "Sobra m³"], ascending=[True, False]).head(10)
        st.dataframe(preparar_visualizacao(ofensores[cols_base]), use_container_width=True, height=360)
    with col_mel:
        st.markdown("#### 10 melhores performances")
        melhores = filtrado[filtrado["Ocupação %"] <= 1].sort_values("Ocupação %", ascending=False).head(10)
        if melhores.empty:
            melhores = filtrado.sort_values("Ocupação %", ascending=False).head(10)
        st.dataframe(preparar_visualizacao(melhores[cols_base]), use_container_width=True, height=360)

    st.markdown("### Simulação")
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

    with st.expander("Ver base crítica completa"):
        criticas = filtrado[filtrado["Ocupação %"] < 0.5].sort_values("Ocupação %")
        st.dataframe(preparar_visualizacao(criticas[cols_base]), use_container_width=True, height=520)
        csv = preparar_visualizacao(criticas[cols_base]).to_csv(index=False, sep=";").encode("utf-8-sig")
        st.download_button("Baixar cargas críticas em CSV", csv, "cargas_criticas.csv", "text/csv")

    with st.expander("Cargas especiais cadastradas"):
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
            cols_esp = [c for c in ["ID de carga", "Tipo especial", "Data Carregamento", "Plano de transporte", "Transportador", "Tipo de Pedido", "Tipo do veículo", "Modalidade", "Linha de produção", "Destino Filial", "Cubagem da carga", "Capacidade do veículo", "Ocupação %", "Sobra m³", "Valor da carga"] if c in base_esp.columns]
            st.dataframe(preparar_visualizacao(base_esp[cols_esp]), use_container_width=True, height=420)

with aba_ns:
    render_resumo_compacto_superior()
    st.subheader("Nível de Serviço 1P / Full")
    st.caption("Baseado em data entrega prevista cliente, Modal Transp, cidade cliente, transportador, ECC/Ecom, situação, prazo_cliente e pedido_pacote.")

    if base_ns.empty:
        st.info("Envie as bases CSV de Nível de Serviço 1P e/ou Full no menu lateral para iniciar a análise.")
    else:
        ns = base_ns.copy()

        f1, f2, f3, f4 = st.columns(4)
        operacoes = sorted(ns["Operação"].dropna().unique().tolist())
        op_sel = f1.multiselect("Operação", operacoes, default=operacoes, key="ns_operacao")
        modal_sel = f2.multiselect("Modal", sorted(ns["Modal"].dropna().astype(str).unique().tolist()), key="ns_modal")
        transp_sel = f3.multiselect("Transportador", sorted(ns["Transportador"].dropna().astype(str).unique().tolist()), key="ns_transportador")
        cidade_sel = f4.multiselect("Cidade", sorted(ns["Cidade"].dropna().astype(str).unique().tolist()), key="ns_cidade")

        f5, f6, f7, f8 = st.columns(4)
        canal_sel = f5.multiselect("ECC / Ecom em casa", sorted(ns["Canal"].dropna().astype(str).unique().tolist()), key="ns_canal") if "Canal" in ns.columns else []
        situacao_sel = f6.multiselect("Situação", sorted(ns["Situação"].dropna().astype(str).unique().tolist()), key="ns_situacao") if "Situação" in ns.columns else []
        prazo_sel = f7.multiselect("Prazo Cliente", sorted(ns["Faixa Prazo"].dropna().astype(str).unique().tolist()), key="ns_prazo") if "Faixa Prazo" in ns.columns else []
        ofensor_sel = f8.multiselect("Ofensor NS", sorted(ns["Ofensor NS"].dropna().astype(str).unique().tolist()), key="ns_ofensor") if "Ofensor NS" in ns.columns else []

        if ns["Data Prevista"].notna().any():
            dmin = ns["Data Prevista"].dropna().min().date()
            dmax = ns["Data Prevista"].dropna().max().date()
            periodo_ns = st.date_input("Período - Data entrega prevista cliente", value=(dmin, dmax), min_value=dmin, max_value=dmax, format="DD/MM/YYYY", key="periodo_ns")
        else:
            periodo_ns = None

        ns_f = ns.copy()
        if op_sel:
            ns_f = ns_f[ns_f["Operação"].isin(op_sel)]
        if modal_sel:
            ns_f = ns_f[ns_f["Modal"].astype(str).isin(modal_sel)]
        if transp_sel:
            ns_f = ns_f[ns_f["Transportador"].astype(str).isin(transp_sel)]
        if cidade_sel:
            ns_f = ns_f[ns_f["Cidade"].astype(str).isin(cidade_sel)]
        if canal_sel and "Canal" in ns_f.columns:
            ns_f = ns_f[ns_f["Canal"].astype(str).isin(canal_sel)]
        if situacao_sel and "Situação" in ns_f.columns:
            ns_f = ns_f[ns_f["Situação"].astype(str).isin(situacao_sel)]
        if prazo_sel and "Faixa Prazo" in ns_f.columns:
            ns_f = ns_f[ns_f["Faixa Prazo"].astype(str).isin(prazo_sel)]
        if ofensor_sel and "Ofensor NS" in ns_f.columns:
            ns_f = ns_f[ns_f["Ofensor NS"].astype(str).isin(ofensor_sel)]
        if periodo_ns and isinstance(periodo_ns, tuple) and len(periodo_ns) == 2:
            ini, fim = pd.to_datetime(periodo_ns[0]), pd.to_datetime(periodo_ns[1])
            ns_f = ns_f[(ns_f["Data Prevista"] >= ini) & (ns_f["Data Prevista"] <= fim)]

        st.markdown("#### Resumo executivo")
        if ns_f.empty:
            st.warning("Nenhum dado encontrado com os filtros selecionados.")
        else:
            c1, c2 = st.columns(2)
            for idx, operacao in enumerate(["1P MAGALU", "3P FULL"]):
                base_op = ns_f[ns_f["Operação"] == operacao]
                total = pd.to_numeric(base_op.get("Total_pedidos", pd.Series([1]*len(base_op))), errors="coerce").fillna(1).sum()
                dentro = pd.to_numeric(base_op["Dentro Prazo"], errors="coerce").fillna(0).sum()
                fora = pd.to_numeric(base_op["Fora Prazo"], errors="coerce").fillna(0).sum()
                ns_val = dentro / total if total else np.nan
                with c1 if idx == 0 else c2:
                    card_ns(operacao, ns_val, total, dentro, fora, operacao)
                    st.caption(meta_texto_ns(operacao))

            st.divider()
            st.markdown("#### NS diário")
            diario = consolidar_ns(ns_f, ["Operação", "Data Prevista"])
            if not diario.empty:
                diario["Data_label"] = pd.to_datetime(diario["Data Prevista"], errors="coerce").dt.strftime("%d/%m/%Y")
                fig = px.line(diario, x="Data_label", y="NS", color="Operação", markers=True, text=diario["NS"].map(lambda x: f"{x*100:.1f}%" if pd.notna(x) else ""), hover_data=["Total_pedidos", "Dentro_prazo", "Fora_prazo"], title="NS diário por operação")
                fig.update_yaxes(tickformat=".0%")
                fig.update_traces(textposition="top center")
                if "1P MAGALU" in diario["Operação"].astype(str).unique():
                    fig.add_hline(y=0.96, line_dash="dash", annotation_text="Meta 1P 96%", annotation_position="top left")
                if "3P FULL" in diario["Operação"].astype(str).unique():
                    fig.add_hline(y=0.95, line_dash="dot", annotation_text="Meta Full 95%", annotation_position="bottom left")
                st.plotly_chart(fig, use_container_width=True)
                st.dataframe(formatar_ns_tabela(diario.sort_values(["Data Prevista", "Operação"])), use_container_width=True, height=300)

            st.markdown("#### Semana a semana")
            ns_semana_base = ns_f.copy()
            if "Início Semana" not in ns_semana_base.columns or ns_semana_base["Início Semana"].isna().all():
                ns_semana_base["Início Semana"] = ns_semana_base["Data Prevista"] - pd.to_timedelta(pd.to_datetime(ns_semana_base["Data Prevista"], errors="coerce").dt.weekday, unit="D")
            ns_semana_base["Semana"] = pd.to_datetime(ns_semana_base["Início Semana"], errors="coerce").dt.strftime("%d/%m")
            semanal_ns = consolidar_ns(ns_semana_base, ["Operação", "Início Semana", "Semana"])
            if not semanal_ns.empty:
                semanal_ns = semanal_ns.sort_values(["Operação", "Início Semana"])
                semanal_ns["Evolução"] = semanal_ns.groupby("Operação")["NS"].diff().apply(lambda x: "Melhorando" if pd.notna(x) and x > 0 else ("Piorando" if pd.notna(x) and x < 0 else "Estável"))
                semanal_ns["Faltam_meta"] = semanal_ns.apply(lambda r: pedidos_para_meta(r["Total_pedidos"], r["Dentro_prazo"], meta_ns_operacao(r["Operação"])), axis=1)
                fig_sem_ns = px.bar(
                    semanal_ns,
                    x="Semana",
                    y="NS",
                    color="Operação",
                    barmode="group",
                    text="NS",
                    hover_data=["Total_pedidos", "Fora_prazo", "Faltam_meta", "Evolução"],
                    title="NS semana a semana por operação"
                )
                fig_sem_ns.update_yaxes(tickformat=".0%")
                aplicar_rotulos_barra(fig_sem_ns, formato="%{y:.1%}")
                if "1P MAGALU" in semanal_ns["Operação"].astype(str).unique():
                    fig_sem_ns.add_hline(y=0.96, line_dash="dash", annotation_text="Meta 1P 96%", annotation_position="top left")
                if "3P FULL" in semanal_ns["Operação"].astype(str).unique():
                    fig_sem_ns.add_hline(y=0.95, line_dash="dot", annotation_text="Meta Full 95%", annotation_position="bottom left")
                st.plotly_chart(fig_sem_ns, use_container_width=True)
                tabela_semana_ns = semanal_ns.drop(columns=[c for c in ["Início Semana"] if c in semanal_ns.columns])
                st.dataframe(formatar_ns_tabela(tabela_semana_ns), use_container_width=True, height=300)

            st.markdown("#### Análise por modal")
            modal = consolidar_ns(ns_f, ["Operação", "Modal"])
            if not modal.empty:
                fig_modal = px.bar(modal.sort_values("NS"), x="Modal", y="NS", color="Operação", barmode="group", text="NS", hover_data=["Total_pedidos", "Dentro_prazo", "Fora_prazo"], title="NS por modal")
                fig_modal.update_yaxes(tickformat=".0%")
                aplicar_rotulos_barra(fig_modal, formato="%{y:.1%}")
                st.plotly_chart(fig_modal, use_container_width=True)
                st.dataframe(formatar_ns_tabela(modal.sort_values(["Operação", "NS"])), use_container_width=True, height=300)

            st.markdown("#### Indicador Ofensor NS")
            st.caption("Mostra os maiores motivos de perda de NS por operação e modal. Exemplo: Atraso Transporte, Cliente Ausente, Endereço, etc. O filtro de Ofensor NS acima altera todos os gráficos e tabelas desta aba.")
            ofensores_modal = consolidar_ofensor_ns_por_modal(ns_f)
            if ofensores_modal.empty:
                st.info("Não encontrei Ofensor NS preenchido para os pedidos fora do prazo com os filtros atuais.")
            else:
                st.dataframe(formatar_ofensor_ns_tabela(ofensores_modal.head(80)), use_container_width=True, height=420)

            st.markdown("#### Análise por prazo cliente")
            prazo = consolidar_ns(ns_f, ["Operação", "Faixa Prazo"]) if "Faixa Prazo" in ns_f.columns else pd.DataFrame()
            if not prazo.empty:
                ordem_prazo = {"D0": 0, "D1": 1, "D2": 2, "D3": 3, "D4": 4, "D5": 5, "Maior D5": 6, "Não informado": 99}
                prazo["ordem"] = prazo["Faixa Prazo"].map(ordem_prazo).fillna(98)
                prazo = prazo.sort_values(["Operação", "ordem"])
                fig_prazo = px.bar(prazo, x="Faixa Prazo", y="NS", color="Operação", barmode="group", text="NS", hover_data=["Total_pedidos", "Dentro_prazo", "Fora_prazo"], title="NS por prazo cliente")
                fig_prazo.update_yaxes(tickformat=".0%")
                aplicar_rotulos_barra(fig_prazo, formato="%{y:.1%}")
                st.plotly_chart(fig_prazo, use_container_width=True)
                st.dataframe(formatar_ns_tabela(prazo.drop(columns=["ordem"])), use_container_width=True, height=300)

            col_cid, col_transp = st.columns(2)
            with col_cid:
                st.markdown("#### Top cidades ofensoras")
                cidade = consolidar_ns(ns_f, ["Operação", "Cidade"])
                cidade = cidade[cidade["Total_pedidos"] >= 10].sort_values(["Fora_prazo", "NS"], ascending=[False, True]).head(30)
                st.dataframe(formatar_ns_tabela(cidade), use_container_width=True, height=480)

            with col_transp:
                st.markdown("#### Top transportadores ofensores")
                transp = consolidar_ns(ns_f, ["Operação", "Transportador"])
                transp = transp[transp["Total_pedidos"] >= 10].sort_values(["Fora_prazo", "NS"], ascending=[False, True]).head(30)
                st.dataframe(formatar_ns_tabela(transp), use_container_width=True, height=480)

            st.markdown("#### Insights automáticos")
            insights_ns = gerar_insights_ns(ns_f)
            if insights_ns.empty:
                st.success("Nenhum insight crítico encontrado com os filtros atuais.")
            else:
                st.dataframe(insights_ns, use_container_width=True, height=520)

            st.markdown("#### Base detalhada filtrada")
            detalhe = ns_f.copy()
            detalhe["Data Prevista"] = pd.to_datetime(detalhe["Data Prevista"], errors="coerce").dt.strftime("%d/%m/%Y")
            if "Total_pedidos" in detalhe.columns:
                detalhe["Total pedidos"] = pd.to_numeric(detalhe["Total_pedidos"], errors="coerce").fillna(1).astype(int)
                detalhe["Dentro prazo"] = pd.to_numeric(detalhe["Dentro Prazo"], errors="coerce").fillna(0).astype(int)
                detalhe["Fora prazo"] = pd.to_numeric(detalhe["Fora Prazo"], errors="coerce").fillna(0).astype(int)
                detalhe["NS"] = detalhe["Dentro prazo"] / detalhe["Total pedidos"].replace(0, np.nan)
                detalhe["NS"] = detalhe["NS"].map(lambda x: br_numero(x, percentual=True) if pd.notna(x) else "")
            else:
                detalhe["Status NS"] = detalhe["Dentro Prazo"].map({1: "Dentro do prazo", 0: "Fora do prazo"})
            detalhe_cols = [c for c in ["Operação", "Data Prevista", "Modal", "Transportador", "Cidade", "Canal", "Situação", "Faixa Prazo", "Ofensor NS", "Pedido", "Status NS", "Total pedidos", "Dentro prazo", "Fora prazo", "NS"] if c in detalhe.columns]
            detalhe = detalhe[detalhe_cols]
            st.dataframe(detalhe, use_container_width=True, height=420)


with aba_eclusa:
    render_resumo_compacto_superior()
    st.subheader("Saída Eclusa")
    st.caption("Meta geral: 99,20%. Abaixo de 98,20% zera. Referência de período: data_nota. Atraso = data_saida_cd_considerada - Data_limite_origem.")
    if eclusa_atual.empty:
        st.info("Envie a base de Saída Eclusa (.csv) no menu lateral para iniciar a análise.")
    else:
        ecl = eclusa_atual.copy()
        f1, f2, f3, f4, f5 = st.columns(5)
        meses = sorted(ecl["Mês"].dropna().astype(str).unique().tolist()) if "Mês" in ecl.columns else []
        semanas = []
        if "Início Semana" in ecl.columns:
            ecl["Semana"] = pd.to_datetime(ecl["Início Semana"], errors="coerce").dt.strftime("%d/%m")
            sem_df = ecl[["Início Semana", "Semana"]].dropna().drop_duplicates().sort_values("Início Semana")
            semanas = sem_df["Semana"].astype(str).tolist()
        with f1:
            meses_sel = st.multiselect("Mês", meses, default=meses[-1:] if meses else [], key="eclusa_mes")
        with f2:
            semanas_sel = st.multiselect("Semana", semanas, key="eclusa_semana")
        with f3:
            ops_sel = st.multiselect("Operação", sorted(ecl["Operação"].dropna().astype(str).unique().tolist()), key="eclusa_operacao")
        with f4:
            mod_sel = st.multiselect("Modal", sorted(ecl["Modal"].dropna().astype(str).unique().tolist()), key="eclusa_modal")
        with f5:
            status_sel = st.multiselect("Status saída", sorted(ecl["Status Saída"].dropna().astype(str).unique().tolist()), key="eclusa_status_saida")

        if "Data Nota" in ecl.columns and ecl["Data Nota"].notna().any():
            data_min_e = min(ecl["Data Nota"].dropna()).date()
            data_max_e = max(ecl["Data Nota"].dropna()).date()
            data_range_e = st.date_input("Data nota", value=(data_min_e, data_max_e), min_value=data_min_e, max_value=data_max_e, format="DD/MM/YYYY", key="eclusa_data_nota")
        else:
            data_range_e = None

        if meses_sel:
            ecl = ecl[ecl["Mês"].astype(str).isin(meses_sel)]
        if semanas_sel and "Semana" in ecl.columns:
            ecl = ecl[ecl["Semana"].astype(str).isin(semanas_sel)]
        if data_range_e and "Data Nota" in ecl.columns and isinstance(data_range_e, tuple) and len(data_range_e) == 2:
            dt_ini_e, dt_fim_e = pd.to_datetime(data_range_e[0]), pd.to_datetime(data_range_e[1])
            ecl = ecl[(ecl["Data Nota"] >= dt_ini_e) & (ecl["Data Nota"] <= dt_fim_e)]
        if ops_sel:
            ecl = ecl[ecl["Operação"].astype(str).isin(ops_sel)]
        if mod_sel:
            ecl = ecl[ecl["Modal"].astype(str).isin(mod_sel)]
        if status_sel:
            ecl = ecl[ecl["Status Saída"].astype(str).isin(status_sel)]

        total = ecl["Total_pedidos"].sum()
        dentro = ecl["Dentro Prazo"].sum()
        fora = ecl["Fora Prazo"].sum()
        ns_val = dentro / total if total else np.nan
        atraso_medio = ecl.loc[ecl["Atraso horas"] > 0, "Atraso horas"].mean()
        atraso_max = ecl["Atraso horas"].max()

        k1, k2, k3, k4 = st.columns(4)
        with k1:
            card_indicador_meta("Saída Eclusa", ns_val, total, dentro, fora, META_ECLUSA, tipo="Eclusa")
        with k2:
            st.metric("Total pedidos", f"{int(total):,}".replace(",", "."))
            st.metric("Fora do prazo", f"{int(fora):,}".replace(",", "."))
        with k3:
            st.metric("Faltam para meta", f"{pedidos_para_meta(total, dentro, META_ECLUSA):,}".replace(",", "."))
            st.metric("Dentro do prazo", f"{int(dentro):,}".replace(",", "."))
        with k4:
            st.metric("Atraso médio", f"{atraso_medio:.1f}h" if pd.notna(atraso_medio) else "0,0h")
            st.metric("Maior atraso", f"{atraso_max:.1f}h" if pd.notna(atraso_max) else "0,0h")

        st.markdown("### Evolução diária e semanal")
        diario = consolidar_eclusa(ecl, ["Data Nota"])
        diario = diario.sort_values("Data Nota")
        if not diario.empty:
            diario["Data_formatada"] = pd.to_datetime(diario["Data Nota"], errors="coerce").dt.strftime("%d/%m/%Y")
            fig_d = px.line(diario, x="Data_formatada", y="NS", markers=True, text="NS", hover_data=["Total_pedidos", "Fora_prazo", "Faltam_meta"], title="NS diário - Saída Eclusa")
            fig_d.update_traces(texttemplate="%{y:.2%}", textposition="top center")
            fig_d.update_yaxes(tickformat=".2%")
            fig_d.add_hline(y=META_ECLUSA, line_dash="dash", line_color="green", annotation_text="Meta 99,20%")
            fig_d.add_hline(y=ZERADO_ECLUSA, line_dash="dot", line_color="purple", annotation_text="Zera 98,20%")
            st.plotly_chart(fig_d, use_container_width=True)

        semanal = consolidar_eclusa(ecl, ["Início Semana"]) if "Início Semana" in ecl.columns else consolidar_eclusa(ecl, ["Semana"])
        if not semanal.empty:
            if "Início Semana" in semanal.columns:
                semanal = semanal.sort_values("Início Semana")
                semanal["Semana"] = pd.to_datetime(semanal["Início Semana"], errors="coerce").dt.strftime("%d/%m")
            else:
                semanal = semanal.sort_values("Semana")
            semanal["Evolução"] = ""
            semanal["Delta"] = semanal["NS"].diff()
            semanal["Evolução"] = semanal["Delta"].apply(lambda x: "Melhorando" if pd.notna(x) and x > 0 else ("Piorando" if pd.notna(x) and x < 0 else "Estável"))
            fig_s = px.bar(semanal, x="Semana", y="NS", text="NS", hover_data=["Total_pedidos", "Fora_prazo", "Faltam_meta", "Evolução"], title="Quadro semanal - Saída Eclusa")
            fig_s.update_yaxes(tickformat=".2%")
            fig_s.add_hline(y=META_ECLUSA, line_dash="dash", line_color="green", annotation_text="Meta 99,20%")
            aplicar_rotulos_barra(fig_s, formato="%{y:.2%}")
            st.plotly_chart(fig_s, use_container_width=True)
            tabela_semanal = semanal.drop(columns=[c for c in ["Início Semana", "Semana Mês"] if c in semanal.columns])
            cols_ordem = ["Semana"] + [c for c in tabela_semanal.columns if c != "Semana"]
            tabela_semanal = tabela_semanal[cols_ordem]
            st.dataframe(formatar_eclusa_tabela(tabela_semanal), use_container_width=True, height=260)

        st.markdown("### Ofensores e análises")
        cmodal, ctransp = st.columns(2)
        with cmodal:
            st.markdown("#### Por modal")
            modal = consolidar_eclusa(ecl, ["Modal"]).sort_values(["NS", "Fora_prazo"], ascending=[True, False])
            fig_m = px.bar(modal, x="Modal", y="NS", text="NS", hover_data=["Total_pedidos", "Fora_prazo", "Faltam_meta"], title="NS por modal")
            fig_m.update_yaxes(tickformat=".2%")
            fig_m.add_hline(y=META_ECLUSA, line_dash="dash", line_color="green")
            aplicar_rotulos_barra(fig_m, formato="%{y:.2%}")
            st.plotly_chart(fig_m, use_container_width=True)
            st.dataframe(formatar_eclusa_tabela(modal), use_container_width=True, height=300)
        with ctransp:
            st.markdown("#### Top transportadores ofensores")
            transp = consolidar_eclusa(ecl, ["Transportador Grupo"])
            transp = transp[transp["Total_pedidos"] >= 10].sort_values(["Fora_prazo", "NS"], ascending=[False, True]).head(30)
            st.dataframe(formatar_eclusa_tabela(transp), use_container_width=True, height=510)

        ccid, catraso = st.columns(2)
        with ccid:
            st.markdown("#### Top cidades ofensoras")
            cidade = consolidar_eclusa(ecl, ["Cidade"])
            cidade = cidade[cidade["Total_pedidos"] >= 10].sort_values(["Fora_prazo", "NS"], ascending=[False, True]).head(30)
            st.dataframe(formatar_eclusa_tabela(cidade), use_container_width=True, height=450)
        with catraso:
            st.markdown("#### Faixa de atraso")
            faixa = ecl.groupby("Faixa Atraso", dropna=False).agg(Pedidos=("Pedido", "nunique")).reset_index().sort_values("Pedidos", ascending=False)
            fig_f = px.bar(faixa, x="Faixa Atraso", y="Pedidos", text="Pedidos", title="Distribuição por faixa de atraso")
            aplicar_rotulos_barra(fig_f)
            st.plotly_chart(fig_f, use_container_width=True)

        st.markdown("### Insights automáticos - Saída Eclusa")
        insights_e = gerar_insights_eclusa(ecl)
        if insights_e.empty:
            st.success("Nenhum insight crítico encontrado com os filtros atuais.")
        else:
            st.dataframe(insights_e, use_container_width=True, height=450)

        st.markdown("### Base detalhada filtrada")
        detalhe_cols = [c for c in ["Pedido", "Data Nota", "Operação", "Modal", "Transportador Grupo", "Transportador", "Cidade", "Status Saída", "Dentro Prazo", "Atraso horas", "Faixa Atraso"] if c in ecl.columns]
        detalhe = ecl[detalhe_cols].copy()
        if "Data Nota" in detalhe.columns:
            detalhe["Data Nota"] = pd.to_datetime(detalhe["Data Nota"], errors="coerce").dt.strftime("%d/%m/%Y")
        detalhe["Status Eclusa"] = detalhe["Dentro Prazo"].map({1: "Dentro do prazo", 0: "Fora do prazo"}) if "Dentro Prazo" in detalhe.columns else ""
        st.dataframe(formatar_eclusa_tabela(detalhe), use_container_width=True, height=420)

with aba_bases:
    render_resumo_compacto_superior()
    st.subheader("Bases / Histórico")
    st.caption("Área operacional para salvar histórico, excluir datas e manter cargas especiais.")

    st.markdown("### Histórico de Cubagem")
    if not sheet_id:
        st.warning("Cole/fixe o ID do Google Sheets para usar o histórico.")
    else:
        c1, c2 = st.columns(2)
        with c1:
            if st.button("Adicionar base atual ao histórico", type="primary", disabled=cargas_atual.empty):
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
                datas_excluir = st.multiselect("Excluir cubagem por data", datas_hist)
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

    st.divider()
    st.markdown("### Histórico de Nível de Serviço")
    st.caption("Salva o NS agregado no Google Sheets para não pesar o app. Ao salvar uma nova base, o app substitui somente as datas presentes no arquivo enviado, corrigindo desconsiderações sem duplicar.")
    if not sheet_id:
        st.warning("Cole/fixe o ID do Google Sheets para salvar o histórico de NS.")
    else:
        h1, h2, h3 = st.columns(3)
        with h1:
            if st.button("Salvar NS 1P no histórico", disabled=ns_1p_atual.empty):
                try:
                    hist_existente = ler_historico_ns(sheet_id, "1P")
                    novo = preparar_ns_para_historico(ns_1p_atual)
                    combinado = juntar_ns_sem_duplicar(hist_existente, novo)
                    escrever_historico_ns(sheet_id, "1P", combinado)
                    st.success(f"Histórico NS 1P atualizado: {len(combinado):,.0f} linha(s) agregada(s).".replace(",", "."))
                    st.cache_data.clear()
                except Exception as e:
                    st.error(f"Erro ao salvar NS 1P: {e}")
        with h2:
            if st.button("Salvar NS Full no histórico", disabled=ns_full_atual.empty):
                try:
                    hist_existente = ler_historico_ns(sheet_id, "Full")
                    novo = preparar_ns_para_historico(ns_full_atual)
                    combinado = juntar_ns_sem_duplicar(hist_existente, novo)
                    escrever_historico_ns(sheet_id, "Full", combinado)
                    st.success(f"Histórico NS Full atualizado: {len(combinado):,.0f} linha(s) agregada(s).".replace(",", "."))
                    st.cache_data.clear()
                except Exception as e:
                    st.error(f"Erro ao salvar NS Full: {e}")
        with h3:
            st.metric("Linhas agregadas NS", f"{(len(historico_ns_1p_raw) + len(historico_ns_full_raw)):,.0f}".replace(",", "."))

    st.divider()
    st.markdown("### Histórico de Saída Eclusa")
    st.caption("Salva a eclusa agregada no Google Sheets. Ao salvar, substitui as datas presentes no novo arquivo, corrigindo reprocessamentos sem duplicar.")
    if not sheet_id:
        st.warning("Cole/fixe o ID do Google Sheets para salvar o histórico de Eclusa.")
    else:
        e1, e2 = st.columns(2)
        with e1:
            if st.button("Salvar Saída Eclusa no histórico", disabled=eclusa_atual.empty):
                try:
                    hist_existente = ler_historico_eclusa(sheet_id)
                    novo = preparar_eclusa_para_historico(eclusa_atual)
                    combinado = juntar_historico_substituindo_datas(hist_existente, novo, "Data Nota", COLUNAS_ECLUSA_HISTORICO)
                    escrever_historico_eclusa(sheet_id, combinado)
                    st.success(f"Histórico Eclusa atualizado: {len(combinado):,.0f} linha(s) agregada(s).".replace(",", "."))
                    st.cache_data.clear()
                except Exception as e:
                    st.error(f"Erro ao salvar Saída Eclusa: {e}")
        with e2:
            st.metric("Linhas agregadas Eclusa", f"{len(historico_eclusa_raw):,.0f}".replace(",", "."))

    st.divider()
    st.markdown("### Cargas especiais")
    st.caption("Cadastre IDs de carga que distorcem a cubagem, como Cofre ou Cerveja.")
    if not sheet_id:
        st.warning("Cole/fixe o ID do Google Sheets para cadastrar e salvar cargas especiais.")
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
