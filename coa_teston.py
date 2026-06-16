"""
COA · TESTON — Dashboard Safra 2026/27
Streamlit + Plotly | v3
Login + SharePoint (Graph API) + Split Colheita/Agro + Período fixo

SECRETS necessários no Streamlit Cloud (ou .streamlit/secrets.toml):
  SP_TENANT_ID         = "..."
  SP_CLIENT_ID         = "..."
  SP_CLIENT_SECRET     = "..."
  SP_SITE_URL          = "https://metalcana.sharepoint.com/sites/MSCOLHEITA"
  SP_APT_PATH          = "/VITOR/AGRITEL/apontamentos.xlsx"
  SP_PERF_PATH         = "/VITOR/AGRITEL/performance.xlsx"
  SP_VEICULOS_PATH     = "/VITOR/AGRITEL/Veículos___AGRITEL.xlsx"
  SP_FROTAS_CC_PATH    = "/VITOR/AGRITEL/Frotas_CentroDeCusto.xlsx"
  SP_PERIODO_PATH      = "/VITOR/AGRITEL/periodo_agritel.jpg"
  PERIODO_INICIO       = "2026-04-01 07:00"
  PERIODO_FIM          = "2026-04-30 06:59"
  APP_TIPO             = "colheita"   # "colheita" | "agro"
  DASHBOARD_SENHA_HASH = "<sha256 da senha>"
"""

import re, io, os, hashlib, pathlib
from datetime import datetime
import pandas as pd
import numpy as np
import streamlit as st
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# ══════════════════════════════════════════════════════════════════════════════
# CONFIGURAÇÕES
# ══════════════════════════════════════════════════════════════════════════════
def _secret(key: str, default: str = "") -> str:
    try:
        return st.secrets[key]
    except Exception:
        return os.environ.get(key, default)

SP_TENANT_ID      = _secret("SP_TENANT_ID")
SP_CLIENT_ID      = _secret("SP_CLIENT_ID")
SP_CLIENT_SECRET  = _secret("SP_CLIENT_SECRET")
SP_SITE_URL       = _secret("SP_SITE_URL",       "https://metalcana.sharepoint.com/sites/MSCOLHEITA")
SP_APT_PATH       = _secret("SP_APT_PATH",       "/VITOR/AGRITEL/apontamentos.xlsx")
SP_PERF_PATH      = _secret("SP_PERF_PATH",      "/VITOR/AGRITEL/performance.xlsx")
SP_VEICULOS_PATH  = _secret("SP_VEICULOS_PATH",  "/VITOR/AGRITEL/Veículos___AGRITEL.xlsx")
SP_FROTAS_CC_PATH = _secret("SP_FROTAS_CC_PATH", "/VITOR/AGRITEL/Frotas_CentroDeCusto.xlsx")
SP_PERIODO_PATH   = _secret("SP_PERIODO_PATH",   "/VITOR/AGRITEL/periodo_agritel.jpg")
PERIODO_INICIO    = _secret("PERIODO_INICIO",    "2026-04-01 07:00")
PERIODO_FIM       = _secret("PERIODO_FIM",       "2026-04-30 06:59")

# "colheita" → mostra colhedoras e transbordos
# "agro"     → mostra tudo que NÃO é colhedora nem transbordo
APP_TIPO = _secret("APP_TIPO", "colheita").lower().strip()

SENHA_HASH = _secret("DASHBOARD_SENHA_HASH",
                     hashlib.sha256("teston2026".encode()).hexdigest())

# ── Títulos / ícones por tipo ─────────────────────────────────────────────────
TIPO_CONFIG = {
    "colheita": {"icone": "🌾", "titulo": "COA · TESTON — Colheita", "subtitulo": "Colhedoras & Transbordos"},
    "agro":     {"icone": "🚜", "titulo": "COA · TESTON — Agro",     "subtitulo": "Máquinas Agrícolas"},
}
CONF = TIPO_CONFIG.get(APP_TIPO, TIPO_CONFIG["colheita"])

# Palavras-chave para classificar máquinas como "colheita"
KEYWORDS_COLHEITA = ["colhedora", "transbordo"]

# ── Cálculo das horas do período ──────────────────────────────────────────────
def calc_period_hours() -> float:
    """Retorna o total de horas do período configurado."""
    try:
        fmt = "%Y-%m-%d %H:%M"
        ini = datetime.strptime(PERIODO_INICIO, fmt)
        fim = datetime.strptime(PERIODO_FIM, fmt)
        return round((fim - ini).total_seconds() / 3600, 1)
    except Exception:
        return 720.0

PERIOD_H = calc_period_hours()

# ══════════════════════════════════════════════════════════════════════════════
# CORES
# ══════════════════════════════════════════════════════════════════════════════
C_PRODUTIVO  = "#1a7a4a"
C_MANUTENCAO = "#c0392b"
C_ESPERA     = "#c97d10"
C_MANOBRA    = "#7F77DD"
C_PARADA     = "#888780"
C_SEM_APT    = "#b4b2a9"
C_F1262      = "#1a7a4a"
C_F1263      = "#1a5fa8"
C_DISP       = "#27ae60"
C_ELEVADOR   = "#9b59b6"
C_MOTOR      = "#2980b9"

CATEGORIA_MAP = {
    "Colhendo":                   "Produtivo",
    "Manobra":                    "Manobra",
    "Deslocando":                 "Manobra",
    "Quebra":                     "Manutenção",
    "Manutenção Programada":      "Manutenção",
    "Troca Faquinha":             "Manutenção",
    "PARADA OPERACIONAL":         "Manutenção",
    "Chuva / Umidade":            "Espera/Imprevisto",
    "Usina Parada":               "Espera/Imprevisto",
    "Falta de Area Liberada":     "Espera/Imprevisto",
    "Falta de Caminhão":          "Espera/Imprevisto",
    "Aguardando Transbordo":      "Espera/Imprevisto",
    "Abastecimento":              "Parada Operacional",
    "Refeição":                   "Parada Operacional",
    "Checklist":                  "Parada Operacional",
    "Final Jornada de Trabalho":  "Parada Operacional",
    "Início Jornada de Trabalho": "Parada Operacional",
}

CAT_CORES = {
    "Produtivo":          C_PRODUTIVO,
    "Manutenção":         C_MANUTENCAO,
    "Espera/Imprevisto":  C_ESPERA,
    "Manobra":            C_MANOBRA,
    "Parada Operacional": C_PARADA,
    "Sem Apontamento":    C_SEM_APT,
}

APT_CORES = {
    "Colhendo":                   C_PRODUTIVO,
    "Aguardando Transbordo":      C_ESPERA,
    "Chuva / Umidade":            C_F1263,
    "Quebra":                     C_MANUTENCAO,
    "Manobra":                    C_MANOBRA,
    "Sem Apontamento":            C_SEM_APT,
    "Manutenção Programada":      "#2ea05f",
    "Abastecimento":              "#5DCAA5",
    "Troca Faquinha":             "#D85A30",
    "Checklist":                  C_PARADA,
    "Refeição":                   C_PARADA,
    "Deslocando":                 "#9b9b9b",
    "Final Jornada de Trabalho":  "#9b9b9b",
    "Início Jornada de Trabalho": "#9b9b9b",
    "Falta de Area Liberada":     C_MANUTENCAO,
    "Falta de Caminhão":          C_MANUTENCAO,
    "PARADA OPERACIONAL":         C_MANUTENCAO,
    "Usina Parada":               "#9b9b9b",
}

# Fallback se Veículos AGRITEL não estiver disponível ainda
DEVICE_MAP_FALLBACK = {
    "671a433018212ed9057ce3ee": "Frota 1262",
    "671a433018212ed9057ce3f1": "Frota 1263",
}

# ── PAGE CONFIG ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title=CONF["titulo"],
    page_icon=CONF["icone"],
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
[data-testid="stMetricValue"]       { font-size: 1.6rem !important; font-weight: 700; }
[data-testid="stMetricLabel"]       { font-size: 0.72rem !important; text-transform: uppercase; letter-spacing: .05em; }
[data-testid="stMetricDelta"]       { font-size: 0.72rem !important; }
[data-testid="stSidebar"]           { background-color: #1c1e18; }
[data-testid="stSidebar"] *         { color: #d8dbd4 !important; }
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3        { color: #a8c832 !important; }
div[data-testid="metric-container"] { background: rgba(255,255,255,0.05); border-radius: 10px; padding: 12px 16px; border-left: 3px solid #1a7a4a; }
.block-container                    { padding-top: 1.4rem; }
.js-plotly-plot .plotly, .js-plotly-plot .plotly .svg-container { background: transparent !important; }
.periodo-badge {
    display: inline-flex; align-items: center; gap: 8px;
    background: rgba(168,200,50,0.12); border: 1px solid rgba(168,200,50,0.3);
    border-radius: 8px; padding: 6px 14px; margin-bottom: 12px;
    font-size: 12px; color: #a8c832;
}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# LOGIN
# ══════════════════════════════════════════════════════════════════════════════
def check_login():
    if st.session_state.get("autenticado"):
        return
    st.markdown(f"""
    <div style="max-width:360px;margin:80px auto;text-align:center">
      <h1 style="color:#a8c832;font-size:2rem">{CONF['icone']} {CONF['titulo']}</h1>
      <p style="color:#888;margin-bottom:28px">{CONF['subtitulo']} · Safra 2026/27</p>
    </div>
    """, unsafe_allow_html=True)
    col_c = st.columns([1, 2, 1])[1]
    with col_c:
        senha = st.text_input("Senha de acesso", type="password", key="senha_input",
                              placeholder="Digite a senha...")
        if st.button("Entrar", use_container_width=True, type="primary"):
            if hashlib.sha256(senha.encode()).hexdigest() == SENHA_HASH:
                st.session_state["autenticado"] = True
                st.rerun()
            else:
                st.error("Senha incorreta.")
    st.stop()

check_login()

# ══════════════════════════════════════════════════════════════════════════════
# PARSERS
# ══════════════════════════════════════════════════════════════════════════════
def parse_horas(txt) -> float:
    if isinstance(txt, (int, float)):
        return float(txt) if not (isinstance(txt, float) and np.isnan(txt)) else 0.0
    if not isinstance(txt, str):
        return 0.0
    m = re.search(r"([\d.,]+)\s*h", txt)
    if m:
        return float(m.group(1).replace(".", "").replace(",", "."))
    return 0.0

def parse_num(txt) -> float:
    if isinstance(txt, (int, float)):
        return float(txt) if not (isinstance(txt, float) and np.isnan(txt)) else 0.0
    if not isinstance(txt, str):
        return 0.0
    cleaned = re.sub(r"[^\d,.]", "", txt).replace(".", "").replace(",", ".")
    try:
        return float(cleaned)
    except Exception:
        return 0.0

def parse_tempo_perf(txt) -> float:
    if isinstance(txt, (int, float)):
        return float(txt) if not (isinstance(txt, float) and np.isnan(txt)) else 0.0
    if not isinstance(txt, str):
        return 0.0
    m = re.search(r"(\d+)[,.]?(\d*)\s*h", txt)
    if not m:
        return 0.0
    return float(m.group(1)) + (float(m.group(2)) / 60 if m.group(2) else 0.0)

def _normalize(s: str) -> str:
    """Remove acentos e coloca em minúsculas para comparação de colunas."""
    s = s.lower()
    for a, b in [("á","a"),("ã","a"),("â","a"),("à","a"),("é","e"),("ê","e"),
                 ("í","i"),("ó","o"),("õ","o"),("ô","o"),("ú","u"),("ç","c")]:
        s = s.replace(a, b)
    return s

def _find_col(df: pd.DataFrame, keywords: list[str]) -> str | None:
    """Encontra a primeira coluna cujo nome normalizado contenha alguma keyword."""
    norm = {_normalize(c): c for c in df.columns}
    for kw in keywords:
        for n, orig in norm.items():
            if kw in n:
                return orig
    return None

# ══════════════════════════════════════════════════════════════════════════════
# SHAREPOINT — Microsoft Graph API
# ══════════════════════════════════════════════════════════════════════════════
@st.cache_data(ttl=600)
def load_from_sharepoint() -> dict:
    """
    Baixa todos os arquivos necessários do SharePoint.
    Retorna dict com chaves: 'apt', 'perf', 'veiculos', 'frotas_cc', 'periodo_img'
    """
    result = {k: None for k in ("apt", "perf", "veiculos", "frotas_cc", "periodo_img")}
    try:
        import requests

        token_resp = requests.post(
            f"https://login.microsoftonline.com/{SP_TENANT_ID}/oauth2/v2.0/token",
            data={
                "grant_type":    "client_credentials",
                "client_id":     SP_CLIENT_ID,
                "client_secret": SP_CLIENT_SECRET,
                "scope":         "https://graph.microsoft.com/.default",
            },
            timeout=15,
        )
        token_resp.raise_for_status()
        headers = {"Authorization": f"Bearer {token_resp.json()['access_token']}"}

        hostname  = SP_SITE_URL.replace("https://", "").split("/")[0]
        site_path = "/".join(SP_SITE_URL.replace("https://", "").split("/")[1:])
        site_resp = requests.get(
            f"https://graph.microsoft.com/v1.0/sites/{hostname}:/{site_path}",
            headers=headers, timeout=15,
        )
        site_resp.raise_for_status()
        site_id = site_resp.json()["id"]

        drive_resp = requests.get(
            f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive",
            headers=headers, timeout=15,
        )
        drive_resp.raise_for_status()
        drive_id = drive_resp.json()["id"]

        def _download(path: str) -> bytes | None:
            try:
                r = requests.get(
                    f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:{path}:/content",
                    headers=headers, timeout=30,
                )
                r.raise_for_status()
                return r.content
            except Exception as e:
                st.warning(f"⚠️ Arquivo não encontrado no SharePoint: `{path}` — {e}")
                return None

        result["apt"]        = _download(SP_APT_PATH)
        result["perf"]       = _download(SP_PERF_PATH)
        result["veiculos"]   = _download(SP_VEICULOS_PATH)
        result["frotas_cc"]  = _download(SP_FROTAS_CC_PATH)
        result["periodo_img"] = _download(SP_PERIODO_PATH)

    except Exception as exc:
        import traceback
        st.error(f"❌ Erro SharePoint: {exc}\n\n```\n{traceback.format_exc()}\n```")

    return result


def _upload_to_sharepoint(file_path: str, sp_path: str) -> bool:
    """Upload de um arquivo local para o SharePoint. Usado pelo script de upload do período."""
    try:
        import requests
        token_resp = requests.post(
            f"https://login.microsoftonline.com/{SP_TENANT_ID}/oauth2/v2.0/token",
            data={"grant_type": "client_credentials", "client_id": SP_CLIENT_ID,
                  "client_secret": SP_CLIENT_SECRET, "scope": "https://graph.microsoft.com/.default"},
            timeout=15,
        )
        token_resp.raise_for_status()
        headers = {"Authorization": f"Bearer {token_resp.json()['access_token']}",
                   "Content-Type": "application/octet-stream"}

        hostname  = SP_SITE_URL.replace("https://", "").split("/")[0]
        site_path = "/".join(SP_SITE_URL.replace("https://", "").split("/")[1:])
        site_resp = requests.get(
            f"https://graph.microsoft.com/v1.0/sites/{hostname}:/{site_path}",
            headers={**headers, "Content-Type": "application/json"}, timeout=15)
        site_resp.raise_for_status()
        site_id = site_resp.json()["id"]

        drive_resp = requests.get(
            f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive",
            headers={**headers, "Content-Type": "application/json"}, timeout=15)
        drive_resp.raise_for_status()
        drive_id = drive_resp.json()["id"]

        with open(file_path, "rb") as f:
            data = f.read()
        r = requests.put(
            f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:{sp_path}:/content",
            headers=headers, data=data, timeout=60)
        r.raise_for_status()
        return True
    except Exception as e:
        print(f"Erro upload: {e}")
        return False

# ══════════════════════════════════════════════════════════════════════════════
# CARREGAMENTO DOS DADOS
# ══════════════════════════════════════════════════════════════════════════════
def _read_df(content: bytes, sheet: int = 0) -> pd.DataFrame:
    if content[:4] == b"PK\x03\x04":
        return pd.read_excel(io.BytesIO(content), sheet_name=sheet)
    return pd.read_csv(io.BytesIO(content))


@st.cache_data
def load_veiculos(content: bytes) -> pd.DataFrame:
    """
    Carrega Veículos AGRITEL para obter DeviceId, Nome e Grupo de cada máquina.

    Colunas esperadas (Agritel export — nomes aproximados):
      - ID do dispositivo  → DeviceId
      - Nome da máquina    → Nome
      - Grupo de máquinas  → Grupo
    """
    df = _read_df(content)

    id_col    = _find_col(df, ["id do dispositivo", "deviceid", "device id", "id"])
    nome_col  = _find_col(df, ["nome da maquina", "nome", "name", "maquina"])
    grupo_col = _find_col(df, ["grupo de maquinas", "grupo", "group", "tipo", "descri"])

    if id_col is None:
        id_col = df.columns[0]
    if nome_col is None:
        nome_col = df.columns[1] if len(df.columns) > 1 else df.columns[0]

    result = pd.DataFrame()
    result["DeviceId"] = df[id_col].astype(str).str.strip()
    result["Nome"]     = df[nome_col].astype(str).str.strip()
    result["Grupo"]    = df[grupo_col].astype(str).str.strip() if grupo_col else ""

    # Classificação do tipo com base em palavras-chave
    def _classify(row):
        texto = (_normalize(row["Nome"]) + " " + _normalize(row["Grupo"]))
        for kw in KEYWORDS_COLHEITA:
            if kw in texto:
                return "colheita"
        return "agro"

    result["Tipo"] = result.apply(_classify, axis=1)
    return result


@st.cache_data
def load_frotas_cc(content: bytes) -> pd.DataFrame:
    """
    Carrega Frotas_CentroDeCusto para obter o centro de custo e tipo de cada frota.

    Colunas esperadas:
      - Frota / Nome da Máquina  → Nome
      - Descrição                → Descricao
      - Centro de Custo          → CC
    """
    df = _read_df(content)

    nome_col  = _find_col(df, ["nome da maquina", "nome", "frota", "maquina", "name"])
    desc_col  = _find_col(df, ["descri", "tipo", "grupo", "category"])
    cc_col    = _find_col(df, ["centro de custo", "cc", "cost center", "centro"])

    if nome_col is None:
        nome_col = df.columns[0]

    result = pd.DataFrame()
    result["Nome"]     = df[nome_col].astype(str).str.strip()
    result["Descricao"] = df[desc_col].astype(str).str.strip() if desc_col else ""
    result["CC"]       = df[cc_col].astype(str).str.strip() if cc_col else ""

    # Classificar tipo pela descrição
    def _classify(row):
        texto = _normalize(row["Descricao"] + " " + row["Nome"])
        for kw in KEYWORDS_COLHEITA:
            if kw in texto:
                return "colheita"
        return "agro"

    result["Tipo"] = result.apply(_classify, axis=1)
    return result


def build_device_map(veiculos_df: pd.DataFrame | None) -> tuple[dict, dict]:
    """
    Retorna:
      device_map  : {DeviceId → nome_curto}   (ex: "Frota 1262")
      tipo_map    : {DeviceId → tipo}          (ex: "colheita" | "agro")
    """
    if veiculos_df is None or veiculos_df.empty:
        # fallback hardcoded
        device_map = DEVICE_MAP_FALLBACK.copy()
        tipo_map   = {k: "colheita" for k in DEVICE_MAP_FALLBACK}
        return device_map, tipo_map

    device_map = {}
    tipo_map   = {}
    for _, row in veiculos_df.iterrows():
        did  = str(row["DeviceId"])
        nome = str(row["Nome"])
        # Extrai "Frota XXXX" do nome, se existir
        m = re.search(r"[Ff]rota\s*(\d+)", nome)
        nome_curto = f"Frota {m.group(1)}" if m else nome
        device_map[did] = nome_curto
        tipo_map[did]   = row["Tipo"]

    return device_map, tipo_map


@st.cache_data
def load_apontamentos(content: bytes, device_map: dict, tipo_map: dict,
                      filtrar_tipo: str | None = None) -> pd.DataFrame:
    df = _read_df(content)
    df.columns = ["Apontamento", "DeviceId", "Codigo", "HorasTexto"]
    df["Horas"]  = df["HorasTexto"].apply(parse_horas)
    df["Codigo"] = df["Codigo"].astype(str).replace("10.004", "10.003")
    df = df.groupby(["Apontamento", "DeviceId", "Codigo"], as_index=False)["Horas"].sum()
    df["Categoria"] = df["Apontamento"].map(CATEGORIA_MAP).fillna("Sem Apontamento")
    df["Frota"]     = df["DeviceId"].map(device_map).fillna(df["DeviceId"])
    df["Tipo"]      = df["DeviceId"].map(tipo_map).fillna("agro")

    if filtrar_tipo:
        df = df[df["Tipo"] == filtrar_tipo].copy()

    return df


@st.cache_data
def load_performance(content: bytes, frotas_colheita: set | None = None,
                     filtrar_tipo: str | None = None) -> pd.DataFrame:
    df = _read_df(content)
    c  = df.columns.tolist()
    def col(name):
        return df[name].apply(parse_num) if name in c else pd.Series(0.0, index=df.index)

    result = pd.DataFrame({
        "NomeFrota":         df["Nome da Máquina"].str.extract(r"(Frota \d+)")[0].fillna(df["Nome da Máquina"]),
        "NomeCompleto":      df["Nome da Máquina"].astype(str),
        "Frota":             df["Nome da Máquina"].str.extract(r"Frota (\d+)")[0],
        "TempoMotorLigado":  df["Tempo Motor Ligado:"].apply(parse_tempo_perf),
        "TempoTrabalho":     df["Tempo em Trabalho"].apply(parse_tempo_perf),
        "TempoManobra":      df["Tempo em Manobra:"].apply(parse_tempo_perf),
        "TempoOcioso":       df["Tempo com Motor Ocioso:"].apply(parse_tempo_perf),
        "TempoElevador":     df["Tempo de Elevador Ligado"].apply(parse_tempo_perf),
        "ConsumoTotal":      col("Consumo Total:"),
        "ConsumoTrabalho":   col("Consumo em Trabalho:"),
        "TaxaMedConsumo":    col("Taxa Média de Consumo de Combustível"),
        "TaxaMedTrab":       col("Taxa Média de Consumo de Combustível em Trabalho"),
        "CargaMax":          col("Carga Máxima do Motor"),
        "CargaMed":          col("Carga Média do Motor"),
        "TempMinMotor":      col("Temperatura Mínima do Motor"),
        "TempMaxMotor":      col("Temperatura Máxima do Motor"),
        "TempMedMotor":      col("Temperatura Média do Motor"),
        "TempMinOleo":       col("Temperatura Mínima do Óleo Hidráulico"),
        "TempMaxOleo":       col("Temperatura Máxima do Óleo Hidráulico"),
        "TempMedOleo":       col("Temperatura Média do Óleo Hidráulico"),
        "Odometro":          col("Odômetro de Trabalho"),
        "VelocidadeMax":     col("Velocidade Máxima de Trabalho"),
        "VelocidadeMed":     col("Velocidade Média de Trabalho:"),
        "RpmMotor":          col("Rpm Médio em Trabalho:"),
        "RpmExtMed":         col("Rpm Médio de Exaustor Primário em Trabalho"),
        "PressaoMed":        col("Pressão Média de Corte Base em Trabalho"),
        "ReversoIndustrial": col("Número de Embuchamentos Detectados"),
    })

    # Classificar tipo baseado no nome completo
    if filtrar_tipo and frotas_colheita is not None:
        # frotas_colheita é um set de NomeFrota (ex: {"Frota 1262", "Frota 1263"})
        if filtrar_tipo == "colheita":
            result = result[result["NomeFrota"].isin(frotas_colheita)].copy()
        else:
            result = result[~result["NomeFrota"].isin(frotas_colheita)].copy()

    return result


# ══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def calc_disponibilidade(df: pd.DataFrame) -> pd.DataFrame:
    total = df.groupby("Frota")["Horas"].sum().rename("Total")
    quebra = df[df["Apontamento"] == "Quebra"].groupby("Frota")["Horas"].sum().rename("Quebra")
    disp = pd.concat([total, quebra], axis=1).fillna(0)
    disp["Disponibilidade"]   = (disp["Total"] - disp["Quebra"]) / disp["Total"] * 100
    disp["Indisponibilidade"] = disp["Quebra"] / disp["Total"] * 100
    return disp.reset_index()


FIG_LAYOUT = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font=dict(family="Inter, sans-serif", size=12, color="#e0e0e0"),
)

# ══════════════════════════════════════════════════════════════════════════════
# GRÁFICOS
# ══════════════════════════════════════════════════════════════════════════════
def fig_donut(df: pd.DataFrame) -> go.Figure:
    agg = df.groupby("Categoria")["Horas"].sum().reset_index()
    agg = agg[agg["Horas"] > 0].sort_values("Horas", ascending=False)
    fig = go.Figure(go.Pie(
        labels=agg["Categoria"], values=agg["Horas"].round(1),
        hole=0.60,
        marker=dict(colors=[CAT_CORES.get(c, "#888") for c in agg["Categoria"]],
                    line=dict(color="rgba(0,0,0,0.3)", width=1)),
        textinfo="percent", textfont=dict(color="#ffffff"),
        hovertemplate="<b>%{label}</b><br>%{value:.1f}h (%{percent})<extra></extra>",
    ))
    fig.update_layout(**FIG_LAYOUT, showlegend=True,
        legend=dict(orientation="v", x=1.0, y=0.5, font=dict(color="#e0e0e0")),
        height=280, margin=dict(l=4, r=4, t=4, b=4))
    return fig


def fig_barras_h(df: pd.DataFrame) -> go.Figure:
    agg = df.groupby(["Apontamento", "Categoria"])["Horas"].sum().reset_index()
    agg = agg[agg["Horas"] > 0].sort_values("Horas")
    max_label = max((len(str(a)) for a in agg["Apontamento"]), default=10) * 7
    fig = go.Figure(go.Bar(
        x=agg["Horas"].round(1), y=agg["Apontamento"],
        orientation="h",
        marker_color=[APT_CORES.get(a, "#888") for a in agg["Apontamento"]],
        text=agg["Horas"].round(1).astype(str) + "h",
        textposition="outside", textfont=dict(color="#e0e0e0"),
        hovertemplate="<b>%{y}</b><br>%{x:.1f}h<extra></extra>",
    ))
    fig.update_layout(**FIG_LAYOUT,
        height=max(300, len(agg) * 30 + 60),
        margin=dict(l=max_label + 20, r=60, t=10, b=10),
        xaxis=dict(showgrid=True, gridcolor="rgba(255,255,255,0.1)",
                   title="horas", color="#e0e0e0", zerolinecolor="rgba(255,255,255,0.2)"),
        yaxis=dict(showgrid=False, color="#e0e0e0"),
    )
    return fig


def fig_comparativo(df: pd.DataFrame) -> go.Figure:
    cats   = ["Colhendo", "Aguardando Transbordo", "Chuva / Umidade",
              "Quebra", "Manobra", "Sem Apontamento"]
    frotas = sorted(df["Frota"].unique())
    paleta = [C_F1262, C_F1263, C_ESPERA, C_MANOBRA, "#9b59b6", "#e67e22"]
    fig    = go.Figure()
    for i, frota in enumerate(frotas):
        sub         = df[df["Frota"] == frota]
        total_frota = sub["Horas"].sum()
        vals  = [round(sub[sub["Apontamento"] == c]["Horas"].sum(), 1) for c in cats]
        texts = [f"{v / total_frota * 100:.1f}%" if total_frota > 0 else "" for v in vals]
        fig.add_trace(go.Bar(
            name=frota, x=cats, y=vals,
            marker_color=paleta[i % len(paleta)],
            text=texts, textposition="outside",
            textfont=dict(color="#e0e0e0", size=9),
            hovertemplate="<b>%{x}</b><br>%{y:.1f}h · %{text}<extra>" + frota + "</extra>",
        ))
    fig.update_layout(**FIG_LAYOUT, barmode="group", height=350,
        xaxis=dict(tickangle=-20, showgrid=False, color="#e0e0e0"),
        margin=dict(l=4, r=4, t=36, b=4),
        yaxis=dict(showgrid=True, gridcolor="rgba(255,255,255,0.1)",
                   title="horas", color="#e0e0e0",
                   zerolinecolor="rgba(255,255,255,0.2)"),
        legend=dict(orientation="h", y=1.12, font=dict(color="#e0e0e0")),
    )
    return fig


def fig_perdas(df: pd.DataFrame) -> go.Figure:
    cats = ["Chuva / Umidade", "Sem Apontamento", "Quebra",
            "Aguardando Transbordo", "Manobra", "Manutenção Programada", "Abastecimento"]
    agg = df[df["Apontamento"].isin(cats)].groupby("Apontamento")["Horas"].sum()
    agg = agg[agg > 0].sort_values()
    if agg.empty:
        return go.Figure()
    max_label = max((len(str(a)) for a in agg.index), default=10) * 7
    fig = go.Figure(go.Bar(
        x=agg.values.round(1), y=agg.index, orientation="h",
        marker_color=[APT_CORES.get(a, "#888") for a in agg.index],
        text=[f"{v:.1f}h" for v in agg.values],
        textposition="outside", textfont=dict(color="#e0e0e0"),
        hovertemplate="<b>%{y}</b><br>%{x:.1f}h<extra></extra>",
    ))
    fig.update_layout(**FIG_LAYOUT,
        height=max(260, len(agg) * 38 + 60),
        margin=dict(l=max_label + 20, r=60, t=10, b=10),
        xaxis=dict(showgrid=True, gridcolor="rgba(255,255,255,0.1)",
                   title="horas", color="#e0e0e0", zerolinecolor="rgba(255,255,255,0.2)"),
        yaxis=dict(showgrid=False, color="#e0e0e0"),
    )
    return fig


def fig_disponibilidade(disp_df: pd.DataFrame) -> go.Figure:
    fig = go.Figure()
    fig.add_trace(go.Bar(
        name="Disponível", x=disp_df["Frota"],
        y=disp_df["Disponibilidade"].round(1),
        marker_color=C_DISP,
        text=disp_df["Disponibilidade"].round(1).astype(str) + "%",
        textposition="inside", textfont=dict(color="#ffffff"),
        hovertemplate="<b>%{x}</b><br>Disponibilidade: %{y:.1f}%<extra></extra>",
    ))
    fig.add_trace(go.Bar(
        name="Quebra", x=disp_df["Frota"],
        y=disp_df["Indisponibilidade"].round(1),
        marker_color=C_MANUTENCAO,
        text=disp_df["Indisponibilidade"].round(1).astype(str) + "%",
        textposition="inside", textfont=dict(color="#ffffff"),
        hovertemplate="<b>%{x}</b><br>Quebra: %{y:.1f}%<extra></extra>",
    ))
    fig.update_layout(**FIG_LAYOUT, barmode="stack", height=260,
        xaxis=dict(showgrid=False, color="#e0e0e0"),
        margin=dict(l=4, r=4, t=4, b=4),
        yaxis=dict(showgrid=True, gridcolor="rgba(255,255,255,0.1)",
                   title="%", range=[0, 105], color="#e0e0e0",
                   zerolinecolor="rgba(255,255,255,0.2)"),
        legend=dict(orientation="h", y=1.1, font=dict(color="#e0e0e0")),
    )
    return fig


def fig_horas_motor_elevador(df_perf: pd.DataFrame) -> go.Figure:
    """Gráfico de barras: Hora Motor Ligado vs Hora Elevador por frota."""
    if df_perf.empty:
        return go.Figure()
    fig = go.Figure()
    fig.add_trace(go.Bar(
        name="Motor Ligado", x=df_perf["NomeFrota"],
        y=df_perf["TempoMotorLigado"].round(1),
        marker_color=C_MOTOR,
        text=df_perf["TempoMotorLigado"].round(1).astype(str) + "h",
        textposition="inside", textfont=dict(color="#ffffff"),
        hovertemplate="<b>%{x}</b><br>Motor: %{y:.1f}h<extra></extra>",
    ))
    fig.add_trace(go.Bar(
        name="Elevador (Trabalho)", x=df_perf["NomeFrota"],
        y=df_perf["TempoElevador"].round(1),
        marker_color=C_ELEVADOR,
        text=df_perf["TempoElevador"].round(1).astype(str) + "h",
        textposition="inside", textfont=dict(color="#ffffff"),
        hovertemplate="<b>%{x}</b><br>Elevador: %{y:.1f}h<extra></extra>",
    ))
    fig.update_layout(**FIG_LAYOUT, barmode="group", height=280,
        xaxis=dict(showgrid=False, color="#e0e0e0"),
        margin=dict(l=4, r=4, t=4, b=4),
        yaxis=dict(showgrid=True, gridcolor="rgba(255,255,255,0.1)",
                   title="horas", color="#e0e0e0",
                   zerolinecolor="rgba(255,255,255,0.2)"),
        legend=dict(orientation="h", y=1.12, font=dict(color="#e0e0e0")),
    )
    return fig


def fig_temperaturas(row: pd.Series) -> go.Figure:
    labels = ["Motor", "Óleo Hidráulico"]
    meds   = [row["TempMedMotor"], row["TempMedOleo"]]
    mins_  = [row["TempMinMotor"], row["TempMinOleo"]]
    maxs_  = [row["TempMaxMotor"], row["TempMaxOleo"]]
    cores  = ["#E24B4A", "#378ADD"]
    fig    = go.Figure()
    for label, med, mn, mx, cor in zip(labels, meds, mins_, maxs_, cores):
        fig.add_trace(go.Bar(
            name=label, x=[label], y=[med], marker_color=cor,
            text=f"mín {mn:.0f}° · méd {med:.0f}° · máx {mx:.0f}°",
            textposition="inside", textfont=dict(color="#ffffff", size=12),
            error_y=dict(type="data", symmetric=False,
                         array=[mx - med], arrayminus=[med - mn]),
            hovertemplate=f"<b>{label}</b><br>min {mn}°C · méd {med}°C · máx {mx}°C<extra></extra>",
        ))
    fig.update_layout(**FIG_LAYOUT, height=260, showlegend=False,
        margin=dict(l=55, r=20, t=20, b=10),
        yaxis=dict(title="°C", showgrid=True, gridcolor="rgba(255,255,255,0.1)",
                   color="#e0e0e0", zerolinecolor="rgba(255,255,255,0.2)"),
        xaxis=dict(showgrid=False, color="#e0e0e0"),
    )
    return fig


def fig_gauge(valor: float, titulo: str, maximo: float = 100,
              cor: str = C_PRODUTIVO, sufixo: str = "%") -> go.Figure:
    fig = go.Figure(go.Indicator(
        mode="gauge+number", value=round(valor, 1),
        number=dict(suffix=sufixo, font=dict(size=22, color=cor)),
        gauge=dict(
            axis=dict(range=[0, maximo], tickwidth=1,
                      tickcolor="rgba(255,255,255,0.3)", tickfont=dict(color="#e0e0e0")),
            bar=dict(color=cor, thickness=0.6),
            bgcolor="rgba(255,255,255,0.05)", borderwidth=0,
            steps=[dict(range=[0, maximo], color="rgba(255,255,255,0.05)")],
        ),
        title=dict(text=titulo, font=dict(size=12, color="#aaaaaa")),
    ))
    fig.update_layout(**FIG_LAYOUT, height=200, margin=dict(l=20, r=20, t=40, b=20))
    return fig


# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown(f"## {CONF['icone']} COA · TESTON")
    st.markdown(f"**{CONF['subtitulo']}** · Safra 2026/27")
    st.markdown("---")

    st.markdown("### Dados")
    fonte = st.radio(
        "Fonte",
        ["SharePoint (automático)", "Upload manual"],
        label_visibility="collapsed",
    )

    sp_data    = {}
    apt_bytes  = None
    perf_bytes = None
    periodo_img_bytes = None

    if fonte == "SharePoint (automático)":
        with st.spinner("Conectando ao SharePoint..."):
            sp_data = load_from_sharepoint()
        apt_bytes         = sp_data.get("apt")
        perf_bytes        = sp_data.get("perf")
        periodo_img_bytes = sp_data.get("periodo_img")
        if apt_bytes and perf_bytes:
            st.success("✓ SharePoint conectado")
        else:
            st.error("Falha ao carregar dados principais.")
    else:
        up_apt   = st.file_uploader("XLSX — Apontamentos",      type=["xlsx","csv"], key="up_apt")
        up_perf  = st.file_uploader("XLSX — Performance",       type=["xlsx","csv"], key="up_perf")
        up_veic  = st.file_uploader("XLSX — Veículos AGRITEL",  type=["xlsx","csv"], key="up_veic")
        up_cc    = st.file_uploader("XLSX — Frotas CC",         type=["xlsx","csv"], key="up_cc")
        apt_bytes  = up_apt.read()  if up_apt  else None
        perf_bytes = up_perf.read() if up_perf else None
        if up_veic:
            sp_data["veiculos"] = up_veic.read()
        if up_cc:
            sp_data["frotas_cc"] = up_cc.read()

    if not apt_bytes or not perf_bytes:
        st.warning("Aguardando os arquivos de dados.")
        st.stop()

    # Construir mapeamentos de dispositivos
    veiculos_df   = load_veiculos(sp_data["veiculos"])   if sp_data.get("veiculos")  else None
    frotas_cc_df  = load_frotas_cc(sp_data["frotas_cc"]) if sp_data.get("frotas_cc") else None
    device_map, tipo_map = build_device_map(veiculos_df)

    # Frotas do tipo "colheita" para filtrar performance
    frotas_colheita = {v for k, v in device_map.items() if tipo_map.get(k) == "colheita"}

    df_apt  = load_apontamentos(apt_bytes,  device_map, tipo_map, filtrar_tipo=APP_TIPO)
    df_perf = load_performance(perf_bytes,  frotas_colheita, filtrar_tipo=APP_TIPO)
    disp_df = calc_disponibilidade(df_apt)

    # Informação do período
    st.markdown("---")
    st.markdown("### Período")
    try:
        ini_dt = datetime.strptime(PERIODO_INICIO, "%Y-%m-%d %H:%M")
        fim_dt = datetime.strptime(PERIODO_FIM,    "%Y-%m-%d %H:%M")
        st.markdown(
            f"**{ini_dt.strftime('%d/%m/%Y %H:%M')}**  \n"
            f"até **{fim_dt.strftime('%d/%m/%Y %H:%M')}**"
        )
        st.caption(f"Referência: **{PERIOD_H:.0f}h/máquina**")
    except Exception:
        st.caption(f"Período: {PERIODO_INICIO} – {PERIODO_FIM}")

    # Mostrar imagem do período se disponível
    if periodo_img_bytes:
        with st.expander("📋 Configuração Agritel"):
            st.image(periodo_img_bytes, use_container_width=True)

    st.markdown("---")
    st.markdown("### Filtros")
    frotas_disp = ["Todas"] + sorted(df_apt["Frota"].unique().tolist())
    frota_sel   = st.selectbox("Frota", frotas_disp)

    st.markdown("---")
    st.markdown("### Navegação")
    pagina = st.radio(
        "Página",
        ["Visão geral", "Apontamentos", "Telemetria", "Por frota"],
        label_visibility="collapsed",
    )

    st.markdown("---")
    if st.button("🚪 Sair", use_container_width=True):
        st.session_state["autenticado"] = False
        st.rerun()

    st.caption(f"MS Colheitas · TESTON · {CONF['subtitulo']} · 2026")

# ══════════════════════════════════════════════════════════════════════════════
# DADOS FILTRADOS
# ══════════════════════════════════════════════════════════════════════════════
df_f      = df_apt  if frota_sel == "Todas" else df_apt[df_apt["Frota"] == frota_sel]
df_perf_f = df_perf if frota_sel == "Todas" else df_perf[df_perf["NomeFrota"] == frota_sel]

num_frotas  = df_f["Frota"].nunique() if not df_f.empty else 1
total_h     = df_f["Horas"].sum()                                            # total apontado (todas as máquinas)
period_ref  = PERIOD_H * num_frotas                                          # período × nº de máquinas

col_h   = df_f[df_f["Apontamento"] == "Colhendo"]["Horas"].sum()
man_h   = df_f[df_f["Apontamento"] == "Manobra"]["Horas"].sum()
sa_h    = df_f[df_f["Categoria"] == "Sem Apontamento"]["Horas"].sum()
qbr_h   = df_f[df_f["Apontamento"] == "Quebra"]["Horas"].sum()
chuva_h = df_f[df_f["Apontamento"] == "Chuva / Umidade"]["Horas"].sum()
agt_h   = df_f[df_f["Apontamento"] == "Aguardando Transbordo"]["Horas"].sum()
improd  = chuva_h + qbr_h + agt_h

ef_col   = col_h / total_h * 100              if total_h > 0         else 0
ef_man   = man_h / (col_h + man_h) * 100      if (col_h + man_h) > 0 else 0
ef_imp   = improd / total_h * 100             if total_h > 0         else 0
ef_sa    = sa_h / total_h * 100               if total_h > 0         else 0
disp_m   = (total_h - qbr_h) / total_h * 100 if total_h > 0         else 0
ef_chuva = chuva_h / total_h * 100            if total_h > 0         else 0
ef_qbr   = qbr_h   / total_h * 100            if total_h > 0         else 0
ef_agt   = agt_h   / total_h * 100            if total_h > 0         else 0

# Telemetria agregada
motor_ligado_total = df_perf_f["TempoMotorLigado"].sum() if not df_perf_f.empty else 0.0
elevador_total     = df_perf_f["TempoElevador"].sum()    if not df_perf_f.empty else 0.0
ef_trabalho        = (elevador_total / motor_ligado_total * 100
                      if motor_ligado_total > 0 else 0.0)

# ══════════════════════════════════════════════════════════════════════════════
# BANNER DO PERÍODO
# ══════════════════════════════════════════════════════════════════════════════
def render_periodo_banner():
    """Exibe um banner informativo com o período de referência."""
    try:
        ini_dt = datetime.strptime(PERIODO_INICIO, "%Y-%m-%d %H:%M")
        fim_dt = datetime.strptime(PERIODO_FIM,    "%Y-%m-%d %H:%M")
        label = (f"📅 Período: {ini_dt.strftime('%d/%m/%Y %H:%M')} – "
                 f"{fim_dt.strftime('%d/%m/%Y %H:%M')} · "
                 f"<b>{PERIOD_H:.0f}h/máquina</b>")
    except Exception:
        label = f"📅 Período configurado: {PERIODO_INICIO} – {PERIODO_FIM}"

    frota_label = "Todas as frotas" if frota_sel == "Todas" else frota_sel
    maquinas_label = (f"{num_frotas} máquina{'s' if num_frotas > 1 else ''} · "
                      f"{period_ref:.0f}h referência total")
    st.markdown(
        f'<div class="periodo-badge">{label} &nbsp;|&nbsp; {CONF["icone"]} '
        f'{frota_label} — {maquinas_label}</div>',
        unsafe_allow_html=True,
    )


# ════════════════════════════════════════════════════════════════════════════
# PÁGINA 1 — VISÃO GERAL
# ════════════════════════════════════════════════════════════════════════════
if pagina == "Visão geral":
    st.markdown(f"## {CONF['icone']} Visão geral — {CONF['subtitulo']}")
    render_periodo_banner()

    # ── Linha 1: KPIs de tempo e eficiência ──────────────────────────────────
    c1, c2, c3, c4, c5 = st.columns(5)

    # PERÍODO DE REFERÊNCIA (corrigido: mostra as horas do período, não a soma)
    c1.metric(
        "⏱️ Período (por máquina)",
        f"{PERIOD_H:.0f}h",
        f"Total apontado: {total_h:.0f}h",
    )
    c2.metric("✅ Disponibilidade",   f"{disp_m:.1f}%",  f"{qbr_h:.1f}h em quebra")
    c3.metric("⚡ Efic. colheita",    f"{ef_col:.1f}%",  f"{col_h:.1f}h colhendo")
    c4.metric("↔️ % Manobra",         f"{ef_man:.1f}%",  f"{man_h:.1f}h de manobra")
    c5.metric("❓ Sem apontamento",    f"{ef_sa:.1f}%",   f"{sa_h:.1f}h sem registro")

    # ── Linha 2: Hora Motor + Hora Elevador + Eficiência de Trabalho ─────────
    st.markdown(
        "<p style='margin:14px 0 6px;font-size:11px;color:#888;"
        "text-transform:uppercase;letter-spacing:.08em'>⚙️ Motor & Trabalho</p>",
        unsafe_allow_html=True,
    )
    cm1, cm2, cm3, cm4 = st.columns(4)
    cm1.metric("🔵 Hora Motor Ligado",
               f"{motor_ligado_total:.1f}h",
               f"{motor_ligado_total/num_frotas:.1f}h/máquina" if num_frotas > 0 else "")
    cm2.metric("🟣 Hora Elevador (Trab.)",
               f"{elevador_total:.1f}h",
               f"{elevador_total/num_frotas:.1f}h/máquina" if num_frotas > 0 else "")
    cm3.metric("📊 Efic. Trabalho (Elev./Motor)",
               f"{ef_trabalho:.1f}%",
               f"Elevador ÷ Motor Ligado")
    cm4.metric("🔄 Máquinas no período",
               f"{num_frotas}",
               f"{total_h:.0f}h apontadas total")

    # ── Linha 3: Breakdown Improdutivo ───────────────────────────────────────
    st.markdown(
        "<p style='margin:14px 0 6px;font-size:11px;color:#888;"
        "text-transform:uppercase;letter-spacing:.08em'>⚠️ Breakdown Improdutivo</p>",
        unsafe_allow_html=True,
    )
    ci0, ci1, ci2, ci3 = st.columns(4)
    ci0.metric("⚠️ Improdutivo total",     f"{ef_imp:.1f}%",   f"{improd:.1f}h acumuladas")
    ci1.metric("🌧️ Chuva / Umidade",      f"{ef_chuva:.1f}%", f"{chuva_h:.1f}h")
    ci2.metric("🔧 Quebra",                f"{ef_qbr:.1f}%",   f"{qbr_h:.1f}h")
    ci3.metric("⏳ Aguardando Transbordo", f"{ef_agt:.1f}%",   f"{agt_h:.1f}h")

    st.divider()

    # ── Performance agregada ─────────────────────────────────────────────────
    if not df_perf_f.empty:
        cons_med  = df_perf_f["TaxaMedTrab"].mean()
        vel_med   = df_perf_f["VelocidadeMed"].mean()
        carga_med = df_perf_f["CargaMed"].mean()
        rev_total = int(df_perf_f["ReversoIndustrial"].sum())
        c6, c7, c8, c9 = st.columns(4)
        c6.metric("🛢️ Consumo médio trab.", f"{cons_med:.1f} L/h")
        c7.metric("🚜 Velocidade média",     f"{vel_med:.2f} km/h")
        c8.metric("⚙️ Carga média motor",    f"{carga_med:.1f}%")
        c9.metric("🔴 Reverso Industrial",   str(rev_total))

    st.divider()

    # ── Gráfico Motor × Elevador ─────────────────────────────────────────────
    if not df_perf_f.empty:
        st.markdown("**Hora Motor Ligado vs Hora Elevador por frota**")
        st.plotly_chart(fig_horas_motor_elevador(df_perf_f),
                        width="stretch", config={"displayModeBar": False})
        st.divider()

    st.markdown("**Disponibilidade mecânica por frota**")
    st.plotly_chart(fig_disponibilidade(disp_df),
                    width="stretch", config={"displayModeBar": False})

    st.divider()

    col_l, col_r = st.columns(2)
    with col_l:
        st.markdown("**Distribuição de horas por categoria**")
        st.plotly_chart(fig_donut(df_f), width="stretch",
                        config={"displayModeBar": False})
    with col_r:
        st.markdown("**Comparativo entre frotas**")
        frotas_totais   = df_apt.groupby("Frota")["Horas"].sum()
        total_geral_apt = df_apt["Horas"].sum()
        kpi_cols = st.columns(len(frotas_totais) + 1)
        kpi_cols[0].metric("⏱️ Total apontado", f"{total_geral_apt:.1f}h")
        for j, (fr, hr) in enumerate(frotas_totais.items()):
            kpi_cols[j + 1].metric(f"⏱️ {fr}", f"{hr:.1f}h")
        st.plotly_chart(fig_comparativo(df_apt), width="stretch",
                        config={"displayModeBar": False})

    st.markdown("**Perdas e paradas operacionais — horas acumuladas**")
    st.plotly_chart(fig_perdas(df_f), width="stretch",
                    config={"displayModeBar": False})


# ════════════════════════════════════════════════════════════════════════════
# PÁGINA 2 — APONTAMENTOS
# ════════════════════════════════════════════════════════════════════════════
elif pagina == "Apontamentos":
    st.markdown(f"## {CONF['icone']} Apontamentos por tipo")
    render_periodo_banner()

    c1, c2, c3 = st.columns(3)
    c1.metric("✅ Disponibilidade mecânica", f"{disp_m:.1f}%",
              f"{qbr_h:.1f}h em quebra · {total_h:.1f}h total apontado")
    c2.metric("⚡ Eficiência colheita", f"{ef_col:.1f}%", f"{col_h:.1f}h")
    c3.metric("❓ Sem apontamento",     f"{ef_sa:.1f}%",  f"{sa_h:.1f}h")

    st.divider()
    st.markdown("**Horas por apontamento**")
    st.plotly_chart(fig_barras_h(df_f), width="stretch",
                    config={"displayModeBar": False})

    st.divider()
    col_l, col_r = st.columns(2)
    with col_l:
        st.markdown("**Distribuição por categoria**")
        st.plotly_chart(fig_donut(df_f), width="stretch",
                        config={"displayModeBar": False})
    with col_r:
        st.markdown("**Resumo por categoria**")
        cats = df_f.groupby("Categoria")["Horas"].sum().reset_index()
        cats = cats[cats["Horas"] > 0].sort_values("Horas", ascending=False)
        cats["% do Total"] = (cats["Horas"] / total_h * 100).round(1).astype(str) + "%"
        cats["Horas"]      = cats["Horas"].round(1)
        max_h_cat          = cats["Horas"].max()
        for _, row in cats.iterrows():
            cor       = CAT_CORES.get(row["Categoria"], "#888")
            pct_barra = row["Horas"] / max_h_cat * 100
            st.markdown(
                f'<div style="margin-bottom:10px">'
                f'<div style="display:flex;justify-content:space-between;margin-bottom:4px">'
                f'<span style="font-size:13px;color:#ccc">{row["Categoria"]}</span>'
                f'<span style="font-size:14px;font-weight:700;color:#e0e0e0">{row["Horas"]}h '
                f'<span style="font-weight:400;color:#888;font-size:12px">({row["% do Total"]})</span></span>'
                f'</div>'
                f'<div style="background:rgba(255,255,255,0.08);border-radius:6px;height:18px;overflow:hidden">'
                f'<div style="width:{pct_barra:.1f}%;height:100%;background:{cor};border-radius:6px;'
                f'display:flex;align-items:center;padding-left:8px"></div>'
                f'</div></div>',
                unsafe_allow_html=True,
            )


# ════════════════════════════════════════════════════════════════════════════
# PÁGINA 3 — TELEMETRIA
# ════════════════════════════════════════════════════════════════════════════
elif pagina == "Telemetria":
    st.markdown(f"## {CONF['icone']} Telemetria — performance")
    render_periodo_banner()
    st.caption("Motor, consumo, temperatura, velocidade e elevador")

    if df_perf_f.empty:
        st.warning("Nenhum dado de performance para a frota selecionada.")
        st.stop()

    cols = st.columns(len(df_perf_f))
    for i, (_, row) in enumerate(df_perf_f.iterrows()):
        with cols[i]:
            st.markdown(f"**{row['NomeFrota']}**")
            d_row = disp_df[disp_df["Frota"] == row["NomeFrota"].replace("Frota ", "")]
            if not d_row.empty:
                disp_val = d_row.iloc[0]["Disponibilidade"]
                qbr_val  = d_row.iloc[0]["Quebra"]
                st.markdown(
                    f'<div style="padding:8px 12px;border-radius:8px;background:rgba(39,174,96,0.15);'
                    f'border-left:3px solid {C_DISP};margin-bottom:8px">'
                    f'<span style="font-size:11px;color:#aaa">Disponibilidade mecânica</span><br>'
                    f'<span style="font-size:20px;font-weight:700;color:{C_DISP}">{disp_val:.1f}%</span>'
                    f'<span style="font-size:10px;color:#888;margin-left:8px">({qbr_val:.1f}h em quebra)</span>'
                    f'</div>',
                    unsafe_allow_html=True,
                )

            # Eficiência de trabalho por frota
            ef_trab_frota = (row["TempoElevador"] / row["TempoMotorLigado"] * 100
                             if row["TempoMotorLigado"] > 0 else 0)
            st.markdown(
                f'<div style="padding:8px 12px;border-radius:8px;background:rgba(155,89,182,0.15);'
                f'border-left:3px solid {C_ELEVADOR};margin-bottom:8px">'
                f'<span style="font-size:11px;color:#aaa">Efic. Trabalho (Elev./Motor)</span><br>'
                f'<span style="font-size:20px;font-weight:700;color:{C_ELEVADOR}">{ef_trab_frota:.1f}%</span>'
                f'<span style="font-size:10px;color:#888;margin-left:8px">'
                f'({row["TempoElevador"]:.1f}h elev. / {row["TempoMotorLigado"]:.1f}h motor)</span>'
                f'</div>',
                unsafe_allow_html=True,
            )

            metricas = [
                ("Motor ligado",       f"{row['TempoMotorLigado']:.1f} h"),
                ("Em trabalho",        f"{row['TempoTrabalho']:.1f} h"),
                ("Elevador ligado",    f"{row['TempoElevador']:.1f} h"),
                ("Motor ocioso",       f"{row['TempoOcioso']:.1f} h"),
                ("Consumo total",      f"{int(row['ConsumoTotal']):,} L".replace(",", ".")),
                ("Taxa em trab.",      f"{row['TaxaMedTrab']:.1f} L/h"),
                ("Carga média",        f"{row['CargaMed']:.1f}%"),
                ("Vel. média",         f"{row['VelocidadeMed']:.2f} km/h"),
                ("RPM motor",          f"{int(row['RpmMotor'])} rpm"),
                ("RPM extrator",       f"{int(row['RpmExtMed'])} rpm"),
                ("Pressão méd.",       f"{row['PressaoMed']:.0f} psi"),
                ("Reverso Industrial", str(int(row["ReversoIndustrial"]))),
            ]
            for label, val in metricas:
                cor_val = C_MANUTENCAO if label == "Reverso Industrial" else (
                          C_ELEVADOR if label == "Elevador ligado" else (
                          C_MOTOR if label == "Motor ligado" else "#e0e0e0"))
                st.markdown(
                    f'<div style="display:flex;justify-content:space-between;padding:6px 10px;'
                    f'background:rgba(255,255,255,0.05);border-radius:6px;margin-bottom:4px">'
                    f'<span style="font-size:11px;color:#999">{label}</span>'
                    f'<span style="font-size:12px;font-weight:500;color:{cor_val}">{val}</span>'
                    f'</div>',
                    unsafe_allow_html=True,
                )

    st.divider()
    st.markdown("**Temperaturas — mín / méd / máx**")
    tcols = st.columns(len(df_perf_f))
    for i, (_, row) in enumerate(df_perf_f.iterrows()):
        with tcols[i]:
            st.markdown(f"*{row['NomeFrota']}*")
            st.plotly_chart(fig_temperaturas(row), width="stretch",
                            config={"displayModeBar": False})

    st.divider()
    st.markdown("**Gauges operacionais**")
    gcols = st.columns(len(df_perf_f) * 4)
    idx   = 0
    for _, row in df_perf_f.iterrows():
        ef_mot  = row["TempoTrabalho"]  / row["TempoMotorLigado"] * 100 if row["TempoMotorLigado"] > 0 else 0
        ocioso  = row["TempoOcioso"]    / row["TempoMotorLigado"] * 100 if row["TempoMotorLigado"] > 0 else 0
        ef_elev = row["TempoElevador"]  / row["TempoMotorLigado"] * 100 if row["TempoMotorLigado"] > 0 else 0
        with gcols[idx]:
            st.plotly_chart(fig_gauge(ef_mot, f"{row['NomeFrota']}\nEfic. motor", cor=C_PRODUTIVO),
                            width="stretch", config={"displayModeBar": False})
        with gcols[idx + 1]:
            st.plotly_chart(fig_gauge(ef_elev, f"{row['NomeFrota']}\nEfic. Elevador", cor=C_ELEVADOR),
                            width="stretch", config={"displayModeBar": False})
        with gcols[idx + 2]:
            st.plotly_chart(fig_gauge(row["CargaMed"], f"{row['NomeFrota']}\nCarga motor", cor=C_ESPERA),
                            width="stretch", config={"displayModeBar": False})
        with gcols[idx + 3]:
            st.plotly_chart(fig_gauge(ocioso, f"{row['NomeFrota']}\nOcioso", cor=C_MANUTENCAO),
                            width="stretch", config={"displayModeBar": False})
        idx += 4


# ════════════════════════════════════════════════════════════════════════════
# PÁGINA 4 — POR FROTA
# ════════════════════════════════════════════════════════════════════════════
elif pagina == "Por frota":
    frotas_lista = sorted(df_apt["Frota"].unique().tolist())
    frota_4 = st.selectbox("Selecione a frota", frotas_lista, key="frota4")

    sub_apt  = df_apt[df_apt["Frota"] == frota_4]
    sub_perf = df_perf[df_perf["NomeFrota"] == frota_4] if not df_perf.empty else pd.DataFrame()
    sub_disp = disp_df[disp_df["Frota"] == frota_4.replace("Frota ", "")]

    st.markdown(f"## {CONF['icone']} {frota_4} — Análise individual")
    render_periodo_banner()
    st.caption("Todos os indicadores desta frota")

    tot_f  = sub_apt["Horas"].sum()
    col_f  = sub_apt[sub_apt["Apontamento"] == "Colhendo"]["Horas"].sum()
    man_f  = sub_apt[sub_apt["Apontamento"] == "Manobra"]["Horas"].sum()
    qbr_f  = sub_apt[sub_apt["Apontamento"] == "Quebra"]["Horas"].sum()
    disp_f = (tot_f - qbr_f) / tot_f * 100 if tot_f > 0 else 0

    kc1, kc2, kc3, kc4 = st.columns(4)
    kc1.metric("✅ Disponibilidade",
               f"{disp_f:.1f}%",
               f"{qbr_f:.1f}h quebra · {tot_f:.1f}h total")
    kc2.metric("⚡ Efic. colheita",
               f"{col_f/tot_f*100:.1f}%" if tot_f > 0 else "–",
               f"{col_f:.1f}h colhendo")
    kc3.metric("↔️ % Manobra",
               f"{man_f/(col_f+man_f)*100:.1f}%" if (col_f + man_f) > 0 else "–",
               f"{man_f:.1f}h")
    kc4.metric("⏱️ Período ref.",
               f"{PERIOD_H:.0f}h",
               f"{tot_f:.1f}h apontado nesta frota")

    if not sub_perf.empty:
        row = sub_perf.iloc[0]
        ef_trab = (row["TempoElevador"] / row["TempoMotorLigado"] * 100
                   if row["TempoMotorLigado"] > 0 else 0)

        st.divider()
        kc5, kc6, kc7, kc8 = st.columns(4)
        kc5.metric("🔵 Motor Ligado",
                   f"{row['TempoMotorLigado']:.1f}h")
        kc6.metric("🟣 Hora Elevador",
                   f"{row['TempoElevador']:.1f}h",
                   f"Efic. Trabalho: {ef_trab:.1f}%")
        kc7.metric("🛢️ Cons. trab.",
                   f"{row['TaxaMedTrab']:.1f} L/h",
                   f"{int(row['ConsumoTrabalho']):,} L total".replace(",", "."))
        kc8.metric("🔴 Reverso Industrial",
                   str(int(row["ReversoIndustrial"])), "detectados")

        st.divider()
        kc9, kc10, kc11, kc12 = st.columns(4)
        kc9.metric("🚜 Vel. média",   f"{row['VelocidadeMed']:.2f} km/h")
        kc10.metric("⚙️ RPM motor",    f"{int(row['RpmMotor'])} rpm")
        oc = row["TempoOcioso"] / row["TempoMotorLigado"] * 100 if row["TempoMotorLigado"] > 0 else 0
        kc11.metric("⏸️ Motor ocioso",
                    f"{oc:.1f}%",
                    f"{row['TempoOcioso']:.1f}h de {row['TempoMotorLigado']:.0f}h")
        kc12.metric("⚙️ Carga média",  f"{row['CargaMed']:.1f}%")

    st.divider()
    cl, cr = st.columns(2)
    with cl:
        st.markdown("**Uso do tempo**")
        st.plotly_chart(fig_donut(sub_apt), width="stretch",
                        config={"displayModeBar": False})
    with cr:
        st.markdown("**Apontamentos detalhados**")
        st.plotly_chart(fig_barras_h(sub_apt), width="stretch",
                        config={"displayModeBar": False})

    if not sub_perf.empty:
        st.divider()
        st.markdown("**Temperaturas**")
        st.plotly_chart(fig_temperaturas(sub_perf.iloc[0]),
                        width="stretch", config={"displayModeBar": False})
        st.divider()
        tcol1, tcol2, tcol3, tcol4 = st.columns(4)
        row = sub_perf.iloc[0]
        ef_mot  = row["TempoTrabalho"] / row["TempoMotorLigado"] * 100 if row["TempoMotorLigado"] > 0 else 0
        ocioso  = row["TempoOcioso"]   / row["TempoMotorLigado"] * 100 if row["TempoMotorLigado"] > 0 else 0
        ef_elev = row["TempoElevador"] / row["TempoMotorLigado"] * 100 if row["TempoMotorLigado"] > 0 else 0
        with tcol1:
            st.plotly_chart(fig_gauge(ef_mot,  "Efic. Motor",    cor=C_PRODUTIVO), width="stretch", config={"displayModeBar": False})
        with tcol2:
            st.plotly_chart(fig_gauge(ef_elev, "Efic. Elevador", cor=C_ELEVADOR),  width="stretch", config={"displayModeBar": False})
        with tcol3:
            st.plotly_chart(fig_gauge(row["CargaMed"], "Carga Motor", cor=C_ESPERA), width="stretch", config={"displayModeBar": False})
        with tcol4:
            st.plotly_chart(fig_gauge(ocioso, "Ocioso", cor=C_MANUTENCAO), width="stretch", config={"displayModeBar": False})
