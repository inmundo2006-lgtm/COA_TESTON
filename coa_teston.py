"""
COA · TESTON — Dashboard Safra 2026/27  |  v6
Streamlit + Plotly | Login + SharePoint

NOVIDADES v6 (vs v5):
  - _load_apt_raw() lê novo formato: Máquina|Data e Horário|Apontamento|Código|Operador|Tempo(min)
    → pré-agrega para nível diário (211k linhas → ~2k), Tempo ÷ 60 = horas
  - Frotas_CentroDeCusto.xlsx é a única fonte de frente + tipo (colheita/agro)
  - AGRITEL_GRUPOS e _FROTA_FRENTE removidos — zero código hardcoded de frotas
  - filter_apt() e filter_perf() usam cc_map {frota_num: (frente, tipo)}
  - Sidebar: frentes vindas do cc_map dinamicamente

SECRETS necessários:
  SP_TENANT_ID, SP_CLIENT_ID, SP_CLIENT_SECRET
  SP_SITE_URL       = "https://metalcana.sharepoint.com/sites/MSCOLHEITA"
  SP_VEICULOS_PATH  = "/VITOR/AGRITEL/Veículos_Agritel.xlsx"
  SP_FROTAS_CC_PATH = "/VITOR/AGRITEL/Frotas_CentroDeCusto.xlsx"
  SP_APT_CONSOL     = "/VITOR/AGRITEL/APONTAMENTOS_CONSOLIDADO.xlsx"
  SP_PERF_CONSOL    = "/VITOR/AGRITEL/PERFORMANCE_CONSOLIDADO.xlsx"
  DASHBOARD_SENHA_HASH = "<sha256>"
"""

import re, io, os, hashlib
from datetime import datetime, date, timedelta
from urllib.parse import quote
import pandas as pd
import numpy as np
import streamlit as st
import plotly.graph_objects as go

# ══════════════════════════════════════════════════════════════════════════════
# CONFIGURAÇÕES
# ══════════════════════════════════════════════════════════════════════════════
def _s(key, default=""):
    try:    return st.secrets[key]
    except: return os.environ.get(key, default)

SP_TENANT_ID      = _s("SP_TENANT_ID")
SP_CLIENT_ID      = _s("SP_CLIENT_ID")
SP_CLIENT_SECRET  = _s("SP_CLIENT_SECRET")
SP_SITE_URL       = _s("SP_SITE_URL",       "https://metalcana.sharepoint.com/sites/MSCOLHEITA")
SP_VEICULOS_PATH  = _s("SP_VEICULOS_PATH",  "/VITOR/AGRITEL/Veículos_Agritel.xlsx")
SP_FROTAS_CC_PATH = _s("SP_FROTAS_CC_PATH", "/VITOR/AGRITEL/Frotas_CentroDeCusto.xlsx")
SP_APT_CONSOL     = _s("SP_APT_CONSOL",     "/VITOR/AGRITEL/APONTAMENTOS_CONSOLIDADO.xlsx")
SP_PERF_CONSOL    = _s("SP_PERF_CONSOL",    "/VITOR/AGRITEL/PERFORMANCE_CONSOLIDADO.xlsx")

SENHA_HASH = _s("DASHBOARD_SENHA_HASH", hashlib.sha256("teston2026".encode()).hexdigest())

TIPO_CFG = {
    "colheita": {"icone": "🌾", "titulo": "COA · TESTON — Colheita", "sub": "Colhedoras & Transbordos"},
    "agro":     {"icone": "🚜", "titulo": "COA · TESTON — Agro",     "sub": "Máquinas Agrícolas"},
}
APP_TIPO = st.session_state.get("tipo_operacao_val", "colheita")
CONF     = TIPO_CFG[APP_TIPO]

# ══════════════════════════════════════════════════════════════════════════════
# CORES
# ══════════════════════════════════════════════════════════════════════════════
C_PRODUTIVO  = "#1a7a4a"
C_MANUTENCAO = "#c0392b"
C_ESPERA     = "#c97d10"
C_MANOBRA    = "#7F77DD"
C_PARADA     = "#888780"
C_SEM_APT    = "#b4b2a9"
C_DISP       = "#27ae60"
C_ELEVADOR   = "#9b59b6"
C_MOTOR      = "#2980b9"

CATEGORIA_MAP = {
    "Colhendo": "Produtivo", "Trabalhando": "Produtivo",
    "Manobra": "Manobra", "Deslocando": "Manobra",
    "Quebra": "Manutenção", "Manutenção Programada": "Manutenção",
    "Troca Faquinha": "Manutenção", "PARADA OPERACIONAL": "Manutenção",
    "Chuva / Umidade": "Espera/Imprevisto", "Usina Parada": "Espera/Imprevisto",
    "Falta de Area Liberada": "Espera/Imprevisto", "Falta de Caminhão": "Espera/Imprevisto",
    "Aguardando Transbordo": "Espera/Imprevisto", "Aguardando Colhedora": "Espera/Imprevisto",
    "Abastecimento": "Parada Operacional", "Refeição": "Parada Operacional",
    "Checklist": "Parada Operacional", "Final Jornada de Trabalho": "Parada Operacional",
    "Início Jornada de Trabalho": "Parada Operacional", "Parado": "Parada Operacional",
}
CAT_CORES = {
    "Produtivo": C_PRODUTIVO, "Manutenção": C_MANUTENCAO,
    "Espera/Imprevisto": C_ESPERA, "Manobra": C_MANOBRA,
    "Parada Operacional": C_PARADA, "Sem Apontamento": C_SEM_APT,
}
APT_CORES = {
    "Colhendo": C_PRODUTIVO, "Trabalhando": C_PRODUTIVO,
    "Aguardando Transbordo": C_ESPERA, "Aguardando Colhedora": C_ESPERA,
    "Chuva / Umidade": "#1a5fa8", "Quebra": C_MANUTENCAO,
    "Manobra": C_MANOBRA, "Sem Apontamento": C_SEM_APT,
    "Manutenção Programada": "#2ea05f", "Abastecimento": "#5DCAA5",
    "Troca Faquinha": "#D85A30", "Checklist": C_PARADA,
    "Refeição": C_PARADA, "Deslocando": "#9b9b9b",
    "Final Jornada de Trabalho": "#9b9b9b", "Início Jornada de Trabalho": "#9b9b9b",
    "Falta de Area Liberada": C_MANUTENCAO, "Falta de Caminhão": C_MANUTENCAO,
    "PARADA OPERACIONAL": C_MANUTENCAO, "Usina Parada": "#9b9b9b",
    "Parado": C_PARADA,
}

# ══════════════════════════════════════════════════════════════════════════════
# PAGE CONFIG + CSS
# ══════════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="COA · TESTON", page_icon="🌿",
    layout="wide", initial_sidebar_state="expanded",
)
st.markdown("""
<style>
[data-testid="stMetricValue"]       {font-size:1.6rem!important;font-weight:700}
[data-testid="stMetricLabel"]       {font-size:.72rem!important;text-transform:uppercase;letter-spacing:.05em}
[data-testid="stMetricDelta"]       {font-size:.72rem!important}
[data-testid="stSidebar"]           {background-color:#1c1e18}
[data-testid="stSidebar"] * {color:#d8dbd4!important}
[data-testid="stSidebar"] h1,[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3        {color:#a8c832!important}
div[data-testid="metric-container"] {background:rgba(255,255,255,.05);border-radius:10px;
                                     padding:12px 16px;border-left:3px solid #1a7a4a}
.block-container{padding-top:1.4rem}
.periodo-badge{display:inline-flex;align-items:center;gap:8px;
  background:rgba(168,200,50,.12);border:1px solid rgba(168,200,50,.3);
  border-radius:8px;padding:6px 14px;margin-bottom:12px;font-size:12px;color:#a8c832}
.api-badge{display:inline-flex;align-items:center;gap:6px;font-size:11px;
  background:rgba(41,128,185,.12);border:1px solid rgba(41,128,185,.3);
  border-radius:6px;padding:4px 10px;color:#2980b9;margin-bottom:8px}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# LOGIN
# ══════════════════════════════════════════════════════════════════════════════
def check_login():
    if st.session_state.get("autenticado"):
        return
    st.markdown("""
    <div style="max-width:360px;margin:80px auto;text-align:center">
      <h1 style="color:#a8c832;font-size:2rem">🌿 COA · TESTON</h1>
      <p style="color:#888;margin-bottom:28px">Safra 2026/27</p>
    </div>""", unsafe_allow_html=True)
    col = st.columns([1,2,1])[1]
    with col:
        pw = st.text_input("Senha de acesso", type="password", placeholder="Digite a senha...")
        if st.button("Entrar", use_container_width=True, type="primary"):
            if hashlib.sha256(pw.encode()).hexdigest() == SENHA_HASH:
                st.session_state["autenticado"] = True
                st.rerun()
            else:
                st.error("Senha incorreta.")
    st.stop()

check_login()

# ══════════════════════════════════════════════════════════════════════════════
# PARSERS / HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def parse_num(v):
    if isinstance(v, (int, float)):
        return float(v) if not (isinstance(v, float) and np.isnan(v)) else 0.0
    c = re.sub(r"[^\d,.]", "", str(v)).replace(".", "").replace(",", ".")
    try:  return float(c)
    except: return 0.0

def _norm(s):
    s = str(s).lower()
    for a, b in [("á","a"),("ã","a"),("â","a"),("é","e"),("ê","e"),
                 ("í","i"),("ó","o"),("õ","o"),("ô","o"),("ú","u"),("ç","c")]:
        s = s.replace(a, b)
    return s

def _find_col(df, kws):
    nm = {_norm(c): c for c in df.columns}
    for kw in kws:
        for n, c in nm.items():
            if kw in n: return c
    return None

def _read_df(content: bytes) -> pd.DataFrame:
    if content[:4] == b"PK\x03\x04":
        return pd.read_excel(io.BytesIO(content), sheet_name=0)
    return pd.read_csv(io.BytesIO(content))

def _tipo_maq_from_name(name: str) -> str:
    """Fallback por palavra-chave — usa a mesma nomenclatura da aba 'cadastro'."""
    n = _norm(str(name))
    if "colhedora" in n:                                    return "Colhedora"
    if "pulveriz" in n:                                     return "Pulverizador"
    if "pa carregadeira" in n or "carregadeira" in n:       return "Pá Carregadeira"
    if "retroescavadeira" in n or "escavadeira" in n:       return "Retroescavadeira"
    if any(k in n for k in ["transbordo","caminhao","trator","reboque"]):
        return "Trator"
    return "Outro"

def _frota_num(frota_str: str) -> str:
    """'Frota 1262' → '1262'"""
    m = re.search(r"(\d+)", str(frota_str))
    return m.group(1) if m else ""

safe = lambda n, d: n / d * 100 if d > 0 else 0.0

# ══════════════════════════════════════════════════════════════════════════════
# SHAREPOINT
# ══════════════════════════════════════════════════════════════════════════════
@st.cache_data(ttl=300)
def load_from_sharepoint():
    result = {k: None for k in ("apt_consol","perf_consol","veiculos","frotas_cc")}
    try:
        import requests as req
        tok = req.post(
            f"https://login.microsoftonline.com/{SP_TENANT_ID}/oauth2/v2.0/token",
            data={"grant_type":"client_credentials","client_id":SP_CLIENT_ID,
                  "client_secret":SP_CLIENT_SECRET,
                  "scope":"https://graph.microsoft.com/.default"},
            timeout=15).json()
        h  = {"Authorization": f"Bearer {tok['access_token']}"}
        host  = SP_SITE_URL.replace("https://","").split("/")[0]
        spath = "/".join(SP_SITE_URL.replace("https://","").split("/")[1:])
        sid   = req.get(f"https://graph.microsoft.com/v1.0/sites/{host}:/{spath}",
                        headers=h, timeout=15).json()["id"]
        did   = req.get(f"https://graph.microsoft.com/v1.0/sites/{sid}/drive",
                        headers=h, timeout=15).json()["id"]

        def _dl(path):
            enc = quote(path, safe="/")
            try:
                r = req.get(
                    f"https://graph.microsoft.com/v1.0/drives/{did}/root:{enc}:/content",
                    headers=h, timeout=60)
                r.raise_for_status()
                return r.content
            except Exception as e:
                st.warning(f"⚠️ SharePoint: `{path}` — {e}")
                return None

        result["apt_consol"]  = _dl(SP_APT_CONSOL)
        result["perf_consol"] = _dl(SP_PERF_CONSOL)
        result["veiculos"]    = _dl(SP_VEICULOS_PATH)
        result["frotas_cc"]   = _dl(SP_FROTAS_CC_PATH)
    except Exception as e:
        st.error(f"❌ SharePoint: {e}")
    return result

# ══════════════════════════════════════════════════════════════════════════════
# CARREGAMENTO DOS ARQUIVOS BASE
# ══════════════════════════════════════════════════════════════════════════════
@st.cache_data
def load_frotas_cc(content: bytes) -> dict:
    """
    Lê Frotas_CentroDeCusto.xlsx.
    Retorna {frota_num: (frente, tipo)}
    ex: {'1262': ('RIO AMAMBAI - F. FIRU', 'colheita'), '1417': ('AGRO CIANORTE - F. CARMONA', 'agro')}
    """
    df = _read_df(content)
    frota_c = _find_col(df, ["frota"])
    cc_c    = _find_col(df, ["centro de custo","centro","cc"])
    if not frota_c: frota_c = df.columns[0]
    if not cc_c:    cc_c    = df.columns[1] if len(df.columns) > 1 else df.columns[0]

    result: dict[str, tuple] = {}
    for _, row in df.iterrows():
        frota_n = str(row[frota_c]).strip().lstrip("0") or "0"
        cc      = str(row[cc_c]).strip()
        cu      = cc.upper()
        tipo    = "agro" if (cu.startswith("AGRO") or cu.startswith("PREPARO")) else "colheita"
        result[frota_n] = (cc, tipo)
    return result


@st.cache_data
def load_veiculos(content: bytes) -> pd.DataFrame:
    df = _read_df(content)
    nome_c = _find_col(df, ["nome","name","maquina"])
    tipo_c = _find_col(df, ["tipo","type","grupo"])
    if not nome_c: nome_c = df.columns[0]
    return pd.DataFrame({
        "Nome":  df[nome_c].astype(str).str.strip(),
        "Grupo": df[tipo_c].astype(str).str.strip() if tipo_c else "",
    })


def load_tipo_maq_lookup(content: bytes) -> dict:
    """
    Lê a aba 'cadastro' do Veículos_Agritel.xlsx.
    Colunas: CHAVE | SAIDA
    Ex: {'Colhedora de Cana': 'Colhedora', 'Trator Transbordo': 'Trator', ...}
    """
    try:
        df = pd.read_excel(io.BytesIO(content), sheet_name="cadastro", dtype=str)
        chave_c = _find_col(df, ["chave","key"]) or df.columns[0]
        saida_c = _find_col(df, ["saida","output","tipo","resultado"]) or df.columns[1]
        return {
            str(k).strip(): str(v).strip()
            for k, v in zip(df[chave_c], df[saida_c])
            if str(k).strip() and str(k).strip() != "nan"
        }
    except Exception:
        return {}


def build_veic_tipomaq(veiculos_df, tipo_lookup: dict | None = None) -> dict:
    """
    Retorna {frota_num: tipo_maq} ex: {'1262': 'Colhedora', '1500': 'Trator'}
    tipo_lookup = dict da aba 'cadastro' do Veículos_Agritel: {CHAVE: SAIDA}
    """
    result: dict[str, str] = {}
    if veiculos_df is None or veiculos_df.empty:
        return result
    for _, row in veiculos_df.iterrows():
        m = re.search(r"[Ff]rota\s*(\d+)", str(row.get("Nome", "")))
        if m:
            fonte = str(row.get("Grupo", row.get("Nome", ""))).strip()
            # Tenta lookup exato da aba cadastro, fallback por palavra-chave
            tm = (tipo_lookup or {}).get(fonte) or _tipo_maq_from_name(fonte)
            result[m.group(1)] = tm
    return result

# ══════════════════════════════════════════════════════════════════════════════
# APONTAMENTOS — novo formato
# ══════════════════════════════════════════════════════════════════════════════
@st.cache_data
def _load_apt_raw(content: bytes) -> pd.DataFrame:
    """
    Lê APONTAMENTOS_CONSOLIDADO.xlsx no formato:
        Máquina | Data e Horário | Apontamento | Código | Operador | Tempo(min)

    Pré-agrega para nível diário:
        Data | Frota | Apontamento | Horas
    (~211k linhas → ~2k linhas em cache)
    """
    df = _read_df(content)
    cols_lc = {c.strip().lower(): c for c in df.columns}

    maq_c  = next((cols_lc[k] for k in ["máquina","maquina"]              if k in cols_lc), None)
    data_c = next((cols_lc[k] for k in ["data e horário","data e horario",
                                          "data e hora","data"]            if k in cols_lc), None)
    apt_c  = next((cols_lc[k] for k in ["apontamento"]                    if k in cols_lc), None)
    tmp_c  = next((cols_lc[k] for k in ["tempo"]                          if k in cols_lc), None)

    if not all([maq_c, data_c, apt_c, tmp_c]):
        st.error(f"Colunas não encontradas no APONTAMENTOS_CONSOLIDADO. "
                 f"Encontrado: {list(df.columns)}")
        return pd.DataFrame()

    # Data e Horário já vem como datetime (openpyxl converte automático)
    df[data_c] = pd.to_datetime(df[data_c], errors="coerce")

    r = pd.DataFrame({
        "Data":        df[data_c].dt.normalize(),              # apenas a data
        "Frota":       df[maq_c].astype(str).str.strip(),     # "Frota 1200"
        "Apontamento": df[apt_c].astype(str).str.strip(),
        "Horas":       pd.to_numeric(df[tmp_c], errors="coerce").fillna(0) / 60.0,
    })

    # Agrega por dia (reduz 211k → ~2k)
    r = (r.groupby(["Data","Frota","Apontamento"], as_index=False)["Horas"].sum())
    return r.dropna(subset=["Data"])


@st.cache_data
def _load_perf_raw(content: bytes) -> pd.DataFrame:
    df = _read_df(content)
    date_col = _find_col(df, ["data","date"])
    if date_col and df.columns[0] != date_col:
        cols = [date_col] + [c for c in df.columns if c != date_col]
        df = df[cols]
    df = df.rename(columns={df.columns[0]: "Data"})
    df["Data"] = pd.to_datetime(df["Data"], dayfirst=True, errors="coerce")
    return df.dropna(subset=["Data"])

# ══════════════════════════════════════════════════════════════════════════════
# FILTROS
# ══════════════════════════════════════════════════════════════════════════════
def filter_apt(df_raw: pd.DataFrame, data_ini: date, data_fim: date,
               cc_map: dict, veic_tipomaq: dict,
               filtrar_tipo: str,
               frentes_sel: list | None = None,
               tipos_maq_sel: list | None = None) -> pd.DataFrame:
    """
    Filtra e classifica apontamentos.
    cc_map = {frota_num: (frente, tipo)}  ← Frotas_CentroDeCusto
    """
    mask = (df_raw["Data"].dt.date >= data_ini) & (df_raw["Data"].dt.date <= data_fim)
    df = df_raw[mask].copy()
    if df.empty:
        return pd.DataFrame(columns=["Frota","Apontamento","Horas","Categoria",
                                     "FrotaNum","Tipo","Frente","TipoMaq"])

    df = df.groupby(["Frota","Apontamento"], as_index=False)["Horas"].sum()
    df["Categoria"] = df["Apontamento"].map(CATEGORIA_MAP).fillna("Sem Apontamento")

    def _classify(frota_str: str):
        fn = _frota_num(frota_str)
        if fn in cc_map:
            frente, tipo = cc_map[fn]
        else:
            frente, tipo = "Sem frente", "colheita"
        tipo_maq = veic_tipomaq.get(fn, "Colhedora" if tipo == "colheita" else "Trator")
        return fn, tipo, frente, tipo_maq

    cl             = df["Frota"].apply(_classify)
    df["FrotaNum"] = cl.apply(lambda x: x[0])
    df["Tipo"]     = cl.apply(lambda x: x[1])
    df["Frente"]   = cl.apply(lambda x: x[2])
    df["TipoMaq"]  = cl.apply(lambda x: x[3])

    df = df[df["Tipo"] == filtrar_tipo]
    if frentes_sel:   df = df[df["Frente"].isin(frentes_sel)]
    if tipos_maq_sel: df = df[
        df["TipoMaq"].isin(tipos_maq_sel) | (df["TipoMaq"] == "Outro")]
    return df


def filter_perf(df_raw: pd.DataFrame, data_ini: date, data_fim: date,
                cc_map: dict, veic_tipomaq: dict,
                filtrar_tipo: str,
                frentes_sel: list | None = None,
                tipos_maq_sel: list | None = None) -> pd.DataFrame:
    """
    Agrega performance por frota.
    Classificação via frota_num → cc_map (sem AGRITEL_GRUPOS).
    """
    mask = (df_raw["Data"].dt.date >= data_ini) & (df_raw["Data"].dt.date <= data_fim)
    df = df_raw[mask].copy()
    if df.empty:
        return pd.DataFrame()

    def col(name):
        return df[name].apply(parse_num) if name in df.columns else pd.Series(0.0, index=df.index)

    def _tempo(name):
        if name not in df.columns:
            return pd.Series(0.0, index=df.index)
        return df[name].apply(lambda v: parse_num(v) if isinstance(v, str)
                              else (float(v) if pd.notna(v) else 0.0))

    desl_col = _find_col(df, ["deslocamento","deslocando"])

    tmp = pd.DataFrame({
        "NomeFrota":         df["Nome da Máquina"].str.extract(r"(Frota \d+)")[0].fillna(df["Nome da Máquina"]),
        "DeviceId":          df["Nome da Máquina"].astype(str),
        "TempoMotorLigado":  _tempo("Tempo Motor Ligado:"),
        "TempoTrabalho":     _tempo("Tempo em Trabalho"),
        "TempoManobra":      _tempo("Tempo em Manobra:"),
        "TempoOcioso":       _tempo("Tempo com Motor Ocioso:"),
        "TempoElevador":     _tempo("Tempo de Elevador Ligado"),
        "TempoDeslocando":   _tempo(desl_col) if desl_col else pd.Series(0.0, index=df.index),
        "ConsumoTotal":      col("Consumo Total:"),
        "ConsumoTrabalho":   col("Consumo em Trabalho:"),
        "TaxaMedTrab":       col("Taxa Média de Consumo de Combustível em Trabalho"),
        "TaxaMedConsumo":    col("Taxa Média de Consumo de Combustível"),
        "TempMaxMotor":      col("Temperatura Máxima do Motor"),
        "TempMedMotor":      col("Temperatura Média do Motor"),
        "TempMinMotor":      col("Temperatura Mínima do Motor"),
        "TempMaxOleo":       col("Temperatura Máxima do Óleo Hidráulico"),
        "TempMedOleo":       col("Temperatura Média do Óleo Hidráulico"),
        "TempMinOleo":       col("Temperatura Mínima do Óleo Hidráulico"),
        "CargaMed":          col("Carga Média do Motor"),
        "CargaMax":          col("Carga Máxima do Motor"),
        "VelocidadeMed":     col("Velocidade Média de Trabalho:"),
        "VelocidadeMax":     col("Velocidade Máxima de Trabalho"),
        "RpmMotor":          col("Rpm Médio em Trabalho:"),
        "RpmExtMed":         col("Rpm Médio de Exaustor Primário em Trabalho"),
        "PressaoMed":        col("Pressão Média de Corte Base em Trabalho"),
        "ReversoIndustrial": col("Número de Embuchamentos Detectados"),
        "Odometro":          col("Odômetro de Trabalho"),
    })

    # Classificação via cc_map (frota_num → frente/tipo)
    def _classify_perf(nome_maquina):
        fn = _frota_num(str(nome_maquina))
        if fn in cc_map:
            frente, tipo = cc_map[fn]
        else:
            frente, tipo = "Sem frente", "colheita"
        # TipoMaq: veic_tipomaq primeiro, fallback por nome
        # "Outro" = desconhecido — não forçar tipo incorreto
        tm = veic_tipomaq.get(fn) or _tipo_maq_from_name(str(nome_maquina))
        return tipo, frente, tm

    cl             = tmp["DeviceId"].apply(_classify_perf)
    tmp["Tipo"]    = cl.apply(lambda x: x[0])
    tmp["Frente"]  = cl.apply(lambda x: x[1])
    tmp["TipoMaq"] = cl.apply(lambda x: x[2])

    tmp = tmp[tmp["Tipo"] == filtrar_tipo]
    if frentes_sel:   tmp = tmp[tmp["Frente"].isin(frentes_sel)]
    # "Outro" = tipo desconhecido → sempre exibe (não filtra fora)
    if tipos_maq_sel: tmp = tmp[
        tmp["TipoMaq"].isin(tipos_maq_sel) | (tmp["TipoMaq"] == "Outro")]
    if tmp.empty:
        return pd.DataFrame()

    agg_sum  = ["TempoMotorLigado","TempoTrabalho","TempoManobra","TempoOcioso",
                "TempoElevador","TempoDeslocando","ConsumoTotal","ConsumoTrabalho",
                "ReversoIndustrial","Odometro"]
    agg_max  = ["TempMaxMotor","TempMaxOleo","CargaMax","VelocidadeMax"]
    agg_mean = ["TaxaMedTrab","TaxaMedConsumo","TempMedMotor","TempMedOleo",
                "TempMinMotor","TempMinOleo","CargaMed","VelocidadeMed",
                "RpmMotor","RpmExtMed","PressaoMed"]

    result = tmp.groupby("NomeFrota").agg(
        **{c: (c,"sum")  for c in agg_sum  if c in tmp.columns},
        **{c: (c,"max")  for c in agg_max  if c in tmp.columns},
        **{c: (c,"mean") for c in agg_mean if c in tmp.columns},
    ).reset_index()

    for col_name in ["TempMedMotor","TempMaxMotor","TempMinMotor",
                     "TempMedOleo","TempMaxOleo","TempMinOleo","TempoDeslocando"]:
        if col_name not in result.columns:
            result[col_name] = 0.0

    result["Frota"]        = result["NomeFrota"].str.extract(r"Frota (\d+)")[0].fillna("")
    result["NomeCompleto"] = result["NomeFrota"]
    return result

# ══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def calc_disponibilidade(df: pd.DataFrame) -> pd.DataFrame:
    total = df.groupby("Frota")["Horas"].sum().rename("Total")
    qbr   = df[df["Apontamento"]=="Quebra"].groupby("Frota")["Horas"].sum().rename("Quebra")
    d = pd.concat([total, qbr], axis=1).fillna(0)
    d["Disponibilidade"]   = (d["Total"] - d["Quebra"]) / d["Total"] * 100
    d["Indisponibilidade"] = d["Quebra"] / d["Total"] * 100
    return d.reset_index()

FIG_L = dict(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
             font=dict(family="Inter,sans-serif", size=12, color="#e0e0e0"))

# ══════════════════════════════════════════════════════════════════════════════
# GRÁFICOS (inalterados desde v5)
# ══════════════════════════════════════════════════════════════════════════════
def fig_donut(df):
    agg = df.groupby("Categoria")["Horas"].sum().reset_index()
    agg = agg[agg["Horas"] > 0].sort_values("Horas", ascending=False)
    fig = go.Figure(go.Pie(
        labels=agg["Categoria"], values=agg["Horas"].round(1), hole=0.60,
        marker=dict(colors=[CAT_CORES.get(c,"#888") for c in agg["Categoria"]],
                    line=dict(color="rgba(0,0,0,0.3)", width=1)),
        textinfo="percent", textfont=dict(color="#fff"),
        hovertemplate="<b>%{label}</b><br>%{value:.1f}h (%{percent})<extra></extra>",
    ))
    fig.update_layout(**FIG_L, showlegend=True,
        legend=dict(orientation="v", x=1.0, y=0.5, font=dict(color="#e0e0e0")),
        height=280, margin=dict(l=4,r=4,t=4,b=4))
    return fig

def fig_barras_h(df):
    agg = df.groupby(["Apontamento","Categoria"])["Horas"].sum().reset_index()
    agg = agg[agg["Horas"] > 0].sort_values("Horas")
    ml  = max((len(str(a)) for a in agg["Apontamento"]), default=10) * 7
    fig = go.Figure(go.Bar(
        x=agg["Horas"].round(1), y=agg["Apontamento"], orientation="h",
        marker_color=[APT_CORES.get(a,"#888") for a in agg["Apontamento"]],
        text=agg["Horas"].round(1).astype(str)+"h",
        textposition="outside", textfont=dict(color="#e0e0e0"),
        hovertemplate="<b>%{y}</b><br>%{x:.1f}h<extra></extra>",
    ))
    fig.update_layout(**FIG_L, height=max(300, len(agg)*30+60),
        margin=dict(l=ml+20,r=60,t=10,b=10),
        xaxis=dict(showgrid=True,gridcolor="rgba(255,255,255,0.1)",title="horas",color="#e0e0e0"),
        yaxis=dict(showgrid=False,color="#e0e0e0"))
    return fig

def fig_comparativo(df):
    cats   = ["Colhendo","Trabalhando","Aguardando Transbordo","Chuva / Umidade",
              "Quebra","Manobra","Sem Apontamento"]
    frotas = sorted(df["Frota"].unique())
    pal    = ["#1a7a4a","#2ecc71","#1a5fa8","#c97d10","#7F77DD","#9b59b6","#e67e22"]
    fig    = go.Figure()
    for i, fr in enumerate(frotas):
        sub = df[df["Frota"]==fr]; tot = sub["Horas"].sum()
        vals  = [round(sub[sub["Apontamento"]==c]["Horas"].sum(), 1) for c in cats]
        texts = [f"{v/tot*100:.1f}%" if tot > 0 else "" for v in vals]
        fig.add_trace(go.Bar(name=fr, x=cats, y=vals, marker_color=pal[i%len(pal)],
            text=texts, textposition="outside", textfont=dict(color="#e0e0e0", size=9),
            hovertemplate="<b>%{x}</b><br>%{y:.1f}h · %{text}<extra>"+fr+"</extra>"))
    fig.update_layout(**FIG_L, barmode="group", height=350,
        xaxis=dict(tickangle=-20,showgrid=False,color="#e0e0e0"),
        margin=dict(l=4,r=4,t=36,b=4),
        yaxis=dict(showgrid=True,gridcolor="rgba(255,255,255,0.1)",title="horas",color="#e0e0e0"),
        legend=dict(orientation="h",y=1.12,font=dict(color="#e0e0e0")))
    return fig

def fig_perdas(df):
    cats = ["Chuva / Umidade","Sem Apontamento","Quebra","Aguardando Transbordo",
            "Aguardando Colhedora","Manobra","Manutenção Programada","Abastecimento"]
    agg  = df[df["Apontamento"].isin(cats)].groupby("Apontamento")["Horas"].sum()
    agg  = agg[agg > 0].sort_values()
    if agg.empty: return go.Figure()
    ml = max((len(str(a)) for a in agg.index), default=10) * 7
    fig = go.Figure(go.Bar(
        x=agg.values.round(1), y=agg.index, orientation="h",
        marker_color=[APT_CORES.get(a,"#888") for a in agg.index],
        text=[f"{v:.1f}h" for v in agg.values],
        textposition="outside", textfont=dict(color="#e0e0e0"),
        hovertemplate="<b>%{y}</b><br>%{x:.1f}h<extra></extra>"))
    fig.update_layout(**FIG_L, height=max(260,len(agg)*38+60),
        margin=dict(l=ml+20,r=60,t=10,b=10),
        xaxis=dict(showgrid=True,gridcolor="rgba(255,255,255,0.1)",title="horas",color="#e0e0e0"),
        yaxis=dict(showgrid=False,color="#e0e0e0"))
    return fig

def fig_disponibilidade(disp_df):
    fig = go.Figure()
    fig.add_trace(go.Bar(name="Disponível",x=disp_df["Frota"],
        y=disp_df["Disponibilidade"].round(1),marker_color=C_DISP,
        text=disp_df["Disponibilidade"].round(1).astype(str)+"%",
        textposition="inside",textfont=dict(color="#fff"),
        hovertemplate="<b>%{x}</b><br>Disp: %{y:.1f}%<extra></extra>"))
    fig.add_trace(go.Bar(name="Quebra",x=disp_df["Frota"],
        y=disp_df["Indisponibilidade"].round(1),marker_color=C_MANUTENCAO,
        text=disp_df["Indisponibilidade"].round(1).astype(str)+"%",
        textposition="inside",textfont=dict(color="#fff"),
        hovertemplate="<b>%{x}</b><br>Quebra: %{y:.1f}%<extra></extra>"))
    fig.update_layout(**FIG_L,barmode="stack",height=260,
        xaxis=dict(showgrid=False,color="#e0e0e0"),margin=dict(l=4,r=4,t=4,b=4),
        yaxis=dict(showgrid=True,gridcolor="rgba(255,255,255,0.1)",title="%",
                   range=[0,105],color="#e0e0e0"),
        legend=dict(orientation="h",y=1.1,font=dict(color="#e0e0e0")))
    return fig

def fig_motor_elevador(df_perf):
    if df_perf.empty: return go.Figure()
    fig = go.Figure()
    fig.add_trace(go.Bar(name="Motor Ligado",x=df_perf["NomeFrota"],
        y=df_perf["TempoMotorLigado"].round(1),marker_color=C_MOTOR,
        text=df_perf["TempoMotorLigado"].round(1).astype(str)+"h",
        textposition="inside",textfont=dict(color="#fff"),
        hovertemplate="<b>%{x}</b><br>Motor: %{y:.1f}h<extra></extra>"))
    fig.add_trace(go.Bar(name="Elevador",x=df_perf["NomeFrota"],
        y=df_perf["TempoElevador"].round(1),marker_color=C_ELEVADOR,
        text=df_perf["TempoElevador"].round(1).astype(str)+"h",
        textposition="inside",textfont=dict(color="#fff"),
        hovertemplate="<b>%{x}</b><br>Elevador: %{y:.1f}h<extra></extra>"))
    fig.update_layout(**FIG_L,barmode="group",height=280,
        xaxis=dict(showgrid=False,color="#e0e0e0"),margin=dict(l=4,r=4,t=4,b=4),
        yaxis=dict(showgrid=True,gridcolor="rgba(255,255,255,0.1)",title="horas",color="#e0e0e0"),
        legend=dict(orientation="h",y=1.12,font=dict(color="#e0e0e0")))
    return fig

def fig_temperaturas(row):
    labels=["Motor","Óleo Hidráulico"]
    meds=[row.get("TempMedMotor",0),row.get("TempMedOleo",0)]
    maxs=[row.get("TempMaxMotor",0),row.get("TempMaxOleo",0)]
    mins=[row.get("TempMinMotor",0),row.get("TempMinOleo",0)]
    cores=["#E24B4A","#378ADD"]
    fig=go.Figure()
    for lb,md,mn,mx,co in zip(labels,meds,mins,maxs,cores):
        fig.add_trace(go.Bar(name=lb,x=[lb],y=[md],marker_color=co,
            text=f"méd {md:.0f}° · máx {mx:.0f}°",
            textposition="inside",textfont=dict(color="#fff",size=11),
            error_y=dict(type="data",symmetric=False,
                         array=[max(0,mx-md)],arrayminus=[max(0,md-mn)]),
            hovertemplate=f"<b>{lb}</b><br>méd {md}°C · máx {mx}°C<extra></extra>"))
    fig.update_layout(**FIG_L,height=260,showlegend=False,
        margin=dict(l=55,r=20,t=20,b=10),
        yaxis=dict(title="°C",showgrid=True,gridcolor="rgba(255,255,255,0.1)",color="#e0e0e0"),
        xaxis=dict(showgrid=False,color="#e0e0e0"))
    return fig

def fig_gauge(valor,titulo,maximo=100,cor=C_PRODUTIVO,sufixo="%"):
    fig=go.Figure(go.Indicator(mode="gauge+number",value=round(valor,1),
        number=dict(suffix=sufixo,font=dict(size=22,color=cor)),
        gauge=dict(axis=dict(range=[0,maximo],tickfont=dict(color="#e0e0e0")),
                   bar=dict(color=cor,thickness=0.6),bgcolor="rgba(255,255,255,0.05)",
                   borderwidth=0,steps=[dict(range=[0,maximo],color="rgba(255,255,255,0.05)")]),
        title=dict(text=titulo,font=dict(size=12,color="#aaa"))))
    fig.update_layout(**FIG_L,height=200,margin=dict(l=20,r=20,t=40,b=20))
    return fig

# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown(f"## {CONF['icone']} COA · TESTON")
    st.markdown(f"**{CONF['sub']}** · Safra 2026/27")
    st.markdown("---")

    st.markdown("### Dados SharePoint")
    fonte = st.radio("Fonte",["SharePoint (auto)","Upload manual"],label_visibility="collapsed")

    sp_data = {}
    if fonte == "SharePoint (auto)":
        with st.spinner("Conectando SharePoint..."):
            sp_data = load_from_sharepoint()
        if sp_data.get("apt_consol"):
            st.success("✓ Conectado")
        else:
            st.error("Falha ao carregar consolidado.")
    else:
        up_apt  = st.file_uploader("XLSX — Apontamentos Consolidado", type=["xlsx","csv"], key="ua")
        up_perf = st.file_uploader("XLSX — Performance Consolidado",  type=["xlsx","csv"], key="up")
        up_veic = st.file_uploader("XLSX — Veículos Agritel",         type=["xlsx","csv"], key="uv")
        up_cc   = st.file_uploader("XLSX — Frotas Centro de Custo",   type=["xlsx","csv"], key="uc")
        if up_apt:  sp_data["apt_consol"]  = up_apt.read()
        if up_perf: sp_data["perf_consol"] = up_perf.read()
        if up_veic: sp_data["veiculos"]    = up_veic.read()
        if up_cc:   sp_data["frotas_cc"]   = up_cc.read()

    if not sp_data.get("apt_consol"):
        st.warning("Aguardando APONTAMENTOS_CONSOLIDADO.xlsx")
        st.stop()

    if not sp_data.get("frotas_cc"):
        st.warning("Aguardando Frotas_CentroDeCusto.xlsx")
        st.stop()

    # ── Mapas base ───────────────────────────────────────────────────────────
    cc_map       = load_frotas_cc(sp_data["frotas_cc"])      # {frota_num: (frente, tipo)}
    veic_df      = load_veiculos(sp_data["veiculos"]) if sp_data.get("veiculos") else None
    tipo_lookup  = load_tipo_maq_lookup(sp_data["veiculos"]) if sp_data.get("veiculos") else {}
    veic_tipomaq = build_veic_tipomaq(veic_df, tipo_lookup)  # {frota_num: tipo_maq}

    st.markdown("---")

    # ── Tipo de operação ─────────────────────────────────────────────────────
    st.markdown("### Operação")
    tipo_opcao = st.radio("Tipo",["🌾  Colheita","🚜  Agro"],
                          horizontal=True,label_visibility="collapsed",key="tipo_operacao")
    APP_TIPO = "colheita" if "Colheita" in tipo_opcao else "agro"
    CONF     = TIPO_CFG[APP_TIPO]
    st.session_state["tipo_operacao_val"] = APP_TIPO

    st.markdown("---")

    # ── Período ──────────────────────────────────────────────────────────────
    st.markdown("### Período")
    hoje      = datetime.now()
    d_ini_def = hoje.replace(day=1).date()
    d_fim_def = hoje.date()

    col_a, col_b = st.columns(2)
    with col_a:
        data_ini = st.date_input("De",  value=d_ini_def, key="d_ini", format="DD/MM/YYYY")
    with col_b:
        data_fim = st.date_input("Até", value=d_fim_def, key="d_fim", format="DD/MM/YYYY")

    if data_ini > data_fim:
        st.error("Data início > data fim.")
        st.stop()

    dias_periodo = (data_fim - data_ini).days + 1
    period_h     = dias_periodo * 24.0
    st.caption(
        f"**{data_ini.strftime('%d/%m/%Y')}** – **{data_fim.strftime('%d/%m/%Y')}**\n"
        f"**{dias_periodo}** dias · **{period_h:.0f}h**/máquina"
    )

    # ── Carrega consolidados ─────────────────────────────────────────────────
    df_apt_raw  = _load_apt_raw(sp_data["apt_consol"])
    df_perf_raw = (_load_perf_raw(sp_data["perf_consol"])
                   if sp_data.get("perf_consol") else pd.DataFrame())

    st.markdown("---")
    st.markdown("### Filtros")

    # Frentes vindas do cc_map (fonte dinâmica — sem hardcode)
    frentes_disponiveis = sorted({
        frente for fn, (frente, tipo) in cc_map.items()
        if tipo == APP_TIPO
    })

    frentes_sel = st.multiselect(
        "Frente", frentes_disponiveis, default=frentes_disponiveis,
        key="frentes_sel", placeholder="Todas as frentes"
    ) or frentes_disponiveis

    # Tipos de máquina
    if veic_tipomaq:
        _tms = {veic_tipomaq.get(fn,"Outro")
                for fn,(fr,tp) in cc_map.items() if tp==APP_TIPO} - {"Outro"}
        tipos_maq_disponiveis = sorted(_tms) or (
            ["Colhedora","Trator"] if APP_TIPO=="colheita" else ["Trator","Pulverizador"])
    else:
        tipos_maq_disponiveis = (["Colhedora","Trator"]
                                  if APP_TIPO=="colheita" else ["Trator","Pulverizador"])

    tipos_maq_sel = st.multiselect(
        "Tipo de máquina", tipos_maq_disponiveis, default=tipos_maq_disponiveis,
        key="tipos_maq_sel", placeholder="Todos os tipos"
    ) or tipos_maq_disponiveis

    # ── Aplica filtros ────────────────────────────────────────────────────────
    df_apt  = filter_apt(df_apt_raw, data_ini, data_fim,
                         cc_map, veic_tipomaq, APP_TIPO, frentes_sel, tipos_maq_sel)
    df_perf = (filter_perf(df_perf_raw, data_ini, data_fim,
                           cc_map, veic_tipomaq, APP_TIPO, frentes_sel, tipos_maq_sel)
               if not df_perf_raw.empty else pd.DataFrame())
    disp_df = calc_disponibilidade(df_apt) if not df_apt.empty else pd.DataFrame()

    frotas_disp = ["Todas"] + sorted(df_apt["Frota"].unique().tolist())
    frota_sel   = st.selectbox("Frota", frotas_disp, key="frota_sel")

    if not df_apt_raw.empty:
        bd_min = df_apt_raw["Data"].min().strftime("%d/%m/%Y")
        bd_max = df_apt_raw["Data"].max().strftime("%d/%m/%Y")
        n_maq  = df_apt_raw["Frota"].nunique()
        st.markdown(
            f'<div class="api-badge">📦 BD: {bd_min} → {bd_max} '
            f'· {n_maq} máquinas</div>',
            unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("### Navegação")
    pagina = st.radio("Página",
        ["Visão geral","Apontamentos","Telemetria","Por frota"],
        label_visibility="collapsed")

    st.markdown("---")
    if st.button("🚪 Sair", use_container_width=True):
        st.session_state["autenticado"] = False
        st.rerun()
    st.caption(f"MS Colheitas · TESTON · {CONF['sub']} · 2026")

# ══════════════════════════════════════════════════════════════════════════════
# DADOS FILTRADOS
# ══════════════════════════════════════════════════════════════════════════════
if df_apt.empty:
    st.warning("Nenhum dado de apontamento para o período e filtros selecionados.")
    st.stop()

df_f      = df_apt  if frota_sel=="Todas" else df_apt[df_apt["Frota"]==frota_sel]
df_perf_f = (df_perf if frota_sel=="Todas"
             else df_perf[df_perf["NomeFrota"]==frota_sel]) \
            if not df_perf.empty else pd.DataFrame()

num_frotas = max(df_f["Frota"].nunique(), 1)
total_h    = df_f["Horas"].sum()
col_h      = df_f[df_f["Apontamento"].isin(["Colhendo","Trabalhando"])]["Horas"].sum()
man_h      = df_f[df_f["Apontamento"]=="Manobra"]["Horas"].sum()
sa_h       = df_f[df_f["Categoria"]=="Sem Apontamento"]["Horas"].sum()
qbr_h      = df_f[df_f["Apontamento"]=="Quebra"]["Horas"].sum()
chuva_h    = df_f[df_f["Apontamento"]=="Chuva / Umidade"]["Horas"].sum()
agt_h      = df_f[df_f["Apontamento"].isin(["Aguardando Transbordo","Aguardando Colhedora"])]["Horas"].sum()
improd     = chuva_h + qbr_h + agt_h

ef_col   = safe(col_h,  total_h)
ef_man   = safe(man_h,  col_h + man_h)
ef_imp   = safe(improd, total_h)
ef_sa    = safe(sa_h,   total_h)
disp_m   = safe(total_h - qbr_h, total_h)
ef_chuva = safe(chuva_h, total_h)
ef_qbr   = safe(qbr_h,   total_h)
ef_agt   = safe(agt_h,   total_h)

motor_h    = df_perf_f["TempoMotorLigado"].sum() if not df_perf_f.empty else 0.0
elevador_h = df_perf_f["TempoElevador"].sum()    if not df_perf_f.empty else 0.0
ef_trab    = safe(elevador_h, motor_h)

# ══════════════════════════════════════════════════════════════════════════════
# BANNER
# ══════════════════════════════════════════════════════════════════════════════
def banner():
    fr_label = "Todas as frotas" if frota_sel=="Todas" else frota_sel
    st.markdown(
        f'<div class="periodo-badge">'
        f'📅 {data_ini.strftime("%d/%m/%Y")} – {data_fim.strftime("%d/%m/%Y")} '
        f'· <b>{dias_periodo}d / {period_h:.0f}h</b> por máquina '
        f'&nbsp;|&nbsp; {CONF["icone"]} {fr_label} · {num_frotas} máquina(s)'
        f'</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# PÁGINAS
# ══════════════════════════════════════════════════════════════════════════════
if pagina == "Visão geral":
    st.markdown(f"## {CONF['icone']} Visão geral — {CONF['sub']}")
    banner()

    c1,c2,c3,c4,c5 = st.columns(5)
    c1.metric("⏱️ Período / máquina", f"{period_h:.0f}h",
              f"{dias_periodo} dias · {total_h:.0f}h apontado")
    c2.metric("✅ Disponibilidade",   f"{disp_m:.1f}%",  f"{qbr_h:.1f}h quebra")
    c3.metric("⚡ Efic. colheita",    f"{ef_col:.1f}%",  f"{col_h:.1f}h colhendo")
    c4.metric("↔️ % Manobra",         f"{ef_man:.1f}%",  f"{man_h:.1f}h")
    c5.metric("❓ Sem apontamento",    f"{ef_sa:.1f}%",   f"{sa_h:.1f}h")

    if not df_perf_f.empty:
        st.markdown("<p style='margin:14px 0 6px;font-size:11px;color:#888;"
                    "text-transform:uppercase;letter-spacing:.08em'>⚙️ Motor & Elevador</p>",
                    unsafe_allow_html=True)
        cm1,cm2,cm3,cm4 = st.columns(4)
        cm1.metric("🔵 Motor Ligado",   f"{motor_h:.1f}h",    f"{motor_h/num_frotas:.1f}h/máq")
        cm2.metric("🟣 Hora Elevador",  f"{elevador_h:.1f}h", f"{elevador_h/num_frotas:.1f}h/máq")
        cm3.metric("📊 Efic. Trabalho", f"{ef_trab:.1f}%",    "Elevador ÷ Motor Ligado")
        cons_total = df_perf_f["ConsumoTotal"].sum()
        cm4.metric("🛢️ Consumo total",  f"{cons_total:,.0f} L".replace(",","."))

    st.markdown("<p style='margin:14px 0 6px;font-size:11px;color:#888;"
                "text-transform:uppercase;letter-spacing:.08em'>⚠️ Breakdown Improdutivo</p>",
                unsafe_allow_html=True)
    ci0,ci1,ci2,ci3 = st.columns(4)
    ci0.metric("⚠️ Improdutivo total",     f"{ef_imp:.1f}%",  f"{improd:.1f}h")
    ci1.metric("🌧️ Chuva / Umidade",      f"{ef_chuva:.1f}%",f"{chuva_h:.1f}h")
    ci2.metric("🔧 Quebra",                f"{ef_qbr:.1f}%",  f"{qbr_h:.1f}h")
    ci3.metric("⏳ Aguardando",            f"{ef_agt:.1f}%",  f"{agt_h:.1f}h")

    if not df_perf_f.empty:
        cons_med = df_perf_f["TaxaMedTrab"].mean()
        st.divider()
        c6,c7,c8,c9 = st.columns(4)
        c6.metric("🛢️ Consumo médio trab.", f"{cons_med:.1f} L/h")
        c7.metric("🌡️ Temp. média motor",   f"{df_perf_f['TempMedMotor'].mean():.1f}°C")
        c8.metric("🌡️ Temp. máx. motor",   f"{df_perf_f['TempMaxMotor'].max():.1f}°C")
        c9.metric("🌡️ Temp. máx. óleo",    f"{df_perf_f['TempMaxOleo'].max():.1f}°C")

    st.divider()
    if not df_perf_f.empty:
        st.markdown("**Hora Motor Ligado vs Hora Elevador**")
        st.plotly_chart(fig_motor_elevador(df_perf_f), use_container_width=True,
                        config={"displayModeBar":False})
        st.divider()

    if not disp_df.empty:
        st.markdown("**Disponibilidade mecânica por frota**")
        st.plotly_chart(fig_disponibilidade(disp_df), use_container_width=True,
                        config={"displayModeBar":False})
        st.divider()

    cl,cr = st.columns(2)
    with cl:
        st.markdown("**Distribuição por categoria**")
        st.plotly_chart(fig_donut(df_f), use_container_width=True, config={"displayModeBar":False})
    with cr:
        st.markdown("**Comparativo entre frotas**")
        tots = df_apt.groupby("Frota")["Horas"].sum()
        kc   = st.columns(min(len(tots)+1,6))
        kc[0].metric("⏱️ Total", f"{df_apt['Horas'].sum():.1f}h")
        for j,(fr,hr) in enumerate(tots.items()):
            if j+1 < len(kc): kc[j+1].metric(f"⏱️ {fr}", f"{hr:.1f}h")
        st.plotly_chart(fig_comparativo(df_apt), use_container_width=True,
                        config={"displayModeBar":False})

    st.markdown("**Perdas e paradas — horas acumuladas**")
    st.plotly_chart(fig_perdas(df_f), use_container_width=True, config={"displayModeBar":False})


elif pagina == "Apontamentos":
    st.markdown(f"## {CONF['icone']} Apontamentos")
    banner()
    c1,c2,c3 = st.columns(3)
    c1.metric("✅ Disponibilidade", f"{disp_m:.1f}%",
              f"{qbr_h:.1f}h quebra · {total_h:.1f}h total")
    c2.metric("⚡ Efic. colheita",  f"{ef_col:.1f}%", f"{col_h:.1f}h")
    c3.metric("❓ Sem apontamento", f"{ef_sa:.1f}%",  f"{sa_h:.1f}h")
    st.divider()
    st.markdown("**Horas por apontamento**")
    st.plotly_chart(fig_barras_h(df_f), use_container_width=True, config={"displayModeBar":False})
    st.divider()
    cl,cr = st.columns(2)
    with cl:
        st.markdown("**Distribuição por categoria**")
        st.plotly_chart(fig_donut(df_f), use_container_width=True, config={"displayModeBar":False})
    with cr:
        st.markdown("**Resumo por categoria**")
        cats = df_f.groupby("Categoria")["Horas"].sum().reset_index()
        cats = cats[cats["Horas"]>0].sort_values("Horas",ascending=False)
        cats["pct"]   = (cats["Horas"]/total_h*100).round(1).astype(str)+"%"
        cats["Horas"] = cats["Horas"].round(1)
        mh = cats["Horas"].max()
        for _,row in cats.iterrows():
            cor = CAT_CORES.get(row["Categoria"],"#888")
            pb  = row["Horas"]/mh*100
            st.markdown(
                f'<div style="margin-bottom:10px">'
                f'<div style="display:flex;justify-content:space-between;margin-bottom:4px">'
                f'<span style="font-size:13px;color:#ccc">{row["Categoria"]}</span>'
                f'<span style="font-size:14px;font-weight:700;color:#e0e0e0">{row["Horas"]}h '
                f'<span style="font-weight:400;color:#888;font-size:12px">({row["pct"]})</span></span></div>'
                f'<div style="background:rgba(255,255,255,.08);border-radius:6px;height:18px;overflow:hidden">'
                f'<div style="width:{pb:.1f}%;height:100%;background:{cor};border-radius:6px"></div>'
                f'</div></div>', unsafe_allow_html=True)


elif pagina == "Telemetria":
    st.markdown(f"## {CONF['icone']} Telemetria — Performance")
    banner()
    if df_perf_f.empty:
        st.warning("Nenhum dado de performance para este período/frota.")
        st.stop()

    cols = st.columns(min(len(df_perf_f),4))
    for i,(_,row) in enumerate(df_perf_f.iterrows()):
        with cols[i % len(cols)]:
            st.markdown(f"**{row['NomeFrota']}**")
            if not disp_df.empty:
                d_row = disp_df[disp_df["Frota"]==row["NomeFrota"]]
                if not d_row.empty:
                    dv = d_row.iloc[0]["Disponibilidade"]
                    qv = d_row.iloc[0]["Quebra"]
                    st.markdown(
                        f'<div style="padding:8px 12px;border-radius:8px;'
                        f'background:rgba(39,174,96,.15);border-left:3px solid {C_DISP};margin-bottom:8px">'
                        f'<span style="font-size:11px;color:#aaa">Disponibilidade mecânica</span><br>'
                        f'<span style="font-size:20px;font-weight:700;color:{C_DISP}">{dv:.1f}%</span>'
                        f'<span style="font-size:10px;color:#888;margin-left:8px">({qv:.1f}h quebra)</span>'
                        f'</div>', unsafe_allow_html=True)
            ef_el = safe(row.get("TempoElevador",0), row.get("TempoMotorLigado",1))
            st.markdown(
                f'<div style="padding:8px 12px;border-radius:8px;'
                f'background:rgba(155,89,182,.15);border-left:3px solid {C_ELEVADOR};margin-bottom:8px">'
                f'<span style="font-size:11px;color:#aaa">Efic. Trabalho (Elev./Motor)</span><br>'
                f'<span style="font-size:20px;font-weight:700;color:{C_ELEVADOR}">{ef_el:.1f}%</span>'
                f'<span style="font-size:10px;color:#888;margin-left:8px">'
                f'({row.get("TempoElevador",0):.1f}h / {row.get("TempoMotorLigado",0):.1f}h)</span>'
                f'</div>', unsafe_allow_html=True)
            for lb,val,cor in [
                ("Motor ligado",       f"{row.get('TempoMotorLigado',0):.1f}h",  C_MOTOR),
                ("Elevador ligado",    f"{row.get('TempoElevador',0):.1f}h",     C_ELEVADOR),
                ("Em trabalho",        f"{row.get('TempoTrabalho',0):.1f}h",     "#e0e0e0"),
                ("Manobra",            f"{row.get('TempoManobra',0):.1f}h",      "#e0e0e0"),
                ("Ocioso",             f"{row.get('TempoOcioso',0):.1f}h",       "#e0e0e0"),
                ("Deslocando",         f"{row.get('TempoDeslocando',0):.1f}h",   "#e0e0e0"),
                ("Consumo total",
                 f"{int(row.get('ConsumoTotal',0)):,} L".replace(",","."),       "#e0e0e0"),
                ("Taxa cons. (trab.)", f"{row.get('TaxaMedTrab',0):.1f} L/h",   "#e0e0e0"),
                ("Temp. méd. motor",   f"{row.get('TempMedMotor',0):.1f}°C",    "#E24B4A"),
                ("Temp. máx. motor",   f"{row.get('TempMaxMotor',0):.1f}°C",    "#E24B4A"),
                ("Temp. méd. óleo",    f"{row.get('TempMedOleo',0):.1f}°C",     "#378ADD"),
                ("Temp. máx. óleo",    f"{row.get('TempMaxOleo',0):.1f}°C",     "#378ADD"),
                ("Reverso Industrial",  f"{int(row.get('ReversoIndustrial',0))} ocorr.", C_MANUTENCAO),
            ]:
                st.markdown(
                    f'<div style="display:flex;justify-content:space-between;padding:6px 10px;'
                    f'background:rgba(255,255,255,.05);border-radius:6px;margin-bottom:4px">'
                    f'<span style="font-size:11px;color:#999">{lb}</span>'
                    f'<span style="font-size:12px;font-weight:500;color:{cor}">{val}</span>'
                    f'</div>', unsafe_allow_html=True)

    st.divider()
    st.markdown("**Temperaturas — méd / máx**")
    tc = st.columns(min(len(df_perf_f),4))
    for i,(_,row) in enumerate(df_perf_f.iterrows()):
        with tc[i % len(tc)]:
            st.markdown(f"*{row['NomeFrota']}*")
            st.plotly_chart(fig_temperaturas(row), use_container_width=True,
                            config={"displayModeBar":False})

    st.divider()
    st.markdown("**Gauges operacionais**")
    for _, row in df_perf_f.iterrows():
        ef_m  = safe(row.get("TempoTrabalho",0), row.get("TempoMotorLigado",1))
        ef_el = safe(row.get("TempoElevador",0), row.get("TempoMotorLigado",1))
        oc    = safe(row.get("TempoOcioso",0),   row.get("TempoMotorLigado",1))
        g1, g2, g3 = st.columns(3)
        with g1:
            st.plotly_chart(fig_gauge(ef_m,  f"{row['NomeFrota']}\nEfic. Motor",    cor=C_PRODUTIVO),
                            use_container_width=True, config={"displayModeBar":False})
        with g2:
            st.plotly_chart(fig_gauge(ef_el, f"{row['NomeFrota']}\nEfic. Elevador", cor=C_ELEVADOR),
                            use_container_width=True, config={"displayModeBar":False})
        with g3:
            st.plotly_chart(fig_gauge(oc,    f"{row['NomeFrota']}\nOcioso",         cor=C_MANUTENCAO),
                            use_container_width=True, config={"displayModeBar":False})


elif pagina == "Por frota":
    frotas_lista = sorted(df_apt["Frota"].unique().tolist())
    fr4 = st.selectbox("Selecione a frota", frotas_lista, key="fr4")

    sub_apt  = df_apt[df_apt["Frota"]==fr4]
    sub_perf = df_perf[df_perf["NomeFrota"]==fr4] if not df_perf.empty else pd.DataFrame()
    sub_disp = disp_df[disp_df["Frota"]==fr4] if not disp_df.empty else pd.DataFrame()

    st.markdown(f"## {CONF['icone']} {fr4} — Análise individual")
    banner()

    tot_f = sub_apt["Horas"].sum()
    col_f = sub_apt[sub_apt["Apontamento"].isin(["Colhendo","Trabalhando"])]["Horas"].sum()
    man_f = sub_apt[sub_apt["Apontamento"]=="Manobra"]["Horas"].sum()
    qbr_f = sub_apt[sub_apt["Apontamento"]=="Quebra"]["Horas"].sum()
    d_f   = safe(tot_f - qbr_f, tot_f)

    k1,k2,k3,k4 = st.columns(4)
    k1.metric("✅ Disponibilidade",  f"{d_f:.1f}%",
              f"{qbr_f:.1f}h quebra · {tot_f:.1f}h total")
    k2.metric("⚡ Efic. colheita",   f"{safe(col_f,tot_f):.1f}%", f"{col_f:.1f}h")
    k3.metric("↔️ % Manobra",        f"{safe(man_f,col_f+man_f):.1f}%", f"{man_f:.1f}h")
    k4.metric("⏱️ Período ref.",     f"{period_h:.0f}h", f"{tot_f:.1f}h apontado")

    if not sub_perf.empty:
        row = sub_perf.iloc[0]
        ef_el = safe(row.get("TempoElevador",0), row.get("TempoMotorLigado",1))
        st.divider()
        k5,k6,k7,k8 = st.columns(4)
        k5.metric("🔵 Motor Ligado",  f"{row.get('TempoMotorLigado',0):.1f}h")
        k6.metric("🟣 Hora Elevador", f"{row.get('TempoElevador',0):.1f}h",
                  f"Efic. {ef_el:.1f}%")
        k7.metric("🛢️ Consumo",
                  f"{int(row.get('ConsumoTotal',0)):,} L".replace(",","."),
                  f"{row.get('TaxaMedTrab',0):.1f} L/h")
        oc = safe(row.get("TempoOcioso",0), row.get("TempoMotorLigado",1))
        k8.metric("⏸️ Ocioso", f"{oc:.1f}%", f"{row.get('TempoOcioso',0):.1f}h")

    st.divider()
    cl,cr = st.columns(2)
    with cl:
        st.markdown("**Uso do tempo**")
        st.plotly_chart(fig_donut(sub_apt), use_container_width=True,
                        config={"displayModeBar":False})
    with cr:
        st.markdown("**Apontamentos detalhados**")
        st.plotly_chart(fig_barras_h(sub_apt), use_container_width=True,
                        config={"displayModeBar":False})

    if not sub_perf.empty:
        st.divider()
        st.markdown("**Temperaturas**")
        st.plotly_chart(fig_temperaturas(sub_perf.iloc[0]), use_container_width=True,
                        config={"displayModeBar":False})