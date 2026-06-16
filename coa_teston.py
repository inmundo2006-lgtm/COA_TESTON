"""
COA · TESTON — Dashboard Safra 2026/27  |  v4
Streamlit + Plotly | Login + SharePoint + Agritel API direto

NOVIDADES v4:
  - Período selecionado no sidebar (date picker) — sem secrets PERIODO_INICIO/FIM
  - Performance buscada direto da Agritel API (/operational-report)
  - Chunking automático (API limita 15 dias por chamada)
  - Cache 10 min alinhado ao rate limit da API
  - apontamentos.xlsx ainda vem do SharePoint (API não tem detalhe de Quebra/Chuva etc.)
  - performance.xlsx removido do SharePoint

SECRETS necessários:
  SP_TENANT_ID, SP_CLIENT_ID, SP_CLIENT_SECRET
  SP_SITE_URL      = "https://metalcana.sharepoint.com/sites/MSCOLHEITA"
  SP_APT_PATH      = "/VITOR/AGRITEL/apontamentos.xlsx"
  SP_VEICULOS_PATH = "/VITOR/AGRITEL/Veículos AGRITEL.xlsx"
  SP_FROTAS_CC_PATH= "/VITOR/AGRITEL/Frotas_CentroDeCusto.xlsx"
  SP_PERIODO_PATH  = "/VITOR/AGRITEL/periodo_agritel.jpg"   (opcional)
  AGRITEL_API_KEY  = "69d65ab7023281608700eedd"
  AGRITEL_API_URL  = "https://api.agritel.com.br"           (opcional)
  APP_TIPO         = "colheita"   # "colheita" | "agro"
  DASHBOARD_SENHA_HASH = "<sha256>"
"""

import re, io, os, hashlib
from datetime import datetime, date, timedelta, timezone
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
SP_APT_PATH       = _s("SP_APT_PATH",       "/VITOR/AGRITEL/apontamentos.xlsx")
SP_VEICULOS_PATH  = _s("SP_VEICULOS_PATH",  "/VITOR/AGRITEL/Veículos_Agritel.xlsx")
SP_FROTAS_CC_PATH = _s("SP_FROTAS_CC_PATH", "/VITOR/AGRITEL/Frotas_CentroDeCusto.xlsx")
SP_PERIODO_PATH   = _s("SP_PERIODO_PATH",   "/VITOR/AGRITEL/periodo_agritel.jpg")

AGRITEL_API_KEY   = _s("AGRITEL_API_KEY",   "69d65ab7023281608700eedd")
AGRITEL_API_URL   = _s("AGRITEL_API_URL",   "https://api.agritel.com.br")

SENHA_HASH = _s("DASHBOARD_SENHA_HASH", hashlib.sha256("teston2026".encode()).hexdigest())

# Fuso horário MS (UTC-4) — dia agrícola 07:00→06:59 local
TZ_OFFSET = timedelta(hours=4)   # soma para converter local→UTC

TIPO_CFG = {
    "colheita": {"icone": "🌾", "titulo": "COA · TESTON — Colheita", "sub": "Colhedoras & Transbordos"},
    "agro":     {"icone": "🚜", "titulo": "COA · TESTON — Agro",     "sub": "Máquinas Agrícolas"},
}
# CONF padrão inicial — sobrescrito no sidebar após o radio selector
APP_TIPO = st.session_state.get("tipo_operacao_val", "colheita")
CONF     = TIPO_CFG[APP_TIPO]
KEYWORDS_COLHEITA = ["colhedora", "transbordo"]

DEVICE_MAP_FALLBACK = {
    "671a433018212ed9057ce3ee": "Frota 1262",
    "671a433018212ed9057ce3f1": "Frota 1263",
}

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
    "Colhendo": "Produtivo", "Manobra": "Manobra", "Deslocando": "Manobra",
    "Quebra": "Manutenção", "Manutenção Programada": "Manutenção",
    "Troca Faquinha": "Manutenção", "PARADA OPERACIONAL": "Manutenção",
    "Chuva / Umidade": "Espera/Imprevisto", "Usina Parada": "Espera/Imprevisto",
    "Falta de Area Liberada": "Espera/Imprevisto", "Falta de Caminhão": "Espera/Imprevisto",
    "Aguardando Transbordo": "Espera/Imprevisto",
    "Abastecimento": "Parada Operacional", "Refeição": "Parada Operacional",
    "Checklist": "Parada Operacional", "Final Jornada de Trabalho": "Parada Operacional",
    "Início Jornada de Trabalho": "Parada Operacional",
}
CAT_CORES = {
    "Produtivo": C_PRODUTIVO, "Manutenção": C_MANUTENCAO,
    "Espera/Imprevisto": C_ESPERA, "Manobra": C_MANOBRA,
    "Parada Operacional": C_PARADA, "Sem Apontamento": C_SEM_APT,
}
APT_CORES = {
    "Colhendo": C_PRODUTIVO, "Aguardando Transbordo": C_ESPERA,
    "Chuva / Umidade": "#1a5fa8", "Quebra": C_MANUTENCAO,
    "Manobra": C_MANOBRA, "Sem Apontamento": C_SEM_APT,
    "Manutenção Programada": "#2ea05f", "Abastecimento": "#5DCAA5",
    "Troca Faquinha": "#D85A30", "Checklist": C_PARADA,
    "Refeição": C_PARADA, "Deslocando": "#9b9b9b",
    "Final Jornada de Trabalho": "#9b9b9b", "Início Jornada de Trabalho": "#9b9b9b",
    "Falta de Area Liberada": C_MANUTENCAO, "Falta de Caminhão": C_MANUTENCAO,
    "PARADA OPERACIONAL": C_MANUTENCAO, "Usina Parada": "#9b9b9b",
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
[data-testid="stSidebar"] *         {color:#d8dbd4!important}
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
# PARSERS
# ══════════════════════════════════════════════════════════════════════════════
def parse_horas(v):
    if isinstance(v, (int, float)):
        return float(v) if not np.isnan(v) else 0.0
    m = re.search(r"([\d.,]+)\s*h", str(v))
    return float(m.group(1).replace(".","").replace(",",".")) if m else 0.0

def parse_num(v):
    if isinstance(v, (int, float)):
        return float(v) if not np.isnan(v) else 0.0
    c = re.sub(r"[^\d,.]", "", str(v)).replace(".","").replace(",",".")
    try: return float(c)
    except: return 0.0

def _norm(s):
    s = str(s).lower()
    for a,b in [("á","a"),("ã","a"),("â","a"),("é","e"),("ê","e"),
                ("í","i"),("ó","o"),("õ","o"),("ô","o"),("ú","u"),("ç","c")]:
        s = s.replace(a,b)
    return s

def _find_col(df, kws):
    nm = {_norm(c): c for c in df.columns}
    for kw in kws:
        for n,c in nm.items():
            if kw in n: return c
    return None

def _read_df(content):
    if content[:4] == b"PK\x03\x04":
        return pd.read_excel(io.BytesIO(content), sheet_name=0)
    return pd.read_csv(io.BytesIO(content))

# ══════════════════════════════════════════════════════════════════════════════
# SHAREPOINT — só apontamentos, veículos e CC (sem performance.xlsx)
# ══════════════════════════════════════════════════════════════════════════════
@st.cache_data(ttl=600)
def load_from_sharepoint():
    result = {k: None for k in ("apt","veiculos","frotas_cc","periodo_img")}
    try:
        import requests as req
        tok = req.post(
            f"https://login.microsoftonline.com/{SP_TENANT_ID}/oauth2/v2.0/token",
            data={"grant_type":"client_credentials","client_id":SP_CLIENT_ID,
                  "client_secret":SP_CLIENT_SECRET,"scope":"https://graph.microsoft.com/.default"},
            timeout=15).json()
        h = {"Authorization": f"Bearer {tok['access_token']}"}
        host  = SP_SITE_URL.replace("https://","").split("/")[0]
        spath = "/".join(SP_SITE_URL.replace("https://","").split("/")[1:])
        sid   = req.get(f"https://graph.microsoft.com/v1.0/sites/{host}:/{spath}",
                        headers=h, timeout=15).json()["id"]
        did   = req.get(f"https://graph.microsoft.com/v1.0/sites/{sid}/drive",
                        headers=h, timeout=15).json()["id"]

        def _dl(path):
            enc = quote(path, safe="/")
            try:
                r = req.get(f"https://graph.microsoft.com/v1.0/drives/{did}/root:{enc}:/content",
                            headers=h, timeout=30)
                r.raise_for_status()
                return r.content
            except Exception as e:
                st.warning(f"⚠️ SharePoint: `{path}` — {e}")
                return None

        result["apt"]        = _dl(SP_APT_PATH)
        result["veiculos"]   = _dl(SP_VEICULOS_PATH)
        result["frotas_cc"]  = _dl(SP_FROTAS_CC_PATH)
        result["periodo_img"]= _dl(SP_PERIODO_PATH)
    except Exception as e:
        st.error(f"❌ SharePoint: {e}")
    return result

# ══════════════════════════════════════════════════════════════════════════════
# AGRITEL API  — operational-report (substituiu performance.xlsx)
# ══════════════════════════════════════════════════════════════════════════════
@st.cache_data(ttl=600)
def _api_chunk(start_iso: str, end_iso: str) -> list | None:
    """Uma chamada à API (≤15 dias). Cacheada 10 min (rate limit da API)."""
    import requests as req
    try:
        r = req.get(
            f"{AGRITEL_API_URL}/telemetry-integration/operational-report",
            headers={"X-Api-Key": AGRITEL_API_KEY},
            params={"startTime": start_iso, "endTime": end_iso},
            timeout=30,
        )
        if r.status_code == 429:
            retry = r.headers.get("Retry-After", "600")
            st.warning(f"⏳ Rate limit API Agritel — aguarde {retry}s e atualize a página.")
            return None
        r.raise_for_status()
        return r.json().get("items", [])
    except Exception as e:
        st.warning(f"⚠️ Agritel API: {e}")
        return None


def fetch_agritel(data_ini: date, data_fim: date) -> tuple[list, bool]:
    """
    Busca o período completo, quebrando em chunks de até 15 dias.
    UTC = local MS (UTC-4) + 4h  →  dia agr. 07:00 local = 11:00 UTC
    Retorna (items_merged, sucesso_total).
    """
    # Converter para UTC
    start_utc = datetime(data_ini.year, data_ini.month, data_ini.day, 7, 0, 0) + TZ_OFFSET
    end_utc   = datetime(data_fim.year, data_fim.month, data_fim.day, 6, 59, 59) + TZ_OFFSET + timedelta(days=1)

    # Quebrar em chunks ≤15 dias
    chunks, cur = [], start_utc
    while cur < end_utc:
        nxt = min(cur + timedelta(days=15), end_utc)
        chunks.append((cur, nxt))
        cur = nxt

    merged: dict[str, dict] = {}
    ok = True

    for cs, ce in chunks:
        items = _api_chunk(
            cs.strftime("%Y-%m-%dT%H:%M:%SZ"),
            ce.strftime("%Y-%m-%dT%H:%M:%SZ"),
        )
        if items is None:
            ok = False
            continue
        for item in items:
            did = item["deviceId"]
            if did not in merged:
                merged[did] = {**item}
            else:
                b = merged[did]
                # Somar campos acumulativos
                for f in ("engineOnSeconds","idlingSeconds","maneuveringSeconds",
                          "workingSeconds","dislocatingSeconds","fuelConsumptionLiters"):
                    b[f] = b.get(f,0) + item.get(f,0)
                # Máximos de temperatura
                for f in ("engineCoolantTemperatureMaxCelsius","hydraulicOilTemperatureMaxCelsius",
                          "engineFuelRateMaxLitersPerHour"):
                    b[f] = max(b.get(f,0), item.get(f,0))
                # Média de taxa de consumo (média simples entre chunks)
                for f in ("engineCoolantTemperatureAvgCelsius","hydraulicOilTemperatureAvgCelsius",
                          "engineFuelRateAvgLitersPerHour"):
                    b[f] = (b.get(f,0) + item.get(f,0)) / 2
                # Horímetro: manter start do primeiro chunk, end do último
                b["elevatorHourmeterEnd"] = item.get("elevatorHourmeterEnd", b.get("elevatorHourmeterEnd",0))
                b["engineHourmeterEnd"]   = item.get("engineHourmeterEnd",   b.get("engineHourmeterEnd",0))

    return list(merged.values()), ok


def build_perf_df(items: list, device_map: dict, tipo_map: dict, filtrar_tipo: str) -> pd.DataFrame:
    """Converte resposta da API no mesmo formato de DataFrame do código anterior."""
    rows = []
    for item in items:
        did  = item.get("deviceId","")
        tipo = tipo_map.get(did, "agro")
        if tipo != filtrar_tipo:
            continue
        nome    = device_map.get(did, item.get("deviceName", did))
        m = re.search(r"\d+", nome)
        frota_n = m.group() if m else ""
        # Horímetro elevador = diferença entre fim e início do período
        elev_h  = (item.get("elevatorHourmeterEnd",0) - item.get("elevatorHourmeterStart",0))
        rows.append({
            "NomeFrota":        nome,
            "NomeCompleto":     item.get("deviceName", nome),
            "Frota":            frota_n,
            "TempoMotorLigado": item.get("engineOnSeconds",0)     / 3600,
            "TempoTrabalho":    item.get("workingSeconds",0)       / 3600,
            "TempoManobra":     item.get("maneuveringSeconds",0)   / 3600,
            "TempoOcioso":      item.get("idlingSeconds",0)        / 3600,
            "TempoDeslocando":  item.get("dislocatingSeconds",0)   / 3600,
            "TempoElevador":    max(0.0, elev_h),
            "ConsumoTotal":     item.get("fuelConsumptionLiters",0),
            "ConsumoTrabalho":  item.get("fuelConsumptionLiters",0),
            "TaxaMedTrab":      item.get("engineFuelRateAvgLitersPerHour",0),
            "TaxaMedConsumo":   item.get("engineFuelRateAvgLitersPerHour",0),
            "TempMaxMotor":     item.get("engineCoolantTemperatureMaxCelsius",0),
            "TempMedMotor":     item.get("engineCoolantTemperatureAvgCelsius",0),
            "TempMinMotor":     0,
            "TempMaxOleo":      item.get("hydraulicOilTemperatureMaxCelsius",0),
            "TempMedOleo":      item.get("hydraulicOilTemperatureAvgCelsius",0),
            "TempMinOleo":      0,
            "CargaMed":         0, "CargaMax":     0,
            "VelocidadeMed":    0, "VelocidadeMax":0,
            "RpmMotor":         0, "RpmExtMed":    0,
            "PressaoMed":       0, "ReversoIndustrial": 0,
            "Odometro":         0,
        })
    return pd.DataFrame(rows) if rows else pd.DataFrame(columns=[
        "NomeFrota","NomeCompleto","Frota","TempoMotorLigado","TempoTrabalho",
        "TempoManobra","TempoOcioso","TempoDeslocando","TempoElevador",
        "ConsumoTotal","ConsumoTrabalho","TaxaMedTrab","TaxaMedConsumo",
        "TempMaxMotor","TempMedMotor","TempMinMotor","TempMaxOleo","TempMedOleo","TempMinOleo",
        "CargaMed","CargaMax","VelocidadeMed","VelocidadeMax",
        "RpmMotor","RpmExtMed","PressaoMed","ReversoIndustrial","Odometro",
    ])

# ══════════════════════════════════════════════════════════════════════════════
# CARREGAMENTO — veículos, CC e apontamentos
# ══════════════════════════════════════════════════════════════════════════════
@st.cache_data
def load_veiculos(content: bytes) -> pd.DataFrame:
    df = _read_df(content)
    id_c   = _find_col(df,["id do dispositivo","deviceid","device id","id"])
    nome_c = _find_col(df,["nome da maquina","nome","name","maquina"])
    grp_c  = _find_col(df,["grupo de maquinas","grupo","group","tipo","descri"])
    if not id_c:   id_c   = df.columns[0]
    if not nome_c: nome_c = df.columns[min(1,len(df.columns)-1)]
    r = pd.DataFrame({
        "DeviceId": df[id_c].astype(str).str.strip(),
        "Nome":     df[nome_c].astype(str).str.strip(),
        "Grupo":    df[grp_c].astype(str).str.strip() if grp_c else "",
    })
    def _tipo(row):
        txt = _norm(row["Nome"]) + " " + _norm(row["Grupo"])
        return "colheita" if any(k in txt for k in KEYWORDS_COLHEITA) else "agro"
    r["Tipo"] = r.apply(_tipo, axis=1)
    return r

@st.cache_data
def load_frotas_cc(content: bytes) -> pd.DataFrame:
    df = _read_df(content)
    nome_c = _find_col(df,["nome da maquina","nome","frota","maquina","name"])
    desc_c = _find_col(df,["descri","tipo","grupo"])
    cc_c   = _find_col(df,["centro de custo","cc","cost"])
    if not nome_c: nome_c = df.columns[0]
    r = pd.DataFrame({
        "Nome":     df[nome_c].astype(str).str.strip(),
        "Descricao":df[desc_c].astype(str).str.strip() if desc_c else "",
        "CC":       df[cc_c].astype(str).str.strip()   if cc_c   else "",
    })
    def _tipo(row):
        txt = _norm(row["Descricao"]) + " " + _norm(row["Nome"])
        return "colheita" if any(k in txt for k in KEYWORDS_COLHEITA) else "agro"
    r["Tipo"] = r.apply(_tipo, axis=1)
    return r

def build_device_map(veiculos_df):
    if veiculos_df is None or veiculos_df.empty:
        return DEVICE_MAP_FALLBACK.copy(), {k:"colheita" for k in DEVICE_MAP_FALLBACK}
    dm, tm = {}, {}
    for _, row in veiculos_df.iterrows():
        did  = str(row["DeviceId"])
        nome = str(row["Nome"])
        m    = re.search(r"[Ff]rota\s*(\d+)", nome)
        dm[did] = f"Frota {m.group(1)}" if m else nome
        tm[did] = row["Tipo"]
    return dm, tm

@st.cache_data
def load_apontamentos(content: bytes, device_map: dict, tipo_map: dict, filtrar_tipo: str):
    df = _read_df(content)
    df.columns = ["Apontamento","DeviceId","Codigo","HorasTexto"]
    df["Horas"]     = df["HorasTexto"].apply(parse_horas)
    df["Codigo"]    = df["Codigo"].astype(str).replace("10.004","10.003")
    df = df.groupby(["Apontamento","DeviceId","Codigo"],as_index=False)["Horas"].sum()
    df["Categoria"] = df["Apontamento"].map(CATEGORIA_MAP).fillna("Sem Apontamento")
    df["Frota"]     = df["DeviceId"].map(device_map).fillna(df["DeviceId"])
    df["Tipo"]      = df["DeviceId"].map(tipo_map).fillna("agro")
    return df[df["Tipo"] == filtrar_tipo].copy()

# ══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def calc_disponibilidade(df):
    total = df.groupby("Frota")["Horas"].sum().rename("Total")
    qbr   = df[df["Apontamento"]=="Quebra"].groupby("Frota")["Horas"].sum().rename("Quebra")
    d = pd.concat([total,qbr],axis=1).fillna(0)
    d["Disponibilidade"]   = (d["Total"]-d["Quebra"])/d["Total"]*100
    d["Indisponibilidade"] = d["Quebra"]/d["Total"]*100
    return d.reset_index()

FIG_L = dict(paper_bgcolor="rgba(0,0,0,0)",plot_bgcolor="rgba(0,0,0,0)",
             font=dict(family="Inter,sans-serif",size=12,color="#e0e0e0"))

# ══════════════════════════════════════════════════════════════════════════════
# GRÁFICOS
# ══════════════════════════════════════════════════════════════════════════════
def fig_donut(df):
    agg = df.groupby("Categoria")["Horas"].sum().reset_index()
    agg = agg[agg["Horas"]>0].sort_values("Horas",ascending=False)
    fig = go.Figure(go.Pie(
        labels=agg["Categoria"], values=agg["Horas"].round(1), hole=0.60,
        marker=dict(colors=[CAT_CORES.get(c,"#888") for c in agg["Categoria"]],
                    line=dict(color="rgba(0,0,0,0.3)",width=1)),
        textinfo="percent", textfont=dict(color="#fff"),
        hovertemplate="<b>%{label}</b><br>%{value:.1f}h (%{percent})<extra></extra>",
    ))
    fig.update_layout(**FIG_L,showlegend=True,
        legend=dict(orientation="v",x=1.0,y=0.5,font=dict(color="#e0e0e0")),
        height=280, margin=dict(l=4,r=4,t=4,b=4))
    return fig

def fig_barras_h(df):
    agg = df.groupby(["Apontamento","Categoria"])["Horas"].sum().reset_index()
    agg = agg[agg["Horas"]>0].sort_values("Horas")
    ml  = max((len(str(a)) for a in agg["Apontamento"]),default=10)*7
    fig = go.Figure(go.Bar(
        x=agg["Horas"].round(1), y=agg["Apontamento"], orientation="h",
        marker_color=[APT_CORES.get(a,"#888") for a in agg["Apontamento"]],
        text=agg["Horas"].round(1).astype(str)+"h",
        textposition="outside", textfont=dict(color="#e0e0e0"),
        hovertemplate="<b>%{y}</b><br>%{x:.1f}h<extra></extra>",
    ))
    fig.update_layout(**FIG_L, height=max(300,len(agg)*30+60),
        margin=dict(l=ml+20,r=60,t=10,b=10),
        xaxis=dict(showgrid=True,gridcolor="rgba(255,255,255,0.1)",title="horas",color="#e0e0e0"),
        yaxis=dict(showgrid=False,color="#e0e0e0"))
    return fig

def fig_comparativo(df):
    cats   = ["Colhendo","Aguardando Transbordo","Chuva / Umidade","Quebra","Manobra","Sem Apontamento"]
    frotas = sorted(df["Frota"].unique())
    pal    = ["#1a7a4a","#1a5fa8","#c97d10","#7F77DD","#9b59b6","#e67e22"]
    fig    = go.Figure()
    for i,fr in enumerate(frotas):
        sub = df[df["Frota"]==fr]; tot = sub["Horas"].sum()
        vals  = [round(sub[sub["Apontamento"]==c]["Horas"].sum(),1) for c in cats]
        texts = [f"{v/tot*100:.1f}%" if tot>0 else "" for v in vals]
        fig.add_trace(go.Bar(name=fr,x=cats,y=vals,marker_color=pal[i%len(pal)],
            text=texts,textposition="outside",textfont=dict(color="#e0e0e0",size=9),
            hovertemplate="<b>%{x}</b><br>%{y:.1f}h · %{text}<extra>"+fr+"</extra>"))
    fig.update_layout(**FIG_L,barmode="group",height=350,
        xaxis=dict(tickangle=-20,showgrid=False,color="#e0e0e0"),
        margin=dict(l=4,r=4,t=36,b=4),
        yaxis=dict(showgrid=True,gridcolor="rgba(255,255,255,0.1)",title="horas",color="#e0e0e0"),
        legend=dict(orientation="h",y=1.12,font=dict(color="#e0e0e0")))
    return fig

def fig_perdas(df):
    cats = ["Chuva / Umidade","Sem Apontamento","Quebra","Aguardando Transbordo",
            "Manobra","Manutenção Programada","Abastecimento"]
    agg  = df[df["Apontamento"].isin(cats)].groupby("Apontamento")["Horas"].sum()
    agg  = agg[agg>0].sort_values()
    if agg.empty: return go.Figure()
    ml   = max((len(str(a)) for a in agg.index),default=10)*7
    fig  = go.Figure(go.Bar(
        x=agg.values.round(1),y=agg.index,orientation="h",
        marker_color=[APT_CORES.get(a,"#888") for a in agg.index],
        text=[f"{v:.1f}h" for v in agg.values],
        textposition="outside",textfont=dict(color="#e0e0e0"),
        hovertemplate="<b>%{y}</b><br>%{x:.1f}h<extra></extra>"))
    fig.update_layout(**FIG_L,height=max(260,len(agg)*38+60),
        margin=dict(l=ml+20,r=60,t=10,b=10),
        xaxis=dict(showgrid=True,gridcolor="rgba(255,255,255,0.1)",title="horas",color="#e0e0e0"),
        yaxis=dict(showgrid=False,color="#e0e0e0"))
    return fig

def fig_disponibilidade(disp_df):
    fig = go.Figure()
    fig.add_trace(go.Bar(name="Disponível",x=disp_df["Frota"],y=disp_df["Disponibilidade"].round(1),
        marker_color=C_DISP,text=disp_df["Disponibilidade"].round(1).astype(str)+"%",
        textposition="inside",textfont=dict(color="#fff"),
        hovertemplate="<b>%{x}</b><br>Disp: %{y:.1f}%<extra></extra>"))
    fig.add_trace(go.Bar(name="Quebra",x=disp_df["Frota"],y=disp_df["Indisponibilidade"].round(1),
        marker_color=C_MANUTENCAO,text=disp_df["Indisponibilidade"].round(1).astype(str)+"%",
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
    meds=[row["TempMedMotor"],row["TempMedOleo"]]
    maxs=[row["TempMaxMotor"],row["TempMaxOleo"]]
    mins=[row["TempMinMotor"],row["TempMinOleo"]]
    cores=["#E24B4A","#378ADD"]
    fig=go.Figure()
    for lb,md,mn,mx,co in zip(labels,meds,mins,maxs,cores):
        fig.add_trace(go.Bar(name=lb,x=[lb],y=[md],marker_color=co,
            text=f"méd {md:.0f}° · máx {mx:.0f}°",
            textposition="inside",textfont=dict(color="#fff",size=11),
            error_y=dict(type="data",symmetric=False,array=[mx-md],arrayminus=[max(0,md-mn)]),
            hovertemplate=f"<b>{lb}</b><br>méd {md}°C · máx {mx}°C<extra></extra>"))
    fig.update_layout(**FIG_L,height=260,showlegend=False,
        margin=dict(l=55,r=20,t=20,b=10),
        yaxis=dict(title="°C",showgrid=True,gridcolor="rgba(255,255,255,0.1)",color="#e0e0e0"),
        xaxis=dict(showgrid=False,color="#e0e0e0"))
    return fig

def fig_gauge(valor, titulo, maximo=100, cor=C_PRODUTIVO, sufixo="%"):
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

    # ── Fonte dos dados base (SharePoint) ────────────────────────────────────
    st.markdown("### Dados SharePoint")
    fonte = st.radio("Fonte",["SharePoint (auto)","Upload manual"],label_visibility="collapsed")

    sp_data   = {}
    apt_bytes = None

    if fonte == "SharePoint (auto)":
        with st.spinner("Conectando SharePoint..."):
            sp_data = load_from_sharepoint()
        apt_bytes = sp_data.get("apt")
        if apt_bytes:
            st.success("✓ Conectado")
        else:
            st.error("Falha ao carregar apontamentos.")
    else:
        up_apt  = st.file_uploader("XLSX — Apontamentos",     type=["xlsx","csv"],key="ua")
        up_veic = st.file_uploader("XLSX — Veículos AGRITEL", type=["xlsx","csv"],key="uv")
        up_cc   = st.file_uploader("XLSX — Frotas CC",        type=["xlsx","csv"],key="uc")
        apt_bytes = up_apt.read() if up_apt else None
        if up_veic: sp_data["veiculos"]  = up_veic.read()
        if up_cc:   sp_data["frotas_cc"] = up_cc.read()

    if not apt_bytes:
        st.warning("Aguardando apontamentos.xlsx")
        st.stop()

    # ── Cadastro de frotas ────────────────────────────────────────────────────
    veic_df = load_veiculos(sp_data["veiculos"]) if sp_data.get("veiculos") else None
    cc_df   = load_frotas_cc(sp_data["frotas_cc"]) if sp_data.get("frotas_cc") else None
    device_map, tipo_map = build_device_map(veic_df)
    frotas_colheita = {v for k,v in device_map.items() if tipo_map.get(k)=="colheita"}

    st.markdown("---")

    # ── Seletor de tipo de operação (1 app, 2 visões) ─────────────────────────
    st.markdown("### Operação")
    tipo_opcao = st.radio(
        "Tipo",
        ["🌾  Colheita", "🚜  Agro"],
        horizontal=True,
        label_visibility="collapsed",
        key="tipo_operacao",
    )
    APP_TIPO = "colheita" if "Colheita" in tipo_opcao else "agro"
    CONF     = TIPO_CFG[APP_TIPO]
    st.session_state["tipo_operacao_val"] = APP_TIPO   # persiste para o próximo rerun

    st.markdown("---")
    st.markdown("### Período")
    hoje = datetime.now()
    d_ini_def = hoje.replace(day=1).date()
    d_fim_def = hoje.date()

    col_a, col_b = st.columns(2)
    with col_a:
        data_ini = st.date_input("De",  value=d_ini_def, key="d_ini",
                                 format="DD/MM/YYYY")
    with col_b:
        data_fim = st.date_input("Até", value=d_fim_def, key="d_fim",
                                 format="DD/MM/YYYY")

    if data_ini > data_fim:
        st.error("Data início > data fim.")
        st.stop()

    dias_periodo = (data_fim - data_ini).days + 1
    period_h     = dias_periodo * 24.0   # horas totais do período (por máquina)

    n_chunks = max(1, -(-dias_periodo // 15))   # ceil division
    st.caption(
        f"**{data_ini.strftime('%d/%m/%Y')}** – **{data_fim.strftime('%d/%m/%Y')}**  \n"
        f"**{dias_periodo}** dias · **{period_h:.0f}h**/máquina"
        + (f" · {n_chunks} chamadas API" if n_chunks > 1 else "")
    )

    # ── Buscar dados da API Agritel ───────────────────────────────────────────
    st.markdown("---")
    st.markdown("### API Agritel")
    with st.spinner(f"Buscando {dias_periodo} dias da API..."):
        api_items, api_ok = fetch_agritel(data_ini, data_fim)

    if api_items:
        st.markdown(
            f'<div class="api-badge">✓ {len(api_items)} máquina(s) · cache 10 min</div>',
            unsafe_allow_html=True,
        )
    else:
        st.warning("⚠️ API sem dados — verifique datas ou rate limit.")

    # ── Montar DataFrames ─────────────────────────────────────────────────────
    df_apt  = load_apontamentos(apt_bytes, device_map, tipo_map, APP_TIPO)
    df_perf = build_perf_df(api_items, device_map, tipo_map, APP_TIPO)
    disp_df = calc_disponibilidade(df_apt)

    # ── Imagem do período (opcional) ──────────────────────────────────────────
    if sp_data.get("periodo_img"):
        with st.expander("📋 Ref. Agritel"):
            st.image(sp_data["periodo_img"], use_container_width=True)

    st.markdown("---")
    st.markdown("### Filtros")
    frotas_disp = ["Todas"] + sorted(df_apt["Frota"].unique().tolist())
    frota_sel   = st.selectbox("Frota", frotas_disp)

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
df_f      = df_apt  if frota_sel=="Todas" else df_apt[df_apt["Frota"]==frota_sel]
df_perf_f = df_perf if frota_sel=="Todas" else df_perf[df_perf["NomeFrota"]==frota_sel]

num_frotas = df_f["Frota"].nunique() or 1
total_h    = df_f["Horas"].sum()
col_h      = df_f[df_f["Apontamento"]=="Colhendo"]["Horas"].sum()
man_h      = df_f[df_f["Apontamento"]=="Manobra"]["Horas"].sum()
sa_h       = df_f[df_f["Categoria"]=="Sem Apontamento"]["Horas"].sum()
qbr_h      = df_f[df_f["Apontamento"]=="Quebra"]["Horas"].sum()
chuva_h    = df_f[df_f["Apontamento"]=="Chuva / Umidade"]["Horas"].sum()
agt_h      = df_f[df_f["Apontamento"]=="Aguardando Transbordo"]["Horas"].sum()
improd     = chuva_h + qbr_h + agt_h

safe = lambda n,d: n/d*100 if d>0 else 0
ef_col   = safe(col_h,   total_h)
ef_man   = safe(man_h,   col_h+man_h)
ef_imp   = safe(improd,  total_h)
ef_sa    = safe(sa_h,    total_h)
disp_m   = safe(total_h-qbr_h, total_h)
ef_chuva = safe(chuva_h, total_h)
ef_qbr   = safe(qbr_h,   total_h)
ef_agt   = safe(agt_h,   total_h)

motor_h    = df_perf_f["TempoMotorLigado"].sum() if not df_perf_f.empty else 0.0
elevador_h = df_perf_f["TempoElevador"].sum()    if not df_perf_f.empty else 0.0
ef_trab    = safe(elevador_h, motor_h)

# ══════════════════════════════════════════════════════════════════════════════
# BANNER DO PERÍODO
# ══════════════════════════════════════════════════════════════════════════════
def banner():
    fr_label = "Todas as frotas" if frota_sel=="Todas" else frota_sel
    st.markdown(
        f'<div class="periodo-badge">'
        f'📅 {data_ini.strftime("%d/%m/%Y")} – {data_fim.strftime("%d/%m/%Y")} '
        f'· <b>{dias_periodo}d / {period_h:.0f}h</b> por máquina '
        f'&nbsp;|&nbsp; {CONF["icone"]} {fr_label} · {num_frotas} máquina(s)'
        f'</div>',
        unsafe_allow_html=True,
    )

# ══════════════════════════════════════════════════════════════════════════════
# PÁGINAS
# ══════════════════════════════════════════════════════════════════════════════
if pagina == "Visão geral":
    st.markdown(f"## {CONF['icone']} Visão geral — {CONF['sub']}")
    banner()

    c1,c2,c3,c4,c5 = st.columns(5)
    c1.metric("⏱️ Período / máquina", f"{period_h:.0f}h",
              f"{dias_periodo} dias · total {total_h:.0f}h apontado")
    c2.metric("✅ Disponibilidade",   f"{disp_m:.1f}%",  f"{qbr_h:.1f}h em quebra")
    c3.metric("⚡ Efic. colheita",    f"{ef_col:.1f}%",  f"{col_h:.1f}h colhendo")
    c4.metric("↔️ % Manobra",         f"{ef_man:.1f}%",  f"{man_h:.1f}h")
    c5.metric("❓ Sem apontamento",    f"{ef_sa:.1f}%",   f"{sa_h:.1f}h")

    st.markdown("<p style='margin:14px 0 6px;font-size:11px;color:#888;"
                "text-transform:uppercase;letter-spacing:.08em'>⚙️ Motor & Elevador — via API Agritel</p>",
                unsafe_allow_html=True)
    cm1,cm2,cm3,cm4 = st.columns(4)
    cm1.metric("🔵 Hora Motor Ligado",   f"{motor_h:.1f}h",
               f"{motor_h/num_frotas:.1f}h/máq")
    cm2.metric("🟣 Hora Elevador",       f"{elevador_h:.1f}h",
               f"{elevador_h/num_frotas:.1f}h/máq")
    cm3.metric("📊 Efic. Trabalho",      f"{ef_trab:.1f}%",
               "Elevador ÷ Motor Ligado")
    cons_total = df_perf_f["ConsumoTotal"].sum() if not df_perf_f.empty else 0
    cm4.metric("🛢️ Consumo total",       f"{cons_total:,.0f} L".replace(",","."))

    st.markdown("<p style='margin:14px 0 6px;font-size:11px;color:#888;"
                "text-transform:uppercase;letter-spacing:.08em'>⚠️ Breakdown Improdutivo</p>",
                unsafe_allow_html=True)
    ci0,ci1,ci2,ci3 = st.columns(4)
    ci0.metric("⚠️ Improdutivo total",     f"{ef_imp:.1f}%", f"{improd:.1f}h")
    ci1.metric("🌧️ Chuva / Umidade",      f"{ef_chuva:.1f}%",f"{chuva_h:.1f}h")
    ci2.metric("🔧 Quebra",                f"{ef_qbr:.1f}%",  f"{qbr_h:.1f}h")
    ci3.metric("⏳ Aguardando Transbordo", f"{ef_agt:.1f}%",  f"{agt_h:.1f}h")

    if not df_perf_f.empty:
        cons_med  = df_perf_f["TaxaMedTrab"].mean()
        carga_med = df_perf_f["CargaMed"].mean()
        st.divider()
        c6,c7,c8,c9 = st.columns(4)
        c6.metric("🛢️ Consumo médio trab.", f"{cons_med:.1f} L/h")
        c7.metric("🌡️ Temp. média motor",
                  f"{df_perf_f['TempMedMotor'].mean():.1f}°C")
        c8.metric("🌡️ Temp. máx. motor",
                  f"{df_perf_f['TempMaxMotor'].max():.1f}°C")
        c9.metric("🌡️ Temp. máx. óleo",
                  f"{df_perf_f['TempMaxOleo'].max():.1f}°C")

    st.divider()
    if not df_perf_f.empty:
        st.markdown("**Hora Motor Ligado vs Hora Elevador**")
        st.plotly_chart(fig_motor_elevador(df_perf_f),
                        width="stretch",config={"displayModeBar":False})
        st.divider()

    st.markdown("**Disponibilidade mecânica por frota**")
    st.plotly_chart(fig_disponibilidade(disp_df),width="stretch",config={"displayModeBar":False})
    st.divider()

    cl,cr = st.columns(2)
    with cl:
        st.markdown("**Distribuição por categoria**")
        st.plotly_chart(fig_donut(df_f),width="stretch",config={"displayModeBar":False})
    with cr:
        st.markdown("**Comparativo entre frotas**")
        tots = df_apt.groupby("Frota")["Horas"].sum()
        kc   = st.columns(len(tots)+1)
        kc[0].metric("⏱️ Total", f"{df_apt['Horas'].sum():.1f}h")
        for j,(fr,hr) in enumerate(tots.items()):
            kc[j+1].metric(f"⏱️ {fr}", f"{hr:.1f}h")
        st.plotly_chart(fig_comparativo(df_apt),width="stretch",config={"displayModeBar":False})

    st.markdown("**Perdas e paradas — horas acumuladas**")
    st.plotly_chart(fig_perdas(df_f),width="stretch",config={"displayModeBar":False})


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
    st.plotly_chart(fig_barras_h(df_f),width="stretch",config={"displayModeBar":False})
    st.divider()
    cl,cr = st.columns(2)
    with cl:
        st.markdown("**Distribuição por categoria**")
        st.plotly_chart(fig_donut(df_f),width="stretch",config={"displayModeBar":False})
    with cr:
        st.markdown("**Resumo por categoria**")
        cats = df_f.groupby("Categoria")["Horas"].sum().reset_index()
        cats = cats[cats["Horas"]>0].sort_values("Horas",ascending=False)
        cats["pct"] = (cats["Horas"]/total_h*100).round(1).astype(str)+"%"
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
                f'</div></div>',unsafe_allow_html=True)


elif pagina == "Telemetria":
    st.markdown(f"## {CONF['icone']} Telemetria — API Agritel")
    banner()
    st.markdown(
        '<div class="api-badge">📡 Dados direto da API · Motor, Elevador, Combustível, Temperaturas</div>',
        unsafe_allow_html=True)

    if df_perf_f.empty:
        st.warning("Nenhum dado de telemetria para este período/frota.")
        st.stop()

    cols = st.columns(len(df_perf_f))
    for i,(_,row) in enumerate(df_perf_f.iterrows()):
        with cols[i]:
            st.markdown(f"**{row['NomeFrota']}**")
            d_row = disp_df[disp_df["Frota"]==row["NomeFrota"].replace("Frota ","")]
            if not d_row.empty:
                dv = d_row.iloc[0]["Disponibilidade"]
                qv = d_row.iloc[0]["Quebra"]
                st.markdown(
                    f'<div style="padding:8px 12px;border-radius:8px;'
                    f'background:rgba(39,174,96,.15);border-left:3px solid {C_DISP};margin-bottom:8px">'
                    f'<span style="font-size:11px;color:#aaa">Disponibilidade mecânica</span><br>'
                    f'<span style="font-size:20px;font-weight:700;color:{C_DISP}">{dv:.1f}%</span>'
                    f'<span style="font-size:10px;color:#888;margin-left:8px">({qv:.1f}h quebra)</span>'
                    f'</div>',unsafe_allow_html=True)
            ef_el = safe(row["TempoElevador"],row["TempoMotorLigado"])
            st.markdown(
                f'<div style="padding:8px 12px;border-radius:8px;'
                f'background:rgba(155,89,182,.15);border-left:3px solid {C_ELEVADOR};margin-bottom:8px">'
                f'<span style="font-size:11px;color:#aaa">Efic. Trabalho (Elev./Motor)</span><br>'
                f'<span style="font-size:20px;font-weight:700;color:{C_ELEVADOR}">{ef_el:.1f}%</span>'
                f'<span style="font-size:10px;color:#888;margin-left:8px">'
                f'({row["TempoElevador"]:.1f}h / {row["TempoMotorLigado"]:.1f}h)</span>'
                f'</div>',unsafe_allow_html=True)
            for lb,val,cor in [
                ("Motor ligado",        f"{row['TempoMotorLigado']:.1f}h",  C_MOTOR),
                ("Elevador ligado",     f"{row['TempoElevador']:.1f}h",     C_ELEVADOR),
                ("Em trabalho",         f"{row['TempoTrabalho']:.1f}h",     "#e0e0e0"),
                ("Manobra",             f"{row['TempoManobra']:.1f}h",      "#e0e0e0"),
                ("Ocioso",              f"{row['TempoOcioso']:.1f}h",       "#e0e0e0"),
                ("Deslocando",          f"{row['TempoDeslocando']:.1f}h",   "#e0e0e0"),
                ("Consumo total",       f"{int(row['ConsumoTotal']):,} L".replace(",","."), "#e0e0e0"),
                ("Taxa cons. (trab.)",  f"{row['TaxaMedTrab']:.1f} L/h",   "#e0e0e0"),
                ("Temp. méd. motor",    f"{row['TempMedMotor']:.1f}°C",     "#E24B4A"),
                ("Temp. máx. motor",    f"{row['TempMaxMotor']:.1f}°C",     "#E24B4A"),
                ("Temp. méd. óleo",     f"{row['TempMedOleo']:.1f}°C",      "#378ADD"),
                ("Temp. máx. óleo",     f"{row['TempMaxOleo']:.1f}°C",      "#378ADD"),
            ]:
                st.markdown(
                    f'<div style="display:flex;justify-content:space-between;padding:6px 10px;'
                    f'background:rgba(255,255,255,.05);border-radius:6px;margin-bottom:4px">'
                    f'<span style="font-size:11px;color:#999">{lb}</span>'
                    f'<span style="font-size:12px;font-weight:500;color:{cor}">{val}</span>'
                    f'</div>',unsafe_allow_html=True)

    st.divider()
    st.markdown("**Temperaturas — méd / máx**")
    tc = st.columns(len(df_perf_f))
    for i,(_,row) in enumerate(df_perf_f.iterrows()):
        with tc[i]:
            st.markdown(f"*{row['NomeFrota']}*")
            st.plotly_chart(fig_temperaturas(row),width="stretch",config={"displayModeBar":False})

    st.divider()
    st.markdown("**Gauges operacionais**")
    gc = st.columns(len(df_perf_f)*3)
    idx = 0
    for _,row in df_perf_f.iterrows():
        ef_m  = safe(row["TempoTrabalho"], row["TempoMotorLigado"])
        ef_el = safe(row["TempoElevador"], row["TempoMotorLigado"])
        oc    = safe(row["TempoOcioso"],   row["TempoMotorLigado"])
        with gc[idx]:
            st.plotly_chart(fig_gauge(ef_m,  f"{row['NomeFrota']}\nEfic. Motor",    cor=C_PRODUTIVO),
                            width="stretch",config={"displayModeBar":False})
        with gc[idx+1]:
            st.plotly_chart(fig_gauge(ef_el, f"{row['NomeFrota']}\nEfic. Elevador", cor=C_ELEVADOR),
                            width="stretch",config={"displayModeBar":False})
        with gc[idx+2]:
            st.plotly_chart(fig_gauge(oc,    f"{row['NomeFrota']}\nOcioso",         cor=C_MANUTENCAO),
                            width="stretch",config={"displayModeBar":False})
        idx += 3


elif pagina == "Por frota":
    frotas_lista = sorted(df_apt["Frota"].unique().tolist())
    fr4 = st.selectbox("Selecione a frota", frotas_lista, key="fr4")
    sub_apt  = df_apt[df_apt["Frota"]==fr4]
    sub_perf = df_perf[df_perf["NomeFrota"]==fr4] if not df_perf.empty else pd.DataFrame()
    sub_disp = disp_df[disp_df["Frota"]==fr4.replace("Frota ","")]

    st.markdown(f"## {CONF['icone']} {fr4} — Análise individual")
    banner()

    tot_f = sub_apt["Horas"].sum()
    col_f = sub_apt[sub_apt["Apontamento"]=="Colhendo"]["Horas"].sum()
    man_f = sub_apt[sub_apt["Apontamento"]=="Manobra"]["Horas"].sum()
    qbr_f = sub_apt[sub_apt["Apontamento"]=="Quebra"]["Horas"].sum()
    d_f   = safe(tot_f-qbr_f, tot_f)

    k1,k2,k3,k4 = st.columns(4)
    k1.metric("✅ Disponibilidade",  f"{d_f:.1f}%",
              f"{qbr_f:.1f}h quebra · {tot_f:.1f}h total")
    k2.metric("⚡ Efic. colheita",
              f"{safe(col_f,tot_f):.1f}%", f"{col_f:.1f}h")
    k3.metric("↔️ % Manobra",
              f"{safe(man_f,col_f+man_f):.1f}%", f"{man_f:.1f}h")
    k4.metric("⏱️ Período ref.",     f"{period_h:.0f}h",
              f"{tot_f:.1f}h apontado")

    if not sub_perf.empty:
        row = sub_perf.iloc[0]
        ef_el = safe(row["TempoElevador"], row["TempoMotorLigado"])
        st.divider()
        k5,k6,k7,k8 = st.columns(4)
        k5.metric("🔵 Motor Ligado",   f"{row['TempoMotorLigado']:.1f}h")
        k6.metric("🟣 Hora Elevador",  f"{row['TempoElevador']:.1f}h",
                  f"Efic. {ef_el:.1f}%")
        k7.metric("🛢️ Consumo",        f"{int(row['ConsumoTotal']):,} L".replace(",","."),
                  f"{row['TaxaMedTrab']:.1f} L/h")
        oc = safe(row["TempoOcioso"], row["TempoMotorLigado"])
        k8.metric("⏸️ Ocioso",         f"{oc:.1f}%",
                  f"{row['TempoOcioso']:.1f}h")

    st.divider()
    cl,cr = st.columns(2)
    with cl:
        st.markdown("**Uso do tempo**")
        st.plotly_chart(fig_donut(sub_apt),width="stretch",config={"displayModeBar":False})
    with cr:
        st.markdown("**Apontamentos detalhados**")
        st.plotly_chart(fig_barras_h(sub_apt),width="stretch",config={"displayModeBar":False})

    if not sub_perf.empty:
        st.divider()
        st.markdown("**Temperaturas**")
        st.plotly_chart(fig_temperaturas(sub_perf.iloc[0]),
                        width="stretch",config={"displayModeBar":False})