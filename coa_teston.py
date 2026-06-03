"""
COA · TESTON — Dashboard Safra 2026/27
Streamlit + Plotly | Login + SharePoint
"""

import re, io, os, pathlib, hashlib
import pandas as pd
import numpy as np
import streamlit as st
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# ══════════════════════════════════════════════════════════════════════════════
# CONFIGURAÇÃO — preencha as credenciais aqui ou via variáveis de ambiente
# ══════════════════════════════════════════════════════════════════════════════
SP_TENANT_ID     = os.environ.get("SP_TENANT_ID",     "SEU_TENANT_ID_AQUI")
SP_CLIENT_ID     = os.environ.get("SP_CLIENT_ID",     "SEU_CLIENT_ID_AQUI")
SP_CLIENT_SECRET = os.environ.get("SP_CLIENT_SECRET", "SEU_CLIENT_SECRET_AQUI")
SP_SITE_URL      = os.environ.get("SP_SITE_URL",      "https://SEU_TENANT.sharepoint.com/sites/SEU_SITE")
SP_APT_PATH      = os.environ.get("SP_APT_PATH",      "/sites/SEU_SITE/Documentos Compartilhados/apontamentos.csv")
SP_PERF_PATH     = os.environ.get("SP_PERF_PATH",     "/sites/SEU_SITE/Documentos Compartilhados/performance.csv")

# Senha do dashboard — altere para a senha desejada
# Gere o hash em Python: import hashlib; hashlib.sha256("sua_senha".encode()).hexdigest()
SENHA_HASH = os.environ.get(
    "DASHBOARD_SENHA_HASH",
    hashlib.sha256("teston2026".encode()).hexdigest()  # senha padrão: teston2026
)

# ── PAGE CONFIG ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="COA · TESTON — Safra 2026/27",
    page_icon="🌿",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CORES ─────────────────────────────────────────────────────────────────────
C_PRODUTIVO  = "#1a7a4a"
C_MANUTENCAO = "#c0392b"
C_ESPERA     = "#c97d10"
C_MANOBRA    = "#7F77DD"
C_PARADA     = "#888780"
C_SEM_APT    = "#b4b2a9"
C_F1262      = "#1a7a4a"
C_F1263      = "#1a5fa8"
C_DISP       = "#27ae60"

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

DEVICE_MAP = {
    "671a433018212ed9057ce3ee": "Frota 1262",
    "671a433018212ed9057ce3f1": "Frota 1263",
}

# ── CSS GLOBAL ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
[data-testid="stMetricValue"]       { font-size:1.6rem !important; font-weight:700; }
[data-testid="stMetricLabel"]       { font-size:0.72rem !important; text-transform:uppercase; letter-spacing:.05em; }
[data-testid="stMetricDelta"]       { font-size:0.72rem !important; }
[data-testid="stSidebar"]           { background-color:#1c1e18; }
[data-testid="stSidebar"] *         { color:#d8dbd4 !important; }
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3        { color:#a8c832 !important; }
div[data-testid="metric-container"] { background:rgba(255,255,255,0.05); border-radius:10px; padding:12px 16px; border-left:3px solid #1a7a4a; }
.block-container                    { padding-top:1.4rem; }
.js-plotly-plot .plotly,
.js-plotly-plot .plotly .svg-container { background:transparent !important; }
/* Tela de login */
.login-wrap { max-width:400px; margin:80px auto; padding:40px; background:rgba(255,255,255,0.04); border-radius:16px; border:0.5px solid rgba(255,255,255,0.1); }
.login-title { font-size:24px; font-weight:700; color:#a8c832; margin-bottom:4px; }
.login-sub   { font-size:13px; color:#888; margin-bottom:28px; }
.login-err   { color:#c0392b; font-size:13px; margin-top:8px; }
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# TELA DE LOGIN
# ══════════════════════════════════════════════════════════════════════════════
def check_login():
    """Retorna True se o usuário está autenticado."""
    return st.session_state.get("autenticado", False)

def tela_login():
    st.markdown("""
    <div class="login-wrap">
        <div class="login-title">🌿 COA · TESTON</div>
        <div class="login-sub">Dashboard Operacional — Safra 2026/27</div>
    </div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("###")
        st.markdown("#### 🔒 Acesso restrito")
        senha = st.text_input("Senha", type="password", placeholder="Digite a senha de acesso",
                              key="input_senha")
        entrar = st.button("Entrar", use_container_width=True, type="primary")

        if entrar:
            senha_hash = hashlib.sha256(senha.encode()).hexdigest()
            if senha_hash == SENHA_HASH:
                st.session_state["autenticado"] = True
                st.rerun()
            else:
                st.error("Senha incorreta. Tente novamente.")

        st.markdown("""
        <div style='text-align:center;margin-top:24px;font-size:11px;color:#555'>
            MS Colheitas e Serviços · Naviraí/MS · Brasil
        </div>
        """, unsafe_allow_html=True)

# Bloqueia tudo se não estiver logado
if not check_login():
    tela_login()
    st.stop()


# ══════════════════════════════════════════════════════════════════════════════
# SHAREPOINT — carregamento dos CSVs
# ══════════════════════════════════════════════════════════════════════════════
@st.cache_data(ttl=600)  # cache de 10 minutos
def load_from_sharepoint():
    """
    Carrega CSVs do SharePoint via Microsoft Graph API REST.
    Usa apenas requests — sem dependência do Office365 library.
    Requer: Sites.Read.All no app Azure AD.
    """
    try:
        import requests

        # 1. Obtém token OAuth2 via client credentials
        token_url = f"https://login.microsoftonline.com/{SP_TENANT_ID}/oauth2/v2.0/token"
        token_resp = requests.post(token_url, data={
            "grant_type":    "client_credentials",
            "client_id":     SP_CLIENT_ID,
            "client_secret": SP_CLIENT_SECRET,
            "scope":         "https://graph.microsoft.com/.default",
        })
        token_resp.raise_for_status()
        access_token = token_resp.json()["access_token"]
        headers = {"Authorization": f"Bearer {access_token}"}

        # 2. Descobre o site ID via Graph
        hostname = SP_SITE_URL.replace("https://", "").split("/")[0]
        site_path = "/".join(SP_SITE_URL.replace("https://", "").split("/")[1:])
        site_resp = requests.get(
            f"https://graph.microsoft.com/v1.0/sites/{hostname}:/{site_path}",
            headers=headers
        )
        site_resp.raise_for_status()
        site_id = site_resp.json()["id"]

        # 3. Obtém drive raiz do site
        drive_resp = requests.get(
            f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive",
            headers=headers
        )
        drive_resp.raise_for_status()
        drive_id = drive_resp.json()["id"]

        # 4. Baixa cada arquivo pelo caminho relativo
        def _download(sp_path):
            # Converte caminho SharePoint para caminho relativo ao drive
            # ex: /sites/MSCOLHEITA/Documentos Compartilhados/VITOR/AGRITEL/apontamentos.csv
            # → VITOR/AGRITEL/apontamentos.csv
            parts = sp_path.replace("/sites/MSCOLHEITA/Documentos Compartilhados/", "").strip("/")
            url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{parts}:/content"
            resp = requests.get(url, headers=headers)
            resp.raise_for_status()
            return resp.content

        apt_bytes  = _download(SP_APT_PATH)
        perf_bytes = _download(SP_PERF_PATH)
        return apt_bytes, perf_bytes

    except Exception as e:
        import traceback
        st.error(f"Erro SharePoint (Graph API): {str(e)}\n\n{traceback.format_exc()}")
        return None, None


def load_fallback_local():
    """Fallback: lê os CSVs da pasta data/ local."""
    DATA_DIR  = pathlib.Path(__file__).parent / "data"
    apt_file  = DATA_DIR / "apontamentos.csv"
    perf_file = DATA_DIR / "performance.csv"
    apt_bytes  = apt_file.read_bytes()  if apt_file.exists()  else None
    perf_bytes = perf_file.read_bytes() if perf_file.exists() else None
    return apt_bytes, perf_bytes


# ── PARSERS ───────────────────────────────────────────────────────────────────
def parse_horas(txt) -> float:
    if not isinstance(txt, str):
        return float(txt) if txt else 0.0
    m = re.search(r"([\d.,]+)\s*h", txt)
    if m:
        return float(m.group(1).replace(".", "").replace(",", "."))
    return 0.0

def parse_num(txt) -> float:
    if not isinstance(txt, str):
        return float(txt) if txt else 0.0
    cleaned = re.sub(r"[^\d,.]", "", txt).replace(".", "").replace(",", ".")
    try:
        return float(cleaned)
    except:
        return 0.0

def parse_tempo_perf(txt) -> float:
    if not isinstance(txt, str):
        return 0.0
    m = re.search(r"(\d+)[,.]?(\d*)\s*h", txt)
    if not m:
        return 0.0
    return float(m.group(1)) + (float(m.group(2)) / 60 if m.group(2) else 0.0)


@st.cache_data
def load_apontamentos(content: bytes) -> pd.DataFrame:
    # Detecta formato pelo header dos bytes
    is_xlsx = content[:4] == b'PK\x03\x04'  # assinatura ZIP/XLSX

    if is_xlsx:
        # XLSX: valores já são decimais limpos
        df = pd.read_excel(io.BytesIO(content))
        df.columns = ["Apontamento", "DeviceId", "Codigo", "Horas"]
        df["Horas"] = pd.to_numeric(df["Horas"], errors="coerce").fillna(0.0)
    else:
        # CSV: detecta se novo formato (linhas encapsuladas) ou padrão
        raw = content.decode("utf-8-sig", errors="replace")
        linhas = raw.strip().split("\n")
        primeira_dado = linhas[1].strip() if len(linhas) > 1 else ""
        novo_formato = primeira_dado.startswith('"') and '""' in primeira_dado

        if novo_formato:
            rows = []
            for line in linhas:
                line = line.strip().strip("\r")
                if line.startswith("Name") or not line:
                    continue
                if line.startswith('"') and line.endswith('"'):
                    line = line[1:-1]
                line = line.replace('""', "\x00")
                parts = line.split(",")
                if len(parts) >= 4:
                    rows.append([parts[0], parts[1], parts[2],
                                  ",".join(parts[3:]).replace("\x00", '"').strip('"')])
            df = pd.DataFrame(rows, columns=["Apontamento","DeviceId","Codigo","HorasTexto"])
        else:
            df = pd.read_csv(io.BytesIO(content))
            df.columns = ["Apontamento","DeviceId","Codigo","HorasTexto"]

        df["Horas"] = df["HorasTexto"].apply(parse_horas)

    df["Codigo"] = df["Codigo"].astype(str).str.replace(".0","",regex=False)
    # Consolida Sem Apontamento (10003 e 10004)
    df.loc[df["Codigo"] == "10004", "Codigo"] = "10003"
    df = df.groupby(["Apontamento","DeviceId","Codigo"], as_index=False)["Horas"].sum()
    df["Categoria"] = df["Apontamento"].map(CATEGORIA_MAP).fillna("Sem Apontamento")
    df["Frota"]     = df["DeviceId"].map(DEVICE_MAP).fillna(df["DeviceId"])
    return df

@st.cache_data
def load_performance(content: bytes) -> pd.DataFrame:
    is_xlsx = content[:4] == b'PK\x03\x04'
    if is_xlsx:
        df = pd.read_excel(io.BytesIO(content))
    else:
        df = pd.read_csv(io.BytesIO(content))
    c = df.columns.tolist()
    def col(name):
        if name not in c:
            return pd.Series(0.0, index=df.index)
        if is_xlsx:
            return pd.to_numeric(df[name], errors="coerce").fillna(0.0)
        return df[name].apply(parse_num)
    return pd.DataFrame({
        "NomeFrota":        df["Nome da Máquina"].str.extract(r"(Frota \d+)")[0].fillna(df["Nome da Máquina"]),
        "Frota":            df["Nome da Máquina"].str.extract(r"Frota (\d+)")[0],
        "TempoMotorLigado": df["Tempo Motor Ligado:"].apply(parse_tempo_perf),
        "TempoTrabalho":    df["Tempo em Trabalho"].apply(parse_tempo_perf),
        "TempoManobra":     df["Tempo em Manobra:"].apply(parse_tempo_perf),
        "TempoOcioso":      df["Tempo com Motor Ocioso:"].apply(parse_tempo_perf),
        "ConsumoTotal":     col("Consumo Total:"),
        "ConsumoTrabalho":  col("Consumo em Trabalho:"),
        "TaxaMedConsumo":   col("Taxa Média de Consumo de Combustível"),
        "TaxaMedTrab":      col("Taxa Média de Consumo de Combustível em Trabalho"),
        "CargaMax":         col("Carga Máxima do Motor"),
        "CargaMed":         col("Carga Média do Motor"),
        "TempMinMotor":     col("Temperatura Mínima do Motor"),
        "TempMaxMotor":     col("Temperatura Máxima do Motor"),
        "TempMedMotor":     col("Temperatura Média do Motor"),
        "TempMinOleo":      col("Temperatura Mínima do Óleo Hidráulico"),
        "TempMaxOleo":      col("Temperatura Máxima do Óleo Hidráulico"),
        "TempMedOleo":      col("Temperatura Média do Óleo Hidráulico"),
        "Odometro":         col("Odômetro de Trabalho"),
        "VelocidadeMax":    col("Velocidade Máxima de Trabalho"),
        "VelocidadeMed":    col("Velocidade Média de Trabalho:"),
        "RpmMotor":         col("Rpm Médio em Trabalho:"),
        "RpmExtMed":        col("Rpm Médio de Exaustor Primário em Trabalho"),
        "PressaoMed":       col("Pressão Média de Corte Base em Trabalho"),
        "Embuchamentos":    col("Número de Embuchamentos Detectados"),
        "TempoElevador":    df["Tempo de Elevador Ligado"].apply(parse_tempo_perf),
    })

def calc_disponibilidade(df: pd.DataFrame) -> pd.DataFrame:
    total  = df.groupby("Frota")["Horas"].sum().rename("Total")
    quebra = df[df["Apontamento"] == "Quebra"].groupby("Frota")["Horas"].sum().rename("Quebra")
    disp   = pd.concat([total, quebra], axis=1).fillna(0)
    disp["Disponibilidade"]   = (disp["Total"] - disp["Quebra"]) / disp["Total"] * 100
    disp["Indisponibilidade"] = disp["Quebra"] / disp["Total"] * 100
    return disp.reset_index()


# ── LAYOUT PLOTLY ─────────────────────────────────────────────────────────────
FIG_LAYOUT = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font=dict(family="Inter, sans-serif", size=12, color="#e0e0e0"),
)

def fig_donut(df):
    agg = df.groupby("Categoria")["Horas"].sum().reset_index()
    agg = agg[agg["Horas"] > 0].sort_values("Horas", ascending=False)
    total = agg["Horas"].sum()
    fig = go.Figure(go.Pie(
        labels=agg["Categoria"], values=agg["Horas"].round(1), hole=0.60,
        marker=dict(colors=[CAT_CORES.get(c,"#888") for c in agg["Categoria"]],
                    line=dict(color="rgba(0,0,0,0.3)", width=1)),
        textinfo="percent", textfont=dict(color="#fff"),
        hovertemplate="<b>%{label}</b><br>%{value:.1f}h (%{percent})<extra></extra>",
    ))
    fig.update_layout(**FIG_LAYOUT, height=280, showlegend=True,
        legend=dict(orientation="v", x=1.0, y=0.5, font=dict(color="#e0e0e0")),
        margin=dict(l=4, r=4, t=4, b=4))
    return fig

def fig_barras_h(df):
    agg = df.groupby(["Apontamento","Categoria"])["Horas"].sum().reset_index()
    agg = agg[agg["Horas"] > 0].sort_values("Horas")
    if agg.empty:
        return go.Figure()
    max_label = max(len(str(a)) for a in agg["Apontamento"]) * 7
    fig = go.Figure(go.Bar(
        x=agg["Horas"].round(1), y=agg["Apontamento"], orientation="h",
        marker_color=[APT_CORES.get(a,"#888") for a in agg["Apontamento"]],
        text=agg["Horas"].round(1).astype(str)+"h", textposition="outside",
        textfont=dict(color="#e0e0e0"),
        hovertemplate="<b>%{y}</b><br>%{x:.1f}h<extra></extra>",
    ))
    fig.update_layout(**FIG_LAYOUT,
        height=max(300, len(agg)*30+60),
        margin=dict(l=max_label+20, r=60, t=10, b=10),
        xaxis=dict(showgrid=True, gridcolor="rgba(255,255,255,0.1)",
                   title="horas", color="#e0e0e0", zerolinecolor="rgba(255,255,255,0.2)"),
        yaxis=dict(showgrid=False, color="#e0e0e0"))
    return fig

def fig_comparativo(df):
    cats   = ["Colhendo","Aguardando Transbordo","Chuva / Umidade","Quebra","Manobra","Sem Apontamento"]
    frotas = sorted(df["Frota"].unique())
    cores  = [C_F1262, C_F1263]
    fig = go.Figure()
    for i, frota in enumerate(frotas):
        sub  = df[df["Frota"]==frota]
        vals = [round(sub[sub["Apontamento"]==c]["Horas"].sum(),1) for c in cats]
        fig.add_trace(go.Bar(
            name=frota, x=cats, y=vals, marker_color=cores[i%2],
            hovertemplate="<b>%{x}</b><br>%{y:.1f}h<extra>"+frota+"</extra>"))
    fig.update_layout(**FIG_LAYOUT, barmode="group", height=300,
        margin=dict(l=4, r=4, t=4, b=4),
        xaxis=dict(tickangle=-20, showgrid=False, color="#e0e0e0"),
        yaxis=dict(showgrid=True, gridcolor="rgba(255,255,255,0.1)",
                   title="horas", color="#e0e0e0"),
        legend=dict(orientation="h", y=1.1, font=dict(color="#e0e0e0")))
    return fig

def fig_perdas(df):
    cats = ["Chuva / Umidade","Sem Apontamento","Quebra","Aguardando Transbordo",
            "Manobra","Manutenção Programada","Abastecimento"]
    agg  = df[df["Apontamento"].isin(cats)].groupby("Apontamento")["Horas"].sum()
    agg  = agg[agg>0].sort_values()
    if agg.empty:
        return go.Figure()
    max_label = max(len(str(a)) for a in agg.index) * 7
    fig = go.Figure(go.Bar(
        x=agg.values.round(1), y=agg.index, orientation="h",
        marker_color=[APT_CORES.get(a,"#888") for a in agg.index],
        text=[f"{v:.1f}h" for v in agg.values], textposition="outside",
        textfont=dict(color="#e0e0e0"),
        hovertemplate="<b>%{y}</b><br>%{x:.1f}h<extra></extra>"))
    fig.update_layout(**FIG_LAYOUT,
        height=max(260, len(agg)*38+60),
        margin=dict(l=max_label+20, r=60, t=10, b=10),
        xaxis=dict(showgrid=True, gridcolor="rgba(255,255,255,0.1)",
                   title="horas", color="#e0e0e0"),
        yaxis=dict(showgrid=False, color="#e0e0e0"))
    return fig

def fig_disponibilidade(disp_df):
    fig = go.Figure()
    fig.add_trace(go.Bar(
        name="Disponível", x=disp_df["Frota"], y=disp_df["Disponibilidade"].round(1),
        marker_color=C_DISP,
        text=disp_df["Disponibilidade"].round(1).astype(str)+"%",
        textposition="inside", textfont=dict(color="#fff"),
        hovertemplate="<b>%{x}</b><br>Disponibilidade: %{y:.1f}%<extra></extra>"))
    fig.add_trace(go.Bar(
        name="Quebra", x=disp_df["Frota"], y=disp_df["Indisponibilidade"].round(1),
        marker_color=C_MANUTENCAO,
        text=disp_df["Indisponibilidade"].round(1).astype(str)+"%",
        textposition="inside", textfont=dict(color="#fff"),
        hovertemplate="<b>%{x}</b><br>Quebra: %{y:.1f}%<extra></extra>"))
    fig.update_layout(**FIG_LAYOUT, barmode="stack", height=260,
        margin=dict(l=4, r=4, t=4, b=4),
        xaxis=dict(showgrid=False, color="#e0e0e0"),
        yaxis=dict(showgrid=True, gridcolor="rgba(255,255,255,0.1)",
                   title="%", range=[0,105], color="#e0e0e0"),
        legend=dict(orientation="h", y=1.1, font=dict(color="#e0e0e0")))
    return fig

def fig_temperaturas(row):
    labels = ["Motor","Óleo Hidráulico"]
    meds   = [row["TempMedMotor"],  row["TempMedOleo"]]
    mins_  = [row["TempMinMotor"],  row["TempMinOleo"]]
    maxs_  = [row["TempMaxMotor"],  row["TempMaxOleo"]]
    cores  = ["#E24B4A","#378ADD"]
    fig = go.Figure()
    for label, med, mn, mx, cor in zip(labels, meds, mins_, maxs_, cores):
        texto = f"mín {mn:.0f}° · méd {med:.0f}° · máx {mx:.0f}°"
        fig.add_trace(go.Bar(
            name=label, x=[label], y=[med], marker_color=cor,
            text=texto, textposition="inside", textfont=dict(color="#fff", size=12),
            error_y=dict(type="data", symmetric=False,
                         array=[mx-med], arrayminus=[med-mn]),
            hovertemplate=f"<b>{label}</b><br>min {mn}°C · méd {med}°C · máx {mx}°C<extra></extra>"))
    fig.update_layout(**FIG_LAYOUT, height=260, showlegend=False,
        margin=dict(l=55, r=20, t=20, b=10),
        yaxis=dict(title="°C", showgrid=True,
                   gridcolor="rgba(255,255,255,0.1)", color="#e0e0e0"),
        xaxis=dict(showgrid=False, color="#e0e0e0"))
    return fig

def fig_gauge(valor, titulo, maximo=100, cor=C_PRODUTIVO, sufixo="%"):
    fig = go.Figure(go.Indicator(
        mode="gauge+number", value=round(valor,1),
        number=dict(suffix=sufixo, font=dict(size=22, color=cor)),
        gauge=dict(
            axis=dict(range=[0,maximo], tickwidth=1,
                      tickcolor="rgba(255,255,255,0.3)",
                      tickfont=dict(color="#e0e0e0")),
            bar=dict(color=cor, thickness=0.6),
            bgcolor="rgba(255,255,255,0.05)", borderwidth=0,
            steps=[dict(range=[0,maximo], color="rgba(255,255,255,0.05)")]),
        title=dict(text=titulo, font=dict(size=12, color="#aaaaaa"))))
    fig.update_layout(**FIG_LAYOUT, height=200,
        margin=dict(l=20, r=20, t=40, b=20))
    return fig


# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("## COA · TESTON")
    st.markdown("**Safra 2026/27** · Gestão de Colhedoras")
    st.markdown("---")

    # ── Fonte dos dados ────────────────────────────────────────────────────
    st.markdown("### Dados")

    # Verifica se credenciais Azure estão configuradas
    azure_ok = "SEU_TENANT_ID" not in SP_TENANT_ID and SP_CLIENT_ID != "SEU_CLIENT_ID_AQUI"

    if azure_ok:
        fonte_opcoes = ["SharePoint (automático)", "Upload manual"]
    else:
        fonte_opcoes = ["Arquivos padrão (data/)", "Upload manual"]
        st.caption("⚠️ Azure não configurado — usando dados locais")

    fonte = st.radio("Fonte", fonte_opcoes, label_visibility="collapsed")

    apt_bytes = perf_bytes = None

    if "SharePoint" in fonte:
        with st.spinner("Conectando ao SharePoint..."):
            apt_bytes, perf_bytes = load_from_sharepoint()
        if apt_bytes:
            st.success("✓ SharePoint conectado")
        else:
            st.error("Falha no SharePoint — verifique as credenciais")

    elif "Upload" in fonte:
        up_apt  = st.file_uploader("CSV Apontamentos", type="csv", key="up_apt")
        up_perf = st.file_uploader("CSV Performance",  type="csv", key="up_perf")
        apt_bytes  = up_apt.read()  if up_apt  else None
        perf_bytes = up_perf.read() if up_perf else None

    else:
        apt_bytes, perf_bytes = load_fallback_local()

    if not apt_bytes or not perf_bytes:
        st.warning("Sem dados disponíveis. Configure o SharePoint ou faça upload.")
        st.stop()

    df_apt  = load_apontamentos(apt_bytes)
    df_perf = load_performance(perf_bytes)
    disp_df = calc_disponibilidade(df_apt)

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
    # Botão de logout
    if st.button("🚪 Sair", use_container_width=True):
        st.session_state["autenticado"] = False
        st.rerun()

    st.caption("MS Colheitas e Serviços · TESTON · 2026")


# ── DADOS FILTRADOS ───────────────────────────────────────────────────────────
df_f    = df_apt if frota_sel == "Todas" else df_apt[df_apt["Frota"]==frota_sel]
total_h = df_f["Horas"].sum()
col_h   = df_f[df_f["Apontamento"]=="Colhendo"]["Horas"].sum()
man_h   = df_f[df_f["Apontamento"]=="Manobra"]["Horas"].sum()
sa_h    = df_f[df_f["Categoria"]=="Sem Apontamento"]["Horas"].sum()
qbr_h   = df_f[df_f["Apontamento"]=="Quebra"]["Horas"].sum()
chuva_h = df_f[df_f["Apontamento"]=="Chuva / Umidade"]["Horas"].sum()
agt_h   = df_f[df_f["Apontamento"]=="Aguardando Transbordo"]["Horas"].sum()
improd  = chuva_h + qbr_h + agt_h

ef_col = col_h/total_h*100           if total_h>0           else 0
ef_man = man_h/(col_h+man_h)*100     if (col_h+man_h)>0     else 0
ef_imp = improd/total_h*100          if total_h>0           else 0
ef_sa  = sa_h/total_h*100            if total_h>0           else 0
disp_m = (total_h-qbr_h)/total_h*100 if total_h>0           else 0


# ══════════════════════════════════════════════════════════════════════════════
# PÁGINA 1 — VISÃO GERAL
# ══════════════════════════════════════════════════════════════════════════════
if pagina == "Visão geral":
    st.markdown("## Visão geral operacional")
    st.caption(f"{'Todas as frotas' if frota_sel=='Todas' else frota_sel} · Safra 2026/27")

    c1,c2,c3,c4,c5 = st.columns(5)
    c1.metric("✅ Disponibilidade",  f"{disp_m:.1f}%",  f"{qbr_h:.1f}h em quebra")
    c2.metric("⚡ Efic. colheita",   f"{ef_col:.1f}%",  f"{col_h:.1f}h colhendo")
    c3.metric("⚠️ Improdutivo",       f"{ef_imp:.1f}%",  "Chuva+Quebra+Ag.Trans.")
    c4.metric("↔️ % Manobra",         f"{ef_man:.1f}%",  f"{man_h:.1f}h de manobra")
    c5.metric("❓ Sem apontamento",   f"{ef_sa:.1f}%",   f"{sa_h:.1f}h sem registro")

    st.divider()
    if not df_perf.empty:
        c6,c7,c8,c9 = st.columns(4)
        c6.metric("🛢️ Consumo médio trab.", f"{df_perf['TaxaMedTrab'].mean():.1f} L/h")
        c7.metric("🚜 Velocidade média",    f"{df_perf['VelocidadeMed'].mean():.2f} km/h")
        c8.metric("⚙️ Carga média motor",   f"{df_perf['CargaMed'].mean():.1f}%")
        c9.metric("🔴 Embuchamentos",       str(int(df_perf["Embuchamentos"].sum())))

    st.divider()
    st.markdown("**Disponibilidade mecânica por frota**")
    st.plotly_chart(fig_disponibilidade(disp_df), width="stretch",
                    config={"displayModeBar":False})
    st.divider()

    cL,cR = st.columns(2)
    with cL:
        st.markdown("**Distribuição de horas por categoria**")
        st.plotly_chart(fig_donut(df_f), width="stretch", config={"displayModeBar":False})
    with cR:
        st.markdown("**Comparativo Frota 1262 vs Frota 1263**")
        st.plotly_chart(fig_comparativo(df_apt), width="stretch", config={"displayModeBar":False})

    st.markdown("**Perdas e paradas operacionais — horas acumuladas**")
    st.plotly_chart(fig_perdas(df_f), width="stretch", config={"displayModeBar":False})


# ══════════════════════════════════════════════════════════════════════════════
# PÁGINA 2 — APONTAMENTOS
# ══════════════════════════════════════════════════════════════════════════════
elif pagina == "Apontamentos":
    st.markdown("## Apontamentos por tipo")
    st.caption(f"{'Todas as frotas' if frota_sel=='Todas' else frota_sel}")

    c1,c2,c3 = st.columns(3)
    c1.metric("✅ Disponibilidade",   f"{disp_m:.1f}%",
              f"{qbr_h:.1f}h quebra · {total_h:.1f}h total")
    c2.metric("⚡ Eficiência colheita",f"{ef_col:.1f}%", f"{col_h:.1f}h")
    c3.metric("❓ Sem apontamento",    f"{ef_sa:.1f}%",  f"{sa_h:.1f}h")

    st.divider()
    st.markdown("**Horas por apontamento**")
    st.plotly_chart(fig_barras_h(df_f), width="stretch", config={"displayModeBar":False})

    st.divider()
    cL,cR = st.columns(2)
    with cL:
        st.markdown("**Distribuição por categoria**")
        st.plotly_chart(fig_donut(df_f), width="stretch", config={"displayModeBar":False})
    with cR:
        st.markdown("**Resumo por categoria**")
        cats = df_f.groupby("Categoria")["Horas"].sum().reset_index()
        cats = cats[cats["Horas"]>0].sort_values("Horas",ascending=False)
        cats["% do Total"] = (cats["Horas"]/total_h*100).round(1).astype(str)+"%"
        cats["Horas"] = cats["Horas"].round(1)
        max_h = cats["Horas"].max()
        for _,row in cats.iterrows():
            cor = CAT_CORES.get(row["Categoria"],"#888")
            pct = row["Horas"]/max_h*100
            st.markdown(
                f'<div style="margin-bottom:10px">'
                f'<div style="display:flex;justify-content:space-between;margin-bottom:4px">'
                f'<span style="font-size:13px;color:#ccc">{row["Categoria"]}</span>'
                f'<span style="font-size:14px;font-weight:700;color:#e0e0e0">{row["Horas"]}h '
                f'<span style="font-weight:400;color:#888;font-size:12px">({row["% do Total"]})</span></span>'
                f'</div>'
                f'<div style="background:rgba(255,255,255,0.08);border-radius:6px;height:18px;overflow:hidden">'
                f'<div style="width:{pct:.1f}%;height:100%;background:{cor};border-radius:6px"></div>'
                f'</div></div>',
                unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# PÁGINA 3 — TELEMETRIA
# ══════════════════════════════════════════════════════════════════════════════
elif pagina == "Telemetria":
    st.markdown("## Telemetria — performance")
    st.caption("Motor, consumo, temperatura e velocidade")

    if df_perf.empty:
        st.warning("Sem dados de performance disponíveis.")
        st.stop()

    cols = st.columns(len(df_perf))
    for i,(_,row) in enumerate(df_perf.iterrows()):
        with cols[i]:
            st.markdown(f"**{row['NomeFrota']}**")
            d_row = disp_df[disp_df["Frota"]==row["NomeFrota"].replace("Frota ","")]
            if not d_row.empty:
                dv = d_row.iloc[0]["Disponibilidade"]
                qv = d_row.iloc[0]["Quebra"]
                st.markdown(
                    f'<div style="padding:8px 12px;border-radius:8px;'
                    f'background:rgba(39,174,96,0.15);border-left:3px solid {C_DISP};margin-bottom:8px">'
                    f'<span style="font-size:11px;color:#aaa">Disponibilidade mecânica</span><br>'
                    f'<span style="font-size:20px;font-weight:700;color:{C_DISP}">{dv:.1f}%</span>'
                    f'<span style="font-size:10px;color:#888;margin-left:8px">({qv:.1f}h em quebra)</span>'
                    f'</div>', unsafe_allow_html=True)
            for label,val in [
                ("Motor ligado",  f"{row['TempoMotorLigado']:.1f} h"),
                ("Em trabalho",   f"{row['TempoTrabalho']:.1f} h"),
                ("Motor ocioso",  f"{row['TempoOcioso']:.1f} h"),
                ("Consumo total", f"{int(row['ConsumoTotal']):,} L".replace(",",".")),
                ("Taxa em trab.", f"{row['TaxaMedTrab']:.1f} L/h"),
                ("Carga média",   f"{row['CargaMed']:.1f}%"),
                ("Vel. média",    f"{row['VelocidadeMed']:.2f} km/h"),
                ("RPM motor",     f"{int(row['RpmMotor'])} rpm"),
                ("RPM extrator",  f"{int(row['RpmExtMed'])} rpm"),
                ("Pressão méd.",  f"{row['PressaoMed']:.0f} psi"),
                ("Embuchamentos", str(int(row["Embuchamentos"]))),
            ]:
                cv = C_MANUTENCAO if label=="Embuchamentos" else "#e0e0e0"
                st.markdown(
                    f'<div style="display:flex;justify-content:space-between;padding:6px 10px;'
                    f'background:rgba(255,255,255,0.05);border-radius:6px;margin-bottom:4px">'
                    f'<span style="font-size:11px;color:#999">{label}</span>'
                    f'<span style="font-size:12px;font-weight:500;color:{cv}">{val}</span>'
                    f'</div>', unsafe_allow_html=True)

    st.divider()
    st.markdown("**Temperaturas — mín / méd / máx**")
    tcols = st.columns(len(df_perf))
    for i,(_,row) in enumerate(df_perf.iterrows()):
        with tcols[i]:
            st.markdown(f"*{row['NomeFrota']}*")
            st.plotly_chart(fig_temperaturas(row), width="stretch",
                            config={"displayModeBar":False})

    st.divider()
    st.markdown("**Gauges operacionais**")
    gcols = st.columns(len(df_perf)*3)
    idx = 0
    for _,row in df_perf.iterrows():
        ef = row["TempoTrabalho"]/row["TempoMotorLigado"]*100 if row["TempoMotorLigado"]>0 else 0
        oc = row["TempoOcioso"]/row["TempoMotorLigado"]*100   if row["TempoMotorLigado"]>0 else 0
        with gcols[idx]:
            st.plotly_chart(fig_gauge(ef, f"{row['NomeFrota']}\nEfic. motor", cor=C_PRODUTIVO),
                            width="stretch", config={"displayModeBar":False})
        with gcols[idx+1]:
            st.plotly_chart(fig_gauge(row["CargaMed"], f"{row['NomeFrota']}\nCarga motor", cor=C_ESPERA),
                            width="stretch", config={"displayModeBar":False})
        with gcols[idx+2]:
            st.plotly_chart(fig_gauge(oc, f"{row['NomeFrota']}\nOcioso", cor=C_MANUTENCAO),
                            width="stretch", config={"displayModeBar":False})
        idx += 3


# ══════════════════════════════════════════════════════════════════════════════
# PÁGINA 4 — POR FROTA
# ══════════════════════════════════════════════════════════════════════════════
elif pagina == "Por frota":
    frotas_lista = sorted(df_apt["Frota"].unique().tolist())
    frota_4  = st.selectbox("Selecione a frota", frotas_lista, key="frota4")
    sub_apt  = df_apt[df_apt["Frota"]==frota_4]
    sub_perf = df_perf[df_perf["NomeFrota"]==frota_4] if not df_perf.empty else pd.DataFrame()

    st.markdown(f"## {frota_4}")
    st.caption("Análise individual — todos os indicadores")

    tot_f  = sub_apt["Horas"].sum()
    col_f  = sub_apt[sub_apt["Apontamento"]=="Colhendo"]["Horas"].sum()
    man_f  = sub_apt[sub_apt["Apontamento"]=="Manobra"]["Horas"].sum()
    qbr_f  = sub_apt[sub_apt["Apontamento"]=="Quebra"]["Horas"].sum()
    disp_f = (tot_f-qbr_f)/tot_f*100 if tot_f>0 else 0

    k1,k2,k3,k4 = st.columns(4)
    k1.metric("✅ Disponibilidade", f"{disp_f:.1f}%",
              f"{qbr_f:.1f}h quebra · {tot_f:.1f}h total")
    k2.metric("⚡ Efic. colheita",  f"{col_f/tot_f*100:.1f}%" if tot_f>0 else "–",
              f"{col_f:.1f}h")
    k3.metric("↔️ % Manobra",       f"{man_f/(col_f+man_f)*100:.1f}%" if (col_f+man_f)>0 else "–",
              f"{man_f:.1f}h")

    if not sub_perf.empty:
        row = sub_perf.iloc[0]
        k4.metric("🔴 Embuchamentos", str(int(row["Embuchamentos"])), "detectados")
        st.divider()
        k5,k6,k7,k8 = st.columns(4)
        k5.metric("🛢️ Cons. trab.",  f"{row['TaxaMedTrab']:.1f} L/h",
                  f"{int(row['ConsumoTrabalho']):,} L".replace(",","."))
        k6.metric("🚜 Vel. média",   f"{row['VelocidadeMed']:.2f} km/h")
        k7.metric("⚙️ RPM motor",    f"{int(row['RpmMotor'])} rpm")
        oc = row["TempoOcioso"]/row["TempoMotorLigado"]*100 if row["TempoMotorLigado"]>0 else 0
        k8.metric("⏸️ Motor ocioso", f"{oc:.1f}%",
                  f"{row['TempoOcioso']:.1f}h de {row['TempoMotorLigado']:.0f}h")

    st.divider()
    cL,cR = st.columns(2)
    with cL:
        st.markdown("**Uso do tempo**")
        st.plotly_chart(fig_donut(sub_apt), width="stretch", config={"displayModeBar":False})
    with cR:
        st.markdown("**Apontamentos detalhados**")
        st.plotly_chart(fig_barras_h(sub_apt), width="stretch", config={"displayModeBar":False})

    if not sub_perf.empty:
        st.divider()
        st.markdown("**Temperaturas**")
        st.plotly_chart(fig_temperaturas(sub_perf.iloc[0]),
                        width="stretch", config={"displayModeBar":False})