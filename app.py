import math
import datetime as dt
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st
from PIL import Image

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter


# =====================================================
# CONFIGURAZIONE PAGINA
# =====================================================
st.set_page_config(
    page_title="SupplyChain AI Starter Kit",
    layout="wide"
)


# =====================================================
# STILE SOBRIO / CORPORATE
# =====================================================
st.markdown("""
<style>
.block-container {padding-top: 2rem;}

.hero-box {
    background: white;
    border: 1px solid #e6e6e6;
    border-radius: 10px;
    padding: 22px;
    height: 420px;
}

.hero-img-box {
    background: white;
    border: 1px solid #e6e6e6;
    border-radius: 10px;
    padding: 10px;
    height: 420px;
    display: flex;
    align-items: center;
    justify-content: center;
}

.subtitle {
    color: #555;
    font-size: 0.95rem;
}

.small-note {
    font-size: 0.85rem;
    color: #777;
}
</style>
""", unsafe_allow_html=True)


# =====================================================
# COSTANTI
# =====================================================
REQUIRED_COLS = [
    "articolo",
    "consumo_mensile",
    "lead_time_giorni",
    "stock_attuale",
    "criticita",
    "valore_unitario",
]


# =====================================================
# FUNZIONI UTILI
# =====================================================
def business_days_in_month(year: int, month: int) -> int:
    start = dt.date(year, month, 1)
    end = dt.date(year + 1, 1, 1) if month == 12 else dt.date(year, month + 1, 1)

    count = 0
    cur = start
    while cur < end:
        if cur.weekday() < 5:
            count += 1
        cur += dt.timedelta(days=1)

    return count


def load_data(file):
    if file.name.lower().endswith(".csv"):
        df = pd.read_csv(file, sep=None, engine="python")
    else:
        df = pd.read_excel(file)

    df.columns = [str(c).strip().lower() for c in df.columns]
    df = df.loc[:, ~df.columns.str.contains("^unnamed", case=False)]
    return df


def validate_columns(df):
    return [c for c in REQUIRED_COLS if c not in df.columns]


def normalize_df(df):
    df = df.copy()
    df["criticita"] = df["criticita"].astype(str).str.lower().str.strip()
    df["criticita"] = df["criticita"].replace(
        {"alto": "alta", "medio": "media", "basso": "bassa"}
    )

    for c in ["consumo_mensile", "lead_time_giorni", "stock_attuale", "valore_unitario"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    df = df.dropna(subset=REQUIRED_COLS)
    return df


def compute_metrics(df, workdays):
    out = df.copy()

    out["consumo_giornaliero"] = out["consumo_mensile"] / workdays
    out["domanda_lt"] = out["consumo_giornaliero"] * out["lead_time_giorni"]

    def safety_factor(c):
        if c == "alta":
            return 0.5
        if c == "media":
            return 0.3
        return 0.15

    out["scorta_sicurezza"] = out["domanda_lt"] * out["criticita"].apply(safety_factor)

    out["punto_riordino"] = (
        out["domanda_lt"] + out["scorta_sicurezza"]
    ).apply(lambda x: math.ceil(x))

    out["qty_suggerita"] = (
        out["punto_riordino"] - out["stock_attuale"]
    ).clip(lower=0).apply(lambda x: math.ceil(x))

    def risk(row):
        if row["stock_attuale"] < row["domanda_lt"]:
            return "alto"
        if row["stock_attuale"] < row["punto_riordino"]:
            return "medio"
        return "basso"

    out["rischio_stockout"] = out.apply(risk, axis=1)
    out["valore_unitario"] = out["valore_unitario"].round(2)

    return out


def genera_prompt(row, year, month, workdays):
    return f"""
Agisci come responsabile supply chain di una PMI.

Mese: {month:02d}/{year} (giorni lavorativi: {workdays})

Articolo: {row['articolo']}
Consumo mensile: {int(row['consumo_mensile'])}
Stock attuale: {int(row['stock_attuale'])}
Lead time (giorni): {int(row['lead_time_giorni'])}
Criticità: {row['criticita']}
Valore unitario (€): {row['valore_unitario']:.2f}

Punto di riordino: {int(row['punto_riordino'])}
Quantità suggerita: {int(row['qty_suggerita'])}
Rischio stockout: {row['rischio_stockout']}

Spiega se riordinare o no, perché, e suggerisci 2 azioni operative.
""".strip()


# =====================================================
# HERO SECTION (TESTO + IMMAGINE)
# =====================================================
col1, col2 = st.columns([1.3, 1])

with col1:
    st.markdown('<div class="hero-box">', unsafe_allow_html=True)
    st.title("SupplyChain AI Starter Kit")
    st.markdown(
        '<p class="subtitle">Strumento decisionale per Responsabili Logistica e Supply Chain.</p>',
        unsafe_allow_html=True,
    )
    st.markdown("""
    • Punto di riordino e rischio stockout  
    • Quantità suggerite arrotondate  
    • Export Excel professionale  
    • Prompt AI pronti all’uso
    """)
    st.markdown("</div>", unsafe_allow_html=True)

with col2:
    st.markdown('<div class="hero-img-box">', unsafe_allow_html=True)
    img_path = Path("assets/logistic_manager_future.png")
    if img_path.exists():
        img = Image.open(img_path)
        st.image(img, width=380)
    else:
        st.warning("Immagine non trovata nella cartella assets.")
    st.markdown("</div>", unsafe_allow_html=True)


# =====================================================
# SIDEBAR
# =====================================================
st.sidebar.header("Impostazioni")

today = dt.date.today()
year = int(st.sidebar.number_input("Anno", 2020, 2100, today.year))
month = int(st.sidebar.selectbox("Mese", list(range(1, 13)), index=today.month - 1))

workdays = business_days_in_month(year, month)
st.sidebar.caption(f"Giorni lavorativi (lun-ven): {workdays}")

uploaded = st.sidebar.file_uploader("Carica CSV o Excel", type=["csv", "xlsx", "xls"])
if not uploaded:
    st.stop()


# =====================================================
# PROCESSING
# =====================================================
df_raw = load_data(uploaded)
missing = validate_columns(df_raw)

if missing:
    st.error(f"Colonne mancanti: {missing}")
    st.stop()

df = normalize_df(df_raw)
metrics = compute_metrics(df, workdays)

st.subheader("Priorità riordino e rischio stockout")
st.dataframe(metrics, use_container_width=True)


# =====================================================
# PROMPT AI
# =====================================================
st.subheader("Prompt AI decisionale")

art = st.selectbox("Seleziona articolo", metrics["articolo"].astype(str))
row = metrics[metrics["articolo"].astype(str) == art].iloc[0]

st.text_area(
    "Prompt pronto",
    genera_prompt(row, year, month, workdays),
    height=260
)
