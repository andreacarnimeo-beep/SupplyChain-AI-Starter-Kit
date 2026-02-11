import math
import datetime as dt
from pathlib import Path

import pandas as pd
import streamlit as st
from PIL import Image


# =========================
# CONFIG
# =========================
st.set_page_config(page_title="SupplyChain AI Starter Kit", layout="wide")

st.markdown(
    """
<style>
.block-container {padding-top: 1.6rem;}
.small-note {color:#6b7280; font-size:0.9rem;}
</style>
""",
    unsafe_allow_html=True,
)

REQUIRED_COLS = [
    "articolo",
    "consumo_mensile",
    "lead_time_giorni",
    "stock_attuale",
    "criticita",
    "valore_unitario",
]


# =========================
# FUNZIONI
# =========================
def business_days_in_month(year: int, month: int) -> int:
    start = dt.date(year, month, 1)
    end = dt.date(year + 1, 1, 1) if month == 12 else dt.date(year, month + 1, 1)
    cur, count = start, 0
    while cur < end:
        if cur.weekday() < 5:
            count += 1
        cur += dt.timedelta(days=1)
    return count


def load_data(file) -> pd.DataFrame:
    if file.name.lower().endswith(".csv"):
        df = pd.read_csv(file, sep=None, engine="python")
    else:
        df = pd.read_excel(file)
    df.columns = [str(c).strip().lower() for c in df.columns]
    df = df.loc[:, ~df.columns.str.contains("^unnamed", case=False, na=False)]
    return df


def validate_columns(df: pd.DataFrame):
    return [c for c in REQUIRED_COLS if c not in df.columns]


def normalize_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["criticita"] = df["criticita"].astype(str).str.lower().str.strip()
    df["criticita"] = df["criticita"].replace({"alto": "alta", "medio": "media", "basso": "bassa"})

    for c in ["consumo_mensile", "lead_time_giorni", "stock_attuale", "valore_unitario"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    df = df.dropna(subset=REQUIRED_COLS)
    return df


def compute_metrics(df: pd.DataFrame, workdays: int) -> pd.DataFrame:
    out = df.copy()

    out["consumo_giornaliero"] = out["consumo_mensile"] / float(workdays)
    out["domanda_lt"] = out["consumo_giornaliero"] * out["lead_time_giorni"]

    def safety_factor(c):
        if c == "alta":
            return 0.50
        if c == "media":
            return 0.30
        return 0.15

    out["scorta_sicurezza"] = out["domanda_lt"] * out["criticita"].apply(safety_factor)

    out["punto_riordino"] = (out["domanda_lt"] + out["scorta_sicurezza"]).apply(lambda x: int(math.ceil(x)))
    out["qty_suggerita"] = (out["punto_riordino"] - out["stock_attuale"]).clip(lower=0).apply(lambda x: int(math.ceil(x)))

    def risk(r):
        if r["stock_attuale"] < r["domanda_lt"]:
            return "alto"
        if r["stock_attuale"] < r["punto_riordino"]:
            return "medio"
        return "basso"

    out["rischio_stockout"] = out.apply(risk, axis=1)
    out["valore_unitario"] = out["valore_unitario"].round(2)
    return out


def genera_prompt(row: pd.Series, year: int, month: int, workdays: int) -> str:
    return f"""
Agisci come responsabile supply chain di una PMI.

Mese: {month:02d}/{year} (giorni lavorativi lun–ven: {workdays})

Articolo: {row['articolo']}
Consumo mensile: {int(row['consumo_mensile'])}
Lead time (giorni): {int(row['lead_time_giorni'])}
Stock attuale: {int(row['stock_attuale'])}
Criticità: {row['criticita']}
Valore unitario (€): {row['valore_unitario']:.2f}

Calcoli:
- Domanda su lead time: {row['domanda_lt']:.2f}
- Scorta sicurezza: {row['scorta_sicurezza']:.2f}
- Punto riordino: {int(row['punto_riordino'])}
- Qty suggerita: {int(row['qty_suggerita'])}
- Rischio stockout: {row['rischio_stockout']}

Spiega se riordinare o no, perché, e suggerisci 2 azioni operative immediate.
""".strip()


# =========================
# HERO (SENZA SPAZI VUOTI)
# =========================
left, right = st.columns([1.35, 1])

with left:
    with st.container(border=True):
        st.markdown("### SupplyChain AI Starter Kit")
        st.markdown(
            "Strumento decisionale per Responsabili Logistica e Supply Chain. "
            "Carica i dati → ottieni priorità di riordino → usa l’AI per decidere meglio."
        )
        st.markdown(
            """
- Punto di riordino e rischio stockout  
- Quantità suggerite arrotondate  
- Prompt AI pronti all’uso  
"""
        )
        st.markdown('<div class="small-note">Nota: consumo_mensile = quantità mensili.</div>', unsafe_allow_html=True)

with right:
    with st.container(border=True):
        img_path = Path("assets/logistic_manager_future.png")
        if img_path.exists():
            img = Image.open(img_path)
            st.image(img, use_container_width=True)
        else:
            st.warning("Immagine non trovata: assets/logistic_manager_future.png")


st.divider()


# =========================
# SIDEBAR
# =========================
st.sidebar.header("Impostazioni")
today = dt.date.today()
year = int(st.sidebar.number_input("Anno", 2020, 2100, today.year))
month = int(st.sidebar.selectbox("Mese", list(range(1, 13)), index=today.month - 1))
workdays = business_days_in_month(year, month)
st.sidebar.caption(f"Giorni lavorativi (lun–ven): {workdays}")

uploaded = st.sidebar.file_uploader("Carica CSV o Excel", type=["csv", "xlsx", "xls"])
if not uploaded:
    st.info("Carica un file per vedere l’analisi.")
    st.stop()


# =========================
# PROCESSING
# =========================
df_raw = load_data(uploaded)
missing = validate_columns(df_raw)
if missing:
    st.error(f"Colonne mancanti: {missing}")
    st.stop()

df = normalize_df(df_raw)
if df.empty:
    st.error("Nessuna riga valida trovata. Controlla che i campi numerici siano corretti.")
    st.stop()

metrics = compute_metrics(df, workdays)


# =========================
# OUTPUT
# =========================
st.subheader("Priorità riordino e rischio stockout")
st.dataframe(metrics, use_container_width=True)

st.subheader("Prompt AI decisionale")
art = st.selectbox("Seleziona articolo", metrics["articolo"].astype(str).tolist())
row = metrics[metrics["articolo"].astype(str) == art].iloc[0]
st.text_area("Prompt pronto (copia e incolla in ChatGPT)", genera_prompt(row, year, month, workdays), height=260)

