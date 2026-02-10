import pandas as pd
import streamlit as st

st.set_page_config(page_title="SupplyChain AI Starter Kit", layout="wide")
st.title("SupplyChain AI Starter Kit — PMI")

REQUIRED_COLS = [
    "articolo",
    "consumo_mensile",
    "lead_time_giorni",
    "stock_attuale",
    "criticita",
    "valore_unitario",
    "unita_misura",
]


def load_data(file):
    name = file.name.lower()

    if name.endswith(".csv"):
        # Prova a leggere con separatore automatico (gestisce ; e ,)
        df = pd.read_csv(file, sep=None, engine="python")
    else:
        df = pd.read_excel(file)

    # pulizia nomi colonne
    df.columns = [str(c).strip().lower() for c in df.columns]

    # rimuove eventuali colonne vuote tipo "unnamed: 0"
    df = df.loc[:, ~df.columns.str.contains("^unnamed")]

    return df

def validate_columns(df):
    return [c for c in REQUIRED_COLS if c not in df.columns]

def compute_metrics(df):
    out = df.copy()

    # consumo giornaliero (approssimazione su 30 giorni)
    out["consumo_giornaliero"] = out["consumo_mensile"] / 30.0

    # domanda durante lead time
    out["domanda_lt"] = out["consumo_giornaliero"] * out["lead_time_giorni"]

    # scorta di sicurezza "semplice" in base alla criticità
    def safety_factor(c):
        c = str(c).strip().lower()
        if c == "alta":
            return 0.50
        if c == "media":
            return 0.30
        return 0.15

    out["fattore_ss"] = out["criticita"].apply(safety_factor)
    out["scorta_sicurezza"] = out["domanda_lt"] * out["fattore_ss"]

    # punto riordino
    out["punto_riordino"] = (out["domanda_lt"] + out["scorta_sicurezza"]).round(0)

    # rischio stockout (semplice)
    def risk(row):
        if row["stock_attuale"] < row["domanda_lt"]:
            return "alto"
        if row["stock_attuale"] < row["punto_riordino"]:
            return "medio"
        return "basso"

    out["rischio_stockout"] = out.apply(risk, axis=1)

    # quantità suggerita per tornare al punto di riordino (approccio prudente)
    out["qty_suggerita"] = (out["punto_riordino"] - out["stock_attuale"]).clip(lower=0).round(0)

    # ordinamento per priorità (alto->medio->basso)
    priority_map = {"alto": 0, "medio": 1, "basso": 2}
    out["priorita_sort"] = out["rischio_stockout"].map(priority_map).fillna(3)
    out = out.sort_values(["priorita_sort", "valore_unitario"], ascending=[True, False]).drop(columns=["priorita_sort"])

    import math

# ... dentro compute_metrics

out["punto_riordino"] = (out["domanda_lt"] + out["scorta_sicurezza"]).apply(lambda x: math.ceil(x))
out["qty_suggerita"] = (out["punto_riordino"] - out["stock_attuale"]).clip(lower=0).apply(lambda x: math.ceil(x))

# prezzi a 2 decimali (solo formato numerico)
out["valore_unitario"] = out["valore_unitario"].round(2)

    return out

# ✏️ MODIFICA 1 — Prompt generator
def genera_prompt(row):
    return f"""
Agisci come responsabile supply chain di una PMI.

Articolo: {row['articolo']}
Consumo mensile: {row['consumo_mensile']}
Lead time (giorni): {row['lead_time_giorni']}
Stock attuale: {row['stock_attuale']}
Criticità: {row['criticita']}
Valore unitario: {row['valore_unitario']}

Analizza la situazione e indica:
- rischio stockout
- se riordinare o no
- quantità consigliata
- motivazione pratica e semplice
"""

# Sidebar upload
st.sidebar.header("Carica file dati")
uploaded = st.sidebar.file_uploader("CSV o Excel", type=["csv", "xlsx", "xls"])

# Mostra esempio (utile per PMI e per evitare errori)
st.subheader("Esempio di dati (formato richiesto)")
example_df = pd.DataFrame(
    [
        {"articolo": "A001", "consumo_mensile": 100, "lead_time_giorni": 10, "stock_attuale": 20, "criticita": "alta", "valore_unitario": 10.5},
        {"articolo": "A002", "consumo_mensile": 50,  "lead_time_giorni": 20, "stock_attuale": 40, "criticita": "media","valore_unitario": 5.0},
        {"articolo": "A003", "consumo_mensile": 200, "lead_time_giorni": 5,  "stock_attuale": 10, "criticita": "bassa","valore_unitario": 15.0},
    ]
)
st.dataframe(example_df, use_container_width=True)

st.info("Carica un file con colonne: " + ", ".join(REQUIRED_COLS))

if not uploaded:
    st.stop()

df = load_data(uploaded)
missing = validate_columns(df)

if missing:
    st.error(f"Colonne mancanti: {missing}")
    st.stop()

metrics = compute_metrics(df)

st.subheader("Priorità riordino e rischio stockout")
st.dataframe(
    metrics[
        ["articolo", "stock_attuale", "domanda_lt", "scorta_sicurezza", "punto_riordino", "rischio_stockout", "qty_suggerita"]
    ],
    use_container_width=True
)

# ✏️ MODIFICA 2 — UI prompt generator
st.subheader("Prompt AI per analisi decisionale")

articolo_sel = st.selectbox(
    "Seleziona articolo",
    metrics["articolo"].astype(str).tolist()
)

row = metrics[metrics["articolo"].astype(str) == articolo_sel].iloc[0]
prompt = genera_prompt(row)

st.text_area(
    "Prompt pronto (copia e incolla in ChatGPT)",
    prompt,
    height=220
)

st.download_button(
    "Scarica risultati (CSV)",
    metrics.to_csv(index=False).encode("utf-8"),
    file_name="supplychain_ai_results.csv",
    mime="text/csv"
)
