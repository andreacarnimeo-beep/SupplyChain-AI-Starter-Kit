import math
import datetime as dt
from io import BytesIO

import pandas as pd
import streamlit as st
from PIL import Image

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter


# =========================
# CONFIG + STYLE
# =========================
st.set_page_config(page_title="SupplyChain AI Starter Kit", layout="wide")

st.markdown(
    """
<style>
.block-container {padding-top: 2rem;}
.sc-card {
  background: white;
  border: 1px solid rgba(0,0,0,0.06);
  border-radius: 16px;
  padding: 18px;
  box-shadow: 0 8px 24px rgba(0,0,0,0.06);
}
.badge {
  display:inline-block; padding: 4px 12px; border-radius: 999px;
  background: rgba(30,136,229,0.12); color: #1e88e5;
  font-size: 0.85rem; border: 1px solid rgba(30,136,229,0.2);
}
.small {color: rgba(0,0,0,0.6); font-size: 0.9rem;}
</style>
""",
    unsafe_allow_html=True,
)


# =========================
# COSTANTI
# =========================
REQUIRED_COLS = [
    "articolo",
    "consumo_mensile",
    "lead_time_giorni",
    "stock_attuale",
    "criticita",
    "valore_unitario",
]
OPTIONAL_DEFAULTS = {"unita_misura": "pz"}


# =========================
# FUNZIONI BASE
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
    df = df.loc[:, ~df.columns.str.contains("^unnamed", case=False)]
    return df


def validate_columns(df):
    return [c for c in REQUIRED_COLS if c not in df.columns]


def normalize_df(df):
    out = df.copy()
    for c, v in OPTIONAL_DEFAULTS.items():
        if c not in out.columns:
            out[c] = v

    out["criticita"] = out["criticita"].astype(str).str.lower().str.strip()
    out["criticita"] = out["criticita"].replace(
        {"alto": "alta", "medio": "media", "basso": "bassa"}
    )

    for c in ["consumo_mensile", "lead_time_giorni", "stock_attuale", "valore_unitario"]:
        out[c] = pd.to_numeric(out[c], errors="coerce")

    out = out.dropna(subset=REQUIRED_COLS)
    out["unita_misura"] = out["unita_misura"].astype(str).str.strip()
    out.loc[out["unita_misura"] == "", "unita_misura"] = "pz"
    return out


def compute_metrics(df, workdays):
    out = df.copy()
    out["consumo_giornaliero"] = out["consumo_mensile"] / workdays
    out["domanda_lt"] = out["consumo_giornaliero"] * out["lead_time_giorni"]

    def sf(c):
        return 0.5 if c == "alta" else 0.3 if c == "media" else 0.15

    out["scorta_sicurezza"] = out["domanda_lt"] * out["criticita"].apply(sf)
    out["punto_riordino"] = (out["domanda_lt"] + out["scorta_sicurezza"]).apply(lambda x: math.ceil(x))
    out["qty_suggerita"] = (out["punto_riordino"] - out["stock_attuale"]).clip(lower=0).apply(lambda x: math.ceil(x))

    def risk(r):
        if r["stock_attuale"] < r["domanda_lt"]:
            return "alto"
        if r["stock_attuale"] < r["punto_riordino"]:
            return "medio"
        return "basso"

    out["rischio_stockout"] = out.apply(risk, axis=1)
    out["valore_unitario"] = out["valore_unitario"].round(2)
    return out.sort_values(["rischio_stockout", "valore_unitario"], ascending=[True, False])


def genera_prompt(row, year, month, workdays):
    um = row["unita_misura"]
    return f"""
Agisci come responsabile supply chain di una PMI.

Mese: {month:02d}/{year} (giorni lavorativi: {workdays})
Articolo: {row['articolo']} ({um})

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


# =========================
# HERO SECTION (TESTO + IMMAGINE)
# =========================
col1, col2 = st.columns([1.2, 1])

with col1:
    st.markdown(
        """
<div class="sc-card">
  <span class="badge">Decision Support • AI Ready</span>
  <h1>SupplyChain AI Starter Kit</h1>
  <p class="small">
    Strumento decisionale per Responsabili Logistica e Supply Chain.
    Carica i dati → ottieni priorità di riordino → esporta Excel → usa l’AI per decidere meglio.
  </p>
  <ul>
    <li>Punto di riordino e rischio stockout</li>
    <li>Quantità suggerite arrotondate</li>
    <li>Excel professionale (Input + Output)</li>
    <li>Prompt AI pronti all’uso</li>
  </ul>
</div>
""",
        unsafe_allow_html=True,
    )

with col2:
    img = Image.open("assets/logistic_manager_future.png")
    st.markdown('<div class="sc-card">', unsafe_allow_html=True)
    st.image(img, use_container_width=True, caption="Il Logistic Manager del Futuro")
    st.markdown('<p class="small">Immagine originale — uso libero</p>', unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)


# =========================
# TEMPLATE DOWNLOAD
# =========================
def build_template_xlsx():
    wb = Workbook()
    ws = wb.active
    ws.title = "Dati"
    headers = ["articolo","consumo_mensile","lead_time_giorni","stock_attuale","criticita","valore_unitario","unita_misura"]
    ws.append(headers)
    ws.append(["A001",100,10,20,"alta",10.50,"pz"])

    header_fill = PatternFill("solid", fgColor="D9EAF7")
    for c in range(1, len(headers)+1):
        ws.cell(1,c).font = Font(bold=True)
        ws.cell(1,c).fill = header_fill
        ws.cell(1,c).alignment = Alignment(horizontal="center")

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"

    dv = DataValidation(type="list", formula1='"bassa,media,alta"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add("E2:E200")

    for col in ["B","C","D"]:
        for r in range(2,201):
            ws[f"{col}{r}"].number_format = "0"
    for r in range(2,201):
        ws[f"F{r}"].number_format = "0.00"

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


st.download_button(
    "⬇️ Scarica template Excel",
    data=build_template_xlsx(),
    file_name="SupplyChain_AI_Template.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)


# =========================
# SIDEBAR + UPLOAD
# =========================
st.sidebar.header("Impostazioni")
today = dt.date.today()
year = int(st.sidebar.number_input("Anno", 2020, 2100, today.year))
month = int(st.sidebar.selectbox("Mese", list(range(1,13)), index=today.month-1))
workdays = business_days_in_month(year, month)
st.sidebar.caption(f"Giorni lavorativi: {workdays}")

uploaded = st.sidebar.file_uploader("Carica CSV o Excel", type=["csv","xlsx","xls"])
if not uploaded:
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
metrics = compute_metrics(df, workdays)

st.subheader("Priorità riordino e rischio stockout")
st.dataframe(metrics, use_container_width=True)


# =========================
# EXCEL OUTPUT
# =========================
def build_results_xlsx(df, metrics):
    wb = Workbook()

    for name, data in {"Input": df, "Output": metrics}.items():
        ws = wb.create_sheet(name)
        ws.append(list(data.columns))
        for c in range(1, len(data.columns)+1):
            ws.cell(1,c).font = Font(bold=True)
            ws.cell(1,c).fill = PatternFill("solid", fgColor="D9EAF7")
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = f"A1:{get_column_letter(len(data.columns))}1"
        for _, r in data.iterrows():
            ws.append(r.tolist())

    wb.remove(wb["Sheet"])
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


st.download_button(
    "⬇️ Scarica risultati Excel (Input + Output)",
    data=build_results_xlsx(df, metrics),
    file_name="SupplyChain_AI_Results.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)


# =========================
# PROMPT AI
# =========================
st.subheader("Prompt AI decisionale")
art = st.selectbox("Seleziona articolo", metrics["articolo"].astype(str))
row = metrics[metrics["articolo"].astype(str)==art].iloc[0]
st.text_area("Prompt pronto", genera_prompt(row, year, month, workdays), height=260)

