from openpyxl.utils import get_column_letter


import math
from io import BytesIO

import pandas as pd
import streamlit as st

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.worksheet.datavalidation import DataValidation


st.set_page_config(page_title="SupplyChain AI Starter Kit", layout="wide")
st.title("SupplyChain AI Starter Kit — PMI")

# Colonne minime richieste (unita_misura verrà aggiunta se manca)
REQUIRED_COLS = [
    "articolo",
    "consumo_mensile",     # quantità MENSILI
    "lead_time_giorni",
    "stock_attuale",
    "criticita",
    "valore_unitario",     # € con 2 decimali
]

OPTIONAL_COLS_DEFAULTS = {
    "unita_misura": "pz"
}


def build_template_xlsx() -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Dati"

    headers = [
        "articolo",
        "consumo_mensile",     # quantità MENSILI
        "lead_time_giorni",
        "stock_attuale",
        "criticita",           # bassa/media/alta
        "valore_unitario",     # € con 2 decimali
        "unita_misura",        # es. pz, kg, lt
    ]
    ws.append(headers)

    # Header in grassetto + centrato
    header_font = Font(bold=True)
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # larghezze
    widths = {"A": 14, "B": 18, "C": 16, "D": 14, "E": 12, "F": 16, "G": 12}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    # riga esempio (utile al cliente)
    ws.append(["A001", 100, 10, 20, "alta", 10.50, "pz"])

    # formati numerici fino a 200 righe
    for r in range(2, 201):
        ws[f"B{r}"].number_format = "0"     # consumo mensile intero
        ws[f"C{r}"].number_format = "0"     # lead time intero
        ws[f"D{r}"].number_format = "0"     # stock intero
        ws[f"F{r}"].number_format = "0.00"  # prezzo 2 decimali

    # menu a tendina criticità
    dv = DataValidation(type="list", formula1='"bassa,media,alta"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add("E2:E200")

    # nota
    ws["I1"] = "NOTE"
    ws["I2"] = "consumo_mensile = quantità mensili (unità in colonna G)."

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()

def build_results_xlsx(input_df: pd.DataFrame, output_df: pd.DataFrame) -> bytes:
    wb = Workbook()

    # ===== Sheet 1: Input =====
    ws_in = wb.active
    ws_in.title = "Input"

    input_cols = ["articolo", "unita_misura", "consumo_mensile", "lead_time_giorni", "stock_attuale", "criticita", "valore_unitario"]
    for c in input_cols:
        if c not in input_df.columns:
            input_df[c] = "pz" if c == "unita_misura" else ""

    df_in = input_df[input_cols].copy()

    ws_in.append(input_cols)
    header_font = Font(bold=True)

    for col_idx, col_name in enumerate(input_cols, start=1):
        cell = ws_in.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Data rows
    for _, r in df_in.iterrows():
        ws_in.append([
            r["articolo"],
            r["unita_misura"],
            int(r["consumo_mensile"]) if pd.notna(r["consumo_mensile"]) else None,
            int(r["lead_time_giorni"]) if pd.notna(r["lead_time_giorni"]) else None,
            int(r["stock_attuale"]) if pd.notna(r["stock_attuale"]) else None,
            str(r["criticita"]),
            float(r["valore_unitario"]) if pd.notna(r["valore_unitario"]) else None,
        ])

    # Formati Input
    for row in range(2, ws_in.max_row + 1):
        ws_in[f"C{row}"].number_format = "0"     # consumo mensile intero
        ws_in[f"D{row}"].number_format = "0"     # lead time intero
        ws_in[f"E{row}"].number_format = "0"     # stock intero
        ws_in[f"G{row}"].number_format = "0.00"  # prezzo 2 decimali

    # Menu a tendina criticità (colonna F)
    dv = DataValidation(type="list", formula1='"bassa,media,alta"', allow_blank=False)
    ws_in.add_data_validation(dv)
    dv.add(f"F2:F{max(2, ws_in.max_row)}")

    # Larghezze colonne Input
    widths_in = [14, 12, 18, 16, 14, 12, 16]
    for i, w in enumerate(widths_in, start=1):
        ws_in.column_dimensions[get_column_letter(i)].width = w

    ws_in["I1"] = "NOTE"
    ws_in["I2"] = "consumo_mensile = quantità mensili (unità in colonna B)."

    # ===== Sheet 2: Output =====
    ws_out = wb.create_sheet("Output")

    output_cols = [
        "articolo", "unita_misura", "consumo_mensile", "stock_attuale",
        "domanda_lt", "scorta_sicurezza", "punto_riordino",
        "rischio_stockout", "qty_suggerita", "valore_unitario"
    ]
    for c in output_cols:
        if c not in output_df.columns:
            output_df[c] = ""

    df_out = output_df[output_cols].copy()

    ws_out.append(output_cols)
    for col_idx, col_name in enumerate(output_cols, start=1):
        cell = ws_out.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for _, r in df_out.iterrows():
        ws_out.append([
            r["articolo"],
            r["unita_misura"],
            int(r["consumo_mensile"]) if pd.notna(r["consumo_mensile"]) else None,
            int(r["stock_attuale"]) if pd.notna(r["stock_attuale"]) else None,
            float(r["domanda_lt"]) if pd.notna(r["domanda_lt"]) else None,
            float(r["scorta_sicurezza"]) if pd.notna(r["scorta_sicurezza"]) else None,
            int(r["punto_riordino"]) if pd.notna(r["punto_riordino"]) else None,
            str(r["rischio_stockout"]),
            int(r["qty_suggerita"]) if pd.notna(r["qty_suggerita"]) else None,
            float(r["valore_unitario"]) if pd.notna(r["valore_unitario"]) else None,
        ])

    # Formati Output
    for row in range(2, ws_out.max_row + 1):
        ws_out[f"C{row}"].number_format = "0"     # consumo mensile
        ws_out[f"D{row}"].number_format = "0"     # stock
        ws_out[f"E{row}"].number_format = "0.00"  # domanda_lt
        ws_out[f"F{row}"].number_format = "0.00"  # scorta_sicurezza
        ws_out[f"G{row}"].number_format = "0"     # punto riordino intero
        ws_out[f"I{row}"].number_format = "0"     # qty suggerita intero
        ws_out[f"J{row}"].number_format = "0.00"  # prezzo 2 decimali

    # Larghezze colonne Output
    widths_out = [14, 12, 18, 14, 14, 16, 16, 16, 14, 16]
    for i, w in enumerate(widths_out, start=1):
        ws_out.column_dimensions[get_column_letter(i)].width = w

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def load_data(file) -> pd.DataFrame:
    name = file.name.lower()

    if name.endswith(".csv"):
        # Gestisce separatori ; e , automaticamente
        df = pd.read_csv(file, sep=None, engine="python")
    else:
        df = pd.read_excel(file)

    # pulizia colonne
    df.columns = [str(c).strip().lower() for c in df.columns]
    # rimuove eventuali colonne tipo "Unnamed: 0"
    df = df.loc[:, ~df.columns.str.contains(r"^unnamed", case=False, na=False)]

    return df


def validate_columns(df: pd.DataFrame):
    return [c for c in REQUIRED_COLS if c not in df.columns]


def normalize_df(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()

    # aggiunge colonne opzionali se mancano
    for col, default in OPTIONAL_COLS_DEFAULTS.items():
        if col not in out.columns:
            out[col] = default

    # pulizia/normalizzazione criticità
    out["criticita"] = out["criticita"].astype(str).str.strip().str.lower()

    # numeric coercion (evita crash se arriva testo)
    for col in ["consumo_mensile", "lead_time_giorni", "stock_attuale", "valore_unitario"]:
        out[col] = pd.to_numeric(out[col], errors="coerce")

    # drop righe con valori fondamentali mancanti
    out = out.dropna(subset=["articolo", "consumo_mensile", "lead_time_giorni", "stock_attuale", "criticita", "valore_unitario"])

    return out


def compute_metrics(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()

    # consumo giornaliero stimato (mese = 30 giorni)
    out["consumo_giornaliero"] = out["consumo_mensile"] / 30.0
    out["domanda_lt"] = out["consumo_giornaliero"] * out["lead_time_giorni"]

    def safety_factor(c):
        c = str(c).strip().lower()
        if c == "alta":
            return 0.50
        if c == "media":
            return 0.30
        return 0.15

    out["fattore_ss"] = out["criticita"].apply(safety_factor)
    out["scorta_sicurezza"] = out["domanda_lt"] * out["fattore_ss"]

    # arrotondamenti: all’intero superiore
    out["punto_riordino"] = (out["domanda_lt"] + out["scorta_sicurezza"]).apply(lambda x: int(math.ceil(x)))
    out["qty_suggerita"] = (out["punto_riordino"] - out["stock_attuale"]).clip(lower=0).apply(lambda x: int(math.ceil(x)))

    def risk(row):
        if row["stock_attuale"] < row["domanda_lt"]:
            return "alto"
        if row["stock_attuale"] < row["punto_riordino"]:
            return "medio"
        return "basso"

    out["rischio_stockout"] = out.apply(risk, axis=1)

    # prezzi a 2 decimali
    out["valore_unitario"] = out["valore_unitario"].round(2)

    # ordinamento priorità (alto -> medio -> basso), poi per valore unitario alto
    priority_map = {"alto": 0, "medio": 1, "basso": 2}
    out["priorita_sort"] = out["rischio_stockout"].map(priority_map).fillna(3)
    out = out.sort_values(["priorita_sort", "valore_unitario"], ascending=[True, False]).drop(columns=["priorita_sort"])

    return out


def genera_prompt(row) -> str:
    return f"""
Agisci come responsabile supply chain di una PMI.

Articolo: {row['articolo']}
Consumo mensile ({row.get('unita_misura','pz')}): {int(row['consumo_mensile'])}
Lead time (giorni): {int(row['lead_time_giorni'])}
Stock attuale ({row.get('unita_misura','pz')}): {int(row['stock_attuale'])}
Criticità: {row['criticita']}
Valore unitario (€): {row['valore_unitario']:.2f}

Calcoli app:
- Punto riordino ({row.get('unita_misura','pz')}): {int(row['punto_riordino'])}
- Qty suggerita ({row.get('unita_misura','pz')}): {int(row['qty_suggerita'])}
- Rischio stockout: {row['rischio_stockout']}

Analizza e indica in modo pratico:
- se riordinare o no
- quantità consigliata e perché
- azioni immediate (es. accelerare consegna, alternativa fornitore, safety stock)
"""


# ===== UI =====

st.download_button(
    "Scarica template Excel (pronto da compilare)",
    data=build_template_xlsx(),
    file_name="SupplyChain_AI_Starter_Kit_Template.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.caption("Nel template: intestazioni in grassetto, criticità con menu a tendina, prezzi a 2 decimali, consumo_mensile = quantità mensili.")

st.sidebar.header("Carica file dati")
uploaded = st.sidebar.file_uploader("CSV o Excel", type=["csv", "xlsx", "xls"])

st.subheader("Esempio di dati (formato richiesto)")
example_df = pd.DataFrame(
    [
        {"articolo": "A001", "consumo_mensile": 100, "lead_time_giorni": 10, "stock_attuale": 20, "criticita": "alta", "valore_unitario": 10.50, "unita_misura": "pz"},
        {"articolo": "A002", "consumo_mensile": 50,  "lead_time_giorni": 20, "stock_attuale": 40, "criticita": "media", "valore_unitario": 5.00,  "unita_misura": "pz"},
        {"articolo": "A003", "consumo_mensile": 200, "lead_time_giorni": 5,  "stock_attuale": 10, "criticita": "bassa", "valore_unitario": 15.00, "unita_misura": "pz"},
    ]
)
st.dataframe(example_df, use_container_width=True)

st.info("Colonne minime richieste: " + ", ".join(REQUIRED_COLS) + " | Colonna consigliata: unita_misura (se manca uso 'pz').")

if not uploaded:
    st.stop()

df = load_data(uploaded)
missing = validate_columns(df)
if missing:
    st.error(f"Colonne mancanti: {missing}")
    st.stop()

df = normalize_df(df)
if df.empty:
    st.error("Il file è stato letto ma non ci sono righe valide (controlla numeri e intestazioni).")
    st.stop()

metrics = compute_metrics(df)

st.subheader("Priorità riordino e rischio stockout")
show_cols = [
    "articolo",
    "unita_misura",
    "consumo_mensile",
    "stock_attuale",
    "domanda_lt",
    "scorta_sicurezza",
    "punto_riordino",
    "rischio_stockout",
    "qty_suggerita",
    "valore_unitario",
]
st.dataframe(metrics[show_cols], use_container_width=True)

st.download_button(
    "Scarica risultati (CSV)",
    metrics.to_csv(index=False).encode("utf-8"),
    file_name="supplychain_ai_results.csv",
    mime="text/csv",
)

st.subheader("Prompt AI per analisi decisionale")
articolo_sel = st.selectbox("Seleziona articolo", metrics["articolo"].astype(str).tolist())
row = metrics[metrics["articolo"].astype(str) == str(articolo_sel)].iloc[0]
prompt = genera_prompt(row)

st.text_area("Prompt pronto (copia e incolla in ChatGPT)", prompt, height=260)
