import io
import re
from datetime import datetime
from typing import Dict, List, Tuple

import streamlit as st
import pandas as pd
import plotly.express as px

# -----------------------------
# Optional PDF Export (WeasyPrint)
# -----------------------------
WEASYPRINT_OK = False
try:
    from weasyprint import HTML  # type: ignore
    WEASYPRINT_OK = True
except Exception:
    WEASYPRINT_OK = False


# -----------------------------
# HTML Helper
# -----------------------------
def fig_to_html(fig) -> str:
    # Plotly Chart als HTML-Fragment (mit CDN PlotlyJS)
    return fig.to_html(full_html=False, include_plotlyjs="cdn")


def build_html_report(
    title: str,
    kpis: Dict[str, str],
    charts_html: List[Tuple[str, str]],
    table_df: pd.DataFrame
) -> str:
    css = """
    :root{--bg:#0f1115;--card:#171a21;--muted:#aab2c0;--text:#f2f4f8;--accent:#4aa3ff;}
    body{margin:0;font-family:Inter,system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif;background:var(--bg);color:var(--text);}
    .wrap{max-width:1200px;margin:0 auto;padding:28px;}
    h1{margin:0 0 6px 0;font-size:28px;letter-spacing:.2px;}
    .sub{color:var(--muted);margin-bottom:18px;}
    .grid{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin:14px 0 18px;}
    .card{background:var(--card);border:1px solid rgba(255,255,255,.07);border-radius:14px;padding:14px 14px 12px;}
    .kpi-label{color:var(--muted);font-size:12px;margin-bottom:6px;}
    .kpi-val{font-size:20px;font-weight:700;}
    .section{margin-top:18px;}
    .section h2{font-size:18px;margin:0 0 10px 0;}
    .chart{background:var(--card);border:1px solid rgba(255,255,255,.07);border-radius:14px;padding:10px;margin-bottom:12px;}
    table{width:100%;border-collapse:collapse;background:var(--card);border:1px solid rgba(255,255,255,.07);border-radius:14px;overflow:hidden;}
    th,td{padding:10px 10px;border-bottom:1px solid rgba(255,255,255,.06);font-size:13px;}
    th{color:var(--muted);text-align:left;font-weight:600;}
    tr:last-child td{border-bottom:none;}
    .pill{display:inline-block;padding:3px 8px;border-radius:999px;background:rgba(74,163,255,.15);border:1px solid rgba(74,163,255,.25);color:#cfe7ff;font-size:12px;}
    @media (max-width:1000px){.grid{grid-template-columns:repeat(2,1fr);}}
    @media (max-width:560px){.grid{grid-template-columns:1fr;}}
    """
    now = datetime.now().strftime("%d.%m.%Y %H:%M")

    kpi_cards = ""
    for label, val in kpis.items():
        kpi_cards += f"""
        <div class="card">
            <div class="kpi-label">{label}</div>
            <div class="kpi-val">{val}</div>
        </div>
        """

    charts_block = ""
    for headline, chart_div in charts_html:
        charts_block += f"""
        <div class="section">
            <h2>{headline}</h2>
            <div class="chart">{chart_div}</div>
        </div>
        """

    preview_table = table_df.copy()
    if len(preview_table) > 200:
        preview_table = preview_table.head(200)

    table_html = preview_table.to_html(index=False, escape=True)

    html = f"""
    <!doctype html>
    <html lang="de">
    <head>
      <meta charset="utf-8">
      <meta name="viewport" content="width=device-width, initial-scale=1">
      <title>{title}</title>
      <style>{css}</style>
    </head>
    <body>
      <div class="wrap">
        <h1>{title}</h1>
        <div class="sub">Generiert am <span class="pill">{now}</span></div>

        <div class="grid">
            {kpi_cards}
        </div>

        {charts_block}

        <div class="section">
          <h2>Beispiel-Listing (max. 200 Zeilen)</h2>
          {table_html}
          <div class="sub" style="margin-top:8px;">Hinweis: Im Streamlit-Dashboard kannst du natürlich alles filtern/sehen.</div>
        </div>
      </div>
    </body>
    </html>
    """
    return html


# -----------------------------
# Streamlit UI
# -----------------------------
st.set_page_config(page_title="Kennzahlenauswertung – Datum & Auswertungsart", layout="wide")
st.title("📊 Kennzahlenauswertung (Excel → Streamlit → HTML/PDF)")

with st.sidebar:
    st.header("Datenquelle")
    uploaded = st.file_uploader("Excel hochladen (.xlsx)", type=["xlsx"])
    st.caption("Optional: lokaler Pfad (nur lokal, nicht Streamlit Cloud).")
    local_path = st.text_input("Lokaler Excel-Pfad (optional)", value="")

    st.divider()
    st.header("Struktur")
    st.caption("Annahmen: Spalte A = Datum, Zeile 2 = Auswertungsarten, Daten ab Zeile 3.")
    show_detail_table = st.checkbox("Detailtabelle anzeigen", value=False)

# Load Excel
excel_bytes = None
if uploaded is not None:
    excel_bytes = uploaded.read()
elif local_path.strip():
    with open(local_path.strip(), "rb") as f:
        excel_bytes = f.read()

if not excel_bytes:
    st.info("Bitte Excel hochladen oder lokalen Pfad angeben.")
    st.stop()

xls = pd.ExcelFile(io.BytesIO(excel_bytes))
sheet = st.selectbox("Tabellenblatt wählen", xls.sheet_names, index=0)

# Einlesen ohne Header, damit wir Zeile 2 als "Spaltennamen" manuell übernehmen können
df_raw = pd.read_excel(io.BytesIO(excel_bytes), sheet_name=sheet, header=None)

st.write("### Rohdaten-Vorschau")
st.dataframe(df_raw.head(20), use_container_width=True)

# -----------------------------
# Struktur-Parameter (deine Vorgabe)
# -----------------------------
DATE_COL = 0        # Spalte A
HEADER_ROW = 1      # Zeile 2 (0-basiert)
DATA_START = 2      # Ab Zeile 3

# -----------------------------
# Datum + Wochentag OHNE locale (Streamlit Cloud safe)
# -----------------------------
df_dates = pd.DataFrame()
df_dates["Datum"] = pd.to_datetime(df_raw.iloc[DATA_START:, DATE_COL], errors="coerce")
df_dates = df_dates[df_dates["Datum"].notna()].copy()

# 0=Mo ... 6=So
wd = df_dates["Datum"].dt.weekday
df_dates["Wochentag"] = wd.map({0: "Mo", 1: "Di", 2: "Mi", 3: "Do", 4: "Fr", 5: "Sa", 6: "So"})

# Optionale feste Sortierung
weekday_order = ["Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"]
df_dates["Wochentag"] = pd.Categorical(df_dates["Wochentag"], categories=weekday_order, ordered=True)

# Wir müssen die Wertezeilen passend zu den gültigen Datumszeilen alignen:
# df_raw Datenbereich ab DATA_START, aber df_dates hat evtl. Zeilen mit ungültigem Datum entfernt.
# Daher merken wir uns die Original-Indizes:
valid_idx = df_raw.iloc[DATA_START:, DATE_COL].index[pd.to_datetime(df_raw.iloc[DATA_START:, DATE_COL], errors="coerce").notna()]

# -----------------------------
# Auswertungsarten aus Zeile 2 übernehmen
# -----------------------------
metrics: Dict[int, str] = {}
for col in range(1, df_raw.shape[1]):  # ab Spalte B
    name = df_raw.iloc[HEADER_ROW, col]
    if pd.isna(name):
        continue
    metric_name = str(name).strip()
    if metric_name == "":
        continue
    metrics[col] = metric_name

if not metrics:
    st.error("Keine Auswertungsarten in Zeile 2 gefunden (Spalten B..?).")
    st.stop()

st.write("### Erkannte Auswertungsarten (aus Zeile 2)")
st.write(list(metrics.values()))

# -----------------------------
# Long-Format bauen: Datum, Wochentag, Auswertungsart, Wert
# -----------------------------
records = []
for col_idx, metric_name in metrics.items():
    values = df_raw.loc[valid_idx, col_idx]  # nur Zeilen mit gültigem Datum
    for i, val in values.items():
        if pd.isna(val):
            continue
        # Datum/Wochentag zu diesem Index holen
        # Position in df_dates über valid_idx
        pos = list(valid_idx).index(i)
        records.append({
            "Datum": df_dates.iloc[pos]["Datum"],
            "Wochentag": df_dates.iloc[pos]["Wochentag"],
            "Auswertungsart": metric_name,
            "Wert": val
        })

long_df = pd.DataFrame(records)
if long_df.empty:
    st.error("Keine Werte gefunden (alles leer/NaN?)")
    st.stop()

# Werte numerisch (falls Excel als Text kommt)
long_df["Wert"] = pd.to_numeric(long_df["Wert"], errors="coerce")
long_df = long_df[long_df["Wert"].notna()].copy()

# -----------------------------
# KPIs
# -----------------------------
st.subheader("Kennzahlen")
c1, c2, c3, c4 = st.columns(4)

c1.metric("Zeitraum Start", long_df["Datum"].min().strftime("%d.%m.%Y"))
c2.metric("Zeitraum Ende", long_df["Datum"].max().strftime("%d.%m.%Y"))
c3.metric("Auswertungsarten", str(long_df["Auswertungsart"].nunique()))
c4.metric("Datenpunkte", str(len(long_df)))

# -----------------------------
# Filter
# -----------------------------
st.subheader("Filter")
sel_metrics = st.multiselect(
    "Auswertungsarten auswählen",
    options=sorted(long_df["Auswertungsart"].unique().tolist()),
    default=sorted(long_df["Auswertungsart"].unique().tolist())
)

filtered = long_df[long_df["Auswertungsart"].isin(sel_metrics)].copy()

# -----------------------------
# Grafiken je Auswertungsart
# -----------------------------
st.subheader("Grafische Auswertung nach Wochentag (Summe)")

figs: List[Tuple[str, object]] = []

for metric in sel_metrics:
    sub = filtered[filtered["Auswertungsart"] == metric].copy()
    if sub.empty:
        continue

    agg = (
        sub.groupby("Wochentag")["Wert"]
        .sum()
        .reindex(weekday_order)
        .dropna()
        .reset_index()
    )

    fig = px.bar(
        agg,
        x="Wochentag",
        y="Wert",
        title=f"{metric} pro Wochentag (Summe)",
        labels={"Wert": metric}
    )
    st.plotly_chart(fig, use_container_width=True)
    figs.append((metric, fig))

if show_detail_table:
    st.subheader("Detailtabelle (Long-Format)")
    st.dataframe(filtered.sort_values(["Auswertungsart", "Datum"]), use_container_width=True)

# -----------------------------
# HTML + PDF Export
# -----------------------------
st.subheader("📄 Export")

kpis = {
    "Zeitraum Start": long_df["Datum"].min().strftime("%d.%m.%Y"),
    "Zeitraum Ende": long_df["Datum"].max().strftime("%d.%m.%Y"),
    "Auswertungsarten": str(long_df["Auswertungsart"].nunique()),
    "Datenpunkte": str(len(long_df)),
}

charts_html: List[Tuple[str, str]] = []
for metric, fig in figs:
    charts_html.append((f"{metric} pro Wochentag (Summe)", fig_to_html(fig)))

html_report = build_html_report(
    title="Kennzahlenauswertung – Datum & Auswertungsart",
    kpis=kpis,
    charts_html=charts_html,
    table_df=filtered
)

st.download_button(
    "⬇️ HTML-Report herunterladen",
    data=html_report.encode("utf-8"),
    file_name="kennzahlenauswertung.html",
    mime="text/html"
)

if WEASYPRINT_OK:
    try:
        pdf_bytes = HTML(string=html_report).write_pdf()
        st.download_button(
            "⬇️ PDF-Report herunterladen",
            data=pdf_bytes,
            file_name="kennzahlenauswertung.pdf",
            mime="application/pdf"
        )
        st.success("PDF-Export bereit (WeasyPrint).")
    except Exception as e:
        st.error(f"PDF-Export fehlgeschlagen: {e}")
else:
    st.warning(
        "PDF-Export ist deaktiviert, weil WeasyPrint nicht installiert ist.\n\n"
        "Lokale Installation: `pip install weasyprint` (ggf. System-Abhängigkeiten nötig)."
    )
