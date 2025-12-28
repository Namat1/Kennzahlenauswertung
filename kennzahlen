import io
from datetime import datetime

import streamlit as st
import pandas as pd
import plotly.express as px

# Optional PDF Export
WEASYPRINT_OK = False
try:
    from weasyprint import HTML
    WEASYPRINT_OK = True
except Exception:
    pass

st.set_page_config(page_title="Fuhrpark – Grafische Auswertung", layout="wide")
st.title("📊 Grafische Auswertung nach Datum & Auswertungsart")

# -----------------------------
# Datenquelle
# -----------------------------
uploaded = st.file_uploader("Excel-Datei hochladen", type=["xlsx"])

if not uploaded:
    st.info("Bitte Excel-Datei hochladen")
    st.stop()

df_raw = pd.read_excel(uploaded, header=None)

# -----------------------------
# Strukturannahmen
# -----------------------------
DATE_COL = 0        # Spalte A
HEADER_ROW = 1      # Zeile 2 (0-basiert)
DATA_START = 2      # Ab Zeile 3

# -----------------------------
# Datum & Wochentag
# -----------------------------
df = df_raw.iloc[DATA_START:].copy()
df["Datum"] = pd.to_datetime(df_raw.iloc[DATA_START:, DATE_COL], errors="coerce")
df["Wochentag"] = df["Datum"].dt.day_name(locale="de_DE")

# Mapping auf Kurzform
weekday_map = {
    "Montag": "Mo",
    "Dienstag": "Di",
    "Mittwoch": "Mi",
    "Donnerstag": "Do",
    "Freitag": "Fr",
    "Samstag": "Sa",
    "Sonntag": "So"
}
df["Wochentag"] = df["Wochentag"].map(weekday_map)

# -----------------------------
# Auswertungsarten aus Zeile 2
# -----------------------------
metrics = {}
for col in range(1, df_raw.shape[1]):
    name = df_raw.iloc[HEADER_ROW, col]
    if pd.notna(name):
        metrics[col] = str(name).strip()

# -----------------------------
# Long-Format bauen
# -----------------------------
records = []

for col_idx, metric_name in metrics.items():
    values = df_raw.iloc[DATA_START:, col_idx]
    for i, val in enumerate(values):
        if pd.isna(val):
            continue
        records.append({
            "Datum": df.iloc[i]["Datum"],
            "Wochentag": df.iloc[i]["Wochentag"],
            "Auswertungsart": metric_name,
            "Wert": val
        })

long_df = pd.DataFrame(records)

# -----------------------------
# Dashboard
# -----------------------------
st.subheader("Kennzahlen")
c1, c2, c3 = st.columns(3)
c1.metric("Zeitraum Start", long_df["Datum"].min().strftime("%d.%m.%Y"))
c2.metric("Zeitraum Ende", long_df["Datum"].max().strftime("%d.%m.%Y"))
c3.metric("Auswertungsarten", long_df["Auswertungsart"].nunique())

# -----------------------------
# Grafiken je Auswertungsart
# -----------------------------
st.subheader("Grafische Auswertung nach Wochentag")

figs = []

for metric in long_df["Auswertungsart"].unique():
    sub = long_df[long_df["Auswertungsart"] == metric]

    agg = (
        sub.groupby("Wochentag")["Wert"]
        .sum()
        .reindex(["Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"])
        .dropna()
        .reset_index()
    )

    fig = px.bar(
        agg,
        x="Wochentag",
        y="Wert",
        title=f"{metric} pro Wochentag",
        labels={"Wert": metric}
    )

    st.plotly_chart(fig, use_container_width=True)
    figs.append((metric, fig))

# -----------------------------
# HTML Export
# -----------------------------
html_blocks = []
for metric, fig in figs:
    html_blocks.append(f"<h2>{metric}</h2>")
    html_blocks.append(fig.to_html(full_html=False, include_plotlyjs="cdn"))

html_report = f"""
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>Fuhrpark Auswertung</title>
</head>
<body>
<h1>Fuhrpark – Grafische Auswertung</h1>
<p>Erstellt am {datetime.now().strftime("%d.%m.%Y %H:%M")}</p>
{''.join(html_blocks)}
</body>
</html>
"""

st.download_button(
    "⬇️ HTML herunterladen",
    data=html_report.encode("utf-8"),
    file_name="fuhrpark_auswertung.html",
    mime="text/html"
)

# -----------------------------
# PDF Export
# -----------------------------
if WEASYPRINT_OK:
    pdf = HTML(string=html_report).write_pdf()
    st.download_button(
        "⬇️ PDF herunterladen",
        data=pdf,
        file_name="fuhrpark_auswertung.pdf",
        mime="application/pdf"
    )
else:
    st.warning("PDF-Export deaktiviert (WeasyPrint nicht installiert)")
