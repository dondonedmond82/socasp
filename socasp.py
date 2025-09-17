import panel as pn
import hvplot.pandas
import pandas as pd
import numpy as np
import pyodbc

pn.extension('tabulator', 'echarts')

# ===================== DATABASE CONNECTION =====================
conn_str = r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=./data/socasp.accdb;"
conn = pyodbc.connect(conn_str)
cursor = conn.cursor()

# ===================== LOAD & CLEAN DATA =====================
df_importations = pd.read_excel("./data/importation.xlsx")
df_importations['anneemois'] = pd.to_datetime(df_importations['anneemois'])
df_importations['annee'] = df_importations['anneemois'].dt.year.astype(str)
df_importations['mois'] = df_importations['anneemois'].dt.month_name()
df_importations['jour'] = df_importations['anneemois'].dt.day_name()

# Normalize column names
df_importations.columns = df_importations.columns.str.lower().str.strip()

required_cols = ['anneemois', 'marketeur', 'origine', 'annee', 'mois', 'jour', 
                 'provenance', 'essence', 'jet', 'petrole', 'gazoil']
for col in required_cols:
    if col not in df_importations.columns:
        raise ValueError(f"Column '{col}' missing from importations data")

# ===================== KPI CARDS =====================
def create_kpi_cards(df):
    if df.empty:
        return pn.pane.Markdown("⚠️ **Aucune donnée pour les filtres sélectionnés.**", style="color:red; font-weight:bold")
    total = df[['essence', 'jet', 'petrole', 'gazoil']].sum()
    total_volume = total.sum()
    card_style = {"background": "#f8f8f8", "padding": "10px", "border-radius": "8px", "text-align": "center", "width": "200px"}
    return pn.Row(
        pn.pane.Markdown(f"### **Total Volume**\n{total_volume:,.0f}", styles=card_style),
        pn.pane.Markdown(f"### **Essence**\n{total['essence']:,.0f}", styles=card_style),
        pn.pane.Markdown(f"### **Jet**\n{total['jet']:,.0f}", styles=card_style),
        pn.pane.Markdown(f"### **Gazoil**\n{total['gazoil']:,.0f}", styles=card_style),
    )

# ===================== FILTER WIDGETS =====================
select_marketeur = pn.widgets.MultiSelect(
    name='Marketeurs',
    options=sorted(df_importations['marketeur'].unique()),
    value=list(df_importations['marketeur'].unique()),
    size=8
)
select_origin = pn.widgets.Select(
    name='Origine',
    options=['Toutes'] + sorted(df_importations['origine'].dropna().unique().tolist()),
    value='Toutes'
)

# ===================== CHART FUNCTIONS =====================
def filter_data(marketeurs, origin):
    df_filtered = df_importations[df_importations['marketeur'].isin(marketeurs)]
    if origin != "Toutes":
        df_filtered = df_filtered[df_filtered['origine'] == origin]
    return df_filtered

def bar_chart(df):
    if df.empty:
        return pn.pane.Markdown("⚠️ Pas de données pour cette sélection.")
    df_grouped = df.groupby(['marketeur'])[['essence', 'jet', 'petrole', 'gazoil']].sum().reset_index()
    return df_grouped.hvplot.bar(x='marketeur', y=['essence', 'jet', 'petrole', 'gazoil'],
                                 stacked=True, width=850, height=400,
                                 title="Comparaison des Carburants par Marketeur (Agrégé)")

def line_chart(df):
    if df.empty:
        return pn.pane.Markdown("⚠️ Pas de données pour cette sélection.")
    df_grouped = df.groupby(['anneemois'])[['essence', 'jet', 'petrole', 'gazoil']].sum().reset_index()
    return df_grouped.hvplot.line(x='anneemois', y=['essence', 'jet', 'petrole', 'gazoil'],
                                  width=850, height=400, title="Tendance Mensuelle des Carburants")

def scatter_chart(df):
    if df.empty:
        return pn.pane.Markdown("⚠️ Pas de données pour cette sélection.")
    df_grouped = df.groupby(['marketeur'])[['essence', 'gazoil']].sum().reset_index()
    return df_grouped.hvplot.scatter(x='essence', y='gazoil', hover_cols=['marketeur'],
                                     width=850, height=400, title="Essence vs Gazoil par Marketeur")

def heatmap_chart(df):
    if df.empty:
        return pn.pane.Markdown("⚠️ Pas de données pour cette sélection.")
    pivot = df.pivot_table(index='marketeur', columns='anneemois', values='gazoil', aggfunc='sum').fillna(0)
    tidy = pivot.stack().reset_index()
    tidy.columns = ['marketeur', 'anneemois', 'gazoil']
    return tidy.hvplot.heatmap(x='anneemois', y='marketeur', C='gazoil', cmap='Viridis',
                               width=850, height=400, title="Heatmap des Gazoil par Marketeur et Mois")

def table_view(df):
    if df.empty:
        return pn.pane.Markdown("⚠️ Pas de données pour cette sélection.")
    df_grouped = df[['origine', 'provenance','marketeur', 'annee', 'mois', 'essence', 'jet', 'petrole', 'gazoil']]
    return pn.widgets.Tabulator(df_grouped, pagination='remote', page_size=10, width=850)

# ===================== DATA ENTRY FORM =====================
date_input = pn.widgets.DatetimeInput(name="Date", value=pd.Timestamp.today())
# marketeur_input = pn.widgets.TextInput(name="Marketeur")
# origine_input = pn.widgets.TextInput(name="Origine")
# provenance_input = pn.widgets.TextInput(name="Provenance")
essence_input = pn.widgets.IntInput(name="Essence", value=0, start=0)
jet_input = pn.widgets.IntInput(name="Jet", value=0, start=0)
petrole_input = pn.widgets.IntInput(name="Petrole", value=0, start=0)
gazoil_input = pn.widgets.IntInput(name="Gazoil", value=0, start=0)
status_message = pn.pane.Markdown("")

df_origine = pd.read_excel("./type/origine.xlsx")
df_provenance = pd.read_excel("./type/provenance.xlsx")
df_marketeur = pd.read_excel("./type/marketeur.xlsx")

unique_origine = list(df_origine['origine'].unique())
unique_provenance =   list(df_provenance['provenance'].unique())
unique_marketeur =  list(df_marketeur['marketeur'].unique())

origine_input = pn.widgets.Select(name="Origine", options=unique_origine)
provenance_input = pn.widgets.Select(name="Provenance", options=unique_provenance)
marketeur_input = pn.widgets.Select(name="Marketeur", options=unique_marketeur)

def add_record(event):
    try:
        cursor.execute("""
            INSERT INTO importations (anneemois, marketeur, origine, provenance, essence, jet, petrole, gazoil)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, date_input.value, marketeur_input.value, origine_input.value, provenance_input.value,
             essence_input.value, jet_input.value, petrole_input.value, gazoil_input.value)
        conn.commit()
        status_message.object = "✅ **Données insérées avec succès !**"
    except Exception as e:
        status_message.object = f"❌ Erreur : {e}"

submit_button = pn.widgets.Button(name="Ajouter", button_type="primary")
submit_button.on_click(add_record)

data_entry_form = pn.Column(
    "### ➕ Ajouter un enregistrement",
    date_input, origine_input, provenance_input, marketeur_input,
    essence_input, jet_input, petrole_input, gazoil_input,
    submit_button, status_message,
    width=325, styles={"background": "#eef6ff", "padding": "10px", "border-radius": "10px"}
)

# ===================== MAIN DASHBOARD =====================
@pn.depends(select_marketeur, select_origin)
def create_dashboard(marketeurs, origin):
    df_filtered = filter_data(marketeurs, origin)
    return pn.Column(
        "## 📊 Aperçu Global",
        create_kpi_cards(df_filtered),
        pn.Tabs(
            ("📊 Bar Chart", bar_chart(df_filtered)),
            ("📈 Line Chart", line_chart(df_filtered)),
            ("🔵 Scatter Plot", scatter_chart(df_filtered)),
            ("🔥 Heatmap", heatmap_chart(df_filtered)),
            ("📋 Table", table_view(df_filtered)),
        ),
        width=900
    )

sidebar = pn.Column(
    "## 🔎 Filtres",
    select_marketeur,
    select_origin,
    data_entry_form,
    sizing_mode="stretch_height",
    width=350,
    styles={"background": "#f5f5f5", "padding": "15px", "border-radius": "10px"}
)

main_area = pn.Column(create_dashboard, sizing_mode="stretch_width")

dashboard = pn.Row(sidebar, pn.Spacer(width=20), main_area, sizing_mode="stretch_width")
dashboard.servable()