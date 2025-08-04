import streamlit as st
import pandas as pd
from openpyxl import Workbook, load_workbook
from datetime import datetime
import io
import base64

st.set_page_config(page_title="Export Moûts FOSS vers Dyostem", layout="wide", initial_sidebar_state="collapsed")

def create_file(ws, Nom_colonne):
    import string
    alphabet = list(string.ascii_uppercase)[:len(Nom_colonne)]
    for count, i in enumerate(Nom_colonne):
        ws[alphabet[count] + str(1)] = i
    return ws

def get_dates_from_file(uploaded_file):
    wb2 = load_workbook(filename=io.BytesIO(uploaded_file.read()))
    ws2 = wb2[wb2.sheetnames[0]]
    
    max_line = ws2.max_row
    dates = set()
    for row in range(2, max_line + 1):
        if ws2.cell(row=row, column=2).value == "Moûts":
            date_value = ws2.cell(row=row, column=4).value
            if date_value:
                try:
                    parsed_date = datetime.strptime(str(date_value)[:10], '%Y-%m-%d').strftime('%d/%m/%Y')
                    dates.add(parsed_date)
                except ValueError:
                    pass
    return sorted(list(dates))

def process_file(uploaded_file, selected_date):
    Nom_colonne = ["Nom de parcelle", "Cépage", "Date analyse", "Code échantillon", "Quantité Sucre (mg/baie)", "TAP (% vol)", "Acidité totale (g H2SO4/l)", "pH", "Acide malique (g/l)", "Acide tartrique", "Azote assimilable (mg/l)", "Potassium (g/l)", "Anthocyanes (mg/l)"]
    
    wb = Workbook()
    ws = wb.active
    ws = create_file(ws, Nom_colonne)
    
    wb2 = load_workbook(filename=io.BytesIO(uploaded_file.getvalue()))
    ws2 = wb2[wb2.sheetnames[0]]
    
    max_line = ws2.max_row
    
    ligne_ws = 1
    for row in range(2, max_line + 1):
        if ws2.cell(row=row, column=2).value == "Moûts":
            date_value = ws2.cell(row=row, column=4).value
            if date_value and datetime.strptime(str(date_value)[:10], '%Y-%m-%d').strftime('%d/%m/%Y') == selected_date:
                ligne_ws += 1
                ws.cell(row=ligne_ws, column=1).value = ws2.cell(row=row, column=3).value  # A: C
                ws.cell(row=ligne_ws, column=2).value = "Sauvignon blanc"  # B: Hardcoded
                ws.cell(row=ligne_ws, column=3).value = datetime.strptime(str(date_value)[:10], '%Y-%m-%d').strftime('%d/%m/%Y')  # C: D formatted
                ws.cell(row=ligne_ws, column=5).value = ws2.cell(row=row, column=5).value  # E: E
                ws.cell(row=ligne_ws, column=6).value = ws2.cell(row=row, column=6).value  # F: F
                ws.cell(row=ligne_ws, column=7).value = ws2.cell(row=row, column=7).value  # G: G
                ws.cell(row=ligne_ws, column=8).value = ws2.cell(row=row, column=8).value  # H: H
                ws.cell(row=ligne_ws, column=9).value = ws2.cell(row=row, column=9).value  # I: I
                ws.cell(row=ligne_ws, column=10).value = ws2.cell(row=row, column=10).value  # J: J
                ws.cell(row=ligne_ws, column=11).value = ws2.cell(row=row, column=11).value  # K: K
                ws.cell(row=ligne_ws, column=12).value = ws2.cell(row=row, column=14).value  # L: N
                ws.cell(row=ligne_ws, column=13).value = ws2.cell(row=row, column=16).value  # M: P
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# Custom CSS for Apple-like modern UI
st.markdown("""
    <style>
    .stApp {
        background-color: #f5f5f7;
        color: #1d1d1f;
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
    }
    .stButton > button {
        background-color: #0071e3;
        color: white;
        border-radius: 8px;
        border: none;
        padding: 10px 24px;
        font-size: 16px;
        font-weight: 600;
        transition: background-color 0.3s;
    }
    .stButton > button:hover {
        background-color: #0077ed;
    }
    .stSelectbox > div > div {
        background-color: white;
        border-radius: 8px;
        border: 1px solid #d2d2d7;
    }
    .stFileUploader > div > div {
        background-color: white;
        border-radius: 8px;
        border: 1px solid #d2d2d7;
        padding: 10px;
    }
    h1 {
        color: #1d1d1f;
        font-size: 32px;
        font-weight: 600;
        text-align: center;
        margin-bottom: 40px;
    }
    .stDownloadButton > button {
        background-color: #0071e3;
        color: white;
        border-radius: 8px;
        border: none;
        padding: 10px 24px;
        font-size: 16px;
        font-weight: 600;
        transition: background-color 0.3s;
        width: 100%;
    }
    .stDownloadButton > button:hover {
        background-color: #0077ed;
    }
    </style>
""", unsafe_allow_html=True)

st.title("Export Moûts FOSS vers Dyostem")

uploaded_file = st.file_uploader("Charger le fichier source (.xlsx)", type="xlsx")

if uploaded_file:
    with st.spinner("Analyse du fichier..."):
        dates = get_dates_from_file(uploaded_file)
    
    if dates:
        selected_date = st.selectbox("Sélectionner la date", dates, index=0)
        
        if st.button("Générer le fichier"):
            with st.spinner("Génération en cours..."):
                processed_file = process_file(uploaded_file, selected_date)
            
            st.download_button(
                label="Télécharger le fichier généré",
                data=processed_file,
                file_name="export_dyostem.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.warning("Aucune date valide trouvée dans le fichier.")