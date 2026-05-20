import streamlit as st
import pandas as pd
from fpdf import FPDF
import io
import json
import os
from datetime import datetime
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Gestione Presenze Corrieri", layout="wide")

# File JSON locale per il salvataggio permanente dei dati
DB_FILE = "database_presenze.json"

# Dizionario per convertire il numero del mese nel nome in italiano maiuscolo
MESI_ITA = {
    1: "GENNAIO", 2: "FEBBRAIO", 3: "MARZO", 4: "APRILE", 5: "MAGGIO", 6: "GIUGNO",
    7: "LUGLIO", 8: "AGOSTO", 9: "SETTEMBRE", 10: "OTTOBRE", 11: "NOVEMBRE", 12: "DICEMBRE"
}

# --- FUNZIONI DI SALVATAGGIO E CARICAMENTO JSON ---
def salva_database_json():
    data_to_save = {
        "furgoni": st.session_state.furgoni.to_dict(orient="records"),
        "anagrafica_corrieri": st.session_state.anagrafica_corrieri.to_dict(orient="records"),
        "responsabili": st.session_state.responsabili.to_dict(orient="records"),
        "stato_giornaliero": st.session_state.stato_giornaliero.to_dict(orient="records")
    }
    with open(DB_FILE, "w", encoding="utf-8") as f:
        json.dump(data_to_save, f, ensure_ascii=False, indent=4)

def carica_database_json():
    if os.path.exists(DB_FILE):
        try:
            with open(DB_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                st.session_state.furgoni = pd.DataFrame(data["furgoni"])
                st.session_state.anagrafica_corrieri = pd.DataFrame(data["anagrafica_corrieri"])
                st.session_state.responsabili = pd.DataFrame(data["responsabili"])
                st.session_state.stato_giornaliero = pd.DataFrame(data["stato_giornaliero"])
                return True
        except Exception as e:
            st.error(f"Errore nel caricamento del database JSON: {e}")
    return False

# --- INIZIALIZZAZIONE DATABASE ---
if 'database_caricato' not in st.session_state:
    if carica_database_json():
        st.session_state.database_caricato = True
    else:
        # Dati di fabbrica (omessi qui per brevità, mantieni i tuoi array originali)
        st.session_state.furgoni = pd.DataFrame([{"MARCA": "Fiat", "MODELLO": "Ducato", "TIPO": "Furgone", "TARGA": "EP800MZ", "DISPONIBILE": "SI"}])
        st.session_state.anagrafica_corrieri = pd.DataFrame([{"COGNOME": "CROCI", "NOME": "MARINO", "CELLULARE": "3314509080", "GIRO_FISSO": "1"}])
        st.session_state.responsabili = pd.DataFrame([{"COGNOME": "ROSSI", "NOME": "LUIGI", "RUOLO": "Responsabile Logistica"}])
        
        df_giorno = st.session_state.anagrafica_corrieri.copy()
        df_giorno["STATO"] = "Presente (Giro Fisso)"
        df_giorno["GIRO_SUPPORTO"] = ""
        df_giorno["MEZZO"] = "Nessuno"
        df_giorno["NOTE"] = ""
        st.session_state.stato_giornaliero = df_giorno
        st.session_state.database_caricato = True

# --- BARRA LATERALE FISSA ---
menu = ["📋 Tabellone Presenze", "🚐 Anagrafica Furgoni", "👥 Anagrafica Personale"]
scelta = st.sidebar.selectbox("Navigazione", menu)

st.sidebar.markdown("---")
st.sidebar.subheader("💾 Salvataggio Dati")
if st.sidebar.button("💾 SALVA SU DATABASE", use_container_width=True, type="primary"):
    salva_database_json()
    st.sidebar.success("Database JSON salvato con successo!")

# --- FUNZIONI DI EXPORT FILE ---
def genera_excel_4_blocchi(data_label, b1, b2, b3, b4):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        b1.to_excel(writer, sheet_name='Piano Giornaliero', index=False, startrow=2)
        ws = writer.sheets['Piano Giornaliero']
        font_data_top = Font(name='Calibri', size=11, bold=True, color='000000')
        ws.cell(row=1, column=2, value="PRESENZE CORRIERI").font = font_data_top
        ws.cell(row=1, column=3, value=f"DATA: {data_label}").font = font_data_top
        
        font_titolo = Font(name='Calibri', size=11, bold=True, color='000000')
        font_header = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
        fill_header = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid') 
        fill_titolo = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid') 
        border_grigio = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'), top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))
        
        def scrivi_blocco_excel(ws, df, titolo_blocco, riga_partenza):
            ws.cell(row=riga_partenza, column=1, value=titolo_blocco).font = font_titolo
            ws.cell(row=riga_partenza, column=1).fill = fill_titolo
            for col_idx, col_name in enumerate(df.columns, start=1):
                cell = ws.cell(row=riga_partenza+1, column=col_idx, value=col_name)
                cell.font = font_header
                cell.fill = fill_header
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = border_grigio
            curr_row = riga_partenza + 2
            for _, riga_dati in df.iterrows():
                for col_idx, val in enumerate(riga_dati, start=1):
                    cell = ws.cell(row=curr_row, column=col_idx, value=str(val) if pd.notna(val) else "")
                    cell.border = border_grigio
                curr_row += 1
            return curr_row + 2

        ws.delete_rows(3, ws.max_row+20)
        prossima_riga = 3
        prossima_riga = scrivi_blocco_excel(ws, b1, "1° BLOCCO: CORRIERI CON NUMERO DI GIRO ASSOCIATO", prossima_riga)
        prossima_riga = scrivi_blocco_excel(ws, b2, "2° BLOCCO: CORRIERI IN SUPPORTO ALTRA FILIALE", prossima_riga)
        prossima_riga = scrivi_blocco_excel(ws, b3, "3° BLOCCO: NOMINATIVI RESPONSABILI PRESENTI", prossima_riga)
        prossima_riga = scrivi_blocco_excel(ws, b4, "4° BLOCCO: CORRIERI ASSENTI", prossima_riga)
        
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 13)
    return output.getvalue()

def genera_pdf_4_blocchi(data_label, b1, b2, b3, b4):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, f"PIANO GIORNALIERO FLOTTA E PRESENZE CORRIERI", align="C")
    pdf.ln(12)
    
    def aggiungi_tabella_pdf(titolo, df):
        pdf.set_font("Arial", "B", 10)
        pdf.set_fill_color(225, 230, 240)
        pdf.cell(0, 7, titolo, fill=True)
        pdf.ln(9)
        pdf.set_font("Arial", "", 9)
        if df.empty:
            pdf.cell(0, 7, " Nessun record.")
            pdf.ln(10)
            return
        pdf.set_fill_color(30, 75, 120)
        pdf.set_text_color(255, 255, 255)
        for col in df.columns:
            pdf.cell(31, 7, str(col), border=1, fill=True)
        pdf.ln(7)
        pdf.set_text_color(0, 0, 0)
        for _, riga in df.iterrows():
            for col in df.columns:
                pdf.cell(31, 7, str(riga[col])[:16], border=1)
            pdf.ln(7)
        pdf.ln(4)
        
    aggiungi_tabella_pdf("1. CORRIERI CON GIRO ASSOCIATO", b1)
    aggiungi_tabella_pdf("2. EVENTUALI CORRIERI IN SUPPORTO", b2)
    aggiungi_tabella_pdf("3. NOMINATIVI RESPONSABILI PRESENTI", b3)
    aggiungi_tabella_pdf("4. CORRIERI ASSENTI", b4)
    return bytes(pdf.output(dest='S'))


# --- 1. FRAMMENTO ISOLATO TABELLONE PRESENZE ---
@st.fragment
def mostra_tabellone_presenze():
    data_lavorazione = st.date_input("Data di lavorazione del Piano Presenze", datetime.today())
    giorno = data_lavorazione.day
    anno = data_lavorazione.year
    mese_testo = MESI_ITA[data_lavorazione.month]
    data_formato_personalizzato = f"{giorno} {mese_testo} {anno}"

    st.info("💡 Le modifiche qui dentro sono isolate: la pagina NON salterà all'inizio ad ogni clic.")
    
    furgoni_attivi = st.session_state.furgoni[st.session_state.furgoni['DISPONIBILE'] != "GUASTO"]
    elenco_furgoni_tendina = ["Nessuno"] + (furgoni_attivi['MARCA'] + " " + furgoni_attivi['MODELLO'] + " [" + furgoni_attivi['TARGA'] + "]").tolist()
    
    def salva_tabellone_giornaliero():
        if "editor_giornaliero_diretto" in st.session_state:
            edits = st.session_state["editor_giornaliero_diretto"]
            df_attuale = st.session_state.stato_giornaliero.copy()
            
            for row_idx, deltas in edits["edited_rows"].items():
                for col, val in deltas.items():
                    df_attuale.iat[row_idx, df_attuale.columns.get_loc(col)] = val
            
            # Controllo incrociato duplicati furgoni
            for idx, row in df_attuale.iterrows():
                nuovo_mezzo = row["MEZZO"]
                if nuovo_mezzo != "Nessuno":
                    for alt_idx, alt_row in df_attuale.iterrows():
                        if alt_idx != idx and alt_row["MEZZO"] == nuovo_mezzo:
                            df_attuale.iat[alt_idx, df_attuale.columns.get_loc("MEZZO")] = "Nessuno"
            
            st.session_state.stato_giornaliero = df_attuale

    st.data_editor(
        st.session_state.stato_giornaliero,
        column_config={
            "COGNOME": st.column_config.TextColumn("Cognome", disabled=True),
            "NOME": st.column_config.TextColumn("Nome", disabled=True),
            "CELLULARE": st.column_config.TextColumn("Cellulare", disabled=True),
            "GIRO_FISSO": st.column_config.TextColumn("Giro Fisso", disabled=True),
            "STATO": st.column_config.SelectboxColumn("Stato Presenza", options=["Presente (Giro Fisso)", "Supporto Altra Filiale", "Assente"], required=True, width="medium"),
            "GIRO_SUPPORTO": st.column_config.TextColumn("Giro di Supporto / Filiale", width="medium"),
            "MEZZO": st.column_config.SelectboxColumn("Furgone Assegnato", options=elenco_furgoni_tendina, required=True, width="large"),
            "NOTE": st.column_config.TextColumn("Note Operative", width="large")
        },
        hide_index=True,
        use_container_width=True,
        key="editor_giornaliero_diretto",
        on_change=salva_tabellone_giornaliero
    )

    # Calcolo dei 4 blocchi basato sullo stato corrente in tempo reale
    df_correnti = st.session_state.stato_giornaliero
    blocco1 = df_correnti[df_correnti['STATO'] == "Presente (Giro Fisso)"][['COGNOME', 'NOME', 'CELLULARE', 'GIRO_FISSO', 'MEZZO', 'NOTE']]
    blocco2 = df_correnti[df_correnti['STATO'] == "Supporto Altra Filiale"][['COGNOME', 'NOME', 'CELLULARE', 'GIRO_SUPPORTO', 'MEZZO', 'NOTE']]
    blocco3 = st.session_state.responsabili
    blocco4 = df_correnti[df_correnti['STATO'] == "Assente"][['COGNOME', 'NOME', 'CELLULARE', 'NOTE']]

    st.markdown("---")
    col_x1, col_x2 = st.columns(2)
    with col_x1:
        st.download_button(
            label=f"📥 Scarica Excel", 
            data=genera_excel_4_blocchi(data_formato_personalizzato, blocco1, blocco2, blocco3, blocco4), 
            file_name=f"Presenze {data_formato_personalizzato}.xlsx", 
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    with col_x2:
        st.download_button(
            label=f"📥 Scarica PDF", 
            data=genera_pdf_4_blocchi(data_formato_personalizzato, blocco1, blocco2, blocco3, blocco4), 
            file_name=f"Presenze {data_formato_personalizzato}.pdf", 
            mime="application/pdf",
            use_container_width=True
        )


# --- 2. FRAMMENTO ISOLATO ANAGRAFICA FURGONI ---
@st.fragment
def mostra_anagrafica_furgoni():
    st.title("🚐 Anagrafica e Stato Mezzi Aziendali")
    
    def salva_furgoni_edits():
        if "tabella_gestione_furgoni" in st.session_state:
            edits = st.session_state["tabella_gestione_furgoni"]
            df_attuale = st.session_state.furgoni.copy()
            for row_idx, deltas in edits["edited_rows"].items():
                for col, val in deltas.items():
                    df_attuale.iat[row_idx, df_attuale.columns.get_loc(col)] = val
            if edits["deleted_rows"]:
                df_attuale = df_attuale.drop(edits["deleted_rows"]).reset_index(drop=True)
            if edits["added_rows"]:
                for new_row in edits["added_rows"]:
                    riga_pulita = {col: new_row.get(col, "") for col in df_attuale.columns}
                    df_attuale = pd.concat([df_attuale, pd.DataFrame([riga_pulita])], ignore_index=True)
            st.session_state.furgoni = df_attuale

    st.data_editor(
        st.session_state.furgoni, 
        num_rows="dynamic", 
        column_config={"DISPONIBILE": st.column_config.SelectboxColumn("Disponibilità", options=["SI", "NO", "GUASTO"], required=True)},
        use_container_width=True,
        key="tabella_gestione_furgoni",
        on_change=salva_furgoni_edits
    )


# --- 3. FRAMMENTO ISOLATO ANAGRAFICA PERSONALE ---
@st.fragment
def mostra_anagrafica_personale():
    st.title("👥 Gestione Personale e Autisti")
    st.subheader("Anagrafica Fissa Corrieri")
    
    def aggiorna_anagrafica_corrieri():
        if "tabella_gestione_corrieri" in st.session_state:
            edits = st.session_state["tabella_gestione_corrieri"]
            df_attuale = st.session_state.anagrafica_corrieri.copy()
            
            for row_idx, deltas in edits["edited_rows"].items():
                for col, val in deltas.items():
                    df_attuale.iat[row_idx, df_attuale.columns.get_loc(col)] = val
            if edits["deleted_rows"]:
                df_attuale = df_attuale.drop(edits["deleted_rows"]).reset_index(drop=True)
            if edits["added_rows"]:
                for new_row in edits["added_rows"]:
                    riga_pulita = {col: new_row.get(col, "") for col in df_attuale.columns}
                    df_attuale = pd.concat([df_attuale, pd.DataFrame([riga_pulita])], ignore_index=True)
            
            st.session_state.anagrafica_corrieri = df_attuale
            
            # Sincronizzazione automatica del tabellone quotidiano
            df_nuovo = df_attuale.copy()
            df_nuovo = df_nuovo.merge(st.session_state.stato_giornaliero[['COGNOME', 'NOME', 'STATO', 'GIRO_SUPPORTO', 'MEZZO', 'NOTE']], on=['COGNOME', 'NOME'], how='left')
            df_nuovo["STATO"] = df_nuovo["STATO"].fillna("Presente (Giro Fisso)")
            df_nuovo["GIRO_SUPPORTO"] = df_nuovo["GIRO_SUPPORTO"].fillna("")
            df_nuovo["MEZZO"] = df_nuovo["MEZZO"].fillna("Nessuno")
            df_nuovo["NOTE"] = df_nuovo["NOTE"].fillna("")
            st.session_state.stato_giornaliero = df_nuovo

    st.data_editor(
        st.session_state.anagrafica_corrieri, 
        num_rows="dynamic", 
        use_container_width=True, 
        key="tabella_gestione_corrieri",
        on_change=aggiorna_anagrafica_corrieri
    )
    
    st.markdown("---")
    st.subheader("Anagrafica Fissa Responsabili / Capi Turno")
    
    def aggiorna_responsabili():
        if "tabella_gestione_responsabili" in st.session_state:
            edits = st.session_state["tabella_gestione_responsabili"]
            df_attuale = st.session_state.responsabili.copy()
            for row_idx, deltas in edits["edited_rows"].items():
                for col, val in deltas.items():
                    df_attuale.iat[row_idx, df_attuale.columns.get_loc(col)] = val
            if edits["deleted_rows"]:
                df_attuale = df_attuale.drop(edits["deleted_rows"]).reset_index(drop=True)
            if edits["added_rows"]:
                for new_row in edits["added_rows"]:
                    riga_pulita = {col: new_row.get(col, "") for col in df_attuale.columns}
                    df_attuale = pd.concat([df_attuale, pd.DataFrame([riga_pulita])], ignore_index=True)
            st.session_state.responsabili = df_attuale

    st.data_editor(
        st.session_state.responsabili, 
        num_rows="dynamic", 
        use_container_width=True, 
        key="tabella_gestione_responsabili",
        on_change=aggiorna_responsabili
    )


# --- LOGICA DEL MENU PRINCIPALE ---
if scelta == "📋 Tabellone Presenze":
    st.title("📋 Inserimento Presenze e Assegnazione Mezzi")
    mostra_tabellone_presenze()

elif scelta == "🚐 Anagrafica Furgoni":
    mostra_anagrafica_furgoni()

elif scelta == "👥 Anagrafica Personale":
    mostra_anagrafica_personale()
