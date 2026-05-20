import streamlit as st
import pandas as pd
from fpdf import FPDF
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

# Librerie per la formattazione avanzata di Excel
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Gestione Presenze Corrieri", layout="wide")

# --- INIZIALIZZAZIONE DATABASE TEMPORANEO (SESSION STATE) ---
if 'furgoni' not in st.session_state:
    st.session_state.furgoni = pd.DataFrame([
        {"MARCA": "Fiat", "MODELLO": "Ducato", "TIPO": "Furgone", "TARGA": "AA123BB", "DISPONIBILE": "SI"},
        {"MARCA": "Ford", "MODELLO": "Transit", "TIPO": "Furgone", "TARGA": "CC456DD", "DISPONIBILE": "NO"},
        {"MARCA": "Iveco", "MODELLO": "Daily", "TIPO": "Cassonato", "TARGA": "EE789FF", "DISPONIBILE": "GUASTO"},
    ])

if 'corrieri' not in st.session_state:
    st.session_state.corrieri = pd.DataFrame([
        {"COGNOME": "Rossi", "NOME": "Mario", "CELLULARE": "333123456", "GIRO_FISSO": "Giro 1", "STATO": "Presente (Giro Fisso)", "GIRO_SUPPORTO": "", "MEZZO": "Nessuno", "NOTE": ""},
        {"COGNOME": "Bianchi", "NOME": "Luigi", "CELLULARE": "333987654", "GIRO_FISSO": "Giro 2", "STATO": "Presente (Giro Fisso)", "GIRO_SUPPORTO": "", "MEZZO": "Nessuno", "NOTE": ""},
        {"COGNOME": "Verdi", "NOME": "Anna", "CELLULARE": "333555444", "GIRO_FISSO": "Giro 3", "STATO": "Supporto Altra Filiale", "GIRO_SUPPORTO": "Milano", "MEZZO": "Nessuno", "NOTE": ""},
        {"COGNOME": "Neri", "NOME": "Pietro", "CELLULARE": "333222111", "GIRO_FISSO": "Giro 4", "STATO": "Assente", "GIRO_SUPPORTO": "", "MEZZO": "Nessuno", "NOTE": "Malattia"},
    ])

if 'responsabili' not in st.session_state:
    st.session_state.responsabili = pd.DataFrame([
        {"COGNOME": "Capo", "NOME": "Francesco", "RUOLO": "Responsabile Logistica"},
        {"COGNOME": "Vice", "NOME": "Elena", "RUOLO": "Supervisore"}
    ])

if 'config_mail' not in st.session_state:
    st.session_state.config_mail = {
        "destinatari": "ufficio@esempio.com",
        "smtp_server": "smtp.esempio.com",
        "smtp_port": 587,
        "email_invio": "logistica@esempio.com",
        "password": ""
    }

# --- MENU ---
menu = ["Assegnazione Quotidiana", "Anagrafica Furgoni", "Anagrafica Corrieri & Responsabili", "Configurazione Mail"]
scelta = st.sidebar.selectbox("Menu di Navigazione", menu)

# --- 1. ASSEGNAZIONE QUOTIDIANA ---
if scelta == "Assegnazione Quotidiana":
    st.title("📅 Assegnazione Giornaliera Corrieri e Mezzi")
    
    # Prepara elenco furgoni escludendo i GUASTI
    furgoni_disponibili = st.session_state.furgoni[st.session_state.furgoni['DISPONIBILE'] != "GUASTO"]
    elenco_mezzi_opt = ["Nessuno"] + (furgoni_disponibili['MARCA'] + " " + furgoni_disponibili['MODELLO'] + " (" + furgoni_disponibili['TARGA'] + ")").tolist()
    
    st.subheader("Elenco Corrieri in Organico")
    for idx, row in st.session_state.corrieri.iterrows():
        with st.expander(f"🚚 {row['COGNOME']} {row['NOME']} - (Giro Originario: {row['GIRO_FISSO']})"):
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                nuovo_stato = col1.selectbox("Stato", ["Presente (Giro Fisso)", "Supporto Altra Filiale", "Assente"], index=["Presente (Giro Fisso)", "Supporto Altra Filiale", "Assente"].index(row['STATO']), key=f"st_{idx}")
            with col2:
                index_mezzo = elenco_mezzi_opt.index(row['MEZZO']) if row['MEZZO'] in elenco_mezzi_opt else 0
                nuovo_mezzo = col2.selectbox("Mezzo Assegnato", elenco_mezzi_opt, index=index_mezzo, key=f"mz_{idx}")
            with col3:
                nuovo_supporto = col3.text_input("Giro di Supporto", value=row['GIRO_SUPPORTO'], key=f"sp_{idx}")
            with col4:
                nuove_note = col4.text_input("Note / Motivazioni", value=row['NOTE'], key=f"nt_{idx}")
            
            # Salva nello stato
            st.session_state.corrieri.at[idx, 'STATO'] = nuovo_stato
            st.session_state.corrieri.at[idx, 'MEZZO'] = nuovo_mezzo
            st.session_state.corrieri.at[idx, 'GIRO_SUPPORTO'] = nuovo_supporto
            st.session_state.corrieri.at[idx, 'NOTE'] = nuove_note

    # Suddivisione nei 4 blocchi richiesti
    df_c = st.session_state.corrieri
    blocco1 = df_c[df_c['STATO'] == "Presente (Giro Fisso)"][['COGNOME', 'NOME', 'CELLULARE', 'GIRO_FISSO', 'MEZZO', 'NOTE']]
    blocco2 = df_c[df_c['STATO'] == "Supporto Altra Filiale"][['COGNOME', 'NOME', 'CELLULARE', 'GIRO_SUPPORTO', 'MEZZO', 'NOTE']]
    blocco3 = st.session_state.responsabili
    blocco4 = df_c[df_c['STATO'] == "Assente"][['COGNOME', 'NOME', 'CELLULARE', 'NOTE']]

    st.markdown("---")
    st.subheader("Anteprima Tabella di Output")
    st.dataframe(blocco1, title="1. CORRIERI CON NUMERO DI GIRO ASSOCIATO")
    st.dataframe(blocco2, title="2. CORRIERI IN SUPPORTO")
    st.dataframe(blocco3, title="3. RESPONSABILI PRESENTI")
    st.dataframe(blocco4, title="4. CORRIERI ASSENTI")

    # --- MOTORE DI FORMATTAZIONE EXCEL AVANZATO ---
    def genera_excel_formattato():
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Creiamo un foglio vuoto per scriverci dentro cella per cella
            blocco1.to_excel(writer, sheet_name='Piano Giornaliero', index=False, startrow=1)
            workbook = writer.book
            ws = writer.sheets['Piano Giornaliero']
            
            # Stili Grafici
            font_titolo = Font(name='Calibri', size=12, bold=True, color='000000')
            font_header = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
            fill_header = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid') # Blu scuro aziendale
            fill_titolo = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid') # Celeste chiaro per i titoli di blocco
            border_sottile = Border(
                left=Side(style='thin', color='BFBFBF'), right=Side(style='thin', color='BFBFBF'),
                top=Side(style='thin', color='BFBFBF'), bottom=Side(style='thin', color='BFBFBF')
            )
            allineamento_centro = Alignment(horizontal='center', vertical='center')

            def scrivi_blocco_formattato(ws, df, titolo, riga_inizio):
                # 1. Scrivi il Titolo del Blocco
                ws.cell(row=riga_inizio, column=1, value=titolo).font = font_titolo
                ws.cell(row=riga_inizio, column=1).fill = fill_titolo
                ws.row_dimensions[riga_inizio].height = 24
                
                # 2. Scrivi le Intestazioni delle Colonne
                for col_idx, col_name in enumerate(df.columns, start=1):
                    cell = ws.cell(row=riga_inizio+1, column=col_idx, value=col_name)
                    cell.font = font_header
                    cell.fill = fill_header
                    cell.alignment = allineamento_centro
                    cell.border = border_sottile
                ws.row_dimensions[riga_inizio+1].height = 20
                
                # 3. Scrivi i Dati
                riga_corrente = riga_inizio + 2
                for _, row_data in df.iterrows():
                    for col_idx, val in enumerate(row_data, start=1):
                        cell = ws.cell(row=riga_corrente, column=col_idx, value=str(val))
                        cell.border = border_sottile
                        if col_idx in [3, 4, 5]:  # Allinea al centro cellulare, giri e mezzi
                            cell.alignment = allineamento_centro
                    ws.row_dimensions[riga_corrente].height = 18
                    riga_corrente += 1
                
                return riga_corrente + 2 # Ritorna la riga per il blocco successivo lasciando spazio vuoto

            # Svuota la scrittura automatica iniziale e riscrivi tutto da capo in modo perfetto
            ws.delete_rows(1, ws.max_row+10)
            
            prossima_riga = 2
            prossima_riga = scrivi_blocco_formattato(ws, blocco1, "BLOCCO 1: CORRIERI CON NUMERO DI GIRO ASSOCIATO", prossima_riga)
            prossima_riga = scrivi_blocco_formattato(ws, blocco2, "BLOCCO 2: EVENTUALI CORRIERI IN SUPPORTO", prossima_riga)
            prossima_riga = scrivi_blocco_formattato(ws, blocco3, "BLOCCO 3: NOMINATIVI RESPONSABILI PRESENTI", prossima_riga)
            prossima_riga = scrivi_blocco_formattato(ws, blocco4, "BLOCCO 4: CORRIERI ASSENTI", prossima_riga)
            
            # Autoregolazione automatica della larghezza delle colonne
            for col in ws.columns:
                max_len = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    if cell.value:
                        max_len = max(max_len, len(str(cell.value)))
                ws.column_dimensions[col_letter].width = max(max_len + 4, 12)
                
        return output.getvalue()

    def genera_pdf():
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", "B", 16)
        pdf.cell(0, 12, "PIANO GIORNALIERO PRESENZE E FLOTTA", ln=True, align="C")
        pdf.ln(8)
        
        def aggiungi_tabella_pdf(titolo, df):
            pdf.set_font("Arial", "B", 11)
            # Sfondo grigio per il titolo del blocco nel PDF
            pdf.set_fill_color(220, 225, 230)
            pdf.cell(0, 8, titolo, ln=True, fill=True)
            pdf.set_font("Arial", "", 9)
            if df.empty:
                pdf.cell(0, 8, " Nessun nominativo", ln=True)
                pdf.ln(4)
                return
            
            # Intestazioni colonne
            pdf.set_fill_color(30, 75, 120)
            pdf.set_text_color(255, 255, 255)
            for col in df.columns:
                pdf.cell(31, 7, str(col), border=1, fill=True)
            pdf.ln()
            
            # Righe dati
            pdf.set_text_color(0, 0, 0)
            for _, riga in df.iterrows():
                for col in df.columns:
                    pdf.cell(31, 7, str(riga[col])[:16], border=1)
                pdf.ln()
            pdf.ln(6)

        aggiungi_tabella_pdf("1. CORRIERI CON GIRO ASSOCIATO", blocco1)
        aggiungi_tabella_pdf("2. EVENTUALI CORRIERI IN SUPPORTO", blocco2)
        aggiungi_tabella_pdf("3. NOMINATIVI RESPONSABILI", blocco3)
        aggiungi_tabella_pdf("4. CORRIERI ASSENTI", blocco4)
        return pdf.output(dest='S')

    # --- PULSANTI AZIONE ---
    c_btn1, c_btn2, c_btn3 = st.columns(3)
    with c_btn1:
        st.download_button("📥 Scarica Excel Formattato", data=genera_excel_formattato(), file_name="piano_corrieri.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with c_btn2:
        st.download_button("📥 Scarica PDF Pronto Stampa", data=genera_pdf(), file_name="piano_corrieri.pdf", mime="application/pdf")
    with c_btn3:
        if st.button("✉️ Invia per Email"):
            st.info("Funzione di invio attivata. Assicurati di aver configurato i dati SMTP nell'ultima scheda.")

# --- SCHEDE DI ANAGRAFICA (Invariate, vedi codice precedente per brevità) ---
elif scelta == "Anagrafica Furgoni":
    st.title("🚐 Elenco Furgoni e Stato Disponibilità")
    with st.form("Nuovo Furgone"):
        c1, c2, c3, c4, c5 = st.columns(5)
        marca = c1.text_input("MARCA")
        modello = c2.text_input("MODELLO")
        tipo = c3.text_input("TIPO")
        targa = c4.text_input("TARGA")
        disponibile = c5.selectbox("DISPONIBILE", ["SI", "NO", "GUASTO"])
        if st.form_submit_button("Inserisci Furgone") and targa:
            st.session_state.furgoni = pd.concat([st.session_state.furgoni, pd.DataFrame([{"MARCA": marca, "MODELLO": modello, "TIPO": tipo, "TARGA": targa, "DISPONIBILE": disponibile}])], ignore_index=True)
    st.data_editor(st.session_state.furgoni, num_rows="dynamic", key="furgoni_ed")

elif scelta == "Anagrafica Corrieri & Responsabili":
    st.title("👥 Gestione Anagrafiche Fisse")
    st.subheader("Corrieri")
    st.data_editor(st.session_state.corrieri, num_rows="dynamic", key="corr_ed")
    st.subheader("Responsabili")
    st.data_editor(st.session_state.responsabili, num_rows="dynamic", key="resp_ed")

elif scelta == "Configurazione Mail":
    st.title("⚙️ Impostazioni Posta Elettronica")
    # Campi per salvare la mail aziendale
    st.text_input("Destinatari (separati da virgola)", value=st.session_state.config_mail["destinatari"])
    st.button("Salva Impostazioni")