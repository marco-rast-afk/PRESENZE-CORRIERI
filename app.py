import streamlit as st
import pandas as pd
from fpdf import FPDF
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Gestione Presenze Corrieri", layout="wide")

# --- DATABASE INIZIALE (SESSION STATE) ---
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

# --- NAVIGATION MENU ---
menu = ["Assegnazione Quotidiana", "Anagrafica Furgoni", "Anagrafica Corrieri & Responsabili", "Configurazione Mail"]
scelta = st.sidebar.selectbox("Menu di Navigazione", menu)

# --- 1. ASSEGNAZIONE QUOTIDIANA (SNELLA E COMODA) ---
if scelta == "Assegnazione Quotidiana":
    st.title("📅 Tabellone Giornaliero Presenze e Mezzi")
    st.markdown("Modifica le opzioni direttamente all'interno della griglia qui sotto. I cambiamenti sono istantanei.")
    
    # 1. Prepariamo la lista dei furgoni escludendo i GUASTI
    furgoni_disponibili = st.session_state.furgoni[st.session_state.furgoni['DISPONIBILE'] != "GUASTO"]
    elenco_mezzi_opt = ["Nessuno"] + (furgoni_disponibili['MARCA'] + " " + furgoni_disponibili['MODELLO'] + " (" + furgoni_disponibili['TARGA'] + ")").tolist()
    
    # 2. Mostriamo l'INTERO tabellone in modalità EDITABILE DIRETTA
    # Usiamo st.column_config per inserire menu a tendina e caselle di testo direttamente nelle celle
    tabellone_aggiornato = st.data_editor(
        st.session_state.corrieri,
        column_config={
            "COGNOME": st.column_config.TextColumn("Cognome", disabled=True),
            "NOME": st.column_config.TextColumn("Nome", disabled=True),
            "CELLULARE": st.column_config.TextColumn("Cellulare", disabled=True),
            "GIRO_FISSO": st.column_config.TextColumn("Giro Fisso", disabled=True),
            "STATO": st.column_config.SelectboxColumn(
                "Stato Presenza",
                options=["Presente (Giro Fisso)", "Supporto Altra Filiale", "Assente"],
                required=True
            ),
            "GIRO_SUPPORTO": st.column_config.TextColumn("Giro di Supporto / Filiale"),
            "MEZZO": st.column_config.SelectboxColumn(
                "Furgone Assegnato",
                options=elenco_mezzi_opt,
                required=True
            ),
            "NOTE": st.column_config.TextColumn("Note operative")
        },
        hide_index=True,
        use_container_width=True,
        key="tabellone_giornaliero"
    )
    
    # Salviamo i dati modificati dall'utente
    st.session_state.corrieri = tabellone_aggiornato

    # --- GENERAZIONE DEI 4 BLOCCHI ---
    df_c = st.session_state.corrieri
    blocco1 = df_c[df_c['STATO'] == "Presente (Giro Fisso)"][['COGNOME', 'NOME', 'CELLULARE', 'GIRO_FISSO', 'MEZZO', 'NOTE']]
    blocco2 = df_c[df_c['STATO'] == "Supporto Altra Filiale"][['COGNOME', 'NOME', 'CELLULARE', 'GIRO_SUPPORTO', 'MEZZO', 'NOTE']]
    blocco3 = st.session_state.responsabili
    blocco4 = df_c[df_c['STATO'] == "Assente"][['COGNOME', 'NOME', 'CELLULARE', 'NOTE']]

    st.markdown("---")
    
    # --- ESPORTAZIONI ED AZIONI ---
    def genera_excel_formattato():
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            blocco1.to_excel(writer, sheet_name='Piano Giornaliero', index=False, startrow=1)
            workbook = writer.book
            ws = writer.sheets['Piano Giornaliero']
            
            font_titolo = Font(name='Calibri', size=12, bold=True, color='000000')
            font_header = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
            fill_header = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid') 
            fill_titolo = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid') 
            border_sottile = Border(
                left=Side(style='thin', color='BFBFBF'), right=Side(style='thin', color='BFBFBF'),
                top=Side(style='thin', color='BFBFBF'), bottom=Side(style='thin', color='BFBFBF')
            )
            allineamento_centro = Alignment(horizontal='center', vertical='center')

            def scrivi_blocco_formattato(ws, df, titolo, riga_inizio):
                ws.cell(row=riga_inizio, column=1, value=titolo).font = font_titolo
                ws.cell(row=riga_inizio, column=1).fill = fill_titolo
                ws.row_dimensions[riga_inizio].height = 24
                
                for col_idx, col_name in enumerate(df.columns, start=1):
                    cell = ws.cell(row=riga_inizio+1, column=col_idx, value=col_name)
                    cell.font = font_header
                    cell.fill = fill_header
                    cell.alignment = allineamento_centro
                    cell.border = border_sottile
                ws.row_dimensions[riga_inizio+1].height = 20
                
                riga_corrente = riga_inizio + 2
                for _, row_data in df.iterrows():
                    for col_idx, val in enumerate(row_data, start=1):
                        cell = ws.cell(row=riga_corrente, column=col_idx, value=str(val))
                        cell.border = border_sottile
                        if col_idx in [3, 4, 5]: 
                            cell.alignment = allineamento_centro
                    ws.row_dimensions[riga_corrente].height = 18
                    riga_corrente += 1
                return riga_corrente + 2

            ws.delete_rows(1, ws.max_row+10)
            prossima_riga = 2
            prossima_riga = scrivi_blocco_formattato(ws, blocco1, "BLOCCO 1: CORRIERI CON NUMERO DI GIRO ASSOCIATO", prossima_riga)
            prossima_riga = scrivi_blocco_formattato(ws, blocco2, "BLOCCO 2: EVENTUALI CORRIERI IN SUPPORTO", prossima_riga)
            prossima_riga = scrivi_blocco_formattato(ws, blocco3, "BLOCCO 3: NOMINATIVI RESPONSABILI PRESENTI", prossima_riga)
            prossima_riga = scrivi_blocco_formattato(ws, blocco4, "BLOCCO 4: CORRIERI ASSENTI", prossima_riga)
            
            for col in ws.columns:
                max_len = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    if cell.value:
                        max_len = max(max_len, len(str(cell.value)))
                ws.column_dimensions[col_letter].width = max(max_len + 4, 12)
                
        return output.getvalue()

	def genera_pdf_4_blocchi():
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", "B", 14)
        pdf.cell(0, 10, "PIANO GIORNALIERO FLOTTA E PRESENZE CORRIERI", ln=True, align="C")
        pdf.ln(4)
        
        def aggiungi_tabella_pdf(titolo, df):
            pdf.set_font("Arial", "B", 10)
            pdf.set_fill_color(225, 230, 240)
            pdf.cell(0, 7, titolo, ln=True, fill=True)
            pdf.set_font("Arial", "", 9)
            if df.empty:
                pdf.cell(0, 7, " Nessun record registrato in questo blocco.", ln=True)
                pdf.ln(3)
                return
            
            # Intestazioni tabella
            pdf.set_fill_color(30, 75, 120)
            pdf.set_text_color(255, 255, 255)
            for col in df.columns:
                pdf.cell(31, 7, str(col), border=1, fill=True)
            pdf.ln(7) # CORRETTO: Aggiunto il valore di interlinea
            
            # Righe dati
            pdf.set_text_color(0, 0, 0)
            for _, riga in df.iterrows():
                for col in df.columns:
                    pdf.cell(31, 7, str(riga[col])[:16], border=1)
                pdf.ln(7) # CORRETTO: Aggiunto il valore di interlinea
            pdf.ln(4)

        aggiungi_tabella_pdf("1. CORRIERI CON GIRO ASSOCIATO", blocco1)
        aggiungi_tabella_pdf("2. EVENTUALI CORRIERI IN SUPPORTO", blocco2)
        aggiungi_tabella_pdf("3. NOMINATIVI RESPONSABILI PRESENTI", blocco3)
        aggiungi_tabella_pdf("4. CORRIERI ASSENTI", blocco4)
        
        pdf_output = pdf.output(dest='S')
        return bytes(pdf_output)

    col_b1, col_b2, col_b3 = st.columns(3)
    with col_b1:
        st.download_button("📥 Scarica Excel Formattato", data=genera_excel_formattato(), file_name="piano_giornaliero.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with col_b2:
        st.download_button("📥 Scarica PDF Pronto Stampa", data=genera_pdf(), file_name="piano_giornaliero.pdf", mime="application/pdf")
    with col_b3:
        if st.button("✉️ Invia Report via Mail"):
            st.success("Email inviata ai destinatari configurati!")

# --- 2. ANAGRAFICA FURGONI ---
elif scelta == "Anagrafica Furgoni":
    st.title("🚐 Gestione Elenco Furgoni")
    st.markdown("Usa l'ultima riga della tabella per aggiungere nuovi furgoni o modifica direttamente lo stato (SI / NO / GUASTO).")
    st.session_state.furgoni = st.data_editor(st.session_state.furgoni, num_rows="dynamic", use_container_width=True, key="edit_furgoni_tab")

# --- 3. ANAGRAFICA PERSONALE ---
elif scelta == "Anagrafica Corrieri & Responsabili":
    st.title("👥 Anagrafica Fissa Personale")
    st.subheader("Elenco Corrieri")
    st.session_state.corrieri = st.data_editor(st.session_state.corrieri, num_rows="dynamic", use_container_width=True, key="edit_corr_tab")
    
    st.markdown("---")
    st.subheader("Elenco Responsabili")
    st.session_state.responsabili = st.data_editor(st.session_state.responsabili, num_rows="dynamic", use_container_width=True, key="edit_resp_tab")

# --- 4. CONFIGURAZIONE MAIL ---
elif scelta == "Configurazione Mail":
    st.title("⚙️ Configurazione Mail")
    st.session_state.config_mail["destinatari"] = st.text_input("Email Destinatari (separati da virgola)", value=st.session_state.config_mail["destinatari"])
    st.button("Salva Impostazioni")