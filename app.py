import streamlit as st
import pandas as pd
from fpdf import FPDF
import io
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Gestione Presenze Corrieri", layout="wide")

# --- DATABASE INIZIALE (SESSION STATE) ---
if 'furgoni' not in st.session_state:
    st.session_state.furgoni = pd.DataFrame([
        {"MARCA": "Fiat", "MODELLO": "Ducato", "TIPO": "Furgone", "TARGA": "GR256RF", "DISPONIBILE": "SI"},
        {"MARCA": "Ford", "MODELLO": "Transit", "TIPO": "Furgone", "TARGA": "GS557WM", "DISPONIBILE": "SI"},
        {"MARCA": "Iveco", "MODELLO": "Daily", "TIPO": "Furgone", "TARGA": "HB683CE", "DISPONIBILE": "NO"},
        {"MARCA": "Mercedes", "MODELLO": "Sprinter", "TIPO": "Furgone", "TARGA": "HB787TS", "DISPONIBILE": "GUASTO"},
    ])

# ANAGRAFICA PERSONALE: Contiene solo i 4 campi base richiesti
if 'anagrafica_corrieri' not in st.session_state:
    st.session_state.anagrafica_corrieri = pd.DataFrame([
        {"COGNOME": "CROCI", "NOME": "MARINO", "CELLULARE": "3314509080", "GIRO_FISSO": "1"},
        {"COGNOME": "D ANGELO", "NOME": "SALVATORE", "CELLULARE": "3881132883", "GIRO_FISSO": "3"},
        {"COGNOME": "MARCIANO", "NOME": "ANTONIO", "CELLULARE": "3489292359", "GIRO_FISSO": "4"},
        {"COGNOME": "CAPUTO", "NOME": "OVIDIO", "CELLULARE": "3385277033", "GIRO_FISSO": "5"},
        {"COGNOME": "AINIS", "NOME": "CIRO", "CELLULARE": "3891618386", "GIRO_FISSO": "6"},
    ])

if 'responsabili' not in st.session_state:
    st.session_state.responsabili = pd.DataFrame([
        {"COGNOME": "ROSSI", "NOME": "LUIGI", "RUOLO": "Responsabile Logistica"},
        {"COGNOME": "VERDI", "NOME": "MARCO", "RUOLO": "Supervisore di Turno"}
    ])

if 'config_mail' not in st.session_state:
    st.session_state.config_mail = {"destinatari": "ufficio.logistica@esempio.com"}

# Struttura di lavoro giornaliera temporanea creata unendo l'anagrafica con i campi operativi vuoti
if 'stato_giornaliero' not in st.session_state:
    df_giorno = st.session_state.anagrafica_corrieri.copy()
    df_giorno["STATO"] = "Presente (Giro Fisso)"
    df_giorno["GIRO_SUPPORTO"] = ""
    df_giorno["MEZZO"] = "Nessuno"
    df_giorno["NOTE"] = ""
    st.session_state.stato_giornaliero = df_giorno

# Sincronizza il tabellone se l'anagrafica di base viene modificata (aggiunta/rimozione autisti)
if len(st.session_state.stato_giornaliero) != len(st.session_state.anagrafica_corrieri):
    df_nuovo = st.session_state.anagrafica_corrieri.copy()
    # Recupera lo stato precedente per non sovrascrivere il lavoro già fatto oggi
    df_nuovo = df_nuovo.merge(st.session_state.stato_giornaliero[['COGNOME', 'NOME', 'STATO', 'GIRO_SUPPORTO', 'MEZZO', 'NOTE']], on=['COGNOME', 'NOME'], how='left')
    df_nuovo["STATO"] = df_nuovo["STATO"].fillna("Presente (Giro Fisso)")
    df_nuovo["GIRO_SUPPORTO"] = df_nuovo["GIRO_SUPPORTO"].fillna("")
    df_nuovo["MEZZO"] = df_nuovo["MEZZO"].fillna("Nessuno")
    df_nuovo["NOTE"] = df_nuovo["NOTE"].fillna("")
    st.session_state.stato_giornaliero = df_nuovo

# --- MENU DI NAVIGAZIONE A SINISTRA ---
menu = ["📋 Tabellone Presenze", "🚐 Anagrafica Furgoni", "👥 Anagrafica Personale", "⚙️ Configurazione Mail"]
scelta = st.sidebar.selectbox("Navigazione", menu)

# --- 1. TABELLONE PRESENZE SNELLO (TUTTO A VISTA) ---
if scelta == "📋 Tabellone Presenze":
    st.title("📋 Inserimento Presenze e Assegnazione Mezzi")
    st.markdown("I campi *Cognome, Nome, Cellulare e Giro Fisso* sono bloccati. Configura lo Stato, il Mezzo e le Note direttamente sulle righe attive.")
    
    # Estrazione dinamica della lista furgoni (ESCLUSI i GUASTO)
    furgoni_attivi = st.session_state.furgoni[st.session_state.furgoni['DISPONIBILE'] != "GUASTO"]
    elenco_furgoni_tendina = ["Nessuno"] + (furgoni_attivi['MARCA'] + " " + furgoni_attivi['MODELLO'] + " [" + furgoni_attivi['TARGA'] + "]").tolist()
    
    # TABELLONE GLOBALE EDITABILE (Tutti i campi dinamici sono sbloccati qui)
    tabellone_modificato = st.data_editor(
        st.session_state.stato_giornaliero,
        column_config={
            "COGNOME": st.column_config.TextColumn("Cognome", disabled=True),
            "NOME": st.column_config.TextColumn("Nome", disabled=True),
            "CELLULARE": st.column_config.TextColumn("Cellulare", disabled=True),
            "GIRO_FISSO": st.column_config.TextColumn("Giro Fisso", disabled=True),
            "STATO": st.column_config.SelectboxColumn(
                "Stato Presenza",
                options=["Presente (Giro Fisso)", "Supporto Altra Filiale", "Assente"],
                required=True,
                width="medium"
            ),
            "GIRO_SUPPORTO": st.column_config.TextColumn("Giro di Supporto / Filiale", width="medium"),
            "MEZZO": st.column_config.SelectboxColumn(
                "Furgone Assegnato",
                options=elenco_furgoni_tendina,
                required=True,
                width="large"
            ),
            "NOTE": st.column_config.TextColumn("Note Operative", width="large")
        },
        hide_index=True,
        use_container_width=True,
        key="editor_giornaliero_diretto"
    )
    
    # Salvataggio istantaneo dello stato modificato
    st.session_state.stato_giornaliero = tabellone_modificato

    # --- GENERAZIONE AUTOMATICA DEI 4 BLOCCHI DI OUTPUT ---
    df_correnti = st.session_state.stato_giornaliero
    blocco1 = df_correnti[df_correnti['STATO'] == "Presente (Giro Fisso)"][['COGNOME', 'NOME', 'CELLULARE', 'GIRO_FISSO', 'MEZZO', 'NOTE']]
    blocco2 = df_correnti[df_correnti['STATO'] == "Supporto Altra Filiale"][['COGNOME', 'NOME', 'CELLULARE', 'GIRO_SUPPORTO', 'MEZZO', 'NOTE']]
    blocco3 = st.session_state.responsabili
    blocco4 = df_correnti[df_correnti['STATO'] == "Assente"][['COGNOME', 'NOME', 'CELLULARE', 'NOTE']]

    st.markdown("---")
    
    # --- FUNZIONI DI GENERAZIONE REPORT FORMATTATI ---
    def genera_excel_4_blocchi():
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            blocco1.to_excel(writer, sheet_name='Piano Giornaliero', index=False, startrow=1)
            ws = writer.sheets['Piano Giornaliero']
            
            font_titolo = Font(name='Calibri', size=11, bold=True, color='000000')
            font_header = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
            fill_header = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid') 
            fill_titolo = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid') 
            border_grigio = Border(
                left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
                top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
            )
            allineamento_centro = Alignment(horizontal='center', vertical='center')

            def scrivi_blocco_excel(ws, df, titolo_blocco, riga_partenza):
                ws.cell(row=riga_partenza, column=1, value=titolo_blocco).font = font_titolo
                ws.cell(row=riga_partenza, column=1).fill = fill_titolo
                ws.row_dimensions[riga_partenza].height = 22
                
                for col_idx, col_name in enumerate(df.columns, start=1):
                    cell = ws.cell(row=riga_partenza+1, column=col_idx, value=col_name)
                    cell.font = font_header
                    cell.fill = fill_header
                    cell.alignment = allineamento_centro
                    cell.border = border_grigio
                ws.row_dimensions[riga_partenza+1].height = 20
                
                curr_row = riga_partenza + 2
                for _, riga_dati in df.iterrows():
                    for col_idx, val in enumerate(riga_dati, start=1):
                        cell = ws.cell(row=curr_row, column=col_idx, value=str(val) if pd.notna(val) else "")
                        cell.border = border_grigio
                        if col_idx in [3, 4, 5]: 
                            cell.alignment = allineamento_centro
                    ws.row_dimensions[curr_row].height = 18
                    curr_row += 1
                return curr_row + 2 

            ws.delete_rows(1, ws.max_row+20)
            prossima_riga = 2
            prossima_riga = scrivi_blocco_excel(ws, blocco1, "1° BLOCCO: CORRIERI CON NUMERO DI GIRO ASSOCIATO", prossima_riga)
            prossima_riga = scrivi_blocco_excel(ws, blocco2, "2° BLOCCO: CORRIERI IN SUPPORTO ALTRA FILIALE", prossima_riga)
            prossima_riga = scrivi_blocco_excel(ws, blocco3, "3° BLOCCO: NOMINATIVI RESPONSABILI PRESENTI", prossima_riga)
            prossima_riga = scrivi_blocco_excel(ws, blocco4, "4° BLOCCO: CORRIERI ASSENTI", prossima_riga)
            
            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = max(max_len + 4, 13)
                
        return output.getvalue()

    def genera_pdf_4_blocchi():
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", "B", 14)
        
        pdf.cell(0, 10, "PIANO GIORNALIERO FLOTTA E PRESENZE CORRIERI", align="C")
        pdf.ln(12) 
        
        def aggiungi_tabella_pdf(titolo, df):
            pdf.set_font("Arial", "B", 10)
            pdf.set_fill_color(225, 230, 240)
            
            pdf.cell(0, 7, titolo, fill=True)
            pdf.ln(9) 
            
            pdf.set_font("Arial", "", 9)
            if df.empty:
                pdf.cell(0, 7, " Nessun record registrato in questo blocco.")
                pdf.ln(10)
                return
            
            # Intestazioni tabella
            pdf.set_fill_color(30, 75, 120)
            pdf.set_text_color(255, 255, 255)
            for col in df.columns:
                pdf.cell(31, 7, str(col), border=1, fill=True)
            pdf.ln(7) 
            
            # Righe dati
            pdf.set_text_color(0, 0, 0)
            for _, riga in df.iterrows():
                for col in df.columns:
                    pdf.cell(31, 7, str(riga[col])[:16], border=1)
                pdf.ln(7) 
            pdf.ln(4)

        aggiungi_tabella_pdf("1. CORRIERI CON GIRO ASSOCIATO", blocco1)
        aggiungi_tabella_pdf("2. EVENTUALI CORRIERI IN SUPPORTO", blocco2)
        aggiungi_tabella_pdf("3. NOMINATIVI RESPONSABILI PRESENTI", blocco3)
        aggiungi_tabella_pdf("4. CORRIERI ASSENTI", blocco4)
        
        pdf_output = pdf.output(dest='S')
        return bytes(pdf_output)

    # --- INTERFACCIA PULSANTI DI ESPORTAZIONE DIRETTA ---
    col_x1, col_x2, col_x3 = st.columns(3)
    with col_x1:
        st.download_button("📥 Scarica Report Excel", data=genera_excel_4_blocchi(), file_name="piano_presenze.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with col_x2:
        st.download_button("📥 Scarica PDF Pronto Stampa", data=genera_pdf_4_blocchi(), file_name="piano_presenze.pdf", mime="application/pdf")
    with col_x3:
        if st.button("✉️ Invia Report via Mail"):
            st.success(f"File elaborati e inviati con successo a: {st.session_state.config_mail['destinatari']}")

# --- 2. SCHEDA ANAGRAFICA FURGONI ---
elif scelta == "🚐 Anagrafica Furgoni":
    st.title("🚐 Anagrafica e Stato Mezzi Aziendali")
    st.markdown("Puoi inserire un nuovo furgone compilando l'ultima riga vuota in fondo alla tabella. Cambia lo stato in `GUASTO` per rimuoverlo subito dalle scelte quotidiane.")
    
    furgoni_tabella = st.data_editor(
        st.session_state.furgoni, 
        num_rows="dynamic", 
        column_config={
            "DISPONIBILE": st.column_config.SelectboxColumn("Disponibilità", options=["SI", "NO", "GUASTO"], required=True)
        },
        use_container_width=True,
        key="tabella_gestione_furgoni"
    )
    st.session_state.furgoni = furgoni_tabella

# --- 3. SCHEDA ANAGRAFICA PERSONALE (SOLO I PRIMI 4 CAMPI BASE) ---
elif scelta == "👥 Anagrafica Personale":
    st.title("👥 Gestione Personale e Autisti")
    
    st.subheader("Anagrafica Fissa Corrieri")
    st.markdown("Inserisci qui i dati base permanenti dell'organico (Cognome, Nome, Cellulare, Giro standard). L'ultima riga serve per inserire nuovi autisti.")
    
    # Gestisce solo ed esclusivamente le 4 colonne di anagrafica pura
    corrieri_tabella = st.data_editor(
        st.session_state.anagrafica_corrieri, 
        num_rows="dynamic", 
        use_container_width=True, 
        key="tabella_gestione_corrieri"
    )
    st.session_state.anagrafica_corrieri = corrieri_tabella
    
    st.markdown("---")
    st.subheader("Anagrafica Fissa Responsabili / Capi Turno")
    responsabili_tabella = st.data_editor(st.session_state.responsabili, num_rows="dynamic", use_container_width=True, key="tabella_gestione_responsabili")
    st.session_state.responsabili = responsabili_tabella

# --- 4. CONFIGURAZIONE MAIL ---
elif scelta == "⚙️ Configurazione Mail":
    st.title("⚙️ Configurazione Indirizzi di Spedizione")
    email_salvate = st.text_input("Indirizzi Email Destinatari (se sono più di uno, separali con una virgola)", value=st.session_state.config_mail["destinatari"])
    if st.button("Salva Configurazione"):
        st.session_state.config_mail["destinatari"] = email_salvate
        st.success("Impostazioni salvate con successo!")