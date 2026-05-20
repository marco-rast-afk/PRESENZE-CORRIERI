import streamlit as st
import pandas as pd
from fpdf import FPDF
import io
from datetime import datetime
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Gestione Presenze Corrieri", layout="wide")

# Dizionario per convertire il numero del mese nel nome in italiano maiuscolo
MESI_ITA = {
    1: "GENNAIO", 2: "FEBBRAIO", 3: "MARZO", 4: "APRILE", 5: "MAGGIO", 6: "GIUGNO",
    7: "LUGLIO", 8: "AGOSTO", 9: "SETTEMBRE", 10: "OTTOBRE", 11: "NOVEMBRE", 12: "DICEMBRE"
}

# --- DATABASE INIZIALE (SESSION STATE) ---
if 'furgoni' not in st.session_state:
    st.session_state.furgoni = pd.DataFrame([
        {"MARCA": "Fiat", "MODELLO": "Ducato", "TIPO": "Furgone", "TARGA": "GR256RF", "DISPONIBILE": "SI"},
        {"MARCA": "Ford", "MODELLO": "Transit", "TIPO": "Furgone", "TARGA": "GS557WM", "DISPONIBILE": "SI"},
        {"MARCA": "Iveco", "MODELLO": "Daily", "TIPO": "Furgone", "TARGA": "HB683CE", "DISPONIBILE": "NO"},
        {"MARCA": "Mercedes", "MODELLO": "Sprinter", "TIPO": "Furgone", "TARGA": "HB787TS", "DISPONIBILE": "GUASTO"},
    ])

# ANAGRAFICA PERSONALE AUTOMATIZZATA CON TUTTI I TUOI CORRIERI
if 'anagrafica_corrieri' not in st.session_state:
    st.session_state.anagrafica_corrieri = pd.DataFrame([
        {"COGNOME": "CROCI", "NOME": "MARINO", "CELLULARE": "3314509080", "GIRO_FISSO": "1"},
        {"COGNOME": "VINCIGUERRA", "NOME": "ANGELO", "CELLULARE": "3351696753", "GIRO_FISSO": "2"},
        {"COGNOME": "D ANGELO", "NOME": "SALVATORE", "CELLULARE": "3881132883", "GIRO_FISSO": "3"},
        {"COGNOME": "MARCIANO", "NOME": "ANTONIO", "CELLULARE": "3489292359", "GIRO_FISSO": "4"},
        {"COGNOME": "CAPUTO", "NOME": "OVIDIO", "CELLULARE": "3385277033", "GIRO_FISSO": "5"},
        {"COGNOME": "AINIS", "NOME": "CIRO", "CELLULARE": "3891618386", "GIRO_FISSO": "6"},
        {"COGNOME": "CAPPELLA", "NOME": "GIUSEPPE", "CELLULARE": "3427598835", "GIRO_FISSO": "7"},
        {"COGNOME": "BEN CHATTI", "NOME": "MARTIN", "CELLULARE": "3451303466", "GIRO_FISSO": "8"},
        {"COGNOME": "POMILIA", "NOME": "PIETRO", "CELLULARE": "3382107031", "GIRO_FISSO": "9"},
        {"COGNOME": "CAPUTO", "NOME": "GIANLUCA", "CELLULARE": "3391178350", "GIRO_FISSO": "10"},
        {"COGNOME": "CAIONI", "NOME": "MANUEL", "CELLULARE": "3703349078", "GIRO_FISSO": "11"},
        {"COGNOME": "MANZO", "NOME": "GIAMPAOLO", "CELLULARE": "3384187899", "GIRO_FISSO": "12"},
        {"COGNOME": "XHANI", "NOME": "ORGES", "CELLULARE": "3515960708", "GIRO_FISSO": "13"},
        {"COGNOME": "CISSE", "NOME": "THIERNO", "CELLULARE": "3513434477", "GIRO_FISSO": "14"},
        {"COGNOME": "CAPPONI", "NOME": "MASSIMILIANO", "CELLULARE": "3206274943", "GIRO_FISSO": "15"},
        {"COGNOME": "MERCURI", "NOME": "SIMONE", "CELLULARE": "3483624304", "GIRO_FISSO": "16"},
        {"COGNOME": "FELICIANI", "NOME": "SAMUELE", "CELLULARE": "3459494393", "GIRO_FISSO": "17"},
        {"COGNOME": "CALUGI", "NOME": "ALESSANDRO", "CELLULARE": "3467894592", "GIRO_FISSO": "18"},
        {"COGNOME": "D'ANTONIO", "NOME": "GIANFRANCO", "CELLULARE": "3791766322", "GIRO_FISSO": "19"},
        {"COGNOME": "SPACCASASSI", "NOME": "MARIO", "CELLULARE": "3298424239", "GIRO_FISSO": "20"},
        {"COGNOME": "MACRILLANTE", "NOME": "LORENZO", "CELLULARE": "3396085295", "GIRO_FISSO": "21"},
        {"COGNOME": "CAPPELLI", "NOME": "SILVIO", "CELLULARE": "3314578329", "GIRO_FISSO": "22"},
        {"COGNOME": "SQUILLACE", "NOME": "MICHELE", "CELLULARE": "3404966459", "GIRO_FISSO": "23"},
        {"COGNOME": "CAPRIOTTI", "NOME": "CRISTIANO", "CELLULARE": "3452137376", "GIRO_FISSO": "24"},
        {"COGNOME": "VINCIGUERRA", "NOME": "VINCENZO", "CELLULARE": "3348361930", "GIRO_FISSO": "25"},
        {"COGNOME": "AMANDONICO", "NOME": "DARIO", "CELLULARE": "3477784701", "GIRO_FISSO": "26"},
        {"COGNOME": "DI PIETRO", "NOME": "SIMONE", "CELLULARE": "3484587690", "GIRO_FISSO": "27"},
        {"COGNOME": "LUCIANI", "NOME": "GIOVANNI", "CELLULARE": "3282063754", "GIRO_FISSO": "28"},
        {"COGNOME": "DE SIMONE", "NOME": "ANTONIO", "CELLULARE": "3472316522", "GIRO_FISSO": "29"},
        {"COGNOME": "MARCONI", "NOME": "ANDREA", "CELLULARE": "3381092067", "GIRO_FISSO": "30"},
        {"COGNOME": "PETROCCHI", "NOME": "ALESSANDRO", "CELLULARE": "3204269922", "GIRO_FISSO": "31"},
        {"COGNOME": "MANAZZA", "NOME": "GIUSEPPE", "CELLULARE": "3926854314", "GIRO_FISSO": "32"},
        {"COGNOME": "SCARPA", "NOME": "CRESCENZO", "CELLULARE": "3494993238", "GIRO_FISSO": "33"},
        {"COGNOME": "FEBO", "NOME": "DANIELE", "CELLULARE": "3297074905", "GIRO_FISSO": "34"},
        {"COGNOME": "DIENG", "NOME": "MAYPE", "CELLULARE": "3203619657", "GIRO_FISSO": "35"},
        {"COGNOME": "MARCONI", "NOME": "ALESSIO", "CELLULARE": "3931574397", "GIRO_FISSO": "36"},
        {"COGNOME": "CORCIONE", "NOME": "GIUSEPPE", "CELLULARE": "3203437844", "GIRO_FISSO": "37"},
        {"COGNOME": "PAVAN", "NOME": "ALESSANDRO", "CELLULARE": "3515953354", "GIRO_FISSO": "38"},
        {"COGNOME": "CAMUSO", "NOME": "LUCA", "CELLULARE": "3511244208", "GIRO_FISSO": "39"},
        {"COGNOME": "NICCOLINI", "NOME": "ANDREA", "CELLULARE": "3202732824", "GIRO_FISSO": "40"},
        {"COGNOME": "VINOTTI", "NOME": "SONNY", "CELLULARE": "3313800865", "GIRO_FISSO": "41"},
        {"COGNOME": "SCIARRA", "NOME": "ALESSANDRO", "CELLULARE": "3515006439", "GIRO_FISSO": "42"},
        {"COGNOME": "ENOW", "NOME": "SOLOMON", "CELLULARE": "3930180387", "GIRO_FISSO": "43"},
        {"COGNOME": "GIACOMELLI", "NOME": "PAOLO", "CELLULARE": "3298884097", "GIRO_FISSO": "201"},
        {"COGNOME": "SPALAZZI", "NOME": "BENEDETTA", "CELLULARE": "3341971927", "GIRO_FISSO": "202"},
        {"COGNOME": "FERRAMINI", "NOME": "PANCRAZIO", "CELLULARE": "3408660269", "GIRO_FISSO": "203"},
        {"COGNOME": "OREFICE", "NOME": "SALVAVORE", "CELLULARE": "3899012716", "GIRO_FISSO": "204"},
        {"COGNOME": "SALVI", "NOME": "GIANLUCA", "CELLULARE": "3289122867", "GIRO_FISSO": "205"},
        {"COGNOME": "GIULIANI", "NOME": "LUIGI", "CELLULARE": "3318356858", "GIRO_FISSO": "206"},
        {"COGNOME": "TOMMARELLI", "NOME": "ALFONSO", "CELLULARE": "3478548756", "GIRO_FISSO": "207"},
        {"COGNOME": "NISTOR", "NOME": "BOGDAN", "CELLULARE": "3293436637", "GIRO_FISSO": "208"},
        {"COGNOME": "DE MINICIS", "NOME": "GIULIANO", "CELLULARE": "3932174950", "GIRO_FISSO": "209"},
        {"COGNOME": "PERILLO", "NOME": "VINCENZO", "CELLULARE": "3349059267", "GIRO_FISSO": "210"},
        {"COGNOME": "BORGIA", "NOME": "FRANCESCO", "CELLULARE": "3280347303", "GIRO_FISSO": "211"},
        {"COGNOME": "VESPERINI", "NOME": "ROBERTO", "CELLULARE": "3203215836", "GIRO_FISSO": "212"},
        {"COGNOME": "PROTASI", "NOME": "FILIPPO", "CELLULARE": "3490626638", "GIRO_FISSO": "213"},
        {"COGNOME": "DE MATTEIS", "NOME": "ALESSIA", "CELLULARE": "3490566977", "GIRO_FISSO": "214"},
        {"COGNOME": "PELLE", "NOME": "EUGENIO", "CELLULARE": "3272314445", "GIRO_FISSO": "215"},
        {"COGNOME": "COSTANTINI", "NOME": "FABRIZIO", "CELLULARE": "3355869387", "GIRO_FISSO": "216"},
        {"COGNOME": "MASCITTI", "NOME": "CLAUDIO", "CELLULARE": "3404501899", "GIRO_FISSO": "217"},
        {"COGNOME": "TRANQUILLI", "NOME": "SIMONE", "CELLULARE": "3807991216", "GIRO_FISSO": "218"},
        {"COGNOME": "MARUSCO", "NOME": "DAVIDE", "CELLULARE": "3296488555", "GIRO_FISSO": "219"},
        {"COGNOME": "MASSICCI", "NOME": "EMILIE", "CELLULARE": "3896382624", "GIRO_FISSO": "220"},
        {"COGNOME": "BOFFINI", "NOME": "VALERIO", "CELLULARE": "3703393921", "GIRO_FISSO": "221"},
        {"COGNOME": "VINCIGUERRA", "NOME": "ROSARIO", "CELLULARE": "3888818725", "GIRO_FISSO": "222"},
        {"COGNOME": "TRANQUILLI", "NOME": "JESSICA", "CELLULARE": "3466767960", "GIRO_FISSO": "223"},
        {"COGNOME": "GERMINI", "NOME": "MASSIMO", "CELLULARE": "3381174017", "GIRO_FISSO": "224"},
        {"COGNOME": "FELICIANI", "NOME": "MATTEO", "CELLULARE": "3459494394", "GIRO_FISSO": "225"},
        {"COGNOME": "PERSICO", "NOME": "GIADA", "CELLULARE": "3500750399", "GIRO_FISSO": "226"},
        {"COGNOME": "MATJA", "NOME": "KLAUDIO", "CELLULARE": "3291692911", "GIRO_FISSO": "227"},
        {"COGNOME": "ROTUNNO", "NOME": "DARIO", "CELLULARE": "3387710023", "GIRO_FISSO": "228"},
        {"COGNOME": "LANZA", "NOME": "CHRISTIAN", "CELLULARE": "3881680533", "GIRO_FISSO": "230"},
        {"COGNOME": "RUZZINI", "NOME": "ROBERTO", "CELLULARE": "3283797115", "GIRO_FISSO": "501"},
        {"COGNOME": "STANGONI", "NOME": "MARCO", "CELLULARE": "3319145878", "GIRO_FISSO": "503"},
        {"COGNOME": "ECHEZURIA", "NOME": "CARLOS", "CELLULARE": "3317412770", "GIRO_FISSO": "508"},
        {"COGNOME": "ROSCIOLI", "NOME": "PIERFRANCESCO", "CELLULARE": "3454280718", "GIRO_FISSO": "510"},
        {"COGNOME": "SPAGNOLI", "NOME": "ANDREA", "CELLULARE": "3756760553", "GIRO_FISSO": "511"},
        {"COGNOME": "MEDICO", "NOME": "PIERPAOLO", "CELLULARE": "3298537578", "GIRO_FISSO": "513"},
        {"COGNOME": "MUHAMMAD", "NOME": "YOUNAS", "CELLULARE": "3463618752", "GIRO_FISSO": "514"},
        {"COGNOME": "IAVARONE", "NOME": "GIUSTINO", "CELLULARE": "3488701779", "GIRO_FISSO": "515"},
        {"COGNOME": "CATALANO", "NOME": "SILVIO", "CELLULARE": "3534763452", "GIRO_FISSO": "517"},
        {"COGNOME": "DI GIROLAMO", "NOME": "ALESSANDRO", "CELLULARE": "3802543375", "GIRO_FISSO": "518"},
    ])

if 'responsabili' not in st.session_state:
    st.session_state.responsabili = pd.DataFrame([
        {"COGNOME": "ROSSI", "NOME": "LUIGI", "RUOLO": "Responsabile Logistica"},
        {"COGNOME": "VERDI", "NOME": "MARCO", "RUOLO": "Supervisore di Turno"}
    ])

# Struttura permanente quotidiana (mantiene i dati salvati)
if 'stato_giornaliero' not in st.session_state:
    df_giorno = st.session_state.anagrafica_corrieri.copy()
    df_giorno["STATO"] = "Presente (Giro Fisso)"
    df_giorno["GIRO_SUPPORTO"] = ""
    df_giorno["MEZZO"] = "Nessuno"
    df_giorno["NOTE"] = ""
    st.session_state.stato_giornaliero = df_giorno

# Sincronizzazione dinamica del tabellone in caso di modifiche in anagrafica fissa
if len(st.session_state.stato_giornaliero) != len(st.session_state.anagrafica_corrieri):
    df_nuovo = st.session_state.anagrafica_corrieri.copy()
    df_nuovo = df_nuovo.merge(st.session_state.stato_giornaliero[['COGNOME', 'NOME', 'STATO', 'GIRO_SUPPORTO', 'MEZZO', 'NOTE']], on=['COGNOME', 'NOME'], how='left')
    df_nuovo["STATO"] = df_nuovo["STATO"].fillna("Presente (Giro Fisso)")
    df_nuovo["GIRO_SUPPORTO"] = df_nuovo["GIRO_SUPPORTO"].fillna("")
    df_nuovo["MEZZO"] = df_nuovo["MEZZO"].fillna("Nessuno")
    df_nuovo["NOTE"] = df_nuovo["NOTE"].fillna("")
    st.session_state.stato_giornaliero = df_nuovo

# --- MENU DI NAVIGAZIONE A SINISTRA ---
menu = ["📋 Tabellone Presenze", "🚐 Anagrafica Furgoni", "👥 Anagrafica Personale"]
scelta = st.sidebar.selectbox("Navigazione", menu)

# --- 1. TABELLONE PRESENZE SNELLO ---
if scelta == "📋 Tabellone Presenze":
    st.title("📋 Inserimento Presenze e Assegnazione Mezzi")
    
    # Selezione della Data di Lavorazione
    data_lavorazione = st.date_input("Data di lavorazione del Piano Presenze", datetime.today())
    
    giorno = data_lavorazione.day
    anno = data_lavorazione.year
    mese_testo = MESI_ITA[data_lavorazione.month]
    data_formato_personalizzato = f"{giorno} {mese_testo} {anno}"

    st.info("💡 I dati rimangono memorizzati dal giorno precedente. Se riassegni un furgone già occupato, il sistema lo rimuoverà automaticamente dal vecchio giro per evitare duplicati.")
    
    furgoni_attivi = st.session_state.furgoni[st.session_state.furgoni['DISPONIBILE'] != "GUASTO"]
    elenco_furgoni_tendina = ["Nessuno"] + (furgoni_attivi['MARCA'] + " " + furgoni_attivi['MODELLO'] + " [" + furgoni_attivi['TARGA'] + "]").tolist()
    
    # Visualizzazione dell'editor
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

    # --- LOGICA DI CONTROLLO RILEVAMENTO DUPLICATI IN TEMPO REALE ---
    # Confrontiamo lo stato vecchio con quello appena modificato per trovare quale riga è cambiata
    if not tabellone_modificato.equals(st.session_state.stato_giornaliero):
        for idx, row in tabellone_modificato.iterrows():
            vecchio_mezzo = st.session_state.stato_giornaliero.at[idx, "MEZZO"]
            nuovo_mezzo = row["MEZZO"]
            
            # Se l'utente ha modificato il furgone assegnandone uno reale (diverso da "Nessuno")
            if nuovo_mezzo != vecchio_mezzo and nuovo_mezzo != "Nessuno":
                # Cerchiamo se questo mezzo è presente in ALTRE righe oltre a quella modificata
                for alt_idx, alt_row in tabellone_modificato.iterrows():
                    if alt_idx != idx and alt_row["MEZZO"] == nuovo_mezzo:
                        # Rilevato duplicato! Lo impostiamo a "Nessuno" nella vecchia riga
                        tabellone_modificato.at[alt_idx, "MEZZO"] = "Nessuno"
                        st.toast(f"⚠️ Mezzo {nuovo_mezzo} rimosso automaticamente dal Giro {alt_row['GIRO_FISSO']} (Riassegnato a Giro {row['GIRO_FISSO']})")
        
        # Salviamo lo stato aggiornato e pulito dai duplicati
        st.session_state.stato_giornaliero = tabellone_modificato
        st.rerun()

    # --- GENERAZIONE AUTOMATICA DEI 4 BLOCCHI DI OUTPUT ---
    df_correnti = st.session_state.stato_giornaliero
    blocco1 = df_correnti[df_correnti['STATO'] == "Presente (Giro Fisso)"][['COGNOME', 'NOME', 'CELLULARE', 'GIRO_FISSO', 'MEZZO', 'NOTE']]
    blocco2 = df_correnti[df_correnti['STATO'] == "Supporto Altra Filiale"][['COGNOME', 'NOME', 'CELLULARE', 'GIRO_SUPPORTO', 'MEZZO', 'NOTE']]
    blocco3 = st.session_state.responsabili
    blocco4 = df_correnti[df_correnti['STATO'] == "Assente"][['COGNOME', 'NOME', 'CELLULARE', 'NOTE']]

    st.markdown("---")
    
    # --- GENERAZIONE DEI FLUSSI PER I FILE ---
    def genera_excel_4_blocchi(data_label):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            blocco1.to_excel(writer, sheet_name='Piano Giornaliero', index=False, startrow=2)
            ws = writer.sheets['Piano Giornaliero']
            
            font_data_top = Font(name='Calibri', size=11, bold=True, color='000000')
            ws.cell(row=1, column=2, value="PRESENZE CORRIERI").font = font_data_top
            ws.cell(row=1, column=3, value=f"DATA: {data_label}").font = font_data_top
            ws.cell(row=1, column=4, value="Filiale SDA PORTO D'ASCOLI").font = font_data_top
            
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

            ws.delete_rows(3, ws.max_row+20)
            prossima_riga = 3
            prossima_riga = scrivi_blocco_excel(ws, blocco1, "1° BLOCCO: CORRIERI CON NUMERO DI GIRO ASSOCIATO", prossima_riga)
            prossima_riga = scrivi_blocco_excel(ws, blocco2, "2° BLOCCO: CORRIERI IN SUPPORTO ALTRA FILIALE", prossima_riga)
            prossima_riga = scrivi_blocco_excel(ws, blocco3, "3° BLOCCO: NOMINATIVI RESPONSABILI PRESENTI", prossima_riga)
            prossima_riga = scrivi_blocco_excel(ws, blocco4, "4° BLOCCO: CORRIERI ASSENTI", prossima_riga)
            
            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = max(max_len + 4, 13)
        return output.getvalue()

    def genera_pdf_4_blocchi(data_label):
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", "B", 14)
        pdf.cell(0, 10, f"PIANO GIORNALIERO FLOTTA E PRESENZE CORRIERI", align="C")
        pdf.ln(6)
        pdf.set_font("Arial", "I", 11)
        pdf.cell(0, 10, f"Data di lavorazione: {data_label}", align="C")
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
            
        aggiungi_tabella_pdf("1. CORRIERI CON GIRO ASSOCIATO", blocco1)
        aggiungi_tabella_pdf("2. EVENTUALI CORRIERI IN SUPPORTO", blocco2)
        aggiungi_tabella_pdf("3. NOMINATIVI RESPONSABILI PRESENTI", blocco3)
        aggiungi_tabella_pdf("4. CORRIERI ASSENTI", blocco4)
        return bytes(pdf.output(dest='S'))

    excel_data = genera_excel_4_blocchi(data_formato_personalizzato)
    pdf_data = genera_pdf_4_blocchi(data_formato_personalizzato)

    nome_file_excel = f"Presenze {data_formato_personalizzato}.xlsx"
    nome_file_pdf = f"Presenze {data_formato_personalizzato}.pdf"

    # --- INTERFACCIA PULSANTI DI ESPORTAZIONE ---
    col_x1, col_x2 = st.columns(2)
    with col_x1:
        st.download_button(
            label=f"📥 Scarica Excel ({nome_file_excel})", 
            data=excel_data, 
            file_name=nome_file_excel, 
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    with col_x2:
        st.download_button(
            label=f"📥 Scarica PDF ({nome_file_pdf})", 
            data=pdf_data, 
            file_name=nome_file_pdf, 
            mime="application/pdf",
            use_container_width=True
        )

# --- 2. SCHEDA ANAGRAFICA FURGONI ---
elif scelta == "🚐 Anagrafica Furgoni":
    st.title("🚐 Anagrafica e Stato Mezzi Aziendali")
    furgoni_tabella = st.data_editor(
        st.session_state.furgoni, 
        num_rows="dynamic", 
        column_config={"DISPONIBILE": st.column_config.SelectboxColumn("Disponibilità", options=["SI", "NO", "GUASTO"], required=True)},
        use_container_width=True,
        key="tabella_gestione_furgoni"
    )
    st.session_state.furgoni = furgoni_tabella

# --- 3. SCHEDA ANAGRAFICA PERSONALE ---
elif scelta == "👥 Anagrafica Personale":
    st.title("👥 Gestione Personale e Autisti")
    st.subheader("Anagrafica Fissa Corrieri")
    corrieri_tabella = st.data_editor(st.session_state.anagrafica_corrieri, num_rows="dynamic", use_container_width=True, key="tabella_gestione_corrieri")
    st.session_state.anagrafica_corrieri = corrieri_tabella
    
    st.markdown("---")
    st.subheader("Anagrafica Fissa Responsabili / Capi Turno")
    responsabili_tabella = st.data_editor(st.session_state.responsabili, num_rows="dynamic", use_container_width=True, key="tabella_gestione_responsabili")
    st.session_state.responsabili = responsabili_tabella