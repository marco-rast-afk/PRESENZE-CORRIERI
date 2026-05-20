# --- 3. SCHEDA ANAGRAFICA PERSONALE ---
elif scelta == "👥 Anagrafica Personale":
    st.title("👥 Gestione Personale e Autisti")
    st.subheader("Anagrafica Fissa Corrieri")
    
    # Funzione di callback per salvare IMMEDIATAMENTE e sincronizzare il tabellone presenze
    def aggiorna_anagrafica_corrieri():
        # Recupera i dati modificati dall'editor tramite la chiave del session_state
        if "tabella_gestione_corrieri" in st.session_state:
            edits = st.session_state["tabella_gestione_corrieri"]
            df_attuale = st.session_state.anagrafica_corrieri.copy()
            
            # 1. Gestione righe modificate
            for row_idx, deltas in edits["edited_rows"].items():
                for col, val in deltas.items():
                    df_attuale.iat[row_idx, df_attuale.columns.get_loc(col)] = val
            
            # 2. Gestione righe eliminate
            if edits["deleted_rows"]:
                df_attuale = df_attuale.drop(edits["deleted_rows"]).reset_index(drop=True)
            
            # 3. Gestione nuove righe aggiunte
            if edits["added_rows"]:
                for new_row in edits["added_rows"]:
                    # Assicurati che le colonne esistano per evitare NaN fastidiosi
                    riga_pulita = {col: new_row.get(col, "") for col in df_attuale.columns}
                    df_attuale = pd.concat([df_attuale, pd.DataFrame([riga_pulita])], ignore_index=True)
            
            # Salva l'anagrafica aggiornata
            st.session_state.anagrafica_corrieri = df_attuale
            
            # Sincronizza FORZATAMENTE lo stato giornaliero per non perdere i nuovi nomi nel tabellone
            df_nuovo = df_attuale.copy()
            df_nuovo = df_nuovo.merge(
                st.session_state.stato_giornaliero[['COGNOME', 'NOME', 'STATO', 'GIRO_SUPPORTO', 'MEZZO', 'NOTE']], 
                on=['COGNOME', 'NOME'], 
                how='left'
            )
            df_nuovo["STATO"] = df_nuovo["STATO"].fillna("Presente (Giro Fisso)")
            df_nuovo["GIRO_SUPPORTO"] = df_nuovo["GIRO_SUPPORTO"].fillna("")
            df_nuovo["MEZZO"] = df_nuovo["MEZZO"].fillna("Nessuno")
            df_nuovo["NOTE"] = df_nuovo["NOTE"].fillna("")
            st.session_state.stato_giornaliero = df_nuovo

    # Visualizza l'editor passando la callback su on_change
    st.data_editor(
        st.session_state.anagrafica_corrieri, 
        num_rows="dynamic", 
        use_container_width=True, 
        key="tabella_gestione_corrieri",
        on_change=aggiorna_anagrafica_corrieri
    )
    
    st.markdown("---")
    st.subheader("Anagrafica Fissa Responsabili / Capi Turno")
    
    # Callback simile per i responsabili
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
