import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
from datetime import datetime, date
import io
import xlsxwriter 

# --- HILFSFUNKTION: EXCEL EXPORT ---
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Kassenbuch')
    output.seek(0)
    return output.read()

# --- KONFIGURATION ---
st.set_page_config(page_title="Vereins-Cockpit", layout="wide", page_icon="⛪")

# DER LINK ZUR TABELLE
SHEET_URL = "https://docs.google.com/spreadsheets/d/1zV6UCDkalRRk9auLXYfJb_kMGEwUAJ_lUHxktNkyCOw/edit"

# --- HILFSFUNKTIONEN ---
def load_data(conn):
    # Daten laden (alle 10 Spalten)
    df = conn.read(spreadsheet=SHEET_URL, usecols=list(range(10)), ttl=0)
    
    # 1. Zeilen löschen, die komplett leer sind
    df = df.dropna(how="all")
    
    # 2. Datum-Fix: Konvertieren und Uhrzeit entfernen
    if "Datum" in df.columns:
        # Zwingt alles in ein Datumsformat, kaputte Werte werden "NaT"
        df["Datum"] = pd.to_datetime(df["Datum"], errors='coerce')
        # Uhrzeit abschneiden -> Nur noch das reine Datum behalten
        df["Datum"] = df["Datum"].dt.date
    
    # 3. Zahlen-Bereinigung
    if not df.empty:
        for col in ["Einnahme", "Ausgabe"]:
            df[col] = df[col].astype(str).str.replace('€', '', regex=False).str.replace(',', '.', regex=False)
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)
            
    return df

# --- VERBINDUNG ---
conn = st.connection("gsheets", type=GSheetsConnection)
try:
    df = load_data(conn)
except Exception as e:
    st.error(f"Fehler beim Laden: {e}")
    st.stop()

# --- SIDEBAR MENÜ ---
st.sidebar.title("⛪ Hatler Minis")
menu = st.sidebar.radio("Menü", ["📊 Cockpit & Journal", "✍️ Neue Buchung", "💸 Offene Zahlungen", "📈 Projekt-Analyse", "📄 Dokumente", '✅ Kassenprüfung', "🔐 Zugangsdaten"])

# ==============================================================================
# 1. COCKPIT & JOURNAL
# ==============================================================================
if menu == "📊 Cockpit & Journal":
    st.title("📊 Finanz-Übersicht")
    
    budget = df["Einnahme"].sum() - df["Ausgabe"].sum()
    real_df = df[df["Status"] == "Erledigt"]
    bank_real = real_df["Einnahme"].sum() - real_df["Ausgabe"].sum()
    
    offen_df = df[(df["Status"] == "Offen") & ((df["Ausgabe"] > 0) | (df["Einnahme"] > 0))]
    offen_summe = offen_df["Ausgabe"].sum()

    col1, col2, col3 = st.columns(3)
    col1.metric("💰 Verfügbares Budget", f"{budget:,.2f} €")
    col2.metric("🏦 Kontostand (N26)", f"{bank_real:,.2f} €", delta=f"- {offen_summe:.2f} € offen", delta_color="inverse")
    col3.metric(label="📄 Offene Posten", value=f"{len(offen_df)} Stück")
    
    st.markdown("---")
    st.subheader("Buchungsjournal")
    
    # Sortieren: Neueste oben
    display_df = df.sort_values(by="Datum", ascending=False)
    
    st.dataframe(
        display_df, 
        use_container_width=True,
        column_config={
            "Datum": st.column_config.DateColumn("Datum", format="DD.MM.YYYY"),
            "Einnahme": st.column_config.NumberColumn(format="%.2f €"),
            "Ausgabe": st.column_config.NumberColumn(format="%.2f €"),
        }
    )

# ==============================================================================
# 2. NEUE BUCHUNG
# ==============================================================================
elif menu == "✍️ Neue Buchung":
    st.header("✍️ Neuen Eintrag erfassen")
    
    with st.form("entry_form"):
        col_a, col_b = st.columns(2)
        datum_in = col_a.date_input("Datum", date.today())
        anlass_in = col_b.text_input("Anlass / Person")
        
        typ = st.radio("Buchungstyp", ["Ausgabe", "Einnahme"], horizontal=True)
        betrag_in = st.number_input("Betrag (€)", min_value=0.01, format="%.2f")
        bemerkung_in = st.text_input("Bemerkung")
        
        c1, c2, c3 = st.columns(3)
        konto_in = c1.selectbox("Konto", ["Minikonto", "Handkassa"])
        rechnung_in = c2.checkbox("Rechnung vorhanden?", value=True)
        
        status_default = "Offen" if konto_in == "Minikonto" else "Erledigt"
        status_in = c3.selectbox("Status", ["Offen", "Erledigt"], index=0 if status_default=="Offen" else 1)
        
        submitted = st.form_submit_button("Speichern")
        
        if submitted:
            if not anlass_in:
                st.error("Bitte Anlass angeben!")
            else:
                einnahme_val = betrag_in if typ == "Einnahme" else 0.0
                ausgabe_val = betrag_in if typ == "Ausgabe" else 0.0
                rechnung_txt = "Ja" if rechnung_in else "Nein"
                
                # 1. Den neuen Datensatz als kleines DataFrame erstellen
                # Die Spaltennamen müssen EXAKT wie im Google Sheet sein
                new_entry = pd.DataFrame([{
                    "Datum": datum_in, # Python-Datumsobjekt
                    "Anlass_Person": anlass_in,
                    "Einnahme": einnahme_val,
                    "Ausgabe": ausgabe_val,
                    "Bemerkung": bemerkung_in,
                    "Konto": konto_in,
                    "Rechnung_Vorhanden": rechnung_txt,
                    "Status": status_in,
                    "Pruefung_OK": "",
                    "Pruefung_Bemerkung": ""
                }])
                
                # 2. Den neuen Eintrag an das bestehende 'df' (das wir oben geladen haben) hängen
                updated_df = pd.concat([df, new_entry], ignore_index=True)
                
                # 3. Das komplette, aktualisierte DataFrame zurück zu Google schicken
                conn.update(spreadsheet=SHEET_URL, worksheet="Buchungen", data=updated_df)
                
                st.success("Gespeichert! Die Seite wird neu geladen...")
                st.rerun()

# ==============================================================================
# 3. OFFENE ZAHLUNGEN
# ==============================================================================
elif menu == "💸 Offene Zahlungen":
    st.header("💸 Offene Überweisungen")
    
    mask_offen = (df["Status"] == "Offen") & ((df["Ausgabe"] > 0) | (df["Einnahme"] > 0))
    todos = df[mask_offen].copy()
    
    if todos.empty:
        st.success("Alles erledigt! 🎉")
    else:
        st.dataframe(
            todos,
            use_container_width=True,
            column_config={
                "Datum": st.column_config.DateColumn(format="DD.MM.YYYY"), 
                "Einnahme": st.column_config.NumberColumn(format="%.2f €"),
                "Ausgabe": st.column_config.NumberColumn(format="%.2f €"),
            }
        )
        entry_to_close = st.selectbox("Welchen Eintrag bezahlen?", todos["Anlass_Person"].unique())
        
        if st.button("Als 'Erledigt' markieren"):
            mask = (df["Anlass_Person"] == entry_to_close) & (df["Status"] == "Offen")
            df.loc[mask, "Status"] = "Erledigt"
            conn.update(spreadsheet=SHEET_URL, worksheet="Buchungen", data=df)
            st.success("Markiert!")
            st.rerun()

# ==============================================================================
# 4. DOKUMENTE
# ==============================================================================
elif menu == "📄 Dokumente":
    from docxtpl import DocxTemplate 
    st.header("📄 Förderantrags-Generator")
    
    with st.form("antrags_form"):
        st.subheader("Details zur Veranstaltung")
        c1, c2 = st.columns(2)
        bezeichnung_in = c1.text_input("Bezeichnung des Projekts")
        ort_in = c2.text_input("Ort der Veranstaltung")
        
        c3, c4 = st.columns(2)
        gruppe_in = c3.selectbox("Gruppe", ["Jungschar", "Minis", "KJ", "Firmlinge"])
        pfarre_in = c4.text_input("Pfarrgemeinde")

        c5, c6, c7 = st.columns(3)
        kids_in = c5.number_input("Anzahl Kinder/Jugendliche", min_value=0)
        begleiter_in = c6.number_input("Anzahl Begleitpersonen", min_value=0)
        naechtigungen_in = c7.number_input("Anzahl Nächtigungen (Lager)", min_value=0)
        
        datum_range = st.text_input("Dauer (Datum von - bis)", placeholder="z.B. 01.08. - 05.08.2026")
        
        st.subheader("Finanzen & Antragsteller")
        gesamtsumme_in = st.number_input("Gesamtsumme (€)", min_value=1.00, format="%.2f")
        antragsteller_in = st.text_input("Antragsteller*in (Name)")
        adresse_in = st.text_input("Adresse")
        kontodaten_in = st.text_input("Kontodaten (IBAN)")
        
        submitted = st.form_submit_button("Antrag erstellen")

    if submitted:
        context = {
            "bezeichnung": bezeichnung_in,
            "ort": ort_in,
            "gruppe": gruppe_in,
            "pfarrgemeinde": pfarre_in,
            "anzahl_kids": kids_in,
            "anzahl_begleiter": begleiter_in,
            "naechtigungen": naechtigungen_in,
            "datum_von_bis": datum_range,
            "gesamtsumme": f"{gesamtsumme_in:,.2f}",
            "antragsteller": antragsteller_in,
            "adresse": adresse_in,
            "kontodaten": kontodaten_in,
            "name_datum": f"{antragsteller_in}, {date.today().strftime('%d.%m.%Y')}"
        }
        
        try:
            doc = DocxTemplate("vorlage_antrag.docx")
            doc.render(context)
            bio = io.BytesIO()
            doc.save(bio)
            st.download_button(
                label="📥 Antrag herunterladen",
                data=bio.getvalue(),
                file_name=f"Foerderantrag_{bezeichnung_in}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception as e:
            st.error(f"Fehler: {e}")

# ==============================================================================
# 5. KASSENPRÜFUNG
# ==============================================================================
elif menu == '✅ Kassenprüfung':
    st.header("✅ Kassenprüfung")

    if 'Pruefung_OK' not in df.columns:
        st.error("Spalte 'Pruefung_OK' fehlt im Sheet!")
    else:
        unverified_df = df[df['Pruefung_OK'].isnull() | (df['Pruefung_OK'] == "")]
        
        if unverified_df.empty:
            st.success("🎉 Alles erledigt!")
            excel_data = to_excel(df)
            st.download_button(label="⬇️ Excel Export", data=excel_data, file_name=f'Bericht_{date.today()}.xlsx')
        else:
            current_idx = unverified_df.index[0]
            row = df.loc[current_idx]
            
            with st.container(border=True):
                st.subheader(f"📅 {row['Datum'].strftime('%d.%m.%Y') if pd.notnull(row['Datum']) else '---'}")
                st.write(f"**Anlass:** {row['Anlass_Person']}")
                st.metric("Betrag", f"{row['Einnahme'] if row['Einnahme'] > 0 else row['Ausgabe']:.2f} €")
            
            with st.form("audit_form"):
                status = st.radio("Prüfung:", ["OK ✅", "Fehler ❌", "Überspringen ⏭️"], horizontal=True)
                bemerkung = st.text_input("Notiz")
                if st.form_submit_button("Speichern"):
                    if status != "Überspringen ⏭️":
                        df.loc[current_idx, 'Pruefung_OK'] = status
                        df.loc[current_idx, 'Pruefung_Bemerkung'] = bemerkung
                        conn.update(spreadsheet=SHEET_URL, worksheet="Buchungen", data=df)
                        st.rerun()

# ==============================================================================
# 6. PROJEKT ANALYSE
# ==============================================================================
elif menu == "📈 Projekt-Analyse":
    st.header("📈 Projekt-Check")
    search_term = st.text_input("Stichwort (z.B. Lager)")

    if search_term:
        mask = (df['Anlass_Person'].astype(str).str.contains(search_term, case=False)) | \
               (df['Bemerkung'].astype(str).str.contains(search_term, case=False))
        project_df = df[mask]

        if not project_df.empty:
            ergebnis = project_df['Einnahme'].sum() - project_df['Ausgabe'].sum()
            st.metric("Gewinn / Verlust", f"{ergebnis:,.2f} €", delta=f"{ergebnis:.2f} €")
            st.dataframe(project_df, column_config={"Datum": st.column_config.DateColumn(format="DD.MM.YYYY")})

# ==============================================================================
# 7. ZUGANGSDATEN
# ==============================================================================
elif menu == "🔐 Zugangsdaten":
    st.header("🔐 Geschützter Bereich")
    if "authenticated" not in st.session_state: st.session_state.authenticated = False

    if not st.session_state.authenticated:
        password = st.text_input("Passwort", type="password")
        if st.button("Einloggen"):
            if password == st.secrets["credentials"]["admin_password"]:
                st.session_state.authenticated = True
                st.rerun()
    else:
        st.write(f"**IBAN:** {st.secrets['credentials']['bank_iban']}")
        if st.button("Ausloggen"):
            st.session_state.authenticated = False
            st.rerun()
