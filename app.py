import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
from datetime import datetime, date
import io

# --- KONFIGURATION ---
st.set_page_config(page_title="Vereins-Cockpit", layout="wide", page_icon="⛪")

# DER LINK ZUR TABELLE (Fest eingebaut)
SHEET_URL = "https://docs.google.com/spreadsheets/d/1zV6UCDkalRRk9auLXYfJb_kMGEwUAJ_lUHxktNkyCOw/edit"

# --- HILFSFUNKTIONEN ---
def load_data(conn):
    # Wir zwingen ihn, diese URL zu nutzen
    df = conn.read(spreadsheet=SHEET_URL, usecols=list(range(8)), ttl=0)
    df = df.dropna(how="all")
    
    # Datums-Konvertierung
    df["Datum"] = pd.to_datetime(df["Datum"], dayfirst=True, errors='coerce')
    
    # Euro-Zeichen und Komma bereinigen
    if not df.empty:
        df["Einnahme"] = df["Einnahme"].astype(str).str.replace('€', '', regex=False).str.replace(',', '.', regex=False)
        df["Ausgabe"] = df["Ausgabe"].astype(str).str.replace('€', '', regex=False).str.replace(',', '.', regex=False)
        df["Einnahme"] = pd.to_numeric(df["Einnahme"], errors='coerce').fillna(0.0)
        df["Ausgabe"] = pd.to_numeric(df["Ausgabe"], errors='coerce').fillna(0.0)
        
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
menu = st.sidebar.radio("Menü", ["📊 Cockpit & Journal", "✍️ Neue Buchung", "💸 Offene Zahlungen", "📄 Dokumente"])

# ==============================================================================
# 1. COCKPIT & JOURNAL
# ==============================================================================
if menu == "📊 Cockpit & Journal":
    st.title("📊 Finanz-Übersicht")
    
    budget = df["Einnahme"].sum() - df["Ausgabe"].sum()
    real_df = df[df["Status"] == "Erledigt"]
    bank_real = real_df["Einnahme"].sum() - real_df["Ausgabe"].sum()
    offen_df = df[(df["Status"] == "Offen") & (df["Ausgabe"] > 0)]
    offen_summe = offen_df["Ausgabe"].sum()

    col1, col2, col3 = st.columns(3)
    col1.metric("💰 Verfügbares Budget", f"{budget:,.2f} €")
    col2.metric("🏦 Kontostand (Real)", f"{bank_real:,.2f} €", delta=f"- {offen_summe:.2f} € offen", delta_color="inverse")
    col3.metric("📄 Offene Rechnungen", f"{len(offen_df)} Stück")
    
    st.markdown("---")
    st.subheader("Buchungsjournal")
    
    display_df = df.copy()
    display_df["Datum"] = display_df["Datum"].dt.strftime("%d.%m.%Y")
    
    st.dataframe(
        display_df.sort_index(ascending=False), 
        use_container_width=True,
        column_config={
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
        konto_in = c1.selectbox("Konto", ["Bank", "Handkassa", "Minikonto"])
        rechnung_in = c2.checkbox("Rechnung vorhanden?", value=True)
        
        status_default = "Offen" if konto_in == "Bank" else "Erledigt"
        status_in = c3.selectbox("Status", ["Offen", "Erledigt"], index=0 if status_default=="Offen" else 1)
        
        submitted = st.form_submit_button("Speichern")
        
        if submitted:
            if not anlass_in:
                st.error("Bitte Anlass angeben!")
            else:
                einnahme_val = betrag_in if typ == "Einnahme" else 0.0
                ausgabe_val = betrag_in if typ == "Ausgabe" else 0.0
                rechnung_txt = "Ja" if rechnung_in else "Nein"
                
                new_entry = pd.DataFrame([{
                    "Datum": datum_in.strftime("%Y-%m-%d"),
                    "Anlass_Person": anlass_in,
                    "Einnahme": einnahme_val,
                    "Ausgabe": ausgabe_val,
                    "Bemerkung": bemerkung_in,
                    "Konto": konto_in,
                    "Rechnung_Vorhanden": rechnung_txt,
                    "Status": status_in
                }])
                
                updated_df = pd.concat([df, new_entry], ignore_index=True)
                
                # AUCH HIER: URL direkt übergeben!
                conn.update(spreadsheet=SHEET_URL, worksheet="Buchungen", data=updated_df)
                
                st.success("Gespeichert! Bitte Seite neu laden (F5).")

# ==============================================================================
# 3. OFFENE ZAHLUNGEN
# ==============================================================================
elif menu == "💸 Offene Zahlungen":
    st.header("💸 Offene Überweisungen")
    
    mask_offen = (df["Status"] == "Offen") & (df["Ausgabe"] > 0)
    todos = df[mask_offen].copy()
    
    if todos.empty:
        st.success("Alles erledigt! 🎉")
    else:
        st.dataframe(todos)
        entry_to_close = st.selectbox("Welchen Eintrag bezahlen?", todos["Anlass_Person"].unique())
        
        if st.button("Als 'Erledigt' markieren"):
            mask = (df["Anlass_Person"] == entry_to_close) & (df["Status"] == "Offen")
            df.loc[mask, "Status"] = "Erledigt"
            
            # AUCH HIER: URL direkt übergeben!
            conn.update(spreadsheet=SHEET_URL, worksheet="Buchungen", data=df)
            st.success("Markiert! Bitte neu laden.")

# ==============================================================================
# 4. DOKUMENTE
# ==============================================================================
# ==============================================================================
# 4. DOKUMENTE
# ==============================================================================
elif menu == "📄 Dokumente":
    from docxtpl import DocxTemplate # Wird nur hier geladen, um Fehler zu vermeiden
    
    st.header("📄 Förderantrags-Generator")
    
    # 1. Eingabefelder (Basierend auf deiner Vorlage)
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
        
        datum_in = st.text_input("Dauer der Veranstaltung (Datum von - bis)", placeholder="z.B. 01.08. - 05.08.2026")
        
        st.subheader("Finanzen & Antragsteller")
        gesamtsumme_in = st.number_input("Gesamtsumme (€)", min_value=1.00, format="%.2f")
        antragsteller_in = st.text_input("Antragsteller*in (Name, Datum)")
        adresse_in = st.text_input("Adresse")
        kontodaten_in = st.text_input("Kontodaten (IBAN)")
        
        submitted = st.form_submit_button("Antrag erstellen")

    if submitted:
        # 2. Daten ins Word-Dokument einfügen (Context)
        context = {
            "bezeichnung": bezeichnung_in,
            "ort": ort_in,
            "gruppe": gruppe_in,
            "pfarrgemeinde": pfarre_in,
            "anzahl_kids": kids_in,
            "anzahl_begleiter": begleiter_in,
            "naechtigungen": naechtigungen_in,
            "datum_von_bis": datum_in,
            "gesamtsumme": f"{gesamtsumme_in:,.2f}", # Formatierung mit Komma
            "antragsteller": antragsteller_in,
            "adresse": adresse_in,
            "kontodaten": kontodaten_in,
            "name_datum": f"{antragsteller_in}, {date.today().strftime('%d.%m.%Y')}"
        }
        
        try:
            doc = DocxTemplate("vorlage_antrag.docx")
            doc.render(context)
            
            # Speichern in Memory Stream für Download
            bio = io.BytesIO()
            doc.save(bio)
            
            st.download_button(
                label="📥 Antrag herunterladen",
                data=bio.getvalue(),
                file_name=f"Foerderantrag_{bezeichnung_in}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            st.success("Dokument generiert und bereit zum Download.")
            
        except FileNotFoundError:
            st.error("Fehler: 'vorlage_antrag.docx' nicht gefunden. Bitte prüfen Sie den Dateinamen auf GitHub.")
        except Exception as e:
            st.error(f"Fehler beim Erstellen des Dokuments. Möglicherweise ein ungültiger Platzhalter. ({e})")
