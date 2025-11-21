import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
from datetime import datetime, date
import io

# --- KONFIGURATION ---
st.set_page_config(page_title="Vereins-Cockpit", layout="wide", page_icon="⛪")

# --- HILFSFUNKTIONEN ---
def load_data(conn):
    # Lädt Daten (ohne expliziten Worksheet-Namen, um Fehler zu vermeiden)
    df = conn.read(usecols=list(range(8)), ttl=0)
    df = df.dropna(how="all")
    
    # Datums-Konvertierung
    df["Datum"] = pd.to_datetime(df["Datum"], dayfirst=True, errors='coerce')
    
    # Euro-Zeichen und Komma bereinigen (WICHTIGER FIX)
    if not df.empty:
        # Wir wandeln alles in String um, entfernen '€' und tauschen Komma gegen Punkt
        df["Einnahme"] = df["Einnahme"].astype(str).str.replace('€', '', regex=False).str.replace(',', '.', regex=False)
        df["Ausgabe"] = df["Ausgabe"].astype(str).str.replace('€', '', regex=False).str.replace(',', '.', regex=False)
        
        # Jetzt in Zahlen umwandeln (Erzwingen)
        df["Einnahme"] = pd.to_numeric(df["Einnahme"], errors='coerce').fillna(0.0)
        df["Ausgabe"] = pd.to_numeric(df["Ausgabe"], errors='coerce').fillna(0.0)
        
    return df

# --- VERBINDUNG ---
conn = st.connection("gsheets", type=GSheetsConnection)
df = load_data(conn)

# --- SIDEBAR MENÜ ---
st.sidebar.title("⛪ Hatler Minis")
menu = st.sidebar.radio("Menü", ["📊 Cockpit & Journal", "✍️ Neue Buchung", "💸 Offene Zahlungen", "📄 Dokumente"])

# ==============================================================================
# 1. COCKPIT & JOURNAL
# ==============================================================================
if menu == "📊 Cockpit & Journal":
    st.title("📊 Finanz-Übersicht")
    
    # Berechnungen
    budget = df["Einnahme"].sum() - df["Ausgabe"].sum()
    
    # Bankstand (Nur Status 'Erledigt')
    real_df = df[df["Status"] == "Erledigt"]
    bank_real = real_df["Einnahme"].sum() - real_df["Ausgabe"].sum()
    
    # Offene Rechnungen
    offen_df = df[(df["Status"] == "Offen") & (df["Ausgabe"] > 0)]
    offen_summe = offen_df["Ausgabe"].sum()

    # Anzeige
    col1, col2, col3 = st.columns(3)
    col1.metric("💰 Verfügbares Budget", f"{budget:,.2f} €")
    col2.metric("🏦 Kontostand (Real)", f"{bank_real:,.2f} €", delta=f"- {offen_summe:.2f} € offen", delta_color="inverse")
    col3.metric("📄 Offene Rechnungen", f"{len(offen_df)} Stück")
    
    st.markdown("---")
    st.subheader("Buchungsjournal")
    
    # Tabelle anzeigen
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
        
        # Status Logik
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
                
                # Neue Zeile erstellen
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
                
                # Anhängen
                updated_df = pd.concat([df, new_entry], ignore_index=True)
                
                # SPEICHERN (Ohne Try/Except Block, damit Fehler 200 ignoriert wird)
                conn.update(worksheet="Buchungen", data=updated_df)
                
                st.success("Gespeichert! Bitte Seite neu laden (F5) für Update.")

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
            # Update Logik im DataFrame
            mask = (df["Anlass_Person"] == entry_to_close) & (df["Status"] == "Offen")
            df.loc[mask, "Status"] = "Erledigt"
            
            conn.update(worksheet="Buchungen", data=df)
            st.success("Markiert! Bitte neu laden.")

# ==============================================================================
# 4. DOKUMENTE
# ==============================================================================
elif menu == "📄 Dokumente":
    st.header("📄 Generator")
    st.info("Funktion bereit. Bitte 'vorlage_antrag.docx' auf GitHub hochladen.")
    # Hier kommt später der Code für docxtpl rein, wenn die Datei da ist.
