import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
from datetime import datetime, date
from docxtpl import DocxTemplate
import io

# --- KONFIGURATION ---
st.set_page_config(page_title="Vereins-Cockpit", layout="wide", page_icon="⛪")

# --- HILFSFUNKTIONEN ---
def load_data(conn):
    # Lädt Spalten A bis H (8 Spalten)
    df = conn.read(usecols=list(range(8)), ttl=0)
    df = df.dropna(how="all")
    # Datums-Konvertierung erzwingen (Tag zuerst, z.B. 01.01.2024)
    df["Datum"] = pd.to_datetime(df["Datum"], dayfirst=True, errors='coerce')
    # Leere Werte auffüllen für Berechnungen
    df["Einnahme"] = df["Einnahme"].fillna(0.0)
    df["Ausgabe"] = df["Ausgabe"].fillna(0.0)
    return df

# --- VERBINDUNG ---
conn = st.connection("gsheets", type=GSheetsConnection)

try:
    df = load_data(conn)
except Exception as e:
    st.error(f"Genauer Fehler: {e}")
    st.stop()

# --- SIDEBAR MENÜ ---
st.sidebar.title("⛪ Hatler Minis")
menu = st.sidebar.radio("Menü", ["📊 Cockpit & Journal", "✍️ Neue Buchung", "💸 Offene Zahlungen", "📄 Dokumente"])

# ==============================================================================
# 1. COCKPIT & JOURNAL
# ==============================================================================
if menu == "📊 Cockpit & Journal":
    st.title("📊 Finanz-Übersicht")
    
    # --- BERECHNUNG DER KENNZAHLEN ---
    
    # 1. Verfügbares Budget (Alles was gebucht ist, egal ob bezahlt oder nicht)
    budget = df["Einnahme"].sum() - df["Ausgabe"].sum()
    
    # 2. Echter Bankstand (Nur was Status "Erledigt" hat UND Konto "Bank" ist)
    # Wir nehmen an: Alles was "Erledigt" ist, ist real geflossen.
    # Achtung: Wir summieren hier alle Konten, die "Erledigt" sind. 
    # Wenn du NUR Bank willst: df[(df["Status"] == "Erledigt") & (df["Konto"] == "Bank")]
    real_df = df[df["Status"] == "Erledigt"]
    bank_real = real_df["Einnahme"].sum() - real_df["Ausgabe"].sum()
    
    # 3. Offene Rechnungen (Summe aller Ausgaben mit Status "Offen")
    offen_df = df[(df["Status"] == "Offen") & (df["Ausgabe"] > 0)]
    offen_summe = offen_df["Ausgabe"].sum()

    # --- ANZEIGE ---
    col1, col2, col3 = st.columns(3)
    
    col1.metric(
        label="💰 Verfügbares Budget",
        value=f"{budget:,.2f} €",
        help="Das darf noch ausgegeben werden (Einnahmen - Ausgaben)"
    )
    
    col2.metric(
        label="🏦 Kontostand (Real)",
        value=f"{bank_real:,.2f} €",
        delta=f"- {offen_summe:.2f} € noch offen",
        delta_color="inverse",
        help="Das liegt tatsächlich auf dem Konto (Status 'Erledigt')"
    )
    
    col3.metric(
        label="📄 Offene Rechnungen",
        value=f"{len(offen_df)} Stück",
        help="Anzahl der Rechnungen mit Status 'Offen'"
    )
    
    st.markdown("---")
    
    # --- JOURNAL TABELLE ---
    st.subheader("Buchungsjournal")
    
    # Wir formatieren das Datum für die Anzeige schön deutsch
    display_df = df.copy()
    display_df["Datum"] = display_df["Datum"].dt.strftime("%d.%m.%Y")
    
    st.dataframe(
        display_df.sort_index(ascending=False), 
        use_container_width=True,
        column_config={
            "Einnahme": st.column_config.NumberColumn(format="%.2f €"),
            "Ausgabe": st.column_config.NumberColumn(format="%.2f €"),
            "Status": st.column_config.Column(
                width="small",
                help="Offen = Noch nicht überwiesen",
            )
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
        anlass_in = col_b.text_input("Anlass / Person", placeholder="z.B. Einkauf Lager")
        
        typ = st.radio("Buchungstyp", ["Ausgabe", "Einnahme"], horizontal=True)
        
        betrag_in = st.number_input("Betrag (€)", min_value=0.01, format="%.2f")
        bemerkung_in = st.text_input("Bemerkung (optional)")
        
        c1, c2, c3 = st.columns(3)
        konto_in = c1.selectbox("Konto", ["Bank", "Handkassa", "Minikonto"])
        rechnung_in = c2.checkbox("Rechnung vorhanden?", value=True)
        
        # Logik: Wenn Handkassa, ist es meist sofort erledigt. Wenn Bank, oft erst "Offen".
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
                conn.update(worksheet="Buchungen", data=updated_df)
                st.success("Buchung gespeichert!")
                #st.rerun()

# ==============================================================================
# 3. OFFENE ZAHLUNGEN (DEIN BEREICH)
# ==============================================================================
elif menu == "💸 Offene Zahlungen":
    st.header("💸 Offene Überweisungen")
    st.info("Hier siehst du alle Ausgaben mit Status 'Offen'. Wenn du überwiesen hast, ändere den Status im Google Sheet oder hier.")

    # Filter: Nur Ausgaben, die Offen sind
    mask_offen = (df["Status"] == "Offen") & (df["Ausgabe"] > 0)
    todos = df[mask_offen].copy()
    
    if todos.empty:
        st.success("Alles erledigt! Keine offenen Rechnungen. 🎉")
    else:
        # Wir zeigen die Liste an
        st.table(todos[["Datum", "Anlass_Person", "Ausgabe", "Konto"]])
        
        st.write("---")
        st.write("**Status ändern:**")
        # Workaround: Da wir keine Datenbank-IDs haben, wählen wir über den Anlass aus
        # (In einer Profi-App hätten wir IDs, hier halten wir es simpel)
        entry_to_close = st.selectbox("Welchen Eintrag hast du bezahlt?", todos["Anlass_Person"].unique())
        
        if st.button("Als 'Erledigt' markieren"):
            # Wir suchen die Zeile im Original-DF
            # Hinweis: Das ändert alle Einträge mit diesem Namen, die offen sind.
            mask_update = (df["Anlass_Person"] == entry_to_close) & (df["Status"] == "Offen")
            
            if mask_update.any():
                df.loc[mask_update, "Status"] = "Erledigt"
                # Optional: Datum auf heute setzen (Überweisungstag)?
                # df.loc[mask_update, "Datum"] = pd.to_datetime(date.today())
                
                # Update Sheet
                conn.update(worksheet="Buchungen", data=df)
                st.balloons()
                st.success(f"{entry_to_close} wurde als bezahlt markiert!")
                st.rerun()
            else:
                st.error("Eintrag nicht gefunden.")

# ==============================================================================
# 4. DOKUMENTE (FÖRDERANTRAG)
# ==============================================================================
elif menu == "📄 Dokumente":
    st.header("📄 Generator für Förderanträge")
    
    st.markdown("Lade eine Datei namens `vorlage_antrag.docx` in dein Verzeichnis, damit das klappt.")
    
    col1, col2 = st.columns(2)
    p_name = col1.text_input("Projektname", "Minilager 2025")
    p_datum = col2.text_input("Zeitraum/Datum", "Sommer 2025")
    p_summe = col1.number_input("Gesamtkosten (€)", value=500.0)
    p_antragsteller = col2.text_input("Antragsteller", "Max Mustermann")
    
    if st.button("Dokument erstellen"):
        context = {
            "projekt_name": p_name,
            "datum": p_datum,
            "gesamtkosten": f"{p_summe:.2f}",
            "antragsteller": p_antragsteller
        }
        
        try:
            doc = DocxTemplate("vorlage_antrag.docx")
            doc.render(context)
            
            # Speichern in Memory Stream für Download
            bio = io.BytesIO()
            doc.save(bio)
            
            st.download_button(
                label="📥 Word-Datei herunterladen",
                data=bio.getvalue(),
                file_name=f"Antrag_{p_name}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            st.success("Dokument generiert!")
            
        except FileNotFoundError:
            st.error("Fehler: Die Datei 'vorlage_antrag.docx' wurde nicht gefunden. Bitte lade sie hoch!")
        except Exception as e:
            st.error(f"Ein Fehler ist aufgetreten: {e}")
