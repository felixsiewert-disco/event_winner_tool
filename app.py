
import streamlit as st
import pandas as pd
import pgeocode
import numpy as np
from io import BytesIO

# ─────────────────────────────────────────────
# HILFSFUNKTIONEN
# ─────────────────────────────────────────────

@st.cache_data
def lade_plz_daten():
    return pgeocode.Nominatim('de')

def berechne_distanz(lat1, lon1, lat2, lon2):
    """Luftlinie in km zwischen zwei GPS-Punkten."""
    R = 6371
    phi1, phi2 = np.radians(lat1), np.radians(lat2)
    dphi = np.radians(lat2 - lat1)
    dlam = np.radians(lon2 - lon1)
    a = np.sin(dphi/2)**2 + np.cos(phi1)*np.cos(phi2)*np.sin(dlam/2)**2
    return R * 2 * np.arctan2(np.sqrt(a), np.sqrt(1-a))

def weise_zone_zu(distanz, grenze_nah, grenze_fern):
    if distanz <= grenze_nah:
        return "Nahbereich"
    elif distanz <= grenze_fern:
        return "Fernbereich"
    else:
        return "Rest"

def erkenne_duplikate(df):
    """Findet doppelte E-Mail-Adressen."""
    return df[df.duplicated(subset=["email"], keep=False)]["email"].unique().tolist()

def klassifiziere_listen(df, history_dfs):
    """
    Vergleicht Bewerber mit historischen Gewinnern und Check-ins.
    Gibt jedem Bewerber einen Status: whitelist / greylist / blacklist
    """
    df = df.copy()
    df["status"] = "whitelist"
    df["status_grund"] = ""

    if not history_dfs:
        return df

    for event_name, gewinner_df, checkin_df in history_dfs:
        gewinner_emails = set(gewinner_df["email"].str.lower().dropna())
        checkin_emails  = set(checkin_df["email"].str.lower().dropna()) if checkin_df is not None else set()

        for idx, row in df.iterrows():
            email = str(row.get("email", "")).lower()
            if email in gewinner_emails:
                if checkin_df is not None:
                    if email not in checkin_emails:
                        # War Gewinner aber nicht erschienen → Blacklist
                        df.at[idx, "status"] = "blacklist"
                        df.at[idx, "status_grund"] = f"No-Show bei {event_name}"
                    else:
                        # War Gewinner und erschienen → Greylist
                        if df.at[idx, "status"] == "whitelist":
                            df.at[idx, "status"] = "greylist"
                            df.at[idx, "status_grund"] = f"Gewinner bei {event_name}"
                else:
                    # Nur Gewinnerliste, kein Check-in → Greylist
                    if df.at[idx, "status"] == "whitelist":
                        df.at[idx, "status"] = "greylist"
                        df.at[idx, "status_grund"] = f"Gewinner bei {event_name} (kein Check-in)"
    return df

def waehle_gewinner(df, anzahl, zone_prozente, geschlecht_quote, freigegebene_emails):
    """
    Hauptauswahl: Filtert nach Status, dann wählt nach Zone + Geschlecht.
    Gibt (gewinner_df, nachruecker_df) zurück.
    """
    # Blacklist/Greylist ausschließen (außer manuell freigegebene)
    aktive = df[
        (df["status"] == "whitelist") |
        (df["email"].str.lower().isin([e.lower() for e in freigegebene_emails]))
    ].copy()

    gewinner_list = []
    zonen = ["Nahbereich", "Fernbereich", "Rest"]
    ziel_pro_zone = {
        "Nahbereich": round(anzahl * zone_prozente[0] / 100),
        "Fernbereich": round(anzahl * zone_prozente[1] / 100),
        "Rest":        anzahl - round(anzahl * zone_prozente[0] / 100) - round(anzahl * zone_prozente[1] / 100)
    }

    weiblich_ziel = round(anzahl * geschlecht_quote / 100)
    weiblich_count = 0

    verwendet_idx = set()

    for zone in zonen:
        pool = aktive[aktive["zone"] == zone].copy()
        pool = pool[~pool.index.isin(verwendet_idx)]
        ziel = ziel_pro_zone[zone]
        ausgewaehlt = 0

        # Frauen zuerst (bis Quote erfüllt)
        for idx, row in pool[pool["geschlecht"].str.lower() == "w"].iterrows():
            if ausgewaehlt >= ziel:
                break
            if weiblich_count < weiblich_ziel:
                gewinner_list.append(idx)
                verwendet_idx.add(idx)
                weiblich_count += 1
                ausgewaehlt += 1

        # Männer auffüllen
        for idx, row in pool[pool["geschlecht"].str.lower() == "m"].iterrows():
            if ausgewaehlt >= ziel:
                break
            gewinner_list.append(idx)
            verwendet_idx.add(idx)
            ausgewaehlt += 1

        # Falls Zone leer: andere Zone auffüllen (Umverteilung)
        if ausgewaehlt < ziel:
            rest_pool = aktive[~aktive.index.isin(verwendet_idx)]
            for idx, row in rest_pool.iterrows():
                if ausgewaehlt >= ziel:
                    break
                gewinner_list.append(idx)
                verwendet_idx.add(idx)
                ausgewaehlt += 1

    gewinner_df   = df.loc[gewinner_list].copy()
    nachruecker_df = aktive[~aktive.index.isin(verwendet_idx)].head(round(len(gewinner_list) * 0.3))

    return gewinner_df, nachruecker_df

def exportiere_excel(gewinner_df, nachruecker_df):
    """Erstellt eine Excel-Datei mit zwei Tabellenblättern."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Pretix-Format
        pretix_cols = {
            "email": "email",
            "vorname": "attendee_name",
            "nachname": "attendee_name"
        }
        pretix_df = gewinner_df.copy()
        if "vorname" in pretix_df.columns and "nachname" in pretix_df.columns:
            pretix_df["attendee_name"] = pretix_df["vorname"] + " " + pretix_df["nachname"]
        pretix_df["item"] = "Gewinner-Ticket"
        pretix_df[["email", "attendee_name", "item"]].to_excel(writer, sheet_name="Gewinner_pretix", index=False)

        # Nachruecker
        nachruecker_df.to_excel(writer, sheet_name="Nachruecker", index=False)

    return output.getvalue()


# ─────────────────────────────────────────────
# STREAMLIT UI
# ─────────────────────────────────────────────

st.set_page_config(page_title="Event Winner Tool", layout="wide")
st.title("🎟️ Event-Bewerber-Management-Tool")
st.markdown("Automatische Gewinner-Auswahl mit Zonen, Quoten & Blacklist/Greylist")

nomi = lade_plz_daten()

# ── 1. BEWERBER IMPORTIEREN ──────────────────
st.header("1️⃣ Bewerber importieren")
bewerber_file = st.file_uploader("Bewerber-Excel hochladen (Pflicht)", type=["xlsx"])

df_bewerber = None
if bewerber_file:
    df_raw = pd.read_excel(bewerber_file)
    st.write("**Spalten erkannt:**", list(df_raw.columns))

    # Automatische Spaltenzuordnung (Mapping)
    spalten_map = {}
    moegliche = {
        "email":      ["email", "e-mail", "mail", "emailadresse"],
        "vorname":    ["vorname", "first name", "firstname", "first_name"],
        "nachname":   ["nachname", "last name", "lastname", "last_name", "name"],
        "geschlecht": ["geschlecht", "gender", "sex", "anrede"],
        "plz":        ["plz", "postleitzahl", "zip", "postal code", "postal_code"],
    }
    for feld, varianten in moegliche.items():
        for col in df_raw.columns:
            if col.lower().strip() in varianten:
                spalten_map[feld] = col
                break

    st.write("**Automatische Spalten-Zuordnung:**", spalten_map)

    # Manuelle Korrektur falls nötig
    with st.expander("Spalten manuell zuordnen (falls falsch)"):
        alle_spalten = ["(nicht vorhanden)"] + list(df_raw.columns)
        for feld in ["email", "vorname", "nachname", "geschlecht", "plz"]:
            aktuell = spalten_map.get(feld, "(nicht vorhanden)")
            auswahl = st.selectbox(f"{feld}", alle_spalten,
                                   index=alle_spalten.index(aktuell) if aktuell in alle_spalten else 0,
                                   key=f"map_{feld}")
            spalten_map[feld] = auswahl

    # DataFrame normalisieren
    df_bewerber = pd.DataFrame()
    for feld, col in spalten_map.items():
        if col != "(nicht vorhanden)" and col in df_raw.columns:
            df_bewerber[feld] = df_raw[col]
        else:
            df_bewerber[feld] = ""

    # Duplikate anzeigen
    duplikate = erkenne_duplikate(df_bewerber)
    if duplikate:
        st.warning(f"⚠️ {len(duplikate)} doppelte E-Mail(s) gefunden: {duplikate[:5]}")
        df_bewerber = df_bewerber.drop_duplicates(subset=["email"], keep="first")
        st.info("Duplikate automatisch zusammengeführt (erste Zeile behalten).")

    st.success(f"✅ {len(df_bewerber)} Bewerber geladen")
    st.dataframe(df_bewerber.head(10))


# ── 2. HISTORY IMPORTIEREN ──────────────────
st.header("2️⃣ Event-Historie (optional)")
history_dfs = []

for i in range(1, 4):
    with st.expander(f"Vergangenes Event {i}"):
        gew_file = st.file_uploader(f"Event {i}: Gewinner-Excel", type=["xlsx"], key=f"gew{i}")
        ci_file  = st.file_uploader(f"Event {i}: Check-in-Excel",  type=["xlsx"], key=f"ci{i}")
        if gew_file:
            gew_df = pd.read_excel(gew_file)
            gew_df.columns = [c.lower().strip() for c in gew_df.columns]
            ci_df = None
            if ci_file:
                ci_df = pd.read_excel(ci_file)
                ci_df.columns = [c.lower().strip() for c in ci_df.columns]
            history_dfs.append((f"Event {i}", gew_df, ci_df))
            st.success(f"Event {i}: {len(gew_df)} Gewinner geladen" +
                       (f", {len(ci_df)} Check-ins" if ci_df is not None else ""))


# ── 3. KONFIGURATION ────────────────────────
st.header("3️⃣ Konfiguration")

col1, col2 = st.columns(2)
with col1:
    event_plz     = st.text_input("Event-PLZ", "10115")
    anzahl_gewinner = st.number_input("Anzahl Gewinner", 1, 500, 20)
    greylist_dauer  = st.slider("Greylist-Dauer (Events)", 1, 10, 3)
    blacklist_dauer = st.slider("Blacklist-Dauer (Events)", 1, 20, 5)

with col2:
    st.markdown("**Zonengrenzen (km)**")
    grenze_nah  = st.slider("Nahbereich bis km", 10, 300, 80)
    grenze_fern = st.slider("Fernbereich bis km", grenze_nah+10, 600, 200)

    st.markdown("**Zonenverteilung (%)**")
    pct_nah  = st.slider("% Nahbereich",  0, 100, 70)
    pct_fern = st.slider("% Fernbereich", 0, 100-pct_nah, 25)
    pct_rest = 100 - pct_nah - pct_fern
    st.markdown(f"% Rest: **{pct_rest}%**")

    geschlecht_quote = st.slider("% Frauen-Anteil", 0, 100, 60)


# ── 4. LISTEN-KLASSIFIZIERUNG & FREIGABE ────
if df_bewerber is not None:
    st.header("4️⃣ Listen-Status & manuelle Freigabe")

    # Distanz & Zone berechnen
    event_geo = nomi.query_postal_code(event_plz)
    event_lat, event_lon = event_geo.latitude, event_geo.longitude

    def berechne_row(row):
        try:
            geo = nomi.query_postal_code(str(row["plz"]).zfill(5))
            dist = berechne_distanz(event_lat, event_lon, geo.latitude, geo.longitude)
            zone = weise_zone_zu(dist, grenze_nah, grenze_fern)
            return pd.Series([round(dist, 1), zone])
        except:
            return pd.Series([None, "Unbekannt"])

    df_bewerber[["distanz_km", "zone"]] = df_bewerber.apply(berechne_row, axis=1)
    df_bewerber = klassifiziere_listen(df_bewerber, history_dfs)

    # Zusammenfassung
    status_counts = df_bewerber["status"].value_counts()
    c1, c2, c3 = st.columns(3)
    c1.metric("✅ Whitelist", status_counts.get("whitelist", 0))
    c2.metric("🟡 Greylist",  status_counts.get("greylist", 0))
    c3.metric("🔴 Blacklist", status_counts.get("blacklist", 0))

    # Manuelle Freigabe
    freigegebene = st.session_state.get("freigegebene", set())

    for liste_name, farbe in [("greylist", "🟡"), ("blacklist", "🔴")]:
        gruppe = df_bewerber[df_bewerber["status"] == liste_name]
        if not gruppe.empty:
            with st.expander(f"{farbe} {liste_name.capitalize()} ({len(gruppe)} Personen)"):
                if st.button(f"Alle {liste_name.capitalize()} freigeben", key=f"alle_{liste_name}"):
                    freigegebene.update(gruppe["email"].str.lower().tolist())
                for _, row in gruppe.iterrows():
                    col_a, col_b = st.columns([4, 1])
                    col_a.write(f"{row.get('vorname','')} {row.get('nachname','')} | {row['email']} | {row['status_grund']}")
                    if col_b.button("Freigeben", key=f"fg_{row['email']}"):
                        freigegebene.add(str(row["email"]).lower())

    st.session_state["freigegebene"] = freigegebene
    if freigegebene:
        st.info(f"🔓 {len(freigegebene)} Person(en) manuell freigegeben")


# ── 5. AUSWAHL & EXPORT ─────────────────────
    st.header("5️⃣ Gewinner auswählen & exportieren")

    if st.button("🎯 Jetzt Gewinner auswählen"):
        freigegebene = st.session_state.get("freigegebene", set())
        gewinner_df, nachruecker_df = waehle_gewinner(
            df_bewerber,
            anzahl_gewinner,
            [pct_nah, pct_fern, pct_rest],
            geschlecht_quote,
            freigegebene
        )

        st.success(f"✅ {len(gewinner_df)} Gewinner ausgewählt, {len(nachruecker_df)} Nachrücker")

        # Zonen-Statistik
        st.subheader("Zonen-Verteilung")
        st.dataframe(gewinner_df["zone"].value_counts().rename_axis("Zone").reset_index(name="Anzahl"))

        # Geschlecht
        st.subheader("Geschlecht-Verteilung")
        st.dataframe(gewinner_df["geschlecht"].value_counts().rename_axis("Geschlecht").reset_index(name="Anzahl"))

        st.subheader("Gewinnerliste Vorschau")
        st.dataframe(gewinner_df)

        # Export
        excel_bytes = exportiere_excel(gewinner_df, nachruecker_df)
        st.download_button(
            "📥 Excel herunterladen (Gewinner + Nachrücker)",
            data=excel_bytes,
            file_name="Gewinnerliste_pretix.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
