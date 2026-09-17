"""DATEV-Buchungsstapel für die LN Rechnungserfassung.

Erzeugt eine Datei im DATEV-Format EXTF, Version 700, Kategorie 21
(Buchungsstapel). Aufbau: eine Kopfzeile mit den Mandantendaten, darunter
die Spaltenüberschriften, danach je Rechnung eine Buchungszeile.

Wichtig und bewusst so gebaut:

* Es wird nichts geraten. Ohne hinterlegtes Kreditoren- und Aufwandskonto
  wird die Rechnung nicht exportiert, sondern als offen gemeldet.
* Das Ergebnis heißt Buchungsvorschlag. Vor dem Import beim Steuerberater
  gehört ein Testlauf mit einem einzelnen Monat.
* Kodierung ist Windows-1252, nicht UTF-8. DATEV erwartet das so; mit UTF-8
  brechen Umlaute.
"""
from __future__ import annotations

import io
import re
from datetime import datetime

from excel_report import _zahl, normalisiere, parse_date, skonto_werte

FORMAT_VERSION = 700
FORMAT_KATEGORIE = 21          # Buchungsstapel
FORMAT_NAME = "Buchungsstapel"
FORMAT_UNTERVERSION = 13

# Spaltenüberschriften der Buchungszeilen (Zeile 2 der Datei).
# Nur die ersten Felder werden gefüllt, der Rest bleibt leer – das ist
# zulässig, die Spaltenzahl muss aber stimmen.
SPALTEN = [
    "Umsatz (ohne Soll/Haben-Kz)", "Soll/Haben-Kennzeichen", "WKZ Umsatz",
    "Kurs", "Basis-Umsatz", "WKZ Basis-Umsatz", "Konto",
    "Gegenkonto (ohne BU-Schlüssel)", "BU-Schlüssel", "Belegdatum",
    "Belegfeld 1", "Belegfeld 2", "Skonto", "Buchungstext", "Postensperre",
    "Diverse Adressnummer", "Geschäftspartnerbank", "Sachverhalt",
    "Zinssperre", "Beleglink",
    "Beleginfo - Art 1", "Beleginfo - Inhalt 1",
    "Beleginfo - Art 2", "Beleginfo - Inhalt 2",
    "Beleginfo - Art 3", "Beleginfo - Inhalt 3",
    "Beleginfo - Art 4", "Beleginfo - Inhalt 4",
    "Beleginfo - Art 5", "Beleginfo - Inhalt 5",
    "Beleginfo - Art 6", "Beleginfo - Inhalt 6",
    "Beleginfo - Art 7", "Beleginfo - Inhalt 7",
    "Beleginfo - Art 8", "Beleginfo - Inhalt 8",
    "KOST1 - Kostenstelle", "KOST2 - Kostenstelle", "Kost-Menge",
    "EU-Land u. UStID", "EU-Steuersatz", "Abw. Versteuerungsart",
    "Sachverhalt L+L", "Funktionsergänzung L+L",
    "BU 49 Hauptfunktionstyp", "BU 49 Hauptfunktionsnummer",
    "BU 49 Funktionsergänzung", "Zusatzinformation - Art 1",
    "Zusatzinformation- Inhalt 1",
]

# Übliche Vorsteuerschlüssel. Sie sind NICHT allgemeingültig – der
# Steuerberater legt fest, was im Mandanten gilt. Deshalb nur als Vorschlag.
STEUERSCHLUESSEL_VORSCHLAG = {
    "SKR03": {19.0: "9", 7.0: "8", 0.0: ""},
    "SKR04": {19.0: "9", 7.0: "8", 0.0: ""},
}


def standard_einstellungen() -> dict:
    return {
        "beraternummer": "",
        "mandantennummer": "",
        "wj_beginn": f"{datetime.now().year}0101",
        "sachkontenlaenge": 4,
        "kontenrahmen": "SKR03",
        "bezeichnung": "Eingangsrechnungen",
        "herkunft": "LN",
        "festschreibung": False,
        "steuerschluessel": {"19": "9", "7": "8", "0": ""},
        "lieferanten": {},        # Name -> {"kreditor": "70001", "aufwand": "4980"}
        "geaendert_am": "",
    }


# ------------------------------------------------------------- Hilfsmittel

def _txt(v) -> str:
    """Textfeld: in Anführungszeichen, eigene Anführungszeichen verdoppelt."""
    s = "" if v is None else str(v)
    return '"' + s.replace('"', '""') + '"'


def _betrag(v) -> str:
    """Betrag mit Komma, ohne Tausenderpunkt, immer zwei Nachkommastellen."""
    return f"{abs(float(v)):.2f}".replace(".", ",")


def _nur_ziffern(v) -> str:
    return re.sub(r"\D", "", str(v or ""))


def kopfzeile(e: dict, von: datetime | None, bis: datetime | None) -> str:
    """Zeile 1 der Datei – Aufbau und Reihenfolge sind von DATEV vorgegeben."""
    jetzt = datetime.now().strftime("%Y%m%d%H%M%S") + "000"
    felder = [
        _txt("EXTF"),
        str(FORMAT_VERSION),
        str(FORMAT_KATEGORIE),
        _txt(FORMAT_NAME),
        str(FORMAT_UNTERVERSION),
        jetzt,
        "",                                   # importiert
        _txt((e.get("herkunft") or "LN")[:2]),
        _txt(""),                             # exportiert von
        _txt(""),                             # importiert von
        _nur_ziffern(e.get("beraternummer")),
        _nur_ziffern(e.get("mandantennummer")),
        _nur_ziffern(e.get("wj_beginn"))[:8],
        str(int(e.get("sachkontenlaenge") or 4)),
        von.strftime("%Y%m%d") if von else "",
        bis.strftime("%Y%m%d") if bis else "",
        _txt(e.get("bezeichnung") or "Eingangsrechnungen"),
        _txt(""),                             # Diktatkürzel
        "1",                                  # Buchungstyp: Finanzbuchführung
        "0",                                  # Rechnungslegungszweck
        "1" if e.get("festschreibung") else "0",
        _txt("EUR"),
        "", _txt(""), "", "",
        _txt("03" if (e.get("kontenrahmen") or "SKR03") == "SKR03" else "04"),
        "", "", _txt(""), _txt(""),
    ]
    return ";".join(felder)


def buchungszeile(inv: dict, e: dict) -> tuple[str | None, str | None]:
    """Eine Buchungszeile bauen. Liefert (Zeile, Fehlergrund)."""
    lieferant = (inv.get("lieferant") or "").strip()
    zuordnung = (e.get("lieferanten") or {}).get(lieferant)
    if not zuordnung or not zuordnung.get("kreditor") or not zuordnung.get("aufwand"):
        return None, ("Keine Kontierung hinterlegt"
                      if lieferant else "Kein Lieferant erkannt")

    brutto = _zahl(inv.get("brutto_gesamt"))
    if brutto is None:
        return None, "Kein Bruttobetrag"
    rd = parse_date(inv.get("rechnungsdatum"))
    if rd is None:
        return None, "Kein Rechnungsdatum"
    nummer = (inv.get("rechnungsnummer") or "").strip()
    if not nummer:
        return None, "Keine Rechnungsnummer"

    satz = _zahl(inv.get("ust_satz_prozent"))
    schluessel = ""
    if satz is not None:
        key = f"{satz:g}"
        schluessel = (e.get("steuerschluessel") or {}).get(key, "")

    sp, st_, _stichtag = skonto_werte(inv)
    skonto = ""
    if sp and brutto:
        skonto = _betrag(brutto * sp / 100)

    text = f"{nummer} {lieferant}"[:60]

    felder = [""] * len(SPALTEN)
    felder[0] = _betrag(brutto)
    felder[1] = "H" if brutto < 0 else "S"      # Gutschrift dreht die Seite
    felder[2] = _txt(inv.get("waehrung") or "EUR")
    felder[6] = _nur_ziffern(zuordnung["aufwand"])      # Konto = Aufwand
    felder[7] = _nur_ziffern(zuordnung["kreditor"])     # Gegenkonto = Kreditor
    felder[8] = _txt(schluessel) if schluessel else ""
    felder[9] = rd.strftime("%d%m")                     # Belegdatum TTMM
    felder[10] = _txt(nummer[:36])
    felder[12] = skonto
    felder[13] = _txt(text)
    return ";".join(felder), None


def datev_datei(invs: list[dict], e: dict) -> tuple[bytes, list[dict], int]:
    """Buchungsstapel erzeugen.

    Liefert (Dateiinhalt, Liste der übersprungenen Rechnungen, Anzahl Zeilen).
    """
    zeilen, offen, daten = [], [], []
    for inv in invs:
        normalisiere(inv)
        zeile, grund = buchungszeile(inv, e)
        if zeile is None:
            offen.append({
                "Rechnung": inv.get("rechnungsnummer") or inv.get("datei"),
                "Lieferant": inv.get("lieferant") or "–",
                "Grund": grund,
            })
            continue
        zeilen.append(zeile)
        rd = parse_date(inv.get("rechnungsdatum"))
        if rd:
            daten.append(rd)

    von, bis = (min(daten), max(daten)) if daten else (None, None)
    inhalt = "\r\n".join(
        [kopfzeile(e, von, bis), ";".join(SPALTEN)] + zeilen
    ) + "\r\n"
    # DATEV erwartet Windows-1252. Zeichen, die es dort nicht gibt, werden
    # ersetzt statt einen Fehler auszulösen.
    return inhalt.encode("cp1252", errors="replace"), offen, len(zeilen)


def pruefe_einstellungen(e: dict) -> list[str]:
    """Was fehlt noch, bevor ein Export sinnvoll ist?"""
    fehlt = []
    if not _nur_ziffern(e.get("beraternummer")):
        fehlt.append("Beraternummer")
    if not _nur_ziffern(e.get("mandantennummer")):
        fehlt.append("Mandantennummer")
    if len(_nur_ziffern(e.get("wj_beginn"))) != 8:
        fehlt.append("Beginn des Wirtschaftsjahres (JJJJMMTT)")
    return fehlt


# ------------------------------------------------------------- Oberfläche

def render_datev(invs: list[dict], laden, speichern) -> None:
    """Einstellungen und Export. `laden`/`speichern` kommen aus der App."""
    import pandas as pd
    import streamlit as st

    if "datev" not in st.session_state:
        gespeichert = laden() or {}
        e = standard_einstellungen()
        e.update(gespeichert)
        st.session_state["datev"] = e
    e = st.session_state["datev"]

    st.markdown("**Mandantendaten**")
    st.caption("Diese Angaben stehen in der Kopfzeile der Datei. "
               "Sie kommen vom Steuerberater – nicht raten.")
    c1, c2, c3 = st.columns(3)
    e["beraternummer"] = c1.text_input("Beraternummer", e.get("beraternummer", ""))
    e["mandantennummer"] = c2.text_input("Mandantennummer",
                                         e.get("mandantennummer", ""))
    e["wj_beginn"] = c3.text_input("Wirtschaftsjahr beginnt am (JJJJMMTT)",
                                   e.get("wj_beginn", ""))
    c4, c5, c6 = st.columns(3)
    e["kontenrahmen"] = c4.selectbox(
        "Kontenrahmen", ["SKR03", "SKR04"],
        index=0 if (e.get("kontenrahmen") or "SKR03") == "SKR03" else 1)
    e["sachkontenlaenge"] = c5.number_input(
        "Sachkontenlänge", 4, 8, int(e.get("sachkontenlaenge") or 4))
    e["bezeichnung"] = c6.text_input("Bezeichnung des Stapels",
                                     e.get("bezeichnung", "Eingangsrechnungen"))
    e["festschreibung"] = st.checkbox(
        "Buchungen festschreiben", bool(e.get("festschreibung")),
        help="Festgeschriebene Buchungen lassen sich nicht mehr ändern. "
             "Im Zweifel aus lassen und den Steuerberater fragen.")

    st.divider()
    st.markdown("**Steuerschlüssel je USt-Satz**")
    st.caption("Die Vorschläge sind üblich, aber nicht allgemeingültig. "
               "Der Steuerberater sagt, was im Mandanten gilt.")
    saetze = sorted({f"{_zahl(i.get('ust_satz_prozent')):g}"
                     for i in invs if _zahl(i.get("ust_satz_prozent")) is not None})
    for s in saetze or ["19", "7", "0"]:
        e.setdefault("steuerschluessel", {}).setdefault(s, "")
    ss = e["steuerschluessel"]
    spalten = st.columns(max(len(ss), 1))
    for sp, satz in zip(spalten, sorted(ss, key=lambda x: -float(x))):
        ss[satz] = sp.text_input(f"{satz} % USt", ss.get(satz, ""),
                                 key=f"bu_{satz}")

    st.divider()
    st.markdown("**Kontierung je Lieferant**")
    st.caption("Kreditorenkonto und Standard-Aufwandskonto. Beides steht im "
               "Kontenplan des Mandanten. Lässt sich direkt in der Tabelle "
               "eintragen.")
    namen = sorted({(i.get("lieferant") or "").strip()
                    for i in invs if (i.get("lieferant") or "").strip()})
    vorhanden = e.get("lieferanten") or {}
    tabelle = pd.DataFrame([
        {"Lieferant": n,
         "Kreditorenkonto": vorhanden.get(n, {}).get("kreditor", ""),
         "Aufwandskonto": vorhanden.get(n, {}).get("aufwand", "")}
        for n in namen
    ])
    if tabelle.empty:
        st.info("Noch keine Lieferanten – zuerst Rechnungen auslesen.")
    else:
        bearbeitet = st.data_editor(
            tabelle, width="stretch", hide_index=True, key="datev_lieferanten",
            column_config={
                "Lieferant": st.column_config.TextColumn(disabled=True,
                                                         width="large"),
                "Kreditorenkonto": st.column_config.TextColumn(width="medium"),
                "Aufwandskonto": st.column_config.TextColumn(width="medium"),
            },
        )
        e["lieferanten"] = {
            r["Lieferant"]: {"kreditor": str(r["Kreditorenkonto"] or "").strip(),
                             "aufwand": str(r["Aufwandskonto"] or "").strip()}
            for _, r in bearbeitet.iterrows()
        }

    st.divider()
    s1, s2 = st.columns([1, 3])
    if s1.button("Einstellungen speichern", type="primary", width="stretch"):
        e["geaendert_am"] = datetime.now().strftime("%d.%m.%Y %H:%M")
        speichern(e)
        st.success("Gespeichert.")
    if e.get("geaendert_am"):
        s2.caption(f"Zuletzt gespeichert: {e['geaendert_am']}")

    st.divider()
    st.markdown("**Buchungsstapel erzeugen**")
    fehlt = pruefe_einstellungen(e)
    if fehlt:
        st.warning("Es fehlt noch: " + ", ".join(fehlt))
        return
    if not invs:
        st.info("Keine Rechnungen vorhanden.")
        return

    inhalt, offen, anzahl = datev_datei(invs, e)
    if anzahl:
        st.download_button(
            f"DATEV-Buchungsstapel ({anzahl} Buchung(en))",
            inhalt,
            file_name=f"EXTF_Buchungsstapel_"
                      f"{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv",
            type="primary",
        )
    else:
        st.warning("Keine Rechnung ist vollständig kontiert – "
                   "kein Stapel erzeugt.")

    if offen:
        st.markdown(f"**{len(offen)} Rechnung(en) nicht im Stapel**")
        st.dataframe(pd.DataFrame(offen), width="stretch", hide_index=True)

    st.caption(
        "Das ist ein Buchungsvorschlag, keine fertige Buchhaltung. "
        "Bitte vor dem ersten echten Import einen einzelnen Monat testweise "
        "einlesen lassen. Kodierung Windows-1252, Format EXTF 700."
    )
