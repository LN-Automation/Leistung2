"""Verbrauchs- und Kostenübersicht für die LN Rechnungserfassung.

Bei jedem Auslesen meldet die Anthropic-API zurück, wie viele Tokens
verbraucht wurden. Diese Zahlen werden hier festgehalten und je Monat
aufsummiert, damit der Kunde sieht, was seine Nutzung kostet.

Bewusst so gebaut:

* Die Tokenzahlen sind exakt – sie kommen aus der API-Antwort, nicht aus
  einer Schätzung.
* Die Preise sind einstellbar. Anbieterpreise ändern sich, und je nachdem,
  ob du zum Einkaufspreis weitergibst oder einen eigenen Satz abrechnest,
  gehört hier ein anderer Wert hinein. Voreingestellt ist der derzeit
  übliche Satz für Claude Sonnet.
* Der Euro-Betrag ist eine Umrechnung. Abgerechnet wird beim Anbieter in
  US-Dollar, der Kurs schwankt. Das steht auch so in der Oberfläche.
"""
from __future__ import annotations

from datetime import datetime

# Voreinstellung: US-Dollar je eine Million Tokens.
# Vor dem Einsatz beim Kunden gegen die aktuelle Preisliste des Anbieters
# prüfen – die Werte lassen sich in der Oberfläche ändern.
PREIS_INPUT_STANDARD = 3.00
PREIS_OUTPUT_STANDARD = 15.00
KURS_STANDARD = 0.92          # USD -> EUR, grober Richtwert


def standard_preise() -> dict:
    return {
        "preis_input": PREIS_INPUT_STANDARD,
        "preis_output": PREIS_OUTPUT_STANDARD,
        "kurs_usd_eur": KURS_STANDARD,
    }


def datensatz(datei: str, modell: str, usage) -> dict:
    """Einen Verbrauchsdatensatz aus der API-Antwort bauen."""
    def feld(name):
        return int(getattr(usage, name, 0) or 0)

    jetzt = datetime.now()
    return {
        "datei": datei,
        "modell": modell,
        "zeitpunkt": jetzt.strftime("%Y-%m-%d %H:%M:%S"),
        "monat": jetzt.strftime("%Y-%m"),
        "input_tokens": feld("input_tokens"),
        "output_tokens": feld("output_tokens"),
        "cache_read": feld("cache_read_input_tokens"),
        "cache_write": feld("cache_creation_input_tokens"),
    }


def kosten(satz: dict, preise: dict) -> tuple[float, float]:
    """Kosten eines Datensatzes in USD und EUR."""
    pi = float(preise.get("preis_input", PREIS_INPUT_STANDARD))
    po = float(preise.get("preis_output", PREIS_OUTPUT_STANDARD))
    kurs = float(preise.get("kurs_usd_eur", KURS_STANDARD))
    # Zwischengespeicherte Eingaben sind günstiger; ohne Prompt-Caching
    # sind diese Felder null und der Term fällt weg.
    usd = (
        (satz.get("input_tokens", 0) + satz.get("cache_write", 0)) / 1e6 * pi
        + satz.get("cache_read", 0) / 1e6 * pi * 0.1
        + satz.get("output_tokens", 0) / 1e6 * po
    )
    return usd, usd * kurs


def je_monat(saetze: list[dict], preise: dict) -> list[dict]:
    monate: dict[str, dict] = {}
    for s in saetze:
        m = monate.setdefault(s.get("monat", "?"), {
            "Monat": s.get("monat", "?"), "Belege": 0,
            "Eingabe-Tokens": 0, "Ausgabe-Tokens": 0, "USD": 0.0, "EUR": 0.0,
        })
        usd, eur = kosten(s, preise)
        m["Belege"] += 1
        m["Eingabe-Tokens"] += s.get("input_tokens", 0) + s.get("cache_read", 0)
        m["Ausgabe-Tokens"] += s.get("output_tokens", 0)
        m["USD"] += usd
        m["EUR"] += eur
    return [monate[k] for k in sorted(monate, reverse=True)]


def _zahl(v, stellen=2):
    return f"{v:,.{stellen}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def render_nutzung(saetze: list[dict], laden, speichern) -> None:
    """Übersicht und Preiseinstellungen."""
    import pandas as pd
    import streamlit as st

    gespeichert = laden() or {}
    preise = standard_preise()
    preise.update({k: gespeichert[k] for k in standard_preise() if k in gespeichert})

    if not saetze:
        st.info("Noch kein Verbrauch erfasst. Die Zählung beginnt mit dem "
                "nächsten Auslesen von Rechnungen.")
    else:
        monate = je_monat(saetze, preise)
        aktuell = monate[0]
        gesamt_usd = sum(m["USD"] for m in monate)
        gesamt_eur = sum(m["EUR"] for m in monate)

        s1, s2, s3, s4 = st.columns(4, gap="large")
        with s1, st.container(border=True):
            st.metric("Laufender Monat (EUR)", _zahl(aktuell["EUR"]))
            st.caption(f"{aktuell['Belege']} Beleg(e)")
        with s2, st.container(border=True):
            st.metric("Laufender Monat (USD)", _zahl(aktuell["USD"]))
            st.caption(aktuell["Monat"])
        with s3, st.container(border=True):
            st.metric("Gesamt bisher (EUR)", _zahl(gesamt_eur))
            st.caption(f"{len(saetze)} Beleg(e) insgesamt")
        with s4, st.container(border=True):
            schnitt = gesamt_eur / len(saetze) if saetze else 0
            st.metric("Je Beleg (EUR)", _zahl(schnitt, 4))
            st.caption("Durchschnitt über alle Belege")

        st.write("")
        st.markdown("**Verbrauch je Monat**")
        df = pd.DataFrame(monate)
        st.dataframe(
            df, width="stretch", hide_index=True,
            column_config={
                "Eingabe-Tokens": st.column_config.NumberColumn(format="%d"),
                "Ausgabe-Tokens": st.column_config.NumberColumn(format="%d"),
                "USD": st.column_config.NumberColumn(format="%.2f"),
                "EUR": st.column_config.NumberColumn(format="%.2f"),
            },
        )

        with st.expander("Einzelne Belege"):
            zeilen = []
            for s in sorted(saetze, key=lambda x: x.get("zeitpunkt", ""),
                            reverse=True)[:500]:
                usd, eur = kosten(s, preise)
                zeilen.append({
                    "Zeitpunkt": s.get("zeitpunkt", ""),
                    "Datei": s.get("datei", ""),
                    "Modell": s.get("modell", ""),
                    "Eingabe": s.get("input_tokens", 0) + s.get("cache_read", 0),
                    "Ausgabe": s.get("output_tokens", 0),
                    "EUR": round(eur, 4),
                })
            st.dataframe(pd.DataFrame(zeilen), width="stretch", hide_index=True)

    st.divider()
    st.markdown("**Preise**")
    st.caption("Der Anbieter rechnet in US-Dollar je eine Million Tokens ab. "
               "Die Sätze ändern sich gelegentlich – hier eintragen, was "
               "tatsächlich berechnet wird.")
    c1, c2, c3 = st.columns(3)
    preise["preis_input"] = c1.number_input(
        "USD je Mio. Eingabe-Tokens", 0.0, 500.0,
        float(preise["preis_input"]), step=0.5, format="%.2f")
    preise["preis_output"] = c2.number_input(
        "USD je Mio. Ausgabe-Tokens", 0.0, 500.0,
        float(preise["preis_output"]), step=0.5, format="%.2f")
    preise["kurs_usd_eur"] = c3.number_input(
        "Umrechnungskurs USD → EUR", 0.1, 3.0,
        float(preise["kurs_usd_eur"]), step=0.01, format="%.2f")

    if st.button("Preise speichern", type="primary"):
        neu = dict(gespeichert)
        neu.update(preise)
        speichern(neu)
        st.success("Gespeichert.")

    st.caption(
        "Die Tokenzahlen stammen unverändert aus der Antwort des KI-Anbieters. "
        "Der Euro-Betrag ist eine Umrechnung zum oben eingetragenen Kurs und "
        "kann von der späteren Rechnung abweichen."
    )
