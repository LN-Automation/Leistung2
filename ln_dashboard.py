"""Dashboard für die LN Rechnungserfassung (Produkt 2).

Zeigt dieselben Auswertungen wie das Analyseblatt im Excel-Report, nur
interaktiv: oben ein Umschalter für die Währung, darunter Kennzahlen und
Grafiken. Getrennte Reiter je Währung braucht es dadurch nicht.

Aufruf aus der App:
    from ln_dashboard import render_dashboard
    render_dashboard(rechnungen)

Die Aufbereitung steckt in reinen Funktionen ohne Streamlit, damit sie sich
einzeln prüfen lässt.
"""
from __future__ import annotations

from datetime import datetime, timedelta

from excel_report import (
    STUFE_OHNE,
    STUFEN,
    _stufe,
    _zahl,
    normalisiere,
    parse_date,
    pruefhinweise,
    skonto_werte,
)

NAVY = "#0F172A"
BLUE = "#1A56DB"
RED = "#B91C1C"
GREEN = "#15803D"
GOLD = "#B48A2F"
GREY = "#94A3B8"
HELL = "#E2E8F0"

# Farben der Fälligkeitsstufen – gleiche Logik wie im Excel-Report
STUFEN_FARBEN = {
    "Überfällig > 30 Tage": "#E02424",
    "Überfällig 1–30 Tage": "#F05252",
    "Fällig in 0–7 Tagen": "#D9A404",
    "Fällig in 8–14 Tagen": "#E3B341",
    "Fällig in 15–30 Tagen": "#31C48D",
    "Fällig in über 30 Tagen": "#0E9F6E",
    STUFE_OHNE: "#9CA3AF",
}

BELEGKLASSEN = [
    ("bis 100", 0, 100),
    ("100 bis 500", 100, 500),
    ("500 bis 1.000", 500, 1000),
    ("1.000 bis 5.000", 1000, 5000),
    ("5.000 bis 10.000", 5000, 10000),
    ("über 10.000", 10000, None),
]


# ---------------------------------------------------------- Aufbereitung

def aufbereiten(invs: list[dict], heute: datetime | None = None) -> list[dict]:
    """Rechnungen einmal durchrechnen: Fälligkeit, Stufe, Skonto, Hinweise."""
    heute = (heute or datetime.now()).replace(hour=0, minute=0, second=0, microsecond=0)
    daten = []
    for inv in invs:
        normalisiere(inv)
        rd = parse_date(inv.get("rechnungsdatum"))
        fd = parse_date(inv.get("faelligkeitsdatum"))
        tage = (fd - heute).days if fd else None
        sp, st_, stichtag = skonto_werte(inv)
        if stichtag is not None:
            skonto_bis = stichtag
        elif rd is not None and st_ is not None:
            skonto_bis = rd + timedelta(days=st_)
        else:
            skonto_bis = None
        brutto = _zahl(inv.get("brutto_gesamt"))
        daten.append({
            "inv": inv,
            "lieferant": (inv.get("lieferant") or "Ohne Lieferant").strip(),
            "rd": rd,
            "fd": fd,
            "tage": tage,
            "stufe": _stufe(tage),
            "waehrung": inv.get("waehrung") or "EUR",
            "netto": _zahl(inv.get("netto_gesamt")),
            "ust": _zahl(inv.get("ust_betrag")),
            "brutto": brutto,
            "satz": _zahl(inv.get("ust_satz_prozent")),
            "skonto_prozent": sp,
            "skonto_betrag": (brutto * sp / 100) if (brutto is not None and sp) else None,
            "skonto_offen": bool(skonto_bis and skonto_bis >= heute),
            "zahlungsziel": (fd - rd).days if (rd and fd) else None,
            "hinweise": pruefhinweise(inv),
            "positionen": inv.get("positionen") or [],
        })
    return daten


def waehrungen_von(daten: list[dict]) -> list[str]:
    return sorted({d["waehrung"] for d in daten})


def kennzahlen(daten: list[dict], w: str) -> dict:
    teil = [d for d in daten if d["waehrung"] == w]
    return {
        "anzahl": len(teil),
        "gesamt": sum(d["brutto"] or 0 for d in teil),
        "ueberfaellig": sum(d["brutto"] or 0 for d in teil
                            if d["tage"] is not None and d["tage"] < 0),
        "anzahl_ueberfaellig": sum(1 for d in teil
                                   if d["tage"] is not None and d["tage"] < 0),
        "faellig14": sum(d["brutto"] or 0 for d in teil
                         if d["tage"] is not None and 0 <= d["tage"] <= 14),
        "skonto_offen": sum(d["skonto_betrag"] or 0 for d in teil if d["skonto_offen"]),
        "skonto_verfallen": sum(d["skonto_betrag"] or 0 for d in teil
                                if d["skonto_betrag"] and not d["skonto_offen"]),
        "mit_hinweis": sum(1 for d in teil if d["hinweise"]),
    }


def faelligkeitsstruktur(daten: list[dict], w: str) -> list[dict]:
    teil = [d for d in daten if d["waehrung"] == w]
    reihen = []
    for label in [s[1] for s in STUFEN] + [STUFE_OHNE]:
        gruppe = [d for d in teil if d["stufe"] == label]
        reihen.append({
            "Stufe": label,
            "Belege": len(gruppe),
            "Betrag": sum(d["brutto"] or 0 for d in gruppe),
        })
    return reihen


def lieferanten(daten: list[dict], w: str, grenze: int = 10) -> list[dict]:
    """Volumen je Lieferant mit kumuliertem Anteil und ABC-Klasse."""
    teil = [d for d in daten if d["waehrung"] == w]
    gesamt = sum(d["brutto"] or 0 for d in teil) or 1.0
    nach = {}
    for d in teil:
        e = nach.setdefault(d["lieferant"], {"betrag": 0.0, "anzahl": 0, "ziele": []})
        e["betrag"] += d["brutto"] or 0
        e["anzahl"] += 1
        if d["zahlungsziel"] is not None:
            e["ziele"].append(d["zahlungsziel"])

    reihen, laufend = [], 0.0
    for name, e in sorted(nach.items(), key=lambda x: -x[1]["betrag"]):
        vorher = laufend / gesamt
        laufend += e["betrag"]
        # Wer die Schwelle überschreitet, zählt noch zur unteren Klasse
        klasse = "A" if vorher < 0.80 else ("B" if vorher < 0.95 else "C")
        reihen.append({
            "Lieferant": name,
            "Belege": e["anzahl"],
            "Betrag": e["betrag"],
            "Kumuliert": laufend / gesamt,
            "Klasse": klasse,
            "Zahlungsziel": (sum(e["ziele"]) / len(e["ziele"])) if e["ziele"] else None,
        })
    return reihen[:grenze] if grenze else reihen


def monatsverlauf(daten: list[dict], w: str) -> list[dict]:
    teil = [d for d in daten if d["waehrung"] == w and d["rd"]]
    nach = {}
    for d in teil:
        key = d["rd"].strftime("%Y-%m")
        e = nach.setdefault(key, {"betrag": 0.0, "anzahl": 0})
        e["betrag"] += d["brutto"] or 0
        e["anzahl"] += 1
    return [{"Monat": k, "Belege": v["anzahl"], "Betrag": v["betrag"]}
            for k, v in sorted(nach.items())]


def belegsgroessen(daten: list[dict], w: str) -> list[dict]:
    teil = [d for d in daten if d["waehrung"] == w]
    reihen = []
    for label, von, bis in BELEGKLASSEN:
        gruppe = [d for d in teil if d["brutto"] is not None
                  and abs(d["brutto"]) >= von
                  and (bis is None or abs(d["brutto"]) < bis)]
        reihen.append({
            "Klasse": label,
            "Belege": len(gruppe),
            "Betrag": sum(d["brutto"] or 0 for d in gruppe),
        })
    return reihen


def skonto_status(daten: list[dict], w: str) -> list[dict]:
    teil = [d for d in daten if d["waehrung"] == w]
    erreichbar = [d for d in teil if d["skonto_betrag"] and d["skonto_offen"]]
    verfallen = [d for d in teil if d["skonto_betrag"] and not d["skonto_offen"]]
    ohne = [d for d in teil if not d["skonto_betrag"]]
    return [
        {"Status": "Frist läuft noch", "Belege": len(erreichbar),
         "Betrag": sum(d["skonto_betrag"] or 0 for d in erreichbar)},
        {"Status": "Frist abgelaufen", "Belege": len(verfallen),
         "Betrag": sum(d["skonto_betrag"] or 0 for d in verfallen)},
        {"Status": "Kein Skonto angeboten", "Belege": len(ohne), "Betrag": 0.0},
    ]


def artikel(daten: list[dict], w: str, grenze: int = 10) -> list[dict]:
    teil = [d for d in daten if d["waehrung"] == w]
    nach = {}
    for d in teil:
        for pos in d["positionen"]:
            bez = str(pos.get("bezeichnung") or "").strip()
            if not bez:
                continue
            e = nach.setdefault(bez[:45], {"betrag": 0.0, "anzahl": 0, "lief": set()})
            e["betrag"] += _zahl(pos.get("gesamt_netto")) or 0
            e["anzahl"] += 1
            e["lief"].add(d["lieferant"])
    reihen = [{"Bezeichnung": k, "Posten": v["anzahl"], "Betrag": v["betrag"],
               "Lieferanten": len(v["lief"])}
              for k, v in sorted(nach.items(), key=lambda x: -x[1]["betrag"])]
    return reihen[:grenze]


def dringend(daten: list[dict], grenze: int = 8) -> list[dict]:
    offen = [d for d in daten if d["tage"] is not None and d["tage"] <= 14]
    return sorted(offen, key=lambda d: d["tage"])[:grenze]


# ------------------------------------------------------------- Darstellung

def _zahl_kurz(v) -> str:
    """Betrag ohne Währung – die steht in der Beschriftung der Kennzahl."""
    if v is None:
        return "–"
    return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _geld(v, w="EUR"):
    if v is None:
        return "–"
    return f"{v:,.2f} {w}".replace(",", "X").replace(".", ",").replace("X", ".")


def _export_fassung(chart):
    """Grafik für den Export gleich einstellen wie im Dashboard.

    Altair rendert außerhalb von Streamlit mit seinem eigenen Standardthema –
    andere Schrift, andere Achsenfarben. Deshalb wird hier dasselbe Aussehen
    ausdrücklich gesetzt und eine feste Plotbreite vergeben, damit alle
    Grafiken im Bild gleich groß sind.
    """
    return (
        chart.properties(width=430, height=270)
        .configure_axis(
            labelColor="#475569", titleColor="#475569",
            gridColor="#E2E8F0", domainColor="#E2E8F0", tickColor="#E2E8F0",
            labelFontSize=11, titleFontSize=12, labelFont="sans-serif",
            titleFont="sans-serif",
        )
        .configure_legend(
            labelColor="#475569", titleColor="#475569",
            labelFontSize=11, titleFontSize=11, labelFont="sans-serif",
        )
        .configure_view(strokeWidth=0)
        .configure_title(color="#0F172A", fontSize=13, font="sans-serif")
    )


def dashboard_bild(grafiken, kopfzeile: str) -> bytes:
    """Alle Dashboard-Grafiken als ein PNG – im Aufbau des Dashboards.

    Zwei Spalten, jede Grafik in einer Karte mit Rahmen und Überschrift,
    alle Karten gleich groß. Braucht vl-convert-python.
    """
    from io import BytesIO

    try:
        import vl_convert as vlc
    except ImportError:
        raise RuntimeError(
            "Für den Bild-Export fehlt das Paket vl-convert-python. "
            "In requirements.txt ergänzen und die App neu starten."
        )
    from PIL import Image, ImageDraw, ImageFont

    S = 2                    # Skalierung, damit das Bild scharf ist
    bilder = []
    for titel, chart in grafiken:
        if chart is None:
            continue
        try:
            roh = vlc.vegalite_to_png(_export_fassung(chart).to_json(), scale=S)
        except Exception:     # noqa: BLE001  – einzelne Grafik überspringen
            continue
        bilder.append((titel, Image.open(BytesIO(roh)).convert("RGB")))
    if not bilder:
        raise RuntimeError("Keine Grafiken vorhanden.")

    def schrift(groesse, fett=False):
        namen = (["DejaVuSans-Bold.ttf", "Arial Bold.ttf", "Helvetica.ttc"]
                 if fett else ["DejaVuSans.ttf", "Arial.ttf", "Helvetica.ttc"])
        ordner = ["/usr/share/fonts/truetype/dejavu/", "/Library/Fonts/",
                  "/System/Library/Fonts/", "C:/Windows/Fonts/", ""]
        for o in ordner:
            for n in namen:
                try:
                    return ImageFont.truetype(o + n, groesse)
                except Exception:  # noqa: BLE001
                    continue
        return ImageFont.load_default()

    # Karten: alle gleich groß, Grafik mittig darin
    innen = 22 * S                       # Innenabstand der Karte
    titelhoehe = 34 * S
    inhalt_b = max(b.width for _, b in bilder)
    inhalt_h = max(b.height for _, b in bilder)
    karte_b = inhalt_b + innen * 2
    karte_h = inhalt_h + innen * 2 + titelhoehe

    rand = 28 * S
    luecke = 18 * S
    kopf_h = 74 * S
    spalten = 2
    zeilen = (len(bilder) + spalten - 1) // spalten

    breite = rand * 2 + karte_b * spalten + luecke
    hoehe = kopf_h + rand + zeilen * karte_h + (zeilen - 1) * luecke + rand

    blatt = Image.new("RGB", (breite, hoehe), (255, 255, 255))
    d = ImageDraw.Draw(blatt)

    # Kopfzeile im Farbton des Dashboards
    d.rectangle([0, 0, breite, kopf_h], fill=(15, 23, 42))
    d.text((rand, 20 * S), "LN Automation – Auswertung",
           font=schrift(15 * S, True), fill=(255, 255, 255))
    d.text((rand, 44 * S), kopfzeile, font=schrift(10 * S),
           fill=(148, 163, 184))

    for i, (titel, bild) in enumerate(bilder):
        sp, ze = i % spalten, i // spalten
        x = rand + sp * (karte_b + luecke)
        y = kopf_h + rand + ze * (karte_h + luecke)
        # Karte: heller Rahmen mit runden Ecken, wie st.container(border=True)
        d.rounded_rectangle([x, y, x + karte_b, y + karte_h], radius=8 * S,
                            outline=(226, 232, 240), width=1 * S,
                            fill=(255, 255, 255))
        d.text((x + innen, y + innen - 4 * S), titel,
               font=schrift(12 * S, True), fill=(15, 23, 42))
        bx = x + innen + (inhalt_b - bild.width) // 2
        by = y + innen + titelhoehe
        blatt.paste(bild, (bx, by))

    raus = BytesIO()
    blatt.save(raus, "PNG", optimize=True)
    return raus.getvalue()


def _zeigen(st, sammlung, titel, chart):
    """Grafik anzeigen und zugleich für den Bild-Export vormerken."""
    sammlung.append((titel, chart))
    st.altair_chart(chart, width="stretch")


def handlungsempfehlungen(daten: list[dict], w: str) -> list[dict]:
    """Konkrete nächste Schritte, ausschließlich aus den Zahlen abgeleitet.

    Es wird nichts geraten – jede Zeile nennt die Anzahl und den Betrag,
    die dahinterstehen.
    """
    teil = [d for d in daten if d["waehrung"] == w]
    punkte = []

    alt_faellig = [d for d in teil if d["tage"] is not None and d["tage"] < -30]
    if alt_faellig:
        punkte.append({
            "art": "rot",
            "titel": f"{len(alt_faellig)} Rechnung(en) über 30 Tage überfällig",
            "text": f"{_zahl_kurz(sum(d['brutto'] or 0 for d in alt_faellig))} {w}. "
                    f"Zuerst klären, ob bereits gezahlt wurde – sonst drohen "
                    f"Mahnkosten und Lieferstopp.",
        })

    bald = [d for d in teil if d["skonto_offen"] and d["skonto_betrag"]]
    if bald:
        punkte.append({
            "art": "gruen",
            "titel": f"Bei {len(bald)} Rechnung(en) läuft die Skontofrist noch",
            "text": f"{_zahl_kurz(sum(d['skonto_betrag'] for d in bald))} {w} "
                    f"sind noch erreichbar. Diese Zahlungen vorziehen.",
        })

    ohne_faellig = [d for d in teil if d["tage"] is None]
    if ohne_faellig:
        punkte.append({
            "art": "gelb",
            "titel": f"{len(ohne_faellig)} Rechnung(en) ohne erkennbare Fälligkeit",
            "text": f"{_zahl_kurz(sum(d['brutto'] or 0 for d in ohne_faellig))} {w} "
                    f"lassen sich nicht einplanen. Zahlungsziel am Beleg "
                    f"nachsehen oder beim Lieferanten erfragen.",
        })

    hinweise = [d for d in teil if d["hinweise"]]
    if hinweise:
        punkte.append({
            "art": "gelb",
            "titel": f"{len(hinweise)} Rechnung(en) mit Prüfhinweis",
            "text": "Rechenproben oder Datumsangaben sind auffällig. "
                    "Vor der Zahlung am Beleg prüfen – Einzelheiten auf dem "
                    "Blatt „Prüfliste“ im Excel-Report.",
        })

    lief = lieferanten(daten, w, grenze=0)
    gesamt = sum(d["brutto"] or 0 for d in teil) or 1.0
    if lief and lief[0]["Betrag"] / gesamt > 0.30:
        punkte.append({
            "art": "blau",
            "titel": f"{lief[0]['Lieferant']} deckt "
                     f"{lief[0]['Betrag'] / gesamt:.0%} des Einkaufsvolumens",
            "text": "Eine solche Konzentration ist ein Verhandlungshebel – "
                    "und zugleich ein Risiko, wenn der Lieferant ausfällt.",
        })

    verfallen = [d for d in teil if d["skonto_betrag"] and not d["skonto_offen"]]
    if verfallen:
        punkte.append({
            "art": "blau",
            "titel": f"Bei {len(verfallen)} Rechnung(en) ist die Skontofrist abgelaufen",
            "text": f"Rechnerisch {_zahl_kurz(sum(d['skonto_betrag'] for d in verfallen))} "
                    f"{w} liegen gelassen. Ob tatsächlich gezahlt wurde, sagt "
                    f"diese Auswertung nicht – aber es lohnt der Blick auf den "
                    f"Zahlungslauf.",
        })

    return punkte


def render_dashboard(invs: list[dict]) -> None:
    """Zeichnet das Dashboard. Erwartet die Rechnungen wie in der App."""
    import altair as alt
    import pandas as pd
    import streamlit as st

    if not invs:
        st.info("Noch keine Rechnungen ausgelesen – im Reiter „Hochladen“ starten.")
        return

    daten = aufbereiten(invs)
    whrs = waehrungen_von(daten)

    if len(whrs) > 1:
        c1, c2 = st.columns([1.2, 3])
        w = c1.radio("Währung", whrs, horizontal=True, key="dash_waehrung")
        c2.caption("Beträge werden nicht umgerechnet – jede Währung wird "
                   "für sich ausgewertet.")
    else:
        w = whrs[0]

    k = kennzahlen(daten, w)

    # ---- Kennzahlen -----------------------------------------------------
    s1, s2, s3, s4 = st.columns(4, gap="large")
    for spalte, (titel, wert, unter) in zip(
            (s1, s2, s3, s4),
            ((f"Offene Posten gesamt ({w})", _zahl_kurz(k["gesamt"]),
              f"{k['anzahl']} Rechnung(en)"),
             (f"Davon überfällig ({w})", _zahl_kurz(k["ueberfaellig"]),
              f"{k['anzahl_ueberfaellig']} Rechnung(en)"),
             (f"Fällig in 14 Tagen ({w})", _zahl_kurz(k["faellig14"]),
              "nächste zwei Wochen"),
             (f"Skonto erreichbar ({w})", _zahl_kurz(k["skonto_offen"]),
              (f"verfallen: {_zahl_kurz(k['skonto_verfallen'])} {w}"
               if k["skonto_verfallen"] else "Frist noch offen")))):
        with spalte, st.container(border=True):
            st.metric(titel, wert)
            st.caption(unter)

    if k["mit_hinweis"]:
        st.warning(
            f"{k['mit_hinweis']} von {k['anzahl']} Rechnung(en) haben einen "
            f"Prüfhinweis. Die Hinweise stammen aus festen Rechenregeln "
            f"(Rechenprobe, Datumslogik, Vollständigkeit) – es wurde nichts korrigiert."
        )

    st.divider()

    grafiken: list = []   # für den Bild-Export

    achse = alt.Axis(labelColor="#475569", titleColor="#475569",
                     grid=True, gridColor=HELL, domainColor=HELL, tickColor=HELL)

    # ---- Reihe 1: Fälligkeit | Lieferantenkonzentration -----------------
    st.write("")
    l, r = st.columns(2, gap="large")

    with l, st.container(border=True):
        st.markdown("**Offene Posten nach Fälligkeit**")
        df = pd.DataFrame(faelligkeitsstruktur(daten, w))
        reihenfolge = df["Stufe"].tolist()
        grafik = (
            alt.Chart(df)
            .mark_bar(cornerRadiusTopLeft=3, cornerRadiusTopRight=3)
            .encode(
                x=alt.X("Stufe:N", sort=reihenfolge, title=None,
                        axis=alt.Axis(labelAngle=-35, labelColor="#475569",
                                      domainColor=HELL, tickColor=HELL)),
                y=alt.Y("Betrag:Q", title=w, axis=achse),
                color=alt.Color("Stufe:N", sort=reihenfolge, legend=None,
                                scale=alt.Scale(domain=list(STUFEN_FARBEN),
                                                range=list(STUFEN_FARBEN.values()))),
                tooltip=["Stufe", "Belege", alt.Tooltip("Betrag:Q", format=",.2f")],
            )
            .properties(height=300)
        )
        _zeigen(st, grafiken, "Offene Posten nach Fälligkeit", grafik)

    with r, st.container(border=True):
        st.markdown("**Lieferantenkonzentration**")
        lief = lieferanten(daten, w)
        if lief:
            df = pd.DataFrame(lief)
            ordnung = df["Lieferant"].tolist()
            balken = (
                alt.Chart(df)
                .mark_bar(color=BLUE, cornerRadiusTopLeft=3, cornerRadiusTopRight=3)
                .encode(
                    x=alt.X("Lieferant:N", sort=ordnung, title=None,
                            axis=alt.Axis(labelAngle=-35, labelColor="#475569",
                                          domainColor=HELL, tickColor=HELL)),
                    y=alt.Y("Betrag:Q", title=w, axis=achse),
                    tooltip=["Lieferant", "Klasse", "Belege",
                             alt.Tooltip("Betrag:Q", format=",.2f")],
                )
            )
            linie = (
                alt.Chart(df)
                .mark_line(color=GOLD, point=alt.OverlayMarkDef(color=GOLD, size=40))
                .encode(
                    x=alt.X("Lieferant:N", sort=ordnung, title=None),
                    y=alt.Y("Kumuliert:Q", title="kumuliert",
                            axis=alt.Axis(format="%", labelColor=GOLD,
                                          titleColor=GOLD, grid=False)),
                    tooltip=[alt.Tooltip("Kumuliert:Q", format=".0%")],
                )
            )
            _zeigen(st, grafiken, "Lieferantenkonzentration",
                    alt.layer(balken, linie).resolve_scale(y="independent")
                    .properties(height=300))
            a = [x for x in lief if x["Klasse"] == "A"]
            if a:
                anteil = sum(x["Betrag"] for x in a) / (k["gesamt"] or 1)
                st.caption(f"{len(a)} Lieferant(en) machen {anteil:.0%} des Volumens aus "
                           f"(Klasse A).")

    # ---- Reihe 2: Monatsverlauf | Belegsgrößen --------------------------
    st.write("")
    l, r = st.columns(2, gap="large")

    with l, st.container(border=True):
        st.markdown("**Rechnungseingang je Monat**")
        mon = monatsverlauf(daten, w)
        if len(mon) >= 2:
            df = pd.DataFrame(mon)
            grafik = (
                alt.Chart(df)
                .mark_line(color="#7E3AF2", strokeWidth=2.5,
                           point=alt.OverlayMarkDef(color="#7E3AF2", size=55))
                .encode(
                    x=alt.X("Monat:N", title=None, axis=achse),
                    y=alt.Y("Betrag:Q", title=w, axis=achse),
                    tooltip=["Monat", "Belege", alt.Tooltip("Betrag:Q", format=",.2f")],
                )
                .properties(height=290)
            )
            _zeigen(st, grafiken, "Rechnungseingang je Monat", grafik)
        else:
            st.caption("Zu wenige Monate für einen Verlauf.")

    with r, st.container(border=True):
        st.markdown("**Belegsgrößen**")
        df = pd.DataFrame(belegsgroessen(daten, w))
        ordnung = df["Klasse"].tolist()
        grafik = (
            alt.Chart(df)
            .mark_bar(cornerRadiusTopLeft=3, cornerRadiusTopRight=3)
            .encode(
                x=alt.X("Klasse:N", sort=ordnung, title=f"Betrag in {w}",
                        axis=alt.Axis(labelAngle=-35, labelColor="#475569",
                                      domainColor=HELL, tickColor=HELL)),
                y=alt.Y("Belege:Q", title="Belege", axis=achse),
                color=alt.Color("Belege:Q", legend=None,
                                scale=alt.Scale(scheme="blues")),
                tooltip=["Klasse", "Belege", alt.Tooltip("Betrag:Q", format=",.2f")],
            )
            .properties(height=290)
        )
        _zeigen(st, grafiken, "Belegsgrößen", grafik)
        klein = sum(x["Belege"] for x in belegsgroessen(daten, w)[:2])
        if k["anzahl"]:
            st.caption(f"{klein} von {k['anzahl']} Belegen liegen unter 500 {w} – "
                       f"der Erfassungsaufwand ist je Beleg gleich.")

    # ---- Reihe 3: Skonto | Zahlungsziele --------------------------------
    st.write("")
    l, r = st.columns(2, gap="large")

    with l, st.container(border=True):
        st.markdown("**Skonto-Status**")
        df = pd.DataFrame(skonto_status(daten, w))
        grafik = (
            alt.Chart(df)
            .mark_arc(innerRadius=62, stroke="#FFFFFF", strokeWidth=2)
            .encode(
                theta=alt.Theta("Belege:Q"),
                color=alt.Color(
                    "Status:N",
                    scale=alt.Scale(
                        domain=["Frist läuft noch", "Frist abgelaufen",
                                "Kein Skonto angeboten"],
                        range=[GREEN, RED, GREY]),
                    legend=alt.Legend(title=None, orient="bottom"),
                ),
                tooltip=["Status", "Belege", alt.Tooltip("Betrag:Q", format=",.2f")],
            )
            .properties(height=290)
        )
        _zeigen(st, grafiken, "Skonto-Status", grafik)
        if k["skonto_verfallen"]:
            st.caption(f"Rechnerisch entgangen: {_geld(k['skonto_verfallen'], w)}. "
                       f"Ob gezahlt wurde, sagt diese Auswertung nicht.")

    with r, st.container(border=True):
        st.markdown("**Zahlungsziele je Lieferant**")
        zz = [x for x in lieferanten(daten, w) if x["Zahlungsziel"] is not None]
        if zz:
            df = pd.DataFrame(zz)
            grafik = (
                alt.Chart(df)
                .mark_bar(color="#0694A2", cornerRadiusTopRight=3,
                          cornerRadiusBottomRight=3)
                .encode(
                    y=alt.Y("Lieferant:N", sort="-x", title=None, axis=achse),
                    x=alt.X("Zahlungsziel:Q", title="Tage", axis=achse),
                    tooltip=["Lieferant", "Belege",
                             alt.Tooltip("Zahlungsziel:Q", format=".1f")],
                )
                .properties(height=290)
            )
            _zeigen(st, grafiken, "Zahlungsziele je Lieferant", grafik)
        else:
            st.caption("Keine Rechnung mit Rechnungs- und Fälligkeitsdatum.")

    # ---- Reihe 4: Artikel | dringende Posten ----------------------------
    st.write("")
    l, r = st.columns(2, gap="large")

    with l, st.container(border=True):
        st.markdown("**Größte Positionen nach Artikel**")
        art = artikel(daten, w)
        if art:
            df = pd.DataFrame(art)
            grafik = (
                alt.Chart(df)
                .mark_bar(color="#0E9F6E", cornerRadiusTopRight=3,
                          cornerRadiusBottomRight=3)
                .encode(
                    y=alt.Y("Bezeichnung:N", sort="-x", title=None, axis=achse),
                    x=alt.X("Betrag:Q", title=f"{w}, netto", axis=achse),
                    tooltip=["Bezeichnung", "Posten", "Lieferanten",
                             alt.Tooltip("Betrag:Q", format=",.2f")],
                )
                .properties(height=290)
            )
            _zeigen(st, grafiken, "Größte Positionen nach Artikel", grafik)
            mehrere = [x for x in art if x["Lieferanten"] > 1]
            if mehrere:
                st.caption(f"{len(mehrere)} Artikel kommen von mehr als einem "
                           f"Lieferanten – dort lohnt ein Preisvergleich.")
        else:
            st.caption("Keine Einzelpositionen ausgelesen.")

    with r, st.container(border=True):
        st.markdown("**Bald fällig oder überfällig**")
        posten = [d for d in dringend(daten) if d["waehrung"] == w]
        if not posten:
            st.caption("Keine Rechnung in den nächsten 14 Tagen fällig.")
        for d in posten:
            t = d["tage"]
            status = (f"überfällig seit {abs(t)} Tagen" if t < 0
                      else ("heute fällig" if t == 0 else f"fällig in {t} Tagen"))
            farbe = RED if t < 0 else NAVY
            st.markdown(
                f"<div style='font-size:0.88rem;color:{farbe};padding:2px 0;'>"
                f"{d['inv'].get('rechnungsnummer') or d['inv'].get('datei')} · "
                f"{d['lieferant']} · {_geld(d['brutto'], w)} · {status}</div>",
                unsafe_allow_html=True,
            )

    # ---- Was als Nächstes zu tun ist -----------------------------------
    st.write("")
    st.markdown("### Was als Nächstes zu tun ist")
    punkte = handlungsempfehlungen(daten, w)
    if not punkte:
        st.success("Nichts Dringendes: keine überfälligen Posten, keine offenen "
                   "Skontofristen, keine Prüfhinweise.")
    else:
        farben = {"rot": RED, "gelb": GOLD, "gruen": GREEN, "blau": BLUE}
        spalten = st.columns(2, gap="large")
        for i, punkt in enumerate(punkte):
            with spalten[i % 2], st.container(border=True):
                st.markdown(
                    f"<div style='color:{farben[punkt['art']]};font-weight:600;"
                    f"font-size:0.98rem;margin-bottom:4px;'>{punkt['titel']}</div>"
                    f"<div style='color:#475569;font-size:0.88rem;'>"
                    f"{punkt['text']}</div>",
                    unsafe_allow_html=True,
                )
        st.caption("Alle Angaben sind aus den ausgelesenen Belegen berechnet. "
                   "Zahlungsstände sind nicht bekannt – die Hinweise ersetzen "
                   "keinen Abgleich mit dem Bankkonto.")

    # Grafiken für den Export in der Seitenleiste bereitstellen
    st.session_state["dash_grafiken"] = grafiken
    st.session_state["dash_kopf"] = (
        f"Stand {datetime.now().strftime('%d.%m.%Y, %H:%M')} Uhr · "
        f"{k['anzahl']} Rechnung(en) in {w}"
    )
    st.session_state["dash_export_waehrung"] = w
    # Kennung der Datenlage: ändert sie sich, wird ein erzeugtes Bild
    # verworfen – so bleibt es bei genau einem Knopf.
    st.session_state["dash_fingerabdruck"] = (len(invs), w, k["gesamt"])


def kurzueberblick(invs: list[dict]) -> None:
    """Kompakte Kennzahlen für den Hochladen-Reiter – je Währung eine Zeile."""
    import streamlit as st

    if not invs:
        return
    daten = aufbereiten(invs)
    for w in waehrungen_von(daten):
        k = kennzahlen(daten, w)
        s1, s2, s3, s4 = st.columns(4, gap="large")
        with s1, st.container(border=True):
            st.metric(f"Offene Posten ({w})", _zahl_kurz(k["gesamt"]))
            st.caption(f"{k['anzahl']} Rechnung(en)")
        with s2, st.container(border=True):
            st.metric(f"Überfällig ({w})", _zahl_kurz(k["ueberfaellig"]))
            st.caption(f"{k['anzahl_ueberfaellig']} Rechnung(en)")
        with s3, st.container(border=True):
            st.metric(f"Fällig in 14 Tagen ({w})", _zahl_kurz(k["faellig14"]))
            st.caption("nächste zwei Wochen")
        with s4, st.container(border=True):
            st.metric("Mit Prüfhinweis", str(k["mit_hinweis"]))
            st.caption(f"von {k['anzahl']} Rechnung(en)")
