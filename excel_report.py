"""Excel-Report für die LN Rechnungserfassung.

Blätter:
  1 Übersicht      Kennzahlen, Fälligkeitsstruktur, drei Kurzlisten
  2 Rechnungen     eine Zeile je Beleg, filterbar
  3 Positionen     flache Tabelle, pivotierbar
  4 Prüfliste      nur die Belege mit Auffälligkeiten
  5 USt-Auswertung Netto und Steuer je Satz

Öffentliche Schnittstelle unverändert:
    build_excel, normalisiere, parse_date, pruefhinweise
"""
from __future__ import annotations

import io
import re
from datetime import datetime, timedelta

NAVY = "0F172A"
BLUE = "1A56DB"
GOLD = "B48A2F"
LIGHT = "F3F6FB"
GREY = "64748B"
RED = "B91C1C"
GREEN = "15803D"
LINE = "D8DEE9"

# Farben für Diagramme: kräftig, aber aus einer Familie
# Jeder Diagrammblock ist gleich hoch – sonst stehen die Grafiken krumm
BLOCK = 20

SERIE = ["1A56DB", "0E9F6E", "B48A2F", "7E3AF2", "E02424", "0694A2",
         "FF8A4C", "5145CD", "057A55", "9F580A"]
AMPEL = {"rot": "E02424", "gelb": "D9A404", "gruen": "0E9F6E", "grau": "9CA3AF"}

RECHNUNG_SPALTEN = [
    ("rechnungsnummer", "Rechnungs-Nr.", 16),
    ("lieferant", "Lieferant", 30),
    ("rechnungsdatum", "Rechnungsdatum", 15),
    ("faelligkeitsdatum", "Fällig am", 13),
    ("faelligkeit_quelle", "Fälligkeit", 12),
    ("tage_bis_faellig", "Tage", 8),
    ("stufe", "Fälligkeitsstufe", 18),
    ("netto_gesamt", "Netto", 13),
    ("ust_satz_prozent", "USt %", 7),
    ("ust_betrag", "USt-Betrag", 13),
    ("brutto_gesamt", "Brutto", 14),
    ("waehrung", "Währung", 9),
    ("skonto_prozent", "Skonto %", 9),
    ("skonto_tage", "Skontofrist", 11),
    ("skonto_bis", "Skonto bis", 12),
    ("skonto_betrag", "Skontobetrag", 13),
    ("anzahl_positionen", "Positionen", 11),
    ("iban", "IBAN", 26),
    ("lieferant_ust_id", "USt-ID Lieferant", 18),
    ("zahlungsbedingungen", "Zahlungsbedingungen", 34),
    ("lieferbedingungen", "Lieferbedingungen", 26),
    ("datei", "Datei", 28),
    ("pruefhinweis", "Prüfhinweis", 46),
]

POSITION_SPALTEN = [
    ("rechnungsnummer", "Rechnungs-Nr.", 16),
    ("lieferant", "Lieferant", 30),
    ("rechnungsdatum", "Rechnungsdatum", 15),
    ("pos_nr", "Pos", 6),
    ("bezeichnung", "Bezeichnung", 46),
    ("menge", "Menge", 10),
    ("einheit", "Einheit", 9),
    ("einzelpreis_netto", "Einzelpreis (netto)", 17),
    ("gesamt_netto", "Gesamt (netto)", 15),
    ("waehrung", "Währung", 9),
    ("ust_satz_prozent", "USt %", 7),
]

# Formulierungen, bei denen die Rechnung am Rechnungsdatum fällig ist
SOFORT_MARKER = (
    "sofort fällig", "sofort faellig", "sofort zahlbar", "zahlbar sofort",
    "sofort ohne abzug", "zahlung sofort", "sofort rein netto", "sofort netto",
    "zahlbar bei erhalt", "bei erhalt fällig", "due immediately", "payable immediately",
    "due upon receipt", "payable on receipt",
)

# Fälligkeitsstufen: (Schlüssel, Beschriftung, von_Tagen, bis_Tagen)
# Tage = faellig - heute. Negativ heißt überfällig.
STUFEN = [
    ("ueber_30", "Überfällig > 30 Tage", None, -31),
    ("ueber_1", "Überfällig 1–30 Tage", -30, -1),
    ("heute_7", "Fällig in 0–7 Tagen", 0, 7),
    ("tage_14", "Fällig in 8–14 Tagen", 8, 14),
    ("tage_30", "Fällig in 15–30 Tagen", 15, 30),
    ("spaeter", "Fällig in über 30 Tagen", 31, None),
]
STUFE_OHNE = "Ohne Fälligkeit"


# --------------------------------------------------------------- Hilfsmittel

def _zahl(v):
    """Robust nach float – akzeptiert auch '1.234,56 €'. Sonst None."""
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    t = str(v).replace("€", "").replace("EUR", "").replace("$", "")
    t = t.replace("USD", "").replace("%", "").replace(" ", "").strip()
    if "," in t and "." in t:
        t = t.replace(".", "").replace(",", ".")
    elif "," in t:
        t = t.replace(",", ".")
    try:
        return float(t)
    except ValueError:
        return None


def parse_date(s):
    if not s:
        return None
    if isinstance(s, datetime):
        return s
    for fmt in ("%d.%m.%Y", "%d.%m.%y", "%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(str(s).strip(), fmt)
        except ValueError:
            continue
    return None


_RE_PROZENT = re.compile(r"(\d{1,2}(?:[.,]\d{1,2})?)\s*(?:%|prozent|percent)", re.I)
_RE_TAGE = re.compile(r"(\d{1,3})\s*(?:tage?n?|days?)", re.I)
_RE_DATUM = re.compile(r"\b(\d{1,2}\.\d{1,2}\.\d{2,4})\b")


def skonto_werte(inv: dict):
    """Skontosatz, Skontofrist und Stichtag aus dem Freitext lösen.

    Liefert (prozent, tage, stichtag). Jeder Wert kann None sein.
    Es wird nichts geraten: ohne klare Prozentangabe gibt es kein Skonto.
    """
    p = _zahl(inv.get("skonto_prozent"))
    t = _zahl(inv.get("skonto_tage"))
    if p is not None:
        return p, (int(t) if t is not None else None), None

    text = " ".join(str(inv.get(k) or "") for k in ("skonto", "zahlungsbedingungen"))
    if not text.strip():
        return None, None, None
    low = text.lower()
    if "skonto" not in low:
        return None, None, None

    mp = _RE_PROZENT.search(text)
    if not mp:
        return None, None, None
    prozent = _zahl(mp.group(1))
    if prozent is None or not (0 < prozent <= 15):
        return None, None, None

    # Tagesangaben sammeln, aber die Nettofrist aussortieren:
    # "… 30 Tage netto" und "sonst 30 Tage" meinen das Zahlungsziel,
    # nicht die Skontofrist.
    kandidaten = []
    for m in _RE_TAGE.finditer(text):
        n = int(m.group(1))
        if n > 120:
            continue
        danach = text[m.end():m.end() + 14].lower().lstrip(" ,.-")
        davor = text[max(0, m.start() - 14):m.start()].lower()
        if danach.startswith("netto") or danach.startswith("rein netto"):
            continue
        if "sonst" in davor or "danach" in davor or "ansonsten" in davor:
            continue
        kandidaten.append((abs(m.start() - mp.start()), n))

    tage = None
    if kandidaten:
        kandidaten.sort()
        tage = kandidaten[0][1]

    # Stichtag ("… bei Zahlung bis 19.07.2026 …") schlägt die Tagesangabe
    stichtag = None
    md = _RE_DATUM.search(text)
    if md:
        stichtag = parse_date(md.group(1))

    return prozent, tage, stichtag


def normalisiere(inv: dict) -> dict:
    """Fehlende Fälligkeit aus 'sofort fällig' o. ä. ableiten.
    Es wird nichts erfunden: nur wenn der Beleg das eindeutig sagt."""
    if not inv.get("faelligkeitsdatum") and inv.get("rechnungsdatum"):
        text = " ".join(
            str(inv.get(k) or "") for k in ("zahlungsbedingungen", "skonto", "lieferbedingungen")
        ).lower()
        if any(m in text for m in SOFORT_MARKER):
            inv["faelligkeitsdatum"] = inv["rechnungsdatum"]
            inv["faelligkeit_abgeleitet"] = True
    return inv


def pruefhinweise(inv: dict) -> list[str]:
    """Auffälligkeiten, die ein Mensch am Beleg prüfen sollte.
    Es wird nichts korrigiert – nur markiert."""
    h: list[str] = []
    heute = datetime.now()

    netto = _zahl(inv.get("netto_gesamt"))
    ust = _zahl(inv.get("ust_betrag"))
    brutto = _zahl(inv.get("brutto_gesamt"))
    satz = _zahl(inv.get("ust_satz_prozent"))

    if netto is not None and ust is not None and brutto is not None:
        if abs(netto + ust - brutto) > 0.02:
            h.append(f"Rechenprobe: Netto + USt ergibt {_de(netto + ust)}, "
                     f"Beleg nennt {_de(brutto)}")

    if netto is not None and ust is not None and satz:
        erwartet = netto * satz / 100
        if abs(erwartet - ust) > max(0.02, abs(netto) * 0.002):
            h.append(f"USt-Betrag passt nicht zu {satz:g} % (erwartet {_de(erwartet)})")

    posten = [_zahl(p.get("gesamt_netto")) for p in (inv.get("positionen") or [])]
    if posten and all(x is not None for x in posten) and netto is not None:
        if abs(sum(posten) - netto) > 0.02:
            h.append(f"Positionen ergeben {_de(sum(posten))} netto, "
                     f"Beleg nennt {_de(netto)}")

    rd = parse_date(inv.get("rechnungsdatum"))
    fd = parse_date(inv.get("faelligkeitsdatum"))
    if rd is None:
        h.append("Kein Rechnungsdatum erkannt")
    elif rd > heute + timedelta(days=365):
        h.append(f"Rechnungsdatum {rd.strftime('%d.%m.%Y')} liegt weit in der Zukunft "
                 f"– am Beleg prüfen")
    elif rd < heute - timedelta(days=1825):
        h.append(f"Rechnungsdatum {rd.strftime('%d.%m.%Y')} liegt mehr als 5 Jahre zurück")

    if fd is None:
        h.append("Keine Fälligkeit erkannt – Zahlungsziel am Beleg prüfen")
    elif rd is not None and fd < rd:
        h.append("Fälligkeit liegt vor dem Rechnungsdatum")

    if inv.get("faelligkeit_abgeleitet"):
        h.append("Fälligkeit aus „sofort fällig“ abgeleitet (= Rechnungsdatum)")

    if brutto is None:
        h.append("Kein Bruttobetrag erkannt")
    if not inv.get("lieferant"):
        h.append("Kein Lieferant erkannt")

    return h


def _de(v):
    """Zahl in deutscher Schreibweise, ohne Währung."""
    if v is None:
        return "–"
    return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _fmt_eur(v, waehrung="EUR"):
    if v is None:
        return "–"
    return f"{_de(v)} {waehrung or ''}".strip()


def _geldformat(w):
    w = (w or "EUR").replace('"', "")
    return f'#,##0.00 "{w}"'


def _stufe(tage):
    if tage is None:
        return STUFE_OHNE
    for _, label, von, bis in STUFEN:
        if von is not None and tage < von:
            continue
        if bis is not None and tage > bis:
            continue
        return label
    return STUFE_OHNE


def _nur_prozent():
    """Datenbeschriftung, die wirklich nur den Prozentwert zeigt."""
    from openpyxl.chart.label import DataLabelList
    dl = DataLabelList()
    dl.showPercent = True
    dl.showVal = False
    dl.showCatName = False
    dl.showSerName = False
    dl.showLegendKey = False
    dl.showBubbleSize = False
    return dl


def _faerbe_serie(chart, farben):
    """Datenreihen einfärben, damit die Grafik nicht nach Vorlage aussieht."""
    from openpyxl.chart.shapes import GraphicalProperties
    for i, serie in enumerate(chart.series):
        farbe = farben[i % len(farben)]
        serie.graphicalProperties = GraphicalProperties(solidFill=farbe)
        serie.graphicalProperties.line.solidFill = farbe


def _faerbe_punkte(chart, farben):
    """Einzelne Segmente oder Säulen getrennt einfärben."""
    from openpyxl.chart.marker import DataPoint
    from openpyxl.chart.shapes import GraphicalProperties
    if not chart.series:
        return
    punkte = []
    for i, farbe in enumerate(farben):
        dp = DataPoint(idx=i)
        dp.graphicalProperties = GraphicalProperties(solidFill=farbe)
        dp.graphicalProperties.line.solidFill = "FFFFFF"
        punkte.append(dp)
    chart.series[0].data_points = punkte


def _schlicht(chart, gitter=True):
    """Achsen sichtbar erzwingen und das Gitternetz beruhigen.

    Ohne delete=False blendet Excel beide Achsen aus – LibreOffice zeigt sie
    trotzdem, deshalb fällt der Fehler beim Testen leicht durch.
    """
    for achse in (getattr(chart, "x_axis", None), getattr(chart, "y_axis", None)):
        if achse is None:
            continue
        achse.delete = False
        achse.majorTickMark = "out"
    if not gitter:
        try:
            chart.y_axis.majorGridlines = None
        except Exception:
            pass
    chart.style = None


def _blocktitel(ws, zeile, text, zusatz=None):
    from openpyxl.styles import Font
    ws.cell(row=zeile, column=2, value=text).font = Font(size=11, bold=True, color=NAVY)
    if zusatz:
        ws.cell(row=zeile, column=4, value=zusatz).font = Font(size=8, color=GREY)
    return zeile + 1


def _textblock_dringend(ws, zeile, daten, heute_tag):
    """Was in den nächsten 14 Tagen zu zahlen ist oder schon überfällig war."""
    from openpyxl.styles import Font
    zeile = _blocktitel(ws, zeile, "Bald fällig oder überfällig",
                        "die sechs dringendsten Posten")
    dringend = sorted([d for d in daten if d["tage"] is not None and d["tage"] <= 14],
                      key=lambda d: d["tage"])[:6]
    if not dringend:
        ws.cell(row=zeile, column=2,
                value="Keine Rechnung in den nächsten 14 Tagen fällig."
                ).font = Font(size=10, color=GREY)
        return zeile + 3
    for d in dringend:
        t = d["tage"]
        status = (f"überfällig seit {abs(t)} Tagen" if t < 0
                  else ("heute fällig" if t == 0 else f"fällig in {t} Tagen"))
        ws.cell(row=zeile, column=2,
                value=f"{d['inv'].get('rechnungsnummer') or d['inv'].get('datei')} · "
                      f"{d['inv'].get('lieferant') or 'Ohne Lieferant'} · "
                      f"{_fmt_eur(d['brutto'], d['waehrung'])} · {status} "
                      f"({d['fd'].strftime('%d.%m.%Y')})"
                ).font = Font(size=9, color=RED if t < 0 else NAVY)
        zeile += 1
    return zeile + 2


def _textblock_hohe_betraege(ws, zeile, daten, leit):
    """Belege, die deutlich über dem Schnitt liegen – lohnt einen zweiten Blick."""
    from openpyxl.styles import Font
    teil = [d for d in daten if d["waehrung"] == leit and d["brutto"]]
    if len(teil) < 5:
        return zeile
    schnitt = sum(abs(d["brutto"]) for d in teil) / len(teil)
    auffaellig = sorted([d for d in teil if abs(d["brutto"]) > schnitt * 2],
                        key=lambda d: -abs(d["brutto"]))[:4]
    zeile = _blocktitel(ws, zeile, "Auffällig hohe Beträge",
                        f"mehr als das Doppelte des Schnitts "
                        f"({_fmt_eur(schnitt, leit)})")
    if not auffaellig:
        ws.cell(row=zeile, column=2,
                value="Keine Ausreißer – die Beträge liegen eng beieinander."
                ).font = Font(size=10, color=GREY)
        return zeile + 3
    for d in auffaellig:
        faktor = abs(d["brutto"]) / schnitt
        ws.cell(row=zeile, column=2,
                value=f"{d['inv'].get('rechnungsnummer') or d['inv'].get('datei')} · "
                      f"{d['inv'].get('lieferant') or 'Ohne Lieferant'} · "
                      f"{_fmt_eur(d['brutto'], d['waehrung'])} "
                      f"({faktor:.1f}-fach über dem Schnitt)"
                ).font = Font(size=9, color=GOLD)
        zeile += 1
    return zeile + 2


def _textblock_pruefen(ws, zeile, daten, leit, mit_hinweis):
    """Kurzfassung der Prüfliste – Einzelheiten stehen auf dem eigenen Blatt."""
    from openpyxl.styles import Font
    betrag = sum(d["brutto"] or 0 for d in mit_hinweis if d["waehrung"] == leit)
    anteil = (len(mit_hinweis) / len(daten) * 100) if daten else 0
    zeile = _blocktitel(ws, zeile, "Bitte am Beleg prüfen",
                        "Einzelheiten auf dem Blatt „Prüfliste“")
    if not mit_hinweis:
        ws.cell(row=zeile, column=2,
                value="Keine Auffälligkeiten – Rechenproben und Daten sind plausibel."
                ).font = Font(size=10, color=GREEN)
        return zeile + 2
    ws.cell(row=zeile, column=2,
            value=f"{len(mit_hinweis)} von {len(daten)} Belegen ({anteil:.0f} %) "
                  f"haben einen Hinweis, Volumen {_fmt_eur(betrag, leit)}. "
                  f"Es wurde nichts korrigiert, nur markiert."
            ).font = Font(size=10, color=RED)
    zeile += 1
    gezaehlt = {}
    for d in mit_hinweis:
        for h in d["hinweise"]:
            kurz = h.split(":")[0].split("–")[0].strip()
            gezaehlt[kurz] = gezaehlt.get(kurz, 0) + 1
    for kurz, n in sorted(gezaehlt.items(), key=lambda x: -x[1])[:4]:
        ws.cell(row=zeile, column=2, value=f"{n}× {kurz}").font = Font(
            size=9, color=NAVY)
        zeile += 1
    return zeile + 2


# ------------------------------------------------------------------- Aufbau

def build_excel(invoices: list[dict]) -> bytes:
    from openpyxl import Workbook
    from openpyxl.chart import BarChart, DoughnutChart, LineChart, PieChart, Reference
    from openpyxl.chart.label import DataLabelList
    from openpyxl.formatting.rule import DataBarRule
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    heute = datetime.now()
    heute_tag = heute.replace(hour=0, minute=0, second=0, microsecond=0)

    thin = Side(style="thin", color=LINE)
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    unten = Border(bottom=Side(style="thin", color=LINE))
    head_fill = PatternFill("solid", fgColor=NAVY)
    head_font = Font(color="FFFFFF", bold=True, size=10)
    kpi_fill = PatternFill("solid", fgColor=LIGHT)
    zebra = PatternFill("solid", fgColor="F8FAFC")
    # Deutlicher als das feine Zebra auf dem Rechnungsblatt: hier soll man
    # auf einen Blick sehen, welche Positionen zu einer Rechnung gehören.
    beleg_streifen = PatternFill("solid", fgColor="E8EFF9")

    # ---- Daten einmal aufbereiten, danach nur noch lesen -------------------
    daten = []
    for inv in invoices:
        normalisiere(inv)
        rd = parse_date(inv.get("rechnungsdatum"))
        fd = parse_date(inv.get("faelligkeitsdatum"))
        tage = (fd - heute_tag).days if fd else None
        sp, st, stichtag = skonto_werte(inv)
        if stichtag is not None:
            skonto_bis = stichtag
            if rd is not None and st is None:
                st = (stichtag - rd).days
        elif rd is not None and st is not None:
            skonto_bis = rd + timedelta(days=st)
        else:
            skonto_bis = None
        brutto = _zahl(inv.get("brutto_gesamt"))
        skonto_betrag = (brutto * sp / 100) if (brutto is not None and sp) else None
        daten.append({
            "inv": inv,
            "rd": rd, "fd": fd, "tage": tage,
            "stufe": _stufe(tage),
            "waehrung": inv.get("waehrung") or "EUR",
            "netto": _zahl(inv.get("netto_gesamt")),
            "ust": _zahl(inv.get("ust_betrag")),
            "brutto": brutto,
            "satz": _zahl(inv.get("ust_satz_prozent")),
            "skonto_prozent": sp,
            "skonto_tage": st,
            "skonto_bis": skonto_bis,
            "skonto_betrag": skonto_betrag,
            "skonto_offen": bool(skonto_bis and skonto_bis >= heute_tag),
            "hinweise": pruefhinweise(inv),
            "positionen": inv.get("positionen") or [],
        })

    waehrungen = sorted({d["waehrung"] for d in daten}) or ["EUR"]
    leit = "EUR" if "EUR" in waehrungen else waehrungen[0]

    # ======================= Blatt 1: Übersicht =============================
    # Ziel: alles auf einen Blick, ohne herauszoomen zu müssen.
    # Links die Textlisten, rechts die Fälligkeitstabelle und eine Grafik.
    ws = wb.active
    ws.title = "Übersicht"
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1

    breiten = {"A": 2, "B": 23, "C": 23, "D": 23, "E": 23, "F": 3}
    for sp_, br in breiten.items():
        ws.column_dimensions[sp_].width = br
    # G/H und L/M rahmen die Tabelle ein, damit sie mittig unter der
    # Grafik sitzt. I/J/K tragen die Tabelle selbst.
    for sp_, br in {"G": 10, "H": 10, "I": 24, "J": 8, "K": 16,
                    "L": 10, "M": 10, "N": 11}.items():
        ws.column_dimensions[sp_].width = br

    ws["B2"] = "Rechnungsübersicht"
    ws["B2"].font = Font(size=18, bold=True, color=NAVY)
    ws["B3"] = (f"LN Automation · Stand {heute.strftime('%d.%m.%Y, %H:%M')} Uhr · "
                f"{len(daten)} Eingangsrechnung(en)")
    ws["B3"].font = Font(size=9, color=GREY)

    zeile = 5

    def kpi_block(z, eintraege):
        """Vier Kennzahlen nebeneinander, je eine Spalte. Echte Zahlen."""
        col = 2
        for label, wert, fmt, farbe in eintraege:
            for r in (z, z + 1):
                ws.cell(row=r, column=col).fill = kpi_fill
            ws.cell(row=z, column=col, value=label).font = Font(size=9, color=GREY)
            zelle = ws.cell(row=z + 1, column=col, value=wert)
            zelle.font = Font(size=16, bold=True, color=farbe)
            if fmt:
                zelle.number_format = fmt
            zelle.alignment = Alignment(vertical="center", horizontal="right")
            ws.row_dimensions[z + 1].height = 26
            col += 1
        return z + 2

    for w in waehrungen:
        teil_w = [d for d in daten if d["waehrung"] == w]
        summe = sum(d["brutto"] or 0 for d in teil_w)
        ueberfaellig = sum(d["brutto"] or 0 for d in teil_w
                           if d["tage"] is not None and d["tage"] < 0)
        faellig14 = sum(d["brutto"] or 0 for d in teil_w
                        if d["tage"] is not None and 0 <= d["tage"] <= 14)
        skonto_offen = sum(d["skonto_betrag"] or 0 for d in teil_w if d["skonto_offen"])

        if len(waehrungen) > 1:
            ws.cell(row=zeile, column=2, value=f"Währung {w}").font = Font(
                size=10, bold=True, color=NAVY)
            zeile += 1

        zeile = kpi_block(zeile, [
            ("Offene Posten gesamt (brutto)", summe, _geldformat(w), NAVY),
            ("Davon überfällig", ueberfaellig, _geldformat(w),
             RED if ueberfaellig else NAVY),
            ("Fällig in 14 Tagen", faellig14, _geldformat(w), NAVY),
            ("Skonto noch erreichbar", skonto_offen, _geldformat(w),
             GREEN if skonto_offen else GREY),
        ])

    if len(waehrungen) > 1:
        ws.cell(row=zeile, column=2,
                value="Mehrere Währungen – Summen stehen getrennt, keine Umrechnung."
                ).font = Font(size=8, italic=True, color=GREY)
    zeile += 4

    inhalt_start = zeile
    mit_hinweis = [d for d in daten if d["hinweise"]]

    # ---- Rechts: Grafik oben, Fälligkeitstabelle mittig darunter --------
    # Die Tabelle ist zugleich die Datengrundlage des Diagramms.
    TAB_SP = 9          # Spalte I – so liegt die Tabelle mittig unter der Grafik
    dia_zeile = 5       # Grafik beginnt auf Höhe der Kennzahlen
    tab_zeile = dia_zeile + 19

    ws.cell(row=tab_zeile, column=TAB_SP, value="Fälligkeitsstruktur").font = Font(
        size=11, bold=True, color=NAVY)
    if len(waehrungen) > 1:
        ws.cell(row=tab_zeile, column=TAB_SP + 2, value=f"nur {leit}").font = Font(
            size=8, color=GREY)

    kopf_stufe = tab_zeile + 1
    for versatz, txt in ((0, "Stufe"), (1, "Anzahl"), (2, "Betrag brutto")):
        z_ = ws.cell(row=kopf_stufe, column=TAB_SP + versatz, value=txt)
        z_.font = Font(size=9, bold=True, color="FFFFFF")
        z_.fill = head_fill
        z_.alignment = Alignment(horizontal="left" if versatz == 0 else "right")

    stufen_start = kopf_stufe + 1
    z_stufe = stufen_start
    for label in [s[1] for s in STUFEN] + [STUFE_OHNE]:
        teil_s = [d for d in daten if d["waehrung"] == leit and d["stufe"] == label]
        betrag = sum(d["brutto"] or 0 for d in teil_s)
        ws.cell(row=z_stufe, column=TAB_SP, value=label).font = Font(
            size=9, color=RED if label.startswith("Überfällig") else NAVY)
        ws.cell(row=z_stufe, column=TAB_SP + 1,
                value=len(teil_s)).alignment = Alignment(horizontal="right")
        bz = ws.cell(row=z_stufe, column=TAB_SP + 2, value=betrag)
        bz.number_format = _geldformat(leit)
        bz.alignment = Alignment(horizontal="right")
        for versatz in range(3):
            ws.cell(row=z_stufe, column=TAB_SP + versatz).border = unten
        z_stufe += 1
    stufen_ende = z_stufe - 1

    sp_betrag = get_column_letter(TAB_SP + 2)
    ws.conditional_formatting.add(
        f"{sp_betrag}{stufen_start}:{sp_betrag}{stufen_ende}",
        DataBarRule(start_type="num", start_value=0, end_type="max",
                    color=BLUE, showValue=True),
    )

    # Zusammenfassung direkt darunter
    zus = stufen_ende + 2
    ws.cell(row=zus, column=TAB_SP, value="Zusammenfassung").font = Font(
        size=9, bold=True, color=GREY)
    ueber_s = sum(d["brutto"] or 0 for d in daten
                  if d["waehrung"] == leit and d["tage"] is not None and d["tage"] < 0)
    offen_s = sum(d["brutto"] or 0 for d in daten
                  if d["waehrung"] == leit and d["tage"] is not None and d["tage"] >= 0)
    ohne_s = sum(d["brutto"] or 0 for d in daten
                 if d["waehrung"] == leit and d["tage"] is None)
    ges_s = (ueber_s + offen_s + ohne_s) or 1
    for i, (lab, wert, farbe) in enumerate(
            (("Überfällig", ueber_s, AMPEL["rot"]),
             ("Noch nicht fällig", offen_s, AMPEL["gruen"]),
             ("Ohne Fälligkeit", ohne_s, AMPEL["grau"])), start=1):
        ws.cell(row=zus + i, column=TAB_SP, value=lab).font = Font(size=9, color=farbe)
        az = ws.cell(row=zus + i, column=TAB_SP + 1, value=wert / ges_s)
        az.number_format = "0 %"
        az.alignment = Alignment(horizontal="right")
        bz = ws.cell(row=zus + i, column=TAB_SP + 2, value=wert)
        bz.number_format = _geldformat(leit)
        bz.alignment = Alignment(horizontal="right")
        for versatz in range(3):
            ws.cell(row=zus + i, column=TAB_SP + versatz).border = unten

    diagramm = BarChart()
    diagramm.type = "col"
    diagramm.title = f"Offene Posten nach Fälligkeit ({leit})"
    diagramm.legend = None
    diagramm.y_axis.numFmt = "#,##0"
    diagramm.gapWidth = 40
    diagramm.add_data(Reference(ws, min_col=TAB_SP + 2, min_row=kopf_stufe,
                                max_row=stufen_ende), titles_from_data=True)
    diagramm.set_categories(Reference(ws, min_col=TAB_SP, min_row=stufen_start,
                                      max_row=stufen_ende))
    _schlicht(diagramm)
    _faerbe_punkte(diagramm, [AMPEL["rot"], "F05252", AMPEL["gelb"], "E3B341",
                              "31C48D", AMPEL["gruen"], AMPEL["grau"]])
    diagramm.height, diagramm.width = 9.4, 15.5
    ws.add_chart(diagramm, f"G{dia_zeile}")

    # ---- Linke Spalte: die drei Textlisten, direkt unter den Kennzahlen ---
    zeile = _textblock_dringend(ws, inhalt_start, daten, heute_tag)
    zeile = _textblock_hohe_betraege(ws, zeile, daten, leit)
    zeile = _textblock_pruefen(ws, zeile, daten, leit, mit_hinweis)

    # ======================= Blatt 2: Rechnungen ============================
    ws2 = wb.create_sheet("Rechnungen")
    ws2.sheet_view.showGridLines = False
    for idx, (_, label, breite) in enumerate(RECHNUNG_SPALTEN, start=1):
        zelle = ws2.cell(row=1, column=idx, value=label)
        zelle.fill = head_fill
        zelle.font = head_font
        zelle.alignment = Alignment(vertical="center", wrap_text=True)
        zelle.border = border
        ws2.column_dimensions[get_column_letter(idx)].width = breite
    n_spalten = len(RECHNUNG_SPALTEN)
    ws2.row_dimensions[1].height = 30

    geld = {"netto_gesamt", "ust_betrag", "brutto_gesamt", "skonto_betrag"}
    datum = {"rechnungsdatum", "faelligkeitsdatum", "skonto_bis"}
    lang = {"zahlungsbedingungen", "lieferbedingungen", "datei", "pruefhinweis"}

    for r, d in enumerate(daten, start=2):
        inv = d["inv"]
        werte = {
            "lieferant": inv.get("lieferant"),
            "rechnungsnummer": inv.get("rechnungsnummer"),
            "rechnungsdatum": d["rd"],
            "faelligkeitsdatum": d["fd"],
            "faelligkeit_quelle": ("abgeleitet" if inv.get("faelligkeit_abgeleitet")
                                   else ("vom Beleg" if d["fd"] else "fehlt")),
            "tage_bis_faellig": d["tage"],
            "stufe": d["stufe"],
            "netto_gesamt": d["netto"],
            "ust_satz_prozent": d["satz"],
            "ust_betrag": d["ust"],
            "brutto_gesamt": d["brutto"],
            "waehrung": d["waehrung"],
            "skonto_prozent": d["skonto_prozent"],
            "skonto_tage": d["skonto_tage"],
            "skonto_bis": d["skonto_bis"],
            "skonto_betrag": d["skonto_betrag"],
            "anzahl_positionen": len(d["positionen"]) or None,
            "iban": inv.get("iban"),
            "lieferant_ust_id": inv.get("lieferant_ust_id"),
            "zahlungsbedingungen": inv.get("zahlungsbedingungen"),
            "lieferbedingungen": inv.get("lieferbedingungen"),
            "datei": inv.get("datei"),
            "pruefhinweis": "; ".join(d["hinweise"]),
        }
        for c, (key, _, _) in enumerate(RECHNUNG_SPALTEN, start=1):
            wert = werte.get(key)
            zelle = ws2.cell(row=r, column=c)
            if key in datum and wert:
                zelle.value = wert
                zelle.number_format = "DD.MM.YYYY"
            elif key in geld and wert is not None:
                zelle.value = float(wert)
                zelle.number_format = _geldformat(d["waehrung"])
                zelle.alignment = Alignment(horizontal="right")
            elif key == "ust_satz_prozent" and wert is not None:
                zelle.value = float(wert) / 100
                zelle.number_format = "0.0 %"
                zelle.alignment = Alignment(horizontal="right")
            elif key == "skonto_prozent" and wert is not None:
                zelle.value = float(wert) / 100
                zelle.number_format = "0.0 %"
                zelle.alignment = Alignment(horizontal="right")
            else:
                zelle.value = wert if wert not in ("", None) else None
            if key in lang:
                zelle.alignment = Alignment(wrap_text=True, vertical="top")
            if key == "pruefhinweis" and wert:
                zelle.font = Font(color=RED, size=9)
            if key == "stufe" and d["tage"] is not None and d["tage"] < 0:
                zelle.font = Font(color=RED, size=10, bold=True)
            zelle.border = border
        if r % 2 == 0:
            for c in range(1, n_spalten + 1):
                if ws2.cell(row=r, column=c).fill.fgColor.rgb in (None, "00000000"):
                    ws2.cell(row=r, column=c).fill = zebra
        ws2.row_dimensions[r].height = 28

    # Summenzeile
    letzte = len(daten) + 1
    summe_zeile = letzte + 1
    sp_netto = get_column_letter(_index("netto_gesamt"))
    sp_ust = get_column_letter(_index("ust_betrag"))
    sp_brutto = get_column_letter(_index("brutto_gesamt"))
    ws2.cell(row=summe_zeile, column=1, value="Summe (alle Währungen gemischt "
                                              "– siehe Übersicht)").font = Font(
        bold=True, size=10, color=GREY)
    for sp_ in (sp_netto, sp_ust, sp_brutto):
        zelle = ws2[f"{sp_}{summe_zeile}"]
        zelle.value = f"=SUBTOTAL(109,{sp_}2:{sp_}{letzte})"
        zelle.font = Font(bold=True, size=11, color=NAVY)
        zelle.number_format = "#,##0.00"
        zelle.alignment = Alignment(horizontal="right")
        zelle.border = Border(top=Side(style="double", color=NAVY))

    ws2.auto_filter.ref = f"A1:{get_column_letter(n_spalten)}{letzte}"
    ws2.freeze_panes = "C2"
    ws2.print_title_rows = "1:1"
    ws2.page_setup.orientation = "landscape"

    # ======================= Blatt 3: Positionen ============================
    ws3 = wb.create_sheet("Positionen")
    ws3.sheet_view.showGridLines = False
    for idx, (_, label, breite) in enumerate(POSITION_SPALTEN, start=1):
        zelle = ws3.cell(row=1, column=idx, value=label)
        zelle.fill = head_fill
        zelle.font = head_font
        zelle.alignment = Alignment(vertical="center", wrap_text=True)
        zelle.border = border
        ws3.column_dimensions[get_column_letter(idx)].width = breite
    ws3.row_dimensions[1].height = 26

    zeile3 = 2
    pos_zeile = {}          # Index der Rechnung -> erste Positionszeile
    for r_idx, d in enumerate(daten):
        inv = d["inv"]
        if d["positionen"]:
            pos_zeile[r_idx] = zeile3
        for pnr, pos in enumerate(d["positionen"], start=1):
            werte = {
                "rechnungsnummer": inv.get("rechnungsnummer"),
                "lieferant": inv.get("lieferant"),
                "rechnungsdatum": d["rd"],
                "pos_nr": pnr,
                "bezeichnung": pos.get("bezeichnung"),
                "menge": _zahl(pos.get("menge")),
                "einheit": pos.get("einheit"),
                "einzelpreis_netto": _zahl(pos.get("einzelpreis_netto")),
                "gesamt_netto": _zahl(pos.get("gesamt_netto")),
                "waehrung": d["waehrung"],
                "ust_satz_prozent": d["satz"],
            }
            for c, (key, _, _) in enumerate(POSITION_SPALTEN, start=1):
                wert = werte.get(key)
                zelle = ws3.cell(row=zeile3, column=c)
                if key == "rechnungsdatum" and wert:
                    zelle.value = wert
                    zelle.number_format = "DD.MM.YYYY"
                elif key in ("einzelpreis_netto", "gesamt_netto") and wert is not None:
                    zelle.value = float(wert)
                    zelle.number_format = _geldformat(d["waehrung"])
                    zelle.alignment = Alignment(horizontal="right")
                elif key == "menge" and wert is not None:
                    zelle.value = float(wert)
                    zelle.number_format = "#,##0.###"
                    zelle.alignment = Alignment(horizontal="right")
                elif key == "ust_satz_prozent" and wert is not None:
                    zelle.value = float(wert) / 100
                    zelle.number_format = "0.0 %"
                    zelle.alignment = Alignment(horizontal="right")
                else:
                    zelle.value = wert
                zelle.border = border
            # Streifen je Rechnung statt je Zeile: die Belege bleiben
            # optisch zusammen, ohne dass Leerzeilen den Filter zerstören.
            if r_idx % 2 == 1:
                for c in range(1, len(POSITION_SPALTEN) + 1):
                    ws3.cell(row=zeile3, column=c).fill = beleg_streifen
            # Rücksprung zur Rechnung
            link = ws3.cell(row=zeile3, column=1)
            link.hyperlink = f"#'Rechnungen'!A{r_idx + 2}"
            link.font = Font(color=BLUE, underline="single", size=10)
            zeile3 += 1

    if zeile3 > 2:
        ws3.auto_filter.ref = f"A1:{get_column_letter(len(POSITION_SPALTEN))}{zeile3 - 1}"
    ws3.freeze_panes = "E2"
    ws3.print_title_rows = "1:1"
    ws3.page_setup.orientation = "landscape"

    # Sprung von der Rechnung zu ihren Positionen. Erst hier möglich,
    # weil die Zielzeilen beim Aufbau von Blatt 2 noch nicht feststehen.
    sp_pos = _index("anzahl_positionen")
    for r_idx, d in enumerate(daten):
        ziel = pos_zeile.get(r_idx)
        if not ziel:
            continue
        zelle = ws2.cell(row=r_idx + 2, column=sp_pos)
        zelle.hyperlink = f"#'Positionen'!A{ziel}"
        zelle.font = Font(color=BLUE, underline="single", size=10)

    # ======================= Blatt 4: Prüfliste =============================
    ws4 = wb.create_sheet("Prüfliste")
    ws4.sheet_view.showGridLines = False
    kopf4 = ["Rechnungs-Nr.", "Lieferant", "Datum", "Brutto", "Währung",
             "Anzahl", "Art der Hinweise", "Was zu prüfen ist", "Datei", "Erledigt"]
    for idx, (label, breite) in enumerate(
            zip(kopf4, (16, 28, 13, 14, 9, 8, 26, 58, 26, 11)), start=1):
        zelle = ws4.cell(row=1, column=idx, value=label)
        zelle.fill = head_fill
        zelle.font = head_font
        zelle.alignment = Alignment(vertical="center", wrap_text=True)
        zelle.border = border
        ws4.column_dimensions[get_column_letter(idx)].width = breite
    ws4.row_dimensions[1].height = 30

    ws4.cell(row=2, column=1,
             value="Die Hinweise stammen aus festen Rechenregeln, nicht aus der "
                   "Texterkennung: Rechenproben, Datumslogik, Vollständigkeit. "
                   "Es wurde nichts korrigiert.").font = Font(
        size=9, italic=True, color=GREY)

    z4 = 3
    zeile_je_beleg = {d["inv"].get("rechnungsnummer"): i + 2
                      for i, d in enumerate(daten)}
    for d in sorted(mit_hinweis, key=lambda x: -len(x["hinweise"])):
        inv = d["inv"]
        arten = sorted({h.split(":")[0].split("–")[0].strip() for h in d["hinweise"]})
        nr = ws4.cell(row=z4, column=1, value=inv.get("rechnungsnummer"))
        ziel = zeile_je_beleg.get(inv.get("rechnungsnummer"))
        if ziel:
            nr.hyperlink = f"#'Rechnungen'!A{ziel}"
            nr.font = Font(color=BLUE, underline="single", size=10)
        ws4.cell(row=z4, column=2, value=inv.get("lieferant") or "–")
        zd = ws4.cell(row=z4, column=3, value=d["rd"])
        if d["rd"]:
            zd.number_format = "DD.MM.YYYY"
        zb = ws4.cell(row=z4, column=4, value=d["brutto"])
        zb.number_format = _geldformat(d["waehrung"])
        zb.alignment = Alignment(horizontal="right")
        ws4.cell(row=z4, column=5, value=d["waehrung"])
        za = ws4.cell(row=z4, column=6, value=len(d["hinweise"]))
        za.alignment = Alignment(horizontal="right")
        za.font = Font(bold=True, color=RED if len(d["hinweise"]) > 1 else NAVY)
        zk = ws4.cell(row=z4, column=7, value=", ".join(arten))
        zk.alignment = Alignment(wrap_text=True, vertical="top")
        zk.font = Font(size=9, color=NAVY)
        zh = ws4.cell(row=z4, column=8, value="\n".join(f"– {h}" for h in d["hinweise"]))
        zh.font = Font(color=RED, size=9)
        zh.alignment = Alignment(wrap_text=True, vertical="top")
        ws4.cell(row=z4, column=9, value=inv.get("datei"))
        ws4.cell(row=z4, column=10, value="")
        for c in range(1, 11):
            ws4.cell(row=z4, column=c).border = border
        ws4.row_dimensions[z4].height = max(28, 13 * len(d["hinweise"]) + 10)
        z4 += 1

    if z4 == 3:
        ws4.cell(row=3, column=1,
                 value="Keine Auffälligkeiten – Rechenproben und Daten sind plausibel."
                 ).font = Font(size=10, color=GREEN)
    else:
        ws4.auto_filter.ref = f"A2:J{z4 - 1}"
    ws4.freeze_panes = "A3"
    ws4.print_title_rows = "1:1"
    ws4.page_setup.orientation = "landscape"

    # ==================== Blatt 5: USt-Auswertung ===========================
    ws5 = wb.create_sheet("USt-Auswertung")
    ws5.sheet_view.showGridLines = False
    ws5.page_setup.orientation = "landscape"
    ws5.sheet_properties.pageSetUpPr.fitToPage = True
    ws5.page_setup.fitToWidth = 1
    ws5["B2"] = "Umsatzsteuer nach Satz"
    ws5["B2"].font = Font(size=14, bold=True, color=NAVY)
    ws5["B3"] = ("Grundlage für die Vorsteuer. Belege ohne erkannten Satz stehen "
                 "eigens ausgewiesen.")
    ws5["B3"].font = Font(size=9, color=GREY)
    for sp_, br in {"A": 2, "B": 16, "C": 10, "D": 18, "E": 18, "F": 18,
                    "G": 2}.items():
        ws5.column_dimensions[sp_].width = br

    z5 = 5
    for w in waehrungen:
        ws5.cell(row=z5, column=2, value=f"Währung {w}").font = Font(
            size=11, bold=True, color=NAVY)
        z5 += 1
        for c, txt in ((2, "USt-Satz"), (3, "Belege"), (4, "Netto"),
                       (5, "USt-Betrag"), (6, "Brutto")):
            zelle = ws5.cell(row=z5, column=c, value=txt)
            zelle.fill = head_fill
            zelle.font = head_font
            zelle.border = border
        z5 += 1
        teil = [d for d in daten if d["waehrung"] == w]
        saetze = sorted({d["satz"] for d in teil if d["satz"] is not None})
        kreis_start = z5
        for satz in saetze:
            gruppe = [d for d in teil if d["satz"] == satz]
            _ust_zeile(ws5, z5, f"{satz:g} %", gruppe, w, border, _geldformat)
            z5 += 1
        ohne = [d for d in teil if d["satz"] is None]
        if ohne:
            _ust_zeile(ws5, z5, "ohne Angabe", ohne, w, border, _geldformat, RED)
            z5 += 1
        kreis_ende = z5 - 1
        if w == leit and kreis_ende >= kreis_start:
            kreis = PieChart()
            kreis.title = f"Nettovolumen je USt-Satz ({w})"
            kreis.add_data(Reference(ws5, min_col=4, min_row=kreis_start - 1,
                                     max_row=kreis_ende), titles_from_data=True)
            kreis.set_categories(Reference(ws5, min_col=2, min_row=kreis_start,
                                           max_row=kreis_ende))
            _faerbe_punkte(kreis, [SERIE[0], SERIE[1], SERIE[2], SERIE[3], SERIE[4]])
            kreis.dataLabels = _nur_prozent()
            kreis.legend.position = "b"
            kreis.height, kreis.width = 8.5, 12
            ws5.add_chart(kreis, f"H{kreis_start - 2}")

        # Summenzeile
        _ust_zeile(ws5, z5, "Summe", teil, w, border, _geldformat, NAVY, fett=True)
        for c in range(2, 7):
            ws5.cell(row=z5, column=c).border = Border(
                top=Side(style="double", color=NAVY))
        z5 += 3

    ws5.cell(row=z5, column=2,
             value="Hinweis: 0 % kann steuerfreie innergemeinschaftliche Lieferung "
                   "oder Reverse Charge sein. Am Beleg prüfen.").font = Font(
        size=9, italic=True, color=GREY)

    for w in waehrungen:
        name = "Analyse" if len(waehrungen) == 1 else f"Analyse ({w})"
        _blatt_analyse(wb, daten, w, heute_tag, border, head_fill, head_font,
                       kpi_fill, blattname=name)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _index(key):
    for i, (k, _, _) in enumerate(RECHNUNG_SPALTEN, start=1):
        if k == key:
            return i
    return 1


def _ust_zeile(ws, zeile, label, gruppe, waehrung, border, geldformat,
               farbe=None, fett=False):
    from openpyxl.styles import Alignment, Font
    netto = sum(d["netto"] or 0 for d in gruppe)
    ust = sum(d["ust"] or 0 for d in gruppe)
    brutto = sum(d["brutto"] or 0 for d in gruppe)
    ws.cell(row=zeile, column=2, value=label).font = Font(
        size=10, bold=fett, color=farbe or "0F172A")
    ws.cell(row=zeile, column=3, value=len(gruppe)).alignment = Alignment(
        horizontal="right")
    for c, wert in ((4, netto), (5, ust), (6, brutto)):
        zelle = ws.cell(row=zeile, column=c, value=wert)
        zelle.number_format = geldformat(waehrung)
        zelle.alignment = Alignment(horizontal="right")
        zelle.font = Font(size=10, bold=fett, color=farbe or "0F172A")
    for c in range(2, 7):
        ws.cell(row=zeile, column=c).border = border


# ===================== Blatt 6: Analyse =================================
#
# Aufbau: zwei Abschnitte je Zeile, jeder Abschnitt besteht aus einem
# farbigen Titelband, einer Tabelle links und einem Diagramm rechts.
# Beide Abschnitte einer Zeile beginnen auf derselben Zeile, damit die
# Diagramme auf gleicher Höhe stehen.

SP_LINKS = 2        # Spalte B: Tabelle des linken Abschnitts
SP_RECHTS = 15      # Spalte O: Tabelle des rechten Abschnitts
DIA_LINKS = "G"     # Anker des linken Diagramms
DIA_RECHTS = "T"    # Anker des rechten Diagramms
ABSCHNITT_BREITE = 12   # Spalten je Abschnitt (Tabelle + Diagrammfläche)
ABSCHNITT_HOEHE = 19    # Zeilen je Abschnittszeile
DIA_H, DIA_B = 6.9, 11.2

BELEGKLASSEN = [
    ("bis 100", 0, 100),
    ("100 bis 500", 100, 500),
    ("500 bis 1.000", 500, 1000),
    ("1.000 bis 5.000", 1000, 5000),
    ("5.000 bis 10.000", 5000, 10000),
    ("über 10.000", 10000, None),
]


def _band(ws, zeile, spalte, titel, zusatz=None):
    """Farbiges Titelband über einem Abschnitt – trennt die Blöcke deutlich."""
    from openpyxl.styles import Alignment, Font, PatternFill
    fill = PatternFill("solid", fgColor=NAVY)
    for c in range(spalte, spalte + ABSCHNITT_BREITE):
        ws.cell(row=zeile, column=c).fill = fill
    zelle = ws.cell(row=zeile, column=spalte, value=titel)
    zelle.font = Font(size=12, bold=True, color="FFFFFF")
    zelle.alignment = Alignment(vertical="center")
    ws.row_dimensions[zeile].height = 22
    if zusatz:
        z2 = ws.cell(row=zeile + 1, column=spalte, value=zusatz)
        z2.font = Font(size=9, italic=True, color=GREY)
    return zeile + 2


def _kopfzeile(ws, zeile, spalte, beschriftungen):
    """Tabellenkopf in gedecktem Grau – eine Stufe unter dem Titelband."""
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    fill = PatternFill("solid", fgColor="E6EBF3")
    unten = Border(bottom=Side(style="medium", color=NAVY))
    for i, txt in enumerate(beschriftungen):
        zelle = ws.cell(row=zeile, column=spalte + i, value=txt)
        zelle.fill = fill
        zelle.font = Font(size=9, bold=True, color=NAVY)
        zelle.border = unten
        zelle.alignment = Alignment(horizontal="left" if i == 0 else "right")
    return zeile + 1


def _zeile_schreiben(ws, zeile, spalte, werte, formate, farben=None):
    """Eine Tabellenzeile mit Zahlenformaten und feiner Trennlinie."""
    from openpyxl.styles import Alignment, Border, Font, Side
    fein = Border(bottom=Side(style="thin", color=LINE))
    for i, wert in enumerate(werte):
        zelle = ws.cell(row=zeile, column=spalte + i, value=wert)
        zelle.border = fein
        zelle.alignment = Alignment(horizontal="left" if i == 0 else "right")
        farbe = (farben or {}).get(i, NAVY if i == 0 else "1F2937")
        zelle.font = Font(size=10, color=farbe)
        if formate[i]:
            zelle.number_format = formate[i]
    return zeile + 1


def _blatt_analyse(wb, daten, leit, heute_tag, border, head_fill, head_font,
                   kpi_fill, blattname="Analyse"):
    """Auswertungen, die über die reine Bestandsübersicht hinausgehen."""
    from openpyxl.chart import BarChart, DoughnutChart, LineChart, Reference
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter

    ws = wb.create_sheet(blattname)
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1

    geld = _geldformat(leit)
    teil = [d for d in daten if d["waehrung"] == leit]
    gesamt = sum(d["brutto"] or 0 for d in teil) or 1.0

    # Spaltenbreiten: zwei gleich gebaute Hälften
    ws.column_dimensions["A"].width = 2
    for basis in (SP_LINKS, SP_RECHTS):
        for versatz, breite in enumerate((26, 7, 14, 14, 2)):
            ws.column_dimensions[get_column_letter(basis + versatz)].width = breite
        for versatz in range(5, ABSCHNITT_BREITE):
            ws.column_dimensions[get_column_letter(basis + versatz)].width = 9
    ws.column_dimensions[get_column_letter(SP_RECHTS - 1)].width = 3

    ws["B2"] = blattname
    ws["B2"].font = Font(size=16, bold=True, color=NAVY)
    ws["B3"] = (f"Auswertung über {len(teil)} Beleg(e) in {leit}. "
                f"Jede Währung hat ein eigenes Blatt – es wird nichts umgerechnet.")
    ws["B3"].font = Font(size=10, color=GREY)

    basis_zeile = 5

    # ================== Zeile 1: Lieferanten | Monatsverlauf ==============
    nach_lief = {}
    for d in teil:
        name = (d["inv"].get("lieferant") or "Ohne Lieferant").strip()
        e = nach_lief.setdefault(name, {"betrag": 0.0, "anzahl": 0, "ziele": []})
        e["betrag"] += d["brutto"] or 0
        e["anzahl"] += 1
        if d["rd"] and d["fd"]:
            e["ziele"].append((d["fd"] - d["rd"]).days)
    rang = sorted(nach_lief.items(), key=lambda x: -x[1]["betrag"])

    z = _band(ws, basis_zeile, SP_LINKS, "Lieferantenkonzentration",
              "A = erste 80 % des Volumens, B = bis 95 %, C = Rest")
    kopf_abc = _kopfzeile(ws, z, SP_LINKS,
                          ["Lieferant", "Belege", "Betrag brutto", "kumuliert"])
    zeile = kopf_abc
    laufend = 0.0
    klassen = {"A": [0, 0.0], "B": [0, 0.0], "C": [0, 0.0]}
    for name, e in rang:
        vorher = laufend / gesamt          # Anteil VOR diesem Lieferanten
        laufend += e["betrag"]
        anteil = laufend / gesamt
        # Der Lieferant, der die Schwelle überschreitet, gehört noch zur
        # unteren Klasse – sonst fällt der größte Lieferant durchs Raster.
        klasse = "A" if vorher < 0.80 else ("B" if vorher < 0.95 else "C")
        klassen[klasse][0] += 1
        klassen[klasse][1] += e["betrag"]
        if zeile - kopf_abc < 9:
            zeile = _zeile_schreiben(
                ws, zeile, SP_LINKS,
                [f"{name[:30]}  ({klasse})", e["anzahl"], e["betrag"], anteil],
                [None, "0", geld, "0 %"])
    ende_abc = zeile - 1

    if ende_abc >= kopf_abc:
        balken = BarChart()
        balken.type = "col"
        balken.y_axis.numFmt = "#,##0"
        balken.gapWidth = 30
        balken.add_data(Reference(ws, min_col=SP_LINKS + 2, min_row=kopf_abc - 1,
                                  max_row=ende_abc), titles_from_data=True)
        balken.set_categories(Reference(ws, min_col=SP_LINKS, min_row=kopf_abc,
                                        max_row=ende_abc))
        linie = LineChart()
        linie.add_data(Reference(ws, min_col=SP_LINKS + 3, min_row=kopf_abc - 1,
                                 max_row=ende_abc), titles_from_data=True)
        linie.y_axis.axId = 200
        linie.y_axis.numFmt = "0 %"
        linie.y_axis.majorGridlines = None
        linie.y_axis.crosses = "max"
        balken += linie
        _schlicht(balken)
        try:
            balken.series[0].graphicalProperties.solidFill = BLUE
            balken.series[1].graphicalProperties.line.solidFill = GOLD
            balken.series[1].graphicalProperties.line.width = 22000
        except Exception:
            pass
        balken.legend.position = "b"
        balken.height, balken.width = DIA_H, DIA_B
        ws.add_chart(balken, f"{DIA_LINKS}{kopf_abc - 1}")

    monate = {}
    for d in teil:
        if d["rd"] is None:
            continue
        key = d["rd"].strftime("%Y-%m")
        e = monate.setdefault(key, {"betrag": 0.0, "anzahl": 0})
        e["betrag"] += d["brutto"] or 0
        e["anzahl"] += 1

    z = _band(ws, basis_zeile, SP_RECHTS, "Rechnungseingang je Monat",
              "nach Rechnungsdatum, nicht nach Eingangsdatum")
    kopf_mon = _kopfzeile(ws, z, SP_RECHTS, ["Monat", "Belege", "Betrag brutto"])
    zeile = kopf_mon
    for key in sorted(monate):
        jahr, mon = key.split("-")
        zeile = _zeile_schreiben(ws, zeile, SP_RECHTS,
                                 [f"{mon}.{jahr}", monate[key]["anzahl"],
                                  monate[key]["betrag"]],
                                 [None, "0", geld])
    ende_mon = zeile - 1

    if len(monate) >= 2:
        d_mon = LineChart()
        d_mon.y_axis.numFmt = "#,##0"
        d_mon.add_data(Reference(ws, min_col=SP_RECHTS + 2, min_row=kopf_mon - 1,
                                 max_row=ende_mon), titles_from_data=True)
        d_mon.set_categories(Reference(ws, min_col=SP_RECHTS, min_row=kopf_mon,
                                       max_row=ende_mon))
        _schlicht(d_mon)
        _faerbe_serie(d_mon, [SERIE[3]])
        d_mon.legend = None
        d_mon.height, d_mon.width = DIA_H, DIA_B
        ws.add_chart(d_mon, f"{DIA_RECHTS}{kopf_mon - 1}")

    basis_zeile += ABSCHNITT_HOEHE

    # ================== Zeile 2: Belegsgrößen | Skonto ====================
    z = _band(ws, basis_zeile, SP_LINKS, "Belegsgrößen",
              "Der Erfassungsaufwand ist je Beleg gleich, unabhängig vom Betrag")
    kopf_gr = _kopfzeile(ws, z, SP_LINKS,
                         [f"Betrag in {leit}", "Belege", "Summe brutto", "Anteil"])
    zeile = kopf_gr
    klein_anzahl = klein_betrag = 0
    for label, von, bis in BELEGKLASSEN:
        gruppe = [d for d in teil if d["brutto"] is not None
                  and abs(d["brutto"]) >= von
                  and (bis is None or abs(d["brutto"]) < bis)]
        betrag = sum(d["brutto"] or 0 for d in gruppe)
        if bis is not None and bis <= 500:
            klein_anzahl += len(gruppe)
            klein_betrag += betrag
        zeile = _zeile_schreiben(
            ws, zeile, SP_LINKS,
            [label, len(gruppe), betrag,
             (len(gruppe) / len(teil)) if teil else 0],
            [None, "0", geld, "0 %"])
    ende_gr = zeile - 1

    d_gr = BarChart()
    d_gr.type = "col"
    d_gr.legend = None
    d_gr.gapWidth = 40
    d_gr.y_axis.numFmt = "#,##0"
    d_gr.y_axis.title = "Belege"
    d_gr.add_data(Reference(ws, min_col=SP_LINKS + 1, min_row=kopf_gr - 1,
                            max_row=ende_gr), titles_from_data=True)
    d_gr.set_categories(Reference(ws, min_col=SP_LINKS, min_row=kopf_gr,
                                  max_row=ende_gr))
    _schlicht(d_gr)
    _faerbe_punkte(d_gr, ["9CC3F5", "76ACF0", "4E8FEA", "2C74E0", "1A56DB", "1443AE"])
    d_gr.height, d_gr.width = DIA_H, DIA_B
    ws.add_chart(d_gr, f"{DIA_LINKS}{kopf_gr - 1}")

    if teil:
        ws.cell(row=basis_zeile + ABSCHNITT_HOEHE - 3, column=SP_LINKS,
                value=f"{klein_anzahl} von {len(teil)} Belegen "
                      f"({klein_anzahl / len(teil) * 100:.0f} %) liegen unter "
                      f"500 {leit} und machen nur "
                      f"{klein_betrag / gesamt * 100:.0f} % des Volumens aus."
                ).font = Font(size=9, italic=True, color=NAVY)

    erreichbar = [d for d in teil if d["skonto_betrag"] and d["skonto_offen"]]
    verfallen = [d for d in teil if d["skonto_betrag"] and not d["skonto_offen"]]
    ohne = [d for d in teil if not d["skonto_betrag"]]

    z = _band(ws, basis_zeile, SP_RECHTS, "Skonto",
              "aus Satz und Frist auf dem Beleg berechnet")
    kopf_sk = _kopfzeile(ws, z, SP_RECHTS, ["Status", "Belege", "Skontobetrag"])
    zeile = kopf_sk
    for label, gruppe, farbe in (("Frist läuft noch", erreichbar, GREEN),
                                 ("Frist abgelaufen", verfallen, RED),
                                 ("Kein Skonto angeboten", ohne, GREY)):
        zeile = _zeile_schreiben(
            ws, zeile, SP_RECHTS,
            [label, len(gruppe), sum(d["skonto_betrag"] or 0 for d in gruppe)],
            [None, "0", geld], farben={0: farbe})
    ende_sk = zeile - 1

    d_sk = DoughnutChart(holeSize=48)
    d_sk.add_data(Reference(ws, min_col=SP_RECHTS + 1, min_row=kopf_sk - 1,
                            max_row=ende_sk), titles_from_data=True)
    d_sk.set_categories(Reference(ws, min_col=SP_RECHTS, min_row=kopf_sk,
                                  max_row=ende_sk))
    _faerbe_punkte(d_sk, [AMPEL["gruen"], AMPEL["rot"], AMPEL["grau"]])
    d_sk.dataLabels = _nur_prozent()
    d_sk.legend.position = "b"
    d_sk.height, d_sk.width = DIA_H, DIA_B
    ws.add_chart(d_sk, f"{DIA_RECHTS}{kopf_sk - 1}")

    v_summe = sum(d["skonto_betrag"] or 0 for d in verfallen)
    hinweis = (f"Bei {len(verfallen)} Beleg(en) ist die Frist abgelaufen, "
               f"rechnerisch entgangen: {_fmt_eur(v_summe, leit)}. "
               f"Ob gezahlt wurde, sagt diese Auswertung nicht."
               if v_summe else "Keine abgelaufene Skontofrist gefunden.")
    ws.cell(row=basis_zeile + ABSCHNITT_HOEHE - 3, column=SP_RECHTS,
            value=hinweis).font = Font(
        size=9, italic=True, color=RED if v_summe else GREEN)

    basis_zeile += ABSCHNITT_HOEHE

    # ================== Zeile 3: Zahlungsziele | Artikel ==================
    mit_ziel = [(n, e) for n, e in rang if e["ziele"]][:9]
    if mit_ziel:
        z = _band(ws, basis_zeile, SP_LINKS, "Zahlungsziele je Lieferant",
                  "Tage zwischen Rechnungs- und Fälligkeitsdatum")
        kopf_zz = _kopfzeile(ws, z, SP_LINKS,
                             ["Lieferant", "Belege", "Ø Tage", "Volumen"])
        zeile = kopf_zz
        for name, e in mit_ziel:
            zeile = _zeile_schreiben(
                ws, zeile, SP_LINKS,
                [name[:30], len(e["ziele"]),
                 round(sum(e["ziele"]) / len(e["ziele"]), 1), e["betrag"]],
                [None, "0", "0.0", geld])
        ende_zz = zeile - 1

        d_zz = BarChart()
        d_zz.type = "bar"
        d_zz.legend = None
        d_zz.gapWidth = 40
        d_zz.x_axis.numFmt = "#,##0"
        d_zz.x_axis.title = "Tage"
        d_zz.add_data(Reference(ws, min_col=SP_LINKS + 2, min_row=kopf_zz - 1,
                                max_row=ende_zz), titles_from_data=True)
        d_zz.set_categories(Reference(ws, min_col=SP_LINKS, min_row=kopf_zz,
                                      max_row=ende_zz))
        _schlicht(d_zz)
        _faerbe_serie(d_zz, [SERIE[5]])
        d_zz.height, d_zz.width = DIA_H, DIA_B
        ws.add_chart(d_zz, f"{DIA_LINKS}{kopf_zz - 1}")

    artikel = {}
    for d in teil:
        for pos in d["positionen"]:
            bez = str(pos.get("bezeichnung") or "").strip()
            if not bez:
                continue
            e = artikel.setdefault(bez[:40], {"betrag": 0.0, "anzahl": 0,
                                              "lieferanten": set()})
            e["betrag"] += _zahl(pos.get("gesamt_netto")) or 0
            e["anzahl"] += 1
            if d["inv"].get("lieferant"):
                e["lieferanten"].add(d["inv"]["lieferant"])
    top_artikel = sorted(artikel.items(), key=lambda x: -x[1]["betrag"])[:9]

    if top_artikel:
        z = _band(ws, basis_zeile, SP_RECHTS, "Größte Positionen nach Artikel",
                  "aus dem Blatt „Positionen“ zusammengefasst")
        kopf_ar = _kopfzeile(ws, z, SP_RECHTS,
                             ["Bezeichnung", "Posten", "Summe netto", "Lieferanten"])
        zeile = kopf_ar
        for bez, e in top_artikel:
            mehrere = len(e["lieferanten"]) > 1
            zeile = _zeile_schreiben(
                ws, zeile, SP_RECHTS,
                [bez, e["anzahl"], e["betrag"], len(e["lieferanten"])],
                [None, "0", geld, "0"],
                farben={3: GOLD} if mehrere else None)
        ende_ar = zeile - 1

        d_ar = BarChart()
        d_ar.type = "bar"
        d_ar.legend = None
        d_ar.gapWidth = 40
        d_ar.x_axis.numFmt = "#,##0"
        d_ar.x_axis.title = leit
        d_ar.add_data(Reference(ws, min_col=SP_RECHTS + 2, min_row=kopf_ar - 1,
                                max_row=ende_ar), titles_from_data=True)
        d_ar.set_categories(Reference(ws, min_col=SP_RECHTS, min_row=kopf_ar,
                                      max_row=ende_ar))
        _schlicht(d_ar)
        _faerbe_serie(d_ar, [SERIE[1]])
        d_ar.height, d_ar.width = DIA_H, DIA_B
        ws.add_chart(d_ar, f"{DIA_RECHTS}{kopf_ar - 1}")

        anzahl_mehrere = sum(1 for _, e in top_artikel if len(e["lieferanten"]) > 1)
        if anzahl_mehrere:
            ws.cell(row=basis_zeile + ABSCHNITT_HOEHE - 3, column=SP_RECHTS,
                    value=f"{anzahl_mehrere} Artikel kommen von mehr als einem "
                          f"Lieferanten – dort lohnt ein Preisvergleich."
                    ).font = Font(size=9, italic=True, color=GOLD)

    # Klassen-Zusammenfassung unten anhängen
    basis_zeile += ABSCHNITT_HOEHE
    ws.cell(row=basis_zeile, column=SP_LINKS,
            value="Einordnung der Lieferanten").font = Font(
        size=11, bold=True, color=NAVY)
    for i, schl in enumerate(("A", "B", "C"), start=1):
        anzahl, betrag = klassen[schl]
        ws.cell(row=basis_zeile + i, column=SP_LINKS,
                value=f"Klasse {schl}: {anzahl} Lieferant(en), "
                      f"{_fmt_eur(betrag, leit)} "
                      f"({betrag / gesamt * 100:.0f} % des Volumens)").font = Font(
            size=10, color=NAVY if schl == "A" else GREY)
