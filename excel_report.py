"""Excel-Report für die LN Rechnungserfassung – 3 Blätter in Profi-Optik."""
from __future__ import annotations

import io
from datetime import datetime, timedelta

NAVY = "0F172A"
BLUE = "1A56DB"
GOLD = "B48A2F"
LIGHT = "F3F6FB"
GREY = "64748B"
RED = "B91C1C"

RECHNUNG_SPALTEN = [
    ("datei", "Datei", 30),
    ("lieferant", "Lieferant", 26),
    ("rechnungsnummer", "Rechnungs-Nr.", 16),
    ("rechnungsdatum", "Rechnungsdatum", 16),
    ("faelligkeitsdatum", "Fällig am", 14),
    ("netto_gesamt", "Netto", 12),
    ("ust_satz_prozent", "USt %", 8),
    ("ust_betrag", "USt-Betrag", 12),
    ("brutto_gesamt", "Brutto", 12),
    ("waehrung", "Währung", 9),
    ("skonto", "Skonto", 22),
    ("zahlungsbedingungen", "Zahlungsbedingungen", 34),
    ("lieferbedingungen", "Lieferbedingungen", 34),
    ("iban", "IBAN", 26),
    ("lieferant_ust_id", "USt-ID Lieferant", 18),
    ("pruefhinweis", "Prüfhinweis", 44),
]


# Formulierungen, bei denen die Rechnung am Rechnungsdatum fällig ist
SOFORT_MARKER = (
    "sofort fällig", "sofort faellig", "sofort zahlbar", "zahlbar sofort",
    "sofort ohne abzug", "zahlung sofort", "sofort rein netto", "sofort netto",
    "zahlbar bei erhalt", "bei erhalt fällig", "due immediately", "payable immediately",
    "due upon receipt", "payable on receipt",
)


def _zahl(v):
    """Robust nach float – akzeptiert auch '1.234,56 €'. Sonst None."""
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    t = str(v).replace("€", "").replace("EUR", "").replace(" ", "").strip()
    if "," in t and "." in t:
        t = t.replace(".", "").replace(",", ".")
    elif "," in t:
        t = t.replace(",", ".")
    try:
        return float(t)
    except ValueError:
        return None


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

    # Rechenprobe Netto + USt = Brutto
    if netto is not None and ust is not None and brutto is not None:
        if abs(netto + ust - brutto) > 0.02:
            h.append(f"Rechenprobe: Netto + USt ergibt {netto + ust:.2f}, "
                     f"Beleg nennt {brutto:.2f}")

    # USt-Betrag zum angegebenen Satz
    if netto is not None and ust is not None and satz:
        erwartet = netto * satz / 100
        if abs(erwartet - ust) > max(0.02, abs(netto) * 0.002):
            h.append(f"USt-Betrag passt nicht zu {satz:g} % (erwartet {erwartet:.2f})")

    # Summe der Einzelpositionen
    posten = [_zahl(p.get("gesamt_netto")) for p in (inv.get("positionen") or [])]
    if posten and all(x is not None for x in posten) and netto is not None:
        if abs(sum(posten) - netto) > 0.02:
            h.append(f"Positionen ergeben {sum(posten):.2f} netto, Beleg nennt {netto:.2f}")

    # Datumsplausibilität
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


def parse_date(s):
    if not s:
        return None
    for fmt in ("%d.%m.%Y", "%d.%m.%y", "%Y-%m-%d"):
        try:
            return datetime.strptime(str(s).strip(), fmt)
        except ValueError:
            continue
    return None


def _fmt_eur(v, waehrung="EUR"):
    if v is None:
        return "–"
    return f"{v:,.2f} {waehrung or ''}".replace(",", "X").replace(".", ",").replace("X", ".")


def build_excel(invoices: list[dict]) -> bytes:
    from openpyxl import Workbook
    from openpyxl.chart import BarChart, Reference
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    heute = datetime.now()
    for _i in invoices:
        normalisiere(_i)

    thin = Side(style="thin", color="D8DEE9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    head_fill = PatternFill("solid", fgColor=NAVY)
    head_font = Font(color="FFFFFF", bold=True, size=11)
    kpi_fill = PatternFill("solid", fgColor=LIGHT)

    def style_header_row(ws, row, ncols):
        for c in range(1, ncols + 1):
            cell = ws.cell(row=row, column=c)
            cell.fill = head_fill
            cell.font = head_font
            cell.alignment = Alignment(vertical="center", horizontal="left")
            cell.border = border
        ws.row_dimensions[row].height = 22

    # ============================ Blatt 1: Übersicht =========================
    ws = wb.active
    ws.title = "Übersicht"
    ws.sheet_view.showGridLines = False
    # damit beim Drucken/PDF alle vier Kennzahlen-Karten auf die Seite passen
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1

    ws["B2"] = "LN Automation – Rechnungsübersicht"
    ws["B2"].font = Font(size=18, bold=True, color=NAVY)
    ws["B3"] = f"Erstellt am {heute.strftime('%d.%m.%Y um %H:%M')} Uhr · {len(invoices)} Eingangsrechnung(en)"
    ws["B3"].font = Font(size=10, color=GREY)

    # Kennzahlen – je Währung ein eigener Block. Beträge verschiedener
    # Währungen werden NICHT addiert und nicht umgerechnet.
    waehrungen = sorted({(i.get("waehrung") or "EUR") for i in invoices}) or ["EUR"]

    zeile_kpi = 5
    for w in waehrungen:
        teil = [i for i in invoices if (i.get("waehrung") or "EUR") == w]
        summe = sum(i.get("brutto_gesamt") or 0 for i in teil)
        faellig14, ueberfaellig = 0.0, 0.0
        for i in teil:
            d = parse_date(i.get("faelligkeitsdatum"))
            if d is None:
                continue
            tage = (d - heute).days
            if tage < 0:
                ueberfaellig += i.get("brutto_gesamt") or 0
            elif tage <= 14:
                faellig14 += i.get("brutto_gesamt") or 0

        if len(waehrungen) > 1:
            ws.cell(row=zeile_kpi, column=2,
                    value=f"Währung {w}").font = Font(size=11, bold=True, color=NAVY)
            zeile_kpi += 1

        kpis = [
            ("Offene Posten gesamt (brutto)", _fmt_eur(summe, w)),
            ("Davon fällig in ≤ 14 Tagen", _fmt_eur(faellig14, w)),
            ("Davon überfällig", _fmt_eur(ueberfaellig, w)),
            ("Anzahl Rechnungen", str(len(teil))),
        ]
        col = 2
        for label, wert in kpis:
            for r in (zeile_kpi, zeile_kpi + 1):
                for c in range(col, col + 3):
                    ws.cell(row=r, column=c).fill = kpi_fill
            ws.cell(row=zeile_kpi, column=col, value=label).font = Font(size=9, color=GREY)
            wcell = ws.cell(row=zeile_kpi + 1, column=col, value=wert)
            wcell.font = Font(size=14, bold=True,
                              color=NAVY if "überfällig" not in label else RED)
            col += 4
        zeile_kpi += 3

    if len(waehrungen) > 1:
        ws.cell(row=zeile_kpi, column=2,
                value="Mehrere Währungen erkannt – die Summen stehen getrennt und werden "
                      "nicht umgerechnet.").font = Font(size=9, italic=True, color=GREY)
        zeile_kpi += 2

    # Diagramm: Daten liegen auf verstecktem Blatt (rendert zuverlässig)
    ws_data = wb.create_sheet("Daten")
    ws_data.sheet_state = "hidden"
    ws_data["A1"], ws_data["B1"] = "Rechnung", "Brutto"
    for idx, inv in enumerate(invoices, start=2):
        _w = inv.get("waehrung") or "EUR"
        _name = f"{inv.get('rechnungsnummer') or inv.get('datei')}"
        if len(waehrungen) > 1:
            _name = f"{_name} ({_w})"
        ws_data.cell(row=idx, column=1, value=_name)
        ws_data.cell(row=idx, column=2, value=inv.get("brutto_gesamt") or 0)
    chart = BarChart()
    chart.type = "col"
    chart.title = ("Bruttobetrag je Rechnung (nicht umgerechnet)"
                   if len(waehrungen) > 1 else "Bruttobetrag je Rechnung")
    chart.legend = None
    chart.y_axis.title = waehrungen[0] if len(waehrungen) == 1 else "Betrag"
    data = Reference(ws_data, min_col=2, min_row=1, max_row=1 + len(invoices))
    cats = Reference(ws_data, min_col=1, min_row=2, max_row=1 + len(invoices))
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.height, chart.width = 8, 18
    ws.add_chart(chart, f"B{zeile_kpi}")

    # Hinweise
    zeile = zeile_kpi + 18
    ws.cell(row=zeile, column=2, value="⚠ Bald fällig / überfällig").font = Font(bold=True, size=12, color=NAVY)
    zeile += 1
    faellige = sorted(
        [i for i in invoices if parse_date(i.get("faelligkeitsdatum"))],
        key=lambda i: parse_date(i.get("faelligkeitsdatum")),
    )
    notiert = 0
    for i in faellige:
        d = parse_date(i.get("faelligkeitsdatum"))
        tage = (d - heute).days
        if tage > 14 or notiert >= 6:
            continue
        status = f"überfällig seit {abs(tage)} Tagen" if tage < 0 else f"fällig in {tage} Tagen"
        c = ws.cell(
            row=zeile, column=2,
            value=f"• {i.get('rechnungsnummer') or i.get('datei')} · {i.get('lieferant') or '?'} · "
                  f"{_fmt_eur(i.get('brutto_gesamt'), i.get('waehrung'))} · {status} ({i.get('faelligkeitsdatum')})",
        )
        c.font = Font(color=RED if tage < 0 else NAVY, size=10)
        zeile += 1
        notiert += 1
    if notiert == 0:
        ws.cell(row=zeile, column=2, value="• Keine Rechnung in den nächsten 14 Tagen fällig.").font = Font(size=10, color=GREY)
        zeile += 1

    zeile += 1
    ws.cell(row=zeile, column=2, value="💡 Auffällig hohe Beträge").font = Font(bold=True, size=12, color=NAVY)
    zeile += 1
    # Durchschnitt je Währung – sonst vergleicht man USD mit EUR
    schnitte = {}
    for w in waehrungen:
        teil = [i.get("brutto_gesamt") or 0 for i in invoices
                if (i.get("waehrung") or "EUR") == w]
        schnitte[w] = (sum(teil) / len(teil)) if teil else 0

    hohe = []
    for i in invoices:
        w = i.get("waehrung") or "EUR"
        anzahl_w = sum(1 for x in invoices if (x.get("waehrung") or "EUR") == w)
        if anzahl_w > 1 and schnitte[w] and (i.get("brutto_gesamt") or 0) > 1.5 * schnitte[w]:
            hohe.append(i)

    for i in sorted(hohe, key=lambda x: -((x.get("brutto_gesamt") or 0)
                                          / (schnitte[x.get("waehrung") or "EUR"] or 1)))[:5]:
        w = i.get("waehrung") or "EUR"
        faktor = (i.get("brutto_gesamt") or 0) / (schnitte[w] or 1)
        ws.cell(
            row=zeile, column=2,
            value=f"• {i.get('rechnungsnummer') or i.get('datei')} · {i.get('lieferant') or '?'} · "
                  f"{_fmt_eur(i.get('brutto_gesamt'), w)} "
                  f"({faktor:.1f}× Durchschnitt in {w})",
        ).font = Font(size=10, color=NAVY)
        zeile += 1
    if not hohe:
        ws.cell(row=zeile, column=2, value="• Keine Ausreißer – Beträge liegen nah beieinander.").font = Font(size=10, color=GREY)
        zeile += 1

    # Prüfhinweise
    zeile += 2
    ws.cell(row=zeile, column=2, value="🔍 Bitte am Beleg prüfen").font = Font(bold=True, size=12, color=NAVY)
    zeile += 1
    auffaellig = [(i, pruefhinweise(i)) for i in invoices]
    auffaellig = [(i, hs) for i, hs in auffaellig if hs]
    if not auffaellig:
        ws.cell(row=zeile, column=2,
                value="• Keine Auffälligkeiten – Rechenproben und Daten sind plausibel."
                ).font = Font(size=10, color=GREY)
    else:
        for i, hs in auffaellig[:10]:
            kopf = (f"• {i.get('rechnungsnummer') or i.get('datei')} · "
                    f"{i.get('lieferant') or '?'}")
            ws.cell(row=zeile, column=2, value=kopf).font = Font(size=10, bold=True, color=NAVY)
            zeile += 1
            for hinweis in hs:
                ws.cell(row=zeile, column=2, value=f"    – {hinweis}").font = Font(size=10, color=RED)
                zeile += 1
        if len(auffaellig) > 10:
            ws.cell(row=zeile, column=2,
                    value=f"• … und {len(auffaellig) - 10} weitere – siehe Spalte "
                          f"„Prüfhinweis“ auf dem Blatt „Rechnungen“."
                    ).font = Font(size=10, color=GREY)

    for c, w in zip("ABCDEFGHIJKLMNOP", (2, 18, 12, 6, 18, 12, 6, 18, 12, 6, 14, 12, 6, 6, 6, 6)):
        ws.column_dimensions[c].width = w

    # ============================ Blatt 2: Rechnungen ========================
    # Anker: In welcher Zeile beginnt jede Rechnung auf dem Positionen-Blatt?
    anchors, _z = [], 1
    for inv in invoices:
        anchors.append(_z)
        _z += 2 + max(len(inv.get("positionen") or []), 1) + 1

    ws2 = wb.create_sheet("Rechnungen")
    ws2.sheet_view.showGridLines = False
    for idx, (_, label, breite) in enumerate(RECHNUNG_SPALTEN, start=1):
        ws2.cell(row=1, column=idx, value=label)
        ws2.column_dimensions[get_column_letter(idx)].width = breite
    link_col = len(RECHNUNG_SPALTEN) + 1
    ws2.cell(row=1, column=link_col, value="Positionen")
    ws2.column_dimensions[get_column_letter(link_col)].width = 14
    style_header_row(ws2, 1, link_col)

    geld_spalten = {"netto_gesamt", "ust_betrag", "brutto_gesamt"}
    datum_spalten = {"rechnungsdatum", "faelligkeitsdatum"}
    lang_spalten = {"zahlungsbedingungen", "lieferbedingungen", "skonto", "datei",
                    "pruefhinweis"}

    for r, inv in enumerate(invoices, start=2):
        for c, (key, _, _) in enumerate(RECHNUNG_SPALTEN, start=1):
            wert = "; ".join(pruefhinweise(inv)) if key == "pruefhinweis" else inv.get(key)
            cell = ws2.cell(row=r, column=c)
            if key in datum_spalten and parse_date(wert):
                cell.value = parse_date(wert)
                cell.number_format = "DD.MM.YYYY"
            elif key in geld_spalten and wert is not None:
                cell.value = float(wert)
                cell.number_format = "#,##0.00"
                cell.alignment = Alignment(horizontal="right")
            else:
                cell.value = wert
            if key in lang_spalten:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
            if key == "pruefhinweis" and wert:
                cell.font = Font(color=RED, size=10)
            cell.border = border
        lcell = ws2.cell(row=r, column=link_col, value="→ ansehen")
        lcell.hyperlink = f"#Positionen!A{anchors[r - 2]}"
        lcell.font = Font(color=BLUE, underline="single")
        lcell.border = border
        lcell.alignment = Alignment(horizontal="center", vertical="center")
        if r % 2 == 0:
            for c in range(1, link_col + 1):
                ws2.cell(row=r, column=c).fill = PatternFill("solid", fgColor="F8FAFC")
        ws2.row_dimensions[r].height = 32
    ws2.auto_filter.ref = f"A1:{get_column_letter(link_col)}{len(invoices) + 1}"
    ws2.freeze_panes = "A2"

    # =========================== Blatt 3: Positionen =========================
    ws3 = wb.create_sheet("Positionen")
    ws3.sheet_view.showGridLines = False
    pos_spalten = [("bezeichnung", "Bezeichnung", 48), ("menge", "Menge", 10),
                   ("einheit", "Einheit", 10), ("einzelpreis_netto", "Einzelpreis (netto)", 18),
                   ("gesamt_netto", "Gesamt (netto)", 16)]
    zeile = 1
    for inv_idx, inv in enumerate(invoices):
        titel = (f"{inv.get('rechnungsnummer') or inv.get('datei')} · "
                 f"{inv.get('lieferant') or 'Unbekannter Lieferant'} · "
                 f"Brutto {_fmt_eur(inv.get('brutto_gesamt'), inv.get('waehrung'))}")
        ws3.merge_cells(start_row=zeile, start_column=1, end_row=zeile, end_column=len(pos_spalten))
        tcell = ws3.cell(row=zeile, column=1, value=titel)
        tcell.font = Font(bold=True, color="FFFFFF", size=11)
        tcell.fill = PatternFill("solid", fgColor=BLUE)
        tcell.alignment = Alignment(vertical="center")
        tcell.hyperlink = f"#Rechnungen!A{inv_idx + 2}"
        ws3.row_dimensions[zeile].height = 20
        zeile += 1
        for c, (_, label, breite) in enumerate(pos_spalten, start=1):
            hcell = ws3.cell(row=zeile, column=c, value=label)
            hcell.font = Font(bold=True, size=9, color=GREY)
            hcell.border = border
            ws3.column_dimensions[get_column_letter(c)].width = breite
        zeile += 1
        for pos in inv.get("positionen") or [{}]:
            for c, (key, _, _) in enumerate(pos_spalten, start=1):
                cell = ws3.cell(row=zeile, column=c, value=pos.get(key))
                cell.border = border
                if key in ("einzelpreis_netto", "gesamt_netto") and pos.get(key) is not None:
                    cell.number_format = "#,##0.00"
                    cell.alignment = Alignment(horizontal="right")
            zeile += 1
        zeile += 1  # Absatz zwischen den Rechnungen

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
