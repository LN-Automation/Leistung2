"""LN Automation – KI-Rechnungserfassung (Produkt 2).

Start lokal:   streamlit run ln_billing_app.py
Kunde lädt Rechnungen (PDF/Scan) hoch -> Claude liest sie strukturiert aus
-> Tabelle in der App -> Excel-Download (Übersicht + Einzelpositionen).
"""
from __future__ import annotations

import base64
import io
import json
import os
from pathlib import Path

import streamlit as st
from dotenv import load_dotenv

load_dotenv()


def setting(name: str, default: str = "") -> str:
    try:
        if name in st.secrets:
            return str(st.secrets[name])
    except Exception:  # noqa: BLE001
        pass
    return os.getenv(name, default)


# Favicon: das Logo, falls es im Ordner liegt – sonst als Rückfall das Emoji
_LOGO_PFAD = Path(__file__).parent / "logo.png"
st.set_page_config(
    page_title="LN Automation – Rechnungserfassung",
    page_icon="logo.png" if _LOGO_PFAD.exists() else "🧾",
    layout="wide",
)

st.markdown(
    """
    <style>
      /* Menü, Deploy-Button und die Cloud-Leiste ("Fork", GitHub) ausblenden.
         Der Header bleibt bestehen – dort sitzt der Aufklapp-Pfeil der Sidebar,
         der darunter gezielt wieder sichtbar gemacht wird. */
      #MainMenu, footer, .stAppDeployButton, [data-testid="stStatusWidget"],
      [data-testid="stToolbar"], [data-testid="stToolbarActions"],
      [data-testid="stAppToolbar"], .stAppToolbar {visibility: hidden;}
      [data-testid="stHeader"] {background: transparent;}
      [data-testid="stSidebarCollapsedControl"],
      [data-testid="stSidebarCollapsedControl"] *,
      [data-testid="collapsedControl"],
      [data-testid="collapsedControl"] * {
        display: flex !important; visibility: visible !important;
        opacity: 1 !important; pointer-events: auto !important; z-index: 99999 !important;
      }
      .block-container {padding-top: 2.2rem;}
      .ln-section {
        font-size: 0.82rem; letter-spacing: 0.12em; text-transform: uppercase;
        color: #64748b; font-weight: 600; margin: 2.2rem 0 0.6rem 0;
      }
    </style>
    """,
    unsafe_allow_html=True,
)

def _logo_block(breite: int = 190, untertitel: str = "") -> str:
    """HTML-Kopf mit zentriertem Logo. Ohne Logo-Datei: Schriftzug."""
    if _LOGO_PFAD.exists():
        b64 = base64.b64encode(_LOGO_PFAD.read_bytes()).decode()
        inneres = f'<img src="data:image/png;base64,{b64}" style="width:{breite}px;max-width:60%;" />'
    else:
        inneres = '<div style="color:#0f172a;font-size:2rem;font-weight:800;">LN Automation</div>'
    unter = (f'<div style="color:#64748b;font-size:1.0rem;margin-top:2px;">{untertitel}</div>'
             if untertitel else "")
    return (f'<div style="text-align:center;padding:6px 0 2px 0;">{inneres}{unter}</div>'
            f'<hr style="border:none;border-top:1px solid #e2e8f0;margin:8px 0 1.2rem 0;" />')


# ------------------------------ Anmeldung -----------------------------------

_kunden: dict[str, str] = {}
try:
    if "kunden" in st.secrets:
        _kunden = {str(k).lower(): str(v) for k, v in dict(st.secrets["kunden"]).items()}
except Exception:  # noqa: BLE001
    pass
_app_pw = setting("APP_PASSWORD")

if (_kunden or _app_pw) and not st.session_state.get("auth_ok"):
    # Die App läuft auf layout="wide" – für die Anmeldeseite die Breite
    # begrenzen, damit sie aussieht wie bei Produkt 1 (layout="centered").
    st.markdown(
        """
        <style>
          .block-container {max-width: 46rem !important; padding-top: 3rem !important;}
        </style>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(
        _logo_block(210, "KI-Rechnungserfassung – Eingangsrechnungen automatisch auslesen"),
        unsafe_allow_html=True,
    )
    with st.form("login"):
        st.markdown("**Anmeldung**")
        firma = st.text_input("Firmen-Kennung") if _kunden else ""
        pw = st.text_input("Zugangspasswort", type="password")
        if st.form_submit_button("Anmelden", type="primary", width="stretch"):
            if _kunden:
                f = firma.strip().lower()
                if f in _kunden and pw == _kunden[f]:
                    st.session_state.auth_ok = True
                    st.session_state.kunde = firma.strip()
                    st.rerun()
                else:
                    st.error("Firmen-Kennung oder Passwort falsch.")
            elif pw == _app_pw:
                st.session_state.auth_ok = True
                st.rerun()
            else:
                st.error("Falsches Passwort.")
    st.stop()

# ------------------------------- Kopfbereich --------------------------------

st.markdown(
    _logo_block(190, "KI-Rechnungserfassung – Eingangsrechnungen automatisch auslesen"),
    unsafe_allow_html=True,
)

with st.sidebar:
    st.header("Einrichtung")
    if st.session_state.get("kunde"):
        st.caption(f"Angemeldet: {st.session_state.kunde}")
        if st.button("Abmelden"):
            for _k in ("auth_ok", "kunde", "invoices", "positions"):
                st.session_state.pop(_k, None)
            st.rerun()
        st.divider()
    if "api_key" not in st.session_state:
        st.session_state.api_key = setting("ANTHROPIC_API_KEY")
    if st.session_state.api_key:
        st.success("API-Key aktiv ✓")
    else:
        with st.form("key_form"):
            _k = st.text_input("Claude API-Key", type="password")
            if st.form_submit_button("Speichern", type="primary"):
                if _k.strip().startswith("sk-ant-"):
                    st.session_state.api_key = _k.strip()
                    st.rerun()
                else:
                    st.error("Key beginnt mit sk-ant-")
    api_key = st.session_state.api_key

# --------------------------- Claude-Extraktion ------------------------------

EXTRAKTIONS_PROMPT = """Du bist ein Buchhaltungs-Assistent. Lies die beigefügte Eingangsrechnung
und extrahiere die Daten EXAKT wie im Dokument angegeben (keine Werte erfinden,
fehlende Felder als null). Antworte AUSSCHLIESSLICH mit einem JSON-Objekt in
genau diesem Schema, ohne Markdown, ohne Erklärungen:

{
  "lieferant": string|null,
  "lieferant_ust_id": string|null,
  "rechnungsnummer": string|null,
  "rechnungsdatum": string|null,
  "leistungsdatum": string|null,
  "faelligkeitsdatum": string|null,
  "zahlungsbedingungen": string|null,
  "skonto": string|null,
  "lieferbedingungen": string|null,
  "waehrung": string|null,
  "netto_gesamt": number|null,
  "ust_satz_prozent": number|null,
  "ust_betrag": number|null,
  "brutto_gesamt": number|null,
  "iban": string|null,
  "positionen": [
    {"bezeichnung": string, "menge": number|null, "einheit": string|null,
     "einzelpreis_netto": number|null, "gesamt_netto": number|null}
  ]
}

Beträge als Zahlen mit Punkt als Dezimaltrenner (1234.56). Daten als TT.MM.JJJJ.

Zum Fälligkeitsdatum: Steht kein Datum auf der Rechnung, lässt sich aber aus
Rechnungsdatum und Zahlungsbedingungen eindeutig berechnen, dann berechne es
("30 Tage netto" -> Rechnungsdatum + 30 Tage; "sofort fällig" oder "zahlbar bei
Erhalt" -> Rechnungsdatum). Ist es nicht eindeutig, gib null zurück. Niemals raten."""


def extract_invoice(filename: str, data: bytes, key: str) -> dict:
    from anthropic import Anthropic

    suffix = Path(filename).suffix.lower()
    if suffix == ".pdf":
        block = {
            "type": "document",
            "source": {
                "type": "base64",
                "media_type": "application/pdf",
                "data": base64.b64encode(data).decode(),
            },
        }
    elif suffix in (".png", ".jpg", ".jpeg", ".webp"):
        media = "image/png" if suffix == ".png" else "image/jpeg"
        block = {
            "type": "image",
            "source": {
                "type": "base64",
                "media_type": media,
                "data": base64.b64encode(data).decode(),
            },
        }
    else:
        raise RuntimeError(f"Format {suffix} wird nicht unterstützt (PDF, PNG, JPG).")

    msg = Anthropic(api_key=key).messages.create(
        model="claude-sonnet-4-6",
        max_tokens=3000,
        messages=[{"role": "user", "content": [block, {"type": "text", "text": EXTRAKTIONS_PROMPT}]}],
    )
    text = "".join(b.text for b in msg.content if b.type == "text").strip()
    text = text.replace("```json", "").replace("```", "").strip()
    result = json.loads(text)
    result["datei"] = filename
    return result


# --------------------------- Speicherung (Qdrant) ---------------------------

import re
import uuid


def _slug(name: str) -> str:
    return re.sub(r"[^a-z0-9]+", "_", str(name).strip().lower()).strip("_") or "lokal"


def storage_on() -> bool:
    return bool(setting("QDRANT_URL"))


@st.cache_resource
def _qdrant():
    from qdrant_client import QdrantClient

    return QdrantClient(url=setting("QDRANT_URL"), api_key=setting("QDRANT_API_KEY") or None)


def _coll() -> str:
    return "rechnungen_" + _slug(st.session_state.get("kunde", "lokal"))


def _ensure():
    from qdrant_client.models import Distance, PayloadSchemaType, VectorParams

    c = _qdrant()
    if not c.collection_exists(_coll()):
        c.create_collection(
            collection_name=_coll(),
            vectors_config=VectorParams(size=1, distance=Distance.COSINE),
        )
    try:
        c.create_payload_index(
            collection_name=_coll(), field_name="datei",
            field_schema=PayloadSchemaType.KEYWORD,
        )
    except Exception:  # noqa: BLE001
        pass
    return c


def load_saved() -> list[dict]:
    if not storage_on():
        return []
    c = _ensure()
    out, offset = [], None
    while True:
        pts, offset = c.scroll(collection_name=_coll(), limit=200, with_payload=True, offset=offset)
        out.extend(p.payload for p in pts)
        if offset is None:
            return out


def save_inv(inv: dict) -> None:
    if not storage_on():
        return
    from qdrant_client.models import FieldCondition, Filter, MatchValue, PointStruct

    c = _ensure()
    c.delete(
        collection_name=_coll(),
        points_selector=Filter(must=[FieldCondition(key="datei", match=MatchValue(value=inv["datei"]))]),
    )
    c.upsert(
        collection_name=_coll(),
        points=[PointStruct(id=str(uuid.uuid4()), vector=[0.0], payload=inv)],
    )


def delete_inv(datei: str) -> None:
    st.session_state.invoices = [i for i in st.session_state.invoices if i.get("datei") != datei]
    if storage_on():
        from qdrant_client.models import FieldCondition, Filter, MatchValue

        _ensure().delete(
            collection_name=_coll(),
            points_selector=Filter(must=[FieldCondition(key="datei", match=MatchValue(value=datei))]),
        )


def clear_inv() -> None:
    st.session_state.invoices = []
    if storage_on():
        from qdrant_client.models import Distance, VectorParams

        c = _qdrant()
        c.delete_collection(_coll())
        c.create_collection(
            collection_name=_coll(),
            vectors_config=VectorParams(size=1, distance=Distance.COSINE),
        )


if "invoices" not in st.session_state:
    st.session_state.invoices = load_saved()

# ------------------------------- Verarbeitung -------------------------------

st.markdown('<div class="ln-section">Rechnungen hochladen</div>', unsafe_allow_html=True)

uploads = st.file_uploader(
    "Eingangsrechnungen als PDF oder Scan (PNG/JPG) – mehrere gleichzeitig möglich",
    type=["pdf", "png", "jpg", "jpeg", "webp"],
    accept_multiple_files=True,
)

if uploads and st.button("Rechnungen auslesen", type="primary"):
    if not api_key:
        st.error("Bitte zuerst links den Claude API-Key eintragen.")
        st.stop()
    prog = st.progress(0.0)
    fehler = []
    for i, up in enumerate(uploads, start=1):
        try:
            inv = extract_invoice(up.name, up.getvalue(), api_key)
            st.session_state.invoices = [
                r for r in st.session_state.invoices if r.get("datei") != up.name
            ] + [inv]
            save_inv(inv)
            prog.progress(i / len(uploads), text=f"{up.name} ausgelesen")
        except Exception as e:  # noqa: BLE001
            fehler.append(f"{up.name}: {e}")
    if fehler:
        st.error("Probleme bei: " + " | ".join(fehler))
    else:
        st.success(f"{len(uploads)} Rechnung(en) ausgelesen und gespeichert.")

# -------------------------------- Ergebnis ----------------------------------

st.markdown('<div class="ln-section">Ausgelesene Rechnungen</div>', unsafe_allow_html=True)

from excel_report import (  # noqa: E402
    build_excel, normalisiere, parse_date, pruefhinweise,
)


def _geld(v, w="EUR"):
    if v is None:
        return "–"
    return f"{v:,.2f} {w or ''}".replace(",", "X").replace(".", ",").replace("X", ".")


invs = [normalisiere(i) for i in st.session_state.invoices]
if not invs:
    st.caption("Noch keine – oben Rechnungen hochladen und auslesen.")
else:
    if not storage_on():
        st.caption("ℹ️ Lokaler Modus ohne Datenbank – Daten gelten nur für diese Sitzung. "
                   "Mit QDRANT_URL in den Secrets bleiben sie dauerhaft gespeichert.")

    c1, c2 = st.columns([2.4, 2.0])
    sortierung = c1.selectbox(
        "Sortieren nach",
        [
            "Fälligkeit (früheste zuerst)",
            "Fälligkeit (späteste zuerst)",
            "Betrag (höchste zuerst)",
            "Betrag (niedrigste zuerst)",
            "Rechnungsdatum (neueste zuerst)",
            "Lieferant (A–Z)",
        ],
    )
    waehrungen = sorted({(i.get("waehrung") or "?") for i in invs})
    if len(waehrungen) > 1:
        auswahl = c2.multiselect("Währung", waehrungen, default=waehrungen)
    else:
        auswahl = waehrungen
        c2.caption(f"Währung: {waehrungen[0]}")

    gefiltert = [i for i in invs if (i.get("waehrung") or "?") in auswahl]

    _spaet = parse_date("31.12.2099")
    if sortierung.startswith("Fälligkeit (frü"):
        gefiltert.sort(key=lambda i: parse_date(i.get("faelligkeitsdatum")) or _spaet)
    elif sortierung.startswith("Fälligkeit (spä"):
        gefiltert.sort(key=lambda i: parse_date(i.get("faelligkeitsdatum")) or _spaet, reverse=True)
    elif sortierung.startswith("Betrag (hö"):
        gefiltert.sort(key=lambda i: i.get("brutto_gesamt") or 0, reverse=True)
    elif sortierung.startswith("Betrag (nie"):
        gefiltert.sort(key=lambda i: i.get("brutto_gesamt") or 0)
    elif sortierung.startswith("Rechnungsdatum"):
        gefiltert.sort(key=lambda i: parse_date(i.get("rechnungsdatum")) or _spaet, reverse=True)
    else:
        gefiltert.sort(key=lambda i: (i.get("lieferant") or "").lower())

    # Summen je Währung – verschiedene Währungen werden nicht addiert
    summen_je_waehrung: dict[str, float] = {}
    for i in gefiltert:
        w = i.get("waehrung") or "?"
        summen_je_waehrung[w] = summen_je_waehrung.get(w, 0.0) + (i.get("brutto_gesamt") or 0)

    if summen_je_waehrung:
        spalten = st.columns(len(summen_je_waehrung) + 1)
        for spalte, (w, betrag) in zip(spalten, sorted(summen_je_waehrung.items())):
            anzahl = sum(1 for i in gefiltert if (i.get("waehrung") or "?") == w)
            spalte.metric(f"Summe brutto ({w})", _geld(betrag, w))
            spalte.caption(f"{anzahl} Rechnung(en)")
        spalten[-1].metric("Rechnungen gesamt", str(len(gefiltert)))
        if len(summen_je_waehrung) > 1:
            spalten[-1].caption("Summen getrennt – keine Umrechnung")

    import pandas as pd

    spalten = [
        "lieferant", "rechnungsnummer", "rechnungsdatum", "faelligkeitsdatum",
        "netto_gesamt", "ust_satz_prozent", "ust_betrag", "brutto_gesamt",
        "waehrung", "skonto", "zahlungsbedingungen", "datei",
    ]
    zeilen = []
    for i in gefiltert:
        hinweise = pruefhinweise(i)
        eintrag = {k: i.get(k) for k in spalten}
        eintrag["pruefhinweis"] = ("⚠️ " + " · ".join(hinweise)) if hinweise else ""
        zeilen.append(eintrag)
    df = pd.DataFrame(zeilen)
    df.columns = ["Lieferant", "Rechnungs-Nr.", "Datum", "Fällig am", "Netto", "USt %",
                  "USt-Betrag", "Brutto", "Währung", "Skonto", "Zahlungsbedingungen", "Datei",
                  "Prüfhinweis"]

    anzahl_auffaellig = sum(1 for i in gefiltert if pruefhinweise(i))
    if anzahl_auffaellig:
        st.warning(
            f"{anzahl_auffaellig} von {len(gefiltert)} Rechnung(en) haben einen Prüfhinweis – "
            f"siehe letzte Spalte. Es wurde nichts korrigiert, nur markiert."
        )
    else:
        st.success("Rechenproben und Datumsangaben sind bei allen Rechnungen plausibel.")

    st.dataframe(df, use_container_width=True, hide_index=True)

    with st.expander("🗂 Verwalten / Entfernen"):
        for i in gefiltert:
            ca, cb = st.columns([6, 1])
            ca.markdown(
                f"**{i.get('rechnungsnummer') or i.get('datei')}** · {i.get('lieferant') or '?'} · "
                f"{_geld(i.get('brutto_gesamt'), i.get('waehrung'))} · fällig {i.get('faelligkeitsdatum') or '–'}"
            )
            if cb.button("Entfernen", key=f"del_{i.get('datei')}"):
                delete_inv(i.get("datei"))
                st.rerun()
        st.divider()
        if st.session_state.get("confirm_clear"):
            st.warning("Wirklich ALLE Rechnungen entfernen?")
            k1, k2 = st.columns(2)
            if k1.button("Ja, alle entfernen", type="primary"):
                clear_inv()
                st.session_state.confirm_clear = False
                st.rerun()
            if k2.button("Abbrechen"):
                st.session_state.confirm_clear = False
                st.rerun()
        elif st.button("🧹 Alle entfernen"):
            st.session_state.confirm_clear = True
            st.rerun()

    st.download_button(
        "📥 Excel-Report herunterladen",
        build_excel(gefiltert),
        file_name="LN_Rechnungserfassung.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
