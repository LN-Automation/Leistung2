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

# Produkt 1 läuft auf layout="centered", das Dashboard hier braucht "wide".
# Deshalb vor set_page_config prüfen, ob überhaupt eine Anmeldung ansteht –
# nur dann schmal. Früher hing das an auth_ok allein; ohne gesetztes Passwort
# wurde der Login übersprungen und die Seite blieb fälschlich schmal.
def _login_noetig() -> bool:
    try:
        kunden = dict(st.secrets.get("kunden", {}) or {})
        pw = st.secrets.get("APP_PASSWORD", "")
    except Exception:  # noqa: BLE001
        kunden, pw = {}, ""
    return bool(kunden or pw) and not st.session_state.get("auth_ok")


st.set_page_config(
    page_title="LN Automation – Rechnungserfassung",
    page_icon="logo.png" if _LOGO_PFAD.exists() else "🧾",
    layout="centered" if _login_noetig() else "wide",
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

def _logo_block(breite: int = 230, untertitel: str = "") -> str:
    """HTML-Kopf mit zentriertem Logo. Ohne Logo-Datei: Schriftzug.
    Bewusst identisch zu Produkt 1, damit beide Apps gleich aussehen."""
    if _LOGO_PFAD.exists():
        b64 = base64.b64encode(_LOGO_PFAD.read_bytes()).decode()
        inneres = f'<img src="data:image/png;base64,{b64}" style="width:{breite}px;max-width:70%;" />'
    else:
        inneres = '<div style="color:#0f172a;font-size:2rem;font-weight:800;">LN Automation</div>'
    unter = (f'<div style="color:#64748b;font-size:1.0rem;margin-top:4px;">{untertitel}</div>'
             if untertitel else "")
    # Trennlinie bewusst nur unter dem Schriftzug, nicht über die ganze
    # Breite – sonst zerschneidet sie das Dashboard optisch.
    return (f'<div style="text-align:center;padding:26px 0 10px 0;">'
            f'{inneres}{unter}'
            f'<hr style="border:none;border-top:1px solid #e2e8f0;'
            f'width:min(560px,60%);margin:22px auto 0 auto;" />'
            f'</div>')


# ------------------------------ Anmeldung -----------------------------------

_kunden: dict[str, str] = {}
try:
    if "kunden" in st.secrets:
        _kunden = {str(k).lower(): str(v) for k, v in dict(st.secrets["kunden"]).items()}
except Exception:  # noqa: BLE001
    pass
_app_pw = setting("APP_PASSWORD")

if (_kunden or _app_pw) and not st.session_state.get("auth_ok"):
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
    _logo_block(190, "KI-Rechnungserfassung – Eingangsrechnungen automatisch auslesen")
    + '<div style="height:26px;"></div>',
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


MODELL = "claude-sonnet-4-6"


def extract_invoice(filename: str, data: bytes, key: str) -> tuple[dict, dict]:
    """Rechnung auslesen. Liefert (Daten, Verbrauchsdatensatz)."""
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
        model=MODELL,
        max_tokens=3000,
        messages=[{"role": "user", "content": [block, {"type": "text", "text": EXTRAKTIONS_PROMPT}]}],
    )
    text = "".join(b.text for b in msg.content if b.type == "text").strip()
    text = text.replace("```json", "").replace("```", "").strip()
    result = json.loads(text)
    result["datei"] = filename
    from ln_nutzung import datensatz

    return result, datensatz(filename, MODELL, getattr(msg, "usage", None))


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


def _coll_einst() -> str:
    return "einstellungen_" + _slug(st.session_state.get("kunde", "lokal"))


def einstellungen_laden() -> dict:
    """DATEV-Einstellungen des Kunden. Ohne Qdrant nur für die Sitzung."""
    if not storage_on():
        return st.session_state.get("_einst_lokal", {})
    from qdrant_client.models import Distance, VectorParams

    c = _qdrant()
    if not c.collection_exists(_coll_einst()):
        c.create_collection(
            collection_name=_coll_einst(),
            vectors_config=VectorParams(size=1, distance=Distance.COSINE),
        )
        return {}
    try:
        pts = c.retrieve(collection_name=_coll_einst(), ids=[1],
                         with_payload=True)
        return dict(pts[0].payload) if pts else {}
    except Exception:  # noqa: BLE001
        return {}


def einstellungen_speichern(daten: dict) -> None:
    st.session_state["_einst_lokal"] = daten
    if not storage_on():
        return
    from qdrant_client.models import Distance, PointStruct, VectorParams

    c = _qdrant()
    if not c.collection_exists(_coll_einst()):
        c.create_collection(
            collection_name=_coll_einst(),
            vectors_config=VectorParams(size=1, distance=Distance.COSINE),
        )
    c.upsert(collection_name=_coll_einst(),
             points=[PointStruct(id=1, vector=[0.0], payload=daten)])


def _coll_nutzung() -> str:
    return "nutzung_" + _slug(st.session_state.get("kunde", "lokal"))


def nutzung_speichern(satz: dict) -> None:
    """Einen Verbrauchsdatensatz ablegen. Ohne Qdrant nur für die Sitzung."""
    st.session_state.setdefault("_nutzung_lokal", []).append(satz)
    if not storage_on():
        return
    from qdrant_client.models import Distance, PointStruct, VectorParams

    c = _qdrant()
    if not c.collection_exists(_coll_nutzung()):
        c.create_collection(
            collection_name=_coll_nutzung(),
            vectors_config=VectorParams(size=1, distance=Distance.COSINE),
        )
    c.upsert(collection_name=_coll_nutzung(),
             points=[PointStruct(id=str(uuid.uuid4()), vector=[0.0],
                                 payload=satz)])


def nutzung_laden() -> list[dict]:
    lokal = st.session_state.get("_nutzung_lokal", [])
    if not storage_on():
        return lokal
    c = _qdrant()
    if not c.collection_exists(_coll_nutzung()):
        return lokal
    out, offset = [], None
    while True:
        pts, offset = c.scroll(collection_name=_coll_nutzung(), limit=500,
                               with_payload=True, offset=offset)
        out.extend(p.payload for p in pts)
        if offset is None:
            return out


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

# --------------------------------- Reiter -----------------------------------

from excel_report import (  # noqa: E402
    build_excel, normalisiere, parse_date, pruefhinweise,
)
from ln_dashboard import (  # noqa: E402
    dashboard_bild, kurzueberblick, render_dashboard,
)
from ln_datev import render_datev  # noqa: E402
from ln_nutzung import render_nutzung  # noqa: E402


def _geld(v, w="EUR"):
    if v is None:
        return "–"
    return f"{v:,.2f} {w or ''}".replace(",", "X").replace(".", ",").replace("X", ".")


invs = [normalisiere(i) for i in st.session_state.invoices]

tab_hoch, tab_liste, tab_dash, tab_export, tab_nutzung = st.tabs(
    ["Hochladen", "Rechnungen", "Dashboard", "Exporte", "Nutzung"]
)

# ------------------------------ Reiter: Hochladen ---------------------------

with tab_hoch:
    st.markdown('<div class="ln-section">Rechnungen hochladen</div>',
                unsafe_allow_html=True)

    meldung = st.session_state.pop("letzte_meldung", None)
    if meldung:
        gelungen, misslungen, fehler = meldung
        if gelungen:
            st.success(f"{gelungen} von {gelungen + misslungen} Rechnung(en) "
                       f"erfolgreich ausgelesen und gespeichert.")
        if misslungen:
            st.error(f"{misslungen} Rechnung(en) konnten nicht gelesen werden: "
                     + " | ".join(fehler[:5])
                     + (" …" if len(fehler) > 5 else ""))

    uploads = st.file_uploader(
        "Eingangsrechnungen als PDF oder Scan (PNG/JPG) – mehrere gleichzeitig möglich",
        type=["pdf", "png", "jpg", "jpeg", "webp"],
        accept_multiple_files=True,
    )

    if uploads and st.button("Rechnungen auslesen", type="primary",
                         width="stretch"):
        if not api_key:
            st.error("Bitte zuerst links den Claude API-Key eintragen.")
            st.stop()
        prog = st.progress(0.0)
        fehler = []
        for i, up in enumerate(uploads, start=1):
            try:
                inv, verbrauch = extract_invoice(up.name, up.getvalue(), api_key)
                st.session_state.invoices = [
                    r for r in st.session_state.invoices if r.get("datei") != up.name
                ] + [inv]
                save_inv(inv)
                nutzung_speichern(verbrauch)
                prog.progress(i / len(uploads), text=f"{up.name} ausgelesen")
            except Exception as e:  # noqa: BLE001
                fehler.append(f"{up.name}: {e}")
        gelungen = len(uploads) - len(fehler)
        st.session_state["letzte_meldung"] = (gelungen, len(fehler), fehler)
        st.rerun()

    if invs:
        st.divider()
        st.markdown('<div class="ln-section">Aktueller Bestand</div>',
                    unsafe_allow_html=True)
        kurzueberblick(invs)
        st.caption("Vollständige Auswertung im Reiter „Dashboard“.")
    if not storage_on():
        st.caption("ℹ️ Lokaler Modus ohne Datenbank – Daten gelten nur für diese "
                   "Sitzung. Mit QDRANT_URL in den Secrets bleiben sie dauerhaft "
                   "gespeichert.")

# ------------------------------ Reiter: Rechnungen --------------------------

with tab_liste:
    st.markdown('<div class="ln-section">Ausgelesene Rechnungen</div>',
                unsafe_allow_html=True)

    if not invs:
        st.caption("Noch keine – im Reiter „Hochladen“ starten.")
    else:
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
            gefiltert.sort(key=lambda i: parse_date(i.get("faelligkeitsdatum")) or _spaet,
                           reverse=True)
        elif sortierung.startswith("Betrag (hö"):
            gefiltert.sort(key=lambda i: i.get("brutto_gesamt") or 0, reverse=True)
        elif sortierung.startswith("Betrag (nie"):
            gefiltert.sort(key=lambda i: i.get("brutto_gesamt") or 0)
        elif sortierung.startswith("Rechnungsdatum"):
            gefiltert.sort(key=lambda i: parse_date(i.get("rechnungsdatum")) or _spaet,
                           reverse=True)
        else:
            gefiltert.sort(key=lambda i: (i.get("lieferant") or "").lower())

        anzahl_auffaellig = sum(1 for i in gefiltert if pruefhinweise(i))
        if anzahl_auffaellig:
            st.warning(
                f"{anzahl_auffaellig} von {len(gefiltert)} Rechnung(en) haben einen "
                f"Prüfhinweis. Die Hinweise stammen aus festen Rechenregeln, nicht "
                f"aus der Texterkennung. Es wurde nichts korrigiert, nur markiert."
            )
        else:
            st.success("Rechenproben und Datumsangaben sind bei allen Rechnungen "
                       "plausibel.")

        import pandas as pd

        zeilen = []
        for i in gefiltert:
            hinweise = pruefhinweise(i)
            zeilen.append({
                "Rechnungs-Nr.": i.get("rechnungsnummer") or "",
                "Lieferant": i.get("lieferant") or "–",
                "Datum": i.get("rechnungsdatum") or "",
                "Fällig am": i.get("faelligkeitsdatum") or "",
                "Netto": i.get("netto_gesamt"),
                "USt %": i.get("ust_satz_prozent"),
                "USt-Betrag": i.get("ust_betrag"),
                "Brutto": i.get("brutto_gesamt"),
                "Währung": i.get("waehrung") or "",
                "Skonto": i.get("skonto") or "",
                "Zahlungsbedingungen": i.get("zahlungsbedingungen") or "",
                "Datei": i.get("datei") or "",
                "Prüfhinweis": " · ".join(hinweise),
            })
        df = pd.DataFrame(zeilen)

        geld = st.column_config.NumberColumn(format="%.2f", width="small")
        st.dataframe(
            df,
            width="stretch",
            hide_index=True,
            column_config={
                "Rechnungs-Nr.": st.column_config.TextColumn(width="small"),
                "Lieferant": st.column_config.TextColumn(width="medium"),
                "Netto": geld,
                "USt-Betrag": geld,
                "Brutto": st.column_config.NumberColumn(format="%.2f", width="small"),
                "USt %": st.column_config.NumberColumn(format="%.0f %%", width="small"),
                "Zahlungsbedingungen": st.column_config.TextColumn(width="medium"),
                "Prüfhinweis": st.column_config.TextColumn(width="large"),
            },
        )

        with st.expander("Verwalten / Entfernen"):
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
            elif st.button(f"Alle {len(gefiltert)} Rechnungen entfernen"):
                st.session_state.confirm_clear = True
                st.rerun()
            st.divider()
            for i in gefiltert:
                ca, cb = st.columns([6, 1])
                ca.markdown(
                    f"**{i.get('rechnungsnummer') or i.get('datei')}** · "
                    f"{i.get('lieferant') or '?'} · "
                    f"{_geld(i.get('brutto_gesamt'), i.get('waehrung'))} · "
                    f"fällig {i.get('faelligkeitsdatum') or '–'}"
                )
                if cb.button("Entfernen", key=f"del_{i.get('datei')}"):
                    delete_inv(i.get("datei"))
                    st.rerun()
        st.caption("Alle Downloads liegen im Reiter „Exporte“.")

# ------------------------------ Reiter: Dashboard ---------------------------

with tab_dash:
    st.markdown('<div class="ln-section">Auswertung</div>', unsafe_allow_html=True)
    render_dashboard(invs)


# ------------------------------- Reiter: Exporte ----------------------------

with tab_export:
    st.markdown('<div class="ln-section">Exporte</div>', unsafe_allow_html=True)

    if not invs:
        st.info("Noch keine Rechnungen – im Reiter „Hochladen“ starten.")
    else:
        e_excel, e_png, e_datev = st.tabs(
            ["Excel-Report", "Grafiken als Bild", "DATEV-Buchungsstapel"]
        )

        with e_excel:
            st.markdown("**Excel-Report**")
            st.caption("Sechs Blätter: Übersicht, Rechnungen, Positionen, "
                       "Prüfliste, USt-Auswertung und Analyse je Währung.")
            st.download_button(
                "Excel-Report herunterladen",
                build_excel(invs),
                file_name="LN_Rechnungserfassung.xlsx",
                mime="application/vnd.openxmlformats-officedocument."
                     "spreadsheetml.sheet",
                type="primary",
                key="dl_excel",
            )

        with e_png:
            st.markdown("**Grafiken als Bild**")
            st.caption("Alle Dashboard-Grafiken in einem PNG, im selben Aufbau. "
                       "Gut zum Einfügen in eine Mail oder Präsentation.")
            if (st.session_state.get("dash_png_fingerabdruck")
                    != st.session_state.get("dash_fingerabdruck")):
                st.session_state["dash_png"] = None
            if st.session_state.get("dash_png"):
                st.download_button(
                    "PNG herunterladen",
                    st.session_state["dash_png"],
                    file_name=f"LN_Dashboard_"
                              f"{st.session_state.get('dash_export_waehrung', 'EUR')}"
                              f".png",
                    mime="image/png",
                    type="primary",
                    key="dl_png",
                )
            else:
                if st.button("PNG erstellen", key="btn_png"):
                    with st.spinner("Bild wird erzeugt …"):
                        try:
                            st.session_state["dash_png"] = dashboard_bild(
                                st.session_state.get("dash_grafiken") or [],
                                st.session_state.get("dash_kopf", ""),
                            )
                            st.session_state["dash_png_fingerabdruck"] = \
                                st.session_state.get("dash_fingerabdruck")
                            st.session_state["dash_png_fehler"] = None
                            st.rerun()
                        except Exception as ex:  # noqa: BLE001
                            st.session_state["dash_png_fehler"] = str(ex)
                if st.session_state.get("dash_png_fehler"):
                    st.warning(st.session_state["dash_png_fehler"])

        with e_datev:
            render_datev(invs, einstellungen_laden, einstellungen_speichern)


# ------------------------------- Reiter: Nutzung ----------------------------

with tab_nutzung:
    st.markdown('<div class="ln-section">Verbrauch und Kosten</div>',
                unsafe_allow_html=True)
    st.caption(
        "Jedes Auslesen einer Rechnung verbraucht Tokens beim KI-Anbieter. "
        "Hier steht, wie viel – nichts davon ist geschätzt, die Zahlen kommen "
        "direkt aus der Antwort der Schnittstelle."
    )
    render_nutzung(nutzung_laden(), einstellungen_laden, einstellungen_speichern)
