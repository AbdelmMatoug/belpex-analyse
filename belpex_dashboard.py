# -*- coding: utf-8 -*-
"""
Scholt — Belpex kwartierprijzen
===============================
Interactieve UI voor de analyse van Belgische kwartierprijzen, opgezet als
een stap-voor-stap wizard (geen zijbalk): elke instelling staat op de stap
waar ze nodig is, en de "Volgende"-knop blokkeert zolang de vereiste data
voor die stap ontbreekt.

Drie stappen: data & periode (inclusief de piek/dal-verhouding), verbruik &
vergoeding, export. De piek/dal-keuze op stap 1 werkt door in alles wat
daarna komt — kerncijfers, eenheidsprijs piek/dal en de Excel-werkmap.

Starten:
    pip install streamlit plotly pandas numpy openpyxl xlsxwriter
    streamlit run belpex_dashboard.py

De browser opent automatisch op http://localhost:8501
"""

from __future__ import annotations

import io
import sys
from datetime import date
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from belpex_excel import build_workbook

# Bij een ingepakte .exe (PyInstaller) hangt het werkmap-pad niet af van de
# map waaruit de gebruiker toevallig dubbelklikt — de databank hoort dan
# naast de .exe te staan, niet in een onvoorspelbare CWD.
if getattr(sys, "frozen", False):
    BASISMAP = Path(sys.executable).resolve().parent
else:
    BASISMAP = Path(__file__).resolve().parent

ELEXYS_URL = ("https://www.elexys.be/en/insights/quarter-hourly-belpex-day-ahead-spot-be"
              "?from={van}&until={tot}")

STORE = BASISMAP / "belpex_databank.csv"   # lokale prijzendatabank

# Belpex noteert pas kwartierprijzen vanaf deze datum; alles daarvoor is
# uurdata en bestaat niet op de kwartierpagina van Elexys.
KWARTIER_START = date(2025, 8, 1)

STAPPEN = ["Data & periode", "Verbruik & vergoeding", "Export"]

# Piek/dal-verhoudingen. De waarden zijn (eerste piekuur, eerste daluur na de
# piek): (8, 20) betekent piek van 08:00 t.e.m. 19:45. De gekozen verhouding
# bepaalt alles — de kerncijfers, de eenheidsprijs piek/dal én de werkmap.
PIEKPROFIELEN = {
    "Standaard — 07:00 t.e.m. 21:45 (15u)": (7, 22),
    "Scholt — 08:00 t.e.m. 19:45 (12u)": (8, 20),
    "Zelf kiezen": None,
}

EENHEDEN = {"EUR/kWh": 1, "cent/kWh": 100, "EUR/MWh": 1000}
DECIMALEN = {"EUR/kWh": 4, "cent/kWh": 2, "EUR/MWh": 2}

# --------------------------------------------------------------------------
# PAGINA-INSTELLINGEN
# --------------------------------------------------------------------------

st.set_page_config(page_title="Scholt — kwartierprijzen", page_icon=":material/bolt:",
                   layout="wide")

# Twee tabbladen bovenaan: de rekentool (deze file) en de handleiding.
# De rekentool blijft gewone code op moduleniveau; enkel voor de handleiding
# wordt een andere pagina uitgevoerd, waarna dit script stopt.
def _rekentool():
    pass


_pagina = st.navigation(
    [st.Page(_rekentool, title="Rekentool", icon=":material/calculate:", default=True),
     st.Page("handleiding.py", title="Handleiding", icon=":material/menu_book:")],
    position="top")
if _pagina.title == "Handleiding":
    _pagina.run()
    st.stop()

# Huisstijl Scholt Energy (scholt.be): rood accent op donkerblauw.
ROOD, NAVY, DONKER, ZAND, GREY = "#e61d3a", "#030737", "#171b25", "#edeae4", "#7f8c8d"
ZAND_RAND = "#c9c3b8"   # rand rond zandkleurige vlakken in grafieken


def nl(x: float, dec: int = 0) -> str:
    """Belgische notatie: 140.537 en 0,1930 (niet 140,537 en 0.1930)."""
    return f"{x:,.{dec}f}".replace(",", "\0").replace(".", ",").replace("\0", ".")

# Compatibiliteit: Streamlit >= 1.60 gebruikt width="stretch" i.p.v. use_container_width
try:
    _v = tuple(int(x) for x in st.__version__.split(".")[:2])
    FULL = {"width": "stretch"} if _v >= (1, 60) else {"use_container_width": True}
except Exception:
    FULL = {"use_container_width": True}

st.markdown("""
<style>
  /* Werkblad-dichtheid: één kolom van 1100px, 8px-raster, rood = actie of piek. */
  :root {
    --inkt: #171b25;        /* tekst */
    --inkt-zacht: #5b6170;  /* labels */
    --inkt-stil: #8a8f9b;   /* metadata */
    --zand: #edeae4;        /* vlakken, dal */
    --lijn: rgba(23,27,37,.10);
    --piek: #e61d3a;        /* Scholt-rood: piek én primaire actie */
    --nacht: #030737;       /* Scholt-donkerblauw */
    --papier: #ffffff;      /* kaartvlak */
  }
  html, body {-webkit-font-smoothing: antialiased;}
  .block-container {padding-top: 1.5rem; max-width: 1100px;}
  h1 {font-size: 1.75rem !important; letter-spacing: -0.01em; text-wrap: balance;}
  [data-testid="stMetricValue"] {font-size: 1.35rem; font-variant-numeric: tabular-nums;}
  [data-testid="stMetricLabel"] p {color: var(--inkt-zacht);}

  /* Stappenbalk: drie benoemde stappen, de actieve in rood. */
  .stappen {display: flex; gap: 8px; margin: 4px 0 24px;}
  .stappen div {flex: 1; padding: 8px 0 0; border-top: 3px solid var(--zand);
                color: var(--inkt-stil); font-size: .85rem; font-weight: 500;}
  .stappen div b {display: block; font-size: .7rem; font-weight: 600;
                  letter-spacing: .06em; text-transform: uppercase;}
  .stappen .klaar {border-top-color: var(--nacht); color: var(--inkt-zacht);}
  .stappen .nu {border-top-color: var(--piek); color: var(--inkt);}

  /* Dagbalk: 96 kwartieren van één dag, piek in rood. */
  .dagbalk {display: grid; grid-template-columns: repeat(96, 1fr); gap: 1px;
            height: 22px; border-radius: 4px; overflow: hidden; margin-top: 4px;}
  .dagbalk span {background: var(--zand);}
  .dagbalk span.p {background: var(--piek);}
  .dagbalk-uren {display: flex; justify-content: space-between; margin-top: 4px;
                 font-size: .72rem; color: var(--inkt-stil); font-variant-numeric: tabular-nums;}
  .dagbalk-uitleg {font-size: .8rem; color: var(--inkt-zacht); margin-top: 6px;}
  .dagbalk-uitleg i {display: inline-block; width: 10px; height: 10px; border-radius: 2px;
                     margin: 0 4px -1px 0; background: var(--zand);}
  .dagbalk-uitleg i.p {background: var(--piek);}

  /* Uitkomst: het antwoord van de tool, groot en als eerste. */
  .uitkomst {padding: 20px 24px; border-radius: 8px; background: var(--papier); margin-bottom: 8px;
             box-shadow: 0 0 0 1px var(--lijn), 0 1px 2px rgba(23,27,37,.05),
                         0 4px 12px rgba(23,27,37,.04);}
  .uitkomst-label {font-size: .72rem; font-weight: 600; letter-spacing: .06em;
                   text-transform: uppercase; color: var(--inkt-zacht);}
  .uitkomst-waarde {font-size: 2.75rem; font-weight: 700; line-height: 1.1; color: var(--inkt);
                    letter-spacing: -0.02em; font-variant-numeric: tabular-nums; margin-top: 4px;}
  .uitkomst-waarde small {font-size: 1rem; font-weight: 500; color: var(--inkt-zacht);
                          letter-spacing: 0; margin-left: 6px;}
  .uitkomst-rij {display: flex; flex-wrap: wrap; gap: 8px 32px; margin-top: 16px; padding-top: 16px;
                 border-top: 1px solid var(--lijn); font-size: .9rem; color: var(--inkt-zacht);}
  .uitkomst-rij b {color: var(--inkt); font-weight: 600; font-variant-numeric: tabular-nums;}
  .uitkomst-rij i {display: inline-block; width: 10px; height: 10px; border-radius: 2px;
                   margin: 0 6px -1px 0; background: var(--zand);}
  .uitkomst-rij i.p {background: var(--piek);}
</style>
""", unsafe_allow_html=True)


def stappenbalk(stap: int) -> None:
    delen = []
    for i, naam in enumerate(STAPPEN, start=1):
        klasse = "nu" if i == stap else ("klaar" if i < stap else "")
        delen.append(f'<div class="{klasse}"><b>Stap {i}</b>{naam}</div>')
    st.markdown(f'<div class="stappen">{"".join(delen)}</div>', unsafe_allow_html=True)


def dagbalk(peak_start: int, peak_end: int, weekdays_only: bool) -> None:
    """Toont welke kwartieren van een dag als piek tellen."""
    cellen = "".join(f'<span class="{"p" if peak_start <= q // 4 < peak_end else ""}"></span>'
                     for q in range(96))
    uren = "".join(f"<span>{u:02d}:00</span>" for u in (0, 6, 12, 18)) + "<span>24:00</span>"
    dagen = ("ma–vr; zaterdag en zondag zijn volledig dal" if weekdays_only
             else "elke dag van de week")
    st.markdown(
        f'<div class="dagbalk">{cellen}</div><div class="dagbalk-uren">{uren}</div>'
        f'<div class="dagbalk-uitleg"><i class="p"></i>piek {peak_start:02d}:00 t.e.m. '
        f'{peak_end - 1:02d}:45 ({peak_end - peak_start} uur), {dagen} &nbsp; '
        f'<i></i>dal: al de rest</div>', unsafe_allow_html=True)


# --------------------------------------------------------------------------
# DATA INLEZEN
# --------------------------------------------------------------------------

def to_numeric(series: pd.Series) -> pd.Series:
    """Zet tekst met EU-notatie ('1.234,56 €') om naar float."""
    if pd.api.types.is_numeric_dtype(series):
        return series.astype(float)
    s = series.astype(str).str.replace(r"[^\d,.\-]", "", regex=True)
    both = s.str.contains(r"\.") & s.str.contains(",")
    s = s.mask(both, s.str.replace(".", "", regex=False))
    return pd.to_numeric(s.str.replace(",", ".", regex=False), errors="coerce")


def _parse_datetime(series: pd.Series) -> pd.Series:
    # ISO (2026-09-01) eerst: met dayfirst=True zou dat 9 januari worden.
    ts = pd.to_datetime(series, errors="coerce", format="ISO8601")
    if ts.notna().mean() > 0.5 and not pd.api.types.is_datetime64_dtype(ts):
        # Tijdzone-offsets (+02:00 / +01:00 rond de zomeruurwissel): omzetten
        # naar Belgische lokale tijd zonder zone, net als de Belpex-prijzen.
        ts = (pd.to_datetime(series, errors="coerce", format="ISO8601", utc=True)
              .dt.tz_convert("Europe/Brussels").dt.tz_localize(None))
    if ts.isna().mean() > 0.5:
        ts = pd.to_datetime(series, errors="coerce", dayfirst=True)
    if ts.isna().mean() > 0.5:
        ts = pd.to_datetime(series, errors="coerce", dayfirst=False)
    return ts


_DATUM_PATROON = r"\d{1,4}[-/.]\d{1,2}[-/.]\d{1,4}"   # 1/09/2026, 2026-09-01, 01.09.26
_UUR_PATROON = r"^\s*\d{1,2}[:uh]\d{2}"                # 00:15, 00u15, 0h15


def _kies_kolommen(df: pd.DataFrame, value_keywords) -> tuple[str, str]:
    """
    Kies tijdstempel- en waardekolom op basis van de INHOUD, niet enkel de naam.
    Exports verschillen sterk (Fluvius 'From;To;TimeFrame;Value;Unit', Elexys,
    eigen Excel-lijsten...): een kolom heet bv. 'TimeFrame' maar bevat
    'Total hours'. Daarom: tijdkolom = eerste kolom waarvan de meeste waarden
    als datum/tijd lezen; waardekolom = numerieke kolom, bij voorkeur met een
    herkenbare naam. Bij twee tijdkolommen (van/tot) wint de eerste: het begin
    van het kwartier.
    """
    staal = df.head(200)

    def is_tijd(c) -> bool:
        if pd.api.types.is_datetime64_any_dtype(staal[c]):
            return True
        tekst = staal[c].astype(str)
        return (tekst.str.contains(_DATUM_PATROON).mean() > 0.8
                and _parse_datetime(tekst).notna().mean() > 0.8)

    def is_getal(c) -> bool:
        tekst = staal[c].astype(str)
        return (not tekst.str.contains(_UUR_PATROON).mean() > 0.5
                and to_numeric(staal[c]).notna().mean() > 0.8)

    tijd_kols = [c for c in df.columns if is_tijd(c)]
    if not tijd_kols:
        raise ValueError("Geen kolom met datum/tijd gevonden.")
    dt_col = tijd_kols[0]

    num_kols = [c for c in df.columns if c not in tijd_kols and is_getal(c)]
    if not num_kols:
        raise ValueError("Geen kolom met getallen gevonden.")
    voorkeur = [c for c in num_kols if any(k in c.lower() for k in value_keywords)]
    return dt_col, voorkeur[0] if voorkeur else num_kols[-1]


def normalise(df: pd.DataFrame, dt_col=None, value_col=None, value_name="price",
              value_keywords=("price", "prijs", "eur", "€", "value", "waarde")) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    if dt_col is None or value_col is None:
        auto_dt, auto_val = _kies_kolommen(df, value_keywords)
        dt_col = dt_col or auto_dt
        value_col = value_col or auto_val

    out = pd.DataFrame({"datetime": _parse_datetime(df[dt_col]),
                        value_name: to_numeric(df[value_col])})
    out = out.dropna().drop_duplicates("datetime").sort_values("datetime")
    out = out.reset_index(drop=True)
    out.attrs["bron"] = (dt_col, value_col)
    return out


# Eenheid van verbruik -> factor naar kWh.
EENHEID_NAAR_KWH = {"kwh": 1.0, "wh": 0.001, "mwh": 1000.0}


def normalise_verbruik(df: pd.DataFrame, dt_col=None, verbruik_col=None,
                       eenheid: str | None = None) -> pd.DataFrame:
    """
    Zet eender welk kwartierverbruik om naar één standaardvorm:
    kolommen `datetime` (begin van het kwartier) en `verbruik` (kWh).

    Bovenop normalise():
    - rijen met injectie (Fluvius-register 'Injectie ...') vallen weg;
    - meerdere rijen per kwartier (bv. registers 'Afname Dag' en 'Afname
      Nacht') worden opgeteld in plaats van op één na weggegooid;
    - een eenheidskolom met Wh of MWh wordt omgerekend naar kWh, tenzij
      `eenheid` ("kwh"/"wh"/"mwh") expliciet is opgegeven (manuele keuze).
    """
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    trefwoorden = ("verbruik", "consumption", "kwh", "energie", "energy", "volume",
                   "hoeveelheid", "afname", "value", "waarde")

    register_col = next((c for c in df.columns if any(k in c.lower() for k in
                         ("register", "richting", "direction", "type"))), None)
    if register_col is not None:
        df = df[~df[register_col].astype(str).str.lower().str.contains("injec")]

    if dt_col is None or verbruik_col is None:
        auto_dt, auto_val = _kies_kolommen(df, trefwoorden)
        dt_col = dt_col or auto_dt
        verbruik_col = verbruik_col or auto_val

    waarde = to_numeric(df[verbruik_col])
    eenheid_col = next((c for c in df.columns if c.lower() in ("unit", "eenheid")), None)
    if eenheid is not None:
        waarde = waarde * EENHEID_NAAR_KWH[eenheid]
    elif eenheid_col is not None:
        factor = df[eenheid_col].astype(str).str.strip().str.lower().map(EENHEID_NAAR_KWH)
        waarde = waarde * factor.fillna(1.0)

    out = pd.DataFrame({"datetime": _parse_datetime(df[dt_col]), "verbruik": waarde}).dropna()
    out = (out.groupby("datetime", as_index=False)["verbruik"].sum()
           .sort_values("datetime").reset_index(drop=True))
    out.attrs["bron"] = (dt_col, verbruik_col)
    return out


def verbruik_melding(vbr: pd.DataFrame, d: pd.DataFrame) -> tuple[str, str]:
    """(niveau, tekst) die toont wat er uit het verbruiksbestand gelezen werd."""
    tijd_kol, waarde_kol = vbr.attrs.get("bron", ("?", "?"))
    stap_min = vbr["datetime"].diff().median().total_seconds() / 60
    if tijd_kol == "_datetime":
        tijd_kol = "datum + uur"
    tekst = (f"Herkend: tijdstip uit kolom **{tijd_kol}**, verbruik uit "
             f"kolom **{waarde_kol}** — {nl(len(vbr))} rijen, "
             f"{vbr['datetime'].min():%d/%m/%Y} t/m {vbr['datetime'].max():%d/%m/%Y}, "
             f"totaal {nl(vbr['verbruik'].sum())} kWh.")
    if len(vbr) > 1 and stap_min != 15:
        tekst += (f" Let op: de rijen liggen {stap_min:.0f} min uit elkaar, geen kwartier: "
                  "alleen tijdstippen die exact op een kwartier vallen tellen mee.")
    gematcht = int(d["datetime"].isin(vbr["datetime"]).sum())
    if gematcht == 0:
        return ("warning", tekst + " Maar geen enkele tijdstempel kwam overeen met de "
                "kwartieren in de gekozen periode — alle verbruik staat op 0. "
                "Klopt de periode op stap 1, of de gekozen kolommen hieronder?")
    if gematcht == len(d):
        return ("success", tekst + f" Alle {nl(gematcht)} kwartieren van de periode gematcht.")
    return ("success", tekst + f" {nl(gematcht)} van de {nl(len(d))} kwartieren gematcht "
            f"({gematcht/len(d)*100:.0f}%). De overige kwartieren zonder match kregen verbruik = 0.")


@st.cache_data(show_spinner=False)
def read_upload(content: bytes, name: str, overslaan: int = 0,
                kopregel: bool = True) -> pd.DataFrame:
    """Lees CSV/Excel. `overslaan` = lijnen boven de tabel; zonder kopregel krijgen
    de kolommen een nummer."""
    buf = io.BytesIO(content)
    header = 0 if kopregel else None
    if name.lower().endswith((".xlsx", ".xls", ".xlsm")):
        return pd.read_excel(buf, skiprows=overslaan, header=header)
    for sep in (";", ",", "\t", "|"):
        buf.seek(0)
        try:
            trial = pd.read_csv(buf, sep=sep, engine="python", skiprows=overslaan, header=header)
            if trial.shape[1] >= 2:
                return trial
        except Exception:
            continue
    raise ValueError("Kon het bestand niet inlezen als CSV of Excel.")


def _combineer_datum_tijd(df: pd.DataFrame) -> pd.DataFrame:
    """
    Sommige exports (Elexys, meterdata) splitsen 'Date' en 'Time' in twee
    aparte kolommen. normalise() pikt dan alleen 'Date' op als tijdstempel
    (eerste kolom die op 'date' matcht) en negeert 'Time' volledig — alle
    rijen krijgen zo dezelfde dag-stempel en verdwijnen op één na bij het
    ontdubbelen. Plak ze daarom eerst samen.
    """
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    datum_col = next((c for c in df.columns
                      if any(k in c.lower() for k in ("date", "datum"))), None)
    tijd_col = next((c for c in df.columns if c != datum_col
                     and any(k in c.lower() for k in ("time", "tijd", "hour", "uur"))), None)
    if datum_col is not None and tijd_col is not None:
        tijd = (df[tijd_col].astype(str)
                .str.replace("u", ":", regex=False)
                .str.replace("h", ":", regex=False)
                .str.strip())
        df["_datetime"] = df[datum_col].astype(str).str.strip() + " " + tijd
        df.attrs["datum_tijd"] = (datum_col, tijd_col)
    return df


def _fetch_web_tabel(sessie, headers, van: date, tot: date) -> pd.DataFrame:
    """
    Terugvaloptie: leest de zichtbare tabel op de Elexys-pagina zelf. Die
    toont maar de laatste ~50 kwartieren (~12,5 uur) tot en met 'tot' — 'van'
    wordt door de site genegeerd. Wordt enkel gebruikt als de Excel-export
    (zie fetch_web) niet gevonden wordt.
    """
    url = ELEXYS_URL.format(van=van.isoformat(), tot=tot.isoformat())
    r = sessie.get(url, headers=headers, timeout=60)
    r.raise_for_status()

    tabellen = pd.read_html(io.StringIO(r.text))
    if not tabellen:
        raise ValueError("Geen tabel gevonden op de pagina.")
    grootste = max(tabellen, key=len)
    if len(grootste) < 10:
        raise ValueError("De tabel is leeg of wordt via JavaScript geladen.")
    grootste = _combineer_datum_tijd(grootste)
    out = normalise(grootste, dt_col="_datetime" if "_datetime" in grootste.columns else None)
    if out.empty:
        raise ValueError("De tabel bevatte geen bruikbare datum- en prijskolom.")
    return out


def _fetch_web_periode(van: date, tot: date) -> pd.DataFrame:
    """
    Haalt de kwartierprijzen op via de Excel-exportknop op de Elexys-pagina
    (niet de zichtbare tabel — die toont maar de laatste ~50 kwartieren).
    Draait op jouw machine met jouw IP, dus dit kan lukken waar een server
    faalt. Lukt het niet, dan blijft uploaden altijd werken.

    Let op: de exportknop levert per opvraging één aaneengesloten reeks van
    Elexys (voor het lopende jaar is dat de volledige year-to-date-reeks,
    ongeacht 'van'/'tot' in de query). 'van'/'tot' worden hier daarom ook
    client-side als filter op het resultaat toegepast. Een bereik over
    meerdere kalenderjaren wordt door fetch_web() jaar per jaar opgevraagd.
    """
    import re
    import requests

    headers = {
        "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                       "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"),
        "Accept-Language": "nl-BE,nl;q=0.9,en;q=0.8",
    }
    sessie = requests.Session()
    pagina_url = ELEXYS_URL.format(van=van.isoformat(), tot=tot.isoformat())
    p = sessie.get(pagina_url, headers=headers, timeout=60)
    p.raise_for_status()

    match = re.search(r"/en/insights/export/xlsx/[\w%:.\-]+", p.text)
    if match is None:
        return _fetch_web_tabel(sessie, headers, van, tot)

    r = sessie.get("https://www.elexys.be" + match.group(0), headers=headers, timeout=60)
    r.raise_for_status()

    ruw = pd.read_excel(io.BytesIO(r.content), header=2)
    ruw = _combineer_datum_tijd(ruw)
    if "_datetime" not in ruw.columns:
        raise ValueError("Onverwachte kolommen in de Elexys-export.")
    out = normalise(ruw, dt_col="_datetime")
    if out.empty:
        raise ValueError("De export bevatte geen bruikbare data.")

    masker = ((out["datetime"] >= pd.Timestamp(van)) &
              (out["datetime"] < pd.Timestamp(tot) + pd.Timedelta(days=1)))
    out = out[masker].reset_index(drop=True)
    if out.empty:
        raise ValueError(f"Geen data in de export voor de periode {van} – {tot}.")
    return out


def jaar_stukken(van: date, tot: date):
    """Knipt een bereik in stukken die elk binnen één kalenderjaar vallen."""
    stukken = []
    for jaar in range(van.year, tot.year + 1):
        stuk_van, stuk_tot = max(van, date(jaar, 1, 1)), min(tot, date(jaar, 12, 31))
        if stuk_van <= stuk_tot:
            stukken.append((stuk_van, stuk_tot))
    return stukken


def fetch_web(van: date, tot: date):
    """
    Haalt de kwartierprijzen op voor een bereik dat meerdere kalenderjaren mag
    beslaan (bv. 2025 + 2026). Elexys geeft zo'n bereik meestal in één export
    terug; lukt dat niet, dan wordt er alsnog jaar per jaar opgehaald en
    samengevoegd.

    Geeft (data, problemen) terug: 'problemen' bevat per mislukt stuk een
    leesbare melding, zodat de geslaagde jaren toch bruikbaar blijven.
    """
    if tot < van:
        raise ValueError("'Tot' ligt voor 'van'.")

    # Vóór KWARTIER_START bestaan er geen kwartierprijzen (Belpex noteerde toen
    # per uur); de Elexys-pagina antwoordt op zo'n periode met een serverfout.
    problemen = []
    if van < KWARTIER_START:
        if tot < KWARTIER_START:
            raise ValueError(
                f"Belpex publiceert pas kwartierprijzen vanaf "
                f"{KWARTIER_START:%d/%m/%Y}; voor die datum bestaan er enkel uurprijzen.")
        problemen.append(
            f"vóór {KWARTIER_START:%d/%m/%Y} bestaan er geen kwartierprijzen — "
            f"die periode is overgeslagen.")
        van = KWARTIER_START

    try:
        data = _fetch_web_periode(van, tot)
    except Exception as e:
        frames = []
        for stuk_van, stuk_tot in jaar_stukken(van, tot):
            try:
                frames.append(_fetch_web_periode(stuk_van, stuk_tot))
            except Exception as jaar_fout:
                problemen.append(f"{stuk_van.year}: {jaar_fout}")
        if not frames:
            raise ValueError("; ".join(problemen) or str(e))
        data = pd.concat(frames, ignore_index=True)

    data = (data.dropna().drop_duplicates("datetime", keep="last")
            .sort_values("datetime").reset_index(drop=True))

    # Elexys levert per export maar een beperkt aantal rijen; als het begin van
    # de gevraagde periode ontbreekt, wordt dat stuk apart nagehaald.
    for _ in range(12):
        oudste = data["datetime"].min().date()
        if oudste <= van:
            break
        try:
            extra = _fetch_web_periode(van, oudste)
        except Exception as e:
            problemen.append(f"tot {oudste:%d/%m/%Y} niet volledig opgehaald: {e}")
            break
        data = (pd.concat([data, extra], ignore_index=True)
                .dropna().drop_duplicates("datetime", keep="last")
                .sort_values("datetime").reset_index(drop=True))
        if data["datetime"].min().date() >= oudste:
            break

    return data, problemen


def load_store() -> pd.DataFrame:
    if STORE.exists():
        d = pd.read_csv(STORE, parse_dates=["datetime"])
        return d.dropna().drop_duplicates("datetime").sort_values("datetime").reset_index(drop=True)
    return pd.DataFrame(columns=["datetime", "price"])


def save_store(new: pd.DataFrame) -> pd.DataFrame:
    merged = pd.concat([load_store(), new[["datetime", "price"]]], ignore_index=True)
    merged = (merged.dropna().drop_duplicates("datetime", keep="last")
              .sort_values("datetime").reset_index(drop=True))
    merged.to_csv(STORE, index=False)
    return merged


@st.cache_data(show_spinner=False)
def demo_data() -> pd.DataFrame:
    """Gesimuleerde prijzen, alleen om het dashboard te tonen."""
    idx = pd.date_range("2026-01-01", "2026-08-27 23:45", freq="15min")
    rng = np.random.default_rng(1)
    h = (idx.hour + idx.minute / 60).to_numpy()
    doy = idx.dayofyear.to_numpy()
    p = (75
         + 25 * np.cos((doy - 15) / 365 * 2 * np.pi)
         - 40 * np.exp(-((h - 13) ** 2) / 8)
         + 35 * np.exp(-((h - 19) ** 2) / 4)
         + 20 * np.exp(-((h - 8) ** 2) / 4)
         + rng.normal(0, 22, len(idx))
         - 90 * (rng.random(len(idx)) < 0.01))
    return pd.DataFrame({"datetime": idx, "price": p.round(2)})


def enrich(df: pd.DataFrame, peak_start: int, peak_end: int, weekdays_only: bool):
    d = df.copy()
    dt = d["datetime"]
    d["date"] = dt.dt.date
    d["month_label"] = dt.dt.to_period("M").astype(str)
    d["quarter_label"] = dt.dt.to_period("Q").astype(str)
    d["dayofweek"] = dt.dt.dayofweek
    d["dayname"] = dt.dt.day_name()
    d["hour"] = dt.dt.hour
    d["qh"] = d["hour"] * 4 + dt.dt.minute // 15
    d["time_label"] = dt.dt.strftime("%H:%M")
    d["is_weekend"] = d["dayofweek"] >= 5
    win = (d["hour"] >= peak_start) & (d["hour"] < peak_end)
    d["is_peak"] = win & ~d["is_weekend"] if weekdays_only else win
    d["is_negative"] = d["price"] < 0
    return d


def monthly_table(d: pd.DataFrame, factor: float = 1.0) -> pd.DataFrame:
    g = d.groupby("month_label")
    out = pd.DataFrame({
        "Baseload": g["price"].mean(),
        "Peak": d[d["is_peak"]].groupby("month_label")["price"].mean(),
        "Off-peak": d[~d["is_peak"]].groupby("month_label")["price"].mean(),
        "Mediaan": g["price"].median(),
        "Min": g["price"].min(),
        "Max": g["price"].max(),
        "Std": g["price"].std(),
        "P10": g["price"].quantile(0.10),
        "P90": g["price"].quantile(0.90),
        "Neg. kwartieren": g["is_negative"].sum(),
        "Neg. %": g["is_negative"].mean() * 100,
    })
    out["Peak ratio"] = out["Peak"] / out["Baseload"]
    prijs_kolommen = ["Baseload", "Peak", "Off-peak", "Mediaan", "Min", "Max", "Std", "P10", "P90"]
    out[prijs_kolommen] = out[prijs_kolommen] * factor
    return out.round(4)


def piek_omschrijving(peak_start: int, peak_end: int, weekdays_only: bool) -> str:
    """Leesbare omschrijving, bv. 'weekdagen 07:00 t.e.m. 21:45 (15u)'."""
    dagen = "weekdagen" if weekdays_only else "elke dag"
    return (f"{dagen} {peak_start:02d}:00 t.e.m. {peak_end - 1:02d}:45 "
            f"({peak_end - peak_start}u)")


def style_fig(fig, height=430, title=None):
    fig.update_layout(
        height=height, title=title, margin=dict(l=10, r=10, t=50, b=10),
        hovermode="x unified", plot_bgcolor="white", separators=",.",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0),
    )
    fig.update_xaxes(showgrid=True, gridcolor="#eee")
    fig.update_yaxes(showgrid=True, gridcolor="#eee")
    return fig


# --------------------------------------------------------------------------
# PERSISTENTE STATUS
# --------------------------------------------------------------------------
# Elke instelling wordt maar op één stap getoond, niet in een permanente
# zijbalk. Om die waarden ook op de andere stappen te kunnen gebruiken
# (bv. de piekuren die stap 1 verzamelt, gebruikt door stap 2 en 3), worden
# ze hier bij wijziging in st.session_state weggeschreven en bovenaan elke
# rerun met een veilige standaardwaarde teruggelezen — ongeacht welke stap
# net actief is.

cache = st.session_state.setdefault("_cache", {})
raw = cache.get("raw", pd.DataFrame(columns=["datetime", "price"]))
source = cache.get("source", "")
d0 = cache.get("d0")
d1 = cache.get("d1")
piekprofiel = cache.get("piekprofiel", list(PIEKPROFIELEN)[0])
peak_start, peak_end = cache.get("peak_uren", PIEKPROFIELEN[list(PIEKPROFIELEN)[0]])
weekdays_only = cache.get("weekdays_only", True)
eenheid = cache.get("eenheid", "EUR/kWh")
# De vergoeding-widgets staan pas op stap 2, ná de Vorige/Volgende-knoppen.
# Lees hun waarde daarom rechtstreeks uit session_state (die is bij een klik al
# bijgewerkt), anders loopt de blokkade één rerun achter op het vinkje.
vergoeding_in = st.session_state.get("vergoeding_input", cache.get("vergoeding", 0.0))
geen_vergoeding = st.session_state.get("w_geen_vergoeding", cache.get("geen_vergoeding", False))
vbr_df = cache.get("vbr_df")            # al genormaliseerde verbruik-DataFrame, of None
vbr_naam = cache.get("vbr_naam")
vbr_melding = cache.get("vbr_melding")  # (niveau, tekst) van de laatste upload-poging

factor = EENHEDEN[eenheid]
dec = DECIMALEN[eenheid]

st.session_state.setdefault("stap", 1)
st.session_state["stap"] = max(1, min(len(STAPPEN), st.session_state["stap"]))
stap = st.session_state["stap"]

st.title("Scholt — Belpex kwartierprijzen")
stappenbalk(stap)

# --------------------------------------------------------------------------
# STAP 1 — DATA & PERIODE
# --------------------------------------------------------------------------
if stap == 1:
    # Meestal staat de databank al klaar: dan gaat het op deze stap enkel om de
    # periode en de piekuren. Prijzen ophalen of uploaden is onderhoud en zit
    # ingeklapt; bij een lege databank staat het open.
    store = load_store()
    if store.empty:
        st.write("De prijzendatabank is nog leeg. Haal eerst de marktprijzen op of upload "
                 "een prijsbestand.")
    else:
        st.write(f"De databank bevat Belpex-kwartierprijzen van "
                 f"{store['datetime'].min():%d/%m/%Y} t/m {store['datetime'].max():%d/%m/%Y}. "
                 "Kies de periode en de piekuren voor deze klant.")

    with st.expander("Prijzen bijwerken of zelf een prijsbestand gebruiken",
                     icon=":material/database:", expanded=store.empty):
        kol1, kol2 = st.columns(2)
        with kol1:
            st.markdown("**Optie A — marktprijzen automatisch ophalen**")
            vandaag = date.today()
            # Standaard het lopende jaar: oudere kwartieren zitten al in de databank
            # en opnieuw ophalen kost alleen tijd. Verder terug mag nog altijd —
            # tot KWARTIER_START, want eerder bestaan er geen kwartierprijzen.
            jaarbegin = max(KWARTIER_START, date(vandaag.year, 1, 1))
            fv = st.date_input("Van", value=jaarbegin, min_value=KWARTIER_START,
                               max_value=vandaag, key="fetch_van")
            ft = st.date_input("Tot", value=vandaag, min_value=KWARTIER_START,
                               max_value=vandaag, key="fetch_tot")
            st.caption(f"Standaard wordt {vandaag.year} opgehaald en samengevoegd met wat er "
                       f"al in de databank zit. Zet 'Van' vroeger om verder terug te gaan — "
                       f"tot {KWARTIER_START:%d/%m/%Y}, want eerder bestaan er geen "
                       f"kwartierprijzen, enkel uurprijzen.")
            if st.button("Marktprijzen ophalen", icon=":material/download:", width="stretch"):
                try:
                    jaren = [jv.year for jv, _ in jaar_stukken(fv, ft)]
                    with st.spinner(f"Ophalen {', '.join(str(j) for j in jaren)} ..."):
                        gehaald, problemen = fetch_web(fv, ft)
                    save_store(gehaald)
                    per_jaar = gehaald["datetime"].dt.year.value_counts().sort_index()
                    detail = " · ".join(f"{j}: {nl(n)}" for j, n in per_jaar.items())
                    st.success(f"{nl(len(gehaald))} kwartieren opgehaald en bewaard ({detail}).")
                    for pr in problemen:
                        st.warning(f"Niet opgehaald — {pr}")
                    st.rerun()
                except Exception as e:
                    st.warning(f"Ophalen lukte niet: {e}")
                    st.caption("Download de prijzen dan handmatig en upload ze rechts.")
        with kol2:
            st.markdown("**Optie B — zelf een prijsbestand uploaden**")
            ups = st.file_uploader("Prijsbestand(en) — CSV of Excel",
                                   type=["csv", "xlsx", "xls", "txt"],
                                   accept_multiple_files=True, key="prijs_upload")
            kp1, kp2 = st.columns(2)
            dt_pick = kp1.text_input("Datum/tijd-kolom", "", placeholder="automatisch",
                                     help="Enkel invullen als de kolom niet herkend wordt.")
            pr_pick = kp2.text_input("Prijskolom", "", placeholder="automatisch",
                                     help="Enkel invullen als de kolom niet herkend wordt.")

            parsed = pd.DataFrame(columns=["datetime", "price"])
            if ups:
                frames, problems = [], []
                for f in ups:
                    try:
                        frames.append(normalise(read_upload(f.getvalue(), f.name),
                                                dt_pick or None, pr_pick or None))
                    except Exception as e:
                        problems.append(f"{f.name}: {e}")
                for msg in problems:
                    st.error(msg)
                if frames:
                    parsed = pd.concat(frames, ignore_index=True).drop_duplicates("datetime")
                    st.success(f"{nl(len(parsed))} kwartieren ingelezen uit {len(frames)} bestand(en)")
                    if st.button("Toevoegen aan prijzendatabank", icon=":material/add:",
                                 width="stretch"):
                        save_store(parsed)
                        st.success("Toegevoegd.")
                        st.rerun()

        store = load_store()
        st.divider()
        bron = st.radio(
            "Welke data wil je gebruiken?",
            ["Databank", "Enkel de upload hierboven", "Demodata (voorbeeld)"],
            index=0 if not store.empty else (1 if ups else 2), horizontal=True)

        if bron.startswith("Databank"):
            if store.empty:
                st.info("De databank is nog leeg. Gebruik optie A of B hierboven.")
                raw, source = pd.DataFrame(columns=["datetime", "price"]), ""
            else:
                raw, source = store, f"databank ({STORE.name})"
        elif bron.startswith("Enkel"):
            if parsed.empty:
                st.info("Upload eerst een bestand hierboven (optie B).")
                raw, source = pd.DataFrame(columns=["datetime", "price"]), ""
            else:
                raw, source = parsed, ", ".join(f.name for f in ups)
        else:
            raw, source = demo_data(), "demodata (gesimuleerd)"

        if not store.empty:
            st.caption(f"Databank: {nl(len(store))} kwartieren · "
                      f"{store['datetime'].min():%d/%m/%Y} → {store['datetime'].max():%d/%m/%Y}")
            db1, db2 = st.columns(2)
            db1.download_button("Databank downloaden (CSV)", store.to_csv(index=False).encode(),
                                "belpex_databank.csv", "text/csv",
                                icon=":material/download:", width="stretch")
            # Wissen is niet ongedaan te maken: eerst bewust aanvinken.
            zeker = db2.checkbox("Ik wil alle prijzen in de databank wissen", key="wis_zeker")
            if db2.button("Databank wissen", icon=":material/delete:", width="stretch",
                          disabled=not zeker):
                STORE.unlink(missing_ok=True)
                st.session_state["wis_zeker"] = False
                st.rerun()

    if not raw.empty:
        st.divider()
        st.markdown("**Periode**")
        lo, hi = raw["datetime"].min().date(), raw["datetime"].max().date()
        jaren = sorted(raw["datetime"].dt.year.unique())
        keuzes = ["Volledig", "Laatste 30 dagen", "Laatste 90 dagen",
                  "Dit jaar", "Per jaar", "Zelf kiezen"]
        # Bij data uit meerdere jaren is "volledig" de zinnigste standaard;
        # bij één jaar blijft dat year-to-date.
        preset = st.radio("Snelkeuze", keuzes,
                          index=0 if len(jaren) > 1 else 3, horizontal=True)
        if preset == "Laatste 30 dagen":
            d0, d1 = max(lo, hi - pd.Timedelta(days=29).to_pytimedelta()), hi
        elif preset == "Laatste 90 dagen":
            d0, d1 = max(lo, hi - pd.Timedelta(days=89).to_pytimedelta()), hi
        elif preset == "Dit jaar":
            d0, d1 = max(lo, date(hi.year, 1, 1)), hi
        elif preset == "Per jaar":
            jaar = st.selectbox("Jaar", jaren, index=len(jaren) - 1)
            d0, d1 = max(lo, date(jaar, 1, 1)), min(hi, date(jaar, 12, 31))
        elif preset == "Zelf kiezen":
            picked = st.date_input("Van / tot", value=(lo, hi), min_value=lo, max_value=hi)
            d0, d1 = picked if isinstance(picked, tuple) and len(picked) == 2 else (lo, hi)
        else:
            d0, d1 = lo, hi
        st.caption(f"{d0:%d/%m/%Y} → {d1:%d/%m/%Y}")

        st.divider()
        st.markdown("**Piek / dal**")
        # Deze keuze staat bewust op stap 1, vóór elke berekening: ze bepaalt de
        # kerncijfers, de eenheidsprijs piek/dal én de Excel-werkmap, dus ze moet
        # vastliggen voor er ook maar iets gerekend wordt.
        st.session_state.setdefault("w_piekprofiel", piekprofiel)
        st.session_state.setdefault("w_peak_uren", (peak_start, peak_end))
        st.session_state.setdefault("w_weekdays_only", weekdays_only)
        st.session_state.setdefault("w_eenheid", eenheid)

        pk1, pk2 = st.columns([3, 2])
        with pk1:
            piekprofiel = st.radio("Verhouding", list(PIEKPROFIELEN), key="w_piekprofiel")
            vaste_uren = PIEKPROFIELEN[piekprofiel]
            if vaste_uren is None:
                peak_start, peak_end = st.slider("Piekuren", 0, 24, key="w_peak_uren")
            else:
                peak_start, peak_end = vaste_uren
            weekdays_only = st.checkbox("Piek enkel op werkdagen (ma-vr)",
                                        key="w_weekdays_only")
        with pk2:
            eenheid = st.radio("Eenheid (prijs)", list(EENHEDEN), key="w_eenheid")
        dagbalk(peak_start, peak_end, weekdays_only)

        cache["piekprofiel"] = piekprofiel
        cache["peak_uren"] = (peak_start, peak_end)
        cache["weekdays_only"] = weekdays_only
        cache["eenheid"] = eenheid
        factor = EENHEDEN[eenheid]
        dec = DECIMALEN[eenheid]

    cache["raw"], cache["source"] = raw, source
    cache["d0"], cache["d1"] = d0, d1

# --------------------------------------------------------------------------
# FILTEREN + VERRIJKEN (elke stap, ongeacht welke actief is)
# --------------------------------------------------------------------------

data_geladen = (not raw.empty) and d0 is not None and d1 is not None
if data_geladen:
    mask = (raw["datetime"] >= pd.Timestamp(d0)) & (raw["datetime"] < pd.Timestamp(d1) + pd.Timedelta(days=1))
    d = enrich(raw[mask], peak_start, peak_end, weekdays_only)
    d["price"] = d["price"] / 1000        # EUR/MWh (bron) -> EUR/kWh (canoniek)
    d["price_disp"] = d["price"] * factor  # canoniek -> gekozen weergave-eenheid
    data_geladen = not d.empty
else:
    d = pd.DataFrame()

# --- Verbruik + vergoeding koppelen (voedt stap 3 en 4) --------------------
heeft_verbruik = vbr_df is not None
vergoeding_bevestigd = geen_vergoeding or vergoeding_in > 0
d_wb = pd.DataFrame()
ep_totaal = ep_piek = ep_dal = kost_piek = kost_dal = None
piek_vast = None
if data_geladen:
    d_wb = d.copy()
    d_wb["vergoeding"] = 0.0 if geen_vergoeding else vergoeding_in
    if heeft_verbruik:
        d_wb = d_wb.merge(vbr_df, on="datetime", how="left")
        d_wb["verbruik"] = d_wb["verbruik"].fillna(0.0)

    piek_vast = (d_wb["hour"] >= peak_start) & (d_wb["hour"] < peak_end)
    if weekdays_only:
        piek_vast &= d_wb["datetime"].dt.dayofweek < 5

    if heeft_verbruik:
        vb = d_wb.dropna(subset=["verbruik"])
        vbr_piek_mask = piek_vast.loc[vb.index]

        def _eenheidsprijs(sel: pd.DataFrame):
            v = sel["verbruik"].sum()
            return ((sel["price"] + sel["vergoeding"]) * sel["verbruik"]).sum() / v if v else None

        ep_totaal = _eenheidsprijs(vb)
        ep_piek = _eenheidsprijs(vb[vbr_piek_mask])
        ep_dal = _eenheidsprijs(vb[~vbr_piek_mask])
        kost_piek = ((vb.loc[vbr_piek_mask, "price"] + vb.loc[vbr_piek_mask, "vergoeding"])
                     * vb.loc[vbr_piek_mask, "verbruik"]).sum()
        kost_dal = ((vb.loc[~vbr_piek_mask, "price"] + vb.loc[~vbr_piek_mask, "vergoeding"])
                    * vb.loc[~vbr_piek_mask, "verbruik"]).sum()

# --------------------------------------------------------------------------
# WIZARD — voortgang + Vorige/Volgende met blokkades
# --------------------------------------------------------------------------

blokkade = None
if stap == 1 and not data_geladen:
    blokkade = "Laad eerst prijsdata en kies een periode voor je verder kan."
elif stap == 2 and not heeft_verbruik:
    blokkade = "Upload eerst het verbruik (bovenaan deze stap) voor je verder kan."
elif stap == 2 and not vergoeding_bevestigd:
    blokkade = "Vul de vergoeding in, of vink 'geen aparte leveringsvergoeding' aan."


def navigatie() -> None:
    """Vorige/Volgende, altijd onderaan de stap: eerst invullen, dan verder."""
    st.divider()
    if blokkade:
        st.warning(blokkade, icon=":material/info:")
    links, _, rechts = st.columns([1, 3, 1])
    with links:
        if st.button("Vorige", icon=":material/arrow_back:", disabled=(stap == 1),
                     width="stretch"):
            st.session_state["stap"] -= 1
            st.rerun()
    with rechts:
        if stap < len(STAPPEN) and st.button(
                "Volgende", icon=":material/arrow_forward:",
                icon_position="right", type="primary", disabled=blokkade is not None,
                width="stretch"):
            st.session_state["stap"] += 1
            st.rerun()


if not data_geladen:
    if stap != 1:
        st.info("Ga naar stap 1 en laad eerst prijsdata.")
    navigatie()
    st.stop()

if stap > 1:
    st.caption(f"Prijzen: {source} · {d['datetime'].min():%d/%m/%Y} t/m "
               f"{d['datetime'].max():%d/%m/%Y} · {nl(len(d))} kwartieren · "
               f"piek {piek_omschrijving(peak_start, peak_end, weekdays_only)}")

# ============================================================================
# ============================================================================
# STAP 2 — VERBRUIK & VERGOEDING
# ============================================================================
if stap == 2:
    st.write("Upload het kwartierverbruik van de klant en vul de leveringsvergoeding in. "
             "De eenheidsprijs verschijnt meteen hieronder.")

    v1, v2 = st.columns(2, gap="large")
    with v1:
        nieuw_bestand = st.file_uploader(
            "Verbruik per kwartier (CSV of Excel)",
            type=["csv", "xlsx", "xls", "txt"], key="verbruik_upload",
            help="Eén lijn per kwartier, met een datum/tijd en een verbruik. "
                 "Kolomnamen en notatie maken niet uit; de tool zoekt ze zelf.")
        # Onthoud het laatst geprobeerde bestand (ook bij een fout), anders wordt
        # een onleesbaar bestand bij elke rerun opnieuw geprobeerd.
        if nieuw_bestand is not None and nieuw_bestand.name != cache.get("vbr_poging"):
            cache["vbr_poging"] = nieuw_bestand.name
            # De ruwe bytes blijven bewaard voor de manuele kolomkeuze hieronder.
            cache["vbr_bytes"] = nieuw_bestand.getvalue()
            try:
                ruw = _combineer_datum_tijd(read_upload(cache["vbr_bytes"], nieuw_bestand.name))
                vbr = normalise_verbruik(ruw, dt_col="_datetime" if "_datetime" in ruw.columns else None)
                if vbr.empty:
                    raise ValueError("geen enkele rij met een geldige datum/tijd en verbruik.")
                cache["vbr_df"] = vbr
                cache["vbr_naam"] = nieuw_bestand.name
                cache["vbr_melding"] = verbruik_melding(vbr, d)
                cache["vbr_manueel_open"] = False
                # Wat automatisch herkend werd, als startwaarde voor de manuele keuze.
                tijd_kol, waarde_kol = vbr.attrs["bron"]
                uur_kol = None
                if tijd_kol == "_datetime":
                    tijd_kol, uur_kol = ruw.attrs["datum_tijd"]
                eenheid_col = next((c for c in ruw.columns
                                    if str(c).lower() in ("unit", "eenheid")), None)
                eenheid = "kWh"
                if eenheid_col is not None:
                    meest = ruw[eenheid_col].astype(str).str.strip().str.lower().mode()
                    eenheid = {"wh": "Wh", "mwh": "MWh"}.get(meest.iloc[0] if len(meest) else "", "kWh")
                cache["vbr_auto"] = {"tijd": tijd_kol, "uur": uur_kol,
                                     "waarde": waarde_kol, "eenheid": eenheid}
            except Exception as e:
                cache["vbr_melding"] = ("error", f"Kon **{nieuw_bestand.name}** niet automatisch "
                                        f"inlezen: {e} Kies hieronder zelf de kolommen.")
                cache["vbr_manueel_open"] = True
                cache["vbr_auto"] = {}
            st.rerun()
        if vbr_naam:
            st.caption(f"Actief bestand: **{vbr_naam}**")
        if vbr_melding:
            niveau, tekst = vbr_melding
            getattr(st, niveau)(tekst)
        if vbr_df is not None:
            standaard = vbr_df.rename(columns={"verbruik": "verbruik_kwh"})
            st.download_button(
                "Gestandaardiseerd verbruik (CSV)",
                standaard.to_csv(index=False, sep=";", decimal=",",
                                 date_format="%Y-%m-%d %H:%M").encode("utf-8-sig"),
                file_name="verbruik_standaard.csv", mime="text/csv",
                icon=":material/download:")
    with v2:
        st.session_state.setdefault("vergoeding_input", vergoeding_in)
        st.session_state.setdefault("w_geen_vergoeding", geen_vergoeding)
        vergoeding_in = st.number_input(
            "Vergoeding (EUR/kWh) — marge leverancier, bovenop de Belpex-prijs",
            min_value=0.0, step=0.001, format="%.4f", key="vergoeding_input")
        geen_vergoeding = st.checkbox("Ik betaal geen aparte leveringsvergoeding (vergoeding = 0)",
                                      key="w_geen_vergoeding")
        cache["vergoeding"] = vergoeding_in
        cache["geen_vergoeding"] = geen_vergoeding
        if not (geen_vergoeding or vergoeding_in > 0):
            st.caption("Vul een bedrag in, of vink het vakje aan.")

    # --- Manuele kolomkeuze: terugval als de automatische herkenning faalt of
    # de verkeerde kolom pakt. Werkt op de bewaarde bytes van de laatste upload.
    vbr_bytes = cache.get("vbr_bytes")
    poging = cache.get("vbr_poging")
    if vbr_bytes is not None:
        with st.expander("Kolommen controleren of zelf kiezen", icon=":material/tune:",
                         expanded=cache.get("vbr_manueel_open", False)):
            st.caption("Vooraf ingevuld met wat automatisch herkend werd. Pas aan wat niet "
                       "klopt en klik op **Toepassen**.")
            m1, m2 = st.columns(2)
            overslaan = m1.number_input("Lijnen overslaan bovenaan (tekst boven de tabel)",
                                        0, 100, 0, key=f"m_skip_{poging}")
            kopregel = m2.checkbox("Eerste lijn bevat kolomnamen", True, key=f"m_kop_{poging}")
            try:
                tabel = read_upload(vbr_bytes, poging, int(overslaan), kopregel)
                tabel.columns = [str(c).strip() for c in tabel.columns]
            except Exception as e:
                st.error(f"Kon de tabel niet lezen met deze instellingen: {e}")
                tabel = None

            if tabel is not None:
                st.caption("Eerste lijnen zoals ingelezen:")
                st.dataframe(tabel.head(5), hide_index=True)
                kols = list(tabel.columns)
                geen = "— geen (datum en uur zitten samen) —"
                auto = cache.get("vbr_auto") or {}

                def _index(opties, waarde, standaard):
                    return opties.index(waarde) if waarde in opties else standaard

                eenheden = ["kWh", "Wh", "MWh"]
                c1, c2, c3, c4 = st.columns(4)
                tijd_kol = c1.selectbox("Datum/tijd (begin kwartier)", kols,
                                        index=_index(kols, auto.get("tijd"), 0),
                                        key=f"m_tijd_{poging}")
                uur_kol = c2.selectbox("Uur apart", [geen] + kols,
                                       index=_index([geen] + kols, auto.get("uur"), 0),
                                       key=f"m_uur_{poging}")
                waarde_kol = c3.selectbox("Verbruik", kols,
                                          index=_index(kols, auto.get("waarde"), len(kols) - 1),
                                          key=f"m_waarde_{poging}")
                eenheid_keuze = c4.selectbox("Eenheid", eenheden,
                                             index=_index(eenheden, auto.get("eenheid"), 0),
                                             key=f"m_eenh_{poging}")

                if st.button("Toepassen", type="primary"):
                    try:
                        bron = tabel.copy()
                        if uur_kol != geen:
                            uur = (bron[uur_kol].astype(str).str.replace("u", ":", regex=False)
                                   .str.replace("h", ":", regex=False).str.strip())
                            bron["_datetime"] = bron[tijd_kol].astype(str).str.strip() + " " + uur
                            dt_kol = "_datetime"
                        else:
                            dt_kol = tijd_kol
                        vbr = normalise_verbruik(bron, dt_col=dt_kol, verbruik_col=waarde_kol,
                                                 eenheid=eenheid_keuze.lower())
                        if uur_kol != geen:
                            vbr.attrs["bron"] = (f"{tijd_kol} + {uur_kol}", waarde_kol)
                        if vbr.empty:
                            raise ValueError("met deze kolommen bleef er geen enkele geldige rij "
                                             "over (datum of getal niet leesbaar).")
                        cache["vbr_df"] = vbr
                        cache["vbr_naam"] = poging
                        cache["vbr_melding"] = verbruik_melding(vbr, d)
                        cache["vbr_manueel_open"] = False
                    except Exception as e:
                        cache["vbr_melding"] = ("error", f"Manuele keuze lukte niet: {e}")
                        cache["vbr_manueel_open"] = True
                    st.rerun()

    st.divider()
    if not heeft_verbruik:
        st.info("Upload het verbruik hierboven om de eenheidsprijs te zien.")
    elif ep_totaal is None:
        st.warning("Er is geen verbruik in de gekozen periode, dus geen eenheidsprijs. "
                   "Kijk de periode op stap 1 na.")
    else:
        def _prijs(v):
            return nl(v * factor, dec) if v is not None else "—"

        verbruik_tot = float(vb["verbruik"].sum())
        verbruik_piek = float(vb.loc[vbr_piek_mask, "verbruik"].sum())
        kost_tot = kost_piek + kost_dal
        # De uitkomst eerst en groot: dit is het getal waarvoor de tool bestaat.
        st.markdown(f"""
<div class="uitkomst">
  <div class="uitkomst-label">Eenheidsprijs {d0:%d/%m/%Y} – {d1:%d/%m/%Y}</div>
  <div class="uitkomst-waarde">{_prijs(ep_totaal)}<small>{eenheid}</small></div>
  <div class="uitkomst-rij">
    <span><i class="p"></i>Piek <b>{_prijs(ep_piek)}</b></span>
    <span><i></i>Dal <b>{_prijs(ep_dal)}</b></span>
    <span>Verbruik <b>{nl(verbruik_tot)} kWh</b>
          ({nl(verbruik_piek / verbruik_tot * 100 if verbruik_tot else 0)}% in piek)</span>
    <span>Kost <b>€ {nl(kost_tot)}</b></span>
  </div>
</div>""", unsafe_allow_html=True)
        st.caption(f"Eenheidsprijs = Σ (Belpex-prijs + vergoeding) × verbruik ÷ Σ verbruik, "
                   f"over {nl(len(d_wb))} kwartieren ({nl(int(piek_vast.sum()))} piek, "
                   f"{nl(int((~piek_vast).sum()))} dal).")

        g1, g2 = st.columns(2, gap="large")
        with g1:
            fig = go.Figure(go.Bar(
                x=["Totaal", "Piek", "Dal"],
                y=[(v * factor if v is not None else 0) for v in (ep_totaal, ep_piek, ep_dal)],
                marker_color=[NAVY, ROOD, ZAND], marker_line_color=[NAVY, ROOD, ZAND_RAND],
                marker_line_width=1))
            fig.update_yaxes(title=eenheid)
            st.plotly_chart(style_fig(fig, 340, "Eenheidsprijs per periode"), **FULL)
        with g2:
            fig = go.Figure(go.Pie(labels=["Piek", "Dal"], values=[kost_piek, kost_dal],
                                   hole=0.55, marker_colors=[ROOD, ZAND], sort=False,
                                   textinfo="label+percent", showlegend=False,
                                   marker_line_color="white", marker_line_width=2))
            st.plotly_chart(style_fig(fig, 340, "Aandeel in de kost (EUR)"), **FULL)

# ============================================================================
# STAP 3 — EXPORT
# ============================================================================
elif stap == 3:
    st.subheader("Eenheidsprijs-werkmap")
    st.caption("Excel met levende formules — prijs, vergoeding en verbruik staan al ingevuld "
              "op het blad 'Kwartierdata'; 'Overzicht' herrekent zichzelf.")
    fname = f"belpex_eenheidsprijs_{d0:%Y%m%d}_{d1:%Y%m%d}.xlsx"
    if st.button("Scholt-werkmap aanmaken", icon=":material/table_view:", type="primary",
                 width="stretch"):
        with st.spinner("Werkmap bouwen ..."):
            st.session_state["wb"] = build_workbook(
                d_wb, peak_start, peak_end, weekdays_only)
            st.session_state["wb_name"] = fname

    if st.session_state.get("wb"):
        st.download_button("Werkmap opslaan", st.session_state["wb"],
                           st.session_state["wb_name"],
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           icon=":material/download:", type="primary", width="stretch")

    st.divider()
    st.subheader("Marktprijzen in de periode")
    st.caption(f"Belpex zelf, los van het verbruik van de klant. Prijzen in {eenheid}.")
    p = d["price"]
    peak_avg = d.loc[d["is_peak"], "price"].mean()
    off_avg = d.loc[~d["is_peak"], "price"].mean()
    daily = d.groupby("date")["price"]
    c = st.columns(6)
    c[0].metric("Gemiddeld", nl(p.mean() * factor, dec),
                help="Baseload: gewoon gemiddelde van alle kwartieren.")
    c[1].metric("Piek", nl(peak_avg * factor, dec) if pd.notna(peak_avg) else "—",
                f"{'+' if peak_avg >= p.mean() else ''}{nl((peak_avg / p.mean() - 1) * 100, 1)}%"
                if pd.notna(peak_avg) and p.mean() else None, delta_color="off")
    c[2].metric("Dal", nl(off_avg * factor, dec) if pd.notna(off_avg) else "—")
    c[3].metric("Schommeling (σ)", nl(p.std() * factor, dec),
                help="Standaardafwijking van de kwartierprijzen.")
    c[4].metric("Negatieve prijs", f"{nl(d['is_negative'].mean() * 100, 1)}%",
                f"{nl(int(d['is_negative'].sum()))} kwartieren", delta_color="off")
    c[5].metric("Gem. dagverschil", nl((daily.max() - daily.min()).mean() * factor, dec),
                help="Gemiddeld verschil tussen de duurste en goedkoopste kwartier van een dag.")

    st.caption(f"Downloads: prijskolommen in {eenheid}; 'Ruwe data' en de CSV blijven in "
               "EUR/kWh.")
    mt = monthly_table(d, factor)
    profiles = {
        "Kwartierprofiel": d.groupby("time_label")["price_disp"].agg(["mean", "median", "min", "max", "std"]).round(4),
        "Uurprofiel": d.groupby("hour")["price_disp"].agg(["mean", "median", "min", "max"]).round(4),
        "Dagreeks": d.groupby("date").agg(baseload=("price_disp", "mean"), min=("price_disp", "min"),
                                          max=("price_disp", "max"),
                                          spread=("price_disp", lambda s: s.max() - s.min()),
                                          neg=("is_negative", "sum")).round(4),
    }
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        mt.to_excel(w, sheet_name="Maandoverzicht")
        for n, f in profiles.items():
            f.to_excel(w, sheet_name=n[:31])
        d[["datetime", "price", "is_peak", "is_negative"]].to_excel(w, sheet_name="Ruwe data", index=False)

    x, y = st.columns(2)
    x.download_button("Analyse als Excel", buf.getvalue(),
                      f"belpex_{d0:%Y%m%d}_{d1:%Y%m%d}.xlsx",
                      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                      icon=":material/download:", **FULL)
    y.download_button("Kwartierprijzen als CSV",
                      d[["datetime", "price", "is_peak", "is_negative"]]
                      .to_csv(sep=";", decimal=",", index=False).encode(),
                      f"belpex_{d0:%Y%m%d}_{d1:%Y%m%d}.csv", "text/csv",
                      icon=":material/download:", **FULL)

    st.caption("Laatste 200 kwartieren van de periode:")
    st.dataframe(d[["datetime", "price", "is_peak", "is_negative"]].tail(200)
                 .rename(columns={"datetime": "Tijdstip", "price": "Prijs (EUR/kWh)",
                                  "is_peak": "Piek", "is_negative": "Negatief"}),
                 hide_index=True, **FULL, height=400)

navigatie()
