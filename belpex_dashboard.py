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

STAPPEN = ["📋 Data & periode", "🧾 Verbruik & vergoeding", "💾 Export"]

# Piek/dal-verhoudingen. De waarden zijn (eerste piekuur, eerste daluur na de
# piek): (8, 20) betekent piek van 08:00 t.e.m. 19:45. De gekozen verhouding
# bepaalt alles — de kerncijfers, de eenheidsprijs piek/dal én de werkmap.
PIEKPROFIELEN = {
    "Standaard — 08:00 t.e.m. 19:45 (12u)": (8, 20),
    "Scholt — 07:00 t.e.m. 21:45 (15u)": (7, 22),
    "Zelf kiezen": None,
}

EENHEDEN = {"EUR/kWh": 1, "cent/kWh": 100, "EUR/MWh": 1000}
DECIMALEN = {"EUR/kWh": 4, "cent/kWh": 2, "EUR/MWh": 2}

# --------------------------------------------------------------------------
# PAGINA-INSTELLINGEN
# --------------------------------------------------------------------------

st.set_page_config(page_title="Scholt — kwartierprijzen", page_icon="⚡", layout="wide")

# Huisstijl Scholt Energy (scholt.be): rood accent op donkerblauw.
ROOD, NAVY, DONKER, ZAND, GREY = "#e61d3a", "#030737", "#171b25", "#edeae4", "#7f8c8d"

# Compatibiliteit: Streamlit >= 1.60 gebruikt width="stretch" i.p.v. use_container_width
try:
    _v = tuple(int(x) for x in st.__version__.split(".")[:2])
    FULL = {"width": "stretch"} if _v >= (1, 60) else {"use_container_width": True}
except Exception:
    FULL = {"use_container_width": True}

st.markdown("""
<style>
  .block-container {padding-top: 2rem; max-width: 1100px;}
  [data-testid="stMetricValue"] {font-size: 1.6rem;}
  h1 {font-size: 1.9rem !important;}
</style>
""", unsafe_allow_html=True)


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


def normalise(df: pd.DataFrame, dt_col=None, value_col=None, value_name="price",
              value_keywords=("price", "prijs", "eur", "€", "value", "waarde")) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    if dt_col is None:
        for c in df.columns:
            if any(k in c.lower() for k in ("date", "datum", "time", "tijd", "period", "mtu")):
                dt_col = c
                break
        dt_col = dt_col or df.columns[0]
    if value_col is None:
        for c in df.columns:
            if c != dt_col and any(k in c.lower() for k in value_keywords):
                value_col = c
                break
        value_col = value_col or [c for c in df.columns if c != dt_col][-1]

    ts = pd.to_datetime(df[dt_col], errors="coerce", dayfirst=True)
    if ts.isna().mean() > 0.5:
        ts = pd.to_datetime(df[dt_col], errors="coerce", dayfirst=False)

    out = pd.DataFrame({"datetime": ts, value_name: to_numeric(df[value_col])})
    out = out.dropna().drop_duplicates("datetime").sort_values("datetime")
    return out.reset_index(drop=True)


def normalise_verbruik(df: pd.DataFrame, dt_col=None, verbruik_col=None) -> pd.DataFrame:
    """Zelfde als normalise(), maar zoekt een verbruikskolom i.p.v. een prijskolom."""
    return normalise(df, dt_col, verbruik_col, value_name="verbruik",
                     value_keywords=("verbruik", "consumption", "kwh", "energie",
                                     "energy", "volume", "hoeveelheid", "afname"))


@st.cache_data(show_spinner=False)
def read_upload(content: bytes, name: str) -> pd.DataFrame:
    buf = io.BytesIO(content)
    if name.lower().endswith((".xlsx", ".xls", ".xlsm")):
        return pd.read_excel(buf)
    for sep in (";", ",", "\t", "|"):
        buf.seek(0)
        try:
            trial = pd.read_csv(buf, sep=sep, engine="python")
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
        hovermode="x unified", plot_bgcolor="white",
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
vergoeding_in = cache.get("vergoeding", 0.0)
geen_vergoeding = cache.get("geen_vergoeding", False)
vbr_df = cache.get("vbr_df")            # al genormaliseerde verbruik-DataFrame, of None
vbr_naam = cache.get("vbr_naam")
vbr_melding = cache.get("vbr_melding")  # (niveau, tekst) van de laatste upload-poging

factor = EENHEDEN[eenheid]
dec = DECIMALEN[eenheid]

st.session_state.setdefault("stap", 1)
st.session_state["stap"] = max(1, min(len(STAPPEN), st.session_state["stap"]))
stap = st.session_state["stap"]

st.title("⚡ Scholt — Belpex kwartierprijzen")

# --------------------------------------------------------------------------
# STAP 1 — DATA & PERIODE
# --------------------------------------------------------------------------
if stap == 1:
    st.write("Laad eerst je prijsdata. Kies één van de twee manieren hieronder.")

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
        if st.button("⚡ Marktprijzen ophalen", type="primary", width="stretch"):
            try:
                jaren = [jv.year for jv, _ in jaar_stukken(fv, ft)]
                with st.spinner(f"Ophalen {', '.join(str(j) for j in jaren)} ..."):
                    gehaald, problemen = fetch_web(fv, ft)
                save_store(gehaald)
                per_jaar = gehaald["datetime"].dt.year.value_counts().sort_index()
                detail = " · ".join(f"{j}: {n:,}" for j, n in per_jaar.items())
                st.success(f"{len(gehaald):,} kwartieren opgehaald en bewaard ({detail}).")
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
        with st.expander("Kolommen (enkel bij detectieproblemen)"):
            dt_pick = st.text_input("Naam datum/tijd-kolom", "", placeholder="automatisch")
            pr_pick = st.text_input("Naam prijskolom", "", placeholder="automatisch")

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
                st.success(f"{len(parsed):,} kwartieren ingelezen uit {len(frames)} bestand(en)")
                if st.button("➕ Toevoegen aan prijzendatabank", width="stretch"):
                    save_store(parsed)
                    st.success("Toegevoegd.")
                    st.rerun()

    store = load_store()
    st.divider()
    bron = st.radio(
        "Welke data wil je gebruiken?",
        ["Databank (opgehaald/geüpload, YTD)", "Enkel de upload hierboven", "Demodata (voorbeeld)"],
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
        st.caption(f"Databank: {len(store):,} kwartieren · "
                  f"{store['datetime'].min():%d/%m/%Y} → {store['datetime'].max():%d/%m/%Y}")
        with st.expander("Databank beheren"):
            st.download_button("Databank downloaden (CSV)",
                               store.to_csv(index=False).encode(),
                               "belpex_databank.csv", "text/csv", width="stretch")
            if st.button("🗑 Databank wissen", width="stretch"):
                STORE.unlink(missing_ok=True)
                st.rerun()

    if not raw.empty:
        st.divider()
        st.markdown("**Periode**")
        lo, hi = raw["datetime"].min().date(), raw["datetime"].max().date()
        jaren = sorted(raw["datetime"].dt.year.unique())
        keuzes = ["Volledig", "Laatste 30 dagen", "Laatste 90 dagen",
                  "Year to date", "Per jaar", "Zelf kiezen"]
        # Bij data uit meerdere jaren is "volledig" de zinnigste standaard;
        # bij één jaar blijft dat year-to-date.
        preset = st.radio("Snelkeuze", keuzes,
                          index=0 if len(jaren) > 1 else 3, horizontal=True)
        if preset == "Laatste 30 dagen":
            d0, d1 = max(lo, hi - pd.Timedelta(days=29).to_pytimedelta()), hi
        elif preset == "Laatste 90 dagen":
            d0, d1 = max(lo, hi - pd.Timedelta(days=89).to_pytimedelta()), hi
        elif preset == "Year to date":
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
        st.caption(f"Piek = {piek_omschrijving(peak_start, peak_end, weekdays_only)}  ·  "
                   f"dal = al de rest.")

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
    d_wb["vergoeding"] = vergoeding_in
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
    blokkade = "Laad eerst prijsdata (en een geldige periode) voor je verder kan."
elif stap == 2 and not heeft_verbruik:
    blokkade = "Upload je verbruik hieronder voor je verder kan."
elif stap == 2 and not vergoeding_bevestigd:
    blokkade = "Vul je vergoeding in, of vink 'geen vergoeding' aan, voor je verder kan."

st.divider()
nav = st.columns([1, 3, 1])
with nav[0]:
    if st.button("◀ Vorige", disabled=(stap == 1), width="stretch"):
        st.session_state["stap"] -= 1
        st.rerun()
with nav[1]:
    st.progress((stap - 1) / (len(STAPPEN) - 1))
    st.caption(f"**Stap {stap} van {len(STAPPEN)}: {STAPPEN[stap - 1]}**")
with nav[2]:
    if st.button("Volgende ▶", disabled=(stap == len(STAPPEN) or blokkade is not None), width="stretch"):
        st.session_state["stap"] += 1
        st.rerun()
if blokkade:
    st.warning(f"⚠️ {blokkade}")

st.divider()

if not data_geladen:
    if stap != 1:
        st.info("⬅️ Ga naar stap 1 en laad eerst prijsdata.")
    st.stop()

# --------------------------------------------------------------------------
# HEADER (stap 2 t.e.m. 4)
# --------------------------------------------------------------------------

if stap > 1:
    st.caption(f"Bron: {source}  ·  {d['datetime'].min():%d/%m/%Y} t/m {d['datetime'].max():%d/%m/%Y}"
              f"  ·  {len(d):,} kwartieren"
              f"  ·  piek {piek_omschrijving(peak_start, peak_end, weekdays_only)}"
              f"  ·  eenheid {eenheid}")

    p = d["price"]
    peak_avg = d.loc[d["is_peak"], "price"].mean()
    off_avg = d.loc[~d["is_peak"], "price"].mean()
    daily = d.groupby("date")["price"]

    c = st.columns(6)
    c[0].metric("Baseload", f"{p.mean()*factor:.{dec}f}", help=eenheid)
    c[1].metric("Peak", f"{peak_avg*factor:.{dec}f}" if pd.notna(peak_avg) else "—",
                f"{(peak_avg/p.mean()-1)*100:+.1f}% vs base" if pd.notna(peak_avg) and p.mean() else None)
    c[2].metric("Off-peak", f"{off_avg*factor:.{dec}f}" if pd.notna(off_avg) else "—")
    c[3].metric("Volatiliteit (σ)", f"{p.std()*factor:.{dec}f}")
    c[4].metric("Negatieve kwartieren", f"{d['is_negative'].mean()*100:.1f}%",
                f"{int(d['is_negative'].sum()):,} stuks", delta_color="off")
    c[5].metric("Gem. dagspread", f"{(daily.max()-daily.min()).mean()*factor:.{dec}f}")
    st.divider()

# ============================================================================
# ============================================================================
# STAP 2 — VERBRUIK & VERGOEDING
# ============================================================================
if stap == 2:
    st.markdown(f"""
Upload je kwartierverbruik en vul je leveringsvergoeding in — de eenheidsprijs
(totaal, piek, dal) wordt dan meteen berekend.

**Piek** = {piek_omschrijving(peak_start, peak_end, weekdays_only)} · **Dal** = al de rest.
Pas je de verhouding aan op stap 1, dan volgen deze cijfers en de werkmap mee.
Eenheidsprijs = totale kost / totaal verbruik (verbruiksgewogen). Kost per kwartier
= (prijs + vergoeding) × verbruik.
""")

    v1, v2 = st.columns(2)
    with v1:
        nieuw_bestand = st.file_uploader(
            "📤 Verbruik uploaden (CSV of Excel, met datum/tijd + verbruik in kWh)",
            type=["csv", "xlsx", "xls", "txt"], key="verbruik_upload")
        if nieuw_bestand is not None and nieuw_bestand.name != vbr_naam:
            try:
                ruw = _combineer_datum_tijd(read_upload(nieuw_bestand.getvalue(), nieuw_bestand.name))
                vbr = normalise_verbruik(ruw, dt_col="_datetime" if "_datetime" in ruw.columns else None)
                if vbr["datetime"].duplicated().any():
                    vbr = vbr.drop_duplicates("datetime", keep="last")
                cache["vbr_df"] = vbr
                cache["vbr_naam"] = nieuw_bestand.name
                gematcht = int(d["datetime"].isin(vbr["datetime"]).sum())
                if gematcht == 0:
                    cache["vbr_melding"] = ("warning", "Geen enkele tijdstempel uit het "
                                            "verbruiksbestand kwam overeen met de kwartieren in "
                                            "de gekozen periode — alle verbruik staat op 0.")
                elif gematcht == len(d):
                    cache["vbr_melding"] = ("success", f"Alle {gematcht:,} kwartieren gematcht "
                                            "met verbruiksdata.")
                else:
                    cache["vbr_melding"] = ("success", f"{gematcht:,} van de {len(d):,} "
                                            f"kwartieren gematcht ({gematcht/len(d)*100:.0f}%). "
                                            "De overige kwartieren zonder match kregen verbruik = 0.")
                st.rerun()
            except Exception as e:
                cache["vbr_melding"] = ("error", f"Kon het verbruiksbestand niet inlezen: {e}")
                st.rerun()
        if vbr_naam:
            st.caption(f"Actief bestand: **{vbr_naam}**")
            if vbr_melding:
                niveau, tekst = vbr_melding
                getattr(st, niveau)(tekst)
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
            st.caption("👆 Vul een bedrag in, of vink het vakje aan.")

    st.divider()
    k1, k2, k3 = st.columns(3)
    k1.metric("Kwartieren in de periode", f"{len(d_wb):,}")
    k2.metric("Waarvan piek", f"{int(piek_vast.sum()):,}" if piek_vast is not None else "—")
    k3.metric("Waarvan dal", f"{int((~piek_vast).sum()):,}" if piek_vast is not None else "—")

    if not heeft_verbruik:
        st.info("⬆️ Upload je verbruik hierboven om de eenheidsprijs te zien.")
    else:
        e1, e2, e3 = st.columns(3)
        for col, label, ep in ((e1, "Eenheidsprijs totaal", ep_totaal),
                               (e2, "Eenheidsprijs piek", ep_piek),
                               (e3, "Eenheidsprijs dal", ep_dal)):
            col.metric(label, f"{ep*factor:.{dec}f}" if ep is not None else "—", help=eenheid)

        g1, g2 = st.columns(2)
        with g1:
            fig = go.Figure(go.Bar(
                x=["Totaal", "Piek", "Dal"],
                y=[(v * factor if v is not None else 0) for v in (ep_totaal, ep_piek, ep_dal)],
                marker_color=[NAVY, ROOD, GREY]))
            fig.update_yaxes(title=eenheid)
            st.plotly_chart(style_fig(fig, 380, "Eenheidsprijs per periode"), **FULL)
        with g2:
            fig = go.Figure(go.Pie(labels=["Piek", "Dal"], values=[kost_piek, kost_dal],
                                   hole=0.5, marker_colors=[ROOD, GREY]))
            st.plotly_chart(style_fig(fig, 380, "Aandeel kost piek vs dal (EUR)"), **FULL)

# ============================================================================
# STAP 3 — EXPORT
# ============================================================================
else:
    st.subheader("Eenheidsprijs-werkmap")
    st.caption("Excel met levende formules — prijs, vergoeding en verbruik staan al ingevuld "
              "op het blad 'Kwartierdata'; 'Overzicht' herrekent zichzelf.")
    fname = f"belpex_eenheidsprijs_{d0:%Y%m%d}_{d1:%Y%m%d}.xlsx"
    if st.button("📗 Scholt-werkmap aanmaken", type="primary", width="stretch"):
        with st.spinner("Werkmap bouwen ..."):
            st.session_state["wb"] = build_workbook(
                d_wb, peak_start, peak_end, weekdays_only)
            st.session_state["wb_name"] = fname

    if st.session_state.get("wb"):
        st.download_button("⬇️ Werkmap opslaan", st.session_state["wb"],
                           st.session_state["wb_name"],
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           type="primary", width="stretch")

    st.divider()
    st.subheader("Volledige analyse")
    st.caption(f"Prijskolommen in de tabellen hieronder staan in {eenheid}; "
              "'Ruwe data' en de CSV blijven in EUR/kWh voor reproduceerbaarheid.")
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
    x.download_button("📊 Excel downloaden", buf.getvalue(),
                      f"belpex_{d0:%Y%m%d}_{d1:%Y%m%d}.xlsx",
                      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                      **FULL)
    y.download_button("📄 CSV downloaden",
                      d[["datetime", "price", "is_peak", "is_negative"]]
                      .to_csv(sep=";", decimal=",", index=False).encode(),
                      f"belpex_{d0:%Y%m%d}_{d1:%Y%m%d}.csv", "text/csv",
                      **FULL)

    st.dataframe(d[["datetime", "price", "is_peak", "is_negative"]].tail(200),
                 **FULL, height=400)
