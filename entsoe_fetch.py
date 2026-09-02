# -*- coding: utf-8 -*-
"""
entsoe_fetch.py — kwartierprijzen day-ahead ophalen bij ENTSO-E.

Officiële bron, gedocumenteerde API, geen scraping.

  Endpoint : https://web-api.tp.entsoe.eu/api
  Artikel  : 12.1.D Day-ahead Prices
  Parameters:
      securityToken = jouw token
      documentType  = A44          (prijsdocument)
      in_Domain     = out_Domain   = biedzone (België: 10YBE----------2)
      periodStart / periodEnd      = yyyyMMddHHmm in UTC

Token aanvragen (eenmalig, gratis):
  1. Registreer op https://transparency.entsoe.eu
  2. Mail naar transparency@entsoe.eu met "Restful API access" als onderwerp
     en je registratie-e-mailadres in de body. Antwoord binnen ~3 werkdagen.
  3. Genereer het token onder 'My Account Settings'.

Beperkingen van de API:
  - maximaal één jaar per request (dit script hakt automatisch in stukken)
  - maximaal 400 requests per minuut per IP en per token
"""

from __future__ import annotations

import xml.etree.ElementTree as ET
from datetime import date, datetime, timedelta, timezone

import pandas as pd

API_URL = "https://web-api.tp.entsoe.eu/api"
TZ = "Europe/Brussels"

BIEDZONES = {
    "België (BE)": "10YBE----------2",
    "Nederland (NL)": "10YNL----------L",
    "Frankrijk (FR)": "10YFR-RTE------C",
    "Duitsland/Luxemburg (DE-LU)": "10Y1001A1001A82H",
}

RESOLUTIES = {
    "PT15M": timedelta(minutes=15),
    "PT30M": timedelta(minutes=30),
    "PT60M": timedelta(hours=1),
    "PT1H": timedelta(hours=1),
}


class EntsoeError(RuntimeError):
    pass


# --------------------------------------------------------------------------
# PARSER
# --------------------------------------------------------------------------

def _tag(elem) -> str:
    """Elementnaam zonder namespace."""
    return elem.tag.split("}")[-1]


def _find(parent, naam):
    for kind in parent:
        if _tag(kind) == naam:
            return kind
    return None


def _findall(parent, naam):
    return [k for k in parent if _tag(k) == naam]


def parse_xml(xml_tekst: str) -> pd.DataFrame:
    """
    Zet een Publication_MarketDocument om naar een DataFrame met
    kolommen 'datetime' (lokale Belgische tijd) en 'price' (EUR/MWh).

    Gaten in de posities worden opgevuld met de vorige waarde: bij curveType
    A03 laat ENTSO-E opeenvolgende gelijke prijzen weg.
    """
    root = ET.fromstring(xml_tekst.encode() if isinstance(xml_tekst, str) else xml_tekst)

    if _tag(root) == "Acknowledgement_MarketDocument":
        reden = _find(root, "Reason")
        tekst = "onbekende reden"
        if reden is not None:
            t = _find(reden, "text")
            c = _find(reden, "code")
            tekst = f"{t.text if t is not None else ''} (code {c.text if c is not None else '?'})"
        raise EntsoeError(f"ENTSO-E gaf geen data terug: {tekst}")

    rijen: list[tuple[datetime, float]] = []

    for ts in _findall(root, "TimeSeries"):
        for periode in _findall(ts, "Period"):
            interval = _find(periode, "timeInterval")
            res_el = _find(periode, "resolution")
            if interval is None or res_el is None:
                continue

            start_el, eind_el = _find(interval, "start"), _find(interval, "end")
            stap = RESOLUTIES.get(res_el.text.strip())
            if stap is None or start_el is None:
                continue

            start = datetime.fromisoformat(start_el.text.replace("Z", "+00:00"))
            eind = (datetime.fromisoformat(eind_el.text.replace("Z", "+00:00"))
                    if eind_el is not None else None)

            punten: dict[int, float] = {}
            for punt in _findall(periode, "Point"):
                pos_el = _find(punt, "position")
                pr_el = _find(punt, "price.amount")
                if pos_el is None or pr_el is None:
                    continue
                punten[int(pos_el.text)] = float(pr_el.text)

            if not punten:
                continue

            # aantal posities: uit de periodelengte, anders uit de hoogste positie
            if eind is not None:
                aantal = int((eind - start) / stap)
            else:
                aantal = max(punten)
            aantal = max(aantal, max(punten))

            vorige = None
            for pos in range(1, aantal + 1):
                if pos in punten:
                    vorige = punten[pos]
                if vorige is None:          # gat vóór het eerste punt
                    continue
                rijen.append((start + (pos - 1) * stap, vorige))

    if not rijen:
        return pd.DataFrame(columns=["utc", "price"])

    d = pd.DataFrame(rijen, columns=["utc", "price"])
    d["utc"] = pd.to_datetime(d["utc"], utc=True)
    # Alles blijft in UTC tot het allerlaatste moment. Zou je hier al naar
    # lokale tijd omzetten, dan vallen op de laatste zondag van oktober twee
    # uren op dezelfde lokale stempel en verdwijnt er één bij het ontdubbelen.
    # Latere TimeSeries zijn recentere revisies: die winnen.
    return (d.drop_duplicates("utc", keep="last")
            .sort_values("utc").reset_index(drop=True))


def naar_lokaal(d: pd.DataFrame) -> pd.DataFrame:
    """UTC-kolom omzetten naar Belgische wandkloktijd zonder tijdzone-info."""
    if d.empty:
        return pd.DataFrame(columns=["datetime", "price"])
    out = pd.DataFrame({
        "datetime": d["utc"].dt.tz_convert(TZ).dt.tz_localize(None),
        "price": d["price"].values,
    })
    return out.sort_values("datetime").reset_index(drop=True)


def naar_kwartieren(d: pd.DataFrame) -> pd.DataFrame:
    """
    Zet een reeks (met UTC-kolom) om naar een volledig kwartierraster.
    Uurwaarden van vóór de overstap naar de 15-minuten-MTU worden herhaald
    over de vier kwartieren. Het raster wordt in UTC opgebouwd, zodat de
    zomer-winteruurwissel intact blijft.
    """
    if d.empty:
        return d
    s = d.set_index("utc")["price"].sort_index()
    raster = pd.date_range(s.index.min(), s.index.max(), freq="15min", tz="UTC")
    uit = (s.reindex(s.index.union(raster)).ffill().reindex(raster)
           .rename_axis("utc").reset_index(name="price"))
    return uit


# --------------------------------------------------------------------------
# OPHALEN
# --------------------------------------------------------------------------

def _stamp(d: date, eind: bool = False) -> str:
    """Lokale dag -> UTC-tijdstempel yyyyMMddHHmm."""
    lokaal = pd.Timestamp(d) + (pd.Timedelta(days=1) if eind else pd.Timedelta(0))
    return (lokaal.tz_localize(TZ).tz_convert(timezone.utc)).strftime("%Y%m%d%H%M")


def fetch(token: str, van: date, tot: date,
          biedzone: str = BIEDZONES["België (BE)"],
          kwartieren: bool = True,
          timeout: int = 120) -> pd.DataFrame:
    """
    Haalt de day-ahead prijzen op tussen 'van' en 'tot' (beide inclusief).
    Hakt automatisch in stukken van maximaal één jaar.
    """
    import requests

    if not token or len(token.strip()) < 10:
        raise EntsoeError("Geen geldig ENTSO-E-token opgegeven.")
    if tot < van:
        raise EntsoeError("De einddatum ligt vóór de startdatum.")

    stukken, cursor = [], van
    while cursor <= tot:
        stuk_eind = min(tot, date(cursor.year, 12, 31))
        stukken.append((cursor, stuk_eind))
        cursor = stuk_eind + timedelta(days=1)

    delen = []
    for a, b in stukken:
        params = {
            "securityToken": token.strip(),
            "documentType": "A44",
            "in_Domain": biedzone,
            "out_Domain": biedzone,
            "periodStart": _stamp(a),
            "periodEnd": _stamp(b, eind=True),
        }
        r = requests.get(API_URL, params=params, timeout=timeout)

        if r.status_code == 401:
            raise EntsoeError("Token geweigerd (401). Controleer je security token.")
        if r.status_code == 429:
            raise EntsoeError("Te veel requests (429). Wacht 10 minuten en probeer opnieuw.")
        if r.status_code == 400:
            raise EntsoeError(f"Ongeldige query (400) voor {a} – {b}: {r.text[:200]}")
        r.raise_for_status()

        delen.append(parse_xml(r.text))

    d = pd.concat(delen, ignore_index=True) if delen else pd.DataFrame(
        columns=["utc", "price"])
    if d.empty:
        return pd.DataFrame(columns=["datetime", "price"])
    d = d.drop_duplicates("utc", keep="last").sort_values("utc")

    if kwartieren:
        d = naar_kwartieren(d)

    d = naar_lokaal(d)
    masker = ((d["datetime"] >= pd.Timestamp(van)) &
              (d["datetime"] < pd.Timestamp(tot) + pd.Timedelta(days=1)))
    return d[masker].reset_index(drop=True)


if __name__ == "__main__":
    import argparse

    p = argparse.ArgumentParser(description="ENTSO-E day-ahead prijzen ophalen")
    p.add_argument("--token", required=True)
    p.add_argument("--van", required=True, help="YYYY-MM-DD")
    p.add_argument("--tot", required=True, help="YYYY-MM-DD")
    p.add_argument("--uit", default="entsoe_prijzen.csv")
    a = p.parse_args()

    df = fetch(a.token, date.fromisoformat(a.van), date.fromisoformat(a.tot))
    df.to_csv(a.uit, index=False)
    print(f"{len(df):,} kwartieren -> {a.uit}")
    print(df.head())
