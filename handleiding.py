"""
Handleiding — aparte pagina naast het dashboard.

Streamlit toont elk bestand in pages/ automatisch als extra tabblad in de
zijbalk. Deze pagina is louter uitleg voor de eindgebruiker; ze rekent niets.
"""
import streamlit as st

st.set_page_config(page_title="Handleiding — Scholt", page_icon=":material/menu_book:", layout="wide")
st.markdown("<style>.block-container {padding-top: 2rem; max-width: 900px;}</style>",
            unsafe_allow_html=True)

st.title("Handleiding")
st.markdown("""
Deze tool berekent wat elektriciteit **per kWh** kost wanneer je aan de Belpex-kwartierprijs
betaalt. Hij legt de marktprijs van elk kwartier naast het verbruik van dat kwartier en
geeft een eenheidsprijs voor **totaal**, **piek** en **dal**. Het resultaat kan je
downloaden als Excel-werkmap.

De tool werkt in **3 stappen**. Met **Volgende** en **Vorige** onderaan elke stap ga je heen en weer.
Wat je instelt, blijft bewaard zolang het tabblad openstaat.
""")

with st.container(border=True):
    st.markdown("""
**Wat heb je nodig?**
- Een **verbruiksbestand per kwartier** (CSV of Excel), bijvoorbeeld de export uit het
  Fluvius-portaal. Elke lijn is één kwartier, met een datum/tijd en een verbruik.
- De **leveringsvergoeding** van de leverancier in EUR/kWh (de marge bovenop de
  Belpex-prijs). Is er geen, dan vink je dat aan.
""")

st.header("Stap 1 — Data & periode")
st.markdown("""
**Prijzen.** De databank met Belpex-kwartierprijzen is normaal al gevuld; bovenaan staat
van wanneer tot wanneer. Ontbreken de laatste dagen, open dan
**Prijzen bijwerken of zelf een prijsbestand gebruiken** en klik op **Marktprijzen ophalen**. De prijzen worden
online opgehaald en aan de databank toegevoegd. Heb je zelf een prijsbestand, upload
het dan rechts (optie B).

Kies daarna onder *Welke data wil je gebruiken?* meestal **Databank**.

**Periode.** Kies een snelkeuze (bv. *Per jaar* of *Laatste 90 dagen*) of *Zelf kiezen*.
Kies een periode waarin het verbruik van de klant valt.
Kwartierprijzen bestaan pas vanaf **1 augustus 2025**.

**Piek / dal.** Bepaalt welke kwartieren als piek tellen:

| Keuze | Piekuren |
|---|---|
| Standaard | 07:00 t.e.m. 21:45 (15 uur) |
| Scholt | 08:00 t.e.m. 19:45 (12 uur) |
| Zelf kiezen | met een schuifbalk |

Met *Piek enkel op werkdagen* zijn zaterdag en zondag volledig dal. De rode balk onder de keuze
toont de 96 kwartieren van een dag: rood is piek, de rest is dal.
Deze keuze werkt door in alle cijfers en in de werkmap.

**Eenheid.** Enkel hoe prijzen getoond worden (EUR/kWh of EUR/MWh); de berekening
verandert niet.
""")

st.header("Stap 2 — Verbruik & vergoeding")
st.markdown("""
**1. Verbruik uploaden.** Sleep het verbruiksbestand in het uploadvak. De tool zoekt zelf
uit welke kolom de datum/tijd is en welke het verbruik. Kolomnamen, scheidingsteken
(`;` `,` tab) en notatie (`31,5` of `31.5`) maken niet uit. Wh en MWh worden omgezet
naar kWh. Bij een Fluvius-export met registers wordt injectie weggelaten en worden dag
en nacht opgeteld.

**2. Melding nakijken.** Na de upload verschijnt een melding, bijvoorbeeld:

> Herkend: tijdstip uit kolom **From**, verbruik uit kolom **Value** — 2.784 rijen,
> 01/09/2026 t/m 29/09/2026, totaal 140.537 kWh.

Kijk of het **totaal in kWh** en de **datums** kloppen met wat je verwacht.

**3. Klopt het niet? Kolommen zelf kiezen.** Open
**Kolommen controleren of zelf kiezen**. Het staat vooraf ingevuld met wat de tool
herkende. Pas aan wat fout is:
- *Lijnen overslaan bovenaan*: als er tekst (titel, klantnaam...) boven de tabel staat.
- *Eerste lijn bevat kolomnamen*: uitvinken als het bestand geen kopregel heeft.
- *Datum/tijd*, *Uur apart* (als datum en uur in twee kolommen staan), *Verbruik*, *Eenheid*.

Klik op **Toepassen**. Lukt de automatische herkenning helemaal niet, dan staat dit
paneel meteen open.

**4. Vergoeding.** Vul de leveringsvergoeding in EUR/kWh in, **of** vink
*Ik betaal geen aparte leveringsvergoeding* aan.

**Resultaat.** Onder de upload verschijnt de **eenheidsprijs** groot, met daaronder de prijs in
**piek** en **dal**, het totale verbruik (en welk deel in piek valt) en de totale kost.
""")
st.latex(r"\text{Eenheidsprijs} = \frac{\sum (\text{prijs} + \text{vergoeding}) \times \text{verbruik}}{\sum \text{verbruik}}")
st.caption("Per kwartier: kost = (Belpex-prijs + vergoeding) × verbruik. "
           "Dure kwartieren met veel verbruik wegen dus zwaarder door.")
st.markdown("""
Met **Gestandaardiseerd verbruik (CSV)** download je het verbruik zoals de tool het
gelezen heeft (één lijn per kwartier, verbruik in kWh), handig om na te kijken.
""")

st.header("Stap 3 — Export")
st.markdown("""
- **Scholt-werkmap aanmaken** → daarna **Werkmap opslaan**. Dit is de Excel met
  prijs, vergoeding en verbruik per kwartier en een overzichtsblad met formules die
  zichzelf herrekenen.
- **Marktprijzen in de periode**: gemiddelde, piek- en dalprijs van Belpex zelf, los van het
  verbruik van de klant.
- **Analyse als Excel** / **Kwartierprijzen als CSV**: de marktprijzen zelf
  (maandoverzicht, kwartier- en uurprofiel, dagreeks, ruwe data).
""")

st.header("Problemen oplossen")
st.markdown("""
| Wat je ziet | Wat te doen |
|---|---|
| **Volgende** blijft grijs | Lees de melding erboven: er ontbreekt nog iets (prijsdata, verbruik of de vergoeding). |
| *Geen enkele tijdstempel kwam overeen…* | De periode op stap 1 valt niet samen met het verbruik. Ga terug en kies een periode die de datums van het verbruik bevat. |
| *Kon het bestand niet automatisch inlezen* | Kies de kolommen zelf in het paneel **Kolommen controleren of zelf kiezen**. |
| *De rijen liggen 60 min uit elkaar* | Het bestand bevat uurwaarden, geen kwartierwaarden. Vraag een export per kwartier. |
| De app toont *“This app has gone to sleep”* | Klik op de knop om ze te wekken en wacht een minuut. Gebeurt na een tijd zonder gebruik. |
| De databank mist recente dagen | Stap 1 → **Prijzen bijwerken…** → **Marktprijzen ophalen**. |
""")
