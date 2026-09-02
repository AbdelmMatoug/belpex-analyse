import pandas as pd
from entsoe_fetch import parse_xml as _px, naar_kwartieren, naar_lokaal, EntsoeError, _stamp
def parse_xml(x): return naar_lokaal(_px(x))
from datetime import date

NS = 'xmlns="urn:iec62325.351:tc57wg16:451-3:publicationdocument:7:0"'

def doc(periods):
    ts = "".join(periods)
    return f'<?xml version="1.0" encoding="UTF-8"?><Publication_MarketDocument {NS}><mRID>x</mRID><type>A44</type><TimeSeries><mRID>1</mRID><businessType>A62</businessType><currency_Unit.name>EUR</currency_Unit.name><price_Measure_Unit.name>MWH</price_Measure_Unit.name><curveType>A01</curveType>{ts}</TimeSeries></Publication_MarketDocument>'

def period(start,end,res,pts):
    p="".join(f"<Point><position>{i}</position><price.amount>{v}</price.amount></Point>" for i,v in pts)
    return f"<Period><timeInterval><start>{start}</start><end>{end}</end></timeInterval><resolution>{res}</resolution>{p}</Period>"

# 1) uurlijkse dag (24 punten), winter -> lokaal 00:00
d = parse_xml(doc([period("2026-01-01T23:00Z","2026-01-02T23:00Z","PT60M",[(i,10.0+i) for i in range(1,25)])]))
assert len(d)==24, len(d)
assert str(d.datetime.iloc[0])=="2026-01-02 00:00:00", d.datetime.iloc[0]
assert d.price.iloc[0]==11.0 and d.price.iloc[-1]==34.0
print("1 uurlijks winter OK", d.datetime.iloc[0], d.datetime.iloc[-1])

# 2) kwartier (96 punten) zomer -> UTC+2
d = parse_xml(doc([period("2026-07-01T22:00Z","2026-07-02T22:00Z","PT15M",[(i,float(i)) for i in range(1,97)])]))
assert len(d)==96
assert str(d.datetime.iloc[0])=="2026-07-02 00:00:00", d.datetime.iloc[0]
assert str(d.datetime.iloc[1])=="2026-07-02 00:15:00"
print("2 kwartier zomer OK")

# 3) gaten (curveType A03): posities 1,5,9 -> forward fill tot 12
d = parse_xml(doc([period("2026-03-02T23:00Z","2026-03-03T02:00Z","PT15M",[(1,5.0),(5,7.0),(9,9.0)])]))
assert len(d)==12, len(d)
assert list(d.price)==[5.0]*4+[7.0]*4+[9.0]*4, list(d.price)
print("3 gaten opvullen OK")

# 4) negatieve prijzen + revisie: tweede TimeSeries wint
xml = doc([period("2026-01-01T23:00Z","2026-01-02T00:00Z","PT60M",[(1,-25.5)])])
xml = xml.replace("</Publication_MarketDocument>",
   f'<TimeSeries><mRID>2</mRID><curveType>A01</curveType>{period("2026-01-01T23:00Z","2026-01-02T00:00Z","PT60M",[(1,-30.0)])}</TimeSeries></Publication_MarketDocument>')
d = parse_xml(xml)
assert len(d)==1 and d.price.iloc[0]==-30.0, d
print("4 negatief + revisie OK")

# 5) acknowledgement
ack='<?xml version="1.0"?><Acknowledgement_MarketDocument xmlns="urn:iec62325.351:tc57wg16:451-1:acknowledgementdocument:7:0"><Reason><code>999</code><text>No matching data found</text></Reason></Acknowledgement_MarketDocument>'
try:
    parse_xml(ack); raise SystemExit("FOUT: had moeten falen")
except EntsoeError as e:
    print("5 acknowledgement OK:", str(e)[:60])

# 6) uur -> kwartier expansie
d = parse_xml(doc([period("2026-01-01T23:00Z","2026-01-02T23:00Z","PT60M",[(i,10.0+i) for i in range(1,25)])]))
q = naar_lokaal(naar_kwartieren(_px(doc([period("2026-01-01T23:00Z","2026-01-02T23:00Z","PT60M",[(i,10.0+i) for i in range(1,25)])]))))
assert len(q)==93, len(q)
assert list(q.price[:5])==[11.0,11.0,11.0,11.0,12.0], list(q.price[:5])
print("6 uur->kwartier OK", len(q), "punten")

# 7) tijdstempels UTC
assert _stamp(date(2026,1,1))=="202512312300", _stamp(date(2026,1,1))
assert _stamp(date(2026,7,1))=="202606302200", _stamp(date(2026,7,1))
assert _stamp(date(2026,8,27),eind=True)=="202608272200", _stamp(date(2026,8,27),eind=True)
print("7 UTC-stempels OK")

# 8) DST-overgang oktober: dubbele uur
d = parse_xml(doc([period("2026-10-24T22:00Z","2026-10-25T23:00Z","PT60M",[(i,float(i)) for i in range(1,26)])]))
assert len(d)==25, len(d)
print("8 DST 25 uur OK", d.datetime.iloc[0], "->", d.datetime.iloc[-1])
print("\nALLE TESTS GESLAAGD")
