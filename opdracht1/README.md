
Verbruiksanalyse Tool (Python)


Dit project bevat een Python-script voor het analyseren en vergelijken van energieverbruik op basis van een configuratiebestand en CBS-data.
Het script voert berekeningen uit over gas- en elektriciteitsverbruik per m², per m³, en in totaal. 
De resultaten worden geschreven naar een Excel-bestand op basis van een template, inclusief een visuele barchart.


📂 Bestandsstructuur


- opdracht1.py        → Hoofdscript voor berekening en output
- template.xlsx       → Excel-bestand met opmaak en velden waarin de resultaten worden geschreven
- cbsdata.xlsx        → Excelbestand met CBS-verbruiksgegevens (categorieën, bouwjaren, gemiddelde gas en elektriciteitsverbruik)
- config.xml          → Configuratiebestand met inputgegevens voor een specifieke analyse
- Opdracht1.xlsx      → De uiteindelijke output (Excelbestand met ingevulde data en grafiek)


⚙️ Benodigdheden


- Python 3.7+
- openpyxl
- pandas

Installeer de vereisten via pip:

pip install openpyxl pandas



🧾 Voorbeeld van config.xml

```xml
<config>
    <Naam>Bedrijf X</Naam>
    <Straat>Nassau Ouwerkerkstraat 3</Straat>
    <Postcode>2596CC</Postcode>
    <Plaats>Den Haag</Plaats>
    <Gas>130000</Gas>
    <Elektriciteit>27400</Elektriciteit>
    <EnergetischeWaardeGasElektra>10</EnergetischeWaardeGasElektra>
    <Verdiepingen>3</Verdiepingen>
    <Bouwjaar>2025</Bouwjaar>
    <Categorie>Detailhandel zonder koeling</Categorie>
    <Oppervlakte>5000</Oppervlakte>
    <HoogtePlafond>3</HoogtePlafond>
    <GemHoogtePlafond>2.7</GemHoogtePlafond>
    <CBSDatafile>cbsdata.xlsx</CBSDatafile>
    <Outputfile>Opdracht1.xlsx</Outputfile>
</config>
```

Voer het script uit met een XML-configuratiebestand als argument:

python opdracht1.py config.xml

Na afloop wordt een Excelbestand (Opdracht1.xlsx) gegenereerd met de verbruikscijfers, een grafiek en ingevulde klantinformatie.
