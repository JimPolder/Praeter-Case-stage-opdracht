===============================
Financiële Besparingsanalyse Tool (Python)
===============================

Dit project bevat een Python-script die een financiële analyse uitvoert op basis van een XML-configuratiebestand. Het berekent onder andere:

- Jaarlijkse kosten en baten
- EBITDA, cashflows, afschrijvingen
- IRR (interne rentevoet), REV (rendement eigen vermogen)
- TVT (terugverdientijd)
- Netto winst

De resultaten worden weggeschreven naar een Excel-template, inclusief gegevensinvoer.

-----------------------------
📂 Bestandsstructuur
-----------------------------

- opdracht2.py        → Hoofdscript voor berekening en output
- template.xlsx       → Excel-bestand met opmaak en velden waarin de resultaten worden geschreven
- config.xml          → Configuratiebestand met inputgegevens voor een specifieke analyse
- Opdracht2.xlsx      → De uiteindelijke output (Excelbestand met ingevulde data)

-----------------------------
⚙️ Benodigdheden
-----------------------------

- Python 3.7+
- openpyxl
- pandas

Installeer de vereisten via pip:

pip install pandas openpyxl

-----------------------------
🧾 Voorbeeld van config.xml
-----------------------------
<config>
    <Termijn>12</Termijn>
    <Inflatie>0.02</Inflatie>
    <Afschrijving>Linear</Afschrijving>
    <EigenVermogen>0.2</EigenVermogen>
    <RenteVV>0.04</RenteVV>
    <Belasting>0.165</Belasting>

    <HerinvesteringJaar>6</HerinvesteringJaar>
    <Investering>160000</Investering>
    <EenmaligeSubsidie>10000</EenmaligeSubsidie>
    <Restwaarde>0</Restwaarde>

    <Besparing>20000</Besparing>
    <JaarlijkseSubsidie>800</JaarlijkseSubsidie>

    <EenmaligeKosten>30000</EenmaligeKosten>
    <VasteExploitatieKosten>2000</VasteExploitatieKosten>
    <Herinvestering>4000</Herinvestering>

    <Outputfile>Opdracht2.xlsx</Outputfile>
</config>

Voer het script uit met een XML-configuratiebestand als argument:

python opdracht2.py config.xml

Na het uitvoeren wordt een bestand Opdracht2.xlsx gegenereerd met alle financiële projecties.

