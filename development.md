# Introductie
we gaan deze library anders implementeren, naar de creatie van 2 templates:
een aanlevertemplate en een infotemplate. 
De aanlevertemplate is de template die gebruikt wordt om de data aan te leveren. Dit is de bestaande template die je met otlmow-template kan maken. Hier moet niets aan aangepast worden, functioneel
De infotemplate is de template die ter informatie kan gebruikt worden om de structuur van de data inzichtelijk te maken. 
Het doel om dus om 2 aparte klassen bij te maken die de twee templates kunnen aanmaken. Daarbij zal er nog een gemeenschappelijke 
klasse worden gebruikt waarin de functies zitten die door beiden worden gebruikt.

er is reeds een script voor een info template te maken, dat is orig_infotemplate_generator.py
Als je subset /home/davidlinux/PycharmProjects/OTLMOW-GUI/otlmow_gui/demo_projects/simpel_vergelijkings_project/simpele_vergelijkings_subset.db
gebruikt (een eenvoudige subset om te testen)
dan krijg je het bestand info_template_voorbeeld.xlsx.

# stap 1: refactoring
We gaan dat dit script nu refactoren zodat het zoveel mogelijk van otlmow-template gebruikt.
bv. subsets inladen via otlmow-modelbuilder en andere zaken...
Maak eerst een unit test adhv bovenstaande info met de subset (kopieer die naar deze repo) zodat we zeker zijn dat dit correct wordt gerefactord
Maak ook eerst meer unit testen om de bestaande functionaliteiten te waarborgen
Daarna pas refactoren naar aparte klassen zoals in de intro

✅ **STAP 1 VOLTOOID**
- Unit tests aangemaakt voor InfotemplateGenerator (UnitTests/Subset/InfotemplateGenerator_test.py)
- Originele script gerefactord naar aparte klassen:
  - `TemplateGeneratorBase` - gemeenschappelijke basisklasse (collector laden, class filtering)
  - `AanlevertemplateGenerator` - aanlevertemplate generatie (uit SubsetTemplateCreator gehaald)
  - `InfotemplateGenerator` - infotemplate generatie (geërfd van TemplateGeneratorBase)
  - `SubsetTemplateCreator` - facade voor backward compatibility
- Gebruikt OSLOCollector uit otlmow-modelbuilder voor subsets en relaties
- Output is identiek aan originele generator

# stap 2 Extra functionaliteiten toevoegen (alleen voor infotemplate)

## 1. Enkel relaties met concrete klasses opnemen in de relaties tab bij eigenschappen
- Filter abstracte/base klasses uit (zoals OTLCodelijst, OTLLegenda, etc.)
- Alleen concrete implementatie-klasses tonen in de relaties tab

## 2. 2de heeftBetrokkene relatie
- Als de subset een `heeftBetrokkene` relatie bevat, moeten er **2 records** aangemaakt worden voor deze relatie
- Normaal zou je deze maar 1x maken, maar voor deze specifieke relatie dubbel
- Met link naar waar de rollen/betrokkenen te vinden zijn

## 3. Kleuren voor generieke attributen
- Attributen met de volgende URI's krijgen een zacht lichtgroene achtergrondkleur in Excel:
  - `https://wegenenverkeer.data.vlaanderen.be/ns/implementatieelement#AIMObject.assetId`
  - `https://wegenenverkeer.data.vlaanderen.be/ns/implementatieelement#AIMObject.datumOprichtingObject`
  - `https://wegenenverkeer.data.vlaanderen.be/ns/implementatieelement#AIMObject.bestekPostNummer`
  - `https://wegenenverkeer.data.vlaanderen.be/ns/implementatieelement#AIMNaamObject.naam`
  - `https://wegenenverkeer.data.vlaanderen.be/ns/implementatieelement#AIMDBStatus.isActief`
  - `https://wegenenverkeer.data.vlaanderen.be/ns/implementatieelement#AIMToestand.toestand`
  - `https://wegenenverkeer.data.vlaanderen.be/ns/implementatieelement#AIMVersie.assetVersie`
  - `https://wegenenverkeer.data.vlaanderen.be/ns/implementatieelement#AIMObject.typeURI`
  - `https://wegenenverkeer.data.vlaanderen.be/ns/implementatieelement#AIMObject.theoretischeLevensduur`
  - `https://wegenenverkeer.data.vlaanderen.be/ns/implementatieelement#AIMObject.standaardBestekPostNummer`
  - `https://wegenenverkeer.data.vlaanderen.be/ns/implementatieelement#AIMObject.notitie`

## 4. Default wie levert aan: aannemer
- In de infotemplate: kolom "wie levert aan" standaard ingevuld op "aannemer"
- Kan ook "AWV" zijn, maar default = "aannemer"

## 5. TypeURI al invullen bij opmerkingen
- **Alleen voor de eigenschap typeURI**: de typeURI van de klasse invullen in de **opmerkingen kolom** van de infotemplate
- Niet voor alle attributen, enkel voor typeURI eigenschap

## 6. Opmaak tabellen: headers in vetgedrukt
- Alle tabellen in de infotemplate: header rij vetgedrukt maken

## 7. Opmerkingen/isActief/toestand default waardes
- **isActief** kolom: default `TRUE`, wordt `FALSE` wanneer het asset uit de databank verwijderd moet worden
- **toestand** kolom: default `"in-gebruik"`

## 8. Keuzelijst URI link in formaat/opmaak/naamconventie
- Bij attributen die een keuzelijst (codelijst) hebben: link naar de URI van die codelijst toevoegen in de kolom **formaat/opmaak/naamconventie**

## 9. Naam is een generiek attribuut
- `naam` (AIMNaamObject.naam) toevoegen aan de lijst van generieke attributen (krijgt dus ook de lichtgroene kleur)

## 10. Filter toevoegen op attributen header rij
- Excel auto-filter (dropdown pijltjes) activeren op de header rij van het **attributen tabblad**

## 11. Relaties: richting toevoegen in E relatie met verzetten naar A
- **Voor alle relaties**: kolom "richting" toevoegen in kolom E (met waarden `<===`, `===>` of `<===>`)
- De bestaande kolom "relatie met" van E-relaties tabblad → naar **A-relaties tabblad** verplaatsen

## 12. Tabbladen alfabetisch sorteren
- **Alle Excel tabbladen (sheets)** alfabetisch A-Z sorteren

## 13. Datatemplate/aanlevertemplate vs info template
- Alleen documenteren (ter informatie), geen functionele wijziging

---

**Voorbeeld bestand:** `info_template_voorbeeld.xlsx` (in repo root en UnitTests/Subset/)
**Test subset:** `/home/davidlinux/PycharmProjects/OTLMOW-GUI/otlmow_gui/demo_projects/simpel_vergelijkings_project/simpele_vergelijkings_subset.db`

TODO:

- ontbrekend type toestand - ✅ GEREFACTORED: opmerking "Is bijna altijd 'in-gebruik'" toegevoegd voor toestand attributen
- keuzelijst link - ✅ GEREFACTORED: codelist URI link toegevoegd voor attributen met type dat begint met 'Kl'
