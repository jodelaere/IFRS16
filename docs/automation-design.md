# IFRS16 Monthly Roll-Forward — Automation Design

## Doel

Automatiseer de maandelijkse update van het IFRS16 Input Board Pack met behulp van
een Office Script (draait in de echte Excel Online-rekenmachine, dus "Refresh All"
en pivottabellen werken) getriggerd via Power Automate.

## Vastgestelde SharePoint-structuur

Site: `behohr-finance` (https://houseofhr.sharepoint.com/sites/behohr-finance)

```
Consolidation  annual statements/
  {jaar}/
    P{periode} {jaar}/                         (bv. "P8 2026", ook gezien als "P08.2025")
      IFRS 16/
        {yyyymm} - IFRS16 - 1 - Entity list Power BI - To refresh.xlsx
        {yyyymm} - IFRS16 - 2 - Review Anaplan vs Fluence.xlsx
        {yyyymm} - IFRS16 - 3 - Input Board Pack.xlsx      <-- doelbestand van deze automatisering
        {yyyymm} - IFRS16 - {PowerHouse} - Verschillen ...xlsx  (per PowerHouse, o.a. Accent, Cohedron)
        {yyyymm} - IFRS16 - Export 3.5 - DD{ddmmyyyy}.xlsx
        Anaplan Exports 2.92.10/
          Updated Lease properties (2).xlsx    <-- bron voor "2.9 input"
          Updated Lease properties (3).xlsx    <-- bron voor "2.10 input"
        Service PH/
```

Let op: de maand-submap-conventie is niet 100% consistent over de jaren heen
("P8 2026" vs "P08.2025" vs "P09-2025"); de Power Automate flow moet hierop
robuust zoeken (bv. op basis van periodenummer + jaar, niet exacte string-match).

## Bevestigde structuur "Input Board Pack.xlsx" (9 tabbladen totaal)

Twee tabbladen volledig ingelezen; de overige 7 (`2.9 input`, `2.10 input`,
`2_9 Output`, `2.10 Output`(?), `Pivots on 2.10`, en 2 andere) kon ik niet volledig
inlezen — de SharePoint-tekstextractie knipt af bij grote workbooks. Onderstaand
is wat we zeker weten, plus wat nog bevestigd moet worden.

### Tab "Mvt Schedule Details"
- Bovenaan: telling nieuwe contracten per PowerHouse (Buildings/Vehicles) voor de lopende periode (`mvt P08`).
- Tabel "BUILDINGS - NEW" (rij 18-25 in augustus-versie): kolommen
  `Key | Entity | Lease description | Lease commencement date | Reasonably certain end date selection | Reasonably certain end date | Lease duration | Fixed payment | Payment frequency | Asset category | Leased capacity | Lease Liability`.
  `Lease Liability` = `Fixed payment * Leased capacity` (kolom L = H*G in het voorbeeld — let op: dit lijkt een vereenvoudigde formule, geen echte discontering; te bevestigen).
  Er is een vergelijkbare tabel te verwachten voor "VEHICLES - NEW" (niet gezien in de uitgeknipte dump, maar de structuur boven refereert er impliciet naar).
- **Open vraag**: waar komt de rij-key (bv. `1109__Antwerpen_JVG7`) vandaan, en wat is exact de regel om te bepalen dat een contract "nieuw deze maand" is? Voorstel (zie script): een contract in de Anaplan-export met `Lease commencement date` in de huidige rapportageperiode én dat nog niet voorkomt in de bestaande "BUILDINGS/VEHICLES - NEW"-lijst van vorige maand.

### Tab "Movement schedule"
- Rij 1 = headers. Kolom G1 bevat een **formule** `=K34` (dus de "August 2026"-tekst
  is al automatisch gekoppeld aan een brondatum in K34 — dit hoeft niet handmatig
  aangepast te worden). Kolom G17 verwijst naar dezelfde cel.
- Kolommen K ("Plug 08 2026") en L ("Plug 07 2026") lijken **handmatige
  correctie-cijfers** (geen formule zichtbaar in de dump) — dit zijn vermoedelijk
  bewuste manuele boekingen door de preparer, dus NIET automatiseerbaar zonder
  menselijk oordeel. Voorstel: deze blijven een manuele stap; het script laat ze
  onaangeroerd en labelt ze in een reviewmail.
- Kolom N ("July 2026") en O ("mvt P08") lijken **statische tekst-headers** die elke
  maand handmatig verschoven worden (dit is vermoedelijk de "paar kolommen
  onderaan" die je noemde). **Te bevestigen**: of dit puur label-tekst is (veilig
  te scripten: N1 krijgt de waarde van het huidige G1, O1 wordt "mvt P{nieuwe
  periode}") of dat er ook celverwijzingen mee verschuiven.
- De cijfers per PowerHouse worden via `XLOOKUP` opgehaald uit `Pivots on 2.10`
  (aantallen) en `2_9 Output` (bedragen/totaal). Zodra `2.10 input`/`2.9 input`
  zijn overschreven en de pivottabellen ververst zijn, updaten deze automatisch —
  **geen** handmatige actie nodig als Office Script `pivotTable.refresh()` en
  `application.calculate(FullRebuild)` aanroept.

### Tabs "2.9 input" / "2.10 input" (niet rechtstreeks ingelezen)
Op basis van je beschrijving ("we overwriten de data in 2.10 en 2.9 input tabs
met de nieuwe excel geëxporteerd uit Anaplan") ga ik ervan uit dat deze tabs
1-op-1 dezelfde 28 kolommen hebben als de Anaplan-export
(`Updated Lease properties (2)/(3).xlsx`, tabblad "Sheet 1"):

```
Entity, Lease description, Cost Center, Local Cost center code,
Lease commencement date, Purchase option?, Exercise of Purchase option date,
Exercise price of purchase option, Reasonably certain end date selection,
Reasonably certain end date, Transfer IN Date, Transfer Out Date,
Lease duration, Fixed payment, Payment frequency,
Payment at beginning of period?, Revision type, Index or Rate,
Lease revision frequency (months), First revision after (months),
Reference index/rate date, Provision for dismantling costs, Status,
Last modification status, Asset category, Leased capacity, Type motor
```

**Te bevestigen**: exacte kolomvolgorde/naam in `2.9 input`/`2.10 input` zelf
(kan afwijken als er een extra sleutel-kolom is toegevoegd), en het beginpunt
(rij 1 = headers, rij 2 = eerste data-rij aangenomen).

### Entity list (bron: MDM via Power BI)

Naast de Anaplan-exports is er een derde brontabel, aangeleverd via
`{yyyymm} - IFRS16 - 1 - Entity list Power BI - To refresh.xlsx` (in dezelfde
`IFRS 16`-map). Dit bestand bevat een Power BI-connectie naar dataset
**BEHOHR-FDP-PRD-FINANCE**, tabel "18. Legal Entity Dimension" — de instructie
staat letterlijk in het bestand: *"1) REFRESH POWER BI TABLE FOR NEW ENTITIES
EVERY MONTH"*. Deze dimensie wordt uiteindelijk gevoed vanuit **MDM** (Master
Data Management) en bevat:

```
Legal Entity Code | Legal Entity Description | LE Powerhouse | LE Boutique | PH Fluence
```

plus een klein los mappingtabelletje (kolom G/H) `FDP/Anaplan-naam ↔ Fluence-naam`
voor PowerHouse-namen die tussen de twee systemen verschillen (bv. "ABY
Engineering" ↔ "House of ABY").

Deze tabel wordt maandelijks (1) ververst vanuit Power BI/MDM in dit losse
bestand, en (2) gekopieerd naar de **"Entity list"-tab** in het Input Board
Pack. Die tab wordt vervolgens gebruikt (vermoedelijk via XLOOKUP/VLOOKUP op
Legal Entity Code) in de "2.9 Output"/"2.10 Output"-tabs om elke leaseregel aan
de juiste PowerHouse te koppelen — dit is dus de schakel die de aantallen in
"Pivots on 2.10" en "Movement schedule" per PowerHouse correct laat optellen.

**Getest — directe Power BI-query is momenteel niet mogelijk.** Ik heb geprobeerd
de "18. Legal Entity Dimension"-tabel rechtstreeks te bevragen via de Power
BI-connector (DAX-query op dataset BEHOHR-FDP-PRD-FINANCE, workspace
`54172bbd-89f7-4b7c-a86d-786b11475eef`, report `5764f6b0-e1e1-426e-9417-790422fa6e62`),
om de tussenliggende Excel-refresh helemaal te kunnen overslaan. Dit gaf
`ArtifactAccessDenied` (403) — dit account/deze sessie heeft geen toegang tot
dat specifieke Power BI-rapport via de connector, ook al is er wel toegang tot
het Excel-bestand met dezelfde onderliggende Power BI-connectie.

**Gekozen aanpak (default) — via het bestaande Excel-bestand**: de
automatisering ververst `Entity list Power BI - To refresh.xlsx` (Refresh All)
en kopieert de resulterende tabel naar de "Entity list"-tab, net zoals het
huidige handmatige proces. Dit vermijdt de afhankelijkheid van directe Power
BI-API-toegang.

**Toekomstige verbetering (optioneel)**: als iemand met de juiste rechten
toegang tot dit Power BI-rapport/dataset laat toevoegen voor deze connector
(of voor de service account die de Power Automate-flow straks draait), kan
stap 3 in de flow hieronder vervangen worden door een rechtstreekse DAX-query
— dat elimineert de afhankelijkheid van een los, kwetsbaar Excel-bestand met
een Power BI-connectie die soms niet ververst.

**Nog wel te bevestigen** (ongeacht welke bron gekozen wordt): kan de Power
BI-connectie in `Entity list Power BI - To refresh.xlsx` automatisch ververst
worden binnen een Office Script/Power Automate-context (een live Power
BI/Analysis Services-connectie kan om interactieve herauthenticatie vragen),
of blijft die refresh een handmatige stap met alleen de kopieerstap naar
"Entity list" geautomatiseerd?

### Tabs "Pivots on 2.10" en "2_9 Output"
Niet ingelezen. Op basis van de XLOOKUP-formules in "Movement schedule" weten we:
- `Pivots on 2.10!B6:B19` / `C6:C19` = PowerHouse-naam → aantal (blok 1)
- `Pivots on 2.10!B27:B41` / `C27:C41` = PowerHouse-naam → aantal (blok 2, "Terminated")
- `Pivots on 2.10!B81:B91` / `C81:C91` = PowerHouse-naam → aantal (blok 3, "Transfers")
- `2_9 Output!F:F` (PowerHouse-naam) → `G:G`/`H:H` (totaal huidige/vorige periode)

Dit is vermoedelijk een echte Excel PivotTable gebouwd op `2.10 input`
(vandaar de naam). **Te bevestigen** door iemand met het bestand open: is dit
een native PivotTable-object (dan volstaat `pivotTable.refresh()`), of een
handmatig opgebouwde tabel met formules?

## Automatiseringsstappen (Power Automate flow)

1. **Trigger**: nieuw/gewijzigd bestand in `Anaplan Exports 2.92.10` van de
   lopende periode-map (of een vaste tijdsplanning, bv. 1x per maand na sluiting Anaplan).
2. **Nieuwe maandmap aanmaken**: kopieer de volledige `IFRS 16`-map (of specifiek
   het Input Board Pack-bestand) van de vorige periode naar de nieuwe periode-map
   (`sharepoint_copy_item`-achtige actie, native Power Automate "Copy file").
3. **Entity list verversen**: ververs (indien automatiseerbaar, zie open vraag
   hierboven) de Power BI-connectie in `Entity list Power BI - To refresh.xlsx`
   en lees de resulterende "Legal Entity Dimension"-tabel.
4. **Anaplan-data inladen**: lees de twee nieuwe exportbestanden (als tabel, via
   "List rows present in a table" — vereist dat de export als Excel-tabel is
   opgemaakt, of via een tussenstap die er een tabel van maakt).
5. **Office Script uitvoeren** op het gekopieerde Input Board Pack-bestand met
   als parameters: de entity list-rijen, de ingelezen Anaplan-rijen (2x), en de
   nieuwe/vorige periodelabels. Zie `scripts/ifrs16-monthly-rollforward.ts`.
6. **Notificatie**: stuur een mail/Teams-bericht met (a) de gedetecteerde nieuwe
   contracten die zijn voorgevuld in "Mvt Schedule Details" ter review, (b) een
   herinnering om de "Plug"-kolommen, eventuele header-kolommen (N/O) en de
   Entity list-refresh (indien niet automatiseerbaar) handmatig te controleren,
   (c) het resultaat van de CHECK-rij (moet 0 zijn).

## Openstaande punten voor validatie (graag bevestigen voor ik het script afrond)

1. Exacte rol van cel `K34` (Movement schedule) — wat staat erin en hoe wordt het gevuld?
2. Zijn "Pivots on 2.10" en "2_9 Output"/"2.10 Output" native PivotTables of formule-tabbladen?
3. Exacte kolomstructuur van "2.9 input"/"2.10 input" tabs (headers + startrij).
4. Bevestig de "VEHICLES - NEW"-tabel in "Mvt Schedule Details" (locatie/kolommen), aangezien die niet in de ingelezen data zat.
5. Is de kolomverschuiving in "Movement schedule" (N/O headers, evt. toevoegen nieuwe kolom elke maand) puur tekst, of moeten er ook formules mee verschoven worden?
6. Exacte locatie/kolomstructuur van de "Entity list"-tab in het Input Board Pack zelf (aangenomen: zelfde kolommen als "Legal Entity Dimension", zie hierboven), en hoe "2.9 Output"/"2.10 Output" die precies opzoeken (welke kolom, exacte range).
7. Kan de Power BI-connectie in `Entity list Power BI - To refresh.xlsx` automatisch ververst worden binnen een Power Automate/Office Script-context, of blijft dat een handmatige stap?
8. (Optioneel, lager prioriteit) Als directe Power BI-toegang later geregeld wordt: bevestig dat dataset BEHOHR-FDP-PRD-FINANCE / tabel "18. Legal Entity Dimension" via de MCP-connector query-baar wordt (nu `ArtifactAccessDenied` — zie hierboven) zodat de tussenliggende Excel-refresh vervangen kan worden door een rechtstreekse DAX-query.
