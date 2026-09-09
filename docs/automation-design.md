# IFRS16 Monthly Roll-Forward — Automation Design

Automatisering van de maandelijkse update van het IFRS16 Input Board Pack, via een
Office Script (draait in de echte Excel-rekenmachine) getriggerd door Power Automate.

Alles in dit document is geverifieerd tegen
`202608 - IFRS16 - 3 - Input Board Pack.xlsx` (P8 2026) — tabnamen, tabelnamen,
celverwijzingen, formules, Power Query M-code en de connecties.

## SharePoint-structuur

Site: `behohr-finance`

```
Consolidation  annual statements/{jaar}/P{periode} {jaar}/IFRS 16/
  {yyyymm} - IFRS16 - 1 - Entity list Power BI - To refresh.xlsx
  {yyyymm} - IFRS16 - 2 - Review Anaplan vs Fluence.xlsx
  {yyyymm} - IFRS16 - 3 - Input Board Pack.xlsx      <-- doelbestand
  {yyyymm} - IFRS16 - Export 3.5 - DD{ddmmyyyy}.xlsx
  Anaplan Exports 2.92.10/
    Updated Lease properties (2).xlsx    -> 2.9 Input   (11.329 rijen)
    Updated Lease properties (3).xlsx    -> 2.10 Input  (24.617 rijen)
```

Map- en bestandsnamen zijn **niet consistent** over de maanden heen
(`P5 2026/IFRS16` vs `P8 2026/IFRS 16`; `2026.06 - …`, `#2026.07 - …`,
`202608 - …`). De flow moet dus zoeken op periode + jaar, niet op exacte string.

## Opbouw van het Input Board Pack (10 tabbladen)

| # | Tab | Inhoud |
|---|---|---|
| 1 | `Mvt Schedule Details` | Aantal nieuwe contracten per PowerHouse (rijen 3-14) + detailtabel "BUILDINGS - NEW" (rij 18+) |
| 2 | `Movement schedule` | De movement schedule zelf: 2 blokken (buildings rij 3-14, vehicles rij 19-30) + presentatietabel rij 33-47 + CHECK rij 49 |
| 3 | `Pivots on 2.10` | 4 PivotTables (IN / OUT / TRANSFER IN / TRANSFER OUT) + formuleblok Transfers |
| 4 | `2_10 Output` | Power Query-output, Excel-tabel `_2_10_Output` (A1:R24618) |
| 5 | `2.10 Input` | Anaplan-export, Excel-tabel `Table1` (A1:AB24618) |
| 6 | `2_9 Output` | Power Query-output, tabel `_2_9_Output` (A2:D157) + PivotTable in F2:I16 |
| 7 | `2.9 Input` | Anaplan-export, Excel-tabel `Table2.9` (A1:AB11330) |
| 8 | `Info` | Documentatie van de 2.9-selectielogica |
| 9 | `Entity List` | **Hard copy** van de entiteitenlijst (4 kolommen, rij 3+) — hier kijken alle lookups naar |
| 10 | `Entity List PowerBI` | Nieuw toegevoegd: live Power BI-connectie, tabel `Table_ExternalData_1` (A3:S361) |

### De dataflow

```
Anaplan export (2) ──> 2.9 Input  (Table2.9) ──PQ──> 2_9 Output ──> PivotTable2 (F2:I16)
                                                                          │
Anaplan export (3) ──> 2.10 Input (Table1)   ──PQ──> 2_10 Output ──> 4 PivotTables
                                                                          │
Entity List (kolom C = PowerHouse) ──XLOOKUP──> PH-kolom in beide outputs │
                                                                          v
                                                        Movement schedule (XLOOKUP)
                                                                          v
                                                        Presentatietabel rij 33-47
```

### Movement schedule — kolomindeling

Beide blokken (buildings rij 3-14, vehicles rij 19-30) hebben dezelfde structuur:

| Kolom | Inhoud | Herkomst |
|---|---|---|
| A | PowerHouse | vast |
| B | December 2025 (opening) | `=B35` / `=G35` — uit de presentatietabel |
| C | New contracts | `XLOOKUP` → `Pivots on 2.10'!$B$6:$B$19` |
| D | M&A | hardcoded 0 |
| E | Terminated contracts | `-XLOOKUP` → `$B$27:$B$41` |
| F | Transfers between PHs | `XLOOKUP` → `$B$81:$B$91` |
| G | Huidige maand | `XLOOKUP` → `2_9 Output'!F:F` → kolom G (buildings) / H (vehicles) |
| I | Calculated | `=SUM(B:F)` |
| J | Difference | `=G-I` |
| L | Plug huidige maand | **manueel** — telt op in de E-formule |
| M | Plug vorige maand | historiek |
| O | Vorige maand | **getypte waarde** (het vorige-maandcijfer) |
| P | mvt P{nn} | `=G-O` — de echte maandmutatie |

### Wat de plug is

De plug in kolom L zit **in de Terminated contracts-formule**:

```
E8: =-IFERROR(XLOOKUP(A8,'Pivots on 2.10'!$B$27:$B$41,…),0)+L8
```

Hij reconcilieert de YTD-opbouw `B+C+D+E+F` naar de 2.9-telling in kolom G, en is
dus het **verschil tussen de 2.10- en de 2.9-cut** — geen vorige-maand-constructie.

Getoetst op P8: Covebo buildings heeft pivot OUT = 51 en plug −1, dus E = −52;
`399+140+0−52+0 = 487 = G`, Difference = 0. Healthcare: `−13 + −2 = −15`;
`118+8−15 = 111 = G` ✓

Gevolg: de plug moet bij elke nieuwe cut opnieuw bepaald worden. De
Difference-kolom is daarbij precies de nog benodigde plug, dus het reviewrapport
van het script geeft per PowerHouse dat bedrag. **Bewust niet automatisch
ingevuld** — de tie forceren zou echte datafouten maskeren.

> ℹ️ In een eerdere versie was kolom O `=SUM(H3:L3)`, wat neerkwam op *huidige
> maand + plug* en dus niet de vorige maand voorstelde. Dat is gecorrigeerd:
> O bevat nu de werkelijke vorige-maandcijfers, waardoor `P = G − O` pas een
> echte maandmutatie geeft. Het script vult O nu automatisch — kolom G bevat
> immers nog de vorige maand tot de nieuwe Anaplan-data ingeladen wordt.

**De maandkop staat op één plek**: `G1` en `G17` zijn formules `=K34`, dus de
huidige-maandlabel wordt in de presentatietabel (rij 34) getypt en beide
blokheaders volgen automatisch. Dit is de "paar kolommen onderaan" uit de
procesbeschrijving.

`B34`/`G34` ("December 2025") en `B35:B46`/`G35:G46` zijn de jaaropening en
blijven het hele jaar staan.

### Mvt Schedule Details

- Rijen 3-14: aantal nieuwe contracten per PowerHouse, buildings (kolom B) en
  vehicles (kolom E). **Nu hardcoded getypt** — gelijk aan `Movement schedule`
  kolom P. `E15` is zelfs een hardcoded 21 in plaats van een `SUM`.
- Rij 18+: tabel "BUILDINGS - NEW" met kolommen
  `key | Entity | Lease description | commencement | end date selection | end date | duration | fixed payment | frequency | asset category | leased capacity | Lease Liability`.
  `Lease Liability` = `=H*G` (fixed payment × duration).
- **Er is géén "VEHICLES - NEW"-tabel.** Vehicles worden geteld, niet
  gedetailleerd — conform het proces zoals beschreven.

## Bevindingen

### 1. ⚠️ De rapportageperiode komt uit de systeemklok

De Power Query M-code leidt de rapportagemaand af uit `DateTime.LocalNow()`:

```m
"New leases before and in current month (IN)" =
    each Date.Month([Lease commencement date]) <= Date.Month(DateTime.LocalNow())-1
```

De conventie is dus: *refresh in maand M → rapportagemaand M−1*. Gevolgen:

- **Niet reproduceerbaar.** Wie het P8-pack in oktober opent en Refresh All doet,
  krijgt september-cijfers in het augustus-pack. Een afgesloten board pack dat
  verandert wanneer je het later opent, is een controleprobleem.
- **De decemberafsluiting breekt.** Voor P12 refresh je in januari:
  `Date.Month(now)-1` = `1-1` = **0**. Geen enkel contract heeft maand ≤ 0, dus
  new contracts en terminated contracts worden allemaal 0 — precies bij de
  jaarafsluiting.
- **Het script kan de maand niet besturen.** Zolang de queries de klok gebruiken,
  kan de automatisering niet bepalen welke periode ze berekenen. Dit is dus geen
  losse verbetering maar een **randvoorwaarde** voor de automatisering.

De oplossing staat bij de eenmalige template-wijzigingen hieronder.

### 2. ⚠️ De nieuwe "Entity List PowerBI"-tab is nog niet aangesloten

Beide output-tabs halen de PowerHouse nog uit de **oude** hard copy:

```
2_9 Output!D3  =XLOOKUP(A3, 'Entity List'!A:A, 'Entity List'!C:C)
2_10 Output!R2 =XLOOKUP(Q2, 'Entity List'!A:A, 'Entity List'!C:C)
```

De nieuwe tab heeft de PowerHouse in **kolom F** (kolom C is `LE Country Long`)
en begint op rij 4. Bovendien heeft de oude tab 356 datarijen en de nieuwe 358 —
de hard copy loopt dus twee entiteiten achter.

### 3. ⚠️ Het Transfers-blok koppelt rijen op positie

`Pivots on 2.10` rijen 82-91 berekenen transfers als `=C50-C68`, `=C51-C69`, …:
harde celverwijzingen naar TRANSFER IN (rij 50-59) en TRANSFER OUT (rij 68-77).
Die twee pivots tonen **alleen PowerHouses die deze maand transfers hebben**.
Deze maand hebben beide exact dezelfde 9 PowerHouses, dus het klopt. Zodra een
PowerHouse alleen transfers *in* of alleen *uit* heeft, verschuiven de rijen en
trekt de formule stilzwijgend de verkeerde PowerHouse af.

Idem voor de lookup-ranges: `$B$6:$B$19` sluit precies aan op PivotTable3
(B6:E20, 12 PowerHouses + Grand Total) — **geen marge**. Een 13e PowerHouse laat
de pivot groeien en de lookup mist hem (via `IFERROR` → 0, dus zonder foutmelding).

### 4. ℹ️ Transfers hebben geen maandgrens (nu zonder gevolg)

`Transfers (IN)` en `Transfers (OUT)` gebruiken enkel `Date.IsInCurrentYear(...)`,
terwijl de IN/OUT-vlaggen elk *twee* voorwaarden hebben (jaar én maand). Een
transfer gedateerd na de rapportagemaand zou dus meetellen.

**Getoetst op de P8-data: geen effect.** Van de 399 transfer-IN en 399
transfer-OUT rijen in 2026 valt er geen enkele na 31/08/2026 — Anaplan registreert
transfers pas wanneer ze gebeuren. (399/399 matcht ook exact de Grand Totals van
beide TRANSFER-pivots.)

Alleen relevant als Anaplan ooit toekomstgedateerde transfers zou gaan bevatten.
Wordt de parameter uit template-wijziging A doorgevoerd, dan is een maandgrens
toevoegen triviaal — maar het is nu geen prioriteit.

### 5. 🔴 Lease Liability is fout voor kwartaalcontracten

`Lease Liability` = `=H*G` = `fixed payment × lease duration (maanden)`. Dat klopt
alleen voor maandelijkse betalingen. Een kwartaalcontract wordt met **factor 3**
overschat, want de kwartaalbetaling wordt 36× geteld in plaats van 12×.

**Bevestigd als fout** — de juiste basis is *betaling × aantal betalingen*.

In de Anaplan-data komen exact twee frequenties voor: `Monthly` en `Quarterly`.
Quarterly zit bijna uitsluitend bij buildings (128 van 129 in 2.10 Input),
precies de populatie die in "BUILDINGS - NEW" belandt — dus dit raakt vrijwel
elke maand waarin een nieuw kwartaalcontract start.

Juiste formule (zie template-wijziging E):

```
=H19*G19/IFS(I19="Monthly",1,I19="Quarterly",3)
```

`IFS` heeft bewust geen fallback: een onbekende frequentie geeft `#N/A` in plaats
van een stil verkeerd cijfer.

**Impact op P8 2026**: rij 19 (ClickCare Antwerpen, € 13.663,75 per kwartaal,
36 maanden) gaat van € 491.895 naar € 163.965. Het totaal `L25` daalt daardoor
van **€ 675.778,87 naar € 347.848,87** — de overige vijf contracten zijn
maandelijks en blijven ongewijzigd. Te beslissen of P8 herzien wordt.

### 6. ℹ️ Entity code `2XXX` — latent risico, geen actueel cijferprobleem

`2XXX - 2XXX-ZorgXchange` is in Anaplan een volwaardige entiteit met een
**placeholder-code**, en komt niet voor in de Legal Entity Dimension (358
entiteiten, geen match). Ze heeft 54 contracten in 2.10 Input en 9 in 2.9 Input,
alle vehicles.

Omdat beide Power Queries de entity code naar `Int64.Type` casten en `"2XXX"`
geen getal is, sneuvelen die rijen in de output. Dat klinkt ernstig, maar
onderzocht blijkt het **geen effect op de cijfers** te hebben:

- **52 van de 54** hebben een Transfer Out, allemaal gedateerd **2024** — het
  wagenpark is op 1 april 2024 overgezet. Elk kenteken komt terug onder een echte
  entiteit met Transfer IN op die datum: 8 onder `2104` (TMI AP B.V.), 1 onder
  `2106`. De `2XXX`-rijen zijn de herkomstzijde van die transfer, en worden per
  ontwerp uitgesloten — de 2.9-query filtert `[Transfer Out Date] = null`
  ("zodat een contract slechts 1x in rekening genomen wordt") en beide pivots
  filteren `Transfers (OUT)` op `(blank)`.
- De **2 rijen zonder Transfer Out** hebben geen commencement, einddatum of
  transfer in 2026 en zetten dus geen enkele vlag aan.

**Netto: 0 rijen die een beweging in 2026 zouden veroorzaken.** De contracten
tellen correct mee onder 2104 en 2106; TMI staat terecht op 425.

Wat blijft: zodra iemand een **nieuw** contract op deze entiteit boekt, verdwijnt
het zonder foutmelding. Daarom rapporteert het roll-forward script voortaan elke
entity code uit de Anaplan-input die niet in de entiteitenlijst voorkomt, met het
aantal contracten. Opruimen (echte code in Anaplan + entiteit in MDM, of de
entiteit afsluiten) is netjes maar niet urgent.

## Eenmalige template-wijzigingen

Uit te voeren in Excel Desktop op het P8-bestand, dat daarna de template is die
elke maand meekopieert.

### A. Rapportageperiode als parameter (randvoorwaarde, zie bevinding 1)

Het setup-script maakt de named cell `ReportingPeriodEnd` aan (Info!B7). Daarna
alleen nog de M-code vervangen — **plak de volledige queries** uit
`scripts/power-query/`, doe geen losse zoek-vervang-acties:

1. Data → Queries & Connections → rechtsklik **2/10 Output** → Edit → Home →
   Advanced Editor → alles selecteren → plak [`2-10-Output.m`](../scripts/power-query/2-10-Output.m) → Done.
2. Idem voor **2/9 Output** met [`2-9-Output.m`](../scripts/power-query/2-9-Output.m).
3. Refresh All en vergelijk met het huidige P8-resultaat. De cijfers moeten
   **identiek** zijn — de wijziging verandert waar de periode vandaan komt, niet
   welke periode het is.

Wat er inhoudelijk verandert:

- `DateTime.LocalNow()` → `ReportingPeriodEnd`. De `-1` verdwijnt daarbij, want
  de parameter ligt *in* de rapportagemaand en `LocalNow()` lag in de maand erna.
  In de 2/9-query wordt `>= Date.Month(LocalNow)` daarom `> Date.Month(RPE)`.
- Elke vlag behoudt `null` als de brondatum leeg is.

> ⚠️ Dat laatste is geen detail. `Date.IsInCurrentYear(null)` geeft `null`, en
> beide TRANSFER-pivots filteren `Transfers (OUT)` op **`(blank)`** — precies die
> null. Een kale `Date.Year(null) = Date.Year(…)` geeft `false` in plaats van
> `null`, waardoor die filter niets meer matcht en de transfercijfers stil
> veranderen. Vandaar de expliciete `if [datum] = null then null else …`.

> ⚠️ De datum wordt als **`=DATE(jjjj;mm;dd)`-formule** weggeschreven, niet als
> tekst. `Date.From()` op tekst hangt af van de locale waarin de refresh draait.

### B. Entity List aansluiten op de Power BI-tab (bevinding 2)

Laat de lookups ongemoeid (24.000+ formules) en voed de oude tab uit de nieuwe.
Zet in `Entity List` op rij 3 en vul door tot rij 360:

```
A3  =IF('Entity List PowerBI'!A4="","",IFERROR(VALUE('Entity List PowerBI'!A4),'Entity List PowerBI'!A4))
B3  ='Entity List PowerBI'!B4
C3  ='Entity List PowerBI'!F4     <-- F, niet C
D3  ='Entity List PowerBI'!G4
```

> ⚠️ **Kolom A moet door `VALUE()`.** De Power BI-tab levert de entity code als
> **tekst** (`'1001'`), terwijl de hard copy hem als **getal** bewaart (`1001`)
> en beide output-tabs opzoeken met een Int64-code. XLOOKUP is typegevoelig, dus
> een kale verwijzing maakt van élke opzoeking `#N/A` en wist de volledige
> PowerHouse-mapping. De `IFERROR` houdt echt niet-numerieke codes als tekst,
> waardoor die zichtbaar onopgelost blijven in plaats van stil te worden omgezet.

### C. Mvt Schedule Details koppelen (haalt een manuele stap weg)

```
B3  ='Movement schedule'!P3    … doorvullen tot B14 (=P14)
E3  ='Movement schedule'!P19   … doorvullen tot E14 (=P30)
E15 =SUM(E3:E14)
```

### D. Marge in het Transfers-blok (bevinding 3)

Vervang de positionele `=C50-C68` door een lookup op PowerHouse-naam, zodat de
berekening blijft kloppen als de pivotrijen verschuiven:

```
C82  =IFERROR(XLOOKUP($B82,$B$49:$B$59,C$49:C$59),0) - IFERROR(XLOOKUP($B82,$B$67:$B$77,C$67:C$77),0)
```

### E. Lease Liability corrigeren voor betalingsfrequentie (bevinding 5)

Vervang in `Mvt Schedule Details` kolom L (rij 19 en verder) `=H19*G19` door:

```
=H19*G19/IFS(I19="Monthly",1,I19="Quarterly",3)
```

Het script zet deze formule vanaf nu zelf bij elke nieuw gedetecteerde building.

## Power Automate flow

Site: `behohr-finance`, bibliotheek `Consolidation  annual statements`.

### 1. Trigger — "Manually trigger a flow"

Begin bewust **handmatig** met één invoerveld:

| Invoer | Type | Voorbeeld |
|---|---|---|
| `PeriodEnd` | Date | `2026-09-30` |

Reden: de close loopt niet elke maand op dezelfde dag, en de preparer weet
wanneer Anaplan klaar is. Een planning of een bestandstrigger kan later, maar
start handmatig zodat de eerste maanden controleerbaar zijn.

### 2. Afgeleide waarden — "Compose"

| Naam | Expressie | Resultaat |
|---|---|---|
| `NewYYYYMM` | `formatDateTime(triggerBody()['date'],'yyyyMM')` | `202609` |
| `NewPeriodCode` | `concat('P',formatDateTime(triggerBody()['date'],'MM'))` | `P09` |
| `NewMonthLabel` | `formatDateTime(triggerBody()['date'],'MMMM yyyy')` | `September 2026` |
| `PrevMonthLabel` | `formatDateTime(addMonths(triggerBody()['date'],-1),'MMMM yyyy')` | `August 2026` |
| `NewFolder` | `concat(formatDateTime(triggerBody()['date'],'yyyy'),'/P',formatDateTime(triggerBody()['date'],'M'),' ',formatDateTime(triggerBody()['date'],'yyyy'),'/IFRS 16')` | `2026/P9 2026/IFRS 16` |
| `PrevFolder` | idem met `addMonths(...,-1)` | `2026/P8 2026/IFRS 16` |

> ⚠️ `formatDateTime` geeft Engelse maandnamen — dat komt overeen met de labels in
> het bestand (`August 2026`). De mapnaam gebruikt `M` zonder voorloopnul
> (`P9 2026`), de bestandsnaam `MM` mét (`202609`). Zie de bestaande mappen.

### 3. Vorige board pack ophalen — "Get files (properties only)"

Op `PrevFolder`, met filter `substringof('Input Board Pack',Name)`. Neem het
eerste resultaat. Faal expliciet als er geen of meer dan één match is — beter
een duidelijke fout dan het verkeerde bestand doorrollen.

### 4. Kopiëren — "Copy file"

Naar `NewFolder`, bestandsnaam
`concat(outputs('NewYYYYMM'),' - IFRS16 - 3 - Input Board Pack.xlsx')`.
Zet "If another file is already there" op **Fail**, zodat een tweede run niet
stilzwijgend werk overschrijft.

Vanaf hier bepaalt de flow zelf de naamgeving, dus de historische
naaminconsistentie speelt geen rol meer.

### 5. Anaplan-exports inlezen — "List rows present in a table" (Excel Online)

Twee keer, op `NewFolder/Anaplan Exports 2.92.10/`:

| Bestand | Tabel | Gaat naar |
|---|---|---|
| `Updated Lease properties (2).xlsx` | de tabel in Sheet 1 | `2.9 Input` |
| `Updated Lease properties (3).xlsx` | de tabel in Sheet 1 | `2.10 Input` |

De exports moeten als Excel-tabel opgemaakt zijn, anders ziet deze actie ze niet.
Zet **Pagination aan** met een limiet boven 25.000 — de 2.10-export heeft
24.617 rijen en haalt anders stilzwijgend maar een deel op.

### 6. Script uitvoeren — "Run script" (Excel Online)

Op het gekopieerde bestand, script `ifrs16-monthly-rollforward`:

| Parameter | Waarde |
|---|---|
| `anaplan29Rows` | output van de 2.9-lijst |
| `anaplan210Rows` | output van de 2.10-lijst |
| `params/periodEndDate` | `PeriodEnd` |
| `params/newMonthLabel` | `NewMonthLabel` |
| `params/previousMonthLabel` | `PrevMonthLabel` |
| `params/newPeriodCode` | `NewPeriodCode` |

Het script zet zelf `ReportingPeriodEnd`, legt kolom G vast vóór het inladen,
ververst en retourneert een reviewrapport als tekst.

### 7. Notificatie — "Send an email (V2)" of Teams-bericht

Body = het reviewrapport uit stap 6, plus een link naar het nieuwe bestand.

### Foutafhandeling

Zet op stap 4 t/m 6 een parallelle "has failed"-tak die mailt met de foutmelding.
Zonder dat faalt de flow stil en staat er een half bijgewerkt board pack in de map.

### Wat het script doet

- Anaplan-data in `Table2.9` en `Table1` schrijven (mét tabel-resize, want de
  queries lezen die tabellen op naam; in blokken van 5.000 rijen).
- Power Queries + alle 5 PivotTables verversen, dan volledig herrekenen.
- Kolom G vastleggen **voor** het inladen (dat is de vorige maand) en na de
  refresh wegschrijven naar kolom O.
- Maandlabels rollen (`F34`/`K34`, `L1`/`M1`, `O1`/`O17`, `P1`/`P17`) en de
  plug-kolom één maand opschuiven (L → M).
- Nieuwe buildings detecteren en voorinvullen in "BUILDINGS - NEW".
- Reviewrapport teruggeven.

### Wat manueel blijft

- **De plug in kolom L** — het reconciliatieverschil tussen de 2.10- en 2.9-cut.
  Het script bewaart de vorige waarde in kolom M, laat L staan, en meldt per
  PowerHouse hoeveel plug er nog nodig is om te sluiten.
- **Beoordeling van de nieuwe contracten** in "BUILDINGS - NEW".
- **De Entity List-refresh.** De connectie is een live MSOLAP-verbinding
  (`Provider=MSOLAP.8; Data Source=pbiazure://api.powerbi.com;
  Integrated Security=ClaimsToken`), geen Power Query. Die vraagt een
  interactieve AAD-token en is vrijwel zeker niet onbemand te verversen.
  Alternatief voor volledige automatisering: Power Automate haalt de lijst op met
  de gevalideerde DAX-query hieronder en het script schrijft de waarden weg.

### Gevalideerde DAX-query (alternatief voor de Entity List-refresh)

Het FDP-model geeft `ArtifactAccessDenied`; het **Treasury-model** DirectQuery't
dezelfde FDP-tabellen en werkt wel.

| Route | artifactId | Resultaat |
|---|---|---|
| `BEHOHR-FDP-PRD-FINANCE` | `5764f6b0-e1e1-426e-9417-790422fa6e62` | ❌ 403 |
| `HOHR Treasury Report - FDP` | `68d39a44-643a-40ce-bfb2-cbf6e2a0e3c6` | ✅ werkt |

```dax
EVALUATE
SELECTCOLUMNS('18. Legal Entity Dimension',
  "Code",        '18. Legal Entity Dimension'[Legal Entity Code],
  "Description", '18. Legal Entity Dimension'[Legal Entity Description],
  "Powerhouse",  '18. Legal Entity Dimension'[LE Powerhouse],
  "Boutique",    '18. Legal Entity Dimension'[LE Boutique])
```

Gevalideerd: 358 entiteiten, 13 PowerHouses, 76 boutiques — exact het aantal
rijen op de `Entity List PowerBI`-tab.

> ⚠️ **Geen `LE Active`-filter.** Voor P&L-cijfers hoort die filter er wel, maar
> dit is een mapping-tabel: een niet-actieve entiteit kan nog lopende of
> historische leases hebben, en filteren laat die contracten zonder PowerHouse.

## Openstaande vragen

1. In P8 staat de Lease Liability van rij 19 nu als `=H19*12` (hardcoded aantal
   kwartalen). Werkt, maar de generieke variant
   `=H19*G19/IFS(I19="Monthly",1,I19="Quarterly",3)` blijft kloppen bij elke
   looptijd en frequentie, en geeft `#N/A` in plaats van een stil verkeerd
   cijfer bij een onbekende frequentie. Overnemen in de template?
