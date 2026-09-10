# IFRS16 Input Board Pack — maandprocedure

Handmatige procedure, zonder Power Automate. Volg de volgorde: stap 3 en 4
moeten vóór de refresh, anders gaan ze stil fout.

## 0. Voorbereiden

Kopieer het **vorige-maandbestand** naar de nieuwe maandmap en hernoem het.

```
2026 / P{nn} 2026 / IFRS 16 / 2026{nn} - IFRS16 - 3 - Input Board Pack.xlsx
```

> Kopieer het bestand met de automatisering erin, niet
> `..._PreAutomatisatie.xlsx`. Dat laatste is het oude bestand van vóór P8 en
> staat er alleen als terugvalpad.

Haal de twee exports uit de Anaplan IFRS16-module (2.9 en 2.10) en zet ze in de
submap `Anaplan Exports` van diezelfde maand.

## 1. Kolom O vullen — vóór alles

Op tabblad **`Movement schedule`**:

| Kopiëren | Plakken **als waarden** in |
|---|---|
| `G3:G15` (buildings) | `O3:O15` |
| `G19:G31` (vehicles) | `O19:O31` |

Kolom G houdt nu nog de vorige maand vast. Na de refresh staat daar de nieuwe
maand en zijn die cijfers weg.

**Plak als waarden** (Ctrl+Shift+V → Waarden). Kolom G bestaat uit XLOOKUP-
formules; plak je die door, dan volgt O gewoon G en is `P = G-O` altijd nul.

Controle: `O15` en `O31` moeten gelijk zijn aan de totalen die je vorige maand
gerapporteerd hebt.

## 2. Plug doorschuiven

Op **`Movement schedule`**, ook als waarden:

| Kopiëren | Plakken in |
|---|---|
| `L3:L14` | `M3:M14` |
| `L19:L30` | `M19:M30` |

De plug in kolom L hoort bij de nieuwe snit en moet opnieuw vastgesteld worden;
M bewaart die van vorige maand. Klopt L straks niet meer, dan zie je dat aan
kolom J (Difference) en aan CHECK — deze stap gaat dus niet stil fout.

## 3. Anaplan-exports inplakken

- `2.10 Input` → tabel **`Table1`**
- `2.9 Input` → tabel **`Table2.9`**

> **De valkuil.** Power Query leest deze tabs als **Excel-tabel**, niet als
> bereik. Heeft de nieuwe export een ander aantal rijen dan vorige maand, dan
> blijft de tabel op de oude omvang staan: extra rijen vallen erbuiten en worden
> genegeerd, of oude rijen blijven erin zitten. Zonder foutmelding.

Twee manieren om dat te voorkomen:

1. Verwijder eerst alle datarijen van de tabel, plak dan vanaf de eerste rij —
   Excel laat de tabel meegroeien; of
2. Plak eroverheen en zet daarna het bereik goed via **Tabelontwerp → Formaat
   wijzigen**.

Controleer achteraf dat het aantal rijen in de tabel gelijk is aan het aantal
rijen in de export.

## 4. Rapportagedatum zetten

Tabblad **`Info`**, cel **`B7`** (heet `ReportingPeriodEnd`): de **laatste dag
van de maand**, bijvoorbeeld `30/09/2026`.

Dit is de enige plek waar de maand staat. Alle koppen, beide Power Queries en de
tabel BUILDINGS - NEW leiden zich hieruit af.

## 5. Refresh All

**Data → Alles vernieuwen.** Zorg dat je in Power BI ingelogd bent, anders
vernieuwt het tabblad `Entity List PowerBI` niet — dat is een rechtstreekse
verbinding met het semantische model en die gaat op jouw token.

Draai de refresh twee keer als de pivots achterlopen: de queries moeten eerst de
outputtabellen opnieuw opbouwen voordat de pivots er iets aan hebben.

## 6. Controleren

| Waar | Moet |
|---|---|
| `Movement schedule` rij 49 (CHECK) | 0 |
| `Movement schedule` kolom J (Difference) | leeg op alle 24 rijen |
| `Movement schedule` G1 / G17 | de nieuwe maand |
| `Mvt Schedule Details` B1 / E1 | `mvt P{nn}` van de nieuwe maand |
| `Mvt Schedule Details` A19 | de nieuwe contracten van deze maand |

Staat er een verschil in kolom J, dan is de plug in kolom L nog niet goed gezet
voor deze snit. Plug hem bij tot J leeg is — maar kijk eerst of het verschil
geen echte datafout is. Forceren maskeert het.

## 7. Beoordelen

Twee dingen die je zelf moet bekijken:

- **De nieuwe contracten** in BUILDINGS - NEW. Zijn dit er plausibel veel voor
  deze maand, en klopt de lease liability?
- **Entiteiten zonder PowerHouse.** Komt er een entiteit bij in MDM die nog niet
  gemapt is, dan vallen zijn contracten uit de pivots. Zichtbaar als een
  entiteit in de inputtabs geen PowerHouse krijgt.

## Wat je níét moet doen

- **Alleen de datum aanpassen en verversen.** De datum bepaalt alleen waar de
  queries de maand afsnijden; de data komt uit de inputtabs. Vervang je de
  export niet, dan snijd je vorige-maandsdata af op de nieuwe maand. CHECK staat
  dan gewoon op 0 — de werkmap klopt met zichzelf, alleen niet met de maand.
- **Het tabblad `Entity List` verwijderen.** Het bevat geen handmatige data meer,
  maar 24.772 XLOOKUPs in de twee output-tabs wijzen ernaar, en het doet de
  typeconversie tussen de tekstcodes van Power BI en de Int64-codes waarmee die
  lookups zoeken. Verbergen mag.
- **Iets naar `Entity List` kopiëren.** Die vult zichzelf uit
  `Entity List PowerBI`, en die haalt zichzelf uit Power BI.

## Waarom de mvt-kolom niet is wat hij lijkt

`Movement schedule` kolom P is **niet** het aantal nieuwe contracten. Het is
`deze maand − wat je vorige maand gerapporteerd hebt`, dus nieuw en beëindigd al
tegen elkaar weggestreept, plus alle herziening van de vorige maand. Zie
[automation-design.md](automation-design.md#wat-kolom-p-wél-en-niet-is).

Wil je weten wat er deze maand bijgekomen is: dat is BUILDINGS - NEW.
