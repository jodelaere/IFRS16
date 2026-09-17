/**
 * IFRS16 Template Setup — one-time Office Script
 *
 * Applies template changes B, C, D, F and H from docs/automation-design.md to
 * the Input Board Pack, and creates the named cell for change A. Run once on
 * the P8 2026 workbook; it then travels forward with every monthly copy.
 *
 * Change A's M-code edit CANNOT be scripted — Office Scripts has no access to
 * Power Query M. Do that by hand afterwards; the script prints the reminder.
 *
 * Safety: every figure this touches is captured before and re-read after, and
 * any change is reported. On P8 2026 nothing should move — these changes
 * restructure how values are derived, not what they are. Checked in advance:
 * change B adds entity codes 1652 and 2006, neither of which has any contract,
 * and no existing code changes PowerHouse; change C matches column P on all 24
 * rows; change H reproduces the six August contracts and the same
 * EUR 347,848.87 total. So any reported difference
 * means something is off — undo (Ctrl+Z works on script edits) and investigate
 * before saving.
 *
 * Safe to re-run: it checks whether each change is already in place.
 */

const SETUP_SHEET_INFO = 'Info'
const SETUP_SHEET_MOVEMENT = 'Movement schedule'
const SETUP_SHEET_DETAILS = 'Mvt Schedule Details'
const SETUP_SHEET_PIVOTS = 'Pivots on 2.10'
const SETUP_SHEET_ENTITY_LIST = 'Entity List'
const SETUP_SHEET_ENTITY_PBI = 'Entity List PowerBI'

const SETUP_PERIOD_NAME = 'ReportingPeriodEnd'
const SETUP_PERIOD_CELL = 'B7'
/** P8 2026 — matches the workbook this is first run against. */
const SETUP_PERIOD_INITIAL = '2026-08-31'

/**
 * Entity List rows 3.. mirror Entity List PowerBI rows 4.., one row lower.
 *
 * MDM held 358 entities when this was written, exactly filling the Power BI
 * table at A3:S361. Sized to 500 so an entity added in MDM still lands here
 * instead of falling off the end, where its contracts would silently lose
 * their PowerHouse. The rows past the source table stay blank: every formula
 * is guarded by IF(source="","",...).
 */
const ENTITY_FIRST_TARGET_ROW = 3
const ENTITY_ROW_COUNT = 500
const ENTITY_SOURCE_OFFSET = 1

/** Transfers block: rows 82-91, fed by the TRANSFER IN and OUT pivots. */
const TRANSFERS_FIRST_ROW = 82
const TRANSFERS_LAST_ROW = 91
const TRANSFER_IN_RANGE = '$B$50:$B$59'
const TRANSFER_OUT_RANGE = '$B$68:$B$77'

const DETAILS_FIRST_PH_ROW = 3
const DETAILS_PH_COUNT = 12
const DETAILS_NEW_FIRST_ROW = 19
/** Rows reserved for the spill: cleared, and formatted so it never lands bare. */
const DETAILS_SPILL_LAST_ROW = 200

/** BUILDINGS - NEW spans A..N since change K split the liability. */
const DETAILS_COLUMN_COUNT = 14

/** Movement schedule: buildings rows 3-14, vehicles 19-30, group total one below. */
const MOVEMENT_BLOCK_FIRST_ROWS = [3, 19]
const MOVEMENT_PH_COUNT = 12

/**
 * Change I columns on Movement schedule.
 *
 * U and V are typed snapshots of last month's New and Terminated, written the
 * same moment as column O. R and S are the deltas derived from them. Lies'
 * version took the deltas straight from the previous month's workbook over an
 * external link; that link carries a hard-coded path to one specific month's
 * file, has to be repointed every close, and breaks the moment a file is
 * renamed — which is exactly what happened when the WorkingVersion was
 * promoted. A snapshot needs no second file.
 */
const MOVEMENT_SPLIT = [
  { source: 'C', delta: 'R', snapshot: 'U', label: 'New contracts', shortLabel: 'New', seed: 'newContracts' },
  { source: 'E', delta: 'S', snapshot: 'V', label: 'Terminated contracts', shortLabel: 'Ended', seed: 'terminated' },
  { source: 'F', delta: 'T', snapshot: 'W', label: 'Transfers between PHs', shortLabel: 'Transf', seed: 'transfers' },
]

/**
 * Columns whose formatting the new ones borrow. C is an ordinary derived
 * column; L is the plug, which like the snapshots is typed by hand each month
 * and is set in italic to say so.
 */
const MOVEMENT_RESULT_TEMPLATE_COLUMN = 'C'
const MOVEMENT_INPUT_TEMPLATE_COLUMN = 'L'
const MOVEMENT_SPLIT_COLUMN_WIDTH = 62
/** The existing count column on Mvt Schedule Details. */
const DETAILS_RESULT_TEMPLATE_COLUMN = 'B'

/**
 * P7 2026 New and Terminated per PowerHouse, in Movement schedule row order,
 * group total last. Read from the July workbook
 * (P7 2026/IFRS 16/#2026.07 - Board Slides (Powerquery).xlsx), columns C and E.
 *
 * These seed the snapshot columns once, because August is the first close that
 * has them and no earlier run photographed July. From P9 onwards the monthly
 * step fills them and this table is never used again — the seed only writes
 * into cells that are still empty, so re-running setup in a later month cannot
 * drag July's figures back in.
 */
const PRIOR_MONTH_SEED: { [block: string]: { [column: string]: number[] } } = {
  buildings: {
    newContracts: [2, 8, 6, 1, 0, 136, 8, 0, 4, 10, 58, 0, 233],
    terminated: [-4, -13, -14, -1, -1, -49, -13, 0, -1, -6, -88, 0, -190],
    transfers: [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
  },
  vehicles: {
    newContracts: [18, 30, 394, 7, 0, 374, 134, 4, 29, 18, 106, 58, 1172],
    terminated: [-45, -45, -463, -36, -10, -413, -473, -12, -163, -24, -62, -121, -1867],
    transfers: [-1, -2, 0, 8, -1, 0, 0, 3, -7, 0, 0, 0, 0],
  },
}

/**
 * Change M: the snapshot sheet behind the terminated/reinstated pair.
 *
 * A2 holds a live list of the contracts that count as terminated right now —
 * the same rule the Terminated pivot uses. D and E hold that same list as it
 * stood LAST month, as typed values, photographed before the new export goes
 * in. Everything else is the difference between the two.
 *
 * The rule was validated against both real exports: 187 at 31 July, matching
 * what the July pack reported, and 193 at 31 August, matching the pivot.
 */
const SNAPSHOT_SHEET = 'Snapshot'
const SNAPSHOT_ROW_COUNT = 800
/** Rows 1-8 explain what the sheet is; the lists start below that. */
const SNAPSHOT_FIRST_DATA_ROW = 9
const SNAPSHOT_NOTE = [
  'Werkblad van de automatisering. Niet met de hand aanpassen, op één stap na (zie onder).',
  '',
  'Kolom A en B: de contracten die op dit moment als stopgezet tellen. Dit is een formule over Table1 — dezelfde',
  'regel als de Terminated-pivot: Land and buildings, einddatum in het rapportagejaar t/m de rapportagemaand,',
  'geen transfer-out-datum.',
  'Kolom D en E: exact diezelfde lijst zoals hij vórige maand was, als getypte waarden.',
  'Het verschil tussen de twee vult BUILDINGS - TERMINATED en BUILDINGS - REINSTATED op Mvt Schedule Details.',
  'ÉÉN HANDELING PER MAAND: kopieer A en B naar D en E als waarden, vóórdat je de nieuwe Anaplan-export inplakt.',
]

/** BUILDINGS - REINSTATED, past the terminated block. */
const DETAILS_BACK_FIRST_COLUMN_INDEX = 27 // AB
const DETAILS_BACK_HEADERS = [
  'BUILDINGS - REINSTATED',
  'Reasonably certain end date last month',
  'Entity',
  'Lease description',
  'Reasonably certain end date now',
]

/**
 * P7 2026 terminated set — the 187 Land and buildings contracts that counted as
 * terminated at 31 July: end date in 2026 up to and including July, no
 * transfer-out date.
 *
 * Extracted from the real July Anaplan export (Exports_Juli2026.xlsx, sheet
 * 2.10). Not from July2026_2.10.xlsx in the P8 Anaplan Exports folder — that
 * one was pulled during the August close and reports 170, seventeen short of
 * the 187 the July pack actually reported.
 *
 * This seeds the snapshot once. August is the first close that compares against
 * a previous month, and nothing photographed July at the time. From P9 onwards
 * the monthly step writes it and this table is never read again.
 *
 * Verified: against the August export this reproduces 17 reinstated and 23
 * newly terminated contracts, which reconciles to Movement schedule column S on
 * all twelve PowerHouses, and matches the hand analysis contract for contract
 * (ABY 1611__5 and 1611__6, three SOLCOM, twelve TimePartner).
 */
const PRIOR_TERMINATED_SEED: string[][] = [
  ['1101__BUILDING2', '2026-06-30'],
  ['1201__P226.001', '2026-03-01'],
  ['1202__P105.001', '2026-01-31'],
  ['1202__P152.001', '2026-01-01'],
  ['1202__P173.001', '2026-02-28'],
  ['1202__P185.001', '2026-02-28'],
  ['1202__P204.001', '2026-01-01'],
  ['1202__P241.001', '2026-04-30'],
  ['1202__P244.001', '2026-06-30'],
  ['1202__P255.001.2', '2026-05-20'],
  ['1202__P280.001', '2026-01-01'],
  ['1202__P307.001', '2026-01-31'],
  ['1202__P341.001', '2026-04-30'],
  ['1202__P360.001', '2026-01-31'],
  ['1402__CONB013', '2026-02-28'],
  ['1501__B047_313', '2026-02-28'],
  ['1501__B049_313', '2026-05-31'],
  ['1501__CO_COVB054_313', '2026-01-01'],
  ['1501__CO_COVB055_313', '2026-03-31'],
  ['1501__CO_COVHOUSING_774', '2026-01-14'],
  ['1501__CO_COVHOUSING_775', '2026-01-31'],
  ['1501__CO_COVHOUSING_776', '2026-02-08'],
  ['1501__CO_COVHOUSING_777', '2026-02-14'],
  ['1501__CO_COVHOUSING_783', '2026-03-31'],
  ['1501__CO_COVHOUSING_784', '2026-03-31'],
  ['1501__CO_COVHOUSING_785', '2026-05-14'],
  ['1501__CO_COVHOUSING_790', '2026-05-31'],
  ['1501__CO_COVHOUSING_800', '2026-06-30'],
  ['1501__CO_COVHOUSING_801', '2026-06-30'],
  ['1501__CO_COVHOUSING_820', '2026-03-16'],
  ['1501__CO_COVHOUSING_823', '2026-01-31'],
  ['1501__CO_COVHOUSING_825', '2026-06-30'],
  ['1501__HOUSING_001', '2026-03-31'],
  ['1501__HOUSING_588', '2026-05-01'],
  ['1501__HOUSING_707', '2026-02-28'],
  ['1501__HOUSING_708', '2026-03-31'],
  ['1501__HOUSING_709', '2026-02-28'],
  ['1501__HOUSING_710', '2026-03-31'],
  ['1501__HOUSING_711', '2026-03-31'],
  ['1501__HOUSING_745', '2026-04-30'],
  ['1501__HOUSING_746', '2026-06-30'],
  ['1501__HOUSING_751', '2026-06-30'],
  ['1501__HOUSING_830', '2026-01-31'],
  ['1501__HOUSING_836', '2026-03-31'],
  ['1501__HOUSING_843', '2026-03-16'],
  ['1501__HOUSING_847', '2026-02-28'],
  ['1501__HOUSING_848', '2026-01-31'],
  ['1501__HOUSING_849', '2026-01-31'],
  ['1501__HOUSING_851', '2026-03-02'],
  ['1501__HOUSING_858', '2026-05-31'],
  ['1501__HOUSING_860', '2026-03-31'],
  ['1501__HOUSING_861', '2026-01-26'],
  ['1501__HOUSING_876', '2026-06-30'],
  ['1501__HOUSING_888', '2026-05-31'],
  ['1501__HOUSING_912', '2026-07-10'],
  ['1501__HOUSING_913', '2026-07-10'],
  ['1501__HOUSING_914', '2026-07-10'],
  ['1501__HOUSING_915', '2026-07-10'],
  ['1501__HOUSING_917', '2026-07-10'],
  ['1501__HOUSING_918', '2026-07-10'],
  ['1552__PRO_NL_BRAN0002', '2026-03-31'],
  ['1552__PRO_NL_BRAN0003', '2026-01-31'],
  ['1552__PRO_NL_BRAN0018', '2026-03-31'],
  ['1602__3', '2026-03-31'],
  ['1610__11', '2026-02-01'],
  ['1611__5', '2026-07-31'],
  ['1611__6', '2026-07-31'],
  ['1704__TP206', '2026-05-31'],
  ['1704__TP207', '2026-01-31'],
  ['1704__TP216.2', '2026-03-31'],
  ['1704__TP220', '2026-04-30'],
  ['1704__TP221', '2026-04-30'],
  ['1704__TP222', '2026-01-31'],
  ['1704__TP240', '2026-07-31'],
  ['1704__TP241', '2026-07-31'],
  ['1704__TP243', '2026-07-31'],
  ['1704__TP254', '2026-07-31'],
  ['1704__TP255', '2026-05-31'],
  ['1704__TP256', '2026-07-31'],
  ['1704__TP264', '2026-04-30'],
  ['1704__TP266', '2026-07-31'],
  ['1704__TP267', '2026-07-31'],
  ['1704__TP268', '2026-04-30'],
  ['1704__TP269', '2026-04-14'],
  ['1704__TP282', '2026-03-31'],
  ['1704__TP284', '2026-04-30'],
  ['1704__TP288', '2026-05-31'],
  ['1704__TP289', '2026-05-31'],
  ['1704__TP290', '2026-06-30'],
  ['1704__TP295', '2026-07-31'],
  ['1704__TP298', '2026-05-31'],
  ['1704__TP300', '2026-05-31'],
  ['1704__TP309', '2026-06-30'],
  ['1704__TP317', '2026-07-31'],
  ['1704__TP325', '2026-06-30'],
  ['1704__TP531', '2026-07-31'],
  ['1704__TP537', '2026-07-31'],
  ['1704__TP544', '2026-07-31'],
  ['1704__TP551', '2026-06-30'],
  ['1704__TP553', '2026-06-30'],
  ['1704__TP554', '2026-06-30'],
  ['1704__TP556', '2026-06-30'],
  ['1704__TP583', '2026-01-31'],
  ['1704__TP584', '2026-01-31'],
  ['1704__TP585', '2026-03-31'],
  ['1704__TP586', '2026-01-31'],
  ['1704__TP587', '2026-01-31'],
  ['1704__TP588', '2026-01-31'],
  ['1704__TP589', '2026-05-31'],
  ['1704__TP619', '2026-06-14'],
  ['1704__TP648', '2026-07-15'],
  ['1704__ZQZAQ-2005', '2026-06-30'],
  ['1710__TP105', '2026-05-31'],
  ['1710__TP106', '2026-07-31'],
  ['1710__TP118', '2026-06-30'],
  ['1710__TP181', '2026-02-28'],
  ['1710__TP187', '2026-04-30'],
  ['1710__TP199', '2026-07-31'],
  ['1710__TP211', '2026-06-30'],
  ['1710__TP25', '2026-07-31'],
  ['1710__TP29', '2026-01-31'],
  ['1710__TP304', '2026-02-28'],
  ['1710__TP306', '2026-07-31'],
  ['1710__TP310', '2026-02-28'],
  ['1710__TP323', '2026-07-31'],
  ['1710__TP35', '2026-03-31'],
  ['1710__TP359', '2026-02-28'],
  ['1710__TP366', '2026-06-30'],
  ['1710__TP367', '2026-06-30'],
  ['1710__TP374', '2026-03-31'],
  ['1710__TP388', '2026-07-31'],
  ['1710__TP389', '2026-07-31'],
  ['1710__TP393', '2026-07-31'],
  ['1710__TP423', '2026-04-30'],
  ['1710__TP424', '2026-04-30'],
  ['1710__TP432', '2026-06-30'],
  ['1710__TP433', '2026-06-30'],
  ['1710__TP434', '2026-02-28'],
  ['1710__TP465', '2026-02-28'],
  ['1710__TP467', '2026-02-28'],
  ['1710__TP488', '2026-01-31'],
  ['1710__TP489', '2026-06-23'],
  ['1710__TP5', '2026-07-31'],
  ['1710__TP545', '2026-06-23'],
  ['1710__TP557', '2026-07-31'],
  ['1710__TP561', '2026-06-30'],
  ['1710__TP570', '2026-03-31'],
  ['1710__TP641', '2026-06-30'],
  ['1710__TP644', '2026-04-30'],
  ['1710__TP65', '2026-05-31'],
  ['1710__TP71', '2026-06-30'],
  ['1710__ZQBZ-256.2', '2026-06-23'],
  ['1710__ZQZAQ-2009', '2026-02-28'],
  ['1710__ZQZAQ-2012', '2026-02-28'],
  ['1824__TRAB008', '2026-05-31'],
  ['1903__UtrechtII', '2026-03-31'],
  ['1903__UtrechtIIICohedron', '2026-03-31'],
  ['1905__Dellaertlaan-Beverwijk', '2026-01-01'],
  ['1914__Utrecht', '2026-03-31'],
  ['1920__ZaandamII', '2026-03-31'],
  ['1921__Houten_Old_1/3/2025', '2026-03-31'],
  ['1922__Arnhem', '2026-03-31'],
  ['1929__Sittard WnH laan 19 BIB Zuid', '2026-03-31'],
  ['1930__Emmen 2A17', '2026-03-31'],
  ['1933__Meerkollaan-Eindhoven', '2026-03-31'],
  ['1933__Sittard WnH laan 19 Reeling', '2026-03-31'],
  ['1936__Eindhoven-Bogert', '2026-03-31'],
  ['1938__Utrecht', '2026-03-31'],
  ['1956__Amsterdamsestraatweg-Baarn', '2026-03-31'],
  ['2003__HAM-Buro2', '2026-07-31'],
  ['2003__HAM-Stellplatz2', '2026-07-31'],
  ['2003__HAM-Stellplatz3', '2026-07-31'],
  ['2003__HAM-Stellplatz4', '2026-07-31'],
  ['2003__HAM-Stellplatz5', '2026-07-31'],
  ['2003__HAM-Stellplatze', '2026-07-31'],
  ['2202__LC-Neuss', '2026-07-31'],
  ['2205__3A37-7D2F', '2026-02-28'],
  ['2205__7699-8E60_1', '2026-02-28'],
  ['2205__7D43-31CD', '2026-01-31'],
  ['2205__ABE6-476A', '2026-01-31'],
  ['2205__F55C-D902', '2026-03-31'],
  ['2210__5E31-F9F1_077', '2026-06-30'],
  ['2221__AVBZ-1', '2026-03-31'],
  ['2221__AVBZ-108', '2026-01-31'],
  ['2221__AVBZ-25', '2026-03-31'],
  ['2221__AVBZ-37', '2026-04-30'],
  ['2221__BZ-98', '2026-05-31'],
]

/** Change J columns on Mvt Schedule Details: buildings G/H, vehicles J/K. */
/**
 * Change L: BUILDINGS - TERMINATED, beside the NEW table rather than below it.
 *
 * Below would have to clear the 182 rows the NEW spill reserves, and the moment
 * a month produced more contracts than that reservation the two would collide
 * with #SPILL!. Side by side, neither can ever reach the other.
 *
 * Eleven columns, the same eleven the NEW table opens with — a terminated
 * contract has no lease liability to add.
 */
const DETAILS_OUT_FIRST_COLUMN_INDEX = 15 // P
const DETAILS_OUT_COLUMN_COUNT = 11 // P..Z
const DETAILS_OUT_HEADERS = [
  'BUILDINGS - TERMINATED',
  'Entity',
  'Lease description',
  'Lease commencement date',
  'Reasonably certain end date selection',
  'Reasonably certain end date',
  'Lease duration',
  'Fixed payment',
  'Payment frequency',
  'Asset category',
  'Leased capacity',
]

const DETAILS_SPLIT_BLOCKS = [
  { columns: ['G', 'H', 'I'], movementFirstRow: 3 },
  { columns: ['J', 'K', 'L'], movementFirstRow: 19 },
]

function main(workbook: ExcelScript.Workbook): string {
  const log: string[] = ['IFRS16 template setup']

  const before = captureCheckFigures(workbook, 'L25')

  log.push(applyReportingPeriodCell(workbook))
  log.push(applyEntityListLink(workbook))
  log.push(applyDetailsCountLink(workbook))
  log.push(applyTransfersLookup(workbook))
  log.push(applyPeriodLabels(workbook))
  log.push(applyNewBuildingsSpill(workbook))
  log.push(applyMovementSplit(workbook))
  log.push(applyDetailsSplit(workbook))
  log.push(applySnapshotSheet(workbook))
  log.push(applyTerminatedBuildings(workbook))
  log.push(applyReinstatedBuildings(workbook))

  workbook.getApplication().calculate(ExcelScript.CalculationType.fullRebuild)

  log.push('')
  log.push(compareCheckFigures(workbook, before))
  log.push('')
  log.push('STILL TO DO BY HAND — template change A: replace DateTime.LocalNow() in')
  log.push('both Power Queries with the ReportingPeriodEnd parameter. See the table')
  log.push('in docs/automation-design.md. Until then the reporting month still comes')
  log.push('from the system clock and the roll-forward script will refuse to run.')

  return log.join('\n')
}

/**
 * Figures that must not move: the CHECK row, both group totals, and the
 * new-lease total. That total is read where it lives at the time — L25 before
 * change H moves it, L17 after — so the comparison stays like-for-like.
 */
function captureCheckFigures(workbook: ExcelScript.Workbook, liabilityCell: string): (number | string)[] {
  const movement = workbook.getWorksheet(SETUP_SHEET_MOVEMENT)
  const details = workbook.getWorksheet(SETUP_SHEET_DETAILS)
  return [
    readFigure(movement.getRange('F49')),
    readFigure(movement.getRange('K49')),
    readFigure(movement.getRange('G15')),
    readFigure(movement.getRange('G31')),
    readFigure(movement.getRange('F15')),
    readFigure(movement.getRange('F31')),
    readFigure(details.getRange('B15')),
    readFigure(details.getRange('E15')),
    readFigure(details.getRange(liabilityCell)),
  ]
}

/**
 * Returns the number, or the error text when the cell holds one. Reading these
 * through Number() would turn #NAME?/#SPILL! into NaN, and every NaN
 * comparison is false — so a broken formula would have passed verification
 * silently, which is the one outcome this check exists to prevent.
 */
function readFigure(range: ExcelScript.Range): number | string {
  const value = range.getValue()
  if (typeof value === 'string' && value.indexOf('#') === 0) return value
  const asNumber = Number(value)
  return isNaN(asNumber) ? String(value) : asNumber
}

function compareCheckFigures(workbook: ExcelScript.Workbook, before: (number | string)[]): string {
  const labels = [
    'CHECK buildings (F49)', 'CHECK vehicles (K49)',
    'Total buildings (G15)', 'Total vehicles (G31)',
    'Transfers buildings (F15)', 'Transfers vehicles (F31)',
    'Mvt buildings (B15)', 'Mvt vehicles (E15)',
    'New lease liability total',
  ]
  const after = captureCheckFigures(workbook, 'L17')
  const moved: string[] = []
  after.forEach((value, index) => {
    const wasNumber = typeof before[index] === 'number'
    const isNumber = typeof value === 'number'
    const changed = !wasNumber || !isNumber
      ? String(before[index]) !== String(value)
      : Math.abs((value as number) - (before[index] as number)) > 0.005
    if (changed) moved.push(`${labels[index]}: ${before[index]} -> ${value}`)
  })
  return moved.length === 0
    ? 'VERIFICATION OK — no reported figure changed.'
    : `FIGURES CHANGED — review before saving:\n  ${moved.join('\n  ')}`
}

/** Change A (the scriptable half): the named cell both Power Queries will read. */
function applyReportingPeriodCell(workbook: ExcelScript.Workbook): string {
  if (workbook.getNamedItem(SETUP_PERIOD_NAME)) {
    return `A: named cell "${SETUP_PERIOD_NAME}" already exists — left alone.`
  }
  const info = workbook.getWorksheet(SETUP_SHEET_INFO)
  info.getRange('A7').setValue('Reporting period end (drives both Power Queries)')
  const cell = info.getRange(SETUP_PERIOD_CELL)
  // Written as a DATE() formula, not a string: Date.From() in the queries would
  // otherwise parse text under whatever locale the refresh runs in.
  const parts = SETUP_PERIOD_INITIAL.split('-')
  cell.setFormula(`=DATE(${parts[0]},${Number(parts[1])},${Number(parts[2])})`)
  cell.setNumberFormat('dd/mm/yyyy')
  workbook.addNamedItem(SETUP_PERIOD_NAME, `=${SETUP_SHEET_INFO}!$${SETUP_PERIOD_CELL[0]}$${SETUP_PERIOD_CELL.substring(1)}`, 'Reporting period end read by the 2.9 and 2.10 Power Queries')
  return `A: created ${SETUP_PERIOD_NAME} at ${SETUP_SHEET_INFO}!${SETUP_PERIOD_CELL} = ${SETUP_PERIOD_INITIAL}.`
}

/**
 * Change B: feed the hard-copy Entity List from the Power BI tab, leaving the
 * 24k+ XLOOKUP formulas in the output tabs untouched. PowerHouse is column F on
 * the source tab — column C there is LE Country Long.
 *
 * The entity code needs VALUE(). The Power BI tab returns codes as text
 * ('1001') while the hard copy holds them as numbers (1001) and both output
 * tabs look them up with an Int64 entity code. XLOOKUP is type-strict, so a
 * plain reference would turn every lookup into #N/A and wipe out the entire
 * PowerHouse mapping. IFERROR keeps genuinely non-numeric codes as text, which
 * correctly leaves them unmatched and visible rather than silently coerced.
 */
function applyEntityListLink(workbook: ExcelScript.Workbook): string {
  const sheet = workbook.getWorksheet(SETUP_SHEET_ENTITY_LIST)
  const sourceColumns = ['A', 'B', 'F', 'G']

  for (let index = 0; index < ENTITY_ROW_COUNT; index++) {
    const targetRow = ENTITY_FIRST_TARGET_ROW + index
    const sourceRow = targetRow + ENTITY_SOURCE_OFFSET
    sourceColumns.forEach((sourceColumn, columnIndex) => {
      const reference = `'${SETUP_SHEET_ENTITY_PBI}'!${sourceColumn}${sourceRow}`
      const value = columnIndex === 0 ? `IFERROR(VALUE(${reference}),${reference})` : reference
      sheet
        .getRangeByIndexes(targetRow - 1, columnIndex, 1, 1)
        .setFormula(`=IF(${reference}="","",${value})`)
    })
  }
  return `B: Entity List rows ${ENTITY_FIRST_TARGET_ROW}-${ENTITY_FIRST_TARGET_ROW + ENTITY_ROW_COUNT - 1} now derive from ${SETUP_SHEET_ENTITY_PBI} (PowerHouse from column F, code converted back to a number).`
}

/** Change C: the count block equals Movement schedule column P — link it instead of retyping. */
function applyDetailsCountLink(workbook: ExcelScript.Workbook): string {
  const sheet = workbook.getWorksheet(SETUP_SHEET_DETAILS)
  for (let index = 0; index < DETAILS_PH_COUNT; index++) {
    const row = DETAILS_FIRST_PH_ROW + index
    sheet.getRange(`B${row}`).setFormula(`='${SETUP_SHEET_MOVEMENT}'!P${3 + index}`)
    sheet.getRange(`E${row}`).setFormula(`='${SETUP_SHEET_MOVEMENT}'!P${19 + index}`)
  }
  // Both totals, not just E15: B15 had been overwritten with a typed value at
  // some point, so it would have gone stale as soon as column B moved.
  sheet.getRange('B15').setFormula('=SUM(B3:B14)')
  sheet.getRange('E15').setFormula('=SUM(E3:E14)')
  return 'C: Mvt Schedule Details count block linked to Movement schedule column P; B15 and E15 are now SUMs.'
}

/**
 * Change D: the transfers block paired pivot rows positionally (=C50-C68), which
 * silently mismatches PowerHouses as soon as the two pivots list different ones.
 * Look them up by name instead.
 */
function applyTransfersLookup(workbook: ExcelScript.Workbook): string {
  const sheet = workbook.getWorksheet(SETUP_SHEET_PIVOTS)
  for (let row = TRANSFERS_FIRST_ROW; row <= TRANSFERS_LAST_ROW; row++) {
    for (const column of ['C', 'D', 'E']) {
      sheet.getRange(`${column}${row}`).setFormula(
        `=IFERROR(XLOOKUP($B${row},${TRANSFER_IN_RANGE},${column}$50:${column}$59),0)` +
        `-IFERROR(XLOOKUP($B${row},${TRANSFER_OUT_RANGE},${column}$68:${column}$77),0)`
      )
    }
  }
  return `D: transfers block rows ${TRANSFERS_FIRST_ROW}-${TRANSFERS_LAST_ROW} now match PowerHouses by name instead of by row position.`
}

/**
 * Change F: derive every period label from ReportingPeriodEnd.
 *
 * K34 was a typed label, so changing the parameter moved the figures but left
 * the heading on the old month — July data under an "August 2026" header, which
 * is exactly the kind of mismatch nobody catches in a board pack. Deriving the
 * labels makes that impossible, and means the monthly roll needs no label edits
 * at all: one cell drives the whole workbook.
 *
 * [$-en-US] forces English month names regardless of the Excel display language,
 * matching the labels already in the file.
 */
function applyPeriodLabels(workbook: ExcelScript.Workbook): string {
  const movement = workbook.getWorksheet(SETUP_SHEET_MOVEMENT)
  const details = workbook.getWorksheet(SETUP_SHEET_DETAILS)

  const currentMonth = `TEXT(${SETUP_PERIOD_NAME},"[$-en-US]mmmm yyyy")`
  const priorMonth = `TEXT(EDATE(${SETUP_PERIOD_NAME},-1),"[$-en-US]mmmm yyyy")`
  const periodCode = `"mvt P"&TEXT(MONTH(${SETUP_PERIOD_NAME}),"00")`
  const plugCurrent = `"Plug "&TEXT(MONTH(${SETUP_PERIOD_NAME}),"00")&" "&YEAR(${SETUP_PERIOD_NAME})`
  const plugPrior = `"Plug "&TEXT(MONTH(EDATE(${SETUP_PERIOD_NAME},-1)),"00")&" "&YEAR(EDATE(${SETUP_PERIOD_NAME},-1))`

  // G1 and G17 already read =K34, so they follow automatically.
  movement.getRange('K34').setFormula(`=${currentMonth}`)
  movement.getRange('F34').setFormula(`=${currentMonth}`)
  movement.getRange('L1').setFormula(`=${plugCurrent}`)
  movement.getRange('M1').setFormula(`=${plugPrior}`)
  for (const row of [1, 17]) {
    movement.getRange(`O${row}`).setFormula(`=${priorMonth}`)
    movement.getRange(`P${row}`).setFormula(`=${periodCode}`)
  }
  details.getRange('B1').setFormula(`=${periodCode}`)
  details.getRange('E1').setFormula(`=${periodCode}`)

  return 'F: period labels (K34, F34, L1, M1, O1/O17, P1/P17, details B1/E1) now derive from ReportingPeriodEnd.'
}

/**
 * Change H: derive BUILDINGS - NEW from the Anaplan data instead of writing it.
 *
 * This was the last part of the workbook that did not follow the reporting
 * parameter — set the period to July and the August contracts stayed put under
 * July figures. One spill formula fixes that permanently, and supersedes the
 * old per-row lease liability formulas (the liability is now column 12 of the
 * spill) and the mismatch warning that would otherwise have guarded it.
 *
 * Source is Table1 (2.10 Input), not 2_10 Output: the query drops Fixed
 * payment, Lease duration, Leased capacity and Payment frequency, which this
 * table needs. Column numbers below are positions in Table1.
 *
 * The total moves to L17, above the header. A spill grows and shrinks, so
 * anything directly beneath it would block it with #SPILL!.
 */
function applyNewBuildingsSpill(workbook: ExcelScript.Workbook): string {
  const sheet = workbook.getWorksheet(SETUP_SHEET_DETAILS)

  // The old static rows and total must go first, or the spill has nowhere to
  // land. Row 19 keeps its formatting: it is the template every spilled row is
  // styled from.
  const spillRowCount = DETAILS_SPILL_LAST_ROW - DETAILS_NEW_FIRST_ROW + 1
  sheet
    .getRangeByIndexes(DETAILS_NEW_FIRST_ROW - 1, 0, spillRowCount, DETAILS_COLUMN_COUNT)
    .clear(ExcelScript.ClearApplyTo.contents)

  // Change K: the liability also splits into non-current and current.
  //
  // Lies added that split by hand per row, and it went wrong twice. The
  // quarterly ClickCare contract was not divided by its payment frequency, so
  // it came out at three times the real figure; and from the second row down
  // the current column held the WHOLE liability while the non-current column
  // still added a slice on top, double-counting every row. Derived here
  // instead, from the same duration and frequency the total already uses, so
  // non-current + current = total by construction.
  //
  // IF rather than MIN/MAX: those collapse an array to one value, and these
  // operands are whole spilled columns.
  sheet.getRange(`A${DETAILS_NEW_FIRST_ROW}`).setFormula(
    '=LET(t,Table1,' +
    'pay,INDEX(t,,15),dur,INDEX(t,,14),freq,INDEX(t,,16),' +
    'div,IFS(freq="Monthly",1,freq="Quarterly",3),' +
    'liab,pay*dur/div,' +
    'cur,pay*IF(dur>12,12,dur)/div,' +
    'noncur,pay*IF(dur>12,dur-12,0)/div,' +
    `keep,(INDEX(t,,26)="Land and buildings")*(TEXT(INDEX(t,,6),"yyyymm")=TEXT(${SETUP_PERIOD_NAME},"yyyymm")),` +
    'FILTER(HSTACK(CHOOSECOLS(t,1,2,3,6,10,11,14,15,16,26,27),liab,noncur,cur),keep,""))'
  )

  // The two new columns inherit the liability column's look — accounting
  // format, fill, borders — before the styling step tiles row 19 downwards.
  sheet
    .getRange(`M${DETAILS_NEW_FIRST_ROW}:N${DETAILS_NEW_FIRST_ROW}`)
    .copyFrom(sheet.getRange(`L${DETAILS_NEW_FIRST_ROW}`), ExcelScript.RangeCopyType.formats)
  const headerRow = DETAILS_NEW_FIRST_ROW - 1
  sheet
    .getRange(`M${headerRow}:N${headerRow}`)
    .copyFrom(sheet.getRange(`L${headerRow}`), ExcelScript.RangeCopyType.formats)
  sheet.getRange(`M${headerRow}`).setValue('Lease Liability Non-Current')
  sheet.getRange(`N${headerRow}`).setValue('Lease Liability Current')

  // The totals sit above the table: a spill grows and shrinks, so anything
  // directly beneath it would block it with #SPILL!.
  const totalLabel = sheet.getRange('K17')
  totalLabel.setValue('Total')
  totalLabel.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right)
  totalLabel.getFormat().getFont().setBold(true)

  // Summing the spill by column index rather than a fixed range is what stops
  // the off-by-one that left the last contract out of Lies' total.
  ;[
    { cell: 'L17', column: 12 },
    { cell: 'M17', column: 13 },
    { cell: 'N17', column: 14 },
  ].forEach((entry) => {
    const total = sheet.getRange(entry.cell)
    total.setFormula(`=SUM(CHOOSECOLS(A${DETAILS_NEW_FIRST_ROW}#,${entry.column}))`)
    total.setNumberFormat('#,##0.00')
    total.getFormat().getFont().setBold(true)
    total.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeTop).setStyle(ExcelScript.BorderLineStyle.continuous)
    total.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setStyle(ExcelScript.BorderLineStyle.double)
  })

  const styled = styleSpillBlock(workbook, 0, DETAILS_COLUMN_COUNT)
  return (
    'H+K: BUILDINGS - NEW spills from Table1 for the reporting month, liability split into ' +
    `non-current and current; totals in L17/M17/N17, ${styled} row(s) styled.`
  )
}

/**
 * Change I: split the monthly movement into new versus terminated contracts.
 *
 * Column P (=G-O) is the net move and nets the two against each other, so a
 * PowerHouse that took on four contracts and lost four reads as nil. Lies
 * separated them on the buildings block by subtracting the previous month's
 * workbook over an external link. Two problems with that: the link names one
 * specific month's file, so it has to be repointed every close and dies when a
 * file is renamed; and it only covered buildings.
 *
 * Same device as column O instead. U and V hold last month's New and
 * Terminated as typed values, written at the same moment as O, and the deltas
 * derive from them. No second workbook, and it extends to vehicles for free.
 *
 * R + S = P is then a free check: the two halves must add back to the net move.
 */
function applyMovementSplit(workbook: ExcelScript.Workbook): string {
  const sheet = workbook.getWorksheet(SETUP_SHEET_MOVEMENT)
  const priorShort = `TEXT(MONTH(EDATE(${SETUP_PERIOD_NAME},-1)),"00")&" "&YEAR(EDATE(${SETUP_PERIOD_NAME},-1))`

  for (const firstRow of MOVEMENT_BLOCK_FIRST_ROWS) {
    const headerRow = firstRow - 2
    const subHeaderRow = firstRow - 1
    const totalRow = firstRow + MOVEMENT_PH_COUNT
    const lastPowerHouseRow = totalRow - 1

    for (const part of MOVEMENT_SPLIT) {
      // A derived column looks like every other derived column, and a snapshot
      // looks like the plug columns it behaves like. Both are copied off cells
      // that already exist rather than restyled from scratch — the sheet's own
      // conventions (magenta header, banded total, italic for anything typed by
      // hand) then hold without having to be restated here.
      styleLikeColumn(sheet, MOVEMENT_RESULT_TEMPLATE_COLUMN, part.delta, headerRow, subHeaderRow, firstRow, lastPowerHouseRow, totalRow)
      styleLikeColumn(sheet, MOVEMENT_INPUT_TEMPLATE_COLUMN, part.snapshot, headerRow, subHeaderRow, firstRow, lastPowerHouseRow, totalRow)
      // The snapshot total still sits in the banded row, just italic with it.
      sheet
        .getRange(`${part.snapshot}${totalRow}`)
        .copyFrom(sheet.getRange(`${MOVEMENT_RESULT_TEMPLATE_COLUMN}${totalRow}`), ExcelScript.RangeCopyType.formats)
      sheet.getRange(`${part.snapshot}${totalRow}`).getFormat().getFont().setItalic(true)

      sheet.getRange(`${part.delta}${headerRow}`).setValue(part.label)
      // Short, like "Plug 07 2026" — the long form wrapped into three lines.
      sheet
        .getRange(`${part.snapshot}${headerRow}`)
        .setFormula(`="${part.shortLabel} "&${priorShort}`)

      // Through the group total row, so the totals split too.
      for (let row = firstRow; row <= totalRow; row++) {
        sheet
          .getRange(`${part.delta}${row}`)
          .setFormula(`=${part.source}${row}-${part.snapshot}${row}`)
      }
    }
  }

  // Wrapped headers need a width, or "Terminated contracts" spills across its
  // neighbours the way it did before.
  sheet.getRange('R:W').getFormat().setColumnWidth(MOVEMENT_SPLIT_COLUMN_WIDTH)

  const seeded = seedPriorMonthSplit(sheet, workbook)

  return (
    'I: Movement schedule R/S/T split the move into new, terminated and transfers, from snapshots ' +
    `in U/V/W — both blocks, styled off columns C and L. ${seeded}`
  )
}

/**
 * Give a column the look of an existing one: header, sub-header, data rows and
 * the group total row, each copied from the matching row of the template
 * column. Formats only — nothing it copies carries a value or a formula.
 */
function styleLikeColumn(
  sheet: ExcelScript.Worksheet,
  templateColumn: string,
  targetColumn: string,
  headerRow: number,
  subHeaderRow: number,
  firstRow: number,
  lastPowerHouseRow: number,
  totalRow: number
) {
  const pairs = [
    { from: `${templateColumn}${headerRow}`, to: `${targetColumn}${headerRow}` },
    { from: `${templateColumn}${subHeaderRow}`, to: `${targetColumn}${subHeaderRow}` },
    { from: `${templateColumn}${firstRow}`, to: `${targetColumn}${firstRow}:${targetColumn}${lastPowerHouseRow}` },
    { from: `${templateColumn}${totalRow}`, to: `${targetColumn}${totalRow}` },
  ]
  for (const pair of pairs) {
    sheet.getRange(pair.to).copyFrom(sheet.getRange(pair.from), ExcelScript.RangeCopyType.formats)
  }
}

/**
 * Write the P7 figures into the snapshot columns — for the P8 close only.
 *
 * An earlier version only wrote into cells that were still empty. That sounds
 * careful and is the opposite: once the columns hold something wrong, the
 * script can no longer put it right, and re-running looks like it did nothing.
 * Which is exactly what happened.
 *
 * Anchored to the period instead. These figures ARE July, so they belong to the
 * August close and to no other. When the workbook is set to August 2026 they
 * are written unconditionally, overwriting whatever is there; in any other
 * month this function does not touch the sheet, so a real snapshot taken during
 * a later close is safe.
 */
function seedPriorMonthSplit(sheet: ExcelScript.Worksheet, workbook: ExcelScript.Workbook): string {
  const period = reportingPeriod(workbook)
  if (!period) return 'Snapshots left alone: reporting period unreadable.'
  if (period.getFullYear() !== 2026 || period.getMonth() + 1 !== 8) {
    return `Snapshots left alone: the P7 seed belongs to August 2026, workbook is on ${period.getFullYear()}-${period.getMonth() + 1}.`
  }

  for (const firstRow of MOVEMENT_BLOCK_FIRST_ROWS) {
    // Through X, one past the snapshots: an earlier run left a stray column
    // there and it has to go, or it reads as a fourth snapshot.
    sheet
      .getRangeByIndexes(firstRow - 1, 20, MOVEMENT_PH_COUNT + 1, 4)
      .clear(ExcelScript.ClearApplyTo.contents)
    sheet.getRangeByIndexes(firstRow - 3, 23, 1, 1).clear(ExcelScript.ClearApplyTo.all)
  }

  let written = 0
  for (const block of [
    { firstRow: MOVEMENT_BLOCK_FIRST_ROWS[0], seed: PRIOR_MONTH_SEED.buildings },
    { firstRow: MOVEMENT_BLOCK_FIRST_ROWS[1], seed: PRIOR_MONTH_SEED.vehicles },
  ]) {
    for (const part of MOVEMENT_SPLIT) {
      const values = block.seed[part.seed]
      const column: number[][] = []
      for (let index = 0; index <= MOVEMENT_PH_COUNT; index++) column.push([values[index]])
      sheet
        .getRange(`${part.snapshot}${block.firstRow}:${part.snapshot}${block.firstRow + MOVEMENT_PH_COUNT}`)
        .setValues(column)
      written += column.length
    }
  }
  return `Wrote the P7 seed into ${written} snapshot cells (August 2026 only).`
}

/** The reporting period as a date, or null when the named cell is missing. */
function reportingPeriod(workbook: ExcelScript.Workbook): Date | null {
  const named = workbook.getNamedItem(SETUP_PERIOD_NAME)
  if (!named) return null
  const serial = Number(named.getRange().getValue())
  if (isNaN(serial) || serial <= 0) return null
  // Excel day 1 is 1900-01-01, and it counts a 1900-02-29 that never existed;
  // 1899-12-30 as the origin absorbs both.
  return new Date(Date.UTC(1899, 11, 30) + serial * 86400000)
}

/**
 * Change J: carry that split onto Mvt Schedule Details, next to the net move.
 *
 * Same shape as change C — the count block already reads Movement schedule
 * column P, so these read R and S from the same rows rather than repeating the
 * arithmetic.
 */
function applyDetailsSplit(workbook: ExcelScript.Workbook): string {
  const sheet = workbook.getWorksheet(SETUP_SHEET_DETAILS)
  const headerRow = 1
  const subHeaderRow = 2
  const totalRow = DETAILS_FIRST_PH_ROW + DETAILS_PH_COUNT
  const lastPowerHouseRow = totalRow - 1

  for (const block of DETAILS_SPLIT_BLOCKS) {
    block.columns.forEach((column, partIndex) => {
      const part = MOVEMENT_SPLIT[partIndex]

      // Column B is the existing count column on this sheet, so the new ones
      // borrow its header, its accounting format and its banded total.
      styleLikeColumn(
        sheet,
        DETAILS_RESULT_TEMPLATE_COLUMN,
        column,
        headerRow,
        subHeaderRow,
        DETAILS_FIRST_PH_ROW,
        lastPowerHouseRow,
        totalRow
      )
      sheet.getRange(`${column}${headerRow}`).setValue(part.label)

      for (let index = 0; index < DETAILS_PH_COUNT; index++) {
        const row = DETAILS_FIRST_PH_ROW + index
        const sourceRow = block.movementFirstRow + index
        sheet
          .getRange(`${column}${row}`)
          .setFormula(`='${SETUP_SHEET_MOVEMENT}'!${part.delta}${sourceRow}`)
      }
      sheet
        .getRange(`${column}${totalRow}`)
        .setFormula(`=SUM(${column}${DETAILS_FIRST_PH_ROW}:${column}${lastPowerHouseRow})`)
    })
  }

  // Column widths are deliberately left alone here: G to L are sized for the
  // BUILDINGS - NEW table further down the same sheet, and a width is a
  // property of the whole column. The headers wrap instead, like A1 and B1.
  return 'J: Mvt Schedule Details shows new, terminated and transfers beside the net move — buildings G/H/I, vehicles J/K/L.'
}

/**
 * Dress a spill block.
 *
 * A spill carries no formatting of its own — it shows whatever the cells
 * already had. Only the originally populated rows were styled, so a longer
 * month landed as raw serial dates and unrounded amounts. Tiling the formatting
 * across all 182 reserved rows fixes that but leaves the key column's blue and
 * the liability column's yellow running far below the data.
 *
 * So tile it over exactly as many rows as the spill actually produced, and
 * strip the formatting off the rest. The monthly script repeats this after its
 * refresh, because the row count changes every month.
 *
 * The first row is the template and is never cleared, even when the month is
 * empty.
 */
function styleSpillBlock(
  workbook: ExcelScript.Workbook,
  firstColumnIndex: number,
  columnCount: number
): number {
  const sheet = workbook.getWorksheet(SETUP_SHEET_DETAILS)

  // The spill was just written or just refreshed; read it after a recalculation
  // or the row count below is the previous month's.
  workbook.getApplication().calculate(ExcelScript.CalculationType.full)

  const reserved = DETAILS_SPILL_LAST_ROW - DETAILS_NEW_FIRST_ROW + 1
  const keys = sheet
    .getRangeByIndexes(DETAILS_NEW_FIRST_ROW - 1, firstColumnIndex, reserved, 1)
    .getValues()
  let filled = 0
  for (let i = 0; i < keys.length; i++) {
    const key = keys[i][0]
    if (key === null || key === undefined || String(key) === '') break
    filled++
  }

  if (filled > 1) {
    sheet
      .getRangeByIndexes(DETAILS_NEW_FIRST_ROW, firstColumnIndex, filled - 1, columnCount)
      .copyFrom(
        sheet.getRangeByIndexes(DETAILS_NEW_FIRST_ROW - 1, firstColumnIndex, 1, columnCount),
        ExcelScript.RangeCopyType.formats
      )
  }

  const firstBlankRow = DETAILS_NEW_FIRST_ROW + (filled > 1 ? filled : 1)
  if (firstBlankRow <= DETAILS_SPILL_LAST_ROW) {
    sheet
      .getRangeByIndexes(
        firstBlankRow - 1,
        firstColumnIndex,
        DETAILS_SPILL_LAST_ROW - firstBlankRow + 1,
        columnCount
      )
      .clear(ExcelScript.ClearApplyTo.formats)
  }

  return filled
}

/**
 * Change L: list the contracts that ended this month, beside the new ones.
 *
 * Definition: in this month's terminated set and not in last month's. Verified
 * against both real Anaplan exports — it gives 23 contracts for P8, which with
 * the 17 reinstated reconciles to column S on all twelve PowerHouses.
 */
function applyTerminatedBuildings(workbook: ExcelScript.Workbook): string {
  const sheet = workbook.getWorksheet(SETUP_SHEET_DETAILS)
  const firstColumn = columnLetter(DETAILS_OUT_FIRST_COLUMN_INDEX)
  const headerRow = DETAILS_NEW_FIRST_ROW - 1

  sheet
    .getRangeByIndexes(
      DETAILS_NEW_FIRST_ROW - 1,
      DETAILS_OUT_FIRST_COLUMN_INDEX,
      DETAILS_SPILL_LAST_ROW - DETAILS_NEW_FIRST_ROW + 1,
      DETAILS_OUT_COLUMN_COUNT
    )
    .clear(ExcelScript.ClearApplyTo.contents)

  // Borrow the NEW table's header and first data row, column for column, so the
  // two tables read as one pair rather than as a bolt-on.
  for (const row of [headerRow, DETAILS_NEW_FIRST_ROW]) {
    sheet
      .getRangeByIndexes(row - 1, DETAILS_OUT_FIRST_COLUMN_INDEX, 1, DETAILS_OUT_COLUMN_COUNT)
      .copyFrom(
        sheet.getRangeByIndexes(row - 1, 0, 1, DETAILS_OUT_COLUMN_COUNT),
        ExcelScript.RangeCopyType.formats
      )
  }
  DETAILS_OUT_HEADERS.forEach((header, index) => {
    sheet.getRangeByIndexes(headerRow - 1, DETAILS_OUT_FIRST_COLUMN_INDEX + index, 1, 1).setValue(header)
  })

  // Newly terminated THIS month, which is not the same as "ends this month".
  // A contract entered late with a July end date counts this month too, and one
  // that already counted last month does not count again. Set difference
  // against the snapshot, so this reconciles to Movement schedule column S.
  const priorKeys =
    `${SNAPSHOT_SHEET}!$D$${SNAPSHOT_FIRST_DATA_ROW}:$D$${SNAPSHOT_FIRST_DATA_ROW + SNAPSHOT_ROW_COUNT - 1}`
  sheet.getRange(`${firstColumn}${DETAILS_NEW_FIRST_ROW}`).setFormula(
    '=LET(t,Table1,' +
    `keep,${TERMINATED_CONDITION}*ISNA(XMATCH(INDEX(t,,1),${priorKeys})),` +
    'FILTER(CHOOSECOLS(t,1,2,3,6,10,11,14,15,16,26,27),keep,""))'
  )

  // Count above the header, for the same reason the NEW totals sit there: a
  // spill grows and shrinks and would hit anything placed below it.
  const countLabel = sheet.getRangeByIndexes(16, DETAILS_OUT_FIRST_COLUMN_INDEX - 1, 1, 1)
  countLabel.setValue('Aantal')
  countLabel.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right)
  countLabel.getFormat().getFont().setBold(true)
  const count = sheet.getRange(`${firstColumn}17`)
  count.setFormula(`=SUM(--(CHOOSECOLS(${firstColumn}${DETAILS_NEW_FIRST_ROW}#,1)<>""))`)
  count.setNumberFormat('#,##0')
  count.getFormat().getFont().setBold(true)

  // Match the NEW table's column widths rather than inventing new ones.
  for (let index = 0; index < DETAILS_OUT_COLUMN_COUNT; index++) {
    const width = sheet.getRangeByIndexes(headerRow - 1, index, 1, 1).getFormat().getColumnWidth()
    sheet
      .getRangeByIndexes(headerRow - 1, DETAILS_OUT_FIRST_COLUMN_INDEX + index, 1, 1)
      .getFormat()
      .setColumnWidth(width)
  }

  const styled = styleSpillBlock(workbook, DETAILS_OUT_FIRST_COLUMN_INDEX, DETAILS_OUT_COLUMN_COUNT)
  return `L: BUILDINGS - TERMINATED spills beside the new ones from ${firstColumn}${DETAILS_NEW_FIRST_ROW}; ${styled} row(s) styled.`
}

/** 0 -> A, 27 -> AB. */
function columnLetter(index: number): string {
  let letters = ''
  let remaining = index
  while (remaining >= 0) {
    letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.charAt(remaining % 26) + letters
    remaining = Math.floor(remaining / 26) - 1
  }
  return letters
}

/**
 * The condition the Terminated pivot applies, as a formula fragment over
 * Table1 bound to `t`: Land and buildings, reasonably certain end date in the
 * reporting year up to and including the reporting month, and no transfer-out
 * date.
 *
 * The transfer-out test is not decoration — the pivot drops transferred-out
 * contracts, and without it none of these blocks reconcile to the figure they
 * sit beside. Checked against both real Anaplan exports: 187 at 31 July and
 * 193 at 31 August, each equal to what the workbook reported.
 *
 * A blank end date is excluded by the year test: YEAR of an empty cell is 1900.
 */
const TERMINATED_CONDITION =
  '(INDEX(t,,26)="Land and buildings")' +
  `*(YEAR(INDEX(t,,11))=YEAR(${SETUP_PERIOD_NAME}))` +
  `*(MONTH(INDEX(t,,11))<=MONTH(${SETUP_PERIOD_NAME}))` +
  '*(INDEX(t,,13)="")'

/**
 * Change M: the snapshot sheet.
 *
 * A2 spills this month's terminated set (key and end date). D and E hold last
 * month's, typed. The monthly step copies A over to D before the new export
 * goes in, exactly like columns O, U, V and W on Movement schedule.
 */
function applySnapshotSheet(workbook: ExcelScript.Workbook): string {
  let sheet = workbook.getWorksheet(SNAPSHOT_SHEET)
  if (!sheet) sheet = workbook.addWorksheet(SNAPSHOT_SHEET)

  // Say on the sheet what it is. Without this it is 187 rows of keys and dates
  // with no explanation, which is worse than useless — it looks like a mistake.
  sheet.getRange('A1').setValue('SNAPSHOT')
  sheet.getRange('A1').getFormat().getFont().setBold(true)
  sheet.getRange('A1').getFormat().getFont().setSize(14)
  SNAPSHOT_NOTE.forEach((line, index) => {
    sheet.getRange(`A${index + 2}`).setValue(line)
  })
  sheet.getRange(`A${SNAPSHOT_NOTE.length + 1}`).getFormat().getFont().setBold(true)

  const header = SNAPSHOT_FIRST_DATA_ROW - 1
  const last = SNAPSHOT_FIRST_DATA_ROW + SNAPSHOT_ROW_COUNT - 1
  sheet.getRange(`A${header}`).setValue('Contract — deze maand')
  sheet.getRange(`B${header}`).setValue('Einddatum')
  sheet.getRange(`D${header}`).setValue('Contract — vorige maand')
  sheet.getRange(`E${header}`).setValue('Einddatum')
  sheet
    .getRange(`A${header}:E${header}`)
    .copyFrom(sheet.getRange(`A${header}`), ExcelScript.RangeCopyType.formats)
  sheet.getRange(`A${header}:E${header}`).getFormat().getFont().setBold(true)

  // A count beside each list, so the two are comparable at a glance without
  // scrolling to the bottom. 193 against 187 for P8.
  sheet.getRange(`B${header - 1}`).setFormula(
    `=SUM(--(CHOOSECOLS(A${SNAPSHOT_FIRST_DATA_ROW}#,1)<>""))&" contracten"`
  )
  sheet.getRange(`E${header - 1}`).setFormula(
    `=SUM(--(D${SNAPSHOT_FIRST_DATA_ROW}:D${last}<>""))&" contracten"`
  )
  sheet.getRange(`B${header - 1}:E${header - 1}`).getFormat().getFont().setItalic(true)

  sheet.getRange(`A${SNAPSHOT_FIRST_DATA_ROW}`).setFormula(
    '=LET(t,Table1,' +
    `keep,${TERMINATED_CONDITION},` +
    'FILTER(HSTACK(INDEX(t,,1),INDEX(t,,11)),keep,""))'
  )
  sheet.getRange(`B${SNAPSHOT_FIRST_DATA_ROW}:B${last}`).setNumberFormat('dd/mm/yyyy')
  sheet.getRange(`E${SNAPSHOT_FIRST_DATA_ROW}:E${last}`).setNumberFormat('dd/mm/yyyy')

  // Dates showed as ###### because the columns were never widened.
  sheet.getRange('A:A').getFormat().setColumnWidth(230)
  sheet.getRange('B:B').getFormat().setColumnWidth(90)
  sheet.getRange('C:C').getFormat().setColumnWidth(24)
  sheet.getRange('D:D').getFormat().setColumnWidth(230)
  sheet.getRange('E:E').getFormat().setColumnWidth(90)

  const seeded = seedPriorTerminated(sheet, workbook)
  return `M: ${SNAPSHOT_SHEET} sheet — explained at the top, lists from row ${SNAPSHOT_FIRST_DATA_ROW}. ${seeded}`
}

/**
 * Anchored to August 2026 for the same reason as the other seed: these are July
 * figures, they belong to the P8 close, and a guard that only fills empty cells
 * cannot repair a wrong one.
 */
function seedPriorTerminated(sheet: ExcelScript.Worksheet, workbook: ExcelScript.Workbook): string {
  const period = reportingPeriod(workbook)
  if (!period) return 'Prior snapshot left alone: reporting period unreadable.'
  if (period.getFullYear() !== 2026 || period.getMonth() + 1 !== 8) {
    return 'Prior snapshot left alone: the P7 seed belongs to August 2026.'
  }
  sheet.getRangeByIndexes(SNAPSHOT_FIRST_DATA_ROW - 1, 3, SNAPSHOT_ROW_COUNT, 2).clear(ExcelScript.ClearApplyTo.contents)
  const rows = PRIOR_TERMINATED_SEED.map((entry) => {
    const parts = entry[1].split('-')
    return [entry[0], new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]))]
  })
  const keys: string[][] = rows.map((row) => [row[0] as string])
  sheet.getRangeByIndexes(SNAPSHOT_FIRST_DATA_ROW - 1, 3, keys.length, 1).setValues(keys)
  // Dates go in as formulas for the same reason the period cell does: a string
  // would be parsed under whatever locale the workbook opens in.
  PRIOR_TERMINATED_SEED.forEach((entry, index) => {
    const parts = entry[1].split('-')
    sheet
      .getRangeByIndexes(SNAPSHOT_FIRST_DATA_ROW - 1 + index, 4, 1, 1)
      .setFormula(`=DATE(${Number(parts[0])},${Number(parts[1])},${Number(parts[2])})`)
  })
  return `Seeded ${PRIOR_TERMINATED_SEED.length} contracts from the P7 export.`
}

/**
 * Change N: the contracts that were terminated last month and are not any more.
 *
 * This is the half that cannot come from one export. A contract whose end date
 * moved out to 2029 looks like any other live contract in this month's data —
 * there is no flag on it. The only way to see it is to have kept last month's
 * list, which is what the snapshot sheet is for.
 *
 * Five columns: the key, the end date it had last month, entity and description
 * for readability, and the end date it has now.
 */
function applyReinstatedBuildings(workbook: ExcelScript.Workbook): string {
  const sheet = workbook.getWorksheet(SETUP_SHEET_DETAILS)
  const first = columnLetter(DETAILS_BACK_FIRST_COLUMN_INDEX)
  const headerRow = DETAILS_NEW_FIRST_ROW - 1
  const width = DETAILS_BACK_HEADERS.length

  sheet
    .getRangeByIndexes(
      DETAILS_NEW_FIRST_ROW - 1,
      DETAILS_BACK_FIRST_COLUMN_INDEX,
      DETAILS_SPILL_LAST_ROW - DETAILS_NEW_FIRST_ROW + 1,
      width
    )
    .clear(ExcelScript.ClearApplyTo.contents)

  for (const row of [headerRow, DETAILS_NEW_FIRST_ROW]) {
    sheet
      .getRangeByIndexes(row - 1, DETAILS_BACK_FIRST_COLUMN_INDEX, 1, width)
      .copyFrom(
        sheet.getRangeByIndexes(row - 1, 0, 1, width),
        ExcelScript.RangeCopyType.formats
      )
  }
  DETAILS_BACK_HEADERS.forEach((header, index) => {
    sheet.getRangeByIndexes(headerRow - 1, DETAILS_BACK_FIRST_COLUMN_INDEX + index, 1, 1).setValue(header)
  })

  const lastRow = SNAPSHOT_FIRST_DATA_ROW + SNAPSHOT_ROW_COUNT - 1
  const priorKeys = `${SNAPSHOT_SHEET}!$D$${SNAPSHOT_FIRST_DATA_ROW}:$D$${lastRow}`
  const priorEnds = `${SNAPSHOT_SHEET}!$E$${SNAPSHOT_FIRST_DATA_ROW}:$E$${lastRow}`
  sheet.getRange(`${first}${DETAILS_NEW_FIRST_ROW}`).setFormula(
    '=LET(t,Table1,' +
    `prior,${priorKeys},priorEnd,${priorEnds},` +
    `cur,FILTER(INDEX(t,,1),${TERMINATED_CONDITION},""),` +
    'keep,(prior<>"")*ISNA(XMATCH(prior,cur)),' +
    'FILTER(HSTACK(prior,priorEnd,' +
    'XLOOKUP(prior,INDEX(t,,1),INDEX(t,,2),""),' +
    'XLOOKUP(prior,INDEX(t,,1),INDEX(t,,3),""),' +
    'XLOOKUP(prior,INDEX(t,,1),INDEX(t,,11),"")),keep,""))'
  )

  const countLabel = sheet.getRangeByIndexes(16, DETAILS_BACK_FIRST_COLUMN_INDEX - 1, 1, 1)
  countLabel.setValue('Aantal')
  countLabel.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right)
  countLabel.getFormat().getFont().setBold(true)
  const count = sheet.getRange(`${first}17`)
  count.setFormula(`=SUM(--(CHOOSECOLS(${first}${DETAILS_NEW_FIRST_ROW}#,1)<>""))`)
  count.setNumberFormat('#,##0')
  count.getFormat().getFont().setBold(true)

  for (let index = 0; index < width; index++) {
    const source = sheet.getRangeByIndexes(headerRow - 1, index, 1, 1).getFormat().getColumnWidth()
    sheet
      .getRangeByIndexes(headerRow - 1, DETAILS_BACK_FIRST_COLUMN_INDEX + index, 1, 1)
      .getFormat()
      .setColumnWidth(source)
  }

  const styled = styleSpillBlock(workbook, DETAILS_BACK_FIRST_COLUMN_INDEX, width)
  return `N: BUILDINGS - REINSTATED spills from ${first}${DETAILS_NEW_FIRST_ROW}; ${styled} row(s) styled.`
}
