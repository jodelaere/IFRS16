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

/** BUILDINGS - NEW spans A..L. */
const DETAILS_COLUMN_COUNT = 12

function main(workbook: ExcelScript.Workbook): string {
  const log: string[] = ['IFRS16 template setup']

  const before = captureCheckFigures(workbook, 'L25')

  log.push(applyReportingPeriodCell(workbook))
  log.push(applyEntityListLink(workbook))
  log.push(applyDetailsCountLink(workbook))
  log.push(applyTransfersLookup(workbook))
  log.push(applyPeriodLabels(workbook))
  log.push(applyNewBuildingsSpill(workbook))

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

  sheet.getRange(`A${DETAILS_NEW_FIRST_ROW}`).setFormula(
    '=LET(t,Table1,' +
    'pay,INDEX(t,,15),dur,INDEX(t,,14),freq,INDEX(t,,16),' +
    'liab,pay*dur/IFS(freq="Monthly",1,freq="Quarterly",3),' +
    `keep,(INDEX(t,,26)="Land and buildings")*(TEXT(INDEX(t,,6),"yyyymm")=TEXT(${SETUP_PERIOD_NAME},"yyyymm")),` +
    'FILTER(HSTACK(CHOOSECOLS(t,1,2,3,6,10,11,14,15,16,26,27),liab),keep,""))'
  )

  // The total sits above the table: a spill grows and shrinks, so anything
  // directly beneath it would block it with #SPILL!.
  const totalLabel = sheet.getRange('K17')
  totalLabel.setValue('Total')
  totalLabel.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right)
  totalLabel.getFormat().getFont().setBold(true)

  const total = sheet.getRange('L17')
  total.setFormula(`=SUM(CHOOSECOLS(A${DETAILS_NEW_FIRST_ROW}#,12))`)
  total.setNumberFormat('#,##0.00')
  total.getFormat().getFont().setBold(true)
  total.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeTop).setStyle(ExcelScript.BorderLineStyle.continuous)
  total.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setStyle(ExcelScript.BorderLineStyle.double)

  const styled = styleNewBuildingsSpill(workbook)
  return `H: BUILDINGS - NEW spills from Table1 for the reporting month; labelled total in K17/L17, ${styled} row(s) styled.`
}

/**
 * Dress the spill.
 *
 * A spill carries no formatting of its own — it shows whatever the cells
 * already had. Only the six originally populated rows were styled, so a longer
 * month landed as raw serial dates and unrounded amounts. Tiling the formatting
 * across all 182 reserved rows fixes that but leaves column A's blue and column
 * L's yellow running far below the data.
 *
 * So tile it over exactly as many rows as the spill actually produced, and
 * strip the formatting off the rest. The monthly script repeats this after its
 * refresh, because the row count changes every month.
 *
 * Row 19 is the template and is never cleared, even when the month is empty.
 */
function styleNewBuildingsSpill(workbook: ExcelScript.Workbook): number {
  const sheet = workbook.getWorksheet(SETUP_SHEET_DETAILS)

  // The spill was just written or just refreshed; read it after a recalculation
  // or the row count below is the previous month's.
  workbook.getApplication().calculate(ExcelScript.CalculationType.full)

  const reserved = DETAILS_SPILL_LAST_ROW - DETAILS_NEW_FIRST_ROW + 1
  const keys = sheet.getRangeByIndexes(DETAILS_NEW_FIRST_ROW - 1, 0, reserved, 1).getValues()
  let filled = 0
  for (let i = 0; i < keys.length; i++) {
    const key = keys[i][0]
    if (key === null || key === undefined || String(key) === '') break
    filled++
  }

  if (filled > 1) {
    sheet
      .getRangeByIndexes(DETAILS_NEW_FIRST_ROW, 0, filled - 1, DETAILS_COLUMN_COUNT)
      .copyFrom(
        sheet.getRangeByIndexes(DETAILS_NEW_FIRST_ROW - 1, 0, 1, DETAILS_COLUMN_COUNT),
        ExcelScript.RangeCopyType.formats
      )
  }

  const firstBlankRow = DETAILS_NEW_FIRST_ROW + (filled > 1 ? filled : 1)
  if (firstBlankRow <= DETAILS_SPILL_LAST_ROW) {
    sheet
      .getRangeByIndexes(
        firstBlankRow - 1,
        0,
        DETAILS_SPILL_LAST_ROW - firstBlankRow + 1,
        DETAILS_COLUMN_COUNT
      )
      .clear(ExcelScript.ClearApplyTo.formats)
  }

  return filled
}
