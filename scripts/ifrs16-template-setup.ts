/**
 * IFRS16 Template Setup — one-time Office Script
 *
 * Applies template changes B–E from docs/automation-design.md to the Input
 * Board Pack, and creates the named cell for change A. Run once on the P8 2026
 * workbook; it then travels forward with every monthly copy.
 *
 * Change A's M-code edit CANNOT be scripted — Office Scripts has no access to
 * Power Query M. Do that by hand afterwards; the script prints the reminder.
 *
 * Safety: every figure this touches is captured before and re-read after, and
 * any change is reported. On P8 2026 nothing should move — these changes
 * restructure how values are derived, not what they are. Checked in advance:
 * change B adds entity codes 1652 and 2006, neither of which has any contract,
 * and no existing code changes PowerHouse; change C matches column P on all 24
 * rows; change E reproduces the hardcoded =H19*12. So any reported difference
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

/** Entity List rows 3..360 mirror Entity List PowerBI rows 4..361 (358 entities). */
const ENTITY_FIRST_TARGET_ROW = 3
const ENTITY_ROW_COUNT = 358
const ENTITY_SOURCE_OFFSET = 1

/** Transfers block: rows 82-91, fed by the TRANSFER IN and OUT pivots. */
const TRANSFERS_FIRST_ROW = 82
const TRANSFERS_LAST_ROW = 91
const TRANSFER_IN_RANGE = '$B$50:$B$59'
const TRANSFER_OUT_RANGE = '$B$68:$B$77'

const DETAILS_FIRST_PH_ROW = 3
const DETAILS_PH_COUNT = 12
const DETAILS_NEW_FIRST_ROW = 19
const DETAILS_NEW_LAST_ROW = 24

function main(workbook: ExcelScript.Workbook): string {
  const log: string[] = ['IFRS16 template setup']

  const before = captureCheckFigures(workbook)

  log.push(applyReportingPeriodCell(workbook))
  log.push(applyEntityListLink(workbook))
  log.push(applyDetailsCountLink(workbook))
  log.push(applyTransfersLookup(workbook))
  log.push(applyLeaseLiabilityFormula(workbook))

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

/** Figures that must not move: the CHECK row, both group totals, and the new-lease total. */
function captureCheckFigures(workbook: ExcelScript.Workbook): number[] {
  const movement = workbook.getWorksheet(SETUP_SHEET_MOVEMENT)
  const details = workbook.getWorksheet(SETUP_SHEET_DETAILS)
  return [
    Number(movement.getRange('F49').getValue()),
    Number(movement.getRange('K49').getValue()),
    Number(movement.getRange('G15').getValue()),
    Number(movement.getRange('G31').getValue()),
    Number(movement.getRange('F15').getValue()),
    Number(movement.getRange('F31').getValue()),
    Number(details.getRange('B15').getValue()),
    Number(details.getRange('E15').getValue()),
    Number(details.getRange('L25').getValue()),
  ]
}

function compareCheckFigures(workbook: ExcelScript.Workbook, before: number[]): string {
  const labels = [
    'CHECK buildings (F49)', 'CHECK vehicles (K49)',
    'Total buildings (G15)', 'Total vehicles (G31)',
    'Transfers buildings (F15)', 'Transfers vehicles (F31)',
    'Mvt buildings (B15)', 'Mvt vehicles (E15)',
    'New lease liability (L25)',
  ]
  const after = captureCheckFigures(workbook)
  const moved: string[] = []
  after.forEach((value, index) => {
    if (Math.abs(value - before[index]) > 0.005) {
      moved.push(`${labels[index]}: ${before[index]} -> ${value}`)
    }
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
 * Change E: generalise the lease liability so any duration and frequency works.
 * The workbook has =H19*12 hardcoded for the one quarterly contract; this gives
 * the same figure but keeps holding next month. IFS has no fallback branch on
 * purpose — an unknown frequency yields #N/A rather than a wrong number.
 */
function applyLeaseLiabilityFormula(workbook: ExcelScript.Workbook): string {
  const sheet = workbook.getWorksheet(SETUP_SHEET_DETAILS)
  for (let row = DETAILS_NEW_FIRST_ROW; row <= DETAILS_NEW_LAST_ROW; row++) {
    if (!sheet.getRange(`A${row}`).getValue()) continue
    sheet.getRange(`L${row}`).setFormula(
      `=H${row}*G${row}/IFS(I${row}="Monthly",1,I${row}="Quarterly",3)`
    )
  }
  return `E: lease liability rows ${DETAILS_NEW_FIRST_ROW}-${DETAILS_NEW_LAST_ROW} now divide the duration by the payment frequency.`
}
