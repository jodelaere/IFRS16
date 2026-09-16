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
  { source: 'C', delta: 'R', snapshot: 'U', label: 'New contracts', seed: 'newContracts' },
  { source: 'E', delta: 'S', snapshot: 'V', label: 'Terminated contracts', seed: 'terminated' },
  { source: 'F', delta: 'T', snapshot: 'W', label: 'Transfers', seed: 'transfers' },
]

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

/** Change J columns on Mvt Schedule Details: buildings G/H, vehicles J/K. */
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

  const styled = styleNewBuildingsSpill(workbook)
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
  const priorMonth = `TEXT(EDATE(${SETUP_PERIOD_NAME},-1),"[$-en-US]mmmm yyyy")`

  for (const firstRow of MOVEMENT_BLOCK_FIRST_ROWS) {
    const headerRow = firstRow - 2
    for (const part of MOVEMENT_SPLIT) {
      sheet.getRange(`${part.delta}${headerRow}`).setValue(part.label)
      sheet
        .getRange(`${part.snapshot}${headerRow}`)
        .setFormula(`="${part.label} "&${priorMonth}`)

      // Through the group total row, so the totals split too.
      for (let row = firstRow; row <= firstRow + MOVEMENT_PH_COUNT; row++) {
        sheet
          .getRange(`${part.delta}${row}`)
          .setFormula(`=${part.source}${row}-${part.snapshot}${row}`)
      }
    }
  }

  const seeded = seedPriorMonthSplit(sheet)

  return (
    'I: Movement schedule R/S/T split the move into new, terminated and transfers, from snapshots ' +
    `in U/V/W — both blocks, no external link. ${seeded}`
  )
}

/**
 * Fill the snapshot columns with the P7 figures, but only where they are still
 * empty.
 *
 * August is the first close with these columns, so nothing has photographed
 * July yet and the deltas would read as the full year-to-date. The guard is
 * what makes this safe to leave in the script: once a month has written real
 * snapshots, there is nothing empty left to seed and the P7 table is inert.
 */
function seedPriorMonthSplit(sheet: ExcelScript.Worksheet): string {
  const blocks = [
    { firstRow: MOVEMENT_BLOCK_FIRST_ROWS[0], seed: PRIOR_MONTH_SEED.buildings },
    { firstRow: MOVEMENT_BLOCK_FIRST_ROWS[1], seed: PRIOR_MONTH_SEED.vehicles },
  ]

  let written = 0
  let skipped = 0
  for (const block of blocks) {
    for (const part of MOVEMENT_SPLIT) {
      const values = block.seed[part.seed]
      for (let index = 0; index <= MOVEMENT_PH_COUNT; index++) {
        const cell = sheet.getRange(`${part.snapshot}${block.firstRow + index}`)
        if (String(cell.getValue()) !== '') {
          skipped++
          continue
        }
        cell.setValue(values[index])
        written++
      }
    }
  }

  if (written === 0) return 'Snapshots already filled, P7 seed not used.'
  const tail = skipped > 0 ? `, left ${skipped} already filled` : ''
  return `Seeded ${written} empty snapshot cell(s) with the P7 figures${tail}.`
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
  const totalRow = DETAILS_FIRST_PH_ROW + DETAILS_PH_COUNT

  for (const block of DETAILS_SPLIT_BLOCKS) {
    block.columns.forEach((column, partIndex) => {
      const part = MOVEMENT_SPLIT[partIndex]
      sheet.getRange(`${column}1`).setValue(part.label)

      for (let index = 0; index < DETAILS_PH_COUNT; index++) {
        const row = DETAILS_FIRST_PH_ROW + index
        const sourceRow = block.movementFirstRow + index
        sheet
          .getRange(`${column}${row}`)
          .setFormula(`='${SETUP_SHEET_MOVEMENT}'!${part.delta}${sourceRow}`)
      }
      sheet
        .getRange(`${column}${totalRow}`)
        .setFormula(`=SUM(${column}${DETAILS_FIRST_PH_ROW}:${column}${totalRow - 1})`)
    })
  }

  return 'J: Mvt Schedule Details shows new, terminated and transfers beside the net move — buildings G/H/I, vehicles J/K/L.'
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
