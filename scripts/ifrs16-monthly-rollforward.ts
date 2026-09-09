/**
 * IFRS16 Monthly Roll-Forward — Office Script
 *
 * Runs against the new period's "Input Board Pack.xlsx", already copied forward
 * from the previous period by the Power Automate flow. Every sheet name, table
 * name, cell reference and formula below was verified against
 * "202608 - IFRS16 - 3 - Input Board Pack.xlsx" (P8 2026).
 *
 * Invoked from Power Automate via "Run script". Power Automate supplies the two
 * Anaplan exports as LeaseRow arrays ("Updated Lease properties (2).xlsx" →
 * 2.9 Input, "(3).xlsx" → 2.10 Input) plus the period parameters.
 *
 * REQUIRES the one-time template changes in docs/automation-design.md — most
 * importantly the Power Query reporting-period parameter. Without it the queries
 * derive the reporting month from DateTime.LocalNow() and this script cannot
 * control which month they compute.
 */

interface LeaseRow {
  Key: string // "1001__2ATC794" — entity code + contract name, split by the query
  Entity: string
  LeaseDescription: string
  CostCenter: string
  LocalCostCenterCode: string
  LeaseCommencementDate: string
  PurchaseOption: string
  ExerciseOfPurchaseOptionDate: string
  ExercisePriceOfPurchaseOption: string
  ReasonablyCertainEndDateSelection: string
  ReasonablyCertainEndDate: string
  TransferInDate: string
  TransferOutDate: string
  LeaseDuration: number
  FixedPayment: number
  PaymentFrequency: string
  PaymentAtBeginningOfPeriod: string
  RevisionType: string
  IndexOrRate: string
  LeaseRevisionFrequencyMonths: string
  FirstRevisionAfterMonths: string
  ReferenceIndexRateDate: string
  ProvisionForDismantlingCosts: string
  Status: string
  LastModificationStatus: string
  AssetCategory: string
  LeasedCapacity: number
  TypeMotor: string
}

/** The 28 columns of the Anaplan export, in sheet order. */
const LEASE_ROW_HEADERS: (keyof LeaseRow)[] = [
  'Key', 'Entity', 'LeaseDescription', 'CostCenter', 'LocalCostCenterCode',
  'LeaseCommencementDate', 'PurchaseOption', 'ExerciseOfPurchaseOptionDate',
  'ExercisePriceOfPurchaseOption', 'ReasonablyCertainEndDateSelection',
  'ReasonablyCertainEndDate', 'TransferInDate', 'TransferOutDate',
  'LeaseDuration', 'FixedPayment', 'PaymentFrequency',
  'PaymentAtBeginningOfPeriod', 'RevisionType', 'IndexOrRate',
  'LeaseRevisionFrequencyMonths', 'FirstRevisionAfterMonths',
  'ReferenceIndexRateDate', 'ProvisionForDismantlingCosts', 'Status',
  'LastModificationStatus', 'AssetCategory', 'LeasedCapacity', 'TypeMotor',
]

interface RollForwardParams {
  /** Last day of the new reporting period, e.g. "2026-09-30". Drives every label. */
  periodEndDate: string
  /** e.g. "September 2026" — the month label typed into the schedule headers. */
  newMonthLabel: string
  /** e.g. "August 2026" — becomes the comparative column header. */
  previousMonthLabel: string
  /** e.g. "P09" — used for the "mvt P09" column headers. */
  newPeriodCode: string
}

// --- Verified sheet, table and cell references (P8 2026 workbook) -----------

const SHEET_MVT_DETAILS = 'Mvt Schedule Details'
const SHEET_MOVEMENT = 'Movement schedule'
const SHEET_ENTITY_LIST_PBI = 'Entity List PowerBI'

/** Anaplan data lands in these Excel Tables; the Power Queries read them by name. */
const TABLE_29_INPUT = 'Table2.9'
const TABLE_210_INPUT = 'Table1'

/** Named cell holding the period end date that both Power Queries read. */
const REPORTING_PERIOD_NAME = 'ReportingPeriodEnd'

/**
 * Movement schedule layout. Two blocks with identical column structure:
 * A PowerHouse | B opening (Dec) | C new | D M&A | E terminated | F transfers |
 * G current month | I calculated | J difference | L plug (new) | M plug (prior) |
 * O prior month | P movement.
 * Buildings rows 3-14, vehicles rows 19-30, 12 PowerHouses each.
 */
const BUILDINGS_FIRST_ROW = 3
const VEHICLES_FIRST_ROW = 19
const POWERHOUSE_COUNT = 12

/**
 * G1 and G17 are formulas (=K34), so the current-month label is edited once in
 * the presentation table at row 34 and both block headers follow.
 */
const CELL_MONTH_LABEL_BUILDINGS = 'F34'
const CELL_MONTH_LABEL_VEHICLES = 'K34'

/** Entity List PowerBI: Table_ExternalData_1 at A3:S361, headers row 3, data row 4. */
const ENTITY_LIST_PBI_FIRST_DATA_ROW = 4
const ENTITY_LIST_PBI_CODE_COLUMN = 1 // A
const ENTITY_LIST_PBI_DESCRIPTION_COLUMN = 2 // B
const ENTITY_LIST_PBI_POWERHOUSE_COLUMN = 6 // F — note: NOT C, which is LE Country Long

/** PowerHouse values that mean "not mapped yet in MDM". */
const UNMAPPED_POWERHOUSE_VALUES = ['', 'N/A']

/** Excel caps a single setValues payload; write large inputs in slices. */
const WRITE_CHUNK_ROWS = 5000

function main(
  workbook: ExcelScript.Workbook,
  anaplan29Rows: LeaseRow[],
  anaplan210Rows: LeaseRow[],
  params: RollForwardParams
): string {
  // Must precede any refresh: the Power Queries derive every IN/OUT flag from
  // this parameter, so setting it afterwards would compute the wrong period.
  setReportingPeriod(workbook, params.periodEndDate)

  // Column G still holds the previous month's counts until the new Anaplan data
  // is loaded, so capture it first — that is exactly what column O needs.
  const previousMonthCounts = capturePreviousMonthCounts(workbook)

  replaceInputTable(workbook, TABLE_29_INPUT, anaplan29Rows)
  replaceInputTable(workbook, TABLE_210_INPUT, anaplan210Rows)

  refreshQueriesAndPivots(workbook)

  writePreviousMonthCounts(workbook, previousMonthCounts)
  rollMovementSchedule(workbook, params)
  appendNewBuildings(workbook, anaplan210Rows, params)

  refreshQueriesAndPivots(workbook)

  return buildReviewReport(workbook, params)
}

/**
 * Writes the reporting period into the named cell the Power Queries read.
 * Without template change A the queries still use DateTime.LocalNow(), which
 * makes the output depend on when the flow happens to run — so this fails loudly
 * rather than silently producing whichever month the clock implies.
 */
function setReportingPeriod(workbook: ExcelScript.Workbook, periodEndDate: string) {
  const namedItem = workbook.getNamedItem(REPORTING_PERIOD_NAME)
  if (!namedItem) {
    throw new Error(
      `Named cell "${REPORTING_PERIOD_NAME}" not found. Apply template change A first — ` +
      'until then the Power Queries take the reporting month from the system clock ' +
      'and this script cannot control which period they compute.'
    )
  }
  namedItem.getRange().setValue(periodEndDate)
}

/** Column G for both blocks, including the group total row. */
function capturePreviousMonthCounts(workbook: ExcelScript.Workbook): (string | number | boolean)[][][] {
  const sheet = workbook.getWorksheet(SHEET_MOVEMENT)
  if (!sheet) throw new Error(`Sheet "${SHEET_MOVEMENT}" not found.`)

  return [BUILDINGS_FIRST_ROW, VEHICLES_FIRST_ROW].map((firstRow) =>
    sheet.getRange(`G${firstRow}:G${firstRow + POWERHOUSE_COUNT}`).getValues()
  )
}

/**
 * Writes the captured counts into column O ("prior month"), which is a typed
 * value column — it used to be =SUM(H:L), which resolved to current month plus
 * plug and therefore did not represent the prior month at all. Column P
 * (=G-O) only yields a real month-on-month movement once O holds actual prior
 * month figures.
 */
function writePreviousMonthCounts(workbook: ExcelScript.Workbook, counts: (string | number | boolean)[][][]) {
  const sheet = workbook.getWorksheet(SHEET_MOVEMENT)
  if (!sheet) throw new Error(`Sheet "${SHEET_MOVEMENT}" not found.`)

  ;[BUILDINGS_FIRST_ROW, VEHICLES_FIRST_ROW].forEach((firstRow, index) => {
    sheet.getRange(`O${firstRow}:O${firstRow + POWERHOUSE_COUNT}`).setValues(counts[index])
  })
}

/**
 * Overwrites an input table with the fresh Anaplan export. Resizing the table
 * matters: the Power Queries read Table1 / Table2.9 by name, so a plain range
 * paste would leave them reading the previous month's row count.
 */
function replaceInputTable(workbook: ExcelScript.Workbook, tableName: string, rows: LeaseRow[]) {
  const table = workbook.getTable(tableName)
  if (!table) {
    throw new Error(`Table "${tableName}" not found — the Power Query reads it by name, so it must exist.`)
  }
  if (rows.length === 0) {
    throw new Error(`No Anaplan rows supplied for "${tableName}" — refusing to empty the input table.`)
  }

  const sheet = table.getWorksheet()
  const headerRange = table.getHeaderRowRange()
  const firstDataRow = headerRange.getRowIndex() + 1
  const firstColumn = headerRange.getColumnIndex()
  const columnCount = LEASE_ROW_HEADERS.length

  const existingBody = table.getRangeBetweenHeaderAndTotal()
  if (existingBody) {
    existingBody.clear(ExcelScript.ClearApplyTo.contents)
  }

  table.resize(sheet.getRangeByIndexes(headerRange.getRowIndex(), firstColumn, rows.length + 1, columnCount))

  for (let offset = 0; offset < rows.length; offset += WRITE_CHUNK_ROWS) {
    const slice = rows.slice(offset, offset + WRITE_CHUNK_ROWS)
    const values = slice.map((row) => LEASE_ROW_HEADERS.map((key) => row[key] ?? ''))
    sheet
      .getRangeByIndexes(firstDataRow + offset, firstColumn, slice.length, columnCount)
      .setValues(values as (string | number)[][])
  }
}

/**
 * Both Power Queries source from $Workbook$ (the input tables), so they need no
 * external credentials. The five PivotTables have refreshOnLoad=false and feed
 * the Movement schedule XLOOKUPs, so they must be refreshed explicitly — and
 * only after the queries have rebuilt their output tables.
 *
 * The Entity List PowerBI tab is a live MSOLAP connection to Power BI and is
 * NOT refreshed here; see docs/automation-design.md.
 */
function refreshQueriesAndPivots(workbook: ExcelScript.Workbook) {
  workbook.refreshAllPowerQueries()
  workbook.getPivotTables().forEach((pivotTable) => pivotTable.refresh())
  workbook.getApplication().calculate(ExcelScript.CalculationType.fullRebuild)
}

/**
 * Rolls the period labels and shifts the plug column one month back.
 *
 * The plug in column L feeds the Terminated contracts formula
 * (E = -XLOOKUP(pivot OUT) + L), so it reconciles the YTD build-up
 * B+C+D+E+F to the 2.9 snapshot count in column G — it is the difference
 * between the 2.10 and 2.9 cuts, not a prior-month device.
 *
 * It therefore has to be re-established for each new cut. This preserves last
 * month's value in column M and leaves L untouched for the preparer; the review
 * report states how much extra plug each PowerHouse needs to tie. Deliberately
 * not auto-plugged: forcing the tie would mask genuine data errors.
 */
function rollMovementSchedule(workbook: ExcelScript.Workbook, params: RollForwardParams) {
  const sheet = workbook.getWorksheet(SHEET_MOVEMENT)
  if (!sheet) throw new Error(`Sheet "${SHEET_MOVEMENT}" not found.`)

  const periodEnd = new Date(params.periodEndDate)
  const month = String(periodEnd.getUTCMonth() + 1).padStart(2, '0')
  const previousMonth = String(periodEnd.getUTCMonth() === 0 ? 12 : periodEnd.getUTCMonth()).padStart(2, '0')
  const year = periodEnd.getUTCFullYear()

  sheet.getRange(CELL_MONTH_LABEL_BUILDINGS).setValue(params.newMonthLabel)
  sheet.getRange(CELL_MONTH_LABEL_VEHICLES).setValue(params.newMonthLabel)

  sheet.getRange('L1').setValue(`Plug ${month} ${year}`)
  sheet.getRange('M1').setValue(`Plug ${previousMonth} ${year}`)

  for (const headerRow of [1, 17]) {
    sheet.getRange(`O${headerRow}`).setValue(params.previousMonthLabel)
    sheet.getRange(`P${headerRow}`).setValue(`mvt ${params.newPeriodCode}`)
  }

  for (const firstRow of [BUILDINGS_FIRST_ROW, VEHICLES_FIRST_ROW]) {
    const plugs = sheet.getRange(`L${firstRow}:L${firstRow + POWERHOUSE_COUNT - 1}`).getValues()
    sheet.getRange(`M${firstRow}:M${firstRow + POWERHOUSE_COUNT - 1}`).setValues(plugs)
  }

  const details = workbook.getWorksheet(SHEET_MVT_DETAILS)
  if (details) {
    details.getRange('B1').setValue(`mvt ${params.newPeriodCode}`)
    details.getRange('E1').setValue(`mvt ${params.newPeriodCode}`)
  }
}

/**
 * Appends the buildings that commenced in the reporting month to the
 * "BUILDINGS - NEW" table on Mvt Schedule Details (starting row 19).
 *
 * There is deliberately no vehicles equivalent — the workbook lists new
 * buildings in detail only; vehicles are counted, not itemised.
 *
 * Column layout (verified): A key | B entity | C description | D commencement |
 * E end date selection | F end date | G duration | H fixed payment |
 * I frequency | J asset category | K leased capacity | L lease liability.
 *
 * L = payment × number of payments, so the duration in months is divided by the
 * months per payment period. The workbook's own =H*G only holds for Monthly; a
 * Quarterly contract was overstated threefold. IFS deliberately has no fallback
 * branch: an unexpected frequency yields #N/A rather than a silently wrong
 * figure. Only Monthly and Quarterly occur in the Anaplan data today.
 */
const PAYMENTS_PER_PERIOD_FORMULA = 'IFS({col}="Monthly",1,{col}="Quarterly",3)'
function appendNewBuildings(workbook: ExcelScript.Workbook, currentExport: LeaseRow[], params: RollForwardParams) {
  const sheet = workbook.getWorksheet(SHEET_MVT_DETAILS)
  if (!sheet) throw new Error(`Sheet "${SHEET_MVT_DETAILS}" not found.`)

  const periodEnd = new Date(params.periodEndDate)
  const periodStart = new Date(Date.UTC(periodEnd.getUTCFullYear(), periodEnd.getUTCMonth(), 1))

  const newBuildings = currentExport.filter((row) => {
    if (row.AssetCategory !== 'Land and buildings') return false
    const commencement = new Date(row.LeaseCommencementDate)
    return commencement >= periodStart && commencement <= periodEnd
  })

  const headerRow = 18
  const firstDataRow = headerRow + 1

  // Clear the previous month's rows plus the total line below them.
  const previousBlock = sheet.getRangeByIndexes(firstDataRow - 1, 0, 200, 12)
  previousBlock.clear(ExcelScript.ClearApplyTo.contents)

  newBuildings.forEach((row, index) => {
    const rowNumber = firstDataRow + index
    sheet.getRangeByIndexes(rowNumber - 1, 0, 1, 11).setValues([[
      row.Key,
      row.Entity,
      row.LeaseDescription,
      row.LeaseCommencementDate,
      row.ReasonablyCertainEndDateSelection,
      row.ReasonablyCertainEndDate,
      row.LeaseDuration,
      row.FixedPayment,
      row.PaymentFrequency,
      row.AssetCategory,
      row.LeasedCapacity,
    ]])
    const divisor = PAYMENTS_PER_PERIOD_FORMULA.replace(/\{col\}/g, `I${rowNumber}`)
    sheet.getRange(`L${rowNumber}`).setFormula(`=H${rowNumber}*G${rowNumber}/${divisor}`)
  })

  if (newBuildings.length > 0) {
    const totalRow = firstDataRow + newBuildings.length
    sheet.getRange(`L${totalRow}`).setFormula(`=SUM(L${firstDataRow}:L${totalRow - 1})`)
  }
}

/**
 * Everything the preparer must look at, returned to Power Automate so it can go
 * straight into the notification instead of being buried in a script log.
 */
function buildReviewReport(workbook: ExcelScript.Workbook, params: RollForwardParams): string {
  const lines: string[] = [`IFRS16 roll-forward to ${params.newMonthLabel} (${params.newPeriodCode})`]

  const movement = workbook.getWorksheet(SHEET_MOVEMENT)
  if (movement) {
    const buildingsCheck = movement.getRange('F49').getValue()
    const vehiclesCheck = movement.getRange('K49').getValue()
    const ok = Number(buildingsCheck) === 0 && Number(vehiclesCheck) === 0
    lines.push(`CHECK row (must be 0): buildings ${buildingsCheck}, vehicles ${vehiclesCheck} — ${ok ? 'OK' : 'NOT TYING, investigate'}`)

    // The Difference column is G minus the YTD build-up, so a non-zero value is
    // exactly the extra plug column L needs for that PowerHouse to tie.
    const plugsNeeded: string[] = []
    for (const [label, firstRow] of [['Buildings', BUILDINGS_FIRST_ROW], ['Vehicles', VEHICLES_FIRST_ROW]] as [string, number][]) {
      const block = movement.getRange(`A${firstRow}:J${firstRow + POWERHOUSE_COUNT - 1}`).getValues()
      block.forEach((row) => {
        const difference = Number(row[9])
        if (difference !== 0) plugsNeeded.push(`${label} ${row[0]}: ${difference > 0 ? '+' : ''}${difference}`)
      })
    }
    lines.push(plugsNeeded.length === 0
      ? 'YTD build-up ties to the 2.9 count for every PowerHouse; no plug adjustment needed.'
      : `Extra plug needed in column L to tie to the 2.9 count — ${plugsNeeded.join(', ')}`)
  }

  lines.push('Column M now carries last month\'s plug; column L is unchanged and is the preparer\'s to set.')
  lines.push(unmappedEntitySummary(workbook))

  return lines.join('\n')
}

/**
 * Entities MDM has not mapped to a PowerHouse drop out of the per-PowerHouse
 * totals without any error, so they are reported rather than left silent.
 */
function unmappedEntitySummary(workbook: ExcelScript.Workbook): string {
  const sheet = workbook.getWorksheet(SHEET_ENTITY_LIST_PBI)
  if (!sheet) return `Sheet "${SHEET_ENTITY_LIST_PBI}" not found — entity mapping not checked.`

  const usedRange = sheet.getUsedRange()
  if (!usedRange) return 'Entity list is empty.'

  const lastRow = usedRange.getRowIndex() + usedRange.getRowCount()
  const rowCount = lastRow - ENTITY_LIST_PBI_FIRST_DATA_ROW + 1
  if (rowCount <= 0) return 'Entity list has no data rows.'

  const values = sheet
    .getRangeByIndexes(ENTITY_LIST_PBI_FIRST_DATA_ROW - 1, 0, rowCount, ENTITY_LIST_PBI_POWERHOUSE_COLUMN)
    .getValues()

  const unmapped: string[] = []
  values.forEach((row) => {
    const code = String(row[ENTITY_LIST_PBI_CODE_COLUMN - 1] ?? '').trim()
    if (!code) return
    const powerhouse = String(row[ENTITY_LIST_PBI_POWERHOUSE_COLUMN - 1] ?? '').trim()
    if (UNMAPPED_POWERHOUSE_VALUES.indexOf(powerhouse) !== -1) {
      unmapped.push(`${code} ${row[ENTITY_LIST_PBI_DESCRIPTION_COLUMN - 1]}`)
    }
  })

  return unmapped.length === 0
    ? `Entity list: ${values.length} rows, all mapped to a PowerHouse.`
    : `Entities without a PowerHouse mapping (fix in MDM): ${unmapped.join(', ')}`
}
