/**
 * IFRS16 Monthly Roll-Forward — Office Script
 *
 * Runs against the "Input Board Pack.xlsx" for the new period (already copied
 * forward from the previous period's file by the Power Automate flow before
 * this script is invoked). See docs/automation-design.md for the full design,
 * assumptions, and open validation points — several ranges below are marked
 * CONFIRM-ME because they could not be verified against the live workbook
 * from this session and must be checked before this script is trusted on the
 * real file.
 *
 * Invoked from Power Automate via the "Run script" action. Power Automate
 * supplies anaplan29Rows / anaplan210Rows as arrays of LeaseRow (e.g. built
 * from "List rows present in a table" on the two Anaplan export files).
 */

interface LeaseRow {
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
  LeaseDuration: string
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

interface RollForwardParams {
  newMonthLabel: string // e.g. "September 2026"
  previousMonthLabel: string // e.g. "August 2026"
  newPeriodCode: string // e.g. "P09"
  previousPeriodCode: string // e.g. "P08"
  /** First day of the new reporting period, used to detect "new this month" leases. */
  periodStartDate: string // ISO date, e.g. "2026-09-01"
  periodEndDate: string // ISO date, e.g. "2026-09-30"
}

const LEASE_ROW_HEADERS: (keyof LeaseRow)[] = [
  'Entity', 'LeaseDescription', 'CostCenter', 'LocalCostCenterCode',
  'LeaseCommencementDate', 'PurchaseOption', 'ExerciseOfPurchaseOptionDate',
  'ExercisePriceOfPurchaseOption', 'ReasonablyCertainEndDateSelection',
  'ReasonablyCertainEndDate', 'TransferInDate', 'TransferOutDate',
  'LeaseDuration', 'FixedPayment', 'PaymentFrequency',
  'PaymentAtBeginningOfPeriod', 'RevisionType', 'IndexOrRate',
  'LeaseRevisionFrequencyMonths', 'FirstRevisionAfterMonths',
  'ReferenceIndexRateDate', 'ProvisionForDismantlingCosts', 'Status',
  'LastModificationStatus', 'AssetCategory', 'LeasedCapacity', 'TypeMotor',
]

function main(
  workbook: ExcelScript.Workbook,
  anaplan29Rows: LeaseRow[],
  anaplan210Rows: LeaseRow[],
  params: RollForwardParams
) {
  overwriteInputTab(workbook, '2.9 input', anaplan29Rows)
  overwriteInputTab(workbook, '2.10 input', anaplan210Rows)

  refreshEverything(workbook)

  // CONFIRM-ME: whether the header roll is pure text or also needs formula
  // shifts. See docs/automation-design.md, open point 5.
  rollMovementScheduleHeaders(workbook, params)

  appendNewLeases(workbook, anaplan210Rows, params)

  refreshEverything(workbook)
}

function overwriteInputTab(workbook: ExcelScript.Workbook, sheetName: string, rows: LeaseRow[]) {
  const sheet = workbook.getWorksheet(sheetName)
  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found — confirm exact tab name before running.`)
  }

  const usedRange = sheet.getUsedRange()
  const headerRow = 1 // CONFIRM-ME: assumes row 1 = headers, data starts row 2
  const lastDataRow = Math.max(usedRange ? usedRange.getRowCount() : headerRow, headerRow)

  // Clear existing data below the header row, keep formatting.
  if (lastDataRow > headerRow) {
    sheet
      .getRangeByIndexes(headerRow, 0, lastDataRow - headerRow, LEASE_ROW_HEADERS.length)
      .clear(ExcelScript.ClearApplyTo.contents)
  }

  const values = rows.map((row) => LEASE_ROW_HEADERS.map((key) => row[key] ?? ''))
  if (values.length > 0) {
    sheet
      .getRangeByIndexes(headerRow, 0, values.length, LEASE_ROW_HEADERS.length)
      .setValues(values as (string | number)[][])
  }
}

function refreshEverything(workbook: ExcelScript.Workbook) {
  workbook.getPivotTables().forEach((pivotTable) => pivotTable.refresh())
  workbook.getApplication().calculate(ExcelScript.CalculationType.fullRebuild)
}

function rollMovementScheduleHeaders(workbook: ExcelScript.Workbook, params: RollForwardParams) {
  const sheet = workbook.getWorksheet('Movement schedule')
  if (!sheet) throw new Error('Sheet "Movement schedule" not found.')

  // G1 / G17 are already formula-driven (=K34) per the sample workbook, so
  // the "current month" header does not need to be touched here.

  // CONFIRM-ME: column letters below are taken from the August 2026 sample
  // and may need adjusting once the exact roll mechanic is confirmed.
  const previousMonthHeaderCell = 'N1' // was "July 2026" in the sample
  const movementColumnHeaderCell = 'O1' // was "mvt P08" in the sample

  sheet.getRange(previousMonthHeaderCell).setValue(params.previousMonthLabel)
  sheet.getRange(movementColumnHeaderCell).setValue(`mvt ${params.newPeriodCode}`)

  // Plug columns (K "Plug 08 2026" / L "Plug 07 2026" in the sample) are left
  // untouched intentionally — see docs/automation-design.md: these are manual
  // correcting entries made by the preparer and must not be auto-populated.
}

function appendNewLeases(workbook: ExcelScript.Workbook, currentExport: LeaseRow[], params: RollForwardParams) {
  const sheet = workbook.getWorksheet('Mvt Schedule Details')
  if (!sheet) throw new Error('Sheet "Mvt Schedule Details" not found.')

  const start = new Date(params.periodStartDate)
  const end = new Date(params.periodEndDate)

  const newLeases = currentExport.filter((row) => {
    const commencement = new Date(row.LeaseCommencementDate)
    return commencement >= start && commencement <= end
  })

  const buildings = newLeases.filter((r) => r.AssetCategory === 'Land and buildings')
  const vehicles = newLeases.filter((r) => r.AssetCategory === 'Vehicles')

  appendToNewLeaseTable(sheet, 'BUILDINGS - NEW', buildings)
  // CONFIRM-ME: "VEHICLES - NEW" table location/columns were not visible in
  // the sample read — this call assumes it mirrors the BUILDINGS table with
  // the same column layout, positioned to the right or below it.
  appendToNewLeaseTable(sheet, 'VEHICLES - NEW', vehicles)
}

/**
 * Appends rows to the "<CATEGORY> - NEW" table in Mvt Schedule Details.
 * Column layout observed for BUILDINGS - NEW:
 *   Key | Entity | Lease description | Lease commencement date |
 *   Reasonably certain end date selection | Reasonably certain end date |
 *   Lease duration | Fixed payment | Payment frequency | Asset category |
 *   Leased capacity | Lease Liability (=FixedPayment * LeasedCapacity)
 */
function appendToNewLeaseTable(sheet: ExcelScript.Worksheet, tableName: string, rows: LeaseRow[]) {
  const table = sheet.getTables().find((t) => t.getName() === tableName)
  if (!table) {
    console.log(`Table "${tableName}" not found by name — confirm it is defined as an Excel Table, or adjust appendToNewLeaseTable to use a fixed range instead.`)
    return
  }

  for (const row of rows) {
    table.addRow(-1, [
      `${row.Entity.split(' - ')[0]}__${row.LeaseDescription}`, // CONFIRM-ME: key format guessed from sample data
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
      row.FixedPayment * row.LeasedCapacity,
    ])
  }
}
