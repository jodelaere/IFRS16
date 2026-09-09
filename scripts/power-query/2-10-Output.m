// Query "2/10 Output" — corrected version (template change A).
//
// Paste this over the whole query: Data > Queries & Connections > right-click
// "2/10 Output" > Edit > Home > Advanced Editor > select all > paste > Done.
//
// Only two things changed versus the original:
//   1. ReportingPeriodEnd is read from the named cell instead of the clock.
//   2. Each flag keeps null when its source date is null.
//
// That second point matters more than it looks. Date.IsInCurrentYear(null)
// returns null, and both TRANSFER pivots filter Transfers (OUT) on "(blank)" —
// i.e. on that null. A plain Date.Year(null) = Date.Year(...) comparison would
// return false instead, the "(blank)" filter would match nothing, and the
// transfer figures would change silently. Hence the explicit null guards.

let
    Source = Excel.CurrentWorkbook(){[Name="Table1"]}[Content],
    ReportingPeriodEnd = Date.From(Excel.CurrentWorkbook(){[Name="ReportingPeriodEnd"]}[Content]{0}[Column1]),
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Column1", type text}, {"Entity", type text}, {"Lease description", type text}, {"Cost Center", type text}, {"Local Cost center code", type text}, {"Lease commencement date", type datetime}, {"Purchase option?", type logical}, {"Exercise of Purchase option date", type any}, {"Exercise price of purchase option", Int64.Type}, {"Reasonably certain end date selection", type text}, {"Reasonably certain end date", type datetime}, {"Transfer IN Date", type datetime}, {"Transfer Out Date", type datetime}, {"Lease duration", type number}, {"Fixed payment", type number}, {"Payment frequency", type text}, {"Payment at beginning of period?", type logical}, {"Revision type", type any}, {"Index or Rate", type any}, {"Lease revision frequency (months)", Int64.Type}, {"First revision after (months)", Int64.Type}, {"Reference index/rate date", type any}, {"Provision for dismantling costs", Int64.Type}, {"Status", type text}, {"Last modification status", type text}, {"Asset category", type text}, {"Leased capacity", Int64.Type}, {"Type motor", type text}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"Local Cost center code", "Cost Center", "Purchase option?", "Exercise of Purchase option date", "Exercise price of purchase option", "Payment frequency", "Fixed payment", "Payment at beginning of period?", "Revision type", "Index or Rate", "Lease revision frequency (months)", "First revision after (months)", "Reference index/rate date", "Provision for dismantling costs", "Status", "Last modification status", "Leased capacity", "Type motor", "Lease duration"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Columns",{"Column1", "Entity", "Asset category", "Lease description", "Lease commencement date", "Reasonably certain end date selection", "Reasonably certain end date", "Transfer IN Date", "Transfer Out Date"}),
    #"Split Column - entity codes" = Table.SplitColumn(#"Reordered Columns", "Column1", Splitter.SplitTextByDelimiter("__", QuoteStyle.Csv), {"Column1.1", "Column1.2"}),
    #"Renamed Column - entity code" = Table.RenameColumns(#"Split Column - entity codes",{{"Column1.1", "Entity Code"}, {"Column1.2", "Contract Name"}}),
    #"Changed Type - entity code" = Table.TransformColumnTypes(#"Renamed Column - entity code",{{"Entity Code", Int64.Type}}),
    #"Added Custom - New leases in current year" = Table.AddColumn(#"Changed Type - entity code", "New leases in current year (IN)", each if [Lease commencement date] = null then null else Date.Year([Lease commencement date]) = Date.Year(ReportingPeriodEnd)),
    #"Added Custom - New leases between 1/01 and current month end" = Table.AddColumn(#"Added Custom - New leases in current year", "New leases before and in current month (IN)", each if [Lease commencement date] = null then null else Date.Month([Lease commencement date]) <= Date.Month(ReportingPeriodEnd)),
    #"Changed Type - Reas. Certain End Date" = Table.TransformColumnTypes(#"Added Custom - New leases between 1/01 and current month end",{{"Reasonably certain end date", type date}}),
    #"Changed Type - Reas. Certain End Date - NL-BE" = Table.TransformColumnTypes(#"Changed Type - Reas. Certain End Date", {{"Reasonably certain end date", type date}}, "nl-BE"),
    #"Added Custom - Ended in current year" = Table.AddColumn(#"Changed Type - Reas. Certain End Date - NL-BE", "Ended in current year (OUT)", each if [Reasonably certain end date] = null then null else Date.Year([Reasonably certain end date]) = Date.Year(ReportingPeriodEnd)),
    #"Added Custom - Ended between 1/01 and current month end" = Table.AddColumn(#"Added Custom - Ended in current year", "Ended before and in current month (OUT)", each if [Reasonably certain end date] = null then null else Date.Month([Reasonably certain end date]) <= Date.Month(ReportingPeriodEnd)),
    #"Added Custom - Transfers (IN)" = Table.AddColumn(#"Added Custom - Ended between 1/01 and current month end", "Transfers (IN)", each if [Transfer IN Date] = null then null else Date.Year([Transfer IN Date]) = Date.Year(ReportingPeriodEnd)),
    #"Added Custom - Transfers (OUT)" = Table.AddColumn(#"Added Custom - Transfers (IN)", "Transfers (OUT)", each if [Transfer Out Date] = null then null else Date.Year([Transfer Out Date]) = Date.Year(ReportingPeriodEnd))
in
    #"Added Custom - Transfers (OUT)"
