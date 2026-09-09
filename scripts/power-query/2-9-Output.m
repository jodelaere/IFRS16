// Query "2/9 Output" — corrected version (template change A).
//
// Paste this over the whole query: Data > Queries & Connections > right-click
// "2/9 Output" > Edit > Home > Advanced Editor > select all > paste > Done.
//
// Changes versus the original, same two as the 2/10 query:
//   1. ReportingPeriodEnd is read from the named cell instead of the clock.
//   2. Each flag keeps null when its source date is null.
//
// Note the "after current month end" flag: the original compared against
// Date.Month(DateTime.LocalNow()) with no -1, because LocalNow sat in the month
// *after* the reporting month. ReportingPeriodEnd sits *in* it, so the
// comparison becomes a strict greater-than against the same month.

let
    Source = Excel.CurrentWorkbook(){[Name="Table2.9"]}[Content],
    ReportingPeriodEnd = Date.From(Excel.CurrentWorkbook(){[Name="ReportingPeriodEnd"]}[Content]{0}[Column1]),
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Column1", type text}, {"Entity", type text}, {"Lease description", type text}, {"Cost Center", type text}, {"Local Cost center code", type text}, {"Lease commencement date", type datetime}, {"Purchase option?", type logical}, {"Exercise of Purchase option date", type any}, {"Exercise price of purchase option", Int64.Type}, {"Reasonably certain end date selection", type text}, {"Reasonably certain end date", type datetime}, {"Transfer IN Date", type datetime}, {"Transfer Out Date", type datetime}, {"Lease duration", type number}, {"Fixed payment", type number}, {"Payment frequency", type text}, {"Payment at beginning of period?", type logical}, {"Revision type", type any}, {"Index or Rate", type any}, {"Lease revision frequency (months)", Int64.Type}, {"First revision after (months)", Int64.Type}, {"Reference index/rate date", type any}, {"Provision for dismantling costs", Int64.Type}, {"Status", type text}, {"Last modification status", type text}, {"Asset category", type text}, {"Leased capacity", Int64.Type}, {"Type motor", type text}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"Cost Center", "Local Cost center code", "Purchase option?", "Exercise of Purchase option date", "Exercise price of purchase option", "Fixed payment", "Payment frequency", "Payment at beginning of period?", "Revision type", "Index or Rate", "Lease revision frequency (months)", "First revision after (months)", "Reference index/rate date", "Provision for dismantling costs", "Status", "Last modification status", "Leased capacity", "Type motor", "Lease duration"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Columns",{"Column1", "Entity", "Asset category", "Lease description", "Lease commencement date", "Reasonably certain end date selection", "Reasonably certain end date", "Transfer IN Date", "Transfer Out Date"}),
    #"Split Column by Delimiter - Entity code vs name" = Table.SplitColumn(#"Reordered Columns", "Entity", Splitter.SplitTextByEachDelimiter({"-"}, QuoteStyle.Csv, false), {"Entity.1", "Entity.2"}),
    #"Renamed Columns - Entity code and name" = Table.RenameColumns(#"Split Column by Delimiter - Entity code vs name",{{"Entity.1", "Entity code"}, {"Entity.2", "Entity name"}}),
    #"Added Custom - lease IN in current year" = Table.AddColumn(#"Renamed Columns - Entity code and name", "Lease commencement in current year", each if [Lease commencement date] = null then null else Date.Year([Lease commencement date]) = Date.Year(ReportingPeriodEnd)),
    #"Added Custom - lease IN after current month" = Table.AddColumn(#"Added Custom - lease IN in current year", "Lease commencement after current month end", each if [Lease commencement date] = null then null else Date.Month([Lease commencement date]) > Date.Month(ReportingPeriodEnd)),
    #"Added Custom - lease IN - to be excluded" = Table.AddColumn(#"Added Custom - lease IN after current month", "Lease IN - to be excluded", each [Lease commencement in current year]=true and [Lease commencement after current month end] =true),
    #"Filtered Rows - Lease IN excluded" = Table.SelectRows(#"Added Custom - lease IN - to be excluded", each ([#"Lease IN - to be excluded"] = false)),
    #"Added Custom - lease OUT in current year" = Table.AddColumn(#"Filtered Rows - Lease IN excluded", "lease OUT in current year", each if [Reasonably certain end date] = null then null else Date.Year([Reasonably certain end date]) = Date.Year(ReportingPeriodEnd)),
    #"Added Custom - lease OUT between 1/01 and current month end" = Table.AddColumn(#"Added Custom - lease OUT in current year", "lease OUT between 1/01 and current month end", each if [Reasonably certain end date] = null then null else Date.Month([Reasonably certain end date]) <= Date.Month(ReportingPeriodEnd)),
    #"Added Custom - lease OUT to be excluded" = Table.AddColumn(#"Added Custom - lease OUT between 1/01 and current month end", "Lease OUT - to be excluded", each [lease OUT in current year] = true and [#"lease OUT between 1/01 and current month end"] = true),
    #"Filtered Rows - lease OUT excluded" = Table.SelectRows(#"Added Custom - lease OUT to be excluded", each ([#"Lease OUT - to be excluded"] = false) and ([Transfer Out Date] = null)),
    #"Grouped Rows" = Table.Group(#"Filtered Rows - lease OUT excluded", {"Entity code", "Asset category"}, {{"Number of lease contracts per entity", each Table.RowCount(_), Int64.Type}}),
    #"Changed Type - Entity Code" = Table.TransformColumnTypes(#"Grouped Rows",{{"Entity code", Int64.Type}})
in
    #"Changed Type - Entity Code"
