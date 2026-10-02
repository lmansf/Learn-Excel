// Lesson 4.3 · Budget2025
// SPOILER: reference solution. Try the task first.

// Query: Budget2025   (From Text/CSV: budget_2025_wide.csv)
// Select FY Total > Remove Columns; select FacilityID..Measure > Transform > Unpivot Columns > Unpivot Other Columns;
// rename Attribute to Month and Value to Amount
let
    Source = Csv.Document(File.Contents(DataFolder & "budget_2025_wide.csv"),[Delimiter=",", Columns=18, Encoding=65001, QuoteStyle=QuoteStyle.None]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"FacilityID", type text}, {"Department", type text}, {"LineType", type text}, {"Category", type text}, {"Measure", type text}, {"Jan", Int64.Type}, {"Feb", Int64.Type}, {"Mar", Int64.Type}, {"Apr", Int64.Type}, {"May", Int64.Type}, {"Jun", Int64.Type}, {"Jul", Int64.Type}, {"Aug", Int64.Type}, {"Sep", Int64.Type}, {"Oct", Int64.Type}, {"Nov", Int64.Type}, {"Dec", Int64.Type}, {"FY Total", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"FY Total"}),
    #"Unpivoted Other Columns" = Table.UnpivotOtherColumns(#"Removed Columns", {"FacilityID", "Department", "LineType", "Category", "Measure"}, "Attribute", "Value"),
    #"Renamed Columns" = Table.RenameColumns(#"Unpivoted Other Columns",{{"Attribute", "Month"}, {"Value", "Amount"}})
in
    #"Renamed Columns"
