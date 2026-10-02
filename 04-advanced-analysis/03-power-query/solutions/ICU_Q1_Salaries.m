// Lesson 4.3 · ICU_Q1_Salaries
// SPOILER: reference solution. Try the task first.

// Query: Budget2025 after task 13: the task 12 query plus one step at the end. Select Department >
// Transform > Split Column > By Delimiter > --Custom-- " - " (space hyphen space) > Split at: Left-most delimiter.
// The dialog names the parts Department.1 and Department.2 (and adds a Changed Type step); typing the final
// names in the step, as here, splits and renames at once.
let
    Source = Csv.Document(File.Contents(DataFolder & "budget_2025_wide.csv"),[Delimiter=",", Columns=18, Encoding=65001, QuoteStyle=QuoteStyle.None]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"FacilityID", type text}, {"Department", type text}, {"LineType", type text}, {"Category", type text}, {"Measure", type text}, {"Jan", Int64.Type}, {"Feb", Int64.Type}, {"Mar", Int64.Type}, {"Apr", Int64.Type}, {"May", Int64.Type}, {"Jun", Int64.Type}, {"Jul", Int64.Type}, {"Aug", Int64.Type}, {"Sep", Int64.Type}, {"Oct", Int64.Type}, {"Nov", Int64.Type}, {"Dec", Int64.Type}, {"FY Total", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"FY Total"}),
    #"Unpivoted Other Columns" = Table.UnpivotOtherColumns(#"Removed Columns", {"FacilityID", "Department", "LineType", "Category", "Measure"}, "Attribute", "Value"),
    #"Renamed Columns" = Table.RenameColumns(#"Unpivoted Other Columns",{{"Attribute", "Month"}, {"Value", "Amount"}}),
    #"Split Column by Delimiter" = Table.SplitColumn(#"Renamed Columns", "Department", Splitter.SplitTextByEachDelimiter({" - "}, QuoteStyle.Csv, false), {"CostCenter", "DeptName"})
in
    #"Split Column by Delimiter"

// Query: ICU_Q1_Salaries   (right-click Budget2025 > Reference)
let
    Source = Budget2025,
    #"Filtered Rows" = Table.SelectRows(Source, each [DeptName] = "Intensive Care Unit"
        and [Category] = "Salaries & Wages" and [Measure] = "Actual"
        and List.Contains({"Jan", "Feb", "Mar"}, [Month])),
    #"Grouped Rows" = Table.Group(#"Filtered Rows", {"DeptName"}, {{"Q1Actual", each List.Sum([Amount]), type nullable number}})
in
    #"Grouped Rows"
