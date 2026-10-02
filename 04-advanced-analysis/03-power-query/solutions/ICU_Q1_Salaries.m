// Lesson 4.3 · ICU_Q1_Salaries
// SPOILER: reference solution. Try the task first.

// Budget2025 gets one more step after #"Renamed Columns" (Transform > Split Column > By Delimiter,
// Custom delimiter " - " (space hyphen space), Split at: Left-most delimiter), and its "in" line now returns it.
// The dialog names the parts Department.1 and Department.2; typing the final names here splits and renames at once.
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
