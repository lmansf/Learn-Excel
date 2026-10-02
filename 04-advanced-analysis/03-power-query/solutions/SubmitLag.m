// Lesson 4.3 · SubmitLag
// SPOILER: reference solution. Try the task first.

// Query: SubmitLag   (right-click Claims > Reference; Add Column > Custom Column)
let
    Source = Claims,
    #"Added Custom" = Table.AddColumn(Source, "DaysToSubmit", each Duration.Days([SubmitDate] - [ServiceDate]), Int64.Type),
    #"Filtered Rows" = Table.SelectRows(#"Added Custom", each [DaysToSubmit] > 30)
in
    #"Filtered Rows"
