// Lesson 4.3 · PaidByPayerType
// SPOILER: reference solution. Try the task first.

// Query: PaidByPayerType   (right-click Claims > Reference; Home > Merge Queries)
let
    Source = Claims,
    #"Merged Queries" = Table.NestedJoin(Source, {"PayerID"}, Payers, {"PayerID"}, "Payers", JoinKind.LeftOuter),
    #"Expanded Payers" = Table.ExpandTableColumn(#"Merged Queries", "Payers", {"PayerType"}, {"PayerType"}),
    #"Grouped Rows" = Table.Group(#"Expanded Payers", {"PayerType"}, {
        {"Claims", each Table.RowCount(_), Int64.Type},
        {"Paid", each List.Sum([PaidAmount]), type nullable number}})
in
    #"Grouped Rows"
