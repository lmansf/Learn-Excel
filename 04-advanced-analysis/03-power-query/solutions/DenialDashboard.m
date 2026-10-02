// Lesson 4.3 · DenialDashboard
// SPOILER: reference solution. Try the task first.

// Query: DenialDashboard   (right-click Claims > Reference)
let
    Source = Claims,
    #"Filtered Rows" = Table.SelectRows(Source, each ([ClaimStatus] = "Denied")),
    #"Merged Queries" = Table.NestedJoin(#"Filtered Rows", {"PayerID"}, Payers, {"PayerID"}, "Payers", JoinKind.LeftOuter),
    #"Expanded Payers" = Table.ExpandTableColumn(#"Merged Queries", "Payers", {"PayerType"}, {"PayerType"}),
    #"Grouped Rows" = Table.Group(#"Expanded Payers", {"PayerType", "DenialReason"}, {
        {"DeniedClaims", each Table.RowCount(_), Int64.Type},
        {"DeniedBilled", each List.Sum([BilledAmount]), type nullable number}}),
    #"Sorted Rows" = Table.Sort(#"Grouped Rows",{{"DeniedBilled", Order.Descending}})
in
    #"Sorted Rows"
