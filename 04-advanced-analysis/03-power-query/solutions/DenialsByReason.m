// Lesson 4.3 · DenialsByReason
// SPOILER: reference solution. Try the task first.

// Query: DenialsByReason, Grouped Rows step edited (gear icon > Advanced > Add aggregation)
let
    Source = Claims,
    #"Filtered Rows" = Table.SelectRows(Source, each ([ClaimStatus] = "Denied")),
    #"Grouped Rows" = Table.Group(#"Filtered Rows", {"DenialReason"}, {
        {"Claims", each Table.RowCount(_), Int64.Type},
        {"BilledDenied", each List.Sum([BilledAmount]), type nullable number},
        {"AvgBilled", each List.Average([BilledAmount]), type nullable number}}),
    #"Sorted Rows" = Table.Sort(#"Grouped Rows",{{"AvgBilled", Order.Descending}})
in
    #"Sorted Rows"
