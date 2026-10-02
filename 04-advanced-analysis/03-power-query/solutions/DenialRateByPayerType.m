// Lesson 4.3 · DenialRateByPayerType
// SPOILER: reference solution. Try the task first.

// Query: DenialRateByPayerType   (right-click Claims > Reference)
let
    Source = Claims,
    #"Merged Queries" = Table.NestedJoin(Source, {"PayerID"}, Payers, {"PayerID"}, "Payers", JoinKind.LeftOuter),
    #"Expanded Payers" = Table.ExpandTableColumn(#"Merged Queries", "Payers", {"PayerType"}, {"PayerType"}),
    #"Added Conditional Column" = Table.AddColumn(#"Expanded Payers", "IsDenied", each if [ClaimStatus] = "Denied" then 1 else 0, Int64.Type),
    #"Grouped Rows" = Table.Group(#"Added Conditional Column", {"PayerType"}, {
        {"AllClaims", each Table.RowCount(_), Int64.Type},
        {"Denied", each List.Sum([IsDenied]), type nullable number}}),
    #"Added Custom" = Table.AddColumn(#"Grouped Rows", "DenialRate", each [Denied] / [AllClaims], Percentage.Type),
    #"Sorted Rows" = Table.Sort(#"Added Custom",{{"DenialRate", Order.Descending}})
in
    #"Sorted Rows"
