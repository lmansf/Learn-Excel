// Lesson 4.3 · DenialMap_and_DenialsByOwner
// SPOILER: reference solution. Try the task first.

// Query: DenialMap   (click inside tblDenialMap > Data > From Table/Range)
let
    Source = Excel.CurrentWorkbook(){[Name="tblDenialMap"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"DenialReason", type text}, {"RevCycleStage", type text}, {"OwnerTeam", type text}}),
    #"Trimmed Text" = Table.TransformColumns(#"Changed Type",{{"DenialReason", Text.Trim, type text}}),
    #"Capitalized Each Word" = Table.TransformColumns(#"Trimmed Text",{{"DenialReason", Text.Proper, type text}})
in
    #"Capitalized Each Word"

// Query: DenialsByOwner   (right-click Claims > Reference)
let
    Source = Claims,
    #"Filtered Rows" = Table.SelectRows(Source, each ([ClaimStatus] = "Denied")),
    #"Merged Queries" = Table.NestedJoin(#"Filtered Rows", {"DenialReason"}, DenialMap, {"DenialReason"}, "DenialMap", JoinKind.LeftOuter),
    #"Expanded DenialMap" = Table.ExpandTableColumn(#"Merged Queries", "DenialMap", {"OwnerTeam"}, {"OwnerTeam"}),
    #"Grouped Rows" = Table.Group(#"Expanded DenialMap", {"OwnerTeam"}, {
        {"Claims", each Table.RowCount(_), Int64.Type},
        {"BilledDenied", each List.Sum([BilledAmount]), type nullable number}})
in
    #"Grouped Rows"
