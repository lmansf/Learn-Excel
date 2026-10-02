// Lesson 4.3 · Unbilled
// SPOILER: reference solution. Try the task first.

// Query: Unbilled   (Home > Merge Queries > Merge Queries as New:
// top = Encounters2025, bottom = Claims, click EncounterID in both, Join Kind = Left Anti)
let
    Source = Table.NestedJoin(Encounters2025, {"EncounterID"}, Claims, {"EncounterID"}, "Claims", JoinKind.LeftAnti),
    #"Removed Columns" = Table.RemoveColumns(Source, {"Claims"})
in
    #"Removed Columns"
