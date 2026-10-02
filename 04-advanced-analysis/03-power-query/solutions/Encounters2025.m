// Lesson 4.3 · Encounters2025
// SPOILER: reference solution. Try the task first.

// Query: Encounters2025   (From Text/CSV: encounters_2025.csv)
let
    Source = Csv.Document(File.Contents(DataFolder & "encounters_2025.csv"),[Delimiter=",", Columns=10, Encoding=65001, QuoteStyle=QuoteStyle.None]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"EncounterID", type text}, {"PatientID", type text}, {"EncounterType", type text}, {"FacilityID", type text}, {"DeptID", type text}, {"AdmitDateTime", type datetime}, {"DischargeDateTime", type datetime}, {"PrimaryDxCode", type text}, {"PayerID", type text}, {"TotalCharges", type number}})
in
    #"Changed Type"
