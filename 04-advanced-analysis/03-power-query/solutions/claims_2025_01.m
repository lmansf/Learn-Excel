// Lesson 4.3 · claims_2025_01
// SPOILER: reference solution. Try the task first.

// Query: claims_2025_01
// Data > Get Data > From File > From Text/CSV > pick the file > Transform Data (or Load)
let
    Source = Csv.Document(File.Contents("C:\PQ\data\claims_monthly\claims_2025_01.csv"),[Delimiter=",", Columns=13, Encoding=65001, QuoteStyle=QuoteStyle.None]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"ClaimID", type text}, {"EncounterID", type text}, {"PatientID", type text}, {"PayerID", type text}, {"ServiceDate", type date}, {"SubmitDate", type date}, {"BilledAmount", type number}, {"AllowedAmount", type number}, {"PatientResponsibility", type number}, {"PaidAmount", type number}, {"ClaimStatus", type text}, {"DenialReason", type text}, {"PaidDate", type date}})
in
    #"Changed Type"
