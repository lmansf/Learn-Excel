// Lesson 4.3 · Payers
// SPOILER: reference solution. Try the task first.

// Query: Payers   (From Text/CSV: payers.csv)
let
    Source = Csv.Document(File.Contents(DataFolder & "payers.csv"),[Delimiter=",", Columns=6, Encoding=65001, QuoteStyle=QuoteStyle.None]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"PayerID", type text}, {"PayerName", type text}, {"PayerType", type text}, {"AvgAllowedPctOfCharges", type number}, {"AvgDaysToPay", Int64.Type}, {"TimelyFilingDays", Int64.Type}})
in
    #"Changed Type"
