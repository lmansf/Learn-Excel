// Lesson 4.3 · Claims
// SPOILER: reference solution. Try the task first.

// Query: Claims   (Data > Get Data > From File > From Folder > pick claims_monthly > Combine & Transform Data,
// choose the first file as the sample, OK, then rename the query from claims_monthly to Claims)
// Excel also creates a "Helper Queries" group (Sample File, Parameter1, Transform Sample File, Transform File).
// The Source line below uses the DataFolder parameter (Guide section 12); the generated code has your full path.
// The helper query Sample File repeats the folder path in its own Source step, so make the same change there:
//     Source = Folder.Files(DataFolder & "claims_monthly"),
// (On a Mac, keep full paths such as "/Users/you/PQ/data/claims_monthly" instead of the parameter.)
// Step names can differ slightly between Excel versions. Optional hardening (Guide section 11): insert
//     #"CSV Only" = Table.SelectRows(Source, each Text.Lower([Extension]) = ".csv"),
// after Source (and make the next step read #"CSV Only") so stray files, such as a Mac .DS_Store file,
// are never combined.
let
    Source = Folder.Files(DataFolder & "claims_monthly"),
    #"Filtered Hidden Files1" = Table.SelectRows(Source, each [Attributes]?[Hidden]? <> true),
    #"Invoke Custom Function1" = Table.AddColumn(#"Filtered Hidden Files1", "Transform File", each #"Transform File"([Content])),
    #"Renamed Columns1" = Table.RenameColumns(#"Invoke Custom Function1", {"Name", "Source.Name"}),
    #"Removed Other Columns1" = Table.SelectColumns(#"Renamed Columns1", {"Source.Name", "Transform File"}),
    #"Expanded Table Column1" = Table.ExpandTableColumn(#"Removed Other Columns1", "Transform File", Table.ColumnNames(#"Transform File"(#"Sample File"))),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded Table Column1",{{"Source.Name", type text}, {"ClaimID", type text}, {"EncounterID", type text}, {"PatientID", type text}, {"PayerID", type text}, {"ServiceDate", type date}, {"SubmitDate", type date}, {"BilledAmount", type number}, {"AllowedAmount", type number}, {"PatientResponsibility", type number}, {"PaidAmount", type number}, {"ClaimStatus", type text}, {"DenialReason", type text}, {"PaidDate", type date}})
in
    #"Changed Type"
