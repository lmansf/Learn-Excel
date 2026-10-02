// Lesson 4.3 · AgedPending
// SPOILER: reference solution. Try the task first.

// AgedPending: Pending claims that have waited too long for payment.
// Paste into Data > Get Data > From Other Sources > Blank Query > Home > Advanced Editor.
// It reads your Claims query (task 3), so that query must exist and be named Claims.
let
    Source = Claims,
    AsOf = #date(2025, 12, 31),
    PendingOnly = Table.SelectRows(Source, each [ClaimStatus] = "Pending"),
    AddDaysPending = Table.AddColumn(PendingOnly, "DaysPending", each Duration.Days(AsOf - [SubmitDate]), Int64.Type),
    Aged = Table.SelectRows(AddDaysPending, each [DaysPending] > 90),
    KeepColumns = Table.SelectColumns(Aged, {"ClaimID", "PayerID", "SubmitDate", "BilledAmount", "DaysPending"}),
    Sorted = Table.Sort(KeepColumns, {{"DaysPending", Order.Descending}})
in
    Sorted
