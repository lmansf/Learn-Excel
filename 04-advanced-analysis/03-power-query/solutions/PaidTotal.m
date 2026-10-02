// Lesson 4.3 · PaidTotal
// SPOILER: reference solution. Try the task first.

// Right-click Claims > Reference, rename the new query PaidTotal,
// select the PaidAmount column > Transform > Statistics > Sum. The preview shows one number.
let
    Source = Claims,
    #"Calculated Sum" = List.Sum(Source[PaidAmount])
in
    #"Calculated Sum"
// Alternative: load Claims to a sheet and type =SUM(Claims[PaidAmount]) in any empty cell.
