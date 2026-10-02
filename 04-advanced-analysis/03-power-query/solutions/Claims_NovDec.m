// Lesson 4.3 · Claims_NovDec
// SPOILER: reference solution. Try the task first.

// Import claims_2025_11.csv and claims_2025_12.csv the same way as task 1, then
// Home > Append Queries > Append Queries as New > Two tables > claims_2025_11 + claims_2025_12
// Query: Claims_NovDec
let
    Source = Table.Combine({claims_2025_11, claims_2025_12})
in
    Source
