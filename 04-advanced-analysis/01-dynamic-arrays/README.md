# Lesson 4.1 · Dynamic Arrays: FILTER, SORT, UNIQUE & More

> **Level:** Advanced · **Time:** about 125 minutes · **Workbook:** [`4.1-dynamic-arrays.xlsx`](4.1-dynamic-arrays.xlsx)
> **Data:** 2,000 encounters sampled from Bluestone Health System's 2025 activity at all four facilities, with facility, department, attending, diagnosis category, and payer names already joined in. Also includes the provider roster (147) and the department list (31, with StaffedBeds 0 for departments that have no inpatient beds). Column definitions are in the [data dictionary](../../data/README.md#encounterscsv).

A case manager asks for *"every Cedar Ridge patient who stayed a week or longer, longest first."* A revenue-cycle lead
wants *"each payer type's encounter count, biggest first."* In older Excel, each request meant filtering, copying,
pasting, and removing duplicates by hand, and the result was a static snapshot that went stale with the next data
refresh. **Dynamic array** functions answer these questions with one formula that returns a whole list and recalculates
the moment the data changes. In this lesson you build live lists, a long-stay worklist, and a one-formula unit
leaderboard from 2,000 Bluestone encounters.

## What you'll learn

- Understand spilling, the # spill reference, and #SPILL! errors
- Extract lists with UNIQUE, FILTER, SORT, and SORTBY
- Generate sequences with SEQUENCE and reshape with TAKE, DROP, CHOOSECOLS, VSTACK, and HSTACK
- Combine functions into one-formula reports

## 📖 Guide

The examples use three Excel Tables from the lesson workbook: **tblEncounters** (Encounters sheet), **tblProviders**
(Providers sheet), and **tblDepartments** (Departments sheet). Structured references such as
`tblEncounters[TotalCharges]` ([Lesson 3.1](../../03-data-analysis/01-tables-named-ranges/README.md)) are the natural
partner for dynamic arrays, because a Table grows when you paste in new rows and every formula that refers to it sees
the new rows automatically.

> 📋 **Version check:** this lesson needs **Microsoft 365** (Windows, Mac, or Excel for the web) or **Excel 2024**.
> Excel 2021 has FILTER, SORT, SORTBY, UNIQUE, SEQUENCE, XLOOKUP, and LET, but not VSTACK, HSTACK, TAKE, DROP, or
> CHOOSECOLS. Section 13 has the full table.

### 1. One formula, many results: spilling

Click cell S6 on the Workspace sheet (the **Guide examples** area, or any empty cell with room below it) and type:

```
=UNIQUE(tblEncounters[EncounterType])
```

Press **Enter** and four values appear, one per row: Inpatient, Emergency, Outpatient, Observation. You typed one
formula, but Excel returned an **array** (a list of values) and **spilled** it into the cells below. The cell where you
typed the formula is the **anchor cell**, and the block of cells the results fill is the **spill range**.

What you see on the sheet tells you what's going on:

- A thin **blue border** surrounds the spill range whenever you select a cell inside it.
- Click any spilled cell below the anchor and the formula bar shows the formula **grayed out**. Only the anchor holds
  the formula, so that's the only cell you can edit.
- Delete the anchor and the whole spill disappears. Change the data and the spill grows or shrinks to fit.

| | Legacy array formula (Excel 2019 and earlier) | Dynamic array formula (Microsoft 365, Excel 2021+) |
|---|---|---|
| How you enter it | Select the output range first, then **Ctrl + Shift + Enter** (Mac: **⌘ + Shift + Return**) | Type in one cell, press **Enter** |
| How it looks | Curly braces in the formula bar: `{=…}` | No braces |
| Result size | Fixed. Too small cuts results off, too big shows `#N/A` | Resizes itself every time it recalculates |
| Editing | Must select the whole array to change it | Edit the anchor only |

> ⚠️ A spilled list needs empty cells to land in. That is why every **answer cell** on the Practice sheet needs a
> formula that returns **one** value. A spill there would run into the Check column. Lists go on the **Workspace**
> sheet, where each exercise has a yellow anchor cell with empty space below it.

### 2. Referring to a whole spill: the `#` operator

Put a `#` after the anchor's address to mean "the entire spill range that starts here." If the payer list spills from
B6, then:

| Formula | Returns |
|---|---|
| `=B6#` | The whole list again (it spills too) |
| `=ROWS(B6#)` | How many items the list holds |
| `=SORT(B6#)` | The same list sorted A to Z |
| `=COUNTIFS(tblEncounters[Payer], B6#)` | One count per payer, spilling next to the list |
| `=Workspace!B6#` | The same reference from another sheet |

The `#` is the **spill range operator**, and a reference such as `B6#` is a **spill reference**. Use a spill reference
instead of a fixed range like `B6:B13`. When a new payer appears in the data, the list in B6 grows to B14, and every
formula that uses `B6#` grows with it. You don't have to type the `#`: while writing a formula, select the spill range
with the mouse and Excel writes `B6#` for you.

Spill references also work outside formulas:

- **Drop-down lists.** In **Data → Data Validation**, choose **Allow: List** and enter `=Workspace!$B$6#` as the
  **Source**. The drop-down then always offers the current payer list.
- **Charts.** In Microsoft 365 and Excel 2024, a chart built from a spill range resizes by itself when the spill grows
  or shrinks. In Excel 2021, chart the spill through a name instead. Choose **Formulas → Name Manager → New**, name it
  (for example, `MonthCounts`), and set **Refers to** `=Workspace!$G$6#`. Then choose **Chart Design → Select Data**,
  edit the series, and set **Series values** to `=Workspace!MonthCounts`. A chart series needs the sheet (or file)
  name in front of a defined name.

### 3. When spilling fails: `#SPILL!` and its relatives

Excel never overwrites a cell to make room for a spill. When it can't spill, the anchor shows `#SPILL!`. Select the
anchor and click the warning icon (a yellow diamond with an exclamation mark) to read the reason:

| Message in the warning menu | What's wrong | Fix |
|---|---|---|
| **Spill range isn't blank** | Something sits in the way: a value, a stray space, even a formula that returns `""` | Choose **Select Obstructing Cells**, then delete or move what's there |
| **Spill range has merged cells** | The results would land in merged cells | Unmerge them (**Home → Merge & Center → Unmerge Cells**) |
| **Spill range in table** | The formula is inside an Excel Table, and Tables don't allow spills | Move the formula outside the Table, or convert the Table to a range |
| **Spill range is too big** | The result would run past the edge of the worksheet, often from a whole-column reference like `=A:A+1` typed below row 1 | Refer to a Table column or a fixed range instead |
| **Spill range is unknown** | The result size changes between calculation passes, as in `=SEQUENCE(RANDBETWEEN(1,100))` | Base the size on something stable instead of a volatile function |

Two related errors come from the functions themselves rather than from the sheet:

- `#CALC!` means a function produced an empty result, most often a FILTER where no row matches. Section 7 shows how
  to replace it with your own message.
- `#VALUE!` from FILTER usually means the condition array isn't the same height as the data you're filtering.

> ⚠️ Typing into a cell inside a spill range blocks the spill, and the anchor turns into `#SPILL!`. Press **Ctrl + Z**
> (Mac: **⌘ + Z**) to undo the stray entry.

### 4. The `@` sign: implicit intersection

Before dynamic arrays, typing `=tblEncounters[TotalCharges]` into a single cell didn't return the whole column. Excel
quietly returned the one value in the **same row** as the formula. That old behavior is called **implicit
intersection**, and Microsoft 365 now writes it with the `@` operator:

| Formula | In Microsoft 365 |
|---|---|
| `=tblEncounters[TotalCharges]` | Spills all 2,000 charges |
| `=@tblEncounters[TotalCharges]` | Returns only the charge in the formula's own row (or `#VALUE!` if that row isn't in the Table) |
| `=[@TotalCharges]` (inside the Table) | The familiar "this row" reference from Lesson 3.1. It is the same idea |

You'll mostly meet `@` when you open a workbook built in an older version. Excel adds `@` to any formula that might
now return several values, so the old workbook keeps calculating exactly as it did. Leave those `@` signs alone unless
you deliberately want the formula to spill.

### 5. UNIQUE: list each value once

```
=UNIQUE(array, [by_col], [exactly_once])
```

| Argument | Meaning | Default |
|---|---|---|
| `array` | The column, columns, or rows to de-duplicate | Required |
| `by_col` | `TRUE` compares columns instead of rows (for data laid out sideways) | `FALSE` |
| `exactly_once` | `TRUE` returns only values that appear **exactly once** | `FALSE` (every distinct value) |

UNIQUE returns values in the order they **first appear** in the data. Wrap it in SORT when you want them in order.

To set a later optional argument without changing an earlier one, **leave the earlier one empty**. In
`UNIQUE(tblEncounters[Attending],,TRUE)` the two commas in a row skip `by_col`, so Excel uses its default (`FALSE`) and
`TRUE` goes to `exactly_once`. The same trick works in every function in this lesson.

Here's what UNIQUE does with the lesson data:

| Formula | Result |
|---|---|
| `=UNIQUE(tblEncounters[EncounterType])` | 4 rows: Inpatient, Emergency, Outpatient, Observation |
| `=SORT(UNIQUE(tblEncounters[DxCategory]))` | The 17 diagnosis categories, A to Z |
| `=UNIQUE(tblEncounters[[EncounterType]:[Facility]])` | 10 rows × 2 columns: every encounter type and facility pair that occurs |
| `=UNIQUE(tblEncounters[Attending],,TRUE)` | The 5 attendings who appear on exactly one encounter |
| `=ROWS(UNIQUE(tblEncounters[Attending]))` | 142, the number of distinct attendings |

Give UNIQUE **several columns** and it compares whole rows. `tblEncounters[[EncounterType]:[Facility]]` is the
structured reference for "the EncounterType through Facility columns," so the result lists each type and facility
combination once. For columns that aren't next to each other, glue them together first with HSTACK (section 10):
`=UNIQUE(HSTACK(tblEncounters[Facility], tblEncounters[PayerType]))`.

**Counting distinct values** used to take an obscure `SUMPRODUCT(1/COUNTIF(…))` trick. Now it's
`=ROWS(UNIQUE(range))`. ROWS counts how many rows the array has and returns one number, so the formula fits in a single
cell.

> ⚠️ **Blanks become zeros.** Readmit30 is empty for every non-inpatient encounter, so
> `=UNIQUE(tblEncounters[Readmit30])` returns N, **0**, Y. The 0 stands for the empty cells. Filter them out first:
> `=UNIQUE(FILTER(tblEncounters[Readmit30], tblEncounters[Readmit30]<>""))`.

> ⚠️ UNIQUE ignores case, so "Medicare" and "MEDICARE" count as one value. Stray spaces still count as differences, so
> clean the data first (Lesson 3.3).

### 6. SORT and SORTBY

```
=SORT(array, [sort_index], [sort_order], [by_col])
=SORTBY(array, by_array1, [sort_order1], [by_array2, sort_order2], …)
```

`sort_order` is `1` for ascending (the default) or `-1` for descending. The two functions differ in **how you name the
sort key**:

| | SORT | SORTBY | Data → Sort (Lesson 1.6) |
|---|---|---|---|
| Sort key | A column **number** inside `array` | Any range or array of the same height, **inside or outside** the result | Columns you pick in a dialog |
| Several keys | `{2,5}` with orders `{1,-1}` | Add more `by_array, order` pairs | Add Level |
| Result | A new, live, sorted copy | A new, live, sorted copy | Rearranges the original rows once |
| Best for | Sorting a block by one of its own columns | Sorting by something you don't want to show, or by a calculated array | One-off tidying |

Worked examples on tblDepartments (31 departments, columns DeptID, Department, Facility, ServiceLine, UnitType,
StaffedBeds):

```
=SORT(tblDepartments[[Department]:[StaffedBeds]], 5, -1)
```

The array has five columns (Department through StaffedBeds), so `5` means StaffedBeds and `-1` means largest first.
The first rows are Medical-Surgical 4 West (36 beds), Medical-Surgical 5 East (32), and Cedar Ridge's Medical-Surgical
unit (30). The 13 departments with 0 beds, such as the emergency departments and clinics, come last.

```
=SORT(tblDepartments[[Department]:[StaffedBeds]], {2,5}, {1,-1})
```

Sort by column 2 (Facility) A to Z, then by column 5 (StaffedBeds) largest first within each facility. Ashby Falls
comes first, starting with Medical-Surgical (28 beds), Labor & Delivery (10), and the Intensive Care Unit (8).

```
=SORTBY(tblDepartments[Department], tblDepartments[StaffedBeds], -1)
```

This returns only the department **names**, ordered by bed count. The beds column decides the order but isn't part of
the result. SORT can't do this because its key must be a column of the array it returns.

```
=SORT(tblEncounters[TotalCharges],,-1)
```

To sort a single column, skip `sort_index` with an empty argument (section 5). Excel then sorts by column 1, the only
column there is, and `-1` puts the largest charge (250,917.29) at the top of the 2,000 values.

> ⚠️ `sort_index` counts columns of the array **you pass in**, not worksheet columns. In
> `SORT(tblDepartments[[Department]:[StaffedBeds]], 5, -1)` the 5 is StaffedBeds even though StaffedBeds is column F
> on the sheet.

> ⚠️ **Watch for empty cells in the sort column.** They come back as **0** in the result, but SORT and SORTBY don't
> rank them as zeros, so in a largest-first sort they can end up above the real numbers. That's why this lesson's
> tblDepartments stores a real 0 for departments without beds. With your own data, remove the empty rows with FILTER
> (section 7) before you sort: `SORT(FILTER(array, sort_column<>""), …)`.

> 💡 **Tip:** When two rows tie on the sort key, add a second key so the order is predictable, for example
> `SORT(…, {3,4}, {-1,-1})` sorts by column 3 and breaks ties with column 4.

### 7. FILTER: keep the rows that match

```
=FILTER(array, include, [if_empty])
```

| Argument | Meaning |
|---|---|
| `array` | What to return: one column, several columns, or a whole Table |
| `include` | A TRUE/FALSE test with **one result per row** of `array` |
| `if_empty` | What to show when nothing matches. Without it, FILTER returns `#CALC!` |

A single condition:

```
=FILTER(tblDepartments[[Department]:[Facility]], tblDepartments[UnitType]="Critical Care")
```

This spills a 3 × 2 block: the Intensive Care Unit at Bluestone Memorial, Ashby Falls, and Cedar Ridge.
`tblDepartments[UnitType]="Critical Care"` compares all 31 rows at once and produces 31 TRUE/FALSE values. FILTER keeps
the rows that are TRUE.

**Several conditions** use arithmetic on those TRUE/FALSE arrays. TRUE behaves like 1 and FALSE like 0
([Lesson 2.1](../../02-formulas-functions/01-logical-functions/README.md)):

| Logic | Write it as | Why it works |
|---|---|---|
| A **and** B | `(A)*(B)` | 1 × 1 = 1. Any FALSE makes the product 0 |
| A **or** B | `(A)+(B)` | The sum is at least 1 when either is TRUE, and FILTER treats any non-zero number as TRUE |
| A and (B or C) | `(A)*((B)+(C))` | Parentheses make the OR happen first |
| Not A | `(col<>"value")` or `NOT(A)` | Flip the test |
| Date range | `(tblEncounters[AdmitDate]>=DATE(2025,7,1))*(tblEncounters[AdmitDate]<DATE(2025,10,1))` | Q3 2025 |
| Text contains | `ISNUMBER(SEARCH("Medicare", tblEncounters[Payer]))` | SEARCH finds text anywhere in the cell (Lesson 2.2) |

Worked example: Ashby Falls' 30-day readmissions.

```
=FILTER(tblEncounters[EncounterID],
        (tblEncounters[Facility]="Ashby Falls Community Hospital")*(tblEncounters[Readmit30]="Y"))
```

It spills 12 encounter IDs, starting with ENC111983. Swap the first argument for `tblEncounters` and you get all 16
columns of those 12 rows. A bare Table name such as `tblEncounters` means every data row and every column of the
Table, without the header row.

**Handling "no matches."** Ashby Falls has no outpatient clinics, so this returns `#CALC!`:

```
=FILTER(tblEncounters[EncounterID],
        (tblEncounters[Facility]="Ashby Falls Community Hospital")*(tblEncounters[EncounterType]="Outpatient"))
```

Add the third argument, `if_empty`, to show a message instead. The formula then returns the text "No matching
encounters" in one cell:

```
=FILTER(tblEncounters[EncounterID],
        (tblEncounters[Facility]="Ashby Falls Community Hospital")*(tblEncounters[EncounterType]="Outpatient"),
        "No matching encounters")
```

> ⚠️ **AND() and OR() don't work inside FILTER.** `AND(range1="x", range2="y")` collapses all the rows into a single
> TRUE or FALSE, and FILTER returns `#VALUE!`. Use `*` and `+`.

> ⚠️ Wrap each comparison in its own parentheses. `(A)*(B)+(C)` means "(A and B) or C," which is rarely what you want.

| Tool | Result | Updates by itself? | Use it when |
|---|---|---|---|
| AutoFilter (Lesson 1.6) | Hides rows in place | Reapply after data changes | Exploring by eye |
| Advanced Filter (Lesson 1.6) | Copies matches once | No | One-off extracts in older Excel |
| COUNTIFS / SUMIFS (Lesson 2.5) | One number | Yes | You only need a count or a total |
| **FILTER** | The matching rows themselves | **Yes** | Lists, worklists, and anything you want to nest |

### 8. Turning a spill into one number

Nesting is the core dynamic-array habit: one function's result becomes the next function's input. Read nested
formulas from the inside out. When you need a single answer, finish with a function that collapses the array:

| You want | Wrap the array in | Example |
|---|---|---|
| How many items | `ROWS(…)` or `COUNTA(…)` | `=ROWS(UNIQUE(FILTER(tblEncounters[DxCode], tblEncounters[EncounterType]="Observation")))` returns 27 distinct diagnosis codes among observation stays |
| A total or average | `SUM(…)`, `AVERAGE(…)`, `MAX(…)` | `=AVERAGE(FILTER(tblEncounters[TotalCharges], tblEncounters[EncounterType]="Observation"))` |
| The *n*th item | `INDEX(…, n)` | `=INDEX(SORT(UNIQUE(tblEncounters[Payer])), 1)` is the first payer A to Z |
| The first item | `TAKE(…, 1)` | `=TAKE(SORTBY(…), 1)` |

**Math works item by item.** When two arrays have the same number of rows, `array1*array2` multiplies the first item
by the first item, the second by the second, and so on. Wrap the result in SUM to add up the products. A TRUE/FALSE
test counts as 1 or 0 here too, so this formula adds up only the beds of the three critical-care units (40):

```
=SUM(tblDepartments[StaffedBeds]*(tblDepartments[UnitType]="Critical Care"))
```

Lesson 4.2 builds on this kind of array math with SUMPRODUCT.

> ⚠️ `ROWS(FILTER(…))` returns `#CALC!` when nothing matches, because FILTER has nothing to count. Use
> `IFERROR(ROWS(FILTER(…)), 0)` or COUNTIFS if zero matches is possible.

### 9. SEQUENCE: numbers and dates on demand

```
=SEQUENCE(rows, [columns], [start], [step])
```

| Formula | Result |
|---|---|
| `=SEQUENCE(5)` | 1, 2, 3, 4, 5 down a column |
| `=SEQUENCE(1, 7)` | 1 to 7 across a row |
| `=SEQUENCE(3, 4, 10, 10)` | A 3 × 4 grid: 10, 20, 30, 40 / 50, 60 … 120 |
| `=DATE(2025,1,6) + SEQUENCE(4, 1, 0, 7)` | Four Mondays: 01/06, 01/13, 01/20, 01/27/2025 (a weekly staffing calendar) |
| `=DATE(2025, SEQUENCE(4, 1, 1, 3), 1)` | Quarter starts: 01/01, 04/01, 07/01, 10/01/2025 |
| `=SEQUENCE(ROWS(B6#))` | Row numbers 1, 2, 3 … that match the length of the list in B6 |

DATE accepts an array for any of its arguments, so `DATE(2025, SEQUENCE(4,1,1,3), 1)` builds four dates in one go. DATE
also rolls over: month 13 of 2025 is January 2026. That makes "first day of the next month" easy to write as
`DATE(YEAR(d), MONTH(d)+1, 1)`.

> ⚠️ Dates are serial numbers (Lesson 2.3). If a spilled list shows 45658 instead of 01/01/2025, format the column as a
> date. The Workspace sheet's month column is already formatted for you.

> ⚠️ Some older date functions, including EDATE and EOMONTH, return `#VALUE!` when you hand them a range or a spill
> reference such as `F6#` directly. Build the dates with DATE as shown above, or add `+0` to turn the range into an
> array: `EOMONTH(F6#+0, 0)`.

> 📋 RANDARRAY is SEQUENCE's random cousin: `=RANDARRAY(5, 1, 1, 100, TRUE)` returns five whole numbers from 1 to 100.
> It recalculates every time anything changes, so never use it for graded or reported numbers.

### 10. Reshaping arrays: TAKE, DROP, CHOOSECOLS, VSTACK, HSTACK, and friends

These functions cut, pick, and glue arrays. A negative number counts from the end.

| Function | Syntax | What it does | Example |
|---|---|---|---|
| **TAKE** | `TAKE(array, rows, [columns])` | Keeps the first (or, with a negative number, last) rows or columns | `TAKE(SORT(…,,-1), 10)` gives a top 10 |
| **DROP** | `DROP(array, rows, [columns])` | Removes the first (or last) rows or columns | `DROP(I6#, 1)` removes the header row of a report that spills from I6 |
| **CHOOSECOLS** | `CHOOSECOLS(array, col1, [col2], …)` | Keeps the listed columns, in the order you list them | `CHOOSECOLS(tblDepartments, 2, 6)` gives Department and StaffedBeds |
| **CHOOSEROWS** | `CHOOSEROWS(array, row1, [row2], …)` | Keeps the listed rows | `CHOOSEROWS(B6#, 1, -1)` gives the first and last items |
| **VSTACK** | `VSTACK(array1, [array2], …)` | Stacks arrays on top of each other | `VSTACK({"Unit","Beds"}, …)` puts a header row on top |
| **HSTACK** | `HSTACK(array1, [array2], …)` | Places arrays side by side | `HSTACK(names, counts)` builds a two-column table |
| **TOCOL** | `TOCOL(array, [ignore], [scan_by_column])` | Flattens a block into one column. `ignore` = 1 skips blanks | `UNIQUE(TOCOL(B6:D20, 1))` lists the distinct values in a block |

Worked example: the three largest units, showing only name and beds, under a header row.

```
=VSTACK({"Unit","Beds"},
        TAKE(CHOOSECOLS(SORT(tblDepartments, 6, -1), 2, 6), 3))
```

From the inside out, SORT orders all 31 departments by column 6 (StaffedBeds), largest first. CHOOSECOLS keeps columns
2 and 6 (Department and StaffedBeds). TAKE keeps the first three rows. VSTACK puts the header row on top. The result is
a 4 × 2 block: Unit/Beds, then Medical-Surgical 4 West 36, Medical-Surgical 5 East 32, and Medical-Surgical 30 (Cedar
Ridge's unit).

The `{"Unit","Beds"}` part is an **array constant**: values typed inside braces. A comma moves to the next column and a
semicolon moves to the next row, so `{"Unit";"Beds"}` would stack the two labels vertically instead.

> ⚠️ VSTACK pads with `#N/A` when the stacked arrays have different widths. Make the widths match, or wrap the result
> in `IFNA(…, "")`.

> 📋 In regions where Excel separates function arguments with semicolons, the separators inside array constants differ
> too. If `{"Unit","Beds"}` is rejected, check which separators your Excel uses.

> 💡 **Tip:** CHOOSECOLS(tblEncounters, …) takes column **numbers**. tblEncounters has 16 columns: EncounterID is 1,
> LOSDays is 5, Facility is 7, Department is 8, and TotalCharges is 15. Click a cell in the Table and count from the
> left, or use `XMATCH("LOSDays", tblEncounters[#Headers])` to look a position up.

### 11. XLOOKUP with arrays

XLOOKUP ([Lesson 2.6](../../02-formulas-functions/06-lookup-functions/README.md)) also works with arrays in two
directions.

**Return several columns at once.** Point `return_array` at more than one column and the result spills across:

```
=XLOOKUP("ENC110714", tblEncounters[EncounterID], tblEncounters[[Facility]:[Department]])
```

This returns Bluestone Memorial Hospital and Medical-Surgical 5 East in two adjacent cells.

**Look up many values at once.** Give `lookup_value` a whole array and you get one result per value. This looks up the
specialty of the attending on every Emergency encounter and lists the distinct specialties:

```
=UNIQUE(XLOOKUP(FILTER(tblEncounters[Attending], tblEncounters[EncounterType]="Emergency"),
                tblProviders[Provider], tblProviders[Specialty]))
```

It returns a single row, Emergency Medicine, which is a quick way to confirm that only emergency physicians attend ED
visits. In older Excel this needed a helper column of lookups.

> ⚠️ If any lookup fails, that position shows `#N/A` and functions such as SUM fail with it. Use XLOOKUP's
> `if_not_found` argument: `XLOOKUP(…, …, …, "Unknown")`.

### 12. Building one-formula reports

Functions such as COUNTIFS, SUMIFS, and AVERAGEIFS normally take one criterion and return one number. Give them an
**array** of criteria instead, and Excel runs the function once per item and returns an array of results. This is
called **lifting**, and it is what turns these functions into one-formula summary tables:

```
=COUNTIFS(tblEncounters[PayerType], UNIQUE(tblEncounters[PayerType]))
```

That's five counts, one per payer type, in the same order as the UNIQUE list. Put the two arrays side by side with
HSTACK and sort by the count:

```
=SORT(HSTACK(UNIQUE(tblEncounters[PayerType]),
             COUNTIFS(tblEncounters[PayerType], UNIQUE(tblEncounters[PayerType]))),
      2, -1)
```

The result is a 5 × 2 block with no header row yet:

| Payer type (column 1) | Count (column 2) |
|---|---:|
| Government | 842 |
| Commercial | 729 |
| Medicare Advantage | 309 |
| Self-Pay | 111 |
| Workers' Comp | 9 |

Add a header with `VSTACK({"Payer type","Encounters"}, …)`, or a share column by dividing the counts by
`ROWS(tblEncounters[PayerType])`. The same pattern sorts a list by a total it never shows:
`SORTBY(UNIQUE(names), SUMIFS(amounts, names_column, UNIQUE(names)), -1)`.

**Criteria you build with `&` lift too.** In
[Lesson 2.5](../../02-formulas-functions/05-conditional-aggregation/README.md) you wrote criteria such as
`">="&DATE(2025,7,1)`. Join an operator to an array and you get one criterion per item, so this formula returns three
counts across a row: 416 encounters with a LOSDays of at least 3, 107 with at least 7, and 12 with at least 14.

```
=COUNTIFS(tblEncounters[LOSDays], ">="&{3,7,14})
```

The list after `&` can be an array constant, a spill reference, or another function's result.

**Two criteria lists at once.** When you lift two criteria at the same time, COUNTIFS pairs the lists up item by item:
the first item of one list goes with the first item of the other, the second with the second, and so on. The lists
must have the same number of rows. Try it in the Workspace's Guide examples area. In U6, spill every encounter type
and facility pair that occurs:

```
=UNIQUE(tblEncounters[[EncounterType]:[Facility]])
```

Then, in W6, split that two-column spill with CHOOSECOLS and count each pair:

```
=COUNTIFS(tblEncounters[EncounterType], CHOOSECOLS(U6#, 1),
          tblEncounters[Facility],      CHOOSECOLS(U6#, 2))
```

W6 spills 10 counts, one per pair, that add up to 2,000. The first three are 386 (Inpatient at Bluestone Memorial
Hospital), 498 (Emergency at Bluestone Memorial), and 679 (Outpatient at the Bluestone Outpatient Pavilion). SUMIFS
and AVERAGEIFS pair their criteria lists the same way.

> ⚠️ **Only the criteria can be arrays.** The range arguments of COUNTIFS, SUMIFS, and AVERAGEIFS (the columns they
> test and add up) must be real cell ranges: a Table column, an ordinary range, or a spill reference such as `U6#`.
> If you pass a calculated array there, as in `COUNTIFS(FILTER(…), "Y")`, Excel refuses the formula with a
> "There's a problem with this formula" message. Filter first and count with ROWS, or add another criteria pair.

**Preview: naming the pieces with LET.** The formula above calls `UNIQUE(tblEncounters[PayerType])` twice. LET lets
you calculate it once, give it a name, and reuse the name:

```
=LET(types, UNIQUE(tblEncounters[PayerType]),
     n,     COUNTIFS(tblEncounters[PayerType], types),
     VSTACK({"Payer type","Encounters"}, SORT(HSTACK(types, n), 2, -1)))
```

LET takes pairs of *name, value*, and its last argument is the result.
[Lesson 4.2](../02-advanced-formulas-let-lambda/README.md) covers LET in depth. For now, use it whenever the same piece
appears more than once in a report formula. It makes the formula shorter, easier to read, and faster.

> 💡 **Build from the inside out.** Write the innermost piece in a spare cell and check that it spills what you expect.
> Then wrap it in the next function and check again. Long formulas are easier to read with line breaks: press
> **Alt + Enter** (Mac: **⌃ + ⌥ + Return**) inside the formula bar, and **Ctrl + Shift + U** (Mac: **⌃ + Shift + U**) to
> expand the formula bar.

> 💡 **Formatting doesn't travel with a spill.** Spilled values take on whatever formatting the destination cells
> already have. Format a generous block in advance, as the Workspace sheet does, or use conditional formatting with a
> rule such as `=$I6<>""`.

> 📋 Current Microsoft 365 builds also have **GROUPBY** and **PIVOTBY**, which build summary tables like this in one
> function. They aren't in Excel 2024, so this course builds reports from the functions above.

### 13. Shortcuts and version notes

| Action | Windows | Mac |
|---|---|---|
| Create a Table from a range | Ctrl + T | ⌃ + T (or ⌘ + T) |
| Edit the active cell | F2 | ⌃ + U |
| Accept a function or Table name from AutoComplete | Tab | Tab |
| Expand or collapse the formula bar | Ctrl + Shift + U | ⌃ + Shift + U |
| Line break inside a formula | Alt + Enter | ⌃ + ⌥ + Return |
| Show formulas instead of results | Ctrl + `` ` `` (grave accent) | ⌃ + `` ` `` |
| Undo (for example, a value that blocks a spill) | Ctrl + Z | ⌘ + Z |
| Enter a legacy array formula (Excel 2019 and earlier) | Ctrl + Shift + Enter | ⌘ + Shift + Return |

| Functions | Available in |
|---|---|
| FILTER, SORT, SORTBY, UNIQUE, SEQUENCE, RANDARRAY, XLOOKUP, XMATCH, LET | Microsoft 365, Excel 2021, Excel 2024, Excel for the web, and the matching Mac versions |
| VSTACK, HSTACK, TAKE, DROP, CHOOSECOLS, CHOOSEROWS, TOCOL, TOROW | Microsoft 365, Excel 2024, Excel for the web, and the matching Mac versions |
| GROUPBY, PIVOTBY | Current Microsoft 365 only |

If someone opens your workbook in Excel 2019 or earlier, they see the formulas with an `_xlfn.` prefix (for example
`=_xlfn._xlws.FILTER(…)`), and the cells return `#NAME?` as soon as they recalculate. Share a copy with the results
pasted as values (**Home → Paste → Values**) when you know your audience runs an older version.

## 🧪 Hands-on practice

Download [`4.1-dynamic-arrays.xlsx`](4.1-dynamic-arrays.xlsx) and open the **Practice** sheet. Type a formula in each
yellow cell (or in the yellow anchor cells on the **Workspace** sheet), and the **Check** column turns green when you're
right.

<!-- BEGIN GENERATED: practice -->
Tasks 1, 2, 12, and 13 are spill exercises: build them in the yellow anchor cells on the Workspace sheet, and the gray cells here read your results. Every other answer cell needs a formula that returns ONE value, so wrap spilling functions in ROWS, SUM, AVERAGE, INDEX, or TAKE(…,1). Refer to the data with structured references such as tblEncounters[TotalCharges].

| # | Task | Hint |
|:-:|------|------|
| 1 | Go to the Workspace sheet. In the yellow cell B6, enter one formula that spills the list of distinct payer names from the Payer column of tblEncounters. The gray cell counts the names in your list. | UNIQUE of one Table column |
| 2 | In Workspace!D6, enter one formula that spills the distinct ServiceLine values from tblEncounters sorted A to Z. When you press Enter, D6 shows #SPILL!. Find out why and fix the problem without moving your formula. The gray cell then summarizes your list. | SORT(UNIQUE(…)). Then click the warning icon next to the error |
| 3 | Department names repeat across hospitals (each hospital has its own 'Emergency Department'). How many distinct Facility + Department combinations (units) appear in tblEncounters? | UNIQUE over two adjacent columns, then count the rows |
| 4 | How many patients (PatientID) have exactly one encounter in this extract? | UNIQUE has an optional third argument |
| 5 | This one uses tblDepartments (on the Departments sheet), but the formula still goes in the yellow cell here. Using FILTER, how many departments have UnitType "Inpatient" and 20 or more StaffedBeds? | Multiply the two conditions with *, then count the rows FILTER returns |
| 6 | How many different attending providers (Attending) treated Emergency encounters? | FILTER first, then UNIQUE, then count |
| 7 | Flu-season review: what were the total charges of Emergency encounters admitted in January or February 2025 whose DxCategory is Respiratory or Infectious? Enter dollars and cents. | AND with *, OR with +. Wrap the OR part in its own parentheses |
| 8 | What is the combined TotalCharges of the five most expensive Emergency encounters? Enter dollars and cents. | SORT the ED charges largest first, TAKE the first 5, then add them |
| 9 | Outliers distort averages. What is the average TotalCharges of Inpatient encounters after dropping the 10 most expensive stays? Enter dollars and cents. | Sort largest first, then DROP the first 10 rows |
| 10 | Which attending provider's encounters add up to the highest TotalCharges? Enter the name exactly as it appears (Last, First). | SORTBY the UNIQUE names by a SUMIFS that uses those same names as its criteria |
| 11 | Look up the specialty of the attending on every Inpatient encounter (tblProviders has a Specialty for each Provider name). How many distinct specialties appear? | XLOOKUP accepts a whole array of lookup values |
| 12 | On the Workspace sheet, make F6 spill the first day of each month of 2025 (01/01/2025 through 12/01/2025) using SEQUENCE. Column F is formatted to show them as Jan 2025, Feb 2025, and so on. Then, in G6, write ONE COUNTIFS formula that refers to your month list with the spill operator (F6#) and spills the number of encounters admitted in each month. The gray cell checks both columns. | DATE accepts an array of months. Count AdmitDate ≥ each month start and < the next month's start |
| 13 | Build a long-stay worklist in Workspace!I6 with ONE formula: a header row (EncounterID, Department, LOSDays, TotalCharges) on top of every Cedar Ridge Medical Center encounter with LOSDays of 7 or more, showing only those four columns, sorted by LOSDays from longest to shortest. | VSTACK(header, SORT(CHOOSECOLS(FILTER(tblEncounters, …), …), …)) |
<!-- END GENERATED: practice -->

## ✅ Answer key

The workbook has a hidden **Answer Key** sheet (right-click any sheet tab → **Unhide…** → *Answer Key*). Its *Live
result* column runs a one-cell version of each sample formula, so you can see it working. The same answers are below,
collapsed so you don't see them by accident.

<!-- BEGIN GENERATED: answers -->
<details>
<summary><b>🔑 Show the answer key</b> — Try every task before opening this.</summary>

**1. Distinct payers (Workspace!B6)**

- **Answer:** 8
- **Solution:** `=UNIQUE(tblEncounters[Payer])`

UNIQUE returns each payer once, in the order it first appears in the table. Only B6 holds the formula, and the other names spill into B7, B8, and so on. Click one of them and the formula bar shows the formula grayed out, and a blue border outlines the whole spill range.

**2. Sorted service lines and a #SPILL! fix (Workspace!D6)**

- **Answer:** 10 service lines · first: Ambulatory · last: Women & Children
- **Solution:** `=SORT(UNIQUE(tblEncounters[ServiceLine]))`

A leftover note ('draft') sits in D11, inside the range the list needs. Excel never overwrites data, so it shows #SPILL! in the anchor instead. Click the warning icon, choose **Select Obstructing Cells**, press Delete, and the list spills at once. SORT puts the UNIQUE results in A to Z order. UNIQUE(SORT(…)) gives the same list.

**3. Distinct units (Facility + Department)**

- **Answer:** 27
- **Solution:** `=ROWS(UNIQUE(tblEncounters[[Facility]:[Department]]))`

Given two columns, UNIQUE compares whole rows, so 'Bluestone Memorial Hospital | Emergency Department' and 'Cedar Ridge Medical Center | Emergency Department' are different units. =ROWS(UNIQUE(tblEncounters[Department])) returns only 20, because it merges the same name across hospitals. ROWS counts the rows of the spilled result and returns one number, so it fits in an answer cell. For columns that aren't next to each other, use UNIQUE(HSTACK(col1, col2)).

**4. Patients with exactly one encounter**

- **Answer:** 1,186
- **Solution:** `=ROWS(UNIQUE(tblEncounters[PatientID],,TRUE))`

The third argument, exactly_once, set to TRUE keeps only values that occur once. The two commas skip the second argument (by_col). Compare =ROWS(UNIQUE(tblEncounters[PatientID])), which counts every distinct patient (1,549). The difference, 363 patients, came back at least twice.

**5. Inpatient units with 20+ staffed beds**

- **Answer:** 7
- **Solution:**

```
=ROWS(FILTER(tblDepartments[Department],(tblDepartments[UnitType]="Inpatient")*(tblDepartments[StaffedBeds]>=20)))
```


Each condition returns TRUE or FALSE for all 31 rows. Multiplying them turns TRUE/FALSE into 1/0, so only rows where both are 1 pass the filter (AND logic). Bluestone Memorial Hospital's Intensive Care Unit has 20 beds but doesn't count, because its UnitType is 'Critical Care'. COUNTIFS gives the same number here. FILTER is worth learning because the same include argument can also return the rows themselves.

**6. Distinct ED attendings**

- **Answer:** 28
- **Solution:** `=ROWS(UNIQUE(FILTER(tblEncounters[Attending],tblEncounters[EncounterType]="Emergency")))`

Read it from the inside out. FILTER keeps the Attending names of Emergency encounters (with repeats), UNIQUE removes the repeats, and ROWS counts what's left. Nesting functions this way is the core dynamic-array habit.

**7. Flu-season ED charges**

- **Answer:** 184,214.42
- **Solution:**

```
=SUM(FILTER(tblEncounters[TotalCharges],(tblEncounters[EncounterType]="Emergency")*(tblEncounters[AdmitDate]<DATE(2025,3,1))*((tblEncounters[DxCategory]="Respiratory")+(tblEncounters[DxCategory]="Infectious"))))
```


Three conditions are multiplied (AND), and the last one is itself an OR built by adding two tests. The extra parentheses around the OR matter: without them, multiplication happens first and the formula would add every Infectious encounter of the year regardless of type or date. 64 encounters match. Because the data covers 2025 only, AdmitDate < 3/1/2025 is enough for 'January or February'.

**8. Top five ED charges (TAKE)**

- **Answer:** 42,030.97
- **Solution:**

```
=SUM(TAKE(SORT(FILTER(tblEncounters[TotalCharges],tblEncounters[EncounterType]="Emergency"),,-1),5))
```


SORT(…,,-1) sorts descending (the empty second argument keeps the default sort column). TAKE(array, 5) keeps the first five rows, and SUM collapses them to one number. =SUM(LARGE(FILTER(…),{1,2,3,4,5})) works too, but TAKE scales better: change 5 to 50 and nothing else changes.

**9. Inpatient average without the top 10 (DROP)**

- **Answer:** 31,772.46
- **Solution:**

```
=AVERAGE(DROP(SORT(FILTER(tblEncounters[TotalCharges],tblEncounters[EncounterType]="Inpatient"),,-1),10))
```


DROP is the mirror image of TAKE: it removes rows from the start (or, with a negative number, from the end). The full inpatient average is 33,987.43, so ten very expensive stays move the mean by about \$2,215. TRIMMEAN (Lesson 2.4) trims a percentage from both ends, while DROP lets you decide exactly what to remove.

**10. Attending with the highest total charges (SORTBY)**

- **Answer:** Taylor, Jonathan
- **Solution:**

```
=TAKE(SORTBY(UNIQUE(tblEncounters[Attending]),SUMIFS(tblEncounters[TotalCharges],tblEncounters[Attending],UNIQUE(tblEncounters[Attending])),-1),1)
```


SUMIFS normally takes one criterion. Hand it the whole UNIQUE list and it returns one total per name (142 totals). SORTBY then orders the names by those totals, even though the totals never appear in the result, and TAKE(…,1) keeps the top name. The runner-up, Hawkins, Bruce, is only \$3,146 behind.

**11. Specialties attending inpatient stays (XLOOKUP with an array)**

- **Answer:** 10
- **Solution:**

```
=ROWS(UNIQUE(XLOOKUP(FILTER(tblEncounters[Attending],tblEncounters[EncounterType]="Inpatient"),tblProviders[Provider],tblProviders[Specialty])))
```


FILTER returns 529 attending names, one per inpatient stay. XLOOKUP looks up each of them and returns an array of the same size, so you get one specialty per stay. UNIQUE and ROWS then count the distinct specialties. In Excel 2019 and earlier you'd need a helper column for this.

**12. Month calendar with SEQUENCE and F6# (Workspace!F6:G6)**

- **Answer:** 12 months · 2000 encounters
- **Solution:**

1. In **F6**: `=DATE(2025,SEQUENCE(12),1)`
2. In **G6**: `=COUNTIFS(tblEncounters[AdmitDate],">="&F6#,tblEncounters[AdmitDate],"<"&DATE(YEAR(F6#),MONTH(F6#)+1,1))`


SEQUENCE(12) spills 1 to 12, and DATE turns each number into that month's first day. In G6, F6# means 'the whole spill that starts in F6', so COUNTIFS receives 12 start dates and returns 12 counts. The upper bound DATE(YEAR(F6#),MONTH(F6#)+1,1) is the next month's first day (month 13 rolls into January 2026). If you change SEQUENCE(12) in F6 to SEQUENCE(24), G6 grows to 24 counts automatically. December is the busiest month (221 encounters).

**13. One-formula long-stay worklist (Workspace!I6)**

- **Answer:** 9 stays · longest: ENC117529 (19 d) · shortest: ENC118758 (7 d)
- **Solution:**

```
=VSTACK({"EncounterID","Department","LOSDays","TotalCharges"},SORT(CHOOSECOLS(FILTER(tblEncounters,(tblEncounters[Facility]="Cedar Ridge Medical Center")*(tblEncounters[LOSDays]>=7)),1,8,5,15),3,-1))
```


Build it from the inside out. FILTER(tblEncounters, …) returns all 16 columns of the matching rows. CHOOSECOLS keeps columns 1, 8, 5, and 15 in the order you list them. SORT(…, 3, -1) sorts by the third of those columns (LOSDays), largest first. VSTACK puts the header row, an array constant in braces, on top. Two stays tie at 14 days. To break ties by charges, use SORT(…, {3,4}, {-1,-1}).

</details>
<!-- END GENERATED: answers -->

## 🏆 Bonus challenge

<!-- BEGIN GENERATED: bonus -->
The CFO wants a 'unit leaderboard' she can refresh every month: every unit (Facility + Department) with at least 50 encounters in the extract, with four columns (Facility, Department, Encounters, AvgCharge), where Encounters is the unit's number of encounters and AvgCharge is the average TotalCharges of those encounters. Sort it by AvgCharge from highest to lowest and put a header row on top. Build it in Workspace!N6 as ONE formula. Section 12 of the lesson guide shows how to lift COUNTIFS and AVERAGEIFS over a two-column list, and LET (previewed there) keeps the formula readable. Then answer the questions below with formulas that refer to your leaderboard through N6#.

Work on the **Bonus** sheet of the workbook.

- **B1.** How many units qualify for the leaderboard? Count data rows only, not the header. *(Hint: ROWS of the spill, minus the header row)*
- **B2.** Which Department tops the leaderboard? *(Hint: Row 2 of the spill is the first data row)*
- **B3.** What is the top unit's average charge? Enter dollars and cents. *(Hint: Same row, fourth column)*
- **B4.** What share of all TotalCharges in tblEncounters comes from the units on the leaderboard? Each unit's total is Encounters × AvgCharge. Enter it as a percentage. *(Hint: DROP the header row, multiply column 3 by column 4, add up, then divide by all charges)*
<!-- END GENERATED: bonus -->

<!-- BEGIN GENERATED: bonus-answers -->
<details>
<summary><b>🔑 Show the bonus solution</b> — Give it a real try first!</summary>

**B1. Units on the leaderboard**

- **Answer:** 13
- **Solution:**

1. In **Workspace!N6**:

```
=LET(units, UNIQUE(tblEncounters[[Facility]:[Department]]),
     fac,   CHOOSECOLS(units, 1),
     dept,  CHOOSECOLS(units, 2),
     n,     COUNTIFS(tblEncounters[Facility], fac, tblEncounters[Department], dept),
     avg,   AVERAGEIFS(tblEncounters[TotalCharges], tblEncounters[Facility], fac, tblEncounters[Department], dept),
     VSTACK({"Facility","Department","Encounters","AvgCharge"},
            SORT(FILTER(HSTACK(units, n, avg), n >= 50), 4, -1)))
```

2. In the answer cell: `=ROWS(Workspace!N6#)-1`


UNIQUE over Facility:Department lists every unit once. CHOOSECOLS splits that two-column list so COUNTIFS and AVERAGEIFS can use each column as a criteria array, and each returns one value per unit. HSTACK glues the four columns together, FILTER keeps units with n ≥ 50, SORT orders by column 4 descending, and VSTACK adds the header. LET names each piece once, so UNIQUE runs once. Without LET, the same UNIQUE(…) would have to be repeated in every part of the formula that needs it.

**B2. Top unit's department**

- **Answer:** Cardiac Step-Down
- **Solution:** `=INDEX(Workspace!N6#,2,2)`

INDEX works on a spill like on any range: row 2 (row 1 is the header), column 2. The top unit is Cardiac Step-Down at Bluestone Memorial Hospital. Inpatient units dominate the top of the board because a stay costs far more than a clinic visit.

**B3. Top unit's average charge**

- **Answer:** 39,862.22
- **Solution:** `=INDEX(Workspace!N6#,2,4)`

Column 4 of the leaderboard is AvgCharge. A Cardiac Step-Down stay averages \$39,862, about 15 times the average Bluestone Memorial emergency visit. If your answer has many decimals, that's fine: the check accepts the unrounded value.

**B4. Leaderboard units' share of charges**

- **Answer:** 62.4%
- **Solution:**

```
=SUM(CHOOSECOLS(DROP(Workspace!N6#,1),3)*CHOOSECOLS(DROP(Workspace!N6#,1),4))/SUM(tblEncounters[TotalCharges])
```


DROP(N6#,1) removes the header so only numbers remain. CHOOSECOLS pulls out the Encounters and AvgCharge columns. Multiplying them row by row rebuilds each unit's total charges, and SUM adds them. The 13 busiest units bring in 62.4% of all charges. The other 14 units see fewer patients, but most of them are ICUs and inpatient floors with high charges per stay, so together they still account for the remaining 37.6%.

</details>
<!-- END GENERATED: bonus-answers -->

## Key takeaways

- A dynamic array formula lives in one **anchor cell** and **spills** its results into the cells around it. Leave
  room, because anything in the way causes `#SPILL!`.
- A **spill reference** such as `A2#` refers to a whole spill and resizes with it. Use it in formulas, drop-down
  lists, and charts.
- **UNIQUE**, **SORT**/**SORTBY**, and **FILTER** replace copy-paste-dedupe routines with live lists. In FILTER, `*`
  means AND and `+` means OR.
- Wrap a spill in **ROWS**, **SUM**, **AVERAGE**, **INDEX**, or **TAKE(…,1)** when you need one number.
- **SEQUENCE** generates numbers and dates. **TAKE**, **DROP**, **CHOOSECOLS**, **VSTACK**, and **HSTACK** cut and
  assemble arrays.
- Lifting COUNTIFS, SUMIFS, or AVERAGEIFS over a UNIQUE list turns them into one-formula summary tables. LET keeps those
  formulas readable.

<!-- BEGIN GENERATED: nav -->
---

⬅️ **Previous:** [3.6 What-If Analysis: Goal Seek, Scenarios, Data Tables & Solver](../../03-data-analysis/06-what-if-analysis/README.md) · 🏠 [Course home](../../README.md) · **Next:** [4.2 Advanced Formulas: LET, LAMBDA & Array Logic](../02-advanced-formulas-let-lambda/README.md) ➡️
<!-- END GENERATED: nav -->
