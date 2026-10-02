# Lesson 4.3 · Power Query: Import, Transform & Combine

> **Level:** Advanced · **Time:** about 70 minutes · **Workbook:** [`4.3-power-query.xlsx`](4.3-power-query.xlsx) · **Files:** [`data/`](data/) (one download: [`4.3-power-query-data.zip`](data/4.3-power-query-data.zip))
> **Data:** CSV exports in the lesson's data folder: twelve monthly files of claims submitted in 2025, a January 2026 claim file (simulated next-month export) for the bonus, the payer list, 11,196 encounters discharged in 2025, and a wide 2025 budget-vs-actual export. The workbook adds a hand-maintained denial-reason mapping table.

Early every month, a revenue-cycle analyst at Bluestone Health downloads the billing system's claims export. She pastes it
under last month's rows, fixes the date columns, looks up each payer's type, and rebuilds the denial report. It takes most
of a morning, and one bad paste quietly breaks the numbers. **Power Query** records each of those steps once and replays
them on new files when you click **Refresh**. In this lesson you turn twelve monthly claim exports, a payer list, an
encounter extract, and a wide budget file into a small pipeline. One query reads a whole folder of files. Other queries
filter, group, merge, and unpivot the results, and a denial dashboard updates itself when the January 2026 file arrives.

## What you'll learn

- Import CSV files and whole folders with Get & Transform
- Clean and reshape data with Applied Steps: types, splits, filters, unpivot, group by
- Merge (join) and append queries
- Build refreshable, repeatable data pipelines and read the M code behind them

## 📖 Guide

The examples use the CSV files in this lesson's `data` folder. Section 2 shows how to get them. Paths are written for
Windows as `C:\PQ\data\`. On a Mac, use a path such as `/Users/you/PQ/data/` instead.

> 📋 **Version check:** Power Query is built into Excel 2016 and later on Windows (Microsoft 365, 2016, 2019, 2021, 2024)
> and into Microsoft 365 for Mac. It needs the **desktop** app, because Excel for the web can't read files from your
> computer. Section 17 lists what differs between versions.

### 1. What Power Query does

**Power Query** is Excel's built-in tool for **ETL**: it *extracts* data from a source (a CSV file, a folder of files,
a database, or a table in the workbook), *transforms* it (filters, cleans, reshapes, combines), and *loads* the result
into a worksheet table, a PivotTable, or the Data Model. Its commands live in the **Get & Transform Data** group on the
**Data** tab, which is why Microsoft also calls the feature **Get & Transform**.

Every change you make in the Power Query Editor is saved as an **applied step**. The ordered list of steps, together with
its source, is a **query**. Three properties make a query different from a formula or a manual cleanup:

- **Non-destructive.** The source file is never changed. The query reads the CSV and writes its result somewhere else.
- **Repeatable.** The steps are stored, so the same cleanup runs the same way every time.
- **Refreshable.** **Data → Refresh All** reruns every query on whatever the source files contain right now.

| Approach | Best for | Weak spot |
|---|---|---|
| Formulas (Modules 1–2, Lesson 4.1) | Calculations that update the moment someone types | Slow on very large data, and they can't import, clean, or combine CSV files |
| Manual cleanup (Lesson 3.3) | One-off fixes to a single file | You redo every step next month |
| **Power Query** (this lesson) | Importing, cleaning, reshaping, and combining the same kind of data again and again | Results change only when you refresh |
| VBA macros (Module 5) | Automating Excel itself: formatting, buttons, emailing reports | Code to maintain, and macro security prompts |

> 💡 **Tip:** A good rule of thumb is to let Power Query get the data into shape, then use formulas, PivotTables, and
> charts on the clean result. Lesson 3.3 cleaned a messy file by hand. This lesson is the "do it once, refresh forever"
> version of that work.

### 2. Get the lesson files

Power Query reads files from your disk, so the workbook alone isn't enough. You need the `data` folder too.

1. Download the data. The quickest way is to open [`data/4.3-power-query-data.zip`](data/4.3-power-query-data.zip) on
   GitHub and click the download button (**Download raw file**). You can also download the whole course with
   **Code → Download ZIP** on the repository's home page.
2. Extract the zip. On Windows, right-click it and choose **Extract All…** (Windows puts the `data` folder inside a new
   folder named `4.3-power-query-data`). On a Mac, double-click it.
3. Move the extracted `data` folder somewhere with a short path, such as `C:\PQ\data` (Mac: `/Users/you/PQ/data`).
4. Download [`4.3-power-query.xlsx`](4.3-power-query.xlsx), open it, and build all of your queries in this workbook.
   Queries are saved inside the workbook, so a normal `.xlsx` file keeps them. No macros are involved.

```
C:\PQ\data\
├── claims_monthly\          claims_2025_01.csv … claims_2025_12.csv (one file per submit month)
├── new_month\               claims_2026_01.csv (bonus only: a simulated January 2026 export)
├── budget_2025_wide.csv     2025 budget and actual, one column per month
├── encounters_2025.csv      encounters discharged in 2025
├── payers.csv               payer list with PayerType
└── 4.3-power-query-data.zip (only if you downloaded the repository)
```

Each claims file has the same 13 columns as the course's [`claims.csv`](../../data/README.md#claimscsv): ClaimID,
EncounterID, PatientID, PayerID, ServiceDate, SubmitDate, BilledAmount, AllowedAmount, PatientResponsibility, PaidAmount,
ClaimStatus, DenialReason, and PaidDate. The **Data Files** sheet in the workbook describes every file. The course
data ends on 12/31/2025, so `claims_2026_01.csv` is a **simulated** next-month export: claims for late-2025 services that
go out in January 2026.

> ⚠️ Import from the **extracted** folder, never from inside the zip. Windows lets you browse into a zip as if it were a
> folder, but Power Query can't read files there.

> ⚠️ Keep the `new_month` folder **outside** `claims_monthly` until the bonus tells you to copy its file in. From Folder
> reads subfolders too, so a folder inside `claims_monthly` would be combined with the 2025 files.

### 3. Import a CSV file or a workbook table

This walk-through imports `payers.csv`, which you need for task 7. Task 1 uses the same steps on a claims file.

1. Click **Data → Get Data → From File → From Text/CSV**. (The **From Text/CSV** button in the Get & Transform Data group
   is a shortcut to the same command.)
2. Select `C:\PQ\data\payers.csv` and click **Import**.
3. A preview window opens. Check its three settings:

   | Setting | What it controls | For this lesson's files |
   |---|---|---|
   | **File Origin** | The text encoding | *65001: Unicode (UTF-8)* or *1252: Western European* both work, because the files are plain ASCII |
   | **Delimiter** | The character between fields | *Comma* |
   | **Data Type Detection** | How Power Query guesses column types: from the first 200 rows, from the entire file, or not at all | *Based on first 200 rows* (the default) |

4. Click **Transform Data**. The **Power Query Editor** opens with the data and three applied steps.

On a Mac, click **Data → Get Data (Power Query)**, choose **Text/CSV** in the data source dialog, browse to the file, and
then choose **Transform Data**.

The preview window offers three ways forward:

| Button | What happens |
|---|---|
| **Load** | Loads the data straight into a new worksheet as an Excel Table |
| **Load To…** (Load's dropdown) | Asks where to load it: Table, PivotTable, connection only, or the Data Model |
| **Transform Data** | Opens the Power Query Editor so you can check and change the data first |

> 💡 **Tip:** Prefer **Transform Data**. You can fix types and remove columns before anything lands in your workbook.

From Text/CSV writes three steps for you, and you can see them under **Applied Steps** on the right:

| Step | What it does | M function |
|---|---|---|
| **Source** | Reads the file and splits each line at the commas | `Csv.Document(File.Contents(...))` |
| **Promoted Headers** | Turns the first line into column names | `Table.PromoteHeaders` |
| **Changed Type** | Sets each column's data type, guessed from the first 200 rows | `Table.TransformColumnTypes` |

> 💡 **Tip:** To bring in every column exactly as the file wrote it, as Text, choose **Do not detect data types** in the
> preview window, or delete the Changed Type step in the editor. Then set each column's type yourself (section 5). This is
> the safe way to import MRNs such as `04979796`, which Excel turns into the number 4979796 when you double-click a CSV
> (Lesson 3.3).

Power Query names the query after the file (payers). Before you load, type a better name in the **Name** box under
Query Settings, such as `Payers`. Then click **Home → Close & Load**. Excel adds a new sheet with an Excel Table named after
the query and opens the **Queries & Connections** pane (**Data → Queries & Connections**), which shows "8 rows loaded."
That pane is the most reliable place to read a query's row count.

> ⚠️ The status bar at the bottom of the editor counts only the **preview**, which holds the first 1,000 rows. A large
> query shows "999+ ROWS" there, so read the real total from Queries & Connections after you load.

**Import a table from this workbook.** Power Query can also read an Excel Table in the workbook you're working in, such
as the hand-maintained **tblDenialMap** on the DenialMap sheet (task 9).

1. Click any cell inside the Table.
2. Click **Data → From Table/Range** in the Get & Transform Data group. Some versions label it **From Table**. If the
   cells aren't a Table yet, Excel first shows the **Create Table** dialog and converts them.
3. The Power Query Editor opens. The query is named after the Table (`tblDenialMap`), so rename it, for example to
   `DenialMap`.

The Source step reads the Table by name, and Power Query adds a Changed Type step after it:

```
= Excel.CurrentWorkbook(){[Name="tblDenialMap"]}[Content]
```

Because the source is a Table, rows that someone adds to it later are picked up on the next refresh. **Close & Load**
writes the query's result to a **new** Table on a new sheet, so a lookup query like this is usually better loaded as
**Only Create Connection** (section 15).

### 4. A tour of the Power Query Editor

To reopen a query later, double-click it in the Queries & Connections pane. On Windows, **Alt + F12** opens the editor
directly.

| Area | Where | What it's for |
|---|---|---|
| **Ribbon** | Top: Home, Transform, Add Column, View | Commands. **Transform** changes a column in place. **Add Column** builds a new column and keeps the original |
| **Queries pane** | Left (click **>** to expand it) | Every query in the workbook. Right-click a query to Reference, Duplicate, Rename, or Delete it |
| **Formula bar** | Above the data (turn it on with **View → Formula Bar**) | The M code of the selected step. You can edit it directly |
| **Data preview** | Middle | The first 1,000 rows as they look after the selected step |
| **Query Settings** | Right | The query's **Name** and its list of **Applied Steps** |
| **Status bar** | Bottom | Column count, preview row count, and the column-profiling scope |

Working with Applied Steps:

- **Click any step** to see the data as it was at that point. Click the last step to return to the end result.
- **Click the gear icon** next to a step (when it has one) to reopen its dialog and change the settings.
- **Click the X** to delete a step. Deleting a step in the middle can break later steps that use a column it created.
- **Right-click a step** to rename it (F2 on Windows), move it, insert a step after it, or **Delete Until End**.
- If you select an earlier step and then add a new command, Power Query asks before inserting the step there.

**View → Column quality**, **Column distribution**, and **Column profile** add small bars and statistics above and
below each column. The green/red/gray **Column quality** bar shows the share of valid values, errors, and empty values,
so a red sliver warns you about conversion errors before you load.

> ⚠️ Column profiling is based on the **top 1,000 rows** by default. To profile everything, click the status bar text
> *Column profiling based on top 1000 rows* and choose *Column profiling based on entire data set*.

To select several columns, hold **Ctrl** (Mac: **⌘**) and click each header, or hold **Shift** and click the first and
last headers of a block.

### 5. Data types, errors, nulls, and locale

The icon at the left of each column header shows its **data type**. Click the icon to change it.

| Icon | Data type | M type | Example in this lesson |
|:-:|---|---|---|
| ABC | Text | `type text` | ClaimID, PayerID, DenialReason |
| 123 | Whole Number | `Int64.Type` | AvgDaysToPay, the budget month columns |
| 1.2 | Decimal Number | `type number` | BilledAmount, PaidAmount |
| $ | Fixed Decimal Number (currency, 4 decimal places) | `Currency.Type` | An alternative for money columns |
| % | Percentage | `Percentage.Type` | A denial rate you calculate |
| Calendar | Date | `type date` | ServiceDate, SubmitDate, PaidDate |
| Calendar + clock | Date/Time | `type datetime` | AdmitDateTime |
| Stopwatch | Duration | `type duration` | The result of subtracting two dates |
| ABC/123 | Any (no type yet) | `type any` | A new custom column: set a type before you load it |

When you change a type right after the Changed Type step, Power Query asks whether to **Replace current** (edit the
existing Changed Type step) or **Add new step**. Replace current keeps the step list short.

To strip the clock time from a Date/Time column such as AdmitDateTime, change its type to **Date**, or select it and click
**Transform → Date → Date Only**. Either way 07/15/2025 14:32 becomes 07/15/2025, which is what a date key needs when you
relate tables in the Data Model (Lesson 4.4).

> ⚠️ Keep ID columns such as ClaimID, MRN, ZIP, and PayerID as **Text**, even when they look numeric. As numbers, IDs
> lose their leading zeros (Lesson 3.3) and won't match the text keys in other tables when you merge.

**Errors.** If a value can't be converted, for example the text "N/A" in a number column, the cell shows *Error* and the
column quality bar turns red. Click the empty space in an error cell to read its message. You can then fix the source,
or use **Home → Remove Rows → Remove Errors**, or right-click the column and choose **Replace Errors**.

**Nulls and empty text.** A blank field in a CSV arrives in Power Query in one of two ways. In number and date columns it
becomes **null** (Power Query's "no value"), so PaidDate is null on every unpaid claim. In text columns it usually stays an
**empty string** `""`, so DenialReason looks blank but isn't null. The **Remove Empty** command in a column's filter menu
removes both, and in M you would write `[DenialReason] <> null and [DenialReason] <> ""`.

**Locale.** The Changed Type step converts text using your computer's regional settings. This lesson's files use ISO dates
(`2025-01-07`) and periods as decimal separators, and ISO dates read correctly everywhere. If your Windows region uses a
comma as the decimal separator, money columns can come in far too large or as errors. Fix it with **right-click the column
→ Change Type → Using Locale…**, then pick the data type and the locale **English (United States)**. The same command
reads US-style dates such as `01/07/2025` correctly on a computer set to a UK or European region.

```
= Table.TransformColumnTypes(#"Promoted Headers", {{"PaidAmount", type number}}, "en-US")
```

### 6. Everyday transforms

Most cleanup takes only a handful of commands. Each one adds a step, and each step is an M function.

| Goal | Command | M function |
|---|---|---|
| Keep only some columns | **Home → Choose Columns**, or select them and right-click → **Remove Other Columns** | `Table.SelectColumns` |
| Drop columns | Select them, then press **Delete** or **Home → Remove Columns** | `Table.RemoveColumns` |
| Rename a column | Double-click the header | `Table.RenameColumns` |
| Filter rows | The header's dropdown arrow | `Table.SelectRows` |
| Sort | The header's dropdown arrow | `Table.Sort` |
| Remove duplicate rows | Select the key columns → **Home → Remove Rows → Remove Duplicates** | `Table.Distinct` |
| Replace a value | **Transform → Replace Values** | `Table.ReplaceValue` |
| Trim spaces, clean, change case | **Transform → Format → Trim / Clean / lowercase / UPPERCASE / Capitalize Each Word** | `Text.Trim`, `Text.Clean`, `Text.Lower`, `Text.Upper`, `Text.Proper` |
| Split one column into several | **Transform → Split Column → By Delimiter** | `Table.SplitColumn` |
| Pull out part of the text | **Transform → Extract → First Characters / Text Before Delimiter / …** | `Text.Start`, `Text.BeforeDelimiter`, … |
| Join columns into one | Select them → **Transform → Merge Columns** | `Table.CombineColumns` |
| Fill blanks with the value above | **Transform → Fill → Down** | `Table.FillDown` |
| Use the first row as headers | **Home → Use First Row as Headers** | `Table.PromoteHeaders` |

> 📋 Format, Extract, and Merge Columns appear on both the **Transform** and **Add Column** tabs. On Transform they
> replace the column. On Add Column they create a new column and keep the original. For example, **Add Column → Extract →
> First Characters** (3) on PrimaryDxCode gives the ICD-10 category (`I50.9` becomes `I50`) next to the full code, which is
> the Power Query version of `=LEFT(code,3)`.

**Worked example: split a "code - name" column.** In `budget_2025_wide.csv`, the Department column holds values such as
`6110 - Medical-Surgical 4 West`: a cost center, a space-hyphen-space separator, and the department name.

1. Select Department, then click **Transform → Split Column → By Delimiter**.
2. Choose **--Custom--** and type a space, a hyphen, and a space: ` - `.
3. Under **Split at**, choose **Left-most delimiter**, then click OK.
4. Rename the new columns Department.1 and Department.2 to CostCenter and DeptName.

| Split at | Result for `6110 - Medical-Surgical 4 West` |
|---|---|
| Left-most delimiter ` - ` | `6110` and `Medical-Surgical 4 West` ✔ |
| Each occurrence of `-` (no spaces) | `6110 `, ` Medical`, and `Surgical 4 West`, which breaks the name ✘ |

Power Query adds a Changed Type step after the split, and it usually makes CostCenter a Whole Number. That's fine here.
Make it Text instead if you'll merge on it with a text key.

**Filtering.** The header dropdown offers a checklist of values plus **Text Filters**, **Number Filters**, or
**Date Filters**. The checklist shows only the first 1,000 distinct values, so click **Load more** when the list says it
may be incomplete.

> ⚠️ When you tick a few values in the checklist, Power Query writes a rule that keeps exactly those values, such as
> `[ClaimStatus] = "Denied" or [ClaimStatus] = "Appealed"`. When you untick a few, it writes the opposite, such as
> `[ClaimStatus] <> "Paid"`. Next month a brand-new status would be kept by the second rule but dropped by the first.
> Click the step and read the formula bar to be sure the rule says what you mean.

### 7. Custom and conditional columns

**Add Column → Conditional Column** builds an `if … then … else` column from a dialog, with no typing. To flag denied
claims, set *New column name* to `IsDenied`, *If* `ClaimStatus` *equals* `Denied` *then* `1`, *Else* `0`. Power Query
writes:

```
= Table.AddColumn(#"Expanded Payers", "IsDenied", each if [ClaimStatus] = "Denied" then 1 else 0)
```

Click **Add Clause** to add *else if* conditions, for example to sort claims into billed-amount bands.

**Add Column → Custom Column** lets you type an expression in M, Power Query's formula language. You refer to a column of
the current row as `[ColumnName]`, and you can double-click a name in the *Available columns* list to insert it.

| You want | Custom column formula |
|---|---|
| Days between two dates | `Duration.Days([SubmitDate] - [ServiceDate])` |
| Year or month of a date | `Date.Year([ServiceDate])`, `Date.Month([ServiceDate])` |
| First day of the month | `Date.StartOfMonth([ServiceDate])` |
| Join text | `[FacilityID] & " / " & [DeptID]` |
| Number as text | `Text.From([AvgDaysToPay])` |
| First three characters | `Text.Start([PrimaryDxCode], 3)` |
| A ratio | `[PaidAmount] / [BilledAmount]` |
| A label from a test | `if [BilledAmount] >= 50000 then "High dollar" else "Standard"` |
| Test for a missing value | `if [PaidDate] = null then "Unpaid" else "Paid"` |

Subtracting one date from another gives a **duration** (for example 4.00:00:00, meaning four days), and `Duration.Days`
turns it into the whole number 4. You can also build that column without typing: select SubmitDate, Ctrl-click
(Mac: ⌘-click) ServiceDate, then click **Add Column → Date → Subtract Days**. The column you select first is the one
Power Query subtracts from.

> ⚠️ **M is case-sensitive.** `"Denied"` doesn't equal `"denied"`, and the function is `Text.Start`, not `text.start`.
> Keywords such as `if`, `then`, `else`, `and`, `or`, and `not` are lowercase. M also never converts numbers to text for you,
> so `"Days: " & 5` is an error. Write `"Days: " & Text.From(5)` instead.

> 💡 **Tip:** A custom column starts with the type Any (ABC/123 icon) unless you pick a type in the dialog. Set its type
> before loading, or Excel receives untyped values and your PivotTables may treat numbers as text.

**Add Column → Column From Examples** is the Power Query version of Flash Fill. You type the result you want in a few
rows and Power Query writes the M formula for you. Check the formula it proposes, because it guesses from your examples.

### 8. Group By and column statistics

**Group By** collapses many rows into one row per group and calculates summaries, much like a PivotTable. The difference
is that the result is an ordinary table you can keep transforming, merging, and loading.

1. Click **Home → Group By** (it's on the Transform tab too).
2. **Basic** groups by one column with one aggregation. **Advanced** lets you **Add grouping** (more columns to group by)
   and **Add aggregation** (more summaries).
3. For each aggregation, enter a *New column name*, an *Operation*, and the *Column* it uses.

| Operation | Returns |
|---|---|
| Count Rows | Number of rows in the group (no column needed) |
| Count Distinct Rows | Number of rows that differ in at least one column |
| Sum, Average, Median, Min, Max | The usual statistics of one numeric column |
| All Rows | A nested table of the group's rows, for advanced follow-up steps |

**Worked example.** Grouping the January 2025 claims file by ClaimStatus with Count Rows returns five rows: Paid 860,
Partially Paid 97, Denied 77, Appealed 33, and Pending 22. The M code is one step:

```
= Table.Group(#"Changed Type", {"ClaimStatus"}, {{"Claims", each Table.RowCount(_), Int64.Type}})
```

Group by two columns, such as PayerType *and* DenialReason, to get one row per combination. Either select both column
headers before you click Group By, or click **Add grouping** in the Advanced dialog. The bonus uses this to build a denial
dashboard.

> 💡 **Tip:** To get a **rate** per group, aggregate the two counts you need, then divide them in a custom column after the
> Group By step. Sum of a 1/0 flag counts the rows where the flag is 1. The bonus calculates denial rates this way.

**One number from a whole column.** To total a column without grouping, select it and click **Transform → Statistics →
Sum**. The Statistics menu sits in the Number Column group and also offers Average, Minimum, Maximum, Count Values, and
more. The query then returns a single number instead of a table, so add this step to a query that **references** your
staging query (section 14), never to the staging query itself. **Transform → Count Rows** works the same way and returns
the number of rows.

```
= List.Sum(Source[PaidAmount])
```

### 9. Unpivot and pivot

Finance and scheduling exports often arrive **wide**, with one column per month. Analysis tools want them **long**, with
one row per month and a single Month column. **Unpivot** turns wide into long, and **Pivot** turns long back into wide.

Wide (two rows of `budget_2025_wide.csv`, first three months shown):

| Department | Category | Measure | Jan | Feb | Mar |
|---|---|---|--:|--:|--:|
| 6110 - Medical-Surgical 4 West | Salaries & Wages | Budget | 648,732 | 585,952 | 648,732 |
| 6110 - Medical-Surgical 4 West | Salaries & Wages | Actual | 642,497 | 586,090 | 616,790 |

Long, after **Unpivot Other Columns** (Attribute renamed to Month, Value renamed to Amount):

| Department | Category | Measure | Month | Amount |
|---|---|---|---|--:|
| 6110 - Medical-Surgical 4 West | Salaries & Wages | Budget | Jan | 648,732 |
| 6110 - Medical-Surgical 4 West | Salaries & Wages | Budget | Feb | 585,952 |
| 6110 - Medical-Surgical 4 West | Salaries & Wages | Budget | Mar | 648,732 |
| 6110 - Medical-Surgical 4 West | Salaries & Wages | Actual | Jan | 642,497 |
| … | | | | |

Now "Q1" is a filter on Month, and a single Group By or PivotTable can total any month range.

**Transform → Unpivot Columns** has three versions. They differ in what they remember for the next refresh:

| Command | You select | M it writes | A new month column next refresh… |
|---|---|---|---|
| **Unpivot Columns** | The month columns | `Table.UnpivotOtherColumns`, listing the columns you did *not* select | is unpivoted too |
| **Unpivot Other Columns** | The label columns | `Table.UnpivotOtherColumns`, listing the columns you selected | is unpivoted too |
| **Unpivot Only Selected Columns** | The month columns | `Table.Unpivot`, listing the columns you selected | stays as its own column |

> ⚠️ Remove total columns such as **FY Total** before you unpivot. Otherwise the total becomes a thirteenth "month" and
> every annual sum doubles.

> ⚠️ Unpivot skips **null** cells. A blank month produces no row at all, so a row count after unpivoting can be smaller
> than rows × months when the source has gaps.

**Pivot** goes the other way. Select the Measure column, click **Transform → Pivot Column**, choose **Amount** as the
*Values Column*, and keep the aggregate function **Sum** (under *Advanced options*). Budget and Actual become side-by-side
columns, and a custom column `[Actual] - [Budget]` gives the variance. For 4 West's January salaries that's
642,497 − 648,732 = −6,235, which is under budget.

```
= Table.Pivot(#"Renamed Columns", List.Distinct(#"Renamed Columns"[Measure]), "Measure", "Amount", List.Sum)
```

### 10. Append queries

**Appending** stacks the rows of two or more queries into one, like pasting one table under another. **Merging**
(section 13) adds columns from matching rows of another table, like a lookup.

| | Append | Merge |
|---|---|---|
| What grows | Rows | Columns |
| Matching rule | Columns line up **by name** | Rows match on **key columns** |
| Healthcare example | November claims + December claims | Claims + the payer list, to get PayerType |
| M function | `Table.Combine` | `Table.NestedJoin` |

1. In the editor, click **Home → Append Queries → Append Queries as New**. (Plain **Append Queries** adds the rows to the
   current query instead of making a new one.)
2. Choose **Two tables** or **Three or more tables**, pick the queries, and click OK.

```
= Table.Combine({claims_2025_11, claims_2025_12})
```

> ⚠️ Columns are matched by **name**. If one export says `PaidAmt` and another says `PaidAmount`, you get two half-empty
> columns, with null in the rows that came from the other file. Rename columns so they match before you append.

### 11. Combine all the files in a folder

Appending twelve monthly queries by hand works, but next month you'd have to add a thirteenth. **From Folder** treats the
folder itself as the source, so new files are picked up automatically.

1. Click **Data → Get Data → From File → From Folder** (Mac: **Data → Get Data (Power Query)**, then choose **Folder**).
2. Browse to `C:\PQ\data\claims_monthly` and click **Open** (or **OK**).
3. A window lists the files. Click **Combine → Combine & Transform Data**.
4. In the **Combine Files** dialog, leave *Sample File* as **First file**, check that the preview looks right, and click OK.
5. Rename the new query (it's named after the folder) to **Claims**.

Power Query builds the combine from several queries, which it places in a **Helper Queries** group:

| Query | What it is |
|---|---|
| **Sample File** | The first file in the folder, used to design the steps |
| **Parameter1** | A parameter that points the transform at one file at a time |
| **Transform Sample File** | The steps applied to every file. Edit this query to change how each file is cleaned |
| **Transform File** | A function made from Transform Sample File. The main query calls it once per file |
| **Claims** (your query) | Lists the files, calls Transform File on each, and expands the results into one table |

The main query keeps a **Source.Name** column with each row's file name, so you can always tell which monthly export a
row came from. The generated code looks like this (step names vary a little between Excel versions):

```
let
    Source = Folder.Files("C:\PQ\data\claims_monthly"),
    #"Filtered Hidden Files1" = Table.SelectRows(Source, each [Attributes]?[Hidden]? <> true),
    #"Invoke Custom Function1" = Table.AddColumn(#"Filtered Hidden Files1", "Transform File", each #"Transform File"([Content])),
    ...
```

> ⚠️ **From Folder includes every file in every subfolder.** A stray file breaks the combine or adds junk rows. Common
> culprits are a `readme.txt`, a copy of the workbook saved into the folder, a subfolder of old files, and on a Mac the
> hidden `.DS_Store` file that Finder creates.

> 💡 **Tip:** For a sturdier folder query, click **Transform Data** instead of **Combine** in step 3. Filter the
> **Extension** column to `.csv`, then click the **Combine Files** button (two down-arrows) in the **Content** column
> header. The filter step protects the query from any file that isn't a CSV.

> ⚠️ Every file must have the same column names. If a future export renames a column, the next refresh fails with
> *Expression.Error: The column 'PaidAmount' of the table wasn't found*, because the Changed Type step names each column.

### 12. Parameters: make the path portable

The Source step of every query you've built holds a full path such as `C:\PQ\data\claims_monthly`. Move the folder, send
the workbook to a colleague, or open it on a Mac, and every query fails with a *DataSource.NotFound* error. A
**parameter** stores the path once so you can change it in one place.

1. In the editor, click **Home → Manage Parameters → New Parameter**.
2. Set *Name* to `DataFolder`, *Type* to **Text**, and *Current Value* to your folder path **with** the final separator:
   `C:\PQ\data\` (Mac: `/Users/you/PQ/data/`). Click OK.
3. Select the Claims query, click its **Source** step, and change the formula bar to:

   ```
   = Folder.Files(DataFolder & "claims_monthly")
   ```

4. Make the same change in the **Sample File** helper query (in the Helper Queries group). Its Source step repeats the
   folder path, and Claims reads its column names from Sample File, so a stale path there still breaks the refresh.
5. In each single-file query, change the path inside `File.Contents(...)` the same way, for example
   `File.Contents(DataFolder & "payers.csv")`.

From then on, a new location means one edit: **Home → Manage Parameters**, change *Current Value*, then **Refresh All**.
The answer key's M code uses DataFolder from task 3 on.

> ⚠️ **On a Mac, skip the parameter.** Microsoft documents that Excel for Mac supports only literal (absolute) paths in
> a data source, so a Source step that builds its path from a parameter can fail there. Keep full paths such as
> `Folder.Files("/Users/you/PQ/data/claims_monthly")`, and when you paste answer-key code, replace
> `DataFolder & "payers.csv"` with the full path in quotes, such as `"/Users/you/PQ/data/payers.csv"`.

> 💡 **Tip:** Without a parameter, you can still repoint queries in bulk with **Data → Get Data → Data Source Settings**,
> select the old path, and click **Change Source…**.

> 📋 Some workbooks read the path from a worksheet cell with `Excel.CurrentWorkbook(){[Name="DataFolder"]}[Content]{0}[Column1]`,
> so people who never open the editor can change it. A step that uses a value from the workbook to open a file can
> trigger a *Formula.Firewall* error, because Power Query's privacy firewall controls how data from different sources is
> combined (section 13). That's one reason parameters are the simpler choice.

### 13. Merge queries (joins)

A **merge** is Power Query's lookup: it matches rows in two tables on one or more **key columns** and brings in data from
the matching rows. Unlike XLOOKUP, a merge runs once per refresh on the whole table, it can match on several columns, and it
can return all matching rows, not just the first.

1. Select the query you want to add columns to (for example Claims), then click **Home → Merge Queries** (or **Merge
   Queries as New** to keep the original unchanged).
2. Click the key column in the top table (PayerID), pick the second table (Payers), and click its key column (PayerID).
   To match on several columns, Ctrl-click (Mac: ⌘-click) them in the same order in both tables.
3. Choose the **Join Kind** and click OK. A new column of nested tables appears.
4. Click the **expand** button (two arrows) in the new column's header, tick only the columns you need (PayerType), and
   untick **Use original column name as prefix** unless you want the column named `Payers.PayerType`.

Before you click OK, read the line at the bottom of the Merge dialog. It reports how many rows of the first table found a
match, which is the quickest check that your keys line up.

| Join Kind | Keeps | Example with this lesson's data |
|---|---|---|
| **Left Outer** (default) | All rows of the first table, plus matches from the second | Every claim, with its payer's PayerType (task 7) |
| **Right Outer** | All rows of the second table, plus matches from the first | Every payer, with its claims if it has any |
| **Full Outer** | All rows from both tables | A reconciliation of two lists that should agree |
| **Inner** | Only rows that match in both | Claims whose PayerID appears in the payer list |
| **Left Anti** | Rows of the first table with **no** match | Encounters that have no claim yet (task 10) |
| **Right Anti** | Rows of the second table with **no** match | Payers with no claims in January 2025: only Workers' Compensation |

The M code is two steps, the join and the expand:

```
#"Merged Queries" = Table.NestedJoin(Claims, {"PayerID"}, Payers, {"PayerID"}, "Payers", JoinKind.LeftOuter),
#"Expanded Payers" = Table.ExpandTableColumn(#"Merged Queries", "Payers", {"PayerType"}, {"PayerType"})
```

> ⚠️ **Merges match exactly.** `"Timely Filing"` and `"Timely filing"` don't match because merges are case-sensitive, and
> `"Authorization Required "` with a trailing space doesn't match `"Authorization Required"`. Excel's XLOOKUP ignores case,
> so this surprises people. Clean keys first with **Transform → Format → Trim** and a consistent case, then check for
> nulls in the expanded column. (On Windows, the *Use fuzzy matching* option in the Merge dialog can also ignore case and
> near-misses, but exact keys are more reliable.)

> ⚠️ **Types must match too.** A Text key `"6110"` never matches a Whole Number key `6110`. Set both key columns to the
> same type before merging.

> ⚠️ **Duplicate keys multiply rows.** If the payer list had PY04 twice, every PY04 claim would appear twice after the
> merge. Lookup tables should have one row per key, so use **Remove Duplicates** on the key if you aren't sure.

> 📋 **Privacy levels.** When a query combines data from two different sources, such as the claims folder and
> `payers.csv` (task 7) or the claims folder and a table inside the workbook (task 9), Excel may show *Information is
> required about data privacy*. Click **Continue** and set each source to **Organizational**. On a computer you control,
> you can instead choose **Data → Get Data → Query Options → Current Workbook → Privacy → Ignore the Privacy Levels**.

### 14. Build a pipeline: Reference, Duplicate, and dependencies

Right-click a query in the Queries pane and you get two ways to copy it:

| Command | What the new query contains | Use it when |
|---|---|---|
| **Reference** | One step: `Source = Claims`. It starts from the *output* of the original query | You want a report built on a shared, cleaned base. Fix the base once and every reference gets the fix |
| **Duplicate** | A full copy of all the original's steps | You want an independent variation and don't mind maintaining two copies |

Good pipelines use one **staging query** per source, which imports and cleans and nothing else, and several report
queries that **reference** it. This lesson's pipeline ends up like this:

```
claims_monthly folder ──► Claims ──┬──► PaidTotal
                                   ├──► DenialsByReason
                                   ├──► PaidByPayerType ◄──── Payers (payers.csv)
                                   ├──► SubmitLag
                                   ├──► DenialsByOwner ◄───── DenialMap (tblDenialMap)
                                   ├──► AgedPending
                                   └──► Unbilled ◄─────────── Encounters2025 (encounters_2025.csv)
budget_2025_wide.csv ──► Budget2025 ──► ICU_Q1_Salaries
```

The bonus adds two more report queries, DenialDashboard and DenialRateByPayerType, that start from Claims and merge
Payers. **View → Query Dependencies** draws the same map for your workbook. It's the first place to look when you inherit a
workbook full of queries.

> 💡 **Tip:** Give queries short, meaningful names without spaces (Claims, Payers, DenialDashboard). Names appear in M
> code, and a name with spaces must be written `#"Denial Dashboard"` every time you refer to it.

### 15. Load options and refresh

**Home → Close & Load** loads the query as a Table on a new sheet. **Close & Load To…** (the dropdown under the same
button) lets you choose:

| Option | Result | Good for |
|---|---|---|
| **Table** | An Excel Table on a worksheet | Results people read, or that formulas and charts use |
| **PivotTable Report** | A PivotTable fed straight from the query | Summaries, without putting the detail rows on a sheet |
| **PivotChart** | A PivotChart (and its PivotTable) | Quick visuals |
| **Only Create Connection** | Nothing on a sheet. The query just runs when others need it | Staging queries such as Claims, Payers, and Encounters2025 |
| **Add this data to the Data Model** (check box) | Loads into Power Pivot's in-memory model (Windows only) | Millions of rows and multi-table models (Lesson 4.4) |

To change a query's load setting later, right-click it in the Queries & Connections pane and choose **Load To…**.

> 📋 A connection-only query shows *Connection only* in the Queries & Connections pane instead of "N rows loaded." When a
> task asks how many rows a query returns, load it to a Table, or reference it and click **Transform → Count Rows** in the
> editor. On a Mac, **Close & Load** puts each query on a worksheet as a Table. Depending on your version, **Close & Load
> To…** and **Only Create Connection** may be missing there, and the Data Model isn't available.

| To refresh | Do this |
|---|---|
| Every query in the workbook | **Data → Refresh All** (Windows: **Ctrl + Alt + F5**) |
| One loaded table | Click inside it and press **Alt + F5** (Windows), or right-click → **Refresh** |
| One query | Right-click it in Queries & Connections → **Refresh** |

Refresh All refreshes every query in the workbook. A report query re-runs the queries it references (Claims, Payers)
as part of its own refresh, so every report sees the files as they are right now. To refresh automatically, right-click a
query → **Properties…** and tick **Refresh data when opening the file**, or set **Refresh every** *n* minutes.

> ⚠️ Don't type over values inside a loaded query table, because the next refresh replaces them. To add your own
> calculation, add a formula column to the right of the loaded table. Excel keeps it as a calculated column and fills it
> for the new rows after each refresh.

> ⚠️ A worksheet holds at most 1,048,576 rows. For bigger results, load to a PivotTable, Only Create Connection, or the
> Data Model.

### 16. Read the M code

Every query is a short program in **M**, Power Query's formula language. Open it with **Home → Advanced Editor** (also on
the View tab). Here is a query that reads January's denials from the Claims query and counts them by reason:

```
let
    Source = Claims,                                                                   // 1
    JanuaryOnly = Table.SelectRows(Source, each [Source.Name] = "claims_2025_01.csv"), // 2
    DeniedOnly = Table.SelectRows(JanuaryOnly, each [ClaimStatus] = "Denied"),         // 3
    ByReason = Table.Group(DeniedOnly, {"DenialReason"},
        {{"Claims", each Table.RowCount(_), Int64.Type}}),                             // 4
    Sorted = Table.Sort(ByReason, {{"Claims", Order.Descending}})                      // 5
in
    Sorted                                                                             // 6
```

1. `let` opens a list of named steps. **Source** starts from another query, Claims.
2. Each step is `Name = expression,` and almost always uses the step above it. This one keeps the rows whose file name is
   January's export.
3. Steps run in order, so this filter sees only January's rows.
4. `Table.Group` takes the table, a **list** of grouping columns in braces, and a list of aggregations.
5. `Order.Descending` puts the biggest group first. Authorization Required (22) and Coding Error (18) top January's list.
6. `in` names the step whose result the query returns, usually the last one.

| Syntax | Meaning |
|---|---|
| `#"Changed Type"` | A step or query name that contains spaces or symbols |
| `each [ClaimStatus] = "Denied"` | A small function that runs once per row. `[ClaimStatus]` is that row's value |
| `_` | "The current item" inside `each`. `Table.RowCount(_)` counts the rows of the current group |
| `{"Jan", "Feb", "Mar"}` | A **list** |
| `[Delimiter=",", Encoding=65001]` | A **record** (named fields). Used for options |
| `Claims[PaidAmount]` | One column of a table, as a list |
| `Source{0}` | The first row of a table (counting starts at 0) |
| `#date(2025, 12, 31)` | A date value. `#datetime` and `#duration` work the same way |
| `// note` and `/* note */` | Comments, ignored when the query runs |

| Function family | Examples you met in this lesson |
|---|---|
| Sources | `Csv.Document`, `File.Contents`, `Folder.Files`, `Excel.CurrentWorkbook` |
| Table shaping | `Table.SelectRows`, `Table.SelectColumns`, `Table.RemoveColumns`, `Table.RenameColumns`, `Table.Sort`, `Table.TransformColumnTypes` |
| New columns | `Table.AddColumn`, `Table.SplitColumn` |
| Combining | `Table.Combine` (append), `Table.NestedJoin` + `Table.ExpandTableColumn` (merge) |
| Reshaping | `Table.Group`, `Table.UnpivotOtherColumns`, `Table.Pivot` |
| Values | `Text.Trim`, `Text.Proper`, `Text.Start`, `Date.Year`, `Duration.Days`, `List.Sum`, `List.Contains` |

**Create a query from M code.** Click **Data → Get Data → From Other Sources → Blank Query** (Mac: choose **Blank
query** in the **Get Data (Power Query)** dialog). The editor opens with an empty query. Type a name in the **Name** box,
click **Home → Advanced Editor**, select everything in the window, paste your code, and click **Done**. The line at the
bottom of the Advanced Editor says *No syntax errors have been detected*, or it names the first error it found. Task 11
starts this way.

You rarely need to write M from scratch. Reading it is what matters: it shows exactly what a query does, which hard-coded
values it depends on (a path, a date, a threshold), and where to change them. Small edits are easiest in the formula bar.

When a query fails, the message names the problem:

| Error | Usual cause | Fix |
|---|---|---|
| *DataSource.NotFound* | The file or folder moved or was renamed | Update the DataFolder parameter or **Data Source Settings → Change Source** |
| *Expression.Error: The column 'X' of the table wasn't found* | A source column was renamed or removed, and a step still names it | Fix the column name in the step (often Changed Type) |
| *Expression.Error: The import X matches no exports* | The code refers to a query that doesn't exist, for example `DenialMap` when the query is still named `tblDenialMap` | Rename the query or fix the reference (names are case-sensitive) |
| *Expression.Error: The name 'X' wasn't recognized* | A misspelled step or function name, such as `table.selectrows` | Fix the spelling and case |
| *DataFormat.Error: We couldn't convert to Number* | Text that isn't a number, or a locale mismatch | Fix the source, use **Change Type → Using Locale…**, or replace errors |
| *Formula.Firewall: … references other queries or steps, so it may not directly access a data source* | One step both opens a data source and uses data from another query or source, such as a path read from a worksheet cell | Use a parameter for the path, move the data access into its own query, or (on your own computer) ignore privacy levels (section 13) |

### 17. Version notes

| Excel | Power Query |
|---|---|
| Microsoft 365, Excel 2021, Excel 2024 (Windows) | Full Power Query Editor. Everything in this lesson works |
| Excel 2016 and 2019 (Windows) | Built in as **Get & Transform**. Early Excel 2016 builds call the entry point **Data → New Query** instead of Get Data |
| Excel 2010 and 2013 (Windows) | A free Microsoft add-in with its own **Power Query** tab. Same ideas, older menus |
| Microsoft 365 for Mac | Power Query Editor with Text/CSV, Table/Range, folder (added in 2024), merge, append, group, and unpivot. Start from **Data → Get Data (Power Query)**. Paths must be typed in full (no DataFolder parameter in a Source step), load destinations are limited, and there's no Data Model and no fuzzy merge. Older perpetual Mac versions such as Excel 2019 for Mac don't have the editor |
| Excel for the web | Can refresh some queries, and recent versions include a basic editor, but it can't read files on your computer. Use desktop Excel for this lesson |

> 📋 Mac support has grown with each update. If a command in this lesson is missing on your Mac, update Excel. Most steps
> can also be pasted as M code from the answer key into a blank query's Advanced Editor. On a Mac, replace each
> `DataFolder & "…"` with the full path in quotes, and use forward slashes instead of backslashes.

## 🧪 Hands-on practice

Open [`4.3-power-query.xlsx`](4.3-power-query.xlsx) and extract the [`data`](data/) folder first (Guide section 2). Build
each query in Power Query, read the result, and type it into the yellow cell on the **Practice** sheet. The **Check** column
turns green when you're right. The workbook's **DenialMap** sheet holds the table for task 9, and the **Starter M** sheet
holds the code for task 11 (also in [`starter/AgedPending.m`](starter/AgedPending.m)).

<!-- BEGIN GENERATED: practice -->
Build every query in this workbook from the CSV files in the lesson's data folder (Guide section 2), and do the tasks in order because later tasks reuse earlier queries. Name each query as the task says, because later tasks and the answer key use those names. Type each result in the yellow cell as a plain number or text. From task 3 on, the answer key's M code uses a DataFolder parameter (Guide section 12). If you skip the parameter, or work on a Mac, your code shows your full folder path in its place.

| # | Task | Hint |
|:-:|------|------|
| 1 | Import claims_2025_01.csv from the claims_monthly folder with Data → Get Data → From File → From Text/CSV and load it to a new sheet. How many claims (rows, not counting the header) does the January 2025 file contain? | After loading, the Queries & Connections pane says 'N rows loaded' |
| 2 | Import claims_2025_11.csv and claims_2025_12.csv as two more queries. Then, in the Power Query Editor, stack them with Home → Append Queries → Append Queries as New, and name the new query Claims_NovDec. How many rows does it return? | Append stacks rows, and columns line up by name |
| 3 | Now combine all twelve 2025 files at once. Choose Data → Get Data → From File → From Folder, select the claims_monthly folder, then Combine & Transform Data. Rename the new query Claims. How many rows does Claims return? | One folder query replaces twelve imports |
| 4 | Check that PaidAmount in Claims has the Decimal Number type (1.2 icon). Then reference Claims in a new query named PaidTotal and total its PaidAmount column. What is the total PaidAmount across all 2025 claims? Enter it to the cent. | Select PaidAmount, then Transform → Statistics → Sum |
| 5 | Reference Claims in a new query named DenialsByReason, keep only rows whose ClaimStatus is Denied, and use Home → Group By on DenialReason with the operation Count Rows. How many denied claims list Authorization Required as the reason? | Filter first, then Group By (Basic) |
| 6 | Edit that Group By step (gear icon → Advanced) and add two aggregations of BilledAmount: Sum and Average. Which denial reason has the highest AVERAGE BilledAmount per denied claim? Type the reason exactly as it appears. | Group By → Advanced → Add aggregation |
| 7 | Import payers.csv as a query named Payers. Then reference Claims in a new query named PaidByPayerType, merge it with Payers on PayerID (Join Kind: Left Outer), expand only PayerType, and group by PayerType with Sum of PaidAmount. What was the total PaidAmount for the Commercial payer type? Enter it to the cent. | Home → Merge Queries, then the expand button (two arrows) in the new column's header |
| 8 | Reference Claims in a new query named SubmitLag and add a custom column DaysToSubmit that holds the number of days from ServiceDate to SubmitDate, with the Whole Number type. How many 2025 claims took MORE than 30 days to submit? | Duration.Days turns a date difference into days. Then Number Filters → Greater Than, and count the rows |
| 9 | Load tblDenialMap (on the DenialMap sheet) with Data → From Table/Range and name the query DenialMap. Reference Claims in a new query named DenialsByOwner, keep the Denied rows, merge them with the map on DenialReason (Left Outer), and expand OwnerTeam. Check that EVERY denied claim found a match (no null OwnerTeam), and fix the keys in the map query if some didn't. What is the total denied BilledAmount owned by Patient Access? Enter it to the cent. | Merges match text exactly: look at Transform → Format → Trim and Capitalize Each Word |
| 10 | Import encounters_2025.csv (one row per encounter discharged in 2025) as a query named Encounters2025. Use Merge Queries as New with Encounters2025 on top, Claims below, EncounterID in both, and Join Kind Left Anti. Name the new query Unbilled. How many 2025 encounters have no claim in the 2025 claim files? | Left Anti = rows only in the first (top) table |
| 11 | Create a blank query (Data → Get Data → From Other Sources → Blank Query), name it AgedPending, open Home → Advanced Editor, and replace its contents with the starter code on the Starter M sheet (also in starter/AgedPending.m). It lists Pending claims more than 60 days old as of 12/31/2025. Read the code, change it to more than 90 days, and report how many Pending claims are more than 90 days old. | Find the step that compares DaysPending with 60 |
| 12 | Import budget_2025_wide.csv as a query named Budget2025. Remove the FY Total column, select the five label columns (FacilityID, Department, LineType, Category, Measure), and choose Transform → Unpivot Columns → Unpivot Other Columns. Rename Attribute to Month and Value to Amount. How many rows does Budget2025 return? | Unpivot Other Columns keeps the selected columns and unpivots the rest |
| 13 | In Budget2025, split Department (for example '6130 - Intensive Care Unit') into CostCenter and DeptName with Transform → Split Column → By Delimiter, using ' - ' (space, hyphen, space) at the left-most delimiter. Then reference Budget2025 in a new query named ICU_Q1_Salaries. What was the Q1 2025 (Jan–Mar) Actual Salaries & Wages for all three departments named Intensive Care Unit combined? Enter whole dollars. | Filter DeptName, Category, Measure, and Month, then Transform → Statistics → Sum on Amount (or Group By) |
<!-- END GENERATED: practice -->

## ✅ Answer key

The workbook has a hidden **Answer Key** sheet (right-click any sheet tab → **Unhide…** → *Answer Key*). Because the work
happens in Power Query, the key lists each result and the M code that produces it, with no live formulas. Each query is
also in the [`solutions`](solutions/) folder, ready to paste into the Advanced Editor. The answers are below, collapsed so
you don't see them by accident.

<!-- BEGIN GENERATED: answers -->
<details>
<summary><b>🔑 Show the answer key</b> — Try every task before opening this.</summary>

**1. Import one CSV file (January 2025)**

- **Answer:** 1,089
- **Solution:**

```powerquery
// Query: claims_2025_01
// Data > Get Data > From File > From Text/CSV > pick the file > Transform Data (or Load)
let
    Source = Csv.Document(File.Contents("C:\PQ\data\claims_monthly\claims_2025_01.csv"),[Delimiter=",", Columns=13, Encoding=65001, QuoteStyle=QuoteStyle.None]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"ClaimID", type text}, {"EncounterID", type text}, {"PatientID", type text}, {"PayerID", type text}, {"ServiceDate", type date}, {"SubmitDate", type date}, {"BilledAmount", type number}, {"AllowedAmount", type number}, {"PatientResponsibility", type number}, {"PaidAmount", type number}, {"ClaimStatus", type text}, {"DenialReason", type text}, {"PaidDate", type date}})
in
    #"Changed Type"
```


From Text/CSV writes three steps for you. **Source** reads the file with `Csv.Document`, **Promoted Headers** turns the first line into column names, and **Changed Type** sets each column's data type from the first 200 rows. The status bar inside the editor only counts the rows in the preview, so read the total from the Queries & Connections pane (or from the loaded Table) after **Close & Load**.

**2. Append two monthly files**

- **Answer:** 2,065
- **Solution:**

```powerquery
// Import claims_2025_11.csv and claims_2025_12.csv the same way as task 1, then
// Home > Append Queries > Append Queries as New > Two tables > claims_2025_11 + claims_2025_12
// Query: Claims_NovDec
let
    Source = Table.Combine({claims_2025_11, claims_2025_12})
in
    Source
```


Appending stacks the rows of one query under another (`Table.Combine`). It matches columns by **name**, not position, so files with the same headers line up even if the column order differs. The result has November's rows plus December's rows. Appending works for two or three files, but importing twelve files one by one doesn't scale, which is what task 3 fixes.

**3. Combine a folder of files (Claims)**

- **Answer:** 11,175
- **Solution:**

```powerquery
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
```


**From Folder** lists every file in the folder (and its subfolders), runs the same transformation on each one through the helper function *Transform File*, and appends the results. It also adds a **Source.Name** column with each row's file name, which is handy for tracing a row back to its export. Next month you drop a new file into the folder and click Refresh, with no new query needed.

**4. Total PaidAmount in Claims**

- **Answer:** 30,816,087.11
- **Solution:**

```powerquery
// Right-click Claims > Reference, rename the new query PaidTotal,
// select the PaidAmount column > Transform > Statistics > Sum. The preview shows one number.
let
    Source = Claims,
    #"Calculated Sum" = List.Sum(Source[PaidAmount])
in
    #"Calculated Sum"
// Alternative: load Claims to a sheet and type =SUM(Claims[PaidAmount]) in any empty cell.
```


**Reference** creates a new query whose source is the output of Claims, so you can summarize without changing Claims itself. **Statistics → Sum** turns the query into a single number (`List.Sum`). Never add that step to Claims, because every query built on Claims would then receive a number instead of a table. If PaidAmount were still Text (ABC icon), Sum would be grayed out.

**5. Group denied claims by reason**

- **Answer:** 297
- **Solution:**

```powerquery
// Query: DenialsByReason   (right-click Claims > Reference)
let
    Source = Claims,
    #"Filtered Rows" = Table.SelectRows(Source, each ([ClaimStatus] = "Denied")),
    #"Grouped Rows" = Table.Group(#"Filtered Rows", {"DenialReason"}, {{"Claims", each Table.RowCount(_), Int64.Type}}),
    #"Sorted Rows" = Table.Sort(#"Grouped Rows",{{"Claims", Order.Descending}})
in
    #"Sorted Rows"
```


Group By collapses the rows into one row per DenialReason and counts the rows in each group, like a PivotTable, but the result is a table that refreshes with the data. The filter must come **before** the Group By step: Applied Steps run top to bottom, and each step works on the output of the step above it. Authorization Required is the most common reason.

**6. Group By with several aggregations**

- **Answer:** Medical Necessity
- **Solution:**

```powerquery
// Query: DenialsByReason, Grouped Rows step edited (gear icon > Advanced > Add aggregation)
let
    Source = Claims,
    #"Filtered Rows" = Table.SelectRows(Source, each ([ClaimStatus] = "Denied")),
    #"Grouped Rows" = Table.Group(#"Filtered Rows", {"DenialReason"}, {
        {"Claims", each Table.RowCount(_), Int64.Type},
        {"BilledDenied", each List.Sum([BilledAmount]), type nullable number},
        {"AvgBilled", each List.Average([BilledAmount]), type nullable number}}),
    #"Sorted Rows" = Table.Sort(#"Grouped Rows",{{"AvgBilled", Order.Descending}})
in
    #"Sorted Rows"
```


Advanced Group By returns several summaries per group in one pass. Authorization Required has the most denials, but Medical Necessity denials average about $13,339 of billed charges each, the highest of any reason. Volume and dollars tell different stories, which is why denial reports show both counts and amounts.

**7. Merge Claims with Payers (PaidAmount by PayerType)**

- **Answer:** 14,941,777.08
- **Solution:**

```powerquery
// Query: Payers   (From Text/CSV: payers.csv)
let
    Source = Csv.Document(File.Contents(DataFolder & "payers.csv"),[Delimiter=",", Columns=6, Encoding=65001, QuoteStyle=QuoteStyle.None]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"PayerID", type text}, {"PayerName", type text}, {"PayerType", type text}, {"AvgAllowedPctOfCharges", type number}, {"AvgDaysToPay", Int64.Type}, {"TimelyFilingDays", Int64.Type}})
in
    #"Changed Type"

// Query: PaidByPayerType   (right-click Claims > Reference; Home > Merge Queries)
let
    Source = Claims,
    #"Merged Queries" = Table.NestedJoin(Source, {"PayerID"}, Payers, {"PayerID"}, "Payers", JoinKind.LeftOuter),
    #"Expanded Payers" = Table.ExpandTableColumn(#"Merged Queries", "Payers", {"PayerType"}, {"PayerType"}),
    #"Grouped Rows" = Table.Group(#"Expanded Payers", {"PayerType"}, {
        {"Claims", each Table.RowCount(_), Int64.Type},
        {"Paid", each List.Sum([PaidAmount]), type nullable number}})
in
    #"Grouped Rows"
```


A merge is Power Query's lookup. **Left Outer** keeps every claim and brings in the matching payer row as a nested table, and the expand button pulls out just the columns you need. Commercial combines 3 different payers, so you can't answer this by PayerID alone. You need the PayerType from the lookup table and then a Group By. Uncheck *Use original column name as prefix* when you expand, or the column is named Payers.PayerType.

**8. Custom column: days from service to submission**

- **Answer:** 362
- **Solution:**

```powerquery
// Query: SubmitLag   (right-click Claims > Reference; Add Column > Custom Column)
let
    Source = Claims,
    #"Added Custom" = Table.AddColumn(Source, "DaysToSubmit", each Duration.Days([SubmitDate] - [ServiceDate]), Int64.Type),
    #"Filtered Rows" = Table.SelectRows(#"Added Custom", each [DaysToSubmit] > 30)
in
    #"Filtered Rows"
```


Subtracting one date from another in M gives a **duration**, not a number. `Duration.Days` turns it into whole days. You can also build it without typing: select SubmitDate, Ctrl-click ServiceDate, then Add Column → Date → Subtract Days (the order you click sets which date comes first). Claims billed more than 30 days after service delay cash and risk timely-filing denials.

**9. Merge with an Excel Table (denials owned by Patient Access)**

- **Answer:** 4,542,077.42
- **Solution:**

```powerquery
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
```


Power Query merges are **exact and case-sensitive**. The hand-typed map has *Authorization Required* with a trailing space and *Timely filing* with a lower-case f, so a plain merge leaves 330 denied claims with a null OwnerTeam and undercounts Patient Access. Trim and Capitalize Each Word in the DenialMap query fix both keys. (Fixing the cells on the DenialMap sheet and refreshing works too.) Always check a merge by filtering the new column for null.

**10. Left Anti merge: encounters with no claim**

- **Answer:** 263
- **Solution:**

```powerquery
// Query: Encounters2025   (From Text/CSV: encounters_2025.csv)
let
    Source = Csv.Document(File.Contents(DataFolder & "encounters_2025.csv"),[Delimiter=",", Columns=10, Encoding=65001, QuoteStyle=QuoteStyle.None]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"EncounterID", type text}, {"PatientID", type text}, {"EncounterType", type text}, {"FacilityID", type text}, {"DeptID", type text}, {"AdmitDateTime", type datetime}, {"DischargeDateTime", type datetime}, {"PrimaryDxCode", type text}, {"PayerID", type text}, {"TotalCharges", type number}})
in
    #"Changed Type"

// Query: Unbilled   (Home > Merge Queries > Merge Queries as New:
// top = Encounters2025, bottom = Claims, click EncounterID in both, Join Kind = Left Anti)
let
    Source = Table.NestedJoin(Encounters2025, {"EncounterID"}, Claims, {"EncounterID"}, "Claims", JoinKind.LeftAnti),
    #"Removed Columns" = Table.RemoveColumns(Source, {"Claims"})
in
    #"Removed Columns"
```


A **Left Anti** join keeps only the rows in the first table that have **no** match in the second. It's the fastest way to answer "what's missing?" questions. These encounters were discharged but not yet billed (hospitals call this *discharged not final billed*, or DNFB), and together they carry $2,371,213.36 of charges. The bonus refreshes this list after the January 2026 claims arrive.

**11. Read and edit M code (aged pending claims)**

- **Answer:** 162
- **Solution:**

```powerquery
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
```


Each line between `let` and `in` is one Applied Step: a name, an equals sign, and an expression that usually uses the step above it. `AsOf` is a step that just holds a date, and the **Aged** step is the filter, so changing `> 60` to `> 90` is the whole edit. With 60 days the query returns 191 claims, and with 90 it returns 162. Claims pending that long need a follow-up call to the payer.

**12. Unpivot the wide budget**

- **Answer:** 5,928
- **Solution:**

```powerquery
// Query: Budget2025   (From Text/CSV: budget_2025_wide.csv)
// Select FY Total > Remove Columns; select FacilityID..Measure > Transform > Unpivot Columns > Unpivot Other Columns;
// rename Attribute to Month and Value to Amount
let
    Source = Csv.Document(File.Contents(DataFolder & "budget_2025_wide.csv"),[Delimiter=",", Columns=18, Encoding=65001, QuoteStyle=QuoteStyle.None]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"FacilityID", type text}, {"Department", type text}, {"LineType", type text}, {"Category", type text}, {"Measure", type text}, {"Jan", Int64.Type}, {"Feb", Int64.Type}, {"Mar", Int64.Type}, {"Apr", Int64.Type}, {"May", Int64.Type}, {"Jun", Int64.Type}, {"Jul", Int64.Type}, {"Aug", Int64.Type}, {"Sep", Int64.Type}, {"Oct", Int64.Type}, {"Nov", Int64.Type}, {"Dec", Int64.Type}, {"FY Total", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"FY Total"}),
    #"Unpivoted Other Columns" = Table.UnpivotOtherColumns(#"Removed Columns", {"FacilityID", "Department", "LineType", "Category", "Measure"}, "Attribute", "Value"),
    #"Renamed Columns" = Table.RenameColumns(#"Unpivoted Other Columns",{{"Attribute", "Month"}, {"Value", "Amount"}})
in
    #"Renamed Columns"
```


The wide file has 494 rows with a column per month. Unpivoting turns each month cell into its own row, so you get 494 × 12 = 5,928 rows with a Month and an Amount column. That long shape is what PivotTables, Group By, and the Data Model want. FY Total has to go first, or it would be unpivoted as a thirteenth 'month' and double every annual total. **Unpivot Other Columns** records the five label columns to keep (`Table.UnpivotOtherColumns`), so a future file with extra month columns still unpivots correctly.

**13. Split a column, then filter and sum (ICU Q1 salaries)**

- **Answer:** 3,483,594
- **Solution:**

```powerquery
// Query: Budget2025 after task 13: the task 12 query plus one step at the end. Select Department >
// Transform > Split Column > By Delimiter > --Custom-- " - " (space hyphen space) > Split at: Left-most delimiter.
// The dialog names the parts Department.1 and Department.2 (and adds a Changed Type step); typing the final
// names in the step, as here, splits and renames at once.
let
    Source = Csv.Document(File.Contents(DataFolder & "budget_2025_wide.csv"),[Delimiter=",", Columns=18, Encoding=65001, QuoteStyle=QuoteStyle.None]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"FacilityID", type text}, {"Department", type text}, {"LineType", type text}, {"Category", type text}, {"Measure", type text}, {"Jan", Int64.Type}, {"Feb", Int64.Type}, {"Mar", Int64.Type}, {"Apr", Int64.Type}, {"May", Int64.Type}, {"Jun", Int64.Type}, {"Jul", Int64.Type}, {"Aug", Int64.Type}, {"Sep", Int64.Type}, {"Oct", Int64.Type}, {"Nov", Int64.Type}, {"Dec", Int64.Type}, {"FY Total", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"FY Total"}),
    #"Unpivoted Other Columns" = Table.UnpivotOtherColumns(#"Removed Columns", {"FacilityID", "Department", "LineType", "Category", "Measure"}, "Attribute", "Value"),
    #"Renamed Columns" = Table.RenameColumns(#"Unpivoted Other Columns",{{"Attribute", "Month"}, {"Value", "Amount"}}),
    #"Split Column by Delimiter" = Table.SplitColumn(#"Renamed Columns", "Department", Splitter.SplitTextByEachDelimiter({" - "}, QuoteStyle.Csv, false), {"CostCenter", "DeptName"})
in
    #"Split Column by Delimiter"

// Query: ICU_Q1_Salaries   (right-click Budget2025 > Reference)
let
    Source = Budget2025,
    #"Filtered Rows" = Table.SelectRows(Source, each [DeptName] = "Intensive Care Unit"
        and [Category] = "Salaries & Wages" and [Measure] = "Actual"
        and List.Contains({"Jan", "Feb", "Mar"}, [Month])),
    #"Grouped Rows" = Table.Group(#"Filtered Rows", {"DeptName"}, {{"Q1Actual", each List.Sum([Amount]), type nullable number}})
in
    #"Grouped Rows"
```


Splitting at ' - ' (with the spaces) separates the cost center from the name without breaking names that contain a plain hyphen, such as *Medical-Surgical*. After the split, all three hospitals' ICUs share the DeptName *Intensive Care Unit* (the file has 24 distinct department names for 31 departments), so one filter catches all three. The unpivoted Month column makes "Q1" a simple filter on Jan, Feb, and Mar.

</details>
<!-- END GENERATED: answers -->

## 🏆 Bonus challenge

<!-- BEGIN GENERATED: bonus -->
It's the first week of February 2026 and the CFO wants a denial dashboard she can refresh every month without anyone rebuilding it. Build it on top of your Claims query. First, make the pipeline portable: if you haven't yet, create the DataFolder parameter (Guide section 12) and use it in the Source steps of Claims and its Sample File helper query (on a Mac, keep the full paths instead). Answer B1–B3 BEFORE you add the January 2026 file, then follow B4 and B5.

Work on the **Bonus** sheet of the workbook.

- **B1.** Build a query named DenialDashboard: reference Claims, keep Denied claims, merge Payers to get PayerType, and group by BOTH PayerType and DenialReason with Count Rows and Sum of BilledAmount. What is the denied BilledAmount for Government + Authorization Required? Enter it to the cent. *(Hint: Select PayerType and DenialReason before you click Group By, or use Add grouping in the Advanced dialog)*
- **B2.** Build a query named DenialRateByPayerType from ALL claims: reference Claims, merge Payers for PayerType, and add a conditional column IsDenied (1 if ClaimStatus is Denied, otherwise 0). Group by PayerType with two aggregations: AllClaims (Count Rows) and Denied (Sum of IsDenied). Which payer type has the highest denial rate (denied claims ÷ all claims)? *(Hint: After the Group By step, add a custom column DenialRate = [Denied] / [AllClaims] and sort it)*
- **B3.** What is that payer type's denial rate? Enter it as a percentage rounded to 1 decimal place. *(Hint: Click the type icon in the DenialRate header and choose Percentage)*
- **B4.** Now copy data\new_month\claims_2026_01.csv into the claims_monthly folder and click Data → Refresh All. You don't edit any query. What is the denied BilledAmount for Government + Authorization Required now? Enter it to the cent. *(Hint: Every query built on Claims re-reads the folder when you refresh)*
- **B5.** Your Unbilled query (task 10) refreshed too. How many 2025 encounters are still unbilled now that the January 2026 claims are in? *(Hint: Look at the Queries & Connections pane after the refresh)*
<!-- END GENERATED: bonus -->

<!-- BEGIN GENERATED: bonus-answers -->
<details>
<summary><b>🔑 Show the bonus solution</b> — Give it a real try first!</summary>

**B1. DenialDashboard by PayerType × DenialReason**

- **Answer:** 1,624,268.55
- **Solution:**

```powerquery
// Query: DenialDashboard   (right-click Claims > Reference)
let
    Source = Claims,
    #"Filtered Rows" = Table.SelectRows(Source, each ([ClaimStatus] = "Denied")),
    #"Merged Queries" = Table.NestedJoin(#"Filtered Rows", {"PayerID"}, Payers, {"PayerID"}, "Payers", JoinKind.LeftOuter),
    #"Expanded Payers" = Table.ExpandTableColumn(#"Merged Queries", "Payers", {"PayerType"}, {"PayerType"}),
    #"Grouped Rows" = Table.Group(#"Expanded Payers", {"PayerType", "DenialReason"}, {
        {"DeniedClaims", each Table.RowCount(_), Int64.Type},
        {"DeniedBilled", each List.Sum([BilledAmount]), type nullable number}}),
    #"Sorted Rows" = Table.Sort(#"Grouped Rows",{{"DeniedBilled", Order.Descending}})
in
    #"Sorted Rows"
```


Grouping by two columns gives one row per combination (20 rows here). Government + Authorization Required is the largest pool of denied dollars, so authorization work for Government-insured patients (Medicare and State Medicaid in this data) is where the revenue-cycle team should start.

**B2. Denial rate by payer type**

- **Answer:** Medicare Advantage
- **Solution:**

```powerquery
// Query: DenialRateByPayerType   (right-click Claims > Reference)
let
    Source = Claims,
    #"Merged Queries" = Table.NestedJoin(Source, {"PayerID"}, Payers, {"PayerID"}, "Payers", JoinKind.LeftOuter),
    #"Expanded Payers" = Table.ExpandTableColumn(#"Merged Queries", "Payers", {"PayerType"}, {"PayerType"}),
    #"Added Conditional Column" = Table.AddColumn(#"Expanded Payers", "IsDenied", each if [ClaimStatus] = "Denied" then 1 else 0, Int64.Type),
    #"Grouped Rows" = Table.Group(#"Added Conditional Column", {"PayerType"}, {
        {"AllClaims", each Table.RowCount(_), Int64.Type},
        {"Denied", each List.Sum([IsDenied]), type nullable number}}),
    #"Added Custom" = Table.AddColumn(#"Grouped Rows", "DenialRate", each [Denied] / [AllClaims], Percentage.Type),
    #"Sorted Rows" = Table.Sort(#"Added Custom",{{"DenialRate", Order.Descending}})
in
    #"Sorted Rows"
```


A rate needs two counts per group: all claims (Count Rows) and denied claims. Summing a 1/0 flag counts the rows where the condition is true, the same trick as SUMPRODUCT with booleans. Dividing in a custom column after the Group By gives the rate. In this data, Medicare Advantage denies about 1 claim in 10, the highest rate of any payer type, which mirrors what many US health systems report about Medicare Advantage plans.

**B3. Highest denial rate (value)**

- **Answer:** 10.2%
- **Solution:** Use the DenialRateByPayerType query from B2 and read the DenialRate value on the Medicare Advantage row. The Percentage type displays it as a percentage.

Medicare Advantage: 10.2% of its claims were denied. Typing 10.2 or 10.2% both pass the check.

**B4. Drop in the January 2026 file and Refresh All**

- **Answer:** 1,634,064.48
- **Solution:**

1. In File Explorer (Mac: Finder), copy **claims_2026_01.csv** from `data\new_month` into `data\claims_monthly`.
2. In Excel, click **Data → Refresh All** (Windows: Ctrl + Alt + F5).
3. Read the Government / Authorization Required row of the loaded DenialDashboard table.


Claims reads whatever files are in the folder, so the thirteenth file flows through every query that references it: Claims now returns 11,379 rows and the dashboard counts 869 denied claims instead of 848. This is the payoff of a folder-based pipeline. If Claims had a filter that kept only 2025 SubmitDates, the new file would have been silently ignored, so keep date filters out of the base query.

**B5. Unbilled encounters after the refresh**

- **Answer:** 59
- **Solution:** Read the row count of the refreshed **Unbilled** query in the Queries & Connections pane (or `=ROWS(Unbilled)` if you loaded it as a Table named Unbilled).

The January 2026 file billed 204 of the 263 encounters on the unbilled list, so 59 remain. (In the course's full claims data, their claims go out between February and May 2026, so later monthly files keep shrinking the list.) Because the Left Anti merge points at Claims, the unbilled list maintains itself as new files arrive. No one has to rerun a lookup.

</details>
<!-- END GENERATED: bonus-answers -->

## Key takeaways

- Power Query **records** your cleanup as Applied Steps and **replays** them on refresh, so a monthly report becomes a
  one-click job. The source files are never changed.
- Check **data types** right after import, keep IDs as Text, and use **Using Locale** when dates or decimals come from a
  different region.
- **From Folder** combines every file in a folder (and its subfolders), so the next month's export joins the pipeline when
  you drop it in and click **Refresh All**.
- **Append** stacks rows and matches columns by name. **Merge** looks up columns on key values, and its join kind decides
  which rows survive. **Left Anti** answers "what's missing?"
- Merge keys must match **exactly**, including case, spaces, and data type. Always check the expanded column for nulls.
- **Unpivot** wide month columns into rows before you analyze them, and remove total columns first.
- Build one cleaned staging query per source, **Reference** it for each report, and read the **M code** to see exactly
  what a query depends on.

<!-- BEGIN GENERATED: nav -->
---

⬅️ **Previous:** [4.2 Advanced Formulas: LET, LAMBDA & Array Logic](../02-advanced-formulas-let-lambda/README.md) · 🏠 [Course home](../../README.md) · **Next:** [4.4 Data Model, Power Pivot & DAX](../04-power-pivot-dax/README.md) ➡️
<!-- END GENERATED: nav -->
