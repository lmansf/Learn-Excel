# Learn Excel: From First Formula to VBA, with Healthcare Data

A free, hands-on Excel course that takes you from **"what's a cell?"** to **Power Query, DAX, dashboards, and VBA macros**.
Every lesson uses realistic (synthetic) data from **Bluestone Health System**, a fictional regional health system with three
hospitals and an outpatient pavilion. Instead of practicing on "Widget Sales," you'll practice on bed census, ED wait times, lab
results, insurance claims, readmissions, staffing, and supply inventory.

## Getting started

1. Open [Lesson 1.1](01-foundations/01-excel-interface-navigation/README.md), or any lesson in the [syllabus](#syllabus).
2. Click the workbook link at the top of the lesson. GitHub shows a preview page, so click **Download raw file** to get the
   `.xlsx` file. To get the whole course at once, including the extra files that Lessons 4.3 and 5.2–6.1 use, click
   **Code → Download ZIP** on the repository's main page.
3. Open the file in Excel. If a yellow **Protected View** bar appears, click **Enable Editing**, because Excel opens
   downloaded files read-only until you do.
4. Read the **Start Here** sheet, follow the guide in the lesson README, then work through the **Practice** and **Bonus**
   sheets.

## How every lesson works

Each lesson folder contains a **README** (the lesson) and an **Excel workbook**. Each workbook opens on a **Start Here** sheet
that explains how to use it and what the cell colors mean.

| Part | Where | What you do |
|---|---|---|
| 📖 **Guide** | Lesson README | Learn the skill: explanations, syntax, worked examples, shortcuts, and pitfalls |
| 🧪 **Hands-on practice** | Workbook → **Practice** sheet | Solve 12–13 tasks on real data. Type answers in the yellow cells, and the **Check** column turns ✔ green when you're right. Gray cells read work you do on another sheet, so don't type over them |
| ✅ **Answer key** | Workbook → hidden **Answer Key** sheet, and a collapsed section in the README | Hidden by default. Unhide it only after you've tried (see below) |
| 🏆 **Bonus challenge** | Workbook → **Bonus** sheet | A harder, multi-step problem in 4–5 parts, with its own hidden **Bonus Key** |

### Revealing a hidden answer sheet

Right-click any sheet tab, choose **Unhide…**, pick **Answer Key** (or **Bonus Key**), and click **OK**. This works the same way
in Excel for Windows, Mac, and the web. Some lessons have extra hidden key sheets, such as **Chart Key** or **Dashboard Key**,
that hold reference charts, formats, or dashboards. In the README, answers sit inside a collapsed **🔑 Show the answer key**
section.

> 💡 **Tip:** The Practice sheet's Check column compares your answer with the hidden key without showing it to you, so you can
> keep trying until it turns green. The few build-it tasks that no formula can check show **See key** instead, so compare your
> work with the key.

## Syllabus

<!-- BEGIN GENERATED: syllabus -->
### [Level 1 · Foundations](01-foundations/README.md)

*Beginner* — Find your way around Excel, enter and format data, write your first formulas, and sort and filter a real dataset.

| # | Lesson | Time | You'll learn to… |
|:-:|---|:-:|---|
| 1.1 | [The Excel Interface & Navigation](01-foundations/01-excel-interface-navigation/README.md) | 70 min | name the parts of the Excel window: ribbon, Quick Access Toolbar, Name Box, formula bar, grid, sheet tabs, status bar; understand workbooks, worksheets, cells, ranges, and cell addresses; move around large datasets fast with keyboard shortcuts and Go To; read quick statistics from the status bar; freeze panes, zoom, and hide/unhide rows, columns, and sheets |
| 1.2 | [Data Entry, AutoFill & Editing](01-foundations/02-data-entry-autofill/README.md) | 80 min | recognize how Excel stores text, numbers, dates, times, and TRUE/FALSE; enter and edit data efficiently (F2, Ctrl + Enter, Ctrl + D, Alt + Enter); create series with AutoFill, the Fill Series dialog, and Flash Fill; use Copy, Paste Special (values, transpose, formats), and Find & Replace; avoid classic traps: lost leading zeros, numbers stored as text, accidental dates |
| 1.3 | [Formatting Cells & Number Formats](01-foundations/03-formatting-cells/README.md) | 100 min | apply number formats: Number, Currency vs Accounting, Percentage, Date, Time, Text; write custom number formats (leading zeros, units, colors, thousands); format fonts, fills, borders, alignment, and wrap text — and avoid merged cells; use Format Painter, cell styles, and themes for consistent reports; understand that formatting changes how a value looks, not the value itself |
| 1.4 | [Your First Formulas & Functions](01-foundations/04-basic-formulas/README.md) | 55 min | write formulas with operators and the correct order of operations; use SUM, AVERAGE, MIN, MAX, COUNT, COUNTA, and COUNTBLANK; calculate rates and percentages such as bed occupancy; copy formulas down a column and read common errors (#DIV/0!, #NAME?, #VALUE!); round results with ROUND |
| 1.5 | [Relative, Absolute & Mixed References](01-foundations/05-cell-references/README.md) | 85 min | predict how relative references change when a formula is copied; lock references with \$ (absolute) and use F4 to toggle; build two-way grids with mixed references (\$A1 and A\$1); reference other sheets and sum the same cell across sheets (3-D references) |
| 1.6 | [Sorting & Filtering Data](01-foundations/06-sorting-filtering/README.md) | 110 min | sort by one or several columns, including custom orders; filter text, numbers, and dates with AutoFilter (and Top 10, by color); summarize only visible rows with SUBTOTAL and AGGREGATE; use Advanced Filter for complex AND/OR criteria and unique lists |

### [Level 2 · Formulas & Functions](02-formulas-functions/README.md)

*Beginner → Intermediate* — The core function families every analyst uses daily: logic, text, dates, statistics, conditional aggregation, and lookups.

| # | Lesson | Time | You'll learn to… |
|:-:|---|:-:|---|
| 2.1 | [Logical Functions: IF, AND, OR, IFS & More](02-formulas-functions/01-logical-functions/README.md) | 105 min | compare values to produce TRUE/FALSE and use booleans in math; make decisions with IF, nested IF, and IFS; combine conditions with AND, OR, NOT (and XOR); map codes to labels with SWITCH; trap errors with IFERROR and IFNA |
| 2.2 | [Text Functions](02-formulas-functions/02-text-functions/README.md) | 110 min | extract parts of text with LEFT, RIGHT, MID, FIND, and SEARCH; clean and standardize text with TRIM, CLEAN, UPPER, LOWER, PROPER, and SUBSTITUTE; join text with &, CONCAT, and TEXTJOIN; format numbers as text with TEXT; use modern TEXTBEFORE, TEXTAFTER, and TEXTSPLIT (Microsoft 365 and Excel 2024) |
| 2.3 | [Dates & Times](02-formulas-functions/03-date-time-functions/README.md) | 115 min | understand date serial numbers and times as fractions of a day; build and take apart dates with DATE, YEAR, MONTH, DAY, WEEKDAY, EOMONTH, and EDATE; calculate ages, lengths of stay, and turnaround times (DATEDIF, YEARFRAC, NETWORKDAYS, WORKDAY); do time math for ED waits and overnight shifts (MOD, [h]:mm) |
| 2.4 | [Math & Statistical Functions](02-formulas-functions/04-math-statistical-functions/README.md) | 100 min | round correctly with ROUND, ROUNDUP, ROUNDDOWN, MROUND, CEILING.MATH, and FLOOR.MATH; use INT, TRUNC, MOD, ABS, and SUMPRODUCT; describe data with MEDIAN, MODE, STDEV, PERCENTILE, and QUARTILE — and know when the mean misleads; rank and pick extremes with RANK.EQ, LARGE, and SMALL |
| 2.5 | [Conditional Counting & Summing](02-formulas-functions/05-conditional-aggregation/README.md) | 110 min | count and sum with conditions using COUNTIF(S), SUMIF(S), and AVERAGEIF(S); find conditional extremes with MAXIFS and MINIFS; write criteria with operators, cell references, wildcards, and date ranges; build summary grids that fill with a single formula |
| 2.6 | [Lookups: VLOOKUP, INDEX/MATCH & XLOOKUP](02-formulas-functions/06-lookup-functions/README.md) | 125 min | look up exact matches with VLOOKUP and XLOOKUP; use approximate matches for tiers like age bands and BMI categories; combine INDEX and MATCH for flexible and two-way lookups; handle missing values with IFNA and XLOOKUP's if_not_found |

### [Level 3 · Data Analysis](03-data-analysis/README.md)

*Intermediate* — Structure, validate, clean, summarize, and visualize data — Tables, PivotTables, charts, and what-if models.

| # | Lesson | Time | You'll learn to… |
|:-:|---|:-:|---|
| 3.1 | [Excel Tables, Structured References & Named Ranges](03-data-analysis/01-tables-named-ranges/README.md) | 105 min | convert ranges to Excel Tables and name them; write structured references like tblInventory[UnitCost] and [@QtyOnHand]; use calculated columns, the Total Row, and slicers on Tables; create, manage, and use named ranges and named constants |
| 3.2 | [Data Validation & Conditional Formatting](03-data-analysis/02-data-validation-conditional-formatting/README.md) | 120 min | restrict entries with list, number, date, length, and custom-formula validation; build dependent drop-down lists; highlight what matters with conditional formatting rules, data bars, color scales, and icon sets; write formula-based rules that format entire rows |
| 3.3 | [Cleaning Messy Data](03-data-analysis/03-data-cleaning/README.md) | 145 min | spot common data problems: stray spaces, inconsistent case and categories, text dates, duplicates; fix them with TRIM, PROPER, SUBSTITUTE, VALUE, DATEVALUE, and lookup mapping tables; use Text to Columns, Flash Fill, Remove Duplicates, and Go To Special; document a repeatable cleaning process |
| 3.4 | [PivotTables](03-data-analysis/04-pivottables/README.md) | 130 min | build PivotTables from a Table and arrange rows, columns, values, and filters; change summaries (Sum, Count, Average) and Show Values As (% of total, difference, running total); group dates and numbers; filter with slicers and timelines; add calculated fields and use GETPIVOTDATA |
| 3.5 | [Charts & Data Visualization](03-data-analysis/05-charts-visualization/README.md) | 135 min | choose the right chart for comparisons, trends, parts of a whole, distributions, and relationships; build and format column, line, bar, scatter, histogram, combo, and Pareto charts; add trendlines, dynamic titles, and sparklines; design clear, accessible charts that tell one story |
| 3.6 | [What-If Analysis: Goal Seek, Scenarios, Data Tables & Solver](03-data-analysis/06-what-if-analysis/README.md) | 125 min | structure a model with separate inputs, calculations, and outputs; find break-even points with Goal Seek; compare cases with Scenario Manager and sensitivity Data Tables; optimize a staffing mix with Solver |

### [Level 4 · Advanced Analysis](04-advanced-analysis/README.md)

*Advanced* — Modern Excel: dynamic arrays, LET/LAMBDA, Power Query, the Data Model with DAX, statistics, and dashboards. Lesson 4.1 and parts of 4.2 need Microsoft 365 or Excel 2024, and Lesson 4.4 needs Excel for Windows.

| # | Lesson | Time | You'll learn to… |
|:-:|---|:-:|---|
| 4.1 | [Dynamic Arrays: FILTER, SORT, UNIQUE & More](04-advanced-analysis/01-dynamic-arrays/README.md) | 125 min | understand spilling, the # spill reference, and #SPILL! errors; extract lists with UNIQUE, FILTER, SORT, and SORTBY; generate sequences with SEQUENCE and reshape with TAKE, DROP, CHOOSECOLS, VSTACK, and HSTACK; combine functions into one-formula reports |
| 4.2 | [Advanced Formulas: LET, LAMBDA & Array Logic](04-advanced-analysis/02-advanced-formulas-let-lambda/README.md) | 150 min | write multi-condition array logic with SUMPRODUCT and boolean math; make complex formulas readable and fast with LET; create reusable custom functions with LAMBDA and the Name Manager; use MAP, REDUCE, SCAN, BYROW, and audit formulas with Evaluate Formula |
| 4.3 | [Power Query: Import, Transform & Combine](04-advanced-analysis/03-power-query/README.md) | 170 min | import CSV files and whole folders with Get & Transform; clean and reshape data with Applied Steps: types, splits, filters, unpivot, group by; merge (join) and append queries; build refreshable, repeatable data pipelines and read the M code behind them |
| 4.4 | [Data Model, Power Pivot & DAX](04-advanced-analysis/04-power-pivot-dax/README.md) | 175 min | design a star schema and load tables into the Data Model; create relationships and a proper date table; write DAX measures with SUM, COUNTROWS, DISTINCTCOUNT, DIVIDE, and CALCULATE; use time intelligence (TOTALYTD, SAMEPERIODLASTYEAR) and iterators (SUMX, AVERAGEX) |
| 4.5 | [Statistics & Forecasting](04-advanced-analysis/05-statistics-forecasting/README.md) | 150 min | summarize distributions with the Analysis ToolPak and histograms; measure relationships with correlation and linear regression; test differences with t-tests and quantify uncertainty with confidence intervals; forecast volumes and monitor processes with control charts |
| 4.6 | [Building Interactive Dashboards](04-advanced-analysis/06-dashboards/README.md) | 160 min | plan a dashboard around audience, questions, and KPIs; build KPI cards and selector-driven formulas; connect slicers to multiple PivotTables and drive charts from selections; polish layout, interactivity, and performance |

### [Level 5 · Automation with Macros & VBA](05-automation-vba/README.md)

*Expert* — Record macros, then write VBA: variables, loops, ranges, custom functions, error handling, events, and UserForms. You need desktop Excel for Windows or Mac, and you save each workbook as a macro-enabled .xlsm file.

| # | Lesson | Time | You'll learn to… |
|:-:|---|:-:|---|
| 5.1 | [Recording Your First Macros](05-automation-vba/01-recording-macros/README.md) | 110 min | show the Developer tab and set macro security safely; record, run, and save macros in a macro-enabled workbook (.xlsm); choose between absolute and relative recording; run macros from buttons, shortcuts, and the Quick Access Toolbar — and read the recorded code |
| 5.2 | [VBA Fundamentals](05-automation-vba/02-vba-fundamentals/README.md) | 160 min | navigate the Visual Basic Editor and organize code in modules; declare variables with the right data types and Option Explicit; control flow with If, Select Case, For, For Each, and Do loops; debug with breakpoints, stepping, the Immediate window, and Debug.Print |
| 5.3 | [VBA: Ranges, Worksheets & Workbooks](05-automation-vba/03-vba-ranges-worksheets/README.md) | 180 min | navigate the object model: Application, Workbooks, Worksheets, Range; find the last row and work with CurrentRegion, Offset, and Resize; loop through sheets; add, copy, rename, and delete them; process data fast with arrays, AutoFilter, and Sort — without Select |
| 5.4 | [VBA: Custom Functions, Dictionaries & Error Handling](05-automation-vba/04-vba-functions-error-handling/README.md) | 185 min | write user-defined functions (UDFs) you can call from cells; pass arguments ByVal/ByRef, use Optional arguments, and return errors with CVErr; count and group with Collections and Scripting.Dictionary; handle errors gracefully with On Error, the Err object, and cleanup code |
| 5.5 | [Events, UserForms & Automated Reports](05-automation-vba/05-events-userforms-automation/README.md) | 195 min | respond to workbook and worksheet events (Open, Change, BeforeSave); build a validated data-entry UserForm that writes to a Table; automate a report: combine files with Dir, export to PDF, save timestamped copies; know the safe limits of automation and when to use Office Scripts or Power Automate |

### [Level 6 · Capstone](06-capstone/README.md)

*Expert* — Bring every skill together in an end-to-end hospital performance review. The refresh macro needs desktop Excel for Windows or Mac.

| # | Lesson | Time | You'll learn to… |
|:-:|---|:-:|---|
| 6.1 | [Capstone: Hospital Performance Review](06-capstone/01-hospital-performance-review/README.md) | 240 min | plan an analysis from business questions to deliverables; prepare multi-table data (cleaning, joins, calculated fields); analyze throughput, quality, utilization, and finance KPIs; deliver a dashboard, an automated refresh macro, and an executive summary |

**30 lessons · about 65 hours of guided practice.**
<!-- END GENERATED: syllabus -->

## What you need

- **Microsoft Excel.** Microsoft 365 (Windows or Mac) is recommended. Excel 2024 also covers every practice task, and only a
  few side notes need Microsoft 365. Excel 2021 lacks LAMBDA, TEXTBEFORE/TEXTSPLIT, and VSTACK/TAKE/CHOOSECOLS, so Lesson 4.1
  and parts of Lessons 2.2 and 4.2 need Microsoft 365 or Excel 2024. Excel 2019 and earlier also lack XLOOKUP, dynamic arrays,
  and LET, which Lessons 2.6, 4.1, and 4.2 teach. The lessons show classic alternatives where they can, and each lesson's
  version notes say what works where.
- **On a Mac**, you can follow almost everything. Lesson 4.4 (Power Pivot and the Data Model) needs Excel for Windows, and the
  Power Query editor in Lesson 4.3 needs Microsoft 365 for Mac. Each lesson flags smaller Mac differences, such as the
  Windows-only Forecast Sheet button in Lesson 4.5.
- **Excel for the web** handles many lessons in Levels 1–4, but some tools the lessons use are missing or limited there:
  Advanced Filter (1.6), Circle Invalid Data (3.2), chart formatting (3.5), the what-if tools and Solver (3.6), Power Query
  on files from your computer (4.3), Power Pivot (4.4), the Analysis ToolPak (4.5), and Form Controls (4.6). It can't run
  VBA macros, so Level 5 and the capstone's refresh macro need desktop Excel for Windows or Mac.
- **No prior experience.** Start at [Lesson 1.1](01-foundations/01-excel-interface-navigation/README.md). If you already know
  the basics, skim the Level 1 practice sheets. If they're all green on the first try, jump ahead.

## The data

All datasets live in [`data/`](data/README.md) with a full **data dictionary** and a diagram of how the tables relate: patients,
encounters, ED visits, claims, lab results, medications, patient-satisfaction surveys, daily census, employees and shifts,
supply inventory, and budgets for January 2024 through December 2025 (about 180,000 rows in all). Each lesson workbook already
contains the data it needs. A few lessons also use files from their own folder: Power Query (4.3) imports CSV files from its
`data` folder, Lesson 5.5 combines the monthly census files in its `data` folder, and the VBA lessons (5.2–5.5) and the
capstone import starter code from a `starter` folder. Each of those lessons tells you what to download.

> **All data is synthetic.** Bluestone Health System, its facilities, staff, and patients are fictional. Phone numbers use the
> 555 area code, and emails use `example.com`. ICD-10 diagnosis codes are real public codes used for realism. Nothing here is
> clinical or financial guidance.

## Repository layout

```
01-foundations/            Level 1 lessons (each lesson = README.md + practice workbook)
02-formulas-functions/     Level 2
03-data-analysis/          Level 3
04-advanced-analysis/      Level 4 (Lesson 4.3 also has data, starter, and solutions folders)
05-automation-vba/         Level 5 (lesson folders also hold VBA starter code and solutions, and 5.5 has a data folder)
06-capstone/               Final project (with a starter and a solution macro)
data/                      Full synthetic datasets + data dictionary
tools/                     Generators that build the data, workbooks, and answer keys (see tools/README.md)
```

## For contributors

Workbooks and answer keys are **generated**, never hand-edited. That way every answer is computed from the data and verified
by recalculating the workbook in LibreOffice. See [`tools/README.md`](tools/README.md) to regenerate the data, rebuild a lesson,
or write a new one.
