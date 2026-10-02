# Learn Excel: From First Formula to VBA, with Healthcare Data

A free, hands-on Excel course that takes you from **"what's a cell?"** to **Power Query, DAX, dashboards, and VBA macros**.
Every lesson uses realistic (synthetic) data from **Bluestone Health System**, a fictional regional health system with three
hospitals and an outpatient pavilion. Instead of practicing on "Widget Sales," you'll practice on bed census, ED wait times, lab
results, insurance claims, readmissions, staffing, and supply inventory.

## How every lesson works

Each lesson folder contains a **README** (the lesson) and an **Excel workbook**.

| Part | Where | What you do |
|---|---|---|
| 📖 **Guide** | Lesson README | Learn the skill: explanations, syntax, worked examples, shortcuts, and pitfalls |
| 🧪 **Hands-on practice** | Workbook → **Practice** sheet | Solve 8–13 tasks on real data. Type answers in the yellow cells, and the **Check** column turns ✔ green when you're right |
| ✅ **Answer key** | Workbook → hidden **Answer Key** sheet, and a collapsed section in the README | Hidden by default. Unhide it only after you've tried (see below) |
| 🏆 **Bonus challenge** | Workbook → **Bonus** sheet | A harder, multi-step problem, with its own hidden **Bonus Key** |

### Revealing a hidden answer sheet

Right-click any sheet tab, choose **Unhide…**, pick **Answer Key** (or **Bonus Key**), and click **OK**. This works the same way
in Excel for Windows, Mac, and the web. In the README, answers sit inside a collapsed **🔑 Show the answer key** section.

> 💡 **Tip:** The Practice sheet's Check column compares your answer with the hidden key without showing it to you, so you can
> keep trying until it turns green.

## Syllabus

<!-- BEGIN GENERATED: syllabus -->
### [Level 1 · Foundations](01-foundations/README.md)

*Beginner* — Find your way around Excel, enter and format data, write your first formulas, and sort and filter a real dataset.

| # | Lesson | Time | You'll learn to… |
|:-:|---|:-:|---|
| 1.1 | [The Excel Interface & Navigation](01-foundations/01-excel-interface-navigation/README.md) | 30 min | name the parts of the Excel window: ribbon, Quick Access Toolbar, Name Box, formula bar, grid, sheet tabs, status bar; understand workbooks, worksheets, cells, ranges, and cell addresses; move around large datasets fast with keyboard shortcuts and Go To |
| 1.2 | [Data Entry, AutoFill & Editing](01-foundations/02-data-entry-autofill/README.md) | 40 min | recognize how Excel stores text, numbers, dates, times, and TRUE/FALSE; enter and edit data efficiently (F2, Ctrl+Enter, Ctrl+D, Alt+Enter); create series with AutoFill, the Fill Series dialog, and Flash Fill |
| 1.3 | [Formatting Cells & Number Formats](01-foundations/03-formatting-cells/README.md) | 45 min | apply number formats: Number, Currency vs Accounting, Percentage, Date, Time, Text; write custom number formats (leading zeros, units, colors, thousands); format fonts, fills, borders, alignment, and wrap text — and avoid merged cells |
| 1.4 | [Your First Formulas & Functions](01-foundations/04-basic-formulas/README.md) | 45 min | write formulas with operators and the correct order of operations; use SUM, AVERAGE, MIN, MAX, COUNT, COUNTA, and COUNTBLANK; calculate rates and percentages such as bed occupancy |
| 1.5 | [Relative, Absolute & Mixed References](01-foundations/05-cell-references/README.md) | 45 min | predict how relative references change when a formula is copied; lock references with \$ (absolute) and use F4 to toggle; build two-way grids with mixed references (\$A1 and A\$1) |
| 1.6 | [Sorting & Filtering Data](01-foundations/06-sorting-filtering/README.md) | 45 min | sort by one or several columns, including custom orders; filter text, numbers, and dates with AutoFilter (and Top 10, by color); summarize only visible rows with SUBTOTAL and AGGREGATE |

### [Level 2 · Formulas & Functions](02-formulas-functions/README.md)

*Beginner → Intermediate* — The core function families every analyst uses daily: logic, text, dates, statistics, conditional aggregation, and lookups.

| # | Lesson | Time | You'll learn to… |
|:-:|---|:-:|---|
| 2.1 | [Logical Functions: IF, AND, OR, IFS & More](02-formulas-functions/01-logical-functions/README.md) | 50 min | compare values to produce TRUE/FALSE and use booleans in math; make decisions with IF, nested IF, and IFS; combine conditions with AND, OR, NOT (and XOR) |
| 2.2 | [Text Functions](02-formulas-functions/02-text-functions/README.md) | 50 min | extract parts of text with LEFT, RIGHT, MID, FIND, and SEARCH; clean and standardize text with TRIM, CLEAN, UPPER, LOWER, PROPER, and SUBSTITUTE; join text with &, CONCAT, and TEXTJOIN; format numbers as text with TEXT |
| 2.3 | [Dates & Times](02-formulas-functions/03-date-time-functions/README.md) | 55 min | understand date serial numbers and times as fractions of a day; build and take apart dates with DATE, YEAR, MONTH, DAY, WEEKDAY, EOMONTH, and EDATE; calculate ages, lengths of stay, and turnaround times (DATEDIF, YEARFRAC, NETWORKDAYS, WORKDAY) |
| 2.4 | [Math & Statistical Functions](02-formulas-functions/04-math-statistical-functions/README.md) | 50 min | round correctly with ROUND, ROUNDUP, ROUNDDOWN, MROUND, CEILING.MATH, and FLOOR.MATH; use INT, TRUNC, MOD, ABS, and SUMPRODUCT; describe data with MEDIAN, MODE, STDEV, PERCENTILE, and QUARTILE — and know when the mean misleads |
| 2.5 | [Conditional Counting & Summing](02-formulas-functions/05-conditional-aggregation/README.md) | 55 min | count and sum with conditions using COUNTIF(S), SUMIF(S), and AVERAGEIF(S); find conditional extremes with MAXIFS and MINIFS; write criteria with operators, cell references, wildcards, and date ranges |
| 2.6 | [Lookups: VLOOKUP, INDEX/MATCH & XLOOKUP](02-formulas-functions/06-lookup-functions/README.md) | 60 min | look up exact matches with VLOOKUP and XLOOKUP; use approximate matches for tiers like age bands and BMI categories; combine INDEX and MATCH for flexible and two-way lookups |

### [Level 3 · Data Analysis](03-data-analysis/README.md)

*Intermediate* — Structure, validate, clean, summarize, and visualize data — Tables, PivotTables, charts, and what-if models.

| # | Lesson | Time | You'll learn to… |
|:-:|---|:-:|---|
| 3.1 | [Excel Tables, Structured References & Named Ranges](03-data-analysis/01-tables-named-ranges/README.md) | 50 min | convert ranges to Excel Tables and name them; write structured references like tblInventory[UnitCost] and [@QtyOnHand]; use calculated columns, the Total Row, and slicers on Tables |
| 3.2 | [Data Validation & Conditional Formatting](03-data-analysis/02-data-validation-conditional-formatting/README.md) | 55 min | restrict entries with list, number, date, length, and custom-formula validation; build dependent drop-down lists; highlight what matters with conditional formatting rules, data bars, color scales, and icon sets |
| 3.3 | [Cleaning Messy Data](03-data-analysis/03-data-cleaning/README.md) | 60 min | spot common data problems: stray spaces, inconsistent case and categories, text dates, duplicates; fix them with TRIM, PROPER, SUBSTITUTE, VALUE, DATEVALUE, and lookup mapping tables; use Text to Columns, Flash Fill, Remove Duplicates, and Go To Special |
| 3.4 | [PivotTables](03-data-analysis/04-pivottables/README.md) | 60 min | build PivotTables from a Table and arrange rows, columns, values, and filters; change summaries (Sum, Count, Average) and Show Values As (% of total, difference, running total); group dates and numbers; filter with slicers and timelines |
| 3.5 | [Charts & Data Visualization](03-data-analysis/05-charts-visualization/README.md) | 55 min | choose the right chart for comparisons, trends, parts of a whole, distributions, and relationships; build and format column, line, bar, scatter, histogram, combo, and Pareto charts; add trendlines, dynamic titles, and sparklines |
| 3.6 | [What-If Analysis: Goal Seek, Scenarios, Data Tables & Solver](03-data-analysis/06-what-if-analysis/README.md) | 60 min | structure a model with separate inputs, calculations, and outputs; find break-even points with Goal Seek; compare cases with Scenario Manager and sensitivity Data Tables |

### [Level 4 · Advanced Analysis](04-advanced-analysis/README.md)

*Advanced* — Modern Excel: dynamic arrays, LET/LAMBDA, Power Query, the Data Model with DAX, statistics, and dashboards.

| # | Lesson | Time | You'll learn to… |
|:-:|---|:-:|---|
| 4.1 | [Dynamic Arrays: FILTER, SORT, UNIQUE & More](04-advanced-analysis/01-dynamic-arrays/README.md) | 60 min | understand spilling, the # spill reference, and #SPILL! errors; extract lists with UNIQUE, FILTER, SORT, and SORTBY; generate sequences with SEQUENCE and reshape with TAKE, DROP, CHOOSECOLS, VSTACK, and HSTACK |
| 4.2 | [Advanced Formulas: LET, LAMBDA & Array Logic](04-advanced-analysis/02-advanced-formulas-let-lambda/README.md) | 65 min | write multi-condition array logic with SUMPRODUCT and boolean math; make complex formulas readable and fast with LET; create reusable custom functions with LAMBDA and the Name Manager |
| 4.3 | [Power Query: Import, Transform & Combine](04-advanced-analysis/03-power-query/README.md) | 70 min | import CSV files and whole folders with Get & Transform; clean and reshape data with Applied Steps: types, splits, filters, unpivot, group by; merge (join) and append queries |
| 4.4 | [Data Model, Power Pivot & DAX](04-advanced-analysis/04-power-pivot-dax/README.md) | 75 min | design a star schema and load tables into the Data Model; create relationships and a proper date table; write DAX measures with SUM, COUNTROWS, DISTINCTCOUNT, DIVIDE, and CALCULATE |
| 4.5 | [Statistics & Forecasting](04-advanced-analysis/05-statistics-forecasting/README.md) | 70 min | summarize distributions with the Analysis ToolPak and histograms; measure relationships with correlation and linear regression; test differences with t-tests and quantify uncertainty with confidence intervals |
| 4.6 | [Building Interactive Dashboards](04-advanced-analysis/06-dashboards/README.md) | 70 min | plan a dashboard around audience, questions, and KPIs; build KPI cards and selector-driven formulas; connect slicers to multiple PivotTables and drive charts from selections |

### [Level 5 · Automation with Macros & VBA](05-automation-vba/README.md)

*Expert* — Record macros, then write VBA: variables, loops, ranges, custom functions, error handling, events, and UserForms.

| # | Lesson | Time | You'll learn to… |
|:-:|---|:-:|---|
| 5.1 | [Recording Your First Macros](05-automation-vba/01-recording-macros/README.md) | 45 min | show the Developer tab and set macro security safely; record, run, and save macros in a macro-enabled workbook (.xlsm); choose between absolute and relative recording |
| 5.2 | [VBA Fundamentals](05-automation-vba/02-vba-fundamentals/README.md) | 60 min | navigate the Visual Basic Editor and organize code in modules; declare variables with the right data types and Option Explicit; control flow with If, Select Case, For, For Each, and Do loops |
| 5.3 | [VBA: Ranges, Worksheets & Workbooks](05-automation-vba/03-vba-ranges-worksheets/README.md) | 65 min | navigate the object model: Application, Workbooks, Worksheets, Range; find the last row and work with CurrentRegion, Offset, and Resize; loop through sheets; add, copy, rename, and delete them |
| 5.4 | [VBA: Custom Functions, Dictionaries & Error Handling](05-automation-vba/04-vba-functions-error-handling/README.md) | 65 min | write user-defined functions (UDFs) you can call from cells; pass arguments ByVal/ByRef, use Optional arguments, and return errors with CVErr; count and group with Collections and Scripting.Dictionary |
| 5.5 | [Events, UserForms & Automated Reports](05-automation-vba/05-events-userforms-automation/README.md) | 75 min | respond to workbook and worksheet events (Open, Change, BeforeSave); build a validated data-entry UserForm that writes to a Table; automate a report: combine files with Dir, export to PDF, save timestamped copies |

### [Level 6 · Capstone](06-capstone/README.md)

*Expert* — Bring every skill together in an end-to-end hospital performance review.

| # | Lesson | Time | You'll learn to… |
|:-:|---|:-:|---|
| 6.1 | [Capstone: Hospital Performance Review](06-capstone/01-hospital-performance-review/README.md) | 180 min | plan an analysis from business questions to deliverables; prepare multi-table data (cleaning, joins, calculated fields); analyze throughput, quality, utilization, and finance KPIs |

**30 lessons · about 30 hours of guided practice.**
<!-- END GENERATED: syllabus -->

## What you need

- **Microsoft Excel**: Microsoft 365 is recommended. Excel 2021 or 2019 works for most lessons, and each lesson notes features
  that need newer versions (XLOOKUP, dynamic arrays, LET/LAMBDA). Mac users can follow almost everything, and the lessons flag
  Windows-only features (Power Pivot, some VBA/UserForm details).
- **Excel for the web** works for Levels 1–3, but it can't run VBA macros or Power Pivot.
- **No prior experience.** Start at Lesson 1.1. If you already know the basics, skim the Level 1 practice sheets. If they're all
  green on the first try, jump ahead.

## The data

All datasets live in [`data/`](data/README.md) with a full **data dictionary** and a diagram of how the tables relate: patients,
encounters, ED visits, claims, lab results, medications, patient-satisfaction surveys, daily census, employees and shifts,
supply inventory, and budgets for January 2024 through December 2025 (about 200,000 rows in all). Each lesson workbook already
contains the slice of data it needs, so you don't have to download anything else.

> **All data is synthetic.** Bluestone Health System, its facilities, staff, and patients are fictional. Phone numbers use the
> 555 area code, and emails use `example.com`. ICD-10 diagnosis codes are real public codes used for realism. Nothing here is
> clinical or financial guidance.

## Repository layout

```
01-foundations/            Level 1 lessons (each lesson = README.md + practice workbook)
02-formulas-functions/     Level 2
03-data-analysis/          Level 3
04-advanced-analysis/      Level 4
05-automation-vba/         Level 5 (VBA starter code + solutions in each lesson)
06-capstone/               Final project
data/                      Full synthetic datasets + data dictionary
tools/                     Generators that build the data, workbooks, and answer keys (see tools/README.md)
```

## For contributors

Workbooks and answer keys are **generated**, never hand-edited. That way every answer is computed from the data and verified
by recalculating the workbook in LibreOffice. See [`tools/README.md`](tools/README.md) to regenerate the data, rebuild a lesson,
or write a new one.
