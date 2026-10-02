"""Course curriculum: the single source of truth for lesson order, folders and scope.

Learner-facing fields (title, level, minutes, objectives) feed the course README syllabus and
each lesson's navigation links. Author-facing fields (spec) tell lesson authors exactly what
to teach, which data to use, what to practice, and what NOT to cover (because another lesson
owns it) so the course progresses without gaps or repetition.
"""
from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class LessonInfo:
    code: str
    module: str
    slug: str
    title: str
    minutes: int
    objectives: list[str]
    spec: str = ""
    level: str = ""

    @property
    def path(self) -> str:
        return f"{self.module}/{self.slug}"


MODULES = {
    "01-foundations": ("Level 1 · Foundations", "Beginner", "Find your way around Excel, enter and format data, write your first formulas, and sort and filter a real dataset."),
    "02-formulas-functions": ("Level 2 · Formulas & Functions", "Beginner → Intermediate", "The core function families every analyst uses daily: logic, text, dates, statistics, conditional aggregation, and lookups."),
    "03-data-analysis": ("Level 3 · Data Analysis", "Intermediate", "Structure, validate, clean, summarize, and visualize data — Tables, PivotTables, charts, and what-if models."),
    "04-advanced-analysis": ("Level 4 · Advanced Analysis", "Advanced", "Modern Excel: dynamic arrays, LET/LAMBDA, Power Query, the Data Model with DAX, statistics, and dashboards. Lesson 4.1 and parts of 4.2 need Microsoft 365 or Excel 2024, and Lesson 4.4 needs Excel for Windows."),
    "05-automation-vba": ("Level 5 · Automation with Macros & VBA", "Expert", "Record macros, then write VBA: variables, loops, ranges, custom functions, error handling, events, and UserForms. You need desktop Excel for Windows or Mac, and you save each workbook as a macro-enabled .xlsm file."),
    "06-capstone": ("Level 6 · Capstone", "Expert", "Bring every skill together in an end-to-end hospital performance review. The refresh macro needs desktop Excel for Windows or Mac."),
}

LESSONS: list[LessonInfo] = [
    # ------------------------------------------------------------------ Level 1
    LessonInfo("1.1", "01-foundations", "01-excel-interface-navigation", "The Excel Interface & Navigation", 70, [
        "Name the parts of the Excel window: ribbon, Quick Access Toolbar, Name Box, formula bar, grid, sheet tabs, status bar",
        "Understand workbooks, worksheets, cells, ranges, and cell addresses",
        "Move around large datasets fast with keyboard shortcuts and Go To",
        "Read quick statistics from the status bar",
        "Freeze panes, zoom, and hide/unhide rows, columns, and sheets",
    ], spec="""
GUIDE must cover: window anatomy (annotated ASCII/diagram table), workbook vs worksheet vs cell vs range, A1 addresses,
selecting (click, Shift+click, Ctrl+click, Ctrl+A, Ctrl+Space, Shift+Space), navigation shortcuts table for Windows AND Mac
(Ctrl/⌘ + arrows, Ctrl+Shift+arrows, Ctrl+Home/End, Page Up/Down, Alt+Page Up/Down, Ctrl+PgUp/PgDn to switch sheets, F5/Ctrl+G Go To,
Name Box jumps), status bar quick stats (Average/Count/Numerical Count/Min/Max/Sum — customize by right-click), Freeze Panes
(top row, first column, at selection), Split, Zoom, hide/unhide rows/columns/sheets, Find (Ctrl+F) incl. Find All count.
This lesson also teaches how to UNHIDE A SHEET — which learners need for every answer key in the course; make that prominent.
DATA: ~500 patients (selected columns) on a 'Patients' sheet; a second small sheet (e.g. 'Departments') that the builder
HIDES on purpose (learners unhide it for one task); optionally hide one column on Patients for an unhide task.
PRACTICE (12–13, no formulas required; learners type what they find): value in a specific cell via Name Box; last data row
number (Ctrl+Down); last used column letter; status-bar Sum/Average/Count of a stated selection (state exactly which cells);
a value found after unhiding the hidden sheet / hidden column; a Find All count (e.g. patients in a town).
Use check='text' for addresses/IDs, numbers for counts/stats (state rounding, e.g. "to 1 decimal").
BONUS: a multi-step 'scavenger hunt' combining Go To, Find All, freeze panes, and the status bar.
DO NOT teach formulas (Lesson 1.4) or formatting (1.3).
"""),
    LessonInfo("1.2", "01-foundations", "02-data-entry-autofill", "Data Entry, AutoFill & Editing", 80, [
        "Recognize how Excel stores text, numbers, dates, times, and TRUE/FALSE",
        "Enter and edit data efficiently (F2, Ctrl + Enter, Ctrl + D, Alt + Enter)",
        "Create series with AutoFill, the Fill Series dialog, and Flash Fill",
        "Use Copy, Paste Special (values, transpose, formats), and Find & Replace",
        "Avoid classic traps: lost leading zeros, numbers stored as text, accidental dates",
    ], spec="""
GUIDE: data types and how Excel decides (alignment clue: text left / numbers right), dates as serial numbers (light touch —
details in 2.3), entering times, the apostrophe for text (MRN '00412345'), editing (F2, Esc, Enter vs Tab direction),
Ctrl+Enter to fill a selection, Ctrl+D / Ctrl+R, Alt+Enter line break; AutoFill (fill handle drag & double-click; numbers,
dates, weekdays, months, custom lists, Ctrl-drag behavior), Home → Fill → Series (step/stop, weekdays only), Flash Fill (Ctrl+E)
basics, Copy/Cut/Paste, Paste Special (Values, Formats, Transpose, Multiply), Find & Replace (Match entire cell contents, Replace All
count), Undo/Redo, insert/delete rows & columns, Excel's auto-correct traps (gene/'MARCH1' style, '1-2' → date, long numbers →
scientific notation, 16+ digit truncation).
DATA: a December 2025 unit staffing-schedule template (blank yellow ranges to AutoFill), a short supply list, and a short list
of 'Last, First' names for Flash Fill. Use `summary` formulas + `fill` self-test values so AutoFill results are auto-checked
(e.g. summary '=IF(COUNTA(Schedule!A5:A35)=0,"",Schedule!A35)' expecting a date).
PRACTICE (12–13): AutoFill dates/weekdays-only, Fill Series with a step (15-minute appointment slots), ID series (e.g. SKU-1001…),
Flash Fill first names, Paste Special Transpose, Find & Replace count (state the exact replacement), entering an MRN with leading zeros.
BONUS: build a 2-week weekday clinic appointment grid (times 08:00–16:45 every 15 min × weekdays) entirely with AutoFill/Fill Series;
check counts and specific cells.
DO NOT teach formulas (1.4) or number formats beyond a mention (1.3).
"""),
    LessonInfo("1.3", "01-foundations", "03-formatting-cells", "Formatting Cells & Number Formats", 100, [
        "Apply number formats: Number, Currency vs Accounting, Percentage, Date, Time, Text",
        "Write custom number formats (leading zeros, units, colors, thousands)",
        "Format fonts, fills, borders, alignment, and wrap text — and avoid merged cells",
        "Use Format Painter, cell styles, and themes for consistent reports",
        "Understand that formatting changes how a value looks, not the value itself",
    ], spec="""
GUIDE: display vs stored value (the single most important idea), General format, Number/Currency/Accounting/Percentage/
Fraction/Scientific/Text/Date/Time formats, keyboard shortcuts (Ctrl+1, Ctrl+Shift+$ % # @ !), increase/decrease decimals,
custom number format syntax (sections positive;negative;zero;text, 0 vs # vs ?, comma scaling for thousands e.g. #,##0,"K",
literal text in quotes e.g. 0.0 "days", [Red]/(…) negatives, conditions like [<1]0.0%;…, leading zeros 00000000 for MRN,
phone (000) 000-0000, dates d/m/y/ddd/mmmm codes, [h]:mm for durations >24h), alignment (wrap, indent, rotate), Center Across
Selection vs Merge & Center (and why merging hurts), borders, fills, Format Painter (double-click to lock), cell styles,
themes, column width/row height autofit, clearing formats.
DATA: 2025 department budget summary (one row per department: DeptName, Budget, Actual, Variance, Variance %) as raw unformatted
numbers; a column of MRNs stored as numbers (lost zeros); some durations in days.
PRACTICE (12–13): mix of (a) "what does the cell display" text answers after applying a format (be explicit, e.g. value 0.0875 with
0.0% → '8.8%'); (b) "what custom format code…" text answers (use `accept` for equivalent codes); (c) stored-value questions
(a cell shows 4.7 — what is the actual value?); (d) manual formatting tasks checked against the key.
BONUS: format a 'Monthly Budget Report' to spec: thousands with K, negative variances red in parentheses, percentages 1 decimal,
header styling — answer key lists the format codes; auto-check what specific cells display (a text-answer question per format).
DO NOT teach conditional formatting (3.2) or the TEXT function (2.2) beyond a forward pointer.
"""),
    LessonInfo("1.4", "01-foundations", "04-basic-formulas", "Your First Formulas & Functions", 55, [
        "Write formulas with operators and the correct order of operations",
        "Use SUM, AVERAGE, MIN, MAX, COUNT, COUNTA, and COUNTBLANK",
        "Calculate rates and percentages such as bed occupancy",
        "Copy formulas down a column and read common errors (#DIV/0!, #NAME?, #VALUE!)",
        "Round results with ROUND",
    ], spec="(Reference lesson written by the course lead. Use it as the pattern for README structure and task design. Its guide is shorter than most, so see tools/README.md for the expected depth.)"),
    LessonInfo("1.5", "01-foundations", "05-cell-references", "Relative, Absolute & Mixed References", 85, [
        "Predict how relative references change when a formula is copied",
        "Lock references with $ (absolute) and use F4 to toggle",
        "Build two-way grids with mixed references ($A1 and A$1)",
        "Reference other sheets and sum the same cell across sheets (3-D references)",
    ], spec="""
GUIDE: what happens when you copy a formula (relative shifts), absolute $A$1, mixed $A1 / A$1, F4 cycling (Mac: Cmd+T or
Fn+F4), % of total pattern, rate-in-one-cell pattern, two-way grid pattern, cross-sheet references (Sheet!A1, 'Sheet Name'!A1),
3-D references SUM(Jan:Mar!B5) and their gotchas (sheet order), referencing other workbooks (brief), Ctrl+` show formulas,
Trace Precedents (brief). Name a cell (brief preview; full named ranges in 3.1).
DATA: (1) 2025 department expense summary for % of total; (2) staffing hours with a single 'overtime premium' / hourly rate
cell; (3) three monthly sheets Oct/Nov/Dec 2025 with identical layout (department × category actuals from budget.csv) for 3-D sums;
(4) a staffing grid: census values down rows and HPPD (hours per patient day) targets across columns.
PRACTICE (12–13): % of total for a stated row; cost with a fixed rate cell; fill a grid with ONE formula using mixed refs
(summary = SUM of grid + fill self-test with the formula); 3-D SUM total; a 'predict what the formula becomes when copied to X'
text task (e.g. '=B$2*$A3' copied two right, three down → '=D$2*$A6').
BONUS: 2026 budget projection grid: 2025 actual × (1+category inflation row) × (1+volume growth column) with one formula;
check the grand total and two specific cells.
DO NOT teach named ranges in depth (3.1) or lookup functions (2.6).
"""),
    LessonInfo("1.6", "01-foundations", "06-sorting-filtering", "Sorting & Filtering Data", 110, [
        "Sort by one or several columns, including custom orders",
        "Filter text, numbers, and dates with AutoFilter (and Top 10, by color)",
        "Summarize only visible rows with SUBTOTAL and AGGREGATE",
        "Use Advanced Filter for complex AND/OR criteria and unique lists",
    ], spec="""
GUIDE: sort basics (A→Z, Z→A, Data → Sort dialog, add level, sort by cell color, custom lists e.g. weekday order or
'Discharged, Admitted, Transferred, LWBS'), header row checkbox, selecting the whole dataset (why sorting one column breaks rows),
AutoFilter (Ctrl+Shift+L), text/number/date filters, search box, Top 10, filter by color, multiple-column filters (AND), clearing,
reapplying; status bar count of visible rows; SUBTOTAL(9/101-111 family) vs SUM on filtered data, AGGREGATE (ignore errors &
hidden rows); Advanced Filter (criteria range, AND on same row, OR on separate rows, copy to another location, unique records only).
DATA: ~1,000 ED visits from 2025 at one facility with a precomputed DoorToProviderMin column (blank for LWBS), ESILevel,
ArrivalMode, ChiefComplaint, EDDisposition, ArrivalDateTime.
PRACTICE (12–13): after a stated multi-level sort, what is the EDVisitID in the first/10th row (text); filtered counts
(e.g. Ambulance arrivals with ESI 1–2); Top 10 by wait — the 10th largest value; SUBTOTAL average of a filtered set;
date filter count (e.g. arrivals in November 2025 on weekends — state how); Advanced Filter unique chief complaint count.
BONUS: Advanced Filter with OR logic (ESI 1, OR ESI 2 arriving by ambulance after 6 pm) copied to a new sheet — count and average
door-to-provider minutes; and a custom-list sort question.
DO NOT teach COUNTIFS (2.5) except as a forward pointer; formulas only via SUBTOTAL/AGGREGATE.
"""),
    # ------------------------------------------------------------------ Level 2
    LessonInfo("2.1", "02-formulas-functions", "01-logical-functions", "Logical Functions: IF, AND, OR, IFS & More", 105, [
        "Compare values to produce TRUE/FALSE and use booleans in math",
        "Make decisions with IF, nested IF, and IFS",
        "Combine conditions with AND, OR, NOT (and XOR)",
        "Map codes to labels with SWITCH; trap errors with IFERROR and IFNA",
    ], spec="""
GUIDE: comparison operators, TRUE/FALSE results, booleans as 1/0 (--), IF syntax, nested IF readability, IFS (and its
TRUE catch-all), AND/OR/NOT/XOR, combining inside IF, SWITCH, IFERROR vs IFNA (and why not to hide every error), common
mistakes (text numbers, ="" blanks, order of nested conditions). Clinical examples: abnormal lab flags, SIRS criteria
(Temp >100.4°F or <96.8°F; HR >90; RR >20 — state the simplified rule clearly as 'educational, not clinical guidance').
DATA: ~400 lab results (TestName, ResultValue, RefLow, RefHigh) and ~500 ED visits with triage vitals and ESILevel.
Add empty learner columns (extra_cols) where tasks ask to fill a column; check with `summary` + `fill`.
PRACTICE (12–13): High/Low/Normal flag column via nested IF (summary = COUNTIF of "High"); IFS version; AND/OR fever+tachycardia
count; SIRS criteria count (0–3) and SIRS-positive (≥2) count; SWITCH ESI level names; IFERROR on a ratio; NOT; an XOR question.
BONUS: a qSOFA-style screening score (RR ≥22, SBP ≤100, altered mentation if ChiefComplaint contains 'Altered' — can use
ISNUMBER(SEARCH()) with a short explanation) → count of patients with score ≥2 and their average heart rate.
DO NOT teach COUNTIFS (2.5) beyond COUNTIF for summaries; text functions only minimal.
"""),
    LessonInfo("2.2", "02-formulas-functions", "02-text-functions", "Text Functions", 110, [
        "Extract parts of text with LEFT, RIGHT, MID, FIND, and SEARCH",
        "Clean and standardize text with TRIM, CLEAN, UPPER, LOWER, PROPER, and SUBSTITUTE",
        "Join text with &, CONCAT, and TEXTJOIN; format numbers as text with TEXT",
        "Use modern TEXTBEFORE, TEXTAFTER, and TEXTSPLIT (Microsoft 365 and Excel 2024)",
    ], spec="""
GUIDE: text vs numbers, LEN, LEFT/RIGHT/MID, FIND (case-sensitive) vs SEARCH (wildcards, case-insensitive), nesting FIND inside
MID/LEFT, TRIM/CLEAN (and CHAR(160) non-breaking spaces via SUBSTITUTE), UPPER/LOWER/PROPER (PROPER pitfalls: McDonald, O'Brien),
& and CONCAT/TEXTJOIN (ignore_empty), SUBSTITUTE vs REPLACE, TEXT(value, format_text) for dates & numbers, VALUE/NUMBERVALUE,
EXACT, REPT, TEXTBEFORE/TEXTAFTER/TEXTSPLIT with version notes. Healthcare: ICD-10 category = LEFT(code,3), MRN padding with
TEXT(…,"00000000"), 'LAST, First' names, phone digits, email domains, ChronicConditions list 'HTN;DM;HF'.
DATA: ~400 patients with a builder-made 'PatientName' ('LAST, First' with mixed case), MRN as number, Phone, Email,
ChronicConditions; ~300 encounters with PrimaryDxCode; providers for label building.
PRACTICE (12–13): first name from 'Last, First'; ICD category; email domain; digits-only phone (SUBSTITUTE chain); provider label
'Dr. First Last, MD'; MRN padded to 8 chars; count of patients whose ChronicConditions contain 'DM' (ISNUMBER(SEARCH)) or via
LEN-SUBSTITUTE counting of items; TEXT a date as 'Mon-YYYY'. Use summary+fill for column tasks; single-cell answers for others.
BONUS: parse the messy 'CityStateZip' strings (from data/messy) into City, State, ZIP (5 digits even when ZIP+4) using
FIND/MID or TEXTBEFORE/TEXTAFTER — check counts/specific outputs; plus count of chronic conditions per patient.
DO NOT teach data cleaning workflow tools (Text to Columns, Remove Duplicates → 3.3) beyond a pointer.
"""),
    LessonInfo("2.3", "02-formulas-functions", "03-date-time-functions", "Dates & Times", 115, [
        "Understand date serial numbers and times as fractions of a day",
        "Build and take apart dates with DATE, YEAR, MONTH, DAY, WEEKDAY, EOMONTH, and EDATE",
        "Calculate ages, lengths of stay, and turnaround times (DATEDIF, YEARFRAC, NETWORKDAYS, WORKDAY)",
        "Do time math for ED waits and overnight shifts (MOD, [h]:mm)",
    ], spec="""
GUIDE: serial numbers (1 = 1/1/1900; 1900 vs 1904 system note for Mac legacy), times as fractions, why TODAY()/NOW() are volatile
and why reports use a fixed 'Report Date' cell (course as-of date is 12/31/2025), DATE, YEAR/MONTH/DAY, WEEKDAY (return types),
TEXT(date,"dddd"), EOMONTH/EDATE, DATEDIF (undocumented; 'Y','M','D','YM'), YEARFRAC, DAYS, NETWORKDAYS(.INTL) and WORKDAY(.INTL)
with holiday lists, date arithmetic (LOS = discharge − admit; INT for calendar days; midnights = INT(d)−INT(a)), time math
(×24 hours, ×1440 minutes), [h]:mm formats, overnight shifts =MOD(out−in,1), DATEVALUE/TIMEVALUE for text dates (light).
DATA: ~300 inpatient encounters (Admit/Discharge datetimes + patient DOB pre-joined), ~300 ED visits (timestamps),
~200 shifts including night shifts, ~300 claims (Service/Submit/Paid dates), a small 2025 holiday list.
PRACTICE (12–13): age at admission (DATEDIF); LOS days (decimal) and calendar days; # of weekend admissions; month-end of a
service date; days from submit to paid (average); door-to-provider minutes for a stated visit and the average; worked hours
of a night shift with MOD; a WORKDAY appeal deadline with holidays; NETWORKDAYS count.
BONUS: CMS-style 'two-midnight' check: count stays crossing ≥2 midnights; and average LOS in hours by weekday of admission
(state which weekday's value to enter).
DO NOT teach COUNTIFS/AVERAGEIFS in depth (2.5) — simple COUNTIF/AVERAGE over helper columns is fine.
"""),
    LessonInfo("2.4", "02-formulas-functions", "04-math-statistical-functions", "Math & Statistical Functions", 100, [
        "Round correctly with ROUND, ROUNDUP, ROUNDDOWN, MROUND, CEILING.MATH, and FLOOR.MATH",
        "Use INT, TRUNC, MOD, ABS, and SUMPRODUCT",
        "Describe data with MEDIAN, MODE, STDEV, PERCENTILE, and QUARTILE — and know when the mean misleads",
        "Rank and pick extremes with RANK.EQ, LARGE, and SMALL",
    ], spec="""
GUIDE: rounding family with a comparison table (incl. negative digits, MROUND to 15-minute billing units, CEILING.MATH for whole
vials/doses), INT vs TRUNC on negatives, MOD (e.g. even/odd, minutes past hour), ABS, SUM vs SUMPRODUCT (dose × unit cost),
center (AVERAGE, MEDIAN, MODE.SNGL, TRIMMEAN) and spread (MIN/MAX/range, STDEV.S vs STDEV.P, VAR.S), PERCENTILE.INC/EXC,
QUARTILE.INC, PERCENTRANK, RANK.EQ vs RANK.AVG, LARGE/SMALL; skewed healthcare data (charges, LOS) → median; precision vs display.
DATA: ~600 inpatient encounters from 2025 with LOS (days) and TotalCharges; ~800 ED visits with DoorToProviderMin; ~300
medication orders (DosesDispensed, UnitCost).
PRACTICE (12–13): mean vs median LOS; mode of ESI; STDEV.S charges; 90th percentile door-to-provider; Q1/Q3; rank of a given
encounter's charge; 3rd largest; MROUND minutes to 15; CEILING.MATH vials; SUMPRODUCT total med cost; ROUNDUP vs ROUNDDOWN result.
BONUS: outlier analysis of charges: IQR fences (Q1−1.5·IQR, Q3+1.5·IQR) → count of high outliers; compare with mean+3·SD rule;
TRIMMEAN(…,10%).
DO NOT teach conditional versions (AVERAGEIFS etc. → 2.5) or regression/forecasting (4.5).
"""),
    LessonInfo("2.5", "02-formulas-functions", "05-conditional-aggregation", "Conditional Counting & Summing", 110, [
        "Count and sum with conditions using COUNTIF(S), SUMIF(S), and AVERAGEIF(S)",
        "Find conditional extremes with MAXIFS and MINIFS",
        "Write criteria with operators, cell references, wildcards, and date ranges",
        "Build summary grids that fill with a single formula",
    ], spec="""
GUIDE: COUNTIF → COUNTIFS (AND logic), OR logic by adding COUNTIFS or SUM(COUNTIFS(…,{"a","b"})), SUMIF vs SUMIFS argument order,
AVERAGEIF(S), MAXIFS/MINIFS (2019+), criteria strings (">=100", "<>"&A1, "*sepsis*", "?" wildcard, "~*" escape), dates
(">="&DATE(2025,7,1)), blanks ("", "<>"), criteria ranges must be same size, building a two-way summary grid with mixed
references (links back to 1.5), rates = COUNTIFS(numerator)/COUNTIFS(denominator), divide-by-zero guard.
DATA: ~3,000 2025 encounters (with FacilityName, DeptName, PayerName, DxCategory, LOS pre-joined) and their claims
(ClaimStatus, DenialReason, BilledAmount, PaidAmount, SubmitDate, PaidDate).
PRACTICE (12–13): counts by type & facility; total charges for a payer; average LOS for heart failure; denied claims count for a
reason; claims submitted in Q3; MAXIFS charges for sepsis; wildcard count of circulatory (I*) diagnoses; a facility × type grid
filled with one formula (summary = grid total, plus one specific cell task); readmission rate for a service line.
BONUS: payer scorecard — denial rate per payer, average days to pay (helper column), and which payer has the highest denial rate
(text answer) plus its rate.
DO NOT use PivotTables (3.4) or SUMPRODUCT array tricks (4.2).
"""),
    LessonInfo("2.6", "02-formulas-functions", "06-lookup-functions", "Lookups: VLOOKUP, INDEX/MATCH & XLOOKUP", 125, [
        "Look up exact matches with VLOOKUP and XLOOKUP",
        "Use approximate matches for tiers like age bands and BMI categories",
        "Combine INDEX and MATCH for flexible and two-way lookups",
        "Handle missing values with IFNA and XLOOKUP's if_not_found",
    ], spec="""
GUIDE: why lookups (joining tables), VLOOKUP (exact FALSE, approximate TRUE on sorted ascending, col_index pitfalls, can't look
left), HLOOKUP (brief), INDEX, MATCH (0/1/-1), INDEX+MATCH (left lookups, two-way INDEX(…,MATCH(),MATCH())), XLOOKUP (all
arguments: if_not_found, match_mode 0/-1/1/2 wildcard, search_mode 1/-1 last match, returning multiple columns), XMATCH,
IFNA vs IFERROR, approximate tier tables (age bands, BMI categories <18.5, 18.5–24.9, 25–29.9, 30+), lookups across sheets,
performance & robustness tips, version notes (XLOOKUP needs 365/2021+).
DATA: ~500 encounters (IDs only) + lookup sheets: Patients, Providers, Diagnoses, Payers, Departments; a BMI tier table; a
2-D table of 2025 budget by department × category.
PRACTICE (12–13): patient last name (VLOOKUP) for a stated encounter; dx description (XLOOKUP); payer type; attending specialty
(INDEX/MATCH); fill a column with payer names (summary = COUNTIF of one payer); BMI category with approximate match; two-way budget
lookup; most recent encounter date for a patient (XLOOKUP search_mode -1); missing-ID handling returning 'Not found'.
BONUS: an 'Encounter lookup card' — type an EncounterID → patient name, age at admission, dx description, attending name &
specialty, payer, claim status; tasks ask for the outputs for two stated IDs (one invalid → 'Not found').
DO NOT teach dynamic-array FILTER (4.1).
"""),
    # ------------------------------------------------------------------ Level 3
    LessonInfo("3.1", "03-data-analysis", "01-tables-named-ranges", "Excel Tables, Structured References & Named Ranges", 105, [
        "Convert ranges to Excel Tables and name them",
        "Write structured references like tblInventory[UnitCost] and [@QtyOnHand]",
        "Use calculated columns, the Total Row, and slicers on Tables",
        "Create, manage, and use named ranges and named constants",
    ], spec="""
GUIDE: why Tables (auto-expand, banding, filters, calculated columns, structured refs, pivot sources), Ctrl+T / Format as Table,
naming tables (Table Design → Table Name; tbl prefix convention), structured reference syntax (Table[Col], [@Col], Table[#All],
[#Headers], [#Totals], [[#This Row],[Col]], column ranges [[ColA]:[ColB]]), calculated columns (auto-fill, undo auto-fill),
Total Row and SUBTOTAL, slicers on Tables, converting back to range; Named ranges (Name Box, Formulas → Define Name, Create from
Selection, scope workbook vs sheet, Name Manager, named constants like ReportDate / ExpiringWindowDays), F3 paste names,
dynamic named ranges (OFFSET vs INDEX) vs Tables.
IMPORTANT TECH NOTE: Practice summary/check formulas must NOT reference a table name or defined name the learner has to create
(Excel may refuse to open a file with a formula pointing to a missing table). Use plain A1 ranges in summaries. Provide the
supply inventory as a PLAIN RANGE (as_table=False) so the learner converts it; in self-test mode a `customize` hook may add the table.
DATA: supply_inventory (~257 rows) + a 'Settings' sheet with ReportDate (12/31/2025) and ExpiringWindowDays (90).
PRACTICE (12–13): calculated column ExtendedValue = [@QtyOnHand]*[@UnitCost] (summary = SUM of that column range); total
inventory value; count of rows at/below reorder point; items expiring within the window of ReportDate; Total Row average; a
named-constant question; a structured-reference 'what does this return' question.
BONUS: reorder report: for rows at/below reorder point compute order cost (ReorderQty × UnitCost) by vendor; answer the vendor with
the largest reorder cost (text) and the amount.
"""),
    LessonInfo("3.2", "03-data-analysis", "02-data-validation-conditional-formatting", "Data Validation & Conditional Formatting", 120, [
        "Restrict entries with list, number, date, length, and custom-formula validation",
        "Build dependent drop-down lists",
        "Highlight what matters with conditional formatting rules, data bars, color scales, and icon sets",
        "Write formula-based rules that format entire rows",
    ], spec="""
GUIDE: Data Validation (Settings: any/whole/decimal/list/date/time/text length/custom; Input Message; Error Alert Stop/Warning/
Information; lists from ranges and named ranges; dependent dropdowns with INDIRECT; custom formulas like =AND(LEN(A2)=8,
ISNUMBER(--A2)); Circle Invalid Data; copying validation). Conditional formatting (Highlight Cells, Top/Bottom, Data Bars, Color
Scales, Icon Sets, duplicate values, dates occurring, formula rules with $-anchoring to color whole rows, rule order and
'Stop If True', Manage Rules, applies-to ranges, performance). Clinical uses: critical lab values, expiring supplies, overflow census.
DATA: ~400 lab results, supply inventory, ~60 days of unit census, and an 'Intake Log' sheet with some invalid entries
(bad MRN lengths, impossible dates, out-of-list payers) for Circle Invalid Data.
PRACTICE (12–13): most tasks ask the learner to create a rule AND answer a checkable count ("how many cells/rows will your rule
highlight?" — compute with Python; solution shows the rule and an equivalent COUNTIFS); validation tasks ask "how many entries are
circled as invalid?" Mark pure build steps check='manual' with clear steps in the solution.
BONUS: unit 'huddle board' — occupancy color scale, icon sets on day-over-day change, overflow rows highlighted, and a dependent
Facility → Department dropdown; checkable counts for each rule.
"""),
    LessonInfo("3.3", "03-data-analysis", "03-data-cleaning", "Cleaning Messy Data", 145, [
        "Spot common data problems: stray spaces, inconsistent case and categories, text dates, duplicates",
        "Fix them with TRIM, PROPER, SUBSTITUTE, VALUE, DATEVALUE, and lookup mapping tables",
        "Use Text to Columns, Flash Fill, Remove Duplicates, and Go To Special",
        "Document a repeatable cleaning process",
    ], spec="""
GUIDE: a cleaning workflow (profile → standardize → convert types → dedupe → validate → document), profiling tricks (filters,
COUNTIF, LEN vs LEN(TRIM)), fixing case/spaces, standardizing categories with a mapping table + XLOOKUP, numbers/dates stored as
text (VALUE, DATEVALUE, --, Text to Columns with column data format DMY/MDY), parsing with Text to Columns (delimited & fixed
width), Flash Fill, Remove Duplicates (which columns define a duplicate; normalize first!), highlighting duplicates with COUNTIF,
Go To Special → Blanks → fill down, Find & Replace with wildcards, keeping raw data untouched (work on a copy), a cleaning log.
Forward pointer: Power Query (4.3) automates this.
DATA: data/messy/patient_registrations_raw.csv (650 rows) on a 'Raw' sheet; a payer mapping table; tools/_truth/
patient_registrations_truth.csv gives each RecordID's true PatientID (use it to compute answers; never put it in the workbook).
PRACTICE (12–13): count rows with extra spaces in PatientName; standardized Sex column (summary = count of F); proper-case
'First Last' names; DOB text in 5 formats → real dates (the formula can be long; show a robust approach) — summary = count of
valid dates or the earliest DOB; phone → 10 digits; MRN → 8-digit text; Insurance variants → PayerID via mapping; exact duplicate
row count; number of distinct patients after cleaning MRNs (should equal the truth: 560).
BONUS: dedupe to one row per patient keeping the most recent RegisteredOn (parse both datetime formats) — how many rows remain
and the RegisteredOn kept for a stated MRN.
"""),
    LessonInfo("3.4", "03-data-analysis", "04-pivottables", "PivotTables", 130, [
        "Build PivotTables from a Table and arrange rows, columns, values, and filters",
        "Change summaries (Sum, Count, Average) and Show Values As (% of total, difference, running total)",
        "Group dates and numbers; filter with slicers and timelines",
        "Add calculated fields and use GETPIVOTDATA",
    ], spec="""
GUIDE: when to pivot, data prerequisites (tabular, one header row, no blanks, Table source), Insert → PivotTable, field list
areas, Value Field Settings (Sum/Count/Average/Max; Count vs Count of numbers; number format), Show Values As (% of Grand Total,
% of Column/Row Total, % Of, Difference From, Running Total In, Rank), grouping (dates → months/quarters/years; numbers → age bands;
manual groups), sorting & filtering (label, value, Top 10), Report Filter, slicers & timelines & report connections, calculated
fields (and their limits: they sum before dividing), refresh/change data source, layouts (Tabular, Repeat Item Labels, subtotals,
grand totals), drill-down (double-click), GETPIVOTDATA (and turning it off), PivotCharts (brief; charts in 3.5), Recommended PivotTables.
DATA: all 2025 encounters (~11k) with pre-joined FacilityName, DeptName, ServiceLine, PayerName, PayerType, DxCategory, AgeAtAdmit,
LOS, Readmit30, TotalCharges, plus a numeric ReadmitFlag (1/0) for averages.
PRACTICE (12–13): learners build pivots and type the numbers: counts by facility/type, average charges by payer type, month with
most observation stays (text), % of ED visits that are Medicaid, inpatient readmission rate by service line (average of ReadmitFlag),
age-band grouping counts, a calculated field result, a GETPIVOTDATA value. Every answer is computed in Python.
BONUS: readmission deep-dive pivot: rate by service line × payer type; which combination (with ≥30 index stays) has the highest
rate (text) and the rate; top 3 diagnoses by readmission count.
"""),
    LessonInfo("3.5", "03-data-analysis", "05-charts-visualization", "Charts & Data Visualization", 135, [
        "Choose the right chart for comparisons, trends, parts of a whole, distributions, and relationships",
        "Build and format column, line, bar, scatter, histogram, combo, and Pareto charts",
        "Add trendlines, dynamic titles, and sparklines",
        "Design clear, accessible charts that tell one story",
    ], spec="""
GUIDE: chart selection guide table, building from Tables/pivots, chart elements (+ button), Format pane, axes (scale, number format,
log scale), data labels, secondary axis & combo charts, histogram & box-and-whisker (2016+), Pareto, waterfall (budget variance),
scatter + trendline (equation, R²), sparklines, dynamic chart titles (="Title "&A1 in a cell, then =Sheet!$A$1 in title), chart
templates, accessibility (alt text, color-blind friendly palettes, direct labels), common mistakes (3-D, pie with many slices,
truncated axes for bars), PivotCharts.
DATA (pre-aggregated by the builder): monthly ED visits by facility (24 months), LWBS rate by month, readmission rate by service
line, payer mix, ED arrivals by hour of day, LOS vs age sample (300 inpatient stays), denial reasons counts.
PRACTICE (12–13): build-a-chart tasks paired with checkable numeric reads of the charted data (peak month — text; slope of a linear
trendline via SLOPE; R² via RSQ; busiest hour; cumulative % of the top two denial reasons). Pure build steps are check='manual'.
Use a `customize` hook to add REFERENCE CHARTS built with openpyxl.chart to the hidden Answer Key (or a hidden 'Chart Key' sheet).
BONUS: combo chart (monthly ED volume columns + LWBS % line on secondary axis) and a Pareto chart of denial reasons; checks on
the numbers behind them.
"""),
    LessonInfo("3.6", "03-data-analysis", "06-what-if-analysis", "What-If Analysis: Goal Seek, Scenarios, Data Tables & Solver", 125, [
        "Structure a model with separate inputs, calculations, and outputs",
        "Find break-even points with Goal Seek",
        "Compare cases with Scenario Manager and sensitivity Data Tables",
        "Optimize a staffing mix with Solver",
    ], spec="""
GUIDE: model layout best practices (inputs in blue font, named inputs, one formula per row, no hard-coded numbers in formulas),
Goal Seek (set cell / to value / by changing; limitations — one input, numeric precision), Scenario Manager (add scenarios, show,
summary report), one-variable and two-variable Data Tables (Data → What-If → Data Table; row/column input cells; {=TABLE()} array;
calculation option 'Automatic except data tables'), Solver add-in (enable it; objective, variable cells, constraints incl. int,
Simplex LP vs GRG Nonlinear, answer report), interpreting sensitivity.
DATA: build a Primary Care Clinic monthly P&L model on a 'Model' sheet from realistic inputs derived from the data (visits/day,
payer mix, average reimbursement by payer from claims for D400, provider/staff costs, supply cost per visit, fixed costs).
PRACTICE (12–13): base-case operating margin (formula); Goal Seek visits/day for break-even (compute exactly in Python — the model
should be linear so the answer is unique; tell learners to round to 1 decimal); required average reimbursement for a 5% margin;
values from a one-variable and a two-variable data table at stated inputs (compute in Python); scenario summary value.
BONUS: Solver staffing mix — choose integer RN/LPN/CNA counts minimizing cost subject to coverage hours, RN ≥ 60% of licensed staff,
and a CNA cap; solve by brute force in Python; ask for the minimum cost and the RN count.
"""),
    # ------------------------------------------------------------------ Level 4
    LessonInfo("4.1", "04-advanced-analysis", "01-dynamic-arrays", "Dynamic Arrays: FILTER, SORT, UNIQUE & More", 125, [
        "Understand spilling, the # spill reference, and #SPILL! errors",
        "Extract lists with UNIQUE, FILTER, SORT, and SORTBY",
        "Generate sequences with SEQUENCE and reshape with TAKE, DROP, CHOOSECOLS, VSTACK, and HSTACK",
        "Combine functions into one-formula reports",
    ], spec="""
GUIDE: what spilling is (vs legacy Ctrl+Shift+Enter arrays), spill range operator (A2#), #SPILL! causes, implicit intersection @,
UNIQUE (by_col, exactly_once), SORT/SORTBY (multi-key), FILTER (AND with *, OR with +, if_empty), SEQUENCE (calendars, numbering),
TAKE/DROP, CHOOSECOLS/CHOOSEROWS, VSTACK/HSTACK, TOCOL, XLOOKUP returning arrays, nesting e.g. SORT(UNIQUE(FILTER())), COUNTA/ROWS
of a spill, referencing spills in charts & validation lists, compatibility (365/2021; older versions show _xlfn.).
TECH NOTES: answer cells on the Practice sheet must hold SCALAR formulas (a spill there would collide with the Check column). For
spill exercises use a 'Workspace' sheet with labeled anchor cells; checks use plain formulas over a generous range (e.g.
COUNTA(Workspace!B5:B200), INDEX positions). The library writes formulas containing dynamic-array functions as single-cell array
formulas automatically so they evaluate correctly; LibreOffice verification supports FILTER/SORT/UNIQUE/XLOOKUP/LET/SEQUENCE/
VSTACK/TAKE but NOT LAMBDA/MAP/BYROW — set live=False for those and verify with Python.
DATA: 2,000 2025 encounters (pre-joined names), providers, departments.
PRACTICE (12–13): count of distinct attending providers in the ED via ROWS(UNIQUE(FILTER())); top-5 charges via TAKE(SORT()) — enter
the 5th value; FILTER count & sum with two criteria; FILTER with OR; SORTBY first item; SEQUENCE-based month starts (enter the 7th);
CHOOSECOLS/VSTACK questions; spill tasks on Workspace with count checks.
BONUS: a one-formula 'department leaderboard': departments with ≥100 encounters sorted by average charge descending (show name,
count, average) — checks: number of rows, first department, its average.
"""),
    LessonInfo("4.2", "04-advanced-analysis", "02-advanced-formulas-let-lambda", "Advanced Formulas: LET, LAMBDA & Array Logic", 150, [
        "Write multi-condition array logic with SUMPRODUCT and boolean math",
        "Make complex formulas readable and fast with LET",
        "Create reusable custom functions with LAMBDA and the Name Manager",
        "Use MAP, REDUCE, SCAN, BYROW, and audit formulas with Evaluate Formula",
    ], spec="""
GUIDE: boolean arrays (TRUE*TRUE=1, -- coercion), SUMPRODUCT multi-criteria sums/counts and when it beats SUMIFS (OR logic,
computed criteria like MONTH(), weighted averages), LET (naming intermediate results; performance), LAMBDA (parameters, testing in a
cell, saving in Name Manager e.g. LOSDAYS, BMI, AGEAT; recursion note), helper functions MAP/REDUCE/SCAN/BYROW/BYCOL/MAKEARRAY with
examples, INDIRECT/OFFSET (volatile — prefer INDEX), CHOOSE, multi-criteria lookups (XLOOKUP(1,(A=x)*(B=y),…)), FREQUENCY, formula
auditing (Evaluate Formula, Trace Precedents/Dependents, Watch Window, F9 on a selection). Version notes (LAMBDA family = 365).
TECH NOTES: LibreOffice cannot verify LAMBDA/MAP/REDUCE/SCAN/BYROW — set live=False and self_test=False for those tasks and make
sure the Python answer is computed independently. Do NOT define LAMBDA names in the workbook (openpyxl defined names with LAMBDA are
fragile); the learner creates them.
DATA: all ~5,900 inpatient encounters (PatientID, Admit/Discharge datetimes, DischargeDisposition, Facility, Dx, Readmit30 omitted!)
for computing 30-day readmissions from scratch; ~1,000 lab results.
PRACTICE (12–13): readmission flag column via COUNTIFS on same patient with AdmitDate in (Discharge, Discharge+30] (summary = count of
1s, must match the dataset's Readmit30 = Y count — define the rule exactly as data/README.md does), SUMPRODUCT revenue with OR logic,
LET-based average LOS for a facility/dx, a LAMBDA BMI or LOS function (answer = its result for given inputs), MAP/BYROW results,
multi-criteria XLOOKUP, FREQUENCY of LOS bins (enter one bin count).
BONUS: readmission rate by facility with CMS-like exclusions (exclude index stays ending in death, transfer, or AMA) — one LET
formula per facility; answer the highest facility and its rate.
"""),
    LessonInfo("4.3", "04-advanced-analysis", "03-power-query", "Power Query: Import, Transform & Combine", 170, [
        "Import CSV files and whole folders with Get & Transform",
        "Clean and reshape data with Applied Steps: types, splits, filters, unpivot, group by",
        "Merge (join) and append queries",
        "Build refreshable, repeatable data pipelines and read the M code behind them",
    ], spec="""
GUIDE: what Power Query is (ETL; non-destructive; refreshable), Data → Get Data (From Text/CSV, From Folder, From Table/Range),
the Power Query Editor tour (ribbon, Applied Steps, formula bar, Advanced Editor, query settings, data type icons & locale),
common transforms (remove/choose columns, rename, change type, filter, replace values, trim/clean/format, split column by delimiter,
extract, merge columns, conditional & custom columns), unpivot/pivot, group by (multiple aggregations), merge queries (join kinds
incl. left anti), append, combine files from a folder (sample file, helper queries), parameters (file path), load options (Table,
PivotTable, Connection only, Data Model), refresh (and Refresh All), query dependencies, M basics (let/in, steps, Table.* functions),
version notes (Windows 365 full; Mac supports most transforms; Excel for web limited).
FILES: the lesson folder must include a `data/` subfolder created by the builder (write via lesson.extra_files or a customize hook):
`claims_monthly/claims_2025_01.csv` … `_12.csv` (2025 claims split by SubmitDate month), `budget_2025_wide.csv` (months as columns),
`payers.csv`, `encounters_2025.csv` (subset of columns), and `new_month/claims_2026_01.csv` for the bonus refresh (create a plausible
January 2026 file by re-dating a sample of December 2025 claims; state clearly it is synthetic).
PRACTICE (12–13): learners run the transformation in Power Query and type resulting numbers: combined row count and total PaidAmount;
unpivoted budget row count and Q1 actual total; denied claims grouped by reason (count for one reason); merged payer type totals;
left-anti count of encounters without claims; custom column results. Solutions are M code (solution_lang='m') + steps; check numbers
are computed in Python.
BONUS: refreshable denial dashboard query: folder combine + merge payers + group by PayerType & DenialReason; then drop the January
2026 file into the folder and refresh — new totals.
"""),
    LessonInfo("4.4", "04-advanced-analysis", "04-power-pivot-dax", "Data Model, Power Pivot & DAX", 175, [
        "Design a star schema and load tables into the Data Model",
        "Create relationships and a proper date table",
        "Write DAX measures with SUM, COUNTROWS, DISTINCTCOUNT, DIVIDE, and CALCULATE",
        "Use time intelligence (TOTALYTD, SAMEPERIODLASTYEAR) and iterators (SUMX, AVERAGEX)",
    ], spec="""
GUIDE: why a Data Model (millions of rows, multiple tables without VLOOKUP, measures), star schema (fact vs dimension), adding tables
(Power Pivot → Add to Data Model, or Power Query load), Diagram View & relationships (1:*, single direction, active/inactive),
date table (CALENDAR/CALENDARAUTO or provided DimDate; Mark as Date Table), calculated columns vs measures (row vs filter context),
DAX essentials: SUM, AVERAGE, COUNTROWS, DISTINCTCOUNT, DIVIDE, CALCULATE with filters, ALL/REMOVEFILTERS, FILTER, RELATED,
SUMX/AVERAGEX, time intelligence TOTALYTD, SAMEPERIODLASTYEAR, DATEADD; implicit vs explicit measures; KPIs; PivotTables from the
Data Model; CUBEVALUE/CUBEMEMBER (brief); version notes (Power Pivot = Windows; Mac can't edit the model).
DATA: separate sheets/tables: FactEncounters (all 2024–2025 encounters, keys + measures incl. LOS days, charges, Readmit flag),
FactClaims, DimPatient (subset of columns + AgeGroup), DimProvider, DimDepartment, DimFacility, DimPayer, DimDiagnosis, DimDate.
PRACTICE (12–13): learners create measures and type the value the pivot shows for a stated filter context, e.g. Total Charges 2025;
Encounter count for Cedar Ridge ED in Q3 2025; Avg LOS inpatient by facility; Distinct patients 2025; Denial Rate for Commercial;
Readmission Rate; YoY % change in ED visits (2025 vs 2024); YTD charges at end of June 2025. Solutions are DAX (solution_lang='dax').
Compute every value in Python.
BONUS: rolling 3-month average ED visits for Dec 2025; % change vs same month last year for one facility; RANKX of attending
providers by inpatient encounters — the #1 provider's name.
"""),
    LessonInfo("4.5", "04-advanced-analysis", "05-statistics-forecasting", "Statistics & Forecasting", 150, [
        "Summarize distributions with the Analysis ToolPak and histograms",
        "Measure relationships with correlation and linear regression",
        "Test differences with t-tests and quantify uncertainty with confidence intervals",
        "Forecast volumes and monitor processes with control charts",
    ], spec="""
GUIDE: enabling the Analysis ToolPak, Descriptive Statistics output, histograms (FREQUENCY, bins), CORREL/PEARSON, scatter +
trendline, SLOPE/INTERCEPT/RSQ/LINEST (and ToolPak Regression output: coefficients, R², p-values, residuals), T.TEST (tails, types)
and the ToolPak t-test, CONFIDENCE.T / CI = mean ± t·s/√n, moving averages, FORECAST.LINEAR/TREND, FORECAST.ETS and Forecast Sheet
(seasonality; note implementations differ so results are approximate), statistical process control (I-chart: mean ± 3σ (use
moving-range σ = MR̄/1.128 or STDEV — say which), p-chart for rates with varying n), correlation ≠ causation.
TECH NOTE: FORECAST.ETS results are implementation-specific → make ETS tasks manual or use a wide tolerance; everything else exact.
DATA: daily ED arrivals per facility (731 days, aggregated from ed_visits), monthly inpatient volumes & readmission counts,
~400 inpatient stays (age, LOS, charges), door-to-provider minutes by facility.
PRACTICE (12–13): CORREL(age, LOS); SLOPE/INTERCEPT/RSQ; predicted LOS for age 80; T.TEST p-value for LOS F01 vs F03 (two-tailed,
unequal variance); 95% CI half-width for mean door-to-provider; FORECAST.LINEAR next month; 7-day moving average on a stated date;
I-chart UCL/LCL; count of days above UCL; a histogram bin count.
BONUS: p-chart of monthly readmission rates (varying n) — months above the UCL (count) and the UCL for a stated month.
"""),
    LessonInfo("4.6", "04-advanced-analysis", "06-dashboards", "Building Interactive Dashboards", 160, [
        "Plan a dashboard around audience, questions, and KPIs",
        "Build KPI cards and selector-driven formulas",
        "Connect slicers to multiple PivotTables and drive charts from selections",
        "Polish layout, interactivity, and performance",
    ], spec="""
GUIDE: dashboard design (audience, 5-second rule, KPI selection, layout grid, alignment, white space, color with meaning,
avoid chartjunk), architecture (Data → Calc/Model → Dashboard sheets), KPI cards (big number, target, variance, trend arrow via
conditional formatting or UNICHAR(9650/9660)), selectors (data-validation dropdowns; Form Controls combo box with cell link; slicers &
timelines with Report Connections to multiple pivots), selector-driven formulas (SUMIFS/AVERAGEIFS/XLOOKUP/FILTER on the model),
dynamic charts (chart series pointing at formula ranges), sparklines, conditional formatting, camera tool/linked picture, hiding
gridlines/headings, protecting the sheet, printing to one page, performance (avoid volatile functions), accessibility.
DATA: a 'KPI_Monthly' table built by the builder: month × facility rows with ED visits, LWBS %, median door-to-provider, inpatient
admits, ALOS, readmission rate, occupancy %, HCAHPS top-box % (OverallRating 9–10), denial rate; plus targets table.
PRACTICE (12–13): with stated selector values (e.g. Facility = Cedar Ridge Medical Center, Month = 2025-10) compute KPI values
with formulas (checkable); variance to target; YoY change; trend direction text ('▲'/'▼' or 'Up'/'Down'); system-wide totals.
BONUS: build the full dashboard to spec; customize hook adds a REFERENCE dashboard sheet (KPI cells + openpyxl charts) inside the
hidden key or a hidden 'Reference Dashboard' sheet; checkable questions about selected KPI values.
"""),
    # ------------------------------------------------------------------ Level 5
    LessonInfo("5.1", "05-automation-vba", "01-recording-macros", "Recording Your First Macros", 110, [
        "Show the Developer tab and set macro security safely",
        "Record, run, and save macros in a macro-enabled workbook (.xlsm)",
        "Choose between absolute and relative recording",
        "Run macros from buttons, shortcuts, and the Quick Access Toolbar — and read the recorded code",
    ], spec="""
GUIDE: what a macro is, Developer tab (Windows & Mac), Trust Center / macro security levels, Mark of the Web & blocked macros in
downloaded files (how to Unblock safely), file formats (.xlsx can't store macros; .xlsm, .xlsb, .xltm, .xlam), Personal Macro
Workbook (PERSONAL.XLSB), Record Macro dialog (naming rules, shortcut keys & conflicts, store in, description), absolute vs
'Use Relative References', stopping, running (Alt+F8, shortcut, button/shape → Assign Macro, QAT), opening the VBE (Alt+F11) to
read recorded code (Sub/End Sub, Range().Select, Selection, With blocks) and what the recorder over-records, simple edits, what the
recorder can't do (loops, decisions) → leads to 5.2; Excel for the web/iPad can't run VBA (Office Scripts alternative, brief).
FILES: workbook is .xlsx (the learner saves as .xlsm). Provide `solutions/` folder with .bas files of reference recorded macros
(clean, commented) — and a README note that solutions are spoilers. Keep code realistic (what the recorder emits, then a cleaned version).
DATA: an unformatted daily census export for one unit (one month) to turn into a report via a recorded macro; ED visits for bonus.
PRACTICE (12–13): knowledge checks with text answers (file extension, relative vs absolute, where Personal Macro Workbook lives,
shortcut to open VBE) + 'after running your macro' checks read via summary formulas from cells the macro writes (e.g. the macro adds
a TOTAL row → summary reads it; use INDIRECT if referencing a sheet the macro creates). Self-test uses fill values.
BONUS: record → then edit a macro that filters ED visits to ESI 1–2, copies to a new sheet 'HighAcuity', sorts by arrival; checks
via INDIRECT("'HighAcuity'!A1") style summaries (count of rows, first EDVisitID).
"""),
    LessonInfo("5.2", "05-automation-vba", "02-vba-fundamentals", "VBA Fundamentals", 160, [
        "Navigate the Visual Basic Editor and organize code in modules",
        "Declare variables with the right data types and Option Explicit",
        "Control flow with If, Select Case, For, For Each, and Do loops",
        "Debug with breakpoints, stepping, the Immediate window, and Debug.Print",
    ], spec="""
GUIDE: VBE tour (Project Explorer, Properties, Code window, Immediate/Locals/Watch windows), modules vs sheet/ThisWorkbook modules,
Sub procedures & calling them, Option Explicit (and Tools → Options → Require Variable Declaration), data types (Long vs Integer,
Double, Currency, String, Date, Boolean, Variant, Object), Dim/Const/scope (procedure, module, Public), operators (\\ integer division,
Mod, &), MsgBox/InputBox (+ Application.InputBox Type), If/ElseIf/Else, Select Case (ranges, Is >=), For…Next (Step), For Each over
a Range, Do While/Until, Exit For/Do, string & date functions in VBA (Left, InStr, Format, DateDiff), reading/writing cells
(Range("A1").Value, Cells(r,c)), debugging (F8, F9, Ctrl+Shift+F8, Locals, Watches, Debug.Print, Stop), common compile/runtime errors.
FILES: `starter/` .bas modules with TODO stubs and `solutions/` .bas with full answers (spoilers). Practice workbook .xlsx with an
'Output' sheet where macros write results (Practice checks read Output cells via summary formulas; self-test uses fill values).
DATA: ~500 lab results.
PRACTICE (12–13): mix of predict-the-output questions (text/number answers about given snippets: loops, \\ and Mod, Select Case)
and macro-output tasks (CountCriticals writes count to Output!B2; average potassium; first critical LabResultID; loop flags written
to a column → summary count).
BONUS: a macro that color-codes rows by AbnormalFlag and writes a per-test summary table (count, # abnormal, % abnormal) to Output;
checks on specific summary cells.
"""),
    LessonInfo("5.3", "05-automation-vba", "03-vba-ranges-worksheets", "VBA: Ranges, Worksheets & Workbooks", 180, [
        "Navigate the object model: Application, Workbooks, Worksheets, Range",
        "Find the last row and work with CurrentRegion, Offset, and Resize",
        "Loop through sheets; add, copy, rename, and delete them",
        "Process data fast with arrays, AutoFilter, and Sort — without Select",
    ], spec="""
GUIDE: object model & dot notation, ThisWorkbook vs ActiveWorkbook, Worksheets("Name") vs CodeName, Range/Cells/Rows/Columns,
Offset/Resize/CurrentRegion/UsedRange, last row (Cells(Rows.Count,"A").End(xlUp).Row) and pitfalls, Value vs Value2 vs Text vs Formula,
With…End With, For Each ws In ThisWorkbook.Worksheets, Worksheets.Add (After:=), .Name, .Copy, .Delete with DisplayAlerts, error if
sheet exists (function SheetExists), reading a range into a Variant array and writing back (speed), AutoFilter & SpecialCells
(xlCellTypeVisible), Range.Sort, ListObjects basics, Application.ScreenUpdating/Calculation/EnableEvents with safe restore,
avoiding Select/Activate, opening/closing other workbooks (Workbooks.Open, SaveAs FileFormat:=xlOpenXMLWorkbook).
FILES: starter/ and solutions/ .bas (spoilers). Checks must read sheets the macros create through INDIRECT (a missing sheet gives
#REF!, which the check treats as 'Not yet') — never a direct reference to a sheet that doesn't exist yet.
DATA: 2,000 2025 encounters.
PRACTICE (12–13): split encounters into one sheet per facility (checks: row counts via INDIRECT); summary sheet of counts and total
charges by EncounterType written by macro (checks); last-row function results; array-based LOS calculation written to a column
(summary = SUM); delete-and-recreate safety question; predict-the-output questions about Offset/Resize addresses.
BONUS: 'monthly packet' macro: for each month of 2025 create a sheet with that month's inpatient encounters sorted by charges desc and
a header total; checks on two months' row counts and top charge.
"""),
    LessonInfo("5.4", "05-automation-vba", "04-vba-functions-error-handling", "VBA: Custom Functions, Dictionaries & Error Handling", 185, [
        "Write user-defined functions (UDFs) you can call from cells",
        "Pass arguments ByVal/ByRef, use Optional arguments, and return errors with CVErr",
        "Count and group with Collections and Scripting.Dictionary",
        "Handle errors gracefully with On Error, the Err object, and cleanup code",
    ], spec="""
GUIDE: Function vs Sub, UDF rules (no side effects on other cells, Application.Volatile, recalculation), arguments (ByVal/ByRef,
Optional with IsMissing/defaults, ParamArray brief), returning errors (CVErr(xlErrValue)), WorksheetFunction vs Application.Match
(error behavior), input validation, On Error GoTo Handler / Resume Next / Resume / On Error GoTo 0, Err.Number/Description/Raise,
cleanup pattern (restore settings in the exit block), Collections vs Scripting.Dictionary (late binding CreateObject, .Exists, .Item,
.Keys, counting & summing patterns), sorting keys (simple bubble sort or worksheet sort), organizing code (modules, naming, comments),
saving UDFs in an add-in (.xlam).
UDF ideas: BMI(weightLb, heightIn), AGEAT(dob, asOf), LOSDAYS(admit, discharge) returning #VALUE! when discharge < admit,
ESINAME(level), DENIALRATE(range).
FILES: starter/ and solutions/ .bas (spoilers). Practice uses learner-filled UDF columns with summary formulas; self-test fills values.
DATA: ~500 patients (height, weight, DOB), ~500 encounters, claims for dictionary counts.
PRACTICE (12–13): UDF results for stated inputs (number answers), column sums produced by UDFs, Dictionary macro output cells (count of
distinct payers, top payer by claims), error-handling predict-the-output questions.
BONUS: a robust 'payer summary' macro: Dictionary of payer → (count, billed, paid, denied) written to a sorted summary table with full
error handling (missing sheet → friendly message); checks on summary cells.
"""),
    LessonInfo("5.5", "05-automation-vba", "05-events-userforms-automation", "Events, UserForms & Automated Reports", 195, [
        "Respond to workbook and worksheet events (Open, Change, BeforeSave)",
        "Build a validated data-entry UserForm that writes to a Table",
        "Automate a report: combine files with Dir, export to PDF, save timestamped copies",
        "Know the safe limits of automation and when to use Office Scripts or Power Automate",
    ], spec="""
GUIDE: event procedures (where they live: ThisWorkbook & sheet modules), Workbook_Open, Workbook_BeforeSave, Worksheet_Change with
Intersect and Target, Application.EnableEvents (and resetting after errors), Worksheet_SelectionChange (brief); UserForms (insert,
controls: Label, TextBox, ComboBox (RowSource vs AddItem), OptionButton, CheckBox, CommandButton; properties; Initialize event;
validation; writing a new ListRow to a ListObject; Unload Me); file automation (Dir loop over *.csv, Workbooks.Open, copying data,
closing), ExportAsFixedFormat to PDF, SaveCopyAs with Format(Now,"yyyy-mm-dd_hhmm"), Application.OnTime (brief), security &
reliability (signing macros, error logging), Office Scripts (TypeScript) & Power Automate as modern alternatives (brief, with a small
Office Script example).
FILES: starter/ and solutions/ (.bas for modules, .cls for sheet/workbook event code, .frm notes — since .frm needs a binary .frx,
give UserForm build steps + code-behind as text), and a `data/census_monthly/` folder with 12 monthly census CSVs (2025) for the
Dir-combine task. Checks read cells the macros write (summary formulas; INDIRECT for created sheets).
DATA: an 'Intake' sheet with a ListObject table for the form; monthly census CSVs.
PRACTICE (12–13): combine 12 census files → total rows and total patient days (checkable); intake form adds a row → checks on table
row count after entering 3 specified patients; Worksheet_Change timestamp (manual); knowledge questions (EnableEvents, Intersect).
BONUS: one-click monthly report macro — combine, summarize by unit (ADC, occupancy), write a summary sheet, export PDF; checks on
summary values.
"""),
    # ------------------------------------------------------------------ Level 6
    LessonInfo("6.1", "06-capstone", "01-hospital-performance-review", "Capstone: Hospital Performance Review", 240, [
        "Plan an analysis from business questions to deliverables",
        "Prepare multi-table data (cleaning, joins, calculated fields)",
        "Analyze throughput, quality, utilization, and finance KPIs",
        "Deliver a dashboard, an automated refresh macro, and an executive summary",
    ], spec="""
GUIDE: a project brief from the (fictional) Bluestone CMO/CFO asking for a 2025 performance review; a suggested work plan (data prep
→ metrics → analysis → dashboard → automation → summary); metric definitions table (ALOS, O/E LOS index using diagnoses.ExpectedLOS,
30-day readmission rate per data/README, ED median door-to-provider, LWBS %, occupancy, HCAHPS top-box %, denial rate, net
collection rate = Paid / Allowed for adjudicated claims, cost per encounter optional), tips that reference earlier lessons by number,
and a rubric/checklist for the deliverables. Encourage any toolset (formulas, pivots, Power Query, DAX, VBA).
DATA: 2025 slices of encounters, ed_visits, claims, patient_satisfaction, daily_census, diagnoses, departments, facilities, payers.
PRACTICE (12–13): the key numbers the executive summary needs, each auto-checked (state definitions precisely so answers are
unambiguous), e.g. system ALOS (inpatient, decimal days), O/E LOS index by facility (enter F03), readmission rate by facility, the
service line with the highest readmission rate (text), ED median door-to-provider by facility, LWBS %, average occupancy for ICUs,
HCAHPS top-box %, denial rate overall and the top denial reason (text), net collection rate.
BONUS: an 'opportunity sizing' question: if the facility with the highest O/E LOS index reduced its inpatient LOS to O/E = 1.00,
how many bed-days would be saved in 2025 (sum over its inpatient stays of max(0, actual − expected) — define precisely) and the
equivalent number of staffed beds (÷ 365).
"""),
]

for _l in LESSONS:
    _l.level = MODULES[_l.module][1]


def by_code(code: str) -> LessonInfo:
    for l in LESSONS:
        if l.code == code:
            return l
    raise KeyError(code)


def neighbors(code: str):
    i = [l.code for l in LESSONS].index(code)
    prev = LESSONS[i - 1] if i > 0 else None
    nxt = LESSONS[i + 1] if i + 1 < len(LESSONS) else None
    return prev, nxt
