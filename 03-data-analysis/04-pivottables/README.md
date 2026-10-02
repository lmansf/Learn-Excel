# Lesson 3.4 · PivotTables

> **Level:** Intermediate · **Time:** about 60 minutes · **Workbook:** [`3.4-pivottables.xlsx`](3.4-pivottables.xlsx)
> **Data:** All 11,145 encounters that began in 2025 at Bluestone's four facilities (one row per encounter), with facility, department, service line, payer, and diagnosis names already joined in, plus age, length of stay, charges, and a 30-day readmission flag.

Every week someone at Bluestone is asked a version of the same question: *how many, how much, and broken down by what?*
How many ED visits did each hospital see? What do Self-Pay visits charge on average? Which service line has the worst
readmission rate, and was December busier than November? You could answer each one with a COUNTIFS or AVERAGEIFS formula
(Lesson 2.5), but every new breakdown means a new grid of labels and formulas. A **PivotTable** builds that grid for you. You
drag a few fields into place, Excel summarizes 11,145 encounters in a second, and you can rearrange the summary just as
fast. In this lesson you build the pivots that finance, quality, and operations teams rely on: volumes by facility, payer
mix, monthly trends, age bands, readmission rates, and charges per patient day.

## What you'll learn

- Build PivotTables from a Table and arrange rows, columns, values, and filters
- Change summaries (Sum, Count, Average) and Show Values As (% of total, difference, running total)
- Group dates and numbers; filter with slicers and timelines
- Add calculated fields and use GETPIVOTDATA

## 📖 Guide

### 1. What a PivotTable does

A **PivotTable** is an interactive summary of a list of records. This lesson often calls it a **pivot** for short. You
choose which columns become row labels, column labels, and filters, and which column is summed, counted, or averaged. Excel
does the grouping and the arithmetic. You don't write any formulas, and you can rearrange (*pivot*) the layout by dragging.

A few words come up constantly:

| Term | Meaning | Example in this lesson |
|---|---|---|
| **Source data** | The list the pivot summarizes | tblEncounters, 11,145 rows |
| **Field** | A column of the source data. The pivot uses the column header as the field name | FacilityName, TotalCharges |
| **Item** | One distinct value of a field | *Inpatient* is an item of EncounterType |
| **Area** | One of the four boxes where you place fields: Rows, Columns, Values, Filters | |
| **Value field** | A field in the Values area, together with its summary | Sum of TotalCharges |

You already know another way to build summaries, so it helps to compare them:

| | COUNTIFS / SUMIFS grid (Lesson 2.5) | PivotTable |
|---|---|---|
| Build a facility × encounter-type summary | Type the labels, write one formula with mixed references, fill it | Drag two fields |
| Change the breakdown (say, payer instead of facility) | Retype labels and edit formulas | Drag a different field |
| New categories appear in the data | You add labels by hand | They appear after a refresh |
| Data changes | Results update immediately | Results update when you click **Refresh** |
| Control over the layout | Complete | Limited to the pivot layouts |
| Best for | Fixed reports, dashboards, one number inside a model | Exploring, one-off questions, recurring summary reports |

Use both. A pivot is the fastest way to explore and find an answer. A formula is better when one number must sit in a fixed
cell of a report. The answer key for every task in this lesson shows the matching COUNTIFS, SUMIFS, or AVERAGEIFS formula, so
you can see the two approaches side by side.

### 2. Get the data ready

A pivot is only as good as its source. Excel expects a **tabular list**: one header row, one record per row, and one kind of
information per column. These problems break pivots or quietly distort them:

| Problem in the source | What goes wrong in the pivot | Fix |
|---|---|---|
| A blank header cell | Excel refuses to build the pivot: *The PivotTable field name is not valid* | Give every column a header |
| Two columns with the same header | Excel renames the second one (Charges2), and it's easy to drag the wrong one | Make headers unique |
| Blank rows inside the data | The range Excel guesses stops at the first blank row | Delete them, or use a Table |
| Subtotal or total rows mixed into the data | Those amounts are counted twice | Delete them. The pivot makes its own totals |
| Numbers or dates stored as text | The field defaults to Count, and dates can't be grouped | Convert them (Lesson 3.3) |
| One column per month (Jan, Feb, Mar…) | You can't put "month" in Rows or filter by it | Reshape to one Month column (Power Query, Lesson 4.3) |
| Merged cells | Broken headers and blank items | Unmerge them |

Build pivots from an **Excel Table** (Lesson 3.1) whenever you can. A pivot built on tblEncounters keeps pointing at the
whole Table as it grows, so new rows appear the next time you refresh. A pivot built on a fixed range such as
`Encounters!$A$1:$R$11146` ignores anything added below row 11,146 until you change its source.

The Encounters sheet in this lesson's workbook is already a Table named **tblEncounters**: one row per encounter, 18
columns, and no blank headers. The **Data Dictionary** sheet explains every column. Four of them deserve a note now:

- **ReadmitFlag** is 1 when an inpatient stay was followed by another inpatient admission within 30 days, 0 when it wasn't,
  and blank on every row that isn't an inpatient stay. Section 5 shows why this one column gives you three useful numbers.
- **LOSDays** is DischargeDate − AdmitDate in days, so a same-day visit has 0.
- **DeptName** repeats across facilities. Every hospital has an *Emergency Department*, so pair DeptName with FacilityName
  when the facility matters.
- **PatientID** identifies the person. A patient with three visits has three rows, so counting PatientID counts encounters,
  not patients.

### 3. Create a PivotTable

1. Click any cell inside tblEncounters on the Encounters sheet.
2. Choose **Insert → PivotTable**. In Microsoft 365 for Windows, if the button opens a short menu, choose
   **From Table/Range**. You can also choose **Table Design → Summarize with PivotTable** (Mac: **Table → Summarize with
   PivotTable**).
3. In the dialog (called *PivotTable from table or range* or *Create PivotTable*, depending on your version), check that
   **Table/Range** shows `tblEncounters`.
4. Under where to place it, choose **New Worksheet**.
5. Leave **Add this data to the Data Model** unticked (the Data Model is Lesson 4.4), and click **OK**.

Excel inserts a new sheet with an empty pivot starting at A3 and opens the **PivotTable Fields** pane on the right. Rows 1
and 2 stay empty to make room for filters.

Whenever a cell inside a pivot is selected, two extra ribbon tabs appear. This lesson uses the current names:

| Excel version | Ribbon tabs for pivots |
|---|---|
| Microsoft 365 and Excel 2021 or later (Windows and Mac) | **PivotTable Analyze** and **Design** |
| Excel 2013 to 2019 (Windows) | **PivotTable Tools → Analyze** and **Design** |
| Excel 2010 (Windows) | **PivotTable Tools → Options** and **Design** |

> ⚠️ If the dialog shows a range like `Encounters!$C$5:$C$40` instead of the Table name, the active cell wasn't inside the
> Table when you started. Click **Cancel**, click inside the Table, and start again.

> 💡 **Tip:** **Insert → Recommended PivotTables** previews several pivots that Excel thinks suit your data. Pick one as a
> starting point, then rearrange it like any other pivot.

> 📋 Excel for the web can create and edit PivotTables, but some tools in this lesson, such as calculated fields, are missing
> or limited there. Use desktop Excel for Windows or Mac for the practice.

### 4. Build the layout in the PivotTable Fields pane

The top of the **PivotTable Fields** pane lists every field in the source, one checkbox per column of tblEncounters. The
bottom holds the four areas:

| Area | What it does | Good fields to put here |
|---|---|---|
| **Rows** | One row per item, down the left side | FacilityName, ServiceLine, months |
| **Columns** | One column per item, across the top | EncounterType, PayerType (fields with few items) |
| **Values** | The numbers summarized where each row meets each column | EncounterID (to count), TotalCharges, ReadmitFlag |
| **Filters** | A filter above the pivot that limits which records count, without adding rows or columns | EncounterType, FacilityName |

To place a field, drag it from the list into an area. Ticking a field's checkbox also works: Excel sends text and date fields
to Rows and number fields to Values. To move a field, drag it to another area. To remove it, untick it or drag it out of the
pane.

**Worked example: encounters by type.** Drag **EncounterType** to Rows, then drag **EncounterID** to Values. Excel names the
value field *Count of EncounterID*:

| Row Labels | Count of EncounterID |
|---|--:|
| Emergency | 3,774 |
| Inpatient | 2,859 |
| Observation | 593 |
| Outpatient | 3,919 |
| **Grand Total** | **11,145** |

The Grand Total equals the number of rows in tblEncounters, so every encounter is counted exactly once. Any field without
blanks counts every row, which is why EncounterID is a safe choice. Counting ReadmitFlag instead would give only 2,859,
because Count skips blank cells and ReadmitFlag is blank outside inpatient stays.

Now try the other areas on the same pivot:

- **Columns:** drag **FacilityName** to Columns. The pivot becomes a grid with one column per facility and a Grand Total
  column on the right.
- **Filters:** drag FacilityName from Columns to Filters instead. A filter cell appears in B1 showing *(All)*. Open its
  button, pick a facility, and click **OK**. The whole pivot now describes that facility only. To pick several, tick
  **Select Multiple Items** first.
- **Two fields in one area:** put ServiceLine and then PayerType in Rows. The field on top is the outer level, so the payer
  types are listed inside each service line. Drag them to swap the order.

> ⚠️ A pivot is a report, not a range you can edit. If you type over a value or try to insert a row inside it, Excel says
> it can't change that part of the PivotTable report. Change the source data and refresh instead.

> 📋 The field list disappears when you click outside the pivot. Click inside the pivot to bring it back. If you closed it,
> choose **PivotTable Analyze → Field List**.

> 💡 **Tip:** On Windows, ticking **Defer Layout Update** at the bottom of the pane lets you make several changes and then
> apply them all at once with **Update**. It saves waiting on large sources.

### 5. Choose how values are summarized

Excel picks a starting summary for each value field: **Sum** when every cell in the column is a number, and **Count** when the
column contains any text or blank cells. Always read the label. *Count of TotalCharges* where you expected *Sum of
TotalCharges* is the first sign that a numeric column contains text or blanks.

To change the summary, open **Value Field Settings** in any of these ways:

- Right-click a number in the pivot and choose **Summarize Values By** for the common functions, or **Value Field
  Settings…** for all of them.
- In the Values area of the pane, click the field and choose **Value Field Settings…** (Mac: click the field's **ⓘ**
  button).
- On Windows, double-click the value field's header cell in the pivot, such as *Sum of TotalCharges*.

| Summarize by | Returns | Bluestone example |
|---|---|---|
| **Sum** | Total | Total charges by facility |
| **Count** | Number of non-empty cells, like COUNTA | Encounters by type (Count of EncounterID) |
| **Average** | Mean | Average length of stay by facility |
| **Max** / **Min** | Largest / smallest value | The longest stay in each service line |
| **Product** | All values multiplied | Rarely useful |
| **Count Numbers** | Number of numeric cells, like COUNT | Rows where a number was recorded |
| **StdDev**, **StdDevp**, **Var**, **Varp** | Spread of the values (sample or population) | Variation in charges (Lesson 4.5) |

**Worked example: average length of stay.** Put EncounterType in Filters and select **Inpatient**, put FacilityName in Rows,
and put LOSDays in Values. Then change *Sum of LOSDays* to **Average**:

| Row Labels | Average of LOSDays |
|---|--:|
| Ashby Falls Community Hospital | 4.53 |
| Bluestone Memorial Hospital | 4.68 |
| Cedar Ridge Medical Center | 4.48 |
| **Grand Total** | **4.63** |

The Grand Total is the average of all 2,859 stays, not the average of the three facility averages (4.56). Bluestone
Memorial has most of the stays, so it pulls the total toward its own figure. That's correct, and it's the same reason a
period rate should be total ÷ total (Lesson 1.4).

#### One flag, three answers

A **flag** is a column of 1s and 0s that marks whether something happened. ReadmitFlag marks 30-day readmissions. Each
inpatient stay is an **index stay**, meaning a stay that could be followed by a readmission. Because the column holds only
1s, 0s, and blanks, three summaries of it answer three different questions:

| Summary of ReadmitFlag | What it means | All 2025 inpatient stays |
|---|---|--:|
| Count | Index stays (blanks are skipped) | 2,859 |
| Sum | Readmissions (each 1 adds one) | 424 |
| Average | Readmission rate = Sum ÷ Count | 14.8% |

Readmit30 holds the same information as Y and N. Text can only be counted, so a pivot can't average Y and N. That's why
analysts add a numeric flag next to a Y/N column.

#### Formats and names

- **Number format:** click **Number Format** inside Value Field Settings to show 14.8% or 2,652.83. That format belongs to the
  value field, so it survives refreshes and layout changes. Formatting the cells with **Home → Number** can be lost when the
  pivot changes shape.
- **Custom name:** type a **Custom Name** in Value Field Settings, or click the field's header cell in the pivot and type
  over it. *Readmit rate* reads better than *Average of ReadmitFlag*. A custom name can't match a source field name exactly,
  so `ReadmitFlag` is refused, but `Readmit rate` works.
- **Several value fields:** drag the same field into Values twice to show two summaries side by side, such as Count and
  Average of ReadmitFlag. With two or more value fields, a **Σ Values** item appears in the Columns area. Drag it to Rows to
  stack the summaries vertically instead.

> 📋 **Counting people instead of rows:** Count of PatientID counts encounters, because a patient with three visits has three
> rows. A true distinct count needs the Data Model. Tick **Add this data to the Data Model** when you create the pivot, and
> **Distinct Count** appears in Value Field Settings. That's Windows only, and Lesson 4.4 covers it.

### 6. Show Values As: shares, changes, running totals, and ranks

**Summarize Values By** decides *what* to calculate. **Show Values As** decides *how to present* each result compared with
other cells in the pivot: as a share, a change, a cumulative total, or a rank. Open it from the **Show Values As** tab of
Value Field Settings, or right-click a number and choose **Show Values As**.

| Show Values As | Each cell shows | Question it answers |
|---|---|---|
| **No Calculation** | The plain summary (the default) | How many ED visits did each facility have? |
| **% of Grand Total** | Cell ÷ the grand total | What share of all 2025 encounters did each payer type bring? |
| **% of Column Total** | Cell ÷ its column's total | Within ED visits, what is the payer mix? |
| **% of Row Total** | Cell ÷ its row's total | At Ashby Falls, what share of encounters were ED visits? |
| **% Of** | Cell ÷ one chosen base item | Each facility's volume as a % of Bluestone Memorial's |
| **% of Parent Row Total** (also Column and Parent Total) | Cell ÷ its parent group's total | With ServiceLine above PayerType in Rows: each payer's share within its service line |
| **Difference From** | Cell − the base item | How many more ED visits than last month? |
| **% Difference From** | (Cell − base item) ÷ base item | Percent change from last month |
| **Running Total In** | Cumulative sum down the base field | Year-to-date admissions |
| **% Running Total In** | Running total ÷ the total | When did we reach half of the year's volume? |
| **Rank Smallest to Largest** / **Rank Largest to Smallest** | 1, 2, 3… within the base field | Which payer type ranks first by volume? |
| **Index** | How far the cell is above (>1) or below (<1) what its row and column totals predict | An advanced check for unusual combinations |

Several options ask for a **base field**, the field to compare along (usually the one in Rows), and a **base item**, the item
to compare with. The base item can be a specific item, such as *Bluestone Memorial Hospital*, or **(previous)** and
**(next)**, which mean the neighboring item.

**% of Row Total.** With FacilityName in Rows, EncounterType in Columns, and Count of EncounterID shown as % of Row Total,
the Ashby Falls row reads Emergency 53.3%, Inpatient 38.6%, Observation 8.1%, and Grand Total 100%. More than half of Ashby
Falls' encounters are ED visits, a typical profile for a community hospital.

**Difference, % difference, and running total.** Put EncounterType in Filters (select Emergency), AdmitDate in Rows grouped
by Quarters (section 7), and EncounterID in Values four times. Leave the first copy as a count. Show the second as
Difference From and the third as % Difference From, both with the quarters field as the base field and **(previous)** as the
base item. Show the fourth as Running Total In the quarters field:

| Row Labels | Count of EncounterID | Difference From (previous) | % Difference From (previous) | Running Total In |
|---|--:|--:|--:|--:|
| Qtr1 | 1,037 | | | 1,037 |
| Qtr2 | 885 | −152 | −14.7% | 1,922 |
| Qtr3 | 872 | −13 | −1.5% | 2,794 |
| Qtr4 | 980 | 108 | 12.4% | 3,774 |
| **Grand Total** | **3,774** | | | |

ED volume dips through spring and summer and climbs again with respiratory season in the fourth quarter. Qtr1 has no
previous quarter, so its difference cells stay blank. The last running total equals the Grand Total of the count, which is a
handy check.

**% of Grand Total and Rank.** With PayerType in Rows and Count of EncounterID added twice, one copy as % of Grand Total and
one as Rank Largest to Smallest (base field PayerType), the pivot shows Government 41.5% (rank 1), Commercial 37.7% (2),
Medicare Advantage 15.1% (3), Self-Pay 5.3% (4), and Workers' Comp 0.5% (5).

> ⚠️ Show Values As replaces the number in that value field. To see the count *and* the percentage, put the field in Values
> twice and change only one copy, as in the examples above.

> ⚠️ Difference From (previous) and Running Total In follow the order of the items in the base field. If you sort the months
> by their counts, "previous" means the row above, not the month before. Sort the months back into calendar order first.

> ⚠️ Percentages use only the records the pivot can see. With a filter on, % of Grand Total divides by the filtered total,
> not by everything in the source.

### 7. Group dates, numbers, and items

**Grouping** combines items into bigger buckets, such as days into months or ages into bands. The source data doesn't
change, and you don't need a helper column.

#### Dates

In Excel 2016 and later, dropping a date field into Rows or Columns often groups it automatically. Excel adds a field such as
**Months (AdmitDate)**, and sometimes **Quarters** and **Years** fields too. The exact names vary by version. To choose the
grouping yourself:

1. Right-click any date (or month) in the pivot and choose **Group…** (or **PivotTable Analyze → Group Field**).
2. Check **Starting at** and **Ending at**. Excel fills them with the first and last dates in the data.
3. Under **By**, click each unit you want. Click a highlighted unit again to deselect it.
4. Click **OK**.

| Select under By | You get | Use it for |
|---|---|---|
| **Months** | Jan … Dec | A monthly trend within one year |
| **Months** and **Years** | 2024 › Jan … Dec, then 2025 › Jan … Dec | A monthly trend across years. Months alone would add January 2024 and January 2025 together |
| **Quarters** (and **Years**) | Qtr1 … Qtr4 | Quarterly reporting |
| **Days** with **Number of days** = 7 | 7-day blocks that start on the Starting at date | Weekly volumes (set Starting at to a Monday) |
| **Hours** | 12 AM … 11 PM | Date-time data, such as ED arrivals by hour of day |

All of this lesson's data is from 2025, so Months alone is safe here.

> 💡 **Tip:** If you don't want automatic grouping, press **Ctrl + Z** (Mac: **⌘ + Z**) right after Excel groups the field. In
> Microsoft 365 and Excel 2019 or later for Windows, you can switch it off for good in **File → Options → Data → Disable
> automatic grouping of Date/Time columns in PivotTables**.

#### Numbers

Right-click a number in a Rows or Columns field, such as an age, and choose **Group…**. The dialog asks for **Starting at**,
**Ending at**, and **By** (the band width). For example, inpatient LOSDays grouped from 0 to 34 by 7:

| Row Labels | Count of EncounterID |
|---|--:|
| 0-6 | 2,294 |
| 7-13 | 517 |
| 14-20 | 44 |
| 21-27 | 3 |
| 28-34 | 1 |
| **Grand Total** | **2,859** |

Each band includes both ends, so *7-13* means 7 through 13 days. Values below Starting at or above Ending at go into two
extra groups whose labels begin with < and >. Every band has the same width. For uneven bands, such as 0–17, 18–64, and
65+, add a band column to the source with IFS (Lesson 2.1) or an approximate-match lookup (Lesson 2.6), then refresh.

#### Items you choose

You can also group any items you select:

1. With PayerName in Rows, click **Medicare**, then Ctrl+click (Mac: ⌘+click) **Silverline Medicare Advantage**.
2. Right-click one of them and choose **Group** (Windows: **Alt + Shift + →**), or choose **PivotTable Analyze → Group
   Selection**.
3. Excel adds a new field, **PayerName2**, and puts the two payers in a group called *Group1*. Click the *Group1* cell, type
   `All Medicare`, and press **Enter**.

The new outer level shows *All Medicare* with 4,313 encounters, and every other payer becomes a one-item group of its own.
Click the **−** buttons to collapse the groups and compare them. To undo any grouping, right-click a grouped item and choose
**Ungroup** (Windows: **Alt + Shift + ←**).

> ⚠️ **"Cannot group that selection."** Excel can group a date or number field only when every value in the source column is
> a real date or number. A single text value or blank cell blocks grouping. Fix the column (Lesson 3.3), refresh, and try
> again.

> ⚠️ **Grouping is shared.** Pivots built from the same source share one **pivot cache**, Excel's internal copy of the data.
> Grouping is stored in the cache, so grouping AgeAtAdmit in one pivot groups it in every pivot built from tblEncounters.

### 8. Sort and filter

**Sorting.** Right-click a number and choose **Sort → Sort Largest to Smallest** to order the items by that value field.
Right-click a label and choose **Sort → Sort A to Z** to order them by name. Grouped months and quarters sort in calendar
order, not alphabetically. You can also drag an item to a new position: select its label cell, point at the cell border, and
drag. In a two-level layout, sorting the inner field orders it *within* each outer item.

**Filtering.** Every Rows and Columns field has a filter button. In the default Compact Form there's one **Row Labels**
button, with a **Select field** box at the top of its menu that chooses which field you're filtering.

| Filter | Keeps | Example |
|---|---|---|
| Checkbox list | The items you tick | Only the three hospitals |
| **Label Filters** | Items whose label matches a text rule (contains, begins with, between…) | DxDescription contains "pain" |
| **Value Filters** | Items whose summary passes a number rule (greater than, between…) | Payer types with at least 30 index stays |
| **Value Filters → Top 10** | The top or bottom N items, the top N percent, or the items that make up a chosen sum | The top 5 ED diagnoses |
| **Filters** area | Records that match, for the whole pivot | EncounterType = Emergency |

**Worked example: top 5 ED diagnoses.** Put EncounterType in Filters (select Emergency), DxDescription in Rows, and
EncounterID in Values. Open **Row Labels → Value Filters → Top 10…**, change 10 to 5, keep *Items* and *Count of
EncounterID*, and click **OK**. Then sort largest to smallest:

| Row Labels | Count of EncounterID |
|---|--:|
| Acute upper respiratory infection, unspecified | 419 |
| Unspecified abdominal pain | 312 |
| Chest pain, unspecified | 264 |
| Urinary tract infection, site not specified | 249 |
| Low back pain, unspecified | 213 |
| **Grand Total** | **1,457** |

Look at the Grand Total: 1,457, not 3,774. A regular PivotTable's totals, and the percentages built on them, include only
the items that pass its filters. When you need a share of *all* ED visits, read the total before you filter, or remove the
filter.

A value filter or Top 10 filter on an inner Rows field works inside each outer item. With ServiceLine above PayerType, a Top 3
filter on PayerType shows the top three payer types within every service line.

> 💡 **Tip:** The Filters area shows *(Multiple Items)* when you pick more than one item, which hides what's selected. A slicer
> (next section) shows every selection at a glance.

> 💡 **Tip:** On Windows, **PivotTable Analyze → Options ▾ → Show Report Filter Pages…** copies the pivot onto a new sheet for
> each item of a Filters field, such as one sheet per facility.

### 9. Slicers and timelines

A **slicer** is a panel of buttons that filters a pivot, one button per item. You used slicers on Tables in Lesson 3.1, and
they work the same way here. A **timeline** is a slicer for a date field: a bar of periods that you click or drag across.

**Insert a slicer:**

1. Click inside the pivot.
2. Choose **PivotTable Analyze → Insert Slicer** (or **Insert → Slicer**).
3. Tick one or more fields, such as **PayerType**, and click **OK**.
4. Click a button to filter. Ctrl+click (Mac: ⌘+click) adds more buttons, or turn on the **Multi-Select** button in the
   slicer's header. Dimmed buttons have no data under the current filters.
5. To clear it, click the **Clear Filter** button (a funnel with a red X) at the slicer's top right, or select the slicer
   and press **Alt + C** (Windows).

**Insert a timeline:**

1. Click inside the pivot and choose **PivotTable Analyze → Insert Timeline** (or **Insert → Timeline**).
2. Tick a date field, such as **AdmitDate**, and click **OK**. Only fields that hold real dates are listed.
3. Use the time-level menu at the timeline's top right to choose **YEARS**, **QUARTERS**, **MONTHS**, or **DAYS**.
4. Click a period to select it. Drag the handles at either end of the selection, or Shift+click another period, to select a
   longer range.

Selections combine just as they did on Tables: buttons in *one* slicer combine with OR (Government *or* Medicare
Advantage), and separate slicers, timelines, and filters combine with AND.

**One slicer, several pivots.** A slicer starts out connected only to the pivot you inserted it from. To make it filter other
pivots, select the slicer and choose **Slicer → Report Connections** (Excel 2010: **PivotTable Connections**), then tick the
pivots to connect. Timelines have the same command on the **Timeline** tab. The pivots must share the same source data, which
pivots built from tblEncounters do.

| | Filters area | Slicer | Timeline |
|---|---|---|---|
| Shows what's selected | Only *(Multiple Items)* when you pick several | Every selected button | The selected period |
| Works on | Any field | Any field | Date fields only |
| Filters several pivots at once | No | Yes, with Report Connections | Yes, with Report Connections |
| Takes up | One cell | A box on the sheet | A wide bar on the sheet |
| Best for | Quick analysis for yourself | Reports other people click through | Picking date ranges |

> ⚠️ A slicer keeps filtering even when it's scrolled out of view. If a pivot's numbers look too small, look for a slicer, a
> timeline, or a Filters selection before you suspect the data.

> 📋 **Version note:** Slicers for PivotTables need Excel 2010 or later on Windows. Timelines need Excel 2013 or later on
> Windows, or Excel 2019 or later on a Mac. Microsoft 365 supports both on Windows and Mac.

### 10. Calculated fields

A **calculated field** is a new field defined by a formula that uses other fields. It's stored in the pivot cache, not in the
source Table, so it's available to every pivot that shares the cache.

1. Click inside the pivot.
2. Choose **PivotTable Analyze → Fields, Items, & Sets → Calculated Field…** (Excel 2010: **Options → Calculations → Fields,
   Items, & Sets**).
3. In **Name**, type a name such as `ChargesPerDay`.
4. In **Formula**, delete the `0`. Then double-click **TotalCharges** in the field list, type `/`, and double-click
   **LOSDays**, so the box reads `=TotalCharges/LOSDays`.
5. Click **Add**, then **OK**. Excel adds *Sum of ChargesPerDay* to Values.

**The one rule to remember: a calculated field sums first, then applies the formula.** For each cell of the pivot, Excel adds
up TotalCharges and LOSDays over the records in that cell, and then divides. For all 2025 inpatient stays, that's

```
Sum of TotalCharges ÷ Sum of LOSDays = 93,956,770.78 ÷ 13,251 = 7,090.54 charges per patient day
```

Averaging each stay's own charges ÷ days would give 8,172.22 instead, and four same-day stays (0 days) would fail with
#DIV/0!. The two numbers answer different questions. The ratio of sums weights each stay by its days, which is what finance
means by charges per patient day. The average of ratios weights every stay equally. Excel labels the field *Sum of
ChargesPerDay* even though it's a ratio.

| A calculated field can… | A calculated field can't… |
|---|---|
| Use + − * / ^, constants, and single-value functions such as IF and ROUND | Refer to cells, ranges, or named ranges |
| Combine any number fields in the source | Count rows, because each field in the formula is already a sum |
| Appear in every pivot that shares the cache | Work in a pivot built on the Data Model. There you write a DAX measure instead (Lesson 4.4) |

To edit or delete a calculated field, open the same dialog, pick it from the **Name** dropdown, and click **Modify** or
**Delete**.

> 💡 **Tip:** To divide by a number of rows, give the source a helper column of 1s, such as `IndexStay` = 1 on every inpatient
> row. Then a calculated field like `=ReadmitFlag/IndexStay` divides readmissions by index stays.

> ⚠️ The same menu offers **Calculated Item**, a formula that combines items of one field, such as Medicare + State Medicaid.
> Calculated items are easy to double-count in totals, and a field that has one can't be grouped. Group the items (section 7)
> or add a column to the source instead.

### 11. GETPIVOTDATA: read a pivot from a formula

Sometimes a formula elsewhere in the workbook needs a number from a pivot, for example on a summary sheet. **GETPIVOTDATA**
looks that number up by its field and item names instead of its address:

```
=GETPIVOTDATA(data_field, pivot_table, [field1, item1], [field2, item2], ...)
```

| Argument | Meaning | Example |
|---|---|---|
| `data_field` | The value field's name in quotes: the source field name, or the name shown in the pivot | `"TotalCharges"` or `"Sum of TotalCharges"` |
| `pivot_table` | Any cell inside the pivot, usually its top-left cell | `$A$3` |
| `field1, item1, …` | Pairs that pick the cell you want | `"FacilityName","Ashby Falls Community Hospital"` |

With no field and item pairs, GETPIVOTDATA returns the grand total of the value field.

You rarely type it yourself. Type `=` in a cell outside the pivot, click a value cell inside the pivot, and press **Enter**.
Excel writes the GETPIVOTDATA formula for you instead of a plain reference like `=B6`. In the section 4 pivot (EncounterType
in Rows, Count of EncounterID in Values, top-left cell A3), clicking the Observation count gives:

```
=GETPIVOTDATA("EncounterID",$A$3,"EncounterType","Observation")      → 593
```

A plain reference such as `=B6` points at a *position*. After someone sorts, filters, or adds a field, B6 may hold a
different number. GETPIVOTDATA points at a *meaning*, so it keeps finding the Observation count wherever it moves.

- **Make it interactive:** replace a typed item with a cell reference, such as `=GETPIVOTDATA("EncounterID",$A$3,
  "EncounterType",F2)`, where F2 holds a drop-down list of encounter types (Lesson 3.2).
- **It needs a visible item:** if the item is filtered out, or its field isn't in the pivot, GETPIVOTDATA returns #REF!.

> 💡 **Tip:** If you want plain references when you click, turn the feature off on Windows with **PivotTable Analyze →
> Options ▾ → Generate GetPivotData** (the small arrow next to **Options** at the left of the tab), or in **File → Options →
> Formulas → Use GetPivotData functions for PivotTable references**. On any platform, you can type the address (`=B6`)
> instead of clicking.

### 12. Layout and design

The **Design** tab controls how the pivot looks. Start with the report layout:

| Report layout | What it looks like | Use it when |
|---|---|---|
| **Compact Form** (default) | All Rows fields share one column, indented, under *Row Labels* | Exploring. It saves width |
| **Outline Form** | Each Rows field in its own column, with subtotals at the top of each group | Long reports read from top to bottom |
| **Tabular Form** | Each Rows field in its own column, with subtotals at the bottom, like a flat table | Reading combinations side by side, or copying results elsewhere |

Here are the first rows of a pivot with FacilityName above EncounterType in Compact Form:

```
Row Labels                         Count of EncounterID
⊟ Ashby Falls Community Hospital                    956
     Emergency                                      510
     Inpatient                                      369
     Observation                                     77
⊟ Bluestone Memorial Hospital                     5,188
     Emergency                                    2,687
     Inpatient                                    2,082
     Observation                                    419
…
```

And the first rows of the same pivot in Tabular Form with **Repeat All Item Labels** and no subtotals:

```
FacilityName                     EncounterType   Count of EncounterID
Ashby Falls Community Hospital   Emergency                        510
Ashby Falls Community Hospital   Inpatient                        369
Ashby Falls Community Hospital   Observation                       77
Bluestone Memorial Hospital      Emergency                      2,687
Bluestone Memorial Hospital      Inpatient                      2,082
Bluestone Memorial Hospital      Observation                      419
…
```

| Option | Where | What it does |
|---|---|---|
| **Repeat All Item Labels** | Design → Report Layout | Repeats the outer label on every row (Outline and Tabular Form) |
| **Subtotals** | Design → Subtotals | Shows subtotals at the top or bottom of each group, or hides them |
| **Grand Totals** | Design → Grand Totals | Turns grand totals on or off for rows and columns separately |
| **Blank Rows** | Design → Blank Rows | Adds an empty row after each group |
| **PivotTable Styles** | Design → styles gallery | Banding, header colors, and borders |
| **For empty cells show** | PivotTable Options → Layout & Format | Shows 0 instead of a blank where a combination has no records |
| **Autofit column widths on update** | PivotTable Options → Layout & Format | Untick it to keep your column widths when you refresh |
| **+/− buttons** | PivotTable Analyze → +/− Buttons | Shows or hides the expand and collapse buttons |

Open **PivotTable Options** by right-clicking inside the pivot, or with **PivotTable Analyze → Options**.

### 13. Refresh, change the source, and drill down

A pivot doesn't recalculate the way a formula does. It summarizes the pivot cache, a snapshot of the source taken when you
built or last refreshed it. When the source changes, **refresh** the pivot to take a new snapshot:

| To do this | Windows | Mac |
|---|---|---|
| Refresh the selected pivot | **Alt + F5**, or **PivotTable Analyze → Refresh** | **PivotTable Analyze → Refresh**, or right-click → **Refresh** |
| Refresh every pivot and query in the workbook | **Ctrl + Alt + F5**, or **Data → Refresh All** | **Data → Refresh All** |
| Refresh automatically each time the file opens | **PivotTable Options → Data → Refresh data when opening the file** | The same option in **PivotTable Options** |

Because the source is a Table, a refresh picks up rows added to tblEncounters. If a pivot was built on a fixed range, choose
**PivotTable Analyze → Change Data Source** and type the Table name, `tblEncounters`, to fix it for good.

**Drill down.** Double-click any value cell, or right-click it and choose **Show Details**. Excel inserts a new sheet with a
Table of every source row behind that number, with all of the source's columns. Use it to answer *which encounters are
these?*: to find an outlier, check a surprising total, or hand a work list to a colleague.

> ⚠️ The detail sheet is a copy. It isn't linked to the pivot and never refreshes, so delete it when you're done.

> ⚠️ **Patient data travels with the pivot.** The pivot cache stores every source row, so anyone who receives a workbook with
> a pivot can double-click it and see patient-level detail, even if you deleted the source sheet. To share only the summary,
> copy the pivot and use **Paste Special → Values** in a new workbook.

> ⚠️ *A PivotTable report cannot overlap another PivotTable report.* Pivots grow when you add fields. Put each pivot on its
> own sheet, or leave plenty of empty space around it.

To move or delete a pivot, use **PivotTable Analyze → Move PivotTable**, or choose **PivotTable Analyze → Select → Entire
PivotTable** and press **Delete**.

### 14. PivotCharts

A **PivotChart** is a chart tied to a pivot. It plots exactly what the pivot shows and changes whenever you filter or rearrange
the pivot. Click inside a pivot and choose **PivotTable Analyze → PivotChart** (or **Insert → PivotChart**). On Windows,
**Alt + F1** adds the chart to the current sheet and **F11** puts it on a sheet of its own. The chart's field buttons and any
connected slicers filter the chart and the pivot together. Lesson 3.5 covers choosing and designing charts.

### 15. Keyboard shortcuts and version notes

| Action | Windows | Mac |
|---|---|---|
| Insert a PivotTable | **Alt, N, V** (in Microsoft 365, then **T** for From Table/Range) | **Insert → PivotTable** |
| Open the field list | **PivotTable Analyze → Field List** | **PivotTable Analyze → Field List** |
| Open a field's filter menu | **Alt + ↓** on its header cell | **⌥ + ↓** on its header cell |
| Value Field Settings | Right-click a value → **Value Field Settings…** | **ⓘ** next to the field in the Values area |
| Group selected items | **Alt + Shift + →** | Right-click → **Group…** |
| Ungroup | **Alt + Shift + ←** | Right-click → **Ungroup** |
| Drill down (Show Details) | Double-click a value | Double-click a value |
| Refresh this pivot | **Alt + F5** | **PivotTable Analyze → Refresh** |
| Refresh all | **Ctrl + Alt + F5** | **Data → Refresh All** |
| Select several slicer buttons | **Ctrl** + click | **⌘** + click |
| Clear the selected slicer | **Alt + C** | **Clear Filter** button |
| PivotChart on this sheet / on its own sheet | **Alt + F1** / **F11** | **Insert → PivotChart** |

| Feature | Available in |
|---|---|
| PivotTables, Value Field Settings, grouping, calculated fields, GETPIVOTDATA | Every desktop version of Excel for Windows and Mac |
| Rank, % of Parent, and % Running Total In under Show Values As, and Repeat All Item Labels | Excel 2010 and later |
| Slicers for PivotTables | Excel 2010 and later (Windows), current Mac versions |
| Timelines | Excel 2013 and later (Windows), Excel 2019 and later (Mac) |
| Automatic date grouping | Excel 2016 and later |
| Distinct Count and the Data Model | Excel for Windows (Lesson 4.4) |
| Excel for the web | Creates and edits PivotTables and slicers. Some options in this lesson are missing or limited |

## 🧪 Hands-on practice

Download [`3.4-pivottables.xlsx`](3.4-pivottables.xlsx) and open the **Practice** sheet. Tasks 1–3 build one pivot step by
step, and tasks 7–9 reuse one monthly pivot. Type each answer in the yellow cell. The **Check** column turns green when
you're right.

<!-- BEGIN GENERATED: practice -->
Every task uses tblEncounters on the Encounters sheet: all 11,145 encounters that began in 2025. Build each PivotTable on a new worksheet (Insert → PivotTable → New Worksheet), or rearrange the one you already have. Then type the number the pivot shows into the yellow cell. Type the value itself rather than a reference to a pivot cell, because a reference like =B7 points somewhere else as soon as you rearrange the pivot. Task 13 is the exception.

| # | Task | Hint |
|:-:|------|------|
| 1 | Create a PivotTable from tblEncounters on a new worksheet. Put FacilityName in Rows and EncounterID in Values. How many 2025 encounters did Cedar Ridge Medical Center have? | Insert → PivotTable, then drag fields into the four areas |
| 2 | Rearrange the pivot. Remove FacilityName, put EncounterType in Filters and select Inpatient, put ServiceLine in Rows, and put TotalCharges in Values. What were the total charges for Cardiovascular inpatient stays? Enter the amount to the cent. | The filter button for the Filters area appears above the pivot |
| 3 | Drill down: double-click the Cardiovascular total in that pivot. Excel lists the stays behind the number on a new sheet. Sort that list by TotalCharges, largest first. What is the EncounterID of the most expensive Cardiovascular inpatient stay? | Double-click a value cell (Show Details) |
| 4 | Build a pivot of average ED charges by payer type: EncounterType = Emergency in Filters, PayerType in Rows, and TotalCharges in Values. Change the summary from Sum to Average. What was the average charge for a Self-Pay ED visit? Round to 2 decimal places. | Right-click a value → Summarize Values By, or Value Field Settings |
| 5 | Build a readmission pivot: EncounterType = Inpatient in Filters, ServiceLine in Rows, and ReadmitFlag in Values. Notice which summary Excel picks, then change it to Average and format it as a percentage. What was the 30-day readmission rate for the Medicine service line? Enter it as a percentage to 1 decimal place. | Average of a 1/0 column is the share of 1s |
| 6 | Build an ED payer-mix pivot: PayerName in Rows, EncounterType in Columns, and EncounterID in Values. Show the values as % of Column Total. What percentage of Emergency encounters were billed to State Medicaid? Enter it to 1 decimal place. | Value Field Settings → Show Values As |
| 7 | Group dates by month: EncounterType = Emergency in Filters, AdmitDate in Rows, and EncounterID in Values. Group AdmitDate by Months. Which month of 2025 had the fewest ED visits? Type the month's three-letter name as the pivot shows it (for example, Mar). | Right-click a date → Group…, then sort by the count |
| 8 | In the same pivot, add EncounterID to Values a second time and show it as Difference From the (previous) month. By how many visits did December's ED volume differ from November's? Type a negative number if December was lower. | Show Values As → Difference From, Base item (previous) |
| 9 | Change the EncounterType filter to Inpatient, and change the second value field to Running Total In the months field. How many inpatient stays began from January 1 through June 30, 2025 (the running total on the Jun row)? | Show Values As → Running Total In |
| 10 | Group numbers into bands: EncounterType = Inpatient in Filters, AgeAtAdmit in Rows, and EncounterID in Values. Group AgeAtAdmit starting at 0, ending at 99, by 10. How many inpatient stays were for patients aged 70–79? | Right-click an age → Group… (Starting at, Ending at, By) |
| 11 | Build a new pivot with FacilityName in Rows, EncounterType in Columns, and EncounterID in Values. Insert a slicer for PayerType and a timeline for AdmitDate. In the slicer, select both Government and Medicare Advantage. In the timeline, switch to QUARTERS and select Q4 2025. How many Emergency encounters does Cedar Ridge Medical Center show? | PivotTable Analyze → Insert Slicer and Insert Timeline. Ctrl+click (Mac: ⌘+click) picks a second button |
| 12 | Add a calculated field named ChargesPerDay with the formula =TotalCharges/LOSDays. Use a pivot with EncounterType = Inpatient in Filters, FacilityName in Rows, and ChargesPerDay in Values. What is ChargesPerDay for Ashby Falls Community Hospital? Round to 2 decimal places. | PivotTable Analyze → Fields, Items, & Sets → Calculated Field |
| 13 | On a new sheet, build a pivot with FacilityName in Rows, EncounterType in Columns, and TotalCharges in Values. Then click this task's yellow cell, type =, switch to the pivot sheet, click the Cedar Ridge Medical Center × Observation cell, and press Enter. Excel writes a GETPIVOTDATA formula. What does it return? Leave the formula in the cell. | If you get a plain reference such as =Sheet7!D8 instead, turn Generate GetPivotData back on |
<!-- END GENERATED: practice -->

## ✅ Answer key

The workbook hides three key sheets. **Answer Key** and **Bonus Key** list each answer with its build steps, and their
*Live result* column runs a COUNTIFS, SUMIFS, or AVERAGEIFS formula that reaches the same number without a pivot.
**Pivot Key** shows every finished pivot in full, so you can compare your whole layout. To see them, right-click any sheet tab
and choose **Unhide…**. The same answers are below, collapsed so you don't see them by accident.

<!-- BEGIN GENERATED: answers -->
<details>
<summary><b>🔑 Show the answer key</b> — Try every task before opening this.</summary>

**1. Create a PivotTable from tblEncounters on a new worksheet. Put FacilityName in Rows…**

- **Answer:** 1,082
- **Solution:**

1. Click any cell in tblEncounters on the Encounters sheet, then choose **Insert → PivotTable** (Microsoft 365 for Windows: **Insert → PivotTable → From Table/Range**). Check that Table/Range says `tblEncounters`, choose **New Worksheet**, and click **OK**.
2. In the PivotTable Fields pane, drag **FacilityName** to **Rows** and **EncounterID** to **Values**. Excel names the value field *Count of EncounterID*.
3. Read the Cedar Ridge Medical Center row.


EncounterID is text, so Excel summarizes it with Count, which adds 1 for every non-empty cell. Each row is one encounter, so the Grand Total (11,145) equals the number of rows in the Table. That's a quick check that the pivot sees all of the data. Cross-check without a pivot: `=COUNTIFS(tblEncounters[FacilityName],"Cedar Ridge Medical Center")`

**2. Rearrange the pivot. Remove FacilityName, put EncounterType in Filters and select…**

- **Answer:** 13,768,733.65
- **Solution:**

1. Uncheck **FacilityName** in the field list (or drag it out of Rows).
2. Drag **EncounterType** to **Filters**. Open the filter button that appears in B1, pick **Inpatient**, and click **OK**.
3. Drag **ServiceLine** to **Rows** and **TotalCharges** to **Values**. Remove *Count of EncounterID* from Values.
4. Read the Cardiovascular row of *Sum of TotalCharges*.


TotalCharges contains only numbers, so Excel sums it by default. The Filters area filters the whole pivot without adding rows or columns, which keeps the layout simple when you only need one slice. Cross-check without a pivot: `=SUMIFS(tblEncounters[TotalCharges],tblEncounters[EncounterType],"Inpatient",tblEncounters[ServiceLine],"Cardiovascular")`

**3. Drill down: double-click the Cardiovascular total in that pivot. Excel lists the stays…**

- **Answer:** ENC114841
- **Solution:**

1. Double-click the *Sum of TotalCharges* cell for Cardiovascular. Excel inserts a new sheet with a Table of the 354 matching rows.
2. Right-click any TotalCharges value in that Table and choose **Sort → Sort Largest to Smallest** (or use the TotalCharges filter button).
3. Read the EncounterID in the first row.


Drilling down answers "which records make up this number?" The detail sheet holds 354 rows, exactly the stays in the total, and the top one is a 4-day stay for Non-ST elevation (NSTEMI) myocardial infarction with charges of $123,078.09. The detail sheet is a static copy: it doesn't update when the data changes, so delete it when you're done. Cross-check without a pivot: `=INDEX(tblEncounters[EncounterID],MATCH(MAXIFS(tblEncounters[TotalCharges],tblEncounters[EncounterType],"Inpatient",tblEncounters[ServiceLine],"Cardiovascular"),tblEncounters[TotalCharges],0))`

**4. Build a pivot of average ED charges by payer type: EncounterType = Emergency in…**

- **Answer:** 2,811.64
- **Solution:**

1. Set the **EncounterType** filter to **Emergency** and put **PayerType** in **Rows** and **TotalCharges** in **Values**.
2. Right-click any number in the pivot and choose **Summarize Values By → Average** (or **Value Field Settings → Summarize Values By → Average**).
3. Click **Number Format** in Value Field Settings and choose Number with 2 decimals, or Currency.
4. Read the Self-Pay row.


Average divides the sum of the charges by the number of visits in each row (219 Self-Pay ED visits). The Grand Total row is the average over every ED visit, not the average of the payer-type averages, so it weights each payer type by its volume. Cross-check without a pivot: `=AVERAGEIFS(tblEncounters[TotalCharges],tblEncounters[EncounterType],"Emergency",tblEncounters[PayerType],"Self-Pay")`

**5. Build a readmission pivot: EncounterType = Inpatient in Filters, ServiceLine in Rows,…**

- **Answer:** 16.1%
- **Solution:**

1. Set the **EncounterType** filter to **Inpatient**, put **ServiceLine** in **Rows**, and drag **ReadmitFlag** to **Values**. Excel shows *Count of ReadmitFlag*.
2. Open **Value Field Settings**, choose **Average**, click **Number Format**, and pick **Percentage** with 1 decimal place.
3. Read the Medicine row.


Excel picks **Count** because ReadmitFlag has blank cells (every non-inpatient row), and a pivot only defaults to Sum when a column is 100% numbers. Count of ReadmitFlag is the number of index stays, Sum is the number of readmissions, and Average (Sum ÷ Count) is the readmission rate: 252 ÷ 1,562 for Medicine. Average ignores blanks, so the non-inpatient rows can't dilute the rate. Cross-check without a pivot: `=AVERAGEIFS(tblEncounters[ReadmitFlag],tblEncounters[EncounterType],"Inpatient",tblEncounters[ServiceLine],"Medicine")`

**6. Build an ED payer-mix pivot: PayerName in Rows, EncounterType in Columns, and…**

- **Answer:** 19.3%
- **Solution:**

1. Clear the EncounterType filter (or start a new pivot). Put **PayerName** in **Rows**, **EncounterType** in **Columns**, and **EncounterID** in **Values**.
2. Right-click a number → **Show Values As → % of Column Total**.
3. Read the State Medicaid row in the Emergency column.


% of Column Total divides each cell by its column's total, so every column adds up to 100% and you read each encounter type's payer mix down the column. % of Row Total would answer a different question: what share of that payer's encounters were ED visits. Filtering EncounterType to Emergency and choosing % of Grand Total gives the same answer. Cross-check without a pivot: `=COUNTIFS(tblEncounters[EncounterType],"Emergency",tblEncounters[PayerName],"State Medicaid")/COUNTIFS(tblEncounters[EncounterType],"Emergency")`

**7. Group dates by month: EncounterType = Emergency in Filters, AdmitDate in Rows, and…**

- **Answer:** Aug
- **Solution:**

1. Set the **EncounterType** filter to **Emergency**, put **AdmitDate** in **Rows**, and put **EncounterID** in **Values**.
2. If you see individual dates, right-click one → **Group…**, select **Months** only, and click **OK**. (Excel 2016 and later may group by month automatically.)
3. Right-click a count → **Sort → Sort Smallest to Largest**. The first month is the answer.


Aug had 259 ED visits, just below Apr with 263. Grouping sorts 3,774 ED visits into 12 monthly buckets without a helper column. Sorting by value reorders the months, so remember to sort back (**Sort A to Z** on the month labels keeps calendar order) before you compare one month with the next. Cross-check without a pivot: `=COUNTIFS(tblEncounters[EncounterType],"Emergency",tblEncounters[AdmitDate],">="&DATE(2025,8,1),tblEncounters[AdmitDate],"<"&DATE(2025,9,1))` returns 259. Repeat it for each month, or see the key's live formula, which checks all twelve at once.

**8. In the same pivot, add EncounterID to Values a second time and show it as Difference…**

- **Answer:** 44
- **Solution:**

1. Sort the months back into calendar order (right-click a month → **Sort → Sort A to Z**).
2. Drag **EncounterID** into **Values** again. Right-click one of the new numbers → **Show Values As → Difference From…**
3. Base field: the field that shows the months (**AdmitDate**, or **Months (AdmitDate)** if Excel created it). Base item: **(previous)**. Click **OK**.
4. Read the Dec row of the new column.


Difference From (previous) subtracts the month above: 365 − 321 = 44. January is blank because it has no previous month in the data. % Difference From would show the same change as a percentage (13.7%). Cross-check without a pivot: `=COUNTIFS(tblEncounters[EncounterType],"Emergency",tblEncounters[AdmitDate],">="&DATE(2025,12,1))-COUNTIFS(tblEncounters[EncounterType],"Emergency",tblEncounters[AdmitDate],">="&DATE(2025,11,1),tblEncounters[AdmitDate],"<"&DATE(2025,12,1))`

**9. Change the EncounterType filter to Inpatient, and change the second value field to…**

- **Answer:** 1,489
- **Solution:**

1. Set the **EncounterType** filter to **Inpatient**.
2. Right-click a number in the second value column → **Show Values As → Running Total In…** → Base field: the months field → **OK**.
3. Read the Jun row.


Running Total In adds each month to everything above it, so the Jun row is the year-to-date total at the end of June, and the Dec row equals the Grand Total (2,859). The running total follows the order of the months, which is another reason to keep them in calendar order. Cross-check without a pivot: `=COUNTIFS(tblEncounters[EncounterType],"Inpatient",tblEncounters[AdmitDate],"<"&DATE(2025,7,1))`

**10. Group numbers into bands: EncounterType = Inpatient in Filters, AgeAtAdmit in Rows,…**

- **Answer:** 600
- **Solution:**

1. Set the **EncounterType** filter to **Inpatient**, put **AgeAtAdmit** in **Rows**, and **EncounterID** in **Values**.
2. Right-click any age → **Group…** → Starting at **0**, Ending at **99**, By **10** → **OK**.
3. Read the **70-79** row.


Grouping turns 99 distinct ages into ten bands labelled 0-9, 10-19, … 90-99. Each band includes both ends, so 70-79 means ages 70 through 79. It's the busiest band for inpatient care. Cross-check without a pivot: `=COUNTIFS(tblEncounters[EncounterType],"Inpatient",tblEncounters[AgeAtAdmit],">=70",tblEncounters[AgeAtAdmit],"<=79")`

**11. Build a new pivot with FacilityName in Rows, EncounterType in Columns, and EncounterID…**

- **Answer:** 88
- **Solution:**

1. Build the pivot on a new sheet: **FacilityName** in **Rows**, **EncounterType** in **Columns**, **EncounterID** in **Values**.
2. **PivotTable Analyze → Insert Slicer** → tick **PayerType** → **OK**. Click **Government**, then Ctrl+click (Mac: ⌘+click) **Medicare Advantage**.
3. **PivotTable Analyze → Insert Timeline** → tick **AdmitDate** → **OK**. Change the time level (top right of the timeline) to **QUARTERS** and click **Q4**.
4. Read the Cedar Ridge Medical Center row in the Emergency column.


Buttons selected in one slicer combine with OR (Government or Medicare Advantage), and separate filters combine with AND (those payer types and Q4). A timeline is a slicer built for dates: it selects a continuous period at the level you choose. Cross-check without a pivot: `=SUM(COUNTIFS(tblEncounters[EncounterType],"Emergency",tblEncounters[FacilityName],"Cedar Ridge Medical Center",tblEncounters[PayerType],{"Government","Medicare Advantage"},tblEncounters[AdmitDate],">="&DATE(2025,10,1)))`

**12. Add a calculated field named ChargesPerDay with the formula =TotalCharges/LOSDays. Use…**

- **Answer:** 7,243.21
- **Solution:**

1. Click inside a pivot built from tblEncounters that no slicer or timeline is filtering (or build a new one). A leftover slicer selection from task 11 would change the result.
2. **PivotTable Analyze → Fields, Items, & Sets → Calculated Field…**
3. Name: `ChargesPerDay`. Formula: `=TotalCharges/LOSDays` (double-click the fields in the list to insert them). Click **Add**, then **OK**.
4. Set the **EncounterType** filter to **Inpatient**, put **FacilityName** in **Rows**, and keep only *Sum of ChargesPerDay* in **Values**.
5. Read the Ashby Falls Community Hospital row.


A calculated field adds up each field first and then applies the formula: Sum of TotalCharges ÷ Sum of LOSDays for the 369 Ashby Falls Community Hospital stays. That's charges per patient day. It is not the average of each stay's charges ÷ days. That would be $7,685.71, and its one same-day stay would return #DIV/0!. Excel labels the field *Sum of ChargesPerDay* even though it's a ratio. Cross-check without a pivot: `=SUMIFS(tblEncounters[TotalCharges],tblEncounters[EncounterType],"Inpatient",tblEncounters[FacilityName],"Ashby Falls Community Hospital")/SUMIFS(tblEncounters[LOSDays],tblEncounters[EncounterType],"Inpatient",tblEncounters[FacilityName],"Ashby Falls Community Hospital")`

**13. On a new sheet, build a pivot with FacilityName in Rows, EncounterType in Columns, and…**

- **Answer:** 842,579.99
- **Solution:**

1. Build the pivot on a new sheet (Excel names it something like *Sheet7*).
2. Click the yellow answer cell on the Practice sheet and type `=`.
3. Switch to the pivot sheet, click the Cedar Ridge Medical Center × Observation cell, and press **Enter**. Excel writes a formula like:

```
=GETPIVOTDATA("TotalCharges",Sheet7!$A$3,"FacilityName","Cedar Ridge Medical Center","EncounterType","Observation")
```


GETPIVOTDATA looks a value up by its field and item names instead of by its cell address, so it keeps returning the right number when the pivot is sorted, filtered, or rearranged, as long as the item stays visible. If someone filters Cedar Ridge out of the pivot, the formula returns #REF!. Cross-check without a pivot: `=SUMIFS(tblEncounters[TotalCharges],tblEncounters[FacilityName],"Cedar Ridge Medical Center",tblEncounters[EncounterType],"Observation")`

</details>
<!-- END GENERATED: answers -->

## 🏆 Bonus challenge

<!-- BEGIN GENERATED: bonus -->
Bluestone's Quality Committee is preparing for its annual readmissions review and wants to know where 30-day readmissions concentrate. An index stay is any inpatient stay that could be followed by a readmission: every row where ReadmitFlag is 1 or 0. Small groups produce extreme rates by chance (1 readmission among 4 stays is 25%), so the committee only reviews groups with at least 30 index stays. Build the pivots on new sheets and answer the committee's questions.

Work on the **Bonus** sheet of the workbook.

- **B1.** Build a pivot with EncounterType = Inpatient in Filters, ServiceLine and then PayerType in Rows, and ReadmitFlag in Values twice: once as Count (rename it Index stays) and once as Average (rename it Readmit rate, formatted as a percentage). Add a value filter on PayerType that keeps only rows with at least 30 index stays. How many service line × payer type combinations remain? *(Hint: PayerType's filter menu → Value Filters → Greater Than Or Equal To. Tabular Form makes rows easy to count)*
- **B2.** Among the remaining combinations, which has the highest readmission rate? Type it as ServiceLine, PayerType (for example: Medicine, Commercial). *(Hint: Sort the Readmit rate column, or scan it with a color scale)*
- **B3.** What is that combination's readmission rate? Enter it as a percentage to 1 decimal place. *(Hint: Read the Readmit rate cell)*
- **B4.** Now rank diagnoses. In a new pivot (EncounterType = Inpatient in Filters), put DxDescription in Rows and ReadmitFlag in Values as Sum, which is the number of readmissions. Apply a Top 10 filter that keeps the top 3 items. Which diagnosis has the third-highest number of readmissions? Type the description exactly as it appears. *(Hint: Row Labels filter → Value Filters → Top 10…)*
- **B5.** What share of all 2025 inpatient readmissions do those three diagnoses account for together? Enter it as a percentage to 1 decimal place. *(Hint: With the Top 3 filter on, the Grand Total adds only the visible rows)*
<!-- END GENERATED: bonus -->

<!-- BEGIN GENERATED: bonus-answers -->
<details>
<summary><b>🔑 Show the bonus solution</b> — Give it a real try first!</summary>

**B1. Build a pivot with EncounterType = Inpatient in Filters, ServiceLine and then…**

- **Answer:** 22
- **Solution:**

1. New pivot: **EncounterType** in **Filters** (select **Inpatient**); **ServiceLine** then **PayerType** in **Rows**.
2. Drag **ReadmitFlag** to **Values** twice. In **Value Field Settings**, set the first to **Count** with Custom Name `Index stays`, and the second to **Average** with Custom Name `Readmit rate` and a percentage number format.
3. **Design → Report Layout → Show in Tabular Form** and **Design → Subtotals → Do Not Show Subtotals** make each combination one row.
4. Open the filter button on the **PayerType** header (in Compact Form, the Row Labels button, then choose **PayerType** under *Select field*) → **Value Filters → Greater Than Or Equal To…** → *Index stays* · `30` → **OK**.
5. Count the remaining PayerType rows (select the Readmit rate cells above the Grand Total and read **Count** on the status bar).


A value filter on the inner row field is applied within each service line, so it keeps or hides each service line × payer type cell separately. 31 combinations exist in the data, and 9 of them have fewer than 30 index stays. Custom names keep the two ReadmitFlag fields apart and make the filter dialog readable. The key's cross-check uses Microsoft 365 functions you'll meet in Lessons 4.1 and 4.2: `=LET(s,UNIQUE(tblEncounters[ServiceLine]),p,TRANSPOSE(UNIQUE(tblEncounters[PayerType])),n,COUNTIFS(tblEncounters[EncounterType],"Inpatient",tblEncounters[ServiceLine],s,tblEncounters[PayerType],p),SUM(--(n>=30)))`

**B2. Among the remaining combinations, which has the highest readmission rate? Type it as…**

- **Answer:** Cardiovascular, Medicare Advantage
- **Solution:**

1. Right-click a *Readmit rate* value on a payer-type row → **Sort → Sort Largest to Smallest**. In a two-level pivot this sorts the payer types *within* each service line, so the highest rate in each service line moves to the top of its group.
2. Compare those top rows, one per service line, and pick the highest. (Sorting the service lines as well orders them by their overall rate, which doesn't guarantee that the best single combination ends up first.)
3. Optional: **Home → Conditional Formatting → Color Scales** on the Readmit rate cells makes the highest rate stand out.


Cardiovascular stays paid by Medicare Advantage plans had 22 readmissions in 86 index stays. Three groups too small to qualify sit right behind it at 25.0%: Cardiovascular / Self-Pay (4 of 16), Critical Care / Self-Pay (2 of 8), and Neuroscience / Self-Pay (1 of 4). With one more readmission, any of them would top the list, which is why the committee sets a minimum volume before it ranks rates.

**B3. What is that combination's readmission rate? Enter it as a percentage to 1 decimal place.**

- **Answer:** 25.6%
- **Solution:** Read the *Readmit rate* cell on the Cardiovascular / Medicare Advantage row.

22 ÷ 86 = 25.6%, against 14.8% for all inpatient stays. Cross-check without a pivot: `=AVERAGEIFS(tblEncounters[ReadmitFlag],tblEncounters[EncounterType],"Inpatient",tblEncounters[ServiceLine],"Cardiovascular",tblEncounters[PayerType],"Medicare Advantage")`

**B4. Now rank diagnoses. In a new pivot (EncounterType = Inpatient in Filters), put…**

- **Answer:** Heart failure, unspecified
- **Solution:**

1. New pivot: **EncounterType** in **Filters** (**Inpatient**), **DxDescription** in **Rows**, **ReadmitFlag** in **Values** (change it to **Sum**).
2. Open the **Row Labels** filter button → **Value Filters → Top 10…** → **Top** `3` **Items** by *Sum of ReadmitFlag* → **OK**.
3. Sort largest to smallest and read the third row.


Sum of a 1/0 flag counts the 1s, so Sum of ReadmitFlag is the number of readmissions. The Top 10 filter keeps the top N items by any value field, not just 10. The top three are Pneumonia, unspecified organism (47), Chronic obstructive pulmonary disease with (acute) exacerbation (45), Heart failure, unspecified (44), and fourth place has 40. Cross-check (Microsoft 365): `=LET(d,UNIQUE(FILTER(tblEncounters[DxDescription],tblEncounters[ReadmitFlag]=1)),n,COUNTIFS(tblEncounters[DxDescription],d,tblEncounters[ReadmitFlag],1),INDEX(SORTBY(d,n,-1),3))`

**B5. What share of all 2025 inpatient readmissions do those three diagnoses account for…**

- **Answer:** 32.1%
- **Solution:**

1. Note the Grand Total with the Top 3 filter on (136 readmissions).
2. Clear the filter (**Row Labels → Clear Filter From "DxDescription"**) and note the new Grand Total (424).
3. Divide: 136 ÷ 424.


A regular PivotTable totals only the items that survive its filters. With the Top 3 filter on, the Grand Total is 136, so % of Grand Total would show the three diagnoses adding up to 100%. The true denominator is all 424 readmissions, so about 32% of readmissions come from just three diagnoses. Pneumonia, COPD, and heart failure are also conditions in Medicare's Hospital Readmissions Reduction Program, so quality teams watch them closely. Cross-check (Microsoft 365): `=LET(d,UNIQUE(FILTER(tblEncounters[DxDescription],tblEncounters[ReadmitFlag]=1)),n,COUNTIFS(tblEncounters[DxDescription],d,tblEncounters[ReadmitFlag],1),SUM(LARGE(n,SEQUENCE(3)))/SUM(tblEncounters[ReadmitFlag]))`

</details>
<!-- END GENERATED: bonus-answers -->

## Key takeaways

- A PivotTable summarizes a tabular list when you drag fields into Rows, Columns, Values, and Filters. Build it from a Table
  so new rows appear when you refresh.
- Excel sums fields that are all numbers and counts everything else. A 1/0 flag gives you index stays (Count), events (Sum),
  and a rate (Average).
- Show Values As turns the same numbers into shares, changes, running totals, and ranks. Add a field to Values twice to see
  the number and the percentage together.
- Grouping makes months, quarters, and age bands without helper columns, and it's shared by every pivot on the same cache.
- Slicers and timelines are visible filters that can drive several pivots. Every filter also shrinks the totals that
  percentages are based on.
- A calculated field sums first and then applies its formula, which is exactly right for ratios like charges per patient day.
- A pivot doesn't recalculate on its own: refresh it after the data changes. When a formula must read a pivot, use
  GETPIVOTDATA.

<!-- BEGIN GENERATED: nav -->
---

⬅️ **Previous:** [3.3 Cleaning Messy Data](../03-data-cleaning/README.md) · 🏠 [Course home](../../README.md) · **Next:** [3.5 Charts & Data Visualization](../05-charts-visualization/README.md) ➡️
<!-- END GENERATED: nav -->
