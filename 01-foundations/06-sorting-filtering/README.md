# Lesson 1.6 · Sorting & Filtering Data

> **Level:** Beginner · **Time:** about 45 minutes · **Workbook:** [`1.6-sorting-filtering.xlsx`](1.6-sorting-filtering.xlsx)
> **Data:** All 937 emergency department visits at Cedar Ridge Medical Center in 2025: arrival time, day and hour, arrival mode, ESI triage level, chief complaint, two triage vital signs, a shock index formula, door-to-provider minutes, and disposition. Rows shaded orange were flagged by the triage sepsis screen.

Every morning the emergency department (ED) manager at Cedar Ridge gets the same kinds of questions. *Who waited longest to see
a provider? How many ambulance patients were critically ill? Did any patient with a positive sepsis screen wait more than half an
hour?* The answers are all in a 937-row export, but nobody can read 937 rows. **Sorting** puts the rows that matter at the top.
**Filtering** hides everything else. **SUBTOTAL** and **AGGREGATE** then calculate on just the rows you can see, and **Advanced
Filter** answers the either/or questions that ordinary filters can't. In this lesson you'll do all four on a full year of Cedar
Ridge's ED visits, and you'll learn the traps that quietly scramble data or give wrong counts.

## What you'll learn

- Sort by one or several columns, including custom orders
- Filter text, numbers, and dates with AutoFilter (and Top 10, by color)
- Summarize only visible rows with SUBTOTAL and AGGREGATE
- Use Advanced Filter for complex AND/OR criteria and unique lists

## 📖 Guide

The examples use the **EDVisits** sheet of the lesson workbook. It holds one row per ED visit: 937 visits in rows 2–938, columns
A–M. Open it and follow along. You can undo any sort or filter with **Ctrl + Z** (Mac: **⌘ + Z**).

| Column | Header | What it holds |
|:-:|---|---|
| A | EDVisitID | Visit ID. IDs follow arrival order, so sorting this column A to Z restores the original order |
| B | PatientID | The patient. Some patients visited more than once |
| C | ArrivalDateTime | Date and time of arrival (a real date-time value) |
| D | ArrivalDay | Day of the week as text: Sun, Mon, Tue, Wed, Thu, Fri, Sat |
| E | ArrivalHour | Hour of arrival from 0 to 23, so 18 means 6:00–6:59 pm |
| F | ArrivalMode | Walk-In, Ambulance, Police, or Air Transport |
| G | ESILevel | Triage urgency from 1 (most urgent) to 5 (least urgent) |
| H | ChiefComplaint | The reason for the visit, in the patient's words as triage recorded them |
| I | HeartRate | Beats per minute at triage |
| J | SystolicBP | The top blood-pressure number at triage. Blank when it wasn't recorded |
| K | ShockIndex | A live formula, `=I2/J2`. Shows `#DIV/0!` where SystolicBP is blank |
| L | DoorToProviderMin | Minutes from arrival to the first provider. Blank for patients who left without being seen |
| M | EDDisposition | Discharged, Admitted, Observation, Transferred, Left AMA, or LWBS |

A few healthcare terms appear throughout the lesson:

- The **ESI level** (Emergency Severity Index) is the triage nurse's urgency rating. ESI 1 patients need immediate life-saving care.
  ESI 5 patients could safely wait.
- **Door-to-provider time** is the number of minutes from arrival to first contact with a physician or advanced practice
  provider. EDs watch it closely, because long waits delay treatment and make patients more likely to leave.
- **LWBS** (left without being seen) means the patient left before a provider saw them. These visits have no door-to-provider time,
  and at Cedar Ridge their blood pressure wasn't recorded either.
- The **shock index** is heart rate ÷ systolic blood pressure. A value near or above 1.0 can be an early warning of shock. (This is
  an educational example, not clinical guidance.)
- Rows shaded **orange** are visits that the triage **sepsis screen** flagged. Sepsis is a life-threatening response to infection,
  and it's treated as time-critical.

### 1. Get the data ready: one tidy list

A **list** is a block of cells with one header row and one record per row. Sorting and filtering behave predictably when your
list follows four rules:

1. **One header row**, with a different name in every column.
2. **No completely blank rows or columns inside the list.** Excel finds the edges of your data by looking for the first fully
   blank row and column, so a blank row in the middle splits the list in two and only half of it gets sorted.
3. **No merged cells.** Excel refuses to sort a range that contains merged cells of different sizes.
4. **One kind of data per column.** Dates belong in a date column and numbers in a number column, because a number stored as text
   sorts as text.

The EDVisits sheet already follows these rules. Click any single cell inside it and press **Ctrl + A** (Mac: **⌘ + A**). Excel
selects A1:M938, the **current region**, which is the block of filled cells around the active cell. When you start a sort or
filter from one cell, Excel works on the current region.

> 💡 **Tip:** Before you sort a list you didn't build, check that it has a column that records the original order, such as an ID or
> a row number. If it doesn't, add one: type 1 and 2 in the first two rows, select both, and double-click the fill handle (Lesson
> 1.2). You can then get back to the starting order at any time. On the EDVisits sheet, EDVisitID does that job.

> 📋 **Excel Tables:** Lesson 3.1 turns lists like this into Excel Tables (**Ctrl + T**; Mac: **⌘ + T**). Tables come with filter
> arrows already on and always sort whole rows. This lesson uses a plain range on purpose, so you meet the dialogs and warnings
> that plain ranges produce.

### 2. Quick sorts on one column

To **sort** is to reorder the rows of a list by the values in one or more columns. The fastest way takes two clicks:

1. Click **one cell** in the column you want to sort by. Don't select the whole column.
2. Click **Data → Sort A to Z** (the A→Z button) for ascending order or **Data → Sort Z to A** for descending order. The same
   commands are under **Home → Sort & Filter** and on the right-click menu under **Sort**.

In the menus, Excel names the commands after the kind of data in the column:

| The column holds | Ascending command | Descending command |
|---|---|---|
| Text, like ChiefComplaint | Sort A to Z | Sort Z to A |
| Numbers, like DoorToProviderMin | Sort Smallest to Largest | Sort Largest to Smallest |
| Dates, like ArrivalDateTime | Sort Oldest to Newest | Sort Newest to Oldest |

Excel moves each **entire row**, so a visit's arrival time, complaint, and disposition stay together. It also decides whether row 1
is a header. Here it recognizes the header because row 1 is formatted differently and holds text above columns of numbers and dates,
so the header stays put.

Try it: click L2 and choose **Sort Smallest to Largest**. The shortest waits rise to the top. Now press **Ctrl + End** (Mac:
**Control + Fn + →**) and look at the bottom. Rows 924–938 have a blank DoorToProviderMin, because those are the 15 LWBS visits.
Sort Largest to Smallest and the blanks are *still* at the bottom.

Excel sorts different kinds of values in a fixed order:

| Ascending (A to Z, smallest to largest) | Descending (Z to A, largest to smallest) |
|---|---|
| 1. Numbers, smallest first (dates and times are numbers) | 1. Error values such as `#DIV/0!` |
| 2. Text, A to Z (capitals and lowercase count as the same) | 2. TRUE, then FALSE |
| 3. FALSE, then TRUE | 3. Text, Z to A |
| 4. Error values | 4. Numbers, largest first |
| 5. **Blank cells, always last** | 5. **Blank cells, always last** |

You can see the error rule on the ShockIndex column. Sort column K Largest to Smallest, and the 15 `#DIV/0!` rows jump *above* the
highest real shock index, because descending order puts errors first.

> ⚠️ **Classic trap: sorting one column on its own.** If you click the column letter H to select the whole column and then click
> Sort A to Z, Excel shows a **Sort Warning** that starts *"Microsoft Excel found data next to your selection."* Choose **Expand the
> selection**. If you choose **Continue with the current selection**, Excel sorts only column H, so every complaint ends up next to
> the wrong patient and nothing warns you afterward. If that happens, press **Ctrl + Z** right away. Clicking a single cell before
> you sort avoids the dialog completely.

> ⚠️ **Hidden rows and columns stay where they are.** Excel doesn't move hidden rows or columns when it sorts, so unhide everything
> first (Lesson 1.1 shows how).

> ⚠️ **Numbers stored as text** sort as text, so "10" lands before "9". When Excel spots them it asks whether to *sort anything that
> looks like a number, as a number*. Lesson 1.2 shows how to fix the cause.

> 📋 **Formulas move with their rows.** K2 holds `=I2/J2`. After any sort, each row's ShockIndex still divides that row's own heart
> rate by that row's own blood pressure, so the column stays correct. A formula that points at a *different* row, such as
> `=C3-C2` (time since the previous arrival), would point at a different pair of visits after a sort. Be careful sorting a list
> that contains formulas like that.

### 3. Sort by several columns with the Sort dialog

A one-click sort uses one column. To sort by several, open the **Sort dialog** with **Data → Sort** (Windows: **Alt, A, S, S**;
Mac: **⌘ + Shift + R**).

Each line in the dialog is a **sort level**. Excel sorts by the first level, then uses the second level only to order rows that
tie on the first, and so on down the list. You can add up to 64 levels.

| Part of the dialog | What it does |
|---|---|
| **Add Level** | Adds a "Then by" line below the selected level |
| **Delete Level**, **Copy Level** | Removes or duplicates the selected level |
| **▲ ▼** arrows | Move the selected level up or down. The top level always wins |
| **Sort On** | What to compare: Cell Values, Cell Color, Font Color, or a conditional-formatting icon |
| **Order** | A to Z, Smallest to Largest, Oldest to Newest, a color, or Custom List… |
| **My data has headers** | Keeps row 1 out of the sort and shows the header names in the Column boxes |
| **Options…** | Case-sensitive sorting, and sorting left to right (columns instead of rows) |

*Worked example: within each arrival mode, who were the most urgent patients, and which of them waited longest?*

1. Click any cell in the data and choose **Data → Sort**. Check that **My data has headers** is ticked.
2. **Sort by** ArrivalMode, Order **A to Z**.
3. **Add Level**: **Then by** ESILevel, **Smallest to Largest**.
4. **Add Level**: **Then by** DoorToProviderMin, **Largest to Smallest**. Click **OK**.

If you start from the original order, the top of the sheet now reads:

| Row | EDVisitID | ArrivalMode | ESILevel | DoorToProviderMin |
|:-:|---|---|:-:|:-:|
| 2 | ED208219 | Air Transport | 1 | 5 |
| 3 | ED207885 | Ambulance | 1 | 5 |
| 4 | ED209966 | Ambulance | 1 | 5 |
| 5 | ED210447 | Ambulance | 1 | 5 |
| 6 | ED209208 | Ambulance | 1 | 4 |

Rows 3–5 tie on all three levels: Ambulance, ESI 1, 5 minutes. Rows that tie on every level stay in whatever order they happened to
be in, and that depends on your earlier sorts. When the exact order matters, for example when you're asked what's in row 11, add a
final **tiebreaker** level such as ArrivalDateTime or EDVisitID, so each row has exactly one possible position.

> 💡 **Tip:** The Sort dialog remembers its levels the next time you open it on the same data. Before a new sort, delete the old
> levels you don't want, or they'll still apply.

### 4. Custom sort orders

Some text has a natural order that isn't alphabetical. Sort the ArrivalDay column A to Z and you get Fri, Mon, Sat, Sun, Thu,
Tue, Wed, which is useless for a weekly staffing report. A **custom list** tells Excel the order you want.

Excel comes with four built-in custom lists: **Sun, Mon, Tue, Wed, Thu, Fri, Sat**; **Sunday, Monday, …**; **Jan, Feb, Mar, …**; and
**January, February, …**. To sort with one:

1. Click a cell in the data and choose **Data → Sort**.
2. Pick the column (ArrivalDay) and set **Order** to **Custom List…**.
3. Select the list you want (here **Sun, Mon, Tue, Wed, Thu, Fri, Sat**) and click **OK**, then **OK** again.

To make your own list, choose **NEW LIST** in that same Custom Lists box, type the entries one per line (press **Enter** after each),
and click **Add**. For example, a report that always lists arrival modes from busiest to quietest would use:

```
Walk-In
Ambulance
Police
Air Transport
```

You can also manage lists outside the Sort dialog: **File → Options → Advanced → General → Edit Custom Lists…** on Windows, or
**Excel → Settings… → Custom Lists** on a Mac (**Excel → Preferences…** in older versions). There you can also import a list
from cells that already hold the entries.

> ⚠️ Spell every entry exactly as it appears in the data, and include every value the column contains. A value that isn't on the
> list doesn't get a place in your order, so it won't land where you expect.

> 📋 Excel saves custom lists on your computer, not inside the workbook. A list you create is available in every workbook you open
> on that computer, and a colleague has to add the list on theirs. Custom lists also drive AutoFill (Lesson 1.2): type *Walk-In*
> and drag the fill handle to get the rest of the list.

### 5. Sort by color

Sometimes the only marker is a color. On the EDVisits sheet, the triage nurse shaded the sepsis-screen visits orange, and no column
says so in words. You can sort by that color:

1. **Data → Sort**. **Sort by** EDVisitID (any column works, because the whole row is shaded).
2. Set **Sort On** to **Cell Color**. Set **Order** to the orange swatch and leave **On Top**.
3. **Add Level**: **Then by** DoorToProviderMin, **Sort On** Cell Values, **Largest to Smallest**. Click **OK**.

The flagged visits now sit at the top, longest wait first. Row 2 is ED207021, a flagged patient who waited 210 minutes to see a
provider, exactly the kind of case a sepsis quality review looks for. For several colors, add one level per color and stack
them in the order you want. A quick shortcut for one color is to right-click an orange cell and choose **Sort → Put Selected Cell
Color On Top**.

> 💡 **Tip:** Color is information only a person can see. No worksheet function can read a fill color, so you can't count orange
> rows with a formula. When a color means something, also record it in a column (for example *SepsisScreen = Positive*). Lesson
> 3.2 shows how conditional formatting can then color the rows from that column automatically.

### 6. AutoFilter: hide the rows you don't need

To **filter** is to hide the rows that don't meet your conditions, without deleting anything. Excel's everyday filter is
**AutoFilter**:

1. Click any cell in the data.
2. Press **Ctrl + Shift + L** (Mac: **⌘ + Shift + F**), or choose **Data → Filter**.

A dropdown **filter arrow** appears in every header cell. Click one (or select the header cell and press **Alt + ↓**; Mac:
**Option + ↓**) to open that column's **filter list**:

```
┌──────────────────────────────────────┐
│ Sort Smallest to Largest             │
│ Sort Largest to Smallest             │
│ Sort by Color                      ▸ │
│ Clear Filter From "ESILevel"         │
│ Filter by Color                    ▸ │
│ Number Filters                     ▸ │  ← Text Filters / Date Filters on other columns
│ ┌──────────────────────────────────┐ │
│ │ Search                        🔍 │ │
│ └──────────────────────────────────┘ │
│ ☑ (Select All)                       │
│ ☑ 1                                  │
│ ☑ 2                                  │
│ ☐ 3                                  │
│ ☐ 4                                  │
│ ☐ 5                                  │
│                   [  OK  ] [Cancel]  │
└──────────────────────────────────────┘
```

Untick the items you don't want, or untick **(Select All)** first and tick only the ones you do, then click **OK**. Excel hides
the other rows. Three clues tell you a filter is on: the row numbers turn blue and skip, the column's arrow changes to a funnel
icon, and the status bar at the bottom of the window reads, for example, **235 of 937 records found**.

> 💡 **Tip:** In large lists, especially ones that contain formulas, the status bar sometimes shows **Filter Mode** instead of the
> count. To count the visible rows anyway, select A2:A938 and read **Count** on the status bar (it counts only visible cells), or
> use `=SUBTOTAL(103,EDVisits!A2:A938)` from section 9.

Filters on several columns combine like this:

| You do this | The logic | Example |
|---|---|---|
| Tick several items in **one** column | **OR**: a row stays if it matches any ticked item | ESILevel 1 *or* 2 |
| Filter **two different** columns | **AND**: a row stays only if it passes every column's filter | ArrivalMode Ambulance *and* ESILevel 1 |
| Need OR **across** columns | AutoFilter can't do it in one step. Use Advanced Filter (section 11) | ESILevel 1 *or* ArrivalMode Police |

The **Search** box at the top of the list finds items that contain what you type, ignoring capitals. Type *pain* in the
ChiefComplaint search box and Excel ticks only the complaints that include "pain" anywhere, such as *Chest Pain*, *Back Pain*, and
*Painful Urination / Fever*. Check the ticked items before you click OK, because a search can catch values you didn't intend. To
combine two searches, run the second one with **Add current selection to filter** ticked.

Clearing and re-running filters:

| To do this | Do this | Windows shortcut | Mac |
|---|---|---|---|
| Turn filter arrows on or off (off also shows every row) | **Data → Filter** | Ctrl + Shift + L | ⌘ + Shift + F |
| Open the selected header's filter list | Click the arrow | Alt + ↓ | Option + ↓ |
| Clear one column's filter | Arrow → **Clear Filter From "…"** | | |
| Clear every filter but keep the arrows | **Data → Clear** | Alt, A, C | Data → Clear |
| Re-run the filters after the data changed | **Data → Reapply** | Ctrl + Alt + L | Data → Reapply |

> ⚠️ **Clear before you start a new question.** Filters stay on until you clear them. If you filter for ambulance arrivals, then
> filter ESILevel for a new question, you're still looking at ambulance arrivals only. Make **Data → Clear** a habit before every
> new filter question.

> ⚠️ **Filters don't update themselves.** If you edit a cell so that a visible row no longer matches, the row stays visible until
> you choose **Data → Reapply**.

> 💡 **Tip:** To filter on the value in the active cell, right-click the cell and choose **Filter → Filter by Selected Cell's
> Value**. Click any *Ambulance* cell and you're one click from an ambulance-only list.

### 7. Text, number, and date filters

Below the sort commands, every filter list offers a submenu of conditions. Excel shows **Text Filters**, **Number Filters**, or
**Date Filters** depending on what most of the column holds.

| Submenu | Conditions you can choose |
|---|---|
| **Text Filters** | Equals, Does Not Equal, Begins With, Ends With, Contains, Does Not Contain, Custom Filter… |
| **Number Filters** | Equals, Does Not Equal, Greater Than, Greater Than Or Equal To, Less Than, Less Than Or Equal To, Between, Top 10…, Above Average, Below Average, Custom Filter… |
| **Date Filters** | Equals, Before, After, Between, Tomorrow, Today, Yesterday, Next/This/Last Week, Month, Quarter, Year, Year to Date, All Dates in the Period ▸ (Quarter 1–4, January–December), Custom Filter… |

Every condition opens the **Custom AutoFilter** dialog, which holds up to two conditions on the same column joined by **And** or
**Or**. You can type the wildcards `*` (any number of characters) and `?` (exactly one character) into text conditions.

Some examples on the lesson data:

- **Text Filters → Begins With** *Fever* on ChiefComplaint shows the 27 *Fever / Altered Mental Status* visits. **Contains** *fever*
  shows far more, because it also catches *Cough / Fever* and *Painful Urination / Fever*.
- **Number Filters → Between** 60 and 120 on DoorToProviderMin shows the 236 visits that waited 60 to 120 minutes. Between includes
  both end values.
- **(Blanks)** at the bottom of the DoorToProviderMin checkbox list shows the 15 LWBS visits, the ones with no wait recorded.

**Dates.** Because ArrivalDateTime holds real dates, its checkbox list is a tree grouped by year, then month, then day. Untick
**(Select All)**, expand **2025**, and tick **December** to see December's 93 visits. **Date Filters → All Dates in the Period →
December** does the same thing for December of every year in the column, which here is just 2025.

> ⚠️ **Between and date-times.** ArrivalDateTime includes a time. If you choose **Date Filters → Between** 12/1/2025 and 12/31/2025,
> Excel reads 12/31/2025 as midnight at the *start* of December 31, so it leaves out every visit later that day. Use the
> checkbox tree, All Dates in the Period, or **is before 1/1/2026** instead.

> 📋 Filters like **Today**, **This Month**, and **Last Year** are measured from today's date on your computer, so they suit live data
> better than a historical export like this one. Run one on this 2025 data next year and the results change.

> 💡 **Tip:** If you'd rather see every date as one flat list instead of the tree, turn off **Group dates in the AutoFilter menu**
> (Windows: **File → Options → Advanced → Display options for this workbook**).

### 8. Top 10, Above Average, and Filter by Color

**Top 10** lives under **Number Filters**. Despite the name, it's a small dialog with three choices:

```
[ Top    ▾ ]  [ 10 ]  [ Items   ▾ ]
  Top / Bottom   1–500    Items / Percent
```

**Top 10 Items** on DoorToProviderMin keeps the ten rows with the longest waits. **Bottom 5 Items** keeps the five shortest.
**Top 10 Percent** keeps the longest tenth of the waits. Top 10 keeps every row whose value is at least the cut-off value, so if two
rows tie at the cut-off you'll see both, and a "Top 10" can show 11 or more rows. Top 10 only filters. It doesn't sort, so the
visible rows stay in their current order. Sort the column afterward if you want them ranked.

**Above Average** and **Below Average** compare each value with the column's mean. Above Average on DoorToProviderMin shows the 359
visits that waited longer than the 48.0-minute average.

**Filter by Color** works like Sort by Color:

1. Open any column's filter arrow (the sepsis rows are shaded across the whole row, so EDVisitID works).
2. Choose **Filter by Color**, then click the orange swatch under **Filter by Cell Color**.

Excel shows the 68 flagged visits. Under the same menu, **No Fill** shows the unflagged ones. You can pick one color per column at a
time, and you can still add ordinary filters on other columns, which combine with AND as usual.

> 📋 **Versions:** Sort and filter by color need Excel 2007 or later, and the Search box needs Excel 2010 or later. Every current
> version of Excel for Windows and Mac has both.

### 9. Count and total only the visible rows: SUBTOTAL

Filtering hides rows, but ordinary functions still see them. Filter ArrivalMode to **Ambulance**, then type these formulas on the
Workspace sheet (any cell that the filter can't hide works):

| Formula | Result (rounded) | What it counts |
|---|---:|---|
| `=COUNTA(EDVisits!A2:A938)` | 937 | Every visit, hidden or not |
| `=SUBTOTAL(103,EDVisits!A2:A938)` | 235 | Only the visible (ambulance) visits |
| `=AVERAGE(EDVisits!L2:L938)` | 48.0 | The average wait of every visit |
| `=SUBTOTAL(101,EDVisits!L2:L938)` | 30.4 | The average wait of the visible visits |

**SUBTOTAL** performs a calculation on only the rows that are still visible:

```
=SUBTOTAL(function_num, ref1, [ref2], ...)
=SUBTOTAL(101, EDVisits!L2:L938)      → average of the visible waits
```

The first argument, **function_num**, picks the calculation. Each calculation has two numbers:

| Calculation | Include rows you hid by hand | Ignore rows you hid by hand |
|---|:-:|:-:|
| AVERAGE | 1 | 101 |
| COUNT (numbers) | 2 | 102 |
| COUNTA (non-empty cells) | 3 | 103 |
| MAX | 4 | 104 |
| MIN | 5 | 105 |
| PRODUCT | 6 | 106 |
| STDEV (sample) | 7 | 107 |
| STDEVP (population) | 8 | 108 |
| SUM | 9 | 109 |
| VAR (sample) | 10 | 110 |
| VARP (population) | 11 | 111 |

Both numbers in each pair **ignore rows hidden by a filter**. They differ only for rows you hide yourself with right-click → Hide:
the 1–11 versions still include those rows and the 101–111 versions skip them. When you just want "what I can see," use the
101–111 versions and you can't go wrong.

A few more things SUBTOTAL does:

- **It ignores other SUBTOTALs** inside its range, so a grand total built with SUBTOTAL doesn't double-count subtotals above it.
- **It only skips hidden rows, not hidden columns.** It's meant for columns of data like the ones in this lesson.
- **It recalculates whenever the filter changes.** Change the filter and the same formulas instantly describe the new set of rows.

To count visible rows reliably, point `SUBTOTAL(103, …)` at a column that's never blank, such as EDVisitID. The LWBS rows are blank
in DoorToProviderMin, so counting that column would come up short.

> 💡 **Tip:** If a filter is on and you click **AutoSum** (Windows: **Alt + =**; Mac: **⌘ + Shift + T**) in the cell just below a
> filtered column, Excel writes `=SUBTOTAL(9, …)` instead of `=SUM(…)`, so the total follows the filter.

> ⚠️ **Put SUBTOTAL where the filter can't hide it.** A formula in row 500 of the data disappears when the filter hides row 500. Put
> summary formulas on another sheet, in rows above the header, or in row 1 at least two columns to the right of the data. The
> filter never hides the header row, and the empty column between keeps Excel from treating your formula as part of the list.

> ⚠️ **A SUBTOTAL result is live.** It changes every time the filter changes. To keep a number, copy the cell and paste it back as a
> value (**Home → Paste → Values**).

> ⚠️ **Errors break SUBTOTAL.** `=SUBTOTAL(104,EDVisits!K2:K938)` should give the highest shock index, but it returns `#DIV/0!`,
> because the range contains `#DIV/0!` cells. MAX, SUM, and AVERAGE do the same. Getting past errors is the job of AGGREGATE.

### 10. AGGREGATE: SUBTOTAL with more functions and a way past errors

**AGGREGATE** works like SUBTOTAL, but it offers 19 calculations instead of 11 and lets you choose what to ignore, including error
values.

```
=AGGREGATE(function_num, options, ref1, [ref2], ...)     reference form (functions 1–13)
=AGGREGATE(function_num, options, array, k)              array form (functions 14–19)

=AGGREGATE(4, 6, EDVisits!K2:K938)        → MAX of the shock index, ignoring errors
=AGGREGATE(14, 5, EDVisits!L2:L938, 2)    → 2nd-largest visible wait
```

| function_num | Calculation | | function_num | Calculation |
|:-:|---|---|:-:|---|
| 1 | AVERAGE | | 11 | VAR.P |
| 2 | COUNT | | 12 | **MEDIAN** |
| 3 | COUNTA | | 13 | **MODE.SNGL** |
| 4 | MAX | | 14 | **LARGE** (needs *k*) |
| 5 | MIN | | 15 | **SMALL** (needs *k*) |
| 6 | PRODUCT | | 16 | **PERCENTILE.INC** (*k* = 0 to 1) |
| 7 | STDEV.S | | 17 | **QUARTILE.INC** (*k* = 0 to 4) |
| 8 | STDEV.P | | 18 | **PERCENTILE.EXC** |
| 9 | SUM | | 19 | **QUARTILE.EXC** |
| 10 | VAR.S | | | |

The bold functions are the ones SUBTOTAL doesn't have. The second argument, **options**, says what to skip:

| options | Ignores hidden rows | Ignores error values | Ignores nested SUBTOTAL / AGGREGATE |
|:-:|:-:|:-:|:-:|
| 0 (or omitted) | | | ✔ |
| 1 | ✔ | | ✔ |
| 2 | | ✔ | ✔ |
| 3 | ✔ | ✔ | ✔ |
| 4 | | | |
| 5 | ✔ | | |
| 6 | | ✔ | |
| 7 | ✔ | ✔ | |

"Hidden rows" here means both rows hidden by a filter and rows you hid by hand.

Examples on the lesson data:

- With no filter on, `=MAX(EDVisits!K2:K938)` returns `#DIV/0!`, but `=AGGREGATE(4,6,EDVisits!K2:K938)` returns about **1.76**,
  the highest shock index of the year.
- With ArrivalMode filtered to Ambulance, `=AGGREGATE(12,5,EDVisits!L2:L938)` returns the **median** visible wait, 23 minutes, and
  `=AGGREGATE(14,5,EDVisits!L2:L938,2)` returns the second-longest visible wait, 124 minutes.

How SUBTOTAL and AGGREGATE compare:

| | SUBTOTAL | AGGREGATE |
|---|---|---|
| Calculations | 11 | 19, including MEDIAN, LARGE, SMALL, PERCENTILE, QUARTILE |
| Rows hidden by a filter | Always ignored | Ignored with options 1, 3, 5, 7 |
| Rows hidden by hand | Ignored with 101–111 | Ignored with options 1, 3, 5, 7 |
| Error values in the range | Returns the error | Ignored with options 2, 3, 6, 7 |
| Works in | Every version | Excel 2010+ (Windows), Excel 2011+ (Mac), the web |

> ⚠️ **Options 0, 2, 4, and 6 count hidden rows.** Unlike SUBTOTAL, AGGREGATE doesn't automatically skip filtered-out rows. If your
> result doesn't change when you change the filter, check the second argument.

> ⚠️ The hidden-row options only work when you give AGGREGATE a plain range like `EDVisits!L2:L938`. If you pass it a calculated
> array (an advanced trick for functions 14–19), it can still skip errors but it can't tell which rows are hidden.

> 💡 **Tip:** If you own the data, the cleaner fix for an error column is to stop the errors at the source. For example,
> `=IFERROR(I2/J2,"")` shows a blank instead of `#DIV/0!` (Lesson 2.1). Use AGGREGATE when you can't or shouldn't change the data,
> such as an export you'll refresh next month.

### 11. Advanced Filter: criteria you write in cells

AutoFilter can't express "ESI 1 **or** arrived by police," because filters on different columns always combine with AND.
**Advanced Filter** reads its conditions from cells you type on the worksheet, so it can do any mix of AND and OR. It can also copy
the matching rows somewhere else and produce a list of unique values.

An Advanced Filter has three parts:

| Part | What it is | Example |
|---|---|---|
| **List range** | The data, including its header row | `EDVisits!$A$1:$M$938` |
| **Criteria range** | A small block of cells: header names on top, conditions below | `Workspace!$A$4:$B$6` |
| **Copy to** (optional) | Where to put the matching rows | `Workspace!$D$4` |

**How to write a criteria range**

1. Copy the header names you need from row 1 of the data. They must match the data headers exactly, so copying is safer than
   typing.
2. Type each condition in the cell under its header.
3. **Conditions on the same row are AND.** **Conditions on different rows are OR.**

AND: ESI 2 patients who came by ambulance.

| ESILevel | ArrivalMode |
|---|---|
| 2 | Ambulance |

OR in one column: ESI 1 or ESI 2.

| ESILevel |
|---|
| 1 |
| 2 |

OR across columns: ESI 1, or anyone who arrived by police. Each row is one way to qualify.

| ESILevel | ArrivalMode |
|---|---|
| 1 | |
| | Police |

Two conditions on the same column, such as a range of values, need the header twice:

| DoorToProviderMin | DoorToProviderMin |
|---|---|
| >=60 | <=120 |

What you can type in a criteria cell:

| You type | It matches | Example result on EDVisits |
|---|---|---|
| `Ambulance` | Text that **begins with** "Ambulance" (not case-sensitive) | Ambulance arrivals |
| `Fever` | Text that **begins with** "Fever", **not** text that merely contains it | 27 *Fever / Altered Mental Status* visits, but no *Cough / Fever* |
| `*Fever*` | Text that contains "Fever" anywhere | Every complaint with fever in it |
| `="=Chest Pain"` | Exactly "Chest Pain" and nothing longer | 63 visits |
| `>=18` | Numbers 18 or more | Arrivals from 6 pm on (ArrivalHour) |
| `<>Discharged` | Everything except "Discharged" | 387 visits |
| `="="` (the cell shows =) | Blank cells | The 15 LWBS rows (DoorToProviderMin) |
| `<>` | Non-blank cells | Every visit with a wait |

To get the exact-match criterion, type `="=Chest Pain"` as shown. The cell then displays **=Chest Pain**, which is what Advanced
Filter reads. (Typing `=Chest Pain` directly doesn't work, because Excel treats anything that starts with = as a formula.) The
blank-cell criterion works the same way: `="="` displays a lone =, which means "equal to nothing."

**How to run it (copying the results to another sheet)**

1. Clear any AutoFilter on the data first (**Data → Clear**).
2. Build the criteria range, for example on the **Workspace** sheet.
3. **Go to the sheet where the results should appear** and click an empty cell there. Excel can only copy filtered data to the
   *active* sheet. Starting from the wrong sheet gives the error *"You can only copy filtered data to the active sheet."*
4. Choose **Data → Advanced** (Windows: **Alt, A, Q**; Mac: **Data → Advanced**).
5. Choose **Copy to another location**. Fill in the **List range**, **Criteria range**, and **Copy to** boxes. You can click in a
   box and then select the cells, even on another sheet.
6. Click **OK**. Excel copies the header row and every matching row.

To filter the data where it is instead, start from the data sheet and choose **Filter the list, in-place**. The rows hide just as
they do with AutoFilter, and **Data → Clear** brings them back.

**Unique lists.** Tick **Unique records only** and Excel keeps only the first copy of each distinct row. If the list range is a
single column, such as `EDVisits!$B$1:$B$938`, you get a list of the different values in that column. Run it on PatientID with an
empty criteria range and you'll find that 588 different patients made the 937 visits.

> ⚠️ **A blank row in the criteria range matches everything.** If your criteria range includes an empty row (for example, you
> selected A4:B7 when your conditions end in row 6), every row qualifies and nothing is filtered. Select only the header row and the
> rows you filled in.

> ⚠️ **Plain text means "begins with."** `Fever` doesn't find *Cough / Fever*. Use `*Fever*` for "contains," and `="=Fever"` for an
> exact match. AutoFilter's Search box behaves differently, because it always searches for "contains."

> ⚠️ **The results are a snapshot.** Copied results don't update when the data changes. Run the Advanced Filter again. **Copy to**
> also writes over whatever is in the destination cells, so point it at an empty area.

> 💡 **Tip:** To copy only some columns, type just those header names (for example EDVisitID, ESILevel, DoorToProviderMin) where the
> results should go, and select those header cells as the **Copy to** range. Excel copies only those columns, in that order.

> 💡 **Tip: formula criteria.** A criteria cell can hold a formula that returns TRUE or FALSE for the first data row. Leave the
> header above it blank (or use a label that isn't a column name), refer to row 2 with a relative reference, and lock everything
> else with `$`. For example, `=L2>AVERAGE($L$2:$L$938)` copies the 359 visits that waited longer than average.

> 📋 After you run an Advanced Filter, Excel creates the names **Criteria** and **Extract** (Formulas → Name Manager). That's normal.
> Excel uses them to remember your last ranges.

> 📋 **Versions:** Advanced Filter is in Excel for Windows and Excel for Mac. Excel for the web doesn't have it.

### 12. Which tool should you use?

| You need to… | Use |
|---|---|
| Put rows in order by one column | A quick sort (Data → A→Z / Z→A) |
| Order by several columns, by a custom order, or by color | The Sort dialog |
| Look at a subset quickly and change your mind often | AutoFilter |
| Count, total, or average only the rows you can see | SUBTOTAL |
| The same, but with errors in the data or a median, percentile, or *k*-th largest | AGGREGATE |
| OR logic across columns, or a lot of conditions | Advanced Filter |
| A copy of matching rows to send or keep | Advanced Filter → Copy to another location |
| A list of the different values in a column | Advanced Filter → Unique records only |

Later lessons add formula-based versions of these tools. COUNTIFS and AVERAGEIFS (Lesson 2.5) calculate on conditions without
filtering anything. The FILTER, SORT, SORTBY, and UNIQUE functions (Lesson 4.1) build live sorted and filtered lists that update by
themselves. Slicers (Lesson 3.1) and PivotTables (Lesson 3.4) add clickable filters to reports.

> 📋 **Shared workbooks:** If several people edit the same workbook at once (saved to OneDrive or SharePoint), a filter or sort you
> apply changes what everyone sees. **View → Sheet View → New** gives you a private view where your sorts and filters don't
> affect anyone else.

### 13. Worked example: the evening ambulance surge

*Question: Cedar Ridge's charge nurses say ambulance arrivals pile up in the evening. How many ambulance patients arrived from 6 pm
on, and how quickly were they seen?*

1. On EDVisits, choose **Data → Clear** (or turn the filter arrows on with **Ctrl + Shift + L** if they're off).
2. Filter **ArrivalMode** to **Ambulance**.
3. Filter **ArrivalHour** with **Number Filters → Greater Than Or Equal To** → **18**. The status bar reads **73 of 937 records
   found**.
4. On the Workspace sheet, write the summary formulas:

   | Formula | Result | Meaning |
   |---|---:|---|
   | `=SUBTOTAL(103,EDVisits!A2:A938)` | 73 | visible visits |
   | `=SUBTOTAL(102,EDVisits!L2:L938)` | 72 | visible visits with a wait (one patient left without being seen) |
   | `=ROUND(SUBTOTAL(101,EDVisits!L2:L938),1)` | 28.7 | average wait in minutes |
   | `=AGGREGATE(12,5,EDVisits!L2:L938)` | 22 | median wait in minutes |
   | `=AGGREGATE(4,5,EDVisits!L2:L938)` | 124 | longest wait in minutes |

5. Sort the visible rows by DoorToProviderMin, Largest to Smallest, to see who waited longest.
6. **Sanity check.** The average wait for *all* visits is 48.0 minutes, with a median of 39. Evening ambulance patients were seen
   faster, which is plausible: ambulance patients tend to be sicker, so triage moves them up the queue. If your average had come out
   at 48.0, the SUBTOTAL would be looking at hidden rows too, so you'd recheck the function number.

The median (22) is lower than the average (28.7) because a handful of long waits pull the average up. That's why emergency
department wait times are usually reported as medians.

## 🧪 Hands-on practice

Download [`1.6-sorting-filtering.xlsx`](1.6-sorting-filtering.xlsx) and open the **Practice** sheet. Most tasks ask you to sort or
filter EDVisits and type what you see. Tasks 11 and 12 ask for formulas. The **Check** column turns green when you're right.

<!-- BEGIN GENERATED: practice -->
Every task uses the EDVisits sheet: 937 visits in rows 2–938, columns A–M. For tasks 1–10, sort or filter, then type what you see: an ID, a row number, or a count. Before each filter task, clear the filters left over from the task before (Data → Clear). Tasks 11 and 12 use SUBTOTAL and AGGREGATE while an ESI 3 filter stays on, so do them last. Mac: Ctrl+Shift+L is ⌘+Shift+F.

| # | Task | Hint |
|:-:|------|------|
| 1 | Sort the EDVisits sheet by DoorToProviderMin, Largest to Smallest. Which EDVisitID is now in row 2? (That's the visit with the longest wait from arrival to first provider contact.) | Click one cell in the column, then use the Z→A button |
| 2 | Use the Sort dialog (Data → Sort) to sort by three levels: ESILevel Smallest to Largest, then DoorToProviderMin Largest to Smallest, then ArrivalDateTime Oldest to Newest. Which EDVisitID is in row 11? | Data → Sort, then Add Level twice. The top level wins |
| 3 | Sort by ArrivalDay using the built-in custom list Sun, Mon, Tue, Wed, Thu, Fri, Sat (Sort dialog → Order → Custom List…). On which row does the first Wednesday (Wed) visit appear? | Sunday, Monday and Tuesday visits come first |
| 4 | Turn on AutoFilter (Ctrl+Shift+L). Show only visits with ArrivalMode = Ambulance and ESILevel 1 or 2. How many visits are visible? | Filter two columns, then read the status bar |
| 5 | Clear the filters (Data → Clear). Use the Search box in the ChiefComplaint filter to show every visit whose complaint contains the word fever anywhere. How many visits are visible? | Search box at the top of the filter list |
| 6 | Clear the filters. Use Number Filters → Top 10 on DoorToProviderMin to show the 10 longest waits. What is the smallest DoorToProviderMin still visible (the 10th-longest wait), in minutes? | Top 10 hides everything except the largest values |
| 7 | Clear the filters. Show only visits that arrived in November 2025 on a Saturday or Sunday. How many are there? | Date tree for the month, then ArrivalDay for the weekend |
| 8 | Clear the filters. Rows shaded orange are visits that triage flagged as a positive sepsis screen. Filter to the orange rows, then also keep only those with DoorToProviderMin greater than 30. How many flagged visits waited more than 30 minutes to see a provider? | Filter by Color, then Number Filters → Greater Than |
| 9 | Clear the filters. Use Advanced Filter to find visits that were ESI level 1 OR waited more than 120 minutes for a provider. Build the criteria range on the Workspace sheet and copy the results there (start Data → Advanced from the Workspace sheet). How many visits does it copy? | Conditions on different rows of the criteria range mean OR |
| 10 | Use Advanced Filter with Unique records only to copy a list of the different ChiefComplaint values to the Workspace sheet. How many different chief complaints are there? Don't count the header. | List range: the ChiefComplaint column only. No criteria range |
| 11 | Filter EDVisits to ESILevel 3 only (clear everything else). In the yellow cell, write a formula that averages DoorToProviderMin for the visible rows only, rounded to 1 decimal place with ROUND. Leave the filter on for task 12. | SUBTOTAL's AVERAGE is function_num 1 (or 101) |
| 12 | Keep the ESI 3 filter on. ShockIndex (HeartRate ÷ SystolicBP) shows #DIV/0! where no blood pressure was recorded, so =SUBTOTAL(104,EDVisits!K2:K938) returns #DIV/0!. Write an AGGREGATE formula that returns the highest ShockIndex among the visible rows while ignoring the errors. (2 decimal places is close enough.) | MAX is function 4. Pick the option that ignores hidden rows AND error values |
<!-- END GENERATED: practice -->

## ✅ Answer key

The workbook has a hidden **Answer Key** sheet (right-click any sheet tab → **Unhide…** → *Answer Key*). Its *Live result* column
recomputes each answer with an independent formula, so you can confirm the numbers without redoing the sort or filter. (The
color task has no live formula, because no formula can see a fill color.) The same
answers are below, collapsed so you don't see them by accident.

<!-- BEGIN GENERATED: answers -->
<details>
<summary><b>🔑 Show the answer key</b> — Try every task before opening this.</summary>

**1. Longest door-to-provider wait (one-column sort)**

- **Answer:** ED208970
- **Solution:**

1. Click any cell in column L (DoorToProviderMin), for example L2.
2. Choose **Data → Sort Largest to Smallest** (the Z→A button). Because you clicked a single cell, Excel sorts the whole block of data with it.
3. Read A2: **ED208970**, a wait of 299 minutes.


Clicking one cell (not selecting the column) lets Excel find the whole data block and keep every row together. Scroll to the bottom: the 15 LWBS visits with a blank wait sit in rows 924–938. Excel always sorts blanks last, whether you sort ascending or descending.

**2. Three-level sort (row 11)**

- **Answer:** ED208106
- **Solution:**

1. Click any cell in the data and choose **Data → Sort**. Check that **My data has headers** is ticked.
2. Sort by **ESILevel**, Sort On **Cell Values**, Order **Smallest to Largest**.
3. Click **Add Level**: Then by **DoorToProviderMin**, **Largest to Smallest**.
4. Click **Add Level** again: Then by **ArrivalDateTime**, **Oldest to Newest**. Click **OK**.
5. Read A11: **ED208106**.


The first level groups the 23 ESI 1 visits at the top (rows 2–24). Inside that group the second level puts the longest waits first. Several ESI 1 patients waited exactly 3 minutes, so the third level (arrival time) breaks the tie. Without a tiebreaker, rows that tie on every level stay in whatever order they happened to be in, which depends on your earlier sorts.

**3. Custom-list sort by weekday**

- **Answer:** 433
- **Solution:**

1. **Data → Sort**. If levels from task 2 are still listed, select each one and click **Delete Level**.
2. Sort by **ArrivalDay**, Sort On **Cell Values**, Order **Custom List…**.
3. In the Custom Lists box pick **Sun, Mon, Tue, Wed, Thu, Fri, Sat**, then **OK** twice.
4. Scroll down column D until the days change from Tue to Wed: row **433**.


A custom list sorts in the order you give it, not alphabetically. Sunday (142 visits), Monday (152) and Tuesday (137) fill rows 2–432, so Wednesday starts in row 433. An A→Z sort would have put the days in the order Fri, Mon, Sat, Sun, Thu, Tue, Wed, and the first Wednesday would land in row 801.

**4. Ambulance arrivals with ESI 1–2 (two-column filter)**

- **Answer:** 137
- **Solution:**

1. Click any cell in the data and press **Ctrl+Shift+L** (Mac: **⌘+Shift+F**), or choose **Data → Filter**.
2. Open the **ArrivalMode** arrow, untick **(Select All)**, tick **Ambulance**, **OK**.
3. Open the **ESILevel** arrow, untick 3, 4 and 5 (or use **Number Filters → Less Than Or Equal To → 2**), **OK**.
4. The status bar reads **137 of 937 records found**.


Filters on different columns combine with AND: a row stays visible only if it passes every column's filter. Ticking 1 and 2 inside one column is OR within that column. If the status bar shows *Filter Mode* instead of a count, select A2:A938 and read **Count** on the status bar.

**5. Chief complaints containing "fever" (search box)**

- **Answer:** 266
- **Solution:**

1. **Data → Clear** removes the filters from task 4 but keeps the filter arrows.
2. Open the **ChiefComplaint** arrow and type **fever** in the **Search** box.
3. The list shrinks to 3 items (Cough / Fever, Fever / Altered Mental Status, Painful Urination / Fever). Click **OK**.
4. The status bar reads **266 of 937 records found**.


The Search box matches text anywhere in the value and ignores case, so it finds complaints that start, end, or contain *fever*. It works the same as **Text Filters → Contains**. Always glance at the ticked items before clicking OK, because a search can catch values you didn't intend.

**6. Top 10 longest waits (10th-longest value)**

- **Answer:** 160
- **Solution:**

1. **Data → Clear**.
2. Open the **DoorToProviderMin** arrow → **Number Filters → Top 10…**.
3. Leave **Top**, **10**, **Items** and click **OK**.
4. Ten rows remain. The smallest of them is **160** minutes (sort the column Largest to Smallest to see it at the bottom).


Top 10 keeps the rows whose value is at least the 10th largest. Here the 10th-longest wait is 160 minutes, so exactly ten rows stay visible. If several rows had tied at the cut-off value, Excel would show all of them, so a Top 10 filter can show more than ten rows. The same dialog does Bottom 10 and Top 10 Percent.

**7. November weekend arrivals (date filter)**

- **Answer:** 26
- **Solution:**

1. **Data → Clear**.
2. Open the **ArrivalDateTime** arrow. In the date tree, untick **(Select All)**, expand **2025**, tick **November**, **OK**. (Or use **Date Filters → All Dates in the Period → November**.)
3. The status bar shows 80 November visits. Now open the **ArrivalDay** arrow, keep only **Sat** and **Sun**, **OK**.
4. The status bar reads **26 of 937 records found**.


Excel groups real dates in the filter list by year, month, and day, so you can tick a whole month at once. There's no built-in *weekend* date filter, which is why ED exports often include a day-of-week column like ArrivalDay. Avoid **Between 11/1/2025 and 11/30/2025** here: the values include times, and 11/30/2025 means midnight at the start of November 30, so visits later that day would be left out.

**8. Sepsis-flagged visits that waited over 30 minutes (filter by color)**

- **Answer:** 29
- **Solution:**

1. **Data → Clear**.
2. Open any column's arrow (EDVisitID works) → **Filter by Color** → pick the orange swatch under **Filter by Cell Color**. 68 rows remain.
3. Open the **DoorToProviderMin** arrow → **Number Filters → Greater Than…** → type **30** → **OK**.
4. The status bar reads **29 of 937 records found**.


Color is the only marker for the sepsis flag, so Filter by Color is the only way to isolate those 68 rows. No worksheet function can read a fill color, which is why this task has no live formula in the key. Combining a color filter on one column with a number filter on another is still AND logic. Sepsis care is time-critical, so these are the waits a quality team reviews first.

**9. Advanced Filter: ESI 1 OR wait over 120 minutes**

- **Answer:** 65
- **Solution:**

1. **Data → Clear** on EDVisits.
2. On **Workspace**, type the criteria range in A4:B6:

   | A | B |
   |---|---|
   | ESILevel | DoorToProviderMin |
   | 1 | |
   | | >120 |

3. Still on Workspace, choose **Data → Advanced**. Select **Copy to another location**. List range: `EDVisits!$A$1:$M$938`. Criteria range: `Workspace!$A$4:$B$6`. Copy to: `Workspace!$D$4`. **OK**.
4. The copy has a header row plus **65** visit rows.


Each criteria row is one way to qualify. Row 5 catches the 23 ESI 1 visits and row 6 catches the 42 waits over 120 minutes. No ESI 1 patient waited that long, so the two groups don't overlap here. When they do, a visit that meets both rows is copied once, not twice. AutoFilter can't do this, because filters on two columns always combine with AND. The criteria headers must match the data headers exactly, so copy them from row 1 of EDVisits rather than typing them.

**10. Unique chief complaints (Advanced Filter)**

- **Answer:** 28
- **Solution:**

1. Click an empty cell on **Workspace** (for example **S4**) and choose **Data → Advanced**.
2. Select **Copy to another location**. List range: `EDVisits!$H$1:$H$938` (just that one column, header included).
3. Leave Criteria range empty. Copy to: `Workspace!$S$4`. Tick **Unique records only**. **OK**.
4. The list has a header plus **28** complaints (select them and read Count on the status bar).


With an empty criteria range every row qualifies, and **Unique records only** keeps the first copy of each distinct value. Because the list range is a single column, the duplicates are judged on that column alone. With the whole table as the list range, a row would only count as a duplicate if every column matched.

**11. SUBTOTAL average of the visible rows (ESI 3)**

- **Answer:** 51.7
- **Solution:** `=ROUND(SUBTOTAL(101,EDVisits!L2:L938),1)`

SUBTOTAL skips rows hidden by a filter, so it averages only the 399 ESI 3 visits. Function 1 would work too, because both 1 and 101 ignore filtered-out rows (101 also ignores rows you hide by hand). AVERAGE-type functions skip blank cells, so the 6 ESI 3 LWBS visits with no wait don't drag the average toward zero. Plain AVERAGE would ignore the filter and average all 922 waits. The key's live cell uses AVERAGEIFS (Lesson 2.5) so it works without a filter. If you clear the filter later, this check turns red, which shows SUBTOTAL responding to the filter.

**12. AGGREGATE maximum, ignoring hidden rows and errors**

- **Answer:** 1.64
- **Solution:** `=AGGREGATE(4,7,EDVisits!K2:K938)`

AGGREGATE(4, 7, range) means MAX (4) while ignoring hidden rows and error values (7). The 6 visible #DIV/0! cells come from LWBS patients whose blood pressure wasn't recorded. SUBTOTAL and MAX return an error as soon as the range holds one, while AGGREGATE steps over it. Option 6 would ignore errors but count the hidden rows, giving the highest shock index of the whole year. A shock index near or above 1.0 suggests the heart is racing to keep blood pressure up, which is why EDs watch it.

</details>
<!-- END GENERATED: answers -->

## 🏆 Bonus challenge

<!-- BEGIN GENERATED: bonus -->
Cedar Ridge's ED medical director is preparing a high-acuity review for the quality committee. It covers every ESI 1 visit, plus ESI 2 patients who arrived by ambulance in the evening (ArrivalHour 18 or later, which means 6:00 pm to 11:59 pm). That rule is an OR across different columns, so AutoFilter can't do it in one step. Before you start, clear every filter on EDVisits (Data → Clear). Practice tasks 11 and 12 will turn red, which is expected. Then insert a new sheet named Review, build the criteria range on the Workspace sheet, and use Advanced Filter (started from the Review sheet) to copy the matching rows to Review!A1.

Work on the **Bonus** sheet of the workbook.

- **B1.** How many visits does your Advanced Filter copy to the Review sheet? Don't count the header row. *(Hint: Two criteria rows: ESI 1 alone, and ESI 2 + Ambulance + >=18 together)*
- **B2.** What is the average DoorToProviderMin of the review visits, rounded to 1 decimal place? *(Hint: AVERAGE over column L of the Review sheet)*
- **B3.** The director lists dispositions by level of care. Create the custom list Admitted, Observation, Transferred, Discharged, Left AMA, LWBS. Sort the Review sheet by EDDisposition with that list, then by DoorToProviderMin Largest to Smallest. Which EDVisitID is in row 44 of the Review sheet? *(Hint: Order → Custom List… → NEW LIST)*
- **B4.** CMS reports ED wait times as medians, because a few very long waits pull an average up. Filter the Review sheet to ESILevel 2 only, then use AGGREGATE to find the median DoorToProviderMin of the visible rows. *(Hint: SUBTOTAL has no median, but AGGREGATE function 12 does)*
<!-- END GENERATED: bonus -->

<!-- BEGIN GENERATED: bonus-answers -->
<details>
<summary><b>🔑 Show the bonus solution</b> — Give it a real try first!</summary>

**B1. Visits in the review list**

- **Answer:** 60
- **Solution:**

1. **Data → Clear** on EDVisits. Insert a sheet (**Shift+F11**) and rename it **Review**.
2. On **Workspace**, build this criteria range in an empty spot (for example A20:C22), with headers copied from EDVisits:

   | ESILevel | ArrivalMode | ArrivalHour |
   |---|---|---|
   | 1 | | |
   | 2 | Ambulance | >=18 |

3. Click **Review!A1**, then **Data → Advanced** → **Copy to another location**. List range `EDVisits!$A$1:$M$938`, Criteria range `Workspace!$A$20:$C$22`, Copy to `Review!$A$1`. **OK**.
4. Review shows a header plus **60** rows (rows 2–61).


Row 21 of the criteria range (ESILevel = 1) catches all 23 ESI 1 visits. Row 22 is an AND: ESI 2 *and* Ambulance *and* ArrivalHour ≥ 18, which adds 37 more. The two rows together are OR. Excel only copies filtered data to the active sheet, so you must start Data → Advanced from the Review sheet, or Excel shows an error.

**B2. Average wait in the review list**

- **Answer:** 15.8
- **Solution:** `=ROUND(AVERAGE(Review!L2:L61),1)` (or `=ROUND(AVERAGE(Review!L:L),1)`, because AVERAGE skips the text header).

The Advanced Filter result is ordinary cells, so regular functions work on it. AVERAGE ignores text and blanks, so even a whole-column reference gives the right answer. The copy is a snapshot: if EDVisits changes, run the Advanced Filter again.

**B3. Custom-list sort of the review list (row 44)**

- **Answer:** ED207754
- **Solution:**

1. On Review, click any cell in the data → **Data → Sort**. Sort by **EDDisposition**, Order **Custom List…**.
2. Select **NEW LIST**, type the six entries one per line (press Enter after each), click **Add**, then **OK**.
3. **Add Level**: Then by **DoorToProviderMin**, **Largest to Smallest**. **OK**.
4. Rows 2–43 hold the 42 Admitted visits, so row 44 is the first Observation visit: **ED207754**.


The review list holds 42 Admitted, 10 Observation and 8 Discharged visits, so the custom order puts Observation in rows 44–53. Row 44 is the Observation patient who waited longest (88 minutes). An A→Z sort would have put Discharged second, and a different visit in row 44. Every disposition appears in the list, even ones the review doesn't contain, so the list works on any month's data. Excel saves a custom list on your computer, so it's available in every workbook you open there.

**B4. Median wait for the ESI 2 evening ambulance patients**

- **Answer:** 19
- **Solution:**

1. On Review, press **Ctrl+Shift+L** and filter **ESILevel** to **2**.
2. Type `=AGGREGATE(12,5,Review!L2:L61)` in a cell the filter can't hide: row 1, two columns past the data (for example O1), or another sheet.
3. It returns **19**.


AGGREGATE(12, 5, range) is MEDIAN (12) that ignores hidden rows (5). SUBTOTAL can't do this, because its eleven functions don't include MEDIAN, LARGE, SMALL, or PERCENTILE. The 37 ESI 2 evening ambulance patients had a median wait of 19 minutes but a mean of 23.7, because a few long waits pull the mean up. That gap is why CMS reports medians.

</details>
<!-- END GENERATED: bonus-answers -->

## Key takeaways

- Click **one cell** before you sort or filter so Excel finds the whole list and keeps every row together. If you see the Sort
  Warning, choose **Expand the selection**.
- The Sort dialog sorts by several levels, by a **custom list** (weekdays, level of care), or by **color**. Add a tiebreaker level
  when the exact order matters. Blanks always sort last.
- AutoFilter ticks within one column are **OR**. Filters on different columns are **AND**. Clear old filters before each new
  question.
- Ordinary functions see hidden rows. **SUBTOTAL** (use 101–111) calculates on visible rows only, and **AGGREGATE** adds MEDIAN,
  LARGE, PERCENTILE, and the ability to skip error values.
- **Advanced Filter** handles OR across columns: conditions on the same criteria row are AND, and different rows are OR. Plain text
  criteria mean "begins with," and you must start from the sheet that receives the copy.
- A color is invisible to formulas. When a color carries meaning, record it in a column as well.

<!-- BEGIN GENERATED: nav -->
---

⬅️ **Previous:** [1.5 Relative, Absolute & Mixed References](../05-cell-references/README.md) · 🏠 [Course home](../../README.md) · **Next:** [2.1 Logical Functions: IF, AND, OR, IFS & More](../../02-formulas-functions/01-logical-functions/README.md) ➡️
<!-- END GENERATED: nav -->
