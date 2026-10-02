# Lesson 4.6 · Building Interactive Dashboards

> **Level:** Advanced · **Time:** about 70 minutes · **Workbook:** [`4.6-dashboards.xlsx`](4.6-dashboards.xlsx)
> **Data:** Monthly KPIs for Bluestone Health's three hospitals, Jan 2024 – Dec 2025 (72 hospital-months), every ED visit's door-to-provider time, and a KPI dictionary with targets and owners. The monthly numbers are built from [`ed_visits.csv`](../../data/README.md#ed_visitscsv), [`encounters.csv`](../../data/README.md#encounterscsv), [`daily_census.csv`](../../data/README.md#daily_censuscsv), [`patient_satisfaction.csv`](../../data/README.md#patient_satisfactioncsv), and [`claims.csv`](../../data/README.md#claimscsv) in the [data dictionary](../../data/README.md).

Every month, Bluestone Health's leaders ask the same questions. Are ED patients waiting too long? Will there be a bed for
the next admission? Are payers denying more claims? An analyst can answer each one with a fresh report, or build a
**dashboard** once and let the leaders answer the questions themselves: pick a hospital and a month, and every number,
status color, and chart on the page changes together. This lesson shows you how to plan a dashboard around the people who'll
read it, build the formulas behind KPI cards, wire dropdowns, slicers, and charts so that one click updates everything, and
then polish the result so it's fast, protected, and readable at a glance.

## What you'll learn

- Plan a dashboard around audience, questions, and KPIs
- Build KPI cards and selector-driven formulas
- Connect slicers to multiple PivotTables and drive charts from selections
- Polish layout, interactivity, and performance

## 📖 Guide

The workbook is laid out the way a real dashboard workbook should be, so you can study it while you read:

| Sheet | Role |
|---|---|
| **Dashboard** | The practice canvas: two yellow selectors (a facility and a month), a finished example card, and room for your chart and combo box |
| **Calc** | The model layer: helper cells that turn the selectors into dates and criteria, plus the 12-month block your chart will read |
| **KPI_Monthly** | Table `tblKPI`: one row per hospital per month, January 2024 – December 2025 |
| **EDWaits** | Table `tblEDWaits`: every ED visit that was seen by a provider, with its door-to-provider minutes |
| **Targets** | Table `tblTargets`: the KPI dictionary, with each KPI's question, definition, direction, target, and owner |
| **Lists** | The dropdown sources: four facility choices (including *All facilities*) and 24 months |
| **Board** | A blank canvas for the bonus |
| **Dashboard Key** *(hidden)* | A finished reference dashboard. Unhide it after the bonus, or whenever you want to see a technique in action |

### 1. What a dashboard is for

A **dashboard** is a single screen that shows a specific audience the few numbers it needs to decide what to do next. That
definition rules out a lot of things people call dashboards:

| | Dashboard | Report | Analysis |
|---|---|---|---|
| Question it answers | "Are we OK, and where should I look?" | "What exactly happened?" | "Why did it happen?" |
| Size | One screen or one printed page | Many rows, many pages | Whatever the question needs |
| Refreshed | Every period, same layout | Every period | Once |
| Reader's time | Seconds | Minutes | An hour |

An **interactive dashboard** adds **selectors**: input cells or controls, such as a facility dropdown, that every number on
the page reads. Instead of 36 copies of the same page (3 hospitals × 12 months), you build one page that can become any of
them.

The test of a good dashboard is the **5-second rule**: within five seconds a reader should know whether things are on track
and where to look first. Try it on your own work. Show the page to a colleague for five seconds, hide it, and ask what they
saw. If they describe the colors instead of the message, the design needs work.

### 2. Plan before you build: audience, questions, and KPIs

Most weak dashboards start in Excel. Good ones start with a conversation about who will read the page and what they need
to decide.

**Start with the audience.** The same data supports very different dashboards:

| Audience | What they decide | Grain | Cadence | Typical KPIs |
|---|---|---|---|---|
| Board quality committee | Where to ask for an improvement plan | System and hospital, by month | Monthly | Readmission rate, HCAHPS top-box %, denial rate |
| Chief Nursing Officer | Staffing and bed flow | Hospital and unit, by day or week | Daily or weekly | Occupancy %, ALOS, ED boarding |
| ED nurse manager | Shift staffing and triage changes | One ED, by hour or shift | Daily | Door-to-provider time, LWBS %, arrivals by hour |

**Then write the questions in plain words**, and pick the one number that answers each one. That number is a **KPI**
(key performance indicator). For Bluestone's monthly operations dashboard:

| Question | KPI |
|---|---|
| Are ED patients leaving before a provider sees them? | LWBS % (left without being seen) |
| How long do ED patients wait to see a provider? | Median door-to-provider minutes |
| Are inpatients staying longer than they need to? | ALOS (average length of stay) |
| Are discharged patients coming back within 30 days? | 30-day readmission rate |
| Will we have a bed for the next admission? | Occupancy % |
| Would patients rate their stay 9 or 10 out of 10? | HCAHPS top-box % |
| Are payers refusing our claims? | Denial rate |

**A KPI is only useful if it's fully defined.** Before you build a card, you should be able to fill in every column of a
**KPI dictionary**, which is what the Targets sheet is:

| Column in `tblTargets` | Why it matters |
|---|---|
| **KPI** | One name, used everywhere. Formulas look the target up by this text, so spell it identically |
| **Question** | Keeps the KPI tied to a decision. A KPI that answers nobody's question doesn't belong on the page |
| **Definition** | Numerator, denominator, which rows count, and which date decides the month. "Readmission rate" means nothing until you say which stays count |
| **Direction** | *Lower is better* or *Higher is better*. Status colors and arrows depend on it |
| **Target** | The line between on target and off target. On target means at or below a lower-is-better target, or at or above a higher-is-better one |
| **Owner** | The person who acts when the KPI goes red. No owner, no action |

A few more planning rules:

- **Keep it short.** Five to nine KPIs fit on one page. If everything is on the dashboard, nothing stands out.
- **Separate volume from performance.** ED visits and discharges give context, but more isn't better or worse, so they get
  no target and no red or green.
- **Mix leading and lagging KPIs.** A *leading* KPI measures a process that drives later results, so it warns you early:
  long door-to-provider times come before patients leave without being seen and before poor survey scores. A *lagging*
  KPI reports an outcome after the fact: a month's readmission rate isn't final until 30 days after the month ends.
  Leaders need both.
- **Sketch the layout on paper first.** Draw boxes for the selectors, cards, and charts. Moving a box on paper is much
  cheaper than moving a chart in Excel.

### 3. Three layers: data, model, and dashboard

Dashboard workbooks that last are built in three layers, each on its own sheet:

```
 DATA                         MODEL (Calc)                       DASHBOARD
 tblKPI, tblEDWaits,   ──►    selected month, last year's  ──►   selectors, KPI cards,
 tblTargets, Lists            month, criteria, chart blocks      charts, labels
```

| Layer | What lives there | Rules |
|---|---|---|
| **Data** | Tables, one fact per row | No formulas that depend on the dashboard. New months are appended at the bottom |
| **Model** | Helper cells and calculations that more than one card needs, plus the blocks that charts plot | Each calculation exists once. Hide the sheet when you're done |
| **Dashboard** | Selectors, cards, charts, and text | Cells hold short formulas that point at the model or the data. No typed-in numbers |

This separation pays off every month. When new data arrives, you append rows to `tblKPI`, and the dashboard updates itself
because it only reads the tables and the selectors. (Add the new month to the Month list on **Lists** too, so the dropdown
offers it.) Open the **Calc** sheet. Its gray cells already turn the selectors into
the pieces that formulas need:

| Calc cell | Formula | Shows (default selection) |
|---|---|---|
| C6 | `=SelMonth` | Nov 2025 |
| C7 | `=EDATE(SelMonth,-1)` | Oct 2025 (the previous month) |
| C8 | `=EDATE(SelMonth,-12)` | Nov 2024 (the same month last year) |
| C9 | `=IF(SelFacility="All facilities","*",SelFacility)` | Cedar Ridge Medical Center |

### 4. Additive and non-additive measures

Open **KPI_Monthly**. Each row is one hospital in one month, and the columns come in two kinds.

An **additive measure** can be summed across rows and still mean something: in January 2025, 473 ED visits plus 79 plus 79
is 631 ED visits for the system. A **non-additive measure**, such as a rate, an average, or a median, can't. Adding three LWBS rates gives a
meaningless number.

| Columns in `tblKPI` | Meaning | Additive? |
|---|---|---|
| Month, Facility | The keys. Month is a real date (the first of the month) formatted as *mmm yyyy* | — |
| EDVisits, LWBS | ED arrivals, and arrivals who left without being seen | Yes |
| LWBSRate | LWBS ÷ EDVisits | No |
| MedianDTP | Median minutes from arrival to first provider contact | No, and it can't be rebuilt from this table |
| IPDischarges, LOSDays | Inpatient discharges and their total length of stay in days | Yes |
| ALOS | LOSDays ÷ IPDischarges | No |
| IndexStays, Readmits | Discharges that count for readmission, and those readmitted within 30 days | Yes |
| ReadmitRate | Readmits ÷ IndexStays | No |
| PatientDays, BedDays | Midnight census and staffed beds, summed over the days of the month | Yes |
| Occupancy | PatientDays ÷ BedDays | No |
| Surveys, TopBox | Surveys returned, and those rating the stay 9 or 10 | Yes |
| TopBoxRate | TopBox ÷ Surveys | No |
| ClaimsAdjudicated, ClaimsDenied | Claims with a payer decision, and those denied or appealed | Yes |
| DenialRate | ClaimsDenied ÷ ClaimsAdjudicated | No |

Here's why the difference matters. In January 2025 the three hospitals reported these LWBS numbers:

| Hospital | LWBS | ED visits | LWBS % |
|---|--:|--:|--:|
| Bluestone Memorial Hospital | 6 | 473 | 1.27% |
| Ashby Falls Community Hospital | 2 | 79 | 2.53% |
| Cedar Ridge Medical Center | 3 | 79 | 3.80% |
| **System (total ÷ total)** | **11** | **631** | **1.74%** |
| Average of the three rates | | | 2.53% |

The average of the rates is off by almost a full point, because it gives two small hospitals the same weight as one that
sees six times as many patients. **Rates roll up as total numerator ÷ total denominator.** That's why `tblKPI` stores the
components next to every rate. A dashboard that has to show *All facilities* or a whole quarter rebuilds each rate from them.

**Medians are worse: you can't rebuild them at all.** In January 2025 the three hospitals' median door-to-provider times
were 41.0, 46.0, and 43.5 minutes. Their average is 43.5, but the true median of all 620 patients seen that month is 42.0.
A median depends on every individual value, so a roll-up median has to come from the detail rows in `tblEDWaits`
(section 6 shows how).

> ⚠️ **Watch for immature data.** Some numbers aren't final when the month closes. In `tblKPI`, December 2025's readmission
> columns are blank because the 30-day window is still open, its survey columns are blank because most surveys haven't come
> back, and its claims columns are blank because most claims are still pending. A blank makes a rate divide by zero, so cards
> should show *n/a* (with `IFERROR`) rather than 0%, which would look like a perfect month.

### 5. Selectors: dropdowns, names, and combo boxes

Excel offers several kinds of selector. [Lesson 3.2](../../03-data-analysis/02-data-validation-conditional-formatting/README.md)
taught dropdowns, and [Lesson 3.4](../../03-data-analysis/04-pivottables/README.md) taught slicers and timelines. This table
compares them as dashboard controls:

| Selector | How you add it | What formulas can read | Strengths | Watch out for |
|---|---|---|---|---|
| Data-validation dropdown | **Data → Data Validation → List** | The cell's value (text or a date) | Simple, and works in every version of Excel, including the web | Someone can paste over it, so protect the sheet (section 11) |
| Form Controls combo box | **Developer → Insert → Combo Box** | Its *cell link*: the chosen item's position number | Floats above the grid, so it can't be typed over | Returns a number, not text. Doesn't run in Excel for the web |
| Slicer on a PivotTable | **PivotTable Analyze → Insert Slicer** | Nothing directly. It filters pivots | Buttons show what's selected, and one slicer can filter many pivots | Pivots need refreshing. Formulas such as SUMIFS don't see the selection |
| Timeline | **PivotTable Analyze → Insert Timeline** | Nothing directly. It filters pivots by date | Click or drag across months, quarters, or years | Works only with PivotTables and date fields |
| Slicer on a Table | **Table Design → Insert Slicer** | Nothing directly. It hides rows | Fast filtering for people browsing the table | SUMIFS still counts the hidden rows. Only SUBTOTAL and AGGREGATE skip them |

For formula-driven cards, use dropdowns or combo boxes. Use slicers and timelines with PivotTables (section 9).

**Build a dropdown selector.** Dashboard!C4 and C5 are already set up this way. You'll repeat these steps for the bonus
dashboard.

1. Put the choices in a list on a helper sheet. On **Lists**, A2:A5 holds *All facilities* and the three hospital names, and
   C2:C25 holds 24 month-start dates formatted as *mmm yyyy*.
2. Select the selector cell (Dashboard!C4) and choose **Data → Data Validation** (Windows: Alt, A, V, V).
3. Under **Allow**, choose **List**. In **Source**, type `=Lists!$A$2:$A$5`. Leave **In-cell dropdown** ticked.
4. On the **Error Alert** tab, keep the *Stop* style so nobody can type a facility that doesn't exist. Click **OK**.

Select the cell and press **Alt + ↓** (Mac: **Option + ↓**) to open the list from the keyboard.

> 💡 **Tip:** Store months as real dates, not text like "Nov 2025". A date list formatted as *mmm yyyy* shows friendly labels
> in the dropdown but puts a real date in the cell, so `EDATE` and `SUMIFS` can work with it.

> ⚠️ A data-validation **Source** can't be a Table reference such as `=tblKPI[Facility]`. Point it at a plain range, at a
> defined name that refers to the column, or (in Microsoft 365) at a spill: put `=SORT(UNIQUE(tblKPI[Facility]))` in a cell
> such as Lists!G2 and use `=Lists!$G$2#` as the Source.

**Name the selector cells.** Click Dashboard!C4, click the **Name Box** (left of the formula bar), type `SelFacility`, and
press **Enter**. The workbook already names C4 `SelFacility` and C5 `SelMonth`. Names make formulas read like sentences
(`tblKPI[Month],SelMonth` instead of `tblKPI[Month],Dashboard!$C$5`), and they keep working if you move the selector. See
every name with **Formulas → Name Manager** (Windows: Ctrl + F3).

**Add a Form Controls combo box.** Form Controls live on the **Developer** tab, which is hidden by default:

- **Windows:** **File → Options → Customize Ribbon**, tick **Developer** in the right-hand list, then **OK**.
- **Mac:** **Excel → Preferences** (or **Settings**) **→ Ribbon & Toolbar**, tick **Developer** under Main Tabs, then **Save**.

Then:

1. Choose **Developer → Insert → Combo Box (Form Control)** (Mac: **Developer → Combo Box**) and drag a rectangle on the
   sheet.
2. Right-click the combo box → **Format Control** → **Control** tab.
3. Set the **Input range** to the list (`Lists!$A$2:$A$5`), the **Cell link** to the cell that receives the choice, and
   **Drop down lines** to the number of items to show at once. Click **OK**.
4. Click any cell to deselect the control, then use it.

The cell link receives the **position** of the chosen item: 1 for the first item, 2 for the second, and so on. Turn the
position back into text with `INDEX`, as Calc!C13 does:

```
=INDEX(Lists!$A$2:$A$5, Calc!C12)
```

> 📋 Form Controls are a desktop feature. Excel for the web doesn't support them, so use data-validation dropdowns if your
> dashboard will be opened in a browser. Avoid **ActiveX** controls (the other half of the Insert menu), because they're
> Windows-only and often blocked by security settings.

### 6. Selector-driven formulas

Every card formula follows the same idea: filter the data by the selectors, then aggregate. These patterns cover almost
everything a dashboard needs:

| Need | Pattern |
|---|---|
| One additive value | `=SUMIFS(tblKPI[EDVisits],tblKPI[Facility],SelFacility,tblKPI[Month],SelMonth)` |
| A rate, rebuilt from components | `=SUMIFS(numerator…)/SUMIFS(denominator…)` with the same criteria |
| A non-additive value for one row | `=XLOOKUP(1,(tblKPI[Facility]=SelFacility)*(tblKPI[Month]=SelMonth),tblKPI[MedianDTP])` |
| The same month last year | Replace `SelMonth` with `EDATE(SelMonth,-12)` |
| The previous month | `EDATE(SelMonth,-1)` |
| A window of months | Two criteria on Month: `">="&EDATE(SelMonth,-2)` and `"<="&SelMonth` |
| Year to date | `">="&DATE(YEAR(SelMonth),1,1)` and `"<="&SelMonth` |
| "All facilities" | `IF(SelFacility="All facilities","*",SelFacility)` as the Facility criterion |
| An average over detail rows | `=AVERAGEIFS(tblEDWaits[DoorToProviderMin],tblEDWaits[Facility],SelFacility, …date window…)` |
| A median across hospitals or months | `=MEDIAN(FILTER(tblEDWaits[DoorToProviderMin], conditions))` |
| "All facilities" inside FILTER | `((tblEDWaits[Facility]=SelFacility)+(SelFacility="All facilities"))` as the facility condition |
| Missing or immature data | Wrap the formula in `IFERROR(…,"n/a")` |

**SUMIFS is the workhorse.** With both selectors as criteria it returns exactly one hospital-month, because `tblKPI` has
one row for each. With a criterion dropped, or with a date window, it adds up several rows correctly. It never returns
`#N/A`, and it recalculates the moment a selector changes. Here are worked examples with SelFacility set to Bluestone
Memorial Hospital and SelMonth set to Jan 2025:

| Card value | Formula | Result |
|---|---|--:|
| ED visits | `=SUMIFS(tblKPI[EDVisits],tblKPI[Facility],SelFacility,tblKPI[Month],SelMonth)` | 473 |
| ED visits, same month last year | `…,tblKPI[Month],EDATE(SelMonth,-12))` | 405 |
| Change vs last year | this year ÷ last year − 1 | +16.8% |
| Rolling 3 months (Nov 2024 – Jan 2025) | `…,tblKPI[Month],">="&EDATE(SelMonth,-2),tblKPI[Month],"<="&SelMonth)` | 1,321 |
| Year to date, with SelMonth set to Mar 2025 | `…,tblKPI[Month],">="&DATE(YEAR(SelMonth),1,1),tblKPI[Month],"<="&SelMonth)` | 1,266 |

**Why EDATE?** `EDATE(date, months)` moves a date by whole months and keeps the day of the month, so the first of November
becomes the first of another month and matches the Month column exactly. Subtracting 365 days lands on the wrong day
whenever a February 29 is in between. Joining an operator to a date, as in `">="&EDATE(SelMonth,-2)`, turns the date into
criteria text that SUMIFS understands.

**Looking up a non-additive value.** For one hospital-month you can read MedianDTP directly. `XLOOKUP(1, (A=x)*(B=y), C)`
finds the row where both conditions are TRUE, because TRUE × TRUE = 1
([Lesson 4.2](../02-advanced-formulas-let-lambda/README.md)). Don't SUMIFS a rate or median column. It happens to work for
one row, but it silently adds the values together as soon as the selection covers several rows.

**AVERAGEIFS belongs on detail rows.** `AVERAGEIFS` takes the same criteria as SUMIFS and returns the mean of the matching
cells, so it's the right tool for a mean wait over the visits in `tblEDWaits`. Pointed at a rate column it goes wrong:
`=AVERAGEIFS(tblKPI[LWBSRate],tblKPI[Month],SelMonth)` gives every hospital equal weight and returns 2.53% for January 2025
instead of the true 1.74% (section 4).

**The "All facilities" trick.** No row in `tblKPI` says *All facilities*, so `SUMIFS(…,tblKPI[Facility],SelFacility,…)`
returns 0 when it's selected. SUMIFS criteria accept wildcards, and `"*"` matches any text, so swapping the selector for
`IF(SelFacility="All facilities","*",SelFacility)` makes the same formula add up all three hospitals. Calc!C9 holds exactly
that helper, so formulas that point at it handle both cases.

**Medians from the detail.** `FILTER` returns the matching detail rows and `MEDIAN` reduces them to one number, so the
formula fits in a single cell without spilling:

```
=MEDIAN(FILTER(tblEDWaits[DoorToProviderMin],
               (tblEDWaits[ArrivalDateTime]>=SelMonth)*(tblEDWaits[ArrivalDateTime]<EDATE(SelMonth,1))))
```

The date test is a **half-open window**: on or after the first day of the month, and before the first day of the next
month. ArrivalDateTime includes a time, so a test of `<=` the last day of the month would miss every arrival after midnight
on that day. To limit the median to one hospital as well, multiply in a third condition: `*(tblEDWaits[Facility]=SelFacility)`.

**"All facilities" inside FILTER.** FILTER compares values exactly and doesn't understand wildcards, so the `"*"` trick
from SUMIFS matches no rows here, and FILTER returns `#CALC!` because it has nothing to return. Write the facility test
with OR logic instead ([Lesson 4.1](../01-dynamic-arrays/README.md)). A row should pass if its facility matches the
selector, or if the selector says *All facilities*. Adding the two tests does exactly that, because TRUE + FALSE = 1 and
FALSE + FALSE = 0:

```
=MEDIAN(FILTER(tblEDWaits[DoorToProviderMin],
               (tblEDWaits[ArrivalDateTime]>=SelMonth)*(tblEDWaits[ArrivalDateTime]<EDATE(SelMonth,1))
               *((tblEDWaits[Facility]=SelFacility)+(SelFacility="All facilities"))))
```

With a hospital selected, only that hospital's visits pass. With *All facilities* selected, the second test is TRUE for
every row, so every visit in the month passes. The bonus needs this version, because its median card has to work both
ways.

> 📋 **Older versions.** FILTER, XLOOKUP, and LET need Microsoft 365 or Excel 2021 or later. In Excel 2010–2019, use
> `=AGGREGATE(17,6,values/(conditions),2)` for a conditional median (function 17 is QUARTILE.INC, quartile 2 is the median,
> and option 6 skips the `#DIV/0!` errors that the division creates for rows that don't match), and INDEX/MATCH for lookups.
> A multi-condition lookup such as `=INDEX(tblKPI[MedianDTP],MATCH(1,(tblKPI[Facility]=SelFacility)*(tblKPI[Month]=SelMonth),0))`
> must be confirmed with **Ctrl + Shift + Enter** (Mac: **⌘ + Shift + Return**) in those versions.

**Name the pieces with LET.** Card formulas get long because the same SUMIFS appears more than once. `LET` names each piece
once and then uses the names:

```
=LET(cur,  SUMIFS(tblKPI[PatientDays],tblKPI[Facility],SelFacility,tblKPI[Month],SelMonth)
          /SUMIFS(tblKPI[BedDays],tblKPI[Facility],SelFacility,tblKPI[Month],SelMonth),
     prev, …the same with EDATE(SelMonth,-1)…,
     IF(cur>prev, "up", "down"))
```

> 💡 **Tip:** Break a long formula over several lines in the formula bar with **Alt + Enter** (Mac: **⌃ + Option +
> Return**). Excel ignores the line breaks when it calculates.

### 7. KPI cards: value, context, status, and trend

A **KPI card** is a small block of cells that shows one KPI. A number alone tells the reader nothing ("Is 1.27% good?"), so
every card adds context:

```
┌───────────────────────────────┐
│ LWBS %                        │  title: what the number is
│            1.27%              │  the value, large
│ Target          ≤ 2.0%        │  the target from tblTargets
│ Status          ✔ On target   │  status in text AND color
│ vs last year    ▼ 1.7 pts     │  direction of change
└───────────────────────────────┘
   Bluestone Memorial Hospital, January 2025
```

**Variance** compares the value with the target. For rates there are two kinds, and a dashboard should say which one it
shows:

| Measure | Formula | Bluestone, Jan 2025 | Read it as |
|---|---|---|---|
| Variance in **percentage points** | value − target | 1.27% − 2.00% = −0.73 pts | 0.73 points better than target |
| Relative variance | value ÷ target − 1 | −36.6% | 36.6% below target |
| Change vs last year (a count) | this year ÷ last year − 1 | 473 vs 405 visits = +16.8% | 16.8% more visits |

A **percentage point** is the difference between two percentages. LWBS falling from 2.96% to 1.27% is a drop of 1.7
points, which is a 57% relative improvement. Mixing the two up is one of the most common dashboard errors.

**Status depends on direction.** A higher value is good for HCAHPS top-box % and bad for every other KPI in `tblTargets`.
Read the direction from the KPI dictionary instead of writing a different formula for each card:

```
=IF(IF(direction="Lower is better", value<=target, value>=target), "On target", "Off target")
```

Look `target` and `direction` up by the KPI's name, for example `XLOOKUP("LWBS %",tblTargets[KPI],tblTargets[Direction])`.
The name must match the KPI column exactly. Because every card reads the same dictionary, one edit to `tblTargets` updates
every card that uses it. `LET` keeps the formula readable: name the value, the target, and the direction, then write the IF.

For a **scorecard** (a line such as "6 of 7 KPIs on target"), return a number instead of text. `--IF(…)` turns TRUE and
FALSE into 1 and 0, so `=SUM(statuses)` counts the KPIs on target and `=COUNT(statuses)` counts those with data. A status of
`""` for missing data is skipped by both.

**Trend arrows.** `UNICHAR(number)` returns any Unicode character, so a formula can return an arrow:

| Code | Character | Typical use |
|--:|:-:|---|
| `UNICHAR(9650)` | ▲ | went up |
| `UNICHAR(9660)` | ▼ | went down |
| `UNICHAR(9644)` | ▬ | no change |
| `UNICHAR(10004)` | ✔ | on target |
| `UNICHAR(10008)` | ✘ | off target |

```
=IF(cur>prev, UNICHAR(9650), IF(cur<prev, UNICHAR(9660), UNICHAR(9644)))
```

There's a second way that keeps the cell numeric: a **custom number format** with the arrows inside it. Select the cell,
press **Ctrl + 1** (Mac: **⌘ + 1**), choose **Number → Custom**, and type the format. A custom format has up to four
sections separated by semicolons: positive numbers; negative numbers; zero; text. The arrow formats below use the first
three. The example card on the Dashboard sheet uses the first row of this table:

| Format code | Shows | Use it for |
|---|---|---|
| `"▲ "0.0%;"▼ "0.0%;"▬ "0.0%` | ▲ 3.4% (for 0.034) | Neutral changes, such as volumes |
| `[Color10]"▲ "0.0%;[Red]"▼ "0.0%;"▬ "0.0%` | ▲ 3.4% in green | Changes where up is good |
| `+0.0%;-0.0%;0.0%` | +3.4% | Variances, so the sign always shows |
| `#,##0.0,"K"` | 1.2K (for 1,234) | Large counts. Each trailing comma divides by 1,000 |
| `0.0" min"` | 31.0 min (for 31) | Units inside the number |

The negative section shows the value without a minus sign, so −0.034 appears as ▼ 3.4%. `[Color10]` is a darker, more
readable green than `[Green]`.

**Color the status with conditional formatting**
([Lesson 3.2](../../03-data-analysis/02-data-validation-conditional-formatting/README.md)), not by hand. Select the status cell,
choose **Home → Conditional Formatting → New Rule → Use a formula to determine which cells to format**, and enter a rule
such as `=LEFT($F$10,1)="✔"` with a green fill. Add a second rule for ✘ with a red fill. To color a value cell by
direction instead, use the direction-aware test as the rule: `=IF($R11="Lower is better",$O11<=$Q11,$O11>=$Q11)`. (Both
rules come from the hidden Dashboard Key: F10 is its LWBS status cell, and row 11 of its model area holds the LWBS
value, target, and direction.)

> ⚠️ **Color has to mean the same thing everywhere.** Red means "off target, act". Don't color volumes (more ED visits isn't
> bad), and don't color an arrow red just because it points down, because a falling LWBS % is good news.

> 💡 **Tip:** Icon sets (**Home → Conditional Formatting → Icon Sets**) put arrows or traffic lights next to numbers. Use
> **Edit Rule** to set the thresholds as numbers, tick **Reverse Icon Order** for lower-is-better KPIs, and tick **Show Icon
> Only** if the number is shown elsewhere.

**Build labels with TEXT.** `="Target "&IF(dir="Lower is better","≤ ","≥ ")&TEXT(target,"0.0%")` produces *Target ≤ 2.0%*.
`TEXT` formats a number inside a text string, which a number format can't do.

### 8. Charts that follow the selectors

A chart plots ranges, not formulas, so a **dynamic chart** is an ordinary chart whose ranges contain formulas that read the
selectors. When a selector changes, the formulas recalculate and the chart redraws.

**Build the chart block on the model sheet:**

1. Choose a block of rows for the window, such as 12 rows for "the last 12 months".
2. In the month column, return the oldest month in the first row, then step forward one month per row. A counter does this
   with one formula copied down: `ROWS(B$17:B17)` returns 1 in the first row, 2 in the next, and so on, because only the
   second reference moves as you copy ([Lesson 1.5](../../01-foundations/05-cell-references/README.md)). So
   `=EDATE(SelMonth,ROWS(B$17:B17)-12)` runs from 11 months back to SelMonth itself.
3. In the value column, use the selector-driven SUMIFS with that row's month as the Month criterion.

In Microsoft 365, `=EDATE(SelMonth,SEQUENCE(12,1,-11))` spills all 12 months from one cell
([Lesson 4.1](../01-dynamic-arrays/README.md)). A chart built from a spill keeps a fixed range, so if a spill can change
size, chart it through a defined name that refers to the spill, such as `=Calc!$B$17#`.

**Insert and place the chart:**

1. Select the block including its header row and choose **Insert → Charts → Line**. (On Windows, **Alt + F1** inserts a
   default chart right away.)
2. Leave the top-left header cell blank. With an empty corner, Excel treats the first column as the axis labels. If Excel
   still plots the months as a second line, right-click the chart → **Select Data** and fix the axis labels there.
3. Cut the chart (**Ctrl + X**, Mac: **⌘ + X**) and paste it on the dashboard. A moved chart still points at the model sheet.

**Make the title dynamic**, the same way as in [Lesson 3.5](../../03-data-analysis/05-charts-visualization/README.md).
Build the title text in a cell, for example
`="ED visits, 12 months to "&TEXT(SelMonth,"mmm yyyy")&" · "&SelFacility`. Then click the chart title, type `=` in the
formula bar, click that cell, and press **Enter**. Now the title follows the selectors too. Both charts on the hidden
Dashboard Key work this way: their titles point at cells O19 and O20 of its model area.

**Handle gaps.** If the window reaches back before the data starts, SUMIFS returns 0 and the line plunges to zero. Return
`NA()` instead: `=IF(COUNTIFS(tblKPI[Facility],SelFacility,tblKPI[Month],B17)=0,NA(),SUMIFS(…))`. Excel doesn't plot `#N/A`
points, so the line simply starts later.

**Add a target line.** Add a column to the block that repeats the target in every row and include it in the chart as a
second series. The Dashboard Key's LWBS chart does this with `=$Q$11`, its LWBS target cell. Format the target series as a
thin dashed line, and readers see the misses without reading a number.

**Choose the right chart for each question:**

| The question | Chart |
|---|---|
| How is it trending? | Line (or columns for counts) |
| How do hospitals or units compare? | Bar, sorted, with the axis starting at zero |
| Is it above the target? | Line or bar with a target line |
| What share is each part? | A sorted bar. Avoid pies with more than two or three slices |
| A single number | A KPI card, not a chart |

**Leave out chartjunk:** 3-D effects, gradients, heavy gridlines, a legend for a single series, and axis labels with more
decimals than the data deserves. Every mark on the chart should help answer the question.

**Sparklines** ([Lesson 3.5](../../03-data-analysis/05-charts-visualization/README.md)) are tiny charts inside one cell,
ideal next to a KPI card. Select the destination cell, choose **Insert → Sparklines → Line**, set the **Data Range** to the
12 values, and click **OK**. Use **Sparkline → Show → High Point** or **Markers** to emphasize the points that matter.

### 9. Slicers and timelines across several PivotTables

PivotTables ([Lesson 3.4](../../03-data-analysis/04-pivottables/README.md)) are a fast way to put summaries on a
dashboard, and slicers make them interactive. This table helps you choose between formulas and pivots:

| | Formula cards (SUMIFS, XLOOKUP) | PivotTables with slicers |
|---|---|---|
| Layout | Exactly where you put each number | The pivot decides, and it grows or shrinks as filters change |
| When a selector changes | Updates immediately | Updates immediately (slicers and timelines) |
| When the data changes | Updates immediately | Only after a refresh |
| Building | Slower, one formula per number | Fast, drag and drop |
| Rates and medians | Any definition you can write | Calculated fields sum first, then divide. No Median summary (except through the Data Model and DAX) |
| Drill-down | No | Double-click any value |

**Connect one slicer to several pivots.** A slicer starts out filtering only the pivot you inserted it from.

1. Build each pivot from the same source, `tblKPI`. Pivots built from the same Table share a **pivot cache** (Excel's stored
   copy of the source data), and only pivots that share a cache can share a slicer.
2. Click a pivot → **PivotTable Analyze → Insert Slicer** → tick the field (such as **Facility**) → **OK**.
3. Select the slicer and choose **Slicer → Report Connections**. Tick every pivot the slicer should filter → **OK**.
4. A **timeline** works the same way: **PivotTable Analyze → Insert Timeline**, then **Timeline → Report Connections**. Use
   the time-level menu at the timeline's top right to switch between YEARS, QUARTERS, MONTHS, and DAYS.

You can also connect from the pivot's side: click in a pivot → **PivotTable Analyze → Filter Connections**.

> ⚠️ If a pivot is missing from the Report Connections list, it was built on a different cache, for example one pivot added
> to the Data Model and one not. Rebuild it from the same Table, the same way as the others.

**Rates in a pivot.** A calculated field (**PivotTable Analyze → Fields, Items & Sets → Calculated Field**) such as
`=LWBS/EDVisits` sums each field over the visible rows first, then divides. Lesson 3.4 warned that this breaks some
calculations, but for a rate built from additive components it's exactly the total ÷ total rule from section 4.

**Make pivots dashboard-friendly:**

- Put the pivots on a helper sheet and show only slicers, PivotCharts, or formula cards on the dashboard. A slicer can sit on
  a different sheet from the pivots it filters.
- Right-click a slicer → **Slicer Settings** → tick **Hide items with no data**. Use **Slicer → Columns** to lay the buttons
  out in a row.
- In **PivotTable Options → Layout & Format**, untick **Autofit column widths on update**, so filtering doesn't resize your
  carefully set columns.
- Feed a formula card from a pivot with `GETPIVOTDATA` (Lesson 3.4) when you need the pivot's number in a fixed place.
- Pivots don't recalculate on their own. Use **Data → Refresh All** (Windows: **Ctrl + Alt + F5**) after the data changes,
  or tick **Refresh data when opening the file** in **PivotTable Options → Data**.

### 10. Layout and visual design

People read a dashboard the way they read a page: top-left first, then across, then down. Put the most important thing
(often the scorecard, or the KPI that's furthest off target) at the top left, and the detail at the bottom.

- **Use the grid.** Give columns consistent widths and make every card the same size. On Windows, hold **Alt** while you drag
  or resize a chart to snap it to the cell edges. Select several shapes or charts and use **Align** on the **Shape Format**
  or **Chart Format** tab to line them up.
- **Use white space** instead of borders to separate groups. Narrow spacer columns between cards do this well.
- **Use color with meaning.** Keep the page neutral (grays plus one brand color) and save red and green for status. If
  everything is colorful, the red cell no longer stands out.
- **Never rely on color alone.** About 1 in 12 men has some color-vision deficiency, most often red-green. Pair color with a
  symbol or a word (✔ On target, ✘ Off target), and check the page in grayscale.
- **Format numbers for reading.** Use the fewest decimals that still separate the values (1.4%, not 1.36986%), put units in
  labels or formats, and use the same format for the same KPI everywhere.
- **Write titles that say what the reader is looking at.** *ED visits, 12 months to Nov 2025 · Cedar Ridge Medical Center*
  beats *Chart 1*.

### 11. Polish: hide, protect, print, and share

**Make it look like a page, not a spreadsheet.**

- On the dashboard sheet, untick **View → Gridlines** and **View → Headings** (both platforms).
- Hide the model and data sheets (right-click a tab → **Hide**) once the dashboard works. To stop readers unhiding them, use
  **Review → Protect Workbook** and tick **Structure**.
- A **linked picture** is a live image of a range that you can place anywhere, which is handy for showing a block from the
  model sheet inside the dashboard layout. Copy the range, then choose **Home → Paste ▾ → Linked Picture** (both platforms).
  On Windows, the **Camera** command does the same thing. Add it through **File → Options → Quick Access Toolbar →
  Commands Not in the Ribbon**.

**Protect the sheet so only the selectors can change.** Every cell starts out *locked*, but locking does nothing until the
sheet is protected.

1. Select the selector cells, press **Ctrl + 1** (Mac: **⌘ + 1**), and on the **Protection** tab untick **Locked** → **OK**.
2. For each slicer, right-click → **Size and Properties** → **Properties** → untick **Locked**.
3. Choose **Review → Protect Sheet**. Keep *Select locked cells* and *Select unlocked cells* ticked, tick **Use PivotTable &
   PivotChart** if the sheet has slicers, and add a password only if you'll store it somewhere safe.

> ⚠️ A combo box can write to its cell link only if that cell is unlocked or on an unprotected sheet. Keep cell links on the
> model sheet.

**Print it on one page.**

1. Select the dashboard area and choose **Page Layout → Print Area → Set Print Area**.
2. Choose **Page Layout → Orientation → Landscape**.
3. In **Page Layout → Scale to Fit**, set Width to **1 page** and Height to **1 page**.
4. Check the result with **File → Print** (**Ctrl + P**, Mac: **⌘ + P**).

**Make it accessible.**

- Right-click each chart → **Edit Alt Text** and describe the message ("LWBS % stayed below the 2% target in 10 of the last
  12 months"), not the chart type.
- Run **Review → Check Accessibility** and fix what it reports.
- Keep text contrast high (dark text on light fills) and avoid text smaller than 9 points.
- Give charts and shapes meaningful names in the **Selection Pane** (Windows: **Home → Find & Select → Selection Pane**), so
  screen-reader users hear "LWBS trend chart" instead of "Chart 3".

### 12. Performance

A dashboard should respond instantly when someone changes a selector. Most slow dashboards have the same cause: volatile
functions.

A **volatile function** recalculates after every change anywhere in the workbook, even when its inputs didn't change, and so
does every formula that depends on it.

| Volatile | Non-volatile alternative |
|---|---|
| `OFFSET` | `INDEX`, or SUMIFS with a date window |
| `INDIRECT` | `INDEX` or `XLOOKUP`, or `CHOOSE` for a few fixed choices |
| `TODAY`, `NOW` | A *report date* cell that you update once per refresh |
| `RAND`, `RANDBETWEEN`, `RANDARRAY` | Generate the numbers once, then paste them as values |
| `CELL`, `INFO` | Avoid them on dashboards |

A common older pattern sums the last three rows of a chart block with OFFSET. Here are three ways to write a rolling
3-month total over the Calc block, from worst to best:

| Formula | Volatile? |
|---|---|
| `=SUM(OFFSET(Calc!C28,-2,0,3,1))` | Yes |
| `=SUM(INDEX(Calc!C17:C28,10):INDEX(Calc!C17:C28,12))` | No. INDEX can return a cell reference, so INDEX:INDEX builds a range |
| `=SUMIFS(tblKPI[EDVisits],tblKPI[Facility],SelFacility,tblKPI[Month],">="&EDATE(SelMonth,-2),tblKPI[Month],"<="&SelMonth)` | No, and it reads the data directly |

For Bluestone Memorial Hospital at Jan 2025, all three return 1,321 (413 + 435 + 473). The SUMIFS version is the most robust
because it doesn't depend on where the block is or how the data is sorted.

More ways to keep a dashboard fast:

- **Calculate once, reference many times.** Put shared pieces (the selected month, last year's month, the facility
  criterion) in model cells, and point the cards at them.
- **Point formulas at Tables, not whole columns.** `FILTER(A:A, …)` processes a million rows. `tblEDWaits[…]` processes only
  the rows that exist, and it grows with the data.
- **Use FILTER sparingly over big detail tables.** One MEDIAN(FILTER(…)) over 12,000 rows is instant, but hundreds of them add
  up. Precompute what you can in the data layer, the way `tblKPI` does.
- **Check the calculation mode.** **Formulas → Calculation Options** should be **Automatic**. If someone set it to Manual, the
  dashboard stops updating until they press **F9** (Mac: **⌘ + =**).

### 13. Worked example: the Inpatient discharges card, start to finish

Open the **Dashboard** sheet. The card under the selectors is finished, so you can click each cell and read its formula. It
was built in four steps:

1. **Title.** B7:C7 is merged and filled navy: *INPATIENT DISCHARGES · example card*.
2. **Value.** B8:C8 is merged, set in large bold type, and holds
   `=SUMIFS(tblKPI[IPDischarges],tblKPI[Facility],SelFacility,tblKPI[Month],SelMonth)`. For Cedar Ridge Medical Center in
   Nov 2025 it shows **30**.
3. **Context.** C9 holds the same SUMIFS with `EDATE(SelMonth,-12)` as the month: **29** discharges in Nov 2024.
4. **Trend.** C10 holds `=IFERROR(B8/C9-1,"n/a")` with the custom format `"▲ "0.0%;"▼ "0.0%;"▬ "0.0%`, so it shows
   **▲ 3.4%** while the cell still holds a number (0.0345).

> 📋 [Lesson 1.3](../../01-foundations/03-formatting-cells/README.md) advised against merged cells because they break
> sorting, filtering, and copying. Nobody sorts or filters a dashboard's display cells, so merging a card's title and big
> number is a common exception. **Center Across Selection** gives the same look without merging.

The card has no target and no color, because discharges are a volume: more isn't better or worse. Now change the Facility
selector to Bluestone Memorial Hospital and the Month to Jan 2025. All three numbers change, and the arrow may flip, without
anyone touching a formula. That's the whole idea of an interactive dashboard. Set the selectors back to Cedar Ridge Medical
Center and Nov 2025 before you start the practice tasks, because they're checked against that selection.

### 14. Shortcuts and version notes

| Action | Windows | Mac |
|---|---|---|
| Open a dropdown list | Alt + ↓ | Option + ↓ |
| Format Cells (custom formats, Locked) | Ctrl + 1 | ⌘ + 1 |
| Name Manager | Ctrl + F3 | **Formulas → Name Manager** |
| Insert a default chart next to the data | Alt + F1 | **Insert → Charts** |
| Line break inside a formula | Alt + Enter | ⌃ + Option + Return |
| Refresh all PivotTables | Ctrl + Alt + F5 | **Data → Refresh All** |
| Recalculate the workbook | F9 | ⌘ + = |
| Select several slicer buttons | Ctrl + click | ⌘ + click |
| Snap a chart to the cell grid | Hold Alt while dragging | Use **Align** on the Format tab |
| Print preview | Ctrl + P | ⌘ + P |

| Feature | Availability |
|---|---|
| SUMIFS, EDATE, data-validation lists, conditional formatting | Every current version, including Excel for the web |
| UNICHAR | Excel 2013 and later |
| XLOOKUP, LET, FILTER, SEQUENCE | Microsoft 365 and Excel 2021 or later |
| Slicers for PivotTables | Excel 2010 and later (Windows), current Mac versions, Excel for the web |
| Timelines | Excel 2013 and later (Windows), Excel 2016 and later (Mac) |
| Form Controls | Excel for Windows and Mac. They don't run in Excel for the web |
| Sparklines | Excel 2010 and later (Windows), Excel 2011 and later (Mac) |
| Linked Picture | Windows and Mac. The Camera command is Windows only |

## 🧪 Hands-on practice

Download [`4.6-dashboards.xlsx`](4.6-dashboards.xlsx) and open the **Practice** sheet. Tasks 1–9 are formulas you type in the
yellow cells. Tasks 10–13 have you build a combo box, a chart block, and connected PivotTables. The **Check** column turns
green when you're right.

<!-- BEGIN GENERATED: practice -->
The yellow selectors on the Dashboard sheet are named SelFacility (Dashboard!C4) and SelMonth (Dashboard!C5). They start at Cedar Ridge Medical Center and Nov 2025. Write every formula with SelFacility and SelMonth, never typed-in names or dates, and keep that selection while you check your answers. (Change it afterwards and watch your formulas follow.) Tasks 1–9 are formulas you type in the yellow cells below. Tasks 10 and 11 are built on the Dashboard and Calc sheets, and their gray cells fill in from your work. For tasks 12 and 13, build PivotTables on a new Pivots sheet and type the number they show.

| # | Task | Hint |
|:-:|------|------|
| 1 | Write a formula that returns the number of ED visits for the selected facility and month (the EDVisits column of tblKPI). | SUMIFS with two criteria: Facility = SelFacility and Month = SelMonth |
| 2 | Return the LWBS % (left without being seen) for the selection. Build it from its components, LWBS ÷ EDVisits, instead of reading the LWBSRate column. Enter it as a percentage. | One SUMIFS for the numerator divided by one SUMIFS for the denominator |
| 3 | System-wide LWBS % for the selected month: all three hospitals combined, ignoring the facility selector. Divide total LWBS by total ED visits instead of averaging the three hospitals' rates. Enter it as a percentage. | Drop the Facility criterion from both SUMIFS |
| 4 | LWBS variance to target: the selected LWBS % minus the LWBS % target from tblTargets. Look the target up with a formula instead of typing 2%. Enter the result as a percentage. A positive result means the hospital is above (worse than) this lower-is-better target. | XLOOKUP("LWBS %", tblTargets[KPI], tblTargets[Target]) |
| 5 | ED visits change versus the same month last year: this year ÷ last year − 1. Find last year's month with EDATE so the formula works for any selected month. Enter it as a percentage. | EDATE(SelMonth,-12) is the same month one year earlier |
| 6 | Rolling 3-month ED visits for the selected facility: the selected month plus the two months before it (Sep–Nov 2025 for the default selection). Use one SUMIFS with a date window, not OFFSET. | Two criteria on the Month column: ">="&EDATE(SelMonth,-2) and "<="&SelMonth |
| 7 | Occupancy trend arrow. Compare the selected month's occupancy (PatientDays ÷ BedDays) with the previous month's. Return ▲ with UNICHAR(9650) if it rose, ▼ with UNICHAR(9660) if it fell, or ▬ with UNICHAR(9644) if it didn't change. | LET(cur, …, prev, …, IF(cur>prev, UNICHAR(9650), …)). EDATE(SelMonth,-1) is the previous month |
| 8 | Status cell for median door-to-provider: return the text On target or Off target for the selection. Read the target and the Direction of "Median door-to-provider" from tblTargets, so the same pattern works for any KPI. A median isn't additive, so look up the hospital's MedianDTP value instead of adding it up. | LET + XLOOKUP(1, (tblKPI[Facility]=SelFacility)*(tblKPI[Month]=SelMonth), tblKPI[MedianDTP]). Then IF on the direction |
| 9 | System-wide median door-to-provider, in minutes, for the selected month. Medians can't be added or averaged across hospitals, so compute it from the visit-level table: the MEDIAN of DoorToProviderMin for every tblEDWaits visit whose ArrivalDateTime falls in the selected month. Keep one decimal place. | MEDIAN(FILTER(…)) with ArrivalDateTime >= SelMonth and < EDATE(SelMonth,1) |
| 10 | Insert a Form Controls combo box on the Dashboard. In Format Control, set its Input range to Lists!$A$2:$A$5 and its Cell link to Calc!$C$12. Then choose Ashby Falls Community Hospital in the combo box. The gray cell shows the number your combo box writes to the cell link. | Developer → Insert → Combo Box (Form Control), then right-click it → Format Control → Control tab |
| 11 | On the Calc sheet, fill the yellow 12-month trend block. In B17:B28, return the 12 months ending at SelMonth, oldest first (use EDATE). In C17:C28, return ED visits for SelFacility in each of those months. Then select B16:C28, insert a line chart, and move it to the Dashboard. The gray cell totals your block: the trailing-12-month ED visits. | EDATE(SelMonth,-11) is the oldest month. A counter such as ROWS(B$17:B17) lets one formula step forward as you copy it down |
| 12 | Insert a new sheet named Pivots. Build PivotTable 1 from tblKPI with Month in Rows and Sum of EDVisits in Values. Insert a Facility slicer and a Month timeline for it. Select Ashby Falls Community Hospital in the slicer, and select 2025 Q3 in the timeline (switch it to QUARTERS). What is PivotTable 1's grand total? | PivotTable Analyze → Insert Slicer, and PivotTable Analyze → Insert Timeline |
| 13 | On the Pivots sheet, build PivotTable 2 from tblKPI with Facility in Rows, plus Sum of LWBS and Sum of EDVisits in Values. Connect the slicer and the timeline to it (Report Connections). Then add a calculated field LWBSPct = LWBS / EDVisits. With Ashby Falls and 2025 Q3 still selected, what LWBS % does PivotTable 2 show? Enter it as a percentage. | Select the slicer → Slicer → Report Connections. Then PivotTable Analyze → Fields, Items & Sets → Calculated Field |
<!-- END GENERATED: practice -->

## ✅ Answer key

The workbook has a hidden **Answer Key** sheet (right-click any sheet tab → **Unhide…**). Its *Live result* column runs each
sample formula against the data, and the PivotTable tasks show a SUMIFS "formula twin" that proves the number. The same
answers are below, collapsed so you don't see them by accident.

<!-- BEGIN GENERATED: answers -->
<details>
<summary><b>🔑 Show the answer key</b> — Try every task before opening this.</summary>

**1. Write a formula that returns the number of ED visits for the selected facility and…**

- **Answer:** 80
- **Solution:** `=SUMIFS(tblKPI[EDVisits],tblKPI[Facility],SelFacility,tblKPI[Month],SelMonth)`

tblKPI has exactly one row per hospital per month, so SUMIFS with both criteria returns that row's value. SUMIFS is the workhorse of selector-driven dashboards. It never returns #N/A, it adds up correctly when a criterion matches several rows (a quarter, or every hospital), and it recalculates the moment a selector changes. Because the criteria are the named cells, choosing a different facility on the Dashboard updates this number without anyone touching the formula.

**2. Return the LWBS % (left without being seen) for the selection. Build it from its…**

- **Answer:** 3.75%
- **Solution:**

```
=SUMIFS(tblKPI[LWBS],tblKPI[Facility],SelFacility,tblKPI[Month],SelMonth)/SUMIFS(tblKPI[EDVisits],tblKPI[Facility],SelFacility,tblKPI[Month],SelMonth)
```


3 of 80 patients left before a provider saw them. For one hospital-month this matches the LWBSRate column. The components version is still the better habit, because the same formula keeps working when the selection covers several rows (a quarter, or all hospitals). There, SUMIFS on LWBSRate would add percentages together, which is meaningless.

**3. System-wide LWBS % for the selected month: all three hospitals combined, ignoring the…**

- **Answer:** 1.53%
- **Solution:**

```
=SUMIFS(tblKPI[LWBS],tblKPI[Month],SelMonth)/SUMIFS(tblKPI[EDVisits],tblKPI[Month],SelMonth)
```


8 ÷ 522 = 1.53%. Averaging the three hospital rates gives 1.71%, because a simple average gives the small hospitals the same weight as Bluestone Memorial, which sees several times as many patients. Rates roll up as total numerator ÷ total denominator. That's why tblKPI stores the components.

**4. LWBS variance to target: the selected LWBS % minus the LWBS % target from tblTargets.…**

- **Answer:** 1.75%
- **Solution:**

```
=SUMIFS(tblKPI[LWBS],tblKPI[Facility],SelFacility,tblKPI[Month],SelMonth)/SUMIFS(tblKPI[EDVisits],tblKPI[Facility],SelFacility,tblKPI[Month],SelMonth)-XLOOKUP("LWBS %",tblTargets[KPI],tblTargets[Target])
```


3.75% − 2.00% = +1.75%. Strictly, that's +1.75 percentage points: the gap between two rates. (The relative variance, actual ÷ target − 1, would be +87.5%.) Keeping targets in a table and looking them up means one edit to tblTargets updates every card that uses the target. On a real card this cell would subtract the target cell from the value cell.

**5. ED visits change versus the same month last year: this year ÷ last year − 1. Find last…**

- **Answer:** -7.0%
- **Solution:**

```
=SUMIFS(tblKPI[EDVisits],tblKPI[Facility],SelFacility,tblKPI[Month],SelMonth)/SUMIFS(tblKPI[EDVisits],tblKPI[Facility],SelFacility,tblKPI[Month],EDATE(SelMonth,-12))-1
```


80 visits this year against 86 in Nov 2024. Comparing with the same month last year removes seasonality, which a month-over-month change can't do: an ED's November differs from its October for seasonal reasons alone. EDATE moves a date by whole months and keeps it on the first of the month, so it matches the Month column for every selection. Subtracting 365 days happens to work for Nov 2025, but it lands a day late whenever a February 29 falls in between: 1/1/2025 − 365 days is 1/2/2024, which matches no row, while EDATE returns 1/1/2024.

**6. Rolling 3-month ED visits for the selected facility: the selected month plus the two…**

- **Answer:** 221
- **Solution:**

```
=SUMIFS(tblKPI[EDVisits],tblKPI[Facility],SelFacility,tblKPI[Month],">="&EDATE(SelMonth,-2),tblKPI[Month],"<="&SelMonth)
```


Joining an operator to a date (">="&EDATE(SelMonth,-2)) turns the date into criteria text that SUMIFS understands. A common older pattern is SUM(OFFSET(…)). OFFSET is volatile, so Excel recalculates it after every edit anywhere in the workbook, and dozens of them make a dashboard sluggish. SUMIFS recalculates only when its inputs change, and it doesn't depend on the table being sorted.

**7. Occupancy trend arrow. Compare the selected month's occupancy (PatientDays ÷ BedDays)…**

- **Answer:** ▲
- **Solution:**

```
=LET(cur,SUMIFS(tblKPI[PatientDays],tblKPI[Facility],SelFacility,tblKPI[Month],SelMonth)/SUMIFS(tblKPI[BedDays],tblKPI[Facility],SelFacility,tblKPI[Month],SelMonth),prev,SUMIFS(tblKPI[PatientDays],tblKPI[Facility],SelFacility,tblKPI[Month],EDATE(SelMonth,-1))/SUMIFS(tblKPI[BedDays],tblKPI[Facility],SelFacility,tblKPI[Month],EDATE(SelMonth,-1)),IF(cur>prev,UNICHAR(9650),IF(cur<prev,UNICHAR(9660),UNICHAR(9644))))
```


Occupancy went from 81.5% in Oct 2025 to 84.7%. LET names the two rates once, so the IF reads like a sentence instead of repeating four SUMIFS. A formula-made arrow is ordinary text, so you can color it with conditional formatting. On a dashboard, rising occupancy deserves amber or red as it approaches the ceiling, even though the arrow itself only says which way it moved.

**8. Status cell for median door-to-provider: return the text On target or Off target for…**

- **Answer:** Off target
- **Solution:**

```
=LET(val,XLOOKUP(1,(tblKPI[Facility]=SelFacility)*(tblKPI[Month]=SelMonth),tblKPI[MedianDTP]),tgt,XLOOKUP("Median door-to-provider",tblTargets[KPI],tblTargets[Target]),dir,XLOOKUP("Median door-to-provider",tblTargets[KPI],tblTargets[Direction]),IF(IF(dir="Lower is better",val<=tgt,val>=tgt),"On target","Off target"))
```


The median is 31.0 minutes against a target of 30 or less, so it's off target, even if only barely. XLOOKUP(1, (condition)*(condition), …) finds the one row where both conditions are TRUE (TRUE×TRUE = 1). The inner IF flips the comparison by direction: lower-is-better KPIs pass at or below the target, and higher-is-better KPIs pass at or above it. Binary statuses hide near misses like this one, which is why many dashboards add an amber band (for example, within 10% of target).

**9. System-wide median door-to-provider, in minutes, for the selected month. Medians can't…**

- **Answer:** 33.5
- **Solution:**

```
=MEDIAN(FILTER(tblEDWaits[DoorToProviderMin],(tblEDWaits[ArrivalDateTime]>=SelMonth)*(tblEDWaits[ArrivalDateTime]<EDATE(SelMonth,1))))
```


The true system median is 33.5 minutes. Averaging the three hospital medians gives 34.5, which is not the median of anything. A median depends on every individual value, so it can only be computed from the detail rows. The date test uses a half-open window (>= first day, < first day of next month), which catches every arrival time on the last day of the month. In Excel 2019 or earlier, use =AGGREGATE(17,6,values/(condition),2) instead, where function 17 is QUARTILE.INC and quartile 2 is the median.

**10. Combo box cell link**

- **Answer:** 3
- **Solution:**

1. Show the **Developer** tab if you haven't (see the guide).
2. **Developer → Insert → Combo Box (Form Control)** (Mac: **Developer → Combo Box**), then drag a box onto the Dashboard.
3. Right-click the combo box → **Format Control** → **Control** tab. Input range: `Lists!$A$2:$A$5`. Cell link: `Calc!$C$12`. Drop down lines: 4. Click **OK**.
4. Click a cell to deselect the control, then choose **Ashby Falls Community Hospital**.

Calc!C13 turns the number back into a name: `=INDEX(Lists!$A$2:$A$5,Calc!C12)`.


A combo box writes the position of the chosen item, not its text: Ashby Falls Community Hospital is item 3 of the input range. INDEX(list, position) converts it back, which Calc!C13 does. To make the combo box drive the whole dashboard, point your formulas (or the SelFacility name) at that INDEX cell. A combo box floats above the grid, so it can't be typed over by accident, but it doesn't work in Excel for the web, and the position-number link breaks if someone re-sorts the list.

**11. 12-month trend block and chart (trailing-12-month ED visits)**

- **Answer:** 924
- **Solution:**

1. In **Calc!B17** type `=EDATE(SelMonth,ROWS(B$17:B17)-12)` and copy it down to B28. The first row gives EDATE(SelMonth,-11) and the last gives SelMonth itself.
2. In **C17** type `=SUMIFS(tblKPI[EDVisits],tblKPI[Facility],SelFacility,tblKPI[Month],B17)` and copy it down.
3. Select **B16:C28** → **Insert → Charts → Line**. The corner cell B16 is blank on purpose, so Excel uses the months as the axis.
4. Cut the chart (Ctrl + X, Mac: ⌘ + X), click a cell on the Dashboard, and paste. The chart keeps pointing at Calc.
5. Change the selectors on the Dashboard and watch the line redraw.


The chart's series points at formula cells, and those cells point at the selectors, so one dropdown redraws the chart. The block for Cedar Ridge Medical Center runs Dec 2024–Nov 2025 and peaks at 87 visits in Feb 2025. In Microsoft 365 you can instead spill the months with =EDATE(SelMonth,SEQUENCE(12,1,-11)). A chart built from a spill keeps a fixed range, so if a spill can change size, chart it through a defined name that refers to it (for example =Calc!$B$17#). If you pick an early month, the window reaches back before January 2024 and SUMIFS returns 0. Wrapping it as IF(COUNTIFS(…)=0,NA(),SUMIFS(…)) makes the line chart leave a gap instead of plunging to zero.

**12. PivotTable 1 with a slicer and a timeline**

- **Answer:** 185
- **Solution:**

1. Click in tblKPI → **Insert → PivotTable** → **New Worksheet** → **OK**. Rename the sheet **Pivots**.
2. Drag **Month** to Rows and **EDVisits** to Values (Sum of EDVisits). Excel may group the dates into Years and Quarters. That's fine.
3. **PivotTable Analyze → Insert Slicer** → tick **Facility** → **OK**. Click **Ashby Falls Community Hospital**.
4. **PivotTable Analyze → Insert Timeline** → tick **Month** → **OK**. Set the time level to **QUARTERS** and click **2025 Q3**.
5. Read the Grand Total.

Formula twin: `=SUMIFS(tblKPI[EDVisits],tblKPI[Facility],"Ashby Falls Community Hospital",tblKPI[Month],">="&DATE(2025,7,1),tblKPI[Month],"<="&DATE(2025,9,1))`


A slicer is a selector made of buttons, and a timeline is a selector for dates. Both filter the pivot they were inserted from. Visible selections are what make pivots dashboard-friendly: a Filters-area dropdown hides what's selected, but a slicer shows it. The formula twin proves the number: ED visits at Ashby Falls in July, August, and September 2025.

**13. PivotTable 2 connected to the same slicer and timeline**

- **Answer:** 1.62%
- **Solution:**

1. Click in tblKPI → **Insert → PivotTable** → **Existing Worksheet**, pick a cell on Pivots a few columns right of PivotTable 1 → **OK**.
2. Drag **Facility** to Rows, then **LWBS** and **EDVisits** to Values.
3. Select the Facility slicer → **Slicer → Report Connections** → tick both PivotTables → **OK**. Do the same for the timeline (**Timeline → Report Connections**). PivotTable 2 now shows only Ashby Falls, Q3 2025.
4. Click in PivotTable 2 → **PivotTable Analyze → Fields, Items & Sets → Calculated Field**. Name: `LWBSPct`. Formula: `=LWBS/EDVisits` → **Add** → **OK**. Format it as a percentage.

Formula twin: `=SUMIFS(tblKPI[LWBS],tblKPI[Facility],"Ashby Falls Community Hospital",tblKPI[Month],">="&DATE(2025,7,1),tblKPI[Month],"<="&DATE(2025,9,1))/SUMIFS(tblKPI[EDVisits],tblKPI[Facility],"Ashby Falls Community Hospital",tblKPI[Month],">="&DATE(2025,7,1),tblKPI[Month],"<="&DATE(2025,9,1))`


Report Connections is what lets one slicer filter several pivots. If you skipped it, PivotTable 2 would still show every hospital and month. A pivot calculated field sums each field first and then divides, so LWBSPct is total LWBS ÷ total visits for the quarter. That's exactly the right rate here: 1.62%. Averaging the three monthly rates would give 1.54%. Both pivots must come from the same source (tblKPI) to share a slicer.

</details>
<!-- END GENERATED: answers -->

## 🏆 Bonus challenge

This is the dashboard you'd actually hand to a board. Sketch the layout on paper first, build the model cells before the
cards, and compare your finished page with the hidden **Dashboard Key** sheet.

<!-- BEGIN GENERATED: bonus -->
Each month the COO presents one page to the board's Quality & Operations Committee. It must answer three questions at a glance: Are we on target? Where are we missing? Which way are things heading? Build it on the Board sheet to this spec:

1. Facility and Month dropdowns fed by the Lists sheet, named BoardFacility and BoardMonth. Facility must allow All facilities.
2. One card for each of the seven KPIs in tblTargets, showing the value, the target, a status colored by conditional formatting, and an arrow versus the same month last year. Every card must work for All facilities, so rebuild rates from their components and compute the median door-to-provider from tblEDWaits.
3. A scorecard line that counts the KPIs on target, such as '6 of 7 KPIs on target'.
4. A 12-month LWBS % trend block and a line chart that follow both dropdowns.
5. Polish: gridlines and headings off, only the two dropdowns unlocked, the sheet protected, and one landscape page when printed.

Then use your finished Board dashboard to answer B1–B4 below. The hidden Dashboard Key sheet is a finished reference build, so compare your numbers with it when you're done.

Work on the **Bonus** sheet of the workbook.

- **B1.** Set your Board dashboard to All facilities and Oct 2025. What is the system-wide 30-day readmission rate? Enter it as a percentage. *(Hint: Swap "All facilities" for the asterisk wildcard in the Facility criterion)*
- **B2.** With All facilities and Oct 2025 still selected, how many of the seven KPIs are on target? Use the Targets sheet's rule: on target means at or below a lower-is-better target, or at or above a higher-is-better one, so a value exactly equal to its target counts as on target. *(Hint: Give each card a status cell that returns 1 or 0, then SUM them. Remember which KPIs are higher-is-better)*
- **B3.** Keep Oct 2025 and switch the Facility dropdown to each hospital in turn. Which hospital has the fewest KPIs on target? *(Hint: Your scorecard line answers this. Change only the Facility dropdown)*
- **B4.** Back on All facilities and Oct 2025, your 12-month LWBS % trend runs Nov 2024–Oct 2025. In which month was the system-wide LWBS % highest? Enter the first day of that month as a date. *(Hint: INDEX(months, MATCH(MAX(rates), rates, 0)) on your trend block, or just read the chart's peak)*
<!-- END GENERATED: bonus -->

<!-- BEGIN GENERATED: bonus-answers -->
<details>
<summary><b>🔑 Show the bonus solution</b> — Give it a real try first!</summary>

**B1. System-wide readmission rate, Oct 2025**

- **Answer:** 18.96%
- **Solution:**

1. Make a helper cell named **FacCrit**: `=IF(BoardFacility="All facilities","*",BoardFacility)`.
2. The readmission card's value:

```
=SUMIFS(tblKPI[Readmits],tblKPI[Facility],FacCrit,tblKPI[Month],BoardMonth)/SUMIFS(tblKPI[IndexStays],tblKPI[Facility],FacCrit,tblKPI[Month],BoardMonth)
```


40 readmissions ÷ 211 index stays = 18.96%. No row in tblKPI says "All facilities", so SUMIFS with that text returns 0. The wildcard `*` matches any text, so the same SUMIFS adds all three hospitals. Every additive component works this way, and every rate is then rebuilt as total ÷ total.

**B2. Scorecard, All facilities, Oct 2025**

- **Answer:** 3
- **Solution:**

Give each card a numeric status cell, for example for LWBS:

```
=IF(ISNUMBER(val),--IF(dir="Lower is better",val<=tgt,val>=tgt),"")
```

Then the scorecard is `=SUM(statuses)&" of "&COUNT(statuses)&" KPIs on target"`.

The median card can't use FacCrit, because FILTER doesn't understand the `*` wildcard. Test the facility with OR logic instead:

```
=MEDIAN(FILTER(tblEDWaits[DoorToProviderMin],
  (tblEDWaits[ArrivalDateTime]>=BoardMonth)*(tblEDWaits[ArrivalDateTime]<EDATE(BoardMonth,1))
  *((tblEDWaits[Facility]=BoardFacility)+(BoardFacility="All facilities"))))
```

The system values for this month:

| KPI | Value | Target | Status |
|---|---|---|---|
| LWBS % | 1.4% | ≤ 2.0% | ✔ On target |
| Median door-to-provider | 38.0 min | ≤ 30.0 min | ✘ Off target |
| ALOS | 4.54 days | ≤ 4.50 days | ✘ Off target |
| 30-day readmission rate | 19.0% | ≤ 15.0% | ✘ Off target |
| Occupancy % | 79.1% | ≤ 85.0% | ✔ On target |
| HCAHPS top-box % | 50.0% | ≥ 50.0% | ✔ On target |
| Denial rate | 14.3% | ≤ 10.0% | ✘ Off target |


3 of 7. The median door-to-provider must come from tblEDWaits, because the three hospital medians can't be combined. HCAHPS top-box is the one higher-is-better KPI, and this month it lands at exactly 50.0%: on target with >=, off target with >. Decide ties in the KPI dictionary, not in each formula. Status cells that return 1 or 0 (not text) make the scorecard a plain SUM and make conditional formatting rules simple.

**B3. Hospital with the fewest KPIs on target, Oct 2025**

- **Answer:** Bluestone Memorial Hospital
- **Solution:**

Read the scorecard line after each switch:

| Hospital | On target |
|---|---|
| Bluestone Memorial Hospital | 2 of 7 |
| Ashby Falls Community Hospital | 5 of 7 |
| Cedar Ridge Medical Center | 4 of 7 |

The Dashboard Key's model area has this comparison table (columns N–V), with `=INDEX(N42:N44,MATCH(MIN(V42:V44),V42:V44,0))` picking the lowest.


Bluestone Memorial Hospital meets 2 of 7 targets in Oct 2025. A selector-driven dashboard answers "which hospital?" in a few clicks without a separate report per hospital. The large hospital being furthest off target is also why the system-wide scorecard lands where it does: its volumes dominate every total ÷ total rate.

**B4. Peak system LWBS % month in the 12-month trend**

- **Answer:** 12/01/2024
- **Solution:**

Trend block on the Board sheet (12 rows): month `=EDATE(BoardMonth,ROWS(first:current)-12)`, and the rate:

```
=IFERROR(SUMIFS(tblKPI[LWBS],tblKPI[Facility],FacCrit,tblKPI[Month],month)
  /SUMIFS(tblKPI[EDVisits],tblKPI[Facility],FacCrit,tblKPI[Month],month),NA())
```

Then `=INDEX(months,MATCH(AGGREGATE(4,6,rates),rates,0))`. AGGREGATE(4,6,…) is MAX that skips error values, so the #N/A gaps of an early month can't break it.


December 2024 peaks at 2.90%, against a target of 2%. Every other month in the window is lower. A target line on the chart (a second series that repeats the target in every row) makes the misses visible at a glance, and the dynamic title tells the reader which facility and months they're looking at.

</details>
<!-- END GENERATED: bonus-answers -->

## Key takeaways

- Plan from the audience and their questions to a short list of KPIs, each with a definition, a direction, a target, and an
  owner. The KPI dictionary is part of the dashboard.
- Build in three layers (data, model, dashboard), so new data flows through without anyone editing the page.
- Store additive components. Rebuild every rate as total ÷ total, and compute roll-up medians from the detail rows.
- Drive every number from named selector cells with SUMIFS, EDATE, and XLOOKUP. The `"*"` wildcard turns the same formula
  into an "All facilities" formula.
- A KPI card needs context: a target, a direction-aware status shown in both text and color, and a trend.
- Charts follow the selectors when they plot formula blocks. Slicers filter several PivotTables through Report Connections.
- Polish for the reader (gridlines off, protection, one printed page, alt text) and avoid volatile functions such as OFFSET
  and INDIRECT.

<!-- BEGIN GENERATED: nav -->
---

⬅️ **Previous:** [4.5 Statistics & Forecasting](../05-statistics-forecasting/README.md) · 🏠 [Course home](../../README.md) · **Next:** [5.1 Recording Your First Macros](../../05-automation-vba/01-recording-macros/README.md) ➡️
<!-- END GENERATED: nav -->
