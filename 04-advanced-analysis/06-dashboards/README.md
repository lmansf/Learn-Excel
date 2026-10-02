# Lesson 4.6 · Building Interactive Dashboards

> **Level:** Advanced · **Time:** about 70 minutes · **Workbook:** [`4.6-dashboards.xlsx`](4.6-dashboards.xlsx)
> **Data:** Monthly KPIs for Bluestone Health's three hospitals, Jan 2024 – Dec 2025 (72 hospital-months), every ED visit's door-to-provider time, and a KPI dictionary with targets and owners.

TODO: one-paragraph hook that explains why this skill matters in a hospital setting.

## What you'll learn

- Plan a dashboard around audience, questions, and KPIs
- Build KPI cards and selector-driven formulas
- Connect slicers to multiple PivotTables and drive charts from selections
- Polish layout, interactivity, and performance

## 📖 Guide

TODO: the teaching content.

## 🧪 Hands-on practice

Download [`4.6-dashboards.xlsx`](4.6-dashboards.xlsx) and open the **Practice** sheet. Type each answer in the yellow cell — the **Check**
column turns green when you're right.

<!-- BEGIN GENERATED: practice -->
The yellow selectors on the Dashboard sheet are named SelFacility (Dashboard!C4) and SelMonth (Dashboard!C5). They start at Cedar Ridge Medical Center and Nov 2025. Write every formula with SelFacility and SelMonth, never typed-in names or dates, and keep that selection while you check your answers. (Change it afterwards and watch your formulas follow.) Tasks 1–9 go in the yellow cells below. Tasks 10–13 have you build on other sheets.

| # | Task | Hint |
|:-:|------|------|
| 1 | Write a formula that returns the number of ED visits for the selected facility and month (the EDVisits column of tblKPI). | SUMIFS with two criteria: Facility = SelFacility and Month = SelMonth |
| 2 | Return the LWBS % (left without being seen) for the selection. Build it from its components, LWBS ÷ EDVisits, instead of reading the LWBSRate column. Enter it as a percentage. | One SUMIFS for the numerator divided by one SUMIFS for the denominator |
| 3 | LWBS variance to target: the selected LWBS % minus the LWBS % target from tblTargets. Look the target up with a formula instead of typing 2%. Enter the result as a percentage. A positive result means the hospital is above (worse than) this lower-is-better target. | XLOOKUP("LWBS %", tblTargets[KPI], tblTargets[Target]) |
| 4 | ED visits change versus the same month last year: this year ÷ last year − 1. Find last year's month with EDATE so the formula works for any selected month. Enter it as a percentage. | EDATE(SelMonth,-12) is the same month one year earlier |
| 5 | Occupancy trend arrow. Compare the selected month's occupancy (PatientDays ÷ BedDays) with the previous month's. Return ▲ with UNICHAR(9650) if it rose, ▼ with UNICHAR(9660) if it fell, or ▬ with UNICHAR(9644) if it didn't change. | LET(cur, …, prev, …, IF(cur>prev, UNICHAR(9650), …)). EDATE(SelMonth,-1) is the previous month |
| 6 | Status cell for median door-to-provider: return the text On target or Off target for the selection. Read the target and the Direction of "Median door-to-provider" from tblTargets, so the same pattern works for any KPI. A median isn't additive, so look up the hospital's MedianDTP value instead of adding it up. | LET + XLOOKUP(1, (tblKPI[Facility]=SelFacility)*(tblKPI[Month]=SelMonth), tblKPI[MedianDTP]); then IF on the direction |
| 7 | System-wide LWBS % for the selected month: all three hospitals combined, ignoring the facility selector. Divide total LWBS by total ED visits instead of averaging the three hospitals' rates. Enter it as a percentage. | Drop the Facility criterion from both SUMIFS |
| 8 | System-wide median door-to-provider, in minutes, for the selected month. Medians can't be added or averaged across hospitals, so compute it from the visit-level table: the MEDIAN of DoorToProviderMin for every tblEDWaits visit whose ArrivalDateTime falls in the selected month. Keep one decimal place. | MEDIAN(FILTER(…)) with ArrivalDateTime >= SelMonth and < EDATE(SelMonth,1) |
| 9 | Rolling 3-month ED visits for the selected facility: the selected month plus the two months before it (Sep–Nov 2025 for the default selection). Use one SUMIFS with a date window, not OFFSET. | Two criteria on the Month column: ">="&EDATE(SelMonth,-2) and "<="&SelMonth |
| 10 | Insert a Form Controls combo box on the Dashboard. In Format Control, set its Input range to Lists!$A$2:$A$5 and its Cell link to Calc!$C$12. Then choose Ashby Falls Community Hospital in the combo box. The gray cell shows the number your combo box writes to the cell link. | Developer → Insert → Combo Box (Form Control), then right-click it → Format Control → Control tab |
| 11 | On the Calc sheet, fill the yellow 12-month trend block. In B17:B28, return the 12 months ending at SelMonth, oldest first (use EDATE). In C17:C28, return ED visits for SelFacility in each of those months. Then select B16:C28, insert a line chart, and move it to the Dashboard. The gray cell totals your block: the trailing-12-month ED visits. | EDATE(SelMonth,-11) is the oldest month. A counter such as ROWS(B$17:B17) lets one formula step forward as you copy it down |
| 12 | Insert a new sheet named Pivots. Build PivotTable 1 from tblKPI with Month in Rows and Sum of EDVisits in Values. Insert a Facility slicer and a Month timeline for it. Select Ashby Falls Community Hospital in the slicer, and select 2025 Q3 in the timeline (switch it to QUARTERS). What is PivotTable 1's grand total? | PivotTable Analyze → Insert Slicer, and PivotTable Analyze → Insert Timeline |
| 13 | On the Pivots sheet, build PivotTable 2 from tblKPI with Facility in Rows, plus Sum of LWBS and Sum of EDVisits in Values. Connect the slicer and the timeline to it (Report Connections). Then add a calculated field LWBSPct = LWBS / EDVisits. With Ashby Falls and 2025 Q3 still selected, what LWBS % does PivotTable 2 show? Enter it as a percentage. | Select the slicer → Slicer → Report Connections. Then PivotTable Analyze → Fields, Items & Sets → Calculated Field |
<!-- END GENERATED: practice -->

## ✅ Answer key

The workbook has a hidden **Answer Key** sheet (right-click any sheet tab → **Unhide…**). The same answers are below,
collapsed so you don't see them by accident.

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

**3. LWBS variance to target: the selected LWBS % minus the LWBS % target from tblTargets.…**

- **Answer:** 1.75%
- **Solution:**

```
=SUMIFS(tblKPI[LWBS],tblKPI[Facility],SelFacility,tblKPI[Month],SelMonth)/SUMIFS(tblKPI[EDVisits],tblKPI[Facility],SelFacility,tblKPI[Month],SelMonth)-XLOOKUP("LWBS %",tblTargets[KPI],tblTargets[Target])
```


3.75% − 2.00% = +1.75%. Strictly, that's +1.75 percentage points: the gap between two rates. (The relative variance, actual ÷ target − 1, would be +87.5%.) Keeping targets in a table and looking them up means one edit to tblTargets updates every card that uses the target. On a real card this cell would subtract the target cell from the value cell.

**4. ED visits change versus the same month last year: this year ÷ last year − 1. Find last…**

- **Answer:** -7.0%
- **Solution:**

```
=SUMIFS(tblKPI[EDVisits],tblKPI[Facility],SelFacility,tblKPI[Month],SelMonth)/SUMIFS(tblKPI[EDVisits],tblKPI[Facility],SelFacility,tblKPI[Month],EDATE(SelMonth,-12))-1
```


80 visits this year against 86 in Nov 2024. Comparing with the same month last year removes seasonality, which a month-over-month change can't do: November always differs from October in an ED. EDATE moves a date by whole months and keeps it on the first of the month, so it matches the Month column exactly. Subtracting 365 days would not.

**5. Occupancy trend arrow. Compare the selected month's occupancy (PatientDays ÷ BedDays)…**

- **Answer:** ▲
- **Solution:**

```
=LET(cur,SUMIFS(tblKPI[PatientDays],tblKPI[Facility],SelFacility,tblKPI[Month],SelMonth)/SUMIFS(tblKPI[BedDays],tblKPI[Facility],SelFacility,tblKPI[Month],SelMonth),prev,SUMIFS(tblKPI[PatientDays],tblKPI[Facility],SelFacility,tblKPI[Month],EDATE(SelMonth,-1))/SUMIFS(tblKPI[BedDays],tblKPI[Facility],SelFacility,tblKPI[Month],EDATE(SelMonth,-1)),IF(cur>prev,UNICHAR(9650),IF(cur<prev,UNICHAR(9660),UNICHAR(9644))))
```


Occupancy went from 81.5% in Oct 2025 to 84.7%. LET names the two rates once, so the IF reads like a sentence instead of repeating four SUMIFS. A formula-made arrow is ordinary text, so you can color it with conditional formatting. On a dashboard, rising occupancy deserves amber or red as it approaches the ceiling, even though the arrow itself only says which way it moved.

**6. Status cell for median door-to-provider: return the text On target or Off target for…**

- **Answer:** Off target
- **Solution:**

```
=LET(val,XLOOKUP(1,(tblKPI[Facility]=SelFacility)*(tblKPI[Month]=SelMonth),tblKPI[MedianDTP]),tgt,XLOOKUP("Median door-to-provider",tblTargets[KPI],tblTargets[Target]),dir,XLOOKUP("Median door-to-provider",tblTargets[KPI],tblTargets[Direction]),IF(IF(dir="Lower is better",val<=tgt,val>=tgt),"On target","Off target"))
```


The median is 31.0 minutes against a target of 30 or less, so it's off target, even if only barely. XLOOKUP(1, (condition)*(condition), …) finds the one row where both conditions are TRUE (TRUE×TRUE = 1). The inner IF flips the comparison by direction: lower-is-better KPIs pass at or below the target, and higher-is-better KPIs pass at or above it. Binary statuses hide near misses like this one, which is why many dashboards add an amber band (for example, within 10% of target).

**7. System-wide LWBS % for the selected month: all three hospitals combined, ignoring the…**

- **Answer:** 1.53%
- **Solution:**

```
=SUMIFS(tblKPI[LWBS],tblKPI[Month],SelMonth)/SUMIFS(tblKPI[EDVisits],tblKPI[Month],SelMonth)
```


8 ÷ 522 = 1.53%. Averaging the three hospital rates gives 1.71%, because a simple average gives the small hospitals the same weight as Bluestone Memorial, which sees several times as many patients. Rates roll up as total numerator ÷ total denominator. That's why tblKPI stores the components.

**8. System-wide median door-to-provider, in minutes, for the selected month. Medians can't…**

- **Answer:** 33.5
- **Solution:**

```
=MEDIAN(FILTER(tblEDWaits[DoorToProviderMin],(tblEDWaits[ArrivalDateTime]>=SelMonth)*(tblEDWaits[ArrivalDateTime]<EDATE(SelMonth,1))))
```


The true system median is 33.5 minutes. Averaging the three hospital medians gives 34.5, which is not the median of anything. A median depends on every individual value, so it can only be computed from the detail rows. The date test uses a half-open window (>= first day, < first day of next month), which catches every arrival time on the last day of the month. In Excel 2019 or earlier, use =AGGREGATE(17,6,values/(condition),2) instead, where function 17 is QUARTILE.INC and quartile 2 is the median.

**9. Rolling 3-month ED visits for the selected facility: the selected month plus the two…**

- **Answer:** 221
- **Solution:**

```
=SUMIFS(tblKPI[EDVisits],tblKPI[Facility],SelFacility,tblKPI[Month],">="&EDATE(SelMonth,-2),tblKPI[Month],"<="&SelMonth)
```


Joining an operator to a date (">="&EDATE(SelMonth,-2)) turns the date into criteria text that SUMIFS understands. A common older pattern is SUM(OFFSET(…)). OFFSET is volatile, so Excel recalculates it after every edit anywhere in the workbook, and dozens of them make a dashboard sluggish. SUMIFS recalculates only when its inputs change, and it doesn't depend on the table being sorted.

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


The chart's series points at formula cells, and those cells point at the selectors, so one dropdown redraws the chart. The block for Cedar Ridge Medical Center runs Dec 2024–Nov 2025 and peaks at 87 visits in Feb 2025. In Microsoft 365 you can instead spill the months with =EDATE(SelMonth,SEQUENCE(12,1,-11)). Charts can't point at a spill directly, though, so you'd chart the spill range through a defined name. If you pick an early month, the window reaches back before January 2024 and SUMIFS returns 0. Wrapping it as IF(COUNTIFS(…)=0,NA(),SUMIFS(…)) makes the line chart leave a gap instead of plunging to zero.

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

<!-- BEGIN GENERATED: bonus -->
Each month the COO presents one page to the board's Quality & Operations Committee. It must answer three questions at a glance: Are we on target? Where are we missing? Which way are things heading? Build it on the Board sheet to this spec:
1. Facility and Month dropdowns fed by the Lists sheet, named BoardFacility and BoardMonth. Facility must allow All facilities.
2. One card for each of the seven KPIs in tblTargets, showing the value, the target, a status colored by conditional formatting, and an arrow versus the same month last year. Every card must work for All facilities, so rebuild rates from their components and compute the median door-to-provider from tblEDWaits.
3. A scorecard line such as '3 of 7 KPIs on target'.
4. A 12-month LWBS % trend block and a line chart that follow both dropdowns.
5. Polish: gridlines and headings off, only the two dropdowns unlocked, the sheet protected, and one landscape page when printed.
The hidden Dashboard Key sheet is a finished reference build. Compare your numbers with it when you're done.

Work on the **Bonus** sheet of the workbook.

- **B1.** Set your Board dashboard to All facilities and Oct 2025. What is the system-wide 30-day readmission rate? Enter it as a percentage. *(Hint: Swap "All facilities" for the wildcard "*" in the Facility criterion)*
- **B2.** With All facilities and Oct 2025 still selected, how many of the seven KPIs are on target? *(Hint: Give each card a status cell that returns 1 or 0, then SUM them. Remember which KPIs are higher-is-better)*
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


40 readmissions ÷ 211 index stays = 18.96%. No row in tblKPI says "All facilities", so SUMIFS with that text returns 0. The wildcard * matches any text, so the same SUMIFS adds all three hospitals. Every additive component works this way, and every rate is then rebuilt as total ÷ total.

**B2. Scorecard, All facilities, Oct 2025**

- **Answer:** 3
- **Solution:**

Give each card a numeric status cell, for example for LWBS:

```
=IF(ISNUMBER(val),--IF(dir="Lower is better",val<=tgt,val>=tgt),"")
```

Then the scorecard is `=SUM(statuses)&" of "&COUNT(statuses)&" KPIs on target"`.

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

- TODO

<!-- BEGIN GENERATED: nav -->
---

⬅️ **Previous:** [4.5 Statistics & Forecasting](../05-statistics-forecasting/README.md) · 🏠 [Course home](../../README.md) · **Next:** [5.1 Recording Your First Macros](../../05-automation-vba/01-recording-macros/README.md) ➡️
<!-- END GENERATED: nav -->
