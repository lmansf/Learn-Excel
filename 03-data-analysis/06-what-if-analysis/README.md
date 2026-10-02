# Lesson 3.6 · What-If Analysis: Goal Seek, Scenarios, Data Tables & Solver

> **Level:** Intermediate · **Time:** about 60 minutes · **Workbook:** [`3.6-what-if-analysis.xlsx`](3.6-what-if-analysis.xlsx)
> **Data:** A monthly operating model for Bluestone's Primary Care Clinic (D400, Bluestone Outpatient Pavilion). Payer mix and reimbursement come from the clinic's 2,079 claims with 2025 service dates, wage rates from Bluestone HR records, and benefits, supply, and fixed costs from the clinic's 2025 budget actuals.

TODO: one-paragraph hook that explains why this skill matters in a hospital setting.

## What you'll learn

- Structure a model with separate inputs, calculations, and outputs
- Find break-even points with Goal Seek
- Compare cases with Scenario Manager and sensitivity Data Tables
- Optimize a staffing mix with Solver

## 📖 Guide

TODO: the teaching content.

## 🧪 Hands-on practice

Download [`3.6-what-if-analysis.xlsx`](3.6-what-if-analysis.xlsx) and open the **Practice** sheet. Type each answer in the yellow cell — the **Check**
column turns green when you're right.

<!-- BEGIN GENERATED: practice -->
Every task uses the Model and Staffing sheets. Tasks 1 and 2 read your Model live, so their checks stay green only while the Model holds its base-case inputs (200 visits per day, Commercial $142, and so on). After each what-if run, put the original values back: click Cancel in the Goal Seek Status box, and choose Restore Original Values in the Solver Results box.

| # | Task | Hint |
|:-:|------|------|
| 1 | On the Model sheet, complete the yellow Operating income cell (B59): net patient revenue minus total operating expenses. The gray cell here reads your formula. What is the clinic's base-case operating income per month? | An output should only point at calculation cells |
| 2 | Complete the yellow Operating margin cell (B60): operating income as a share of net patient revenue. What is the base-case operating margin? (Format it as a percentage.) | Margin = income ÷ revenue |
| 3 | Audit the model. One formula in the CALCULATIONS section (rows 42–56) has a number typed into it instead of a reference to its input cell. Type that cell's address (for example B99). Then fix the formula so it points to the input. Task 10 depends on the fix. | Show Formulas, or Trace Dependents on each input |
| 4 | Use Goal Seek to find the break-even volume: set Operating income to 0 by changing Visits per day (Model!B8). How many visits per day does the clinic need? Round to 1 decimal place, then click Cancel to restore 200. | Data → What-If Analysis → Goal Seek |
| 5 | The CFO is renegotiating commercial contracts. Use Goal Seek to set Operating margin to 5% (type 0.05) by changing the Commercial $ per visit (Model!C17). What commercial reimbursement per visit is needed? Round to the nearest dollar, then click Cancel to restore $142. | Set cell = the margin cell; it must contain a formula |
| 6 | Build the one-variable Data Table in Model!F8:H17: put =B59 in G8 and =B60 in H8, select F8:H17, and use Column input cell B8. What operating margin does your table show at 190 visits per day? Enter it as a percentage to 1 decimal place. | Data → What-If Analysis → Data Table; the visits run down a column |
| 7 | Look down the Operating income column of your one-variable table. What is the lowest visits-per-day value in the table at which the clinic makes a profit (operating income above 0)? | Read the table, or let MINIFS find it |
| 8 | Build the two-variable Data Table: put =B59 in the corner cell F22, select F22:K31, and use Row input cell C17 (commercial $ across row 22) and Column input cell B8 (visits per day down column F). What operating income does the table show at 210 visits per day and $160 per commercial visit? Round to the nearest dollar. | Row input = the input whose values run across the top row |
| 9 | How many of the 45 combinations in your two-variable table (G23:K31) are profitable (operating income above 0)? Use a formula. | COUNTIF with ">0" |
| 10 | Open Scenario Manager on the Model and add three scenarios that change B8, B16, B17 and B36. Base plan: visits per day 200, Medicaid share 14.9%, Commercial share 37.6%, billing fee 4.0%; Downside: visits per day 190, Medicaid share 17.9%, Commercial share 34.6%, billing fee 5.0%; Upside: visits per day 206, Medicaid share 12.9%, Commercial share 39.6%, billing fee 3.5%. Create a Scenario Summary with result cells B59 and B60. What is operating income in the Downside scenario? Round to the nearest dollar. | Data → What-If Analysis → Scenario Manager |
| 11 | From the same Scenario Summary, what operating margin does the Upside scenario produce? Enter it as a percentage to 1 decimal place. | Same summary, different column |
| 12 | On the Staffing sheet, use Solver to minimize the monthly staff cost (B31) by changing the RN, LPN, and CNA FTEs (B18:B20), subject to the three constraints in rows 25–27 (leave out the CNA cap in row 28). Keep Make Unconstrained Variables Non-Negative ticked and choose Simplex LP. What is the minimum monthly cost? Round to the nearest dollar. | Data → Solver (enable the Solver add-in first) |
| 13 | Run Solver again (same setup) and select Sensitivity under Reports before you click OK. On the Sensitivity Report sheet, what is the Shadow Price of the Total support hours constraint (Staffing!B25)? Enter it in dollars to 2 decimal places. | Shadow price = cost change per one-unit increase in the constraint's right-hand side |
<!-- END GENERATED: practice -->

## ✅ Answer key

The workbook has a hidden **Answer Key** sheet (right-click any sheet tab → **Unhide…**). The same answers are below,
collapsed so you don't see them by accident.

<!-- BEGIN GENERATED: answers -->
<details>
<summary><b>🔑 Show the answer key</b> — Try every task before opening this.</summary>

**1. Operating income (Model!B59)**

- **Answer:** -8,618.72
- **Solution:** `=B44-B56`

Operating income is revenue minus expenses, and both already exist as calculation rows (B44 and B56). Because those cells are named, `=NetRevenue-TotalExpenses` works too and reads better. The clinic loses about $8,619 a month in the base case. That's common for hospital-owned primary care, which is why the rest of the lesson asks what it would take to break even.

**2. Operating margin (Model!B60)**

- **Answer:** -1.9%
- **Solution:** `=B59/B44`

Margin divides operating income by revenue (`=OperatingIncome/NetRevenue`). A margin lets you compare a small clinic with a large hospital, which a dollar figure can't. Goal Seek will aim at this cell in task 5.

**3. Find the hard-coded number**

- **Answer:** B52
- **Solution:**

1. Press **Ctrl + `** (Mac: **⌃ + `**) to show formulas, and read down rows 42–56.
2. B52 reads `=NetRevenue*0.04`. The 4% is typed in, although the fee has its own input cell, B36.
3. Change B52 to `=NetRevenue*BillingFeePct` (or `=B44*B36`), then press Ctrl + ` again to show values.


Today both versions return the same number, so nothing looks wrong. The trouble starts when someone changes the fee input in B36: the model ignores it. Another way to catch this is **Formulas → Trace Dependents** on B36. Excel finds no dependents and beeps (or says so in a message), because no formula reads that input.

**4. Goal Seek: break-even visits per day**

- **Answer:** 204.2
- **Solution:**

1. Choose **Data → What-If Analysis → Goal Seek**.
2. **Set cell:** `B59` · **To value:** `0` · **By changing cell:** `B8`.
3. Click **OK**. B8 shows about 204.2440. Note it, then click **Cancel**.


Each extra visit per day adds one visit on each of the 21 clinic days. Each of those visits brings in $96.71 after the billing fee, supplies, and vaccines (its **contribution margin**). Fixed costs are $414,783 a month, so break-even = $414,783 ÷ ($96.71 × 21) ≈ 204.2. That is 98% of the clinic's 208-visit daily capacity, so volume alone is a fragile fix. The live result in the key does this algebra, and Goal Seek reaches the same answer by trial and error.

**5. Goal Seek: commercial rate for a 5% margin**

- **Answer:** about $164 (Goal Seek shows 163.58)
- **Solution:**

1. **Data → What-If Analysis → Goal Seek**.
2. **Set cell:** `B60` · **To value:** `0.05` · **By changing cell:** `C17`.
3. Click **OK**, read C17, and click **Cancel**.


Commercial plans would have to pay about $164 instead of $142, an increase of 15%. Only 37.6% of visits are commercial, so each extra commercial dollar moves the average reimbursement by just 37.6 cents. Margin is a ratio, so Goal Seek has to iterate here, and it stops when the margin is within its tolerance of 5%. That's why the check accepts a small range around the exact value.

**6. One-variable Data Table: margin at 190 visits/day**

- **Answer:** -6.8%
- **Solution:** `=Model!H11`

The visits per day values run **down a column** (F9:F17), so B8 is the **column** input cell and the row input box stays empty. For each value, Excel puts it in B8, recalculates the model, and writes the results of the formulas in row 8 into that row of the table. You can type the number or point at H11.

**7. One-variable Data Table: first profitable volume**

- **Answer:** 205
- **Solution:** `=MINIFS(Model!F9:F17,Model!G9:G17,">0")`

The table jumps in steps of 5, so it brackets the break-even point instead of finding it: 200 visits loses money and 205 makes money, which agrees with Goal Seek's 204.2. Use a Data Table to see the whole curve, and Goal Seek to pin down the exact crossing.

**8. Two-variable Data Table: 210 visits/day × $160**

- **Answer:** 40,342
- **Solution:** `=Model!J29`

A two-variable table has exactly one formula, in its top-left corner. Excel substitutes each top-row value into the **row** input cell (C17) and each left-column value into the **column** input cell (B8), and fills every intersection. If you swap the two input cells, the table still fills without any warning, but with wrong numbers. So check one cell by hand: at 200 visits and $140 the result should be close to the base case (−$8,619), because the base case is 200 visits at $142.

**9. Two-variable Data Table: profitable combinations**

- **Answer:** 24
- **Solution:** `=COUNTIF(Model!G23:K31,">0")`

Ordinary formulas can read a Data Table's results, so you can summarize or chart them, or add conditional formatting. The profitable combinations sit in the bottom-right of the table, where volume and rate are both high. That pattern tells the CFO that a better contract and more visits each lower the volume the clinic needs to break even.

**10. Scenario Manager: Downside operating income**

- **Answer:** -41,824
- **Solution:**

1. Select B8, then Ctrl-click (Mac: ⌘-click) B16:B17 and B36.
2. **Data → What-If Analysis → Scenario Manager → Add…**. Name it *Base plan* and click **OK**. The values box shows the current values, so click **OK** again.
3. Click **Add…** again for *Downside* and *Upside*. Type the values in the order B8, B16, B17, B36 (decimals like 0.179 for 17.9%).
4. Click **Summary…**, choose **Scenario summary**, set **Result cells** to `B59,B60`, and click **OK**. Read the Downside column on the new Scenario Summary sheet.


Fewer visits, a shift from commercial to Medicaid, and a higher vendor fee all hit income at once. If you skipped the fix in task 3, the model ignores the 5% fee and you'd see −$37,660 instead. That's exactly the kind of silent error a hard-coded number causes. The summary sheet is a snapshot. It doesn't update when the model changes, so create it again after any edit.

**11. Scenario Manager: Upside margin**

- **Answer:** 2.6%
- **Solution:** Read the Upside column of the OperatingMargin row on the Scenario Summary sheet.

The Upside case brings the clinic to roughly 2.6%. It takes three things at once: 6 more visits a day, 2 points of payer mix shifted from Medicaid to Commercial, and a cheaper vendor contract (3.5%). Because B59 and B60 are named, the summary labels those rows OperatingIncome and OperatingMargin instead of $B$59 and $B$60.

**12. Solver: lowest-cost staffing mix (FTEs)**

- **Answer:** 88,342
- **Solution:** `=Staffing!B31`

Solver parameters: **Set Objective** `$B$31`, **To** Min, **By Changing Variable Cells** `$B$18:$B$20`, **Subject to the Constraints** `$B$25:$B$27 >= $D$25:$D$27`, Simplex LP. Then keep the solution and point at (or type) B31. Solver chooses 4.2 RN, 2.8 LPN and 8.0 CNA FTEs. This answer makes sense: CNAs are the cheapest staff, so they cover every hour that doesn't need a licensed nurse. The licensed hours are then split at exactly the 60% RN minimum, because LPNs cost less than RNs. Today's 5/3/8 staffing costs $96,798, so the plan saves about $8,457 a month.

**13. Solver Sensitivity Report: shadow price**

- **Answer:** 25.01
- **Solution:**

1. **Data → Solver**, same parameters as task 12, **Solve**.
2. In **Solver Results**, select **Sensitivity** under *Reports*, then click **OK**.
3. On the new *Sensitivity Report 1* sheet, find the row for `$B$25` under *Constraints* and read **Shadow Price**.


Requiring one more support hour raises the minimum cost by $25.01. That's one CNA hour at $19.74 plus 26.7% benefits, because the cheapest way to cover an hour that doesn't need a licensed nurse is a CNA hour. The live result in the key repeats that arithmetic. The licensed-hours constraint has a higher shadow price ($21.53). An extra licensed hour costs 0.6 RN hours plus 0.4 LPN hours, but it also replaces a CNA hour that's no longer needed.

</details>
<!-- END GENERATED: answers -->

## 🏆 Bonus challenge

<!-- BEGIN GENERATED: bonus -->
Bluestone wants the Primary Care Clinic to grow to 220 visits per day in 2026 by adding a fourth APP (capacity rises from 208 to 226 visits per day). HR hires whole people, so every role needs a whole number of 1.0-FTE staff. Only 6 CNA positions are approved (Staffing!B12). Find the cheapest staffing plan, then test the full plan in the P&L.

Work on the **Bonus** sheet of the workbook.

- **B1.** On the Staffing sheet, change B6 to 220. Re-run Solver with two more constraints: B18:B20 = int (integer) and B28 <= D28 (the CNA cap). In Solver Options, set Integer Optimality (%) to 0. What is the minimum monthly staff cost? Round to the nearest dollar. *(Hint: Add an int constraint on the three variable cells)*
- **B2.** How many RNs does that plan hire? *(Hint: Look at B18)*
- **B3.** How many LPNs does that plan hire? *(Hint: Is rounding the fractional answer up the same thing?)*
- **B4.** Test the plan in the P&L. On the Model, add two more scenarios that change B8, B10 and B21:B23: Today (200, 3, 5, 3, 8) and Growth 2026 (220, 4, and your RN, LPN, and CNA counts). What operating income does Growth 2026 produce? Round to the nearest dollar. *(Hint: Scenario Manager can switch five inputs at once, and switch them back)*
<!-- END GENERATED: bonus -->

<!-- BEGIN GENERATED: bonus-answers -->
<details>
<summary><b>🔑 Show the bonus solution</b> — Give it a real try first!</summary>

**B1. Integer staffing plan: minimum cost**

- **Answer:** 112,488
- **Solution:**

1. Type 220 in Staffing!B6.
2. **Data → Solver**. Keep the objective, variable cells, and the three constraints from task 12.
3. **Add** `$B$18:$B$20` **int** and **Add** `$B$28 <= $D$28`.
4. **Options** → *All Methods* tab → **Integer Optimality (%)** = 0 → **OK**. Check that **Ignore Integer Constraints** is not ticked.
5. **Solve**, keep the solution, and read B31.


With fractions allowed, the cheapest plan costs $107,302 (6.3 RN, 4.2 LPN, 6.0 CNA). Requiring whole people raises that to $112,488, so indivisibility costs $5,186 a month. Solver handles integer constraints by **branch and bound**, solving many LPs with tighter and tighter bounds. Its default Integer Optimality of 1% lets it stop at any plan within 1% of the best possible, so set it to 0 when the exact answer matters.

**B2. Integer plan: RNs**

- **Answer:** 7
- **Solution:** Read Staffing!B18 after Solver finishes.

7 RNs out of 11 licensed staff is 63.6%, which meets the 60% rule.

**B3. Integer plan: LPNs**

- **Answer:** 4
- **Solution:** Read Staffing!B19 after Solver finishes.

The fractional plan needs 10.5 licensed FTEs, and 11 whole people already cover that. A tempting shortcut is to round each fractional value up, which gives 7 RN, 5 LPN and 6 CNA. That plan hires one licensed person too many and costs $5,902 a month more. It also breaks the RN rule, because 7 of 12 licensed staff is only 58.3%. Rounding each variable separately can't see how the constraints interact, but Solver's integer search can.

**B4. Growth 2026 scenario: operating income**

- **Answer:** 2,624
- **Solution:**

1. On the Model, select B8, B10 and B21:B23 (Ctrl-click; Mac: ⌘-click).
2. **Scenario Manager → Add…** *Today*: 200, 3, 5, 3, 8.
3. **Add…** *Growth 2026*: 220, 4, 7, 4, 6.
4. **Summary…** with result cells `B59,B60`, or select Growth 2026, click **Show**, read B59, then **Show** Today to restore the base case.


The growth plan turns a −$8,619 monthly loss into $2,624 (0.5% margin). The new APP costs $10,800 plus benefits, and the larger staff costs $15,690 more a month than today's. The extra 420 visits a month each contribute $96.71, which covers both. Using scenarios instead of overwriting inputs means one click (Show Today) puts the base case back for every other task.

</details>
<!-- END GENERATED: bonus-answers -->

## Key takeaways

- TODO

<!-- BEGIN GENERATED: nav -->
---

⬅️ **Previous:** [3.5 Charts & Data Visualization](../05-charts-visualization/README.md) · 🏠 [Course home](../../README.md) · **Next:** [4.1 Dynamic Arrays: FILTER, SORT, UNIQUE & More](../../04-advanced-analysis/01-dynamic-arrays/README.md) ➡️
<!-- END GENERATED: nav -->
