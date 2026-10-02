# Lesson 1.4 · Your First Formulas & Functions

> **Level:** Beginner · **Time:** about 55 minutes · **Workbook:** [`1.4-basic-formulas.xlsx`](1.4-basic-formulas.xlsx)
> **Data:** Daily census for Medical-Surgical 4 West at Bluestone Memorial Hospital, Jan–Mar 2025 (90 days), plus the unit's supply room stock.

Every morning, a nurse manager has to answer questions like *"How full were we last month?"*, *"How long are patients staying?"*,
and *"What's sitting in our supply room?"* With a calculator, each answer takes several minutes. With a formula it takes
seconds, and the answer **updates itself** when the data changes. In this lesson you'll write your first formulas against a real-looking
unit census and learn the handful of functions you'll use more than any others.

## What you'll learn

- Write formulas with operators and the correct order of operations
- Use SUM, AVERAGE, MIN, MAX, COUNT, COUNTA, and COUNTBLANK
- Calculate rates and percentages such as bed occupancy
- Copy formulas down a column and read common errors (#DIV/0!, #NAME?, #VALUE!)
- Round results with ROUND

## 📖 Guide

### 1. What is a formula?

A **formula** is an instruction that tells Excel to calculate something. Every formula starts with an equals sign `=`.

| You type | Excel shows | Why |
|---|---|---|
| `36` | 36 | A constant: just a number |
| `=36-33` | 3 | A formula with two constants |
| `=B2-E2` | 3 (if B2 = 36 and E2 = 33) | A formula with **cell references** |

Cell references are what make spreadsheets powerful. `=B2-E2` means "take whatever is in B2 and subtract whatever is in E2."
Change E2 to 30 and the result instantly becomes 6. **Prefer references to typed-in numbers.** If you write `=36-33`, Excel can't
update it for you.

> 💡 **Tip:** You don't have to type references. After typing `=`, click a cell (or drag across a range) and Excel inserts the
> address for you. Press **Enter** to finish, or **Esc** to cancel.

The cell shows the **result**. The **formula bar** shows the formula. Click a cell and look at the formula bar to see how its value was
calculated. Press **Ctrl + `` ` ``** (the grave accent key, next to 1; Mac: **Control + `` ` ``**) to toggle
*Show Formulas* for the whole sheet.

### 2. Operators and the order of operations

| Operator | Meaning | Example | Result |
|:-:|---|---|---|
| `+` | add | `=33+2` | 35 |
| `-` | subtract | `=36-33` | 3 |
| `*` | multiply | `=12*3` | 36 |
| `/` | divide | `=33/36` | 0.9167 |
| `^` | exponent (power) | `=2^3` | 8 |
| `&` | join text | `="4"&" West"` | 4 West |
| `%` | percent | `=85%` | 0.85 |

Excel follows the same order of operations you learned in math class (often remembered as **PEMDAS**):

1. **P**arentheses `( )`
2. **E**xponents `^`
3. **M**ultiplication and **D**ivision `*` `/`, left to right
4. **A**ddition and **S**ubtraction `+` `-`, left to right

So `=10+20/5` is **14** (the division happens first), but `=(10+20)/5` is **6**. When in doubt, add parentheses. They never hurt and
they make your intent obvious to the next reader.

> ⚠️ **Classic trap:** average of two numbers. `=A2+B2/2` adds A2 to *half* of B2. You meant `=(A2+B2)/2`.

### 3. Functions

A **function** is a built-in formula with a name. You give it **arguments** inside parentheses, separated by commas:

```
=FUNCTIONNAME(argument1, argument2, ...)
=SUM(C2:C91)
```

`C2:C91` is a **range**: every cell from C2 down to C91. The colon means "through."

Ways to enter a function:

- **Type it.** As you type `=SU`, Excel shows a list of matching functions. Press **Tab** to accept one. A ScreenTip then shows the
  arguments the function expects.
- **AutoSum.** Select the cell just below a column of numbers and press **Alt + =** (Mac: **⌘ + Shift + T**). Excel guesses the range
  and writes `=SUM(...)` for you. The drop-down arrow on the **Home → AutoSum** button also offers Average, Count Numbers, Max, and Min.
- **Insert Function.** Click **fx** next to the formula bar to search for a function and fill in its arguments in a dialog.

### 4. The seven functions you'll use every day

| Function | What it does | Example on the Census sheet |
|---|---|---|
| `SUM(range)` | Adds the numbers | `=SUM(C2:C91)` gives total admissions |
| `AVERAGE(range)` | Mean of the numbers (sum ÷ count) | `=AVERAGE(E2:E91)` gives average daily census |
| `MIN(range)` | Smallest number | `=MIN(E2:E91)` gives the quietest night |
| `MAX(range)` | Largest number | `=MAX(E2:E91)` gives the busiest night |
| `COUNT(range)` | How many cells contain **numbers** (dates count!) | `=COUNT(A2:A91)` gives the number of days |
| `COUNTA(range)` | How many cells are **not empty** (numbers, text, anything) | `=COUNTA(F2:F91)` gives days with a note |
| `COUNTBLANK(range)` | How many cells are **empty** | `=COUNTBLANK(F2:F91)` gives days without a note |

A few things beginners trip over:

- **COUNT ignores text.** `=COUNT(F2:F91)` on the Notes column returns 0, because notes are text. Use COUNTA.
- **Dates are numbers.** Excel stores 01/01/2025 as the serial number 45658, so COUNT, MIN, and MAX all work on dates. (Lesson 2.3
  covers this in depth.)
- **AVERAGE skips blanks but includes zeros.** An empty cell is ignored. A cell containing 0 pulls the average down.
- Functions accept several ranges: `=SUM(C2:C91, D2:D91)` adds both columns.

### 5. Healthcare metrics you can build from a census

A **midnight census** is the number of patients in beds at midnight, counted once per day. From that column you can build
most of the standard unit metrics:

| Metric | Definition | Formula idea |
|---|---|---|
| **Patient days** | Sum of the daily midnight census | `=SUM(census)` |
| **Average daily census (ADC)** | Patient days ÷ number of days | `=AVERAGE(census)` |
| **Occupancy rate** | Patient days ÷ staffed-bed days | `=SUM(census)/SUM(beds)` |
| **Average length of stay (ALOS)** | Patient days ÷ discharges | `=SUM(census)/SUM(discharges)` |

> 💡 **Tip:** For a rate over a period, compute *total numerator ÷ total denominator*, as in `=SUM(census)/SUM(beds)`. Don't
> average the daily percentages. When the denominators differ from day to day (for example, if beds open and close), the
> average of daily rates is slightly wrong. The SUM/SUM pattern is always right.

To show a rate as a percentage, select the cell and press **Ctrl + Shift + %** (Mac: **Control + Shift + %**) or click **Home → %**.
The stored value is still 0.9126…. Formatting only changes how it's *displayed* (see Lesson 1.3).

### 6. Copying formulas down a column

Often you need the same calculation on every row, like daily occupancy = census ÷ beds. You don't write 90 formulas. You write one and
copy it:

1. In the first data row, type `=E2/B2` and press **Enter**.
2. Select that cell and **double-click the fill handle** (the small square at the cell's bottom-right corner). Excel copies the formula
   down to the last row of data. You can also drag the fill handle, or select the range and press **Ctrl + D** (Mac: **⌘ + D**).

Each copied formula **adjusts to its own row**: row 3 gets `=E3/B3`, row 4 gets `=E4/B4`, and so on. These are called
**relative references**, and Lesson 1.5 explains exactly how they work (and how to stop a reference from moving).

> 📋 **Excel Tables:** the Census data is formatted as an Excel Table (blue banded rows). When you type a formula in the first cell of
> a blank Table column, Excel often fills the whole column automatically. It may also write the formula as
> `=[@MidnightCensus]/[@StaffedBeds]`. That's a *structured reference* meaning "this row's MidnightCensus." Both styles give
> the same result. You'll learn structured references in Lesson 3.1.

### 7. Rounding with ROUND

```
=ROUND(number, num_digits)
=ROUND(4.6667, 1)   → 4.7
=ROUND(4.6667, 0)   → 5
=ROUND(1234.5, -2)  → 1200
```

**Formatting vs. rounding:** If you format 4.6667 to show one decimal, the cell *displays* 4.7 but still *stores* 4.6667, and later
calculations use 4.6667. `ROUND` changes the stored value itself. Use ROUND when the rounded number is what you mean, such as a
reported ALOS or a billable quantity. Use formatting when you only want a tidier display.

Related functions you'll meet later: `ROUNDUP` and `ROUNDDOWN` always round away from or toward zero (useful for "you can't
staff half a nurse"), and `MROUND`, `CEILING.MATH`, and `FLOOR.MATH` round to multiples (Lesson 2.4).

### 8. When formulas go wrong

| You see | It means | Usual fix |
|---|---|---|
| `#DIV/0!` | Dividing by zero or by an empty cell | Check the denominator. Later you'll wrap it in `IFERROR` (Lesson 2.1) |
| `#NAME?` | Excel doesn't recognize a name, often a typo like `=SUMM(...)` or text missing its quotes | Fix the spelling and use AutoComplete |
| `#VALUE!` | Wrong type of data, such as math on text (`="abc"*2`) | Check for text that only *looks* like a number |
| `#REF!` | A referenced cell was deleted | Undo (Ctrl + Z, Mac: ⌘ + Z) or rewrite the reference |
| `#####` | The column is too narrow to show the number or date | Widen the column (double-click its right border) |

To investigate any formula, select the cell and press **F2** (Mac: **Control + U**, or **Fn + F2**). Excel color-codes each reference and outlines the matching cells on the
sheet.

### 9. Worked example: one formula, start to finish

*Question: What was 4 West's occupancy rate for January?*

1. January is rows 2–32 on the Census sheet.
2. In an empty cell, type `=SUM(` and drag over **E2:E32** (midnight census). Type `)/SUM(` and drag over **B2:B32** (staffed beds).
   Type `)` and press **Enter**.
3. Your formula reads `=SUM(E2:E32)/SUM(B2:B32)`. Format the cell as a percentage.
4. Sanity check: 4 West has 36 beds. A census in the low 30s means occupancy around 90%. If you got 9,000% or 0.9%, recheck your
   ranges.

That last step matters. **Always ask whether the answer is plausible.** Formulas are fast, but a wrong range gives a wrong answer
just as fast.

### 10. Shortcuts and version notes

| Action | Windows | Mac |
|---|---|---|
| Finish / cancel a formula | Enter / Esc | Return / Esc |
| Accept a function from the AutoComplete list | Tab | Tab |
| AutoSum | Alt + = | ⌘ + Shift + T |
| Show or hide formulas on the whole sheet | Ctrl + `` ` `` | Control + `` ` `` |
| Percentage format | Ctrl + Shift + % | Control + Shift + % |
| Fill down | Ctrl + D | ⌘ + D |
| Edit the active cell and color-code its references | F2 | Control + U (or Fn + F2) |
| Undo | Ctrl + Z | ⌘ + Z |

| Feature | Version |
|---|---|
| SUM, AVERAGE, MIN, MAX, COUNT, COUNTA, COUNTBLANK, ROUND, ROUNDUP, ROUNDDOWN | Every Excel version, including Excel for the web |
| **fx** (Insert Function) | Every version. On a Mac it opens the **Formula Builder** pane instead of a dialog |
| **Formulas → Show Formulas** (the button version of Ctrl + `` ` ``) | Every current desktop version, Windows and Mac |
| CEILING.MATH and FLOOR.MATH (Lesson 2.4) | Excel 2013 or later on Windows, Excel 2016 or later on a Mac |

## 🧪 Hands-on practice

Download [`1.4-basic-formulas.xlsx`](1.4-basic-formulas.xlsx) and open the **Practice** sheet. Type each answer in the yellow cell, as a
formula wherever possible. The **Check** column turns green when you're right.

<!-- BEGIN GENERATED: practice -->
All tasks use the Census sheet (4 West, Q1 2025) unless they say otherwise. Use cell ranges like Census!C2:C91, or click and drag to select them while typing a formula.

| # | Task | Hint |
|:-:|------|------|
| 1 | How many patients were admitted to 4 West during the quarter (total of the Admissions column)? | SUM |
| 2 | What was the average midnight census? (Keep full precision. The check compares the value to 2 decimal places.) | AVERAGE |
| 3 | What was the highest midnight census on any day? | MAX |
| 4 | What was the lowest midnight census on any day? | MIN |
| 5 | How many days of census data are there? Count the dates in the CensusDate column. | COUNT counts numbers, and dates are numbers |
| 6 | On how many days did the charge nurse write a note? | COUNTA counts anything that isn't empty |
| 7 | On how many days was the Notes cell left empty? | COUNTBLANK |
| 8 | Patient days = the sum of every day's midnight census. How many patient days did 4 West provide in Q1? | It's a SUM |
| 9 | What was the unit's occupancy rate for the whole quarter? Divide total patient days by total staffed-bed days. Enter it as a percentage. | SUM(…)/SUM(…), then format as % |
| 10 | In the Census sheet, fill the yellow Occupancy column with a formula for each day (MidnightCensus ÷ StaffedBeds). The gray cell counts the days your column shows 100% or more. (Type it in G2, then double-click the fill handle to copy it down.) | Relative references shift down as you copy |
| 11 | Average length of stay (ALOS) = patient days ÷ discharges. Calculate it for the quarter, rounded to 1 decimal place with ROUND. | ROUND(number, 1) |
| 12 | Without typing it into Excel first, what does =(30-6)/4+2^3 return? Then type it to confirm. | Parentheses → exponents → × ÷ → + − |
| 13 | Switch to the Supplies sheet. Fill the yellow StockValue column with UnitCost × QtyOnHand for every item. The gray cell totals your column. What is the total value of 4 West's supply room? | Multiply with *; copy down |
<!-- END GENERATED: practice -->

## ✅ Answer key

The workbook has a hidden **Answer Key** sheet (right-click any sheet tab → **Unhide…** → *Answer Key*). Its *Live result* column
runs each sample formula, so you can see it working. The same answers are below, collapsed so you don't see them by accident.

<!-- BEGIN GENERATED: answers -->
<details>
<summary><b>🔑 Show the answer key</b> (try every task before you open this)</summary>

**1. How many patients were admitted to 4 West during the quarter (total of the Admissions…**

- **Answer:** 699
- **Solution:** `=SUM(Census!C2:C91)`

SUM adds every number in the range. Typing =SUM( and then dragging over the column fills in the range for you.

**2. What was the average midnight census? (Keep full precision. The check compares the…**

- **Answer:** 32.86
- **Solution:** `=AVERAGE(Census!E2:E91)`

AVERAGE = SUM ÷ COUNT of the numbers in the range. This is the unit's *average daily census* (ADC).

**3. What was the highest midnight census on any day?**

- **Answer:** 37
- **Solution:** `=MAX(Census!E2:E91)`

**4. What was the lowest midnight census on any day?**

- **Answer:** 28
- **Solution:** `=MIN(Census!E2:E91)`

**5. How many days of census data are there? Count the dates in the CensusDate column.**

- **Answer:** 90
- **Solution:** `=COUNT(Census!A2:A91)`

Excel stores dates as numbers (serial numbers), so COUNT includes them. COUNT ignores text and blanks.

**6. On how many days did the charge nurse write a note?**

- **Answer:** 21
- **Solution:** `=COUNTA(Census!F2:F91)`

COUNTA counts every non-empty cell, including text. COUNT would return 0 here because the notes are text.

**7. On how many days was the Notes cell left empty?**

- **Answer:** 69
- **Solution:** `=COUNTBLANK(Census!F2:F91)`

COUNTA + COUNTBLANK always equals the number of cells in the range (here 90).

**8. Patient days = the sum of every day's midnight census. How many patient days did 4…**

- **Answer:** 2,957
- **Solution:** `=SUM(Census!E2:E91)`

Each patient in a bed at midnight counts as one patient day. Patient days drive staffing and cost per day.

**9. What was the unit's occupancy rate for the whole quarter? Divide total patient days by…**

- **Answer:** 91.3%
- **Solution:** `=SUM(Census!E2:E91)/SUM(Census!B2:B91)`

A rate for a whole period should be total numerator ÷ total denominator. Format the cell as a percentage (Ctrl + Shift + %, Mac: Control + Shift + %) to see 91.3% instead of 0.9126…

**10. Daily occupancy column (days at or over 100%)**

- **Answer:** 6
- **Solution:** `=E2/B2`

Type =E2/B2 in the first Occupancy cell and copy it down. Each row's formula points to its own row (=E3/B3, =E4/B4 …). Because the data is an Excel Table, typing the formula in one cell may fill the whole column automatically, and you may see it written as =[@MidnightCensus]/[@StaffedBeds]. Both are correct.

**11. Average length of stay (ALOS) = patient days ÷ discharges. Calculate it for the…**

- **Answer:** 4.2
- **Solution:** `=ROUND(SUM(Census!E2:E91)/SUM(Census!D2:D91),1)`

ROUND changes the stored value, not just how it looks. The check here is strict: an unrounded value won't match.

**12. Without typing it into Excel first, what does =(30-6)/4+2^3 return? Then type it to…**

- **Answer:** 14
- **Solution:** `=(30-6)/4+2^3`

(30−6)=24 first (parentheses), then 2^3=8 (exponent), then 24/4=6 (division), then 6+8=14.

**13. StockValue column (total supply value)**

- **Answer:** 31,730.68
- **Solution:** `=E2*F2`

One formula, copied down, gives each item's value, and then SUM adds up the column. (The live formula in the key uses SUMPRODUCT, which multiplies and adds in one step. You'll meet it in Lesson 2.4.)

</details>
<!-- END GENERATED: answers -->

## 🏆 Bonus challenge

<!-- BEGIN GENERATED: bonus -->
The Chief Nursing Officer says February felt 'impossibly full' on 4 West and asks whether the unit needs more staffed beds. Hospitals often plan beds so average occupancy is about 85%, which leaves enough slack to absorb surges. February is rows 33–60 of the Census sheet.

Work on the **Bonus** sheet of the workbook.

- **B1.** What was February's average daily census (ADC)? Use only the February rows. *(Hint: AVERAGE over rows 33–60)*
- **B2.** How many beds would 4 West have needed in February for that ADC to equal 85% occupancy? (ADC ÷ 0.85, keep the decimals.) *(Hint: If ADC is 85% of the beds, beds = ADC ÷ 0.85)*
- **B3.** You can't open part of a bed. How many EXTRA beds (beyond today's 36) should the unit open? Round up to a whole bed. *(Hint: ROUNDUP works like ROUND but always rounds up: ROUNDUP(number, 0))*
- **B4.** Sanity check: on how many February days did the midnight census exceed 85% of 36 beds (i.e. more than 30.6 patients)? *(Hint: COUNTIF(range, ">30.6"), a preview of Lesson 2.5)*
<!-- END GENERATED: bonus -->

<!-- BEGIN GENERATED: bonus-answers -->
<details>
<summary><b>🔑 Show the bonus solution</b> (give it a real try first)</summary>

**B1. What was February's average daily census (ADC)? Use only the February rows.**

- **Answer:** 32.82
- **Solution:** `=AVERAGE(Census!E33:E60)`

Select only the February rows. Selecting the whole column would give the quarterly ADC instead.

**B2. How many beds would 4 West have needed in February for that ADC to equal 85%…**

- **Answer:** 38.61
- **Solution:** `=AVERAGE(Census!E33:E60)/0.85`

**B3. You can't open part of a bed. How many EXTRA beds (beyond today's 36) should the unit…**

- **Answer:** 3
- **Solution:** `=ROUNDUP(AVERAGE(Census!E33:E60)/0.85-Census!B33,0)`

Beds needed (≈38.6) minus the 36 staffed beds ≈ 2.6, rounded up to 3. ROUND would have given 3 here as well, but only ROUNDUP guarantees you never under-build: 2.1 extra beds still means opening 3.

**B4. Sanity check: on how many February days did the midnight census exceed 85% of 36 beds…**

- **Answer:** 25
- **Solution:** `=COUNTIF(Census!E33:E60,">"&0.85*36)`

COUNTIF counts cells that meet a condition, and you'll master it in Lesson 2.5. Here it confirms that almost every February day ran above the 85% planning target.

</details>
<!-- END GENERATED: bonus-answers -->

## Key takeaways

- Every formula starts with `=`. Reference cells instead of typing numbers so results update automatically.
- Excel follows PEMDAS. Use parentheses to make your intent explicit.
- **SUM, AVERAGE, MIN, MAX** summarize numbers. **COUNT** counts numbers (including dates), **COUNTA** counts anything non-empty, and
  **COUNTBLANK** counts empties.
- Period rates are *total ÷ total*: occupancy = SUM(census) ÷ SUM(beds), ALOS = patient days ÷ discharges.
- Write one formula, then copy it down with the fill handle. References adjust row by row.
- `ROUND` changes the value. Formatting changes only how it looks.

<!-- BEGIN GENERATED: nav -->
---

⬅️ **Previous:** [1.3 Formatting Cells & Number Formats](../03-formatting-cells/README.md) · 🏠 [Course home](../../README.md) · **Next:** [1.5 Relative, Absolute & Mixed References](../05-cell-references/README.md) ➡️
<!-- END GENERATED: nav -->
