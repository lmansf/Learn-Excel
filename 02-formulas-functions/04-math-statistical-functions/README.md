# Lesson 2.4 · Math & Statistical Functions

> **Level:** Beginner → Intermediate · **Time:** about 100 minutes · **Workbook:** [`2.4-math-statistical-functions.xlsx`](2.4-math-statistical-functions.xlsx)
> **Data:** All 573 inpatient stays discharged from Bluestone Memorial Hospital in Q4 2025 (October–December), with age, length of stay, expected length of stay, and charges. All 785 emergency department visits at Bluestone Memorial in November–December 2025, with wait and length-of-stay minutes. The 388 medication orders written for that quarter's Medical-Surgical 5 East stays.

A draft board report says Bluestone Memorial's average inpatient charge last quarter was \$32,042. That figure is correct, and
it's also misleading, because about two out of three stays were charged less than that. Hospitals run on numbers like this one: average
length of stay, the 90th-percentile ED wait, the cost of a pharmacy order, the number of vials to pull for a dose. Each one needs
the right function. Pharmacy can't open half a vial, so it has to round up. A typical patient's experience is better told by the
median than the mean. In this lesson you'll learn the math functions that round and calculate the way real operations require,
and the statistical functions that tell the truth about skewed healthcare data.

## What you'll learn

- Round correctly with ROUND, ROUNDUP, ROUNDDOWN, MROUND, CEILING.MATH, and FLOOR.MATH
- Use INT, TRUNC, MOD, ABS, and SUMPRODUCT
- Describe data with MEDIAN, MODE, STDEV, PERCENTILE, and QUARTILE — and know when the mean misleads
- Rank and pick extremes with RANK.EQ, LARGE, and SMALL

## 📖 Guide

The examples use the lesson workbook, so open it and try each formula as you read.

| Sheet | Columns |
|---|---|
| **Stays** (rows 2–574) | A EncounterID, B Unit, C AdmitDate, D DischargeDate, E Age, F PrimaryDxCode, G DxDescription, H ExpectedLOS, I LOSDays, J TotalCharges, K AgeBand (yellow, for task 10) |
| **ED** (rows 2–786) | A EDVisitID, B ArrivalDateTime, C ArrivalMode, D ESILevel, E ChiefComplaint, F EDDisposition, G DoorToProviderMin, H EDLOSMin |
| **Meds** (rows 2–389) | A MedOrderID, B EncounterID, C OrderDateTime, D MedicationName, E Dose, F Route, G Frequency, H DosesDispensed, I UnitCost |

**LOSDays** is the length of stay in days, counted as the number of midnights between admission and discharge. **ExpectedLOS**
is a benchmark length of stay for the patient's diagnosis. **DoorToProviderMin** is the number of minutes from ED arrival until a
provider first saw the patient. It's blank for the 16 patients who left without being seen. **EDLOSMin** is the number of
minutes from ED arrival to departure. On the Meds sheet, **DosesDispensed** is the number of doses pharmacy sent for an order, and
**UnitCost** is what the hospital paid for one dose, its **acquisition cost**.

> 📋 The three data sheets are Excel Tables named tblStays, tblED, and tblMeds. If you drag across a whole Table column while
> building a formula, Excel writes a structured reference such as `tblStays[TotalCharges]` instead of `Stays!J2:J574`. Both give
> the same results (Lesson 2.1). This guide types plain addresses, and Lesson 3.1 covers structured references.

> 💡 **Tip:** Try examples in an empty cell at least one column away from the data, such as column M on the Stays sheet. A
> formula typed in the column right next to a Table becomes a new Table column.

### 1. ROUND, ROUNDUP, and ROUNDDOWN

```
=ROUND(number, num_digits)
=ROUNDUP(number, num_digits)
=ROUNDDOWN(number, num_digits)
```

All three take the same two arguments. **num_digits** says where to round. Positive values count decimal places, zero rounds to a
whole number, and negative values round to the left of the decimal point. Stays row 6 is a stay charged \$83,723.83:

| num_digits | Rounds to | `=ROUND(J6, num_digits)` |
|:-:|---|---|
| 1 | tenths | 83,723.8 |
| 0 | a whole number | 83,724 |
| -1 | tens | 83,720 |
| -2 | hundreds | 83,700 |
| -3 | thousands | 84,000 |

Negative num_digits is how you produce the rounded figures in a board report ("about \$84,000"). `ROUND(x,-3)/1000` gives
84 for a column headed "\$ thousands."

ROUND goes to the **nearest** value. ROUNDUP and ROUNDDOWN ignore the next digit and always go one way:

| Formula | Result | Why |
|---|---|---|
| `=ROUND(J6,-2)` | 83,700 | 83,723.83 is closer to 83,700 than to 83,800 |
| `=ROUNDUP(J6,-2)` | 83,800 | ROUNDUP always moves **away from zero** |
| `=ROUNDDOWN(J6,-2)` | 83,700 | ROUNDDOWN always moves **toward zero** |
| `=ROUND(AVERAGE(I2:I574),1)` | 4.7 | The average LOS is 4.7155… days |
| `=ROUNDUP(AVERAGE(I2:I574),1)` | 4.8 | Up to the next tenth, even though the next digit is 1 |

Pick the function that matches the real-world rule, not the one that looks tidiest:

| Situation | Function | Example |
|---|---|---|
| Reporting a number | ROUND | Average LOS of 4.7 days |
| Falling short isn't acceptable (staff, beds, supplies) | ROUNDUP | 27 patients at one nurse per 5 patients: `=ROUNDUP(27/5,0)` gives 6 nurses |
| Only completed units count | ROUNDDOWN | Full doses left in a multi-dose vial, or completed years of service |

When the next digit is exactly 5, ROUND moves away from zero, so `=ROUND(2.5,0)` is 3 and `=ROUND(-2.5,0)` is -3.

> 📋 VBA's own `Round` function works differently. It rounds a 5 to the nearest even number, so `Round(2.5)` is 2. You'll meet
> this in Level 5. The worksheet ROUND function always behaves as described here.

> 💡 **Tip:** Round once, at the end. If you round every order's cost to the cent and then add them up, the total can drift a few
> cents away from the rounded total of the unrounded costs. Keep full precision through a calculation and ROUND the final result.

### 2. Rounding to a multiple: MROUND, CEILING.MATH, and FLOOR.MATH

Sometimes the answer has to be a multiple of something: 15-minute blocks, 500 mg vials, 10-year age bands. MROUND calls that
number the *multiple*, and CEILING.MATH and FLOOR.MATH call it the **significance**.

```
=MROUND(number, multiple)                       nearest multiple
=CEILING.MATH(number, [significance], [mode])   up to a multiple
=FLOOR.MATH(number, [significance], [mode])     down to a multiple
```

Significance is optional for CEILING.MATH and FLOOR.MATH and defaults to 1. The mode argument only affects negative numbers
(section 3). Here are three door-to-provider times from the ED sheet, rounded to 15-minute blocks:

| DoorToProviderMin | `=MROUND(G3,15)` | `=CEILING.MATH(G3,15)` | `=FLOOR.MATH(G3,15)` |
|---|:-:|:-:|:-:|
| 28 (row 3) | 30 | 30 | 15 |
| 35 (row 4) | 30 | 45 | 30 |
| 109 (row 5) | 105 | 120 | 105 |

- **MROUND** picks the closest multiple. A value exactly halfway rounds away from zero, so `=MROUND(22.5,15)` is 30. Use it for
  "to the nearest 15 minutes."
- **CEILING.MATH** rounds up, so any started block counts as a whole block. Use it for vials, boxes, staff, and billing blocks.
- **FLOOR.MATH** rounds down, so only completed blocks count. Use it to put values into **bands**, such as age groups or charge
  ranges.

| Question | Formula | Result |
|---|---|---|
| How many 500 mg single-dose vials for a 750 mg dose? | `=CEILING.MATH(750,500)/500` | 2 vials (1,000 mg opened, 250 mg discarded) |
| How many nurses for 27 patients at one nurse per 5? | `=CEILING.MATH(27/5)` | 6 |
| Which 5-year age band is Stays row 2 (age 86) in? | `=FLOOR.MATH(E2,5)` | 85 |
| ED row 4 arrived at 09:32. What's the nearest quarter hour? | `=MROUND(B4,"0:15")` | 11/01/2025 09:30 |
| …and the next quarter hour? | `=CEILING.MATH(B4,"0:15")` | 11/01/2025 09:45 |

CEILING.MATH returns the rounded *amount*, 1,000 mg, not the number of vials. Divide by the vial size to count vials.
`=ROUNDUP(750/500,0)` gives 2 directly, and either approach is fine.

For times, the multiple `"0:15"` is text that Excel converts to 15 minutes (15 ÷ 1,440 of a day). If the result shows a number
such as 45962.4, format the cell as a date-time (Lesson 1.3).

> ⚠️ **MROUND needs matching signs.** In Excel, `=MROUND(-2.5,1)` returns #NUM! because the number is negative and the multiple
> is positive. Write `=MROUND(-2.5,-1)` instead, which returns -3.

> ⚠️ **Times on a boundary.** Times are fractions of a day, and most fractions can't be stored exactly. A time that sits exactly on
> a block boundary, such as 09:30 with a 15-minute significance, can occasionally land in the neighboring block with CEILING.MATH
> or FLOOR.MATH. When the block must be exact, work in whole minutes:
> `=CEILING.MATH(ROUND(B4*1440,0),15)/1440`. Whole numbers of minutes, like the ones on the ED sheet, never have this problem.

> 📋 **Version note:** CEILING.MATH and FLOOR.MATH need Excel 2013 or later (Excel 2016 or later on a Mac). Older workbooks use
> `CEILING(number, significance)` and `FLOOR(number, significance)`, which require the significance argument and handle negative
> numbers less predictably. MROUND is built into every current version.

### 3. INT and TRUNC, and rounding negative numbers

```
=INT(number)                    rounds down to the next whole number
=TRUNC(number, [num_digits])    cuts off digits (num_digits defaults to 0)
```

For positive numbers INT and TRUNC agree: both turn 4.7155 into 4. For negative numbers they don't, and negative numbers show up
whenever you subtract. Stays row 11 (ENC118855) stayed 3 days against an expected 4.6, so its **difference from expected**,
`=I11-H11`, is -1.6:

| Formula | Result | What it did |
|---|:-:|---|
| `=INT(I11-H11)` | -2 | Moved down the number line to the next lower whole number |
| `=TRUNC(I11-H11)` | -1 | Cut off the decimal part |
| `=ROUND(I11-H11,0)` | -2 | Went to the nearest whole number |

Which one is right depends on the question. "How many *full* days shorter than expected?" is -1, because the 0.6 is a partial day,
so use TRUNC. INT reports -2, which overstates it.

With num_digits, TRUNC works exactly like ROUNDDOWN: `=TRUNC(4.7155,1)` is 4.7 and `=TRUNC(83723.83,-3)` is 83,000.

The whole rounding family behaves differently below zero. This table compares every function on the same four values:

| Function | 2.4 | 2.5 | -2.4 | -2.5 | Direction |
|---|:-:|:-:|:-:|:-:|---|
| `ROUND(x,0)` | 2 | 3 | -2 | -3 | Nearest. A 5 goes away from zero |
| `ROUNDUP(x,0)` | 3 | 3 | -3 | -3 | Away from zero |
| `ROUNDDOWN(x,0)` | 2 | 2 | -2 | -2 | Toward zero |
| `TRUNC(x)` | 2 | 2 | -2 | -2 | Toward zero (same as ROUNDDOWN) |
| `INT(x)` | 2 | 2 | -3 | -3 | Down the number line |
| `CEILING.MATH(x)` | 3 | 3 | -2 | -2 | Up the number line |
| `FLOOR.MATH(x)` | 2 | 2 | -3 | -3 | Down the number line |
| `MROUND(x,1)` | 2 | 3 | #NUM! | #NUM! | Nearest multiple. Use `MROUND(x,-1)` for negatives |

The words "up" and "down" mean two different things here. In ROUNDUP and ROUNDDOWN they mean *away from zero* and *toward zero*.
In INT, CEILING.MATH, and FLOOR.MATH they mean *up or down the number line*. That's why `=ROUNDUP(-2.4,0)` is -3 but
`=CEILING.MATH(-2.4)` is -2.

The **mode** argument of CEILING.MATH and FLOOR.MATH flips their direction for negative numbers only:
`=CEILING.MATH(-2.5,1,1)` is -3 and `=FLOOR.MATH(-2.5,1,1)` is -2.

### 4. MOD: what's left over

```
=MOD(number, divisor)    the remainder after dividing number by divisor
```

MOD pairs naturally with INT. `INT(number/divisor)` is how many whole times the divisor fits, and `MOD(number, divisor)` is what's
left over. ED row 4 (ED211180) has an EDLOSMin of 473:

| Formula | Result | Meaning |
|---|---|---|
| `=INT(H4/60)` | 7 | 7 whole hours |
| `=MOD(H4,60)` | 53 | 53 minutes left over |
| `=INT(H4/60)&" h "&MOD(H4,60)&" min"` | 7 h 53 min | The & operator joins the pieces into text (Lessons 1.4 and 2.2) |
| `=TEXT(H4/1440,"[h]:mm")` | 7:53 | Another way: turn minutes into a fraction of a day and format it as a time (Lesson 2.3) |

Other everyday uses:

| Question | Formula |
|---|---|
| Is this row number even? (for shading every other row, Lesson 3.2) | `=MOD(ROW(),2)=0` |
| Is this every 3rd row? | `=MOD(ROW(),3)=0` |
| What time of day was this arrival, without the date? | `=MOD(B4,1)` gives 09:32 once formatted as a time |

> ⚠️ **MOD takes the sign of the divisor.** `=MOD(-7,60)` returns 53, not -7, and `=MOD(7,-60)` returns -53. Keep the divisor
> positive and you'll get the remainders you expect.

> 📋 `=QUOTIENT(473,60)` returns 7, the whole-number part of a division. It's the same as `TRUNC(473/60)`.

### 5. ABS: distance without direction

```
=ABS(number)    the number without its sign
```

`=ABS(I11-H11)` turns row 11's difference from expected, -1.6 days, into 1.6. That matters as soon as you summarize differences,
because positive and negative differences cancel each other out:

| Question | Formula | Result |
|---|---|---|
| On balance, do stays run longer or shorter than expected? | `=AVERAGE(I2:I574)-AVERAGE(H2:H574)` | +1.04 days (longer) |
| How far off is a typical stay, in either direction? | `=SUMPRODUCT(ABS(I2:I574-H2:H574))/COUNT(I2:I574)` | 1.84 days |

The first number says that stays run about a day long *on balance*. The second says a typical stay misses its benchmark by almost
two days, because a 3-day overrun and a 3-day underrun average to zero without ABS but to 3 with it.

ABS also turns a two-sided test into a single comparison. "Within 2 days of expected, either way" is `ABS(I2-H2)<=2`, which is
shorter than `AND(I2-H2>=-2, I2-H2<=2)`. To count the rows that pass over the whole column, use the `SUMPRODUCT(--(…))` pattern
from Lesson 2.1:

```
=SUMPRODUCT(--(ABS(I2:I574-H2:H574)<=2))     → 403 stays within 2 days of expected
```

### 6. SUMPRODUCT: multiply row by row, then add

```
=SUMPRODUCT(array1, [array2], ...)
```

SUMPRODUCT multiplies matching rows of two or more ranges and adds up the products. The first three medication orders show how:

| Meds row | MedicationName | DosesDispensed (H) | UnitCost (I) | H × I |
|:-:|---|:-:|:-:|--:|
| 2 | Sodium Chloride 0.9% | 16 | 1.75 | 28.00 |
| 3 | Azithromycin | 7 | 1.10 | 7.70 |
| 4 | Acetaminophen | 15 | 0.08 | 1.20 |
| | | | **Total** | **36.90** |

`=SUMPRODUCT(H2:H4,I2:I4)` returns 36.90 in one cell. Without it you'd fill a helper column with `=H2*I2` and SUM it.

- Every range must be the same size, or SUMPRODUCT returns #VALUE!.
- When you list ranges separated by commas, text and blank cells count as 0.
- You can also do arithmetic on whole ranges inside it. `=SUMPRODUCT(H2:H4*I2:I4)` gives the same 36.90. Lesson 2.1 used
  `SUMPRODUCT(--(condition))` to count, and section 5 used it with ABS. SUMPRODUCT handles that range math in every version of
  Excel without Ctrl + Shift + Enter (Mac: ⌘ + Shift + Return). With arithmetic, a text cell such as the header in row 1 turns
  the whole result into #VALUE!, so use the data rows (`H2:H389`), not the whole column.
- A comparison can act as a filter. On the Stays sheet, `(I2:I574>10)` is TRUE or FALSE for each stay. Multiplying it by the
  charges turns TRUE into 1 and FALSE into 0, so each charge is either kept (× 1) or zeroed out (× 0). Then SUMPRODUCT adds what's
  left: `=SUMPRODUCT((I2:I574>10)*J2:J574)` returns \$2,000,462.07, the total charges for stays longer than 10 days. Lesson 2.5
  does the same job with SUMIF.
- To get a **weighted average**, divide by the sum of the weights: `=SUMPRODUCT(weights, values)/SUM(weights)`. For example, the
  average cost per dose dispensed weights each order's UnitCost by its DosesDispensed.

> ⚠️ `=SUM(H2:H389)*SUM(I2:I389)` is not the same thing. Multiplying two totals pairs every order's doses with every other order's
> price, and gives a meaningless, much larger number.

> 📋 In Microsoft 365 and Excel 2021 or later, `=SUM(H2:H389*I2:I389)` also works, because Excel calculates the range math
> automatically. SUMPRODUCT works everywhere.

### 7. Precision versus display

Formatting and rounding look alike but do different things (Lesson 1.3):

| | Number format (Home → Decrease Decimal) | ROUND function |
|---|---|---|
| Changes | What the cell *shows* | What the cell *stores* |
| `AVERAGE(I2:I574)` formatted to 1 decimal | Shows 4.7, stores 4.71553… | `=ROUND(AVERAGE(I2:I574),1)` shows and stores 4.7 |
| Later formulas use | The full stored value | The rounded value |

That difference explains the classic complaint that "the totals don't add up." Split a \$100.00 supply charge evenly across three
cost centers and each cell shows \$33.33, yet their total shows \$100.00, not the \$99.99 that the visible cells add up to. Neither is
wrong. Decide which you need: round each piece with ROUND (and then put the leftover cent somewhere on purpose), or keep full
precision and accept that the displayed pieces may not add up exactly.

> ⚠️ Avoid **Set precision as displayed** (File → Options → Advanced → *When calculating this workbook*, or on a Mac, Excel →
> Settings → Calculation, called Preferences in older versions). It permanently rounds every stored number to its displayed format, and Undo can't bring the decimals
> back.

Excel stores numbers in binary with about 15 significant digits, so some decimals can't be stored exactly. Row 11's difference
from expected, 3 − 4.6, is stored as -1.5999999999999996 even though every display shows -1.6. This almost never matters. It can matter when a
result sits exactly on a boundary, such as an equality test or INT on a value that should be a whole number. In those cases ROUND
first, for example `=ROUND(I11-H11,1)=-1.6`.

### 8. The center: mean, median, mode, and trimmed mean

| Measure | Function | What it is |
|---|---|---|
| **Mean** | `AVERAGE` | The sum divided by the count |
| **Median** | `MEDIAN` | The middle value when the values are sorted (the average of the two middle values when the count is even) |
| **Mode** | `MODE.SNGL` | The value that occurs most often |
| **Trimmed mean** | `TRIMMEAN` | The mean after dropping a share of the highest and lowest values |

Here they are on two lesson columns:

| Measure | Formula | LOSDays (Stays) | DoorToProviderMin (ED) |
|---|---|:-:|:-:|
| Mean | `=AVERAGE(range)` | 4.72 days | 49.2 min |
| Median | `=MEDIAN(range)` | 4 days | 38 min |
| Mode | `=MODE.SNGL(range)` | 3 days | 28 min |
| 10% trimmed mean | `=TRIMMEAN(range,0.1)` | 4.42 days | 45.9 min |

Both columns are **right-skewed**: most values are small, and a long tail of large values stretches to the right. A stay can't be
shorter than 0 days, but one lasted 34. No wait was shorter than 1 minute, but one lasted 276. Every value in that tail pulls the
mean up, while the median only cares about which value sits in the middle. That's why, in right-skewed data, the mean is usually
larger than the median, which is usually larger than the mode.

**When the mean misleads.** If a manager asks "how long does a typical patient wait?", the honest answer is the median, 38 minutes.
The mean, 49 minutes, is longer than what 62% of patients actually waited. Most healthcare measures are skewed this way:
length of stay, charges, wait times, turnaround times. Report the median, and add a high percentile (section 10) to describe the
tail.

| Use | When | Example |
|---|---|---|
| `AVERAGE` | The data is roughly symmetric, or you need to rebuild a total (mean × count = total) | Budgeting: average charge per stay × number of stays |
| `MEDIAN` | The data is skewed, or the question is "what's typical?" | Length of stay, charges, ED waits |
| `MODE.SNGL` | Codes, categories, or small whole-number scales | The most common ESI level or LOS |
| `TRIMMEAN` | You want a mean that a few extreme values can't drag around | Comparing units with a few unusual stays |

**MODE details.** The mode is the right "average" for codes. ESI triage levels run from 1 to 5, and their mean, 3.12, isn't a
level at all. A few things to know:

- `MODE.SNGL` returns #N/A when no value repeats. `=MODE.SNGL(J2:J574)` on charges returns #N/A, because no two stays were charged
  exactly the same amount to the cent. The mode is only meaningful for values that repeat.
- When two values tie for most frequent, MODE.SNGL returns whichever appears first in the range.
- `MODE.MULT` returns every tied mode. In Microsoft 365 and Excel 2021 or later it spills them into several cells (Lesson 4.1).
- The older `MODE` function works exactly like MODE.SNGL.

**TRIMMEAN details.**

```
=TRIMMEAN(array, percent)
```

`percent` is the *total* share to drop, split evenly between the top and the bottom. Excel rounds the number of dropped values
down to an even number. With 573 stays and 10%, that's 57.3, rounded down to 56, so TRIMMEAN drops the 28 shortest and the 28
longest stays and averages the rest.

**What these functions skip.** AVERAGE, MEDIAN, MODE.SNGL, and TRIMMEAN ignore blank cells, text, and TRUE/FALSE values in cells,
but they include zeros. The 16 blank DoorToProviderMin cells (patients who left without being seen) are correctly left out.

> ⚠️ Never type 0 to mean "no value." If someone typed 0 into those 16 blank cells, every statistic on the column would drop, and
> the ED would look faster than it is.

### 9. The spread: range, standard deviation, and variance

Two units can share a median length of stay and still feel completely different to manage, because one is predictable and the
other swings. **Spread** measures that.

- The **range** is `=MAX(I2:I574)-MIN(I2:I574)`, 34 days here. It's easy to explain, but a single extreme value decides it.
- The **standard deviation** is, roughly, the typical distance between a value and the mean. A bigger standard deviation means the
  values are more spread out. `=STDEV.S(I2:I574)` returns 3.21 days.
- The **variance** is the standard deviation squared, so `=VAR.S(I2:I574)` returns 10.30 in "days squared." Statistical formulas
  use it, but you'll report the standard deviation because it's in the same units as the data.

| Function | Divides by | Use when |
|---|---|---|
| `STDEV.S`, `VAR.S` | n − 1 | The data is a **sample** of a larger process. Most hospital data is: this quarter's stays stand for "our stays in general" |
| `STDEV.P`, `VAR.P` | n | The data is the entire **population** and you care only about it |

On 573 stays the two barely differ: `STDEV.S` gives 3.209 and `STDEV.P` gives 3.206. On a small sample the gap is bigger. When in
doubt, use STDEV.S. The older `STDEV` and `VAR` functions are the same as STDEV.S and VAR.S, and `STDEVP` is the same as STDEV.P.

To compare spread between measures with different units or scales, divide the standard deviation by the mean. This is the
**coefficient of variation**. `=STDEV.S(I2:I574)/AVERAGE(I2:I574)` gives 0.68, so LOS varies by about 68% of its mean.

**The standard deviation describes bell-shaped data best.** For data shaped like a bell curve, about 95% of values fall within
2 standard deviations of the mean. Try it on LOS: 4.72 − 2 × 3.21 is -1.70 days, which is impossible. When the mean minus two
standard deviations goes below the smallest possible value, the data is skewed, and percentiles describe it better.

### 10. Percentiles, quartiles, and percent rank

A **percentile** is the value that a given share of the data falls at or below. The 90th percentile of ED waits is the wait that
90% of patients didn't exceed.

```
=PERCENTILE.INC(array, k)          k from 0 to 1, such as 0.9 for the 90th percentile
=QUARTILE.INC(array, quart)        quart is 0, 1, 2, 3, or 4
=PERCENTRANK.INC(array, x)         the reverse: what percentile is the value x?
```

**Quartiles** split sorted data into four equal parts. QUARTILE.INC is a shortcut for the matching percentile:

| quart | Returns | Same as |
|:-:|---|---|
| 0 | The minimum | `MIN`, `PERCENTILE.INC(array,0)` |
| 1 | **Q1**, the 25th percentile | `PERCENTILE.INC(array,0.25)` |
| 2 | The median | `MEDIAN`, `PERCENTILE.INC(array,0.5)` |
| 3 | **Q3**, the 75th percentile | `PERCENTILE.INC(array,0.75)` |
| 4 | The maximum | `MAX`, `PERCENTILE.INC(array,1)` |

On the Stays sheet, `=QUARTILE.INC(I2:I574,1)` is 3 and `=QUARTILE.INC(I2:I574,3)` is 6, so the middle half of stays lasted 3 to 6
days. The distance between them, Q3 − Q1, is the **interquartile range (IQR)**. It's 3 days here. The IQR is the spread of the
middle half of the data, so a few extreme values can't stretch it the way they stretch the standard deviation. On the ED sheet,
`=PERCENTILE.INC(G2:G786,0.75)` is 68 minutes and `=PERCENTILE.INC(G2:G786,0.95)` is 126 minutes.

**How Excel calculates a percentile.** Take five waits sorted from shortest to longest: 22, 35, 41, 58, and 90 minutes. For the
90th percentile, PERCENTILE.INC finds position 1 + 0.9 × (5 − 1) = 4.6. That's 60% of the way from the 4th value (58) to the 5th
(90), so the result is 58 + 0.6 × (90 − 58) = **77.2**. A percentile can be a value that isn't in the data, and that's normal.

There are two versions of each function:

| | PERCENTILE.INC, QUARTILE.INC | PERCENTILE.EXC, QUARTILE.EXC |
|---|---|---|
| Position for percentile k | 1 + k × (n − 1) | k × (n + 1) |
| Allows k = 0 and k = 1 | Yes (they return MIN and MAX) | No (#NUM!) |
| 25th percentile of the five waits | 35 | 28.5 |
| 90th percentile of the five waits | 77.2 | #NUM! (too few values for that percentile) |
| Older function that matches it | `PERCENTILE`, `QUARTILE` | None |

With hundreds of rows the two methods give nearly identical answers. Use .INC unless you need to match a report built with the
exclusive method. Other statistics programs choose among several methods, so small differences from Excel are normal.

**PERCENTRANK.INC** answers the opposite question. `=PERCENTRANK.INC(G2:G786,60)` returns 0.699, so a 60-minute wait was longer
than about 70% of the other waits. By default it returns three decimal places, and it cuts off the rest instead of rounding.

**Blanks and errors.** The percentile and quartile functions skip blank cells, so the patients who left without being seen are
left out automatically. A single error value in the range, such as #N/A, makes the whole result an error. `AGGREGATE` (Lesson 1.6)
can skip errors: `=AGGREGATE(16,6,G2:G786,0.9)` is the 90th percentile ignoring errors, where 16 means PERCENTILE.INC and 6 means
"ignore error values." Function 12 is MEDIAN, 14 is LARGE, 15 is SMALL, and 17 is QUARTILE.INC.

### 11. Ranking and picking extremes: RANK.EQ, RANK.AVG, LARGE, and SMALL

```
=RANK.EQ(number, ref, [order])    the position of number in ref
=RANK.AVG(number, ref, [order])   the same, but ties get their average position
=LARGE(array, k)                  the k-th largest value
=SMALL(array, k)                  the k-th smallest value
```

**RANK.EQ** finds where a value stands. Leave out `order` (or use 0) to rank the largest value as 1. Set `order` to 1 to rank
the smallest value as 1 instead. Stays row 2 (ENC118816) was charged \$32,556.40:

| Formula | Result | Meaning |
|---|:-:|---|
| `=RANK.EQ(J2,J2:J574)` | 191 | The 191st most expensive stay |
| `=RANK.EQ(J2,J2:J574,1)` | 383 | The 383rd cheapest stay |

**Ties.** Twenty-eight stays lasted exactly 9 days, and 35 stays were longer. `=RANK.EQ(9,I2:I574)` gives all 28 of them rank 36.
The next rank used is 64, because ranks 37 through 63 are "used up" by the tie. RANK.AVG gives the tied stays the average of
positions 36 through 63 instead:

| | RANK.EQ | RANK.AVG |
|---|---|---|
| Tied values get | The best position in the tie | The average position in the tie |
| The 9-day stays | 36 | 49.5 |
| Typical use | Leaderboards and "top 10" lists | Statistics that need average ranks |

> ⚠️ **Lock the ref before you copy.** To rank every stay in a helper column, write `=RANK.EQ(I2,$I$2:$I$574)` and fill it down.
> Without the dollar signs the range slides down with each row (Lesson 1.5). Press F4 (Mac: ⌘ + T) to add them. The ref must be a
> range of cells, not a calculated array.

**LARGE and SMALL** work the other way around: you give the position and get the value.

| Formula | Result |
|---|---|
| `=LARGE(G2:G786,1)` | 276, the longest door-to-provider wait (the same as MAX) |
| `=LARGE(G2:G786,2)` | 211, the 2nd-longest |
| `=LARGE(G2:G786,3)` | 196, the 3rd-longest |
| `=SMALL(I2:I574,1)` | 0, the shortest stay (the same as MIN) |
| `=SMALL(I2:I574,2)` | 1, the 2nd-shortest |

Duplicates count separately. Thirty stays lasted 1 day, so `SMALL(I2:I574,k)` returns 1 for every k from 2 through 31.

| Question | Function |
|---|---|
| "Where does this stay rank?" | RANK.EQ |
| "What is the 3rd-highest value?" | LARGE |
| "What is the 3rd-lowest value?" | SMALL |

> 💡 **Tip:** To average the three longest waits in one formula, give LARGE a list of positions:
> `=AVERAGE(LARGE(G2:G786,{1,2,3}))` returns 227.7 minutes. The braces `{1,2,3}` are an **array constant**, a list typed straight
> into the formula, and this works in every version of Excel. Lesson 4.1 builds complete top-N lists with SORT and TAKE.

### 12. Worked example: profiling length of stay and spotting outliers

*Question: How long do Bluestone Memorial's inpatients stay, and which stays are unusually long?* Use the whole Stays sheet. Type
each label in column M and its formula in column N, starting in row 2, so every number has a name next to it.

| Cell | Statistic | Formula | Result |
|:-:|---|---|:-:|
| N2 | Stays | `=COUNT(I2:I574)` | 573 |
| N3 | Mean | `=AVERAGE(I2:I574)` | 4.72 |
| N4 | Median | `=MEDIAN(I2:I574)` | 4 |
| N5 | Mode | `=MODE.SNGL(I2:I574)` | 3 |
| N6 | 10% trimmed mean | `=TRIMMEAN(I2:I574,0.1)` | 4.42 |
| N7 | Standard deviation | `=STDEV.S(I2:I574)` | 3.21 |
| N8 | Q1 | `=QUARTILE.INC(I2:I574,1)` | 3 |
| N9 | Q3 | `=QUARTILE.INC(I2:I574,3)` | 6 |
| N10 | 90th percentile | `=PERCENTILE.INC(I2:I574,0.9)` | 9 |
| N11 | Longest stay | `=MAX(I2:I574)` | 34 |

Read it from top to bottom. The mean is above the median, which is above the mode, so LOS is right-skewed. A typical stay is
4 days, and the middle half of stays lasted 3 to 6 days. At least 90% of stays lasted 9 days or less. The longest stay is far
beyond everything else.

**An outlier** is a value far from the rest of the data. Two rules are common for flagging them:

1. **The IQR rule** (also called Tukey's fences). Values below Q1 − 1.5 × IQR or above Q3 + 1.5 × IQR are outliers. Put the IQR
   in N12 with `=N9-N8`, which gives 6 − 3 = 3 days. Put the upper fence in N13 with `=N9+1.5*N12`, which gives
   6 + 1.5 × 3 = **10.5 days**. Then `=COUNTIF(I2:I574,">"&N13)` counts **27** stays longer than the fence. The lower fence is
   3 − 4.5 = -1.5 days, which no stay can be below.
2. **The mean ± 3 standard deviations rule.** The upper limit is 4.716 + 3 × 3.209 = **14.34 days** (`=N3+3*N7`), and only
   **7** stays exceed it.

`">"&N13` joins the operator to the cell's value to build the criterion text ">10.5". COUNTIF comes from Lesson 2.1, and Lesson
2.5 covers criteria like this one fully.

Why do the rules disagree? The long stays themselves inflate the mean and the standard deviation, and that raises the bar they
have to clear. The extreme values partly hide themselves. Quartiles barely move when a few values change, so the IQR rule is the
standard screen for skewed data like LOS and charges. The 3-standard-deviation rule works best for roughly bell-shaped data.

> ⚠️ An outlier isn't an error. A 34-day stay is probably a real patient with complications. Flag outliers for review and report
> them separately, but don't delete them.

### 13. Working efficiently, and version notes

- **Find functions by category:** **Formulas → Math & Trig** lists ROUND, MROUND, SUMPRODUCT, and the rest of sections 1–6.
  **Formulas → More Functions → Statistical** lists sections 8–11.
- **AutoSum's dropdown** (**Home → AutoSum ▾**) inserts Average, Count Numbers, Max, and Min. **Alt + =** (Mac: **⌘ + Shift + T**)
  inserts SUM directly.
- **Shift + F3** opens the Insert Function dialog on Windows or the Formula Builder on a Mac, with a box for each argument.
- **Check your formula against the status bar.** Select a range and the status bar shows its Average, Count, and Sum. Right-click
  the status bar (Mac: Control-click) to add Minimum, Maximum, and Numerical Count.
- **The Analysis ToolPak** can produce a whole table of descriptive statistics at once. Lesson 4.5 covers it.

| Functions | Available in |
|---|---|
| ROUND, ROUNDUP, ROUNDDOWN, INT, TRUNC, MOD, ABS, SUMPRODUCT, AVERAGE, MEDIAN, TRIMMEAN, LARGE, SMALL, MROUND | Every current version |
| STDEV.S, STDEV.P, VAR.S, VAR.P, MODE.SNGL, MODE.MULT, PERCENTILE.INC/.EXC, QUARTILE.INC/.EXC, PERCENTRANK.INC/.EXC, RANK.EQ, RANK.AVG, AGGREGATE | Excel 2010 and later (Excel 2011 and later on a Mac) |
| CEILING.MATH, FLOOR.MATH | Excel 2013 and later (Excel 2016 and later on a Mac) |
| STDEV, STDEVP, VAR, MODE, PERCENTILE, QUARTILE, PERCENTRANK, RANK, CEILING, FLOOR | Older "compatibility" versions. They still work, but prefer the newer names |

Microsoft 365, Excel 2021 and later, and Excel for the web have every function in this lesson.

## 🧪 Hands-on practice

Download [`2.4-math-statistical-functions.xlsx`](2.4-math-statistical-functions.xlsx) and open the **Practice** sheet. Type each
answer in the yellow cell, as a formula wherever possible. The **Check** column turns green when you're right.

<!-- BEGIN GENERATED: practice -->
Tasks use three data sheets. Stays holds 573 inpatient stays (rows 2–574), ED holds 785 visits (rows 2–786), and Meds holds 388 medication orders (rows 2–389). Reference a column's data rows, like Stays!J2:J574, rather than a whole column like J:J, because the header text in row 1 breaks the row-by-row math in SUMPRODUCT. Structured references such as tblStays[TotalCharges] work too. When a task asks you to round, round with a function so the stored value matches, not just the display.

| # | Task | Hint |
|:-:|------|------|
| 1 | What is the most common ESI triage level among the ED visits? | MODE.SNGL |
| 2 | How much higher is the mean (average) TotalCharges per stay than the median TotalCharges? Enter the difference in dollars, to the cent. | AVERAGE(…) − MEDIAN(…) |
| 3 | What is the sample standard deviation of TotalCharges, rounded to the nearest \$100 with ROUND? | STDEV.S, then ROUND with a negative num_digits |
| 4 | ED leaders review the three longest ED visits for boarding delays, where an admitted patient waits in the ED for an inpatient bed. They report each visit's length of stay in whole hours and count any started hour as a full hour. Find the 3rd-longest EDLOSMin on the ED sheet, convert it to hours, and round it up to a whole number of hours with ROUNDUP. | LARGE(array, k) finds the visit. Divide by 60, then ROUNDUP(…, 0) |
| 5 | The charge-capture team audits the cheapest stays, because an unusually low charge often means some charges were never posted. What is the 3rd-lowest TotalCharges? | SMALL is LARGE's mirror image |
| 6 | Encounter ENC119052 (Stays row 70) stayed 10 days. Rank its LOSDays among all 573 stays with RANK.EQ, where the longest stay is rank 1. | RANK.EQ(number, ref). Leaving out order ranks the largest value as 1 |
| 7 | Ninety percent of the patients who saw a provider waited at most how many minutes? Calculate the 90th percentile of DoorToProviderMin. (The 16 blank cells are patients who left without being seen.) | PERCENTILE.INC(array, 0.9) |
| 8 | What is the interquartile range (IQR = Q3 − Q1) of EDLOSMin, in minutes? Use QUARTILE.INC. | QUARTILE.INC(array, 3) − QUARTILE.INC(array, 1) |
| 9 | What was the total acquisition cost of the medications dispensed for the 5 East stays? Multiply each order's DosesDispensed by its UnitCost and add up all the orders, in one formula. Enter the total in dollars, to the cent. | SUMPRODUCT(array1, array2) |
| 10 | Fill the yellow AgeBand column on the Stays sheet with each stay's 10-year age band using FLOOR.MATH, so age 78 becomes 70 and age 80 becomes 80. Type the formula in K2. Stays is an Excel Table, so Excel fills the formula down the whole column for you. If it doesn't, double-click the fill handle. The gray cell counts the stays in the 80–89 band. | FLOOR.MATH(number, significance) rounds down to a multiple |
| 11 | Order RX824601 (Meds row 255) is vancomycin 1,250 mg every 12 hours, with 24 doses dispensed. Assume the IV room mixes every dose from 500 mg single-dose vials and throws away whatever is left in an opened vial. How many vials did this order use? | Work out vials per dose first: round 1,250 mg UP to a whole number of 500 mg vials with CEILING.MATH (or ROUNDUP) |
| 12 | Case managers call a stay "on target" when its LOSDays is within 1 day of its ExpectedLOS in either direction (a difference of 1.0 day or less, longer or shorter). How many of the 573 stays were on target? Use one formula. | ABS makes −0.9 and +0.9 the same distance. Count TRUE results with SUMPRODUCT(--(…)) from Lesson 2.1 |
| 13 | The ED status board rounds each visit's length of stay to the nearest 15 minutes and shows it as hours and minutes, like "6 h 45 min". What does it show for visit ED211176 (ED row 2, EDLOSMin = 624)? Build the text with one formula that refers to the EDLOSMin cell. | MROUND to 15 first. Then INT(minutes/60) gives hours and MOD(minutes,60) gives the minutes left over. Join with & |
<!-- END GENERATED: practice -->

## ✅ Answer key

The workbook has a hidden **Answer Key** sheet (right-click any sheet tab → **Unhide…** → *Answer Key*). Its *Live result* column
runs each sample formula, so you can see it working. The same answers are below, collapsed so you don't see them by accident.

<!-- BEGIN GENERATED: answers -->
<details>
<summary><b>🔑 Show the answer key</b> — Try every task before opening this.</summary>

**1. What is the most common ESI triage level among the ED visits?**

- **Answer:** 3
- **Solution:** `=MODE.SNGL(ED!D2:D786)`

MODE.SNGL returns the value that occurs most often: level 3 appears 328 times out of 785. ESI levels are codes on a 1–5 scale, so their mean (3.12) isn't a real triage level. For codes and categories, the mode is the honest "typical" value. The older `MODE` function gives the same result.

**2. How much higher is the mean (average) TotalCharges per stay than the median…**

- **Answer:** 8,076.03
- **Solution:** `=AVERAGE(Stays!J2:J574)-MEDIAN(Stays!J2:J574)`

The mean is \$32,041.72 but the median is \$23,965.69, so half of all stays were charged less than \$23,966. In fact 375 of the 573 stays (65%) fall below the mean. A handful of very long, very expensive stays pull the mean up, while the median only cares about the middle value. When data is skewed like this, the median describes a typical stay better.

**3. What is the sample standard deviation of TotalCharges, rounded to the nearest \$100…**

- **Answer:** 24,100
- **Solution:** `=ROUND(STDEV.S(Stays!J2:J574),-2)`

STDEV.S returns \$24,056.39. A num_digits of -2 rounds to the hundreds place, giving \$24,100. STDEV.S treats the stays as a sample of an ongoing process, which is the usual choice. STDEV.P gives \$24,035.39, which rounds to \$24,000, so the check tells the two apart.

**4. ED leaders review the three longest ED visits for boarding delays, where an admitted…**

- **Answer:** 20
- **Solution:** `=ROUNDUP(LARGE(ED!H2:H786,3)/60,0)`

LARGE returns the k-th largest value, so `LARGE(range,1)` is the same as MAX. The three longest visits lasted 1,304, 1,281, and 1,169 minutes. The 3rd is 1,169 ÷ 60 = 19.48 hours, and ROUNDUP turns the started hour into a full one, giving 20. `ROUND` and `ROUNDDOWN` both give 19 here, which would under-report the visit because the rule counts every started hour. `=CEILING.MATH(LARGE(…,3)/60)` gives the same 20. Change k to 2 or 1 to see the other two visits.

**5. The charge-capture team audits the cheapest stays, because an unusually low charge…**

- **Answer:** 6,249.27
- **Solution:** `=SMALL(Stays!J2:J574,3)`

SMALL(array, k) returns the k-th smallest value, so `SMALL(range,1)` equals MIN. Like LARGE, it ignores blank cells and text.

**6. Encounter ENC119052 (Stays row 70) stayed 10 days. Rank its LOSDays among all 573…**

- **Answer:** 28
- **Solution:** `=RANK.EQ(Stays!I70,Stays!I2:I574)`

27 stays were longer than 10 days, so this one ranks 28. The other 7 stays of exactly 10 days share rank 28, and the next rank used is 36. `RANK.AVG` returns 31.5 instead, which is the average of positions 28–35. If you copy a RANK formula down a column, lock the ref with \$ (`$I$2:$I$574`) so it doesn't slide.

**7. Ninety percent of the patients who saw a provider waited at most how many minutes?…**

- **Answer:** 102
- **Solution:** `=PERCENTILE.INC(ED!G2:G786,0.9)`

PERCENTILE.INC ignores blank cells, so it uses only the 769 visits with a provider time. The median wait was 38 minutes, yet 74 of the 769 waits (9.6%) were longer than 102 minutes. That long tail is why ED leaders track the 90th percentile alongside the median. If someone had typed 0 into the blank cells, the percentile would drop and the report would look better than reality.

**8. What is the interquartile range (IQR = Q3 − Q1) of EDLOSMin, in minutes? Use QUARTILE.INC.**

- **Answer:** 155
- **Solution:** `=QUARTILE.INC(ED!H2:H786,3)-QUARTILE.INC(ED!H2:H786,1)`

Q1 is 167 and Q3 is 322, so the middle half of ED visits lasted between 2 h 47 min and 5 h 22 min. The IQR ignores the extreme visits at both ends, so a few boarding patients can't stretch it. QUARTILE.EXC interpolates slightly differently, but with 785 visits it lands on the same Q1 and Q3 here.

**9. What was the total acquisition cost of the medications dispensed for the 5 East stays?…**

- **Answer:** 9,252.38
- **Solution:** `=SUMPRODUCT(Meds!H2:H389,Meds!I2:I389)`

SUMPRODUCT multiplies the two columns row by row, then adds the 388 products. It gives the same result as a helper column of `=H2*I2` totaled with SUM. Multiplying two totals, as in `=SUM(H…)*SUM(I…)`, is wrong (it gives \$3,108,585.48) because it pairs every order's doses with every other order's price. Divide by `SUM(Meds!H2:H389)` to get the average cost per dose, \$3.08.

**10. AgeBand column with FLOOR.MATH (stays aged 80–89)**

- **Answer:** 44
- **Solution:** `=FLOOR.MATH(E2,10)`

FLOOR.MATH rounds down to the nearest multiple of 10, so every age from 80 to 89 lands in the 80 band. `ROUND(E2,-1)` would be wrong, because it sends ages 75–84 to 80 and gives 93 instead of 44. `=INT(E2/10)*10`, `=TRUNC(E2,-1)`, and `=ROUNDDOWN(E2,-1)` all work too. In the Table you may see `=FLOOR.MATH([@Age],10)`. The most common band is the 70s.

**11. Order RX824601 (Meds row 255) is vancomycin 1,250 mg every 12 hours, with 24 doses…**

- **Answer:** 72
- **Solution:** `=CEILING.MATH(1250,500)/500*Meds!H255`

`CEILING.MATH(1250,500)` rounds up to 1,500 mg, which is 3 vials per dose. Two vials would hold only 1,000 mg, so ROUNDDOWN would under-dose. 3 vials × 24 doses = 72 vials. Rounding the order's total instead (`CEILING.MATH(1250*24,500)/500` = 60) assumes leftover drug carries over to the next dose, which a single-dose vial doesn't allow. `=ROUNDUP(1250/500,0)*Meds!H255` also works. The order wasted 6,000 mg in partly used vials.

**12. Case managers call a stay "on target" when its LOSDays is within 1 day of its…**

- **Answer:** 264
- **Solution:** `=SUMPRODUCT(--(ABS(Stays!I2:I574-Stays!H2:H574)<=1))`

Subtracting the two columns gives each stay's difference from expected, which is negative when the stay was shorter than expected. ABS turns every difference into a distance, the comparison turns each distance into TRUE or FALSE, `--` turns those into 1s and 0s, and SUMPRODUCT adds them. Without ABS, every stay that ended early would count as on target, and you'd get 348. A helper column of `=ABS(I2-H2)` counted with `COUNTIF(range,"<=1")` gives the same 264.

**13. The ED status board rounds each visit's length of stay to the nearest 15 minutes and…**

- **Answer:** 10 h 30 min
- **Solution:** `=INT(MROUND(ED!H2,15)/60)&" h "&MOD(MROUND(ED!H2,15),60)&" min"`

`MROUND(624,15)` returns 630, because 624 is closer to 630 than to 615. Then `INT(630/60)` is 10 whole hours and `MOD(630,60)` is the 30 minutes left over. The & operator joins the pieces into "10 h 30 min". Without the rounding the board would show "10 h 24 min". `TRUNC` or `ROUNDDOWN(…,0)` can replace INT here because the minutes are positive.

</details>
<!-- END GENERATED: answers -->

## 🏆 Bonus challenge

<!-- BEGIN GENERATED: bonus -->
The CFO's draft board report says the average inpatient charge in Q4 was \$32,042. A board member asks whether that number is typical and which stays are unusually expensive. Flag the high-charge outliers with two common rules, measure how much they matter, and find a more representative average. Use the TotalCharges column on the Stays sheet (Stays!J2:J574). Your B1 answer lands in cell D6 of this sheet, so later parts can refer to it.

Work on the **Bonus** sheet of the workbook.

- **B1.** Calculate the upper outlier fence with the IQR rule: Q3 + 1.5 × (Q3 − Q1), using QUARTILE.INC. Enter it in dollars, to the cent. *(Hint: Find Q1 and Q3 with QUARTILE.INC, then combine them)*
- **B2.** How many stays have TotalCharges above the fence from B1? *(Hint: COUNTIF(range, ">"&cell) joins the operator to your B1 cell)*
- **B3.** What share of all Q4 TotalCharges came from the stays above the fence? Enter it as a percentage with one decimal place. *(Hint: (range>D6) is TRUE or FALSE for each stay. Multiply it by the charges inside SUMPRODUCT, then divide by the total)*
- **B4.** A second common rule flags any value more than 3 standard deviations above the mean. How many stays have TotalCharges above AVERAGE + 3 × STDEV.S? (Hint: Build the threshold inside the criterion: ">"&(AVERAGE(…)+3*STDEV.S(…)))
- **B5.** Calculate a 10% trimmed mean of TotalCharges with TRIMMEAN, which drops about 5% of stays from each end before averaging. Enter it in dollars, to the cent. *(Hint: TRIMMEAN(array, percent). The percent is the TOTAL share to drop)*
<!-- END GENERATED: bonus -->

<!-- BEGIN GENERATED: bonus-answers -->
<details>
<summary><b>🔑 Show the bonus solution</b> — Give it a real try first!</summary>

**B1. Calculate the upper outlier fence with the IQR rule: Q3 + 1.5 × (Q3 − Q1), using…**

- **Answer:** 70,379.05
- **Solution:**

```
=QUARTILE.INC(Stays!J2:J574,3)+1.5*(QUARTILE.INC(Stays!J2:J574,3)-QUARTILE.INC(Stays!J2:J574,1))
```


Q1 is \$16,511.45 and Q3 is \$38,058.49, so the IQR is \$21,547.04 and the fence is \$38,058.49 + 1.5 × \$21,547.04 = \$70,379.05. The lower fence, Q1 − 1.5 × IQR, is −\$15,809.11. That's below zero, so the rule can't flag any low outliers. Right-skewed data like charges usually has outliers on the high side only.

**B2. How many stays have TotalCharges above the fence from B1?**

- **Answer:** 50
- **Solution:** `=COUNTIF(Stays!J2:J574,">"&D6)`

`">"&D6` builds the criterion text ">70379.05", so COUNTIF counts charges above your fence. That's 50 stays, or 8.7% of all stays.

**B3. What share of all Q4 TotalCharges came from the stays above the fence? Enter it as a…**

- **Answer:** 25.7%
- **Solution:** `=SUMPRODUCT((Stays!J2:J574>D6)*Stays!J2:J574)/SUM(Stays!J2:J574)`

`(Stays!J2:J574>D6)` is TRUE for an outlier and FALSE otherwise. Multiplying by the charges turns TRUE into 1 and FALSE into 0, so only the outlier charges survive, and SUMPRODUCT adds them. Just 8.7% of stays produced 25.7% of all charges. That concentration is exactly why the mean sits so far above the median. (`=SUMIF(Stays!J2:J574,">"&D6)/SUM(Stays!J2:J574)` from Lesson 2.5 gives the same answer.)

**B4. A second common rule flags any value more than 3 standard deviations above the mean.…**

- **Answer:** 13
- **Solution:** `=COUNTIF(Stays!J2:J574,">"&(AVERAGE(Stays!J2:J574)+3*STDEV.S(Stays!J2:J574)))`

The threshold is \$32,041.72 + 3 × \$24,056.39 = \$104,210.89, and only 13 stays exceed it, compared with 50 under the IQR rule. The outliers themselves inflate the mean and the standard deviation, which raises the bar they have to clear. This effect is called **masking**. The 3-SD rule assumes roughly symmetric data. Quartiles barely move when a few extreme values change, so the IQR rule is the better screen for skewed data like charges and length of stay.

**B5. Calculate a 10% trimmed mean of TotalCharges with TRIMMEAN, which drops about 5% of…**

- **Answer:** 29,241.27
- **Solution:** `=TRIMMEAN(Stays!J2:J574,0.1)`

573 × 10% = 57.3 stays. TRIMMEAN rounds that down to an even number (56) and drops 28 from each end. The result, \$29,241.27, sits between the median (\$23,965.69) and the mean (\$32,041.72). For the board, report the median as the typical stay and keep the mean for budgeting, because mean × number of stays = total charges. Then list the outlier stays separately so nobody mistakes them for the norm.

</details>
<!-- END GENERATED: bonus-answers -->

## Key takeaways

- Match the rounding function to the real-world rule. Use ROUND to report, ROUNDUP or CEILING.MATH when you can't fall short
  (vials, staff), ROUNDDOWN or FLOOR.MATH for completed units and bands, and MROUND for the nearest increment.
- Rounding changes the stored value, and formatting changes only the display.
- INT and TRUNC agree on positive numbers but not negative ones. MOD returns the remainder, which turns minutes into hours and
  minutes, and ABS turns a signed difference into a distance.
- SUMPRODUCT multiplies columns row by row and adds the products, so a total cost needs no helper column.
- Healthcare data is usually right-skewed. Report the median and a high percentile, not just the mean, and use the mode for codes.
- Use STDEV.S for samples. For skewed data the IQR and percentiles describe spread better than the standard deviation, and the IQR
  rule flags outliers more reliably than mean ± 3 standard deviations.
- RANK.EQ answers "where does this value rank?" LARGE and SMALL answer "what's the k-th value?" Tied values share a rank.

<!-- BEGIN GENERATED: nav -->
---

⬅️ **Previous:** [2.3 Dates & Times](../03-date-time-functions/README.md) · 🏠 [Course home](../../README.md) · **Next:** [2.5 Conditional Counting & Summing](../05-conditional-aggregation/README.md) ➡️
<!-- END GENERATED: nav -->
