# Lesson 4.5 · Statistics & Forecasting

> **Level:** Advanced · **Time:** about 150 minutes · **Workbook:** [`4.5-statistics-forecasting.xlsx`](4.5-statistics-forecasting.xlsx)
> **Data:** A random sample of 400 inpatient stays discharged in 2025 (200 each from Bluestone Memorial and Cedar Ridge), every Q4 2025 emergency department visit at the three hospitals with door-to-provider minutes, daily ED arrivals per hospital for 2024–2025, and 24 months of system-wide ED, LWBS, inpatient, and 30-day readmission counts.

Every month, the quality committee looks at last month's readmission rate and someone asks, *"Is that real, or just noise?"*
The CFO asks how much one extra day in hospital adds to the bill. The ED director wants to know whether Cedar Ridge really
moves patients faster than Bluestone Memorial, and the staffing office needs next month's volume before the schedule is
built. Averages and PivotTables can't answer those questions, because every one of them is about **variation**: how much
numbers bounce around on their own, and when a difference is bigger than that bounce. In this lesson you'll describe
distributions, measure relationships, test differences, put honest error bars on estimates, forecast volumes, and build the
control charts hospitals use to tell a real change from random noise.

## What you'll learn

- Summarize distributions with the Analysis ToolPak and histograms
- Measure relationships with correlation and linear regression
- Test differences with t-tests and quantify uncertainty with confidence intervals
- Forecast volumes and monitor processes with control charts

## 📖 Guide

The examples use the lesson workbook, so open it and try each formula as you read. Every worked example uses a different
column, hospital, or period from the practice tasks, so you can copy the method without seeing an answer.

### 1. The data

| Sheet | Rows | Columns |
|---|---|---|
| **Stays** | 2–401 (Bluestone Memorial 2–201, Cedar Ridge 202–401) | A EncounterID, B Facility, C DischargeDate, D Age, E DxCategory, F LOSDays, G TotalCharges |
| **Bins** | 4–10 | A LOSBin: the bin limits for the LOS histogram in Task 2 |
| **EDWaits** | 2–1629 (Bluestone Memorial 2–1157, Ashby Falls 1158–1383, Cedar Ridge 1384–1629) | A EDVisitID, B Facility, C ArrivalDateTime, D ESILevel, E DoorToProviderMin |
| **EDDaily** | 2–732 (2024 is rows 2–367, 2025 is rows 368–732) | A Date, B Weekday, C Memorial, D AshbyFalls, E CedarRidge, F AllEDs |
| **Monthly** | 2–25 (Jan 2024 to Dec 2025) | A MonthStart, B MonthNum, C EDVisits, D LWBS, E InpatientDischarges, F IndexStays, G Readmits30 |

A few definitions you'll need:

- **LOSDays** (length of stay) is the number of midnights between admission and discharge, with a minimum of 1 day. It's the
  same definition Lesson 4.2 used.
- **Age** on the Stays sheet is the patient's age in whole years on the admission date.
- **DoorToProviderMin** is the number of minutes from ED arrival until a provider first saw the patient. It's blank for
  patients who **left without being seen (LWBS)**.
- The EDDaily counts are ED arrivals per calendar day. Days with no arrivals show 0.
- On the Monthly sheet, **IndexStays** is the month's inpatient discharges minus the patients who died, because a patient who
  died can't be readmitted. **Readmits30** counts the index stays followed by another inpatient admission within 30 days, using
  the rule in the [data dictionary](../../data/README.md). Months are discharge months.

The Stays sheet is a **random sample**: 200 stays drawn at random from each hospital's 2025 discharges. Statistics is mostly
about samples. You measure some patients and use them to say something about all patients, and the tools in this lesson tell
you how far to trust that leap.

> 💡 **Tip:** Try examples in an empty column at least one column away from a data Table. A formula typed in the column right next
> to a Table becomes a new Table column, and Excel may copy it into every row.

### 2. Turn on the Analysis ToolPak

The **Analysis ToolPak** is a free add-in that ships with Excel. It adds a **Data Analysis** button that runs about twenty
statistical procedures through dialog boxes and writes the results as a table.

**Windows:**

1. Select **File → Options → Add-ins**.
2. At the bottom, set **Manage** to *Excel Add-ins* and select **Go…**.
3. Tick **Analysis ToolPak** and select **OK**. (You don't need *Analysis ToolPak – VBA*, which is for macros.)
4. A **Data Analysis** button appears at the right end of the **Data** tab.

**Mac:** select **Tools → Excel Add-ins…**, tick **Analysis ToolPak**, and select **OK**. The button appears on the **Data** tab.

**Excel for the web** doesn't have the ToolPak. Every ToolPak number this lesson uses can also be calculated with the worksheet
functions shown here, and those work everywhere.

The tools you'll use here:

| Data Analysis tool | What it produces | Worksheet-function equivalent |
|---|---|---|
| Descriptive Statistics | A table of mean, median, standard deviation, skewness, and more | AVERAGE, MEDIAN, STDEV.S, SKEW, … (section 3) |
| Histogram | Bin counts and an optional chart | FREQUENCY, COUNTIFS (section 4) |
| Correlation | A matrix of correlations between several columns | CORREL (section 5) |
| Regression | Coefficients, R², p-values, residuals | LINEST, SLOPE, INTERCEPT, RSQ (sections 6–7) |
| t-Test (three versions) | Means, t statistic, p-values | T.TEST (section 8) |
| Moving Average | A column of moving averages and a chart | AVERAGE over a sliding range (section 10) |

> ⚠️ **ToolPak output is frozen.** The tools write plain numbers, not formulas. If the data changes, the output doesn't, and
> nothing on the sheet warns you. Use the ToolPak to explore, and use worksheet functions for anything that has to stay current,
> such as a monthly report.

Every tool's dialog works the same way:

- Fill in **Input Range** with a cell range such as `Stays!$G$1:$G$401`. Click in the box and drag over the cells to fill it in.
  A range typed without a sheet name, such as `$G$1:$G$401`, refers to the sheet that was active when you opened the tool, so
  start each tool from the sheet that holds the data.
- Tick **Labels in first row** when your range includes the header, so the output uses the column name.
- **Output options** are an Output Range on the current sheet, a **New Worksheet Ply** (a new sheet), or a New Workbook. A new
  sheet is safest, because the tool overwrites whatever is already in the output range.

### 3. Describing a distribution: Descriptive Statistics

A **distribution** is the pattern of values a measure takes: where most values sit, how spread out they are, and whether a long
tail stretches to one side. Lesson 2.4 built these numbers one function at a time. The ToolPak builds them all at once.

1. Select the **Stays** sheet, then select **Data → Data Analysis → Descriptive Statistics → OK**.
2. Set **Input Range** to `$G$1:$G$401` (TotalCharges), **Grouped By** to *Columns*, and tick **Labels in first row**.
3. Choose **New Worksheet Ply**.
4. Tick **Summary statistics** and **Confidence Level for Mean** (leave it at 95%). Optionally tick **Kth Largest** and **Kth
   Smallest**.
5. Select **OK**.

Here is the output for TotalCharges, with the function that gives each row:

| Output row | TotalCharges | Same as |
|---|--:|---|
| Mean | 34,277.55 | `AVERAGE` |
| Standard Error | 1,423.44 | `STDEV.S(range)/SQRT(COUNT(range))` |
| Median | 23,897.57 | `MEDIAN` |
| Mode | #N/A | `MODE.SNGL`: no two stays were charged exactly the same amount |
| Standard Deviation | 28,468.72 | `STDEV.S` |
| Sample Variance | 810,467,880.68 | `VAR.S` |
| Kurtosis | 10.13 | `KURT` |
| Skewness | 2.59 | `SKEW` |
| Range | 232,177.28 | `MAX(range)-MIN(range)` |
| Minimum | 5,449.93 | `MIN` |
| Maximum | 237,627.21 | `MAX` |
| Sum | 13,711,019.22 | `SUM` |
| Count | 400 | `COUNT` |
| Confidence Level(95.0%) | 2,798.37 | `CONFIDENCE.T(0.05,STDEV.S(range),COUNT(range))` (section 9) |

Three rows are new since Lesson 2.4:

- **Skewness** measures how lopsided a distribution is. A value near 0 means roughly symmetric. A positive value means a long
  tail to the right (a few very large values), and a negative value means a long tail to the left. A common rule of thumb calls
  anything beyond ±1 strongly skewed. Charges, LOS, and wait times are almost always right-skewed, which is why their mean sits
  above their median.
- **Kurtosis** measures how heavy the tails are compared with a bell curve. Excel reports *excess* kurtosis, so a bell curve
  scores 0. A value of 10 means extreme values turn up far more often than a bell curve would predict.
- The **standard error** of the mean is the standard deviation divided by √n. The standard deviation describes how much
  individual stays vary. The standard error describes how much the *average* would vary if you drew another sample of 400
  stays. It shrinks as the sample grows, and it's the building block of confidence intervals and t-tests.

> 📋 `SKEW` and `KURT` use sample formulas, the ones the ToolPak uses. `SKEW.P` (Excel 2013 and later) gives the population
> version, which is slightly smaller for small samples.

### 4. Histograms

A **histogram** shows a distribution by sorting values into **bins**, which are consecutive ranges of equal or chosen width, and
counting the values in each bin. You have three ways to build one.

**FREQUENCY.** You met this function in Lesson 4.2. It counts values into bins in one formula:

```
=FREQUENCY(data_array, bins_array)
```

`bins_array` lists each bin's **upper limit** in ascending order. A bin counts the values *greater than the previous limit and
less than or equal to its own limit*. FREQUENCY returns one more count than there are limits, and that last count holds
everything above the top limit.

Here is TotalCharges in \$10,000 bins. With the limits 10,000, 20,000, …, 100,000 typed in I2:I11 of the Stays sheet,
`=FREQUENCY(G2:G401,I2:I11)` returns:

| Bin limit | Counts stays charged | Stays | Cumulative % |
|--:|---|--:|--:|
| 10,000 | \$10,000 or less | 18 | 4.5% |
| 20,000 | over \$10,000, up to \$20,000 | 138 | 39.0% |
| 30,000 | over \$20,000, up to \$30,000 | 93 | 62.3% |
| 40,000 | over \$30,000, up to \$40,000 | 44 | 73.3% |
| 50,000 | over \$40,000, up to \$50,000 | 28 | 80.3% |
| 60,000 | over \$50,000, up to \$60,000 | 19 | 85.0% |
| 70,000 | over \$60,000, up to \$70,000 | 20 | 90.0% |
| 80,000 | over \$70,000, up to \$80,000 | 10 | 92.5% |
| 90,000 | over \$80,000, up to \$90,000 | 9 | 94.8% |
| 100,000 | over \$90,000, up to \$100,000 | 9 | 97.0% |
| (extra 11th count) | over \$100,000 | 12 | 100.0% |

In Microsoft 365 and Excel 2021 or later, FREQUENCY spills all 11 counts. In Excel 2019 and earlier, select 11 cells first, type
the formula, and confirm it with **Ctrl + Shift + Enter** (Mac: **⌘ + Shift + Return**). To get a single bin's count in one cell,
wrap FREQUENCY in INDEX: `=INDEX(FREQUENCY(G2:G401,I2:I11),3)` returns 93. COUNTIFS does the same job one bin at a time:
`=COUNTIFS(G2:G401,">20000",G2:G401,"<=30000")` also returns 93.

**The Histogram tool.** **Data → Data Analysis → Histogram** asks for an **Input Range** (the data) and a **Bin Range** (the
limits, sorted ascending). It counts exactly the way FREQUENCY does and labels the last row *More*. Tick **Chart Output** for a
column chart, and **Cumulative Percentage** for the running total in the last column of the table above. If you leave the Bin
Range empty, Excel picks evenly spaced bins between the minimum and maximum, which rarely land on round numbers.

> ⚠️ **The Histogram tool reads the Bin Range from the active sheet.** It keeps the sheet name on the Input Range but drops it from
> the Bin Range, so bins on another sheet can be silently replaced by the same cells on the sheet you started from. Keep the bins on
> the same sheet as the data, or start the tool while the *bins* sheet is active. For Task 2, select the Bins sheet, open the tool,
> type `Stays!$F$2:$F$401` as the Input Range and `$A$4:$A$10` as the Bin Range, and leave **Labels** unticked.

**The Histogram chart (Excel 2016 and later).** Lesson 3.5, section 9, built this chart with **Insert → Insert Statistic Chart →
Histogram** and set its bins in **Format Axis** (bin width, number of bins, overflow, and underflow). The chart labels each bar in
interval notation, so the bar `(20000, 30000]` holds the same stays as FREQUENCY's 30,000 bin. Here is how the three methods
compare:

| | FREQUENCY | Histogram tool | Histogram chart |
|---|---|---|---|
| Updates when data changes | Yes | No | Yes |
| You choose the bin limits | Yes | Yes | Width, count, overflow, and underflow only |
| Counts in cells you can use in formulas | Yes | Yes (frozen) | No |
| Versions | All | Desktop with the ToolPak | Excel 2016 and later |

> ⚠️ **Read the bin label as an upper limit.** A bin labeled 30000 in FREQUENCY or the Histogram tool means "over 20,000 up to
> 30,000," not "30,000 and up." With whole-number data such as LOS, this decides which bin a value that sits exactly on a limit
> falls into. Misreading the limits is the most common histogram mistake.

> 💡 **Tip:** Pick bin widths that are round numbers a reader can say out loud (\$10,000, 2 days, 15 minutes), and aim for roughly
> 5 to 20 bins. Too few bins hide the shape. Too many turn it into noise.

The charges histogram shows a tall peak between \$10,000 and \$30,000 and a long tail to the right, which is exactly what a
skewness of 2.59 describes. **Insert → Insert Statistic Chart → Box and Whisker** draws the same distribution as a box from Q1 to
Q3 with the median inside it, and marks values beyond the IQR fences from Lesson 2.4 as separate dots.

### 5. Relationships: scatter plots and correlation

A **scatter plot** (the scatter chart from Lesson 3.5) puts one measure on each axis and draws a dot for every row. Select two
columns and choose **Insert → Insert Scatter (X, Y) or Bubble Chart → Scatter**. Excel puts the left column on the horizontal (x)
axis, so arrange the columns with the explanatory measure first.

The **correlation coefficient**, written **r**, measures how closely the dots follow a straight line. It runs from −1 to +1:

```
=CORREL(array1, array2)
=PEARSON(array1, array2)        identical to CORREL
```

| r | Meaning |
|---|---|
| +1 | Perfect straight line going up |
| about +0.5 to +0.9 | Strong positive relationship |
| about +0.3 to +0.5 | Moderate positive relationship |
| about +0.1 to +0.3 | Weak positive relationship |
| near 0 | No straight-line relationship |
| negative | Same scale, but one measure falls as the other rises |

These bands are rules of thumb, not laws. What counts as "strong" depends on the field. Examples from the workbook:

| Pair | Formula | r |
|---|---|--:|
| Monthly ED visits vs. inpatient discharges | `=CORREL(Monthly!C2:C25,Monthly!E2:E25)` | 0.752 |
| ESI level vs. door-to-provider minutes | `=CORREL(EDWaits!D2:D1629,EDWaits!E2:E1629)` | 0.474 |
| Daily arrivals, Memorial vs. Cedar Ridge | `=CORREL(EDDaily!C2:C732,EDDaily!E2:E732)` | 0.102 |
| Age vs. TotalCharges | `=CORREL(Stays!D2:D401,Stays!G2:G401)` | 0.081 |

The order of the two ranges doesn't matter for CORREL. When one cell of a pair is blank, as for LWBS patients on EDWaits, CORREL
skips that pair. **Data → Data Analysis → Correlation** builds a table of every pairwise r for several adjacent columns at once.

**r measures straight lines only.** ESI level and wait time have r = 0.474, yet the group medians show the relationship isn't a
straight line. The median wait climbs steeply from ESI 1 (3 minutes) to ESI 4 (61 minutes), then levels off, and ESI 5 patients
(59 minutes) wait slightly less than ESI 4 patients, because the least urgent patients are often fast-tracked. A curve, a U
shape, or two separate clusters can all produce a misleading r. **Always look at the scatter plot before you trust r.** A single
extreme point can also create or hide a correlation on its own.

**r² (r squared)** is the share of the variation in one measure that a straight line through the other explains. With r = 0.752,
r² = 0.565, so ED volume accounts for about 57% of the month-to-month variation in inpatient discharges.

> ⚠️ **Correlation isn't causation.** Months with more ED visits have more inpatient discharges. Part of that is direct, because
> some ED patients are admitted. But winter respiratory season raises *both* numbers at once, so some of the correlation comes
> from a **common cause**. When two measures move together, consider four explanations: A causes B, B causes A, something else
> causes both, or chance. Only a designed comparison, such as a trial or a careful before-and-after study, separates them.

### 6. Simple linear regression

**Linear regression** fits the straight line that best predicts one measure, **y**, from another, **x**:

```
y = intercept + slope × x
```

"Best" means **least squares**: the line that makes the sum of the squared vertical distances from the dots to the line as small
as possible. Each of those vertical distances, actual y minus predicted y, is a **residual**.

| Function | Returns |
|---|---|
| `SLOPE(known_y's, known_x's)` | How much y changes, on average, for each 1-unit increase in x |
| `INTERCEPT(known_y's, known_x's)` | The predicted y when x = 0 |
| `RSQ(known_y's, known_x's)` | R², the share of y's variation the line explains (the same as CORREL squared) |
| `STEYX(known_y's, known_x's)` | The **standard error of the estimate**: the typical size of a residual, in y's units |
| `FORECAST.LINEAR(x, known_y's, known_x's)` | The predicted y for a new x |
| `TREND(known_y's, known_x's, new_x's)` | Predictions for one or many new x values (and several x columns, section 7) |
| `LINEST(known_y's, known_x's, TRUE, TRUE)` | The whole regression in one array (section 7) |

Every one of these lists **known_y's before known_x's**. Swap the ranges and you get a different line, one that predicts x from
y. FORECAST.LINEAR also takes the new x as its first argument, ahead of both ranges.

**Worked example.** How many inpatient discharges should the system expect in a month with a given number of ED visits?

| Question | Formula | Result |
|---|---|--:|
| Slope | `=SLOPE(Monthly!E2:E25,Monthly!C2:C25)` | 0.3746 |
| Intercept | `=INTERCEPT(Monthly!E2:E25,Monthly!C2:C25)` | 40.91 |
| R² | `=RSQ(Monthly!E2:E25,Monthly!C2:C25)` | 0.5654 |
| Typical miss | `=STEYX(Monthly!E2:E25,Monthly!C2:C25)` | 22.31 |
| Prediction for a 600-visit month | `=FORECAST.LINEAR(600,Monthly!E2:E25,Monthly!C2:C25)` | 265.6 |

Read it like this. Each additional ED visit in a month goes with about 0.37 more inpatient discharges. A 600-visit month should
bring about 266 discharges, give or take roughly 22 (the STEYX). The intercept, 40.9, is the line's value at zero ED visits.
No month comes anywhere near zero ED visits, so the intercept is just where the line crosses the axis, not a real-world quantity.

To see the same line on a chart, add a linear trendline as in Lesson 3.5. Build the scatter plot (x = EDVisits,
y = InpatientDischarges), select **Chart Elements** (the **+** button) **→ Trendline → More Options** (Mac: **Chart Design → Add
Chart Element → Trendline → More Trendline Options**), and tick **Display Equation on chart** and **Display R-squared value on
chart**. The chart shows `y = 0.3746x + 40.911` and `R² = 0.5654`, the same numbers SLOPE, INTERCEPT, and RSQ return.

> ⚠️ **Don't extrapolate.** The line was fit on months with 403 to 631 ED visits. A prediction for 1,000 visits assumes the
> straight line keeps going, and nothing in the data shows that it does.

> 📋 **Version note:** `FORECAST.LINEAR` (Excel 2016 and later) replaced the older `FORECAST`, which takes the same arguments
> and still works.

### 7. Regression output: is the slope real?

A slope calculated from a sample is an estimate. A different sample would give a slightly different slope, so the question is
whether the true slope could plausibly be zero. Regression output answers it.

**The Regression tool.** Select the **Monthly** sheet, then select **Data → Data Analysis → Regression**. Set **Input Y Range** to
`$E$1:$E$25` (InpatientDischarges) and **Input X Range** to `$C$1:$C$25` (EDVisits), tick **Labels**, choose **New Worksheet
Ply**, and optionally tick **Residuals** and **Residual Plots**. The output has three blocks:

**Regression Statistics**

| Row | Value | Meaning |
|---|--:|---|
| Multiple R | 0.7519 | The absolute value of r |
| R Square | 0.5654 | Share of the variation in y explained (RSQ) |
| Adjusted R Square | 0.5456 | R² with a penalty for each x variable. Use it to compare models with different numbers of x's |
| Standard Error | 22.31 | The typical residual (STEYX) |
| Observations | 24 | n |

**ANOVA**

| | df | SS | MS | F | Significance F |
|---|--:|--:|--:|--:|--:|
| Regression | 1 | 14,241.30 | 14,241.30 | 28.62 | 0.0000227 |
| Residual | 22 | 10,947.20 | 497.60 | | |
| Total | 23 | 25,188.50 | | | |

**Coefficients**

| | Coefficients | Standard Error | t Stat | P-value | Lower 95% | Upper 95% |
|---|--:|--:|--:|--:|--:|--:|
| Intercept | 40.91 | 36.15 | 1.13 | 0.2699 | −34.05 | 115.88 |
| EDVisits | 0.3746 | 0.0700 | 5.35 | 0.0000227 | 0.2294 | 0.5198 |

The coefficients table is where the answers are:

- **Standard Error** of a coefficient measures how much that estimate would vary from sample to sample.
- **t Stat** = coefficient ÷ its standard error. It counts how many standard errors the estimate is away from zero.
- **P-value** is the probability of a t Stat at least this far from zero *if the true coefficient were zero* (section 8 explains
  p-values). EDVisits has p = 0.0000227, so a slope this steep would almost never appear by chance. The intercept's p of 0.27
  says it isn't distinguishable from zero, which doesn't matter here.
- **Lower 95% and Upper 95%** form a confidence interval for each coefficient. The slope is probably between 0.23 and 0.52
  discharges per ED visit.

With one x variable, Significance F equals the slope's p-value. It becomes useful with several x's, where it tests whether the
model as a whole explains anything.

> ⚠️ The Regression tool stops with an error if any Y or X cell is blank or holds text. Remove or filter out incomplete rows first.

**LINEST: the same output as a live formula.**

```
=LINEST(known_y's, known_x's, const, stats)
```

Set `const` to TRUE to fit an intercept (FALSE forces the line through zero) and `stats` to TRUE for the full statistics. With
one x column, LINEST returns a block of 5 rows × 2 columns:

| Row | Column 1 | Column 2 |
|:-:|---|---|
| 1 | slope | intercept |
| 2 | standard error of the slope | standard error of the intercept |
| 3 | R² | standard error of the estimate (STEYX) |
| 4 | F statistic | residual degrees of freedom (n − 2) |
| 5 | regression sum of squares | residual sum of squares |

In Microsoft 365 and Excel 2021 or later, the block spills. In Excel 2019 and earlier, select 5 × 2 cells first and confirm
with **Ctrl + Shift + Enter** (Mac: **⌘ + Shift + Return**). To pull out one number, wrap LINEST in INDEX(…, row, column), which
works in every version:

| You want | Formula | Result |
|---|---|--:|
| Slope | `=INDEX(LINEST(Monthly!E2:E25,Monthly!C2:C25,TRUE,TRUE),1,1)` | 0.3746 |
| Its standard error | `=INDEX(LINEST(Monthly!E2:E25,Monthly!C2:C25,TRUE,TRUE),2,1)` | 0.0700 |
| Residual df | `=INDEX(LINEST(Monthly!E2:E25,Monthly!C2:C25,TRUE,TRUE),4,2)` | 22 |

LINEST doesn't report p-values, but you can build one. Divide the slope by its standard error to get t, then convert t to a
two-tailed probability with `T.DIST.2T(ABS(t), df)`. LET (Lesson 4.2) keeps it readable:

```
=LET(fit, LINEST(Monthly!E2:E25, Monthly!C2:C25, TRUE, TRUE),
     tstat, INDEX(fit,1,1) / INDEX(fit,2,1),
     T.DIST.2T(ABS(tstat), INDEX(fit,4,2)))                          → 0.0000227
```

T.DIST.2T needs a non-negative t, which is why the formula uses ABS.

The coefficient's 95% confidence interval comes from the same pieces: slope ± t × its standard error, where t is now the
*critical* value `T.INV.2T(0.05, df)`. For EDVisits, `=T.INV.2T(0.05,22)` is 2.0739, so the interval is 0.3746 ± 2.0739 × 0.0700,
which runs from 0.2294 to 0.5198. Those are the Lower 95% and Upper 95% columns of the Regression output. In one formula:

```
=LET(fit, LINEST(Monthly!E2:E25, Monthly!C2:C25, TRUE, TRUE),
     INDEX(fit,1,1) - T.INV.2T(0.05, INDEX(fit,4,2)) * INDEX(fit,2,1))         → 0.2294
```

**Several x variables.** LINEST and the Regression tool both accept more than one x column, as long as the columns are next to
each other. Inpatient volume grew in 2025, so add MonthNum (column B) beside EDVisits (column C):

```
=LINEST(Monthly!E2:E25, Monthly!B2:C25, TRUE, TRUE)
```

Row 1 of the result is **in reverse order**: the coefficient of the *last* x column comes first. Here it reads 0.3765 (EDVisits),
1.040 (MonthNum), and 26.93 (the intercept). R² rises to 0.615, and adjusted R² to 0.578. Each coefficient now means "holding the
other x constant." The MonthNum coefficient says discharges drifted up by about 1 a month at the same ED volume, but its p-value
is 0.116, so the drift isn't statistically distinguishable from zero with only 24 months.

> 💡 **Tip:** Tick **Residual Plots** in the Regression tool, or chart the residuals against x yourself. Residuals should look like
> a shapeless band around zero. A curve or a funnel shape means a straight line is the wrong model, whatever R² says.

### 8. Hypothesis tests and t-tests

A **hypothesis test** asks whether a difference in your sample is large enough that chance alone is an unlikely explanation. It
starts from the **null hypothesis**, the boring explanation: "there's no real difference, and the gap is random sampling noise."
The test then calculates a **p-value**: the probability of seeing a difference at least as large as yours *if the null hypothesis
were true*.

- A small p-value means chance alone would rarely produce your result, so you **reject** the null hypothesis and call the
  difference **statistically significant**.
- The cutoff is the **significance level**, written **alpha**, usually 0.05.
- A large p-value means the data are *consistent with* no difference. It doesn't prove the groups are equal. The sample may simply
  be too small to see the difference.

> ⚠️ **A p-value isn't the probability that the null hypothesis is true**, and it says nothing about how *big* or how *important*
> a difference is. With thousands of rows, trivially small differences become significant. Always report the difference itself
> (in days, minutes, or dollars) next to the p-value.

**T.TEST** compares two means and returns the p-value directly:

```
=T.TEST(array1, array2, tails, type)
```

| Argument | Value | Use when |
|---|---|---|
| tails | 2 | You're asking "are they different?" in either direction. The usual choice |
| tails | 1 | You decided *before looking at the data* that only one direction matters. Gives half the two-tailed p-value |
| type | 1 | **Paired**: the same patients, units, or months measured twice (before and after). Both arrays must be the same length and in matching order |
| type | 2 | Two independent groups, assuming equal variances |
| type | 3 | Two independent groups, *not* assuming equal variances (**Welch's t-test**). The safe default for independent groups |

**Example 1: independent groups.** Do Bluestone Memorial and Cedar Ridge ED patients wait different amounts of time? Their Q4 2025
visits are separate blocks of the EDWaits sheet:

```
=T.TEST(EDWaits!E2:E1157, EDWaits!E1384:E1629, 2, 3)     → 0.772
```

Memorial averaged 48.97 minutes and Cedar Ridge 48.21. The p-value of 0.772 says random variation would produce a gap at least
this large most of the time, so there's no evidence of a difference. T.TEST skips the blank LWBS cells.

**Example 2: paired data.** Did monthly inpatient discharges rise from 2024 to 2025? Each 2025 month has a natural partner, the
same month in 2024, and both share that month's season:

| Test | Formula | p-value |
|---|---|--:|
| Paired | `=T.TEST(Monthly!E2:E13,Monthly!E14:E25,2,1)` | 0.0302 |
| Unpaired, equal variances | `=T.TEST(Monthly!E2:E13,Monthly!E14:E25,2,2)` | 0.1722 |
| Unpaired, Welch | `=T.TEST(Monthly!E2:E13,Monthly!E14:E25,2,3)` | 0.1743 |

Discharges rose from 2,681 in 2024 to 2,905 in 2025, an average of 18.7 more per month. The unpaired tests compare the two years
as loose piles of numbers, so the big winter-to-summer swing within each year swamps the difference. The paired test looks only at
each month's change from the year before, which removes the season, and the increase becomes significant. **When the data come in
natural pairs, use type 1.**

**The ToolPak versions.** **Data → Data Analysis** offers *t-Test: Paired Two Sample for Means*, *t-Test: Two-Sample Assuming
Equal Variances*, and *t-Test: Two-Sample Assuming Unequal Variances*. Each reports both groups' means, variances, and counts, the
t Stat, and both one-tail and two-tail p-values (`P(T<=t) one-tail` and `P(T<=t) two-tail`) with the matching critical t values.
Leave **Hypothesized Mean Difference** blank, which means 0. The two-tail p-value is what T.TEST returns.

> 📋 For the unequal-variance test, the ToolPak rounds the degrees of freedom to a whole number before calculating its p-value, so
> its result can differ from T.TEST's in the later decimal places. With samples of a few hundred, the difference is negligible.

### 9. Confidence intervals

A sample mean is an estimate. A **confidence interval** puts a range around it that, by a procedure that works 95% of the time,
contains the true mean:

```
mean ± t × s ÷ √n
```

Here s is the sample standard deviation, n is the number of values, and t is the critical value of Student's t distribution
for n − 1 degrees of freedom: `=T.INV.2T(0.05, n-1)`, about 1.97 for a few hundred values. The part after the ± is the
**margin of error**, or **half-width**, and one function calculates it:

```
=CONFIDENCE.T(alpha, standard_dev, size)
```

Alpha is 1 minus the confidence level, so 0.05 for 95% and 0.10 for 90%. For door-to-provider times at Bluestone Memorial:

```
=CONFIDENCE.T(0.05, STDEV.S(EDWaits!E2:E1157), COUNT(EDWaits!E2:E1157))     → 2.23
```

| Hospital (Q4 2025) | n with a provider time | Mean (min) | Std. dev. | ± margin | 95% interval |
|---|--:|--:|--:|--:|---|
| Bluestone Memorial | 1,135 | 48.97 | 38.29 | 2.23 | 46.7 to 51.2 |
| Ashby Falls | (Task 9) | | | | |
| Cedar Ridge | 242 | 48.21 | 37.00 | 4.69 | 43.5 to 52.9 |

The two hospitals have almost the same spread, but Cedar Ridge's interval is about twice as wide, because it has about a fifth
as many patients. The margin shrinks with √n, so **cutting the margin in half takes four times as many patients**.

> ⚠️ **Use COUNT for size.** The blank DoorToProviderMin cells are patients who left without being seen, not measurements. COUNT
> counts only numbers. ROWS counts every row, blanks included, and makes the interval too narrow. COUNTA skips truly empty cells,
> so it happens to agree with COUNT here, but it also counts text, such as an export that writes "LWBS" instead of leaving the
> cell blank.

| Function | Uses | When |
|---|---|---|
| `CONFIDENCE.T` | Student's t critical value | Almost always: you estimated the standard deviation from the same sample |
| `CONFIDENCE.NORM` | 1.96 from the normal distribution | Only when the standard deviation is known in advance. The older `CONFIDENCE` is the same as CONFIDENCE.NORM |

**What 95% means.** If you repeated the sampling many times and built an interval each time, about 95% of those intervals would
contain the true mean. It does **not** mean 95% of patients wait between 46.7 and 51.2 minutes. Individual waits range from 1
minute to over 4 hours. The interval is about the *average*.

> 💡 **Tip:** The interval relies on the sample mean behaving like a bell curve, which holds for skewed data like waits once n
> reaches a few dozen. For small samples of very skewed data, report the median and a percentile range (Lesson 2.4) instead.

### 10. Moving averages

Daily counts are noisy, and ED arrivals carry a strong weekly rhythm. From 2024–2025, Memorial averaged:

| Mon | Tue | Wed | Thu | Fri | Sat | Sun |
|--:|--:|--:|--:|--:|--:|--:|
| 13.58 | 12.22 | 12.18 | 12.35 | 12.04 | 10.89 | 10.51 |

A **moving average** replaces each value with the average of a window of nearby values. A 7-day window always holds exactly one of
each weekday, so the weekly pattern cancels out and the underlying level shows through.

| Type | Window for a given day | Use it for |
|---|---|---|
| **Trailing** | That day and the 6 days before it | Monitoring and reports: you can calculate it today |
| **Centered** | 3 days before, the day, and 3 days after | Describing the past: it lines up with the middle of the window |

The first trailing 7-day average you can calculate is for 01/07/2024, row 8 of EDDaily, because it needs 7 days of history:

```
=AVERAGE(EDDaily!C2:C8)     → 15.29     (Mon 12, Tue 16, Wed 21, Thu 19, Fri 11, Sat 15, Sun 13)
```

Fill it down and the range slides one row at a time: row 9 averages C3:C9, row 10 averages C4:C10, and so on. Leave the
references relative, with no \$ signs.

Other ways to get the same numbers:

- **Data → Data Analysis → Moving Average.** Set **Interval** to 7 and tick **Chart Output**. The tool produces trailing averages
  and writes #N/A for the first 6 rows, where the window isn't full yet.
- **A chart trendline.** On a line chart of daily arrivals, add **Trendline → More Options → Moving Average** with **Period** 7.
  It's also trailing.

> ⚠️ **Moving averages inside a Table.** If you type `=AVERAGE(C2:C8)` in a new column next to the EDDaily Table, Excel turns it
> into a calculated column and copies it into rows 2–7 too. There the 7-day window reaches above the first day, so those cells
> show errors or quietly average fewer than 7 days. Press **Ctrl + Z** (Mac: **⌘ + Z**) once right after the automatic fill to keep
> the formula only where you typed it, then fill it down from row 8 yourself. Or build the column one blank column away from the
> Table.

### 11. Forecasting with a trend

`FORECAST.LINEAR` and `TREND` turn the regression line of section 6 into a forecast by using *time* as x:

```
=FORECAST.LINEAR(x, known_y's, known_x's)
=TREND(known_y's, known_x's, new_x's)
```

Use a simple period number such as MonthNum (1–24) for x, and ask for period 25. Real dates work too, because dates are serial
numbers, but then the slope is per *day*, and months have 28 to 31 days. A period number keeps every step the same size.

**Worked example: inpatient discharges.**

```
=FORECAST.LINEAR(25, Monthly!E2:E25, Monthly!B2:B25)     → 244.99
```

The trend adds about 0.98 discharges a month (`=SLOPE(Monthly!E2:E25,Monthly!B2:B25)`), so the line forecasts about 245
discharges for January 2026. But look at the data. The two Januaries had 233 and 287 discharges, and winter months run high every
year. The trend line explains only 4% of the variation (`RSQ` = 0.044), because most of the variation is **seasonality**, a
pattern that repeats every year. A straight line can't see it.

**Adding seasonality with a seasonal index.** A classic fix multiplies the trend forecast by a **seasonal index** for the month:
how much that calendar month usually runs above or below the trend.

1. In column I of the Monthly sheet (one blank column away from the Table), calculate the trend value for every month: type
   `=TREND($E$2:$E$25,$B$2:$B$25,B2)` in I2 and fill it down to I25. Press **F4** (Mac: **⌘ + T**) to add the \$ signs.
2. In column J, divide actual by trend: `=E2/I2`, filled down. January 2024's 233 discharges ÷ a trend value of 221.5 = 1.052.
   January 2025 gives 287 ÷ 233.2 = 1.230.
3. Average the ratios for each calendar month. January's seasonal index is (1.052 + 1.230) ÷ 2 = **1.141**: Januaries run about
   14% above trend. `AVERAGEIFS` with `MONTH()` in a helper column does this for all 12 months at once.
4. Forecast = trend forecast × index = 244.99 × 1.141 ≈ **279.5** discharges for January 2026.

With only two years of history, each index rests on two values, so treat it as rough. Three or more years make it much more
stable.

> 💡 **Tip:** Test a forecasting method before you trust it. Fit it on 2024 only, forecast 2025, and compare the forecasts with what
> actually happened. The average absolute error, `=AVERAGE(ABS(actual-forecast))`, tells you how far off to expect the next
> forecast to be.

### 12. Exponential smoothing: FORECAST.ETS and Forecast Sheet

Excel 2016 added a more powerful method, **exponential triple smoothing (ETS)**, specifically the AAA version. It tracks three
things that update as each new value arrives: the current **level**, the **trend**, and a **seasonal** adjustment for each point
in the cycle. Recent values get more weight than old ones, and Excel picks the weights itself by fitting the history.

**Forecast Sheet (Windows).**

1. Select the timeline and the values. For monthly ED visits, select `A1:A25` on the Monthly sheet, then hold **Ctrl** and select
   `C1:C25`.
2. Select **Data → Forecast Sheet**.
3. Choose a line or column chart and set **Forecast End**, for example 12/1/2026.
4. Open **Options**:

| Option | What it does |
|---|---|
| Forecast Start | Where the forecast begins. Set it before the last actual value to see how the method would have done |
| Confidence Interval | The width of the shaded band. 95% by default |
| Seasonality | *Detect automatically*, or *Set manually* to the cycle length: 12 for monthly data with a yearly pattern, 7 for daily data with a weekly pattern |
| Include forecast statistics | Adds a table of the smoothing weights and error measures |
| Timeline Range / Values Range | The ranges you selected |
| Fill Missing Points Using | Interpolation (default) or zeros |
| Aggregate Duplicates Using | How to combine two values with the same date (Average by default) |

5. Select **Create**. Excel adds a new sheet with a Table of the history plus forecast rows, and a chart with the forecast and its
   confidence band.

The new sheet's forecast columns are ordinary formulas you can read and reuse:

| Function | Returns |
|---|---|
| `FORECAST.ETS(target_date, values, timeline, [seasonality], [data_completion], [aggregation])` | The forecast for target_date |
| `FORECAST.ETS.CONFINT(target_date, values, timeline, [confidence_level], …)` | The ± margin around that forecast. The sheet's Lower and Upper Confidence Bound columns are the forecast minus and plus this |
| `FORECAST.ETS.SEASONALITY(values, timeline, …)` | The season length Excel detected (0 means none) |
| `FORECAST.ETS.STAT(values, timeline, statistic_type, …)` | Smoothing weights and accuracy measures such as MAE and RMSE |

The `seasonality` argument is 1 (or omitted) to detect it automatically, 0 for none, or a whole number for the cycle length.

**Rules for the timeline.** The timeline needs a consistent step: every day, the 1st of every month, every year, or a period number
such as MonthNum. Excel can fill a small share of missing points and average duplicate dates, using the options above.

Try it twice on the monthly ED visits: once with seasonality detected automatically and once set manually to 12. With only two
years of monthly history, automatic detection has very little to go on. Compare the January 2026 forecast with the straight-line
forecast from Task 11, and with the two Januaries already in the data.

> 📋 **Version note:** The FORECAST.ETS functions work in Excel 2016 and later on Windows and Mac. The **Forecast Sheet**
> button is Windows-only, so on a Mac you type the FORECAST.ETS formulas yourself and build the chart by hand. Excel optimizes the
> smoothing weights internally, and other programs implement ETS differently, so expect small differences between tools and even
> between Excel versions.

> 📋 The ToolPak's **Exponential Smoothing** tool is *simple* smoothing: one weight, no trend, and no seasonality. Its "damping
> factor" is 1 minus the smoothing weight. For forecasting volumes with seasons, use FORECAST.ETS.

### 13. Control charts: separating signal from noise

Every process varies. Daily ED arrivals differ even when nothing about the community or the hospital has changed. **Statistical
process control (SPC)** names two kinds of variation:

- **Common-cause variation** is the everyday noise built into a stable process. Reacting to it, for example by praising a good day
  or investigating a bad one, wastes effort and often makes things worse.
- **Special-cause variation** comes from something outside the usual system: a flu outbreak, a bus crash, a new triage protocol, a
  data-entry error. It deserves investigation.

A **control chart** is a line chart of the measure over time with three horizontal lines: the **center line** (the average) and
the **upper and lower control limits** (UCL and LCL), usually 3 standard deviations either side. Points inside the limits are
treated as common-cause noise. Points outside are **signals**.

> ⚠️ **Control limits aren't targets.** They describe what the process *does*, calculated from its own data. A target or
> specification ("door-to-provider under 30 minutes") describes what you *want*. A stable process can consistently miss its target.

Which chart to use:

| Data | Chart | Example |
|---|---|---|
| One measurement or count per period | **I-chart** (individuals, also called XmR) | Daily ED arrivals, monthly ALOS |
| A proportion with a different denominator each period | **p-chart** | Monthly readmission rate, LWBS rate |
| Counts of rare events per period, same area of opportunity | c-chart (an I-chart also works) | Falls per month on one unit |
| Several measurements per period | X̄ and R charts | Five lab turnaround times sampled each day |

> 📋 A c-chart sets its limits from the count's own average, which assumes the events follow a Poisson pattern. The I-chart makes
> no such assumption, so it's the safer general-purpose choice for counts like daily ED arrivals.

**The I-chart.** You can't estimate the noise with STDEV.S of all the points, because STDEV.S also captures slow seasonal swings
and real shifts. Those inflate the limits and hide real signals. Instead, the I-chart estimates short-term noise from the
**moving range (MR)**: the absolute change from one point to the next.

```
MR (each point after the first) = ABS(this value − previous value)
MR̄  = average of the moving ranges
σ̂   = MR̄ ÷ 1.128
UCL = mean + 3 × σ̂        (the same as mean + 2.66 × MR̄)
LCL = mean − 3 × σ̂        (set to 0 if it's negative and the measure can't go below 0)
```

**MR̄** (say "MR-bar") is the average moving range, and **σ̂** (say "sigma-hat") is the estimated standard deviation of the
short-term noise. The constant 1.128, called d₂, converts the average range of two consecutive points into a standard
deviation. 3 ÷ 1.128 ≈ 2.66, which is why many references write the limits as mean ± 2.66 × MR̄.

**Worked example: Cedar Ridge daily arrivals, 2024 baseline** (EDDaily column E, rows 2–367).

| Step | Formula | Result |
|---|---|--:|
| Mean | `=AVERAGE(EDDaily!E2:E367)` | 2.544 |
| MR̄ | `=SUMPRODUCT(ABS(EDDaily!E3:E367-EDDaily!E2:E366))/COUNT(EDDaily!E3:E367)` | 1.726 |
| σ̂ | MR̄ ÷ 1.128 | 1.530 |
| UCL | mean + 3 × σ̂ | 7.13 |
| LCL | mean − 3 × σ̂ = −2.05 | 0 |

The MR̄ formula subtracts the whole column from itself shifted by one row: `E3:E367-E2:E366` pairs each day with the day before.
ABS makes every change positive, and SUMPRODUCT adds them up in any Excel version, without Ctrl + Shift + Enter. A helper column of
`=ABS(E3-E2)` filled down and then averaged gives the same 1.726. There are 365 moving ranges for 366 days, so divide by the count
of moving ranges, not the count of days.

**Baseline, then monitor.** Calculate the limits from a stable **baseline** period, freeze them, and then plot new data against
them. If you recalculated the limits every month, a real change would gradually pull the limits along with it, and you'd never see
it. Against the 2024 limits, `=COUNTIF(EDDaily!E368:E732,">"&7.13)` finds 1 day in 2025 above the UCL: 02/25/2025, with 8
arrivals. (In a real chart, point the criterion at the cell that holds the UCL rather than typing the number.)

**Signal rules.** A point outside the limits is the main signal, but SPC also flags patterns inside them. These are the most widely
used rules. Exact run lengths vary between references, so pick one set and use it consistently:

| Rule | Signal | Suggests |
|---|---|---|
| 1 | One point outside the control limits | A sudden special cause |
| 2 | A **shift**: 8 or more points in a row on the same side of the center line | The process level has changed |
| 3 | A **trend**: 6 or more points in a row all rising or all falling | Gradual drift |
| 4 | 2 out of 3 points in a row more than 2σ̂ from the center, on the same side | An early warning of a shift |

**Build the chart.** Excel has no control-chart type, so you build one from a line chart:

1. On the EDDaily sheet, put the center line, UCL, and LCL in three cells, say M2, M3, and M4. Then, one blank column away from
   the Table, add CL, UCL, and LCL columns that point at those cells with absolute references (`=$M$2`, `=$M$3`, `=$M$4`), so
   every row shows the same value.
2. Select the dates, the values, and the three limit columns, holding **Ctrl** (Mac: **⌘**) to add non-adjacent ranges. Select
   **Insert → Insert Line or Area Chart → Line**.
3. Format the limit series as dashed red lines and the center line as a solid gray line, without markers.
4. To mark the 2025 signals, type `=IF(E368>$M$3,E368,NA())` in row 368 of one more column, fill it down to row 732, and chart
   it as markers only. NA() returns #N/A, which a line chart doesn't plot, so only the signal points appear.

### 14. p-charts for rates

A **p-chart** monitors a proportion, such as the share of ED patients who leave without being seen. Each period's rate is a count
of events (x) divided by a denominator (n), and the denominator changes from month to month. A rate from a small month is less
reliable than a rate from a big one, so each period gets its **own** limits around one shared center line, **p̄** (say
"p-bar"):

```
p̄    = total events ÷ total denominator          (not the average of the monthly rates)
σᵢ   = SQRT(p̄ × (1 − p̄) ÷ nᵢ)                    for period i
UCLᵢ = p̄ + 3 × σᵢ
LCLᵢ = p̄ − 3 × σᵢ                                  (set to 0 if negative)
```

Small months get wide limits and big months get narrow ones, so the limit lines look like a staircase instead of straight lines.

**Worked example: the system-wide LWBS rate** (Monthly sheet, LWBS in column D and EDVisits in column C).

1. Center line: `=SUM(Monthly!D2:D25)/SUM(Monthly!C2:C25)` gives 237 ÷ 12,292 = **1.93%**. Put it in a cell well clear of the
   Table, say P2, and refer to it with `$P$2`. That leaves columns K to O free for the bonus.
2. In three helper columns, calculate each month's rate `=D2/C2`, its UCL `=$P$2+3*SQRT($P$2*(1-$P$2)/C2)`, and its LCL
   `=MAX(0,$P$2-3*SQRT($P$2*(1-$P$2)/C2))`. If you type them in columns H, I, and J, right next to the Table, the Table grows to
   include them and Excel fills each formula down to row 25 for you. That's what you want here, because every row needs the
   formula. If I and J still hold your seasonal-index work from section 11, clear them first.
3. Flag any month where the rate is above its UCL or below its LCL, and count the flags.

February 2024 (row 3) is the closest call. 21 of 586 patients left without being seen, a rate of 3.58%, just under its UCL of
3.63%. September 2024 had only 403 visits, so its LCL works out negative (−0.13%) and is set to 0. No month falls outside its
limits, so the LWBS rate shows only common-cause variation.

A useful way to judge how unusual a month is regardless of its n is the **z-score**: (rateᵢ − p̄) ÷ σᵢ. It counts how many
standard errors a month sits from the center line, so anything beyond ±3 is a signal. February 2024's z-score is 2.91.

> 📋 When every denominator is in the thousands, p-chart limits become so narrow that almost every point signals, because real
> processes vary a little more than the binomial formula assumes. A variant called the Laney p′-chart corrects for this. For
> monthly hospital rates with denominators in the hundreds, the standard p-chart works well.

> ⚠️ **Check that every period is complete.** A rate whose numerator needs follow-up time, such as 30-day readmissions or
> 90-day mortality, will look artificially low for the most recent periods, because their follow-up hasn't finished.

### 15. Choosing the right tool

| Question | Tool |
|---|---|
| What does this measure look like: typical value, spread, shape? | Descriptive Statistics, histogram, box and whisker (sections 3–4) |
| Do two measures move together? | Scatter plot, CORREL (section 5) |
| How much does y change per unit of x, and what's the prediction for a new x? | SLOPE, INTERCEPT, FORECAST.LINEAR, TREND (section 6) |
| Is a relationship real, and how much does it explain? | Regression tool or LINEST: p-value and R² (section 7) |
| Do two groups differ on average? | T.TEST, type 3 for independent groups and type 1 for pairs (section 8) |
| How precise is this average? | CONFIDENCE.T (section 9) |
| What's the underlying level of a noisy daily series? | 7-day moving average (section 10) |
| What will next month look like? | FORECAST.LINEAR with a seasonal index, or FORECAST.ETS (sections 11–12) |
| Has something changed, or is this noise? | I-chart for single values, p-chart for rates (sections 13–14) |

### 16. Shortcuts and version notes

| Action | Windows | Mac |
|---|---|---|
| Turn on the Analysis ToolPak | **File → Options → Add-ins**, set **Manage** to *Excel Add-ins*, then **Go…** | **Tools → Excel Add-ins…** |
| Open the ToolPak tools | **Data → Data Analysis** | **Data → Data Analysis** |
| Enter a legacy array formula (FREQUENCY or LINEST in Excel 2019 and earlier) | Ctrl + Shift + Enter | ⌘ + Shift + Return |
| Toggle absolute references (\$) while editing a formula | F4 | ⌘ + T |
| Undo a Table's automatic calculated-column fill | Ctrl + Z | ⌘ + Z |
| Add a non-adjacent range to a selection | Ctrl + drag | ⌘ + drag |
| Build a forecast with a chart | **Data → Forecast Sheet** | Not available (type FORECAST.ETS formulas) |

| Functions or features | Available in |
|---|---|
| CORREL, PEARSON, SLOPE, INTERCEPT, RSQ, STEYX, LINEST, TREND, FREQUENCY, SKEW, KURT, FORECAST | Every current version |
| T.TEST, T.DIST.2T, T.INV.2T, CONFIDENCE.T, CONFIDENCE.NORM, STDEV.S, VAR.S | Excel 2010 and later (Excel 2011 and later on a Mac) |
| FORECAST.LINEAR, the FORECAST.ETS functions, Histogram and Box and Whisker charts | Excel 2016 and later |
| Forecast Sheet | Excel 2016 and later for Windows |
| LET, XLOOKUP, spilling FREQUENCY and LINEST | Microsoft 365 and Excel 2021 or later |
| Analysis ToolPak | Desktop Excel for Windows and Mac (not Excel for the web) |

The older names TTEST, TDIST, TINV, CONFIDENCE, and FORECAST still work and match T.TEST, the t distribution functions,
CONFIDENCE.NORM, and FORECAST.LINEAR. Prefer the newer names, whose arguments say exactly which version of each calculation you
get.

## 🧪 Hands-on practice

Download [`4.5-statistics-forecasting.xlsx`](4.5-statistics-forecasting.xlsx) and open the **Practice** sheet. Type each answer
in the yellow cell, as a formula wherever possible. Where a task mentions a ToolPak tool, run the tool to see its output, then
enter the matching worksheet formula so the answer stays live. The **Check** column turns green when you're right.

<!-- BEGIN GENERATED: practice -->
Stays (rows 2–401) holds the 400 sampled inpatient stays: Bluestone Memorial in rows 2–201 and Cedar Ridge in rows 202–401. Bins (A4:A10) holds the LOS histogram bins. EDWaits holds Q4 2025 ED visits sorted by hospital. EDDaily holds one row per day from 01/01/2024 (row 2) to 12/31/2025 (row 732), with 2025 starting in row 368. Monthly holds Jan 2024 (row 2) to Dec 2025 (row 25). Answer with formulas where you can. Where a task says the ToolPak, a worksheet formula gives the same number. Each yellow cell holds one number, so wrap FREQUENCY or LINEST in INDEX there. Otherwise they spill into the cells below.

| # | Task | Hint |
|:-:|------|------|
| 1 | Run **Data → Data Analysis → Descriptive Statistics** on LOSDays for all 400 stays (tick Summary statistics). What Skewness does the output report? Enter it to 2 decimal places, or use the worksheet function that calculates it. | The ToolPak's Skewness row is the SKEW function |
| 2 | Build a histogram of LOSDays (all 400 stays) using the bins on the Bins sheet, with the Histogram tool or FREQUENCY. How many stays fall in the bin labeled 6? Enter that one count. | A bin's number is its upper limit. Which LOS values does that bin collect? |
| 3 | How strongly is a patient's age related to length of stay? Calculate the correlation coefficient between Age and LOSDays for all 400 stays, to 3 decimal places. | CORREL(array1, array2) |
| 4 | Finance wants to know how much a day of stay adds to the bill. Fit a straight line that predicts TotalCharges from LOSDays (all 400 stays). What is the slope, in dollars per day? Enter it to 2 decimal places. | SLOPE(known_y's, known_x's). The thing you predict goes first |
| 5 | Using the straight line that predicts LOSDays from Age (all 400 stays), what LOS does it predict for an 80-year-old patient? Enter days to 2 decimal places. | FORECAST.LINEAR(x, known_y's, known_x's), or INTERCEPT + SLOPE × 80 |
| 6 | Is the age effect from Task 5 real, or could it be chance? Run **Data → Data Analysis → Regression** with LOSDays as the Y range and Age as the X range, or build the p-value from LINEST as in guide section 7. What p-value does it report for the Age coefficient? Enter it to 4 decimal places. | t Stat = coefficient ÷ its standard error. T.DIST.2T turns t into a two-tailed p-value |
| 7 | Finance would rather quote a range than a single number. For the TotalCharges-on-LOSDays line from Task 4, what is the lower end of the 95% confidence interval for the slope? It's the "Lower 95%" value the Regression tool reports for LOSDays. Enter dollars per day to 2 decimal places. | slope − t × (standard error of the slope), where t = T.INV.2T(0.05, residual df) |
| 8 | Do Bluestone Memorial and Cedar Ridge differ in average LOS? Run a two-tailed t-test that does not assume equal variances, comparing Memorial's LOSDays (rows 2–201) with Cedar Ridge's (rows 202–401). Enter the p-value to 3 decimal places. | T.TEST(array1, array2, tails, type). Type 3 = two-sample, unequal variance |
| 9 | Estimate Ashby Falls' true average door-to-provider time with a 95% confidence interval, using its Q4 2025 visits (EDWaits rows 1158–1383). What is the interval's half-width (the ± margin), in minutes to 2 decimal places? Blank cells are patients who left without being seen. | CONFIDENCE.T(alpha, standard_dev, size). Alpha for 95% is 0.05 |
| 10 | A trailing 7-day moving average smooths out the weekday pattern. What was the 7-day moving average of Memorial's daily arrivals on 02/25/2025 (that day and the 6 days before it)? Enter it to 2 decimal places. | 02/25/2025 is EDDaily row 423. Average 7 rows ending there |
| 11 | Fit a linear trend to the 24 monthly EDVisits values, using MonthNum (1–24) as x. What does it forecast for month 25 (January 2026)? Enter it to 1 decimal place. | FORECAST.LINEAR(x, known_y's, known_x's) |
| 12 | Build an individuals (I) chart for Memorial's daily arrivals with 2024 as the baseline (EDDaily rows 2–367). The moving ranges are the absolute day-to-day changes, σ̂ = average moving range ÷ 1.128, and UCL = mean + 3 × σ̂. What is the UCL? Enter it to 2 decimal places. | Average \|today − yesterday\| over 2024, divide by 1.128, triple it, add the mean |
| 13 | Monitor 2025 against the 2024 limits. On how many days in 2025 (EDDaily rows 368–732) did Memorial's arrivals exceed the UCL from Task 12? | COUNTIF with ">"& your Task 12 cell (D17) |
<!-- END GENERATED: practice -->

## ✅ Answer key

The workbook has a hidden **Answer Key** sheet (right-click any sheet tab → **Unhide…** → *Answer Key*). Its *Live result* column
runs each sample formula against the data. The same answers are below, collapsed so you don't see them by accident.

<!-- BEGIN GENERATED: answers -->
<details>
<summary><b>🔑 Show the answer key</b> (try every task before you open this)</summary>

**1. Skewness of LOSDays (Descriptive Statistics)**

- **Answer:** 1.78
- **Solution:** `=SKEW(Stays!F2:F401)`

Skewness measures how lopsided a distribution is. Zero means symmetric, and a positive value means a long tail to the right. LOS has a skewness of 1.78, because most stays are short (median 4 days) while a few run to 19 days. The ToolPak's Descriptive Statistics output is a block of typed-in numbers, not formulas, so it won't update if the data changes. `SKEW` will. The output's Kurtosis row (4.65) is `KURT`. A positive value means heavier tails than a bell curve.

**2. Build a histogram of LOSDays (all 400 stays) using the bins on the Bins sheet, with…**

- **Answer:** 92
- **Solution:** `=INDEX(FREQUENCY(Stays!F2:F401,Bins!A4:A10),3)`

Each bin counts the values above the previous limit, up to and including its own limit. Bin 6 therefore counts stays of 5 or 6 days: more than 4, at most 6. FREQUENCY returns one count per bin plus a final "More" count (5 stays over 14 days), and INDEX picks count number 3. `=COUNTIFS(Stays!F2:F401,">4",Stays!F2:F401,"<=6")` gives the same 92. If you read the label as "6 to 7 days" you'd get 66, which is the most common histogram mistake.

**3. How strongly is a patient's age related to length of stay? Calculate the correlation…**

- **Answer:** 0.130
- **Solution:** `=CORREL(Stays!D2:D401,Stays!F2:F401)`

r = 0.130 is a weak positive correlation. Older patients stay slightly longer on average, but the scatter plot is a cloud, not a line. Squaring r gives R² = 0.017, so age accounts for only about 1.7% of the variation in LOS. `PEARSON` returns the same value, and the order of the two ranges doesn't matter for correlation.

**4. Finance wants to know how much a day of stay adds to the bill. Fit a straight line…**

- **Answer:** 6,665.43
- **Solution:** `=SLOPE(Stays!G2:G401,Stays!F2:F401)`

Each extra day of stay goes with about \$6,665.43 more in charges. The line is Charges = \$2,600.11 + \$6,665.43 × LOSDays, and `=INTERCEPT(Stays!G2:G401,Stays!F2:F401)` gives the intercept. `=RSQ(Stays!G2:G401,Stays!F2:F401)` gives R² = 43.8%, so LOS explains a little under half of the stay-to-stay differences in charges. The rest comes from diagnosis, procedures, ICU days, and so on. Swapping the arguments gives the slope of LOS on charges (0.000066 days per dollar), which answers a different question. SLOPE always takes the y-values (the outcome) first.

**5. Using the straight line that predicts LOSDays from Age (all 400 stays), what LOS does…**

- **Answer:** 5.11
- **Solution:** `=FORECAST.LINEAR(80,Stays!F2:F401,Stays!D2:D401)`

The line is LOS = 3.726 + 0.0173 × Age, so at age 80 it predicts 5.11 days. That's the average LOS the line expects for 80-year-olds, not a forecast for one patient. The typical miss around the line (`STEYX`) is 2.80 days, far more than the 0.52-day difference the line predicts between a 50-year-old and an 80-year-old. `=TREND(Stays!F2:F401,Stays!D2:D401,80)` and `=INTERCEPT(Stays!F2:F401,Stays!D2:D401)+SLOPE(Stays!F2:F401,Stays!D2:D401)*80` give the same result. The older `FORECAST` function does too.

**6. p-value of the Age slope (Regression tool or LINEST)**

- **Answer:** 0.0093
- **Solution:**

```
=LET(fit,LINEST(Stays!F2:F401,Stays!D2:D401,TRUE,TRUE),tstat,INDEX(fit,1,1)/INDEX(fit,2,1),T.DIST.2T(ABS(tstat),INDEX(fit,4,2)))
```


With stats set to TRUE, LINEST returns a 5 × 2 block: row 1 holds the slope (0.0173) and intercept, row 2 their standard errors (0.0066 for the slope), and row 4 column 2 the residual degrees of freedom (398). The t Stat is 0.0173 ÷ 0.0066 = 2.61, and `T.DIST.2T` gives p = 0.0093. That's below 0.05, so the slope is **statistically significant**: a true slope of zero would rarely produce a sample like this one. Yet R² is only 1.7%, so age explains almost none of the variation in LOS. Significant means "probably not zero," not "large" or "useful." Without LET (Excel 2019 and earlier), write `=T.DIST.2T(ABS(INDEX(LINEST(Stays!F2:F401,Stays!D2:D401,TRUE,TRUE),1,1)/INDEX(LINEST(Stays!F2:F401,Stays!D2:D401,TRUE,TRUE),2,1)),INDEX(LINEST(Stays!F2:F401,Stays!D2:D401,TRUE,TRUE),4,2))`.

**7. Finance would rather quote a range than a single number. For the…**

- **Answer:** 5,920.66
- **Solution:**

```
=LET(fit,LINEST(Stays!G2:G401,Stays!F2:F401,TRUE,TRUE),INDEX(fit,1,1)-T.INV.2T(0.05,INDEX(fit,4,2))*INDEX(fit,2,1))
```


The slope is \$6,665.43 with a standard error of \$378.83. With 398 residual degrees of freedom, `T.INV.2T(0.05,398)` = 1.9659, so the interval is \$6,665.43 ± 1.9659 × \$378.83, which runs from \$5,920.66 to \$7,410.19 per day. Finance can say each extra day goes with roughly \$5,921 to \$7,410 more in charges. It's the same t × standard error recipe CONFIDENCE.T uses for a mean (guide section 9), applied to a slope. Without LET, repeat the LINEST call inside each INDEX.

**8. Do Bluestone Memorial and Cedar Ridge differ in average LOS? Run a two-tailed t-test…**

- **Answer:** 0.513
- **Solution:** `=T.TEST(Stays!F2:F201,Stays!F202:F401,2,3)`

The sample means are 4.845 days (Memorial) and 4.660 days (Cedar Ridge), a difference of 0.185 days. T.TEST returns the p-value directly: 0.513. If the hospitals' true averages were equal, random sampling alone would produce a gap at least this large about 51% of the time, so you **can't conclude** they differ. That's not proof they're the same, only that this sample can't tell them apart. A one-tailed test (tails = 1) gives half the p-value (0.257). Type 2 (equal variances) gives 0.513 too, because with equal group sizes both tests use the same t statistic (0.654) and differ only slightly in degrees of freedom. The ToolPak's "t-Test: Two-Sample Assuming Unequal Variances" reports the same value as P(T<=t) two-tail.

**9. Estimate Ashby Falls' true average door-to-provider time with a 95% confidence…**

- **Answer:** 6.09
- **Solution:** `=CONFIDENCE.T(0.05,STDEV.S(EDWaits!E1158:E1383),COUNT(EDWaits!E1158:E1383))`

217 patients saw a provider. Their mean wait was 49.15 minutes with a standard deviation of 45.52. The margin is t × s ÷ √n = 1.9710 × 45.52 ÷ √217 = 6.09, so the 95% confidence interval is 43.1 to 55.2 minutes. Use COUNT, not ROWS: ROWS also counts the 9 blank rows, which aren't measurements, and gives 5.97. COUNTA agrees with COUNT here only because the blank cells are truly empty. It would also count a text entry such as "LWBS" if an export had one. `CONFIDENCE.NORM` (and the old `CONFIDENCE`) uses 1.96 instead of t and gives 6.06, slightly too narrow. The ToolPak's "Confidence Level(95.0%)" row is CONFIDENCE.T.

**10. A trailing 7-day moving average smooths out the weekday pattern. What was the 7-day…**

- **Answer:** 17.57
- **Solution:** `=AVERAGE(EDDaily!C417:C423)`

Rows 417–423 cover 02/19 through 02/25/2025, one of each weekday, so the Monday peak and the weekend dip cancel out. 17.57 arrivals a day was the highest 7-day average of 2025, against a 2025 daily mean of 12.17. A centered average (3 days either side) gives 14.86 for this date. Trailing averages are what you can calculate today, and centered ones are for describing the past. The ToolPak's Moving Average tool and a chart's Moving Average trendline are both trailing.

**11. Fit a linear trend to the 24 monthly EDVisits values, using MonthNum (1–24) as x. What…**

- **Answer:** 510.1
- **Solution:** `=FORECAST.LINEAR(25,Monthly!C2:C25,Monthly!B2:B25)`

The trend line is almost flat (slope -0.16 visits a month, R² = 0.0003), so the forecast is close to the 24-month average of 512.2. But winter is the busy season: the two Januaries had 580 and 631 visits, and the two Decembers 621 and 595. A straight line can't see seasons, so it under-forecasts winter and over-forecasts summer. A seasonal index or FORECAST.ETS (see the guide) handles that. `=TREND(…,…,25)` gives the same number as FORECAST.LINEAR.

**12. Build an individuals (I) chart for Memorial's daily arrivals with 2024 as the baseline…**

- **Answer:** 22.80
- **Solution:**

```
=AVERAGE(EDDaily!C2:C367)+3*SUMPRODUCT(ABS(EDDaily!C3:C367-EDDaily!C2:C366))/COUNT(EDDaily!C3:C367)/1.128
```


2024 averaged 11.773 arrivals a day. The 365 moving ranges average 4.148, so σ̂ = 4.148 ÷ 1.128 = 3.677 and UCL = 11.773 + 3 × 3.677 = 22.80. The LCL is 0.74. Many references write this as mean ± 2.66 × MR̄, which is the same thing, because 3 ÷ 1.128 ≈ 2.66. A helper column of `=ABS(C3-C2)` filled down and averaged works too. The moving range measures short-term, day-to-day noise, so a slow seasonal swing doesn't widen the limits the way STDEV.S does (it would put the UCL at 23.35).

**13. Monitor 2025 against the 2024 limits. On how many days in 2025 (EDDaily rows 368–732)…**

- **Answer:** 4
- **Solution:** `=COUNTIF(EDDaily!C368:C732,">"&D17)`

4 days broke the limit: 02/19 (24), 05/02 (23), 12/08 (24), 12/23 (23). Each is a **special-cause** signal worth a look, and in an ED the usual causes are respiratory season, a mass-casualty event, or a neighboring ED on diversion. No day fell below the LCL. Freezing the limits from a baseline year is standard practice: if you recalculate them with the new data, a lasting change gets absorbed into the limits. Limits from STDEV.S would flag only 2 days.

</details>
<!-- END GENERATED: answers -->

## 🏆 Bonus challenge

<!-- BEGIN GENERATED: bonus -->
The quality committee wants a p-chart of Bluestone's system-wide 30-day readmission rate by discharge month, January 2024 to December 2025 (Monthly sheet, rows 2–25). For each month, n = IndexStays (inpatient discharges of patients who didn't die) and the rate = Readmits30 ÷ IndexStays. The center line p̄ = total readmissions ÷ total index stays, and each month gets its own limits: p̄ ± 3 × √(p̄ × (1 − p̄) ÷ n). Your B1 answer lands in cell D6 of the Bonus sheet, so later parts can refer to it. Helper columns for each month's rate, UCL, and LCL make this much easier. Put them in empty columns to the right of the Monthly Table.

Enter your answers on the **Bonus** sheet of the workbook, and build any helper columns on the **Monthly** sheet.

- **B1.** What is the center line p̄ for all 24 months? Enter it as a percentage to 2 decimal places. *(Hint: Total readmissions ÷ total index stays. Don't average the 24 monthly rates)*
- **B2.** September 2024 had the fewest index stays (177). What is its upper control limit? Enter it as a percentage to 2 decimal places. *(Hint: p̄ + 3 × SQRT(p̄ × (1 − p̄) / n), with this month's n)*
- **B3.** How many of the 24 months fall outside their control limits (above the UCL or below the LCL)? *(Hint: Compare each month's rate with its own limits. Add helper columns for the rate, UCL, LCL, and a TRUE/FALSE flag, then COUNTIF the flags. One SUMPRODUCT also works)*
- **B4.** Which month is it? Enter its MonthStart date. *(Hint: Read it off your helper columns, or XLOOKUP(TRUE, your test, the MonthStart column))*
- **B5.** Drop December 2025 and recompute p̄ from the other 23 months (rows 2–24). The highest remaining month is February 2025 (row 15). How many standard errors above the new center line is it? Calculate z = (rate − p̄) ÷ √(p̄ × (1 − p̄) ÷ n) and enter it to 2 decimal places. (Hint: LET(p, new p̄, n, that month's IndexStays, (rate − p) / SQRT(p*(1−p)/n)))
<!-- END GENERATED: bonus -->

<!-- BEGIN GENERATED: bonus-answers -->
<details>
<summary><b>🔑 Show the bonus solution</b> (give it a real try first)</summary>

**B1. What is the center line p̄ for all 24 months? Enter it as a percentage to 2 decimal…**

- **Answer:** 14.62%
- **Solution:** `=SUM(Monthly!G2:G25)/SUM(Monthly!F2:F25)`

800 readmissions ÷ 5,473 index stays = 14.62%. Averaging the 24 monthly rates gives 14.66% instead, because it weights a small month the same as a big one. A rate for a period is total ÷ total (Lesson 1.4).

**B2. September 2024 had the fewest index stays (177). What is its upper control limit?…**

- **Answer:** 22.58%
- **Solution:** `=D6+3*SQRT(D6*(1-D6)/Monthly!F10)`

√(0.1462 × 0.8538 ÷ 177) = 0.0266, so the UCL is 0.1462 + 3 × that = 22.58%. The biggest month (Dec 2025, n = 306) gets a UCL of 20.68%. Smaller months get wider limits because a rate from fewer patients bounces around more. That's why a p-chart's limit lines look like a staircase.

**B3. How many of the 24 months fall outside their control limits (above the UCL or below…**

- **Answer:** 1
- **Solution:** `=SUMPRODUCT(--(ABS(Monthly!G2:G25/Monthly!F2:F25-D6)>3*SQRT(D6*(1-D6)/Monthly!F2:F25)))`

Only one month signals, and it falls below its LCL. No month is above its UCL. `ABS(rate − p̄) > 3σ` catches both directions in one test because every LCL here is above zero. When n is small, p̄ − 3σ can go negative and the LCL is set to 0, so test the two sides separately. With helper columns on the Monthly sheet: Rate `=G2/F2`, UCL `=Bonus!$D$6+3*SQRT(Bonus!$D$6*(1-Bonus!$D$6)/F2)`, LCL the same with − instead of +, and a flag column that tests whether the rate is above the UCL or below the LCL. Then COUNTIF the flags.

**B4. Which month is it? Enter its MonthStart date.**

- **Answer:** 12/01/2025
- **Solution:**

```
=XLOOKUP(TRUE,ABS(Monthly!G2:G25/Monthly!F2:F25-D6)>3*SQRT(D6*(1-D6)/Monthly!F2:F25),Monthly!A2:A25)
```


December 2025: 23 readmissions out of 306 index stays = 7.52%, below its LCL of 8.56%. Before anyone celebrates, ask why. The data ends on 12/31/2025, so a patient discharged on 12/20 has only 11 days of follow-up in the data, and any readmission in January 2026 hasn't happened yet. December's rate is low because the 30-day window isn't complete. That's a **data artifact**, not an improvement. Excluding months with incomplete follow-up is standard for readmission reporting.

**B5. Drop December 2025 and recompute p̄ from the other 23 months (rows 2–24). The highest…**

- **Answer:** 1.93
- **Solution:**

```
=LET(p,SUM(Monthly!G2:G24)/SUM(Monthly!F2:F24),n,Monthly!F15,(Monthly!G15/n-p)/SQRT(p*(1-p)/n))
```


The corrected p̄ is 15.04%, higher than 14.62% because the artificially low December no longer drags it down. February 2025's rate of 19.15% is z = 1.93 standard errors above it (it was 2.15 against the old center line). A point is a special-cause signal only beyond ±3, so every month is inside the limits: readmissions are **stable**, showing common-cause variation only. The chart says the committee shouldn't react to any single month. To lower the rate, change the process (discharge planning, follow-up calls) and then watch for a shift, such as 8 months in a row below the center line.

</details>
<!-- END GENERATED: bonus-answers -->

## Key takeaways

- The Analysis ToolPak produces whole tables of statistics in one go, but its output is frozen. Use the worksheet functions
  (SKEW, FREQUENCY, LINEST, T.TEST, CONFIDENCE.T) for anything that must stay current.
- A histogram bin's number is its upper limit, inclusive. Check how your tool defines bins before you read the counts.
- r measures straight-line association only, R² is the share of variation explained, and neither proves causation. Look at the
  scatter plot first.
- Regression functions list known_y's before known_x's. Read a slope with its standard error and p-value: significant means
  "probably not zero," not "large."
- A p-value is the chance of a difference at least this big if nothing real were going on. Use T.TEST type 3 for independent
  groups and type 1 for paired data, and report the size of the difference alongside it.
- A 95% confidence interval for a mean is mean ± CONFIDENCE.T(0.05, s, n). Its width shrinks with √n.
- Straight-line forecasts miss seasons. Use a seasonal index or FORECAST.ETS. Control charts use frozen baseline limits from the
  moving range (I-chart) or from each period's n (p-chart) to separate special causes from everyday noise.

<!-- BEGIN GENERATED: nav -->
---

⬅️ **Previous:** [4.4 Data Model, Power Pivot & DAX](../04-power-pivot-dax/README.md) · 🏠 [Course home](../../README.md) · **Next:** [4.6 Building Interactive Dashboards](../06-dashboards/README.md) ➡️
<!-- END GENERATED: nav -->
