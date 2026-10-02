# Lesson 2.5 · Conditional Counting & Summing

> **Level:** Beginner → Intermediate · **Time:** about 55 minutes · **Workbook:** [`2.5-conditional-aggregation.xlsx`](2.5-conditional-aggregation.xlsx)
> **Data:** 2,724 encounters that ended in 2025 across all four Bluestone Health System facilities (every 2025 encounter for a fixed sample of about 900 patients), with facility, department, service line, diagnosis, payer, charges, and length of stay. The 2,666 claims for those encounters that had been submitted by 12/31/2025, with status, denial reason, and payment dates.

Every month, someone at Bluestone Health System answers the same questions. How many patients did each hospital discharge? How
much did we bill Medicare? Which insurer denied the most claims, and how much money is stuck in those denials? What is the 30-day
readmission rate for cardiology? Each answer is a count, a sum, or an average of only the rows that meet some conditions. You
could filter the data and read the total off the status bar, but you'd have to do it again next month, and nobody could see which
filters you used. A conditional formula sits in a fixed cell on a report, recalculates as soon as new rows arrive, and shows exactly
which rule produced the number. In this lesson you'll learn the functions behind almost every hospital dashboard, then use them to
build a payer mix grid and an inpatient denial scorecard.

## What you'll learn

- Count and sum with conditions using COUNTIF(S), SUMIF(S), and AVERAGEIF(S)
- Find conditional extremes with MAXIFS and MINIFS
- Write criteria with operators, cell references, wildcards, and date ranges
- Build summary grids that fill with a single formula

## 📖 Guide

The examples use the lesson workbook, so open it and try each formula as you read.

| Sheet | Columns |
|---|---|
| **Encounters** (rows 2–2725) | A EncounterID, B PatientID, C EncounterType, D FacilityName, E DeptName, F ServiceLine, G AdmitDate, H DischargeDate, I LOSDays, J PrimaryDxCode, K DxDescription, L DxCategory, M DischargeDisposition, N PayerName, O TotalCharges, P Readmit30 |
| **Claims** (rows 2–2667) | A ClaimID, B EncounterID, C PayerName, D EncounterType, E ServiceDate, F SubmitDate, G BilledAmount, H PaidAmount, I ClaimStatus, J DenialReason, K PaidDate, L DaysToPay (yellow, for the bonus) |
| **Payer Mix** | The grid you fill in practice task 12 |
| **Scorecard** | The bonus scorecard |

A few columns need a definition. The data dictionary describes every column of
[`encounters.csv`](../../data/README.md#encounterscsv) and [`claims.csv`](../../data/README.md#claimscsv), the files these sheets
are built from.

- **EncounterType** is Inpatient (an admitted stay), Observation (a short hospital stay that isn't an admission), Emergency (an
  emergency department visit that didn't become a stay), or Outpatient (a clinic visit at the Bluestone Outpatient Pavilion).
- **LOSDays** is the length of stay, counted as the number of midnights between admission and discharge. It's filled only for
  Inpatient and Observation stays and is blank for Emergency and Outpatient visits.
- **Readmit30** is Y when the patient was admitted again within 30 days of this discharge and N when they weren't. It's filled only
  for Inpatient stays.
- **ClaimStatus** is Paid, Partially Paid, Denied, Appealed (denied, and the hospital is contesting it), or Pending (the insurer
  hasn't decided yet). **DenialReason** is blank on claims that went through cleanly, and **PaidDate** is blank until money arrives.

> 💡 **Tip:** The formulas below are written as you'd type them on the sheet that holds the data. Try them in an empty column at
> least one column away from the Table, such as column T on the Encounters sheet or column N on the Claims sheet, because a formula
> typed in the column right next to a Table becomes a new Table column. A few examples read an input from column R of the Encounters
> sheet, such as a threshold in R2. On any other sheet, such as Practice, put the sheet name in front of each range:
> `Encounters!C2:C2725`.

You don't have to drag across 2,724 rows. While you're typing a formula, click the first data cell (for example C2), then press
**Ctrl + Shift + ↓** (Mac: **⌘ + Shift + ↓**), and Excel extends the reference down to the last filled cell. In a column with gaps,
such as LOSDays, the jump stops at the first blank cell, so type the range there instead.

> 📋 **Tables:** Encounters and Claims are Excel Tables. When you select a whole Table column while you build a formula, for example
> with the shortcut above, Excel writes a **structured reference** such as `tblEncounters[EncounterType]` instead of
> `Encounters!C2:C2725`. Both point at the same cells and give the same result in a single formula. They behave differently when you
> fill a formula to the right, which matters for the summary grid in section 11. Lesson 3.1 covers structured references.

### 1. COUNTIF: count the rows that meet one condition

```
=COUNTIF(range, criteria)
```

**range** is the cells to test. **criteria** is the condition a cell must meet to be counted. COUNTIF returns how many cells meet it.

| Formula (Encounters sheet) | Result | What it counts |
|---|--:|---|
| `=COUNTIF(C2:C2725,"Inpatient")` | 688 | Inpatient encounters |
| `=COUNTIF(C2:C2725,"observation")` | 144 | Observation stays (case doesn't matter) |
| `=COUNTIF(D2:D2725,"Ashby Falls Community Hospital")` | 221 | Encounters at Ashby Falls |
| `=COUNTIF(P2:P2725,"Y")` | 96 | Inpatient stays followed by a readmission |

Text criteria follow four rules:

- **Put text in double quotes.** Without them, Excel treats Inpatient as a name it doesn't know and returns #NAME?.
- **The whole cell must match.** `"Medicare"` matches cells that say exactly Medicare, not Silverline Medicare Advantage. Section 8
  shows how wildcards match part of a cell.
- **Upper and lower case don't matter.** "INPATIENT", "inpatient", and "Inpatient" all match.
- **Spaces do matter.** A cell holding "Inpatient " with a trailing space doesn't match "Inpatient". TRIM (Lesson 2.2) removes the
  extra spaces, and Lesson 3.3 covers cleaning a whole dataset.

### 2. Writing criteria: operators, cell references, and blanks

The criteria argument is text that Excel reads as an **operator** followed by a **value**. When there's no operator, Excel assumes
"equal to." The operators are the comparison operators from Lesson 2.1: `=`, `<>`, `>`, `>=`, `<`, and `<=`.

| Criteria | Meaning | Example | Result |
|---|---|---|--:|
| `"Inpatient"` or `"=Inpatient"` | equal to | `=COUNTIF(C2:C2725,"Inpatient")` | 688 |
| `"<>Outpatient"` | not equal to | `=COUNTIF(C2:C2725,"<>Outpatient")` | 1,745 |
| `">=100000"` | greater than or equal to | `=COUNTIF(O2:O2725,">=100000")` | 19 |
| `"<1000"` | less than | `=COUNTIF(O2:O2725,"<1000")` | 1,032 |
| `0` or `"0"` | equal to a number | `=COUNTIF(I2:I2725,0)` | 13 |
| `""` | blank | `=COUNTIF(P2:P2725,"")` | 2,036 |
| `"<>"` | not blank | `=COUNTIF(P2:P2725,"<>")` | 688 |

The operator and the value sit together **inside one pair of quotes**, so Excel reads ">=100000" as a single instruction. A plain
number needs no quotes. `=COUNTIF(I2:I2725,0)` finds the 13 stays that ended on the day they started, and it skips the 1,892 blank
LOSDays cells, because a blank cell isn't a zero.

**Blanks.** Readmit30 is blank for every encounter that isn't an Inpatient stay, so `""` counts the 2,036 other encounters and `"<>"`
counts the 688 stays that have a Y or an N. On the Claims sheet, `=COUNTIF(J2:J2667,"")` counts the 2,257 claims with no
DenialReason and `=COUNTIF(J2:J2667,"<>")` counts the 409 that have one. The `""` criterion also counts cells whose formula returns an
empty text string (`""`), such as the unpaid rows of the helper column you build in the bonus. `"<>"` counts those cells too, because
a cell that holds a formula isn't empty. To count only the cells that hold a number, use COUNT.

**Cell references and functions.** To compare against a value that lives in a cell, write the operator in quotes and join the
reference to it with `&`:

| You want | Criteria | Formula |
|---|---|---|
| equal to whatever R1 holds | `R1` | `=COUNTIF(C2:C2725,R1)` |
| at least the number in R2 | `">="&R2` | `=COUNTIF(O2:O2725,">="&R2)` |
| not equal to R1 | `"<>"&R1` | `=COUNTIF(C2:C2725,"<>"&R1)` |
| above the average charge | `">"&AVERAGE(O2:O2725)` | `=COUNTIF(O2:O2725,">"&AVERAGE(O2:O2725))` |

With 100000 in R2, `">="&R2` builds the text ">=100000" and the formula returns 19. Change R2 to 50000 and the result becomes 139
without anyone touching the formula. That's why reports keep thresholds in labeled cells, like the rate-in-one-cell pattern from
Lesson 1.5. The last row works the same way: AVERAGE calculates the average charge first, and `">"&` turns it into a criterion, so the
formula counts the 730 encounters that cost more than average.

> ⚠️ **The reference goes outside the quotes.** `=COUNTIF(O2:O2725,">=R2")` returns 0, because Excel compares every charge with the
> letters "R2". Write the operator in quotes, then `&`, then the reference: `">="&R2`.

> ⚠️ **Numbers stored as text don't meet number criteria.** `">=100000"` skips a cell that holds the *text* "125000". If a count looks
> too low, check for numbers stored as text (Lesson 2.1, section 9) and convert them (Lesson 3.3).

> ⚠️ **Whole columns and "not" criteria.** You can write `C:C` instead of `C2:C2725`, and COUNTIFS handles whole columns quickly.
> But a "not equal" or blank criterion also matches the empty cells below the data. On the Claims sheet, `=COUNTIF(I:I,"<>Pending")`
> returns 1,048,297, because it counts the header and more than a million empty cells. Use exact ranges (or a Table column) with
> `"<>…"` and `""` criteria.

### 3. COUNTIFS: several conditions at once

```
=COUNTIFS(criteria_range1, criteria1, [criteria_range2, criteria2], …)
```

COUNTIFS takes its conditions in **pairs**: a range, then the criteria for that range. A row counts only when it passes **every**
pair, so the conditions are joined with AND.

| Formula (Encounters sheet) | Result | What it counts |
|---|--:|---|
| `=COUNTIFS(C2:C2725,"Observation",D2:D2725,"Bluestone Memorial Hospital")` | 111 | Observation stays at Bluestone Memorial |
| `=COUNTIFS(C2:C2725,"Inpatient",I2:I2725,">=10")` | 52 | Inpatient stays of 10 days or more |
| `=COUNTIFS(C2:C2725,"Inpatient",I2:I2725,">=3",I2:I2725,"<=5")` | 326 | Inpatient stays of 3 to 5 days |
| `=COUNTIFS(E2:E2725,"Intensive Care Unit",D2:D2725,"Ashby Falls Community Hospital")` | 8 | ICU encounters at Ashby Falls |

The third formula uses the LOSDays column twice. That's how you write **between**: one condition for the lower bound and one for the
upper bound. A single criterion such as `"3-5"` doesn't work.

The fourth formula shows why you often need more than one condition. Three hospitals each have a department called Intensive Care
Unit, so `=COUNTIF(E2:E2725,"Intensive Care Unit")` returns 42 encounters across the whole system. Only the FacilityName condition
narrows it to Ashby Falls.

COUNTIFS accepts up to 127 pairs, and it also works with a single pair. You can use COUNTIFS for every count and never rewrite a
formula when a second condition comes along.

> ⚠️ **Every range must be the same size.** `=COUNTIFS(C2:C2725,"Inpatient",D2:D2700,"Ashby Falls Community Hospital")` returns
> #VALUE!, because Excel lines the ranges up row by row and they don't match. Use the same first and last row in every range of the
> formula.

### 4. OR logic: add counts together

COUNTIFS can't say OR. `=COUNTIFS(J2:J2725,"J18.9",J2:J2725,"U07.1")` asks for rows whose code is J18.9 *and* U07.1 at the same time,
so it returns 0. To count rows that meet one condition *or* another, count each one separately and add the results.

**Two values in the same column.** These count the inpatient stays for pneumonia (J18.9) or COVID-19 (U07.1):

```
=COUNTIFS(C2:C2725,"Inpatient",J2:J2725,"J18.9")+COUNTIFS(C2:C2725,"Inpatient",J2:J2725,"U07.1")     → 103
=SUM(COUNTIFS(C2:C2725,"Inpatient",J2:J2725,{"J18.9","U07.1"}))                                    → 103
```

The second version uses an **array constant**: a list of values typed inside curly braces and separated by commas. You type the
braces yourself. Given a list as its criteria, COUNTIFS returns a list of counts, one for each value (84 pneumonia stays and 19
COVID-19 stays), and SUM adds them up. The pattern works in every version of Excel without Ctrl + Shift + Enter, and it grows easily:
`{"J18.9","U07.1","J96.01"}`. Adding is safe here because no row can have two different codes at once.

> 📋 In Microsoft 365, `=COUNTIFS(C2:C2725,"Inpatient",J2:J2725,{"J18.9","U07.1"})` without SUM spills the two counts into two cells
> side by side. Older versions show only the first count. Wrap it in SUM whenever you want one total.

**Conditions on different columns.** Now both conditions can be true for the same row, so adding the counts would count those rows
twice. Subtract the overlap once:

```
=COUNTIF(C2:C2725,"Observation")+COUNTIF(N2:N2725,"Medicare")-COUNTIFS(C2:C2725,"Observation",N2:N2725,"Medicare")     → 699
```

That's 144 Observation stays plus 582 Medicare encounters minus the 27 Observation stays that were billed to Medicare. Without the
subtraction you'd get 726. When an OR spans three or more columns, the overlaps multiply. Use a helper column with the OR function
and count its TRUEs, or the SUMPRODUCT approach from Lesson 2.1 (section 7). Lesson 4.2 goes further.

### 5. SUMIF and SUMIFS: add up the matching rows

```
=SUMIF(range, criteria, [sum_range])
=SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], …)
```

Both test one or more columns and add up another column. The arguments come in a different order, which trips up almost everyone at
least once:

| | SUMIF | SUMIFS |
|---|---|---|
| Number of conditions | exactly one | 1 to 127 |
| Where the column to add goes | **last**, and it's optional | **first**, and it's required |
| If the ranges are different sizes | silently shifts the sum range (see below) | returns #VALUE! |
| Available in | every version | Excel 2007 and later |

| Formula (Encounters sheet) | Result | What it adds |
|---|--:|---|
| `=SUMIF(C2:C2725,"Inpatient",O2:O2725)` | 23,508,195.40 | Charges for all Inpatient stays |
| `=SUMIFS(O2:O2725,D2:D2725,"Ashby Falls Community Hospital",C2:C2725,"Emergency")` | 302,329.95 | Emergency visit charges at Ashby Falls |
| `=SUMIF(O2:O2725,">=100000")` | 2,417,401.18 | The 19 charges of \$100,000 or more |

The last formula has no sum_range. When you leave it out, SUMIF adds the same cells it tested, which is handy when the condition is
on the amounts themselves.

On the Claims sheet, SUMIFS compares what was billed with what was collected. `=SUMIFS(G2:G2667,C2:C2667,"Keystone Health Partners")`
returns 4,692,944.96 billed, and `=SUMIFS(H2:H2667,C2:C2667,"Keystone Health Partners")` returns 1,747,255.43 paid. Insurers pay a
contracted rate rather than the full charge, so collections run far below charges.

**OR logic works the same way.** Give SUMIFS an array constant and wrap it in SUM, just as you did with COUNTIFS in section 4. This
totals the charges of the 103 inpatient stays for pneumonia or COVID-19:

```
=SUM(SUMIFS(O2:O2725,C2:C2725,"Inpatient",J2:J2725,{"J18.9","U07.1"}))     → 2,330,562.31
```

Don't wrap AVERAGEIFS this way. `SUM(AVERAGEIFS(…,{"J18.9","U07.1"}))` adds two averages together (44,373.17), and even
`AVERAGE(AVERAGEIFS(…))` gives an average of averages (22,186.58), which section 6 explains is usually wrong. For the true average
of the combined group, divide the total by the count: `SUM(SUMIFS(…))/SUM(COUNTIFS(…))` returns 22,626.82.

> ⚠️ **SUMIF quietly resizes sum_range.** SUMIF uses only the top-left cell of sum_range, then takes a block the same size as range.
> `=SUMIF(C2:C2725,"Inpatient",O5:O10)` doesn't return an error. It adds cells from O5:O2728, three rows out of step with the
> EncounterType it tested, and returns 7,813,006.39 with no warning. SUMIFS would have returned #VALUE!, which is one more reason to
> use SUMIFS for everything.

> ⚠️ **Turning SUMIF into SUMIFS?** Move the column to add to the front. `=SUMIFS(C2:C2725,"Inpatient",O2:O2725)` isn't SUMIF with
> an S added. It doesn't work, because SUMIFS expects the column to add first and a range in the second position.

### 6. AVERAGEIF and AVERAGEIFS

```
=AVERAGEIF(range, criteria, [average_range])
=AVERAGEIFS(average_range, criteria_range1, criteria1, [criteria_range2, criteria2], …)
```

The argument order mirrors SUMIF and SUMIFS. The column to average comes last in AVERAGEIF and first in AVERAGEIFS.

| Formula (Encounters sheet) | Result | What it averages |
|---|--:|---|
| `=AVERAGEIF(C2:C2725,"Inpatient",I2:I2725)` | 4.88 | Inpatient length of stay, the **average length of stay (ALOS)** |
| `=AVERAGEIF(C2:C2725,"Observation",I2:I2725)` | 1.10 | Observation length of stay |
| `=AVERAGEIFS(I2:I2725,C2:C2725,"Inpatient",D2:D2725,"Ashby Falls Community Hospital")` | 5.13 | Inpatient ALOS at Ashby Falls |
| `=AVERAGEIF(C2:C2725,"Observation",O2:O2725)` | 8,460.26 | Average charge for an Observation stay |

Three behaviors are worth knowing:

- **Blank and text cells in the average range are skipped.** LOSDays is blank for Emergency and Outpatient visits, so even the plain
  `=AVERAGE(I2:I2725)` averages only the 832 stays that have a value (4.23 days). A **zero** is a value, though, so a stay of 0 days
  pulls the average down.
- **No matching rows means #DIV/0!.** `=AVERAGEIFS(I2:I2725,D2:D2725,"Bluestone Outpatient Pavilion")` returns #DIV/0!, because none of
  the Pavilion's clinic visits has a LOSDays value. Section 10 shows how to guard against it.
- **An average of averages is usually wrong.** The three hospitals' inpatient ALOS values are 4.84, 5.13, and 4.88. Their simple average
  is 4.95, but the system-wide ALOS is 4.88, because Bluestone Memorial's 513 stays outweigh Ashby Falls's 82. Average the rows, not
  the averages, just as Lesson 1.4 calculated a period's rate as total ÷ total.

### 7. MAXIFS and MINIFS: conditional extremes

```
=MAXIFS(max_range, criteria_range1, criteria1, [criteria_range2, criteria2], …)
=MINIFS(min_range, criteria_range1, criteria1, [criteria_range2, criteria2], …)
```

These work like SUMIFS: the column to search comes first, then the range/criteria pairs.

| Formula | Result | What it finds |
|---|--:|---|
| `=MAXIFS(I2:I2725,C2:C2725,"Inpatient",D2:D2725,"Ashby Falls Community Hospital")` | 16 | Ashby Falls's longest inpatient stay, in days |
| `=MINIFS(O2:O2725,C2:C2725,"Inpatient")` | 5,195.41 | The least expensive inpatient stay |
| `=MAXIFS(K2:K2667,C2:C2667,"Medicare")` (Claims sheet) | 12/31/2025 | The most recent Medicare payment |

Dates are numbers, so MAXIFS finds the latest date and MINIFS finds the earliest. The result may show as a serial number (46022).
Format the cell as a date with **Ctrl + Shift + #** (Mac: **⌃ + Shift + #**) or **Home → Number Format → Short Date**.

> ⚠️ **No match returns 0, not an error.** `=MAXIFS(I2:I2725,D2:D2725,"Bluestone Outpatient Pavilion")` returns 0, which looks like a
> real answer ("the longest stay was 0 days"). When nothing might match, count first:
> `=IF(COUNTIFS(D2:D2725,"Bluestone Outpatient Pavilion",I2:I2725,"<>")=0,"none",MAXIFS(I2:I2725,D2:D2725,"Bluestone Outpatient Pavilion"))`
> returns "none".

> 📋 **Version note:** MAXIFS and MINIFS need Excel 2019 or later (Windows or Mac), Microsoft 365, or Excel for the web. Older
> one-time-purchase versions, such as Excel 2016, show #NAME?. In those versions, use
> `=MAX(IF((C2:C2725="Inpatient")*(D2:D2725="Ashby Falls Community Hospital"),I2:I2725))` and finish it with
> **Ctrl + Shift + Enter** (Mac: **⌘ + Shift + Return**). IF returns FALSE for every row that fails the test, and MAX ignores the
> FALSEs. This workaround can't use wildcards.

### 8. Wildcards: match part of a cell

Text criteria can include three special characters, called **wildcards**:

| Wildcard | Matches | Example (Encounters sheet) | Result |
|:-:|---|---|--:|
| `*` | any number of characters, including none | `=COUNTIF(J2:J2725,"E11*")` | 146 |
| `?` | exactly one character | `=COUNTIF(J2:J2725,"E11.?")` | 73 |
| `???` | exactly three characters | `=COUNTIF(J2:J2725,"???")` | 282 |
| `*` on both sides | contains | `=COUNTIF(K2:K2725,"*fracture*")` | 21 |
| `~` | the next `*`, `?`, or `~` is a plain character | `=COUNTIF(K2:K2725,"*~?*")` | cells that contain a question mark (0 here) |

Here's how to read the examples:

- **"E11\*"** matches every type 2 diabetes code: E11.65 (with hyperglycemia) and E11.9 (without complications).
- **"E11.?"** allows exactly one character after the dot, so it matches E11.9 but not E11.65. Half of the diabetes encounters drop out.
- **"???"** matches codes that are exactly three characters long: I10, R55, Z23, O80, and O82.
- **"\*fracture\*"** finds the word fracture anywhere in the description, whether at the start, in the middle, or at the end.
- **"\*~?\*"** puts a tilde in front of the question mark, so the `?` is a plain character instead of a wildcard. The criterion
  finds descriptions that contain a real question mark. A literal asterisk works the same way: `"*~**"` finds cells that contain
  an asterisk, and `"~*"` alone matches a cell holding just one asterisk.

Wildcards are powerful, and that's the danger. `=COUNTIF(N2:N2725,"*Medicare*")` returns 987, because it catches Silverline Medicare
Advantage (405 encounters) as well as Medicare (582). `=COUNTIF(D2:D2725,"Bluestone*")` returns 2,256, because two facilities start with
Bluestone: Bluestone Memorial Hospital and Bluestone Outpatient Pavilion. Before you trust a wildcard count, filter the column and
check which values it caught.

A wildcard can surround a cell reference, which turns that cell into a search box. With the word unspecified typed in R3,
`=COUNTIF(K2:K2725,"*"&R3&"*")` returns 1,485. Clinical documentation integrity (CDI) teams track this number, because a vague code
such as "Heart failure, unspecified" can pay less than a specific one. Add a condition to see the inpatient picture:
`=COUNTIFS(C2:C2725,"Inpatient",K2:K2725,"*"&R3&"*")` returns 384 of the 688 stays.

> ⚠️ **Wildcards only match text.** ICD-10 codes, MRNs, and ZIP codes stored as text work. A number, such as a charge of 4,250, doesn't
> match `"4*"`. To test a number, use a numeric range such as `">=4000"` and `"<5000"`.

> ⚠️ **Wildcards work in criteria, not in `=` comparisons.** `=K2="*sepsis*"` compares K2 with the exact text \*sepsis\*,
> asterisks included, so it returns FALSE. Wildcards work in COUNTIF(S), SUMIF(S), AVERAGEIF(S), MAXIFS, MINIFS, SEARCH (Lesson 2.2),
> and the lookup functions in Lesson 2.6.

### 9. Date criteria

Dates are serial numbers (Lesson 2.3), so the comparison operators work on them. Build the date with DATE and join it to the
operator with `&`:

```
=COUNTIFS(H2:H2725,">="&DATE(2025,10,1),H2:H2725,"<"&DATE(2026,1,1))     → 719 encounters discharged in Q4 2025
```

It's the "between" pattern from section 3, with dates: one condition for the start and one for the end, both on DischargeDate.

Three habits make date criteria reliable:

- **Build dates with DATE or take them from a cell.** A criterion like `">=10/1/2025"` can work, but Excel reads that text using your
  computer's regional settings, and in a day-month-year region 10/1/2025 is January 10. `DATE(2025,10,1)` means October 1 everywhere.
  A cell holding a real date is even better, because the report period can change without editing any formula. With 10/01/2025 in R5
  and 12/31/2025 in R6, the criteria are `">="&R5` and `"<="&R6`.
- **End with "less than the first day of the next period."** `"<"&DATE(2026,1,1)` catches every moment of December 31, even in a column
  that holds times. `"<="&DATE(2025,12,31)` stops at midnight at the *start* of December 31, so it drops that day's rows when the column
  holds date-times (Lesson 2.3). EDATE calculates the next period's start for you: `"<"&EDATE(R5,3)` is three months after the date in R5.
- **Keep each condition on a real column.** Every criteria_range must be a range of cells, so Excel won't accept
  `=COUNTIFS(MONTH(H2:H2725),10)`. Write the month as a date range instead, add a helper column with `=MONTH(H2)`, or use SUMPRODUCT
  (Lesson 2.1).

On the Claims sheet, this formula totals the money collected in December 2025 (185 payments):

```
=SUMIFS(H2:H2667,K2:K2667,">="&DATE(2025,12,1),K2:K2667,"<="&DATE(2025,12,31))     → 681,430.02
```

PaidDate holds whole dates with no times, so `"<="` the last day is safe here.

### 10. Rates, and the divide-by-zero guard

Most quality and finance measures are **rates**: readmissions per discharge, denials per decided claim, falls per 1,000 patient days.
Each one is a count divided by a count:

```
rate = COUNTIFS(denominator conditions, plus the outcome) / COUNTIFS(denominator conditions)
```

The **denominator** describes everyone who could have had the outcome. The **numerator** is the same group with one more condition
for the outcome itself.

| Measure | Formula | Result |
|---|---|--:|
| 30-day readmission rate, all hospitals | `=COUNTIFS(C2:C2725,"Inpatient",P2:P2725,"Y")/COUNTIF(C2:C2725,"Inpatient")` | 14.0% |
| 30-day readmission rate, Ashby Falls | `=COUNTIFS(D2:D2725,"Ashby Falls Community Hospital",C2:C2725,"Inpatient",P2:P2725,"Y")/COUNTIFS(D2:D2725,"Ashby Falls Community Hospital",C2:C2725,"Inpatient")` | 11.0% |
| Claim denial rate (Claims sheet) | `=SUM(COUNTIFS(I2:I2667,{"Denied","Appealed"}))/COUNTIF(I2:I2667,"<>Pending")` | 11.8% |

The denial rate counts Appealed claims as denials, because every appeal starts with a denial. It leaves Pending claims out of the
denominator, because the insurer hasn't decided them yet. That gives 282 denials out of 2,387 decided claims. To show a rate as a
percentage, press **Ctrl + Shift + %** (Mac: **⌃ + Shift + %**), then add a decimal place with **Home → Increase Decimal**.

> ⚠️ **The denominator must be the right population.** If it includes rows that could never have the outcome, the rate comes out too
> low. A clinic visit can't be followed by a readmission (Readmit30 applies only to inpatient stays), so clinic visits don't belong in
> a readmission denominator, even when they share a service line with the inpatient units.

**The divide-by-zero guard.** A rate for a group with no rows divides 0 by 0. The Ambulatory service line has only clinic visits, so
its readmission rate is an error:

```
=COUNTIFS(F2:F2725,"Ambulatory",P2:P2725,"Y")/COUNTIFS(F2:F2725,"Ambulatory",C2:C2725,"Inpatient")     → #DIV/0!
```

You can guard it in two ways:

```
=IF(COUNTIFS(F2:F2725,"Ambulatory",C2:C2725,"Inpatient")=0,"n/a",COUNTIFS(F2:F2725,"Ambulatory",P2:P2725,"Y")/COUNTIFS(F2:F2725,"Ambulatory",C2:C2725,"Inpatient"))
=IFERROR(COUNTIFS(F2:F2725,"Ambulatory",P2:P2725,"Y")/COUNTIFS(F2:F2725,"Ambulatory",C2:C2725,"Inpatient"),"n/a")
```

| Guard | What it catches | Use it when |
|---|---|---|
| `IF(denominator=0,"n/a",…)` | only the empty-group case | you divide by a count you can test, which covers most rates |
| `IFERROR(…,"n/a")` | every error, including typos and ranges of different sizes | the function has no separate denominator to test, as with AVERAGEIFS |

The IF version is longer but safer, because a broken range still shows its error instead of a calm "n/a" (Lesson 2.1 makes the same
point). Text such as "n/a" is skipped by MAX, MIN, and AVERAGE, so a column of guarded rates can still be summarized.

### 11. Summary grids: one formula for every cell

A **summary grid** has category labels down the side, labels across the top, and a conditional count, sum, or average in every cell.
In Lesson 1.5 you filled a grid with one formula by using mixed references. The same technique works with COUNTIFS, and it's how
many hospital monthly reports are built.

**Worked example: inpatient discharges by month and hospital.** On a new sheet, type the three hospital names in B4:D4. For the
months, type 01/01/2025 in A5 and `=EDATE(A5,1)` in A6, then copy A6 down to A16, so column A holds the first day of every month of
2025. EDATE returns a serial number, so if A6:A16 show numbers such as 45689, format them as dates. The finished grid looks like
this:

|   | A | B | C | D |
|---|---|--:|--:|--:|
| **4** | *Month ↓ Hospital →* | Bluestone Memorial Hospital | Ashby Falls Community Hospital | Cedar Ridge Medical Center |
| **5** | 01/01/2025 | 47 | 7 | 11 |
| **6** | 02/01/2025 | 48 | 13 | 13 |
| **7** | 03/01/2025 | 37 | 14 | 7 |
| … | … | … | … | … |
| **16** | 12/01/2025 | 62 | 7 | 13 |

Type one formula in B5. It's shown on two lines to fit the page, but you type it on one line:

```
=COUNTIFS(Encounters!$D$2:$D$2725,B$4,Encounters!$C$2:$C$2725,"Inpatient",
          Encounters!$H$2:$H$2725,">="&$A5,Encounters!$H$2:$H$2725,"<"&EDATE($A5,1))
```

Then copy it across to D5 and down to row 16. Each reference has its own job:

| Reference | What's locked | Why |
|---|---|---|
| `Encounters!$D$2:$D$2725` and the other data ranges | row and column | The data never moves. Without the \$ signs, the ranges would slide right and down as you copy, and the counts would quietly go wrong. |
| `B$4` | the row | Every cell reads the hospital name in row 4 of its own column. |
| `$A5` | the column | Every cell reads the month in column A of its own row. |
| `EDATE($A5,1)` | the column | The first day of the following month, which gives the "less than the next period" end from section 9. |

To decide where each \$ goes, ask the two questions from Lesson 1.5: *should this reference move when I copy across?* and *should it
move when I copy down?* Here's what one wrong \$ does:

| Mistake in B5 | What you see |
|---|---|
| `$B$4` instead of `B$4` | Every column repeats Bluestone Memorial's counts |
| `A5` instead of `$A5` | Column C reads B5, a count, as its month, so it shows only 0s |
| Data ranges without \$ signs | The ranges drift to other columns and rows, so most cells show 0 or a count that's slightly off |

> ⚠️ **Type the data ranges as addresses, not Table references.** If you build the formula by selecting whole Table columns, the
> ranges appear as structured references such as `tblEncounters[FacilityName]`, which have no \$ signs to lock. Copy and paste leaves
> them alone, but filling to the right with the fill handle shifts each one to the next Table column. `tblEncounters[FacilityName]`
> becomes `tblEncounters[DeptName]`, and the copied columns fill with 0s. Type the ranges as addresses such as
> `Encounters!$D$2:$D$2725`, or fill the grid with copy and paste. (Lesson 3.1 shows how to lock a Table column with
> `tblEncounters[[FacilityName]:[FacilityName]]`.)

Check a finished grid three ways:

1. **Grand total.** Add up the whole grid and compare it with one count of the whole population. This grid should add up to
   `=COUNTIF(Encounters!C2:C2725,"Inpatient")`, which is 688.
2. **Row or column totals.** Each column should equal a simpler count for that hospital, such as
   `=COUNTIFS(Encounters!D2:D2725,B4,Encounters!C2:C2725,"Inpatient")`.
3. **Spot check.** Filter the Encounters sheet to one hospital, Inpatient, and one month, then compare the row count in the status bar
   with that cell.

> 💡 **Tip (Microsoft 365):** If you give COUNTIFS a whole column of labels for one criterion and a whole row of labels for another,
> one formula spills the entire grid. With hospital names in A5:A7 and encounter types in B4:E4,
> `=COUNTIFS(Encounters!D2:D2725,A5:A7,Encounters!C2:C2725,B4:E4)` fills a 3 × 4 block. The copy-and-lock method above works in every
> version, so the practice task uses it.

Where do the labels come from? You can type them, copy the column and use **Data → Remove Duplicates** on the copy, or, in Microsoft
365, use the UNIQUE function (Lesson 4.1).

> 📋 A PivotTable (Lesson 3.4) builds a grid like this by drag and drop. Formula grids are still worth knowing, because they keep a
> fixed layout from month to month, can mix counts, sums, and rates in one table, and recalculate without a Refresh.

### 12. Quick reference and version notes

| Function | Syntax | Available in |
|---|---|---|
| COUNTIF | `COUNTIF(range, criteria)` | every version |
| COUNTIFS | `COUNTIFS(criteria_range1, criteria1, …)` | Excel 2007 and later |
| SUMIF | `SUMIF(range, criteria, [sum_range])` | every version |
| SUMIFS | `SUMIFS(sum_range, criteria_range1, criteria1, …)` | Excel 2007 and later |
| AVERAGEIF | `AVERAGEIF(range, criteria, [average_range])` | Excel 2007 and later |
| AVERAGEIFS | `AVERAGEIFS(average_range, criteria_range1, criteria1, …)` | Excel 2007 and later |
| MAXIFS, MINIFS | `MAXIFS(max_range, criteria_range1, criteria1, …)` | Excel 2019 and later, Microsoft 365, Excel for the web |

A memory aid for the argument order: the **-IFS** functions (SUMIFS, AVERAGEIFS, MAXIFS, MINIFS) all put the result column **first**,
followed by the pairs. The older single-condition SUMIF and AVERAGEIF put it **last**. COUNTIF and COUNTIFS have no result column at
all, because they only count.

A few more rules apply to the whole family:

- Text criteria ignore case. For a case-sensitive count, use EXACT (Lesson 2.2) inside SUMPRODUCT (Lesson 2.1):
  `=SUMPRODUCT(--EXACT(C2:C2725,"Inpatient"))`.
- When the condition is a calculation on each row, such as the month of a date, LOSDays compared with a benchmark, or the length of a
  code, the -IFS functions can't do it directly. Add a helper column, or use SUMPRODUCT (Lesson 2.1) or FILTER (Lesson 4.1).

> ⚠️ **Closed workbooks.** COUNTIF, SUMIF, and the rest of the family return #VALUE! when their ranges point to *another workbook that
> is closed*. Open the source workbook, or bring the data into your report workbook (Power Query, Lesson 4.3).

Keyboard shortcuts for this lesson:

| Action | Windows | Mac |
|---|---|---|
| Cycle a reference through the \$ options | F4 | ⌘ + T |
| Extend a reference to the last filled row while typing a formula | Ctrl + Shift + ↓ | ⌘ + Shift + ↓ |
| Enter the same formula in every selected cell | Ctrl + Enter | ⌘ + Return |
| Fill down or fill right | Ctrl + D or Ctrl + R | ⌘ + D or ⌘ + R |
| Open the Insert Function dialog (Formula Builder on a Mac) | Shift + F3 | Shift + F3 |
| Format as a percentage or a date | Ctrl + Shift + % or Ctrl + Shift + # | ⌃ + Shift + % or ⌃ + Shift + # |

## 🧪 Hands-on practice

Download [`2.5-conditional-aggregation.xlsx`](2.5-conditional-aggregation.xlsx) and open the **Practice** sheet. Type each answer in
the yellow cell as a formula, with the sheet name in front of each range (`Encounters!C2:C2725`). The **Check** column turns green when
you're right.

<!-- BEGIN GENERATED: practice -->
The Encounters sheet holds rows 2–2725 and the Claims sheet holds rows 2–2667. Type text criteria exactly as the task spells them. Upper and lower case don't matter, but every other character does. Task 12 is filled on the Payer Mix sheet.

| # | Task | Hint |
|:-:|------|------|
| 1 | How many encounters were Emergency visits (EncounterType = Emergency)? | COUNTIF(range, criteria) |
| 2 | How many Inpatient encounters were at Cedar Ridge Medical Center? | COUNTIFS takes range/criteria pairs, and a row must pass every pair |
| 3 | What were the total charges (TotalCharges) of encounters billed to Medicare? Use PayerName = Medicare exactly, because Silverline Medicare Advantage is a separate payer. | SUMIF(range, criteria, sum_range) |
| 4 | How much was billed (BilledAmount) on claims with ClaimStatus = Denied and DenialReason = Authorization Required? Use the Claims sheet. | SUMIFS puts the sum_range FIRST |
| 5 | What was the average length of stay (LOSDays) of Inpatient encounters with PrimaryDxCode I50.9 (heart failure)? Round to 1 decimal place with ROUND. | AVERAGEIFS(average_range, criteria_range1, criteria1, …), wrapped in ROUND |
| 6 | How many claims were submitted in Q3 2025 (SubmitDate from 07/01/2025 through 09/30/2025)? Build the dates with DATE. | Use the same column twice, once for the start and once for the end |
| 7 | How many encounters had a primary diagnosis code in the circulatory chapter of ICD-10, meaning a PrimaryDxCode that starts with the letter I? | An asterisk matches any number of characters |
| 8 | How many Inpatient encounters had TotalCharges above the average TotalCharges of all Inpatient encounters? Calculate the average inside the formula rather than typing it. | Join an operator to a function: ">"&AVERAGEIF(…) |
| 9 | How many encounters at Bluestone Memorial Hospital were hospital stays, meaning EncounterType is Inpatient OR Observation? | COUNTIFS joins conditions with AND. For OR, add two counts together |
| 10 | What was the highest TotalCharges for an encounter whose DxDescription contains the word sepsis? | MAXIFS(max_range, criteria_range1, criteria1, …). "Contains" needs an asterisk on both sides |
| 11 | Collections wants the oldest claim still waiting on an insurance company. What is the earliest SubmitDate among claims with ClaimStatus = Pending and a PayerName other than Self-Pay? Enter it as a date. | MINIFS works on dates too, and "<>" means not equal to |
| 12 | On the Payer Mix sheet, fill the yellow grid B5:E12 with ONE COUNTIFS formula: the number of encounters for the payer in column A and the encounter type in row 4. Type it in B5, then copy it across and down. Type the data ranges as cell addresses with \$ signs instead of selecting Table columns (guide section 11 explains why). The gray cell adds up your grid. | Lock the data ranges fully. For each label, lock only the part (row or column) that must stay put |
| 13 | What was the 30-day readmission rate for the Cardiovascular service line? Divide the Cardiovascular Inpatient stays with Readmit30 = Y by all Cardiovascular Inpatient stays. Enter it as a percentage. | Rate = COUNTIFS(numerator) / COUNTIFS(denominator). Same filters, plus one more on top |
<!-- END GENERATED: practice -->

## ✅ Answer key

The workbook has a hidden **Answer Key** sheet (right-click any sheet tab → **Unhide…** → *Answer Key*). Its *Live result* column runs
each sample formula, so you can see it working. The same answers are below, collapsed so you don't see them by accident.

<!-- BEGIN GENERATED: answers -->
<details>
<summary><b>🔑 Show the answer key</b> — Try every task before opening this.</summary>

**1. Emergency visits**

- **Answer:** 913
- **Solution:** `=COUNTIF(Encounters!C2:C2725,"Emergency")`

COUNTIF checks every cell in the range against one criterion and counts the matches. A text criterion goes in double quotes and must match the whole cell, but upper and lower case don't matter, so "emergency" works too.

**2. Inpatient stays at Cedar Ridge**

- **Answer:** 93
- **Solution:**

```
=COUNTIFS(Encounters!C2:C2725,"Inpatient",Encounters!D2:D2725,"Cedar Ridge Medical Center")
```


COUNTIFS joins its conditions with AND: a row counts only when EncounterType is Inpatient *and* FacilityName is Cedar Ridge Medical Center. The order of the pairs doesn't matter.

**3. Total charges billed to Medicare**

- **Answer:** 6,830,385.52
- **Solution:** `=SUMIF(Encounters!N2:N2725,"Medicare",Encounters!O2:O2725)`

SUMIF tests one column (PayerName) and adds the matching rows of another (TotalCharges). Without wildcards, "Medicare" matches only cells that say exactly Medicare. The criterion `"*Medicare*"` would also pick up Silverline Medicare Advantage and return 11,665,796.53.

**4. Billed dollars denied for missing authorization**

- **Answer:** 824,904.94
- **Solution:** `=SUMIFS(Claims!G2:G2667,Claims!I2:I2667,"Denied",Claims!J2:J2667,"Authorization Required")`

SUMIFS starts with the column to add, then lists range/criteria pairs. The status criterion matters: 88 claims carry the reason Authorization Required, but only 77 of them are still plain Denied. The other 11 are all under appeal (ClaimStatus = Appealed), so they don't belong in a Denied total.

**5. Heart failure average length of stay**

- **Answer:** 5.2
- **Solution:**

```
=ROUND(AVERAGEIFS(Encounters!I2:I2725,Encounters!C2:C2725,"Inpatient",Encounters!J2:J2725,"I50.9"),1)
```


AVERAGEIFS averages only the rows that pass every condition. Leave out the Inpatient criterion and the heart failure Observation stays join the average, which drops to 4.9. Bluestone's benchmark (ExpectedLOS) for I50.9 is 4.1 days, so these stays run long.

**6. Claims submitted in Q3 2025**

- **Answer:** 652
- **Solution:** `=COUNTIFS(Claims!F2:F2667,">="&DATE(2025,7,1),Claims!F2:F2667,"<="&DATE(2025,9,30))`

A date range needs two conditions on the same column. `">="&DATE(2025,7,1)` joins the operator to the date's serial number, so the criterion works in any regional date setting. Because SubmitDate holds whole dates, `"<="&DATE(2025,9,30)` is safe here. `"<"&DATE(2025,10,1)` gives the same answer and also works on columns that include times.

**7. Circulatory diagnoses (I codes)**

- **Answer:** 315
- **Solution:** `=COUNTIF(Encounters!J2:J2725,"I*")`

The asterisk stands for "anything, or nothing," so `"I*"` matches I10, I21.4, I50.9, and every other code that starts with I. Wildcards only work on text, which is fine here because ICD-10 codes are text. Cross-check: COUNTIF on DxCategory = Circulatory also returns 315.

**8. Inpatient stays above the average inpatient charge**

- **Answer:** 233
- **Solution:**

```
=COUNTIFS(Encounters!C2:C2725,"Inpatient",Encounters!O2:O2725,">"&AVERAGEIF(Encounters!C2:C2725,"Inpatient",Encounters!O2:O2725))
```


AVERAGEIF works out the average Inpatient charge (34,168.89), and `">"&` turns it into the criterion text ">34168.89…". Only 233 of 688 stays (34%) beat the average, because a few very expensive stays pull the mean up (Lesson 2.4).

**9. Hospital stays at Bluestone Memorial (OR logic)**

- **Answer:** 624
- **Solution:**

```
=SUM(COUNTIFS(Encounters!D2:D2725,"Bluestone Memorial Hospital",Encounters!C2:C2725,{"Inpatient","Observation"}))
```


One COUNTIFS can't say OR, because every row would need to be Inpatient and Observation at once. Add two counts instead: `=COUNTIFS(…,"Inpatient")+COUNTIFS(…,"Observation")`. The SUM version does the same thing in one formula. The array constant {"Inpatient","Observation"} makes COUNTIFS return two counts, and SUM adds them. This is safe because no row can match both values.

**10. Most expensive sepsis encounter**

- **Answer:** 165,467.20
- **Solution:** `=MAXIFS(Encounters!O2:O2725,Encounters!K2:K2725,"*sepsis*")`

MAXIFS returns the largest value among the rows that pass every condition. `"*sepsis*"` means "sepsis with anything before or after it," which is how you write *contains* in a criterion. `"sepsis"` alone would only match a cell holding exactly that one word.

**11. Oldest pending insurance claim**

- **Answer:** 10/10/2025
- **Solution:** `=MINIFS(Claims!F2:F2667,Claims!I2:I2667,"Pending",Claims!C2:C2667,"<>Self-Pay")`

Dates are numbers, so the smallest SubmitDate is the oldest. `"<>Self-Pay"` keeps every payer except Self-Pay. Without it, the answer would be 01/04/2025: a self-pay balance the patient still owes, which belongs to a different work queue. If the cell shows a number like 45940, format it as a date.

**12. Payer mix grid filled with one formula**

- **Answer:** 2,724
- **Solution:** `=COUNTIFS(Encounters!$N$2:$N$2725,$A5,Encounters!$C$2:$C$2725,B$4)`

The data ranges never move, so they get full \$ signs. `$A5` lets the row change but always reads column A, and `B$4` lets the column change but always reads row 4. Every encounter has exactly one payer and one type, so a correct grid adds up to the number of rows on the Encounters sheet. Each row also matches the gray COUNTIF check in column G.

**13. Cardiovascular 30-day readmission rate**

- **Answer:** 19.5%
- **Solution:**

```
=COUNTIFS(Encounters!F2:F2725,"Cardiovascular",Encounters!C2:C2725,"Inpatient",Encounters!P2:P2725,"Y")/COUNTIFS(Encounters!F2:F2725,"Cardiovascular",Encounters!C2:C2725,"Inpatient")
```


The numerator is the denominator plus one extra condition (Readmit30 = Y). That's the pattern for every rate. The Inpatient criterion matters in both: the Cardiovascular service line also includes Cardiology Clinic visits, so `COUNTIF(ServiceLine,"Cardiovascular")` alone returns 119 instead of 77. Formal CMS measures also leave out some stays, such as deaths, transfers, and planned readmissions. This is the simple internal version.

</details>
<!-- END GENERATED: answers -->

## 🏆 Bonus challenge

<!-- BEGIN GENERATED: bonus -->
Bluestone's revenue cycle director is preparing for contract talks and wants a 2025 scorecard of inpatient claims by payer. Use these definitions. An **adjudicated** claim is any claim whose ClaimStatus is not Pending (the payer has made a decision). A **denied** claim is one whose ClaimStatus is Denied or Appealed, because an appealed claim was denied first. **Denial rate** = denied ÷ adjudicated. **Days to pay** = PaidDate − SubmitDate for every claim that has a PaidDate. Build the scorecard on the Scorecard sheet, one row per payer, using only claims whose EncounterType is Inpatient.

Work on the **Bonus** sheet of the workbook.

- **B1.** On the Claims sheet, fill the yellow DaysToPay column (L2:L2667) with PaidDate − SubmitDate, or an empty text string ("") when PaidDate is blank. Type one formula in L2. Claims is a Table, so Excel fills it down the whole column for you. If you click cells instead of typing their addresses, Excel writes [@PaidDate], which means the PaidDate on the same row. The gray cell averages your column across all claim types. *(Hint: IF(PaidDate="", "", PaidDate − SubmitDate))*
- **B2.** On the Scorecard sheet, fill columns B (adjudicated inpatient claims), C (denied inpatient claims), and D (denial rate) for all 8 payers, one formula per column copied down. A payer with no adjudicated inpatient claims must show n/a in column D instead of #DIV/0!. The gray cell shows the highest rate in your column D. (If it shows #DIV/0!, a row still needs the guard.) *(Hint: Column B: "<>Pending". Column C: the OR pattern from guide section 4. Column D: test column B for 0 before you divide)*
- **B3.** Which payer has the highest inpatient denial rate? Type its PayerName. *(Hint: Compare the rates in column D)*
- **B4.** Dollars at risk: for the payer you named in task B3, what is the total BilledAmount of its inpatient claims that are Denied or Appealed? *(Hint: The same OR trick works with SUMIFS)*
- **B5.** Fill column E of the Scorecard with each payer's average DaysToPay for inpatient claims (show n/a when a payer has none). How many more days, on average, does the payer you named in task B3 take to pay an inpatient claim than Medicare does? Subtract the full-precision averages (point at the cells instead of retyping the 1-decimal values that column E displays), then round the difference to 1 decimal place. *(Hint: AVERAGEIFS on your DaysToPay column, wrapped in IFERROR for the payer with no claims)*
<!-- END GENERATED: bonus -->

<!-- BEGIN GENERATED: bonus-answers -->
<details>
<summary><b>🔑 Show the bonus solution</b> — Give it a real try first!</summary>

**B1. DaysToPay helper column**

- **Answer:** 26.97
- **Solution:** `=IF(K2="","",K2-F2)`

The -IFS functions can only test and average columns that already exist. None of them can subtract two columns row by row, so a **helper column** does the subtraction first, and then AVERAGEIFS can average it by payer. Without the IF, an unpaid claim computes 0 − SubmitDate, a large negative number of days that wrecks every average. The empty text "" is skipped by AVERAGE and AVERAGEIFS. The key's live formula shows a no-helper version: the average of the differences is (sum of PaidDates − sum of their SubmitDates) ÷ the number of paid claims.

**B2. Denial rates (with a divide-by-zero guard)**

- **Answer:** 30.9%
- **Solution:**

B5: `=COUNTIFS(Claims!$C$2:$C$2667,$A5,Claims!$D$2:$D$2667,"Inpatient",Claims!$I$2:$I$2667,"<>Pending")`

C5: `=SUM(COUNTIFS(Claims!$C$2:$C$2667,$A5,Claims!$D$2:$D$2667,"Inpatient",Claims!$I$2:$I$2667,{"Denied","Appealed"}))`

D5: `=IF(B5=0,"n/a",C5/B5)`

Copy each one down to row 12.


Column B uses `"<>Pending"` to count every decision. Column C needs OR logic (Denied or Appealed), so it wraps COUNTIFS with an array constant in SUM. `$A5` keeps each formula reading its own payer as you copy down. Workers' Compensation had no inpatient claims, so its row divides 0 by 0. `IF(B5=0,"n/a",…)` tests for exactly that case. `IFERROR(C5/B5,"n/a")` also works, but it would hide any other mistake too. Counting only Denied (not Appealed) would understate the worst payer's rate at 19.1%.

**B3. Payer with the highest denial rate**

- **Answer:** Silverline Medicare Advantage
- **Solution:** Read it from column D of your scorecard: the largest rate belongs to **Silverline Medicare Advantage**.

Silverline Medicare Advantage denied 30.9% of its adjudicated inpatient claims. Next is State Medicaid at 17.7%, and traditional Medicare is at 7.1%. A formula can find the name for you too: `=INDEX(A5:A12,MATCH(MAX(D5:D12),D5:D12,0))` on the Scorecard sheet. Lesson 2.6 teaches INDEX and MATCH.

**B4. Billed dollars at risk with the worst payer**

- **Answer:** 1,031,768.03
- **Solution:**

```
=SUM(SUMIFS(Claims!G2:G2667,Claims!C2:C2667,"Silverline Medicare Advantage",Claims!D2:D2667,"Inpatient",Claims!I2:I2667,{"Denied","Appealed"}))
```


SUMIFS with the array constant returns two totals (Denied and Appealed), and SUM adds them. About 1.0 million dollars of inpatient charges are tied up in denials with this one payer, which is the number that gets a contract meeting's attention.

**B5. How much slower the worst payer pays**

- **Answer:** 14.1
- **Solution:**

```
=ROUND(AVERAGEIFS(Claims!L2:L2667,Claims!C2:C2667,"Silverline Medicare Advantage",Claims!D2:D2667,"Inpatient")-AVERAGEIFS(Claims!L2:L2667,Claims!C2:C2667,"Medicare",Claims!D2:D2667,"Inpatient"),1)
```


E5 is `=IFERROR(AVERAGEIFS(Claims!$L$2:$L$2667,Claims!$C$2:$C$2667,$A5,Claims!$D$2:$D$2667,"Inpatient"),"n/a")`, copied down. AVERAGEIFS skips the "" cells of unpaid claims, and IFERROR turns Workers' Compensation's #DIV/0! into n/a. Here IFERROR is the right guard, because AVERAGEIFS has no separate denominator you could test. Silverline Medicare Advantage averages 28.95 days and Medicare 14.82, so `=ROUND(Scorecard!E6-Scorecard!E5,1)` in the answer cell gives the same 14.1. Compared with traditional Medicare, Silverline Medicare Advantage denies about four times as often *and* takes about two weeks longer to pay.

</details>
<!-- END GENERATED: bonus-answers -->

## Key takeaways

- COUNTIFS, SUMIFS, AVERAGEIFS, MAXIFS, and MINIFS keep only the rows that pass every condition. The -IFS functions put the column to
  add, average, or search **first**, while SUMIF and AVERAGEIF put it last.
- Criteria are text: an operator and a value inside quotes (`">=10"`), or an operator in quotes joined to a cell or function with `&`
  (`">="&R2`). Build date criteria with DATE or a date cell, and end a period with `"<"` the first day of the next one.
- For OR, add counts together, either as COUNTIFS + COUNTIFS or as `SUM(COUNTIFS(…,{"a","b"}))`. When the conditions are on different
  columns, subtract the overlap.
- Wildcards (`*`, `?`, `~`) match part of a text cell. Check which values a wildcard caught before you trust its count.
- A rate is COUNTIFS of the right population plus the outcome, divided by COUNTIFS of that population. Guard it with
  `IF(denominator=0,"n/a",…)`, and remember that MAXIFS and MINIFS return 0 when nothing matches.
- One COUNTIFS with fully locked data ranges and mixed-reference labels (`$A5`, `B$4`) fills a whole summary grid.

<!-- BEGIN GENERATED: nav -->
---

⬅️ **Previous:** [2.4 Math & Statistical Functions](../04-math-statistical-functions/README.md) · 🏠 [Course home](../../README.md) · **Next:** [2.6 Lookups: VLOOKUP, INDEX/MATCH & XLOOKUP](../06-lookup-functions/README.md) ➡️
<!-- END GENERATED: nav -->
