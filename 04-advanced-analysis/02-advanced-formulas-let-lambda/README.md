# Lesson 4.2 · Advanced Formulas: LET, LAMBDA & Array Logic

> **Level:** Advanced · **Time:** about 65 minutes · **Workbook:** [`4.2-advanced-formulas-let-lambda.xlsx`](4.2-advanced-formulas-let-lambda.xlsx)
> **Data:** All 5,586 inpatient stays at Bluestone's three hospitals, Jan 2024–Dec 2025 (with the Readmit30 flag removed so you can rebuild it), the diagnosis lookup with benchmark LOS, 1,018 lab results from Ashby Falls stays discharged Jul–Dec 2025, and the Bluestone Memorial ICU's daily census for 2024–2025.

The quality director wants 30-day readmission rates by hospital, calculated with the exclusions that **CMS** (the Centers for
Medicare & Medicaid Services) uses, and she wants them to update every month without anyone rebuilding a PivotTable. That question can't be answered with SUMIFS alone. Each stay has to be
compared with every *other* stay of the same patient, some stays have to be excluded, and the result has to be divided
hospital by hospital. In this lesson you'll learn the formula tools that make questions like this a single, readable,
reusable formula: boolean array logic, **LET** for naming the steps, **LAMBDA** for turning the steps into your own
functions, the LAMBDA helpers (MAP, BYROW, SCAN, REDUCE), and the auditing tools that let you trust what you built.

## What you'll learn

- Write multi-condition array logic with SUMPRODUCT and boolean math
- Make complex formulas readable and fast with LET
- Create reusable custom functions with LAMBDA and the Name Manager
- Use MAP, REDUCE, SCAN, BYROW, and audit formulas with Evaluate Formula

## 📖 Guide

### 1. The data and two definitions

| Sheet | Excel Table | What's in it |
|---|---|---|
| **Stays** | `tblStays` | Every inpatient stay, 2024–2025 (5,586 rows): IDs, facility, primary diagnosis, admit and discharge date-times, disposition, charges, and an empty yellow **Readmit30** column for Task 5 |
| **Diagnoses** | `tblDx` | 51 ICD-10 codes with a description, a category, and **ExpectedLOS** (a benchmark length of stay in days) |
| **Labs** | `tblLabs` | 1,018 lab results from Ashby Falls stays discharged July–December 2025, sorted oldest to newest |
| **Census** | `tblCensus` | Bluestone Memorial ICU (20 staffed beds), one row per day for 2024–2025 (731 rows) |
| **Audit** | none | A colleague's broken formula for Task 13 |
| **Sandbox** | none | An empty sheet for watching array formulas spill |

Almost every formula in this lesson uses **structured references** (Lesson 3.1). `tblStays[FacilityID]` means the whole FacilityID
column of the Stays table, which is the same cells as `Stays!$C$2:$C$5587`. Structured references are easier to read in long
formulas, and they grow automatically when rows are added.

The tasks use two definitions throughout:

- **LOS days** (length of stay) = discharge date − admit date, counted in midnights, with a minimum of 1 day. AdmitDateTime and
  DischargeDateTime hold a date *and* a time, so `INT()` strips the time first: `INT(discharge) - INT(admit)`. A patient
  admitted 01/04/2024 at 19:46 and discharged 01/07/2024 at 16:29 has 3 LOS days. A stay admitted and discharged on the same
  date has 0 midnights, so it counts as 1 day.
- **30-day readmission**: a stay is followed by a readmission when the same patient has a *later* inpatient admission whose
  admit **date** is 0–30 days after this stay's discharge **date**, at any Bluestone hospital. This is the rule behind the
  `Readmit30` column in the full dataset (see the [data dictionary](../../data/README.md)).

### 2. Boolean arrays: how TRUE and FALSE become numbers

Compare a whole column with a value and Excel returns an **array**: one result per row, held in memory. Here's what two
comparisons return for the first five stays on the Stays sheet:

| Row | FacilityID | DischargeDisposition | A: `(FacilityID="F01")` | B: `(Disposition="Home/Self-Care")` | `A*B` (AND) | `A+B` (OR) |
|:-:|:-:|---|:-:|:-:|:-:|:-:|
| 2 | F01 | Home Health | TRUE | FALSE | 0 | 1 |
| 3 | F03 | Home/Self-Care | FALSE | TRUE | 0 | 1 |
| 4 | F02 | Home/Self-Care | FALSE | TRUE | 0 | 1 |
| 5 | F01 | Home/Self-Care | TRUE | TRUE | **1** | **2** |
| 6 | F03 | Home/Self-Care | FALSE | TRUE | 0 | 1 |

Arithmetic turns TRUE into 1 and FALSE into 0. You used that rule in Lessons 2.1 and 2.4 to count with
`SUMPRODUCT(--(condition))`. Combining several conditions this way is called **boolean math**:

| Expression | Result | Use it for |
|---|---|---|
| `=TRUE*TRUE` | 1 | **AND**: a row is 1 only when every condition is TRUE |
| `=TRUE*FALSE` | 0 | |
| `=TRUE+FALSE` | 1 | **OR**: a row is at least 1 when any condition is TRUE |
| `=TRUE+TRUE` | 2 | ⚠️ a row that meets both OR conditions counts twice |
| `=--TRUE` | 1 | The **double unary** `--` converts TRUE/FALSE to 1/0 without changing anything else |
| `=1-(A1="F01")` | 0 or 1 | **NOT** |

So `(tblStays[FacilityID]="F01")*(tblStays[DischargeDisposition]="Home/Self-Care")` is an array of 5,586 ones and zeros
with a 1 on every F01 stay that went home, and `SUM` or `SUMPRODUCT` of that array counts them.

> ⚠️ **AND() and OR() don't work row by row.** `AND(tblStays[FacilityID]="F01", …)` looks at all 5,586 values at once and
> returns a single TRUE or FALSE. Inside array formulas, use `*` for AND and `+` for OR. Section 7 shows how MAP and BYROW
> let you use AND and OR safely.

> ⚠️ **Numbers never equal text.** `MONTH(tblStays[AdmitDateTime])` returns numbers, so `MONTH(…)="12"` is FALSE on every
> row, even for December admissions. Write `MONTH(…)=12` without quotes. COUNTIFS accepts a criterion like `"12"` because it
> converts criteria strings to numbers, but a plain `=` comparison doesn't convert anything.

> 💡 **Tip:** To see a short array, type a test formula such as `=SUM((Stays!C2:C6="F01")*1)`, select the part
> `Stays!C2:C6="F01"` in the formula bar, and press **F9** (Mac: **Fn + F9**). Excel shows `{TRUE;FALSE;FALSE;TRUE;FALSE}`.
> Press **Esc** afterwards. If you press Enter instead, Excel replaces that part of the formula with the constant values. F9
> can't display a whole 5,586-row column, because the result would be longer than Excel's 8,192-character formula limit. To
> inspect a full column, type the piece into a cell on the **Sandbox** sheet and let it spill.

### 3. SUMPRODUCT: counts and sums with any conditions

```
=SUMPRODUCT(array1, [array2], …)          multiplies the arrays row by row, then adds the products
=SUMPRODUCT((condition1)*(condition2))                       count of rows meeting both conditions
=SUMPRODUCT((condition1)*(condition2), values)               sum of values on those rows
=SUMPRODUCT((condition1)*(condition2), values)/SUMPRODUCT((condition1)*(condition2))   conditional average
```

**Worked example:** how many Cedar Ridge (F03) stays discharged in 2025 went to a Skilled Nursing Facility, and what were
their charges?

```
=SUMPRODUCT((tblStays[FacilityID]="F03")*(YEAR(tblStays[DischargeDateTime])=2025)*(tblStays[DischargeDisposition]="Skilled Nursing Facility"))
```

The result is **51**. Add the charges column as a second argument and the same conditions give **$1,860,774.45**:

```
=SUMPRODUCT((tblStays[FacilityID]="F03")*(YEAR(tblStays[DischargeDateTime])=2025)*(tblStays[DischargeDisposition]="Skilled Nursing Facility"), tblStays[TotalCharges])
```

You met SUMPRODUCT in Lesson 2.4 and COUNTIFS and SUMIFS in Lesson 2.5. COUNTIFS and SUMIFS are faster and simpler when
they can do the job, so use SUMPRODUCT when they can't:

| Situation | COUNTIFS / SUMIFS | SUMPRODUCT |
|---|---|---|
| Simple AND conditions on columns as they are (`FacilityID = "F03"`) | ✅ best choice | ✅ works |
| A condition on a **calculated** value: `YEAR(date)=2025`, `WEEKDAY(date,2)>5`, LOS = `INT(dis)-INT(adm)` | ❌ needs a helper column | ✅ calculate inside the formula |
| **OR** across different columns (sepsis OR long stay) | ❌ awkward, and overlaps double count | ✅ `((A)+(B))>0` |
| Multiplying two columns before adding (dose × unit cost, weighted averages) | ❌ | ✅ |
| Speed on very large tables | ✅ faster | slower, because every row is evaluated for every condition |

`WEEKDAY(date, 2)` numbers the days Monday = 1 through Sunday = 7, so `WEEKDAY(…,2)>5` is TRUE on Saturdays and Sundays.

**OR logic without double counting.** Count Ashby Falls (F02) stays discharged in 2025 that were sepsis (A41.9) *or* had LOS
days of 10 or more:

```
=SUMPRODUCT((((tblStays[PrimaryDxCode]="A41.9")+(INT(tblStays[DischargeDateTime])-INT(tblStays[AdmitDateTime])>=10))>0)
            *(tblStays[FacilityID]="F02")*(YEAR(tblStays[DischargeDateTime])=2025))
```

There are 38 sepsis stays and 20 long stays, but 10 stays are both. Without the `>0`, those 10 stays score 2 and the formula
returns 58. With `((A)+(B))>0`, each stay is either TRUE or FALSE, and the correct answer is **48**.

**Conditional averages.** Divide a SUMPRODUCT sum by a SUMPRODUCT count. The average charge for F02 sepsis stays discharged
in 2025 is **$73,787.82**:

```
=SUMPRODUCT((tblStays[FacilityID]="F02")*(tblStays[PrimaryDxCode]="A41.9")*(YEAR(tblStays[DischargeDateTime])=2025), tblStays[TotalCharges])
 /SUMPRODUCT((tblStays[FacilityID]="F02")*(tblStays[PrimaryDxCode]="A41.9")*(YEAR(tblStays[DischargeDateTime])=2025))
```

That repeats the conditions twice, which is exactly the problem LET solves in section 5.

**Testing against a list.** To ask whether each row's value is one of several values, MATCH the whole column against an
**array constant**, a list typed between braces (Lesson 2.5). MATCH returns a position for every value it finds and `#N/A` for
every value it doesn't, so ISNUMBER and ISNA turn its result into TRUE/FALSE:

| Test | TRUE on rows whose value is… |
|---|---|
| `ISNUMBER(MATCH(column, {list}, 0))` | on the list |
| `ISNA(MATCH(column, {list}, 0))` | **not** on the list |

Of the 2,905 stays discharged in 2025, **2,402** went home (Home/Self-Care or Home Health) and the other **503** didn't:

```
=SUMPRODUCT(ISNUMBER(MATCH(tblStays[DischargeDisposition], {"Home/Self-Care","Home Health"}, 0))*(YEAR(tblStays[DischargeDateTime])=2025))
=SUMPRODUCT(ISNA(MATCH(tblStays[DischargeDisposition], {"Home/Self-Care","Home Health"}, 0))*(YEAR(tblStays[DischargeDateTime])=2025))
```

A list is shorter than one comparison per value, and you can add a value to it without touching the rest of the formula.

> 📋 **SUM or SUMPRODUCT?** In Microsoft 365 and Excel 2021+, `=SUM((condition1)*(condition2))` works too, because these
> versions evaluate arrays in any formula. In Excel 2019 and earlier, SUM needed **Ctrl + Shift + Enter** (Mac: **⌘ + Shift +
> Return**) to do array math, but SUMPRODUCT never did. That's why you'll see SUMPRODUCT in older workbooks.

> ⚠️ **Every array must be the same size.** `tblStays[FacilityID]` and `Stays!$F$2:$F$5000` don't line up, so the formula
> returns an error (`#VALUE!` or `#N/A`). A range that's the right size but starts one row too low is worse, because it returns a
> wrong number with no error. Structured references avoid both problems.

> ⚠️ **Avoid whole-column references** such as `C:C` inside SUMPRODUCT. Excel then evaluates more than a million rows for every
> condition. Use Table columns instead.

### 4. More array logic: lookups, FREQUENCY, and comparing each row with the whole table

#### 4a. Lookups with several conditions

XLOOKUP (Lesson 2.6) normally looks up one value in one column. To match on several columns, build a 1/0 array with boolean
math and look up the number **1** in it:

```
=XLOOKUP(1, (condition1)*(condition2), return_array, "Not found", 0, search_mode)
```

The Labs sheet is sorted oldest to newest. Patient PT11261 has three white blood cell counts (TestCode WBC):

```
=XLOOKUP(1, (tblLabs[PatientID]="PT11261")*(tblLabs[TestCode]="WBC"), tblLabs[ResultValue], "Not found")       → 7.3
=XLOOKUP(1, (tblLabs[PatientID]="PT11261")*(tblLabs[TestCode]="WBC"), tblLabs[ResultValue], "Not found", 0, -1) → 21.4
```

The first formula searches top-down (search_mode 1, the default) and returns the *first* match from 10/26/2025. The second
uses search_mode **-1** to search bottom-up, so it returns the *most recent* result from 12/04/2025. On a sorted sheet, -1
means "latest".

> 📋 **Before XLOOKUP**, the same lookup was written `=INDEX(return_range, MATCH(1, (condition1)*(condition2), 0))`. In
> Excel 2019 and earlier it has to be confirmed with Ctrl + Shift + Enter (Mac: ⌘ + Shift + Return).

#### 4b. FREQUENCY: distributions in one formula

```
=FREQUENCY(data_array, bins_array)
```

**FREQUENCY** counts how many values fall into each bin. Each bin number is an *upper limit*, and FREQUENCY returns one extra
count for everything above the last bin. For LOS days of F02 stays discharged in 2025 with the bins `{2,4,7}`:

| Position | Bin | Counts LOS days that are… | Count |
|:-:|:-:|---|:-:|
| 1 | 2 | ≤ 2 | 74 |
| 2 | 4 | > 2 and ≤ 4 | 153 |
| 3 | 7 | > 4 and ≤ 7 | 102 |
| 4 | (overflow) | > 7 | 45 |

```
=FREQUENCY(FILTER(INT(tblStays[DischargeDateTime])-INT(tblStays[AdmitDateTime]),
                  (tblStays[FacilityID]="F02")*(YEAR(tblStays[DischargeDateTime])=2025)), {2,4,7})
```

This formula spills four numbers, so try it on the Sandbox sheet. To return just one of the counts, wrap it in INDEX:
`=INDEX(FREQUENCY(…), 4)` returns 45. FILTER (Lesson 4.1) keeps only the stays you want. The older pattern
`FREQUENCY(IF(condition, values), bins)` also works, because FREQUENCY ignores the FALSE values that IF returns.

#### 4c. COUNTIFS that compares each row with the whole table

Some questions compare a row with *other* rows: "Did this patient come back within 30 days?" COUNTIFS can answer it,
because its criteria can come from the current row while its ranges cover the whole table. Patient PT13853 shows the idea:

| EncounterID | Facility | Admit | Discharge | Next admission | Days later | Readmitted? |
|---|:-:|---|---|---|:-:|:-:|
| ENC100149 | F01 | 01/04/2024 19:46 | 01/07/2024 16:29 | 01/24/2024 (at F03) | 17 | 1 |
| ENC100812 | F03 | 01/24/2024 20:05 | 01/29/2024 14:06 | 11/30/2024 | 306 | 0 |
| ENC109619 | F03 | 11/30/2024 23:15 | 12/04/2024 16:31 | none | | 0 |

The first stay counts as readmitted even though the patient came back to a *different* hospital, because the rule is "any
Bluestone hospital." To test one row, you count the rows that meet three conditions:

1. Same patient: `tblStays[PatientID]` equals this row's PatientID.
2. Admitted after this discharge: `tblStays[AdmitDateTime]` is greater than this row's DischargeDateTime.
3. Admit date no more than 30 days later: `tblStays[AdmitDateTime]` is less than `INT(discharge) + 31`, which is midnight at the
   start of the 31st day after the discharge date.

A count above 0 means the stay was followed by a readmission. In a Table, `[@PatientID]` means "this row's PatientID," so the
criteria come from the current row. Criteria such as `">"&[@DischargeDateTime]` join a comparison operator to a value, as in
Lesson 2.5. Type the formula in the first cell of an empty Table column and Excel fills it down every row as a **calculated
column** (Lesson 3.1).

**Every row in one formula.** COUNTIFS also accepts a whole column as a *criteria* argument. It then counts once for each value
in that column and returns an array with one count per stay. On the Sandbox sheet,
`=COUNTIFS(tblStays[PatientID], tblStays[PatientID])` spills 5,586 numbers: how many stays each row's patient has. Test the
array and add it up to get one answer:

```
=SUM(--(COUNTIFS(tblStays[PatientID], tblStays[PatientID])>1))
```

The result is **4,414**, the stays that belong to patients with more than one stay. Criteria that you build with `&` work the
same way: `">"&tblStays[DischargeDateTime]` is an array of 5,586 criteria strings, one for each stay. Only the *criteria*
arguments can be arrays. The *range* arguments must still be real ranges, such as Table columns (section 5 shows how this
plays out inside LET).

> ⚠️ **Dates, not date-times.** The rule counts calendar days. If you test `AdmitDateTime <= DischargeDateTime + 30`, a patient
> discharged at 10:00 and readmitted at 15:00 on day 30 is missed, because 15:00 is later than 10:00. `INT()` puts both sides
> on whole dates.

> 📋 **Performance:** a COUNTIFS column on 5,586 rows compares every row with every other row, about 31 million comparisons
> per criterion, and the one-formula version does the same work. Excel handles that in a second or two. On much larger tables,
> consider Power Query (Lesson 4.3) instead.

#### 4d. CHOOSE, and why INDEX beats OFFSET and INDIRECT

```
=CHOOSE(index_num, value1, value2, …)       returns value1 when index_num is 1, value2 when it's 2, …
```

**CHOOSE** picks one of several values or ranges by position. With a selector cell B1 that holds 1, 2, or 3,
`=SUM(CHOOSE(B1, tblCensus[Admissions], tblCensus[Discharges], tblCensus[MidnightCensus]))` totals whichever column the user
picked.

**OFFSET** and **INDIRECT** also build references on the fly, but they're **volatile**: Excel recalculates them after *every*
change anywhere in the workbook, which slows a workbook with thousands of formulas. INDEX returns a reference too, and it isn't
volatile. To average the ICU's last 7 days of census:

| Approach | Formula | Volatile? |
|---|---|:-:|
| OFFSET | `=AVERAGE(OFFSET(Census!E2, ROWS(tblCensus[MidnightCensus])-7, 0, 7, 1))` | yes |
| INDEX:INDEX | `=AVERAGE(INDEX(tblCensus[MidnightCensus], ROWS(tblCensus[MidnightCensus])-6):INDEX(tblCensus[MidnightCensus], ROWS(tblCensus[MidnightCensus])))` | no |
| TAKE (365/2024) | `=AVERAGE(TAKE(tblCensus[MidnightCensus], -7))` | no |

All three return 18.71 (the last seven nights of 2025). INDEX joined to INDEX with a colon builds a range from the 725th to the
731st cell of the column.

### 5. LET: name the steps of a formula

```
=LET(name1, value1, [name2, value2, …], calculation)
```

**LET** assigns names to values *inside* one formula. Each name can use the names before it, and the last argument is the
calculation that uses them. Compare two ways to get the average LOS days of heart failure stays (I50.9) at Ashby Falls (F02)
discharged in 2025:

**Without LET**, the LOS expression appears twice and the conditions twice:

```
=SUMPRODUCT(IF(INT(tblStays[DischargeDateTime])-INT(tblStays[AdmitDateTime])<1, 1, INT(tblStays[DischargeDateTime])-INT(tblStays[AdmitDateTime]))
   *(tblStays[FacilityID]="F02")*(tblStays[PrimaryDxCode]="I50.9")*(YEAR(tblStays[DischargeDateTime])=2025))
 /SUMPRODUCT((tblStays[FacilityID]="F02")*(tblStays[PrimaryDxCode]="I50.9")*(YEAR(tblStays[DischargeDateTime])=2025))
```

**With LET**, every step has a name and is written once:

```
=LET(
  adm, tblStays[AdmitDateTime],
  dis, tblStays[DischargeDateTime],
  nights, INT(dis)-INT(adm),
  los, IF(nights<1, 1, nights),
  keep, (tblStays[FacilityID]="F02")*(tblStays[PrimaryDxCode]="I50.9")*(YEAR(dis)=2025),
  SUM(los*keep)/SUM(keep)
)
```

Both return **5.17**. Read the LET version from top to bottom: `nights` is the midnight count for every stay, `los` applies
the 1-day minimum, `keep` is 1 on the 36 stays you want, and the last line is a conditional average. `los*keep` zeroes out
every other stay, and `SUM(keep)` counts the kept stays.

Why LET is worth it:

| Benefit | Why |
|---|---|
| **Readable** | Names like `los` and `keep` say what each piece means |
| **Faster** | Excel calculates each name once. Without LET, a repeated expression is calculated every time it appears |
| **Safer to edit** | Change the LOS rule in one place instead of hunting for every copy |
| **Easy to debug** | Replace the last argument with any name to see that step's result (see section 8) |

Rules for names:

- Start with a letter, and use only letters, numbers, underscores, and periods. No spaces.
- A name can't look like a cell address. `los` is fine, but `los1` isn't, because LOS1 is a real cell (column LOS, row 1).
- Avoid names that match a function, such as `days`, `rate`, or `index`, so a reader never wonders which one you mean.
- Names exist only inside their own formula. Another cell's LET can reuse `los` for something else.

> ⚠️ **Ranges stay ranges, calculations become arrays.** A name that points straight at a column (`adm, tblStays[AdmitDateTime]`)
> is still a range, so you can pass it to COUNTIFS. A calculated name such as `nights` is an array, and COUNTIFS, SUMIFS, and
> COUNTIF only accept ranges in their range arguments, so `COUNTIF(nights, ">5")` fails. Use `SUM(--(nights>5))` instead.

**Lookups inside LET.** Give XLOOKUP a whole column as its lookup value and it returns one result per row. That's how you
attach the benchmark ExpectedLOS to every stay without a helper column. For pneumonia (J18.9) stays discharged in 2025,
observed LOS days ÷ expected LOS days is **1.23**:

```
=LET(
  adm, tblStays[AdmitDateTime],
  dis, tblStays[DischargeDateTime],
  nights, INT(dis)-INT(adm),
  los, IF(nights<1, 1, nights),
  explos, XLOOKUP(tblStays[PrimaryDxCode], tblDx[DxCode], tblDx[ExpectedLOS]),
  keep, (tblStays[PrimaryDxCode]="J18.9")*(YEAR(dis)=2025),
  SUM(los*keep)/SUM(explos*keep)
)
```

This ratio is the **LOS index**, often written **O/E** (observed ÷ expected). A value above 1.00 means patients stayed longer
than the benchmark. Divide the totals rather than averaging each stay's ratio, so long stays carry their proper weight.

> 💡 **Tip:** Long formulas are easier to read on several lines. Press **Alt + Enter** (Mac: **⌃ + ⌥ + Return**) inside the
> formula bar to start a new line, and add spaces to indent. Press **Ctrl + Shift + U** (Mac: **⌃ + Shift + U**) to expand the
> formula bar so you can see every line.

### 6. LAMBDA: build your own functions

```
=LAMBDA([parameter1, parameter2, …], calculation)
```

**LAMBDA** turns a calculation into a function with **parameters** (named inputs). You build one in three steps.

**Step 1. Write the calculation as a LET**, with the inputs as names. You already did this for LOS. The whole LET then
becomes the LAMBDA's calculation, for example `LAMBDA(weight_lb, height_in, LET(bmi, 703*weight_lb/height_in^2, ROUND(bmi, 1)))`.

**Step 2. Test it in a cell.** A LAMBDA on its own returns `#CALC!`, because nothing has called it yet. Add the arguments in a
second pair of parentheses right after it:

```
=LAMBDA(weight_lb, height_in, 703*weight_lb/height_in^2)(180, 70)          → 25.8 (body mass index)
=LAMBDA(admit, discharge, INT(discharge)-INT(admit))(Stays!E2, Stays!F2)  → 3 (midnights for the first stay)
```

**Step 3. Save it in the Name Manager** so you can call it by name anywhere in the workbook:

1. Choose **Formulas → Name Manager** (Windows: **Ctrl + F3**) and click **New**. On a Mac, **Formulas → Define Name** opens the
   same fields.
2. Type the function's name in **Name**, for example `MIDNIGHTS`. The same rules as LET names apply.
3. Leave **Scope** as *Workbook*.
4. In **Comment**, describe the inputs and the result. Excel shows the comment as a tooltip when you type the function.
5. In **Refers to**, paste the LAMBDA *without* the test arguments: `=LAMBDA(admit, discharge, INT(discharge)-INT(admit))`.
6. Click **OK**. Now `=MIDNIGHTS(Stays!E2, Stays!F2)` returns 3.

A named LAMBDA works on whole columns as well as single cells, because the parameters simply receive whatever you pass:
`=SUM(MIDNIGHTS(tblStays[AdmitDateTime], tblStays[DischargeDateTime]))` adds up the midnights of all 5,586 stays.

> ⚠️ **Aggregating functions collapse arrays.** Inside a LAMBDA that you call with whole columns, `MAX`, `MIN`, `SUM`, `AND`,
> and `OR` act on the *entire* column. `MAX(1, INT(discharge)-INT(admit))` doesn't apply a minimum to each stay. It returns
> one number, the longest stay in the table. Use `IF(x<1, 1, x)`, which works row by row, or call the function through MAP
> (section 7), which passes one value at a time.

More things you can do with LAMBDA:

- **Optional parameters.** Put a parameter in square brackets to make it optional, and test it with **ISOMITTED**:
  `=LAMBDA(result, [low], IF(ISOMITTED(low), result<70, result<low))` flags a glucose result below 70 mg/dL unless you pass
  another limit.
- **Recursion.** A named LAMBDA can call itself by name. Excel limits how deep those calls can go, so prefer REDUCE or SCAN for
  long loops.
- **Reuse in other workbooks.** Names live inside the workbook. Copy the *Refers to* text into the other workbook's Name Manager.
  For long LAMBDAs, Microsoft's free **Excel Labs** add-in includes an *Advanced Formula Environment* with a proper code editor.

| Error | Usual cause |
|---|---|
| `#CALC!` | The LAMBDA was never called (no arguments after it), or a helper's LAMBDA returned more than one value per call |
| `#VALUE!` | The function was called with the wrong number of arguments |
| `#NAME?` | The name is misspelled, or the Excel version doesn't support LAMBDA |

> 📋 **Version note:** LAMBDA and the helper functions in section 7 need **Microsoft 365** (Windows, Mac, or web) or **Excel 2024**.
> Excel 2021 has LET and XLOOKUP but not LAMBDA. A workbook that uses LAMBDA shows `#NAME?` in older versions.

### 7. The LAMBDA helpers: MAP, BYROW, BYCOL, SCAN, REDUCE, MAKEARRAY

These functions run a LAMBDA many times for you, each in a different pattern:

| Function | Calls your LAMBDA… | The LAMBDA receives | Returns | Hospital example |
|---|---|---|---|---|
| **MAP** | once per item | one value from each array you pass | an array the same size | Flag each stay with AND/OR logic |
| **BYROW** | once per row | a whole row (1 × n) | one value per row | Is this lab result outside its range? |
| **BYCOL** | once per column | a whole column | one value per column | Maximum of each census column |
| **SCAN** | once per item, in order | the running result so far and the next item | every running result | Running totals and streaks |
| **REDUCE** | once per item, in order | the running result so far and the next item | only the final result | Loop over a list of codes |
| **MAKEARRAY** | once per cell of a grid you size | a row number and a column number | the grid | Facility × year summary |

#### MAP

```
=MAP(array1, [array2, …], LAMBDA(item1, [item2, …], calculation))
```

MAP passes one item at a time, so AND, OR, and MAX behave. Count F02 stays discharged in 2025 that were both long (LOS days of
10 or more) and expensive ($100,000 or more):

```
=SUM(MAP(tblStays[FacilityID], tblStays[AdmitDateTime], tblStays[DischargeDateTime], tblStays[TotalCharges],
     LAMBDA(fac, adm, dis, chg, IF(AND(fac="F02", YEAR(dis)=2025, INT(dis)-INT(adm)>=10, chg>=100000), 1, 0))))
```

The result is **12**. The arrays must all be the same size, and the LAMBDA needs one parameter per array, in the same order.

#### BYROW and BYCOL

```
=BYROW(array, LAMBDA(r, calculation))
=BYCOL(array, LAMBDA(col, calculation))
```

BYROW hands the LAMBDA one row of a multi-column range. Use `INDEX(r, n)` to pick the nth cell of that row `r`. On how many ICU
days did discharges outnumber admissions?

```
=SUM(BYROW(tblCensus[[Admissions]:[Discharges]], LAMBDA(r, IF(INDEX(r,2)>INDEX(r,1), 1, 0))))
```

The result is **189**. `tblCensus[[Admissions]:[Discharges]]` is the structured reference for two adjacent columns (Lesson 3.1).
BYROW really pays off when each row needs a function that would otherwise swallow the whole range, such as
`BYROW(range, LAMBDA(r, AND(r>0)))` ("every cell in this row is positive") or `BYROW(range, LAMBDA(r, MAX(r)))`.
`=BYCOL(tblCensus[[Admissions]:[MidnightCensus]], LAMBDA(col, MAX(col)))` returns `{12, 12, 20}`, the maximum of each column.

#### SCAN

```
=SCAN(initial_value, array, LAMBDA(accumulator, item, calculation))
```

SCAN carries a running result, called the **accumulator**, down the array. Each call receives the accumulator and the next item
and returns the new accumulator. SCAN returns every intermediate value. A running total of ICU admissions:

```
=SCAN(0, tblCensus[Admissions], LAMBDA(total, x, total + x))
```

| CensusDate | Admissions | Running total |
|---|:-:|:-:|
| 01/01/2024 | 7 | 7 |
| 01/02/2024 | 6 | 13 |
| 01/03/2024 | 4 | 17 |
| 01/04/2024 | 6 | 23 |
| 01/05/2024 | 8 | 31 |

The calculation can also *reset* the accumulator. `IF(condition, acc+1, 0)` counts consecutive TRUE values and drops back to 0
on a FALSE, so it measures **streaks**. With this illustration input:

| Item | TRUE | TRUE | FALSE | TRUE | TRUE | TRUE | FALSE |
|---|:-:|:-:|:-:|:-:|:-:|:-:|:-:|
| `IF(item, acc+1, 0)` | 1 | 2 | 0 | 1 | 2 | 3 | 0 |

`MAX` of that output is the longest streak (3 here). For example, `=MAX(SCAN(0, tblCensus[MidnightCensus]>=tblCensus[StaffedBeds],
LAMBDA(run, full, IF(full, run+1, 0))))` returns **4**, the longest run of nights when every ICU bed was full. Streaks depend on
the previous row's result, so SUMPRODUCT and COUNTIFS can't calculate them without a helper column.

#### REDUCE

```
=REDUCE(initial_value, array, LAMBDA(accumulator, item, calculation))
```

REDUCE works like SCAN but returns only the final accumulator. It's a loop: "start here, then do this for each item." Count the
stays that ended in death or hospice by looping over a list of dispositions:

```
=REDUCE(0, {"Expired","Hospice"}, LAMBDA(total, dispo, total + COUNTIF(tblStays[DischargeDisposition], dispo)))
```

The result is **165**. `{"Expired","Hospice"}` is an array constant like the ones in section 3. Commas separate items across
a row, and semicolons separate items down a column, so `{"F01";"F02";"F03"}` (used by MAKEARRAY below) is a vertical list of
three facility IDs. REDUCE, MAP, and SCAN accept either shape.

#### MAKEARRAY

```
=MAKEARRAY(rows, columns, LAMBDA(r, c, calculation))
```

MAKEARRAY builds a grid by calling the LAMBDA for every row number and column number. This one counts discharges by facility
(rows) and year (columns):

```
=MAKEARRAY(3, 2, LAMBDA(r, c, COUNTIFS(tblStays[FacilityID], INDEX({"F01";"F02";"F03"}, r),
                                        tblStays[DischargeDateTime], ">="&DATE(2023+c, 1, 1),
                                        tblStays[DischargeDateTime], "<"&DATE(2024+c, 1, 1))))
```

| | 2024 | 2025 |
|---|:-:|:-:|
| F01 | 1,985 | 2,113 |
| F02 | 327 | 374 |
| F03 | 369 | 418 |

> ⚠️ **One value per call.** The LAMBDA inside MAP, BYROW, BYCOL, and MAKEARRAY must return a single value each time. If it
> returns an array (for example, a whole FILTER result per row), Excel shows `#CALC!`, because it can't nest arrays.

> ⚠️ **Answer cells hold one value.** MAP, BYROW, SCAN, and MAKEARRAY spill arrays. On the Practice sheet, wrap them in `SUM`,
> `MAX`, or `INDEX` so the result is one number. Use the Sandbox sheet to watch the full spill first.

> 💡 **Tip:** Every helper task in the practice also has a boolean-math or helper-column solution. The helpers earn
> their place when the per-row logic reads naturally with AND, OR, or MAX, when one row depends on the previous one (SCAN), or
> when you want to call your own named LAMBDA on each item.

### 8. Auditing complex formulas

A long formula that returns a believable number can still be wrong. These tools show you what each part calculates.

**Evaluate Formula** steps through a formula one calculation at a time:

1. Select the cell, then choose **Formulas → Evaluate Formula** (in the Formula Auditing group).
2. The part Excel will calculate next is underlined. Click **Evaluate** to replace it with its result.
3. Keep clicking. When a part refers to another formula cell, **Step In** shows that cell's formula and **Step Out** returns.
4. Stop at any step whose result can't be right, such as an error, a number that looks wrong, or a comparison between different
   data types. That's the step to fix. **Restart** begins again.

> ⚠️ **You can only read the start of a long array.** On a Table column, Evaluate Formula shows thousands of values, and the
> first ones may not be typical. The Stays sheet starts in January 2024, so a correct test for 2025 shows FALSE on every value
> you can see. Look at the data types instead. Text appears in quotes, such as `"F02"`, and numbers and dates appear without
> quotes. To check a whole condition, count its TRUE values (see *Count each condition on its own* below).

> 📋 Evaluate Formula is in every Windows version of Excel, but Excel for Mac doesn't have it. On a Mac, copy pieces of the
> formula into Sandbox cells and let them spill, or select a piece in the formula bar and press **Fn + F9** (then **Esc**).
> In Microsoft 365, selecting a piece of a formula while you edit it also shows its value in a small tooltip.

| Tool | Where | What it tells you |
|---|---|---|
| **F2** (Mac: **Fn + F2** or **⌃ + U**) | Edit the cell | Each reference gets a color, and a matching colored box outlines the cells on the sheet. A box that's one row too low is easy to spot |
| **F9** (Mac: **Fn + F9**) on a selected part | Formula bar | The value of just that part, if the result is short enough to display. Press **Esc** afterwards |
| **Trace Precedents / Trace Dependents** | Formulas → Formula Auditing | Arrows to the cells a formula uses, or to the formulas that use this cell. **Remove Arrows** clears them |
| **Ctrl + [** (Mac: **⌃ + [**) | Keyboard | Selects the cells the formula refers to, even on another sheet |
| **Watch Window** | Formulas → Watch Window | Keeps chosen cells' values in view while you work elsewhere (on a Mac: Microsoft 365, or Excel 2021 and later) |
| **Show Formulas**, **Ctrl + `** (Mac: **⌃ + `**) | Keyboard | Shows formulas instead of results in every cell |

**Debugging a LET.** Temporarily replace the last argument with one of the names. If the average looks wrong, change
`SUM(los*keep)/SUM(keep)` to `SUM(keep)` and press Enter. For the heart failure example in section 5 you should see 36. If you see 0,
one of the conditions in `keep` is never TRUE. Put the calculation back when you're done.

**Count each condition on its own.** When a count with several conditions looks wrong, type each condition into its own
`SUM(--(…))` on the Sandbox sheet. For the heart failure example in section 5:

| Formula | Result |
|---|:-:|
| `=SUM(--(tblStays[FacilityID]="F02"))` | 701 |
| `=SUM(--(tblStays[PrimaryDxCode]="I50.9"))` | 447 |
| `=SUM(--(YEAR(tblStays[DischargeDateTime])=2025))` | 2,905 |
| All three conditions multiplied together | 36 |

A condition that returns 0 is never TRUE, so that's the one to fix. The combined count can't be larger than the smallest single
count, so a combined count above 447 here would also point to a mistake.

A checklist for any long formula:

1. Read it aloud from the inside out. Each condition should match a phrase in the request.
2. Check that every array covers the same rows (F2 shows the boxes).
3. Check data types: numbers compared with numbers, text with text, dates with dates.
4. Test it on a case you can verify by hand, such as a single patient or a single day.
5. Compare the total with a simpler method, such as a COUNTIFS or a filtered column's status-bar count.

### 9. Performance and compatibility

| Function | Excel 2016/2019 | Excel 2021 | Excel 2024 | Microsoft 365 & web |
|---|:-:|:-:|:-:|:-:|
| SUMPRODUCT, FREQUENCY, CHOOSE, INDEX, COUNTIFS | ✅ | ✅ | ✅ | ✅ |
| XLOOKUP, FILTER, LET | ❌ | ✅ | ✅ | ✅ |
| TAKE, DROP, VSTACK | ❌ | ❌ | ✅ | ✅ |
| LAMBDA, MAP, BYROW, BYCOL, SCAN, REDUCE, MAKEARRAY, ISOMITTED | ❌ | ❌ | ✅ | ✅ |

- **Use Tables, not whole columns.** `tblStays[FacilityID]` covers exactly the data rows.
- **Name repeated work with LET.** Excel calculates each name once, however many times the formula uses it.
- **Avoid volatile functions** (OFFSET, INDIRECT, TODAY, NOW, RAND) in large models.
- **Choose between a helper column and one formula.** A helper column such as Readmit30 is easy to inspect and filter. A single
  LET formula is self-contained and can't be broken by someone deleting a column. Both are good choices for different audiences.
- **Calculation mode.** If a big workbook feels slow while you build it, switch to **Formulas → Calculation Options → Manual**,
  press **F9** (Mac: **Fn + F9**) to recalculate, and switch back to Automatic when you're done.

## 🧪 Hands-on practice

Download [`4.2-advanced-formulas-let-lambda.xlsx`](4.2-advanced-formulas-let-lambda.xlsx) and open the **Practice** sheet. Type
each formula in the yellow cell, and the **Check** column turns green when you're right. Tasks 8–12 and bonus B3 need Microsoft 365
or Excel 2024. Tasks 3, 4, 6, 7, and the bonus use XLOOKUP, FILTER, or LET, which need Excel 2021 or later.

<!-- BEGIN GENERATED: practice -->
Every task works on the Excel Tables in this workbook. Table references such as tblStays[FacilityID] are the easiest to read, and A1 ranges such as Stays!$C$2:$C$5587 work too. LOS days = discharge date − admit date (midnights), with a minimum of 1 day. Each answer cell must return ONE value, so try array formulas on the Sandbox sheet first.

| # | Task | Hint |
|:-:|------|------|
| 1 | How many inpatient stays were admitted on a Saturday or Sunday in 2025? Use AdmitDateTime for both the weekday and the year. | WEEKDAY(date, 2) numbers Monday as 1 and Sunday as 7. Multiply two TRUE/FALSE arrays inside SUMPRODUCT. |
| 2 | Case management reviews every 2025 discharge that had LOS days of 7 or more OR went to a Skilled Nursing Facility. What were the total charges (TotalCharges) of the stays on that review list? Count each stay once, and enter dollars and cents. | Adding two conditions gives OR logic, but a stay that meets both scores 2. Wrap the OR in (…>0). |
| 3 | The Labs sheet is sorted from oldest to newest CollectedDateTime. What was patient PT11882's most recent creatinine (TestCode CREAT) result? | XLOOKUP can look for 1 in an array of 1s and 0s that you build from two conditions. Its 6th argument chooses the search direction. |
| 4 | Use FREQUENCY with the bins {1,2,3,5,7,14} on the LOS days of every stay discharged in 2025. FREQUENCY returns 7 counts. Enter the 7th: the number of 2025 discharges with LOS days greater than 14. | FILTER the LOS values to 2025 first, then pick one count out of FREQUENCY's result with INDEX. |
| 5 | On the Stays sheet, fill the yellow Readmit30 column: 1 if the same patient has another stay whose AdmitDateTime is after this stay's DischargeDateTime and whose admit DATE is 0–30 days after this stay's discharge DATE, otherwise 0. The gray cell totals your column. How many stays were followed by a 30-day readmission? | COUNTIFS over the whole table, using this row's PatientID and DischargeDateTime as criteria. Compare dates with INT(). |
| 6 | Write ONE LET formula that names the admit and discharge columns, computes LOS days (minimum 1), and returns the average LOS days for sepsis stays (PrimaryDxCode A41.9) at Bluestone Memorial (F01) discharged in 2025. Round to 2 decimal places with ROUND. | Name each step (nights, then los, then a 1/0 keep array). A conditional average is SUM(los*keep)/SUM(keep). |
| 7 | The LOS index (observed ÷ expected) compares actual LOS days with the benchmark ExpectedLOS of each stay's diagnosis (Diagnoses sheet). Calculate it for Cedar Ridge (F03) stays discharged in 2025: total LOS days ÷ total ExpectedLOS. Keep full precision (the check accepts 2 decimal places). | Inside LET, XLOOKUP the whole PrimaryDxCode column. You get one ExpectedLOS for every stay. |
| 8 | Create a named function LOSDAYS(admit, discharge) in the Name Manager that returns LOS days (discharge date − admit date, minimum 1). Then call it on whole columns. What is the total of LOS days for Ashby Falls (F02) stays discharged in 2025? | Formulas → Name Manager → New. Test the LAMBDA in a cell first, for example =LAMBDA(…)(Stays!E2,Stays!F2). |
| 9 | Utilization review audits inpatient stays that spanned fewer than 2 midnights (discharge date − admit date < 2). Using MAP with a LAMBDA that uses AND, count Bluestone Memorial (F01) stays discharged in 2025 that spanned fewer than 2 midnights. | MAP can take several same-size arrays and passes one value from each to your LAMBDA. Wrap the result in SUM. |
| 10 | On the Labs sheet, count the results outside their reference range (ResultValue below RefLow or above RefHigh). Use BYROW over the three adjacent columns ResultValue:RefHigh with a LAMBDA that uses OR. | Inside BYROW, the LAMBDA gets one row of 3 cells: INDEX(r,1) is the value, INDEX(r,2) the low, INDEX(r,3) the high. |
| 11 | The Census sheet holds the Bluestone Memorial ICU's daily census for 2024–2025. A day is 'strained' when MidnightCensus ÷ StaffedBeds is 90% or more. Using SCAN, find the longest run of consecutive strained days. | Keep a running count that adds 1 on a strained day and resets to 0 otherwise. Then take the MAX. |
| 12 | The four HRRP conditions in the data are AMI (I21.4), heart failure (I50.9), pneumonia (J18.9), and COPD (J44.1). Using REDUCE to loop over the array {"I21.4","I50.9","J18.9","J44.1"}, count Bluestone Memorial (F01) stays discharged in 2025 with any of these primary diagnoses. | REDUCE(0, codes, LAMBDA(total, dxcode, total + …)) starts at 0 and adds one COUNTIFS per code. |
| 13 | The Audit sheet has a colleague's formula that should count Ashby Falls (F02) stays discharged in 2025 to Home Health, but it returns 0. Step through it with Evaluate Formula (Windows; on a Mac, test its pieces on the Sandbox sheet), fix the bug, and enter the correct count. | Watch what YEAR returns and what it is compared with. |
<!-- END GENERATED: practice -->

## ✅ Answer key

The workbook has a hidden **Answer Key** sheet (right-click any sheet tab → **Unhide…** → *Answer Key*). Its *Live result*
column runs each sample formula, except the LAMBDA tasks, which show "—" so the file also opens cleanly in older versions. The
same answers are below, collapsed so you don't see them by accident.

<!-- BEGIN GENERATED: answers -->
<details>
<summary><b>🔑 Show the answer key</b> — Try every task before opening this.</summary>

**1. Weekend admissions in 2025**

- **Answer:** 668
- **Solution:** `=SUMPRODUCT((WEEKDAY(tblStays[AdmitDateTime],2)>5)*(YEAR(tblStays[AdmitDateTime])=2025))`

Each comparison returns an array of 5,586 TRUE/FALSE values. Multiplying the two arrays turns TRUE into 1 and FALSE into 0, so a row is 1 only when both conditions hold (AND logic). SUMPRODUCT adds the 1s. COUNTIFS could handle the year with two date criteria, but it can't apply WEEKDAY to the range first, so the weekend test needs either a helper column or array logic.

**2. Charges on the case-management review list (OR logic)**

- **Answer:** 40,025,644.41
- **Solution:**

```
=SUMPRODUCT((((INT(tblStays[DischargeDateTime])-INT(tblStays[AdmitDateTime])>=7)+(tblStays[DischargeDisposition]="Skilled Nursing Facility"))>0)*(YEAR(tblStays[DischargeDateTime])=2025),tblStays[TotalCharges])
```


(A)+(B) is 0, 1, or 2. The 49 long SNF stays score 2, so multiplying by charges without the `>0` test counts their charges twice and gives $42,880,891.91. `((A)+(B))>0` turns the sum back into TRUE/FALSE, so every stay counts once. LOS days of 7 or more only depends on the midnight count, so `INT(dis)-INT(adm)>=7` is enough here (the 1-day minimum only changes 0-night stays).

**3. Most recent creatinine for PT11882 (multi-criteria XLOOKUP)**

- **Answer:** 3.90
- **Solution:**

```
=XLOOKUP(1,(tblLabs[PatientID]="PT11882")*(tblLabs[TestCode]="CREAT"),tblLabs[ResultValue],"Not found",0,-1)
```


Multiplying the two conditions builds an array that is 1 only on this patient's creatinine rows. XLOOKUP finds the value 1 in that array. search_mode -1 searches from the bottom up, and because the sheet is sorted oldest to newest, the bottom match is the latest result. A top-down search returns the first result (2.75), which is how creatinine looked months earlier. Rising creatinine can signal kidney injury.

**4. FREQUENCY of LOS days: stays longer than 14 days**

- **Answer:** 33
- **Solution:**

```
=INDEX(FREQUENCY(FILTER(INT(tblStays[DischargeDateTime])-INT(tblStays[AdmitDateTime]),YEAR(tblStays[DischargeDateTime])=2025),{1,2,3,5,7,14}),7)
```


FREQUENCY counts values ≤ 1, then >1 to 2, >2 to 3, >3 to 5, >5 to 7, >7 to 14, and finally everything above the last bin. That extra overflow count is why 6 bins return 7 numbers. INDEX(…, 7) returns just the overflow count, so the answer cell holds one value. The solution skips the 1-day minimum because a 0-night stay lands in the ≤ 1 bin either way, so applying it gives the same counts. `FREQUENCY(IF(YEAR(dis)=2025, INT(dis)-INT(adm)), bins)` also works, because FREQUENCY ignores the FALSE values that IF returns for other years.

**5. Readmit30 flag column (stays followed by a readmission)**

- **Answer:** 800
- **Solution:**

```
=IF(COUNTIFS(tblStays[PatientID],[@PatientID],tblStays[AdmitDateTime],">"&[@DischargeDateTime],tblStays[AdmitDateTime],"<"&(INT([@DischargeDateTime])+31))>0,1,0)
```


For each row, COUNTIFS counts the stays that have the same PatientID, an AdmitDateTime after this discharge, and an AdmitDateTime earlier than INT(discharge)+31, which is midnight at the start of the 31st day after the discharge date. That last test is the same as an admit DATE no more than 30 days later. Any count above 0 means a readmission. In A1 style the first row is `=IF(COUNTIFS($B$2:$B$5587,B2,$E$2:$E$5587,">"&F2,$E$2:$E$5587,"<"&(INT(F2)+31))>0,1,0)`. Your total should be 800, the same as the official Readmit30 = Y count in the full dataset. A window measured to the minute (admit ≤ discharge + 30) finds only 793 because it misses readmissions on day 30 that happen later in the day than the discharge did. The key's live formula does all rows in one cell by giving COUNTIFS whole columns as criteria.

**6. Average LOS days for F01 sepsis stays (LET)**

- **Answer:** 7.19
- **Solution:**

```
=LET(
  adm, tblStays[AdmitDateTime],
  dis, tblStays[DischargeDateTime],
  nights, INT(dis)-INT(adm),
  los, IF(nights<1, 1, nights),
  keep, (tblStays[FacilityID]="F01")*(tblStays[PrimaryDxCode]="A41.9")*(YEAR(dis)=2025),
  ROUND(SUM(los*keep)/SUM(keep), 2)
)
```


Each name is calculated once and reused, so INT(dis)-INT(adm) isn't repeated. `los*keep` zeroes out the rows you don't want, and SUM(keep) counts the rows you do (167 stays), which gives a conditional average without a helper column. `AVERAGE(FILTER(los, keep))` is an equally good last step.

**7. LOS index (O/E) for Cedar Ridge in 2025 (LET + XLOOKUP)**

- **Answer:** 1.25
- **Solution:**

```
=LET(
  adm, tblStays[AdmitDateTime],
  dis, tblStays[DischargeDateTime],
  nights, INT(dis)-INT(adm),
  los, IF(nights<1, 1, nights),
  explos, XLOOKUP(tblStays[PrimaryDxCode], tblDx[DxCode], tblDx[ExpectedLOS]),
  keep, (tblStays[FacilityID]="F03")*(YEAR(dis)=2025),
  SUM(los*keep)/SUM(explos*keep)
)
```


Given a whole column as its lookup value, XLOOKUP returns one benchmark per stay. Naming that array `explos` keeps the final line readable: observed days ÷ expected days for the kept rows. An index above 1.00 means patients stayed longer than the benchmark. Don't average the per-stay ratios. A ratio of totals weights long and short stays correctly.

**8. LOSDAYS named LAMBDA: total LOS days for F02 in 2025**

- **Answer:** 1,694
- **Solution:**

```
=SUM(LOSDAYS(tblStays[AdmitDateTime],tblStays[DischargeDateTime])*(tblStays[FacilityID]="F02")*(YEAR(tblStays[DischargeDateTime])=2025))
```


Define LOSDAYS as `=LAMBDA(admit, discharge, LET(nights, INT(discharge)-INT(admit), IF(nights<1, 1, nights)))`. Called with two whole columns, it returns 5,586 LOS values, and the two conditions keep the 374 F02 stays from 2025. Use IF for the minimum, not MAX. `MAX(1, INT(discharge)-INT(admit))` collapses the whole column into one number (the longest stay, 34 days), and the total becomes 12,716.

**9. Short stays at F01 in 2025 (MAP)**

- **Answer:** 118
- **Solution:**

```
=SUM(MAP(tblStays[FacilityID], tblStays[AdmitDateTime], tblStays[DischargeDateTime],
  LAMBDA(fac, adm, dis, IF(AND(fac="F01", YEAR(dis)=2025, INT(dis)-INT(adm)<2), 1, 0))))
```


MAP calls the LAMBDA once per stay, so `fac`, `adm`, and `dis` are single values and AND works. Outside MAP, AND(...) on whole columns would return one TRUE/FALSE for all 5,586 rows. The boolean-math version `=SUMPRODUCT((tblStays[FacilityID]="F01")*(YEAR(tblStays[DischargeDateTime])=2025)*(INT(tblStays[DischargeDateTime])-INT(tblStays[AdmitDateTime])<2))` gives the same count. MAP is worth it when the per-row logic reads better with AND, OR, or MAX.

**10. Lab results outside the reference range (BYROW)**

- **Answer:** 365
- **Solution:**

```
=SUM(BYROW(tblLabs[[ResultValue]:[RefHigh]],
  LAMBDA(r, IF(OR(INDEX(r,1)<INDEX(r,2), INDEX(r,1)>INDEX(r,3)), 1, 0))))
```


BYROW hands the LAMBDA one 1×3 row at a time, so OR tests just that row and returns one TRUE/FALSE. BYROW stacks the results into a column of 1s and 0s for SUM. `tblLabs[[ResultValue]:[RefHigh]]` is the structured reference for the three adjacent columns. Cross-check: the lab's own abnormal flags (not included here) mark the same 365 results.

**11. Longest run of strained ICU days (SCAN)**

- **Answer:** 15
- **Solution:**

```
=MAX(SCAN(0, tblCensus[MidnightCensus]/tblCensus[StaffedBeds]>=0.9,
  LAMBDA(run, strained, IF(strained, run+1, 0))))
```


SCAN walks down the 731 TRUE/FALSE values and keeps a running number (the accumulator `run`). On a strained day it adds 1, and on any other day it resets to 0. The result is one running count per day, so its MAX is the longest streak. A streak depends on the previous row's result, which SUMPRODUCT and COUNTIFS can't express without a helper column.

**12. HRRP-condition stays at F01 in 2025 (REDUCE)**

- **Answer:** 561
- **Solution:**

```
=REDUCE(0, {"I21.4","I50.9","J18.9","J44.1"},
  LAMBDA(total, dxcode, total + COUNTIFS(tblStays[PrimaryDxCode], dxcode, tblStays[FacilityID], "F01",
    tblStays[DischargeDateTime], ">="&DATE(2025,1,1), tblStays[DischargeDateTime], "<"&DATE(2026,1,1))))
```


REDUCE starts the accumulator `total` at 0, then calls the LAMBDA once per code, adding that code's COUNTIFS each time. It returns only the final total. The codes can't overlap (a stay has one primary diagnosis), so adding is safe. `SUM(COUNTIFS(…, {codes}, …))` gives the same answer. REDUCE earns its keep when each step needs more logic than one function call.

**13. Audit: fix the colleague's formula**

- **Answer:** 48
- **Solution:**

```
=SUMPRODUCT((tblStays[FacilityID]="F02")*(YEAR(tblStays[DischargeDateTime])=2025)*(tblStays[DischargeDisposition]="Home Health"))
```


YEAR returns the number 2025, but the formula compares it with the text "2025". In Excel a number never equals text, so that array is all FALSE and every product is 0. Evaluate Formula shows the YEAR step as numbers and the comparison as all FALSE (on a Mac, `=YEAR(tblStays[DischargeDateTime])="2025"` on the Sandbox sheet spills a column of FALSE). COUNTIFS would have accepted "2025" because it converts criteria strings, and that habit is how this bug usually gets in.

</details>
<!-- END GENERATED: answers -->

## 🏆 Bonus challenge

<!-- BEGIN GENERATED: bonus -->
Bluestone's quality committee wants 30-day all-cause readmission rates by hospital, built with CMS-style rules (simplified: CMS also risk-adjusts its rates and ignores planned readmissions). Index stays are inpatient stays discharged from 01/01/2025 through 11/30/2025, so every stay has a full 30 days of follow-up in the data. Exclude stays that ended in death (Expired), a transfer (Transfer to Another Hospital), or a discharge against medical advice (Left AMA). A readmission is any later inpatient admission of the same patient, at any Bluestone hospital, with an admit date 0–30 days after the index discharge date (the Task 5 rule). Rate = index stays followed by a readmission ÷ index stays. Build each answer as one LET formula over tblStays.

Work on the **Bonus** sheet of the workbook.

- **B1.** How many index stays are there system-wide? *(Hint: ISNA(MATCH(disposition, {list}, 0)) is TRUE when a disposition is NOT on the list.)*
- **B2.** Write one LET formula that returns the readmission rate for Bluestone Memorial (F01). Enter it as a percentage (1 decimal place is enough). *(Hint: Name readmit (the Task 5 test for every row at once), eligible, and keep, then divide SUM(keep*readmit) by SUM(keep).)*
- **B3.** Which hospital has the highest rate? Enter its FacilityID. *(Hint: Wrap your B2 formula in LAMBDA(facility, …), save it as READMITRATE, and MAP it over {"F01";"F02";"F03"}. Or edit the facility ID three times.)*
- **B4.** What is that hospital's rate? Enter it as a percentage (1 decimal place is enough). *(Hint: Reuse B2 with the other facility ID, or call your READMITRATE function.)*
- **B5.** CMS also publishes readmission rates by condition. What is the system-wide rate for index stays whose primary diagnosis is one of the four HRRP conditions from Task 12? Enter it as a percentage (1 decimal place is enough). *(Hint: Swap the facility test for a diagnosis-list test: ISNUMBER(MATCH(dx, {list}, 0)).)*
<!-- END GENERATED: bonus -->

<!-- BEGIN GENERATED: bonus-answers -->
<details>
<summary><b>🔑 Show the bonus solution</b> — Give it a real try first!</summary>

**B1. Eligible index stays (system-wide)**

- **Answer:** 2,452
- **Solution:**

```
=LET(
  dis, tblStays[DischargeDateTime],
  eligible, (dis>=DATE(2025,1,1))*(dis<DATE(2025,12,1))
            *ISNA(MATCH(tblStays[DischargeDisposition], {"Expired","Transfer to Another Hospital","Left AMA"}, 0)),
  SUM(eligible)
)
```


MATCH returns #N/A for every disposition that isn't one of the three exclusions, so ISNA turns "not excluded" into TRUE. The two date tests keep discharges from 01/01/2025 up to (not including) 12/01/2025. Without the exclusions there would be 2,595 stays in the window.

**B2. Readmission rate for F01**

- **Answer:** 15.8%
- **Solution:**

```
=LET(
  adm, tblStays[AdmitDateTime],
  dis, tblStays[DischargeDateTime],
  readmit, COUNTIFS(tblStays[PatientID], tblStays[PatientID], adm, ">"&dis, adm, "<"&(INT(dis)+31))>0,
  eligible, (dis>=DATE(2025,1,1))*(dis<DATE(2025,12,1))
            *ISNA(MATCH(tblStays[DischargeDisposition], {"Expired","Transfer to Another Hospital","Left AMA"}, 0)),
  keep, eligible*(tblStays[FacilityID]="F01"),
  SUM(keep*readmit)/SUM(keep)
)
```


`readmit` hands COUNTIFS whole columns as criteria, so it returns one count per stay (5,586 of them), and `>0` turns those counts into TRUE/FALSE. LET calculates that expensive array once. `keep` narrows the index stays to F01 (1,788 stays). If you finished Task 5, `SUM(keep*tblStays[Readmit30])/SUM(keep)` is a shorter last step that reuses your column. The self-contained version keeps working even if someone deletes that column.

**B3. Hospital with the highest readmission rate**

- **Answer:** F02
- **Solution:**

```
=LET(
  facs, {"F01";"F02";"F03"},
  rates, MAP(facs, LAMBDA(f, READMITRATE(f))),
  XLOOKUP(MAX(rates), rates, facs)
)
```


Define READMITRATE in the Name Manager as your B2 formula wrapped in `=LAMBDA(facility, LET(…))`, with "F01" replaced by `facility`. MAP passes each ID in turn to `LAMBDA(f, READMITRATE(f))`, and XLOOKUP returns the ID next to the largest rate. The rates are F01 15.8%, F02 18.6%, F03 16.5%. Ashby Falls Community Hospital has the highest rate even though it's the smallest hospital, which is why rates, not counts, are compared.

**B4. Highest hospital readmission rate**

- **Answer:** 18.6%
- **Solution:**

```
=LET(
  adm, tblStays[AdmitDateTime],
  dis, tblStays[DischargeDateTime],
  readmit, COUNTIFS(tblStays[PatientID], tblStays[PatientID], adm, ">"&dis, adm, "<"&(INT(dis)+31))>0,
  eligible, (dis>=DATE(2025,1,1))*(dis<DATE(2025,12,1))
            *ISNA(MATCH(tblStays[DischargeDisposition], {"Expired","Transfer to Another Hospital","Left AMA"}, 0)),
  keep, eligible*(tblStays[FacilityID]="F02"),
  SUM(keep*readmit)/SUM(keep)
)
```


Same formula as B2 with "F02". With the named LAMBDA it is simply `=READMITRATE("F02")`. Turning a long LET into a LAMBDA means one tested definition serves every hospital, so a fix to the rule only has to be made once.

**B5. System-wide readmission rate for HRRP conditions**

- **Answer:** 20.7%
- **Solution:**

```
=LET(
  adm, tblStays[AdmitDateTime],
  dis, tblStays[DischargeDateTime],
  readmit, COUNTIFS(tblStays[PatientID], tblStays[PatientID], adm, ">"&dis, adm, "<"&(INT(dis)+31))>0,
  eligible, (dis>=DATE(2025,1,1))*(dis<DATE(2025,12,1))
            *ISNA(MATCH(tblStays[DischargeDisposition], {"Expired","Transfer to Another Hospital","Left AMA"}, 0)),
  keep, eligible*ISNUMBER(MATCH(tblStays[PrimaryDxCode], {"I21.4","I50.9","J18.9","J44.1"}, 0)),
  SUM(keep*readmit)/SUM(keep)
)
```


ISNUMBER(MATCH(…)) is the mirror image of the exclusion test: TRUE when the diagnosis IS on the list. The 682 HRRP index stays are readmitted more often than index stays overall (16.3%), which is why CMS targets these conditions.

</details>
<!-- END GENERATED: bonus-answers -->

## Key takeaways

- Comparisons on whole columns return TRUE/FALSE arrays. Multiply them for AND, add them for OR, and wrap an OR in `(…)>0` so no
  row counts twice.
- Use SUMPRODUCT, or SUM in Microsoft 365, when the conditions need calculations such as YEAR, WEEKDAY, or LOS. Keep COUNTIFS and
  SUMIFS for simple conditions.
- `XLOOKUP(1, (A)*(B), …)` matches on several columns, and search_mode -1 returns the latest match on a sorted sheet.
- LET names each step once, which makes long formulas readable, faster, and easy to debug.
- LAMBDA turns a tested LET into your own function. Test it in a cell, save it in the Name Manager, and use IF rather than MAX
  or AND so it works on whole columns.
- MAP, BYROW, SCAN, and REDUCE loop for you. SCAN handles running totals and streaks that depend on the previous row.
- Audit before you trust: Evaluate Formula, F9 on a selection, colored range boxes, and a hand-checked test case.

<!-- BEGIN GENERATED: nav -->
---

⬅️ **Previous:** [4.1 Dynamic Arrays: FILTER, SORT, UNIQUE & More](../01-dynamic-arrays/README.md) · 🏠 [Course home](../../README.md) · **Next:** [4.3 Power Query: Import, Transform & Combine](../03-power-query/README.md) ➡️
<!-- END GENERATED: nav -->
