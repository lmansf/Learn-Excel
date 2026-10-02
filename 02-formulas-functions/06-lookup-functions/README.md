# Lesson 2.6 · Lookups: VLOOKUP, INDEX/MATCH & XLOOKUP

> **Level:** Beginner → Intermediate · **Time:** about 60 minutes · **Workbook:** [`2.6-lookup-functions.xlsx`](2.6-lookup-functions.xlsx)
> **Data:** Every 2025 encounter (529) for the 185 patients on Dr. Isabella Nguyen's internal medicine panel, stored as IDs only, plus the lookup tables that turn those IDs into names: the panel roster, providers, ICD-10 diagnoses, payers, departments, and claims. Also a 2025 department budget grid, BMI and age-band tier tables, and the practice's December 2025 referral list.

Open any extract from an electronic health record and you'll see rows like `ENC110776 · PT12169 · PRV1137 · J06.9 · PY05`.
Databases store IDs because IDs never change and never repeat, but nobody can read a report made of them. The patient's name
lives in one table, the diagnosis description in another, and the payer in a third. A **lookup** takes an ID, finds it in the
table that describes it, and brings back the detail you need. Analysts use lookups every day: they put payer names on claims
reports, diagnosis descriptions on quality reviews, attending names on case lists, and risk tiers on patient rosters. In this
lesson you'll use the three lookup tools every Excel user meets (VLOOKUP, INDEX/MATCH, and XLOOKUP) on a primary care panel's
encounters, and you'll learn to handle the IDs that aren't there.

## What you'll learn

- Look up exact matches with VLOOKUP and XLOOKUP
- Use approximate matches for tiers like age bands and BMI categories
- Combine INDEX and MATCH for flexible and two-way lookups
- Handle missing values with IFNA and XLOOKUP's if_not_found

## 📖 Guide

The examples use the lesson workbook, so open it and try each formula as you read. Dr. Isabella Nguyen is an internist at the
Bluestone Outpatient Pavilion. Her panel is the 185 patients who list her as their primary care provider, and the Encounters sheet
holds every 2025 visit or stay those patients had anywhere in the Bluestone Health System.

| Sheet | Rows | Columns |
|---|---|---|
| **Encounters** | 2–530 | A EncounterID, B PatientID, C EncounterType, D AdmitDate, E DischargeDate, F DeptID, G AttendingProviderID, H PrimaryDxCode, I PayerID, J TotalCharges, K PayerName (yellow, task 5) |
| **Patients** | 2–186 | A PatientID, B MRN, C FirstName, D LastName, E Sex, F DOB, G Age, H HeightIn, I WeightLb, J BMI, K PrimaryPayerID, L BMICategory (yellow, task 7) |
| **Referrals** | 2–25 | A ReferralID, B ReceivedDate, C MRN, D PatientName, E ReferredBy, F Reason, G PatientID (yellow, task 13) |
| **Providers** | 2–148 | A ProviderID, B FirstName, C LastName, D Credential, E Specialty, F PrimaryDeptID, G FacilityID |
| **Diagnoses** | 2–52 | A DxCode, B DxDescription, C DxCategory, D ChronicCondition, E ExpectedLOS |
| **Payers** | 2–9 | A PayerID, B PayerName, C PayerType, D AvgAllowedPctOfCharges, E AvgDaysToPay, F TimelyFilingDays |
| **Departments** | 2–32 | A DeptID, B DeptName, C FacilityID, D ServiceLine, E UnitType, F StaffedBeds, G CostCenter |
| **Claims** | 2–530 | A ClaimID, B EncounterID, C PayerID, D SubmitDate, E BilledAmount, F PaidAmount, G ClaimStatus, H DenialReason |
| **Budget** | 2–32 | A DeptID, B DeptName, C FacilityID, D–J the 2025 annual budget for seven expense categories (Salaries & Wages … Other Operating) |
| **BMITiers** | 2–7 | A MinBMI, B Category, C BMIRange |
| **AgeBands** | 2–7 | A MinAge, B AgeBand |

Encounters is sorted by AdmitDate, oldest first. **MRN** (medical record number) is stored as 8-character text, so leading zeros
are kept. **Age** is the patient's age in completed years on 12/31/2025, and **BMI** (body mass index) is 703 × WeightLb ÷
HeightIn², rounded to 1 decimal place.

> 💡 **Tip:** Try each example as you read. When a formula refers to a cell such as `B3` without a sheet name, type it on the sheet
> the example names (usually Encounters or Patients), in an empty column a few columns to the right of the data, such as column N.
> Don't use the column right next to the data, because a formula typed there becomes a new Table column.

### 1. What a lookup does

Here is row 3 of the Encounters sheet, and what each ID means once you look it up:

| Encounters column | Value in row 3 | Look it up in | What comes back |
|---|---|---|---|
| B PatientID | PT12169 | Patients | Christine Miller |
| F DeptID | D400 | Departments | Primary Care Clinic |
| G AttendingProviderID | PRV1137 | Providers | Isabella Nguyen, MD (Internal Medicine) |
| H PrimaryDxCode | J06.9 | Diagnoses | Acute upper respiratory infection, unspecified |
| I PayerID | PY05 | Payers | Evergreen Mutual Insurance |

Every lookup has the same four parts. Learn these terms, because the rest of the lesson uses them:

- The **lookup value** is the thing you already have, such as the PatientID PT12169.
- The **lookup table** is the sheet that describes it, such as Patients.
- The **key column** is the column of the lookup table that holds the IDs, such as Patients column A. Each ID should appear in it
  exactly once.
- The **return column** holds the detail you want back, such as LastName.

So every lookup formula answers three questions: *what am I looking for, where do I search for it, and what do I bring back?*
The functions in this lesson differ only in how you tell them.

> 📋 Database people call this a **join**. A lookup is a join done one cell at a time, which is why lookups show up in nearly
> every healthcare report.

### 2. VLOOKUP: the classic exact match

```
=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
```

| Argument | What it means | Example |
|---|---|---|
| lookup_value | The ID you're looking for | `B3` (PT12169) |
| table_array | The whole lookup table. Its **first column** must be the key column | `Patients!$A$2:$K$186` |
| col_index_num | Which column of table_array to return, counting the key column as 1 | `3` for FirstName |
| range_lookup | `FALSE` (or `0`) for an exact match. `TRUE` (or omitted) for an approximate match | `FALSE` |

On the Encounters sheet, this formula returns the first name of row 3's patient:

```
=VLOOKUP(B3,Patients!$A$2:$K$186,3,FALSE)      → Christine
```

Here's what Excel does:

1. It searches the first column of `Patients!$A$2:$K$186` (PatientID) from top to bottom for PT12169.
2. It stops at the first exact match, which is row 105 of the Patients sheet.
3. It counts across that row to the 3rd column of the table: PatientID (1), MRN (2), FirstName (3).
4. It returns that cell, Christine.

The V stands for *vertical*, because VLOOKUP searches down a column.

> ⚠️ **Always type FALSE for IDs.** If you leave range_lookup out, VLOOKUP does an approximate match.
> `=VLOOKUP("PT12170",Patients!$A$2:$K$186,4)` returns **Miller**, even though PT12170 isn't on the Patients sheet. PT12170 is a
> Bluestone patient who sees a different primary care doctor, so she isn't on Dr. Nguyen's panel. Excel quietly falls back to the
> nearest smaller ID, PT12169 (Christine Miller), and returns her last name, so a report would show the wrong patient. With FALSE,
> the same formula returns #N/A, which is the honest answer.

**What VLOOKUP can't do well:**

- **It can't look left.** The key column must be the first column of table_array. The table doesn't have to start at column A,
  though. To find a last name from an MRN, start the table at the MRN column:
  `=VLOOKUP("03449108",Patients!$B$2:$D$186,3,FALSE)` returns Cruz (B is column 1, C is 2, D is 3). But no VLOOKUP can return the
  PatientID for an MRN, because PatientID sits to the left of MRN.
- **col_index_num is a typed number.** If someone inserts a column into the Patients sheet, `3` now points at a different column,
  and every formula returns the wrong field without any error.
- **A col_index_num larger than the table** returns #REF!. `=VLOOKUP(B3,Patients!$A$2:$K$186,15,FALSE)` fails because the table
  has only 11 columns.
- **It returns the first match only.** If an ID appears twice, you get the first row and never hear about the second.

> 💡 **Tip:** The lookup table usually sits on another sheet, and you don't have to type its sheet name. While you're typing the
> formula, click the **Patients** tab and drag over the table. Excel writes `Patients!A2:K186` for you. Press **F4** (Mac:
> **⌘ + T**) right away to lock it as `Patients!$A$2:$K$186`, so the table stays put when you copy the formula down. If a sheet
> name contains a space, Excel wraps it in single quotes, as in `'Patient List'!$A$2:$K$186`. Lesson 1.5 explains absolute and
> cross-sheet references.

> 📋 **HLOOKUP** is VLOOKUP turned on its side. It searches the **top row** of a table and returns a value from a row further
> down. On the Budget sheet, `=HLOOKUP("Medical Supplies",Budget!$D$1:$J$32,2,FALSE)` finds the Medical Supplies header and returns
> the 2nd row of the range. That's the Medical Supplies budget for D100, Bluestone Memorial's Emergency Department: 1,391,039.
> You'll rarely need HLOOKUP, because XLOOKUP and INDEX/MATCH search rows and columns equally well.

### 3. XLOOKUP: the modern lookup

```
=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])
```

| Argument | What it means |
|---|---|
| lookup_value | The value you're looking for |
| lookup_array | The **one column** (or row) to search |
| return_array | The column (or columns) to return from. It must have the same number of rows as lookup_array |
| if_not_found | Optional. What to show when there's no match, such as `"Not found"`. Leave it out to get #N/A |
| match_mode | Optional. How to match (table below). The default, 0, is an exact match |
| search_mode | Optional. Which direction to search (table below). The default, 1, is first to last |

| match_mode | Meaning | Use it for |
|:-:|---|---|
| `0` (default) | Exact match | IDs, codes, names |
| `-1` | Exact match, or else the next **smaller** item | Tiers such as BMI categories (section 6) |
| `1` | Exact match, or else the next **larger** item | "Round up to the next tier" tables |
| `2` | Wildcard match with `*`, `?`, and `~` | "Description contains knee" (section 7) |

| search_mode | Meaning |
|:-:|---|
| `1` (default) | Search from the first row to the last |
| `-1` | Search from the last row to the first, so you get the **last** match (section 7) |
| `2` / `-2` | Binary search on data sorted ascending / descending. Fast on huge lists, but wrong answers if the data isn't sorted |

On the Encounters sheet, row 3's diagnosis description takes one short formula, with nothing to count:

```
=XLOOKUP(H3,Diagnoses!$A$2:$A$52,Diagnoses!$B$2:$B$52)      → Acute upper respiratory infection, unspecified
```

To skip an optional argument and use a later one, leave its place empty between commas. `=XLOOKUP(A2,B2:B9,C2:C9,,2)` has no
if_not_found and a match_mode of 2.

Why XLOOKUP is easier than VLOOKUP:

- It matches exactly unless you ask for something else, so you can't forget a FALSE.
- The search column and the return column are separate arguments, so it **looks left** as easily as right.
- Inserting a column between them doesn't break it, because there's no column number.
- It has if_not_found built in, and it can search from the bottom up.

**Returning several columns at once.** If return_array is several columns wide, XLOOKUP returns the whole matching row. Typed in
an empty cell, `=XLOOKUP("PT12169",Patients!$A$2:$A$186,Patients!$C$2:$D$186)` shows Christine in that cell and Miller in the
cell to its right. That's called a **spill**, and Lesson 4.1 covers it in depth. A spill needs empty cells to land in, and if
something is in the way you'll see #SPILL!. A multi-column result is also useful inside another function. To total a department's
whole budget row, give XLOOKUP all seven category columns and wrap it in SUM:

```
=SUM(XLOOKUP("D110",Budget!$A$2:$A$32,Budget!$D$2:$J$32))      → 12,523,079
```

That's the full 2025 expense budget for Medical-Surgical 4 West.

> ⚠️ lookup_array and return_array must cover the same rows. `Patients!$A$2:$A$186` with `Patients!$D$2:$D$185` returns #VALUE!.
> Selecting whole matching ranges, or using Table columns, avoids this.

> 📋 **Version note:** XLOOKUP and XMATCH need Microsoft 365, Excel 2021, or Excel 2024 (Windows or Mac), or Excel for the web.
> Excel 2019 and earlier don't have them. If someone opens your file in an older version, the formula shows `_xlfn.XLOOKUP` and
> returns #NAME? when it recalculates. If your colleagues use older Excel, use INDEX/MATCH (next section), which works everywhere.

### 4. INDEX and MATCH

INDEX and MATCH are two separate functions that make a flexible lookup when you combine them.

**INDEX** returns the item at a position you give it:

```
=INDEX(array, row_num, [column_num])
=INDEX(Providers!$C$2:$C$148,137)      → Nguyen
```

Position 137 of the LastName range is Nguyen. Positions count from the top of the range, not from the top of the sheet, so
position 137 is sheet row 138.

**MATCH** does the reverse. It returns the position of a value in a range:

```
=MATCH(lookup_value, lookup_array, [match_type])
=MATCH("PRV1137",Providers!$A$2:$A$148,0)      → 137
```

| match_type | Finds | The lookup_array must be |
|:-:|---|---|
| `0` | The first exact match | In any order |
| `1` (the default) | The largest value that is less than or equal to lookup_value | Sorted ascending |
| `-1` | The smallest value that is greater than or equal to lookup_value | Sorted descending |

> ⚠️ MATCH's default is **1**, an approximate match. Always type the `0` for IDs, just like VLOOKUP's FALSE.

**Put them together.** MATCH finds the position, and INDEX returns the value at that position of another column:

```
=INDEX(Providers!$C$2:$C$148,MATCH("PRV1137",Providers!$A$2:$A$148,0))      → Nguyen
```

Read it from the inside out. MATCH answers "where is PRV1137 in the ProviderID list?" (position 137), and INDEX answers "what's at
position 137 of the LastName list?"
Because the search column and the return column are separate, INDEX/MATCH can look left. This finds the PatientID for an MRN,
which VLOOKUP can't do:

```
=INDEX(Patients!$A$2:$A$186,MATCH("03449108",Patients!$B$2:$B$186,0))      → PT10073
```

> ⚠️ The INDEX range and the MATCH range must start on the same row and be the same size. If one starts at row 2 and the other
> at row 1, every answer comes from the wrong row, and nothing warns you.

**XMATCH** is the modern MATCH. It matches exactly by default and takes the same match_mode and search_mode options as XLOOKUP:
`=XMATCH("PRV1137",Providers!$A$2:$A$148)` also returns 137. Like XLOOKUP, it needs Microsoft 365 or Excel 2021 or later.

**MATCH as a yes/no test.** `=ISNUMBER(MATCH("03449108",Patients!$B$2:$B$186,0))` returns TRUE when the MRN is on the panel and
FALSE when it isn't, because MATCH returns a number when it finds the value and #N/A when it doesn't.

**Which one should you use?**

| | VLOOKUP | INDEX/MATCH | XLOOKUP |
|---|---|---|---|
| Default match | Approximate (you must type FALSE) | Approximate (you must type 0) | Exact |
| Look left | No | Yes | Yes |
| Survives inserted columns | No (typed column number) | Yes | Yes |
| Built-in "not found" message | No, wrap it in IFNA | No, wrap it in IFNA | Yes, if_not_found |
| Last match (search from the bottom) | No | Not simply | Yes, search_mode -1 |
| Return several columns at once | Not simply | Whole row with column_num 0 | Yes |
| Works in | Every version | Every version | Microsoft 365, Excel 2021+ |

Use XLOOKUP when everyone who opens the file has a current version. Use INDEX/MATCH when the file has to work in older Excel.
Learn VLOOKUP anyway, because you'll find it in almost every workbook you inherit.

### 5. Two-way lookups

The Budget sheet is a **grid**: departments down the side, expense categories across the top. To pull one number you need two
matches, one for the row and one for the column. INDEX takes both:

```
=INDEX(Budget!$D$2:$J$32,
       MATCH("D110",Budget!$A$2:$A$32,0),
       MATCH("Medical Supplies",Budget!$D$1:$J$1,0))      → 1,155,058
```

1. `MATCH("D110",Budget!$A$2:$A$32,0)` returns 2, because D110 is the 2nd department.
2. `MATCH("Medical Supplies",Budget!$D$1:$J$1,0)` returns 3, because Medical Supplies is the 3rd category header.
3. `INDEX(Budget!$D$2:$J$32,2,3)` returns the cell where row 2 and column 3 of the grid cross, which is F3: 4 West's 2025
   Medical Supplies budget.

You can type the formula on one line. The line breaks above only make it easier to read. In Excel, **Alt + Enter** (Mac:
**⌃ + ⌥ + Return**) adds a line break inside the formula bar.

> ⚠️ The pieces have to line up. The header range (`D1:J1`) must start in the same column as the grid (`D2:J32`), and the row-label
> range (`A2:A32`) must start on the same row as the grid. If they're offset by one, you get the neighbor's number.

XLOOKUP does the same job with one XLOOKUP inside another:

```
=XLOOKUP("D110",Budget!$A$2:$A$32,XLOOKUP("Medical Supplies",Budget!$D$1:$J$1,Budget!$D$2:$J$32))
```

The inner XLOOKUP returns the entire Medical Supplies column, and the outer one picks the D110 row out of it.

**Whole rows.** A column_num of 0 tells INDEX to return the whole row, so `=SUM(INDEX(Budget!$D$2:$J$32,2,0))` adds up all seven
categories for D110. Section 3 showed the XLOOKUP version.

> ⚠️ **Look up by ID, not by name.** Three hospitals each have a department called Intensive Care Unit (D130, D230, and D330).
> `=XLOOKUP("Intensive Care Unit",Budget!$B$2:$B$32,Budget!$A$2:$A$32)` returns D130 and silently ignores the other two. Patient
> names repeat even more often. Four patients on this panel have the last name Miller, and a lookup on "Miller" returns the first
> one, PT10413. That's why every lookup in this lesson uses an ID.

> 💡 **Tip:** Put the row label and the column header in their own cells, such as `D110` in L2 and `Medical Supplies` in L3 of the
> Budget sheet, and refer to those cells in the formula. Now you can answer a different question by typing over L2 or L3. Lesson 3.2 shows how to
> turn those cells into drop-down lists.

### 6. Approximate match: tiers and bands

Many healthcare measures are reported in **tiers**: BMI categories, age bands, risk scores, payment brackets. You could write a
long nested IF, but a **tier table** is easier to read and to change. A tier table lists the **lower bound** of each tier, sorted
from smallest to largest:

| MinBMI | Category | BMIRange |
|:-:|---|---|
| 0.0 | Underweight | Below 18.5 |
| 18.5 | Healthy weight | 18.5 to under 25 |
| 25.0 | Overweight | 25 to under 30 |
| 30.0 | Obesity class 1 | 30 to under 35 |
| 35.0 | Obesity class 2 | 35 to under 40 |
| 40.0 | Obesity class 3 | 40 and above |

An **approximate match** finds the largest lower bound that is less than or equal to the value. Patients row 4 has a BMI of
27.1. The bounds at or below 27.1 are 0, 18.5, and 25.0. The largest of those is 25.0, so the category is Overweight. Typed on
the Patients sheet, each of these formulas returns Overweight for row 4:

```
=VLOOKUP(J4,BMITiers!$A$2:$B$7,2,TRUE)
=XLOOKUP(J4,BMITiers!$A$2:$A$7,BMITiers!$B$2:$B$7,,-1)
=INDEX(BMITiers!$B$2:$B$7,MATCH(J4,BMITiers!$A$2:$A$7,1))
```

The same idea works for any banded measure. Patients row 3 is 73 years old, and on the Patients sheet
`=VLOOKUP(G3,AgeBands!$A$2:$B$7,2,TRUE)` returns the 65-74 band.

**Boundaries belong to the tier that starts there.** Patients row 65 has a BMI of exactly 25.0. The largest bound at or below
25.0 is 25.0 itself, so that patient is Overweight, not Healthy weight. That matches the clinical definition ("25 to under 30"),
and it's why a tier table stores lower bounds.

> ⚠️ **Sort the tier table ascending** for VLOOKUP with TRUE and MATCH with 1. They use a fast search that assumes sorted data. On
> an unsorted table they return wrong tiers without any error. XLOOKUP with match_mode -1 checks every row, so it works on unsorted
> tables too, but keep tier tables sorted anyway so people can read them.

> ⚠️ **Start the table at the lowest possible value.** A value below the first bound has no tier and returns #N/A. That's why the
> BMI table starts at 0, not at 18.5.

> ⚠️ **Mind the decimals.** A BMI of 24.96 displays as 25.0 when the cell shows one decimal place, but the stored value is still
> 24.96, so the lookup puts it in Healthy weight. When a cutoff matters, ROUND the value first (Lesson 1.4). The BMI column in this
> workbook is already rounded to 1 decimal place.

> ⚠️ **Tiers need an approximate match, and IDs need an exact one.** An exact-match lookup on a tier table fails:
> `=XLOOKUP(J4,BMITiers!$A$2:$A$7,BMITiers!$B$2:$B$7)` returns #N/A, because 27.1 isn't one of the bounds. The reverse mistake, an
> approximate match on IDs, is worse, because it returns a wrong row instead of an error (section 2).

Compare the tier lookup with the IFS approach from Lesson 2.1:

```
=IFS(J4>=40,"Obesity class 3",J4>=35,"Obesity class 2",J4>=30,"Obesity class 1",
     J4>=25,"Overweight",J4>=18.5,"Healthy weight",TRUE,"Underweight")
```

Both return Overweight. If the cutoffs change, though, you'd have to edit every IFS formula in the workbook, while the lookup
version needs one change in one cell of the tier table.

### 7. Wildcards and searching from the bottom

**Wildcards** let you match part of a text value:

| Wildcard | Matches | Example | Matches… |
|:-:|---|---|---|
| `*` | Any number of characters, including none | `"*sepsis*"` | "Sepsis, unspecified organism" |
| `?` | Exactly one character | `"E11.?"` | E11.9, but not E11.65 (two characters after the dot) |
| `~` | Treats the next `*` or `?` as a plain character | `"~*"` | A real asterisk |

XLOOKUP uses wildcards only when match_mode is 2:

```
=XLOOKUP("*sepsis*",Diagnoses!$B$2:$B$52,Diagnoses!$A$2:$A$52,,2)      → A41.9
```

MATCH with match_type 0, and VLOOKUP with FALSE, accept wildcards automatically. `=INDEX(Diagnoses!$A$2:$A$52,
MATCH("*sepsis*",Diagnoses!$B$2:$B$52,0))` also returns A41.9. Lookups ignore upper and lower case, with or without wildcards,
so "SEPSIS" finds the same row.

> ⚠️ A wildcard lookup still returns only the **first** match. `"Type 2*"` matches two descriptions,
> E11.65 (with hyperglycemia) and E11.9 (without complications), and XLOOKUP returns E11.65 because it comes first. Make the
> pattern specific enough to match one row, and check with `=COUNTIF(Diagnoses!$B$2:$B$52,"Type 2*")`, which returns 2 here.

**The last match.** A normal lookup stops at the first match. Patient PT10444 has nine encounters, and the Encounters sheet is
sorted oldest to newest, so the first match is the oldest:

```
=XLOOKUP("PT10444",Encounters!$B$2:$B$530,Encounters!$D$2:$D$530)            → 01/01/2025
```

search_mode -1 starts at the bottom, so the first match it meets is the most recent encounter:

```
=XLOOKUP("PT10444",Encounters!$B$2:$B$530,Encounters!$D$2:$D$530,,0,-1)      → 12/15/2025
=XLOOKUP("PT10444",Encounters!$B$2:$B$530,Encounters!$C$2:$C$530,,0,-1)      → Emergency
```

"Last row" means "most recent" only because the sheet is sorted by date. Sort first if you aren't sure. If the result shows a
number such as 46006, the cell needs a date format (**Home → Number Format → Short Date**). MAXIFS from Lesson 2.5 would also find
the latest date, but only a lookup can bring back other columns from that row, such as the EncounterType.

> 📋 XLOOKUP returns one match. To list *every* encounter for a patient, you'll use the FILTER function in Lesson 4.1.

### 8. When a lookup fails: #N/A, IFNA, and if_not_found

**#N/A** means "no match." Sometimes that's the right answer: a referred patient really isn't on the panel. Often, though, the
value is there and the lookup still can't see it. These are the usual causes:

| Cause | Example | Fix |
|---|---|---|
| The value really isn't in the table | A referral for a patient who isn't on the panel | Expected. Show a clear message instead of #N/A |
| A typo or a look-alike character | `ENC1O0776` typed with the letter O instead of a zero | Fix the source. `=LEN()` and `=EXACT()` help you compare |
| Extra spaces | `"PT12169 "` with a trailing space | `TRIM` the lookup value (Lesson 2.2) |
| Number vs. text | The number 3449108 vs. the text MRN "03449108" | Convert one side so both are text (or both are numbers) |
| The ranges miss the row | The lookup range stops at row 150 but the ID is on row 160 | Select the full columns, or use Table columns |
| A value below the first tier | A BMI table that starts at 18.5 | Start the tier table at 0 |

**Numbers vs. text** causes more failed lookups than anything else. MRNs, ZIP codes, and some account numbers are stored as text
so their leading zeros survive. When another system sends the same MRN as a number, the zeros disappear and the two values no
longer match:

```
=XLOOKUP(3449108,Patients!$B$2:$B$186,Patients!$A$2:$A$186)                    → #N/A
=XLOOKUP(TEXT(3449108,"00000000"),Patients!$B$2:$B$186,Patients!$A$2:$A$186)   → PT10073
```

`TEXT(value,"00000000")` turns the number back into 8-character text with the leading zeros restored, so the match works. Going
the other way, when the IDs in the table are real numbers and your lookup value is text, convert it with `--A2` or `VALUE(A2)`.

> 💡 **Tip:** To see whether a cell holds text or a number, look at its alignment. Text sits on the left of the cell and numbers sit
> on the right, unless someone changed the alignment. `=ISTEXT(B2)` tells you for certain.

**Showing a message instead of #N/A.** When "not there" is a real possibility, replace the error with words. XLOOKUP has its own
argument for this, and the older functions use IFNA:

```
=XLOOKUP("04683061",Patients!$B$2:$B$186,Patients!$A$2:$A$186,"Not found")                → Not found
=IFNA(VLOOKUP("PT12170",Patients!$A$2:$K$186,4,FALSE),"Not found")                       → Not found
=IFNA(INDEX(Patients!$A$2:$A$186,MATCH("04683061",Patients!$B$2:$B$186,0)),"Not found")  → Not found
```

Why IFNA and not IFERROR? Lesson 2.1 showed that IFERROR hides every error type. A lookup that fails because of a broken formula
should *look* broken:

| Formula (on Encounters row 3) | Result | What happened |
|---|---|---|
| `=VLOOKUP(B3,Patients!$A$2:$K$186,15,FALSE)` | #REF! | Column 15 doesn't exist. That's a bug |
| `=IFERROR(VLOOKUP(B3,Patients!$A$2:$K$186,15,FALSE),"Not found")` | Not found | The bug is disguised as a missing patient |
| `=IFNA(VLOOKUP(B3,Patients!$A$2:$K$186,15,FALSE),"Not found")` | #REF! | IFNA traps only #N/A, so the bug stays visible |

XLOOKUP's if_not_found behaves like IFNA. It replaces only the "no match" result.

> ⚠️ Choose a fallback that can't be mistaken for data. `""` looks like a blank field, and `0` looks like a real amount (a $0
> charge, or a patient aged 0). Text such as `Not found` is unmistakable, and `COUNTIF(range,"Not found")` counts the misses.

> 📋 **Version note:** IFNA needs Excel 2013 or later (Excel 2016 or later on a Mac). IFERROR works in Excel 2007 and later.

### 9. Chaining lookups

Sometimes one lookup isn't enough. If you have only an EncounterID and you want the patient's last name, you need two hops:
EncounterID to PatientID (on Encounters), then PatientID to LastName (on Patients). You can put one lookup inside the other:

```
=XLOOKUP(XLOOKUP("ENC110776",Encounters!$A$2:$A$530,Encounters!$B$2:$B$530),
         Patients!$A$2:$A$186,Patients!$D$2:$D$186)      → Miller
```

Read it from the inside out. The inner XLOOKUP turns ENC110776 into PT12169, and the outer XLOOKUP turns PT12169 into Miller.

Nesting gets hard to read quickly. A **helper cell** is usually clearer: put the first hop (the PatientID) in its own cell, then
point every later lookup at that cell. You can check each hop on its own, and you look up the PatientID only once instead of once
per formula.

When the first hop might fail, the "not found" result travels down the chain. If the inner XLOOKUP returns `"Not found"`, the outer
XLOOKUP searches the PatientID column for the text "Not found", doesn't find it, and returns its own if_not_found. Formulas that
*calculate* with a looked-up value, like DATEDIF on a date of birth, need an IF test first, because a calculation on the text "Not
found" returns #VALUE!.

You can also join lookups with `&` to build a label, as you did with text in Lesson 2.2. For row 3's attending provider, on the
Encounters sheet:

```
=XLOOKUP(G3,Providers!$A$2:$A$148,Providers!$B$2:$B$148)&" "&
 XLOOKUP(G3,Providers!$A$2:$A$148,Providers!$C$2:$C$148)&", "&
 XLOOKUP(G3,Providers!$A$2:$A$148,Providers!$D$2:$D$148)      → Isabella Nguyen, MD
```

### 10. Filling a lookup column and keeping lookups reliable

Most of the time you'll add a lookup as a new column, such as a payer name next to every PayerID. Here's the procedure:

1. Click the first empty cell of the new column (for example, K2 on Encounters).
2. Type the lookup for that row, such as `=XLOOKUP(I2,Payers!A2:A9,Payers!B2:B9)`.
3. Click into each lookup range and press **F4** (Mac: **⌘ + T**) to make it absolute: `Payers!$A$2:$A$9`. Leave `I2` relative so it
   moves down with each row.
4. Press **Enter**. On an Excel Table, Excel usually fills the rest of the column for you. If it doesn't, select the cell and
   double-click the fill handle, or select the column's cells and press **Ctrl + D** (Mac: **⌘ + D**).
5. Check the result. Filter the column for #N/A or "Not found", or count them with COUNTIF, before you trust it.

> 📋 **Excel Tables:** if you click cells instead of typing addresses, Excel may write
> `=XLOOKUP([@PayerID],tblPayers[PayerID],tblPayers[PayerName])`. Those are **structured references** to the Table and its columns.
> They never slide when you copy, and they grow when the table grows. Lesson 3.1 covers them.

**Habits that keep lookups reliable:**

- **Make sure keys are unique.** `=COUNTIF(Patients!$A$2:$A$186,A2)` should be 1 for every ID. A duplicate key means a lookup
  silently returns the first one.
- **Use exact matches for IDs.** Type FALSE, 0, or leave XLOOKUP's match_mode at its default.
- **Don't hard-code column numbers you might break.** VLOOKUP can find its own column number with MATCH:
  `=VLOOKUP(B3,Patients!$A$1:$K$186,MATCH("LastName",Patients!$A$1:$K$1,0),FALSE)` keeps returning the last name even if someone
  inserts a column. (The table starts at row 1 here so that its headers line up with the MATCH range. The header row never matches
  a real ID.)
- **Keep ranges the same size and on the same rows.** If the lookup range and the return range are different sizes, XLOOKUP
  returns #VALUE!. If they're the same size but start on different rows, INDEX/MATCH returns values from the wrong rows.
- **Check the type of your keys.** Text "123" never matches the number 123.
- **Keep lookup tables in the same workbook when you can.** A lookup into another file creates an **external link**. It shows the
  last saved values when the other file is closed, and it breaks if that file is moved or renamed. To see and fix a workbook's links,
  use **Data → Edit Links** (called **Workbook Links** in newer Microsoft 365 versions).
- **On very large sheets, think about speed.** An exact match checks rows one by one. A few thousand lookups over a few thousand
  rows is instant. Hundreds of thousands can slow a workbook down. Then you can sort the key column and use XLOOKUP's binary search
  (search_mode 2), or let one MATCH in a helper column feed several INDEX formulas.

### 11. Worked example: turning an ID-only row into a readable line

*Question: what happened in encounter ENC110776, in words a care manager can read?*

1. Find the encounter. Press **Ctrl + F** (Mac: **⌘ + F**) on the Encounters sheet and search for ENC110776. It's row 3. Type
   the formulas below in empty cells of column N on the Encounters sheet.
2. Patient: `=XLOOKUP(B3,Patients!$A$2:$A$186,Patients!$C$2:$C$186)&" "&XLOOKUP(B3,Patients!$A$2:$A$186,Patients!$D$2:$D$186)`
   returns Christine Miller.
3. Diagnosis: `=XLOOKUP(H3,Diagnoses!$A$2:$A$52,Diagnoses!$B$2:$B$52)` returns Acute upper respiratory infection, unspecified.
4. Where: `=XLOOKUP(F3,Departments!$A$2:$A$32,Departments!$B$2:$B$32)` returns Primary Care Clinic.
5. Who: the joined provider label from section 9 returns Isabella Nguyen, MD, and
   `=XLOOKUP(G3,Providers!$A$2:$A$148,Providers!$E$2:$E$148)` returns Internal Medicine.
6. Payer: `=XLOOKUP(I3,Payers!$A$2:$A$9,Payers!$B$2:$B$9)` returns Evergreen Mutual Insurance.
7. Claim: the Claims sheet has its own ClaimID key, but column B holds the EncounterID, and XLOOKUP can search any column:
   `=XLOOKUP(A3,Claims!$B$2:$B$530,Claims!$G$2:$G$530)` returns Paid.

The readable version: *Christine Miller saw Dr. Nguyen in the Primary Care Clinic on 01/02/2025 for an upper respiratory infection.
Evergreen Mutual Insurance paid the claim.*

Then sanity-check it. The visit was an Outpatient encounter in Dr. Nguyen's own clinic, so an internist as the attending and a
small charge ($304.68) both make sense. If a lookup had returned a cardiac surgeon or a $90,000 charge, you'd recheck the ranges.

### 12. Quick reference

| Function | Syntax | Returns | Available in |
|---|---|---|---|
| VLOOKUP | `VLOOKUP(value, table, col_num, FALSE)` | A value from a column to the right of the key column | Every version |
| HLOOKUP | `HLOOKUP(value, table, row_num, FALSE)` | A value from a row below the key row | Every version |
| MATCH | `MATCH(value, range, 0)` | The position of the value in the range | Every version |
| INDEX | `INDEX(range, row_num, [col_num])` | The value at a position (col_num 0 = whole row) | Every version |
| INDEX/MATCH | `INDEX(return_range, MATCH(value, lookup_range, 0))` | A lookup in any direction | Every version |
| XLOOKUP | `XLOOKUP(value, lookup_range, return_range, [if_not_found], [match_mode], [search_mode])` | A lookup in any direction, one or many columns | Microsoft 365, Excel 2021+ |
| XMATCH | `XMATCH(value, range, [match_mode], [search_mode])` | The position, exact by default | Microsoft 365, Excel 2021+ |
| IFNA | `IFNA(value, value_if_na)` | The value, or the fallback for #N/A only | Excel 2013+ (Mac 2016+) |

| Shortcut | Windows | Mac |
|---|---|---|
| Toggle `$` in a reference | **F4** | **⌘ + T** |
| Edit the active cell (shows each range in color) | **F2** | **⌃ + U** |
| Fill the selection down | **Ctrl + D** | **⌘ + D** |
| Find an ID on a sheet | **Ctrl + F** | **⌘ + F** |
| Show formulas instead of results | **Ctrl + `** | **⌃ + `** |
| Line break inside a formula | **Alt + Enter** | **⌃ + ⌥ + Return** |

## 🧪 Hands-on practice

Download [`2.6-lookup-functions.xlsx`](2.6-lookup-functions.xlsx) and open the **Practice** sheet. Type each answer as a formula in
the yellow cell. The **Check** column turns green when you're right. Tasks 5, 7, and 13 ask you to fill a column on a data sheet,
and their gray cells summarize your work.

<!-- BEGIN GENERATED: practice -->
The Encounters sheet stores IDs only. Every answer comes from looking an ID up in another sheet. Write each answer as a formula. Your answers sit on this Practice sheet, so a reference to a data cell needs its sheet name, such as Encounters!B9. Click the cell on the data sheet and Excel writes the sheet name for you. Lock lookup ranges with $ (F4, Mac: ⌘ + T) so they don't slide when you copy.

| # | Task | Hint |
|:-:|------|------|
| 1 | Encounter ENC110900 is on row 9 of the Encounters sheet. Use VLOOKUP with its PatientID (cell Encounters!B9) to return the patient's last name from the Patients sheet. | VLOOKUP(lookup_value, table_array, col_index_num, FALSE). Count columns from PatientID (column 1) to LastName |
| 2 | Encounter ENC111256 (Encounters row 20) was billed to payer PY02. Use VLOOKUP on the Payers sheet to return the payer's PayerType. | Count the columns of the Payers table to find PayerType's column number |
| 3 | Encounter ENC113996 (Encounters row 150) has primary diagnosis code E11.65. Use XLOOKUP to return its description from the Diagnoses sheet. | XLOOKUP(lookup_value, lookup_array, return_array) |
| 4 | Who attended encounter ENC112339 (Encounters row 73)? Use INDEX and MATCH with its AttendingProviderID to return the provider's Specialty from the Providers sheet. | INDEX(Specialty column, MATCH(id, ProviderID column, 0)) |
| 5 | Fill the yellow PayerName column on the Encounters sheet with a lookup that returns each encounter's payer name (start in K2 and copy down to row 530). The gray cell counts how many encounters your column shows as exactly "Medicare". | Lock the Payers ranges with $ before you copy the formula down |
| 6 | Patient PT10395 is on Patients row 21, and the Age column shows 75. Use VLOOKUP with an approximate match (TRUE) on the AgeBands sheet to return the patient's age band. | An approximate match returns the band whose MinAge is the largest one that is ≤ the age |
| 7 | Fill the yellow BMICategory column on the Patients sheet: look up each patient's BMI in the BMITiers table with an approximate match (start in L2 and copy down to row 186). The gray cell counts patients in any obesity class (BMI 30 or higher). | VLOOKUP(…, TRUE) or XLOOKUP with match_mode -1 (exact match or next smaller item) |
| 8 | The lab's interface file sent a result for MRN 50339. It arrived as a number, so its leading zeros were dropped, but the Patients sheet stores MRNs as 8-character text. Return the PatientID for this MRN. (PatientID is to the left of MRN, so VLOOKUP can't do it.) | Rebuild the text MRN with TEXT(number,"00000000"), then use INDEX/MATCH or XLOOKUP |
| 9 | A clinic note mentions "osteoarthritis of the right knee." Use a wildcard lookup to return the DxCode whose DxDescription contains the word knee. | An asterisk wildcard stands for any characters, so put one on each side of the word. XLOOKUP needs match_mode 2 |
| 10 | Two-way lookup: on the Budget sheet, what is the 2025 Pharmaceuticals budget for department D130 (Intensive Care Unit at Bluestone Memorial Hospital)? Find the row by DeptID and the column by category name. | INDEX(grid, MATCH(row label…), MATCH(column header…)) |
| 11 | What is the TOTAL 2025 expense budget (all seven categories) for department D400, Primary Care Clinic, where Dr. Nguyen practices? Use one lookup that returns the department's whole row of the grid, wrapped in SUM. | XLOOKUP's return_array can be several columns wide |
| 12 | Patient PT13163 visited more often than anyone else on the panel (11 encounters in 2025). The Encounters sheet is sorted oldest to newest. Use XLOOKUP searching from the bottom up to return the AdmitDate of this patient's most recent encounter. | search_mode is XLOOKUP's 6th argument, and -1 searches last-to-first |
| 13 | The Referrals sheet lists 24 referrals received in December. Fill its yellow PatientID column by looking up each MRN in the Patients sheet (start in G2). Patients who aren't on the panel yet must show the text Not found instead of #N/A. The gray cell counts the Not found rows. | XLOOKUP's 4th argument (if_not_found), or wrap the lookup in IFNA(…, "Not found") |
<!-- END GENERATED: practice -->

## ✅ Answer key

The workbook has a hidden **Answer Key** sheet (right-click any sheet tab → **Unhide…** → *Answer Key*). Its *Live result* column
runs each sample formula, so you can see it working. The same answers are below, collapsed so you don't see them by accident.

<!-- BEGIN GENERATED: answers -->
<details>
<summary><b>🔑 Show the answer key</b> — Try every task before opening this.</summary>

**1. Encounter ENC110900 is on row 9 of the Encounters sheet. Use VLOOKUP with its…**

- **Answer:** Robertson
- **Solution:** `=VLOOKUP(Encounters!B9,Patients!$A$2:$K$186,4,FALSE)`

VLOOKUP searches the first column of the table (PatientID) for the ID, then returns the value from the 4th column of the same row: PatientID, MRN, FirstName, **LastName**. FALSE asks for an exact match, which is what you want for any ID.

**2. Encounter ENC111256 (Encounters row 20) was billed to payer PY02. Use VLOOKUP on the…**

- **Answer:** Medicare Advantage
- **Solution:** `=VLOOKUP(Encounters!I20,Payers!$A$2:$F$9,3,FALSE)`

PayerType is the 3rd column of Payers (PayerID, PayerName, **PayerType**). The payer's name, Silverline Medicare Advantage, contains the word Medicare, but its type is Medicare Advantage, a private plan that contracts with Medicare. That difference matters for billing rules, which is why reports look up the type instead of guessing it from the name.

**3. Encounter ENC113996 (Encounters row 150) has primary diagnosis code E11.65. Use…**

- **Answer:** Type 2 diabetes mellitus with hyperglycemia
- **Solution:** `=XLOOKUP(Encounters!H150,Diagnoses!$A$2:$A$52,Diagnoses!$B$2:$B$52)`

XLOOKUP takes two separate columns, one to search and one to return, so there's no column number to count. It also defaults to an exact match, so you can't forget the FALSE that VLOOKUP needs.

**4. Who attended encounter ENC112339 (Encounters row 73)? Use INDEX and MATCH with its…**

- **Answer:** Pulmonary & Critical Care
- **Solution:** `=INDEX(Providers!$E$2:$E$148,MATCH(Encounters!G73,Providers!$A$2:$A$148,0))`

MATCH finds the position of the ProviderID in the ProviderID column (the 0 means exact match). INDEX then returns the item at that same position in the Specialty column. Because the two columns are separate arguments, INDEX/MATCH works in every Excel version and doesn't care where the columns sit.

**5. PayerName column (count of Medicare encounters)**

- **Answer:** 120
- **Solution:** `=XLOOKUP(I2,Payers!$A$2:$A$9,Payers!$B$2:$B$9)`

This is the everyday use of a lookup: adding a readable column to an ID-only extract. The $ signs keep the Payers ranges fixed while the lookup value (I2, then I3, …) moves down a row at a time. `=VLOOKUP(I2,Payers!$A$2:$F$9,2,FALSE)` works just as well. COUNTIF with "Medicare" counts exact matches only, so Silverline Medicare Advantage isn't included.

**6. Patient PT10395 is on Patients row 21, and the Age column shows 75. Use VLOOKUP with…**

- **Answer:** 75-84
- **Solution:** `=VLOOKUP(Patients!G21,AgeBands!$A$2:$B$7,2,TRUE)`

With TRUE, VLOOKUP looks for the largest MinAge that is less than or equal to 75. MinAge 75 qualifies, and the next one (85) is too big, so the answer is the 75-84 band. A value exactly on a boundary belongs to the band that **starts** there. That's why a tier table lists each band's lower bound, sorted smallest to largest.

**7. BMICategory column (patients in an obesity class)**

- **Answer:** 82
- **Solution:** `=VLOOKUP(J2,BMITiers!$A$2:$B$7,2,TRUE)`

Each BMI falls between two lower bounds in BMITiers, and the approximate match returns the tier whose lower bound is the largest one at or below the BMI. 5 patients sit exactly on a boundary (25.0 or 30.0) and correctly land in the higher tier. The XLOOKUP version is `=XLOOKUP(J2,BMITiers!$A$2:$A$7,BMITiers!$B$2:$B$7,,-1)`. The panel has 41 patients in class 1, 28 in class 2, and 13 in class 3.

**8. The lab's interface file sent a result for MRN 50339. It arrived as a number, so its…**

- **Answer:** PT11047
- **Solution:** `=XLOOKUP(TEXT(50339,"00000000"),Patients!$B$2:$B$186,Patients!$A$2:$A$186)`

Looking up the number 50339 returns #N/A, because a number never equals the text "00050339". TEXT(…,"00000000") pads it back to 8 characters as text, and then the match works. The INDEX/MATCH version is `=INDEX(Patients!$A$2:$A$186,MATCH(TEXT(50339,"00000000"),Patients!$B$2:$B$186,0))`. Both can return a column to the left of the one they search.

**9. A clinic note mentions "osteoarthritis of the right knee." Use a wildcard lookup to…**

- **Answer:** M17.11
- **Solution:** `=XLOOKUP("*knee*",Diagnoses!$B$2:$B$52,Diagnoses!$A$2:$A$52,,2)`

The asterisk stands for any number of characters, so `"*knee*"` matches "Unilateral primary osteoarthritis, right knee". XLOOKUP only uses wildcards when match_mode is 2. MATCH with match_type 0 uses them automatically: `=INDEX(Diagnoses!$A$2:$A$52,MATCH("*knee*",Diagnoses!$B$2:$B$52,0))`. If several descriptions matched, you'd get the first one.

**10. Two-way lookup: on the Budget sheet, what is the 2025 Pharmaceuticals budget for…**

- **Answer:** 993,601
- **Solution:**

```
=INDEX(Budget!$D$2:$J$32,MATCH("D130",Budget!$A$2:$A$32,0),MATCH("Pharmaceuticals",Budget!$D$1:$J$1,0))
```


The first MATCH finds the department's row inside the grid and the second finds the category's column in the header row. INDEX returns the cell where they cross. The nested XLOOKUP version is `=XLOOKUP("D130",Budget!$A$2:$A$32,XLOOKUP("Pharmaceuticals",Budget!$D$1:$J$1,Budget!$D$2:$J$32))`: the inner XLOOKUP returns the whole Pharmaceuticals column, and the outer one picks the department's row from it. Look up by DeptID, not DeptName. Three facilities each have an "Intensive Care Unit", and a lookup on that name would return only the first one.

**11. What is the TOTAL 2025 expense budget (all seven categories) for department D400,…**

- **Answer:** 5,348,882
- **Solution:** `=SUM(XLOOKUP("D400",Budget!$A$2:$A$32,Budget!$D$2:$J$32))`

When return_array is seven columns wide, XLOOKUP returns all seven values from the matching row, and SUM adds them. On its own in an empty cell the same XLOOKUP would spill across seven cells. INDEX can do this too, because a column number of 0 means "the whole row": `=SUM(INDEX(Budget!$D$2:$J$32,MATCH("D400",Budget!$A$2:$A$32,0),0))`.

**12. Patient PT13163 visited more often than anyone else on the panel (11 encounters in…**

- **Answer:** 12/29/2025
- **Solution:** `=XLOOKUP("PT13163",Encounters!$B$2:$B$530,Encounters!$D$2:$D$530,,0,-1)`

A normal lookup stops at the first match, which here is the oldest encounter (01/19/2025). search_mode -1 starts at the bottom, so the first match it meets is the most recent row. If the result shows a number like 46020, format the cell as a date. `=MAXIFS(…)` (Lesson 2.5) gives the same date, but only the lookup can return other columns from that row, like its EncounterID or diagnosis.

**13. Referral PatientID column (count of Not found)**

- **Answer:** 7
- **Solution:** `=XLOOKUP(C2,Patients!$B$2:$B$186,Patients!$A$2:$A$186,"Not found")`

XLOOKUP's if_not_found argument replaces #N/A with your own text, but only when the lookup really finds nothing. The older equivalent is `=IFNA(INDEX(Patients!$A$2:$A$186,MATCH(C2,Patients!$B$2:$B$186,0)),"Not found")`. Prefer IFNA to IFERROR here. IFERROR would also hide a #REF! or #NAME? caused by a broken formula, and you'd never know. The 7 Not found rows are new patients for the practice to register.

</details>
<!-- END GENERATED: answers -->

## 🏆 Bonus challenge

<!-- BEGIN GENERATED: bonus -->
Dr. Nguyen's care manager answers the same questions all day: who is this encounter's patient, how old were they, what were they treated for, who was the attending, who pays, and where does the claim stand? Build a reusable lookup card on the Card sheet. Each column takes one EncounterID in row 5 and returns ten facts below it. Rows 6 and 7 are helper cells (PatientID and Attending ProviderID) that the other rows can reuse. Write every formula in column B, then copy B6:B15 to C6:C15. Card A holds ENC117295. Card B holds ENC119O88, an ID copied from a handwritten note, and every row of Card B must show Not found instead of an error. Keep both IDs in place while you check your answers. The gray cells on the Bonus sheet read your card.

Work on the **Bonus** sheet of the workbook.

- **B1.** Card A (ENC117295): what does your Patient name row (Card!B8) show? Format: First Last. *(Hint: Chain two lookups: EncounterID → PatientID (row 6), then PatientID → names. Join with &" "&)*
- **B2.** Card A: how old was the patient, in completed years, on the encounter's AdmitDate (Card!B9)? *(Hint: Look up DOB and AdmitDate, then DATEDIF(…, …, "Y") from Lesson 2.3)*
- **B3.** Card A: what does your Attending row (Card!B12) show? Format: First Last, Credential (for example, Isabella Nguyen, MD). *(Hint: Use the Attending ProviderID helper in row 7, three lookups, and & to join them)*
- **B4.** Card A: what is the claim status (Card!B15)? Claims are matched by EncounterID, which is column B of the Claims sheet. *(Hint: XLOOKUP can search any column. VLOOKUP would need its table to start at column B)*
- **B5.** Card B (ENC119O88): how many of the ten output rows (Card!C6:C15) show exactly Not found? All ten should. If any shows #N/A or #VALUE!, fix that row's formula in column B and copy it right again. *(Hint: Give every lookup an if_not_found, and let IF skip calculations (like DATEDIF) when the helper cell says Not found)*
<!-- END GENERATED: bonus -->

<!-- BEGIN GENERATED: bonus-answers -->
<details>
<summary><b>🔑 Show the bonus solution</b> — Give it a real try first!</summary>

**B1. Card A · Patient name (First Last)**

- **Answer:** Carlos Ortiz
- **Solution:**

```
=IF(B6="Not found","Not found",XLOOKUP(B6,Patients!$A$2:$A$186,Patients!$C$2:$C$186)&" "&XLOOKUP(B6,Patients!$A$2:$A$186,Patients!$D$2:$D$186))
```


Row 6 does the first hop (EncounterID → PatientID) once, and every patient row reuses it. Chaining through a helper cell keeps each formula short and lets you check each hop on its own. The IF in front returns Not found when row 6 already says so, which keeps Card B clean.

**B2. Card A · Age at admission (years)**

- **Answer:** 78
- **Solution:**

```
=IF(B6="Not found","Not found",DATEDIF(XLOOKUP(B6,Patients!$A$2:$A$186,Patients!$F$2:$F$186),XLOOKUP(B$5,Encounters!$A$2:$A$530,Encounters!$D$2:$D$530),"Y"))
```


DATEDIF with "Y" counts completed years. The patient was born on 09/18/1946 and admitted on 08/01/2025, 48 days before turning 79, so the answer is 78. YEAR(admit) − YEAR(DOB) would give 79, which is one year too old. The DOB comes from Patients (via row 6) and the AdmitDate from Encounters (via row 5), so this one formula reads two different tables.

**B3. Card A · Attending (First Last, Credential)**

- **Answer:** Judith Cox, DO
- **Solution:**

```
=IF(B7="Not found","Not found",XLOOKUP(B7,Providers!$A$2:$A$148,Providers!$B$2:$B$148)&" "&XLOOKUP(B7,Providers!$A$2:$A$148,Providers!$C$2:$C$148)&", "&XLOOKUP(B7,Providers!$A$2:$A$148,Providers!$D$2:$D$148))
```


Three lookups on the same ProviderID return FirstName, LastName, and Credential, and & glues them together with a space and a comma. Showing the credential is safer than putting "Dr." in front of every name, because some attendings are nurse practitioners (NP) or physician assistants (PA). Judith Cox is a DO, a doctor of osteopathic medicine.

**B4. Card A · Claim status**

- **Answer:** Denied
- **Solution:** `=XLOOKUP(B$5,Claims!$B$2:$B$530,Claims!$G$2:$G$530,"Not found")`

The Claims key is ClaimID, but the card knows only the EncounterID. XLOOKUP searches the EncounterID column wherever it is. With VLOOKUP you'd start the table at column B: `=VLOOKUP(B$5,Claims!$B$2:$H$530,6,FALSE)`. This claim was denied (Authorization Required), so it goes on the care manager's follow-up list.

**B5. Card B · rows showing Not found**

- **Answer:** 10
- **Solution:** `=XLOOKUP(B$5,Encounters!$A$2:$A$530,Encounters!$B$2:$B$530,"Not found")`

ENC119O88 contains the letter O where the real ID ENC119088 has a zero. To Excel they're different text, so every lookup comes back empty-handed. That's the most common reason a lookup fails on data that looks right. The sample solution shown is row 6. Three techniques make the whole card fail politely:

1. Every XLOOKUP gets an if_not_found of "Not found".
2. Chained lookups pass "Not found" along. The inner XLOOKUP returns it, and the outer one can't find a DxCode called "Not found", so it returns its own "Not found".
3. Rows that join or calculate (the two name rows and the age row) test their helper cell first with `IF(B6="Not found","Not found",…)`. Without that test, DATEDIF of the text "Not found" returns #VALUE!. A name join would show #N/A, or "Not found Not found" if each of its lookups had its own if_not_found.

Your finished card should match this:

| Row | Field | Card A (ENC117295) | Card B (ENC119O88) |
|:-:|---|---|---|
| 6 | PatientID | PT12752 | Not found |
| 7 | Attending ProviderID | PRV1030 | Not found |
| 8 | Patient name (First Last) | Carlos Ortiz | Not found |
| 9 | Age at admission (years) | 78 | Not found |
| 10 | Primary diagnosis | Chronic obstructive pulmonary disease with (acute) exacerbation | Not found |
| 11 | Unit (department) | Medical-Surgical 4 West | Not found |
| 12 | Attending (First Last, Credential) | Judith Cox, DO | Not found |
| 13 | Attending specialty | Hospital Medicine | Not found |
| 14 | Payer | Silverline Medicare Advantage | Not found |
| 15 | Claim status | Denied | Not found |

Every formula on the card (type them in column B, then copy them to column C):

- Row 6 · PatientID: `=XLOOKUP(B$5,Encounters!$A$2:$A$530,Encounters!$B$2:$B$530,"Not found")`
- Row 7 · Attending ProviderID: `=XLOOKUP(B$5,Encounters!$A$2:$A$530,Encounters!$G$2:$G$530,"Not found")`
- Row 8 · Patient name (First Last): `=IF(B6="Not found","Not found",XLOOKUP(B6,Patients!$A$2:$A$186,Patients!$C$2:$C$186)&" "&XLOOKUP(B6,Patients!$A$2:$A$186,Patients!$D$2:$D$186))`
- Row 9 · Age at admission (years): `=IF(B6="Not found","Not found",DATEDIF(XLOOKUP(B6,Patients!$A$2:$A$186,Patients!$F$2:$F$186),XLOOKUP(B$5,Encounters!$A$2:$A$530,Encounters!$D$2:$D$530),"Y"))`
- Row 10 · Primary diagnosis: `=XLOOKUP(XLOOKUP(B$5,Encounters!$A$2:$A$530,Encounters!$H$2:$H$530,"Not found"),Diagnoses!$A$2:$A$52,Diagnoses!$B$2:$B$52,"Not found")`
- Row 11 · Unit (department): `=XLOOKUP(XLOOKUP(B$5,Encounters!$A$2:$A$530,Encounters!$F$2:$F$530,"Not found"),Departments!$A$2:$A$32,Departments!$B$2:$B$32,"Not found")`
- Row 12 · Attending (First Last, Credential): `=IF(B7="Not found","Not found",XLOOKUP(B7,Providers!$A$2:$A$148,Providers!$B$2:$B$148)&" "&XLOOKUP(B7,Providers!$A$2:$A$148,Providers!$C$2:$C$148)&", "&XLOOKUP(B7,Providers!$A$2:$A$148,Providers!$D$2:$D$148))`
- Row 13 · Attending specialty: `=XLOOKUP(B7,Providers!$A$2:$A$148,Providers!$E$2:$E$148,"Not found")`
- Row 14 · Payer: `=XLOOKUP(XLOOKUP(B$5,Encounters!$A$2:$A$530,Encounters!$I$2:$I$530,"Not found"),Payers!$A$2:$A$9,Payers!$B$2:$B$9,"Not found")`
- Row 15 · Claim status: `=XLOOKUP(B$5,Claims!$B$2:$B$530,Claims!$G$2:$G$530,"Not found")`

</details>
<!-- END GENERATED: bonus-answers -->

## Key takeaways

- A lookup joins two tables: it takes an ID, finds it in a key column, and returns a detail from the same row. Look up by IDs,
  never by names, because names repeat.
- Use exact matches for IDs. VLOOKUP needs FALSE and MATCH needs 0. XLOOKUP matches exactly by default.
- XLOOKUP looks in any direction, survives inserted columns, and can search from the bottom or return several columns. Use
  INDEX/MATCH when the file must work in Excel 2019 or earlier.
- Two-way lookups cross a row match with a column match: `INDEX(grid, MATCH(row), MATCH(column))`.
- Approximate matches turn a value into a tier. The tier table lists lower bounds sorted ascending and starts at the lowest possible
  value, and a value on a boundary belongs to the tier that starts there.
- #N/A means "no match." Check for typos, spaces, and numbers stored as text before you trap it, then use if_not_found or IFNA
  rather than IFERROR so real mistakes stay visible.

<!-- BEGIN GENERATED: nav -->
---

⬅️ **Previous:** [2.5 Conditional Counting & Summing](../05-conditional-aggregation/README.md) · 🏠 [Course home](../../README.md) · **Next:** [3.1 Excel Tables, Structured References & Named Ranges](../../03-data-analysis/01-tables-named-ranges/README.md) ➡️
<!-- END GENERATED: nav -->
