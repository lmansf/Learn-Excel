# Lesson 2.3 · Dates & Times

> **Level:** Beginner → Intermediate · **Time:** about 55 minutes · **Workbook:** [`2.3-date-time-functions.xlsx`](2.3-date-time-functions.xlsx)
> **Data:** All 369 inpatient stays admitted at Ashby Falls Community Hospital in 2025 (with each patient's date of birth and, for ED admissions, the ED arrival time) and the insurance claim for each stay; the hospital's 226 emergency department visits in Q4 2025; one pay week (December 14–20, 2025) of worked shifts on Medical-Surgical 4 West at Bluestone Memorial Hospital; and the business office's 2025–2026 holiday calendar.

Almost every number a hospital reports has a clock behind it. Length of stay drives bed planning and reimbursement. Door-to-provider
minutes show how well the emergency department keeps up. Payers set filing and appeal deadlines in days, and payroll needs the
paid hours of a night shift that started yesterday and ended this morning. Each of those answers comes from date and time math, and
it goes wrong in quiet ways: an age that is a year off near a birthday, a month count that misses the last evening of the month, or a
night shift that comes out as negative hours. In this lesson you'll learn how Excel stores dates and times, then use that knowledge
to calculate ages, lengths of stay, business-day deadlines, ED waits, and shift hours that hold up under review.

## What you'll learn

- Understand date serial numbers and times as fractions of a day
- Build and take apart dates with DATE, YEAR, MONTH, DAY, WEEKDAY, EOMONTH, and EDATE
- Calculate ages, lengths of stay, and turnaround times (DATEDIF, YEARFRAC, NETWORKDAYS, WORKDAY)
- Do time math for ED waits and overnight shifts (MOD, [h]:mm)

## 📖 Guide

The examples use the lesson workbook. On the **Stays** sheet, column E is DOB, F is EDArrivalDateTime, G is AdmitDateTime, and H
is DischargeDateTime. On **Claims**, E is ServiceDate, F is SubmitDate, and G is PaidDate. On **ED**, C is ArrivalDateTime, D is
TriageDateTime, E is ProviderSeenDateTime, and F is DepartureDateTime. On **Shifts**, F is ClockIn and G is ClockOut. **Settings!B2**
holds the report date, 12/31/2025, and **Holidays!A2:A19** lists the business office's holidays.

Many examples use **Stays row 5**: a patient born 09/01/1947 who arrived in the ED on Friday, January 3, 2025 at 12:51, was admitted
at 15:42, and went home on January 9 at 16:47. Open the workbook and try each example in an empty cell as you read. Type each
example on the sheet it describes (Stays examples on the Stays sheet, claim examples on the Claims sheet, and so on), or add the
sheet name, as in `=Stays!H5-Stays!G5`.

### 1. Dates are numbers

Excel stores every date as a **serial number**, which is a count of days. January 1, 1900 is day 1, so January 1, 2025 is day
45,658. What you see in the cell, such as 01/03/2025, is a number format laid over that number.

| You see | Excel stores | Why |
|---|---|---|
| 01/01/1900 | 1 | Day 1 of Excel's calendar |
| 01/01/2025 | 45658 | 45,657 days after day 1 |
| 12/31/2025 (Settings!B2) | 46022 | The report date |
| 01/03/2025 15:42 (Stays!G5) | 45660.6542 | A date plus a fraction of a day (section 2) |

To see the number behind a date, select the cell and press **Ctrl + Shift + ~** (Mac: **Control + Shift + ~**) to apply the General
format. Press **Ctrl + Z** (Mac: **⌘ + Z**) to get the date format back. **Ctrl + Shift + #** (Mac: **Control + Shift + #**)
applies a quick date format, and **Ctrl + 1** (Mac: **⌘ + 1**) opens Format Cells, where the custom date codes from Lesson 1.3
(`mm/dd/yyyy`, `ddd mmm d`) live.

Because a date is a number, ordinary math works on it:

| Formula | Result | What it does |
|---|---|---|
| `=H5-G5` | 6.0451 | Days between admission and discharge |
| `=Claims!E270+30` | 11/14/2025 | The date 30 days after a claim's service date |
| `=G5>=DATE(2025,1,1)` | TRUE | Compares two dates |
| `=MIN(G2:G370)` | 01/01/2025 21:09 | The earliest admission. MAX, COUNT, sorting, and filters work too |

**Entering dates.** Type a date in your regional order, such as 3/14/2025 in the US. Excel recognizes it, stores the serial number,
and right-aligns it like any number. A date that stays left-aligned was stored as text (section 13). Excel also recognizes the
international form 2025-03-14 whatever your regional settings, which makes it the safest way to type dates in a shared workbook.
**Ctrl + ;** (Mac: **Control + ;**) types today's date and **Ctrl + Shift + ;** (Mac: **⌘ + ;**) types the current time. Both are
fixed values that never change, which makes them useful as time stamps.

> ⚠️ **Two-digit years.** By default Excel reads a two-digit year from 00 to 29 as 2000–2029 and from 30 to 99 as 1930–1999. So a
> date of birth typed as 5/27/24 becomes 2024, which is right for an infant, but 1/1/30 becomes 1930. Always type four-digit years
> for dates of birth.

> 📋 **The 1904 date system.** Excel for Mac 2008 and earlier counted days from January 1, 1904, and any workbook can still be set
> to that system (Windows: **File → Options → Advanced → Use 1904 date system**; Mac: **Excel → Settings → Calculation**, called
> **Preferences** in older versions). Dates
> copied between a 1900 workbook and a 1904 workbook shift by 1,462 days, which is four years and a day. If pasted dates jump by four
> years, check that setting. Excel also treats 1900 as a leap year, a bug kept for compatibility with Lotus 1-2-3, so serial numbers
> before March 1, 1900 are off by one day. That never affects modern dates. Dates before 1900 can't be stored as dates at all.

### 2. Times are fractions of a day

If a whole day is 1, then an hour is 1/24 and a minute is 1/1,440. Noon is 0.5, 6:00 AM is 0.25, and 30 minutes is 0.0208. A
**date-time**, like AdmitDateTime, is a serial number with a decimal part: 45660.6542 means day 45,660 (01/03/2025) and 0.6542 of the
way through that day (15:42).

When you subtract two date-times, the result is a **duration** in days. Multiply to change the unit. In Stays row 5, the patient
spent `=G5-F5` = 0.11875 days in the ED before admission:

| To express a duration in… | Multiply by | Stays row 5: `=(G5-F5)*…` |
|---|---|---|
| hours | 24 | 2.85 |
| minutes | 1,440 (24 × 60) | 171 |
| seconds | 86,400 (24 × 60 × 60) | 10,260 |

To go the other way, divide: 90 minutes as a time value is `=90/1440`.

| Function | Returns | Example on Stays row 5 | Result |
|---|---|---|---|
| `TIME(hour, minute, second)` | A time value | `=TIME(0,30,0)` | 0.0208 (30 minutes) |
| `HOUR(serial)` | The hour on the clock, 0–23 | `=HOUR(G5)` | 15 |
| `MINUTE(serial)` | The minute on the clock, 0–59 | `=MINUTE(G5)` | 42 |
| `INT(serial)` | The date part | `=INT(G5)` | 45660, which is 01/03/2025 |
| `MOD(serial, 1)` | The time part | `=MOD(G5,1)` | 0.6542, which is 15:42 |

INT keeps the whole number, which is the date. `MOD(number, 1)` returns what is left after dividing by 1, which is the decimal part,
so it gives the time of day. You'll use both again for midnights (section 7) and overnight shifts (section 11).

> ⚠️ **HOUR and MINUTE read a clock, not a duration.** A duration of 30 hours is 1.25 days, and `=HOUR(1.25)` returns 6 because
> the clock shows 6:00 one day later. To get a duration in hours, multiply by 24.

### 3. TODAY, NOW, and a fixed report date

`=TODAY()` returns the current date and `=NOW()` returns the current date and time. Both are **volatile**, which means Excel
recalculates them every time the workbook recalculates. A report built on TODAY() gives one answer today and a different one
tomorrow, so nobody can reproduce last month's numbers.

For anything you report, put the "as of" date in one cell and point every formula at it. This workbook keeps that **report date**
in Settings!B2 (12/31/2025, the date the course data is current to).

| Question | Changes every day | Repeatable |
|---|---|---|
| Age of the Stays row 5 patient | `=DATEDIF(E5,TODAY(),"Y")` | `=DATEDIF(E5,Settings!$B$2,"Y")` → 78 |
| Days since a claim was submitted | `=TODAY()-Claims!F2` | `=Settings!$B$2-Claims!F2` |

The `$` signs lock the reference, so it stays on Settings!B2 when you copy the formula down a column (Lesson 1.5; press **F4**, or
**⌘ + T** on a Mac, to add them). To rerun the report for another date, you change one cell.

### 4. Building and taking apart dates: DATE, YEAR, MONTH, DAY

```
=DATE(year, month, day)
=YEAR(date)    =MONTH(date)    =DAY(date)
```

YEAR, MONTH, and DAY take a date apart into three numbers, and DATE puts three numbers back together into a real date.

| Formula | Result |
|---|---|
| `=YEAR(G5)` | 2025 |
| `=MONTH(G5)` | 1 |
| `=DAY(G5)` | 3 |
| `=DATE(2025,1,3)` | 45660, which displays as 01/03/2025 |

DATE **rolls over** months and days that are out of range, which makes it a small calendar calculator:

| Formula | Result | What happened |
|---|---|---|
| `=DATE(2025,13,1)` | 01/01/2026 | Month 13 rolls into the next year |
| `=DATE(2025,3,0)` | 02/28/2025 | Day 0 is the last day of the month before |
| `=DATE(2025,1,3+90)` | 04/03/2025 | 90 days after January 3 |
| `=DATE(YEAR(G5),MONTH(G5),1)` | 01/01/2025 | The first day of the admission month |
| `=DATE(YEAR(E5)+65,MONTH(E5),DAY(E5))` | 09/01/2012 | The row 5 patient's 65th birthday |

A few patterns come up in almost every report:

| You want | Formula | Stays row 5 |
|---|---|---|
| The quarter number | `=ROUNDUP(MONTH(G5)/3,0)` | 1 |
| A quarter label | `=YEAR(G5)&"-Q"&ROUNDUP(MONTH(G5)/3,0)` | 2025-Q1 |
| The same date next year | `=DATE(YEAR(G5)+1,MONTH(G5),DAY(G5))` | 01/03/2026 |

**Testing whether a date falls in a period.** Build the boundaries with DATE. "In October 2025" means on or after October 1 *and*
before November 1. With the SUMPRODUCT counting pattern from Lesson 2.1, this counts the October admissions:

```
=SUMPRODUCT((G2:G370>=DATE(2025,10,1))*(G2:G370<DATE(2025,11,1)))     → 41
```

> ⚠️ **Date-times and "on or before".** `DATE(2025,10,31)` means October 31 at 00:00, the very first moment of the day. An admission
> at 10/31/2025 21:49 is *greater* than that. So an upper bound of `<=DATE(2025,10,31)` silently drops the three admissions on
> October 31 and returns 38 instead of 41. Use `<` the first day of the next period, which works whether or not the column holds
> times. Comparing months also works: `=SUMPRODUCT((MONTH(G2:G370)=10)*(YEAR(G2:G370)=2025))`.

> ⚠️ **Don't type dates as text inside formulas.** `=G5>="10/1/2025"` compares a number with the *text* "10/1/2025". Excel ranks
> every piece of text above every number, so the result is FALSE for every row. Build the date with DATE, or point to a cell that
> holds one. (COUNTIFS is the exception: it reads a criterion like `">=10/1/2025"` as a date. Lesson 2.5 shows the safer
> `">="&DATE(2025,10,1)` form.)

### 5. Days of the week: WEEKDAY and TEXT

```
=WEEKDAY(serial_number, [return_type])
```

WEEKDAY turns a date into a day number. The **return_type** decides which day is 1:

| return_type | Numbering | `WEEKDAY(G5, …)` for Friday 01/03/2025 |
|---|---|---|
| 1 or omitted | Sunday = 1 … Saturday = 7 | 6 |
| 2 | Monday = 1 … Sunday = 7 | 5 |
| 3 | Monday = 0 … Sunday = 6 | 4 |
| 11 to 17 | Week starts on the day you choose (11 = Monday, 12 = Tuesday, … 17 = Sunday) | — |

Return type 2 makes weekend tests easy, because Saturday and Sunday are exactly the values above 5. This counts the ED's weekend
arrivals:

```
=SUMPRODUCT(--(WEEKDAY(ED!C2:C227,2)>5))     → 44 of the 226 visits
```

With the default return type, the weekend is 1 and 7, so you'd need two tests: `(WEEKDAY(d)=1)+(WEEKDAY(d)=7)`.

When you want the *name* of the day or month, use TEXT with a date format code (Lesson 2.2):

| Formula | Result |
|---|---|
| `=TEXT(G5,"dddd")` | Friday |
| `=TEXT(G5,"ddd")` | Fri |
| `=TEXT(G5,"mmmm")` | January |
| `=TEXT(G5,"mmm yyyy")` | Jan 2025 |
| `=TEXT(G5,"yyyy-mm")` | 2025-01 (sorts correctly as text) |

A helper column of weekday names lets you group and average by day. For example, `AVERAGEIF(day_names,"Friday",values)` averages
the values on Friday rows. Lesson 2.5 covers AVERAGEIF fully.

> ⚠️ **TEXT returns text, not a date.** You can't add days to "Friday", and "Apr 2025" sorts before "Jan 2025" alphabetically. Keep
> the real date in its own column and use TEXT only for labels. The format codes also depend on the language of your Office
> installation: a German Excel uses `TTTT` instead of `dddd` and returns "Freitag".

For week numbers, `=WEEKNUM(G5)` counts weeks that start on Sunday, with the week containing January 1 as week 1, and
`=ISOWEEKNUM(G5)` uses the ISO standard (weeks start on Monday). Both return 1 for 01/03/2025.

> 📋 **Version note:** WEEKDAY return types 11–17 need Excel 2010 or later. ISOWEEKNUM needs Excel 2013 or later on Windows or Excel
> 2016 or later on a Mac.

### 6. Whole months: EDATE and EOMONTH

Months have 28 to 31 days, so "one month later" can't be done by adding a fixed number of days. Two functions do month math for you:

```
=EDATE(start_date, months)       the same day number, that many months later (negative = earlier)
=EOMONTH(start_date, months)     the LAST day of the month, that many months later
```

Claims row 270 is claim CLM519179, with a ServiceDate of 10/15/2025:

| Formula | Result | Meaning |
|---|---|---|
| `=E270+30` | 11/14/2025 | 30 days later (not the same as a month) |
| `=EDATE(E270,1)` | 11/15/2025 | One month later, same day number |
| `=EDATE(E270,-6)` | 04/15/2025 | Six months earlier |
| `=EOMONTH(E270,0)` | 10/31/2025 | The last day of the service month |
| `=EOMONTH(E270,1)` | 11/30/2025 | The last day of the following month |
| `=EOMONTH(E270,-1)+1` | 10/01/2025 | The first day of the service month |
| `=DAY(EOMONTH(E270,0))` | 31 | The number of days in the service month |

When the target month is shorter, EDATE stops at its last day: `=EDATE(DATE(2025,8,31),1)` returns 09/30/2025. That makes EDATE the
right tool for birthdays and anniversaries. `=EDATE(Stays!E5,65*12)` returns 09/01/2012, the same 65th birthday as the DATE formula
in section 4. In a hospital you'll use EDATE for six-month follow-up visits and credential renewals, and EOMONTH for month-end close and
"by the end of next month" deadlines.

> ⚠️ **Format the result as a date.** EDATE, EOMONTH, and WORKDAY return a serial number, and a General cell shows it as a number
> like 46021. Press **Ctrl + Shift + #** (Mac: **Control + Shift + #**) to see the date.

> ⚠️ **One date at a time.** Give EDATE and EOMONTH a single date. Passing a whole range, as in `EOMONTH(E2:E370,0)`, returns
> #VALUE!, even in Microsoft 365. Work row by row in a helper column instead.

### 7. Date arithmetic: lengths of stay and turnaround times

Subtracting one date from another gives the days between them. When the cells hold times too, the answer has a decimal part.

```
=end - start                     elapsed days (with a decimal when the cells hold times)
=DAYS(end_date, start_date)      the same subtraction as a function; note that the END date comes first
```

**Length of stay** (LOS) can be measured three ways, and they answer different questions. Compare Stays row 5 with row 140, a patient
admitted on 05/03/2025 at 23:22 and discharged on 05/05/2025 at 08:57:

| Measure | Formula | Row 5 | Row 140 |
|---|---|---|---|
| Elapsed days | `=H5-G5` | 6.05 | 1.40 |
| Elapsed hours | `=(H5-G5)*24` | 145.1 | 33.6 |
| Midnights (calendar days) | `=INT(H5)-INT(G5)` | 6 | 2 |

**Elapsed days** measure the actual time in a bed, to the minute, which suits hour-based targets and throughput analysis. A
**midnight** count compares the two dates and ignores the clock, so row 140 crosses two midnights in under 34 hours. Midnights are how
a midnight census counts patient days (Lesson 1.4) and how room-and-board days are usually billed. Many hospitals count a same-day
stay as one day: `=MAX(1,INT(H2)-INT(G2))`.

> 📋 **Excel Tables:** the data sheets are Excel Tables. When you type a formula like `=H2-G2` into the first cell of an empty Table
> column, Excel fills the whole column for you. If you click the cells instead of typing their addresses, Excel may write
> `=[@DischargeDateTime]-[@AdmitDateTime]`, where `[@DischargeDateTime]` means "DischargeDateTime in this row." Both give the same
> result. Lesson 3.1 covers these structured references.

**Turnaround and aging** use the same subtraction. Claim CLM519179 (Claims row 270) was served on 10/15/2025, submitted on
10/21/2025, and paid on 12/10/2025:

| Question | Formula | Result |
|---|---|---|
| Days from service to submission | `=F270-E270` | 6 |
| Days from submission to payment | `=G270-F270` or `=DAYS(G270,F270)` | 50 |
| Days since submission, as of the report date | `=Settings!$B$2-F270` | 71 |

> ⚠️ **Blank dates count as zero.** An unpaid claim has no PaidDate, and Excel treats the empty cell as 0 (day zero, before 1900).
> `=G2-F2` then returns minus the submit date's serial number, about −46,000 days for a 2025 claim, which wrecks any average.
> Guard the formula: `=IF(G2="","",G2-F2)`. The empty text `""` keeps the row blank, and AVERAGE ignores text.

> ⚠️ **A day count that looks like a date.** If you subtract dates in a cell with the General format, Excel often copies the date
> format from the cells you referenced. Then 50 days displays as 02/19/1900. The number is right, so change the cell's format to
> General or Number.

> 📋 **Version note:** DAYS needs Excel 2013 or later on Windows, Excel 2016 or later on a Mac, or Excel for the web. Plain
> subtraction works everywhere.

### 8. Ages: DATEDIF and YEARFRAC

```
=DATEDIF(start_date, end_date, unit)
```

**DATEDIF** counts the complete years, months, or days between two dates. The start date comes first, and the unit goes in quotes.

| unit | Returns | `=DATEDIF(E5,G5,unit)` on Stays row 5 |
|---|---|---|
| `"Y"` | Complete years | 77 |
| `"M"` | Complete months | 928 |
| `"D"` | Days | 28249 |
| `"YM"` | Months left over after the complete years | 4 |

Join them to state an age the way a chart does: `=DATEDIF(E5,G5,"Y")&" y "&DATEDIF(E5,G5,"YM")&" m"` returns **77 y 4 m**.

> ⚠️ **DATEDIF is hidden.** Excel keeps DATEDIF for compatibility with old Lotus 1-2-3 workbooks, so it doesn't appear in
> AutoComplete or the Insert Function list, and no ScreenTip shows its arguments. Type it in full. If the start date is later than
> the end date, it returns #NUM!. Microsoft also warns that the `"MD"` unit can return wrong results, so don't use it.

**Why not just divide by 365?** Stays row 10 is a patient born 01/10/1973 and admitted 01/06/2025, four days before turning 52.
`=DATEDIF(E10,G10,"Y")` correctly returns 51. `=INT((G10-E10)/365)` returns 52, because the 13 leap days the patient has lived push
the count past the birthday early. Dividing by 365.25 is closer but still slips by a day near some birthdays. DATEDIF compares
calendar dates, so it's always exact.

```
=YEARFRAC(start_date, end_date, [basis])
```

**YEARFRAC** returns the years between two dates *with* a decimal. The **basis** argument chooses how days are counted:

| basis | Day count | Typical use |
|---|---|---|
| 0 or omitted | US 30/360: every month counts as 30 days | Bond interest (it's the default!) |
| 1 | Actual/actual: real days ÷ real year length | Ages, years of service |
| 2 | Actual/360 | Some loan interest |
| 3 | Actual/365 | Some loan interest |
| 4 | European 30/360 | European bonds |

`=YEARFRAC(E5,G5,1)` returns 77.34 for the row 5 patient. Over long spans the bases barely differ, but over months they can disagree
in the second decimal place, so always pass basis 1 for ages.

| You need | Use |
|---|---|
| Age as stated on a chart ("77") | `DATEDIF(dob, date, "Y")` |
| An infant's age in months | `DATEDIF(dob, date, "M")` |
| Age or years of service with decimals | `YEARFRAC(start, end, 1)` |
| Age as of the report date | `DATEDIF(dob, Settings!$B$2, "Y")` |

> 💡 **Tip:** DATEDIF and YEARFRAC use only the date part of a date-time, so you can point them straight at AdmitDateTime.

### 9. Business days: NETWORKDAYS and WORKDAY

Billing offices, appeals teams, and HR count **business days**: Monday through Friday, minus holidays. Two functions do the counting:

| Function | Question it answers | Is the start date counted? |
|---|---|---|
| `NETWORKDAYS(start_date, end_date, [holidays])` | How many business days from start through end? | Yes. Both ends count |
| `WORKDAY(start_date, days, [holidays])` | Which date is *days* business days after start? | No |

The optional **holidays** argument is a range of dates to skip. In this workbook that's `Holidays!$A$2:$A$19`.

| Formula | Result |
|---|---|
| `=NETWORKDAYS(DATE(2025,12,1),DATE(2025,12,31),Holidays!$A$2:$A$19)` | 22 business days in December 2025: 23 weekdays minus Christmas |
| `=NETWORKDAYS(F270,G270,Holidays!$A$2:$A$19)` | 35 business days for claim CLM519179, against 50 calendar days |
| `=WORKDAY(DATE(2025,12,22),5,Holidays!$A$2:$A$19)` | 12/30/2025. It skips Christmas. Without the holiday list, 12/29/2025 |
| `=WORKDAY(DATE(2026,1,9),-6,Holidays!$A$2:$A$19)` | 12/31/2025. A negative count goes backward, here skipping New Year's Day |

A few rules keep these functions honest:

- **Lock the holiday range** with `$` (or use the Table reference `tblHolidays[Date]`), so it doesn't slide when you copy the
  formula down.
- **Use observed dates.** Independence Day 2026 falls on a Saturday, so the list holds Friday 07/03/2026, the day the office is
  actually closed. A holiday that lands on a weekend changes nothing, because weekends are already skipped.
- **Both functions use only the date part**, so a date-time works as a start or end date.

> ⚠️ **NETWORKDAYS counts both ends.** A claim submitted and paid on the same Tuesday has a NETWORKDAYS of 1, not 0. For "business
> days elapsed," subtract 1.

When the weekend isn't Saturday and Sunday, use the **.INTL** versions, which add a **weekend** argument:

```
=NETWORKDAYS.INTL(start_date, end_date, [weekend], [holidays])
=WORKDAY.INTL(start_date, days, [weekend], [holidays])
```

| weekend | Days off |
|---|---|
| 1 or omitted | Saturday and Sunday |
| 2 | Sunday and Monday |
| 7 | Friday and Saturday |
| 11 | Sunday only |
| 17 | Saturday only |
| `"0000011"` | Seven digits for Monday through Sunday, where 1 means a day off (this one is Saturday and Sunday) |

A clinic open Monday through Saturday had `=NETWORKDAYS.INTL(DATE(2025,12,1),DATE(2025,12,31),11,Holidays!$A$2:$A$19)` = 26 working
days in December 2025.

> 📋 **Version note:** NETWORKDAYS and WORKDAY are built into Excel 2007 and later (Excel 2003 needed the Analysis ToolPak add-in).
> The .INTL versions need Excel 2010 or later.

### 10. Time math: ED waits in minutes

ED row 2 is visit ED210677. The patient arrived on 10/01/2025 at 18:52, was triaged at 19:01, was seen by a provider at 19:28, and
left at 23:13.

| Measure | Formula | Result |
|---|---|---|
| Door-to-triage minutes | `=(D2-C2)*1440` | 9 |
| Door-to-provider minutes | `=(E2-C2)*1440` | 36 |
| ED length of stay in hours | `=(F2-C2)*24` | 4.35 |
| ED length of stay as a time | `=F2-C2`, formatted `[h]:mm` | 4:21 |

Because the ED sheet stores the full date *and* time, a wait that crosses midnight needs no special handling. Visit ED212063 (row 198)
arrived on 12/18/2025 at 23:34 and was seen on 12/19/2025 at 01:17, and `=(E198-C198)*1440` correctly returns 103.

A patient who **left without being seen** (LWBS) has no ProviderSeenDateTime. Guard those rows the same way as unpaid claims:
`=IF(E2="","",(E2-C2)*1440)`.

> ⚠️ **Rounding dust.** Most times can't be stored exactly in binary, so a difference in minutes can come out a hair above or below
> the whole number, such as 35.9999999999 instead of 36. The cell displays 36, but an exact test like `=(E2-C2)*1440=36` or
> `<=30` can then give the wrong answer. Round before you compare with a threshold: `=ROUND((E2-C2)*1440,0)<=30`.

### 11. Clock times without dates: MOD for overnight shifts

Many time-clock exports list punch times only, with no date. The Shifts sheet is like that: ClockIn and ClockOut are times of day.
Subtraction works for a day shift, but not for a shift that ends the next morning:

| Shift | ClockIn | ClockOut | ClockOut − ClockIn | MOD(ClockOut − ClockIn, 1) |
|---|---|---|---|---|
| Row 2, day shift | 06:57 | 19:33 | `=G2-F2` → 12:36 | `=MOD(G2-F2,1)` → 12:36 |
| Row 21, night shift | 18:55 | 07:39 | `=G21-F21` → ######## (−0.4694 of a day) | `=MOD(G21-F21,1)` → 12:44 |

07:39 minus 18:55 is negative, and Excel can't display a negative time, so the cell fills with `#`. The fix is MOD:

```
=MOD(ClockOut - ClockIn, 1)
```

`MOD(number, 1)` always returns a value of at least 0 and less than 1. A positive duration passes through unchanged. A negative one
gets a whole day added, so −0.4694 becomes 0.5306, which is 12:44. `=IF(G21<F21,G21+1-F21,G21-F21)` does the same thing in a longer way.

To get paid time, subtract the 30-minute unpaid meal break as a time: `=MOD(G21-F21,1)-TIME(0,30,0)` returns 12:14.

> ⚠️ **Half an hour is not 0.5.** In Excel's time math, 0.5 is half a *day*. Subtract `TIME(0,30,0)`, `30/1440`, or `"0:30"`.

> ⚠️ **MOD only works for spans under 24 hours.** With times alone, a 26-hour on-call shift looks exactly like a 2-hour one. When a
> shift can run past 24 hours, record full date-times and subtract them normally.

### 12. Showing durations: h:mm versus [h]:mm

Durations are stored in days, and the number format decides how you read them. Unit secretary EMP2124 worked five 8-hour shifts this
pay week, and their paid time adds up to 1.6958 days:

| Format | Shows | What it means |
|---|---|---|
| General | 1.695833 | Days |
| `h:mm` | 16:42 | Hours of the clock. It wraps every 24 hours and shows only what's left over |
| `[h]:mm` | 40:42 | Total hours. The square brackets let hours go past 24 |
| `[m]` | 2442 | Total minutes |
| `=total*24`, formatted `0.00` | 40.70 | Decimal hours, the form payroll uses |

To apply one, press **Ctrl + 1** (Mac: **⌘ + 1**), choose **Number → Custom**, and type the code in the **Type** box.

> ⚠️ **Multiply by 24 before you multiply by a rate.** At EMP2124's rate of $19.83 an hour, `=total*19.83` pays $33.63, because
> the total is 1.6958 *days*. `=total*24*19.83` correctly pays $807.08.

### 13. Text that looks like a date: DATEVALUE and TIMEVALUE

Exports from other systems sometimes store dates as text. Signs of text dates: they stay left-aligned, `=ISNUMBER(A2)` returns FALSE,
SUM and MAX skip them, and comparisons like `>=DATE(2025,1,1)` treat them as text (section 4). Plain arithmetic such as `=A2+1` may
still work, because Excel converts date-like text on the fly, but only when the text matches your regional date order. Otherwise it
returns #VALUE!. Convert text dates once, with DATEVALUE and TIMEVALUE, so every formula sees a real date:

| Formula | Result |
|---|---|
| `=DATEVALUE("2025-12-31")` | 46022 |
| `=TIMEVALUE("7:30 PM")` | 0.8125 |
| `=DATEVALUE("12/31/2025")` | 46022 when your computer's regional settings put the month first, as in the US, but #VALUE! where dates are written day first |
| `=--"2025-12-31"` | 46022. The double minus converts text to a number (Lesson 2.1) |

The ISO form yyyy-mm-dd converts correctly under any regional setting. Lesson 3.3 covers cleaning whole columns of mixed text dates.

### 14. Quick reference

| Action | Windows | Mac |
|---|---|---|
| Insert today's date (fixed value) | Ctrl + ; | Control + ; |
| Insert the current time (fixed value) | Ctrl + Shift + ; | ⌘ + ; |
| Apply a date format | Ctrl + Shift + # | Control + Shift + # |
| Apply a time format | Ctrl + Shift + @ | Control + Shift + @ |
| Show the serial number (General format) | Ctrl + Shift + ~ | Control + Shift + ~ |
| Format Cells, for codes like `[h]:mm` | Ctrl + 1 | ⌘ + 1 |
| Lock a reference with `$` | F4 | ⌘ + T |

| Functions | Available in |
|---|---|
| DATE, YEAR, MONTH, DAY, WEEKDAY, TIME, HOUR, MINUTE, TODAY, NOW, DATEVALUE, TIMEVALUE, INT, MOD | Every version |
| DATEDIF | Every version, but hidden from AutoComplete |
| EDATE, EOMONTH, YEARFRAC, NETWORKDAYS, WORKDAY, WEEKNUM | Excel 2007 and later |
| NETWORKDAYS.INTL, WORKDAY.INTL | Excel 2010 and later |
| DAYS, ISOWEEKNUM | Excel 2013 and later (Excel 2016 and later on a Mac) |

All of them work in Microsoft 365 and Excel for the web.

## 🧪 Hands-on practice

Download [`2.3-date-time-functions.xlsx`](2.3-date-time-functions.xlsx) and open the **Practice** sheet. Type each answer in the
yellow cell, as a formula wherever possible. The **Check** column turns green when you're right.

<!-- BEGIN GENERATED: practice -->
Tasks 1–7 use the Stays sheet, 8–11 the Claims sheet, 12 the ED sheet, and 13 the Shifts sheet. Settings!B2 holds the report date (12/31/2025) and the Holidays sheet lists the business office's holidays. Several tasks ask you to fill a yellow column on a data sheet: type the formula in the first data row and the Table fills the rest (if it doesn't, double-click the fill handle). The gray cell on this sheet then summarizes your column.

| # | Task | Hint |
|:-:|------|------|
| 1 | Stays!G2 shows the first admission of 2025: 01/01/2025 21:09. What number does Excel actually store in that cell? Point a formula at the cell (the answer cell is already formatted to show 4 decimal places). | A date-time is one number: whole days since 1900, plus a fraction of a day |
| 2 | The patient in Stays row 15 (encounter ENC110973) was 64 at admission. On what date does the patient turn 65, the usual age of Medicare eligibility? Use the DOB in that row. | EDATE moves a date by whole months. How many months are in 65 years? |
| 3 | Fill the yellow AgeAtAdmit column on the Stays sheet with each patient's age in completed years on the admit date. Start in I2. The gray cell averages your column. What was the average age at admission? | DATEDIF(start_date, end_date, "Y") counts completed years. Type it in full, because Excel won't suggest it |
| 4 | The youngest patient (Stays row 42, encounter ENC112029) was an infant. Calculate the exact age in years at admission with YEARFRAC, using basis 1 (actual/actual). Round to 2 decimal places. | YEARFRAC(start_date, end_date, basis). Don't skip the third argument |
| 5 | How many stays were admitted on a weekend (Saturday or Sunday)? Use the AdmitDateTime column. | WEEKDAY with return_type 2 numbers Monday as 1 and Sunday as 7 |
| 6 | How many stays were admitted in April 2025? Compare AdmitDateTime with dates you build with DATE. | On or after April 1, and before May 1 |
| 7 | Fill the yellow LOSDays column with each stay's length of stay in days, including the fraction of a day (DischargeDateTime minus AdmitDateTime). Start in J2. The gray cell averages your column. What was the average length of stay? | Later date-time minus earlier date-time gives days |
| 8 | Ashby Falls' billing standard says a claim must be submitted by the last day of the month after the month of service. What is the deadline for claim CLM511227 (Claims row 29, ServiceDate 01/30/2025)? | EOMONTH(start_date, months) returns the last day of a month |
| 9 | Fill the yellow DaysToPay column on the Claims sheet with the calendar days from SubmitDate to PaidDate. 94 claims have no PaidDate yet, so make those rows return "" (empty text). Start in H2. The gray cell averages your column. What is the average days to pay? | Test for a blank PaidDate with IF before you subtract |
| 10 | Claim CLM520056 (Claims row 307) is still Pending. As of the report date in Settings!B2, how many business days has it been waiting? Count from its SubmitDate through the report date, both days included, skipping weekends and the dates on the Holidays sheet. | NETWORKDAYS(start_date, end_date, holidays) |
| 11 | Claim CLM517132 (Claims row 207) was only partially paid by Evergreen Mutual Insurance, and the payment posted on its PaidDate. Bluestone Health's policy gives the appeals team 10 business days to send an underpayment appeal, so the deadline is the 10th business day after the PaidDate (the PaidDate itself doesn't count). Weekends and the dates on the Holidays sheet aren't business days. What is the deadline? | WORKDAY(start_date, days, holidays) returns a date |
| 12 | Fill the yellow DoorToProviderMin column on the ED sheet with the minutes from ArrivalDateTime to ProviderSeenDateTime. 9 patients left without being seen and have no ProviderSeenDateTime, so make those rows return "". Start in H2. The gray cell averages your column. What is the average door-to-provider time in minutes? | A difference of date-times is in days. A day has 24 × 60 = 1,440 minutes |
| 13 | The Shifts sheet's ClockIn and ClockOut columns hold clock times only, with no dates, and 69 shifts end after midnight. Fill the yellow PaidTime column with each shift's paid time: clock-out minus clock-in, corrected for shifts that cross midnight, minus a 30-minute unpaid meal break. Keep each result as a time value (don't multiply by 24). Start in H2. The gray cell totals your column in [h]:mm format. What is the total paid time for the week? | MOD(…, 1) turns a negative time difference into the right positive one. Half an hour is a time value, not 0.5 |
<!-- END GENERATED: practice -->

## ✅ Answer key

The workbook has a hidden **Answer Key** sheet (right-click any sheet tab → **Unhide…** → *Answer Key*). Its *Live result* column
runs a working formula for every formula-based answer, so you can watch it calculate. The same answers are below, collapsed so you don't see them by accident.

<!-- BEGIN GENERATED: answers -->
<details>
<summary><b>🔑 Show the answer key</b> — Try every task before opening this.</summary>

**1. Stays!G2 shows the first admission of 2025: 01/01/2025 21:09. What number does Excel…**

- **Answer:** 45,658.8812
- **Solution:** `=Stays!G2`

The whole part, 45,658, is the **date serial number**: January 1, 2025 is day 45,658 counting from January 1, 1900. The decimal part, 0.8812, is the time as a fraction of a 24-hour day: 21:09 is 1,269 minutes ÷ 1,440 minutes per day. You can also see the number by giving the cell the General format with Ctrl + Shift + ~ (Mac: Control + Shift + ~). Because dates and times are numbers, you can add, subtract, and compare them.

**2. The patient in Stays row 15 (encounter ENC110973) was 64 at admission. On what date…**

- **Answer:** 03/19/2025
- **Solution:** `=EDATE(Stays!E15,65*12)`

EDATE moves a date forward (or back) by whole months and keeps the day number, so 65 × 12 = 780 months after the DOB is the 65th birthday. `=DATE(YEAR(Stays!E15)+65,MONTH(Stays!E15),DAY(Stays!E15))` gives the same date. It takes the date apart with YEAR, MONTH, and DAY, adds 65 to the year, and builds it again with DATE. Adding 65 × 365 days would land 16 days early, because it ignores the 16 leap days in between.

**3. AgeAtAdmit column with DATEDIF (average age)**

- **Answer:** 58.12
- **Solution:** `=DATEDIF(E2,G2,"Y")`

DATEDIF counts whole birthdays passed, which is how age is stated on a chart. Shortcuts like `=INT((G2-E2)/365)` drift by a day for every leap year lived, and in this data they make 14 patients a year too old. Excel doesn't list DATEDIF in AutoComplete or Insert Function, so type it in full. The start date must come first, or DATEDIF returns #NUM!. 159 of the 369 patients were 65 or older at admission.

**4. The youngest patient (Stays row 42, encounter ENC112029) was an infant. Calculate the…**

- **Answer:** 0.70
- **Solution:** `=ROUND(YEARFRAC(Stays!E42,Stays!G42,1),2)`

YEARFRAC returns the fraction of a year between two dates: 0.6959 here. Basis 1 counts the real days (254 of them) and divides by the real length of the year. Without the basis argument YEARFRAC uses basis 0 (the 30/360 banking convention, where every month has 30 days) and returns 0.6889, which rounds to 0.69. Use basis 1 for ages. For infants, clinicians usually state age in months instead: `=DATEDIF(Stays!E42,Stays!G42,"M")` returns 8.

**5. How many stays were admitted on a weekend (Saturday or Sunday)? Use the AdmitDateTime…**

- **Answer:** 95
- **Solution:** `=SUMPRODUCT(--(WEEKDAY(Stays!G2:G370,2)>5))`

`WEEKDAY(date,2)` returns 1 for Monday through 7 for Sunday, so Saturday and Sunday are exactly the values above 5. Comparing the whole column gives a list of TRUE/FALSE values, `--` turns them into 1s and 0s, and SUMPRODUCT adds them (Lesson 2.1). With the default return_type 1 (Sunday = 1, Saturday = 7) you'd need two tests: `=SUMPRODUCT((WEEKDAY(Stays!G2:G370)=1)+(WEEKDAY(Stays!G2:G370)=7))`.

**6. How many stays were admitted in April 2025? Compare AdmitDateTime with dates you build…**

- **Answer:** 27
- **Solution:** `=SUMPRODUCT((Stays!G2:G370>=DATE(2025,4,1))*(Stays!G2:G370<DATE(2025,5,1)))`

DATE(2025,4,1) builds the serial number for April 1, and multiplying the two TRUE/FALSE lists keeps only rows that pass both tests (AND logic). The upper bound is **before May 1**, not on or before April 30. AdmitDateTime includes a time, so 04/30/2025 19:59 is *greater* than DATE(2025,4,30), which means midnight. `<=DATE(2025,4,30)` would miss the 3 admissions on April 30 and return 24. `=SUMPRODUCT((MONTH(Stays!G2:G370)=4)*(YEAR(Stays!G2:G370)=2025))` also works.

**7. LOSDays column (average length of stay)**

- **Answer:** 4.54
- **Solution:** `=H2-G2`

Subtracting two date-times gives the elapsed days, with the hours as a decimal: 2.50 is two and a half days. In a cell with the General format, Excel may copy the date format from the cells you referenced and show 4.5 days as 01/04/1900 12:00. Change the format to Number when that happens. Multiply by 24 to get hours. The live formula in the key uses a shortcut: the sum of the differences equals the difference of the sums.

**8. Ashby Falls' billing standard says a claim must be submitted by the last day of the…**

- **Answer:** 02/28/2025
- **Solution:** `=EOMONTH(Claims!E29,1)`

`EOMONTH(date,0)` is the last day of the date's own month, and `EOMONTH(date,1)` is the last day of the next month. It handles 28-, 29-, 30-, and 31-day months for you. Adding 30 days instead would give 03/01/2025, which is in the wrong month. `=EOMONTH(date,-1)+1` gives the first day of the date's month, another pattern you'll use often.

**9. DaysToPay column (average days to pay)**

- **Answer:** 27.6
- **Solution:** `=IF(G2="","",G2-F2)`

An empty PaidDate counts as 0 in arithmetic, so a bare `=G2-F2` returns minus the submit date's serial number for every unpaid claim (−45,666 for the first one), which wrecks the average. The IF returns empty text for those rows, and AVERAGE ignores text. `=IF(G2="","",DAYS(G2,F2))` gives the same result. Notice that DAYS takes the **end** date first, and that it needs the same blank guard.

**10. Claim CLM520056 (Claims row 307) is still Pending. As of the report date in…**

- **Answer:** 28
- **Solution:** `=NETWORKDAYS(Claims!F307,Settings!$B$2,Holidays!$A$2:$A$19)`

NETWORKDAYS counts Monday-to-Friday dates from the start date through the end date, **including both**, and skips any date in the holidays range. Without the holiday list the answer would be 31. The three missing days are Thanksgiving, the day after, and Christmas. The `$` signs lock the Settings and Holidays references, so you can copy the formula down a column without them moving.

**11. Claim CLM517132 (Claims row 207) was only partially paid by Evergreen Mutual…**

- **Answer:** 12/03/2025
- **Solution:** `=WORKDAY(Claims!G207,10,Holidays!$A$2:$A$19)`

WORKDAY steps forward the given number of business days, never counting the start date itself, and skips weekends and holidays. This window crosses 2 holidays (Thanksgiving Day, Day after Thanksgiving). Without the holiday list WORKDAY would return 12/01/2025, 2 days too early. WORKDAY returns a serial number, so if you see a number like 46000 instead of a date, give the cell a Date format. NETWORKDAYS measures a span you already have, and WORKDAY finds the date at the end of a span.

**12. DoorToProviderMin column (average minutes)**

- **Answer:** 49.1
- **Solution:** `=IF(E2="","",(E2-C2)*1440)`

Subtracting gives a fraction of a day (0.0347 is 50 minutes), and multiplying by 1,440 converts days to minutes. Because these columns hold the full date *and* time, visit ED211281 (arrived 23:51, seen at 00:03 the next morning) still gives a positive 12 minutes, with no special handling. The IF keeps LWBS visits out of the average instead of counting them as huge negative waits.

**13. PaidTime column with MOD (total paid time for the week)**

- **Answer:** 2213:48 (h:mm)
- **Solution:** `=MOD(G2-F2,1)-TIME(0,30,0)`

Shift SH915484 clocked in at 18:48 and out at 08:20. 08:20 − 18:48 is negative (−0.4361 of a day). MOD(…, 1) adds one whole day to a negative result and leaves positive results alone, so the shift becomes 13:32. Subtract the meal break as a time, `TIME(0,30,0)` or `"0:30"` or `30/1440`. Subtracting 0.5 would remove half a *day*. The total is 2,213.8 hours, so it needs **[h]:mm**. Plain h:mm would wrap past every 24 hours and show only the leftover hours. Without MOD the week totals 557.8 hours, because every overnight shift comes out negative. Format your column as [h]:mm (or h:mm) to read each shift.

</details>
<!-- END GENERATED: answers -->

## 🏆 Bonus challenge

<!-- BEGIN GENERATED: bonus -->
Ashby Falls' utilization review nurse is preparing for a Medicare audit. Under the CMS two-midnight benchmark, an inpatient admission is generally expected to span at least two midnights of hospital care. This bonus uses a simplified, educational version (not billing guidance): count the midnights between two date-times by comparing their dates, not the hours between them. The nurse also wants to know whether the day of the week a patient is admitted affects how long they stay. Fill the yellow Midnights, BenchMidnights, and AdmitDay columns on the Stays sheet as you go. B2, B4, and B5 also use your LOSDays column from task 7.

Work on the **Bonus** sheet of the workbook.

- **B1.** Fill the Midnights column with the number of midnights each stay crossed between AdmitDateTime and DischargeDateTime. Start in K2. The gray cell counts your rows with 2 or more. How many stays crossed at least two midnights? *(Hint: INT strips the time from a date-time. Subtract the two dates)*
- **B2.** How many stays crossed two or more midnights even though they lasted LESS than 48 hours? *(Hint: Two conditions on two of your columns: multiply the TRUE/FALSE lists. 48 hours is 2 days)*
- **B3.** CMS starts the benchmark clock when hospital care begins, which for ED admissions is the ED arrival, not the inpatient admit order. Fill the BenchMidnights column: midnights from EDArrivalDateTime to DischargeDateTime, or from AdmitDateTime when EDArrivalDateTime is blank. Start in L2. The gray cell counts your rows with 2 or more. How many stays meet the benchmark when ED time counts? *(Hint: Use IF to pick the clock start for each row, then count midnights the same way as in B1)*
- **B4.** Fill the AdmitDay column with the weekday name of each AdmitDateTime (Monday, Tuesday, …). What was the average length of stay, in HOURS, for patients admitted on a Friday? Round to 1 decimal place. *(Hint: TEXT(date, "dddd") gives the weekday name. AVERAGEIF (a preview of Lesson 2.5) or AVERAGE(IF(…)) averages the Friday rows)*
- **B5.** Which weekday of admission has the LONGEST average length of stay? Type the weekday name. *(Hint: Build a small seven-row table of averages, one per weekday)*
<!-- END GENERATED: bonus -->

<!-- BEGIN GENERATED: bonus-answers -->
<details>
<summary><b>🔑 Show the bonus solution</b> — Give it a real try first!</summary>

**B1. Midnights column (stays with 2 or more midnights)**

- **Answer:** 356
- **Solution:** `=INT(H2)-INT(G2)`

INT(date-time) drops the decimal part, leaving the date at midnight. The difference between two dates is the number of midnights crossed: admitted Monday 23:00 and discharged Wednesday 01:00 is 2 midnights in only 26 hours. Rounding LOSDays doesn't work: that stay's LOSDays is 1.08, which rounds to 1. 13 stays fall short of the benchmark.

**B2. How many stays crossed two or more midnights even though they lasted LESS than 48 hours?**

- **Answer:** 28
- **Solution:** `=SUMPRODUCT((Stays!K2:K370>=2)*(Stays!J2:J370<2))`

LOSDays measures elapsed time, and 2 days is 48 hours. Midnights measures calendar dates. A patient admitted late in the evening crosses a midnight within minutes, so a stay of 30 hours can span 2 midnights. That's why a two-midnight review can't use "LOS ≥ 2 days" as a shortcut.

**B3. BenchMidnights column (stays meeting the benchmark with ED time)**

- **Answer:** 357
- **Solution:** `=INT(H2)-INT(IF(F2="",G2,F2))`

The IF picks the start of care for each row, and INT turns it into a date. 23 ED patients arrived before midnight and were admitted after it, so each gains a midnight. Only 1 of them moves from below the benchmark to meeting it. The rest already had 2 or more. `=INT(H2)-INT(MIN(F2:G2))` is a shorter trick that works because MIN ignores the blank cell, but the IF says what you mean.

**B4. Fill the AdmitDay column with the weekday name of each AdmitDateTime (Monday, Tuesday,…**

- **Answer:** 108.3
- **Solution:** `=ROUND(AVERAGEIF(Stays!M2:M370,"Friday",Stays!J2:J370)*24,1)`

Fill AdmitDay with `=TEXT(G2,"dddd")`, which returns full weekday names. AVERAGEIF averages the LOSDays values on rows where AdmitDay is "Friday", and × 24 converts days to hours. In Microsoft 365 and Excel 2021, `=ROUND(AVERAGE(IF(Stays!M2:M370="Friday",Stays!J2:J370))*24,1)` also works, as you saw in Lesson 2.1. TEXT returns weekday names in your Office language, so a German Excel shows "Freitag".

**B5. Which weekday of admission has the LONGEST average length of stay? Type the weekday name.**

- **Answer:** Thursday
- **Solution:**

1. On the Bonus sheet, type the seven weekday names, Monday to Sunday, in G2:G8.
2. In H2, enter `=AVERAGEIF(Stays!$M$2:$M$370,G2,Stays!$J$2:$J$370)*24` and fill it down to H8.
3. Find the largest average by eye, or let Excel find it: `=INDEX(G2:G8,MATCH(MAX(H2:H8),H2:H8,0))` (INDEX and MATCH are covered in Lesson 2.6).


Thursday admissions stay longest, at 121.8 hours on average, against 113.1 for Monday, the runner-up, and 98.8 for Sunday. A weekly pattern like this often points to weekend gaps in services: a patient admitted late in the week may wait through the weekend for a test, a procedure, or a discharge placement.

</details>
<!-- END GENERATED: bonus-answers -->

## Key takeaways

- A date is a **serial number** of days since 1900, and a time is a **fraction of a day**. Multiply a duration by 24 for hours or by
  1,440 for minutes.
- Put the "as of" date in one cell, like Settings!B2, instead of using TODAY(), so a report gives the same answer every time you
  open it.
- **DATE** builds dates, and **YEAR, MONTH, DAY, and WEEKDAY** take them apart. Test a period with `>=` its first day and `<` the
  first day of the next period, because date-times on the last day are later than midnight.
- **EDATE** moves by whole months and **EOMONTH** jumps to a month's last day. Both return serial numbers, so format the result as a
  date.
- State ages with **DATEDIF(…,"Y")**, and use **YEARFRAC** with basis 1 when you need decimals. Never divide by 365.
- **NETWORKDAYS** counts business days (both ends included), and **WORKDAY** finds the date a number of business days away. Give
  both a locked holiday list.
- For clock times without dates, **MOD(out − in, 1)** fixes overnight shifts, and **[h]:mm** shows totals past 24 hours.

<!-- BEGIN GENERATED: nav -->
---

⬅️ **Previous:** [2.2 Text Functions](../02-text-functions/README.md) · 🏠 [Course home](../../README.md) · **Next:** [2.4 Math & Statistical Functions](../04-math-statistical-functions/README.md) ➡️
<!-- END GENERATED: nav -->
