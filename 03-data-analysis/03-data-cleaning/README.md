# Lesson 3.3 · Cleaning Messy Data

> **Level:** Intermediate · **Time:** about 145 minutes · **Workbook:** [`3.3-data-cleaning.xlsx`](3.3-data-cleaning.xlsx)
> **Data:** A 650-row patient registration export from Bluestone's legacy registration system (names, DOBs, phones, insurance, and MRNs typed every which way, plus duplicates), a payer mapping table, a one-week bed-board census export, and a 116-line lab interface feed from Cedar Ridge Medical Center. The raw export is also available as [`data/messy/patient_registrations_raw.csv`](../../data/messy/patient_registrations_raw.csv), and the [data dictionary](../../data/README.md) describes every course file.

Every hospital report starts with data that someone typed or some system exported. Registration clerks type names in capitals or
in lowercase, an old interface strips the leading zeros from MRNs, and one insurance plan ends up spelled five different ways.
Messy data almost never causes an error message. It causes **wrong answers**: a COUNTIF that misses `Medicare ` because of a
trailing space, a lookup that can't find MRN `1955231`, and a patient counted twice because she registered once as "Amber White"
and once as "WHITE, AMBER". In this lesson you clean a realistic registration export step by step. You keep the original
untouched, fix each column with formulas and Excel's cleaning tools, remove duplicates safely, and record what you changed so that
you or a colleague can repeat the job next month.

## What you'll learn

- Spot common data problems: stray spaces, inconsistent case and categories, text dates, duplicates
- Fix them with TRIM, PROPER, SUBSTITUTE, VALUE, DATEVALUE, and lookup mapping tables
- Use Text to Columns, Flash Fill, Remove Duplicates, and Go To Special
- Document a repeatable cleaning process

## 📖 Guide

### 1. A cleaning workflow you can repeat

**Data cleaning** turns data that many people and systems typed, exported, and merged into data you can count, look up, and join
reliably. It goes faster, and more safely, when you always work in the same order:

| Step | The question | Tools in this lesson |
|---|---|---|
| 1. Protect | Is the original export safe? | A copy of the sheet, and formulas in new columns |
| 2. Profile | What is wrong, and in how many rows? | Filter lists, COUNTIF, LEN vs LEN(TRIM), EXACT, COUNT vs COUNTA |
| 3. Standardize text | Is the same thing typed in different ways? | TRIM, CLEAN, UPPER, LOWER, PROPER, SUBSTITUTE, Find & Replace |
| 4. Convert types | Are numbers and dates stored as text? | VALUE, DATEVALUE, DATE, TEXT, Text to Columns |
| 5. Map categories | Does one value have many spellings? | A mapping table and XLOOKUP |
| 6. Restructure | Is there one fact per cell, and a label on every row? | Text to Columns, Flash Fill, Go To Special |
| 7. Deduplicate | Is the same patient in the file twice? | Remove Duplicates, COUNTIF, UNIQUE |
| 8. Validate and document | Did it work, and can someone repeat it? | Check formulas and a cleaning log |

The order matters most at step 7. Excel can only recognize a duplicate after the rows look alike, so you standardize first and
deduplicate second. Practice task 3 shows what happens when you skip ahead.

### 2. Protect the raw data

**Raw data** is the export exactly as you received it. Never type over it. You need it to re-run your cleaning when you find a
mistake, to prove what the source system said, and to compare against next month's export.

The workbook follows the pattern you should use at work:

| Sheet | Role |
|---|---|
| **Raw** (gray tab) | The export as received. Read it, never edit it |
| **Clean** | A copy of Raw with yellow cleaning columns on the right. Each cleaning column holds a formula that reads the messy value in the same row |
| **PayerMap** | A mapping table that turns insurance spellings into payer IDs (section 9) |
| **CleaningLog** | The record of every step: the problem, the fix, and how many rows it changed (section 15) |

To copy a sheet, right-click its tab → **Move or Copy…** → tick **Create a copy** → **OK**. The shortcut is to hold **Ctrl**
(Mac: **Option**) while you drag the tab.

Raw and Clean share the same columns and rows, so a worked example below such as `=TRIM(B10)` means "row 10's PatientName". Try
the examples on the Clean sheet in column V or further right, never on Raw. Keep them out of the yellow columns, because a formula
typed into an empty Table column fills the whole column (a **calculated column**, [Lesson 3.1](../01-tables-named-ranges/README.md)).
Keep them out of column T too, because typing right next to a Table adds that column to the Table.

If you point at a cell with the mouse while you build a formula in a yellow column, Excel writes a structured reference such as
`=TRIM([@PatientName])` instead of `=TRIM(B2)`. From column V, outside the Table, it writes `tblClean[@PatientName]`. All of
these mean "this row's PatientName", so any form is correct. The answer key uses the A1 form because it is shorter.

Cleaning in new columns, instead of overwriting the messy ones, has three advantages. You see the before and after side by side,
the logic is visible in the formula bar, and you can fix a mistake by editing one formula. When the clean data is final, you can
freeze it: select the clean columns, copy them, press Ctrl + Alt + V (Mac: ⌃ + ⌘ + V) to open **Paste Special**, choose
**Values**, and click **OK**.

> 📋 **Why Raw is all text.** When you double-click a CSV file, Excel converts values as it opens them: `04979796` becomes the
> number 4979796, `03-Apr-1957` becomes a date, `01.11.2003` stays text, and long ID numbers lose every digit after the 15th. The
> Raw sheet keeps every value exactly as the source system wrote it, which is what you get when every column is imported as Text
> ([Lesson 4.3](../../04-advanced-analysis/03-power-query/README.md) shows how with Power Query). In Microsoft 365 you can switch some of these conversions off under **File → Options →
> Data → Automatic data conversion** (Mac: **Excel → Preferences → Edit**).

### 3. Profile before you fix

**Profiling** means measuring what is wrong, and how often, before you change anything. It tells you which fixes you need, and its
counts become the "before" numbers in your cleaning log.

| Question | Quick look | Formula |
|---|---|---|
| Which spellings exist? | Turn on filters with Ctrl + Shift + L (Mac: ⌘ + Shift + F) and open the column's dropdown | `=UNIQUE(Raw!D2:D651)` (Microsoft 365 or Excel 2021+) |
| How many rows hold one value? | Filter, then read the status bar | `=COUNTIF(Raw!D2:D651,"Female")` (counts `FEMALE` too) |
| How many values have extra spaces? | You can't see them | `=SUMPRODUCT(--(LEN(Raw!H2:H651)<>LEN(TRIM(Raw!H2:H651))))` |
| How many are in ALL CAPS? | Scroll and squint | `=SUMPRODUCT(--EXACT(Raw!B2:B651,UPPER(Raw!B2:B651)))` |
| How many match a pattern? | Find All with Ctrl + F (Mac: ⌃ + F) | `=COUNTIF(Raw!B2:B651,"*,*")` counts names that contain a comma |
| Real numbers and dates, or text? | Alignment, green triangles, status bar | `=COUNT(Raw!C2:C651)` compared with `=COUNTA(Raw!C2:C651)` |
| Placeholder text posing as a blank? | The filter list | `=COUNTIF(Raw!F2:F651,"N/A")` |

The extra-space formula deserves a closer look, because you'll use the pattern throughout this lesson.
`LEN(Raw!H2:H651)<>LEN(TRIM(Raw!H2:H651))` compares all 650 rows at once and returns 650 TRUE/FALSE results. TRUE means TRIM would
shorten that value, so it has extra spaces. The double minus `--` turns TRUE into 1 and FALSE into 0, and SUMPRODUCT adds them up.
On the Insurance column the answer is **37**: 25 rows of `Medicare ` and 12 rows of `Silverline Medicare Advantage `, each with a
trailing space. In Microsoft 365 or Excel 2021+, `=SUM(…)` works in place of SUMPRODUCT.

**Two traps while profiling.** The filter dropdown and UNIQUE both ignore capitals, so `F` and `f` appear as one entry. To count
case problems, use **EXACT**, which compares two texts character by character and is case-sensitive. Spaces, on the other hand,
do count, so `F` and `·F` (with a leading space) appear as separate entries. Also glance at the **status bar** when you select a
column. If it shows only *Count* and no *Sum* or *Average*, nothing in the selection is a number, so a column of "dates" is
really text.

Here is what profiling finds in the Raw sheet (each `·` below marks a space):

| Column | What profiling finds |
|---|---|
| PatientName | Two layouts: 92 rows are `First Last` and the rest `Last, First`. 131 are in ALL CAPS and 60 in lowercase. Some have extra spaces, such as row 10, `··Alexander·,··Gary·` |
| DOB | Text, not dates: `=COUNT(Raw!C2:C651)` is 0. Five layouts: `2002-03-31` (153 rows), `03-Apr-1957` (92), `08/06/1989` or `3/22/1961` (360), and `01.11.2003` (45) |
| Sex | 10 spellings of two values, including `Female`, `MALE`, `f`, and `·F` with a leading space |
| Phone | Six layouts: `(555) 476-7432`, `555-722-8468`, `555.468.7891`, `5553869767`, `+1 555 597 3811`, and blank |
| Email | 150 truly blank, plus 132 **placeholders** (69 `N/A` and 63 `none`), 65 padded with spaces, and 50 in capitals |
| CityStateZip | ZIP+4 codes mixed with 5-digit ZIPs, some cities in capitals, and some rows with no comma |
| Insurance | 30 spellings of 7 payers, 37 of them with a trailing space |
| MRN | 86 with an `MRN-` prefix, and 141 shorter than 8 digits because their leading zeros were lost |
| RegisteredOn | Text in two layouts: `2015-11-06 18:33:00` (403 rows) and `05/30/2022 05:42 PM` (247) |
| Whole rows | Exact copies, and the same patients registered twice with different typing |

A **placeholder** is a value that means "nothing here", such as `N/A`, `none`, `UNKNOWN`, `0`, or `1/1/1900`. Placeholders make counts
lie: `=COUNTA(Raw!F2:F651)` counts `N/A` as an email address on file. Turn them into real blanks during cleaning, and list the
placeholders you found in your log.

### 4. Spaces and capitals

[Lesson 2.2](../../02-formulas-functions/02-text-functions/README.md) introduced these functions one cell at a time. In cleaning
you apply them to whole columns:

| Function | What it does | Use it for |
|---|---|---|
| `TRIM(text)` | Removes spaces at both ends and shrinks inner runs of spaces to one | Every text column |
| `CLEAN(text)` | Removes non-printing characters such as line breaks | Text pasted from electronic health record (EHR) notes |
| `SUBSTITUTE(text,UNICHAR(160)," ")` | Turns non-breaking spaces into normal ones, which TRIM can then remove | Text copied from web pages |
| `UPPER(text)` | ALL CAPITALS | Codes such as sex, state, or lookup keys |
| `LOWER(text)` | all lowercase | Email addresses |
| `PROPER(text)` | First Letter Of Each Word | Names and cities |

**Worked example: cleaning the Email column.** Emails should be lowercase and trimmed, and placeholders should become blank:

| Email (row) | Formula | Result |
|---|---|---|
| `·carlos.santos91@example.com·` (31) | `=LOWER(TRIM(F31))` | `carlos.santos91@example.com` |
| `THERESA.RYAN88@EXAMPLE.COM` (17) | `=LOWER(TRIM(F17))` | `theresa.ryan88@example.com` |
| `none` (8) | `=IF(OR(TRIM(F8)="N/A",TRIM(F8)="none"),"",LOWER(TRIM(F8)))` | *(empty)* |

The last formula works on every row: it returns an empty result for the two placeholders and a clean address otherwise. The `=`
comparison ignores case, so it would also catch `NONE` or `None`.

**Standardizing a short code.** The Sex column has 10 spellings, but every one of them starts with the right letter once the
leading space is gone. So TRIM, then take the first character with LEFT, then capitalize it with UPPER. That shortcut works only
because F and M are each identified by their first letter. Medicare and Medicaid share a first letter, so the Insurance column needs
the mapping table in section 9.

**PROPER's blind spots.** PROPER capitalizes every letter that follows a non-letter, so `O'BRIEN` correctly becomes `O'Brien`.
It also turns McDonald into `Mcdonald`, and this file has two McDonalds (rows 201 and 507). Fix known exceptions by hand, or with
`=SUBSTITUTE(PROPER(A2),"Mcd","McD")`, and note them in the cleaning log.

> ⚠️ `=` and COUNTIF ignore case: `="f"="F"` is TRUE. When a check must notice capitals, such as "how many cells are exactly F",
> use EXACT: `=SUMPRODUCT(--EXACT(range,"F"))`.

### 5. One column, two layouts: patient names

92 names are written `First Last` and the rest `Last, First`. When one column mixes layouts, your formula has to recognize the
layout first and then apply the right fix:

```
=IF(test for layout A, fix for layout A, fix for layout B)
```

The test here is whether the name contains a comma: `ISNUMBER(FIND(",",B2))`. FIND returns the comma's position, or #VALUE! when
there is no comma, and ISNUMBER turns that into TRUE or FALSE. Here are the pieces for row 10, `··Alexander·,··Gary·`:

| Formula | Result | What it does |
|---|---|---|
| `=FIND(",",B10)` | 13 | Position of the comma |
| `=MID(B10,FIND(",",B10)+1,LEN(B10))` | `··Gary·` | Everything after the comma. LEN(B10) is simply a length that is long enough |
| `=TRIM(MID(B10,FIND(",",B10)+1,LEN(B10)))` | `Gary` | The first name, trimmed |
| `=TRIM(LEFT(B10,FIND(",",B10)-1))` | `Alexander` | Everything before the comma: the last name |

Join the first name, a space, and the last name with `&`. That is the fix for `Last, First` rows. The fix for names that are
already `First Last` is just TRIM. Wrap the whole IF in PROPER so that both layouts get the same capitals. Practice task 10 asks you
to put this together. In Microsoft 365 or Excel 2024, `TRIM(TEXTAFTER(B10,","))` and `TRIM(TEXTBEFORE(B10,","))` give the
same two pieces.

> ⚠️ **Names are not identifiers.** Different patients can share a name, and one patient's name can change through marriage, a
> legal name change, or a typo. Use names to read the data, and the MRN to match and deduplicate it.

### 6. Removing characters with SUBSTITUTE: phone numbers

```
=SUBSTITUTE(text, old_text, new_text)
```

Replace a character with `""` (nothing) to delete it, and nest SUBSTITUTEs to delete several characters. Read a nested formula
from the inside out: the innermost SUBSTITUTE runs first and hands its result to the next one. Row 3 holds `(555) 476-7432`:

| Formula | Result |
|---|---|
| `=SUBSTITUTE(E3,"(","")` | `555) 476-7432` |
| `=SUBSTITUTE(SUBSTITUTE(E3,"(",""),")","")` | `555 476-7432` |
| `=SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(E3,"(",""),")","")," ","")` | `555476-7432` |

Before you write the formula, list which characters appear in each layout. That tells you how many layers you need:

| Layout | Characters to remove |
|---|---|
| `(555) 476-7432` | `(` `)` space `-` |
| `555-722-8468` | `-` |
| `555.468.7891` | `.` |
| `5553869767` | none |
| `+1 555 597 3811` | `+`, the spaces, and the country code `1` |

The country code is the odd one out. After the other characters are gone, `+1 555 597 3811` is `+15555973811`. A U.S. phone number is
always the **last** 10 digits, so `RIGHT(…,10)` around the whole formula drops the `+1` and leaves 10-digit numbers unchanged.

> 💡 **Tip:** Build a nested formula one layer at a time, and press Enter after each layer to check the result.

> 📋 **Microsoft 365:** `=REGEXREPLACE(E2,"[^0-9]","")` removes every character that isn't a digit in one step. `[^0-9]` means "any
> character that is not 0 through 9". REGEXREPLACE is new (2024) and only in Microsoft 365.

### 7. Numbers stored as text, and identifiers that must stay text

Two kinds of data look like numbers, and they need opposite treatment:

| Kind | Examples | Store as | Why |
|---|---|---|---|
| **Quantities** | Charges, lab results, census counts, doses | Numbers | You add, average, and compare them |
| **Identifiers** | MRN, ZIP, phone, NPI (National Provider Identifier), account number | Text | Leading zeros matter, nobody adds MRNs, and IDs longer than 15 digits lose precision as numbers |

**Spotting numbers stored as text.** They line up on the left, Excel marks them with a small green triangle (*Number Stored as
Text*), SUM ignores them, and COUNT is lower than COUNTA. `=ISTEXT(A2)` tells you for sure.

**When quantities arrive as text,** convert them:

| Method | How | Result |
|---|---|---|
| VALUE | `=VALUE(A2)` | A number, in a new column |
| Double minus | `=--A2` | The same, shorter |
| Error button | Select the cells, click the warning icon → **Convert to Number** | Converted in place |
| Paste Special → Multiply | Copy a cell that contains 1, select the cells, then Paste Special → **Multiply** ([Lesson 1.2](../../01-foundations/02-data-entry-autofill/README.md)) | Converted in place |
| Text to Columns | Select one column → **Data → Text to Columns** → **Finish** | Converted in place (section 10) |

**When identifiers have lost their zeros,** rebuild them as text. Row 11's MRN is `1955231`:

| Formula | Result | How it works |
|---|---|---|
| `=TEXT(VALUE(I11),"00000000")` | `01955231` | Make it a number, then write it as text with exactly 8 digits. Each 0 in the format is a required digit |
| `=RIGHT("0000000"&I11,8)` | `01955231` | Glue seven zeros on the front, then keep the last 8 characters |

Both formulas leave an MRN that is already correct, like `05927600`, unchanged. Row 5 has a third problem, an `MRN-` prefix
(`MRN-05927600`), so remove that with SUBSTITUTE before you pad.

> ⚠️ A custom number format of `00000000` only changes what you see. The cell still stores 1955231, and an XLOOKUP for `"01955231"`
> won't find it, because lookups treat the number 1955231 and the text `01955231` as different values
> ([Lesson 2.6](../../02-formulas-functions/06-lookup-functions/README.md), section 8).

> ⚠️ **Find & Replace can strip zeros too.** If you replace `MRN-` with nothing in a General-formatted column, Excel re-reads each
> result as if you had typed it, so `05927600` becomes the number 5927600. Format the column as Text first (**Home → Number Format
> → Text**), or use a formula.

### 8. Dates stored as text

A **text date** looks like a date but is stored as characters. The signs: it sits on the left of the cell, changing the number
format does nothing, it sorts alphabetically (so `01.11.2003` lands above `08/06/1989`), and `=COUNT(Raw!C2:C651)` returns 0.

Three functions turn date text into real dates and times, which Excel stores as serial numbers
([Lesson 2.3](../../02-formulas-functions/03-date-time-functions/README.md)). Format the results as dates with
**Home → Number Format → Short Date**.

| Function | Reads | Example | Result |
|---|---|---|---|
| `DATEVALUE(text)` | A date. Any time in the text is ignored | `=DATEVALUE("3/22/1961")` | 22362 (3/22/1961) |
| `VALUE(text)` or `--text` | A date, a time, or both | `=VALUE("05/30/2022 05:42 PM")` | 44711.7375 |
| `TIMEVALUE(text)` | A time | `=TIMEVALUE("05:42 PM")` | 0.7375 (17:42) |

A **date-time** is the date's serial number plus the time's fraction of a day: 44711.7375 is 5/30/2022 (44711) plus 17:42
(0.7375). So when the date and the time need different handling, convert them separately and add them.
`=DATE(2022,5,30)+TIMEVALUE("05:42 PM")` returns the same 44711.7375. Format a date-time column with a custom format such as
`mm/dd/yyyy hh:mm` to see both parts.

**How DATEVALUE reads the five DOB layouts** on a computer with U.S. regional settings:

| DOB (row) | `=DATEVALUE(C…)` | Why |
|---|---|---|
| `2002-03-31` (2) | 3/31/2002 | ISO year-month-day is read the same way in every region |
| `03-Apr-1957` (3) | 4/3/1957 | The month is a word, so there's no doubt |
| `08/06/1989` (4) | 8/6/1989 | U.S. settings read slashes month first |
| `3/22/1961` (20) | 3/22/1961 | One-digit months and days are fine |
| `01.11.2003` (30) | #VALUE! | U.S. settings don't accept a dot as a date separator |

The dot is easy to fix: swap the dots for slashes first, so `=DATEVALUE(SUBSTITUTE(C30,".","/"))` returns 1/11/2003.

> ⚠️ **Regional settings decide how DATEVALUE reads slashes.** On a computer set to a day-first region, such as the UK or most of
> Europe, `08/06/1989` silently becomes 8 June 1989, and `3/22/1961` returns #VALUE! because there is no month 22. What matters is the
> order the *source system* used. This export comes from a U.S. registration system, so its slash dates are month first, whatever
> your computer says.

**The region-proof method: build the date from its pieces.** `DATE(year, month, day)` doesn't depend on regional settings. Take each
piece out of the text with LEFT, MID, and RIGHT (Lesson 2.2):

| Layout | year | month | day |
|---|---|---|---|
| `2002-03-31` | `LEFT(C2,4)` | `MID(C2,6,2)` | `RIGHT(C2,2)` |
| `03-Apr-1957` | `RIGHT(C3,4)` | from the month name (below) | `LEFT(C3,2)` |
| `08/06/1989` | `RIGHT(C4,4)` | `LEFT(C4,2)` | `MID(C4,4,2)` |
| `3/22/1961` | `RIGHT(C20,4)` | the text before the first `/` | the text between the slashes |

Two tricks complete the table:

- **A month name to a number.** `=(SEARCH(MID(C3,4,3),"JanFebMarAprMayJunJulAugSepOctNovDec")+2)/3` returns 4. SEARCH finds `Apr`
  at position 10 of the 36-letter string, and because every name is 3 letters long, (10+2)/3 gives the month number.
- **Months and days of different lengths.** In `3/22/1961` the first slash is at position 2, and in `08/06/1989` it's at position 3.
  Find it with `FIND("/",C20)` and call that position *p*. The month is `LEFT(C20,p-1)`. The year and its slash are always the last
  5 characters, so the day is `MID(C20,p+1,LEN(C20)-p-5)`. For dotted dates, search `SUBSTITUTE(C20,".","/")` instead of `C20`. It
  has the separators in the same positions, so LEFT and MID still work on the original text.

To choose between the layouts, test where the first dash is. `MID(C2,5,1)="-"` is TRUE only for ISO dates, and `MID(C2,3,1)="-"`
only for `03-Apr-1957`. Nest the tests in IF ([Lesson 2.1](../../02-formulas-functions/01-logical-functions/README.md)), and you have practice task 11's formula. LEFT, MID, and RIGHT return
text, but DATE accepts text digits like `"1961"` and converts them for you.

> ⚠️ **DATE never complains.** `=DATE(1961,22,3)` doesn't return an error. It returns 10/3/1962, because month 22 rolls over into
> the following year. A formula that puts the day where the month belongs produces real-looking, wrong dates. So validate every
> converted date column:
>
> - `=COUNT(range)` equals the number of rows, so every row converted.
> - `=MIN(range)` and `=MAX(range)` are plausible: no birth dates after the report date, and none before 1900.
> - Days above 12 exist. A month number is never above 12, so rows with month and day swapped can't have them. In a correct column
>   of real birthdays, more than half of the days are 13 or later.
> - Spot-check one row of each raw layout against the text.

**Text to Columns as a date converter.** Select a column of text dates → **Data → Text to Columns** → **Delimited** → **Next** →
untick every delimiter → **Next** → choose **Date** and the order the data uses (**MDY** for this export) → **Finish**. Because you
state the order, the result doesn't depend on your regional settings. It applies one order to the whole column, though, and this DOB
column mixes year-first ISO dates with month-first ones, so practice task 11 uses a formula instead.

### 9. Standardizing categories with a mapping table

The Insurance column holds 30 spellings of 7 payers. `=COUNTIF(Raw!H2:H651,"Medicare")` returns 87, because it catches `Medicare`,
`MEDICARE`, and `medicare` but misses `Medicare ` (with a trailing space) and `Medicare Part A`. In fact 135 rows are Medicare
patients. A **mapping table** (also called a *crosswalk*) lists every spelling you have seen next to the one standard value it means:

| Variant | PayerID | StandardName |
|---|---|---|
| MEDICARE | PY01 | Medicare |
| MEDICARE PART A | PY01 | Medicare |
| SILVERLINE MA | PY02 | Silverline Medicare Advantage |
| UNINSURED | PY07 | Self-Pay |
| … | … | … |

| Approach | Example | Verdict |
|---|---|---|
| Nested IF or SWITCH | `=IF(H2="Medicare","PY01",IF(H2="Medicare Part A","PY01",…))` | 30 branches buried in a formula. A new spelling means editing every copy |
| Mapping table + lookup | `=XLOOKUP(TRIM(H2),tblPayerMap[Variant],tblPayerMap[PayerID],"UNMAPPED")` | One short formula. The map is visible, reviewable, and grows |

To build and use a map:

1. **List the distinct spellings.** Copy the column to an empty sheet and run **Data → Remove Duplicates**, or in Microsoft 365
   or Excel 2021+ type `=UNIQUE(TRIM(Raw!H2:H651))`. Both ignore capitals. UNIQUE(TRIM(…)) returns 19 spellings, one for each
   spelling a lookup has to recognize. Remove Duplicates leaves 21, because it keeps `Medicare ` and
   `Silverline Medicare Advantage ` (with their trailing spaces) as separate entries, so those two need trimming before they join
   the map. Paste the UNIQUE result as values before you edit it.
2. **Type the standard value next to each spelling.** Where the meaning isn't obvious, ask the data's owner and write the rule
   down. Here, Patient Access says "Uninsured" means Self-Pay, so the map records that in its Note column.
3. **Look up with TRIM.** XLOOKUP ignores capitals but not spaces, so `TRIM(H2)` makes `Medicare ` match `MEDICARE`.
4. **Return "UNMAPPED" for anything missing,** using XLOOKUP's *if_not_found* argument, instead of #N/A. Next month, filter the
   result column for UNMAPPED, add each new spelling as a new row of the map, and every formula picks it up.

PayerMap is an Excel Table, so refer to it as `tblPayerMap[Variant]` (Lesson 3.1). A Table reference grows when you add a row, but
a fixed range such as `PayerMap!$A$2:$A$20` would miss row 21. Without XLOOKUP (Excel 2019 and earlier), use
`=IFNA(VLOOKUP(TRIM(H2),tblPayerMap,2,FALSE),"UNMAPPED")`.

### 10. Splitting one column into several: Text to Columns

Interface feeds and older exports often cram several fields into one cell, separated by a character such as `|`, a comma, or a tab.
The LabFeed sheet holds 116 lab results like this one:

```
07539656|NA|138|mmol/L|12/01/2025 14:58
```

**Text to Columns** splits a column into several columns, in place:

1. Select the cells to split (one column only).
2. **Data → Text to Columns** (Windows key tips: Alt, A, E).
3. **Step 1:** choose **Delimited** when a character separates the fields, or **Fixed width** when every field starts at the same
   position, as in old mainframe reports. Click **Next**.
4. **Step 2:** tick the delimiters. For `|`, tick **Other** and type `|` in the box. The preview shows where the splits will fall.
   Click **Next**.
5. **Step 3:** click each column in the preview and set its **Column data format**, then click **Finish**.

**Fixed width** splits at character positions instead of at a delimiter. In step 2 the preview shows a ruler with break lines:
click the ruler to add a break, drag a line to move it, and double-click a line to remove it. Excel guesses breaks where it sees
columns of spaces, so check every one before you click **Next**. A bed list printed as `D330 12B 07539656` (unit in characters
1–4, bed in 6–8, MRN in 10–17) splits cleanly this way, and step 3 still lets you keep the MRN as Text.

| Step 3 format | What Excel does | Use it for |
|---|---|---|
| **General** | Numbers become numbers, date-like text becomes dates, everything else stays text | Quantities such as results and counts |
| **Text** | Keeps the characters exactly as they are | IDs with leading zeros: MRN, ZIP, NPI |
| **Date** (MDY, DMY, YMD…) | Reads the text as a date in the order you choose | Text dates, whatever your regional settings |
| **Do not import column (skip)** | Leaves that piece out | Fields you don't need |

> ⚠️ **General destroys leading zeros.** Under General, `07539656` becomes the number 7539656. Every MRN in the LabFeed starts with 0,
> so choose **Text** for the MRN column in step 3.

> ⚠️ **The split overwrites the columns to the right.** Excel puts the pieces in the selected column and the columns next to it. If
> those cells aren't empty, it asks *"There's already data here. Do you want to replace it?"* Click **Cancel** and set a different
> **Destination** in step 3, or insert empty columns first.

> 💡 **Tip:** Excel remembers your delimiter for the rest of the session, so text you paste later may split itself at every `|`. To
> stop that, run Text to Columns once more with every delimiter unticked.

Other ways to split text, compared:

| Tool | Result | Updates when the data changes? | Best for |
|---|---|---|---|
| Text to Columns | Values, in place | No | One-off splits, and converting a column's type |
| TEXTSPLIT, TEXTBEFORE, TEXTAFTER (Microsoft 365 or Excel 2024, Lesson 2.2) | Formulas | Yes | Splits you'll repeat on new data |
| Flash Fill | Values | No | Irregular patterns you can show by example |
| Power Query (Lesson 4.3) | A refreshable table | Yes, when you click Refresh | Monthly feeds |

### 11. Flash Fill for one-off patterns

**Flash Fill** (Lesson 1.2) fills a column from examples. Type the result you want for a row or two, then press **Ctrl + E** (or
choose **Data → Flash Fill**, on both Windows and Mac). Excel looks for a rule that reproduces every example you typed and applies
it to the rest of the column.

The examples you choose decide which rule it finds. Suppose you want the 5-digit ZIP from CityStateZip and type two examples:
`45501` from `Bluestone OH 45501` and `45720` from `CEDAR RIDGE, OH 45720`. Several rules fit both of them: "the first number",
"the last number", and "the last five characters". They agree on plain ZIPs, but on `Bluestone, OH 45501-8106` (row 11) they give
`45501`, `8106`, and `-8106`. One more example, typed on a ZIP+4 row, leaves only the right rule. So give Flash Fill **one example of
each layout**, especially the unusual ones.

Then check the result with a formula, because Flash Fill guesses and never warns you. For ZIPs, `=SUMPRODUCT(--(LEN(range)<>5))`
should be 0. If some rows come out wrong, press Ctrl + Z (Mac: ⌘ + Z), type the correct value on one of the wrong rows as an extra
example, and run Flash Fill again. The **Flash Fill Options** button that appears next to the column lets you undo the fill or select
the cells it changed.

| Use Flash Fill when… | Use a formula when… |
|---|---|
| It's a one-off job on today's data | The export arrives every month |
| The pattern is easy to show by example | The rule depends on conditions (two name layouts, five date layouts) |
| You'll check every result | You want the logic documented in the cell |

> ⚠️ Flash Fill writes typed-in values, not formulas. It won't update when the source changes, and it may base its rule on a
> different column of the table than the one you had in mind. That's why the answer key always shows a formula alternative.

### 12. Filling gaps with Go To Special → Blanks

Reports printed for people often show a label only on the first row of its block, as the CensusExport sheet does:

| Facility | Unit | CensusDate | MidnightCensus |
|---|---|---|---|
| Bluestone Memorial Hospital | Cardiac Step-Down | 12/15/2025 | … |
| | | 12/16/2025 | … |
| | | … | … |
| | Intensive Care Unit | 12/15/2025 | … |

That's easy to read but useless for analysis, because a filter, a sort, a SUMIFS, or a PivotTable sees blank facility and unit names
on most rows. **Go To Special** selects cells by type, and its **Blanks** option selects only the empty cells inside your selection.
Combined with Ctrl + Enter, it fills every gap with the label above it:

1. Select the label columns of the data only, for example **A2:B50** on CensusExport. Leave out the header row and any empty rows
   below the data.
2. **Home → Find & Select → Go To Special…** (or press **F5** or **Ctrl + G**, Mac: **⌃ + G**, then click **Special…**). Choose
   **Blanks** → **OK**. Only the empty cells stay selected.
3. Type `=` and press the **Up arrow**. The active cell is the first blank, A3, so the formula reads `=A2`, "the cell above me".
4. Press **Ctrl + Enter** (Mac: **⌘ + Return**). Excel puts the formula into every selected cell, and because the reference is
   relative, each blank points to the cell directly above itself.
5. Freeze the result: select A2:B50, copy, then use **Paste Special → Values**. A formula that means "one row up" points somewhere
   else as soon as you sort.

Other Go To Special options that help with cleaning:

| Option | Selects | Cleaning use |
|---|---|---|
| **Blanks** | Empty cells | Fill gaps, or see where data is missing |
| **Constants** | Typed values, not formulas | Find numbers someone typed over a formula column |
| **Formulas**, with only **Errors** ticked | Formulas that return an error | Find every #VALUE! a DATEVALUE column produced |
| **Visible cells only** (Alt + ;, Mac: ⌘ + Shift + Z) | Only the rows a filter shows | Copy filtered rows without the hidden ones |
| **Row differences** | Cells in each row that differ from the cell in the active cell's column | Compare a raw and a clean column side by side |

### 13. Find & Replace with wildcards

Lesson 1.2 covered Find & Replace (Ctrl + H, Mac: ⌃ + H) and its options. For cleaning, its **wildcards** help you find every
variant of a messy value:

| Wildcard | Matches | Example in Find what | Finds |
|---|---|---|---|
| `*` | Any number of characters | `* ,*` | Names with a space before the comma, like row 10 |
| `?` | Exactly one character | `self?pay` | `Self Pay`, `Self-Pay`, `self-pay`, and `SELF PAY` |
| `~` | Treats the next `*` or `?` as a real character | `~*` | Cells containing an asterisk |

Find ignores capitals unless you tick **Match case**, which is why `self?pay` also finds `SELF PAY`. Before any **Replace All**,
click **Find All** and read the count at the bottom of the list. That's how many cells the replacement will change, and the number
belongs in your log.

> ⚠️ A wildcard replaces **everything it matches**. Replacing `MRN-*` with nothing doesn't strip the prefix. It empties the whole
> cell. To remove just the prefix, replace `MRN-` with no wildcard.

> ⚠️ Replace re-enters each changed value as if you had typed it. Results made only of digits turn into numbers and lose their
> leading zeros, and results that look like dates turn into dates. Work on a copy, and format ID columns as Text first.

### 14. Duplicates: define, find, count, remove

Start by deciding what makes two rows "the same". An **exact duplicate** matches in every column. A **duplicate record** describes
the same real-world thing, here the same patient, even when the typing differs. Rows 18 and 410 are the same patient:

| Column | Row 18 (raw) | Row 410 (raw) | Both, after cleaning |
|---|---|---|---|
| PatientName | `Amber White` | `WHITE, AMBER` | Amber White |
| DOB | `1996-06-24` | `6/24/1996` | 06/24/1996 |
| Sex | `FEMALE` | `·F` | F |
| Phone | `(555) 556-3833` | `5555563833` | 5555563833 |
| Email | `amber.white81@example.com` | `amber.white81@example.com` | amber.white81@example.com |
| CityStateZip | `Cedar Ridge OH 45720` | `Cedar Ridge, OH 45720` | ZIP5 45720 |
| Insurance | `Medicaid` | `STATE MEDICAID` | PY03 |
| MRN | `03929231` | `3929231` | 03929231 |
| RegisteredOn | `2024-07-10 16:38:00` | `07/10/2024 07:15 AM` | Two times on the same day |

Only the email address matches character for character, so Remove Duplicates, which needs every ticked column to match, sees two
different rows. After cleaning, the two rows agree on every cleaned column except the time of registration, and the clean MRN
identifies them as one patient. That's why you **standardize first and deduplicate second, on an ID column**.

| Goal | Tool or formula |
|---|---|
| See duplicates | **Home → Conditional Formatting → Highlight Cells Rules → Duplicate Values** ([Lesson 3.2](../02-data-validation-conditional-formatting/README.md)). It colors every copy, the first one included, and ignores capitals |
| Highlight whole rows that repeat an ID | Select A2:S651 on Clean with A2 as the active cell, then **Home → Conditional Formatting → New Rule → Use a formula**: `=COUNTIF($Q$2:$Q$651,$Q2)>1`. The `$Q2` reference keeps the column fixed, so every cell in a row tests that row's MRNClean |
| Number each copy | `=COUNTIF($Q$2:Q2,Q2)` copied down gives 1 on a value's first row, 2 on its second, and so on. The range `$Q$2:Q2` grows by one row each time the formula moves down |
| Count all copies of a value | `=COUNTIF($Q$2:$Q$651,Q2)` |
| Count distinct values | `=SUMPRODUCT(1/COUNTIF(Q2:Q651,Q2:Q651))` in any version, or `=ROWS(UNIQUE(Q2:Q651))` in Microsoft 365 or Excel 2021+ |
| Delete duplicates | **Data → Remove Duplicates** (Windows key tips: Alt, A, M) |

The distinct-count formula relies on a COUNTIF behavior you haven't needed before. When the *criteria* argument is a whole range
instead of one value, COUNTIF runs once for each cell in it and returns one count per row. So `COUNTIF(Q2:Q651,Q2:Q651)` returns
650 counts, each one saying how many rows share that row's MRN. A value that appears 3 times contributes 1/3 + 1/3 + 1/3 = 1 to
the sum, so the total is the number of distinct values. COUNTIFS works the same way, with a range in any of its criteria. If the
range can contain blanks, use `=SUMPRODUCT((Q2:Q651<>"")/COUNTIF(Q2:Q651,Q2:Q651&""))`, which skips them instead of dividing by
zero. In Microsoft 365 or Excel 2021+, SUM can replace SUMPRODUCT in these formulas. In Excel 2019 and earlier, keep
SUMPRODUCT, because it handles the per-row arrays without any special entry.

> ⚠️ COUNTIF and COUNTIFS treat text that looks like a number as a number. On the raw MRN column they would count `1955231` and
> `01955231` as the same value, and they can confuse ID numbers longer than 15 digits. That's harmless once every MRN is padded to
> the same 8 characters, but on raw IDs, count with `=SUMPRODUCT(--(range="01955231"))`, which compares the text exactly.

**Remove Duplicates** asks which columns define a duplicate, then deletes every row that repeats an earlier row in those columns:

1. Click any cell in the data, then choose **Data → Remove Duplicates**. Excel selects the whole block of data around that cell.
2. Keep **My data has headers** ticked, so Excel lists the columns by name and never deletes the header row.
3. Tick only the columns that define a duplicate. **Select All** and **Unselect All** save clicks. Untick a unique ID such as
   RecordID. No two rows share it, so with it ticked Excel finds no duplicates at all.
4. Click **OK** and read the message.

Know its rules before you trust it:

- It keeps the **first** row of each group and deletes the later ones. It removes the whole row of your selection, not just the
  ticked columns.
- It ignores capitals but not spaces: `Medicare` matches `MEDICARE` but not `Medicare `.
- It compares what the cells **display**, so the same date shown in two different formats counts as two different values.
- It finishes with a message such as *"12 duplicate values found and removed; 238 unique values remain."* Both numbers belong in
  your log.
- Ctrl + Z (Mac: ⌘ + Z) undoes it right away. Once the file is saved and closed, the deleted rows are gone, so run it on a copy.

> 💡 **Tip: keep the newest row, not the first.** Because Remove Duplicates keeps the first row it meets, sort the data first with
> the newest registration at the top, then remove duplicates on the ID column. The bonus does the same job with formulas, so you can
> see exactly which row survives.

### 15. Validate and document

**Validation** proves that the cleaning worked. Run checks like these on every clean column:

| Check | Formula idea | You expect |
|---|---|---|
| Every row converted | `=COUNT(Clean!L2:L651)` for a date column | The number of rows |
| No failed lookups | `=COUNTIF(Clean!P2:P651,"UNMAPPED")` | 0 |
| IDs have the right length | `=SUMPRODUCT(--(LEN(Clean!Q2:Q651)<>8))` | 0 |
| Values are plausible | `=MIN(…)` and `=MAX(…)` | No birth dates after the report date |
| No formula errors | **Go To Special → Formulas → Errors** | Nothing selected |
| Rows add up | Rows in = rows kept + rows removed | Matches your log |
| Spot checks | Filter one example of each raw layout | Matches the raw text by eye |

A **cleaning log** records what you did so that someone else can review it and you can repeat it. Keep it in the workbook, as the
CleaningLog sheet does. For each step, record:

- the column and the problem you found, with the count from profiling,
- the fix: the formula, or the tool and its settings (for example "Text to Columns, delimiter |, MRN as Text"),
- the **rows changed**,
- any business rule or exception, such as "Uninsured = Self-Pay" or "McDonald fixed by hand in rows 201 and 507",
- who did it, and when.

The rows-changed count compares the old and new values case-sensitively. For example, this counts how many emails the
`LOWER(TRIM(…))` fix from section 4 changes:

```
=SUMPRODUCT(--NOT(EXACT(Raw!F2:F651,LOWER(TRIM(Raw!F2:F651)))))      → 184
```

EXACT returns TRUE where nothing changed, NOT flips it, and SUMPRODUCT counts the rows. Use EXACT rather than `=`, because `=` treats
`f` and `F` as equal and would miss every change that only fixed capitals. A count that is far higher or lower than you expected is
often the first sign of a wrong formula.

**Make it repeatable.** When next month's export arrives, paste it into the first ten columns of Clean (and into Raw, for the
record), and every formula column recalculates. The steps you did by hand, such as Flash Fill, Text to Columns, Go To Special, and
Remove Duplicates, don't repeat themselves, so write them in the log as numbered steps. Lesson 4.3's **Power Query** records every
cleaning step and replays them all with one click on **Refresh**, which is the next level of this lesson.

### 16. Shortcuts and version notes

| Action | Windows | Mac |
|---|---|---|
| Filters on/off | Ctrl + Shift + L | ⌘ + Shift + F |
| Go To (then **Special…**) | F5 or Ctrl + G | ⌃ + G |
| Put one entry in every selected cell | Ctrl + Enter | ⌘ + Return |
| Flash Fill | Ctrl + E | **Data → Flash Fill** |
| Find / Replace | Ctrl + F / Ctrl + H | ⌃ + F / ⌃ + H |
| Paste Special | Ctrl + Alt + V | ⌃ + ⌘ + V |
| Select visible cells only | Alt + ; | ⌘ + Shift + Z |
| Undo | Ctrl + Z | ⌘ + Z |
| Text to Columns | Alt, A, E | **Data → Text to Columns** |
| Remove Duplicates | Alt, A, M | **Data → Remove Duplicates** |
| Copy a sheet | Ctrl + drag the tab | Option + drag the tab |

| Feature | Needs |
|---|---|
| TRIM, PROPER, SUBSTITUTE, VALUE, DATEVALUE, TEXT, Text to Columns, Remove Duplicates, Go To Special | Every Excel version |
| Flash Fill | Excel 2013+ on Windows, Excel 2019 or Microsoft 365 on a Mac |
| MAXIFS (bonus) | Excel 2019+ |
| XLOOKUP, UNIQUE | Microsoft 365 or Excel 2021+ |
| TEXTBEFORE, TEXTAFTER, TEXTSPLIT | Microsoft 365 or Excel 2024 |
| REGEXREPLACE, automatic data conversion settings | Microsoft 365 |

## 🧪 Hands-on practice

Download [`3.3-data-cleaning.xlsx`](3.3-data-cleaning.xlsx) and open the **Practice** sheet. Most tasks fill a yellow column on the
Clean sheet, and a gray cell on the Practice sheet checks the whole column. The **Check** column turns green when you're right.

<!-- BEGIN GENERATED: practice -->
Raw is the untouched export: never edit it. Clean is your working copy of the same 650 rows, with yellow columns to fill. For a column task, type the formula in row 2 of the yellow column. Clean is an Excel Table, so the formula fills down by itself. A gray cell here then checks the whole column. Other tasks use the CensusExport, LabFeed, and CleaningLog sheets.

| # | Task | Hint |
|:-:|------|------|
| 1 | Profile first. On the Raw sheet, how many PatientName values contain extra spaces (leading, trailing, or doubled)? Compare each name's length with the length of its TRIMmed version. | LEN(x)<>LEN(TRIM(x)) is TRUE when TRIM would remove something. SUMPRODUCT(--(…)) counts the TRUEs |
| 2 | On the Clean sheet, fill the yellow SexClean column (M) with a single capital letter, F or M. The raw Sex column has 10 spellings, including Female, MALE, lowercase f, and values with a stray space. The gray cell counts cells that are exactly F (case-sensitive). | TRIM first, then take the first letter, then capitalize it |
| 3 | Make a copy of the Raw sheet (right-click its tab → Move or Copy → tick Create a copy). On the copy, run Data → Remove Duplicates on all the data (A1:J651) with every column ticked except RecordID (each row has its own RecordID, so leaving it ticked finds nothing). How many duplicate rows does Excel remove? | Untick RecordID in the Remove Duplicates dialog. Excel's message tells you the count |
| 4 | Switch to CensusExport, a bed-board report that prints each Facility and Unit only on the first row of its block. Fill every blank cell in A2:B50 with the value above it: select that range, use Go To Special → Blanks, type = and press the Up arrow, then press Ctrl + Enter (Mac: ⌘ + Return). The gray cell stays blank until every gap is filled, then totals the week's midnight census for the Intensive Care Unit at Cedar Ridge Medical Center (ICU patient days). | Go To Special selects only the blanks, and Ctrl + Enter fills them all with one relative formula |
| 5 | Back on Clean, use Flash Fill to fill the yellow ZIP5 column (O) with the 5-digit ZIP from CityStateZip. 101 rows carry a ZIP+4 such as 45501-8106, and those must become 45501. Type three examples yourself: rows 2 and 3, plus the first ZIP+4 row (row 11). Then select the first empty cell and press Ctrl + E (Mac: Data → Flash Fill). The gray cell counts the distinct ZIP codes in your column. | Give Flash Fill an example from each layout, especially a ZIP+4 row |
| 6 | Fill the yellow PhoneClean column (N) with each phone as exactly 10 digits and nothing else, or an empty result when Phone is blank. Phones arrive as (555) 476-7432, 555-722-8468, 555.468.7891, 5553869767, and +1 555 597 3811. The gray cell counts rows that end up with exactly 10 digits. | Nest one SUBSTITUTE per unwanted character, then keep the RIGHT 10 characters |
| 7 | Fill the yellow MRNClean column (Q) with every MRN as 8-digit text: 05927600, not MRN-05927600, and 01955231, not 1955231. The gray cell first checks that every value is 8 characters long, then counts the distinct MRNs, which is the number of real patients in the file. | Remove the prefix with SUBSTITUTE, make it a number with VALUE, then pad it back with TEXT(…,"00000000") |
| 8 | Fill the yellow PayerID column (P) by looking up each Insurance value in the PayerMap table, so all 30 spellings become one of seven PayerIDs. Some values carry a trailing space. Return UNMAPPED for anything missing from the map. The gray cell shows a warning if any row is unmapped. Otherwise it counts PY02 (Silverline Medicare Advantage) rows. | XLOOKUP the TRIMmed Insurance value in tblPayerMap, and use XLOOKUP's if_not_found argument for UNMAPPED |
| 9 | Switch to LabFeed: 116 STAT results from a lab interface, each crammed into one cell as MRN\|TestCode\|Result\|Units\|CollectedDateTime. Split A1:A117 into five columns with Data → Text to Columns (Delimited, Other: \|). In step 3 of the wizard, set the MRN column's format to Text so its leading zeros survive. The gray cell checks the MRNs, then shows the average potassium (TestCode K) result in mmol/L. | Text to Columns step 3: click the first column in the preview, then choose Text |
| 10 | Fill the yellow NameClean column (K) with every name as First Last in Proper Case, with single spaces. For example, `"Haddad, Jonathan"` becomes Jonathan Haddad, `"  Alexander ,  Gary "` becomes Gary Alexander, and `"Amber White"` stays Amber White. 92 rows are already First Last (no comma). The gray cell warns if any name still has extra spaces or is ALL CAPS or all lowercase. Otherwise it counts distinct names. | IF(ISNUMBER(FIND(",",B2)), flip the two parts, just TRIM), then wrap everything in PROPER |
| 11 | Fill the yellow DOBClean column (L) with real dates. DOB is text in five layouts: 2002-03-31, 03-Apr-1957, 08/06/1989, 3/22/1961, and 01.11.2003. The export is from a U.S. system, so the slash and dot layouts are month first. The gray cell checks that every row is a real date, then counts DOBs that fall on the 13th or later of their month (a quick test that month and day weren't swapped). | DATEVALUE can't read 01.11.2003 until the dots become slashes. For a version that works with any regional setting, build DATE(year, month, day) from the pieces |
| 12 | Document your work. On the CleaningLog sheet, put a formula in the yellow RowsChanged cell for step 2 (E3) that counts how many rows your SexClean column actually changed: rows where the raw Sex value is not exactly the same as SexClean (case-sensitive). The gray cell here reads your log entry. | EXACT compares case-sensitively, NOT flips TRUE and FALSE, and SUMPRODUCT(--…) counts the TRUEs |
<!-- END GENERATED: practice -->

## ✅ Answer key

The workbook has a hidden **Answer Key** sheet (right-click any sheet tab → **Unhide…** → *Answer Key*). Its *Live result* column
proves most answers straight from the Raw sheet. The same answers are below, collapsed so you don't see them by accident.

<!-- BEGIN GENERATED: answers -->
<details>
<summary><b>🔑 Show the answer key</b> — Try every task before opening this.</summary>

**1. Names with extra spaces (profiling)**

- **Answer:** 69
- **Solution:** `=SUMPRODUCT(--(LEN(Raw!B2:B651)<>LEN(TRIM(Raw!B2:B651))))`

TRIM removes leading and trailing spaces and shrinks inner runs to one space, so any name it shortens had extra spaces. Comparing lengths tests all 650 rows at once: the comparison gives TRUE/FALSE for each row, the double minus turns those into 1/0, and SUMPRODUCT adds them. You can't see a trailing space by looking at a cell, so a count like this is how you find out a problem exists before you fix it. In Microsoft 365 or Excel 2021+, `=SUM(--(LEN(…)<>LEN(TRIM(…))))` works too.

**2. SexClean column (rows that are exactly F)**

- **Answer:** 335
- **Solution:** `=UPPER(LEFT(TRIM(D2),1))`

Every spelling starts with the right letter once the stray space is gone, so the fix is TRIM, then `LEFT(…,1)`, then UPPER. The order matters: `LEFT(" F",1)` is a space. The check uses EXACT because COUNTIF ignores case and would count a lowercase f as F. This shortcut works only because each value is identified by its first letter. Medicare and Medicaid share a first letter, so Insurance needs a mapping table instead (task 8).

**3. Remove Duplicates on the raw columns**

- **Answer:** 40
- **Solution:**

1. Right-click the **Raw** tab → **Move or Copy…** → tick **Create a copy** → **OK**. Work on **Raw (2)**.
2. Click any cell in the data, then **Data → Remove Duplicates**.
3. Keep **My data has headers** ticked. Untick **RecordID** and leave the other nine columns ticked. Click **OK**.
4. Excel reports *40 duplicate values found and removed; 610 unique values remain.*

Formula check without deleting anything (Microsoft 365 or Excel 2021+): `=ROWS(Raw!B2:J651)-ROWS(UNIQUE(Raw!B2:J651))`


Remove Duplicates only removes rows that match in **every** ticked column, so it finds the 40 rows that were exported twice, character for character. That is not all of them, as task 7 will show when you count the real patients. Many more rows are re-registrations typed differently. For example, row 18 ("Amber White", (555) 556-3833) and row 410 ("WHITE, AMBER", 5555563833) are the same patient, but Remove Duplicates sees two different rows. Standardize first, then remove duplicates on the cleaned key column. UNIQUE on all nine columns counts the same thing without deleting anything. Both ignore case.

**4. Go To Special → Blanks fill-down (Cedar Ridge ICU patient days)**

- **Answer:** 71
- **Solution:**

1. Select **A2:B50** on CensusExport (not the whole columns, and not the header row).
2. **Home → Find & Select → Go To Special…** (or press **F5**, Mac: **⌃ + G**, then **Special…**). Choose **Blanks** → **OK**. Only the empty cells stay selected.
3. Without clicking anywhere, type **=** and press **↑**. The active cell is the first blank, A3, so the formula reads `=A2`.
4. Press **Ctrl + Enter** (Mac: **⌘ + Return**) to put that formula in every selected blank. Each copy points to the cell above itself.
5. Turn the formulas into values: select A2:B50, copy, then **Paste Special → Values**.

Check with `=SUMIFS(CensusExport!E2:E50,CensusExport!A2:A50,"Cedar Ridge Medical Center",CensusExport!B2:B50,"Intensive Care Unit")`


After Go To Special, Excel has selected only the 88 empty cells, and the formula you type goes into all of them at once. Because the reference is relative, each blank points to the cell directly above it, and that cell either holds the label or points further up. The unit name "Intensive Care Unit" appears at all three hospitals, so the SUMIFS needs both columns filled. With the blanks left in, it returns 0: no Cedar Ridge ICU row carries both labels, because the facility name was printed only once, on the facility's first unit. Paste the result as values before you sort, because a formula that points "one row up" points somewhere else after a sort.

**5. Flash Fill ZIP5 (distinct ZIP codes)**

- **Answer:** 12
- **Solution:**

1. In **O2** type **45501** (from *Bluestone OH 45501*), and in **O3** type **45720** (from *CEDAR RIDGE, OH 45720*).
2. In the first ZIP+4 row, **O11** (*Bluestone, OH 45501-8106*), type **45501**.
3. Select **O4**, the first empty cell, and press **Ctrl + E** (or **Data → Flash Fill**).
4. Check the result: `=SUMPRODUCT(--(LEN(Clean!O2:O651)<>5))` should return 0. If some rows are wrong, press Ctrl + Z, type the correct ZIP on one of the wrong rows as an extra example, and run Flash Fill again.

Formula alternative (Microsoft 365 or Excel 2024): `=LEFT(TEXTAFTER(G2," ",-1),5)`


Flash Fill looks for a rule that explains every example you typed. Two plain ZIPs fit several rules (the first number, the last number, the last five characters), and on 45501-8106 those give different answers. The ZIP+4 example leaves only "the first number". If the +4 survives, the distinct count jumps to 108. Flash Fill writes typed-in values, not formulas, so it won't update when next month's export arrives. The Lesson 2.2 bonus did the same job with TEXTAFTER. The formula is the repeatable choice, and Flash Fill is the fast one for a one-off.

**6. PhoneClean column (rows with exactly 10 digits)**

- **Answer:** 620
- **Solution:**

```
=RIGHT(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(E2,"(",""),")",""),"-",""),".","")," ",""),10)
```


Each SUBSTITUTE deletes one kind of character: parentheses, dashes, dots, and spaces. That leaves 5554767432 for most rows, but +15555973811 for the 40 numbers with a country code, so `RIGHT(…,10)` keeps the last ten digits, which is always the U.S. number. A blank phone needs no special handling, because SUBSTITUTE and RIGHT on an empty cell return an empty result. The 30 blanks are why the count is 620, not 650. Keep phones as text, because they are identifiers, not quantities. In Microsoft 365, `=RIGHT(REGEXREPLACE(E2,"[^0-9]",""),10)` removes every non-digit in one step.

**7. MRNClean column (distinct patients)**

- **Answer:** 560
- **Solution:** `=TEXT(VALUE(SUBSTITUTE(I2,"MRN-","")),"00000000")`

Two problems hide in this column. 86 rows carry an MRN- prefix, which SUBSTITUTE removes. 141 rows lost their leading zeros somewhere upstream, because a system treated the MRN as a number. VALUE turns the remaining text into a number, and `TEXT(…,"00000000")` writes it back as text with exactly 8 digits, adding the zeros it needs. `=RIGHT("0000000"&SUBSTITUTE(I2,"MRN-",""),8)` does the same without VALUE. The result: 650 rows but only 560 patients, so 90 rows are duplicates.

**8. PayerID column via mapping table (PY02 rows)**

- **Answer:** 68
- **Solution:** `=XLOOKUP(TRIM(H2),tblPayerMap[Variant],tblPayerMap[PayerID],"UNMAPPED")`

A mapping table turns a messy category into a standard code with one lookup, and it documents your decisions where everyone can see them, such as the rule that Uninsured means PY07 (Self-Pay). XLOOKUP ignores case, so the map needs only one row per spelling, not one per capitalization. It does not ignore spaces, so TRIM comes first. Without TRIM, the 12 rows typed "Silverline Medicare Advantage " with a trailing space come back UNMAPPED. The "UNMAPPED" result makes new spellings easy to find next month: filter for it, add the spelling as a new row of the map, and every formula picks it up. The Table references (`tblPayerMap[Variant]`) grow with the map. A fixed range such as `PayerMap!$A$2:$A$20` works today but would miss a row added below it. Without XLOOKUP, use `=IFNA(VLOOKUP(TRIM(H2),tblPayerMap,2,FALSE),"UNMAPPED")`.

**9. Text to Columns on the lab feed (average potassium)**

- **Answer:** 4.18
- **Solution:**

1. On **LabFeed**, select **A1:A117** (header included).
2. **Data → Text to Columns**. Choose **Delimited** → **Next**.
3. Untick **Tab**, tick **Other**, and type a vertical bar **|** in the box. The preview splits into five columns. Click **Next**.
4. Click the **MRN** column in the preview and choose **Text** as its column data format. Optionally click the last column and choose **Date: MDY**. Click **Finish**.

Check with `=AVERAGEIF(LabFeed!B2:B117,"K",LabFeed!C2:C117)`


Text to Columns splits each cell at every |, and step 3 decides what each piece becomes. Under **General**, Excel turns 07539656 (row 2) into the number 7539656. All 116 MRNs in this feed start with 0, so every one would break and stop matching the registration file. Choosing **Text** keeps the digits exactly as sent. The results (for example 138 and 4.6 in rows 2 and 3) should stay General so they become real numbers you can average. If you get it wrong, press Ctrl + Z (Mac: ⌘ + Z) to undo the split and run it again.

**10. NameClean column (distinct names)**

- **Answer:** 557
- **Solution:**

```
=PROPER(IF(ISNUMBER(FIND(",",B2)),TRIM(MID(B2,FIND(",",B2)+1,LEN(B2)))&" "&TRIM(LEFT(B2,FIND(",",B2)-1)),TRIM(B2)))
```


`FIND(",",B2)` returns a position when there is a comma and #VALUE! when there isn't, so `ISNUMBER(FIND(…))` tells the two layouts apart. For Last, First rows, MID takes everything after the comma (the first name) and LEFT takes everything before it (the last name). TRIM each piece before joining, because the spaces sit around the comma. PROPER on the outside fixes the case of both layouts at once. Two checks on the result. First, the flip matters: `=PROPER(TRIM(B2))` alone gives 578 distinct names, and deleting the comma without flipping still gives 569, because the same patient then appears as both "Amber White" (row 18) and "White Amber" (row 410). Second, the count is 557, not 560, because 3 names belong to two different patients (Justin Lewis, Lisa Morris, Mei Myers). That's why you deduplicate on MRN, never on name. PROPER also turns McDonald into Mcdonald, so fix known exceptions by hand and note them in the cleaning log.

**11. DOBClean column (DOBs on day 13 or later)**

- **Answer:** 397
- **Solution:**

```
=IF(MID(C2,5,1)="-",DATE(LEFT(C2,4),MID(C2,6,2),RIGHT(C2,2)),IF(MID(C2,3,1)="-",DATE(RIGHT(C2,4),(SEARCH(MID(C2,4,3),"JanFebMarAprMayJunJulAugSepOctNovDec")+2)/3,LEFT(C2,2)),DATE(RIGHT(C2,4),LEFT(C2,FIND("/",SUBSTITUTE(C2,".","/"))-1),MID(C2,FIND("/",SUBSTITUTE(C2,".","/"))+1,LEN(C2)-FIND("/",SUBSTITUTE(C2,".","/"))-5))))
```


This solution works on any computer: it reads the layout and builds DATE(year, month, day) from the pieces. ISO dates have a dash in position 5. In 03-Apr-1957 the dash is in position 3, and the month comes from where Apr sits in "JanFebMarAprMayJunJulAugSepOctNovDec" (position 10, and (10+2)/3 = 4). Everything else is month/day/year with a / or . after the month. If your Windows or Mac region is United States, the much shorter `=DATEVALUE(SUBSTITUTE(C2,".","/"))` gives identical results, because U.S. settings read the ISO and 03-Apr-1957 layouts and read slashes month first. In a day-first region it reads 08/06/1989 as 8 June and fails on 3/22/1961, and in a language whose month names differ from English it can fail on 03-Apr-1957 too. The day-13 test catches a swap: a month number is never above 12, so if the slash and dot rows had month and day swapped, none of them could show a day above 12 and the count would fall to 145 instead of 397.

**12. Cleaning log: rows changed by the Sex step**

- **Answer:** 424
- **Solution:** `=SUMPRODUCT(--NOT(EXACT(Clean!D2:D651,Clean!M2:M651)))`

A cleaning log records what you changed and how much, so the next person (or you, next month) can repeat it and audit it. "Rows changed" is the most useful number in it: here 424 of 650 rows changed and 226 were already F or M. EXACT matters again, because with `=` the values "f" and "F" count as equal and those rows would look unchanged. A count that is suspiciously high or low is often the first sign that a cleaning formula is wrong.

</details>
<!-- END GENERATED: answers -->

## 🏆 Bonus challenge

<!-- BEGIN GENERATED: bonus -->
Bluestone is moving to a new EHR, and the vendor needs a master patient load file with exactly one row per patient. The rule from Patient Access is: keep each patient's most recent registration (latest RegisteredOn). Fill the Clean sheet's last two yellow columns (RegisteredClean and Keep), using your MRNClean and PhoneClean columns from the practice tasks. RegisteredOn has two layouts: 2015-11-06 18:33:00 (24-hour clock) and 05/30/2022 05:42 PM (12-hour clock). If you want helper columns, put them on the Clean sheet in column V or further right, outside the Table.

Work on the **Bonus** sheet of the workbook.

- **B1.** Fill the yellow RegisteredClean column (R) with real date-times. Make sure 05:42 PM becomes 17:42. The gray cell checks every row, then counts registrations made at 5:00 PM or later (the evening Patient Access shift). *(Hint: In both layouts the time starts at character 12, and TIMEVALUE understands both 18:33:00 and 05:42 PM. Build the date part with DATE, then add the time to it)*
- **B2.** Fill the yellow Keep column (S) with TRUE on exactly one row per patient (MRNClean): the row with that patient's latest RegisteredClean. If the latest time appears on two identical rows, keep only the first of them. Every other row is FALSE. The gray cell counts the TRUEs. *(Hint: MAXIFS (Lesson 2.5) finds the patient's latest time. A COUNTIFS over an expanding range (`$Q$2:Q2`) breaks ties)*
- **B3.** Which RegisteredOn did you keep for MRN 01276260? Enter it as a date and time. *(Hint: MAXIFS with two conditions: MRNClean is this MRN, and Keep is TRUE. Or filter Clean on those two columns and read RegisteredClean)*
- **B4.** Remove Duplicates always keeps the first row it meets. For how many patients is the row you kept NOT that patient's first row in the file? These are the patients that Remove Duplicates on MRNClean would have gotten wrong. *(Hint: A row is a patient's first occurrence when MATCH(its MRN, the MRN column, 0) returns its own position. Or add a helper column with `=COUNTIF($Q$2:Q2,Q2)=1`)*
- **B5.** The rule "keep the latest row" decides which registration survives into the master list, and it can throw away good data. How many kept rows have an empty PhoneClean even though another registration for the same patient has a phone number? (Hint: For each row, COUNTIFS(MRN column, this MRN, PhoneClean column, `"?*"`) counts that patient's rows that have a phone)
<!-- END GENERATED: bonus -->

<!-- BEGIN GENERATED: bonus-answers -->
<details>
<summary><b>🔑 Show the bonus solution</b> — Give it a real try first!</summary>

**B1. RegisteredClean column (evening registrations)**

- **Answer:** 118
- **Solution:**

```
=IF(MID(J2,5,1)="-",DATE(LEFT(J2,4),MID(J2,6,2),MID(J2,9,2)),DATE(MID(J2,7,4),LEFT(J2,2),MID(J2,4,2)))+TIMEVALUE(MID(J2,12,8))
```


Both layouts put the time at character 12, so `TIMEVALUE(MID(J2,12,8))` reads "18:33:00" and "05:42 PM" alike, and AM/PM is handled for you. The date part differs, so the IF builds it with DATE from the right pieces. In the ISO layout the day is `MID(J2,9,2)`, not `RIGHT(J2,2)`, because the time follows it. A date-time is just date + time (a whole number plus a fraction of a day). If PM were ignored, the evening count would drop to 76. With U.S. regional settings, `=--J2` (or `=VALUE(J2)`) converts both layouts in one step.

**B2. Keep column (rows in the master list)**

- **Answer:** 560
- **Solution:** `=AND(R2=MAXIFS($R$2:$R$651,$Q$2:$Q$651,Q2),COUNTIFS($Q$2:Q2,Q2,$R$2:R2,R2)=1)`

The first test, RegisteredClean = MAXIFS(all RegisteredClean, all MRNClean, this MRN), is TRUE on the patient's latest row. Exact copies share that time, so the second test counts how many rows so far (the range `$Q$2:Q2` grows as the formula goes down) have this MRN and this time. It equals 1 only on the first of them. The total must equal the 560 distinct MRNs from practice task 7, which is a good cross-check. Avoid comparing date-times with `">"&R2` inside COUNTIFS, because that turns the time into text with 15 digits and can miss by a rounding error.

**B3. Registration kept for MRN 01276260**

- **Answer:** 12/21/2019 14:06
- **Solution:** `=MAXIFS(Clean!R2:R651,Clean!Q2:Q651,"01276260",Clean!S2:S651,TRUE)`

This patient has three rows. Two of them say 1276260 (the MRN lost its leading zero) with the time 08:22, and the third says 01276260 with 12/21/2019 02:06 PM. You only find all three after cleaning the MRN, and you only pick the right one if 02:06 PM became 14:06: read as 02:06 AM it would look older than 08:22.

**B4. Kept rows that are not the first occurrence**

- **Answer:** 27
- **Solution:** `=SUMPRODUCT(Clean!S2:S651*(MATCH(Clean!Q2:Q651,Clean!Q2:Q651,0)<>ROW(Clean!Q2:Q651)-1))`

MATCH(MRN, MRN column, 0) returns the position of the first row with that MRN. `ROW(…)-1` is each row's own position in the column (row 2 is position 1). Where they differ, the row is a repeat. Multiplying by Keep counts kept rows that are repeats. For these 27 patients a newer registration sits further down the file. To make Remove Duplicates keep the newest row, sort by RegisteredClean (newest to oldest) first, then remove duplicates on MRNClean: the first row it meets is then the newest.

**B5. Kept rows that lose a phone number**

- **Answer:** 2
- **Solution:**

```
=SUMPRODUCT(Clean!S2:S651*(Clean!N2:N651="")*(COUNTIFS(Clean!Q2:Q651,Clean!Q2:Q651,Clean!N2:N651,"?*")>0))
```


When its criteria argument is a range instead of a single value, COUNTIFS returns one count per row: how many rows share this row's MRN and have a phone. The criteria `"?*"` means at least one character, which works because PhoneClean is text. Multiply by Keep and by an empty PhoneClean to find the 2 patients who would arrive in the new EHR with no phone, although an older registration has one. Real master patient index (MPI) loads prevent this with **survivorship rules**, which decide field by field which value survives. A common rule takes each field from the most recent row that has it filled in, instead of taking every field from the most recent row.

</details>
<!-- END GENERATED: bonus-answers -->

## Key takeaways

- Never edit the raw export. Clean a copy, with formulas in new columns, so every fix is visible and repeatable.
- Profile before you fix. LEN vs LEN(TRIM), EXACT, COUNTIF with wildcards, and COUNT vs COUNTA tell you what's wrong and how often.
- Identifiers such as MRNs, ZIPs, and phone numbers stay text. Rebuild lost zeros with `TEXT(VALUE(…),"00000000")`, and keep
  General format away from them in Text to Columns and Find & Replace.
- Text dates need care. DATEVALUE follows your computer's regional settings, `DATE(year, month, day)` works everywhere, and because
  DATE never complains, every converted date column needs a validation check.
- A mapping table with `XLOOKUP(TRIM(…), …, "UNMAPPED")` turns any number of spellings into one code and documents the business
  rules behind it.
- Standardize first, then deduplicate on an ID, never on a name. Remove Duplicates keeps the first row it meets.
- Log every step with a rows-changed count. Formulas repeat themselves next month, but tool steps don't, which is the problem Power
  Query (Lesson 4.3) solves.

<!-- BEGIN GENERATED: nav -->
---

⬅️ **Previous:** [3.2 Data Validation & Conditional Formatting](../02-data-validation-conditional-formatting/README.md) · 🏠 [Course home](../../README.md) · **Next:** [3.4 PivotTables](../04-pivottables/README.md) ➡️
<!-- END GENERATED: nav -->
