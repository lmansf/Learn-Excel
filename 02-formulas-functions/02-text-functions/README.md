# Lesson 2.2 · Text Functions

> **Level:** Beginner → Intermediate · **Time:** about 50 minutes · **Workbook:** [`2.2-text-functions.xlsx`](2.2-text-functions.xlsx)
> **Data:** 400 patients from Bluestone's registration system (names, MRNs, phones, emails, chronic-condition flags), their 283 encounters in Q4 2025, the 147-provider roster, and a 650-row legacy registration export with messy addresses.

Hospital data is full of text that is *almost* right. The registration system exports names as `ABBOTT, Edward`, but the
patient letter needs `Edward Abbott`. A medical record number loses its leading zeros on its way through a spreadsheet. A
quality report needs ICD-10 codes rolled up to categories, and the appointment-reminder vendor rejects any phone number that
isn't exactly 10 digits. You could retype 400 rows by hand, but that takes an afternoon, invites typos, and has to be repeated
next month. Text functions fix every row at once, the same way every time. In this lesson you'll take apart, clean, and
rebuild realistic patient, encounter, and provider text from Bluestone Health System.

## What you'll learn

- Extract parts of text with LEFT, RIGHT, MID, FIND, and SEARCH
- Clean and standardize text with TRIM, CLEAN, UPPER, LOWER, PROPER, and SUBSTITUTE
- Join text with &, CONCAT, and TEXTJOIN; format numbers as text with TEXT
- Use modern TEXTBEFORE, TEXTAFTER, and TEXTSPLIT (Microsoft 365)

## 📖 Guide

All examples use the lesson workbook. **Patients** has 400 patients (MRN in column B, PatientName in C, Phone in D, Email in E,
ChronicConditions in F), **Encounters** has their Q4 2025 visits (AdmitDate in D, PrimaryDxCode in E, DxDescription in F),
and **Providers** lists every physician and advanced practice provider.

### 1. Text, numbers, and LEN

Every cell holds either a **number** (dates and times are numbers too) or **text**, which programmers call a *string*. Excel
treats them differently. You can add numbers but not text, and two values that look identical won't match if one is text and
the other is a number.

| Clue | Number | Text |
|---|---|---|
| Default alignment | Right | Left |
| `=ISNUMBER(cell)` | TRUE | FALSE |
| `=ISTEXT(cell)` | FALSE | TRUE |
| Lesson example | MRN in Patients!B3 shows `897724` | Phone in Patients!D2 shows `(555) 875-0698` |

**Identifiers** such as MRNs, ZIP codes, phone numbers, and ICD-10 codes should be stored as text. You never do math on them,
and storing them as numbers destroys leading zeros. That is exactly what happened to the MRN column on the Patients sheet: the
8-digit MRN `00897724` arrived as the number `897724`. You'll put the zeros back with TEXT in section 8.

Two rules follow from this:

- **Text functions always return text**, even when the result looks like a number. `=LEFT("45720-1280",5)` returns the text
  `45720`.
- **Text that looks like a number does not equal the number.** `="45501"=45501` returns FALSE. Section 8 shows how to convert
  in either direction.

The first function to learn is **LEN**, which returns the number of characters in a cell. It counts letters, digits,
punctuation, and spaces, including spaces you can't see.

```
=LEN(text)
```

| Formula | Result | Why |
|---|:-:|---|
| `=LEN(Patients!C2)` | 14 | `ABBOTT, Edward` is 6 letters + comma + space + 6 letters |
| `=LEN(Patients!C5)` | 18 | `hamilton, george` *looks* like 16 characters, but it ends with two invisible spaces |
| `=LEN(Patients!B3)` | 6 | Text functions accept numbers and work on their digits: `897724` has 6 |
| `=LEN(Encounters!D128)` | 5 | The date 11/14/2025 is stored as the serial number 45975 (see Lesson 2.3) |

> 💡 **Tip:** LEN is your X-ray for invisible characters. If `LEN(C5)` is larger than `LEN(TRIM(C5))`, the cell has stray
> spaces (section 4).

### 2. Taking text apart with LEFT, RIGHT, and MID

```
=LEFT(text, [num_chars])           the first num_chars characters (num_chars defaults to 1)
=RIGHT(text, [num_chars])          the last num_chars characters
=MID(text, start_num, num_chars)   num_chars characters, starting at position start_num
```

Positions count from 1. In `(555) 875-0698`, the opening parenthesis is position 1 and the first 5 is position 2.

| Formula | Result | Use |
|---|---|---|
| `=LEFT(Encounters!E3,3)` | `J44` | ICD-10 category of `J44.1` |
| `=RIGHT(Patients!D2,4)` | `0698` | Last 4 digits of the phone, for an identity check |
| `=MID(Patients!D2,2,3)` | `555` | Area code |
| `=MID(Patients!D2,7,3)` | `875` | The three digits after the area code |
| `=MID("ABBOTT, Edward",9,100)` | `Edward` | Asking for more characters than remain just returns the rest |

An **ICD-10 code** is a diagnosis code. Its first three characters are the **category**: a letter and two digits, such as
`J44` for chronic obstructive pulmonary disease (COPD). Everything after the dot adds detail. `J44.1` is COPD *with acute
exacerbation*, and in `S93.401A` (an ankle sprain) the final `A` means *initial encounter*. Codes have different lengths, but
the category is always the first three characters. That's why `=LEFT(E2,3)` safely rolls every code up to its category, so a
quality report can count `E11.65` and `E11.9` together as type 2 diabetes (`E11`).

LEFT, RIGHT, and MID with a typed number work only when the piece you want is always the same length, like an ICD-10 category
or the last 4 digits of a consistently formatted phone. When the length varies, as with a last name, you need FIND to locate a
landmark first.

> ⚠️ The results are text. `=RIGHT(Patients!D2,4)` returns the text `0698` and keeps its leading zero, which is correct for a
> phone fragment. If you typed `0698` into a cell yourself, Excel would store the number 698.

### 3. Finding positions with FIND and SEARCH

```
=FIND(find_text, within_text, [start_num])
=SEARCH(find_text, within_text, [start_num])
```

Both return the position where `find_text` first appears inside `within_text`. The optional `start_num` says where to start
looking (the default is 1). They differ in two ways:

| | FIND | SEARCH |
|---|---|---|
| Upper/lower case | **Case-sensitive** | Ignores case |
| Wildcards | No | Yes: `?` matches any one character, `*` any run of characters, `~` before a literal `?` or `*` |
| Not found | `#VALUE!` | `#VALUE!` |
| On Encounters!F2 (`Chest pain, unspecified`) | `=FIND("pain",F2)` → 7 <br> `=FIND("Pain",F2)` → `#VALUE!` | `=SEARCH("PAIN",F2)` → 7 |

FIND is most useful *inside* other functions. In a `LAST, First` name, the comma is a landmark: everything before it is the
last name. **Nesting** means using one function's result as another function's argument:

```
=LEFT(C2, FIND(",", C2) - 1)
```

Excel evaluates a nested formula from the inside out. For `ABBOTT, Edward` in Patients!C2:

| Step | Piece | Result |
|:-:|---|---|
| 1 | `FIND(",",C2)` | 7 (the comma is the 7th character) |
| 2 | `7 - 1` | 6 (the last name is the 6 characters before the comma) |
| 3 | `LEFT(C2,6)` | `ABBOTT` |

The first name starts just after the comma, so MID takes over:

```
=MID(C2, FIND(",", C2) + 2, LEN(C2))      → Edward
```

`+2` skips the comma and the space. `LEN(C2)` is simply a length that is always long enough to reach the end.

> ⚠️ `+2` assumes exactly one space after the comma, and the Patients sheet is messier than that. `CLARK,ARTHUR` (row 9) has no
> space, so `+2` cuts off the A and returns `RTHUR`. `Romero,  Albert` (row 6) has two spaces, so the result starts with a
> space. The robust version uses `+1` and lets TRIM remove however many spaces are there (section 4).

To find a **second occurrence**, start searching one character after the first. In `HTN;CKD;OA` (Patients!F12),
`=FIND(";",F12)` is 4, and `=FIND(";",F12,FIND(";",F12)+1)` is 8.

**The "contains" test.** SEARCH returns a number when it finds the text and `#VALUE!` when it doesn't. Wrap it in **ISNUMBER**
and you get TRUE or FALSE, which you can count or feed into IF (Lesson 2.1):

```
=ISNUMBER(SEARCH("DM", F2))                                  → TRUE for HTN;DM, FALSE for HTN or a blank cell
=IF(ISNUMBER(SEARCH("DM", F2)), "Invite to diabetes class", "")
```

> ⚠️ SEARCH finds text anywhere, even inside a longer word. The condition codes in this lesson never overlap, but if a list
> could contain both `HF` and `CHF`, `SEARCH("HF",…)` would match both. Wrap the list and the code in delimiters so only whole
> items match: `=ISNUMBER(SEARCH(";HF;", ";"&F2&";"))`.

### 4. Cleaning spaces and invisible characters: TRIM, CLEAN, and the non-breaking space

```
=TRIM(text)
=CLEAN(text)
```

**TRIM** removes spaces from the start and end of the text and shrinks every run of spaces inside it to a single space.
**CLEAN** removes non-printing characters (character codes 0 to 31), such as the line breaks that come along when you copy
from a multi-line text box in an EHR.

| Row | PatientName (`·` marks a space) | `=TRIM(C…)` | LEN before → after |
|:-:|---|---|:-:|
| 5 | `hamilton,·george··` | `hamilton, george` | 18 → 16 |
| 6 | `Romero,··Albert` | `Romero, Albert` | 15 → 14 |
| 8 | `·Chavez,·Carol` | `Chavez, Carol` | 14 → 13 |
| 14 | `ANDREWS,` + non-breaking space + `BRIAN` | unchanged! | 14 → 14 |

Row 14 is the trap. A **non-breaking space** looks exactly like a space, but it is a different character: code 160 instead of
32. Web pages and patient portals use it to keep words together, so it arrives whenever you copy text out of a browser.
**TRIM and CLEAN both ignore it.** To prove it's there, check the character after the comma: `=UNICODE(MID(C14,9,1))` returns
160, while an ordinary space returns 32.

The fix is to turn non-breaking spaces into ordinary spaces with SUBSTITUTE (section 6), then TRIM:

```
=TRIM(SUBSTITUTE(C14, UNICHAR(160), " "))      → ANDREWS, BRIAN
```

Combine all three for an all-purpose cleaner. Read it from the inside out: swap non-breaking spaces for normal ones, remove
non-printing characters, then trim.

```
=TRIM(CLEAN(SUBSTITUTE(C2, UNICHAR(160), " ")))
```

> 📋 **CHAR or UNICHAR?** Many websites show this fix with `CHAR(160)`. CHAR uses the operating system's character set, so
> `CHAR(160)` is a non-breaking space in Excel for Windows but can be a different character in Excel for Mac. **UNICHAR(160)**
> is the same character on every platform (Excel 2013 and later), so this lesson uses UNICHAR.

> 💡 **Tip:** To remove non-breaking spaces without a formula, open Find & Replace (Ctrl + H; Mac: ⌃ + H). In **Find what**, hold
> Alt and type 0160 on the numeric keypad (Mac: press Option + Space). Type one ordinary space in **Replace with**, then click
> **Replace All**.

### 5. Changing case with UPPER, LOWER, and PROPER

```
=UPPER(text)     =LOWER(text)     =PROPER(text)
```

| Formula (Patients!C5 is `hamilton, george  `) | Result |
|---|---|
| `=UPPER(TRIM(C5))` | `HAMILTON, GEORGE` |
| `=LOWER(TRIM(C5))` | `hamilton, george` |
| `=PROPER(TRIM(C5))` | `Hamilton, George` |

**PROPER** capitalizes the first letter of the text and every letter that follows a character that isn't a letter (a space,
hyphen, apostrophe, or digit). It lowercases everything else. That rule works for most names and fails in predictable places:

| Input | PROPER gives | Right? |
|---|---|---|
| `O'BRIEN, TERRY` (row 39) | `O'Brien, Terry` | Yes: the letter after the apostrophe is capitalized |
| `MCDONALD` (row 294) | `Mcdonald` | No: should be McDonald |
| `patient's chart` | `Patient'S Chart` | No: the s after an apostrophe is capitalized too |
| `MRI follow-up` | `Mri Follow-Up` | No: acronyms lose their capitals |
| `JOHN SMITH III` | `John Smith Iii` | No: Roman numerals |

Run PROPER, then patch the exceptions you know about with SUBSTITUTE: `=SUBSTITUTE(PROPER(A2),"Mcd","McD")`. For patient names,
confirm with the source system rather than guessing, because some families really do spell it Mcdonald.

**Comparing text.** The `=` operator ignores case: `="PETROV"="Petrov"` is TRUE. COUNTIF, MATCH, and the lookups in Lesson 2.6
ignore case as well. When case matters, use **EXACT**, which compares two texts character by character:

```
=EXACT(text1, text2)
=EXACT("PETROV", "Petrov")      → FALSE
```

The Check cells for practice task 13 and bonus B3 use EXACT, so `Hannah O'brien` won't pass.

### 6. Replacing text with SUBSTITUTE and REPLACE

```
=SUBSTITUTE(text, old_text, new_text, [instance_num])
=REPLACE(old_text, start_num, num_chars, new_text)
```

| | SUBSTITUTE | REPLACE |
|---|---|---|
| Finds what to change by | **Content**: what the text says | **Position**: where it sits |
| Changes | Every occurrence, or only the nth one with `instance_num` | One stretch of characters |
| Case-sensitive? | Yes: `SUBSTITUTE("Dm","DM","")` changes nothing | Not applicable |
| Typical use | Removing punctuation, swapping delimiters | Masking or overwriting a fixed position |

| Formula | Result | Why you'd do it |
|---|---|---|
| `=SUBSTITUTE(Encounters!E5,".","")` | `S93401A` | Electronic claim files list ICD-10 codes without the dot |
| `=SUBSTITUTE(Patients!F12,";",", ")` | `HTN, CKD, OA` | A readable problem list for a discharge summary |
| `=SUBSTITUTE(Patients!F12,";"," and ",2)` | `HTN;CKD and OA` | `instance_num` 2 changes only the second semicolon |
| `=REPLACE(TEXT(Patients!B3,"00000000"),1,4,"****")` | `****7724` | Masked MRN for a report that leaves the care team |

**Nesting SUBSTITUTE** removes several different characters in one formula. Each layer removes one character and hands its
result to the next layer out. For the phone in Patients!D2:

| Layer | Formula | Result |
|:-:|---|---|
| 1 | `SUBSTITUTE(D2,"(","")` | `555) 875-0698` |
| 2 | `SUBSTITUTE(layer 1,")","")` | `555 875-0698` |
| 3 | `SUBSTITUTE(layer 2," ","")` | `555875-0698` |
| 4 | `SUBSTITUTE(layer 3,"-","")` | `5558750698` |

```
=SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(D2,"(",""),")","")," ",""),"-","")
```

**Counting items with LEN and SUBSTITUTE.** Removing every semicolon shortens the text by exactly the number of semicolons, so
`=LEN(F12)-LEN(SUBSTITUTE(F12,";",""))` returns 2 for `HTN;CKD;OA`. A list always has one more item than separators, so add 1
to get 3 conditions.

> ⚠️ A blank cell has zero separators, so `+1` makes it look like one item. Guard with IF:
> `=IF(F2="",0,LEN(F2)-LEN(SUBSTITUTE(F2,";",""))+1)`.

**REPT** repeats text a given number of times: `=REPT(text, number_times)`. It's handy for padding, for example
`=REPT("0",8-LEN(B3))&B3` returns `00897724`, and for masks such as `=REPT("*",4)&RIGHT(B3,4)`.

### 7. Joining text with &, CONCAT, and TEXTJOIN

The **& operator** (ampersand) joins pieces of text. Any spaces or punctuation you want in the result go inside double quotes
as literal text. For the first provider on the Providers sheet:

```
="Dr. "&B2&" "&C2&", "&D2      → Dr. Roy Ferguson, MD
```

| Tool | Example | Accepts a range? | Delimiter | Version |
|---|---|:-:|---|---|
| `&` | `=B2&" "&C2` | No | Type each one | All versions |
| `CONCATENATE` | `=CONCATENATE(B2," ",C2)` | No | Type each one | All versions (kept for compatibility) |
| `CONCAT` | `=CONCAT(B2:D2)` | Yes | Type each one | Excel 2019 or later, Microsoft 365 |
| `TEXTJOIN` | `=TEXTJOIN(" ",TRUE,B2:C2)` | Yes | Once, placed between every item | Excel 2019 or later, Microsoft 365 |

`=CONCAT(B2:D2)` returns `RoyFergusonMD` because CONCAT adds nothing between items. **TEXTJOIN** solves that:

```
=TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...)
```

The **delimiter** goes between every pair of items, and **ignore_empty** decides what happens to blank cells. Patient PT10002
(row 3 of Patients) has a phone but no email:

| Formula | Result |
|---|---|
| `=TEXTJOIN(" \| ",TRUE,D3,E3)` | `(555) 373-1790` |
| `=TEXTJOIN(" \| ",FALSE,D3,E3)` | `(555) 373-1790 \| ` (dangling separator) |
| `=D3&" \| "&E3` | `(555) 373-1790 \| ` (same problem) |

> ⚠️ Joining uses the **stored value**, not what the cell displays. `="Admitted "&Encounters!D128` gives `Admitted 45975`,
> because the date 11/14/2025 is stored as 45975. Wrap numbers and dates in TEXT (next section) before you join them.

> 📋 CONCAT and TEXTJOIN need Excel 2019 or later, or Microsoft 365. Older versions show `#NAME?`, so use `&` if your workbook
> must open in them. A cell holds at most 32,767 characters, and TEXTJOIN returns `#VALUE!` if its result would be longer.

### 8. Turning numbers and dates into text with TEXT (and back again)

```
=TEXT(value, format_text)
```

**TEXT** converts a number or date into text, using the same format codes as custom number formats (Lesson 1.3). Put the format
code in double quotes.

| Formula | Result | Healthcare use |
|---|---|---|
| `=TEXT(Patients!B3,"00000000")` | `00897724` | 8-digit MRN with its leading zeros |
| `=TEXT(Encounters!D128,"mm/dd/yyyy")` | `11/14/2025` | A date inside a sentence |
| `=TEXT(Encounters!D128,"mmm-yyyy")` | `Nov-2025` | Month label on a quality report |
| `=TEXT(Encounters!D128,"mmmm d, yyyy")` | `November 14, 2025` | Patient letter |
| `=TEXT(Encounters!D128,"dddd")` | `Friday` | Day-of-week analysis |
| `=TEXT(0.91256,"0.0%")` | `91.3%` | Occupancy in a headline |
| `=TEXT(12345.678,"$#,##0.00")` | `$12,345.68` | Charges in a sentence |
| `=TEXT(TIME(14,5,0),"h:mm AM/PM")` | `2:05 PM` | Appointment time |

In a format code, each `0` is a required digit, so `"00000000"` pads with zeros to 8 digits. `mmm` is the short month name,
`mmmm` the full name, `dddd` the full weekday, and `yyyy` the 4-digit year. Now the join from section 7 works:

```
="Admitted "&TEXT(Encounters!D128,"mm/dd/yyyy")      → Admitted 11/14/2025
```

**Formatting versus TEXT.** Both can make an MRN *look* like `00897724`, but only TEXT changes the value:

| | Custom number format `00000000` | `=TEXT(B3,"00000000")` |
|---|---|---|
| The cell shows | `00897724` | `00897724` |
| The cell stores | the number 897724 | the text `00897724` |
| `="MRN "&cell` gives | `MRN 897724` | `MRN 00897724` |
| Math works? | Yes | No, it's text |

> ⚠️ TEXT results are text, so they sort alphabetically: `Dec-2025` sorts before `Nov-2025` and `Oct-2025`. Keep the real date
> in its own column for sorting and math, and use TEXT only for labels.

> ⚠️ Format codes depend on Excel's language. The codes in this course are for English versions. A German version, for
> example, uses `JJJJ` for the year. If a TEXT formula gives odd results on a colleague's computer, check the language first.

**Going the other way.** When a number arrives as text (from a CSV, a web form, or a text function), convert it before doing
math:

| Formula | Result | When to use it |
|---|---|---|
| `=VALUE("00897724")` | 897724 | Text digits to a number. The zeros disappear, as they should for a number |
| `=--"45501"` | 45501 | The same thing, shorter. Two minus signs (the *double unary*) force the conversion |
| `=NUMBERVALUE("1.234,5", ",", ".")` | 1234.5 | Text from a system that uses a comma as the decimal separator |

### 9. Modern splitting: TEXTBEFORE, TEXTAFTER, and TEXTSPLIT

> 📋 **Version note:** TEXTBEFORE, TEXTAFTER, and TEXTSPLIT need **Microsoft 365, Excel for the web, or Excel 2024** (Windows or
> Mac). In Excel 2021 and earlier they show `#NAME?`. If you share a workbook with someone on an older version, use the classic
> formulas from sections 2 to 6, or paste your results as values.

```
=TEXTBEFORE(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])
=TEXTAFTER(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])
=TEXTSPLIT(text, col_delimiter, [row_delimiter], [ignore_empty], [match_mode], [pad_with])
```

TEXTBEFORE returns everything before a **delimiter** (a landmark such as `","` or `"@"`), and TEXTAFTER returns everything after
it. Together they replace most FIND-inside-LEFT-or-MID nesting.

| Argument | Meaning | Default |
|---|---|---|
| `instance_num` | Which occurrence of the delimiter. **Negative numbers count from the end**: -1 is the last one | 1 |
| `match_mode` | 0 = case-sensitive, 1 = ignore case | 0 |
| `match_end` | 1 = treat the end of the text as a delimiter | 0 |
| `if_not_found` | What to return when the delimiter isn't there | `#N/A` |

| Goal | Classic formula | Microsoft 365 |
|---|---|---|
| Last name from `ABBOTT, Edward` | `=LEFT(C2,FIND(",",C2)-1)` | `=TEXTBEFORE(C2,",")` |
| First name | `=TRIM(MID(C2,FIND(",",C2)+1,LEN(C2)))` | `=TRIM(TEXTAFTER(C2,","))` |
| Portal username from an email | `=LEFT(E2,FIND("@",E2)-1)` | `=TEXTBEFORE(E2,"@")` |
| Email domain | `=MID(E2,FIND("@",E2)+1,LEN(E2))` | `=TEXTAFTER(E2,"@")` |
| Last condition in `HTN;CKD;OA` | `=TRIM(RIGHT(SUBSTITUTE(F12,";",REPT(" ",100)),100))` | `=TEXTAFTER(F12,";",-1)` |

The last row shows why the new functions matter. FIND only searches left to right, so finding the *last* delimiter takes a
trick: the classic formula swaps every semicolon for 100 spaces, grabs the last 100 characters (which now hold only the last
item plus padding), and trims the padding. A negative `instance_num` does the same job directly. You'll use it in the bonus,
where `TEXTAFTER(C2," ",-1)` grabs the ZIP code at the end of an address.

**When the delimiter is missing.** A patient with a single condition has no semicolon, so `=TEXTBEFORE("HTN",";")` returns
`#N/A`. Either of these fixes returns `HTN`:

```
=TEXTBEFORE(F2,";",,,1)        match_end = 1: the end of the text counts as a delimiter
=TEXTBEFORE(F2,";",,,,F2)      if_not_found: return the whole list when there's no semicolon
```

To skip optional arguments, leave them empty between commas, as above.

**TEXTSPLIT** breaks text into pieces and **spills** them: one formula fills as many cells as it needs. Type
`=TEXTSPLIT(Patients!F12,";")` in an empty cell and `HTN`, `CKD`, and `OA` land in three cells side by side. If any of those
cells already holds something, you get a `#SPILL!` error instead. Lesson 4.1 covers spilling in depth. To get a single value
instead of a spill, wrap TEXTSPLIT in another function:

| Formula | Result |
|---|---|
| `=COUNTA(TEXTSPLIT(Patients!F12,";"))` | 3 (how many conditions) |
| `=INDEX(TEXTSPLIT(Patients!F12,";"),2)` | `CKD` (the second one) |
| `=TEXTSPLIT("HTN;DM\|CKD;OA",";","\|")` | A 2 × 2 grid, because `\|` is the row delimiter |

> ⚠️ `COUNTA(TEXTSPLIT(F2,";"))` returns 1 for a **blank** cell, not 0, because TEXTSPLIT always returns at least one cell.
> Guard it the same way as the LEN trick: `=IF(F2="",0,COUNTA(TEXTSPLIT(F2,";")))`.

> 📋 Recent Microsoft 365 versions also include REGEXTEST, REGEXEXTRACT, and REGEXREPLACE, which match patterns such as "any run
> of digits." They're beyond this lesson. Everything here works with the functions above.

### 10. Worked example: cleaning one messy name, step by step

*Task: turn `·MCDONALD,··VIRGINIA` (Patients!C294, with a leading space and two spaces after the comma) into `Virginia McDonald`.*

| Step | Formula | Result |
|:-:|---|---|
| 1. Find the landmark | `=FIND(",",C294)` | 10 (the leading space is position 1) |
| 2. Last name | `=TRIM(LEFT(C294,FIND(",",C294)-1))` | `MCDONALD` |
| 3. First name | `=TRIM(MID(C294,FIND(",",C294)+1,LEN(C294)))` | `VIRGINIA` |
| 4. Join and fix case | `=PROPER(step 3&" "&step 2)` | `Virginia Mcdonald` |
| 5. Patch the exception | `=SUBSTITUTE(step 4,"Mcd","McD")` | `Virginia McDonald` |

Step 3 uses `+1` and TRIM instead of `+2`, so it works whether the name has zero, one, or two spaces after the comma. Nested
into one formula (without the step 5 patch):

```
=PROPER(TRIM(MID(C294,FIND(",",C294)+1,LEN(C294)))&" "&TRIM(LEFT(C294,FIND(",",C294)-1)))
```

In Microsoft 365 the same result is shorter:

```
=PROPER(TRIM(TEXTAFTER(C294,","))&" "&TRIM(TEXTBEFORE(C294,",")))
```

Then **check your work**. Sort or filter the new column, scan the shortest and longest values, compare LEN with what you
expect, and look specifically for Mc, Mac, O', and hyphenated names. Formulas are fast, but a wrong formula is wrong 400 times.

> 💡 **Tip:** Build long formulas in **helper columns** first, one column per step like the table above. When every step works,
> nest them into one formula, or keep the helper columns, which are easier for the next person to audit.

> 💡 **Tip:** A cleaned column still depends on the original. Before you delete the messy column, copy the cleaned one and use
> **Paste Special → Values** (Ctrl + Alt + V, then V; Mac: ⌃ + ⌘ + V) to replace the formulas with their results.

### 11. Working efficiently with text formulas

| Action | Windows | Mac |
|---|---|---|
| Edit the active cell | F2 | ⌃ + U |
| Expand the formula bar for a long formula | Ctrl + Shift + U | ⌃ + Shift + U |
| Fill the top cell's formula down a selected range | Ctrl + D | ⌘ + D |
| Show formulas instead of results | Ctrl + `` ` `` (grave accent) | ⌃ + `` ` `` |
| Paste Special (for example, to paste values) | Ctrl + Alt + V | ⌃ + ⌘ + V |
| Find & Replace | Ctrl + H | ⌃ + H |

Double-clicking the fill handle copies a formula down to the last row of data on both platforms. The lesson's data sheets are
**Excel Tables**, so a formula typed in the first row of a yellow column usually fills the whole column on its own (Lesson 3.1
explains Tables). On Windows, **Formulas → Evaluate Formula** steps through a nested formula one layer at a time, which is the
fastest way to see where a long text formula goes wrong.

Formulas aren't the only way to reshape text. Choose the tool by how often you'll repeat the job:

| Tool | How | Updates when the data changes? | Best for |
|---|---|:-:|---|
| Text formulas (this lesson) | `=TRIM(MID(…))` | Yes | Monthly reports and anything you need to audit |
| Flash Fill | Type an example or two, then **Data → Flash Fill** (Windows: Ctrl + E) | No | Quick one-off fixes (Lesson 1.2) |
| Text to Columns | **Data → Text to Columns** | No | Splitting a whole column once at a delimiter (Lesson 3.3) |
| Power Query | **Data → Get Data** | Yes, on refresh | Cleaning the same export every month (Lesson 4.3) |

## 🧪 Hands-on practice

Download [`2.2-text-functions.xlsx`](2.2-text-functions.xlsx) and open the **Practice** sheet. Type each answer in its yellow
cell, as a formula wherever possible, and the **Check** column turns green when you're right. For the column tasks (gray answer
cells), you fill a yellow column on a data sheet and the gray cell summarizes your work.

<!-- BEGIN GENERATED: practice -->
Tasks use the Patients, Encounters, and Providers sheets. When a task names one patient, it also gives the row, so you can point your formula at that row's cells (for example Patients!C25). Column tasks ask you to fill a yellow column on a data sheet; a gray cell here then summarizes your column.

| # | Task | Hint |
|:-:|------|------|
| 1 | Front-desk staff confirm a caller's identity with the last 4 digits of their phone number. Return the last 4 digits of patient PT10015's phone (row 16 of the Patients sheet). | RIGHT(text, num_chars) |
| 2 | On the Encounters sheet, fill the yellow DxCategory column with the ICD-10 category of each PrimaryDxCode (its first 3 characters). The gray cell counts encounters in category E11 (type 2 diabetes). Type your formula in G2, then copy it down. | LEFT(text, num_chars) |
| 3 | Extract the last name of patient PT10031 (row 32 of the Patients sheet, "PETROV, Ana"): everything before the comma. | FIND gives the comma's position; LEFT takes everything before it |
| 4 | The appointment-reminder system needs phone numbers as 10 digits with no punctuation. Convert patient PT10058's phone (row 59, "(555) 529-7071") to digits only. | Nest one SUBSTITUTE per character you want to remove |
| 5 | Patient PT10002's MRN (row 3) shows 897724 because the export stored it as a number and dropped the leading zeros. Return it as the 8-character text MRN printed on wristbands. | TEXT(value, format_text) with a format of eight 0s |
| 6 | Monthly quality reports label encounters by month. Return the admit month of encounter ENC120217 (row 128 of the Encounters sheet) as text in the form Mon-YYYY, for example Jan-2026. | TEXT with a date format code |
| 7 | Build the ID-badge label for provider PRV1054 (row 55 of the Providers sheet) in the form Dr. First Last, Credential. For example, provider PRV1001's label is Dr. Roy Ferguson, MD. | Join pieces with &; spaces and punctuation go inside quotes |
| 8 | On the Patients sheet, fill the yellow HasDM column with TRUE when the patient's ChronicConditions list includes DM (diabetes) and FALSE otherwise. Blank lists should give FALSE. The gray cell counts the TRUEs. | SEARCH returns a position or #VALUE!; ISNUMBER turns that into TRUE/FALSE |
| 9 | The patient portal username is the part of the email address before the @. Return the username for patient PT10123 (row 124). | TEXTBEFORE (Microsoft 365), or LEFT + FIND |
| 10 | How many chronic conditions does patient PT10101 (row 102) have? Split the semicolon-separated ChronicConditions list with TEXTSPLIT and count the pieces. | COUNTA(TEXTSPLIT(…)) returns one number instead of a spill |
| 11 | On the Patients sheet, fill the yellow ContactLine column with the phone and email joined by " \| " (space, vertical bar, space), for example (555) 875-0698 \| edward.abbott94@example.com. When there's no email, show just the phone, with no dangling separator. The gray cell counts lines that contain a \|. | TEXTJOIN(delimiter, ignore_empty, …) |
| 12 | On the Patients sheet, fill the yellow FirstName column with each patient's first name: everything after the comma in PatientName, with no extra spaces. Watch out: some names have doubled, leading, or trailing spaces, and 28 were pasted from the patient portal with non-breaking spaces (UNICHAR(160)). Case doesn't matter here. The gray cell adds up the lengths of all your first names, so any leftover space changes the total. | TRIM can't remove a non-breaking space; SUBSTITUTE it with a normal space first |
| 13 | Patient PT10325's name was typed as "o'brien,hannah  " (row 326, with two trailing spaces). Return a clean display name in the form First Last, in Proper Case. This check is case-sensitive. | Extract both parts, join them with " ", then wrap everything in PROPER |
<!-- END GENERATED: practice -->

## ✅ Answer key

The workbook has a hidden **Answer Key** sheet (right-click any sheet tab → **Unhide…** → *Answer Key*). Its *Live result*
column runs each sample formula, so you can see it working. The same answers are below, collapsed so you don't see them by
accident.

<!-- BEGIN GENERATED: answers -->
<details>
<summary><b>🔑 Show the answer key</b> — Try every task before opening this.</summary>

**1. Front-desk staff confirm a caller's identity with the last 4 digits of their phone…**

- **Answer:** 7258
- **Solution:** `=RIGHT(Patients!D16,4)`

RIGHT returns characters from the end of the text. The phone is stored as text in a fixed (555) 123-4567 layout, so its last 4 characters are always the last 4 digits. The result is text: a phone ending in 0698 would correctly keep its leading 0.

**2. DxCategory column (encounters in category E11)**

- **Answer:** 15
- **Solution:** `=LEFT(E2,3)`

An ICD-10 code's first three characters are its category, so LEFT(code,3) groups E11.65 and E11.9 together as E11. Codes have different lengths (I10, E11.65, S72.001A), but the category is always the first 3 characters, which is why LEFT works with a fixed count here.

**3. Extract the last name of patient PT10031 (row 32 of the Patients sheet, "PETROV,…**

- **Answer:** PETROV
- **Solution:** `=LEFT(Patients!C32,FIND(",",Patients!C32)-1)`

FIND(",",…) returns the comma's position (7). The last name is the 6 characters before it, so subtract 1. In Microsoft 365 you can also write =TEXTBEFORE(Patients!C32,",").

**4. The appointment-reminder system needs phone numbers as 10 digits with no punctuation.…**

- **Answer:** 5555297071
- **Solution:** `=SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(Patients!D59,"(",""),")","")," ",""),"-","")`

Each SUBSTITUTE replaces one character with nothing (""). Nesting them removes the parentheses, the space, and the dash in one formula. Work from the inside out: the innermost SUBSTITUTE runs first and passes its result to the next one. The result is text, which is what you want for phone numbers.

**5. Patient PT10002's MRN (row 3) shows 897724 because the export stored it as a number…**

- **Answer:** 00897724
- **Solution:** `=TEXT(Patients!B3,"00000000")`

In a format code, each 0 is a required digit, so "00000000" pads the number with leading zeros to 8 digits. The result "00897724" is text, so the zeros are part of the value: they survive joining with & and match MRNs stored as text in other systems. A custom number format (Lesson 1.3) only changes how the cell looks, so ="MRN "&Patients!B3 would still give MRN 897724.

**6. Monthly quality reports label encounters by month. Return the admit month of encounter…**

- **Answer:** Nov-2025
- **Solution:** `=TEXT(Encounters!D128,"mmm-yyyy")`

In a date format code, mmm is the short month name and yyyy is the 4-digit year. The dash is copied as-is. Without TEXT, a date joined to text shows its serial number: ="Admitted "&Encounters!D128 gives Admitted 45975, because a date is stored as a number.

**7. Build the ID-badge label for provider PRV1054 (row 55 of the Providers sheet) in the…**

- **Answer:** Dr. Jonathan Taylor, MD
- **Solution:** `="Dr. "&Providers!B55&" "&Providers!C55&", "&Providers!D55`

The & operator joins text. Literal text, including every space, the period after Dr, and the comma, goes inside double quotes. =CONCAT("Dr. ",Providers!B55," ",Providers!C55,", ",Providers!D55) gives the same result.

**8. HasDM column (patients with diabetes)**

- **Answer:** 52
- **Solution:** `=ISNUMBER(SEARCH("DM",F2))`

SEARCH returns the position where DM starts, or #VALUE! when it isn't there (including in blank cells). ISNUMBER converts any position to TRUE and the error to FALSE. SEARCH ignores case, so it would also find dm. If a code could hide inside a longer code, search for ";DM;" inside ";"&F2&";" instead.

**9. The patient portal username is the part of the email address before the @. Return the…**

- **Answer:** sara.scott36
- **Solution:** `=TEXTBEFORE(Patients!E124,"@")`

TEXTBEFORE returns everything before the first @. The classic version, which works in every Excel, is =LEFT(Patients!E124,FIND("@",Patients!E124)-1).

**10. How many chronic conditions does patient PT10101 (row 102) have? Split the…**

- **Answer:** 4
- **Solution:** `=COUNTA(TEXTSPLIT(Patients!F102,";"))`

TEXTSPLIT spills one condition per cell to the right. Wrapping it in COUNTA counts the pieces and returns a single number, so nothing spills into the Check column. Without TEXTSPLIT, count the semicolons and add 1: =LEN(Patients!F102)-LEN(SUBSTITUTE(Patients!F102,";",""))+1.

**11. ContactLine column (lines with both phone and email)**

- **Answer:** 229
- **Solution:** `=TEXTJOIN(" | ",TRUE,D2,E2)`

TEXTJOIN puts the delimiter between items. With ignore_empty set to TRUE, it skips the blank email, so those rows show only the phone. With FALSE, or with D2&" | "&E2, every row gets a separator and the blank rows end in a dangling " | ".

**12. FirstName column (checksum: total characters)**

- **Answer:** 2,324
- **Solution:** `=TRIM(MID(SUBSTITUTE(C2,UNICHAR(160)," "),FIND(",",C2)+1,LEN(C2)))`

Work from the inside out. SUBSTITUTE turns each non-breaking space into an ordinary space, MID takes everything after the comma (LEN is simply a length that's long enough), and TRIM removes the leading, trailing, and doubled spaces. Without the SUBSTITUTE, the 28 portal rows keep an invisible character and the total comes out 28 too high. On Windows, CHAR(160) is the same character as UNICHAR(160). A Microsoft 365 version is =TRIM(SUBSTITUTE(TEXTAFTER(C2,","),UNICHAR(160)," ")).

**13. Patient PT10325's name was typed as "o'brien,hannah  " (row 326, with two trailing…**

- **Answer:** Hannah O'Brien
- **Solution:**

```
=PROPER(TRIM(MID(Patients!C326,FIND(",",Patients!C326)+1,LEN(Patients!C326)))&" "&LEFT(Patients!C326,FIND(",",Patients!C326)-1))
```


MID + FIND takes the first name and TRIM drops the trailing spaces. LEFT + FIND takes the last name. Join them with a space and wrap the whole thing in PROPER, which capitalizes the first letter of each word and every letter that follows a non-letter. That's why the B after the apostrophe in O'Brien comes out right. In Microsoft 365: =PROPER(TRIM(TEXTAFTER(Patients!C326,","))&" "&TEXTBEFORE(Patients!C326,",")).

</details>
<!-- END GENERATED: answers -->

## 🏆 Bonus challenge

<!-- BEGIN GENERATED: bonus -->
Bluestone is opening a diabetes education clinic in Cedar Ridge. Marketing wants a clean mailing list from the legacy registration export (the Registrations sheet), and care management wants to know how many current patients are 'complex' (three or more chronic conditions). The export crams city, state, and ZIP into one CityStateZip column in several styles: Cedar Ridge, OH 45720 · CEDAR RIDGE, OH 45720 · Cedar Ridge OH 45720 (no comma) · Cedar Ridge, OH 45720-1280 (ZIP+4). Some city names are two words. (The export also has duplicate records. Removing those is a Lesson 3.3 job, so leave them in.)

Work on the **Bonus** sheet of the workbook.

- **B1.** On the Registrations sheet, fill the yellow City column with just the city name in Proper Case, with no comma, state, or ZIP (for example Lakeview Heights). The gray cell counts rows whose City is exactly Lakeview Heights (case-sensitive). That town appears in every messy style, so it's a good test. *(Hint: The city is everything before the second-to-last space. TEXTBEFORE accepts a negative instance_num)*
- **B2.** Fill the yellow ZIP5 column with the 5-digit ZIP code as text. ZIP+4 codes like 45720-1280 must become 45720. The gray cell counts rows in ZIP 45501 (downtown Bluestone). *(Hint: The ZIP is the last word; keep only its first 5 characters)*
- **B3.** Write one formula that turns record R1002's CityStateZip (row 3, "CEDAR RIDGE, OH 45720") into the mailing-label line City, ST 12345: city in Proper Case, a comma and a space, the 2-letter state in capitals, a space, and the 5-digit ZIP. This check is case-sensitive. *(Hint: Build city, state, and ZIP separately, then join them with &. Which piece needs PROPER?)*
- **B4.** Back on the Patients sheet, fill the yellow ConditionCount column with the number of chronic conditions each patient has: 0 when ChronicConditions is blank, 1 for HTN, 2 for HTN;DM, and so on. The gray cell counts 'complex' patients with 3 or more conditions. *(Hint: Items = semicolons + 1. How many semicolons? Compare LEN before and after removing them)*
- **B5.** What is the average number of chronic conditions per patient across all 400 patients, including those with none? Use your ConditionCount column. (The check accepts 2 decimal places.) *(Hint: AVERAGE of the column you just filled)*
<!-- END GENERATED: bonus -->

<!-- BEGIN GENERATED: bonus-answers -->
<details>
<summary><b>🔑 Show the bonus solution</b> — Give it a real try first!</summary>

**B1. City column (rows exactly 'Lakeview Heights')**

- **Answer:** 51
- **Solution:** `=PROPER(TRIM(SUBSTITUTE(TEXTBEFORE(C2," ",-2),",","")))`

The commas are unreliable, but the spaces are not: the last two spaces always separate the city, the state, and the ZIP. TEXTBEFORE(C2," ",-2) counts spaces from the end, so it returns everything before the second-to-last space ("Cedar Ridge," or "CEDAR RIDGE"). SUBSTITUTE removes a comma if there is one, TRIM tidies up, and PROPER fixes the case. A FIND(",") approach fails with #VALUE! on the rows that have no comma. Without Microsoft 365, the classic trick is =PROPER(SUBSTITUTE(LEFT(C2,LEN(C2)-LEN(TRIM(RIGHT(SUBSTITUTE(C2," ",REPT(" ",100)),100)))-4),",","")): the TRIM(RIGHT(SUBSTITUTE(…))) part pulls out the last word (the ZIP), and LEFT keeps everything except the ZIP and the 4 characters of " OH ".

**B2. ZIP5 column (rows in ZIP 45501)**

- **Answer:** 113
- **Solution:** `=LEFT(TEXTAFTER(C2," ",-1),5)`

TEXTAFTER(C2," ",-1) returns everything after the last space: the whole ZIP or ZIP+4. LEFT(…,5) keeps the first five digits. RIGHT(C2,5) looks tempting but returns "-1280" for 45720-1280. 23 of the 113 rows in 45501 have a ZIP+4. Keep ZIPs as text: a ZIP like 02134 would lose its leading zero as a number. Classic version: =LEFT(TRIM(RIGHT(SUBSTITUTE(C2," ",REPT(" ",100)),100)),5). SUBSTITUTE swaps every space for 100 spaces, RIGHT(…,100) then grabs a chunk that holds only the last word plus padding, and TRIM strips the padding.

**B3. Write one formula that turns record R1002's CityStateZip (row 3, "CEDAR RIDGE, OH…**

- **Answer:** Cedar Ridge, OH 45720
- **Solution:**

```
=PROPER(TRIM(SUBSTITUTE(TEXTBEFORE(Registrations!C3," ",-2),",","")))&", "&UPPER(TEXTBEFORE(TEXTAFTER(Registrations!C3," ",-2)," "))&" "&LEFT(TEXTAFTER(Registrations!C3," ",-1),5)
```


The state is the word between the last two spaces: TEXTAFTER(…," ",-2) gives "OH 45720", and TEXTBEFORE(…," ") keeps "OH". Apply PROPER to the city only. Wrapping the whole line in PROPER is the classic mistake: it turns OH into Oh. If you finished B1 and B2, =Registrations!D3&", OH "&Registrations!E3 also works, but extracting the state keeps the formula correct for out-of-state patients.

**B4. ConditionCount column (patients with 3+ conditions)**

- **Answer:** 35
- **Solution:** `=IF(F2="",0,LEN(F2)-LEN(SUBSTITUTE(F2,";",""))+1)`

LEN(F2)-LEN(SUBSTITUTE(F2,";","")) is the number of semicolons, because removing them shortens the text by exactly that many characters. A list always has one more item than separators, so add 1. The IF handles blank lists: without it, a blank cell has 0 semicolons and would count as 1 condition. COUNTA(TEXTSPLIT(F2,";")) has the same problem on blank cells, so it needs the same IF.

**B5. What is the average number of chronic conditions per patient across all 400 patients,…**

- **Answer:** 0.88
- **Solution:** `=AVERAGE(Patients!J2:J401)`

If your average comes out as 1.33, your column counts the 181 blank lists as 1 condition each. That mistake doesn't change B4's count, because those rows stay below 3, but it inflates the average. Testing a formula on the edge cases (blank, one item, many items) catches errors like this.

</details>
<!-- END GENERATED: bonus-answers -->

## Key takeaways

- Store identifiers such as MRNs, ZIP codes, phone numbers, and ICD-10 codes as **text**. Text functions always return text, and
  VALUE or `--` converts back when you really need a number.
- **LEFT, RIGHT, and MID** take fixed-length pieces. When the length varies, **FIND** (case-sensitive) or **SEARCH**
  (case-insensitive, wildcards) locates a landmark first, and `ISNUMBER(SEARCH(…))` is the standard "contains" test.
- **LEN** is your X-ray. **TRIM** removes ordinary spaces and **CLEAN** removes non-printing characters, but neither touches the
  non-breaking space, so run `SUBSTITUTE(…,UNICHAR(160)," ")` first.
- **PROPER** fixes most names but not McDonald, acronyms, or Roman numerals. Use **EXACT** when case matters, because `=` and
  COUNTIF ignore it.
- Join with **&**, **CONCAT**, or **TEXTJOIN** (whose ignore_empty argument skips blanks), and wrap numbers and dates in **TEXT**
  so they keep their format inside a sentence.
- In Microsoft 365, **TEXTBEFORE**, **TEXTAFTER**, and **TEXTSPLIT** replace most FIND/MID nesting, and a negative instance_num
  searches from the end.

<!-- BEGIN GENERATED: nav -->
---

⬅️ **Previous:** [2.1 Logical Functions: IF, AND, OR, IFS & More](../01-logical-functions/README.md) · 🏠 [Course home](../../README.md) · **Next:** [2.3 Dates & Times](../03-date-time-functions/README.md) ➡️
<!-- END GENERATED: nav -->
