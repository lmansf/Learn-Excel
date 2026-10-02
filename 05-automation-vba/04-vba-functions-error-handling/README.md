# Lesson 5.4 · VBA: Custom Functions, Dictionaries & Error Handling

> **Level:** Expert · **Time:** about 185 minutes · **Workbook:** [`5.4-vba-functions-error-handling.xlsx`](5.4-vba-functions-error-handling.xlsx)
> **Data:** Cedar Ridge Medical Center, 2025: adult patients with height and weight, inpatient and observation stays, every claim with its payer and status, and the 8 payers. A few rows contain deliberate data-entry errors for your code to catch.

Analysts at Cedar Ridge Medical Center calculate the same things over and over: a patient's BMI, an age on the
admission date, a length of stay, a denial rate. As worksheet formulas, each rule gets copied into dozens of workbooks,
and every copy can drift. A **user-defined function** (UDF) keeps the rule in one place and lets anyone call it from a
cell, like `=LOSDAYS(E2,F2)`. In this lesson you'll write four of them, summarize a year of claims in a single pass
with Collections and dictionaries, and make a monthly macro survive the things that break real reports: a renamed sheet,
a blank amount, and a discharge date keyed in the wrong year.

## What you'll learn

- Write user-defined functions (UDFs) you can call from cells
- Pass arguments ByVal/ByRef, use Optional arguments, and return errors with CVErr
- Count and group with Collections and Scripting.Dictionary
- Handle errors gracefully with On Error, the Err object, and cleanup code

## 📖 Guide

### 1. Set up the workbook

You need the workbook and the four starter modules in this lesson's [`starter/`](starter/) folder. On GitHub, open each
`.bas` file and click **Download raw file**.

The lesson workbook is an .xlsx file, and .xlsx files can't store macros. Do this once before you write any code:

1. Open `5.4-vba-functions-error-handling.xlsx`, choose **File → Save As**, and pick **Excel Macro-Enabled Workbook
   (\*.xlsm)** as the file type.
2. Open the **Visual Basic Editor** (VBE) with **Alt + F11** (Mac: **Option + F11**), or click **Developer → Visual
   Basic**.
3. Choose **File → Import File…** (Windows shortcut: **Ctrl + M**) and import the four starter files one at a time. They
   appear in the **Modules** folder of the Project Explorer.
4. Choose **Debug → Compile VBAProject**. If nothing happens, the code compiles.

| Starter module | What's in it | Tasks |
|---|---|---|
| [`HealthUDFs.bas`](starter/HealthUDFs.bas) | Stubs for BMI, AGEAT, LOSDAYS, and DENIALRATE, plus two finished helpers | 1–6 |
| [`ClaimDictionaries.bas`](starter/ClaimDictionaries.bas) | Stubs for three Collection and dictionary macros and the AddIfNew helper, plus a finished HeaderColumn helper | 8–10 |
| [`Snippets.bas`](starter/Snippets.bas) | Four finished snippets for the predict-the-output tasks | 7, 11–13 |
| [`PayerSummary.bas`](starter/PayerSummary.bas) | A skeleton for the bonus macro, with finished helpers | Bonus |

The workbook has two sheets you haven't met before. Your macros write their results to the **Output** sheet, and the
gray cells on the Practice sheet read them from there. The **Snippets** sheet shows the code for the predict-the-output
tasks.

> ⚠️ If Excel blocks the macros in a file you downloaded, the file carries the *Mark of the Web*. Lesson 5.1 shows how
> to unblock it safely (on Windows: right-click the file → **Properties** → tick **Unblock**).

> 📋 The [`solutions/`](solutions/) folder has complete answers (spoilers). The procedures in a solution module have the
> same names as the ones in your starter module. Before you import a solution, remove your own module (right-click it in
> the Project Explorer → **Remove**) or import it into a spare copy of the workbook. Otherwise every call to those names
> is ambiguous: a cell or button can run the wrong copy, and code that calls one of them stops with *Ambiguous name
> detected*.

### 2. Functions vs. Subs

Your macros so far have been **Sub procedures**. A Sub *does* something, like writing values or formatting a sheet. A
**Function procedure** *calculates* something and hands the result back to whoever called it.

| | Sub | Function |
|---|---|---|
| Hands back a value | No | Yes: assign it to the function's own name |
| Run it from the Macros list (Alt + F8; Mac: Option + F8) or a button | Yes, if it takes no arguments | No: functions never appear in the Macros list |
| Call it from a worksheet cell | No | Yes, if it's Public and in a standard module |
| Call it from other VBA code | `BuildPayerSummary` | `x = BMI(197.9, 70.4)` |
| Change cells, sheets, or settings | Yes | Only when VBA calls it, never when a cell calls it |

```vba
[Public | Private] Function Name([arg1 As Type, arg2 As Type, ...]) As ReturnType
    ' ...calculate...
    Name = result          ' assigning to the function's own name sets the return value
End Function
```

Here is the smallest useful UDF:

```vba
Public Function BMI(weightLb As Double, heightIn As Double) As Double
    BMI = 703 * weightLb / heightIn ^ 2
End Function
```

To try it, don't paste it into a new module. HealthUDFs already has a BMI stub, and two public functions named BMI in
one workbook make `=BMI(` ambiguous, so a cell can end up running the empty stub. Instead, add the line
`BMI = 703 * w / h ^ 2` to the stub, just above `End Function`. The stub has already put the weight in `w` and the
height in `h`, and task 1 adds the error checks later.

**Worked example.** Albert Romero (PT10005, row 3 of the Patients sheet) weighs 197.9 lb and is 70.4 in tall. Type
`=BMI(Patients!G3,Patients!F3)` in an empty cell on the Practice sheet and Excel shows about 28.07 (how many decimals
you see depends on the column width). In the VBE's **Immediate window** (Ctrl + G, or **View → Immediate
Window** on a Mac), type `? BMI(197.9, 70.4)` and press Enter to see `28.0708653473657`. The `?` is short for
`Debug.Print`, so the Immediate window is the fastest way to test a function.

> ⚠️ If a function never assigns to its own name, it returns the default for its type (0, "", or Empty) without any
> error. A cell that shows 0 when you expected a number usually means a branch of your code forgot the assignment.

Use `Exit Function` to leave early, for example right after you return an error value.

### 3. The rules for user-defined functions

1. **Put it in a standard module.** Use **Insert → Module** in the VBE. A function inside a sheet module (Sheet1) or
   ThisWorkbook can't be called from a cell, so the cell shows #NAME?.
2. **Make it Public** (the default). A **Private** function doesn't appear in AutoComplete or the Insert Function list,
   and code in other modules can't call it, which is exactly what you want for helpers. Private hides a function but
   doesn't lock it: a formula that spells out its exact name can still run it.
3. **Don't name it like a cell address.** `LOS1` and `BMI2` are cell addresses (column LOS, row 1), so Excel reads them
   as references and the formula fails. That's why this lesson uses `LOSDAYS`.
4. **Return a value and change nothing else.** A UDF that a cell calls can't change other cells, formats, sheets, or
   Excel settings. If it tries, the attempt fails and the cell usually shows #VALUE!. Put that kind of work in a Sub.
5. **Expect silence when it fails.** A run-time error inside a UDF doesn't show the usual error dialog. The cell just
   shows #VALUE!. Never use MsgBox in a UDF either, because Excel may call it many times during one recalculation.

Your UDFs show up in Excel's AutoComplete list as you type `=BM`, and in the **User Defined** category of the Insert
Function dialog (click **fx** next to the formula bar). On Windows, type `=BMI` and press **Ctrl + Shift + A** to
insert the parentheses and argument names.

**When does a UDF recalculate?** Excel tracks the cells you pass in as arguments, and nothing else.

| Situation | What Excel does |
|---|---|
| A cell passed as an argument changes | Recalculates the UDF |
| A cell the code reads directly (not an argument) changes | Nothing, so the result goes stale |
| The function calls `Application.Volatile` | Recalculates it whenever anything recalculates, like TODAY and NOW |
| You edit the VBA code | Nothing yet. Re-enter the cell: F2, then Enter (Mac: ⌃ + U, then Return) |
| You press Ctrl + Alt + F9 (Windows) | Recalculates every formula in every open workbook |

To refresh a whole column of UDF cells after you fix your code, on Windows or on a Mac, select the cells, open
**Replace** (Ctrl + H; Mac: ⌃ + H), and replace `=` with `=`. Excel re-enters every formula in the selection, so each
cell runs your new code.

> ⚠️ Pass every input as an argument. A function that reads `Range("ReportDate")` inside its code won't update when
> ReportDate changes, because Excel doesn't know the function depends on that cell.

> 💡 **Tip:** To debug a UDF, click a line inside it and press **F9** (Mac: **Debug → Toggle Breakpoint**) to set a
> breakpoint. Then select a cell that uses the function and re-enter it (F2, then Enter; Mac: ⌃ + U, then Return). The
> VBE stops at the breakpoint, and you step through with F8 (Mac: **Debug → Step Into**) as in Lesson 5.2. Remove the
> breakpoint when you're done, or every recalculation stops there.

> 📋 **Version note:** UDFs run in Excel for Windows and Excel for Mac. Excel for the web, iPad, and iPhone can't run
> VBA, so UDF cells can't recalculate there. If you need a reusable formula that works everywhere, a named LAMBDA
> (Lesson 4.2, which needs Microsoft 365 or Excel 2024) is the macro-free alternative for pure calculations.

> 📋 Each UDF call has some overhead. A few thousand UDF cells are fine. Hundreds of thousands can make a workbook
> sluggish, and a built-in formula is usually faster when one can do the job.

### 4. Passing arguments: ByRef, ByVal, Optional, and ParamArray

VBA passes arguments **ByRef** (by reference) unless you say otherwise. The procedure receives the caller's variable
itself, so it can change it. **ByVal** (by value) passes a copy, so the caller's variable is safe.

```vba
Sub AddLateFee(ByRef balance As Double)
    balance = balance + 35              ' changes the caller's variable
End Sub

Sub PreviewLateFee(ByVal balance As Double)
    balance = balance + 35              ' changes only the local copy
End Sub

Sub Demo()
    Dim due As Double
    due = 200
    PreviewLateFee due                  ' due is still 200
    AddLateFee due                      ' due is now 235
    Debug.Print due                     ' 235
End Sub
```

How you *call* a procedure also matters:

| Call | What a ByRef parameter receives | Can the caller's variable change? |
|---|---|---|
| `AddLateFee due` | the variable | Yes |
| `Call AddLateFee(due)` | the variable | Yes |
| `AddLateFee (due)` | a temporary copy, because the parentheses turn `due` into an expression | No |
| `x = SomeFunction(due)` | the variable (these parentheses belong to the function call) | Yes |

> ⚠️ Call a Sub without parentheses: `AddLateFee due`. If you type `AddLateFee(due)`, the VBE quietly reformats it as
> `AddLateFee (due)`, and VBA passes a copy. A Sub with two arguments won't even accept parentheses unless you use
> `Call`.

> 💡 **Tip:** Write `ByVal` for inputs a procedure shouldn't change. It documents your intent and prevents accidents.
> When a *cell* calls a UDF, ByRef and ByVal make no visible difference, because Excel passes values and ranges, not
> your variables.

**Optional arguments** let the caller leave an input out. They must come after all the required arguments.

| Style | Declaration | How you tell it was left out |
|---|---|---|
| Variant, no default | `Optional asOf As Variant` | `IsMissing(asOf)` is True |
| Typed, with a default | `Optional statusToCount As String = "Denied"` | You can't, and you don't need to: it holds the default |

`IsMissing` works only on an Optional **Variant**. On a typed Optional argument it is always False, and a left-out
argument just holds the type's default (0, "", or False). Excel's documentation marks optional arguments with square
brackets, so this lesson writes AGEAT as `AGEAT(dob, [asOf])`.

```vba
Public Function AGEAT(dob As Variant, Optional asOf As Variant) As Variant
    Dim refValue As Variant
    If IsMissing(asOf) Then
        refValue = Date             ' left out: use today's date
        Application.Volatile        ' today changes, so recalculate every time
    Else
        refValue = ArgValue(asOf)   ' ArgValue is a helper from the starter module (section 5)
    End If
    ' ...validate, then calculate the age...
```

With this design, `=AGEAT(D2)` gives the patient's age today and `=AGEAT(D2,E2)` gives the age on the admission date.

**Age in completed years.** `Year(asOf) - Year(dob)` counts calendar years, so it's one too many whenever the birthday
comes later in the year than asOf. Build this year's birthday with `DateSerial` and compare:

```vba
years = Year(refDate) - Year(birth)
If DateSerial(Year(refDate), Month(birth), Day(birth)) > refDate Then years = years - 1
```

**Worked example.** Encounter ENC111079 (row 14 of the Encounters sheet) is for a patient born 10/09/1961 and admitted
on 01/10/2025. 2025 − 1961 = 64, but the 2025 birthday (October 9) hadn't happened yet on January 10, so the age at
admission is 63. `DateDiff("yyyy", dob, asOf)` also returns 64, because it counts year boundaries crossed, not
birthdays.

**ParamArray** (briefly) accepts any number of arguments as one Variant array. It must be the last argument and can't
be combined with Optional.

```vba
' =COUNTFILLED(C2, D2, E2) counts how many of the cells you pass are not blank
Public Function COUNTFILLED(ParamArray inputs() As Variant) As Long
    Dim v As Variant
    For Each v In inputs
        If Len(Trim$(CStr(ArgValue(v)))) > 0 Then COUNTFILLED = COUNTFILLED + 1
    Next v
End Function
```

### 5. Range arguments and input validation

When a formula passes a cell (`=BMI(G3,F3)`), a Variant argument receives a **Range object**, not a number. When it
passes a number or the result of another formula (`=BMI(197.9,70.4)`), the argument receives the value. The starter
module's `ArgValue` helper hides the difference:

```vba
' A cell reference arrives as a Range object; a typed number or a formula
' result arrives as a plain value. Return the value either way.
Private Function ArgValue(arg As Variant) As Variant
    If IsObject(arg) Then
        ArgValue = arg.Cells(1, 1).Value
    Else
        ArgValue = arg
    End If
End Function
```

The type you declare for an argument decides who handles bad input:

| Declaration | A cell with text such as N/A | A blank cell | Who picks the error |
|---|---|---|---|
| `weightLb As Double` | Excel returns #VALUE! before your code runs | Arrives as 0, silently | Excel |
| `weightLb As Variant` plus your own checks | Your code returns the error you choose | Arrives as Empty, so you can reject it | You |

To check a value, don't trust `IsNumeric`. It's built to answer "could this text be converted to a number?", which is a
different question:

| Value | `IsNumeric(v)` | `VarType(v)` | `IsNumberValue(v)` |
|---|---|---|---|
| a blank cell (Empty) | **True** ⚠️ | 0 `vbEmpty` | False |
| the text "12" | **True** ⚠️ | 8 `vbString` | False |
| the text "N/A" | False | 8 `vbString` | False |
| 180.4 | True | 5 `vbDouble` | True |
| a date | **False** ⚠️ | 7 `vbDate` | True |

The starter's `IsNumberValue` helper tests the type instead, so it accepts real numbers and dates and rejects
everything else:

```vba
Private Function IsNumberValue(v As Variant) As Boolean
    Select Case VarType(v)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbDate, vbByte
            IsNumberValue = True
        Case Else
            IsNumberValue = False
    End Select
End Function
```

> ⚠️ `IsNumeric(Empty)` is True, so a test like `If IsNumeric(amount) Then total = total + amount` treats a blank
> BilledAmount as \$0 and counts the claim. That's the trap the bonus is built around.

### 6. Returning Excel errors with CVErr

`CVErr(number)` builds an Excel **error value** that a cell displays as #VALUE!, #N/A, and so on. A function can only
return one if its return type is **Variant**. A function declared `As Double` can't hold an error value.

| Constant | Value | Cell shows | Return it when… |
|---|---|---|---|
| `xlErrValue` | 2015 | #VALUE! | an input is the wrong kind, or the data can't be right (a discharge before the admission) |
| `xlErrNum` | 2036 | #NUM! | a number has the right type but is impossible (a height of 0) |
| `xlErrNA` | 2042 | #N/A | no answer exists (a lookup that finds nothing) |
| `xlErrDiv0` | 2007 | #DIV/0! | the denominator is empty (a rate over zero claims) |
| `xlErrRef` | 2023 | #REF! | a reference is invalid |
| `xlErrName` | 2029 | #NAME? | rarely returned on purpose |
| `xlErrNull` | 2000 | #NULL! | rarely returned on purpose |

Here is a UDF that turns an Emergency Severity Index (ESI) level into its name. It returns #VALUE! for the wrong kind of
input and #N/A for a number that isn't an ESI level:

```vba
' =ESINAME(2) returns "Emergent". Anything but a whole number from 1 to 5 returns an error.
Public Function ESINAME(esiLevel As Variant) As Variant
    Dim lvl As Variant
    lvl = ArgValue(esiLevel)
    If Not IsNumberValue(lvl) Then
        ESINAME = CVErr(xlErrValue)            ' blank or text: the wrong kind of input
        Exit Function
    End If
    Select Case lvl
        Case 1: ESINAME = "Resuscitation"
        Case 2: ESINAME = "Emergent"
        Case 3: ESINAME = "Urgent"
        Case 4: ESINAME = "Less urgent"
        Case 5: ESINAME = "Non-urgent"
        Case Else: ESINAME = CVErr(xlErrNA)    ' a number, but not an ESI level
    End Select
End Function
```

Your LOSDAYS function in task 4 follows the same shape: check the inputs, return `CVErr(xlErrValue)` for anything
impossible, and only then calculate. For the calculation, `DateDiff("d", admit, discharge)` counts the calendar-day
boundaries (midnights) between two date-times.

**Worked example.** Encounter ENC111251 (row 26 of the Encounters sheet) is an observation stay admitted 01/15/2025 at
21:16 and discharged 01/16/2025 at 10:27. That's only 13 hours, but it crosses one midnight, so
`DateDiff("d", …)` returns 1. Counting midnights is how hospitals count inpatient days.

Why return an error instead of 0 or a blank? A 0 looks exactly like a same-day stay, so a typo would quietly pull
the average length of stay down. An error can't be mistaken for data. You can still summarize around it:
`AGGREGATE(9,6,range)` sums a range while ignoring error cells (Excel 2010 and later), and `COUNTIF(range,"#VALUE!")`
counts one specific error.

**A whole range as an argument.** Some functions need a block of cells, such as a column of claim statuses. Declare that
argument `As Range` and loop over its cells with `For Each`, as in Lesson 5.2. A Table column reference such as
`tblEncounters[LOSDays]` arrives as the column's data cells, without the header. Two tests help inside the loop:

| Test | True when the cell… |
|---|---|
| `IsEmpty(cell.Value)` | is blank |
| `IsError(cell.Value)` | shows an error value such as #VALUE! or #N/A |

This function averages the numbers in a range and skips everything else. Excel's AVERAGE returns an error if any cell in
its range shows an error, so after task 4 `=AVGNUMBERS(tblEncounters[LOSDays])` gives the average length of stay of the
valid rows. (`AGGREGATE(1,6,range)` does the same job without a macro.)

```vba
' =AVGNUMBERS(range): the average of the numbers in range. Blanks, text, and
' error cells are skipped. #DIV/0! if the range holds no numbers at all.
Public Function AVGNUMBERS(values As Range) As Variant
    Dim cell As Range, total As Double, n As Long
    For Each cell In values.Cells
        If IsNumberValue(cell.Value) Then      ' False for Empty, text, and error values
            total = total + cell.Value
            n = n + 1
        End If
    Next cell
    If n = 0 Then
        AVGNUMBERS = CVErr(xlErrDiv0)          ' nothing to divide by
    Else
        AVGNUMBERS = total / n
    End If
End Function
```

Because `values` is declared `As Range`, a formula that passes a typed number instead, such as `=AVGNUMBERS(5)`, gets
#VALUE! from Excel before your code runs. To try ESINAME or AVGNUMBERS, paste it into HealthUDFs, because both call the
starter's Private helpers, and a Private procedure can only be called from its own module.

To compare text without caring about case, use `StrComp(a, b, vbTextCompare) = 0` (Lesson 5.3) or compare `UCase(a)`
with `UCase(b)` (Lesson 5.2).

> ⚠️ Test `IsError` before you compare or convert a value. `cell.Value = "Denied"` and `CStr(cell.Value)` both raise a
> Type mismatch error when the cell holds an error value, and inside a UDF that turns the whole result into #VALUE!.

### 7. Calling Excel's worksheet functions from VBA

VBA has its own functions for text and dates (Trim, InStr, Year, DateDiff), but MATCH, VLOOKUP, SUMIFS, and MEDIAN only
exist in Excel. You can call them two ways, and they report failure differently.

| | `Application.WorksheetFunction.Match(...)` | `Application.Match(...)` |
|---|---|---|
| Argument hints (IntelliSense) | Yes | No |
| When the function fails (no match, #DIV/0!…) | Raises run-time error **1004** and stops the code, or jumps to your handler | Returns an error **value** (Error 2042 = #N/A) |
| Variable that stores the result | any matching type | must be a Variant |
| How you check for failure | an error handler | `IsError(result)` |

Use `Application.Match` when "not found" is a normal outcome you want to test for. This helper from the starter
modules finds a column by its header text, so your macros keep working if someone inserts a column:

```vba
Private Function HeaderColumn(ws As Worksheet, ByVal headerText As String) As Long
    Dim pos As Variant
    pos = Application.Match(headerText, ws.Rows(1), 0)   ' an error VALUE, not a crash
    If IsError(pos) Then
        Err.Raise vbObjectError + 513, "HeaderColumn", _
                  "Column '" & headerText & "' was not found on sheet '" & ws.Name & "'."
    End If
    HeaderColumn = CLng(pos)
End Function
```

When it can't find the header, it raises a clear error of its own (section 12 explains `Err.Raise`).

### 8. Collections

A **Collection** is VBA's built-in list. You add **items**, optionally each with a text **key**, and get them back by
position or by key. Collections work in every version of Excel on Windows and Mac.

```vba
Dim reasons As Collection
Set reasons = New Collection

reasons.Add Item:="Coding Error", Key:="Coding Error"
reasons.Add Item:="Timely Filing", Key:="Timely Filing"

Debug.Print reasons.Count               ' 2
Debug.Print reasons(1)                  ' Coding Error  (positions start at 1)
Debug.Print reasons("Timely Filing")    ' Timely Filing (look up by key)

Dim r As Variant
For Each r In reasons
    Debug.Print r
Next r

reasons.Remove "Coding Error"
```

| Rule | What it means for you |
|---|---|
| Positions start at 1 | `reasons(1)` is the first item |
| Keys are text and not case-sensitive | "coding error" and "Coding Error" are the same key |
| Adding a key that's already there | raises run-time error 457 |
| Reading a key that isn't there | raises run-time error 5 |
| There's no `.Exists` and no list of keys | you test for a key by trapping the error |
| You can't change an item in place | Remove it, then Add it again |

The "already there" error is useful, because it makes a Collection a natural **distinct list**. Add every value with
itself as the key, and ignore the error for duplicates:

```vba
' Adds keyText to a Collection once. Returns False if it was already there.
Private Function AddIfNew(col As Collection, ByVal keyText As String) As Boolean
    On Error Resume Next
    col.Add Item:=keyText, Key:=keyText    ' error 457 if the key already exists
    AddIfNew = (Err.Number = 0)
    On Error GoTo 0
End Function
```

`On Error Resume Next` is switched on for exactly one statement, the result is read from `Err.Number` straight away, and
`On Error GoTo 0` switches normal error handling back on. Section 12 explains why that discipline matters.

### 9. Scripting.Dictionary

A **dictionary** stores **key → item** pairs, like a two-column lookup table that lives in memory. Unlike a Collection,
it can tell you whether a key exists, list its keys, and change an item in place. It comes from the Microsoft Scripting
Runtime library, which is part of Windows.

| | Late binding (used in this lesson) | Early binding |
|---|---|---|
| Code | `Dim d As Object`<br>`Set d = CreateObject("Scripting.Dictionary")` | **Tools → References** → tick *Microsoft Scripting Runtime*, then<br>`Dim d As Scripting.Dictionary`<br>`Set d = New Scripting.Dictionary` |
| Argument hints (IntelliSense) | No | Yes |
| Setup on someone else's PC | None | None on Windows, because the reference is saved with the file |

| Member | What it does | Example |
|---|---|---|
| `.Add key, item` | Adds a pair. Raises error 457 if the key exists | `d.Add "PY01", "Medicare"` |
| `d(key)` or `.Item(key)` | Reads or writes an item. Writing to a new key adds it | `d("PY01") = "Medicare"` |
| `.Exists(key)` | True if the key is there | `If d.Exists(id) Then` |
| `.Count` | Number of keys | `Debug.Print d.Count` |
| `.Keys`, `.Items` | All keys, or all items, as a 0-based array | `For Each k In d.Keys` |
| `.Remove key`, `.RemoveAll` | Deletes one pair, or all of them | `d.Remove "PY08"` |
| `.CompareMode` | `vbBinaryCompare` (default: "PY01" and "py01" differ) or `vbTextCompare`. Set it before you add anything | `d.CompareMode = vbTextCompare` |

> ⚠️ The argument order is reversed between the two objects: `Collection.Add item, key` but `Dictionary.Add key,
> item`. Named arguments (`Item:=`, `Key:=`) make Collection code easier to read.

Almost every dictionary job is one of four patterns:

```vba
' 1. Distinct: keep one of each
If Not d.Exists(patientID) Then d.Add patientID, True

' 2. Count per key: reading a missing key gives Empty, and Empty + 1 = 1
d(status) = d(status) + 1

' 3. Sum per key
d(payerID) = d(payerID) + paidAmount

' 4. Lookup: load once, then ask as often as you like
payerNames(payers(r, 1)) = payers(r, 2)                ' while loading the Payers sheet
If payerNames.Exists(id) Then nm = payerNames(id) Else nm = "(unknown payer)"
```

**Worked example.** This macro counts the stays on the Encounters sheet by EncounterType. It reads the whole table into
an array first (Lesson 5.3), so the loop never touches the sheet:

```vba
Sub CountStaysByType()
    Dim stays As Variant, r As Long, k As Variant
    Dim byType As Object
    Set byType = CreateObject("Scripting.Dictionary")
    stays = ThisWorkbook.Worksheets("Encounters").Range("A1").CurrentRegion.Value
    For r = 2 To UBound(stays, 1)                      ' row 1 is the header
        byType(stays(r, 3)) = byType(stays(r, 3)) + 1  ' column 3 = EncounterType
    Next r
    For Each k In byType.Keys
        Debug.Print k, byType(k)
    Next k
End Sub
```

The Immediate window shows `Inpatient 408` and `Observation 97`, in the order each type first appeared.

> ⚠️ **Reading a key adds it.** `If d(id) = "" Then` quietly adds `id` with an Empty item, so `.Count` grows. Use
> `.Exists` whenever you only want to ask.

> ⚠️ **Keys keep their type.** The number 1001 (read from a cell) and the text "1001" are different keys. Convert
> with `CStr` when an ID could arrive either way.

**Several numbers per key.** To keep a count *and* totals for each payer, store an array as the item. You can't change
the array inside the dictionary directly, because `stats(id)(0) = …` changes a temporary copy that the dictionary never
sees. Copy it out, change it, and put it back:

```vba
If Not stats.Exists(id) Then stats.Add id, Array(0, 0)   ' (count, total paid)
s = stats(id)              ' 1. copy the array out
s(0) = s(0) + 1            ' 2. change the copy
s(1) = s(1) + paidAmount
stats(id) = s              ' 3. put it back
```

### 10. Collections and dictionaries on a Mac

Excel for Mac has no Scripting Runtime, so `CreateObject("Scripting.Dictionary")` stops with run-time error 429. A
Collection works everywhere, so Mac code builds on it.

| | Collection | Scripting.Dictionary |
|---|---|---|
| Available | Windows and Mac | Windows only |
| Create | `Set c = New Collection` | `Set d = CreateObject("Scripting.Dictionary")` |
| Add | `c.Add item, key` | `d.Add key, item` or `d(key) = item` |
| Keys | optional, text only, not case-sensitive | required, any type, case-sensitive by default |
| Is this key here? | trap error 5 | `d.Exists(key)` |
| Change an item | Remove, then Add | `d(key) = newValue` |
| List the keys | not possible | `d.Keys` |
| First position | 1 | 0 (for `.Keys` and `.Items`) |

To count or total on a Mac, let a Collection map each key to a **slot** number, and keep the numbers you update in
ordinary arrays indexed by that slot:

```vba
' Returns the slot number stored under keyText. A new key gets the next free slot.
Private Function SlotFor(slots As Collection, ByVal keyText As String) As Long
    On Error Resume Next
    SlotFor = slots(keyText)               ' error 5 if the key isn't there yet
    If Err.Number <> 0 Then
        slots.Add Item:=slots.Count + 1, Key:=keyText
        SlotFor = slots.Count
    End If
    On Error GoTo 0
End Function

' In the loop: one slot per payer, counts in a parallel array
i = SlotFor(slots, CStr(claims(r, payerCol)))
ids(i) = CStr(claims(r, payerCol))
counts(i) = counts(i) + 1
```

The files ending in `Mac_Solution.bas` in [`solutions/`](solutions/) replace every dictionary in the lesson with
Collections: AddIfNew for distinct lists, and SlotFor plus arrays for counts and totals. They also run on Windows.

### 11. Sorting the keys

A dictionary hands back its keys in the order you added them. You have two ways to sort:

1. **Write the results to a sheet, then sort the range.** This is fast for any size and works on Windows and Mac:

   ```vba
   With ws.Range("A1").Resize(n + 1, 3)          ' header row + n result rows
       .Sort Key1:=.Cells(1, 3), Order1:=xlDescending, Header:=xlYes
   End With
   ```

2. **Sort the keys array in memory** with a simple swap sort. It compares every pair, so keep it for a few hundred
   keys at most:

   ```vba
   Dim keyList As Variant, i As Long, j As Long, tmp As Variant
   keyList = d.Keys                          ' a 0-based array
   For i = LBound(keyList) To UBound(keyList) - 1
       For j = i + 1 To UBound(keyList)
           If keyList(j) < keyList(i) Then   ' to sort by count instead: d(keyList(j)) > d(keyList(i))
               tmp = keyList(i): keyList(i) = keyList(j): keyList(j) = tmp
           End If
       Next j
   Next i
   ```

> ⚠️ Sort *before* you add a total row. `Range.Sort` doesn't know a row is special, so a TOTAL row inside the sorted
> range moves to the top.

### 12. Handling run-time errors

A **run-time error** happens while code runs: a sheet that doesn't exist, text where a number should be, a division by
zero. With no error handling, VBA stops and shows a dialog with **End** and **Debug** buttons. Every setting your macro
changed stays changed, and the user is left with a half-finished report.

| Statement | After an error, VBA… | Use it for |
|---|---|---|
| `On Error GoTo Fail` | jumps to the line labeled `Fail:` | most macros: one handler per procedure |
| `On Error Resume Next` | skips the failing statement and carries on | one statement you expect might fail, followed straight away by a check of `Err.Number` |
| `On Error GoTo 0` | stops with the normal dialog again | ending a Resume Next section |

Inside a handler, a **Resume** statement decides where the code continues:

| Statement | Continues at | Watch out |
|---|---|---|
| `Resume` | the statement that failed (it retries) | if the cause isn't fixed, it loops forever |
| `Resume Next` | the statement after the one that failed | the failed statement's work is simply skipped |
| `Resume CleanExit` | a label you choose | the usual way out: run the cleanup code |

The **Err object** describes the most recent error:

| Member | Meaning |
|---|---|
| `Err.Number` | The error number. 0 means no error |
| `Err.Description` | The message, such as "Type mismatch" |
| `Err.Source` | Where it came from |
| `Err.Raise number, source, description` | Raises an error of your own |
| `Err.Clear` | Resets the number to 0. Any Resume, Exit Sub, Exit Function, or On Error statement also clears it |

The errors you'll meet most often:

| Number | Message | Typical cause |
|:-:|---|---|
| 5 | Invalid procedure call or argument | asking a Collection for a key it doesn't have |
| 6 | Overflow | a row count stored in an Integer |
| 9 | Subscript out of range | `Worksheets("Claims")` after someone renamed the sheet |
| 11 | Division by zero | a rate for a payer with no claims |
| 13 | Type mismatch | `CLng("N/A")`, or adding text to a number |
| 91 | Object variable or With block variable not set | using a Range variable before `Set` |
| 429 | ActiveX component can't create object | `Scripting.Dictionary` on a Mac |
| 457 | This key is already associated with an element of this collection | adding a duplicate key |
| 1004 | Application-defined or object-defined error | Excel object errors, such as `WorksheetFunction.Match` finding nothing |

**Raise your own errors** when your code detects a problem it can't fix. Add your number to `vbObjectError`, so it
can't clash with VBA's own numbers (513 to 65535 are yours to use):

```vba
Err.Raise vbObjectError + 513, "HeaderColumn", "Column 'BilledAmount' was not found on sheet 'Claims'."
```

**Errors travel up the call stack.** If a procedure has no handler, VBA passes the error to the procedure that called
it, and so on until one has a handler. That's why HeaderColumn can raise an error and the macro that called it reports
the message in its own handler.

**Resume Next done right.** Keep it to the one statement that might fail, read `Err.Number`, then switch it off:

```vba
' True if this workbook has a sheet with that name
Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)    ' error 9 if there is no such sheet
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function
```

> ⚠️ Three classic mistakes:
> - Forgetting `Exit Sub` above the handler label, so the handler code runs every time, even after success.
> - Expecting the handler to catch an error raised *inside* the handler. It can't. That error goes to the caller, or
>   to the user.
> - Leaving `On Error Resume Next` switched on for a whole procedure. Every later bug is hidden, and variables quietly
>   keep stale values.

> 💡 **Tip:** In the VBE, **Tools → Options → General → Error Trapping** controls when VBA stops. **Break on Unhandled
> Errors** lets your handlers do their job. **Break in Class Module** behaves the same way for the standard-module code
> in this lesson, so either is fine. **Break on All Errors** stops on every error, even handled ones, which helps while
> debugging but makes working handlers (such as AddIfNew's duplicate-key trap) look broken. If your handlers never seem
> to run, check this setting first.

### 13. The cleanup pattern: validate, trap, restore

Robust macros follow the same shape. They **validate** what they can predict and explain it in plain words. They
**trap** everything else in one handler. And they **restore** every setting they changed, whether the macro succeeded
or failed.

```vba
Public Sub BuildReport()
    Dim prevCalc As Long

    ' 1. Validate what you can predict, before changing anything
    If Not SheetExists("Claims") Then
        MsgBox "I can't find a sheet named 'Claims'. Was it renamed?", vbExclamation, "Build report"
        Exit Sub
    End If

    ' 2. Save the settings you'll change, then switch the handler on
    prevCalc = Application.Calculation
    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' 3. ...the real work...

CleanExit:                                  ' 4. runs after success AND after an error
    On Error Resume Next                    '    cleanup must never raise an error of its own
    Application.Calculation = prevCalc
    Application.ScreenUpdating = True
    Exit Sub

Fail:                                       ' 5. explain what happened, then clean up
    MsgBox "BuildReport stopped: " & Err.Description & " (error " & Err.Number & ")", vbCritical, "Build report"
    Resume CleanExit
End Sub
```

1. When everything works, the code falls through into `CleanExit`, restores the settings, and leaves at `Exit Sub`. The
   handler below never runs.
2. When any statement fails, VBA jumps to `Fail`, shows the message, and `Resume CleanExit` sends it back up to the same
   cleanup code.

Restoring matters because some settings outlive your macro. Excel usually switches ScreenUpdating back on when a macro
ends, but **calculation mode stays manual** for every open workbook until someone changes it, and `EnableEvents` (Lesson 5.5)
stays off. A macro that crashes halfway can leave a whole department looking at numbers that no longer recalculate.

> 💡 **Tip:** Read data with `.Value2` when you'll test types. `.Value` returns Currency for currency-formatted cells
> and Date for date cells. `.Value2` returns plain Doubles for both, so a check like `VarType(v) = vbDouble` stays
> simple.

### 14. Organizing and sharing your code

- **One module per purpose.** Name each module in the Properties window (F4), for example HealthUDFs or
  ClaimDictionaries.
- **`Option Explicit`** at the top of every module, as in Lesson 5.2.
- **`Private`** for helpers, so they stay out of the Insert Function list and out of other modules' way.
- **Variable names that don't clash with Excel's.** The VBE keeps one spelling per name across the project, so a
  variable called `keys` or `value` changes how `.Keys` and `.Value` are displayed everywhere. That's why the helpers
  use names like `keyText` and `headerText`.
- **A comment above each procedure** that says what it returns, what it expects, and which errors it can return.
- **Debug → Compile VBAProject** before you save. It catches typos in procedures you haven't run yet.
- **Export your modules** (right-click a module → **Export File…**) to back them up or track them in version control.
  That's how the .bas files in this lesson were made.

On Windows, you can give a UDF a description and argument help in the Insert Function dialog with
`Application.MacroOptions` (`ArgumentDescriptions` needs Excel 2010 or later). Excel doesn't reliably keep these
descriptions after the file closes, so call the macro every time the file opens, from the Workbook_Open event in
Lesson 5.5:

```vba
Sub DescribeHealthUDFs()
    Application.MacroOptions Macro:="BMI", Category:="Bluestone Health", _
        Description:="Body-mass index from weight in pounds and height in inches.", _
        ArgumentDescriptions:=Array("Weight in pounds", "Height in inches")
End Sub
```

**Where to keep your UDFs**

| Store them in | Call them as | Good for |
|---|---|---|
| This workbook (.xlsm) | `=BMI(G2,F2)` | functions that belong to one report |
| Your Personal Macro Workbook (PERSONAL.XLSB) | `=PERSONAL.XLSB!BMI(G2,F2)` | your own shortcuts (other people can't recalculate them) |
| An add-in (.xlam) | `=BMI(G2,F2)` in any workbook while the add-in is loaded | sharing functions with a team |

To turn your UDF modules into an **add-in**:

1. Copy the modules into a new workbook (drag them in the Project Explorer, or export and import them), then compile.
2. Choose **File → Save As**, and pick **Excel Add-in (\*.xlam)** as the file type. On Windows, Excel switches to your
   AddIns folder. Keep it there, or choose a shared folder your team can reach.
3. Load it. On Windows, choose **File → Options → Add-ins**, set **Manage** to *Excel Add-ins*, click **Go…**, and
   then click **Browse…**. On a Mac, choose **Tools → Excel Add-ins…** and click **Browse**. The **Developer → Excel
   Add-ins** button opens the same dialog on both.
4. Tick the add-in and click **OK**. Its functions now work in every workbook you open.

> ⚠️ A workbook that uses add-in functions stores the add-in's location. On a colleague's computer without the add-in,
> those cells can't recalculate. Share the add-in, or keep the functions in the workbook itself.

## 🧪 Hands-on practice

Download [`5.4-vba-functions-error-handling.xlsx`](5.4-vba-functions-error-handling.xlsx), save it as .xlsm, and import
the starter modules (guide section 1). Type each answer in its yellow cell. A gray cell fills itself from your work
elsewhere: a yellow column of UDF formulas on a data sheet, or a cell your macro writes on the Output sheet. The
**Check** column turns green when you're right.

<!-- BEGIN GENERATED: practice -->
Save the workbook as .xlsm and import the starter modules first (see the Start Here sheet). In tasks 1–6 you write user-defined functions and use them in cells: tasks 1 and 6 in the yellow cell, and tasks 2–5 in yellow columns on the data sheets, which the gray cells summarize. Tasks 7 and 11–13 ask you to predict what a snippet on the Snippets sheet does. In tasks 8–10 you run macros, and the gray cells read what they wrote to the Output sheet.

| # | Task | Hint |
|:-:|------|------|
| 1 | Write BMI(weightLb, heightIn) in a standard module (start from starter/HealthUDFs.bas). BMI = 703 × weight in pounds ÷ (height in inches)². It must return #VALUE! for a blank or text input and #NUM! for a zero or negative one. In the yellow cell, call it for the first patient on the Patients sheet (row 2, PT10002). Keep full precision. The check accepts 2 decimal places. | Assign the result to the function's own name. Return errors with CVErr |
| 2 | Fill the yellow BMI column on the Patients sheet with your BMI function for every patient. Don't round inside the function: the column's number format already shows 1 decimal place. The gray cell counts the patients whose unrounded BMI is 30 or more (the usual adult obesity threshold). | Type it once in the first BMI cell. The Table fills the rest |
| 3 | Write AGEAT(dob, [asOf]): a person's age in completed years on the asOf date, or on today's date when asOf is left out (use IsMissing). Then fill the yellow AgeAtAdmit column on the Encounters sheet with each patient's age on the admission date. The gray cell averages your column, and the check compares that average to 2 decimal places. | Optional asOf As Variant + IsMissing. Then check whether this year's birthday has happened |
| 4 | Write LOSDAYS(admitDateTime, dischargeDateTime): the number of midnights between admission and discharge, or #VALUE! when an input isn't a date or the discharge date-time is earlier than the admission date-time. Fill the yellow LOSDays column on the Encounters sheet. The gray cell adds up your column and skips error cells. What is the total? | CVErr(xlErrValue) for bad rows. DateDiff("d", …) counts midnights |
| 5 | How many rows does your LOSDays column flag with #VALUE!? The gray cell counts them. (Each one is a data-entry error to send back to Health Information Management.) | Filter the LOSDays column to see the error rows |
| 6 | Write DENIALRATE(statusRange, [statusToCount]): the share of non-blank cells in statusRange whose text equals statusToCount (ignoring case). statusToCount is Optional and defaults to "Denied". Return #DIV/0! if the range has no non-blank cells. In the yellow cell, call DENIALRATE on the ClaimStatus column of the Claims sheet and leave out the second argument. The cell is already formatted as a percentage. | Optional statusToCount As String = "Denied", then For Each cell In statusRange.Cells |
| 7 | Snippet A on the Snippets sheet passes a \$100 charge to AddFeeByVal and AddFeeByRef. What number does Debug.Print charge show? Predict first, then run it to check. | ByRef shares the caller's variable. What do parentheses around an argument do? |
| 8 | Complete CountDenialReasons in starter/ClaimDictionaries.bas: add every non-blank DenialReason on the Claims sheet to a Collection, using the reason itself as the key, so each reason is kept once. The macro writes the Collection's Count to Output!B4, and the gray cell reads it. | col.Add Item:=reason, Key:=reason raises an error for a key it already has |
| 9 | Complete CountDistinctPatients: add every PatientID on the Claims sheet to a Scripting.Dictionary (Mac: a Collection) once, then write the number of keys to Output!B5. How many different patients had a claim? | CreateObject("Scripting.Dictionary"), then .Exists and .Add |
| 10 | Complete TopDeniedPayer: build one dictionary that maps PayerID → PayerName (from the Payers sheet) and another that counts the claims with ClaimStatus "Denied" for each PayerID (Mac: a Collection for the names, and SlotFor plus an array for the counts). Write the PayerName with the most denied claims to Output!B6 and its count to Output!B7. The gray cell reads B6. | d(key) = d(key) + 1 counts. Then loop over .Keys to find the largest |
| 11 | Snippet B reads five blood-glucose readings typed as text and adds them up, with an error handler that skips bad entries. What total does the last line print? Predict first, then run it to check. | Resume Next continues after the line that failed |
| 12 | Snippet C reads the same five values with On Error Resume Next instead of a handler. What total does it print? Predict first, then run it to check. | When an assignment fails, what's left in the variable? |
| 13 | Snippet D looks up a payer ID that doesn't exist (PY09), first with Application.Match and then with Application.WorksheetFunction.Match. The first prints a message. The second stops the macro with a run-time error. Type the error number. Predict first, then run it to check. | The 'application-defined or object-defined' error number |
<!-- END GENERATED: practice -->

**Code for the predict-the-output tasks.** The same code is on the workbook's **Snippets** sheet and in
[`starter/Snippets.bas`](starter/Snippets.bas). Predict first, then click inside the snippet's Sub and press **F5**
(Mac: **Run → Run Sub/UserForm**) to check. Debug.Print writes to the Immediate window. Snippet D stops with a run-time
error dialog on purpose: note the number, then click **End**.

**Snippet A** (task 7): ByVal vs ByRef

```vba
Sub AddFeeByVal(ByVal amount As Double)
    amount = amount + 25
End Sub

Sub AddFeeByRef(ByRef amount As Double)
    amount = amount + 25
End Sub

Sub SnippetA()
    Dim charge As Double
    charge = 100
    AddFeeByVal charge
    AddFeeByRef charge
    AddFeeByRef (charge)
    Debug.Print charge
End Sub
```

**Snippet B** (task 11): an error handler inside a loop

```vba
Sub SnippetB()
    Dim readings As Variant, i As Long, total As Long
    readings = Array("120", "abc", "95", "", "88")
    On Error GoTo BadReading
    For i = LBound(readings) To UBound(readings)
        total = total + CLng(readings(i))
    Next i
    Debug.Print "Total:"; total
    Exit Sub
BadReading:
    Debug.Print "Skipped item"; i; "- error"; Err.Number
    Resume Next
End Sub
```

**Snippet C** (task 12): On Error Resume Next

```vba
Sub SnippetC()
    Dim readings As Variant, i As Long, reading As Long, total As Long
    readings = Array("120", "abc", "95", "", "88")
    On Error Resume Next
    For i = LBound(readings) To UBound(readings)
        reading = CLng(readings(i))
        total = total + reading
    Next i
    On Error GoTo 0
    Debug.Print "Total:"; total
End Sub
```

**Snippet D** (task 13): two ways to call MATCH from VBA

```vba
Sub SnippetD()
    Dim payerList As Range, pos As Variant
    Set payerList = ThisWorkbook.Worksheets("Payers").Range("A2:A9")

    pos = Application.Match("PY09", payerList, 0)
    If IsError(pos) Then
        Debug.Print "Application.Match: not found"
    Else
        Debug.Print "Application.Match: row"; pos
    End If

    pos = Application.WorksheetFunction.Match("PY09", payerList, 0)
    Debug.Print "WorksheetFunction.Match: row"; pos
End Sub
```

## ✅ Answer key

The workbook has hidden **Answer Key** and **Bonus Key** sheets (right-click any sheet tab → **Unhide…**). The .xlsx
file can't contain your UDFs, so for those tasks the key's *Live result* column runs an equivalent worksheet formula
instead. The complete modules are in [`solutions/`](solutions/). The answers are also below, collapsed so you don't see
them by accident.

<!-- BEGIN GENERATED: answers -->
<details>
<summary><b>🔑 Show the answer key</b> — Try every task before opening this.</summary>

**1. BMI for the first patient**

- **Answer:** 33.21
- **Solution:**

```vba
' In the yellow cell: =BMI(Patients!G2,Patients!F2)

' Tasks 1-2: body-mass index = 703 x weight (lb) / height (in) ^ 2
Public Function BMI(weightLb As Variant, heightIn As Variant) As Variant
    Dim w As Variant, h As Variant
    w = ArgValue(weightLb)
    h = ArgValue(heightIn)
    If Not IsNumberValue(w) Or Not IsNumberValue(h) Then
        BMI = CVErr(xlErrValue)            ' blank or text input: #VALUE!
    ElseIf w <= 0 Or h <= 0 Then
        BMI = CVErr(xlErrNum)              ' impossible measurement: #NUM!
    Else
        BMI = 703 * w / h ^ 2
    End If
End Function
```


A **Function** hands back a value by assigning it to its own name (`BMI = …`). Because it's Public and sits in a standard module, Excel lists it with the built-in functions, so `=BMI(` works in any cell. The return type is Variant so the same function can return a number or an error value made by `CVErr`. The two helpers come from the starter module: `ArgValue` turns a cell reference into its value, and `IsNumberValue` rejects blanks and text. Brenda Peterson weighs 180.4 lb at 61.8 in, so BMI = 703 × 180.4 ÷ 61.8² ≈ 33.21. If the cell shows #NAME?, the function isn't in a standard module of *this* workbook, or macros are disabled.

**2. BMI column (patients with BMI ≥ 30)**

- **Answer:** 258
- **Solution:** `=BMI([@WeightLb],[@HeightIn])`

Type the formula in H2 (or `=BMI(G2,F2)`, which does the same). A UDF behaves like any other function: write it once in the first row and the Excel Table copies it down as a calculated column, so each of the 591 rows recalculates whenever its height or weight changes. That's the payoff of a UDF over a macro, because a macro would write values once and go stale. The key's live formula does the same math with SUMPRODUCT, which shows a UDF is a convenience: anything it does with simple arithmetic you could also do with a formula.

**3. AgeAtAdmit column (average age at admission)**

- **Answer:** 60.08
- **Solution:**

```vba
' In the AgeAtAdmit column (G2): =AGEAT([@DOB],[@AdmitDateTime])

' Task 3: age in completed years on asOf. Leave asOf out to use today's date.
Public Function AGEAT(dob As Variant, Optional asOf As Variant) As Variant
    Dim birthValue As Variant, refValue As Variant
    Dim birth As Date, refDate As Date, years As Long

    birthValue = ArgValue(dob)
    If IsMissing(asOf) Then
        refValue = Date                    ' second argument left out: today
        Application.Volatile               ' today changes, so recalculate every time
    Else
        refValue = ArgValue(asOf)
    End If
    If Not IsNumberValue(birthValue) Or Not IsNumberValue(refValue) Then
        AGEAT = CVErr(xlErrValue)
        Exit Function
    End If

    birth = CDate(birthValue)
    refDate = CDate(refValue)
    If refDate < birth Then
        AGEAT = CVErr(xlErrNum)            ' asOf is before the date of birth
        Exit Function
    End If

    years = Year(refDate) - Year(birth)
    ' Birthday still to come in the asOf year? Then the person is a year younger.
    If DateSerial(Year(refDate), Month(birth), Day(birth)) > refDate Then years = years - 1
    AGEAT = years
End Function
```


`IsMissing` works only on an **Optional Variant** argument, which is why `asOf` is declared `As Variant`. The tricky part is the birthday: `Year(asOf) - Year(dob)` counts calendar years, so it's one too high whenever the birthday falls later in the year than the admission. `DateSerial(Year(asOf), Month(dob), Day(dob))` builds this year's birthday. If that's after asOf, subtract 1. `DateDiff("yyyy", dob, asOf)` has the same flaw because it counts year boundaries, not birthdays. On this data it is wrong for 272 of 505 stays and gives an average of 60.62 instead of 60.08.

**4. LOSDays column (total of the valid stays)**

- **Answer:** 1,926
- **Solution:**

```vba
' In the LOSDays column (H2): =LOSDAYS([@AdmitDateTime],[@DischargeDateTime])

' Tasks 4-5: length of stay = midnights between admission and discharge.
' #VALUE! when an input isn't a date or the discharge is before the admission.
Public Function LOSDAYS(admitDateTime As Variant, dischargeDateTime As Variant) As Variant
    Dim a As Variant, d As Variant
    a = ArgValue(admitDateTime)
    d = ArgValue(dischargeDateTime)
    If Not IsNumberValue(a) Or Not IsNumberValue(d) Then
        LOSDAYS = CVErr(xlErrValue)
    ElseIf CDate(d) < CDate(a) Then
        LOSDAYS = CVErr(xlErrValue)        ' discharge before admission: a data-entry error
    Else
        LOSDAYS = DateDiff("d", CDate(a), CDate(d))   ' "d" counts the midnights crossed
    End If
End Function
```


`DateDiff("d", admit, discharge)` counts the calendar-day boundaries crossed, so a patient admitted at 23:00 and discharged at 01:00 the next morning has a stay of 1 midnight, which is how inpatient days are counted. Returning `CVErr(xlErrValue)` makes the bad rows impossible to miss, and it keeps them out of totals: the gray cell uses `AGGREGATE(9,6,…)` (9 = SUM, 6 = ignore error values). If you had returned 0 or a negative number instead, it would have silently dragged the total down.

**5. LOSDays rows flagged #VALUE!**

- **Answer:** 4
- **Solution:** `=COUNTIF(tblEncounters[LOSDays],"#VALUE!")`

Nothing new to type here: the gray cell already holds this formula. COUNTIF can count one specific error value when you give the error's text as the criteria. Three discharges were keyed with the year 2024, and one same-day stay has its discharge time keyed 12 hours early (an AM/PM slip). Compare full date-times, not just dates: a check like `DateValue(d) < DateValue(a)` finds only 3 of the 4 problems, because the AM/PM slip has the right date. The check counts #VALUE! only, so a function that returns #N/A or #NUM! for these rows shows ✘.

**6. DENIALRATE on the ClaimStatus column**

- **Answer:** 7.7%
- **Solution:**

```vba
' In the yellow cell: =DENIALRATE(tblClaims[ClaimStatus])
' (or =DENIALRATE(Claims!H2:H1094))

' Task 6: share of non-blank cells in statusRange that equal statusToCount.
Public Function DENIALRATE(statusRange As Range, Optional statusToCount As String = "Denied") As Variant
    Dim cell As Range, total As Long, hits As Long
    For Each cell In statusRange.Cells
        If Not IsEmpty(cell.Value) And Not IsError(cell.Value) Then
            total = total + 1
            If StrComp(Trim$(CStr(cell.Value)), statusToCount, vbTextCompare) = 0 Then hits = hits + 1
        End If
    Next cell
    If total = 0 Then
        DENIALRATE = CVErr(xlErrDiv0)      ' nothing to divide by: #DIV/0!
    Else
        DENIALRATE = hits / total
    End If
End Function
```


A typed Optional argument gets its default in the declaration (`= "Denied"`), so you don't need IsMissing. Declaring `statusRange As Range` lets you loop over its cells, and it means Excel returns #VALUE! by itself if someone passes a plain number. 84 of 1,093 claims are Denied. Because statusToCount is a parameter, the same function answers other questions too: `=DENIALRATE(tblClaims[ClaimStatus],"Pending")` gives the pending rate. The key's live cell uses COUNTIF/COUNTA because your UDF isn't in the downloaded .xlsx.

**7. Snippet A: ByVal vs ByRef**

- **Answer:** 125
- **Solution:** Trace it: `AddFeeByVal charge` adds 25 to a **copy** (charge stays 100). `AddFeeByRef charge` adds 25 to the caller's variable itself (125). `AddFeeByRef (charge)` looks the same, but the parentheses turn `charge` into an expression, so VBA passes a temporary copy and charge stays **125**.

VBA passes arguments **ByRef by default**, so a procedure can change the caller's variable. Write `ByVal` when a procedure must not do that. The parentheses trap catches experienced developers: when you call a Sub without `Call`, wrapping one argument in parentheses evaluates it first and passes the result, which silently turns ByRef into ByVal. Call Subs without parentheses, or use `Call AddFeeByRef(charge)`.

**8. CountDenialReasons (Collection)**

- **Answer:** 7
- **Solution:**

```vba
' Task 8: how many different denial reasons? A Collection refuses a duplicate
' key, so adding every reason with Key:=reason keeps exactly one of each.
Public Sub CountDenialReasons()
    Dim ws As Worksheet, claims As Variant
    Dim reasonCol As Long, r As Long, reason As String
    Dim reasons As Collection

    Set ws = ThisWorkbook.Worksheets("Claims")
    claims = ws.Range("A1").CurrentRegion.Value        ' header + data in one 2-D array
    reasonCol = HeaderColumn(ws, "DenialReason")
    Set reasons = New Collection

    For r = 2 To UBound(claims, 1)
        reason = Trim$(CStr(claims(r, reasonCol)))     ' a blank cell becomes ""
        If Len(reason) > 0 Then AddIfNew reasons, reason
    Next r

    ThisWorkbook.Worksheets("Output").Range("B4").Value = reasons.Count
End Sub

' Adds keyText to a Collection once. Returns False if it was already there.
Private Function AddIfNew(col As Collection, ByVal keyText As String) As Boolean
    On Error Resume Next
    col.Add Item:=keyText, Key:=keyText    ' error 457 if the key already exists
    AddIfNew = (Err.Number = 0)
    On Error GoTo 0
End Function
```


A Collection key must be unique, so adding a duplicate key raises run-time error 457. `AddIfNew` traps that error with `On Error Resume Next`, reads `Err.Number`, and switches trapping off again straight away with `On Error GoTo 0`. The 7 reasons are: Authorization Required, Coding Error, Duplicate Claim, Eligibility / Coverage, Medical Necessity, Missing Documentation, Timely Filing. Skipping blanks matters: most claims have no denial reason, and counting the empty string as a reason would give 8.

**9. CountDistinctPatients (Dictionary)**

- **Answer:** 631
- **Solution:**

```vba
' Task 9: how many different patients have a claim? Dictionary keys are unique.
Public Sub CountDistinctPatients()
    Dim ws As Worksheet, claims As Variant
    Dim patientCol As Long, r As Long
    Dim patients As Object                             ' late-bound Scripting.Dictionary

    Set ws = ThisWorkbook.Worksheets("Claims")
    claims = ws.Range("A1").CurrentRegion.Value
    patientCol = HeaderColumn(ws, "PatientID")
    Set patients = CreateObject("Scripting.Dictionary")

    For r = 2 To UBound(claims, 1)
        If Not patients.Exists(claims(r, patientCol)) Then
            patients.Add claims(r, patientCol), 1
        End If
    Next r

    ThisWorkbook.Worksheets("Output").Range("B5").Value = patients.Count
End Sub
```


`.Exists` asks whether a key is already in the dictionary, so each PatientID is added once and `.Count` is the number of different patients. 1,093 claims come from 631 patients because 265 patients had more than one claim. `Dim patients As Object` with `CreateObject` is **late binding**: it needs no reference to the Scripting Runtime library, so the file works on any Windows PC. In Microsoft 365 the key's `=ROWS(UNIQUE(…))` gives the same answer, but the dictionary pattern scales to jobs a formula can't do, like the bonus.

**10. TopDeniedPayer (two dictionaries)**

- **Answer:** Silverline Medicare Advantage (22 denied claims)
- **Solution:**

```vba
' Task 10: which payer has the most Denied claims? Two dictionaries:
' one to look up payer names, one to count.
Public Sub TopDeniedPayer()
    Dim wsClaims As Worksheet, claims As Variant, payers As Variant
    Dim payerCol As Long, statusCol As Long, r As Long
    Dim payerNames As Object, deniedCount As Object
    Dim payerID As Variant, bestID As String, bestCount As Long

    ' 1. Lookup dictionary: PayerID -> PayerName
    Set payerNames = CreateObject("Scripting.Dictionary")
    payers = ThisWorkbook.Worksheets("Payers").Range("A1").CurrentRegion.Value
    For r = 2 To UBound(payers, 1)
        payerNames(payers(r, 1)) = payers(r, 2)        ' assigning to a new key adds it
    Next r

    ' 2. Counting dictionary: PayerID -> number of Denied claims
    Set wsClaims = ThisWorkbook.Worksheets("Claims")
    claims = wsClaims.Range("A1").CurrentRegion.Value
    payerCol = HeaderColumn(wsClaims, "PayerID")
    statusCol = HeaderColumn(wsClaims, "ClaimStatus")
    Set deniedCount = CreateObject("Scripting.Dictionary")
    For r = 2 To UBound(claims, 1)
        If claims(r, statusCol) = "Denied" Then
            ' Reading a key that isn't there returns Empty (and adds it); Empty + 1 = 1
            deniedCount(claims(r, payerCol)) = deniedCount(claims(r, payerCol)) + 1
        End If
    Next r

    ' 3. Keep the largest count (on a tie, the payer seen first wins)
    For Each payerID In deniedCount.Keys
        If deniedCount(payerID) > bestCount Then
            bestCount = deniedCount(payerID)
            bestID = payerID
        End If
    Next payerID

    With ThisWorkbook.Worksheets("Output")
        .Range("B6").Value = payerNames(bestID)
        .Range("B7").Value = bestCount
    End With
End Sub
```


`deniedCount(id) = deniedCount(id) + 1` is the **counting pattern**: reading a key that isn't there yet returns Empty (and quietly adds the key), and Empty + 1 = 1. Then one loop over `.Keys` keeps the largest count. The lookup dictionary replaces a VLOOKUP inside the loop. Note that the payer with the most claims overall is Medicare (270 claims), but the most *denials* come from Silverline Medicare Advantage: 22 of its 180 claims.

**11. Snippet B: On Error GoTo + Resume Next**

- **Answer:** 303
- **Solution:** CLng fails (run-time error 13, Type mismatch) on "abc" and on the empty string. Each time, VBA jumps to BadReading, prints a line, and `Resume Next` continues with the statement **after** the one that failed, which is `Next i`. So the bad items add nothing: 120 + 95 + 88 = **303**.

`On Error GoTo label` sends any run-time error to the handler, and the `Exit Sub` before the label keeps normal runs out of it. Inside the handler, `Err.Number` says what went wrong (13 = Type mismatch). `Resume Next` skips the failing statement, `Resume` retries it (an endless loop here, because "abc" never becomes a number), and `Resume SomeLabel` continues at a label.

**12. Snippet C: the On Error Resume Next trap**

- **Answer:** 518
- **Solution:** When `CLng("abc")` fails, the assignment never happens, so `reading` still holds 120 from the line before, and it's added again. The same thing happens for "" (reading is still 95). 120 + 120 + 95 + 95 + 88 = **518**.

`On Error Resume Next` doesn't fix an error. It hides it. The variable keeps its previous value and the code carries on with wrong data: here 215 phantom units of glucose. Keep Resume Next to the one line you expect might fail, check `Err.Number` right after it, and switch back with `On Error GoTo 0`. AddIfNew in task 8 follows that pattern.

**13. Snippet D: Application.Match vs WorksheetFunction.Match**

- **Answer:** 1004
- **Solution:** `Application.Match` returns the error **value** Error 2042 (#N/A), which `IsError` detects, so the code prints "Application.Match: not found". `WorksheetFunction.Match` raises a run-time **error** instead: **1004**, *Unable to get the Match property of the WorksheetFunction class*.

Both call Excel's MATCH, but they report failure differently. Use `Application.Match` (result in a Variant, then `IsError`) when "not found" is a normal outcome, as in the HeaderColumn helper. `WorksheetFunction.Match` only makes sense when a miss is a real error that your handler should catch. The same split applies to VLOOKUP, INDEX, and the other lookup functions.

</details>
<!-- END GENERATED: answers -->

## 🏆 Bonus challenge

<!-- BEGIN GENERATED: bonus -->
Cedar Ridge's revenue-cycle director wants a payer scorecard she can rebuild with one click every month. Write BuildPayerSummary (start from starter/PayerSummary.bas) so that it does five things:
1. If the Claims or Payers sheet is missing, it shows a friendly message and stops without a run-time error.
2. It reads the Claims table into an array and uses a Scripting.Dictionary keyed by PayerID (Mac: a Collection plus arrays) to work out four numbers for each payer: Claims (number of rows), Billed (sum of BilledAmount), Paid (sum of PaidAmount), and Denied (rows with ClaimStatus "Denied").
3. It skips, and counts, every row whose BilledAmount or PaidAmount is not a number. A blank cell is not a number, and a skipped row is left out of every column.
4. It writes a sheet named PayerSummary, clearing it first if it already exists. Row 1 holds the headers PayerID, PayerName, Claims, Billed, Paid, Denied, DenialRate (Denied ÷ Claims). Below it go one row per payer, sorted by Billed from largest to smallest, and then a TOTAL row: the word TOTAL in column A, the totals in C:F, and the overall rate in G. Two rows below TOTAL, column A holds the text Rows skipped and column B holds the count.
5. It restores ScreenUpdating and Calculation even when an error stops the macro.

Work on the **Bonus** sheet of the workbook.

- **B1.** Run BuildPayerSummary. Which payer has the third-highest billed charges? (It's the PayerName in cell B4 of PayerSummary, and the gray cell reads it.) *(Hint: Range.Sort with Key1:=the Billed header cell, Order1:=xlDescending, Header:=xlYes)*
- **B2.** What DenialRate does your PayerSummary show for Silverline Medicare Advantage? (The gray cell finds Silverline's row on PayerSummary and reads column G.) *(Hint: The array-in-a-dictionary pattern: copy it out, change it, put it back)*
- **B3.** What total Paid does the TOTAL row of PayerSummary show? (The gray cell reads column E of the TOTAL row, and the check compares it to the cent.) *(Hint: Accumulate the totals in the same loop that fills the output array)*
- **B4.** How many claim rows did your macro skip because BilledAmount or PaidAmount was not a number? (The gray cell reads the number next to 'Rows skipped'.) *(Hint: A blank cell read into an array is Empty, and VarType(Empty) is vbEmpty)*
- **B5.** Test the error handling. (1) Rename the Claims sheet to Claims_old and run BuildPayerSummary: you should get your friendly message, not a run-time error. (2) Rename it back. (3) Temporarily change the header text in the Claims sheet's BilledAmount cell (E1) to Billed and run the macro again: the Fail handler should report the missing column, and Formulas → Calculation Options should still show Automatic afterwards. (4) Put the header back and run the macro once more. *(Hint: Check before you change settings, and clean up in one place)*
<!-- END GENERATED: bonus -->

The reference solutions are in [`solutions/PayerSummary_Solution.bas`](solutions/PayerSummary_Solution.bas) (Windows,
Scripting.Dictionary) and [`solutions/PayerSummaryMac_Solution.bas`](solutions/PayerSummaryMac_Solution.bas) (Mac,
Collections only, and it also runs on Windows). Both are spoilers.

<!-- BEGIN GENERATED: bonus-answers -->
<details>
<summary><b>🔑 Show the bonus solution</b> — Give it a real try first!</summary>

**B1. Third-highest billed payer**

- **Answer:** Keystone Health Partners
- **Solution:**

```vba
Public Sub BuildPayerSummary()
    Dim wsClaims As Worksheet, wsOut As Worksheet
    Dim claims As Variant, payers As Variant, outData() As Variant
    Dim payerNames As Object, stats As Object
    Dim colPayer As Long, colBilled As Long, colPaid As Long, colStatus As Long
    Dim r As Long, n As Long, totalRow As Long, skipped As Long
    Dim payerID As Variant, s As Variant
    Dim totClaims As Long, totDenied As Long, totBilled As Double, totPaid As Double
    Dim prevCalc As Long

    ' 1. Check first: a friendly message beats a crash
    If Not SheetExists("Claims") Or Not SheetExists("Payers") Then
        MsgBox "BuildPayerSummary needs a sheet named 'Claims' and a sheet named 'Payers'." & vbNewLine & _
               "One of them is missing or has been renamed. Fix that, then run the macro again.", _
               vbExclamation, "Payer summary"
        Exit Sub
    End If

    ' 2. From here on, any run-time error jumps to Fail, which restores Excel's settings
    prevCalc = Application.Calculation
    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set wsClaims = ThisWorkbook.Worksheets("Claims")
    claims = wsClaims.Range("A1").CurrentRegion.Value2     ' Value2: plain numbers, no Currency/Date types
    colPayer = HeaderColumn(wsClaims, "PayerID")
    colBilled = HeaderColumn(wsClaims, "BilledAmount")
    colPaid = HeaderColumn(wsClaims, "PaidAmount")
    colStatus = HeaderColumn(wsClaims, "ClaimStatus")
    payers = ThisWorkbook.Worksheets("Payers").Range("A1").CurrentRegion.Value2

    ' 3. Lookup dictionary: PayerID -> PayerName
    Set payerNames = CreateObject("Scripting.Dictionary")
    For r = 2 To UBound(payers, 1)
        payerNames(payers(r, 1)) = payers(r, 2)
    Next r

    ' 4. One pass over the claims: PayerID -> Array(claims, billed, paid, denied)
    Set stats = CreateObject("Scripting.Dictionary")
    For r = 2 To UBound(claims, 1)
        If IsNumberValue(claims(r, colBilled)) And IsNumberValue(claims(r, colPaid)) Then
            payerID = claims(r, colPayer)
            If Not stats.Exists(payerID) Then stats.Add payerID, Array(0, 0, 0, 0)
            s = stats(payerID)                 ' copy the array out of the dictionary...
            s(0) = s(0) + 1
            s(1) = s(1) + claims(r, colBilled)
            s(2) = s(2) + claims(r, colPaid)
            If claims(r, colStatus) = "Denied" Then s(3) = s(3) + 1
            stats(payerID) = s                 ' ...and put it back: stats(payerID)(0) = 1 would not stick
        Else
            skipped = skipped + 1              ' blank or text amount
        End If
    Next r
    n = stats.Count
    If n = 0 Then Err.Raise vbObjectError + 514, "BuildPayerSummary", "No claim rows with numeric amounts were found."

    ' 5. Header + one row per payer, built in memory, plus the running totals
    ReDim outData(1 To n + 1, 1 To 7)
    outData(1, 1) = "PayerID": outData(1, 2) = "PayerName": outData(1, 3) = "Claims"
    outData(1, 4) = "Billed": outData(1, 5) = "Paid": outData(1, 6) = "Denied": outData(1, 7) = "DenialRate"
    r = 1
    For Each payerID In stats.Keys
        r = r + 1
        s = stats(payerID)
        outData(r, 1) = payerID
        If payerNames.Exists(payerID) Then
            outData(r, 2) = payerNames(payerID)
        Else
            outData(r, 2) = "(unknown payer)"
        End If
        outData(r, 3) = s(0): outData(r, 4) = s(1): outData(r, 5) = s(2): outData(r, 6) = s(3)
        outData(r, 7) = s(3) / s(0)
        totClaims = totClaims + s(0): totBilled = totBilled + s(1)
        totPaid = totPaid + s(2): totDenied = totDenied + s(3)
    Next payerID

    ' 6. A fresh output sheet: clear it if it exists, add it if it doesn't
    If SheetExists(OUT_SHEET) Then
        Set wsOut = ThisWorkbook.Worksheets(OUT_SHEET)
        wsOut.Cells.Clear
    Else
        Set wsOut = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets("Payers"))
        wsOut.Name = OUT_SHEET
    End If

    ' 7. Write the block in one go, sort the payer rows by Billed (largest
    '    first), then add the TOTAL row and the skipped-rows line
    With wsOut
        .Range("A1").Resize(n + 1, 7).Value = outData
        .Range("A1").Resize(n + 1, 7).Sort Key1:=.Range("D1"), Order1:=xlDescending, Header:=xlYes
        totalRow = n + 2
        .Cells(totalRow, 1).Value = "TOTAL"
        .Cells(totalRow, 3).Value = totClaims
        .Cells(totalRow, 4).Value = totBilled
        .Cells(totalRow, 5).Value = totPaid
        .Cells(totalRow, 6).Value = totDenied
        .Cells(totalRow, 7).Value = totDenied / totClaims
        .Cells(totalRow + 2, 1).Value = "Rows skipped"
        .Cells(totalRow + 2, 2).Value = skipped
        .Range("D:E").NumberFormat = "#,##0.00"
        .Range("G:G").NumberFormat = "0.0%"
        .Rows(1).Font.Bold = True
        .Rows(totalRow).Font.Bold = True
        .Columns("A:G").AutoFit
    End With

CleanExit:
    On Error Resume Next                   ' cleanup must never raise an error of its own
    Application.Calculation = prevCalc
    Application.ScreenUpdating = True
    Exit Sub

Fail:
    MsgBox "BuildPayerSummary stopped: " & Err.Description & " (error " & Err.Number & ")", _
           vbCritical, "Payer summary"
    Resume CleanExit
End Sub

' ---------------------------------------------------------------------
' Helpers
' ---------------------------------------------------------------------

' True if this workbook has a sheet with that name.
Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)    ' error 9 if there is no such sheet
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

' Column number of a header in row 1. Raises a clear error if it is missing.
Private Function HeaderColumn(ws As Worksheet, ByVal headerText As String) As Long
    Dim pos As Variant
    pos = Application.Match(headerText, ws.Rows(1), 0)
    If IsError(pos) Then
        Err.Raise vbObjectError + 513, "HeaderColumn", _
                  "Column '" & headerText & "' was not found on sheet '" & ws.Name & "'."
    End If
    HeaderColumn = CLng(pos)
End Function

' True for real numbers (a blank cell or text is not a number).
Private Function IsNumberValue(v As Variant) As Boolean
    Select Case VarType(v)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte
            IsNumberValue = True
        Case Else
            IsNumberValue = False
    End Select
End Function
```


The dictionary loop collects the totals in whatever order payers first appear. Sorting the written block, header excluded, puts the biggest payers on top: Medicare, Silverline Medicare Advantage, Keystone Health Partners. Third place is close (Keystone Health Partners billed 2,552,628.44 against 2,544,075.24 for State Medicaid), so an unsorted or wrongly sorted table shows ✘. The macro writes the whole block with one `.Value = outData` assignment, which is much faster than writing cell by cell. **Mac:** the Collection version is in solutions/PayerSummaryMac_Solution.bas.

**B2. Silverline Medicare Advantage denial rate**

- **Answer:** 11.3%
- **Solution:** Read it from PayerSummary column G. Formula cross-check (blank and text amounts excluded): `=COUNTIFS(tblClaims[PayerID],"PY02",tblClaims[ClaimStatus],"Denied",tblClaims[BilledAmount],">=0")/COUNTIFS(tblClaims[PayerID],"PY02",tblClaims[BilledAmount],">=0")`

Silverline has 20 denied claims out of 177 kept rows. One Silverline denial has a blank BilledAmount. `IsNumeric(Empty)` returns True, so a macro that tests amounts with IsNumeric keeps that row and reports 11.8% instead of 11.3%. Test the type with `VarType` instead, as IsNumberValue does. Also watch the array trap: `stats(id)(0) = stats(id)(0) + 1` changes a temporary copy and the dictionary never sees it. Copy the array out, change it, and assign it back.

**B3. Total Paid on the TOTAL row**

- **Answer:** 4,440,116.72
- **Solution:** Read it from the TOTAL row, column E. Formula cross-check: `=SUMIFS(tblClaims[PaidAmount],tblClaims[BilledAmount],">=0")`

The macro adds each payer's figures to running totals while it builds the output array, then writes the TOTAL row after sorting, so the sort can't move it. It covers 1,088 claims and \$16,722,301.02 billed. Writing the TOTAL row before sorting is a classic bug: Range.Sort would treat it as a payer and sort it to the top.

**B4. Rows skipped**

- **Answer:** 5
- **Solution:** Read it from the cell next to *Rows skipped*. Formula cross-check: `=ROWS(tblClaims[BilledAmount])-COUNT(tblClaims[BilledAmount])`

Filter the Claims sheet's BilledAmount column to find them: two are blank and three hold text (N/A, VOID). Without the check, `s(1) + "N/A"` raises run-time error 13 (Type mismatch), the handler stops the macro, and nobody gets a scorecard. Counting and reporting skipped rows is better than silently dropping them, because the director can see that 5 claims need fixing at the source.

**B5. Robustness test**

- **Solution:**

1. With Claims renamed, `SheetExists("Claims")` is False, so the macro shows the friendly MsgBox and exits **before** it changes any setting.
2. With the header renamed, `HeaderColumn` raises error vbObjectError + 513. HeaderColumn has no handler of its own, so the error travels up to BuildPayerSummary's `Fail` handler, which shows *Column 'BilledAmount' was not found on sheet 'Claims'*. Then `Resume CleanExit` restores Calculation and ScreenUpdating.
3. If Calculation Options shows Manual after the error, your error path skipped the cleanup block. Set it back with **Formulas → Calculation Options → Automatic**, then route the handler through `Resume CleanExit`. (Excel usually switches ScreenUpdating back on by itself when a macro ends, but the calculation mode stays manual for every open workbook until someone changes it.)


Good error handling has two layers. **Validate** the things you can predict (a missing sheet) and explain them in plain language. **Trap** everything else in one handler that tells the user what happened and always runs the cleanup code. Because errors travel up the call stack until they meet a handler, the custom error raised in a helper function is reported by the main macro's handler.

</details>
<!-- END GENERATED: bonus-answers -->

## Key takeaways

- A UDF is a Public Function in a standard module. It returns a value and changes nothing else, and it recalculates
  only when its arguments change, so pass every input as an argument.
- Arguments are ByRef unless you write ByVal. Optional arguments come last: use a Variant with `IsMissing`, or a typed
  argument with a default. Parentheses around a single Sub argument pass a copy.
- Declare the return type as Variant and return `CVErr(xlErrValue)`, `xlErrNum`, `xlErrNA`, or `xlErrDiv0` for bad
  input, because an error can't be mistaken for real data.
- Check types with `VarType`, not `IsNumeric`, because `IsNumeric(Empty)` is True.
- A Collection is built in and works on a Mac. `Scripting.Dictionary` (Windows) adds `.Exists`, `.Keys`, and items you
  can change. Almost every job is a distinct list, a count, a sum, or a lookup.
- Validate what you can predict, trap the rest with `On Error GoTo Fail`, and send every exit through one cleanup block
  that restores Excel's settings. Keep `On Error Resume Next` to a single statement.

<!-- BEGIN GENERATED: nav -->
---

⬅️ **Previous:** [5.3 VBA: Ranges, Worksheets & Workbooks](../03-vba-ranges-worksheets/README.md) · 🏠 [Course home](../../README.md) · **Next:** [5.5 Events, UserForms & Automated Reports](../05-events-userforms-automation/README.md) ➡️
<!-- END GENERATED: nav -->
