# Lesson 5.2 · VBA Fundamentals

> **Level:** Expert · **Time:** about 60 minutes · **Workbook:** [`5.2-vba-fundamentals.xlsx`](5.2-vba-fundamentals.xlsx)
> **Data:** 505 lab results from Bluestone's three intensive care units (Bluestone Memorial, Ashby Falls, Cedar Ridge), collected January–April 2025 and listed in the order the results were released.

In Lesson 5.1 the macro recorder wrote code for you. The recorder can repeat your clicks, but it can't make a decision
(*"is this potassium critical?"*) and it can't repeat a step for every row (*"check all 505 results"*). For that you write
**VBA** (Visual Basic for Applications) yourself. Picture a lab supervisor who scrolls through every STAT result each
morning to find the ones that missed the 60-minute turnaround target. That takes the better part of an hour. A 20-line macro
does the same job in under a second, flags every late result, and never skips a row because the phone rang. Here you write
your first real VBA against three ICUs' lab results: variables, decisions, loops, and the debugging tools you'll reach for
every time something doesn't work.

## What you'll learn

- Navigate the Visual Basic Editor and organize code in modules
- Declare variables with the right data types and Option Explicit
- Control flow with If, Select Case, For, For Each, and Do loops
- Debug with breakpoints, stepping, the Immediate window, and Debug.Print

## 📖 Guide

### 1. Set up your workbook

You need three files from this lesson's folder: the workbook and the two starter modules
[`starter/LabMacros.bas`](starter/LabMacros.bas) and [`starter/Snippets.bas`](starter/Snippets.bas). On GitHub, open each
`.bas` file and click **Download raw file**.

1. Open `5.2-vba-fundamentals.xlsx` and save it as a macro-enabled workbook: **File → Save As**, then choose
   **Excel Macro-Enabled Workbook (\*.xlsm)** as the file type. An `.xlsx` file can't store macros, so Excel would throw your
   code away when you close it.
2. Open the **Visual Basic Editor** (the **VBE**): press **Alt + F11** (Mac: **Option + F11**), or click
   **Developer → Visual Basic**. To get back to Excel, press the same keys again or click the Excel window.
3. In the Project Explorer (**Ctrl + R**; Mac: **View → Project Explorer**), click **VBAProject (5.2-vba-fundamentals.xlsm)**.
   The VBE imports into whichever project is selected. If you made a Personal Macro Workbook in Lesson 5.1,
   **VBAProject (PERSONAL.XLSB)** is listed too, so make sure you've selected this workbook's project. Then choose
   **File → Import File…** (Windows shortcut: **Ctrl + M**) and pick `LabMacros.bas`. Repeat for `Snippets.bas`. Both appear
   in the **Modules** folder of this workbook's project.
4. Double-click **LabMacros**, click anywhere inside `Sub Warmup()`, and press **F5** (Mac: **Run → Run Sub/UserForm**).
5. Switch to Excel and look at the **Output** sheet. Cell B5 now holds the number of lab rows. That's Practice task 1 done.
6. Back in the VBE, open the **Immediate window** (**Ctrl + G**; Mac: **View → Immediate Window**). Warmup also printed a line
   there with `Debug.Print`. Keep this window open, because the snippets in tasks 2–7 print their results there.

> ⚠️ **Undo can't undo a macro.** Ctrl + Z does nothing after a macro changes cells, so save before you run new code. If a macro
> wrecks your data, close without saving and reopen.

> 📋 **Version note:** VBA runs in the desktop versions of Excel for Windows and Excel for Mac. Excel for the web,
> iPad, iPhone, and Android can open an `.xlsm` file but can't run or edit its macros. Lesson 5.5 shows the web alternative,
> Office Scripts. If Excel shows a yellow **Enable Content** bar or a red **Security Risk** bar, Lesson 5.1 explains macro security.

### 2. A tour of the Visual Basic Editor

The VBE is a separate window with several panes. If one is missing, open it from the **View** menu.

| Pane | What it shows | Windows | Mac |
|---|---|---|---|
| **Project Explorer** | Every open workbook (each one is a *VBA project*) with its sheets, `ThisWorkbook`, and modules | Ctrl + R | View → Project Explorer |
| **Properties** | Settings of the selected item, such as a module's `(Name)` | F4 | View → Properties Window |
| **Code window** | The code of the item you double-clicked | F7, or double-click | double-click |
| **Immediate window** | What `Debug.Print` writes. You can also type a line here and press Enter to run it | Ctrl + G | View → Immediate Window |
| **Locals window** | Every variable in the paused procedure and its current value | View → Locals Window | View → Locals Window |
| **Watch window** | Expressions you asked the VBE to keep an eye on | View → Watch Window | View → Watch Window |

A typical layout puts the Project Explorer at top left, Properties below it, the Code window on the right, and the Immediate
window along the bottom. Drag a pane's title bar to dock it somewhere else.

The editor helps you as you type:

- **IntelliSense.** After you type `ws.` a list of everything a worksheet can do pops up. Press **Tab** to accept an item.
  **Ctrl + Space** completes a half-typed name, and **Ctrl + J** lists members on demand (Windows).
- **Auto-capitalization.** When you press Enter, the VBE rewrites keywords and declared variable names in their proper case.
  Type variable names in lowercase. If `criticalcount` doesn't change to `criticalCount`, you've misspelled it.
- **Color coding.** Keywords are blue, comments are green, and a line with a syntax error turns red.

> 💡 **Tip:** In **Tools → Options → Editor**, turn on **Require Variable Declaration** (section 5 explains why) and consider
> turning off **Auto Syntax Check**. Bad lines still turn red, but the VBE stops interrupting you with a pop-up every time you
> leave a half-finished line.

### 3. Where code lives: modules

A **module** is a container for code inside a workbook. There are several kinds, and code only works where it belongs.

| Container | How you get one | What goes in it |
|---|---|---|
| **Standard module** (`Module1`, `LabMacros`) | **Insert → Module** | Your macros. Everything in this lesson lives here |
| **Sheet module** (one per sheet, listed under *Microsoft Excel Objects* with the tab name in parentheses) | Double-click a sheet in the Project Explorer | Code that reacts to events on that sheet, such as a cell changing (Lesson 5.5) |
| **ThisWorkbook** | Double-click `ThisWorkbook` | Code that reacts to workbook events, such as opening or saving (Lesson 5.5) |
| **UserForm** | **Insert → UserForm** | Custom dialog boxes (Lesson 5.5) |

Keep your modules organized:

- **Rename** a module by selecting it and changing `(Name)` in the Properties window. Use names like `LabMacros` or
  `ReportTools`. Don't give a module the same name as a procedure, because VBA then can't tell which one you mean.
- **Export** a module (right-click → **Export File…**) to save it as a `.bas` text file. That's how you back up code, share it,
  or move it to another workbook. **Import** reverses it.
- **Remove** a module with right-click → **Remove…**. The VBE offers to export it first.

### 4. Sub procedures

A **procedure** is a named block of code. A **Sub** procedure performs actions. You run it, and it does its work.

```vba
Sub CountCriticals()                 ' the name you run
    ' statements go here, one per line
End Sub
```

- **Names** start with a letter and contain only letters, digits, and underscores. No spaces, and no VBA keywords such as `Next`.
- **Comments** start with an apostrophe `'`. VBA ignores everything after it on that line.
- A long line can continue on the next one if you end it with a space and an underscore ` _`.

There are several ways to run a Sub:

| From | How |
|---|---|
| The VBE | Click inside the Sub and press **F5** (Mac: **Run → Run Sub/UserForm** or the ▶ button) |
| Excel | **Alt + F8** (Mac: **Option + F8**) or **Developer → Macros**, pick the macro, click **Run** |
| Another procedure | Write its name on a line of its own |
| The Immediate window | Type its name and press Enter |

One Sub can call others, which lets you build a big job out of small, testable pieces:

```vba
Sub MorningChecks()
    CountCriticals          ' runs CountCriticals, then comes back here
    FlagSlowStat
End Sub
```

A Sub can also take **arguments**, which are values it needs to do its job. Declare them in the parentheses:

```vba
Sub WriteCountAbove(testCode As String, limit As Double)
    ' ... count the results of testCode above limit ...
End Sub

Sub CheckCreatinine()
    WriteCountAbove "CREAT", 2           ' call it with no parentheses
    Call WriteCountAbove("CREAT", 2)     ' or use Call WITH parentheses
End Sub
```

When `CheckCreatinine` calls it, the names `testCode` and `limit` inside `WriteCountAbove` hold "CREAT" and 2 for that run.

> ⚠️ `WriteCountAbove("CREAT", 2)` on its own line gives **Compile error: Expected: =**. Without `Call`, VBA reads parentheses
> around several arguments as the start of a function call whose result you forgot to store. Drop the parentheses or add `Call`.

> 📋 A Sub that takes arguments doesn't appear in the **Macros** dialog, because Excel has no way to ask you for the
> arguments. A `Private Sub` is hidden too. Run them from another procedure or from the Immediate window.

### 5. Variables, data types, and Option Explicit

A **variable** is a named place to store a value while the macro runs. You create one with `Dim` and give it a **data
type**, which says what kind of value it holds.

```vba
Dim r As Long                 ' a row number
Dim total As Double           ' a running total of ResultValue
Dim testCode As String        ' a code such as "GLU"
Dim collected As Date         ' a date and time
Dim found As Boolean          ' True or False
Dim ws As Worksheet           ' an object: a whole worksheet

r = 2                                      ' ordinary values: just =
Set ws = ThisWorkbook.Worksheets("Labs")   ' objects need Set
```

| Type | Holds | Range | Starts as | Use it for |
|---|---|---|---|---|
| `Long` | Whole numbers | about ±2.1 billion | 0 | Counts and row numbers. Your default for whole numbers |
| `Integer` | Whole numbers | −32,768 to 32,767 | 0 | Almost nothing. It overflows on anything big |
| `Double` | Decimal numbers | about 15 significant digits | 0 | Lab values, averages, ratios |
| `Currency` | Money, exactly 4 decimal places | about ±922 trillion | 0 | Charges and payments, with no rounding drift |
| `String` | Text | up to about 2 billion characters | `""` | Codes, names, IDs |
| `Date` | A date and time | 1/1/100 to 12/31/9999 | 12:00:00 AM | Timestamps such as CollectedDateTime |
| `Boolean` | `True` or `False` | | `False` | Yes/no flags such as `found` |
| `Variant` | Anything | | `Empty` | Values whose type you can't know in advance |
| `Worksheet`, `Range`, `Workbook` | A reference to one kind of Excel object | | `Nothing` | Sheets and cells. Assign with `Set` |
| `Object` | A reference to any kind of object | | `Nothing` | Rarely needed here. A specific type such as `Worksheet` catches mistakes and gives you IntelliSense |

> ⚠️ `Dim r, n As Long` declares **only `n`** as a Long. `r` silently becomes a Variant. Give every variable its own type:
> `Dim r As Long, n As Long`.

**Option Explicit.** Put `Option Explicit` on the first line of every module. It tells VBA to refuse to run any code that uses
a variable you haven't declared. Instead, VBA stops with **Compile error: Variable not defined** and highlights the name.
Without it, a misspelled name silently creates a brand-new empty variable, and your macro gives a wrong answer without any
error at all. Snippet A in the practice shows this.

Turn on **Tools → Options → Editor → Require Variable Declaration** so the VBE types `Option Explicit` into every *new* module
for you. It doesn't change existing modules, so add the line to those by hand. On a Mac the same check box is in the VBE's
options, and typing `Option Explicit` yourself does exactly the same job on any version.

**Constants.** A **constant** is a named value that can't change while the macro runs. Use one instead of a "magic number"
buried in the code:

```vba
Const STAT_TARGET_MIN As Long = 60      ' STAT results should be back within 60 minutes
If minutes > STAT_TARGET_MIN Then ...
```

**Scope.** Where you declare a variable decides who can see it and how long it lives:

| Declared | Visible to | Keeps its value |
|---|---|---|
| `Dim` inside a Sub | That Sub only | Until the Sub ends. Each run starts fresh |
| `Dim` or `Private` at the top of a module (below `Option Explicit`) | Every procedure in that module | Between runs, until the VBA project is reset |
| `Public` at the top of a module | Every module in the workbook | Between runs, until the VBA project is reset |

> 💡 **Tip:** Declare variables inside the procedure that uses them. Module-level variables remember values from the last
> run, and a counter that starts at 37 instead of 0 is a hard bug to spot.

### 6. Operators and expressions

| Operator | Meaning | Example | Result |
|:-:|---|---|---|
| `+ - * /` | Arithmetic | `135 / 60` | 2.25 |
| `\` | **Integer division**: divide and drop the remainder | `135 \ 60` | 2 |
| `Mod` | **Remainder** after integer division | `135 Mod 60` | 15 |
| `^` | Power | `2 ^ 3` | 8 |
| `&` | Join text | `"LAB" & 725446` | `"LAB725446"` |
| `= <> < > <= >=` | Compare. The result is True or False | `4.6 > 4` | True |
| `And`, `Or`, `Not` | Combine True/False values | `flag = "HH" Or flag = "LL"` | True if either is true |

VBA works through an expression in this order: `^`, then negation, then `*` and `/`, then `\`, then `Mod`, then `+` and `-`,
then `&`, then comparisons, then `Not`, `And`, `Or`. So in `135 \ 60 & " h"` the division happens first and the join last.
When in doubt, add parentheses.

The `=` sign does two jobs. On a line of its own, `x = 5` **assigns** 5 to `x`. Inside `If x = 5 Then` it **compares**.

*Worked example.* A STAT lactate took 135 minutes from collection to result. Turn that into hours and minutes:

```vba
Dim tat As Long
tat = 135
Debug.Print tat \ 60 & " h " & tat Mod 60 & " min"     ' prints: 2 h 15 min
```

> ⚠️ **Join text with `&`, not `+`.** `"5" + 2` gives 7 because VBA turns the text into a number, while `"5" & 2` gives
> `"52"`. Using `&` always means "join."

> ⚠️ **`And` and `Or` always evaluate both sides.** Some languages stop as soon as the answer is known. VBA doesn't. So
> `If r > 1 And ws.Cells(r, 5).Value > 180 Then` still reads the cell when `r` is 1. When the second test could fail on
> its own, nest it inside a separate `If`.

> ⚠️ `\` and `Mod` first round both numbers to whole numbers, so use them with whole numbers.

VBA looks like the worksheet, but a few rules differ. These differences cause real bugs:

| Expression | In VBA | On the worksheet |
|---|---|---|
| Round a half | `Round(2.5)` gives **2**. VBA uses *banker's rounding*, so halves go to the nearest even number | `=ROUND(2.5,0)` gives **3** |
| Remainder of a negative | `-7 Mod 3` gives **−1** | `=MOD(-7,3)` gives **2** |
| Negative base with a power | `-2 ^ 2` gives **−4** (power first) | `=-2^2` gives **4** (negation first) |
| Compare text | `"GLU" = "Glu"` is **False** (case-sensitive) | `="GLU"="Glu"` is **TRUE** |

> 💡 **Tip:** When you want the worksheet's behavior, call the worksheet function from VBA:
> `Application.WorksheetFunction.Round(2.5, 0)` gives 3, and `WorksheetFunction.CountIf(...)` works like `COUNTIF`.

### 7. Talking to the user: MsgBox and InputBox

**MsgBox** shows a message. Add buttons, an icon, and a title with the optional arguments:

```vba
MsgBox criticalCount & " critical results found."                         ' just OK
MsgBox "Snapshot finished.", vbInformation, "Lab snapshot"               ' icon + title

Dim answer As VbMsgBoxResult
answer = MsgBox("Clear the old TATFlag column first?", vbYesNo + vbQuestion, "FlagSlowStat")
If answer = vbNo Then Exit Sub
```

| Buttons | Icons | Return values |
|---|---|---|
| `vbOKOnly` (default), `vbOKCancel`, `vbYesNo`, `vbYesNoCancel` | `vbInformation`, `vbExclamation`, `vbCritical`, `vbQuestion` | `vbOK`, `vbCancel`, `vbYes`, `vbNo` |

Add a button constant and an icon constant together with `+`, and start a new line inside the message with `vbNewLine`.
The parentheses rule from section 4 applies here too. Use parentheses only when you store the answer, as in
`answer = MsgBox(...)`.

**InputBox** asks for a value. There are two input boxes, VBA's own `InputBox` and Excel's `Application.InputBox`, and they
behave differently:

| | `InputBox` (VBA) | `Application.InputBox` (Excel) |
|---|---|---|
| Returns | Always a String | What you ask for with `Type` |
| Checks the input | No | Yes. `Type:=1` number, `2` text, `4` True/False, `8` a cell range the user selects |
| When the user clicks Cancel | Returns `""` | Returns `False` |
| Use it for | Free text such as a test code | Numbers, or picking cells |

```vba
Dim code As String
code = InputBox("Which TestCode?", "Count results above a limit")
If code = "" Then Exit Sub                       ' Cancel, or nothing typed

Dim limit As Variant                             ' Variant: holds a number OR False
limit = Application.InputBox("Above what value?", "Count results above a limit", Type:=1)
If VarType(limit) = vbBoolean Then Exit Sub      ' the user clicked Cancel
```

With `Type:=1`, Excel rejects text such as "two" and asks again, so by the next line you know you have a number.

> ⚠️ Don't test for Cancel with `If limit = False`. In VBA `False` equals 0, so a user who types **0** looks exactly like a
> user who clicked Cancel. `VarType(limit) = vbBoolean` is True only when Cancel really returned `False`.

> ⚠️ **Passing a Variant to a typed argument.** `WriteCountAbove code, limit` stops with **Compile error: ByRef argument type
> mismatch**, because `limit` is a Variant and `WriteCountAbove` declares its argument `As Double`. Convert the value as you
> pass it: `WriteCountAbove code, CDbl(limit)`. (Lesson 5.4 explains ByRef and ByVal.)

### 8. Reading and writing cells

Your macro reaches a cell through its worksheet. Store the worksheet in an object variable once, then use it. Each line
after the `Set` below points at cells on the Labs sheet, and its comment says which ones:

```vba
Dim ws As Worksheet
Set ws = ThisWorkbook.Worksheets("Labs")

ws.Range("A2").Value            ' "LAB725446", the first LabResultID
ws.Cells(2, 1).Value            ' the same cell: Cells(row, column)
ws.Cells(r, 5).Value            ' ResultValue in row r, where r can change inside a loop
ws.Range("I2:I506")             ' a block of cells
ws.Range("A" & r & ":L" & r)    ' columns A to L of row r
```

`Cells(row, column)` uses numbers for both, which is ideal in loops. These are the column numbers on the Labs sheet (they're
also at the top of `LabMacros.bas`):

| # | Column | # | Column | # | Column |
|:-:|---|:-:|---|:-:|---|
| 1 | A LabResultID | 6 | F Units | 11 | K CollectedDateTime |
| 2 | B EncounterID | 7 | G RefLow | 12 | L ResultedDateTime |
| 3 | C TestCode | 8 | H RefHigh | 13 | M TATFlag (you fill it) |
| 4 | D TestName | 9 | I AbnormalFlag | | |
| 5 | E ResultValue | 10 | J Priority | | |

Your macros read from the Labs sheet and write their answers to the Output sheet, so most of them use two worksheet
variables:

```vba
Dim ws As Worksheet, wsOut As Worksheet
Set ws = ThisWorkbook.Worksheets("Labs")
Set wsOut = ThisWorkbook.Worksheets("Output")

wsOut.Range("B6").Value = criticalCount      ' write a result into Output!B6
```

`ThisWorkbook.Worksheets("Output").Range("B6").Value = criticalCount` does the same job in one line, as `Warmup` does.

This table lists what you can do with a cell or range. For a line that starts with a dot, put the range in front of it, as in
`ws.Range("M2:M506").ClearContents`:

| Code | What it does |
|---|---|
| `x = ws.Cells(r, 5).Value` | Read a value |
| `wsOut.Range("B6").Value = criticalCount` | Write a value |
| `ws.Cells(r, 13).Value = "Review"` | Write text |
| `.ClearContents` | Empty the cells but keep their formatting |
| `.Interior.Color = RGB(255, 235, 156)` | Fill with a color (red, green, blue, each 0–255) |
| `.Interior.ColorIndex = xlNone` | Remove the fill |
| `.NumberFormat = "0.0%"` | Apply a number format |
| `.Font.Bold = True` | Make the text bold |

**The last row.** Your loops need to know where the data ends. Don't hard-code 506, because next month's export will be a
different size. This line starts at the very bottom of column A and jumps up to the last filled cell, just like pressing
**Ctrl + ↑** (Mac: **⌘ + ↑**):

```vba
lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
```

Lesson 5.3 explains how it works and when it fails. For now, copy it.

> ⚠️ **Always say which sheet.** In a standard module, a bare `Range("A1")` or `Cells(2, 1)` means *whichever sheet is active*.
> Run the macro while the Output sheet is showing and `Cells(r, 5)` quietly reads the Output sheet. `ws.Cells(r, 5)` is always
> the Labs sheet.

> ⚠️ **Forgetting `Set`.** `ws = ThisWorkbook.Worksheets("Labs")` without `Set` stops with run-time error 91. Plain values use
> `=`, and objects use `Set ... =`.

> 📋 The Labs data is an Excel Table (`tblLabs`). VBA reads and writes a Table's cells like any other cells. Lesson 5.3
> introduces the Table object (`ListObject`).

### 9. Making decisions: If and Select Case

**If…Then** runs code only when a condition is True. Add `ElseIf` and `Else` branches as needed:

```vba
If condition Then
    ' runs when condition is True
ElseIf otherCondition Then
    ' runs when the first is False and this one is True
Else
    ' runs when none of them is True
End If

If n > 0 Then avg = total / n        ' one-line form: no End If
```

*Worked example.* Classify a potassium result. The thresholds below are illustrative, not a lab policy.

```vba
Dim k As Double, status As String
k = ws.Cells(r, 5).Value                         ' a potassium ResultValue, in mmol/L
If k < 2.5 Or k > 6.5 Then
    status = "Critical"
ElseIf k < 3.5 Or k > 5.1 Then
    status = "Abnormal"
Else
    status = "Normal"
End If
```

The order matters. A value of 7.0 also satisfies `k > 5.1`, but VBA tests the critical branch first and stops at the first
branch that is True.

**Select Case** is cleaner when you compare *one* value against several possibilities:

```vba
Select Case glucose
    Case Is < 70
        band = "Low"
    Case 70 To 180
        band = "In range"
    Case Is > 180
        band = "High"
    Case Else
        band = "Check value"      ' not reached for numbers, but a safe default
End Select
```

| Case form | Matches |
|---|---|
| `Case "HH", "LL"` | Either value (a comma means *or*) |
| `Case 70 To 180` | Anything from 70 through 180 |
| `Case Is > 180` | Anything greater than 180. Use `Is` before a comparison |
| `Case Else` | Anything no earlier Case matched |

> ⚠️ **The first matching Case wins.** VBA runs that one branch and jumps to `End Select`, so it never looks at later Cases,
> even ones that also match. Put the most extreme or most specific Case first.

> ⚠️ Don't write `Case flag = "HH"`. That compares the tested value with True or False. Write `Case "HH"`.

### 10. Loops

A **loop** repeats a block of code. Each repetition is a **pass**.

**For…Next** repeats a set number of times with a **counter** variable:

```vba
For r = 2 To lastRow            ' r = 2, 3, 4 … lastRow
    ' work on row r
Next r
```

`Step` sets how much the counter changes after each pass. Without it, the counter goes up by 1.

```vba
For hr = 0 To 18 Step 6         ' hr = 0, 6, 12, 18
    ' a vital-signs check every 6 hours
Next hr

For r = lastRow To 2 Step -1    ' backwards, which you need when deleting rows
    ' work on row r
Next r
```

- VBA works out the start, the end, and the `Step` once, before the first pass.
- `Next` adds the Step to the counter and then checks it against the end. So when a loop finishes normally, the counter
  holds the first value that went *past* the end, not the last value used.

**For Each** visits every cell in a range (or every item in a collection) without a counter:

```vba
Dim cell As Range
For Each cell In ws.Range("I2:I" & lastRow)      ' every AbnormalFlag cell
    If cell.Value = "N" Then normalCount = normalCount + 1
Next cell
```

It goes through the range row by row, left to right. Use it when you need each cell but not a row number (and if you need
one later, `cell.Row` gives it). For Each also loops over collections such as `ThisWorkbook.Worksheets` (Lesson 5.3).

**Do loops** repeat while (or until) a condition holds, for when you don't know the number of passes in advance:

| Form | Tests the condition | Can run zero times? |
|---|---|---|
| `Do While cond` … `Loop` | Before each pass. Keeps going while it's True | Yes |
| `Do Until cond` … `Loop` | Before each pass. Keeps going until it's True | Yes |
| `Do` … `Loop While cond` | After each pass | No, it always runs at least once |
| `Do` … `Loop Until cond` | After each pass | No, it always runs at least once |

```vba
k = 5
Do Until wsOut.Cells(k, 5).Value = ""      ' find the first empty cell in Output column E
    k = k + 1
Loop
```

**Leaving early.** `Exit For` leaves the innermost For loop at once, and `Exit Do` leaves a Do loop. `Exit Sub` leaves the
whole procedure. Use them when you've found what you were looking for.

**Common loop patterns.** Most macros in this lesson are built from four patterns:

| Pattern | Idea | Sketch |
|---|---|---|
| **Counter** | Add 1 each time a condition is true | `If cond Then n = n + 1` |
| **Accumulator** | Keep a running total (and count), divide at the end | `total = total + value: n = n + 1` |
| **First match** | Remember the match, then stop | `firstID = ...: Exit For` |
| **Running maximum** | Keep the best value seen so far | `If value > best Then best = value` |

The sketches use a colon `:` to put two statements on one line. That's legal VBA, but in your own macros one statement per
line is easier to read and to step through.

**Nested loops.** A loop inside a loop runs completely for every pass of the outer loop. A common use is searching a list:

```vba
' Is "LACT" already listed in Output!E5:E40?
found = False
For k = 5 To 40
    If wsOut.Cells(k, 5).Value = "LACT" Then
        found = True
        Exit For                 ' k still points at the matching row
    End If
Next k
If found Then Debug.Print "LACT is in row " & k Else Debug.Print "Not listed yet"
```

Put that search inside a loop over the lab rows and you have a nested loop.

> ⚠️ **Infinite loops.** A Do loop whose condition never changes runs forever and Excel stops responding. Press **Esc** or
> **Ctrl + Break** (Mac: **⌘ + .** or **Esc**) to interrupt it, then click **End** or **Debug**. On a laptop without a Break key,
> Esc usually works.

> 💡 **Tip:** Reading 505 rows one cell at a time is instant. On tens of thousands of rows it gets slow. Lesson 5.3 shows the
> fast way, which reads the whole range into memory at once.

### 11. Text and date functions in VBA

VBA has its own text and date functions. Many look like worksheet functions, but they're written without the `=`.

| Function | Example | Result |
|---|---|---|
| `Left(text, n)` | `Left("LAB725446", 3)` | `"LAB"` |
| `Right(text, n)` | `Right("LAB725446", 6)` | `"725446"` |
| `Mid(text, start, [n])` | `Mid("ENC110732", 4)` | `"110732"` |
| `Len(text)` | `Len("LAB725446")` | 9 |
| `InStr(text, find)` | `InStr("White Blood Cell Count", "Cell")` | 13, the position where "Cell" starts (0 if not found) |
| `UCase` / `LCase` | `UCase("Glu")` | `"GLU"` |
| `Trim(text)` | `Trim("  K ")` | `"K"` |
| `Replace(text, old, new)` | `Replace("Med-Surg", "-", " ")` | `"Med Surg"` |
| `Format(value, pattern)` | `Format(0.0832, "0.0%")` | `"8.3%"` |
| `Format(date, pattern)` | `Format(#4/28/2025#, "mmm d, yyyy")` | `"Apr 28, 2025"` |

**Case-sensitive text.** VBA compares text exactly, so `"GLU" = "Glu"` is False. To ignore case, convert both sides with
`UCase`, or put `Option Compare Text` at the top of the module (below `Option Explicit`) to make every comparison in that module
case-insensitive.

**Dates.** A date literal goes between `#` signs and is always written month/day/year in code, whatever your regional settings:
`#4/28/2025 6:47 PM#`.

| Function | Example | Result |
|---|---|---|
| `DateDiff(interval, start, end)` | `DateDiff("n", #1/1/2025 4:56 PM#, #1/1/2025 5:22 PM#)` | 26 (minutes) |
| | `DateDiff("d", #1/1/2025#, #4/30/2025#)` | 119 (days) |
| `DateAdd(interval, n, date)` | `DateAdd("d", 30, #1/15/2025#)` | 2/14/2025 |
| `DateSerial(year, month, day)` | `DateSerial(2025, 12, 31)` | 12/31/2025 |
| `Year`, `Month`, `Day`, `Hour`, `Minute` | `Hour(#1/1/2025 4:56 PM#)` | 16 |

The interval codes are `"yyyy"` years, `"q"` quarters, `"m"` months, `"d"` days, `"ww"` weeks, `"h"` hours, `"n"` minutes, and
`"s"` seconds.

> ⚠️ **`"m"` means months.** Minutes are `"n"`. `DateDiff("m", collected, resulted)` returns 0 for almost every lab result.

> ⚠️ **DateDiff counts boundaries crossed**, not full units. `DateDiff("yyyy", #12/31/2024#, #1/1/2025#)` is 1 even though the
> dates are one day apart. That's fine for minutes between two timestamps, but wrong for a patient's age (Lesson 5.4 builds a
> proper age function).

> 📋 Dates are numbers in VBA too, so `resulted - collected` gives days as a decimal. Multiplying by 1440 gives minutes, but
> floating-point error can turn 60 into 59.9999999. `DateDiff("n", …)` returns a clean whole number.

### 12. Debugging

Every programmer writes bugs. What separates beginners from experts is how fast they find them. VBA reports four kinds of
problem:

| Kind | When you see it | Example | What to do |
|---|---|---|---|
| **Syntax error** | As you type. The line turns red | `If x > 5` with no `Then` | Fix the line |
| **Compile error** | When you run, or use **Debug → Compile VBAProject** | *Variable not defined*, *Next without For*, *Expected: =*, *ByRef argument type mismatch* | The VBE highlights the problem. Fix it and run again |
| **Run-time error** | While the code runs. A dialog shows a number | *13: Type mismatch* | Click **Debug** to see the failing line |
| **Logic error** | Never. The code runs and gives a wrong answer | A count of 0 when the data clearly has matches | Step through the code and inspect values |

Logic errors are the dangerous ones because nothing warns you. That's why you check a macro's result against a worksheet
formula such as `COUNTIFS` the first time you run it.

> ⚠️ **A compile error in one Sub can block the whole module.** VBA compiles a module before it runs any procedure in it, so a
> half-finished `FlagSlowStat` with an undeclared variable can stop `Warmup` from running too. When a macro you didn't touch
> suddenly reports a compile error, choose **Debug → Compile VBAProject**. The VBE jumps to the line that needs fixing.

These run-time errors cover most of what you'll meet early on:

| Number | Message | Typical cause |
|:-:|---|---|
| 6 | Overflow | A value too big for its type, such as more than 32,767 in an `Integer`. Also `0 / 0`, for example an average whose total and count both stayed at 0 |
| 9 | Subscript out of range | A misspelled sheet name: `Worksheets("Lab")` |
| 11 | Division by zero | Dividing a number other than 0 by 0, such as `total / n` when `n` is 0 but `total` isn't |
| 13 | Type mismatch | Comparing or calculating with the wrong kind of value, such as text with a number |
| 91 | Object variable or With block variable not set | A missing `Set`, or using an object variable before assigning it |
| 1004 | Application-defined or object-defined error | An impossible cell reference such as `Cells(0, 1)` |

**Break mode.** When a run-time error dialog appears, click **Debug**. The VBE highlights the failing line in yellow and
pauses the code there. This paused state is **break mode**. In break mode you can:

- **Hover** over any variable to see its value in a tooltip.
- Read every variable at once in the **Locals** window. Click the **+** next to an object such as `ws` to explore it.
- Type questions into the **Immediate** window (see below).
- Fix the code, then press **F5** to continue, or click **Reset** (the square ■ button, or **Run → Reset**) to stop.

> ⚠️ While the VBE is in break mode, your macro is still "running." Excel won't run other macros, and some Excel commands are
> unavailable. Click **Reset** when you're done investigating.

**Stepping through code.** Instead of waiting for an error, you can pause the code where *you* choose and run it one line at a
time while you watch the variables change.

| Action | Windows | Mac | What it does |
|---|---|---|---|
| Toggle breakpoint | **F9**, or click the gray margin | Click the gray margin, or use the Debug menu | Marks a line with a red dot. The code pauses *before* running it |
| Step Into | **F8** | Debug menu (in recent versions ⇧ + ⌘ + I) | Runs one line. If it calls another Sub, steps into that Sub |
| Step Over | **Shift + F8** | Debug menu | Runs one line. A called Sub runs all at once |
| Step Out | **Ctrl + Shift + F8** | Debug menu | Runs the rest of the current Sub and pauses after it returns |
| Run to Cursor | **Ctrl + F8** | Debug menu | Runs until the line the cursor is on |
| Continue | **F5** | Run menu or ▶ | Runs to the end or the next breakpoint |
| Clear all breakpoints | **Ctrl + Shift + F9** | Debug menu | Removes every red dot |
| Interrupt a running macro | **Esc** or **Ctrl + Break** | **⌘ + .** or **Esc** | Pauses it and offers Debug |

Breakpoints disappear when you close the workbook. If you want a pause point that's saved with the code, write `Stop` on a line
of its own. Delete it before you share the workbook.

**The Immediate window** (Ctrl + G; Mac: **View → Immediate Window**) is your scratchpad:

- `Debug.Print` writes to it while the code runs. Put a few in a loop to see what happens without stepping:
  `Debug.Print r, ws.Cells(r, 3).Value, ws.Cells(r, 5).Value`. Commas line the values up in columns, and a semicolon prints
  the next value straight after the last one.
- Type `?` and an expression, then press Enter to evaluate it: `? 135 \ 60`, or in break mode `? ws.Cells(r, 3).Value`.
  Break mode lets you use the paused procedure's variables.
- Type a statement and press Enter to run it: `ThisWorkbook.Worksheets("Output").Range("B6").ClearContents`.
- Call a Sub with arguments: `WriteCountAbove "CREAT", 2`.
- You can't use `Dim` here, and each line runs on its own.

**The Watch window** keeps an expression in view while you step. Right-click a variable → **Add Watch…**. Choose
**Break When Value Changes** or **Break When Value Is True** to make the code pause by itself, for example when `r = 120` or when
`highCount > 0`. On Windows, **Shift + F9** shows a **Quick Watch** of the selected expression.

A debugging routine that works:

1. **Read the message.** The error number and the yellow line usually point straight at the problem.
2. **Inspect the values** on that line by hovering, or in the Locals window. Which value isn't what you expected?
3. **Test your theory** in the Immediate window. For example, `? ws.Cells(r, 5).Value` shows what the code is really comparing.
4. **Go back further** if needed. Click Reset, set a breakpoint a few lines earlier, run again, and step with F8.
5. **Fix it, compile, rerun, and cross-check.** Use **Debug → Compile VBAProject**, run the macro, and compare the result with a
   worksheet formula.

### 13. Worked example: from question to working macro

*Question: which routine (non-STAT) result had the longest collect-to-result turnaround, and how long did it take?*

First, write the plan in plain words: *for each row, if it's Routine, work out its minutes. If that's more than the longest so
far, remember it and its LabResultID.* That's a For loop, an If, DateDiff, and the running-maximum pattern. Now the code:

```vba
Sub LongestRoutineTAT()
    Dim ws As Worksheet
    Dim r As Long, lastRow As Long
    Dim minutes As Long, longest As Long
    Dim longestID As String

    Set ws = ThisWorkbook.Worksheets("Labs")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow
        If ws.Cells(r, 10).Value = "Routine" Then                     ' column J = Priority
            minutes = DateDiff("n", ws.Cells(r, 11).Value, ws.Cells(r, 12).Value)
            If minutes > longest Then                                 ' running maximum
                longest = minutes
                longestID = ws.Cells(r, 1).Value
            End If
        End If
    Next r

    MsgBox "Longest routine turnaround: " & longest \ 60 & " h " & longest Mod 60 & " min (" & longestID & ")", _
           vbInformation, "Routine TAT"
End Sub
```

1. Paste it into a new module (**Insert → Module**), put the cursor inside it, and press **F8** a few times. Watch the yellow
   line move and see `r`, `minutes`, and `longest` change in the Locals window.
2. Press **F5** to run the rest. The message reads **Longest routine turnaround: 11 h 42 min (LAB731518)**, a platelet count.
3. Cross-check with a worksheet formula in any empty cell:
   `=ROUND(MAX(IF(Labs!J2:J506="Routine",(Labs!L2:L506-Labs!K2:K506)*1440)),0)` returns 702 minutes, which is 11 h 42 min.
   In Excel 2019 or earlier, confirm it with **Ctrl + Shift + Enter** (Mac: **⌘ + Shift + Return**).

Routine samples are batched, so long turnaround is expected. The same macro with `"STAT"` would find a real problem.

## 🧪 Hands-on practice

Set up the workbook as described in section 1. Then type each answer in a yellow cell on the **Practice** sheet, or run your
macro so the gray cell fills in by itself. The **Check** column turns green when you're right.

Tasks 2–7 ask you to predict what each snippet below prints. The same code is on the workbook's **Snippets** sheet and in
`starter/Snippets.bas`. Commit to a prediction, then run the snippet (cursor inside it, then **F5**; Mac:
**Run → Run Sub/UserForm**) and compare it with the Immediate window. Snippet B is supposed to fail: note the error number,
then click **End**.

**Snippet A** (task 2)

```vba
Sub SnippetA_Typo()
    Dim drawCount As Long
    drawCount = 5
    drawCuont = drawCount + 1       ' add one more blood draw
    Debug.Print drawCount
End Sub
```

**Snippet B** (task 3)

```vba
Sub SnippetB_Overflow()
    Dim labRows As Integer
    labRows = 51527                 ' rows in the full two-year lab file
    Debug.Print labRows
End Sub
```

**Snippet C** (task 4)

```vba
Sub SnippetC_Turnaround()
    Dim tatMinutes As Long
    tatMinutes = 199                ' slowest STAT result on the Labs sheet
    Debug.Print tatMinutes \ 60 & " h " & tatMinutes Mod 60 & " min"
End Sub
```

**Snippet D** (task 5)

```vba
Sub SnippetD_VitalSigns()
    Dim hr As Long, checks As Long
    For hr = 0 To 23 Step 4         ' vital signs every 4 hours (q4h)
        checks = checks + 1
    Next hr
    Debug.Print hr
End Sub
```

**Snippet E** (task 6)

```vba
Sub SnippetE_Lactate()
    Dim lactate As Double, category As String
    lactate = 4.6                   ' mmol/L
    Select Case lactate
        Case Is < 0.5
            category = "Low"
        Case Is <= 2
            category = "Normal"
        Case Is > 2
            category = "Elevated"
        Case Is > 4
            category = "Critical"
        Case Else
            category = "Check value"
    End Select
    Debug.Print category
End Sub
```

**Snippet F** (task 7)

```vba
Sub SnippetF_HalfLife()
    Dim level As Double, hours As Long
    level = 400                     ' ng/mL right after the dose
    Do While level > 50
        level = level / 2           ' the drug's half-life is 6 hours
        hours = hours + 6
    Loop
    Debug.Print hours
End Sub
```

<!-- BEGIN GENERATED: practice -->
Save this workbook as .xlsm, then import starter/LabMacros.bas and starter/Snippets.bas (VBE → File → Import File…). Tasks 2–7 are predict-the-output questions about the code on the Snippets sheet: type your prediction first, then run the snippet to check it. In tasks 1 and 8–13 you run macros, and the gray cells read what your macros wrote to the Output sheet (or to the TATFlag column on the Labs sheet).

| # | Task | Hint |
|:-:|------|------|
| 1 | Import the starter module, then run the Warmup macro (it's already written). It writes the number of lab rows to Output!B5, and the gray cell reads it from there. | Click inside Warmup and press F5, or in Excel press Alt+F8 (Mac: Option+F8), pick Warmup, and click Run |
| 2 | Snippet A: what number does the Immediate window show? Predict first, then run SnippetA_Typo to check. | Which variable does Debug.Print actually read? |
| 3 | Snippet B: running SnippetB_Overflow stops with a run-time error. Type the error number. | Check the Integer row of the data-types table |
| 4 | Snippet C: what exactly does the Immediate window show? Type the whole line in the same pattern, for example 1 h 5 min. | \ keeps the whole part of a division, and Mod keeps the remainder |
| 5 | Snippet D: what number does Debug.Print hr show after the loop finishes? | List every value hr takes, then apply Step one more time |
| 6 | Snippet E: which category does the Immediate window show? (Type the word.) | Select Case runs only the first Case that matches |
| 7 | Snippet F: how many hours does the Immediate window show? | Write down level and hours after each pass, and test the condition before every pass |
| 8 | Complete CountCriticals: use For Each to loop over the AbnormalFlag cells (Labs column I) and count the results flagged HH or LL (critical values). Write the count to Output!B6. | For Each cell In ws.Range("I2:I" & lastRow) … If … Or … Then |
| 9 | Complete AveragePotassium: loop over the rows with For…Next, add up ResultValue for every potassium result (TestCode K) and count them, then write the average to Output!B7. (Full precision or 2 decimal places are both accepted.) | Two accumulators: a Double for the total and a Long for the count |
| 10 | Complete FindFirstCritical: find the first row (from the top) whose AbnormalFlag is HH or LL, write its LabResultID to Output!B8, and leave the loop with Exit For. | Case "HH", "LL" matches either value, and Exit For stops the loop |
| 11 | Complete FlagSlowStat: for every STAT result whose turnaround (ResultedDateTime minus CollectedDateTime, in minutes, with DateDiff) is more than 60 minutes, write SLOW in that row's TATFlag cell (Labs column M). Leave all other rows empty. The gray cell counts the SLOW flags. | Use DateDiff("n", start, finish), then combine the two tests with And or a nested If |
| 12 | Complete CountAbove and WriteCountAbove. CountAbove asks for a TestCode (InputBox) and a limit (Application.InputBox with Type:=1), then calls WriteCountAbove, which counts the results of that test above the limit, writes the count to Output!B9, and shows it in a MsgBox. Run CountAbove and enter CREAT and 2. | Call a Sub that takes arguments without parentheses: SubName arg1, arg2. Convert the Variant limit with CDbl |
| 13 | BuggyGlucoseCount (already in the starter module) should count glucose (GLU) results above 180 mg/dL and write the count to Output!B10. It has two bugs: the first stops it with a run-time error, and after you fix that one it writes 0. Use F8, the Locals window, and the Immediate window to find and fix both, then run it. | When it stops, click Debug and hover over r. While it's paused, try ? ws.Cells(5, 3).Value (a glucose row) in the Immediate window |
<!-- END GENERATED: practice -->

## ✅ Answer key

The workbook has a hidden **Answer Key** sheet (right-click any sheet tab → **Unhide…** → *Answer Key*). Its *Live result*
column checks every macro answer with a worksheet formula. Complete reference macros are in
[`solutions/LabMacros_Solution.bas`](solutions/LabMacros_Solution.bas) (spoilers). Import it into a spare copy of the workbook so
its procedure names don't clash with yours. The answers are also below, collapsed so you don't see them by accident.

<!-- BEGIN GENERATED: answers -->
<details>
<summary><b>🔑 Show the answer key</b> — Try every task before opening this.</summary>

**1. Warmup: run your first macro**

- **Answer:** 505
- **Solution:**

```vba
Sub Warmup()
    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = ThisWorkbook.Worksheets("Labs")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row      ' last filled row in column A
    ThisWorkbook.Worksheets("Output").Range("B5").Value = lastRow - 1   ' minus the header row
    Debug.Print "Warmup ran: " & (lastRow - 1) & " lab rows"
End Sub
```


Warmup finds the last filled row in column A, subtracts 1 for the header row, and writes the result into a cell. If the gray cell stays empty, check that you saved as .xlsm, enabled macros, and ran Warmup in **this** workbook. If you made a Personal Macro Workbook in Lesson 5.1, also check that LabMacros was imported into this workbook's project and not into PERSONAL.XLSB. Open the Immediate window (Ctrl + G; Mac: View → Immediate Window) to see the line Debug.Print wrote.

**2. Snippet A: what number does the Immediate window show? Predict first, then run…**

- **Answer:** 5
- **Solution:** Trace it: `drawCount` is set to 5. The next line assigns to `drawCuont`, which is a *different* variable, so `drawCount` is still 5 when it's printed.

Without **Option Explicit**, VBA silently creates a new Variant for any misspelled name, so the typo doesn't raise an error and the result is quietly wrong. Add `Option Explicit` at the top of the module and the same code stops with *Compile error: Variable not defined*, highlighting `drawCuont`.

**3. Snippet B: running SnippetB_Overflow stops with a run-time error. Type the error number.**

- **Answer:** 6
- **Solution:** Run-time error **6: Overflow**. An `Integer` holds only −32,768 to 32,767, and 51,527 is bigger.

`Integer` is a 16-bit type left over from early versions of VBA. Use `Long` (up to about 2.1 billion) for counts and row numbers. A worksheet has 1,048,576 rows, so an `Integer` row counter fails on any large sheet. `Long` is also no slower on modern computers.

**4. Snippet C: what exactly does the Immediate window show? Type the whole line in the…**

- **Answer:** 3 h 19 min
- **Solution:** `199 \ 60` is 3 (whole hours) and `199 Mod 60` is 19 (leftover minutes). `&` joins everything into one string: **3 h 19 min**.

Arithmetic operators run before `&`, so each piece is calculated first and then joined. Integer division `\` and `Mod` are the standard way to split minutes into hours and minutes. (199 minutes was the slowest STAT turnaround on the Labs sheet, against a 60-minute target.)

**5. Snippet D: what number does Debug.Print hr show after the loop finishes?**

- **Answer:** 24
- **Solution:** The loop body runs for hr = 0, 4, 8, 12, 16, 20 (6 checks). `Next` then adds the Step again (20 + 4 = 24), which is past 23, so the loop ends with hr = **24**.

After a For…Next loop finishes normally, the counter holds the first value that *failed* the test, not the last value used. That's why code that needs "the last row processed" should save it in its own variable or leave the loop with Exit For.

**6. Snippet E: which category does the Immediate window show? (Type the word.)**

- **Answer:** Elevated
- **Solution:** 4.6 is not < 0.5 and not ≤ 2, but it **is** > 2, so `Case Is > 2` matches and sets **Elevated**. VBA then jumps to End Select, so `Case Is > 4` is never tested.

Order matters in Select Case. Put the most specific (most extreme) test first: `Case Is > 4` before `Case Is > 2`. Here a critical lactate of 4.6 mmol/L is mislabeled "Elevated", which is exactly the kind of silent bug that testing with real values catches.

**7. Snippet F: how many hours does the Immediate window show?**

- **Answer:** 18
- **Solution:** level 400 → 200 (6 h) → 100 (12 h) → 50 (18 h). Before the next pass VBA tests `50 > 50`, which is False, so the loop stops: **18**.

`Do While … Loop` tests the condition *before* each pass, so the body may run zero times. Use `Do … Loop While` when the body must run at least once. Watch the boundary: `> 50` stops at exactly 50, while `>= 50` would run one more pass and print 24.

**8. Complete CountCriticals: use For Each to loop over the AbnormalFlag cells (Labs column…**

- **Answer:** 24
- **Solution:**

```vba
Sub CountCriticals()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim cell As Range
    Dim criticalCount As Long

    Set ws = ThisWorkbook.Worksheets("Labs")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For Each cell In ws.Range("I2:I" & lastRow)          ' the AbnormalFlag column
        If cell.Value = "HH" Or cell.Value = "LL" Then
            criticalCount = criticalCount + 1
        End If
    Next cell

    ThisWorkbook.Worksheets("Output").Range("B6").Value = criticalCount
    MsgBox criticalCount & " critical results", vbInformation, "Count criticals"
End Sub
```


This is the **counter pattern**: start a Long at 0 and add 1 each time a condition is true. Each `Or` side must be a full comparison: `cell.Value = "HH" Or cell.Value = "LL"`. Writing `cell.Value = "HH" Or "LL"` gives a Type mismatch error. Cross-check with a formula: `=COUNTIF('Labs'!I2:I506,"HH")+COUNTIF('Labs'!I2:I506,"LL")`.

**9. Complete AveragePotassium: loop over the rows with For…Next, add up ResultValue for…**

- **Answer:** 4.07
- **Solution:**

```vba
Sub AveragePotassium()
    Dim ws As Worksheet
    Dim r As Long, lastRow As Long
    Dim total As Double, n As Long

    Set ws = ThisWorkbook.Worksheets("Labs")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow
        If ws.Cells(r, 3).Value = "K" Then           ' column C = TestCode
            total = total + ws.Cells(r, 5).Value     ' column E = ResultValue
            n = n + 1
        End If
    Next r

    If n > 0 Then
        ThisWorkbook.Worksheets("Output").Range("B7").Value = total / n
        Debug.Print n & " potassium results, average " & Format(total / n, "0.00")
    Else
        MsgBox "No potassium results found.", vbExclamation
    End If
End Sub
```


The **accumulator pattern** keeps a running total and a running count, then divides once at the end. The `If n > 0` guard matters when no rows match: then `total` and `n` are both still 0, and `0 / 0` stops with run-time error 6 (*Overflow*). Any other number divided by 0 gives error 11 (*Division by zero*). The 24 potassium results average 4.07 mmol/L, which is comfortably inside the 3.5–5.1 reference range. If you round in VBA, remember that `Round` uses banker's rounding.

**10. Complete FindFirstCritical: find the first row (from the top) whose AbnormalFlag is HH…**

- **Answer:** LAB725484
- **Solution:**

```vba
Sub FindFirstCritical()
    Dim ws As Worksheet
    Dim r As Long, lastRow As Long
    Dim firstID As String

    Set ws = ThisWorkbook.Worksheets("Labs")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow
        Select Case ws.Cells(r, 9).Value             ' column I = AbnormalFlag
            Case "HH", "LL"
                firstID = ws.Cells(r, 1).Value       ' column A = LabResultID
                Exit For                             ' found it: stop looping
        End Select
    Next r

    If firstID = "" Then firstID = "None found"
    ThisWorkbook.Worksheets("Output").Range("B8").Value = firstID
    Debug.Print "First critical: " & firstID & " (row " & r & ")"
End Sub
```


`Exit For` leaves the loop the moment the first match is found (sheet row 7), so the macro doesn't waste time on the remaining rows and `firstID` can't be overwritten by a later match. A Case with a comma-separated list (`Case "HH", "LL"`) matches any of the values.

**11. FlagSlowStat (count of SLOW flags)**

- **Answer:** 87
- **Solution:**

```vba
Sub FlagSlowStat()
    Dim ws As Worksheet
    Dim r As Long, lastRow As Long
    Dim minutes As Long, slowCount As Long

    Set ws = ThisWorkbook.Worksheets("Labs")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow
        ws.Cells(r, 13).ClearContents                ' start clean so a re-run is safe
        If ws.Cells(r, 10).Value = "STAT" Then       ' column J = Priority
            ' "n" = minutes (not "m", which is months)
            minutes = DateDiff("n", ws.Cells(r, 11).Value, ws.Cells(r, 12).Value)
            If minutes > 60 Then
                ws.Cells(r, 13).Value = "SLOW"       ' column M = TATFlag
                slowCount = slowCount + 1
            End If
        End If
    Next r

    Debug.Print slowCount & " slow STAT results"
End Sub
```


`DateDiff("n", …)` returns whole minutes. The interval code is `"n"` because `"m"` means *months*. 87 of the 352 STAT results missed the 60-minute target. Clearing column M at the start of each pass makes the macro safe to run twice, because an old SLOW flag can't survive a re-run.

**12. CountAbove (CREAT above 2)**

- **Answer:** 12
- **Solution:**

```vba
Sub CountAbove()
    Dim testCode As String
    Dim limitInput As Variant

    testCode = UCase(Trim(InputBox("Which TestCode? (for example CREAT)", "Count results above a limit")))
    If testCode = "" Then Exit Sub                   ' Cancel (or nothing typed)

    limitInput = Application.InputBox("Count results above what value?", "Count results above a limit", Type:=1)
    If VarType(limitInput) = vbBoolean Then Exit Sub ' Cancel returns False

    WriteCountAbove testCode, CDbl(limitInput)
End Sub

Sub WriteCountAbove(testCode As String, limit As Double)
    Dim ws As Worksheet
    Dim r As Long, lastRow As Long, n As Long

    Set ws = ThisWorkbook.Worksheets("Labs")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow
        If ws.Cells(r, 3).Value = testCode Then
            If ws.Cells(r, 5).Value > limit Then n = n + 1
        End If
    Next r

    ThisWorkbook.Worksheets("Output").Range("B9").Value = n
    MsgBox n & " " & testCode & " results above " & limit, vbInformation, "Count results above a limit"
End Sub
```


Splitting the work into two Subs keeps the dialog code apart from the counting code. You can test `WriteCountAbove` straight from the Immediate window (`WriteCountAbove "CREAT", 2`) without clicking through dialogs. `Application.InputBox` with `Type:=1` refuses non-numbers and returns **False** on Cancel, which is why the result goes into a Variant and is checked with `VarType`. Creatinine above 2 mg/dL is a rough screen for impaired kidney function. Formal acute kidney injury criteria compare each patient with their own baseline.

**13. Debug BuggyGlucoseCount**

- **Answer:** 20
- **Solution:**

```vba
Sub BuggyGlucoseCount()
    Dim ws As Worksheet
    Dim r As Long, lastRow As Long
    Dim highCount As Long

    Set ws = ThisWorkbook.Worksheets("Labs")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow                             ' FIX 1: start below the header row
        If ws.Cells(r, 3).Value = "GLU" And ws.Cells(r, 5).Value > 180 Then   ' FIX 2: "GLU", not "Glu"
            highCount = highCount + 1
        End If
    Next r

    ThisWorkbook.Worksheets("Output").Range("B10").Value = highCount
End Sub
```


**Bug 1:** the loop starts at row 1, the header row. `ws.Cells(1, 5).Value` is the text "ResultValue", and comparing text with the number 180 raises *Run-time error 13: Type mismatch*. VBA's `And` evaluates both sides even when the first is already False, so the TestCode test doesn't protect you. **Bug 2:** text comparison in VBA is case-sensitive by default, so `"GLU" = "Glu"` is False and nothing is ever counted. Fix it by matching the data exactly (`"GLU"`), or compare `UCase(ws.Cells(r, 3).Value) = "GLU"`.

</details>
<!-- END GENERATED: answers -->

## 🏆 Bonus challenge

<!-- BEGIN GENERATED: bonus -->
The ICU medical director wants a one-click 'lab snapshot' for the morning huddle. Complete the LabSnapshot stub in the starter module so that it does four things:
1. Clears the old fills on Labs!A2:L506, the old summary in Output!E5:H40, and Output!J5, so you can run it again safely.
2. Colors each data row (columns A:L) by its AbnormalFlag: light red RGB(255, 199, 206) for HH or LL, amber RGB(255, 235, 156) for H or L, and no fill for N.
3. Builds a summary table on the Output sheet from row 5 down, with one row per TestCode in the order each code first appears. The columns are TestCode (E), Results (F), Abnormal (G, any flag other than N), and PctAbnormal (H, Abnormal ÷ Results, formatted 0.0%).
4. Writes the TestCode with the highest PctAbnormal to Output!J5.

Don't use Scripting.Dictionary, because that's Lesson 5.4. A nested loop that searches the table you're building is enough.

Work on the **Bonus** sheet of the workbook.

- **B1.** Run LabSnapshot. Then click the filter arrow on the Labs table's AbnormalFlag header and choose Filter by Color → your amber fill. How many rows are amber? (The status bar shows 'x of 505 records found'.) *(Hint: Select Case flag: Case "HH", "LL" … Case "H", "L" …)*
- **B2.** How many TestCode rows does your summary table have? (The gray cell counts the codes in Output!E5:E40.) *(Hint: Search rows 5 to nextRow - 1 with an inner For loop. If the code isn't there, add a row)*
- **B3.** How many creatinine (CREAT) results are abnormal (any flag other than N)? (The gray cell looks up CREAT in your table.) *(Hint: If flag <> "N" Then add 1 to column G of the test's row)*
- **B4.** What percentage of lactate (LACT) results are abnormal? (The gray cell looks up LACT in your table.) *(Hint: Compute the percentages in a second loop, after every row has been counted)*
- **B5.** Which TestCode has the highest percentage of abnormal results? (The gray cell reads Output!J5.) *(Hint: Track the best value so far: If pct > topPct Then …)*
<!-- END GENERATED: bonus -->

The reference solution is in [`solutions/LabSnapshot_Solution.bas`](solutions/LabSnapshot_Solution.bas) (spoilers).

<!-- BEGIN GENERATED: bonus-answers -->
<details>
<summary><b>🔑 Show the bonus solution</b> — Give it a real try first!</summary>

**B1. Run LabSnapshot, then count the amber rows**

- **Answer:** 215
- **Solution:**

```vba
Sub LabSnapshot()
    Dim wsLab As Worksheet, wsOut As Worksheet
    Dim r As Long, lastRow As Long
    Dim k As Long, nextRow As Long
    Dim flag As String, code As String
    Dim found As Boolean
    Dim pct As Double, topPct As Double, topCode As String

    Set wsLab = ThisWorkbook.Worksheets("Labs")
    Set wsOut = ThisWorkbook.Worksheets("Output")
    lastRow = wsLab.Cells(wsLab.Rows.Count, 1).End(xlUp).Row

    ' 1. Reset: remove old fills and the old summary so the macro can be re-run
    wsLab.Range("A2:L" & lastRow).Interior.ColorIndex = xlNone
    wsOut.Range("E5:H40").ClearContents
    wsOut.Range("J5").ClearContents
    nextRow = 5                                      ' first empty row of the summary table

    For r = 2 To lastRow
        flag = wsLab.Cells(r, 9).Value               ' column I = AbnormalFlag
        code = wsLab.Cells(r, 3).Value               ' column C = TestCode

        ' 2. Color the row (columns A:L) by its flag
        Select Case flag
            Case "HH", "LL"
                wsLab.Range("A" & r & ":L" & r).Interior.Color = RGB(255, 199, 206)
            Case "H", "L"
                wsLab.Range("A" & r & ":L" & r).Interior.Color = RGB(255, 235, 156)
        End Select

        ' 3. Find this test in the summary table; add a new row if it isn't there yet
        found = False
        For k = 5 To nextRow - 1
            If wsOut.Cells(k, 5).Value = code Then
                found = True
                Exit For                             ' k now points at the test's row
            End If
        Next k
        If Not found Then
            k = nextRow
            wsOut.Cells(k, 5).Value = code
            wsOut.Cells(k, 6).Value = 0
            wsOut.Cells(k, 7).Value = 0
            nextRow = nextRow + 1
        End If

        ' 4. Update the counts
        wsOut.Cells(k, 6).Value = wsOut.Cells(k, 6).Value + 1          ' Results
        If flag <> "N" Then
            wsOut.Cells(k, 7).Value = wsOut.Cells(k, 7).Value + 1      ' Abnormal
        End If
    Next r

    ' 5. Percent abnormal for each test, tracking the highest
    For k = 5 To nextRow - 1
        pct = wsOut.Cells(k, 7).Value / wsOut.Cells(k, 6).Value
        wsOut.Cells(k, 8).Value = pct
        If pct > topPct Then
            topPct = pct
            topCode = wsOut.Cells(k, 5).Value
        End If
    Next k
    wsOut.Range("H5:H40").NumberFormat = "0.0%"
    wsOut.Range("J5").Value = topCode

    MsgBox "Tests summarized: " & (nextRow - 5) & vbNewLine & _
           "Highest % abnormal: " & topCode & " (" & Format(topPct, "0.0%") & ")", _
           vbInformation, "Lab snapshot"
End Sub
```


A worksheet formula can't see fill colors, so Filter by Color (Lesson 1.6) is how you check the coloring. Clearing the old fills first matters: without it, a row that was amber yesterday would stay amber even if its flag changed. Clear the filter afterwards (Data → Clear). The full macro is also in solutions/LabSnapshot_Solution.bas.

**B2. How many TestCode rows does your summary table have? (The gray cell counts the codes…**

- **Answer:** 10
- **Solution:** Run LabSnapshot (full code in B1). The gray cell uses `=COUNTA(Output!E5:E40)`.

The inner loop searches only the rows already written (5 to nextRow − 1). If it finishes without a match, the code is new, so the macro writes it in `nextRow` and moves `nextRow` down one row. That search-or-add pattern is what a Dictionary does for you in Lesson 5.4. The ICU slice has 10 distinct tests.

**B3. How many creatinine (CREAT) results are abnormal (any flag other than N)? (The gray…**

- **Answer:** 40
- **Solution:** Run LabSnapshot. The gray cell uses `=INDEX(Output!G5:G40,MATCH("CREAT",Output!E5:E40,0))`.

`<> "N"` counts H, L, HH, and LL together, which is what *abnormal* means here. Updating the count in the found row `k` works because `Exit For` left `k` pointing at that row.

**B4. What percentage of lactate (LACT) results are abnormal? (The gray cell looks up LACT…**

- **Answer:** 76.9%
- **Solution:** Run LabSnapshot. The gray cell uses `=INDEX(Output!H5:H40,MATCH("LACT",Output!E5:E40,0))`.

Store the ratio (0.769…) and let `.NumberFormat = "0.0%"` display it. If you store 76.9 instead, the check accepts that too, but a true fraction is easier to reuse in later formulas. Dividing in a separate loop after counting means each test is divided once, with its final totals.

**B5. Which TestCode has the highest percentage of abnormal results? (The gray cell reads…**

- **Answer:** GLU
- **Solution:** Run LabSnapshot. Step 5 keeps the best value so far in `topPct` and its code in `topCode`.

The **running maximum** pattern compares each value with the best one seen so far. In this ICU sample, glucose (69 of 85, 81.2%) edges out lactate (76.9%). That's plausible, because stress hyperglycemia is common in critically ill patients. Lactate has the most *critical* values, though (21 of the 24 HH or LL results), so "most often abnormal" and "most dangerous" are different questions.

</details>
<!-- END GENERATED: bonus-answers -->

## Key takeaways

- Write macros in **standard modules**, save the workbook as **.xlsm**, and start every module with **`Option Explicit`**.
- Declare every variable with a type: **`Long`** for whole numbers (not `Integer`), **`Double`** for decimals, **`String`**,
  **`Date`**, **`Boolean`**, and object types assigned with **`Set`**.
- Qualify every cell with its sheet (`ws.Cells(r, c)`), and find the last row instead of hard-coding it.
- **If/ElseIf** and **Select Case** make decisions, and the **first** branch that matches wins. Text comparisons are
  case-sensitive.
- **For…Next** walks rows, **For Each** walks cells, and **Do While/Until** loops until a condition changes. Counter,
  accumulator, first-match, and running-maximum patterns cover most everyday macros.
- Debug in **break mode**: F8 to step, F9 for breakpoints, the Locals window for values, and the Immediate window for
  `Debug.Print` and `?` questions. Cross-check every new macro against a worksheet formula.

<!-- BEGIN GENERATED: nav -->
---

⬅️ **Previous:** [5.1 Recording Your First Macros](../01-recording-macros/README.md) · 🏠 [Course home](../../README.md) · **Next:** [5.3 VBA: Ranges, Worksheets & Workbooks](../03-vba-ranges-worksheets/README.md) ➡️
<!-- END GENERATED: nav -->
