# Lesson 5.3 · VBA: Ranges, Worksheets & Workbooks

> **Level:** Expert · **Time:** about 180 minutes · **Workbook:** [`5.3-vba-ranges-worksheets.xlsx`](5.3-vba-ranges-worksheets.xlsx)
> **Data:** 2,000 encounters admitted in 2025 at all four Bluestone facilities (inpatient, observation, emergency, and outpatient), exported from the EHR as a plain range, plus the facility list as an Excel Table.

Every month, Bluestone's finance team receives the encounter export from the EHR: a few thousand rows, one per visit or
stay, from four facilities. Before anyone can use it, someone filters it four times to give each facility its own sheet,
builds a summary by encounter type, saves a separate file for Ashby Falls, and checks that no rows went missing on the way.
By hand that takes the better part of an hour, and next month's export has a different number of rows. In this lesson you'll
write macros that do the whole job in seconds and work on an export of any size. You'll learn the objects your code moves
around (workbooks, worksheets, and ranges), the idioms that find the data wherever it ends, and the techniques that keep
macros fast: arrays, AutoFilter, and Sort, all without selecting a single cell.

## What you'll learn

- Navigate the object model: Application, Workbooks, Worksheets, Range
- Find the last row and work with CurrentRegion, Offset, and Resize
- Loop through sheets; add, copy, rename, and delete them
- Process data fast with arrays, AutoFilter, and Sort — without Select

## 📖 Guide

### 1. Set up your workbook

You need three files from this lesson's folder: the workbook and the two starter modules
[`starter/EncounterMacros.bas`](starter/EncounterMacros.bas) and [`starter/Snippets.bas`](starter/Snippets.bas). On GitHub,
open each `.bas` file and click **Download raw file**.

1. Open `5.3-vba-ranges-worksheets.xlsx` and save it as a macro-enabled workbook: **File → Save As**, then choose
   **Excel Macro-Enabled Workbook (\*.xlsm)**. Save it in an ordinary folder on your computer rather than a synced OneDrive
   folder, because task 11 saves a second file next to it (section 9 explains why).
2. Open the Visual Basic Editor (the VBE) with **Alt + F11** (Mac: **Developer → Visual Basic**).
3. Choose **File → Import File…** (Windows shortcut: **Ctrl + M**) and import `EncounterMacros.bas`, then `Snippets.bas`.
   Both appear in the **Modules** folder of the Project Explorer.
4. Open the **Immediate window** with **Ctrl + G** (Mac: **View → Immediate Window**). The snippets print there, and you can
   type `?` and an expression to ask Excel a question at any time.

The workbook's sheets:

| Sheet | What it holds |
|---|---|
| **Practice** | Tasks 1–12: yellow cells for your predictions, gray cells that read what your macros create |
| **Snippets** | The code for tasks 1–6 (the same code is in `Snippets.bas`) |
| **Encounters** | 2,000 encounters from 2025 in columns A:K, plus an empty LOSDays column (L) for task 12. It's a plain range, like a raw export |
| **Facilities** | The four facilities, as an Excel Table named `tblFacilities` |
| **Output** | Cells your macros write results to |
| **Bonus**, **PacketTemplate** | The bonus challenge and the template sheet it copies |

These are the Encounters columns. They're also listed at the top of `EncounterMacros.bas`.

| # | Column | # | Column | # | Column |
|:-:|---|:-:|---|:-:|---|
| 1 | A EncounterID | 5 | E DeptID | 9 | I PrimaryDxCode |
| 2 | B PatientID | 6 | F AdmitSource | 10 | J PayerID |
| 3 | C EncounterType | 7 | G AdmitDateTime | 11 | K TotalCharges |
| 4 | D FacilityID | 8 | H DischargeDateTime | 12 | L LOSDays (you fill it) |

> ⚠️ **Save before you run.** Undo, **Ctrl + Z** (Mac: **⌘ + Z**), can't undo a macro, and the macros in this lesson add and
> delete whole sheets. If something goes wrong, close without saving and reopen.

> 📋 **Version note:** VBA runs in desktop Excel for Windows and Excel for Mac. Excel for the web, iPad, iPhone, and Android can
> open an `.xlsm` file but can't run its macros. Everything in this lesson works in Excel 2013 or later on Windows and Excel 2016
> or later on a Mac.

### 2. The object model

An **object** is anything in Excel your code can work with: the Excel application, a workbook, a worksheet, a range of cells.
Objects are organized in a hierarchy called the **object model**, where each object contains others:

```
Application                      Excel itself
└── Workbooks                    every open workbook
    └── Workbook                 one file, such as 5.3-vba-ranges-worksheets.xlsm
        └── Worksheets           the workbook's worksheets
            └── Worksheet        one sheet, such as Encounters
                ├── Range        a cell or a block of cells, such as K2 or A2:K2001
                └── ListObjects  the sheet's Excel Tables
```

You move down the hierarchy with dots. This line names one cell by its full address, starting at the top:

```vba
Application.Workbooks("5.3-vba-ranges-worksheets.xlsm").Worksheets("Encounters").Range("K2").Value
```

The things you can use on an object are its **members**. There are two kinds, properties and methods. Many properties return
a third kind of thing, a collection:

| Idea | What it is | Examples |
|---|---|---|
| **Property** | A fact or setting you can read, and often change | `ws.Name`, `rng.Value`, `rng.Address`, `Application.ScreenUpdating` |
| **Method** | An action the object performs | `ws.Delete`, `rng.Sort`, `rng.AutoFilter`, `wb.Close` |
| **Collection** | An object that holds a set of objects of one kind. A property of the parent returns it, and its name is the plural | `Workbooks`, `Worksheets`, `ListObjects` |

Every collection works the same way. You pick one item by name or by position, count the items, or loop over them:

```vba
ThisWorkbook.Worksheets("Encounters")      ' one item, by name
ThisWorkbook.Worksheets(1)                 ' one item, by position: the leftmost tab
ThisWorkbook.Worksheets.Count              ' how many items there are
For Each ws In ThisWorkbook.Worksheets     ' every item in turn (section 7)
```

Many properties return another object. `ws.Range("A2")` returns a Range, and `ws.Parent` returns the workbook the sheet belongs
to. That's why you can keep adding dots.

> 💡 **Tip:** To explore an object, declare a variable of its type, type the variable and a dot, and IntelliSense lists every
> member. Press **F2** in the VBE (Mac: **View → Object Browser**) to open the **Object Browser**, where you can search for a
> member such as `CurrentRegion`. On Windows, select a member and press **F1** for its help page.

### 3. Pointing at the right workbook and sheet

Your code has to say *which* workbook and *which* sheet it means. Three references sound alike but behave very differently:

| Reference | Means | Changes when |
|---|---|---|
| `ThisWorkbook` | The workbook that contains the running code | Never |
| `ActiveWorkbook` | The workbook whose window is in front | You or your code open, create, or switch to another workbook |
| `ActiveSheet` | The sheet showing in the active workbook | Someone clicks a tab, or code activates a sheet |

Use `ThisWorkbook` for your own sheets. In task 11 your macro opens a second workbook, and from that moment `ActiveWorkbook` is
the other file.

There are also three ways to name one sheet:

| Way | Example | Breaks when |
|---|---|---|
| By **tab name** | `ThisWorkbook.Worksheets("Encounters")` | Someone renames the tab (run-time error 9) |
| By **index**, its position | `ThisWorkbook.Worksheets(4)` | Someone adds, moves, or deletes a sheet. Hidden sheets count too |
| By **code name** | `Sheet4.Range("A1")` | Only when someone edits the code name in the VBE |

The **code name** is a second name that only VBA sees. The Project Explorer shows it first, with the tab name in parentheses,
for example `Sheet4 (Encounters)`. To change it, select the sheet in the Project Explorer and edit **(Name)** in the Properties
window, which **F4** opens (Mac: **View → Properties Window**). A code name only works inside the workbook that holds the
code, and a sheet your macro creates gets a code name such as `Sheet12` that you can't know in advance. That's why this lesson
uses tab names.

`Worksheets` holds only worksheets, while `Sheets` also includes chart sheets. Use `Worksheets` unless you really mean charts too.

**The implicit-reference rule.** In a standard module, VBA quietly fills in any object you leave out:

| You write | VBA reads it as |
|---|---|
| `Range("A2")`, `Cells(2, 1)`, `Rows(1)` | `ActiveSheet.Range("A2")`, `ActiveSheet.Cells(2, 1)`, `ActiveSheet.Rows(1)` |
| `Worksheets("Encounters")` | `ActiveWorkbook.Worksheets("Encounters")` |
| `ws.Range(Cells(2, 1), Cells(2, 11))` | `ws.Range(ActiveSheet.Cells(2, 1), ActiveSheet.Cells(2, 11))` |

The third line of that table is the sneakiest bug in VBA. The `ws.` in front makes the code look qualified, but the two `Cells` inside belong to
the active sheet. The line works while `ws` happens to be the active sheet and fails as soon as another sheet is in front,
because a range can't start on one sheet and end on another. **Qualify every reference.**

> 📋 In a **sheet module**, the code behind one sheet (Lesson 5.5), an unqualified `Range` means that sheet, not the active
> sheet. Code copied between modules can silently change meaning, which is one more reason to qualify everything.

**Object variables.** Store any object you use more than once in a variable, assigned with `Set`:

```vba
Dim wb As Workbook, ws As Worksheet
Set wb = ThisWorkbook
Set ws = wb.Worksheets("Encounters")
Debug.Print ws.Name, ws.Range("A2").Value          ' Encounters    ENC110713
```

**With … End With** saves typing the same object again and again. Inside the block, start each member with a dot:

```vba
With ThisWorkbook.Worksheets("Encounters")
    Debug.Print .Range("A2").Value                           ' ENC110713
    Debug.Print .Range(.Cells(2, 1), .Cells(2, 11)).Address  ' $A$2:$K$2
End With
```

Look at the dots in front of both `.Cells`. Without them, you're back to the active sheet.

### 4. Range, Cells, Rows, and Columns

There are many ways to point at cells. These cover almost everything:

| Code | Refers to |
|---|---|
| `ws.Range("K2")` | One cell |
| `ws.Range("A2:K2001")` | A block of cells |
| `ws.Range("K2:K" & lastR)` | A block whose end you work out while the code runs |
| `ws.Cells(2, 11)` or `ws.Cells(2, "K")` | Row 2, column 11 (K). Ideal inside loops |
| `ws.Range(ws.Cells(2, 1), ws.Cells(lastR, 11))` | The block between two corner cells |
| `ws.Rows(1)`, `ws.Columns("K")` | A whole row or column |
| `rng.EntireRow`, `rng.EntireColumn` | The whole rows or columns a range touches |
| `ws.Range("C5").Cells(2, 2)` | **Relative**: row 2, column 2 *of that range*, which is D6 |

Every range can tell you about itself:

| Property | Returns | Example on Encounters |
|---|---|---|
| `.Address` | Its address, with `$` signs | `ws.Range("A2:K2001").Address` → `$A$2:$K$2001` |
| `.Address(False, False)` | Its address without `$` signs | `A2:K2001` |
| `.Row`, `.Column` | The number of its first row and first column | `ws.Range("K2").Column` → 11 |
| `.Rows.Count`, `.Columns.Count` | Its height and width | 2,000 and 11 for A2:K2001 |
| `.Count` | How many cells it has | 22,000 for A2:K2001 |

**Value, Value2, and Text.** A cell holds one value but can show it in different ways, so VBA has several properties for its
contents. Here's row 2 of Encounters, ENC110713, an inpatient stay admitted at 02:12 on New Year's Day:

| Property | G2 (AdmitDateTime) | K2 (TotalCharges) | Use it for |
|---|---|---|---|
| `.Value` | A Date. The Immediate window shows it in your regional format, such as `1/1/2025 2:12:00 AM` | `10673.63` | Everyday reading and writing |
| `.Value2` | `45658.0916666667`, the date's serial number as a Double | `10673.63` | Speed, and dates as plain numbers |
| `.Text` | `"01/01/2025 02:12"`, exactly what the cell displays | `"10,673.63"` | Copying what the user sees |

`.Formula` returns or sets a cell's formula as text, as in
`ThisWorkbook.Worksheets("Output").Range("E4").Formula = "=SUM(Encounters!K2:K2001)"`. It always uses English function names
and commas, whatever language your Excel uses.

> ⚠️ `.Text` is always a String, and it returns `####` when the column is too narrow to show the value. Never calculate with it.

> 📋 The only difference between `.Value` and `.Value2` is that Value2 never uses the Date or Currency data types. For a cell
> formatted as currency, `.Value` returns a Currency value, which keeps only 4 decimal places. The Encounters charges use an
> ordinary number format, so both properties give the same Double.

### 5. Moving and resizing: Offset, Resize, CurrentRegion, UsedRange

```vba
rng.Offset(rowOffset, columnOffset)    ' the same size, moved
rng.Resize(rowSize, columnSize)        ' the same top-left corner, a new size
```

- **Offset** moves a range without changing its size. Positive numbers go down and right, and negative numbers go up and left.
  Leave out an argument, or use 0, to stay put in that direction.
- **Resize** keeps the top-left corner and sets a new number of rows and columns. Leave out an argument to keep that
  dimension as it is.

| Code | Result | Why |
|---|---|---|
| `Range("B3").Offset(1, 1)` | C4 | One row down and one column right |
| `Range("K2").Offset(0, -4)` | G2 | Four columns left, from TotalCharges to AdmitDateTime |
| `Range("A1").Resize(3, 2)` | A1:B3 | 3 rows by 2 columns, starting at A1 |
| `Range("A2").Resize(1, 11)` | A2:K2 | One complete data row of Encounters |
| `Range("D10").Offset(-3, 2).Resize(4, 3)` | F7:H10 | Move to F7, then make the range 4 rows by 3 columns |

Read a chain from left to right. Each step works on the result of the step before it.

> ⚠️ Moving past the edge of the sheet, as in `Range("A1").Offset(-1, 0)`, stops with run-time error 1004.

**CurrentRegion** is the block of data around a cell. Excel grows the range from that cell until every side meets a completely
empty row or column (or the edge of the sheet). On Encounters, `Range("A1").CurrentRegion` is A1:L2001. The LOSDays header in L1
is part of it even though the cells below it are empty, because a header is data too. To see a cell's current region in Excel,
select the cell, press **F5** (Mac: **Ctrl + G**), click **Special…**, and choose **Current region**.

Watch for two things with CurrentRegion:

- A completely empty row inside the data cuts the region short. Exports with blank separator rows need another method.
- Anything typed next to the data, such as a note in M1, becomes part of the region.

**UsedRange** is the rectangle from the first to the last cell the sheet has ever used, including cells that only have
formatting. It can be bigger than the data and doesn't always start at A1. Use it for a quick look, not to find where the data
ends.

**The data-body idiom.** Most of the time you want the data without its header row. Take the region, move it down one row, and
trim the extra row that now hangs off the bottom:

```vba
Dim block As Range, body As Range
Set block = ws.Range("A1").CurrentRegion                   ' header + data
Set body = block.Offset(1).Resize(block.Rows.Count - 1)    ' data only
```

This works however many rows the export has. Snippet B uses exactly this idiom.

> 💡 **Tip:** Ranges are invisible until you ask about them, so check one before your code changes it. Type
> `? Worksheets("Encounters").Range("A1").CurrentRegion.Address` in the Immediate window and press Enter.

### 6. Finding the last row

Lesson 5.2 gave you one line to copy that finds where the data ends. This lesson calls it the **last-row idiom**. Here it is
taken apart. This version names the column by its letter and stores the result in `lastR`:

```vba
lastR = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
```

1. `ws.Rows.Count` is the number of rows on a worksheet: 1,048,576.
2. `ws.Cells(1048576, "A")` is the very bottom cell of column A.
3. `.End(xlUp)` jumps up from there to the first filled cell, exactly like pressing **Ctrl + ↑** (Mac: **⌘ + ↑**).
4. `.Row` turns that cell into its row number, which is 2001 on Encounters.

Starting at the bottom and jumping up is the reliable direction, because it skips every gap above the last value.

`End` takes four directions, which match the Ctrl + arrow keys (Mac: ⌘ + arrow keys):

| Direction | Like pressing | Typical use |
|---|---|---|
| `xlUp` | Ctrl + ↑ | The last row: `ws.Cells(ws.Rows.Count, "A").End(xlUp).Row` |
| `xlToLeft` | Ctrl + ← | The last column: `ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column` |
| `xlDown` | Ctrl + ↓ | The end of the *first* block of filled cells. It stops before the first blank |
| `xlToRight` | Ctrl + → | The end of the first block along a row |

Several methods find where data ends, and each one fails in its own way:

| Method | Returns | Goes wrong when |
|---|---|---|
| `Cells(Rows.Count, "A").End(xlUp).Row` | The last filled cell in column A | The column has blanks at the bottom, or a filter hides the last rows |
| `Range("A1").End(xlDown).Row` | The last cell before the first gap | The column has any blank cell. With nothing below A1 it returns 1,048,576 |
| `Range("A1").CurrentRegion.Rows.Count` | The height of the block around A1 | A completely empty row splits the data |
| `UsedRange.Rows.Count` | The height of everything ever used | Formatting or deleted data makes it too big |
| `ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row` | The last row with anything in any column | The sheet is empty: Find returns Nothing and `.Row` fails |
| `lo.ListRows.Count` | The number of rows in an Excel Table | It only works for Tables (section 13) |

Three rules keep the idiom honest:

- **Use a column that's filled on every row.** On Encounters that's column A, the EncounterID. AdmitSource (column F) is
  blank for emergency and outpatient visits, so the idiom on column F can stop early. Snippet C shows by how much.
- **An empty column returns 1**, the header row. Code such as `Resize(lastR - 1)` then asks for 0 rows and fails with run-time
  error 1004. Check `If lastR < 2 Then Exit Sub` before you use the data.
- **Turn filters off first.** When a filter hides the last rows, `End(xlUp)` can stop at the last *visible* row instead
  (section 11).

The starter module wraps the idiom in a **function**: a procedure that hands back a value, which you set by assigning to the
function's own name. Lesson 5.4 covers functions in depth.

```vba
Function LastRow(ws As Worksheet, Optional col As Variant = 1) As Long
    LastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function

lastR = LastRow(ThisWorkbook.Worksheets("Encounters"), "A")    ' 2001
```

`Optional col As Variant = 1` means you may leave the column out, and then the function uses column 1 (A).

Two more facts about functions help with task 7, where you write one yourself. `Exit Function` leaves a function at once, just
as `Exit Sub` leaves a Sub (Lesson 5.2). And a function whose name is never assigned returns its type's starting value: 0 for
a Long, False for a Boolean. So a yes-or-no function only has to set its name to True when the answer is yes.

> ⚠️ VBA ignores case, so `lastRow` and `LastRow` are the same name to it. Lesson 5.2 stored the last row in a variable called
> `lastRow`, but once your project has the LastRow function, `Dim lastRow As Long` hides the function inside that procedure.
> The line `lastRow = LastRow(ws, "A")` then stops with the compile error *Expected array*. That's why the code in this
> lesson calls the variable `lastR`.

### 7. Looping through worksheets

`For Each` (Lesson 5.2) visits every item of any collection, so it can visit every sheet:

```vba
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
    Debug.Print ws.Index, ws.Name, ws.Visible
Next ws
```

On a fresh copy of the lesson workbook this prints ten lines, from `1  Start Here  -1` to `10  PacketTemplate  -1`. Two
things stand out. First, the loop visits hidden sheets too: lines 8 and 9 are the answer keys, with a Visible value of 0. So
test a sheet's name before your code changes it. Second, `Visible` is a number:

| Constant | Value | Meaning |
|---|:-:|---|
| `xlSheetVisible` | −1 | A normal tab |
| `xlSheetHidden` | 0 | Hidden. Right-click a tab → **Unhide…** shows it again |
| `xlSheetVeryHidden` | 2 | Hidden so that only VBA can show it again |

**Picking sheets by name pattern.** The `Like` operator compares text with a pattern:

| Pattern character | Matches | Example |
|---|---|---|
| `#` | Any one digit | `"F03" Like "F0#"` is True |
| `?` | Any one character | `"F0A" Like "F0?"` is True |
| `*` | Any number of characters, including none | `"2025-03" Like "2025-*"` is True |
| `[A-F]` | One character from a list or range | `"D7" Like "[A-F]#"` is True |

`Like` is case-sensitive, so `"f03" Like "F0#"` is False.

```vba
For Each ws In ThisWorkbook.Worksheets
    If ws.Name Like "2025-*" Then Debug.Print ws.Name     ' every monthly packet from the bonus
Next ws
```

**Index loops.** `For i = 1 To ThisWorkbook.Worksheets.Count` gives you a counter as well. When you *delete* sheets in a loop,
count backwards (`For i = ThisWorkbook.Worksheets.Count To 1 Step -1`), because deleting a sheet shifts the position of every
sheet after it.

### 8. Adding, naming, copying, and deleting sheets

| Action | Code | Notes |
|---|---|---|
| Add a sheet at the end | `Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))` | `Add` returns the new sheet. Without `Before` or `After`, it goes in front of the active sheet |
| Add before a given sheet | `Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets("Output"))` | Use `Before` or `After`, never both |
| Rename | `ws.Name = "F01"` | Follow the naming rules below |
| Copy inside the workbook | `wsTpl.Copy After:=ThisWorkbook.Worksheets("Bonus")` | The copy is named `PacketTemplate (2)` and becomes the **active** sheet |
| Copy to a new workbook | `ws.Copy` | No arguments: Excel creates a new workbook that holds only the copy (section 9) |
| Move | `ws.Move Before:=ThisWorkbook.Worksheets(1)` | Takes the same arguments as Copy |
| Delete | `ws.Delete` | Asks the user to confirm, unless DisplayAlerts is False |
| Hide or show | `ws.Visible = xlSheetHidden` | At least one sheet must stay visible |

`Copy` doesn't hand back the new sheet the way `Add` does. The copy is the active sheet straight afterwards, so grab it on the
very next line with `Set wsNew = ActiveSheet`. This is one of the few places where ActiveSheet is the right tool.

**Sheet names** must follow these rules, or the line that sets `.Name` stops with run-time error 1004:

- 1 to 31 characters.
- None of these characters: `\ / ? * [ ] :`
- It can't start or end with an apostrophe, and it can't be the reserved word `History`.
- **It must be unique in the workbook, ignoring case.** `F01` and `f01` count as the same name, so the error message reads
  *That name is already taken*.

**Deleting safely.** `ws.Delete` normally asks *Microsoft Excel will permanently delete this sheet. Do you want to continue?*
Set `Application.DisplayAlerts = False` first to accept that prompt silently, and set it back to True straight afterwards. Two
more rules apply:

- `ThisWorkbook.Worksheets("TypeSummary_2024")` stops with run-time error 9, *Subscript out of range*, when there's no sheet
  with that name. DisplayAlerts can't prevent it, because the error happens before Delete ever runs.
- Excel won't delete the last visible sheet in a workbook (run-time error 1004).

**SheetExists and the delete-and-recreate pattern.** A macro that builds a report sheet should work on its tenth run as well as
its first. The **delete-and-recreate pattern** makes that happen:

1. If the old sheet exists, delete it without a prompt.
2. Add a new sheet and name it.
3. Fill it.

Step 1 needs a way to ask "is there a sheet with this name?" that doesn't trigger error 9. Loop over the worksheets and compare
each name with the one you want. Excel ignores case in sheet names, so compare with
`StrComp(ws.Name, sheetName, vbTextCompare)`, which returns 0 when the two names match regardless of case. You'll write that
function, SheetExists, in task 7. The starter module already has the procedure that uses it:

```vba
Sub DeleteSheetIfExists(ByVal sheetName As String)
    If SheetExists(sheetName) Then
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets(sheetName).Delete
        Application.DisplayAlerts = True
    End If
End Sub
```

`ByVal` in front of the argument gives the procedure its own copy of the text, so you can pass a String variable, a Variant,
or a cell's value without a type error. Lesson 5.4 explains ByVal and ByRef.

> 💡 **Tip:** Another way to test for a sheet is to try to get it and see whether an error happens, using
> `On Error Resume Next`. That works, but it also hides every other error. Lesson 5.4 covers error handling. The loop is clearer.

### 9. Workbooks: open, create, save, and close

| Task | Code |
|---|---|
| Open a file | `Set wb = Workbooks.Open(Filename:=filePath, ReadOnly:=True)` |
| Create a blank workbook | `Set wb = Workbooks.Add` |
| Copy one sheet into a new workbook | `ws.Copy`, then `Set wb = ActiveWorkbook` |
| Save under a new name or format | `wb.SaveAs Filename:=filePath, FileFormat:=xlOpenXMLWorkbook` |
| Save in place | `wb.Save` |
| Close | `wb.Close SaveChanges:=False` (or `True`) |
| Refer to a workbook that's already open | `Workbooks("F02_encounters_2025.xlsx")`, using its file name with the extension |

`Workbooks.Open` and `Workbooks.Add` return the workbook, so keep it in a variable. `ws.Copy` doesn't, but the new workbook is
the active one straight afterwards.

**FileFormat must match the extension.** If they disagree, Excel refuses to save or writes a file that won't open:

| Constant | Value | Extension |
|---|:-:|---|
| `xlOpenXMLWorkbook` | 51 | .xlsx (no macros) |
| `xlOpenXMLWorkbookMacroEnabled` | 52 | .xlsm |
| `xlExcel12` | 50 | .xlsb (binary) |
| `xlCSVUTF8` | 62 | .csv in UTF-8 (Microsoft 365 and Excel 2019 or later) |
| `xlCSV` | 6 | .csv |

**Paths.** `ThisWorkbook.Path` is the folder the workbook is saved in. It's empty if the workbook has never been saved. Join a
folder and a file name with `Application.PathSeparator`, which is `\` on Windows and `/` on a Mac:

```vba
filePath = ThisWorkbook.Path & Application.PathSeparator & "F02_encounters_2025.xlsx"
```

> ⚠️ **ActiveWorkbook moves.** The moment your code opens or creates a workbook, that workbook is active. An unqualified
> `Worksheets("Output")` then looks in the new file and fails with run-time error 9. Write `ThisWorkbook.Worksheets("Output")`.

> ⚠️ **SaveAs over an existing file asks first.** Set DisplayAlerts to False around SaveAs so that a second run overwrites the
> file silently. If that file is open in Excel, SaveAs fails with run-time error 1004 no matter what.

> ⚠️ **Read what you need before you close.** After `wb.Close`, the variable no longer points at an open workbook, so `wb.Name`
> fails. Store the name first.

> 📋 **Mac:** Excel for Mac runs in a sandbox. The first time a macro saves or opens a file in a folder, Excel may ask you to
> grant it access to that folder. Allow it and the macro carries on.

> 📋 **OneDrive and SharePoint:** for a workbook that syncs through OneDrive, `ThisWorkbook.Path` can be a web address
> (`https://…`) instead of a folder, and building a file path from it fails. That's why section 1 asks you to save the lesson
> workbook in an ordinary local folder.

`wb.SaveCopyAs` saves a copy without changing which file you're working in. Lesson 5.5 uses it for timestamped backups.

### 10. Arrays: read once, write once

Every time VBA reads or writes a cell, it makes a round trip into Excel. One trip is quick, but a loop that reads 2 columns of
2,000 rows one cell at a time makes 4,000 of them, and writing the results makes 2,000 more. The fix is to move a whole block in
one trip:

```vba
Dim data As Variant
data = ws.Range("C2:K2001").Value      ' ONE read: a 2,000 x 9 array
```

`data` is now a **2-D array**, a grid of values in memory. You read it as `data(row, column)`, and both directions are
numbered from 1, starting at the range's first cell. So `data(1, 1)` is C2 and `data(1, 9)` is K2.

| Expression | Value |
|---|---|
| `UBound(data, 1)` | 2000, the number of rows |
| `UBound(data, 2)` | 9, the number of columns |
| `data(1, 1)` | `"Inpatient"`, from C2 |
| `data(1, 9)` | `10673.63`, from K2 |

Then loop over the array. That part is instant, because everything is already in memory:

```vba
Sub InpatientAverageCharge()
    Dim ws As Worksheet, data As Variant
    Dim i As Long, n As Long, total As Double

    Set ws = ThisWorkbook.Worksheets("Encounters")
    data = ws.Range("C2:K" & LastRow(ws, "A")).Value     ' columns C (EncounterType) to K (TotalCharges)

    For i = 1 To UBound(data, 1)
        If data(i, 1) = "Inpatient" Then                  ' array column 1 = sheet column C
            n = n + 1
            total = total + data(i, 9)                    ' array column 9 = sheet column K
        End If
    Next i
    Debug.Print n & " inpatient stays, average charge " & Format(total / n, "#,##0.00")
End Sub
```

It prints *539 inpatient stays, average charge 31,321.36*.

**Writing back.** Fill a second array, then assign it to a range of exactly the same shape:

```vba
Dim results() As Variant
ReDim results(1 To UBound(data, 1), 1 To 1)          ' one column, one row per data row
' ... fill results(i, 1) inside the loop ...
firstCell.Resize(UBound(results, 1), 1).Value = results   ' ONE write
```

Here `firstCell` is the top cell of the column you're filling.

**Declaring arrays and ReDim.** Two declarations look alike but do different jobs:

| Declaration | What it creates | Use it for |
|---|---|---|
| `Dim data As Variant` | One Variant, which can hold a whole array | Receiving a range: `data = rng.Value` |
| `Dim results() As Variant` | A **dynamic array**: an array whose size you set later | An array you fill yourself |
| `ReDim results(1 To n, 1 To 1)` | Sets the size of a dynamic array while the code runs | Sizes you only know at run time, such as `UBound(data, 1)` |

`ReDim` also empties the array, so running it again gives you a fresh, empty array of the new size. A loop that needs a clean
array on every pass can ReDim it at the top of the pass.

**One-dimensional arrays and the Array function.** `Array(...)` builds a short list on the spot, which suits a fixed set of
labels such as the four encounter types. It returns a **1-D array**, a single row of values numbered from **0**, not 1:

```vba
Dim types As Variant, t As Variant, i As Long
types = Array("Emergency", "Inpatient", "Observation", "Outpatient")
Debug.Print types(0), UBound(types)          ' Emergency     3   (four items, numbered 0 to 3)

For Each t In types                           ' the loop variable must be a Variant
    Debug.Print t
Next t
For i = 0 To UBound(types)                    ' or count from 0 when you need a position
    Debug.Print i, types(i)
Next i

ws.Range("A1:C1").Value = Array("EncounterType", "Encounters", "TotalCharges")   ' three headers across a row
```

`LBound(types)` returns the first number, 0, just as `UBound` returns the last. A `For Each` loop over an array needs a Variant
loop variable because the array's items are Variants, and `Dim t As String` stops with a compile error.

**Collecting matching rows.** To copy only some rows of the export, loop over the source array and copy each matching row into
a second array, counting as you go. Size the second array for the worst case, where every row matches, and then write only the
rows you filled. Here `wsObs` is a sheet you've added for the result:

```vba
Dim src As Variant, buf() As Variant
Dim i As Long, j As Long, n As Long

src = ws.Range("A2:K" & lastR).Value                      ' the export, columns A:K
ReDim buf(1 To UBound(src, 1), 1 To UBound(src, 2))        ' room for every row
n = 0                                                      ' rows of buf used so far
For i = 1 To UBound(src, 1)
    If src(i, 3) = "Observation" Then                      ' array column 3 = sheet column C, EncounterType
        n = n + 1                                          ' the next free row of buf
        For j = 1 To UBound(src, 2)                        ' copy the whole row, one column at a time
            buf(n, j) = src(i, j)
        Next j
    End If
Next i
If n > 0 Then wsObs.Range("A2").Resize(n, UBound(buf, 2)).Value = buf   ' ONE write: only the top n rows
```

On Encounters this collects the 89 observation stays. `Resize(n, …)` makes the target exactly n rows tall, and a range smaller
than the array takes only the array's top-left part, so the empty rows at the bottom of `buf` are never written. The `If n > 0`
test matters because `Resize(0, …)` stops with run-time error 1004.

These rules cover most array mistakes:

| Rule | Why it matters |
|---|---|
| A multi-cell range always gives a 2-D array numbered from 1, even a single column | Write `data(i, 1)`, not `data(i)` |
| A single cell gives a single value, not an array | A one-column range with only one row gives no array, and `UBound` then fails with run-time error 13. Check `IsArray(data)` when the data might be that small |
| The range you write to must match the array's shape | `Resize(UBound(arr, 1), UBound(arr, 2))` makes it fit |
| A 1-D array (such as one made with `Array(...)`) writes **across** a row | To write down a column, use a 2-D array dimensioned `(1 To n, 1 To 1)` |
| A range smaller than the array takes only the array's top-left part | Handy for a buffer that's bigger than you need, as in *Collecting matching rows* |
| `.Value` gives dates as Date values, and `.Value2` gives them as Doubles | With Value2, `Int(x)` strips the time from a date-time |

> 💡 **Tip:** Measure the difference with `Timer`, which returns the seconds since midnight. Store `t = Timer` before the work
> and `Debug.Print Timer - t` after it. On a few thousand rows both versions feel instant, but on 100,000 rows a cell-by-cell loop
> can take many seconds while the array version takes a fraction of one.

### 11. AutoFilter and SpecialCells

`Range.AutoFilter` is VBA's version of **Data → Filter**:

```vba
rng.AutoFilter Field:=4, Criteria1:="F02"
```

- `rng` is the whole block, header included. Usually that's `ws.Range("A1").CurrentRegion`.
- `Field` is a column number **within that block**. 4 means the block's fourth column, which is D (FacilityID) on Encounters.
- Calling AutoFilter again with a different Field adds a second condition, just like choosing filters in two columns.

| Filter | Code |
|---|---|
| Equals | `.AutoFilter Field:=3, Criteria1:="Inpatient"` |
| Greater than or equal to | `.AutoFilter Field:=11, Criteria1:=">=100000"` |
| Between (two conditions on one column) | `.AutoFilter Field:=11, Criteria1:=">=50000", Operator:=xlAnd, Criteria2:="<100000"` |
| Either of two values | `.AutoFilter Field:=4, Criteria1:="F01", Operator:=xlOr, Criteria2:="F03"` |
| Any value in a list | `.AutoFilter Field:=10, Criteria1:=Array("PY01", "PY03", "PY04"), Operator:=xlFilterValues` |
| Blank cells | `.AutoFilter Field:=6, Criteria1:="="` |
| Clear one column's filter | `.AutoFilter Field:=4` |

**Copying the visible rows.** `SpecialCells(xlCellTypeVisible)` returns only the cells that a filter leaves showing, header
included:

```vba
With ws.Range("A1").CurrentRegion
    .AutoFilter Field:=4, Criteria1:="F02"
    .SpecialCells(xlCellTypeVisible).Copy Destination:=wsNew.Range("A1")
End With
ws.AutoFilterMode = False              ' remove the filter
```

`Copy Destination:=` copies values, formulas, and formats straight to the destination, without selecting anything.

**Counting the visible rows.** After a filter, the visible cells form many separate blocks, which VBA calls **areas**. That makes
`.Rows.Count` on the visible cells a trap, because it counts the rows of the first area only. Use one of these instead:

```vba
nVisible = Application.WorksheetFunction.Subtotal(103, ws.Range("A1").CurrentRegion.Columns(1)) - 1
nVisible = ws.Range("A1").CurrentRegion.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1
```

`SUBTOTAL(103, …)` is COUNTA that skips hidden rows, and `.Count` counts the cells of every area. Both subtract 1 for the
header.

**Turning a filter off.**

| Code | What it does |
|---|---|
| `ws.AutoFilterMode = False` | Removes the filter arrows and shows every row |
| `If ws.FilterMode Then ws.ShowAllData` | Keeps the arrows and shows every row. ShowAllData fails with error 1004 when nothing is filtered, hence the If |

> ⚠️ **No matching rows.** When nothing matches, the header is still visible, so SpecialCells returns just the header and you copy
> only the header. If you call SpecialCells on the data rows alone and none are visible, it stops with run-time error 1004,
> *No cells were found*.

> ⚠️ **Dates in AutoFilter criteria are fragile.** Criteria are passed as text, and how Excel reads a date written as text depends
> on the date format and the computer's regional settings, so a date filter that works on your computer can fail on a
> colleague's. For a date range, pass serial numbers (`Criteria1:=">=" & CLng(DateSerial(2025, 3, 1))`) or skip the filter and
> loop over an array, as the bonus does.

> 📋 **Excel Tables:** on a Table, filter the Table's own range (`lo.Range.AutoFilter …`, section 13). A Table keeps its own
> filter, so `ws.AutoFilterMode = False` doesn't clear it. Use `If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData`.

### 12. Sorting with Range.Sort

`Range.Sort` is VBA's version of **Data → Sort**. Call it on the whole block, header included:

```vba
rng.Sort Key1:=ws.Range("K1"), Order1:=xlDescending, Header:=xlYes
```

| Argument | Meaning |
|---|---|
| `Key1` | A cell in the column to sort by. One cell is enough: `ws.Range("K1")` means column K |
| `Order1` | `xlAscending` (smallest first, the default) or `xlDescending` (largest first) |
| `Key2`, `Order2`, `Key3`, `Order3` | A second and third sort level, used to break ties |
| `Header` | `xlYes` keeps the first row in place as a header, and `xlNo` sorts every row. Avoid `xlGuess`, which lets Excel decide |

This two-level sort puts the facilities in order and, within each facility, the most expensive encounter first:

```vba
ws.Range("A1").CurrentRegion.Sort _
    Key1:=ws.Range("D1"), Order1:=xlAscending, _
    Key2:=ws.Range("K1"), Order2:=xlDescending, Header:=xlYes
```

> ⚠️ **Sort the whole block, never one column.** `ws.Range("K2:K2001").Sort …` reorders the charges and leaves every other column
> where it was, so each encounter ends up with another patient's charge. Always sort the CurrentRegion or the Table.

> ⚠️ **Try sorts on a facility sheet, not on Encounters.** Tasks 3 and 4 and the examples in this guide expect the export in its
> original order, oldest admission first. If you sort Encounters by accident, put it back with
> `ws.Range("A1").CurrentRegion.Sort Key1:=ws.Range("G1"), Order1:=xlAscending, Header:=xlYes`, which sorts by AdmitDateTime.
> Task 9 sorts the facility sheets, which are copies, so the export itself never moves.

The macro recorder (Lesson 5.1) writes a longer version that uses the sheet's **Sort object**:

```vba
With ws.Sort
    .SortFields.Clear
    .SortFields.Add2 Key:=ws.Range("K2:K2001"), Order:=xlDescending
    .SetRange ws.Range("A1:L2001")
    .Header = xlYes
    .Apply
End With
```

Both forms work in Excel. The Sort object allows more than three levels and can sort by cell color, but the recorder hard-codes
the ranges (`K2:K2001`), and those are the first thing to fix in recorded code. Older versions record `.Add` instead of
`.Add2`, and `.Add` works in every version.

### 13. Excel Tables in VBA: ListObjects

In VBA an Excel Table is a **ListObject**, and every worksheet has a `ListObjects` collection:

```vba
Dim lo As ListObject
Set lo = ThisWorkbook.Worksheets("Facilities").ListObjects("tblFacilities")
```

| Member | Returns | On tblFacilities |
|---|---|---|
| `lo.Range` | The whole Table, header included | `$A$1:$D$5` |
| `lo.HeaderRowRange` | The header row | `$A$1:$D$1` |
| `lo.DataBodyRange` | The data rows without the header | `$A$2:$D$5` |
| `lo.ListRows.Count` | The number of data rows | 4 |
| `lo.ListColumns("FacilityName").DataBodyRange` | One column's data, found by its header | `$B$2:$B$5` |
| `lo.ListColumns("FacilityName").Index` | That column's position in the Table | 2 |
| `lo.ListRows.Add` | A new empty row at the bottom, so the Table grows | A ListRow object |

You loop over one column of a Table like any other range:

```vba
Dim cell As Range
For Each cell In lo.ListColumns("FacilityName").DataBodyRange
    Debug.Print cell.Value         ' Bluestone Memorial Hospital, Ashby Falls Community Hospital, ...
Next cell
```

Tables suit VBA well. A Table always knows where it ends, so you don't need the last-row idiom. Code that finds columns by their
header names keeps working when someone inserts a new column, while code that says "column 4" quietly reads the wrong data.

> ⚠️ `lo.DataBodyRange` is `Nothing` when the Table has no data rows, and using it then stops with run-time error 91. Check
> `If lo.ListRows.Count > 0` first.

### 14. Working without Select and Activate

The macro recorder writes code that repeats what you did: select a sheet, select a cell, then act on the **Selection**. Code you
write yourself should act on objects directly. Here's what the recorder produces when you filter Encounters to Ashby Falls and
copy the result to a new sheet:

```vba
' Recorded
Sheets("Encounters").Select
Range("A1").Select
Selection.AutoFilter
ActiveSheet.Range("$A$1:$L$2001").AutoFilter Field:=4, Criteria1:="F02"
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Sheets.Add After:=ActiveSheet
ActiveSheet.Paste
```

And here's the same job written directly:

```vba
' Direct
Dim wsEnc As Worksheet, wsNew As Worksheet
Set wsEnc = ThisWorkbook.Worksheets("Encounters")
Set wsNew = ThisWorkbook.Worksheets.Add(After:=wsEnc)
With wsEnc.Range("A1").CurrentRegion
    .AutoFilter Field:=4, Criteria1:="F02"
    .SpecialCells(xlCellTypeVisible).Copy Destination:=wsNew.Range("A1")
End With
wsEnc.AutoFilterMode = False
```

The direct version is shorter and faster, because Excel doesn't redraw the screen for every selection. It also doesn't depend
on which sheet or cell happens to be active when the macro starts, and it doesn't hard-code row 2001.

| Recorded pattern | Direct version |
|---|---|
| `Sheets("Encounters").Select`, `Range("A1").Select`, `Selection.Value = …` | `wsEnc.Range("A1").Value = …` |
| `Range(Selection, Selection.End(xlDown)).Select` | `wsEnc.Range("A1").CurrentRegion`, or the last-row idiom |
| `Selection.Copy`, select the target, `ActiveSheet.Paste` | `source.Copy Destination:=target` |
| `Sheets.Add`, then `ActiveSheet.Name = …` | `Set ws = ThisWorkbook.Worksheets.Add(…)`, then `ws.Name = …` |
| `ActiveWorkbook.Save` | `ThisWorkbook.Save`, or the variable that holds the other workbook |

Use `Select` and `Activate` only for what the user needs to see, such as showing the finished report at the end of a macro with
`wsReport.Activate`. Use `ActiveSheet` and `ActiveWorkbook` only on the line straight after `Copy`, to grab the new object.

> 💡 **Tip:** To copy values only, without formats and without the clipboard, assign them:
> `target.Resize(source.Rows.Count, source.Columns.Count).Value = source.Value`.

### 15. Speed and safety settings

Four Application properties make big macros faster and quieter. Each one changes Excel for the user too, so you must put it
back.

| Setting | What it controls | While your macro works | Does Excel put it back when the code ends? |
|---|---|---|---|
| `Application.ScreenUpdating` | Redrawing the screen after every change | `False` | Usually, but restore it yourself |
| `Application.Calculation` | Recalculating formulas after every change | `xlCalculationManual` | **No** |
| `Application.EnableEvents` | Running event macros when cells change (Lesson 5.5) | `False` | **No** |
| `Application.DisplayAlerts` | Confirmation prompts such as *delete this sheet?* | `False`, only around Delete or SaveAs | Yes |

`Application.StatusBar = "Building packets…"` shows a progress message, and `Application.StatusBar = False` gives the status
bar back to Excel. It also stays until you reset it.

Calculation and EnableEvents stay changed after the macro ends, even if it stopped with an error. A colleague's workbook that
"stopped recalculating" is often the victim of a macro that set Calculation to manual and then crashed. So use the
**safe-restore pattern**: remember the user's settings, change them, send any error to a cleanup block, and restore everything
there.

```vba
Sub SafeSettingsPattern()
    Dim oldCalc As XlCalculation
    oldCalc = Application.Calculation               ' remember the user's setting

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    On Error GoTo CleanUp                           ' any error jumps to CleanUp

    ' ... the real work ...

CleanUp:                                            ' reached after success AND after an error
    Application.Calculation = oldCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then MsgBox "Stopped: " & Err.Description, vbExclamation
End Sub
```

When the work succeeds, the code runs straight on into CleanUp, and `Err.Number` is still 0. When something fails, VBA jumps to
CleanUp with the error's details in `Err`. Lesson 5.4 explains `On Error` fully. SplitByFacility and the bonus use this pattern.

> 💡 **Tip:** If a macro crashed and Excel now ignores event macros or won't recalculate, type
> `Application.EnableEvents = True` or `Application.Calculation = xlCalculationAutomatic` in the Immediate window and press
> Enter.

### 16. Worked example: the high-charge list

*Task: Finance wants a sheet that lists every 2025 encounter with charges of \$100,000 or more, most expensive first. It must be
rebuilt from scratch each time the macro runs.*

Write the plan in plain words first: delete the old HighCharges sheet, add a new one, filter Encounters to TotalCharges of at
least 100,000, copy the visible rows, turn the filter off, sort the new sheet, and report the count. Each step comes from a
section of this guide:

```vba
Sub HighChargeList()
    Dim wsEnc As Worksheet, wsHi As Worksheet
    Dim n As Long

    Set wsEnc = ThisWorkbook.Worksheets("Encounters")
    If wsEnc.AutoFilterMode Then wsEnc.AutoFilterMode = False     ' start unfiltered (section 11)

    DeleteSheetIfExists "HighCharges"                             ' delete-and-recreate (section 8)
    Set wsHi = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsHi.Name = "HighCharges"

    With wsEnc.Range("A1").CurrentRegion                          ' the whole export (section 5)
        .AutoFilter Field:=11, Criteria1:=">=100000"              ' column K, TotalCharges
        .SpecialCells(xlCellTypeVisible).Copy Destination:=wsHi.Range("A1")
    End With
    wsEnc.AutoFilterMode = False                                  ' leave the export as you found it

    With wsHi.Range("A1").CurrentRegion
        .Sort Key1:=wsHi.Range("K1"), Order1:=xlDescending, Header:=xlYes    ' section 12
        n = .Rows.Count - 1                   ' one solid block now, so Rows.Count is safe
        .Columns.AutoFit
    End With
    MsgBox n & " encounters of $100,000 or more.", vbInformation, "High-charge list"
End Sub
```

1. Paste it into the EncounterMacros module, below the other procedures. It calls DeleteSheetIfExists, which calls your
   SheetExists, so finish SheetExists (task 7) first.
2. Click inside the macro and press **F8** (Mac: **Debug → Step Into**) a few times. Watch Encounters filter and the new sheet
   fill, then press **F5** (Mac: **Run → Run Sub/UserForm**) to finish.
3. The message reads *12 encounters of \$100,000 or more.* All 12 are inpatient stays, and the top row is ENC113090, an Ashby
   Falls stay with \$185,019.66 in charges.
4. Cross-check with a formula in an empty cell on the Output sheet: `=COUNTIF(Encounters!K2:K2001,">=100000")` returns 12.
   Don't type it on Encounters itself, because a cell next to the data becomes part of its CurrentRegion (section 5).
5. Run the macro again. The old sheet is replaced rather than duplicated, and no prompt appears.

Task 8 uses the same steps inside a loop, once for each facility.

## 🧪 Hands-on practice

Set up the workbook as described in section 1. Then type each prediction in a yellow cell on the **Practice** sheet, or run
your macro so the gray cell fills in by itself. The **Check** column turns green when you're right.

Tasks 1–6 ask what each snippet below prints, or which error stops it. The same code is on the workbook's **Snippets** sheet and
in `starter/Snippets.bas`. Commit to a prediction, then run the snippet: click inside it and press **F5**
(Mac: **Run → Run Sub/UserForm**), and compare the Immediate window with your prediction. Snippets D and E are supposed to
fail: read the error number, then click **End**.

**Snippet A** (task 1)

```vba
Sub SnippetA_OffsetResize()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Encounters")
    Debug.Print ws.Range("C2").Offset(3, 4).Resize(5).Address
End Sub
```

**Snippet B** (task 2)

```vba
Sub SnippetB_DataBody()
    Dim ws As Worksheet, block As Range
    Set ws = ThisWorkbook.Worksheets("Encounters")
    Set block = ws.Range("E50").CurrentRegion            ' any cell inside the data works
    Set block = block.Offset(1).Resize(block.Rows.Count - 1)
    Debug.Print block.Address(False, False)
End Sub
```

**Snippet C** (tasks 3 and 4)

```vba
Sub SnippetC_LastRow()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Encounters")
    Debug.Print ws.Cells(ws.Rows.Count, "A").End(xlUp).Row     ' line 1: EncounterID
    Debug.Print ws.Cells(ws.Rows.Count, "F").End(xlUp).Row     ' line 2: AdmitSource
    Debug.Print ws.Range("F1").End(xlDown).Row                 ' line 3: AdmitSource again
End Sub
```

**Snippet D** (task 5)

```vba
Sub SnippetD_WhichSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Encounters")
    ThisWorkbook.Worksheets("Practice").Activate               ' Practice is now the active sheet
    Debug.Print ws.Range(Cells(2, 1), Cells(2, 11)).Address    ' row 2, columns A to K
End Sub
```

**Snippet E** (task 6)

```vba
Sub SnippetE_DeleteOld()
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("TypeSummary_2024").Delete         ' last year's summary sheet
    Application.DisplayAlerts = True
    Debug.Print "Old summary deleted"
End Sub
```

Tasks 7–12 build on each other, so do them in order. SplitByFacility (task 8) creates the facility sheets that tasks 9–11 use.

<!-- BEGIN GENERATED: practice -->
Save the workbook as .xlsm, then import starter/EncounterMacros.bas and starter/Snippets.bas (VBE → File → Import File…). Tasks 1–6 ask what the code on the Snippets sheet prints: type your prediction, then run the snippet to check it. In tasks 7–12 you complete macros, and the gray cells read the sheets and cells your macros create.

| # | Task | Hint |
|:-:|------|------|
| 1 | Snippet A: what address does the Immediate window show? Type it as printed (the \$ signs are optional). | Offset moves the top-left corner, and Resize sets the size from that corner |
| 2 | Snippet B: what address does the Immediate window show? | CurrentRegion is the whole block around E50 (Go To Special → Current region shows it). The last two steps drop the header row |
| 3 | Snippet C: line 1 prints 2001, the last row of column A. What number does line 2 print (the same idiom on column F, AdmitSource)? | Which encounter types leave AdmitSource blank? Look at the bottom of column F |
| 4 | Snippet C: what number does line 3 print? | End(xlDown) works like Ctrl + ↓: it stops at the last filled cell before a gap |
| 5 | Snippet D: running SnippetD_WhichSheet stops with a run-time error. Type the error number. | Which sheet does a Cells(…) with nothing in front of it belong to? |
| 6 | Snippet E: running SnippetE_DeleteOld stops with a run-time error. Type the error number. | What does a collection do when you ask for an item it doesn't have? |
| 7 | Complete SheetExists and BuildTypeSummary in the starter module. BuildTypeSummary deletes any old TypeSummary sheet (with DeleteSheetIfExists, which calls your SheetExists), adds a new sheet named TypeSummary at the end of the workbook, and writes one row per EncounterType (Emergency, Inpatient, Observation, Outpatient) with the columns EncounterType (A), Encounters (B, the count), and TotalCharges (C), headers in row 1. Run it twice: the second run must not stop with an error or ask a question. The gray cell looks up the Inpatient TotalCharges on your sheet. | SheetExists: For Each ws In ThisWorkbook.Worksheets … StrComp(ws.Name, sheetName, vbTextCompare) = 0 |
| 8 | Complete SplitByFacility: for each FacilityID in the Facilities table (tblFacilities), delete any old sheet with that name, add a new sheet named after the ID (F01, F02, F03, F04), AutoFilter the Encounters block on FacilityID, and copy the visible cells (header included) to A1 of the new sheet. Turn the filter off at the end. The gray cell lists the number of data rows on F01, F02, F03, and F04. | block.AutoFilter Field:=4, Criteria1:=facID, then block.SpecialCells(xlCellTypeVisible).Copy |
| 9 | Complete SortAndReconcile, part 1: loop through every worksheet and, for each sheet whose name is Like "F0#", sort its block by TotalCharges (column K), largest first, with Range.Sort and a header row. Which EncounterID is now in A2 of sheet F03? (The gray cell reads F03!A2.) | Test ws.Name Like "F0#" inside the loop, then sort ws.Range("A1").CurrentRegion with Key1 in column K and Order1:=xlDescending (guide section 12) |
| 10 | SortAndReconcile, part 2: in the same loop, add each facility sheet's data rows and TotalCharges to two running totals, then write the row total to Output!B4 and the charge total to Output!B5. What is the total charge on the facility sheets? (The gray cell reads Output!B5.) | Find each sheet's last row with LastRow(ws, "A"), then WorksheetFunction.Sum the K cells |
| 11 | Complete ExportFacility: copy sheet F02 into a new workbook (Worksheet.Copy with no arguments), save it in the same folder as this workbook as F02_encounters_2025.xlsx with SaveAs and FileFormat:=xlOpenXMLWorkbook, and close it. Then reopen the file with Workbooks.Open, add up its TotalCharges column, write the total to Output!B7 (and the file name to Output!B6), and close it without saving. The gray cell reads Output!B7. | After Worksheet.Copy the new workbook is the ActiveWorkbook, so write results through ThisWorkbook |
| 12 | Complete FillLOSDays: read AdmitDateTime and DischargeDateTime (Encounters G2:H2001) into a Variant array with one .Value2 read, calculate each row's length of stay in days with Int(discharge) - Int(admit) (the number of midnights, ignoring the times), and write all the results to the yellow LOSDays column (L) with one assignment. The gray cell adds up your column. | stay = ws.Range("G2:H" & lastR).Value2, then ReDim los(1 To UBound(stay, 1), 1 To 1) |
<!-- END GENERATED: practice -->

## ✅ Answer key

The workbook has a hidden **Answer Key** sheet (right-click any sheet tab → **Unhide…** → *Answer Key*). Its *Live result*
column checks every macro answer with a worksheet formula. Complete reference macros are in
[`solutions/EncounterMacros_Solution.bas`](solutions/EncounterMacros_Solution.bas) (spoilers). Import it into a spare copy of the
workbook so its procedure names don't clash with yours. The answers are also below, collapsed so you don't see them by accident.

<!-- BEGIN GENERATED: answers -->
<details>
<summary><b>🔑 Show the answer key</b> — Try every task before opening this.</summary>

**1. Snippet A: Offset and Resize**

- **Answer:** \$G\$5:\$G\$9
- **Solution:** `Range("C2").Offset(3, 4)` moves 3 rows down and 4 columns right, to **G5**. `Resize(5)` makes the range 5 rows tall from that corner and keeps its width of 1 column: **\$G\$5:\$G\$9**.

Offset never changes a range's size, and Resize never moves its top-left corner, so you can read a chain of them left to right. Leaving out Resize's second argument keeps the column count as it is. `.Address` returns absolute references (\$G\$5) unless you ask for `.Address(False, False)`.

**2. Snippet B: the data-body idiom**

- **Answer:** A2:L2001
- **Solution:** `Range("E50").CurrentRegion` grows from E50 to the whole block bounded by empty rows and columns: A1:L2001 (the LOSDays header in L1 makes the block 12 columns wide). `Offset(1)` shifts the block down one row, and `Resize(Rows.Count - 1)` trims the row that fell off the bottom: **A2:L2001**.

This **data-body idiom** gives you the data without its header, however many rows the export has. It works from any cell inside the block, which is why the snippet starts in E50. CurrentRegion stops at the first completely empty row and column, so a blank row in the middle of an export cuts it short.

**3. Snippet C, line 2: the last row of AdmitSource**

- **Answer:** 1995
- **Solution:** End(xlUp) from the bottom of column F stops at the last cell with something in it: row **1995**. The last 6 encounters are emergency or outpatient visits, which have no AdmitSource.

The last-row idiom finds the last filled cell **in the column you give it**. AdmitSource is blank for emergency and outpatient visits, so a loop that stopped at this row would silently skip the last 6 encounters. Always use a column that is filled on every row, such as the ID column.

**4. Snippet C, line 3: End(xlDown)**

- **Answer:** 2
- **Solution:** F1 and F2 are both filled and F3 is empty, so End(xlDown) stops at the end of that first block: row **2**.

`Range("F1").End(xlDown)` is the top-down version of the idiom, and it fails on the first blank cell. On a column with no gaps it gives the same answer as End(xlUp), which is why it seems to work until the day a blank appears. Start from the bottom and go up with `Cells(Rows.Count, col).End(xlUp)`.

**5. Snippet D: the unqualified Cells error**

- **Answer:** 1004
- **Solution:** Run-time error **1004**: *Method 'Range' of object '_Worksheet' failed*. The two bare `Cells(...)` belong to the active sheet (Practice), and `ws.Range(...)` can't build an Encounters range from Practice cells.

In a standard module, `Range` and `Cells` with no object in front mean *the active sheet*. Qualify every one: `ws.Range(ws.Cells(2, 1), ws.Cells(2, 11))`, or use a With block and a leading dot: `.Range(.Cells(2, 1), .Cells(2, 11))`. The same macro works when Encounters happens to be active, which is what makes this bug hard to spot.

**6. Snippet E: deleting a sheet that isn't there**

- **Answer:** 9
- **Solution:** Run-time error **9**: *Subscript out of range*. There is no sheet named TypeSummary_2024, so `Worksheets("TypeSummary_2024")` fails before `.Delete` even runs.

`DisplayAlerts = False` only silences Excel's *confirmation* prompt. It doesn't make a missing sheet exist. That's why delete-and-recreate macros check first with a function such as **SheetExists**. The macro also stopped before its `DisplayAlerts = True` line. Excel resets DisplayAlerts by itself when code finishes, but it doesn't do that for Calculation or EnableEvents, which is why restore lines belong in cleanup code that always runs (guide section 15).

**7. SheetExists + BuildTypeSummary (Inpatient TotalCharges)**

- **Answer:** 16,882,215.53
- **Solution:**

```vba
Function SheetExists(ByVal sheetName As String) As Boolean
    ' True if THIS workbook has a worksheet with that name (ignoring case,
    ' because Excel treats "typesummary" and "TypeSummary" as the same name).
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, sheetName, vbTextCompare) = 0 Then
            SheetExists = True
            Exit Function
        End If
    Next ws
End Function

Sub BuildTypeSummary()
    Dim wsEnc As Worksheet, wsSum As Worksheet
    Dim rngType As Range, rngCharges As Range
    Dim types As Variant
    Dim lastR As Long, i As Long

    Set wsEnc = ThisWorkbook.Worksheets("Encounters")
    lastR = LastRow(wsEnc, "A")
    Set rngType = wsEnc.Range("C2:C" & lastR)          ' EncounterType
    Set rngCharges = wsEnc.Range("K2:K" & lastR)       ' TotalCharges

    ' Delete-and-recreate: the macro can run again and again
    DeleteSheetIfExists "TypeSummary"
    Set wsSum = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsSum.Name = "TypeSummary"

    wsSum.Range("A1:C1").Value = Array("EncounterType", "Encounters", "TotalCharges")
    types = Array("Emergency", "Inpatient", "Observation", "Outpatient")
    For i = 0 To UBound(types)                          ' Array() numbers its items from 0
        With wsSum.Cells(i + 2, 1)                      ' rows 2 to 5, column A
            .Value = types(i)
            .Offset(0, 1).Value = Application.WorksheetFunction.CountIf(rngType, types(i))
            .Offset(0, 2).Value = Application.WorksheetFunction.SumIf(rngType, types(i), rngCharges)
        End With
    Next i

    wsSum.Range("A1:C1").Font.Bold = True
    wsSum.Range("C2:C5").NumberFormat = "#,##0.00"
    wsSum.Columns("A:C").AutoFit
End Sub
```


**Delete-and-recreate** makes a macro safe to rerun: without the delete, the second run stops at `wsSum.Name = "TypeSummary"` with run-time error 1004 (*That name is already taken*), because sheet names must be unique. `Worksheets.Add` returns the new sheet, so `Set wsSum = …` gives you a variable for it and you never need ActiveSheet. SheetExists compares with `vbTextCompare` because Excel sheet names ignore case. The finished sheet: Emergency 663 encounters / \$1,753,021.67, Inpatient 539 encounters / \$16,882,215.53, Observation 89 encounters / \$702,704.46, Outpatient 709 encounters / \$899,052.36. Inpatient stays are 27% of the encounters but 83% of the charges.

**8. SplitByFacility (data rows per facility sheet)**

- **Answer:** 928 / 175 / 188 / 709
- **Solution:**

```vba
Sub SplitByFacility()
    Dim wsEnc As Worksheet, wsNew As Worksheet
    Dim lo As ListObject
    Dim idCell As Range, block As Range
    Dim facID As String

    Set wsEnc = ThisWorkbook.Worksheets("Encounters")
    Set lo = ThisWorkbook.Worksheets("Facilities").ListObjects("tblFacilities")
    If wsEnc.AutoFilterMode Then wsEnc.AutoFilterMode = False   ' start with no filter
    Set block = wsEnc.Range("A1").CurrentRegion                 ' header + every data row

    Application.ScreenUpdating = False
    On Error GoTo CleanUp                                       ' always restore settings

    For Each idCell In lo.ListColumns("FacilityID").DataBodyRange
        facID = idCell.Value
        DeleteSheetIfExists facID
        Set wsNew = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsNew.Name = facID

        block.AutoFilter Field:=4, Criteria1:=facID               ' column D = FacilityID
        block.SpecialCells(xlCellTypeVisible).Copy Destination:=wsNew.Range("A1")
        wsNew.Range("A1").CurrentRegion.Columns.AutoFit
    Next idCell

CleanUp:
    wsEnc.AutoFilterMode = False                                ' leave the export unfiltered
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then MsgBox "SplitByFacility stopped: " & Err.Description, vbExclamation
End Sub
```


AutoFilter hides the rows that don't match, and `SpecialCells(xlCellTypeVisible)` picks only the rows still showing, so the copy brings the header and that facility's rows and nothing else. Looping over the Table's FacilityID column means a fifth facility would get its own sheet with no code change. The `On Error GoTo CleanUp` block turns the filter off and ScreenUpdating back on even if something fails. If a count is one too high or too low, check whether your copy included the header row, because the gray cell subtracts 1 for it.

**9. SortAndReconcile: top EncounterID on F03**

- **Answer:** ENC112749
- **Solution:**

```vba
Sub SortAndReconcile()
    Dim ws As Worksheet
    Dim lastR As Long, rowTotal As Long, sheetCount As Long
    Dim chargeTotal As Double

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like "F0#" Then                  ' F01 to F09 (# = any one digit)
            ' Task 9: largest charge first. One cell is enough to name the sort column.
            ws.Range("A1").CurrentRegion.Sort Key1:=ws.Range("K1"), Order1:=xlDescending, Header:=xlYes

            ' Task 10: add this sheet to the control totals
            lastR = LastRow(ws, "A")
            rowTotal = rowTotal + (lastR - 1)       ' minus the header row
            chargeTotal = chargeTotal + Application.WorksheetFunction.Sum(ws.Range("K2:K" & lastR))
            sheetCount = sheetCount + 1
        End If
    Next ws

    With ThisWorkbook.Worksheets("Output")
        .Range("B4").Value = rowTotal
        .Range("B5").Value = chargeTotal
    End With
    MsgBox sheetCount & " facility sheets: " & rowTotal & " rows, " & _
           Format(chargeTotal, "#,##0.00") & " in charges.", vbInformation, "Reconcile the split"
End Sub
```


`For Each ws In ThisWorkbook.Worksheets` visits every worksheet, including the hidden key sheets, so the `Like "F0#"` test (# means any one digit) picks out the facility sheets. `Header:=xlYes` keeps row 1 in place as a header. Leave it out and Range.Sort treats row 1 as data, so an ascending sort would bury the header among the rows. ENC112749 is Cedar Ridge's most expensive 2025 encounter in the sample (\$167,509.41, inpatient).

**10. SortAndReconcile: control total of charges**

- **Answer:** 20,236,994.02
- **Solution:** The second half of the loop in task 9's macro: `rowTotal = rowTotal + (lastR - 1)` and `chargeTotal = chargeTotal + Application.WorksheetFunction.Sum(ws.Range("K2:K" & lastR))`, then both are written to Output after `Next ws`.

This is a **reconciliation**: the facility sheets together must hold exactly the 2,000 rows and \$20,236,994.02 of the export (compare with `=SUM(Encounters!K2:K2001)`). If the totals differ, the split lost or duplicated rows, and you know before anyone reads the sheets. The same control-total habit applies to any macro that moves data around.

**11. ExportFacility (F02 total read back from the saved file)**

- **Answer:** 2,357,196.80
- **Solution:**

```vba
Sub ExportFacility()
    Const FAC_ID As String = "F02"
    Dim wbNew As Workbook, wbCheck As Workbook
    Dim filePath As String, savedName As String
    Dim lastR As Long
    Dim total As Double

    If Not SheetExists(FAC_ID) Then
        MsgBox "There's no " & FAC_ID & " sheet. Run SplitByFacility first.", vbExclamation
        Exit Sub
    End If
    If ThisWorkbook.Path = "" Then
        MsgBox "Save this workbook first, so the export has a folder to go in.", vbExclamation
        Exit Sub
    End If
    filePath = ThisWorkbook.Path & Application.PathSeparator & FAC_ID & "_encounters_2025.xlsx"

    ' 1. Copy the sheet into a brand-new workbook and save that as .xlsx
    ThisWorkbook.Worksheets(FAC_ID).Copy            ' no Before/After: Excel makes a new workbook
    Set wbNew = ActiveWorkbook                      ' right after Copy, the new workbook is active
    Application.DisplayAlerts = False               ' replace last run's file without asking
    wbNew.SaveAs Filename:=filePath, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    savedName = wbNew.Name                          ' read it BEFORE closing
    wbNew.Close SaveChanges:=False

    ' 2. Reopen the saved file, read it, and close it again
    Set wbCheck = Workbooks.Open(Filename:=filePath, ReadOnly:=True)
    With wbCheck.Worksheets(1)
        lastR = .Cells(.Rows.Count, "A").End(xlUp).Row
        total = Application.WorksheetFunction.Sum(.Range("K2:K" & lastR))
    End With
    wbCheck.Close SaveChanges:=False

    ' 3. Report in THIS workbook, whichever workbook is active now
    With ThisWorkbook.Worksheets("Output")
        .Range("B6").Value = savedName
        .Range("B7").Value = total
    End With
End Sub
```


Once a second workbook is open, **ActiveWorkbook** and **ThisWorkbook** are different files. ThisWorkbook is always the file that holds the code, so it's the safe way back to Output. `FileFormat:=xlOpenXMLWorkbook` (51) must match the .xlsx extension, or Excel refuses or saves a file that won't open. DisplayAlerts False lets a second run overwrite the old file. Opening with `ReadOnly:=True` and closing with `SaveChanges:=False` guarantees the check never changes the export.

**12. FillLOSDays (sum of the LOSDays column)**

- **Answer:** 2,580
- **Solution:**

```vba
Sub FillLOSDays()
    Dim ws As Worksheet
    Dim lastR As Long, i As Long
    Dim stay As Variant          ' will hold an n x 2 array: admit, discharge
    Dim los() As Variant         ' n x 1 array for the results

    Set ws = ThisWorkbook.Worksheets("Encounters")
    lastR = LastRow(ws, "A")

    stay = ws.Range("G2:H" & lastR).Value2          ' ONE read. Value2 gives dates as serial numbers
    ReDim los(1 To UBound(stay, 1), 1 To 1)

    For i = 1 To UBound(stay, 1)
        ' whole days between the two dates = midnights in the hospital
        los(i, 1) = Int(stay(i, 2)) - Int(stay(i, 1))
    Next i

    ws.Range("L2").Resize(UBound(los, 1), 1).Value = los   ' ONE write
End Sub
```


Every read or write of a cell crosses from VBA into Excel, and that crossing is the slow part. The array version crosses twice (one read, one write) instead of 6,000 times (4,000 reads and 2,000 writes) for 2,000 rows. A multi-cell range's `.Value2` is always a **2-D** array numbered from 1, even for one column, so the results array must be 2-D too: `los(1 To n, 1 To 1)`. Value2 hands back dates as serial numbers, so `Int` strips the time. Inpatient stays account for 2,406 of the 2,580 days. Emergency and observation visits that crossed midnight count 1 or 2.

</details>
<!-- END GENERATED: answers -->

## 🏆 Bonus challenge

<!-- BEGIN GENERATED: bonus -->
Finance wants a monthly packet of inpatient encounters, one sheet per month of 2025, each built from the PacketTemplate sheet. Complete BuildMonthlyPackets so that, for each month 1 to 12, it does four things:
1. Deletes any old sheet named for the month in yyyy-mm form (for example 2025-03), copies PacketTemplate to the end of the workbook, and renames the copy to that name.
2. Writes that month's Inpatient encounters (AdmitDateTime in the month) to the packet: all 11 columns A:K, in the same order as Encounters. The template's headers are in row 6, so data starts in A7. Use arrays: read the Encounters data once, before the month loop, collect each month's rows in a second array, and write them with one assignment.
3. Sorts the packet's rows by TotalCharges, largest first.
4. Fills the header block: B2 the first day of the month, B3 the number of encounters, and B4 their total charges.

After the loop, write the name of the packet sheet with the largest total charges to Output!B9 (it's formatted as Text, so Excel keeps 2025-xx as text and doesn't turn it into a date). Turn ScreenUpdating off while the macro runs, and make sure it comes back on even if something fails.

Work on the **Bonus** sheet of the workbook.

- **B1.** How many inpatient encounters are in the 2025-03 packet? (The gray cell counts the data rows from row 7 down.) *(Hint: Test src(i, 3) = "Inpatient" And src(i, 7) >= monthStart And src(i, 7) < nextMonth)*
- **B2.** What total charges does the 2025-07 packet show in its header (cell B4)? *(Hint: Reset total = 0 at the start of every month)*
- **B3.** What is the top charge in the 2025-11 packet (TotalCharges in K7, after sorting)? *(Hint: Row 5 of the template is empty, so CurrentRegion from A6 is just the table)*
- **B4.** Which month's packet has the largest total charges? (The gray cell reads Output!B9.) *(Hint: The running-maximum pattern from Lesson 5.2, one level up: once per month, not once per row)*
<!-- END GENERATED: bonus -->

The reference solution is in [`solutions/MonthlyPackets_Solution.bas`](solutions/MonthlyPackets_Solution.bas) (spoilers).

<!-- BEGIN GENERATED: bonus-answers -->
<details>
<summary><b>🔑 Show the bonus solution</b> — Give it a real try first!</summary>

**B1. Encounters in the 2025-03 packet**

- **Answer:** 37
- **Solution:**

```vba
Sub BuildMonthlyPackets()
    Const NCOLS As Long = 11                     ' Encounters columns A:K
    Dim wsEnc As Worksheet, wsTpl As Worksheet, wsPk As Worksheet
    Dim src As Variant, buf() As Variant
    Dim lastR As Long, i As Long, j As Long, n As Long, m As Long
    Dim monthStart As Date, nextMonth As Date
    Dim sheetName As String, bestName As String
    Dim total As Double, bestTotal As Double

    Set wsEnc = ThisWorkbook.Worksheets("Encounters")
    Set wsTpl = ThisWorkbook.Worksheets("PacketTemplate")
    lastR = LastRow(wsEnc, "A")
    src = wsEnc.Range("A2").Resize(lastR - 1, NCOLS).Value   ' read the export ONCE

    Application.ScreenUpdating = False
    On Error GoTo CleanUp

    For m = 1 To 12
        monthStart = DateSerial(2025, m, 1)
        nextMonth = DateSerial(2025, m + 1, 1)   ' month 13 rolls over to January 2026
        sheetName = Format(monthStart, "yyyy-mm")

        ' 1. Collect this month's inpatient rows in a buffer as big as the export
        ReDim buf(1 To UBound(src, 1), 1 To NCOLS)
        n = 0
        total = 0
        For i = 1 To UBound(src, 1)
            If src(i, 3) = "Inpatient" And src(i, 7) >= monthStart And src(i, 7) < nextMonth Then
                n = n + 1
                For j = 1 To NCOLS
                    buf(n, j) = src(i, j)
                Next j
                total = total + src(i, 11)
            End If
        Next i

        ' 2. A fresh copy of the template, named for the month
        DeleteSheetIfExists sheetName
        wsTpl.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
        Set wsPk = ActiveSheet                   ' a copied sheet is always the active one
        wsPk.Name = sheetName

        ' 3. Header block, data, sort
        wsPk.Range("B2").Value = monthStart
        wsPk.Range("B3").Value = n
        wsPk.Range("B4").Value = total
        If n > 0 Then
            ' A range smaller than the array takes just the top-left n rows of it
            wsPk.Range("A7").Resize(n, NCOLS).Value = buf
            wsPk.Range("A6").CurrentRegion.Sort Key1:=wsPk.Range("K6"), Order1:=xlDescending, Header:=xlYes
        End If

        ' 4. Remember the biggest month so far
        If total > bestTotal Then
            bestTotal = total
            bestName = sheetName
        End If
    Next m

    ThisWorkbook.Worksheets("Output").Range("B9").Value = bestName

CleanUp:
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        MsgBox "BuildMonthlyPackets stopped: " & Err.Description, vbExclamation
    Else
        MsgBox "12 packets built. Largest month: " & bestName, vbInformation
    End If
End Sub
```


Comparing against the **first day of the next month** with `<` catches every admission on the last day of the month, whatever the time. `<= DateSerial(2025, 3, 31)` would miss a patient admitted at 3/31 14:20, because that datetime is larger than midnight on 3/31. `DateSerial(2025, m + 1, 1)` works for December too: month 13 rolls over to January 2026. The `ReDim buf(1 To UBound(src, 1), …)` buffer is as big as the whole export, and `Range("A7").Resize(n, 11).Value = buf` writes only its top-left n rows.

**B2. Header total on the 2025-07 packet**

- **Answer:** 1,362,384.70
- **Solution:** Add each matching row's TotalCharges (`src(i, 11)`) to `total` inside the collecting loop, then write `wsPk.Range("B4").Value = total` after copying the template.

If every month after January shows a bigger total than the one before, `total` wasn't reset inside the month loop, so each packet carries the previous months' charges. The same goes for `n`. Accumulators that belong to one pass of an outer loop must be reset at the top of that pass.

**B3. Top charge on the 2025-11 packet**

- **Answer:** 144,338.01
- **Solution:** Sort the block that starts at the header row: `wsPk.Range("A6").CurrentRegion.Sort Key1:=wsPk.Range("K6"), Order1:=xlDescending, Header:=xlYes`.

The template leaves row 5 empty on purpose. CurrentRegion stops at an empty row, so `Range("A6").CurrentRegion` is the header plus the data and never includes the labels in A2:B4. Without that empty row, the region would reach up to the title and labels in rows 1–4, and the sort would shuffle them in among the encounters. ENC120639 tops the 2025-11 packet at \$144,338.01.

**B4. Packet with the largest total charges**

- **Answer:** 2025-11
- **Solution:** Track a running maximum in the month loop. After each month, `If total > bestTotal Then` store `total` in `bestTotal` and `sheetName` in `bestName`. After `Next m`, write `bestName` to Output!B9.

**2025-11**: \$1,925,363.45 from 53 stays. The quietest month was 2025-06 (\$868,931.28, 31 stays). Output!B9 is formatted as Text because Excel may read a value like 2025-01 written into a General cell as a date (January 2025). The check accepts either form. In the solution, the `On Error GoTo CleanUp` line sends any error to the cleanup block, which turns ScreenUpdating back on before reporting the problem.

</details>
<!-- END GENERATED: bonus-answers -->

## Key takeaways

- Excel is a hierarchy of objects, from Application down through Workbooks and Worksheets to Range. Start from
  **ThisWorkbook**, and keep any object you use twice in a variable assigned with `Set`.
- Qualify every `Range` and `Cells`, including the ones inside `ws.Range(...)`. An unqualified reference belongs to whichever
  sheet or workbook is active.
- Find where data ends with `Cells(Rows.Count, col).End(xlUp)` on a column that's filled on every row, and shape ranges with
  **CurrentRegion**, **Offset**, and **Resize** so macros work on any number of rows.
- Loop over sheets with `For Each ws In ThisWorkbook.Worksheets`, and make report macros safe to rerun with
  **delete-and-recreate**: check with SheetExists, delete with DisplayAlerts off, then Add and name the new sheet.
- Move data in bulk: read a range into an array once, calculate in memory, and write back once. Extract rows with
  **AutoFilter** and **SpecialCells(xlCellTypeVisible)**, and order them with **Range.Sort** on the whole block.
- Don't Select. Act on objects directly, and restore ScreenUpdating, Calculation, and EnableEvents in a cleanup block that runs
  even after an error.

<!-- BEGIN GENERATED: nav -->
---

⬅️ **Previous:** [5.2 VBA Fundamentals](../02-vba-fundamentals/README.md) · 🏠 [Course home](../../README.md) · **Next:** [5.4 VBA: Custom Functions, Dictionaries & Error Handling](../04-vba-functions-error-handling/README.md) ➡️
<!-- END GENERATED: nav -->
