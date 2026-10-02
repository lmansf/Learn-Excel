# Lesson 5.5 · Events, UserForms & Automated Reports

> **Level:** Expert · **Time:** about 75 minutes · **Workbook:** [`5.5-events-userforms-automation.xlsx`](5.5-events-userforms-automation.xlsx)
> **Data:** Bluestone Memorial Hospital: the 4 West bed board (36 beds, 12/31/2025) and the ED intake log (Dec 25–31, 2025), plus 12 monthly census CSV files covering all 18 inpatient units in the health system in 2025 (6,570 rows).

At 2 p.m. on New Year's Eve the 4 West charge nurse marks bed 405B *Dirty*, and environmental services needs to know
exactly when. Down in the emergency department, a registration clerk keys in an ambulance arrival while the phone rings,
and one dropped leading zero sends a chart to the wrong patient. Upstairs, the Chief Nursing Officer wants the monthly
census packet on her desk the first morning of every month. None of these jobs should depend on someone remembering a
manual step. In this lesson you'll write macros that **run by themselves** when something happens (events), a
**data-entry form** that refuses bad input before it reaches the sheet, and a **one-click report** that combines a folder
of exports, files a PDF, and keeps a backup. You'll also learn where VBA stops being the right tool, and what to use
instead.

## What you'll learn

- Respond to workbook and worksheet events (Open, Change, BeforeSave)
- Build a validated data-entry UserForm that writes to a Table
- Automate a report: combine files with Dir, export to PDF, save timestamped copies
- Know the safe limits of automation and when to use Office Scripts or Power Automate

## 📖 Guide

This lesson builds on everything in Level 5 so far: recording and macro security (5.1), the VBE and control flow (5.2),
ranges, sheets and workbooks (5.3), and error handling (5.4). VBA runs only in desktop Excel for Windows and Mac.

### 1. Set up your workbook and files

The practice workbook is an `.xlsx` file, which can't store code. You'll save it as a macro-enabled workbook first, then
bring in the starter code.

1. Make a folder on your computer for this lesson, for example `C:\ExcelCourse\Lesson5.5` (Mac: a folder in
   **Documents**). Choose a folder that OneDrive or SharePoint **doesn't** sync (section 8 explains why).
2. Download the workbook into that folder. Then download [`data/census_monthly.zip`](data/census_monthly.zip), unzip it,
   and make sure the `census_monthly` folder sits **next to** the workbook. On Windows, **Extract All** suggests a
   destination that ends in `census_monthly`, which puts the folder inside a second `census_monthly` folder. Delete that
   last part of the path so the files land in `Lesson5.5\census_monthly\`. You can also browse the
   [`data/census_monthly`](data/census_monthly) folder online. If you downloaded the whole course from GitHub
   (**Code → Download ZIP**), the folder is inside this lesson's `data` folder, so copy it out and put it next to your
   workbook.
3. Open the workbook. If a yellow **Protected View** bar appears, click **Enable Editing**. (If your saved `.xlsm` later
   shows a red *Security Risk* bar, unblock it as shown in Lesson 5.1.)
4. Choose **File → Save As**, pick **Excel Macro-Enabled Workbook (\*.xlsm)**, and save it in the same folder.
5. Press **Alt + F11** (Mac: **Option + F11**, or **Developer → Visual Basic**) to open the Visual Basic Editor (VBE).
   Choose **File → Import File…** and import [`starter/modLog.bas`](starter/modLog.bas) and
   [`starter/modReports.bas`](starter/modReports.bas).

Your folder should look like this. The macros create `Reports` and `Archive` themselves.

```
Lesson5.5/
├── 5.5-events-userforms-automation.xlsm
└── census_monthly/
    ├── census_2025-01.csv
    ├── census_2025-02.csv
    ├── …
    ├── census_2025-12.csv
    └── README.txt
```

The starter and solution files come in three kinds, and each kind goes into the VBE differently:

| File | What it holds | How it goes into the VBE |
|---|---|---|
| [`starter/modLog.bas`](starter/modLog.bas) | `WriteLog` (adds a row to the Log sheet) and `EventsOn` (repair macro). Already complete | **File → Import File…** |
| [`starter/modReports.bas`](starter/modReports.bas) | `CombineCensusFiles` (task 11), `SaveTimestampedCopy` (section 9) and `BuildCensusReport` (bonus) | **File → Import File…** |
| [`starter/ThisWorkbook.cls`](starter/ThisWorkbook.cls) | `Workbook_Open` and `Workbook_BeforeSave` (task 6) | Open it in a text editor, copy, and **paste** into the **ThisWorkbook** code window, replacing what's there |
| [`starter/BedBoard.cls`](starter/BedBoard.cls) | `Worksheet_Change` (task 7) | **Paste** into the **BedBoard** sheet's code window, replacing what's there |
| [`starter/frmIntake.vba`](starter/frmIntake.vba) | The intake form's code (tasks 8–10) | Build the form (section 7), then **paste** into its code window |

Every part you need to write is marked with a numbered comment in the starter file that says what goes there.

> ⚠️ **Paste event code. Don't import it.** Event procedures only run from the module that belongs to the object, such
> as ThisWorkbook or the BedBoard sheet. **File → Import File…** always creates a new, separate module (such as
> `Module1` or `ThisWorkbook1`), and event procedures in it never fire. Clear the code window before you paste, because the files start
> with `Option Explicit` and a module can't contain it twice.

The [`solutions/`](solutions) folder holds the finished code for every task. It's a spoiler, so try first.

The workbook's sheets:

| Sheet | What it is | Used in |
|---|---|---|
| **BedBoard** | Table `tblBeds`: 36 beds on 4 West with BedID, Room, Status, StatusTime, Notes | Tasks 1, 7 |
| **Intake** | Table `tblIntake`: 69 ED arrivals at Bluestone Memorial, Dec 25 to 11:40 on Dec 31 | Tasks 8–10 |
| **Log** | Timestamp, Event, Detail, User. `WriteLog` appends here | Tasks 6, 7 |
| **Combined** | Headers only. `CombineCensusFiles` fills it from the CSV files | Tasks 11, 12, bonus |
| **Settings** | ReportMonth, CsvFolder, ReportFolder, LastExport: cells your macros read and write | Task 11, bonus |
| **Lists** | Arrival modes and chief complaints for the form's drop-downs | Tasks 8–10 |

VBE shortcuts you'll use in this lesson:

| Action | Windows | Mac |
|---|---|---|
| Open or switch to the VBE | **Alt + F11** | **Option + F11**, or **Developer → Visual Basic** |
| Properties window | **F4** | **View → Properties Window** |
| Immediate window | **Ctrl + G** | **View → Immediate Window** |
| Run the Sub or UserForm the cursor is in | **F5** | **Run → Run Sub/UserForm** |
| Step through code one line at a time | **F8** | **Debug → Step Into** |
| Toggle a breakpoint | **F9** | **Debug → Toggle Breakpoint** |
| Show a form's code / its designer | **F7** / **Shift + F7** | **View → Code** / **View → Object** |
| Run a macro from Excel | **Alt + F8** | **Developer → Macros** |

### 2. Events: code that runs by itself

Until now, every macro you wrote ran because you started it. An **event** is something that happens to a workbook, a
sheet or a form: the file opens, a cell changes, a button is clicked. An **event procedure** is a Sub with a reserved
name that Excel runs automatically when its event happens. You never call it yourself.

Event procedures must live in the module that belongs to their object:

| Module | In the Project Explorer it looks like | Events it can handle | Examples |
|---|---|---|---|
| ThisWorkbook | `ThisWorkbook` | The workbook's events | `Workbook_Open`, `Workbook_BeforeSave`, `Workbook_SheetChange` |
| A sheet module | `Sheet3 (BedBoard)` | Events of **that sheet only** | `Worksheet_Change`, `Worksheet_SelectionChange` |
| A UserForm module | `frmIntake` | The form's and its controls' events | `UserForm_Initialize`, `cmdSave_Click` |
| A standard module | `modLog`, `Module1` | None. A Sub named `Workbook_Open` here is just an ordinary macro | `WriteLog`, `CombineCensusFiles` |

The safest way to start an event procedure is to let the VBE write its first and last lines:

1. In the Project Explorer, double-click **ThisWorkbook** (or a sheet such as **Sheet3 (BedBoard)**).
2. In the left drop-down above the code window, choose **Workbook** (or **Worksheet**).
3. In the right drop-down, choose the event, for example **BeforeSave**.

The VBE inserts the exact name and argument list. A typed signature with one argument wrong fails with
*"Procedure declaration does not match description of event or procedure having the same name."*

> 💡 **Tip:** Choosing **Workbook** in the left list also inserts an empty `Workbook_Open`. Delete it if you don't need it.

The events you'll use most:

| Event | Module | Fires when | Hospital example |
|---|---|---|---|
| `Workbook_Open()` | ThisWorkbook | The file opens with macros enabled | Log who opened the bed board, then jump to the BedBoard sheet |
| `Workbook_BeforeSave(SaveAsUI, Cancel)` | ThisWorkbook | Just before every save | Refuse to save an intake log with blank MRNs |
| `Workbook_BeforeClose(Cancel)` | ThisWorkbook | Just before the file closes | Cancel scheduled `OnTime` runs |
| `Workbook_SheetChange(Sh, Target)` | ThisWorkbook | Cells change on **any** sheet | One audit trail for the whole workbook |
| `Worksheet_Change(Target)` | Sheet | Cells on **this** sheet are changed by a person or by VBA | Stamp the time a bed status changed |
| `Worksheet_SelectionChange(Target)` | Sheet | The selection moves | Show a bed's notes in the status bar |
| `Worksheet_Activate()` | Sheet | The sheet becomes the active sheet | Refresh a summary |
| `Worksheet_Calculate()` | Sheet | The sheet recalculates | React when a formula result crosses a limit |

> ⚠️ `Worksheet_Change` fires when cell **values** are entered, pasted, cleared or written by VBA. It does **not** fire
> when a formula's result changes because of a recalculation, and it doesn't fire for formatting changes. Use
> `Worksheet_Calculate` to react to formula results.

> 📋 Events need macros enabled. In Protected View, with macros disabled, and in Excel for the web, nothing runs.

### 3. Workbook events: Open and BeforeSave

```vba
Private Sub Workbook_Open()
    ' runs after the workbook opens
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    ' runs before the save; set Cancel = True to stop it
End Sub
```

`Workbook_BeforeSave` gives you two arguments:

| Argument | Meaning |
|---|---|
| `SaveAsUI` | `True` when the **Save As** dialog is about to appear (Save As, or the first save of a new file) |
| `Cancel` | Starts as `False`. Set it to `True` and Excel cancels the save. It works because `Cancel` is passed **ByRef** (Lesson 5.4): your change goes back to Excel |

**Worked example.** The registration supervisor wants a trail of who opened the intake workbook, and she never wants it
saved with a blank MRN or ESI level. The `WriteLog` helper from `modLog` adds one row to the Log sheet:

```vba
Private Sub Workbook_Open()
    WriteLog "Open", "Opened by " & Application.UserName
    Me.Worksheets("BedBoard").Activate             ' Me = this workbook
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Dim lo As ListObject, missing As Long

    Set lo = Me.Worksheets("Intake").ListObjects("tblIntake")
    missing = Application.WorksheetFunction.CountBlank(lo.ListColumns("MRN").DataBodyRange) _
            + Application.WorksheetFunction.CountBlank(lo.ListColumns("ESILevel").DataBodyRange)

    If missing > 0 Then
        Cancel = True                                ' stop the save
        WriteLog "Save blocked", missing & " blank MRN/ESILevel cell(s)"
        MsgBox missing & " required Intake cell(s) are blank.", vbExclamation, "Save cancelled"
        Exit Sub
    End If

    WriteLog "Save", IIf(SaveAsUI, "Save As dialog", "Save")
End Sub
```

Inside the ThisWorkbook module, **`Me`** means this workbook, so `Me.Worksheets("Intake")` is the same as
`ThisWorkbook.Worksheets("Intake")`. **`IIf(condition, valueIfTrue, valueIfFalse)`** is VBA's one-line version of the
worksheet IF function, so the last line logs *Save As dialog* when `SaveAsUI` is True and *Save* otherwise. Because
BeforeSave runs *before* Excel writes the file, the Save row it logs is included in the saved file.

To test it, save (Ctrl + S, Mac: ⌘ + S), close the workbook, reopen it, and click **Enable Content** if asked. The Log
sheet now shows a Save row and an Open row.

> ⚠️ Because `Workbook_Open` writes to the Log, Excel asks whether to save when you close the file, even if you changed
> nothing else. That's expected: the workbook really did change.

> ⚠️ If the file is stored in OneDrive or SharePoint with **AutoSave** on, Excel saves for you every few seconds, and those
> saves can run `Workbook_BeforeSave` too. Keep the code fast, and avoid message boxes that would pop up over and over.

> 📋 `Workbook_Open` also runs when another macro opens the file with `Workbooks.Open`, unless events are switched off.
> Older workbooks use a Sub named `Auto_Open` in a standard module instead. It still works, but `Workbook_Open` is the
> modern choice.

### 4. Worksheet_Change, Target and Intersect

```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Target = every cell changed by this one action
End Sub
```

**`Target`** is a Range holding every cell that changed in one action. Beginners often assume it's a single cell, and
it often isn't:

| What you do on the BedBoard sheet | Target |
|---|---|
| Type *Dirty* in C11 and press Enter | `C11` (1 cell) |
| Select C20:C22, type *Clean*, press **Ctrl + Enter** (Mac: **⌘ + Return**) | `C20:C22` (3 cells, one event) |
| Paste a 5 × 2 block | the 10 pasted cells |
| Select B10:D12 and press Delete | `B10:D12` (9 cells, including ones that were already empty) |
| Your macro runs `Range("D11").Value = Now` | `D11`. VBA writes fire the event too |
| A formula elsewhere recalculates | No event |

Usually you care about one column only, such as Status. **`Intersect`** returns the cells two (or more) ranges have in
common, or **`Nothing`** if they don't overlap:

```
Intersect(range1, range2, ...)        → Range, or Nothing
Intersect(Range("B10:D12"), Range("C2:C37"))   → C10:C12   (3 cells)
Intersect(Range("E5"), Range("C2:C37"))        → Nothing
```

Picture two rectangles on the grid. Intersect keeps only the cells inside **both**, so it trims away other columns
**and** rows outside the second range. Because `Nothing` isn't a range, you test it with `Is`, not `=`:

```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim watched As Range, changed As Range, cell As Range

    Set watched = Me.ListObjects("tblBeds").ListColumns("Status").DataBodyRange
    Set changed = Intersect(Target, watched)
    If changed Is Nothing Then Exit Sub            ' the edit was somewhere else

    For Each cell In changed.Cells                 ' one pass per changed Status cell
        Debug.Print cell.Address, cell.Value
    Next cell
End Sub
```

Paste this into the BedBoard sheet's module, open the Immediate window, and experiment. Change one status, then fill
three with Ctrl + Enter, then type in a Notes cell. You'll see one line, then three lines, then nothing.

A few details make this code robust:

- **`Me`** in a sheet module is that sheet. `Me.ListObjects("tblBeds")` is the Table, and
  `.ListColumns("Status").DataBodyRange` is the Status column without its header (Lessons 3.1 and 5.3). This keeps
  working when someone adds beds or moves columns. `Me.Range("tblBeds[Status]")` is an equivalent shortcut.
- **Loop over `changed.Cells`, not over `Target`.** Then a big selection that only partly touches the Status column
  is handled correctly, cell by cell.
- **To reach the same row in another column**, work out the row's position inside the Table:
  `r = cell.Row - lo.DataBodyRange.Row + 1`, then use `lo.ListColumns("StatusTime").DataBodyRange.Cells(r)`.

> ⚠️ Avoid `If Target.Column = 3 Then`. It tests only the first cell of Target, misses multi-cell edits, and breaks as
> soon as someone inserts a column.

**Worksheet_SelectionChange** fires every time the selection moves, which is often. Keep it light. This one shows the
selected bed's notes in Excel's status bar. It looks only at the first selected cell, `Target.Cells(1)`, and uses the
same row-inside-the-Table formula as above:

```vba
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim lo As ListObject, r As Long

    Set lo = Me.ListObjects("tblBeds")
    If Intersect(Target.Cells(1), lo.DataBodyRange) Is Nothing Then
        Application.StatusBar = False                ' hand the status bar back to Excel
    Else
        r = Target.Row - lo.DataBodyRange.Row + 1    ' row inside the table
        Application.StatusBar = "Bed " & lo.ListColumns("BedID").DataBodyRange.Cells(r).Value & _
                                ": " & lo.ListColumns("Notes").DataBodyRange.Cells(r).Value
    End If
End Sub
```

### 5. Application.EnableEvents: stop the event loop

The bed board handler needs to **write** cells: a timestamp in StatusTime, and a tidied status ("dirty" becomes
"Dirty"). Writing a cell is a change, so it fires `Worksheet_Change` again. The handler then writes again, which fires
it again, and so on. Depending on the code, Excel freezes for a while, stops with an *"Out of stack space"* error, or
quietly gives up after many nested calls.

**`Application.EnableEvents`** switches Excel's workbook and worksheet events (such as `Workbook_BeforeSave` and
`Worksheet_Change`) on or off. Turn events off just before your handler writes,
and **always** turn them back on, even when an error occurs:

```vba
    ' ...inside Worksheet_Change, after the Intersect test from section 4
    ' (lo = Me.ListObjects("tblBeds"), changed = the Status cells that changed)
    On Error GoTo CleanUp
    Application.EnableEvents = False          ' our own writes won't fire events

    For Each cell In changed.Cells
        cell.Value = StrConv(Trim$(cell.Value), vbProperCase)              ' "dirty" -> "Dirty"
        r = cell.Row - lo.DataBodyRange.Row + 1
        lo.ListColumns("StatusTime").DataBodyRange.Cells(r).Value = Now
        WriteLog "Bed status", lo.ListColumns("BedID").DataBodyRange.Cells(r).Value & " -> " & cell.Value
    Next cell

CleanUp:
    Application.EnableEvents = True           ' runs on success AND after an error
    If Err.Number <> 0 Then MsgBox "Bed board update failed: " & Err.Description, vbExclamation
End Sub
```

**`StrConv(text, vbProperCase)`** capitalizes the first letter of each word, so *dirty* and *DIRTY* both become
*Dirty*. On success, the code runs on into `CleanUp:`. After an error, `On Error GoTo CleanUp` jumps there. Either way
events come back on. This is the cleanup pattern from Lesson 5.4.

What you need to know about EnableEvents:

| Fact | Why it matters |
|---|---|
| It belongs to the **Application**, not to one workbook | Switching it off affects every open workbook |
| It **stays** off after your macro ends | One crash between `False` and `True` silently disables every handler until you fix it |
| It doesn't affect UserForm events | A button's `Click` still runs while EnableEvents is False |

When your handlers suddenly stop working, open the Immediate window (Ctrl + G), type this line, and press Enter. You
can also run the `EventsOn` macro in `modLog`.

```vba
Application.EnableEvents = True
```

> ⚠️ **Event macros clear Undo.** When any macro changes the workbook, Excel empties its Undo list. With a handler that
> stamps a time on every status change, **Ctrl + Z** can't undo the user's last edit on that sheet. Tell your users, and
> keep write-back handlers to sheets where that's acceptable.

**Debugging events.** Click the handler's first line and press **F9** to set a breakpoint, then edit a cell in Excel.
The VBE stops on that line, and you can step through with **F8** and hover over `Target` to see its address. If the
code never stops there, run through this checklist:

| Symptom | Likely cause | Fix |
|---|---|---|
| Nothing happens at all | Code in a standard module or on the wrong sheet | Move it into the sheet's own module (or ThisWorkbook) |
| Worked yesterday, nothing happens today | An earlier error left EnableEvents False | `Application.EnableEvents = True` in the Immediate window |
| Multi-cell edits log only one row | The code reads `Target.Value` or `Target.Row` | Loop over `Intersect(...)` cell by cell |
| Excel freezes or shows *Out of stack space* | The handler's own writes fire it again | Wrap the writes in `EnableEvents = False` / `True` |
| *Procedure declaration does not match…* | The signature was typed by hand | Recreate it from the drop-downs |

### 6. UserForms: controls, properties and events

A **UserForm** is a dialog box you design yourself. People type into its controls, your code checks the input, and
only clean data reaches the sheet. To add one, choose **Insert → UserForm** in the VBE. The **Toolbox** appears next to
it (if it doesn't, choose **View → Toolbox**). Draw a control by clicking a Toolbox button and dragging on the form.
Then select the control and set its properties in the **Properties window** (F4).

The first property to set on every control is **(Name)**, the name your code uses. A three-letter prefix tells you what
kind of control a name refers to:

| Control | Prefix | What your code reads | Key properties |
|---|---|---|---|
| Label | `lbl` | (nothing, it only shows text) | `Caption` |
| TextBox | `txt` | `.Value`, always a **String** | `MaxLength`, `Value` |
| ComboBox | `cbo` | `.Value` (the chosen text) and `.ListIndex` (−1 when nothing is chosen) | `Style`, `RowSource`, `List` |
| OptionButton | `opt` | `.Value`, `True` for the chosen one | `Caption`, `GroupName` |
| CheckBox | `chk` | `.Value`, `True` or `False` | `Caption` |
| Frame | `fra` | (groups other controls) | `Caption` |
| CommandButton | `cmd` | Its `Click` event | `Caption`, `Default`, `Cancel` |

Some properties that matter for a data-entry form:

- **ComboBox `Style`:** `0 - fmStyleDropDownCombo` lets people type anything. `2 - fmStyleDropDownList` allows only items
  from the list, which is what you want for coded fields such as arrival mode.
- **OptionButtons in one Frame** form a group, so only one of them can be selected. (All option buttons placed directly
  on the form form a single group. To split them into separate groups without frames, give each group its own
  `GroupName`.)
- **CommandButton `Default = True`** makes Enter click it. **`Cancel = True`** makes Esc click it.
- **Tab order** is the order the cursor moves when people press Tab. Right-click the form and choose **Tab Order** to
  set it.

There are three ways to fill a ComboBox:

| Method | Example | Good for | Watch out |
|---|---|---|---|
| `RowSource` property | `Lists!A2:A5` (Properties window or code) | A quick link to a fixed range of cells | It's resolved against the **active** workbook. It doesn't grow when items are added below the range. You can't call `AddItem` while it's set |
| `AddItem` method | `cboArrivalMode.AddItem "Ambulance"` | Short lists, or items your code builds | One call per item |
| `List` property | `cboComplaint.List = rng.Value` | Loading a whole column of cells at once | `rng` must be more than one cell, because one cell's `.Value` isn't an array |

A form has a short life cycle, and each stage has its own statement or event:

| Code | What it does |
|---|---|
| `frmIntake.Show` | Loads the form (running `UserForm_Initialize`) and shows it **modally**: the user must close it before touching the sheet |
| `frmIntake.Show vbModeless` | Shows it without blocking the sheet |
| `Private Sub UserForm_Initialize()` | Runs once as the form loads, before it appears. Fill lists and set defaults here |
| `Me` | Inside the form's code, the form itself, so `Me.Controls("optESI3")` is the third ESI button |
| `Me.Hide` | Hides the form but keeps it in memory with its values |
| `Unload Me` | Closes the form and clears it from memory. Next time it starts fresh |

To create a control's event procedure, double-click the control in the designer. For example, double-clicking
`cmdSave` creates `Private Sub cmdSave_Click()`.

> 📋 **Mac:** Excel for Mac (Microsoft 365, 2021, 2019) runs UserForms and can edit them, but its form designer is more
> limited than the Windows one. If a property isn't available in the Mac Properties window, set it in code inside
> `UserForm_Initialize` instead.

### 7. Build the intake form, step by step

Here's the form you'll build. Registration clerks use it to add ED arrivals to `tblIntake` on the Intake sheet.

```
┌─ ED Quick Registration ──────────────────────────────────────┐
│  MRN         [00787672]        Arrival  [2025-12-31 11:44]   │
│  Last name   [Cruz        ]    First name [Bobby       ]     │
│  Arrival mode [Walk-In          ▼]                           │
│  ┌ ESI level ──────────────────────────┐                     │
│  │ ( )1   ( )2   ( )3   (•)4   ( )5    │   [ ] Interpreter   │
│  └─────────────────────────────────────┘       needed        │
│  Chief complaint [Vomiting / Dizziness              ▼]       │
│  Saved: King, Gerald (MRN 02718484)                          │
│                              [ Save patient ]  [ Close ]     │
└──────────────────────────────────────────────────────────────┘
```

1. In the VBE, choose **Insert → UserForm**. In the Properties window set **(Name)** = `frmIntake` and
   **Caption** = `ED Quick Registration`.
2. Draw the controls below. Add a Label next to each input so people know what to type.

   | Control | (Name) | Set these properties |
   |---|---|---|
   | TextBox | `txtMRN` | `MaxLength` = 8 |
   | TextBox | `txtLast` | |
   | TextBox | `txtFirst` | |
   | TextBox | `txtArrival` | (filled with the current time by code) |
   | ComboBox | `cboArrivalMode` | `Style` = 2 - fmStyleDropDownList |
   | Frame | `fraESI` | `Caption` = ESI level |
   | 5 OptionButtons **inside** the frame | `optESI1` … `optESI5` | `Caption` = 1 … 5 |
   | CheckBox | `chkInterpreter` | `Caption` = Interpreter needed |
   | ComboBox | `cboComplaint` | `Style` = 2 - fmStyleDropDownList |
   | Label | `lblStatus` | `Caption` = (empty) |
   | CommandButton | `cmdSave` | `Caption` = Save patient, `Default` = True |
   | CommandButton | `cmdClose` | `Caption` = Close, `Cancel` = True |

3. Right-click the form, choose **Tab Order**, and put the inputs in top-to-bottom order. If your designer has no
   **Tab Order** command (as can happen on a Mac), set each control's `TabIndex` property instead: 0 for the first
   input, 1 for the next, and so on.
4. Press **F7** (Mac: **View → Code**) to open the form's code window, delete anything in it, and paste
   [`starter/frmIntake.vba`](starter/frmIntake.vba). Finish its numbered to-do steps using the rest of this section.
5. Insert a standard module (**Insert → Module**) with a macro that opens the form, then assign it to a button on the
   Intake sheet (**Insert → Shapes**, draw a shape, right-click it → **Assign Macro…**, as in Lesson 5.1):

   ```vba
   Public Sub ShowIntakeForm()
       frmIntake.Show
   End Sub
   ```

   The same module is in [`solutions/modIntake.bas`](solutions/modIntake.bas). Add it only after the form exists,
   because code that names `frmIntake` won't compile until then. While you build the form, you can also click inside
   its code window and press **F5** (Mac: **Run → Run Sub/UserForm**) to open it.

**Filling the lists.** `UserForm_Initialize` loads the arrival modes one at a time with `AddItem` and loads the 37 chief
complaints in one step with `List`. It then calls `ClearForm`, which empties the boxes and puts the current time in
`txtArrival` with `Format(Now, "yyyy-mm-dd hh:nn")`.

```vba
For Each c In wsLists.Range("A2", wsLists.Cells(wsLists.Rows.Count, "A").End(xlUp))
    cboArrivalMode.AddItem c.Value
Next c
Set complaints = wsLists.Range("C2", wsLists.Cells(wsLists.Rows.Count, "C").End(xlUp))
cboComplaint.List = complaints.Value
```

**Validating before you write.** The Save button calls a function that returns a list of problems, or `""` when there
are none. Only then does it touch the sheet. If anything is wrong, the form stays open so the clerk can fix it.

```vba
Private Sub cmdSave_Click()
    Dim msg As String
    msg = ValidationMessage()
    If Len(msg) > 0 Then
        MsgBox "Please fix the following:" & vbLf & vbLf & msg, vbExclamation, "Intake not saved"
        Exit Sub
    End If
    AddIntakeRow
    ClearForm                                  ' ready for the next patient
    txtMRN.SetFocus
End Sub
```

Each rule is one line of `ValidationMessage`. These are the tools for writing them:

| Rule | Test | Notes |
|---|---|---|
| MRN is exactly 8 digits | `Trim$(txtMRN.Value) Like "########"` | In `Like`, `#` matches one digit, `?` any one character, `*` any run of characters (including none), `[A-Z]` one capital letter |
| Name is filled in | `Len(Trim$(txtLast.Value)) > 0` | `Trim$` stops a lone space from counting as a name |
| Arrival is a real date-time | `IsDate(txtArrival.Value)` | Ask for `yyyy-mm-dd hh:mm`, which no regional setting can misread |
| Something is chosen in a list | `cboArrivalMode.ListIndex <> -1` | −1 means no item is selected |
| One ESI button is chosen | `SelectedESI() <> 0` | See below |

`Like` examples: `"00787672" Like "########"` is True. `"787672" Like "########"` is False (too short).
`"0078767A" Like "########"` is False (A isn't a digit).

Option buttons don't share a single value, so ask each one in turn. `Me.Controls("optESI" & i)` builds the control's
name as text, which lets one loop check all five buttons:

```vba
Private Function SelectedESI() As Long
    Dim i As Long
    For i = 1 To 5
        If Me.Controls("optESI" & i).Value = True Then
            SelectedESI = i
            Exit Function
        End If
    Next i                                     ' none chosen: returns 0
End Function
```

**Writing a new row to the Table.** `ListRows.Add` adds an empty row at the bottom of the Table and returns it as a
**ListRow**. The Table grows, so its banding, formats and every formula that refers to `tblIntake` include the new
row automatically. Write each field by **column name** so the code survives someone reordering the columns:

```vba
Set lo = ThisWorkbook.Worksheets("Intake").ListObjects("tblIntake")
Set newRow = lo.ListRows.Add                                   ' new empty row at the bottom
newRow.Range.Cells(1, lo.ListColumns("ESILevel").Index).Value = SelectedESI()
```

The starter wraps that last line in a small helper, `SetField lo, newRow, "ESILevel", SelectedESI()`.

**Convert every value to the right type.** A TextBox always hands you text. When VBA writes text that *looks like* a
number or a date into a General cell, Excel converts it, exactly as if you had typed it. That's helpful for some fields
and harmful for others:

| Field | The control gives you | Write it as | Why |
|---|---|---|---|
| MRN | Text `"00787672"` | Set the cell's `NumberFormat = "@"` **first**, then write the text | In a General cell Excel stores the number 787672, and the leading zeros are gone for good |
| ArrivalDateTime | Text `"2025-12-31 11:44"` | `CDate(txtArrival.Value)` | A real date-time sorts, filters and calculates. Text dates can have day and month swapped on non-US systems |
| ESILevel | A Long from `SelectedESI()` | The number | `COUNTIFS(...,"<=2")`, AVERAGE and PivotTables need numbers |
| Interpreter | `True` / `False` | `CBool(chkInterpreter.Value)` | Shows as TRUE / FALSE and filters cleanly |

> ⚠️ Don't rely on the column's existing format. A new Table row copies the formats of the row above it, and in this
> workbook the MRN column is General. Set `"@"` on the new cell yourself (or write `"'" & mrn`).

> 💡 **Tip:** `IntakeID` is the next number in sequence:
> `Application.WorksheetFunction.Max(lo.ListColumns("IntakeID").DataBodyRange) + 1`.

### 8. Combine a folder of files with Dir

Each month the bed-management system drops a CSV export into a folder. **`Dir`** lists the files in a folder, one name
at a time:

| Call | Returns |
|---|---|
| `Dir(folder)` or `Dir(folder & "*.csv")` | **Starts a new search** and returns the first matching file name (name only, no folder), or `""` |
| `Dir()` (no arguments) | The **next** name from the same search, or `""` when there are no more |
| `Dir(path, vbDirectory)` | The folder's name if it exists, or `""`. Use it to check a folder before you use it |

Three rules follow from how Dir works:

1. **Call `Dir()` with no arguments inside the loop.** Calling `Dir` *with* an argument anywhere inside the loop starts
   a new search and breaks the one you're in. Dir can only run one search at a time.
2. **Don't rely on the order.** Windows usually returns names alphabetically, but nothing guarantees it. If order
   matters, collect the names first and sort them.
3. **On a Mac, Dir doesn't support wildcards** such as `*.csv`. List every file with `Dir(folder)` and test each name
   with `Like`. This works on Windows too, so the lesson's code uses it everywhere.

```vba
Sub ListCensusFiles()
    Dim folder As String, fileName As String

    folder = ThisWorkbook.Path & Application.PathSeparator & "census_monthly" & Application.PathSeparator
    fileName = Dir(folder)                       ' first file (any type)
    Do While Len(fileName) > 0
        If LCase$(fileName) Like "*.csv" Then Debug.Print fileName
        fileName = Dir()                         ' next file
    Loop
End Sub
```

`ThisWorkbook.Path` is the folder the workbook is saved in, and **`Application.PathSeparator`** is `\` on Windows and `/`
on a Mac. Building paths from them means the same code runs on both. Run `ListCensusFiles` and the Immediate window
lists the 12 CSV files. README.txt is skipped.

**Open, copy, close.** For each CSV, open it as a workbook, copy its data rows by value, and close it without saving:

```vba
Set wbSrc = Workbooks.Open(Filename:=folder & fileName)
Set body = wbSrc.Worksheets(1).Range("A1").CurrentRegion          ' header + data
Set body = body.Offset(1, 0).Resize(body.Rows.Count - 1)          ' data only: down 1 row, 1 row shorter
wsOut.Cells(nextRow, 1).Resize(body.Rows.Count, body.Columns.Count).Value = body.Value
nextRow = nextRow + body.Rows.Count
wbSrc.Close SaveChanges:=False
```

- A CSV opens as a workbook with one sheet. Because the dates in these files are written as `yyyy-mm-dd`, Excel reads
  them as real dates on any system.
- `.Value = .Value` copies values without the clipboard, which is faster and safer than Copy/Paste (Lesson 5.3).
- **Clear the old rows first** (row 2 down) so that running the macro twice doesn't double the data.
- `Application.ScreenUpdating = False` stops the screen flickering as each file opens. Turn it back on at the end.

> 📋 VBA opens CSV files with US settings by default: commas between fields and month/day dates. A CSV that uses
> semicolons, as many European systems do, needs `Workbooks.Open(..., Local:=True)` to use your computer's regional
> settings.

Put an error handler around the loop. If one file is damaged, the handler should close that file (so it isn't left open
in the background), restore ScreenUpdating, and report which file failed:

```vba
Fail:
    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    MsgBox "Combine failed on " & fileName & ":" & vbLf & Err.Description, vbExclamation
    Resume CleanUp
```

> ⚠️ **OneDrive and SharePoint paths.** When a workbook is opened from OneDrive or SharePoint (for example with AutoSave
> on), `ThisWorkbook.Path` can return a web address such as `https://…` instead of a folder, and Dir can't use it. Keep
> automation workbooks in a local folder that isn't synced, or check for `"http"` at the start of the path and stop with
> a clear message, as the solution does.

> ⚠️ **Mac:** Excel for Mac runs in a security sandbox. The first time your macro opens a file in a new folder, macOS may
> show a *Grant File Access* dialog. Approve the `census_monthly` folder and the macro continues.

> 📋 Power Query (Lesson 4.3) can also combine a folder of CSV files with **Data → Get Data → From File → From Folder**,
> with no code, and refresh it in one click. Choose VBA when combining is one step in a longer job, such as the bonus
> report, which combines, summarizes, exports and archives in one run.

**Summarize the combined rows by unit (for the bonus).** The bonus report turns the combined rows into one row per
unit for a single month. You already have most of the tools. Read the Combined sheet into an array with one `.Value`
read (Lesson 5.3), loop over the rows in the report month, and keep running totals per DeptID with a Collection that
maps each key to a **slot** number (the `SlotFor` idea from Lesson 5.4, section 10). Three details are new:

| Tool | Example | What it does |
|---|---|---|
| `DateSerial(y, m + 1, 0)` | `DateSerial(2025, 12 + 1, 0)` returns 12/31/2025 | Day 0 of the next month is the last day of this month. `Day()` of that date is the number of days in the month (31 for December, 28 for February 2025) |
| `ReDim Preserve` | `ReDim Preserve pd(1 To n)` | Resizes a dynamic array to `n` slots and **keeps** the values already in it. A plain `ReDim` empties the array. Declare the array with empty parentheses first (`Dim pd() As Double`) |
| `UnitIndex(idx, key)` | `k = UnitIndex(unitIdx, "D130")` | Already in `starter/modReports.bas`. Returns the slot number stored under a DeptID in the Collection, or 0 for a unit you haven't seen yet |

When `UnitIndex` returns 0, the row belongs to a new unit. Add 1 to your unit count `n`, grow each per-unit array with
`ReDim Preserve`, and store `n` under the DeptID with `unitIdx.Add Item:=n, Key:=deptID`. Then add the row's
MidnightCensus and StaffedBeds to slot `k` (or to slot `n` for a new unit).

### 9. Export to PDF and save timestamped copies

**`ExportAsFixedFormat`** saves a workbook, a sheet, a range or a chart as a PDF:

```vba
wsRep.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, Quality:=xlQualityStandard, _
    IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
```

| Argument | Typical value | Meaning |
|---|---|---|
| `Type` | `xlTypePDF` | The file type |
| `Filename` | A full path ending in `.pdf` | Where to save. An existing file with that name is replaced, but if it's open in a PDF viewer the export fails |
| `Quality` | `xlQualityStandard` | `xlQualityMinimum` makes a smaller file |
| `IncludeDocProperties` | `True` | Copies the title and author into the PDF |
| `IgnorePrintAreas` | `False` | `False` respects a print area set on the sheet |
| `From`, `To` | Page numbers | Export only some pages |
| `OpenAfterPublish` | `False` | `True` opens the PDF in your viewer, which is handy while testing |

The PDF follows the sheet's page setup, so set it before you export:

```vba
With wsRep.PageSetup
    .Orientation = xlLandscape
    .Zoom = False                 ' must be False before FitToPages* takes effect
    .FitToPagesWide = 1           ' all columns on one page width
    .FitToPagesTall = False       ' as many pages tall as needed
    .CenterFooter = "Page &P of &N"
End With
```

> 💡 **Tip:** PageSetup is slow because Excel talks to the printer driver for every property. If you set many properties,
> put `Application.PrintCommunication = False` before them and `Application.PrintCommunication = True` after.

**Timestamps in file names.** VBA's `Format` function turns a date into text using codes:

| Code | Meaning | 12/31/2025 5:07 PM gives |
|---|---|---|
| `yyyy` | 4-digit year | `2025` |
| `mm` | Month 01–12 (but **minutes** when it comes right after `h` or `hh`) | `12` |
| `mmm` / `mmmm` | Month name | `Dec` / `December` |
| `dd` | Day 01–31 | `31` |
| `hh` | Hour 00–23 (01–12 if the format also has `AM/PM`) | `17` |
| `nn` | Minutes 00–59 | `07` |
| `ss` | Seconds | `00` |

So `Format(#12/31/2025 9:30:00 AM#, "yyyy-mm-dd_hhnn")` returns `2025-12-31_0930`. Two habits keep file names
trouble-free:

- **Put the year first** (`yyyy-mm-dd`). Then sorting files by name also sorts them by date.
- **Never put `:` in a file name.** Windows doesn't allow it (or any of `\ / * ? " < > |`), and it causes trouble on a
  Mac too. That rules out `hh:mm`, so use `hhnn` or `hh-nn`.

> 📋 `nn` works only in VBA's `Format`. The worksheet `TEXT` function writes minutes as `mm` after `h`.

**Save, SaveAs or SaveCopyAs?** All three write a file, but they leave you in different places:

| Method | What happens to the workbook you're in | Use it for |
|---|---|---|
| `ThisWorkbook.Save` | Saved under its current name | Ordinary saves |
| `ThisWorkbook.SaveAs Filename:=…, FileFormat:=…` | **Renamed**: from now on you're editing the new file. The old file stays on disk as last saved | Creating a new version, or changing format (`FileFormat:=xlOpenXMLWorkbookMacroEnabled` is .xlsm) |
| `ThisWorkbook.SaveCopyAs Filename:=…` | **Unchanged**: same name, same folder, same saved/unsaved state. A copy is written to disk | Backups and archive snapshots |

`SaveCopyAs` has no `FileFormat` argument, because the copy is always the same format as the original. Give it the
same extension (`.xlsm` here). Otherwise Excel may refuse to open the copy because its format and extension don't
match.

```vba
Public Sub SaveTimestampedCopy()
    Dim sep As String, folder As String, ext As String, copyPath As String

    sep = Application.PathSeparator
    folder = ThisWorkbook.Path & sep & "Archive"
    If Len(Dir(folder, vbDirectory)) = 0 Then MkDir folder        ' MkDir creates one folder level

    ext = Mid$(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, "."))  ' ".xlsm"
    copyPath = folder & sep & "CensusReport_" & Format(Now, "yyyy-mm-dd_hhnn") & ext
    ThisWorkbook.SaveCopyAs copyPath
    WriteLog "Archive", copyPath
End Sub
```

### 10. Scheduling with Application.OnTime

`Application.OnTime` asks Excel to run a macro later:

```vba
Public nextRun As Date                       ' module level, so you can cancel it later

Sub ScheduleRefresh()
    nextRun = Now + TimeValue("00:15:00")    ' 15 minutes from now
    Application.OnTime EarliestTime:=nextRun, Procedure:="RefreshBedBoard"
End Sub

Sub RefreshBedBoard()
    ' ... do the work ...
    ScheduleRefresh                          ' schedule the next run
End Sub

Sub CancelRefresh()
    On Error Resume Next                     ' an error just means nothing was scheduled
    Application.OnTime EarliestTime:=nextRun, Procedure:="RefreshBedBoard", Schedule:=False
End Sub
```

OnTime only works while **desktop Excel is running** on that computer. If someone is typing in a cell when the time
comes, the run waits until they finish. To cancel a run, you must pass the exact time it was scheduled for, which is why
`nextRun` is stored. If the workbook has been closed when the time comes, Excel reopens it to run the macro, so call
`CancelRefresh` from `Workbook_BeforeClose`. For anything that must run while nobody is logged in, see section 12.

### 11. Safe limits: security and reliability

Automation in a hospital touches patient data and clinical workflows, so it has to be trustworthy as well as clever.

**Security**

- **Don't lower macro security to make a file work.** Files from email or the internet are blocked by the Mark of the
  Web (Lesson 5.1). Unblock a file you trust, or ask IT for a **Trusted Location** for approved automation workbooks.
- **Sign macros you distribute.** In the VBE, **Tools → Digital Signature…** attaches a code-signing certificate. Your IT
  department issues real certificates. On Windows, `SelfCert.exe` (*Digital Certificate for VBA Projects*, in the Office
  folder) makes a test certificate that only your own computer trusts. If you edit signed code, sign it again.
  Organizations can then choose **Disable VBA macros except digitally signed macros** in the Trust Center.
- **Protect patient information.** Keep identifiers such as MRNs and names out of file names and log details. Export PDFs
  only to secured folders, and don't send PHI by automated email. The bed board log in this lesson records bed IDs, not
  patients, for that reason.

**Reliability**

- Use `Option Explicit`, and validate inputs before writing (section 7).
- Restore everything you switch off (`EnableEvents`, `ScreenUpdating`, `DisplayAlerts`) in a cleanup block that also
  runs after errors.
- Log what the automation did and any errors (`WriteLog`), so you can answer "when did this last run, and did it work?"
- Keep paths and settings in cells (the **Settings** sheet), not hard-coded in the code.
- Test on a copy, and take a `SaveCopyAs` backup before anything destructive, such as deleting sheets.
- Let automation **support** clinical work, not make clinical decisions. A macro can flag an ESI 1 arrival. A clinician
  decides what happens next.

**Where VBA can't run:** Excel for the web, Excel on iPad, iPhone and Android, Protected View, and any computer with
macros disabled. A workbook stored in the cloud can still contain macros, but they only run when someone opens it in
desktop Excel.

| You need to… | Best tool |
|---|---|
| React instantly to an edit someone makes in desktop Excel | A VBA event |
| Validated data entry on desktop Excel | A UserForm, or a Table with data validation (Lesson 3.2) |
| Combine and clean the same files every month | Power Query (Lesson 4.3), or VBA when it's one step of a bigger job |
| Run on a schedule with nobody logged in, on files in OneDrive or SharePoint | Power Automate + an Office Script |
| Automate a workbook people use in Excel for the web or Teams | An Office Script |
| Collect entries from many people at once | Microsoft Forms or Power Apps feeding a Table through Power Automate |

### 12. Office Scripts and Power Automate

**Office Scripts** are Excel's modern automation language. You write and run them on the **Automate** tab of Excel for
the web, and of desktop Excel in Microsoft 365. They're written in **TypeScript** (JavaScript with types) and saved in
your OneDrive, not inside the workbook. The **Record Actions** button writes a script for you, much like the macro
recorder. Office Scripts need a Microsoft 365 work or school account.

This script counts the ESI 1–2 arrivals in `tblIntake`, the same question as practice task 10:

```typescript
function main(workbook: ExcelScript.Workbook): number {
  const table = workbook.getTable("tblIntake");
  if (!table) {
    throw new Error("Table tblIntake not found.");
  }
  const esiColumn = table.getColumnByName("ESILevel");
  if (!esiColumn) {
    throw new Error("Column ESILevel not found.");
  }
  const values = esiColumn.getRangeBetweenHeaderAndTotal().getValues();
  let urgent = 0;
  for (const row of values) {
    if (typeof row[0] === "number" && row[0] <= 2) {
      urgent++;
    }
  }
  console.log(`ESI 1-2 arrivals: ${urgent}`);
  return urgent;               // Power Automate can use this value in its next step
}
```

The ideas map directly onto VBA. `workbook` plays the role of `ThisWorkbook`, `getTable` the role of `ListObjects`, and
`getValues()` returns a 2-D array like `Range.Value`. A script can return a value, which is how it hands results to
Power Automate.

**Power Automate** runs **flows** in Microsoft's cloud, so they work with no Excel open and nobody logged in. The 06:00
census summary from practice task 5 is a three-step flow:

1. **Trigger:** *Recurrence*, every day at 06:00 in your time zone.
2. **Action:** *Excel Online (Business) → Run script*. Pick the workbook in OneDrive or SharePoint and the script.
3. **Action:** *Microsoft Teams → Post message in a chat or channel*, using the script's result.

| | VBA | Office Scripts | Power Automate |
|---|---|---|---|
| Runs in | Desktop Excel for Windows and Mac | Excel for the web and Microsoft 365 desktop Excel | Microsoft's cloud |
| Language | VBA | TypeScript | A visual flow designer |
| Started by | Events, buttons, shortcuts, OnTime | A button or the Automate tab, or a flow | Schedules, new emails, new files, forms, Teams messages and more |
| Reacts to cell edits as they happen | ✔ (`Worksheet_Change`) | ✘ | ✘ |
| Custom dialog boxes | ✔ (UserForms) | ✘ | Via Power Apps or Microsoft Forms |
| Local files and folders | ✔ (`Dir`, `Workbooks.Open`) | ✘ (only the workbook it runs on) | Cloud files through connectors |
| Needs someone's Excel to be open | ✔ | ✔ when run by hand. ✘ when a flow runs it | ✘ |

Use VBA for rich desktop tools like the ones in this lesson. Move to Office Scripts and Power Automate when the work must
run unattended, in the browser, or across Microsoft 365 services.

## 🧪 Hands-on practice

Set up your folder as described in section 1, then open the **Practice** sheet of your `.xlsm`. Type answers in the
yellow cells. The gray cells read what your macros write, and the **Check** column turns green when you're right.

<!-- BEGIN GENERATED: practice -->
Tasks 1–5 are quick concept checks you can answer from the guide. Tasks 6–12 are build tasks: save this workbook as an .xlsm first, import the starter modules, and work in order. The gray cells read what your macros write, so they stay blank until your code has run.

| # | Task | Hint |
|:-:|------|------|
| 1 | On the BedBoard sheet, the Status column of tblBeds is C2:C37. Suppose you select A30:E40 and press Delete. Worksheet_Change runs once with Target = A30:E40. How many cells are in Intersect(Target, Range("C2:C37"))? | Intersect keeps only the cells that are in both ranges: shared columns AND shared rows |
| 2 | A run-time error stopped your Worksheet_Change handler right after it set Application.EnableEvents = False, and now nothing happens when you edit a Status cell. Type the exact statement you would run in the Immediate window (Ctrl + G) to switch events back on. | It's the same property your handler switched off |
| 3 | Your archive macro runs at 5:07 PM on 12/31/2025 and calls ThisWorkbook.SaveCopyAs folder & "CensusReport_" & Format(Now, "yyyy-mm-dd_hhnn") & ".xlsm". What file name does it create? Type the name only, without the folder. | Look up each code in the Format table in Guide §9 (hh is a 24-hour clock unless you add AM/PM) |
| 4 | After ThisWorkbook.SaveCopyAs runs, which workbook are you working in? Type A, B, or C.<br>A = the original file (the copy was written to disk but isn't open)<br>B = the new copy (Excel switched to it, the way Save As does)<br>C = both files are open in Excel | Compare Save, SaveAs and SaveCopyAs in Guide §9 |
| 5 | The house supervisor wants the census summary rebuilt at 06:00 every morning and posted to a Microsoft Teams channel, even on days when nobody has Excel open. Which tool fits? Type A, B, or C.<br>A = Application.OnTime in your .xlsm<br>B = a Workbook_Open macro<br>C = an Office Script run by a scheduled Power Automate flow | Which option doesn't need desktop Excel to be running? |
| 6 | Save the workbook as .xlsm and import modLog. Paste starter/ThisWorkbook.cls into the ThisWorkbook module and finish Workbook_Open and Workbook_BeforeSave so they call WriteLog "Open", … and WriteLog "Save", …. Save, close, reopen with macros enabled, then save again. The gray cell shows TRUE when the Log sheet has at least one Open row and one Save row. | Choose Workbook in the left drop-down of the ThisWorkbook code window, then the event on the right |
| 7 | Paste starter/BedBoard.cls into the BedBoard sheet's module and finish its Worksheet_Change handler. For each changed Status cell it should stamp StatusTime with Now and call WriteLog "Bed status", … once. Delete any "Bed status" rows that your testing left on the Log sheet, then make exactly these edits:<br>• Change 4W-405B's Status to Dirty.<br>• Select the Status cells of 4W-410A, 4W-410B and 4W-411A, type Clean, and press Ctrl + Enter (Mac: ⌘ + Return) to fill all three at once.<br>• In 4W-415A's Notes cell, type Bed alarm on.<br>The gray cell counts the Log's "Bed status" rows. | Intersect Target with the Status column, then loop For Each over the result |
| 8 | Build frmIntake (Guide §6–7) and use it to add these three ED arrivals on 12/31/2025. Type each arrival as yyyy-mm-dd hh:mm, for example 2025-12-31 11:41.<br>• 11:41 · MRN 02718484 · King, Gerald · Walk-In · ESI 3 · Vomiting / Dizziness · interpreter needed<br>• 11:44 · MRN 00787672 · Cruz, Bobby · Walk-In · ESI 4 · Vomiting / Dizziness · no interpreter<br>• 11:52 · MRN 05525049 · Rice, Jennifer · Ambulance · ESI 2 · Shortness of Breath · no interpreter<br>Then try to save a fourth record with MRN 787672. Your form must refuse it. The gray cell shows how many rows tblIntake has now. | ListRows.Add, then write each field into the new row |
| 9 | How many MRNs in tblIntake are stored as 8-character text? The gray cell counts them with ISTEXT and LEN. If the count is lower than your row count in task 8, your form wrote some MRNs as numbers. Delete those rows, fix AddIntakeRow, and enter the patients again. | A TextBox gives you text, but Excel converts number-like text that lands in a General cell |
| 10 | How many arrivals in tblIntake are ESI level 1 or 2 now? The gray cell uses COUNTIFS(tblIntake[ESILevel],"<=2"), which counts only real numbers. | Loop over the five option buttons with Me.Controls("optESI" & i) |
| 11 | Finish CombineCensusFiles in modReports so it loops through every .csv file in the census_monthly folder with Dir, opens each one, copies its data rows (not its header) to the next empty row of the Combined sheet, and closes it without saving. Run it. The gray cell counts the data rows on Combined, or shows a message if column A holds text or blank rows. | Dir(folder) gives the first file and Dir() each next one. CurrentRegion.Offset(1).Resize(…) drops the header |
| 12 | Now analyze what your macro imported. In the yellow cell, write one formula that reads the Combined sheet: what was the 2025 occupancy of Bluestone Memorial's Intensive Care Unit (DeptID D130)? Use patient days (sum of MidnightCensus) ÷ bed days (sum of StaffedBeds), as a percentage to 1 decimal place. | SUMIFS ÷ SUMIFS on the Combined columns |
<!-- END GENERATED: practice -->

## ✅ Answer key

The workbook has a hidden **Answer Key** sheet (right-click any sheet tab → **Unhide…** → *Answer Key*). Its *Live
result* column shows a dash for every task, because the key can't run your macros. The finished code for every task is in
[`solutions/`](solutions), and the key below shows the parts each task needs, collapsed so you don't see them by
accident.

<!-- BEGIN GENERATED: answers -->
<details>
<summary><b>🔑 Show the answer key</b> — Try every task before opening this.</summary>

**1. Cells in Intersect(Target, Status column)**

- **Answer:** 8
- **Solution:** Draw both rectangles. Target covers columns A–E, rows 30–40. The Status column covers column C, rows 2–37. They share C30:C37.

Rows 38–40 are below the table and columns other than C aren't Status cells, so they drop out, leaving 8 cells. This is why a handler loops over Intersect(Target, …) instead of over Target: it touches only the Status cells that changed, however big the selection was.

**2. Turn events back on from the Immediate window**

- **Answer:** Application.EnableEvents = True
- **Solution:** `Application.EnableEvents = True` (type it in the Immediate window and press Enter)

EnableEvents belongs to the whole Excel application, not to one workbook or sheet. It stays False after the macro stops, for every open workbook, until something sets it back. That's why every handler that turns it off needs an error handler that always turns it back on. The EventsOn macro in modLog does the same job from Alt + F8 (Mac: Developer → Macros).

**3. Timestamped file name from Format**

- **Answer:** CensusReport_2025-12-31_1707.xlsm
- **Solution:** `CensusReport_2025-12-31_1707.xlsm`

yyyy-mm-dd gives 2025-12-31. hh gives the hour on a 24-hour clock (17) because the format has no AM/PM code, and nn gives the minutes (07). Year-month-day order makes the archive files sort by date when you sort them by name. VBA would also read mm right after hh as minutes, but nn is never ambiguous, so prefer it.

**4. Where you are after SaveCopyAs**

- **Answer:** A
- **Solution:** **A.** SaveCopyAs writes a snapshot to a new file and leaves the open workbook alone.

SaveAs renames the open workbook, so afterwards you're editing the new file and the original file is no longer open. SaveCopyAs is the right tool for backups because nothing about the open workbook changes: same name, same folder, and the copy keeps the original's file format, so give it the same extension (.xlsm).

**5. Unattended daily automation**

- **Answer:** C
- **Solution:** **C.** A scheduled (Recurrence) Power Automate flow runs an Office Script in the cloud.

OnTime and Workbook_Open both need desktop Excel running on someone's PC with the workbook open (or being opened). Neither can fire at 06:00 if that PC is off, nobody is logged on, or Excel was closed overnight. A Power Automate flow with a Recurrence trigger runs in Microsoft's cloud, calls an Office Script on the workbook stored in OneDrive or SharePoint, and can post the result to Teams. VBA can't run in that setting at all.

**6. Workbook_Open and Workbook_BeforeSave write to the Log**

- **Answer:** TRUE
- **Solution:**

```vba
Option Explicit

' Runs every time the workbook opens with macros enabled.
Private Sub Workbook_Open()
    WriteLog "Open", "Opened by " & Application.UserName
    Me.Worksheets("BedBoard").Activate
End Sub

' Runs just before every save (Ctrl+S, File > Save, Save As).
' Setting Cancel = True stops the save.
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Dim lo As ListObject
    Dim missing As Long

    ' Optional rule: refuse to save while required intake fields are blank.
    Set lo = Me.Worksheets("Intake").ListObjects("tblIntake")
    missing = Application.WorksheetFunction.CountBlank(lo.ListColumns("MRN").DataBodyRange) _
            + Application.WorksheetFunction.CountBlank(lo.ListColumns("ESILevel").DataBodyRange)

    If missing > 0 Then
        Cancel = True
        WriteLog "Save blocked", missing & " blank MRN/ESILevel cell(s) in tblIntake"
        MsgBox missing & " required Intake cell(s) are blank (MRN or ESILevel)." & vbLf & _
               "Fill them in, then save again.", vbExclamation, "Save cancelled"
        Exit Sub
    End If

    WriteLog "Save", IIf(SaveAsUI, "Save As dialog", "Save")
End Sub
```


Workbook events run only from the ThisWorkbook module, and their names and arguments must match exactly, which is why picking them from the drop-downs beats typing. BeforeSave runs before Excel writes the file, so the Save row ends up inside the saved file. Setting Cancel = True stops the save. The solution uses that to refuse saving while a required Intake field is blank, and it logs the attempt before it exits.

**7. Bed board change log (Worksheet_Change + Intersect)**

- **Answer:** 4
- **Solution:**

```vba
Option Explicit

' Runs after cells on this sheet are changed by a person or by VBA.
' Target holds EVERY cell changed by that one action (it can be many cells).
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim lo As ListObject
    Dim changed As Range, cell As Range
    Dim r As Long

    Set lo = Me.ListObjects("tblBeds")

    ' React only to the Status column; ignore edits anywhere else on the sheet.
    Set changed = Intersect(Target, lo.ListColumns("Status").DataBodyRange)
    If changed Is Nothing Then Exit Sub

    On Error GoTo CleanUp
    Application.EnableEvents = False      ' our own writes below must not fire this event again

    For Each cell In changed.Cells
        ' Tidy what was typed ("dirty" -> "Dirty"). Without EnableEvents = False,
        ' this write would fire Worksheet_Change again.
        If VarType(cell.Value) = vbString Then cell.Value = StrConv(Trim$(cell.Value), vbProperCase)

        r = cell.Row - lo.DataBodyRange.Row + 1          ' row number inside the table
        lo.ListColumns("StatusTime").DataBodyRange.Cells(r).Value = Now
        WriteLog "Bed status", lo.ListColumns("BedID").DataBodyRange.Cells(r).Value & " -> " & cell.Value
    Next cell

CleanUp:
    Application.EnableEvents = True       ' runs on success AND after an error
    If Err.Number <> 0 Then MsgBox "Bed board update failed: " & Err.Description, vbExclamation
End Sub
```


Ctrl + Enter changes three cells in one action, so the event fires once with a 3-cell Target. A handler that reads only Target.Value or Target.Row logs 1 row instead of 3. The Notes edit fires the event too, but its Intersect with the Status column is Nothing, so the handler exits. Total: 1 + 3 = 4. EnableEvents = False stops the handler's own writes (the tidied status and the StatusTime stamp) from firing it again, and the CleanUp label turns events back on even after an error.

**8. Intake form: rows in tblIntake**

- **Answer:** 72
- **Solution:**

```vba
Private Sub cmdSave_Click()
    Dim msg As String

    msg = ValidationMessage()
    If Len(msg) > 0 Then
        MsgBox "Please fix the following:" & vbLf & vbLf & msg, vbExclamation, "Intake not saved"
        Exit Sub                                   ' keep the form open so the user can fix it
    End If

    AddIntakeRow
    lblStatus.Caption = "Saved: " & Trim$(txtLast.Value) & ", " & Trim$(txtFirst.Value) & _
                        " (MRN " & Trim$(txtMRN.Value) & ")"
    ClearForm                                      ' ready for the next patient
    txtMRN.SetFocus
End Sub

Private Sub AddIntakeRow()
    Dim lo As ListObject
    Dim newRow As ListRow
    Dim nextID As Long

    Set lo = ThisWorkbook.Worksheets(TABLE_SHEET).ListObjects(TABLE_NAME)
    nextID = Application.WorksheetFunction.Max(lo.ListColumns("IntakeID").DataBodyRange) + 1

    Set newRow = lo.ListRows.Add                   ' new empty row at the bottom; the Table grows

    SetField lo, newRow, "IntakeID", nextID
    SetField lo, newRow, "ArrivalDateTime", CDate(txtArrival.Value)     ' a real date-time, not text
    newRow.Range.Cells(1, lo.ListColumns("MRN").Index).NumberFormat = "@"   ' Text format first...
    SetField lo, newRow, "MRN", Trim$(txtMRN.Value)                     ' ...so leading zeros survive
    SetField lo, newRow, "LastName", Trim$(txtLast.Value)
    SetField lo, newRow, "FirstName", Trim$(txtFirst.Value)
    SetField lo, newRow, "ArrivalMode", cboArrivalMode.Value
    SetField lo, newRow, "ESILevel", SelectedESI()                      ' a number (Long)
    SetField lo, newRow, "ChiefComplaint", cboComplaint.Value
    SetField lo, newRow, "Interpreter", CBool(chkInterpreter.Value)     ' TRUE / FALSE
    SetField lo, newRow, "EnteredBy", Application.UserName
End Sub

' Writes one value into the named column of a ListRow.
Private Sub SetField(ByVal lo As ListObject, ByVal r As ListRow, ByVal columnName As String, ByVal v As Variant)
    r.Range.Cells(1, lo.ListColumns(columnName).Index).Value = v
End Sub
```


The table started with 69 arrivals, and three valid saves make 72. If you see 73, the 6-digit MRN got through, so check your Like "########" test (# matches exactly one digit). An extra, half-empty row instead means a run-time error stopped AddIntakeRow after ListRows.Add: delete that row (right-click → Delete → Table Rows) and fix the error. ListRows.Add grows the Table itself, so formats, formulas and anything that refers to tblIntake pick up the new row automatically. The full form code is in solutions/frmIntake.vba.

**9. Intake form: MRNs kept as text**

- **Answer:** 72
- **Solution:**

```vba
newRow.Range.Cells(1, lo.ListColumns("MRN").Index).NumberFormat = "@"   ' Text format first...
SetField lo, newRow, "MRN", Trim$(txtMRN.Value)                     ' ...so leading zeros survive
```


txtMRN.Value is the text "00787672". When VBA writes number-like text into a General-format cell, Excel converts it just as if you had typed it, so the cell stores 787672 and the leading zeros are gone. Formatting the cell as Text ("@") before writing keeps it as text, and so does writing "'" & mrn. Don't rely on the column's existing format: new Table rows copy whatever format the row above has.

**10. Intake form: ESI 1–2 arrivals**

- **Answer:** 18
- **Solution:**

```vba
' Option buttons have no single "value", so ask each one in turn.
Private Function SelectedESI() As Long
    Dim i As Long
    For i = 1 To 5
        If Me.Controls("optESI" & i).Value = True Then
            SelectedESI = i
            Exit Function
        End If
    Next i
    ' none selected: the function returns 0
End Function
```


Option buttons don't share a single value, so SelectedESI asks optESI1 to optESI5 in turn and returns the number of the one that's selected (0 if none, which validation rejects). Writing that Long keeps ESILevel numeric, so COUNTIFS, AVERAGE and PivotTables all work. 17 of the original 69 arrivals were ESI 1–2, and Jennifer Rice (ESI 2) brings the total to 18.

**11. CombineCensusFiles: rows on Combined**

- **Answer:** 6,570
- **Solution:**

```vba
Public Sub CombineCensusFiles()
    Dim sep As String, folder As String, fileName As String
    Dim wsOut As Worksheet, wbSrc As Workbook, body As Range
    Dim nextRow As Long, nFiles As Long

    sep = Application.PathSeparator
    If Len(ThisWorkbook.Path) = 0 Or LCase$(Left$(ThisWorkbook.Path, 4)) = "http" Then
        MsgBox "Save this workbook in a folder on your computer first " & _
               "(a OneDrive/SharePoint web path won't work with Dir).", vbExclamation
        Exit Sub
    End If

    ' The CSV folder sits next to this workbook; its name is in Settings!B4.
    folder = ThisWorkbook.Path & sep & ThisWorkbook.Worksheets("Settings").Range("B4").Value
    If Len(Dir(folder, vbDirectory)) = 0 Then
        MsgBox "Folder not found:" & vbLf & folder, vbExclamation
        Exit Sub
    End If
    folder = folder & sep

    On Error GoTo Fail
    Application.ScreenUpdating = False
    Set wsOut = ThisWorkbook.Worksheets("Combined")
    wsOut.Range("A2:H" & wsOut.Rows.Count).ClearContents     ' keep the headers in row 1
    nextRow = 2

    fileName = Dir(folder)                       ' first file in the folder (any type)
    Do While Len(fileName) > 0
        If LCase$(fileName) Like "*.csv" Then    ' skip README.txt and anything else
            Set wbSrc = Workbooks.Open(Filename:=folder & fileName)
            Set body = wbSrc.Worksheets(1).Range("A1").CurrentRegion
            If body.Rows.Count > 1 Then
                Set body = body.Offset(1, 0).Resize(body.Rows.Count - 1)    ' drop the header row
                wsOut.Cells(nextRow, 1).Resize(body.Rows.Count, body.Columns.Count).Value = body.Value
                nextRow = nextRow + body.Rows.Count
            End If
            wbSrc.Close SaveChanges:=False
            Set wbSrc = Nothing
            nFiles = nFiles + 1
        End If
        fileName = Dir()                         ' next file: no arguments!
    Loop

    wsOut.Columns("A").NumberFormat = "mm/dd/yyyy"
    WriteLog "Combine", nFiles & " files, " & (nextRow - 2) & " rows"

CleanUp:
    Application.ScreenUpdating = True
    Exit Sub

Fail:
    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False   ' never leave a source file open
    MsgBox "Combine failed on " & fileName & ":" & vbLf & Err.Description, vbExclamation
    Resume CleanUp
End Sub
```


12 files × 18 units × the days in each month = 6,570 rows. The gray cell also inspects column A (CensusDate), where every cell should be a real date. "Text rows found" means header rows or README.txt lines were copied in: drop each file's header with Offset(1, 0).Resize(rows − 1), and skip other files with LCase\$(fileName) Like "*.csv". "Blank rows found" means each file left an empty row behind. That happens with Offset(1, 0) alone, because the shifted block keeps the header's row in its count and so ends one row below the data. Because the macro clears row 2 down first, running it twice gives the same count instead of doubling it.

**12. D130 occupancy from the combined data**

- **Answer:** 83.4%
- **Solution:** `=SUMIFS(Combined!H:H,Combined!C:C,"D130")/SUMIFS(Combined!E:E,Combined!C:C,"D130")`

Occupancy for a period is total patient days ÷ total bed days, the same total-over-total rule as in Lesson 1.4. Whole-column references such as Combined!H:H suit a sheet that a macro refills, because the formula keeps working however many rows arrive next month.

</details>
<!-- END GENERATED: answers -->

## 🏆 Bonus challenge

<!-- BEGIN GENERATED: bonus -->
The Chief Nursing Officer wants a monthly census packet without anyone copying and pasting. Write BuildCensusReport in modReports. It reads the month from Settings!B3 (ReportMonth) and then:
1. Runs CombineCensusFiles.
2. Deletes and re-creates a sheet named Report_yyyy-mm (for example Report_2025-12).
3. Puts the headers DeptID, Unit, Facility, PatientDays, BedDays, ADC, Occupancy in A3:G3 and one row per unit from row 4, sorted by Occupancy (highest first). PatientDays = sum of MidnightCensus, BedDays = sum of StaffedBeds, ADC = PatientDays ÷ days in the month, and Occupancy = PatientDays ÷ BedDays. Store full precision, and format ADC as 0.0 and Occupancy as 0.0%.
4. Adds a row under the units with Total in column A and the system-wide values in D:G.
5. Exports the report sheet to CensusReport_yyyy-mm.pdf in the Reports folder next to the workbook (the folder name is in Settings!B5), writes the PDF's full path into Settings!B6, and saves a timestamped backup with SaveCopyAs.

Run it for December 2025 (the month already in Settings!B3), then set Settings!B3 to 01/01/2025 and run it again. The gray cells read your report sheets.

Work on the **Bonus** sheet of the workbook.

- **B1.** What was the December 2025 average daily census (ADC) of Bluestone Memorial's Intensive Care Unit (D130)? The gray cell looks it up on Report_2025-12. *(Hint: ADC = patient days ÷ 31 for December)*
- **B2.** Which unit (DeptID) had the highest occupancy in December 2025? The gray cell reads the first unit row of Report_2025-12, so your sort must be right. *(Hint: Range.Sort on the Occupancy column, descending, before you add the Total row)*
- **B3.** What was the system-wide inpatient occupancy in December 2025? The gray cell reads it from the Total row of Report_2025-12. *(Hint: Total patient days ÷ total bed days, not the average of the unit percentages)*
- **B4.** After rerunning for January 2025, which DeptID tops Report_2025-01? *(Hint: Change Settings!B3, run the same macro again)*
- **B5.** What is the file name (without the folder) of the last PDF your macro exported? The gray cell pulls it from the full path in Settings!B6. *(Hint: Format(reportMonth, "yyyy-mm") inside the file name)*
<!-- END GENERATED: bonus -->

<!-- BEGIN GENERATED: bonus-answers -->
<details>
<summary><b>🔑 Show the bonus solution</b> — Give it a real try first!</summary>

**B1. December ADC for D130**

- **Answer:** 18.1
- **Solution:**

```vba
Public Sub BuildCensusReport()
    Const FIRST_ROW As Long = 4                  ' first unit row on the report
    Dim wsSet As Worksheet, wsData As Worksheet, wsRep As Worksheet
    Dim reportMonth As Date, monthEnd As Date, nDays As Long
    Dim data As Variant, r As Long, i As Long, k As Long, n As Long
    Dim unitIdx As Collection
    Dim ids() As String, unitNames() As String, facs() As String
    Dim pd() As Double, bd() As Double, totPD As Double, totBD As Double
    Dim sheetName As String, sep As String, folder As String, pdfPath As String, ext As String
    Dim lastRow As Long, totalRow As Long

    On Error GoTo Fail
    Set wsSet = ThisWorkbook.Worksheets("Settings")
    If Not IsDate(wsSet.Range("B3").Value) Then Err.Raise vbObjectError + 513, , "Settings!B3 (ReportMonth) must be a date."
    reportMonth = DateSerial(Year(wsSet.Range("B3").Value), Month(wsSet.Range("B3").Value), 1)
    monthEnd = DateSerial(Year(reportMonth), Month(reportMonth) + 1, 0)   ' day 0 of next month = last day
    nDays = Day(monthEnd)

    ' Step 1: refresh the Combined sheet from the CSV folder
    CombineCensusFiles
    Set wsData = ThisWorkbook.Worksheets("Combined")
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Err.Raise vbObjectError + 514, , "The Combined sheet is empty. Check the CSV folder."
    data = wsData.Range("A2:H" & lastRow).Value          ' one fast read into memory

    ' Step 2: total patient days and bed days per unit for the report month.
    ' unitIdx maps DeptID -> position in the arrays (a Collection works on Mac too).
    Set unitIdx = New Collection
    For r = 1 To UBound(data, 1)
        If data(r, 1) >= reportMonth And data(r, 1) <= monthEnd Then
            k = UnitIndex(unitIdx, CStr(data(r, 3)))
            If k = 0 Then                                  ' first row for this unit
                n = n + 1
                ReDim Preserve ids(1 To n): ReDim Preserve unitNames(1 To n): ReDim Preserve facs(1 To n)
                ReDim Preserve pd(1 To n): ReDim Preserve bd(1 To n)
                ids(n) = CStr(data(r, 3))
                unitNames(n) = CStr(data(r, 4))
                facs(n) = CStr(data(r, 2))
                unitIdx.Add n, ids(n)
                k = n
            End If
            pd(k) = pd(k) + data(r, 8)      ' MidnightCensus -> patient days
            bd(k) = bd(k) + data(r, 5)      ' StaffedBeds    -> bed days
        End If
    Next r
    If n = 0 Then Err.Raise vbObjectError + 515, , "No census rows for " & Format(reportMonth, "mmmm yyyy") & "."

    ' Step 3: rebuild this month's report sheet
    sheetName = "Report_" & Format(reportMonth, "yyyy-mm")
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(sheetName).Delete              ' no error if it doesn't exist yet
    On Error GoTo Fail
    Application.DisplayAlerts = True
    Set wsRep = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsRep.Name = sheetName

    With wsRep
        .Range("A1").Value = "Bluestone Health System - Inpatient Census Report - " & Format(reportMonth, "mmmm yyyy")
        .Range("A3:G3").Value = Array("DeptID", "Unit", "Facility", "PatientDays", "BedDays", "ADC", "Occupancy")
        For i = 1 To n
            .Cells(FIRST_ROW + i - 1, 1).Resize(1, 7).Value = _
                Array(ids(i), unitNames(i), facs(i), pd(i), bd(i), pd(i) / nDays, pd(i) / bd(i))
            totPD = totPD + pd(i)
            totBD = totBD + bd(i)
        Next i
        .Range("A" & FIRST_ROW).Resize(n, 7).Sort Key1:=.Range("G" & FIRST_ROW), Order1:=xlDescending, Header:=xlNo
        totalRow = FIRST_ROW + n
        .Cells(totalRow, 1).Resize(1, 7).Value = _
            Array("Total", "All units", "", totPD, totBD, totPD / nDays, totPD / totBD)

        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A3:G3").Font.Bold = True
        .Range("A" & totalRow & ":G" & totalRow).Font.Bold = True
        .Range("D4:E" & totalRow).NumberFormat = "#,##0"
        .Range("F4:F" & totalRow).NumberFormat = "0.0"
        .Range("G4:G" & totalRow).NumberFormat = "0.0%"
        .Range("A3:G" & totalRow).Columns.AutoFit
        With .PageSetup
            .Orientation = xlLandscape
            .Zoom = False                  ' required before FitToPages* takes effect
            .FitToPagesWide = 1
            .FitToPagesTall = False
            .CenterFooter = "Page &P of &N"
        End With
    End With

    ' Step 4: export the report sheet to PDF and record where it went
    sep = Application.PathSeparator
    folder = ThisWorkbook.Path & sep & wsSet.Range("B5").Value
    If Len(Dir(folder, vbDirectory)) = 0 Then MkDir folder
    pdfPath = folder & sep & "CensusReport_" & Format(reportMonth, "yyyy-mm") & ".pdf"
    wsRep.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    wsSet.Range("B6").Value = pdfPath

    ' Step 5: timestamped backup of the whole workbook (you stay in the original)
    ext = Mid$(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, "."))
    ThisWorkbook.SaveCopyAs folder & sep & "CensusReport_" & Format(reportMonth, "yyyy-mm") & _
                            "_built_" & Format(Now, "yyyy-mm-dd_hhnn") & ext

    WriteLog "Report", sheetName & " -> " & pdfPath
    wsRep.Activate
    MsgBox "Report ready:" & vbLf & pdfPath, vbInformation

CleanUp:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

Fail:
    MsgBox "Report failed: " & Err.Description, vbExclamation
    WriteLog "Report error", Err.Description
    Resume CleanUp
End Sub

' Position stored for key in the Collection, or 0 if the key isn't there yet.
Private Function UnitIndex(ByVal idx As Collection, ByVal key As String) As Long
    On Error Resume Next
    UnitIndex = idx.Item(key)              ' error 5 when the key is missing -> stays 0
    On Error GoTo 0
End Function
```


The macro reads Combined into an array once, then loops in memory. That's far faster than reading cells one by one. A Collection maps each DeptID to its position in the parallel arrays (UnitIndex returns 0 for a new unit), which works on Windows and Mac. Scripting.Dictionary (Lesson 5.4) would also work, but it's Windows-only. DateSerial(y, m + 1, 0) returns the last day of the month, so the same code handles 28-, 30- and 31-day months.

**B2. December's fullest unit**

- **Answer:** D210
- **Solution:**

```vba
.Range("A" & FIRST_ROW).Resize(n, 7).Sort Key1:=.Range("G" & FIRST_ROW), _
    Order1:=xlDescending, Header:=xlNo
```


D210 (Medical-Surgical, F02) ran at 94.5%, just ahead of D190 at 93.9%. Sort before you write the Total row. Otherwise the Total row gets sorted in with the units.

**B3. December system occupancy (Total row)**

- **Answer:** 84.2%
- **Solution:**

```vba
totPD / totBD   ' total patient days / total bed days across all units
```


9,140 patient days ÷ 10,850 bed days = 84.2%. Averaging the 18 unit percentages would give every unit equal weight, so the 8-bed Intensive Care Unit at Ashby Falls would count as much as the 36-bed 4 West.

**B4. January's fullest unit (parameterized rerun)**

- **Answer:** D110
- **Solution:**

```vba
reportMonth = DateSerial(Year(wsSet.Range("B3").Value), Month(wsSet.Range("B3").Value), 1)
sheetName = "Report_" & Format(reportMonth, "yyyy-mm")
```


January's leader is D110 (Medical-Surgical 4 West) at 94.8%. Reading the month from a Settings cell instead of hard-coding it is what turns a one-off macro into a reusable report: next month, someone changes one cell and clicks once.

**B5. PDF file name recorded in Settings!B6**

- **Answer:** CensusReport_2025-01.pdf
- **Solution:**

```vba
pdfPath = folder & sep & "CensusReport_" & Format(reportMonth, "yyyy-mm") & ".pdf"
wsRep.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, Quality:=xlQualityStandard, _
    IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
wsSet.Range("B6").Value = pdfPath
```


After the January rerun the answer is CensusReport_2025-01.pdf (the check also accepts the December name if you ran December last). Recording the output path in a cell leaves an audit trail and lets the next step, such as an email or a Power Automate flow, find the file. The gray formula turns every \ into /, then keeps the text after the last /, so it works with Windows and Mac paths.

</details>
<!-- END GENERATED: bonus-answers -->

## Key takeaways

- Event procedures run by themselves, but only from their object's module (ThisWorkbook or the sheet's own module), and
  only with macros enabled. Let the VBE's drop-downs write their signatures.
- `Target` can be many cells. Use `Intersect` with the column you care about, test for `Nothing`, and loop over the
  result.
- Wrap a handler's own writes in `Application.EnableEvents = False` / `True`, with a cleanup block that turns events
  back on even after an error. If events die, run `Application.EnableEvents = True` in the Immediate window.
- A UserForm should validate everything before it writes, then add the row with `ListRows.Add` and convert each value:
  text MRNs, real dates, numeric codes.
- Combine files with `Dir(folder)`, then `Dir()` in the loop. Filter names with `Like` so the code works on a Mac. Open
  each file, copy values, and close it without saving.
- `ExportAsFixedFormat` makes the PDF, and `SaveCopyAs` with `Format(Now, "yyyy-mm-dd_hhnn")` makes a sortable backup
  without moving you out of the original.
- VBA needs desktop Excel. For scheduled, unattended or browser-based automation, use Office Scripts with Power
  Automate.

<!-- BEGIN GENERATED: nav -->
---

⬅️ **Previous:** [5.4 VBA: Custom Functions, Dictionaries & Error Handling](../04-vba-functions-error-handling/README.md) · 🏠 [Course home](../../README.md) · **Next:** [6.1 Capstone: Hospital Performance Review](../../06-capstone/01-hospital-performance-review/README.md) ➡️
<!-- END GENERATED: nav -->
