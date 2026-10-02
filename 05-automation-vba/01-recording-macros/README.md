# Lesson 5.1 · Recording Your First Macros

> **Level:** Expert · **Time:** about 45 minutes · **Workbook:** [`5.1-recording-macros.xlsx`](5.1-recording-macros.xlsx)
> **Data:** Raw daily census exports for the Medical-Surgical unit at Cedar Ridge Medical Center (November and December 2025), plus the hospital's raw Emergency Department visit exports for the same two months (bonus).

On the first working day of every month, someone on Cedar Ridge's Medical-Surgical unit downloads last month's census from the
bed-management system and repeats the same fifteen minutes of clicks. They bold the headers, fix the dates, add a total row,
work out occupancy, and freeze the top row. Down the hall, the Emergency Department does the same with its list of high-acuity
visits. A **macro** does all of those clicks in under a second, the same way every time, so the report is never late and never
formatted differently. In this lesson you'll record your first macros, run them from a shortcut and a button, and open the code
that Excel writes for you. Reading that code is your first step into VBA.

## What you'll learn

- Show the Developer tab and set macro security safely
- Record, run, and save macros in a macro-enabled workbook (.xlsm)
- Choose between absolute and relative recording
- Run macros from buttons, shortcuts, and the Quick Access Toolbar — and read the recorded code

## 📖 Guide

### 1. What a macro is

A **macro** is a saved list of instructions that Excel can replay whenever you ask. Excel macros are written in **VBA** (Visual
Basic for Applications), a programming language built into desktop Excel. You don't have to write any VBA to make one, because
the **macro recorder** watches what you do and writes the VBA for you. You click **Record**, do the job once by hand, and click
**Stop**. From then on the macro repeats the job on demand.

| Good jobs for a recorded macro | Poor fits (for now) |
|---|---|
| The same formatting on every monthly export | A one-off analysis you'll never repeat |
| Adding a standard TOTAL row or header block | Steps that need a decision, such as "if occupancy is over 95%, flag it" |
| Applying a filter and a sort, then copying the result to a new sheet | Repeating a step for every row, sheet, or file (that needs a loop, Lesson 5.2) |
| Setting up the same print layout on every report | Anything that must run in Excel for the web or on an iPad (section 12) |

The recorder stores the **result** of each action: the cell you selected, the value you typed, the format you applied. It doesn't
record mouse movements or pauses, so you can work at a calm pace. It does record mistakes, so if you format the wrong cell and then
fix it, both steps end up in the macro.

> 💡 **Tip:** Read through the steps once before you record. A clean recording is easier to read and edit afterwards.

### 2. Show the Developer tab

The **Developer tab** holds the macro tools, and it's hidden until you switch it on.

| Platform | Steps |
|---|---|
| Windows | **File → Options → Customize Ribbon**. In the right-hand list, tick **Developer**, then click **OK**. (Faster: right-click any ribbon tab → **Customize the Ribbon…**) |
| Mac | **Excel → Settings** (**Preferences** in older versions) → **Ribbon & Toolbar**. In the *Customize the Ribbon* list, tick **Developer**, then click **Save**. |

The **Code** group on the Developer tab has the five buttons this lesson uses:

| Button | What it does | Shortcut |
|---|---|---|
| **Visual Basic** | Opens the **Visual Basic Editor (VBE)**, the window where macro code lives | Alt + F11 (Mac: Option + F11) |
| **Macros** | Lists the macros you can run, edit, or delete | Alt + F8 (Mac: Option + F8) |
| **Record Macro** | Starts the recorder. While recording, the button changes to **Stop Recording** | |
| **Use Relative References** | Switches between absolute and relative recording (section 6) | |
| **Macro Security** | Opens the macro settings (section 3) | |

On a Mac laptop whose top-row keys control brightness and volume, hold **Fn** as well: Fn + Option + F11.

> 📋 You can record without the Developer tab: **View → Macros → Record Macro**. On Windows, the status bar also has a small
> record button at the bottom left. If it's missing, right-click the status bar and tick **Macro Recording**.

### 3. Macro security: run only code you trust

A macro can do anything you can do in Excel, and VBA can reach beyond Excel too: it can delete files, send email, or download
software. Macros hidden in email attachments have long been a favorite way for attackers to break into organizations, hospitals
included. That's why Excel switches macros off in a file until you decide to trust it.

**Trust Center settings.** On Windows, open **Developer → Macro Security** (or **File → Options → Trust Center → Trust Center
Settings → Macro Settings**). You'll see four choices, plus a **Trust access to the VBA project object model** box. Leave that box
unticked, because it lets programs rewrite the code inside your workbooks.

| Setting | What happens when you open a file that contains macros | Use it? |
|---|---|---|
| Disable VBA macros without notification | Macros stay off and Excel doesn't tell you | Too strict for this course: you'd never be offered the chance to enable them |
| **Disable VBA macros with notification** (the default) | A yellow bar appears with an **Enable Content** button | ✅ Recommended |
| Disable VBA macros except digitally signed macros | Only macros signed by a publisher you trust can run | Often chosen by IT departments |
| Enable VBA macros (not recommended…) | Every macro in every file runs immediately | ❌ Never |

On a Mac the same idea lives in **Excel → Settings → Security**, with three options: disable all macros without notification,
disable all macros with notification (the default), and enable all macros. Keep the default.

> 📋 In many hospitals IT locks these settings centrally. If the options are grayed out, that's expected. Ask IT how macros are
> handled in your organization.

**Enable Content and trusted documents.** When you open your own .xlsm, Excel shows a yellow bar: *SECURITY WARNING Macros have
been disabled.* Click **Enable Content** only when you know where the file came from. On Windows, Excel then remembers the file as a
**trusted document** and stops asking. On a Mac, a dialog box asks instead, and you click **Enable Macros** each time you open the
file.

**The Mark of the Web.** Windows tags files you download from the internet or save from email attachments with a hidden marker
called the **Mark of the Web**. Excel blocks macros in those files completely and shows a red bar instead: *SECURITY RISK Microsoft
has blocked macros from running because the source of this file is untrusted.* The red bar has no Enable Content button. If you've
confirmed with the sender that the file is genuine:

1. Close the file. If it came by email, save the attachment to a folder on your computer first.
2. In File Explorer, right-click the file → **Properties**.
3. On the **General** tab, tick **Unblock** (next to *This file came from another computer…*), then click **OK**.
4. Reopen the file and click **Enable Content** on the yellow bar.

A **trusted location** is a folder whose files always open with macros enabled (**Trust Center Settings → Trusted Locations**).
Excel's startup folder, XLSTART, is one by default. Keep the list short, and never add your Downloads folder.

> ⚠️ Never enable macros in a file you weren't expecting, even if it seems to come from a colleague. Confirm with the sender by
> phone or in a new message first. A macro can open any file you can open, including files with patient information.

> 💡 **Tip:** The lesson workbook comes from GitHub, so it carries the Mark of the Web too. It contains no macros, so it only opens
> in Protected View (click **Enable Editing**). Unblock it right after you download it (steps 2–3 above) so the .xlsm you save from it
> starts clean. If your .xlsm ever shows the red bar, unblock it the same way.

**Version notes:** Mark-of-the-Web blocking arrived in Microsoft 365 for Windows in 2022 and was added to Office 2016–2021 through
updates. Older or un-updated versions show the yellow bar for downloaded files instead.

### 4. Save in a format that can hold macros

| Extension | Name in the Save As list | Holds macros? | Use it for |
|---|---|:-:|---|
| `.xlsx` | Excel Workbook | ❌ | Everyday files with no code |
| `.xlsm` | Excel Macro-Enabled Workbook | ✅ | Workbooks with macros (this lesson) |
| `.xlsb` | Excel Binary Workbook | ✅ | Very large files, which open and save faster in this format. The Personal Macro Workbook uses it too |
| `.xltm` | Excel Macro-Enabled Template | ✅ | A template that creates new macro workbooks |
| `.xlam` | Excel Add-in | ✅ | Macros and custom functions you load into every workbook (Lesson 5.4) |
| `.xls` | Excel 97–2003 Workbook | ✅ | The old format. Avoid it |

To save as .xlsm, choose **File → Save As** (F12; Mac: ⌘ + Shift + S), set *Save as type* (Mac: *File Format*) to **Excel
Macro-Enabled Workbook (*.xlsm)**, and click **Save**.

> ⚠️ **The classic trap:** you record a macro in an .xlsx file and press Ctrl + S. Excel warns that *the following features cannot
> be saved in macro-free workbooks: VB project*. If you click **Yes**, Excel saves the .xlsx and **throws your macros away**. Click
> **No**, then pick .xlsm in the Save As dialog. Save as .xlsm *before* you record, and the warning never comes up.

> 📋 Anyone you share an .xlsm with gets the same security prompts you do, and some email systems block .xlsm attachments outright.
> Share macro workbooks through a network folder or SharePoint instead.

### 5. The Record Macro dialog

Click **Developer → Record Macro**. Fill in the four fields and click **OK**. Recording starts the moment you click OK.

| Field | Example | Rules and advice |
|---|---|---|
| **Macro name** | `FormatCensusReport` | Start with a letter, then use letters, numbers, or underscores. No spaces. Avoid names that are also cell addresses (`FY2025` is column FY, row 2025), because Excel may reject them. A verb plus a noun reads well. |
| **Shortcut key** | an uppercase **R** | Optional. A lowercase letter gives Ctrl + letter. An uppercase letter gives Ctrl + Shift + letter. Mac: Option + ⌘ + letter. |
| **Store macro in** | This Workbook | **This Workbook**, **New Workbook**, or **Personal Macro Workbook** (section 10). |
| **Description** | Formats the raw daily census export and adds a TOTAL row. | Shown in the Macro dialog and as a comment at the top of the code. |

**Shortcut conflicts.** A macro's shortcut overrides Excel's own shortcut in every open workbook, for as long as the workbook
holding the macro is open. If you give a macro the shortcut Ctrl + c, pressing Ctrl + C runs your macro instead of copying. Prefer
Ctrl + Shift + letter, and avoid letters you already use with Shift (Ctrl + Shift + L toggles filters on Windows, for example). You can change
a shortcut later: **Macros** (Alt + F8; Mac: Option + F8) → select the macro → **Options…**

While you record, the Record Macro button reads **Stop Recording**, and on Windows a small square appears in the status bar. Click
either one to stop.

### 6. Absolute vs relative recording

By default the recorder uses **absolute recording**: it stores the exact address of every cell you select. Click **Use Relative
References** (it turns highlighted) and the recorder switches to **relative recording**: it stores each selection as a *move*
from the cell that was active before.

| | Absolute recording (default) | Relative recording |
|---|---|---|
| A1 is active and you click D2 | `Range("D2").Select` | `ActiveCell.Offset(1, 3).Range("A1").Select` |
| Meaning | Go to D2 | Go 1 row down and 3 columns right of the active cell |
| Replay it with B5 active | Selects D2 | Selects E6 |
| Use it for | Fixed places: the header row, column A, a title cell | "Wherever I am" steps: the row below the data, the next empty cell |

Read `ActiveCell.Offset(1, 3)` as "the cell 1 row down and 3 columns right of the active cell". Offset takes rows first, then
columns, and negative numbers move up or left. The trailing `.Range("A1")` means "the top-left cell of that spot", so it changes
nothing. For a block of cells the recorder writes something like `.Range("A1:D1")` instead: four cells across, starting there.

Four rules are worth memorizing:

1. **Only actions after you click OK are stored.** The cell that's active when recording starts isn't recorded. With relative
   recording, the moves count from that cell while you record, and from whichever cell is active when you run the macro later.
2. **You can switch mid-recording.** Click Use Relative References on and off as you go, and the recorder follows along. The
   census macro in section 7 does exactly that.
3. **The button stays on.** Use Relative References stays on for your next recording until you click it off (or restart Excel),
   so glance at it before you record.
4. **Ctrl + arrow keys are recorded the same way either way.** Ctrl + ↓ (Mac: ⌘ + ↓) becomes `Selection.End(xlDown).Select`,
   which means "jump to the last filled cell in this direction", wherever that is today. Ctrl + Shift + ↓ (Mac: ⌘ + Shift + ↓)
   becomes `Range(Selection, Selection.End(xlDown)).Select`. One jump with End plus one relative step finds the first empty row
   on any month.

**Formulas follow their own rules.** Use Relative References only affects *selections*. Every formula you type is stored in
**R1C1 notation**, which keeps your relative and absolute (`$`) references exactly as you typed them:

| R1C1 piece | Meaning | A1 equivalent, typed in E32 |
|---|---|---|
| `R2C5` | Row 2, column 5 (no brackets = fixed) | `$E$2` |
| `R2C` | Row 2, this column | `E$2` |
| `R[-1]C` | 1 row up, this column (brackets = relative) | `E31` |
| `RC[-4]` | This row, 4 columns left | `A32` |

So the formula you type decides how a TOTAL row behaves when next month is longer:

| You type in E32 (November, 30 days) | The recorder stores | Replayed in E33 (December, 31 days) |
|---|---|---|
| `=SUM(E2:E31)`, or you click AutoSum | `=SUM(R[-30]C:R[-1]C)` | `=SUM(E3:E32)`: still 30 rows, so 12/01 is missed ❌ |
| `=SUM(E$2:E31)` | `=SUM(R2C:R[-1]C)` | `=SUM(E$2:E32)`: row 2 down to the row above ✅ |

The `$` in front of the 2 anchors the first data row. It's the same mixed reference you used in Lesson 1.5, and it behaves the same
way here. The bottom of the range stays relative, so it always ends on the row just above the TOTAL row, however many days the
month has.

### 7. Walkthrough: record FormatCensusReport

Open the workbook, save it as .xlsm, and click the **Census_Nov** tab. It looks like a typical raw export: plain headers, dates
showing as serial numbers (Excel stores 11/01/2025 as 45962), squashed columns, and no totals.

| CensusDate | FacilityID | DeptID | StaffedBeds | Admissions | Discharges | MidnightCensus |
|---|---|---|---|---|---|---|
| 45962 | F03 | D310 | 30 | 8 | 5 | 28 |
| 45963 | F03 | D310 | 30 | 11 | 11 | 28 |
| 45964 | F03 | D310 | 30 | 6 | 5 | 29 |
| … 30 rows in all, ending in row 31 | | | | | | |

Your macro will turn it into a report with a TOTAL row (sums of columns D:G) and an occupancy figure in column H (patient days ÷
staffed-bed days). Check that **Use Relative References** is **off**, then follow these steps exactly.

**Start recording**

1. **Developer → Record Macro.** Macro name: **FormatCensusReport**. Shortcut key: an uppercase **R** (the dialog shows
   Ctrl + Shift + R; Mac: Option + ⌘ + Shift + R). Store macro in: **This Workbook**. Description: *Formats the raw daily census
   export and adds a TOTAL row.* Click **OK**.

**Fixed places (absolute)**

2. Select **A1:H1**: drag across it, or type `A1:H1` in the Name Box and press Enter. Press **Ctrl + B** (Mac: ⌘ + B), then
   open the **Home → Fill Color** arrow and pick the light shade in the fifth column, second row. Its ScreenTip ends in
   *Accent 1, Lighter 80%*.
3. Click **H1**, type **Occupancy**, and press **Enter**.
4. Click the **column A** heading. Press **Ctrl + 1** (Mac: ⌘ + 1) → **Number → Custom**, type `mm/dd/yyyy` in the *Type* box,
   and click **OK**.
5. Click **A1**, then press **Ctrl + ↓** (Mac: ⌘ + ↓). You land on the last date, A31.

**The row after the data (relative)**

6. Click **Use Relative References** so it's highlighted.
7. Press **↓** once. Type **TOTAL** and press **Tab**.
8. Drag across the four cells of this row under StaffedBeds, Admissions, Discharges, and MidnightCensus (**D32:G32**).
9. Type `=SUM(D$2:D31)` and press **Ctrl + Enter** (Mac: ⌘ + Return). Ctrl + Enter puts the formula in every selected cell, and
   each copy adjusts to its own column: `E$2:E31`, `F$2:F31`, `G$2:G31`.
10. Click **H32**, type `=G32/D32`, and press **Ctrl + Enter** (Mac: ⌘ + Return) so the cell stays selected. Click
    **Home → Percent Style** (Ctrl + Shift + %, the same keys on a Mac), then **Increase Decimal** once.
11. Press **Shift + Space** to select the whole row, then **Ctrl + B** (Mac: ⌘ + B) to make it bold.

**Back to fixed places (absolute)**

12. Click **Use Relative References** again to turn it **off**.
13. Click the **A** column heading, Shift + click the **H** heading, and double-click the boundary between any two selected
    headings to AutoFit the widths (Lesson 1.3).
14. Press **Ctrl + Home** (Mac: Control + Home, or Control + Fn + ←) to go back to A1. This also scrolls the sheet back to the top.
15. **View → Freeze Panes → Freeze Top Row**. As Lesson 1.1 showed, Freeze Panes works on what's on screen. Step 14 put row 1 back
    at the top, so the header row is the one that freezes.
16. **Developer → Stop Recording**.
17. Save with **Ctrl + S** (Mac: ⌘ + S).

Here is why each part is recorded the way it is:

| Steps | Mode | Recorded as | Why it works next month |
|---|---|---|---|
| 2–4 header, Occupancy label, date format | Absolute | `Range("A1:H1")`, `Range("H1")`, `Columns("A:A")` | Every export has its header in row 1 and its dates in column A |
| 5 A1, then Ctrl + ↓ | Absolute | `Range("A1").Select`, `Selection.End(xlDown).Select` | Always starts at A1, then jumps to whatever the last date is |
| 7–11 TOTAL row | Relative | `ActiveCell.Offset(1, 0).Range("A1").Select` and similar | Lands one row below the last date, whatever row that is |
| 9–10 formulas | R1C1 | `"=SUM(R2C:R[-1]C)"`, `"=RC[-1]/RC[-4]"` | The sums run from row 2 to the row above. Occupancy uses its own row |
| 13–15 widths, back to A1, frozen header | Absolute | `Columns("A:H")`, `Range("A1").Select`, `.SplitRow = 1`, `ActiveWindow.FreezePanes` | The same on every sheet. `SplitRow = 1` means "freeze the first row on screen", and `Range("A1").Select` scrolls row 1 back to the top first |

> ⚠️ **Common recording mistakes.** If Use Relative References is already on when you start (section 6, rule 3), steps 2–5 are
> stored as moves from whichever cell was active, so the header formatting and the jump to the last date land in the wrong place
> when you run the macro from another cell. If you click AutoSum in step 9 instead of typing the `$`, the total always adds exactly
> 30 rows. If you forget step 6, pressing ↓ in step 7 is stored as `Range("A32").Select`, so the macro always writes TOTAL into
> row 32.

**A safe routine for testing.** Ctrl + Z (Mac: ⌘ + Z) can't undo a macro (section 8), so protect your data before every test run:

1. Save the .xlsm right after you stop recording.
2. Run the macro on the next sheet (Census_Dec) and check the result.
3. If it's wrong, close the workbook **without saving** and reopen it. Census_Dec is back to its raw state, and your macro is still
   there because you saved it in step 1.
4. To record it again, first delete the old macro: **Macros** (Alt + F8; Mac: Option + F8) → select it → **Delete**. Census_Nov is
   already formatted, so record on a copy of the raw December export instead: right-click the Census_Dec tab → **Move or Copy…** →
   tick **Create a copy** → **OK**. December has one more day, so every row number in steps 5–10 goes up by one: the last date is
   in A32, the TOTAL row is row 33, and you type `=SUM(D$2:D32)` and `=G33/D33`. When you've stopped recording, delete the copy
   (right-click its tab → **Delete**), save, and test the new macro on Census_Dec.

> ⚠️ Don't delete or rename the Census_Nov, Census_Dec, ED_Nov, or ED_Dec sheets. The Practice and Bonus sheets read them, and a
> formula that loses its sheet shows #REF! for good.

### 8. Run a macro

| Method | How | Best for |
|---|---|---|
| **Macro dialog** | **Developer → Macros** or Alt + F8 (Mac: Option + F8), pick the macro, click **Run** | Macros you run now and then |
| **Shortcut key** | The keys you chose when recording, such as Ctrl + Shift + R | Macros you run often |
| **Shape button** | **Insert → Shapes**, draw a shape, type a label, then right-click its border → **Assign Macro…** | Workbooks other people use |
| **Form Control button** | **Developer → Insert → Button (Form Control)**, draw it, and pick the macro in the dialog that opens | The same, with a classic gray button look |
| **Quick Access Toolbar (QAT)** | **File → Options → Quick Access Toolbar** → *Choose commands from:* **Macros** → select → **Add >>** → **Modify…** to pick an icon | Macros you use every day |
| **Custom ribbon group** | **File → Options → Customize Ribbon** → **New Tab** or **New Group** → add macros to it | A team toolset |
| **VBE** | Click anywhere inside the macro's code, then press **F5** (Windows) or choose **Run → Run Sub/UserForm** (both platforms) | Testing while you edit |

Some details matter in practice:

- **A macro runs on the active sheet** unless its code names a sheet. Click the right tab before you run it.
- **Shapes and Form Control buttons** work on Windows and Mac. Avoid **ActiveX** controls, which are Windows-only and less reliable.
  To select a shape button without running it, Ctrl + click it (Mac: ⌘ + click) or right-click it.
- **The QAT** can be hidden in recent Microsoft 365 versions. If you don't see it, right-click the ribbon → **Show Quick Access
  Toolbar**. A QAT
  button remembers which workbook holds the macro and opens that file if it isn't open. In Excel for Mac, look under
  **Excel → Settings → Ribbon & Toolbar → Quick Access Toolbar**. If *Macros* isn't offered there in your version, use a sheet button.
- **Macros you can't see** in the Macro dialog are usually in a workbook that isn't open, or the *Macros in:* box is set to a
  different workbook. Set it to **All Open Workbooks**.

> ⚠️ **No undo.** Ctrl + Z (Mac: ⌘ + Z) can't reverse what a macro changed, and running a macro also clears Excel's undo history.
> Save before you run a macro you haven't tested, so you can close without saving if it goes wrong.

### 9. Read the recorded code

Open the VBE with **Alt + F11** (Mac: **Option + F11**) or **Developer → Visual Basic**. On the left, the **Project Explorer**
(**View → Project Explorer** if it's hidden; Windows: Ctrl + R) lists every open workbook as a *VBAProject*. Expand yours, then **Modules**, and double-click
**Module1**. A **module** is a container for code, and the recorder puts your macros in a standard module named Module1. After
you reopen the file, new recordings may go into Module2.

```vba
Sub FormatCensusReport()                 ' the macro starts: Sub + its name + ()
'
' FormatCensusReport Macro                ' lines that start with an apostrophe are comments
' Formats the raw daily census export and adds a TOTAL row.
'
' Keyboard Shortcut: Ctrl+Shift+R
'
    Range("A1:H1").Select                 ' one instruction per line, run top to bottom
    Selection.Font.Bold = True
    ...
End Sub                                   ' the macro ends
```

The `' Keyboard Shortcut:` line is only a note. The real shortcut is stored out of sight, so editing the comment changes nothing.
Use **Macros → Options…** to change it.

Here is what the lines you'll meet most often mean:

| Recorded line | Plain English |
|---|---|
| `Range("A1:H1").Select` | Select A1:H1 |
| `Selection.Font.Bold = True` | Make the selected cells bold |
| `With Selection.Interior` … `End With` | Set several properties of one thing (here, the selection's fill) without repeating its name on each line |
| `ActiveCell.FormulaR1C1 = "TOTAL"` | Type TOTAL into the active cell (the recorder uses FormulaR1C1 even for plain text) |
| `Columns("A:A").Select` | Select column A |
| `Selection.NumberFormat = "mm/dd/yyyy"` | Apply a number format |
| `Selection.End(xlDown).Select` | Ctrl + ↓: jump to the last filled cell |
| `Range(Selection, Selection.End(xlToRight)).Select` | Ctrl + Shift + →: extend the selection to the last filled cell on the right |
| `ActiveCell.Offset(1, 0).Range("A1").Select` | Move one row down (relative recording) |
| `Sheets("ED_Nov").Select` | Click the ED_Nov tab |
| `Sheets.Add After:=ActiveSheet` | Insert a new sheet after the current one |
| `ActiveSheet.Paste` and `Application.CutCopyMode = False` | Paste, then clear the moving border around the copied cells |
| `ActiveSheet.Range("$A$1:$J$81").AutoFilter Field:=5, Criteria1:="<=2"` | Filter the 5th column of A1:J81 to values ≤ 2 |
| `Selection.AutoFilter` | Data → Filter: turn the filter arrows on, or off if they're already on |
| `….Sort.SortFields.Add2 Key:=Range("E2:E21")` … `.Apply` | Data → Sort with the levels you chose in the dialog |
| `ActiveWindow.FreezePanes = True` | Freeze panes |

**What the recorder over-records.**

- **Select, then act.** Nearly every action becomes two lines: select something, then change `Selection`. The pair
  `Range("A1:H1").Select` / `Selection.Font.Bold = True` does the same as one line, `Range("A1:H1").Font.Bold = True`, and skipping
  the Select is faster and doesn't move the user's cursor.
- **Every setting in a dialog.** Change one option on the Alignment tab of Format Cells and the recorder writes every property on
  that tab (horizontal and vertical alignment, wrap text, orientation, indent, and more). The fill-color block above sets
  `.Pattern`, `.PatternColorIndex`, and `.PatternTintAndShade` to their defaults. Only `.ThemeColor` and `.TintAndShade`, the color
  you actually picked, matter.
- **Scrolling.** Scrolling while you record adds lines such as `ActiveWindow.SmallScroll Down:=12`. They do nothing useful.
- **Today's sizes.** Addresses like `$A$1:$J$81` or `A1:J21` describe this month's data, not next month's.
- **Temporary names.** Renaming a new sheet is recorded as `Sheets("Sheet1").Name = "…"`, but the next new sheet may be called
  Sheet2.

**Safe edits you can make today.**

1. Change a value in quotes, such as the label `"TOTAL"` or the format `"mm/dd/yyyy"`.
2. Delete scroll lines and Select lines that nothing uses.
3. Replace a hard-coded block with `Range("A1").CurrentRegion`, which means "the block of data around A1". Like pressing Ctrl + A
   (Mac: ⌘ + A) inside the data, it grows and shrinks with the export.
4. Replace a temporary sheet name with `ActiveSheet` (the sheet that's selected right now).
5. Shrink a recorded sort key to one cell. In `SortFields.Add2 Key:=Range("E2:E21")` the key only tells Excel which column to sort
   by, so `Key:=Range("E1")` does the same job however many rows there are. The rows that get sorted come from the `.SetRange`
   line, so give that line `Range("A1").CurrentRegion`.
6. Change a sheet name everywhere at once with **Edit → Replace** in the VBE (Windows: Ctrl + H).

After editing, choose **Debug → Compile VBAProject** to catch typos, save, and test on a saved copy (section 7). To get back to
Excel, press Alt + F11 again (Mac: Option + F11) or click the Excel window.

**When a macro stops with an error.** If a line can't run, Excel stops the macro and shows a **run-time error** message with a
number. Two you'll meet often with recorded code are *Run-time error '9': Subscript out of range*, which usually means the code
names a sheet that doesn't exist, and *Run-time error '1004'*, which means Excel refused the action, for example because another
sheet already has the name you're giving a sheet. Every line before the failing one has already run. To recover:

1. Click **Debug**. The VBE opens with the failing line highlighted in yellow, and the macro is paused on it. Excel can't run any
   macro while one is paused.
2. Click **Reset** (the square ■ button on the VBE toolbar, or **Run → Reset**) to stop the macro. Clicking **End** in the error
   message stops it too, without showing you the line.
3. Clean up the half-finished run. It may have left a new sheet with a temporary name, or a filter switched on. Fix those by
   hand, or close the workbook without saving and reopen your saved copy.
4. Fix the line, save, and run the macro again.

> 📋 Lesson 5.2 tours the VBE properly: its windows, stepping through code one line at a time with F8, and the Immediate window.

### 10. The Personal Macro Workbook

Set *Store macro in* to **Personal Macro Workbook** and Excel saves the macro in a hidden workbook called **PERSONAL.XLSB**,
creating it the first time. Excel opens every file in its startup folder, **XLSTART**, each time it starts, so PERSONAL.XLSB opens
hidden in the background and its macros work in every workbook you open. In the Macro dialog they appear as
`PERSONAL.XLSB!MacroName`.

| | Windows | Mac |
|---|---|---|
| File | `PERSONAL.XLSB` | `Personal Macro Workbook` |
| Folder | `C:\Users\<you>\AppData\Roaming\Microsoft\Excel\XLSTART` (paste `%APPDATA%\Microsoft\Excel\XLSTART` into File Explorer's address bar) | `~/Library/Group Containers/UBF8T346G9.Office/User Content/Startup/Excel` |

- When you quit Excel after recording into it, Excel asks whether to save changes to the Personal Macro Workbook. Click **Save**,
  or the macro is lost.
- To edit or delete one of its macros from the Macro dialog, unhide it first (**View → Unhide → PERSONAL.XLSB**), then hide it
  again when you're done. You can also edit it directly in the VBE under *VBAProject (PERSONAL.XLSB)*.

| Store macro in | Where the macro works | Good for |
|---|---|---|
| **This Workbook** | Whenever this file is open, for anyone who opens it | Report macros tied to one file, like FormatCensusReport |
| **Personal Macro Workbook** | In every workbook, but only on your computer | Your own utilities, such as "bold and freeze any header row" |
| **New Workbook** | In a new file Excel creates for it | Rarely needed |

> ⚠️ Your Personal Macro Workbook doesn't travel. A button in a shared report that points to `PERSONAL.XLSB` fails on a colleague's
> computer with *Cannot run the macro…*. Keep shared report macros in the .xlsm itself.

### 11. What the recorder can't do

The recorder writes one straight line of steps. It can't:

- make decisions, such as "if occupancy is over 95%, color the row red",
- repeat steps for each row, sheet, or file,
- remember values, ask the user a question, or recover from an error,
- adapt to new data except through the tricks in this lesson (Ctrl + arrow keys, relative references, `$` anchors, CurrentRegion).

A few actions record incompletely or not at all, including some chart and PivotTable formatting. All of these are jobs for VBA you
write yourself, which starts in Lesson 5.2. The recorder stays useful even then: when you don't know the VBA for something, record
yourself doing it once and read the line it writes.

### 12. Excel for the web, iPad, and Office Scripts

VBA macros run only in desktop Excel for Windows and Mac. Excel for the web, iPad, iPhone, and Android can open an .xlsm and edit
its cells, but they can't run or edit its macros. Microsoft's cloud alternative is **Office Scripts**, written in TypeScript.
Office Scripts need a Microsoft 365 work or school account. With one, **Automate → Record Actions** in Excel for the web (and in
current desktop Excel for Windows and Mac) records an Office Script much as the macro recorder records VBA:

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Bold the header row of the active sheet
  const sheet = workbook.getActiveWorksheet();
  sheet.getRange("A1:H1").getFormat().getFont().setBold(true);
}
```

| | VBA macros | Office Scripts |
|---|---|---|
| Runs in | Desktop Excel for Windows and Mac | Excel for the web and current desktop Excel, with a Microsoft 365 work or school account |
| Language | VBA | TypeScript |
| Stored | Inside the .xlsm (or PERSONAL.XLSB) | In OneDrive or SharePoint, separate from the workbook |
| Can reach | Files, folders, and other Office apps on your computer | Only the workbook, plus Power Automate flows |

Lesson 5.5 returns to Office Scripts and Power Automate.

**Version notes:** VBA and the recorder are in every desktop version of Excel for Windows and in Excel 2011 or later for Mac
(Excel 2008 for Mac had no VBA). Microsoft 365 records sorts with `SortFields.Add2`, and older versions record
`SortFields.Add`. Both work in current Excel, but `Add2` doesn't exist in older versions, so use `.Add` in code that must run
everywhere.

## 🧪 Hands-on practice

Download [`5.1-recording-macros.xlsx`](5.1-recording-macros.xlsx), unblock it (section 3), and save it as an .xlsm before you start.
Type each answer in the yellow cell on the **Practice** sheet. The **Check** column turns green when you're right, and the gray
cells check the TOTAL rows your macro writes.

<!-- BEGIN GENERATED: practice -->
Tasks 1–6, 9 and 13 are quick checks: type your answer in the yellow cell. Tasks 7, 8, 11 and 12 have gray cells that read the TOTAL rows your macro writes on Census_Nov and Census_Dec, so you don't type anything there. Task 10 has nothing to check: compare your buttons with the answer key. Save this file as an .xlsm before you record anything (task 1).

| # | Task | Hint |
|:-:|------|------|
| 1 | This workbook is an .xlsx file, which can't store macros. Which file extension does the Excel Macro-Enabled Workbook format use, the format you save it in before you record? Type the extension, like .xlsx | File → Save As shows each format's extension in the Save as type list |
| 2 | A colleague emails you Census_Report.xlsm. When you open it, a red bar says Microsoft has blocked macros because the source of this file is untrusted. Which is the safe way to run its macros? Type the letter.<br>A: Change the Trust Center to Enable VBA macros.<br>B: Confirm with the colleague that they sent it, save it to your computer, then right-click the file → Properties → tick Unblock.<br>C: Click Enable Content on the red bar.<br>D: Rename the file to .xlsx. | The red bar has no Enable button. Look at the file's Properties |
| 3 | Which function key do you press with Alt (Mac: Option) to open the Visual Basic Editor? Type just the key, like F5. | It's also on the Developer tab: Visual Basic |
| 4 | On Windows, what is the file name (with extension) of the Personal Macro Workbook, the hidden workbook that makes a macro available in every workbook you open? | It's a binary workbook stored in the XLSTART folder |
| 5 | With Use Relative References OFF, you click Record Macro, click cell B3, type Verified, press Enter, and stop recording. Later you select H10 and run the macro. Which cell receives Verified? Type the address, like A1. | Absolute recording stores the address you clicked |
| 6 | This time A1 is selected when you click Record Macro. You turn ON Use Relative References, click B3, type Verified, press Enter, and stop. Later you select H10 and run the macro. Which cell receives Verified? | Count how far you moved from A1 to B3, then make the same move from H10 |
| 7 | Record the FormatCensusReport macro on the Census_Nov sheet by following the recipe in Guide section 7 (TOTAL row recorded with relative references, shortcut Ctrl + Shift + R). The gray cell finds the row your macro labeled TOTAL and reads its MidnightCensus cell, which should equal November's patient days. | Guide section 7 has the click-by-click recipe |
| 8 | Step 10 of the recipe in Guide section 7 put the month's occupancy in column H of your TOTAL row on Census_Nov: total patient days divided by total staffed-bed days, formatted as a percentage with 1 decimal place. The gray cell reads that cell. | Patient days are in column G, staffed-bed days in column D |
| 9 | In the Record Macro dialog you typed an uppercase R in the Shortcut key box. Which key combination runs the macro on Windows? Type it like Ctrl+Alt+X. | An uppercase letter adds a key |
| 10 | Add two more ways to run FormatCensusReport: (a) a button on the Census_Dec sheet labeled Build report, and (b) a button on the Quick Access Toolbar. Don't click them yet. | Right-click a shape → Assign Macro |
| 11 | Go to Census_Dec and run your macro (Ctrl + Shift + R or your button). December has 31 days. The gray cell finds your TOTAL row and counts the daily rows above it. It should show all 31 days. | If it shows 30, your TOTAL row overwrote a day |
| 12 | On Census_Dec, the gray cell reads the MidnightCensus total in your TOTAL row. It should equal December's patient days (all 31 days). | Did you anchor the first row with a $ when you typed the SUM? |
| 13 | Read this recorded line: ActiveCell.FormulaR1C1 = "=SUM(R[-31]C[-1]:R[-1]C[-1])". If it runs while H33 is the active cell, which range does the SUM add up? Type it like A1:A9. | Square brackets count from the formula's own cell: R for rows, C for columns |
<!-- END GENERATED: practice -->

## ✅ Answer key

The workbook has a hidden **Answer Key** sheet (right-click any sheet tab → **Unhide…** → *Answer Key*). Reference macros are in
the [`solutions/`](solutions/) folder. [`modCensusReport.bas`](solutions/modCensusReport.bas) holds FormatCensusReport exactly as the
recorder writes it plus a tidied version, and [`modHighAcuity.bas`](solutions/modHighAcuity.bas) holds the bonus macro as recorded
on ED_Nov plus the edited version. They're spoilers, so record your own first. To try one, open the VBE and choose
**File → Import File…**. An imported macro has no shortcut key, so set one with **Macros → Options…** if you want it. The answers
are also below, collapsed so you don't see them by accident.

<!-- BEGIN GENERATED: answers -->
<details>
<summary><b>🔑 Show the answer key</b> — Try every task before opening this.</summary>

**1. This workbook is an .xlsx file, which can't store macros. Which file extension does…**

- **Answer:** .xlsm
- **Solution:** **.xlsm**: File → Save As (F12; Mac: ⌘ + Shift + S) → *Save as type:* **Excel Macro-Enabled Workbook (*.xlsm)**.

An .xlsx file can't contain VBA code. If you save a workbook that has macros as .xlsx, Excel warns you that the *VB project* can't be saved, and if you click **Yes** it throws the code away. .xlsb, .xltm, .xlam and the old .xls can hold macros too, but .xlsm is the everyday format for a workbook with macros. Save as .xlsm now, before you record, so nothing gets lost.

**2. A colleague emails you Census_Report.xlsm. When you open it, a red bar says Microsoft…**

- **Answer:** B
- **Solution:** **B.** Confirm the sender, save the file, then File Explorer → right-click → **Properties** → **General** → tick **Unblock** → **OK**. Reopen it and click **Enable Content** on the yellow bar.

Email attachments and downloads carry the **Mark of the Web**, so Excel blocks their macros outright and the red bar has no Enable Content button (C is impossible). Enabling every macro (A) would let any file run code, and renaming to .xlsx (D) doesn't run anything because .xlsx can't hold macros. Unblocking one file you've verified is the targeted, safe fix.

**3. Which function key do you press with Alt (Mac: Option) to open the Visual Basic…**

- **Answer:** F11
- **Solution:** **F11**: Alt + F11 on Windows, Option + F11 on a Mac (add Fn if your top-row keys control brightness and volume).

Alt + F11 opens the Visual Basic Editor (VBE), where the recorder puts your code. Pressing it again flips back to Excel. Alt + F8 (Mac: Option + F8) opens the Macro dialog instead.

**4. On Windows, what is the file name (with extension) of the Personal Macro Workbook, the…**

- **Answer:** PERSONAL.XLSB
- **Solution:** **PERSONAL.XLSB**, stored in `%APPDATA%\Microsoft\Excel\XLSTART`.

Choosing *Store macro in: Personal Macro Workbook* creates PERSONAL.XLSB the first time. Excel opens every file in XLSTART automatically and hides this one, so its macros are always available on your computer (but not on anyone else's).

**5. With Use Relative References OFF, you click Record Macro, click cell B3, type…**

- **Answer:** B3
- **Solution:** The recorder wrote `Range("B3").Select`, so the macro always goes to **B3**.

Absolute recording (the default) stores exact addresses: Range("B3").Select, then ActiveCell.FormulaR1C1 = "Verified", then Range("B4").Select for the Enter key. The cell you had selected before running (H10) doesn't matter.

**6. This time A1 is selected when you click Record Macro. You turn ON Use Relative…**

- **Answer:** I12
- **Solution:** The recorder wrote `ActiveCell.Offset(2, 1).Range("A1").Select`: 2 rows down and 1 column right of the active cell. From H10 that is **I12**.

Relative recording stores moves, not addresses. From A1 to B3 is 2 rows down and 1 column right, so the macro makes that same move from wherever it starts. From H10 that lands on I12. The Enter key is stored as another move, ActiveCell.Offset(1, 0).

**7. Record FormatCensusReport on Census_Nov (TOTAL patient days)**

- **Answer:** 809
- **Solution:**

```vba
Sub FormatCensusReport()
'
' FormatCensusReport Macro
' Formats the raw daily census export and adds a TOTAL row.
'
' Keyboard Shortcut: Ctrl+Shift+R
'
    ' --- Use Relative References is OFF: these lines name exact cells ---
    Range("A1:H1").Select
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "Occupancy"
    Range("H2").Select
    Columns("A:A").Select
    Selection.NumberFormat = "mm/dd/yyyy"
    Range("A1").Select
    Selection.End(xlDown).Select
    ' --- Use Relative References turned ON: moves are stored as offsets ---
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "TOTAL"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.Offset(0, 2).Range("A1:D1").Select
    Selection.FormulaR1C1 = "=SUM(R2C:R[-1]C)"
    ActiveCell.Offset(0, 4).Range("A1").Select
    Selection.FormulaR1C1 = "=RC[-1]/RC[-4]"
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.0%"
    ActiveCell.Rows("1:1").EntireRow.Select
    Selection.Font.Bold = True
    ' --- Use Relative References turned OFF again ---
    Columns("A:H").Select
    Columns("A:H").EntireColumn.AutoFit
    Range("A1").Select
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
End Sub
```


This is what the recorder writes for the recipe. Your code may differ in small ways, such as an extra `Offset` line for each Tab you pressed. The lines that matter are `Range("A1").Select` and `Selection.End(xlDown).Select` (recorded with relative references off, so they always start at A1 and jump to the last date), then the relative `ActiveCell.Offset(1, 0)` that steps into the empty row below. November has 30 days in rows 2–31, so TOTAL lands in row 32. The second `Range("A1").Select`, near the end, is the Ctrl + Home step. It scrolls back to the top, so `.SplitRow = 1` freezes the header row and not whichever row happens to be at the top of the screen.

**8. Census_Nov TOTAL row occupancy**

- **Answer:** 89.9%
- **Solution:** In H32 type `=G32/D32`, then Percent Style and Increase Decimal once. The recorder stores it as `"=RC[-1]/RC[-4]"`: same row, 1 and 4 columns to the left.

Total patient days ÷ total staffed-bed days is the unit's occupancy for the month. The formula only refers to cells in its own row, so in R1C1 notation it is the same on every sheet and works on December too.

**9. In the Record Macro dialog you typed an uppercase R in the Shortcut key box. Which key…**

- **Answer:** Ctrl+Shift+R
- **Solution:** **Ctrl + Shift + R** (Mac: Option + ⌘ + Shift + R).

A lowercase letter gives Ctrl + letter, which replaces Excel's own shortcut while this workbook is open (Ctrl + r would stop doing Fill Right). An uppercase letter adds Shift, which collides with far fewer built-in shortcuts. Change it later in Alt + F8 → select the macro → Options…

**10. Button on the sheet and on the Quick Access Toolbar**

- **Solution:**

**Sheet button**

1. On Census_Dec, choose **Insert → Shapes → Rectangle: Rounded Corners** and draw it to the right of the data (for example over columns J:K).
2. Type **Build report** while the shape is selected.
3. Right-click the shape's border → **Assign Macro…** → pick **FormatCensusReport** → **OK**.
4. Click a cell to deselect. The pointer becomes a hand over the shape, and one click runs the macro.

**Quick Access Toolbar** (Windows)

1. **File → Options → Quick Access Toolbar**.
2. *Choose commands from:* **Macros** → select **FormatCensusReport** → **Add >>**.
3. Click **Modify…**, pick an icon, set the display name to *Build census report* → **OK** → **OK**.

If you can't see the toolbar, right-click the ribbon → **Show Quick Access Toolbar**.


A shape button belongs to one sheet and travels with the workbook. A Quick Access Toolbar button belongs to your copy of Excel and remembers which workbook holds the macro, so it opens that workbook if needed. To edit a shape later without running the macro, Ctrl + click it (Mac: ⌘ + click) or right-click it.

**11. Run on Census_Dec (days above the TOTAL row)**

- **Answer:** 31
- **Solution:** Click the **Census_Dec** tab and press **Ctrl + Shift + R**. `Selection.End(xlDown)` finds the last date (row 32) and `ActiveCell.Offset(1, 0)` steps to row 33.

If you recorded the TOTAL steps with relative references off, the code says Range("A32").Select, so on December it types TOTAL over 12/31/2025 in row 32 and the gray cell shows 30. Ctrl + Down plus a relative one-row move finds the first empty row on any month. Macros can't be undone, so close without saving (or reopen your saved copy), fix the recording, and run it again.

**12. Census_Dec TOTAL patient days**

- **Answer:** 863
- **Solution:** The recorded formula `=SUM(R2C:R[-1]C)` becomes `=SUM(G$2:G32)` in row 33.

The recorder stores formulas in R1C1 notation. With =SUM(D$2:D31) the $ makes the top row absolute (R2C) and the bottom row relative (R[-1]C, the row above), so the range grows with the month. With AutoSum or =SUM(D2:D31) the recorder stores =SUM(R[-30]C:R[-1]C), which always adds exactly 30 rows. On December that skips 12/01 and the total comes out low.

**13. Read this recorded line: ActiveCell.FormulaR1C1 = "=SUM(R[-31]C[-1]:R[-1]C[-1])". If…**

- **Answer:** G2:G32
- **Solution:** **G2:G32**: R[-31] is 31 rows above row 33 (row 2), R[-1] is the row above (row 32), and C[-1] is one column left of H, which is G.

It's the R1C1 form of =SUM(G2:G32) typed in H33. Square brackets mean *relative to the formula's own cell*. A number without brackets is a fixed row or column, so R2C means row 2 of this column. Every part of this formula is relative, so in H34 it would add G3:G33. Reading R1C1 like this tells you whether a recorded formula will still be right when the data has a different number of rows.

</details>
<!-- END GENERATED: answers -->

## 🏆 Bonus challenge

<!-- BEGIN GENERATED: bonus -->
Cedar Ridge's ED medical director reviews every high-acuity visit (ESI 1 or 2) each month. The tracking system exports the month's visits sorted by PatientID, and someone spends 20 minutes turning that export into a list. Automate it. The macro must:

1. Filter the export's ESILevel column to values less than or equal to 2.
2. Copy the visible rows, with the header, to a new sheet named HighAcuity.
3. Sort that sheet by ESILevel (smallest first), then ArrivalDateTime (oldest first), and AutoFit the columns.
4. Go back to the export and turn its filter off.

Recording tips: Check that Use Relative References is off. Click A1 and turn the filter on with Data → Filter, then use the ESILevel filter arrow → Number Filters → Less Than Or Equal To → 2 → OK. Select the data with Ctrl + Shift + → and then Ctrl + Shift + ↓ (Mac: ⌘ + Shift + arrows), and copy it with Ctrl + C (Mac: ⌘ + C). Add a sheet with the + (New sheet) button next to the sheet tabs, paste with Ctrl + V (Mac: ⌘ + V), and rename the sheet by double-clicking its tab. In Data → Sort (Lesson 1.6), sort by ESILevel (Smallest to Largest), then click Add Level and pick ArrivalDateTime (Oldest to Newest). AutoFit every column: click the Select All button (the triangle where the row and column headings meet), then double-click any column boundary. Finish by clicking the ED_Nov tab, then A1, then Data → Filter again.

Record it on ED_Nov (80 visits) as ExtractHighAcuity (shortcut Ctrl + Shift + H). Then delete the HighAcuity sheet it made (right-click its tab → Delete), open the VBE, and edit the code so it works on ED_Dec (93 visits). Guide section 9 lists the edits recorded code usually needs and what to do if the macro stops with a run-time error. Run it on ED_Dec. The gray cells read the HighAcuity sheet your macro builds from December, so they stay blank until it exists. While the sheet from your November recording is still there, they show ✘ Not yet or stay blank.

Work on the **Bonus** sheet of the workbook.

- **B1.** How many visits (data rows, not counting the header) are on the HighAcuity sheet built from ED_Dec? *(Hint: COUNTIF on ED_Dec's ESILevel column gives you a target to compare with)*
- **B2.** Which EDVisitID is in the first data row (A2) of HighAcuity? *(Hint: It should be the earliest ESI 1 arrival in December)*
- **B3.** Which EDVisitID is in cell A5 of HighAcuity? *(Hint: This cell depends on both sort levels)*
- **B4.** Which EDVisitID is in the last data row of HighAcuity? *(Hint: Rows below row 21 are only sorted if the sort range adapts)*
- **B5.** After your macro finishes, how many rows of ED_Dec are visible? The gray cell counts them with SUBTOTAL(103, …), which skips rows a filter hides. It should be every December visit. It stays blank until HighAcuity holds December visits. *(Hint: Which sheet does the end of your recorded code go back to?)*
<!-- END GENERATED: bonus -->

<!-- BEGIN GENERATED: bonus-answers -->
<details>
<summary><b>🔑 Show the bonus solution</b> — Give it a real try first!</summary>

**B1. Visits on HighAcuity (from ED_Dec)**

- **Answer:** 26
- **Solution:**

```vba
Sub ExtractHighAcuity()
'
' ExtractHighAcuity Macro
' Copies ESI 1-2 visits to a new HighAcuity sheet, sorted by ESI level and arrival.
'
' Keyboard Shortcut: Ctrl+Shift+H
'
'   Recorded on ED_Nov, then edited to run on ED_Dec. Each EDIT comment marks a change.
'   Delete any old HighAcuity sheet before you run it: a second sheet with that
'   name is not allowed (run-time error 1004).
    Sheets("ED_Dec").Select                             ' EDIT: added, start on December
    Range("A1").Select
    Selection.AutoFilter
    Range("A1").CurrentRegion.AutoFilter Field:=5, Criteria1:="<=2", _
        Operator:=xlAnd                                 ' EDIT: was ActiveSheet.Range("$A$1:$J$81")
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveSheet.Name = "HighAcuity"                     ' EDIT: was Sheets("Sheet1").Select / .Name
    ActiveWorkbook.Worksheets("HighAcuity").Sort.SortFields.Clear
    ' EDIT: keys were Range("E2:E21") and Range("C2:C21"). One cell is enough
    ' to name the key column. (.Add works in Excel 2007 and later; Excel 365 records .Add2.)
    ActiveWorkbook.Worksheets("HighAcuity").Sort.SortFields.Add Key:=Range("E1"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("HighAcuity").Sort.SortFields.Add Key:=Range("C1"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("HighAcuity").Sort
        .SetRange Range("A1").CurrentRegion             ' EDIT: was Range("A1:J21")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Sheets("ED_Dec").Select                             ' EDIT: was Sheets("ED_Nov")
    Range("A1").Select
    Selection.AutoFilter
End Sub
```


Recording on ED_Nov bakes November into the code in four places. Edit each one:

1. **The filter range.** `ActiveSheet.Range("$A$1:$J$81").AutoFilter` names November's exact block, but December runs to row 94. Change it to `Range("A1").CurrentRegion`, the whole block around A1, which grows and shrinks with the export. The unedited line may still work here, but only because the line above it already switched the filter on for the whole block. Had the filter been on before you started recording, the recorder would have skipped `Selection.AutoFilter`, and on December the hard-coded line would filter rows 1–81 only.
2. **The rename.** `Sheets("Sheet1").Select` and `Sheets("Sheet1").Name = "HighAcuity"` use the new sheet's temporary name. The next new sheet may be Sheet2, and then the macro stops with run-time error 9 (Subscript out of range). The new sheet is active right after it's added, so use `ActiveSheet.Name = "HighAcuity"`.
3. **The sort ranges.** `Key:=Range("E2:E21")`, `Key:=Range("C2:C21")` and `.SetRange Range("A1:J21")` fit November's 20 visits. Use `Range("E1")`, `Range("C1")` and `Range("A1").CurrentRegion`. One cell is enough to name a sort key's column.
4. **The last sheet.** `Sheets("ED_Nov").Select` sends the macro back to November, where `Selection.AutoFilter` switches a filter on, and ED_Dec stays filtered. Change it to `Sheets("ED_Dec")`.

Adding `Sheets("ED_Dec").Select` as the first line makes the macro start on the right sheet whichever sheet is active. Both versions, as recorded and as edited, are in `solutions/modHighAcuity.bas`. Lesson 5.3 replaces the remaining Select lines and hard-coded sheet names with variables.

**B2. First EDVisitID on HighAcuity**

- **Answer:** ED211706
- **Solution:** **ED211706**, the earliest-arriving ESI 1 visit (12/01/2025 12:18).

Sorting by ESILevel first puts the 3 ESI 1 visits on top. The second sort level, ArrivalDateTime, orders visits within each level. If you see a different ID, the sort range probably stopped at row 21 (November's size), so only part of the list was sorted. In that case A2 shows ED211749. The live formula in the key finds the right visit with MINIFS: the earliest arrival among ESI 1 visits.

**B3. EDVisitID in A5 of HighAcuity**

- **Answer:** ED211715
- **Solution:** **ED211715**, the earliest-arriving ESI 2 visit (12/01/2025 20:18).

The 3 ESI 1 visits fill rows 2–4, so row 5 holds the first ESI 2 visit, which is the earliest ESI 2 arrival. This cell shows whether both sort levels worked. Sorted by ArrivalDateTime alone, the first and last rows happen to be the same as the correct list, but A5 shows ED211742. With the sort range stuck at row 21, it shows ED211739.

**B4. Last EDVisitID on HighAcuity**

- **Answer:** ED212209
- **Solution:** **ED212209**, the latest-arriving ESI 2 visit (12/26/2025 00:03).

December has 26 high-acuity visits, so the list ends in row 27. A sort range recorded as A1:J21 leaves the last 6 rows in export (PatientID) order, so the last row shows ED211905 instead. `Range("A1").CurrentRegion` always covers every row.

**B5. ED_Dec filter turned off at the end**

- **Answer:** 93
- **Solution:** End the macro with `Sheets("ED_Dec").Select`, `Range("A1").Select`, `Selection.AutoFilter`. That toggles the filter off, so all 93 visits show again.

The recording ends with Sheets("ED_Nov").Select, so on December it toggles a filter ON on ED_Nov and leaves ED_Dec showing only 26 rows. A good macro leaves the source data the way it found it. Selection.AutoFilter is a toggle (like Data → Filter), so it turns the filter off only when one is on.

</details>
<!-- END GENERATED: bonus-answers -->

## Key takeaways

- A macro replays recorded steps as VBA code. Save the workbook as **.xlsm** (or .xlsb) to keep it, because saving as .xlsx
  discards the code.
- Keep macro security at **Disable with notification**. Enable content only in files you trust, and unblock a downloaded file only
  after you've confirmed where it came from.
- **Absolute recording** stores addresses and **relative recording** stores moves. Use absolute steps for fixed places, and
  Ctrl + arrow plus a relative step for "the row after the data".
- Formulas are recorded in R1C1 notation and keep your `$` signs, so anchor the first row (`D$2`) to make totals grow with the data.
- Run macros with Alt + F8 (Mac: Option + F8), a Ctrl + Shift shortcut, a button, or the Quick Access Toolbar. Ctrl + Z can't undo
  them, so save first.
- Read the recorded code before you trust a macro. Strip the Select pairs and scroll lines, and replace hard-coded ranges and
  sheet names. If an edited macro stops with a run-time error, click Debug to see the line, then Reset, clean up, and fix it.

<!-- BEGIN GENERATED: nav -->
---

⬅️ **Previous:** [4.6 Building Interactive Dashboards](../../04-advanced-analysis/06-dashboards/README.md) · 🏠 [Course home](../../README.md) · **Next:** [5.2 VBA Fundamentals](../02-vba-fundamentals/README.md) ➡️
<!-- END GENERATED: nav -->
