# Lesson 1.1 · The Excel Interface & Navigation

> **Level:** Beginner · **Time:** about 70 minutes · **Workbook:** [`1.1-excel-interface-navigation.xlsx`](1.1-excel-interface-navigation.xlsx)
> **Data:** 500 registered patients from the Bluestone Health System patient index (every 8th record), plus a small payer list that starts out hidden. Column definitions are in the [data dictionary](../../data/README.md#patientscsv).

In your first week in any hospital job, someone will send you a spreadsheet with a quick question. *How many patients on this
list have no primary care provider? Who is the oldest? How many prefer Spanish, so we can schedule interpreters?* The list
might have 500 rows or 50,000. If you scroll and count, you'll spend ten minutes and probably miscount. If you know how to move
around Excel, select exactly the cells you need, and read the status bar, you'll have each answer in seconds without writing
a single formula. This lesson teaches those moves on a patient registration list from the fictional Bluestone Health System.

## What you'll learn

- Name the parts of the Excel window: ribbon, Quick Access Toolbar, Name Box, formula bar, grid, sheet tabs, status bar
- Understand workbooks, worksheets, cells, ranges, and cell addresses
- Move around large datasets fast with keyboard shortcuts and Go To
- Read quick statistics from the status bar
- Freeze panes, zoom, and hide/unhide rows, columns, and sheets

## 📖 Guide

This guide describes Excel for Microsoft 365, which looks the same as Excel 2021 and 2024 for everything in this lesson.
Excel 2016 and 2019 differ only where a version note says so. Shortcuts are given for Windows first, then Mac. In the Mac
shortcuts, ⌘ is the Command key, and the Control and Option keys are written as words: **Control + G** means hold Control
and press G. The rest of the course writes Mac shortcuts the same way.

### 1. A tour of the Excel window

Start by getting the workbook onto your computer:

1. Click the workbook link at the top of this page. GitHub opens a preview page rather than the file itself.
2. Click the download button (**Download raw file**) at the top right of the preview, then open the file in Excel.
3. If a yellow **Protected View** bar appears, click **Enable Editing**. Excel opens downloaded files read-only until you do,
   so you could look around but you couldn't type your answers.

Now click the **Patients** tab at the bottom of the window. Then click in the white box just above column A (the Name Box),
type `N2:N501`, and press **Enter** to select every patient's height. You'll see something like this:

```
+----------------------------------------------------------------------------+
| [Save] [Undo] [Redo]    1.1-excel-interface-navigation.xlsx    [Search]    |  <- Quick Access Toolbar, title bar
+----------------------------------------------------------------------------+
| File  Home  Insert  Page Layout  Formulas  Data  Review  View  Help        |  <- ribbon tabs
| [Clipboard] [Font] [Alignment] [Number] [Styles] [Cells] [Editing]         |  <- groups on the Home tab
+-----------+------+---------------------------------------------------------+
| N2      v |  fx  | 66.9                                                    |  <- Name Box, formula bar
+---+-------------+------------+------------+------------+-------------------+
|   |  A          |  B         |  C         |  D         |  ...              |  <- column headers
+---+-------------+------------+------------+------------+-------------------+
| 1 |  PatientID  |  MRN       |  LastName  |  FirstName |  ...              |
| 2 |  PT10008    |  04700089  |  Clark     |  Arthur    |  ...              |  <- the grid of cells
| 3 |  PT10016    |  07426502  |  Kennedy   |  Danielle  |  ...              |
+---+-------------+------------+------------+------------+-------------------+
  ^ row headers (the empty box above row 1 is the Select All button)
+----------------------------------------------------------------------------+
| <  >   Start Here | Practice | Patients | Bonus |  (+)                     |  <- sheet tabs
+----------------------------------------------------------------------------+
| Ready    Average: 63.21   Count: 500   Sum: 31605      [views]  -o-  100%  |  <- status bar, zoom slider
+----------------------------------------------------------------------------+
```

| Part | Where it is | What it does |
|---|---|---|
| **Ribbon** | Across the top | Holds the commands, organized into **tabs** (Home, Insert, Data, View…) and, inside each tab, **groups** (Font, Alignment, Editing…). In this course, a path like **View → Freeze Panes** means "click the View tab, then the Freeze Panes button." |
| **File tab** | Left end of the ribbon | Opens the Backstage view: Save As, Open, Print, Export, and Options. |
| **Quick Access Toolbar** (QAT) | Above or below the ribbon | A small strip of favorite commands (Save, Undo, Redo) that stays visible whichever tab is open. Click its drop-down arrow to add more. |
| **Search box** | Title bar | Type what you want to do, such as "freeze", and Excel finds the command. Windows: **Alt + Q** jumps there. |
| **Name Box** | Left of the formula bar | Shows the address of the active cell. Type an address here and press Enter to jump to it (section 5). |
| **Formula bar** | Right of the Name Box | Shows what's really stored in the active cell. A cell can *display* something shorter or rounded, but the formula bar shows the true content. |
| **Column headers** and **row headers** | Letters across the top, numbers down the left | Name each column and row. Click one to select that whole column or row. |
| **Grid** and **active cell** | The middle | The cells themselves. The active cell has a thick green border, and whatever you type goes into it. |
| **Select All button** | The small triangle where the row and column headers meet | Selects every cell on the sheet. |
| **Sheet tabs** | Bottom left | One tab per visible worksheet. Click a tab to switch sheets, and click **(+)** to add one. Right-click a tab for Rename, Hide, Unhide, and more. |
| **Status bar** | Very bottom | Shows the mode (Ready, Enter, Edit), quick statistics for the selected cells (section 6), the view buttons, and the zoom slider. |

> 📋 **Mac:** Excel for Mac has the same ribbon tabs, plus the macOS menu bar at the top of the screen (Excel, File, Edit, View,
> Insert, Format, Tools, Data, Window, Help). Many commands live in both places. The Quick Access Toolbar sits in the window's
> title bar, and settings are under **Excel → Settings** (called Preferences in older versions) instead of **File → Options**.

> 📋 **Version note:** In recent Microsoft 365 updates for Windows, the Quick Access Toolbar can be hidden by default. If you
> don't see it, right-click the ribbon and choose **Show Quick Access Toolbar**.

> 💡 **Tip:** Need more room for data? Press **Ctrl + F1** (Mac: **⌘ + Option + R**) to collapse the ribbon to just its tab
> names, and press it again to bring it back. On Windows you can also press and release **Alt** to show **KeyTips**, the small
> letters on every ribbon command. Typing them in order runs the command without the mouse: **Alt, W, F, F** is
> **View → Freeze Panes → Freeze Panes**. Excel for Mac doesn't have KeyTips.

### 2. Workbooks, worksheets, cells, and ranges

These five terms come up in every lesson.

- A **workbook** is an Excel file, such as `1.1-excel-interface-navigation.xlsx`.
- A **worksheet** (or just **sheet**) is one grid inside the workbook, shown as one tab. This workbook has Start Here,
  Practice, Patients, Bonus, and a few sheets you can't see yet.
- A **cell** is one box in the grid, where a column and a row cross.
- A **cell address** names a cell by its column letter and then its row number. **C347** is column C, row 347. The letter
  always comes first. The **active cell** is the cell that's selected right now, and its address is always in the Name Box.
- A **range** is a rectangle of cells, written as its top-left address, a colon, and its bottom-right address. **O2:O501**
  means every cell from O2 down to O501.

| You write | It means | Cells |
|---|---|--:|
| `C347` | One cell | 1 |
| `O2:O501` | Column O from row 2 to row 501 (every patient's weight) | 500 |
| `B2:D4` | A block 3 columns wide and 3 rows tall | 9 |
| `A1:D1` | Row 1 from column A to column D | 4 |
| `O:O` | All of column O | 1,048,576 |
| `2:2` | All of row 2 | 16,384 |
| `A1:A5,C1:C5` | Two separate ranges (the comma means "and") | 10 |

**How big is a worksheet?** Every worksheet has **1,048,576 rows** and **16,384 columns**. The column letters run A to Z, then
AA to AZ, BA to BZ, and so on, up to **XFD**. That's more than 17 billion cells, so a 500-patient list fills a tiny corner.

> 📋 **Version note:** These limits apply to `.xlsx` files in Excel 2007 and later. An old `.xls` file opens in
> **Compatibility Mode** (the title bar says so) and is limited to 65,536 rows and 256 columns (A to IV). Save it as `.xlsx`
> with **File → Save As** to get the full grid.

> 💡 **Tip:** If your column headers show numbers (1, 2, 3…) instead of letters, the **R1C1 reference style** is turned on.
> Turn it off in **File → Options → Formulas** by clearing **R1C1 reference style** (Mac: **Excel → Settings → Calculation**,
> or **General** in older versions).

The Patients data is formatted as an **Excel Table** named `tblPatients`, which is why it has banded rows and filter
arrows in the header row. Tables change how a few selection shortcuts behave (section 3). You'll learn Tables properly in
Lesson 3.1.

**The Patients sheet at a glance.** The practice tasks refer to columns by letter, so keep this map handy. Row 1 holds the
headers and rows 2 to 501 hold one patient each.

| Column | Header | What it holds |
|:-:|---|---|
| A | PatientID | The patient's ID in this dataset, such as PT10008. Every row has one. |
| B | MRN | The **medical record number (MRN)**, the 8-digit number on the patient's chart. It's stored as text so its leading zeros stay. |
| C, D | LastName, FirstName | The patient's name |
| E | Sex | F or M |
| F | DOB | Date of birth |
| G, H | City, ZIP | Where the patient lives |
| I | Phone | Phone number. This column starts out hidden (section 9). |
| J | Email | Email address, blank when none is on file |
| K | PreferredLanguage | The language the patient prefers for their care, such as English or Spanish |
| L | PrimaryPayerID | A code for the patient's main **payer**, the insurer or program that pays for their care. The hidden Payers sheet lists the payer names (section 10). |
| M | PCPProviderID | A code for the patient's **primary care provider (PCP)**, blank when no PCP is on file |
| N, O | HeightIn, WeightLb | Height in inches and weight in pounds |
| P | ChronicConditions | Long-term conditions as codes separated by semicolons, such as `HTN;DM` (hypertension and diabetes). Blank when none are recorded. |
| Q | RegistrationDate | The date the patient was first registered |

### 3. Selecting cells and ranges

Almost everything in Excel starts with a selection: you select cells, then act on them. Here are the ways to select.

| To select | Mouse | Windows keys | Mac keys |
|---|---|---|---|
| One cell | Click it | Arrow keys | Arrow keys |
| A range | Drag across it | **Shift + arrow** | **Shift + arrow** |
| A range from here to a clicked cell | **Shift + click** the far corner | | |
| Extra, separate cells or ranges | **Ctrl + click** or **Ctrl + drag** (Mac: **⌘ + click** or **⌘ + drag**) | | |
| From here to the edge of the data | | **Ctrl + Shift + arrow** | **⌘ + Shift + arrow** |
| The whole column | Click the column letter | **Ctrl + Space** | **Control + Space** |
| The whole row | Click the row number | **Shift + Space** | **Shift + Space** |
| The data block around the active cell, then the whole sheet | Select All button (whole sheet) | **Ctrl + A** (press twice for the whole sheet) | **⌘ + A** |

While you drag, the Name Box shows the size of the selection, such as **500R x 1C** (500 rows by 1 column). It's a quick way
to confirm you grabbed the whole column and nothing more.

> 📋 **Inside a Table** these shortcuts work in steps. In `tblPatients`, Ctrl + Space first selects just that column's data
> (rows 2 to 501), a second press adds the header, and a third selects the entire worksheet column. Ctrl + A works the same
> way for the whole Table: data, then data plus headers, then the whole sheet.

> ⚠️ **Mac:** macOS reserves several **Control + arrow** combinations for Mission Control and switching desktops, so in
> Excel for Mac use **⌘ + arrow** to jump. Also, if you have more than one keyboard language installed, **Control + Space**
> may switch input languages instead of selecting a column. Click the column letter instead, or change the shortcut in
> **System Settings → Keyboard → Keyboard Shortcuts → Input Sources**.

### 4. Moving around fast

Scrolling through 500 rows with the mouse wheel is slow and error-prone. These shortcuts take you anywhere in a keystroke.

| Move to | Windows | Mac |
|---|---|---|
| The edge of the current block of data | **Ctrl + arrow** | **⌘ + arrow** |
| The edge of the data, selecting along the way | **Ctrl + Shift + arrow** | **⌘ + Shift + arrow** |
| The start of the row (column A) | **Home** | **Home**, or **Fn + ←** |
| Cell A1 | **Ctrl + Home** | **Control + Home**, or **Control + Fn + ←** |
| The last used cell on the sheet | **Ctrl + End** | **Control + End**, or **Control + Fn + →** |
| One screen down / up | **Page Down** / **Page Up** | **Fn + ↓** / **Fn + ↑** |
| One screen right / left | **Alt + Page Down** / **Alt + Page Up** | **Fn + Option + ↓** / **Fn + Option + ↑** |
| The next / previous sheet | **Ctrl + Page Down** / **Ctrl + Page Up** | **Option + →** / **Option + ←** |
| Back to the active cell after scrolling away | **Ctrl + Backspace** | **Control + Delete** |
| Any address you type | **F5** or **Ctrl + G** (Go To) | **Control + G** |
| Any text or number you type | **Ctrl + F** (Find) | **Control + F** |

> 📋 **Mac:** Wherever this lesson says **Ctrl + arrow** or **Ctrl + Shift + arrow**, press **⌘** instead of Ctrl. Other
> shortcuts, such as Ctrl + Home and Ctrl + F, use the **Control** key on a Mac, as the table shows.

**How Ctrl + arrow decides where to stop.** Excel looks at the cell you're on and the next cell in the direction of the
arrow, then follows three rules:

1. If both cells are filled, it moves to the **last filled cell** before a blank.
2. If the next cell is blank (or you start on a blank cell), it jumps over the blanks to the **next filled cell**.
3. If there is no filled cell left in that direction, it goes to the **edge of the worksheet** (row 1,048,576 or column XFD).

Rule 1 is what makes Ctrl + ↓ so useful. In column A, every patient has a PatientID, so pressing Ctrl + ↓ from A1 lands on
the last patient in row 501. Gaps are the trap, because Ctrl + arrow stops at every one. Some patients have no email
address, so the Email column (J) has gaps. From J1, Ctrl + ↓ lands on J3, because J2 is blank and J3 is the next filled cell (rule 2).
Press it again and you stop at J5, then J8. Each gap interrupts the jump. In a column whose first gap is further down,
rule 1 stops you on the last filled cell above the gap, and that cell is easy to mistake for the bottom of the list.

> ⚠️ **Blanks stop Ctrl + arrow early.** Before you trust Ctrl + ↓ (or Ctrl + Shift + ↓) to find the bottom of a column, check
> the row number in the Name Box. To find the true last entry of a column with gaps, start *below* the data and press
> Ctrl + ↑. Excel then stops on the last filled cell. Better still, use a column with no blanks, such as an ID column.

**Ctrl + End** goes to the **last used cell**: the cell where the last used row and the last used column meet. It's a fast
way to see how big a dataset is.

> ⚠️ Ctrl + End can overshoot. After you delete rows or clear formatting at the bottom of a sheet, Excel may still remember
> the old last cell until you save the workbook.

> 💡 **Tip:** Lost after scrolling? Press **Ctrl + Backspace** (Mac: **Control + Delete**) to scroll the active cell back into
> view without moving it.

### 5. The Name Box and Go To

The **Name Box** does more than show addresses. Click inside it, type, and press **Enter**:

| Type in the Name Box | What happens |
|---|---|
| `C347` | Jumps to cell C347 |
| `O2:O501` | Selects the 500 weights, however far they stretch off screen |
| `A600` | Jumps to an empty cell below the data |
| `tblPatients` | Selects the data in the Table named tblPatients |
| `N2:N501,O2:O501` | Selects two ranges at once |

**Go To** does the same job through a dialog. Press **F5** or **Ctrl + G** (Mac: **Control + G**), type an address or range in
the **Reference** box, and press **Enter**. Go To also remembers where you came from: open it again and the Reference box
already holds your previous location, so **F5, Enter** takes you straight back.

The **Special…** button in the Go To dialog opens **Go To Special**, which selects cells by type rather than by address:

| Go To Special option | Selects |
|---|---|
| Blanks | Every empty cell in the selection (a quick way to see the gaps in a column) |
| Constants | Every cell holding a typed value rather than a formula |
| Last cell | The last used cell, the same as Ctrl + End |
| Visible cells only | Only the cells you can see, skipping hidden rows and columns |

### 6. Quick statistics on the status bar

Select two or more cells and the **status bar** summarizes them instantly. Nothing is written into the sheet, so it's the
fastest way to answer "how many" or "how much" before you decide whether a formula is worth writing.

| Statistic | What it reports | On by default? |
|---|---|:-:|
| **Average** | The mean of the numbers in the selection | ✓ |
| **Count** | How many cells are **not empty**: numbers, text, and dates all count | ✓ |
| **Numerical Count** | How many cells hold **numbers** (dates count, because Excel stores dates as numbers) | |
| **Minimum** | The smallest number or the earliest date | |
| **Maximum** | The largest number or the latest date | |
| **Sum** | The total of the numbers | ✓ |

To show or hide a statistic, **right-click the status bar** (Mac: Control + click, or click with two fingers) and tick it in
the **Customize Status Bar** menu. Your choices stick for every workbook you open afterward.

**Worked example: heights.** Type `N2:N501` in the Name Box and press Enter. The status bar reads
**Average: 63.21**, **Count: 500**, **Sum: 31605**. Turn on Minimum and Maximum and you'll see **Minimum: 19** and
**Maximum: 76.9** (inches). A 19-inch patient looks like a typo until you check the DOB column: the list includes babies, and
they pull the average height down. Always read the extremes before you trust an average.

**Worked example: Count vs Numerical Count.** Turn on Numerical Count in the status bar, then select the MRNs, `B2:B501`.
Count shows **500**, but Average and Sum disappear, and Numerical Count doesn't count a single MRN. The MRNs look like
numbers, but they're stored as text on purpose so their leading zeros survive. (The small green triangles in the corner of each MRN cell are Excel pointing this out.
Lesson 1.2 explains them.) Add the heights to the selection by typing `B2:B501,N2:N501` in the Name Box: Count jumps to
**1000**, while Numerical Count shows **500**, because only the heights are numbers.

> ⚠️ **Count is not Numerical Count.** Count includes text, so on a column of names or IDs it tells you how many cells are
> filled in. Numerical Count ignores text, which also means it ignores numbers that are stored as text.

> 💡 **Tip:** The status bar uses the cells' number format. Select dates and Minimum shows a date. If you ever see a plain
> number where you expected a date, that's the date's **serial number**: Excel counts days with 1/1/1900 as day 1
> (Lesson 2.3).

> 💡 **Tip:** In Excel for Microsoft 365 (version 2206 or later) and Excel 2024 on Windows, click a number on the status bar
> to copy it, then paste it into any cell.

### 7. Find and Find All

**Find** searches the sheet for text or numbers. Press **Ctrl + F** (Mac: **Control + F**), type what you're looking for in
**Find what**, and choose a button:

- **Find Next** selects the next matching cell. After you close the dialog, **Shift + F4** (Mac: **⌘ + G**) finds the next
  one again.
- **Find All** lists every match in the bottom of the dialog and counts them: *"41 cell(s) found."* Click any row in the list
  to jump to that cell. On Windows, press **Ctrl + A** inside the list to select every found cell at once, and then the
  status bar can summarize them.

> ⚠️ **Find searches only what you've selected.** If more than one cell is selected when you press Ctrl + F, Excel searches
> only inside that selection. That's handy when you mean it, but it's a classic reason for a puzzling "We couldn't find what
> you were looking for" message. Click a single cell first (A1 is fine) whenever you want to search the whole sheet. Find
> also searches only the **active sheet** unless you change **Within** to Workbook, so click the right sheet tab before you
> search.

Click **Options >>** to see the settings that control what counts as a match:

| Option | Choices | Use it when |
|---|---|---|
| **Within** | Sheet (default) or Workbook | You want to search every sheet at once |
| **Search** | By Rows or By Columns | You care about the order Find Next visits cells |
| **Look in** | Formulas (default), Values, Notes, Comments | Formulas searches what's stored in each cell, including hidden cells. Values searches what each cell displays and skips hidden cells |
| **Match case** | Off (default) or on | `HTN` and `htn` should be treated as different |
| **Match entire cell contents** | Off (default) or on | The whole cell must equal your search text, not just contain it |

**Worked example: why "Match entire cell contents" matters.** Suppose you want the number of female patients. The Sex
column holds `F` or `M`, so you Find All for `F`. Excel reports more than 400 cells, because with the default settings a
match can be *anywhere* inside a cell: every Fairhaven, every Fisher and Crawford, every email address with an f in it. Tick **Match
entire cell contents**, Find All again, and the count drops to **281**, the cells that are exactly `F`.

Find also understands **wildcards**:

| Wildcard | Matches | Example |
|:-:|---|---|
| `*` | Any number of characters | `Lake*` with Match entire cell contents finds cells that *start* with Lake: 34 patients in Lakeview Heights |
| `?` | Exactly one character | `PT1000?` finds PT10008 |
| `~` | Treats the next `*` or `?` as a normal character | `~*` finds an actual asterisk |

> ⚠️ **Find All counts cells, not patients.** A count of cells equals a count of patients only when the search text appears in
> one column. Before you report a Find All number, scan the **Cell** column of the results list to make sure every hit is in
> the column you meant.

> 📋 **Version note:** Excel for Mac added Find All in version 16.60 (April 2022). On a Mac, **Control + F** opens the Find
> and Replace dialog. **⌘ + F** may instead put the cursor in a search box, which has no Find All button. The matching option
> may be labeled **Find entire cells only**.

### 8. Freeze Panes, Split, and Zoom

The Patients sheet opens with nothing frozen. Scroll down a few screens and the headers disappear, so you can no longer tell
whether a number is a height or a weight. **Freeze Panes** fixes that by locking rows and columns in place while the rest of
the sheet scrolls. All three commands are on the **View** tab under **Freeze Panes** (Windows and Mac).

| To keep on screen | Select this cell first | Then choose |
|---|---|---|
| Row 1 only | Any cell | **Freeze Top Row** |
| Column A only | Any cell | **Freeze First Column** |
| Row 1 and column A | **B2** | **Freeze Panes** |
| Rows 1 to 3 | **A4** | **Freeze Panes** |
| Rows 1 to 2 and columns A to C | **D3** | **Freeze Panes** |

The rule behind the table: **Freeze Panes freezes every row above the selected cell and every column to its left.** A thin
line marks the frozen edge. To undo it, choose **View → Freeze Panes → Unfreeze Panes** (on Windows, **Alt, W, F, F** toggles
it).

> ⚠️ Press **Ctrl + Home** (Mac: **Control + Home**) before you freeze. Freeze Panes freezes whatever is on screen above and
> left of the selected cell. If you've scrolled down to row 200 and freeze at row 210, rows 200 to 209 stay frozen and rows
> 1 to 199 can't be reached until you unfreeze.

> 📋 Once panes are frozen, **Ctrl + Home** goes to the first cell below and right of the frozen area (B2 if you froze at B2),
> not to A1.

> 💡 **Tip:** Inside an Excel Table you get a bonus. When you scroll down with a Table cell selected, the column letters
> (A, B, C…) are replaced by the Table's headers (PatientID, MRN, LastName…), even without freezing.

**Split** (**View → Split**) divides the window at the active cell into two or four panes, and each pane scrolls on its own.
Use it to compare the top of a long list with the bottom. Click **Split** again to remove it. You can freeze or split, but not
both at once.

**Zoom** changes how large the sheet looks on screen. It doesn't change your data or how the sheet prints.

- Drag the **zoom slider** at the right end of the status bar, or click the percentage next to it to open the **Zoom**
  dialog and type any value from 10% to 400%.
- **View → Zoom to Selection** fits the selected cells to the window, and **View → 100%** resets.
- Windows: **Ctrl + mouse wheel** zooms, and so do **Ctrl + Alt + =** and **Ctrl + Alt + -** (minus) in Microsoft 365.
  On a Mac, use the slider or **View → Zoom**.

Zoom is saved separately for each worksheet, so the Patients sheet can sit at 85% while the Practice sheet stays at 100%.

### 9. Hide and unhide rows and columns

Hiding removes rows or columns from view without deleting them. You'll see it in real workbooks all the time, often to tuck
away helper columns or sensitive fields. You can spot hidden rows or columns two ways: the headers **skip** (column H is
followed by J), and a double line sits between the headers where the hidden ones are.

| Action | Mouse | Windows | Mac |
|---|---|---|---|
| Hide rows | Select the rows, right-click a row number → **Hide** | **Ctrl + 9** | **Control + 9** |
| Unhide rows | Select the rows on **both sides** of the gap, right-click → **Unhide** | **Ctrl + Shift + 9** | **Control + Shift + 9** |
| Hide columns | Select the columns, right-click a column letter → **Hide** | **Ctrl + 0** | **Control + 0** |
| Unhide columns | Select the columns on **both sides** of the gap, right-click → **Unhide** | **Ctrl + Shift + 0** | **Control + Shift + 0** |

You can also use the ribbon: **Home → Format → Hide & Unhide**.

> ⚠️ On many Windows PCs, **Ctrl + Shift + 0** does nothing, because a Windows keyboard-language setting claims that key
> combination. Use right-click → **Unhide** instead.

> 💡 **Tip:** To unhide *everything* on a sheet, click the Select All button (or press Ctrl + A, Mac: ⌘ + A, until the whole
> sheet is selected), then right-click any column letter → **Unhide**, and right-click any row number → **Unhide**.

> 💡 **Tip:** A hidden column A is awkward because there's no column to its left to select. Type `A1` in the Name Box, press
> Enter, then choose **Home → Format → Hide & Unhide → Unhide Columns**.

> 💡 **Tip:** Rows you hide by hand still come along when you copy. To copy only what you can see, select the range, press
> **Alt + ;** (Mac: **⌘ + Shift + Z**) to select only the visible cells, and then copy.

> ⚠️ **Hidden is not deleted, and it's not private.** Hidden cells still hold their data, formulas still use them, and anyone
> who opens the file can unhide them. Never hide patient identifiers as a way to "remove" them before sharing a file.

### 10. Unhide a sheet: how you'll open every answer key in this course

Whole worksheets can be hidden too, and every workbook in this course hides its **Answer Key** and **Bonus Key** sheets that
way. A hidden sheet has no tab at all, so you have to know where to look.

1. Right-click **any** sheet tab (Mac: Control + click, or click with two fingers).
2. Choose **Unhide…**. If it's grayed out, the workbook has no hidden sheets.
3. Click the sheet you want, such as **Answer Key**, and click **OK**.

To hide a sheet again, right-click its tab and choose **Hide**. The ribbon route is **Home → Format → Hide & Unhide →
Unhide Sheet**, and on a Mac the menu bar also has **Format → Sheet → Unhide**.

> 📋 **Version note:** In Microsoft 365 (Windows version 16.0.13525 or later, Mac 16.45 or later, and Excel for the web) you can
> Ctrl + click (Mac: ⌘ + click) or Shift + click to unhide several sheets in one go. Older versions unhide one sheet at a time.

> 📋 If **Unhide…** is grayed out but you're sure something is hidden, either the workbook's structure is protected
> (**Review → Protect Workbook**) or the sheet is "very hidden," a state that can only be set and undone in the VBA editor.
> You'll meet very hidden sheets in Lesson 5.3.

A workbook must always have at least one visible sheet, so Excel won't let you hide the last one.

### 11. Worked example: a one-minute profile of the patient list

Here's how the pieces fit together on the Patients sheet, without a single formula.

1. Press **Ctrl + Home** (Mac: **Control + Home**) to start at A1, then choose **View → Freeze Panes → Freeze Top Row** so
   the headers stay visible.
2. **How many patients?** Press **Ctrl + ↓** (Mac: **⌘ + ↓**) in column A. The Name Box shows A501. Row 1 is the header,
   so the list holds 500 patients.
3. **How tall are they?** Type `N2:N501` in the Name Box and press Enter. The status bar shows an average height of 63.21
   inches, and Minimum and Maximum show the range: 19 to 76.9 inches.
4. **Do we need Spanish interpreters?** Click A1 so only one cell is selected, then press **Ctrl + F** (Mac: **Control + F**),
   type `Spanish`, and click **Find All**. The dialog reports *41 cell(s) found*, and the **Cell** column shows they're all
   in column K, PreferredLanguage. So 41 of the 500 patients prefer Spanish.
5. **Sanity check.** 41 out of 500 is about 8%. A number like 410 would mean you'd searched something too broad, so you'd
   go back and look at the results list.

That last step matters as much as the shortcuts. **Always ask whether a number is plausible** before you pass it on.

### 12. Shortcuts and version notes

| Action | Windows | Mac |
|---|---|---|
| Jump to the edge of the data | Ctrl + arrow | ⌘ + arrow |
| Select to the edge of the data | Ctrl + Shift + arrow | ⌘ + Shift + arrow |
| Go to A1 | Ctrl + Home | Control + Home, or Control + Fn + ← |
| Go to the last used cell | Ctrl + End | Control + End, or Control + Fn + → |
| Select the column / row | Ctrl + Space / Shift + Space | Control + Space / Shift + Space |
| Select the data block, then the whole sheet | Ctrl + A | ⌘ + A |
| Next / previous sheet | Ctrl + Page Down / Ctrl + Page Up | Option + → / Option + ← |
| Scroll back to the active cell | Ctrl + Backspace | Control + Delete |
| Go To | F5 or Ctrl + G | Control + G |
| Find / Find Next after closing the dialog | Ctrl + F / Shift + F4 | Control + F / ⌘ + G |
| Hide / unhide rows | Ctrl + 9 / Ctrl + Shift + 9 | Control + 9 / Control + Shift + 9 |
| Hide / unhide columns | Ctrl + 0 / Ctrl + Shift + 0 (or right-click → **Unhide**) | Control + 0 / Control + Shift + 0 |
| Select visible cells only | Alt + ; | ⌘ + Shift + Z |
| Collapse or expand the ribbon | Ctrl + F1 | ⌘ + Option + R |
| Freeze Panes | Alt, W, F, F | **View → Freeze Panes** |

| Feature | Version |
|---|---|
| 1,048,576 rows × 16,384 columns (A to XFD) | `.xlsx` files in Excel 2007 and later. An `.xls` file in Compatibility Mode has 65,536 rows × 256 columns |
| Find All on a Mac | Excel for Mac 16.60 (April 2022) and later |
| KeyTips (Alt, W, F, F) | Excel for Windows only |
| Quick Access Toolbar hidden by default | Recent Microsoft 365 updates for Windows. Right-click the ribbon → **Show Quick Access Toolbar** |
| Click a status bar value to copy it | Microsoft 365 (version 2206 or later) and Excel 2024, on Windows |
| Unhide several sheets at once | Microsoft 365 (Windows 16.0.13525 or later, Mac 16.45 or later) and Excel for the web |
| Settings on a Mac | **Excel → Settings**, called **Excel → Preferences** in older versions |

## 🧪 Hands-on practice

Download [`1.1-excel-interface-navigation.xlsx`](1.1-excel-interface-navigation.xlsx) and open the **Practice** sheet. Each
answer is something you read off the screen, so type the value itself (a name, a number, a date, or an address) into the
yellow cell and press **Enter**. The **Check** column turns green when you're right.

<!-- BEGIN GENERATED: practice -->
Every task uses the Patients sheet (500 patients, rows 2–501) unless it says otherwise. No formulas needed: navigate, look, and type what you find. Mac users: for Ctrl + arrow shortcuts such as Ctrl + ↓ and Ctrl + Shift + ↓, press ⌘ instead of Ctrl.

| # | Task | Hint |
|:-:|------|------|
| 1 | On the Patients sheet, click in the Name Box (the box at the left end of the formula bar), type C347, and press Enter. What last name is in that cell? | The Name Box jumps to any address you type |
| 2 | What is the address of the bottom-right cell of the patient data? From A1, press Ctrl + ↓ to find the last row, go back with Ctrl + ↑, then press Ctrl + → to find the last column. Type the address as column letter + row number (like B12). | Column letter from Ctrl + →, row number from Ctrl + ↓. Ctrl + End confirms |
| 3 | Use the Name Box to go to cell A600, which is in an empty row below the data, then press Ctrl + →. Excel races across the empty row to the very last column of the worksheet. What is the address of the cell you land on? | A worksheet has 16,384 columns |
| 4 | Click M1 (the PCPProviderID header) and press Ctrl + ↓ once. On which row number does Excel stop? | Ctrl + arrow stops at the edge of a block of filled cells |
| 5 | Select the Email data cells J2:J501: type J2:J501 in the Name Box and press Enter. How many patients have an email address on file? Read Count on the status bar. | Count = cells that aren't empty |
| 6 | Select the WeightLb values O2:O501 with Go To: press F5 or Ctrl + G (Mac: Control + G), type O2:O501, and press Enter. What is the average weight in pounds? Round to 1 decimal place. | Average is on the status bar by default |
| 7 | Turn on Minimum in the status bar: right-click the status bar and tick Minimum. Then click cell F2 (the first DOB) and press Ctrl + Shift + ↓ to select every date of birth. What is the date of birth of the oldest patient? Type it as month/day/year, the way the status bar shows it. (If your computer uses day/month dates, type the month as a word instead, such as 15 Mar 1950.) | The oldest patient has the earliest date, so look at Minimum |
| 8 | How many patients live in Millbrook? On the Patients sheet, click a single cell such as A1 so Find searches the whole sheet. Then press Ctrl + F (Mac: Control + F), type Millbrook, click Find All, and read the count at the bottom of the dialog. | Find All shows '… cell(s) found' |
| 9 | Column I is hidden (the column letters jump from H to J). Unhide it, then click cell A1 and use Ctrl + F (Mac: Control + F) to find patient PT12376. What is that patient's phone number? | Select the columns on both sides of the gap, right-click → **Unhide** |
| 10 | Patient PT13312 has a PrimaryPayerID in column L. The payer names are on the Payers sheet, which is hidden. Unhide it (right-click any sheet tab → **Unhide…**) and type the name of this patient's payer. | Right-click a sheet tab → **Unhide…** |
| 11 | How many worksheets does this workbook contain in total, hidden ones included? Count the tabs you can see, then right-click a tab → **Unhide…** to see what's still hidden. Look, but don't unhide the answer keys yet! | Visible tabs + the names listed in the Unhide dialog |
| 12 | You want row 1 (the headers) and columns A:B (PatientID and MRN) to stay on screen while you scroll. Which cell must you select before choosing **View → Freeze Panes → Freeze Panes**? Type its address. | Excel freezes everything above and to the left of the selected cell |
<!-- END GENERATED: practice -->

## ✅ Answer key

The workbook has a hidden **Answer Key** sheet. Open it the way section 10 shows: right-click any sheet tab → **Unhide…** →
*Answer Key*. Its *Live result* column uses formulas (you'll write your own from Lesson 1.4 on) to confirm each answer straight
from the data. The same answers are below, collapsed so you don't see them by accident.

<!-- BEGIN GENERATED: answers -->
<details>
<summary><b>🔑 Show the answer key</b> (try every task before you open this)</summary>

**1. Name Box jump to C347**

- **Answer:** Coleman
- **Solution:**

1. Click the Name Box, type C347, press Enter.
2. Read the cell (or the formula bar).


The Name Box always shows the address of the active cell, and it works in reverse too: type an address, press Enter, and Excel takes you there. That's faster than scrolling 345 rows. Column C holds LastName, so C347 is the last name of the patient in row 347.

**2. Bottom-right cell of the data**

- **Answer:** Q501
- **Solution:**

1. Click A1 and press Ctrl + ↓: the active cell becomes A501.
2. Press Ctrl + ↑ to return to A1, then Ctrl + →: the active cell becomes Q1.
3. Combine them: Q501. Press Ctrl + End (Mac: Control + End, or Control + Fn + →) to confirm. It jumps straight there.


Ctrl + arrow jumps to the edge of the block of filled cells. Column A and row 1 have no gaps, so the jumps land on the true last row (501: 500 patients plus the header row) and the true last column (Q). Ctrl + End goes straight to the last used cell, Q501. The hidden column I doesn't change the answer, because hiding a column doesn't remove it.

**3. The last column of a worksheet**

- **Answer:** XFD600
- **Solution:**

1. Type A600 in the Name Box and press Enter.
2. Press Ctrl + →. The Name Box shows XFD600.
3. Press Ctrl + ← to come back to column A.


When the row is empty, Ctrl + → has no data to stop at, so it goes to the edge of the sheet: column XFD, the 16,384th column. Ctrl + ↓ in an empty column goes to row 1,048,576. Every modern worksheet has exactly 1,048,576 rows × 16,384 columns.

**4. Ctrl + ↓ stops at a gap**

- **Answer:** 34
- **Solution:**

1. Click M1 (or type M1 in the Name Box).
2. Press Ctrl + ↓. The Name Box shows M34.


M35 is empty because that patient has no primary care provider on file. Ctrl + ↓ stops at the last filled cell before the gap, M34, even though the data continues to row 501. Never assume Ctrl + ↓ found the bottom of a column with blanks. Check the row number, or press Ctrl + ↓ again to keep jumping.

**5. Status bar Count (patients with an email)**

- **Answer:** 298
- **Solution:**

1. Type J2:J501 in the Name Box and press Enter.
2. Read Count on the status bar.


The status bar's Count counts every non-empty cell, text included, so it counts the email addresses and skips the blanks. Typing the range in the Name Box selects exactly 500 cells. Ctrl + Shift + ↓ from J2 would be a trap, because the Email column has gaps: it would select only J2:J3. In a Table you can also click any Email cell and press Ctrl + Space (Mac: Control + Space) to select just that column's data.

**6. Status bar Average (weight, lb)**

- **Answer:** 167.8
- **Solution:**

1. Press F5 (or Ctrl + G), type O2:O501 in Reference, press Enter.
2. Read Average on the status bar and round it to 1 decimal place.


Go To selects any range you type, however large. The status bar's Average is 167.8442, which rounds to 167.8. The average is pulled down by children in the list: the lightest patient weighs 9.7 lb, which is a baby and not a typo. Always glance at Minimum and Maximum before trusting an average.

**7. Status bar Minimum (oldest patient's DOB)**

- **Answer:** 02/06/1926
- **Solution:**

1. Right-click the status bar and tick Minimum (and Maximum while you're there).
2. Click cell F2 and press Ctrl + Shift + ↓ to select F2:F501.
3. Read Minimum on the status bar.


Excel stores dates as numbers that grow by 1 each day, so the earliest date is the smallest number and Minimum finds it. The status bar shows it as a date because the cells are formatted as dates. (If you see a plain number such as 9534 instead, that's the date's serial number, and the check accepts it.) Ctrl + Shift + ↓ is safe here because the DOB column has no blanks.

**8. Find All: patients in Millbrook**

- **Answer:** 31
- **Solution:**

1. Click A1 (one cell, so Find searches the whole sheet).
2. Press Ctrl + F, type Millbrook, and click Find All.
3. Read '… cell(s) found' at the bottom of the dialog.


Find All lists every matching cell on the sheet and counts them. That's a count of patients here only because the word Millbrook appears in the City column and nowhere else. If a search term could also appear inside other text, such as an email address, turn on Match entire cell contents under Options. Clicking a single cell first matters too. When several cells are selected, Find searches only inside the selection, so with the DOB column still selected from the previous task, Find All would report nothing.

**9. Unhide column I (phone number)**

- **Answer:** (555) 740-8462
- **Solution:**

1. Click the H column header, Shift + click the J header, then right-click → **Unhide**.
2. Click A1, press Ctrl + F, type PT12376, click Find Next, then close the dialog.
3. Read column I on row 298.


You can't click a hidden column, so you select the columns on both sides of it (H and J) and choose Unhide. Click A1 before you search, because while those three columns are selected, Find looks only inside them. The patient is on row 298. Hiding never deletes data: the phone numbers were there all along, so a hidden column is not a safe place for confidential data.

**10. Unhide the Payers sheet**

- **Answer:** Summit Choice PPO
- **Solution:**

1. On the Patients sheet, click A1, press Ctrl + F, and find PT13312: row 415, payer ID PY06.
2. Right-click any sheet tab → **Unhide…**, pick Payers, and click OK.
3. On the Payers sheet, PY06 is Summit Choice PPO.


Hidden sheets don't show a tab, so the only clue is the Unhide… command becoming available. This is the same move you'll use to open the Answer Key in every lesson. Small lookup lists like this one are often hidden to keep a workbook tidy, and in Lesson 2.6 you'll learn to pull names from them automatically with a lookup formula.

**11. Worksheets in this workbook**

- **Answer:** 7
- **Solution:**

1. Count the visible tabs: Start Here, Practice, Patients, Payers (now unhidden), Bonus.
2. Right-click a tab → **Unhide…**: the list shows Answer Key and Bonus Key. Click Cancel.
3. 5 + 2 = 7.


A workbook is the file, and each worksheet is one tab inside it. This file has 7 worksheets, and 2 of them are still hidden: the Answer Key and the Bonus Key. The total is the same whether or not you unhid Payers first, because hiding a sheet doesn't remove it.

**12. Freeze Panes: which cell to select**

- **Answer:** C2
- **Solution:**

1. Click C2.
2. Choose **View → Freeze Panes → Freeze Panes**.
3. Scroll down and right: row 1 and columns A:B stay put. (**View → Freeze Panes → Unfreeze Panes** undoes it.)


Freeze Panes freezes the rows above the active cell and the columns to its left. To keep 1 row and 2 columns, select the cell just below row 1 and just right of column B: C2. Selecting B2 would keep row 1 but only column A, and A2 would freeze only row 1, which is the same as Freeze Top Row.

</details>
<!-- END GENERATED: answers -->

## 🏆 Bonus challenge

<!-- BEGIN GENERATED: bonus -->
The population-health team is starting a blood-pressure outreach program and will build its mailing list from this patient index. Before the letters go out, the data steward asks you to check four facts. Set up the Patients sheet first: if you froze panes earlier, choose **View → Freeze Panes → Unfreeze Panes**. Then press Ctrl + Home (Mac: Control + Home), click B2, and choose **View → Freeze Panes → Freeze Panes** so the headers and the PatientID column stay in view. Also turn on Minimum and Maximum in the status bar. You'll look everything up on the Patients sheet and type your answers on the Bonus sheet, without a single formula.

Work on the **Bonus** sheet of the workbook.

- **B1.** Each outreach letter is signed by the patient's primary care provider (PCP), and a blank PCPProviderID means no PCP is on file. Use Go To (F5 or Ctrl + G, Mac: Control + G) to select M2:M501. How many of the 500 patients have no PCP? *(Hint: The status bar counts filled cells, but you want the empty ones)*
- **B2.** The program targets every patient with hypertension, coded HTN. Click a single cell such as A1, so Find searches the whole sheet rather than the column you just selected. Press Ctrl + F (Mac: Control + F), click Options >> and make sure Match entire cell contents is OFF, then Find All for HTN. How many patients have HTN anywhere in their ChronicConditions list? *(Hint: Partial matches count: 'HTN;DM' contains HTN)*
- **B3.** Patients whose only condition is hypertension will get a shorter letter. Turn ON Match entire cell contents and Find All for HTN again. How many patients have hypertension as their ONLY recorded chronic condition? *(Hint: Match entire cell contents finds cells that are exactly HTN)*
- **B4.** Finally, the steward wants to confirm that the largest weight on file is real and not a typo. Which patient is the heaviest? Select O2:O501, read Maximum on the status bar, then use Find to locate that weight and read the PatientID in the frozen column A. Type the PatientID. *(Hint: Status bar Maximum, then Ctrl + F for that number)*
<!-- END GENERATED: bonus -->

<!-- BEGIN GENERATED: bonus-answers -->
<details>
<summary><b>🔑 Show the bonus solution</b> (give it a real try first)</summary>

**B1. Patients with no PCP on file**

- **Answer:** 31
- **Solution:**

1. Press F5, type M2:M501, press Enter.
2. The status bar shows Count: 469.
3. 500 − 469 = 31.


Count only counts non-empty cells, so it tells you how many patients do have a PCP (469). The range M2:M501 covers rows 2 to 501, which is 500 cells, so 500 − 469 leaves 31 blanks. Any of those patients who have hypertension need a PCP assigned before a letter can go out. (If you select by dragging instead, the Name Box shows the size of the selection while you drag, such as 500R x 1C.)

**B2. Find All: HTN anywhere in the list**

- **Answer:** 144
- **Solution:**

1. Click A1, press Ctrl + F, click **Options >>**, and untick Match entire cell contents.
2. Type HTN and click Find All.
3. Read the count at the bottom of the dialog.


By default Find matches text anywhere inside a cell, so it finds HTN on its own and also inside lists like HTN;DM and HTN;HF. No other column contains the letters HTN, so every hit is a patient with hypertension.

**B3. Find All: HTN as the only condition**

- **Answer:** 48
- **Solution:**

1. In the Find dialog, tick Match entire cell contents.
2. Find All for HTN.
3. Read the count.


With Match entire cell contents on, a cell must equal HTN exactly, so HTN;DM no longer counts. The gap between the two answers (144 − 48 = 96) is the number of hypertensive patients with at least one other condition. When you finish this challenge, turn the option off again, because Find keeps your settings until you close Excel.

**B4. The heaviest patient**

- **Answer:** PT13784
- **Solution:**

1. Select O2:O501 (Name Box or Go To) and read Maximum: 319.2.
2. With the weights still selected, press Ctrl + F, type 319.2 in Find what, and click Find Next. (Match entire cell contents can be on or off here, because no other cell on the sheet contains 319.2.)
3. Excel selects O474. The frozen column A shows PT13784.


The status bar tells you what the largest value is but not where it is. Find tells you where, and leaving the weights selected helps, because Find then searches only inside the selection. Because panes are frozen at B2, column A stays on screen when Find scrolls over to the WeightLb column, so you can read the ID without losing your place. Teams use this check to plan bariatric beds and lift equipment, and it's also how you'd spot a typo such as 3192 lb.

</details>
<!-- END GENERATED: bonus-answers -->

## Key takeaways

- The **Name Box** shows where you are and takes you anywhere: type an address or a range and press Enter. **F5** (Go To)
  does the same and remembers where you came from.
- **Ctrl + arrow** (Mac: ⌘ + arrow) jumps to the edge of a block of data, and blanks stop it early. Add **Shift** to select
  as you jump, and check the row number before you trust it.
- The **status bar** summarizes any selection. **Count** counts every filled cell, while **Numerical Count** counts only
  numbers and dates. Right-click the status bar to add Minimum and Maximum.
- **Find All** counts matching cells. Click a single cell first so Find searches the whole sheet, and turn on **Match entire
  cell contents** when partial matches would inflate the count.
- **Freeze Panes** keeps every row above and every column left of the selected cell on screen.
- **Hidden is not deleted.** Hidden rows, columns, and sheets still hold their data, and right-click a sheet tab →
  **Unhide…** opens every answer key in this course.

<!-- BEGIN GENERATED: nav -->
---

🏠 [Course home](../../README.md) · **Next:** [1.2 Data Entry, AutoFill & Editing](../02-data-entry-autofill/README.md) ➡️
<!-- END GENERATED: nav -->
