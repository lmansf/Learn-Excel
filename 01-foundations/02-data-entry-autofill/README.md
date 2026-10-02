# Lesson 1.2 · Data Entry, AutoFill & Editing

> **Level:** Beginner · **Time:** about 40 minutes · **Workbook:** [`1.2-data-entry-autofill.xlsx`](1.2-data-entry-autofill.xlsx)
> **Data:** A December 2025 rotating-RN schedule for Medical-Surgical 4 West, a registration record, a Q15 observation form, 4 West's bed list, supply room, and December admissions, plus a January 2026 clinic template for the bonus. The schedule, admissions, and supplies come from [`shifts.csv`](../../data/README.md#shiftscsv), [`encounters.csv`](../../data/README.md#encounterscsv), and [`supply_inventory.csv`](../../data/README.md#supply_inventorycsv) in the data dictionary.

Every number a hospital reports started as something a person typed or pasted. A registrar types an MRN, a charge nurse
builds next month's schedule, and a buyer pastes restock quantities into a requisition. When an MRN loses its leading zeros, the
patient's records stop matching. When "3-12" turns into March 12, a bed assignment is lost. When a schedule is typed cell by cell,
it takes an hour instead of a minute. In this lesson you'll learn how Excel decides what you typed, how to fill dates, times,
and IDs in seconds, how to reshape and clean what you paste, and how to avoid the traps that quietly corrupt healthcare data.

## What you'll learn

- Recognize how Excel stores text, numbers, dates, times, and TRUE/FALSE
- Enter and edit data efficiently (F2, Ctrl+Enter, Ctrl+D, Alt+Enter)
- Create series with AutoFill, the Fill Series dialog, and Flash Fill
- Use Copy, Paste Special (values, transpose, formats), and Find & Replace
- Avoid classic traps: lost leading zeros, numbers stored as text, accidental dates

## 📖 Guide

### 1. How Excel decides what you typed

When you press Enter, Excel looks at what you typed and stores it as one of a few **data types**. The type decides what you can do
with the entry later. Numbers can be added and sorted numerically, while text can only be matched and sorted alphabetically.

| Data type | You type | Excel stores | Default alignment |
|---|---|---|:-:|
| **Text** | `Medina, Lisa` · `NKDA` · `4W-12` | the characters exactly | left |
| **Number** | `98.6` · `36` · `$25` · `85%` | a number (the `$` and `%` become formatting, so 85% is stored as 0.85) | right |
| **Date** | `12/1/2025` · `Dec 1, 2025` · `1-Dec-25` | a number: 45,992 (see section 2) | right |
| **Time** | `7:30` · `7:30 PM` · `19:30` | a fraction of a day: 7:30 is 0.3125 | right |
| **Logical** | `true` · `FALSE` | TRUE or FALSE (Excel capitalizes it) | center |
| **Error** | `#N/A` | an error value | center |

The **alignment clue** is the fastest way to check a type. If you haven't changed a cell's alignment, numbers (including dates and
times) sit on the right, text sits on the left, and TRUE/FALSE sit in the center. A number that sits on the left is really text.

Two more places tell you the truth:

- **The formula bar** shows what's actually stored. A cell can display `12-Mar` while the formula bar shows `3/12/2025`.
- **The status bar.** Select a range and the status bar shows **Count** (every non-empty cell) and, if you turn it on, **Numerical
  Count** (only the numbers). Right-click the status bar to choose which statistics appear. Text and TRUE/FALSE are never part of
  Numerical Count.

> ⚠️ **Numbers stored as text.** Numbers pasted from other systems (a bedside scale, a lab interface, a web page) often arrive as
> text. Excel marks them with a small green triangle in the cell's top-left corner and lines them up on the left. They look
> fine, but they're skipped by SUM, sort in the wrong order, and won't match real numbers. To fix them, select the cells, click
> the **error button** that appears beside them (a yellow diamond with an exclamation mark), and choose **Convert to Number**.
> Section 8 shows a second fix with Paste Special.

### 2. Dates and times are numbers

Excel stores every date as a **serial number**: the count of days since the start of 1900, where 1/1/1900 is day 1. A time is
stored as a **fraction of a day**. Put them together and a single number holds both a date and a time.

| You see | Excel stores | Why |
|---|---|---|
| 01/01/1900 | 1 | Day 1 |
| 12/01/2025 | 45,992 | Day 45,992, counting 1/1/1900 as day 1 |
| 12/31/2025 | 46,022 | 30 days later |
| 12:00 | 0.5 | Half a day |
| 07:00 | 0.2917 | 7 ÷ 24 |
| 0:15 (15 minutes) | 0.0104 | 15 ÷ 1,440 minutes in a day |
| 12/01/2025 07:00 | 45,992.2917 | The date plus the time |

To see the stored number, select the cell and apply the **General** format: press **Ctrl + Shift + ~** (Mac: **⌃ + Shift + ~**) or
choose **Home → Number Format → General**. Press **Ctrl + Z** (Mac: **⌘ + Z**) to put the date format back. Because dates are
numbers, Excel can sort them, count the days between them, and fill them in a series. Lesson 2.3 covers date and time math in depth.

**Entering dates and times**

- Type dates the way your computer's region expects. This course uses the US order, month/day/year, so `12/1/2025` is
  December 1. On a computer set to UK format, the same keystrokes mean 12 January.
- If you leave out the year (`12/1`), Excel uses the current year.
- Type times with a colon: `7:30`, `19:30`, or `7:30 PM` (with a space before PM). `7:30` on its own means 7:30 AM.
- **Ctrl + ;** (Mac: **⌃ + ;**) enters today's date and **Ctrl + Shift + ;** (Mac: **⌘ + ;**) enters the current time. Both are
  typed-in values that never change, which makes them good for logging when something happened.

> ⚠️ **Two-digit years.** Excel reads `00`–`29` as 2000–2029 and `30`–`99` as 1930–1999. A patient born on 8/28/1926 and typed as
> `8/28/26` gets a date of birth in **2026**. Always type four-digit years for dates of birth.

### 3. Keeping text as text: the apostrophe and the Text format

Some entries look like numbers but are really **identifiers**: medical record numbers (MRNs), ZIP codes, phone numbers,
insurance member IDs, and lot numbers. You never do math on them, and their leading zeros matter. Store them as text.

| Method | How | When to use it |
|---|---|---|
| **Apostrophe** | Type `'00412345`. The apostrophe isn't displayed or stored as part of the value. | One-off entries |
| **Text format first** | Select the column, press **Ctrl + 1** (Mac: **⌘ + 1**), choose **Text**, then type. | A whole column of IDs |
| Custom format `00000000` | Shows zeros in front of a number | ⚠️ Display only. The stored value is still the number 412345 |

The rule of thumb: **if you would never add it up, store it as text.**

Excel flags a text entry that looks like a number with the green triangle described in section 1. For an MRN that's expected,
so choose **Ignore Error** from the error button instead of converting it.

> ⚠️ **Format first, then type.** Changing a cell to Text *after* you typed `00412345` doesn't bring the zeros back, because they
> were already gone. Re-type the entry after formatting.

> 📋 A cell formatted as Text treats everything as text, including formulas. If you type `=SUM(B2:B10)` into a Text cell, you
> see the formula itself instead of a result. Set the cell back to General and re-enter the formula.

### 4. Entering and editing efficiently

**Moving while you type.** Enter confirms an entry and moves down, while Tab confirms and moves right. Shift reverses the
direction. If you type across a row with Tab and then press Enter, Excel jumps back to the column where you started, one row
down. That makes typing records row by row fast.

**Enter mode and Edit mode.** The status bar shows which mode you're in. When you start typing in a cell you're in **Enter** mode,
and the arrow keys confirm the entry and move to the next cell. Press **F2** (Mac: **⌃ + U**, or **F2** with the fn key) to switch
to **Edit** mode, where the arrow keys move the cursor inside the text so you can fix a typo without retyping. You can also edit
in the formula bar. Press **Esc** to cancel an entry before you confirm it.

| Action | Windows | Mac |
|---|---|---|
| Confirm and move down / right | Enter / Tab | Return / Tab |
| Confirm and move up / left | Shift + Enter / Shift + Tab | Shift + Return / Shift + Tab |
| Cancel the entry | Esc | Esc |
| Edit the active cell | F2 | ⌃ + U (or fn + F2) |
| Put the same entry in every selected cell | Ctrl + Enter | ⌘ + Return |
| Copy the top cell down / the left cell right | Ctrl + D / Ctrl + R | ⌘ + D / ⌘ + R |
| Start a new line inside the cell | Alt + Enter | ⌃ + ⌥ + Return (or ⌥ + Return) |
| Today's date / current time | Ctrl + ; / Ctrl + Shift + ; | ⌃ + ; / ⌘ + ; |
| Pick from entries already in the column | Alt + ↓ | ⌥ + ↓ |
| Undo / Redo | Ctrl + Z / Ctrl + Y | ⌘ + Z / ⌘ + Y |

**Ctrl + Enter** is the most useful shortcut in this lesson. Select several cells, type once, and press Ctrl + Enter (Mac: ⌘ + Return).
The entry goes into *every* selected cell. The cells don't have to touch: hold **Ctrl** (Mac: **⌘**) while you click or drag to
add cells to the selection. Whatever you type appears in the last cell you clicked (the **active cell**), and Ctrl + Enter copies
it to the rest. On the Schedule sheet, that lets you mark every weekend row in one move.

**Ctrl + D** (Mac: **⌘ + D**), called **fill down**, copies the top cell of the selection into every cell below it. With a single
cell selected, it copies the cell directly above. **Ctrl + R** (Mac: **⌘ + R**) does the same to the right.

**Alt + Enter** (Mac: **⌃ + ⌥ + Return**) adds a line break *inside* a cell, which is handy for two-line notes such as an allergy
and an isolation status. Excel turns on **Wrap Text** for you so both lines show.

> 💡 **Tip:** **AutoComplete** suggests a finish for text you type, based on entries already in the same column. Type `Pen` in a
> column that already contains `Penicillin` and Excel offers the rest. Press Enter to accept, or keep typing to ignore it. It's
> fast, but it can also slip in the wrong word, because Enter, Tab, and Ctrl + Enter all accept the suggestion. Watch the cell
> before you confirm, and press **Backspace** (Mac: **delete**) to remove a suggestion you don't want. You can turn it off
> under **File → Options → Advanced → Enable AutoComplete for cell values** (Mac: **Excel → Settings → AutoComplete**).

### 5. AutoFill and the fill handle

The **fill handle** is the small square at the bottom-right corner of the selected cell or range. When you point at it, the
pointer becomes a thin black **+**. Drag it down or across and Excel **AutoFills**: it either copies the selection or continues a
series, depending on what you started with.

| You start with | Drag the fill handle | Hold Ctrl while you drag (Windows) |
|---|---|---|
| `7` | 7, 7, 7 (copies) | 8, 9, 10 |
| `7` and `8`, both selected | 9, 10, 11 | 7, 8, 7, 8 (repeats) |
| `12/01/2025` | 12/02/2025, 12/03/2025 … | 12/01/2025 (copies) |
| `07:00` | 08:00, 09:00 … (one **hour** at a time) | 07:00 (copies) |
| `07:00` and `07:15`, both selected | 07:30, 07:45 … | 07:00, 07:15, 07:00 … |
| `Mon` or `Monday` | Tue, Wed … or Tuesday, Wednesday … | copies |
| `Dec` or `December` | Jan, Feb … or January, February … | copies |
| `4W-01` (text ending in a number) | 4W-02, 4W-03 … | 4W-01 (copies) |
| `D` (text with no number) | D, D, D | D, D, D |

The pattern behind the table: **one number is copied, but one date, time, or text-ending-in-a-number counts up.** Two selected
cells always tell Excel the step size, so `07:00` and `07:15` produce 15-minute steps.

**Double-click instead of dragging.** Double-click the fill handle and Excel fills *down* as far as the data in the neighboring
column goes. On the Beds sheet, the Room column runs to row 39, so double-clicking the first bed label fills the list down to
row 39, one label per bed.

> ⚠️ Double-click stops at the first empty cell in the neighboring column. If that column has a gap, check that the fill reached the
> bottom. It only fills downward, so for a row of headers you have to drag.

**The Auto Fill Options button.** After you fill, a small button appears at the corner of the range. Click it to change what
happened: **Copy Cells**, **Fill Series**, **Fill Formatting Only**, **Fill Without Formatting**, and for dates **Fill Days**,
**Fill Weekdays**, **Fill Months**, or **Fill Years**. On a Mac, this button is the simplest way to switch between copying and a
series. On Windows you can also drag with the *right* mouse button to choose from the same menu before anything is filled.

> 💡 **Custom lists.** Weekday and month names fill automatically because they're built-in **custom lists**. You can add your
> own, such as `Day, Evening, Night` or your hospital's unit names, under **File → Options → Advanced → General → Edit Custom
> Lists…** (Mac: **Excel → Settings → Custom Lists**). After that, typing `Day` and dragging gives Evening, Night, Day …

> 📋 AutoFill also copies formatting. If you fill from a white cell into yellow input cells, the yellow disappears. Choose **Fill
> Without Formatting** from the Auto Fill Options button if you want to keep the destination's look.

### 6. The Fill Series dialog: exact steps and stop values

Dragging is quick, but sometimes you need precision: "every 15 minutes until 18:45" or "every weekday until the end of the month."
Select the starting cell (or the whole range to fill), then choose **Home → Fill → Series…**. The path is the same on a Mac, and on
Windows the KeyTips are **Alt, H, F, I, S**.

| Setting | What it does |
|---|---|
| **Series in** | **Rows** fills across, **Columns** fills down |
| **Type** | **Linear** adds the step each time. **Growth** multiplies by the step. **Date** steps in date units. **AutoFill** behaves like dragging |
| **Date unit** | **Day**, **Weekday** (skips Saturdays and Sundays), **Month**, or **Year** |
| **Step value** | How much to add each time. For times, type a time such as `0:15` |
| **Stop value** | The last value allowed. Excel stops before it would go past this |
| **Trend** | Fits a straight line through several selected values. You won't need it in this lesson |

There are two ways to tell Excel how far to go. If you select the whole range first, Excel fills exactly that range and you can
leave Stop value empty. If you select only the first cell, you must give a Stop value.

Clinic days for December 2025, weekdays only, with 12/01/2025 typed in A2 and only A2 selected:

```
Series in:   Columns
Type:        Date
Date unit:   Weekday
Step value:  1
Stop value:  12/31/2025
Result:      23 dates, Mon 12/01/2025 through Wed 12/31/2025, no weekends
```

Post-op vital signs every 30 minutes for two hours, with 08:00 typed in A2:

```
Series in:   Columns
Type:        Linear
Step value:  0:30
Stop value:  10:05
Result:      08:00, 08:30, 09:00, 09:30, 10:00
```

> ⚠️ **Set time stop values a few minutes late.** Times are fractions with long decimal tails, so after many steps the running
> total can differ from the exact time by a tiny amount. If the last value lands a hair past the stop, Excel leaves it out. A stop
> of 10:05 instead of 10:00 avoids that.

> 📋 If your version rejects a time such as `0:30` in the Step value box, type the step as a fraction of a day instead:
> 30 ÷ 1,440 = 0.0208333333 for 30 minutes, or 15 ÷ 1,440 = 0.0104166667 for 15 minutes. Or skip the dialog and use the
> two-cell AutoFill pattern from section 5.

### 7. Flash Fill

**Flash Fill** watches you type an example or two next to your data, works out the pattern, and fills the rest of the column.
It's ideal for pulling apart or recombining text, such as first names out of "Last, First" or initials out of a full name.

1. Type the result you want for the first row in the column next to your data, and press **Enter**.
2. Press **Ctrl + E**, or choose **Data → Flash Fill** (this menu works on both Windows and Mac). Often Excel shows a gray preview
   as you start the second row, and pressing Enter accepts it.
3. **Check the results.** Flash Fill is guessing from your examples. If some rows come out wrong, press **Ctrl + Z**
   (Mac: **⌘ + Z**), type a second or third example that shows the pattern more clearly, and run Flash Fill again. Fix any odd
   row by hand.

> 💡 **Tip:** Pick a first example that can only mean one thing. For `Richardson, Rebecca`, the label `Rebecca R.` could use the
> initial of either name, because both start with R. Type a second example from a row where the initials differ, such as
> `Russell Y.` for `Yilmaz, Russell`, before you run Flash Fill.

| Source | Your example | Flash Fill gives the rest |
|---|---|---|
| `Kennedy, Linda` | `Linda` | first names |
| `Kennedy, Linda` | `L. Kennedy` | initial plus last name |
| `Kennedy, Linda` | `KENNEDY` | last names in capitals |

> ⚠️ Flash Fill produces **typed-in values**, not formulas. If a source name changes later, the Flash Fill result does not. When
> the source data will keep changing, use a text formula instead (Lesson 2.2).

> 📋 **Versions:** Flash Fill needs Excel 2013 or later on Windows, or Excel 2019 / Microsoft 365 on a Mac. If nothing happens,
> make sure **File → Options → Advanced → Automatically Flash Fill** is ticked, and that your example column touches the data.

### 8. Copy, Paste, and Paste Special

**Copy** (Ctrl + C, Mac: ⌘ + C) leaves the original in place, while **Cut** (Ctrl + X, Mac: ⌘ + X) moves it. Both show a moving
dashed border, sometimes called "marching ants," until you paste or press Esc. **Paste** (Ctrl + V, Mac: ⌘ + V) brings
*everything*: values, formulas, and formatting.

**Paste Special** lets you choose *which part* of the copied cells to paste. Open it with **Ctrl + Alt + V** (Mac: **⌃ + ⌘ + V**),
or click the arrow under **Home → Paste**.

| Option | What it pastes | Healthcare example |
|---|---|---|
| **Values** | Only the results, never formulas | Freeze restock quantities before sending a requisition |
| **Formats** | Only the look (fonts, fills, number formats) | Make a new month's census sheet match last month's |
| **Formulas** | Formulas without formatting | Reuse a calculation in a differently styled report |
| **Column widths** | Only the widths | Line up a copied table with the original |
| **Transpose** (checkbox) | Rows become columns and columns become rows | Turn a dates-down schedule into a staff-down roster |
| **Multiply / Add / Subtract / Divide** | Combines the copied number with the cells you paste onto | Apply a 4% vendor price increase |
| **Skip blanks** (checkbox) | Doesn't overwrite destination cells with copied blanks | Merge two partial lists |

**Paste Values matters most.** When you copy a cell that holds a formula, a normal paste copies the *formula*, and its cell
references shift to the new location. On the Supplies sheet, SuggestedOrder in H4 is `=MAX(0,G4-F4)`. Paste it four columns to the
left and it becomes `=MAX(0,C4-B4)`, which points at the wrong cells. Paste Values pastes the number the formula produced instead.

- Fastest route: after pasting normally, press **Ctrl** to open the Paste Options button and then **V** (Windows), or use
  **Home → Paste ▾ → Values** (the clipboard icon with **123**).
- In the dialog: **Ctrl + Alt + V**, then **V**, then **Enter**.
- Recent Microsoft 365 updates add **Ctrl + Shift + V** (Mac: **⌘ + Shift + V**) to paste values directly. If it does nothing in
  your version, use one of the routes above.

**Transpose** turns a block on its side. The top row becomes the left column. Tick **Transpose** in the Paste Special dialog, or
choose **Home → Paste ▾ → Transpose**. The paste area must not overlap the copied cells, and Transpose isn't offered after a Cut.
If Transpose is grayed out because the copied cells are inside an Excel Table, convert the Table to a range first (Lesson 3.1).
The TRANSPOSE *function* does the same job with a live link to the source. Its result spills into neighboring cells, which
Lesson 4.1 explains.

**Paste Special with an operation** does math on cells that hold typed-in numbers. Suppose every item in 4 West's supply room goes
up 4% on January 1:

1. Type `1.04` in an empty cell and copy it.
2. Select the UnitCost cells on the Supplies sheet (E4:E17).
3. Open Paste Special, choose **Values** and **Multiply**, and click **OK**. Choosing Values keeps the cells' own formatting.

Medium nitrile gloves go from 7.80 to 8.112 (displayed as 8.11) and the 18 Fr Foley tray from 11.62 to 12.0848 (displayed as
12.08). Press Ctrl + Z (Mac: ⌘ + Z) to undo it, because
the practice tasks use the original prices. The same trick converts numbers stored as text into real numbers: copy a cell containing
`1` and Paste Special → **Multiply** onto them.

> ⚠️ Paste Special is only available after **Copy**. After **Cut**, Excel can only do a normal paste.

### 9. Find & Replace

Lesson 1.1 used **Find** (Ctrl + F, Mac: ⌃ + F) and Find All to count matches. **Replace** (Ctrl + H, Mac: ⌃ + H, or
**Edit → Find → Replace**) opens the same dialog on its Replace tab and changes the matches. If the settings below aren't
showing, click **Options >>**:

| Option | What it does |
|---|---|
| **Within** | **Sheet** (the active sheet) or **Workbook** (every sheet) |
| **Search** | By Rows or By Columns. This only changes the order of results |
| **Look in** | **Formulas** searches what was typed, **Values** searches what's displayed. Leave it on Formulas for typed entries |
| **Match case** | `N` no longer matches `n` |
| **Match entire cell contents** (Mac: **Find entire cells only**) | The *whole* cell must equal what you typed, not just contain it |

**What gets searched depends on your selection.** With a single cell selected, Excel searches the whole sheet. With a range selected,
Excel searches only that range, even several separate blocks selected with Ctrl (Mac: ⌘). That's a precise way to limit a
replacement.

**Replace All** finishes with a message such as *"All done. We made 14 replacements."* That number is your check. **Find All** lists
every match with its cell address and shows a count at the bottom such as *"14 cell(s) found"*. On Windows, click a result and
press **Ctrl + A** to select every found cell on the sheet.

> ⚠️ **Short codes need Match entire cell contents.** Find & Replace matches text *inside* cells and ignores capitals unless you
> tell it otherwise. Replacing shift code `N` with `N12` without that option damages everything that contains an n:
>
> | Cell before | After Replace All (option off) | After Replace All (option on) |
> |---|---|---|
> | `N` | `N12` | `N12` |
> | `Jennifer Patterson` | `JeN12N12ifer PattersoN12` | `Jennifer Patterson` |
> | `Mon` | `MoN12` | `Mon` |
>
> If that happens, press **Ctrl + Z** (Mac: **⌘ + Z**) right away. Undo reverses the whole Replace All in one step.

> 💡 **Tip:** The wildcards from Lesson 1.1 work in Replace too. `*` matches any characters and `?` matches exactly one, so
> `Non-Rebreather*` finds both mask sizes. To search for a real asterisk or question mark, put a tilde in front: `~*`. On
> Windows, press **Ctrl + J** in the Find what box to search for the line breaks that Alt + Enter creates.

### 10. Undo, insert, and delete

**Undo** (Ctrl + Z, Mac: ⌘ + Z) steps back through up to 100 recent actions, and **Redo** (Ctrl + Y, Mac: ⌘ + Y) steps forward
again. Click the arrow next to Undo to undo several steps at once. Saving the file doesn't clear the undo list, but closing the
workbook (and running a macro) does.

**Clearing is not deleting.** The **Delete** key (Mac: **delete**) clears what's *in* the selected cells and leaves the empty cells in
place. To remove rows or columns themselves:

| Task | Windows | Mac |
|---|---|---|
| Select the entire row / column | Shift + Space / Ctrl + Space | Shift + Space / ⌃ + Space (or click the column letter) |
| Insert rows or columns | Select whole rows or columns, then Ctrl + Shift + + | Control-click the row numbers or column letters → **Insert** |
| Delete rows or columns | Select whole rows or columns, then Ctrl + - | Control-click the row numbers or column letters → **Delete** |

**Home → Insert** and **Home → Delete** do the same on both platforms. Excel inserts as many rows as you selected, above the
selection, and inserts columns to the left.

> ⚠️ **Delete whole rows, not cells.** If you select a few cells (not entire rows) and delete them, Excel asks whether to shift the
> cells below *up*. Choosing that in one column slides every value below it out of line with its own record, so a medication ends up
> next to the wrong patient. Select the entire row first.

### 11. Excel's auto-conversion traps

Excel tries to be helpful by converting what you type. With healthcare identifiers, that help can destroy information. Each trap
below happens the moment you press Enter.

| You type | Excel stores | What you meant | Prevent it |
|---|---|---|---|
| `00412345` (an MRN) | 412345 | an 8-digit ID | `'00412345` or Text format first |
| `3-12` (pod 3, bed 12) | March 12 of the current year | a location code | `'3-12` |
| `1/2` (a half-tab dose) | January 2 | 0.5 | `0 1/2` (zero, space, fraction) or `0.5` |
| `MARCH1`, `SEPT2` (gene names) | 1-Mar, 2-Sep | gene symbols | `'MARCH1`. In 2020 the official names became MARCHF1 and SEPTIN2, partly because of Excel |
| `1E5` (a lot number) | 100000, shown as `1.00E+05` | a lot code | `'1E5` |
| `4410287366519087` (16-digit member ID) | 4410287366519080, shown as `4.41029E+15` | the exact ID | `'4410287366519087`. Excel keeps only 15 significant digits, so the last digit is lost for good |
| `123456789012` (12 digits) | the full number, but it's shown as `1.23457E+11` | the full number on screen | Apply the Number format (Lesson 1.3), or store it as text if it's an ID |
| `8/28/26` (a date of birth in 1926) | 08/28/2026 | 1926 | type all four digits of the year |

> 📋 **Version note:** Microsoft 365 (2023 and later) lets you switch some of these conversions off under **File → Options → Data →
> Automatic Data Conversion**: removing leading zeros, truncating long numbers, converting digits around an "E", and turning
> letter-number combinations like MARCH1 into dates. On a Mac, look under **Excel → Settings → Edit**. The settings apply to
> your copy of Excel only, so a colleague opening the same file can still hit the traps. The apostrophe works everywhere.

> 💡 **Tip:** Opening a CSV file by double-clicking it applies the same conversions to every column. Lesson 4.3 shows how to
> import a CSV with Power Query and set each column's type (Text for MRNs, for example) before anything is converted.

### 12. Worked example: a Q4h vital-signs log in under a minute

*Task: 4 West records vital signs every 4 hours (Q4h). Build a log of check times for three days, starting at midnight on
12/16/2025.*

1. In A2, type `12/16/2025 0:00` and press **Enter**. Excel stores one number holding both the date and the time.
2. In A3, type `12/16/2025 4:00` and press **Enter**. The two cells define the 4-hour step.
3. Select **A2:A3** and drag the fill handle down to **A19**.
4. Check the last cell. Three days × 6 checks per day is 18 times, so A2:A19 should end at **12/18/2025 20:00**.
5. Check the count. With A2:A19 selected, the status bar's Count shows **18**.
6. Check the type. Apply General to A3 and you'll see **46007.16667**: day 46,007 (12/16/2025) plus 4 ÷ 24 of a day. Undo to
   restore the date format.

The same log with the Fill Series dialog: select A2, choose **Columns** and **Linear**, Step value `4:00`, and Stop value
`12/18/2025 21:00`.

Whichever method you use, **look at the last value and the count.** A fill that stopped one row early or ran one row too far is
the most common AutoFill mistake, and both checks take two seconds.

## 🧪 Hands-on practice

Download [`1.2-data-entry-autofill.xlsx`](1.2-data-entry-autofill.xlsx) and open the **Practice** sheet. Most tasks send you to
another sheet to do the work, and a gray cell on the Practice sheet reads it. The **Check** column turns green when you're right.

<!-- BEGIN GENERATED: practice -->
Most tasks are done on the other sheets (Entries, Schedule, Beds, Q15 Log, Admissions, Supplies, Order, Roster). Type in the yellow cells here. A gray cell is a pre-filled formula that reads your work on another sheet, and its Check turns green when that work is right. Do the tasks in order, because the Schedule tasks build on each other.

| # | Task | Hint |
|:-:|------|------|
| 1 | On the Entries sheet, select the Entry column (B4:B15). How many of those 12 entries did Excel store as numbers? Dates and times count as numbers. Read the status bar (right-click it to turn on Numerical Count) or use the alignment clue. | By default, numbers line up on the right of a cell and text on the left |
| 2 | Entries!B7 shows the ED arrival as a date and time. What number does Excel actually store in that cell? Switch the cell to General format to see it, then type that number in the yellow cell, rounded to 2 decimal places. | The whole part counts days and the decimal part is the time of day |
| 3 | The MRN on the Entries sheet lost its leading zeros. In the yellow cell, enter the patient's real 8-digit MRN, 00393694, so that Excel keeps it as text with both zeros. | Start the entry with an apostrophe |
| 4 | Type this two-line handoff note into the yellow cell, with "Allergy: Penicillin" on the first line and "Isolation: Contact" on the second line of the same cell. | Enter on its own leaves the cell, so you need a different key combination |
| 5 | On the Schedule sheet, A5 holds 12/01/2025. Use the fill handle to fill the yellow cells A6:A35 with the rest of December, one day per row. The gray cell shows your last date. | Drag the fill handle. A single date counts up one day at a time |
| 6 | On the Schedule sheet, B5 holds "Mon". Fill the weekday names down to B35 by double-clicking the fill handle instead of dragging. The gray cell shows the day name in B35. | Finish task 5 first: double-click fills as far as the neighboring column goes |
| 7 | On the Beds sheet, A4 holds the first bed label, 4W-01. Fill the yellow cells below it so the list runs through all 36 of 4 West's staffed beds. The gray cell shows the label in the last row (A39). | AutoFill increases the number at the end of a text entry |
| 8 | On the Q15 Log sheet, A4 holds 07:00. Fill the yellow cells below it with a check time every 15 minutes, ending at 18:45. Use Home → Fill → Series, or the two-cell AutoFill pattern. The gray cell shows the latest time in column A. | Step value 0:15, with the Stop value a few minutes past the last time |
| 9 | 4 West gets one float-pool RN on every Saturday and Sunday. In the Weekend float column (O) of the Schedule sheet, put FLOAT in every weekend row and nowhere else, using a single entry: Ctrl+click (Mac: ⌘+click) the weekend cells, type FLOAT, and press Ctrl+Enter (Mac: ⌘+Return). The gray cell counts correctly placed FLOATs minus any entries on weekdays. | Finish tasks 5–6 first so you can see which rows are weekends |
| 10 | On the Admissions sheet, use Flash Fill to fill the yellow WhiteboardName column with each patient's first name and last initial, such as "Russell Y." for "Yilmaz, Russell". The gray cell counts how many of the 24 labels are exactly right. | Type an example that can only mean one thing, then press Ctrl + E (Mac: Data → Flash Fill) |
| 11 | The buyer needs 4 West's suggested restock quantities on the Order sheet as plain numbers. Copy Supplies!H4:H17 (SuggestedOrder) and paste only the values into the yellow cells Order!D4:D17. The gray cell totals your Order Qty column. | A normal paste brings the formulas, and their references move |
| 12 | Most people read schedules with staff down the side and dates across the top. Copy Schedule!A4:N35 and use Paste Special → Transpose with the top-left corner in Roster!A3. The gray cell checks that Jean Herrera landed in column A, then counts that nurse's December shifts. | It's a checkbox in the Paste Special dialog |
| 13 | The new scheduling system uses N12 instead of N for a 12-hour night shift. On the Schedule sheet, use Find & Replace to change every N code to N12 without touching any names or day labels. The gray cell counts N12 codes in the grid (C5:N35) and warns you if a name or day label changed too. | Look under Options >> before you click Replace All |
<!-- END GENERATED: practice -->

## ✅ Answer key

The workbook has a hidden **Answer Key** sheet (right-click any sheet tab → **Unhide…** → *Answer Key*). Where a task has a
formula equivalent, its *Live result* column recomputes the answer from the original data. The same answers are below,
collapsed so you don't see them by accident.

<!-- BEGIN GENERATED: answers -->
<details>
<summary><b>🔑 Show the answer key</b> — Try every task before opening this.</summary>

**1. How many entries are stored as numbers?**

- **Answer:** 8
- **Solution:**

1. Select **Entries!B4:B15**.
2. Right-click the status bar and tick **Numerical Count** if it isn't showing.
3. Read Numerical Count (Count shows 12, because it counts every non-empty cell).


Two entries surprise most people. The pod-bed `3-12` became a **date** (March 12), so it's a number. The weight was pasted in as **text**, so it isn't a number even though it looks like one. That's why it sits on the left with a green triangle. TRUE is a logical value (centered), and the name and allergy are text. Everything else, including the date of birth, the arrival time, the copay, and the long member ID, is a number.

**2. Stored value of a date and time**

- **Answer:** 46,005.96
- **Solution:** Select **Entries!B7** and press **Ctrl + Shift + ~** (Mac: **⌃ + Shift + ~**), or choose **Home → Number Format → General**. Read the number, then press **Ctrl + Z** (Mac: **⌘ + Z**) to put the date format back.

Excel stores 12/14/2025 as the serial number 46,005 (day 1 is 1/1/1900). The time 23:04 is 1,384 minutes out of 1,440 in a day, which is 0.9611. Add them and you get 46,005.9611. Formatting only changes how that one number is displayed.

**3. Enter an MRN with leading zeros**

- **Answer:** 00393694
- **Solution:** Type **'00393694** (start with an apostrophe) and press **Enter**. The apostrophe tells Excel "this is text" and is not shown in the cell.

Typed plainly, `00393694` becomes the number 393694, because numbers don't have leading zeros. An apostrophe (or formatting the cell as **Text** *before* you type) stores the characters exactly as typed. A custom number format like `00000000` only *displays* zeros on top of the number, so the stored value would still be wrong for matching and lookups. (The live formula in the key rebuilds the text with TEXT, which you'll meet in Lesson 2.2.)

**4. A two-line note with a line break**

- **Answer:** Allergy: Penicillin ⏎ Isolation: Contact (two lines in one cell)
- **Solution:** Type `Allergy: Penicillin`, press **Alt + Enter** (Mac: **⌃ + ⌥ + Return**), type `Isolation: Contact`, then press **Enter**.

Alt + Enter inserts a line-break character inside the cell and turns on Wrap Text for you. Pressing Enter on its own would finish the entry and jump to the next cell. The check ignores capitals and spaces, but it needs the line break.

**5. AutoFill the December dates**

- **Answer:** 12/31/2025
- **Solution:** Select **A5**, point at the fill handle (the small square at its bottom-right corner) until the pointer becomes a thin **+**, and drag down to **A35**.

AutoFill recognizes a date and adds one day per cell. If every cell shows 12/01/2025 instead, you held Ctrl while dragging (which copies) or chose Copy Cells from the Auto Fill Options button.

**6. AutoFill weekday names (double-click)**

- **Answer:** Wed
- **Solution:** Select **B5** and **double-click** its fill handle. Excel fills down as far as the dates in the neighboring column A go.

Weekday names (Mon, Tue… and Monday, Tuesday…) are a built-in **custom list**, so AutoFill cycles through them. Double-clicking the fill handle copies down to the last row of the adjacent column's data, which is why column A had to be filled first.

**7. AutoFill an ID series (4W-01 …)**

- **Answer:** 4W-36
- **Solution:** Select **Beds!A4** and double-click the fill handle (column B is full, so Excel fills to row 39), or drag it down to **A39**.

When an entry is text that ends in a number, AutoFill increases that number and keeps the rest of the text, including the leading zero (4W-01, 4W-02 … 4W-36). Text with no number in it is just copied.

**8. Fill Series: Q15 check times from 07:00 to 18:45**

- **Answer:** 18:45 (h:mm)
- **Solution:**

**Option A (Fill Series):** select **A4**, choose **Home → Fill → Series…**, pick **Columns** and **Linear**, type **0:15** as the Step value and **18:50** as the Stop value, then click **OK**.

**Option B (AutoFill):** type **7:15** in A5, select A4:A5, and drag the fill handle down to **A51** (18:45).


Times are fractions of a day, so 15 minutes is 0:15 (0.0104…). The Fill Series dialog adds that step until it reaches the Stop value. A stop of 18:50 rather than 18:45 protects you from tiny rounding errors that can drop the last time. With two starting cells, AutoFill copies the gap between them. Either way you get 48 check times. A single time dragged on its own steps by a whole **hour**, not 15 minutes.

**9. Ctrl+Enter into a non-adjacent selection**

- **Answer:** 8
- **Solution:**

1. Click the first weekend cell in column O (the first Sat row), then **Ctrl+click** (Mac: **⌘+click**) every other Sat and Sun row.
2. Type `FLOAT` (it appears in the last cell you clicked).
3. Press **Ctrl + Enter** (Mac: **⌘ + Return**) to enter it in every selected cell at once.


Ctrl + Enter puts the same entry into every selected cell, even when the cells aren't next to each other. Plain Enter would fill only the active cell. The gray formula subtracts any FLOAT typed on a weekday, so it only reaches the full count when every weekend row is marked and no weekday is.

**10. Flash Fill whiteboard names (First L.)**

- **Answer:** 24
- **Solution:**

1. In **Admissions!D2**, type `Rebecca R.` and press **Enter**.
2. In **D3**, type `Russell Y.` and press **Enter**. If Excel shows a gray preview of the remaining labels while you type, you can press Enter to accept it.
3. Otherwise press **Ctrl + E** (or choose **Data → Flash Fill**, on both Windows and Mac).
4. Scan the results, and correct any row Flash Fill got wrong.


Flash Fill studies your examples, finds the pattern ("the text after the comma, a space, the first letter, a period"), and applies it to every row. The first row needs help. In `Richardson, Rebecca` both names start with R, so `Rebecca R.` doesn't show *which* initial you want, and Flash Fill could turn `Yilmaz, Russell` into `Russell R.`. The second example, `Russell Y.`, settles it. The results are typed-in values, not formulas, so they won't update if a name changes. A formula (Lesson 2.2) would. Many units use first name plus last initial on hallway whiteboards to protect patient privacy.

**11. Paste Special → Values**

- **Answer:** 2,526
- **Solution:**

1. Select **Supplies!H4:H17** and press **Ctrl + C** (Mac: **⌘ + C**).
2. Click **Order!D4**.
3. Press **Ctrl + Alt + V** (Mac: **⌃ + ⌘ + V**) to open Paste Special, choose **Values**, and click **OK**. Or use **Home → Paste ▾ → Values (123)**.


SuggestedOrder holds formulas such as `=MAX(0,G4-F4)`. A normal paste copies the *formula*, and because its references are relative, it ends up pointing at the wrong cells on the Order sheet, so you see #VALUE! or wrong numbers. Paste Special → Values pastes only the *results*, which is what you want when the numbers must stay fixed or leave the workbook.

**12. Paste Special → Transpose**

- **Answer:** 14
- **Solution:**

1. Select **Schedule!A4:N35** and press **Ctrl + C** (Mac: **⌘ + C**).
2. Click **Roster!A3**.
3. Press **Ctrl + Alt + V** (Mac: **⌃ + ⌘ + V**), tick **Transpose**, and click **OK**. Or use **Home → Paste ▾ → Transpose**.


Transpose turns the copied block on its side: row 4 (the names) becomes column A, and each date row becomes a column. Jean Herrera's column of codes becomes row 11. Transpose isn't available after **Cut**, and the paste area must not overlap the copied cells.

**13. Find & Replace with Match entire cell contents**

- **Answer:** 61
- **Solution:**

1. Click any single cell on the Schedule sheet (so Excel searches the whole sheet).
2. Press **Ctrl + H** (Mac: **⌃ + H**). Find what: `N`, Replace with: `N12`.
3. Tick **Match entire cell contents** (click **Options >>** first if you can't see it). On a Mac the box is called **Find entire cells only**.
4. Click **Replace All**. Excel reports how many replacements it made, which matches the gray cell.


Without **Match entire cell contents**, Excel replaces the letter n *anywhere* in a cell, and it ignores capitals, so "Mon" would become "MoN12" and every name with an n would be damaged. With the option ticked, only cells containing exactly N change. Selecting just the code grid before Replace All is another safe approach, because Excel then searches only the selection.

</details>
<!-- END GENERATED: answers -->

## 🏆 Bonus challenge

<!-- BEGIN GENERATED: bonus -->
Julie Osei, MD, at the Primary Care Clinic in the Bluestone Outpatient Pavilion, needs an appointment template for the two weeks of Monday 01/05/2026 through Friday 01/16/2026. The clinic books 15-minute slots from 08:00 to 16:45 on weekdays only. Lunch (12:00–12:45) is blocked every day, Wednesday afternoons (13:00 onward) are admin time, and Friday afternoons become video visits. Build the whole grid on the Clinic Grid sheet with AutoFill, Fill Series, Ctrl+Enter, and Find & Replace, without typing cell by cell. B3 (01/05/2026) and A4 (08:00) are filled in for you. The first four parts are checked by gray cells that read the Clinic Grid, so don't type over them. The last part has a yellow cell for your answer.

Work on the **Bonus** sheet of the workbook.

- **B1.** Fill the date header C3:K3 with the next nine weekdays, skipping Saturdays and Sundays. The gray cell shows the date in K3. *(Hint: Date unit: Weekday (or Fill Weekdays))*
- **B2.** Fill the time column A5:A39 with 15-minute slots after 08:00, ending at 16:45. The gray cell shows the latest time in column A. *(Hint: Same technique as the Q15 Log, with a different stop)*
- **B3.** Fill the whole grid B4:K39 with Open in one entry. Then type Lunch over the 12:00–12:45 rows for every day, and Admin over both Wednesday afternoons (13:00–16:45). Use one Ctrl+Enter for each step. The gray cell counts Lunch and Admin slots in the right places, minus any in the wrong place. *(Hint: Holding Ctrl (Mac: ⌘) while you drag adds a second block to the selection)*
- **B4.** Select both Friday-afternoon blocks (13:00–16:45), then use Find & Replace to change Open to Telehealth inside that selection only. Excel reports how many replacements it made. The gray cell counts Telehealth slots in the right places, minus any in the wrong place, so the two numbers should agree. *(Hint: With several cells selected, Replace All stays inside the selection)*
- **B5.** How many slots are still Open for in-person booking? Work it out from the size of the grid and the slots you blocked or converted, then confirm it on the Clinic Grid with Find All. Type the number in the yellow cell. *(Hint: Total slots minus everything you blocked or converted)*
<!-- END GENERATED: bonus -->

<!-- BEGIN GENERATED: bonus-answers -->
<details>
<summary><b>🔑 Show the bonus solution</b> — Give it a real try first!</summary>

**B1. Weekday-only date header**

- **Answer:** 01/16/2026
- **Solution:** Select **B3:K3**, then choose **Home → Fill → Series…**: Rows, Type **Date**, Date unit **Weekday**, Step 1, and click **OK**. Or drag B3's fill handle to K3 and choose **Auto Fill Options → Fill Weekdays**.

Ten weekdays from Monday 01/05 end on Friday 01/16. If you see 01/14 in the last cell, the series included the weekend of 01/10–01/11. (The live formula in the key uses WORKDAY, a date function from Lesson 2.3.)

**B2. 15-minute slot times**

- **Answer:** 16:45 (h:mm)
- **Solution:** Select **A4**, choose **Home → Fill → Series…**: Columns, Linear, Step **0:15**, Stop **16:50**, and click **OK**. Or type 8:15 in A5, select both cells, and drag the fill handle to **A39**.

08:00 to 16:45 every 15 minutes is 36 slots per day, so the grid has 36 × 10 = 360 cells.

**B3. Block lunch and admin time with Ctrl+Enter**

- **Answer:** 72
- **Solution:**

1. Select **B4:K39**, type `Open`, and press **Ctrl + Enter** (Mac: **⌘ + Return**).
2. Select **B20:K23** (12:00–12:45), type `Lunch`, and press **Ctrl + Enter**.
3. Select **D24:D39**, Ctrl+drag (Mac: ⌘+drag) **I24:I39** (the two Wednesday afternoons), type `Admin`, and press **Ctrl + Enter**.


Lunch is 4 slots × 10 days = 40, and Admin is 16 afternoon slots × 2 Wednesdays = 32, so 72 slots are blocked. Typing over a selection with Ctrl + Enter replaces whatever was there, which is why you can paint Open everywhere first and then overwrite the exceptions.

**B4. Find & Replace inside a selection**

- **Answer:** 32
- **Solution:**

1. Select **F24:F39**, then Ctrl+drag (Mac: ⌘+drag) **K24:K39**.
2. Press **Ctrl + H** (Mac: **⌃ + H**). Find what: `Open`, Replace with: `Telehealth`. Tick **Match entire cell contents** for safety.
3. Click **Replace All**. Because more than one cell is selected, Excel searches only the selection. (You can also do one Friday at a time and add the two counts.)


Two Fridays × 16 afternoon slots = 32. If you had only one cell selected, Excel would have replaced every Open on the sheet, so the whole template would have turned into Telehealth.

**B5. Open slots remaining**

- **Answer:** 256
- **Solution:** Work it out: 36 rows × 10 days = 360 slots, minus 40 Lunch, 32 Admin, and 32 Telehealth = **256**. To confirm, click one cell on the Clinic Grid, press **Ctrl + F** (Mac: **⌃ + F**), type `Open`, tick **Match entire cell contents**, and click **Find All**. The dialog reports *"256 cell(s) found"*.

Lunch takes 4 slots on each of 10 days (40). Admin and Telehealth each take the 16 afternoon slots on 2 days (32 and 32). That leaves 360 − 40 − 32 − 32 = 256 Open slots. If Find All reports more, part of a block is missing. If it reports fewer, something extra was overwritten. Working the number out first and then counting is a quick way to catch a fill that went wrong.

</details>
<!-- END GENERATED: bonus-answers -->

## Key takeaways

- Excel stores every entry as text, a number, TRUE/FALSE, or an error. Dates and times are numbers too: a date is a day count and
  a time is a fraction of a day. The alignment clue and the formula bar show you which type you have.
- Store identifiers such as MRNs, ZIP codes, and member IDs as **text** (apostrophe or Text format *before* typing), so leading
  zeros and long digit strings survive.
- **Ctrl + Enter** (Mac: ⌘ + Return) fills every selected cell at once, even cells that aren't next to each other. **Ctrl + D**
  (Mac: ⌘ + D) copies down, and **Alt + Enter** (Mac: ⌃ + ⌥ + Return) adds a line break inside a cell.
- AutoFill continues dates, times, weekday names, and text-with-numbers, and copies everything else. Two starting cells set the
  step. **Fill Series** gives exact steps, weekday-only dates, and stop values.
- **Flash Fill** (Ctrl + E, or **Data → Flash Fill**) learns a text pattern from your examples. Check its results, because they are values that won't update.
- **Paste Special → Values** freezes results, **Transpose** swaps rows and columns, and **Multiply** changes many numbers at once.
- In **Find & Replace**, tick **Match entire cell contents** for short codes, and select a range first when only part of the sheet
  should change.

<!-- BEGIN GENERATED: nav -->
---

⬅️ **Previous:** [1.1 The Excel Interface & Navigation](../01-excel-interface-navigation/README.md) · 🏠 [Course home](../../README.md) · **Next:** [1.3 Formatting Cells & Number Formats](../03-formatting-cells/README.md) ➡️
<!-- END GENERATED: nav -->
