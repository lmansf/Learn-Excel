# Lesson 1.5 · Relative, Absolute & Mixed References

> **Level:** Beginner · **Time:** about 85 minutes · **Workbook:** [`1.5-cell-references.xlsx`](1.5-cell-references.xlsx)
> **Data:** Bluestone Memorial Hospital's 2025 operating expenses by department and category, the same breakdown for October, November, and December 2025, December 2025 overtime for the hourly staff of Medical-Surgical 4 West, and a 4 West staffing grid.

A hospital finance analyst doesn't write 112 formulas for 16 departments and 7 expense categories. The analyst writes one formula
and copies it. A nurse manager builds a staffing grid the same way, and a CFO changes a single inflation assumption and watches a whole
budget update. All of that depends on one skill: knowing what a formula's references do when you copy it. Get it right and one
formula does the work of hundreds. Get it wrong, often by one missing `$` sign, and Excel fills a column with `#DIV/0!` or, worse,
with numbers that look plausible and are wrong. In this lesson you'll learn to predict exactly how a copied formula changes, lock
the parts that must not move, and pull numbers from other sheets, including the same cell on several sheets at once.

## What you'll learn

- Predict how relative references change when a formula is copied
- Lock references with \$ (absolute) and use F4 to toggle
- Build two-way grids with mixed references (\$A1 and A\$1)
- Reference other sheets and sum the same cell across sheets (3-D references)

## 📖 Guide

The examples use the lesson workbook. The **Oct**, **Nov**, **Dec**, and **Expenses** sheets share one layout: departments in rows
5–20, the seven expense categories in columns C–I, a **Total** column J, and a **Hospital total** row 21. Open the Oct sheet and
follow along.

### 1. Relative references: what Excel really remembers

Click **Oct!J5**. The formula bar shows `=SUM(C5:I5)`, the Emergency Department's October total of 1,435,434. You read that as "add
C5 through I5." Excel stores something different: "add the seven cells to my left, on my row." A reference stored this way is a
**relative reference**, and it's the default for every reference you type.

That's why copying works. When you copy J5 down, each copy keeps the same *directions*, so each one adds its own row:

| Formula copied to | Becomes | Adds |
|---|---|---|
| J6 (1 row down) | `=SUM(C6:I6)` | 4 West's categories |
| J7 (2 rows down) | `=SUM(C7:I7)` | 5 East's categories |
| J20 (15 rows down) | `=SUM(C20:I20)` | Pharmacy's categories |

The same happens sideways. **Oct!C21** holds `=SUM(C5:C20)`, the hospital's October Salaries & Wages. Copy it one column right and it
becomes `=SUM(D5:D20)`, Employee Benefits.

The rule for predicting any copy is mechanical:

1. Count how far the formula moves: *r* rows down and *c* columns right (up and left count as negative).
2. Add *r* to every row number in the formula.
3. Shift every column letter *c* letters along the alphabet.

For example, suppose K6 holds `=J6-J5`. Copied to M10, it moves 4 rows down and 2 columns right. Row 6 becomes 10, row 5
becomes 9, and column J becomes L, so M10 holds `=L10-L9`.

All the usual ways of copying a formula follow this rule: dragging or double-clicking the **fill handle**, **Ctrl + D** to fill
down (Mac: **⌘ + D**), **Ctrl + R** to fill right (Mac: **⌘ + R**), and **Ctrl + C** then **Ctrl + V** (Mac: **⌘ + C**, **⌘ + V**).

> 📋 **Cutting is different from copying.** Cut and paste (**Ctrl + X**, then **Ctrl + V**; Mac: **⌘ + X**, **⌘ + V**) *moves* a
> formula, and a moved formula keeps its references exactly as they were. Excel also updates any formula that points at a cell you
> move. To copy a formula's text without adjusting it, select the text in the formula bar, copy it, press **Esc**, and paste it into
> the other cell's formula bar.

> ⚠️ A copied reference can't point above row 1 or left of column A. If a copy would push it off the sheet, Excel writes `#REF!`
> instead. For example, `=J5-J4` copied from row 5 up to row 1 would need row 0.

### 2. When copying breaks a formula: the % of total problem

Relative references are right most of the time. They go wrong when every copy must point at the **same** cell. The classic case is
"% of total": each department's total divided by the hospital total in J21.

Suppose you type `=J5/J21` in **Oct!K5**, the ED's share of October's spending, and copy it down:

| Cell | Formula after copying | Result |
|---|---|---|
| K5 | `=J5/J21` | 11.3% ✔ |
| K6 | `=J6/J22` | #DIV/0! |
| K7 | `=J7/J23` | #DIV/0! |

The numerator should move, because each row needs its own department. The denominator shouldn't. It slid down to J22, an empty
cell, and dividing by an empty cell gives `#DIV/0!`. (If something had been sitting in J22, you'd get a wrong number and no error at
all, which is worse.)

The fix is to lock the denominator with dollar signs, as in `=J5/$J$21`:

| Cell | Formula after copying | Result |
|---|---|---|
| K5 | `=J5/$J$21` | 11.3% |
| K6 | `=J6/$J$21` | 7.9% |
| K7 | `=J7/$J$21` | 7.6% |

> 💡 **Tip:** A % of total column always adds up to 100%. Put `=SUM(K5:K20)` under it as a free sanity check.

### 3. Absolute and mixed references: the \$ sign

A `$` in a reference means "don't change the next thing when this formula is copied." A reference can lock the column, the row,
both, or neither:

| Reference | Name | Copied down a row | Copied right a column | Use it for |
|---|---|---|---|---|
| `J21` | **relative** | `J22` | `K21` | Row-by-row data, like each department's own total |
| `$J$21` | **absolute** | `$J$21` | `$J$21` | One fixed cell, like a grand total or a rate |
| `J$21` | **mixed** (row locked) | `J$21` | `K$21` | A row of values across the top, like targets per column |
| `$J21` | **mixed** (column locked) | `$J22` | `$J21` | A column of values down the side, like labels per row |

The `$` before the letter pins the column. The `$` before the number pins the row. The dollar sign has nothing to do with money.

The `$` signs only matter when a formula is **copied**. A formula you type once and never copy works the same with or without
them, and the result in the original cell never changes.

### 4. F4: add the \$ signs for you

You can type the `$` signs, but **F4** (Mac: **⌘ + T**, or **Fn + F4**) is faster. While you're typing or editing a formula, put the
cursor in or just after a reference and press it. Each press cycles to the next form:

```
J21  →  $J$21  →  J$21  →  $J21  →  J21 (back to the start)
```

1. Type `=J5/J21`, but don't press Enter yet. The cursor is right after J21.
2. Press **F4** once. The formula becomes `=J5/$J$21`.
3. Press **Enter**.

A few details help:

- F4 changes only the reference the cursor is in. To change several at once, select them all in the formula bar first, and F4
  cycles every selected reference together.
- A range counts as one reference. With the cursor in `C5:C20`, F4 gives `$C$5:$C$20`.
- To fix a formula that's already in a cell, press **F2** (Mac: **Control + U**) to edit it, click into the reference, and press F4.

> ⚠️ On Windows, F4 outside a formula means **Repeat last action**. If you press it while you're not editing, Excel may re-apply
> your last format, insert, or delete. Press **Ctrl + Z** (Mac: **⌘ + Z**) to undo it.

> 💡 **Tip:** On many laptops the F-keys control volume or brightness. Hold **Fn** while you press F4, or use **⌘ + T** on a Mac.

### 5. The rate-in-one-cell pattern

Healthcare models are full of **assumptions**: an overtime multiplier, a benefits rate, an inflation rate, or a staffing target
such as **HPPD** (nursing hours per patient day). Put each assumption in its **own labeled cell** and point every formula at it
with an absolute reference. This is the **rate-in-one-cell pattern**.

You can try an example in the empty columns to the right of the Expenses sheet. Finance estimates each department's benefits
cost as 28% of its salaries, which are in column C. Type the label *Benefits rate* in **M3** and the rate, 28%, in **N3**. Then
type one formula in **M5** and copy it down to row 20:

```
M5:  =C5*$N$3        ← C5 moves with the row, but $N$3 stays put
M6:  =C6*$N$3
M7:  =C7*$N$3
```

The estimates land close to the actual Employee Benefits in column D, because Bluestone Memorial's benefits run at about 28% of
salaries overall. Now compare the formula with typing the rate into every cell, as in `=C5*28%`. Both give the same answer today.
But when the rate changes, the first version needs one edit and the second needs 16, one for each department. Numbers typed
inside formulas are called **hard-coded** values, and they're one of the most common causes of stale, wrong spreadsheets. They're
also invisible, because nobody reviewing the sheet can see which rate it assumes.

The pattern also makes **what-if** questions fast. Change the assumption cell, read the new results, then change it back. To
compare before and after, first copy the result you care about and paste it into a spare cell with **Paste Special → Values**
(Lesson 1.2), so the old number stays put while the formulas recalculate. (Lesson 3.6 automates what-if analysis with Goal Seek
and Data Tables.)

The **4 West OT** sheet is set up for this pattern. Cell **B3** holds the overtime multiplier, 1.5 (time-and-a-half), and
practice tasks 5–7 have you point formulas at it.

> 📋 **Rates that grow a number.** To raise a value by a percentage, multiply it by 1 plus the rate. A 3.5% price increase
> turns 100,000 into `=100000*(1+3.5%)`, which is 103,500. To go the other way and find the percent change from an old value to
> a new one, divide and subtract 1: `=103500/100000-1` gives 3.5%.

> 💡 **Tip:** Make assumption cells look different from calculations. Finance teams often use **blue font** for inputs and black for
> formulas, which is the convention in this workbook (see cell B3 on the 4 West OT sheet, and row 3 and column B of Plan 2026).
> Put a label next to every input.

> 📋 **Preview: naming a cell.** Click cell **B3** on the 4 West OT sheet, click the **Name Box** (left of the formula bar), type
> `OT_Multiplier`, and press **Enter**. Any formula in the workbook can now use `OT_Multiplier` instead of `'4 West OT'!$B$3`, and
> the name behaves like an absolute reference. Names can't contain spaces or look like a cell address (`OT1` is a cell, so it
> can't be a name). Lesson 3.1 covers named ranges in depth.

### 6. Mixed references: one formula for a whole grid

A **two-way grid** has labels down the side, labels across the top, and a calculation in every cell that combines its row label with
its column label. Mixed references let one formula fill the whole grid.

Here's a small example: expected average census for three unit sizes at three occupancy targets. Staffed beds run down column A and
targets run across row 4, so every cell is beds × target:

|   | A | B | C | D |
|---|---|---|---|---|
| **4** | *Beds ↓ Target →* | 80% | 85% | 90% |
| **5** | 20 | 16.0 | 17.0 | 18.0 |
| **6** | 28 | 22.4 | 23.8 | 25.2 |
| **7** | 36 | 28.8 | 30.6 | 32.4 |

Type one formula in B5 and copy it across and down:

```
=$A5*B$4
```

- `$A5`: the beds are always in **column A**, so lock the column. The row must change, so each row reads its own bed count.
- `B$4`: the targets are always in **row 4**, so lock the row. The column must change, so each column reads its own target.

With Show Formulas turned on, the grid reads like this. Every cell multiplies its own row label by its own column label:

|   | B | C | D |
|---|---|---|---|
| **5** | `=$A5*B$4` | `=$A5*C$4` | `=$A5*D$4` |
| **6** | `=$A6*B$4` | `=$A6*C$4` | `=$A6*D$4` |
| **7** | `=$A7*B$4` | `=$A7*C$4` | `=$A7*D$4` |

Only one combination of `$` signs works. Here's what C6 becomes under each version of the B5 formula:

| Formula in B5 | Copied to C6 | What goes wrong |
|---|---|---|
| `=A5*B4` | `=B6*C5` | Multiplies two neighboring *results*, not the labels |
| `=$A$5*$B$4` | `=$A$5*$B$4` | Every cell repeats 16.0 |
| `=A$5*$B4` | `=B$5*$B5` | Locks the wrong parts and multiplies grid cells together |
| `=$A5*B$4` | `=$A6*C$4` | ✔ 28 beds × 85% = 23.8 |

To decide where the `$` goes in any formula, ask two questions about each reference: *when I copy across, should it move?* and
*when I copy down, should it move?*

| The reference points at… | Move when copied across? | Move when copied down? | Write it like |
|---|:-:|:-:|---|
| The matching data cell (different for every grid cell) | yes | yes | `C5` |
| One fixed cell (a grand total, a rate) | no | no | `$J$21` |
| A row of labels across the top | yes | no | `C$4` |
| A column of labels down the side | no | yes | `$A5` |

Mixed references are useful outside grids too. Try this in the empty columns of the Oct sheet: type `=C5/$J5` in **L5**, then
copy it across to R5 and down to row 20, and format the block as a percentage. Each cell shows a category's share of its *own
department's* total, because `$J5` always reads the Total column on the current row. The ED spent 59.9% of its October total on Salaries & Wages. With `=C5/C$21` instead,
each cell shows a department's share of its *category's* hospital total, because `C$21` always reads the Hospital total row. The ED
accounts for 12.3% of the hospital's October Salaries & Wages.

**Three ways to fill a grid from one formula:**

1. Type the formula in the top-left cell, drag the fill handle across the first row, then drag the whole row down.
2. Type it in the top-left cell, copy it (Ctrl + C; Mac: ⌘ + C), select the whole grid, and paste.
3. Select the whole grid by dragging from its top-left cell, so that cell is the active cell. Type the formula for that cell and
   press **Ctrl + Enter** (Mac: **⌘ + Return**). Excel enters it in every selected cell and adjusts the references exactly as if
   you had copied it.

> ⚠️ **Double-clicking the fill handle fills down as far as the neighboring column has data.** On the Expenses sheet, column J runs
> through the Hospital total in row 21, so double-clicking K5 fills K21 too. Drag the fill handle when you want to stop at a
> specific row.

### 7. Referencing other sheets

To use a cell from another sheet, put the sheet name and an exclamation mark in front of the address:

```
=Expenses!J5                    one cell on the Expenses sheet
=SUM(Oct!C5:C20)                a range on the Oct sheet
='4 West OT'!B3                 a sheet name with spaces needs single quotes
=Oct!J21-Nov!J21                mixing sheets in one formula
```

The single quotes are required when a sheet name contains a space or a special character such as `-` or `&`, starts with a
digit, or looks like a cell address. Quotes never hurt, so `='Oct'!C5` works too. If a sheet name contains an apostrophe, double
it: `='Women''s Health'!B5`.

You rarely need to type any of this. Build the reference by clicking:

1. Type `=` in the cell where you want the result.
2. Click the other sheet's tab, then click the cell. Excel writes the sheet name, the quotes, and the `!` for you.
3. If you need more cells from that sheet, type an operator such as `/` and click the next cell.
4. Press **Enter**. Excel takes you back to the sheet you started on.

> ⚠️ Press **Enter** while the other sheet is still showing. Clicking your own sheet's tab doesn't finish the formula. Excel is
> still building it, so the next cell you click goes into the formula instead of ending it.

A reference without a sheet name always means the formula's own sheet. So when a formula mixes a cell on another sheet with
cells on its own sheet, click the other sheet's cell and **type** the addresses on your own sheet. For example, on the Expenses
sheet, `=Oct!J5/J5` divides the ED's October total (on Oct) by its 2025 total (on Expenses). If you click back on your own tab to
pick a cell instead, Excel writes the sheet name in front of it, as in `=Oct!J5/Expenses!J5`. That works the same and copies the
same way. It's just longer to read.

Cross-sheet references follow the same copying rules as any other reference. `=Expenses!C5` copied one column right becomes
`=Expenses!D5`, and `=Expenses!$J$21` doesn't move at all. If you rename a sheet, Excel updates every formula that points at it. If
you delete a sheet, formulas that pointed at it show `#REF!`.

### 8. 3-D references: the same cell on many sheets

Monthly reports often live on identical sheets, one per month. A **3-D reference** reaches through a stack of sheets and uses the
same cell, or range, on each one:

```
=SUM(Oct:Dec!C5)                      C5 on Oct, Nov, and Dec, added together
=SUM(Oct:Dec!C5:I20)                  a whole block on all three sheets
=AVERAGE(Oct:Dec!J21)                 the average of the three monthly grand totals
=SUM('Oct 2025:Dec 2025'!C5)          sheet names with spaces: one pair of quotes around both
```

`Oct:Dec!` means "the Oct tab, the Dec tab, and every tab between them." On this workbook's data:

| Formula | Result | What it answers |
|---|---:|---|
| `=SUM(Oct:Dec!C5)` | 2,618,249 | ED Salaries & Wages for Q4 (859,807 + 866,792 + 891,650) |
| `=AVERAGE(Oct:Dec!J21)` | 12,766,652.33 | Average monthly hospital expense in Q4 |
| `=MAX(Oct:Dec!J9)` | 1,008,007 | ICU's most expensive Q4 month (December) |

To build one by clicking:

1. Type `=SUM(`.
2. Click the **Oct** tab, then hold **Shift** and click the **Dec** tab. All three tabs highlight.
3. Click the cell (or drag the range) you want, type `)`, and press **Enter**.

The cell part of a 3-D reference is an ordinary reference, so it copies the usual way. `=SUM(Oct:Dec!C5)` copied one column right
becomes `=SUM(Oct:Dec!D5)`. That means one 3-D formula can fill a whole quarterly summary that has the same layout as the monthly
sheets.

3-D references work with SUM, AVERAGE, COUNT, COUNTA, MAX, MIN, PRODUCT, the "A" versions AVERAGEA, MAXA, and MINA, and the
standard-deviation and variance functions (STDEV.S, STDEV.P, VAR.S, VAR.P, and their older and "A" versions). COUNTIF, SUMIF, and
the other conditional functions don't accept them and return `#VALUE!`.

**The gotchas.** A 3-D reference is defined by **tab positions**, not by sheet names or months, and it adds cells by **position**,
not by label:

| If you… | The 3-D formula… |
|---|---|
| Insert or drag a sheet between Oct and Dec | Silently **includes** it |
| Drag a middle sheet (Nov) outside the Oct…Dec span | Silently **drops** it |
| Move an end sheet (Oct or Dec) | Recalculates over the new span between the two end tabs |
| Delete an end sheet | Shrinks to the remaining sheets between the end tabs |
| Put the summary sheet *between* the end tabs | Includes itself and triggers a **circular reference** warning |
| Rearrange rows on one monthly sheet | Adds mismatched departments with no error at all |

> 💡 **Tip:** Some analysts add two empty "bookend" sheets, such as **First** and **Last**, and write `=SUM(First:Last!C5)`. Any
> monthly sheet dropped between the bookends is included automatically, and the bookends make the range visible in the tab bar.

> ⚠️ **Grouped sheets.** Shift + clicking tabs while you're *not* typing a formula **groups** the sheets: the title bar shows "Group,"
> and anything you type or delete then happens on every grouped sheet at once. Right-click any tab and choose **Ungroup Sheets**
> before you edit again.

### 9. References to other workbooks (briefly)

A formula can also point into another workbook. The workbook name goes in square brackets, in front of the sheet name:

```
=[Budget2026.xlsx]Plan!$C$5                    while Budget2026.xlsx is open
='[Budget 2026.xlsx]Plan'!$C$5                 a space in the workbook name: quotes around [workbook]sheet
='C:\Finance\[Budget 2026.xlsx]Plan'!$C$5      after it's closed (Excel adds the full path)
```

The quoting rule is the same as for sheet names. If the workbook name, sheet name, or path contains a space or a special
character, wrap everything before the `!` in single quotes. If you build the reference by clicking, Excel adds the quotes for you.

When you build one by clicking into the other workbook, Excel makes the reference absolute (`$C$5`), so remove the `$` signs if
you plan to copy it. Excel calls these connections **links**. When you open a file with links, Excel may show a security bar asking
whether to update them. Manage them with **Data → Workbook Links** in Microsoft 365, or **Data → Edit Links** in older versions.
Links break when the source file is renamed or moved, so for anything important it's safer to copy the data in or use Power Query
(Lesson 4.3).

### 10. Seeing and checking references

You can't see a copying mistake by looking at results. These tools show the formulas themselves:

| Tool | How | What it shows |
|---|---|---|
| **Show Formulas** | **Ctrl + `` ` ``** (Mac: **Control + `` ` ``**), or **Formulas → Show Formulas** | Every formula instead of its result. Press it again to switch back |
| **Edit mode** | **F2** (Mac: **Control + U**) | Each reference in a different color, with a matching colored box around the cells it uses |
| **Trace Precedents** | **Formulas → Trace Precedents** | Arrows from the cells a formula uses. A dashed arrow with a small sheet icon means the source is on another sheet |
| **Trace Dependents** | **Formulas → Trace Dependents** | Arrows to the formulas that use the selected cell. Clear them with **Remove Arrows** |

Show Formulas is the fastest way to audit a copied block. In a correct grid the formulas form a clean pattern: the locked parts
are identical in every cell and the moving parts step by one. A cell that breaks the pattern stands out immediately.

> 📋 **Excel Tables use a different system.** In an Excel Table (Lesson 3.1), formulas use *structured references* such as
> `[@OTHours]`, which point at columns by name and don't use `$` signs. That's why this lesson's data sheets are plain ranges.

### 11. Shortcuts and version notes

| Action | Windows | Mac |
|---|---|---|
| Cycle a reference through its `$` forms while typing or editing a formula | F4 | ⌘ + T (or Fn + F4) |
| Edit the active cell and color-code its references | F2 | Control + U |
| Fill down / fill right | Ctrl + D / Ctrl + R | ⌘ + D / ⌘ + R |
| Copy / paste | Ctrl + C / Ctrl + V | ⌘ + C / ⌘ + V |
| Cut, to move a formula without adjusting its references | Ctrl + X | ⌘ + X |
| Enter the same formula in every selected cell | Ctrl + Enter | ⌘ + Return |
| Show or hide formulas on the whole sheet | Ctrl + `` ` `` | Control + `` ` `` |
| Pick the last sheet of a 3-D range while building the formula | Shift + click its tab | Shift + click its tab |
| Cancel an edit / undo | Esc / Ctrl + Z | Esc / ⌘ + Z |

| Feature | Version |
|---|---|
| Relative, absolute, and mixed references, references to other sheets, and 3-D references | Excel 2016 and later and Microsoft 365, on Windows and Mac |
| **Formulas → Show Formulas**, **Trace Precedents**, and **Trace Dependents** | Every current desktop version, Windows and Mac |
| **Data → Workbook Links** | Microsoft 365. Older versions use **Data → Edit Links** |

## 🧪 Hands-on practice

Download [`1.5-cell-references.xlsx`](1.5-cell-references.xlsx) and open the **Practice** sheet. Type each answer in the yellow cell,
as a formula wherever possible. Tasks with a gray cell are done on another sheet, and the gray cell checks your work there. The
**Check** column turns green when you're right.

<!-- BEGIN GENERATED: practice -->
Tasks 1, 4, and 8 ask you to predict a formula or reference. Type your answer as text without the leading = (for example AVERAGE(B2:B9)), because Excel would calculate it if you typed the =. The other tasks want a formula in the yellow cell, or work on another sheet that a gray cell checks.

| # | Task | Hint |
|:-:|------|------|
| 1 | On the Expenses sheet, cell J5 contains =SUM(C5:I5), the Emergency Department's total. If you copy J5 and paste it into J14, what formula will J14 contain? Type it without the =, then click J14 to check. | Relative references keep their distance from the formula cell |
| 2 | What share of Bluestone Memorial's 2025 operating expense came from the Intensive Care Unit? Divide the ICU's total (Expenses!J9) by the hospital total (Expenses!J21). Enter it as a percentage. | Type =, click the Expenses tab, click the cell, type /, click the other cell, press Enter |
| 3 | On the Expenses sheet, fill the yellow % of Total column (K5:K20) with ONE formula: type it in K5, lock the reference to the hospital total in J21, then drag the fill handle down to row 20. The gray cell adds up your column. If it shows #DIV/0!, the total reference moved when you copied. | Click J21 in the formula, then press F4 (Mac: ⌘ + T) |
| 4 | You type =C21 in a cell, leave the cursor in the reference, and press F4 three times (Mac: ⌘ + T three times). Which reference do you end up with? Type it without the =. | Absolute, then row locked, then column locked… |
| 5 | On the 4 West OT sheet, Jessica Aguilar (row 33) worked 9.10 overtime hours in December. On this Practice sheet, write a formula for this employee's overtime pay: OTHours × HourlyRate × the overtime multiplier in B3. Refer to the cells; don't type the numbers. | The sheet name starts with a digit and contains spaces, so it needs single quotes |
| 6 | On the 4 West OT sheet, fill the yellow OTPay column (H6:H72) with one formula: OTHours × HourlyRate × the overtime multiplier in B3. Type it in H6 and copy it down. The gray cell totals your column: what did December's overtime cost? | Lock B3; let the row references move |
| 7 | Finance asks what December's overtime would have cost at double time. Change '4 West OT'!B3 to 2, read task 6's gray cell, and type that amount here as a number. Then set B3 back to 1.5. | One edit updates the whole column. That's the point of the rate cell |
| 8 | A grid has labels across row 2 and down column A, and cell B3 holds =B\$2*\$A3. If you copy B3 to D6, what formula will D6 contain? Type it without the =. | A \$ freezes only the part right after it |
| 9 | On the Staffing Grid sheet, fill the yellow grid B6:H12 with ONE formula: nursing hours needed per day = census (column A) × HPPD target (row 5). Type it in B6, copy it across to H6, then down to row 12. The gray cell adds up the whole grid. | Lock the column of the census and the row of the HPPD |
| 10 | Using one 3-D reference, what did Oncology (row 11 on the Oct, Nov, and Dec sheets) spend on Pharmaceuticals (column F) in Q4 2025? | Type =SUM(, click the Oct tab, Shift + click the Dec tab, then click the cell |
| 11 | Fill the yellow grid on the Q4 Summary sheet (C5:I20) with ONE 3-D formula that adds the same cell on the Oct, Nov, and Dec sheets. Type it in C5, copy it across to I5, then down to row 20. The gray cell adds up your grid: what was Bluestone Memorial's Q4 operating expense? | 3-D references copy like ordinary relative references |
| 12 | A colleague drags the Nov tab to the right of the Dec tab. What would =SUM(Oct:Dec!J21) return then? (J21 is each month's hospital total.) Answer with a formula that uses ordinary sheet references, not a 3-D reference, so it stays correct with the tabs in their usual order. If you test the move, drag Nov back between Oct and Dec afterwards. | A 3-D range is defined by tab positions, not by month names |
<!-- END GENERATED: practice -->

## ✅ Answer key

The workbook has a hidden **Answer Key** sheet (right-click any sheet tab → **Unhide…** → *Answer Key*). Its *Live result* column
runs each sample formula against the data, so you can see it working. The same answers are below, collapsed so you don't see them
by accident.

<!-- BEGIN GENERATED: answers -->
<details>
<summary><b>🔑 Show the answer key</b> (try every task before you open this)</summary>

**1. Predict: J5 copied to J14**

- **Answer:** `=SUM(C14:I14)`
- **Solution:** Count the move: from row 5 to row 14 is 9 rows down and 0 columns across. Add 9 to every row number in the formula and leave the column letters alone.

Excel stores =SUM(C5:I5) in J5 as "the seven cells to my left, on my row." Pasted into J14, the same instruction points at C14:I14. That's why one total formula works for every department.

**2. ICU share of total expense**

- **Answer:** 7.5%
- **Solution:** `=Expenses!J9/Expenses!J21`

A reference to another sheet is the sheet name, an exclamation mark, then the cell. If you build it by clicking, Excel writes the Expenses! part for you. Press Enter while you're still on the Expenses sheet, and Excel takes you back to Practice.

**3. % of Total column (absolute reference)**

- **Answer:** 100.0%
- **Solution:** `=J5/$J$21`

Without the \$ signs, K6 would hold =J6/J22. Row 22 is empty, so the result is #DIV/0!. \$J\$21 stays put wherever you copy it, while J5 moves to each department's row. =J5/J\$21 also works here, because you only copy down. Every % of total column adds up to 100%, so the gray cell doubles as a sanity check.

**4. F4 pressed three times**

- **Answer:** `$C21`
- **Solution:** F4 cycles C21 → \$C\$21 → C\$21 → \$C21 → back to C21. The third press gives **\$C21**.

\$C21 locks the column (C) but lets the row change. Try it: type the formula, press F4 three times, and watch the formula bar. Press Esc afterwards so you don't leave the test formula in the sheet.

**5. Overtime pay for Jessica Aguilar**

- **Answer:** 559.65
- **Solution:** `='4 West OT'!G33*'4 West OT'!D33*'4 West OT'!B3`

Sheet names that contain spaces or start with a digit must be wrapped in single quotes: '4 West OT'!B3. When you click the cells instead of typing, Excel adds the quotes for you. No \$ signs are needed here, because this formula is never copied.

**6. OTPay column (rate in one cell)**

- **Answer:** 16,029.43
- **Solution:** `=G6*D6*$B$3`

G6 and D6 are relative, so each row uses its own employee's hours and rate. \$B\$3 is absolute, so all 67 rows share the one multiplier cell. If you had typed *1.5 into every formula, a policy change would mean editing 67 formulas.

**7. What-if: double time**

- **Answer:** 21,372.58
- **Solution:**

1. Click **'4 West OT'!B3**, type `2`, and press **Enter**.
2. Read the gray cell in task 6 and type that amount into this task's yellow cell.
3. Change B3 back to `1.5` so tasks 5 and 6 show ✔ again.


Because every OTPay formula points at B3, one edit recalculates all 67 rows and the total. This is the payoff of the rate-in-one-cell pattern: assumptions live in one labeled cell, so what-if questions take seconds.

**8. Predict: B3 copied to D6**

- **Answer:** `=D$2*$A6`
- **Solution:** The move is 2 columns right (B to D) and 3 rows down (row 3 to row 6). In B\$2 the row is locked, so only the column moves and it becomes D\$2. In \$A3 the column is locked, so only the row moves and it becomes \$A6. The answer is **=D\$2*\$A6**.

B\$2 always reads row 2 (the labels across the top) in the current column. \$A3 always reads column A (the labels down the side) in the current row. Together they make every cell multiply its own column label by its own row label, which is exactly what a two-way grid needs. You'll build one in the next task.

**9. Staffing grid (mixed references)**

- **Answer:** 13,328
- **Solution:** `=$A6*B$5`

\$A6 keeps every formula looking at column A, and B\$5 keeps every formula looking at row 5. With =\$A\$6*\$B\$5, every cell would repeat 182, the B6 result. With no \$ at all, C6 would compute =B6*C5, multiplying the previous result by the HPPD instead of using the census. Cross-check: the sum of a multiplication grid equals (sum of the census values) × (sum of the HPPD targets), which is how the key's live formula works.

**10. 3-D SUM: Oncology pharmaceuticals in Q4**

- **Answer:** 654,633
- **Solution:** `=SUM(Oct:Dec!F11)`

Oct:Dec! means every sheet from the Oct tab through the Dec tab, so the formula adds F11 on all three. It's the same as =Oct!F11+Nov!F11+Dec!F11, but it stays short no matter how many monthly sheets sit between the two end tabs.

**11. Q4 Summary built with a 3-D formula**

- **Answer:** 38,299,957
- **Solution:** `=SUM(Oct:Dec!C5)`

The cell part of Oct:Dec!C5 is relative, so copying moves it exactly like a normal reference: I20 ends up as =SUM(Oct:Dec!I20). This only works because the three monthly sheets share one layout. A 3-D reference adds cells by position, not by department name. Q4 Summary sits to the left of Oct, outside the Oct:Dec range, so it never adds itself.

**12. 3-D gotcha: a tab moved out of the range**

- **Answer:** 25,941,245
- **Solution:** `=Oct!J21+Dec!J21`

Oct:Dec! means "Oct, Dec, and whatever tabs sit between them right now." Once Nov is dragged past Dec, it's outside the sandwich, so the total silently drops November. A sheet dragged in between the end tabs gets added just as silently. Keep 3-D sheets in order, and consider empty "bookend" sheets (for example First and Last) that mark the range.

</details>
<!-- END GENERATED: answers -->

## 🏆 Bonus challenge

<!-- BEGIN GENERATED: bonus -->
Finance needs a first-draft 2026 operating expense plan for Bluestone Memorial. The rule: each 2026 cell = the 2025 actual for the same department and category (Expenses sheet) × (1 + that category's price inflation in row 3 of the Plan 2026 sheet) × (1 + that department's volume growth in column B). The Plan 2026 sheet uses the same rows and columns as the Expenses sheet. Column J and row 21 of Plan 2026 already total whatever you put in the grid.

Build the plan on the **Plan 2026** sheet, and type your answers in the yellow cells on the **Bonus** sheet.

- **B1.** Fill the yellow grid on the Plan 2026 sheet (C5:I20) with ONE formula typed in C5 and copied across and down. The gray cell adds up your grid: what is the 2026 plan total? (Don't round inside the formula.) *(Hint: The 2025 actual moves both ways, the inflation row is locked, and the growth column is locked)*
- **B2.** What is the 2026 plan for Laboratory · Medical Supplies (Plan 2026, row 18, column E)? Reference the cell in your grid. *(Hint: Check it by hand: 2025 amount × (1 + inflation) × (1 + growth))*
- **B3.** By what percentage would Bluestone Memorial's total operating expense grow from 2025 to 2026 under this plan? Enter it as a percentage. *(Hint: New ÷ old − 1)*
- **B4.** Which department's total expense is planned to grow by the largest percentage? Type the department name exactly as it appears in column A. *(Hint: Add a helper column that divides each department's 2026 total by its 2025 total)*
- **B5.** Pharmaceutical prices are the shakiest assumption. If pharmaceutical inflation were 9.0% instead of 7.5%, how many dollars higher would the 2026 plan total be? Change the one input cell, compare the totals, and type the difference here as a number, rounded to the nearest dollar. Then put the input back. *(Hint: Because of the \$ signs, one edit flows to all 16 Pharmaceuticals cells)*
<!-- END GENERATED: bonus -->

<!-- BEGIN GENERATED: bonus-answers -->
<details>
<summary><b>🔑 Show the bonus solution</b> (give it a real try first)</summary>

**B1. 2026 plan grid (three kinds of reference in one formula)**

- **Answer:** 162,219,455.28
- **Solution:** `=Expenses!C5*(1+C$3)*(1+$B5)`

Expenses!C5 is fully relative, because each plan cell needs the matching 2025 cell. C\$3 locks the row, so every department reads the inflation in row 3 of its own category column. \$B5 locks the column, so every category reads the growth in column B of its own department row. One formula, three reference types, 112 correct cells.

**B2. Spot-check: Laboratory · Medical Supplies**

- **Answer:** 3,424,829.93
- **Solution:** `='Plan 2026'!E18`

Spot-checking a cell far from where you typed the formula proves the copy worked. E18 should hold =Expenses!E18*(1+E\$3)*(1+\$B18): the inflation is still read from row 3 and the growth from column B.

**B3. Hospital-wide growth, 2025 → 2026**

- **Answer:** 6.8%
- **Solution:** `='Plan 2026'!J21/Expenses!J21-1`

'Plan 2026'!J21 is the 2026 grand total and Expenses!J21 is the 2025 grand total. New ÷ old − 1 turns the two totals into a growth rate.

**B4. Fastest-growing department**

- **Answer:** Oncology
- **Solution:**

1. In 'Plan 2026'!K5, type `=J5/Expenses!J5-1` and copy it down to K20. Both references are relative, so each row compares its own department.
2. Format K5:K20 as a percentage and find the largest value (`=MAX(K5:K20)` helps).
3. Read the department name in column A of that row.


Oncology grows 9.7%, ahead of Observation Unit at 9.5%, even though Observation Unit has the highest volume growth (5.0%). Oncology's spending is heavy in Pharmaceuticals, the category with the steepest inflation (7.5%), so its mix pushes the total up faster.

**B5. Sensitivity: pharmaceutical inflation**

- **Answer:** 176,282
- **Solution:**

1. Keep a copy of the current plan total: copy 'Plan 2026'!J21 and paste it as a value (**Paste Special → Values**) into a spare cell such as 'Plan 2026'!L21.
2. Change 'Plan 2026'!F3 from 7.5% to 9.0%.
3. Subtract the old total from the new one (`=J21-L21` on Plan 2026), round to the nearest dollar, and type that number into this task's yellow cell.
4. Set F3 back to 7.5% so B1 and B3 show ✔ again.


Only the Pharmaceuticals column reads that input, so only those 16 cells change. Each one rises by its 2025 amount × (1 + growth) × 0.015, which is the 1.5-point rise in the rate. Testing the shakiest assumption like this is called a sensitivity check, and it only takes seconds because the assumption lives in one cell.

</details>
<!-- END GENERATED: bonus-answers -->

## Key takeaways

- A relative reference stores *directions* from the formula's cell, so a copied formula keeps its distance: move it *r* rows and
  *c* columns, and every relative reference moves the same way.
- A `$` locks the part that follows it. Use `$J$21` for one fixed cell, `C$4` for a row of labels, and `$A5` for a column of labels.
  Press **F4** (Mac: **⌘ + T**) to cycle through the four forms.
- Put every assumption in its own labeled cell and point formulas at it with an absolute reference. One edit then updates the whole
  model, which makes what-if questions fast.
- One formula with mixed references (`=$A5*B$4`) fills an entire two-way grid. Check it with Show Formulas
  (**Ctrl + `` ` ``**, Mac: **Control + `` ` ``**).
- Other sheets are `Sheet!A1`, with single quotes for names that need them: `'4 West OT'!B3`. Build them by clicking.
- `=SUM(Oct:Dec!C5)` adds the same cell across a stack of identically laid-out sheets, but the stack is defined by tab
  *positions*, so moving a tab can silently change the total.

<!-- BEGIN GENERATED: nav -->
---

⬅️ **Previous:** [1.4 Your First Formulas & Functions](../04-basic-formulas/README.md) · 🏠 [Course home](../../README.md) · **Next:** [1.6 Sorting & Filtering Data](../06-sorting-filtering/README.md) ➡️
<!-- END GENERATED: nav -->
