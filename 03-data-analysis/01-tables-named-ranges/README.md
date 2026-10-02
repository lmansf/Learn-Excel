# Lesson 3.1 · Excel Tables, Structured References & Named Ranges

> **Level:** Intermediate · **Time:** about 50 minutes · **Workbook:** [`3.1-tables-named-ranges.xlsx`](3.1-tables-named-ranges.xlsx)
> **Data:** Supply inventory snapshot for 15 storerooms across three Bluestone hospitals (257 stock rows, as of 12/31/2025), a Settings sheet with the report date and expiry window, and Purchasing's vendor list.

Bluestone's supply chain team gets a fresh inventory snapshot every week. Today it has 257 stock rows, and next Monday it will
have a few more or a few less. A formula like `=SUM(Inventory!J2:J258)` still stops at row 258 when next week's rows are
pasted below it, so it quietly misses them, and nobody notices until a unit runs out of saline flushes. **Excel Tables** fix this because they grow with
the data, and their formulas read like the question you're asking: `=SUM(tblInventory[QtyOnHand])`. **Named ranges** do the
same for the settings a report depends on, such as the report date and the 90-day expiry window. In this lesson you turn a raw
inventory export into a Table, use it to find stockouts, expired stock, and overdue counts, and then build a reorder report
that updates itself.

## What you'll learn

- Convert ranges to Excel Tables and name them
- Write structured references like tblInventory[UnitCost] and [@QtyOnHand]
- Use calculated columns, the Total Row, and slicers on Tables
- Create, manage, and use named ranges and named constants

## 📖 Guide

### 1. Why use a Table?

An **Excel Table** (this lesson just says *Table*) is a range that Excel manages as one object. It knows where its header row
is, where its data starts and ends, and what each column is called. A plain **range** is just cells, and Excel knows nothing
about how they fit together.

| | Plain range | Excel Table |
|---|---|---|
| New rows below the data | Missed by `=SUM(J2:J258)` | Included automatically in `=SUM(tblInventory[QtyOnHand])` |
| Row-by-row formulas | `=J2*G2`, different on every row | `=[@QtyOnHand]*[@UnitCost]`, identical on every row |
| Copying a formula down | You fill it yourself | A **calculated column** fills itself |
| Totals | You write SUBTOTAL formulas | Tick **Total Row** and pick from a dropdown |
| Filtering | Turn on AutoFilter yourself | Filter buttons are built in, and you can add **slicers** |
| Formatting | Manual | Banded rows that stay banded when you sort or add rows |
| PivotTables and charts built on it | You update the source range by hand | The source grows with the Table (Lessons 3.4 and 3.5) |

Use a Table for any list of records, meaning one row per item and one column per attribute: stock rows, encounters, claims,
shifts. Keep plain ranges for small blocks that aren't lists, such as the two-row Settings block in this workbook.

### 2. Turn a range into a Table

1. Click any single cell inside the data. On the Inventory sheet, A2 works.
2. Press **Ctrl + T** (Mac: **Control + T**), or choose **Insert → Table**. On Windows, **Ctrl + L** does the same thing.
3. Excel guesses the range from the block of filled cells around your cell and shows it in the dialog, for example
   `=$A$1:$S$258`. Check it.
4. Make sure **My table has headers** is ticked, then click **OK**.

Excel applies a banded style, adds a filter button to each header, and shows the **Table Design** tab whenever a cell inside
the Table is selected. In Excel 2019 and earlier that tab is called **Table Tools → Design**, and on a Mac it's called **Table**.
**Home → Format as Table** does the same job, but it asks you to pick a style first.

Excel can only guess the right range when the data is a clean list:

- **One header row**, with a unique, non-blank name in every column. Excel renames duplicate headers (QtyOnHand, QtyOnHand2)
  and fills blank ones with Column1, Column2, and so on.
- **No completely blank rows or columns** inside the data. A blank row makes Excel stop the range early.
- **No merged cells.** A Table can't contain them.
- **One record per row**, with no subtotal rows mixed in.

> ⚠️ If **My table has headers** is unticked, Excel inserts a new header row (Column1, Column2…) and treats your real headers
> as data. Press **Ctrl + Z** (Mac: **⌘ + Z**) and convert again.

> 💡 **Tip:** On a sheet without frozen panes, scroll down until the header row is out of view, and Excel replaces the
> column letters (A, B, C…) with the Table's column names as long as the active cell is inside the Table. You always know
> which column you're in. The Inventory sheet freezes its header row instead, so there the headers simply stay on screen.

### 3. Name the Table

Excel names Tables Table1, Table2, and so on. Rename every Table right away, because the name appears in every formula that
uses it.

1. Click any cell in the Table.
2. On the **Table Design** tab (Mac: **Table** tab), click in the **Table Name** box at the far left.
3. Type the new name, such as `tblInventory`, and press **Enter**.

Many analysts start Table names with **tbl**. The prefix tells you at a glance that `tblInventory` is a Table and not a named
cell, and when you type `=tbl`, AutoComplete lists every Table in the workbook. The Name Box, left of the formula bar, lists
Tables too, so you can jump to one by picking its name.

Table names follow the same rules as every other name in Excel (section 10): no spaces, start with a letter, an underscore,
or a backslash, and don't look like a cell address.

> ⚠️ `tbl1` is not a valid name, because TBL1 is a real cell address: Excel has 16,384 columns, and TBL is one of them. Use a
> descriptive name like `tblInventory` instead.

### 4. Structured references

A **structured reference** points to part of a Table by name instead of by address. The two forms you'll use most are:

```
tblInventory[QtyOnHand]     every data cell in the QtyOnHand column
[@QtyOnHand]                QtyOnHand in the formula's own row (used inside the Table)
```

The full syntax is the Table name followed by square brackets that hold a column name, a **special item** such as `#All`, or
both:

| Structured reference | Refers to | Example |
|---|---|---|
| `tblInventory[QtyOnHand]` | The data cells of one column (no header, no Total Row) | `=SUM(tblInventory[QtyOnHand])` returns 75,473 units |
| `tblInventory` or `tblInventory[#Data]` | All data cells in every column | `=ROWS(tblInventory[#Data])` returns 257 |
| `tblInventory[#All]` | Header row, data, and Total Row | The whole Table, edges included |
| `tblInventory[#Headers]` | The header row | `=COUNTA(tblInventory[#Headers])` counts the columns |
| `tblInventory[#Totals]` | The Total Row (only while it's turned on) | |
| `tblInventory[[#Headers],[UnitCost]]` | One header cell | Returns the text UnitCost |
| `tblInventory[[#Totals],[UnitCost]]` | One Total Row cell | Returns that column's total |
| `[@UnitCost]` | UnitCost in the same row as the formula | Used in calculated columns |
| `tblInventory[@UnitCost]` | The same, written from outside the Table | Works only on a row the Table occupies |
| `tblInventory[[QtyOnHand]:[ReorderQty]]` | Several adjacent columns | `=SUM(tblInventory[[QtyOnHand]:[ReorderQty]])` |

> 📋 Excel 2007 wrote "this row" as `tblInventory[[#This Row],[UnitCost]]`. Excel 2010 and later show the shorter
> `[@UnitCost]`. Both mean the same thing.

#### Let Excel write them for you

You rarely type a structured reference in full:

- **AutoComplete.** Type `=SUM(tblInventory[` and Excel lists the columns and special items. Pick one with the arrow keys,
  press **Tab**, then type `]`.
- **Click.** While you write a formula, click a cell or drag over a column inside the Table, and Excel writes the structured
  reference. Inside the Table, clicking a cell in the same row gives `[@ColumnName]`. Clicking a Total Row cell from another
  sheet gives `tblInventory[[#Totals],[ColumnName]]`.

> 💡 **Tip:** If clicking gives you `J2` instead of `[@QtyOnHand]`, the option is switched off. Turn it back on in
> **File → Options → Formulas → Use table names in formulas** (Mac: **Excel → Preferences → Tables & Filters**).

#### Worked examples on the inventory

| Question | Formula | Result |
|---|---|---|
| Units on hand, all locations | `=SUM(tblInventory[QtyOnHand])` | 75,473 |
| Units of PPE on hand | `=SUMIFS(tblInventory[QtyOnHand],tblInventory[Category],"PPE")` | 6,207 |
| Stock rows at Ashby Falls | `=COUNTIFS(tblInventory[Facility],"Ashby Falls Community Hospital")` | 37 |
| Value of ICU stock (after task 3) | `=SUMIFS(tblInventory[ExtendedValue],tblInventory[Location],"Intensive Care Unit")` | 146,255.45 |

You never have to check that the ranges are the same size, because every column of a Table has the same rows. That
satisfies the SUMIFS and COUNTIFS rule from Lesson 2.5 automatically. You don't need a sheet name either, because a Table
name works from any sheet in the workbook.

#### Column names with spaces or symbols

This course uses headers without spaces, such as QtyOnHand. If a header contains a space, the this-row form needs an extra
pair of brackets, as in `[@[Qty On Hand]]`, while the whole-column form stays `tblInventory[Qty On Hand]`. A header with a
symbol such as `$` or `%` needs the extra brackets in every form, as in `tblInventory[[Cost $]]`. A few characters, such as
`[`, `]`, `#`, `'`, and `@`, also need an apostrophe in front of them. Excel adds the brackets and apostrophes for you when
you click, but short headers without spaces or symbols keep every formula easier to read.

> ⚠️ **Dragging sideways changes the column.** If you drag `=SUM(tblInventory[QtyOnHand])` one cell to the right with the
> fill handle, it becomes `=SUM(tblInventory[ParLevel])`, because Excel treats the column name like a relative reference.
> Copy and paste (Ctrl + C, Ctrl + V; Mac: ⌘ + C, ⌘ + V) leaves it alone. To lock a column so it never moves, write it as a
> one-column range: `tblInventory[[QtyOnHand]:[QtyOnHand]]`.

### 5. Calculated columns

A **calculated column** is a Table column in which every cell holds the same formula. You create one by entering a formula in
any cell of an empty Table column:

1. Click R2, the first cell of the yellow ExtendedValue column.
2. Type `=`, click J2 (QtyOnHand), type `*`, and click G2 (UnitCost). Excel writes `=[@QtyOnHand]*[@UnitCost]`.
3. Press **Enter**. Excel fills the formula into every row of the column.

Row 2 holds small nitrile exam gloves in Cedar Ridge's Emergency Department: 527 boxes × $8.22 = 4,331.94. Every other row
gets the same formula and its own result.

Right after the fill, a lightning-bolt **AutoCorrect Options** button appears next to the cell. Use it if you didn't want a
calculated column:

| AutoCorrect option | What it does |
|---|---|
| **Undo Calculated Column** | Keeps the formula in the one cell you typed and removes it from the rest of the column |
| **Stop Automatically Creating Calculated Columns** | Turns the feature off for every Table. Turn it back on in **File → Options → Proofing → AutoCorrect Options → AutoFormat As You Type → Fill formulas in tables to create calculated columns** |

A calculated column stays consistent as you work:

- **Change the formula in any cell**, and Excel updates the whole column.
- **Add a row**, and the new row gets the formula automatically.
- **Type a value over one formula**, and that cell becomes an **exception**. Excel marks it with a green triangle
  (*Inconsistent Calculated Column Formula*). An exception usually means someone overtyped a result by hand, so treat it as an
  error to fix.

A calculated column can also return TRUE or FALSE. `=[@QtyOnHand]<=[@ReorderPoint]` flags every row at or below its reorder
point. You can then filter on the flags, or count them with `COUNTIF(tblInventory[NeedsReorder],TRUE)`.

Inside a Table you can leave out the Table name when you refer to a whole column. This calculated column shows each row's
share of the system's stock value:

```
=[@ExtendedValue]/SUM([ExtendedValue])
```

For the gloves in row 2, that's 0.43%.

> ⚠️ A calculated column starts automatically only when the column is empty. If the column already holds values, Excel puts
> your formula in just one cell and offers **Overwrite all cells in this column with this formula** on the AutoCorrect
> Options button.

### 6. Tables grow with your data

**AutoExpansion** adds new rows and columns to a Table as you type next to it:

- **New row:** type in the first empty row directly below the Table, or press **Tab** in the last cell of the last row.
- **New column:** type a header in the first empty cell to the right of the header row.
- **Paste:** paste rows directly below the Table.
- **Resize:** use **Table Design → Resize Table**, or drag the small handle at the Table's bottom-right corner.

Every structured reference follows the new size. `=SUM(tblInventory[QtyOnHand])` includes new rows without any editing, and
so do charts and drop-down lists built on the Table. A PivotTable built on the Table picks up the new rows the next time you
refresh it, without any change to its source (Lesson 3.4).

> 💡 **Try it:** After task 1, and before you turn on the Total Row, type `=SUM(tblInventory[QtyOnHand])` in an empty cell on
> the Practice sheet. On the Inventory sheet, type `TEST` in A259 and `100` in J259, and watch the sum grow by 100. Then press
> **Ctrl + Z** (Mac: **⌘ + Z**) until the test row is gone, so your practice answers stay correct.

> ⚠️ **The Total Row blocks AutoExpansion below the Table.** With the Total Row on, typing under it doesn't add a row. Press
> **Tab** in the last data cell instead, or right-click a row in the Table and choose **Insert → Table Rows Above**.

> 📋 If nothing expands when you type, check that **File → Options → Proofing → AutoCorrect Options → AutoFormat As You Type
> → Include new rows and columns in table** is ticked.

### 7. The Total Row

The **Total Row** is an extra row at the bottom of the Table that summarizes each column.

1. Click inside the Table.
2. Tick **Table Design → Total Row** (Windows shortcut: **Ctrl + Shift + T**; Mac: **Table → Total Row**).
3. Excel writes *Total* in the first column and adds a total under the last column.
4. Click any cell in the Total Row, open its dropdown, and pick the summary you want for that column.

Each choice writes a SUBTOTAL formula, which you met in Lesson 1.6:

| Dropdown choice | Formula Excel writes | Typical use |
|---|---|---|
| Sum | `=SUBTOTAL(109,[ExtendedValue])` | Dollar totals |
| Average | `=SUBTOTAL(101,[LeadTimeDays])` | Typical values |
| Count | `=SUBTOTAL(103,[StockID])` | Non-empty cells, text included |
| Count Numbers | `=SUBTOTAL(102,[QtyOnHand])` | Numeric cells only |
| Max / Min | `=SUBTOTAL(104,[UnitCost])` / `=SUBTOTAL(105,[UnitCost])` | Extremes |
| StdDev / Var | `=SUBTOTAL(107,…)` / `=SUBTOTAL(110,…)` | Spread (Lesson 2.4) |
| More Functions… | Whatever function you pick | Anything else |
| None | Clears the cell | |

The 100-series codes skip every hidden row, whether a filter or slicer hid it or you hid it by hand. So the Total Row always
describes exactly the rows you can see. With no filter on, Max under UnitCost shows 2,868.04, the cost of one Hip Femoral
Stem, Size 11, stocked in Perioperative Services.

To use a total in another formula, click the Total Row cell while you write the formula. Excel writes a reference such as
`tblInventory[[#Totals],[LeadTimeDays]]`, which keeps pointing at the Total Row as the Table grows.

> ⚠️ If you turn the Total Row off, every `[#Totals]` reference returns #REF!.

> 💡 **Tip:** `=SUM(tblInventory[ExtendedValue])` in a normal cell always adds every row, filtered or not. Only SUBTOTAL and
> AGGREGATE (Lesson 1.6) react to filters. Use the Total Row when you mean "the rows I can see," and SUM when you mean
> "every row."

### 8. Filter with slicers

A **slicer** is a panel of buttons, one for each value in a column, that filters the Table when you click a button.

1. Click inside the Table.
2. Choose **Table Design → Insert Slicer** (Mac: **Table → Insert Slicer**). **Insert → Slicer** works too.
3. Tick the columns you want, for example **Facility** and **Category**, and click **OK**.
4. Click a button to show only the matching rows.

| To do this | Do this |
|---|---|
| Select several values | Hold **Ctrl** (Mac: **⌘**) and click, or turn on the **Multi-Select** button in the slicer's header |
| Clear a slicer | Click **Clear Filter** (the funnel with a red X) at its top right, or select the slicer and press **Alt + C** (Windows) |
| Change its size, columns, or style | Select the slicer and use the **Slicer** tab |
| Remove a slicer | Select it and press **Delete** |

Values you select in one slicer combine with OR. Picking two buttons in a Location slicer shows Intensive Care Unit *or*
Pediatrics rows. Different slicers combine with AND, so a Facility slicer and a Category slicer together show Ashby Falls
*and* IV Supplies rows. A slicer is a friendlier front end for the Table's AutoFilter, so the header's filter button
shows the same filter, and the Total Row recalculates for the visible rows.

> ⚠️ A slicer hides rows. It doesn't remove them, so formulas such as `=SUM(tblInventory[ExtendedValue])` still include the
> hidden rows, and the Total Row stays filtered until you clear the slicer.

> 📋 **Version note:** Slicers on Tables need Excel 2013 or later on Windows. Current versions of Excel for Mac and Excel for
> the web support them too. Excel 2010 had slicers for PivotTables only (Lesson 3.4).

### 9. Table housekeeping

| To do this | Use |
|---|---|
| Change the style or banding | **Table Design → Table Styles**, plus the **Banded Rows** and **First Column** checkboxes |
| Hide the filter buttons | Untick **Table Design → Filter Button**, or press **Ctrl + Shift + L** (Mac: **⌘ + Shift + F**) |
| Remove duplicate rows | **Table Design → Remove Duplicates** (Lesson 3.3) |
| Turn the Table back into a plain range | **Table Design → Convert to Range** |

**Convert to Range** keeps the formatting but removes everything else that makes it a Table. Formulas that used structured
references switch to ordinary cell references, so `=SUM(tblInventory[QtyOnHand])` becomes something like
`=SUM(Inventory!$J$2:$J$258)`, and those references stop growing.

A Table can't contain some things:

- **Merged cells.** Unmerge them first, or use **Center Across Selection** instead (Lesson 1.3).
- **Spilled dynamic-array formulas** such as UNIQUE or FILTER (Lesson 4.1). Inside a Table they return #SPILL!, so put them
  next to the Table instead.
- **Multi-cell array formulas** from older versions of Excel.

> ⚠️ A structured reference to a Table in another workbook works only while that workbook is open. When it's closed, the
> formula returns #REF!. Keep the Table in the same workbook, or bring the data in with Power Query (Lesson 4.3).

### 10. Named ranges

A **defined name** is a label you give to a cell, a range, a constant, or a formula. A **named range** is a defined name that
points to cells, such as *ReportDate* for Settings!B2. Once it exists, you write `=ReportDate+7` instead of
`=Settings!$B$2+7`.

Names help in three ways:

- **Readable formulas.** `">"&ReportDate-7` says what it compares against. `">"&Settings!$B$2-7` doesn't.
- **One place to change an assumption.** Type a new date in Settings!B2, and every formula that uses ReportDate updates.
- **Fewer reference mistakes.** Names you create with the Name Box or Create from Selection are absolute references, so they
  never shift when you copy a formula.

There are three ways to create a name:

| Method | Steps | Best for |
|---|---|---|
| **Name Box** | Select the cell, click the Name Box (left of the formula bar), type `ReportDate`, and press **Enter** | One cell or range, quickly |
| **Define Name** | **Formulas → Define Name**, then fill in Name, Scope, Comment, and Refers to | Constants, formulas, choosing a scope, adding a comment |
| **Create from Selection** | Select the labels and the values together, choose **Formulas → Create from Selection** (Windows: **Ctrl + Shift + F3**), and tick where the labels are | Naming many cells at once from labels already on the sheet |

The Settings sheet is laid out for Create from Selection. Select A2:B3, tick **Left column** only, and Excel names B2
*ReportDate* and B3 *ExpiringWindowDays* after the labels beside them.

> ⚠️ Press **Enter** after typing in the Name Box. If you click away instead, Excel doesn't create the name.

> ⚠️ In the Create from Selection dialog, Excel guesses which boxes to tick. If **Top row** is also ticked, Excel treats row 2
> as a row of column labels, so ReportDate isn't created the way you want. Untick everything except **Left column**.

#### Naming rules

| Rule | Allowed | Not allowed |
|---|---|---|
| First character | A letter, an underscore, or a backslash | `90DayWindow` |
| Other characters | Letters, numbers, periods, and underscores | `Report Date` (space), `Lead-Time` (hyphen) |
| Can't look like a cell address | `ReportDate`, `Q1Budget` | `Q1`, `JAN2025`, `TBL1`, `R1C1` |
| Single letters | Any letter except C and R | `C` and `R`, which Excel reserves for R1C1 notation |
| Length | Up to 255 characters | |
| Case | Not case-sensitive: `reportdate` and `ReportDate` are the same name | |

Defined names and Table names share one list, so a name can't match a Table name. Create from Selection repairs labels
that break the rules, so a label like *Report Date* becomes the name `Report_Date`.

#### Use names in formulas

- **Type it.** Names appear in the AutoComplete list next to functions as soon as you type the first letters.
- **Paste it.** Press **F3** (Windows) to open the **Paste Name** list, or use **Formulas → Use in Formula** on either
  platform.
- **Go to it.** Pick a name from the Name Box dropdown, or press **F5** (Mac: **Control + G**), to jump to its cells and
  select them.

A name works in a criterion exactly like a cell reference. Join the operator to the name with `&`. For example, this counts
the stock rows counted in the last 7 days, meaning after 12/24/2025:

```
=COUNTIF(tblInventory[LastCountDate],">"&ReportDate-7)      → 37
```

#### Scope: workbook or worksheet

Every name has a **scope**, the part of the workbook where you can use it by its name alone.

| Scope | How you use it | When to choose it |
|---|---|---|
| **Workbook** (the default) | `=ReportDate` from any sheet | Almost always |
| **Worksheet**, such as Settings | `=ReportDate` on the Settings sheet, and `=Settings!ReportDate` from other sheets | When the same name needs a different value on each sheet, such as a `MonthStart` on each monthly sheet |

You choose the scope in the New Name dialog when you create the name. You can't change it afterward. To change it, delete the
name and create it again.

### 11. Named constants and named formulas

A name doesn't have to point to cells. A **named constant** is a name whose *Refers to* box holds a value instead of a cell
address:

1. Choose **Formulas → Define Name**.
2. Type `CycleCountDays` in **Name**.
3. Replace whatever is in **Refers to** with `=30`.
4. Optionally add a **Comment**, such as *Policy: count every stock row at least every 30 days*, then click **OK**.

Now `=ReportDate-CycleCountDays` works in any formula and returns 12/01/2025. A **named formula** is the same idea with a
formula in *Refers to*. For example, a name *WindowEnd* that refers to `=ReportDate+ExpiringWindowDays` gives you the last day
of the expiry window anywhere in the workbook.

| | Named cell (ReportDate) | Named constant (CycleCountDays) |
|---|---|---|
| Where the value lives | In a cell everyone can see | Only in Name Manager |
| How you change it | Type a new value in the cell | Edit the name in Name Manager |
| Risk | Someone can type over it | Hard to change by accident, but also easy to forget about |
| Best for | Inputs people change often, such as report dates and scenario assumptions | Policy values and conversion factors, such as cycle-count days or minutes per day |

> 💡 **Tip:** Put a report's "today" in a named cell such as ReportDate instead of using `TODAY()`. The workbook then gives the
> same answers tomorrow, and you can rerun last month's report by changing one cell (Lesson 2.3).

### 12. Manage names with Name Manager

**Formulas → Name Manager** (Windows: **Ctrl + F3**) lists every name in the workbook, Table names included:

| Column | Shows |
|---|---|
| Name | The name, with an icon that tells defined names and Tables apart |
| Value | The current value, or the first few values of a range |
| Refers To | The cells, constant, or formula behind the name |
| Scope | *Workbook*, or the sheet the name belongs to |
| Comment | Your note about the name |

From here you can **Edit** a name to rename it or change what it refers to, **Delete** it, or use **Filter** to show only
names with errors, only Table names, or only names of one scope. When you rename a name, every formula that uses it updates.

> ⚠️ If you delete the cells a name points to, the name refers to `#REF!`, and every formula that uses it breaks. After you
> delete rows, columns, or sheets, filter Name Manager by **Names with Errors**, then fix or delete what it shows.

On Windows, two more tools help when you inherit someone else's workbook:

- **F3 → Paste List** writes every name and what it refers to onto the sheet, starting at the active cell. It's a quick way to
  document a model.
- **Formulas → Define Name ▾ → Apply Names** rewrites existing formulas to use names, so `=Settings!$B$2+7` becomes
  `=ReportDate+7`.

> 📋 Older Mac versions of Excel have no Name Manager button. There, **Formulas → Define Name** shows the list of names and
> lets you add or delete them.

### 13. Dynamic named ranges, and why Tables replace them

Before Tables existed, analysts built **dynamic named ranges**: names whose size is calculated by a formula, so they grow when
rows are added. Here are the two classic versions for the QtyOnHand column, written in the *Refers to* box of a name such as
*QtyList*:

```
OFFSET version:  =OFFSET(Inventory!$J$2,0,0,COUNTA(Inventory!$A:$A)-1,1)
INDEX version:   =Inventory!$J$2:INDEX(Inventory!$J:$J,COUNTA(Inventory!$A:$A))
```

Both count the filled cells in column A (257 StockIDs plus the header, so 258) to find where the data ends. Then
`=SUM(QtyList)` adds J2:J258 and keeps up as rows are added.

| | OFFSET name | INDEX name | Table column |
|---|---|---|---|
| Grows with new rows | Yes | Yes | Yes |
| Recalculates | After every change anywhere, because OFFSET is **volatile** | Only when its inputs change | Only when its inputs change |
| Breaks if column A has a blank cell or a stray value below the data | Yes | Yes | No |
| Readable | Hard | Hard | `tblInventory[QtyOnHand]` |
| Where the logic lives | Hidden in Name Manager | Hidden in Name Manager | In the formula itself |

For lists of records, a Table is simpler, faster, and harder to break. You'll still meet dynamic names in older workbooks,
and they're useful for blocks that aren't shaped like a list, so it's worth recognizing them.

> ⚠️ COUNTA counts every filled cell in column A. With a Table's Total Row turned on, the word *Total* sits in column A, so
> both dynamic names grow by one row and include the total. A Table column never includes its own Total Row.

### 14. Tables or names?

| You have | Use |
|---|---|
| A list of records that grows or shrinks | A Table, with structured references |
| A single input or assumption on a sheet | A named cell |
| A fixed policy value or conversion factor | A named constant |
| A calculation you reuse in many formulas | A named formula |
| A block that isn't a list, such as a rate grid | A named range |

Tables and names work together. In the practice tasks, Table columns supply the data and names supply the settings.

### 15. Keyboard shortcuts at a glance

| Action | Windows | Mac |
|---|---|---|
| Create a Table | Ctrl + T (or Ctrl + L) | Control + T |
| Toggle the Total Row | Ctrl + Shift + T | **Table → Total Row** |
| Toggle the filter buttons | Ctrl + Shift + L | ⌘ + Shift + F |
| Open a header's filter menu | Alt + ↓ | ⌥ + ↓ |
| Select the Table's data (press again to add the headers and Total Row) | Ctrl + A | ⌘ + A |
| Select the current Table column | Ctrl + Space | Control + Space |
| Select the current Table row | Shift + Space | Shift + Space |
| Clear the selected slicer | Alt + C | Click **Clear Filter** |
| Open Name Manager | Ctrl + F3 | **Formulas → Name Manager** |
| Create names from a selection | Ctrl + Shift + F3 | **Formulas → Create from Selection** |
| Paste a name into a formula | F3 | **Formulas → Use in Formula** |
| Go To (lists named ranges) | F5 or Ctrl + G | Control + G |

## 🧪 Hands-on practice

Download [`3.1-tables-named-ranges.xlsx`](3.1-tables-named-ranges.xlsx) and open the **Practice** sheet. Work through the tasks
in order, because task 1 builds the Table that every later task uses. Type each answer in a yellow cell. Gray cells read the
work you do on the Inventory and Settings sheets. The **Check** column turns green when you're right.

<!-- BEGIN GENERATED: practice -->
The tasks use the Inventory sheet: supply stock in 15 storerooms across three Bluestone hospitals, in the snapshot taken on 12/31/2025. Task 9 also uses the Settings sheet. Do the tasks in order, because later tasks use the Table, columns, Total Row, and names you create in earlier ones. After task 1, write formulas with structured references such as tblInventory[QtyOnHand] instead of cell ranges.

| # | Task | Hint |
|:-:|------|------|
| 1 | On the Inventory sheet, click any cell in the data and convert the range to an Excel Table with Ctrl + T (Mac: Control + T). Then rename the Table tblInventory in the Table Name box on the Table Design tab (Mac: Table tab). The gray cell finds a Table with that exact name and counts its data rows. It stays blank until the Table exists. | Make sure 'My table has headers' is ticked |
| 2 | How many stock rows are completely out of stock (QtyOnHand = 0)? Use a structured reference to the QtyOnHand column instead of a cell range. | COUNTIF. Type tblInventory[ and pick the column from the list |
| 3 | Fill the yellow ExtendedValue column (column R) with a calculated column: each row's QtyOnHand × UnitCost. Type the formula once in R2 and press Enter, and the Table fills every row. The gray cell totals the column: what's the value of all stock on hand across the system? | Click the QtyOnHand cell in the same row while you type, and Excel writes [@QtyOnHand] |
| 4 | What share of the system's total inventory value (ExtendedValue) is Orthopedic Implants stock? Divide the category's value by the total of all rows and enter it as a percentage, to 1 decimal place. | SUMIFS(…)/SUM(…), both with structured references |
| 5 | Fill the yellow NeedsReorder column (column S) with a calculated column that returns TRUE when QtyOnHand is at or below ReorderPoint, and FALSE otherwise. The gray cell counts the TRUE rows: how many stock rows need reordering? | A comparison already returns TRUE or FALSE, so you don't need IF |
| 6 | Turn on the Table's Total Row and set the LeadTimeDays total to Average. Then, in the yellow cell, type = and click that Total Row cell, so Excel writes a [#Totals] reference. What's the average vendor lead time in days, to 2 decimal places? | Table Design → Total Row (Windows: Ctrl + Shift + T), then use the dropdown in the Total Row cell |
| 7 | With the Total Row still on, what does =ROWS(tblInventory[#All]) return? Predict it first, then type the formula to check. | #All is everything: which rows does it include that tblInventory[QtyOnHand] doesn't? |
| 8 | Insert slicers for Facility and Category (Table Design → Insert Slicer). Select Cedar Ridge Medical Center in the Facility slicer and Surgical in the Category slicer. Set the Total Row's ExtendedValue cell to Sum. What value does it show? Type the number, then clear both slicers, because while they filter the Table the Total Row (and task 6) only summarizes the visible rows. | Slicers filter the Table, and SUBTOTAL ignores rows the filter hides |
| 9 | On the Settings sheet, select A2:B3 and use Formulas → Create from Selection (tick only Left column) to name cell B2 ReportDate and cell B3 ExpiringWindowDays. The gray cell looks up both names and adds them: what date does it show? | The labels in column A become the names of the cells beside them |
| 10 | How many stock rows expire inside the expiry window? Count rows whose ExpirationDate is later than ReportDate and no later than ReportDate + ExpiringWindowDays. Use the two names in a COUNTIFS (rows with no ExpirationDate don't count). | Join an operator to a name the way you join it to a cell: ">"&ReportDate |
| 11 | Expired stock must be pulled from the shelves and written off. What's the total ExtendedValue of rows whose ExpirationDate is before ReportDate? | SUMIFS with a "<"&ReportDate criterion |
| 12 | Policy says every stock row must be cycle-counted at least every 30 days. Create a named constant CycleCountDays that refers to =30 with Formulas → Define Name. It doesn't live in any cell. How many rows are overdue, meaning ReportDate − LastCountDate is greater than CycleCountDays? | Rearrange: overdue means LastCountDate < ReportDate − CycleCountDays |
<!-- END GENERATED: practice -->

## ✅ Answer key

The workbook has a hidden **Answer Key** sheet (right-click any sheet tab → **Unhide…**). Its *Live result* column recalculates
each answer with ordinary cell ranges, because the Table and names don't exist until you create them. The answers are also
below, collapsed so you don't see them by accident.

<!-- BEGIN GENERATED: answers -->
<details>
<summary><b>🔑 Show the answer key</b> — Try every task before opening this.</summary>

**1. Convert the range to a Table named tblInventory (data rows)**

- **Answer:** 257
- **Solution:**

1. Click any cell inside the data, for example **A2**.
2. Press **Ctrl + T** (Mac: **Control + T**), or choose **Insert → Table**.
3. Check that the range shows **$A$1:$S$258** and that **My table has headers** is ticked, then click **OK**.
4. On the **Table Design** tab (Mac: **Table** tab), click in the **Table Name** box at the far left, type `tblInventory`, and press **Enter**.


Ctrl + T selects the whole block of connected cells around the active cell, so you don't need to select the data first. Excel names new Tables Table1, Table2, and so on. Renaming the Table is what makes formulas readable: `tblInventory[UnitCost]` tells the next reader what the formula uses, while `Table1[UnitCost]` doesn't. The gray cell uses `INDIRECT("tblInventory")`, which looks the Table up by name and returns its data rows (headers and Total Row excluded).

**2. How many stock rows are completely out of stock (QtyOnHand = 0)? Use a structured…**

- **Answer:** 9
- **Solution:** `=COUNTIF(tblInventory[QtyOnHand],0)`

`tblInventory[QtyOnHand]` means "every data cell in the QtyOnHand column." It excludes the header and the Total Row, and it grows when rows are added. After you type `tblInventory[`, Excel lists the columns. Pick one with the arrow keys and press **Tab**. These rows are *stockouts*: a nurse who reaches for the item finds an empty bin.

**3. ExtendedValue calculated column (total stock value)**

- **Answer:** 1,011,928.89
- **Solution:** `=[@QtyOnHand]*[@UnitCost]`

When you enter a formula in an empty Table column, Excel copies it to every row of that column. That's a *calculated column*. `[@QtyOnHand]` means "QtyOnHand in this row," so the formula reads the same on all 257 rows and says what it multiplies. The A1 version, `=J2*G2`, changes on every row and tells the next reader nothing about the columns it uses.

**4. What share of the system's total inventory value (ExtendedValue) is Orthopedic…**

- **Answer:** 14.0%
- **Solution:**

```
=SUMIFS(tblInventory[ExtendedValue],tblInventory[Category],"Orthopedic Implants")/SUM(tblInventory[ExtendedValue])
```


Every argument in SUMIFS is a whole Table column, so the ranges are automatically the same size, which SUMIFS requires. Only 15 of the 257 rows are implants, yet they hold 14% of the value on the shelves. High-cost, low-volume items like these are the ones a supply chain team counts most carefully.

**5. NeedsReorder calculated column (rows to reorder)**

- **Answer:** 41
- **Solution:** `=[@QtyOnHand]<=[@ReorderPoint]`

A comparison such as `[@QtyOnHand]<=[@ReorderPoint]` evaluates to TRUE or FALSE by itself (Lesson 2.1). Storing the result as a column turns a row-by-row test into something you can filter, count with `COUNTIF(tblInventory[NeedsReorder],TRUE)`, and reuse in other formulas, as the bonus does. Use "at or below" (`<=`): an item sitting exactly at its reorder point is due.

**6. Turn on the Table's Total Row and set the LeadTimeDays total to Average. Then, in the…**

- **Answer:** 5.35
- **Solution:** `=tblInventory[[#Totals],[LeadTimeDays]]`

The Total Row's dropdown writes `=SUBTOTAL(101,[LeadTimeDays])`. Function number 101 means AVERAGE of the visible rows only, ignoring rows hidden by a filter or hidden by hand (109 = SUM, 103 = COUNTA, 104 = MAX). `tblInventory[[#Totals],[LeadTimeDays]]` points at that total cell by name, so it still works if rows are added. If you turn the Total Row off, the reference returns #REF!.

**7. With the Total Row still on, what does =ROWS(tblInventory[#All]) return? Predict it…**

- **Answer:** 259
- **Solution:** `=ROWS(tblInventory[#All])`

`[#All]` covers the header row, the 257 data rows, and the Total Row: 1 + 257 + 1 = 259. `tblInventory[#Data]` (or just `tblInventory`) is the 257 data rows, `[#Headers]` is the header row alone, and `[#Totals]` is the Total Row alone.

**8. Slicers: Surgical stock at Cedar Ridge Medical Center**

- **Answer:** 47,775.51
- **Solution:**

1. Click inside the Table, then **Table Design → Insert Slicer** (Mac: **Table → Insert Slicer**). Tick **Facility** and **Category** and click **OK**.
2. In the Facility slicer click **Cedar Ridge Medical Center**. In the Category slicer click **Surgical**. The Table now shows 6 rows.
3. In the Total Row, click the ExtendedValue cell, open its dropdown, and choose **Sum**.
4. Read the total and type it in the answer cell: **47,775.51**.
5. Click the **Clear Filter** button (funnel with a red X) at the top right of each slicer.


Slicers are buttons that apply the Table's AutoFilter for you. Choices in different slicers combine with AND, so only rows that match both stay visible. The Total Row uses SUBTOTAL(109,…), which adds only the visible rows. That's why it changes as you click. You can check the number without slicers using `=SUMIFS(tblInventory[ExtendedValue],tblInventory[Facility],"Cedar Ridge Medical Center",tblInventory[Category],"Surgical")`.

**9. Name the Settings cells with Create from Selection**

- **Answer:** 03/31/2026
- **Solution:**

1. Go to the **Settings** sheet and select **A2:B3** (labels and values together).
2. Choose **Formulas → Create from Selection** (Windows shortcut: **Ctrl + Shift + F3**).
3. Tick **Left column** only, untick **Top row**, and click **OK**.
4. Check the result: click B2 and the Name Box shows **ReportDate**. Open **Formulas → Name Manager** to see both names and what they refer to.


Create from Selection turns each label in the left column into a name for the cell to its right: ReportDate → `=Settings!$B$2` and ExpiringWindowDays → `=Settings!$B$3`. Names are absolute references, so they never shift when you copy a formula. 12/31/2025 + 90 days = 03/31/2026, because dates are day counts (Lesson 2.3). If the gray cell stays blank, check the spelling of each name in Name Manager.

**10. How many stock rows expire inside the expiry window? Count rows whose ExpirationDate…**

- **Answer:** 7
- **Solution:**

```
=COUNTIFS(tblInventory[ExpirationDate],">"&ReportDate,tblInventory[ExpirationDate],"<="&ReportDate+ExpiringWindowDays)
```


Criteria are text, so join the operator to the name with `&`, exactly as you would with a cell reference (Lesson 2.5). Blank ExpirationDate cells never match a date criterion, so non-perishables drop out automatically. The formula reads like the policy, and when the next snapshot arrives you only change the date in Settings!B2.

**11. Expired stock must be pulled from the shelves and written off. What's the total…**

- **Answer:** 17,115.69
- **Solution:** `=SUMIFS(tblInventory[ExtendedValue],tblInventory[ExpirationDate],"<"&ReportDate)`

In this snapshot, 10 rows are past their expiration date. That's a patient-safety issue as well as a cost: surveyors such as The Joint Commission look for expired supplies in clinical areas. The live formula in the key guards against blanks with `<>""`, but SUMIFS doesn't need that guard, because an empty cell never matches "<"&ReportDate.

**12. Policy says every stock row must be cycle-counted at least every 30 days. Create a…**

- **Answer:** 78
- **Solution:** `=COUNTIF(tblInventory[LastCountDate],"<"&ReportDate-CycleCountDays)`

In the New Name dialog, type `CycleCountDays` as the Name and `=30` in **Refers to**. ReportDate − CycleCountDays is 12/01/2025, so any row last counted before that date is more than 30 days old. A named constant suits a policy value that shouldn't be typed over by accident. A named cell, like ReportDate, suits an input that people change often.

</details>
<!-- END GENERATED: answers -->

## 🏆 Bonus challenge

<!-- BEGIN GENERATED: bonus -->
Every Monday, Bluestone's purchasing team turns the inventory snapshot into purchase orders. Each stock row at or below its reorder point gets an order for its standard ReorderQty, and all the lines for one vendor go on a single purchase order. Build the report with Tables and names, so next week it updates itself when the new snapshot is pasted in. Finish practice tasks 1–12 first, because this uses tblInventory, NeedsReorder, and the names.

Work on the **Bonus** sheet of the workbook.

- **B1.** Add a new column to tblInventory by typing OrderCost in cell T1, just right of the NeedsReorder header. The Table expands to include it. Make it a calculated column that returns ReorderQty × UnitCost for rows where NeedsReorder is TRUE, and 0 for every other row. The gray cell totals the column: what's the total cost of this week's orders? *(Hint: IF can test a TRUE/FALSE column directly: IF([@NeedsReorder], …, 0))*
- **B2.** On the Vendors sheet, convert the vendor list to a Table named tblVendors. Fill its yellow ReorderCost column with a calculated column that adds up tblInventory[OrderCost] for that row's vendor. Which vendor gets the largest purchase order? (Type the vendor's name, or return it with a formula.) *(Hint: In tblVendors, SUMIFS can add up another Table's column, with [@Vendor] as the criterion)*
- **B3.** What's the value of that vendor's purchase order? *(Hint: The largest value in tblVendors[ReorderCost])*
- **B4.** The group purchasing contract gives a 2% discount on any single vendor order of $20,000 or more. Define two named constants, DiscountMin (=20000) and DiscountPct (=0.02), then calculate the total discount on this week's orders, to the cent. *(Hint: SUMIFS can use the same column as the sum range and the criteria range)*
<!-- END GENERATED: bonus -->

<!-- BEGIN GENERATED: bonus-answers -->
<details>
<summary><b>🔑 Show the bonus solution</b> — Give it a real try first!</summary>

**B1. OrderCost calculated column (total of all orders)**

- **Answer:** 173,950.91
- **Solution:** `=IF([@NeedsReorder],[@ReorderQty]*[@UnitCost],0)`

Typing in the first empty column next to a Table adds a column to it (Excel's *AutoExpansion*). `IF([@NeedsReorder],…)` reuses the calculated column you built in task 5 instead of repeating the comparison. The 41 rows that need an order return their cost and the rest return 0, so the column total is the cost of the whole report.

**B2. On the Vendors sheet, convert the vendor list to a Table named tblVendors. Fill its…**

- **Answer:** Summit Orthopedic Systems
- **Solution:** `=INDEX(tblVendors[Vendor],MATCH(MAX(tblVendors[ReorderCost]),tblVendors[ReorderCost],0))`

The calculated column in tblVendors is `=SUMIFS(tblInventory[OrderCost],tblInventory[Vendor],[@Vendor])`. It mixes whole columns from *another* Table with `[@Vendor]` from this row, which is how Tables talk to each other. INDEX/MATCH (Lesson 2.6) then returns the vendor on the row with the largest total. `=XLOOKUP(MAX(tblVendors[ReorderCost]),tblVendors[ReorderCost],tblVendors[Vendor])` works too in Excel 2021 or Microsoft 365. All four reorder lines for Summit Orthopedic Systems are orthopedic implants, and three of them cost more than $2,000 per unit, so a handful of lines makes the largest order.

**B3. What's the value of that vendor's purchase order?**

- **Answer:** 62,855.33
- **Solution:** `=MAX(tblVendors[ReorderCost])`

That's 36% of the week's spend on a single purchase order. In practice, a buyer would confirm implant orders with the OR schedule before sending them.

**B4. The group purchasing contract gives a 2% discount on any single vendor order of…**

- **Answer:** 2,431.66
- **Solution:** `=SUMIFS(tblVendors[ReorderCost],tblVendors[ReorderCost],">="&DiscountMin)*DiscountPct`

Orders from 3 vendors reach $20,000. Together they total $121,583.18, and 2% of that is the discount. With the thresholds in named constants, Purchasing can test a new contract (say 3% at $15,000) by editing two names in Name Manager, and every formula that uses them updates.

</details>
<!-- END GENERATED: bonus-answers -->

## Key takeaways

- Convert every list of records to a Table (**Ctrl + T**; Mac: **Control + T**) and give it a descriptive name such as
  `tblInventory` right away.
- Structured references like `tblInventory[QtyOnHand]` and `[@QtyOnHand]` are always the right size, grow with the data,
  and read like the question you're asking.
- A formula entered in an empty Table column becomes a calculated column: one formula, every row, no copying.
- Slicers filter the Table, and the Total Row uses SUBTOTAL, so it summarizes only the rows the slicers leave visible.
  `SUM(tblInventory[…])` always includes every row.
- Name your inputs (ReportDate) and your policy values (CycleCountDays) so formulas explain themselves and every assumption
  lives in one place.
- Check Name Manager for `#REF!` names after deleting cells, and prefer Tables to OFFSET or INDEX dynamic ranges for lists that
  grow.

<!-- BEGIN GENERATED: nav -->
---

⬅️ **Previous:** [2.6 Lookups: VLOOKUP, INDEX/MATCH & XLOOKUP](../../02-formulas-functions/06-lookup-functions/README.md) · 🏠 [Course home](../../README.md) · **Next:** [3.2 Data Validation & Conditional Formatting](../02-data-validation-conditional-formatting/README.md) ➡️
<!-- END GENERATED: nav -->
