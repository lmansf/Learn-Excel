# Lesson 3.2 · Data Validation & Conditional Formatting

> **Level:** Intermediate · **Time:** about 120 minutes · **Workbook:** [`3.2-data-validation-conditional-formatting.xlsx`](3.2-data-validation-conditional-formatting.xlsx)
> **Data:** A December 2025 Patient Access intake log with entry errors, STAT lab results from the three ICUs (Jul–Dec 2025), the system supply inventory (12/31/2025 snapshot), and Medical-Surgical 5 East's daily census (Nov–Dec 2025). The bonus adds a bed-huddle board for every inpatient unit. Column definitions are in the [data dictionary](../../data/README.md).

A registrar who types "Self Pay" instead of "Self-Pay" creates a claim that bounces weeks later. An MRN that lost its
leading zero matches no patient, so a lab result lands in limbo. A charge nurse scrolling 300 ICU results can miss the one
critical lactate in the middle. Excel has two tools for these problems. **Data validation** stops bad entries at the moment
someone types them, and **conditional formatting** makes the rows that matter impossible to miss. In this lesson you'll add
both to real-looking Bluestone Health workbooks: a registration log, ICU lab results, supply stock, and a bed-huddle board.

## What you'll learn

- Restrict entries with list, number, date, length, and custom-formula validation
- Build dependent drop-down lists
- Highlight what matters with conditional formatting rules, data bars, color scales, and icon sets
- Write formula-based rules that format entire rows

## 📖 Guide

### 1. Two tools, one idea

Both tools attach a **rule** to a range of cells. The rule is a test that Excel runs on every cell in the range, and the
result decides what happens.

| | Data validation | Conditional formatting |
|---|---|---|
| **Job** | Prevent bad entries | Make patterns and exceptions visible |
| **When it acts** | The moment someone types into the cell | Continuously, every time values change |
| **What a TRUE test means** | The entry is allowed | The cell gets the format |
| **Where to find it** | **Data → Data Validation** | **Home → Conditional Formatting** |
| **Changes the stored value?** | No | No |
| **Healthcare example** | Payer must come from the contracted-payer list | Critical lab results turn red |

Two habits apply to both tools, and the rest of the lesson builds on them:

1. **Select the whole range first, then write the rule for the active cell.** The **active cell** is the white cell inside
   the selection (the one shown in the Name Box). Excel copies your rule to every other cell in the range and adjusts its
   references the same way it adjusts a formula you copy down.
2. **Rules don't change data.** A validation rule can't fix what's already in a cell, and a red fill doesn't make a value
   any different to a formula.

### 2. Data validation basics

Select the cells, then open **Data → Data Validation** (Windows: **Alt, A, V, V**. Mac: **Data** tab → **Data Validation**).
The dialog has three tabs:

- **Settings** holds the rule itself: what to **Allow**, the comparison, and the limits.
- **Input Message** shows a note (like a ScreenTip) when someone selects the cell.
- **Error Alert** decides what happens when someone types something the rule rejects.

The **Allow** box offers eight kinds of rule:

| Allow | Accepts | Healthcare example |
|---|---|---|
| Any value | Anything (the default, which means no rule) | Patient name |
| Whole number | Integers only, within limits | Pain score 0 to 10 |
| Decimal | Any number within limits | Copay between \$0 and \$100 |
| List | Only items from a list, with an optional drop-down arrow | Payer, Facility |
| Date | Dates within limits | Visit date in December 2025 |
| Time | Times within limits | Appointment between 07:00 and 18:00 |
| Text length | Entries with a certain number of characters | ZIP code of exactly 5 characters |
| Custom | Anything for which your formula returns TRUE | MRN of exactly 8 digits |

For number, date, time, and length rules, the **Data** box sets the comparison: *between*, *not between*, *equal to*,
*not equal to*, *greater than*, *less than*, *greater than or equal to*, or *less than or equal to*. **Between includes both
limits**, so a copay rule "between 0 and 100" accepts 0 and 100.

The limit boxes accept numbers, cell references, named cells, and formulas. That matters for dates, because typing
12/1/2025 means December 1 in the US and January 12 in much of the world.

```
Allow: Date    Data: between
Start date:  =DATE(2025,12,1)
End date:    =DATE(2025,12,31)
```

Writing `=DATE(year, month, day)` makes the rule mean the same thing on every computer.

The **Ignore blank** checkbox (on by default) treats an empty cell as valid. Leave it on unless the field is required.

> ⚠️ **Validation only checks typing.** Values that were in the cells before you added the rule stay put, and so do values
> that arrive by paste, Fill, a formula, or a macro. Section 7 shows how to find them. Worse, a normal paste over a validated
> cell **replaces its rule** with whatever rule (or no rule) the copied cell had. Paste with **Paste Special → Values** to keep
> the destination's rule.

### 3. Input messages and error alerts

An **input message** is a short prompt that appears when someone selects the cell, before they type. Use it to state the
rule in plain words: *"Enter the amount collected, \$0 to \$100."* The title holds up to 32 characters and the message up to
255.

The **error alert** appears after an invalid entry. Its **Style** decides how strict the rule is:

| Style | Icon | What happens after an invalid entry | Use it when |
|---|---|---|---|
| **Stop** | ⛔ | Excel refuses the entry. **Retry** lets you fix it, **Cancel** restores the old value | The value must be right (MRN, payer) |
| **Warning** | ⚠️ | Asks "Continue?" **Yes** keeps the entry, **No** lets you edit, **Cancel** restores the old value | Unusual but possible values (a \$150 copay) |
| **Information** | ℹ️ | Shows your message. **OK** keeps the entry, **Cancel** restores the old value | Gentle reminders |

Write your own **Title** and **Error message**, because Excel's default ("This value doesn't match the data validation
restrictions defined for this cell") doesn't tell the registrar what to type instead. A good error message says what's
allowed: *"Copay must be between \$0 and \$100. Enter refunds in the Refund log."*

> ⚠️ Warning and Information alerts let invalid data in after one click. If a field feeds billing or patient matching, use Stop.

### 4. Drop-down lists

A **List** rule restricts a cell to a set of items and, with **In-cell dropdown** ticked, adds an arrow next to the cell.
Press **Alt + ↓** (Mac: **Option + ↓**) to open the list from the keyboard. The **Source** box accepts any of these:

| Source | Example | Notes |
|---|---|---|
| Typed list | `Medicare,State Medicaid,Self-Pay` | Quick for a few fixed items. Separate items with your regional list separator (a comma in the US, often a semicolon in Europe). Limit: 255 characters |
| Range on the same or another sheet | `=Lists!$A$2:$A$9` | Click the range while the Source box is active. Other-sheet ranges work in Excel 2010 and later |
| Named range | `=PayerList` | Easiest to read and reuse. [Lesson 3.1](../01-tables-named-ranges/README.md) shows how to name a range |
| Table column | `=INDIRECT("tblPayers[PayerName]")` | Excel won't accept a structured reference typed straight into Source, so wrap it in INDIRECT (or name the column). The list then grows when the Table grows |

Worked example (task 1): the Lists sheet holds the eight contracted payers in **A2:A9**. Select **IntakeLog!F2:F51**, choose
**Allow: List**, click in **Source**, and select **Lists!A2:A9**. The arrow appears on every Payer cell, and a registrar can no
longer type "Self Pay" by hand.

> ⚠️ A List rule compares the whole entry with each item. "Self Pay" and "Self-Pay" are different, and so is "Medicare " with
> a trailing space. When the Source is a range, matching ignores upper and lower case. A list typed into the Source box is
> case-sensitive.

> 📋 **Version note:** In current Microsoft 365 versions, a drop-down list filters its items as you type in the cell, so
> long lists stay usable.

### 5. Custom formula rules

**Allow: Custom** accepts any formula that returns TRUE or FALSE. TRUE means the entry is valid. Write the formula for the
active cell, exactly as you would type it in that cell, and Excel adjusts it for the rest of the range.

Worked example (task 5): a valid MRN has exactly 8 characters, all digits. Select **C2:C51** with **C2** active and enter:

```
=AND(LEN(C2)=8, ISNUMBER(--C2))
```

| Part | What it checks | Fails for |
|---|---|---|
| `LEN(C2)=8` | The entry has 8 characters | 0256901 (7), 027762471 (9), an MRN stored as a number that lost its leading zero |
| `--C2` | Converts the text to a number. The double minus is two negations, so the value doesn't change | |
| `ISNUMBER(--C2)` | The conversion worked, so every character is a digit | O9127341 (the letter O instead of zero), because `--"O9127341"` returns #VALUE! |
| `AND(…)` | Both tests pass | Anything that fails either test |

A cell can hold only **one** validation rule, so combine every condition for a column into a single formula with AND or OR.
If the formula returns an error, Excel treats the entry as invalid.

More custom rules you can adapt:

| Rule | Custom formula (active cell in row 2) |
|---|---|
| No duplicate IntakeIDs | `=COUNTIF($A$2:$A$51, A2)=1` |
| IntakeID must start with "INT-" | `=LEFT(A2,4)="INT-"` |
| Visit date must be a weekday | `=WEEKDAY(B2,2)<6` |
| Discharge can't be before admission (two date columns) | `=E2>=D2` |
| Code must be typed in capitals | `=EXACT(A2, UPPER(A2))` |

Notice the `$` signs in the duplicate rule. `$A$2:$A$51` must stay fixed for every cell, while `A2` must move to A3, A4, and
so on. That's the same mixed-reference thinking you learned in [Lesson 1.5](../../01-foundations/05-cell-references/README.md).

### 6. Dependent drop-down lists

A **dependent drop-down** is a list whose items depend on the choice in another cell. The cell that makes the first choice
is the **parent**, and the cell whose list changes is the **child**. In the intake log, Facility is the parent and Department
is the child, so the Department list should show only the departments of the facility chosen in the same row. Department
names repeat across facilities ("Emergency Department", "Intensive Care Unit"), so a single long list would let a registrar
pick a department that doesn't exist at that hospital.

The classic method uses one named range per parent value and the **INDIRECT** function, which turns text into a reference.
`=INDIRECT("Cedar_Ridge_Medical_Center")` returns whatever range that name points to.

1. **Lay out the child lists.** On the Lists sheet, each column E:H has a facility name as its header and that facility's
   departments below it.
2. **Name each list after its parent.** Select **E1:E16** (header plus items) and choose **Formulas → Create from Selection**
   (Windows: **Ctrl + Shift + F3**), tick only **Top row**, and click **OK**. Repeat for F1:F5, G1:G6, and H1:H7. Names can't
   contain spaces, so Excel names the first range `Bluestone_Memorial_Hospital`. Check the result in **Formulas → Name
   Manager** (Windows: **Ctrl + F3**).
3. **Create the parent list.** Select **IntakeLog!H2:H51** and add a List rule with Source `=Lists!$C$2:$C$5`.
4. **Create the child list.** Select **IntakeLog!I2:I51** with **I2** active and add a List rule with this Source:

   ```
   =INDIRECT(SUBSTITUTE($H2," ","_"))
   ```

   SUBSTITUTE turns "Cedar Ridge Medical Center" into "Cedar_Ridge_Medical_Center", and INDIRECT finds the range with that
   name. `$H2` locks the column but not the row, so each row reads its own facility.

> ⚠️ If the parent cell is empty when you create the rule, Excel warns that "the source currently evaluates to an error."
> Click **Yes** to continue. The list works as soon as the parent has a value.

> ⚠️ **Changing the parent doesn't clear the child.** If someone switches row 10 from Bluestone Outpatient Pavilion to
> Bluestone Memorial Hospital, "Pediatric Clinic" stays in the Department cell even though it's no longer valid. Circle
> Invalid Data (next section) finds these mismatches. (Lesson 5.5 teaches Worksheet_Change event macros, which can clear the
> child automatically.)

> 💡 **Tip:** Keep parent values to plain words: letters and spaces, starting with a letter. SUBSTITUTE only swaps spaces,
> but Create from Selection also changes anything else a name can't contain (a leading digit, `&`, `-`), so the name and the
> SUBSTITUTE result would no longer match.

**Without named ranges.** The Lists sheet also holds every facility–department pair in long form: facilities in J2:J31 and
their departments in K2:K31, grouped by facility. When the parent values are grouped like this, OFFSET can cut the child list
straight out of the long list. OFFSET(start, rows, columns, height, width) returns a range that starts `rows` below `start`
and is `height` cells tall (you met it in Lesson 3.1's dynamic named ranges). For the Department column, the Source would be:

```
=OFFSET(Lists!$K$1, MATCH($H2,Lists!$J$2:$J$31,0), 0, COUNTIF(Lists!$J$2:$J$31,$H2), 1)
```

Take a row whose facility is Ashby Falls Community Hospital. Its first pair is in J17, the 16th cell of J2:J31, so MATCH
returns 16. COUNTIF returns 4 because Ashby Falls has four departments. OFFSET then starts 16 rows below K1, at K17, and
returns the 4 cells K17:K20. The start is the header cell K1, not K2, because MATCH counts from 1.

This version needs no names and no SUBSTITUTE, so it works for any facility name. The catch is that the long list must stay
grouped by parent. If one Ashby Falls pair sat at the bottom of the list, OFFSET would miss it. The bonus asks you to adapt
this formula to the huddle board.

> 📋 **Version note:** In Microsoft 365 and Excel 2021 or later, a single selector cell, such as a facility chosen in B4,
> has one more option. Put `=FILTER(Lists!$K$2:$K$31, Lists!$J$2:$J$31=$B$4)` in a helper cell (say X2) and set the child's
> Source to its spill range, `=$X$2#`. This only works for one parent cell, not for a whole column of rows. Lesson 4.1 covers
> FILTER and spill references.

### 7. Finding and fixing invalid data

Because validation only checks new typing, you need a way to audit what's already there:

1. Click the arrow on **Data → Data Validation** and choose **Circle Invalid Data**. Excel draws a red oval around every cell
   that breaks its own rule, whether the value was typed, pasted, or calculated.
2. Fix the circled entries. A circle disappears when its cell becomes valid.
3. Choose **Clear Validation Circles** when you're done. Circles also vanish when you save, and they never print.

Circle Invalid Data checks every validated cell on the sheet at once. After you've added rules to several columns, you'll
see circles in all of them, so count only the column you're auditing. Excel draws at most 255 circles at a time, so on a
large sheet fix the first batch and run it again.

> 📋 **Version note:** Circle Invalid Data is a desktop feature. Excel for the web doesn't have it. If you can't find it in
> your version (on a Mac, look under the arrow next to **Data Validation**), apply a temporary conditional formatting rule
> with the same test (Section 13), or type the equivalent formula from the answer key. Each one counts the same invalid cells.

Other validation housekeeping:

| Task | How |
|---|---|
| Find every validated cell | **Home → Find & Select → Data Validation**, or **Ctrl + G** (Mac: **Control + G**) → **Special** → **Data validation** (**All** or **Same**) |
| Copy a rule without copying values | Copy the cell, then **Paste Special** (Ctrl + Alt + V. Mac: Control + ⌘ + V) → **Validation** |
| Edit every cell that shares a rule | Select one cell, open **Data Validation**, change the rule, tick **Apply these changes to all other cells with the same settings** |
| Remove a rule | Select the cells → **Data Validation** → **Clear All** |

> ⚠️ Validation is a convenience, not security. Anyone can clear it, and Paste ignores it. For a shared log, combine it with
> sheet protection, and audit with Circle Invalid Data before you report from the data.

### 8. Conditional formatting basics

**Conditional formatting** applies a format (fill, font color, border, number format, bar, color, or icon) to a cell only
while a rule is TRUE for that cell. Excel re-checks every rule whenever values change, so the formatting stays current.

Open **Home → Conditional Formatting** (Windows: **Alt, H, L**). The menu has five families of rule, plus **New Rule**:

| Family | What it does | Typical use |
|---|---|---|
| **Highlight Cells Rules** | Formats cells that meet a simple test | TAT over 60 minutes, duplicate IDs |
| **Top/Bottom Rules** | Formats the highest or lowest values, or values above or below average | Ten slowest results |
| **Data Bars** | Draws a bar whose length shows the value | Stock level as % of par |
| **Color Scales** | Shades each cell by its position in the range | Occupancy heat map |
| **Icon Sets** | Adds an icon for each band of values | On time / at risk / late |
| **New Rule → Use a formula** | Formats when your own formula is TRUE | Whole row of a critical result |

**Highlight Cells Rules** cover the everyday tests: *Greater Than*, *Less Than*, *Between*, *Equal To*, *Text that
Contains*, *A Date Occurring*, and *Duplicate Values*. Each opens a small dialog with a value box and a preset format
(choose **Custom Format…** at the bottom of the list for your own).

Worked example: on the Labs sheet, select **L2:L322** (TATMin) and choose **Highlight Cells Rules → Greater Than**, type
60, and keep the light red fill. Excel highlights 80 results. Three results took exactly 60 minutes, and they stay plain
because *Greater Than* is strict. If your target is "60 minutes or less," a result at 60 met it, so *Greater Than* is the right test.

> 💡 **Tip:** Formatting doesn't filter or sort anything, but you can filter and sort by it. Click a column's filter arrow and
> choose **Filter by Color** or **Sort by Color** ([Lesson 1.6](../../01-foundations/06-sorting-filtering/README.md)). The
> menu lists every fill and font color your rules applied, and **Filter by Cell Icon** lists the icons from an icon set. The
> status bar then shows how many rows are left. The data sheets in this workbook are Tables, so their filter arrows are
> already on. (On a plain range, turn them on with **Ctrl + Shift + L**. Mac: **⌘ + Shift + F**. In a Table, the same
> shortcut turns them off.) A data bar isn't a fill color, so count bars with a formula such as COUNTIF instead.

> 📋 In Excel for Windows, selecting a range shows the **Quick Analysis** button (**Ctrl + Q**). Its **Formatting** tab
> previews data bars, color scales, and icon sets as you hover.

### 9. Top/Bottom rules, duplicates, and dates

**Top/Bottom Rules** rank values within the range:

| Rule | Formats | Example on TATMin (L2:L322) |
|---|---|---|
| Top 10 Items | The N largest values (N from 1 to 1000) | The 10 slowest results |
| Top 10 % | The largest N% of the cells | The slowest 10% |
| Bottom 10 Items / Bottom 10 % | The smallest values | The fastest results |
| Above Average | Values greater than the range's average | 130 results above the 49.3-minute mean |
| Below Average | Values less than the average | |

> ⚠️ **Ties are included.** If several cells tie with the 10th largest value, Top 10 Items highlights all of them, so you can
> see 11 or 12 highlighted cells. The smallest highlighted value is always `LARGE(range, 10)`.

**Duplicate Values** (under Highlight Cells Rules) formats every value that appears more than once. It highlights **every
copy, including the first one**. If one LabResultID appears three times, three cells turn red even though only two rows are
extra. Switch the first box to **Unique** to highlight values that appear exactly once instead. The comparison ignores case.

**A Date Occurring** offers *Yesterday*, *Today*, *Tomorrow*, *In the last 7 days*, *Last week*, *This week*, *Next week*,
*Last month*, *This month*, and *Next month*.

> ⚠️ Those options compare against your computer's clock, so the same workbook highlights different rows tomorrow. That's
> fine for a live worklist. A report with a fixed "as of" date, like this course's 12/31/2025, needs a formula rule that
> compares with a **ReportDate** cell instead (Section 13).

### 10. Data bars

**Data bars** draw a bar inside each cell. The longer the bar, the larger the value. Choose **Data Bars** and pick a
gradient or solid fill.

By default, the bar's **Minimum** and **Maximum** are **Automatic**: the longest bar belongs to the largest value in the
range. That's often not what you mean. On the Supplies sheet, PctOfPar is QtyOnHand ÷ ParLevel, and the largest value is
about 135%. With automatic limits, an item stocked exactly to par gets a bar only about three-quarters long, which reads as
"short."

To make bar length meaningful, fix the limits:

1. Select the range and add data bars.
2. **Conditional Formatting → Manage Rules**, select the rule, and click **Edit Rule**.
3. Set **Minimum** to Type **Number**, Value **0**, and **Maximum** to Type **Number**, Value **1**.

Now a half bar means 50% of par and a full bar means "at or above par." Values beyond the Maximum are drawn as full bars.
Tick **Show Bar Only** to hide the numbers. Negative values draw to the left of an axis (Excel 2010 and later).

### 11. Color scales

A **color scale** shades every cell in the range by where its value falls between a minimum and a maximum. A 2-color scale
blends between two colors. A 3-color scale adds a **midpoint** color.

The name of each preset lists its colors from **highest to lowest** value. **Green - Yellow - Red** paints the highest values
green, which suits "higher is better" measures like patient satisfaction. **Red - Yellow - Green** paints the highest values
red, which suits occupancy or wait times.

Data bars, color scales, and icon sets all describe their thresholds with the same **Type** options:

| Type | The threshold is… | Example |
|---|---|---|
| Lowest Value / Highest Value | The smallest or largest value in the range | Color-scale defaults |
| Number | A fixed value you type | 0.85, an occupancy target |
| Percent | A position between the lowest and highest value: lowest + p% × (highest − lowest) | 50 Percent of 84.4%–106.3% is 95.3% |
| Percentile | A rank: p% of the values fall below it | Percentile 50 is the median |
| Formula | The result of a formula | `=ReportDate` or `=$B$2` |

The default 3-color scale uses **Lowest Value**, **Percentile 50**, and **Highest Value**. Percentile 50 is the **median**,
so the yellow midpoint marks a typical day for this data, not a target. To color against a target such as 85% occupancy,
edit the rule and set the midpoint Type to **Number** and Value to **0.85**.

You can also build a scale from scratch with **New Rule → Format all cells based on their values**. Choose **2-Color Scale**
or **3-Color Scale** as the **Format Style**, then set the Type, Value, and Color of each point. It's the same dialog that
**Edit Rule** opens.

> ⚠️ Percent and Percentile sound alike but differ. *Percent* measures distance along the range from lowest to highest, so
> one extreme value stretches every threshold. *Percentile* ranks the values, so it ignores how far the extremes are.

> 💡 **Tip:** About 1 in 12 men has some red-green color deficiency. When red versus green carries the message, add a
> second cue: an icon, bold text, or a blue-to-orange scale.

### 12. Icon sets

An **icon set** puts a small symbol in each cell: arrows, traffic lights, flags, ratings, and more. Each icon covers one band
of values, and each band's test is written as **>= value** (or **> value**), checked from the top icon down.

The default 3-icon set uses **Percent 67** and **Percent 33**. Those split the span between the lowest and highest value
into thirds, so they rarely mean anything clinical. On the Labs sheet, TATMin runs from 15 to 137 minutes. The default green
light would go to every result of about 97 minutes or more, which is backwards for a turnaround time.

Worked example (task 11): select **L2:L322**, choose **Icon Sets → 3 Traffic Lights (Unrimmed)**, then **Manage Rules →
Edit Rule**:

1. Click **Reverse Icon Order** so the red light sits at the top (the highest values).
2. Red: **>=**, Value **60**, Type **Number**.
3. Yellow: **>=**, Value **45**, Type **Number**.
4. Green covers everything else, which is under 45 minutes.

| Option | Use it to |
|---|---|
| **Reverse Icon Order** | Put the "bad" icon on high values (waits, TAT, census) |
| **Type: Number** | Make thresholds literal minutes, dollars, or percentages |
| **> instead of >=** | Exclude the boundary value from the higher icon |
| **Show Icon Only** | Hide the number and keep a compact dashboard |
| Choosing **No Cell Icon** for a band | Show icons only for exceptions (Excel 2010 and later) |

### 13. Formula rules that format entire rows

The most flexible rule is **New Rule → Use a formula to determine which cells to format**. You write a formula that returns
TRUE or FALSE for the active cell, click **Format…** to choose the look, and Excel applies it wherever the formula is TRUE.

To color a whole row, apply the rule to every column and **lock the column, not the row**:

1. Select the data rows, for example **A2:L322** on the Labs sheet. The quickest way is to type `A2:L322` into the **Name Box**
   and press **Enter**, which also makes A2 the active cell.
2. **Home → Conditional Formatting → New Rule → Use a formula to determine which cells to format**.
3. Type the formula for row 2:

   ```
   =OR($K2="HH", $K2="LL")
   ```

4. Click **Format… → Fill**, pick a color, and click **OK** twice.

Every cell in row 5 now asks "is K5 HH or LL?" Row 5 holds a lactate of 4.4 mmol/L flagged HH, so the whole row turns red.

| You type | A cell in column C, row 5 tests | Result |
|---|---|---|
| `=$K2="HH"` | K5 | ✔ The whole row follows column K |
| `=K2="HH"` | M5, because the column shifts two places like a copied formula | ✘ Only scattered cells highlight |
| `=$K$2="HH"` | K2, for every cell | ✘ Every row copies row 2's answer |

More row rules on the lesson data:

| Highlight rows where… | Applies to | Formula |
|---|---|---|
| The unit's midnight census was above its staffed beds | Census!A2:F62 | `=$E2>$B2` |
| A STAT result missed a 60-minute target and was abnormal | Labs!A2:L322 | `=AND($L2>60, $K2<>"N")` |
| A supply has expired | Supplies!A2:M258 | `=AND($L2<>"", $L2<=ReportDate)` |

> ⚠️ **A blank cell counts as 0**, and 0 is the date 1/0/1900. So `=$L2<=ReportDate` is TRUE for every supply with no
> expiration date. Guard date tests with `$L2<>""`.

> ⚠️ **Structured references don't work inside rules.** A rule formula can't use `tblLabs[@AbnormalFlag]` or
> `tblLabs[TATMin]`. Use ordinary references such as `$K2`, a defined name such as `ReportDate`, or wrap a column reference in
> INDIRECT, as in `INDIRECT("tblLabs[TATMin]")`. Rules can refer to other sheets in Excel 2010 and later.

> 💡 **Tip:** Test a rule formula in a spare column first. Type it in the first row, fill it down, and check that TRUE appears
> on exactly the rows you expect. Then copy it into the rule and delete the test column.

> 💡 **Tip:** While you edit a rule formula, the arrow keys may insert cell references instead of moving the cursor. Press
> **F2** (Mac: **Fn + F2**) to switch to Edit mode. **F4** (Mac: **⌘ + T**) cycles a reference through `$K$2`,
> `K$2`, `$K2`, and `K2`.

### 14. Managing rules: order, conflicts, and Stop If True

**Home → Conditional Formatting → Manage Rules** (Windows: **Alt, H, L, R**) opens the Conditional Formatting **Rules
Manager**. Every "Manage Rules" step in this lesson means this dialog. Set **Show formatting rules for** to **This
Worksheet** to see every rule on the sheet, not just the ones on the selected cells.

| Column or button | What it does |
|---|---|
| **New Rule / Edit Rule / Delete Rule** | Create, change, or remove the selected rule |
| **Duplicate Rule** | Copy a rule so you can change only its formula or color |
| **▲ ▼** | Move a rule up or down. **Rules higher in the list win** |
| **Applies to** | The range the rule covers. Edit it here instead of re-creating the rule |
| **Stop If True** | When this rule is TRUE for a cell, skip every rule below it for that cell |

When two rules are TRUE for the same cell:

- If they set **different** properties (one bolds the font, the other fills the cell), the cell gets both.
- If they set the **same** property (two fills), the rule higher in the list wins.

> ⚠️ **Excel adds each new rule at the top of the list.** If you create the most urgent rule first and a milder rule second,
> the milder rule ends up on top and wins. Open the Rules Manager after you add related rules and move the most urgent one up.

Worked example (task 13): on the Supplies sheet, Rule 1 fills expired items red and Rule 2 fills items that expire within
90 days of ReportDate amber. Every expired item also expires before ReportDate + 90, so Rule 2 is TRUE on every expired row
as well. With the red rule on top, expired rows show red and only the items that are still good (but expiring soon) show
amber. If the amber rule is on top, every one of those rows turns amber and the expired stock disappears into the crowd.

You need **Stop If True** only when a lower rule sets a *different* property that you want to suppress, such as an icon set
or a bold font that shouldn't appear on rows that are already red. The checkbox is unavailable for data bars, color scales,
and icon sets themselves, but a formula rule above them can stop them.

> ⚠️ **Rules fragment.** Copying, pasting, and inserting rows can split one rule into many copies with odd Applies-to ranges,
> such as `$A$2:$L$40,$A$42:$L$322`. Open the Rules Manager now and then, delete the extra copies, and fix the range on the
> one you keep. Hundreds of fragments, whole-column ranges, and volatile functions such as TODAY() and INDIRECT in rules all
> slow a workbook down.

To remove rules, use **Conditional Formatting → Clear Rules → Clear Rules from Selected Cells** (or **from Entire Sheet**).
The Format Painter copies conditional formatting along with ordinary formatting.

### 15. Shortcuts and version notes

| Action | Windows | Mac |
|---|---|---|
| Data Validation dialog | Alt, A, V, V | **Data** tab → **Data Validation** |
| Open a drop-down list in a cell | Alt + ↓ | Option + ↓ |
| Conditional Formatting menu | Alt, H, L | **Home** tab → **Conditional Formatting** |
| Rules Manager | Alt, H, L, R | **Conditional Formatting → Manage Rules** |
| Create names from a selection | Ctrl + Shift + F3 | **Formulas → Create from Selection** |
| Name Manager | Ctrl + F3 | **Formulas → Name Manager** |
| Go To (then **Special**) | Ctrl + G or F5 | Control + G |
| Paste Special | Ctrl + Alt + V | Control + ⌘ + V |
| Toggle \$ in a reference | F4 | ⌘ + T |
| Turn the filter on or off | Ctrl + Shift + L | ⌘ + Shift + F |

| Feature | Version |
|---|---|
| Data bars, color scales, icon sets, Top/Bottom rules, unlimited rules per range | Excel 2007 and later |
| Rule and list formulas that refer to other sheets, negative data bars, solid bars, **No Cell Icon** | Excel 2010 and later |
| Circle Invalid Data | Excel for Windows. On a Mac, check the arrow next to **Data Validation**. Not in Excel for the web |
| Quick Analysis (Ctrl + Q) | Excel for Windows |
| Drop-down lists that filter as you type | Current Microsoft 365 |
| FILTER and spill references for list sources | Microsoft 365 and Excel 2021 or later |

## 🧪 Hands-on practice

Download [`3.2-data-validation-conditional-formatting.xlsx`](3.2-data-validation-conditional-formatting.xlsx) and open the
**Practice** sheet. Build each rule on the sheet the task names, then type what it shows into the yellow cell (a count or a
value, or a formula that calculates it). The **Check** column turns green when you're right.

<!-- BEGIN GENERATED: practice -->
Part A (tasks 1–6) uses the IntakeLog and Lists sheets. Add each validation rule, then click the arrow on **Data → Data Validation**, choose **Circle Invalid Data**, and count the red circles. Part B (tasks 7–13) uses the Labs, Supplies, and Census sheets. Create each conditional formatting rule, then answer the question about what it highlights. Type the number you count, or write a formula that calculates it.

| # | Task | Hint |
|:-:|------|------|
| 1 | IntakeLog, Payer column (F2:F51): add a List validation whose Source is the payer list on the Lists sheet (Lists!\$A\$2:\$A\$9). Then circle invalid data. How many Payer cells are circled? | **Data → Data Validation**, Allow: List |
| 2 | IntakeLog, VisitDate (B2:B51): the log covers December 2025 only. Allow a Date between 12/01/2025 and 12/31/2025 (type =DATE(2025,12,1) and =DATE(2025,12,31) in the boxes so it works in any regional setting). How many VisitDate cells are circled? | Allow: Date, Data: between |
| 3 | IntakeLog, CopayAmt (G2:G51): allow a Decimal between 0 and 100. On the Input Message tab, add the title Copay and the message "Enter the amount collected, \$0 to \$100." Leave the Error Alert style as Stop. How many CopayAmt cells are circled? | Allow: Decimal. Between includes both limits |
| 4 | IntakeLog, ZIP (E2:E51): allow Text length equal to 5. How many ZIP cells are circled? | Allow: Text length, Data: equal to |
| 5 | IntakeLog, MRN (C2:C51): an MRN must be exactly 8 characters, and all of them digits. Select the column with C2 as the active cell and add a Custom validation rule that is TRUE only for a valid MRN. How many MRN cells are circled? | Combine a LEN test with an ISNUMBER(--C2) test inside AND, written for the active cell C2 |
| 6 | Dependent drop-down: on the Lists sheet, name each department column after its facility (**Formulas → Create from Selection**, Top row, one column at a time: E1:E16, F1:F5, G1:G6, H1:H7). Give IntakeLog Facility (H2:H51) a List validation from Lists!\$C\$2:\$C\$5. Then give Department (I2:I51) a List validation whose Source is =INDIRECT(SUBSTITUTE(\$H2," ","_")). How many Department cells are circled? | Names can't contain spaces, so Create from Selection uses underscores |
| 7 | Labs sheet: the lab interface re-sent some results, so a few LabResultIDs appear more than once. Select A2:A322 and apply **Highlight Cells Rules → Duplicate Values**. How many cells are highlighted? | **Home → Conditional Formatting → Highlight Cells Rules** |
| 8 | Labs, TATMin (L2:L322): turnaround minutes from specimen collection to result. Apply **Top/Bottom Rules → Top 10 Items**. What is the smallest TAT that your rule highlights? Enter it in minutes. | Ties with the 10th value are highlighted too. LARGE gives the k-th largest |
| 9 | Supplies, PctOfPar (K2:K258) is QtyOnHand ÷ ParLevel. Add Data Bars, then edit the rule so Minimum is Type Number, Value 0 and Maximum is Type Number, Value 1. A full bar now means "stocked to par or above." How many items show a full bar? | Every value at or above the Maximum gets a full bar. Filter by Color can't see bars, but COUNTIF can count them |
| 10 | Census sheet (Medical-Surgical 5 East, Nov–Dec 2025): apply the Red - Yellow - Green Color Scale to Occupancy (F2:F62) so the fullest days are red. In **Manage Rules → Edit Rule**, you'll see the midpoint is the 50th percentile. Which occupancy gets the pure yellow midpoint color? Enter it as a percentage to 1 decimal place. | The 50th percentile has a more common name |
| 11 | Labs, TATMin (L2:L322): apply **Icon Sets → 3 Traffic Lights (Unrimmed)**. Edit the rule: click Reverse Icon Order, set both Types to Number, and make red show when the value is >= 60 and yellow when it is >= 45 (green below 45). How many cells show a yellow light? | Each icon's test is ">=". The default Type is Percent, not Number. **Filter by Color → Filter by Cell Icon** counts icons |
| 12 | Labs: highlight the entire row of every critical result (AbnormalFlag HH or LL). Select A2:L322 with A2 active, choose **New Rule → Use a formula to determine which cells to format**, and set a red fill. How many rows are highlighted? | Lock the column with \$, not the row |
| 13 | Supplies: add two formula rules to A2:M258. Rule 1 (red fill): the item has expired, meaning ExpirationDate is not blank and is on or before ReportDate. Rule 2 (amber fill): ExpirationDate is not blank and is on or before ReportDate + ExpiringWindowDays. ReportDate (12/31/2025) and ExpiringWindowDays (90) are named cells on the Lists sheet. Excel puts each new rule at the top of the list, so open Manage Rules and move the red rule above the amber rule. How many rows are amber? | A blank cell counts as 0, which is "on or before" any date. The rule higher in Manage Rules wins the fill |
<!-- END GENERATED: practice -->

## ✅ Answer key

The workbook has a hidden **Answer Key** sheet (right-click any sheet tab → **Unhide…** → *Answer Key*). Each solution lists
the steps and an equivalent formula, and the key's *Live result* column runs that formula so you can see the count is right.
The same answers are below, collapsed so you don't see them by accident.

<!-- BEGIN GENERATED: answers -->
<details>
<summary><b>🔑 Show the answer key</b> (try every task before you open this)</summary>

**1. Payer drop-down list (circled entries)**

- **Answer:** 5
- **Solution:**

1. On **IntakeLog**, select **F2:F51**.
2. **Data → Data Validation** (Windows: Alt, A, V, V). On the **Settings** tab set **Allow** to **List**.
3. Click in **Source**, switch to the **Lists** sheet, and select **A2:A9**. The box reads `=Lists!$A$2:$A$9`. Click **OK**.
4. Click the arrow on **Data → Data Validation**, then **Circle Invalid Data**. (**Clear Validation Circles** removes them.) You see **5** circles.

Equivalent formula: `=SUMPRODUCT(--(COUNTIF(Lists!$A$2:$A$9,IntakeLog!F2:F51)=0))`


A validation rule only checks values as someone types them, so the bad payers already in the log stay until you circle them. Each one is a near miss of a real payer: "Self Pay" instead of "Self-Pay", "Medicaid" instead of "State Medicaid". A drop-down list prevents exactly this. The formula counts entries whose COUNTIF in the payer list is 0, which means they aren't on it.

**2. December-only VisitDate rule (circled entries)**

- **Answer:** 4
- **Solution:**

1. Select **B2:B51** on IntakeLog.
2. **Data → Data Validation → Settings**: **Allow** = **Date**, **Data** = **between**, **Start date** = `=DATE(2025,12,1)`, **End date** = `=DATE(2025,12,31)`. Click **OK**.
3. Click the arrow on **Data → Data Validation**, then **Circle Invalid Data**. (**Clear Validation Circles** removes them.) You see **4** circles.

Equivalent formula: `=ROWS(IntakeLog!B2:B51)-COUNTIFS(IntakeLog!B2:B51,">="&DATE(2025,12,1),IntakeLog!B2:B51,"<="&DATE(2025,12,31))`


Two entries are text that only looks like a date: 12/32/2025 doesn't exist, and 19/12/2025 was imported day-first. Both sit at the left of the cell, which is the clue that they're text. A Date rule rejects text automatically. The other two are real dates with the wrong year (2026 and 2015). COUNTIFS skips text, so "all rows minus rows with a December 2025 date" counts every bad entry.

**3. Copay amount rule with an input message (circled entries)**

- **Answer:** 3
- **Solution:**

1. Select **G2:G51**.
2. **Data → Data Validation → Settings**: **Allow** = **Decimal**, **Data** = **between**, **Minimum** = 0, **Maximum** = 100.
3. **Input Message** tab: Title `Copay`, message `Enter the amount collected, $0 to $100.`
4. **Error Alert** tab: keep **Style** = **Stop**. Click **OK**.
5. Click the arrow on **Data → Data Validation**, then **Circle Invalid Data**. (**Clear Validation Circles** removes them.) You see **3** circles.

Equivalent formula: `=COUNTIF(IntakeLog!G2:G51,"<0")+COUNTIF(IntakeLog!G2:G51,">100")`


"Between" includes both limits, so the \$0.00 and \$100.00 copays in the log are valid. The circled values are a negative amount (a refund typed into the wrong field), 2,500 (almost certainly \$25.00 typed without the decimal point), and 350. Click any cell in the column to see your input message.

**4. ZIP code length rule (circled entries)**

- **Answer:** 3
- **Solution:**

1. Select **E2:E51**.
2. **Data → Data Validation → Settings**: **Allow** = **Text length**, **Data** = **equal to**, **Length** = 5. Click **OK**.
3. Click the arrow on **Data → Data Validation**, then **Circle Invalid Data**. (**Clear Validation Circles** removes them.) You see **3** circles.

Equivalent formula: `=SUMPRODUCT(--(LEN(IntakeLog!E2:E51)<>5))`


Text length counts characters, so the 4-digit ZIP, the 6-digit ZIP, and the ZIP+4 entry (10 characters including the hyphen) fail. The rule also works on numbers. One ZIP in the log was stored as the number 45502 (it's right-aligned), and it passes because it has 5 characters.

**5. Custom MRN rule (circled entries)**

- **Answer:** 4
- **Solution:**

1. Type **C2:C51** in the Name Box and press **Enter**. The range is selected and **C2** is the active cell.
2. **Data → Data Validation → Settings**: **Allow** = **Custom**, **Formula** = `=AND(LEN(C2)=8,ISNUMBER(--C2))`. Click **OK**.
3. Click the arrow on **Data → Data Validation**, then **Circle Invalid Data**. (**Clear Validation Circles** removes them.) You see **4** circles.

Equivalent formula: `=ROWS(IntakeLog!C2:C51)-SUMPRODUCT((LEN(IntakeLog!C2:C51)=8)*ISNUMBER(--IntakeLog!C2:C51))`


LEN catches the 7- and 9-character MRNs and the MRN that was stored as a number (Excel dropped its leading zero, leaving 7 digits). ISNUMBER(--C2) catches O9127341, where the letter O replaced a zero: the double minus tries to turn the text into a number and fails. Each test misses something the other catches, so both go inside AND. You write the formula for C2 only, and Excel adjusts it for C3, C4, and so on.

**6. Dependent Facility → Department list (circled entries)**

- **Answer:** 4
- **Solution:**

1. On **Lists**, select **E1:E16** and choose **Formulas → Create from Selection**, tick only **Top row**, and click **OK**. Repeat for **F1:F5**, **G1:G6**, **H1:H7**. **Formulas → Name Manager** (Windows: Ctrl + F3) now includes four new names, such as `Cedar_Ridge_Medical_Center`.
2. Select **H2:H51** on IntakeLog → **Data Validation** → **List**, Source `=Lists!$C$2:$C$5`.
3. Select **I2:I51** with **I2** active → **Data Validation** → **List**, Source `=INDIRECT(SUBSTITUTE($H2," ","_"))`. Click **OK**.
4. Click the arrow on **Data → Data Validation**, then **Circle Invalid Data**. (**Clear Validation Circles** removes them.) You see **4** circles.

Equivalent formula (uses the facility–department map in Lists!J:K): `=SUMPRODUCT(--(COUNTIFS(Lists!$J$2:$J$31,IntakeLog!H2:H51,Lists!$K$2:$K$31,IntakeLog!I2:I51)=0))`


Create from Selection turns the header "Cedar Ridge Medical Center" into the name Cedar_Ridge_Medical_Center, because names can't contain spaces. SUBSTITUTE makes the same change to the facility in column H, and INDIRECT turns that text into a reference to the named range. Because \$H2 has no \$ before the row, every row builds its list from its own facility. The circled entries are real departments at the wrong facility: Pediatric Clinic, for example, is at the Outpatient Pavilion, not at Bluestone Memorial Hospital.

**7. Duplicate LabResultIDs (highlighted cells)**

- **Answer:** 13
- **Solution:**

1. On **Labs**, select **A2:A322**.
2. **Home → Conditional Formatting → Highlight Cells Rules → Duplicate Values**. Keep **Duplicate** and the light red fill, then click **OK**.
3. Count the highlighted cells (or filter the column by color and read the status bar): **13**.

Equivalent formula: `=SUMPRODUCT(--(COUNTIF(Labs!A2:A322,Labs!A2:A322)>1))`


Duplicate Values highlights every copy of a repeated value, including the first one. There are 6 repeated IDs: five appear twice and one appears three times, so 13 cells light up even though only 7 rows are extra. When you clean the feed, you delete 7 rows, not 13.

**8. Top 10 turnaround times (smallest highlighted value)**

- **Answer:** 99
- **Solution:**

1. Select **L2:L322**.
2. **Home → Conditional Formatting → Top/Bottom Rules → Top 10 Items**. Keep **10**, click **OK**.
3. Sort or filter by color to read the smallest highlighted value.

Equivalent formula: `=LARGE(Labs!L2:L322,10)`


Top 10 Items highlights the 10 largest values plus any cell tied with the 10th, so a tie at the boundary can light up 11 or 12 cells. Here there's no tie at the boundary, so exactly 10 cells are highlighted. Either way, the smallest highlighted value is LARGE(range,10).

**9. Data bars on % of par (full bars)**

- **Answer:** 78
- **Solution:**

1. On **Supplies**, select **K2:K258**.
2. **Home → Conditional Formatting → Data Bars** and pick any fill.
3. **Conditional Formatting → Manage Rules → Edit Rule**. Set **Minimum** Type = **Number**, Value = 0, and **Maximum** Type = **Number**, Value = 1. Click **OK** twice.
4. A data bar isn't a fill color, so Filter by Color can't find the full ones. Count them with the formula below instead, because every value of 1 (100%) or more gets a full bar: **78**.

Equivalent formula: `=COUNTIF(Supplies!K2:K258,">=1")`


With the default Automatic settings, the longest bar belongs to the largest value (about 135% of par), so an item stocked exactly to par looks only about three-quarters full. Fixing the Maximum at 1 gives bar length a meaning: every item at or above 100% of par gets a full bar, and an item at 50% gets a half bar.

**10. Color scale midpoint (50th percentile)**

- **Answer:** 90.6%
- **Solution:**

1. On **Census**, select **F2:F62**.
2. **Home → Conditional Formatting → Color Scales → Red - Yellow - Green Color Scale**. The first color in the name goes to the highest values, so high occupancy is red.
3. **Conditional Formatting → Manage Rules → Edit Rule** shows Minimum = Lowest Value, Midpoint = Percentile 50, Maximum = Highest Value.
4. Percentile 50 is the median, so the formula below gives the occupancy that gets pure yellow.

Equivalent formula: `=MEDIAN(Census!F2:F62)`


The 50th percentile is the median, the middle value when you sort the days by occupancy. Here 17 days have exactly that occupancy, so they all show pure yellow. The default color scale's middle color therefore marks a typical day for this unit, not a target. To color against a target such as 85%, change the midpoint Type to Number (you'll do that in the bonus).

**11. Traffic-light icons on turnaround time (yellow count)**

- **Answer:** 79
- **Solution:**

1. Select **L2:L322** → **Home → Conditional Formatting → Icon Sets → 3 Traffic Lights (Unrimmed)**.
2. **Manage Rules → Edit Rule**. Click **Reverse Icon Order** so red is on top.
3. Red: **>=**, Value **60**, Type **Number**. Yellow: **>=**, Value **45**, Type **Number**. Green covers everything below 45. Click **OK** twice.
4. Click the TATMin filter arrow → **Filter by Color** → **Filter by Cell Icon** → the yellow light, and read the count on the status bar: **79**. Clear the filter afterward.

Equivalent formula: `=COUNTIFS(Labs!L2:L322,">=45",Labs!L2:L322,"<60")`


The default thresholds are Percent 67 and 33. They split the span between the lowest and highest TAT into thirds, so they don't mean minutes at all. Switching the Type to Number makes the thresholds literal. Reverse Icon Order puts red on the high (slow) values. Yellow covers 45 ≤ TAT < 60 because each icon's test is ">=" and the red test is checked first. (83 results are red.)

**12. Critical results: whole-row formula rule**

- **Answer:** 14
- **Solution:**

1. On **Labs**, type **A2:L322** in the Name Box and press **Enter**. The rows are selected and A2 is the active cell.
2. **Home → Conditional Formatting → New Rule → Use a formula to determine which cells to format**.
3. Formula: `=OR($K2="HH",$K2="LL")`. Click **Format → Fill**, pick red, then **OK** twice.
4. Filter any column from B to K by that red fill (no earlier rule colors those columns) and read the count on the status bar: **14**.

Equivalent formula: `=COUNTIF(Labs!K2:K322,"HH")+COUNTIF(Labs!K2:K322,"LL")`


You write the rule for the active cell's row (row 2), and Excel shifts it for every other cell in the Applies-to range. \$K keeps every cell in the row looking at column K, and the unlocked 2 lets each row check its own flag. Without the \$, cell B2 would test L2 and the rule breaks. There are 12 HH and 2 LL results, so 14 rows turn red.

**13. Expired vs expiring supplies: two rules and rule order**

- **Answer:** 7
- **Solution:**

1. On **Supplies**, select **A2:M258** with A2 active.
2. **New Rule → Use a formula**: `=AND($L2<>"",$L2<=ReportDate)` with a red fill → **OK**.
3. With the same range selected, add a second rule: `=AND($L2<>"",$L2<=ReportDate+ExpiringWindowDays)` with an amber (light orange) fill → **OK**.
4. **Conditional Formatting → Manage Rules**, set **Show formatting rules for: This Worksheet**. The amber rule is on top because Excel adds each new rule at the top of the list. Select the red rule and click **▲** (Move Up) so it sits above the amber rule, then click **OK**.
5. Filter column A by the amber fill and read the count on the status bar: **7**.

Equivalent formula: `=COUNTIFS(Supplies!L2:L258,">"&ReportDate,Supplies!L2:L258,"<="&ReportDate+ExpiringWindowDays)`


Rule 2 is TRUE for 17 rows, because every expired item also expires before ReportDate + 90. The red rule sits above it and both rules set a fill, so red wins on the 10 expired rows and 7 rows stay amber. If you skip the move, the newer amber rule stays on top and you'd see 17 amber rows and no red. The \$L2<>"" guard matters too: an empty cell counts as 0 (the date 1/0/1900), which is "on or before ReportDate," so without the guard all 121 non-perishable items would turn red.

</details>
<!-- END GENERATED: answers -->

## 🏆 Bonus challenge

<!-- BEGIN GENERATED: bonus -->
It's 7 a.m. on Wednesday, November 26, 2025, the day before Thanksgiving. The house supervisor runs the system bed huddle from the Huddle sheet: one row per inpatient unit, with the midnight census for 11/24 and 11/25, the overnight change, and occupancy. Make the board readable at a glance, then build the unit selector that drives the gray unit card (Huddle!B6:B9). The card reads the NovCensus sheet (every unit, Nov 1–25). The board's data is in rows 12–29.

Build every rule and the selector on the **Huddle** sheet, and type your answers to B1–B3 in the yellow cells on the **Bonus** sheet.

- **B1.** On Huddle, apply a 3-color scale to Occupancy (G12:G29): Minimum = Lowest Value (green), Midpoint = Number 0.85 (yellow, the planning target), Maximum = Highest Value (red). How many units are shaded on the red side of yellow (occupancy above 85%)? (Hint: **New Rule → Format all cells based on their values**, Format Style 3-Color Scale)
- **B2.** On Huddle, apply **Icon Sets → 3 Arrows (Colored)** to Change (F12:F29). Edit the rule so both Types are Number: up arrow when the value is >= 1, sideways arrow when it is >= 0, down arrow otherwise. How many units show an up arrow? *(Hint: Type = Number, not the default Percent)*
- **B3.** On Huddle, add two formula rules to the whole board (A12:G29): a red fill when the 11/25 census is above staffed beds (overflow), and an amber fill when occupancy is at least 90%. In Manage Rules, put the overflow rule above the amber rule. How many rows end up amber? *(Hint: Both rules are TRUE for an overflow unit, and the rule on top wins the fill)*
- **B4.** On Huddle, build the unit selector. Give B4 a List validation from the Hospitals list (I12:I14). Give B5 a dependent List built from the board's own Facility and Unit columns, with no named ranges. Then pick Cedar Ridge Medical Center in B4 and Intensive Care Unit in B5. The gray answer cell on the Bonus sheet copies the card's Nov 1–25 occupancy (Huddle!B8), so it fills in and turns ✔ when your selector works. *(Hint: Adapt the OFFSET example in guide section 6: MATCH finds the facility's first row and COUNTIF its height)*
<!-- END GENERATED: bonus -->

When you finish, unhide the **Huddle Key** sheet to compare your board with a finished one. It has every rule and both
drop-down lists in place.

<!-- BEGIN GENERATED: bonus-answers -->
<details>
<summary><b>🔑 Show the bonus solution</b> (give it a real try first)</summary>

**B1. Color scale with an 85% target midpoint**

- **Answer:** 9
- **Solution:**

1. Select **G12:G29** on Huddle → **Conditional Formatting → New Rule → Format all cells based on their values**.
2. **Format Style** = **3-Color Scale**. Minimum: Lowest Value, green. Midpoint: Type **Number**, Value **0.85**, yellow. Maximum: Highest Value, red. Click **OK**.

Equivalent formula: `=COUNTIF(Huddle!G12:G29,">0.85")`


With a Number midpoint, yellow means "exactly at target" instead of "a typical unit." Every unit above 85% shades from yellow toward red, and the reddest cell is simply the fullest unit.

**B2. Arrows on the overnight census change**

- **Answer:** 7
- **Solution:**

1. Select **F12:F29** → **Conditional Formatting → Icon Sets → 3 Arrows (Colored)**.
2. **Manage Rules → Edit Rule**: green up arrow **>=** 1 **Number**, yellow sideways arrow **>=** 0 **Number**. The red down arrow covers negative changes. Click **OK** twice.

Equivalent formula: `=COUNTIF(Huddle!F12:F29,">=1")`


Census rose overnight on 7 units, held steady on 10, and fell on 1. Number thresholds keep the arrows honest on any day. The default Percent thresholds depend on the day's smallest and largest change, so a unit that gained one patient could show a sideways arrow on a busier day.

**B3. Overflow (red) above near-capacity (amber) row rules**

- **Answer:** 3
- **Solution:**

1. Select **A12:G29** with A12 active.
2. **New Rule → Use a formula**: `=$E12>$C12`, red fill.
3. **New Rule → Use a formula**: `=$G12>=0.9`, amber fill.
4. **Manage Rules** (This Worksheet): Excel added the amber rule at the top because it's newer. Select the overflow rule and click **▲** so it's on top.

Equivalent formula: `=SUMPRODUCT((Huddle!G12:G29>=0.9)*(Huddle!E12:E29<=Huddle!C12:C29))`


There are 2 overflow units, and both rules are TRUE for them. The overflow rule is on top, so those rows stay red, and 3 rows are amber. Two of those amber units sit at exactly 90.0%, so writing > 0.9 instead of >= 0.9 would lose them. With the rules in the wrong order every overflow row would turn amber too.

**B4. Dependent unit selector driving the card**

- **Answer:** 83.7%
- **Solution:**

1. Select **B4** on Huddle → **Data Validation → List**, Source `=$I$12:$I$14`.
2. Select **B5** → **Data Validation → List**, Source:

   `=OFFSET($B$11,MATCH($B$4,$A$12:$A$29,0),0,COUNTIF($A$12:$A$29,$B$4),1)`

   MATCH finds the facility's first row on the board, and COUNTIF counts its units. The board lists each facility's units together, so OFFSET returns exactly that facility's block of unit names. If B4 is still empty, Excel warns that the source currently evaluates to an error. Click **Yes**, or pick a facility in B4 first.
3. Pick **Cedar Ridge Medical Center** in B4, then **Intensive Care Unit** in B5. The card fills in, and the gray answer cell on the Bonus sheet shows its occupancy.

The guide's version starts at Lists!K1 and uses \$H2, because every intake row has its own facility. Here the start is the board's Unit header (B11) and the parent is the single cell \$B\$4.

An INDIRECT version would need a second set of named lists that hold only the inpatient units. The names from task 6 (such as `Cedar_Ridge_Medical_Center`) already point at every department on the Lists sheet, so new names would have to differ from them. OFFSET skips that setup because it reads the board itself.


Three hospitals each have an Intensive Care Unit, so the unit name alone is ambiguous. The dependent list forces a facility first and then offers only that facility's units, and the card's SUMIFS uses both cells. The OFFSET version needs no named ranges. If you insert a new unit's row inside its facility's block, the board ranges in the formula grow to include it, so the list keeps working.

</details>
<!-- END GENERATED: bonus-answers -->

## Key takeaways

- Data validation stops bad entries as they're typed, and conditional formatting makes what matters visible. Neither one
  changes a stored value.
- Select the whole range, then write the rule for the active cell. Excel adjusts it for every other cell, just like a copied
  formula.
- Validation doesn't check data that was already there or data that arrives by paste. Audit with **Circle Invalid Data** (or
  an equivalent formula).
- A dependent drop-down pairs named ranges with `INDIRECT(SUBSTITUTE(parent," ","_"))`, or uses OFFSET to cut the child list
  out of a long list grouped by parent. Changing the parent doesn't clear the child, so audit it too.
- Fix thresholds with **Type: Number** when they mean something (par level, an 85% target, a 60-minute TAT). The defaults
  (Automatic, Percent, Percentile) only describe the spread of today's data.
- To color whole rows, lock the column and leave the row free (`$K2`), and guard date tests against blanks.
- Rule order matters. The higher rule wins when two rules set the same format, so put the most urgent rule on top.

<!-- BEGIN GENERATED: nav -->
---

⬅️ **Previous:** [3.1 Excel Tables, Structured References & Named Ranges](../01-tables-named-ranges/README.md) · 🏠 [Course home](../../README.md) · **Next:** [3.3 Cleaning Messy Data](../03-data-cleaning/README.md) ➡️
<!-- END GENERATED: nav -->
