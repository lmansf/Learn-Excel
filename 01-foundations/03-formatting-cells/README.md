# Lesson 1.3 · Formatting Cells & Number Formats

> **Level:** Beginner · **Time:** about 45 minutes · **Workbook:** [`1.3-formatting-cells.xlsx`](1.3-formatting-cells.xlsx)
> **Data:** Bluestone Memorial Hospital's 2025 expense budget vs actual by department, two days of ED registrations (12/30–12/31/2025), December 2025 discharges from Cardiac Step-Down, and the Emergency Department's December 2025 budget report (bonus). Every number arrives unformatted, the way it comes out of a source system.

Finance exports the 2025 budget, and Perioperative Services' variance arrives as `-290135` with a variance percentage of
`-0.015859001`. Both numbers are correct, yet a busy reader has to decode them. Formatted, the same cells read `($290K)` in red
and `-1.6%`, and the director of surgery sees the overspend in a second. Formatting is how you make a spreadsheet usable by a CFO,
a nurse manager, or a quality committee: dollars in thousands, overspending in red parentheses, lengths of stay in hours, and MRNs
with all eight digits so staff can find the right chart. In this lesson you'll turn raw exports from Bluestone Memorial Hospital
into presentation-ready reports, guided by one rule that keeps you safe: **formatting changes how a value looks, never the value
itself.**

## What you'll learn

- Apply number formats: Number, Currency vs Accounting, Percentage, Date, Time, Text
- Write custom number formats (leading zeros, units, colors, thousands)
- Format fonts, fills, borders, alignment, and wrap text — and avoid merged cells
- Use Format Painter, cell styles, and themes for consistent reports
- Understand that formatting changes how a value looks, not the value itself

## 📖 Guide

### 1. What you see vs. what Excel stores

Every cell has two faces. The **stored value** is the number Excel keeps and calculates with. The **displayed value** is what you
see in the cell after Excel applies the cell's **number format**, a set of display rules such as "two decimals" or "show as a
percentage." The formula bar always shows the stored value, so it's where you go to find out what a cell really holds.
The table shows cells from this lesson's workbook with formats you'll apply during the practice.

| Cell in this lesson's workbook | Stored value (formula bar) | Number format | The cell displays |
|---|---|---|---|
| Budget!C7, Radiology's budget | 9165952 | `$#,##0,"K"` | `$9,166K` |
| Budget!F7, Radiology's variance % | 0.0281259383 | `0.0%` | `2.8%` |
| Registry!B6, an MRN | 1927734 | `00000000` | `01927734` |
| Stays!E8, a length of stay in days | 4.8243055… | `0.0` | `4.8` |
| Stays!E8, the same cell | 4.8243055… | `[h]:mm` | `115:47` |

The last two rows are the same cell with two different formats. Nothing about the stay changed, only the way Excel describes it.

**Calculations always use the stored value.** On the Budget sheet, once every budget shows in thousands, add up the 16
department budgets *as displayed* and you get 150,080. The Total row holds the true total of the stored values, 150,078,282,
so it shows `$150,078K`, and a SUM of the budget column would return the same 150,078,282. Neither number is wrong. The
displayed values were each rounded, and the small differences add up. When a report needs the
rounded number itself, round the value with `ROUND` (Lesson 1.4) instead of relying on the format.

> 💡 **Tip:** To see a cell's raw stored value, apply the **General** format with **Ctrl + Shift + ~** (Mac: **⌃ + Shift + ~**),
> then press **Ctrl + Z** (Mac: **⌘ + Z**) to put the format back.

> ⚠️ **Don't turn on "Set precision as displayed."** It's under **File → Options → Advanced** (Mac: **Excel → Settings →
> Calculation**; Preferences in older versions), and it permanently rounds every stored value in the workbook to what's displayed. The lost digits can't be
> recovered, so leave it off.

**General** is the format every new cell starts with. It shows a number with as many digits as fit in the column, with no
thousands separators, and it switches to scientific notation (`1.23457E+11`) for numbers with 12 or more digits. If a formatted
number or date doesn't fit, Excel shows `#####` instead of a misleading partial number. That isn't an error in the value, so widen
the column (section 14).

### 2. Where number formats live

You'll use two places, plus keyboard shortcuts (section 3).

**The Number group on the Home tab** holds the everyday tools:

| Control | What it does |
|---|---|
| **Number Format** box (shows *General* at first) | Lists 11 common formats. The box always names the active cell's format. |
| **$** Accounting Number Format | Applies Accounting with two decimals. Its arrow offers other currency symbols. |
| **%** Percent Style | Applies `0%`. |
| **,** Comma Style | Applies Accounting without a currency symbol. |
| Increase Decimal / Decrease Decimal | Adds or removes one displayed decimal place per click (Windows KeyTips **Alt, H, 0** and **Alt, H, 9**). |
| The small arrow at the group's bottom-right corner | Opens the Format Cells dialog. |

**The Format Cells dialog** has every option. Open it with **Ctrl + 1** (Mac: **⌘ + 1**), or right-click a cell and choose
**Format Cells…**. It has six tabs:

| Tab | Controls |
|---|---|
| **Number** | Number formats, including Custom (section 8). The **Sample** box previews the active cell. |
| **Alignment** | Horizontal and vertical alignment, wrap text, shrink to fit, merge, indent, text orientation. |
| **Font** | Font, style, size, underline, color, strikethrough. |
| **Border** | Line style and color for each edge of the selection. |
| **Fill** | Background color and patterns. |
| **Protection** | Locked and hidden settings, which only matter once you protect the sheet. |

### 3. The built-in number formats

| Format | What it shows | Example: stored → displayed |
|---|---|---|
| **General** | As many digits as fit, no separators | 9165952 → `9165952` |
| **Number** | A fixed number of decimals (two at first), an optional thousands separator (the *Use 1000 Separator* box in Format Cells), and a choice of negative styles | 9165952 → `9165952.00`, or `9,165,952.00` with the separator |
| **Currency** | A currency symbol right next to the number | 257801 → `$257,801.00` |
| **Accounting** | A currency symbol pinned to the cell's left edge, aligned decimals, zero as a dash | see section 4 |
| **Date** | Built-in date patterns. *Short Date* follows your computer's region | 45992 → `12/1/2025` |
| **Time** | Built-in time patterns | 0.3916667 → `9:24 AM` |
| **Percentage** | The number × 100, with a % sign | 0.0281259 → `2.81%` |
| **Fraction** | The decimal part as a fraction | 2.75 → `2 3/4` |
| **Scientific** | Powers of ten | 9165952 → `9.17E+06` |
| **Text** | Treats the entry as text, exactly as typed (section 7) | `00412345` stays `00412345` |
| **Special** | Region-specific patterns. US options include Zip Code, Zip Code + 4, Phone Number, and Social Security Number | 5551234567 → `(555) 123-4567` |
| **Custom** | Any format you can write as a code (section 8) | 9165952 → `$9,166K` |

The fastest way to apply the common ones is the keyboard. The symbols `~ ! @ # $ % ^` are the shifted keys along the top row,
from the key left of 1 through 6, which makes them easy to remember.

| Windows | Mac | Applies | Example |
|---|---|---|---|
| **Ctrl + Shift + ~** | **⌃ + Shift + ~** | General | `9165952` |
| **Ctrl + Shift + !** | **⌃ + Shift + !** | Number: two decimals, thousands separator | `9,165,952.00` |
| **Ctrl + Shift + @** | **⌃ + Shift + @** | Time: hour, minute, AM/PM | `9:24 AM` |
| **Ctrl + Shift + #** | **⌃ + Shift + #** | Date: day, month, year | `1-Dec-25` |
| **Ctrl + Shift + $** | **⌃ + Shift + $** | Currency: two decimals, negatives in parentheses | `$257,801.00` |
| **Ctrl + Shift + %** | **⌃ + Shift + %** | Percentage with no decimals | `3%` |
| **Ctrl + Shift + ^** | **⌃ + Shift + ^** | Scientific with two decimals | `9.17E+06` |
| **Ctrl + 1** | **⌘ + 1** | Opens Format Cells | |

> 📋 On a Mac, **⌃ + Shift + $** also shows negative amounts in red. On both platforms you can press the shortcut with the whole
> column selected, so new numbers typed into it pick up the format too.

### 4. Currency vs. Accounting

Both formats show money, and both can show the same number of decimals. They differ in layout, which is why finance reports
use Accounting and price lists or single amounts use Currency.

| | **Currency** | **Accounting** |
|---|---|---|
| Where the $ sits | Right next to the number: `$212.08` | Pinned to the left edge of the cell, with the number at the right |
| A column of amounts | The $ signs form a ragged edge, because each one sits next to its own number | The $ signs form a straight line on the left, and the decimal points line up |
| Zero | `$0.00` | A dash |
| Negative numbers | You choose: `-$1,234.10`, red, `($1,234.10)`, or red with parentheses | Always in parentheses |
| How to apply | **Ctrl + Shift + $** (Mac: **⌃ + Shift + $**) or the Number Format box | The Accounting Number Format button (the dollar-sign button on the Home tab), or the Number Format box |
| Typical use | A copay on a registration form, a price per unit | Budget reports, financial statements, any column of money you'll total |

> ⚠️ **The $ button is not Currency.** The big **$** button on the Home tab applies **Accounting**. The keyboard shortcut
> **Ctrl + Shift + $** applies **Currency**. Check the Number Format box if you're not sure which one a cell has.

**Comma Style** (the **,** button) is Accounting without the dollar sign. It's a good choice for counts and dollar columns in a
table where only the first row and the total show a $ sign, a common layout in hospital financial statements.

### 5. Percentages

A **percentage** format multiplies the stored value by 100 for display and adds a % sign. Radiology's variance % is stored as
0.0281259383, so `0%` shows `3%`, `0.0%` shows `2.8%`, and `0.00%` shows `2.81%`. Excel rounds only the display.

Excel handles typing in two ways:

- In a cell that's still General, typing `2.8%` stores 0.028 and applies a percentage format for you.
- In a cell that's *already* formatted as a percentage, typing `2.8` also stores 0.028. This is called *automatic percent
  entry*, and it's on by default (**File → Options → Advanced → Enable automatic percent entry**; Mac: **Excel → Settings →
  Edit**).

> ⚠️ **Percent of what?** The percentage format assumes the cell holds a fraction. If a report already stores readmission rates as
> whole percentages, such as 14.2 for 14.2%, applying `0.0%` shows `1420.0%`. Either divide those numbers by 100 first (Lesson 1.2's
> Paste Special → Divide does it in one step) or keep them as plain numbers with a header like "Readmission rate (%)".

### 6. Dates, times, and durations

Lesson 1.2 showed that a date is a **serial number** (12/01/2025 is 45992) and a time is a fraction of a day (09:24 is 0.391667).
Formatting decides how that number reads. The Stays sheet's Admitted column arrives as raw serial numbers, so you'll format it
yourself in task 8.

The **date and time codes** below can be combined in any order with spaces, slashes, commas, colons, and hyphens. The examples show
Stays!C7, which stores 45992.391667 (Monday, 12/01/2025 at 09:24).

| Code | Shows | Example |
|---|---|---|
| `d` / `dd` | Day of the month without / with a leading zero | `1` / `01` |
| `ddd` / `dddd` | Short / full day name | `Mon` / `Monday` |
| `m` / `mm` | Month number without / with a leading zero | `12` / `12` (in March: `3` / `03`) |
| `mmm` / `mmmm` / `mmmmm` | Short month name / full name / first letter | `Dec` / `December` / `D` |
| `yy` / `yyyy` | Two-digit / four-digit year | `25` / `2025` |
| `h` / `hh` | Hour without / with a leading zero (24-hour clock unless AM/PM is in the code) | `9` / `09` |
| `m` / `mm` *after h or before s* | Minutes | `24` |
| `s` / `ss` | Seconds | `0` / `00` |
| `AM/PM` | 12-hour clock with AM or PM | `AM` |
| `[h]` / `[m]` | **Elapsed** hours / minutes that keep counting past 24 hours / 60 minutes | see below |

Full codes on the same cell:

| Format code | Displays |
|---|---|
| `mm/dd/yyyy` | `12/01/2025` |
| `d-mmm-yy` (what Ctrl + Shift + # applies) | `1-Dec-25` |
| `dddd, mmmm d, yyyy` | `Monday, December 1, 2025` |
| `mmm yyyy` | `Dec 2025` |
| `yyyy-mm-dd` | `2025-12-01` |
| `h:mm AM/PM` | `9:24 AM` |
| `mm/dd/yyyy hh:mm` | `12/01/2025 09:24` |

> ⚠️ **m means month or minute.** Excel reads `m` and `mm` as minutes only when they come right after `h`/`hh` or right before
> `ss`. Everywhere else they mean month. So `hh:mm` is hours and minutes, but `mm:hh` shows the month first. Use `mm` rather than `m` for
> minutes, so 9:05 doesn't display as `9:5`.

**Durations.** A length of stay, an ED boarding time, or a shift's worked hours is a *duration*, not a time of day. Ordinary hour
codes roll back to 0 every 24 hours, the way a clock does. Square brackets turn the code into an elapsed count. Stays!E8 stores a
4.8243-day stay:

| Format code | Displays | Reads as |
|---|---|---|
| `h:mm` | `19:47` | ⚠️ The clock time 4.8 days later. The 96 hours of the first four days are dropped. |
| `[h]:mm` | `115:47` | 115 hours and 47 minutes, the real length of stay |
| `[m]` | `6947` | Total minutes |
| `0.0` | `4.8` | Days, with one decimal |

> ⚠️ Excel can't display a negative time or date in the default date system, so a negative duration (for example, a discharge
> time typed before the admit time) shows `#####` however wide the column is. Fix the data, not the column.

### 7. The Text format

The **Text** format tells Excel to keep whatever you type next exactly as typed, which protects identifiers such as MRNs, ZIP codes,
and lot numbers from losing leading zeros. Lesson 1.2 covered it in detail. Three rules matter here:

1. **Format first, then type.** Applying Text to a number that already lost its zeros doesn't bring them back.
2. A Text cell shows formulas as text instead of calculating them. Set it back to General and re-enter the formula.
3. Text and a custom format like `00000000` are not the same thing:

| | Text format (`'00412345` or Text then type) | Custom format `00000000` on the number 412345 |
|---|---|---|
| Stored value | The text `00412345` | The number 412345 |
| Matches the MRN `00412345` from another system | Yes | No, because a number never equals text |
| Good for | IDs you match, look up, or export | Displaying numbers that are already numbers, such as a printed list of an export you can't change |

### 8. Custom number formats

When no built-in format does what you need, you write a **custom number format**: a short code made of placeholders and literal
characters. Select the cells, press **Ctrl + 1** (Mac: **⌘ + 1**), choose **Custom** on the Number tab, and type the code in the
**Type** box. The Sample box previews the result, and the list below the box shows the codes already in the workbook. Clicking a
cell first and then opening the dialog shows *that cell's* code, which is the easiest way to read or tweak an existing format.

#### 8a. Placeholders

| Symbol | Meaning | Code → result |
|---|---|---|
| `0` | A digit that always shows, even if it's a zero | 4.5 with `0.00` → `4.50` · 12 with `000` → `012` |
| `#` | A digit that shows only if it's significant | 4.5 with `#.##` → `4.5` · 0.75 with `#.00` → `.75` |
| `?` | Like `0`, but shows a space instead of an insignificant zero, so decimals line up | 2.75 with `# ?/?` → `2 3/4` |
| `.` | Decimal point | |
| `,` between digit placeholders | Thousands separator | 9165952 with `#,##0` → `9,165,952` |
| `,` after the last digit placeholder | **Scales** the display: each trailing comma divides by 1,000 | 9165952 with `#,##0,` → `9,166` |
| `%` | Multiplies by 100 and shows % | 0.028 with `0.0%` → `2.8%` |
| `E+00` | Scientific notation | 9165952 with `0.00E+00` → `9.17E+06` |
| `@` | The cell's text, in the text section | |

Scaling commas are how finance shows thousands and millions without changing the data:

| Code | 9165952 displays |
|---|---|
| `#,##0` | `9,165,952` |
| `#,##0,` | `9,166` |
| `$#,##0,"K"` | `$9,166K` |
| `$#,##0.0,,"M"` | `$9.2M` |

#### 8b. Literal text: units and labels

Text you want printed goes in double quotes, such as `0.0 "days"` (4.8 days) or `0 "beds"` (36 beds). A single character can
follow a backslash instead, as in `0\h`. These characters need no quotes at all: `$ + - ( ) : / ! ^ & ' ~ { } < > =` and the
space. That's why a ZIP+4 format like `00000-0000` works as typed. Letters such as `d`, `m`, `y`, `h`, `s`, and `E` are codes, so
always quote words that contain them.

#### 8c. Sections: positive; negative; zero; text

A code can have up to four **sections** separated by semicolons. Each section formats a different kind of value:

```
#,##0;(#,##0);"-";@
```

That code shows 9,165,952 as `9,165,952`, −353,595 as `(353,595)`, zero as `-`, and any text as typed.

| Sections in the code | How Excel uses them |
|---|---|
| 1 | Every number uses it. Negative numbers get a minus sign in front automatically. |
| 2 | Positive numbers and zero use the first, negative numbers the second. |
| 3 | Positive; negative; zero. |
| 4 | Positive; negative; zero; text. |

> ⚠️ **The minus sign disappears in a negative section.** As soon as a code has a second section, Excel stops adding a minus
> sign to negative numbers and displays them exactly as that section says. With `#,##0;#,##0`, −353,595 displays as `353,595`, which
> looks positive. Put the sign you want into the negative section yourself: `#,##0;-#,##0` or `#,##0;(#,##0)`.

You can leave a section empty to hide that kind of value. `#,##0;-#,##0;` (nothing after the last semicolon) shows zeros as a
blank cell, and `;;;` hides everything.

> ⚠️ A hidden value is still in the cell. It shows in the formula bar, it's used in calculations, and it travels with every copy of
> the file. Never use `;;;` to "protect" patient information.

#### 8d. Colors

Put a color name in square brackets at the very start of a section: `[Black]`, `[Blue]`, `[Cyan]`, `[Green]`, `[Magenta]`, `[Red]`,
`[White]`, or `[Yellow]`. `[Color1]` through `[Color56]` pick from Excel's classic palette, where `[Color10]` is a readable dark
green (plain `[Green]` is very bright). For example, `#,##0;[Red](#,##0)` shows over-budget amounts in red parentheses.

> ⚠️ About 1 in 12 men has red-green color vision deficiency, and many reports are printed in black and white. Make sure the
> *text* carries the meaning too, with a sign, parentheses, or an arrow, and use color only as reinforcement.

#### 8e. Conditions

A condition in square brackets at the start of a section replaces the usual positive/negative meaning. Excel checks the sections
from left to right and uses the first one whose condition is true. You can use up to two conditions, and the last section catches
everything else:

```
[>=1000000]$#,##0.0,,"M";[>=1000]$#,##0,"K";$#,##0
```

With that code, 9,165,952 shows `$9.2M`, 26,907 shows `$27K`, and 950 shows `$950`. To combine a condition with a color, put the
color first, as in `[Red][<0]`.

#### 8f. Spacing tricks you'll see in built-in codes

Select a cell you formatted with Ctrl + Shift + $, open Format Cells, and click **Custom**. The Type box shows
`$#,##0.00_);($#,##0.00)` (on a Mac, `$#,##0.00_);[Red]($#,##0.00)`). The `_)` means "leave a space as wide as a closing parenthesis," so positive numbers line up with
negative ones in parentheses. An asterisk repeats the next
character to fill the cell, so `* ` in the Accounting code pushes the $ to the left edge.

#### 8g. Healthcare recipes

| Need | Code | Stored → displayed |
|---|---|---|
| 8-digit MRN from a number | `00000000` | 1927734 → `01927734` |
| 5-digit ZIP from a number | `00000` | 5401 → `05401` |
| Length of stay with a unit | `0.0 "days"` | 4.8243 → `4.8 days` |
| Staffed beds with a unit | `0 "beds"` | 36 → `36 beds` |
| Temperature | `0.0 "°F"` | 98.6 → `98.6 °F` |
| Budget in thousands | `$#,##0,"K"` | 9165952 → `$9,166K` |
| Variance with sign on both sides | `+0.0%;-0.0%;0.0%` | 0.028 → `+2.8%` |
| Hide zeros in a staffing grid | `0;-0;` | 0 → (blank) |
| Elapsed hours | `[h]:mm` | 4.8243 → `115:47` |

> 📋 Custom formats are saved in the workbook, not on your computer. To reuse one in another file, copy a formatted cell into it
> or use Format Painter (section 12). To remove one you no longer need, select it in the Custom list and click **Delete**.

### 9. Alignment, wrap text, and orientation

Alignment controls where content sits in its cell. The buttons are in **Home → Alignment**, and every option is on the
**Alignment** tab of Format Cells.

| Setting | Options and use |
|---|---|
| **Horizontal** | General (numbers right, text left), Left, Center, Right, Fill, Justify, **Center Across Selection**, Distributed. Keep numbers right-aligned so digits line up. |
| **Vertical** | Top, Center, Bottom. Use Center or Top when some cells in a row wrap onto several lines. |
| **Wrap Text** | Breaks long text onto more lines inside the cell and grows the row height. Windows: **Alt, H, W**. Mac: **Home → Wrap Text**. |
| **Indent** | Shifts content in from the edge, a clean way to show subcategories under a heading. Use the **Increase Indent** button. |
| **Orientation** | Angles or stacks text with the **ab↗** button, for narrow column headers in a day-by-day census grid. |
| **Shrink to fit** | Reduces the font size until the content fits. It can make text unreadably small, so use it sparingly. |

To start a new line at an exact spot, press **Alt + Enter** (Mac: **⌃ + ⌥ + Return**) while typing, as you did in Lesson 1.2. Excel
turns on Wrap Text for you.

### 10. Merge & Center vs. Center Across Selection

**Merge & Center** joins several cells into one big cell and centers the content. It looks tidy, which is why so many reports use it
for titles and group labels. It also causes some of Excel's most annoying problems:

| What you try | What happens with merged cells |
|---|---|
| Merge cells that each hold a value | Excel warns that it keeps only the **upper-left value** and deletes the rest |
| Sort the list | Excel refuses: "To do this, all the merged cells need to be the same size." |
| Filter by a merged label | Only the first row of each block matches, because the other rows are empty |
| Copy and paste a range | Excel refuses to paste into a merged area with a different shape |
| Select a few cells in one column by dragging through a merged area | The selection jumps to every column the merged cell spans |
| Point a formula at a cell inside the block | It returns 0 or blank, because only the top-left cell holds the value |
| Turn the range into an Excel Table (Lesson 3.1) | Tables don't allow merged cells |

**Center Across Selection** gives the same look without merging. Each cell stays independent, so every one of the problems above goes
away. Use it for titles:

1. Type the title in the left cell, such as A1.
2. Select the title cell and the cells it should span, such as A1:F1.
3. Press **Ctrl + 1** (Mac: **⌘ + 1**), open the **Alignment** tab, and set **Horizontal** to **Center Across Selection**.

Center Across Selection only works across columns. For *group labels down a column*, such as a service line next to each of its
departments, don't merge at all. Repeat the label on every row. That's what data tools like sorting, filtering, and PivotTables
expect, and you can make the repeats look quieter with a gray font if you like.

**Fixing merged cells someone else made:**

1. Select the merged cells and click **Home → Merge & Center** to unmerge them (or choose **Merge & Center ▾ → Unmerge Cells**).
   Windows KeyTips: **Alt, H, M, U**.
2. Each label now sits in the top cell of its old block, with blanks below. Select a block from its label down and press
   **Ctrl + D** (Mac: **⌘ + D**) to fill the label down.

> 💡 **Tip:** For a long list with many blocks, select just the label cells (A5:A20 on the Budget sheet, not the whole column,
> which would also catch the blank rows above and below the list), press **F5** or **Ctrl + G** (Mac: **⌃ + G**), and click
> **Special… → Blanks → OK**. Type `=`, press **↑**, and press **Ctrl + Enter** (Mac: **⌘ + Return**). Every blank now copies the
> label above it. Then copy the column and use Paste Special → Values (Lesson 1.2) to replace those formulas with plain text.

### 11. Fonts, fills, and borders

| Action | Windows | Mac |
|---|---|---|
| Bold / Italic / Underline | **Ctrl + B** / **Ctrl + I** / **Ctrl + U** | **⌘ + B** / **⌘ + I** / **⌘ + U** |
| Strikethrough | **Ctrl + 5** | **⌘ + Shift + X** |
| Fill color / Font color | **Alt, H, H** / **Alt, H, F, C** | **Home → Fill Color** / **Home → Font Color** |
| Borders menu | **Alt, H, B** | **Home → Borders ▾** |
| Outline border around the selection | **Ctrl + Shift + &** | **⌘ + ⌥ + 0** |
| Remove borders from the selection | **Ctrl + Shift + _** | **⌘ + ⌥ + -** |

**Colors.** The Fill Color and Font Color palettes have two parts. **Theme colors** (the top block) come from the workbook's theme
and change if the theme changes. **Standard colors** (the bottom row) never change. Use theme colors for anything that should match
your organization's look.

**Borders.** The **Borders ▾** menu applies common borders in one click: **All Borders**, **Outside Borders**, **Thick Outside
Borders**, **Bottom Double Border**, and more. For a specific line style or color, use the **Border** tab of Format Cells: pick the
style and color first, then click the edges in the preview.

> 📋 The light gray **gridlines** you see on every sheet are not borders. They don't print unless you tick **Page Layout → Sheet
> Options → Gridlines → Print**, and you can hide them on screen with **View → Gridlines**. Add borders where the printed report
> needs lines.

**Design habits for health-system reports:**

- Make the header row stand out once (bold, a dark fill with white text) and keep the body plain.
- Keep the same number of decimals down a column, and right-align numbers.
- Put units in the header ("LOS (days)") or in the format (`0.0 "days"`), not as text typed into each cell.
- Use one or two accent colors. A report where everything is highlighted highlights nothing.
- To color cells *based on their values* (for example, red when a lab result is critical), use conditional formatting, which you'll
  learn in Lesson 3.2.

### 12. Format Painter and Paste Special → Formats

**Format Painter** copies all of one cell's formatting (number format, font, fill, borders, alignment, and even conditional
formatting) to other cells. The values aren't touched.

1. Select the cell whose formatting you want to copy.
2. Click **Home → Format Painter** (the paintbrush). Windows KeyTips: **Alt, H, F, P**. The pointer becomes a brush.
3. Click or drag over the target cells. The brush turns off after one use.

**Double-click** the paintbrush instead of clicking it to lock it on. Then you can paint as many ranges as you like, even on other
sheets, until you press **Esc** or click the paintbrush again. If you copy formatting from one cell and paint a larger range, Excel
repeats that cell's formatting across the whole range.

**Paste Special → Formats** does the same job through the clipboard: copy the formatted cells, select the target, press
**Ctrl + Alt + V** (Mac: **⌃ + ⌘ + V**), and choose **Formats**.

> ⚠️ Format Painter replaces *all* the target's formatting. If you paint a currency cell onto a percentage column, the percentages
> become dollar amounts. Press **Ctrl + Z** (Mac: **⌘ + Z**) right away if you paint the wrong range.

### 13. Cell styles and themes

A **cell style** is a named bundle of formatting, such as a font, fill, border, and number format, that you apply in one click from
**Home → Cell Styles**. Styles keep reports consistent, and when you change a style, every cell that uses it updates.

| Group in the gallery | Styles | Use them for |
|---|---|---|
| Good, Bad and Neutral | Normal, Bad, Good, Neutral | Quick status flags |
| Data and Model | Calculation, Check Cell, Explanatory Text, Input, Linked Cell, Note, Output, Warning Text | Marking input cells vs. calculated cells in a model |
| Titles and Headings | Heading 1–4, Title, **Total** | Report structure. **Total** makes a grand-total row bold and draws lines above and below it in the accounting style |
| Themed Cell Styles | Accent1–Accent6 and their 20%, 40%, and 60% tints | Banding and highlights in theme colors |
| Number Format | Comma, Comma [0], Currency, Currency [0], Percent | The same formats as the Home-tab buttons, as styles. The style named **Currency** is the one the **$** button applies, so it uses the Accounting format, not Currency |

- To change a style everywhere, right-click it in the gallery and choose **Modify…**.
- To save your own look, format one cell and choose **Cell Styles → New Cell Style…**.
- **Normal** is the style every cell starts with. Modify Normal to change the default font of the whole workbook.

A **theme** is a workbook-wide set of colors, fonts, and effects. Change it in **Page Layout → Themes**, or change only the
**Colors** or **Fonts** next to it. Everything that uses theme colors, theme fonts, or theme-based cell styles updates at once.
Standard colors and fonts you picked by name stay as they are. A health system can save its brand colors and fonts with
**Themes → Save Current Theme…** so every department's report looks like it came from the same organization.

> 📋 **Version note:** Microsoft 365 switched its default theme fonts in 2023–2024. New workbooks use **Aptos Narrow** in cells and
> Aptos Display for headings. Excel 2021 and earlier use **Calibri** and Calibri Light. Both are fine, and a workbook keeps the
> fonts it was created with.

### 14. Column widths, row heights, and clearing formats

**AutoFit** sizes a column to its widest entry. Double-click the right boundary of the column header, or select several columns and
double-click any boundary between them. Windows KeyTips: **Alt, H, O, I** (column width) and **Alt, H, O, A** (row height). On
both platforms, **Home → Format** also has AutoFit Column Width and Column Width… for an exact number.

> ⚠️ **AutoFit on a whole column measures every cell in it,** including a long report title or a note above the table. On the
> Budget sheet, AutoFit on column A would make it as wide as the sentence in A2. Select just the table's cells (A4:F21) and choose
> **Home → Format → AutoFit Column Width** instead. That command fits the columns to the selected cells only.

> 💡 **Tip:** AutoFit a column after formatting it, not before. A number that fit as `9165952` may need more room as `$9,165,952.00`,
> and Excel shows `#####` until it gets that room.

**Clearing formats.** **Home → Clear (eraser) → Clear Formats** (Windows KeyTips: **Alt, H, E, F**) removes every kind of
formatting and sets the number format back to General. **Clear All** removes the content too.

> ⚠️ **Delete clears contents, not formats.** A cell that once held a date keeps its date format after you press Delete. Type 36
> into it later and Excel shows a date in February 1900, because serial number 36 is 2/5/1900. Clear the formats, or press
> **Ctrl + Shift + ~** (Mac: **⌃ + Shift + ~**) to reset the number format.

### 15. Worked example: variance arrows for a dashboard

*Goal: on the Budget sheet, show each department's Variance % as a green ▲ when it's under budget and a red ▼ when it's over,
with one decimal place.*

1. **Start from what you know.** `0.0%` shows Radiology's 0.0281259383 as `2.8%` and Pharmacy's −0.0332 as `-3.3%`.
2. **Give negatives their own section.** `0.0%;0.0%` shows Pharmacy as `3.3%`. The minus sign disappeared, exactly as section 8c
   warned. That's fine here, because the arrow will carry the direction.
3. **Add the arrows as literal text.** `▲ 0.0%;▼ 0.0%` shows `▲ 2.8%` and `▼ 3.3%`. Paste the arrows in from this page, or insert
   them with **Insert → Symbol** (on Windows you can also type **Alt + 30** and **Alt + 31** on the numeric keypad).
4. **Add a zero section** so a department exactly on budget shows no arrow: `▲ 0.0%;▼ 0.0%;0.0%`.
5. **Add colors** at the start of the first two sections: `[Color10]▲ 0.0%;[Red]▼ 0.0%;0.0%`.

The finished code shows Medical-Surgical 4 West as a green `▲ 2.8%` and Orthopedics & Spine as a red ▼. The arrows mean
the report still reads correctly when printed in black and white. The stored values are untouched, so any formula that uses
the Variance % column still gets the real numbers.

## 🧪 Hands-on practice

Download [`1.3-formatting-cells.xlsx`](1.3-formatting-cells.xlsx) and open the **Practice** sheet. Apply each format on the sheet the
task names, then type your answer in the yellow cell. The **Check** column turns green when you're right.

<!-- BEGIN GENERATED: practice -->
Tasks 1–4 and 12–13 use the Budget sheet, tasks 5–7 the Registry sheet, and tasks 8–11 the Stays sheet. Work in order, because some tasks build on earlier ones. Most tasks ask you to apply a format and then type exactly what the cell displays, including any dollar sign, commas, parentheses, minus sign, or % sign. Those yellow cells are formatted as Text, so Excel keeps your entry exactly as you type it. Tasks 3, 9, and 12 ask for a plain number instead. Answers assume US regional settings.

| # | Task | Hint |
|:-:|------|------|
| 1 | On the Budget sheet, select the Variance cells E5:E21 and press Ctrl + Shift + $ (Mac: ⌃ + Shift + $). What does E11 (Emergency Department, which spent more than its budget) display now? | Negative amounts in this format don't use a minus sign |
| 2 | Select the Variance % cells F5:F21, press Ctrl + Shift + % (Mac: ⌃ + Shift + %), then click Home → Increase Decimal once. What does F17 (Orthopedics & Spine) display? | The shortcut shows 0 decimals; each Increase Decimal click adds one |
| 3 | C5 already has a finance format: it shows the Laboratory budget in thousands with a K. What number is actually stored in C5? Click the cell and read the formula bar. | The cell shows one thing; the formula bar shows another |
| 4 | Copy C5's format to the rest of the money columns with Format Painter: select C5, double-click Home → Format Painter, drag over C5:E20, then drag over the Total row C21:E21, and press Esc. What does E6 (Pharmacy's variance) display now? | A format with only one section puts a minus sign in front of negatives |
| 5 | On the Registry sheet, select the PatientDue cells F5:F24 and click Home → Accounting Number Format (the $ button). Cox, Ronald (row 5) owes nothing for this visit. What character does F5 show where the 0 used to be? | Compare Accounting with Currency in the guide's table |
| 6 | The MRN column lost its leading zeros on the way out of the registration system. Bluestone MRNs are always 8 digits. Select B5:B24, open Format Cells (Ctrl + 1; Mac: ⌘ + 1), choose Custom, and type the code 00000000 in the Type box. What does B13 display? | Each 0 in the code is a digit that always shows, even when it's a zero |
| 7 | The Phone column (D) stores 10-digit numbers. Write a custom number format that shows each one the usual US way: the first three digits in parentheses, a space, three digits, a hyphen, and the last four digits. For example, D5 (5555298330) should display as (555) 529-8330. Apply your format to D5:D24, then type the format code you used. | 0 is a digit placeholder; parentheses, spaces, and hyphens can be typed as they are |
| 8 | On the Stays sheet, the Admitted column (C) shows serial numbers because the export lost its date format. Select C5:C40 and apply the custom format ddd mm/dd/yyyy h:mm AM/PM. What does C6 display? | ddd is the short day name; mm right after h means minutes |
| 9 | The LOS (days) column is formatted to show one decimal place. E7 shows 3.1. What value is actually stored in that cell? Read the formula bar and enter at least 4 decimal places. | Formatting rounds the display, not the value |
| 10 | Hospitals often track length of stay in hours. Select E5:E40 and apply the custom format [h]:mm. What does E15 (the longest stay of the month) display? | Square brackets let the hours keep counting past 24 |
| 11 | Column G (LOS vs Expected) stores LOS minus the diagnosis's expected LOS, in days. Write a two-section custom format for G5:G40: positive values in red with a plus sign and the word days (like +2.1 days), and negative values with a minus sign (like -0.4 days), each with one decimal. What does G5 display? | Sections are separated by semicolons: positive;negative. The second section needs its own minus sign |
| 12 | Back on the Budget sheet, column A labels each department's service line with merged cells. Select A5:A20 and read Count on the status bar (it counts cells that aren't empty). How many of those 16 cells actually contain a service-line label? | A merged block stores its text in one cell only |
| 13 | Finish the Budget report so it looks like the hidden Budget Key sheet: (1) unmerge A1:F1 and center the title with Center Across Selection; (2) unmerge column A and fill each service line down so every row has its label; (3) make header row 4 bold with white text on a dark blue fill, wrapped and centered; (4) add All Borders to A4:F21; (5) apply the Total cell style to A21:F21; (6) select A4:F21 and AutoFit the columns with Home → Format → AutoFit Column Width. Then zoom in on row 21: what kind of line does the Total style draw along the bottom of the row? | Cell Styles is on the Home tab; look closely at the bottom edge of the total row |
<!-- END GENERATED: practice -->

## ✅ Answer key

The workbook has a hidden **Answer Key** sheet (right-click any sheet tab → **Unhide…**). Its *Live result* column uses the TEXT
function, which you'll meet in Lesson 2.2, to prove what each format code displays. Two more hidden sheets show the finished work:
**Budget Key** for task 13 and **Report Key** for the bonus. The same answers are below, collapsed so you don't see them by accident.

<!-- BEGIN GENERATED: answers -->
<details>
<summary><b>🔑 Show the answer key</b> — Try every task before opening this.</summary>

**1. Currency shortcut on a negative variance**

- **Answer:** ($572,762.00)
- **Solution:**

1. Select **Budget!E5:E21**.
2. Press **Ctrl + Shift + $** (Mac: **⌃ + Shift + $**).
3. Read **E11**.


The shortcut applies the Currency format with two decimals, whose code is `$#,##0.00_);($#,##0.00)`. The part after the semicolon is used for negative numbers, and it wraps them in parentheses instead of showing a minus sign, the way accountants write losses. The stored value is still -572762. Click the cell and look at the formula bar to confirm. (The key's live formula uses TEXT, a Lesson 2.2 function that returns what a format would display.)

**2. Percentage with one decimal place**

- **Answer:** -4.6%
- **Solution:**

1. Select **Budget!F5:F21**.
2. Press **Ctrl + Shift + %** (Mac: **⌃ + Shift + %**). The cells show whole percentages.
3. Click **Home → Increase Decimal** (the .00 button with the left arrow) once.
4. Read **F17**.


The cell stores -0.046203… The Percentage format multiplies by 100 for display only and adds the % sign, so it shows -5% at first and -4.6% after one Increase Decimal click. Excel rounds the display, but the stored value keeps every digit.

**3. Stored value behind a thousands (K) display**

- **Answer:** 11,736,897
- **Solution:** Click **Budget!C5** and read the formula bar. It shows the full dollar amount.

The cell displays `$11,737K`, but it stores 11,736,897. The custom code `$#,##0,"K"` ends with a comma after the last 0, which tells Excel to *display* the number divided by 1,000. Any formula that uses C5 still gets the full amount, so totals stay exact even when every cell is shown in thousands.

**4. Format Painter: thousands format on a negative number**

- **Answer:** -$231K
- **Solution:**

1. Select **Budget!C5**.
2. **Double-click** **Home → Format Painter** (the paintbrush). Double-clicking keeps it switched on.
3. Drag over **C5:E20**, then drag over the Total row, **C21:E21**.
4. Press **Esc** (or click the paintbrush again) to switch it off.
5. Read **E6**.


Format Painter copies *all* of C5's formatting (number format, font, fill, borders, alignment) onto the cells you paint. `$#,##0,"K"` has a single section, so Excel uses it for every number and simply puts a minus sign in front of negatives: -230,692 shows as `-$231K`. The parentheses from task 1 are gone because each cell has only one number format, and the newest one wins. A single click on Format Painter paints once. A double-click lets you paint several ranges until you press Esc.

**5. Accounting format: how a zero looks**

- **Answer:** -
- **Solution:**

1. Select **Registry!F5:F24**.
2. Click the **$** button in the **Number** group of the **Home** tab (its ScreenTip says Accounting Number Format).
3. Look at **F5**: the $ sits at the left edge and a dash sits near the right.


The Accounting code is `_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)`. Its third section (zero) prints a dash, so a column of balances shows which patients owe nothing at a glance. Accounting also pins the $ to the left edge of the cell and lines up the decimal points, and it always shows negatives in parentheses. Currency puts the $ right next to the number and shows `$0.00` for zero. The cell still stores 0.

**6. Custom format 00000000 for MRNs**

- **Answer:** 00440415
- **Solution:**

1. Select **Registry!B5:B24** and press **Ctrl + 1** (Mac: **⌘ + 1**).
2. On the **Number** tab choose **Custom**, replace the Type box with `00000000`, and click **OK**.
3. Read **B13**.


The cell stores the number 440415. Each `0` in `00000000` is a placeholder that forces a digit, so Excel pads the display to 8 digits: `00440415`. This fixes how the MRN *looks* on a printed list. The value is still a number, though, so it won't match the text MRN `00440415` in another system or in a lookup. When you control data entry, store IDs as text instead (Lesson 1.2).

**7. Write a phone-number format code**

- **Answer:** (000) 000-0000
- **Solution:**

1. Select **Registry!D5:D24** and press **Ctrl + 1** (Mac: **⌘ + 1**).
2. Choose **Custom** and type `(000) 000-0000` in the Type box. Click **OK**.
3. Type `(000) 000-0000` in the answer cell.


`(000) 000-0000` has ten 0 placeholders, filled from the right with the number's ten digits. Parentheses, the space, and the hyphen are characters Excel prints as-is, without quotes. The check also accepts `(###) ###-####` (the same for 10-digit numbers) and Excel's built-in **Special → Phone Number** format, `[<=9999999]###-####;(###) ###-####`, which you'll see in the Custom box if you used Special.

**8. Custom date-and-time format**

- **Answer:** Tue 11/25/2025 11:58 AM
- **Solution:**

1. Select **Stays!C5:C40** and press **Ctrl + 1** (Mac: **⌘ + 1**).
2. Choose **Custom**, type `ddd mm/dd/yyyy h:mm AM/PM`, and click **OK**.
3. Read **C6**.


The cell stores 45986.49861: the whole number is the date and the decimal is the time of day. `ddd` gives the short day name, `mm/dd/yyyy` the date with leading zeros, and `h:mm AM/PM` a 12-hour time without a leading zero. Because `mm` follows `h`, Excel reads it as minutes, not months.

**9. Stored value behind a one-decimal display**

- **Answer:** 3.14097222… (any entry within 0.0001 is accepted)
- **Solution:** Click **Stays!E7** and read the formula bar.

The stay lasted 4,523 minutes, and 4,523 ÷ 1,440 minutes per day = 3.14097222… days. The `0.0` format only rounds what you see. If you add up the LOS column, Excel adds the full values, so a total can differ slightly from the sum of the rounded numbers on screen. Use ROUND (Lesson 1.4) when the rounded number is the one you mean.

**10. Elapsed hours with [h]:mm**

- **Answer:** 412:47
- **Solution:**

1. Select **Stays!E5:E40** and press **Ctrl + 1** (Mac: **⌘ + 1**).
2. Choose **Custom**, type `[h]:mm`, and click **OK**.
3. Read **E15**.


The cell stores 17.1993 days. One day is 24 hours, so that's 412 hours and 47 minutes. `[h]` shows *elapsed* hours. Plain `h:mm` would show `4:47`, the clock time, because ordinary hours roll back to 0 every 24 hours. Use `[h]:mm` for any duration that can pass a day: LOS, shift hours, ED boarding time.

**11. Two sections, a color, and a unit**

- **Answer:** -2.0 days
- **Solution:**

1. Select **Stays!G5:G40** and press **Ctrl + 1** (Mac: **⌘ + 1**).
2. Choose **Custom**, type `[Red]+0.0 "days";-0.0 "days"`, and click **OK**.
3. Read **G5**.


`[Red]+0.0 "days";-0.0 "days"` has two sections. The first (positive and zero) starts with the color `[Red]`, then a literal +, the number, and the text "days" in quotes. The second section (negative) has no color. When a format has a negative section, Excel stops adding the minus sign for you, so you type it yourself. Without it, -1.975 would display as `2.0 days` and look like a positive number. 15 of the 36 stays show in red because they lasted longer than expected.

**12. What merged cells really store**

- **Answer:** 11
- **Solution:** Select **Budget!A5:A20** and read **Count** on the status bar.

There are 16 departments but only 11 labels, because a merged block keeps its value in the top-left cell and the other cells are empty. For example, only the first Medicine row really says Medicine. Sort, filter, copy, or count by service line and the rows under each block behave as if they had no service line. Excel even refuses to sort a range whose merged cells are different sizes. Task 13 replaces the merges with a label on every row. (After task 13, the key's live COUNTA shows 16, because every row then has its label.)

**13. Style the report (Center Across Selection, header, borders, Total style)**

- **Answer:** double
- **Solution:**

1. **Title:** select **A1** and click **Home → Merge & Center** to unmerge it. Select **A1:F1**, press **Ctrl + 1** (Mac: **⌘ + 1**), and on the **Alignment** tab set **Horizontal** to **Center Across Selection**.
2. **Service lines:** select **A5:A20** and choose **Home → Merge & Center ▾ → Unmerge Cells**. Then, for each block that now has blank cells under its label (A5:A7, A12:A14, A19:A20), select the block and press **Ctrl + D** (Mac: **⌘ + D**) to fill the label down.
3. **Header:** select **A4:F4**. Press **Ctrl + B** (Mac: **⌘ + B**), pick **Home → Font Color ▾ → White**, **Home → Fill Color ▾ → Dark Blue** (or any dark theme color), then click **Wrap Text** and **Center**.
4. **Borders:** select **A4:F21** and choose **Home → Borders ▾ → All Borders**.
5. **Total row:** select **A21:F21** and choose **Home → Cell Styles → Total** (in the *Titles and Headings* group).
6. **AutoFit:** select **A4:F21** (just the table, not whole columns) and choose **Home → Format → AutoFit Column Width** (Windows KeyTips **Alt, H, O, I**). Selecting whole columns would also measure the long note in A2, and column A would grow as wide as that sentence.
7. Unhide **Budget Key** (right-click a sheet tab → **Unhide…**) and compare.


The built-in **Total** style makes the text bold and draws a thin line above the row and a **double** line below it, the accounting convention for a grand total. Because it's a *style*, every total row you apply it to looks the same. Its lines use a theme color, so if you switched the workbook to another theme in **Page Layout → Themes**, the Total lines would change color while a header fill picked from Standard Colors (Dark Blue) would stay put. Center Across Selection looks exactly like Merge & Center but leaves every cell independent, so sorting, selecting columns, and copying keep working.

</details>
<!-- END GENERATED: answers -->

## 🏆 Bonus challenge

<!-- BEGIN GENERATED: bonus -->
The Emergency Department director presents the December 2025 operating report to the CFO next week. Finance's style guide has five rules for department reports. Apply them on the Dec Report sheet, writing every custom format code yourself, then type what each listed cell displays. When you finish, compare your sheet with the hidden Report Key sheet.

- Rule 1: Center the title across A1:I1 without merging. Make the header row bold white text on a dark blue fill, wrapped and centered. Give the Total row the Total cell style.
- Rule 2: Show Dec Budget and Dec Actual in thousands with one decimal place, a dollar sign, and a K, so 843,879 displays as $843.9K.
- Rule 3: Show both Variance columns the same way, except that negative variances (over budget) are red and in parentheses instead of having a minus sign, so -25,403 displays as ($25.4K).
- Rule 4: Show both Var % columns with one decimal place, with negatives red and in parentheses, so -0.1075… displays as (10.8%).
- Rule 5: Show YTD Budget and YTD Actual with a dollar sign: in millions with two decimals and an M when the amount is 1,000,000 or more (10,290,782 displays as $10.29M), and otherwise in thousands with no decimals and a K.

Work on the **Bonus** sheet of the workbook.

- **B1.** Rule 2: what does C5 (Employee Benefits, Dec Actual) display? *(Hint: One comma after the last digit placeholder divides the display by 1,000)*
- **B2.** Rule 3: what does D11 (Salaries & Wages, Dec Variance) display? *(Hint: Two sections: positive;negative. Put the color first in the negative section)*
- **B3.** Rule 4: what does I6 (Equipment & Maintenance, YTD Var %) display? *(Hint: A % in a custom code multiplies by 100, in every section where it appears)*
- **B4.** Rule 5: what does G12 (the Total row's YTD Actual) display? *(Hint: Conditions go in square brackets at the start of a section, like [>=1000000])*
- **B5.** Rule 5: what does F8 (Other Operating, YTD Budget) display? *(Hint: A value that fails the first condition moves on to the next section)*
<!-- END GENERATED: bonus -->

<!-- BEGIN GENERATED: bonus-answers -->
<details>
<summary><b>🔑 Show the bonus solution</b> — Give it a real try first!</summary>

**B1. Thousands with one decimal**

- **Answer:** $261.7K
- **Solution:** Select **'Dec Report'!B5:C12**, press **Ctrl + 1** (Mac: **⌘ + 1**), choose **Custom**, and type `$#,##0.0,"K"`.

261,689 ÷ 1,000 = 261.689, shown with one decimal as `$261.7K`. The comma between `#` and `##0` is the thousands separator. The comma after `0.0` is the scaling comma.

**B2. Negative variance: red, parentheses, thousands**

- **Answer:** ($47.8K)
- **Solution:** Select **'Dec Report'!D5:D12**, then Ctrl+click **H5:H12**. Apply the custom format `$#,##0.0,"K";[Red]($#,##0.0,"K")`.

The first section formats positive variances (under budget). The second section starts with `[Red]` and wraps the same thousands pattern in parentheses, so -47,771 shows as `($47.8K)` in red. Because the format has a negative section, Excel adds no minus sign. In December, 5 of the 8 Dec Variance cells (including the total) turn red.

**B3. Negative percentage in red parentheses**

- **Answer:** (8.2%)
- **Solution:** Select **'Dec Report'!E5:E12** and Ctrl+click **I5:I12**. Apply the custom format `0.0%;[Red](0.0%)`.

The cell stores -0.082270… Each section has its own `%`, so both multiply by 100. The negative section adds `[Red]` and parentheses: `(8.2%)`.

**B4. Conditional format: millions**

- **Answer:** $17.56M
- **Solution:** Select **'Dec Report'!F5:G12** and apply the custom format `[>=1000000]$#,##0.00,,"M";[>=1000]$#,##0,"K";$#,##0`.

A condition in square brackets replaces the usual positive/negative meaning of a section. `[>=1000000]` sends 17,563,319 to the first section, where two scaling commas divide by 1,000,000: `$17.56M`. This total matches the Emergency Department's 2025 actual on the Budget sheet, because December closes the fiscal year.

**B5. Conditional format: thousands**

- **Answer:** $298K
- **Solution:** Same format as the previous part. Check that values under 1,000,000 fall through to the K section.

298,078 is less than 1,000,000, so Excel skips the first section and uses `[>=1000]$#,##0,"K"`: `$298K`. The full code is `[>=1000000]$#,##0.00,,"M";[>=1000]$#,##0,"K";$#,##0`. Its third section catches anything under 1,000. Without conditions, every value would get the same scale, and a `$298K` line would show as `$0.30M`.

</details>
<!-- END GENERATED: bonus-answers -->

## Key takeaways

- Formatting changes the **displayed value**, never the **stored value**. The formula bar shows what's really in the cell, and every
  calculation uses that.
- **Ctrl + 1** (Mac: **⌘ + 1**) opens Format Cells. **Ctrl + Shift** with `~ ! @ # $ % ^` applies General, Number, Time, Date,
  Currency, Percentage, and Scientific.
- **Accounting** (the $ button) pins the $ to the left, aligns decimals, and shows zero as a dash. **Currency** (Ctrl + Shift + $)
  keeps the $ next to the number.
- Custom codes are built from placeholders (`0 # ?`), scaling commas, quoted text, `[Color]` tags, `[conditions]`, and up to four
  sections: positive; negative; zero; text. A negative section needs its own sign.
- Use `[h]:mm` for durations that can pass 24 hours, and remember that `m` means minutes only next to `h` or `s`.
- Center titles with **Center Across Selection** and repeat group labels on every row instead of merging cells.
- **Format Painter** (double-click to lock it), **cell styles**, and **themes** make every report in a workbook look the same.

<!-- BEGIN GENERATED: nav -->
---

⬅️ **Previous:** [1.2 Data Entry, AutoFill & Editing](../02-data-entry-autofill/README.md) · 🏠 [Course home](../../README.md) · **Next:** [1.4 Your First Formulas & Functions](../04-basic-formulas/README.md) ➡️
<!-- END GENERATED: nav -->
