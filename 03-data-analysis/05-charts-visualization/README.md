# Lesson 3.5 · Charts & Data Visualization

> **Level:** Intermediate · **Time:** about 135 minutes · **Workbook:** [`3.5-charts-visualization.xlsx`](3.5-charts-visualization.xlsx)
> **Data:** Bluestone Health System, pre-summarized for charting: monthly ED visits by hospital with counts of patients who left without being seen (LWBS), Jan 2024 – Dec 2025, 2025 ED arrivals and average door-to-provider minutes by hour of day, 2025 readmission rates by service line, 2025 payer mix, a sample of 300 inpatient stays from 2025, 2025 claim denials by reason, and 4 West's FY2025 budget-to-actual margin bridge. The summaries come from [`ed_visits.csv`](../../data/README.md#ed_visitscsv), [`encounters.csv`](../../data/README.md#encounterscsv), [`claims.csv`](../../data/README.md#claimscsv), and [`budget.csv`](../../data/README.md#budgetcsv) in the data dictionary.

A table of 24 months × 3 hospitals holds every number the ED director needs, but nobody sees the winter surge in it.
A line chart shows that surge in five seconds. Hospital leaders read dozens of reports a week, so the analyst whose chart
answers the question at a glance is the one whose work gets used. A badly built chart does the opposite. An axis that
starts at 5% can make the gap between two readmission rates look several times bigger than it is, and a rainbow of colors
can hide the one bar that matters. In this lesson you'll build the charts that operations, quality, and finance teams use every
week: arrivals by hour, monthly volume trends, ranked readmission rates, payer mix, length-of-stay distributions, scatter
charts with trendlines, combo charts, Pareto charts of claim denials, a budget waterfall, and sparklines. You'll also learn to
choose the right chart for a question, link titles to live data, and design charts that everyone can read.

## What you'll learn

- Choose the right chart for comparisons, trends, parts of a whole, distributions, and relationships
- Build and format column, line, bar, scatter, histogram, combo, and Pareto charts
- Add trendlines, dynamic titles, and sparklines
- Design clear, accessible charts that tell one story

## 📖 Guide

### 1. Start with the question, then choose the chart

A **chart** turns numbers into shapes so readers can compare them at a glance. Each chart type is built for one kind of
question, so decide what your audience needs to see first. The chart type usually follows.

| The question is about… | Bluestone example | Best choice | Avoid |
|---|---|---|---|
| **Comparison** of one measure across categories | ED arrivals by hour of day | Column chart, or a bar chart when labels are long | Pie, 3-D columns |
| **Ranking** | Readmission rate by service line | Sorted bar chart | Alphabetical order |
| **Trend** over time | Monthly ED visits, 2024–2025 | Line chart for many periods, column chart for a few | Pie, or bars with time running down the page |
| **Part of a whole** | 2025 payer mix | Pie or doughnut for 2–5 parts, sorted bar for more, 100% stacked bar to compare mixes | Pie with many slices, several pies side by side |
| **Distribution** of one measure | Length of stay for 300 stays | Histogram, or box and whisker to compare groups | A column for every raw value |
| **Relationship** between two measures | Length of stay vs. charges | Scatter chart with a trendline | Line chart (it joins points in row order) |
| **Causes ranked by impact** | Claim denials by reason | Pareto chart | Pie |
| **How a total changes**, step by step | Budgeted margin → actual margin | Waterfall chart | Stacked columns |
| **Two measures with different units** | Arrivals and minutes to see a provider | Combo chart with a secondary axis | Both series squeezed onto one axis |
| **A trend beside the numbers** in a table | Each hospital's 24 months | Sparklines | A full chart for every row |

> 💡 **Tip:** Write the chart's headline before you build it, for example *"ED waits rise and fall with arrivals."* If you
> can't write a one-sentence headline, you aren't ready to choose a chart. Section 15 shows how to put that headline in the
> title.

### 2. Insert a chart

1. Select the data, including the header row. The headers become the series names in the legend. To chart a whole
   Table, or a block with no empty rows or columns, you can click a single cell inside it instead, and Excel charts
   the whole block.
2. To chart columns that aren't next to each other, select the first block, then hold **Ctrl** (Mac: **⌘**) and select the
   next one. Every block must cover the same rows.
3. On the **Insert** tab, pick a button in the **Charts** group, or choose **Recommended Charts** to preview several
   options with your own data.
4. Drag the chart to where you want it, and drag a corner handle to resize it.

On Windows, **Alt + F1** inserts the default chart type (a clustered column chart unless you've changed it) on the current
sheet, and **F11** puts it on a new **chart sheet**, a sheet that holds only the chart. On a Mac, use the Insert tab. To move
any chart to its own sheet later, choose **Chart Design → Move Chart**.

When a chart is selected, two or three extra ribbon tabs appear:

| Excel version | Chart tabs |
|---|---|
| Microsoft 365 and Excel 2021 or later (Windows), Excel 2016 or later (Mac) | **Chart Design** and **Format** |
| Excel 2013 to 2019 (Windows) | **Chart Tools → Design** and **Format** |
| Excel 2010 (Windows) | **Chart Tools → Design**, **Layout**, and **Format** |

**How Excel reads your selection.** Excel decides which cells are **categories** (the labels along the category axis) and
which are **series** (the sets of numbers it plots):

| Your selection | What Excel does |
|---|---|
| A text column on the left, numbers to its right | The text becomes the categories, and each number column becomes a series |
| More rows than columns | Each column becomes a series |
| More columns than rows | Each row becomes a series. **Chart Design → Switch Row/Column** flips it |
| Real dates in the left column | The dates become a **date axis** (Section 4) |
| Numbers in the left column, such as years or hours stored as numbers | Excel plots them as one more series instead of using them as labels (unless that column's header cell is empty) |

> ⚠️ ED_Hourly stores its hours as text (00:00, 01:00 …), so Excel uses them as labels. If a category column holds numbers
> such as 2024 and 2025, Excel draws them as a series of tiny columns. Fix it in **Chart Design → Select Data**: remove the
> unwanted series, click **Edit** under *Horizontal (Category) Axis Labels*, and select the label cells.

**Select Data** (**Chart Design → Select Data**) is where you repair any chart whose data came out wrong. It lists each
series with **Add**, **Edit**, and **Remove** buttons, and the category labels beside them. Its **Hidden and Empty Cells**
button controls two behaviors that surprise people:

- Charts skip rows that are hidden, including rows hidden by a filter. Tick **Show data in hidden rows and columns** if you
  want them plotted.
- Empty cells can show as gaps, as zero, or (for lines) as a line connecting the neighbors.

> 💡 **Tip:** Build charts from an Excel Table (Lesson 3.1). When you add a row to the Table, every chart that uses its
> columns extends automatically, even if you selected the cells by address. A chart built on a plain range that isn't a
> Table ignores a new row 26 until you edit its data in **Select Data**. Every data sheet in this lesson's workbook is a
> Table.

### 3. Chart anatomy and the Format pane

| Element | What it is |
|---|---|
| **Chart title** | The text above the plot |
| **Axes** | The **category axis** (along the bottom of a column chart) and the **value axis** (the numbers) |
| **Axis titles** | What each axis measures, with units |
| **Gridlines** | Reference lines across the plot |
| **Data series** | One set of plotted numbers, such as Arrivals |
| **Data point** | One value in a series, such as the 09:00 column |
| **Data labels** | Values printed on the points |
| **Legend** | The key that matches colors to series |
| **Plot area** and **chart area** | The area inside the axes, and the whole chart object |
| **Trendline** | A line or curve fitted to a series (Section 10) |

There are three ways to add, remove, and format these elements:

- **Floating buttons (Windows).** Three buttons appear beside a selected chart: **Chart Elements** (**+**) adds or removes
  elements, **Chart Styles** (a brush) applies preset looks and colors, and **Chart Filters** (a funnel) hides series or
  categories without touching the data.
- **The ribbon (Windows and Mac).** **Chart Design → Add Chart Element** lists every element. Excel for Mac has no floating
  buttons, so use this menu.
- **The Format pane.** Double-click any element, or select it and press **Ctrl + 1** (Mac: **⌘ + 1**). The pane's contents
  change with the selection. Its icons group the settings: **Fill & Line**, **Effects**, **Size & Properties**, and the
  element's own options (for example **Axis Options** or **Series Options**).

**Read exact values.** Hover the pointer over any column, bar, or point, and a tooltip shows its series, category, and
value, for example *Series "Arrivals" Point "09:00" Value: 313*. Most practice tasks ask you to read a value this way. Data
labels (Section 5) show the values permanently.

To format one point only, click the series once to select every point, pause, and click the point again. Now only that point
is selected. On Windows, the dropdown in **Format → Current Selection** lists every element of the chart, which helps with
small targets such as an axis title.

### 4. Axes: scale, number format, and order

Double-click an axis to open **Format Axis**. These **Axis Options** are the ones you'll use most:

| Option | What it does | Use it when |
|---|---|---|
| **Bounds: Minimum / Maximum** | Fixes the ends of the axis. **Reset** returns to automatic | Comparing charts on one scale, or capping a percentage axis at 100% |
| **Units: Major / Minor** | Sets the spacing of gridlines and labels | Labels crowd together |
| **Display units** | Shows 1,500,000 as 1.5 with a *Millions* label | Dollar axes |
| **Logarithmic scale** | Makes each gridline 10 times the one below | Values span several orders of magnitude, such as clinic-visit and ICU-stay charges on one chart. It can't show zero or negatives |
| **Categories in reverse order** | Flips the order of the categories | Bar charts (Section 6) |
| **Axis crosses** | Sets where the other axis meets this one | After you reverse the categories |
| **Axis Type: Date / Text** | Spaces points by real time, or evenly | Date categories (below) |
| **Number** | Sets the label format. **Linked to source** copies the cells' format | Percentages, dollars, dates |

**Bars and columns start at zero.** Readers judge a bar by its length, so cutting off the bottom of the axis distorts every
comparison. Suppose two units run at 92% and 88% occupancy. On an axis that starts at 85%, the bars are 7 and 3 points
tall, so the first looks more than twice as big even though the values differ by less than 5%. Line charts are different,
because their message is in the slope. A line chart's axis can start above zero as long as it's clearly labeled.

> ⚠️ Excel sets the bounds automatically. When the values sit in a narrow band far from zero, it can start a column chart's
> axis above zero on its own. Check the value axis of every bar and column chart, and set **Minimum** to 0 when it isn't.

**Date axis vs. text axis.** When the category labels are real dates, Excel builds a **date axis**. It spaces points by
actual time and chooses a **base unit** (days, months, or years), so a missing month leaves a visible gap. A **text axis**
spaces the labels evenly, whatever they say. ED_Monthly's Month column holds real dates (the first of each month), so a line
chart of it gets a date axis with a base unit of months. Change the label format under **Number**, for example to `mmm yy`.

### 5. Data labels and legends

**Data labels** print the values on the chart, so readers don't have to trace gridlines. Add them with **Chart Elements →
Data Labels** (Mac: **Chart Design → Add Chart Element → Data Labels**), then choose **More Options** to set what each label
shows:

| Label option | Shows | Typical use |
|---|---|---|
| **Value** | The number | Columns, bars, the last point of a line |
| **Percentage** | Each point's share of the total (pie and doughnut charts only) | Payer mix |
| **Category Name** | The category label | Pie slices, instead of a legend |
| **Series Name** | The series name | Naming a line at its end, instead of a legend |
| **Value From Cells** | Text from any range you choose | Notes such as "Target" or "New process" |

Under **Number** in the same pane, set the label format, such as **Percentage** with 1 decimal place. The cells keep their
own format.

**Direct labels beat legends.** A legend makes readers look back and forth to match colors. Delete the legend when there's
only one series. When there are a few lines, label each one at its last point: select the line, click its last point again,
add a data label, and tick **Series Name** instead of **Value**.

### 6. Column and bar charts

| Variant | Shows | Bluestone example |
|---|---|---|
| **Clustered** | Values side by side within each category | Arrivals by hour; three hospitals per quarter |
| **Stacked** | Parts stacked into a total | ED visits by hospital, stacked into the system total |
| **100% stacked** | Each category's mix, with every column at 100% | Payer mix at each hospital |

A **bar chart** is a column chart turned on its side. Choose bars when the category names are long (service lines, denial
reasons), when there are many categories, or when you're showing a ranking.

> ⚠️ **Bar charts draw the first row at the bottom.** Sort a table from highest to lowest and chart it, and the highest bar
> lands at the bottom. You don't have to re-sort the data. Double-click the vertical (category) axis and tick **Categories
> in reverse order**. That also moves the value axis to the top, so set **Horizontal axis crosses → At maximum category**
> to bring it back down.

Two settings in **Format Data Series → Series Options** control the look:

- **Gap Width** is the space between columns, as a percentage of one column's width. Around 50–80% looks solid.
- **Series Overlap** (clustered charts) makes the columns within a cluster touch (0%) or overlap (above 0%).

**Highlight the bar that matters.** Color every bar gray. Then select the one bar you're talking about (click the series,
then click that bar again) and give it a strong color. The reader's eye goes straight to it, and the title explains why.

### 7. Line charts

Use a **line chart** when the categories are ordered in time and the message is the change: a seasonal pattern, a step after
a new process, or steady growth. Keep it readable:

- **Plot four or five lines at most.** More lines turn into spaghetti. Split them into several small charts with the same
  axis scale.
- **Watch the scale.** On a shared axis, the biggest series sets the scale and flattens the small ones. Bluestone Memorial
  sees several times the ED volume of Ashby Falls or Cedar Ridge, so their lines hug the bottom of a combined chart. Give
  each hospital its own chart, or use sparklines (Section 14) to compare shapes.
- **Use markers sparingly.** They help with a handful of points. With 24 points, a plain line is cleaner.
- **Don't smooth.** **Smoothed line** (**Format Data Series → Fill & Line**) bends the line between points, so it can show
  values that never happened.
- **Use lines only for ordered categories.** A line across service lines suggests a progression that doesn't exist. Use
  columns or bars instead.

> 💡 **Tip:** To see the trend through a seasonal series, add a moving-average trendline: right-click the line → **Add
> Trendline → Moving Average** with **Period** `3`. A 12-period moving average smooths out a full year of seasons.

### 8. Pie and doughnut charts

A **pie chart** shows how one whole splits into parts. It works when there are only a few parts (two to five), the parts
add up to a meaningful whole (all 2025 encounters), and one or two parts are clearly bigger than the rest.

To make a pie readable, sort the data from largest to smallest so the slices run clockwise from biggest to smallest,
starting at 12 o'clock. (**Format Data Series → Angle of first slice** rotates the pie if you need to.) Label the slices with
**Category Name** and **Percentage** instead of a legend.

| Situation | Better choice |
|---|---|
| Six or more slices, or slices of similar size | A sorted bar chart |
| Comparing the mix of two hospitals or two years | A 100% stacked bar, one bar per hospital or year |
| Showing a change over time | A line or column chart |

A **doughnut chart** is a pie with a hole, and some designers put the total in the middle with a text box. It has the same
limits as a pie.

> ⚠️ Never use a 3-D or "exploded" pie. Tilting the pie makes the slices at the front look bigger than equal slices at the
> back.

### 9. Distributions: histograms and box and whisker charts

A **distribution** shows how values spread out: where most of them sit, how wide the spread is, and whether there's a long
tail of extreme values. Length of stay (LOS) is a classic example, because most stays are short and a few are very long.

**Histograms.** A **histogram** sorts values into equal-width ranges called **bins** and draws one column per bin, with no
gaps between them. Select one column of numbers (with its header) and choose **Insert → Insert Statistic Chart → Histogram**
(Mac: the same **Statistic Chart** button on the Insert tab). Then double-click the horizontal axis and set the bins:

| Bins option | What it does |
|---|---|
| **By Category** | Groups identical text categories and adds up their values, instead of binning numbers |
| **Automatic** | Excel chooses the bin width (using Scott's normal reference rule) |
| **Bin width** | Makes every bin this wide, for example 1 day |
| **Number of bins** | Divides the range into this many equal bins |
| **Overflow bin** | Puts every value **above** this number into one last bin, labeled with > |
| **Underflow bin** | Puts every value **at or below** this number into one first bin, labeled with ≤ |

The bin labels use interval notation. A square bracket includes its end point and a round bracket excludes it, so **(3, 4]**
means "more than 3, up to and including 4." A stay of exactly 4.0 days falls in (3, 4], not (4, 5]. Overflow and underflow
bins keep a few extreme values from stretching the axis. LOSDays on the Stays sheet is the exact length of stay (discharge
time minus admit time, in days, to 1 decimal place). It runs from 1.0 to 17.0 days in the sample, so without
an overflow bin the right-hand side of the chart would be a row of nearly empty bins.

> 📋 Histogram, Pareto, box and whisker, and waterfall charts need Excel 2016 or later. In older versions, count each bin
> with COUNTIFS (Lesson 2.5), chart the counts as a column chart, and set **Gap Width** to 0%. Lesson 4.5 builds histograms
> with FREQUENCY and the Analysis ToolPak.

**Box and whisker charts.** A **box and whisker chart** (a box plot) summarizes a distribution with a few landmarks, which
makes it the best way to compare distributions across groups:

| Part | Meaning |
|---|---|
| Box | The middle 50% of the values, from the first quartile (Q1) to the third quartile (Q3) |
| Line inside the box | The median |
| X marker | The mean |
| Whiskers | Reach to the lowest and highest values that lie within 1.5 times the box's length (the interquartile range) of the box |
| Dots beyond the whiskers | Outliers |

To compare length of stay by service line, select B1:B301 on the Stays sheet, Ctrl+click (Mac: ⌘+click) D1:D301, and
choose **Insert → Insert Statistic Chart → Box and Whisker**. Excel draws one box for each service line, and you can see at
once that Behavioral Health has the highest median length of stay. In **Format Data Series** you can show or hide inner points,
outlier points, mean markers, and the mean line, and choose the **Quartile Calculation** (Inclusive median or Exclusive
median). The two methods can draw slightly different boxes, so note which one you used.

### 10. Scatter charts and trendlines

A **scatter chart** (also called an XY chart) plots pairs of numbers, one point per row, with one measure across and the
other up. Use it to ask whether two measures move together.

- Select two adjacent columns with X on the left and Y on the right, then choose **Insert → Insert Scatter (X, Y) or Bubble
  Chart → Scatter**. For columns that aren't adjacent, Ctrl+click (Mac: ⌘+click) the second one. The left-hand column still
  becomes X.
- If Excel puts the wrong measure on X, open **Select Data**, select the series, click **Edit**, and swap the **Series X
  values** and **Series Y values** ranges.
- A line chart isn't a substitute. It places points evenly in row order and ignores the X values.

**Trendlines.** A **trendline** is a line or curve fitted to a series. Right-click the series → **Add Trendline**, or use
**Chart Elements → Trendline** (Mac: **Chart Design → Add Chart Element → Trendline**). Choose the type in **Format
Trendline**:

| Type | Fits | Notes |
|---|---|---|
| **Linear** | A straight line, y = mx + b | The default and the one to try first |
| **Exponential** | Growth or decline by a constant percentage | Every Y value must be positive |
| **Logarithmic** | A fast change that levels off | Every X value must be positive |
| **Polynomial** | A curve with bends (order 2 to 6) | Order 2 has one bend. Higher orders chase noise |
| **Power** | y = a·xᵇ | X and Y must be positive |
| **Moving Average** | The average of the last N points | For time series. It has no equation or R² |

Tick **Display Equation on chart** and **Display R-squared value on chart** to show the fit. To control the decimals, click
the label and set **Format Trendline Label → Number**. **Forecast Forward** extends the line past the data, and **Set
Intercept** forces the line through a chosen value where X is 0.

**Read the equation.** In y = mx + b, the **slope** (m) is the change in Y for each extra unit of X, and the **intercept** (b)
is the predicted Y when X is 0. If a chart of supply cost against patient days showed y = 85x + 1,200, each extra patient day
would add about \$85 of supplies.

**Read R².** **R²** (R-squared) runs from 0 to 1 and tells you how much of the variation in Y the line explains. Near 1, the
points hug the line. Near 0, the points form a cloud and the line tells you little. A steep slope can look impressive while
R² shows the relationship is weak, so always show both.

**Check the trendline with formulas.** A linear trendline is the least-squares line, so these worksheet functions return the
same numbers the chart shows:

```
=SLOPE(known_y's, known_x's)       the m in y = mx + b
=INTERCEPT(known_y's, known_x's)   the b
=RSQ(known_y's, known_x's)         R²
=CORREL(array1, array2)            r, the correlation coefficient (R² = r² for a straight line)
```

> ⚠️ SLOPE and INTERCEPT take the **Y range first**. `=SLOPE(Stays!D2:D301,Stays!C2:C301)` is days of stay (Y) per year of
> age (X). Swap the ranges and you get a different, wrong number. RSQ and CORREL give the same result in either order, but
> keep the Y-first habit.

> ⚠️ **Correlation isn't causation.** A trendline shows that two measures move together in this data, not that one causes
> the other. Lesson 4.5 covers correlation and regression in depth.

### 11. Combo charts and secondary axes

A **combo chart** mixes chart types, usually columns for one series and a line for another. A **secondary axis** is a second
value axis on the right with its own scale. Use the pair when two measures share the same categories but differ in units or
size, such as ED arrivals (hundreds of visits) and the minutes patients wait to see a provider (tens of minutes).

1. Select the categories and both series.
2. Choose **Insert → Insert Combo Chart → Clustered Column – Line on Secondary Axis**.
3. To convert an existing chart, choose **Chart Design → Change Chart Type → Combo**. For each series, pick a chart type,
   and tick **Secondary Axis** for the one that needs its own scale.

Another route works in every version and on the Mac: select one series, open **Format Data Series → Series Options → Plot
Series On → Secondary Axis**, then change that series to a line. On Windows, right-click it → **Change Series Chart Type**. On
a Mac, keep it selected and choose **Chart Design → Change Chart Type**.

> ⚠️ A dual-axis chart can mislead, because changing either axis's scale moves the lines wherever you like. Give both axes
> titles with units, and color each axis's labels to match its series. If the two measures don't share a story, use two
> charts, one above the other, instead.

### 12. Pareto charts

The **Pareto principle** says that a few causes often account for most of an effect. People call them the "vital few" and
often quote the split as 80/20. A **Pareto chart** shows it: columns sorted from largest to smallest, plus a line on a
secondary axis that shows the cumulative percentage from 0% to 100%. Revenue-cycle and quality teams use it to decide which
denial reasons or safety events to work on first.

**The built-in Pareto (Excel 2016 or later).** Select the category column and the value column, and choose **Insert → Insert
Statistic Chart → Pareto**. Excel sorts the categories, adds up rows that share a category label, and draws the cumulative
line. You don't need to sort the data first.

**A Pareto by hand (any version).** It takes a few more steps, but you control every part of it, so you can add an 80%
reference line or label the cut-off:

1. Sort the table by the value column, largest first.
2. Add a cumulative-share column. With the values in C2:C8, type `=SUM($C$2:C2)/SUM($C$2:$C$8)` in D2 and fill down. In
   a Table, Excel fills the column for you. `$C$2:C2` is an **expanding range** (Lesson 3.3). Its start is absolute and its
   end is relative (Lesson 1.5), so the start stays fixed while the end moves down a row at a time. Each row adds up every
   value from the top to that row and divides by the grand total.
3. Select the categories, the values, and the cumulative shares, and insert a **Clustered Column – Line on Secondary Axis**
   combo chart.
4. Set the secondary axis **Minimum** to 0 and **Maximum** to 1 (100%), and set the columns' **Gap Width** to about 10%.

### 13. Waterfall charts

A **waterfall chart** (also called a bridge chart) shows how a starting total becomes an ending total through a series of
increases and decreases. Finance teams use it to explain budget variance: start at the budgeted margin, add or subtract each
category's variance, and finish at the actual margin.

1. Lay out two columns: labels, and signed amounts. The first row is the starting total, the middle rows are the changes
   (positive when favorable, negative when unfavorable), and the last row is the ending total.
2. Select both columns and choose **Insert → Insert Waterfall, Funnel, Stock, Surface, or Radar Chart → Waterfall** (Mac:
   the **Waterfall** button on the Insert tab).
3. Excel first draws every row as a floating change, the totals included. Click the last bar once to select the series and
   again to select just that bar, then right-click it → **Set as Total**. Do the same for the first bar. Totals then stand on
   the axis.
4. The legend shows **Increase**, **Decrease**, and **Total**, each in its own color. The colors come from the
   workbook's theme, so they vary between Excel versions. To recolor a group, click the legend once, then click that
   group's entry. Every bar in the group is now selected, so choose **Format → Shape Fill** and pick a color.

> 💡 **Tip:** Keep one sign convention and write it on the sheet, as the Budget sheet does: positive means favorable. For an
> expense, favorable means *under* budget, so its variance is budget minus actual.

> 📋 Before Excel 2016, people built waterfalls as stacked column charts with an invisible "base" series under each floating
> bar. The hidden Chart Key sheet in this workbook uses that method, because the program that generates the workbook can't
> create Excel 2016's chart types.

### 14. Sparklines

A **sparkline** is a tiny chart inside one cell. It shows the shape of a series right next to the numbers, which makes it
ideal for tables and dashboards (Lesson 4.6).

| Type | Shows | Example |
|---|---|---|
| **Line** | The shape of a trend | 24 months of ED visits for each hospital |
| **Column** | The size of each value | Monthly admissions |
| **Win/Loss** | Only whether each value is positive or negative | Months over or under budget |

1. Select the empty cells where the sparklines will go: one cell per series, all in one row or one column.
2. Choose **Insert → Sparklines → Line** (or **Column**, or **Win/Loss**).
3. In **Data Range**, select the numbers. **Location Range** already shows the cells you selected. Excel makes one sparkline
   for each row or column of the data, matching the location cells.

When you select a sparkline cell, the **Sparkline** tab appears (in older versions: **Sparkline Tools → Design**):

| Option | What it does |
|---|---|
| **Show:** High Point, Low Point, First Point, Last Point, Negative Points, Markers | Marks those points in a contrasting color |
| **Style**, **Sparkline Color**, **Marker Color** | Sets the colors |
| **Axis → Vertical Axis Minimum / Maximum Value Options** | **Automatic for Each Sparkline** (the default), **Same for All Sparklines**, or a custom value |
| **Axis → Date Axis Type** | Spaces the points by real dates |
| **Group / Ungroup** | Sparklines created together share their settings until you ungroup them |
| **Clear** | Removes the selected sparklines |

> ⚠️ By default each sparkline stretches to fill its cell, from its own minimum to its own maximum. That's right for reading
> each series' shape, but a hospital whose volume barely moves looks as volatile as one with big swings. To compare sizes
> across sparklines, set both the minimum and the maximum to **Same for All Sparklines**.

> ⚠️ The **Delete** key doesn't remove a sparkline. Select the cells and choose **Sparkline → Clear**.

> 💡 **Tip:** A sparkline sits behind the cell's contents, so you can still type a label in the same cell. Widen the column
> and raise the row height to make the sparkline easier to read.

### 15. Dynamic chart titles

A **dynamic title** is a chart title linked to a cell, so it changes when the data changes. A chart title can hold typed
text or a link to one cell, but not a formula. So you build the full title in a cell with a formula, then link the title to
that cell.

**Step 1: build the text in a cell.** Join text and values with `&`. Wrap numbers and dates in **TEXT** (Lesson 2.2) so they
keep a readable format. Without TEXT, a date shows up as its serial number and a rate as 0.183615819….

```
="Highest 30-day readmission rate: "&XLOOKUP(MAX(tblReadmits[ReadmitRate]),tblReadmits[ReadmitRate],tblReadmits[ServiceLine])&" ("&TEXT(MAX(tblReadmits[ReadmitRate]),"0.0%")&")"
```

This returns *Highest 30-day readmission rate: Cardiovascular (18.4%)*. XLOOKUP needs Microsoft 365 or Excel 2021 or later.
In older versions, use `INDEX(tblReadmits[ServiceLine],MATCH(MAX(tblReadmits[ReadmitRate]),tblReadmits[ReadmitRate],0))`
instead (Lesson 2.6). Useful TEXT format codes for titles:

| Format code | Example result |
|---|---|
| `"mmm yyyy"` | Jan 2025 |
| `"mmmm yyyy"` | January 2025 |
| `"#,##0"` | 12,345 |
| `"0.0%"` | 18.4% |
| `"$#,##0"` | \$12,345 |

> 📋 TEXT's date codes follow the language of your Excel installation. These codes work in English versions. In some other
> languages the year and day letters differ (German Excel uses `JJJJ` for the year, for example).

**Step 2: link the title.**

1. Click the chart title to select it.
2. Click in the formula bar and type `=`.
3. Click the cell that holds the title text, then press **Enter**. If that cell is on another sheet, click the sheet's tab
   first, then the cell.

The formula bar now shows a reference such as `=Readmits!$F$2`. The same trick works for axis titles, text boxes, and a single
data label.

> 💡 **Tip:** Use structured references to the Table (Lesson 3.1) in the title formula. When next month's row is added, the
> chart, the date range, and the title all update together.

**Write titles that state the finding.** A label title such as *Readmission rates by service line* tells readers what
they're looking at. A headline title such as *Cardiovascular patients are readmitted most often* tells them what to notice.
Use headline titles for reports and presentations.

### 16. Design for clarity and accessibility

A good chart makes one point quickly, and everyone can read it, including colleagues with low vision or color-vision
deficiency (about 1 in 12 men and 1 in 200 women).

**Declutter.** Every element should earn its place:

- Delete the legend when there's one series, or replace it with direct labels.
- Lighten or delete the gridlines when data labels carry the values.
- Remove borders, shadows, gradients, and 3-D effects.
- Cut decimals. 18.4% reads faster than 18.36158%.
- Keep text horizontal. If the category labels need rotating, switch to a bar chart.

**Use color with meaning.**

- Make most of the chart gray, and color only what the story is about.
- Keep each thing's color the same across charts. If Ashby Falls is orange in one chart, keep it orange in the next.
- Avoid pairing red with green, because they look alike to people with the most common form of color-vision deficiency. Use
  blue and orange instead.
- Don't rely on color alone. Add labels, markers, or different line styles.

The Okabe–Ito palette was designed to stay distinguishable for people with color-vision deficiency. The reference charts in
this lesson's Chart Key use it:

| Color | Hex | Color | Hex |
|---|---|---|---|
| Blue | `#0072B2` | Orange | `#E69F00` |
| Sky blue | `#56B4E9` | Vermillion | `#D55E00` |
| Bluish green | `#009E73` | Reddish purple | `#CC79A7` |
| Yellow | `#F0E442` | Black | `#000000` |

To apply one, select the element and choose **Format → Shape Fill → More Fill Colors**, then enter the hex code. On Windows,
Microsoft 365 has a **Hex** box on the dialog's **Custom** tab. On a Mac, the color picker opens: choose its sliders tab, then
**RGB Sliders**, and type the code in **Hex Color #**. For quick monochrome sets, use **Chart Design → Change
Colors**.

**Add alt text.** **Alt text** is a short description that screen readers announce. Right-click the chart → **Edit Alt
Text** (some Microsoft 365 versions call it **View Alt Text**). Describe the chart type, what it shows, and the takeaway, for
example: *"Line chart of monthly ED visits by hospital, January 2024 to December 2025. Bluestone Memorial peaks every
winter."* **Review → Check Accessibility** lists any chart that's missing alt text.

**Check every chart before you share it:**

| Check | Why |
|---|---|
| The title states the finding | Readers know what to look for |
| Bar and column axes start at zero | Lengths stay honest |
| Axes have titles with units | Nobody wonders whether it's visits or minutes |
| Categories are sorted, unless they have a natural order such as months | Rankings are visible |
| No 3-D effects, and no pie with many slices | Shapes aren't distorted |
| Colors are color-blind safe and mean something | Everyone sees the same story |
| Alt text is filled in | Screen-reader users get the message |
| The source and date range are noted | Readers can trust and trace the numbers |

### 17. Common chart mistakes

| Mistake | Why it misleads | Fix |
|---|---|---|
| A bar or column axis that starts above zero | Exaggerates differences | Set **Minimum** to 0 |
| 3-D charts | Perspective distorts sizes | Use 2-D charts |
| A pie with many or similar slices | Angles are hard to compare | Use a sorted bar chart |
| Rainbow colors (**Vary colors by point**) | Color carries no meaning, and the legend repeats the axis labels | One color plus one highlight |
| Two value axes without titles | Readers read values off the wrong scale | Title and color both axes, or use two charts |
| A line across unordered categories | Implies a progression | Use columns or bars |
| Smoothed lines | Shows values between points that never happened | Use straight segments |
| A total plotted next to its own parts | The total dwarfs the parts | Stacked columns, or a separate chart |
| Filtering the data and forgetting the chart | Charts skip hidden rows | **Select Data → Hidden and Empty Cells** |
| Averaging the rates of groups of very different sizes | Small groups count as much as big ones | Total numerator ÷ total denominator (Lesson 1.4) |

### 18. Reuse your work: templates, copied formats, and PivotCharts

- **Chart templates.** Once a chart looks the way you want, right-click it → **Save as Template**. Excel saves a `.crtx`
  file. Apply it to another chart with **Chart Design → Change Chart Type → Templates** (on Windows also **Insert →
  Recommended Charts → All Charts → Templates**). On Windows, right-click a template in that list → **Set as Default Chart**
  to make **Alt + F1** use it.
- **Copy formatting between charts (Windows).** Copy the formatted chart (**Ctrl + C**), select another chart, and choose
  **Home → Paste → Paste Special → Formats**.
- **PivotCharts.** A **PivotChart** plots exactly what its PivotTable shows, and it changes when you filter or rearrange the
  pivot (Lesson 3.4). Insert one with **PivotTable Analyze → PivotChart**. A PivotChart can't be a scatter, bubble, or stock
  chart, or any of the Excel 2016 types such as a histogram, Pareto, box and whisker, or waterfall. For those, build a regular
  chart from the values.

### 19. Keyboard shortcuts

| Action | Windows | Mac |
|---|---|---|
| Insert the default chart on the current sheet | **Alt + F1** | Use the **Insert** tab |
| Insert the default chart on a new chart sheet | **F11** | Use the **Insert** tab, then **Chart Design → Move Chart** |
| Open the Format pane for the selected element | **Ctrl + 1** | **⌘ + 1** |
| Add a non-adjacent range to the selection | **Ctrl** + drag | **⌘** + drag |
| Quick Analysis (its Charts and Sparklines tabs) | **Ctrl + Q** | Not available |
| Repeat the last action | **F4** or **Ctrl + Y** | **⌘ + Y** |
| Undo | **Ctrl + Z** | **⌘ + Z** |
| Link a title to a cell | Select the title, type **=** in the formula bar, click the cell, press **Enter** | Same |

### 20. Version notes

| Feature | Available in |
|---|---|
| Recommended Charts | Excel 2013 or later (Windows), Excel 2016 or later (Mac) |
| Chart Elements, Chart Styles, and Chart Filters buttons | Excel 2013 or later for Windows only |
| Data labels from cells (**Value From Cells**) | Excel 2013 or later (Windows), Excel 2016 or later (Mac) |
| Histogram, Pareto, box and whisker, waterfall, treemap, and sunburst charts | Excel 2016 or later (Windows and Mac) and Microsoft 365 |
| Funnel and map charts | Excel 2019 or later and Microsoft 365 |
| Sparklines | Excel 2010 or later (Windows), Excel 2011 or later (Mac) |
| Excel for the web | Shows all of these charts and creates most of them, with fewer formatting options. Use desktop Excel for this lesson |

## 🧪 Hands-on practice

Download [`3.5-charts-visualization.xlsx`](3.5-charts-visualization.xlsx) and open the **Practice** sheet. The **Check**
column turns green when the value you read from your chart is right.

<!-- BEGIN GENERATED: practice -->
Each task names the sheet to work on. Build each chart on the same sheet as its data, to the right of the table, then type the number or name your chart shows into the yellow cell. The Check column can't see your chart, so every task asks a question the finished chart answers. The hidden Chart Key sheet shows a reference version of each chart. The yellow LWBSRate and CumulativePct columns are for the Bonus.

| # | Task | Hint |
|:-:|------|------|
| 1 | On ED_Hourly, select A1:B25 (ArrivalHour and Arrivals) and insert a clustered column chart. Which hour of the day had the most ED arrivals in 2025? Type the hour as a number from 0 to 23 (for example, 9 for the 09:00 hour). | Insert → Insert Column or Bar Chart → Clustered Column. Hover over the tallest column |
| 2 | On ED_Monthly, select A1:D25 (Month and the three hospitals) and insert a line chart. In which month did Cedar Ridge have its most ED visits? Type the month and year (for example, Jun 2025). | Insert → Insert Line or Area Chart → Line. Hover over the highest point of the Cedar Ridge line |
| 3 | On Readmits, sort the table by ReadmitRate from largest to smallest. Then select ServiceLine and ReadmitRate (A1:A9, then Ctrl+click D1:D9, or ⌘+click on a Mac) and insert a clustered bar chart. Before you change anything else, which service line's bar is at the TOP of the chart? (Afterwards, fix the order so the highest rate is on top.) | Read the chart, not the table. Guide section 6 explains the bar order and how to fix it |
| 4 | The Makeover sheet has a colleague's column chart of the same readmission rates. Its vertical axis starts at 5%, not 0%. Measured from that axis, how many times taller is the Cardiovascular bar than the Women & Children bar? Round to 1 decimal place. Then fix the chart in place, using the checklist in Guide section 16, and add alt text. | A bar's drawn height is its value minus the axis minimum (5%) |
| 5 | On PayerMix, insert a pie chart of 2025 encounters by PayerType. Add data labels that show the Percentage (not the Value), formatted with 1 decimal place. What does the Commercial label show? | Format Data Labels → Label Options: tick Percentage, untick Value. Then Number → Percentage, 1 decimal |
| 6 | On Stays, insert a histogram of LOSDays (all 300 stays). Format the horizontal axis with Bin width 1, Overflow bin 10, and Underflow bin 1. How many stays fall in the overflow bin (longer than 10 days)? | Select D1:D301, then Insert → Insert Statistic Chart → Histogram. Double-click the horizontal axis |
| 7 | On Stays, select D1:E301 (LOSDays and TotalCharges) and insert an XY scatter chart, so LOSDays is on the horizontal axis. Add a linear trendline and display its equation. By how many dollars do charges rise for each extra day in the hospital? Confirm with SLOPE and round to the nearest dollar. | Insert → Insert Scatter (X, Y) or Bubble Chart → Scatter. Right-click a point → Add Trendline, then tick Display Equation on chart. SLOPE takes the Y range first |
| 8 | Make a second scatter chart on Stays with AgeAtAdmit on the horizontal axis and LOSDays on the vertical axis (C1:D301). Add a linear trendline and display the R-squared value. What is R²? Confirm with RSQ and enter it to 4 decimal places. | RSQ(known_y's, known_x's) |
| 9 | Back on ED_Hourly, select A1:C25 and insert a combo chart with Arrivals as clustered columns and AvgDoorToProviderMin as a line on the secondary axis. Which arrival hour has the longest average wait to see a provider? Type the hour as a number from 0 to 23. | Insert → Insert Combo Chart → Clustered Column – Line on Secondary Axis |
| 10 | On Denials, select A1:B8 (DenialReason and Claims) and insert a Pareto chart. What cumulative percentage does the line reach at the second bar (the top two reasons together)? Enter it as a percentage to 1 decimal place. | Insert → Insert Statistic Chart → Pareto. Confirm with LARGE and SUM |
| 11 | On Budget, select A4:B14 and insert a waterfall chart. Set the first and last bars as totals. Which step is the largest drop (the longest downward bar)? Type the Step name as it appears in the table. | Insert → Insert Waterfall, Funnel, Stock, Surface, or Radar Chart → Waterfall. Right-click a bar → Set as Total |
| 12 | On ED_Monthly, insert line sparklines in the yellow cells B27:D27 (Data Range B2:D25), one per hospital. Turn on First Point and Last Point markers. For how many of the three hospitals is the last point (Dec 2025) higher than the first point (Jan 2024)? | Insert → Sparklines → Line. Then the Sparkline tab → Show → First Point, Last Point |
| 13 | Write a formula in the yellow cell that builds this chart title from tblEDMonthly: ED visits by facility, [first month] to [last month] ([total visits] visits). Show each month as a three-letter month and year, and the total with a thousands separator. Example of the pattern: ED visits by facility, Mar 2023 to Feb 2024 (9,876 visits). Then link the title of your Task 2 line chart to this cell. | TEXT(MIN(…),"mmm yyyy") and TEXT(SUM(…),"#,##0"), joined with & |
<!-- END GENERATED: practice -->

## ✅ Answer key

The workbook has three hidden key sheets (right-click any sheet tab → **Unhide…**). The **Answer Key** and **Bonus Key**
list every answer, and their *Live result* column runs a formula that confirms each chart reading. The **Chart Key** shows a
reference version of every chart. The same answers are below, collapsed so you don't see them by accident.

<!-- BEGIN GENERATED: answers -->
<details>
<summary><b>🔑 Show the answer key</b> — Try every task before opening this.</summary>

**1. Column chart: busiest arrival hour**

- **Answer:** 16
- **Solution:**

1. On **ED_Hourly**, select **A1:B25**: the ArrivalHour labels and the Arrivals numbers, headers included.
2. Choose **Insert → Insert Column or Bar Chart → 2-D Column → Clustered Column**. On Windows you can also press **Alt + F1**, which inserts the default chart type (clustered column unless someone changed it).
3. Drag the chart to the right of the table. Hover over the tallest column: the tooltip reads *Series "Arrivals" Point "16:00" Value: 408*.


The 16:00 hour had 408 arrivals, just ahead of 12:00 with 404. The quietest hour was 04:00 with 77. A column chart suits this question because it compares one number across categories, and the eye compares column heights quickly. ArrivalHour holds text labels such as 16:00, so Excel uses it as the category axis automatically. If the hours were numbers, Excel would plot them as a second series of short columns. Cross-check with a formula: `=VALUE(LEFT(XLOOKUP(MAX(tblEDHourly[Arrivals]),tblEDHourly[Arrivals],tblEDHourly[ArrivalHour]),2))`

**2. Line chart: Cedar Ridge's busiest month**

- **Answer:** Mar 2024
- **Solution:**

1. On **ED_Monthly**, select **A1:D25**.
2. Choose **Insert → Insert Line or Area Chart → 2-D Line → Line** (or **Line with Markers**). Excel recognizes the Month column as dates and builds a date axis, one point per month.
3. Hover over the highest point of the Cedar Ridge line, or click the line once and read the values in the tooltips.


Cedar Ridge peaked at 99 visits in Mar 2024. A line chart is the standard choice for a trend over time because the slope between points shows the change. Notice how flat the Cedar Ridge and Ashby Falls lines look: Bluestone Memorial (peak 473 in Jan 2025) sets the scale, so the smaller hospitals' swings shrink. When a small series matters, give it its own chart. Cross-check with a formula: `=XLOOKUP(MAX(tblEDMonthly[Cedar Ridge]),tblEDMonthly[Cedar Ridge],tblEDMonthly[Month])`

**3. Bar chart: which bar Excel puts on top**

- **Answer:** Women & Children
- **Solution:**

1. Click any ReadmitRate cell and choose **Data → Sort Largest to Smallest** (Mac: **Data → Sort**, or the column's filter button).
2. Select **A1:A9**, hold **Ctrl** (Mac: **⌘**), and select **D1:D9**.
3. Choose **Insert → Insert Column or Bar Chart → 2-D Bar → Clustered Bar**.
4. Read the top bar: it's Women & Children, the lowest rate.
5. To fix the order, double-click the vertical (category) axis to open **Format Axis**, then under **Axis Options** tick **Categories in reverse order**. Set **Horizontal axis crosses** to **At maximum category** so the value axis stays at the bottom.


A bar chart plots the first category next to the origin, which is the bottom of a bar chart. So a table sorted from highest to lowest produces a chart with the highest bar at the bottom and the lowest (Women & Children, 6.4%) at the top, which is upside down for a ranking. **Categories in reverse order** flips it without re-sorting the data. Bars suit this data better than columns because the service-line names are long and read easily on the left. Cross-check with a formula: `=XLOOKUP(MIN(tblReadmits[ReadmitRate]),tblReadmits[ReadmitRate],tblReadmits[ServiceLine])`

**4. Makeover: how much a truncated axis exaggerates**

- **Answer:** 9.6
- **Solution:**

1. Each bar is drawn up from the axis minimum, so its height is its rate minus 5%.
2. (0.1836 − 0.05) ÷ (0.0639 − 0.05) = **9.6**.
3. Fix the chart: double-click the vertical axis, and in **Format Axis → Axis Options → Bounds** set **Minimum** to `0` (or click **Reset** so Excel chooses 0). Turn off **Vary colors by point** (**Format Data Series → Fill & Line → Fill**), delete the legend, add data labels, and give the chart a title that states the finding.
4. Right-click the chart → **Edit Alt Text** and describe it, for example: *Column chart of 2025 30-day readmission rates by service line. Cardiovascular is highest at 18.4% and Women & Children lowest at 6.4%.*


The real ratio is 18.4% ÷ 6.4% = 2.9, but the truncated chart draws the Cardiovascular bar 9.6 times as tall. A bar's length is how readers judge its value, so every bar and column chart needs a value axis that starts at zero. Line charts are different: their message is in the slope, so a line chart's axis may start above zero as long as it's labeled. The other problems on the Makeover chart (rainbow colors and a legend that repeats the axis labels) add color without adding information. Cross-check with a formula: `=(MAX(tblReadmits[ReadmitRate])-0.05)/(MIN(tblReadmits[ReadmitRate])-0.05)`

**5. Pie chart: Commercial share**

- **Answer:** 37.7%
- **Solution:**

1. Click any cell in tblPayerMix and choose **Insert → Insert Pie or Doughnut Chart → 2-D Pie**.
2. Add labels: click the **Chart Elements** button (**+**) → **Data Labels → More Options…** (Mac: **Chart Design → Add Chart Element → Data Labels → More Data Label Options**).
3. In **Label Options**, tick **Percentage** and **Category Name** and untick **Value**. Under **Number**, choose **Percentage** with **1** decimal place.
4. Read the Commercial slice.


Excel computes each slice's percentage itself: 4,198 ÷ 11,145 = 37.7%. The default label format has no decimals, so it would round to 38%. A pie works here because there are only five parts of one whole and two of them dominate. Workers' Comp (0.5%) is barely a sliver, which is the usual limit of a pie: with more or smaller slices, a sorted bar chart compares parts far more accurately. Cross-check with a formula: `=XLOOKUP("Commercial",tblPayerMix[PayerType],tblPayerMix[Encounters])/SUM(tblPayerMix[Encounters])`

**6. Histogram: stays in the overflow bin**

- **Answer:** 9
- **Solution:**

1. On **Stays**, select **D1:D301** (the LOSDays column with its header).
2. Choose **Insert → Insert Statistic Chart → Histogram**. On a Mac, the same Statistic Chart button is on the Insert tab.
3. Double-click the horizontal axis. Under **Axis Options → Bins**, choose **Bin width** `1`, tick **Overflow bin** and type `10`, and tick **Underflow bin** and type `1`.
4. Hover over the last column, labeled **>10**.


A histogram answers "how are the values distributed?" by counting values in equal-width bins. Here the tallest bin is (3, 4] days with 75 stays, and the long right tail ends in the >10 overflow bin with 9 stays. That skew is why the median length of stay (3.90 days) sits below the mean (4.59 days). The few long stays pull the mean up but barely move the median. A bin label such as (3, 4] means "more than 3, up to and including 4". The overflow and underflow bins stop a few extreme stays from stretching the axis. Cross-check with a formula: `=COUNTIF(tblStays[LOSDays],">10")`

**7. Scatter + trendline: dollars per extra day**

- **Answer:** 4,922 (dollars per extra day)
- **Solution:**

1. Select **D1:E301**. The left column (LOSDays) becomes X, the right one (TotalCharges) becomes Y.
2. Choose **Insert → Insert Scatter (X, Y) or Bubble Chart → Scatter**.
3. Right-click any point → **Add Trendline…**. In **Format Trendline**, keep **Linear** and tick **Display Equation on chart** and **Display R-squared value on chart**.
4. The equation reads about *y = 4922x + 9225*. The slope is the number in front of x.
5. Confirm in the yellow cell: `=SLOPE(Stays!E2:E301,Stays!D2:D301)`.


A linear trendline is the least-squares line through the points, the same line that SLOPE and INTERCEPT calculate. Its slope says that each extra day adds about \$4,922 in charges on average. R² is 0.31, so length of stay explains about 31% of the variation in charges. The rest comes from what happened during the stay, such as surgery, ICU days, and imaging. SLOPE takes the Y range first, which is easy to get backwards. Cross-check with a formula: `=SLOPE(tblStays[TotalCharges],tblStays[LOSDays])`

**8. Scatter + trendline: R² for age vs. length of stay**

- **Answer:** 0.0185
- **Solution:**

1. Select **C1:D301** and choose **Insert → Insert Scatter (X, Y) or Bubble Chart → Scatter**.
2. Right-click a point → **Add Trendline…** → **Linear**, and tick **Display R-squared value on chart**.
3. The label reads *R² = 0.0185*.
4. Confirm in the yellow cell: `=RSQ(Stays!D2:D301,Stays!C2:C301)`.


R² = 0.0185 means age explains under 2% of the variation in length of stay. The trendline still slopes upward (about 0.017 days per year of age, or 0.17 days per decade), but the points scatter widely around it. Always show R² next to a trendline, because a line can be drawn through any cloud of points, related or not. Even a high R² shows association, not cause. Cross-check with a formula: `=RSQ(tblStays[LOSDays],tblStays[AgeAtAdmit])`

**9. Combo chart: hour with the longest wait**

- **Answer:** 13
- **Solution:**

1. Select **A1:C25** on **ED_Hourly**.
2. Choose **Insert → Insert Combo Chart → Clustered Column – Line on Secondary Axis**. (Or insert any chart, then **Chart Design → Change Chart Type → Combo**, set AvgDoorToProviderMin to **Line** and tick its **Secondary Axis** box.)
3. Add axis titles with **Chart Elements (+) → Axis Titles** (Mac: **Chart Design → Add Chart Element → Axis Titles**): *Arrivals* on the left and *Avg minutes to provider* on the right.
4. Hover over the highest point of the line.


Arrivals peak at 16:00, and the average wait peaks at 13:00 (57.1 minutes, against 27.8 at 04:00). The two series use different units (visits and minutes) and very different sizes, so the line needs its own axis. Otherwise it would be squashed flat. The two shapes match closely (correlation 0.96), so waits rise and fall with arrivals. That suggests staffing doesn't keep pace in the busy hours, but the chart alone can't prove the cause. Always title both axes on a dual-axis chart so nobody reads minutes off the arrivals scale. Cross-check with a formula: `=VALUE(LEFT(XLOOKUP(MAX(tblEDHourly[AvgDoorToProviderMin]),tblEDHourly[AvgDoorToProviderMin],tblEDHourly[ArrivalHour]),2))`

**10. Pareto: cumulative share of the top two denial reasons**

- **Answer:** 55.9%
- **Solution:**

1. Select **A1:B8** on **Denials**.
2. Choose **Insert → Insert Statistic Chart → Pareto**. Excel sorts the reasons from most to fewest claims and adds a cumulative-percentage line on a secondary axis that runs to 100%.
3. Hover over the line at the second bar.
4. To get the exact value, type `=(LARGE(Denials!B2:B8,1)+LARGE(Denials!B2:B8,2))/SUM(Denials!B2:B8)` and format it as a percentage with 1 decimal place.


Authorization Required (395 claims) and Medical Necessity (243) make up 638 of 1,142 denied claims. A Pareto chart ranks causes so a team can see the "vital few" worth fixing first. Here the first two reasons cover 55.9% of denials and the first four cover 84.3%. The built-in Pareto sorts the data for you, so the table itself can stay in any order. LARGE(range,1) and LARGE(range,2) return the largest and second-largest claim counts, which are the first two bars. The shorter cross-check formula that follows hands LARGE the array constant {1,2} (Lesson 2.5) so it returns both at once, and SUM adds them. Cross-check with a formula: `=SUM(LARGE(tblDenials[Claims],{1,2}))/SUM(tblDenials[Claims])`

**11. Waterfall: the largest unfavorable variance**

- **Answer:** Net patient revenue
- **Solution:**

1. Select **A4:B14** on **Budget**.
2. Choose **Insert → Insert Waterfall, Funnel, Stock, Surface, or Radar Chart → Waterfall** (Mac: **Insert → Waterfall**).
3. Click the first bar once to select the series and once more to select only that bar. Right-click it → **Set as Total**. Do the same for the last bar.
4. Find the longest bar in the **Decrease** color (the legend shows which color that is). Hover over it to read its Step name.


A waterfall shows how a starting total becomes an ending total through a series of increases and decreases. 4 West budgeted a margin of \$751,385 and earned \$735,410. Revenue came in \$366,593 under budget, the biggest drop, and salaries & wages saved \$284,237, the biggest rise, so expense control offset most of the revenue shortfall. Without **Set as Total**, Excel treats the last row as one more increase and floats it on top of the running total. Cross-check with a formula: `=XLOOKUP(MIN(tblBudget[Amount]),tblBudget[Amount],tblBudget[Step])`

**12. Sparklines: hospitals that ended higher than they started**

- **Answer:** 2
- **Solution:**

1. Select **B27:D27** on **ED_Monthly**.
2. Choose **Insert → Sparklines → Line**. In **Data Range** type `B2:D25`. **Location Range** already shows `$B$27:$D$27`. Click **OK**. Excel draws one sparkline per column.
3. On the **Sparkline** tab, tick **First Point** and **Last Point** (and **High Point** if you like).
4. Compare each sparkline's two end markers. To check, compare row 25 with row 2.


Bluestone Memorial and Cedar Ridge ended higher than they started. Ashby Falls ended lower (85 visits in Jan 2024, 79 in Dec 2025). A sparkline is a word-sized chart in a cell, so it shows each series' shape next to the numbers. By default every sparkline gets its own vertical scale, which is right for comparing a series with itself but wrong for comparing hospitals. When heights should be comparable, open **Sparkline → Axis** and set both the minimum and the maximum to **Same for All Sparklines**. Cross-check with a formula: `=SUMPRODUCT(--(ED_Monthly!$B$25:$D$25>ED_Monthly!$B$2:$D$2))`

**13. Dynamic chart title**

- **Answer:** ED visits by facility, Jan 2024 to Dec 2025 (12,292 visits)
- **Solution:**

```
="ED visits by facility, "&TEXT(MIN(tblEDMonthly[Month]),"mmm yyyy")&" to "&TEXT(MAX(tblEDMonthly[Month]),"mmm yyyy")&" ("&TEXT(SUM(tblEDMonthly[Total]),"#,##0")&" visits)"
```


Build the title in a cell, because a chart title can show text or one cell reference, but not a formula. To link it, click the chart title on ED_Monthly, type `=` in the formula bar, click the **Practice** sheet tab, click this yellow cell, and press **Enter**. The formula bar then shows a reference such as `=Practice!$D$18`. TEXT turns the dates and the total into formatted text. Without it, `&` would join the raw serial number 45292 instead of Jan 2024. Because the formula uses tblEDMonthly, adding January 2026 as a new row updates the chart and its title together.

</details>
<!-- END GENERATED: answers -->

## 🏆 Bonus challenge

<!-- BEGIN GENERATED: bonus -->
Two directors bring questions to Bluestone's monthly operations review. The ED medical director believes patients leave without being seen (LWBS) mainly in busy months, and the ED's target is an LWBS rate of 2% or less. The revenue-cycle director wants to know which denial reasons to work first, ranked by dollars at risk (DeniedCharges) rather than by claim counts. Build a combo chart and a scatter chart for the first question and a Pareto chart for the second, then check the numbers behind them.

Work on the **Bonus** sheet of the workbook.

- **B1.** On ED_Monthly, fill the yellow LWBSRate column with LWBS ÷ Total for each month. The column is already formatted as a percentage. The gray cell counts the months above the 2% target. How many of the 24 months missed the target? *(Hint: Write one formula in G2 with [@Column] references (Lesson 3.1). The Table fills the other 23 rows)*
- **B2.** Build a combo chart on ED_Monthly with Total as clustered columns and LWBSRate as a line on the secondary axis (select Month, then Ctrl+click or ⌘+click Total and LWBSRate). Which month had the highest LWBS rate? Type the month and year. *(Hint: Select A1:A25, Ctrl+click E1:E25 and G1:G25, then Insert → Insert Combo Chart)*
- **B3.** Test the director's theory with a scatter chart of Total (horizontal axis) against LWBSRate (vertical axis), with a linear trendline and its R². What is R²? Enter it to 3 decimal places. *(Hint: Select E1:E25, Ctrl+click G1:G25, then Insert → Scatter. RSQ(known_y's, known_x's) confirms it)*
- **B4.** On Denials, build a Pareto chart by DeniedCharges by hand: sort the table by DeniedCharges (largest first), fill CumulativePct with each row's running share of total DeniedCharges, then insert a combo chart with DeniedCharges as columns and CumulativePct as a line on the secondary axis (fix that axis at 0% to 100%). How many reasons does it take to reach at least 80% of denied charges? *(Hint: Follow 'A Pareto by hand' in Guide section 12. DeniedCharges is column C. Sort before you chart)*
- **B5.** Compare this Pareto with the one you built by claim count in Task 10. Exactly one reason ranks higher by denied charges than by number of claims. Which one? Type it as it appears in the table. *(Hint: Compare the order of the bars in the two Pareto charts)*
<!-- END GENERATED: bonus -->

<!-- BEGIN GENERATED: bonus-answers -->
<details>
<summary><b>🔑 Show the bonus solution</b> — Give it a real try first!</summary>

**B1. LWBSRate column (months above the 2% target)**

- **Answer:** 8
- **Solution:** `=[@LWBS]/[@Total]`

Type the formula in G2 (`=F2/E2` works too). [@LWBS] means "the LWBS value in this row", and the Table fills the column for you. Rates make months comparable even though volume swings by 228 visits between the busiest and quietest months. The gray cell runs `COUNTIF(G2:G25,">0.02")` on your column, and 8 of the 24 months were above 2%.

**B2. Combo chart: month with the highest LWBS rate**

- **Answer:** Sep 2024
- **Solution:**

1. Select **A1:A25**, then hold **Ctrl** (Mac: **⌘**) and select **E1:E25** and **G1:G25**.
2. Choose **Insert → Insert Combo Chart → Clustered Column – Line on Secondary Axis**.
3. Title both axes (*ED visits* and *LWBS rate*). If the secondary axis shows too many decimals, set its number format to a percentage with 1 decimal place (**Format Axis → Number**).
4. Hover over the highest point of the line.


Sep 2024 had 15 LWBS out of 403 visits (3.72%), yet it was the quietest month of the two years. The busiest month, Jan 2025 (631 visits), had an LWBS rate of only 1.74%. On the combo chart, the line's peaks don't line up with the tallest columns, which is the first hint that volume isn't the whole story. Cross-check with a formula: `=XLOOKUP(MAX(tblEDMonthly[LWBS]/tblEDMonthly[Total]),tblEDMonthly[LWBS]/tblEDMonthly[Total],tblEDMonthly[Month])`

**B3. Scatter: how much volume explains LWBS**

- **Answer:** 0.121
- **Solution:**

1. Select **E1:E25**, then **Ctrl+click** (Mac: **⌘+click**) **G1:G25**. The left column (Total) becomes X.
2. Choose **Insert → Insert Scatter (X, Y) or Bubble Chart → Scatter**.
3. Right-click a point → **Add Trendline…** → **Linear**, tick **Display R-squared value on chart**.
4. Confirm: `=RSQ(ED_Monthly!G2:G25,ED_Monthly!E2:E25)`.


R² = 0.121. The trendline slopes upward (about 0.43 percentage points of LWBS per 100 extra visits), so busier months do run slightly higher, but volume explains only about 12% of the month-to-month variation. The director should look at staffing, boarding, and triage flow as well. Twenty-four points is a small sample, so treat any pattern as a lead to investigate rather than proof. Cross-check with a formula: `=RSQ(tblEDMonthly[LWBS]/tblEDMonthly[Total],tblEDMonthly[Total])`

**B4. Manual Pareto: reasons needed to reach 80% of denied charges**

- **Answer:** 4
- **Solution:**

1. Click a DeniedCharges cell → **Data → Sort Largest to Smallest**.
2. In **D2** type `=SUM($C$2:C2)/SUM($C$2:$C$8)` and press **Enter**. The Table fills it down, and the column is already formatted as a percentage.
3. Select **A1:A8**, then **Ctrl+click** (Mac: **⌘+click**) **C1:D8**, and choose **Insert → Insert Combo Chart → Clustered Column – Line on Secondary Axis**.
4. Double-click the secondary axis and set **Minimum** `0` and **Maximum** `1`. Optionally set the column **Gap Width** to about 10%.
5. Find the first bar where the line reaches 80%.


Cumulative shares by dollars: Authorization Required 36.1%, Medical Necessity 59.4%, Coding Error 76.3%, Missing Documentation 86.1%, Eligibility / Coverage 95.1%, Timely Filing 97.9%, Duplicate Claim 100.0%. The line first passes 80% at reason 4, so 4 of the 7 reasons hold 86.1% of the \$13,906,332 at risk. The expanding range `$C$2:C2` (Lesson 3.3) has an absolute start and a relative end (Lesson 1.5), so its start stays fixed while its end moves down one row at a time. Building a Pareto by hand takes longer than the built-in chart, but it works in every Excel version and lets you add an 80% reference line or label the cut-off. Cross-check (counts the reasons whose running share is still below 80%, then adds one): `=SUMPRODUCT(--(SUMIF(tblDenials[DeniedCharges],">="&tblDenials[DeniedCharges])<0.8*SUM(tblDenials[DeniedCharges])))+1`

**B5. Which reason climbs when you rank by dollars**

- **Answer:** Missing Documentation
- **Solution:** Read the bar order in both charts. By claims, Missing Documentation is number 5, and by denied charges it's number 4, so it swaps places with Eligibility / Coverage.

Missing Documentation denials are fewer (97 claims against 125 for Eligibility / Coverage) but larger: about \$14,026 per claim against \$10,083. That's why the director asked for dollars. A count Pareto ranks the work queue by volume, and a dollar Pareto ranks it by money at risk. Put the two charts side by side, at the same size, so the committee sees the swap at a glance. Cross-check (COUNTIF(range,">"&range) gives each reason's rank minus 1, so MATCH finds the reason whose dollar rank beats its claim rank. In Excel 2019 or earlier, confirm it with Ctrl + Shift + Enter): `=INDEX(tblDenials[DenialReason],MATCH(1,--(COUNTIF(tblDenials[DeniedCharges],">"&tblDenials[DeniedCharges])<COUNTIF(tblDenials[Claims],">"&tblDenials[Claims])),0))`

</details>
<!-- END GENERATED: bonus-answers -->

## Key takeaways

- Pick the chart from the question: columns or bars to compare, lines for trends, a pie only for a few parts of a whole,
  histograms and box plots for distributions, and scatter charts for relationships.
- Bar and column axes start at zero. Bar charts draw the first row at the bottom, so reverse the categories for a ranking.
- A trendline is only as convincing as its R². SLOPE, INTERCEPT, and RSQ return the same numbers as a linear trendline.
- Combo charts with a secondary axis show two measures with different units, as long as both axes are titled.
- Pareto charts rank causes, waterfall charts explain how a total changed, and sparklines put a trend beside the numbers.
- Build a dynamic title in a cell with `&` and TEXT, then link the chart title to that cell.
- Declutter, use color to carry meaning with a color-blind-safe palette, and add alt text to every chart.

<!-- BEGIN GENERATED: nav -->
---

⬅️ **Previous:** [3.4 PivotTables](../04-pivottables/README.md) · 🏠 [Course home](../../README.md) · **Next:** [3.6 What-If Analysis: Goal Seek, Scenarios, Data Tables & Solver](../06-what-if-analysis/README.md) ➡️
<!-- END GENERATED: nav -->
