# Lesson 3.6 · What-If Analysis: Goal Seek, Scenarios, Data Tables & Solver

> **Level:** Intermediate · **Time:** about 60 minutes · **Workbook:** [`3.6-what-if-analysis.xlsx`](3.6-what-if-analysis.xlsx)
> **Data:** A monthly operating model for Bluestone's Primary Care Clinic (D400, Bluestone Outpatient Pavilion). Payer mix and reimbursement come from the clinic's 2,079 claims with 2025 service dates, wage rates from Bluestone HR records, and benefits, supply, and fixed costs from the clinic's 2025 budget actuals. The source files are [`claims.csv`](../../data/README.md#claimscsv), [`encounters.csv`](../../data/README.md#encounterscsv), [`employees.csv`](../../data/README.md#employeescsv), and [`budget.csv`](../../data/README.md#budgetcsv) in the data dictionary.

Bluestone's Primary Care Clinic loses a little money every month, like many hospital-owned primary care practices. The
questions land on the practice manager's desk in waves. *How many more patients a day would we need to break even? What if
more of our patients move to Medicaid? What rate should we ask for in the commercial contract renewal? What's the cheapest
nurse staffing mix that still covers the work safely?* Each answer means changing an assumption and watching the bottom line
move. Excel's **what-if analysis** tools do that systematically. Goal Seek works backward from a target, Scenario Manager
compares a few named stories, Data Tables show a whole range of answers at once, and Solver finds the best combination under
your rules. In this lesson you run all four on one realistic clinic model.

## What you'll learn

- Structure a model with separate inputs, calculations, and outputs
- Find break-even points with Goal Seek
- Compare cases with Scenario Manager and sensitivity Data Tables
- Optimize a staffing mix with Solver

## 📖 Guide

### 1. Four tools, four kinds of question

**What-if analysis** means changing the inputs of a model to see how its results respond. A **model** is a set of formulas
that turns assumptions into results. Its **inputs** are the assumptions you type in, such as visits per day or the RN hourly
rate. Its **outputs** are the results you care about, such as operating income.

Excel has four what-if tools, and each one answers a different kind of question:

| Tool | The question it answers | Inputs it changes | What you get | Example in this lesson |
|---|---|:-:|---|---|
| **Goal Seek** | What value of this input gives me exactly this result? | 1 | One number, written into the input cell | Visits per day needed to break even |
| **Scenario Manager** | What happens under each of these few named sets of assumptions? | Up to 32 | Saved scenarios you can switch between, and a summary report | Base plan vs. Downside vs. Upside |
| **Data Table** | How does a result change across a range of input values? | 1 or 2 | A live grid of results | Operating margin at 180, 185, … 220 visits per day |
| **Solver** | What's the best combination of inputs that follows my rules? | Up to 200 | The best values, plus optional reports | The cheapest RN/LPN/CNA mix that covers the workload |

Goal Seek, Scenario Manager, and Data Table sit under **Data → What-If Analysis**. Solver is a free add-in that ships with
Excel, and it appears on the **Data** tab once you turn it on (section 6).

> 📋 **Version note:** this lesson needs desktop Excel for Windows or Mac. Excel for the web has limited or no support for
> these tools, and its Solver is a separate add-in that works differently.

### 2. Build a model you can trust

Every what-if tool works the same way underneath. It changes one or more input cells, lets Excel recalculate, and reads an
output cell. So the tools only work as well as the model's structure. Open the **Model** sheet. It's the clinic's monthly
**profit and loss (P&L)** model: revenue at the top, then costs, then operating income. Look at how it's laid out:

| Section | Where | What's in it |
|---|---|---|
| **INPUTS** | Rows 5–39 | Volume and capacity, payer mix and $ per visit, staffing and pay, other costs. Every number is typed once, in blue. |
| **CALCULATIONS** | Rows 41–56 | One formula per row, flowing from visits per month down to total operating expenses |
| **OUTPUTS** | Rows 58–63 | Operating income and operating margin (you complete these), cost per visit, capacity |
| **SENSITIVITY AREA** | F5:K32 | Empty frames where you build two Data Tables |

The layout follows a handful of rules that analysts use for any model they expect to change:

| Rule | Why it matters for what-if analysis |
|---|---|
| Type every assumption once, in its own input cell | The what-if tools can only change cells. A number buried inside a formula is invisible to them. |
| Color-code cells: **blue** for inputs, **black** for formulas, **green** for links from another sheet | You, and anyone who inherits the model, can see at a glance which cells are safe to change |
| Write one formula per row, flowing top to bottom | Each step is easy to check, and every output traces back to the inputs |
| Never type over a formula | Overwriting a formula with a number freezes that part of the model without any warning |
| Name the inputs and outputs you'll target | Formulas read like sentences, and Scenario Manager's reports show the names instead of addresses |
| Add a check | Row 19 confirms the payer mix totals 100%. A model that checks itself catches typos before they spread. |

The Model names twelve cells, including **VisitsPerDay** (B8), **ClinicDays** (B7), **CommercialRate** (C17),
**NetRevenue** (B44), and **TotalExpenses** (B56). To see them all, choose **Formulas → Name Manager** (Windows: **Ctrl +
F3**). Lesson 3.1 covers naming in detail. Named cells make formulas self-explanatory. Compare these two versions of the
first calculation:

```
=B8*B7                         ← correct, but you have to look up both cells
=VisitsPerDay*ClinicDays       ← the same formula, readable at a glance
```

**Follow the money through the model.** The clinic plans 200 visits a day on 21 clinic days, so B42 holds 4,200 visits per
month. B43 averages the five payers' rates, weighted by their share of visits:

```
=SUMPRODUCT(B14:B18, C14:C18)    → 0.251×87 + 0.166×88 + 0.149×66 + 0.376×142 + 0.058×120 = $106.63 per visit
```

Net patient revenue is then 4,200 × $106.63 = $447,850 per month. Below that, wages, benefits, supplies, the billing fee,
and three fixed cost lines add up to total operating expenses. The **Sources** sheet shows where each data-derived input
came from. For example, the Commercial rate of $142 is the average allowed amount on the clinic's 2025 commercial and
workers' comp claims, and the RN rate of $42.73 is the median hourly rate of Bluestone's active RNs.

**How to find a hard-coded number.** A **hard-coded number** is a constant typed inside a formula, such as the 21 in
`=VisitsPerDay*21`. It gives the right answer today, so nothing looks wrong. But change ClinicDays to 22 and that formula
ignores you, along with every Data Table, scenario, and Goal Seek that depends on it. Three ways to hunt for one:

1. **Show Formulas.** Press **Ctrl + `` ` ``** (Mac: **⌃ + `` ` ``**), or choose **Formulas → Show Formulas**, to display
   every formula instead of its result. The `` ` `` is the grave accent key, left of 1. Read down the calculation rows and
   look for digits that aren't cell references. Press the shortcut again to switch back.
2. **Trace Dependents.** Select an input and choose **Formulas → Trace Dependents**. Blue arrows point to every formula that
   uses it, and a dashed arrow to a small sheet icon means a formula on another sheet uses it. If Excel draws no arrows and
   says no formula refers to the active cell, that input is an orphan, and some formula probably holds a typed-in copy of
   its value. **Formulas → Remove Arrows** clears the arrows.
3. **Go To Special.** Select the calculation cells (B42:B56) and choose **Home → Find & Select → Go To Special →
   Constants** (Windows: **F5 → Special…**). Excel selects every cell that holds a typed value instead of a formula. In a
   clean calculation section it finds nothing and says *No cells were found*.

> ⚠️ Go To Special finds cells that *are* constants, not constants hidden *inside* formulas. A formula like
> `=VisitsPerDay*21` passes the Go To Special test. Show Formulas and Trace Dependents catch it.

> 💡 **Tip:** Before you run any what-if tool, make sure the outputs are formulas and the inputs are values. Goal Seek
> refuses to target a cell without a formula. Goal Seek, Scenario Manager's **Show** button, and Solver all write numbers
> into the input cells they change, so a formula in one of those cells would be replaced.

### 3. Goal Seek: work backward from a target

**Goal Seek** finds the value of one input that makes a formula cell equal a target you choose. It's the tool for
break-even questions, where you know the answer you want and need the input that produces it. The **break-even point** is
the input value at which operating income is exactly $0, so the clinic neither makes nor loses money.

The Goal Seek dialog has three boxes:

| Box | What to enter | Rule | Example |
|---|---|---|---|
| **Set cell** | The output you want to hit | Must contain a formula | `B59` (operating income) |
| **To value** | The target | Must be a typed number. A cell reference isn't allowed. | `10000` |
| **By changing cell** | The one input Goal Seek may change | Must contain a value, not a formula | `B8` (visits per day) |

To run it:

1. Go to the sheet that holds the input, here the **Model** sheet. The **By changing cell** must be on the active
   sheet, and Excel rejects a reference to another sheet. Click the output cell (for example B59), so Excel pre-fills
   **Set cell** for you.
2. Choose **Data → What-If Analysis → Goal Seek** (Windows key tips: **Alt, A, W, G**).
3. Type the target in **To value**, click in **By changing cell**, click the input cell, and click **OK**.
4. The **Goal Seek Status** box reports whether it found a solution, and the input cell already shows that solution. Click
   **OK** to keep the new input value, or **Cancel** to put the original value back.

**Worked example: what volume earns $10,000 a month?** Click B59 (once you've completed it in task 1), open Goal Seek, and
set B59 to `10000` by changing `B8`. Goal Seek reports a solution, and B8 shows about **209.2** visits per day. Note the
number, then click **Cancel** so the Model returns to 200.

You can check that answer with a little algebra, because this model is **linear** in visits per day: every extra visit adds
the same amount of income. That amount is the **contribution margin** per visit, meaning what a visit brings in after the
costs that grow with volume:

```
Contribution margin per visit = $106.63 × (1 − 4% billing fee) − $4.39 supplies − $1.27 vaccines = $96.71
Fixed costs per month         = wages + benefits + facility + IT + other operating        = $414,783
Operating income              = visits per month × $96.71 − $414,783

Target $10,000:  visits per month = (414,783 + 10,000) ÷ 96.71 ≈ 4,392  →  ÷ 21 days ≈ 209.2 visits per day
```

Goal Seek got there by trial and error, so you didn't need the algebra. But the algebra tells you *why* the answer is 209.2,
and it reveals a problem Goal Seek can't see. The Model's capacity (B62) is 208 visits a day, so this target is out of reach
without another provider. **Goal Seek has no constraints.** It returns whatever number hits the target, even an impossible
one.

**How precise is Goal Seek?** Goal Seek tries a value, measures how far the set cell is from the target, and adjusts. It
stops when the result is close enough or after 100 tries. Both limits come from the iteration settings in **File →
Options → Formulas**: **Maximum Iterations** (default 100) and **Maximum Change** (default 0.001) (Mac: **Excel →
Preferences → Calculation**). For a linear model like this one, Goal Seek lands on the exact answer almost at once. For a
ratio such as operating margin, it can stop anywhere within about 0.001 of the target, which is 0.1 percentage point of
margin. Lower Maximum Change (for example, to 0.000001) when you need more precision. The two boxes are grayed out until
you tick **Enable iterative calculation**, so tick it, change the value, and then untick it again so Excel keeps warning
you about circular references.

| Goal Seek limitation | What to do instead |
|---|---|
| Changes only one input | Use Solver, which changes many |
| Ignores capacity, minimums, and other rules | Use Solver, which takes constraints |
| Finds one answer even when several exist (nonlinear models) | Start from a different input value and run it again |
| Clicking OK overwrites your base case | Click Cancel, or save a Base scenario first (section 4) |

> ⚠️ The Goal Seek Status box says *found a solution* whenever it gets close to the target. It doesn't mean the solution
> makes business sense. Always ask whether the new input value is achievable.

### 4. Scenario Manager: save and compare named cases

A **scenario** is a named set of values for a group of input cells. Those input cells are called the **changing cells**.
Planning teams rarely move one assumption at a time. A downside case usually combines lower volume, a worse payer mix, and
higher costs, and Scenario Manager stores each combination under a name so you can switch between them or compare them side
by side.

To add a scenario:

1. Select the changing cells. Hold **Ctrl** (Mac: **⌘**) and click to select cells that aren't next to each other, for
   example B8, B16:B17, and B36.
2. Choose **Data → What-If Analysis → Scenario Manager** (Windows key tips: **Alt, A, W, S**) and click **Add…**.
3. Type a **Scenario name**. The **Changing cells** box shows your selection. Add a comment if you like, then click **OK**.
4. The **Scenario Values** box lists each changing cell, by name when the cell has one, with its current value. Type the
   values for this scenario and click **OK**, or click **Add** to save it and start the next scenario straight away.

Type percentages as decimals in the Scenario Values box, for example `0.179` for 17.9%.

Back in Scenario Manager, select a scenario and click **Show** to write its values into the changing cells. The model
recalculates instantly. Click **Summary…** to build a report:

- **Scenario summary** inserts a new sheet named *Scenario Summary*. It has one column for the current values and one column
  per scenario. The top block lists the changing cells, and the bottom block lists the **result cells** you chose, such as
  B59 and B60.
- **Scenario PivotTable report** puts the same results in a PivotTable. It's useful when you have many scenarios.

A summary for this lesson's scenarios has this shape:

| | Current Values | Base plan | Downside | Upside |
|---|--:|--:|--:|--:|
| *Changing Cells:* | | | | |
| VisitsPerDay | 200 | 200 | 190 | 206 |
| MedicaidShare | 14.9% | 14.9% | 17.9% | 12.9% |
| … | | | | |
| *Result Cells:* | | | | |
| OperatingIncome | … | … | … | … |
| OperatingMargin | … | … | … | … |

Because the cells are named, the summary shows *VisitsPerDay* and *OperatingIncome* instead of `$B$8` and `$B$59`.

> 💡 **Tip:** Make your first scenario the **Base plan**, holding today's values. It's your reset button: after you show
> another scenario, or a Goal Seek or Solver run goes wrong, select Base plan and click **Show**.

> ⚠️ The Scenario Summary is a snapshot. It doesn't update when you edit the model, so create a fresh summary after any
> change. Also check that every changing cell holds a typed value. If a changing cell holds a formula, showing a scenario
> replaces that formula with a constant.

A few more things worth knowing:

- **A scenario changes only its own changing cells.** Every other input keeps the value it has when you click **Show**.
  Scenarios on one sheet can use different changing cells. For example, a staffing scenario that changes visits and FTEs
  leaves the payer mix alone. So if you last showed Downside, a staffing scenario shown next still runs on Downside's
  payer mix and fee. Show Base plan first whenever you need the other inputs back at the base case.
- **Edit…** changes a saved scenario's name, changing cells, or values, and **Delete** removes it. Fix a typo with
  Edit instead of adding the scenario again.
- Scenarios belong to the sheet they were created on. **Merge…** copies scenarios from another sheet or workbook.
- Each scenario can change up to 32 cells.
- Solver's results box has a **Save Scenario…** button, so you can store an optimal plan as a scenario and compare it with
  the others (section 6).

### 5. Data Tables: see a whole range of answers at once

A **Data Table** runs your model for a whole list of input values and lays the results out in a grid. Analysts often call
the result a **sensitivity table**, because it shows how sensitive an output is to an input.

> ⚠️ A Data Table has nothing to do with an Excel **Table** (Ctrl + T, Mac: Control + T) from Lesson 3.1. The names are
> confusingly similar. This section is only about the what-if tool.

**A one-variable Data Table** varies one input. The Model's Sensitivity area has the frame ready:

```
        F                  G                    H
 7   Visits per day     Operating income     Operating margin
 8   (leave empty)      =B59                 =B60              ← formula row: one formula per result you want
 9   180                ·                    ·
10   185                ·                    ·
 …    …                 ·                    ·
17   220                ·                    ·                 ← Excel fills the · cells
```

1. List the input values down a column (F9:F17 already holds 180 to 220 in steps of 5).
2. In the row just above the first value, starting one column to the right, enter formulas that point at the outputs
   you want: `=B59` in G8 and `=B60` in H8. You can add more formula columns.
3. Select the whole block, including the formula row and the input column (F8:H17).
4. Choose **Data → What-If Analysis → Data Table** (Windows key tips: **Alt, A, W, T**).
5. The input values run down a column, so click **Column input cell**, then click B8 (visits per day). Leave **Row input
   cell** empty. Click **OK**.

For each value in column F, Excel puts that value into B8, recalculates the model, and copies the results of G8 and H8 into
that row of the table. Then it puts the original value back in B8. Click any result cell and the formula bar shows
`{=TABLE(,B8)}`. The empty first argument means "no row input." The braces mean the whole result block is one array, so
you can't edit or delete one cell on its own.

**A two-variable Data Table** varies two inputs and shows one output. The formula goes in the top-left corner, one input's
values run across the top row, and the other input's values run down the left column:

```
        F                G        H        I        J        K
21   ↓ Visits per day   Commercial $ per visit →
22   =B59               130      140      150      160      170     ← corner formula, then row input values
23   180                ·        ·        ·        ·        ·
 …    …                 ·        ·        ·        ·        ·
31   220                ·        ·        ·        ·        ·
```

Select F22:K31, open **Data Table**, set **Row input cell** to C17 (the Commercial $ per visit) and **Column input cell** to
B8 (visits per day), and click **OK**. Each result cell shows `{=TABLE(C17,B8)}`. The first argument is always the row
input and the second is always the column input.

| Your input values run… | Put this input in… | Example |
|---|---|---|
| Across the top row | **Row input cell** | Commercial $ per visit (C17) |
| Down the left column | **Column input cell** | Visits per day (B8) |

> ⚠️ **Swapped input cells don't raise an error.** If you swap them, Excel fills the table with numbers that look fine and
> are wrong. Sanity-check one result against the Model. The base case is 200 visits at $142, so the cell at 200 visits and
> $140 should be a little below the Model's own operating income in B59, because commercial visits pay $2 less. With the
> inputs swapped, that cell would show 140 visits a day at $200, which is a much bigger loss.

> ⚠️ **The input cell must be on the same sheet as the Data Table.** Otherwise Excel says *Input cell reference is not
> valid*. That's why the Sensitivity area lives on the Model sheet. The table can show results from anywhere, but the input
> cells must be local.

More Data Table rules:

- **The formula cells show the base case.** After you build the tables, the formula row (G8:H8) and the corner cell (F22)
  still show the Model's current results. They aren't part of the grid. Some analysts hide the corner value with the
  custom number format `;;;` from Lesson 1.3, which hides a value without deleting the formula.
- **To delete a table**, select the entire result block (for example G9:H17) and press **Delete**. Editing a single cell
  triggers *Cannot change part of a data table*.
- **Type the axis values as constants.** If an axis value were a formula that depends on the input cell, the table would
  feed back into itself.
- **Read the results with ordinary formulas.** `=COUNTIF(G23:K31,">0")` counts the profitable combinations, and charts and
  conditional formatting work on the result block too.
- **Big tables slow everything down.** Excel recalculates every Data Table each time the workbook recalculates, even when
  the change has nothing to do with the table. In a large model, choose **Formulas → Calculation Options → Automatic
  Except for Data Tables** (Mac: also in **Excel → Preferences → Calculation**). Press **F9** (Mac: **⌘ + =**) whenever
  you want the tables refreshed.

> 💡 **Tip:** If a Data Table seems stuck showing old numbers, check the calculation option first. The calculation mode
> applies to every open workbook, and Excel takes it from the first workbook you open in a session, so it can be switched
> without you noticing.

**Data Table or Scenario Manager?**

| | Data Table | Scenario Manager |
|---|---|---|
| Inputs varied | 1 or 2 | Up to 32 per scenario |
| Values tested | An evenly spaced grid | A few coherent stories |
| Results | Live. They update when the model changes. | The summary is a static snapshot |
| Best for | Seeing the shape of a relationship, and where it crosses zero | Presenting a short list of named cases to leaders |

Data Tables and Goal Seek work well together. The one-variable table shows that the clinic crosses into profit somewhere
between two of its rows, and Goal Seek then pins down the exact crossing point.

### 6. Solver: find the best answer that follows your rules

**Solver** is an optimizer. You tell it which cell to make as small or as large as possible, which cells it may change, and
which rules the answer must obey. Solver then searches for the best values. It's the right tool when several inputs interact
and some combinations aren't allowed, as in staffing, scheduling, and supply ordering.

**Turn Solver on (once per computer):**

- **Windows:** **File → Options → Add-ins**. At the bottom, set **Manage** to *Excel Add-ins* and click **Go…**. Tick
  **Solver Add-in** and click **OK**. Solver appears at the right end of the **Data** tab.
- **Mac:** **Tools → Excel Add-ins…**, tick **Solver Add-In**, and click **OK**. Solver appears on the **Data** tab.

The lesson's Solver model lives on the **Staffing** sheet. It decides how many **FTEs** of each nursing role the clinic
staffs. An FTE (full-time equivalent) is one full-time position, so two half-time nurses make 1.0 FTE. The roles are
registered nurses (**RNs**) and licensed practical nurses (**LPNs**), who are licensed nurses, and certified nursing
assistants (**CNAs**), who aren't.

Every Solver model has three parts:

| Part | Meaning | On the Staffing sheet |
|---|---|---|
| **Objective cell** | The single formula cell to minimize, maximize, or set to a value | B31, monthly clinical support staff cost |
| **Variable cells** | The input cells Solver may change, also called decision variables | B18:B20, RN, LPN, and CNA FTEs (green) |
| **Constraints** | Rules the answer must satisfy, written as *cell, relation, limit* | Rows 25–28: hours covered, RN share, CNA cap |

**Worked example: translate the clinic's staffing rules.** Open the **Staffing** sheet. The clinic needs 0.60 support hours
per visit, and 0.28 of those hours must come from licensed nurses (RNs or LPNs). Policy says RNs make up at least 60% of
licensed staff. Each FTE works 168 paid hours a month. Here is how each rule becomes Solver language:

| The clinic's rule | Solver language |
|---|---|
| Spend as little as possible on clinical support staff | Objective `$B$31`, **To: Min** |
| Decide how many RN, LPN, and CNA FTEs to staff | Variable cells `$B$18:$B$20` |
| Cover every support hour: 168 × (RN + LPN + CNA) ≥ visits per month × 0.60 | `$B$25 >= $D$25` |
| Cover every licensed hour: 168 × (RN + LPN) ≥ visits per month × 0.28 | `$B$26 >= $D$26` |
| RNs are at least 60% of licensed FTEs | `$B$27 >= $D$27` (B27 holds RN − 0.6 × (RN + LPN), and D27 holds 0) |
| No negative staff | Tick **Make Unconstrained Variables Non-Negative** |
| *(Bonus)* At most 6 CNAs | `$B$28 <= $D$28` |
| *(Bonus)* Hire whole people | `$B$18:$B$20` **int** |

Look closely at the RN rule. The natural way to write it is RN ÷ (RN + LPN) ≥ 60%, but that formula divides by variable
cells, which makes it **nonlinear**. Multiplying both sides by (RN + LPN) gives RN ≥ 0.6 × (RN + LPN), which moves to
RN − 0.6 × (RN + LPN) ≥ 0. Now every variable is only multiplied by a constant and added. That makes the constraint
**linear**, and the fastest and most reliable Solver method requires it.

**Set up and run Solver:**

1. Go to the **Staffing** sheet and choose **Data → Solver**. The objective and variable cells must be on the active
   sheet.
2. **Set Objective:** `$B$31`. Choose **Min**. (**Max** maximizes. **Value Of** sets the objective to a number, the way Goal
   Seek does.)
3. **By Changing Variable Cells:** `$B$18:$B$20`.
4. Click **Add** for each constraint. Enter the **Cell Reference**, pick the relation (`<=`, `=`, `>=`, `int`, `bin`, or
   `dif`), and enter the **Constraint**. Constraints with the same relation can share one line: `$B$25:$B$27 >= $D$25:$D$27`.
5. Leave **Make Unconstrained Variables Non-Negative** ticked.
6. **Select a Solving Method:** *Simplex LP*.
7. Click **Solve**.

| Solving method | Use it when | Example |
|---|---|---|
| **Simplex LP** | The objective and all constraints are linear (sums of variables times constants) | Staffing cost, supply orders, bed allocation |
| **GRG Nonlinear** | The model is smooth but nonlinear: ratios, or variables multiplied together | Choosing visits and rates to maximize margin |
| **Evolutionary** | The model uses IF, lookups, ROUND, or other jumps | Rosters that use lookup tables |

Simplex LP is exact and fast, and its Sensitivity Report shows the shadow prices and allowable ranges described in section
7. (GRG Nonlinear's report shows reduced gradients and Lagrange multipliers instead, without the ranges.) If you pick
Simplex LP for a nonlinear model, Solver stops with *The linearity conditions required by this LP Solver are not
satisfied*.

**The Solver Results box** appears when Solver finishes:

- **Keep Solver Solution** leaves the new values in the variable cells. **Restore Original Values** puts the old ones back.
- **Reports** lists *Answer*, *Sensitivity*, and *Limits*. Select one or more (Ctrl-click, or ⌘-click on a Mac) before you
  click **OK**, and Excel adds each report as a new sheet.
- **Save Scenario…** stores the solution as a scenario (section 4).

**Solver Options worth knowing** (click **Options** in the Solver Parameters box, then the **All Methods** tab):

| Option | Default | When to change it |
|---|---|---|
| **Integer Optimality (%)** | 1 | Set it to 0 for integer models. At 1, Solver may stop at any plan within 1% of the best possible one. |
| **Ignore Integer Constraints** | Off | Turn it on to solve the fractional version of an integer model quickly |

Solver saves its settings with the worksheet, so the next time you open Solver on that sheet your objective, variables, and
constraints are still there. **Reset All** clears them.

> ⚠️ **Common Solver messages.** *Solver could not find a feasible solution* means no combination satisfies every
> constraint, so look for a constraint with the wrong sign. *The Objective Cell values do not converge* means the objective
> can keep improving forever, which usually means a rule is missing. For example, maximizing visits with no capacity
> constraint has no best answer.

> ⚠️ You can't count on **Ctrl + Z** to undo a Solver run. If you only want the reports, choose **Restore Original
> Values** in the Solver Results box.

**Integer variables.** Staffing whole people is an **integer constraint**. Solver handles it with *branch and bound*: it
solves the fractional problem, then repeatedly splits it, for example into "RN ≤ 6" and "RN ≥ 7", until it proves which whole
numbers are best. Two consequences matter in practice:

- **Rounding the fractional answer isn't the same as solving for whole numbers.** Rounding each variable on its own can break
  a constraint or cost more than necessary. The bonus shows a case where it does both.
- Excel doesn't offer the Sensitivity or Limits reports for models with integer constraints. To get shadow prices, solve
  the fractional version (tick **Ignore Integer Constraints**).

### 7. Read the Solver reports

Reports turn Solver's answer into an explanation that you can bring to a staffing meeting.

**The Answer Report** lists the objective and variable cells before and after the run. Then it lists every constraint with
two key columns:

| Column | Meaning |
|---|---|
| **Status** | **Binding** means the constraint is exactly at its limit and is holding the answer back. **Not Binding** means there's room to spare. |
| **Slack** | How far the constraint is from its limit. Binding constraints have a slack of 0. |

**The Sensitivity Report** (Simplex LP, no integer constraints) answers "what if a limit or a cost changed a little?"

| Section | Column | Meaning |
|---|---|---|
| Variable Cells | **Reduced Cost** | For a variable at 0: how much its cost would have to fall before Solver used any of it. It's 0 for variables already in the plan. |
| Variable Cells | **Objective Coefficient**, **Allowable Increase / Decrease** | The variable's cost per unit, and how far that cost can move before the best plan changes |
| Constraints | **Shadow Price** | How much the objective changes if the constraint's limit (its right-hand side) rises by 1 unit |
| Constraints | **Constraint R.H. Side**, **Allowable Increase / Decrease** | The limit, and the range over which that shadow price stays valid |

A **shadow price** puts a dollar value on a rule. Non-binding constraints have a shadow price of 0, because loosening a rule
that isn't holding you back saves nothing.

**Worked example: what does the 60% RN rule cost?** The Staffing sheet's column D shows the loaded cost of each FTE per month,
meaning wages plus 26.7% benefits for 168 hours: $9,095.34 for an RN, $5,902.50 for an LPN, and $4,201.78 for a CNA. When the
RN rule is binding, raising its right-hand side from 0 to 1 forces Solver to swap one LPN FTE for one RN FTE while keeping
the licensed hours the same. So the rule's shadow price is

```
$9,095.34 − $5,902.50 = $3,192.84 per month for each FTE moved from LPN to RN
```

That's the kind of number a nurse manager can take to leadership: a stricter RN policy costs about $3,200 a month for each
LPN FTE it replaces. Task 13 asks you to read and explain the shadow price of a different constraint.

**The Limits Report** shows how far each variable can move on its own while the other variables stay fixed and the
constraints still hold. It's used less often than the other two.

### 8. Putting it together

A practical order of work for any what-if question:

1. **Build the model** with separate inputs, calculations, and outputs, and add checks.
2. **Save a Base scenario** so you can always get back to the starting point.
3. **Run a Data Table** on the most important input to see the shape of the relationship.
4. **Use Goal Seek** to find the exact point where the result crosses your target.
5. **Use Scenario Manager** to present a handful of realistic combined cases.
6. **Use Solver** when several inputs interact and the answer must follow rules.

| Action | Windows | Mac |
|---|---|---|
| Show or hide formulas | Ctrl + `` ` `` | ⌃ + `` ` `` |
| Goal Seek | Alt, A, W, G | Data → What-If Analysis → Goal Seek |
| Scenario Manager | Alt, A, W, S | Data → What-If Analysis → Scenario Manager |
| Data Table | Alt, A, W, T | Data → What-If Analysis → Data Table |
| Solver | Data → Solver | Data → Solver |
| Recalculate (with *Automatic Except for Data Tables*) | F9 | ⌘ + = |
| Select cells that aren't next to each other | Ctrl + click | ⌘ + click |
| Name Manager | Ctrl + F3 | Formulas → Name Manager |

> 📋 **Version notes:** Goal Seek, Scenario Manager, and Data Tables are in every current desktop version of Excel for
> Windows and Mac. The Solver with the Simplex LP, GRG Nonlinear, and Evolutionary methods is included with Excel 2010 and
> later for Windows and Excel 2016 and later for Mac. Scenarios and Solver settings are saved inside the workbook, so a
> regular `.xlsx` file keeps them, and no macros are needed.

## 🧪 Hands-on practice

Download [`3.6-what-if-analysis.xlsx`](3.6-what-if-analysis.xlsx) and open the **Practice** sheet. Work through the tasks in
order, because later tasks build on the outputs you complete and the fix you make in the first three. Type each answer in
the yellow cell, and the **Check** column turns green when you're right.

<!-- BEGIN GENERATED: practice -->
Every task uses the Model and Staffing sheets. Tasks 1 and 2 read your Model live, so their checks stay green only while the Model holds its base-case inputs (200 visits per day, Commercial $142, and so on). The same is true of any answer you type as a link to a Model cell, such as =Model!H11. After each Goal Seek run, click Cancel in the Goal Seek Status box so the Model keeps those inputs. The Staffing sheet doesn't feed the Model, so you can keep Solver's solutions there.

| # | Task | Hint |
|:-:|------|------|
| 1 | On the Model sheet, complete the yellow Operating income cell (B59): net patient revenue minus total operating expenses. The gray cell here reads your formula. What is the clinic's base-case operating income per month? | An output should only point at calculation cells |
| 2 | Complete the yellow Operating margin cell (B60): operating income as a share of net patient revenue. What is the base-case operating margin? The cell is already formatted as a percentage. | Margin = income ÷ revenue |
| 3 | Audit the model. One formula in the CALCULATIONS section (rows 42–56) has a number typed into it instead of a reference to its input cell. Type that cell's address (for example B99). Then fix the formula so it points to the input. Task 10 depends on the fix. | Show Formulas, or Trace Dependents on each input |
| 4 | On the Model sheet, use Goal Seek to find the break-even volume: set Operating income (B59) to 0 by changing Visits per day (B8). How many visits per day does the clinic need? Round to 1 decimal place, then click Cancel to restore 200. | Data → What-If Analysis → Goal Seek |
| 5 | The CFO is renegotiating commercial contracts. On the Model sheet, use Goal Seek to set Operating margin (B60) to 5% (type 0.05) by changing the Commercial $ per visit (C17). What commercial reimbursement per visit is needed? Round to the nearest dollar, then click Cancel to restore $142. | The Set cell must contain a formula, so use the margin cell |
| 6 | Build the one-variable Data Table in Model!F8:H17: put =B59 in G8 and =B60 in H8, select F8:H17, and use Column input cell B8. What operating margin does your table show at 190 visits per day? Enter it as a percentage to 1 decimal place. | The visits run down a column, so use the Column input cell |
| 7 | Look down the Operating income column of your one-variable table. What is the lowest visits-per-day value in the table at which the clinic makes a profit (operating income above 0)? | Read the table, or let MINIFS find it |
| 8 | Build the two-variable Data Table: put =B59 in the corner cell F22, select F22:K31, and use Row input cell C17 (commercial $ across row 22) and Column input cell B8 (visits per day down column F). What operating income does the table show at 210 visits per day and $160 per commercial visit? Round to the nearest dollar. | Row input = the input whose values run across the top row |
| 9 | How many of the 45 combinations in your two-variable table (G23:K31) are profitable (operating income above 0)? Use a formula. | COUNTIF with ">0" |
| 10 | Open Scenario Manager on the Model and add three scenarios that change B8, B16, B17 and B36. Base plan: visits per day 200, Medicaid share 14.9%, Commercial share 37.6%, billing fee 4.0%. Downside: visits per day 190, Medicaid share 17.9%, Commercial share 34.6%, billing fee 5.0%. Upside: visits per day 206, Medicaid share 12.9%, Commercial share 39.6%, billing fee 3.5%. Create a Scenario Summary with result cells B59 and B60. What is operating income in the Downside scenario? Round to the nearest dollar. | Data → What-If Analysis → Scenario Manager |
| 11 | From the same Scenario Summary, what operating margin does the Upside scenario produce? Enter it as a percentage to 1 decimal place. | Same summary, different column |
| 12 | On the Staffing sheet, use Solver to minimize the monthly staff cost (B31) by changing the RN, LPN, and CNA FTEs (B18:B20), subject to the three constraints in rows 25–27 (leave out the CNA cap in row 28). Keep Make Unconstrained Variables Non-Negative ticked and choose Simplex LP. What is the minimum monthly cost? Round to the nearest dollar, and type the number rather than a link to B31, because the bonus runs Solver on this sheet again. | Data → Solver (enable the Solver add-in first) |
| 13 | Run Solver again (same setup) and select Sensitivity under Reports before you click OK. On the Sensitivity Report sheet, what is the Shadow Price of the Total support hours constraint (Staffing!B25)? Enter it in dollars to 2 decimal places. | Shadow price = cost change per one-unit increase in the constraint's right-hand side |
<!-- END GENERATED: practice -->

## ✅ Answer key

The workbook has hidden **Answer Key** and **Bonus Key** sheets (right-click any sheet tab → **Unhide…**). For the Goal Seek
and Solver tasks, the key's *Live result* column checks the answer with algebra like the worked examples in sections 3 and
7. The same answers are below, collapsed so you don't see them by accident.

<!-- BEGIN GENERATED: answers -->
<details>
<summary><b>🔑 Show the answer key</b> — Try every task before opening this.</summary>

**1. Operating income (Model!B59)**

- **Answer:** -8,618.72
- **Solution:** `=B44-B56`

Operating income is revenue minus expenses, and both already exist as calculation rows (B44 and B56). Because those cells are named, `=NetRevenue-TotalExpenses` works too and reads better. The clinic loses about $8,619 a month in the base case. That's common for hospital-owned primary care, which is why the rest of the lesson asks what it would take to break even.

**2. Operating margin (Model!B60)**

- **Answer:** -1.9%
- **Solution:** `=B59/B44`

Margin divides operating income by revenue (`=OperatingIncome/NetRevenue`). A margin lets you compare a small clinic with a large hospital, which a dollar figure can't. Goal Seek will aim at this cell in task 5.

**3. Find the hard-coded number**

- **Answer:** B52
- **Solution:**

1. On the Model sheet, choose **Formulas → Show Formulas** (or press the Show Formulas shortcut from section 2), and read down rows 42–56.
2. B52 reads `=NetRevenue*0.04`. The 4% is typed in, although the fee has its own input cell, B36.
3. Change B52 to `=NetRevenue*BillingFeePct` (or `=B44*B36`), then turn **Show Formulas** off again to see values.


Today both versions return the same number, so nothing looks wrong. The trouble starts when someone changes the fee input in B36: the model ignores it. Another way to catch this is **Formulas → Trace Dependents** on B36. Excel draws no arrows and tells you no formula refers to the active cell, because no formula reads that input.

**4. Goal Seek: break-even visits per day**

- **Answer:** 204.2
- **Solution:**

1. On the Model sheet, choose **Data → What-If Analysis → Goal Seek**.
2. **Set cell:** `B59` · **To value:** `0` · **By changing cell:** `B8`.
3. Click **OK**. B8 shows 204.2 (the cell holds 204.2440…). Note it, then click **Cancel**.


Each extra visit per day adds one visit on each of the 21 clinic days. Each of those visits brings in $96.71 after the billing fee, supplies, and vaccines (its **contribution margin**). Fixed costs are $414,783 a month, so break-even = $414,783 ÷ ($96.71 × 21) ≈ 204.2. That is 98% of the clinic's 208-visit daily capacity, so volume alone is a fragile fix. The live result in the key does this algebra, and Goal Seek reaches the same answer by trial and error.

**5. Goal Seek: commercial rate for a 5% margin**

- **Answer:** about $164 (Goal Seek shows 163.58)
- **Solution:**

1. On the Model sheet, choose **Data → What-If Analysis → Goal Seek**.
2. **Set cell:** `B60` · **To value:** `0.05` · **By changing cell:** `C17`.
3. Click **OK**, read C17, and click **Cancel**.


Commercial plans would have to pay about $164 instead of $142, an increase of 15%. Only 37.6% of visits are commercial, so each extra commercial dollar moves the average reimbursement by just 37.6 cents. Margin is a ratio, so Goal Seek has to iterate here, and it stops when the margin is within its tolerance of 5%. That's why the check accepts a small range around the exact value.

**6. One-variable Data Table: margin at 190 visits/day**

- **Answer:** -6.8%
- **Solution:** `=Model!H11`

The visits per day values run **down a column** (F9:F17), so B8 is the **column** input cell and the row input box stays empty. For each value, Excel puts it in B8, recalculates the model, and writes the results of the formulas in row 8 into that row of the table. You can type the number or point at H11.

**7. One-variable Data Table: first profitable volume**

- **Answer:** 205
- **Solution:** `=MINIFS(Model!F9:F17,Model!G9:G17,">0")`

The table jumps in steps of 5, so it brackets the break-even point instead of finding it: 200 visits loses money and 205 makes money, which agrees with Goal Seek's 204.2. Use a Data Table to see the whole curve, and Goal Seek to pin down the exact crossing. MINIFS needs Excel 2019 or later. In older versions, read the value off the table and type it.

**8. Two-variable Data Table: 210 visits/day × $160**

- **Answer:** 40,342
- **Solution:** `=Model!J29`

A two-variable table has exactly one formula, in its top-left corner. Excel substitutes each top-row value into the **row** input cell (C17) and each left-column value into the **column** input cell (B8), and fills every intersection. If you swap the two input cells, the table still fills without any warning, but with wrong numbers. So check one cell by hand. The base case is 200 visits at $142 (−$8,619), so the cell at 200 visits and $140 should be about $3,032 lower, because the clinic's 1,579 commercial visits a month each pay $2 less, minus the 4% fee on those dollars. With swapped input cells that cell would show 140 visits a day at $200, a far bigger loss.

**9. Two-variable Data Table: profitable combinations**

- **Answer:** 24
- **Solution:** `=COUNTIF(Model!G23:K31,">0")`

Ordinary formulas can read a Data Table's results, so you can count them, chart them, or add conditional formatting. The profitable combinations sit in the bottom-right of the table, where volume and rate are both high. Reading down each column shows the trade-off the CFO cares about: at $130 per commercial visit the clinic first makes money at 215 visits a day, but at $170 it does at 185. A better contract lowers the volume the clinic needs to break even.

**10. Scenario Manager: Downside operating income**

- **Answer:** -41,824
- **Solution:**

1. Select B8, then Ctrl-click (Mac: ⌘-click) B16:B17 and B36.
2. **Data → What-If Analysis → Scenario Manager → Add…**. Name it *Base plan* and click **OK**. The values box shows the current values, so click **OK** again.
3. Click **Add…** again for *Downside* and *Upside*. Type the values in the order B8, B16, B17, B36 (decimals like 0.179 for 17.9%).
4. Click **Summary…**, choose **Scenario summary**, set **Result cells** to `B59,B60`, and click **OK**. Read the Downside column on the new Scenario Summary sheet.


Fewer visits, a shift from commercial to Medicaid, and a higher vendor fee all hit income at once. If you skipped the fix in task 3, the model ignores the 5% fee and you'd see −$37,660 instead. That's exactly the kind of silent error a hard-coded number causes. The summary sheet is a snapshot. It doesn't update when the model changes, so create it again after any edit.

**11. Scenario Manager: Upside margin**

- **Answer:** 2.6%
- **Solution:** Read the Upside column of the OperatingMargin row on the Scenario Summary sheet.

The Upside case brings the clinic to roughly 2.6%. It takes three things at once: 6 more visits a day, 2 points of payer mix shifted from Medicaid to Commercial, and a cheaper vendor contract (3.5%). Because B59 and B60 are named, the summary labels those rows OperatingIncome and OperatingMargin instead of $B$59 and $B$60.

**12. Solver: lowest-cost staffing mix (FTEs)**

- **Answer:** 88,342
- **Solution:**

1. On the Staffing sheet, choose **Data → Solver**.
2. **Set Objective:** `$B$31` · **To:** Min · **By Changing Variable Cells:** `$B$18:$B$20`.
3. Click **Add** and enter `$B$25:$B$27` **>=** `$D$25:$D$27`, then click **OK**.
4. Leave **Make Unconstrained Variables Non-Negative** ticked, choose **Simplex LP**, and click **Solve**.
5. Choose **Keep Solver Solution**, click **OK**, and type the value of B31, rounded to the dollar, in the answer cell.


Solver chooses 4.2 RN, 2.8 LPN and 8.0 CNA FTEs. This answer makes sense: CNAs are the cheapest staff, so they cover every hour that doesn't need a licensed nurse. The licensed hours are then split at exactly the 60% RN minimum, because LPNs cost less than RNs. Today's 5/3/8 staffing costs $96,798, so the plan saves about $8,457 a month. The key's live result rebuilds this optimum with algebra from the Staffing sheet's own cells.

**13. Solver Sensitivity Report: shadow price**

- **Answer:** 25.01
- **Solution:**

1. **Data → Solver**, same parameters as task 12, **Solve**.
2. In **Solver Results**, select **Sensitivity** under *Reports*, then click **OK**.
3. On the new *Sensitivity Report 1* sheet, find the row for `$B$25` under *Constraints* and read **Shadow Price**.


Requiring one more support hour raises the minimum cost by $25.01. That's one CNA hour at $19.74 plus 26.7% benefits, because the cheapest way to cover an hour that doesn't need a licensed nurse is a CNA hour. The live result in the key repeats that arithmetic. The licensed-hours constraint (B26) has a shadow price of only $21.53, even though licensed nurses cost more. An extra licensed hour costs 0.6 RN hours plus 0.4 LPN hours, but licensed hours also count toward the total, so it replaces a CNA hour that's no longer needed.

</details>
<!-- END GENERATED: answers -->

## 🏆 Bonus challenge

<!-- BEGIN GENERATED: bonus -->
Bluestone wants the Primary Care Clinic to grow to 220 visits per day in 2026 by adding a fourth advanced practice provider (APP: a nurse practitioner or physician assistant). Capacity rises from 208 to 226 visits per day. HR hires whole people, so every role needs a whole number of 1.0-FTE staff. Only 6 CNA positions are approved (Staffing!B12). Find the cheapest staffing plan, then test the full plan in the Model's P&L.

Work on the **Bonus** sheet of the workbook.

- **B1.** On the Staffing sheet, change B6 to 220. Re-run Solver with two more constraints: B18:B20 = int (integer) and B28 <= D28 (the CNA cap). In Solver Options, set Integer Optimality (%) to 0. What is the minimum monthly staff cost? Round to the nearest dollar. *(Hint: In Add Constraint, pick int from the middle list. Integer Optimality is on the All Methods tab of Options)*
- **B2.** How many RNs does that plan hire? *(Hint: Read Staffing!B18. The RN rule in row 27 should still say OK)*
- **B3.** How many LPNs does that plan hire? *(Hint: Is rounding the fractional answer up the same thing?)*
- **B4.** Test the plan in the P&L. On the Model, first select the Base plan scenario and click Show, so the payer mix and billing fee are back at their base values. Then add two more scenarios that change B8, B10, and B21:B23 (visits per day, APP FTEs, and RN, LPN, and CNA FTEs, in that order): Today (200, 3, 5, 3, 8) and Growth 2026 (220, 4, and the RN, LPN, and CNA counts in Staffing!B18:B20). What operating income does Growth 2026 produce? Round to the nearest dollar. *(Hint: A scenario changes only its own cells. Every other input keeps the value it has when you click Show)*
<!-- END GENERATED: bonus -->

<!-- BEGIN GENERATED: bonus-answers -->
<details>
<summary><b>🔑 Show the bonus solution</b> — Give it a real try first!</summary>

**B1. Integer staffing plan: minimum cost**

- **Answer:** 112,488
- **Solution:**

1. Type 220 in Staffing!B6.
2. **Data → Solver**. Keep the objective, variable cells, and the three constraints from task 12.
3. **Add** `$B$18:$B$20` **int** and **Add** `$B$28 <= $D$28`.
4. **Options** → *All Methods* tab → **Integer Optimality (%)** = 0 → **OK**. Check that **Ignore Integer Constraints** is not ticked.
5. **Solve**, keep the solution, and read B31.


With fractions allowed, the cheapest plan costs $107,302 (6.3 RN, 4.2 LPN, 6.0 CNA). Requiring whole people raises that to $112,488, so hiring whole people costs $5,186 a month more. Solver handles integer constraints by **branch and bound**, solving many LPs with tighter and tighter bounds. Its default Integer Optimality of 1% lets it stop at any plan within 1% of the best possible, so set it to 0 when the exact answer matters.

**B2. Integer plan: RNs**

- **Answer:** 7
- **Solution:** Read Staffing!B18 after Solver finishes.

7 RNs out of 11 licensed staff is 63.6%, which meets the 60% rule.

**B3. Integer plan: LPNs**

- **Answer:** 4
- **Solution:** Read Staffing!B19 after Solver finishes.

The fractional plan needs 10.5 licensed FTEs, and 11 whole people already cover that. A tempting shortcut is to round each fractional value up, which gives 7 RN, 5 LPN and 6 CNA. That plan hires one licensed person too many and costs $5,902 a month more. It also breaks the RN rule, because 7 of 12 licensed staff is only 58.3%. Rounding each variable separately can't see how the constraints interact, but Solver's integer search can.

**B4. Growth 2026 scenario: operating income**

- **Answer:** 2,624
- **Solution:**

1. On the Model, open **Scenario Manager**, select *Base plan*, and click **Show**. Close the dialog.
2. Select B8, B10, and B21:B23 (Ctrl-click, or ⌘-click on a Mac).
3. **Scenario Manager → Add…** *Today*: 200, 3, 5, 3, 8.
4. **Add…** *Growth 2026*: 220, 4, 7, 4, 6.
5. **Summary…** with result cells `B59,B60`, or select Growth 2026, click **Show**, read B59, then **Show** Today to restore the base case.


The growth plan turns a −$8,619 monthly loss into $2,624 (0.5% margin). The new APP costs $10,800 plus benefits, and the larger staff costs $15,690 more a month than today's. The extra 420 visits a month each contribute $96.71, which covers both. Growth 2026 changes only its five cells, so it inherits the payer mix and billing fee already on the Model. If Upside was the last scenario you showed, Growth 2026 would quietly use Upside's mix and fee, which is why you show Base plan first. Using scenarios instead of overwriting inputs also means one click (Show Today) puts the base case back for every other task.

</details>
<!-- END GENERATED: bonus-answers -->

## Key takeaways

- Keep inputs, calculations, and outputs separate, and type every assumption once. A number hidden inside a formula
  silently breaks every what-if tool.
- **Goal Seek** answers "what input hits this target?" for one input. Check its answer for plausibility, because it ignores
  capacity and every other rule.
- **Scenario Manager** stores named combinations of inputs. Save a Base plan first, and remember that the summary is a
  snapshot.
- **Data Tables** show a live grid of results for one or two inputs. The input cells must be on the table's sheet, and the
  row and column inputs are easy to swap.
- **Solver** optimizes an objective subject to constraints. Write constraints in linear form, use Simplex LP, and set
  Integer Optimality to 0 when you need whole numbers.
- A shadow price turns a rule into dollars. It tells you what one more unit of a requirement costs.

<!-- BEGIN GENERATED: nav -->
---

⬅️ **Previous:** [3.5 Charts & Data Visualization](../05-charts-visualization/README.md) · 🏠 [Course home](../../README.md) · **Next:** [4.1 Dynamic Arrays: FILTER, SORT, UNIQUE & More](../../04-advanced-analysis/01-dynamic-arrays/README.md) ➡️
<!-- END GENERATED: nav -->
