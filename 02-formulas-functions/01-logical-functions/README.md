# Lesson 2.1 · Logical Functions: IF, AND, OR, IFS & More

> **Level:** Beginner → Intermediate · **Time:** about 50 minutes · **Workbook:** [`2.1-logical-functions.xlsx`](2.1-logical-functions.xlsx)
> **Data:** 391 lab results collected December 1–4, 2025 across Bluestone Health, and all 423 emergency department visits at Bluestone Memorial Hospital in December 2025 with triage vital signs. Triage blood pressure was not recorded for 12 of those visits.

Hospitals run on rules. A glucose above 99 mg/dL gets flagged High. Two abnormal vital signs trigger a sepsis screen. A walk-in
with a sprained ankle goes to fast track instead of a main ED bed. Every one of those rules is a yes-or-no question, and Excel
answers yes-or-no questions with **logical functions**. In this lesson you'll turn rules like these into formulas that flag a
whole column of lab results or ED visits at once. Then a question like *"How many patients met the criteria last month?"*
takes seconds instead of a chart review.

## What you'll learn

- Compare values to produce TRUE/FALSE and use booleans in math
- Make decisions with IF, nested IF, and IFS
- Combine conditions with AND, OR, NOT (and XOR)
- Map codes to labels with SWITCH; trap errors with IFERROR and IFNA

## 📖 Guide

The examples use the lesson workbook. On the **ED** sheet, column C is ArrivalMode, D is ESILevel, E is ChiefComplaint, F is
TempF, G is HeartRate, H is RespRate, I is SystolicBP, J is SpO2, K is PainScore, and L is EDDisposition. On the **Labs** sheet,
E is ResultValue, G is RefLow, H is RefHigh, and I is Priority. Open the workbook and try each example as you read. Type it
on the same sheet as the data it uses, in an empty cell at least one column away from the Table (for example, column T on the
ED sheet or column L on the Labs sheet). A cell right next to a Table becomes part of the Table as soon as you type in it.

### 1. Comparisons return TRUE or FALSE

A **comparison** asks Excel a yes-or-no question about two values. The answer is always one of two **logical values**: `TRUE`
or `FALSE`. (Programmers call them *booleans*.) You don't need a function. A comparison operator between two values is a complete
formula.

The table uses ED row 5, visit ED211701: a walk-in at ESI level 4 with a temperature of 100.7 °F, heart rate 109, respiratory
rate 15, and SpO2 96%, who was admitted.

| Operator | Meaning | Example on ED row 5 | Result |
|:-:|---|---|---|
| `=` | equal to | `=C5="Walk-In"` | TRUE |
| `<>` | not equal to | `=L5<>"Discharged"` | TRUE (the patient was admitted) |
| `>` | greater than | `=F5>100.4` | TRUE (100.7 °F) |
| `<` | less than | `=J5<92` | FALSE (SpO2 is 96) |
| `>=` | greater than or equal to | `=D5>=4` | TRUE (ESI 4) |
| `<=` | less than or equal to | `=H5<=20` | TRUE (15 breaths per minute) |

Four rules cover almost every comparison you'll write:

- **Equal is not greater.** Row 9 (ED211707) has a heart rate of exactly 90. `=G9>90` returns FALSE, but `=G9>=90` returns TRUE.
  Clinical definitions are precise about this ("above 90" versus "90 or more"), so read the rule and pick the matching operator.
- **Text goes in double quotes, and numbers don't.** Write `=C5="Walk-In"` and `=F5>100.4`.
- **Text comparisons ignore case.** `=C5="walk-in"` is also TRUE. For a case-sensitive test, use EXACT (Lesson 2.2).
- **Dates are numbers**, so `=B5>=DATE(2025,12,15)` asks "did this visit arrive on or after December 15?" Lesson 2.3 covers dates.

> ⚠️ **Don't put quotes around numbers.** `=F5>"100.4"` compares a temperature with the *text* "100.4". Excel ranks every number
> below every piece of text, so this formula returns FALSE for every temperature. Write `=F5>100.4`.

> ⚠️ **Excel can't read math-class chains.** `=90<G5<130` looks like "between 90 and 130," but Excel works left to right. First
> `90<G5` becomes TRUE. Then Excel evaluates `TRUE<130`, and logical values rank above numbers, so the result is FALSE no matter
> what the heart rate is. Write `=AND(G5>90,G5<130)` instead (section 6).

### 2. TRUE and FALSE are numbers in disguise

When you do arithmetic on a logical value, Excel treats TRUE as 1 and FALSE as 0. That one fact lets you count and score with plain
math.

| Formula | Result | Why |
|---|---|---|
| `=TRUE+TRUE` | 2 | 1 + 1 |
| `=(F5>100.4)+(G5>90)` | 2 | Row 5 has a fever **and** a heart rate above 90 |
| `=(F9>100.4)+(G9>90)` | 0 | Row 9 has neither (97.6 °F, and a heart rate of exactly 90) |
| `=--(F5>100.4)` | 1 | The **double minus** `--` turns TRUE into 1 and FALSE into 0 |
| `=N(F9>100.4)` | 0 | N, `*1`, and `+0` also convert. `--` is the most common in shared workbooks |

Wrap each comparison in its own parentheses, because Excel does arithmetic before it compares. Without them,
`=F5>100.4+G5>90` would add 100.4 and G5 first and give a meaningless result.

Adding comparisons is how clinical scores work: one point for each criterion met, summed. You'll build a SIRS score exactly
this way in the practice, with no IF at all.

> ⚠️ **SUM ignores TRUE and FALSE stored in cells.** If a helper column holds TRUE/FALSE results, SUM over that column returns 0.
> Count the TRUEs with `=COUNTIF(range,TRUE)`, or make the helper column hold 1s and 0s in the first place with `--`.

### 3. IF: one test, two outcomes

```
=IF(logical_test, value_if_true, value_if_false)
```

| Argument | What goes there | Example |
|---|---|---|
| `logical_test` | Any comparison or function that returns TRUE/FALSE | `E3>H3` |
| `value_if_true` | What to return when the test is TRUE: text in quotes, a number, a cell, or another formula | `"High"` |
| `value_if_false` | What to return when the test is FALSE | `""` |

Labs row 3 is a STAT glucose of 132 mg/dL with a reference range of 70–99:

| Formula | Result | What it does |
|---|---|---|
| `=IF(E3>H3,"High","")` | High | Labels the result, otherwise shows nothing |
| `=IF(E3>H3,E3-H3,0)` | 33 | How far above the upper limit the result is |
| `=IF(I3="STAT",60,240)` | 60 | A turnaround target in minutes: 60 for STAT, 240 for routine |
| `=IF(E3>H3,"High")` | High | Works, but returns FALSE whenever the test fails, because value_if_false is missing |

- `""` (two double quotes with nothing between them) is **empty text**. It makes a cell look blank. Section 9 explains why it
  isn't quite the same as a blank cell.
- Always give IF both outcomes. If you leave out value_if_false, IF shows the word FALSE.

> 💡 **Tip:** As you type `=IF(`, Excel shows a ScreenTip with the argument names, and the argument you're on is **bold**. Press
> **Tab** to accept a function name from the AutoComplete list. **Shift + F3** opens the Insert Function dialog on Windows or the
> Formula Builder on a Mac. If you've already typed `=IF(`, it shows a separate box for each argument.

> 📋 **Excel Tables:** the Labs and ED sheets are Excel Tables. When you type a formula into the first cell of an empty Table
> column, Excel fills the whole column for you (a *calculated column*). If you click cells instead of typing their addresses, Excel
> may write `=IF([@ResultValue]>[@RefHigh],"High","")`, where `[@ResultValue]` means "ResultValue in this row." Both styles give
> the same results. Lesson 3.1 covers these structured references.

### 4. Nested IF: more than two outcomes

An IF chooses between exactly two outcomes. For three or more, put a second IF in the value_if_false argument of the first. This is
a **nested IF**.

Example: classify oxygen saturation. Below 90 is "Critical," below 95 is "Low," and anything else is "Normal."

```
=IF(J2<90, "Critical", IF(J2<95, "Low", "Normal"))
```

Here is how Excel reads it for row 23, where SpO2 is 93:

1. Is 93 < 90? No, so Excel moves to value_if_false, which is the second IF.
2. Is 93 < 95? Yes, so the formula returns "Low."

Excel stops at the first test that is TRUE. Row 2 (SpO2 89) passes the first test and returns "Critical" without ever looking at
the second IF. Filled down the ED sheet, this formula labels 54 visits Critical, 55 Low, and 314 Normal.

**The order of the tests matters.** Swap them and the formula breaks without any error message:

```
=IF(J2<95, "Low", IF(J2<90, "Critical", "Normal"))     ← wrong
```

Now row 2 (SpO2 89) passes the first test, because 89 is below 95, and returns "Low." The "Critical" branch can never run, because
any value below 90 is also below 95. When all the tests are thresholds on one number, start with the most extreme band and work
inward, or walk the bands in order from one end to the other.

A two-sided rule, like a lab flag that is High above one limit and Low below another, is different. Its two tests can't both be
TRUE, so either order works, and every value between the limits falls through to the last value_if_false.

Long nested formulas are easier to read and fix with three habits:

- **Count the closing parentheses.** Each IF needs its own, so a three-outcome nested IF ends with `))`. Excel colors each matching
  pair while you edit.
- **Put each IF on its own line.** Press **Alt + Enter** (Mac: **Control + Option + Return**) inside the formula bar. Line breaks
  don't change the result.
- **Widen the formula bar.** Press **Ctrl + Shift + U** (Mac: **Control + Shift + U**) or drag its bottom edge.

Excel allows 64 levels of nesting, but anything past three or four outcomes is hard to check. Use IFS or SWITCH (next sections)
or a lookup table (Lesson 2.6) instead.

### 5. IFS: a flat list of tests

IFS checks pairs of *test, result* in order and returns the result of the **first** test that is TRUE. There's no nesting and
only one closing parenthesis.

```
=IFS(test1, result1, test2, result2, ..., TRUE, result_for_everything_else)
```

Example: temperature categories on the ED sheet.

```
=IFS(F2<96.8, "Low", F2<=100.4, "Normal", F2<103, "Fever", TRUE, "High fever")
```

Row 11 (102.3 °F) fails the first two tests, passes `F2<103`, and returns "Fever." Filled down the sheet, the formula finds 2 Low,
357 Normal, 61 Fever, and 3 High fever readings. Because the tests run in order, `F2<103` only ever sees temperatures above 100.4.
You don't need to write `AND(F2>100.4,F2<103)`.

The last pair, `TRUE, "High fever"`, is the **catch-all**. TRUE is always true, so it catches every row that failed the earlier
tests. IFS has no value_if_false argument, so this is how you say "otherwise." Without a catch-all, a row that fails every test
returns #N/A.

| | Nested IF | IFS |
|---|---|---|
| Available in | Every version | Excel 2019 and later, Microsoft 365, Excel for the web |
| "Otherwise" result | The last value_if_false | `TRUE, result` as the last pair |
| When no test is TRUE | Returns the last value_if_false | Returns #N/A unless you add the TRUE catch-all |
| Limit | 64 levels | 127 test/result pairs |
| Readability | Hard to read past 3 outcomes | Easy to scan |

The rule about order is the same for both: the first TRUE test wins.

> ⚠️ **Version note:** IFS and SWITCH aren't available in Excel 2016 (the one-time-purchase version) or earlier, on Windows or Mac.
> A workbook that uses them shows #NAME? there. If colleagues use older versions, use nested IF.

### 6. AND, OR, NOT, and XOR: combining tests

| Function | Returns TRUE when… | Example on ED row 5 | Result |
|---|---|---|---|
| `AND(test1, test2, …)` | **every** test is TRUE | `=AND(F5>100.4,G5>90)` | TRUE |
| `OR(test1, test2, …)` | **at least one** test is TRUE | `=OR(J5<92,G5>130)` | FALSE |
| `NOT(test)` | the test is FALSE (NOT flips the answer) | `=NOT(C5="Ambulance")` | TRUE |
| `XOR(test1, test2, …)` | an **odd number** of tests are TRUE (with two tests: exactly one) | `=XOR(F5>100.4,G5>90)` | FALSE (both are TRUE) |

Here is every combination for two tests, A and B:

| A | B | AND(A, B) | OR(A, B) | XOR(A, B) |
|:-:|:-:|:-:|:-:|:-:|
| TRUE | TRUE | TRUE | TRUE | FALSE |
| TRUE | FALSE | FALSE | TRUE | TRUE |
| FALSE | TRUE | FALSE | TRUE | TRUE |
| FALSE | FALSE | FALSE | FALSE | FALSE |

On their own these functions return TRUE or FALSE. Put them inside IF's logical_test to return a label:

```
=IF(AND(D2<=2, C2="Walk-In"), "Sick walk-in", "")
```

This flags high-acuity patients (ESI 1 or 2) who walked in instead of arriving by ambulance. In December, 36 visits get the flag.

You can also nest them inside each other. `NOT(OR(…))` means "none of these":

```
=IF(NOT(OR(L2="LWBS", L2="Left AMA")), "Completed visit", "Left early")
```

**De Morgan's laws** let you rewrite a NOT in another form. Use whichever reads more clearly.

| This | is the same as |
|---|---|
| `NOT(OR(A, B))` | `AND(NOT(A), NOT(B))` |
| `NOT(AND(A, B))` | `OR(NOT(A), NOT(B))` |
| `NOT(L2="LWBS")` | `L2<>"LWBS"` |

NOT is most useful for flipping a function that already returns TRUE/FALSE, such as `NOT(ISNUMBER(E2))` in section 9, or for
"none of these" lists like the one above.

> ⚠️ **XOR with three or more tests** is TRUE when an *odd* number of them are TRUE, so `=XOR(TRUE,TRUE,TRUE)` returns TRUE.
> Use XOR to mean "exactly one" only when there are two tests.

> 💡 **Tip:** AND and OR accept up to 255 tests. A long rule is easiest to check with one test per line (Alt + Enter, or
> Control + Option + Return on a Mac).

AND, OR, and NOT work in every version of Excel. XOR needs Excel 2013 or later on Windows, or Excel 2016 or later on a Mac.

### 7. Logic over a whole column in one formula

So far every formula has tested one row. A quality report usually wants a count instead: "How many visits had both signs?" You
can get it two ways.

**Way 1: a helper column.** Put `=AND(F2>100.4,G2>90)` in a spare column, fill it down, and count the TRUEs with
`=COUNTIF(range,TRUE)`. It's easy to check row by row.

**Way 2: one formula that does math on logical values.** When you give a comparison a whole range, it returns a whole list of TRUE/FALSE values,
one per row. (Excel calls a list like this an **array**.) Turn the list into 1s and 0s and add them up:

```
=SUMPRODUCT(--(F2:F424>100.4))     → 64 visits with a fever
```

SUMPRODUCT adds up arrays. It works in every version of Excel and needs no special keystroke. (Lesson 2.4 shows its other job,
multiplying columns together.) In Microsoft 365 and Excel 2021 or later, `=SUM(--(F2:F424>100.4))` gives the same answer.

> ⚠️ **AND, OR, and XOR don't work row by row on ranges.** `=AND(F2:F424>100.4,G2:G424>90)` doesn't return a list. It collapses
> everything into one answer: "is *every* visit febrile and tachycardic?" To combine conditions row by row inside one formula, use
> arithmetic instead.

| Logic | One row (helper column) | Whole range (inside SUMPRODUCT) |
|---|---|---|
| AND | `AND(F2>100.4, G2>90)` | `(F2:F424>100.4)*(G2:G424>90)` |
| OR | `OR(J2<92, G2>130)` | `--(((J2:J424<92)+(G2:G424>130))>0)` |
| NOT | `NOT(F2>100.4)` | `1-(F2:F424>100.4)` |
| XOR (two tests) | `XOR(F2>100.4, G2>90)` | `--((F2:F424>100.4)<>(G2:G424>90))` |

Each translation works for a simple reason:

- **Multiplying is AND.** 1 × 1 = 1 only when both tests are TRUE. A single 0 makes the product 0.
- **Adding is OR.** The sum is at least 1 when any test is TRUE. The `>0` turns a row with two TRUEs (a sum of 2) back into a
  single TRUE, so that row counts once. You can skip it only when the tests can't both be TRUE, like a temperature that is too
  high or too low.
- **Not-equal is XOR.** Two logical values are unequal exactly when one is TRUE and the other is FALSE.
- **One minus is NOT.** 1 − 1 = 0 and 1 − 0 = 1.

For example, this counts STAT lab results that came back above their upper limit:

```
=SUMPRODUCT((I2:I392="STAT")*(E2:E392>H2:H392))     → 52
```

Multiplying already converts TRUE/FALSE to numbers, so you don't need `--` there. You do need it whenever the last step is a
comparison, because a comparison always returns TRUE/FALSE, even when its inputs were numbers. That covers a single test, the `>0` in the OR pattern,
and the `<>` in the XOR pattern, which is why those rows of the table start with `--`. Leave it out and Excel returns 0:
`=SUMPRODUCT(F2:F424>100.4)` is 0, because SUMPRODUCT treats logical values as zeros. Lesson 4.2 takes this kind of array logic
much further.

**IF over a range.** IF also accepts a whole range. `IF(F2:F424>100.4, G2:G424)` returns the heart rate on febrile rows and FALSE
on all the others. AVERAGE skips the FALSEs, so this averages heart rate for febrile visits only:

```
=AVERAGE(IF(F2:F424>100.4, G2:G424))     → 96.34 (the average heart rate of the 64 febrile visits)
```

**"Any" and "all" questions.** Giving AND or OR a range is useful when you *want* a single answer. `=OR(G2:G424>150)` asks "did
any visit have a heart rate above 150?" (TRUE, one visit at 157.) `=AND(J2:J424>=80)` asks "was every SpO2 at least 80?" (TRUE.)

> 📋 **Version note:** Microsoft 365 and Excel 2021 or later calculate array formulas like `AVERAGE(IF(…))`, `SUM(--(…))`, and
> `OR(range>…)` automatically. In Excel 2019 and earlier, finish them with **Ctrl + Shift + Enter** (Mac: **⌘ + Shift + Return**),
> and Excel shows the formula inside braces `{ }`. SUMPRODUCT never needs this. AVERAGEIF (Lesson 2.5) also handles simple
> conditional averages.

### 8. SWITCH: map codes to labels

When you compare one cell against a list of exact values, SWITCH is shorter than a nested IF, and you write the cell reference
only once.

```
=SWITCH(expression, value1, result1, value2, result2, ..., [default])
```

Example: group ED dispositions for a throughput report.

```
=SWITCH(L2, "Discharged","Home", "Admitted","Inpatient", "Observation","Inpatient", "Transferred","Other hospital", "Left early")
```

Row 5 (Admitted) returns "Inpatient." Row 6 (LWBS, *left without being seen*) matches none of the listed values, so it gets the
**default**, "Left early." Filled down, December has 246 Home, 159 Inpatient, 3 Other hospital, and 15 Left early.

- The default is the final argument with no partner. Without one, a value that matches nothing returns #N/A.
- SWITCH only tests whether values are **exactly equal**. For bands like "4 to 6," use IFS.
- Text matching ignores case, just like `=`.

| Use | When |
|---|---|
| IF | One test, two outcomes |
| Nested IF | Three or four outcomes, or colleagues on Excel 2016 and earlier |
| IFS | Several outcomes from thresholds or mixed tests |
| SWITCH | One cell compared with a list of exact codes |
| Lookup table (Lesson 2.6) | Many codes, or codes that change over time |

> 💡 **Tip:** Leaving out the default can be a feature. If an ESI code of 6 ever sneaks into the data, a SWITCH with no default
> returns #N/A and you notice. A default like "Unknown" would hide the bad code.

### 9. Blanks, text that looks like a number, and "contains"

Logic formulas run into messy data. These three traps cause most wrong answers.

**An empty cell counts as 0, and also as empty text.** Triage blood pressure wasn't recorded for 12 December visits, such as row 13
(ED211716). For that row:

| Formula | Result | Why |
|---|---|---|
| `=I13<=100` | TRUE | The empty cell is treated as 0, and 0 ≤ 100 |
| `=I13=""` | TRUE | An empty cell also equals empty text |
| `=ISBLANK(I13)` | TRUE | The cell holds nothing at all |
| `=AND(I13<>"", I13<=100)` | FALSE | The guard `I13<>""` fails, so AND is FALSE |

So a rule like "systolic BP of 100 or less" scores every missing BP as low unless you guard it with `I2<>""`. You'll need that
guard in the bonus.

**Empty text is not the same as an empty cell.** A formula like `=IFERROR(G2/I2,"")` can show nothing, but the cell still holds a
formula. After you fill the ShockIndex column in task 12, cell Q13 looks empty, yet `=ISBLANK(Q13)` returns FALSE while `=Q13=""`
returns TRUE. Test with `=""` when formulas might return empty text. Also, never type a space to "clear" a cell. A space is a
character, so `=A1=""` returns FALSE for it.

**Text that looks like a number.** Excel ranks every number below every piece of text, and TRUE/FALSE above both. So `="9">10`
returns TRUE. Lab systems sometimes export a result as text, like `"<3"` for a troponin below the detection limit. That text is
"greater than" any number, so a plain `E2>H2` test would flag it High. Check the type first:

```
=IF(NOT(ISNUMBER(E2)), "Check result", IF(E2>H2, "High", "Not high"))
```

Every result on the Labs sheet is a real number, so you won't hit this in the practice. You will at work.

| Function | Returns TRUE when the value is… |
|---|---|
| `ISBLANK(x)` | a truly empty cell |
| `ISNUMBER(x)` | a number (dates and times count) |
| `ISTEXT(x)` | text, including the empty text `""` |
| `ISLOGICAL(x)` | TRUE or FALSE |
| `ISERROR(x)` | any error value |
| `ISNA(x)` | the #N/A error only |

**"Contains" tests.** SEARCH looks for text inside other text and returns its position, or the #VALUE! error if the text isn't
there. ISNUMBER turns that into TRUE or FALSE:

```
=ISNUMBER(SEARCH("fever", E5))     → TRUE  (row 5's complaint is "Painful Urination / Fever")
```

SEARCH ignores case, so "fever" finds "Fever." In December, 116 chief complaints contain "fever." Lesson 2.2 covers SEARCH and FIND
in depth, and this pattern is all you need for now.

### 10. Trapping errors with IFERROR and IFNA

| Error | Usual cause in logic formulas |
|---|---|
| `#DIV/0!` | Dividing by zero or by an empty cell, such as a missing blood pressure |
| `#N/A` | "Value not available": IFS or SWITCH found no match, or a lookup found nothing (Lesson 2.6) |
| `#VALUE!` | The wrong type of data, such as arithmetic on text |
| `#NAME?` | A misspelled function, text missing its quotes, or a function your version lacks (IFS in Excel 2016) |

```
=IFERROR(value, value_if_error)
=IFNA(value, value_if_na)
```

Both functions return `value` when it calculates normally. They differ in which errors they replace:

| | IFERROR | IFNA |
|---|---|---|
| Replaces | Every error type | Only #N/A |
| Typical use | A calculation with one known failure case | Lookups, and SWITCH or IFS with no match |
| Available in | Excel 2007 and later | Excel 2013 and later (Excel 2016 and later on a Mac) |

For example, a **shock index** is heart rate ÷ systolic BP. Row 50 (ED211772) has no recorded BP, so `=G50/I50` returns #DIV/0!.
`=IFERROR(G50/I50,"")` returns empty text instead, and the column stays readable.

**Don't hide every error.** IFERROR can't tell an expected problem from a mistake. `=IFERROR(AVERAG(G2:G424),0)` returns 0,
because the misspelled function name produces #NAME? and IFERROR quietly replaces it. The report then shows an average heart rate
of 0 and nobody notices. Four habits prevent this:

1. Get the formula working *before* you wrap it in IFERROR.
2. Prefer a test for the specific problem you expect. `=IF(I2="","",G2/I2)` handles the missing BP and still shows any other
   error.
3. Use IFNA when #N/A is the only error you expect, so everything else stays visible.
4. Choose the fallback carefully. `""` or `"n/a"` says "no value." A 0 says "the value is zero," and zeros drag down averages.

### 11. Worked example: a red-flag vital-sign check

*Question: How many December ED visits had at least one "red flag" vital sign: SpO2 of 91 or less, a heart rate of 131 or more,
or a respiratory rate of 25 or more?* (These thresholds are borrowed from early-warning scores and simplified for teaching. They
are not clinical guidance.)

1. **Test each part on one row.** Pick a row you can check by eye. Row 6 (ED211703) has SpO2 82, heart rate 134, and respiratory
   rate 39. In three spare cells, type `=J6<=91`, `=G6>=131`, and `=H6>=25`. All three return TRUE.
2. **Combine the parts.** `=OR(J6<=91,G6>=131,H6>=25)` returns TRUE. Try it on row 4 (SpO2 98, heart rate 94, respiratory rate
   20) as well. It returns FALSE.
3. **Score instead of flag, if severity matters.** `=(J6<=91)+(G6>=131)+(H6>=25)` returns 3 for row 6. Row 8 (SpO2 88, heart rate
   120, respiratory rate 37) scores 2.
4. **Label every row.** `=IF(OR(J2<=91,G2>=131,H2>=25),"Red flag","")` in a helper column, filled down.
5. **Count the whole sheet in one cell.**

   ```
   =SUMPRODUCT(--(((J2:J424<=91)+(G2:G424>=131)+(H2:H424>=25))>0))     → 103
   ```

   The `>0` matters here, because one patient can have two or three red flags. Without it, those patients would count two or three
   times.
6. **Sanity-check the answer.** Count each test alone: 79 low SpO2 readings, 20 high heart rates, and 67 high respiratory rates.
   An OR count can't be smaller than the largest single count (79) or bigger than their sum (166). 103 sits between them, so the
   formula passes. Changing `>0` to `>=2` shows that 53 visits had two or more red flags.

> 💡 **Tip:** Step 1 is the most reliable way to debug any logic formula: test each piece in its own cell before you combine them.
> In Excel for Windows, **Formulas → Evaluate Formula** also steps through a formula one calculation at a time, and you can select
> part of a formula in the formula bar and press **F9** to see its value. Press **Esc** afterwards so the value doesn't replace
> that part of your formula.

### 12. Quick reference

| Function | Syntax | Returns | Available in |
|---|---|---|---|
| IF | `IF(test, value_if_true, [value_if_false])` | One of two results | Every version |
| IFS | `IFS(test1, result1, …, TRUE, otherwise)` | Result of the first TRUE test | Excel 2019+, Microsoft 365 |
| AND / OR | `AND(test1, test2, …)` | TRUE if all / any are TRUE | Every version |
| NOT | `NOT(test)` | The opposite of the test | Every version |
| XOR | `XOR(test1, test2, …)` | TRUE if an odd number are TRUE | Excel 2013+ (Mac: 2016+) |
| SWITCH | `SWITCH(value, match1, result1, …, [default])` | Result for the first exact match | Excel 2019+, Microsoft 365 |
| IFERROR | `IFERROR(value, value_if_error)` | value, or the fallback for any error | Excel 2007+ |
| IFNA | `IFNA(value, value_if_na)` | value, or the fallback for #N/A | Excel 2013+ (Mac: 2016+) |
| SUMPRODUCT | `SUMPRODUCT(--(range>x))` | The number of TRUEs in a test over a range | Every version |

| Action | Windows | Mac |
|---|---|---|
| Accept a function from the AutoComplete list | Tab | Tab |
| Insert Function (Windows) or Formula Builder (Mac) | Shift + F3 | Shift + F3 |
| Edit the active cell | F2 | Control + U |
| New line inside a formula | Alt + Enter | Control + Option + Return |
| Expand or collapse the formula bar | Ctrl + Shift + U | Control + Shift + U |
| Show formulas instead of results | Ctrl + ` | Control + ` |
| Enter a legacy array formula (Excel 2019 and earlier) | Ctrl + Shift + Enter | ⌘ + Shift + Return |

## 🧪 Hands-on practice

Download [`2.1-logical-functions.xlsx`](2.1-logical-functions.xlsx) and open the **Practice** sheet. Type each answer in the yellow
cell, as a formula wherever possible. For the column tasks, fill the yellow column on the Labs or ED sheet and the gray cell on the
Practice sheet summarizes your work. The **Check** column turns green when you're right.

<!-- BEGIN GENERATED: practice -->
Tasks 1–3 use the Labs sheet. Tasks 4–13 use the ED sheet. Several tasks ask you to fill one of the yellow columns on a data sheet: type the formula in the first data row and Excel fills the rest of the Table column for you (if it doesn't, double-click the fill handle). The gray cell on this sheet then summarizes your column. When a formula on this sheet points at a data sheet, include the sheet name (for example Labs!E38), or click the cell on that sheet and Excel adds the name for you.

| # | Task | Hint |
|:-:|------|------|
| 1 | Labs row 38 holds a White Blood Cell Count result of 11 K/uL. Write a comparison that returns TRUE if that ResultValue is greater than its RefHigh, and FALSE if not. | A comparison needs no function: just =cell>cell. Include the sheet name (Labs!) in both references |
| 2 | How many of the lab results are above their reference high (ResultValue greater than RefHigh)? Use one formula that turns every row's comparison into a 1 or 0 and adds them up. | -- turns TRUE/FALSE into 1/0, and SUMPRODUCT adds them up |
| 3 | Fill the yellow Flag column on the Labs sheet with a nested IF: "High" if ResultValue is greater than RefHigh, "Low" if it is less than RefLow, otherwise "Normal". A result exactly on a limit is Normal. Start in J2. The gray cell counts your "Normal" flags. | IF(test, "High", IF(test, "Low", "Normal")) |
| 4 | Fill the yellow PainLevel column on the ED sheet with IFS: PainScore 0 is "None", 1–3 is "Mild", 4–6 is "Moderate", and 7–10 is "Severe". Start in M2. The gray cell counts your "Moderate" rows. | Test from the lowest band up, and end with TRUE as the catch-all |
| 5 | How many ED visits had an abnormal temperature: TempF above 100.4 OR below 96.8? | Adding two TRUE/FALSE lists works like OR when both can't be TRUE at once |
| 6 | How many visits had BOTH a fever (TempF above 100.4) AND tachycardia (HeartRate above 90)? | Multiplying two TRUE/FALSE lists works like AND |
| 7 | The sepsis coordinator already reviews visits with both signs. How many visits had exactly ONE of the two: a fever (TempF above 100.4) or tachycardia (HeartRate above 90), but not both? | This is XOR logic. Use XOR in a helper column, or compare the two TRUE/FALSE lists with <> |
| 8 | Fill the yellow FastTrack column: "Fast Track" when ESILevel is 4 or 5 AND ArrivalMode is "Walk-In" AND the ChiefComplaint is NOT "Chest Pain" or "Shortness of Breath" (it is neither one). Every other visit is "Main ED". Start in N2. The gray cell counts your "Fast Track" rows. | IF(AND(…, …, NOT(OR(…, …))), "Fast Track", "Main ED") |
| 9 | Fill the yellow SIRS column with a vital-sign SIRS score from 0 to 3: one point for an abnormal temperature (TempF above 100.4 OR below 96.8), one for HeartRate above 90, and one for RespRate above 20. Start in O2. The gray cell adds up your whole column. | TRUE + TRUE = 2. Add three TRUE/FALSE tests together |
| 10 | How many visits are SIRS-positive, meaning a score of 2 or more in your SIRS column? | COUNTIF with a ">=2" criterion |
| 11 | Fill the yellow ESIName column with SWITCH: ESILevel 1 is "Resuscitation", 2 is "Emergent", 3 is "Urgent", 4 is "Less Urgent", 5 is "Non-Urgent". Start in P2. The gray cell counts your "Emergent" rows. | SWITCH(value, 1, "…", 2, "…", …) |
| 12 | Shock index = HeartRate ÷ SystolicBP, and 1.0 or more is a warning sign. Fill the yellow ShockIndex column, wrapping the division in IFERROR so the 12 visits with no SystolicBP show a blank ("") instead of #DIV/0!. Start in Q2. The gray cell counts visits with a shock index of 1.0 or higher (it asks you to fix any errors first). | IFERROR(value, value_if_error) |
| 13 | Visit ED211716 (ED row 13) has no SystolicBP. In the yellow cell, calculate its shock index again, but wrap the division in IFNA(…, "") instead of IFERROR. The check turns green when your cell shows the error that IFNA lets through. | IFNA traps only one kind of error |
<!-- END GENERATED: practice -->

## ✅ Answer key

The workbook has a hidden **Answer Key** sheet (right-click any sheet tab → **Unhide…** → *Answer Key*). Its *Live result* column
runs a one-cell version of each solution, so you can see it working. The same answers are below, collapsed so you don't see them
by accident.

<!-- BEGIN GENERATED: answers -->
<details>
<summary><b>🔑 Show the answer key</b> — Try every task before opening this.</summary>

**1. Labs row 38 holds a White Blood Cell Count result of 11 K/uL. Write a comparison that…**

- **Answer:** FALSE
- **Solution:** `=Labs!E38>Labs!H38`

A comparison always returns TRUE or FALSE. Here the result (11) is exactly equal to the upper limit (11), and *equal* is not *greater than*, so the answer is FALSE. With `>=` it would be TRUE. Choosing between `>` and `>=` is a definition decision, so always check the rule you were given.

**2. How many of the lab results are above their reference high (ResultValue greater than…**

- **Answer:** 114
- **Solution:** `=SUMPRODUCT(--(Labs!E2:E392>Labs!H2:H392))`

Comparing two ranges row by row gives a list of TRUE/FALSE values. The double minus (`--`) converts them to 1s and 0s, and SUMPRODUCT adds them. In Microsoft 365 and Excel 2021 or later, `=SUM(--(…>…))` works too. Plain `SUM` over a column of TRUE/FALSE *cells* returns 0, because SUM ignores logical values stored in cells.

**3. Flag column with nested IF (count of Normal)**

- **Answer:** 254
- **Solution:** `=IF(E2>H2,"High",IF(E2<G2,"Low","Normal"))`

The first IF handles High. Its *value_if_false* argument is a second IF that handles Low, and "Normal" is whatever is left. Because the tests use `>` and `<`, the 14 results sitting exactly on a limit fall through to Normal. In the Table you may see it as `=IF([@ResultValue]>[@RefHigh],"High",IF([@ResultValue]<[@RefLow],"Low","Normal"))`, which is the same formula.

**4. PainLevel column with IFS (count of Moderate)**

- **Answer:** 40
- **Solution:** `=IFS(K2=0,"None",K2<=3,"Mild",K2<=6,"Moderate",TRUE,"Severe")`

IFS returns the result of the **first** test that is TRUE, so order matters. Testing `<=3` before `<=6` means a score of 2 stops at "Mild" and never reaches "Moderate". The final `TRUE,"Severe"` pair is the catch-all for everything left (7–10). Without it, any value that fails every test returns #N/A.

**5. How many ED visits had an abnormal temperature: TempF above 100.4 OR below 96.8?**

- **Answer:** 66
- **Solution:** `=SUMPRODUCT((ED!F2:F424>100.4)+(ED!F2:F424<96.8))`

`+` acts as OR over whole ranges, because each row scores 1 if either test is TRUE. That's safe here because no temperature can be both above 100.4 and below 96.8. (When both tests *can* be TRUE, write `--((…)+(…)>0)` so a row counts once. The `--` is needed because `>0` turns the sums back into TRUE/FALSE.) `=COUNTIF(ED!F2:F424,">100.4")+COUNTIF(ED!F2:F424,"<96.8")` gives the same answer.

**6. How many visits had BOTH a fever (TempF above 100.4) AND tachycardia (HeartRate above 90)?**

- **Answer:** 44
- **Solution:** `=SUMPRODUCT((ED!F2:F424>100.4)*(ED!G2:G424>90))`

`*` acts as AND over whole ranges. 1 × 1 = 1 only when both tests are TRUE, and anything × 0 = 0. `=SUMPRODUCT(--AND(…))` does **not** work, because AND collapses the whole range into a single TRUE or FALSE instead of testing row by row. To use the AND function itself, put `=AND(F2>100.4,G2>90)` in a helper column and count the TRUEs with `COUNTIF(range,TRUE)`.

**7. The sepsis coordinator already reviews visits with both signs. How many visits had…**

- **Answer:** 189
- **Solution:** `=SUMPRODUCT(--((ED!F2:F424>100.4)<>(ED!G2:G424>90)))`

"Exactly one of two" is XOR. `=XOR(F2>100.4,G2>90)` answers it for one row, so a helper column plus `COUNTIF(range,TRUE)` works. Over whole ranges, `<>` is XOR: TRUE<>FALSE is TRUE, while TRUE<>TRUE and FALSE<>FALSE are FALSE. Like AND and OR, the XOR function collapses a range into one value, so `SUMPRODUCT(--XOR(range>…, range>…))` gives the wrong answer. Check: fever count + tachycardia count − 2 × both = exactly one.

**8. FastTrack column with IF + AND + OR + NOT (count of Fast Track)**

- **Answer:** 136
- **Solution:**

```
=IF(AND(D2>=4,C2="Walk-In",NOT(OR(E2="Chest Pain",E2="Shortness of Breath"))),"Fast Track","Main ED")
```


AND needs all three parts to be TRUE. The third part, `NOT(OR(complaint="Chest Pain", complaint="Shortness of Breath"))`, is TRUE only when the complaint is neither one. `AND(…<>"Chest Pain", …<>"Shortness of Breath")` is the same test written the other way round (De Morgan's law). `ESILevel>=4` covers 4 and 5 because ESI only goes up to 5. Text comparisons ignore case, so "walk-in" also matches.

**9. SIRS score column (total points)**

- **Answer:** 397
- **Solution:** `=OR(F2>100.4,F2<96.8)+(G2>90)+(H2>20)`

When you do arithmetic on TRUE/FALSE, Excel treats TRUE as 1 and FALSE as 0, so adding three tests gives a score from 0 to 3 with no IF at all. The temperature point uses OR because *either* direction counts, but only once. If your column shows TRUE/FALSE instead of numbers, you used AND/OR around everything. The full SIRS criteria also include the white-blood-cell count, which isn't on this sheet.

**10. How many visits are SIRS-positive, meaning a score of 2 or more in your SIRS column?**

- **Answer:** 115
- **Solution:** `=COUNTIF(ED!O2:O424,">=2")`

Once the score is a number, COUNTIF can count any threshold. The key's live formula does the whole thing in one cell: it builds the score for every row as an array, then counts the rows where it is `>=2`.

**11. ESIName column with SWITCH (count of Emergent)**

- **Answer:** 87
- **Solution:** `=SWITCH(D2,1,"Resuscitation",2,"Emergent",3,"Urgent",4,"Less Urgent",5,"Non-Urgent")`

SWITCH compares one value against a list and returns the result paired with the first exact match. It's shorter than five nested IFs, and you write the cell reference only once. Leaving out a default is a deliberate choice here. If a bad code like 6 ever appears, SWITCH returns #N/A and the problem stays visible.

**12. ShockIndex column with IFERROR (count of 1.0 or higher)**

- **Answer:** 54
- **Solution:** `=IFERROR(G2/I2,"")`

Dividing by an empty cell divides by zero, so those rows show #DIV/0!. `IFERROR(G2/I2,"")` returns the division when it works and an empty text string when it fails. COUNTIF then skips the blanks because they are text. `=IF(I2="","",G2/I2)` is an even better formula, because it handles the one problem you expect (a missing BP) and still lets any *other* mistake show up as an error.

**13. IFNA vs IFERROR experiment**

- **Answer:** #DIV/0! (IFNA traps only #N/A, so the divide-by-zero error passes through)
- **Solution:** `=IFNA(ED!G13/ED!I13,"")`

IFERROR catches every error type. IFNA catches only #N/A, the "value not available" error that lookups return when they can't find a match (Lesson 2.6). A divide-by-zero is a different error, so IFNA hands it straight back. That's the point of IFNA. It hides the one error you expect and leaves every other mistake visible.

</details>
<!-- END GENERATED: answers -->

## 🏆 Bonus challenge

<!-- BEGIN GENERATED: bonus -->
The sepsis committee wants to know what a quick qSOFA-style screen would flag if triage nurses scored every ED patient. This simplified, educational version (not clinical guidance) gives one point for each of three findings. (1) RespRate is 22 or more. (2) SystolicBP is 100 or less. A blank (not recorded) SystolicBP earns NO point. (3) Altered mentation, meaning the ChiefComplaint contains the word "Altered" or "Confusion" anywhere. A score of 2 or more is screen-positive. Build the score in the yellow qSOFA column on the ED sheet, then answer the questions.

Work on the **Bonus** sheet of the workbook.

- **B1.** Fill the qSOFA column (0–3) on the ED sheet, starting in R2. The gray cell adds up your whole column. What is the total number of qSOFA points? *(Hint: Add three tests. Guard the BP test with I2<>"". ISNUMBER(SEARCH("altered", E2)) tests 'contains')*
- **B2.** How many visits are screen-positive (a qSOFA score of 2 or more)? *(Hint: COUNTIF on your qSOFA column)*
- **B3.** What was the average HeartRate of the screen-positive visits? Round to 1 decimal place. *(Hint: AVERAGE(IF(test_range>=2, values_range)). AVERAGEIF (Lesson 2.5) also works)*
- **B4.** Safety check: how many screen-positive visits ended with EDDisposition "Discharged" (sent home)? *(Hint: Two conditions over whole columns: multiply the TRUE/FALSE lists)*
- **B5.** The two screens don't always agree. How many visits are qSOFA screen-positive (2 or more) but NOT SIRS-positive (SIRS score below 2 in your SIRS column from task 9)? *(Hint: Same pattern as B4. 'NOT SIRS-positive' is simply SIRS<2)*
<!-- END GENERATED: bonus -->

<!-- BEGIN GENERATED: bonus-answers -->
<details>
<summary><b>🔑 Show the bonus solution</b> — Give it a real try first!</summary>

**B1. qSOFA column (total points)**

- **Answer:** 165
- **Solution:**

```
=(H2>=22)+AND(I2<>"",I2<=100)+OR(ISNUMBER(SEARCH("altered",E2)),ISNUMBER(SEARCH("confusion",E2)))
```


Each part is a TRUE/FALSE test, and adding them gives the score. Watch for two traps. First, an empty cell counts as 0 in a comparison, so a bare `I2<=100` gives a point to every visit with no BP. `AND(I2<>"",I2<=100)` scores only a recorded BP. Without that guard the 12 visits with no BP each gain a point, so the total is 177 instead of 165. Second, SEARCH returns the position of the text, or #VALUE! when it isn't there, so `ISNUMBER(SEARCH(…))` turns that into TRUE/FALSE. SEARCH ignores case, so "altered" finds "Altered". Wrapping the two SEARCH tests in OR keeps the mentation point at 1 even if a complaint ever mentions both words.

**B2. How many visits are screen-positive (a qSOFA score of 2 or more)?**

- **Answer:** 31
- **Solution:** `=COUNTIF(ED!R2:R424,">=2")`

Once the score is a number, COUNTIF counts any threshold. If your column scored the missing BPs, you'd count 35 here instead of 31, so the blank-cell guard changes a real quality number.

**B3. What was the average HeartRate of the screen-positive visits? Round to 1 decimal place.**

- **Answer:** 115.8
- **Solution:** `=ROUND(AVERAGE(IF(ED!R2:R424>=2,ED!G2:G424)),1)`

`IF(range>=2, HeartRate)` returns the heart rate for screen-positive rows and FALSE for the rest, and AVERAGE skips FALSE. Microsoft 365 and Excel 2021 evaluate this array formula automatically. In Excel 2019 and earlier, confirm it with Ctrl + Shift + Enter (Mac: ⌘ + Shift + Return). `=AVERAGEIF(ED!R2:R424,">=2",ED!G2:G424)` gives the same result.

**B4. Safety check: how many screen-positive visits ended with EDDisposition "Discharged"…**

- **Answer:** 19
- **Solution:** `=SUMPRODUCT((ED!R2:R424>=2)*(ED!L2:L424="Discharged"))`

Multiplying the two lists keeps a 1 only where both are TRUE (AND logic), and SUMPRODUCT counts them. These are the charts a sepsis committee would pull for review, because the patients screened positive but went home.

**B5. The two screens don't always agree. How many visits are qSOFA screen-positive (2 or…**

- **Answer:** 3
- **Solution:** `=SUMPRODUCT((ED!R2:R424>=2)*(ED!O2:O424<2))`

NOT(score>=2) is the same as score<2, so you don't need the NOT function over a range. SIRS looks at temperature, heart rate, and breathing. qSOFA looks at breathing, blood pressure, and mental status. Each catches patients the other misses, which is why hospitals don't rely on one screen alone.

</details>
<!-- END GENERATED: bonus-answers -->

## Key takeaways

- A comparison (`=`, `<>`, `>`, `<`, `>=`, `<=`) returns TRUE or FALSE. Equal is not greater, so match the operator to the exact
  wording of the rule.
- In arithmetic TRUE is 1 and FALSE is 0. Adding tests builds a score, and `SUMPRODUCT(--(…))` counts the rows that pass a test.
- IF picks one of two outcomes. For more, nest IFs or use IFS, put the tests in an order where the first TRUE is the right one,
  and end IFS with a `TRUE` catch-all.
- AND means all, OR means any, NOT flips, and XOR means exactly one of two. Over whole ranges, use `*`, `+` with `>0`, `1-`, and
  `<>` instead, and put `--` in front when the last step is a comparison.
- SWITCH maps exact codes to labels. Leave out the default when you want unexpected codes to show up as #N/A.
- An empty cell counts as 0 in a comparison, so guard tests on optional values with `<>""`.
- IFERROR hides every error and IFNA hides only #N/A. Test for the problem you expect, and wrap only formulas you know work.

<!-- BEGIN GENERATED: nav -->
---

⬅️ **Previous:** [1.6 Sorting & Filtering Data](../../01-foundations/06-sorting-filtering/README.md) · 🏠 [Course home](../../README.md) · **Next:** [2.2 Text Functions](../02-text-functions/README.md) ➡️
<!-- END GENERATED: nav -->
