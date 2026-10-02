# Course tooling & authoring guide

Everything learners download (lesson workbooks, the answer sections of each lesson README, the datasets) is
**generated** from code in this folder, so answers can never drift out of sync with the data.

```
tools/
├── generate_data.py      # seeded generator → data/*.csv (+ data/README.md data dictionary)
├── catalogs.py           # reference lists used by the generator (facilities, ICD-10 codes, drugs, supplies…)
├── curriculum.py         # lesson order, folders, objectives, and the author spec for every lesson
├── build.py              # build + verify lessons; refresh course READMEs
├── xlcourse/             # shared library
│   ├── lesson.py         #   Lesson / Task → workbook + README blocks
│   ├── data.py           #   typed CSV loaders
│   ├── xlfn.py           #   formula → file-format conversion (_xlfn. prefixes, [@Col])
│   ├── lo.py             #   LibreOffice (UNO) recalculation used by the verifier
│   └── vba.py            #   best-effort VBA smoke tests in LibreOffice
├── lessons/lesson_X_Y.py # one builder per lesson (build() -> Lesson)
└── _truth/               # answer-key-only truth files (never shown to learners)
```

## Setup

```bash
pip install openpyxl
# LibreOffice Calc 26.x + Python UNO bridge (Ubuntu 24.04: from noble-backports)
sudo apt-get install -y --no-install-recommends -t noble-backports libreoffice-calc libreoffice-core python3-uno
```

LibreOffice 26.x evaluates XLOOKUP, XMATCH, FILTER, SORT, SORTBY, UNIQUE, SEQUENCE, LET, TEXTBEFORE/AFTER, TEXTSPLIT,
VSTACK/HSTACK, TAKE/DROP, IFS, SWITCH, MAXIFS, TEXTJOIN, … It does **not** evaluate LAMBDA, MAP, REDUCE, SCAN, BYROW,
BYCOL, or MAKEARRAY.

## Commands

```bash
python tools/generate_data.py          # regenerate data/ (deterministic)
python tools/build.py 1.4              # build lesson 1.4's workbook + README blocks, then verify in LibreOffice
python tools/build.py 1.4 --no-verify  # faster iteration
python tools/build.py all              # every lesson with a builder
python tools/build.py readme           # refresh the syllabus in README.md and the module READMEs
```

## What `build.py` verifies

1. Every **live formula** in the hidden Answer Key equals the Python-computed answer (within tolerance).
2. A **self-test** copy of the workbook, with each sample solution typed into its answer cell (or the `fill` formulas
   filled down), shows `✔ Correct` in every Check cell.
3. The pristine workbook's Check cells are blank (nothing is pre-answered), key sheets are hidden, and there are no error
   values on the Practice/Bonus/Key sheets.

A lesson is done only when `python tools/build.py X.Y` prints `verification: PASS`.

---

# Authoring a lesson

Read the reference lesson first: [`lessons/lesson_1_4.py`](lessons/lesson_1_4.py) and
[`01-foundations/04-basic-formulas/README.md`](../01-foundations/04-basic-formulas/README.md). Your lesson's spec
(objectives, data, practice ideas, bonus idea, and what NOT to cover) is in [`curriculum.py`](curriculum.py).

## Files you create

| File | Purpose |
|---|---|
| `tools/lessons/lesson_X_Y.py` | `build() -> Lesson`: data sheets, tasks, bonus, optional `customize` hooks |
| `<module>/<slug>/README.md` | The lesson guide. Generated blocks (between `<!-- BEGIN/END GENERATED: … -->`) are filled by the build |
| `<module>/<slug>/<code>-<name>.xlsx` | Written by the build; never edit by hand |
| optional `starter/`, `solutions/`, `data/` subfolders | VBA modules, Power Query source files, etc. (write them from the builder via `lesson.extra_files` or by hand) |

## Lesson README structure (keep this order)

1. `# Lesson X.Y · Title`, then a quote line: Level · Time · Workbook link · Data description
2. A short hook paragraph: why this skill matters in a hospital/health-system job
3. `## What you'll learn` (the curriculum objectives)
4. `## 📖 Guide`: the real teaching. Numbered `###` sections, syntax blocks, tables comparing functions, worked examples on
   the lesson data, ⚠️ pitfalls, 💡 tips, keyboard shortcuts for **Windows and Mac**, and **version notes** (e.g. "XLOOKUP
   needs Microsoft 365 or Excel 2021+"). Teach everything the practice tasks need. Aim for the depth of the reference lesson.
5. `## 🧪 Hands-on practice`: one or two sentences, then the `practice` generated block
6. `## ✅ Answer key`: one or two sentences about the hidden sheet, then the `answers` generated block (collapsed `<details>`)
7. `## 🏆 Bonus challenge`: the `bonus` block, then the `bonus-answers` block (collapsed)
8. `## Key takeaways`: 4–7 bullets
9. The `nav` generated block (previous/next links)

If the README doesn't exist, the first build writes a skeleton with all markers. Replace every `TODO`.

## Writing style for READMEs

Learners read these lessons in a browser, often while switching back and forth to Excel. Write so they can follow along
without rereading.

- **Address the learner as "you."** Use present tense and imperative mood for steps: "Select the column, then press
  Ctrl + Shift + L."
- **Write in plain sentences.** Put one idea in each sentence and connect ideas with words like *because*, *so*, and *then*
  rather than dashes, arrows, or semicolons. Arrows are fine for menu paths (**Data → Sort**) and in tables.
- **Define a term the first time you use it**, in bold, and use that same term for the rest of the lesson. Don't switch
  between synonyms.
- **Use healthcare examples from the lesson data** instead of generic ones like "sales" or "widgets."
- **Name shortcuts with both platforms:** "Ctrl + Shift + L (Mac: ⌘ + Shift + F)".
- **Use the callouts** `💡 Tip:` and `⚠️` (pitfall), and add `📋` when you need a side note. Don't add other kinds of
  callout.
- **Use tables** to compare functions or options and **numbered lists** for procedures. Use prose for explanations.
- **Write explanations in tasks the same way.** The `explanation` field appears in the README answer key, so it should say
  *why* the solution works, not just restate the formula.
- **Leave out filler:** no "In this section we will…", no "Let's dive in," and no closing pep talk.

## Task design rules (the important part)

- **Compute every answer in Python from the data**, never by hand. Put the computation next to the `Task` so a reviewer
  can follow it.
- **Make prompts unambiguous.** State the exact rows or filters, the units, the rounding ("to 1 decimal place"), and the
  format ("as a percentage"). If two reasonable readings give different answers, rewrite the prompt.
- **Don't number the prompts.** The library numbers tasks (1, 2, … and B1, B2, …).
- **Hints nudge.** Give a function name or approach, never the answer.
- **Write solutions the way a learner would type them** into the yellow answer cell on the Practice sheet. Reference the
  data sheets explicitly (`Census!C2:C91`, or Table syntax like `tblCensus[Admissions]`). Use the `SheetData` helpers
  (`sd.rng("Col")`, `sd.col("Col")`, `sd.first_row`, `sd.last_row`) instead of hard-coding addresses.
- **Use `fmt`** for anything that isn't an integer: `'0.0%'` for percentages, which makes the check accept both
  `0.913` and `91.3`; `'#,##0.00'` for money; `'mm/dd/yyyy'` for dates.
- **Choose the check type.** It's inferred from the answer type (number, percent, date, datetime, duration via
  `timedelta`, bool, text). Text checks are case-insensitive and trimmed; add `accept=[...]` for equivalent answers
  (`"00000000"`, `"0000000#"`). Integers must match exactly. Floats default to ±0.0051, so tell learners how to round,
  or set `tol=`.
- **Column-filling tasks** ("fill this new column"): add the column with `extra_cols=[...]` on the data sheet, set
  `summary=` to a gray formula that condenses the learner's column into one number, and set
  `fill={"range": "Sheet!G2:G91", "formula": "=E2/B2"}` so the self-test can fill it. **The summary must return `""`
  until the learner starts**, for example `=IF(COUNT(range)=0,"",SUM(range))`. Set `live=` to an equivalent one-cell
  formula for the key, or `live=False`.
- **Manual tasks** (build a chart, apply a format, record a macro): use `check="manual"` (answer `None`) and write clear
  numbered steps in `solution` as Markdown. Prefer pairing every manual build step with a **checkable question** about
  the result ("…then type how many rows your filter shows").
- **Spill and dynamic-array formulas never go in answer cells,** because they'd spill into the Check column. Ask for a
  scalar (`ROWS(UNIQUE(...))`, `INDEX(SORT(...),5)`), or use a separate `Workspace` sheet with anchor cells and a summary
  that counts the spill over a generous range.
- **Never reference a Table name, defined name, or sheet that doesn't exist in the file.** Excel can refuse to open
  the file. For sheets that a learner's macro will create, use `INDIRECT("'NewSheet'!A1")`, which safely returns `#REF!`
  until the sheet exists.
- **Never use `TODAY()`/`NOW()`** in graded answers. The course "as of" date is **12/31/2025**: put it in a
  `ReportDate` cell when a task needs "today."
- **Use `live=False`** when a solution refers to Practice-sheet cells (the key lives on another sheet) or uses functions
  LibreOffice can't evaluate (LAMBDA family). Set `self_test=False` only when the solution genuinely can't be evaluated
  in LibreOffice, and double-check those answers in Python.
- Use **8–13 practice tasks** that progress from easy to harder and cover every objective, plus **one bonus problem**
  with 2–5 parts that is clearly harder: multi-step, realistic, combining skills.
- `solution_lang="vba" | "m" | "dax"` renders the solution as a code block in the README.

## Excel vs. LibreOffice pitfalls the verifier can't catch

LibreOffice is only a stand-in for Excel, and a few behaviors differ. Write summaries and checks that are correct in **Excel**:

- **`COUNTA(INDIRECT("'Missing'!A:A"))`** returns `#REF!` in LibreOffice but **1** in Excel, because COUNTA counts the
  error value. To test whether a macro-created sheet exists, use `ISREF(INDIRECT("'Sheet'!A1"))` and wrap the summary:
  `=IF(NOT(ISREF(INDIRECT("'HighAcuity'!A1"))),"",COUNTA(INDIRECT("'HighAcuity'!A:A"))-1)`.
- **COUNT and COUNTA with error arguments** differ between the two apps. Don't pass possibly-erroring expressions to
  them directly.
- In VBA smoke tests, LibreOffice raises error 91 on `Debug.Print` and has no `Worksheet.Sort`, `ListObjects`, or
  `Scripting.Dictionary`. See `xlcourse/vba.py` for the full list.
- **Booleans:** LibreOffice treats TRUE/FALSE as numbers, so `COUNT` counts them and `SUMPRODUCT(range>5)` works without
  `--`. In Excel, `COUNT` ignores logical values and `SUMPRODUCT` treats them as 0. Always coerce with `--` in solutions.
  Also, `=90<A1<130` behaves differently in the two apps.
- **Bare Table names:** LibreOffice reads `tblX` as including the header row, so `ROWS(tblX)` is one higher and
  `FILTER(tblX, …)` errors. Use column references such as `ROWS(tblX[ID])` in summaries.
- **Spill references and arrays:** the `A1#` spill operator returns `#NAME?` in LibreOffice. XLOOKUP with an *array* of
  lookup values silently returns only the first result. TEXTBEFORE/TEXTAFTER don't vectorize over ranges.
- **Number-like text:** COUNTIF/MAXIFS compare number-like text as text in LibreOffice, but Excel coerces it to numbers.
- **`_xlfn.NETWORKDAYS.INTL`** (the correct stored form for Excel) returns `#NAME?` in LibreOffice, so use `live=False`
  or a NETWORKDAYS cross-check.
- **Blank cells in SORT/UNIQUE:** Excel shows empty cells as 0, while LibreOffice leaves them blank.

## Optional Lesson settings

| Attribute / argument | Use it when |
|---|---|
| `lesson.practice_instructions`, `lesson.bonus_instructions` | The generic "Type a formula or value in each yellow cell" line doesn't fit (e.g. macro-fed gray cells) |
| `lesson.practice_how` | You want a different Start Here "2. Practice" text |
| `lesson.key_note` | You want a different subtitle on the hidden key sheets |
| `lesson.bonus_where` | The bonus work happens on another sheet (`""` omits the README line) |
| `lesson.sheet_notes = [(sheet, description)]` | You add sheets in a customize hook and want them listed on Start Here |
| `lesson.verify_scan = [(sheet, "A1:Z99")]` | The verifier should also scan a custom sheet for error values |
| `add_table_sheet(..., freeze=False, hidden=True, hidden_cols=[...])` | The lesson needs unfrozen, hidden sheets or hidden columns |
| `Task(fill=[{...}, {...}])` | The self-test must fill several ranges |
| `Task(self_test="summary")` | A customize hook simulates the work behind a summary task (no `fill`) |
| `Task(key_solution="...")` | The key's "Sample solution" cell should show different text from the README |

The library also strips Markdown from prompts, hints, intros and non-code solutions when it writes them into Excel cells,
escapes `$` in generated README blocks so GitHub doesn't render `$…$` as math, and writes 3-D references
(`Jan:Mar!B2`) as ordinary formulas because Excel rejects them inside array formulas. `build.py` warns about `$` pairs in
hand-written README text and about broken relative links.

## Data sheets

`lesson.add_table_sheet(name, rows, columns=[...], table="tblX", extra_cols=[...], formats={...}, widths={...})`
writes a formatted Excel Table. Rows are dicts from `xlcourse.data.load("encounters")` (typed: dates are `date`,
timestamps `datetime`). Pre-join or pre-compute columns in Python when the lesson isn't about that join. Keep workbooks
reasonably small: a few hundred rows for beginner lessons, and at most ~12k rows for pivot and Data Model lessons.
Pick subsets deterministically (sort, then slice), never randomly without a seed.

## Custom content

`@lesson.customize` registers `fn(wb, lesson, selftest)` that runs after the standard sheets are built. Use it for extra
sheets (Settings, Workspace, Model), charts (`openpyxl.chart`), data validation, conditional formatting, hidden helper
sheets, or reference dashboards inside the hidden key. Write formulas with `lesson.set_formula(ws, "B2", "=...")` so
modern functions get their `_xlfn.` prefixes and dynamic-array semantics. When `selftest` is true you may simulate the
learner's work (for example, converting a range to a Table) so summaries evaluate.

## VBA lessons

Workbooks are `.xlsx`, and learners save them as `.xlsm`. Put starter code in `starter/*.bas` and reference solutions in
`solutions/*.bas`. Begin each `.bas` with `Attribute VB_Name = "ModuleName"` so it imports cleanly through
**File → Import File** in the VBE. Make macros write their results into cells (an `Output` sheet), and have Practice
summaries read those cells so the Check column works. The self-test uses `fill={"range": ..., "values": [...]}` to
simulate the macro's output.

Smoke-test your solutions with `xlcourse.vba.run(...)`. It runs the macros in LibreOffice's VBA-compatibility mode and
reads cells back. It can't run `Scripting.Dictionary`, UserForms, or Outlook; note that in the solution comments when
it applies. Code must also be correct for **real Excel**: use `Option Explicit`, declare every variable, and use correct
object-model members and constants.
