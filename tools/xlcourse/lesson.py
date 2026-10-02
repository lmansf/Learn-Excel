"""Lesson builder: one source of truth for a lesson's workbook and README answer sections.

A lesson module (``tools/lessons/lX_YY_slug.py``) creates a :class:`Lesson`, adds data
sheets, defines :class:`Task` objects for the Practice and Bonus sheets, and returns it.
``python tools/build.py <code>`` then:

1. writes the practice workbook (Start Here, Practice, data sheets, Bonus, and the hidden
   Answer Key / Bonus Key sheets) into the lesson folder,
2. refreshes the generated blocks of the lesson README (task list, collapsed answers,
   bonus, collapsed bonus answers, navigation), and
3. verifies everything in LibreOffice: every live formula in the answer key must equal the
   expected answer, and a "self-test" copy with the sample solutions typed in must show
   "✔ Correct" in every Check cell.
"""
from __future__ import annotations

import math
import re
from dataclasses import dataclass, field
from datetime import date, datetime, time, timedelta
from pathlib import Path
from typing import Any, Callable

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.formula.translate import Translator
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter, quote_sheetname
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.table import Table, TableStyleInfo

from .xlfn import to_file_formula

ROOT = Path(__file__).resolve().parents[2]
REPO_URL = "https://github.com/lmansf/learn-excel"

# ---------------------------------------------------------------------------
# Styles
# ---------------------------------------------------------------------------
NAVY = "1F4E79"
TEAL = "2E75B6"
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
PREFILL_FILL = PatternFill("solid", fgColor="EDEDED")
HEADER_FILL = PatternFill("solid", fgColor=NAVY)
KEY_HEADER_FILL = PatternFill("solid", fgColor="7B2C2C")
BAND_FILL = PatternFill("solid", fgColor="DDEBF7")
GOOD_FILL = PatternFill("solid", fgColor="C6EFCE")
BAD_FILL = PatternFill("solid", fgColor="FFC7CE")
NOTE_FILL = PatternFill("solid", fgColor="F3F3F3")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
INPUT_BORDER = Border(left=Side(style="thin", color="BF9000"), right=Side(style="thin", color="BF9000"),
                      top=Side(style="thin", color="BF9000"), bottom=Side(style="thin", color="BF9000"))
WRAP_TOP = Alignment(wrap_text=True, vertical="top")
CORRECT = "✔ Correct"
NOT_YET = "✘ Not yet"

DATE_FMT = "mm/dd/yyyy"
DATETIME_FMT = "mm/dd/yyyy hh:mm"
MONEY_FMT = "#,##0.00"
MONEY_HINTS = ("Amount", "Charges", "Cost", "Revenue", "Rate", "Paid", "Billed", "Allowed", "Responsibility")


def _sheet_ref(sheet: str, cell: str) -> str:
    return f"{quote_sheetname(sheet)}!{cell}"


def _estimate_lines(text: str, width_chars: float) -> int:
    if not text:
        return 1
    lines = 0
    for para in str(text).split("\n"):
        lines += max(1, math.ceil(len(para) / max(width_chars, 1)))
    return lines


# ---------------------------------------------------------------------------
# Tasks
# ---------------------------------------------------------------------------
@dataclass
class Task:
    """One exercise on the Practice or Bonus sheet.

    prompt       What the learner must do (plain text; shown in the workbook and README).
    answer       Expected value computed in Python from the data (number, text, date, datetime, bool).
                 None means the task is checked manually against the key (``check='manual'``).
    solution     Sample solution. A formula string starting with "=" (Excel syntax exactly as the
                 learner would type it in the Practice answer cell), or Markdown steps for
                 non-formula tasks, or VBA code when ``solution_lang='vba'``.
    explanation  Why the solution works / common pitfalls (Markdown allowed; shown in key + README).
    hint         Nudge shown next to the task (function name, not the answer).
    fmt          Number format for the answer cell and key (e.g. '#,##0.00', '0.0%', 'mm/dd/yyyy').
    check        'auto' (from answer type), 'number', 'percent', 'text', 'date', 'datetime', 'duration',
                 'bool', 'manual' or 'custom'.
    tol          Absolute tolerance for numeric checks (default 0.0051 for floats, 0 for ints).
    accept       Extra accepted answers (numbers or case-insensitive text).
    custom_check Formula fragment that is TRUE when correct; use {cell} for the learner's cell and
                 {key} for the key's expected-value cell. Only with check='custom'.
    summary      Prefilled formula placed in the answer cell instead of a yellow input. Use it when
                 the learner works somewhere else (e.g. fills a new column in a data sheet) and the
                 summary condenses their work into one checkable number.
    fill         Self-test helper for ``summary`` tasks: {'range': "Sheet!N2:N501", 'formula': '=...'}
                 (formula for the first cell; filled down) or {'range': ..., 'values': [...]}.
    live         True: put ``solution`` as a live formula in the key; a string: use that formula
                 instead; False: no live formula (e.g. it refers to Practice-sheet cells).
    self_test    Whether the verifier should type the solution in and expect "✔ Correct".
    title        Short label used in the README answer list.
    answer_display  Override how the answer is printed in README/key (e.g. a small table as text).
    solution_lang   'excel' (default for formulas), 'markdown', 'vba', 'm' (Power Query) or 'dax'.
    """

    prompt: str
    answer: Any = None
    solution: str = ""
    explanation: str = ""
    hint: str = ""
    fmt: str | None = None
    check: str = "auto"
    tol: float | None = None
    accept: list | None = None
    custom_check: str | None = None
    summary: str | None = None
    fill: dict | None = None
    live: bool | str = True
    self_test: bool = True
    title: str = ""
    answer_display: str | None = None
    solution_lang: str | None = None
    table: str | None = None  # table name for [@Col] references in fill formulas

    # filled in during build
    number: str = field(default="", init=False)
    answer_cell: str = field(default="", init=False)
    check_cell: str = field(default="", init=False)
    key_cell: str = field(default="", init=False)
    live_cell: str = field(default="", init=False)

    @property
    def is_formula(self) -> bool:
        return isinstance(self.solution, str) and self.solution.startswith("=")

    @property
    def lang(self) -> str:
        if self.solution_lang:
            return self.solution_lang
        return "excel" if self.is_formula else "markdown"

    def kind(self) -> str:
        if self.check != "auto":
            return self.check
        a = self.answer
        if a is None:
            return "manual"
        if isinstance(a, bool):
            return "bool"
        if isinstance(a, datetime):
            return "datetime"
        if isinstance(a, date):
            return "date"
        if isinstance(a, timedelta):
            return "duration"
        if isinstance(a, (int, float)):
            if self.fmt and "%" in self.fmt:
                return "percent"
            return "number"
        return "text"

    def key_value(self):
        a = self.answer
        if isinstance(a, timedelta):
            return a.total_seconds() / 86400
        return a

    def tolerance(self) -> float:
        if self.tol is not None:
            return self.tol
        k = self.kind()
        if k in ("datetime", "duration"):
            return 1 / 1440 + 1e-9
        if k == "date":
            return 0.0
        if isinstance(self.answer, int) and not isinstance(self.answer, bool):
            return 0.0001
        if k == "percent":
            return 0.00051
        return 0.0051


# ---------------------------------------------------------------------------
# Data sheet handle
# ---------------------------------------------------------------------------
@dataclass
class SheetData:
    name: str
    headers: list[str]
    first_row: int
    last_row: int
    table: str | None

    def col(self, header: str) -> str:
        return get_column_letter(self.headers.index(header) + 1)

    def rng(self, header: str, absolute: bool = True, sheet: bool = True) -> str:
        c = self.col(header)
        r = f"${c}${self.first_row}:${c}${self.last_row}" if absolute else f"{c}{self.first_row}:{c}{self.last_row}"
        return f"{quote_sheetname(self.name)}!{r}" if sheet else r

    def cell(self, header: str, i: int, sheet: bool = True) -> str:
        """Cell for the i-th data row (0-based)."""
        a = f"{self.col(header)}{self.first_row + i}"
        return f"{quote_sheetname(self.name)}!{a}" if sheet else a

    def all_range(self, sheet: bool = True) -> str:
        r = f"$A${self.first_row}:${get_column_letter(len(self.headers))}${self.last_row}"
        return f"{quote_sheetname(self.name)}!{r}" if sheet else r

    @property
    def n(self) -> int:
        return self.last_row - self.first_row + 1


# ---------------------------------------------------------------------------
# Lesson
# ---------------------------------------------------------------------------
class Lesson:
    def __init__(self, code: str, module_dir: str, slug: str, title: str, level: str, minutes: int,
                 objectives: list[str], data_note: str = "", workbook_name: str | None = None):
        self.code = code
        self.module_dir = module_dir
        self.slug = slug
        self.title = title
        self.level = level
        self.minutes = minutes
        self.objectives = objectives
        self.data_note = data_note
        self.dir = ROOT / module_dir / slug
        base = re.sub(r"^\d+-", "", slug)
        self.workbook_name = workbook_name or f"{code}-{base}.xlsx"
        self.tasks: list[Task] = []
        self.bonus: list[Task] = []
        self.practice_intro = ""
        self.bonus_title = "Bonus challenge"
        self.bonus_scenario = ""
        self.start_notes: list[str] = []
        self._data_specs: list[dict] = []
        self._custom: list[Callable] = []
        self.practice_sheet = "Practice"
        self.bonus_sheet = "Bonus"
        self.key_sheet = "Answer Key"
        self.bonus_key_sheet = "Bonus Key"
        self.extra_files: dict[str, bytes | str] = {}
        self.sheet_order: list[str] | None = None
        self.data_sheets: dict[str, SheetData] = {}
        self._dynamic: set[tuple[str, str]] = set()

    # ------------------------------------------------------------------ data
    def add_table_sheet(self, name: str, rows: list[dict], columns: list | None = None, table: str | None = None,
                        formats: dict | None = None, widths: dict | None = None, extra_cols: list[str] | None = None,
                        style: str = "TableStyleMedium2", tab_color: str | None = None, start_row: int = 1,
                        notes: list[str] | None = None, as_table: bool = True) -> SheetData:
        """Add a worksheet holding ``rows`` (list of dicts) as an Excel Table.

        columns     list of column names, or (source_key, header) tuples, to include (default: all keys).
        table       Excel Table name (default: 'tbl' + sheet name without spaces).
        formats     {header: number_format} overrides.
        extra_cols  Empty columns appended for the learner to fill (highlighted yellow header).
        start_row   Header row (use >1 to leave room for notes above the table).
        notes       Lines of text written above the table (requires start_row > len(notes)).
        Returns a SheetData helper with ranges/column letters for building formulas.
        """
        if columns is None:
            columns = list(rows[0].keys()) if rows else []
        cols = [(c, c) if isinstance(c, str) else c for c in columns]
        headers = [h for _, h in cols] + list(extra_cols or [])
        tname = table or ("tbl" + re.sub(r"[^A-Za-z0-9]", "", name)) if as_table else None
        sd = SheetData(name, headers, start_row + 1, start_row + max(len(rows), 1), tname)
        self._data_specs.append(dict(name=name, rows=rows, cols=cols, extra=list(extra_cols or []), table=tname,
                                     formats=formats or {}, widths=widths or {}, style=style, tab_color=tab_color,
                                     start_row=start_row, notes=notes or [], sd=sd))
        self.data_sheets[name] = sd
        return sd

    def customize(self, fn: Callable):
        """Register fn(wb, lesson, selftest: bool) to add charts, extra sheets, validation, etc."""
        self._custom.append(fn)
        return fn

    def set_formula(self, ws, coord: str, formula: str, table: str | None = None, dynamic: bool = True):
        """Write a formula the way Microsoft 365 would: as a dynamic-array formula (no implicit
        intersection), with _xlfn prefixes added. Use this for any formula you place in a custom hook."""
        f = to_file_formula(formula, table)
        if dynamic:
            ws[coord] = ArrayFormula(coord, f)
            self._dynamic.add((ws.title, coord))
        else:
            ws[coord] = f

    # ------------------------------------------------------------------ build
    def build_workbook(self, selftest: bool = False) -> Workbook:
        self._dynamic = set()
        wb = Workbook()
        wb.remove(wb.active)
        start = wb.create_sheet("Start Here")
        practice = wb.create_sheet(self.practice_sheet)
        for spec in self._data_specs:
            self._write_data_sheet(wb, spec)
        bonus = wb.create_sheet(self.bonus_sheet) if self.bonus else None
        key = wb.create_sheet(self.key_sheet)
        bkey = wb.create_sheet(self.bonus_key_sheet) if self.bonus else None

        for i, t in enumerate(self.tasks, 1):
            t.number = str(i)
        for i, t in enumerate(self.bonus, 1):
            t.number = f"B{i}"

        self._write_key(key, self.tasks, f"Answer Key — Lesson {self.code}")
        self._write_tasks(practice, self.tasks, self.key_sheet, f"Lesson {self.code} · Practice", self.practice_intro)
        if self.bonus:
            self._write_key(bkey, self.bonus, f"Bonus Key — Lesson {self.code}")
            self._write_tasks(bonus, self.bonus, self.bonus_key_sheet, f"Lesson {self.code} · {self.bonus_title}",
                              self.bonus_scenario, bonus=True)
        self._write_start(start)

        for fn in self._custom:
            fn(wb, self, selftest)

        if selftest:
            self._apply_selftest(wb)

        key.sheet_state = "hidden"
        if bkey is not None:
            bkey.sheet_state = "hidden"
        for name in ("Start Here", self.practice_sheet, self.bonus_sheet, self.key_sheet, self.bonus_key_sheet):
            if name in wb.sheetnames:
                _fit_width(wb[name], landscape=name != "Start Here")
        if self.sheet_order:
            order = [wb[n] for n in self.sheet_order if n in wb.sheetnames]
            order += [ws for ws in wb.worksheets if ws.title not in self.sheet_order]
            wb._sheets = order
        wb.active = wb.sheetnames.index("Start Here")
        for ws in wb.worksheets:
            ws.sheet_view.tabSelected = ws.title == "Start Here"
        wb.calculation.fullCalcOnLoad = True
        wb.properties.title = f"Lesson {self.code} — {self.title}"
        wb.properties.creator = "Learn Excel · Bluestone Health course"
        return wb

    def _write_data_sheet(self, wb: Workbook, spec: dict):
        ws = wb.create_sheet(spec["name"])
        if spec["tab_color"]:
            ws.sheet_properties.tabColor = spec["tab_color"]
        sr = spec["start_row"]
        for i, line in enumerate(spec["notes"], 1):
            c = ws.cell(row=i, column=1, value=line)
            c.font = Font(bold=(i == 1), size=13 if i == 1 else 10, color=NAVY if i == 1 else "595959", italic=(i > 1))
        cols, extra, rows = spec["cols"], spec["extra"], spec["rows"]
        headers = [h for _, h in cols] + extra
        for j, h in enumerate(headers, 1):
            c = ws.cell(row=sr, column=j, value=h)
            c.font = Font(bold=True, color="FFFFFF" if not spec["table"] else None)
            if not spec["table"]:
                c.fill = HEADER_FILL
            if h in extra:
                c.fill = PatternFill("solid", fgColor="FFD966")
                c.font = Font(bold=True, color="000000")
        widths = [len(str(h)) + 2 for h in headers]
        fmts = {}
        for i, row in enumerate(rows):
            r = sr + 1 + i
            for j, (src, h) in enumerate(cols, 1):
                v = row.get(src) if isinstance(row, dict) else row[j - 1]
                c = ws.cell(row=r, column=j, value=v)
                if h in spec["formats"]:
                    c.number_format = spec["formats"][h]
                elif isinstance(v, datetime):
                    c.number_format = DATETIME_FMT if (v.hour or v.minute) or "Date" not in h or "Time" in h else DATE_FMT
                    if "Time" in h or "Start" in h or "End" in h or "Clock" in h:
                        c.number_format = DATETIME_FMT
                elif isinstance(v, date):
                    c.number_format = DATE_FMT
                elif isinstance(v, float) and any(k in h for k in MONEY_HINTS):
                    c.number_format = MONEY_FMT
                if i < 400:
                    s = v.strftime("%m/%d/%Y %H:%M") if isinstance(v, datetime) else str(v) if v is not None else ""
                    widths[j - 1] = max(widths[j - 1], min(len(s) + 2, 48))
            for k, h in enumerate(extra):
                c = ws.cell(row=r, column=len(cols) + k + 1)
                if h in spec["formats"]:
                    c.number_format = spec["formats"][h]
                c.fill = INPUT_FILL
        for j, h in enumerate(headers, 1):
            ws.column_dimensions[get_column_letter(j)].width = spec["widths"].get(h, max(widths[j - 1], 9))
        last = sr + max(len(rows), 1)
        if spec["table"]:
            ref = f"A{sr}:{get_column_letter(len(headers))}{last}"
            tbl = Table(displayName=spec["table"], ref=ref)
            tbl.tableStyleInfo = TableStyleInfo(name=spec["style"], showRowStripes=True)
            ws.add_table(tbl)
        ws.freeze_panes = ws.cell(row=sr + 1, column=1)

    def _write_tasks(self, ws, tasks: list[Task], key_sheet: str, title: str, intro: str, bonus: bool = False):
        ws.sheet_properties.tabColor = "BF9000" if bonus else "548235"
        ws["A1"] = title
        ws["A1"].font = Font(bold=True, size=16, color=NAVY)
        row = 2
        intro_lines = [intro] if intro else []
        intro_lines.append("Type a formula or value in each yellow cell. The Check column turns green when your answer matches. "
                           f"Stuck? Read the hint, then the lesson guide. Answers: right-click a sheet tab → Unhide… → '{key_sheet}'.")
        for line in intro_lines:
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
            c = ws.cell(row=row, column=1, value=line)
            c.alignment = WRAP_TOP
            c.font = Font(italic=line != intro, color="404040", size=11)
            ws.row_dimensions[row].height = 15 * _estimate_lines(line, 150) + 4
            row += 1
        score_row = row
        row += 1
        hdr = row
        for j, h in enumerate(["#", "Task", "Hint", "Your Answer", "Check"], 1):
            c = ws.cell(row=hdr, column=j, value=h)
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = HEADER_FILL
            c.alignment = Alignment(vertical="center", wrap_text=True)
            c.border = BOX
        first = hdr + 1
        for i, t in enumerate(tasks):
            r = first + i
            kr = 5 + i
            t.answer_cell = f"D{r}"
            t.check_cell = f"E{r}"
            t._sheet = ws.title
            t.key_cell = _sheet_ref(key_sheet, f"$C${kr}")
            ws.cell(row=r, column=1, value=t.number).alignment = Alignment(horizontal="center", vertical="top")
            p = ws.cell(row=r, column=2, value=t.prompt)
            p.alignment = WRAP_TOP
            h = ws.cell(row=r, column=3, value=t.hint or "")
            h.alignment = WRAP_TOP
            h.font = Font(italic=True, color="7F7F7F")
            a = ws.cell(row=r, column=4)
            if t.summary:
                self.set_formula(ws, f"D{r}", t.summary, t.table)
                a = ws.cell(row=r, column=4)
                a.fill = PREFILL_FILL
                a.font = Font(italic=True, color="404040")
            else:
                a.fill = INPUT_FILL
                a.border = INPUT_BORDER
            if t.fmt:
                a.number_format = t.fmt
            a.alignment = Alignment(vertical="top", horizontal="right" if t.kind() not in ("text", "manual") else "left")
            ck = ws.cell(row=r, column=5, value=self._check_formula(t, f"D{r}", t.key_cell))
            ck.font = Font(bold=True)
            ck.alignment = Alignment(vertical="top")
            for col in range(1, 6):
                ws.cell(row=r, column=col).border = INPUT_BORDER if (col == 4 and not t.summary) else BOX
            ws.row_dimensions[r].height = max(30, 15 * max(_estimate_lines(t.prompt, 78), _estimate_lines(t.hint, 40)) + 6)
        last = first + len(tasks) - 1
        auto = [t for t in tasks if t.kind() != "manual"]
        sc = ws.cell(row=score_row, column=1,
                     value=f'="Score: "&COUNTIF(E{first}:E{last},"✔*")&" of {len(auto)} auto-checked tasks correct"')
        ws.merge_cells(start_row=score_row, start_column=1, end_row=score_row, end_column=3)
        sc.font = Font(bold=True, size=12, color="375623")
        rng = f"E{first}:E{last}"
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT(E{first},1)="✔"'], fill=GOOD_FILL, font=Font(bold=True, color="006100")))
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT(E{first},1)="✘"'], fill=BAD_FILL, font=Font(bold=True, color="9C0006")))
        for col, wdt in zip("ABCDE", (5, 72, 38, 22, 14)):
            ws.column_dimensions[col].width = wdt
        ws.freeze_panes = ws.cell(row=first, column=1)

    def _check_formula(self, t: Task, cell: str, key: str) -> str:
        k = t.kind()
        if k == "manual":
            return "See key"
        tol = t.tolerance()
        if k == "number":
            conds = [f"ABS({cell}-{key})<={tol}"]
            conds += [f"ABS({cell}-{a})<={tol}" for a in (t.accept or [])]
            cond = f"AND(ISNUMBER({cell}),OR({','.join(conds)}))"
        elif k == "percent":
            cond = f"AND(ISNUMBER({cell}),OR(ABS({cell}-{key})<={tol},ABS({cell}-{key}*100)<={tol * 100}))"
        elif k == "date":
            cond = f"AND(ISNUMBER({cell}),INT({cell})=INT({key}))"
        elif k in ("datetime", "duration"):
            cond = f"AND(ISNUMBER({cell}),ABS({cell}-{key})<={tol})"
        elif k == "bool":
            cond = f"{cell}={key}"
        elif k == "text":
            alts = [f'LOWER(TRIM({cell}&""))=LOWER(TRIM({key}&""))']
            alts += ['LOWER(TRIM({0}&""))="{1}"'.format(cell, str(a).strip().lower().replace('"', '""')) for a in (t.accept or [])]
            cond = f"OR({','.join(alts)})"
        elif k == "custom":
            cond = t.custom_check.replace("{cell}", cell).replace("{key}", key)
        else:
            raise ValueError(f"unknown check kind {k}")
        return to_file_formula(f'=IF(IFERROR({cell}="",FALSE),"",IF(IFERROR({cond},FALSE),"{CORRECT}","{NOT_YET}"))')

    def _write_key(self, ws, tasks: list[Task], title: str):
        ws.sheet_properties.tabColor = "C00000"
        ws["A1"] = "🔑 " + title
        ws["A1"].font = Font(bold=True, size=16, color="7B2C2C")
        ws["A2"] = ("Spoiler alert — try every task before reading this sheet. Column C is the value the Check column compares "
                    "against. Column D is a sample solution; column E is that same formula, live, proving it works.")
        ws["A2"].font = Font(italic=True, color="595959")
        ws.merge_cells("A2:F2")
        ws["A2"].alignment = WRAP_TOP
        ws.row_dimensions[2].height = 32
        for j, h in enumerate(["#", "Task", "Answer", "Sample solution", "Live result", "Explanation"], 1):
            c = ws.cell(row=4, column=j, value=h)
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = KEY_HEADER_FILL
            c.border = BOX
        for i, t in enumerate(tasks):
            r = 5 + i
            ws.cell(row=r, column=1, value=t.number).alignment = Alignment(horizontal="center", vertical="top")
            ws.cell(row=r, column=2, value=t.prompt).alignment = WRAP_TOP
            kv = t.key_value()
            c = ws.cell(row=r, column=3)
            if t.answer_display is not None and t.answer is None:
                c.value = t.answer_display
            else:
                c.value = kv
            if t.fmt:
                c.number_format = t.fmt
            elif isinstance(kv, datetime):
                c.number_format = DATETIME_FMT
            elif isinstance(kv, date):
                c.number_format = DATE_FMT
            c.alignment = Alignment(vertical="top", wrap_text=True)
            c.font = Font(bold=True)
            s = ws.cell(row=r, column=4)
            s.value = t.solution
            s.data_type = "s"
            s.alignment = WRAP_TOP
            s.font = Font(name="Consolas", size=10)
            s._style.quotePrefix = 1
            live = t.live if isinstance(t.live, str) else (t.solution if (t.live and t.is_formula) else None)
            t.live_cell = ""
            if live:
                self.set_formula(ws, f"E{r}", live, t.table)
                lc = ws.cell(row=r, column=5)
                if t.fmt:
                    lc.number_format = t.fmt
                elif isinstance(kv, datetime):
                    lc.number_format = DATETIME_FMT
                elif isinstance(kv, date):
                    lc.number_format = DATE_FMT
                lc.alignment = Alignment(vertical="top")
                t.live_cell = f"E{r}"
            else:
                ws.cell(row=r, column=5, value="—").alignment = Alignment(horizontal="center", vertical="top")
            ex = ws.cell(row=r, column=6, value=_plain(t.explanation))
            ex.alignment = WRAP_TOP
            for col in range(1, 7):
                ws.cell(row=r, column=col).border = BOX
            ws.row_dimensions[r].height = min(409, max(30, 15 * max(_estimate_lines(t.prompt, 55), _estimate_lines(_plain(t.explanation), 60),
                                                                     _estimate_lines(t.solution, 45)) + 6))
        for col, wdt in zip("ABCDEF", (5, 55, 18, 48, 16, 62)):
            ws.column_dimensions[col].width = wdt
        ws.freeze_panes = "A5"

    def _write_start(self, ws):
        ws.sheet_properties.tabColor = NAVY
        ws.column_dimensions["A"].width = 3
        ws.column_dimensions["B"].width = 26
        ws.column_dimensions["C"].width = 95
        ws["B2"] = f"Lesson {self.code}"
        ws["B2"].font = Font(bold=True, size=12, color=TEAL)
        ws["B3"] = self.title
        ws["B3"].font = Font(bold=True, size=22, color=NAVY)
        ws["B4"] = f"{self.level}  ·  about {self.minutes} minutes  ·  Bluestone Health System (fictional)"
        ws["B4"].font = Font(italic=True, color="595959")
        r = 6

        def section(title):
            nonlocal r
            c = ws.cell(row=r, column=2, value=title)
            c.font = Font(bold=True, size=13, color="FFFFFF")
            for col in (2, 3):
                ws.cell(row=r, column=col).fill = HEADER_FILL
            r += 1

        def line(label, text, fill=None, bold_label=True):
            nonlocal r
            a = ws.cell(row=r, column=2, value=label)
            a.font = Font(bold=bold_label)
            a.alignment = WRAP_TOP
            b = ws.cell(row=r, column=3, value=text)
            b.alignment = WRAP_TOP
            if fill:
                a.fill = fill
            ws.row_dimensions[r].height = 15 * _estimate_lines(text, 100) + 3
            r += 1

        section("What you'll learn")
        for o in self.objectives:
            line("•", o, bold_label=False)
        r += 1
        section("How to use this workbook")
        line("1. Read the guide", f"The lesson guide (README) explains each skill with examples: {REPO_URL}/tree/main/{self.module_dir}/{self.slug}")
        line("2. Practice", f"Go to the '{self.practice_sheet}' sheet. Type a formula or value into each yellow cell.")
        line("3. Check yourself", "The Check column shows ✔ Correct (green) or ✘ Not yet (red). Tasks marked 'See key' are checked by comparing with the answer key.")
        if self.bonus:
            line("4. Bonus", f"Finished? Try the harder '{self.bonus_sheet}' sheet.")
        line("5. Answers", f"The '{self.key_sheet}' and '{self.bonus_key_sheet}' sheets are hidden. Right-click any sheet tab → Unhide… → pick one → OK. (Try first!)")
        for note in self.start_notes:
            line("Note", note)
        r += 1
        section("Color legend")
        line("Yellow cell", "Your answer goes here.", fill=INPUT_FILL)
        line("Gray cell", "Pre-filled summary formula — it reads the work you did on another sheet.", fill=PREFILL_FILL)
        line("✔ Correct", "Your answer matches the key.", fill=GOOD_FILL)
        line("✘ Not yet", "Not matching yet — check the hint and try again.", fill=BAD_FILL)
        r += 1
        section("Sheets in this workbook")
        for spec in self._data_specs:
            line(spec["name"], f"Data: {len(spec['rows']):,} rows" + (f" — Excel Table '{spec['table']}'" if spec["table"] else ""))
        if self.data_note:
            line("About the data", self.data_note)
        line("Disclaimer", "All people, places, and numbers are synthetic and fictional. Not clinical guidance.")

    def _apply_selftest(self, wb: Workbook):
        """Type each sample solution into its answer cell so the verifier can expect ✔ everywhere."""
        for t in self.tasks + self.bonus:
            if not t.self_test or t.kind() == "manual":
                continue
            ws = wb[t._sheet]
            if t.fill:
                sheet, rng = t.fill["range"].split("!")
                sheet = sheet.strip("'")
                target = wb[sheet]
                cells = [c for row in target[rng.replace("$", "")] for c in row]
                if "values" in t.fill:
                    for c, v in zip(cells, t.fill["values"]):
                        c.value = v
                else:
                    origin = cells[0].coordinate
                    f0 = to_file_formula(t.fill["formula"], t.fill.get("table", t.table))
                    for c in cells:
                        fc = Translator(f0, origin=origin).translate_formula(c.coordinate) if c.coordinate != origin else f0
                        if t.fill.get("array"):
                            target[c.coordinate] = ArrayFormula(c.coordinate, fc)
                            self._dynamic.add((target.title, c.coordinate))
                        else:
                            c.value = fc
            elif t.summary:
                continue
            elif t.is_formula:
                self.set_formula(ws, t.answer_cell, t.solution, t.table)
            elif t.answer is not None:
                v = t.key_value()
                ws[t.answer_cell] = v

    def save(self, path: Path, selftest: bool = False) -> Path:
        wb = self.build_workbook(selftest=selftest)
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        wb.save(path)
        sheet_files = {ws.title: f"xl/worksheets/sheet{i}.xml" for i, ws in enumerate(wb.worksheets, 1)}
        _add_dynamic_array_metadata(path, {(sheet_files[s], c) for s, c in self._dynamic})
        return path

    def write_extra_files(self):
        for rel, content in self.extra_files.items():
            p = self.dir / rel
            p.parent.mkdir(parents=True, exist_ok=True)
            if isinstance(content, bytes):
                p.write_bytes(content)
            else:
                p.write_text(content, encoding="utf-8", newline="\n")

    # ------------------------------------------------------------------ README
    def readme_blocks(self, nav: str = "") -> dict[str, str]:
        return {
            "practice": _md_tasks(self.tasks, self.practice_intro),
            "answers": _md_answers(self.tasks, "🔑 Show the answer key", "Try every task before opening this."),
            "bonus": _md_bonus(self),
            "bonus-answers": _md_answers(self.bonus, "🔑 Show the bonus solution", "Give it a real try first!") if self.bonus else "",
            "nav": nav,
        }


def _fit_width(ws, landscape: bool = True):
    """Print/PDF setup: fit all columns on one page width."""
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


# ---------------------------------------------------------------------------
# Markdown helpers
# ---------------------------------------------------------------------------
def _plain(md: str) -> str:
    """Strip a little Markdown for display inside Excel cells."""
    s = re.sub(r"```[a-zA-Z]*\n?", "", md or "")
    s = s.replace("**", "").replace("`", "")
    return s.strip()


def fmt_value(v, fmt: str | None = None) -> str:
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, timedelta):
        mins = round(v.total_seconds() / 60)
        return f"{mins // 60}:{mins % 60:02d} (h:mm)"
    if isinstance(v, datetime):
        return v.strftime("%m/%d/%Y %H:%M")
    if isinstance(v, date):
        return v.strftime("%m/%d/%Y")
    if isinstance(v, (int, float)):
        f = fmt or ""
        if "%" in f:
            dec = len(f.split(".")[1].rstrip("%")) if "." in f else 0
            return f"{v * 100:,.{dec}f}%"
        if f.startswith("$") or "$" in f:
            dec = len(f.split(".")[1].split(";")[0].rstrip(")_ ")) if "." in f else 0
            return f"${v:,.{dec}f}"
        if "." in f:
            dec = len(re.match(r"[0#,]*\.([0#]+)", f.replace("$", "")).group(1)) if re.match(r"[0#,]*\.([0#]+)", f.replace("$", "")) else 2
            return f"{v:,.{dec}f}"
        if isinstance(v, int) or float(v).is_integer():
            return f"{int(round(v)):,}"
        return f"{v:,.4f}".rstrip("0").rstrip(".")
    return str(v)


def _cell(s: str) -> str:
    return (s or "").replace("|", "\\|").replace("\n", "<br>")


def _md_tasks(tasks: list[Task], intro: str) -> str:
    lines = []
    if intro:
        lines += [intro, ""]
    lines += ["| # | Task | Hint |", "|:-:|------|------|"]
    for t in tasks:
        lines.append(f"| {t.number} | {_cell(t.prompt)} | {_cell(t.hint)} |")
    return "\n".join(lines)


def _code_block(t: Task) -> str:
    sol = t.solution or ""
    lang = t.lang
    if lang == "excel":
        if "\n" in sol or len(sol) > 90:
            return f"```\n{sol}\n```"
        return f"`{sol}`"
    if lang in ("vba", "m", "dax"):
        fence = {"vba": "vba", "m": "powerquery", "dax": "dax"}[lang]
        return f"```{fence}\n{sol}\n```"
    return sol


def _md_answers(tasks: list[Task], summary: str, warn: str) -> str:
    out = ["<details>", f"<summary><b>{summary}</b> — {warn}</summary>", ""]
    for t in tasks:
        label = t.title or (t.prompt if len(t.prompt) <= 90 else t.prompt[:87].rsplit(" ", 1)[0] + "…")
        out.append(f"**{t.number}. {label}**")
        out.append("")
        if t.answer is not None or t.answer_display:
            shown = t.answer_display if t.answer_display is not None else fmt_value(t.answer, t.fmt)
            if "\n" in shown:
                out += ["- **Answer:**", "", shown, ""]
            else:
                out.append(f"- **Answer:** {shown}")
        sol = _code_block(t)
        if sol:
            if sol.startswith("```") or "\n" in sol:
                out += ["- **Solution:**", "", sol, ""]
            else:
                out.append(f"- **Solution:** {sol}")
        if t.explanation:
            out += ["", t.explanation.strip()]
        out.append("")
    out.append("</details>")
    return "\n".join(out)


def _md_bonus(lesson: Lesson) -> str:
    if not lesson.bonus:
        return ""
    lines = []
    if lesson.bonus_scenario:
        lines += [lesson.bonus_scenario, ""]
    lines.append(f"Work on the **{lesson.bonus_sheet}** sheet of the workbook.")
    lines.append("")
    for t in lesson.bonus:
        lines.append(f"- **{t.number}.** {t.prompt}" + (f" *(Hint: {t.hint})*" if t.hint else ""))
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Dynamic-array metadata (so Excel 365 treats formulas exactly as if typed: no implicit @)
# ---------------------------------------------------------------------------
_METADATA_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    'xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">'
    '<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" '
    'pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" '
    'coerce="1" cellMeta="1"/></metadataTypes>'
    '<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}">'
    '<xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>'
    '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>'
)


def _add_dynamic_array_metadata(path: Path, cells: set[tuple[str, str]]):
    """Mark array formulas as dynamic-array formulas (cm="1"), mirroring what Excel 365 writes."""
    import shutil
    import tempfile
    import zipfile

    if not cells:
        return
    by_part: dict[str, set[str]] = {}
    for part, coord in cells:
        by_part.setdefault(part, set()).add(coord)
    tmp = Path(tempfile.mkstemp(suffix=".xlsx")[1])
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in by_part:
                xml = data.decode("utf-8")
                for coord in by_part[item.filename]:
                    xml = re.sub(rf'<c r="{coord}"(?![^>]*\bcm=)', f'<c r="{coord}" cm="1"', xml, count=1)
                data = xml.encode("utf-8")
            elif item.filename == "[Content_Types].xml":
                xml = data.decode("utf-8")
                if "/xl/metadata.xml" not in xml:
                    xml = xml.replace("</Types>", '<Override PartName="/xl/metadata.xml" ContentType='
                                      '"application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/></Types>')
                data = xml.encode("utf-8")
            elif item.filename == "xl/_rels/workbook.xml.rels":
                xml = data.decode("utf-8")
                if "metadata.xml" not in xml:
                    xml = xml.replace("</Relationships>", '<Relationship Id="rIdXlcMeta" Type="http://schemas.openxmlformats.org/'
                                      'officeDocument/2006/relationships/sheetMetadata" Target="metadata.xml"/></Relationships>')
                data = xml.encode("utf-8")
            zout.writestr(item, data)
        if "xl/metadata.xml" not in zin.namelist():
            zout.writestr("xl/metadata.xml", _METADATA_XML)
    shutil.move(str(tmp), str(path))
