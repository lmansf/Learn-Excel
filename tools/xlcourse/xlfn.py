"""Convert human-friendly Excel formulas into the form stored inside .xlsx files.

Excel stores functions introduced after Excel 2007 with a ``_xlfn.`` prefix (and a few
with ``_xlfn._xlws.``). LET/LAMBDA parameter names are stored with ``_xlpm.``. Formulas
written by openpyxl without these prefixes show ``#NAME?`` until re-entered, so every
formula the course writes into a workbook goes through :func:`to_file_formula`.

``[@Column]`` (this-row structured references) is converted to the stored form
``Table[[#This Row],[Column]]`` when the table name is known.
"""
from __future__ import annotations

import re

XLWS = {"FILTER", "SORT"}
XLFN = {
    "ACOT", "ACOTH", "AGGREGATE", "ARABIC", "ARRAYTOTEXT", "BASE", "BETA.DIST", "BETA.INV", "BINOM.DIST", "BINOM.DIST.RANGE",
    "BINOM.INV", "BITAND", "BITLSHIFT", "BITOR", "BITRSHIFT", "BITXOR", "BYCOL", "BYROW", "CEILING.MATH", "CEILING.PRECISE",
    "CHISQ.DIST", "CHISQ.DIST.RT", "CHISQ.INV", "CHISQ.INV.RT", "CHISQ.TEST", "CHOOSECOLS", "CHOOSEROWS", "COMBINA", "CONCAT",
    "CONFIDENCE.NORM", "CONFIDENCE.T", "COT", "COTH", "COVARIANCE.P", "COVARIANCE.S", "CSC", "CSCH", "DAYS", "DECIMAL", "DROP",
    "ENCODEURL", "ERF.PRECISE", "ERFC.PRECISE", "EXPAND", "EXPON.DIST", "F.DIST", "F.DIST.RT", "F.INV", "F.INV.RT", "F.TEST",
    "FILTERXML", "FLOOR.MATH", "FLOOR.PRECISE", "FORECAST.ETS", "FORECAST.ETS.CONFINT", "FORECAST.ETS.SEASONALITY",
    "FORECAST.ETS.STAT", "FORECAST.LINEAR", "FORMULATEXT", "GAMMA", "GAMMA.DIST", "GAMMA.INV", "GAMMALN.PRECISE", "GAUSS",
    "GROUPBY", "HSTACK", "HYPGEOM.DIST", "IFNA", "IFS", "IMAGE", "ISFORMULA", "ISOMITTED", "ISOWEEKNUM", "LAMBDA", "LET",
    "LOGNORM.DIST", "LOGNORM.INV", "MAKEARRAY", "MAP", "MAXIFS", "MINIFS", "MODE.MULT", "MODE.SNGL", "MUNIT", "NEGBINOM.DIST",
    "NETWORKDAYS.INTL", "NORM.DIST", "NORM.INV", "NORM.S.DIST", "NORM.S.INV", "NUMBERVALUE", "PDURATION", "PERCENTILE.EXC",
    "PERCENTILE.INC", "PERCENTOF", "PERCENTRANK.EXC", "PERCENTRANK.INC", "PERMUTATIONA", "PHI", "PIVOTBY", "POISSON.DIST",
    "QUARTILE.EXC", "QUARTILE.INC", "RANDARRAY", "RANK.AVG", "RANK.EQ", "REDUCE", "REGEXEXTRACT", "REGEXREPLACE", "REGEXTEST",
    "RRI", "SCAN", "SEC", "SECH", "SEQUENCE", "SHEET", "SHEETS", "SKEW.P", "SORTBY", "STDEV.P", "STDEV.S", "SWITCH", "T.DIST",
    "T.DIST.2T", "T.DIST.RT", "T.INV", "T.INV.2T", "T.TEST", "TAKE", "TEXTAFTER", "TEXTBEFORE", "TEXTJOIN", "TEXTSPLIT",
    "TOCOL", "TOROW", "TRIMRANGE", "UNICHAR", "UNICODE", "UNIQUE", "VALUETOTEXT", "VAR.P", "VAR.S", "VSTACK", "WEBSERVICE",
    "WEIBULL.DIST", "WORKDAY.INTL", "WRAPCOLS", "WRAPROWS", "XLOOKUP", "XMATCH", "XOR", "Z.TEST", "ANCHORARRAY", "SINGLE",
}

_FUNC_RE = re.compile(r"(?<![A-Za-z0-9_.\]])([A-Za-z][A-Za-z0-9.]*)\s*\(")
_THIS_ROW_RE = re.compile(r"(?<![\]A-Za-z0-9_])(?:([A-Za-z_][A-Za-z0-9_.]*))?\[@(\[[^\]]+\]|[^\[\]]+)\]")


def _split_strings(formula: str):
    """Yield (is_string_literal, text) segments so we never rewrite inside quotes."""
    out, buf, i, n = [], [], 0, len(formula)
    while i < n:
        ch = formula[i]
        if ch == '"':
            if buf:
                out.append((False, "".join(buf)))
                buf = []
            j = i + 1
            while j < n:
                if formula[j] == '"' and j + 1 < n and formula[j + 1] == '"':
                    j += 2
                    continue
                if formula[j] == '"':
                    break
                j += 1
            out.append((True, formula[i:j + 1]))
            i = j + 1
        else:
            buf.append(ch)
            i += 1
    if buf:
        out.append((False, "".join(buf)))
    return out


def _prefix_functions(code: str) -> str:
    def repl(m):
        name = m.group(1)
        up = name.upper()
        if up.startswith("_XL"):
            return m.group(0)
        if up in XLWS:
            return f"_xlfn._xlws.{up}("
        if up in XLFN:
            return f"_xlfn.{up}("
        return m.group(0)
    return _FUNC_RE.sub(repl, code)


def _lambda_params(code: str) -> list[str]:
    """Collect LET variable names and LAMBDA parameter names (best effort parser)."""
    names = []
    for m in re.finditer(r"\b(LET|LAMBDA)\s*\(", code, flags=re.I):
        kind = m.group(1).upper()
        depth, i, args, cur = 0, m.end(), [], []
        while i < len(code):
            ch = code[i]
            if ch in "([{":
                depth += 1
            elif ch in ")]}":
                if depth == 0:
                    args.append("".join(cur).strip())
                    break
                depth -= 1
            elif ch == "," and depth == 0:
                args.append("".join(cur).strip())
                cur = []
                i += 1
                continue
            cur.append(ch)
            i += 1
        if kind == "LET":
            names += [a for k, a in enumerate(args[:-1]) if k % 2 == 0]
        else:
            names += args[:-1]
    return [n for n in names if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_.]*", n)]


def to_file_formula(formula: str, table: str | None = None) -> str:
    """Return the formula as Excel stores it in the file (prefixes, this-row references)."""
    if not formula.startswith("="):
        return formula
    segments = _split_strings(formula[1:])
    code_only = "".join(t for s, t in segments if not s)
    params = _lambda_params(code_only)
    out = []
    for is_str, text in segments:
        if is_str:
            out.append(text)
            continue
        def this_row(m):
            tbl = m.group(1) or table
            col = m.group(2)
            if col.startswith("["):
                col = col[1:-1]
            if not tbl:
                raise ValueError(f"[@{col}] needs a table name in: {formula}")
            return f"{tbl}[[#This Row],[{col}]]"
        text = _THIS_ROW_RE.sub(this_row, text)
        text = _prefix_functions(text)
        for p in sorted(set(params), key=len, reverse=True):
            text = re.sub(rf"(?<![A-Za-z0-9_.\]\[!$]){re.escape(p)}(?![A-Za-z0-9_(!\[])", f"_xlpm.{p}", text, flags=re.I)
        out.append(text)
    return "=" + "".join(out)


def strip_file_formula(formula: str) -> str:
    """Inverse of to_file_formula for display purposes."""
    return re.sub(r"_xlfn\.(_xlws\.)?|_xlpm\.", "", formula)
