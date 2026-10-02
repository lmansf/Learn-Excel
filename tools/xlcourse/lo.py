"""Evaluate .xlsx workbooks with LibreOffice (headless, via UNO) to verify formulas.

Requires LibreOffice Calc 24.8+ (26.x recommended for XLOOKUP/FILTER/LET support) and the
Python UNO bridge (``import uno``). Each call starts a private soffice process with its own
profile directory and pipe name, so many verifications can run in parallel safely.
"""
from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
import time
import uuid
from contextlib import contextmanager
from pathlib import Path


@contextmanager
def office():
    import uno  # noqa: F401  (system python must have the UNO bridge)
    from com.sun.star.beans import PropertyValue  # noqa: F401

    name = f"xlc_{os.getpid()}_{uuid.uuid4().hex[:8]}"
    prof = Path(tempfile.mkdtemp(prefix="lo_profile_"))
    cmd = ["soffice", "--headless", "--invisible", "--nologo", "--norestore", "--nodefault", "--nolockcheck",
           f"-env:UserInstallation=file://{prof}", f"--accept=pipe,name={name};urp;StarOffice.ComponentContext"]
    proc = subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
    ctx = None
    for _ in range(240):
        try:
            ctx = resolver.resolve(f"uno:pipe,name={name};urp;StarOffice.ComponentContext")
            break
        except Exception:
            if proc.poll() is not None:
                raise RuntimeError("soffice exited early")
            time.sleep(0.25)
    if ctx is None:
        proc.kill()
        raise RuntimeError("could not connect to LibreOffice")
    desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    try:
        yield desktop
    finally:
        try:
            desktop.terminate()
        except Exception:
            pass
        try:
            proc.wait(timeout=20)
        except Exception:
            proc.kill()
        shutil.rmtree(prof, ignore_errors=True)


def _pv(name, value):
    from com.sun.star.beans import PropertyValue
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def read_cells(desktop, path: Path, cells: list[tuple[str, str]], scan: list[tuple[str, str]] | None = None) -> dict:
    """Load a workbook, recalculate everything, and return values for (sheet, address) pairs.

    Returns {"cells": {(sheet, addr): {"kind": "value"|"text"|"empty"|"error", "value": float|None, "text": str}},
             "errors": [(sheet, addr, text)] for any error cells found inside the `scan` ranges}
    """
    import uno
    url = uno.systemPathToFileUrl(str(Path(path).resolve()))
    doc = desktop.loadComponentFromURL(url, "_blank", 0, (_pv("Hidden", True), _pv("ReadOnly", True)))
    if doc is None:
        raise RuntimeError(f"LibreOffice could not open {path}")
    try:
        doc.calculateAll()
        sheets = doc.Sheets
        out = {}
        for sheet, addr in cells:
            if not sheets.hasByName(sheet):
                out[(sheet, addr)] = {"kind": "missing-sheet", "value": None, "text": ""}
                continue
            c = sheets.getByName(sheet).getCellRangeByName(addr)
            out[(sheet, addr)] = _cell_info(c)
        errors = []
        for sheet, rng in scan or []:
            if not sheets.hasByName(sheet):
                continue
            r = sheets.getByName(sheet).getCellRangeByName(rng)
            addr = r.getRangeAddress()
            for row in range(addr.StartRow, addr.EndRow + 1):
                for col in range(addr.StartColumn, addr.EndColumn + 1):
                    c = sheets.getByName(sheet).getCellByPosition(col, row)
                    if c.getType().value == "FORMULA" and c.getError() != 0:
                        errors.append((sheet, _a1(col, row), c.getString() or f"Err:{c.getError()}"))
        return {"cells": out, "errors": errors}
    finally:
        doc.close(True)


def _a1(col, row):
    s, n = "", col + 1
    while n:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return f"{s}{row + 1}"


def _cell_info(c):
    t = c.getType().value  # EMPTY, VALUE, TEXT, FORMULA
    if t == "EMPTY":
        return {"kind": "empty", "value": None, "text": ""}
    if t == "VALUE":
        return {"kind": "value", "value": c.getValue(), "text": c.getString()}
    if t == "TEXT":
        return {"kind": "text", "value": None, "text": c.getString()}
    if c.getError() != 0:
        return {"kind": "error", "value": None, "text": c.getString() or f"Err:{c.getError()}"}
    # formula: result type 1 = value, 2 = string (FormulaResultType2)
    rt = c.FormulaResultType2
    if rt == 2:
        s = c.getString()
        return {"kind": "text" if s != "" else "empty", "value": None, "text": s}
    return {"kind": "value", "value": c.getValue(), "text": c.getString()}
