"""Best-effort smoke tests for VBA solutions using LibreOffice's VBA compatibility mode.

    from xlcourse import vba
    res = vba.run(
        workbook="05-automation-vba/02-vba-fundamentals/5.2-vba-fundamentals.xlsx",
        modules=["05-automation-vba/02-vba-fundamentals/solutions/Module1.bas"],
        macros=["CountCriticals", "SummarizeLabs"],
        read=[("Output", "B2"), ("Output", "B3")],
    )
    print(res["cells"], res["errors"], res["sheets"])

What works in LibreOffice (tested): Worksheets/ThisWorkbook/Range/Cells/Offset/Resize/CurrentRegion,
End(xlUp) last-row idiom, Worksheets.Add/Name, loops, Select Case, MsgBox suppressed, InStr/Left/Format,
Application.WorksheetFunction.Sum/CountIf/…, user-defined functions called from cells.
What does NOT work: CreateObject("Scripting.Dictionary") and other Windows COM objects, UserForms, Outlook,
ListObjects, the Worksheet.Sort object (use Range.Sort in a smoke-test copy), Worksheet.AutoFilterMode = False,
multi-area Range.Copy, Worksheet.Copy into a new workbook, ExportAsFixedFormat, and Debug.Print (raises error 91 in
LibreOffice even though the rest of the macro ran — such errors are annotated in the result). A failure here is a hint to double-check the code,
not proof that it fails in Excel — and a pass is strong (not absolute) evidence it works in Excel.
"""
from __future__ import annotations

import re
from pathlib import Path

from . import lo

ROOT = Path(__file__).resolve().parents[2]


def _clean_module(text: str) -> str:
    lines = []
    for line in text.splitlines():
        s = line.strip()
        if s.startswith("Attribute VB_") or s.upper().startswith("VERSION ") or s.upper().startswith("OPTION VBASUPPORT"):
            continue
        lines.append(line)
    body = "\n".join(lines)
    return "Option VBASupport 1\n" + body


def run(workbook: str | Path, modules: list[str | Path], macros: list[str], read: list[tuple[str, str]] = (),
        save_as: str | Path | None = None, extra_code: str = "") -> dict:
    """Run macros against a copy of `workbook` in LibreOffice and read back cells.

    modules  .bas file paths (Attribute lines are stripped) or raw code strings.
    macros   Sub names to run, in order. Each runs inside an error trap; errors are reported, not raised.
    read     (sheet, address) cells to return after all macros have run (after a recalculation).
    save_as  optional path to save the resulting workbook (xlsx) for inspection.
    """
    import uno

    wb_path = Path(workbook)
    if not wb_path.is_absolute():
        wb_path = ROOT / wb_path
    codes = []
    for m in modules:
        if isinstance(m, str) and ("\n" in m or "Sub " in m or "Function " in m):
            codes.append(_clean_module(m))  # inline code string
            continue
        p = Path(m)
        if True:
            if not p.is_absolute():
                p = ROOT / p
            codes.append(_clean_module(p.read_text(encoding="utf-8-sig")))
    wrapper = ["Option VBASupport 1"]
    for i, name in enumerate(macros):
        wrapper.append(f"""Sub XlcRun{i}()
  On Error GoTo EH
  {name}
  Exit Sub
EH:
  XlcLog "{name}: error " & Err & ": " & Error$
End Sub""")
    wrapper.append("""Sub XlcLog(msg As String)
  Dim sh As Object
  Set sh = ThisComponent.Sheets.getByName("XlcLog")
  Dim r As Long
  r = 0
  Do While sh.getCellByPosition(0, r).getString() <> ""
    r = r + 1
  Loop
  sh.getCellByPosition(0, r).setString(msg)
End Sub""")
    wrapper.append("""Sub XlcPing()
  ThisComponent.Sheets.getByName("XlcLog").getCellByPosition(1, 0).setString("pong")
End Sub""")
    if extra_code:
        wrapper.append(extra_code)
    result = {"cells": {}, "errors": [], "sheets": []}
    with lo.office() as desk:
        url = uno.systemPathToFileUrl(str(wb_path.resolve()))
        doc = desk.loadComponentFromURL(url, "_blank", 0, (lo._pv("Hidden", False), lo._pv("MacroExecutionMode", 4)))
        try:
            doc.BasicLibraries.VBACompatibilityMode = True
            if not doc.BasicLibraries.hasByName("Standard"):
                doc.BasicLibraries.createLibrary("Standard")
            lib = doc.BasicLibraries.getByName("Standard")
            for i, code in enumerate(codes):
                lib.insertByName(f"XlcModule{i + 1}", code)
            lib.insertByName("XlcHarness", "\n".join(wrapper))
            doc.Sheets.insertNewByName("XlcLog", doc.Sheets.Count)
            sp = doc.getScriptProvider()
            try:
                sp.getScript("vnd.sun.star.script:Standard.XlcHarness.XlcPing?language=Basic&location=document").invoke((), (), ())
            except Exception:
                pass
            if doc.Sheets.getByName("XlcLog").getCellByPosition(1, 0).getString() != "pong":
                result["errors"].append("COMPILE: the modules did not compile in LibreOffice Basic (syntax LibreOffice does not "
                                        "support, a constant expression error such as 1/0, or a real syntax error). Nothing ran.")
            for i, name in enumerate(macros):
                try:
                    sp.getScript(f"vnd.sun.star.script:Standard.XlcHarness.XlcRun{i}?language=Basic&location=document").invoke((), (), ())
                except Exception as e:  # compile errors surface here
                    result["errors"].append(f"{name}: could not run ({str(e).splitlines()[0][:200]})")
            doc.calculateAll()
            log = doc.Sheets.getByName("XlcLog")
            r = 0
            while log.getCellByPosition(0, r).getString() or (r == 0 and log.getCellByPosition(1, 0).getString()):
                msg = log.getCellByPosition(0, r).getString()
                r += 1
                if not msg:
                    continue
                if ": error 91:" in msg:
                    msg += " (if the macro uses Debug.Print, this is LibreOffice's known Debug.Print limitation — check the output cells)"
                result["errors"].append(msg)
            doc.Sheets.removeByName("XlcLog")
            result["sheets"] = [doc.Sheets.getByIndex(i).Name for i in range(doc.Sheets.Count)]
            for sheet, addr in read:
                if not doc.Sheets.hasByName(sheet):
                    result["cells"][(sheet, addr)] = None
                    continue
                c = doc.Sheets.getByName(sheet).getCellRangeByName(addr)
                info = lo._cell_info(c)
                result["cells"][(sheet, addr)] = info["value"] if info["kind"] == "value" else info["text"]
            if save_as:
                out = Path(save_as)
                filt = "Calc MS Excel 2007 XML"
                doc.storeToURL(uno.systemPathToFileUrl(str(out.resolve())), (lo._pv("FilterName", filt),))
        finally:
            doc.close(True)
    return result
