Attribute VB_Name = "modCensusReport"
' ==========================================================================
' Lesson 5.1 - Recording Your First Macros - reference solution
' SPOILER: record your own macro first, then compare.
' 
' To import: open the VBE with Alt + F11 (Mac: Option + F11, or Developer >
' Visual Basic), choose File > Import File..., and pick this .bas file. If your
' workbook already has a macro with the same name, rename or delete one of them
' first so it's clear which one runs.
' An imported macro has no shortcut key (the 'Keyboard Shortcut' line is only a
' comment). To set one, open Macros (Alt + F8; Mac: Option + F8), select it, and
' click Options...
' Save the workbook as .xlsm to keep the code.
' ==========================================================================
Option Explicit

Sub FormatCensusReport()
'
' FormatCensusReport Macro
' Formats the raw daily census export and adds a TOTAL row.
'
' Keyboard Shortcut: Ctrl+Shift+R
'
    ' --- Use Relative References is OFF: these lines name exact cells ---
    Range("A1:H1").Select
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "Occupancy"
    Range("H2").Select
    Columns("A:A").Select
    Selection.NumberFormat = "mm/dd/yyyy"
    Range("A1").Select
    Selection.End(xlDown).Select
    ' --- Use Relative References turned ON: moves are stored as offsets ---
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "TOTAL"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.Offset(0, 2).Range("A1:D1").Select
    Selection.FormulaR1C1 = "=SUM(R2C:R[-1]C)"
    ActiveCell.Offset(0, 4).Range("A1").Select
    Selection.FormulaR1C1 = "=RC[-1]/RC[-4]"
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.0%"
    ActiveCell.Rows("1:1").EntireRow.Select
    Selection.Font.Bold = True
    ' --- Use Relative References turned OFF again ---
    Columns("A:H").Select
    Columns("A:H").EntireColumn.AutoFit
    Range("A1").Select
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
End Sub

Sub FormatCensusReport_Clean()
'   The same report as FormatCensusReport, tidied by hand:
'   no Select/Selection pairs and no default properties.
'   Run it with a raw census export as the active sheet.
    With Range("A1:H1")
        .Font.Bold = True
        .Interior.ThemeColor = xlThemeColorAccent1
        .Interior.TintAndShade = 0.8
    End With
    Range("H1").Value = "Occupancy"
    Columns("A").NumberFormat = "mm/dd/yyyy"

    ' The TOTAL row goes one row below the last date in column A
    ' (the same cell that Ctrl + Down arrow, then Down arrow, would reach).
    With Range("A1").End(xlDown).Offset(1, 0)
        .Value = "TOTAL"
        ' Columns D:G: from row 2 (fixed) down to the row above (relative).
        .Offset(0, 3).Resize(1, 4).FormulaR1C1 = "=SUM(R2C:R[-1]C)"
        ' Column H: patient days (G) divided by staffed-bed days (D).
        .Offset(0, 7).FormulaR1C1 = "=RC[-1]/RC[-4]"
        .Offset(0, 7).NumberFormat = "0.0%"
        .EntireRow.Font.Bold = True
    End With

    Columns("A:H").AutoFit
    ' Freeze the header row (window settings always act on the active sheet).
    ActiveWindow.FreezePanes = False
    ActiveWindow.ScrollRow = 1
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
        .FreezePanes = True
    End With
End Sub
