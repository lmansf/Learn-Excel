Attribute VB_Name = "modHighAcuity"
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

Sub ExtractHighAcuity_AsRecorded()
'
' ExtractHighAcuity Macro
' Copies ESI 1-2 visits to a new HighAcuity sheet, sorted by ESI level and arrival.
'
' Keyboard Shortcut: Ctrl+Shift+H
'
'   What the recorder wrote on ED_Nov (the <-- notes are added). It works on
'   November. Lines marked <-- are tied to November and need editing for ED_Dec.
    Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$J$81").AutoFilter Field:=5, Criteria1:="<=2", _
        Operator:=xlAnd                                 ' <-- November's exact range
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Sheets("Sheet1").Select                             ' <-- the new sheet's
    Sheets("Sheet1").Name = "HighAcuity"                ' <-- temporary name
    ActiveWorkbook.Worksheets("HighAcuity").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("HighAcuity").Sort.SortFields.Add2 Key:=Range( _
        "E2:E21"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal                                    ' <-- November's row count
    ActiveWorkbook.Worksheets("HighAcuity").Sort.SortFields.Add2 Key:=Range( _
        "C2:C21"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal                                    ' <-- November's row count
    With ActiveWorkbook.Worksheets("HighAcuity").Sort
        .SetRange Range("A1:J21")                       ' <-- November's row count
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Sheets("ED_Nov").Select                             ' <-- November's sheet
    Range("A1").Select
    Selection.AutoFilter
End Sub

Sub ExtractHighAcuity()
'
' ExtractHighAcuity Macro
' Copies ESI 1-2 visits to a new HighAcuity sheet, sorted by ESI level and arrival.
'
' Keyboard Shortcut: Ctrl+Shift+H
'
'   Recorded on ED_Nov, then edited to run on ED_Dec. Each EDIT comment marks a change.
'   Delete any old HighAcuity sheet before you run it: a second sheet with that
'   name is not allowed (run-time error 1004).
    Sheets("ED_Dec").Select                             ' EDIT: added, start on December
    Range("A1").Select
    Selection.AutoFilter
    Range("A1").CurrentRegion.AutoFilter Field:=5, Criteria1:="<=2", _
        Operator:=xlAnd                                 ' EDIT: was ActiveSheet.Range("$A$1:$J$81")
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveSheet.Name = "HighAcuity"                     ' EDIT: was Sheets("Sheet1").Select / .Name
    ActiveWorkbook.Worksheets("HighAcuity").Sort.SortFields.Clear
    ' EDIT: keys were Range("E2:E21") and Range("C2:C21"). One cell is enough
    ' to name the key column. (.Add works in Excel 2007 and later; Microsoft 365 records .Add2.)
    ActiveWorkbook.Worksheets("HighAcuity").Sort.SortFields.Add Key:=Range("E1"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("HighAcuity").Sort.SortFields.Add Key:=Range("C1"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("HighAcuity").Sort
        .SetRange Range("A1").CurrentRegion             ' EDIT: was Range("A1:J21")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Sheets("ED_Dec").Select                             ' EDIT: was Sheets("ED_Nov")
    Range("A1").Select
    Selection.AutoFilter
End Sub
