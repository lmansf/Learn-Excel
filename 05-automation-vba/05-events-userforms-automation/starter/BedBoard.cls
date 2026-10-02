'==============================================================================
' Lesson 5.5 - BedBoard sheet event code (STARTER) - practice task 7
'
' HOW TO USE: do NOT import this file. In the VBE Project Explorer, double-click
' the BedBoard sheet (it is listed as "SheetN (BedBoard)"), delete anything
' already in its code window (a second Option Explicit won't compile), then
' paste everything below the dashed line. Needs WriteLog from modLog.
'------------------------------------------------------------------------------
Option Explicit

' Runs after cells on this sheet are changed by a person or by VBA.
' Target holds EVERY cell changed by that one action (it can be many cells).
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim lo As ListObject
    Dim changed As Range, cell As Range
    Dim r As Long

    Set lo = Me.ListObjects("tblBeds")

    ' TODO 1: Set changed = the part of Target that is inside the Status column
    '         (Intersect with lo.ListColumns("Status").DataBodyRange).
    '         If changed Is Nothing, Exit Sub.

    ' TODO 2: On Error GoTo CleanUp, then switch events off so your own writes
    '         don't fire this handler again.

    ' TODO 3: For Each cell In changed.Cells
    '            - (optional) tidy the text with StrConv(..., vbProperCase)
    '            - r = cell.Row - lo.DataBodyRange.Row + 1   (row inside the table)
    '            - stamp Now into lo.ListColumns("StatusTime").DataBodyRange.Cells(r)
    '            - WriteLog "Bed status", <BedID of row r> & " -> " & cell.Value
    '         Next cell

CleanUp:
    ' TODO 4: switch events back on (this line must run even after an error)
End Sub
