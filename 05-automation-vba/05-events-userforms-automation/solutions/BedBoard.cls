'==============================================================================
' Lesson 5.5 - BedBoard sheet event code (SOLUTION - spoiler!)
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

    ' React only to the Status column; ignore edits anywhere else on the sheet.
    Set changed = Intersect(Target, lo.ListColumns("Status").DataBodyRange)
    If changed Is Nothing Then Exit Sub

    On Error GoTo CleanUp
    Application.EnableEvents = False      ' our own writes below must not fire this event again

    For Each cell In changed.Cells
        ' Tidy what was typed ("dirty" -> "Dirty"). Without EnableEvents = False,
        ' this write would fire Worksheet_Change again.
        If VarType(cell.Value) = vbString Then cell.Value = StrConv(Trim$(cell.Value), vbProperCase)

        r = cell.Row - lo.DataBodyRange.Row + 1          ' row number inside the table
        lo.ListColumns("StatusTime").DataBodyRange.Cells(r).Value = Now
        WriteLog "Bed status", lo.ListColumns("BedID").DataBodyRange.Cells(r).Value & " -> " & cell.Value
    Next cell

CleanUp:
    Application.EnableEvents = True       ' runs on success AND after an error
    If Err.Number <> 0 Then MsgBox "Bed board update failed: " & Err.Description, vbExclamation
End Sub
