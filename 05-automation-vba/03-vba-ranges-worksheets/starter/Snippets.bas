Attribute VB_Name = "Snippets"
Option Explicit

' =====================================================================
' Lesson 5.3 - predict-the-output snippets (Practice tasks 1-6)
'
' For each snippet: predict what it prints (or which error it stops
' with), type your prediction on the Practice sheet, THEN click inside
' the Sub and press F5 to check. Debug.Print writes to the Immediate
' window (Ctrl + G; Mac: View > Immediate Window).
' =====================================================================

' Snippet A (task 1)
Sub SnippetA_OffsetResize()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Encounters")
    Debug.Print ws.Range("C2").Offset(3, 4).Resize(5).Address
End Sub

' Snippet B (task 2)
Sub SnippetB_DataBody()
    Dim ws As Worksheet, block As Range
    Set ws = ThisWorkbook.Worksheets("Encounters")
    Set block = ws.Range("E50").CurrentRegion            ' any cell inside the data works
    Set block = block.Offset(1).Resize(block.Rows.Count - 1)
    Debug.Print block.Address(False, False)
End Sub

' Snippet C (tasks 3 and 4)
Sub SnippetC_LastRow()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Encounters")
    Debug.Print ws.Cells(ws.Rows.Count, "A").End(xlUp).Row     ' line 1: EncounterID
    Debug.Print ws.Cells(ws.Rows.Count, "F").End(xlUp).Row     ' line 2: AdmitSource
    Debug.Print ws.Range("F1").End(xlDown).Row                 ' line 3: AdmitSource again
End Sub

' Snippet D (task 5)
Sub SnippetD_WhichSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Encounters")
    ThisWorkbook.Worksheets("Practice").Activate               ' Practice is now the active sheet
    Debug.Print ws.Range(Cells(2, 1), Cells(2, 11)).Address    ' row 2, columns A to K
End Sub

' Snippet E (task 6)
Sub SnippetE_DeleteOld()
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("TypeSummary_2024").Delete         ' last year's summary sheet
    Application.DisplayAlerts = True
    Debug.Print "Old summary deleted"
End Sub
