Attribute VB_Name = "Snippets"
Option Explicit

' =====================================================================
' Lesson 5.4 - predict-the-output snippets (tasks 7, 11, 12 and 13)
'
' Read a snippet and type your prediction on the Practice sheet FIRST.
' Then click inside the snippet's Sub and press F5 to run it (Mac:
' Run > Run Sub/UserForm). Debug.Print writes to the Immediate window
' (View > Immediate Window; Ctrl+G on Windows). Snippet D stops with a
' run-time error on purpose: note the number, then click End.
' =====================================================================

' ---- Snippet A (task 7): ByVal vs ByRef ----
Sub AddFeeByVal(ByVal amount As Double)
    amount = amount + 25
End Sub

Sub AddFeeByRef(ByRef amount As Double)
    amount = amount + 25
End Sub

Sub SnippetA()
    Dim charge As Double
    charge = 100
    AddFeeByVal charge
    AddFeeByRef charge
    AddFeeByRef (charge)
    Debug.Print charge
End Sub

' ---- Snippet B (task 11): an error handler inside a loop ----
Sub SnippetB()
    Dim readings As Variant, i As Long, total As Long
    readings = Array("120", "abc", "95", "", "88")
    On Error GoTo BadReading
    For i = LBound(readings) To UBound(readings)
        total = total + CLng(readings(i))
    Next i
    Debug.Print "Total:"; total
    Exit Sub
BadReading:
    Debug.Print "Skipped item"; i; "- error"; Err.Number
    Resume Next
End Sub

' ---- Snippet C (task 12): On Error Resume Next ----
Sub SnippetC()
    Dim readings As Variant, i As Long, reading As Long, total As Long
    readings = Array("120", "abc", "95", "", "88")
    On Error Resume Next
    For i = LBound(readings) To UBound(readings)
        reading = CLng(readings(i))
        total = total + reading
    Next i
    On Error GoTo 0
    Debug.Print "Total:"; total
End Sub

' ---- Snippet D (task 13): two ways to call MATCH from VBA ----
Sub SnippetD()
    Dim payerList As Range, pos As Variant
    Set payerList = ThisWorkbook.Worksheets("Payers").Range("A2:A9")

    pos = Application.Match("PY09", payerList, 0)
    If IsError(pos) Then
        Debug.Print "Application.Match: not found"
    Else
        Debug.Print "Application.Match: row"; pos
    End If

    pos = Application.WorksheetFunction.Match("PY09", payerList, 0)
    Debug.Print "WorksheetFunction.Match: row"; pos
End Sub
