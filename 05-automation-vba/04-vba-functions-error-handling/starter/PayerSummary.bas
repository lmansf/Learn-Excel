Attribute VB_Name = "PayerSummary"
Option Explicit

' =====================================================================
' Lesson 5.4 - bonus starter: BuildPayerSummary
' The full spec is on the Bonus sheet. The helpers at the bottom are
' finished for you. Windows: use Scripting.Dictionary. Mac: use a
' Collection plus arrays (see the guide).
' =====================================================================

Private Const OUT_SHEET As String = "PayerSummary"

Public Sub BuildPayerSummary()
    ' Declare your variables here (Option Explicit is on)

    ' 1. If the Claims or Payers sheet is missing: show a friendly MsgBox and Exit Sub

    ' 2. Remember Application.Calculation, then On Error GoTo Fail, then turn off
    '    ScreenUpdating and switch to manual calculation

    ' 3. Read Claims and Payers into arrays (CurrentRegion.Value2) and find the
    '    columns you need with HeaderColumn

    ' 4. Lookup dictionary PayerID -> PayerName

    ' 5. One pass over the claims: skip (and count) rows whose BilledAmount or
    '    PaidAmount is not a number; otherwise update the payer's claims, billed,
    '    paid and denied totals

    ' 6. Build the output array: header row + one row per payer, and the totals

    ' 7. Clear or create the PayerSummary sheet, write the array, sort the payer
    '    rows by Billed (largest first), then add the TOTAL row and, two rows
    '    below it, "Rows skipped" with the count

CleanExit:
    ' 8. Restore ScreenUpdating and Calculation. This block runs after success
    '    AND after an error.
    Exit Sub

Fail:
    ' 9. Tell the user what went wrong (Err.Description, Err.Number), then
    '    Resume CleanExit
End Sub

' ---------------------------------------------------------------------
' Helpers
' ---------------------------------------------------------------------

' True if this workbook has a sheet with that name.
Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)    ' error 9 if there is no such sheet
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

' Column number of a header in row 1. Raises a clear error if it is missing.
Private Function HeaderColumn(ws As Worksheet, ByVal headerText As String) As Long
    Dim pos As Variant
    pos = Application.Match(headerText, ws.Rows(1), 0)
    If IsError(pos) Then
        Err.Raise vbObjectError + 513, "HeaderColumn", _
                  "Column '" & headerText & "' was not found on sheet '" & ws.Name & "'."
    End If
    HeaderColumn = CLng(pos)
End Function

' True for real numbers (a blank cell or text is not a number).
Private Function IsNumberValue(v As Variant) As Boolean
    Select Case VarType(v)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte
            IsNumberValue = True
        Case Else
            IsNumberValue = False
    End Select
End Function
