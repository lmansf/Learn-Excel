Attribute VB_Name = "PayerSummarySolution"
Option Explicit

' =====================================================================
' Lesson 5.4 - reference solution for the bonus (SPOILERS)
' Windows: Scripting.Dictionary. Excel for Mac: use PayerSummaryMac_Solution.bas.
'
' BuildPayerSummary writes a payer scorecard to the PayerSummary sheet:
'   row 1        PayerID | PayerName | Claims | Billed | Paid | Denied | DenialRate
'   rows 2..n+1  one row per payer, sorted by Billed (largest first)
'   row n+2      TOTAL
'   row n+4      "Rows skipped" and how many claim rows had a blank/text amount
' Remove your own PayerSummary module first (or import into a spare copy).
' =====================================================================

Private Const OUT_SHEET As String = "PayerSummary"

Public Sub BuildPayerSummary()
    Dim wsClaims As Worksheet, wsOut As Worksheet
    Dim claims As Variant, payers As Variant, outData() As Variant
    Dim payerNames As Object, stats As Object
    Dim colPayer As Long, colBilled As Long, colPaid As Long, colStatus As Long
    Dim r As Long, n As Long, totalRow As Long, skipped As Long
    Dim payerID As Variant, s As Variant
    Dim totClaims As Long, totDenied As Long, totBilled As Double, totPaid As Double
    Dim prevCalc As Long

    ' 1. Check first: a friendly message beats a crash
    If Not SheetExists("Claims") Or Not SheetExists("Payers") Then
        MsgBox "BuildPayerSummary needs a sheet named 'Claims' and a sheet named 'Payers'." & vbNewLine & _
               "One of them is missing or has been renamed. Fix that, then run the macro again.", _
               vbExclamation, "Payer summary"
        Exit Sub
    End If

    ' 2. From here on, any run-time error jumps to Fail, which restores Excel's settings
    prevCalc = Application.Calculation
    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set wsClaims = ThisWorkbook.Worksheets("Claims")
    claims = wsClaims.Range("A1").CurrentRegion.Value2     ' Value2: plain numbers, no Currency/Date types
    colPayer = HeaderColumn(wsClaims, "PayerID")
    colBilled = HeaderColumn(wsClaims, "BilledAmount")
    colPaid = HeaderColumn(wsClaims, "PaidAmount")
    colStatus = HeaderColumn(wsClaims, "ClaimStatus")
    payers = ThisWorkbook.Worksheets("Payers").Range("A1").CurrentRegion.Value2

    ' 3. Lookup dictionary: PayerID -> PayerName
    Set payerNames = CreateObject("Scripting.Dictionary")
    For r = 2 To UBound(payers, 1)
        payerNames(payers(r, 1)) = payers(r, 2)
    Next r

    ' 4. One pass over the claims: PayerID -> Array(claims, billed, paid, denied)
    Set stats = CreateObject("Scripting.Dictionary")
    For r = 2 To UBound(claims, 1)
        If IsNumberValue(claims(r, colBilled)) And IsNumberValue(claims(r, colPaid)) Then
            payerID = claims(r, colPayer)
            If Not stats.Exists(payerID) Then stats.Add payerID, Array(0, 0, 0, 0)
            s = stats(payerID)                 ' copy the array out of the dictionary...
            s(0) = s(0) + 1
            s(1) = s(1) + claims(r, colBilled)
            s(2) = s(2) + claims(r, colPaid)
            If claims(r, colStatus) = "Denied" Then s(3) = s(3) + 1
            stats(payerID) = s                 ' ...and put it back: stats(payerID)(0) = 1 would not stick
        Else
            skipped = skipped + 1              ' blank or text amount
        End If
    Next r
    n = stats.Count
    If n = 0 Then Err.Raise vbObjectError + 514, "BuildPayerSummary", "No claim rows with numeric amounts were found."

    ' 5. Header + one row per payer, built in memory, plus the running totals
    ReDim outData(1 To n + 1, 1 To 7)
    outData(1, 1) = "PayerID": outData(1, 2) = "PayerName": outData(1, 3) = "Claims"
    outData(1, 4) = "Billed": outData(1, 5) = "Paid": outData(1, 6) = "Denied": outData(1, 7) = "DenialRate"
    r = 1
    For Each payerID In stats.Keys
        r = r + 1
        s = stats(payerID)
        outData(r, 1) = payerID
        If payerNames.Exists(payerID) Then
            outData(r, 2) = payerNames(payerID)
        Else
            outData(r, 2) = "(unknown payer)"
        End If
        outData(r, 3) = s(0): outData(r, 4) = s(1): outData(r, 5) = s(2): outData(r, 6) = s(3)
        outData(r, 7) = s(3) / s(0)
        totClaims = totClaims + s(0): totBilled = totBilled + s(1)
        totPaid = totPaid + s(2): totDenied = totDenied + s(3)
    Next payerID

    ' 6. A fresh output sheet: clear it if it exists, add it if it doesn't
    If SheetExists(OUT_SHEET) Then
        Set wsOut = ThisWorkbook.Worksheets(OUT_SHEET)
        wsOut.Cells.Clear
    Else
        Set wsOut = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets("Payers"))
        wsOut.Name = OUT_SHEET
    End If

    ' 7. Write the block in one go, sort the payer rows by Billed (largest
    '    first), then add the TOTAL row and the skipped-rows line
    With wsOut
        .Range("A1").Resize(n + 1, 7).Value = outData
        .Range("A1").Resize(n + 1, 7).Sort Key1:=.Range("D1"), Order1:=xlDescending, Header:=xlYes
        totalRow = n + 2
        .Cells(totalRow, 1).Value = "TOTAL"
        .Cells(totalRow, 3).Value = totClaims
        .Cells(totalRow, 4).Value = totBilled
        .Cells(totalRow, 5).Value = totPaid
        .Cells(totalRow, 6).Value = totDenied
        .Cells(totalRow, 7).Value = totDenied / totClaims
        .Cells(totalRow + 2, 1).Value = "Rows skipped"
        .Cells(totalRow + 2, 2).Value = skipped
        .Range("D:E").NumberFormat = "#,##0.00"
        .Range("G:G").NumberFormat = "0.0%"
        .Rows(1).Font.Bold = True
        .Rows(totalRow).Font.Bold = True
        .Columns("A:G").AutoFit
    End With

CleanExit:
    On Error Resume Next                   ' cleanup must never raise an error of its own
    Application.Calculation = prevCalc
    Application.ScreenUpdating = True
    Exit Sub

Fail:
    MsgBox "BuildPayerSummary stopped: " & Err.Description & " (error " & Err.Number & ")", _
           vbCritical, "Payer summary"
    Resume CleanExit
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
