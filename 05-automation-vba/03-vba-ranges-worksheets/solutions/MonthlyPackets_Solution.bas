Attribute VB_Name = "MonthlyPacketsSolution"
Option Explicit

' =====================================================================
' Lesson 5.3 - reference solution for the bonus (SPOILERS)
' Import into a spare copy of the workbook, or remove your own
' BuildMonthlyPackets first. The helpers below are Private copies.
' =====================================================================

Private Function LastRow(ws As Worksheet, Optional col As Variant = 1) As Long
    ' The last filled row in one column: start at the bottom of the sheet and
    ' jump up, like Ctrl + Up arrow. col can be a letter ("A") or a number (1).
    LastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function

Private Function SheetExists(ByVal sheetName As String) As Boolean
    ' True if THIS workbook has a worksheet with that name (ignoring case,
    ' because Excel treats "typesummary" and "TypeSummary" as the same name).
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, sheetName, vbTextCompare) = 0 Then
            SheetExists = True
            Exit Function
        End If
    Next ws
End Function

Private Sub DeleteSheetIfExists(ByVal sheetName As String)
    ' Deletes the sheet if it's there, without Excel's "are you sure?" prompt.
    ' ByVal lets you pass any text: a String, a Variant, or a cell's value.
    If SheetExists(sheetName) Then
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets(sheetName).Delete
        Application.DisplayAlerts = True
    End If
End Sub

Sub BuildMonthlyPackets()
    Const NCOLS As Long = 11                     ' Encounters columns A:K
    Dim wsEnc As Worksheet, wsTpl As Worksheet, wsPk As Worksheet
    Dim src As Variant, buf() As Variant
    Dim lastR As Long, i As Long, j As Long, n As Long, m As Long
    Dim monthStart As Date, nextMonth As Date
    Dim sheetName As String, bestName As String
    Dim total As Double, bestTotal As Double

    Set wsEnc = ThisWorkbook.Worksheets("Encounters")
    Set wsTpl = ThisWorkbook.Worksheets("PacketTemplate")
    lastR = LastRow(wsEnc, "A")
    src = wsEnc.Range("A2").Resize(lastR - 1, NCOLS).Value   ' read the export ONCE

    Application.ScreenUpdating = False
    On Error GoTo CleanUp

    For m = 1 To 12
        monthStart = DateSerial(2025, m, 1)
        nextMonth = DateSerial(2025, m + 1, 1)   ' month 13 rolls over to January 2026
        sheetName = Format(monthStart, "yyyy-mm")

        ' 1. Collect this month's inpatient rows in a buffer as big as the export
        ReDim buf(1 To UBound(src, 1), 1 To NCOLS)
        n = 0
        total = 0
        For i = 1 To UBound(src, 1)
            If src(i, 3) = "Inpatient" And src(i, 7) >= monthStart And src(i, 7) < nextMonth Then
                n = n + 1
                For j = 1 To NCOLS
                    buf(n, j) = src(i, j)
                Next j
                total = total + src(i, 11)
            End If
        Next i

        ' 2. A fresh copy of the template, named for the month
        DeleteSheetIfExists sheetName
        wsTpl.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
        Set wsPk = ActiveSheet                   ' a copied sheet is always the active one
        wsPk.Name = sheetName

        ' 3. Header block, data, sort
        wsPk.Range("B2").Value = monthStart
        wsPk.Range("B3").Value = n
        wsPk.Range("B4").Value = total
        If n > 0 Then
            ' A range smaller than the array takes just the top-left n rows of it
            wsPk.Range("A7").Resize(n, NCOLS).Value = buf
            wsPk.Range("A6").CurrentRegion.Sort Key1:=wsPk.Range("K6"), Order1:=xlDescending, Header:=xlYes
        End If

        ' 4. Remember the biggest month so far
        If total > bestTotal Then
            bestTotal = total
            bestName = sheetName
        End If
    Next m

    ThisWorkbook.Worksheets("Output").Range("B9").Value = bestName

CleanUp:
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        MsgBox "BuildMonthlyPackets stopped: " & Err.Description, vbExclamation
    Else
        MsgBox "12 packets built. Largest month: " & bestName, vbInformation
    End If
End Sub
