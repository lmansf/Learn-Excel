Attribute VB_Name = "LabSnapshotSolution"
Option Explicit

' =====================================================================
' Lesson 5.2 - reference solution for the bonus (SPOILERS)
' Import into a spare copy of the workbook, or remove your own
' LabSnapshot first.
' =====================================================================

Sub LabSnapshot()
    Dim wsLab As Worksheet, wsOut As Worksheet
    Dim r As Long, lastRow As Long
    Dim k As Long, nextRow As Long
    Dim flag As String, code As String
    Dim found As Boolean
    Dim pct As Double, topPct As Double, topCode As String

    Set wsLab = ThisWorkbook.Worksheets("Labs")
    Set wsOut = ThisWorkbook.Worksheets("Output")
    lastRow = wsLab.Cells(wsLab.Rows.Count, 1).End(xlUp).Row

    ' 1. Reset: remove old fills and the old summary so the macro can be re-run
    wsLab.Range("A2:L" & lastRow).Interior.ColorIndex = xlNone
    wsOut.Range("E5:H40").ClearContents
    wsOut.Range("J5").ClearContents
    nextRow = 5                                      ' first empty row of the summary table

    For r = 2 To lastRow
        flag = wsLab.Cells(r, 9).Value               ' column I = AbnormalFlag
        code = wsLab.Cells(r, 3).Value               ' column C = TestCode

        ' 2. Color the row (columns A:L) by its flag
        Select Case flag
            Case "HH", "LL"
                wsLab.Range("A" & r & ":L" & r).Interior.Color = RGB(255, 199, 206)
            Case "H", "L"
                wsLab.Range("A" & r & ":L" & r).Interior.Color = RGB(255, 235, 156)
        End Select

        ' 3. Find this test in the summary table; add a new row if it isn't there yet
        found = False
        For k = 5 To nextRow - 1
            If wsOut.Cells(k, 5).Value = code Then
                found = True
                Exit For                             ' k now points at the test's row
            End If
        Next k
        If Not found Then
            k = nextRow
            wsOut.Cells(k, 5).Value = code
            wsOut.Cells(k, 6).Value = 0
            wsOut.Cells(k, 7).Value = 0
            nextRow = nextRow + 1
        End If

        ' 4. Update the counts
        wsOut.Cells(k, 6).Value = wsOut.Cells(k, 6).Value + 1          ' Results
        If flag <> "N" Then
            wsOut.Cells(k, 7).Value = wsOut.Cells(k, 7).Value + 1      ' Abnormal
        End If
    Next r

    ' 5. Percent abnormal for each test, tracking the highest
    For k = 5 To nextRow - 1
        pct = wsOut.Cells(k, 7).Value / wsOut.Cells(k, 6).Value
        wsOut.Cells(k, 8).Value = pct
        If pct > topPct Then
            topPct = pct
            topCode = wsOut.Cells(k, 5).Value
        End If
    Next k
    wsOut.Range("H5:H40").NumberFormat = "0.0%"
    wsOut.Range("J5").Value = topCode

    MsgBox "Tests summarized: " & (nextRow - 5) & vbNewLine & _
           "Highest % abnormal: " & topCode & " (" & Format(topPct, "0.0%") & ")", _
           vbInformation, "Lab snapshot"
End Sub
