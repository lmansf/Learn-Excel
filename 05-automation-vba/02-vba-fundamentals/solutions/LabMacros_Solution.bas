Attribute VB_Name = "LabMacrosSolution"
Option Explicit

' =====================================================================
' Lesson 5.2 - reference solutions for Practice tasks 8-13 (SPOILERS)
'
' Import this into a spare copy of the workbook, or delete your own
' LabMacros module first, so the two modules don't fight over names.
'
' Labs sheet columns (row 1 = headers, data starts in row 2):
'    1 A LabResultID      5 E ResultValue      9 I AbnormalFlag
'    2 B EncounterID      6 F Units           10 J Priority
'    3 C TestCode         7 G RefLow          11 K CollectedDateTime
'    4 D TestName         8 H RefHigh         12 L ResultedDateTime
'                                             13 M TATFlag (task 11 fills it)
'
' Output sheet cells: B5 Warmup, B6 CountCriticals, B7 AveragePotassium,
'   B8 FindFirstCritical, B9 WriteCountAbove, B10 BuggyGlucoseCount.
'   Bonus: summary table E5:H40, top test in J5.
' =====================================================================

' Task 1
Sub Warmup()
    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = ThisWorkbook.Worksheets("Labs")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row      ' last filled row in column A
    ThisWorkbook.Worksheets("Output").Range("B5").Value = lastRow - 1   ' minus the header row
    Debug.Print "Warmup ran: " & (lastRow - 1) & " lab rows"
End Sub

' Task 8
Sub CountCriticals()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim cell As Range
    Dim criticalCount As Long

    Set ws = ThisWorkbook.Worksheets("Labs")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For Each cell In ws.Range("I2:I" & lastRow)          ' the AbnormalFlag column
        If cell.Value = "HH" Or cell.Value = "LL" Then
            criticalCount = criticalCount + 1
        End If
    Next cell

    ThisWorkbook.Worksheets("Output").Range("B6").Value = criticalCount
    MsgBox criticalCount & " critical results", vbInformation, "Count criticals"
End Sub

' Task 9
Sub AveragePotassium()
    Dim ws As Worksheet
    Dim r As Long, lastRow As Long
    Dim total As Double, n As Long

    Set ws = ThisWorkbook.Worksheets("Labs")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow
        If ws.Cells(r, 3).Value = "K" Then           ' column C = TestCode
            total = total + ws.Cells(r, 5).Value     ' column E = ResultValue
            n = n + 1
        End If
    Next r

    If n > 0 Then
        ThisWorkbook.Worksheets("Output").Range("B7").Value = total / n
        Debug.Print n & " potassium results, average " & Format(total / n, "0.00")
    Else
        MsgBox "No potassium results found.", vbExclamation
    End If
End Sub

' Task 10
Sub FindFirstCritical()
    Dim ws As Worksheet
    Dim r As Long, lastRow As Long
    Dim firstID As String

    Set ws = ThisWorkbook.Worksheets("Labs")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow
        Select Case ws.Cells(r, 9).Value             ' column I = AbnormalFlag
            Case "HH", "LL"
                firstID = ws.Cells(r, 1).Value       ' column A = LabResultID
                Exit For                             ' found it: stop looping
        End Select
    Next r

    If firstID = "" Then firstID = "None found"
    ThisWorkbook.Worksheets("Output").Range("B8").Value = firstID
    Debug.Print "First critical: " & firstID & " (row " & r & ")"
End Sub

' Task 11
Sub FlagSlowStat()
    Dim ws As Worksheet
    Dim r As Long, lastRow As Long
    Dim minutes As Long, slowCount As Long

    Set ws = ThisWorkbook.Worksheets("Labs")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow
        ws.Cells(r, 13).ClearContents                ' start clean so a re-run is safe
        If ws.Cells(r, 10).Value = "STAT" Then       ' column J = Priority
            ' "n" = minutes (not "m", which is months)
            minutes = DateDiff("n", ws.Cells(r, 11).Value, ws.Cells(r, 12).Value)
            If minutes > 60 Then
                ws.Cells(r, 13).Value = "SLOW"       ' column M = TATFlag
                slowCount = slowCount + 1
            End If
        End If
    Next r

    Debug.Print slowCount & " slow STAT results"
End Sub

' Task 12
Sub CountAbove()
    Dim testCode As String
    Dim limitInput As Variant

    testCode = UCase(Trim(InputBox("Which TestCode? (for example CREAT)", "Count results above a limit")))
    If testCode = "" Then Exit Sub                   ' Cancel (or nothing typed)

    limitInput = Application.InputBox("Count results above what value?", "Count results above a limit", Type:=1)
    If VarType(limitInput) = vbBoolean Then Exit Sub ' Cancel returns False

    WriteCountAbove testCode, CDbl(limitInput)
End Sub

Sub WriteCountAbove(testCode As String, limit As Double)
    Dim ws As Worksheet
    Dim r As Long, lastRow As Long, n As Long

    Set ws = ThisWorkbook.Worksheets("Labs")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow
        If ws.Cells(r, 3).Value = testCode Then
            If ws.Cells(r, 5).Value > limit Then n = n + 1
        End If
    Next r

    ThisWorkbook.Worksheets("Output").Range("B9").Value = n
    MsgBox n & " " & testCode & " results above " & limit, vbInformation, "Count results above a limit"
End Sub

' Task 13
Sub BuggyGlucoseCount()
    Dim ws As Worksheet
    Dim r As Long, lastRow As Long
    Dim highCount As Long

    Set ws = ThisWorkbook.Worksheets("Labs")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow                             ' FIX 1: start below the header row
        If ws.Cells(r, 3).Value = "GLU" And ws.Cells(r, 5).Value > 180 Then   ' FIX 2: "GLU", not "Glu"
            highCount = highCount + 1
        End If
    Next r

    ThisWorkbook.Worksheets("Output").Range("B10").Value = highCount
End Sub
