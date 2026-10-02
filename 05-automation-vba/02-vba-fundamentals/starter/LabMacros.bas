Attribute VB_Name = "LabMacros"
Option Explicit

' =====================================================================
' Lesson 5.2 - VBA Fundamentals - starter module (Bluestone Health, fictional)
'
' Import it: in the Visual Basic Editor choose File > Import File... and
' pick this file. Write your code between each Sub and End Sub, then run
' it with F5 (cursor inside the Sub) or from Developer > Macros.
'
' A compile error in one Sub (a red line, or an undeclared variable) can
' stop every macro in this module from running, even Warmup. If that
' happens, choose Debug > Compile VBAProject to jump to the problem line.
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

' Task 1 - already complete. Run it to check that macros work in your copy.
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
    ' TODO: count the results whose AbnormalFlag (column I) is HH or LL.
    ' Use For Each over the range I2:I<lastRow>, then write the count
    ' to Output!B6. (Copy the Dim, Set ws, and lastRow lines from Warmup.)
End Sub

' Task 9
Sub AveragePotassium()
    ' TODO: with For r = 2 To lastRow, add up ResultValue (column E) for
    ' every row whose TestCode (column C) is K, and count those rows.
    ' Write total / count to Output!B7.
End Sub

' Task 10
Sub FindFirstCritical()
    ' TODO: find the FIRST row (from the top) whose AbnormalFlag is HH or LL.
    ' Write its LabResultID (column A) to Output!B8 and leave the loop
    ' with Exit For.
End Sub

' Task 11
Sub FlagSlowStat()
    ' TODO: for every STAT row (Priority, column J), work out the
    ' turnaround in minutes: DateDiff("n", Collected (K), Resulted (L)).
    ' If it is more than 60, write SLOW in column M (TATFlag) of that row.
    ' Leave every other row's column M empty.
End Sub

' Task 12
Sub CountAbove()
    ' TODO: ask for a TestCode with InputBox and for a limit with
    ' Application.InputBox(..., Type:=1). Stop if the user cancels.
    ' Then call WriteCountAbove with the two answers. Pass the Variant
    ' limit as CDbl(yourVariable): a Variant passed to an argument declared
    ' As Double stops with Compile error: ByRef argument type mismatch.
End Sub

Sub WriteCountAbove(testCode As String, limit As Double)
    ' TODO: count the rows whose TestCode = testCode and ResultValue > limit.
    ' Write the count to Output!B9 and show it in a MsgBox.
End Sub

' Task 13 - this macro should count glucose (GLU) results above 180 mg/dL
' and write the count to Output!B10. It has TWO bugs. Debug it and fix it.
Sub BuggyGlucoseCount()
    Dim ws As Worksheet
    Dim r As Long, lastRow As Long
    Dim highCount As Long

    Set ws = ThisWorkbook.Worksheets("Labs")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = 1 To lastRow
        If ws.Cells(r, 3).Value = "Glu" And ws.Cells(r, 5).Value > 180 Then
            highCount = highCount + 1
        End If
    Next r

    ThisWorkbook.Worksheets("Output").Range("B10").Value = highCount
End Sub

' Bonus - see the Bonus sheet for the full brief.
Sub LabSnapshot()
    ' TODO (bonus):
    ' 1. Clear old fills on Labs A2:L<lastRow>, Output!E5:H40 and Output!J5.
    ' 2. Color each row A:L: HH/LL RGB(255, 199, 206), H/L RGB(255, 235, 156).
    ' 3. Build the summary table at Output!E5 down: TestCode, Results,
    '    Abnormal (flag is not N), PctAbnormal (Abnormal / Results).
    ' 4. Write the TestCode with the highest PctAbnormal to Output!J5.
End Sub
