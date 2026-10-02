Attribute VB_Name = "modRefresh"
'==============================================================================
' Lesson 6.1 - Capstone: one-click refresh for the 2025 performance review
' STARTER: write the code for STEP 1 to STEP 5 in RefreshReview.
'
' Import: in the VBE (Alt+F11; Mac: Option+F11) choose File > Import File...
' Run:    Alt+F8 (Mac: Option+F8) > RefreshReview > Run, or a button on the
'         Dashboard. Save the workbook as .xlsm to keep the code.
'
' RefreshReview does five things, in order:
'   1. Refreshes every Power Query query and PivotTable (Data > Refresh All).
'   2. Removes duplicate rows from the Surveys sheet (SurveyID is the key).
'   3. Recalculates every formula.
'   4. Appends an audit row to the RefreshLog sheet (created on the first run):
'      A RefreshedAt | B Encounters | C ED_Visits | D Claims | E Surveys | F DuplicatesRemoved
'   5. Writes the refresh time into Dashboard!H2.
'==============================================================================
Option Explicit

Public Sub RefreshReview()
    Dim oldCalc As XlCalculation
    Dim surveysBefore As Long, surveysAfter As Long
    Dim wsLog As Worksheet
    Dim r As Long

    oldCalc = Application.Calculation
    On Error GoTo Fail
    Application.ScreenUpdating = False

    ' STEP 1 (your code): refresh every query and PivotTable in this workbook (one line).
    '         Then wait for background queries with Application.CalculateUntilAsyncQueriesDone
    '         (wrap that line in On Error Resume Next / On Error GoTo Fail so a failed wait
    '         doesn't stop the macro).

    ' STEP 2 (your code): count the Surveys data rows with DataRowCount("Surveys") into surveysBefore,
    '         remove duplicate rows (SurveyID, column 1, is the key; the sheet has a header row),
    '         then count again into surveysAfter.
    '         Hint: Range("A1").CurrentRegion.RemoveDuplicates Columns:=..., Header:=...

    ' STEP 3 (your code): switch calculation to automatic and recalculate everything.

    ' STEP 4 (your code): Set wsLog = LogSheet(), find the next empty row r in column A, and write:
    '         A = Now, B = Encounters rows, C = ED_Visits rows, D = Claims rows,
    '         E = surveysAfter, F = surveysBefore - surveysAfter

    ' STEP 5 (your code): write Now into Dashboard!H2 and activate the Dashboard sheet.

Done:
    Application.Calculation = oldCalc
    Application.ScreenUpdating = True
    Exit Sub

Fail:
    MsgBox "RefreshReview stopped: " & Err.Description & " (error " & Err.Number & ")", vbExclamation, "Refresh"
    Resume Done
End Sub

' Number of data rows under the header row: last used cell in column A, minus the header.
Private Function DataRowCount(ByVal sheetName As String) As Long
    With ThisWorkbook.Worksheets(sheetName)
        DataRowCount = .Cells(.Rows.Count, "A").End(xlUp).Row - 1
    End With
End Function

' Returns the RefreshLog sheet, creating it with headers the first time.
Private Function LogSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("RefreshLog")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "RefreshLog"
        ws.Range("A1:F1").Value = Array("RefreshedAt", "Encounters", "ED_Visits", "Claims", "Surveys", "DuplicatesRemoved")
        ws.Range("A1:F1").Font.Bold = True
        ws.Columns("A").NumberFormat = "mm/dd/yyyy hh:mm"
        ws.Columns("A:F").ColumnWidth = 14
    End If
    Set LogSheet = ws
End Function
