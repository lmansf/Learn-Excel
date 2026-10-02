Attribute VB_Name = "modRefresh"
'==============================================================================
' Lesson 6.1 - Capstone: one-click refresh for the 2025 performance review
' SOLUTION (spoiler): try starter/RefreshReview_Starter.bas first.
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

    ' 1. Refresh queries and PivotTables, then wait for background queries
    ThisWorkbook.RefreshAll
    On Error Resume Next                      ' a failed wait must not stop the review
    Application.CalculateUntilAsyncQueriesDone
    On Error GoTo Fail

    ' 2. Remove duplicate surveys, keeping the first copy of each SurveyID
    surveysBefore = DataRowCount("Surveys")
    ThisWorkbook.Worksheets("Surveys").Range("A1").CurrentRegion.RemoveDuplicates Columns:=1, Header:=xlYes
    surveysAfter = DataRowCount("Surveys")

    ' 3. Recalculate every formula in every open workbook
    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFull

    ' 4. Append the audit row on the next empty row of RefreshLog
    Set wsLog = LogSheet()
    r = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row + 1
    wsLog.Cells(r, "A").Value = Now
    wsLog.Cells(r, "B").Value = DataRowCount("Encounters")
    wsLog.Cells(r, "C").Value = DataRowCount("ED_Visits")
    wsLog.Cells(r, "D").Value = DataRowCount("Claims")
    wsLog.Cells(r, "E").Value = surveysAfter
    wsLog.Cells(r, "F").Value = surveysBefore - surveysAfter

    ' 5. Stamp the dashboard and bring it to the front
    With ThisWorkbook.Worksheets("Dashboard")
        .Range("H2").Value = Now
        .Activate
    End With

    ' Optional: save a PDF of the dashboard next to the workbook (the workbook must be saved first)
    ' ThisWorkbook.Worksheets("Dashboard").ExportAsFixedFormat Type:=xlTypePDF, _
    '     Filename:=ThisWorkbook.Path & Application.PathSeparator & "Dashboard_" & Format$(Now, "yyyy-mm-dd") & ".pdf"

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
