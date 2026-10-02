Attribute VB_Name = "modReports"
'==============================================================================
' Lesson 5.5 - modReports (SOLUTION - spoiler!)
' Import with File > Import File... in the Visual Basic Editor.
'
' CombineCensusFiles   practice task 11: every CSV in the census folder -> Combined sheet
' SaveTimestampedCopy  guide section 9: a dated backup copy with SaveCopyAs
' BuildCensusReport    bonus: combine, summarize by unit, export PDF, back up
'
' Works on Windows and Mac: paths use Application.PathSeparator, and the Dir
' loop filters names with Like instead of a "*.csv" pattern (Mac's Dir ignores
' wildcards). Needs WriteLog from modLog.
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Practice task 11
'------------------------------------------------------------------------------
Public Sub CombineCensusFiles()
    Dim sep As String, folder As String, fileName As String
    Dim wsOut As Worksheet, wbSrc As Workbook, body As Range
    Dim nextRow As Long, nFiles As Long

    sep = Application.PathSeparator
    If Len(ThisWorkbook.Path) = 0 Or LCase$(Left$(ThisWorkbook.Path, 4)) = "http" Then
        MsgBox "Save this workbook in a folder on your computer first " & _
               "(a OneDrive/SharePoint web path won't work with Dir).", vbExclamation
        Exit Sub
    End If

    ' The CSV folder sits next to this workbook; its name is in Settings!B4.
    folder = ThisWorkbook.Path & sep & ThisWorkbook.Worksheets("Settings").Range("B4").Value
    If Len(Dir(folder, vbDirectory)) = 0 Then
        MsgBox "Folder not found:" & vbLf & folder, vbExclamation
        Exit Sub
    End If
    folder = folder & sep

    On Error GoTo Fail
    Application.ScreenUpdating = False
    Set wsOut = ThisWorkbook.Worksheets("Combined")
    wsOut.Range("A2:H" & wsOut.Rows.Count).ClearContents     ' keep the headers in row 1
    nextRow = 2

    fileName = Dir(folder)                       ' first file in the folder (any type)
    Do While Len(fileName) > 0
        If LCase$(fileName) Like "*.csv" Then    ' skip README.txt and anything else
            Set wbSrc = Workbooks.Open(Filename:=folder & fileName)
            Set body = wbSrc.Worksheets(1).Range("A1").CurrentRegion
            If body.Rows.Count > 1 Then
                Set body = body.Offset(1, 0).Resize(body.Rows.Count - 1)    ' drop the header row
                wsOut.Cells(nextRow, 1).Resize(body.Rows.Count, body.Columns.Count).Value = body.Value
                nextRow = nextRow + body.Rows.Count
            End If
            wbSrc.Close SaveChanges:=False
            Set wbSrc = Nothing
            nFiles = nFiles + 1
        End If
        fileName = Dir()                         ' next file: no arguments!
    Loop

    wsOut.Columns("A").NumberFormat = "mm/dd/yyyy"
    WriteLog "Combine", nFiles & " files, " & (nextRow - 2) & " rows"

CleanUp:
    Application.ScreenUpdating = True
    Exit Sub

Fail:
    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False   ' never leave a source file open
    MsgBox "Combine failed on " & fileName & ":" & vbLf & Err.Description, vbExclamation
    Resume CleanUp
End Sub

'------------------------------------------------------------------------------
' Guide section 9: a timestamped backup. The workbook you're in stays open
' under its own name; the copy is written to disk only.
'------------------------------------------------------------------------------
Public Sub SaveTimestampedCopy()
    Dim sep As String, folder As String, ext As String, copyPath As String

    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "Save the workbook first.", vbExclamation
        Exit Sub
    End If
    sep = Application.PathSeparator
    folder = ThisWorkbook.Path & sep & "Archive"
    If Len(Dir(folder, vbDirectory)) = 0 Then MkDir folder

    ext = Mid$(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, "."))     ' ".xlsm" - a copy keeps the format
    copyPath = folder & sep & "CensusReport_" & Format(Now, "yyyy-mm-dd_hhnn") & ext
    ThisWorkbook.SaveCopyAs copyPath
    WriteLog "Archive", copyPath
End Sub

'------------------------------------------------------------------------------
' Bonus: one-click monthly census report
'------------------------------------------------------------------------------
Public Sub BuildCensusReport()
    Const FIRST_ROW As Long = 4                  ' first unit row on the report
    Dim wsSet As Worksheet, wsData As Worksheet, wsRep As Worksheet
    Dim reportMonth As Date, monthEnd As Date, nDays As Long
    Dim data As Variant, r As Long, i As Long, k As Long, n As Long
    Dim unitIdx As Collection
    Dim ids() As String, unitNames() As String, facs() As String
    Dim pd() As Double, bd() As Double, totPD As Double, totBD As Double
    Dim sheetName As String, sep As String, folder As String, pdfPath As String, ext As String
    Dim lastRow As Long, totalRow As Long

    On Error GoTo Fail
    Set wsSet = ThisWorkbook.Worksheets("Settings")
    If Not IsDate(wsSet.Range("B3").Value) Then Err.Raise vbObjectError + 513, , "Settings!B3 (ReportMonth) must be a date."
    reportMonth = DateSerial(Year(wsSet.Range("B3").Value), Month(wsSet.Range("B3").Value), 1)
    monthEnd = DateSerial(Year(reportMonth), Month(reportMonth) + 1, 0)   ' day 0 of next month = last day
    nDays = Day(monthEnd)

    ' Step 1: refresh the Combined sheet from the CSV folder
    CombineCensusFiles
    Set wsData = ThisWorkbook.Worksheets("Combined")
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Err.Raise vbObjectError + 514, , "The Combined sheet is empty. Check the CSV folder."
    data = wsData.Range("A2:H" & lastRow).Value          ' one fast read into memory

    ' Step 2: total patient days and bed days per unit for the report month.
    ' unitIdx maps DeptID -> position in the arrays (a Collection works on Mac too).
    Set unitIdx = New Collection
    For r = 1 To UBound(data, 1)
        If data(r, 1) >= reportMonth And data(r, 1) <= monthEnd Then
            k = UnitIndex(unitIdx, CStr(data(r, 3)))
            If k = 0 Then                                  ' first row for this unit
                n = n + 1
                ReDim Preserve ids(1 To n): ReDim Preserve unitNames(1 To n): ReDim Preserve facs(1 To n)
                ReDim Preserve pd(1 To n): ReDim Preserve bd(1 To n)
                ids(n) = CStr(data(r, 3))
                unitNames(n) = CStr(data(r, 4))
                facs(n) = CStr(data(r, 2))
                unitIdx.Add n, ids(n)
                k = n
            End If
            pd(k) = pd(k) + data(r, 8)      ' MidnightCensus -> patient days
            bd(k) = bd(k) + data(r, 5)      ' StaffedBeds    -> bed days
        End If
    Next r
    If n = 0 Then Err.Raise vbObjectError + 515, , "No census rows for " & Format(reportMonth, "mmmm yyyy") & "."

    ' Step 3: rebuild this month's report sheet
    sheetName = "Report_" & Format(reportMonth, "yyyy-mm")
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(sheetName).Delete              ' no error if it doesn't exist yet
    On Error GoTo Fail
    Application.DisplayAlerts = True
    Set wsRep = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsRep.Name = sheetName

    With wsRep
        .Range("A1").Value = "Bluestone Health System - Inpatient Census Report - " & Format(reportMonth, "mmmm yyyy")
        .Range("A3:G3").Value = Array("DeptID", "Unit", "Facility", "PatientDays", "BedDays", "ADC", "Occupancy")
        For i = 1 To n
            .Cells(FIRST_ROW + i - 1, 1).Resize(1, 7).Value = _
                Array(ids(i), unitNames(i), facs(i), pd(i), bd(i), pd(i) / nDays, pd(i) / bd(i))
            totPD = totPD + pd(i)
            totBD = totBD + bd(i)
        Next i
        .Range("A" & FIRST_ROW).Resize(n, 7).Sort Key1:=.Range("G" & FIRST_ROW), Order1:=xlDescending, Header:=xlNo
        totalRow = FIRST_ROW + n
        .Cells(totalRow, 1).Resize(1, 7).Value = _
            Array("Total", "All units", "", totPD, totBD, totPD / nDays, totPD / totBD)

        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A3:G3").Font.Bold = True
        .Range("A" & totalRow & ":G" & totalRow).Font.Bold = True
        .Range("D4:E" & totalRow).NumberFormat = "#,##0"
        .Range("F4:F" & totalRow).NumberFormat = "0.0"
        .Range("G4:G" & totalRow).NumberFormat = "0.0%"
        .Range("A3:G" & totalRow).Columns.AutoFit
        With .PageSetup
            .Orientation = xlLandscape
            .Zoom = False                  ' required before FitToPages* takes effect
            .FitToPagesWide = 1
            .FitToPagesTall = False
            .CenterFooter = "Page &P of &N"
        End With
    End With

    ' Step 4: export the report sheet to PDF and record where it went
    sep = Application.PathSeparator
    folder = ThisWorkbook.Path & sep & wsSet.Range("B5").Value
    If Len(Dir(folder, vbDirectory)) = 0 Then MkDir folder
    pdfPath = folder & sep & "CensusReport_" & Format(reportMonth, "yyyy-mm") & ".pdf"
    wsRep.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    wsSet.Range("B6").Value = pdfPath

    ' Step 5: timestamped backup of the whole workbook (you stay in the original)
    ext = Mid$(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, "."))
    ThisWorkbook.SaveCopyAs folder & sep & "CensusReport_" & Format(reportMonth, "yyyy-mm") & _
                            "_built_" & Format(Now, "yyyy-mm-dd_hhnn") & ext

    WriteLog "Report", sheetName & " -> " & pdfPath
    wsRep.Activate
    MsgBox "Report ready:" & vbLf & pdfPath, vbInformation

CleanUp:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

Fail:
    MsgBox "Report failed: " & Err.Description, vbExclamation
    WriteLog "Report error", Err.Description
    Resume CleanUp
End Sub

' Position stored for key in the Collection, or 0 if the key isn't there yet.
Private Function UnitIndex(ByVal idx As Collection, ByVal key As String) As Long
    On Error Resume Next
    UnitIndex = idx.Item(key)              ' error 5 when the key is missing -> stays 0
    On Error GoTo 0
End Function
