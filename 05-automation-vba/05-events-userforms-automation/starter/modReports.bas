Attribute VB_Name = "modReports"
'==============================================================================
' Lesson 5.5 - modReports (STARTER)
' Import with File > Import File... in the Visual Basic Editor, then complete
' each TODO. Needs WriteLog from modLog.
'
' CombineCensusFiles   practice task 11
' SaveTimestampedCopy  guide section 9 (optional)
' BuildCensusReport    bonus
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Practice task 11: every CSV in the census folder -> the Combined sheet
'------------------------------------------------------------------------------
Public Sub CombineCensusFiles()
    Dim sep As String, folder As String, fileName As String
    Dim wsOut As Worksheet, wbSrc As Workbook, body As Range
    Dim nextRow As Long, nFiles As Long

    sep = Application.PathSeparator               ' "\" on Windows, "/" on Mac
    ' The CSV folder sits next to this workbook; its name is in Settings!B4.
    folder = ThisWorkbook.Path & sep & ThisWorkbook.Worksheets("Settings").Range("B4").Value
    If Len(Dir(folder, vbDirectory)) = 0 Then
        MsgBox "Folder not found:" & vbLf & folder, vbExclamation
        Exit Sub
    End If
    folder = folder & sep

    Application.ScreenUpdating = False
    Set wsOut = ThisWorkbook.Worksheets("Combined")
    ' TODO 1: clear any old data from row 2 down (keep the headers in row 1)
    nextRow = 2

    fileName = Dir(folder)                        ' first file in the folder (any type)
    Do While Len(fileName) > 0
        ' TODO 2: only handle names that end in .csv   (hint: LCase$(fileName) Like "*.csv")
        ' TODO 3: Set wbSrc = Workbooks.Open(...)
        ' TODO 4: Set body = the file's data WITHOUT its header row
        '         (hint: Range("A1").CurrentRegion, then Offset(1, 0).Resize(rows - 1))
        ' TODO 5: copy body.Value to wsOut starting at row nextRow, then move nextRow down
        ' TODO 6: close wbSrc without saving and count the file in nFiles
        fileName = Dir()                          ' next file - keep this as the loop's last line
    Loop

    wsOut.Columns("A").NumberFormat = "mm/dd/yyyy"
    Application.ScreenUpdating = True
    WriteLog "Combine", nFiles & " files, " & (nextRow - 2) & " rows"
    ' Stretch goal: add On Error GoTo with a handler that closes wbSrc and
    ' restores ScreenUpdating (see the solution).
End Sub

'------------------------------------------------------------------------------
' Guide section 9 (optional): a timestamped backup with SaveCopyAs
'------------------------------------------------------------------------------
Public Sub SaveTimestampedCopy()
    ' TODO: build  <workbook folder>/Archive/CensusReport_yyyy-mm-dd_hhnn.xlsm
    '       (MkDir the Archive folder if Dir says it doesn't exist),
    '       call ThisWorkbook.SaveCopyAs with it, and WriteLog "Archive", <path>
End Sub

'------------------------------------------------------------------------------
' Bonus: one-click monthly census report (read the bonus brief in the README)
'------------------------------------------------------------------------------
Public Sub BuildCensusReport()
    ' Suggested steps:
    ' 1. Read ReportMonth from Settings!B3; work out the month's first day,
    '    last day (DateSerial(y, m + 1, 0)) and number of days.
    ' 2. Call CombineCensusFiles, then read Combined!A2:H<last row> into a Variant array.
    ' 3. For rows in the report month, total MidnightCensus (patient days) and
    '    StaffedBeds (bed days) per DeptID. UnitIndex below lets a Collection act
    '    as a DeptID -> array-position lookup that also works on a Mac.
    ' 4. Delete and re-create a sheet named "Report_" & Format(month, "yyyy-mm"):
    '    headers in A3:G3, one row per unit from row 4, sorted by Occupancy
    '    (highest first), then a "Total" row. Format ADC as 0.0, Occupancy as 0.0%.
    ' 5. ExportAsFixedFormat the sheet to <workbook folder>/Reports/CensusReport_yyyy-mm.pdf,
    '    write that full path into Settings!B6, and SaveCopyAs a timestamped backup.
End Sub

' Position stored for key in the Collection, or 0 if the key isn't there yet.
Private Function UnitIndex(ByVal idx As Collection, ByVal key As String) As Long
    On Error Resume Next
    UnitIndex = idx.Item(key)              ' error 5 when the key is missing -> stays 0
    On Error GoTo 0
End Function
