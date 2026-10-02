Attribute VB_Name = "EncounterMacrosSolution"
Option Explicit

' =====================================================================
' Lesson 5.3 - reference solutions for Practice tasks 7-12 (SPOILERS)
'
' Import this into a spare copy of the workbook (or remove your own
' EncounterMacros module first) so the macro names don't clash.
' The helpers are Private here, so they never clash with yours.
'
' Encounters sheet (row 1 = headers, data from row 2, no blank rows):
'    1 A EncounterID      5 E DeptID              9 I PrimaryDxCode
'    2 B PatientID        6 F AdmitSource        10 J PayerID
'    3 C EncounterType    7 G AdmitDateTime      11 K TotalCharges
'    4 D FacilityID       8 H DischargeDateTime  12 L LOSDays (task 12 fills it)
'
' Facilities sheet: Excel Table tblFacilities (FacilityID, FacilityName, City, FacilityType)
' Output sheet: B4 and B5 SortAndReconcile, B6 and B7 ExportFacility,
'   B9 BuildMonthlyPackets (bonus, formatted as Text).
' =====================================================================

Private Function LastRow(ws As Worksheet, Optional col As Variant = 1) As Long
    ' The last filled row in one column: start at the bottom of the sheet and
    ' jump up, like Ctrl + Up arrow. col can be a letter ("A") or a number (1).
    LastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function

' Task 7, part 1
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

' Task 7, part 2
Sub BuildTypeSummary()
    Dim wsEnc As Worksheet, wsSum As Worksheet
    Dim rngType As Range, rngCharges As Range
    Dim types As Variant
    Dim lastR As Long, i As Long

    Set wsEnc = ThisWorkbook.Worksheets("Encounters")
    lastR = LastRow(wsEnc, "A")
    Set rngType = wsEnc.Range("C2:C" & lastR)          ' EncounterType
    Set rngCharges = wsEnc.Range("K2:K" & lastR)       ' TotalCharges

    ' Delete-and-recreate: the macro can run again and again
    DeleteSheetIfExists "TypeSummary"
    Set wsSum = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsSum.Name = "TypeSummary"

    wsSum.Range("A1:C1").Value = Array("EncounterType", "Encounters", "TotalCharges")
    types = Array("Emergency", "Inpatient", "Observation", "Outpatient")
    For i = 0 To UBound(types)                          ' Array() numbers its items from 0
        With wsSum.Cells(i + 2, 1)                      ' rows 2 to 5, column A
            .Value = types(i)
            .Offset(0, 1).Value = Application.WorksheetFunction.CountIf(rngType, types(i))
            .Offset(0, 2).Value = Application.WorksheetFunction.SumIf(rngType, types(i), rngCharges)
        End With
    Next i

    wsSum.Range("A1:C1").Font.Bold = True
    wsSum.Range("C2:C5").NumberFormat = "#,##0.00"
    wsSum.Columns("A:C").AutoFit
End Sub

' Task 8
Sub SplitByFacility()
    Dim wsEnc As Worksheet, wsNew As Worksheet
    Dim lo As ListObject
    Dim idCell As Range, block As Range
    Dim facID As String

    Set wsEnc = ThisWorkbook.Worksheets("Encounters")
    Set lo = ThisWorkbook.Worksheets("Facilities").ListObjects("tblFacilities")
    If wsEnc.AutoFilterMode Then wsEnc.AutoFilterMode = False   ' start with no filter
    Set block = wsEnc.Range("A1").CurrentRegion                 ' header + every data row

    Application.ScreenUpdating = False
    On Error GoTo CleanUp                                       ' always restore settings

    For Each idCell In lo.ListColumns("FacilityID").DataBodyRange
        facID = idCell.Value
        DeleteSheetIfExists facID
        Set wsNew = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsNew.Name = facID

        block.AutoFilter Field:=4, Criteria1:=facID               ' column D = FacilityID
        block.SpecialCells(xlCellTypeVisible).Copy Destination:=wsNew.Range("A1")
        wsNew.Range("A1").CurrentRegion.Columns.AutoFit
    Next idCell

CleanUp:
    wsEnc.AutoFilterMode = False                                ' leave the export unfiltered
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then MsgBox "SplitByFacility stopped: " & Err.Description, vbExclamation
End Sub

' Tasks 9 and 10
Sub SortAndReconcile()
    Dim ws As Worksheet
    Dim lastR As Long, rowTotal As Long, sheetCount As Long
    Dim chargeTotal As Double

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like "F0#" Then                  ' F01 to F09 (# = any one digit)
            ' Task 9: largest charge first. One cell is enough to name the sort column.
            ws.Range("A1").CurrentRegion.Sort Key1:=ws.Range("K1"), Order1:=xlDescending, Header:=xlYes

            ' Task 10: add this sheet to the control totals
            lastR = LastRow(ws, "A")
            rowTotal = rowTotal + (lastR - 1)       ' minus the header row
            chargeTotal = chargeTotal + Application.WorksheetFunction.Sum(ws.Range("K2:K" & lastR))
            sheetCount = sheetCount + 1
        End If
    Next ws

    With ThisWorkbook.Worksheets("Output")
        .Range("B4").Value = rowTotal
        .Range("B5").Value = chargeTotal
    End With
    MsgBox sheetCount & " facility sheets: " & rowTotal & " rows, " & _
           Format(chargeTotal, "#,##0.00") & " in charges.", vbInformation, "Reconcile the split"
End Sub

' Task 11
Sub ExportFacility()
    Const FAC_ID As String = "F02"
    Dim wbNew As Workbook, wbCheck As Workbook
    Dim filePath As String, savedName As String
    Dim lastR As Long
    Dim total As Double

    If Not SheetExists(FAC_ID) Then
        MsgBox "There's no " & FAC_ID & " sheet. Run SplitByFacility first.", vbExclamation
        Exit Sub
    End If
    If ThisWorkbook.Path = "" Then
        MsgBox "Save this workbook first, so the export has a folder to go in.", vbExclamation
        Exit Sub
    End If
    filePath = ThisWorkbook.Path & Application.PathSeparator & FAC_ID & "_encounters_2025.xlsx"

    ' 1. Copy the sheet into a brand-new workbook and save that as .xlsx
    ThisWorkbook.Worksheets(FAC_ID).Copy            ' no Before/After: Excel makes a new workbook
    Set wbNew = ActiveWorkbook                      ' right after Copy, the new workbook is active
    Application.DisplayAlerts = False               ' replace last run's file without asking
    wbNew.SaveAs Filename:=filePath, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    savedName = wbNew.Name                          ' read it BEFORE closing
    wbNew.Close SaveChanges:=False

    ' 2. Reopen the saved file, read it, and close it again
    Set wbCheck = Workbooks.Open(Filename:=filePath, ReadOnly:=True)
    With wbCheck.Worksheets(1)
        lastR = .Cells(.Rows.Count, "A").End(xlUp).Row
        total = Application.WorksheetFunction.Sum(.Range("K2:K" & lastR))
    End With
    wbCheck.Close SaveChanges:=False

    ' 3. Report in THIS workbook, whichever workbook is active now
    With ThisWorkbook.Worksheets("Output")
        .Range("B6").Value = savedName
        .Range("B7").Value = total
    End With
End Sub

' Task 12
Sub FillLOSDays()
    Dim ws As Worksheet
    Dim lastR As Long, i As Long
    Dim stay As Variant          ' will hold an n x 2 array: admit, discharge
    Dim los() As Variant         ' n x 1 array for the results

    Set ws = ThisWorkbook.Worksheets("Encounters")
    lastR = LastRow(ws, "A")

    stay = ws.Range("G2:H" & lastR).Value2          ' ONE read. Value2 gives dates as serial numbers
    ReDim los(1 To UBound(stay, 1), 1 To 1)

    For i = 1 To UBound(stay, 1)
        ' whole days between the two dates = midnights in the hospital
        los(i, 1) = Int(stay(i, 2)) - Int(stay(i, 1))
    Next i

    ws.Range("L2").Resize(UBound(los, 1), 1).Value = los   ' ONE write
End Sub
