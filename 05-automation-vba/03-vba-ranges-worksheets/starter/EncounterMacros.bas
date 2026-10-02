Attribute VB_Name = "EncounterMacros"
Option Explicit

' =====================================================================
' Lesson 5.3 - VBA: Ranges, Worksheets & Workbooks - starter module
' (Bluestone Health System, fictional data)
'
' Import it: in the Visual Basic Editor choose File > Import File... and
' pick this file. Complete each TODO, then run the macro with F5 (cursor
' inside the Sub) or from Developer > Macros. Save before every run:
' Undo can't undo a macro.
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

' Helper (complete). Use it in your macros: lastR = LastRow(ws, "A")
Function LastRow(ws As Worksheet, Optional col As Variant = 1) As Long
    ' The last filled row in one column: start at the bottom of the sheet and
    ' jump up, like Ctrl + Up arrow. col can be a letter ("A") or a number (1).
    LastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function

' Task 7, part 1
Function SheetExists(ByVal sheetName As String) As Boolean
    ' TODO: loop over ThisWorkbook.Worksheets with For Each.
    ' When a sheet's name matches, ignoring case:
    '     If StrComp(ws.Name, sheetName, vbTextCompare) = 0 Then
    ' set SheetExists = True and leave with Exit Function.
    ' (If the loop ends without a match, the function returns False.)
End Function

' Helper (complete). It calls your SheetExists, so it only works once
' SheetExists does.
Sub DeleteSheetIfExists(ByVal sheetName As String)
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
    ' TODO:
    ' 1. DeleteSheetIfExists "TypeSummary"
    ' 2. Add a sheet at the end of the workbook and name it TypeSummary:
    '      Set wsSum = ThisWorkbook.Worksheets.Add( _
    '          After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ' 3. Headers in A1:C1: EncounterType, Encounters, TotalCharges
    ' 4. For each type in Array("Emergency", "Inpatient", "Observation", "Outpatient")
    '    write the type, its count, and its total charges in rows 2-5.
    '    (Array() numbers its items from 0: see guide section 10.)
    '    Application.WorksheetFunction.CountIf(rngType, t) counts,
    '    Application.WorksheetFunction.SumIf(rngType, t, rngCharges) totals.
End Sub

' Task 8
Sub SplitByFacility()
    ' TODO: for each FacilityID cell in the Facilities table:
    '     For Each idCell In ThisWorkbook.Worksheets("Facilities") _
    '         .ListObjects("tblFacilities").ListColumns("FacilityID").DataBodyRange
    ' 1. DeleteSheetIfExists, then add a sheet at the end named after the ID.
    ' 2. Filter the Encounters block (Range("A1").CurrentRegion) on
    '    column 4 (FacilityID) with .AutoFilter Field:=4, Criteria1:=...
    ' 3. Copy the visible cells, header included, to A1 of the new sheet:
    '    .SpecialCells(xlCellTypeVisible).Copy Destination:=...
    ' After the loop, turn the filter off: wsEnc.AutoFilterMode = False
    ' Bonus points: ScreenUpdating off while it runs, and back on at the end.
End Sub

' Tasks 9 and 10
Sub SortAndReconcile()
    ' TODO: loop over every worksheet with For Each ws In ThisWorkbook.Worksheets.
    ' For each sheet whose name is Like "F0#":
    '   Task 9:  sort its block (Range("A1").CurrentRegion) by TotalCharges
    '            (column K), largest first, with Header:=xlYes.
    '   Task 10: add its data rows and its TotalCharges to two running totals.
    ' After the loop, write the row total to Output!B4 and the charge total
    ' to Output!B5.
End Sub

' Task 11
Sub ExportFacility()
    ' TODO:
    ' 1. Build the path: ThisWorkbook.Path & Application.PathSeparator & _
    '                    "F02_encounters_2025.xlsx"
    ' 2. ThisWorkbook.Worksheets("F02").Copy   (no arguments: a new workbook)
    '    Set wbNew = ActiveWorkbook
    ' 3. wbNew.SaveAs Filename:=..., FileFormat:=xlOpenXMLWorkbook
    '    (DisplayAlerts False around it, so a second run can overwrite),
    '    then wbNew.Close SaveChanges:=False
    ' 4. Set wbCheck = Workbooks.Open(Filename:=..., ReadOnly:=True),
    '    add up its TotalCharges (column K), and close it without saving.
    ' 5. Write the file name to Output!B6 and the total to Output!B7.
End Sub

' Task 12
Sub FillLOSDays()
    ' TODO:
    ' 1. Read G2:H<lastRow> (AdmitDateTime, DischargeDateTime) into a Variant
    '    with ONE .Value2 read. Value2 gives each date as a serial number.
    ' 2. ReDim a results array: ReDim los(1 To UBound(stay, 1), 1 To 1)
    ' 3. For each row: los(i, 1) = Int(discharge) - Int(admit)
    ' 4. Write the whole array to L2 with ONE assignment:
    '    ws.Range("L2").Resize(UBound(los, 1), 1).Value = los
End Sub

' Bonus - see the Bonus sheet for the full brief
Sub BuildMonthlyPackets()
    ' TODO (bonus): for m = 1 To 12
    ' 1. DeleteSheetIfExists Format(DateSerial(2025, m, 1), "yyyy-mm"), then
    '    copy PacketTemplate to the end and rename the copy to that name.
    ' 2. Collect the month's Inpatient rows (AdmitDateTime in the month) from
    '    an array you read ONCE before the loop, and write them from A7.
    ' 3. Sort the block at A6 by TotalCharges (K), largest first.
    ' 4. B2 = first day of the month, B3 = encounters, B4 = total charges.
    ' After the loop, write the name of the sheet with the largest total
    ' charges to Output!B9. ScreenUpdating off, and always back on.
End Sub
