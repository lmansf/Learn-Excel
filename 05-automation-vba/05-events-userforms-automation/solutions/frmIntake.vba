'==============================================================================
' Lesson 5.5 - frmIntake code-behind (SOLUTION - spoiler!)
'
' HOW TO USE: this is not an importable file. First build the form
' (Insert > UserForm, name it frmIntake, and add the controls listed in the
' lesson guide, section 7). Then double-click the form, delete anything in its
' code window, and paste everything below the dashed line.
'
' Controls this code expects:
'   txtMRN, txtLast, txtFirst, txtArrival          TextBoxes
'   cboArrivalMode, cboComplaint                    ComboBoxes (Style = 2 - fmStyleDropDownList)
'   optESI1 ... optESI5                             OptionButtons inside one Frame
'   chkInterpreter                                  CheckBox
'   cmdSave, cmdClose                               CommandButtons
'   lblStatus                                       Label
'------------------------------------------------------------------------------
Option Explicit

Private Const TABLE_SHEET As String = "Intake"
Private Const TABLE_NAME As String = "tblIntake"

'--- Runs once, just before the form appears ---------------------------------
Private Sub UserForm_Initialize()
    Dim wsLists As Worksheet
    Dim c As Range
    Dim complaints As Range

    Set wsLists = ThisWorkbook.Worksheets("Lists")

    ' Arrival mode: a short list, added one item at a time with AddItem
    For Each c In wsLists.Range("A2", wsLists.Cells(wsLists.Rows.Count, "A").End(xlUp))
        cboArrivalMode.AddItem c.Value
    Next c

    ' Chief complaint: a longer list, loaded in one step by assigning the
    ' cells' values to .List (always reads THIS workbook, unlike RowSource)
    Set complaints = wsLists.Range("C2", wsLists.Cells(wsLists.Rows.Count, "C").End(xlUp))
    cboComplaint.List = complaints.Value

    lblStatus.Caption = ""
    ClearForm
End Sub

'--- Buttons ------------------------------------------------------------------
Private Sub cmdSave_Click()
    Dim msg As String

    msg = ValidationMessage()
    If Len(msg) > 0 Then
        MsgBox "Please fix the following:" & vbLf & vbLf & msg, vbExclamation, "Intake not saved"
        Exit Sub                                   ' keep the form open so the user can fix it
    End If

    AddIntakeRow
    lblStatus.Caption = "Saved: " & Trim$(txtLast.Value) & ", " & Trim$(txtFirst.Value) & _
                        " (MRN " & Trim$(txtMRN.Value) & ")"
    ClearForm                                      ' ready for the next patient
    txtMRN.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

'--- Validation: returns "" when everything is OK ------------------------------
Private Function ValidationMessage() As String
    Dim msg As String

    If Not Trim$(txtMRN.Value) Like "########" Then msg = msg & "- MRN must be exactly 8 digits (keep the leading zeros)." & vbLf
    If Len(Trim$(txtLast.Value)) = 0 Then msg = msg & "- Last name is required." & vbLf
    If Len(Trim$(txtFirst.Value)) = 0 Then msg = msg & "- First name is required." & vbLf
    If Not IsDate(txtArrival.Value) Then msg = msg & "- Arrival must be a date and time, e.g. 2025-12-31 11:41." & vbLf
    If cboArrivalMode.ListIndex = -1 Then msg = msg & "- Choose an arrival mode." & vbLf
    If SelectedESI() = 0 Then msg = msg & "- Choose an ESI level (1-5)." & vbLf
    If cboComplaint.ListIndex = -1 Then msg = msg & "- Choose a chief complaint." & vbLf

    ValidationMessage = msg
End Function

' Option buttons have no single "value", so ask each one in turn.
Private Function SelectedESI() As Long
    Dim i As Long
    For i = 1 To 5
        If Me.Controls("optESI" & i).Value = True Then
            SelectedESI = i
            Exit Function
        End If
    Next i
    ' none selected: the function returns 0
End Function

'--- Write one new row to the Table -------------------------------------------
Private Sub AddIntakeRow()
    Dim lo As ListObject
    Dim newRow As ListRow
    Dim nextID As Long

    Set lo = ThisWorkbook.Worksheets(TABLE_SHEET).ListObjects(TABLE_NAME)
    nextID = Application.WorksheetFunction.Max(lo.ListColumns("IntakeID").DataBodyRange) + 1

    Set newRow = lo.ListRows.Add                   ' new empty row at the bottom; the Table grows

    SetField lo, newRow, "IntakeID", nextID
    SetField lo, newRow, "ArrivalDateTime", CDate(txtArrival.Value)     ' a real date-time, not text
    newRow.Range.Cells(1, lo.ListColumns("MRN").Index).NumberFormat = "@"   ' Text format first...
    SetField lo, newRow, "MRN", Trim$(txtMRN.Value)                     ' ...so leading zeros survive
    SetField lo, newRow, "LastName", Trim$(txtLast.Value)
    SetField lo, newRow, "FirstName", Trim$(txtFirst.Value)
    SetField lo, newRow, "ArrivalMode", cboArrivalMode.Value
    SetField lo, newRow, "ESILevel", SelectedESI()                      ' a number (Long)
    SetField lo, newRow, "ChiefComplaint", cboComplaint.Value
    SetField lo, newRow, "Interpreter", CBool(chkInterpreter.Value)     ' TRUE / FALSE
    SetField lo, newRow, "EnteredBy", Application.UserName
End Sub

' Writes one value into the named column of a ListRow.
Private Sub SetField(ByVal lo As ListObject, ByVal r As ListRow, ByVal columnName As String, ByVal v As Variant)
    r.Range.Cells(1, lo.ListColumns(columnName).Index).Value = v
End Sub

'--- Reset the controls for the next patient ----------------------------------
Private Sub ClearForm()
    Dim i As Long

    txtMRN.Value = ""
    txtLast.Value = ""
    txtFirst.Value = ""
    txtArrival.Value = Format(Now, "yyyy-mm-dd hh:nn")
    cboArrivalMode.ListIndex = -1
    cboComplaint.ListIndex = -1
    For i = 1 To 5
        Me.Controls("optESI" & i).Value = False
    Next i
    chkInterpreter.Value = False
End Sub
