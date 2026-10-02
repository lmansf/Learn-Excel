'==============================================================================
' Lesson 5.5 - frmIntake code-behind (STARTER) - practice tasks 8-10
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

    ' TODO 1: load the chief complaints from Lists!C2 down into cboComplaint
    '         (assign the range's .Value to cboComplaint.List)

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
    lblStatus.Caption = "Saved: " & Trim$(txtLast.Value) & ", " & Trim$(txtFirst.Value)
    ClearForm
    txtMRN.SetFocus
End Sub

Private Sub cmdClose_Click()
    ' TODO 2: close and remove the form from memory
End Sub

'--- Validation: returns "" when everything is OK ------------------------------
Private Function ValidationMessage() As String
    Dim msg As String

    If Len(Trim$(txtLast.Value)) = 0 Then msg = msg & "- Last name is required." & vbLf
    ' TODO 3: add a check for each rule below, in the same style:
    '   - MRN must be exactly 8 digits (hint: Like "########")
    '   - First name is required
    '   - Arrival must be a valid date and time (hint: IsDate)
    '   - An arrival mode, an ESI level, and a chief complaint must be chosen
    '     (hint: ListIndex = -1 means nothing is selected; SelectedESI() = 0 means no ESI)

    ValidationMessage = msg
End Function

' Option buttons have no single "value", so ask each one in turn.
Private Function SelectedESI() As Long
    ' TODO 4: loop i = 1 To 5 over Me.Controls("optESI" & i);
    '         return i for the button whose .Value is True (0 if none)
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
    ' TODO 5: write the other fields with SetField, converting each to the right type:
    '   ArrivalDateTime  CDate(txtArrival.Value)
    '   MRN              set the cell's NumberFormat to "@" FIRST, then write the text
    '   LastName, FirstName, ArrivalMode, ChiefComplaint   the control values (trimmed)
    '   ESILevel         SelectedESI()  (a number)
    '   Interpreter      CBool(chkInterpreter.Value)
    '   EnteredBy        Application.UserName
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
