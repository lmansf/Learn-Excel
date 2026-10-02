'==============================================================================
' Lesson 5.5 - ThisWorkbook event code (SOLUTION - spoiler!)
'
' HOW TO USE: do NOT import this file (File > Import would create a separate new
' module, and events only run from the real ThisWorkbook module).
' In the VBE Project Explorer, double-click ThisWorkbook, delete anything already
' in its code window (a second Option Explicit won't compile), then paste
' everything below the dashed line. Needs WriteLog from modLog.
'------------------------------------------------------------------------------
Option Explicit

' Runs every time the workbook opens with macros enabled.
Private Sub Workbook_Open()
    WriteLog "Open", "Opened by " & Application.UserName
    Me.Worksheets("BedBoard").Activate
End Sub

' Runs just before every save (Ctrl + S, File > Save, Save As).
' Setting Cancel = True stops the save.
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Dim lo As ListObject
    Dim missing As Long

    ' Optional rule: refuse to save while required intake fields are blank.
    Set lo = Me.Worksheets("Intake").ListObjects("tblIntake")
    missing = Application.WorksheetFunction.CountBlank(lo.ListColumns("MRN").DataBodyRange) _
            + Application.WorksheetFunction.CountBlank(lo.ListColumns("ESILevel").DataBodyRange)

    If missing > 0 Then
        Cancel = True
        WriteLog "Save blocked", missing & " blank MRN/ESILevel cell(s) in tblIntake"
        MsgBox missing & " required Intake cell(s) are blank (MRN or ESILevel)." & vbLf & _
               "Fill them in, then save again.", vbExclamation, "Save cancelled"
        Exit Sub
    End If

    WriteLog "Save", IIf(SaveAsUI, "Save As dialog", "Save")
End Sub
