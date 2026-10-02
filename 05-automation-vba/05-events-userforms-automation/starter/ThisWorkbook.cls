'==============================================================================
' Lesson 5.5 - ThisWorkbook event code (STARTER) - practice task 6
'
' HOW TO USE: do NOT import this file (File > Import would create a separate new
' module, and events only run from the real ThisWorkbook module).
' In the VBE Project Explorer, double-click ThisWorkbook, delete anything already
' in its code window (a second Option Explicit won't compile), then paste
' everything below the dashed line. Needs WriteLog from modLog.
'
' Tip: instead of typing an event's first line, pick "Workbook" in the left
' dropdown at the top of the code window and the event in the right dropdown.
'------------------------------------------------------------------------------
Option Explicit

' Runs every time the workbook opens with macros enabled.
Private Sub Workbook_Open()
    ' TODO 1: log the open, e.g.  WriteLog "Open", "Opened by " & Application.UserName
    ' TODO 2 (optional): activate the BedBoard sheet
End Sub

' Runs just before every save. Setting Cancel = True stops the save.
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    ' TODO 3: log the save, e.g.  WriteLog "Save", "Save"
    ' TODO 4 (optional): count blank MRN/ESILevel cells in tblIntake; if there are any,
    '         show a message, log "Save blocked", set Cancel = True, and Exit Sub
    '         BEFORE the "Save" log line runs.
End Sub
