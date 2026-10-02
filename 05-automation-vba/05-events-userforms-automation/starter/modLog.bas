Attribute VB_Name = "modLog"
'==============================================================================
' Lesson 5.5 - modLog (STARTER - this helper module is complete; just import it)
' Import with File > Import File... in the Visual Basic Editor.
'
' WriteLog  appends one row to the Log sheet: Timestamp | Event | Detail | User
' EventsOn  one-click repair when your event handlers stop firing
'==============================================================================
Option Explicit

Public Sub WriteLog(ByVal eventName As String, ByVal detail As String)
    Dim ws As Worksheet
    Dim nextRow As Long

    Set ws = ThisWorkbook.Worksheets("Log")
    nextRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row + 1   ' first empty row under the Event column
    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 1).NumberFormat = "mm/dd/yyyy hh:mm:ss"
    ws.Cells(nextRow, 2).Value = eventName
    ws.Cells(nextRow, 3).Value = detail
    ws.Cells(nextRow, 4).Value = Application.UserName
End Sub

' Run this (Alt+F8 > EventsOn > Run) if an error left events switched off.
Public Sub EventsOn()
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Events, screen updating and alerts are switched back on.", vbInformation
End Sub
