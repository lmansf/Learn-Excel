Attribute VB_Name = "ClaimDictionaries"
Option Explicit

' =====================================================================
' Lesson 5.4 - starter module for tasks 8-10 (Collections & Dictionaries)
'
' Each macro writes its answer to the Output sheet; the gray cells on the
' Practice sheet read it from there.
'   Windows: Set d = CreateObject("Scripting.Dictionary")
'   Mac:     Scripting.Dictionary doesn't exist. Use a Collection
'            (see "Collections and dictionaries on a Mac" in the guide).
' Claims sheet: A ClaimID, B PatientID, C PayerID, D ServiceDate,
'   E BilledAmount, F AllowedAmount, G PaidAmount, H ClaimStatus,
'   I DenialReason.    Payers sheet: A PayerID, B PayerName, C PayerType.
' =====================================================================

' Task 8 -> Output!B4: how many different (non-blank) denial reasons?
Public Sub CountDenialReasons()
    ' TODO: read the Claims sheet into an array (CurrentRegion.Value), add every
    '       non-blank DenialReason to a Collection with AddIfNew, then write
    '       the Collection's Count to Output!B4
End Sub

' Task 9 -> Output!B5: how many different patients have a claim?
Public Sub CountDistinctPatients()
    ' TODO: add every PatientID to a Dictionary (Mac: a Collection) once, then
    '       write the Count to Output!B5
End Sub

' Task 10 -> Output!B6 (PayerName) and Output!B7 (number of Denied claims)
Public Sub TopDeniedPayer()
    ' TODO 1: lookup dictionary PayerID -> PayerName from the Payers sheet
    ' TODO 2: counting dictionary PayerID -> number of rows with ClaimStatus "Denied"
    ' TODO 3: loop over the keys to find the largest count, then write the payer's
    '         name to Output!B6 and the count to Output!B7
End Sub

' ---------------------------------------------------------------------
' Helpers
' ---------------------------------------------------------------------

' Adds keyText to a Collection once. Returns False if it was already there.
Private Function AddIfNew(col As Collection, ByVal keyText As String) As Boolean
    ' TODO: trap the error a duplicate key raises (On Error Resume Next), add the
    '       key, set AddIfNew from Err.Number, then switch trapping off again
End Function

' Column number of a header in row 1. Raises a clear error if it is missing.
Private Function HeaderColumn(ws As Worksheet, ByVal headerText As String) As Long
    Dim pos As Variant
    pos = Application.Match(headerText, ws.Rows(1), 0)   ' an error VALUE, not a crash
    If IsError(pos) Then
        Err.Raise vbObjectError + 513, "HeaderColumn", _
                  "Column '" & headerText & "' was not found on sheet '" & ws.Name & "'."
    End If
    HeaderColumn = CLng(pos)
End Function
