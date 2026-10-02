Attribute VB_Name = "ClaimCollectionsMacSolution"
Option Explicit

' =====================================================================
' Lesson 5.4 - Mac-friendly solutions for tasks 8-10 (SPOILERS)
'
' Excel for Mac has no Scripting.Dictionary (CreateObject fails with
' run-time error 429), so these versions use only the built-in Collection.
' They also run on Windows. Remove your own ClaimDictionaries module first
' (or import into a spare copy) so two modules don't share macro names.
' =====================================================================

' Task 8 (same on Windows and Mac): how many different denial reasons? A Collection refuses a duplicate
' key, so adding every reason with Key:=reason keeps exactly one of each.
Public Sub CountDenialReasons()
    Dim ws As Worksheet, claims As Variant
    Dim reasonCol As Long, r As Long, reason As String
    Dim reasons As Collection

    Set ws = ThisWorkbook.Worksheets("Claims")
    claims = ws.Range("A1").CurrentRegion.Value        ' header + data in one 2-D array
    reasonCol = HeaderColumn(ws, "DenialReason")
    Set reasons = New Collection

    For r = 2 To UBound(claims, 1)
        reason = Trim$(CStr(claims(r, reasonCol)))     ' a blank cell becomes ""
        If Len(reason) > 0 Then AddIfNew reasons, reason
    Next r

    ThisWorkbook.Worksheets("Output").Range("B4").Value = reasons.Count
End Sub

' Task 9 (Mac): a keyed Collection keeps one entry per patient.
Public Sub CountDistinctPatients()
    Dim ws As Worksheet, claims As Variant
    Dim patientCol As Long, r As Long
    Dim patients As Collection

    Set ws = ThisWorkbook.Worksheets("Claims")
    claims = ws.Range("A1").CurrentRegion.Value
    patientCol = HeaderColumn(ws, "PatientID")
    Set patients = New Collection

    For r = 2 To UBound(claims, 1)
        AddIfNew patients, CStr(claims(r, patientCol))
    Next r

    ThisWorkbook.Worksheets("Output").Range("B5").Value = patients.Count
End Sub

' Task 10 (Mac): a Collection maps PayerID -> slot; an array holds the counts.
Public Sub TopDeniedPayer()
    Dim wsClaims As Worksheet, claims As Variant, payers As Variant
    Dim payerCol As Long, statusCol As Long, r As Long, i As Long
    Dim payerNames As Collection, slots As Collection
    Dim ids() As String, counts() As Long
    Dim bestID As String, bestCount As Long

    ' 1. Lookup Collection: Item = PayerName, Key = PayerID
    Set payerNames = New Collection
    payers = ThisWorkbook.Worksheets("Payers").Range("A1").CurrentRegion.Value
    For r = 2 To UBound(payers, 1)
        payerNames.Add Item:=payers(r, 2), Key:=CStr(payers(r, 1))
    Next r

    ' 2. Count Denied claims per payer in slot arrays
    Set wsClaims = ThisWorkbook.Worksheets("Claims")
    claims = wsClaims.Range("A1").CurrentRegion.Value
    payerCol = HeaderColumn(wsClaims, "PayerID")
    statusCol = HeaderColumn(wsClaims, "ClaimStatus")
    Set slots = New Collection
    ReDim ids(1 To UBound(claims, 1))
    ReDim counts(1 To UBound(claims, 1))
    For r = 2 To UBound(claims, 1)
        If claims(r, statusCol) = "Denied" Then
            i = SlotFor(slots, CStr(claims(r, payerCol)))
            ids(i) = CStr(claims(r, payerCol))
            counts(i) = counts(i) + 1
        End If
    Next r

    ' 3. Keep the largest count (on a tie, the payer seen first wins)
    For i = 1 To slots.Count
        If counts(i) > bestCount Then
            bestCount = counts(i)
            bestID = ids(i)
        End If
    Next i

    With ThisWorkbook.Worksheets("Output")
        .Range("B6").Value = LookupOr(payerNames, bestID, "(unknown payer " & bestID & ")")
        .Range("B7").Value = bestCount
    End With
End Sub

' ---------------------------------------------------------------------
' Helpers
' ---------------------------------------------------------------------

' Adds keyText to a Collection once. Returns False if it was already there.
Private Function AddIfNew(col As Collection, ByVal keyText As String) As Boolean
    On Error Resume Next
    col.Add Item:=keyText, Key:=keyText    ' error 457 if the key already exists
    AddIfNew = (Err.Number = 0)
    On Error GoTo 0
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

' Returns the slot number stored under keyText. The first time a key is seen it
' gets the next free slot (Count + 1). A Collection can't change an item, so
' the numbers you want to update live in arrays indexed by this slot.
Private Function SlotFor(slots As Collection, ByVal keyText As String) As Long
    On Error Resume Next
    SlotFor = slots(keyText)               ' error 5 if the key isn't there yet
    If Err.Number <> 0 Then
        slots.Add Item:=slots.Count + 1, Key:=keyText
        SlotFor = slots.Count
    End If
    On Error GoTo 0
End Function

' Collection lookup that returns a default instead of raising error 5.
Private Function LookupOr(col As Collection, ByVal keyText As String, ByVal fallback As String) As String
    LookupOr = fallback
    On Error Resume Next
    LookupOr = col(keyText)                ' leaves the fallback in place if the key is missing
    On Error GoTo 0
End Function
