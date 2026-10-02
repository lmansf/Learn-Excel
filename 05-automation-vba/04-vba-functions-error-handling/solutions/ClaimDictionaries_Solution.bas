Attribute VB_Name = "ClaimDictionariesSolution"
Option Explicit

' =====================================================================
' Lesson 5.4 - reference solutions for tasks 8-10 (SPOILERS)
'
' Windows: uses Scripting.Dictionary through late binding, so you don't
' need a Tools > References setting. Excel for Mac has no
' Scripting.Dictionary: use ClaimCollectionsMac_Solution.bas instead.
' Remove your own ClaimDictionaries module first (or import into a spare
' copy) so two modules don't define the same macro names.
'
' Claims sheet: A ClaimID, B PatientID, C PayerID, D ServiceDate,
'   E BilledAmount, F AllowedAmount, G PaidAmount, H ClaimStatus,
'   I DenialReason.    Payers sheet: A PayerID, B PayerName, C PayerType.
' Results go to the Output sheet: B4 (task 8), B5 (task 9), B6:B7 (task 10).
' =====================================================================

' Task 8: how many different denial reasons? A Collection refuses a duplicate
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

' Task 9: how many different patients have a claim? Dictionary keys are unique.
Public Sub CountDistinctPatients()
    Dim ws As Worksheet, claims As Variant
    Dim patientCol As Long, r As Long
    Dim patients As Object                             ' late-bound Scripting.Dictionary

    Set ws = ThisWorkbook.Worksheets("Claims")
    claims = ws.Range("A1").CurrentRegion.Value
    patientCol = HeaderColumn(ws, "PatientID")
    Set patients = CreateObject("Scripting.Dictionary")

    For r = 2 To UBound(claims, 1)
        If Not patients.Exists(claims(r, patientCol)) Then
            patients.Add claims(r, patientCol), 1
        End If
    Next r

    ThisWorkbook.Worksheets("Output").Range("B5").Value = patients.Count
End Sub

' Task 10: which payer has the most Denied claims? Two dictionaries:
' one to look up payer names, one to count.
Public Sub TopDeniedPayer()
    Dim wsClaims As Worksheet, claims As Variant, payers As Variant
    Dim payerCol As Long, statusCol As Long, r As Long
    Dim payerNames As Object, deniedCount As Object
    Dim payerID As Variant, bestID As String, bestCount As Long

    ' 1. Lookup dictionary: PayerID -> PayerName
    Set payerNames = CreateObject("Scripting.Dictionary")
    payers = ThisWorkbook.Worksheets("Payers").Range("A1").CurrentRegion.Value
    For r = 2 To UBound(payers, 1)
        payerNames(payers(r, 1)) = payers(r, 2)        ' assigning to a new key adds it
    Next r

    ' 2. Counting dictionary: PayerID -> number of Denied claims
    Set wsClaims = ThisWorkbook.Worksheets("Claims")
    claims = wsClaims.Range("A1").CurrentRegion.Value
    payerCol = HeaderColumn(wsClaims, "PayerID")
    statusCol = HeaderColumn(wsClaims, "ClaimStatus")
    Set deniedCount = CreateObject("Scripting.Dictionary")
    For r = 2 To UBound(claims, 1)
        If claims(r, statusCol) = "Denied" Then
            ' Reading a key that isn't there returns Empty (and adds it); Empty + 1 = 1
            deniedCount(claims(r, payerCol)) = deniedCount(claims(r, payerCol)) + 1
        End If
    Next r

    ' 3. Keep the largest count (on a tie, the payer seen first wins)
    For Each payerID In deniedCount.Keys
        If deniedCount(payerID) > bestCount Then
            bestCount = deniedCount(payerID)
            bestID = payerID
        End If
    Next payerID

    With ThisWorkbook.Worksheets("Output")
        .Range("B6").Value = payerNames(bestID)
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
