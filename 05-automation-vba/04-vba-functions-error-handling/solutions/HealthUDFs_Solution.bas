Attribute VB_Name = "HealthUDFsSolution"
Option Explicit

' =====================================================================
' Lesson 5.4 - reference solutions for the UDF tasks 1-6 (SPOILERS)
'
' Import this into a spare copy of the workbook, or remove your own
' HealthUDFs module first (right-click it > Remove HealthUDFs), so two
' modules don't define the same function names.
'
' Worksheet functions in this module:
'   =BMI(weightLb, heightIn)                    body-mass index
'   =AGEAT(dob, [asOf])                         age in completed years
'   =LOSDAYS(admitDateTime, dischargeDateTime)  length of stay (midnights)
'   =DENIALRATE(statusRange, [statusToCount])   share of claims with a status
' UDFs must live in a standard module (Insert > Module), not in a sheet
' module or ThisWorkbook; otherwise the cell shows #NAME?.
' =====================================================================

' Tasks 1-2: body-mass index = 703 x weight (lb) / height (in) ^ 2
Public Function BMI(weightLb As Variant, heightIn As Variant) As Variant
    Dim w As Variant, h As Variant
    w = ArgValue(weightLb)
    h = ArgValue(heightIn)
    If Not IsNumberValue(w) Or Not IsNumberValue(h) Then
        BMI = CVErr(xlErrValue)            ' blank or text input: #VALUE!
    ElseIf w <= 0 Or h <= 0 Then
        BMI = CVErr(xlErrNum)              ' impossible measurement: #NUM!
    Else
        BMI = 703 * w / h ^ 2
    End If
End Function

' Task 3: age in completed years on asOf. Leave asOf out to use today's date.
Public Function AGEAT(dob As Variant, Optional asOf As Variant) As Variant
    Dim birthValue As Variant, refValue As Variant
    Dim birth As Date, refDate As Date, years As Long

    birthValue = ArgValue(dob)
    If IsMissing(asOf) Then
        refValue = Date                    ' second argument left out: today
        Application.Volatile               ' today changes, so recalculate every time
    Else
        refValue = ArgValue(asOf)
    End If
    If Not IsNumberValue(birthValue) Or Not IsNumberValue(refValue) Then
        AGEAT = CVErr(xlErrValue)
        Exit Function
    End If

    birth = CDate(birthValue)
    refDate = CDate(refValue)
    If refDate < birth Then
        AGEAT = CVErr(xlErrNum)            ' asOf is before the date of birth
        Exit Function
    End If

    years = Year(refDate) - Year(birth)
    ' Birthday still to come in the asOf year? Then the person is a year younger.
    If DateSerial(Year(refDate), Month(birth), Day(birth)) > refDate Then years = years - 1
    AGEAT = years
End Function

' Tasks 4-5: length of stay = midnights between admission and discharge.
' #VALUE! when an input isn't a date or the discharge is before the admission.
Public Function LOSDAYS(admitDateTime As Variant, dischargeDateTime As Variant) As Variant
    Dim a As Variant, d As Variant
    a = ArgValue(admitDateTime)
    d = ArgValue(dischargeDateTime)
    If Not IsNumberValue(a) Or Not IsNumberValue(d) Then
        LOSDAYS = CVErr(xlErrValue)
    ElseIf CDate(d) < CDate(a) Then
        LOSDAYS = CVErr(xlErrValue)        ' discharge before admission: a data-entry error
    Else
        LOSDAYS = DateDiff("d", CDate(a), CDate(d))   ' "d" counts the midnights crossed
    End If
End Function

' Task 6: share of non-blank cells in statusRange that equal statusToCount.
Public Function DENIALRATE(statusRange As Range, Optional statusToCount As String = "Denied") As Variant
    Dim cell As Range, total As Long, hits As Long
    For Each cell In statusRange.Cells
        If Not IsEmpty(cell.Value) And Not IsError(cell.Value) Then
            total = total + 1
            If StrComp(Trim$(CStr(cell.Value)), statusToCount, vbTextCompare) = 0 Then hits = hits + 1
        End If
    Next cell
    If total = 0 Then
        DENIALRATE = CVErr(xlErrDiv0)      ' nothing to divide by: #DIV/0!
    Else
        DENIALRATE = hits / total
    End If
End Function

' ---------------------------------------------------------------------
' Helpers. Private = code in other modules can't call them, and they
' stay out of AutoComplete and the Insert Function list.
' ---------------------------------------------------------------------

' A cell reference arrives as a Range object; a typed number or a formula
' result arrives as a plain value. Return the value either way.
Private Function ArgValue(arg As Variant) As Variant
    If IsObject(arg) Then
        ArgValue = arg.Cells(1, 1).Value
    Else
        ArgValue = arg
    End If
End Function

' True for real numbers and dates (dates are numbers in Excel).
' False for blanks, text, and error values. IsNumeric is too forgiving:
' IsNumeric(Empty) and IsNumeric("12") are both True.
Private Function IsNumberValue(v As Variant) As Boolean
    Select Case VarType(v)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbDate, vbByte
            IsNumberValue = True
        Case Else
            IsNumberValue = False
    End Select
End Function
