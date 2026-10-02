Attribute VB_Name = "HealthUDFs"
Option Explicit

' =====================================================================
' Lesson 5.4 - starter module for the UDF tasks 1-6
'
' 1. Save the workbook as .xlsm.  2. In the VBE: File > Import File...
' 3. Replace each TODO. 4. Use the functions in cells, e.g. =BMI(G2,F2).
' Public functions in a standard module like this one can be called from
' worksheet cells. The two helpers at the bottom are finished for you.
' =====================================================================

' Tasks 1-2: body-mass index = 703 x weight (lb) / height (in) ^ 2
Public Function BMI(weightLb As Variant, heightIn As Variant) As Variant
    Dim w As Variant, h As Variant
    w = ArgValue(weightLb)
    h = ArgValue(heightIn)
    ' TODO 1: if w or h is not a number (use IsNumberValue), return a #VALUE! error
    ' TODO 2: if w or h is zero or negative, return a #NUM! error
    ' TODO 3: otherwise return the BMI
End Function

' Task 3: age in completed years on asOf. Leave asOf out to use today's date.
Public Function AGEAT(dob As Variant, Optional asOf As Variant) As Variant
    ' TODO 1: read dob with ArgValue. If asOf is missing use Date, otherwise ArgValue(asOf)
    ' TODO 2: return #VALUE! if either value is not a number or date
    ' TODO 3: work out the age in completed years (careful: a birthday that is still
    '         to come in the asOf year doesn't count yet)
End Function

' Tasks 4-5: length of stay = midnights between admission and discharge
Public Function LOSDAYS(admitDateTime As Variant, dischargeDateTime As Variant) As Variant
    ' TODO 1: read both arguments with ArgValue; #VALUE! if either is not a date
    ' TODO 2: #VALUE! if the discharge date-time is earlier than the admission
    ' TODO 3: otherwise return the number of midnights between the two (DateDiff)
End Function

' Task 6: share of non-blank cells in statusRange that equal statusToCount
Public Function DENIALRATE(statusRange As Range, Optional statusToCount As String = "Denied") As Variant
    ' TODO 1: loop over statusRange.Cells: count the non-blank cells (total) and the
    '         cells whose text equals statusToCount, ignoring case (hits)
    ' TODO 2: return #DIV/0! if total = 0, otherwise hits / total
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
