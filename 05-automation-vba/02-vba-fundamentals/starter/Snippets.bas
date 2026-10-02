Attribute VB_Name = "Snippets"
' =====================================================================
' Lesson 5.2 - predict-the-output snippets (Practice tasks 2-7)
'
' This module deliberately has NO Option Explicit line, so that
' Snippet A can run. Never leave it out of your own modules.
'
' For each snippet: predict what it prints, type your prediction on the
' Practice sheet, THEN click inside the Sub and press F5 to check.
' Debug.Print writes to the Immediate window (View > Immediate Window).
' =====================================================================

' Snippet A
Sub SnippetA_Typo()
    Dim drawCount As Long
    drawCount = 5
    drawCuont = drawCount + 1       ' add one more blood draw
    Debug.Print drawCount
End Sub

' Snippet B
Sub SnippetB_Overflow()
    Dim labRows As Integer
    labRows = 51527                 ' rows in the full two-year lab file
    Debug.Print labRows
End Sub

' Snippet C
Sub SnippetC_Turnaround()
    Dim tatMinutes As Long
    tatMinutes = 199                ' slowest STAT result on the Labs sheet
    Debug.Print tatMinutes \ 60 & " h " & tatMinutes Mod 60 & " min"
End Sub

' Snippet D
Sub SnippetD_VitalSigns()
    Dim hr As Long, checks As Long
    For hr = 0 To 23 Step 4         ' vital signs every 4 hours (q4h)
        checks = checks + 1
    Next hr
    Debug.Print hr
End Sub

' Snippet E
Sub SnippetE_Lactate()
    Dim lactate As Double, category As String
    lactate = 4.6                   ' mmol/L
    Select Case lactate
        Case Is < 0.5
            category = "Low"
        Case Is <= 2
            category = "Normal"
        Case Is > 2
            category = "Elevated"
        Case Is > 4
            category = "Critical"
        Case Else
            category = "Check value"
    End Select
    Debug.Print category
End Sub

' Snippet F
Sub SnippetF_HalfLife()
    Dim level As Double, hours As Long
    level = 400                     ' ng/mL right after the dose
    Do While level > 50
        level = level / 2           ' the drug's half-life is 6 hours
        hours = hours + 6
    Loop
    Debug.Print hours
End Sub
