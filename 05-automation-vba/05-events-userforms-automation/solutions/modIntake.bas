Attribute VB_Name = "modIntake"
'==============================================================================
' Lesson 5.5 - modIntake (SOLUTION)
' Import this AFTER you have created frmIntake; otherwise "Debug > Compile"
' reports "Variable not defined" because the form doesn't exist yet.
' Assign ShowIntakeForm to a button on the Intake sheet
' (Insert > Shapes, draw a shape, right-click > Assign Macro...).
'==============================================================================
Option Explicit

Public Sub ShowIntakeForm()
    frmIntake.Show          ' modal: the user closes the form before working in the sheet again
End Sub
