Attribute VB_Name = "If_Error_Code"
Option Explicit

Public Function calcDeltaPerc(firstVal As Variant, secondVal As Variant)
    Dim fixForm As Double
    fixForm = 1
    On Error GoTo ErrorHandler
    
    If IsError(firstVal / secondVal) Then
        calcDeltaPerc = fixForm
    Else
        calcDeltaPerc = firstVal / secondVal
    End If
    
ErrorHandler:
    Resume Next
    
End Function
