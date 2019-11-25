Attribute VB_Name = "Run_Ind_Part_Deltas"
Option Explicit

Public Sub RunIndPartDelta()
    Dim s As Integer
    Dim cell As Range
    
    'Next two lines keep the sheet from breaking, takes too long to load otherwise
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    'Build_Suppliers method is called because parts are attached to supplier
    'Uses Inheritance structure to determine supplier values
    Call Globals 'Only for debugging this specific module
    Call Build_Suppliers_Orig
    Call Build_Suppliers_New
    Call Build_Parts_Orig
    Call Build_Parts_New
    
    For s = 1 To PartNumbers.Count
        With PartNumbers(s)
            partRange(s, 1).Value = .PartNumber
            partRange(s, 2).Value = .ConCatNum
            partRange(s, 3).Value = .DeltaDollars
            partRange(s, 4).Value = .DeltaPercent
            partRange(s, 5).Value = .DeltaUnit
            partRange(s, 6).Value = .DeltaQuantity
            partRange(s, 7).Value = .DeltaGroup
            partRange(s, 8).Value = .thisUP
            partRange(s, 9).Value = .thisUM
            partRange(s, 10).Value = .thisFR
            partRange(s, 11).Value = .thisQuan
            partRange(s, 12).Value = .thisExP
            partRange(s, 13).Value = .newUP
            partRange(s, 14).Value = .newUM
            partRange(s, 15).Value = .newFR
            partRange(s, 16).Value = .newQuan
            partRange(s, 17).Value = .newExP
        End With
        Next s
        
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub
