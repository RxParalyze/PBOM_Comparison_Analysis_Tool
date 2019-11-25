Attribute VB_Name = "Run_Sup_Dif_Sum"
Option Explicit

Public Sub RunSupDifSum()
    Dim s As Variant
    Dim suppCall As Supplier
    Dim cell As Range
    Dim s2 As Integer
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    'these method calls are required to make supplier calculations
    Call Globals 'Only Globals Call, Resets Globals
    Call Build_Suppliers_Orig
    Call Build_Suppliers_New
    Call Build_Parts_Orig
    Call Build_Parts_New
    
    Dim q As Variant
    Dim t As Integer
    
    'It creates an extra supplier with key 45551 at the end but with a space
    'in front. Blows up
    Debug.Print Suppliers.Keys(26), Suppliers.Items(26).SupplierName
    Debug.Print Suppliers.Keys(0), Suppliers.Items(0).SupplierName
    'Suppliers.Remove (Suppliers.Keys(0))
    For t = 0 To Suppliers.Count - 2
        'Debug.Print Suppliers.Items(t).SupplierName, Suppliers.Keys(t)
        Next t
    
    'For Each q In Suppliers.Keys
     '   t = t + 1
      '  Debug.Print t
       ' Debug.Print Suppliers(q).SupplierName
        'Next q
    
    'Add Suppliers to Supplier Differences Summary
    Dim c As Integer
    c = 1
    For Each s In Suppliers.Keys
        Set suppCall = Suppliers(s)
        
        'Supplier Method Calls
        suppCall.PartCountCalc
        
        supRange(c, 1).Value = suppCall.BestCodeVal
        supRange(c, 2).Value = suppCall.SupplierName
        supRange(c, 3).Value = suppCall.OrigVal
        supRange(c, 4).Value = suppCall.OrigPartCount
        supRange(c, 5).Value = suppCall.NewVal
        supRange(c, 6).Value = suppCall.NewPartCount
        supRange(c, 7).Value = suppCall.DeltaDollars
        supRange(c, 8).Value = suppCall.DeltaPercent
        
        Debug.Print suppCall.SupplierName
        c = c + 1
        Next s
    
    'Calculate check values for error checking
    Summary_Ws.Range("B4").Value = Suppliers.Count
    
    For s2 = 1 To PartNumbers.Count
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
        Next s2
        
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub


