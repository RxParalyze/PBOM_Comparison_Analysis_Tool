Attribute VB_Name = "Build_Orig_Supp_List"
Option Explicit

Public Sub Build_Suppliers_Orig()
    Dim cell As Range
    Dim Supp As Supplier
    Dim errCheck As Integer
    
    For Each cell In Orig_Pbom_BC_Rng.Cells
        If cell.Value = "" Then
            err.Clear
            GoTo EndSub
            'Exit For
        End If
        Set Supp = New Supplier
        'Initialize a new supplier and add values to vars
        err.Clear
        GoTo OrigSupp
        
OrigSupp:
        On Error GoTo ErrHandler
        Supp.BestCodeVal = cell.Value
        Supp.SupplierName = cell.Offset(0, 1).Value
        'error next line
        Suppliers.Add Supp.BestCodeVal, Supp
        'Debug.Print Supp.BestCodeVal
        GoTo nextCell
        
SuppExists:
        GoTo nextCell
        
nextCell:
        errCheck = 0
        Next cell
        
ErrHandler:
        errCheck = errCheck + 1
        Select Case errCheck
            Case 1
                Resume SuppExists
            Case 2
                Resume EndSub
            Case Else
                MsgBox (err.Source)
        End Select
        
EndSub:
        End Sub
