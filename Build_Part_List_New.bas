Attribute VB_Name = "Build_Part_List_New"
Option Explicit

Public Sub Build_Parts_New()
'Possibly add function to determine how many parts to add onto the Part List?
    Dim cell As Range
    Dim findSup As Variant
    Dim myBool As Boolean
    Dim partMod As Integer
    Dim newPart As Variant
    Dim errCheck As Integer
    Dim addPartBool As Boolean
    Dim attachPartBool As Boolean
    Dim thisPart As Variant
    Dim checkParts As Variant
    
    For Each cell In New_Pbom_BC_Rng.Cells
        If cell.Value = "" Then
            'MsgBox ("Empty Cell")
            err.Clear
            GoTo EndSub
            'Exit For
        End If
        Set newPart = New Part
        'Initialize a new part and add values to vars
        err.Clear
        GoTo OrigSupp


OrigSupp:
        On Error GoTo ErrHandler
        
        newPart.ConCatNum = Concatenate(cell.Offset(0, 2).Value, cell.Offset(0, 4).Value, cell.Offset(0, 3).Value, cell.Value, cell.Offset(0, 5).Value)
        cell.Offset(0, 6).Value = newPart.ConCatNum
        newPart.PartNumber = cell.Offset(0, 2).Value
        newPart.newUP = cell.Offset(0, 7).Value
        newPart.newUM = cell.Offset(0, 8).Value
        newPart.newFR = cell.Offset(0, 9).Value
        newPart.newQuan = cell.Offset(0, 10).Value
        newPart.newExP = cell.Offset(0, 11).Value
        
        PartNumbers.Add Item:=newPart, key:=newPart.PartNumber
        addPartBool = True
        attachPartBool = False
        
        GoTo FinishPart
        
FinishPart:
        'Tries add method if PartExists was run and new part is needed
        If addPartBool = False Then
            PartNumbers.Add Item:=newPart, key:=newPart.PartNumber
            
            'If PartExists was run but new part not needed, uses following methods
        Else
            If attachPartBool = True Then
                thisPart.newUP = cell.Offset(0, 7).Value
                thisPart.newUM = cell.Offset(0, 8).Value
                thisPart.newFR = cell.Offset(0, 9).Value
                thisPart.newQuan = thisPart.newQuan + cell.Offset(0, 10).Value
                thisPart.newExP = thisPart.newExP + cell.Offset(0, 11).Value
            End If
        End If
        
        'Find part's supplier, attach to private part list, add total value
        Dim thisSupplier As Supplier
        Set thisSupplier = Suppliers(cell.Value)
        Dim thisSuppPartList As Collection
        Set thisSuppPartList = thisSupplier.newPartList
        
        findSup = HasKey(thisSuppPartList, newPart.PartNumber) 'To execute value increase
        
        With thisSupplier
            'this next line is not working for some reason
            .NewVal = .NewVal + cell.Offset(0, 11).Value 'Increases total value
            If (findSup = False) Then
                thisSupplier.AddNewPart newPart.PartNumber 'adds part if doesn't exist
                End If
            Suppliers.Remove newPart.BestCodeVal
            Suppliers.Add Item:=thisSupplier, key:=newPart.BestCodeVal
            End With
        GoTo nextCell

'If supps are equal, attach to original part with additional BOE and CLIN
PartExists:
        'This piece attaches a new digit to the Part Number so that it keeps parts separate
        myBool = True
        Set thisPart = PartNumbers.Item(newPart.PartNumber)
        partMod = 2
        'This is only comparing the first iteration of a part right now
        If newPart.BestCodeVal <> thisPart.BestCodeVal Then
            newPart.PartNumber = newPart.PartNumber & "_#1"
            Do Until myBool = False
                newPart.PartNumber = removeDigit(newPart.PartNumber)
                newPart.PartNumber = newPart.PartNumber & partMod
                partMod = partMod + 1
                myBool = HasKey(PartNumbers, newPart.PartNumber)
                Loop
            attachPartBool = False
        Else
            attachPartBool = True
            addPartBool = True
        End If
        GoTo FinishPart



nextCell:
        errCheck = 0
        addPartBool = False
        attachPartBool = False
        Next cell

ErrHandler:
        errCheck = errCheck + 1
        Select Case errCheck
            Case 1
                Resume PartExists
            Case 2
                Resume EndSub
            Case Else
                MsgBox (err.Source)
        End Select

EndSub:
        End Sub
        

Function removeDigit(modName As String)
    'Dim NEWSTRING As String
    removeDigit = Left(modName, Len(modName) - 1)
    'removeDigit = NEWSTRING
End Function

