Attribute VB_Name = "Write_Globals"
Option Explicit
Global Suppliers As Scripting.Dictionary
Global PartNumbers As Scripting.Dictionary
Global Summary_Rng_BestCode As Range
Global Orig_Pbom_BC_Rng As Range
Global New_Pbom_BC_Rng As Range
Global Orig_Pbom_Vals As Range
Global New_Pbom_Vals As Range
Global partRange As Range
Global origPbom As Worksheet
Global newPbom As Worksheet
Global Ind_Part_Del As Worksheet
Global Summary_Ws As Worksheet
Global BestCodeRange As String
Global ValuesRange As String
Global supRange As Range
Global EndSupList As Integer

Public Sub Globals()
    Set origPbom = ThisWorkbook.Sheets("Original PBOM")
    Set newPbom = ThisWorkbook.Sheets("New PBOM")
    Set Summary_Ws = ThisWorkbook.Sheets("Supplier Differences Summary")
    Set Ind_Part_Del = ThisWorkbook.Sheets("Individual Part Deltas")
    BestCodeRange = "B10:B100010"
    ValuesRange = "C10:Z100010"
    EndSupList = 4006
    
    'Can I soft code these globals to make more user friendly? So code figures out which line is Best Code?
    Set Summary_Rng_BestCode = Summary_Ws.Range("B6:B" & EndSupList)
    Set supRange = Summary_Ws.Range("B6:J" & EndSupList)
    Set partRange = Ind_Part_Del.Range(BestCodeRange)
    Set Orig_Pbom_BC_Rng = origPbom.Range(BestCodeRange)
    Set New_Pbom_BC_Rng = newPbom.Range(BestCodeRange)
    Set Orig_Pbom_Vals = origPbom.Range(ValuesRange)
    Set New_Pbom_Vals = newPbom.Range(ValuesRange)
    
    'partRange.ClearContents
    'supRange.ClearContents
    Set PartNumbers = New Dictionary
    Set Suppliers = New Dictionary
    Suppliers.CompareMode = TextCompare

End Sub
