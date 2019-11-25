Attribute VB_Name = "Concatenate_Parts"
Option Explicit

Public Function Concatenate(partNum As String, CLIN As String, BOE As String, bestCode As String, EAS As String)
    
    Concatenate = partNum & "_" & CLIN & "_" & BOE & "_" & bestCode & "_" & EAS
    
End Function
