VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Supplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Declare Supplier Object Properties
Private bestCode As String
Private suppName As String
Private origPbom As Double
Private origPbomPartCount As Integer
Private newPbom As Double
Private newPbomPartCount As Integer
Private deltaDollar As Double
Private deltaPerc As Double
Public origPartList As New Collection
Public newPartList As New Collection

Public Property Get BestCodeVal() As String
    BestCodeVal = bestCode
End Property
Public Property Let BestCodeVal(Value As String)
    bestCode = Value
End Property
Public Property Get SupplierName() As String
    SupplierName = suppName
End Property
Public Property Let SupplierName(Value As String)
    suppName = Value
End Property
Public Property Get OrigVal() As Double
    OrigVal = origPbom
End Property
Public Property Let OrigVal(Value As Double)
    origPbom = Value
End Property
Public Property Get OrigPartCount() As Integer
    OrigPartCount = origPbomPartCount
End Property
Public Function PartCountCalc()
    origPbomPartCount = origPartList.Count
    newPbomPartCount = newPartList.Count
End Function
Public Property Get NewVal() As Double
    NewVal = newPbom
End Property
Public Property Let NewVal(Value As Double)
    NewVal = Value
End Property
Public Property Get NewPartCount() As Integer
    NewPartCount = newPbomPartCount
End Property
Public Property Get DeltaDollars() As Double
    Call Delta
    DeltaDollars = deltaDollar
End Property
Public Property Get DeltaPercent() As Double
    Call Delta
    DeltaPercent = deltaPerc
End Property
Public Sub AddOrigPart(ByRef newPart As String)
    origPartList.Add newPart
End Sub
Public Sub AddNewPart(ByRef newPart As String)
    newPartList.Add newPart
End Sub
Private Sub Delta()
    deltaDollar = newPbom - origPbom
    deltaPerc = calcDeltaPerc(deltaDollar, origPbom)
End Sub

Private Sub Class_Initialize()
    Set origPartList = New Collection
    Set newPartList = New Collection
End Sub
