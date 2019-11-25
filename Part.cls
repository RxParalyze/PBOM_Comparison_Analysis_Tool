VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Part"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Universal Part Object Properties
Private partNum As String
Private conCatPartNum As String

'Original Pbom Private Instance Vars
Private unitPrice As Double
Private unitMeasure As String
Private factorRollup As Double
Private quantity As Integer
Private extPrice As Double

'New Pbom Private Instance Vars
Private unitPriceNew As Double
Private unitMeasureNew As String
Private factorRollupNew As Double
Private quantityNew As Integer
Private extPriceNew As Double

'Comparison Vars
Private deltaDollar As Double
Private deltaPerc As Double
Private deltaQuan As Integer
Private deltaUnitPrice As Double
Private deltaGrp As String

'Universal Properties
Public Property Get PartNumber() As String
    PartNumber = partNum
End Property
Public Property Let PartNumber(Value As String)
    partNum = Value
End Property
Public Property Get ConCatNum() As String
    ConCatNum = conCatPartNum
End Property
Public Property Let ConCatNum(Value As String)
    conCatPartNum = Value
End Property

'Original PBOM Properties
Public Property Get thisUM() As String
    thisUM = unitMeasure
End Property
Public Property Let thisUM(Value As String)
    unitMeasure = Value
End Property
Public Property Get thisUP() As Double
    thisUP = unitPrice
End Property
Public Property Let thisUP(Value As Double)
    unitPrice = Value
End Property
Public Property Get thisFR() As Double
    thisFR = factorRollup
End Property
Public Property Let thisFR(Value As Double)
    factorRollup = Value
End Property
Public Property Get thisExP() As Double
    thisExP = extPrice
End Property
Public Property Let thisExP(Value As Double)
    extPrice = Value
End Property
Public Property Get thisQuan() As Integer
    thisQuan = quantity
End Property
Public Property Let thisQuan(Value As Integer)
    quantity = Value
End Property

'New PBOM Properties
Public Property Get newUM() As String
    newUM = unitMeasureNew
End Property
Public Property Let newUM(Value As String)
    unitMeasureNew = Value
End Property
Public Property Get newUP() As Double
    newUP = unitPriceNew
End Property
Public Property Let newUP(Value As Double)
    unitPriceNew = Value
End Property
Public Property Get newFR() As Double
    newFR = factorRollupNew
End Property
Public Property Let newFR(Value As Double)
    factorRollupNew = Value
End Property
Public Property Get newExP() As Double
    newExP = extPriceNew
End Property
Public Property Let newExP(Value As Double)
    extPriceNew = Value
End Property
Public Property Get newQuan() As Integer
    newQuan = quantityNew
End Property
Public Property Let newQuan(Value As Integer)
    quantityNew = Value
End Property

'Delta Properties
Public Property Get DeltaDollars() As Double
    Call Delta
    DeltaDollars = deltaDollar
End Property
Public Property Get DeltaPercent() As Double
    DeltaPercent = deltaPerc
End Property
Public Property Get DeltaQuantity() As Integer
    DeltaQuantity = deltaQuan
End Property
Public Property Get DeltaUnit() As Double
    DeltaUnit = deltaUnitPrice
End Property
Public Property Get DeltaGroup() As String
    Call DeltaGroupSplit
    DeltaGroup = deltaGrp
End Property
Private Sub Delta()
    deltaDollar = extPriceNew - extPrice
    deltaQuan = quantityNew - quantity
    deltaPerc = calcDeltaPerc(deltaDollar, extPrice)
    deltaUnitPrice = unitPriceNew - unitPrice
End Sub


Private Sub DeltaGroupSplit()
    Dim thisPercentVal As Double
    thisPercentVal = Round((deltaPerc * 100) + 0.00000001, 0) '0.00000001 is to prevent vba round errors
    'Here i'm trying to convert the delta %s into groups for Column K on Ind Part Deltas
    Select Case thisPercentVal
        Case -2000000000 To -1 'arbitrary number to cover all negatives
            deltaGrp = "Decreased"
        Case 0
            deltaGrp = "No Change"
        Case 1 To 10
            deltaGrp = "Less than 10%"
        Case 10 To 20
            deltaGrp = "10-20%"
        Case 20 To 50
            deltaGrp = "20-50%"
        Case 50 To 100
            deltaGrp = "50-100%"
        Case Else
            deltaGrp = "Over 100% Increase"
        End Select
End Sub




