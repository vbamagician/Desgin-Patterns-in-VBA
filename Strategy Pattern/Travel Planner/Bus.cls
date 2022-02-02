VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Bus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IStrategy

Private Const MYMODE As String = "Bus"

Private Property Get IStrategy_costPerKilometer() As Single
    IStrategy_costPerKilometer = 20
End Property

Private Property Get GetGeneralMethods() As Misc
    Set GetGeneralMethods = New Misc
End Property

Private Sub IStrategy_drive(ByVal kilometers As Long)
    Dim totalCost As Double
    totalCost = GetGeneralMethods.CalculateCostOfTravel(kilometers, IStrategy_costPerKilometer)
    GetGeneralMethods.ShowResultOnImmediateWindow MYMODE, kilometers, totalCost
End Sub

Private Function IStrategy_GetCost(ByVal kilometers As Long)
    Dim totalCost As Double
    totalCost = GetGeneralMethods.CalculateCostOfTravel(kilometers, IStrategy_costPerKilometer)
    IStrategy_GetCost = GetGeneralMethods.GetResult(MYMODE, kilometers, totalCost)
End Function
