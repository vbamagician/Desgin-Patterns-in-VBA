VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TravelPlanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public TravelStrategy As IStrategy

Public Sub SetTravelStrategy(ByVal myStrategy As IStrategy)
    Set TravelStrategy = myStrategy
End Sub

Public Sub Drive(ByVal TravelDistance)
    TravelStrategy.Drive TravelDistance
End Sub

Public Function GetCostofTravel(ByVal TravelDistance)
    GetCostofTravel = TravelStrategy.GetCost(TravelDistance)
End Function
