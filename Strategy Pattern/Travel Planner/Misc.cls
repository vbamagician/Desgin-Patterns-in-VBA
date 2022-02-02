VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Misc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub ShowResultOnImmediateWindow(ByVal TravelWith As String, ByVal kilometers As Long, ByVal totalCost As Double)
    Debug.Print "Total cost of travel with " & TravelWith & " for " & kilometers & " KM is : " & VBA.Format(totalCost, "Standard")
End Sub

Public Function CalculateCostOfTravel(ByVal kilometers As Single, ByVal costOfPerKilometers As Long) As Double
    CalculateCostOfTravel = kilometers * costOfPerKilometers
End Function

Public Function GetResult(ByVal TravelWith As String, ByVal kilometers As Long, ByVal totalCost As Double) As String
    GetResult = "Total cost of travel with " & TravelWith & " for " & kilometers & " KM is : " & VBA.Format(totalCost, "Standard")
End Function

