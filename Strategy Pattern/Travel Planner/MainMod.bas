Attribute VB_Name = "MainMod"
Option Explicit

Private Enum TransportationMode
    ByBus = 1
    ByTrain
    ByPlane
End Enum

Public Sub Main()
    
    Dim myPlanner As TravelPlanner
    Set myPlanner = New TravelPlanner
    
    myPlanner.SetTravelStrategy Travelby(ByBus)
    myPlanner.Drive 1000
    
    myPlanner.SetTravelStrategy Travelby(ByTrain)
    myPlanner.Drive 1000
    
    myPlanner.SetTravelStrategy Travelby(ByPlane)
    myPlanner.Drive 900
    
    Set myPlanner = Nothing
    
End Sub

Private Function Travelby(ByVal Mode As TransportationMode) As IStrategy
    Set Travelby = GetListOfObjects.Item(VBA.CStr(Mode))
End Function

Private Function GetListOfObjects() As Collection
    Static Coll As Collection
    If Coll Is Nothing Then
        Set Coll = New Collection
        Coll.Add New Bus, VBA.CStr(TransportationMode.ByBus)
        Coll.Add New Train, VBA.CStr(TransportationMode.ByTrain)
        Coll.Add New Plane, VBA.CStr(TransportationMode.ByPlane)
    End If
    Set GetListOfObjects = Coll
End Function
