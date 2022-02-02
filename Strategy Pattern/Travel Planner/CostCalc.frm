VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CostCalc 
   Caption         =   "Travel Planner"
   ClientHeight    =   3435
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7215
   OleObjectBlob   =   "CostCalc.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CostCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum TransportationMode
    ByBus = 1
    ByTrain
    ByPlane
End Enum

Private Sub CommandButton1_Click()
    
    With Me
        If .ComboBox1.Value = vbNullString Then Exit Sub
        If .TextBox1.Value = vbNullString Then Exit Sub
        If VBA.IsNumeric(.TextBox1.Value) = False Then
            .TextBox1.Value = vbNullString
            .TextBox1.SetFocus
        End If
        
        Dim myPlanner As TravelPlanner
        Set myPlanner = New TravelPlanner
        
        myPlanner.SetTravelStrategy Travelby(.ComboBox1.ListIndex + 1)
        .Label4.Caption = myPlanner.GetCostofTravel(VBA.Val(.TextBox1.Value))
        Set myPlanner = Nothing
    End With
    
End Sub

Private Sub CommandButton2_Click()
    With Me
        .TextBox1.Value = vbNullString
        .ComboBox1.Value = vbNullString
        .Label4.Caption = vbNullString
        .ComboBox1.SetFocus
    End With
End Sub

Private Sub CommandButton3_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    With Me.ComboBox1
        .AddItem "By Bus"
        .AddItem "By Train"
        .AddItem "By Airplane"
    End With
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
