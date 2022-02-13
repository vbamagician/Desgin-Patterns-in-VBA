VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} User2 
   Caption         =   "User 2 : Chat Application"
   ClientHeight    =   7335
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8640
   OleObjectBlob   =   "User2.frx":0000
End
Attribute VB_Name = "User2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ThisUser2 As IObserver

Private Sub UserForm_Initialize()
    Set ThisUser2 = New ChatApplication
    With ThisUser2
        Set .TargetFrame = User2.Frame1
        Set .TypingBox = User2.TextBox1
        .UserName = "USER 2"
        .AttachClickEventToSendButton User2.CommandButton1
        .AttachClickEventToClearButton User2.CommandButton2
        .Init
    End With
    'register thisuser1 for listening and notification
    modGlobal.GetNotified.AttachedUser ThisUser2
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Set ThisUser2 = Nothing
End Sub
