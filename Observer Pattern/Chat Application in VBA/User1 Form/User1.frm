VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} User1 
   Caption         =   "User 1 : Chat Application"
   ClientHeight    =   7335
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8640
   OleObjectBlob   =   "User1.frx":0000
End
Attribute VB_Name = "User1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ThisUser1 As IObserver

Private Sub UserForm_Initialize()
    Set ThisUser1 = New ChatApplication
    With ThisUser1
        Set .TargetFrame = User1.Frame1
        Set .TypingBox = User1.TextBox1
        .UserName = "USER 1"
        .AttachClickEventToSendButton User1.CommandButton1
        .AttachClickEventToClearButton User1.CommandButton2
        .Init
    End With
    'register thisuser1 for listening and notification
    modGlobal.GetNotified.AttachedUser ThisUser1
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Set ThisUser1 = Nothing
End Sub
