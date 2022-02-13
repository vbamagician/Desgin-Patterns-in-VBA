VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IObserver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub UpdateChatWindow(ByVal Sender As ISubject) 'main procedure for pattern
    
End Sub

Public Property Get TargetFrame() As MSForms.Frame

End Property

Public Property Set TargetFrame(ByVal vNewValue As MSForms.Frame)

End Property

Public Property Get TypingBox() As MSForms.TextBox

End Property

Public Property Set TypingBox(ByVal vNewValue As MSForms.TextBox)

End Property

Public Property Get UserName() As String

End Property

Public Property Let UserName(ByVal vNewValue As String)

End Property

Public Sub Init()

End Sub

Public Sub AttachClickEventToSendButton(ByVal SendButton As MSForms.CommandButton)

End Sub

Public Sub AttachClickEventToClearButton(ByVal ClearButton As MSForms.CommandButton)

End Sub
