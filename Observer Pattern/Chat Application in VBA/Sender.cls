VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Sender or Notifier
Implements ISubject

Private Type SubjectParameters
    SenderMessage As Range
End Type

Private listOfObservers As Collection
Private this As SubjectParameters

Private Property Get ISubject_SenderMessage() As Range
    Set ISubject_SenderMessage = this.SenderMessage
End Property

Private Property Set ISubject_SenderMessage(ByVal RHS As Range)
    Set this.SenderMessage = RHS
    Call ISubject_NotifyUser
End Property

Private Sub ISubject_AttachedUser(ByVal User As IObserver)
    listOfObservers.Add User
End Sub

Private Sub ISubject_NotifyUser()
    Dim Observer As IObserver
    For Each Observer In listOfObservers
        Observer.UpdateChatWindow Me 'me is a instance of weather station
    Next Observer
End Sub

Private Sub Class_Initialize()
    Set listOfObservers = New Collection
End Sub

Private Sub Class_Terminate()
    Set listOfObservers = Nothing
End Sub
