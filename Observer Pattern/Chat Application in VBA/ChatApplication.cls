VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChatApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IObserver

'Public Events
Private WithEvents btnSend As MSForms.CommandButton
Attribute btnSend.VB_VarHelpID = -1
Private WithEvents btnClear As MSForms.CommandButton
Attribute btnClear.VB_VarHelpID = -1

'Private Constants / Object's Settings
Private Const MESSAGE_FIRST_ROW_TOP_MARGIN As Byte = 24
Private Const MESSAGE_LEFT_GUTTER As Byte = 12
Private Const GAP_BETWEEN_MESSAGE As Byte = 12
Private Const MESSAGE_FIXED_WIDTH As Byte = 200
Private Const MESSAGE_FIXED_MIN_HEIGHT As Byte = 12
Private Const INCREASE_FRAME_SCROLLING As Byte = 48
Private Const UsernameRequired As Boolean = False
Private Const CONTROLTYPE_LABEL As String = "Label"

'Private Instance Variables
Private ChatWindowHeight As Long
Private Messagecounter As Long

'Private Types
Private Type ChatApplicationComponents
    TargetFrame As MSForms.Frame
    TypingBox As MSForms.TextBox
    UserName As String
    Message As String
    DataSheet As Worksheet
End Type

'Private Enums
Private Enum ChatData
    CDID = 1
    CDSender
    CDMessage
End Enum

'Private Type Referal
Private this As ChatApplicationComponents

'Main Method that triggers by notifier or Subject
Private Sub IObserver_UpdateChatWindow(ByVal Sender As ISubject)
    InjectMessageInChatWindow Sender.SenderMessage
End Sub

'Public Properties
Private Property Get IObserver_TargetFrame() As MSForms.Frame
    Set IObserver_TargetFrame = this.TargetFrame
End Property

Private Property Set IObserver_TargetFrame(ByVal RHS As MSForms.Frame)
    Set this.TargetFrame = RHS
End Property

Private Property Get IObserver_TypingBox() As MSForms.TextBox
    Set IObserver_TypingBox = this.TypingBox
End Property

Private Property Set IObserver_TypingBox(ByVal RHS As MSForms.TextBox)
    Set this.TypingBox = RHS
End Property

Private Property Get IObserver_UserName() As String
    IObserver_UserName = this.UserName
End Property

Private Property Let IObserver_UserName(ByVal vNewValue As String)
    this.UserName = vNewValue
End Property

Private Property Get IObserver_Message() As String
    IObserver_Message = this.Message
End Property

Private Property Let IObserver_Message(ByVal vNewValue As String)
    this.Message = vNewValue
End Property

'Private Properties
Private Property Get DataSheet() As Worksheet
    Set DataSheet = Sheet2
End Property

'Public Methods
Private Sub IObserver_Init()
    With IObserver_TargetFrame
        .ScrollHeight = .Height
        .Font.Size = 13
    End With
    IObserver_TypingBox.SetFocus
End Sub

Private Sub IObserver_AttachClickEventToSendButton(ByVal SendButton As MSForms.CommandButton)
    Set btnSend = SendButton
End Sub

Private Sub IObserver_AttachClickEventToClearButton(ByVal ClearButton As MSForms.CommandButton)
    Set btnClear = ClearButton
End Sub

'Private Methods
Private Sub btnClear_Click()
    Call ClearChatWindow
End Sub

Private Sub btnSend_Click()
    Dim LastRow As Long
    IObserver_Message = IObserver_TypingBox.Text
    AddMessageToCentralDatabase
    LastRow = GetLastRow
    'register thisuser1 for listening and notification
    Set modGlobal.GetNotified.SenderMessage = DataSheet.Range("A" & LastRow & ":C" & LastRow)
End Sub

Private Sub ClearChatWindow()
    Dim iControl As MSForms.Control
    For Each iControl In IObserver_TargetFrame.Controls
        If IsControlTypeLabel(iControl) Then
            IObserver_TargetFrame.Controls.Remove iControl.Name
        End If
    Next iControl
    IObserver_TypingBox.SetFocus
    Set iControl = Nothing
End Sub

Private Sub InjectMessageInChatWindow(ByVal MessageMetaData As Range)
    Dim iMessage        As MSForms.Label
    Dim currentTop      As Long
    Dim currentHeight   As Long
    Messagecounter = Messagecounter + 1
    If Not AnyOpenMessage Then
        Set iMessage = CreateLabelForChatWindow(MESSAGE_LEFT_GUTTER, MESSAGE_FIRST_ROW_TOP_MARGIN, MESSAGE_FIXED_WIDTH, MESSAGE_FIXED_MIN_HEIGHT, MessageMetaData)
    Else
        Set iMessage = CreateLabelForChatWindow(MESSAGE_LEFT_GUTTER, ChatWindowHeight + GAP_BETWEEN_MESSAGE, MESSAGE_FIXED_WIDTH, MESSAGE_FIXED_MIN_HEIGHT, MessageMetaData)
    End If
    With iMessage
        currentTop = .Top
        currentHeight = .Height
    End With
    ChatWindowHeight = currentTop + currentHeight
    With IObserver_TargetFrame
        If .ScrollHeight < ChatWindowHeight Then
            If currentHeight > INCREASE_FRAME_SCROLLING Then
                .ScrollHeight = .ScrollHeight + currentHeight
            Else
                .ScrollHeight = .ScrollHeight + INCREASE_FRAME_SCROLLING
            End If
        End If
        .Scroll xAction:=fmScrollActionNoChange, yAction:=fmScrollActionEnd
    End With
    With IObserver_TypingBox
        .Value = vbNullString
        .SetFocus
    End With
    Set iMessage = Nothing
End Sub

'Private Functions
Private Function CreateLabelForChatWindow(ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, ByVal MessageMetaData As Range) As MSForms.Label
    Dim currentUserName As String
    Dim currentMessage As String
    currentUserName = MessageMetaData.Cells(1, 2).Value
    currentMessage = MessageMetaData.Cells(1, 3).Value
    Set CreateLabelForChatWindow = IObserver_TargetFrame.Controls.Add("Forms.Label.1", "Message" & Messagecounter, True)
    With CreateLabelForChatWindow
        .Left = Left
        .Top = Top
        .Width = Width
        .Height = Height
        .BackColor = IIf(currentUserName = IObserver_UserName, vbYellow, vbGreen)
        .Caption = IIf(currentUserName = IObserver_UserName, "Me", currentUserName) & ": " & currentMessage
        .WordWrap = True
        .AutoSize = True
    End With
End Function

Private Sub AddMessageToCentralDatabase()
    Dim NewRow As Long
    NewRow = GetLastRow + 1
    With DataSheet
        .Cells(NewRow, ChatData.CDID).Value = NewRow - 1
        .Cells(NewRow, ChatData.CDSender).Value = IObserver_UserName
        .Cells(NewRow, ChatData.CDMessage).Value = IObserver_Message
    End With
End Sub

Private Function AnyOpenMessage() As Boolean
    Dim iControl As MSForms.Control
    Dim MessageFoundOnChatWindow As Boolean
    MessageFoundOnChatWindow = False
    For Each iControl In IObserver_TargetFrame.Controls
        If IsControlTypeLabel(iControl) Then
            MessageFoundOnChatWindow = True
            Exit For
        End If
    Next iControl
    AnyOpenMessage = MessageFoundOnChatWindow
    Set iControl = Nothing
End Function

Private Function IsControlTypeLabel(ByVal ControlToCheck As MSForms.Control) As Boolean
    IsControlTypeLabel = False
    If VBA.TypeName(ControlToCheck) = CONTROLTYPE_LABEL Then
        IsControlTypeLabel = True
    End If
End Function

Private Function GetLastRow() As Long
    With DataSheet
        GetLastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
    End With
End Function

