Attribute VB_Name = "modGlobal"
Option Explicit

'Designed & Developed By Kamal Bharakhda
'E-Mail Address: kamal.9328093207@gmail.com
'Phone: +91-9328093207
'Project: Chat Application

Private Notifier As ISubject

'Creating Global Instance of Sender Class
Public Function GetNotified() As ISubject
    If Notifier Is Nothing Then Set Notifier = New Sender
    Set GetNotified = Notifier
End Function
