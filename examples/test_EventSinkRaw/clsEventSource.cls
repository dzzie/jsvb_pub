VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEventSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents t As CommandButton
Attribute t.VB_VarHelpID = -1

Event TestEvent(ByVal Message As String, ByVal Number As Long)
Event AnotherEvent(ByVal Value As Double)

Public Sub FireTestEvent(msg As String, num As Long)
    RaiseEvent TestEvent(msg, num)
End Sub

Public Sub FireAnotherEvent(val As Double)
    RaiseEvent AnotherEvent(val)
End Sub

Private Sub t_Click()
    MsgBox "class t click!"
End Sub
