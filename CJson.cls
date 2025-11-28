VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CJson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Author:  David Zimmer <dzzie@yahoo.com> + Claude.ai
'Site:    http://sandsprite.com
'License: MIT

'this is to give vb6 access to json formatted data

Option Explicit

' Singleton is fine
Private m_jsonInterp As CInterpreter

Friend Property Get interp() As CInterpreter
    Set interp = m_jsonInterp
End Property

Public Function ParseToCVal(jsonText As String) As CValue
    On Error Resume Next
    Set ParseToCVal = ParseJSON(jsonText)
End Function

Public Function Parse(json As String) As Object
    If m_jsonInterp Is Nothing Then
        Set m_jsonInterp = New CInterpreter
    End If
    
    m_jsonInterp.Execute "var __j = " & json & ";"
    Set Parse = m_jsonInterp.EvalAsObject("__j")
End Function

Public Function Stringify() As String
    Stringify = m_jsonInterp.Eval("JSON.stringify(__j)").strVal
End Function

Private Sub Class_Initialize()
    'Debug.Print ">>> CJson.Initilize " & hex(ObjPtr(Me))
End Sub

Private Sub Class_Terminate()
    'Debug.Print ">>> CJson.Terminate " & hex(ObjPtr(Me))
End Sub
