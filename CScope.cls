VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CScope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Author:  David Zimmer <dzzie@yahoo.com>
'AI:      Claude.ai
'Site:    http://sandsprite.com
'License: MIT

'------------------------------------------------------------
' CScope.cls - Variable scope storage
' "Where variables live"
'------------------------------------------------------------

Option Explicit

Private m_vars As Collection
Private m_parent As CScope    ' Parent scope (for closures later)
Private m_varNames As Collection  'track names

Private Sub Class_Initialize()
    Set m_vars = New Collection
    Set m_varNames = New Collection
End Sub

' Define a new variable in this scope
Public Sub DefineVar(Name As String, Value As CValue)
    On Error Resume Next
    ' Remove if exists
    m_vars.Remove Name
    m_varNames.Remove Name
    On Error GoTo 0
    
    ' Add new value
    m_vars.add Value, Name
    m_varNames.add Name, Name  'track the name
End Sub

Public Sub DeleteVar(varName As String)
    On Error Resume Next
    m_vars.Remove varName
    On Error GoTo 0
End Sub

Public Function GetVar(Name As String) As CValue
    On Error Resume Next
    Dim val As CValue
    Set val = m_vars(Name)
    
    ' DEBUG
    'Debug.Print ">>> GetVar(" & name & ") - retrieved type: " & val.vType
    If val.vType = vtCOMObject Then
        'Debug.Print ">>> objVal Is Nothing: " & (val.objVal Is Nothing)
    End If
    
    If Err.Number <> 0 Then
        ' Not found in this scope
        If Not m_parent Is Nothing Then
            Set GetVar = m_parent.GetVar(Name)
        Else
            Dim undef As New CValue
            undef.vType = vtUndefined
            Set GetVar = undef
        End If
    Else
        Set GetVar = val
    End If
    On Error GoTo 0
End Function

Public Sub SetVar(Name As String, Value As CValue)
    Dim existing As CValue
    Dim found As Boolean
    
    ' Try to find in this scope
    On Error Resume Next
    Set existing = m_vars(Name)
    found = (Err.Number = 0)
    On Error GoTo 0
    
    If found Then
        ' Found in this scope - update it
        m_vars.Remove Name
        m_vars.add Value, Name
    ElseIf Not m_parent Is Nothing Then
        ' Try parent scope
        m_parent.SetVar Name, Value
    Else
        ' Not found anywhere - create in global scope (sloppy mode)
        ' This prevents infinite recursion
        m_vars.add Value, Name
        m_varNames.add Name, Name
    End If
End Sub

' Check if variable exists
Public Function HasVar(Name As String) As Boolean
    On Error Resume Next
    Dim val As CValue
    Set val = m_vars(Name)
    HasVar = (Err.Number = 0)
    On Error GoTo 0
End Function

' Set parent scope (for nested scopes)
Public Property Set Parent(p As CScope)
    Set m_parent = p
End Property

Public Function GetAllVariables() As Collection 'of CVarItem
    Dim result As New Collection
    Dim i As Long
    Dim varName As String
    Dim varValue As CValue
    Dim vi As CVarItem

    For i = 1 To m_varNames.count
        varName = m_varNames(i)
        Set vi = New CVarItem
        Set vi.Value = m_vars(varName)
        vi.Name = varName
        vi.isGlobal = (m_parent Is Nothing)
        result.add vi
    Next
    
    Set GetAllVariables = result
End Function



Public Function CountAllCValues() As Long
    Dim total As Long
    total = 0
    
    Dim i As Long
    For i = 1 To m_vars.count
        Dim val As CValue
        Set val = m_vars(i)
        
        total = total + 1  ' Count this CValue
        total = total + CountNestedCValues(val)  ' Count children
    Next
    
    CountAllCValues = total
End Function

Private Function CountNestedCValues(val As CValue) As Long
    Dim count As Long
    count = 0
    
    If val.vType = vtObject Then
        ' Count property CValues
        Dim i As Long
        For i = 1 To val.objectProps.count
            Dim prop As CValue
            Set prop = val.objectProps(i)
            count = count + 1
            count = count + CountNestedCValues(prop)  ' Recursive
        Next
    ElseIf val.vType = vtArray Then
        ' Count array element CValues
        For i = 1 To val.arrayVal.count
            Dim elem As CValue
            Set elem = val.arrayVal(i)
            count = count + 1
            count = count + CountNestedCValues(elem)  ' Recursive
        Next
    End If
    
    CountNestedCValues = count
End Function

