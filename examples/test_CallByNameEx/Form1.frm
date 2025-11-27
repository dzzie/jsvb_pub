VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'vb6 callbyname sucks to call things blindly.
'if the function returns an object you need to use a set=
'if you call the function twice to figure it out, you just called the function twice which isnt always right


Private Declare Sub SendDbgMsg Lib "dynproxy.dll" (ByVal msg As String)
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Var() As Any) As Long

Private Declare Function CallByNameEx Lib "dynproxy.dll" ( _
    ByVal obj As IUnknown, _
    ByVal memberName As String, _
    ByVal invokeFlags As Integer, _
    ByRef args() As Variant, _
    ByRef result As Variant, _
    ByRef isObject As Boolean _
) As Long

Private Const DISPATCH_METHOD = 1
Private Const DISPATCH_PROPERTYGET = 2
Private Const DISPATCH_PROPERTYPUT = 4
Private Const DISPATCH_PROPERTYPUTREF = 8

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

Private m_refCountTest As Long

Private Sub TestRefCounts()
    Dim obj1 As Object, obj2 As Object
    Dim r As Long, isObj As Boolean, v As Variant
    
    Dim args() As Variant
    ReDim args(1)
    args(0) = "1"
    args(1) = "b"
    
    ' Get object twice
    r = CallByNameEx(Me, "getme", DISPATCH_METHOD, args, v, isObj)
    Set obj1 = v
    
    r = CallByNameEx(Me, "getme", DISPATCH_METHOD, args, v, isObj)
    Set obj2 = v
    
    SendDbgMsg "obj1 = " & TypeName(obj1) & " ObjPtr=" & Hex(ObjPtr(obj1))
    SendDbgMsg "obj2 = " & TypeName(obj2) & " ObjPtr=" & Hex(ObjPtr(obj2))
    
    Set obj1 = Nothing
    SendDbgMsg "After obj1=Nothing, obj2 still valid: " & TypeName(obj2)
    
    Set obj2 = Nothing
    SendDbgMsg "After obj2=Nothing, form still alive: " & Me.Name
End Sub

 
Property Set MyObj(x As Object)
    SendDbgMsg ">> from vb6 Property Set MyObj=" & TypeName(x)
End Property

Property Get a() As Long
    a = 21
End Property

Property Let b(x)
    SendDbgMsg ">> from vb6 let b=" & x
End Property
 
Private Sub Form_Load()
    
    Dim h As Long
    h = LoadLibrary(App.Path & "\..\..\dynproxy.dll")
    If h = 0 Then End
    
    Dim v As Variant, isObj As Boolean, r As Long
    Dim vv As Variant
    
    SendDbgMsg "<cls>"
    
    TestRefCounts
    Exit Sub
    
    
    Dim args() As Variant
    ReDim args(1)
    args(0) = "1"
    args(1) = "b"
    
    SendDbgMsg ">> from vb6 Hex(VarPtrArray(args))=" & Hex(VarPtrArray(args))
    SendDbgMsg ">> from vb6 Hex(varptr(v))=" & Hex(VarPtr(v))
    
    r = CallByNameEx(Me, "getme", DISPATCH_METHOD, args, v, isObj)
    SendDbgMsg ">> from vb6 >> after getme return typename: " & TypeName(v) & " r=" & r & " isObj: " & isObj
    
    Erase args
    r = CallByNameEx(Me, "a", DISPATCH_PROPERTYGET, args, v, isObj)
    SendDbgMsg ">> from vb6 >> after get.a: " & v & " r=" & r & " isobj=" & isObj
    
    ReDim args(0)
    args(0) = 777
    r = CallByNameEx(Me, "b", DISPATCH_PROPERTYPUT, args, v, isObj)
    SendDbgMsg ">> from vb6 >> after let.b: " & v & " r=" & r & " isobj=" & isObj
   
    ReDim args(0)
    Set args(0) = Me  ' Object assignment
    r = CallByNameEx(Me, "MyObj", DISPATCH_PROPERTYPUTREF, args, v, isObj)
    SendDbgMsg ">> from vb6 >> after set.MyObj: r=" & r & " isobj=" & isObj

End Sub

Function getme(a As Long, b) As Object
    SendDbgMsg ">> from vb6 >> in getme! a=" & a & " b=" & b
    Set getme = Me
End Function

'problem this approach calls the method twice!
Function testCallByName()
    On Error Resume Next
    Dim v As Variant
    
    v = CallByName(Me, "getme", VbMethod)
    
    If Err.Number <> 0 Then
        Debug.Print Err.Number & ": " & Err.Description
        Set v = CallByName(Me, "getme", VbMethod)
    End If
    
    Debug.Print "Type: " & TypeName(v)
    
'    in getme!
'    450: Wrong number of arguments or invalid property assignment
'    in getme!
'    Type: Form1
End Function



