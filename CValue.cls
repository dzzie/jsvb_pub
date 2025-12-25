VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Author:  David Zimmer <dzzie@yahoo.com>
'AI:      Claude.ai
'Site:    http://sandsprite.com
'License: MIT
'part of intrepreter support

'------------------------------------------------------------
' CValue.cls - Runtime value wrapper
' "Every value in JavaScript is one of these"
'------------------------------------------------------------

Option Explicit

Public Enum ValueType
    vtUndefined = 0
    vtNull = 1
    vtBoolean = 2
    vtNumber = 3
    vtString = 4
    vtObject = 5
    vtInt64 = 6
    vtCOMObject = 7
    vtfunction = 8
    vtArray = 9         ' NEW: Arrays
End Enum


' The actual value storage
Public vType As ValueType
Public numVal As Double
Public strVal As String
Public boolVal As Boolean
Public objVal As Object        ' for functions this will be a CFunction
Public int64Val As Currency
Public arrayVal As Collection      ' For arrays

'Parallel Collections - index - matched:
    'objectKeys(1) = "name" -> objectProps(1) = CValue("John")
    'objectKeys(2) = "age" -> objectProps(2) = CValue(30)
    'objectKeys(3) = "city" -> objectProps(3) = CValue("NYC")
    
Public objectProps As Collection    ' For object properties (key-value pairs)
Public objectKeys As Collection

Const dbg_mode As Boolean = False

Function GetTypeName() As String
    Select Case vType
        Case vtUndefined: GetTypeName = "undefined"
        Case vtNull: GetTypeName = "null"
        Case vtBoolean: GetTypeName = "boolean"
        Case vtNumber: GetTypeName = "number"
        Case vtString: GetTypeName = "string"
        Case vtObject: GetTypeName = "object"
        Case vtArray: GetTypeName = "array"
        Case vtfunction: GetTypeName = "function"
        Case vtInt64: GetTypeName = "bigint"
        Case vtCOMObject: GetTypeName = "object"
        Case Else: GetTypeName = "unknown"
    End Select
End Function

' Helper: Get array length
Public Function ArrayLength() As Long
    If vType = vtArray And Not arrayVal Is Nothing Then
        ArrayLength = arrayVal.count
    Else
        ArrayLength = 0
    End If
End Function

Public Function LoadNumFromStr(v As Variant, Optional isUnsigned As Boolean = True) As ValueType
    Dim sVal As String
    Dim U64 As New ULong64
    
    On Error GoTo errHandler
    
    If VarType(v) = vbString Then
        sVal = Trim$(v)
    ElseIf IsNumeric(v) Then
        sVal = CStr(v)
    Else
        GoTo errHandler
    End If
    
    If Len(sVal) = 0 Then GoTo errHandler
    If Not U64.fromString(sVal, mUnsigned) Then GoTo errHandler
    
    If U64.Is32BitSafeEx(isUnsigned) Then
        vType = vtNumber
        numVal = CDbl(U64.lo)
    Else
        vType = vtInt64
        int64Val = U64.rawValue
    End If
    
    LoadNumFromStr = vType
    Exit Function
    
errHandler:
    vType = vtUndefined
    LoadNumFromStr = vType
End Function

'this is the default method of the class
Public Function ToString() As String
Attribute ToString.VB_UserMemId = 0
    Select Case vType
        Case vtUndefined:
                ToString = "undefined"
        Case vtNull:
                ToString = "null"
        Case vtBoolean:
                ToString = IIf(boolVal, "true", "false")
        Case vtNumber:
                ToString = CStr(numVal)
        Case vtString:
                ToString = strVal
        Case vtObject:
                ' Check if it's a Date object
                Dim timestampProp As CValue
                Set timestampProp = GetProperty("_timestamp")
                
                If Not timestampProp Is Nothing Then
                    ' It's a Date object - format it
                    Dim dateVal As Date
                    dateVal = DateAdd("s", timestampProp.numVal / 1000, #1/1/1970#)
                    ToString = Format$(dateVal, "ddd mmm dd yyyy hh:nn:ss")
                Else
                    ToString = "[object Object]"
                End If
        Case vtArray:
                ToString = "[Array]"
        Case vtfunction:
                If Not objVal Is Nothing Then
                    If Len(objVal.Name) > 0 Then
                        ToString = "[Function: " & objVal.Name & "]"
                    Else
                        ToString = "[Function (anonymous)]"
                    End If
                Else
                    ToString = "[Function]"
                End If
            
        Case vtInt64:
                Dim U64 As New ULong64
                U64.mode = mSigned
                U64.rawValue = int64Val
                U64.use0x = True
                'ToString = U64.ToString(mUnsigned) '& "n"
                ToString = U64.ToString(mHex) '& "n"
            
        Case vtCOMObject:
                ' Show the actual COM type name
                If Not objVal Is Nothing Then
                    On Error Resume Next
                    ToString = "[" & TypeName(objVal) & "]"
                    If Err.Number <> 0 Then
                        ToString = "[COM Object]"
                    End If
                    On Error GoTo 0
                Else
                    ToString = "[COM Object (null)]"
                End If
            
        Case Else:
                ToString = "[unknown]"
    End Select
End Function

Public Function ToBoolean() As Boolean
    Select Case vType
        Case vtUndefined, vtNull:
            ToBoolean = False
        Case vtBoolean:
            ToBoolean = boolVal
        Case vtNumber:
            ToBoolean = (numVal <> 0)
        Case vtString:
            ToBoolean = (Len(strVal) > 0)
        Case vtObject, vtArray, vtfunction, vtCOMObject:
            ToBoolean = True  ' Objects always truthy
        Case Else:
            ToBoolean = False
    End Select
End Function




' Helper: Coerce to number
Public Function ToNumber() As Double
    Select Case vType
        Case vtUndefined:
            ToNumber = 0
        Case vtNull:
            ToNumber = 0
        Case vtBoolean:
            ToNumber = IIf(boolVal, 1, 0)
        Case vtNumber:
            ToNumber = numVal
        Case vtString:
            On Error Resume Next
            ToNumber = CDbl(strVal)
            If Err.Number <> 0 Then ToNumber = 0
            On Error GoTo 0
        Case vtInt64:
            ' Convert BigInt to Number (may lose precision for large values!)
            Dim U64 As New ULong64
            U64.rawValue = int64Val
            
            ' Get hi and lo parts
            Dim hi As Long, lo As Long
            U64.GetLongs hi, lo
            
            ' Convert to double (handle unsigned longs properly)
            Dim dblHi As Double, dblLo As Double
            dblHi = CDbl(hi)
            If dblHi < 0 Then dblHi = dblHi + 4294967296#
            
            dblLo = CDbl(lo)
            If dblLo < 0 Then dblLo = dblLo + 4294967296#
            
            ToNumber = (dblHi * 4294967296#) + dblLo
        Case Else:
            ToNumber = 0
    End Select
End Function

' Convert to Int64 (Currency)
Public Function ToInt64() As Currency
    Select Case vType
        Case vtInt64:
            ToInt64 = int64Val
            
        Case vtNumber:
            ' Convert double to Currency properly
            Dim U64 As New ULong64
            Dim hi As Long, lo As Long
            Dim dblVal As Double
            dblVal = numVal
            
            If dblVal < 0 Then
                ' Negative number - convert to two's complement
                Dim posVal As Double
                posVal = 1.84467440737096E+19 + dblVal    ' 2^64 + negative value
                hi = DoubleToSignedLong(Int(posVal / 4294967296#))
                lo = DoubleToSignedLong(posVal - (Int(posVal / 4294967296#) * 4294967296#))
            ElseIf dblVal >= 4294967296# Then
                ' Large positive number > 32-bit
                hi = DoubleToSignedLong(Int(dblVal / 4294967296#))
                lo = DoubleToSignedLong(dblVal - (Int(dblVal / 4294967296#) * 4294967296#))
            Else
                ' Positive number in 32-bit range
                hi = 0
                lo = DoubleToSignedLong(dblVal)
            End If
            
            U64.SetLongs hi, lo
            ToInt64 = U64.rawValue
            
        Case vtString:
            ' Parse string as Int64
            Dim u64str As New ULong64
            u64str.mode = mSigned
            If u64str.fromString(strVal, mUnsigned) Then
                ToInt64 = u64str.rawValue
            Else
                ToInt64 = 0
            End If
            
        Case Else:
            ToInt64 = 0
    End Select
End Function

' Helper function: Convert unsigned double value to signed Long
' Handles values from 0 to 4294967295 (0xFFFFFFFF)
Private Function DoubleToSignedLong(ByVal dblVal As Double) As Long
    ' If value is > 2147483647, it wraps to negative in signed representation
    If dblVal > 2147483647# Then
        DoubleToSignedLong = CLng(dblVal - 4294967296#)  ' Subtract 2^32 to get signed equivalent
    Else
        DoubleToSignedLong = CLng(dblVal)
    End If
End Function

Private Sub Class_Initialize()
    ' Initialize collections if needed
    Set objectProps = New Collection
    Set objectKeys = New Collection
    Set arrayVal = New Collection
End Sub

' ----------------------------------------------------------------
' OBJECT PROPERTY ACCESS
' ----------------------------------------------------------------

' Get a property from this object
' ================================================================
' OBJECT PROPERTY ACCESS - Pure VB6 Collection
' ================================================================

' Get a property from this object
'Case sensitive!
'Public Function GetProperty(propName As String) As CValue
'
'    If dbg_mode Then Debug.Print ">>> CValue.GetProperty ENTRY: searching for '" & propName & "'"
'    If dbg_mode Then Debug.Print ">>> objectKeys.Count = " & objectKeys.count
'
'
'    If vType <> vtObject Then
'        'Debug.Print ">>> CValue.GetProperty: Not an object! (type=" & vType & ")"
'        Set GetProperty = Nothing
'        Exit Function
'    End If
'
'    If objectKeys Is Nothing Or objectProps Is Nothing Then
'        If dbg_mode Then Debug.Print ">>> CValue.GetProperty: No properties!"
'        Set GetProperty = Nothing
'        Exit Function
'    End If
'
'    ' Find the property by key (linear search)
'    Dim i As Long
'    For i = 1 To objectKeys.count
'        If objectKeys(i) = propName Then
'            Set GetProperty = objectProps(i)
'            If dbg_mode Then Debug.Print ">>> CValue.GetProperty('" & propName & "') -> found at index " & i
'            Exit Function
'        End If
'    Next
'
'    ' Not found
'    If dbg_mode Then Debug.Print ">>> CValue.GetProperty('" & propName & "') -> not found"
'    Set GetProperty = Nothing
'End Function

Public Function GetProperty(propName As String) As CValue
    On Error Resume Next
    
    ' Direct lookup by key (case-insensitive!)
    Set GetProperty = objectProps(propName)
    
    If Err.Number <> 0 Then
        Set GetProperty = Nothing
    End If
    
    On Error GoTo 0
End Function

''case insensitive..
'Public Function GetProperty(propName As String) As CValue
'
'    If dbg_mode Then Debug.Print ">>> CValue.GetProperty ENTRY: searching for '" & propName & "'"
'    If dbg_mode Then Debug.Print ">>> objectKeys.Count = " & objectKeys.count
'
'    If vType <> vtObject Then
'        If dbg_mode Then Debug.Print ">>> CValue.GetProperty: Not an object! (type=" & vType & ")"
'        Set GetProperty = Nothing
'        Exit Function
'    End If
'
'    If objectKeys Is Nothing Or objectProps Is Nothing Then
'        If dbg_mode Then Debug.Print ">>> CValue.GetProperty: No properties!"
'        Set GetProperty = Nothing
'        Exit Function
'    End If
'
'    ' Case-INSENSITIVE search!
'    Dim searchName As String
'    searchName = LCase$(propName)
'
'    Dim i As Long
'    For i = 1 To objectKeys.count
'        If LCase$(objectKeys(i)) = searchName Then  ' ? Case insensitive!
'            Set GetProperty = objectProps(i)
'            If dbg_mode Then Debug.Print ">>> CValue.GetProperty('" & propName & "') -> found at index " & i
'            Exit Function
'        End If
'    Next
'
'    If dbg_mode Then Debug.Print ">>> CValue.GetProperty('" & propName & "') -> not found"
'    Set GetProperty = Nothing
'End Function

' Set a property on this object
'case sensitive
Public Sub SetProperty(propName As String, Value As CValue)

    ' Allow both vtObject AND vtCOMObject to have properties!
    If vType <> vtObject And vType <> vtCOMObject Then
        Err.Raise 5, "CValue.SetProperty", "Not an object"
    End If
    
    ' Initialize collections if needed
    If objectKeys Is Nothing Then Set objectKeys = New Collection
    If objectProps Is Nothing Then Set objectProps = New Collection
    
    ' Find if property exists
    Dim foundIndex As Long
    foundIndex = 0
    
    Dim i As Long
    For i = 1 To objectKeys.count
        If objectKeys(i) = propName Then
            foundIndex = i
            Exit For
        End If
    Next
    
    If foundIndex > 0 Then
        ' UPDATE: Remove old, add new at same position
        objectKeys.Remove foundIndex
        objectProps.Remove foundIndex
        
        ' Re-add at same position
        If foundIndex <= objectKeys.count Then
            objectKeys.add propName, , foundIndex
            objectProps.add Value, , foundIndex
        Else
            objectKeys.add propName
            objectProps.add Value
        End If
        
        If dbg_mode Then Debug.Print ">>> CValue.SetProperty('" & propName & "') = " & Value.vType & " (updated)"
    Else
        ' NEW: Just add
        objectKeys.add propName
        objectProps.add Value, propName
        If dbg_mode Then Debug.Print ">>> CValue.SetProperty('" & propName & "') = " & Value.vType & " (new)"
    End If
End Sub

' Check if property exists
Public Function HasProperty(propName As String) As Boolean

    If (vType <> vtObject And vType <> vtCOMObject) Or objectKeys Is Nothing Then
        HasProperty = False
        Exit Function
    End If
    
    ' Linear search through keys
    Dim i As Long
    For i = 1 To objectKeys.count
        If objectKeys(i) = propName Then
            HasProperty = True
            Exit Function
        End If
    Next
    
    HasProperty = False
End Function

' Get all property names - returns Collection
Public Function GetPropertyNames() As Collection
    Set GetPropertyNames = objectKeys
End Function

' Get property count
Public Function PropertyCount() As Long
    If (vType <> vtObject And vType <> vtCOMObject) And Not objectKeys Is Nothing Then
        PropertyCount = objectKeys.count
    Else
        PropertyCount = 0
    End If
End Function

' Iterate properties (for debugging)
Public Function DumpProperties() As String
    If (vType <> vtObject And vType <> vtCOMObject) Or objectKeys Is Nothing Then
        DumpProperties = "{}"
        Exit Function
    End If
    
    Dim result As String
    result = "{ "
    
    Dim i As Long
    For i = 1 To objectKeys.count
        Dim key As String
        Dim val As CValue
        
        key = objectKeys(i)
        Set val = objectProps(i)
        
        If i > 1 Then result = result & ", "
        result = result & key & ": " & val.ToString()
    Next
    
    result = result & " }"
    DumpProperties = result
End Function
