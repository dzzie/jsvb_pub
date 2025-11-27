Attribute VB_Name = "modGlobals"
'Author:  David Zimmer <dzzie@yahoo.com> + Claude.ai
'Site:    http://sandsprite.com
'License: MIT

Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' Helper functions (add to a module or class)

Public Declare Function DispCallFunc Lib "oleaut32.dll" ( _
    ByVal pvInstance As Long, ByVal offsetinVft As Long, _
    ByVal CallConv As Long, ByVal retTYP As Integer, _
    ByVal paCNT As Long, ByRef paTypes As Integer, _
    ByRef paValues As Long, ByRef ret As Variant) As Long
    
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Destination As Any, source As Any, ByVal length As Long)

Public Declare Function CreateProxyForProgIDRaw Lib "dynproxy.dll" (ByVal progId As Long, ByVal resolverDisp As Long) As Long
Public Declare Function CreateProxyForObjectRaw Lib "dynproxy.dll" (ByVal innerDispPtr As Long, ByVal resolverDispPtr As Long) As Long
Public Declare Sub ReleaseDispatchRaw Lib "dynproxy.dll" (ByVal pDisp As Long)
Public Declare Sub SetProxyResolverWins Lib "dynproxy.dll" (ByVal proxyPtr As Long, ByVal enable As Long)
Public Declare Sub ClearProxyNameCache Lib "dynproxy.dll" (ByVal proxyPtr As Long)
Public Declare Function CreateProxyForObjectRawEx Lib "dynproxy.dll" (ByVal innerPtr As Long, ByVal resolverPtr As Long, ByVal resolverWins As Long) As Long
Public Declare Sub SetProxyOverride Lib "dynproxy.dll" (ByVal proxyPtr As Long, ByVal nameBSTR As Long, ByVal dispid As Long)
Public Declare Function ComTypeName Lib "dynproxy.dll" (ByVal obj As IUnknown) As Variant
Public Declare Sub SendDbgMsg Lib "dynproxy.dll" (ByVal msg As String)   'to the PersistantDbgPrint window

Public Declare Function CallByNameEx Lib "dynproxy.dll" ( _
    ByVal obj As IUnknown, _
    ByVal memberName As String, _
    ByVal invokeFlags As Integer, _
    ByRef args() As Variant, _
    ByRef result As Variant, _
    ByRef isObject As Boolean _
) As Long

Public Const DISPATCH_METHOD = 1
Public Const DISPATCH_PROPERTYGET = 2
Public Const DISPATCH_PROPERTYPUT = 4
Public Const DISPATCH_PROPERTYPUTREF = 8

Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Public hLibDynProxy As Long



Sub push(ary, Value) 'this modifies parent ary object
    On Error GoTo init
    Dim X As Long
    X = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = Value
    Exit Sub
init:     ReDim ary(0): ary(0) = Value
End Sub

Function AryIsEmpty(ary) As Boolean
  Dim i As Long
  On Error GoTo oops
    i = UBound(ary)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function

Function ensureDynProxy() As Boolean
     
    If hLibDynProxy <> 0 Then
        ensureDynProxy = True
        Exit Function
    End If
    
    hLibDynProxy = LoadLibrary(App.path & "\dynproxy.dll")
    If hLibDynProxy <> 0 Then
        ensureDynProxy = True
    Else
        Debug.Print ">>> dynproxy.dll not found!"
    End If
             
End Function

Public Function PtrFromObject(obj As Object) As Long
    Dim unk As IUnknown
    Set unk = obj
    PtrFromObject = ObjPtr(unk)
End Function

'Public Function ObjectFromPtr(ByVal ptr As Long) As Object
'    Dim obj As Object
'    CopyMemory obj, ptr, 4
'    Set ObjectFromPtr = obj
'    CopyMemory obj, 0&, 4
'End Function

Public Function ObjectFromPtr(ByVal ptr As Long) As Object
    CopyMemory ObjectFromPtr, ptr, 4  ' Direct assignment, no AddRef!
End Function

    
' Helper: URL encoding
Public Function EncodeURIComponent(str As String) As String
    Dim result As String
    Dim i As Long
    Dim ch As String
    Dim code As Integer
    
    result = ""
    
    For i = 1 To Len(str)
        ch = Mid$(str, i, 1)
        code = Asc(ch)
        
        ' Safe characters: A-Z a-z 0-9 - _ . ! ~ * ' ( )
        If (code >= 65 And code <= 90) Or _
           (code >= 97 And code <= 122) Or _
           (code >= 48 And code <= 57) Or _
           ch = "-" Or ch = "_" Or ch = "." Or ch = "!" Or _
           ch = "~" Or ch = "*" Or ch = "'" Or ch = "(" Or ch = ")" Then
            result = result & ch
        Else
            ' Encode as %XX
            result = result & "%" & Right$("0" & hex$(code), 2)
        End If
    Next
    
    EncodeURIComponent = result
End Function

' Helper: URL encoding (preserves more characters for full URIs)
Public Function EncodeURI(str As String) As String
    Dim result As String
    Dim i As Long
    Dim ch As String
    Dim code As Integer
    
    result = ""
    
    For i = 1 To Len(str)
        ch = Mid$(str, i, 1)
        code = Asc(ch)
        
        ' Safe + reserved characters for URIs
        If (code >= 65 And code <= 90) Or _
           (code >= 97 And code <= 122) Or _
           (code >= 48 And code <= 57) Or _
           ch = "-" Or ch = "_" Or ch = "." Or ch = "!" Or _
           ch = "~" Or ch = "*" Or ch = "'" Or ch = "(" Or ch = ")" Or _
           ch = ":" Or ch = "/" Or ch = "?" Or ch = "#" Or ch = "[" Or ch = "]" Or _
           ch = "@" Or ch = "$" Or ch = "&" Or ch = "+" Or ch = "," Or ch = ";" Or ch = "=" Then
            result = result & ch
        Else
            result = result & "%" & Right$("0" & hex$(code), 2)
        End If
    Next
    
    EncodeURI = result
End Function

' Helper: URL decoding
Public Function DecodeURIComponent(str As String) As String
    Dim result As String
    Dim i As Long
    Dim ch As String
    
    result = ""
    i = 1
    
    Do While i <= Len(str)
        ch = Mid$(str, i, 1)
        
        If ch = "%" And i + 2 <= Len(str) Then
            ' Decode %XX
            Dim hex As String
            hex = Mid$(str, i + 1, 2)
            
            On Error Resume Next
            result = result & Chr$("&H" & hex)
            If Err.Number <> 0 Then
                result = result & ch
            Else
                i = i + 2  ' Skip the hex digits
            End If
            On Error GoTo 0
        ElseIf ch = "+" Then
            ' + means space in URL encoding
            result = result & " "
        Else
            result = result & ch
        End If
        
        i = i + 1
    Loop
    
    DecodeURIComponent = result
End Function

' Helper: Old-style escape
Public Function EscapeString(str As String) As String
    Dim result As String
    Dim i As Long
    Dim ch As String
    Dim code As Integer
    
    result = ""
    
    For i = 1 To Len(str)
        ch = Mid$(str, i, 1)
        code = Asc(ch)
        
        ' Characters that don't need escaping: A-Z a-z 0-9 @ * _ + - . /
        If (code >= 65 And code <= 90) Or _
           (code >= 97 And code <= 122) Or _
           (code >= 48 And code <= 57) Or _
           ch = "@" Or ch = "*" Or ch = "_" Or ch = "+" Or ch = "-" Or ch = "." Or ch = "/" Then
            result = result & ch
        ElseIf code < 256 Then
            result = result & "%" & Right$("0" & hex$(code), 2)
        Else
            ' Unicode: %uXXXX
            result = result & "%u" & Right$("000" & hex$(code), 4)
        End If
    Next
    
    EscapeString = result
End Function

' Helper: Old-style unescape
Public Function UnescapeString(str As String) As String
    Dim result As String
    Dim i As Long
    Dim ch As String
    
    result = ""
    i = 1
    
    Do While i <= Len(str)
        ch = Mid$(str, i, 1)
        
        If ch = "%" Then
            If i + 5 <= Len(str) And Mid$(str, i + 1, 1) = "u" Then
                ' Unicode: %uXXXX
                Dim uhex As String
                uhex = Mid$(str, i + 2, 4)
                
                On Error Resume Next
                result = result & ChrW$("&H" & uhex)
                If Err.Number <> 0 Then
                    result = result & ch
                Else
                    i = i + 5
                End If
                On Error GoTo 0
            ElseIf i + 2 <= Len(str) Then
                ' Regular: %XX
                Dim hex As String
                hex = Mid$(str, i + 1, 2)
                
                On Error Resume Next
                result = result & Chr$("&H" & hex)
                If Err.Number <> 0 Then
                    result = result & ch
                Else
                    i = i + 2
                End If
                On Error GoTo 0
            Else
                result = result & ch
            End If
        Else
            result = result & ch
        End If
        
        i = i + 1
    Loop
    
    UnescapeString = result
End Function

Function IsHexString(s As String) As Boolean
    ' Check if string is a hex number (0x... format)
    If Len(s) < 3 Then
        IsHexString = False
        Exit Function
    End If
    
    If LCase(Left$(s, 2)) <> "0x" Then
        IsHexString = False
        Exit Function
    End If
    
    ' Check if rest are hex digits
    Dim i As Long
    For i = 3 To Len(s)
        Dim c As String
        c = UCase(Mid$(s, i, 1))
        If Not ((c >= "0" And c <= "9") Or (c >= "A" And c <= "F")) Then
            IsHexString = False
            Exit Function
        End If
    Next
    
    IsHexString = True
End Function

