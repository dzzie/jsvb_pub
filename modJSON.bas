Attribute VB_Name = "modJSON"
'Author:  David Zimmer <dzzie@yahoo.com>
'AI:      Claude.ai
'Site:    http://sandsprite.com
'License: MIT

Option Explicit

Public Function IsJSONMethod(methodName As String) As Boolean
    IsJSONMethod = (methodName = "parse" Or methodName = "stringify")
End Function

Public Function CallJSONMethod(methodName As String, args As Collection) As CValue
    Dim result As New CValue
    
    Select Case methodName
        Case "parse"
            ' JSON.parse(text)
            If args.count = 0 Then
                result.vType = vtUndefined
            Else
                Dim jsonText As String
                Dim arg As CValue
                Set arg = args(1)
                jsonText = arg.ToString()
                
                ' Parse JSON string to CValue
                Set result = ParseJSON(jsonText)
            End If
            
        Case "stringify"
            ' JSON.stringify(value)
            If args.count = 0 Then
                result.vType = vtUndefined
            Else
                Set arg = args(1)
                
                ' Convert CValue to JSON string
                result.vType = vtString
                result.strVal = StringifyJSON(arg)
            End If
            
        Case Else
            result.vType = vtUndefined
    End Select
    
    Set CallJSONMethod = result
End Function

' Parse JSON string into CValue
Public Function ParseJSON(jsonText As String) As CValue
    Dim result As New CValue
    Dim pos As Long
    pos = 1
    
    ' Skip whitespace
    pos = SkipWhitespace(jsonText, pos)
    
    ' Parse value
    Set result = ParseJSONValue(jsonText, pos)
    
    Set ParseJSON = result
End Function

Public Function SkipWhitespace(text As String, pos As Long) As Long
    Do While pos <= Len(text)
        Dim ch As String
        ch = Mid$(text, pos, 1)
        If ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Then
            pos = pos + 1
        Else
            Exit Do
        End If
    Loop
    SkipWhitespace = pos
End Function

Public Function ParseJSONValue(text As String, ByRef pos As Long) As CValue
    Dim result As New CValue
    
    pos = SkipWhitespace(text, pos)
    
    If pos > Len(text) Then
        result.vType = vtUndefined
        Set ParseJSONValue = result
        Exit Function
    End If
    
    Dim ch As String
    ch = Mid$(text, pos, 1)
    
    Select Case ch
        Case "{"
            Set result = ParseJSONObject(text, pos)
        Case "["
            Set result = ParseJSONArray(text, pos)
        Case """"
            Set result = ParseJSONString(text, pos)
        Case "t", "f"
            Set result = ParseJSONBoolean(text, pos)
        Case "n"
            Set result = ParseJSONNull(text, pos)
        Case "-", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
            Set result = ParseJSONNumber(text, pos)
        Case Else
            result.vType = vtUndefined
    End Select
    
    Set ParseJSONValue = result
End Function

' Update ParseJSONObject in CInterpreter.cls

Private Function ParseJSONObject(text As String, ByRef pos As Long) As CValue
    Dim result As New CValue
    result.vType = vtObject
    Set result.objectProps = New Collection
    Set result.objectKeys = New Collection
    
    pos = pos + 1  ' Skip {
    pos = SkipWhitespace(text, pos)
    
    ' Empty object
    If Mid$(text, pos, 1) = "}" Then
        pos = pos + 1
        Set ParseJSONObject = result
        Exit Function
    End If
    
    Do While pos <= Len(text)
        ' Parse key
        pos = SkipWhitespace(text, pos)
        If Mid$(text, pos, 1) <> """" Then Exit Do
        
        Dim key As CValue
        Set key = ParseJSONString(text, pos)
        
        ' Expect colon
        pos = SkipWhitespace(text, pos)
        If Mid$(text, pos, 1) <> ":" Then Exit Do
        pos = pos + 1
        
        ' Parse value
        Dim val As CValue
        Set val = ParseJSONValue(text, pos)
        
        ' Add to object
        On Error Resume Next
        result.objectProps.Remove key.strVal
        result.objectKeys.Remove key.strVal
        On Error GoTo 0
        result.objectProps.add val, key.strVal
        result.objectKeys.add key.strVal, key.strVal  ' NEW: Track the key
        
        ' Check for comma or end
        pos = SkipWhitespace(text, pos)
        If Mid$(text, pos, 1) = "}" Then
            pos = pos + 1
            Exit Do
        ElseIf Mid$(text, pos, 1) = "," Then
            pos = pos + 1
        Else
            Exit Do
        End If
    Loop
    
    Set ParseJSONObject = result
End Function


Public Function ParseJSONArray(text As String, ByRef pos As Long) As CValue
    Dim result As New CValue
    result.vType = vtArray
    Set result.arrayVal = New Collection
    
    pos = pos + 1  ' Skip [
    pos = SkipWhitespace(text, pos)
    
    ' Empty array
    If Mid$(text, pos, 1) = "]" Then
        pos = pos + 1
        Set ParseJSONArray = result
        Exit Function
    End If
    
    Do While pos <= Len(text)
        ' Parse value
        Dim val As CValue
        Set val = ParseJSONValue(text, pos)
        result.arrayVal.add val
        
        ' Check for comma or end
        pos = SkipWhitespace(text, pos)
        If Mid$(text, pos, 1) = "]" Then
            pos = pos + 1
            Exit Do
        ElseIf Mid$(text, pos, 1) = "," Then
            pos = pos + 1
        Else
            Exit Do
        End If
    Loop
    
    Set ParseJSONArray = result
End Function

Public Function ParseJSONString(text As String, ByRef pos As Long) As CValue
    Dim result As New CValue
    result.vType = vtString
    
    pos = pos + 1  ' Skip opening "
    
    Dim str As String
    str = ""
    
    Do While pos <= Len(text)
        Dim ch As String
        ch = Mid$(text, pos, 1)
        
        If ch = """" Then
            pos = pos + 1
            Exit Do
        ElseIf ch = "\" Then
            pos = pos + 1
            If pos <= Len(text) Then
                ch = Mid$(text, pos, 1)
                Select Case ch
                    Case "n": str = str & vbLf
                    Case "r": str = str & vbCr
                    Case "t": str = str & vbTab
                    Case "\": str = str & "\"
                    Case "/": str = str & "/"
                    Case """": str = str & """"
                    Case "b", "f": ' Ignore
                    Case "u"
                        ' Unicode escape (simplified)
                        If pos + 4 <= Len(text) Then
                            Dim hex As String
                            hex = Mid$(text, pos + 1, 4)
                            On Error Resume Next
                            str = str & ChrW$("&H" & hex)
                            On Error GoTo 0
                            pos = pos + 4
                        End If
                    Case Else
                        str = str & ch
                End Select
                pos = pos + 1
            End If
        Else
            str = str & ch
            pos = pos + 1
        End If
    Loop
    
    result.strVal = str
    
    ' SNEAKY AUTO-CONVERT: strict validation only - 0x hex strings or numbers as quoted strings we convert to int/bigint unsigned.
    If Len(str) > 0 And InStr(str, " ") = 0 Then
        ' Try hex (0x prefix)
        If Len(str) > 2 And Left$(str, 2) = "0x" Then
            Dim hexPart As String
            hexPart = Mid$(str, 3)
            
            If IsStrictHex(hexPart) Then
                If result.LoadNumFromStr(str) <> vtUndefined Then
                    ' Converted
                End If
            End If
        
        ' Try decimal (strict - only digits, optional leading -)
        ElseIf IsStrictNumeric(str) Then
            If result.LoadNumFromStr(str) <> vtUndefined Then
                ' Converted
            End If
        End If
    End If
    
    Set ParseJSONString = result
End Function

Private Function IsStrictHex(s As String) As Boolean
    If Len(s) = 0 Then Exit Function
    Dim i As Long, ch As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If Not ((ch >= "0" And ch <= "9") Or _
                (ch >= "a" And ch <= "f") Or _
                (ch >= "A" And ch <= "F")) Then
            Exit Function
        End If
    Next
    IsStrictHex = True
End Function

Private Function IsStrictNumeric(s As String) As Boolean
    If Len(s) = 0 Then Exit Function
    Dim i As Long
    Dim ch As String
    Dim start As Long
    
    start = 1
    ' Allow leading minus
    If Left$(s, 1) = "-" Then
        If Len(s) = 1 Then Exit Function  ' Just "-" is invalid
        start = 2
    End If
    
    ' Rest must be digits only
    For i = start To Len(s)
        ch = Mid$(s, i, 1)
        If Not (ch >= "0" And ch <= "9") Then
            Exit Function
        End If
    Next
    
    IsStrictNumeric = True
End Function

Public Function ParseJSONNumber(text As String, ByRef pos As Long) As CValue
    Dim result As New CValue
    result.vType = vtNumber
    
    Dim numStr As String
    numStr = ""
    
    ' Optional minus
    If Mid$(text, pos, 1) = "-" Then
        numStr = "-"
        pos = pos + 1
    End If
    
    ' Digits
    Do While pos <= Len(text)
        Dim ch As String
        ch = Mid$(text, pos, 1)
        If ch >= "0" And ch <= "9" Then
            numStr = numStr & ch
            pos = pos + 1
        Else
            Exit Do
        End If
    Loop
    
    ' Decimal point
    If pos <= Len(text) And Mid$(text, pos, 1) = "." Then
        numStr = numStr & "."
        pos = pos + 1
        
        Do While pos <= Len(text)
            ch = Mid$(text, pos, 1)
            If ch >= "0" And ch <= "9" Then
                numStr = numStr & ch
                pos = pos + 1
            Else
                Exit Do
            End If
        Loop
    End If
    
    ' Exponent
    If pos <= Len(text) Then
        ch = Mid$(text, pos, 1)
        If ch = "e" Or ch = "E" Then
            numStr = numStr & "E"
            pos = pos + 1
            
            If pos <= Len(text) Then
                ch = Mid$(text, pos, 1)
                If ch = "+" Or ch = "-" Then
                    numStr = numStr & ch
                    pos = pos + 1
                End If
            End If
            
            Do While pos <= Len(text)
                ch = Mid$(text, pos, 1)
                If ch >= "0" And ch <= "9" Then
                    numStr = numStr & ch
                    pos = pos + 1
                Else
                    Exit Do
                End If
            Loop
        End If
    End If
    
    result.numVal = CDbl(numStr)
    Set ParseJSONNumber = result
End Function

Public Function ParseJSONBoolean(text As String, ByRef pos As Long) As CValue
    Dim result As New CValue
    result.vType = vtBoolean
    
    If Mid$(text, pos, 4) = "true" Then
        result.boolVal = True
        pos = pos + 4
    ElseIf Mid$(text, pos, 5) = "false" Then
        result.boolVal = False
        pos = pos + 5
    End If
    
    Set ParseJSONBoolean = result
End Function

Public Function ParseJSONNull(text As String, ByRef pos As Long) As CValue
    Dim result As New CValue
    result.vType = vtNull
    
    If Mid$(text, pos, 4) = "null" Then
        pos = pos + 4
    End If
    
    Set ParseJSONNull = result
End Function

' Convert CValue to JSON string

Public Function StringifyJSONArray(arr As Collection) As String
    Dim result As String
    Dim i As Long
    
    result = "["
    
    For i = 1 To arr.count
        If i > 1 Then result = result & ","
        
        Dim val As CValue
        Set val = arr(i)
        result = result & StringifyJSON(val)
    Next
    
    result = result & "]"
    StringifyJSONArray = result
End Function

Private Function StringifyJSONObject(props As Collection, Keys As Collection) As String
    Dim result As String
    Dim i As Long
    Dim keyStr As String
    Dim val As CValue
    
    result = "{"
    
    For i = 1 To Keys.count
        If i > 1 Then result = result & ","
        
        ' Get key
        keyStr = Keys(i)
        
        ' Get value
        Set val = props(keyStr)
        
        ' Add to result
        result = result & """" & EscapeJSONString(keyStr) & """:" & StringifyJSON(val)
    Next
    
    result = result & "}"
    StringifyJSONObject = result
End Function

' Update StringifyJSON to pass keys:

Function StringifyJSON(val As CValue) As String
    Select Case val.vType
        Case vtUndefined
            StringifyJSON = "undefined"
            
        Case vtNull
            StringifyJSON = "null"
            
        Case vtBoolean
            StringifyJSON = IIf(val.boolVal, "true", "false")
            
        Case vtNumber
            StringifyJSON = CStr(val.numVal)
            
        Case vtString
            StringifyJSON = """" & EscapeJSONString(val.strVal) & """"
            
        Case vtArray
            StringifyJSON = StringifyJSONArray(val.arrayVal)
            
        Case vtObject
            StringifyJSON = StringifyJSONObject(val.objectProps, val.objectKeys)  ' CHANGED
            
        Case vtfunction
            StringifyJSON = "undefined"
            
        Case vtCOMObject
            StringifyJSON = "null"
            
        Case vtInt64
            Dim U64 As New ULong64
            U64.mode = mSigned
            U64.rawValue = val.int64Val
            StringifyJSON = U64.ToString(mUnsigned)
            
        Case Else
            StringifyJSON = "null"
    End Select
End Function






Public Function EscapeJSONString(str As String) As String
    Dim result As String
    Dim i As Long
    Dim ch As String
    
    result = ""
    
    For i = 1 To Len(str)
        ch = Mid$(str, i, 1)
        
        Select Case ch
            Case """"
                result = result & "\"""
            Case "\"
                result = result & "\\"
            Case vbCr
                result = result & "\r"
            Case vbLf
                result = result & "\n"
            Case vbTab
                result = result & "\t"
            Case Else
                result = result & ch
        End Select
    Next
    
    EscapeJSONString = result
End Function
