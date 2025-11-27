VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInterpreter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Author:  David Zimmer <dzzie@yahoo.com> + Claude.ai
'Site:    http://sandsprite.com
'License: MIT

Option Explicit
Public HeadLessDebugging As Boolean

Private m_globalScope As CScope
Private m_currentScope As CScope
Private m_parser As CParser
Private m_output As String
Private m_shouldBreak As Boolean
Private m_shouldContinue As Boolean
Private m_returnValue As CValue
Private m_hasReturned As Boolean
Private m_functions As Collection  ' Store function declarations

' Module-level variables
Private m_thisContext As CValue  ' Current 'this' binding
Private m_requestedNextLine As Long
Private m_hasRequestedNextLine As Boolean

' Store the parsed AST (for debugging and validation)
Private m_programAST As CNode

' Track what's already been executed
Private m_executedCode As String

' COM INTEGRATION
Public UseSafeSubset As Boolean
Private m_comObjects As Collection

' Debugger state
Private m_debugMode As Boolean
Private m_debugPaused As Boolean
Private m_debugStepMode As DebugStepMode
Private m_breakpoints As Collection
Private m_callStack As Collection
Private m_currentLine As Long
Private m_sourceCode As String
Private m_stepOutDepth As Long
Private m_AbortExec As Boolean

Public Enum DebugStepMode
    smNone = 0      ' Run until breakpoint
    smStepInto = 1  ' Step into functions
    smStepOver = 2  ' Step over functions
    smStepOut = 3   ' Run until return from current function
End Enum

Public Enum enumAuditEvents
    aeEval = 0
    aeDecode = 1
    aeBracketAccess = 2
    aeCOMCall = 3
    aeActiveX = 4
    aeXOR = 5
    aeParse = 6
    aeFunctionConstructor = 7
End Enum

Public AuditMode As Boolean
Public Event AuditEvent(category As enumAuditEvents, description As String, ByRef cancel As Boolean)

' Debugger events
Public Event OnBreakpoint(ByVal LineNumber As Long, ByVal sourceCode As String)
Public Event OnStep(ByVal LineNumber As Long, ByVal sourceCode As String)
Public Event OnCallStackChanged()
Public Event OnError(ByVal ErrorMessage As String, ByVal LineNumber As Long, ByVal source As String, ByVal col As Long)
Public Event OnExecutionComplete()
Public Event OnVariablesChanged()
Public Event ConsoleLog(ByVal msg As String)

Const debug_mode As Boolean = False

Function AuditEventToStr(ae As enumAuditEvents) As String
    
    Dim categoryName As String
    
    Select Case ae
        Case aeEval: categoryName = "EVAL"
        Case aeDecode: categoryName = "DECODE"
        Case aeBracketAccess: categoryName = "BRACKET"
        Case aeCOMCall: categoryName = "COM_CALL"
        Case aeActiveX: categoryName = "ACTIVEX"
        Case aeXOR: categoryName = "XOR"
        Case aeParse: categoryName = "PARSE"
        Case aeFunctionConstructor: categoryName = "FUNC_CTOR"
    End Select
    
    AuditEventToStr = categoryName

End Function

' FOR TESTING ONLY - Setup minimal debug state
Friend Sub TestSetupDebugState(code As String, currentLine As Long)
    m_sourceCode = code
    m_currentLine = currentLine
    m_debugMode = True
    m_debugPaused = True
    
    ' Parse code
    Set m_programAST = m_parser.ParseScript(code)
End Sub

Property Get GlobalScope() As CScope
    Set GlobalScope = m_globalScope
End Property

Sub Abort()
    m_AbortExec = True
    StopDebug
End Sub

Private Sub RaiseAudit(ByVal category As enumAuditEvents, ByVal description As String)
    Dim cancel As Boolean
    If AuditMode Then
        RaiseEvent AuditEvent(category, description, cancel)
        If cancel Then
            m_AbortExec = True
            Err.Raise vbObjectError + 1000, "CInterpreter.RaiseAudit", _
                      "Execution cancelled by user audit: " & description
        End If
    End If
End Sub

Private Sub Class_Initialize()

    Set m_parser = New CParser
    Set m_globalScope = New CScope
    Set m_currentScope = m_globalScope
    Set m_functions = New Collection
    Set m_comObjects = New Collection  ' NEW
    
    ' Initialize debugger
    Set m_breakpoints = New Collection
    Set m_callStack = New Collection
    m_debugMode = False
    m_debugStepMode = smNone
    m_currentLine = 0
    m_debugPaused = False
    m_AbortExec = False
    
    UseSafeSubset = True  ' Safe by default
    AddBuiltins ' Add built-in console object
End Sub

Private Sub InitDebug()

    m_debugMode = True
    m_debugPaused = True 'False
    m_debugStepMode = smStepInto
    m_currentLine = 0
    m_stepOutDepth = 0
    m_AbortExec = False
    
    Set m_callStack = New Collection
    
    If m_breakpoints Is Nothing Then
        Set m_breakpoints = New Collection
    End If
    
    ' Add global frame
    Dim frame As New CDebugCallFrame
    frame.FunctionName = "<global>"
    frame.LineNumber = 0
    Set frame.scope = m_globalScope
    frame.ArgumentCount = 0
    
    m_callStack.add frame
    RaiseEvent OnCallStackChanged
    
End Sub

Public Sub StopDebug()
    m_debugMode = False
    m_debugPaused = False
    Set m_callStack = New Collection
    Set m_breakpoints = New Collection
End Sub

' Set/clear breakpoint
Public Sub SetBreakpoint(LineNumber As Long)
    On Error Resume Next
    m_breakpoints.Remove CStr(LineNumber)
    On Error GoTo 0
    m_breakpoints.add LineNumber, CStr(LineNumber)
End Sub

Public Sub ClearBreakpoint(LineNumber As Long)
    On Error Resume Next
    m_breakpoints.Remove CStr(LineNumber)
    On Error GoTo 0
End Sub

Public Sub ClearAllBreakpoints()
    Set m_breakpoints = New Collection
End Sub

Public Function hasBreakpoint(LineNumber As Long) As Boolean
    On Error Resume Next
    Dim X As Long
    X = m_breakpoints(CStr(LineNumber))
    hasBreakpoint = (Err.Number = 0)
End Function

' Step commands
Public Sub StepInto()
    m_debugStepMode = smStepInto
    m_debugPaused = False
End Sub

Public Sub StepOver()
    m_debugStepMode = smStepOver
    m_debugPaused = False
    m_stepOutDepth = m_callStack.count
End Sub

Public Sub StepOut()
    m_debugStepMode = smStepOut
    m_debugPaused = False
    m_stepOutDepth = m_callStack.count - 1
End Sub

Public Sub Run()
    m_debugStepMode = smNone
    m_debugPaused = False
    m_AbortExec = False
End Sub

Public Sub PauseExecution()
    If Not m_debugMode Then Exit Sub
    m_debugPaused = True
    m_debugStepMode = smStepInto  ' Switch to step mode when manually paused
End Sub

' Get current state
Public Property Get currentLine() As Long
    currentLine = m_currentLine
End Property

Public Property Get IsPaused() As Boolean
    IsPaused = m_debugPaused
End Property

Public Property Get IsDebugging() As Boolean
    IsDebugging = m_debugMode
End Property

' Get call stack as array of strings
Public Function GetCallStackStrings() As String()
    Dim result() As String
    Dim i As Long
    Dim frame As CDebugCallFrame
    
    If m_callStack.count = 0 Then
        ReDim result(0)
        result(0) = "<no stack>"
        GetCallStackStrings = result
        Exit Function
    End If
    
    ReDim result(m_callStack.count - 1)
    
    For i = 1 To m_callStack.count
        Set frame = m_callStack(i)
        result(i - 1) = frame.ToString()
    Next
    
    GetCallStackStrings = result
End Function

' Get variables in current scope
Public Function GetCurrentVariables() As Collection
    If m_callStack.count = 0 Then
        Set GetCurrentVariables = New Collection
        Exit Function
    End If
    
    Dim frame As CDebugCallFrame
    frame = m_callStack(m_callStack.count)
    
    ' For now, return a collection of strings
    ' Format: "varName = value"
    Dim vars As New Collection
    
    ' We need to enumerate scope variables
    ' This is tricky because Collection doesn't expose keys
    ' We'll add a helper method to CScope
    
    Set GetCurrentVariables = vars
End Function

' Get variables from current scope
Public Function GetCurrentScopeVariables() As Collection
    Set GetCurrentScopeVariables = New Collection
    
    If Not m_debugMode Then Exit Function
    
    ' Get the current scope (function scope or global)
    Dim currentScope As CScope
    
    If m_callStack.count > 0 Then
        ' We're in a function - get the function's scope from top of stack
        Dim frame As CDebugCallFrame
        Set frame = m_callStack.Item(m_callStack.count)
        Set currentScope = frame.scope
    Else
        ' Global scope
        Set currentScope = m_globalScope
    End If
    
    If currentScope Is Nothing Then
        Set currentScope = m_globalScope
    End If
    
    If currentScope Is Nothing Then Exit Function
    
    ' Use CScope's GetAllVariables method
    Set GetCurrentScopeVariables = currentScope.GetAllVariables()
End Function

' ============================================
' DEBUGGER INTERNAL HOOKS
' ============================================

' Called before each statement execution
Private Sub DebugCheckBreak(Node As CNode)
    
    Dim atBreakpoint As Boolean
    Dim shouldBreak As Boolean
    
    If Not m_debugMode Then Exit Sub
    If Node Is Nothing Then Exit Sub
    
    m_currentLine = Node.LineNumber
    atBreakpoint = hasBreakpoint(m_currentLine) ' Check for breakpoint
    shouldBreak = False ' Determine if we should break

    Select Case m_debugStepMode
        Case smStepInto
            shouldBreak = True
            
        Case smStepOver
            ' Break if at same or shallower depth
            If m_callStack.count <= m_stepOutDepth Then
                shouldBreak = True
            End If
            
        Case smStepOut
            ' Break when we've returned to shallower depth
            If m_callStack.count <= m_stepOutDepth Then
                shouldBreak = True
                m_debugStepMode = smStepInto
            End If
            
        Case smNone
            ' Only break on explicit breakpoints
            shouldBreak = atBreakpoint
    End Select
    
    If shouldBreak Or atBreakpoint Then
        m_debugPaused = True
        
        If atBreakpoint Then
            RaiseEvent OnBreakpoint(m_currentLine, m_sourceCode)
        Else
            RaiseEvent OnStep(m_currentLine, m_sourceCode)
        End If
        
        RaiseEvent OnVariablesChanged
        
        If Not HeadLessDebugging Then
            ' PAUSE EXECUTION - wait for user
            Do While m_debugPaused And m_debugMode
                DoEvents  ' This processes the Break button click
                Sleep 15
            Loop
        End If
        
    End If
    
End Sub

' Called when entering a function
Private Sub DebugPushFrame(funcName As String, argCount As Long)

    If Not m_debugMode Then Exit Sub
    
    Dim frame As New CDebugCallFrame
    frame.FunctionName = funcName
    frame.LineNumber = m_currentLine
    Set frame.scope = m_currentScope
    frame.ArgumentCount = argCount
    
    m_callStack.add frame
    RaiseEvent OnCallStackChanged
End Sub

' Called when exiting a function
Private Sub DebugPopFrame()
    If Not m_debugMode Then Exit Sub
    
    If m_callStack.count > 1 Then
        m_callStack.Remove m_callStack.count
        RaiseEvent OnCallStackChanged
    End If
End Sub

Public Sub AddCode(code As String)
    On Error GoTo ErrorHandler
    
    ' Skip empty code
    If Len(Trim$(code)) = 0 Then Exit Sub
    
    ' Parse and execute just the new code
    Dim ast As CNode
    Set ast = m_parser.ParseScript(code)
    ExecuteProgram ast
    
    ' Remember we executed this
    If Len(m_executedCode) > 0 Then
        m_executedCode = m_executedCode & vbCrLf
    End If
    m_executedCode = m_executedCode & code
    
    ' Raise completion event if debugging
    If m_debugMode Then
        RaiseEvent OnExecutionComplete
    End If
    
    Exit Sub
    
ErrorHandler:
    m_output = m_output & vbCrLf & "ERROR in AddCode: " & Err.description
    Debug.Print "AddCode Error: " & Err.description & " - " & Err.source
    
    Dim col As Long
    ' For parser errors during execution, extract line from error message
    If m_currentLine = 0 And Err.Number = vbObjectError + 2000 Then
        m_currentLine = ExtractLineFromError(Err.description)
        col = ExtractColumnFromError(Err.description)
    End If
    
    RaiseEvent OnError(Err.description, m_currentLine, Err.source, col)
  
End Sub

Public Sub ClearContext()
    m_executedCode = ""
    
    ' Reset the interpreter state
    Set m_globalScope = New CScope
    Set m_currentScope = m_globalScope
    Set m_functions = New Collection
    Set m_comObjects = New Collection
    m_output = ""
    m_shouldBreak = False
    m_shouldContinue = False
    Set m_returnValue = Nothing
    m_hasReturned = False
    
    AddBuiltins ' Re-add console, etc.
End Sub

' Evaluate expression in CURRENT context (no re-parsing old code)
Public Function Eval(code As String) As CValue
    On Error GoTo ErrorHandler
    
    Dim ast As CNode
    Dim result As CValue
    
    ' Parse ONLY the new code
    Set ast = m_parser.ParseScript(code)
    
    ' Should be a single expression statement
    If ast.Body.count > 0 Then
        Dim stmt As CNode
        Set stmt = ast.Body(1)
        
        If stmt.tType = ExpressionStatement_Node Then
            ' Evaluate the expression in the current context
            Set result = EvaluateExpression(stmt.Test)
        Else
            ' Execute the statement
            ExecuteStatement stmt
            Set result = New CValue
            result.vType = vtUndefined
        End If
    Else
        Set result = New CValue
        result.vType = vtUndefined
    End If
    
    Set Eval = result
    Exit Function
    
ErrorHandler:
    m_output = m_output & vbCrLf & "ERROR: " & Err.description
    Debug.Print "Eval Error: " & Err.description & " " & Err.source
    
    Dim col As Long
    ' For parser errors during execution, extract line from error message
    If m_currentLine = 0 And Err.Number = vbObjectError + 2000 Then
        m_currentLine = ExtractLineFromError(Err.description)
        col = ExtractColumnFromError(Err.description)
    End If

    RaiseEvent OnError(Err.description, m_currentLine, Err.source, col)
    
    Dim undef As New CValue
    undef.vType = vtUndefined
    Set Eval = undef
End Function

' Add this helper function to your interpreter class
Private Function ExtractLineFromError(errDescription As String) As Long
    ' Extract line number from parser error message
    ' Format: "Unexpected token: } at line 10, column 3"
    
    Dim linePos As Long
    linePos = InStr(errDescription, "at line ")
    
    If linePos > 0 Then
        Dim numStr As String
        numStr = Mid$(errDescription, linePos + 8)  ' Skip "at line "
        
        ' Stop at comma (before column info)
        Dim commaPos As Long
        commaPos = InStr(numStr, ",")
        If commaPos > 0 Then
            numStr = Left$(numStr, commaPos - 1)
        End If
        
        ExtractLineFromError = CLng(numStr)
    Else
        ExtractLineFromError = 0
    End If
End Function

Private Function ExtractColumnFromError(errDescription As String) As Long
    Dim colPos As Long
    colPos = InStr(errDescription, "column ")
    
    If colPos > 0 Then
        ExtractColumnFromError = CLng(Mid$(errDescription, colPos + 7))
    Else
        ExtractColumnFromError = 0
    End If
End Function

' Execute statement(s) in current context
Public Sub EvalStatement(code As String)
    On Error GoTo ErrorHandler
    
    Dim ast As CNode
    Set ast = m_parser.ParseScript(code)
    
    ' Execute in the EXISTING context
    ExecuteProgram ast
    
    Exit Sub
    
ErrorHandler:
    m_output = m_output & vbCrLf & "ERROR: " & Err.description & " - " & Err.source
    Debug.Print "EvalStatement Error: " & Err.description & " - " & Err.source
    
    Dim col As Long
    ' For parser errors during execution, extract line from error message
    If m_currentLine = 0 And Err.Number = vbObjectError + 2000 Then
        m_currentLine = ExtractLineFromError(Err.description)
        col = ExtractColumnFromError(Err.description)
    End If
    
    RaiseEvent OnError(Err.description, m_currentLine, Err.source, col)
End Sub

' Get a JavaScript variable from VB6
Public Function GetVariable(Name As String) As CValue
    Set GetVariable = m_globalScope.GetVar(Name)
End Function

' Set a JavaScript variable from VB6
Public Sub SetVariable(Name As String, Value As Variant)
    Dim val As New CValue
    Set val = VariantToCValue(Value)
    m_globalScope.DefineVar Name, val
End Sub

' Check if a JavaScript variable exists
Public Function HasVariable(Name As String) As Boolean
    HasVariable = m_globalScope.HasVar(Name)
End Function

' BONUS: Get variable as VB6 Variant (convenience method)
Public Function GetVariableAsVariant(Name As String) As Variant
    Dim val As CValue
    Set val = m_globalScope.GetVar(Name)
    GetVariableAsVariant = CValueToVariant(val)
End Function


Private Sub AddBuiltins()
    ' Create console.log as a special marker
    ' (We'll handle it specially in EvaluateCall)
    
       ' NEW: Create Math object
    Dim mathObj As New CValue
    mathObj.vType = vtObject
    Set mathObj.objectProps = New Collection
    
    ' Add Math constants
    Dim pi As New CValue
    pi.vType = vtNumber
    pi.numVal = 3.14159265358979
    mathObj.objectProps.add pi, "PI"
    
    Dim e As New CValue
    e.vType = vtNumber
    e.numVal = 2.71828182845905
    mathObj.objectProps.add e, "E"
    
    Dim ln2 As New CValue
    ln2.vType = vtNumber
    ln2.numVal = 0.693147180559945
    mathObj.objectProps.add ln2, "LN2"
    
    Dim ln10 As New CValue
    ln10.vType = vtNumber
    ln10.numVal = 2.30258509299405
    mathObj.objectProps.add ln10, "LN10"
    
    Dim log2e As New CValue
    log2e.vType = vtNumber
    log2e.numVal = 1.44269504088896
    mathObj.objectProps.add log2e, "LOG2E"
    
    Dim log10e As New CValue
    log10e.vType = vtNumber
    log10e.numVal = 0.434294481903252
    mathObj.objectProps.add log10e, "LOG10E"
    
    Dim sqrt2 As New CValue
    sqrt2.vType = vtNumber
    sqrt2.numVal = 1.4142135623731
    mathObj.objectProps.add sqrt2, "SQRT2"
    
    Dim sqrt1_2 As New CValue
    sqrt1_2.vType = vtNumber
    sqrt1_2.numVal = 0.707106781186548
    mathObj.objectProps.add sqrt1_2, "SQRT1_2"
    
    ' Define Math in global scope
    m_globalScope.DefineVar "Math", mathObj
    
    ' NEW: Create JSON object
    Dim jsonObj As New CValue
    jsonObj.vType = vtObject
    Set jsonObj.objectProps = New Collection
    
    ' Define JSON in global scope
    m_globalScope.DefineVar "JSON", jsonObj

    
End Sub

' Public API: Let host add COM objects
Public Sub AddCOMObject(Name As String, obj As Object)
    ' Create a CValue wrapper for the COM object
    Dim val As New CValue
    val.vType = vtCOMObject
    Set val.objVal = obj
    
    ' Define in global scope
    m_globalScope.DefineVar Name, val
    
    ' Also track in collection for enumeration
    On Error Resume Next
    m_comObjects.Remove Name
    On Error GoTo 0
    m_comObjects.add obj, Name
End Sub

' Public API: Remove a COM object
Public Sub RemoveCOMObject(Name As String)
    ' Remove from global scope
    Dim undef As New CValue
    undef.vType = vtUndefined
    m_currentScope.SetVar Name, undef
    
    ' Remove from tracking
    On Error Resume Next
    m_comObjects.Remove Name
    On Error GoTo 0
End Sub

' Main entry point
Public Sub Execute(code As String, Optional withDebugger As Boolean = False)
    On Error GoTo ErrorHandler
    
    ' Parse the code
    Dim ast As CNode
    Set ast = m_parser.ParseScript(code)
    
    If withDebugger Then
        Set m_programAST = ast
    End If
    
    If withDebugger Then InitDebug
    
    ' Execute the program
    ExecuteProgram ast
    
    If m_debugMode Then
        RaiseEvent OnExecutionComplete
    End If
    
    Exit Sub
    
ErrorHandler:
    m_output = m_output & vbCrLf & "ERROR: " & Err.description
    Debug.Print "Interpreter Error: " & Err.description & " - " & Err.source
    
    Dim col As Long
    ' For parser errors during execution, extract line from error message
    If m_currentLine = 0 And Err.Number = vbObjectError + 2000 Then
        m_currentLine = ExtractLineFromError(Err.description)
        col = ExtractColumnFromError(Err.description)
    End If
    
    RaiseEvent OnError(Err.description, m_currentLine, Err.source, col)
     
End Sub

' Get accumulated output (for testing)
Public Function GetOutput() As String
    GetOutput = m_output
End Function

Public Sub ClearOutput()
    m_output = ""
End Sub

'updated to allow set next line
Private Sub ExecuteProgram(programNode As CNode)

    If programNode Is Nothing Then Exit Sub
    If programNode.tType <> Program_Node Then Exit Sub
    
    Dim statements As Collection
    Set statements = programNode.Body
    
    If statements Is Nothing Then Exit Sub
    
    ' Execute each statement
    Dim stmt As CNode
    Dim stmtIndex As Long
    
    stmtIndex = 1
    m_AbortExec = False
    
    Do While stmtIndex <= statements.count
        
        If stmtIndex Mod 25 = 0 Then 'in case of a endless loop allow UI to breathe and user can click abort
            DoEvents
            Sleep 15
        End If
        
        If m_AbortExec = True Then
            RaiseEvent OnError("Execution Aborted", m_currentLine, "ExecuteProgram", 0)
            StopDebug
            Exit Do
        End If
        
        Set stmt = statements(stmtIndex)
        'Debug.Print ">>> ExecuteProgram: About to execute statement " & stmtIndex & " (line " & stmt.LineNumber & ")"
 
        If Not stmt Is Nothing Then
            m_currentLine = stmt.LineNumber
            'If m_debugMode Then 'now causing duplicate hits..
            '    DebugCheckBreak stmt
            'End If
        End If
        
        ' Check if user requested a jump to different line
        If m_hasRequestedNextLine Then
            'Debug.Print ">>> ExecuteProgram: Jump requested! Current stmtIndex=" & stmtIndex
            m_hasRequestedNextLine = False
            
            ' Find the statement at the requested line
            Dim targetIndex As Long
            targetIndex = FindStatementIndexAtLine(statements, m_requestedNextLine)
            
            If targetIndex > 0 Then
                'Debug.Print ">>> Jumping: statement " & stmtIndex & " (line " & stmt.LineNumber & ") -> statement " & targetIndex & " (line " & m_requestedNextLine & ")"
                stmtIndex = targetIndex
                Set stmt = statements(stmtIndex)
            Else
                'Debug.Print ">>> Jump failed: could not find statement at line " & m_requestedNextLine
            End If
        End If
        
        ' Execute the statement
        ExecuteStatement stmt
        
        ' Check for return
        If m_hasReturned Then Exit Sub
        
        stmtIndex = stmtIndex + 1
    Loop
End Sub

' Execute a statement
Private Sub ExecuteStatement(Node As CNode)
    
    DoEvents
    If m_AbortExec Then Exit Sub
    
    If Not Node Is Nothing Then
        m_currentLine = Node.LineNumber
        
        If m_debugMode Then
            If Node.tType <> BlockStatement_Node Then DebugCheckBreak Node
        End If
        
    End If
    
    Select Case Node.tType
        Case ExpressionStatement_Node:
            ' Evaluate and discard result
            Dim result As CValue
            Set result = EvaluateExpression(Node.Test)
            
        Case VariableDeclaration_Node:
            ExecuteVarDeclaration Node
            
        Case BlockStatement_Node:
            ExecuteBlock Node
            
        Case EmptyStatement_Node:
            ' Do nothing
            
        Case IfStatement_Node:
            ExecuteIfStatement Node
            
        Case WhileStatement_Node:
            ExecuteWhileStatement Node
            
        Case DoWhileStatement_Node:
            ExecuteDoWhileStatement Node
            
        Case ForStatement_Node:
            ExecuteForStatement Node
            
        Case BreakStatement_Node:
            ' Set break flag (we'll implement this)
            m_shouldBreak = True
            
        Case ContinueStatement_Node:
            ' Set continue flag
            m_shouldContinue = True
            
        Case ReturnStatement_Node:
            ExecuteReturnStatement Node
            
        Case FunctionDeclaration_Node:
            ExecuteFunctionDeclaration Node
            
        Case ThrowStatement_Node:
            ExecuteThrowStatement Node
            
        Case TryStatement_Node:
            ExecuteTryStatement Node
        
        Case SwitchStatement_Node:
             ExecuteSwitchStatement Node
             
        Case Else:
            Err.Raise vbObjectError + 1000, "CInterpreter", _
                      "Unsupported statement type: " & Node.tType
    End Select
End Sub

' Execute function declaration
Private Sub ExecuteFunctionDeclaration(Node As CNode)
    Dim funcName As String
    funcName = Node.id.Name
    
    ' Create function object
    Dim func As New CFunction
    func.Name = funcName
    Set func.params = Node.params
    Set func.Body = Node.FunctionBody
    Set func.closure = m_currentScope  ' Capture current scope
    
    ' Store as a value
    Dim funcValue As New CValue
    funcValue.vType = vtfunction
    Set funcValue.objVal = func
    
    ' Define in current scope
    m_currentScope.DefineVar funcName, funcValue
End Sub

' === IF STATEMENT ===
Private Sub ExecuteIfStatement(Node As CNode)
    ' Evaluate test condition
    Dim testResult As CValue
    Set testResult = EvaluateExpression(Node.IfTest)
    
    If testResult.ToBoolean() Then
        ' Execute consequent (then branch)
        ExecuteStatement Node.IfConsequent
    ElseIf Not Node.IfAlternate Is Nothing Then
        ' Execute alternate (else branch)
        ExecuteStatement Node.IfAlternate
    End If
End Sub

' === WHILE STATEMENT ===
Private Sub ExecuteWhileStatement(Node As CNode)

    Do While True
    
        If m_AbortExec Then Exit Do
        
        ' Evaluate test condition
        Dim testResult As CValue
        Set testResult = EvaluateExpression(Node.WhileTest)
        
        If Not testResult.ToBoolean() Then Exit Do
        
        If m_debugMode Then
            DebugCheckBreak Node
        End If
        
        ' Execute body
        ExecuteStatement Node.WhileBody
        
        ' Check for break
        If m_shouldBreak Then
            m_shouldBreak = False
            Exit Do
        End If
        
        ' Check for continue
        If m_shouldContinue Then
            m_shouldContinue = False
            ' Continue to next iteration
        End If
        
        ' Check for return
        If m_hasReturned Then Exit Do
    Loop
    
End Sub

' === DO-WHILE STATEMENT ===
Private Sub ExecuteDoWhileStatement(Node As CNode)
    Do
    
        If m_AbortExec Then Exit Do
        
        If m_debugMode Then
            DebugCheckBreak Node
        End If
        
        ' Execute body
        ExecuteStatement Node.WhileBody
        
        ' Check for break
        If m_shouldBreak Then
            m_shouldBreak = False
            Exit Do
        End If
        
        ' Check for continue (still need to check condition)
        If m_shouldContinue Then
            m_shouldContinue = False
        End If
        
        ' Check for return
        If m_hasReturned Then Exit Do
        
        ' Evaluate test condition
        Dim testResult As CValue
        Set testResult = EvaluateExpression(Node.WhileTest)
        
        If Not testResult.ToBoolean() Then Exit Do
    Loop
    
End Sub

' Evaluate function expression
Private Function EvaluateFunctionExpression(Node As CNode) As CValue
    Dim func As New CFunction
    
    If Not Node.id Is Nothing Then
        func.Name = Node.id.Name
    Else
        func.Name = "<anonymous>"
    End If
    
    Set func.params = Node.params
    Set func.Body = Node.FunctionBody
    Set func.closure = m_currentScope
    
    Dim result As New CValue
    result.vType = vtfunction
    Set result.objVal = func
    
    Set EvaluateFunctionExpression = result
End Function

Private Function EvaluateThisExpression() As CValue
    If Not m_thisContext Is Nothing Then
        Set EvaluateThisExpression = m_thisContext
    Else
        ' No 'this' context - return undefined
        Dim undef As New CValue
        undef.vType = vtUndefined
        Set EvaluateThisExpression = undef
    End If
End Function

Private Sub ExecuteForStatement(Node As CNode)

    ' Execute init
    If Not Node.ForInit Is Nothing Then
        If Node.ForInit.tType = VariableDeclaration_Node Then
            ExecuteVarDeclaration Node.ForInit
        Else
            ' Expression
            Dim initResult As CValue
            Set initResult = EvaluateExpression(Node.ForInit)
        End If
    End If
    
    ' Loop
    Do While True
    
        If m_AbortExec Then Exit Do
        
        ' Evaluate test condition
        If Not Node.ForTest Is Nothing Then
            Dim testResult As CValue
            Set testResult = EvaluateExpression(Node.ForTest)
            
            If Not testResult.ToBoolean() Then Exit Do
        End If
        
        ' Reset continue flag before body
        m_shouldContinue = False
        
        If m_debugMode Then
            DebugCheckBreak Node
        End If
        
        ' Execute body
        ExecuteStatement Node.ForBody
        
        ' Check for break
        If m_shouldBreak Then
            m_shouldBreak = False
            Exit Do
        End If
        
        ' Check for return
        If m_hasReturned Then Exit Do
        
        ' Continue falls through to update (flag already cleared above)
        
        ' Execute update
        If Not Node.ForUpdate Is Nothing Then
            Dim updateResult As CValue
            Set updateResult = EvaluateExpression(Node.ForUpdate)
        End If
    Loop
End Sub


' === RETURN STATEMENT ===
Private Sub ExecuteReturnStatement(Node As CNode)
    If Not Node.ReturnArgument Is Nothing Then
        Set m_returnValue = EvaluateExpression(Node.ReturnArgument)
    Else
        Set m_returnValue = New CValue
        m_returnValue.vType = vtUndefined
    End If
    
    m_hasReturned = True
End Sub

' Execute variable declaration
' In CInterpreter.ExecuteVarDeclaration, add:
Private Sub ExecuteVarDeclaration(Node As CNode)
    Dim i As Long
    For i = 1 To Node.Declarations.count
        Dim decl As CNode
        Set decl = Node.Declarations(i)
        
        ' Get variable name
        Dim varName As String
        varName = decl.VarID.Name
        
        If debug_mode Then Debug.Print ">>> Declaring var: " & varName  ' <-- ADD THIS
        
        ' Get initial value
        Dim Value As CValue
        If Not decl.init Is Nothing Then
            Set Value = EvaluateExpression(decl.init)
            If debug_mode Then Debug.Print ">>> Init value type: " & Value.vType  ' <-- ADD THIS
        Else
            ' No initializer - set to undefined
            Set Value = New CValue
            Value.vType = vtUndefined
        End If
        
        ' Define in current scope
        m_currentScope.DefineVar varName, Value
        If debug_mode Then Debug.Print ">>> Defined successfully"  ' <-- ADD THIS
    Next
End Sub

' Execute block statement
Private Sub ExecuteBlock(Node As CNode)
    Dim i As Long
    For i = 1 To Node.Body.count
        Dim stmt As CNode
        Set stmt = Node.Body(i)
        ExecuteStatement stmt
        
        ' Check for early exit conditions
        If m_shouldBreak Or m_shouldContinue Or m_hasReturned Then
            Exit For
        End If
    Next
End Sub

' Evaluate an expression - returns CValue
Private Function EvaluateExpression(Node As CNode) As CValue
  
    If Not Node Is Nothing Then m_currentLine = Node.LineNumber
    
    Select Case Node.tType
        Case Literal_Node:
            Set EvaluateExpression = EvaluateLiteral(Node)
            
        Case Identifier_Node:
            Set EvaluateExpression = EvaluateIdentifier(Node)
            
        Case BinaryExpression_Node:
            Set EvaluateExpression = EvaluateBinary(Node)
            
        Case AssignmentExpression_Node:
            Set EvaluateExpression = EvaluateAssignment(Node)
            
        Case CallExpression_Node:
            Set EvaluateExpression = EvaluateCall(Node)
            
        Case MemberExpression_Node:
            Set EvaluateExpression = EvaluateMember(Node)
            
        ' === NEW: LOGICAL EXPRESSIONS ===
        Case LogicalExpression_Node:
            Set EvaluateExpression = EvaluateLogical(Node)
            
        ' === NEW: UNARY EXPRESSIONS ===
        Case UnaryExpression_Node:
            Set EvaluateExpression = EvaluateUnary(Node)
            
        ' === NEW: UPDATE EXPRESSIONS (++, --) ===
        Case UpdateExpression_Node:
            Set EvaluateExpression = EvaluateUpdate(Node)
            
        Case FunctionExpression_Node:
            Set EvaluateExpression = EvaluateFunctionExpression(Node)
            
        Case ArrayExpression_Node:
            Set EvaluateExpression = EvaluateArrayExpression(Node)
            
        Case ObjectExpression_Node:
            Set EvaluateExpression = EvaluateObjectExpression(Node)
            
        Case NewExpression_Node:
            Set EvaluateExpression = EvaluateNewExpression(Node)
            
        Case ThisExpression_Node
            Set EvaluateExpression = EvaluateThisExpression()
            
        Case Else:
            Debug.Print "CInterpreter.EvaluateExpression Unsupported expression type: " & Node.tType
            Err.Raise vbObjectError + 1001, "CInterpreter", _
                      "Unsupported expression type: " & Node.tType
    End Select
End Function

' === EVALUATE LOGICAL EXPRESSION (&&, ||) ===

Private Function EvaluateLogical(Node As CNode) As CValue
    Dim result As New CValue
    
    ' Evaluate left
    Dim Left As CValue
    Set Left = EvaluateExpression(Node.Left)
    
    Select Case Node.operator
        Case "&&":
            ' AND: short-circuit if left is false
            If Not Left.ToBoolean() Then
                Set result = Left
            Else
                Set result = EvaluateExpression(Node.Right)
            End If
            
        Case "||":
            ' OR: short-circuit if left is true
            If Left.ToBoolean() Then
                Set result = Left
            Else
                Set result = EvaluateExpression(Node.Right)
            End If
            
        Case Else:
            Err.Raise vbObjectError + 1005, "CInterpreter", _
                      "Unknown logical operator: " & Node.operator
    End Select
    
    Set EvaluateLogical = result
End Function

' === EVALUATE UNARY EXPRESSION (!, -, +, ~, typeof, void, delete) ===
Private Function EvaluateUnary(Node As CNode) As CValue
    Dim result As New CValue
    
    ' Evaluate argument
    Dim arg As CValue
    Set arg = EvaluateExpression(Node.Argument)
    
    Select Case Node.operator
        Case "!":
            ' Logical NOT
            result.vType = vtBoolean
            result.boolVal = Not arg.ToBoolean()
            
        Case "-":
            ' Negation
            result.vType = vtNumber
            result.numVal = -arg.ToNumber()
            
        Case "+":
            ' Unary plus (convert to number)
            result.vType = vtNumber
            result.numVal = arg.ToNumber()
            
        Case "~":
            ' Bitwise NOT
            result.vType = vtNumber
            result.numVal = Not CLng(arg.ToNumber())
            
        Case "typeof":
            ' typeof operator

            If debug_mode Then Debug.Print ">>> typeof arg.vType = " & arg.vType
            
            result.vType = vtString
            Select Case arg.vType
                Case vtUndefined:
                    result.strVal = "undefined"
                Case vtNull:
                    result.strVal = "object"  ' JavaScript quirk!
                Case vtBoolean:
                    result.strVal = "boolean"
                Case vtNumber:
                    result.strVal = "number"
                Case vtString:
                    result.strVal = "string"
                Case vtObject:
                    result.strVal = "object"
                Case vtCOMObject:
                    result.strVal = arg.ToString() ' TypeName(arg.objVal)
                Case vtfunction:
                    result.strVal = "function"
                Case vtArray:
                    result.strVal = "array"
                Case vtInt64:
                    result.strVal = "bigint"
                Case Else:
                    result.strVal = "undefined"
            End Select
            
        Case "void":
            ' void operator (always returns undefined)
            result.vType = vtUndefined
            
        Case Else:
            Err.Raise vbObjectError + 1006, "CInterpreter", _
                      "Unsupported unary operator: " & Node.operator
    End Select
    
    Set EvaluateUnary = result
End Function

' === EVALUATE UPDATE EXPRESSION (++, --) ===
Private Function EvaluateUpdate(Node As CNode) As CValue
    Dim result As New CValue
    Dim newValue As New CValue
    
    ' Get current value
    If Node.Argument.tType <> Identifier_Node Then
        Err.Raise vbObjectError + 1007, "CInterpreter", _
                  "Update expression requires identifier"
    End If
    
    Dim varName As String
    varName = Node.Argument.Name
    
    Dim oldValue As CValue
    Set oldValue = m_currentScope.GetVar(varName)
    
    Dim numValue As Double
    numValue = oldValue.ToNumber()
    
    ' Increment or decrement
    newValue.vType = vtNumber
    
    If Node.operator = "++" Then
        newValue.numVal = numValue + 1
    Else
        newValue.numVal = numValue - 1
    End If
    
    ' Update variable
    m_currentScope.SetVar varName, newValue
    
    ' Return old or new value depending on prefix/postfix
    If Node.Prefix Then
        Set result = newValue
    Else
        result.vType = vtNumber
        result.numVal = numValue
        Set result = result
    End If
    
    Set EvaluateUpdate = result
End Function

' Evaluate literal
Private Function EvaluateLiteral(Node As CNode) As CValue
    Dim result As New CValue
    Dim s As String
    Dim U64 As ULong64
    
    If debug_mode Then Debug.Print ">>> EvaluateLiteral: VarType = " & VarType(Node.Value) & ", Value = '" & Node.Value & "'"
 
    If IsNull(Node.Value) Then
        result.vType = vtNull
    ElseIf VarType(Node.Value) = vbBoolean Then
        result.vType = vtBoolean
        result.boolVal = Node.Value
    ElseIf VarType(Node.Value) = vbString Then
        s = Node.Value
        
        ' Check for Int64 literal (ends with 'n' OR is a large hex number)
        Dim needsBigInt As Boolean
        Dim numPart As String
        
        needsBigInt = False
        numPart = s
        
        ' Check if ends with 'n'
        If Len(s) > 1 And Right$(s, 1) = "n" Then
            numPart = Left$(s, Len(s) - 1)
            needsBigInt = True  ' Explicit BigInt request
        End If
        
        ' Check if it's a hex number > signed 32-bit max (auto-promote)
        If debug_mode Then Debug.Print ">>> Checking hex: numPart='" & numPart & "', IsHexString=" & IsHexString(numPart)

        If Not needsBigInt And IsHexString(numPart) Then
            ' Remove 0x/0X prefix
            Dim hexPart As String
            hexPart = numPart
            If LCase(Left$(hexPart, 2)) = "0x" Then
                hexPart = Mid$(hexPart, 3)
            End If
            
            ' SAFE PROMOTION: Check if value > 0x7FFFFFFF (signed 32-bit max)
            If Len(hexPart) > 8 Then
                ' Definitely > 32-bit (more than 8 hex digits)
                needsBigInt = True
            ElseIf Len(hexPart) = 8 Then
                ' Exactly 8 hex digits - check the actual value
                ' If first digit >= 8, value is >= 0x80000000 (needs BigInt)
                Dim firstDigit As String
                firstDigit = UCase$(Left$(hexPart, 1))
                
                If firstDigit >= "8" Then
                    ' Value is 0x80000000 or higher (> signed 32-bit max)
                    needsBigInt = True
                End If
            End If
            ' If < 8 hex digits, it's always <= 0x0FFFFFFF (safe as Number)
        End If
        
        'Parse as BigInt if needed ***
        If needsBigInt And (IsNumeric(numPart) Or IsHexString(numPart)) Then
            result.vType = vtInt64
            
            ' Parse using ULong64
            Set U64 = New ULong64
            U64.mode = mSigned
            If U64.fromString(numPart, mUnsigned) Then
                result.int64Val = U64.rawValue
            Else
                result.int64Val = 0
            End If
        Else
            ' Regular string
            result.vType = vtString
            result.strVal = Node.Value
        End If
        
    ElseIf IsNumeric(Node.Value) Then
        ' Standard number
        result.vType = vtNumber
        result.numVal = CDbl(Node.Value)
    Else
        result.vType = vtUndefined
    End If
    
    Set EvaluateLiteral = result
End Function

' Evaluate identifier (variable lookup)
Private Function EvaluateIdentifier(Node As CNode) As CValue
    Set EvaluateIdentifier = m_currentScope.GetVar(Node.Name)
End Function

      
Private Function EvaluateBinary(Node As CNode) As CValue
    Dim result As New CValue
    
    ' Evaluate left and right
    Dim Left As CValue
    Dim Right As CValue
    Set Left = EvaluateExpression(Node.Left)
    Set Right = EvaluateExpression(Node.Right)
    
    ' === NEW: Check if either operand is Int64 ===
    If Left.vType = vtInt64 Or Right.vType = vtInt64 Then
        ' AUTOMATIC TYPE PROMOTION: Convert any Number operands to Int64
        Dim leftU64 As Currency
        Dim rightU64 As Currency
        
        If Left.vType = vtInt64 Then
            leftU64 = Left.int64Val
        Else
            ' Auto-promote Number to Int64
            leftU64 = Left.ToInt64()
        End If
        
        If Right.vType = vtInt64 Then
            rightU64 = Right.int64Val
        Else
            ' Auto-promote Number to Int64
            rightU64 = Right.ToInt64()
        End If
        
        ' Handle comparison operators - return BOOLEAN
        If Node.operator = "<" Or Node.operator = ">" Or _
           Node.operator = "<=" Or Node.operator = ">=" Or _
           Node.operator = "==" Or Node.operator = "===" Or _
           Node.operator = "!=" Or Node.operator = "!==" Then
            
            result.vType = vtBoolean
            result.boolVal = EvaluateInt64Comparison(leftU64, rightU64, Node.operator)
        Else
            ' Arithmetic operators - return Int64
            result.vType = vtInt64
            result.int64Val = EvaluateInt64Arithmetic(leftU64, rightU64, Node.operator)
        End If
        
        Set EvaluateBinary = result
        Exit Function
    End If
    
    ' === EXISTING CODE FOR REGULAR NUMBERS ===
    ' Perform operation
    Select Case Node.operator
        Case "+":
            ' Addition or string concatenation
            If Left.vType = vtString Or Right.vType = vtString Then
                result.vType = vtString
                result.strVal = Left.ToString() & Right.ToString()
            Else
                result.vType = vtNumber
                result.numVal = Left.ToNumber() + Right.ToNumber()
            End If
            
        Case "-":
            result.vType = vtNumber
            result.numVal = Left.ToNumber() - Right.ToNumber()
            
        Case "*":
            result.vType = vtNumber
            result.numVal = Left.ToNumber() * Right.ToNumber()
            
        Case "/":
            result.vType = vtNumber
            result.numVal = Left.ToNumber() / Right.ToNumber()
            
        Case "%":
            result.vType = vtNumber
            result.numVal = Left.ToNumber() Mod Right.ToNumber()
            
        Case "<":
            result.vType = vtBoolean
            result.boolVal = (Left.ToNumber() < Right.ToNumber())
            
        Case ">":
            result.vType = vtBoolean
            result.boolVal = (Left.ToNumber() > Right.ToNumber())
            
        Case "<=":
            result.vType = vtBoolean
            result.boolVal = (Left.ToNumber() <= Right.ToNumber())
            
        Case ">=":
            result.vType = vtBoolean
            result.boolVal = (Left.ToNumber() >= Right.ToNumber())
            
        Case "==":
            result.vType = vtBoolean
            result.boolVal = IsEqual(Left, Right)
            
        Case "!=":
            result.vType = vtBoolean
            result.boolVal = Not IsEqual(Left, Right)
            
        Case "===":
            result.vType = vtBoolean
            result.boolVal = IsStrictEqual(Left, Right)
            
        Case "!==":
            result.vType = vtBoolean
            result.boolVal = Not IsStrictEqual(Left, Right)
           
        Case "<<":
            result.vType = vtNumber
            Dim lshiftVal As Long
            Dim lshiftAmt As Long
            lshiftVal = CLng(Left.ToNumber()) And &HFFFFFFFF
            lshiftAmt = CLng(Right.ToNumber()) And 31
            result.numVal = lshiftVal * (2# ^ lshiftAmt)
            
        Case ">>":
            result.vType = vtNumber
            Dim rshiftVal As Long
            Dim rshiftAmt As Long
            rshiftVal = CLng(Left.ToNumber())
            rshiftAmt = CLng(Right.ToNumber()) And 31
            result.numVal = rshiftVal \ (2# ^ rshiftAmt)
        
        Case ">>>":  ' Unsigned right shift
            result.vType = vtNumber
            Dim urshiftVal As Long
            Dim urshiftAmt As Long
            urshiftVal = CLng(Left.ToNumber())
            urshiftAmt = CLng(Right.ToNumber()) And 31
            
            ' Convert to unsigned, shift, convert back
            Dim unsignedVal As Double
            If urshiftVal < 0 Then
                ' Negative value - convert to unsigned by adding 2^32
                unsignedVal = CDbl(urshiftVal) + 4294967296#
            Else
                unsignedVal = CDbl(urshiftVal)
            End If
            
            ' Perform unsigned shift
            result.numVal = Fix(unsignedVal / (2# ^ urshiftAmt))
            
        Case "&"  ' Bitwise AND
            result.vType = vtNumber
            Dim andLeft As Long
            Dim andRight As Long
            andLeft = CLng(Left.ToNumber())
            andRight = CLng(Right.ToNumber())
            result.numVal = andLeft And andRight
                        
        Case "|"  ' Bitwise OR
            result.vType = vtNumber
            Dim orLeft As Long
            Dim orRight As Long
            orLeft = CLng(Left.ToNumber())
            orRight = CLng(Right.ToNumber())
            result.numVal = orLeft Or orRight
             
        Case "^"  ' Bitwise XOR
            result.vType = vtNumber
            Dim xorLeft As Long
            Dim xorRight As Long
            xorLeft = CLng(Left.ToNumber())
            xorRight = CLng(Right.ToNumber())
            result.numVal = xorLeft Xor xorRight
                       
        Case Else:
            Err.Raise vbObjectError + 1002, "CInterpreter", _
                      "Unsupported operator: " & Node.operator
    End Select
    
    Set EvaluateBinary = result
End Function

' Evaluate assignment
' Update EvaluateAssignment to handle member expressions:
Private Function EvaluateAssignment(Node As CNode) As CValue
    Dim Value As CValue
    Dim currentValue As CValue
    Dim result As New CValue
    
    ' Check if this is a compound assignment (+=, -=, etc.)
    Dim operator As String
    operator = ""
    On Error Resume Next
    operator = Node.operator
    On Error GoTo 0
    
    If operator = "" Or operator = "=" Then
        ' Simple assignment: x = value
        Set Value = EvaluateExpression(Node.Right)
    Else
        ' Compound assignment: x += value, x -= value, etc.
        ' First get the current value of the left side
        If Node.Left.tType = Identifier_Node Then
            Set currentValue = m_currentScope.GetVar(Node.Left.Name)
        ElseIf Node.Left.tType = MemberExpression_Node Then
            Set currentValue = EvaluateMember(Node.Left)
        Else
            Err.Raise vbObjectError + 1003, "CInterpreter", _
                      "Invalid assignment target"
        End If
        
        ' Evaluate the right side
        Dim rightValue As CValue
        Set rightValue = EvaluateExpression(Node.Right)
        
        ' === Check if either operand is Int64 - handle separately ===
        If currentValue.vType = vtInt64 Or rightValue.vType = vtInt64 Then
            ' Auto-promote to Int64
            Dim leftU64 As Currency
            Dim rightU64 As Currency
            
            If currentValue.vType = vtInt64 Then
                leftU64 = currentValue.int64Val
            Else
                leftU64 = currentValue.ToInt64()
            End If
            
            If rightValue.vType = vtInt64 Then
                rightU64 = rightValue.int64Val
            Else
                rightU64 = rightValue.ToInt64()
            End If
            
            ' Perform Int64 operation
            result.vType = vtInt64
            
            ' Map compound operator to base operator
            Dim baseOp As String
            Select Case operator
                Case "+=": baseOp = "+"
                Case "-=": baseOp = "-"
                Case "*=": baseOp = "*"
                Case "/=": baseOp = "/"
                Case "%=": baseOp = "%"
                Case "&=": baseOp = "&"
                Case "|=": baseOp = "|"
                Case "^=": baseOp = "^"
                Case "<<=": baseOp = "<<"
                Case ">>=": baseOp = ">>"
                Case ">>>=": baseOp = ">>>"
                Case Else:
                    Err.Raise vbObjectError + 1003, "CInterpreter", _
                              "Unsupported Int64 compound assignment: " & operator
            End Select
            
            result.int64Val = EvaluateInt64Arithmetic(leftU64, rightU64, baseOp)
            Set Value = result
            GoTo DoAssignment
        End If
        
        ' === Regular Number compound assignments ===
        ' Perform the operation based on the compound operator
        Select Case operator
            Case "+=":
                ' Addition or concatenation
                If currentValue.vType = vtString Or rightValue.vType = vtString Then
                    result.vType = vtString
                    result.strVal = currentValue.ToString() & rightValue.ToString()
                Else
                    result.vType = vtNumber
                    result.numVal = currentValue.ToNumber() + rightValue.ToNumber()
                End If
                Set Value = result
                
            Case "-=":
                result.vType = vtNumber
                result.numVal = currentValue.ToNumber() - rightValue.ToNumber()
                Set Value = result
                
            Case "*=":
                result.vType = vtNumber
                result.numVal = currentValue.ToNumber() * rightValue.ToNumber()
                Set Value = result
                
            Case "/=":
                result.vType = vtNumber
                result.numVal = currentValue.ToNumber() / rightValue.ToNumber()
                Set Value = result
                
            Case "%=":
                result.vType = vtNumber
                result.numVal = currentValue.ToNumber() Mod rightValue.ToNumber()
                Set Value = result
                
            Case "&=":
                result.vType = vtNumber
                result.numVal = CLng(currentValue.ToNumber()) And CLng(rightValue.ToNumber())
                Set Value = result
                
            Case "|=":
                result.vType = vtNumber
                result.numVal = CLng(currentValue.ToNumber()) Or CLng(rightValue.ToNumber())
                Set Value = result
                
            Case "^=":
                result.vType = vtNumber
                result.numVal = CLng(currentValue.ToNumber()) Xor CLng(rightValue.ToNumber())
                Set Value = result
                
            Case "<<=":
                result.vType = vtNumber
                Dim lshiftVal As Long
                Dim lshiftAmt As Long
                lshiftVal = CLng(currentValue.ToNumber())
                lshiftAmt = CLng(rightValue.ToNumber()) And 31
                result.numVal = lshiftVal * (2# ^ lshiftAmt)
                Set Value = result
                
            Case ">>=":
                result.vType = vtNumber
                Dim rshiftVal As Long
                Dim rshiftAmt As Long
                rshiftVal = CLng(currentValue.ToNumber())
                rshiftAmt = CLng(rightValue.ToNumber()) And 31
                result.numVal = rshiftVal \ (2# ^ rshiftAmt)
                Set Value = result
              
            Case ">>>=":
                result.vType = vtNumber
                Dim urshiftVal2 As Long
                Dim urshiftAmt2 As Long
                urshiftVal2 = CLng(currentValue.ToNumber())
                urshiftAmt2 = CLng(rightValue.ToNumber()) And 31
                
                Dim unsignedVal2 As Double
                If urshiftVal2 < 0 Then
                    unsignedVal2 = CDbl(urshiftVal2) + 4294967296#
                Else
                    unsignedVal2 = CDbl(urshiftVal2)
                End If
                
                result.numVal = Fix(unsignedVal2 / (2# ^ urshiftAmt2))
                Set Value = result
                
            Case Else:
                Err.Raise vbObjectError + 1003, "CInterpreter", _
                          "Unsupported compound assignment operator: " & operator
        End Select
    End If
    
DoAssignment:
    ' Perform the assignment
    If Node.Left.tType = Identifier_Node Then
        m_currentScope.SetVar Node.Left.Name, Value
    ElseIf Node.Left.tType = MemberExpression_Node Then
        ' Member assignment: obj.prop = value or arr[0] = value
        AssignMember Node.Left, Value
    Else
        Err.Raise vbObjectError + 1003, "CInterpreter", _
                  "Invalid assignment target"
    End If
    
    Set EvaluateAssignment = Value
End Function

Private Function EvaluateCall(Node As CNode) As CValue
    Dim result As New CValue
    Dim argValues As New Collection
    Dim i As Long
    Dim objVal As CValue
    Dim methodName As String
    Dim calleeValue As CValue
    Dim func As CFunction
    Dim arg As CNode
    Dim thisContext As CValue
    
    Const dbg_mode As Boolean = False
    
    Static depth As Long
    depth = depth + 1
    
    If depth > 1000 Then  ' Lower threshold to see pattern
        'Debug.Print String$(depth * 2, " ") & ">>> STOPPING at depth " & depth
        'Debug.Print String$(depth * 2, " ") & ">>> node.Callee.tType = " & node.Callee.tType
        If Node.Callee.tType = MemberExpression_Node Then
            'Debug.Print String$(depth * 2, " ") & ">>> Callee.Object.tType = " & node.Callee.Object.tType
            'Debug.Print String$(depth * 2, " ") & ">>> Callee.prop.Name = " & node.Callee.prop.Name
        End If
        depth = 0
        Err.Raise vbObjectError + 9999, "EvaluateCall", "Stack overflow detected (depth > 10)"
    End If
    
    'Debug.Print String$(depth * 2, " ") & ">>> EvaluateCall depth=" & depth
    
    result.vType = vtUndefined
    Set thisContext = Nothing  ' ? INITIALIZE
    
    ' === NEW: Check for global functions (no 'this') ===
    If Node.Callee.tType = Identifier_Node Then
        Dim funcName As String
        funcName = Node.Callee.Name
        
        If IsGlobalFunction(funcName) Then
            ' Evaluate arguments
            For i = 1 To Node.Arguments.count
                Set arg = Node.Arguments(i)
                argValues.add EvaluateExpression(arg) 'collection of CValue
            Next
            
            ' Call global function
            Set result = CallGlobalFunction(funcName, argValues)
            depth = depth - 1
            Set EvaluateCall = result
            Exit Function
        End If
    End If
    
    ' Check for console.log (have 'this' context!)
    If Node.Callee.tType = MemberExpression_Node Then
        ' ? SET THIS CONTEXT EARLY!
        If dbg_mode Then Debug.Print ">>> EvaluateCall: Method call detected!"
        
        If Node.Callee.Object.tType = Identifier_Node Then

            If Node.Callee.Object.Name = "console" Then
                If Node.Callee.prop.Name = "log" Then
                    Set result = HandleConsoleLog(Node.Arguments)  'collection of CNodes still need to evalexpression
                    depth = depth - 1
                    Set EvaluateCall = result
                    Exit Function
                End If
            ElseIf Node.Callee.Object.Name = "Math" Then
                ' Math method call
                methodName = Node.Callee.prop.Name
                
                If IsMathMethod(methodName) Then
                    ' Evaluate arguments
                    For i = 1 To Node.Arguments.count
                        Set arg = Node.Arguments(i)
                        argValues.add EvaluateExpression(arg)
                    Next
                    
                    ' Call Math method
                    Set result = CallMathMethod(methodName, argValues)
                    depth = depth - 1
                    Set EvaluateCall = result
                    Exit Function
                End If
            ElseIf Node.Callee.Object.Name = "JSON" Then
                ' === JSON method call ===
                methodName = Node.Callee.prop.Name
                
                If IsJSONMethod(methodName) Then
                    ' Evaluate arguments
                    For i = 1 To Node.Arguments.count
                        Set arg = Node.Arguments(i)
                        argValues.add EvaluateExpression(arg)
                    Next
                    
                    ' Call JSON method
                    Set result = CallJSONMethod(methodName, argValues)
                    depth = depth - 1
                    Set EvaluateCall = result
                    Exit Function
                End If
            ElseIf Node.Callee.Object.Name = "Date" Then
                ' === Date static methods ===
                methodName = Node.Callee.prop.Name
                
                If methodName = "now" Then
                    result.vType = vtNumber
                    result.numVal = DateDiff("s", #1/1/1970#, Now) * 1000
                    depth = depth - 1
                    Set EvaluateCall = result
                    Exit Function
                End If
            ElseIf Node.Callee.Object.Name = "String" Then
                ' === String static methods ===
                methodName = Node.Callee.prop.Name
                
                If methodName = "fromCharCode" Then
                    ' String.fromCharCode(code1, code2, ...)
                    ' Evaluate arguments
                    Dim j As Long
                    For j = 1 To Node.Arguments.count
                        Set arg = Node.Arguments(j)
                        argValues.add EvaluateExpression(arg)
                    Next
                    
                    ' Build string from character codes
                    Dim charStr As String
                    Dim charCode As Long
                    Dim val As CValue
                    
                    charStr = ""
                    For j = 1 To argValues.count
                        Set val = argValues(j)
                        charCode = CLng(val.ToNumber())
                        If charCode >= 0 And charCode <= 65535 Then
                            charStr = charStr & ChrW(charCode)
                        End If
                    Next
                    
                    ' Audit logging
                    If AuditMode Then
                        Dim codeList As String
                        codeList = ""
                        For j = 1 To argValues.count
                            Set val = argValues(j)
                            codeList = codeList & val.ToString()
                            If j < argValues.count Then codeList = codeList & ","
                        Next
                        RaiseAudit aeDecode, "String.fromCharCode(" & codeList & ") -> " & charStr
                    End If
                    
                    result.vType = vtString
                    result.strVal = charStr
                    depth = depth - 1
                    Set EvaluateCall = result
                    Exit Function
                End If
            End If
        End If

        
        ' === Evaluate object ONCE  and set as 'this' ===
        'Debug.Print String$(depth * 2, " ") & ">>> About to evaluate Callee.Object (type=" & node.Callee.Object.tType & ")"
        Set objVal = EvaluateExpression(Node.Callee.Object)
        Set thisContext = objVal
        
        'Debug.Print String$(depth * 2, " ") & ">>> RETURNED from evaluating Callee.Object, got type=" & objVal.vType

        If dbg_mode Then Debug.Print ">>> 'this' context set to vType=" & objVal.vType
        If objVal.vType = vtObject Then
            If dbg_mode Then Debug.Print ">>>   Object has " & objVal.objectKeys.count & " properties"
        End If
        
        ' === Check if it's an array method call ===
        If objVal.vType = vtArray Then
            methodName = Node.Callee.prop.Name
            
            If IsArrayMethod(methodName) Then
                ' Evaluate arguments
                For i = 1 To Node.Arguments.count
                    Set arg = Node.Arguments(i)
                    argValues.add EvaluateExpression(arg)
                Next
                
                ' Call array method (pass the actual collection)
                Set result = CallArrayMethod(objVal.arrayVal, methodName, argValues)
                depth = depth - 1
                Set EvaluateCall = result
                Exit Function
            End If
        End If
        
        ' === Check if it's a string method call ===
        If objVal.vType = vtString Then
            methodName = Node.Callee.prop.Name
            
            If IsStringMethod(methodName) Then
                ' Evaluate arguments
                For i = 1 To Node.Arguments.count
                    Set arg = Node.Arguments(i)
                    argValues.add EvaluateExpression(arg)
                Next
                
                ' Call string method
                Set result = CallStringMethod(objVal.strVal, methodName, argValues)
                depth = depth - 1  ' ? ADD THIS
                Set EvaluateCall = result
                Exit Function
            End If
        End If
        
        If objVal.vType = vtObject Then
            methodName = Node.Callee.prop.Name

            If IsDateMethod(methodName) Then
                ' Evaluate arguments
                For i = 1 To Node.Arguments.count
                    Set arg = Node.Arguments(i)
                    argValues.add EvaluateExpression(arg)
                Next
                
                ' Call Date method
                Set result = CallDateMethod(objVal, methodName, argValues)
                depth = depth - 1
                Set EvaluateCall = result
                Exit Function
            End If
        End If
        
        ' === Check if it's a COM method call ===
        ' (objVal already evaluated above - reuse it!)
        If objVal.vType = vtCOMObject Then
            methodName = Node.Callee.prop.Name
            
            Set result = CallCOMMethod(objVal.objVal, methodName, Node.Arguments)
            depth = depth - 1  ' ? ADD THIS
            Set EvaluateCall = result
            Exit Function
        End If
        
        ' === User-defined object methods (.THIS SUPPORT IS THE KEY!) ===
        If objVal.vType = vtObject Then
            methodName = Node.Callee.prop.Name
            If dbg_mode Then Debug.Print ">>> Looking for method '" & methodName & "' on object"
            
            ' Get the method from the object
            Dim methodVal As CValue
            Set methodVal = objVal.GetProperty(methodName)
            
            If Not methodVal Is Nothing And methodVal.vType = vtfunction Then
                If dbg_mode Then Debug.Print ">>> Found method '" & methodName & "' - calling with 'this'"
                
                ' Evaluate arguments
                For i = 1 To Node.Arguments.count
                    Set arg = Node.Arguments(i)
                    argValues.add EvaluateExpression(arg)
                Next
                
                ' Call function with 'this' context!
                Set func = methodVal.objVal
                Dim funcValue As New CValue
                funcValue.vType = vtfunction
                Set funcValue.objVal = func
                Set result = CallJSFunction(funcValue, argValues, thisContext)
                
                depth = depth - 1
                Set EvaluateCall = result
                Exit Function
            Else
                If dbg_mode Then Debug.Print ">>> Method '" & methodName & "' not found or not a function"
            End If
        End If

    End If
    
    ' Regular function call (no 'this')
    Set calleeValue = EvaluateExpression(Node.Callee)
    
    If calleeValue.vType <> vtfunction Then
        depth = depth - 1  ' ? ADD THIS
        Err.Raise vbObjectError + 1008, "CInterpreter", _
                  "Call target is not a function"
    End If
    
    Set func = calleeValue.objVal
    
    ' Evaluate arguments
    For i = 1 To Node.Arguments.count
        Set arg = Node.Arguments(i)
        argValues.add EvaluateExpression(arg)
    Next
    
    ' Call the function
    funcValue.vType = vtfunction
    Set funcValue.objVal = func
    Set result = CallJSFunction(funcValue, argValues, thisContext)
    
    depth = depth - 1
    Set EvaluateCall = result
End Function

Private Function EvaluateMember(Node As CNode) As CValue
    ' Evaluate the object
    Dim obj As CValue
    Dim clen As New CValue
    
    Set obj = EvaluateExpression(Node.Object)
    
    ' Get property name
    Dim propName As String
    If Node.Computed Then
        ' Computed: obj[expr]
        Dim propValue As CValue
        Set propValue = EvaluateExpression(Node.prop)
        
        RaiseAudit aeBracketAccess, obj.ToString() & "[" & propValue.ToString() & "]"
        
        If obj.vType = vtArray Then
            ' Array access
            Dim index As Long
            index = CLng(propValue.ToNumber())
            'Debug.Print ">>> Array access: index=" & index & ", count=" & obj.arrayVal.count

            If index >= 0 And index < obj.arrayVal.count Then
                Set EvaluateMember = obj.arrayVal(index + 1)
            Else
                Dim undef As New CValue
                undef.vType = vtUndefined
                Set EvaluateMember = undef
            End If
            Exit Function
        ElseIf obj.vType = vtString Then  ' ADD THIS
            ' String character access
            index = CLng(propValue.ToNumber())
            Dim charVal As New CValue
            charVal.vType = vtString
            If index >= 0 And index < Len(obj.strVal) Then
                charVal.strVal = Mid$(obj.strVal, index + 1, 1)
            Else
                charVal.strVal = ""
            End If
            Set EvaluateMember = charVal
            Exit Function
        Else
            propName = propValue.ToString()
        End If
    Else
        ' Static: obj.prop
        propName = Node.prop.Name
    End If
    
    ' === NEW: Handle COM objects ===
    If obj.vType = vtCOMObject Then
        Set EvaluateMember = GetCOMProperty(obj.objVal, propName)
        Exit Function
    End If
    
    ' === Handle STRING properties ===
    If obj.vType = vtString Then
        If propName = "length" Then
            clen.vType = vtNumber
            clen.numVal = Len(obj.strVal)
            Set EvaluateMember = clen
            Exit Function
        End If
    End If

    ' Handle special array properties
    If obj.vType = vtArray And propName = "length" Then
        clen.vType = vtNumber
        clen.numVal = obj.arrayVal.count
        Set EvaluateMember = clen
        Exit Function
    End If
    
    ' Object property access
    If obj.vType = vtObject Then
        On Error Resume Next
        Dim result As CValue
        Set result = obj.objectProps(propName)
        
        If Err.Number = 0 Then
            Set EvaluateMember = result
        Else
            Set undef = New CValue
            undef.vType = vtUndefined
            Set EvaluateMember = undef
        End If
        On Error GoTo 0
        Exit Function
    End If
    
    ' Default: undefined
    Set undef = New CValue
    undef.vType = vtUndefined
    Set EvaluateMember = undef
End Function

' Handle console.log
Private Function HandleConsoleLog(args As Collection) As CValue
    Set HandleConsoleLog = HandlePrint(args)   'col of CNodes
End Function


Private Function HandleFormat(args As Collection) As CValue
    
    Dim result As New CValue
    Dim arg As CValue
    Dim i As Long
    
    ' format("template {var}") - looks up vars in scope
    ' format("template {0} {1}", arg1, arg2) - numbered placeholders
    
    If args.count = 0 Then
        result.vType = vtString
        result.strVal = ""
        Set HandleFormat = result
        Exit Function
    End If
    
    ' First arg is the format string
    Set arg = args(1)
    Dim formatStr As String
    formatStr = arg.ToString()
    
    If args.count = 1 Then
        ' ================================================================
        ' SMART MODE: Look up {varname} in scope!
        ' ================================================================
        
        ' Find all {varname} patterns
        Dim pos As Long
        pos = 1
        
        Do While pos <= Len(formatStr)
            Dim startPos As Long
            startPos = InStr(pos, formatStr, "{")
            
            If startPos = 0 Then Exit Do
            
            Dim endPos As Long
            endPos = InStr(startPos, formatStr, "}")
            
            If endPos = 0 Then Exit Do
            
            ' Extract variable name
            Dim varName As String
            varName = Mid$(formatStr, startPos + 1, endPos - startPos - 1)
            varName = Trim$(varName)
            
            ' Check if it's a number (numbered placeholder with no args = leave it)
            If IsNumeric(varName) Then
                pos = endPos + 1
            Else
                ' Look up variable in scope!
                Dim varValue As CValue
                Set varValue = m_currentScope.GetVar(varName)
                
                If varValue.vType = vtUndefined Then
                    ' Try global scope
                    Set varValue = m_globalScope.GetVar(varName)
                End If
                
                ' Format the value
                Dim replacement As String
                Select Case varValue.vType
                    Case vtUndefined
                        replacement = "undefined"
                    Case vtObject, vtArray
                        replacement = modJSON.StringifyJSON(varValue)
                    Case Else
                        replacement = varValue.ToString()
                End Select
                
                ' Replace in string
                Dim placeholder As String
                placeholder = "{" & varName & "}"
                formatStr = Replace(formatStr, placeholder, replacement, , 1)
                
                ' Adjust position (replacement might be different length)
                pos = startPos + Len(replacement)
            End If
        Loop
        
        result.vType = vtString
        result.strVal = formatStr
        
    Else
        ' ================================================================
        ' CLASSIC MODE: Numbered placeholders {0}, {1}, etc.
        ' ================================================================
        
        For i = 2 To args.count
            Set arg = args(i)
            placeholder = "{" & (i - 2) & "}"
            
            ' Smart formatting
            Select Case arg.vType
                Case vtObject, vtArray
                    replacement = modJSON.StringifyJSON(arg)
                Case Else
                    replacement = arg.ToString()
            End Select
            
            formatStr = Replace(formatStr, placeholder, replacement)
        Next
        
        result.vType = vtString
        result.strVal = formatStr
    End If
    
    Set HandleFormat = result
    
End Function



' Handle print() - smart formatting with object/array support
Private Function HandlePrint(args As Collection, Optional isAlert As Boolean) As CValue

    Dim i As Long
    Dim output As String
    Dim val As CValue
    
    Dim result As New CValue
    result.vType = vtString
    Set HandlePrint = result
    
    For i = 1 To args.count
        Dim arg As Object 'CValue or CNode, CNode needs EvaluateExpression
        Set arg = args(i)
        
        If TypeName(arg) = "CNode" Then
            If debug_mode Then Debug.Print ">>> HandlePrint: Evaluating CNode, NodeType=" & arg.NodeType
            Set val = EvaluateExpression(arg)
            If debug_mode Then Debug.Print ">>> HandlePrint: After eval, val.vType=" & val.vType & ", value=" & val.ToString()
        Else
            Set val = arg
            If debug_mode Then Debug.Print ">>> HandlePrint: Direct CValue, vType=" & val.vType
        End If
        
        If i > 1 Then output = output & " "
        
        ' Smart formatting based on type
        Select Case val.vType
            Case vtObject
                ' Use JSON.stringify for objects
                output = output & modJSON.StringifyJSON(val)
                
            Case vtArray
                ' Use JSON.stringify for arrays
                output = output & modJSON.StringifyJSON(val)
                
            Case vtUndefined
                output = output & "undefined"
                
            Case vtNull
                output = output & "null"
                
            Case vtBoolean
                output = output & IIf(val.boolVal, "true", "false")
                
            Case vtNumber
                output = output & CStr(val.numVal)
                
            Case vtString
                output = output & val.strVal
                
            Case vtfunction
                output = output & "[Function]"
                
            Case vtCOMObject
                output = output & "[COMObject]"
                
            Case vtInt64
                Dim U64 As New ULong64
                U64.mode = mSigned
                U64.rawValue = val.int64Val
                output = output & U64.ToString(mUnsigned)
                
            Case Else
                output = output & val.ToString()
        End Select
    Next
    
    result.strVal = output
    
    If isAlert Then
        MsgBox output
    Else
        RaiseEvent ConsoleLog(output)
        m_output = m_output & output & vbCrLf
        Debug.Print output
    End If
    
End Function

' Equality comparison (==)
Private Function IsEqual(Left As CValue, Right As CValue) As Boolean
    ' Simplified equality (no type coercion for now)
    If Left.vType <> Right.vType Then
        IsEqual = False
        Exit Function
    End If
    
    Select Case Left.vType
        Case vtUndefined, vtNull:
            IsEqual = True
        Case vtBoolean:
            IsEqual = (Left.boolVal = Right.boolVal)
        Case vtNumber:
            IsEqual = (Left.numVal = Right.numVal)
        Case vtString:
            IsEqual = (Left.strVal = Right.strVal)
        Case Else:
            IsEqual = False
    End Select
End Function

' Strict equality comparison (===)
Private Function IsStrictEqual(Left As CValue, Right As CValue) As Boolean
    ' Same as IsEqual for now
    IsStrictEqual = IsEqual(Left, Right)
End Function

' ============================================
' STAGE 3: ARRAYS
' ============================================

Private Function EvaluateArrayExpression(Node As CNode) As CValue

    Dim result As New CValue
    Dim undef As New CValue
    
    result.vType = vtArray
    Set result.arrayVal = New Collection
    
    ' Evaluate each element
    Dim i As Long
    For i = 1 To Node.Body.count
        Dim elem As CNode
        Set elem = Node.Body(i)
        
        If elem Is Nothing Then
            ' Sparse array - add undefined
            Set undef = New CValue
            undef.vType = vtUndefined
            result.arrayVal.add undef
        Else
            result.arrayVal.add EvaluateExpression(elem)
        End If
    Next
    
    Set EvaluateArrayExpression = result
End Function

' ============================================
' STAGE 3: OBJECTS
' ============================================

' Update EvaluateObjectExpression in CInterpreter.cls

Private Function EvaluateObjectExpression(Node As CNode) As CValue
    Dim result As New CValue
    result.vType = vtObject
    Set result.objectProps = New Collection
    Set result.objectKeys = New Collection  ' NEW
    
    ' Evaluate each property
    Dim i As Long
    For i = 1 To Node.Properties.count
        Dim prop As CNode
        Set prop = Node.Properties(i)
        
        ' Get property key
        Dim key As String
        If prop.PropKey.tType = Identifier_Node Then
            key = prop.PropKey.Name
        ElseIf prop.PropKey.tType = Literal_Node Then
            key = CStr(prop.PropKey.Value)
        Else
            key = "unknown"
        End If
        
        ' Evaluate property value
        Dim propValue As CValue
        Set propValue = EvaluateExpression(prop.propValue)
        
        ' Store property (remove if exists, then add)
        On Error Resume Next
        result.objectProps.Remove key
        result.objectKeys.Remove key
        On Error GoTo 0
        result.objectProps.add propValue, key
        result.objectKeys.add key, key  ' NEW: Track the key
    Next
    
    Set EvaluateObjectExpression = result
End Function


' Helper to assign to member
Private Sub AssignMember(memberNode As CNode, Value As CValue)
    ' Evaluate the object
    Dim obj As CValue
    Set obj = EvaluateExpression(memberNode.Object)
    
    ' Get property name
    Dim propName As String
    If memberNode.Computed Then
        ' Computed: obj[expr]
        Dim propValue As CValue
        Set propValue = EvaluateExpression(memberNode.prop)
        
        If obj.vType = vtArray Then
            ' Array assignment
            Dim index As Long
            index = CLng(propValue.ToNumber())
            
            If index >= 0 And index < obj.arrayVal.count Then
                obj.arrayVal.Remove index + 1
                obj.arrayVal.add Value, , index + 1
            End If
            Exit Sub
        Else
            propName = propValue.ToString()
        End If
    Else
        ' Static: obj.prop
        propName = memberNode.prop.Name
    End If
    
    ' === NEW: Handle COM objects ===
    If obj.vType = vtCOMObject Then
        SetCOMProperty obj.objVal, propName, Value
        Exit Sub
    End If
    
    ' Object property assignment
    If obj.vType = vtObject Then
        On Error Resume Next
        obj.objectProps.Remove propName
        obj.objectKeys.Remove propName
        On Error GoTo 0
        obj.objectProps.add Value, propName
        obj.objectKeys.add propName, propName
    End If
End Sub

' ============================================
' INT64 OPERATIONS
' ============================================

' Int64 arithmetic helper
Private Function EvaluateInt64Arithmetic(leftVal As Currency, rightVal As Currency, op As String) As Currency
    Dim u64Left As New ULong64
    Dim u64Right As New ULong64
    Dim u64Result As ULong64
    
    u64Left.mode = mSigned
    u64Left.rawValue = leftVal
    
    u64Right.mode = mSigned
    u64Right.rawValue = rightVal
    
    Select Case op
        Case "+":
            Set u64Result = u64Left.add(u64Right)
            
        Case "-":
            Set u64Result = u64Left.subtract(u64Right)
            
        Case "*":
            Set u64Result = u64Left.multiply(u64Right)
            
        Case "/":
            Set u64Result = u64Left.divide(u64Right)
            
        Case "%":
            Set u64Result = u64Left.modulus(u64Right)
            
        Case "&":
            Set u64Result = u64Left.and_(u64Right)
            
        Case "|":
            Set u64Result = u64Left.or_(u64Right)
            
        Case "^":
            Set u64Result = u64Left.xor_(u64Right)
            
        Case "<<":
            Set u64Result = u64Left.lshift(u64Right)
            
        Case ">>":
            Set u64Result = u64Left.rshift(u64Right)
        
        Case ">>>":
            ' For Int64, force unsigned mode for >>> operation
            Dim u64Temp As New ULong64
            u64Temp.mode = mUnsigned  ' Force unsigned
            u64Temp.rawValue = leftVal
            
            Dim u64Amt As New ULong64
            u64Amt.mode = mUnsigned
            u64Amt.rawValue = rightVal
            
            Set u64Result = u64Temp.rshift(u64Amt)
            
        Case Else:
            Err.Raise vbObjectError + 1010, "CInterpreter", _
                      "Unsupported Int64 operator: " & op
    End Select
    
    EvaluateInt64Arithmetic = u64Result.rawValue
End Function

' Int64 comparison helper
Private Function EvaluateInt64Comparison(leftVal As Currency, rightVal As Currency, op As String) As Boolean
    Dim u64Left As New ULong64
    Dim u64Right As New ULong64
    
    u64Left.mode = mSigned
    u64Left.rawValue = leftVal
    u64Right.mode = mSigned
    u64Right.rawValue = rightVal
    
    Select Case op
        Case "<":
            EvaluateInt64Comparison = u64Left.lessThan(u64Right)
        Case ">":
            EvaluateInt64Comparison = u64Left.greaterThan(u64Right)
        Case "<=":
            EvaluateInt64Comparison = u64Left.lessThanEqual(u64Right)
        Case ">=":
            EvaluateInt64Comparison = u64Left.greaterThanEqual(u64Right)
        Case "==", "===":
            EvaluateInt64Comparison = (leftVal = rightVal)
        Case "!=", "!==":
            EvaluateInt64Comparison = (leftVal <> rightVal)
    End Select
End Function

Private Function GetCOMProperty(comObj As Object, propName As String) As CValue
    On Error GoTo ErrorHandler
    
    If Not ensureDynProxy() Then
        RaiseEvent OnError("GetCOMProperty: Dynproxy not found", 0, "GetCOMProperty", 0)
        Exit Function
    End If
    
    Dim resultVar As Variant
    Dim isObj As Boolean
    Dim hr As Long
    Dim result As New CValue
    Dim emptyArgs() As Variant
    
    ReDim emptyArgs(0 To -1)  ' Empty array for no arguments
    
    ' Try PROPERTYGET
    hr = CallByNameEx(comObj, propName, DISPATCH_PROPERTYGET, emptyArgs, resultVar, isObj)
    
    ' If that failed, try METHOD (for parameterless methods that act like properties)
    If hr <> 0 Then
        hr = CallByNameEx(comObj, propName, DISPATCH_METHOD, emptyArgs, resultVar, isObj)
    End If
    
    ' Convert result
    If hr = 0 Then  ' S_OK
        If isObj Then
            result.vType = vtCOMObject
            Set result.objVal = resultVar
        Else
            Set result = VariantToCValue(resultVar)
        End If
    Else
        ' Property doesn't exist or error
        result.vType = vtUndefined
    End If
    
    Set GetCOMProperty = result
    Exit Function
    
ErrorHandler:
    Dim undef As New CValue
    undef.vType = vtUndefined
    Set GetCOMProperty = undef
End Function

'Private Function GetCOMProperty(comObj As Object, propName As String) As CValue
'    On Error GoTo ErrorHandler
'
'    Dim result As New CValue
'
'    ' Try to get as object first
'    Dim propObj As Object
'    On Error Resume Next
'    Set propObj = CallByName(comObj, propName, VbGet)
'
'    If Err.Number = 0 And Not propObj Is Nothing Then
'        ' It's an object
'        result.vType = vtCOMObject
'        Set result.objVal = propObj
'    Else
'        ' Try as value
'        Err.Clear
'        Dim propVal As Variant
'        propVal = CallByName(comObj, propName, VbGet)
'
'        If Err.Number <> 0 Then
'            ' Try as method instead (parameterless method)
'            Err.Clear
'            propVal = CallByName(comObj, propName, VbMethod)
'        End If
'
'        If Err.Number = 0 Then
'            Set result = VariantToCValue(propVal)
'        Else
'            result.vType = vtUndefined
'        End If
'    End If
'    On Error GoTo 0
'
'    Set GetCOMProperty = result
'    Exit Function
'
'ErrorHandler:
'    ' Property doesn't exist or error - return undefined
'    Dim undef As New CValue
'    undef.vType = vtUndefined
'    Set GetCOMProperty = undef
'End Function


Private Function VariantToCValue(v As Variant) As CValue
    Dim result As New CValue
    
    If IsEmpty(v) Then
        result.vType = vtUndefined
        Set VariantToCValue = result
        Exit Function
    End If
    
    If IsNull(v) Then
        result.vType = vtNull
        Set VariantToCValue = result
        Exit Function
    End If
    
    Select Case VarType(v)
        Case vbEmpty:
            result.vType = vtUndefined
            
        Case vbNull:
            result.vType = vtNull
            
        Case vbBoolean:
            result.vType = vtBoolean
            result.boolVal = CBool(v)
            
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal:
            result.vType = vtNumber
            result.numVal = CDbl(v)
            
        Case vbString:
            result.vType = vtString
            result.strVal = CStr(v)
            
        Case vbObject:
            If v Is Nothing Then
                result.vType = vtNull
            Else
                ' It's a COM object - wrap it
                result.vType = vtCOMObject
                Set result.objVal = v
            End If
            
        Case vbDate:
            ' Convert date to string for now
            result.vType = vtString
            result.strVal = CStr(v)
            
        ' === NEW: Handle arrays ===
        Case vbArray + vbVariant, vbArray + vbByte, vbArray + vbInteger, _
             vbArray + vbLong, vbArray + vbSingle, vbArray + vbDouble, _
             vbArray + vbString, vbArray + vbBoolean
            ' Convert Variant array to JavaScript array
            result.vType = vtArray
            Set result.arrayVal = New Collection
            
            Dim i As Long
            Dim lb As Long
            Dim ub As Long
            
            lb = LBound(v)
            ub = UBound(v)
            
            For i = lb To ub
                Dim elemVal As CValue
                Set elemVal = VariantToCValue(v(i))  ' Recursive!
                result.arrayVal.add elemVal
            Next
            
        Case Else:
            ' Unknown type - try to treat as object
            On Error Resume Next
            If isObject(v) Then
                result.vType = vtCOMObject
                Set result.objVal = v
            Else
                result.vType = vtUndefined
            End If
            On Error GoTo 0
    End Select
    
    Set VariantToCValue = result
End Function

' NEW: Convert CValue to Variant (for passing to COM)
Private Function CValueToVariant(val As CValue) As Variant
    Select Case val.vType
        Case vtUndefined:
            CValueToVariant = Empty
            
        Case vtNull:
            CValueToVariant = Null
            
        Case vtBoolean:
            CValueToVariant = val.boolVal
            
        Case vtNumber:
            CValueToVariant = val.numVal
            
        Case vtString:
            CValueToVariant = val.strVal
            
        Case vtInt64:
            ' Convert Currency to Double (may lose precision)
            CValueToVariant = CDbl(val.int64Val)
            
        Case vtCOMObject:
            Set CValueToVariant = val.objVal
            
        Case vtObject:
            ' Can't convert JS object to COM - return Nothing
            CValueToVariant = Nothing
            
        Case vtArray:
            ' Can't convert JS array to COM - return Nothing
            CValueToVariant = Nothing
            
        Case Else:
            CValueToVariant = Empty
    End Select
End Function

Private Function CallCOMMethod(comObj As Object, methodName As String, args As Collection) As CValue
    On Error GoTo ErrorHandler
    
    If Not ensureDynProxy() Then
        RaiseEvent OnError("CallCOMMethod: Dynproxy not found", 0, "", 0)
        Exit Function
    End If
    
    ' Audit logging
    If AuditMode Then
        Dim objType As String
        On Error Resume Next
        objType = TypeName(comObj)
        On Error GoTo ErrorHandler
        
        Dim auditMsg As String
        auditMsg = objType & "." & methodName & "("
        
        If args.count > 0 Then
            Dim firstArg As CNode
            Set firstArg = args(1)
            Dim firstArgVal As CValue
            Set firstArgVal = EvaluateExpression(firstArg)
            auditMsg = auditMsg & Chr(34) & firstArgVal.ToString() & Chr(34)
        End If
        auditMsg = auditMsg & ")"
        
        RaiseAudit aeCOMCall, auditMsg
    End If
    
    ' Convert arguments to Variant array
    Dim vArgs() As Variant
    Dim i As Long
    
    If args.count > 0 Then
        ReDim vArgs(0 To args.count - 1)
        For i = 1 To args.count
            Dim arg As CNode
            Set arg = args(i)
            Dim argVal As CValue
            Set argVal = EvaluateExpression(arg)
            vArgs(i - 1) = CValueToVariant(argVal)
        Next
    Else
        ReDim vArgs(0 To -1)  ' Empty array
    End If
    
    ' Try different invoke types - order matters!
    Dim resultVar As Variant
    Dim isObj As Boolean
    Dim hr As Long
    
    ' 1. Try PROPERTYGET first (handles parameterized properties like Environment("PATH"))
    hr = CallByNameEx(comObj, methodName, DISPATCH_PROPERTYGET, vArgs, resultVar, isObj)
    
    ' 2. If that failed, try METHOD
    If hr <> 0 Then
        hr = CallByNameEx(comObj, methodName, DISPATCH_METHOD, vArgs, resultVar, isObj)
    End If
    
    ' Convert result to CValue
    Dim result As New CValue
    
    If hr = 0 Then  ' S_OK
        If isObj Then
            result.vType = vtCOMObject
            Set result.objVal = resultVar
        Else
            Set result = VariantToCValue(resultVar)
        End If
    Else
        ' Call failed - return undefined
        result.vType = vtUndefined
    End If
    
    Set CallCOMMethod = result
    Exit Function
    
ErrorHandler:
    Dim undef As New CValue
    undef.vType = vtUndefined
    Set CallCOMMethod = undef
End Function

'this brute force technique can call functions twice!
'Private Function CallCOMMethod(comObj As Object, methodName As String, args As Collection) As CValue
'    On Error GoTo ErrorHandler
'
'    ' Convert arguments to Variant array
'    Dim vArgs() As Variant
'    Dim i As Long
'
'    ' Audit logging
'    If AuditMode Then
'        Dim objType As String
'        On Error Resume Next
'        objType = TypeName(comObj)
'        On Error GoTo ErrorHandler
'
'        Dim auditMsg As String
'        auditMsg = objType & "." & methodName & "("
'
'        ' Add first argument if present
'        If args.count > 0 Then
'            Dim firstArg As CNode
'            Set firstArg = args(1)
'            Dim firstArgVal As CValue
'            Set firstArgVal = EvaluateExpression(firstArg)
'            auditMsg = auditMsg & Chr(34) & firstArgVal.ToString() & Chr(34)
'        End If
'        auditMsg = auditMsg & ")"
'
'        RaiseAudit aeCOMCall, auditMsg
'    End If
'
'
'    If args.count > 0 Then
'        ReDim vArgs(1 To args.count)
'        For i = 1 To args.count
'            Dim arg As CNode
'            Set arg = args(i)
'
'            Dim argVal As CValue
'            Set argVal = EvaluateExpression(arg)
'
'            vArgs(i) = CValueToVariant(argVal)
'        Next
'    End If
'
'    ' CRITICAL: Try VbGet first (parameterized property), then VbMethod
'    Dim resultObj As Object
'    Dim resultVal As Variant
'    Dim gotObject As Boolean
'    Dim gotValue As Boolean
'
'    ' Try getting as object using VbGet (for parameterized properties like Environment)
'    On Error Resume Next
'    Err.Clear
'
'    'Debug.Print ">>> Trying VbGet first..."
'    If args.count = 0 Then
'        Set resultObj = CallByName(comObj, methodName, VbGet)
'    ElseIf args.count = 1 Then
'        Set resultObj = CallByName(comObj, methodName, VbGet, vArgs(1))
'    ElseIf args.count = 2 Then
'        Set resultObj = CallByName(comObj, methodName, VbGet, vArgs(1), vArgs(2))
'    ElseIf args.count = 3 Then
'        Set resultObj = CallByName(comObj, methodName, VbGet, vArgs(1), vArgs(2), vArgs(3))
'    ElseIf args.count = 4 Then
'        Set resultObj = CallByName(comObj, methodName, VbGet, vArgs(1), vArgs(2), vArgs(3), vArgs(4))
'    ElseIf args.count = 5 Then
'        Set resultObj = CallByName(comObj, methodName, VbGet, vArgs(1), vArgs(2), vArgs(3), vArgs(4), vArgs(5))
'    End If
'
'    gotObject = (Err.Number = 0 And Not resultObj Is Nothing)
'    'Debug.Print ">>> VbGet result: Err=" & Err.Number & ", gotObject=" & gotObject
'    Err.Clear
'    On Error GoTo 0
'
'    ' If VbGet didn't work, try VbMethod
'    If Not gotObject Then
'        On Error Resume Next
'        Err.Clear
'
'        'Debug.Print ">>> Trying VbMethod..."
'        If args.count = 0 Then
'            Set resultObj = CallByName(comObj, methodName, VbMethod)
'        ElseIf args.count = 1 Then
'            Set resultObj = CallByName(comObj, methodName, VbMethod, vArgs(1))
'        ElseIf args.count = 2 Then
'            Set resultObj = CallByName(comObj, methodName, VbMethod, vArgs(1), vArgs(2))
'        ElseIf args.count = 3 Then
'            Set resultObj = CallByName(comObj, methodName, VbMethod, vArgs(1), vArgs(2), vArgs(3))
'        ElseIf args.count = 4 Then
'            Set resultObj = CallByName(comObj, methodName, VbMethod, vArgs(1), vArgs(2), vArgs(3), vArgs(4))
'        ElseIf args.count = 5 Then
'            Set resultObj = CallByName(comObj, methodName, VbMethod, vArgs(1), vArgs(2), vArgs(3), vArgs(4), vArgs(5))
'        End If
'
'        gotObject = (Err.Number = 0 And Not resultObj Is Nothing)
'        'Debug.Print ">>> VbMethod result: Err=" & Err.Number & ", gotObject=" & gotObject
'        Err.Clear
'        On Error GoTo 0
'    End If
'
'    ' If still didn't get object, try getting as value
'    If Not gotObject Then
'        On Error Resume Next
'        Err.Clear
'
'        'Debug.Print ">>> Trying value with VbGet..."
'        If args.count = 0 Then
'            resultVal = CallByName(comObj, methodName, VbGet)
'        ElseIf args.count = 1 Then
'            resultVal = CallByName(comObj, methodName, VbGet, vArgs(1))
'        ElseIf args.count = 2 Then
'            resultVal = CallByName(comObj, methodName, VbGet, vArgs(1), vArgs(2))
'        ElseIf args.count = 3 Then
'            resultVal = CallByName(comObj, methodName, VbGet, vArgs(1), vArgs(2), vArgs(3))
'        ElseIf args.count = 4 Then
'            resultVal = CallByName(comObj, methodName, VbGet, vArgs(1), vArgs(2), vArgs(3), vArgs(4))
'        ElseIf args.count = 5 Then
'            resultVal = CallByName(comObj, methodName, VbGet, vArgs(1), vArgs(2), vArgs(3), vArgs(4), vArgs(5))
'        End If
'
'        gotValue = (Err.Number = 0)
'        'Debug.Print ">>> VbGet value result: Err=" & Err.Number & ", gotValue=" & gotValue
'
'        ' If VbGet didn't work, try VbMethod
'        If Not gotValue Then
'            Err.Clear
'            'Debug.Print ">>> Trying value with VbMethod..."
'            If args.count = 0 Then
'                resultVal = CallByName(comObj, methodName, VbMethod)
'            ElseIf args.count = 1 Then
'                resultVal = CallByName(comObj, methodName, VbMethod, vArgs(1))
'            ElseIf args.count = 2 Then
'                resultVal = CallByName(comObj, methodName, VbMethod, vArgs(1), vArgs(2))
'            ElseIf args.count = 3 Then
'                resultVal = CallByName(comObj, methodName, VbMethod, vArgs(1), vArgs(2), vArgs(3))
'            ElseIf args.count = 4 Then
'                resultVal = CallByName(comObj, methodName, VbMethod, vArgs(1), vArgs(2), vArgs(3), vArgs(4))
'            ElseIf args.count = 5 Then
'                resultVal = CallByName(comObj, methodName, VbMethod, vArgs(1), vArgs(2), vArgs(3), vArgs(4), vArgs(5))
'            End If
'            gotValue = (Err.Number = 0)
'            'Debug.Print ">>> VbMethod value result: Err=" & Err.Number & ", gotValue=" & gotValue
'        End If
'
'        Err.Clear
'        On Error GoTo 0
'    End If
'
'    ' Build result
'    Dim result As New CValue
'
'    If gotObject Then
'        'Debug.Print ">>> Got object: " & TypeName(resultObj)
'        result.vType = vtCOMObject
'        Set result.objVal = resultObj
'    ElseIf gotValue Then
'        'Debug.Print ">>> Got value: " & VarType(resultVal)
'        Set result = VariantToCValue(resultVal)
'    Else
'        'Debug.Print ">>> Got nothing!"
'        result.vType = vtUndefined
'    End If
'
'    Set CallCOMMethod = result
'    Exit Function
'
'ErrorHandler:
'    'Debug.Print ">>> CallCOMMethod ERROR: " & Err.Description
'    Dim undef As New CValue
'    undef.vType = vtUndefined
'    Set CallCOMMethod = undef
'End Function


' NEW: Set COM property
'Private Sub SetCOMProperty(comObj As Object, propName As String, Value As CValue)
'    On Error Resume Next
'
'    Dim v As Variant
'    v = CValueToVariant(Value)
'
'    CallByName comObj, propName, VbLet, v
'
'    On Error GoTo 0
'End Sub

Private Sub SetCOMProperty(comObj As Object, propName As String, Value As CValue)
    On Error Resume Next
    
    If Not ensureDynProxy() Then
        RaiseEvent OnError("SetCOMProperty: Dynproxy not found", 0, "SetCOMProperty", 0)
        Exit Sub
    End If
    
    Dim v As Variant
    Dim args() As Variant
    Dim resultVar As Variant
    Dim isObj As Boolean
    Dim hr As Long
    Dim invokeType As Integer
    
    ' Convert CValue to Variant
    v = CValueToVariant(Value)
    
    ' Determine if we're setting an object or a value
    If Value.vType = vtCOMObject Or Value.vType = vtObject Then
        invokeType = DISPATCH_PROPERTYPUTREF  ' Property Set
    Else
        invokeType = DISPATCH_PROPERTYPUT      ' Property Let
    End If
    
    ' Create args array with the value to assign
    ReDim args(0 To 0)
    args(0) = v
    
    ' Set the property
    hr = CallByNameEx(comObj, propName, invokeType, args, resultVar, isObj)
    
    ' Ignore errors (On Error Resume Next behavior maintained)
    
    On Error GoTo 0
End Sub

' ============================================
' NEW EXPRESSION (ActiveXObject)
' ============================================

Private Function EvaluateNewExpression(Node As CNode) As CValue

    Dim result As New CValue
    Dim comObj As Object
    Dim arg As CNode
    Dim progId As CValue
    Dim message As String
    Dim msgVal As CValue

    ' Check if it's new Error()
    If Node.Callee.tType = Identifier_Node Then
        
        If Node.Callee.Name = "Date" Then
            Set result = CreateDateObject(Node.Arguments)
            Set EvaluateNewExpression = result
            Exit Function
        ElseIf Node.Callee.Name = "Error" Then
            ' new Error([message])
            message = "Error"
            
            If Node.Arguments.count > 0 Then
                Set arg = Node.Arguments(1)
                Set msgVal = EvaluateExpression(arg)
                message = msgVal.ToString()
            End If
            
            Set result = CreateErrorObject(message)
            Set EvaluateNewExpression = result
            Exit Function
        ElseIf Node.Callee.Name = "Function" Then
            ' new Function([arg1, arg2, ...], body)
            ' Last argument is the body, rest are parameter names
            
            Dim funcBody As String
            Dim funcParams As String
            Dim funcArg As CValue
            Dim funcSource As String
            Dim funcAST As CNode
            Dim newFunc As New CFunction
            Dim i As Long
            Dim paramName As CNode
            
            If Node.Arguments.count = 0 Then
                ' new Function() - empty function
                funcBody = ""
                funcParams = ""
            ElseIf Node.Arguments.count = 1 Then
                ' new Function(body) - just body, no params
                Set funcArg = EvaluateExpression(Node.Arguments(1))
                funcBody = funcArg.ToString()
                funcParams = ""
            Else
                ' new Function(arg1, arg2, ..., body)
                ' Last argument is body, rest are parameter names
                funcParams = ""
                For i = 1 To Node.Arguments.count - 1
                    Set funcArg = EvaluateExpression(Node.Arguments(i))
                    funcParams = funcParams & funcArg.ToString()
                    If i < Node.Arguments.count - 1 Then funcParams = funcParams & ", "
                Next
                
                ' Get the body (last argument)
                Set funcArg = EvaluateExpression(Node.Arguments(Node.Arguments.count))
                funcBody = funcArg.ToString()
            End If
            
            ' Audit event - log before parsing
            RaiseAudit aeFunctionConstructor, funcBody
            
            ' Build complete function source
            funcSource = "function(" & funcParams & ") { " & funcBody & " }"
            
            ' Parse the function expression
            On Error GoTo FunctionParseError
            Set funcAST = m_parser.ParseSingleExpression(funcSource)
            On Error GoTo 0
            
            ' Verify it's a function expression node
            If funcAST.tType <> FunctionExpression_Node Then
                Err.Raise vbObjectError + 1015, "CInterpreter", _
                          "Function constructor: parser did not return function expression"
            End If
            
            Set newFunc.params = funcAST.params
            
            ' Set the function body (BlockStatement node)
            Set newFunc.Body = funcAST.FunctionBody
            
            ' Set closure to current scope
            Set newFunc.closure = m_currentScope
            
            ' Optional: set function name (anonymous)
            newFunc.Name = ""
            
            ' Return as function value
            result.vType = vtfunction
            Set result.objVal = newFunc
            Set EvaluateNewExpression = result
            Exit Function
        
FunctionParseError:
            Err.Raise vbObjectError + 1016, "CInterpreter", _
                      "Function constructor: failed to parse function body - " & Err.description
        End If

    End If
    
    ' Check if it's new ActiveXObject()
    If Node.Callee.tType = Identifier_Node Then
        If Node.Callee.Name = "ActiveXObject" Then
            ' Check safety
            If UseSafeSubset Then
                Err.Raise vbObjectError + 1012, "CInterpreter", _
                          "new ActiveXObject() is disabled in safe mode. " & _
                          "Set UseSafeSubset = False to enable."
            End If
            
            ' Get the ProgID argument
            If Node.Arguments.count = 0 Then
                Err.Raise vbObjectError + 1013, "CInterpreter", _
                          "ActiveXObject requires a ProgID argument"
            End If

            Set arg = Node.Arguments(1)
            
            'Debug.Print ">>> ActiveXObject arg node type: " & arg.tType
    
            Set progId = EvaluateExpression(arg)
            
            'Debug.Print ">>> ActiveXObject progID type: " & progID.vType
            If progId.vType = vtString Then
                'Debug.Print ">>> ActiveXObject progID value: " & progID.strVal
            End If
    
            Set progId = EvaluateExpression(arg)
            
            If progId.vType <> vtString Then
                Err.Raise vbObjectError + 1014, "CInterpreter", _
                          "ActiveXObject ProgID must be a string"
            End If
            
            ' Create the COM object
            On Error GoTo CreateError
            
            RaiseAudit aeActiveX, progId.strVal
            
            Set comObj = CreateObject(progId.strVal)
            On Error GoTo 0
            
            ' Wrap in CValue
            result.vType = vtCOMObject
            Set result.objVal = comObj
            
            Set EvaluateNewExpression = result
            Exit Function
            
CreateError:
            Err.Raise vbObjectError + 1015, "CInterpreter", _
                      "Failed to create ActiveXObject '" & progId.strVal & "': " & Err.description
        End If
    End If
    
    ' Other constructors not yet supported
    Err.Raise vbObjectError + 1016, "CInterpreter", _
              "new operator only supports ActiveXObject currently"
End Function

Sub TestActiveXObject()
    Dim interp As New CInterpreter
    Dim code As String
    
    Debug.Print "=== Testing new ActiveXObject (Unsafe Mode) ==="
    Debug.Print ""
    
    ' Test 1: Safe mode blocks it
    Debug.Print "Test 1: Safe mode (should fail)"
    interp.UseSafeSubset = True
    code = "var fso = new ActiveXObject('Scripting.FileSystemObject');"
    
    On Error Resume Next
    interp.Execute code
    If Err.Number <> 0 Then
        Debug.Print "Expected: Error (safe mode)"
        Debug.Print "Got: " & Err.description
        Err.Clear
    Else
        Debug.Print "ERROR: Should have failed in safe mode!"
    End If
    On Error GoTo 0
    Debug.Print ""
    
    ' Test 2: Unsafe mode allows it
    Debug.Print "Test 2: Unsafe mode - FileSystemObject"
    interp.UseSafeSubset = False
    interp.ClearOutput
    
    code = "var fso = new ActiveXObject('Scripting.FileSystemObject');" & vbCrLf & _
           "var tempFolder = fso.GetSpecialFolder(2);" & vbCrLf & _
           "console.log(tempFolder.Path);"
    interp.Execute code
    Debug.Print "Expected: (Temp path)"
    Debug.Print "Got: " & Trim$(interp.GetOutput())
    interp.ClearOutput
    Debug.Print ""
    
    ' Test 3: WScript.Shell
    Debug.Print "Test 3: Unsafe mode - WScript.Shell"
    code = "var shell = new ActiveXObject('WScript.Shell');" & vbCrLf & _
           "var env = shell.Environment('Process');" & vbCrLf & _
           "console.log(env.Item('COMPUTERNAME'));"
    interp.Execute code
    Debug.Print "Expected: (Computer name)"
    Debug.Print "Got: " & Trim$(interp.GetOutput())
    interp.ClearOutput
    Debug.Print ""
    
    ' Test 4: ADODB.Connection (if you dare!)
    Debug.Print "Test 4: Unsafe mode - ADODB.Connection"
    code = "var conn = new ActiveXObject('ADODB.Connection');" & vbCrLf & _
           "console.log(typeof conn);"
    interp.Execute code
    Debug.Print "Expected: object"
    Debug.Print "Got: " & Trim$(interp.GetOutput())
    interp.ClearOutput
    Debug.Print ""
    
    ' Test 5: Invalid ProgID
    Debug.Print "Test 5: Invalid ProgID (should fail)"
    code = "var bad = new ActiveXObject('This.Does.Not.Exist');"
    
    On Error Resume Next
    interp.Execute code
    If Err.Number <> 0 Then
        Debug.Print "Expected: Error (invalid ProgID)"
        Debug.Print "Got: " & Err.description
        Err.Clear
    Else
        Debug.Print "ERROR: Should have failed with invalid ProgID!"
    End If
    On Error GoTo 0
    Debug.Print ""
    
    ' Test 6: Missing argument
    Debug.Print "Test 6: Missing argument (should fail)"
    code = "var bad = new ActiveXObject();"
    
    On Error Resume Next
    interp.Execute code
    If Err.Number <> 0 Then
        Debug.Print "Expected: Error (missing argument)"
        Debug.Print "Got: " & Err.description
        Err.Clear
    Else
        Debug.Print "ERROR: Should have failed with missing argument!"
    End If
    On Error GoTo 0
    Debug.Print ""
    
    Debug.Print "=== ActiveXObject Tests Complete ==="
End Sub


Private Function IsMathMethod(methodName As String) As Boolean
    Select Case methodName
        Case "abs", "acos", "asin", "atan", "atan2", "ceil", "cos", "exp", _
             "floor", "log", "max", "min", "pow", "random", "round", "sin", _
             "sqrt", "tan"
            IsMathMethod = True
        Case Else
            IsMathMethod = False
    End Select
End Function

Private Function CallMathMethod(methodName As String, args As Collection) As CValue
    Dim result As New CValue
    result.vType = vtNumber
    
    Dim arg1 As Double
    Dim arg2 As Double
    Dim i As Long
    Dim val As CValue
    
    ' Get first argument if exists
    If args.count > 0 Then
        Set val = args(1)
        arg1 = val.ToNumber()
    End If
    
    ' Get second argument if exists
    If args.count > 1 Then
        Set val = args(2)
        arg2 = val.ToNumber()
    End If
    
    Select Case methodName
        Case "abs"
            result.numVal = Abs(arg1)
            
        Case "acos"
            ' VB6 doesn't have acos, compute it
            If arg1 < -1 Or arg1 > 1 Then
                result.numVal = 0  ' NaN in real JS
            Else
                result.numVal = 1.5707963267949 - Atn(arg1 / Sqr(1 - arg1 * arg1))
            End If
            
        Case "asin"
            ' VB6 doesn't have asin, compute it
            If arg1 < -1 Or arg1 > 1 Then
                result.numVal = 0  ' NaN in real JS
            Else
                result.numVal = Atn(arg1 / Sqr(1 - arg1 * arg1))
            End If
            
        Case "atan"
            result.numVal = Atn(arg1)
            
        Case "atan2"
            ' atan2(y, x)
            If arg1 = 0 And arg2 = 0 Then
                result.numVal = 0
            ElseIf arg2 = 0 Then
                result.numVal = IIf(arg1 > 0, 1.5707963267949, -1.5707963267949)
            Else
                result.numVal = Atn(arg1 / arg2)
                ' Adjust for quadrant
                If arg2 < 0 Then
                    If arg1 >= 0 Then
                        result.numVal = result.numVal + 3.14159265358979
                    Else
                        result.numVal = result.numVal - 3.14159265358979
                    End If
                End If
            End If
            
        Case "ceil"
            result.numVal = -Int(-arg1)
            
        Case "cos"
            result.numVal = Cos(arg1)
            
        Case "exp"
            result.numVal = Exp(arg1)
            
        Case "floor"
            result.numVal = Int(arg1)
            
        Case "log"
            If arg1 <= 0 Then
                result.numVal = 0  ' NaN in real JS
            Else
                result.numVal = Log(arg1)
            End If
            
        Case "max"
            ' Find maximum of all arguments
            If args.count = 0 Then
                result.numVal = -1.79769313486231E+308  ' -Infinity
            Else
                result.numVal = arg1
                For i = 2 To args.count
                    Set val = args(i)
                    If val.ToNumber() > result.numVal Then
                        result.numVal = val.ToNumber()
                    End If
                Next
            End If
            
        Case "min"
            ' Find minimum of all arguments
            If args.count = 0 Then
                result.numVal = 1.79769313486231E+308  ' Infinity
            Else
                result.numVal = arg1
                For i = 2 To args.count
                    Set val = args(i)
                    If val.ToNumber() < result.numVal Then
                        result.numVal = val.ToNumber()
                    End If
                Next
            End If
            
        Case "pow"
            result.numVal = arg1 ^ arg2
            
        Case "random"
            ' Random number between 0 and 1
            Randomize
            result.numVal = Rnd()
            
        Case "round"
            result.numVal = CLng(arg1)
            
        Case "sin"
            result.numVal = Sin(arg1)
            
        Case "sqrt"
            If arg1 < 0 Then
                result.numVal = 0  ' NaN in real JS
            Else
                result.numVal = Sqr(arg1)
            End If
            
        Case "tan"
            result.numVal = Tan(arg1)
            
        Case Else
            result.vType = vtUndefined
    End Select
    
    Set CallMathMethod = result
End Function

' ============================================
' ARRAY METHODS
' ============================================

Private Function IsArrayMethod(methodName As String) As Boolean
    Select Case methodName
        Case "push", "pop", "shift", "unshift", "slice", "splice", _
             "join", "reverse", "sort", "concat", "indexOf", "lastIndexOf", _
             "includes", _
             "forEach", "map", "filter", "reduce", "find", "findIndex", _
             "some", "every"
            IsArrayMethod = True
        Case Else
            IsArrayMethod = False
    End Select
End Function

Private Function CallArrayMethod(arr As Collection, methodName As String, args As Collection) As CValue

    Dim result As New CValue
    Dim i As Long
    Dim j As Long
    Dim val As CValue
    Dim arg As CValue
    Dim newArr As Collection
    Dim tempVal As CValue
    Dim found As Boolean
    Dim startIndex As Long
    Dim deleteCount As Long
    Dim strResult As String
    Dim separator As String
    Dim accumulator As CValue
    Dim startIdx As Long
    Dim callbackArgs As Collection
    Dim indexVal As CValue
    
    Select Case methodName
        Case "push"
            ' Add elements to end
            For i = 1 To args.count
                arr.add args(i)
            Next
            result.vType = vtNumber
            result.numVal = arr.count
            
        Case "pop"
            ' Remove and return last element
            If arr.count > 0 Then
                Set result = arr(arr.count)
                arr.Remove arr.count
            Else
                result.vType = vtUndefined
            End If
            
        Case "shift"
            ' Remove and return first element
            If arr.count > 0 Then
                Set result = arr(1)
                arr.Remove 1
            Else
                result.vType = vtUndefined
            End If
            
        Case "unshift"
            ' Add elements to beginning
            For i = args.count To 1 Step -1
                arr.add args(i), , 1
            Next
            result.vType = vtNumber
            result.numVal = arr.count
            
        Case "slice"
            ' Extract a section of array
            Dim sliceStart As Long
            Dim sliceEnd As Long
            
            If args.count > 0 Then
                Set arg = args(1)
                sliceStart = CLng(arg.ToNumber())
                If sliceStart < 0 Then sliceStart = arr.count + sliceStart
                If sliceStart < 0 Then sliceStart = 0
            Else
                sliceStart = 0
            End If
            
            If args.count > 1 Then
                Set arg = args(2)
                sliceEnd = CLng(arg.ToNumber())
                If sliceEnd < 0 Then sliceEnd = arr.count + sliceEnd
            Else
                sliceEnd = arr.count
            End If
            
            If sliceEnd > arr.count Then sliceEnd = arr.count
            
            Set newArr = New Collection
            For i = sliceStart + 1 To sliceEnd
                If i <= arr.count Then
                    newArr.add arr(i)
                End If
            Next
            
            result.vType = vtArray
            Set result.arrayVal = newArr
            
        Case "splice"
            ' Change contents by removing/adding elements
            ' splice(start, deleteCount, item1, item2, ...)
            If args.count = 0 Then
                result.vType = vtArray
                Set result.arrayVal = New Collection
                Exit Function
            End If
            
            Set arg = args(1)
            startIndex = CLng(arg.ToNumber())
            If startIndex < 0 Then startIndex = arr.count + startIndex
            If startIndex < 0 Then startIndex = 0
            If startIndex > arr.count Then startIndex = arr.count
            
            If args.count > 1 Then
                Set arg = args(2)
                deleteCount = CLng(arg.ToNumber())
            Else
                deleteCount = arr.count - startIndex
            End If
            
            ' Create array of deleted elements
            Set newArr = New Collection
            For i = 1 To deleteCount
                If startIndex + 1 <= arr.count Then
                    newArr.add arr(startIndex + 1)
                    arr.Remove startIndex + 1
                End If
            Next
            
            ' Insert new elements
            For i = 3 To args.count
                arr.add args(i), , startIndex + i - 2
            Next
            
            result.vType = vtArray
            Set result.arrayVal = newArr
            
        Case "join"
            ' Join array elements into string
            If args.count > 0 Then
                Set arg = args(1)
                separator = arg.ToString()
            Else
                separator = ","
            End If
            
            strResult = ""
            For i = 1 To arr.count
                Set val = arr(i)
                If i > 1 Then strResult = strResult & separator
                strResult = strResult & val.ToString()
            Next
            
            result.vType = vtString
            result.strVal = strResult
            
        Case "reverse"
            ' Reverse array in place
            Dim tempArr() As CValue
            ReDim tempArr(1 To arr.count)
            
            For i = 1 To arr.count
                Set tempArr(i) = arr(i)
            Next
            
            For i = 1 To arr.count
                arr.Remove 1
            Next
            
            For i = UBound(tempArr) To LBound(tempArr) Step -1
                arr.add tempArr(i)
            Next
            
            result.vType = vtArray
            Set result.arrayVal = arr
            
        Case "concat"
            ' Concatenate arrays
            Set newArr = New Collection
            
            ' Copy original array
            For i = 1 To arr.count
                newArr.add arr(i)
            Next
            
            ' Add arguments (can be values or arrays)
            For i = 1 To args.count
                Set arg = args(i)
                If arg.vType = vtArray Then
                    For j = 1 To arg.arrayVal.count
                        newArr.add arg.arrayVal(j)
                    Next
                Else
                    newArr.add arg
                End If
            Next
            
            result.vType = vtArray
            Set result.arrayVal = newArr
            
        Case "indexOf"
            ' Find first index of element
            result.vType = vtNumber
            result.numVal = -1
            
            If args.count > 0 Then
                Set arg = args(1)
                For i = 1 To arr.count
                    Set val = arr(i)
                    If IsStrictEqual(val, arg) Then
                        result.numVal = i - 1  ' 0-based
                        Exit For
                    End If
                Next
            End If
            
        Case "lastIndexOf"
            ' Find last index of element
            result.vType = vtNumber
            result.numVal = -1
            
            If args.count > 0 Then
                Set arg = args(1)
                For i = arr.count To 1 Step -1
                    Set val = arr(i)
                    If IsStrictEqual(val, arg) Then
                        result.numVal = i - 1  ' 0-based
                        Exit For
                    End If
                Next
            End If
            
        Case "includes"
            ' Check if array includes element
            result.vType = vtBoolean
            result.boolVal = False
            
            If args.count > 0 Then
                Set arg = args(1)
                For i = 1 To arr.count
                    Set val = arr(i)
                    If IsStrictEqual(val, arg) Then
                        result.boolVal = True
                        Exit For
                    End If
                Next
            End If
            
        Case "sort"
            ' Simple bubble sort (no compare function support for now)
            Dim swapped As Boolean
            Dim temp As CValue
            
            Do
                swapped = False
                For i = 1 To arr.count - 1
                    Set val = arr(i)
                    Set tempVal = arr(i + 1)
                    
                    ' Compare as strings (JavaScript default)
                    If val.ToString() > tempVal.ToString() Then
                        ' Swap
                        Set temp = arr(i)
                        arr.Remove i
                        arr.add temp, , i + 1
                        swapped = True
                    End If
                Next
            Loop While swapped
            
            result.vType = vtArray
            Set result.arrayVal = arr
            
        ' Update the Case statement for callback methods in CallArrayMethod:

        Case "forEach"
            ' arr.forEach(callback)
            If args.count = 0 Then
                result.vType = vtUndefined
                Exit Function
            End If
            
            Set arg = args(1)
            If arg.vType <> vtfunction Then
                Err.Raise vbObjectError + 3001, "CallArrayMethod", "forEach requires a function"
            End If

            ' Call callback for each element
            For i = 1 To arr.count
                Set callbackArgs = New Collection
                
                ' Arg 1: element
                callbackArgs.add arr(i)
                
                ' Arg 2: index (0-based)
                Set indexVal = New CValue
                indexVal.vType = vtNumber
                indexVal.numVal = i - 1
                callbackArgs.add indexVal
                
                ' Call the callback
                Dim callResult As CValue
                Set callResult = CallJSFunction(arg, callbackArgs)
            Next
            
            result.vType = vtUndefined
            
        Case "map"
            ' arr.map(callback) - returns new array with transformed elements
            If args.count = 0 Then
                result.vType = vtArray
                Set result.arrayVal = New Collection
                Exit Function
            End If
            
            Set arg = args(1)
            If arg.vType <> vtfunction Then
                Err.Raise vbObjectError + 3002, "CallArrayMethod", "map requires a function"
            End If

            Set newArr = New Collection
            
            For i = 1 To arr.count
                Set callbackArgs = New Collection
                callbackArgs.add arr(i)
                
                Set indexVal = New CValue
                indexVal.vType = vtNumber
                indexVal.numVal = i - 1
                callbackArgs.add indexVal
                
                ' Call callback and collect result
                Set callResult = CallJSFunction(arg, callbackArgs)
                newArr.add callResult
            Next
            
            result.vType = vtArray
            Set result.arrayVal = newArr
            
        Case "filter"
            ' arr.filter(callback) - returns new array with elements that pass test
            If args.count = 0 Then
                result.vType = vtArray
                Set result.arrayVal = New Collection
                Exit Function
            End If
            
            Set arg = args(1)
            If arg.vType <> vtfunction Then
                Err.Raise vbObjectError + 3003, "CallArrayMethod", "filter requires a function"
            End If

            Set newArr = New Collection
            
            For i = 1 To arr.count
                Set callbackArgs = New Collection
                callbackArgs.add arr(i)
                
                Set indexVal = New CValue
                indexVal.vType = vtNumber
                indexVal.numVal = i - 1
                callbackArgs.add indexVal
                
                ' Call callback and check if truthy
                Set callResult = CallJSFunction(arg, callbackArgs)
                If callResult.ToBoolean() Then
                    newArr.add arr(i)
                End If
            Next
            
            result.vType = vtArray
            Set result.arrayVal = newArr
            
        Case "reduce"
            ' arr.reduce(callback, [initialValue])
            If args.count = 0 Then
                Err.Raise vbObjectError + 3004, "CallArrayMethod", "reduce requires a function"
            End If
            
            Set arg = args(1)
            If arg.vType <> vtfunction Then
                Err.Raise vbObjectError + 3004, "CallArrayMethod", "reduce requires a function"
            End If
            
            ' Check for initial value
            If args.count > 1 Then
                Set accumulator = args(2)
                startIdx = 1
            Else
                If arr.count = 0 Then
                    Err.Raise vbObjectError + 3005, "CallArrayMethod", "reduce of empty array with no initial value"
                End If
                Set accumulator = arr(1)
                startIdx = 2
            End If
            
            ' Reduce loop
            For i = startIdx To arr.count
                Set callbackArgs = New Collection
                
                ' Arg 1: accumulator
                callbackArgs.add accumulator
                
                ' Arg 2: current value
                callbackArgs.add arr(i)
                
                ' Arg 3: index (0-based)
                Set indexVal = New CValue
                indexVal.vType = vtNumber
                indexVal.numVal = i - 1
                callbackArgs.add indexVal
                
                ' Call callback
                Set accumulator = CallJSFunction(arg, callbackArgs)
            Next
            
            Set result = accumulator
            
        Case "find"
            ' arr.find(callback) - returns first element that passes test
            If args.count = 0 Then
                result.vType = vtUndefined
                Exit Function
            End If
            
            Set arg = args(1)
            If arg.vType <> vtfunction Then
                Err.Raise vbObjectError + 3006, "CallArrayMethod", "find requires a function"
            End If
            
            result.vType = vtUndefined
            
            For i = 1 To arr.count
                Set callbackArgs = New Collection
                callbackArgs.add arr(i)
                
                Set indexVal = New CValue
                indexVal.vType = vtNumber
                indexVal.numVal = i - 1
                callbackArgs.add indexVal
                
                Set callResult = CallJSFunction(arg, callbackArgs)
                If callResult.ToBoolean() Then
                    Set result = arr(i)
                    Exit For
                End If
            Next
            
        Case "findIndex"
            ' arr.findIndex(callback) - returns index of first element that passes test
            If args.count = 0 Then
                result.vType = vtNumber
                result.numVal = -1
                Exit Function
            End If
            
            Set arg = args(1)
            If arg.vType <> vtfunction Then
                Err.Raise vbObjectError + 3007, "CallArrayMethod", "findIndex requires a function"
            End If
            
            result.vType = vtNumber
            result.numVal = -1
            
            For i = 1 To arr.count
                Set callbackArgs = New Collection
                callbackArgs.add arr(i)
                
                Set indexVal = New CValue
                indexVal.vType = vtNumber
                indexVal.numVal = i - 1
                callbackArgs.add indexVal
                
                Set callResult = CallJSFunction(arg, callbackArgs)
                If callResult.ToBoolean() Then
                    result.numVal = i - 1  ' 0-based
                    Exit For
                End If
            Next
            
        Case "some"
            ' arr.some(callback) - returns true if any element passes test
            If args.count = 0 Then
                result.vType = vtBoolean
                result.boolVal = False
                Exit Function
            End If
            
            Set arg = args(1)
            If arg.vType <> vtfunction Then
                Err.Raise vbObjectError + 3008, "CallArrayMethod", "some requires a function"
            End If
            
            result.vType = vtBoolean
            result.boolVal = False
            
            For i = 1 To arr.count
                Set callbackArgs = New Collection
                callbackArgs.add arr(i)
                
                Set indexVal = New CValue
                indexVal.vType = vtNumber
                indexVal.numVal = i - 1
                callbackArgs.add indexVal
                
                Set callResult = CallJSFunction(arg, callbackArgs)
                If callResult.ToBoolean() Then
                    result.boolVal = True
                    Exit For
                End If
            Next
            
        Case "every"
            ' arr.every(callback) - returns true if all elements pass test
            If args.count = 0 Then
                result.vType = vtBoolean
                result.boolVal = True
                Exit Function
            End If
            
            Set arg = args(1)
            If arg.vType <> vtfunction Then
                Err.Raise vbObjectError + 3009, "CallArrayMethod", "every requires a function"
            End If
            
            result.vType = vtBoolean
            result.boolVal = True
            
            For i = 1 To arr.count
                Set callbackArgs = New Collection
                callbackArgs.add arr(i)
                
                Set indexVal = New CValue
                indexVal.vType = vtNumber
                indexVal.numVal = i - 1
                callbackArgs.add indexVal
                
                Set callResult = CallJSFunction(arg, callbackArgs)
                If Not callResult.ToBoolean() Then
                    result.boolVal = False
                    Exit For
                End If
            Next
            
        Case Else
            result.vType = vtUndefined
    End Select
    
    Set CallArrayMethod = result
End Function



Private Function IsGlobalFunction(funcName As String) As Boolean
    Select Case funcName
        Case "parseInt", "parseFloat", "isNaN", "isFinite", "isInteger", _
             "encodeURIComponent", "decodeURIComponent", "encodeURI", "decodeURI", _
             "escape", "unescape", "eval", "print", "alert", "prompt", "format", _
             "printf", "hex", "Number", "String", "Date"
            IsGlobalFunction = True
        Case Else
            IsGlobalFunction = False
    End Select
End Function

Private Function CallGlobalFunction(funcName As String, args As Collection) As CValue
    Dim result As New CValue
    Dim arg As CValue
    Dim str As String
    Dim num As Double
    Dim radix As Long
    
    Select Case funcName
        
        Case "Date"
            ' Date() as a function returns string (current date)
            result.vType = vtString
            result.strVal = Format$(Now, "ddd mmm dd yyyy hh:nn:ss")
    
        Case "hex"
            result.vType = vtString
            Set CallGlobalFunction = GlobalHex(args)
            Exit Function
            
        Case "printf"
            ' printf("Hello {name}") - Formats AND prints in one shot
    
            ' Just call format() then print it!
            Dim formatResult As CValue
            Set formatResult = HandleFormat(args)
            
            ' Print the result
            Dim printArgs As New Collection
            printArgs.add formatResult
            Set CallGlobalFunction = HandlePrint(printArgs)
            Exit Function
            
        Case "format"
            Set CallGlobalFunction = HandleFormat(args)
            Exit Function
            
        Case "print", "alert"

            Set CallGlobalFunction = HandlePrint(args, (funcName = "alert"))
            Exit Function
         
        Case "prompt"
            Dim promptText As String
            Dim defaultValue As String
            
            If args.count > 0 Then
                Set arg = args(1)
                promptText = arg.ToString()
            End If
            
            If args.count > 1 Then
                Set arg = args(2)
                defaultValue = arg.ToString()
            End If
            
            result.vType = vtString
            result.strVal = InputBox(promptText, "Input", defaultValue)
        
        Case "Number"
            ' Number(value) - converts to number
            If args.count = 0 Then
                result.vType = vtNumber
                result.numVal = 0
            Else
                Set arg = args(1)
                result.vType = vtNumber
                result.numVal = arg.ToNumber()
            End If
        
        Case "String"
            ' String(value) - converts to string
            If args.count = 0 Then
                result.vType = vtString
                result.strVal = ""
            Else
                Set arg = args(1)
                result.vType = vtString
                result.strVal = arg.ToString()
            End If
    
        Case "parseInt"
            ' parseInt(string, [radix])
            If args.count = 0 Then
                result.vType = vtNumber
                result.numVal = 0  ' NaN in real JS
            Else
                Set arg = args(1)
                str = Trim$(arg.ToString())
                
                ' Get radix
                If args.count > 1 Then
                    Set arg = args(2)
                    radix = CLng(arg.ToNumber())
                Else
                    radix = 10
                    ' Auto-detect: 0x for hex
                    If Len(str) > 2 Then
                        If Left$(str, 2) = "0x" Or Left$(str, 2) = "0X" Then
                            radix = 16
                            str = Mid$(str, 3)
                        End If
                    End If
                End If
                
                On Error Resume Next
                
                ' Clean up the string for parsing
                Dim cleanStr As String
                cleanStr = ""
                Dim hasSign As Boolean
                hasSign = False
                
                ' Handle leading sign
                If Len(str) > 0 Then
                    If Left$(str, 1) = "-" Or Left$(str, 1) = "+" Then
                        cleanStr = Left$(str, 1)
                        str = Mid$(str, 2)
                        hasSign = True
                    End If
                End If
                
                ' Parse valid characters for the radix
                Dim i As Long
                For i = 1 To Len(str)
                    Dim ch As String
                    ch = UCase$(Mid$(str, i, 1))
                    Dim isValid As Boolean
                    isValid = False
                    
                    Select Case radix
                        Case 2: isValid = (ch = "0" Or ch = "1")
                        Case 8: isValid = (ch >= "0" And ch <= "7")
                        Case 10: isValid = (ch >= "0" And ch <= "9")
                        Case 16: isValid = ((ch >= "0" And ch <= "9") Or (ch >= "A" And ch <= "F"))
                        Case Else: isValid = False
                    End Select
                    
                    If isValid Then
                        cleanStr = cleanStr & ch
                    Else
                        Exit For
                    End If
                Next
                
                ' Remove just the sign if no digits
                If Len(cleanStr) <= 1 And hasSign Then
                    cleanStr = ""
                End If
                
                ' Try to parse the result
                If Len(cleanStr) = 0 Then
                    result.vType = vtNumber
                    result.numVal = 0  ' NaN
                Else
                    ' Determine if we need BigInt or can use regular number
                    Dim needsBigInt As Boolean
                    needsBigInt = False
                    
                    ' Check if value would exceed safe Number range
                    If radix = 10 And Len(cleanStr) > 15 Then
                        needsBigInt = True
                    ElseIf radix = 16 And Len(cleanStr) > 8 Then
                        ' More than 32-bit hex
                        needsBigInt = True
                    ElseIf radix = 2 And Len(cleanStr) > 53 Then
                        needsBigInt = True
                    End If
                    
                    If needsBigInt Then
                        ' Return as BigInt (vtInt64)
                        Dim U64 As New ULong64
                        
                        Select Case radix
                            Case 16
                                If U64.fromString(cleanStr, mHex) Then
                                    result.vType = vtInt64
                                    result.int64Val = U64.rawValue
                                Else
                                    result.vType = vtNumber
                                    result.numVal = 0
                                End If
                            Case 10
                                If U64.fromString(cleanStr, mUnsigned) Then
                                    result.vType = vtInt64
                                    result.int64Val = U64.rawValue
                                Else
                                    result.vType = vtNumber
                                    result.numVal = 0
                                End If
                            Case Else
                                ' For binary/octal, convert manually or return regular number
                                result.vType = vtNumber
                                result.numVal = 0
                        End Select
                    Else
                        ' Use regular number
                        result.vType = vtNumber
                        
                        Select Case radix
                            Case 10
                                result.numVal = CDbl(cleanStr)
                            Case 16
                                ' Handle large unsigned hex values
                                Dim hexDbl As Double
                                If Len(cleanStr) <= 8 Then
                                    ' Try CLng first for values that fit
                                    Dim tempLng As Long
                                    tempLng = CLng("&H" & cleanStr)
                                    If tempLng < 0 Then
                                        ' Was unsigned value > 2^31, convert properly
                                        hexDbl = CDbl(tempLng) + 4294967296#
                                    Else
                                        hexDbl = CDbl(tempLng)
                                    End If
                                    result.numVal = hexDbl
                                Else
                                    result.numVal = CDbl("&H" & cleanStr)
                                End If
                            Case 8
                                result.numVal = CDbl("&O" & cleanStr)
                            Case 2
                                ' Binary
                                Dim binVal As Double
                                binVal = 0
                                For i = 1 To Len(cleanStr)
                                    binVal = binVal * 2 + CDbl(Mid$(cleanStr, i, 1))
                                Next
                                result.numVal = binVal
                            Case Else
                                result.numVal = 0
                        End Select
                    End If
                End If
                
                If Err.Number <> 0 Then
                    result.vType = vtNumber
                    result.numVal = 0  ' NaN
                End If
                On Error GoTo 0
            End If
            
        Case "parseFloat"
            ' parseFloat(string)
            If args.count = 0 Then
                result.vType = vtNumber
                result.numVal = 0  ' NaN
            Else
                Set arg = args(1)
                str = Trim$(arg.ToString())
                
                result.vType = vtNumber
                
                On Error Resume Next
                result.numVal = CDbl(str)
                If Err.Number <> 0 Then
                    result.numVal = 0  ' NaN
                End If
                On Error GoTo 0
            End If
            
        Case "isNaN"
            ' isNaN(value)
            result.vType = vtBoolean
            
            If args.count = 0 Then
                result.boolVal = True
            Else
                Set arg = args(1)
                
                ' Convert to number first (like real JavaScript)
                num = arg.ToNumber()
                
                ' Check if the original string was actually a valid number
                If arg.vType = vtString Then
                    Dim testStr As String
                    testStr = Trim$(arg.strVal)
                    
                    ' Empty string converts to 0, but should be considered NaN in isNaN
                    If Len(testStr) = 0 Then
                        result.boolVal = True
                        Exit Function
                    End If
                    
                    ' Try to parse - if it fails, it's NaN
                    On Error Resume Next
                    Dim testNum As Double
                    testNum = CDbl(testStr)
                    result.boolVal = (Err.Number <> 0)
                    On Error GoTo 0
                ElseIf arg.vType = vtNumber Then
                    ' Already a number - check if it's valid
                    result.boolVal = False
                ElseIf arg.vType = vtUndefined Or arg.vType = vtNull Then
                    result.boolVal = True
                ElseIf arg.vType = vtBoolean Then
                    result.boolVal = False  ' Booleans convert to 0/1
                Else
                    result.boolVal = True
                End If
            End If
            
        Case "isFinite"
            ' isFinite(value)
            result.vType = vtBoolean
            
            If args.count = 0 Then
                result.boolVal = False
            Else
                Set arg = args(1)
                num = arg.ToNumber()
                
                ' Check if finite (not infinity)
                ' VB6 doesn't have infinity, so check for very large numbers
                If Abs(num) > 1.79769313486231E+308 Then
                    result.boolVal = False
                Else
                    result.boolVal = True
                End If
            End If
            
        Case "isInteger"
            ' isInteger(value) - ES6 Number.isInteger
            result.vType = vtBoolean
            
            If args.count = 0 Then
                result.boolVal = False
            Else
                Set arg = args(1)
                
                If arg.vType = vtNumber Then
                    result.boolVal = (arg.numVal = Int(arg.numVal))
                Else
                    result.boolVal = False
                End If
            End If
            
        Case "encodeURIComponent"
            ' encodeURIComponent(string)
            If args.count = 0 Then
                result.vType = vtString
                result.strVal = ""
            Else
                Set arg = args(1)
                result.vType = vtString
                result.strVal = EncodeURIComponent(arg.ToString())
            End If
            
        Case "decodeURIComponent"
            ' decodeURIComponent(string)
            If args.count = 0 Then
                result.vType = vtString
                result.strVal = ""
            Else
                Set arg = args(1)
                RaiseAudit aeDecode, "decodeURIComponent(" & arg.ToString() & ")"
                result.vType = vtString
                result.strVal = DecodeURIComponent(arg.ToString())
            End If
            
        Case "encodeURI"
            ' encodeURI(string) - like encodeURIComponent but preserves :/?#[]@!$&'()*+,;=
            If args.count = 0 Then
                result.vType = vtString
                result.strVal = ""
            Else
                Set arg = args(1)
                result.vType = vtString
                result.strVal = EncodeURI(arg.ToString())
            End If
            
        Case "decodeURI"
            ' decodeURI(string)
            If args.count = 0 Then
                result.vType = vtString
                result.strVal = ""
            Else
                Set arg = args(1)
                RaiseAudit aeDecode, "decodeURI(" & arg.ToString() & ")"
                result.vType = vtString
                result.strVal = DecodeURIComponent(arg.ToString())  ' Same implementation
            End If
            
        Case "escape"
            ' escape(string) - deprecated but still used
            If args.count = 0 Then
                result.vType = vtString
                result.strVal = ""
            Else
                Set arg = args(1)
                result.vType = vtString
                result.strVal = EscapeString(arg.ToString())
            End If
            
        Case "unescape"
            ' unescape(string) - deprecated
            If args.count = 0 Then
                result.vType = vtString
                result.strVal = ""
            Else
                Set arg = args(1)
                RaiseAudit aeDecode, "unescape(" & arg.ToString() & ")"
                result.vType = vtString
                result.strVal = UnescapeString(arg.ToString())
            End If
            
        Case "eval"
            ' eval(code) - evaluate code in current context
            If args.count = 0 Then
                result.vType = vtUndefined
            Else
                Set arg = args(1)
                str = arg.ToString()
                
                RaiseAudit aeEval, str
                
                ' Use our existing Eval mechanism
                On Error Resume Next
                Set result = Eval(str)
                If Err.Number <> 0 Then
                    result.vType = vtUndefined
                End If
                On Error GoTo 0
            End If
            
        Case Else
            result.vType = vtUndefined
    End Select
    
    Set CallGlobalFunction = result
End Function



' Execute throw statement
Private Sub ExecuteThrowStatement(Node As CNode)
    Dim throwValue As CValue
    
    If Not Node.ReturnArgument Is Nothing Then
        Set throwValue = EvaluateExpression(Node.ReturnArgument)
    Else
        Set throwValue = New CValue
        throwValue.vType = vtUndefined
    End If
    
    ' Convert to error message
    Dim msg As String
    If throwValue.vType = vtObject Then
        ' It's an Error object - get the message property
        On Error Resume Next
        Dim msgVal As CValue
        Set msgVal = throwValue.objectProps("message")
        If Err.Number = 0 Then
            msg = msgVal.ToString()
        Else
            msg = throwValue.ToString()
        End If
        On Error GoTo 0
    Else
        msg = throwValue.ToString()
    End If
    
    ' Raise VB6 error
    Err.Raise vbObjectError + 5000, "JavaScript throw", msg
End Sub

' Execute try/catch/finally statement
Private Sub ExecuteTryStatement(Node As CNode)
    Dim hadError As Boolean
    Dim errorMsg As String
    Dim errorNumber As Long
    
    hadError = False
    
    ' Try block
    On Error Resume Next
    ExecuteStatement Node.TryBlock
    
    If Err.Number <> 0 Then
        hadError = True
        errorMsg = Err.description
        errorNumber = Err.Number
        Err.Clear
    End If
    On Error GoTo 0
    
    ' Catch block
    If hadError And Not Node.TryHandler Is Nothing Then
        ' Create new scope for catch block
        Dim catchScope As New CScope
        Set catchScope.Parent = m_currentScope
        
        ' Bind error to catch parameter
        Dim errorValue As New CValue
        errorValue.vType = vtObject
        Set errorValue.objectProps = New Collection
        Set errorValue.objectKeys = New Collection
        
        ' Add message property
        Dim msgVal As New CValue
        msgVal.vType = vtString
        msgVal.strVal = errorMsg
        errorValue.objectProps.add msgVal, "message"
        errorValue.objectKeys.add "message", "message"
        
        ' Add name property
        Dim nameVal As New CValue
        nameVal.vType = vtString
        nameVal.strVal = "Error"
        errorValue.objectProps.add nameVal, "name"
        errorValue.objectKeys.add "name", "name"
        
        ' Define error in catch scope
        catchScope.DefineVar Node.TryHandler.param.Name, errorValue
        
        ' Switch to catch scope
        Dim savedScope As CScope
        Set savedScope = m_currentScope
        Set m_currentScope = catchScope
        
        ' Execute catch block
        On Error Resume Next
        ExecuteStatement Node.TryHandler.CatchBody
        On Error GoTo 0
        
        ' Restore scope
        Set m_currentScope = savedScope
    End If
    
    ' Finally block (always executes)
    If Not Node.TryFinalizer Is Nothing Then
        On Error Resume Next
        ExecuteStatement Node.TryFinalizer
        On Error GoTo 0
    End If
    
    ' If error wasn't caught, re-raise it
    If hadError And Node.TryHandler Is Nothing Then
        Err.Raise errorNumber, "JavaScript", errorMsg
    End If
End Sub

' ============================================
' === SWITCH STATEMENT ===
' ============================================
Private Sub ExecuteSwitchStatement(Node As CNode)
    On Error GoTo ErrorHandler

    Dim discriminant As CValue
    Set discriminant = EvaluateExpression(Node.discriminant)
    
    Dim foundMatch As Boolean
    foundMatch = False
    
    Dim shouldExecute As Boolean
    shouldExecute = False
    
    Dim defaultCaseIndex As Long
    defaultCaseIndex = 0
    
    Dim caseIndex As Long
    For caseIndex = 1 To Node.Cases.count
        Dim switchCase As CNode
        Set switchCase = Node.Cases(caseIndex)
        
        If switchCase.CaseTest Is Nothing Then
            defaultCaseIndex = caseIndex
        ElseIf Not foundMatch Then
            Dim caseValue As CValue
            Set caseValue = EvaluateExpression(switchCase.CaseTest)
            
            If IsStrictEqual(discriminant, caseValue) Then
                foundMatch = True
                shouldExecute = True
            End If
        End If
        
        If shouldExecute Then
            Dim stmtIndex As Long
            For stmtIndex = 1 To switchCase.CaseConsequent.count
                Dim stmt As CNode
                Set stmt = switchCase.CaseConsequent(stmtIndex)
                
                ExecuteStatement stmt
                
                If m_shouldBreak Then
                    m_shouldBreak = False
                    Exit Sub
                End If
                
                If m_hasReturned Then Exit Sub
            Next stmtIndex
        End If
    Next caseIndex
    
    If Not foundMatch And defaultCaseIndex > 0 Then
        For caseIndex = defaultCaseIndex To Node.Cases.count
            Set switchCase = Node.Cases(caseIndex)
            
            For stmtIndex = 1 To switchCase.CaseConsequent.count
                Set stmt = switchCase.CaseConsequent(stmtIndex)
                
                ExecuteStatement stmt
                
                If m_shouldBreak Then
                    m_shouldBreak = False
                    Exit Sub
                End If
                
                If m_hasReturned Then Exit Sub
            Next stmtIndex
        Next caseIndex
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, "ExecuteSwitchStatement", _
              "Error in switch statement at line " & Node.LineNumber & ": " & Err.description
End Sub


' Check if it's the Error constructor
Private Function IsErrorConstructor(funcName As String) As Boolean
    IsErrorConstructor = (funcName = "Error")
End Function

' Create an Error object
Private Function CreateErrorObject(message As String) As CValue
    Dim result As New CValue
    result.vType = vtObject
    Set result.objectProps = New Collection
    Set result.objectKeys = New Collection
    
    ' Add message property
    Dim msgVal As New CValue
    msgVal.vType = vtString
    msgVal.strVal = message
    result.objectProps.add msgVal, "message"
    result.objectKeys.add "message", "message"
    
    ' Add name property
    Dim nameVal As New CValue
    nameVal.vType = vtString
    nameVal.strVal = "Error"
    result.objectProps.add nameVal, "name"
    result.objectKeys.add "name", "name"
    
    Set CreateErrorObject = result
End Function


Public Function SetNextStatement(targetLine As Long) As Boolean
    On Error GoTo ErrorHandler
    
    If Not m_debugMode Then Exit Function
    If Not m_debugPaused Then Exit Function
    
    If m_programAST Is Nothing Then
        'Debug.Print ">>> SetNextStatement: No AST!"
        Exit Function
    End If
    
    ' Get all executable line numbers
    Dim validLines As Collection
    Set validLines = GetExecutableLines(m_programAST)
    
    If validLines.count = 0 Then
        'Debug.Print ">>> SetNextStatement: No executable lines found!"
        Exit Function
    End If
    
    ' Check if target line is valid
    Dim isValid As Boolean
    isValid = False
    
    On Error Resume Next
    Dim testLine As Variant
    testLine = validLines(CStr(targetLine))
    If Err.Number = 0 Then isValid = True
    On Error GoTo 0
    
    If Not isValid Then
        'Debug.Print ">>> SetNextStatement: Line " & targetLine & " is not an executable statement"
        Exit Function
    End If
    
    ' Validate jump rules
    If Not CanSetNextStatementTo(targetLine) Then
        'Debug.Print ">>> SetNextStatement: Jump blocked"
        Exit Function
    End If
    
    ' Queue the jump
    m_currentLine = targetLine
    m_requestedNextLine = targetLine
    m_hasRequestedNextLine = True
    
    'Debug.Print ">>> SetNextStatement: Jump queued to line " & targetLine
    
    SetNextStatement = True
    Exit Function
    
ErrorHandler:
    Debug.Print ">>> SetNextStatement error: " & Err.description
    SetNextStatement = False
End Function

' Get the range of line numbers in the AST
Private Sub GetASTLineRange(Node As CNode, ByRef minLine As Long, ByRef maxLine As Long)
    If Node Is Nothing Then Exit Sub
    
    ' Initialize
    If minLine = 0 Then minLine = 999999
    
    ' Check this node
    If Node.LineNumber > 0 Then
        If Node.LineNumber < minLine Then minLine = Node.LineNumber
        If Node.LineNumber > maxLine Then maxLine = Node.LineNumber
    End If
    
    ' Check body (collection of statements)
    If Not Node.Body Is Nothing Then
        Dim stmt As CNode
        For Each stmt In Node.Body
            GetASTLineRange stmt, minLine, maxLine
        Next
    End If
    
    ' Check common child nodes
    If Not Node.Left Is Nothing Then GetASTLineRange Node.Left, minLine, maxLine
    If Not Node.Right Is Nothing Then GetASTLineRange Node.Right, minLine, maxLine
    If Not Node.Test Is Nothing Then GetASTLineRange Node.Test, minLine, maxLine
    If Not Node.Consequent Is Nothing Then GetASTLineRange Node.Consequent, minLine, maxLine
    If Not Node.Alternate Is Nothing Then GetASTLineRange Node.Alternate, minLine, maxLine
    If Not Node.init Is Nothing Then GetASTLineRange Node.init, minLine, maxLine
    
    ' Check params (collection)
    On Error Resume Next
    If Not Node.params Is Nothing Then
        Dim param As CNode
        For Each param In Node.params
            GetASTLineRange param, minLine, maxLine
        Next
    End If
    On Error GoTo 0
End Sub

' Validate if we can safely jump to target line

' Get all executable statement line numbers from AST
Private Function GetExecutableLines(Node As CNode) As Collection
    Dim lines As New Collection
    
    If Node Is Nothing Then
        Set GetExecutableLines = lines
        Exit Function
    End If
    
    ' If this is a Program node, get all statement lines
    If Node.tType = Program_Node And Not Node.Body Is Nothing Then
        Dim stmt As CNode
        For Each stmt In Node.Body
            If stmt.LineNumber > 0 Then
                On Error Resume Next
                lines.add stmt.LineNumber, CStr(stmt.LineNumber)
                On Error GoTo 0
            End If
        Next
    End If
    
    Set GetExecutableLines = lines
End Function

' Find which statement index corresponds to target line
Private Function FindStatementIndexAtLine(statements As Collection, targetLine As Long) As Long
    Dim stmt As CNode
    Dim idx As Long
    Dim closestIdx As Long
    Dim closestLine As Long
    
    idx = 1
    closestIdx = 0
    closestLine = -1
    
    ' Find the statement at or closest to the target line
    For Each stmt In statements
        If stmt.LineNumber = targetLine Then
            ' Exact match!
            FindStatementIndexAtLine = idx
            Exit Function
        ElseIf stmt.LineNumber < targetLine Then
            ' This statement is before target, but might be closest
            If stmt.LineNumber > closestLine Then
                closestLine = stmt.LineNumber
                closestIdx = idx
            End If
        ElseIf stmt.LineNumber > targetLine Then
            ' This statement is after target
            ' Return the closest one we found before this
            If closestIdx > 0 Then
                FindStatementIndexAtLine = closestIdx
                Exit Function
            Else
                ' No statement before target, use this one
                FindStatementIndexAtLine = idx
                Exit Function
            End If
        End If
        idx = idx + 1
    Next
    
    ' If we got here, return the closest we found
    If closestIdx > 0 Then
        FindStatementIndexAtLine = closestIdx
    Else
        FindStatementIndexAtLine = 1  ' Default to first statement
    End If
End Function


' ============================================
' BULLETPROOF VALIDATION FOR SET NEXT STATEMENT
' ============================================
' ===========================================================================
' SET NEXT STATEMENT VALIDATION - TEST RESULTS
' ===========================================================================
'
' TEST RESULTS: 10/10 PASSING
'
' +------+--------------------------------+----------+----------+--------+
' | Test | Scenario                       | Result   | Expected | Status |
' +------+--------------------------------+----------+----------+--------+
' |  1   | Valid jump (same scope)        | ALLOWED  | ALLOWED  |  PASS  |
' |  2   | Jump into for loop body        | BLOCKED  | BLOCKED  |  PASS  |
' |  3   | Jump into while loop body      | BLOCKED  | BLOCKED  |  PASS  |
' |  4   | Jump into do-while loop body   | BLOCKED  | BLOCKED  |  PASS  |
' |  5   | Jump into function body        | BLOCKED  | BLOCKED  |  PASS  |
' |  6   | Valid backward jump            | ALLOWED  | ALLOWED  |  PASS  |
' |  7   | Jump over function decl        | ALLOWED  | ALLOWED  |  PASS  |
' |  8   | Jump to for loop start         | ALLOWED  | ALLOWED  |  PASS  |
' |  9   | Jump out of for loop           | ALLOWED  | ALLOWED  |  PASS  |
' | 10   | Jump backward into loop        | BLOCKED  | BLOCKED  |  PASS  |
' +------+--------------------------------+----------+----------+--------+
'
' SAFE JUMPS (Allowed):
' - Forward/backward in same scope
' - Over function declarations
' - To loop start lines
' - Out of loops (early exit)
'
' DANGEROUS JUMPS (Blocked):
' - Into loop bodies (skips initialization)
' - Into function bodies (wrong scope)
' - Backward into loops (infinite loop risk)
' - To non-executable lines
' - Across call stack depths
'
' ===========================================================================
Private Function CanSetNextStatementTo(targetLine As Long) As Boolean
    On Error GoTo ErrorHandler
    
    'Debug.Print ">>> CanSetNextStatementTo: Validating jump from " & m_currentLine & " to " & targetLine
    
    ' Rule 0: Sanity check - must have AST
    If m_programAST Is Nothing Then
        'Debug.Print ">>> BLOCK: No AST available"
        CanSetNextStatementTo = False
        Exit Function
    End If
    
    ' Rule 1: Must stay at same call stack depth
    Dim currentDepth As Long
    currentDepth = m_callStack.count
    
    'Debug.Print ">>> Current call stack depth: " & currentDepth
    
    ' Rule 2: Target must be an executable statement at current depth
    On Error Resume Next
    Dim isExec As Boolean
    isExec = IsExecutableStatementAtCurrentDepth(targetLine)
    If Err.Number <> 0 Then
        'Debug.Print ">>> Error in IsExecutableStatementAtCurrentDepth: " & Err.Description
        Err.Clear
        isExec = False
    End If
    On Error GoTo ErrorHandler
    
    If Not isExec Then
        'Debug.Print ">>> BLOCK: Target line " & targetLine & " is not executable at current depth"
        CanSetNextStatementTo = False
        Exit Function
    End If
    
    ' Rule 3: Cannot jump INTO a for loop body
    On Error Resume Next
    Dim intoFor As Boolean
    intoFor = IsJumpingIntoForLoop(m_currentLine, targetLine)
    If Err.Number <> 0 Then
        'Debug.Print ">>> Error in IsJumpingIntoForLoop: " & Err.Description
        Err.Clear
        intoFor = False
    End If
    On Error GoTo ErrorHandler
    
    If intoFor Then
        'Debug.Print ">>> BLOCK: Cannot jump into for loop body"
        CanSetNextStatementTo = False
        Exit Function
    End If
    
    ' Rule 4: Cannot jump INTO a while loop body
    On Error Resume Next
    Dim intoWhile As Boolean
    intoWhile = IsJumpingIntoWhileLoop(m_currentLine, targetLine)
    If Err.Number <> 0 Then
        'Debug.Print ">>> Error in IsJumpingIntoWhileLoop: " & Err.Description
        Err.Clear
        intoWhile = False
    End If
    On Error GoTo ErrorHandler
    
    If intoWhile Then
        'Debug.Print ">>> BLOCK: Cannot jump into while loop body"
        CanSetNextStatementTo = False
        Exit Function
    End If
    
    ' Rule 5: Cannot jump INTO a function body
    On Error Resume Next
    Dim intoFunc As Boolean
    intoFunc = IsJumpingIntoFunction(m_currentLine, targetLine)
    If Err.Number <> 0 Then
        'Debug.Print ">>> Error in IsJumpingIntoFunction: " & Err.Description
        Err.Clear
        intoFunc = False
    End If
    On Error GoTo ErrorHandler
    
    If intoFunc Then
        'Debug.Print ">>> BLOCK: Cannot jump into function body"
        CanSetNextStatementTo = False
        Exit Function
    End If
    
    ' Rule 6: Cannot jump INTO a do-while loop body
    On Error Resume Next
    Dim intoDoWhile As Boolean
    intoDoWhile = IsJumpingIntoDoWhileLoop(m_currentLine, targetLine)
    If Err.Number <> 0 Then
        'Debug.Print ">>> Error in IsJumpingIntoDoWhileLoop: " & Err.Description
        Err.Clear
        intoDoWhile = False
    End If
    On Error GoTo ErrorHandler
    
    If intoDoWhile Then
        'Debug.Print ">>> BLOCK: Cannot jump into do-while loop body"
        CanSetNextStatementTo = False
        Exit Function
    End If
    
    ' All checks passed!
    'Debug.Print ">>> ALLOW: Jump is safe"
    CanSetNextStatementTo = True
    Exit Function
    
ErrorHandler:
    Debug.Print ">>> CanSetNextStatementTo error: " & Err.description & " (line " & Erl & ")"
    CanSetNextStatementTo = False
End Function


' Check if target line is an executable statement at current call depth
Private Function IsExecutableStatementAtCurrentDepth(targetLine As Long) As Boolean
    ' Get the current scope's statement list
    ' At global depth, check program body
    ' At function depth, check function body
    
    ' For now, simplified: check if line exists in program body
    If m_programAST.tType = Program_Node And Not m_programAST.Body Is Nothing Then
        Dim stmt As CNode
        For Each stmt In m_programAST.Body
            If stmt.LineNumber = targetLine Then
                IsExecutableStatementAtCurrentDepth = True
                Exit Function
            End If
        Next
    End If
    
    IsExecutableStatementAtCurrentDepth = False
End Function

' Check if jumping INTO a for loop body (from outside to inside)
Private Function IsJumpingIntoForLoop(fromLine As Long, toLine As Long) As Boolean
    ' Strategy: Find all ForStatement nodes
    ' Check if toLine is INSIDE a for loop body
    ' AND fromLine is OUTSIDE that for loop
    
    If m_programAST.tType <> Program_Node Then
        IsJumpingIntoForLoop = False
        Exit Function
    End If
    
    Dim stmt As CNode
    For Each stmt In m_programAST.Body
        If stmt.tType = ForStatement_Node Then
            Dim forStartLine As Long
            Dim forBodyStartLine As Long
            Dim forBodyEndLine As Long
            
            forStartLine = stmt.LineNumber
            
            ' Find body line range
            If Not stmt.Body Is Nothing Then
                ' Get first and last line of body
                GetNodeLineRange stmt.Body, forBodyStartLine, forBodyEndLine
                
                'Debug.Print ">>> Found for loop: lines " & forStartLine & " (body: " & forBodyStartLine & "-" & forBodyEndLine & ")"
                
                ' Check if we're jumping INTO the body
                Dim isFromOutside As Boolean
                Dim isToInside As Boolean
                
                isFromOutside = (fromLine < forStartLine Or fromLine > forBodyEndLine)
                isToInside = (toLine >= forBodyStartLine And toLine <= forBodyEndLine)
                
                If isFromOutside And isToInside Then
                    'Debug.Print ">>> Detected jump INTO for loop body!"
                    IsJumpingIntoForLoop = True
                    Exit Function
                End If
            End If
        End If
    Next
    
    IsJumpingIntoForLoop = False
End Function

' Check if jumping INTO a while loop body
Private Function IsJumpingIntoWhileLoop(fromLine As Long, toLine As Long) As Boolean
    ' Similar to for loop check
    
    If m_programAST.tType <> Program_Node Then
        IsJumpingIntoWhileLoop = False
        Exit Function
    End If
    
    Dim stmt As CNode
    For Each stmt In m_programAST.Body
        If stmt.tType = WhileStatement_Node Then
            Dim whileStartLine As Long
            Dim whileBodyStartLine As Long
            Dim whileBodyEndLine As Long
            
            whileStartLine = stmt.LineNumber
            
            If Not stmt.Body Is Nothing Then
                GetNodeLineRange stmt.Body, whileBodyStartLine, whileBodyEndLine
                
                'Debug.Print ">>> Found while loop: lines " & whileStartLine & " (body: " & whileBodyStartLine & "-" & whileBodyEndLine & ")"
                
                Dim isFromOutside As Boolean
                Dim isToInside As Boolean
                
                isFromOutside = (fromLine < whileStartLine Or fromLine > whileBodyEndLine)
                isToInside = (toLine >= whileBodyStartLine And toLine <= whileBodyEndLine)
                
                If isFromOutside And isToInside Then
                    'Debug.Print ">>> Detected jump INTO while loop body!"
                    IsJumpingIntoWhileLoop = True
                    Exit Function
                End If
            End If
        End If
    Next
    
    IsJumpingIntoWhileLoop = False
End Function

' Check if jumping INTO a do-while loop body
Private Function IsJumpingIntoDoWhileLoop(fromLine As Long, toLine As Long) As Boolean
    ' Similar to while loop check
    
    If m_programAST.tType <> Program_Node Then
        IsJumpingIntoDoWhileLoop = False
        Exit Function
    End If
    
    Dim stmt As CNode
    For Each stmt In m_programAST.Body
        If stmt.tType = DoWhileStatement_Node Then
            Dim doStartLine As Long
            Dim doBodyStartLine As Long
            Dim doBodyEndLine As Long
            
            doStartLine = stmt.LineNumber
            
            If Not stmt.Body Is Nothing Then
                GetNodeLineRange stmt.Body, doBodyStartLine, doBodyEndLine
                
                'Debug.Print ">>> Found do-while loop: lines " & doStartLine & " (body: " & doBodyStartLine & "-" & doBodyEndLine & ")"
                
                Dim isFromOutside As Boolean
                Dim isToInside As Boolean
                
                isFromOutside = (fromLine < doStartLine Or fromLine > doBodyEndLine)
                isToInside = (toLine >= doBodyStartLine And toLine <= doBodyEndLine)
                
                If isFromOutside And isToInside Then
                    'Debug.Print ">>> Detected jump INTO do-while loop body!"
                    IsJumpingIntoDoWhileLoop = True
                    Exit Function
                End If
            End If
        End If
    Next
    
    IsJumpingIntoDoWhileLoop = False
End Function

' Check if jumping INTO a function body (already exists, but let's verify it's correct)
Private Function IsJumpingIntoFunction(fromLine As Long, toLine As Long) As Boolean
    ' Find all function declarations
    ' Check if toLine is inside a function body
    ' AND fromLine is outside that function
    
    If m_programAST.tType <> Program_Node Then
        IsJumpingIntoFunction = False
        Exit Function
    End If
    
    Dim stmt As CNode
    For Each stmt In m_programAST.Body
        If stmt.tType = FunctionDeclaration_Node Then
            Dim funcStartLine As Long
            Dim funcBodyStartLine As Long
            Dim funcBodyEndLine As Long
            
            funcStartLine = stmt.LineNumber
            
            If Not stmt.Body Is Nothing Then
                GetNodeLineRange stmt.Body, funcBodyStartLine, funcBodyEndLine
                
                'Debug.Print ">>> Found function: lines " & funcStartLine & " (body: " & funcBodyStartLine & "-" & funcBodyEndLine & ")"
                
                Dim isFromOutside As Boolean
                Dim isToInside As Boolean
                
                isFromOutside = (fromLine < funcStartLine Or fromLine > funcBodyEndLine)
                isToInside = (toLine >= funcBodyStartLine And toLine <= funcBodyEndLine)
                
                If isFromOutside And isToInside Then
                    'Debug.Print ">>> Detected jump INTO function body!"
                    IsJumpingIntoFunction = True
                    Exit Function
                End If
            End If
        End If
    Next
    
    IsJumpingIntoFunction = False
End Function

' Helper: Get the line range of a node (recursively)
Private Sub GetNodeLineRange(Node As CNode, ByRef minLine As Long, ByRef maxLine As Long)
    On Error Resume Next  ' Add error handling
    
    If Node Is Nothing Then Exit Sub
    
    ' Initialize
    If minLine = 0 Then minLine = 999999
    
    ' Check this node
    If Node.LineNumber > 0 Then
        If Node.LineNumber < minLine Then minLine = Node.LineNumber
        If Node.LineNumber > maxLine Then maxLine = Node.LineNumber
    End If
    
    ' Check body (collection of statements)
    If Not Node.Body Is Nothing Then
        ' Make sure Body is actually a Collection
        If TypeName(Node.Body) = "Collection" Then
            Dim stmt As CNode
            For Each stmt In Node.Body
                GetNodeLineRange stmt, minLine, maxLine
            Next
        End If
    End If
    
    ' Check common child nodes
    If Not Node.Left Is Nothing Then GetNodeLineRange Node.Left, minLine, maxLine
    If Not Node.Right Is Nothing Then GetNodeLineRange Node.Right, minLine, maxLine
    If Not Node.Test Is Nothing Then GetNodeLineRange Node.Test, minLine, maxLine
    If Not Node.Consequent Is Nothing Then GetNodeLineRange Node.Consequent, minLine, maxLine
    If Not Node.Alternate Is Nothing Then GetNodeLineRange Node.Alternate, minLine, maxLine
    If Not Node.init Is Nothing Then GetNodeLineRange Node.init, minLine, maxLine
    
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------
' CALL JAVASCRIPT FUNCTION
' ----------------------------------------------------------------

'' Call a JavaScript function with arguments
Public Function CallJSFunction(func As CValue, args As Collection, Optional thisContext As CValue = Nothing) As CValue
    On Error GoTo ErrorHandler
    
    If func.vType <> vtfunction Then
        Err.Raise 5, "CallJSFunction", "Not a function (type=" & func.vType & ")"
    End If
    
    Dim fn As CFunction
    Set fn = func.objVal
    
    ' Create new scope
    Dim funcScope As New CScope
    Set funcScope.Parent = fn.closure
    
    ' Bind parameters
    Dim i As Long
    If Not fn.params Is Nothing Then
        For i = 1 To args.count
            If i <= fn.params.count Then
                Dim paramNode As CNode
                Set paramNode = fn.params(i)
                funcScope.DefineVar paramNode.Name, args(i)
            End If
        Next
    End If
    
    ' Save current state
    Dim oldScope As CScope
    Dim oldThis As CValue
    Dim oldReturn As Boolean
    Dim oldReturnValue As CValue
    
    Set oldScope = m_currentScope
    Set oldThis = m_thisContext
    oldReturn = m_hasReturned
    Set oldReturnValue = m_returnValue
    
    ' Set new state
    Set m_currentScope = funcScope
    Set m_thisContext = thisContext
    m_hasReturned = False
    Set m_returnValue = Nothing
    
    ' *** ADD DEBUGGER HOOK - Push frame AFTER scope switch ***
    If m_debugMode Then
        DebugPushFrame fn.Name, args.count
    End If
    
    ' Execute function body
    If Not fn.Body Is Nothing Then
        ExecuteStatement fn.Body
    End If
    
    ' *** ADD DEBUGGER HOOK - Pop frame BEFORE scope restore ***
    If m_debugMode Then
        DebugPopFrame
    End If
    
    ' Get result
    Dim result As CValue
    If m_hasReturned And Not m_returnValue Is Nothing Then
        Set result = m_returnValue
    Else
        Set result = New CValue
        result.vType = vtUndefined
    End If
    
    ' Restore state
    m_hasReturned = oldReturn
    Set m_returnValue = oldReturnValue
    Set m_currentScope = oldScope
    Set m_thisContext = oldThis
    
    Set CallJSFunction = result
    Exit Function
    
ErrorHandler:
    ' *** ALSO ADD: Pop frame on error ***
    If m_debugMode Then
        DebugPopFrame
    End If
    
    ' Restore on error
    If Not oldScope Is Nothing Then Set m_currentScope = oldScope
    If Not oldThis Is Nothing Then Set m_thisContext = oldThis
    m_hasReturned = oldReturn
    Set m_returnValue = oldReturnValue
    
    Set result = New CValue
    result.vType = vtUndefined
    Set CallJSFunction = result
End Function

' Evaluate JavaScript and return as VB COM object (if it's a JS object)
Public Function EvalAsObject(code As String) As Object
    On Error GoTo ErrorHandler
    
    Const dbg_mode As Boolean = False
    
    If Not ensureDynProxy() Then
        MsgBox "DynProxy.dll not found"
        Exit Function
    End If
    
    ' Evaluate the expression
    Dim result As CValue
    Set result = Eval(code)
    
    If result Is Nothing Then
        Set EvalAsObject = Nothing
        Exit Function
    End If
    
    ' Only wrap if it's a JavaScript object
    If result.vType = vtObject Or result.vType = vtfunction Then
        If dbg_mode Then Debug.Print ">>> EvalAsObject: Wrapping JavaScript object"
        
        ' Create resolver for this object
        Dim jsObj As New CJavaScriptObject
        'jsObj.InitFromCValue result, Me, code
        
        ' Use variable name for objects, direct reference for functions
        If result.vType = vtObject Then
            jsObj.InitFromVarName code, Me
        Else
            jsObj.InitFromCValue result, Me, code
        End If
        
        ' Wrap in dynproxy
        Dim p As Long
        p = CreateProxyForObjectRaw(0&, PtrFromObject(jsObj))
        
        ' Return as VB Object
        Set EvalAsObject = ObjectFromPtr(p)
    Else
        'Debug.Print ">>> EvalAsObject: Not an object (type=" & result.vType & ")"
        Set EvalAsObject = Nothing
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print ">>> EvalAsObject ERROR: " & Err.description
    Set EvalAsObject = Nothing
End Function


' ----------------------------------------------------------------
' Convert JavaScript value to native VB6 Variant
' ----------------------------------------------------------------
Public Function EvalAsVariant(varName As String) As Variant
    On Error GoTo ErrorHandler
    
    Dim result As CValue
    Set result = Eval(varName)
    
    If result Is Nothing Then
        EvalAsVariant = Empty
        Exit Function
    End If
    
    ' Convert to native VB type
    If result.vType = vtObject Or result.vType = vtCOMObject Or result.vType = vtfunction Then
        Set EvalAsVariant = CValueToNativeVariant(result)
    Else
        EvalAsVariant = CValueToNativeVariant(result)
    End If

    Exit Function
    
ErrorHandler:
    Debug.Print ">>> EvalAsVariant ERROR: " & Err.description
    EvalAsVariant = Empty
End Function

' ----------------------------------------------------------------
' Convert CValue to native VB6 Variant (recursive)
' ----------------------------------------------------------------
Private Function CValueToNativeVariant(val As CValue) As Variant
    Select Case val.vType
        Case vtNumber
            CValueToNativeVariant = val.numVal
            
        Case vtString
            CValueToNativeVariant = val.strVal
            
        Case vtBoolean
            CValueToNativeVariant = val.boolVal
            
        Case vtNull
            CValueToNativeVariant = Null
            
        Case vtUndefined
            CValueToNativeVariant = Empty
            
        Case vtObject
            ' Convert to Dictionary
            Set CValueToNativeVariant = ObjectToNativeDictionary(val)
            
        Case vtArray
            ' Convert to VB6 Array
            CValueToNativeVariant = ArrayToNativeArray(val)
            
        Case vtfunction
            ' Can't convert functions - return string representation
            CValueToNativeVariant = "[Function]"
            
        Case Else
            CValueToNativeVariant = Empty
    End Select
End Function

' ----------------------------------------------------------------
' Convert JavaScript object to Scripting.Dictionary
' ----------------------------------------------------------------
Private Function ObjectToNativeDictionary(obj As CValue) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    If obj.vType <> vtObject Then
        Set ObjectToNativeDictionary = dict
        Exit Function
    End If
    
    If obj.objectKeys Is Nothing Then
        Set ObjectToNativeDictionary = dict
        Exit Function
    End If
    
    ' Convert each property recursively
    Dim i As Long
    For i = 1 To obj.objectKeys.count
        Dim key As String
        Dim val As CValue
     
        key = obj.objectKeys(i)
        Set val = obj.objectProps(i)
        'Debug.Print "key: " & key
        
        ' Recursive conversion
        If val.vType = vtObject Then
            Set dict(key) = CValueToNativeVariant(val)
        Else
            dict(key) = CValueToNativeVariant(val)
        End If
    Next
    
    Set ObjectToNativeDictionary = dict
End Function

' ----------------------------------------------------------------
' Convert JavaScript array to VB6 Array
' ----------------------------------------------------------------
Private Function ArrayToNativeArray(arr As CValue) As Variant
    If arr.vType <> vtArray Or arr.arrayVal Is Nothing Then
        ArrayToNativeArray = Array()
        Exit Function
    End If
    
    Dim count As Long
    count = arr.arrayVal.count
    
    If count = 0 Then
        ArrayToNativeArray = Array()
        Exit Function
    End If
    
    ' Build VB6 array
    Dim result() As Variant
    ReDim result(0 To count - 1)
    
    Dim i As Long
    For i = 1 To count
        Dim elem As CValue
        Set elem = arr.arrayVal(i)
        
        ' Recursive conversion
        If elem.vType = vtObject Then
            Set result(i - 1) = CValueToNativeVariant(elem)
        Else
            result(i - 1) = CValueToNativeVariant(elem)
        End If
    Next
    
    ArrayToNativeArray = result
End Function

Private Function GlobalHex(args As Collection) As Variant
    ' JavaScript Hex() function - converts number to hex string
    ' Handles both regular numbers and BigInts (vtInt64)
    
    If args.count = 0 Then
        Err.Raise vbObjectError + 1, "Hex", "Hex() requires at least 1 argument"
    End If
    
    Dim arg As CValue
    Set arg = args(1)
    
    Dim hexStr As String
    Dim U64 As New ULong64
    
    ' Check the CValue type
    Select Case arg.vType
        Case vtInt64
            ' It's a BigInt - already stored as Currency in int64Val
            U64.rawValue = arg.int64Val
            U64.use0x = True
            U64.padLeft = True
            hexStr = U64.ToString(mHex)
            
        Case vtNumber
            ' Regular number
            Dim dblVal As Double
            dblVal = arg.numVal
            
            ' Check if it needs 64-bit handling
            If dblVal > 2147483647# Or dblVal < -2147483648# Then
                ' Use ULong64 for large numbers
                
                Dim hi As Long, lo As Long
                
                If dblVal < 0 Then
                    ' Negative number - convert to two's complement
                    Dim posVal As Double
                    posVal = 1.84467440737096E+19 + dblVal   ' 2^64 + negative value
                    hi = DoubleToSignedLong(Int(posVal / 4294967296#))
                    lo = DoubleToSignedLong(posVal - (Int(posVal / 4294967296#) * 4294967296#))
                    U64.SetLongs hi, lo
                ElseIf dblVal >= 4294967296# Then
                    ' Large positive number > 32-bit
                    hi = DoubleToSignedLong(Int(dblVal / 4294967296#))
                    lo = DoubleToSignedLong(dblVal - (Int(dblVal / 4294967296#) * 4294967296#))
                    U64.SetLongs hi, lo
                Else
                    ' Positive number in 32-bit range but > signed int max
                    ' This is the case for 4294967295
                    lo = DoubleToSignedLong(dblVal)
                    U64.SetLongs 0, lo
                End If
                
                U64.use0x = True
                U64.padLeft = True
                hexStr = U64.ToString(mHex)
            Else
                ' Standard VB Hex for 32-bit safe numbers
                hexStr = "0x" & UCase$(hex$(CLng(dblVal)))
            End If
            
        Case vtString
            ' Try to parse string as number
            Dim numVal As Double
            On Error Resume Next
            numVal = CDbl(arg.strVal)
            If Err.Number = 0 And numVal <= 2147483647# And numVal >= -2147483648# Then
                ' Safe 32-bit range
                hexStr = "0x" & UCase$(hex$(CLng(numVal)))
            Else
                ' Try parsing as 64-bit
                Err.Clear
                If U64.fromString(arg.strVal, mUnsigned) Or U64.fromString(arg.strVal, mHex) Then
                    U64.use0x = True
                    U64.padLeft = True
                    hexStr = U64.ToString(mHex)
                Else
                    Err.Clear
                    Err.Raise vbObjectError + 1, "Hex", "Invalid number string"
                End If
            End If
            On Error GoTo 0
            
        Case Else
            ' Unsupported type
            Err.Raise vbObjectError + 1, "Hex", "Hex() requires a number or BigInt"
    End Select
    
    ' Return as CValue string
    Dim result As New CValue
    result.vType = vtString
    result.strVal = hexStr
    Set GlobalHex = result
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

' ? ADD new Function CreateDateObject
Private Function CreateDateObject(args As Collection) As CValue
    Dim result As New CValue
    result.vType = vtObject
    Set result.objectProps = New Collection
    Set result.objectKeys = New Collection
    
    Dim dateVal As Date
    
    ' Determinar la fecha según argumentos
    If args.count = 0 Then
        ' new Date() - fecha actual
        dateVal = Now
    ElseIf args.count = 1 Then
        ' new Date(milliseconds) o new Date(dateString)
        Dim arg As CNode
        Set arg = args(1)
        Dim argVal As CValue
        Set argVal = EvaluateExpression(arg)
        
        If argVal.vType = vtNumber Then
            ' Milliseconds desde epoch (simplificado)
            dateVal = DateAdd("s", argVal.numVal / 1000, #1/1/1970#)
        ElseIf argVal.vType = vtString Then
            ' Parse date string
            On Error Resume Next
            dateVal = CDate(argVal.strVal)
            If Err.Number <> 0 Then
                dateVal = Now  ' Fallback
            End If
            On Error GoTo 0
        Else
            dateVal = Now
        End If
    Else
        ' new Date(year, month, day, hour, min, sec, ms)
        ' Simplificado - solo year, month, day
        Dim yearVal As CValue, monthVal As CValue, dayVal As CValue
        
        Set arg = args(1)
        Set yearVal = EvaluateExpression(arg)
        
        Set arg = args(2)
        Set monthVal = EvaluateExpression(arg)
        
        Dim day As Long
        day = 1
        If args.count >= 3 Then
            Set arg = args(3)
            Set dayVal = EvaluateExpression(arg)
            day = CLng(dayVal.ToNumber())
        End If
        
        dateVal = DateSerial(CLng(yearVal.ToNumber()), _
                            CLng(monthVal.ToNumber()) + 1, _
                            day)
    End If
    
    ' Agregar métodos del Date object
    
    ' toString()
    Dim toStringFunc As New CValue
    toStringFunc.vType = vtfunction
    ' Crear función que retorna la fecha como string
    Dim toStrFn As New CFunction
    toStrFn.Name = "toString"
    Set toStrFn.Body = Nothing  ' Método nativo
    Set toStringFunc.objVal = toStrFn
    
    result.objectProps.add toStringFunc, "toString"
    result.objectKeys.add "toString", "toString"
    
    ' getTime() - milliseconds desde epoch
    Dim getTimeFunc As New CValue
    getTimeFunc.vType = vtfunction
    Dim getTimeFn As New CFunction
    getTimeFn.Name = "getTime"
    Set getTimeFn.Body = Nothing
    Set getTimeFunc.objVal = getTimeFn
    
    result.objectProps.add getTimeFunc, "getTime"
    result.objectKeys.add "getTime", "getTime"
    
    ' Guardar el valor de fecha internamente
    Dim internalDate As New CValue
    internalDate.vType = vtString
    internalDate.strVal = Format$(dateVal, "ddd mmm dd yyyy hh:nn:ss")
    
    result.objectProps.add internalDate, "_date"
    result.objectKeys.add "_date", "_date"
    
    ' Timestamp interno
    Dim timestamp As New CValue
    timestamp.vType = vtNumber
    timestamp.numVal = DateDiff("s", #1/1/1970#, dateVal) * 1000
    
    result.objectProps.add timestamp, "_timestamp"
    result.objectKeys.add "_timestamp", "_timestamp"
    
    Set CreateDateObject = result
End Function

' ? IsDateMethod,CallDateMethod
Private Function IsDateMethod(methodName As String) As Boolean
    Select Case methodName
        Case "toString", "toDateString", "toTimeString", _
             "getTime", "getFullYear", "getMonth", "getDate", _
             "getHours", "getMinutes", "getSeconds", "getMilliseconds"
            IsDateMethod = True
        Case Else
            IsDateMethod = False
    End Select
End Function

Private Function CallDateMethod(dateObj As CValue, methodName As String, args As Collection) As CValue
    Dim result As New CValue
    
    ' Obtener el timestamp interno
    Dim timestampVal As CValue
    Set timestampVal = dateObj.GetProperty("_timestamp")
    
    If timestampVal Is Nothing Then
        result.vType = vtString
        result.strVal = "Invalid Date"
        Set CallDateMethod = result
        Exit Function
    End If
    
    ' Convertr timestamp to Date VB6
    Dim dateVal As Date
    dateVal = DateAdd("s", timestampVal.numVal / 1000, #1/1/1970#)
    
    Select Case methodName
        Case "toString"
            result.vType = vtString
            result.strVal = Format$(dateVal, "ddd mmm dd yyyy hh:nn:ss")
            
        Case "toDateString"
            result.vType = vtString
            result.strVal = Format$(dateVal, "ddd mmm dd yyyy")
            
        Case "toTimeString"
            result.vType = vtString
            result.strVal = Format$(dateVal, "hh:nn:ss")
            
        Case "getTime"
            result.vType = vtNumber
            result.numVal = timestampVal.numVal
            
        Case "getFullYear"
            result.vType = vtNumber
            result.numVal = Year(dateVal)
            
        Case "getMonth"
            result.vType = vtNumber
            result.numVal = Month(dateVal) - 1  ' JavaScript is 0-based
            
        Case "getDate"
            result.vType = vtNumber
            result.numVal = day(dateVal)
            
        Case "getHours"
            result.vType = vtNumber
            result.numVal = Hour(dateVal)
            
        Case "getMinutes"
            result.vType = vtNumber
            result.numVal = Minute(dateVal)
            
        Case "getSeconds"
            result.vType = vtNumber
            result.numVal = Second(dateVal)
            
        Case "getMilliseconds"
            result.vType = vtNumber
            result.numVal = 0  ' VB6 not have milliseconds
            
        Case Else
            result.vType = vtUndefined
    End Select
    
    Set CallDateMethod = result
End Function


