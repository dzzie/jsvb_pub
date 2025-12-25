VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Basic Script Examples"
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   705
   ClientWidth     =   11040
   LinkTopic       =   "Form1"
   ScaleHeight     =   8670
   ScaleWidth      =   11040
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPaste 
      Caption         =   "Paste"
      Height          =   495
      Left            =   5280
      TabIndex        =   7
      Top             =   6000
      Width           =   975
   End
   Begin VB.CheckBox chkSafeSubset 
      Caption         =   "Safe Subset"
      Height          =   375
      Left            =   2655
      TabIndex        =   6
      Top             =   6030
      Width           =   1725
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   1575
      Top             =   5940
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAbort 
      Caption         =   "Abort"
      Height          =   465
      Left            =   7920
      TabIndex        =   5
      Top             =   5985
      Width           =   1365
   End
   Begin VB.TextBox txtOut 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2085
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   6525
      Width           =   10905
   End
   Begin VB.CommandButton cmdExec 
      Caption         =   "Execute"
      Height          =   465
      Left            =   9420
      TabIndex        =   1
      Top             =   5985
      Width           =   1455
   End
   Begin VB.TextBox txtScript 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5550
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   315
      Width           =   10905
   End
   Begin VB.Label Label1 
      Caption         =   "Output"
      Height          =   285
      Index           =   1
      Left            =   45
      TabIndex        =   4
      Top             =   6165
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "Script"
      Height          =   285
      Index           =   0
      Left            =   45
      TabIndex        =   3
      Top             =   45
      Width           =   915
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save As"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents intrep As CInterpreter
Attribute intrep.VB_VarHelpID = -1
Dim last As String

Private Sub cmdAbort_Click()
    If Not intrep Is Nothing Then intrep.Abort
End Sub

Private Sub cmdPaste_Click()
    txtScript = unixToDOS(Clipboard.GetText)
End Sub

Private Sub intrep_AuditEvent(category As js4vb.enumAuditEvents, description As String, ByRef cancel As Boolean)
    
    Debug.Print "AuditEvent: " & intrep.AuditEventToStr(category) & " : " & description
    
    'If category = aeActiveX Then cancel = True 'works
    
End Sub

Private Sub intrep_ConsoleLog(ByVal msg As String)
    txtOut = txtOut & msg & vbCrLf
End Sub

Private Sub intrep_OnError(ByVal ErrorMessage As String, ByVal LineNumber As Long, ByVal source As String, ByVal col As Long)
    txtOut = txtOut & "Error Line " & LineNumber & ": " & ErrorMessage & vbCrLf
End Sub
 


Private Sub cmdExec_Click()
    On Error Resume Next
    
    Me.Caption = "Running"
    txtOut.Text = Empty
    Set intrep = New CInterpreter
    
    WriteFile last, txtScript.Text
    
    intrep.AuditMode = True
    intrep.UseSafeSubset = (chkSafeSubset.Value = 1)
    intrep.AddCOMObject "form", Me
    intrep.AttachCOMHelper "form", "taco", "function(cmd){print('its tuesday ' + cmd + '!');}"
    
    intrep.Execute txtScript.Text
    
    If Err.Number <> 0 Then MsgBox Err.description
    Me.Caption = "Stopped"
    
End Sub

Private Sub Form_Load()
         
    On Error Resume Next
    
    last = App.path & "\lastScript.txt"
    If FileExists(last) Then
        txtScript.Text = ReadFile(last)
    Else
        txtScript = "console.log(form.caption)" & vbCrLf & _
                "Form.Caption = 'test from vb!'"
    End If

End Sub

Function t(data)
    txtOut.Text = txtOut.Text & data & vbCrLf
    txtOut.SelStart = Len(txtOut.Text)
    DoEvents
End Function

Private Sub Form_Unload(cancel As Integer)
    On Error Resume Next
    WriteFile last, txtScript.Text
End Sub

Function FileExists(path As String) As Boolean
  On Error GoTo hell
    
  If Len(path) = 0 Then Exit Function
  If Right(path, 1) = "\" Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
  
  Exit Function
hell: FileExists = False
End Function

Function ReadFile(filename)
  Dim f As Long, temp
  f = FreeFile
  temp = ""
   Open filename For Binary As #f        ' Open file.(can be text or image)
     temp = Input(FileLen(filename), #f) ' Get entire Files data
   Close #f
   ReadFile = temp
End Function

Sub WriteFile(path, it)
    Dim f As Long
    f = FreeFile
    Open path For Output As #f
    Print #f, it
    Close f
End Sub

Function unixToDOS(ByVal tmp As String)
    Dim isMixed As Boolean
    isMixed = (InStr(tmp, vbCrLf) > 0)
    If isMixed Then tmp = VBA.Replace(tmp, vbCrLf, Chr(5))
    tmp = VBA.Replace(tmp, vbLf, vbCrLf)
    If isMixed Then tmp = VBA.Replace(tmp, Chr(5), vbCrLf)
    unixToDOS = tmp
End Function



Private Sub mnuOpen_Click()
    Dim f As String
    dlg.InitDir = App.path
    dlg.ShowOpen
    f = dlg.filename
    If FileExists(f) Then txtScript = ReadFile(f)
End Sub

Private Sub mnuSaveAs_Click()
    Dim f As String
    dlg.InitDir = App.path
    dlg.ShowSave
    f = dlg.filename
    WriteFile f, txtScript
End Sub
