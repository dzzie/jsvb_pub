VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   ScaleHeight     =   6270
   ScaleWidth      =   8100
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   315
      Left            =   6000
      TabIndex        =   2
      Top             =   5760
      Width           =   1275
   End
   Begin MSComctlLib.ListView lv 
      Height          =   1695
      Left            =   540
      TabIndex        =   1
      Top             =   3960
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   2990
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.ListBox List1 
      Height          =   3570
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   8055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)

Private Declare Sub IPCDebugMode Lib "dynproxy.dll" (ByVal enabled As Long)
Private Declare Function SinkEventsDisconnect Lib "dynproxy.dll" (ByVal pHandle As Long) As Long
Private Declare Sub SendDbgMsg Lib "dynproxy.dll" (ByVal msg As String)

Private Declare Function SinkEventsAuto Lib "dynproxy.dll" ( _
    ByVal pSourceRaw As Long, _
    ByVal sourceName As Long, _
    ByVal pCallbackRaw As Long, _
    ByRef ppHandle As Long) As Long


Dim hSink As Long
Dim oCallback As clsEventCallback
Dim hr As Long

Dim test As New clsEventSource

Function d(x)
    Debug.Print x
    List1.AddItem x
End Function

 
Private Sub Form_Load()
    
    Dim hLib As Long
    Dim pth As String
    
    
    Set oCallback = New clsEventCallback
    
    pth = "D:\_code\js4vb\dynproxy\Debug\dynproxy.dll"
    pth = App.Path & "\..\..\dynproxy.dll"
    
    hLib = LoadLibrary(pth)
    
    If hLib = 0 Then
        d "Failed to find dynproxy.dll?"
        Exit Sub
    End If
 
    IPCDebugMode 1
    SendDbgMsg "<cls>"
    SendDbgMsg "Starting: " & Now
    
    'TestSink 'works
    'lvTest   'works
    
    'this does not work..vb intrinsic controls are not full COM objects
    'they do not support enumconnection points, we extracted their event class IIDs and tested
    'but it looks like they expect early bound vtable based call backs and not IDispatch based lies
    'like we tell the others.
    'hr = SinkEventsAuto(ObjPtr(Command1), StrPtr("cmd1"), ObjPtr(oCallback), hSink)
    'Form1.List1.AddItem "Command1 Result: 0x" & Hex$(hr) & " handle=" & hSink

    Set test.t = Command1

End Sub

'working
Function lvTest()
    lv.ListItems.Add , , "Item 1"
    lv.ListItems.Add , , "Item 2"
    
    d "Hooking ListView events..."
    hr = SinkEventsAuto(ObjPtr(lv.Object), StrPtr("lv"), ObjPtr(oCallback), hSink)
    d "SinkEventsAuto: 0x" & Hex$(hr) & " handle=" & hSink
End Function

Public Function PtrFromObject(obj As Object) As Long
    Dim unk As IUnknown
    Set unk = obj
    PtrFromObject = ObjPtr(unk)
End Function

Sub TestSink()
    Dim oSource As New clsEventSource
    Dim hSink As Long
    Dim hr As Long

    Form1.List1.AddItem "Calling SinkEventsAuto..."
    hr = SinkEventsAuto(PtrFromObject(oSource), StrPtr("mySource"), PtrFromObject(oCallback), hSink)
    Form1.List1.AddItem "Result: 0x" & Hex$(hr) & " handle=" & hSink
    oSource.FireTestEvent "test", 12
    
End Sub


'Public Sub TestWithVB6Class()
'    Dim hr As Long
'    Dim oSource As clsEventSource
'    Dim oCallback As clsEventCallback
'    Dim hSink As Long
'
'    Set oCallback = New clsEventCallback
'    Set oSource = New clsEventSource
'
'    Form1.List1.AddItem "Calling SinkEventsAuto..."
'    hr = SinkEventsAuto(oSource, StrPtr("mySource"), oCallback, hSink)
'
'    Form1.List1.AddItem "SinkEventsAuto returned: 0x" & Hex$(hr)
'
'    If hr < 0 Then
'        Form1.List1.AddItem "Failed - check IPC debug output"
'        Exit Sub
'    End If
'
'    Form1.List1.AddItem "Firing events..."
'    oSource.FireTestEvent "Hello", 42
'    oSource.FireAnotherEvent 3.14159
'
'    SinkEventsDisconnect hSink
'    Form1.List1.AddItem "Done"
'End Sub



