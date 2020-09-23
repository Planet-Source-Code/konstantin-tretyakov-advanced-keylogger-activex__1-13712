VERSION 5.00
Object = "{C4D77E92-252D-11D4-B358-C9A9F1AA7152}#1.0#0"; "KBDMONITOR.OCX"
Begin VB.Form Form1 
   Caption         =   "KeyMon Test"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin KbdMonitor.KeyMon KeyMon1 
      Left            =   3960
      Top             =   120
      _ExtentX        =   979
      _ExtentY        =   979
      Active          =   0
      Passthru        =   -1
      AcceptRepeats   =   0   'False
      DifferNumpad    =   0   'False
      LogType         =   0
      MacroLogType    =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   3600
      Top             =   120
   End
   Begin VB.CheckBox chkActive 
      Caption         =   "&Active"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Value           =   1  'Checked
      Width           =   3615
   End
   Begin VB.TextBox txtLog 
      Enabled         =   0   'False
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   720
      Width           =   4455
   End
   Begin VB.CheckBox chkPass 
      Caption         =   "&Passthru"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Value           =   1  'Checked
      Width           =   3615
   End
   Begin VB.Label Label1 
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'Test program for the KeyMon control
'Try changing the control's properties and see the result
'Also note, that if run the program and click on the Passthru checkbox, to de-check it, then
'the control will suddenly "stop functioning" (logging)
'That is a VB bug. Just compile the program, perform the same actions and see the difference
'Copyright© 2000 Konstantin Tretyakov
Private Sub chkPass_Click()
    KeyMon1.Passthru = chkPass.Value
End Sub
Private Sub chkActive_Click()
    KeyMon1.Active = chkActive.Value
End Sub

Private Sub Form_Load()
    KeyMon1.Activate
    KeyMon1.Macro("Macro1") = "{Shift}{Ctrl}{Ctrl} A"
    KeyMon1.Macro("Macro2") = "{Tab}{Tab}123"
    Label1.Caption = "Try typing the following key sequences: " + vbCrLf + vbLf + "Shift Ctrl Ctrl Space A" + vbCrLf + "Tab Tab 1 2 3"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    KeyMon1.Deactivate      'Always deactivate the control upon exit
End Sub

Private Sub KeyMon1_KeyDownEx(KeyText As String)
    txtLog = txtLog + KeyText
End Sub

Private Sub KeyMon1_Macro(MacroKey As String)
    MsgBox MacroKey
End Sub

Private Sub Timer1_Timer() 'This sub is for debugging purposes
    chkActive.Value = IIf(KeyMon1.Active, 1, 0)
    chkPass.Value = IIf(KeyMon1.Passthru, 1, 0)
End Sub
