VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AGENTCTL.DLL"
Begin VB.Form Form1 
   BackColor       =   &H00008000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   2040
      Top             =   3600
      _cx             =   847
      _cy             =   847
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Starting Up UOS Hit Control to continue"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1560
      TabIndex        =   1
      Top             =   2280
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to UOS 2003"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   11895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If vbKeyControl Then
Form2.Show
End If
End Sub

Private Sub Form_Load()
MsgBox "Your Screen Resolution Must Be at 800 by 600.  If not we suggest to exit out and change it."
Anim = "Wizard"
Agent1.Characters.Load Anim, Anim & ".acs"
Set Char = Agent1.Characters(Anim)
Char.Show
Char.MoveTo 300, 300
Char.Speak ("What up and welcome to UOS StartUp!")
Char.Play "Explain"
Char.Speak ("Hi!  I'll introduce myself, but not now.")
Char.Speak ("To continue please hit the control button on your keyboard.")
Char.Speak ("The control button actually says on your keyboard, CTrl.")
Char.Speak ("Well go on!  But hide me before you hit CTrl.  Right click on me to hide me.  I'll appear on the Desktop, in which if you hit CTrl you'll be headed.")
Char.Play "DoMagic1"
Char.Play "DoMagic2"
End Sub

Private Sub Timer1_Timer()
Label3.Caption = "Starting Up: All APPs"
ProgressBar1.Value = 10
Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
Label3.Caption = "Starting Up: Everything Else"
ProgressBar1.Value = 40
Timer1.Enabled = False
End Sub

Private Sub Timer3_Timer()
Label3.Caption = "Starting Up: Everything Else"
ProgressBar1.Value = 60
Timer1.Enabled = False
End Sub

Private Sub Timer4_Timer()
Label3.Caption = "Starting Up: Everything Else"
ProgressBar1.Value = 80
Timer1.Enabled = False
End Sub

Private Sub Timer5_Timer()
Label3.Caption = "Starting Up: Exiting Current OS"
ProgressBar1.Value = 100
Timer1.Enabled = False
Form2.Show
Unload Me
End Sub
