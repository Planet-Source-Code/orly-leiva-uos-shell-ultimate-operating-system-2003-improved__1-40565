VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AGENTCTL.DLL"
Begin VB.Form Form3 
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   3525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   3525
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Scan Your PC for Viruses"
      Height          =   495
      Left            =   360
      TabIndex        =   13
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "About UOS"
      Height          =   495
      Left            =   1200
      TabIndex        =   12
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New User?  Clik here!"
      Height          =   495
      Left            =   0
      TabIndex        =   11
      Top             =   2520
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   360
      Picture         =   "Settings.frx":0000
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   2
      Top             =   1440
      Width           =   495
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   3000
      Top             =   2880
      _cx             =   847
      _cy             =   847
   End
   Begin VB.Image imgAddRemove 
      Height          =   480
      Left            =   480
      Picture         =   "Settings.frx":030A
      Top             =   480
      Width           =   480
   End
   Begin VB.Label lblAddRemove 
      BackStyle       =   0  'Transparent
      Caption         =   "Add / Remove - Programs"
      Height          =   735
      Left            =   240
      TabIndex        =   10
      Top             =   960
      Width           =   1095
   End
   Begin VB.Image imgAddHardware 
      Height          =   480
      Left            =   1800
      Picture         =   "Settings.frx":074C
      Top             =   360
      Width           =   480
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Add New Hardware"
      Height          =   735
      Left            =   1560
      TabIndex        =   9
      Top             =   960
      Width           =   1095
   End
   Begin VB.Image imgDisplayProperties 
      Height          =   480
      Left            =   2760
      Picture         =   "Settings.frx":0B8E
      Top             =   480
      Width           =   480
   End
   Begin VB.Label lblDisplayProperties 
      BackStyle       =   0  'Transparent
      Caption         =   "Display Properties"
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   960
      Width           =   855
   End
   Begin VB.Image imgKeyboardProperties 
      Height          =   480
      Left            =   3960
      Picture         =   "Settings.frx":0FD0
      Top             =   480
      Width           =   480
   End
   Begin VB.Label lblKeyboardProperties 
      BackStyle       =   0  'Transparent
      Caption         =   "Keyboard Properties"
      Height          =   495
      Left            =   3840
      TabIndex        =   7
      Top             =   960
      Width           =   855
   End
   Begin VB.Image imgNetworkProperties 
      Height          =   480
      Left            =   1560
      Picture         =   "Settings.frx":1412
      Top             =   1440
      Width           =   480
   End
   Begin VB.Label lblNetworkProperties 
      BackStyle       =   0  'Transparent
      Caption         =   "Network Properties"
      Height          =   495
      Left            =   1440
      TabIndex        =   6
      Top             =   1920
      Width           =   735
   End
   Begin VB.Image imgSystemProperties 
      Height          =   480
      Left            =   2640
      Picture         =   "Settings.frx":1854
      Top             =   1440
      Width           =   480
   End
   Begin VB.Label lblSysProperties 
      BackStyle       =   0  'Transparent
      Caption         =   "System Properties"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   1920
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3720
      Picture         =   "Settings.frx":1C96
      Top             =   1440
      Width           =   480
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Desktop Settings"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "UOS Helper Settings"
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      Height          =   255
      Left            =   4200
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "Settings"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form6.Show
End Sub


Private Sub Command2_Click()
Form7.Show
End Sub

Private Sub Command3_Click()
frmScan.Show
End Sub

Private Sub Form_Load()
If Form2.Task1.Caption = "" Then
Form2.Task1.Visible = True
Form2.Task1.Caption = "Settings"
ElseIf Form2.Task2.Caption = "" Then
Form2.Task2.Visible = True
Form2.Task2.Caption = "Settings"
ElseIf Form2.Task3.Caption = "" Then
Form2.Task3.Visible = True
Form2.Task3.Caption = "Settings"
ElseIf Form2.Task4.Caption = "" Then
Form2.Task4.Visible = True
Form2.Task4.Caption = "Settings"
ElseIf Form2.Task5.Caption = "" Then
Form2.Task5.Visible = True
Form2.Task5.Caption = "Settings"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Form2.Task1.Caption = "Settings" Then
Form2.Task1.Visible = False
Form2.Task1.Caption = ""
ElseIf Form2.Task2.Caption = "Settings" Then
Form2.Task2.Visible = False
Form2.Task2.Caption = ""
ElseIf Form2.Task3.Caption = "Settings" Then
Form2.Task3.Visible = False
Form2.Task3.Caption = ""
ElseIf Form2.Task4.Caption = "Settings" Then
Form2.Task4.Visible = False
Form2.Task4.Caption = ""
ElseIf Form2.Task5.Caption = "Settings" Then
Form2.Task5.Visible = False
Form2.Task5.Caption = ""
End If
End Sub

Private Sub Image1_Click()
Form4.Show
End Sub

Private Sub imgAddHardware_Click()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl @1", 5)
End Sub

Private Sub imgAddRemove_Click()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL appwiz.cpl,,1", 5)
End Sub

Private Sub imgDisplayProperties_Click()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,0", 5)
End Sub

Private Sub imgKeyboardProperties_Click()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL main.cpl @1", 5)
End Sub

Private Sub imgNetworkProperties_Click()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL netcpl.cpl", 5)
End Sub

Private Sub imgSystemProperties_Click()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl,,0", 5)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormMove Me
End Sub

Private Sub Label2_Click()
Unload Me
End Sub

Private Sub Picture1_Click()
Agent1.ShowDefaultCharacterProperties
End Sub
