VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AGENTCTL.DLL"
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form2"
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picSystray 
      Height          =   495
      Left            =   10320
      ScaleHeight     =   435
      ScaleWidth      =   915
      TabIndex        =   24
      Top             =   960
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Left            =   8160
      Top             =   360
   End
   Begin VB.CommandButton Task5 
      Height          =   375
      Left            =   7320
      TabIndex        =   22
      Top             =   8520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Task4 
      Height          =   375
      Left            =   6120
      TabIndex        =   21
      Top             =   8520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Task3 
      Height          =   375
      Left            =   4920
      TabIndex        =   20
      Top             =   8520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Task2 
      Height          =   375
      Left            =   3720
      TabIndex        =   19
      Top             =   8520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Task1 
      Height          =   375
      Left            =   2640
      TabIndex        =   18
      Top             =   8520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "No Menu"
      Height          =   375
      Left            =   1320
      TabIndex        =   16
      Top             =   8520
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Menu"
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   8520
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1920
      Picture         =   "DeskTop.frx":0000
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   4
      Top             =   360
      Width           =   615
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   480
      Picture         =   "DeskTop.frx":0442
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   3
      Top             =   1440
      Width           =   492
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2040
      Picture         =   "DeskTop.frx":0884
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   2
      Top             =   1440
      Width           =   495
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   600
      Picture         =   "DeskTop.frx":0CC6
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   360
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2775
      Left            =   -120
      TabIndex        =   8
      Top             =   5640
      Visible         =   0   'False
      Width           =   1695
      Begin VB.CommandButton Command1 
         Caption         =   "Programs"
         Height          =   375
         Left            =   0
         TabIndex        =   13
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Power Out"
         Height          =   375
         Left            =   0
         TabIndex        =   12
         Top             =   2280
         Width           =   1575
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Run"
         Height          =   375
         Left            =   0
         TabIndex        =   11
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Find"
         Height          =   375
         Left            =   0
         TabIndex        =   10
         Top             =   1560
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Settings"
         Height          =   375
         Left            =   0
         TabIndex        =   9
         Top             =   1200
         Width           =   1575
      End
   End
   Begin VB.Shape Shape3 
      Height          =   2895
      Left            =   0
      Top             =   5520
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblTime 
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      TabIndex        =   23
      Top             =   8520
      Width           =   1815
   End
   Begin VB.Shape Shape12 
      Height          =   495
      Left            =   1200
      Top             =   8400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Shape Shape11 
      Height          =   495
      Left            =   120
      Top             =   8400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Shape10 
      Height          =   375
      Left            =   1800
      Top             =   7560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Shape6 
      Height          =   375
      Left            =   1800
      Top             =   7920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Shape5 
      Height          =   375
      Left            =   1800
      Top             =   7200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Shape4 
      Height          =   375
      Left            =   1800
      Top             =   6840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      Height          =   375
      Left            =   1800
      Top             =   6480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "User: "
      Height          =   255
      Left            =   1920
      TabIndex        =   17
      Top             =   0
      Width           =   3375
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C00000&
      Height          =   615
      Left            =   0
      TabIndex        =   14
      Top             =   8400
      Width           =   12015
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "PC Navigator/ I Browser"
      ForeColor       =   &H80000007&
      Height          =   495
      Left            =   1680
      TabIndex        =   7
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Settings"
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Trash Bin"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   2040
      Width           =   735
   End
   Begin VB.Shape Shape7 
      Height          =   1095
      Left            =   1560
      Top             =   240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Shape Shape8 
      Height          =   1095
      Left            =   120
      Top             =   1320
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Shape Shape9 
      Height          =   1095
      Left            =   1560
      Top             =   1320
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      Height          =   1095
      Left            =   120
      Top             =   240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "UOS Stuff"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   9480
      Top             =   240
      _cx             =   847
      _cy             =   847
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Agent1_Click(ByVal CharacterID As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
If Button = 1 Then
Anim = "Wizard"
Set Char = Agent1.Characters(Anim)
Char.Play ("Explain")
Char.Speak ("Well It seems you need help getting around UOS.")
Pause (5)
Shape1.Visible = True
Char.Speak ("If you click here, you'll go to the command center of UOS.  There you can explore your harddrive, folders and settings.")
Pause (10)
Shape1.Visible = False
Shape7.Visible = True
Char.Speak ("Click here to access I Browser and PC Navigator.  With I Browser you can access the internet with PC Navigator you can explore your own C Drive (Or harddrive if you call it that)")
Pause (10)
Shape7.Visible = False
Shape8.Visible = True
Char.Speak ("Click here to access you Computer Settings and Options.  Also you can access UOS Settings and Options.")
Pause (10)
Shape8.Visible = False
Shape9.Visible = True
Char.Speak ("Click here to go to the trash bin.  There you'll find your deleted files.  To continue please open up the menu.  Click on the Menu button to do so.")
Pause (20)
Shape9.Visible = False
Shape2.Visible = True
Char.Speak ("Click here to run Windows Based Programs. ")
Pause (10)
Shape2.Visible = False
Shape3.Visible = True
Char.Speak ("This is the Menu.")
Pause (5)
Shape3.Visible = False
Shape4.Visible = True
Char.Speak ("Click here to access Settings.")
Pause (5)
Shape4.Visible = False
Shape5.Visible = True
Char.Speak ("Click here to Search For Files in you Computer.")
Pause (5)
Shape5.Visible = False
Shape10.Visible = True
Char.Speak ("Click here to Run a Program.")
Pause (5)
Shape10.Visible = False
Shape6.Visible = True
Char.Speak ("Click here to Power off the computer or return to Windows.")
Pause (5)
Shape6.Visible = False
Shape11.Visible = True
Char.Speak ("Click here to make the Menu show up.")
Pause (5)
Shape11.Visible = False
Shape12.Visible = True
Char.Speak ("Click here to hide Menu.")
Pause (5)
Shape12.Visible = False
Char.Speak ("Well, you know everything there is to know on how to use UOS.  Oopss!  I forgot to mension the Stuff Bar.  It is the long Blue bar on the bottom of the screen.")
Char.Speak ("Click on me again to repeat my help.  Right Click on me to hide me.  Also if you don't want to hide me but want me to stay out of your way please click and drag me anywhere on the screen.  Well have Fun!!!!")
End If
End Sub

Private Sub Command1_Click()
frmBrowser.Show
frmBrowser.brwWebBrowser.Navigate ("C:\Windows\Start Menu\Programs\")
frmBrowser.Label1.Caption = "Windows Progs"
End Sub

Private Sub Command3_Click()
If Button = 1 Then
Anim = "Merlin"
Set Char = Agent1.Characters(Anim)
Char.Play ("Explain")
Char.Speak ("Well It seems you need help getting around UOS.")
Shape1.Visible = True
Char.Speak ("If you click here, you'll go to the command center of UOS.  There you can explore your hardrive and folders and settings.")
Pause (5)
Shape1.Visible = False
Shape7.Visible = True
Char.Speak ("Click here to access I Browser and PC Navigator.  With I Browser you can access the internet with PC Navigator you can explore you own C:\")
Pause (5)
Shape7.Visible = False
Shape8.Visible = True
Char.Speak ("Click here to acess you Computer Settings and Options.  Also you can access UOS Settings and Options.")
Pause (5)
Shape8.Visible = False
Shape9.Visible = True
Char.Speak ("Click here to go to the trash bin.  There you'll find your deleted programs.")
Pause (5)
Shape9.Visible = False
Shape2.Visible = True
Char.Speak ("Click here to run Windows Based Programs.")
Pause (5)
Shape2.Visible = False
Shape3.Visible = True
Char.Speak ("Click here to get me for more help.")
Pause (5)
Shape3.Visible = False
Shape4.Visible = True
Char.Speak ("Click here to access Settings.")
Pause (5)
Shape4.Visible = False
Shape5.Visible = True
Char.Speak ("Clik here to Search For Files in you Computer.")
Pause (5)
Shape5.Visible = False
Shape10.Visible = True
Char.Speak ("Click here to Run a Program.")
Pause (5)
Shape10.Visible = False
Shape6.Visible = True
Char.Speak ("Click here to Power off the computer or return to Windows.")
Pause (5)
Shape6.Visible = False
Shape11.Visible = True
Char.Speak ("Click here to make the Menu Bar show up.")
Pause (5)
Shape11.Visible = False
Shape12.Visible = True
Char.Speak ("Click here to hide Menu Bar.")
Pause (5)
Shape12.Visible = False
Char.Speak ("Well, you know every thing there is to know on how to use UOS.")
Char.Speak ("Click on me again to repeat his help.  Right Click on me to hide me.  Also if you don't want to hide me but want me go stay out of your way please click and drag me anywhere on the screen.  Well have Fun!!!!")
End If
End Sub

Private Sub Command4_Click()
Form3.Show
End Sub

Private Sub Command5_Click()
Call ShowFindDialog
End Sub

Private Sub Command6_Click()
Call ShowRunDialog(Me, "UOS Run", _
        "Select the file you want to open.")
End Sub

Private Sub Command7_Click()
Dialog.Show
End Sub

Private Sub Command8_Click()
Frame1.Visible = True
End Sub

Private Sub Command9_Click()
Frame1.Visible = False
End Sub

Private Sub Form_Load()
Anim = "Wizard"
Agent1.Characters.Load Anim, Anim & ".acs"
Set Char = Agent1.Characters(Anim)
Char.Show
Char.MoveTo 300, 300
Char.Speak ("Yo!  What Up My Man! Yeah!!  Welcome to UOS 2003!  Made by OTS and OI!")
Char.Play "Explain"
Char.Speak ("If you don't know me, I'll introduce myself.  My name is Merlin you can call me Mr.Dudely, Dr.Jude, or Principal Judely, or Sir Dudealot!")
Char.Speak ("Click on me and I'll help you get around UOS.")
Char.Speak ("You can click and drag me anywhere you want to.")
Char.Speak ("Well you can start using UOS now Dude!  Have Fun!!! And remember I'll be around if you need any help.  See ya later!")
Char.Play "DoMagic1"
Char.Play "DoMagic2"
lblTime.Caption = Time
End Sub



Private Sub Picture1_Click()
Shape7.Visible = True
End Sub

Private Sub Picture1_DblClick()
Shape7.Visible = False
frmBrowser.Show
End Sub

Private Sub Picture2_Click()
Shape8.Visible = True
End Sub

Private Sub Picture2_DblClick()
Shape8.Visible = False
Form3.Show
End Sub

Private Sub Picture2_DragDrop(Source As Control, X As Single, Y As Single)
Shape8.Visible = False
Form3.Show
End Sub

Private Sub Picture3_Click()
Shape9.Visible = True
End Sub

Private Sub Picture3_DblClick()
Shape9.Visible = False
frmBrowser.Show
frmBrowser.Label1.Caption = "Trash"
frmBrowser.brwWebBrowser.Navigate ("C:\Recycled")
End Sub

Private Sub Picture4_Click()
Shape1.Visible = True
End Sub

Private Sub Picture4_DblClick()
Shape1.Visible = False
Form5.Show
End Sub

Private Sub Task1_Click()
If Task1.Caption = "Virus Scanner" Then
frmScan.Show
End If
If Task1.Caption = "UOS Stuff" Then
Form5.Show
End If
If Task1.Caption = "PC Navig/I Brow" Then
frmBrowser.Show
End If
If Task1.Caption = "Settings" Then
Form3.Show
End If
End Sub

Private Sub Task2_Click()
If Task2.Caption = "Virus Scanner" Then
frmScan.Show
End If
If Task2.Caption = "UOS Stuff" Then
Form5.Show
End If
If Task2.Caption = "PC Navig/I Brow" Then
frmBrowser.Show
End If
If Task2.Caption = "Settings" Then
Form3.Show
End If
End Sub

Private Sub Task3_Click()
If Task3.Caption = "Virus Scanner" Then
frmScan.Show
End If
If Task3.Caption = "UOS Stuff" Then
Form5.Show
End If
If Task3.Caption = "PC Navig/I Brow" Then
frmBrowser.Show
End If
If Task3.Caption = "Settings" Then
Form3.Show
End If
End Sub

Private Sub Task4_Click()
If Task4.Caption = "Virus Scanner" Then
frmScan.Show
End If
If Task4.Caption = "UOS Stuff" Then
Form5.Show
End If
If Task4.Caption = "PC Navig/I Brow" Then
frmBrowser.Show
End If
If Task4.Caption = "Settings" Then
Form3.Show
End If
End Sub

Private Sub Task5_Click()
If Task5.Caption = "Virus Scanner" Then
frmScan.Show
End If
If Task5.Caption = "UOS Stuff" Then
Form5.Show
End If
If Task5.Caption = "PC Navig/I Brow" Then
frmBrowser.Show
End If
If Task5.Caption = "Settings" Then
Form3.Show
End If
End Sub

Private Sub Timer1_Timer()
lblTime.Caption = Time
End Sub
