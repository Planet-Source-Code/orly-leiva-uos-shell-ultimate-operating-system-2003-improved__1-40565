VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Form5 
   BorderStyle     =   0  'None
   Caption         =   "Form5"
   ClientHeight    =   5895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7650
   LinkTopic       =   "Form5"
   ScaleHeight     =   5895
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000F&
      Height          =   495
      Left            =   2280
      Picture         =   "UOS Stuff.frx":0000
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   1
      Top             =   480
      Width           =   495
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   7095
      ExtentX         =   12515
      ExtentY         =   8070
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      Height          =   255
      Left            =   7200
      TabIndex        =   8
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label6 
      BackColor       =   &H000000FF&
      Caption         =   "UOS Stuff"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   7695
   End
   Begin VB.Label Label5 
      Caption         =   "C:\"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   615
   End
   Begin VB.Image imgAddRemove 
      Height          =   600
      Left            =   1200
      Picture         =   "UOS Stuff.frx":0CCA
      Stretch         =   -1  'True
      Top             =   360
      Width           =   600
   End
   Begin VB.Label Label1 
      Caption         =   "CD ROM Drives"
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Floppy Disk"
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Settings"
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3240
      Picture         =   "UOS Stuff.frx":1A0C
      Top             =   480
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   600
      Left            =   4200
      Picture         =   "UOS Stuff.frx":41AE
      Stretch         =   -1  'True
      Top             =   360
      Width           =   720
   End
   Begin VB.Label Label4 
      Caption         =   "CD Rom Drives 2"
      Height          =   255
      Left            =   3960
      TabIndex        =   2
      Top             =   960
      Width           =   1695
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   240
      Picture         =   "UOS Stuff.frx":4EF0
      Top             =   480
      Width           =   480
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
If Form2.Task1.Caption = "" Then
Form2.Task1.Visible = True
Form2.Task1.Caption = "UOS Stuff"
ElseIf Form2.Task2.Caption = "" Then
Form2.Task2.Visible = True
Form2.Task2.Caption = "UOS Stuff"
ElseIf Form2.Task3.Caption = "" Then
Form2.Task3.Visible = True
Form2.Task3.Caption = "UOS Stuff"
ElseIf Form2.Task4.Caption = "" Then
Form2.Task4.Visible = True
Form2.Task4.Caption = "UOS Stuff"
ElseIf Form2.Task5.Caption = "" Then
Form2.Task5.Visible = True
Form2.Task5.Caption = "UOS Stuff"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Form2.Task1.Caption = "UOS Stuff" Then
Form2.Task1.Visible = False
Form2.Task1.Caption = ""
ElseIf Form2.Task2.Caption = "UOS Stuff" Then
Form2.Task2.Visible = False
Form2.Task2.Caption = ""
ElseIf Form2.Task3.Caption = "UOS Stuff" Then
Form2.Task3.Visible = False
Form2.Task3.Caption = ""
ElseIf Form2.Task4.Caption = "UOS Stuff" Then
Form2.Task4.Visible = False
Form2.Task4.Caption = ""
ElseIf Form2.Task5.Caption = "UOS Stuff" Then
Form2.Task5.Visible = False
Form2.Task5.Caption = ""
End If
End Sub

Private Sub Image1_Click()
Form3.Show
Unload Me
End Sub

Private Sub Image2_Click()
WebBrowser1.Navigate ("D:\")
End Sub

Private Sub Image3_Click()
WebBrowser1.Navigate ("C:\")
End Sub

Private Sub imgAddRemove_Click()
WebBrowser1.Navigate ("F:\")
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormMove Me
End Sub

Private Sub Label7_Click()
Unload Me
End Sub

Private Sub Picture1_Click()
WebBrowser1.Navigate ("A:\")
End Sub
