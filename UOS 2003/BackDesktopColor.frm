VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form4"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command12 
      Caption         =   "OK for User Login"
      Height          =   495
      Left            =   360
      TabIndex        =   17
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   840
      TabIndex        =   15
      Top             =   1920
      Width           =   3615
   End
   Begin VB.PictureBox Command1 
      BackColor       =   &H80000001&
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   11
      Top             =   720
      Width           =   615
   End
   Begin VB.PictureBox Command2 
      BackColor       =   &H00FF0000&
      Height          =   375
      Left            =   840
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   10
      Top             =   720
      Width           =   615
   End
   Begin VB.PictureBox Command3 
      BackColor       =   &H0000FF00&
      Height          =   375
      Left            =   1560
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   9
      Top             =   720
      Width           =   615
   End
   Begin VB.PictureBox Command4 
      BackColor       =   &H000000FF&
      Height          =   375
      Left            =   2280
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   8
      Top             =   720
      Width           =   615
   End
   Begin VB.PictureBox Command5 
      BackColor       =   &H0000FFFF&
      Height          =   375
      Left            =   3000
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   7
      Top             =   720
      Width           =   615
   End
   Begin VB.PictureBox Command6 
      Height          =   375
      Left            =   3720
      ScaleHeight     =   315
      ScaleWidth      =   675
      TabIndex        =   6
      Top             =   720
      Width           =   735
   End
   Begin VB.PictureBox Command8 
      BackColor       =   &H8000000D&
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   5
      Top             =   1080
      Width           =   615
   End
   Begin VB.PictureBox Command9 
      BackColor       =   &H80000007&
      Height          =   375
      Left            =   840
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   4
      Top             =   1080
      Width           =   615
   End
   Begin VB.PictureBox Command11 
      BackColor       =   &H80000002&
      Height          =   375
      Left            =   1560
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   3
      Top             =   1080
      Width           =   615
   End
   Begin VB.PictureBox Command10 
      BackColor       =   &H80000009&
      Height          =   375
      Left            =   2280
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   2
      Top             =   1080
      Width           =   615
   End
   Begin VB.PictureBox Command7 
      BackColor       =   &H80000003&
      Height          =   375
      Left            =   3000
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   1
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "User Login"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      Height          =   255
      Left            =   4200
      TabIndex        =   13
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Background Color"
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "Desktop Settings"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.BackColor = vbDesktop

End Sub

Private Sub Command10_Click()
Form2.BackColor = vbWindowBackground
End Sub

Private Sub Command11_Click()
Form2.BackColor = vbActiveTitleBar
End Sub

Private Sub Command12_Click()
Form2.Label6.Caption = "User:" & Text1.Text
End Sub

Private Sub Command2_Click()
Form2.BackColor = vbBlue
End Sub

Private Sub Command3_Click()
Form2.BackColor = vbGreen
End Sub

Private Sub Command4_Click()
Form2.BackColor = vbRed
End Sub

Private Sub Command5_Click()
Form2.BackColor = vbYellow
End Sub

Private Sub Command6_Click()
Form2.BackColor = vbButtonFace
End Sub

Private Sub Command7_Click()
Form2.BackColor = vbInactiveTitleBar
End Sub

Private Sub Command8_Click()
Form2.BackColor = vbHighlight
End Sub

Private Sub Command9_Click()
Form2.BackColor = vbMenuText
End Sub


Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormMove Me
End Sub

Private Sub Label3_Click()
Unload Me
End Sub
