VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   0  'None
   Caption         =   "Dialog Caption"
   ClientHeight    =   1275
   ClientLeft      =   2715
   ClientTop       =   3420
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton Option4 
      Caption         =   "Restart the Computer"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1935
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Restrat UOS"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1935
   End
   Begin VB.OptionButton Option2 
      Caption         =   "ShutDown the computer"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Return To Windows"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      Height          =   255
      Left            =   5520
      TabIndex        =   7
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "ShutDown?"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6015
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormMove Me
End Sub

Private Sub Label2_Click()
Unload Me
End Sub

Private Sub OKButton_Click()
If Option1.Value = True Then
End
End If
If Option2.Value = True Then
Call ShutDown
End If
If Option3.Value = True Then
Form1.Show
Unload Me
End If
If Option4.Value = True Then
Call Restart
End If
End Sub
