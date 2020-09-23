VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmScan 
   Caption         =   "Virus Scanner"
   ClientHeight    =   2295
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5655
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   5655
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar Pb1 
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar Stb1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   2040
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6968
            MinWidth        =   6968
            Text            =   "Virus Scanner is Idle."
            TextSave        =   "Virus Scanner is Idle."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstScan 
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   3201
      View            =   3
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Problem [String]"
         Object.Width           =   7621
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Priority"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.ListBox lstLog 
      BackColor       =   &H80000004&
      ForeColor       =   &H00FF0000&
      Height          =   1620
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
   Begin VB.Menu mnuExit 
      Caption         =   "&Exit"
   End
   Begin VB.Menu mnuReset 
      Caption         =   "&Reset"
   End
   Begin VB.Menu a 
      Caption         =   "About"
   End
End
Attribute VB_Name = "frmScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intFiles        As Integer ' All of our Variables used
Dim intX            As Integer
Dim intLenFile      As Integer
Dim lngBufferLen    As Long
Dim strFilePath     As String
Dim strBuffer       As String
Dim bByte           As Byte
Dim StartTime       As Single
Dim wBorder         As Single
Dim lstAdd          As ListItem

Private Sub a_Click()
Form7.Show
End Sub

Private Sub Form_Load()
ScaleMode = vbTwips ' Sets ScaleMode type
    wBorder = (Width - ScaleWidth) / 4
    Pb1.Move Stb1.Panels(2).Left + 30, Stb1.Top + wBorder + 20, Stb1.Panels(2).Width - 50, Stb1.Height - wBorder - 30 ' Moves Progressbar firmly into Panel 2
Pb1.Visible = True ' Show the Progressbar after drawing it
If Form2.Task1.Caption = "" Then
Form2.Task1.Visible = True
Form2.Task1.Caption = "Virus Scanner"
ElseIf Form2.Task2.Caption = "" Then
Form2.Task2.Visible = True
Form2.Task2.Caption = "Virus Scanner"
ElseIf Form2.Task3.Caption = "" Then
Form2.Task3.Visible = True
Form2.Task3.Caption = "Virus Scanner"
ElseIf Form2.Task4.Caption = "" Then
Form2.Task4.Visible = True
Form2.Task4.Caption = "Virus Scanner"
ElseIf Form2.Task5.Caption = "" Then
Form2.Task5.Visible = True
Form2.Task5.Caption = "Virus Scanner"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If MsgBox("Are you sure you want to Exit?", vbYesNo + vbQuestion, "Confirm") = vbYes Then
    Unload Me
    Else
    Cancel = 1
End If
If Form2.Task1.Caption = "Virus Scanner" Then
Form2.Task1.Visible = False
Form2.Task1.Caption = ""
ElseIf Form2.Task2.Caption = "Virus Scanner" Then
Form2.Task2.Visible = False
Form2.Task2.Caption = ""
ElseIf Form2.Task3.Caption = "Virus Scanner" Then
Form2.Task3.Visible = False
Form2.Task3.Caption = ""
ElseIf Form2.Task4.Caption = "Virus Scanner" Then
Form2.Task4.Visible = False
Form2.Task4.Caption = ""
ElseIf Form2.Task5.Caption = "Virus Scanner" Then
Form2.Task5.Visible = False
Form2.Task5.Caption = ""
End If
End Sub

Private Sub lstLog_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
lstLog.AddItem "Free File Dropped."
    Stb1.Panels(1).Text = "Accepting File..."
    DoEvents
intFiles = Data.Files.Count
For intX = 1 To intFiles
    If (GetAttr(Data.Files(intX)) And vbDirectory) = vbDirectory Then ' Check for Directory
        lstLog.AddItem "Unable to Scan Directory - This feature is not yet included.": Exit Sub
        Stb1.Panels(1).Text = "Cannot Load Directory": lstLog.Clear
        Else
        intLenFile = Len(Data.Files(intX)) ' Get Length
        strFilePath = Left(Data.Files(intX), intLenFile) ' Get File Path
        lstLog.AddItem "Reading File Path: " & strFilePath
    End If
Next intX
lstLog.AddItem "Accessing Binary File (Loading Data to Buffer)..."
Stb1.Panels(1).Text = "Reading File..."
Open (strFilePath) For Binary As #1 ' Open to read Bytes
    DoEvents ' Make sure everything doesn't freeze as we progress
    Stb1.Panels(1).Text = "Writing Bytes to Buffer Space..."
    lstLog.AddItem "... Reading File Length: " & Format(LOF(1) / 1000, "###,###,###.###") & " k."
    Pb1.Max = LOF(1) + 1 'This will error if + 1 is gone
    lstLog.AddItem "... Reading and Writing Bytes"
    Do While Not EOF(1) ' Do not keep reading if End of File
        DoEvents
        Get 1, , bByte ' Getting the Bytes 1 by 1 (Slow Method)
        Pb1.Value = Pb1.Value + 1 ' Updating Progression
        strBuffer = strBuffer & Chr(bByte) ' Turning each Byte into ASCII and storing in Buffer
    Loop
    lstLog.AddItem "Created Buffer Space"
    Pb1.Value = 0
Close #1
lstLog.AddItem "Closed File for Read as Binary"
    lstLog.Visible = False: lstScan.Visible = True
    Stb1.Panels(1).Text = "Scanning Buffer for Virus Strings..."
    DoEvents
    SearchBuffer strBuffer ' Call our Search Function
    Stb1.Panels(1).Text = "Buffer Scan Complete."
    lngBufferLen = Len(strBuffer) ' Get Length of buffer
    strBuffer = Space(lngBufferLen) ' Create Empty Space with same amount in Buffer
    Pause 2
Stb1.Panels(1).Text = "Buffer Emptied."
End Sub

Private Sub SearchBuffer(TheBuffer As String)
'This will give you the basic idea on how to search the buffer for specific strings.
'How accurate everything is, is unknown, some strings may have spaces throughout...
'so it will be hard to perfectly check everything.  Find a site that gives the...
'Virus strings, and use those.
If InStr(1, TheBuffer, "win32.exe") Then
    AddList "File Loads itself in 'Win32' Application", "Low"
End If
If InStr(1, TheBuffer, "remove") Then
    AddList "File Removes an Object", "Medium"
End If
If InStr(1, TheBuffer, "kill") Then
    AddList "File Destroys an Object", "Medium"
End If
If InStr(1, TheBuffer, "@hotmail.com") Then
    AddList "File possibly sends Email to a Hotmail Account", "Low"
End If
If InStr(1, TheBuffer, "win.ini") Then
    AddList "File reads the 'Win.ini' File", "High"
End If
If InStr(1, TheBuffer, "system.ini") Then
    AddList "File reads the 'System.ini' File", "High"
End If
If InStr(1, TheBuffer, "STEALER1") Then
    AddList "File is an AOL Trojan", "High"
End If
End Sub

Private Sub Pause(WaitSeconds As Single)
StartTime = Timer
Do While Timer < StartTime + WaitSeconds
    DoEvents
Loop
End Sub

Private Sub AddList(TheProblem As String, ThePriority As String)
Set lstAdd = lstScan.ListItems.Add(, , TheProblem) ' Add the Problem to the first column
    lstAdd.SubItems(1) = ThePriority ' Add our Priority Level to the second column
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuReset_Click()
lstScan.Visible = False
    lstLog.Visible = True
    lstLog.Clear
Stb1.Panels(1).Text = "Awaiting New File..."
End Sub
