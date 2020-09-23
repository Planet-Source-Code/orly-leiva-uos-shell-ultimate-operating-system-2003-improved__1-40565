Attribute VB_Name = "modAPI"
'Declaring stuff for the Find Dialog
Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation _
    As String, ByVal lpFile As String, ByVal lpParameters _
    As String, ByVal lpDirectory As String, ByVal nShowCmd _
    As Long) As Long
   
Const SW_SHOW = 5

'Declaring stuff for the Reboot Dialog
Private Declare Function SHRestartSystemMB Lib _
"shell32" Alias "#59" (ByVal hOwner As Long, ByVal _
sExtraPrompt As String, ByVal uFlags As Long) As Long

Private Const SystemChangeRestart = 4


'Declaring stuff for Disabling the CTRL+ALT+DEL box
Private Declare Function SystemParametersInfo Lib _
"user32" Alias "SystemParametersInfoA" (ByVal uAction _
As Long, ByVal uParam As Long, ByVal lpvParam As Any, _
ByVal fuWinIni As Long) As Long

'Declaring stuff for the Message Box
Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long

'Message Box Types
Public Const MB_ABORTRETRYIGNORE = &H2& 'Abort, Retry, Ignore
Public Const MB_YESNO = &H4& ' Yes and No
Public Const MB_YESNOCANCEL = &H3& 'Yes, No, Cancel
Public Const MB_RETRYCANCEL = &H5& 'Retry and Cancel
Public Const MB_OKCANCEL = &H1& 'Ok and Cancel
Public Const MB_OK = &H0& 'Just OK

'Icons
Public Const MB_ICONSTOP = &H10& 'Stop Icon
Public Const MB_ICONQUESTION = &H20& 'Question Mark Icon
Public Const MB_ICONASTERISK = &H40& 'Asterisk Icon
Public Const MB_ICONEXCLAMATION = &H30& 'Exclamation Icon

'Button Types
Public Const IDYES = 6 'Yes Button
Public Const IDNO = 7 'No Button
Public Const IDABORT = 3 'Abort Button
Public Const IDCANCEL = 2 'Cancel Button
Public Const IDIGNORE = 5 'Ignore Button
Public Const IDRETRY = 4 'Retry Button
Public Const IDOK = 1 'Ok Button

Private Declare Function SHRunDialog Lib "shell32" _
    Alias "#61" (ByVal hOwner As Long, ByVal UnknownP1 _
    As Long, ByVal UnknownP2 As Long, ByVal szTitle _
    As String, ByVal szPrompt As String, ByVal uFlags _
    As Long) As Long
    
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Const EWX_FORCE = 4
Public Const EWX_LOGOFF = 0
Public Const EWX_REBOOT = 2
Public Const EWX_SHUTDOWN = 1

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Declare Sub ReleaseCapture Lib "user32" ()
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long

Option Explicit

Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function GetForegroundWindow Lib "user32" () As Long
Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare Function TerminateProcess Lib "KERNEL32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
Declare Function OpenProcess Lib "KERNEL32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Declare Function ProcessFirst Lib "KERNEL32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Declare Function ProcessNext Lib "KERNEL32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Declare Function CreateToolhelpSnapshot Lib "KERNEL32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Declare Function CloseHandle Lib "KERNEL32" (ByVal hObject As Long) As Long
Declare Function WritePrivateProfileString Lib "KERNEL32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private Declare Function GetPrivateProfileString& Lib "KERNEL32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String)

Public Declare Function RegisterServiceProcess Lib "KERNEL32" (ByVal ProcessID As Long, ByVal ServiceFlags As Long) As Long
Public Declare Function GetCurrentProcessId Lib "KERNEL32" () As Long

Public Const SW_RESTORE = 9
Public Const GW_OWNER = 4
Public Const GWL_HWNDPARENT = (-8)
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_TOOLWINDOW = &H80
Public Const WS_EX_APPWINDOW = &H40000
Public Const LB_ADDSTRING = &H180
Public Const LB_SETITEMDATA = &H19A
Const MAX_PATH& = 260

Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * MAX_PATH
End Type


Public Sub pSetForegroundWindow(ByVal hWnd As Long)
Dim lForeThreadID As Long
Dim lThisThreadID As Long
Dim lReturn       As Long

If hWnd <> GetForegroundWindow() Then
    
    lForeThreadID = GetWindowThreadProcessId(GetForegroundWindow, ByVal 0&)
    lThisThreadID = GetWindowThreadProcessId(hWnd, ByVal 0&)
    
    If lForeThreadID <> lThisThreadID Then
        Call AttachThreadInput(lForeThreadID, lThisThreadID, True)
        lReturn = SetForegroundWindow(hWnd)
        Call AttachThreadInput(lForeThreadID, lThisThreadID, False)
    Else
       lReturn = SetForegroundWindow(hWnd)
    End If
       If IsIconic(hWnd) Then
       Call ShowWindow(hWnd, SW_RESTORE)
    Else
       Call ShowWindow(hWnd, SW_SHOW)
    End If
End If
End Sub
Public Function WhichWindows(lst As ListBox) As Long
With lst
    .Clear
    Call EnumWindows(AddressOf WhichWindowsCallBack, .hWnd)
    WhichWindows = .ListCount
End With
End Function

Private Function WhichWindowsCallBack(ByVal hWnd As Long, ByVal lParam As Long) As Long
Dim lReturn     As Long
Dim lExStyle    As Long
Dim bNoOwner    As Boolean
Dim sWindowText As String

If hWnd <> frmDesktop.hWnd Then
    If IsWindowVisible(hWnd) Then
        If GetParent(hWnd) = 0 Then
            bNoOwner = (GetWindow(hWnd, GW_OWNER) = 0)
            lExStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
            
            If (((lExStyle And WS_EX_TOOLWINDOW) = 0) And bNoOwner) Or _
                ((lExStyle And WS_EX_APPWINDOW) And Not bNoOwner) Then
                
                sWindowText = Space$(256)
                lReturn = GetWindowText(hWnd, sWindowText, Len(sWindowText))
                If lReturn Then
                  
                   sWindowText = Left$(sWindowText, lReturn)
                   lReturn = SendMessage(lParam, LB_ADDSTRING, 0, ByVal sWindowText)
                   Call SendMessage(lParam, LB_SETITEMDATA, lReturn, ByVal hWnd)
                End If
            End If
        End If
    End If
End If
WhichWindowsCallBack = True
End Function

Sub Restart()
Call ExitWindowsEx(EWX_REBOOT, 0)
End Sub

Public Sub ShowFindDialog(Optional InitialDirectory As String)

ShellExecute 0, "find", _
  IIf(InitialDirectory = "", "", InitialDirectory), _
  vbNullString, vbNullString, SW_SHOW

End Sub

Public Sub SettingsChanged(FormName As Form)
    SHRestartSystemMB FormName.hWnd, vbNullString, SystemChangeRestart
End Sub

Sub DisableCAD(bDisabled As Boolean)
    Dim x As Long
    x = SystemParametersInfo(97, bDisabled, CStr(1), 0)
End Sub

Public Sub ShowRunDialog(ByRef CallingForm As Form, _
    Optional Title As String, _
    Optional Description As String)
    
    If Title = "" Then Title = "Run"
    
    If Description = "" Then Description = _
    "Type the name of a program to open, " & _
        "then click OK when finished."
    
    SHRunDialog CallingForm.hWnd, 0, 0, _
        Title, Description, 0
        
End Sub

Sub ShutDown()
Call ExitWindowsEx(EWX_SHUTDOWN, 0)
End Sub

Sub FormMove(TheForm As Form)
    ReleaseCapture
    Call SendMessage(TheForm.hWnd, &HA1, 2, 0&)
End Sub

Sub Pause(Interval)
    Dim Current
    
    Current = Timer
    Do While Timer - Current < Val(Interval)
        DoEvents
    Loop
End Sub


