Attribute VB_Name = "ModSystray"
Declare Function RegisterTray Lib "tray.dll" _
               (ByVal hInst As Long, _
               ByVal trayXpos As Long, _
               ByVal trayYpos As Long, _
               ByVal trayHeight As Long, _
               ByVal useEmptyTray As Long, _
               ByVal autoHiding As Long) As Long
'User Defined Type NOTIGYICONDATA Which is required by the API to store info on the icon data
      Public Type NOTIFYICONDATA
       cbSize As Long 'Size of this Data Type
       hWnd As Long 'Visual output
       uId As Long
       uFlags As Long ' Various Command Parameters\Flags to sent to Api
       uCallBackMessage As Long
       hIcon As Long 'Where the icon is store.
       szTip As String * 64 'Tool TIp Text
      End Type

      'constants required by Shell_NotifyIcon API call:
      Public Const WM_LBUTTONDOWN = &H201     'Button down
      Public Const WM_LBUTTONUP = &H202       'Button up
      Public Const WM_LBUTTONDBLCLK = &H203   'Double-click
      Public Const WM_RBUTTONDOWN = &H204     'Button down
      Public Const WM_RBUTTONUP = &H205       'Button up
      Public Const WM_RBUTTONDBLCLK = &H206   'Double-click

      Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
      Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
      Public IconInfo As NOTIFYICONDATA 'Declare varible to store icon info
Public Const WM_SETHOTKEY = &H32

Public Enum enm_NIM_Shell
    NIM_ADD = &H0
    NIM_MODIFY = &H1
    NIM_DELETE = &H2
    NIF_MESSAGE = &H1
    NIF_ICON = &H2
    NIF_TIP = &H4
    WM_MOUSEMOVE = &H200
End Enum

Public nidProgramData As NOTIFYICONDATA

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long


Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Parent As Long
Public SysBox As Long
Public Sub SysTrayBootUp()
    Dim hWnd As Long, rctemp As RECT
    
    hWnd = FindWindow("Shell_TrayWnd", vbNullString)
    hWnd = FindWindowEx(hWnd, 0, "TrayNotifyWnd", vbNullString)
    SysBox = hWnd
    Parent = GetParent(SysBox)
    SetParent SysBox, Form2.picSystray.hWnd
    SetWindowPos SysBox, 0, 0, 0, 150, 100, 0
End Sub

Public Sub AddSysTrayIcon()
With IconInfo
        .cbSize = Len(IconInfo)
        .hWnd = frmDesktop.hWnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Form2.Icon
        .szTip = "Holy Grail" & vbNullChar
       End With
       Shell_NotifyIcon NIM_ADD, IconInfo
End Sub

Public Sub DelSysTrayIcon()
       Shell_NotifyIcon NIM_DELETE, IconInfo
End Sub

