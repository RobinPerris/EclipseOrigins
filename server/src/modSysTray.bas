Attribute VB_Name = "modSysTray"
Option Explicit

Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Const KEYEVENTF_KEYUP = &H2
Public Const VK_LWIN = &H5B

'Declare a user-defined variable to pass to the Shell_NotifyIcon
'function.
Public Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

'Declare the constants for the API function. These constants can be
'found in the header file Shellapi.h.
'The following constants are the messages sent to the
'Shell_NotifyIcon function to add, modify, or delete an icon from the System Tray
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
'The following constant is the message sent when a mouse event occurs
'within the rectangular boundaries of the icon in the System Tray
'area.
Public Const WM_MOUSEMOVE = &H200
'The following constants are the flags that indicate the valid
'members of the NOTIFYICONDATA data type.
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
'The following constants are used to determine the mouse input on the
'the icon in the taskbar status area.
'Left-click constants.
Public Const WM_LBUTTONDBLCLK = &H203   'Double-click
Public Const WM_LBUTTONDOWN = &H201     'Button down
Public Const WM_LBUTTONUP = &H202       'Button up
'Right-click constants.
Public Const WM_RBUTTONDBLCLK = &H206   'Double-click
Public Const WM_RBUTTONDOWN = &H204     'Button down
Public Const WM_RBUTTONUP = &H205       'Button up
'Declare the API function call.
Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
'Dimension a variable as the user-defined data type.
Global nid As NOTIFYICONDATA

Public Sub DestroySystemTray()
    nid.cbSize = Len(nid)
    nid.hWnd = frmServer.hWnd
    nid.uId = vbNull
    nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    nid.uCallBackMessage = WM_MOUSEMOVE
    nid.hIcon = frmServer.Icon
    nid.szTip = "Server" & vbNullChar
    Call Shell_NotifyIcon(NIM_DELETE, nid) ' Add to the sys tray
End Sub

Public Sub LoadSystemTray()
    nid.cbSize = Len(nid)
    nid.hWnd = frmServer.hWnd
    nid.uId = vbNull
    nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    nid.uCallBackMessage = WM_MOUSEMOVE
    nid.hIcon = frmServer.Icon
    nid.szTip = "Server" & vbNullChar   'You can add your game name or something.
    Call Shell_NotifyIcon(NIM_ADD, nid) 'Add to the sys tray
End Sub
