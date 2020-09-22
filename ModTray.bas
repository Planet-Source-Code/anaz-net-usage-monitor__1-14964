Attribute VB_Name = "Tray"
Option Explicit 'What can I say about it.Needed so that confusion is minimun

'Tip-Use Api declaration loader to be fast. Everything here is there,in Win32api
Public Declare Function SHAppBarMessage Lib "shell32.dll" (ByVal dwMessage As Long, pData As APPBARDATA) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Global Notify As NOTIFYICONDATA
Global BarData As APPBARDATA
Global Const WM_MOUSEMOVE = &H200
Global Const ABM_GETTASKBARPOS = &H5&

Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uId As Long
        uFlags As Long
        ucallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

Type APPBARDATA
        cbSize As Long
        hwnd As Long
        ucallbackMessage As Long
        uEdge As Long
        rc As RECT
        lParam As Long
End Type

Public Iconobj As Object
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIF_MESSAGE = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_RBUTTONUP = &H205
Public Const WM_LBUTTONDBLCLK = &H203



Public Sub modIcon(frmMain As Form, IconID As Long, icon As Object)
    Dim Result As Long
    Notify.cbSize = 88&
    Notify.hwnd = frmMain.hwnd
    Notify.uId = IconID
    Notify.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    Notify.ucallbackMessage = WM_MOUSEMOVE
    Notify.hIcon = icon
    Notify.szTip = ""
    Result = Shell_NotifyIcon(NIM_MODIFY, Notify)
End Sub
Public Sub AddIcon(frmMain As Form, IconID As Long, icon As Object)
    Dim Result As Long
    BarData.cbSize = 36&
    Result = SHAppBarMessage(ABM_GETTASKBARPOS, BarData)
    Notify.cbSize = 88&
    Notify.hwnd = frmMain.hwnd
    Notify.uId = IconID
    Notify.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    Notify.ucallbackMessage = WM_MOUSEMOVE
    Notify.hIcon = icon
    Notify.szTip = "Monitering.." & Chr$(0)
   Result = Shell_NotifyIcon(NIM_ADD, Notify)
End Sub
Public Sub delIcon(IconID As Long)
    Dim Result As Long
    Notify.uId = IconID
    Result = Shell_NotifyIcon(NIM_DELETE, Notify)
End Sub
