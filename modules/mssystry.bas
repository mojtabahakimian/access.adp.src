Attribute VB_Name = "mssystry"
Option Explicit

Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Public Type NOTIFYICONDATA
  cbSize As Long
  hWnd As Long
  UID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 128
  dwState As Long
  dwStateMask As Long
  szInfo As String * 256
  uTimeOutAndVersion As Long
  szInfoTitle As String * 64
  dwInfoFlags As Long
End Type

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2

'API Constants
Public Const WM_USER = &H400
Public Const WM_USER_TRAY = WM_USER + 1
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MBUTTONDBLCLK = &H209

Public Const NIF_INFO = &H10
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4

Public tico As NOTIFYICONDATA

Public Sub End_Tray(nForm As Form)
tico.cbSize = Len(tico)
tico.hWnd = nForm.hWnd
tico.UID = 1&
Shell_NotifyIcon NIM_DELETE, tico
End Sub

'////////////////////////////////////////////////////////////////////////////////////////////////////////////
Public Sub ShowNotifyBalloonTip(Optional ByVal sText As String = "", Optional ByVal sTitle As String = "", Optional ByVal nIconType As Integer = -1, Optional ByVal nTimeOut As Integer = -1)
With tico
    .uFlags = .uFlags Or NIF_INFO
    .uTimeOutAndVersion = IIf(nTimeOut = -1, 3000, nTimeOut)
    .szInfo = sText & Chr$(0)
    .szInfoTitle = sTitle & Chr$(0)
    .dwInfoFlags = nIconType
End With
Call Shell_NotifyIcon(NIM_MODIFY, tico)
tico.uFlags = tico.uFlags And Not NIF_INFO
End Sub

'////////////////////////////////////////////////////////////////////////////////////////////////////////////
Public Sub Add_Tray(nForm As Form, IconTooltip As String)
tico.cbSize = Len(tico)
tico.hWnd = nForm.hWnd
tico.UID = 1&
tico.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
tico.uCallbackMessage = WM_MOUSEMOVE

tico.szTip = IconTooltip
Shell_NotifyIcon NIM_ADD, tico
End Sub

Public Sub Modify_Tray(IconTooltip As String)
tico.szTip = IconTooltip
Shell_NotifyIcon NIM_MODIFY, tico
End Sub
