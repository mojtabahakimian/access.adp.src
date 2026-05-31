Attribute VB_Name = "Module2"
Option Compare Database
Option Explicit

' =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' API INFO:
'   The ReleaseCapture function releases the mouse capture from a window in the current
'   thread and restores normal mouse input processing. A window that has captured the mouse
'   receives all mouse input, regardless of the position of the cursor, except when a mouse
'   button is clicked while the cursor hot spot is in the window of another thread.
' =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
#If VBA7 Then
    Private Declare PtrSafe Function ReleaseCapture Lib "user32" () As Long
#Else
    Private Declare Function ReleaseCapture Lib "user32" () As Long
#End If
' =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' API INFO:
'   The SendMessage function sends the specified message to a window or windows. It calls
'   the window procedure for the specified window and does not return until the window
'   procedure has processed the message.
'
' PARAMETER(S):
'   hWnd
'       [in] Handle to the window whose window procedure will receive the message. If this
'            parameter is HWND_BROADCAST, the message is sent to all top-level windows in
'            the system, including disabled or invisible unowned windows, overlapped windows,
'            and pop-up windows; but the message is not sent to child windows.
'   Msg
'       [in] Specifies the message to be sent.
'   wParam
'       [in] Specifies additional message-specific information.
'   lParam
'       [in] Specifies additional message-specific information.
' =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
#If VBA7 Then
        Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" ( _
                ByVal hWnd As LongPtr, _
                    ByVal wMsg As Long, _
                        ByVal wParam As LongPtr, _
                            lParam As Any) As LongPtr
#Else
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
                                ByVal hWnd As Long, _
                                    ByVal wMsg As Long, _
                                        ByVal wParam As Long, _
                                            lParam As Any) As Long
#End If
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private Const HWND_BROADCAST = &HFFFF

Public Function fDragWindow(frm As Access.Form)

    With frm

        ReleaseCapture
        SendMessage .hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0

    End With

End Function
