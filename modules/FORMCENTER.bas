Attribute VB_Name = "FORMCENTER"
Option Compare Database

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Declare Function GetWindowRect Lib "user32" _
    (ByVal hWnd As Long, lpRect As RECTL) As Long

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, _
  ByVal hdc As Long) As Long

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, _
  ByVal nIndex As Long) As Long

Private Type RECTL
        Left As Long
        Top As Long
        RIGHT As Long
        Bottom As Long
End Type

Public Const WU_LOGPIXELSX = 88
Public Const WU_LOGPIXELSY = 90

Public Function TwipsPerPixel(strDirection As String) As Long
'Purpose  : Get monitor's Twips per pixel
    'Handle to device
    Dim lngDC                       As Long
    Dim lngPixelsPerInch            As Long
    Const nTwipsPerInch = 1440
    lngDC = GetDC(0)

    If strDirection = "X" Then                                  'Horizontal
        lngPixelsPerInch = GetDeviceCaps(lngDC, WU_LOGPIXELSX)
    Else                                                        'Vertical
        lngPixelsPerInch = GetDeviceCaps(lngDC, WU_LOGPIXELSY)
    End If
    lngDC = ReleaseDC(0, lngDC)
    TwipsPerPixel = nTwipsPerInch / lngPixelsPerInch

End Function

Public Sub WindowSize(ByRef Height As Long, ByRef Width As Long)
'Purpose  : Get Access window size.
    Dim hWnd                        As Long
    Dim rct                         As RECTL

    'hwnd = FindWindow(vbNullString, "Microsoft Access")
    hWnd = Application.hWndAccessApp
    If hWnd <> 0 And GetWindowRect(hWnd, rct) <> 0 Then
        Height = (rct.Bottom - rct.Top) * TwipsPerPixel("Y")
        Width = (rct.RIGHT - rct.Left) * TwipsPerPixel("X")
    End If
End Sub

Public Function CenterMe(frm As Form)
'Purpose  : Center form on screen.
'Requires :
'   Code
'       TwipsPerPixel()
'       WindowSize()
'       Type RECTL
'   API Libraries
'       FindWindow
'       GetWindowRect
'       GetDC
'       ReleaseDC
'       GetDeviceCaps
           On Error Resume Next
    Dim lngWinWidth As Long
    Dim lngWinHeight As Long
    Dim lngFrmWidth As Long
    Dim lngFrmHeight As Long

    Call WindowSize(lngWinHeight, lngWinWidth)
  '  frm.SetFocus

    DoCmd.MoveSize (lngWinWidth - frm.WindowWidth) / 2, Abs(lngWinHeight - frm.WindowHeight) / 2

End Function
