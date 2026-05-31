Attribute VB_Name = "Module3"
Option Compare Database
Option Explicit

' https://www.access-programmers.co.uk/forums/threads/position-and-size-the-access-window.172827/
'
' modified by arnelgp for x64
'
#If VBA7 Then
Private Declare PtrSafe Function apiMoveWindow Lib "user32" Alias "MoveWindow" _
         (ByVal hWnd As LongPtr, ByVal x As Long, ByVal Y As Long, ByVal _
         nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) _
         As Long

Private Declare PtrSafe Function apiGetActiveWindow Lib "user32" Alias "GetActiveWindow" () As LongPtr

Private Declare PtrSafe Function apiGetParent Lib "user32" Alias "GetParent" (ByVal hWnd As LongPtr) As LongPtr
#Else
Private Declare Function apiMoveWindow Lib "user32" Alias "MoveWindow" _
         (ByVal hWnd As Long, ByVal x As Long, ByVal Y As Long, ByVal _
         nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) _
         As Long

Private Declare Function apiGetActiveWindow Lib "user32" Alias "GetActiveWindow" () As Long

Private Declare Function apiGetParent Lib "user32" Alias "GetParent" (ByVal hWnd As Long) As Long
#End If

Function AccessMoveSize(iX As Integer, iY As Integer, iWidth As _
         Integer, iHeight As Integer)
    apiMoveWindow GetAccesshWnd(), iX, iY, iWidth, iHeight, True
End Function


Function GetAccesshWnd()
#If VBA7 Then
    Dim hWnd As LongPtr
    Dim hWndAccess As LongPtr
#Else
    Dim hWnd As Long
    Dim hWndAccess As Long
#End If
    ' Get the handle to the currently active window.
    hWnd = apiGetActiveWindow()
    hWndAccess = hWnd

    ' Find the top window (which has no parent window).
    While hWnd <> 0
        hWndAccess = hWnd
        hWnd = apiGetParent(hWnd)
    Wend

    GetAccesshWnd = hWndAccess

End Function
