Attribute VB_Name = "Module4"
Option Compare Database
Option Explicit

Public Type RECT
   x1 As Long
   Y1 As Long
   x2 As Long
   Y2 As Long
End Type

' https://www.access-programmers.co.uk/forums/threads/position-and-size-the-access-window.172827/
'
' modified by arnelgp for x64
'
'*****************************************************************
' DECLARATIONS SECTION
'*****************************************************************



' NOTE: The following declare statements are case sensitive.

#If VBA7 Then
    Public Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr
    Public Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, rectangle As RECT) As Long
#Else
    Public Declare Function GetDesktopWindow Lib "user32" () As Long
    Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, rectangle As RECT) As Long
#End If

'*****************************************************************
' FUNCTION: GetScreenResolution()
'
' PURPOSE:
'   To determine the current screen size or resolution.
'
' RETURN:
'   The current screen resolution. Typically one of the following:
'      640 x 480
'      800 x 600
'     1024 x 768
'
'*****************************************************************
Function GetScreenResolution() As String

   Dim HorizPos As Integer
   Dim R As RECT
#If VBA7 Then
   Dim hWnd As LongPtr
#Else
   Dim hWnd As Long
#End If
   Dim RETVAL As Long

   hWnd = GetDesktopWindow()
   RETVAL = GetWindowRect(hWnd, R)
   GetScreenResolution = (R.x2 - R.x1) & "x" & (R.Y2 - R.Y1)

HorizPos = CInt(R.x2 - R.x1)

'Call AccessMoveSize(HorizPos, 0, 0, 0)

End Function

Function GetFormSize(Optional ByVal frm As Form) As String
#If VBA7 Then
    Dim H As LongPtr
#Else
    Dim H As Long
#End If
    Dim R As RECT
    Dim RETVAL As Long
    If frm Is Nothing Then
        H = Application.hWndAccessApp
    Else
        H = frm.hWnd
    End If
    RETVAL = GetWindowRect(H, R)
    GetFormSize = (R.x2 - R.x1) & "x" & (R.Y2 - R.Y1)
End Function





Function GetX(RESOL As String) As Integer
    GetX = Left(RESOL, InStr(1, RESOL, "x") - 1)
End Function
