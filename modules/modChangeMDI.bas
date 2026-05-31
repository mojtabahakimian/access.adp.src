Attribute VB_Name = "modChangeMDI"
Option Compare Database
Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    RIGHT As Long
    Bottom As Long
End Type

Private Type SIZEL
    cx As Long
    cy As Long
End Type


Private Type RGBQUAD
  rgbBlue As Byte
  rgbGreen As Byte
  rgbRed As Byte
  rgblReserved As Byte
End Type

Private Type BITMAPINFOHEADER '40 bytes
  biSize As Long ' 40
  biWidth As Long ' 32
  biHeight As Long ' 64
  biPlanes As Integer '1
  biBitCount As Integer '1
  biCompression As Long 'ERGBCompression
  biSizeImage As Long
  biXPelsPerMeter As Long
  biYPelsPerMeter As Long
  biClrUsed As Long
  biClrImportant As Long
End Type


Private Type BITMAPINFO
  bmiHeader As BITMAPINFOHEADER
  bmiColors(1) As RGBQUAD
End Type


Private Type BITMAP
  bmType As Long
  bmWidth As Long
  bmHeight As Long
  bmWidthBytes As Long
  bmPlanes As Integer
  bmBitsPixel As Integer
  bmBits As Long
End Type



' Logical Brush (or Pattern)
Private Type LOGBRUSH
        lbStyle As Long
        lbColor As Long
        lbHatch As Long
End Type

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function apiGetDeviceCaps Lib "gdi32" _
Alias "GetDeviceCaps" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Private Declare Function CreateSolidBrush Lib "gdi32" _
 _
(ByVal crColor As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" _
(ByVal hObject As Long) As Long

Private Declare Function GetStockObject Lib "gdi32" _
(ByVal nIndex As Long) As Long


Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long

' Class field offsets for GetClassLong() and GetClassWord()
Private Const GCL_MENUNAME = (-8)
Private Const GCL_HBRBACKGROUND = (-10)
Private Const GCL_HCURSOR = (-12)
Private Const GCL_HICON = (-14)
Private Const GCL_HMODULE = (-16)
Private Const GCL_CBWNDEXTRA = (-18)
Private Const GCL_CBCLSEXTRA = (-20)
Private Const GCL_WNDPROC = (-24)
Private Const GCL_STYLE = (-26)
Private Const GCW_ATOM = (-32)






' Stock Logical Objects
Private Const WHITE_BRUSH = 0
Private Const LTGRAY_BRUSH = 1
Private Const GRAY_BRUSH = 2
Private Const DKGRAY_BRUSH = 3
Private Const BLACK_BRUSH = 4
Private Const NULL_BRUSH = 5
Private Const HOLLOW_BRUSH = NULL_BRUSH


Private Const CLR_INVALID = &HFFFF

' Brush Styles
Private Const BS_SOLID = 0
Private Const BS_NULL = 1
Private Const BS_HOLLOW = BS_NULL
Private Const BS_HATCHED = 2
Private Const BS_PATTERN = 3
Private Const BS_INDEXED = 4
Private Const BS_DIBPATTERN = 5
Private Const BS_DIBPATTERNPT = 6
Private Const BS_PATTERN8X8 = 7
Private Const BS_DIBPATTERN8X8 = 8

'  Hatch Styles
Private Const HS_HORIZONTAL = 0              '  -----
Private Const HS_VERTICAL = 1                '  |||||
Private Const HS_FDIAGONAL = 2               '  \\\\\
Private Const HS_BDIAGONAL = 3               '  /////
Private Const HS_CROSS = 4                   '  +++++
Private Const HS_DIAGCROSS = 5               '  xxxxx
Private Const HS_FDIAGONAL1 = 6
Private Const HS_BDIAGONAL1 = 7
Private Const HS_SOLID = 8
Private Const HS_DENSE1 = 9
Private Const HS_DENSE2 = 10
Private Const HS_DENSE3 = 11
Private Const HS_DENSE4 = 12
Private Const HS_DENSE5 = 13
Private Const HS_DENSE6 = 14
Private Const HS_DENSE7 = 15
Private Const HS_DENSE8 = 16
Private Const HS_NOSHADE = 17
Private Const HS_HALFTONE = 18
Private Const HS_SOLIDCLR = 19
Private Const HS_DITHEREDCLR = 20
Private Const HS_SOLIDTEXTCLR = 21
Private Const HS_DITHEREDTEXTCLR = 22
Private Const HS_SOLIDBKCLR = 23
Private Const HS_DITHEREDBKCLR = 24
Private Const HS_API_MAX = 25

Private Const TITLE = ""
Private Const API_TRUE As Long = 1&
' Handle to original WINDOWCLASS Brush for MDI window
Private prevHBrush As Long
' Handle to our new Brush for the MDI window
Private hBrush As Long

' Handle to MDI window
Private HwndMDI As Long



Public Function SetMDIBackGround(ByVal crColor As Long) As Boolean
' junk var
Dim lngRet As Long

' Window Rect
Dim rc As RECT

' Grab the Stock WHITE brush
'hBrush = GetStockObject(WHITE_BRUSH)
hBrush = CreateSolidBrush(crColor)

' find MDIClient first
HwndMDI = FindWindowEx(Application.hWndAccessApp, 0&, "MDIClient", TITLE)
' Get current dimensions
lngRet = GetWindowRect(HwndMDI, rc)

With rc
.Bottom = .Bottom - .Top
.Top = 0
.RIGHT = .RIGHT - .Left
.Left = 0
End With
prevHBrush = SetClassLong(HwndMDI, GCL_HBRBACKGROUND, hBrush)
' Force a redraw
Call InvalidateRect(HwndMDI, rc, API_TRUE)
SetMDIBackGround = True
End Function


Public Function RestoreMDIBackground() As Boolean
HwndMDI = FindWindowEx(Application.hWndAccessApp, 0&, "MDIClient", TITLE)
Dim lngRet As Long

' Window Rect
Dim rc As RECT

' Get current dimensions
lngRet = GetWindowRect(HwndMDI, rc)

With rc
.Bottom = .Bottom - .Top
.Top = 0
.RIGHT = .RIGHT - .Left
.Left = 0
End With

hBrush = SetClassLong(HwndMDI, GCL_HBRBACKGROUND, prevHBrush)
' Force a redraw
Call InvalidateRect(HwndMDI, rc, API_TRUE)

Call DeleteObject(hBrush)
RestoreMDIBackground = True
End Function
