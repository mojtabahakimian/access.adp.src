Attribute VB_Name = "api_TwipsPerPixels"
Option Compare Database
Option Explicit

Private Declare Function GetDC Lib "user32" ( _
                ByVal hWnd As Long) As Long

Private Declare Function ReleaseDC Lib "user32" ( _
                                                    ByVal hWnd As Long, _
                                                    ByVal hdc As Long) As Long

Private Declare Function GetDeviceCaps Lib "gdi32" ( _
                ByVal hdc As Long, _
                ByVal nIndex As Long) As Long


'// mod constants
Private Const LogPixelsX = 88
Private Const LOGPIXELSY = 90
Private Const TWIPSPERINCH = 1440



Public Function TwipsPerPixelX(PixelWidthX As Long) As Integer
'// general declarations
Dim hdc As Long
Dim PixelsPerInchX As Long

    '// get device handle
    hdc = GetDC(0)
    '// get and return pixels
    PixelsPerInchX = GetDeviceCaps(hdc, LogPixelsX)
    '// calculate and return measurements in twips
    TwipsPerPixelX = (PixelWidthX / PixelsPerInchX) * TWIPSPERINCH
    '// clean up
    ReleaseDC 0, hdc

End Function

Public Function TwipsPerPixelY(PixelHeightY As Long) As Integer
'// general declarations
Dim hdc As Long
Dim PixelsPerInchY As Long

    '// get device handle
    hdc = GetDC(0)
    '// get and return pixels
    PixelsPerInchY = GetDeviceCaps(hdc, LOGPIXELSY)
    '// calculate and return measurements in twips
    TwipsPerPixelY = (PixelHeightY / PixelsPerInchY) * TWIPSPERINCH
    '// clean up
    ReleaseDC 0, hdc

End Function



Public Function PixelsPerTwipsX(TwipsWidthX As Long) As Integer
'// general declarations
Dim hdc As Long
Dim lngPixelsPerInch As Long

    '// get device handle
    hdc = GetDC(0)
    '// get and return pixels
    lngPixelsPerInch = GetDeviceCaps(hdc, LogPixelsX)
    '// calculate and return measurements in pixels
    PixelsPerTwipsX = (TwipsWidthX / TWIPSPERINCH) * lngPixelsPerInch
    '// clean up
    ReleaseDC 0, hdc

End Function

Public Function PixelsPerTwipsY(TwipsHeightY As Long) As Integer
'// general declarations
Dim hdc As Long
Dim lngPixelsPerInch As Long

    '// get device handle
    hdc = GetDC(0)
    '// get and return pixels
    lngPixelsPerInch = GetDeviceCaps(hdc, LOGPIXELSY)
    '// calculate and return measurements in pixels
    PixelsPerTwipsY = (TwipsHeightY / TWIPSPERINCH) * lngPixelsPerInch
    '// clean up
    ReleaseDC 0, hdc

End Function
