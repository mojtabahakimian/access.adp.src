Attribute VB_Name = "modTextWidth"
Option Compare Database
Option Explicit

Private Type RECT
        Left As Long
        Top As Long
        RIGHT As Long
        Bottom As Long
End Type


' Declare API functions
Private Declare Function apiCreateFont Lib "gdi32" Alias "CreateFontA" _
(ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, _
ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal s As Long, _
ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, _
ByVal PAF As Long, ByVal f As String) As Long

Private Declare Function apiSelectObject Lib "gdi32" Alias "SelectObject" (ByVal hdc As Long, _
ByVal hObject As Long) As Long

Private Declare Function apiDeleteObject Lib "gdi32" _
  Alias "DeleteObject" (ByVal hObject As Long) As Long

Private Declare Function apiGetDeviceCaps Lib "gdi32" Alias "GetDeviceCaps" _
(ByVal hdc As Long, ByVal nIndex As Long) As Long

Private Declare Function apiMulDiv Lib "kernel32" Alias "MulDiv" (ByVal nNumber As Long, _
ByVal nNumerator As Long, ByVal nDenominator As Long) As Long

Private Declare Function apiCreateIC Lib "gdi32" Alias "CreateICA" _
(ByVal lpDriverName As String, ByVal lpDeviceName As String, _
ByVal lpOutput As String, lpInitData As Any) As Long

Private Declare Function apiGetDC Lib "user32" _
  Alias "GetDC" (ByVal hWnd As Long) As Long

Private Declare Function apiReleaseDC Lib "user32" _
 Alias "ReleaseDC" (ByVal hWnd As Long, _
 ByVal hdc As Long) As Long

Private Declare Function apiDeleteDC Lib "gdi32" _
  Alias "DeleteDC" (ByVal hdc As Long) As Long

Private Declare Function apiDrawText Lib "user32" Alias "DrawTextA" _
(ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, _
lpRect As RECT, ByVal wFormat As Long) As Long


' CONSTANTS
Private Const TWIPSPERINCH = 1440
' Used to ask System for the Logical pixels/inch in Y axis
Private Const LOGPIXELSY = 90

' DrawText() Format Flags
Private Const DT_TOP = &H0
Private Const DT_LEFT = &H0
Private Const DT_SINGLELINE = &H20
Private Const DT_CALCRECT = &H400
Private Const DT_WORDBREAK = &H10






    Public Function fTextWidth(frm As Form, varTemp As Variant) As Long

    'Name                   FUNCTION() fTextWidth
    '
    'Purpose:               Returns Control Width & Height needed to
    '                       display the contents of the DataSheet passed
    '                       to this function. This function
    '                       uses the Control's font attributes to build
    '                       a Font for the required calculations.
    '
    'Version:               1.0
    '
    'Calls:                 Text API stuff. DrawText performs the actual
    '                       calculation to determine Control Width/Height.

    'Returns:               Width of the string in TWIPS.
    '
    'Created by:            Stephen Lebans
    '
    'Credits:               Go ahead...take some! :-)
    '
    'Date:                  November 16, 2000
    '
    'Time:                  11:44:25pm
    '
    'Feedback:              Stephen@lebans.com
    '
    'My Web Page:           www.lebans.com
    '
    'Copyright:             Lebans Holdings Ltd.
    '                       Please feel free to use this code
    '                       without restriction in any application you develop.
    '                       This code may not be resold by itself or as
    '                       part of a collection.
    '
    '
    'Bugs:
    'Not tested enough to tell. Let me know
    'NEEDS ERROR CHECKING!!!!!
    '
    'Enjoy
    'Stephen Lebans

    '***************Code Start***************

    ' Did we get a valid Form in DataSheet view passed to us?
    If frm.CurrentView <> 2 Then
        fTextWidth = -1
      '  Exit Function
    End If

    ' Did we get a valid string passed to us?
    If Len(varTemp & vbNullString) = 0 Then
        fTextWidth = -1
        Exit Function
    End If

    ' Structure for DrawText calc
    Dim sRect As RECT

    ' Handle to Form's window
    Dim hWnd As Long

    ' Reports Device Context
    Dim hdc As Long

    ' Holds the current screen resolution
    Dim lngYdpi As Long

    Dim newfont As Long
    ' Handle to our Font Object we created.
    ' We must destroy it before exiting main function

    Dim oldfont As Long
    ' Device COntext's Font we must Select back into the DC
    ' before we exit this function.

    ' Temporary holder for returns from API calls
    Dim lngRet As Long

    ' Calculate screen Font height
    Dim fheight As Long

    ' Get Controls Parents Window handle
    hWnd = frm.hWnd
    If hWnd = 0 Then Exit Function

    ' retrieve a handle to a display device context (DC)
    ' for the client area of the specified window
    hdc = apiGetDC(hWnd)


    ' Because Access control's do not have a permanent Device Context,
    ' we cannot depend on what we find selected into the DC unless
    ' the Control has the focus. In this case we are simply using the
    ' Control's Font attributes to build our own font in whatever
    ' DC is handy. We must Save this DC's Font so we can restore
    ' the Font when we exit this function. The DataSheet may be different
    ' but since this works I'll stay with this logic for now.

    ' Clear our return value
    lngRet = 0

    ' Temporary Information Context for Screen info.
    Dim lngIC As Long

    ' Modified to allow for different screen resolutions
    ' and printer output. Needed to Calculate Font size
    lngIC = apiCreateIC("DISPLAY", vbNullString, vbNullString, vbNullString)
    If lngIC <> 0 Then
        lngYdpi = apiGetDeviceCaps(lngIC, LOGPIXELSY)
        apiDeleteDC (lngIC)
    Else
        lngYdpi = 120 'Default average value
    End If

    ' Calculate/Convert requested Font Height
    ' into Font's Device Coordinate space
    fheight = apiMulDiv(frm.DatasheetFontHeight, lngYdpi, 72)

    ' We use a negative value to signify
    ' to the CreateFont function that we want a Glyph
    ' outline of this size not a bounding box.
    With frm
        newfont = apiCreateFont(-fheight, 0, _
        0, 0, .DatasheetFontWeight, _
        .DatasheetFontItalic, .DatasheetFontUnderline, _
        0, 0, 0, _
        0, 0, 0, .DatasheetFontName)
    End With

    ' Select the new font into our DC.
    oldfont = apiSelectObject(hdc, newfont)

    ' Use DrawText to Calculate height of Rectangle required to hold
    ' the current contents of the String(variant) passed to this function.
    ' IF we pass a value in the .Right member of our RECT structure
    ' the DrawText will return the value in the .Bottom member required
    ' to fully display the control's contents.

With sRect
    .Left = 0
    .Top = 0
    .Bottom = 0 'ctl.Height / (TWIPSPERINCH / lngYdpi)
    .RIGHT = 0 'ctl.Width / (TWIPSPERINCH / lngYdpi)
    lngRet = apiDrawText(hdc, varTemp, -1, sRect, DT_CALCRECT + DT_TOP + DT_LEFT)

    ' Cleanup
    lngRet = apiSelectObject(hdc, oldfont)
    ' Delete the Font we created
    apiDeleteObject (newfont)

    lngRet = apiReleaseDC(hWnd, hdc)

    ' Convert RECT values to TWIPS
    '.Bottom = .Bottom * (TWIPSPERINCH / lngYdpi)
    .RIGHT = .RIGHT * (TWIPSPERINCH / lngYdpi)

    ' Fudge factor to allow for Left and Right margins.
    ' Add 8 pixels for readability
    .RIGHT = .RIGHT + (TWIPSPERINCH / lngYdpi) * 8
End With


fTextWidth = sRect.RIGHT
End Function
