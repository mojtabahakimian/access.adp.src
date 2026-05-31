Attribute VB_Name = "api_CreateRoundRectRgn"
Option Compare Database
Option Explicit

' =======================================================================================
' API INFO:
'   The CreateRoundRectRgn function creates a rectangular region with rounded corners.
'
' PARAMETERS:
'   nLeftRect
'       [in] Specifies the x-coordinate of the upper-left corner of the region in logical units.
'   nTopRect
'       [in] Specifies the y-coordinate of the upper-left corner of the region in logical units.
'   nRightRect
'       [in] Specifies the x-coordinate of the lower-right corner of the region in logical units.
'   nBottomRect
'       [in] Specifies the y-coordinate of the lower-right corner of the region in logical units.
'   nWidthEllipse
'       [in] Specifies the width of the ellipse used to create the rounded corners in logical units.
'   nHeightEllipse
'       [in] Specifies the height of the ellipse used to create the rounded corners in logical units.
'
' RETURNS:
'   If the function succeeds, the return value is the handle to the region.
'   If the function fails, the return value is NULL.
' =======================================================================================
Private Declare Function CreateRoundRectRgn Lib "gdi32" ( _
                                                            ByVal nLeftRect As Long, _
                                                            ByVal nTopRect As Long, _
                                                            ByVal nRightRect As Long, _
                                                            ByVal nBottomRect As Long, _
                                                            ByVal nWidthEllipse As Long, _
                                                            ByVal nHeightEllipse As Long) As Long


' =======================================================================================
' API INFO:
'   The SetWindowRgn function sets the window region of a window. The window region determines the area
'   within the window where the system permits drawing. The system does not display any portion of a
'   window that lies outside of the window region
'
' PARAMETERS:
'   hWnd
'       [in] Handle to the window whose window region is to be set.
'   hRgn
'       [in] Handle to a region. The function sets the window region of the window to this region.
'            If hRgn is NULL, the function sets the window region to NULL.
'   bRedraw
'       [in] Specifies whether the system redraws the window after setting the window region.
'            If bRedraw is TRUE, the system does so; otherwise, it does not.
'
'           Typically, you set bRedraw to TRUE if the window is visible.
'
' RETURNS:
'   If the function succeeds, the return value is nonzero.
'   If the function fails, the return value is zero.
' =======================================================================================
Private Declare Function SetWindowRgn Lib "user32" ( _
                                                            ByVal hWnd As Long, _
                                                            ByVal hRgn As Long, _
                                                            ByVal bRedraw As Boolean) As Long


Public Function UISetRoundRect( _
                                ByVal UIForm As Form, _
                                ByVal CornersInPixels As Byte, _
                                Optional ByVal TopCornersOnly As Boolean = True) As Boolean

Dim intRight As Integer
Dim intHeight As Integer
Dim hRgn As Long

    With UIForm
        '// convert form's measurements from twips to pixels
        intRight = PixelsPerTwipsX(.WindowWidth)
        intHeight = PixelsPerTwipsY(.WindowHeight)
        '// process ui style based on selection
        If TopCornersOnly Then
            intHeight = intHeight + CornersInPixels
        Else
            intHeight = intHeight + 1
        End If
        '// create new region
        hRgn = CreateRoundRectRgn(0, 0, intRight, intHeight, CornersInPixels, CornersInPixels)
        '// apply new region
        SetWindowRgn .hWnd, hRgn, True
    End With

End Function
