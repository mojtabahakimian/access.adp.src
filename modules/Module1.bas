Attribute VB_Name = "Module1"
Option Compare Database
Option Explicit

#If VBA7 Then
Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
                ByVal hWnd As LongPtr, _
                    ByVal nIndex As Long) As LongPtr

Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
                ByVal hWnd As LongPtr, _
                    ByVal nIndex As Long, _
                        ByVal dwNewLong As LongPtr) As LongPtr

Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" ( _
                ByVal hWnd As LongPtr, _
                ByVal crKey As Long, _
                ByVal bAlpha As Byte, _
                ByVal dwflags As Long) As Long
#Else
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
                ByVal hWnd As Long, _
                ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
                ByVal hWnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) As Long

Private Declare Function SetLayeredWindowAttributes Lib "user32" ( _
                ByVal hWnd As Long, _
                ByVal crKey As Long, _
                ByVal bAlpha As Byte, _
                ByVal dwflags As Long) As Long
#End If
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2


#If VBA7 Then
Public Function TransparentFormWithVisibleControl(f As Access.Form, Optional Color As Variant = vbCyan) As LongPtr
    'f.Detail.BackColor = vbMagenta
        Dim lhnd As LongPtr
#Else
Public Function TransparentFormWithVisibleControl(f As Access.Form, Optional Color As Variant = vbCyan) As Long
        Dim lhnd As Long
#End If
    lhnd = SetWindowLong(f.hWnd, GWL_EXSTYLE, GetWindowLong(f.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
    SetLayeredWindowAttributes f.hWnd, Color, 0&, LWA_COLORKEY
    TransparentFormWithVisibleControl = lhnd
End Function
