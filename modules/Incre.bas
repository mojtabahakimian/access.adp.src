Attribute VB_Name = "Incre"
Option Compare Database
Option Explicit
Dim COD As Long, ss As String
Function INCR(i As String) As Long
    COD = COD + 1
    INCR = COD
End Function
Function INCRC(i As String) As Long
    If ss <> i Then
        COD = COD + 1
    End If
    ss = i
    INCRC = COD
End Function
