Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Compare Database
Option Explicit

Private Sub Detail_Format(CANCEL As Integer, FormatCount As Integer)
Me.ServerFilter = "(USER_NAME = N'" & UCurrentUser() & "')"
                 On Error Resume Next
Dim tm As Single
Me.FontSize = Me.CODE1.FontSize
Me.FontName = Me.CODE1.FontName
Me.FontBold = Me.CODE1.FontBold
    If Me.TextWidth(KALA) / 567 > 7.8 Then
        Me.FontSize = 9
        If Me.TextWidth(KALA) / 567 > 7.8 Then
            Me.FontSize = 8
            If Me.TextWidth(KALA) / 567 > 7.8 Then
               Me.FontSize = 7
               If Me.TextWidth(KALA) / 567 > 7.8 Then
                    Me.CODE1.FontSize = 6
               Else
                     Me.CODE1.FontSize = 7
               End If
            Else
                 Me.CODE1.FontSize = 8
            End If
        Else
            Me.CODE1.FontSize = 9
        End If
    Else
        Me.CODE1.FontSize = 10
    End If
If MEGHk - Fix(MEGHk) > 0 Then
    tm = MEGHk - Fix(MEGHk)
    Me.MEGHk.DecimalPlaces = Len(tm) - 2
Else
    Me.MEGHk.DecimalPlaces = 0
End If
If MEGH - Fix(MEGH) > 0 Then
    tm = MEGH - Fix(MEGH)
    Me.MEGH.DecimalPlaces = Len(tm) - 2
Else
    Me.MEGH.DecimalPlaces = 0
End If

Me.FontSize = Me.MEGHk.FontSize
Me.FontName = Me.MEGHk.FontName
Me.FontBold = Me.MEGHk.FontBold
    If Me.TextWidth(MEGHk) / 567 > 2.23 Then
        Me.FontSize = 11
        If Me.TextWidth(MEGHk) / 567 > 2.23 Then
            Me.FontSize = 10
            If Me.TextWidth(MEGHk) / 567 > 2.23 Then
               Me.FontSize = 9
               If Me.TextWidth(MEGHk) / 567 > 2.23 Then
                    Me.MEGHk.FontSize = 8
               Else
                     Me.MEGHk.FontSize = 9
               End If
            Else
                 Me.MEGHk.FontSize = 10
            End If
        Else
            Me.MEGHk.FontSize = 11
        End If
    Else
        Me.MEGHk.FontSize = 12
    End If
Me.FontSize = Me.MEGH.FontSize
Me.FontName = Me.MEGH.FontName
Me.FontBold = Me.MEGH.FontBold
    If Me.TextWidth(MEGH) / 567 > 1.24 Then
        Me.FontSize = 11
        If Me.TextWidth(MEGH) / 567 > 1.24 Then
            Me.FontSize = 10
            If Me.TextWidth(MEGH) / 567 > 1.24 Then
               Me.FontSize = 9
               If Me.TextWidth(MEGH) / 567 > 1.24 Then
                    Me.MEGH.FontSize = 8
               Else
                     Me.MEGH.FontSize = 9
               End If
            Else
                 Me.MEGH.FontSize = 10
            End If
        Else
            Me.MEGH.FontSize = 11
        End If
    Else
        Me.MEGH.FontSize = 12
    End If

End Sub

Private Sub GroupFooter3_Format(CANCEL As Integer, FormatCount As Integer)
Dim rst As New ADODB.Recordset
                 On Error Resume Next
If Not IsNull(Me.MOLAH) Then
 If RIGHT(Me.MOLAH, 1) = "." Then
     rst.Open "SELECT     SUM(dbo.STUF_DEF.VAZN * dbo.ORDR_LST.MEGHk) AS Weight FROM         dbo.ORDR_LST INNER JOIN   dbo.STUF_DEF ON dbo.ORDR_LST.CODE = dbo.STUF_DEF.CODE WHERE      (dbo.ORDR_LST.ID = " & Forms![ORDR_HED]![ID] & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        If Not IsNull(rst.Fields("Weight")) And rst.Fields("Weight") > 0 Then
             Me.VAZN.CAPTION = "وزن كل به كيلو : " & Format(Round(rst.Fields("Weight")), "#,##")
        End If
    End If
 End If
 End If
End Sub

Private Sub Report_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
