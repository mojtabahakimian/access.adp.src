Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit


Private Sub Detail_Format(CANCEL As Integer, FormatCount As Integer)
                 On Error Resume Next
Me.FontSize = 10
Me.FontName = Me.KALA.FontName
Me.FontBold = Me.KALA.FontBold
    If Me.TextWidth(KALA) / 567 > 6.7 Then
        Me.FontSize = 9
        If Me.TextWidth(KALA) / 567 > 6.7 Then
            Me.FontSize = 8
            If Me.TextWidth(KALA) / 567 > 6.7 Then
               Me.FontSize = 7
               If Me.TextWidth(KALA) / 567 > 6.7 Then
                    Me.KALA.FontSize = 6
                    If Me.TextWidth(KALA) / 567 > 6.7 Then
                         Me.KALA.FontSize = 5
                    Else
                          Me.KALA.FontSize = 6
                    End If
               Else
                     Me.KALA.FontSize = 7
               End If
            Else
                 Me.KALA.FontSize = 8
            End If
        Else
            Me.KALA.FontSize = 9
        End If
    Else
        Me.KALA.FontSize = 10
    End If

End Sub

Private Sub GroupFooter3_Format(CANCEL As Integer, FormatCount As Integer)
  Call Showemza(Me, "SGN0134TX", "SGN0234TX", "SGN0334TX")

End Sub

Private Sub Report_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
Dim rst As New ADODB.Recordset
                 On Error Resume Next
  '  If IsNull([Forms]![BASEKNOW]![ISO_FROOSH]) Or Trim([Forms]![BASEKNOW]![ISO_FROOSH]) = "" Then
    '   Me.Label195.Visible = False
  '  End If
    rst.Open "SELECT     SUM(dbo.STUF_DEF.VAZN * dbo.INVO_LST.MEGHk) AS Weight FROM         dbo.INVO_LST INNER JOIN   dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE WHERE     (dbo.INVO_LST.TAG = 2) AND (dbo.INVO_LST.NUMBER = " & Forms![head_lst_havl]![NUMBER] & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        If Not IsNull(rst.Fields("Weight")) Then
             Me.VAZN.CAPTION = "وزن كل به كيلو : " & Format(Round(rst.Fields("Weight")), "#,##")
        End If
    End If
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
