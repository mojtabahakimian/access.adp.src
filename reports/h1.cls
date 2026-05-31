Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Compare Database
Option Explicit

Private Sub Detail_Format(CANCEL As Integer, FormatCount As Integer)
                 On Error Resume Next
Me.FontSize = 12
Me.FontName = Me.CODE.FontName
Me.FontBold = Me.CODE.FontBold
    If Me.TextWidth(KALA) / 567 > 8 Then
        Me.FontSize = 11
        If Me.TextWidth(KALA) / 567 > 8 Then
            Me.FontSize = 10
            If Me.TextWidth(KALA) / 567 > 8 Then
               Me.FontSize = 9
               If Me.TextWidth(KALA) / 567 > 8 Then
                    Me.CODE.FontSize = 8
                    If Me.TextWidth(KALA) / 567 > 8 Then
                         Me.CODE.FontSize = 7
                    Else
                          Me.CODE.FontSize = 8
                    End If
               Else
                     Me.CODE.FontSize = 9
               End If
            Else
                 Me.CODE.FontSize = 10
            End If
        Else
            Me.CODE.FontSize = 11
        End If
    Else
        Me.CODE.FontSize = 12
    End If
End Sub

Private Sub GroupFooter3_Format(CANCEL As Integer, FormatCount As Integer)
  Call Showemza(Me, "ANB_HAVL_FROSHTX", "ANB_HAVL_ANBTX", "ANB_HAVL_MODIRTX")
End Sub

Private Sub Report_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
                 On Error Resume Next
Dim rst As New ADODB.Recordset
  If MID(Forms![BASEKNOW]![OPTIONSS], 18, 1) = "5" Then
    If IsLoaded("head_lst_havl") Then
        rst.Open "SELECT     SUM(dbo.STUF_DEF.VAZN * dbo.INVO_LST.MEGHk) AS Weight FROM   dbo.INVO_LST INNER JOIN   dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE WHERE     (dbo.INVO_LST.TAG = 2) AND (dbo.INVO_LST.NUMBER = " & [Forms]![head_lst_havl]![NUMBER] & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    Else
        rst.Open "SELECT     SUM(dbo.STUF_DEF.VAZN * dbo.INVO_LST.MEGHk) AS Weight FROM   dbo.INVO_LST INNER JOIN   dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE WHERE     (dbo.INVO_LST.TAG = 2) AND (dbo.INVO_LST.NUMBER = " & [Forms]![HEAD_LST_FROOSH22]![NUMBER] & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    End If
    If rst.RecordCount > 0 Then
        If Not IsNull(rst.Fields("Weight")) Then
             Me.VAZN.CAPTION = "وزن كل به كيلو : " & Round(rst.Fields("Weight"))
        End If
    End If
   End If
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
