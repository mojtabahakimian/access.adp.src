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
                    Me.KALA.FontSize = 6.7
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
Dim rst As New ADODB.Recordset
                 On Error Resume Next
If Not IsNull(Me.MOLAH) Then
 If RIGHT(Me.MOLAH, 1) = "." Then
    rst.Open "SELECT     SUM(dbo.STUF_DEF.VAZN * dbo.INVO_LST.MEGHk) AS Weight FROM         dbo.INVO_LST INNER JOIN   dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE WHERE     (dbo.INVO_LST.TAG = 11) AND (dbo.INVO_LST.NUMBER = " & Forms![havalah_exit_sayer]![NUMBER] & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        If Not IsNull(rst.Fields("Weight")) Then
             Me.VAZN.CAPTION = "وزن كل به كيلو : " & Format(Round(rst.Fields("Weight")), "#,##")
             Me.VAZN.Visible = True
        End If
    End If
 End If
 End If
  Call Showemza(Me, "ANB_KHOROGS_MODIRTTX", "ANB_KHOROGS_ANBTX", "ANB_KHOROGS_ANBTX")
End Sub

Private Sub Report_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
                      On Error Resume Next
    If IsNull([Forms]![BASEKNOW]![ISO_MAVAD]) Or Trim([Forms]![BASEKNOW]![ISO_MAVAD]) = "" Then
       Me.Label171.Visible = False
    End If
    If MID(Forms![BASEKNOW]![OPTIONSS], 30, 1) <> "5" Then
       N_RASID.RowSource = "SELECT  dbo.HEAD_MANF.FNUMB, ISNULL(dbo.HEAD_MANF.NAMES, dbo.STUF_DEF.NAME) AS NAM, dbo.HEAD_MANF.FNUMB AS Expr1,  dbo.DTL_MANF.CODE FROM         dbo.STUF_DEF RIGHT OUTER JOIN dbo.HEAD_MANF ON dbo.STUF_DEF.CODE = dbo.HEAD_MANF.CODE LEFT OUTER JOIN dbo.DTL_MANF ON dbo.HEAD_MANF.FNUMB = dbo.DTL_MANF.FNUMB WHERE     (dbo.DTL_MANF.CODE IS NULL)"
       N_RASID.Requery
    End If
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
