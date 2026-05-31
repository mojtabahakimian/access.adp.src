Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Detail_Format(CANCEL As Integer, FormatCount As Integer)
                 On Error Resume Next
Me.FontSize = 10
Me.FontName = Me.CODE1.FontName
Me.FontBold = Me.CODE1.FontBold
    If Me.TextWidth(KALA) / 567 > 7.3 Then
        Me.FontSize = 9
        If Me.TextWidth(KALA) / 567 > 7.3 Then
            Me.FontSize = 8
            If Me.TextWidth(KALA) / 567 > 7.3 Then
               Me.FontSize = 7
               If Me.TextWidth(KALA) / 567 > 7.3 Then
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

End Sub

Private Sub GroupFooter3_Format(CANCEL As Integer, FormatCount As Integer)
                 On Error Resume Next
    Dim JST As New ADODB.Recordset, JAMF, KH, BAA, TF As Double
    JAMF = 0
    TF = 0
    BAA = 0
    KH = 0
   JST.Open "SELECT     SUM(MABL_K) AS JAMF, SUM(N_MOIN) AS TF, SUM(IMBAA) AS BAA FROM dbo.INVO_LST WHERE (NUMBER = " & [Forms]![HEAD_LST_PISHFROOSH2]![NUMBER] & ") And (TAG = 20)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
        JAMF = JST.Fields("JAMF")
        TF = JST.Fields("TF")
        BAA = JST.Fields("BAA")
   End If
   Set JST = New ADODB.Recordset
   JST.Open "SELECT     SUM(MABL_HAZ) AS KH FROM dbo.HEAD_LST WHERE (NUMBER = " & [Forms]![HEAD_LST_PISHFROOSH2]![NUMBER] & ") And (TAG = 20)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
        KH = JST.Fields("KH")
   End If
   Me.MABK.CAPTION = Format(JAMF, "#,###")
   Me.TF.CAPTION = Format(TF, "#,###")
   Me.KH.CAPTION = Format(KH, "#,###")
   Me.MABAA.CAPTION = Format(BAA, "#,###")
   Me.PAY.CAPTION = Format(JAMF + KH + BAA - TF, "#,###")
   Me.Label112.CAPTION = "مبلغ به حروف: " & ALPHANUM(JAMF + KH + BAA - TF) & " " & "ريال"
                 On Error Resume Next
Dim rst As New ADODB.Recordset
    rst.Open "SELECT     SUM(dbo.STUF_DEF.VAZN * dbo.INVO_LST.MEGHk) AS Weight FROM   dbo.INVO_LST INNER JOIN   dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE WHERE     (dbo.INVO_LST.TAG = 20) AND (dbo.INVO_LST.NUMBER = " & [Forms]![HEAD_LST_PISHFROOSH2]![NUMBER] & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        If Not IsNull(rst.Fields("Weight")) Then
             Me.VAZN.CAPTION = "وزن كل به كيلو : " & Round(rst.Fields("Weight"))
        End If
    End If
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
  Call Showemza(Me, "FFRP_FROOSHTX", "FFRP_ANBTX", "FFRP_HESABTX")
End Sub

Private Sub Report_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
                      On Error Resume Next
     If [Forms]![HEAD_LST_PISHFROOSH2]![TAKHFIF] > 0 Then
    '   Me.Label155.CAPTION = Round([Forms]![HEAD_LST_PISHFROOSH2]![TAKHFIF] / [Forms]![HEAD_LST_PISHFROOSH2]![Text59] * 100) & "% تخفيف :"
     End If
   Me.Label224.CAPTION = "%ماليات و عوارض:"
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
