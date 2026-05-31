Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Compare Database
Option Explicit

Private Sub GroupFooter3_Format(CANCEL As Integer, FormatCount As Integer)
Dim rst As New ADODB.Recordset
                 On Error Resume Next
rst.Open "SELECT     SUM(dbo.INVO_LST.MEGHk) AS MEGHkS, SUM(dbo.INVO_LST.MABL_K) AS MABL_KS, SUM(dbo.INVO_LST.N_MOIN) AS TAKH,  SUM(dbo.INVO_LST.IMBAA) AS IMBAA, dbo.HEAD_LST.MABL_HAZ, dbo.HEAD_LST.MBAA, dbo.HEAD_LST.TAKHFIF, SUM(dbo.STUF_DEF.VAZN * dbo.INVO_LST.MEGHk) As WEGHT FROM         dbo.INVO_LST INNER JOIN dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG INNER JOIN    dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE WHERE     (dbo.INVO_LST.NUMBER = " & Forms![HEAD_LST_PISHFROOSH2]![NUMBER] & ") AND (dbo.INVO_LST.TAG = 20) GROUP BY dbo.HEAD_LST.MABL_HAZ, dbo.HEAD_LST.MBAA, dbo.HEAD_LST.TAKHFIF", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
If rst.RecordCount > 0 Then
    Me.JF.CAPTION = Format(rst.Fields("MABL_KS"), "#,###")
    Me.takh.CAPTION = Format(rst.Fields("TAKH"), "#,###")
    Me.MBAA.CAPTION = Format(rst.Fields("IMBAA"), "#,###")
    Me.KHAD.CAPTION = Format(rst.Fields("MABL_HAZ"), "#,###")
    Me.GHABEL.CAPTION = Format(rst.Fields("MABL_KS") + rst.Fields("MABL_HAZ") + rst.Fields("IMBAA") - rst.Fields("TAKH"), "#,###")
    Me.HR.CAPTION = ALPHANUM(rst.Fields("MABL_KS") + rst.Fields("MABL_HAZ") + rst.Fields("IMBAA") - rst.Fields("TAKH")) & " " & "ريال"

End If
Set rst = New ADODB.Recordset
If Not IsNull(Trim(Me.MOLAH)) Then
 If RIGHT(Me.MOLAH, 1) = "." Then
    rst.Open "SELECT     SUM(dbo.STUF_DEF.VAZN * dbo.INVO_LST.MEGHk) AS Weight FROM         dbo.INVO_LST INNER JOIN   dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE WHERE     (dbo.INVO_LST.TAG = 20) AND (dbo.INVO_LST.NUMBER = " & Forms![HEAD_LST_PISHFROOSH2]![NUMBER] & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        If Not IsNull(rst.Fields("Weight")) Then
             Me.VAZN.CAPTION = "وزن كل به كيلو : " & Format(Round(rst.Fields("Weight")), "#,##")
        End If
    End If
 End If
 End If

End Sub

Private Sub GroupHeader2_Format(CANCEL As Integer, FormatCount As Integer)
 Me.CAPTION = Me.NUMBER
End Sub

Private Sub Report_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
                      On Error Resume Next
   Me.Label202.CAPTION = "%ماليات و عوارض:"
     If [Forms]![HEAD_LST_PISHFROOSH2]![TAKHFIF] > 0 Then
       Me.Label155.CAPTION = Round([Forms]![HEAD_LST_PISHFROOSH2]![TAKHFIF] / [Forms]![HEAD_LST_PISHFROOSH2]![Text59] * 100) & "% تخفيف :"
     End If
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
