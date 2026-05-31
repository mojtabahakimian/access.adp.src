Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Compare Database
Option Explicit

Private Sub GroupFooter3_Format(CANCEL As Integer, FormatCount As Integer)
 Dim rst As New ADODB.Recordset
                 On Error Resume Next
If Not IsNull(Me.MOLAH) Then
 If RIGHT(Me.MOLAH, 1) = "." Then
    rst.Open "SELECT     SUM(dbo.STUF_DEF.VAZN * dbo.INVO_LST.MEGHk) AS Weight FROM         dbo.INVO_LST INNER JOIN   dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE WHERE     (dbo.INVO_LST.TAG = 9) AND (dbo.INVO_LST.NUMBER = " & Forms![havalah_enter]![NUMBER] & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        If Not IsNull(rst.Fields("Weight")) Then
             Me.VAZN.CAPTION = "وزن كل به كيلو : " & Format(Round(rst.Fields("Weight")), "#,##")
        End If
    End If
 End If
 End If

End Sub

Private Sub GroupHeader2_Format(CANCEL As Integer, FormatCount As Integer)
                    On Error Resume Next
   If IsNull(FNUMCO) Or FNUMCO = "" Then
        FNUMCO.Visible = False
        Me.Label198.Visible = False
   End If
End Sub

Private Sub Report_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
                    On Error Resume Next
    If IsNull([Forms]![BASEKNOW]![ISO_TOLID]) Or Trim([Forms]![BASEKNOW]![ISO_TOLID]) = "" Then
       Me.Label171.Visible = False
    End If
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
