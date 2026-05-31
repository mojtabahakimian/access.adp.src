Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Detail_Format(CANCEL As Integer, FormatCount As Integer)
                   On Error Resume Next
   Dim rst As New ADODB.Recordset
   rst.Open "SELECT CODE, VAHED,NESBAT FROM MODULE_D WHERE CODE = '" & Me.CODE & "' AND VAHED = " & Me.VAHED_K, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount > 0 Then
     Me.vahed = rst.Fields("NESBAT")
   End If
End Sub


Private Sub PageHeaderSection_Format(CANCEL As Integer, FormatCount As Integer)
                 On Error Resume Next
        If Left(Me.SHARAYET, 1) <> "." Or Me.SHARAYET = "" Or IsNull(Me.SHARAYET) Then
            Me.DATE_N.Visible = True
        Else
            Me.DATE_N.Visible = False
        End If
End Sub

Private Sub Report_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
Dim rst As New ADODB.Recordset

                 On Error Resume Next
If Me.OpenArgs = 2 Then
  Me.Label208.Visible = True
End If
    If IsNull([Forms]![BASEKNOW]![ISO_FROOSH]) Or Trim([Forms]![BASEKNOW]![ISO_FROOSH]) = "" Then
       Me.Label195.Visible = False
    End If
    rst.Open "SELECT     SUM(dbo.STUF_DEF.VAZN * dbo.INVO_LST.MEGHk) AS Weight FROM         dbo.INVO_LST INNER JOIN   dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE WHERE     (dbo.INVO_LST.TAG = 2) AND (dbo.INVO_LST.NUMBER = " & Forms![head_lst_havl]![NUMBER] & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        If Not IsNull(rst.Fields("Weight")) And rst.Fields("Weight") <> 0 Then
             Me.VAZN.CAPTION = "وزن كل به كيلو : " & Format(Round(rst.Fields("Weight")), "#,##")
        Else
            Me.VAZN.Visible = False
        End If
    End If
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
