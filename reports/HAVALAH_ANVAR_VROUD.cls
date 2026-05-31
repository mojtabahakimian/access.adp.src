Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit

Private Sub Report_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
Dim rst As New ADODB.Recordset
                 On Error Resume Next
    If IsNull([Forms]![BASEKNOW]![ISO_KHAREED]) Or Trim([Forms]![BASEKNOW]![ISO_KHAREED]) = "" Then
       Me.Label171.Visible = False
    End If
    rst.Open "SELECT     SUM(dbo.STUF_DEF.VAZN * dbo.INVO_LST.MEGHk) AS Weight FROM         dbo.INVO_LST INNER JOIN   dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE WHERE     (dbo.INVO_LST.TAG = 1) AND (dbo.INVO_LST.NUMBER = " & Forms![head_lst_rasid]![NUMBER] & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        If Not IsNull(rst.Fields("Weight")) Then
             Me.VAZN.CAPTION = "وزن كل به كيلو : " & Format(Round(rst.Fields("Weight")), "#,##")
        End If
    End If
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
