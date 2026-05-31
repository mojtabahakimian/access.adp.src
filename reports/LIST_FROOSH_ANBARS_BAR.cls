Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Compare Database
Option Explicit

Private Sub Report_Close()
    DoCmd.Close acForm, "F_MENU_ANBAR_FRKH_bar"
End Sub

Private Sub Report_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
Dim rst As New ADODB.Recordset
                 On Error Resume Next
 rst.Open "SELECT     SUM(dbo.INVO_LST.MEGHk * dbo.STUF_DEF.VAZN) AS vazn FROM         dbo.INVO_LST INNER JOIN    dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE INNER JOIN  dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG INNER JOIN dbo.OTHER_DTL ON dbo.HEAD_LST.NUMBER = dbo.OTHER_DTL.NUMBER AND dbo.HEAD_LST.TAG = dbo.OTHER_DTL.TAG WHERE     (dbo.OTHER_DTL.DRIVER = N'" & Forms![F_MENU_ANBAR_FRKH_bar]![driver] & "') AND (dbo.HEAD_LST.DATE_N = " & Forms![F_MENU_ANBAR_FRKH_bar]![DT1] & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        If Not IsNull(rst.Fields("VAZN")) Then
             Me.VAZN.CAPTION = "وزن كل به كيلو : " & Round(rst.Fields("VAZN"))
        End If
    End If
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
