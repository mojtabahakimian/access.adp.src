Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub CODE_DblClick(CANCEL As Integer)
   DoCmd.OpenForm "TOLID_COUNTALL_MP1", acFormDS
End Sub

Private Sub Form_Close()
                   On Error Resume Next
  DoCmd.Close acForm, "F_MENU_DATE"
End Sub


Private Sub Form_Load()
   Dim rst As New ADODB.Recordset, VS As String, svv As String
 If IsNull(Me.OpenArgs) Then
    Err.Clear
    svv = ""
    Me.RecordSource = "SELECT     dbo.INVO_LST.CODE, SUM(dbo.INVO_LST.MEGHk) AS MEGHT, dbo.STUF_DEF.NAME, dbo.TCOD_STUFGROUP.NAMES, dbo.TCOD_VAHEDS.NAMES AS vahed, dbo.STUF_DEF.N_FANI FROM         dbo.HEAD_LST INNER JOIN     dbo.INVO_LST ON dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER AND dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG INNER JOIN   dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE INNER JOIN   dbo.TCOD_STUFGROUP ON dbo.STUF_DEF.RADAH = dbo.TCOD_STUFGROUP.CODE INNER JOIN     dbo.TCOD_VAHEDS ON dbo.STUF_DEF.VAHED = dbo.TCOD_VAHEDS.CODE WHERE     (dbo.HEAD_LST.TAG = 9) AND (dbo.HEAD_LST.DATE_N >= " & Forms![F_MENU_DATE]![DT1] & ") AND (dbo.HEAD_LST.DATE_N <= " & Forms![F_MENU_DATE]![DT2] & ")  GROUP BY dbo.INVO_LST.CODE, dbo.STUF_DEF.NAME, dbo.TCOD_STUFGROUP.NAMES, dbo.TCOD_VAHEDS.NAMES, dbo.STUF_DEF.N_FANI "
    If LETSGO("chartfilter") Then ' اگر محدود به  زير مجموعه خودش هست
        VS = UserOnChart(Forms![BASEKNOW]![USERCOD])
    End If

    'محدود به زير مجموعه خودش باشد
    If LETSGO("chartfilter") Then
        If svv = "" Then
           svv = VS
        Else
           svv = svv & " and  (" & VS & ")"
        End If
        Me.RecordSource = "SELECT     dbo.INVO_LST.CODE, SUM(dbo.INVO_LST.MEGHk) AS MEGHT, dbo.STUF_DEF.NAME, dbo.TCOD_STUFGROUP.NAMES, dbo.TCOD_VAHEDS.NAMES AS vahed, dbo.STUF_DEF.N_FANI FROM         dbo.HEAD_LST INNER JOIN dbo.INVO_LST ON dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER AND dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG INNER JOIN dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE INNER JOIN dbo.TCOD_STUFGROUP ON dbo.STUF_DEF.RADAH = dbo.TCOD_STUFGROUP.CODE INNER JOIN dbo.TCOD_VAHEDS ON dbo.STUF_DEF.VAHED = dbo.TCOD_VAHEDS.CODE WHERE     (dbo.HEAD_LST.TAG = 9) AND (dbo.HEAD_LST.DATE_N >= " & Forms![F_MENU_DATE]![DT1] & ") AND (dbo.HEAD_LST.DATE_N <= " & Forms![F_MENU_DATE]![DT2] & ") AND " & svv & " GROUP BY dbo.INVO_LST.CODE, dbo.STUF_DEF.NAME, dbo.TCOD_STUFGROUP.NAMES, dbo.TCOD_VAHEDS.NAMES, dbo.STUF_DEF.N_FANI"
  End If
    'محدود به دپارتمان خودش باشد
    If LETSGO("DEPEMAL") Then
        If svv = "" Then
           ' Me.RecordSource = "SELECT TOP (100) PERCENT HEAD_LST.* FROM HEAD_LST WHERE (TAG = 9) and (DEPATMAN = " & Forms![Default]![TFSAZMAN] & ") ORDER BY NUMBER "
           Me.RecordSource = "SELECT     dbo.INVO_LST.CODE, SUM(dbo.INVO_LST.MEGHk) AS MEGHT, dbo.STUF_DEF.NAME, dbo.TCOD_STUFGROUP.NAMES, dbo.TCOD_VAHEDS.NAMES AS vahed, dbo.STUF_DEF.N_FANI FROM         dbo.HEAD_LST INNER JOIN     dbo.INVO_LST ON dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER AND dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG INNER JOIN   dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE INNER JOIN   dbo.TCOD_STUFGROUP ON dbo.STUF_DEF.RADAH = dbo.TCOD_STUFGROUP.CODE INNER JOIN     dbo.TCOD_VAHEDS ON dbo.STUF_DEF.VAHED = dbo.TCOD_VAHEDS.CODE WHERE     (dbo.HEAD_LST.TAG = 9) AND (dbo.HEAD_LST.DATE_N >= " & Forms![F_MENU_DATE]![DT1] & ") AND (dbo.HEAD_LST.DATE_N <= " & Forms![F_MENU_DATE]![DT2] & ") AND (dbo.HEAD_LST.DEPATMAN = " & Forms![Default]![TFSAZMAN] & ") GROUP BY dbo.INVO_LST.CODE, dbo.STUF_DEF.NAME, dbo.TCOD_STUFGROUP.NAMES, dbo.TCOD_VAHEDS.NAMES, dbo.STUF_DEF.N_FANI "
        Else
          svv = Replace(svv, "USER_NAME", "dbo.HEAD_LST.USER_NAME")
        '   Me.RecordSource = "SELECT TOP (100) PERCENT HEAD_LST.* FROM HEAD_LST WHERE (TAG = 9) and (DEPATMAN = " & Forms![Default]![TFSAZMAN] & ") and " & svv & "ORDER BY NUMBER "
        '   Me.RecordSource = "SELECT     dbo.INVO_LST.CODE, SUM(dbo.INVO_LST.MEGHk) AS MEGHT, dbo.STUF_DEF.NAME, dbo.TCOD_STUFGROUP.NAMES, dbo.TCOD_VAHEDS.NAMES AS vahed, dbo.STUF_DEF.N_FANI FROM         dbo.HEAD_LST INNER JOIN     dbo.INVO_LST ON dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER AND dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG INNER JOIN   dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE INNER JOIN   dbo.TCOD_STUFGROUP ON dbo.STUF_DEF.RADAH = dbo.TCOD_STUFGROUP.CODE INNER JOIN     dbo.TCOD_VAHEDS ON dbo.STUF_DEF.VAHED = dbo.TCOD_VAHEDS.CODE WHERE     (dbo.HEAD_LST.TAG = 9) AND (dbo.HEAD_LST.DATE_N >= " & Forms![F_MENU_DATE]![DT1] & ") AND (dbo.HEAD_LST.DATE_N <= " & Forms![F_MENU_DATE]![DT2] & ") AND (dbo.HEAD_LST.DEPATMAN = " & Forms![Default]![TFSAZMAN] & ") and (" & svv & ") GROUP BY dbo.INVO_LST.CODE, dbo.STUF_DEF.NAME, dbo.TCOD_STUFGROUP.NAMES, dbo.TCOD_VAHEDS.NAMES, dbo.STUF_DEF.N_FANI "
        Me.RecordSource = "SELECT     dbo.INVO_LST.CODE, SUM(dbo.INVO_LST.MEGHk) AS MEGHT, dbo.STUF_DEF.NAME, dbo.TCOD_STUFGROUP.NAMES, dbo.TCOD_VAHEDS.NAMES AS vahed, dbo.STUF_DEF.N_FANI FROM         dbo.HEAD_LST INNER JOIN dbo.INVO_LST ON dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER AND dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG INNER JOIN dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE INNER JOIN dbo.TCOD_STUFGROUP ON dbo.STUF_DEF.RADAH = dbo.TCOD_STUFGROUP.CODE INNER JOIN dbo.TCOD_VAHEDS ON dbo.STUF_DEF.VAHED = dbo.TCOD_VAHEDS.CODE WHERE     (dbo.HEAD_LST.TAG = 9) AND (dbo.HEAD_LST.DATE_N >= " & Forms![F_MENU_DATE]![DT1] & ") AND (dbo.HEAD_LST.DATE_N <= " & Forms![F_MENU_DATE]![DT2] & ") AND (dbo.HEAD_LST.DEPATMAN = " & Forms![Default]![TFSAZMAN] & ") AND " & svv & " GROUP BY dbo.INVO_LST.CODE, dbo.STUF_DEF.NAME, dbo.TCOD_STUFGROUP.NAMES, dbo.TCOD_VAHEDS.NAMES, dbo.STUF_DEF.N_FANI"
        End If
    End If
End If
Me.Refresh
End Sub

Private Sub Form_Open(CANCEL As Integer)
   Dim rst As New ADODB.Recordset, VS As String, svv As String
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
   Me.NAMEE.ColumnWidth = -2


End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
