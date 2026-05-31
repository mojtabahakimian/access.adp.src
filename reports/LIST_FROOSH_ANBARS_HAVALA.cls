Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Compare Database
Option Explicit

Private Sub Report_Close()
    DoCmd.Close acForm, "F_MENU_ANBAR_FRKH_GRP"
End Sub

Private Sub Report_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
Dim rst As New ADODB.Recordset
                 On Error Resume Next
    If Me.OpenArgs = "" Or IsNull(Me.OpenArgs) Then
        Me.RecordSource = "SELECT     INVO_LST.CODE, SUM(INVO_LST.MEGH) AS SMEGH, SUM(INVO_LST.MEGHk) AS SMEGHK, TCOD_VAHEDS.NAMES AS vahed,  TCOD_ANBAR.NAMES AS anbar, STUF_DEF.NAME  AS KALA, SUM(INVO_LST.MABL_K)  AS smabk   FROM         INVO_LST INNER JOIN   STUF_DEF ON INVO_LST.CODE = STUF_DEF.CODE INNER JOIN   TCOD_VAHEDS ON INVO_LST.VAHED_K = TCOD_VAHEDS.CODE INNER JOIN   TCOD_ANBAR ON INVO_LST.ANBAR = TCOD_ANBAR.CODE WHERE     (INVO_LST.TAG = 2) AND (INVO_LST.NUMBER BETWEEN " & Forms![F_MENU_ANBAR_FRKH_GRP]![F1] & " AND " & Forms![F_MENU_ANBAR_FRKH_GRP]![F2] & " ) AND (INVO_LST.ANBAR = " & Forms![F_MENU_ANBAR_FRKH_GRP]![CANBAR] & ")GROUP BY INVO_LST.CODE, TCOD_VAHEDS.NAMES, TCOD_ANBAR.NAMES, STUF_DEF.NAME"
        Me.F1F2.CAPTION = "از فاكتور شماره " & Forms![F_MENU_ANBAR_FRKH_GRP]![F1] & "  تا فاكتور شماره " & Forms![F_MENU_ANBAR_FRKH_GRP]![F2]
        rst.Open "SELECT     SUM(dbo.STUF_DEF.MAX_M * dbo.INVO_LST.MEGHk) AS Weight FROM         dbo.INVO_LST INNER JOIN   dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE INNER JOIN  dbo.TCOD_VAHEDS ON dbo.INVO_LST.VAHED_K = dbo.TCOD_VAHEDS.CODE INNER JOIN  dbo.TCOD_ANBAR ON dbo.INVO_LST.ANBAR = dbo.TCOD_ANBAR.CODE WHERE     (dbo.INVO_LST.TAG = 2) AND (dbo.INVO_LST.NUMBER BETWEEN " & Forms![F_MENU_ANBAR_FRKH_GRP]![F1] & " AND " & Forms![F_MENU_ANBAR_FRKH_GRP]![F2] & " ) AND (INVO_LST.ANBAR = " & Forms![F_MENU_ANBAR_FRKH_GRP]![CANBAR] & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            If Not IsNull(rst.Fields("Weight")) Then
                 Me.VAZN.CAPTION = "وزن كل به كيلو : " & Round(rst.Fields("Weight"))
            End If
        End If
    Else
        Me.RecordSource = "SELECT     INVO_LST.CODE, SUM(INVO_LST.MEGH) AS SMEGH, SUM(INVO_LST.MEGHk) AS SMEGHK, TCOD_VAHEDS.NAMES AS vahed,  TCOD_ANBAR.NAMES AS anbar, STUF_DEF.NAME  AS KALA , SUM(INVO_LST.MABL_K)  AS smabk FROM         INVO_LST INNER JOIN   STUF_DEF ON INVO_LST.CODE = STUF_DEF.CODE INNER JOIN   TCOD_VAHEDS ON INVO_LST.VAHED_K = TCOD_VAHEDS.CODE INNER JOIN   TCOD_ANBAR ON INVO_LST.ANBAR = TCOD_ANBAR.CODE WHERE     (INVO_LST.TAG = 2) AND (INVO_LST.NUMBER BETWEEN " & Forms![F_MENU_ANBAR_FRKH_GRP]![F1] & " AND " & Forms![F_MENU_ANBAR_FRKH_GRP]![F2] & " OR " & Me.OpenArgs & " ) AND (INVO_LST.ANBAR = " & Forms![F_MENU_ANBAR_FRKH_GRP]![CANBAR] & ")GROUP BY INVO_LST.CODE, TCOD_VAHEDS.NAMES, TCOD_ANBAR.NAMES, STUF_DEF.NAME "
        Me.F1F2.CAPTION = "از فاكتور " & Forms![F_MENU_ANBAR_FRKH_GRP]![F1] & "  تا فاكتور " & Forms![F_MENU_ANBAR_FRKH_GRP]![F2] & " و " & Forms![F_MENU_ANBAR_FRKH_GRP]![VF]
         rst.Open "SELECT     SUM(dbo.STUF_DEF.MAX_M  * dbo.INVO_LST.MEGHk) AS Weight FROM         dbo.INVO_LST INNER JOIN   dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE INNER JOIN  dbo.TCOD_VAHEDS ON dbo.INVO_LST.VAHED_K = dbo.TCOD_VAHEDS.CODE INNER JOIN  dbo.TCOD_ANBAR ON dbo.INVO_LST.ANBAR = dbo.TCOD_ANBAR.CODE WHERE     (dbo.INVO_LST.TAG = 2) AND (dbo.INVO_LST.NUMBER BETWEEN " & Forms![F_MENU_ANBAR_FRKH_GRP]![F1] & " AND " & Forms![F_MENU_ANBAR_FRKH_GRP]![F2] & " OR " & Me.OpenArgs & " ) AND (INVO_LST.ANBAR = " & Forms![F_MENU_ANBAR_FRKH_GRP]![CANBAR] & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            If Not IsNull(rst.Fields("Weight")) Then
                 Me.VAZN.CAPTION = "وزن كل به كيلو : " & Round(rst.Fields("Weight"))
            End If
        End If
   End If
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
