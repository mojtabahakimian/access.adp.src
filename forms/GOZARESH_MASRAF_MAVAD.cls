Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Close()
   DoCmd.Close acForm, "F_MENU_DATE"
End Sub

Private Sub Form_Open(CANCEL As Integer)
 On Error Resume Next
    Dim itemm
    For Each itemm In Me.Form
       Me(itemm.NAME).ColumnWidth = -2
    Next
   Dim VS As String, svv As String, sql As String
  Err.Clear
    CurrentProject.Connection.Execute "IF EXISTS (SELECT * FROM sys.objects WHERE type = 'P' AND name = 'sp_AMAR_MASRAF') DROP PROCEDURE [dbo].[sp_AMAR_MASRAF]"
    sql = "CREATE PROCEDURE [dbo].[sp_AMAR_MASRAF] @FROMDATE BIGINT, @TODATE BIGINT, @KONTROL INT AS BEGIN WITH MASRAF_SUB_CTE AS ( SELECT dbo.INVO_LST.CODE, SUM(dbo.INVO_LST.MEGHk) AS MEGH, SUM(dbo.INVO_LST.MABL_K) AS MABL, dbo.STUF_DEF.NAME AS KALA, dbo.TCOD_VAHEDS.NAMES AS VAHED, dbo.TCOD_STUFGROUP.NAMES AS GRP, dbo.mabstandard.price, SUM(dbo.INVO_LST.MEGHk * dbo.mabstandard.price) AS mabst, MASRAFST.MABLST, (MASRAFST.MABLST / dbo.mabstandard.price) AS MEGHST FROM dbo.INVO_LST INNER JOIN dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG INNER JOIN dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE INNER JOIN dbo.TCOD_VAHEDS ON dbo.STUF_DEF.VAHED = dbo.TCOD_VAHEDS.CODE INNER JOIN dbo.TCOD_STUFGROUP ON dbo.STUF_DEF.RADAH = dbo.TCOD_STUFGROUP.CODE INNER JOIN dbo.mabstandard ON dbo.STUF_DEF.CODE = dbo.mabstandard.CODE INNER JOIN dbo.MASRAFST(@FROMDATE, @TODATE, @KONTROL) AS MASRAFST ON " & _
    "dbo.STUF_DEF.CODE = MASRAFST.HES_T WHERE (dbo.HEAD_LST.TAG = 10 OR dbo.HEAD_LST.TAG = 11) AND (dbo.HEAD_LST.DATE_N >= @FROMDATE) AND (dbo.HEAD_LST.DATE_N <= @TODATE) GROUP BY dbo.INVO_LST.CODE, dbo.STUF_DEF.NAME, dbo.TCOD_VAHEDS.NAMES, dbo.TCOD_STUFGROUP.NAMES, dbo.mabstandard.price, MASRAFST.MABLST ) SELECT CODE, MEGH, MABL, KALA, VAHED, GRP, price, mabst, (mabst - MABL) AS MOGNERKH, MABLST, MEGHST, (MEGHST - MEGH) AS MOGMAS, ((MEGHST - MEGH) * price) AS MOGRMAS FROM MASRAF_SUB_CTE; END"
    DoCmd.RunSQL sql
 If IsNull(Me.OpenArgs) Then
    Err.Clear
    svv = ""
'    Me.RecordSource = "SELECT     CODE, MEGH, MABL, KALA, VAHED, GRP, price, mabst, mabst - MABL AS MOGNERKH, MABLST, MEGHST, MEGHST - MEGH AS MOGMAS, (MEGHST - MEGH) * price AS MOGRMAS FROM         dbo.AMAR_MASRAF_SUB(" & Forms![F_MENU_DATE]![DT1] & ", " & Forms![F_MENU_DATE]![DT1] & ", " & Forms![BASEKNOW]![CONKAL] & ") AMAR_MASRAF_SUB"
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
    End If
    'محدود به دپارتمان خودش باشد
    If LETSGO("DEPEMAL") Then
        If svv = "" Then
            CurrentProject.Connection.Execute "IF EXISTS (SELECT * FROM sys.objects WHERE type = 'P' AND name = 'sp_AMAR_MASRAF') DROP PROCEDURE [dbo].[sp_AMAR_MASRAF]"
            sql = "CREATE PROCEDURE [dbo].[sp_AMAR_MASRAF] @FROMDATE BIGINT, @TODATE BIGINT, @KONTROL INT AS BEGIN WITH MASRAF_SUB_CTE AS ( SELECT dbo.INVO_LST.CODE, SUM(dbo.INVO_LST.MEGHk) AS MEGH, SUM(dbo.INVO_LST.MABL_K) AS MABL, dbo.STUF_DEF.NAME AS KALA, dbo.TCOD_VAHEDS.NAMES AS VAHED, dbo.TCOD_STUFGROUP.NAMES AS GRP, dbo.mabstandard.price, SUM(dbo.INVO_LST.MEGHk * dbo.mabstandard.price) AS mabst, MASRAFST.MABLST, (MASRAFST.MABLST / dbo.mabstandard.price) AS MEGHST FROM dbo.INVO_LST INNER JOIN dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG INNER JOIN dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE INNER JOIN dbo.TCOD_VAHEDS ON dbo.STUF_DEF.VAHED = dbo.TCOD_VAHEDS.CODE INNER JOIN dbo.TCOD_STUFGROUP ON dbo.STUF_DEF.RADAH = dbo.TCOD_STUFGROUP.CODE INNER JOIN dbo.mabstandard ON dbo.STUF_DEF.CODE = dbo.mabstandard.CODE INNER JOIN dbo.MASRAFST(@FROMDATE, @TODATE, @KONTROL) AS MASRAFST ON " & _
            "dbo.STUF_DEF.CODE = MASRAFST.HES_T WHERE (dbo.HEAD_LST.TAG = 10 OR dbo.HEAD_LST.TAG = 11) AND (dbo.HEAD_LST.DATE_N >= @FROMDATE) AND (dbo.HEAD_LST.DATE_N <= @TODATE) AND (dbo.HEAD_LST.DEPATMAN = " & Forms![Default]![TFSAZMAN] & ") GROUP BY dbo.INVO_LST.CODE, dbo.STUF_DEF.NAME, dbo.TCOD_VAHEDS.NAMES, dbo.TCOD_STUFGROUP.NAMES, dbo.mabstandard.price, MASRAFST.MABLST ) SELECT CODE, MEGH, MABL, KALA, VAHED, GRP, price, mabst, (mabst - MABL) AS MOGNERKH, MABLST, MEGHST, (MEGHST - MEGH) AS MOGMAS, ((MEGHST - MEGH) * price) AS MOGRMAS FROM MASRAF_SUB_CTE; END"
            DoCmd.RunSQL sql
        Else
          svv = Replace(svv, "USER_NAME", "dbo.HEAD_LST.USER_NAME")
            CurrentProject.Connection.Execute "IF EXISTS (SELECT * FROM sys.objects WHERE type = 'P' AND name = 'sp_AMAR_MASRAF') DROP PROCEDURE [dbo].[sp_AMAR_MASRAF]"
            sql = "CREATE PROCEDURE [dbo].[sp_AMAR_MASRAF] @FROMDATE BIGINT, @TODATE BIGINT, @KONTROL INT AS BEGIN WITH MASRAF_SUB_CTE AS ( SELECT dbo.INVO_LST.CODE, SUM(dbo.INVO_LST.MEGHk) AS MEGH, SUM(dbo.INVO_LST.MABL_K) AS MABL, dbo.STUF_DEF.NAME AS KALA, dbo.TCOD_VAHEDS.NAMES AS VAHED, dbo.TCOD_STUFGROUP.NAMES AS GRP, dbo.mabstandard.price, SUM(dbo.INVO_LST.MEGHk * dbo.mabstandard.price) AS mabst, MASRAFST.MABLST, (MASRAFST.MABLST / dbo.mabstandard.price) AS MEGHST FROM dbo.INVO_LST INNER JOIN dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG INNER JOIN dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE INNER JOIN dbo.TCOD_VAHEDS ON dbo.STUF_DEF.VAHED = dbo.TCOD_VAHEDS.CODE INNER JOIN dbo.TCOD_STUFGROUP ON dbo.STUF_DEF.RADAH = dbo.TCOD_STUFGROUP.CODE INNER JOIN dbo.mabstandard ON dbo.STUF_DEF.CODE = dbo.mabstandard.CODE INNER JOIN dbo.MASRAFST(@FROMDATE, @TODATE, @KONTROL) AS MASRAFST ON " & _
            "dbo.STUF_DEF.CODE = MASRAFST.HES_T WHERE (dbo.HEAD_LST.TAG = 10 OR dbo.HEAD_LST.TAG = 11) AND (dbo.HEAD_LST.DATE_N >= @FROMDATE) AND (dbo.HEAD_LST.DATE_N <= @TODATE) AND (dbo.HEAD_LST.DEPATMAN = " & Forms![Default]![TFSAZMAN] & ") and " & svv & " GROUP BY dbo.INVO_LST.CODE, dbo.STUF_DEF.NAME, dbo.TCOD_VAHEDS.NAMES, dbo.TCOD_STUFGROUP.NAMES, dbo.mabstandard.price, MASRAFST.MABLST ) SELECT CODE, MEGH, MABL, KALA, VAHED, GRP, price, mabst, (mabst - MABL) AS MOGNERKH, MABLST, MEGHST, (MEGHST - MEGH) AS MOGMAS, ((MEGHST - MEGH) * price) AS MOGRMAS FROM MASRAF_SUB_CTE; END"
            DoCmd.RunSQL sql
        End If
    End If
End If
Me.Refresh
End Sub
