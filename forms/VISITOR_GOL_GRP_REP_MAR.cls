Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command55_Click()
     DoCmd.OpenReport "Visit_gol_dtl_REP_print_grp", acViewPreview
End Sub

Private Sub FG_AfterUpdate()
    Dim DT1 As Long, DT2 As Long, rst As New ADODB.Recordset, EMRUZ As Long
    DT1 = CLng(Forms![BASEKNOW]![YEA] & RIGHT("00" & Me.Frame13, 2) & "01")
    DT2 = CLng(Forms![BASEKNOW]![YEA] & RIGHT("00" & Me.Frame13, 2) & "31")
    EMRUZ = DIFF(DT1, FARSIDATE(Now()))
    If EMRUZ - 31 > 0 Then
        EMRUZ = 0
    Else
        EMRUZ = Abs(EMRUZ - 31)
    End If
    If Me.FG = 1 Then
        Me.VISITOR_GOL_DTL_GRP_REP.Form.RecordSource = "SELECT     TOP 100 PERCENT dbo.visitgol_dtl.HES, dbo.visitgol_dtl.MAH, dbo.TCODE_MENUITEM.NAMES, SUM(ISNULL(VISITOR_DTL_KALA_MARA.MABL_K, 0))  AS MABL_KS, SUM(ISNULL(VISITOR_DTL_KALA_MARA.MABMAR, 0)) AS MABMARS, SUM(ISNULL(VISITOR_DTL_KALA_MARA.MEGH_MAR, 0)) AS MEGH_MARS,  SUM(ISNULL(VISITOR_DTL_KALA_MARA.MEGHk, 0)) AS MEGHkS, SUM(dbo.visitgol_dtl.MEGHk) AS MEGHkGOL, SUM(ISNULL(dbo.visitgol_dtl.MEGHk, 0) - ISNULL(VISITOR_DTL_KALA_MARA.MEGHk, 0)) AS MANDMEGH, SUM(ISNULL(VISITOR_DTL_KALA_MARA.MEGHk, 0)) / sum(dbo.visitgol_dtl.MEGHk )* 100 AS DARSADFR," & EMRUZ & "  AS DAYMAND, dbo.CUST_HESAB.NAME AS VISITOR, dbo.TCODE_MENUITEM.CODE AS MENUIT FROM         dbo.STUF_DEF INNER JOIN" & _
                                            " dbo.visitgol_dtl ON dbo.STUF_DEF.CODE = dbo.visitgol_dtl.CODE INNER JOIN   dbo.TCODE_MENUITEM ON dbo.STUF_DEF.MENUIT = dbo.TCODE_MENUITEM.CODE INNER JOIN   dbo.CUST_HESAB ON dbo.visitgol_dtl.HES = dbo.CUST_HESAB.hes LEFT OUTER JOIN  dbo.VISITOR_DTL_KALA_MARA(" & DT1 & "," & DT2 & ",'" & Me.HES & "') VISITOR_DTL_KALA_MARA ON dbo.visitgol_dtl.CODE = VISITOR_DTL_KALA_MARA.CODE AND   dbo.visitgol_dtl.HES = VISITOR_DTL_KALA_MARA.CUST_NO GROUP BY dbo.visitgol_dtl.HES, dbo.visitgol_dtl.MAH, dbo.TCODE_MENUITEM.NAMES, dbo.CUST_HESAB.NAME, dbo.TCODE_MENUITEM.CODE HAVING      (dbo.visitgol_dtl.MAH = " & Me.Frame13 & ") AND (dbo.visitgol_dtl.HES = N'" & Me.HES & "')"
    Else
        Me.VISITOR_GOL_DTL_GRP_REP.Form.RecordSource = "SELECT     TOP 100 PERCENT dbo.visitgol_dtl.HES, dbo.visitgol_dtl.MAH, dbo.TCOD_STUFGROUP.NAMES, SUM(ISNULL(VISITOR_DTL_KALA_mara.MABL_K, 0)) AS MABL_KS, SUM(ISNULL(VISITOR_DTL_KALA_mara.MABMAR, 0)) AS MABMARS, SUM(ISNULL(VISITOR_DTL_KALA_mara.MEGH_MAR, 0)) AS MEGH_MARS, SUM(ISNULL(VISITOR_DTL_KALA_mara.MEGHk, 0)) AS MEGHkS, SUM(dbo.visitgol_dtl.MEGHk) AS MEGHkGOL, SUM(ISNULL(dbo.visitgol_dtl.MEGHk, 0) - ISNULL(VISITOR_DTL_KALA_mara.MEGHk, 0)) AS MANDMEGH, SUM(ISNULL(VISITOR_DTL_KALA_mara.MEGHk, 0)) / SUM(dbo.visitgol_dtl.MEGHk) * 100 AS DARSADFR, " & EMRUZ & " AS DAYMAND, dbo.CUST_HESAB.NAME AS VISITOR, dbo.TCOD_STUFGROUP.CODE AS MENUIT FROM dbo.STUF_DEF INNER JOIN  dbo.visitgol_dtl ON dbo.STUF_DEF.CODE = dbo.visitgol_dtl.CODE INNER JOIN  dbo.CUST_HESAB ON dbo.visitgol_dtl.HES = dbo.CUST_HESAB.hes INNER JOIN  dbo.TCOD_STUFGROUP ON dbo.STUF_DEF.RADAH = dbo.TCOD_STUFGROUP.CODE LEFT OUTER JOIN" & _
                                                       " dbo.VISITOR_DTL_KALA_mara(" & DT1 & "," & DT2 & ",'" & Me.HES & "')VISITOR_DTL_KALA_mara ON dbo.visitgol_dtl.CODE = VISITOR_DTL_KALA_mara.CODE AND  dbo.visitgol_dtl.hes = VISITOR_DTL_KALA_mara.CUST_NO GROUP BY dbo.visitgol_dtl.HES, dbo.visitgol_dtl.MAH, dbo.TCOD_STUFGROUP.NAMES, dbo.CUST_HESAB.NAME, dbo.TCOD_STUFGROUP.CODE HAVING       (dbo.visitgol_dtl.MAH = " & Me.Frame13 & ") AND (dbo.visitgol_dtl.HES = N'" & Me.HES & "')"
    End If

End Sub

Private Sub Form_Open(CANCEL As Integer)
                 On Error Resume Next
        Forms![BASEKNOW]![hhwin] = Me.hWnd
    Dim DT1 As Long, DT2 As Long, rst As New ADODB.Recordset, EMRUZ As Long
    rst.Open "SELECT     HES FROM   visitgol_head", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        Me.HES = rst.Fields("hes")
        Me.HES2 = rst.Fields("hes")
        DT1 = CLng(Forms![BASEKNOW]![YEA] & "0101")
        DT2 = CLng(Forms![BASEKNOW]![YEA] & "0131")
        EMRUZ = DIFF(DT1, FARSIDATE(Now()))
        If EMRUZ - 31 > 0 Then
            EMRUZ = 0
        Else
            EMRUZ = Abs(EMRUZ - 31)
        End If
        If Me.FG = 1 Then
            Me.VISITOR_GOL_DTL_GRP_REP.Form.RecordSource = "SELECT     TOP 100 PERCENT dbo.visitgol_dtl.HES, dbo.visitgol_dtl.MAH, dbo.TCODE_MENUITEM.NAMES, SUM(ISNULL(VISITOR_DTL_KALA_MARA.MABL_K, 0))  AS MABL_KS, SUM(ISNULL(VISITOR_DTL_KALA_MARA.MABMAR, 0)) AS MABMARS, SUM(ISNULL(VISITOR_DTL_KALA_MARA.MEGH_MAR, 0)) AS MEGH_MARS,  SUM(ISNULL(VISITOR_DTL_KALA_MARA.MEGHk, 0)) AS MEGHkS, SUM(dbo.visitgol_dtl.MEGHk) AS MEGHkGOL, SUM(ISNULL(dbo.visitgol_dtl.MEGHk, 0) - ISNULL(VISITOR_DTL_KALA_MARA.MEGHk, 0)) AS MANDMEGH, SUM(ISNULL(VISITOR_DTL_KALA_MARA.MEGHk, 0)) / sum(dbo.visitgol_dtl.MEGHk )* 100 AS DARSADFR," & EMRUZ & "  AS DAYMAND, dbo.CUST_HESAB.NAME AS VISITOR, dbo.TCODE_MENUITEM.CODE AS MENUIT FROM         dbo.STUF_DEF INNER JOIN" & _
                                                " dbo.visitgol_dtl ON dbo.STUF_DEF.CODE = dbo.visitgol_dtl.CODE INNER JOIN   dbo.TCODE_MENUITEM ON dbo.STUF_DEF.MENUIT = dbo.TCODE_MENUITEM.CODE INNER JOIN   dbo.CUST_HESAB ON dbo.visitgol_dtl.HES = dbo.CUST_HESAB.hes LEFT OUTER JOIN  dbo.VISITOR_DTL_KALA_MARA(" & DT1 & "," & DT2 & ",'" & Me.HES & "') VISITOR_DTL_KALA_MARA ON dbo.visitgol_dtl.CODE = VISITOR_DTL_KALA_MARA.CODE AND   dbo.visitgol_dtl.HES = VISITOR_DTL_KALA_MARA.CUST_NO GROUP BY dbo.visitgol_dtl.HES, dbo.visitgol_dtl.MAH, dbo.TCODE_MENUITEM.NAMES, dbo.CUST_HESAB.NAME, dbo.TCODE_MENUITEM.CODE HAVING      (dbo.visitgol_dtl.MAH = " & Me.Frame13 & ") AND (dbo.visitgol_dtl.HES = N'" & Me.HES & "')"
        Else
            Me.VISITOR_GOL_DTL_GRP_REP.Form.RecordSource = "SELECT     TOP 100 PERCENT dbo.visitgol_dtl.HES, dbo.visitgol_dtl.MAH, dbo.TCOD_STUFGROUP.NAMES, SUM(ISNULL(VISITOR_DTL_KALA_mara.MABL_K, 0)) AS MABL_KS, SUM(ISNULL(VISITOR_DTL_KALA_mara.MABMAR, 0)) AS MABMARS, SUM(ISNULL(VISITOR_DTL_KALA_mara.MEGH_MAR, 0)) AS MEGH_MARS, SUM(ISNULL(VISITOR_DTL_KALA_mara.MEGHk, 0)) AS MEGHkS, SUM(dbo.visitgol_dtl.MEGHk) AS MEGHkGOL, SUM(ISNULL(dbo.visitgol_dtl.MEGHk, 0) - ISNULL(VISITOR_DTL_KALA_mara.MEGHk, 0)) AS MANDMEGH, SUM(ISNULL(VISITOR_DTL_KALA_mara.MEGHk, 0)) / SUM(dbo.visitgol_dtl.MEGHk) * 100 AS DARSADFR, " & EMRUZ & " AS DAYMAND, dbo.CUST_HESAB.NAME AS VISITOR, dbo.TCOD_STUFGROUP.CODE AS MENUIT FROM dbo.STUF_DEF INNER JOIN  dbo.visitgol_dtl ON dbo.STUF_DEF.CODE = dbo.visitgol_dtl.CODE INNER JOIN  dbo.CUST_HESAB ON dbo.visitgol_dtl.HES = dbo.CUST_HESAB.hes INNER JOIN  dbo.TCOD_STUFGROUP ON dbo.STUF_DEF.RADAH = dbo.TCOD_STUFGROUP.CODE LEFT OUTER JOIN" & _
                                                           " dbo.VISITOR_DTL_KALA_mara(950101, 950131, '213-1-6') VISITOR_DTL_KALA_mara ON dbo.visitgol_dtl.CODE = VISITOR_DTL_KALA_mara.CODE AND  dbo.visitgol_dtl.hes = VISITOR_DTL_KALA_mara.CUST_NO GROUP BY dbo.visitgol_dtl.HES, dbo.visitgol_dtl.MAH, dbo.TCOD_STUFGROUP.NAMES, dbo.CUST_HESAB.NAME, dbo.TCOD_STUFGROUP.CODE HAVING       (dbo.visitgol_dtl.MAH = " & Me.Frame13 & ") AND (dbo.visitgol_dtl.HES = N'" & Me.HES & "')"
        End If
    End If
End Sub

Private Sub Frame13_Click()
                 On Error Resume Next
    Dim DT1 As Long, DT2 As Long, rst As New ADODB.Recordset, EMRUZ As Long
    DT1 = CLng(Forms![BASEKNOW]![YEA] & RIGHT("00" & Me.Frame13, 2) & "01")
    DT2 = CLng(Forms![BASEKNOW]![YEA] & RIGHT("00" & Me.Frame13, 2) & "31")
    EMRUZ = DIFF(DT1, FARSIDATE(Now()))
    If EMRUZ - 31 > 0 Then
        EMRUZ = 0
    Else
        EMRUZ = Abs(EMRUZ - 31)
    End If
    If Me.FG = 1 Then
        Me.VISITOR_GOL_DTL_GRP_REP.Form.RecordSource = "SELECT     TOP 100 PERCENT dbo.visitgol_dtl.HES, dbo.visitgol_dtl.MAH, dbo.TCODE_MENUITEM.NAMES, SUM(ISNULL(VISITOR_DTL_KALA_MARA.MABL_K, 0))  AS MABL_KS, SUM(ISNULL(VISITOR_DTL_KALA_MARA.MABMAR, 0)) AS MABMARS, SUM(ISNULL(VISITOR_DTL_KALA_MARA.MEGH_MAR, 0)) AS MEGH_MARS,  SUM(ISNULL(VISITOR_DTL_KALA_MARA.MEGHk, 0)) AS MEGHkS, SUM(dbo.visitgol_dtl.MEGHk) AS MEGHkGOL, SUM(ISNULL(dbo.visitgol_dtl.MEGHk, 0) - ISNULL(VISITOR_DTL_KALA_MARA.MEGHk, 0)) AS MANDMEGH, SUM(ISNULL(VISITOR_DTL_KALA_MARA.MEGHk, 0)) / sum(dbo.visitgol_dtl.MEGHk )* 100 AS DARSADFR," & EMRUZ & "  AS DAYMAND, dbo.CUST_HESAB.NAME AS VISITOR, dbo.TCODE_MENUITEM.CODE AS MENUIT FROM         dbo.STUF_DEF INNER JOIN" & _
                                            " dbo.visitgol_dtl ON dbo.STUF_DEF.CODE = dbo.visitgol_dtl.CODE INNER JOIN   dbo.TCODE_MENUITEM ON dbo.STUF_DEF.MENUIT = dbo.TCODE_MENUITEM.CODE INNER JOIN   dbo.CUST_HESAB ON dbo.visitgol_dtl.HES = dbo.CUST_HESAB.hes LEFT OUTER JOIN  dbo.VISITOR_DTL_KALA_MARA(" & DT1 & "," & DT2 & ",'" & Me.HES & "') VISITOR_DTL_KALA_MARA ON dbo.visitgol_dtl.CODE = VISITOR_DTL_KALA_MARA.CODE AND   dbo.visitgol_dtl.HES = VISITOR_DTL_KALA_MARA.CUST_NO GROUP BY dbo.visitgol_dtl.HES, dbo.visitgol_dtl.MAH, dbo.TCODE_MENUITEM.NAMES, dbo.CUST_HESAB.NAME, dbo.TCODE_MENUITEM.CODE HAVING      (dbo.visitgol_dtl.MAH = " & Me.Frame13 & ") AND (dbo.visitgol_dtl.HES = N'" & Me.HES & "')"
    Else
        Me.VISITOR_GOL_DTL_GRP_REP.Form.RecordSource = "SELECT     TOP 100 PERCENT dbo.visitgol_dtl.HES, dbo.visitgol_dtl.MAH, dbo.TCOD_STUFGROUP.NAMES, SUM(ISNULL(VISITOR_DTL_KALA_mara.MABL_K, 0)) AS MABL_KS, SUM(ISNULL(VISITOR_DTL_KALA_mara.MABMAR, 0)) AS MABMARS, SUM(ISNULL(VISITOR_DTL_KALA_mara.MEGH_MAR, 0)) AS MEGH_MARS, SUM(ISNULL(VISITOR_DTL_KALA_mara.MEGHk, 0)) AS MEGHkS, SUM(dbo.visitgol_dtl.MEGHk) AS MEGHkGOL, SUM(ISNULL(dbo.visitgol_dtl.MEGHk, 0) - ISNULL(VISITOR_DTL_KALA_mara.MEGHk, 0)) AS MANDMEGH, SUM(ISNULL(VISITOR_DTL_KALA_mara.MEGHk, 0)) / SUM(dbo.visitgol_dtl.MEGHk) * 100 AS DARSADFR, " & EMRUZ & " AS DAYMAND, dbo.CUST_HESAB.NAME AS VISITOR, dbo.TCOD_STUFGROUP.CODE AS MENUIT FROM dbo.STUF_DEF INNER JOIN  dbo.visitgol_dtl ON dbo.STUF_DEF.CODE = dbo.visitgol_dtl.CODE INNER JOIN  dbo.CUST_HESAB ON dbo.visitgol_dtl.HES = dbo.CUST_HESAB.hes INNER JOIN  dbo.TCOD_STUFGROUP ON dbo.STUF_DEF.RADAH = dbo.TCOD_STUFGROUP.CODE LEFT OUTER JOIN" & _
                                                       " dbo.VISITOR_DTL_KALA_mara(950101, 950131, '213-1-6') VISITOR_DTL_KALA_mara ON dbo.visitgol_dtl.CODE = VISITOR_DTL_KALA_mara.CODE AND  dbo.visitgol_dtl.hes = VISITOR_DTL_KALA_mara.CUST_NO GROUP BY dbo.visitgol_dtl.HES, dbo.visitgol_dtl.MAH, dbo.TCOD_STUFGROUP.NAMES, dbo.CUST_HESAB.NAME, dbo.TCOD_STUFGROUP.CODE HAVING       (dbo.visitgol_dtl.MAH = " & Me.Frame13 & ") AND (dbo.visitgol_dtl.HES = N'" & Me.HES & "')"
    End If
End Sub

Private Sub HES_AfterUpdate()
                 On Error Resume Next
    Me.HES2 = Me.HES
    Dim DT1 As Long, DT2 As Long, rst As New ADODB.Recordset, EMRUZ As Long
    DT1 = CLng(Forms![BASEKNOW]![YEA] & RIGHT("00" & Me.Frame13, 2) & "01")
    DT2 = CLng(Forms![BASEKNOW]![YEA] & RIGHT("00" & Me.Frame13, 2) & "31")
    EMRUZ = DIFF(DT1, FARSIDATE(Now()))
    If EMRUZ - 31 > 0 Then
        EMRUZ = 0
    Else
        EMRUZ = Abs(EMRUZ - 31)
    End If
    If Me.FG = 1 Then
        Me.VISITOR_GOL_DTL_GRP_REP.Form.RecordSource = "SELECT     TOP 100 PERCENT dbo.visitgol_dtl.HES, dbo.visitgol_dtl.MAH, dbo.TCODE_MENUITEM.NAMES, SUM(ISNULL(VISITOR_DTL_KALA_MARA.MABL_K, 0))  AS MABL_KS, SUM(ISNULL(VISITOR_DTL_KALA_MARA.MABMAR, 0)) AS MABMARS, SUM(ISNULL(VISITOR_DTL_KALA_MARA.MEGH_MAR, 0)) AS MEGH_MARS,  SUM(ISNULL(VISITOR_DTL_KALA_MARA.MEGHk, 0)) AS MEGHkS, SUM(dbo.visitgol_dtl.MEGHk) AS MEGHkGOL, SUM(ISNULL(dbo.visitgol_dtl.MEGHk, 0) - ISNULL(VISITOR_DTL_KALA_MARA.MEGHk, 0)) AS MANDMEGH, SUM(ISNULL(VISITOR_DTL_KALA_MARA.MEGHk, 0)) / sum(dbo.visitgol_dtl.MEGHk )* 100 AS DARSADFR," & EMRUZ & "  AS DAYMAND, dbo.CUST_HESAB.NAME AS VISITOR, dbo.TCODE_MENUITEM.CODE AS MENUIT FROM         dbo.STUF_DEF INNER JOIN" & _
                                            " dbo.visitgol_dtl ON dbo.STUF_DEF.CODE = dbo.visitgol_dtl.CODE INNER JOIN   dbo.TCODE_MENUITEM ON dbo.STUF_DEF.MENUIT = dbo.TCODE_MENUITEM.CODE INNER JOIN   dbo.CUST_HESAB ON dbo.visitgol_dtl.HES = dbo.CUST_HESAB.hes LEFT OUTER JOIN  dbo.VISITOR_DTL_KALA_MARA(" & DT1 & "," & DT2 & ",'" & Me.HES & "') VISITOR_DTL_KALA_MARA ON dbo.visitgol_dtl.CODE = VISITOR_DTL_KALA_MARA.CODE AND   dbo.visitgol_dtl.HES = VISITOR_DTL_KALA_MARA.CUST_NO GROUP BY dbo.visitgol_dtl.HES, dbo.visitgol_dtl.MAH, dbo.TCODE_MENUITEM.NAMES, dbo.CUST_HESAB.NAME, dbo.TCODE_MENUITEM.CODE HAVING      (dbo.visitgol_dtl.MAH = " & Me.Frame13 & ") AND (dbo.visitgol_dtl.HES = N'" & Me.HES & "')"
    Else
        Me.VISITOR_GOL_DTL_GRP_REP.Form.RecordSource = "SELECT     TOP 100 PERCENT dbo.visitgol_dtl.HES, dbo.visitgol_dtl.MAH, dbo.TCOD_STUFGROUP.NAMES, SUM(ISNULL(VISITOR_DTL_KALA_mara.MABL_K, 0)) AS MABL_KS, SUM(ISNULL(VISITOR_DTL_KALA_mara.MABMAR, 0)) AS MABMARS, SUM(ISNULL(VISITOR_DTL_KALA_mara.MEGH_MAR, 0)) AS MEGH_MARS, SUM(ISNULL(VISITOR_DTL_KALA_mara.MEGHk, 0)) AS MEGHkS, SUM(dbo.visitgol_dtl.MEGHk) AS MEGHkGOL, SUM(ISNULL(dbo.visitgol_dtl.MEGHk, 0) - ISNULL(VISITOR_DTL_KALA_mara.MEGHk, 0)) AS MANDMEGH, SUM(ISNULL(VISITOR_DTL_KALA_mara.MEGHk, 0)) / SUM(dbo.visitgol_dtl.MEGHk) * 100 AS DARSADFR, " & EMRUZ & " AS DAYMAND, dbo.CUST_HESAB.NAME AS VISITOR, dbo.TCOD_STUFGROUP.CODE AS MENUIT FROM dbo.STUF_DEF INNER JOIN  dbo.visitgol_dtl ON dbo.STUF_DEF.CODE = dbo.visitgol_dtl.CODE INNER JOIN  dbo.CUST_HESAB ON dbo.visitgol_dtl.HES = dbo.CUST_HESAB.hes INNER JOIN  dbo.TCOD_STUFGROUP ON dbo.STUF_DEF.RADAH = dbo.TCOD_STUFGROUP.CODE LEFT OUTER JOIN" & _
                                                       " dbo.VISITOR_DTL_KALA_mara(950101, 950131, '213-1-6') VISITOR_DTL_KALA_mara ON dbo.visitgol_dtl.CODE = VISITOR_DTL_KALA_mara.CODE AND  dbo.visitgol_dtl.hes = VISITOR_DTL_KALA_mara.CUST_NO GROUP BY dbo.visitgol_dtl.HES, dbo.visitgol_dtl.MAH, dbo.TCOD_STUFGROUP.NAMES, dbo.CUST_HESAB.NAME, dbo.TCOD_STUFGROUP.CODE HAVING       (dbo.visitgol_dtl.MAH = " & Me.Frame13 & ") AND (dbo.visitgol_dtl.HES = N'" & Me.HES & "')"
    End If
End Sub

Private Sub HES2_AfterUpdate()
    Me.HES = Me.HES2
End Sub
