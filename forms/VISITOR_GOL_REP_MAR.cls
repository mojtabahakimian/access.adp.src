Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command55_Click()
     DoCmd.OpenReport "Visit_gol_dtl_REP_print", acViewPreview
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
        Me.Visit_gol_dtl_REP.Form.RecordSource = "SELECT     TOP 100 PERCENT dbo.visitgol_dtl.CODE, VISITOR_DTL_KALA_marA.CUST_NO, dbo.visitgol_dtl.HES, dbo.visitgol_dtl.MAH,dbo.STUF_DEF.NAME AS kala, ISNULL(VISITOR_DTL_KALA_marA.MABL_K, 0) AS MABL_K, ISNULL(VISITOR_DTL_KALA_marA.MABMAR, 0) AS MABMAR,ISNULL(VISITOR_DTL_KALA_marA.MEGH_MAR, 0) AS MEGH_MAR, ISNULL(VISITOR_DTL_KALA_marA.MEGHk, 0) AS MEGHk,VISITOR_DTL_KALA_marA.VISITOR, dbo.visitgol_dtl.MEGHk AS MEGHkGOL, ISNULL(dbo.visitgol_dtl.MEGHk - VISITOR_DTL_KALA_marA.MEGHk, 0) AS MANDMEGH, ISNULL(VISITOR_DTL_KALA_marA.MEGHk / dbo.UIIF(dbo.visitgol_dtl.MEGHk, N'=', 0, 1, dbo.visitgol_dtl.MEGHk) * 100, 0) AS DARSADFR, 0 AS DAYMAND FROM dbo.STUF_DEF INNER JOIN dbo.visitgol_dtl ON dbo.STUF_DEF.CODE = dbo.visitgol_dtl.CODE LEFT OUTER JOIN dbo.VISITOR_DTL_KALA_marA(" & DT1 & "," & DT2 & ",'" & Me.HES & "') VISITOR_DTL_KALA_marA ON dbo.visitgol_dtl.CODE = VISITOR_DTL_KALA_marA.CODE AND  dbo.visitgol_dtl.hes = VISITOR_DTL_KALA_marA.CUST_NO " & _
                                                 " WHERE     (dbo.visitgol_dtl.MAH = " & Me.Frame13 & ") AND (dbo.visitgol_dtl.HES = '" & Me.HES & "') ORDER BY dbo.STUF_DEF.NAME"

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
        Me.Visit_gol_dtl_REP.Form.RecordSource = "SELECT     TOP 100 PERCENT dbo.visitgol_dtl.CODE, VISITOR_DTL_KALA_marA.CUST_NO, dbo.visitgol_dtl.HES, dbo.visitgol_dtl.MAH,dbo.STUF_DEF.NAME AS kala, ISNULL(VISITOR_DTL_KALA_marA.MABL_K, 0) AS MABL_K, ISNULL(VISITOR_DTL_KALA_marA.MABMAR, 0) AS MABMAR,ISNULL(VISITOR_DTL_KALA_marA.MEGH_MAR, 0) AS MEGH_MAR, ISNULL(VISITOR_DTL_KALA_marA.MEGHk, 0) AS MEGHk,VISITOR_DTL_KALA_marA.VISITOR, dbo.visitgol_dtl.MEGHk AS MEGHkGOL, ISNULL(dbo.visitgol_dtl.MEGHk - VISITOR_DTL_KALA_marA.MEGHk, 0) AS MANDMEGH, ISNULL(VISITOR_DTL_KALA_marA.MEGHk / dbo.UIIF(dbo.visitgol_dtl.MEGHk, N'=', 0, 1, dbo.visitgol_dtl.MEGHk) * 100, 0) AS DARSADFR, 0 AS DAYMAND FROM dbo.STUF_DEF INNER JOIN dbo.visitgol_dtl ON dbo.STUF_DEF.CODE = dbo.visitgol_dtl.CODE LEFT OUTER JOIN dbo.VISITOR_DTL_KALA_marA(" & DT1 & "," & DT2 & ",'" & Me.HES & "') VISITOR_DTL_KALA_marA ON dbo.visitgol_dtl.CODE = VISITOR_DTL_KALA_marA.CODE AND  dbo.visitgol_dtl.hes = VISITOR_DTL_KALA_marA.CUST_NO " & _
                                                 " WHERE     (dbo.visitgol_dtl.MAH = " & Me.Frame13 & ") AND (dbo.visitgol_dtl.HES = '" & Me.HES & "') ORDER BY dbo.STUF_DEF.NAME"

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
        Me.Visit_gol_dtl_REP.Form.RecordSource = "SELECT     TOP 100 PERCENT dbo.visitgol_dtl.CODE, VISITOR_DTL_KALA_marA.CUST_NO, dbo.visitgol_dtl.HES, dbo.visitgol_dtl.MAH,dbo.STUF_DEF.NAME AS kala, ISNULL(VISITOR_DTL_KALA_marA.MABL_K, 0) AS MABL_K, ISNULL(VISITOR_DTL_KALA_marA.MABMAR, 0) AS MABMAR,ISNULL(VISITOR_DTL_KALA_marA.MEGH_MAR, 0) AS MEGH_MAR, ISNULL(VISITOR_DTL_KALA_marA.MEGHk, 0) AS MEGHk,VISITOR_DTL_KALA_marA.VISITOR, dbo.visitgol_dtl.MEGHk AS MEGHkGOL, ISNULL(dbo.visitgol_dtl.MEGHk - VISITOR_DTL_KALA_marA.MEGHk, 0) AS MANDMEGH, ISNULL(VISITOR_DTL_KALA_marA.MEGHk / dbo.UIIF(dbo.visitgol_dtl.MEGHk, N'=', 0, 1, dbo.visitgol_dtl.MEGHk) * 100, 0) AS DARSADFR, 0 AS DAYMAND FROM dbo.STUF_DEF INNER JOIN dbo.visitgol_dtl ON dbo.STUF_DEF.CODE = dbo.visitgol_dtl.CODE LEFT OUTER JOIN dbo.VISITOR_DTL_KALA_marA(" & DT1 & "," & DT2 & ",'" & Me.HES & "') VISITOR_DTL_KALA_marA ON dbo.visitgol_dtl.CODE = VISITOR_DTL_KALA_marA.CODE AND  dbo.visitgol_dtl.hes = VISITOR_DTL_KALA_marA.CUST_NO " & _
                                                 " WHERE     (dbo.visitgol_dtl.MAH = " & Me.Frame13 & ") AND (dbo.visitgol_dtl.HES = '" & Me.HES & "') ORDER BY dbo.STUF_DEF.NAME"

End Sub

Private Sub HES2_AfterUpdate()
    Me.HES = Me.HES2
End Sub

Private Sub Toggle54_Click()
End Sub
