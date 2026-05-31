Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Sub Command5_Click()
                    On Error Resume Next
  Dim sql   As String
  Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, PATH As String, i As Integer, CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4
  If IsNull(Me.dt) Or IsNull(Me.S2) Then
    DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "پارامترها كافي نيست!"
    Exit Sub
  End If
  Forms![BASEKNOW]![Text44] = False
  DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "آيا درمورد صدور سند اختتاميه اطمينان داريد"
  If Forms![BASEKNOW]![Text44] Then
       rst.Open "DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
       rst.Filter = "N_S = " & Me.S2
       If rst.RecordCount = 0 Then
           rst.AddNew
           rst.Fields("N_S") = Me.S2
           rst.Fields("DATE_S") = Me.dt
           rst.Fields("SHARH_S") = Left(Me.SH, 255)
           rst.Fields("NO_S") = 0
           rst.Fields("OKF") = True
           rst.Fields("USER_NAME") = UCurrentUser()
           rst.update
           rst.Close
       Else
           DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "شماره سند تكراري مي باشد."
           Exit Sub
       End If
       If Me.KIND = 1 Then
            sql = "INSERT INTO dbo.DEED_DTL (N_S, HES_K, HES_M, HES_T, HES_T2, HES_T3, HES_T4, HES, BES, BED, SHARH) SELECT     " & Me.S2 & " AS N_S, HES_K, HES_M, HES_T, HES_T2, HES_T3, HES_T4,HES, dbo.UIIF(BEST, '>', 0, BEST, 0) AS BES, dbo.UIIF(BEST, '<', 0, BEST * - 1, 0) AS BED, 'سند اختتاميه' AS SH FROM dbo.[BED-BES] INNER JOIN  dbo.TOTA_HES ON dbo.[BED-BES].HES_K = dbo.TOTA_HES.NUMBER WHERE (HES <> '" & Forms![BASEKNOW]![ADA] & "' And HES <> '" & Forms![BASEKNOW]![APA] & "') AND (HES <> '" & Forms![BASEKNOW]![ADV] & "' And HES <> '" & Forms![BASEKNOW]![APV] & "') And (dbo.TOTA_HES.M_D = 1)"
        Else
            sql = "INSERT INTO dbo.DEED_DTL (N_S, HES_K, HES_M, HES_T, HES_T2, HES_T3, HES_T4, HES, BES, BED, SHARH) SELECT     " & Me.S2 & " AS N_S, HES_K, HES_M, HES_T, HES_T2, HES_T3, HES_T4,HES, dbo.UIIF(BEST, '>', 0, BEST, 0) AS BES, dbo.UIIF(BEST, '<', 0, BEST * - 1, 0) AS BED, 'سند اختتاميه' AS SH FROM dbo.[BED-BES] WHERE (HES <> '" & Forms![BASEKNOW]![ADA] & "' And HES <> '" & Forms![BASEKNOW]![APA] & "') AND (HES <> '" & Forms![BASEKNOW]![ADV] & "' And HES <> '" & Forms![BASEKNOW]![APV] & "')"
       End If
       DoCmd.RunSQL sql
       'اسناد تجاري
       Call GETTAF3(Forms![BASEKNOW]![ADA], CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4)
       sql = "INSERT INTO dbo.DEED_DTL  (N_SERI, BANK, BES, HES_K, HES_M, HES_T, HES_T2, HES_T3, HES_T4, HES, N_S, SHARH) SELECT     dbo.PAY_GETD.N_SERI, dbo.PAY_GETD.BANK, dbo.PAY_GETD.MABL, " & CKOL & "  AS KOL, " & CMOIN & " AS MOIN, " & CTAF & " AS TAF, " & IIf(IsNull(CTAF2), "NULL", CTAF2) & " AS TAF2, " & IIf(IsNull(CTAF3), "NULL", CTAF3) & " AS TAF3, " & IIf(IsNull(CTAF4), "NULL", CTAF4) & " AS TAF4,  '" & Forms![BASEKNOW]![ADA] & "' AS Expr2, " & Me.S2 & " AS NS, RIGHT(' چك' + CAST(CAST(dbo.PAY_GETD.N_SERI AS numeric) AS nvarchar(50))+ ' بانك ' +  ISNULL(dbo.TCOD_BANKS.NAMES, '  ') + ' ' + ISNULL(dbo.PAY_GETD.SHOBEH, '  ') + ' مورخ ' + CAST(dbo.PAY_GETD.DATE_S AS NVARCHAR),60) AS Expr1 FROM  dbo.TCOD_BANKS INNER JOIN   dbo.PAY_GETD ON dbo.TCOD_BANKS.CODE = dbo.PAY_GETD.BANK WHERE     (dbo.PAY_GETD.N_KOL IS NULL OR   dbo.PAY_GETD.N_KOL = " & Forms![BASEKNOW]![BANKHA] & ") AND (dbo.PAY_GETD.N_KOL2 IS NULL) AND (dbo.PAY_GETD.N_KOL3 IS NULL) and (KIND = 1 OR KIND IS NULL)"
       DoCmd.RunSQL sql
       'اسناد غير تجاري
       Call GETTAF3(Forms![BASEKNOW]![ADV], CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4)
       sql = "INSERT INTO dbo.DEED_DTL  (N_SERI, BANK, BES, HES_K, HES_M, HES_T, HES_T2, HES_T3, HES_T4, HES, N_S, SHARH) SELECT     dbo.PAY_GETD.N_SERI, dbo.PAY_GETD.BANK, dbo.PAY_GETD.MABL, " & CKOL & "  AS KOL, " & CMOIN & " AS MOIN, " & CTAF & " AS TAF, " & IIf(IsNull(CTAF2), "NULL", CTAF2) & " AS TAF2, " & IIf(IsNull(CTAF3), "NULL", CTAF3) & " AS TAF3, " & IIf(IsNull(CTAF4), "NULL", CTAF4) & " AS TAF4,    '" & Forms![BASEKNOW]![ADV] & "' AS Expr2, " & Me.S2 & " AS NS, RIGHT(' چك' + CAST(CAST(dbo.PAY_GETD.N_SERI AS numeric) AS nvarchar(50))+ ' بانك ' +  ISNULL(dbo.TCOD_BANKS.NAMES, '  ') + ' ' + ISNULL(dbo.PAY_GETD.SHOBEH, '  ') + ' مورخ ' + CAST(dbo.PAY_GETD.DATE_S AS NVARCHAR),60) AS Expr1 FROM  dbo.TCOD_BANKS INNER JOIN   dbo.PAY_GETD ON dbo.TCOD_BANKS.CODE = dbo.PAY_GETD.BANK WHERE     (dbo.PAY_GETD.N_KOL IS NULL OR   dbo.PAY_GETD.N_KOL = " & Forms![BASEKNOW]![BANKHA] & ") AND (dbo.PAY_GETD.N_KOL2 IS NULL) AND (dbo.PAY_GETD.N_KOL3 IS NULL) and kind <> 1"
       DoCmd.RunSQL sql
       Call GETTAF3(Forms![BASEKNOW]![APA], CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4)
       sql = "INSERT INTO dbo.DEED_DTL  (N_SERI, BANK, BED, HES_K, HES_M, HES_T, HES_T2, HES_T3, HES_T4, HES, N_S, SHARH) SELECT     dbo.PAY_GETP.N_SERI, dbo.PAY_GETP.BANK, dbo.PAY_GETP.MABL, " & CKOL & "  AS KOL, " & CMOIN & " AS MOIN, " & CTAF & " AS TAF, " & IIf(IsNull(CTAF2), "NULL", CTAF2) & " AS TAF2, " & IIf(IsNull(CTAF3), "NULL", CTAF3) & " AS TAF3, " & IIf(IsNull(CTAF4), "NULL", CTAF4) & " AS TAF4,   '" & Forms![BASEKNOW]![APA] & "' AS Expr2, " & Me.S2 & " AS NS, RIGHT(' چك' + CAST(CAST(dbo.PAY_GETp.N_SERI AS numeric) AS nvarchar(50))+ ' بانك ' + dbo.TCOD_BANKS.NAMES + ' ' + ISNULL(dbo.PAY_GETP.SHOBEH,'  ') + ' مورخ ' + CAST(dbo.PAY_GETP.DATE_S AS NVARCHAR),60) AS Expr1 FROM  dbo.TCOD_BANKS INNER JOIN   dbo.PAY_GETP ON dbo.TCOD_BANKS.CODE = dbo.PAY_GETP.BANK WHERE      (dbo.PAY_GETP.N_KOL2 IS NULL) AND (dbo.PAY_GETP.N_KOL3 IS NULL) and kind = 1 and dbo.PAY_GETP.N_KOL<> 911"
       DoCmd.RunSQL sql
       'اسناد غير تجاري
       Call GETTAF3(Forms![BASEKNOW]![APV], CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4)
       sql = "INSERT INTO dbo.DEED_DTL  (N_SERI, BANK, BED, HES_K, HES_M, HES_T, HES_T2, HES_T3, HES_T4, HES, N_S, SHARH) SELECT     dbo.PAY_GETP.N_SERI, dbo.PAY_GETP.BANK, dbo.PAY_GETP.MABL, " & CKOL & "  AS KOL, " & CMOIN & " AS MOIN, " & CTAF & " AS TAF, " & IIf(IsNull(CTAF2), "NULL", CTAF2) & " AS TAF2, " & IIf(IsNull(CTAF3), "NULL", CTAF3) & " AS TAF3, " & IIf(IsNull(CTAF4), "NULL", CTAF4) & " AS TAF4,   '" & Forms![BASEKNOW]![APV] & "' AS Expr2, " & Me.S2 & " AS NS, RIGHT(' چك' + CAST(CAST(dbo.PAY_GETp.N_SERI AS numeric) AS nvarchar(50))+ ' بانك ' + dbo.TCOD_BANKS.NAMES + ' ' + ISNULL(dbo.PAY_GETP.SHOBEH, '  ') + ' مورخ ' + CAST(dbo.PAY_GETP.DATE_S AS NVARCHAR),60) AS Expr1 FROM  dbo.TCOD_BANKS INNER JOIN   dbo.PAY_GETP ON dbo.TCOD_BANKS.CODE = dbo.PAY_GETP.BANK WHERE      (dbo.PAY_GETP.N_KOL2 IS NULL) AND (dbo.PAY_GETP.N_KOL3 IS NULL) and kind <> 1 and dbo.PAY_GETP.N_KOL<> 911"
       DoCmd.RunSQL sql

       DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, " سند اختتاميه با موفقيت صادر گرديد"
  End If
  DoCmd.Close acForm, Me.NAME
End Sub

Private Sub Command6_Click()
                        On Error Resume Next
   DoCmd.Close acForm, Me.NAME
End Sub

Private Sub dt_BeforeUpdate(CANCEL As Integer)
                      On Error Resume Next: CANCEL = CHEKDATEM(dt, True)
End Sub

Private Sub Form_Load()
                   On Error Resume Next
  Call SETSECURITY(Me.NAME, "EKHTE", 3)
End Sub

Private Sub Form_Open(CANCEL As Integer)
                       On Error Resume Next
  Dim rst As New ADODB.Recordset, PATH As String, i As Integer
   rst.Open "SELECT Max(DEED_HED.N_S) AS MaxOfN_S FROM DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount > 0 Then
      S2 = rst.Fields(0) + 1
   Else
      S2 = 1
   End If
   rst.Close

End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                        On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
