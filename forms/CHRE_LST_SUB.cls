Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Compare Database
Option Explicit
 Dim DELT As Boolean

Private Sub BANK_Enter()
                      On Error Resume Next
   Me.BANK.RowSource = "SELECT dbo.PAY_GETD.BANK, dbo.TCOD_BANKS.NAMES FROM dbo.TCOD_BANKS INNER JOIN dbo.PAY_GETD ON dbo.TCOD_BANKS.CODE = dbo.PAY_GETD.BANK LEFT OUTER JOIN dbo.CHRE_LST ON dbo.PAY_GETD.N_SERI = dbo.CHRE_LST.N_SERI AND dbo.PAY_GETD.BANK = dbo.CHRE_LST.BANK AND  dbo.PAY_GETD.DATE_S = dbo.CHRE_LST.DATE_S WHERE (dbo.PAY_GETD.N_SERI = " & Me.N_SERI & ") And (dbo.CHRE_LST.N_SERI Is Null) GROUP BY dbo.PAY_GETD.BANK, dbo.TCOD_BANKS.NAMES"
   Me.BANK.Requery
End Sub

Private Sub BANK_Exit(CANCEL As Integer)
                     On Error Resume Next
   Me.BANK.RowSource = "SELECT PAY_GETD.BANK, TCOD_BANKS.NAMES FROM TCOD_BANKS INNER JOIN PAY_GETD ON TCOD_BANKS.CODE = PAY_GETD.BANK GROUP BY PAY_GETD.BANK, TCOD_BANKS.NAMES"
   Me.BANK.Requery
End Sub

Private Sub DATE_S_Enter()
                      On Error Resume Next
   Me.DATE_S.RowSource = "SELECT     TOP 100 PERCENT DATE_S AS DT FROM dbo.PAY_GETD WHERE (N_SERI = " & [Forms]![CHKREC_H]![CHRE_LST_SUB].[Form]![N_SERI] & ") And (BANK = " & [Forms]![CHKREC_H]![CHRE_LST_SUB].[Form]![BANK] & ") GROUP BY DATE_S ORDER BY DATE_S "
Me.DATE_S.Requery
End Sub

Private Sub date_s_Exit(CANCEL As Integer)
                      On Error Resume Next
   Me.DATE_S.RowSource = "SELECT     TOP 100 PERCENT DATE_S AS DT FROM dbo.PAY_GETD GROUP BY DATE_S ORDER BY DATE_S"
   Me.DATE_S.Requery
End Sub

Private Sub Form_AfterUpdate()
                       On Error Resume Next
   Forms![CHKREC_H]![LETSANAD] = True
    If Me.RecordsetClone.RecordCount > 0 Then
        Forms![CHKREC_H]![PRINTC].Enabled = True
    Else
        Forms![CHKREC_H]![PRINTC].Enabled = True
    End If
End Sub


Private Sub Form_BeforeUpdate(CANCEL As Integer)
                    On Error Resume Next
  If Not Me.NewRecord And Forms![BASEKNOW]![WAR] = 1 And Not DELT Then
    Forms![BASEKNOW]![Text44] = False
    DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "تغيرات داده شده ثبت شود؟"
    If Not Forms![BASEKNOW]![Text44] Then
        CANCEL = True
    End If
   End If
    If IsNull(Me.N_KOL3) And Not DELT Then
        Me.N_KOL3 = Forms![BASEKNOW]![BANKHA]
        Me.N_MOIN3 = Me.N_MOIN
        Me.N_TAF3 = Me.N_TAF
    End If
End Sub


Private Sub Form_Delete(CANCEL As Integer)
                     On Error Resume Next
    Dim rst As New ADODB.Recordset
   If Me.SelHeight > 1 Then
      DoCmd.OpenForm "mesag", , , , , acDialog, "اعلام وصول چند چك را نمي توانيد باهم حذف نمائيد. در هر بار فقط يك سطر قابل حذف است."
      CANCEL = True
      Exit Sub
   End If
   Forms![BASEKNOW]![Text44] = False
    DoCmd.OpenForm "deleteform", , , , , acDialog
    If Forms![BASEKNOW]![Text44] Then
        rst.Open "select * from pay_getd where n_seri =" & Me.N_SERI & " AND BANK = " & Me.BANK & " AND DATE_S = " & Me.DATE_S, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 1 Then
            rst.Fields("N_KOL3") = Null
            rst.Fields("N_MOIN3") = Null
            rst.Fields("N_TAF3") = Null
            rst.Fields("HES3") = Null
            rst.Fields("N_S") = Null
            rst.update
            DoCmd.RunSQL "DELETE FROM dbo.CHRE_LST WHERE     N_SERI = " & Me.N_SERI & " AND BANK = " & Me.BANK & " AND DATE_S = " & Me.DATE_S
            CANCEL = True
            DELT = True
            Call GETDLOG(1, Me.N_SERI, Me.BANK, Me.DATE_S, rst.Fields("SANDUGH"))
            rst.Close
        End If
    Else
     CANCEL = True
     DELT = True
    End If
   Forms![CHKREC_H]![LETSANAD] = True
   Me.Requery

End Sub

Private Sub Form_Open(CANCEL As Integer)
                    On Error Resume Next
   Me.N_SERI.RowSource = "SELECT N_SERI FROM PAY_GETD WHERE (N_KOL3 IS NULL) AND (NOT (N_MOIN IS NULL)) AND (N_KOL = " & Forms![BASEKNOW]![BANKHA] & ") AND (N_KOL2 IS NULL) GROUP BY N_SERI"
   Me.N_SERI.Requery
End Sub

Private Sub N_SERI_AfterUpdate()
                       On Error Resume Next
    Me.BANK.Requery
    Me.DATE_S.Requery
End Sub

Private Sub Form_Current()
                      On Error Resume Next
   DELT = False
'   Me.N_SERI.RowSource = "SELECT N_SERI FROM PAY_GETD WHERE (N_KOL3 IS NULL) AND (NOT (N_MOIN IS NULL)) AND (N_KOL = " & Forms![baseknow]![BANKHA] & ") AND (N_KOL2 IS NULL) GROUP BY N_SERI"
'   Me.N_SERI.Requery
'   Me.BANK.RowSource = "SELECT PAY_GETD.BANK, TCOD_BANKS.NAMES FROM TCOD_BANKS INNER JOIN PAY_GETD ON TCOD_BANKS.CODE = PAY_GETD.BANK GROUP BY PAY_GETD.BANK, TCOD_BANKS.NAMES"
'   Me.BANK.Requery
'   Me.DATE_S.RowSource = "SELECT Format([DATE_S],'##/##/##') AS DT FROM PAY_GETD GROUP BY Format([DATE_S],'##/##/##') ORDER BY Format([DATE_S],'##/##/##'"
'   Me.DATE_S.Requery
'   Me.N_TAF.RowSource = "SELECT     TDETA_HES.TNUMBER, TDETA_HES.NAME FROM TDETA_HES WHERE     (((TDETA_HES.NUMBER) = " & Me.N_MOIN & ") AND ((TDETA_HES.N_KOL)  = " & [Forms]![baseknow]![BANKHA] & ")) GROUP BY TDETA_HES.TNUMBER, TDETA_HES.NAME ORDER BY TDETA_HES.TNUMBER;"
'   Me.N_TAF.Requery
   Call SETSECURITYSUB("CHKREC_H", Me.NAME, "VCHD", 3)
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
