Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database
Option Explicit
 Dim DELT As Boolean
Private Sub BANK_Enter()
                      On Error Resume Next
   Me.BANK.RowSource = "SELECT dbo.PAY_GETP.BANK, dbo.TCOD_BANKS.NAMES FROM dbo.TCOD_BANKS INNER JOIN dbo.PAY_GETP ON dbo.TCOD_BANKS.CODE = dbo.PAY_GETP.BANK LEFT OUTER JOIN dbo.CHRE_LSP ON dbo.PAY_GETP.N_SERI = dbo.CHRE_LSP.N_SERI AND dbo.PAY_GETP.BANK = dbo.CHRE_LSP.BANK AND  dbo.PAY_GETP.DATE_S = dbo.CHRE_LSP.DATE_S WHERE (dbo.PAY_GETP.N_SERI = " & Me.N_SERI & ") And (dbo.CHRE_LSP.N_SERI Is Null) GROUP BY dbo.PAY_GETP.BANK, dbo.TCOD_BANKS.NAMES"
   Me.BANK.Requery
End Sub
Private Sub BANK_Exit(CANCEL As Integer)
                    On Error Resume Next
   Me.BANK.RowSource = "SELECT PAY_GETP.BANK, TCOD_BANKS.NAMES FROM TCOD_BANKS INNER JOIN PAY_GETP ON TCOD_BANKS.CODE = PAY_GETP.BANK GROUP BY PAY_GETP.BANK, TCOD_BANKS.NAMES"
   Me.BANK.Requery
End Sub
Private Sub Form_Current()
                      On Error Resume Next
   DELT = False
   Me.N_SERI.RowSource = "SELECT N_SERI FROM PAY_GETP WHERE (N_KOL3 IS NULL) AND (NOT (N_MOIN IS NULL)) AND (N_KOL = " & Forms![BASEKNOW]![BANKHA] & ") AND (N_KOL2 IS NULL) GROUP BY N_SERI"
   Me.N_SERI.Requery
   Me.BANK.RowSource = "SELECT PAY_GETP.BANK, TCOD_BANKS.NAMES FROM TCOD_BANKS INNER JOIN PAY_GETP ON TCOD_BANKS.CODE = PAY_GETP.BANK GROUP BY PAY_GETP.BANK, TCOD_BANKS.NAMES"
   Me.BANK.Requery
   Me.N_TAF.RowSource = "SELECT     TDETA_HES.TNUMBER, TDETA_HES.NAME FROM TDETA_HES WHERE     (((TDETA_HES.NUMBER) = " & Me.N_MOIN & ") AND ((TDETA_HES.N_KOL)  = " & [Forms]![BASEKNOW]![BANKHA] & ")) GROUP BY TDETA_HES.TNUMBER, TDETA_HES.NAME ORDER BY TDETA_HES.TNUMBER;"
   Me.N_TAF.Requery
   Call SETSECURITYSUB("CHREC_HP", Me.NAME, "VCHP", 3)
End Sub
Private Sub Form_Open(CANCEL As Integer)
                    On Error Resume Next
   Me.N_SERI.RowSource = "SELECT N_SERI FROM PAY_GETP WHERE (N_KOL3 IS NULL) AND (NOT (N_MOIN IS NULL)) AND (N_KOL = " & Forms![BASEKNOW]![BANKHA] & ") AND (N_KOL2 IS NULL) GROUP BY N_SERI"
End Sub
Private Sub N_SERI_AfterUpdate()
                        On Error Resume Next
    Me.BANK.Requery
    Me.DATE_S.Requery
End Sub
Private Sub Form_AfterUpdate()
                       On Error Resume Next
        Forms![CHREC_HP]![LETSANAD] = True
    If Me.RecordsetClone.RecordCount > 0 Then
        Forms![CHREC_HP]![PRINTC].Enabled = True
    Else
        Forms![CHREC_HP]![PRINTC].Enabled = True
    End If
Me.N_TAF.Requery
End Sub

Private Sub Form_BeforeDelConfirm(CANCEL As Integer, Response As Integer)
                      On Error Resume Next
   Response = acDataErrContinue
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
        rst.Open "select * from pay_getP where n_seri =" & Me.N_SERI & " AND BANK = " & Me.BANK & " AND DATE_S = " & Me.DATE_S, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 1 Then
            rst.Fields("N_KOL3") = Null
            rst.Fields("N_MOIN3") = Null
            rst.Fields("N_TAF3") = Null
            rst.Fields("N_S") = Null
            rst.update
            rst.Close
            DoCmd.RunSQL "DELETE FROM dbo.CHRE_LSP WHERE     N_SERI = " & Me.N_SERI & " AND BANK = " & Me.BANK & " AND DATE_S = " & Me.DATE_S
            CANCEL = True
            DELT = True
        End If
    Else
     CANCEL = True
     DELT = True
    End If
    Forms![CHREC_HP]![LETSANAD] = True
    Me.Requery
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
Private Sub N_TAF_Enter()
                    On Error Resume Next
   Me.N_TAF.RowSource = "SELECT     TDETA_HES.TNUMBER, TDETA_HES.NAME FROM TDETA_HES WHERE     (((TDETA_HES.NUMBER) = " & Me.N_MOIN & ") AND ((TDETA_HES.N_KOL)  = " & [Forms]![BASEKNOW]![BANKHA] & ")) GROUP BY TDETA_HES.TNUMBER, TDETA_HES.NAME ORDER BY TDETA_HES.TNUMBER;"
End Sub
