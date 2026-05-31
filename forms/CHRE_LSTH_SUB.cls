Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Compare Database
Option Explicit
 Dim DELT As Boolean

Private Sub BANK_Enter()
                      On Error Resume Next
   Me.BANK.RowSource = "SELECT PAY_GETD.BANK, TCOD_BANKS.NAMES FROM TCOD_BANKS INNER JOIN  PAY_GETD ON TCOD_BANKS.CODE = PAY_GETD.BANK WHERE (PAY_GETD.N_SERI = " & Me.N_SERI & ") GROUP BY PAY_GETD.BANK, TCOD_BANKS.NAMES"
   Me.BANK.Requery
End Sub

Private Sub BANK_Exit(CANCEL As Integer)
                     On Error Resume Next
   Me.BANK.RowSource = "SELECT PAY_GETD.BANK, TCOD_BANKS.NAMES FROM TCOD_BANKS INNER JOIN PAY_GETD ON TCOD_BANKS.CODE = PAY_GETD.BANK GROUP BY PAY_GETD.BANK, TCOD_BANKS.NAMES"
   Me.BANK.Requery
End Sub

Private Sub DATE_S_Enter()
                      On Error Resume Next
    Me.DATE_S.RowSource = "SELECT     TOP 100 PERCENT DATE_S AS DT FROM dbo.PAY_GETD WHERE (N_SERI = " & Me.N_SERI & ") And (BANK = " & Me.BANK & ") GROUP BY DATE_S ORDER BY DATE_S "
    Me.DATE_S.Requery
End Sub

Private Sub date_s_Exit(CANCEL As Integer)
                      On Error Resume Next
   Me.DATE_S.RowSource = "SELECT     TOP 100 PERCENT DATE_S AS DT FROM dbo.PAY_GETD GROUP BY DATE_S ORDER BY DATE_S"
   Me.DATE_S.Requery
End Sub

Private Sub Form_AfterUpdate()
                       On Error Resume Next
   Forms![CHKREC_HES]![LETSANAD] = True
    If Me.RecordsetClone.RecordCount > 0 Then
        Forms![CHKREC_HES]![PRINTC].Enabled = True
    Else
        Forms![CHKREC_HES]![PRINTC].Enabled = True
    End If
    Me.N_TAF.Requery
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
End Sub


Private Sub Form_Delete(CANCEL As Integer)
                      On Error Resume Next
   Dim rst As New ADODB.Recordset
    Forms![BASEKNOW]![Text44] = False

    DoCmd.OpenForm "deleteform", , , , , acDialog
    If Forms![BASEKNOW]![Text44] Then
        Me.N_S = Null
        Me.N_KOL = Null
        Me.N_MOIN = Null
        Me.N_TAF = Null
        Me.HES = Null
        DELT = True
        DoCmd.RunSQL "DELETE FROM dbo.CHRE_LSPH WHERE     (N_SERI = " & Me.N_SERI & " ) AND (BANK = " & Me.BANK & ") AND (DATE_S = " & Me.DATE_S & ")"
        CANCEL = True
        Me.Requery
    Else
     CANCEL = True
     DELT = False
    End If
   Forms![CHKREC_HES]![LETSANAD] = True
End Sub

Private Sub N_SERI_AfterUpdate()
                     On Error Resume Next
    Me.BANK.Requery
    Me.DATE_S.Requery
End Sub

Private Sub Form_Current()
                      On Error Resume Next
   DELT = False
   Me.BANK.RowSource = "SELECT PAY_GETD.BANK, TCOD_BANKS.NAMES FROM TCOD_BANKS INNER JOIN PAY_GETD ON TCOD_BANKS.CODE = PAY_GETD.BANK GROUP BY PAY_GETD.BANK, TCOD_BANKS.NAMES"
   Me.BANK.Requery
   Me.DATE_S.RowSource = "SELECT Format([DATE_S],'##/##/##') AS DT FROM PAY_GETD GROUP BY Format([DATE_S],'##/##/##') ORDER BY Format([DATE_S],'##/##/##'"
   Me.DATE_S.Requery
  ' Me.N_TAF.RowSource = "SELECT     TDETA_HES.TNUMBER, TDETA_HES.NAME FROM TDETA_HES WHERE     (((TDETA_HES.NUMBER) = 1) AND ((TDETA_HES.N_KOL)  = " & [Forms]![baseknow]![BANKHA] & ")) GROUP BY TDETA_HES.TNUMBER, TDETA_HES.NAME ORDER BY TDETA_HES.TNUMBER"
  ' Me.N_TAF.Requery
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub N_SERI_DblClick(CANCEL As Integer)
                     On Error Resume Next
   Dim rst As New ADODB.Recordset
    Forms![BASEKNOW]![Text44] = False

    DoCmd.OpenForm "deleteform", , , , , acDialog
    If Forms![BASEKNOW]![Text44] Then
        DELT = True
        DoCmd.RunSQL "update PAY_GETD  SET  N_S = Null,N_KOL = Null,N_MOIN = Null,N_TAF = Null,HES1 = Null WHERE     (N_SERI = " & Me.N_SERI & " ) AND (BANK = " & Me.BANK & ") AND (DATE_S = " & Me.DATE_S & ")"
        DoCmd.RunSQL "DELETE FROM dbo.CHRE_LSPH WHERE     (N_SERI = " & Me.N_SERI & " ) AND (BANK = " & Me.BANK & ") AND (DATE_S = " & Me.DATE_S & ")"
        CANCEL = True
        Me.Requery
    Else
     CANCEL = True
     DELT = False
    End If
   Forms![CHKREC_HES]![LETSANAD] = True

End Sub
