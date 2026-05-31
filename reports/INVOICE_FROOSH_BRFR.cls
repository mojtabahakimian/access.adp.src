Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Compare Database
Option Explicit

Private Sub GroupFooter3_Format(CANCEL As Integer, FormatCount As Integer)
  Call Showemza(Me, "FFRB_FROOSHTX", "FFRB_ANBTX", "FFRB_HESABTX")

End Sub

Private Sub Report_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
                     On Error Resume Next
    Dim rst As New ADODB.Recordset, JST As New ADODB.Recordset, JCHK, JAMF, HAZ, NAGHD, VAR, HAV, taf, MBAA As Double
    rst.Open "SELECT PAY_GETP.N_SERI, TCOD_BANKS.NAMES, PAY_GETP.SHOBEH, PAY_GETP.DATE, PAY_GETP.DATE_S, PAY_GETP.MABL, PAY_GETP.NUMBER, PAY_GETP.TAG FROM TCOD_BANKS INNER JOIN PAY_GETP ON TCOD_BANKS.CODE = PAY_GETP.BANK WHERE (((PAY_GETP.NUMBER)=" & [Forms]![head_lst_brfr]![NUMBER] & ") AND ((PAY_GETP.TAG)=24))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
      rst.MoveLast
      rst.MoveFirst
      JCHK = 0
      With rst
            COMM.CAPTION = "چكهاي پرداخت شده " & .RecordCount & " فقره جمعاًبه مبلغ :" & Format(Forms![head_lst_brfr]![NCHK], "### ريال") & "  "
            Do Until .EOF()
                COMM.CAPTION = COMM.CAPTION & "ـ سريال:" & .Fields(0) & " بانك:" & .Fields(1) & " شعبه:" & Trim(.Fields(2))
                 JCHK = JCHK + rst.Fields("MABL")
                .MoveNext
            Loop
       End With
    Else
    End If
   JAMF = 0
   HAZ = 0
   NAGHD = 0
   VAR = 0
   HAV = 0
   taf = 0
   MBAA = 0
   rst.Close
   JST.Open "SELECT Sum(INVO_LST.MABL_K) AS SumOfMABL_K FROM INVO_LST WHERE (((INVO_LST.NUMBER)= " & [Forms]![head_lst_brfr]![NUMBER] & " ) AND ((INVO_LST.TAG)=24))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
        JAMF = JST.Fields(0)
   End If
   JST.Close
   JST.Open "SELECT HEAD_LST.NUMBER, HEAD_LST.TAG AS htag, HEAD_LST.ANBAR, HEAD_LST.NUMBER1, HEAD_LST.DATE_N, HEAD_LST.TAH, HEAD_LST.MAS, HEAD_LST.VAS, HEAD_LST.N_S, HEAD_LST.CUST_NO, HEAD_LST.MOLAH, HEAD_LST.M_NAGHD, HEAD_LST.MABL_VAR, HEAD_LST.MOIN_VAR, HEAD_LST.MABL_HAV, HEAD_LST.MOIN_HAV, HEAD_LST.MABL_HAZ, HEAD_LST.MOIN_HAZ, HEAD_LST.TAKHFIF, HEAD_LST.MOIN_KHF, HEAD_LST.ANBARF, HEAD_LST.FNUMCO, HEAD_LST.MBAA FROM HEAD_LST WHERE (((HEAD_LST.NUMBER)= " & [Forms]![head_lst_brfr]![NUMBER] & " ) AND  ((HEAD_LST.TAG)=25))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
        HAZ = JST.Fields("MABL_HAZ")
        VAR = JST.Fields("MABL_VAR")
        HAV = JST.Fields("MABL_HAV")
        NAGHD = JST.Fields("M_NAGHD")
        taf = JST.Fields("TAKHFIF")
        MBAA = JST.Fields("MBAA")
   End If
   JST.Close
   Me.JF.CAPTION = Format(JAMF, "#,##0;#,##0-")
   Me.HKH.CAPTION = Format(HAZ, "#,##0;#,##0-")
   Me.MBAA.CAPTION = Format(MBAA, "#,##0;#,##0-")
   Me.TF.CAPTION = Format(taf, "#,##0;#,##0-")
   Me.GABEL.CAPTION = Format(JAMF + HAZ - taf + MBAA, "#,##0;#,##0-")
   Me.JPAY.CAPTION = Format(NAGHD + VAR + HAV + JCHK, "#,##0;#,##0-")
   Me.MAN.CAPTION = Format(JAMF + HAZ + MBAA - (NAGHD + VAR + HAV + JCHK + taf), "#,##0;#,##0-")
   Me.HR.CAPTION = ALPHANUM(JAMF + HAZ - taf + MBAA) & " " & "ريال"
   Me.Label202.CAPTION = "%ماليات و عوارض:"
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
