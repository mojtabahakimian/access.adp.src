Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Report_Close()
                    On Error Resume Next
   DoCmd.Close acForm, "F_MENU_KOL_MOIN_tafzil"
End Sub
Private Sub Report_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
                   On Error Resume Next
   Dim rst As New ADODB.Recordset
     rst.Open "SELECT Sum(DEED_DTL.BED) AS SumOfBED, Sum(DEED_DTL.BES) AS SumOfBES  FROM DEED_DTL WHERE (((DEED_DTL.HES_K)=" & Forms![F_MENU_KOL_MOIN_TAFZIL]![HKOL] & ") AND ((DEED_DTL.HES_M)=" & Forms![F_MENU_KOL_MOIN_TAFZIL]![HMOIN] & ") AND ((DEED_DTL.HES_t)=" & Forms![F_MENU_KOL_MOIN_TAFZIL]![HTAF] & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     If rst.RecordCount > 0 Then
        SBK.CAPTION = Format(rst.Fields(0), "#,###")
        SBS.CAPTION = Format(rst.Fields(1), "#,###")
        SBSB.CAPTION = Format(Abs(rst.Fields(0) - rst.Fields(1)), "#,###")
        TASH.CAPTION = IIf(rst.Fields(0) - rst.Fields(1) > 0, "بد", "بس")
     End If
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
