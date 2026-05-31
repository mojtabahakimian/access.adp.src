Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

Private Sub GroupHeader2_Format(CANCEL As Integer, FormatCount As Integer)
                     On Error Resume Next
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT NUMBER1,FNUMCO FROM HEAD_LST WHERE TAG = 27 AND NUMBER  = " & Me.NUMBER1, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
      Me.FFR = rst.Fields("FNUMCO")
    End If

End Sub

Private Sub Report_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
                       On Error Resume Next
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT PAY_GETD.N_SERI, TCOD_BANKS.NAMES, PAY_GETD.SHOBEH, PAY_GETD.DATE, PAY_GETD.DATE_S, PAY_GETD.MABL, PAY_GETD.NUMBER, PAY_GETD.TAG FROM TCOD_BANKS INNER JOIN PAY_GETD ON TCOD_BANKS.CODE = PAY_GETD.BANK WHERE (((PAY_GETD.NUMBER)=" & [Forms]![HEAD_LST_KH_BACK_azad]![NUMBER] & ") AND ((PAY_GETD.TAG)=" & [Forms]![HEAD_LST_KH_BACK_azad]![HTAG] & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
     rst.MoveLast
    rst.MoveFirst
      With rst
            COMM.CAPTION = "چكهاي دريافت  شده " & .RecordCount & " فقره جمعاًبه مبلغ :" & Format(Forms![HEAD_LST_KH_BACK_azad]![NCHK], "### ريال") & "  "
            Do Until .EOF()
                COMM.CAPTION = COMM.CAPTION & "ـ سريال:" & .Fields(0) & " بانك:" & .Fields(1) & " شعبه:" & Trim(.Fields(2))
                .MoveNext
            Loop
       End With
    Else
    End If
    rst.Close
   Me.Label224.CAPTION = "%ماليات و عوارض:"
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
