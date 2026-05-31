Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Compare Database
Option Explicit

Private Sub Report_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
                       On Error Resume Next
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT PAY_GETP.N_SERI, TCOD_BANKS.NAMES, PAY_GETP.SHOBEH, PAY_GETP.DATE, PAY_GETP.DATE_S, PAY_GETP.MABL, PAY_GETP.NUMBER, PAY_GETP.TAG FROM TCOD_BANKS INNER JOIN PAY_GETP ON TCOD_BANKS.CODE = PAY_GETP.BANK WHERE (((PAY_GETP.NUMBER)=" & [Forms]![HEAD_LST_FROOSH_BACK]![NUMBER] & ") AND ((PAY_GETP.TAG)=" & [Forms]![HEAD_LST_FROOSH_BACK]![HTAG] & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
     rst.MoveLast
     rst.MoveFirst
      With rst
            COMM.CAPTION = "چكهاي پرداخت شده " & .RecordCount & " فقره جمعاًبه مبلغ :" & Format(Forms![HEAD_LST_FROOSH_BACK]![NCHK], "### ريال") & "  "
            Do Until .EOF()
                COMM.CAPTION = COMM.CAPTION & "ـ سريال:" & .Fields(0) & " بانك:" & .Fields(1) & " شعبه:" & Trim(.Fields(2))
                .MoveNext
            Loop
       End With
    Else
    End If
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
