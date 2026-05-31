Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                   On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"

End Sub

Private Sub Form_Open(CANCEL As Integer)
    Dim rst As New ADODB.Recordset
    SUBUSERCO.RowSource = "0;" & Chr(34) & "همه" & Chr(34) & ";"
    rst.Open "SELECT SALA_DTL.IDD, SAL_NAME  FROM SALA_DTL WHERE     (ENABL = 0)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
        SUBUSERCO.RowSource = SUBUSERCO.RowSource & rst.Fields("idd") & ";" & Chr(34) & DECODEUN(rst.Fields("SAL_NAME")) & Chr(34) & ";"
        rst.MoveNext
    Wend
    SUBUSERCO.RowSource = MID(SUBUSERCO.RowSource, 1, Len(SUBUSERCO.RowSource) - 1)
  Me.Refresh

End Sub

Private Sub SUBUSERCO_DblClick(CANCEL As Integer)
                 On Error Resume Next
Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset
    RST2.Open "CHARTSAZMANI", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Open "select * from sala_dtl where idd > 1 ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
        RST2.AddNew
        RST2.Fields("USERCO") = Forms![F_USER_PERMITION FORMS]![USER_USER_SUB].Form![IDD]
        RST2.Fields("SUBUSERCO") = rst.Fields("idd")
        RST2.update
        rst.MoveNext
    Wend
    Me.Requery
End Sub
