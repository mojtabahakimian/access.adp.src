Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub ANBCO_DblClick(CANCEL As Integer)
                 On Error Resume Next
Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset
    RST2.Open "OPANBACCESS", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Open "select * from tcod_anbar where code > 0 ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
        RST2.AddNew
        RST2.Fields("USERCO") = Forms![F_USER_PERMITION FORMS]![USER_ANB_SUB].Form![IDD]
        RST2.Fields("ANBCO") = rst.Fields("CODE")
        RST2.update
        rst.MoveNext
    Wend
    Me.Requery
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                       On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"

End Sub

Private Sub Form_Open(CANCEL As Integer)
  Me.Refresh
End Sub
