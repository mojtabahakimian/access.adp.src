Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Current()
                   On Error Resume Next
  Call SETSECURITY(Me.NAME, "RAAS", 3)
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Load()
                   On Error Resume Next
  Call SETSECURITY(Me.NAME, "RAAS", 3)
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub Text0_AfterUpdate()
Dim rst As New ADODB.Recordset
If Not IsNull(Me.Text0) Then
    rst.Open "raas", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
            rst.Fields("ROOS") = DIFF(Text0, rst.Fields("DATE_S"))
            rst.Fields("MABK") = rst.Fields("ROOS") * rst.Fields("MABL")
            rst.update
            rst.MoveNext
        Wend
        Me.RAAS_SUB.Requery
 End If
End Sub

Private Sub Text0_Exit(CANCEL As Integer)
                  On Error Resume Next
 If IsNull(Text0) Then
   CANCEL = True
 Else
                      On Error Resume Next: CANCEL = (CHEKDATEM(Text0, False))
 End If
End Sub
