Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_BeforeInsert(CANCEL As Integer)
                   On Error Resume Next
  Dim rst As New ADODB.Recordset
   If IsNull(Me.KNUM) Then
        rst.Open "SELECT Max(KAR_HEAD.KNUM) AS KN FROM KAR_HEAD HAVING ((Not (Max(KAR_HEAD.KNUM)) Is Null))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            Me.KNUM = rst.Fields(0) + 1
        Else
            Me.KNUM = 1
        End If
   End If
End Sub

Private Sub Form_Current()
                   On Error Resume Next
  Call SETSECURITY(Me.NAME, "KARB", 3)
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                        On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Load()
                   On Error Resume Next
  Call SETSECURITY(Me.NAME, "KARB", 3)
End Sub

Private Sub Form_Open(CANCEL As Integer)
                   On Error Resume Next
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
