Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub FIRSTNUM_BeforeUpdate(CANCEL As Integer)
                 On Error Resume Next
   Dim rst As New ADODB.Recordset
   If Not Me.NewRecord Then
        rst.Open "SELECT  ANBAR FROM PAY_GETD WHERE ANBAR = " & Me.BOOKNUM, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
          CANCEL = True
        End If
   End If
End Sub

Private Sub Form_Delete(CANCEL As Integer)
                 On Error Resume Next
   Dim rst As New ADODB.Recordset
   rst.Open "SELECT  ANBAR FROM PAY_GETD WHERE ANBAR = " & Me.BOOKNUM, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount > 0 Then
     CANCEL = True
   End If
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
