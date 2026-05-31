Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Command8_Click()
                   On Error Resume Next
  Dim i As Integer, rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, IDD As Long
   If IsNull(pass1) Then
     pass1 = ""
   End If
   If IsNull(pass2) Then
     pass2 = ""
   End If
   RST2.Open "select max(idd) as idd from sala_dtl", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If RST2.RecordCount > 0 Then
     IDD = RST2.Fields(0) + 1
   Else
     IDD = 1
   End If
   If pass1 <> pass2 Then
     DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "خطا: كلمه عبور مساوي  نيست لطفاً مجددا سعي كنيد"
     pass1 = Null
     pass2 = Null
   Else
     If IsNull(usname) Then
       DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "خطا: نام كاربر صحيح نيست"
     Else
        rst.Open "SALA_DTL", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        rst.AddNew
        rst.Fields("SAL_NAME") = Trim(CODESAL(Fixp(usname)))
        rst.Fields("PSAL_NAME") = Trim(CODEPAL(pass1))
        rst.Fields("GRSAL") = Me.GRP
        rst.Fields("IDD") = IDD
        rst.update
        If Err <> 0 Then
        Else
            DoCmd.RunSQL "INSERT INTO dbo.SAL_CHEK  (USERCO, OBJECT,RUN) SELECT     " & IDD & " AS USERCO, IDH AS OBJECT,0 AS RUN FROM  dbo.TFORMS"
        End If
        rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "SIGN", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        rst.AddNew
        rst.Fields("USERCO") = IDD
        rst.update
        DoCmd.Close acForm, Me.NAME
     End If
  End If
End Sub

Private Sub Command9_Click()
                     On Error Resume Next
  DoCmd.Close
End Sub

Private Sub Form_Current()
                   On Error Resume Next
  Call SETSECURITY(Me.NAME, "USERS", 3)
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Load()
                   On Error Resume Next
  Call SETSECURITY(Me.NAME, "USERS", 3)
End Sub
