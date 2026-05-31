Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit

Private Sub Command8_Click()
                     On Error Resume Next
    Dim rst As New ADODB.Recordset
  Dim i As Integer
   If IsNull(pass1) Then
     pass1 = ""
   End If
   If IsNull(pass2) Then
     pass2 = ""
   End If

   If pass1 <> pass2 Then
     DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "خطا: كلمه عبور مساوي  نيست لطفاً مجددا سعي كنيد"
     pass1 = Null
     pass2 = Null
   Else
     If IsNull(usname) Then
       DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "خطا: کد كاربر صحيح نيست"
     Else
       If Len(pass1) > 40 Then
         DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "خطا:طول رمز  بايد كمتر از 40 كاراكتر باشد مانند: 110ab"
       Else
            rst.Open "select * from SALA_DTL where idd = " & usname, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount = 0 Then
                 DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "آي دي كاربر يا رمز عبور صحيح نمي باشد."
                 Exit Sub
            Else
                If OLDPASS = DECODEPS(rst.Fields("PSAL_NAME")) Then
                    rst.Fields("PSAL_NAME") = Trim(CODEPAL(pass1))
                    rst.update
                Else
                     DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "آي دي  كاربر يا رمز عبور صحيح نمي باشد."
                End If
                DoCmd.Close acForm, Me.NAME
            End If
       End If
     End If
  End If
End Sub

Private Sub Command9_Click()
                     On Error Resume Next
  DoCmd.Close
End Sub

Private Sub Form_Current()
                   On Error Resume Next
  Call SETSECURITY(Me.NAME, "NEWPASSWORD", 3)
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Load()
                   On Error Resume Next
  Call SETSECURITY(Me.NAME, "NEWPASSWORD", 3)
End Sub

Private Sub Form_Open(CANCEL As Integer)
Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub usname_AfterUpdate()
    Me.usna = GETUSERNAME(Me.usname)
End Sub
