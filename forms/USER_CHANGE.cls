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
     If IsNull(usname) Then
       DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "خطا: نام كاربر صحيح نيست"
     Else
       If Len(pass1) > 40 Then
         DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "خطا:طول رمز  بايد كمتر از 40 كاراكتر باشد مانند: 110ab"
       Else
          If Left(UCurrentUser(), 10) = ChrW(1605) & ChrW(1583) & ChrW(1740) & ChrW(1585) & " " & ChrW(1587) & ChrW(1740) & ChrW(1587) & ChrW(1578) & ChrW(1605) Or Left(UCurrentUser(), 10) = "مدير سيستم" Then
               If usname = ChrW(1605) & ChrW(1583) & ChrW(1740) & ChrW(1585) & ChrW(1587) & ChrW(1740) & ChrW(1587) & ChrW(1578) & ChrW(1605) Or usname = "َAdminister" Then
                  DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "اين كاربر قابل حذف يا تغيير نام نيست"
                Else
                    rst.Open "SALA_DTL", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    rst.Filter = "SAL_NAME = '" & Trim(CODESAL(Fixp(usname))) & "'"
                    Set rst = rst
                    If rst.RecordCount = 0 Then
                         DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "نام كاربر يا رمز عبور صحيح نمي باشد."
                         Exit Sub
                    Else
                        If NEWUSER = "" Or IsNull(NEWUSER) Then
                            Err.Clear
                            rst.delete
                            If Err <> 0 Then
                                DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "عمليات ناموفق بود"
                            Else
                                DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "كاربر حذف شد"
                            End If
                        Else
                            Err.Clear
                            rst.Fields("SAL_NAME") = Trim(CODESAL(Fixp(NEWUSER)))
                            rst.update
                            If Err <> 0 Then
                                DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "عمليات ناموفق بود"
                            Else
                                 DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "نام كاربر با موفقيت اصلاح شد"
                            End If
                         End If
                    End If
                 End If
          Else
            rst.Open "SALA_DTL", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            rst.Filter = "SAL_NAME = '" & Trim(CODESAL(Fixp(usname))) & "'"
            Set rst = rst
            If rst.RecordCount = 0 Then
                 DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "نام كاربر يا رمز عبور صحيح نمي باشد."
                 Exit Sub
            Else
                If pass1 = DECODEPS(rst.Fields("PSAL_NAME")) Then
                    If usname = ChrW(1605) & ChrW(1583) & ChrW(1740) & ChrW(1585) & ChrW(1587) & ChrW(1740) & ChrW(1587) & ChrW(1578) & ChrW(1605) Or usname = "َAdminister" Then
                        DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "اين كاربر قابل حذف يا تغيير نام نيست"
                    Else
                        If NEWUSER = "" Or IsNull(NEWUSER) Then
                            Err.Clear
                            rst.delete
                            If Err <> 0 Then
                                DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "عمليات ناموفق بود"
                            Else
                                DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "كاربر حذف شد"
                            End If
                        Else
                            Err.Clear
                            rst.Fields("SAL_NAME") = Trim(CODESAL(Fixp(NEWUSER)))
                            rst.update
                            If Err <> 0 Then
                                DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "عمليات ناموفق بود"
                            Else
                                 DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "نام كاربر با موفقيت اصلاح شد"
                            End If
                        End If
                     End If
                Else
                     DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "نام كاربر يا رمز عبور صحيح نمي باشد."
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
