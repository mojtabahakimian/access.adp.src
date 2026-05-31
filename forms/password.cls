Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim appAccess As Access.Application, frm As Form


Private Sub Command2_Click()
                  On Error Resume Next
   Dim dbs As Database
  Select Case Text0
        Case "UNLOCK":
              '         If UCurrentUser() = ChrW(1605) & ChrW(1583) & ChrW(1740) & ChrW(1585) & ChrW(1587) & ChrW(1740) & ChrW(1587) & ChrW(1578) & ChrW(1605) Then
                          DoCmd.RunSQL ("UPDATE    dbo.DEED_HED SET  GHATEI = 0")
                          DoCmd.OpenForm "mesag", , , , , acDialog, "انجام شد.!"
               '        End If
        Case "amafhh": ChangeProperty "AllowBypassKey", False

        Case "sud":
                    If LETSGO("SUDP") Then
                        DoCmd.OpenForm "SURAT_MALI"
                    Else
                          DoCmd.OpenForm "mesag", , , , , acDialog, "شما اجازه لازم براي استفاده از اين بخش را نداريد.!"
                    End If

        Case "sudd":
                    If LETSGO("SUDP") Then
                        DoCmd.OpenForm "suddays", acFormPivotChart
                    Else
                          DoCmd.OpenForm "mesag", , , , , acDialog, "شما اجازه لازم براي استفاده از اين بخش را نداريد.!"
                    End If
        Case "suddd":
                    If LETSGO("SUDP") Then
                        DoCmd.OpenForm "suddays", acFormPivotChart, , , , , "1"
                    Else
                          DoCmd.OpenForm "mesag", , , , , acDialog, "شما اجازه لازم براي استفاده از اين بخش را نداريد.!"
                    End If
        Case "sudm":
                    If LETSGO("SUDP") Then
                        Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, rst3 As New ADODB.Recordset, SUDO As Double, i As Integer, K As Integer, dt11 As Long
                        dt11 = CLng(Forms![BASEKNOW]![YEA] & "0101")
                        CurrentProject.Connection.Execute ("DELETE FROM SUDmah")
                        RST2.Open "SELECT     dbo.Umonth(dbo.DEED_HED.DATE_S) AS mm, SUM(dbo.DEED_DTL.BED) AS BEDS, SUM(dbo.DEED_DTL.BES) AS BESS, SUM(dbo.DEED_DTL.BED - dbo.DEED_DTL.BES) AS sudday, 0 AS SUD  FROM         dbo.DEED_DTL INNER JOIN                       dbo.DEED_HED ON dbo.DEED_DTL.N_S = dbo.DEED_HED.N_S INNER JOIN               dbo.TOTA_HES ON dbo.DEED_DTL.HES_K = dbo.TOTA_HES.NUMBER INNER JOIN         dbo.TARAZHES ON dbo.TOTA_HES.NUMBER = dbo.TARAZHES.NUMBER GROUP BY dbo.Umonth(dbo.DEED_HED.DATE_S) order by dbo.Umonth(dbo.DEED_HED.DATE_S)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                        rst.Open "SUDmah", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                        While Not RST2.EOF
                            SUDO = SUDO + RST2.Fields("sudday")
                            rst.AddNew
                            rst.Fields("DATE_S") = RST2.Fields("mm")
                            rst.Fields("BEDS") = RST2.Fields("BEDS")
                            rst.Fields("BESS") = RST2.Fields("BESS")
                            rst.Fields("sudday") = RST2.Fields("sudday")
                            rst.Fields("SUD") = SUDO
                            rst.update
                            RST2.MoveNext
                            Set rst3 = New ADODB.Recordset
                        Wend
                        DoCmd.OpenForm "sudmah", acFormPivotChart, , , , , "1"
                    Else
                        DoCmd.OpenForm "mesag", , , , , acDialog, "شما اجازه لازم براي استفاده از اين بخش را نداريد.!"
                    End If
        Case "CORRECT":
                    DoCmd.OpenForm "board1"
                     DoCmd.OpenForm "form2"
                 Forms![BASEKNOW]![mrcorrect] = True
     Case "yafatemah": DoCmd.RunMacro "startup"
                            ChangeProperty "AllowBypassKey", True
     Case "tables":
                               DoCmd.OpenForm "databasetb", acFormDS
     Case "filter":
                          DoCmd.OpenForm "mesag", , , , , acDialog, frm.ServerFilter & "    RS : " & frm.RecordSource
     Case "sanati":
                               DoCmd.OpenForm "zjazbenheraf", , , , , acDialog
     Case "zparam":
                               DoCmd.OpenForm "zinvoextended", acFormDS
     Case "enddate":
                               Forms![About]![Whether].Locked = False
     Case Else
           If Text0 = "SANADSKN" Then
              DoCmd.OpenForm "SANAD_AMALKARD", , , , , acDialog
            End If
   End Select
   DoCmd.Close acForm, "password"
End Sub


Private Sub Command3_Click()
                   On Error Resume Next
  DoCmd.Quit
End Sub


Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
 Sub SetStartupProperties()
                   On Error Resume Next
   '   ChangeProperty " StartupForm ", dbText, "Customers"
   '   ChangeProperty " StartupShowDBWindow ", dbBoolean, False
   '   ChangeProperty " StartupShowStatusBar ", dbBoolean, False
   '   ChangeProperty " AllowBuiltinToolbars ", dbBoolean, False
   '   ChangeProperty " AllowFullMenus ", dbBoolean, True
   '   ChangeProperty " AllowBreakIntoCode ", dbBoolean, False
   '   ChangeProperty " AllowSpecialKeys ", dbBoolean, True
      ChangeProperty " AllowBypassKey ", True

 End Sub

  Function ChangeProperty(strPropName As String, varPropValue As Variant) As Integer
      Dim dbs As Database, prp As Property, PR As Application
      Const conPropNotFoundError = 2455
    '  Call PRRR
       On Error GoTo Change_Err
      CurrentProject.Properties(strPropName) = varPropValue
      ChangeProperty = True

Change_Bye:
      Exit Function

Change_Err:
      If Err = conPropNotFoundError Then      ' Property not found.
         CurrentProject.Properties.ADD strPropName, varPropValue
          Resume Next
      Else
          ' Unknown error.
          ChangeProperty = False
          Resume Change_Bye
      End If
  End Function






Public Sub PRRR()
   Dim i As Integer
   For i = 0 To 200
     ' DEBUG.PRINT CurrentDb.Properties(i).NAME
     ' DEBUG.PRINT CurrentProject.Properties(i)
   Next i

End Sub

Private Sub Form_Open(CANCEL As Integer)
                     On Error Resume Next
  Set frm = Screen.ActiveForm
End Sub
