Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

Private Sub Command14_Click()
                      On Error Resume Next
   DoCmd.Close
End Sub

Private Sub Command2_Click()
                      On Error Resume Next
   Err.Clear
   Dim errorno As Integer, sql As String, RST2 As New ADODB.Recordset, rst As New ADODB.Recordset, SERH As New ADODB.Recordset
   Dim CON As Long, i As Long, SERNUM As Long
   Select Case Me.OpenArgs
     Case "FO":
                            On Error Resume Next
         rst.Open "SELECT Count(OHEAD_LST.NUMBER) AS CountOfNUMBER FROM OHEAD_LST WHERE (((OHEAD_LST.TAG)=2))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
         Me.Text0 = "n"
         Me.Text0.Visible = True
         Me.Box7.Visible = True
         Me.Refresh
         Me.Text0 = Me.Text0 + "n"
         Me.Refresh
         Me.Text0.Requery
         Me.Repaint
          CON = rst.Fields(0) / 100
         If CON < 1 Then
            CON = 1
         End If
        i = 1
         Randomize
         rst.Open "SELECT OHEAD_LST.* FROM OHEAD_LST WHERE (((OHEAD_LST.TAG)=2))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
         rst.MoveLast
         rst.MoveFirst
         Err.Clear
         While Not rst.EOF()
              RST2.Open "SELECT Max(HEAD_LST.NUMBER) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=2))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
              If RST2.RecordCount = 0 Or IsNull(RST2.Fields(0)) Then
                  SERNUM = 1
              Else
                  SERNUM = RST2.Fields(0) + 1
              End If
              RST2.Close
              SERH.Open "HEAD_LST", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
              SERH.AddNew
              SERH.Fields("NUMBER") = SERNUM
              SERH.Fields("TAG") = 2
              SERH.Fields("MOLAH") = rst.Fields("MOLAH") & "  " & rst.Fields("NUMBER")
              SERH.Fields("USER_NAME") = rst.Fields("USER_NAME")
              SERH.Fields("CUST_KIND") = rst.Fields("CUST_KIND")
              SERH.Fields("SHIFT") = rst.Fields("SHIFT")
              SERH.Fields("DEPATMAN") = rst.Fields("DEPATMAN")
              SERH.Fields("MOIN_KHF") = rst.Fields("MOIN_KHF")
              SERH.Fields("TAKHFIF") = rst.Fields("TAKHFIF")
              SERH.Fields("MOIN_HAZ") = rst.Fields("MOIN_HAZ")
              SERH.Fields("MABL_HAZ") = rst.Fields("MABL_HAZ")
              SERH.Fields("MOIN_HAV") = rst.Fields("MOIN_HAV")
              SERH.Fields("MABL_HAV") = rst.Fields("MABL_HAV")
              SERH.Fields("MOIN_VAR") = rst.Fields("MOIN_VAR")
              SERH.Fields("MABL_VAR") = rst.Fields("MABL_VAR")
              SERH.Fields("M_NAGHD") = rst.Fields("M_NAGHD")
              SERH.Fields("CUST_NO") = rst.Fields("CUST_NO")
              SERH.Fields("VAS") = rst.Fields("VAS")
              SERH.Fields("MAS") = rst.Fields("MAS")
              SERH.Fields("DATE_N") = rst.Fields("DATE_N")
              SERH.update
              Me.SERN = SERNUM
              Me.ONUM = rst.Fields("NUMBER")
              DoCmd.OpenQuery "OAPPEN_INVO"
              DoCmd.OpenQuery "OAPPEN_PAYGETD"
              Call GENSANADFROOSH(SERNUM, SERNUM)
              rst.delete
              If (i Mod CON) = 0 Then
                  Me.Text0 = Me.Text0 + "n"
                  Me.Refresh
                  Me.Text0.Requery
                  Me.Repaint
              End If
              Me.TEDAD = i
              Me.Repaint
              i = i + 1
          rst.MoveNext
         Wend
    Case "FK":
                            On Error Resume Next
         rst.Open "SELECT Count(OHEAD_LST.NUMBER) AS CountOfNUMBER FROM OHEAD_LST WHERE (((OHEAD_LST.TAG)=14))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
         Me.Text0 = "n"
         Me.Text0.Visible = True
         Me.Box7.Visible = True
         Me.Refresh
         Me.Text0 = Me.Text0 + "n"
         Me.Refresh
         Me.Text0.Requery
         Me.Repaint
          CON = rst.Fields(0) / 100
         If CON < 1 Then
            CON = 1
         End If
        i = 1
         Randomize
         rst.Open "SELECT OHEAD_LST.* FROM OHEAD_LST WHERE (((OHEAD_LST.TAG)=14))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
         rst.MoveLast
         rst.MoveFirst
         Err.Clear
         While Not rst.EOF()
              RST2.Open "SELECT Max(HEAD_LST.NUMBER) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=14))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
              If RST2.RecordCount = 0 Or IsNull(RST2.Fields(0)) Then
                  SERNUM = 1
              Else
                  SERNUM = RST2.Fields(0) + 1
              End If
              RST2.Close
              SERH.Open "HEAD_LST", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
              SERH.AddNew
              SERH.Fields("NUMBER") = SERNUM
              SERH.Fields("TAG") = 14
              SERH.Fields("MOLAH") = rst.Fields("MOLAH") & "  " & rst.Fields("NUMBER")
              SERH.Fields("USER_NAME") = rst.Fields("USER_NAME")
              SERH.Fields("CUST_KIND") = rst.Fields("CUST_KIND")
              SERH.Fields("SHIFT") = rst.Fields("SHIFT")
              SERH.Fields("DEPATMAN") = rst.Fields("DEPATMAN")
              SERH.Fields("MOIN_KHF") = rst.Fields("MOIN_KHF")
              SERH.Fields("TAKHFIF") = rst.Fields("TAKHFIF")
              SERH.Fields("MOIN_HAZ") = rst.Fields("MOIN_HAZ")
              SERH.Fields("MABL_HAZ") = rst.Fields("MABL_HAZ")
              SERH.Fields("MOIN_HAV") = rst.Fields("MOIN_HAV")
              SERH.Fields("MABL_HAV") = rst.Fields("MABL_HAV")
              SERH.Fields("MOIN_VAR") = rst.Fields("MOIN_VAR")
              SERH.Fields("MABL_VAR") = rst.Fields("MABL_VAR")
              SERH.Fields("M_NAGHD") = rst.Fields("M_NAGHD")
              SERH.Fields("CUST_NO") = rst.Fields("CUST_NO")
              SERH.Fields("VAS") = rst.Fields("VAS")
              SERH.Fields("MAS") = rst.Fields("MAS")
              SERH.Fields("DATE_N") = rst.Fields("DATE_N")
              SERH.update
              Me.SERN = SERNUM
              Me.ONUM = rst.Fields("NUMBER")
              DoCmd.OpenQuery "OAPPEN_INVOKH"
              DoCmd.OpenQuery "OAPPEN_PAYGETDKH"
              Call SANADKHAD(SERNUM, SERNUM)
              rst.delete
              If (i Mod CON) = 0 Then
                  Me.Text0 = Me.Text0 + "n"
                  Me.Refresh
                  Me.Text0.Requery
                  Me.Repaint
              End If
              Me.TEDAD = i
              Me.Repaint
              i = i + 1
              rst.MoveNext
         Wend
End Select
   If Err = 0 Then
    DoCmd.Hourglass False
    DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "عمليات انتقال فاكتورها با موفقيت انجام شد"
    Me.Text0.Visible = False
    Me.Box7.Visible = False
    Me.Refresh
    DoCmd.Close acForm, Me.NAME
   End If

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Timer()
                     On Error Resume Next
   Me.Refresh
End Sub
