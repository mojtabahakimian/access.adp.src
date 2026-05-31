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
   Dim errorno As Integer, sql As String, RST2 As New ADODB.Recordset, rst As New ADODB.Recordset, SERH As New ADODB.Recordset
   Dim CON As Long, i As Long, SERNUM As Long
                           On Error Resume Next
         rst.Open "SELECT Count(ODEED_HED.N_S) AS CountOfN_S FROM ODEED_HED WHERE (((ODEED_HED.NO_S)=0))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
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
         rst.Open "SELECT ODEED_HED.* FROM ODEED_HED WHERE (((ODEED_HED.NO_S)=0))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
         rst.MoveLast
         rst.MoveFirst
         Err.Clear
         While Not rst.EOF()
              RST2.Open "SELECT Max(DEED_HED.N_S) AS MaxOfN_S FROM DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
              If RST2.RecordCount = 0 Or IsNull(RST2.Fields(0)) Then
                  SERNUM = 1
              Else
                  SERNUM = RST2.Fields(0) + 1
              End If
              RST2.Close
              SERH.Open "DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
              SERH.AddNew
              SERH.Fields("N_S") = SERNUM
              SERH.Fields("SHARH_S") = rst.Fields("SHARH_S") & "  " & rst.Fields("BASE")
              SERH.Fields("DATE_S") = rst.Fields("DATE_S")
              SERH.Fields("NO_S") = rst.Fields("NO_S")
              SERH.Fields("GHATEI") = rst.Fields("GHATEI")
              SERH.Fields("USER_NAME") = rst.Fields("USER_NAME")
              SERH.update
              Me.SERN = SERNUM
              Me.ONUM = rst.Fields("N_S")
              DoCmd.OpenQuery "OAPPEN_CHEKD_SANAD"
              DoCmd.OpenQuery "OAPPEN_CHEKP_SANAD"
              DoCmd.OpenQuery "OAPPEN_DEED_DTL"
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
   If Err = 0 Then
    DoCmd.Hourglass False
    DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "عمليات انتقال اطلاعات حسابداري داري با موفقيت انجام شد"
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
