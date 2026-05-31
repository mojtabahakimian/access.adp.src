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
   Dim errorno As Integer, sql As String, RST2 As New ADODB.Recordset, rst As New ADODB.Recordset, SERH As New ADODB.Recordset, RSID As New ADODB.Recordset
   Dim CON As Long, i As Long, SERNUM As Long
                            On Error Resume Next
         rst.Open "SELECT Count(OPGET_HED.ID) AS CountOfID FROM OPGET_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
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
         rst.Open "OPGET_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
         rst.MoveLast
         rst.MoveFirst
         Err.Clear
         While Not rst.EOF()
              RSID.Open "SELECT Max(PGET_HED.ID) AS MaxOfID FROM PGET_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
              SERNUM = RSID.Fields("MaxOfID") + 1
              RSID.Close
              SERH.Open "PGET_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
              SERH.AddNew
              SERH.Fields("MOLAH") = rst.Fields("MOLAH") & "  " & rst.Fields("ID")
              SERH.Fields("DATE") = rst.Fields("DATE")
              SERH.Fields("CUST_KIND") = rst.Fields("CUST_KIND")
              SERH.Fields("SHIFT") = rst.Fields("SHIFT")
              SERH.Fields("DEPATMAN") = rst.Fields("DEPATMAN")
              SERH.Fields("USER_NAME") = rst.Fields("USER_NAME")
              SERH.update
              Me.SERN = SERNUM
              Me.ONUM = rst.Fields("ID")
              DoCmd.OpenQuery "OAPPEN_CHEK_DARYAFTY"
              DoCmd.OpenQuery "OAPPEN_CHEK_PARDAKHTY"
              DoCmd.OpenQuery "OAPPEN_PGET_LST"
              SANADDP (SERNUM)
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
    DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "عمليات انتقال اطلاعات خزانه داري با موفقيت انجام شد"
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
