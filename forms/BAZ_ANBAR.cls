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
   Dim errorno As Integer, sql As String, RST2 As New ADODB.Recordset, rst As New ADODB.Recordset
   Dim CON, i As Long
                      On Error Resume Next
   rst.Open "SELECT INVO_LST.RADIF FROM INVO_LST ORDER BY INVO_LST.RADIF", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Me.Text0 = "n"
   Me.Text0.Visible = True
   Me.Box7.Visible = True
   Me.Refresh
   Me.Text0 = Me.Text0 + "n"
   Me.Refresh
   Me.Text0.Requery
   Me.Repaint
    CON = rst.RecordCount / 100
   If CON < 1 Then
      CON = 1
   End If
  i = 1
   Randomize
   While Not rst.EOF()

     rst.Fields("radif") = i
     rst.update
     If (i Mod CON) = 0 Then
        Me.Text0 = Me.Text0 + "n"
        Me.Refresh
        Me.Text0.Requery
        Me.Repaint
    End If
    i = i + 1

    rst.MoveNext
   Wend
   rst.Close
   Set rst = New ADODB.Recordset
   DoCmd.SetWarnings (False)
   DoCmd.OpenStoredProcedure "BAZSAZI"
   DoCmd.OpenStoredProcedure "BAZSAZIA"
   DoCmd.OpenStoredProcedure "BAZSAZIF"
   DoCmd.OpenStoredProcedure "BAZMAR"
   DoCmd.OpenStoredProcedure "BAZMANDO"
   Me.Text0 = "n"
   Me.Text0.Visible = True
   Me.Box7.Visible = True
   Me.Refresh
   Me.Text0 = Me.Text0 + "n"
   Me.Refresh
   Me.Text0.Requery
   Me.Repaint
   DoCmd.RunSQL "UPDATE    dbo.STUF_STK SET MOGODI_A = 0, MOGODI = 0"
   rst.Close
   Set rst = New ADODB.Recordset
   rst.Open "STUF_STK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   CON = rst.RecordCount / 100
   If CON < 1 Then
      CON = 1
   End If
   RST2.Open "B_MOG_ANBARHA", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   While Not rst.EOF
       If rst.Fields("CODE") = "40" Then
          DoCmd.Beep
       End If
        RST2.Filter = ("CODE = '" & rst.Fields("CODE") & "' AND ANBAR = " & rst.Fields("ANBAR"))
        Set RST2 = RST2
        If RST2.RecordCount = 0 Then
        Else
            If rst.Fields("MOGODI_A") + rst.Fields("MOGODI") <> RST2.Fields("MAND") Then

                rst.Fields("MOGODI") = RST2.Fields("MAND") - rst.Fields("MOGODI_A")
                rst.update
            End If
            If (i Mod CON) = 0 Then
                Me.Text0 = Me.Text0 + "n"
                Me.Refresh
                Me.Text0.Requery
                Me.Repaint
            End If
            i = i + 1
        End If

        rst.Fields("MOGODI") = Round(rst.Fields("MOGODI") * 10 ^ Forms![BASEKNOW]![DIG]) / 10 ^ Forms![BASEKNOW]![DIG]
        rst.update
        rst.MoveNext
   Wend
   RST2.Close
   rst.Close
    DoCmd.Hourglass False
    DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "عمليات بازسازي  موجودي انبارها با موفقيت انجام شد"
    Me.Text0.Visible = False
    Me.Box7.Visible = False
    Me.Refresh
    DoCmd.Close acForm, Me.NAME

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub Form_Timer()
                     On Error Resume Next
   Me.Refresh
End Sub
