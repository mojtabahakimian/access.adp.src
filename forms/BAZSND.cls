Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

Private Sub Command15_Click()
                   On Error Resume Next
  Dim rst As New ADODB.Recordset, inc, CON, i As Long
  inc = 2000000000
  rst.Open "SELECT DEED_HED.DATE_S, DEED_HED.N_S FROM DEED_HED ORDER BY DEED_HED.DATE_S DESC , DEED_HED.N_S DESC;", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Text0 = "n"
   Text0.Visible = True
   Box7.Visible = True
   Me.Refresh
   Me.Text0 = Me.Text0 + "n"
   Me.Refresh
   Me.Text0.Requery
   Me.Repaint
    CON = rst.RecordCount / 50
   If CON < 1 Then
      CON = 1
   End If
  i = 1
   Randomize
  While Not rst.EOF()
     rst.Fields("N_S") = inc
     rst.update
     inc = inc - 1
     rst.MoveNext
     If (i Mod CON) = 0 Then
        Me.Text0 = Me.Text0 + "n"
        Me.Refresh
        Me.Text0.Requery
        Me.Repaint
    End If
    i = i + 1
  Wend
  inc = 1
   Text0 = "n"
  rst.MoveLast
  While Not rst.BOF()
     rst.Fields("N_S") = inc
     rst.update
     inc = inc + 1
     If (i Mod CON) = 0 Then
        Me.Text0 = Me.Text0 + "n"
        Me.Refresh
        Me.Text0.Requery
        Me.Repaint
    End If
    i = i + 1
     rst.MovePrevious
 Wend
 rst.Close
 DoCmd.OpenForm "MESAGEFORM", , , , , acDialog, "عمليات با موفقيت انجام شد"
 DoCmd.Close acForm, Me.NAME
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
