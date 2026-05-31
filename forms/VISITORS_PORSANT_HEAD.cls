Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command100_Click()
    DoCmd.OpenReport "OLGU_POORSANT", acPreview, , "PORID = " & Me.PORID

End Sub

Private Sub Command23_Click()
    DoCmd.OpenForm "visit_STUFGR_SEL_KALA", acFormDS
End Sub

Private Sub Command23_LostFocus()
   Me.VISITORS_PORSANT_KALAF.Requery
End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
  If IsNull(Me.PORID) Then
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT Max(porid) AS MaxOfpid FROM VISITORS_PORSANT", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
         Me.PORID = 1
    Else
        Me.PORID = rst.Fields(0) + 1
    End If
    rst.Close
  End If
  Me.VISITORS_PORSANT_KALAF.Locked = False
  Me.VISITORS_PORSANT_KALAF.Enabled = True
  Me.VISITORS_PORSANT_KALAF.Requery
End Sub


Private Sub Form_Current()
                 On Error Resume Next
If Me.NewRecord Then
     Me.VISITORS_PORSANT_KALAF.Locked = True
     Me.Command23.Enabled = False
  Else
     Me.VISITORS_PORSANT_KALAF.Locked = False
     Me.Command23.Enabled = True
 End If

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
                  On Error Resume Next
   DoCmd.GoToRecord acDataForm, Me.NAME, acNewRec
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
