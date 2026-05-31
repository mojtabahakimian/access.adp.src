Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command100_Click()
      On Error Resume Next
    DoCmd.OpenReport "tozie_dasti", acPreview, , , , Me.TID
End Sub

Private Sub Command22_Click()
      On Error Resume Next
    DoCmd.OpenReport "TOZIE_FROOSH_ANBARS_HAVALA", acPreview, , , , Me.TID
End Sub

Private Sub Command23_Click()
    DoCmd.OpenForm "tozieselect1", , , , , acDialog

End Sub

Private Sub Form_AfterInsert()
  On Error Resume Next
     Me.TOZIE_SUB.Locked = False
End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
    Dim rst As New ADODB.Recordset
  If IsNull(Me.TID) Then
    rst.Open "SELECT     MAX(TID) AS Expr1 FROM         dbo.TOZIE", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
         Me.TID = 1
    Else
        Me.TID = rst.Fields(0) + 1
    End If
    Me.Command23.Enabled = True
    rst.Close
  End If

End Sub

Private Sub Form_Current()
  On Error Resume Next
If Me.NewRecord Then
     Me.TOZIE_SUB.Locked = True
  Else
     Me.TOZIE_SUB.Locked = False
 End If
 If IsNull(Me.TID) Then
    Me.Command23.Enabled = False
 Else
    Me.Command23.Enabled = True
 End If
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
     On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If MID(Forms![BASEKNOW]![OPTIONSS], 48, 1) = "5" Then
        Select Case KeyAscii
            Case 1610, 1609, 1656, 1744, 1741
                 KeyAscii = 1740
            Case 1603, 1706, 1890, 1708, 1707
                 KeyAscii = 1705
        End Select
    End If

End Sub

Private Sub Form_Open(CANCEL As Integer)
     On Error Resume Next
   DoCmd.GoToRecord acDataForm, Me.NAME, acNewRec
   DoCmd.Maximize
End Sub

Private Sub TDATE_AfterUpdate()
    If Not IsNull(Me.TDATE) Then
        DoCmd.RunCommand acCmdSaveRecord
    End If
End Sub
