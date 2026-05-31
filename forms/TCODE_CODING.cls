Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Close()
    Me.ServerFilter = ""
End Sub

Private Sub Form_Current()
                     On Error Resume Next
   DoCmd.RunCommand (acCmdSelectRecord)

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
                  On Error Resume Next
   Dim rst As New ADODB.Recordset
 If KeyAscii = 13 Then
    Forms![head_lst_havl]![MAS] = Me.CITYCODE
    DoCmd.Close acForm, Me.NAME
  End If
End Sub

Private Sub Form_Open(CANCEL As Integer)
                 On Error Resume Next

   DoCmd.GoToRecord acDataForm, Me.NAME, acNewRec
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub


Private Sub OSCODE_DblClick(CANCEL As Integer)
   DoCmd.OpenForm "TCOD_OSTAN", acFormDS, , , , acDialog

End Sub
