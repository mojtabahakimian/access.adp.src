Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
                     On Error Resume Next
  If KeyAscii = 13 Then
     If Me.MAND > 0 Then
        Forms![head_lst_rasid]![NUMBER1] = Me.NUMBER
    End If
    DoCmd.Close acForm, Me.NAME
End If
End Sub


Private Sub Form_Open(CANCEL As Integer)
   If Me.OpenArgs = 1 Then
    Me.RecordSource = "SELECT NUMBER, TAG, DATE_N, TAH, FNUMCO, SADER, MEGHk, MEGHKS, MAND FROM MEGH_DAR WHERE (TAG = 23) AND (MAND > 0 ) ORDER BY NUMBER DESC"
   End If
End Sub

Private Sub NUMBER_DblClick(CANCEL As Integer)
   DoCmd.OpenForm "HEAD_LST_REQUEST", , , "NUMBER = " & Me.NUMBER, , , 1
End Sub

Private Sub TAH_DblClick(CANCEL As Integer)
   DoCmd.OpenForm "HAVLISTDAR_SUB2", acFormDS
End Sub
