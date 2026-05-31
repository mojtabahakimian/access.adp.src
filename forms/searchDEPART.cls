Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Private Sub Form_Current()
                     On Error Resume Next
   DoCmd.RunCommand (acCmdSelectRecord)
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
                    On Error Resume Next
    If KeyAscii = 13 Then
        Select Case Me.OpenArgs
           Case "1": Forms![HEAD_LST_FROOSH22]![DEPATMAN] = Me.DEPATMAN
           Case "2": Forms![head_lst_brf_free]![DEPATMAN] = Me.DEPATMAN
           Case "3": Forms![HEAD_LST_PISHFROOSH2]![DEPATMAN] = Me.DEPATMAN
           Case "4": Forms![head_lst_brfr]![DEPATMAN] = Me.DEPATMAN
        End Select
        DoCmd.Close acForm, Me.NAME
    End If
    If KeyAscii = 27 Then
        DoCmd.Close acForm, Me.NAME
    End If
End Sub
