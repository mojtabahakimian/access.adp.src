Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit
Dim CTL As Controls

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
       Forms![F_MENU_KOL_MOIN_DATE]![HHMOIN] = Me.NUMBER
       Forms![F_MENU_KOL_MOIN_DATE]![HMOIN] = Me.NUMBER
    DoCmd.Close
End If
 If KeyAscii = 27 Then
   DoCmd.Close
 End If
End Sub
