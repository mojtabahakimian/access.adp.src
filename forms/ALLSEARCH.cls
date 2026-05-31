Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
                      On Error Resume Next
   If KeyAscii = 27 Then
      DoCmd.Close acForm, Me.NAME
   End If
   If KeyAscii = 13 Then
      If IsNull(Text0) Then
            DoCmd.OpenForm "ALLSERCHM"
      Else
            DoCmd.OpenForm "ALLSERCHM", , , , , , Text0
      End If
      DoCmd.Close acForm, Me.NAME
     End If
End Sub
