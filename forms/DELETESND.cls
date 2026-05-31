Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit

Sub Command0_Click()
 On Error GoTo Err_Command0_Click


    Screen.PreviousControl.SetFocus
    DoCmd.FindNext

Exit_Command0_Click:
    Exit Sub

Err_Command0_Click:
    MsgBox Err.Description
    Resume Exit_Command0_Click

End Sub

Private Sub Command1_Click()
                     On Error Resume Next
  Forms![BASEKNOW]![Text44] = True
  DoCmd.Close acForm, Me.NAME
End Sub


Private Sub Command13_Click()
                     On Error Resume Next
  DoCmd.Close
End Sub

Private Sub Form_Load()
                     On Error Resume Next
  DoCmd.restore
  Me.MSG.CAPTION = Me.MSG.CAPTION & Me.OpenArgs
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
