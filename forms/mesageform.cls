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
  DoCmd.Close acForm, "mesageform"
End Sub



Private Sub Form_Open(CANCEL As Integer)
                   On Error Resume Next
  If Len(Me.OpenArgs) > 200 Then
    Me.Width = Me.Width * 2
    Me.MSG.Width = Me.MSG.Width * 2
    DoCmd.Maximize
  Else
    DoCmd.restore
  End If
  Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub


Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
