Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Option Compare Database
Option Explicit

Private Sub Command46_Click()
   DoCmd.OpenForm "AMAR_TFT_FROOSH", acFormDS

End Sub

Private Sub Command47_Click()
    DoCmd.OpenReport "froosh_vaheds", acViewPreview, , "DEPATMAN =" & Me.DEPATMAN

End Sub

Private Sub Form_Close()
                    On Error Resume Next
   DoCmd.Close acForm, "F_MENU_DATE"
End Sub

Private Sub Form_Current()
                 On Error Resume Next
  Me.FROOSH_COUNT_SUB.Form.Filter = "DEPATMAN = " & Me.DEPATMAN
  Me.FROOSH_COUNT_SUB.Form.FilterOn = True
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
