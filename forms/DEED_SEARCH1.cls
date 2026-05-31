Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
                      On Error Resume Next
   If KeyAscii = 96 Or KeyAscii = 129 Then
      DoCmd.OpenForm "DEED_HEAD", , , "N_S = " & Me.N_S
   End If
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Load()
                     On Error Resume Next
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub N_S_DblClick(CANCEL As Integer)
                      On Error Resume Next
   DoCmd.RunCommand acCmdSaveRecord
   DoCmd.OpenForm "DEED_HEAD", , , "N_S = " & Me.N_S
End Sub
