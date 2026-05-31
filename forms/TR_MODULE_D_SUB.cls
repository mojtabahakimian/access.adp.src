Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Current()
                   On Error Resume Next
  Call SETSECURITYSUB("STUF_DEF", "Child32", "KALA", 3)
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Load()
                   On Error Resume Next
  Call SETSECURITYSUB("STUF_DEF", "Child32", "KALA", 3)
End Sub

Private Sub MABL_F_KeyPress(KeyAscii As Integer)
                      On Error Resume Next
sendkeyu (KeyAscii)
End Sub

Private Sub NESBAT_KeyPress(KeyAscii As Integer)
                      On Error Resume Next
sendkeyu (KeyAscii)
End Sub

Private Sub VAHED_DblClick(CANCEL As Integer)
                    On Error Resume Next
   Me.vahed.Requery
End Sub
