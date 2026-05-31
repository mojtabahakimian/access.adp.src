Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

Private Sub DATE_S_AfterUpdate()
                 On Error Resume Next
    ROOS = DIFF(Forms![RAAS]![Text0], Me.DATE_S)
   Me.MABK = ROOS * Me.mabl
End Sub

Private Sub DATE_S_BeforeUpdate(CANCEL As Integer)
                     On Error Resume Next: CANCEL = (CHEKDATEM(DATE_S, False))
End Sub

Private Sub date_s_Exit(CANCEL As Integer)
                 On Error Resume Next: CANCEL = (CHEKDATEM(DATE_S, False))
End Sub

Private Sub Form_Current()
                   On Error Resume Next
  Call SETSECURITYSUB("RAAS", Me.NAME, "RAAS", 3)
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Load()
                   On Error Resume Next
  Call SETSECURITYSUB("RAAS", Me.NAME, "RAAS", 3)
End Sub

Private Sub MABL_AfterUpdate()
                 On Error Resume Next
   Me.MABK = ROOS * Me.mabl
End Sub

Private Sub MABL_KeyPress(KeyAscii As Integer)
                      On Error Resume Next
  sendkeyu (KeyAscii)
End Sub
