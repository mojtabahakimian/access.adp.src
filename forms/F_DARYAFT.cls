Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit
Private Sub Form_Current()
                   On Error Resume Next
  Call SETSECURITY(Me.NAME, "F_DARYAFT", 3)
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
Private Sub Form_Load()
                   On Error Resume Next
  Call SETSECURITY(Me.NAME, "F_DARYAFT", 3)
End Sub

Private Sub Form_Timer()
                      On Error Resume Next
   Me.Refresh
End Sub
