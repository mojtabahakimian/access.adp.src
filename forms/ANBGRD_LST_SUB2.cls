Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Current()
                   On Error Resume Next
  Call SETSECURITYSUB("ANBGRD_HEAD", Me.NAME, "ANGD", 3)
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                 On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub NUM2_AfterUpdate()
      Me.NUM3 = Me.NUM2
End Sub
