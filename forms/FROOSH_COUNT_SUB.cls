Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
                 On Error Resume Next
  Me.Filter = "DEPATMAN = " & Forms![GOZARESH_FROOSH_MAHSUL]![DEPATMAN]
  Me.FilterOn = True
End Sub
