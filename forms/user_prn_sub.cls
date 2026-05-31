Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private Sub Form_Current()
  DoCmd.RunCommand acCmdSelectRecord
  Forms![F_USER_PERMITION FORMS]![USER_PRITER_SUB].Requery
End Sub

Private Sub Form_Open(CANCEL As Integer)
                  On Error Resume Next
 Forms![F_USER_PERMITION FORMS]![USER_PRITER_SUB].Requery
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                       On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub SAL_NAME1_Click()
                  On Error Resume Next
  Forms![F_USER_PERMITION FORMS]![USER_PRITER_SUB].Requery
End Sub
