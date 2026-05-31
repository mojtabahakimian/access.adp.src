Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private Sub Form_Current()
  DoCmd.RunCommand acCmdSelectRecord
  Forms![F_USER_PERMITION FORMS]![BLOCK_HES_SUB].Requery
  Forms![F_USER_PERMITION FORMS]![BLOCKNON_HES_SUB].Requery
End Sub

Private Sub Form_Open(CANCEL As Integer)
  Forms![F_USER_PERMITION FORMS]![BLOCK_HES_SUB].Requery
  Forms![F_USER_PERMITION FORMS]![BLOCKNON_HES_SUB].Requery
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                       On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub SAL_NAME1_Click()
  Forms![F_USER_PERMITION FORMS]![BLOCK_HES_SUB].Requery
  Forms![F_USER_PERMITION FORMS]![BLOCKNON_HES_SUB].Requery
End Sub
