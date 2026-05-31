Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit


Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"

End Sub

Private Sub MODATE_AfterUpdate()
                      On Error Resume Next
   DoCmd.RunCommand acCmdSaveRecord
  Me.MAND = MORAKH_MAND(Me.MODATE, Me.CODE)
  If Me.MAND = 110110110 Then
    Me.MORAKHDAY = 0
    Me.MAND = 0
  End If
End Sub

Private Sub MODATE_BeforeUpdate(CANCEL As Integer)
                     On Error Resume Next: CANCEL = CHEKDATEM(MODATE, False)
End Sub

Private Sub MORAKHDAY_AfterUpdate()
                     On Error Resume Next
  DoCmd.RunCommand acCmdSaveRecord
  Me.MAND = MORAKH_MAND(Me.MODATE, Me.CODE)
  If Me.MAND = 110110110 Then
    Me.MORAKHDAY = 0
    Me.MAND = 0
  End If
End Sub
