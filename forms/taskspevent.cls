Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database





Private Sub Form_Error(DataErr As Integer, Response As Integer)
                   On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_GotFocus()
      On Error Resume Next
DoCmd.RunCommand acCmdRecordsGoToNew
End Sub


Private Sub Form_Open(CANCEL As Integer)
      On Error Resume Next
DoCmd.RunCommand acCmdRecordsGoToNew
End Sub

Private Sub IDNUM_DblClick(CANCEL As Integer)
         On Error Resume Next
   DoCmd.GoToRecord acActiveDataObject, , acNewRec
End Sub
