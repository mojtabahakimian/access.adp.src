Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub ENDBLK_Click()
  If ENDBLK Then
    Me.ENDBLKDT = FARSIDATE(Now())
  Else
      Me.ENDBLKDT = Null
  End If
  Me.EUSER = UCurrentUser()
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                       On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"

End Sub

Private Sub Form_Open(CANCEL As Integer)
  If Me.OpenArgs = "1" Then
        DoCmd.Maximize
  End If
End Sub
