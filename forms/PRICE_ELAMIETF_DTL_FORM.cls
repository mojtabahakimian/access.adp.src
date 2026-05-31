Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_BeforeUpdate(CANCEL As Integer)
    If IsNull(Me.PETID) Or PETID = 0 Then
        PETID = GetLIDD("PRICE_ELAMIETF_DTL", "PETID")
    End If
End Sub

Private Sub GSVALUE_KeyPress(KeyAscii As Integer)
                On Error Resume Next
    sendkeyu (KeyAscii)
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
