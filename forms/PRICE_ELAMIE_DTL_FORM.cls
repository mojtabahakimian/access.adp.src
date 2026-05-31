Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_BeforeUpdate(CANCEL As Integer)
    If IsNull(Me.PERID) Or PERID = 0 Then
        PERID = GetLIDD("PRICE_ELAMIE_DTL", "PERID")
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

Private Sub PRICE1_KeyPress(KeyAscii As Integer)
    sendkeyu (KeyAscii)
End Sub
