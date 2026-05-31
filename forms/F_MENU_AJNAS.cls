Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Option Compare Database
Option Explicit

Private Sub Command5_Click()
                   On Error Resume Next
    If M0 Then
        DoCmd.OpenReport "STUF_DEF", acViewPreview
    Else
        DoCmd.OpenReport "STUF_DEF", acViewPreview, , "MANDah > 0 "
    End If
   DoCmd.OpenForm "F_MENU_AJNAS", acNormal, , , , acHidden

End Sub

Private Sub Command6_Click()
                    On Error Resume Next
    DoCmd.Close
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
