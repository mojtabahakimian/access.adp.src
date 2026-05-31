Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit

Private Sub bes_Label_DblClick(CANCEL As Integer)
    DoCmd.OpenForm "TARAZ_4_MAH"
End Sub

Private Sub Form_Close()
                      On Error Resume Next
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"

End Sub

Private Sub Form_Open(CANCEL As Integer)
                     On Error Resume Next
   If Forms![FMENU_TARAZ_4]![VI] = 2 Then
        Me.Filter = "BED <> 0 OR BES <> 0"
        Me.FilterOn = True
   End If
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub كد_مركز_DblClick(CANCEL As Integer)
                      On Error Resume Next
    DoCmd.OpenReport "R_MARKAZ", acViewPreview
End Sub
