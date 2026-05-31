Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Current()
   DoCmd.RunCommand (acCmdSelectRecord)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
              On Error Resume Next
  If KeyAscii = 13 Then
    DoCmd.OpenForm "BUGET_DEFAUL", , , , , acDialog, Me.BGID
  End If
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
