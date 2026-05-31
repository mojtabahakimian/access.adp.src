Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub NUMBER_DblClick(CANCEL As Integer)
    If LETSGO("FACTKH") Then
        DoCmd.OpenForm "HEAD_LST_FROOSH22", , , , , , "NUMBER = " & Me.NUMBER1
    End If
End Sub
