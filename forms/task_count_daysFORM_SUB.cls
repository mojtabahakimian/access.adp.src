Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub USERNAME_DblClick(CANCEL As Integer)
  If IsNull(Me.USERNAME) Or Me.USERNAME = "" Then
    CANCEL = True
    Exit Sub
  End If
  DoCmd.OpenForm "TASKS_COMPANY", acFormDS, , , , , Me.USERNAME

End Sub
