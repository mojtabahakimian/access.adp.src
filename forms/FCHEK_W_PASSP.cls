Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub N_S_DblClick(CANCEL As Integer)
                   On Error Resume Next
  DoCmd.OpenForm "DEED_HEAD", , , "N_S=" & Me.N_S
End Sub
