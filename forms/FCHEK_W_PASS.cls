Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub N_SERI_DblClick(CANCEL As Integer)
                  On Error Resume Next
 DoCmd.OpenForm "PAY_GETD", , , "N_SERI=" & Me.N_SERI & " AND BANK=" & Me.BANK
End Sub
