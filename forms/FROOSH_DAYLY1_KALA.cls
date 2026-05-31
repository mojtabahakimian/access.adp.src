Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub NUMBER_DblClick(CANCEL As Integer)
  DoCmd.OpenReport "R_FROOSH_DAYLY1_dakhel_kala", acViewPreview, , "date_n >= " & Forms![F_MENU_DATE]![DT1] & " and  date_n <= " & Forms![F_MENU_DATE]![DT2] & " and hes = '" & Me.HES & "'"
End Sub

Private Sub NUMBER1_DblClick(CANCEL As Integer)
  DoCmd.OpenReport "R_FROOSH_DAYLY1_dakhel_kala", acViewPreview, , "date_n >= " & Forms![F_MENU_DATE]![DT1] & " and  date_n <= " & Forms![F_MENU_DATE]![DT2]
End Sub
