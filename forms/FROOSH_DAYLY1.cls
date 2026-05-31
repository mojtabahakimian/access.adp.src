Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub CUST_NO_DblClick(CANCEL As Integer)
  DoCmd.OpenForm "FROOSH_DAYLY1_KALA", acFormDS, , "date_n >= " & Forms![F_MENU_DATE]![DT1] & " and  date_n <= " & Forms![F_MENU_DATE]![DT2] & " and hes = '" & Me.CUST_NO & "'"
End Sub

Private Sub MABL_KS_DblClick(CANCEL As Integer)
  DoCmd.OpenForm "FROOSH_DAYLY1_KALA", acFormDS, , "date_n >= " & Forms![F_MENU_DATE]![DT1] & " and  date_n <= " & Forms![F_MENU_DATE]![DT2]
End Sub

Private Sub NUMBER_DblClick(CANCEL As Integer)
  DoCmd.OpenReport "R_FROOSH_DAYLY1_dakhel", acViewPreview, , "date_n >= " & Forms![F_MENU_DATE]![DT1] & " and  date_n <= " & Forms![F_MENU_DATE]![DT2] & " and hes = '" & Me.CUST_NO & "'"
End Sub

Private Sub NUMBER1_DblClick(CANCEL As Integer)
  DoCmd.OpenReport "R_FROOSH_DAYLY1_dakhel", acViewPreview, , "date_n >= " & Forms![F_MENU_DATE]![DT1] & " and  date_n <= " & Forms![F_MENU_DATE]![DT2]
End Sub
