Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Close()
   DoCmd.Close acForm, "F_MENU_DATE"
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub FROOSH_DblClick(CANCEL As Integer)
  DoCmd.OpenReport "R_FROOSH_DAYLY_person_mba2", acViewPreview, , "sader = 0 "
End Sub
Private Sub FROOSHKH_DblClick(CANCEL As Integer)
  DoCmd.OpenReport "R_FROOSH_DAYLY1_sader", acViewPreview, , "date_n >= " & Forms![F_MENU_DATE]![DT1] & " and  date_n <= " & Forms![F_MENU_DATE]![DT2]
End Sub
Private Sub GHABEL_DblClick(CANCEL As Integer)
  DoCmd.OpenForm "FROOSH_DAYLY1", acFormDS, , "date_n >= " & Forms![F_MENU_DATE]![DT1] & " and  date_n <= " & Forms![F_MENU_DATE]![DT2]
End Sub
Private Sub hes_DblClick(CANCEL As Integer)
  DoCmd.OpenReport "R_FROOSH_DAYLY_person", acViewPreview
End Sub
Private Sub NAME_DblClick(CANCEL As Integer)
  DoCmd.OpenReport "R_FROOSH_DAYLY_person_mba", acViewPreview
End Sub
Private Sub SMBAA_DblClick(CANCEL As Integer)
  DoCmd.OpenReport "R_FROOSH_DAYLY1_sader", acViewPreview, , "date_n >= " & Forms![F_MENU_DATE]![DT1] & " and  date_n <= " & Forms![F_MENU_DATE]![DT2] & " and hes = '" & Me.HES & "'"
End Sub
Private Sub SumOfTAKHFIF_DblClick(CANCEL As Integer)
  DoCmd.OpenReport "R_FROOSH_DAYLY_person_mba2", acViewPreview, , "sader = 1 "
End Sub
