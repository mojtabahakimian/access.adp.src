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

Private Sub hes_DblClick(CANCEL As Integer)
    DoCmd.OpenReport "R_khareed_DAYLY_person", acViewPreview
End Sub
