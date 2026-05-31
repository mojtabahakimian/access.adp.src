Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Close()
  DoCmd.Close acForm, "F_MENU_KART"
End Sub

Private Sub Form_Open(CANCEL As Integer)
   Me.RecordSource = "exec dbo.usp_KART_ANBAR_TOTAL_CLEAN_V3 " & Forms![ak_mogudi_anbar_list]![ANBAR]
End Sub
