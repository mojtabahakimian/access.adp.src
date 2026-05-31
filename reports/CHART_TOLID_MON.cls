Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Report_Close()
   DoCmd.Close acForm, "F_MENU_MON"
End Sub

Private Sub Report_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
  Me.Label3.CAPTION = "ميزان توليد محصولات در ماه " & MAHNAME([Forms]![F_MENU_MON]![MON])

End Sub
