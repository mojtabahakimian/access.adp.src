Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub Detail_Format(CANCEL As Integer, FormatCount As Integer)
                On Error Resume Next
    Me.Line (0, 0)-(20, Me.ScaleHeight), , BF
    Me.Line (5200, 0)-(5220, Me.ScaleHeight), , BF
    Me.Line (10510, 0)-(10530, Me.ScaleHeight), , BF


End Sub

Private Sub Report_Close()
                On Error Resume Next
   DoCmd.Close acForm, "FMENU_taraz_4"
End Sub

Private Sub Report_Open(CANCEL As Integer)
                On Error Resume Next
  Me.dt.CAPTION = Format(Forms![FMENU_TARAZ_4]![DT1], "####/##/##")
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd

End Sub
