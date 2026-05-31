Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub ORDR_HED_DATE_DblClick(CANCEL As Integer)
  DoCmd.OpenForm "ORDR_HED2", acFormDS
End Sub
