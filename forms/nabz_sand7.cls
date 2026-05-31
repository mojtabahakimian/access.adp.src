Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private Sub Form_Close()
  DoCmd.Close acForm, "F_MENU_MON"
End Sub

Private Sub Form_Open(CANCEL As Integer)
  '  Me.RecordSource = "alter view nabz_mali7 as SELECT top(7) DEED_HED.DATE_S, SUM(DEED_DTL.BED) AS mab FROM DEED_DTL INNER JOIN DEED_HED ON DEED_DTL.N_S = DEED_HED.N_S WHERE (DEED_DTL.HES = N'" & Forms![baseknow]![HESNAGHD] & "') GROUP BY DEED_HED.DATE_S"
    DoCmd.Maximize
End Sub
