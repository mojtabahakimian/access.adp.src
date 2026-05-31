Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub CODE_DblClick(CANCEL As Integer)
                    On Error Resume Next
   DoCmd.OpenForm "F_MENU_KART"
   Forms![F_MENU_KART]![ANBAR] = Me.ANBAR
   DoCmd.OpenReport "R_KA_KALA", acViewPreview, , "code = '" & Me.CODE & "'"
   DoCmd.OpenForm "F_MENU_KART", , , , , acHidden
End Sub

Private Sub Form_Close()
    DoCmd.Close acForm, "F_MENU_ANBAR"
End Sub

Private Sub Form_Load()
                   On Error Resume Next
    Forms![BASEKNOW]![hhwin] = Me.hWnd
    Dim tbl As String
    tbl = "mogdt" & Forms![BASEKNOW]![USERCOD]
    Me.RecordSource = "SELECT    " & tbl & ".CODE, " & tbl & ".ANBAR, STUF_DEF.NAME, " & tbl & ".EGDATE, " & tbl & ".MAND, " & tbl & ".FII, " & tbl & ".MABLK, " & tbl & ".NAMES , " & tbl & ".GRCOD FROM  " & tbl & " INNER JOIN STUF_DEF ON " & tbl & ".CODE = STUF_DEF.CODE ORDER BY " & tbl & ".MAND DESC"
  BringWindowToTop (Forms![BASEKNOW]![hhwin])
  DoCmd.Maximize

End Sub
