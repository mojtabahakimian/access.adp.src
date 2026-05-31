Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Report_Open(CANCEL As Integer)
   Me.RecordSource = "SELECT NUMBER, NAM, CODE, NAMES, BEDEHE, grp, GR, BEDEHECH, BEDEHEE, DARAEECH, BEDEHECH + DARAEECH * - 1 AS change FROM SURATMANABE(" & Forms![FMENU_TARAZ_4]![SNDNUM1] & ", " & Forms![FMENU_TARAZ_4]![DT1] & ", " & Forms![FMENU_TARAZ_4]![SNDNUM2] & ", " & Forms![FMENU_TARAZ_4]![DT2] & ") SURATMANABE WHERE     (GR = 2) AND (BEDEHECH >= 0) OR  (GR = 1) AND (DARAEECH < 0)"
End Sub
