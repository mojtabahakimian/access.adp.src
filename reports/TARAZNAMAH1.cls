Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Report_Open(CANCEL As Integer)
   Me.RecordSource = "SELECT NUMBER, NAM, CODE, NAMES, daraee, GRP FROM TARAZNAMAH(" & Forms![F_MENU_SND]![DT1] & ") TARAZNAMAH WHERE (GRP = 1) AND (daraee <> 0)"
End Sub
