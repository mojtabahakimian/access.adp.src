Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub NUMBER_DblClick(CANCEL As Integer)
   DoCmd.OpenForm "HEAD_LST_KHAREED1", , , , , , "NUMBER = " & Me.NUMBER
End Sub
