Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub GroupFooter3_Format(CANCEL As Integer, FormatCount As Integer)
 If [GRP] = 1 Then
    Command33.Visible = True
    Command34.Visible = True
    Command37.Visible = False
    Command38.Visible = False
 Else
    Command33.Visible = False
    Command34.Visible = False
    Command37.Visible = True
    Command38.Visible = True
End If

End Sub
